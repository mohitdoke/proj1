// Shared upload-processing pipeline: Excel buffer in, validated + persisted
// MIS data out. Used by BOTH api/admin/mis/upload.js (the real HTTP upload
// route) and scripts/seed.mjs (the initial data import) — one implementation,
// so a seeded company and a manager-uploaded company are processed exactly
// the same way.
import crypto from "node:crypto";
import * as XLSX_NS from "xlsx";
const XLSX = XLSX_NS.default || XLSX_NS;
import {
  parseWorkbook,
  parseCompanyInfo,
  detectCompanyConfigStrict,
  buildDataset,
} from "../src/lib/misEngine.js";
import { getSupabaseAdmin, MIS_BUCKET } from "./supabaseAdmin.js";
import { flattenParsedToMetricRows } from "./misMetricsShape.js";

/**
 * @param {Object} opts
 * @param {Buffer} opts.buffer - raw .xlsx file bytes
 * @param {string} opts.originalFilename
 * @param {string} [opts.uploadedBy] - free-text identifier of who uploaded it
 * @param {string} [opts.companySlugHint] - if the caller already knows which
 *   company this SHOULD be (e.g. an admin picked one from a dropdown before
 *   uploading), cross-check it against what's actually detected from the
 *   sheet's own row names and reject on mismatch. Optional — detection never
 *   depends on this or on the filename.
 */
export async function processMisUpload({ buffer, originalFilename, uploadedBy, companySlugHint }) {
  const supabase = getSupabaseAdmin();

  // 1. Parse the workbook.
  let wb, parsed;
  try {
    wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
    parsed = parseWorkbook(wb);
  } catch (err) {
    return { ok: false, error: `Could not read this file as a valid MIS workbook: ${err.message || err}` };
  }

  // 2. Detect the company from the sheet's own row names — never from the
  // filename, and never a silent fallback to an existing company's template.
  const detected = detectCompanyConfigStrict(Object.keys(parsed.kpis));
  if (!detected) {
    return { ok: false, error: "Unable to confidently identify company from MIS." };
  }
  if (companySlugHint && companySlugHint !== detected.id) {
    return {
      ok: false,
      error: `This file's contents match "${detected.id}", not the company you selected ("${companySlugHint}"). Double-check you're uploading the right file.`,
    };
  }

  // 3. The detected company must already exist as a `companies` row (adding a
  // brand-new company is a deliberate, separate step — see README "Adding a
  // company" — not an implicit side effect of an upload).
  const { data: companyRow, error: companyErr } = await supabase
    .from("companies")
    .select("id, slug, name")
    .eq("slug", detected.id)
    .maybeSingle();
  if (companyErr) return { ok: false, error: `Database error looking up company: ${companyErr.message}` };
  if (!companyRow) {
    return {
      ok: false,
      error: `Detected company "${detected.id}" has no matching row in the companies table yet. Add it (supabase/seed/reference_data.sql or an insert into companies/company_configs/fund_companies) before uploading its MIS.`,
    };
  }

  // 4. Validate the parse actually produces a usable dashboard before storing
  // anything — never persist a file that would break the dashboard.
  const companyInfo = parseCompanyInfo(wb) || detected.defaultDescription || null;
  let ds;
  try {
    ds = buildDataset(parsed, companyInfo, detected);
  } catch (err) {
    return { ok: false, error: `Failed to compute dashboard data from this workbook: ${err.message || err}` };
  }
  if (!ds.hasRevenue) {
    return {
      ok: false,
      error: `This workbook parsed, but no recognizable revenue row ("${detected.revenueBaseKey}") was found for ${companyRow.name} — refusing to store an unusable dataset.`,
    };
  }

  // 5. Create the upload row (status "processing"), then store the raw file
  // and normalized rows. Nothing about the PREVIOUS current dataset is
  // touched until every step below succeeds.
  const uploadId = crypto.randomUUID();
  const storagePath = `${companyRow.slug}/${uploadId}-${originalFilename}`;

  const { error: insertErr } = await supabase.from("mis_uploads").insert({
    id: uploadId,
    company_id: companyRow.id,
    original_filename: originalFilename,
    storage_path: storagePath,
    uploaded_by: uploadedBy || null,
    processing_status: "processing",
    detected_config_key: detected.id,
    months_count: ds.months.length,
    // The authoritative month axis (see supabase/migrations/0001_init.sql
    // and misMetricsShape.js's reconstructParsedFromMetricRows) — buildDataset's
    // own forceConsecutiveMonths-corrected list, not the raw parsed.months.
    months: ds.months,
    kpi_count: ds.kpiKeys.length,
    is_current: false,
  });
  if (insertErr) return { ok: false, error: `Database error creating upload record: ${insertErr.message}` };

  try {
    const { error: storageErr } = await supabase.storage.from(MIS_BUCKET).upload(storagePath, buffer, {
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      upsert: false,
    });
    if (storageErr) throw new Error(`Storage upload failed: ${storageErr.message}`);

    // Every real MIS row becomes metric rows keyed by its EXACT original
    // label (what the calculation engine's regex matchers key off of).
    // Genuinely blank cells are skipped entirely — never coerced to 0. This
    // uses the SAME flatten function scripts/parity_check.mjs exercises
    // directly, so that check is proving the real write-path code, not a
    // reimplementation of it.
    //
    // Store against `ds.months` (buildDataset()'s own, already
    // forceConsecutiveMonths-corrected month list — see misEngine.js
    // resolveMonths()), NOT the raw `parsed.months` — a handful of source
    // workbooks (Fundamento) have a genuine year-labeling bug in their own
    // header row that gives two different columns the same nominal (y, m);
    // storing against the raw, uncorrected dates would collide those two
    // columns into one period_date and silently drop a month's data.
    // buildDataset() also never mutates parsed.kpis (see its own comment),
    // so flattening parsed.kpis here after the validation buildDataset()
    // call above is safe — it still reflects only genuine sheet rows, not
    // any company-specific alias it computed for validation.
    const rows = flattenParsedToMetricRows({ ...parsed, months: ds.months }).map(row => ({ ...row, upload_id: uploadId, company_id: companyRow.id }));

    const BATCH = 500;
    for (let i = 0; i < rows.length; i += BATCH) {
      const { error: metricsErr } = await supabase.from("mis_metrics").insert(rows.slice(i, i + BATCH));
      if (metricsErr) throw new Error(`Database error storing metrics: ${metricsErr.message}`);
    }

    // 6. Only now — file stored, every row validated and persisted — make
    // this the current version. Un-flag the previous current upload first,
    // then flag this one, so there's never a moment with zero OR two
    // "current" uploads for this company.
    const { error: unsetErr } = await supabase
      .from("mis_uploads")
      .update({ is_current: false })
      .eq("company_id", companyRow.id)
      .eq("is_current", true);
    if (unsetErr) throw new Error(`Database error updating previous version: ${unsetErr.message}`);

    const { error: finalizeErr } = await supabase
      .from("mis_uploads")
      .update({ processing_status: "succeeded", is_current: true })
      .eq("id", uploadId);
    if (finalizeErr) throw new Error(`Database error finalizing upload: ${finalizeErr.message}`);

    // 7. Persist the parsed "Company Info" sheet (if this workbook had one) as
    // display-only metadata on company_configs.overrides, keyed so the READ
    // path (lib/dashboardRead.js) can reconstruct `companyInfo` without
    // re-fetching and re-parsing the stored Excel file on every dashboard
    // request. This is NOT calculation logic — just the same optional
    // Business Description / tags / scale-metrics fields buildDataset()
    // already treats as pure display data. A workbook with no "Company Info"
    // sheet leaves this untouched, and the read path falls back to the
    // company's code-side `defaultDescription`, exactly like the original
    // client-only dashboard did.
    const parsedCompanyInfo = parseCompanyInfo(wb);
    if (parsedCompanyInfo) {
      const { data: cfgRow, error: cfgReadErr } = await supabase
        .from("company_configs")
        .select("overrides")
        .eq("company_id", companyRow.id)
        .maybeSingle();
      if (!cfgReadErr) {
        const nextOverrides = { ...(cfgRow?.overrides || {}), company_info: parsedCompanyInfo };
        await supabase
          .from("company_configs")
          .update({ overrides: nextOverrides })
          .eq("company_id", companyRow.id);
        // Best-effort: a failure here doesn't affect the successfully-stored
        // upload/metrics, so it's not treated as upload failure — the
        // dashboard read path simply falls back to defaultDescription.
      }
    }

    return {
      ok: true,
      uploadId,
      companyId: companyRow.id,
      companySlug: companyRow.slug,
      companyName: companyRow.name,
      monthsCount: ds.months.length,
      kpiCount: ds.kpiKeys.length,
      fyCount: ds.fyData.length,
    };
  } catch (err) {
    // Processing failed after the upload row was created — mark THIS upload
    // failed and stop. The previous "is_current" upload was never touched,
    // so the dashboard users see is unaffected.
    await supabase
      .from("mis_uploads")
      .update({ processing_status: "failed", processing_error: String(err.message || err) })
      .eq("id", uploadId);
    return { ok: false, error: String(err.message || err) };
  }
}
