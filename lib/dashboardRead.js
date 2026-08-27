// Read-side counterpart to lib/misProcessing.js. Everything here is a plain
// Supabase read plus a reshape back into the exact `parsed` object shape
// `parseWorkbook()` produces from a live Excel file — { months, kpis,
// headcount } — so the frontend can hand it straight to the shared
// `buildDataset()` (src/lib/misEngine.js) and get byte-identical output to
// the pre-migration, upload-driven dashboard. No calculation logic lives
// here: this module only reconstructs inputs to buildDataset(), it never
// computes revenue/margins/growth/etc itself.
import { COMPANY_CONFIGS } from "../src/lib/misEngine.js";
import { getSupabaseAdmin } from "./supabaseAdmin.js";
import { reconstructParsedFromMetricRows } from "./misMetricsShape.js";

// Postgres `uuid` columns reject any non-uuid literal outright — even inside
// an `.or(id.eq.X,slug.eq.X)` filter, a non-uuid X makes the WHOLE query
// error ("invalid input syntax for type uuid"), not just that one branch.
// Since callers pass either a real id OR a human slug ("fund-1",
// "easyrewardz"), we must pick exactly one column to query against instead
// of ever sending a non-uuid string to an `id.eq` comparison.
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function byIdOrSlug(query, idOrSlug) {
  return UUID_RE.test(idOrSlug) ? query.eq("id", idOrSlug) : query.eq("slug", idOrSlug);
}

/**
 * All funds, each with its ordered, active company list (slug + name only —
 * enough for the fund/company picker UI). Powers GET /api/funds.
 */
export async function listFunds() {
  const supabase = getSupabaseAdmin();

  const { data: funds, error: fundsErr } = await supabase
    .from("funds")
    .select("id, name, slug")
    .order("name");
  if (fundsErr) throw new Error(`Database error listing funds: ${fundsErr.message}`);

  const { data: memberships, error: memErr } = await supabase
    .from("fund_companies")
    .select("fund_id, display_order, companies(id, slug, name, active)")
    .eq("active", true)
    .order("display_order");
  if (memErr) throw new Error(`Database error listing fund companies: ${memErr.message}`);

  return funds.map(f => ({
    id: f.id,
    slug: f.slug,
    name: f.name,
    companies: memberships
      .filter(m => m.fund_id === f.id && m.companies?.active)
      .map(m => ({ id: m.companies.id, slug: m.companies.slug, name: m.companies.name })),
  }));
}

/**
 * The companies belonging to one fund (by fund id OR slug), ordered for
 * display. Powers GET /api/funds/:fundId/companies.
 */
export async function listCompaniesForFund(fundIdOrSlug) {
  const supabase = getSupabaseAdmin();

  const { data: fund, error: fundErr } = await byIdOrSlug(
    supabase.from("funds").select("id, name, slug"),
    fundIdOrSlug
  ).maybeSingle();
  if (fundErr) throw new Error(`Database error looking up fund: ${fundErr.message}`);
  if (!fund) return null;

  const { data: rows, error } = await supabase
    .from("fund_companies")
    .select("display_order, companies(id, slug, name, active)")
    .eq("fund_id", fund.id)
    .eq("active", true)
    .order("display_order");
  if (error) throw new Error(`Database error listing companies: ${error.message}`);

  return {
    fund,
    companies: rows.filter(r => r.companies?.active).map(r => ({ id: r.companies.id, slug: r.companies.slug, name: r.companies.name })),
  };
}

/**
 * Basic info for one company (by id OR slug) — enough for a header/lookup,
 * without pulling its full dashboard dataset. Powers GET /api/companies/:companyId.
 */
export async function getCompany(companyIdOrSlug) {
  const supabase = getSupabaseAdmin();
  const { data: company, error } = await byIdOrSlug(
    supabase.from("companies").select("id, slug, name, legal_name, active"),
    companyIdOrSlug
  ).maybeSingle();
  if (error) throw new Error(`Database error looking up company: ${error.message}`);
  return company || null;
}

/**
 * Company row + its resolved company_configs entry + the reconstructed
 * companyInfo (stored "Company Info" sheet override, or the code-side
 * defaultDescription) — everything both the dashboard read path AND the
 * research service (lib/research.js) need about a company, without pulling
 * its full metric history. Returns { ok:false, error } the same way the
 * rest of this module does, so callers can handle both uniformly.
 */
export async function getCompanyProfile(companyIdOrSlug) {
  const supabase = getSupabaseAdmin();

  const company = await getCompany(companyIdOrSlug);
  if (!company) return { ok: false, error: "Company not found." };

  const { data: configRow, error: configErr } = await supabase
    .from("company_configs")
    .select("config_key, overrides")
    .eq("company_id", company.id)
    .maybeSingle();
  if (configErr) return { ok: false, error: `Database error looking up company config: ${configErr.message}` };
  const configKey = configRow?.config_key;
  if (!configKey || !COMPANY_CONFIGS[configKey]) {
    return { ok: false, error: `Company "${company.slug}" has no valid company_configs.config_key matching a known COMPANY_CONFIGS entry.` };
  }

  const companyInfo = configRow.overrides?.company_info || COMPANY_CONFIGS[configKey].defaultDescription || null;
  return { ok: true, company, configKey, configRow, companyInfo };
}

/**
 * The full reconstructed input to buildDataset() for one company's CURRENT
 * upload: { parsed, companyInfo, configKey, uploadMeta }. Powers
 * GET /api/companies/:companyId/dashboard. The frontend calls
 * buildDataset(parsed, companyInfo, COMPANY_CONFIGS[configKey]) locally —
 * this function never runs that calculation itself, it only assembles the
 * same raw shape parseWorkbook() would have produced.
 */
export async function getCompanyDashboardInput(companyIdOrSlug) {
  const supabase = getSupabaseAdmin();

  const profile = await getCompanyProfile(companyIdOrSlug);
  if (!profile.ok) return profile;
  const { company, configKey, companyInfo } = profile;

  const { data: upload, error: uploadErr } = await supabase
    .from("mis_uploads")
    .select("id, original_filename, uploaded_at, uploaded_by, months_count, months, kpi_count, detected_config_key")
    .eq("company_id", company.id)
    .eq("is_current", true)
    .eq("processing_status", "succeeded")
    .maybeSingle();
  if (uploadErr) return { ok: false, error: `Database error looking up current upload: ${uploadErr.message}` };
  if (!upload) {
    return { ok: false, error: `No successfully processed MIS upload exists yet for "${company.name}". A manager needs to upload its MIS first.` };
  }
  if (!upload.months || !upload.months.length) {
    return { ok: false, error: `Current upload for "${company.name}" is missing its stored month axis — it may predate the mis_uploads.months column; re-upload to fix.` };
  }

  const { data: metricRows, error: metricsErr } = await supabase
    .from("mis_metrics")
    .select("original_metric_label, period_date, value")
    .eq("upload_id", upload.id)
    .order("period_date");
  if (metricsErr) return { ok: false, error: `Database error loading metrics: ${metricsErr.message}` };
  if (!metricRows.length) {
    return { ok: false, error: `Current upload for "${company.name}" has no stored metric rows.` };
  }

  // Rebuild the exact `{ months, kpis, headcount }` shape parseWorkbook()
  // would have produced, using the SAME reconstruction function
  // scripts/parity_check.mjs exercises directly against misProcessing.js's
  // flatten function — proving this round trip is lossless, not just
  // asserting it. `upload.months` (not inferred from the metric rows) is
  // the authoritative month axis — see reconstructParsedFromMetricRows's
  // own comment for why that matters.
  const parsed = reconstructParsedFromMetricRows(metricRows, upload.months);

  // companyInfo already resolved by getCompanyProfile() above: prefers what
  // was actually in this company's uploaded workbook (persisted at upload
  // time — see misProcessing.js step 7), falling back to the code-side
  // default for that company, exactly like the original client-only
  // dashboard did when no "Company Info" sheet existed.

  return {
    ok: true,
    parsed,
    companyInfo,
    configKey,
    company: { id: company.id, slug: company.slug, name: company.name, legalName: company.legal_name },
    uploadMeta: {
      uploadId: upload.id,
      originalFilename: upload.original_filename,
      uploadedAt: upload.uploaded_at,
      uploadedBy: upload.uploaded_by,
    },
  };
}
