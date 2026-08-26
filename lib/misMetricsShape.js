// Pure, DB-independent conversions between the in-memory `parsed` shape
// ({ months, kpis, headcount }, as produced by misEngine.js's
// parseWorkbook()) and the flat row shape mis_metrics stores one row per
// (metric, month) as. Used by BOTH:
//   - lib/misProcessing.js (write path: parsed -> rows, before insert)
//   - lib/dashboardRead.js (read path: rows -> parsed, after select)
// Having one implementation of each direction — rather than each of those
// two modules doing its own ad hoc flatten/reconstruct — is what lets
// scripts/parity_check.mjs prove the DB round trip is lossless without a
// live database: it calls these same two functions back to back, in
// memory, and diffs the result against a direct parse.

/**
 * parsed -> flat rows, ready to insert into mis_metrics (minus the
 * caller-supplied upload_id/company_id foreign keys). Genuinely blank
 * cells (null/undefined) are skipped entirely, never coerced to 0 — a
 * missing value must stay missing all the way through storage and back.
 */
export function flattenParsedToMetricRows(parsed) {
  const monthDates = parsed.months.map(m => `${m.y}-${String(m.m).padStart(2, "0")}-01`);
  const rows = [];

  Object.entries(parsed.kpis).forEach(([label, values]) => {
    const metric_key = label.trim().toLowerCase().replace(/\s+/g, " ");
    let anyStored = false;
    values.forEach((v, i) => {
      if (v === null || v === undefined) return;
      anyStored = true;
      rows.push({ metric_key, original_metric_label: label, period_date: monthDates[i], value: v, unit: null });
    });
    // A row that exists in the source sheet but is blank for every single
    // month (e.g. a KPI split the company simply doesn't track — see
    // COMPANY_CONFIGS[*].fallbackKPIs) must still come back as a KEY in
    // reconstructed kpis on the read side, all-null, exactly like
    // parseWorkbook() would produce from the live file — NOT vanish
    // entirely, which would silently change kpiKeys/cardConfigs and break
    // any company-specific fallback that only kicks in for "row present,
    // no real data" (as opposed to "row absent"). One explicit null-value
    // placeholder row (on the first month) is enough to preserve that.
    if (!anyStored && values.length) {
      rows.push({ metric_key, original_metric_label: label, period_date: monthDates[0], value: null, unit: null });
    }
  });

  if (parsed.headcount) {
    let anyStored = false;
    parsed.headcount.forEach((v, i) => {
      if (v === null || v === undefined) return;
      anyStored = true;
      rows.push({ metric_key: "headcount", original_metric_label: "Headcount", period_date: monthDates[i], value: v, unit: "count" });
    });
    if (!anyStored && parsed.headcount.length) {
      rows.push({ metric_key: "headcount", original_metric_label: "Headcount", period_date: monthDates[0], value: null, unit: "count" });
    }
  }

  return rows;
}

/**
 * The inverse: flat mis_metrics rows (each needing at least
 * `original_metric_label`, `period_date`, `value`) -> the same `{ months,
 * kpis, headcount }` shape parseWorkbook() produces from a live Excel
 * file. "Headcount" is pulled back out into its own array, mirroring
 * parseWorkbook()'s own special-casing of that row.
 *
 * `months` is the upload's OWN authoritative, ordered month axis (stored
 * verbatim on mis_uploads.months at write time — see misProcessing.js) —
 * NOT inferred from which period_dates happen to appear in metricRows. A
 * calendar month where literally every KPI (including revenue) is blank
 * would otherwise leave zero rows anywhere for that month, making it
 * indistinguishable from "this month was never part of the sheet" if we
 * tried to infer the axis from the rows alone (a real case in at least one
 * source workbook, which pre-lists not-yet-reported trailing months).
 */
export function reconstructParsedFromMetricRows(metricRows, months) {
  const monthIndexByKey = new Map(months.map((mo, i) => [mo.key, i]));

  const kpis = {};
  let headcount = null;
  metricRows.forEach(r => {
    const idx = monthIndexByKey.get(r.period_date.slice(0, 7));
    if (idx === undefined) return;
    const value = r.value === null || r.value === undefined ? null : Number(r.value);
    if (r.original_metric_label.trim().toLowerCase() === "headcount") {
      if (!headcount) headcount = new Array(months.length).fill(null);
      headcount[idx] = value;
    } else {
      if (!kpis[r.original_metric_label]) kpis[r.original_metric_label] = new Array(months.length).fill(null);
      kpis[r.original_metric_label][idx] = value;
    }
  });

  return { sheetName: "Monthly Data", months, kpis, headcount };
}
