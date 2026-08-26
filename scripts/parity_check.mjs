#!/usr/bin/env node
// Task #8 — proves "OLD DASHBOARD OUTPUT === NEW BACKEND-DRIVEN OUTPUT" for
// every company this repo has a real MIS file for, WITHOUT needing a live
// Supabase project. It exercises the real production code on both sides:
//
//   OLD (direct-parse) path:      parseWorkbook(wb) -> buildDataset()
//   NEW (backend-driven) path:    parseWorkbook(wb)
//                                  -> flattenParsedToMetricRows()   [misProcessing.js's write path]
//                                  -> reconstructParsedFromMetricRows()  [dashboardRead.js's read path]
//                                  -> buildDataset()
//
// then deep-diffs the two buildDataset() outputs. Any difference here would
// mean the DB round trip (flatten -> store -> reconstruct) lost or altered
// data before the SAME calculation engine ever saw it — exactly the class
// of regression the master spec's "old vs new must match" requirement is
// there to catch. It intentionally does NOT touch Supabase: that would only
// prove the network/SQL plumbing works, not that the round trip is
// lossless, and it lets this run in CI with no credentials.
import * as XLSX_NS from "xlsx";
const XLSX = XLSX_NS.default || XLSX_NS;
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseWorkbook, parseCompanyInfo, detectCompanyConfigStrict, buildDataset, COMPANY_CONFIGS } from "../src/lib/misEngine.js";
import { flattenParsedToMetricRows, reconstructParsedFromMetricRows } from "../lib/misMetricsShape.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");

const FILES = [
  ["excel_template/Easyrewardz_Live.xlsx", "easyrewardz"],
  ["excel_template/GrayQuest_Mastersheet_Template.xlsx", "grayquest"],
  ["excel_template/Riskcovry_Mastersheet_Template.xlsx", "riskcovry"],
  ["excel_template/Multipl_Mastersheet_Template.xlsx", "multipl"],
  ["misdata/FASTSURANCE_Standardized_MIS_Template.xlsx", "fastsurance"],
  ["misdata/Leegality_Standardized_MIS_Template.xlsx", "leegality"],
  ["misdata/APEX_Vitra_Standardized_MIS_Template_CORRECTED.xlsx", "apexFutureLabs"],
  ["misdata/FINBOX_Standardized_MIS_Template.xlsx", "finbox"],
  ["misdata/FUNDAMENTO_Standardized_MIS_Template.xlsx", "fundamento"],
];

// cardConfigs entries carry a `fmt` function reference (a formatter, not
// data) — strip those before diffing since functions were never part of
// the "output" the master spec cares about, and two separately-created
// arrow functions are never === to each other anyway (false-positive diffs).
function stripFns(value) {
  if (Array.isArray(value)) return value.map(stripFns);
  if (value && typeof value === "object") {
    const out = {};
    for (const [k, v] of Object.entries(value)) {
      if (typeof v === "function") continue;
      out[k] = stripFns(v);
    }
    return out;
  }
  return value;
}

// Only compare the parts of `ds` that represent actual dashboard OUTPUT —
// not companyConfig (the config object itself, which is identical by
// reference on both sides here and carries regex/function fields that
// aren't meaningfully diffable anyway).
function comparableOutput(ds) {
  return stripFns({
    months: ds.months,
    kpiKeys: ds.kpiKeys,
    hasRevenue: ds.hasRevenue,
    hasGP: ds.hasGP,
    hasEBITDA: ds.hasEBITDA,
    hasNet: ds.hasNet,
    fyData: ds.fyData,
    qData: ds.qData,
    cardConfigs: ds.cardConfigs,
  });
}

function deepDiff(a, b, pathPrefix = "") {
  if (a === b) return [];
  if (typeof a !== typeof b) return [`${pathPrefix}: type ${typeof a} vs ${typeof b}`];
  if (a === null || b === null || typeof a !== "object") {
    return (a === b || (Number.isNaN(a) && Number.isNaN(b))) ? [] : [`${pathPrefix}: ${JSON.stringify(a)} !== ${JSON.stringify(b)}`];
  }
  const diffs = [];
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  for (const k of keys) {
    diffs.push(...deepDiff(a[k], b[k], pathPrefix ? `${pathPrefix}.${k}` : k));
    if (diffs.length > 20) break; // don't flood the console for one badly-mismatched file
  }
  return diffs;
}

function parseOnlyArg() {
  const arg = process.argv.find(a => a.startsWith("--only="));
  if (!arg) return null;
  return new Set(arg.slice("--only=".length).split(",").map(s => s.trim()).filter(Boolean));
}

const only = parseOnlyArg();
let allOk = true;

for (const [rel, expectedSlug] of FILES) {
  if (only && !only.has(expectedSlug)) continue;
  const full = path.join(root, rel);
  const buffer = fs.readFileSync(full);
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });

  // --- OLD path: exactly what the original client-only App() did ---
  const parsed = parseWorkbook(wb);
  const companyInfo = parseCompanyInfo(wb);
  const detected = detectCompanyConfigStrict(Object.keys(parsed.kpis));
  const detectionOk = detected?.id === expectedSlug;
  const config = COMPANY_CONFIGS[expectedSlug];
  const dsOld = buildDataset(parsed, companyInfo, config);

  // --- NEW path: the exact functions misProcessing.js / dashboardRead.js use,
  // in the exact same order (validate via buildDataset FIRST, exactly like
  // misProcessing.js does, then flatten) — this is what caught buildDataset()
  // previously mutating `parsed.kpis` with synthesized alias keys (FASTSURANCE)
  // and flattening against raw, uncorrected month dates (Fundamento) before
  // both were fixed; running old-then-new in this order keeps this check
  // honest about the real call sequence instead of a more forgiving one. ---
  const rows = flattenParsedToMetricRows({ ...parsed, months: dsOld.months });
  // dsOld.months is the authoritative axis stored on mis_uploads.months in
  // production (see misProcessing.js) — passed explicitly here too, exactly
  // like dashboardRead.js does, rather than inferred from the rows alone.
  const reconstructed = reconstructParsedFromMetricRows(rows, dsOld.months);
  const dsNew = buildDataset(reconstructed, companyInfo, config);

  const diffs = deepDiff(comparableOutput(dsOld), comparableOutput(dsNew));
  const ok = detectionOk && diffs.length === 0;
  if (!ok) allOk = false;

  console.log(`\n${rel}`);
  console.log(`  company detection: ${detected ? detected.id : "NONE"} ${detectionOk ? "OK" : `*** expected ${expectedSlug} ***`}`);
  console.log(`  old vs new dashboard output: ${diffs.length === 0 ? "IDENTICAL" : `*** ${diffs.length} DIFFERENCE(S) ***`}`);
  diffs.slice(0, 10).forEach(d => console.log(`    - ${d}`));
}

console.log(allOk ? "\nALL PARITY CHECKS PASSED — old and new dashboards produce identical output." : "\n*** SOME PARITY CHECKS FAILED — see above ***");
process.exit(allOk ? 0 : 1);
