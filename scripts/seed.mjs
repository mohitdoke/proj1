#!/usr/bin/env node
// One-time / repeatable data import: loads each company's real master
// Excel file straight through the SAME processMisUpload() pipeline a
// manager's HTTP upload uses (lib/misProcessing.js) — never by typing
// Excel values into SQL. Safe to re-run: each file just becomes that
// company's newest version (see the versioning rules in
// lib/misProcessing.js — a bad file never corrupts a good one).
//
// Prereqs before running this:
//   1. Apply supabase/migrations/0001_init.sql to your Supabase project.
//   2. Run supabase/seed/reference_data.sql (funds/companies/fund_companies/
//      company_configs) — this script only imports MIS METRICS, it assumes
//      the company/fund rows already exist.
//   3. Create a private Storage bucket named "mis-files" in that project.
//   4. Set SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY (see .env.example).
//
// Usage:  node scripts/seed.mjs
//         node scripts/seed.mjs --only=grayquest,riskcovry
import path from "node:path";
import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { processMisUpload } from "../lib/misProcessing.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");

// Every currently-available company MIS file this repo has on hand, mapped
// to the company slug it's expected to resolve to (used only to print a
// clear mismatch warning — detection itself still runs purely off the
// sheet's own row names, exactly like a manager upload would).
const SOURCE_FILES = [
  { slug: "easyrewardz", file: "excel_template/Easyrewardz_Live.xlsx" },
  { slug: "grayquest", file: "excel_template/GrayQuest_Mastersheet_Template.xlsx" },
  { slug: "riskcovry", file: "excel_template/Riskcovry_Mastersheet_Template.xlsx" },
  { slug: "multipl", file: "excel_template/Multipl_Mastersheet_Template.xlsx" },
  { slug: "apexFutureLabs", file: "misdata/APEX_Vitra_Standardized_MIS_Template_CORRECTED.xlsx" },
  { slug: "fastsurance", file: "misdata/FASTSURANCE_Standardized_MIS_Template.xlsx" },
  { slug: "finbox", file: "misdata/FINBOX_Standardized_MIS_Template.xlsx" },
  { slug: "fundamento", file: "misdata/FUNDAMENTO_Standardized_MIS_Template.xlsx" },
  { slug: "leegality", file: "misdata/Leegality_Standardized_MIS_Template.xlsx" },
  { slug: "knightFintech", file: "misdata/KnightFinTech_Standardized_MIS_Template.xlsx" },
  { slug: "traqcheck", file: "misdata/Traqcheck_Standardized_MIS_Template.xlsx" },
  { slug: "castler", file: "misdata/Castler_Standardized_MIS_Template.xlsx" },
  { slug: "datasutram", file: "misdata/DataSutram_Standardized_MIS_Template.xlsx" },
];

function parseOnlyArg() {
  const arg = process.argv.find(a => a.startsWith("--only="));
  if (!arg) return null;
  return new Set(arg.slice("--only=".length).split(",").map(s => s.trim()).filter(Boolean));
}

async function main() {
  if (!process.env.SUPABASE_URL || !process.env.SUPABASE_SERVICE_ROLE_KEY) {
    console.error("SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY must be set in the environment (see .env.example). Aborting.");
    process.exit(1);
  }

  const only = parseOnlyArg();
  const targets = only ? SOURCE_FILES.filter(s => only.has(s.slug)) : SOURCE_FILES;
  if (!targets.length) {
    console.error(`--only matched no known company. Known slugs: ${SOURCE_FILES.map(s => s.slug).join(", ")}`);
    process.exit(1);
  }

  let succeeded = 0, failed = 0;
  for (const { slug, file } of targets) {
    const fullPath = path.join(ROOT, file);
    process.stdout.write(`\n[${slug}] importing ${file} ... `);
    let buffer;
    try {
      buffer = await fs.readFile(fullPath);
    } catch (err) {
      console.log(`SKIPPED (file not found: ${err.message})`);
      failed++;
      continue;
    }

    const result = await processMisUpload({
      buffer,
      originalFilename: path.basename(file),
      uploadedBy: "seed-script",
      companySlugHint: slug,
    });

    if (result.ok) {
      console.log(`OK — ${result.companyName}: ${result.monthsCount} months, ${result.kpiCount} KPIs, ${result.fyCount} FY periods.`);
      succeeded++;
    } else {
      console.log(`FAILED — ${result.error}`);
      failed++;
    }
  }

  console.log(`\nSeed complete: ${succeeded} succeeded, ${failed} failed out of ${targets.length}.`);
  if (failed > 0) process.exitCode = 1;
}

main().catch(err => {
  console.error("Seed script crashed:", err);
  process.exit(1);
});
