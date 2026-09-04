import * as XLSX_NS from 'xlsx';
const XLSX = XLSX_NS.default || XLSX_NS;
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { parseWorkbook, parseCompanyInfo, detectCompanyConfig, buildDataset, COMPANY_CONFIGS } from '../src/lib/misEngine.js';
import { buildReportContent } from '../src/lib/reportContent.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projRoot = path.resolve(__dirname, '..');
const misRoot = path.resolve(projRoot, '..');
const outDir = path.resolve(misRoot, 'report_generator', 'reports');
fs.mkdirSync(outDir, { recursive: true });

const scriptPy = path.resolve(projRoot, 'api', 'reports', '_generate_company_report.py');

const COMPANIES = [
  {
    file: "datatemplate/APEX_Vitra_Standardized_MIS_Template_CORRECTED.xlsx",
    legal_entity_name: "Apex Future Labs Private Limited", brand: "Vitra.ai",
  },
  {
    file: "datatemplate/EasyRewardz_Mastersheet_Template.xlsx",
    legal_entity_name: "EasyRewardz", brand: "EasyRewardz",
  },
  {
    file: "datatemplate/FASTSURANCE_Standardized_MIS_Template_Final.xlsx",
    legal_entity_name: "Fastsurance Consultants Private Limited", brand: "Fastsurance",
  },
  {
    file: "datatemplate/FINBOX_Standardized_MIS_Template (2).xlsx",
    legal_entity_name: "Moshpit Technologies Private Limited", brand: "FinBox",
  },
  {
    file: "datatemplate/FUNDAMENTO_Standardized_MIS_Template.xlsx",
    legal_entity_name: "Fundamento", brand: "Fundamento",
  },
  {
    file: "datatemplate/GrayQuest_Mastersheet_Template.xlsx",
    legal_entity_name: "GrayQuest Education Finance", brand: "GrayQuest",
  },
  {
    file: "standardized/Leegality_Standardized_MIS_Template.xlsx",
    legal_entity_name: "Grey Swift Private Limited", brand: "Leegality",
  },
  {
    file: "datatemplate/MULTIPL_Mastersheet_Template.xlsx",
    legal_entity_name: "MULTIPL", brand: "MULTIPL",
  },
  {
    file: "datatemplate/Riskcovry_Mastersheet_Template.xlsx",
    legal_entity_name: "Riskcovry", brand: "Riskcovry",
  },
  {
    file: "standardized/KnightFinTech_Standardized_MIS_Template.xlsx",
    legal_entity_name: "Knight FinTech", brand: "Knight FinTech",
  },
  {
    file: "standardized/Traqcheck_Standardized_MIS_Template.xlsx",
    legal_entity_name: "Traqcheck", brand: "Traqcheck",
  },
  {
    file: "standardized/Castler_Standardized_MIS_Template.xlsx",
    legal_entity_name: "Ncome Tech Solutions Private Limited", brand: "Castler",
  },
];

const results = [];

for (const c of COMPANIES) {
  const xlsxPath = path.resolve(misRoot, c.file);
  try {
    const wb = XLSX.readFile(xlsxPath);
    const parsed = parseWorkbook(wb);
    const info = parseCompanyInfo(wb);
    const detected = detectCompanyConfig(Object.keys(parsed.kpis));
    if (!detected) {
      throw new Error(`Could not detect company config for ${c.brand}`);
    }

    const ds = buildDataset(parsed, info, detected);
    const payload = buildReportContent(ds, {
      legalEntityName: c.legal_entity_name,
      brand: c.brand,
    });

    const safeName = c.brand.replace(/[^a-zA-Z0-9 _-]/g, '').trim().replace(/\s+/g, '_') || 'Report';
    const outPath = path.join(outDir, `${safeName}_Report.pptx`);
    const legacyUnderscore = path.join(outDir, `${c.brand.replace(/[^a-zA-Z0-9_-]+/g, '_')}_Report.pptx`);
    if (legacyUnderscore !== outPath && fs.existsSync(legacyUnderscore)) {
      try { fs.unlinkSync(legacyUnderscore); } catch {}
    }
    const tempJson = path.join(outDir, `temp_${safeName}.json`);

    fs.writeFileSync(tempJson, JSON.stringify(payload, null, 2), 'utf-8');

    try {
      try {
        execFileSync('py', [scriptPy, tempJson, outPath], { stdio: 'pipe' });
      } catch {
        execFileSync('python', [scriptPy, tempJson, outPath], { stdio: 'pipe' });
      }
    } finally {
      if (fs.existsSync(tempJson)) fs.unlinkSync(tempJson);
    }

    results.push({
      brand: c.brand,
      status: 'OK',
      periods: payload.period_labels,
      kpiRows: payload.kpi_section?.rows?.length || 0,
      finRows: payload.financial_section?.rows?.length || 0,
      file: path.basename(outPath),
    });
  } catch (err) {
    results.push({
      brand: c.brand,
      status: `FAILED: ${err.message}`,
    });
  }
}

for (const r of results) {
  if (r.status === 'OK') {
    console.log(`${r.brand.padEnd(16)} OK -> ${r.file.padEnd(25)} | periods: ${JSON.stringify(r.periods)} | KPI rows: ${r.kpiRows} | Fin rows: ${r.finRows}`);
  } else {
    console.error(`${r.brand.padEnd(16)} ${r.status}`);
  }
}
