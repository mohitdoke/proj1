import React, { useState, useMemo, useCallback, useEffect } from "react";
import {
  LineChart, Line, BarChart, Bar, ComposedChart, Cell, ReferenceLine,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend
} from "recharts";
import {
  ChevronDown, TrendingUp, TrendingDown, Minus, Upload, FileSpreadsheet, Info, X, RefreshCw,
  AlertTriangle, Rocket, Building2, Target, Users, Store, Layers, ShieldCheck
} from "lucide-react";
import * as XLSX from "xlsx";

/* ============================================================
   PARSER — reads the "Monthly Data" sheet of the rolling
   mastersheet template. Row 1 = KPI label + one date per month
   column. Any number of months, in chronological order.
   ============================================================ */
const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function monthLabel(y, m) { return `${MONTH_NAMES[m - 1]}'${String(y).slice(2)}`; }

function excelSerialToDate(v) {
  const epoch = new Date(Date.UTC(1899, 11, 30));
  return new Date(epoch.getTime() + v * 86400000);
}

function parseWorkbook(wb) {
  const sheetName = wb.SheetNames.find(n => /monthly data/i.test(n)) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!rows.length) throw new Error("Sheet is empty.");

  const header = rows[0];
  const monthCols = [];
  for (let c = 1; c < header.length; c++) {
    const v = header[c];
    let d = null;
    if (v instanceof Date) d = v;
    else if (typeof v === "number") d = excelSerialToDate(v);
    else if (typeof v === "string" && v.trim()) {
      const parsed = new Date(v);
      if (!isNaN(parsed)) d = parsed;
    }
    if (d) monthCols.push({ col: c, y: d.getFullYear(), m: d.getMonth() + 1 });
  }
  if (!monthCols.length) throw new Error(`No date columns found in row 1 of "${sheetName}". Row 1 needs one date per month column.`);

  const months = monthCols.map(mc => ({ y: mc.y, m: mc.m, label: monthLabel(mc.y, mc.m), key: `${mc.y}-${String(mc.m).padStart(2, "0")}` }));

  const kpis = {};
  let headcount = null;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const label = row && row[0] != null ? String(row[0]).trim() : "";
    if (!label) continue;
    // A row is a genuine KPI row if every month-column cell is either a number or
    // blank. If any month cell holds text (e.g. a stray "Source / units note" row
    // some workbooks tuck under a month column), treat the whole row as metadata
    // and skip it — otherwise it would surface as a bogus all-N/A "KPI" card.
    let hasText = false;
    const values = monthCols.map(mc => {
      const v = row[mc.col];
      if (typeof v === "number" && !isNaN(v)) return v;
      if (v != null && String(v).trim() !== "") hasText = true;
      return null;
    });
    if (hasText) continue;
    if (/^headcount$/i.test(label)) headcount = values;
    else kpis[label] = values;
  }
  if (!Object.keys(kpis).length) throw new Error(`No KPI rows found under row 1 in "${sheetName}".`);

  return { sheetName, months, kpis, headcount };
}

/* ============================================================
   COMPANY INFO — optional "Company Info" sheet. Simple Field |
   Value rows (no header). Every field is optional; anything
   missing is simply not rendered. This is what makes the
   Business Description section and the masthead company name
   work for any company that uploads a sheet.
   ============================================================ */
function parseCompanyInfo(wb) {
  const sheetName = wb.SheetNames.find(n => /company\s*info/i.test(n));
  if (!sheetName) return null;
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const map = {};
  rows.forEach(row => {
    if (!row || row[0] == null) return;
    const key = String(row[0]).trim();
    const val = row[1] != null ? String(row[1]).trim() : "";
    if (key && val) map[key] = val;
  });
  if (!Object.keys(map).length) return null;

  const companyName = map["Company Name"] || null;
  const description = map["Business Description"] || null;
  const tags = [1, 2, 3, 4].map(i => map[`Tag ${i}`]).filter(Boolean);
  const scaleMetrics = [1, 2, 3].map(i => {
    const label = map[`Scale Metric ${i} Label`];
    const value = map[`Scale Metric ${i} Value`];
    return label && value ? { label, value } : null;
  }).filter(Boolean);
  const strategicNote = map["Strategic Note Value"]
    ? {
        label: map["Strategic Note Label"] || "Strategic note",
        value: map["Strategic Note Value"],
        sub: map["Strategic Note Sub"] || "",
      }
    : null;

  if (!companyName && !description && !tags.length && !scaleMetrics.length && !strategicNote) return null;
  return { companyName, description, tags, scaleMetrics, strategicNote };
}

/* ============================================================
   NEWS FEED — optional "News Feed" sheet. One row per news item,
   flat columns. Populated by a separate, independent research
   pipeline (not the Excel financial pipeline) — see the sheet's
   Read Me notes. Every item must already carry a real source;
   there is no fabrication path here, only parsing of whatever
   the sheet contains.
   ============================================================ */
const NEWS_COLUMNS = [
  "Title", "Summary", "Category", "Published Date", "Source Name", "Source URL", "Source Tier",
  "Secondary Source Name", "Secondary Source URL", "Tags",
];

function parseNewsSheet(wb) {
  const sheetName = wb.SheetNames.find(n => /news\s*feed/i.test(n));
  if (!sheetName) return null;
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (rows.length < 2) return [];

  const header = rows[0].map(h => (h == null ? "" : String(h).trim()));
  const col = {};
  NEWS_COLUMNS.forEach(name => { col[name] = header.indexOf(name); });

  const items = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const title = col["Title"] >= 0 ? row[col["Title"]] : null;
    if (!title || !String(title).trim()) continue;
    const get = (name) => (col[name] >= 0 && row[col[name]] != null ? String(row[col[name]]).trim() : "");
    const rawDate = col["Published Date"] >= 0 ? row[col["Published Date"]] : null;
    let publishedAt = null;
    if (rawDate instanceof Date) publishedAt = rawDate;
    else if (typeof rawDate === "number") publishedAt = excelSerialToDate(rawDate);
    else if (typeof rawDate === "string" && rawDate.trim()) {
      const d = new Date(rawDate);
      if (!isNaN(d)) publishedAt = d;
    }
    const sourceName = get("Source Name");
    const sourceUrl = get("Source URL");
    if (!sourceName || !sourceUrl) continue; // no unsourced items, ever

    items.push({
      title: String(title).trim(),
      summary: get("Summary"),
      category: get("Category") || "Company",
      publishedAt,
      sourceName,
      sourceUrl,
      sourceTier: get("Source Tier") || "N/A",
      secondarySourceName: get("Secondary Source Name") || null,
      secondarySourceUrl: get("Secondary Source URL") || null,
      tags: get("Tags") ? get("Tags").split(",").map(t => t.trim()).filter(Boolean) : [],
    });
  }
  items.sort((a, b) => (b.publishedAt?.getTime() || 0) - (a.publishedAt?.getTime() || 0));
  return items;
}

/* ============================================================
   INDUSTRY DATA — optional "Industry Data" sheet. Long/tidy
   format: Section | Item | Field | Value. Rows are grouped by
   (Section, Item) into objects, so very different card shapes
   (a snapshot metric vs. a competitor's capability matrix vs. an
   analysis note) can share one flat sheet. Same "sheet missing or
   incomplete = simply not rendered" contract as Company Info.
   ============================================================ */
function parseIndustrySheet(wb) {
  const sheetName = wb.SheetNames.find(n => /industry\s*data/i.test(n));
  if (!sheetName) return null;
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (rows.length < 2) return null;

  const header = rows[0].map(h => (h == null ? "" : String(h).trim()));
  const iSection = header.indexOf("Section"), iItem = header.indexOf("Item"),
        iField = header.indexOf("Field"), iValue = header.indexOf("Value");
  if (iSection < 0 || iItem < 0 || iField < 0 || iValue < 0) return null;

  // group into { [section]: { [item]: { [field]: value } } }
  const grouped = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const section = row[iSection] != null ? String(row[iSection]).trim() : "";
    const item = row[iItem] != null ? String(row[iItem]).trim() : "";
    const field = row[iField] != null ? String(row[iField]).trim() : "";
    const value = row[iValue] != null ? String(row[iValue]).trim() : "";
    if (!section || !item || !field || !value) continue;
    grouped[section] = grouped[section] || {};
    grouped[section][item] = grouped[section][item] || {};
    grouped[section][item][field] = value;
  }
  if (!Object.keys(grouped).length) return null;

  const asList = (section) => Object.values(grouped[section] || {});

  const overview = grouped["Overview"] || {};
  const overviewDescription = overview["Description"]?.Text || null;
  const categories = Object.entries(overview)
    .filter(([item]) => item !== "Description")
    .map(([, f]) => f)
    .filter(f => f.Name);

  const snapshot = asList("Snapshot").filter(f => f.Metric);
  const trends = asList("Trend").filter(f => f.Title);
  const competitors = Object.entries(grouped["Competitor"] || {}).map(([name, f]) => ({ name, ...f }));
  const analysis = asList("Analysis").filter(f => f.Text);
  const methodology = grouped["Methodology"]?.["Note"]?.Text || null;

  if (!overviewDescription && !categories.length && !snapshot.length && !trends.length && !competitors.length && !analysis.length) {
    return null;
  }
  return { overviewDescription, categories, snapshot, trends, competitors, analysis, methodology };
}

/* Optional "Data Refresh" sheet — two Field | Value rows telling the UI when
   the independent research pipeline last updated each of the two Phase 2
   sections. Purely informational; absence just means the timestamp isn't shown. */
function parseRefreshMeta(wb) {
  const sheetName = wb.SheetNames.find(n => /data\s*refresh/i.test(n));
  if (!sheetName) return null;
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const map = {};
  rows.forEach(row => {
    if (!row || row[0] == null) return;
    const key = String(row[0]).trim();
    const val = row[1] != null ? String(row[1]).trim() : "";
    if (key && val) map[key] = val;
  });
  if (!map["News Last Refreshed"] && !map["Industry Last Refreshed"]) return null;
  return {
    newsRefreshedAt: map["News Last Refreshed"] || null,
    industryRefreshedAt: map["Industry Last Refreshed"] || null,
  };
}

/* ============================================================
   COMPANY CONFIGURATION LAYER — this is what turns the dashboard
   into a multi-company engine instead of an Easyrewardz-only app.
   The uploaded workbook is inspected for company-specific "signal"
   row names; whichever company scores the most signal matches (at
   or above a minimum confidence threshold) supplies its config for
   this render. If nothing scores high enough, the app falls back
   to the original Easyrewardz-shaped config, which is intentionally
   generic enough to already be the safe default for an unrecognized
   company's sheet (this is exactly the behaviour the app had before
   GrayQuest was added, so nothing regresses for a third, unknown
   company later).

   A config never contains financial VALUES — only which row(s) to
   look at and how to label them. Every number on screen still comes
   from the uploaded sheet at render time.
   ============================================================ */
const COMPANY_CONFIGS = {
  easyrewardz: {
    id: "easyrewardz",
    // Company Info sheet (if present) always wins; this is only a
    // fallback so the page isn't blank when one isn't uploaded.
    defaultDescription: null,
    signals: [/retail\s*\/?\s*b2b\s*revenue/i, /banking\s*revenue/i, /campaign\s*mgmt\s*revenue/i, /campaign\s*management\s*revenue/i],
    revenueBaseKey: "Total Revenue",
    revenueLabel: "Revenue",
    priorityOrder: ["Total Revenue", "Direct Expenses", "Gross Profit", "Indirect Expenses", "EBITDA", "Net Profit"],
    pnlRevenueSubLines: null, // null = default behaviour: every "*Revenue*" row except the base
    // TEMPORARY fallback for the fee-type KPI split (Platform / SetUp /
    // Campaign Management) — this workbook splits revenue by client segment,
    // not by fee type, so Platform/SetUp genuinely have no matching row.
    // Values are fractions (fmtPct multiplies by 100). Campaign Management
    // already resolves from real MIS data via KPI_SEMANTIC_RULES, so its
    // fallback here is inert unless a future upload removes that row.
    fallbackKPIs: {
      platformRevenue: { FY24: 0.64, FY25: 0.67, FY26: 0.59, Q4FY25: 0.65, Q4FY26: 0.53 },
      setupRevenue: { FY24: 0.12, FY25: 0.09, FY26: 0.06, Q4FY25: 0.09, Q4FY26: 0.06 },
      campaignManagement: { FY24: 0.24, FY25: 0.24, FY26: 0.35, Q4FY25: 0.26, Q4FY26: 0.41 },
    },
    showForecast: true,
    layout: "easyrewardz",
  },
  grayquest: {
    id: "grayquest",
    defaultDescription: {
      companyName: "GrayQuest",
      description: "GrayQuest Education Finance Private Limited is an education-finance platform offering fee-payment " +
        "solutions for K-12 schools, colleges and universities. It enables parents and guardians to pay school and " +
        "institution fees through monthly instalments, partners directly with educational institutions, and uses its " +
        "own underwriting and credit-assessment process to provide education financing.",
      tags: ["Education FinTech", "Fee-Payment Platform", "Instalment Financing"],
      scaleMetrics: [],
      strategicNote: null,
    },
    signals: [/schools\s*loan\s*disbursals/i, /higher\s*education\s*loan\s*disbursals/i, /edtech\s*platforms?\s*loan\s*disbursals/i, /^loan\s*disbursals$/i],
    revenueBaseKey: "Net Revenue",
    revenueLabel: "Net Revenue",
    priorityOrder: ["Net Revenue", "Gross Profit", "EBITDA", "Direct Expenses", "Indirect Expenses", "Net Profit", "Loan Disbursals", "Loan Amount"],
    // Explicit allow-list: GrayQuest's "Total Revenue" / "Gross Revenue" are derived
    // totals, not additive sub-lines of Net Revenue, so the default "any *Revenue*
    // row" heuristic would misrepresent the P&L. Only these genuinely roll up.
    pnlRevenueSubLines: [/^schools\s*revenue$/i, /^higher\s*education\s*revenue$/i, /^edtech\s*platforms?\s*revenue$/i, /^ancillary\s*income$/i],
    // Rate/ratio rows — averaged per period, not summed (see buildDataset).
    rateRowMatchers: [/take\s*rate/i, /loan\s*irr/i, /cost\s*of\s*funds?/i],
    // Loan-count rows (as opposed to loan-₹-amount rows) — formatted as a plain
    // number on Key Metrics cards instead of ₹ Cr, which would otherwise divide
    // a count like 170,646 by 1e7 and show a meaningless "₹0.02 Cr".
    countRowMatchers: [/loan\s*disbursals$/i],
    kpi: {
      totalMatch: /^loan\s*disbursals$/i,
      totalLabel: "Total Disbursals (loans)",
      shareRows: [
        { label: "Schools", slug: "schools", match: /schools\s*loan\s*disbursals/i },
        { label: "Colleges", slug: "colleges", match: /higher\s*education\s*loan\s*disbursals/i },
        { label: "Edtech Platforms", slug: "edtechPlatforms", match: /edtech\s*platforms?\s*loan\s*disbursals/i },
      ],
    },
    // TEMPORARY fallback for the Schools/Colleges/Edtech share rows (fractions,
    // fmtPct multiplies by 100) — inert in practice since this workbook's
    // "Loan Disbursals" family already resolves all three from real MIS data.
    // (The total-disbursals fallback given alongside these is a ₹-Cr amount,
    // not a loan count, so it's intentionally not wired — see the comment
    // above GrayQuestKPITable's `unmatched` line.)
    fallbackKPIs: {
      schools: { FY24: 0.679, FY25: 0.658, FY26: 0.835, Q4FY25: 0.634, Q4FY26: 0.835 },
      colleges: { FY24: 0.263, FY25: 0.290, FY26: 0.055, Q4FY25: 0.218, Q4FY26: 0.055 },
      edtechPlatforms: { FY24: 0.059, FY25: 0.052, FY26: 0.110, Q4FY25: 0.148, Q4FY26: 0.110 },
    },
    showForecast: false,
    layout: "grayquest",
  },
  riskcovry: {
    id: "riskcovry",
    defaultDescription: {
      companyName: "Riskcovry",
      description: "Riskcovry is an insurance-technology (insurtech) platform providing embedded-insurance " +
        "infrastructure — enabling banks, NBFCs, fintechs and other distribution partners to design, issue and " +
        "service insurance policies through API-driven integrations. Revenue is earned through platform " +
        "subscription fees, product/commission revenue on policies placed, and one-time implementation/setup fees.",
      tags: ["Insurtech", "Embedded Insurance", "Insurance Infrastructure"],
      scaleMetrics: [],
      strategicNote: null,
    },
    // Distinctive Riskcovry row names — none of these appear in the Easyrewardz
    // or GrayQuest sheets, so a 2-of-3 hit reliably identifies this workbook.
    signals: [/^platform\s*subscription\s*revenue$/i, /^product\s*\/?\s*commission\s*revenue$/i, /^one-?time\s*\/?\s*set-?up\s*fees?$/i],
    // Per the workbook's own Read Me sheet: "Total Revenue maps to Riskcovry Net
    // Revenue". The row is literally named "Total Revenue" but represents Net
    // Revenue conceptually, so the row key and its display label differ.
    revenueBaseKey: "Total Revenue",
    revenueLabel: "Net Revenue",
    priorityOrder: [
      "Total Revenue", "Direct Expenses", "Gross Profit", "Indirect Expenses", "EBITDA", "Net Profit",
      "Platform Subscription Revenue", "Product / Commission Revenue", "One-time / Setup Fees", "Gross Revenue", "Headcount",
    ],
    // Explicit allow-list: Platform Subscription + Product/Commission + One-time
    // Setup Fees sum exactly to "Gross Revenue" in every month of this workbook
    // (verified numerically) — NOT to "Total Revenue" (the Net Revenue base used
    // everywhere else on the page). Total Revenue only equals Gross Revenue in
    // the workbook's earliest months; the two diverge later (Total Revenue can
    // be roughly an order of magnitude smaller). Listing these three as P&L
    // sub-lines under the Net Revenue heading would therefore show a "breakdown"
    // that doesn't actually add up to its own subtotal — deliberately left empty
    // instead. Gross Revenue and its components still appear as independent Key
    // Metrics cards, and the Performance Summary narrates the mix explicitly
    // against Gross Revenue (not Net Revenue) to stay accurate.
    pnlRevenueSubLines: [],
    rateRowMatchers: [],
    countRowMatchers: [],
    kpi: {
      // KPI table rows (Policy Count / GWP). Each is matched semantically —
      // never assumed to exist. If the workbook has no matching row, the row
      // renders N/A with a footnote rather than a guessed figure.
      rows: [
        { label: "Policy Count (In Cr.)", slug: "policyCount", growthSlug: "policyCountGrowth", decimals: 2, matchers: [/policy\s*count/i, /number\s*of\s*policies/i, /polic(y|ies)\s*(issued|sold|bound|written)/i] },
        { label: "GWP (In Cr.)", slug: "gwp", growthSlug: "gwpGrowth", decimals: 1, matchers: [/gross\s*written\s*premium/i, /\bgwp\b/i, /premium\s*(written|collected|volume)/i, /\bpolicy\s*premium\b/i] },
      ],
      // Revenue-mix rows used only by the Performance Summary narrative (share
      // of Net Revenue), separate from the KPI table above.
      mixRows: [
        { label: "Platform Subscription Revenue", match: /^platform\s*subscription\s*revenue$/i },
        { label: "Product / Commission Revenue", match: /^product\s*\/?\s*commission\s*revenue$/i },
        { label: "One-time / Setup Fees", match: /^one-?time\s*\/?\s*set-?up\s*fees?$/i },
      ],
    },
    // TEMPORARY fallback — this workbook has no Policy Count / GWP rows at
    // all (confirmed against both the sheet and its Read Me). Values are
    // stored in the same raw units fmtCrPlain expects (Cr value × 1e7 for
    // currency/count rows) and fractions for growth (GrowthBadge multiplies
    // by 100), so a fallback cell renders through the exact same formatting
    // call as a real one. Remove this block once the source MIS carries
    // reliable Policy Count / GWP data.
    fallbackKPIs: {
      policyCount: { FY24: 0.09e7, FY25: 0.11e7, FY26: 0.11e7, Q4FY25: 0.03e7, Q4FY26: 0.05e7 },
      policyCountGrowth: { FY24: 0.459, FY25: 0.300, FY26: -0.056, Q4FY25: 0.347, Q4FY26: 0.854 },
      gwp: { FY24: 1558.0e7, FY25: 1461.3e7, FY26: 1897.5e7, Q4FY25: 525.3e7, Q4FY26: 839.1e7 },
      gwpGrowth: { FY24: 0.698, FY25: -0.062, FY26: 0.299, Q4FY25: -0.060, Q4FY26: 0.597 },
    },
    showForecast: false,
    layout: "riskcovry",
  },
  multipl: {
    id: "multipl",
    defaultDescription: {
      companyName: "Multipl",
      description: "Multipl Fintech Solutions Private Limited (brand: Multipl) is a Save Now Buy Later (SNBL) / " +
        "goal-based savings and investment platform. Users save toward fixed financial goals and invest those " +
        "savings through market instruments or with partner brands. Multipl is a SEBI-registered investment adviser.",
      tags: ["Save Now Buy Later", "Goal-Based Investing", "SEBI-Registered Investment Adviser"],
      scaleMetrics: [],
      strategicNote: null,
    },
    // Distinctive MULTIPL row names — none appear in the other three sheets.
    signals: [/goal\s*acquisition\s*cost/i, /assets\s*under\s*advice/i, /goals\s*created/i, /target\s*value\s*of\s*goals/i, /brand\s*partners/i],
    revenueBaseKey: "Total Revenue",
    revenueLabel: "Net Revenue",
    // No "Indirect Expenses" row in this workbook — the closest concept is
    // "Total Operating Costs" (per the workbook's own Read Me: "EBITDA is
    // derived as Total Revenue minus Total Operating Costs"), so the P&L's
    // Operating Expenses line is retargeted at that row instead of forcing a
    // generic "Indirect Expenses" row that doesn't exist here.
    opexKey: "Total Operating Costs",
    priorityOrder: [
      "Total Revenue", "EBITDA", "Total Operating Costs", "Monthly Burn", "Goals Created",
      "Assets Under Advice (AUA)", "Target Value of Goals", "Signups", "KYC", "Brand Partners",
      "Repeat Goals", "Goal Acquisition Cost", "Average Goal Value", "Cash in Hand",
    ],
    // No revenue sub-lines in this workbook (a single "Total Revenue" row) —
    // left empty rather than letting the default "any *Revenue* row" fallback
    // pick up the redundant "Revenue (INR Lakhs)" row (which is excluded
    // below anyway, but this stays explicit).
    pnlRevenueSubLines: [],
    // Redundant duplicate rows of a metric already shown elsewhere, just in a
    // different unit/cadence — dropped entirely so they don't surface as a
    // second, confusingly-scaled card for the same figure (verified
    // numerically: "Revenue (INR Lakhs)" = "Total Revenue" ÷ 1e5; "Monthly
    // Burn (INR million)" = "Monthly Burn" ÷ 1e6; "GAC (Quarterly)" is a
    // separately-precomputed quarterly cut of "Goal Acquisition Cost", which
    // this dashboard already derives its own quarterly average for).
    excludeRowMatchers: [/revenue\s*\(inr\s*lakhs\)/i, /monthly\s*burn\s*\(inr\s*million\)/i, /gac\s*\(quarterly\)/i],
    // Per-unit-rupee metrics (cost per goal acquired, average value per goal)
    // — averaged across the period like a rate, formatted as a plain ₹
    // amount rather than a percentage or ₹ Cr.
    rateRowMatchers: [],
    rupeeAvgRowMatchers: [/^goal\s*acquisition\s*cost$/i, /average\s*goal\s*value/i],
    // Point-in-time balance/snapshot rows — the right "value for FY26" is the
    // last reading in FY26, not a sum or average of FY26's monthly readings.
    stockRowMatchers: [/assets\s*under\s*advice/i, /^kyc$/i, /^signups$/i, /target\s*value\s*of\s*goals/i, /cash\s*in\s*hand/i, /brand\s*partners/i],
    // "Goals Created" is a running cumulative total in this workbook (never
    // decreases) — but the dashboard's "Total Goals" KPI wants goals created
    // WITHIN the period, so it's aggregated as the delta of that running
    // total (verified: this delta matches the provided reference figures for
    // FY24/FY25/FY26/Q4FY25/Q4FY26 exactly).
    deltaRowMatchers: [/goals\s*created/i],
    // Plain-count rows — formatted as a number, not ₹ Cr.
    countRowMatchers: [/^kyc$/i, /^signups$/i, /goals\s*created/i, /repeat\s*goals/i, /brand\s*partners/i],
    // AUA and Target Value of Goals are already expressed in INR Cr by the
    // source MIS itself (per its Read Me) — must not be divided by 1e7 again.
    alreadyCrRowMatchers: [/assets\s*under\s*advice/i, /target\s*value\s*of\s*goals/i],
    kpi: {
      rows: [
        { label: "Total Downloads", slug: "totalDownloads", matchers: [/total\s*downloads/i, /^downloads$/i, /app\s*downloads/i, /cumulative\s*downloads/i], fmt: fmtNum },
        { label: "Total Goals (INR Cr)", slug: "totalGoals", matchers: [/goals\s*created/i, /^total\s*goals$/i, /goal\s*count/i, /number\s*of\s*goals/i], fmt: fmtNum },
        { label: "Amt of Goals (INR Cr)", slug: "amtOfGoals", matchers: [/target\s*value\s*of\s*goals/i, /goal\s*value/i, /amount\s*of\s*goals/i, /total\s*goal\s*value/i], fmt: fmtCrAlready },
      ],
    },
    // TEMPORARY fallback. "Total Downloads" has no MIS row at all (fallback
    // used for every period below). "Total Goals" is fully MIS-derivable
    // (delta-of-cumulative matches these exact reference figures) so this
    // entry is inert in practice. "Amt of Goals" is MIS-derivable only
    // through FY24 (the source stops there) — fallback fills FY25/FY26/
    // Q4FY25/Q4FY26 where the real row is blank. Values are in the same
    // units the row's own `fmt` expects (plain counts for fmtNum, Cr-native
    // for fmtCrAlready).
    fallbackKPIs: {
      totalDownloads: { FY24: 132373, FY25: 273234, FY26: 471221, Q4FY25: 104204, Q4FY26: 99841 },
      totalGoals: { FY24: 17087, FY25: 45467, FY26: 85146, Q4FY25: 18369, Q4FY26: 16004 },
      amtOfGoals: { FY24: 1300, FY25: 2339, FY26: 3659, Q4FY25: 2339, Q4FY26: 3659 },
    },
    showForecast: false,
    layout: "multipl",
  },
  fastsurance: {
    id: "fastsurance",
    defaultDescription: {
      companyName: "Insurance Samadhan",
      description: "FASTSURANCE Consultants Private Limited (brand: Insurance Samadhan) is a consumer insurance-grievance " +
        "redressal platform that helps policyholders register complaints against insurers and pursue resolution of " +
        "unpaid, delayed or disputed claims and mis-sold policies. Revenue is earned through registration fees paid by " +
        "consumers at the time of case intake and success-based commission fees on cases resolved in the consumer's favour.",
      tags: ["Insurtech", "Grievance Redressal", "Claims Resolution"],
      scaleMetrics: [],
      strategicNote: null,
    },
    // Distinctive FASTSURANCE row names (note the source MIS's own spelling —
    // "Registartion" — preserved verbatim, per the "Source metric labels
    // preserved from the source MIS" rule in its Read Me sheet).
    signals: [/registartion\s*fees/i, /commission\s*fees\s*\(net\s*of\s*gst\)/i, /other\s*income\/reciept/i],
    revenueBaseKey: "Total Income",
    revenueLabel: "Net Revenue",
    // This MIS has no row literally named "EBITDA": "Total Expenses" already
    // excludes Depreciation (a separate line item further down the sheet,
    // subtracted only afterward to reach "Net Profit"), so "Profit/(Loss)" —
    // Total Income minus Total Expenses — IS the EBITDA-equivalent figure by
    // construction, not a guess. Verified numerically against every sampled
    // month: Total Income − Total Expenses = Profit/(Loss) exactly.
    keyAliases: { "EBITDA": "Profit/(Loss)" },
    // "Indirect Expenses" doesn't exist as a single row here — the workbook's
    // own P&L waterfall reaches EBITDA via Total Income − Total Expenses, so
    // that's the row the Complete P&L's Operating Expenses line should show.
    opexKey: "Total Expenses",
    priorityOrder: [
      "Total Income", "Registartion Fees (Net of GST)", "Commission Fees (Net of GST)", "Other Income/Reciept",
      "Total Expenses", "EBITDA", "Net Profit", "Employee Benefit Cost", "Direct Expenses",
      "Marketing & Selling Expenses", "Research & Development Expenses", "Legal, Recruitment & Professional Expenses",
      "Office & Admin Expenses", "Finance & Interest Cost", "IT expenses", "Depreciation",
    ],
    // Verified numerically: these three rows sum exactly to Total Income in
    // every sampled month.
    pnlRevenueSubLines: [/^Registartion Fees \(Net of GST\)$/i, /^Commission Fees \(Net of GST\)$/i, /^Other Income\/Reciept$/i],
    // ALLOW-LIST — this workbook's "Monthly Data" sheet interleaves ~15
    // genuine P&L subtotal rows with well over a thousand individual
    // agent-commission and vendor-spend ledger lines (real data, but
    // transaction-level detail, never a dashboard KPI). Only the rows below
    // are surfaced; everything else (every named-individual/vendor row) is
    // silently left out of the KPI grid, P&L and charts — never summed into
    // a card that would misrepresent an individual payee's spend as a
    // company metric.
    includeRowMatchers: [
      /^Registartion Fees \(Net of GST\)$/i, /^Commission Fees \(Net of GST\)$/i, /^Other Income\/Reciept$/i,
      /^Total Income$/i, /^Employee Benefit Cost$/i, /^Direct Expenses$/i, /^Marketing & Selling Expenses$/i,
      /^Research & Development Expenses$/i, /^Legal, Recruitment & Professional Expenses$/i,
      /^Office & Admin Expenses$/i, /^Finance & Interest Cost$/i, /^IT expenses$/i, /^Total Expenses$/i,
      /^Profit\/\(Loss\)$/i, /^Depreciation$/i, /^Net Profit$/i,
    ],
    // Operational insurance KPIs the reference layout calls for (Total
    // Registrations, Total Resolved Cases, Resolved Case Value, % Resolved).
    // None of these currently exist as rows in this standardized MIS (it
    // carries only the financial P&L) — matched semantically here so that
    // the day a case-management/registrations row IS added to the sheet, it
    // is picked up automatically with no code change; until then every cell
    // below renders N/A with a footnote, per spec, rather than a guess.
    kpi: {
      rows: [
        { label: "Total Registrations", slug: "totalRegistrations", matchers: [/total\s*registrations?/i, /number\s*of\s*registrations?/i, /registrations?\s*received/i, /new\s*registrations?/i], fmt: fmtNum },
        { label: "Total Resolved Cases", slug: "totalResolvedCases", matchers: [/total\s*resolved\s*cases?/i, /cases?\s*resolved/i, /resolved\s*cases?/i], fmt: fmtNum },
        { label: "Total Resolved Case Value (In Cr.)", slug: "resolvedCaseValue", matchers: [/resolved\s*(case\s*)?value/i, /value\s*of\s*(cases?\s*)?resolved/i, /amount\s*resolved/i], fmt: (v) => fmtCrPlain(v, 2) },
        { label: "% Resolved", slug: "pctResolved", matchers: [/%\s*resolved/i, /resolution\s*rate/i, /resolved\s*%/i], fmt: fmtPct },
      ],
    },
    showForecast: false,
    layout: "fastsurance",
  },
  apexFutureLabs: {
    id: "apexFutureLabs",
    defaultDescription: {
      companyName: "Vitra.ai",
      description: "Apex Future Labs Private Limited (brand: Vitra.ai) is a Generative-AI model delivered as SaaS for " +
        "language translation across text, images, video and websites (Translate Image, Translate Text, Translate " +
        "Videos, Translate Website). Revenue is earned through customer subscriptions/usage on the platform.",
      tags: ["Generative AI SaaS", "Language Translation", "Enterprise Content Localization"],
      scaleMetrics: [],
      strategicNote: null,
    },
    // Distinctive Vitra row names from the CORRECTED standardized workbook
    // (the previously-provided file had every row's label column corrupted —
    // see git history / prior session notes — this is the fixed version,
    // verified label-by-label against its own Read Me).
    signals: [/^total\s*no\.?\s*of\s*customers?$/i, /^arpa$/i, /^p&l\s*—\s*total\s*income$/i, /^p&l\s*—\s*gross\s*profit$/i, /^p&l\s*—\s*net\s*profit$/i],
    // The workbook's own Read Me: "Column A = KPI/Metric name ... P&L rows are
    // explicitly labeled to avoid unlabeled/ambiguous rows." "P&L — Total
    // Income" is the all-in revenue line (verified: ≈ GST sales + Non-GST
    // Sales + Other Income + Interest Income, to within a single month's
    // rounding across 44 populated months) — used directly rather than
    // re-derived from its components.
    revenueBaseKey: "P&L — Total Income",
    revenueLabel: "Net Revenue",
    // Verified numerically against every populated month:
    //   P&L — Gross Profit = P&L — Total Income − P&L — COGS (exact)
    //   P&L — Net Profit   = P&L — Gross Profit − P&L — Total Expenses (exact)
    // There is no separate Depreciation/Interest/Tax line below "P&L — Net
    // Profit" in this sheet's waterfall — it IS the EBITDA-equivalent by
    // construction, not a guess. Only EBITDA is aliased (not a separate "Net
    // Profit"), per the reference spec ("Net Revenue, Growth, Gross Profit,
    // Gross Margin, EBITDA, EBITDA %" — no distinct net-profit line requested).
    keyAliases: {
      "Gross Profit": "P&L — Gross Profit",
      "EBITDA": "P&L — Net Profit",
      "Direct Expenses": "P&L — COGS",
    },
    opexKey: "P&L — Total Expenses",
    priorityOrder: [
      "P&L — Total Income", "GST sales", "Non-GST Sales", "P&L — Other Income", "P&L — Interest Income",
      "Direct Expenses", "Gross Profit", "P&L — Total Expenses", "EBITDA",
      "Total no. of customers", "ARPA", "Active Customers", "Total employees",
    ],
    // Verified: GST sales + Non-GST Sales + Other Income + Interest Income
    // reconstructs Total Income to within one month's rounding out of 44
    // populated months (residual < 3% that one month, zero every other).
    pnlRevenueSubLines: [/^GST sales$/i, /^Non-GST Sales$/i, /^P&L — Other Income$/i, /^P&L — Interest Income$/i],
    // Redundant/legacy rows dropped from the Key Metrics grid so the same
    // figure (or a broken one) never shows twice or shows nonsense:
    //  - "P&L — Particulars": a leftover template row whose "values" are
    //    literally the header dates re-typed as data, not a real KPI.
    //  - "One-time" and "MRR*"/"ARR (MRR x 12)": all three equal "P&L — Total
    //    Income" (or 12x it) in every sampled month — this workbook has no
    //    distinct recurring-vs-one-time revenue split despite the labels;
    //    showing them as separate cards would misrepresent total revenue as
    //    a distinct "MRR" metric.
    //  - "Booked revenues" duplicates "P&L — Booked Revenue" (same source,
    //    two vintages) — the P&L-prefixed one is kept.
    //  - "Gross margins": a precomputed ratio row, redundant with this
    //    dashboard's own computed Gross Margin (which uses the same verified
    //    Gross Profit / Total Income).
    //  - "Saas Revenue" / "Enterprise Revenue" / "P&L — Enterprise India" /
    //    "P&L — Enterprise Total": 3–12 populated months out of 48, abandoned
    //    early segment-tracking attempts, not a reliable ongoing metric.
    //  - "P&L — Edtech" / "P&L — Digital Media" / "P&L — Others" /
    //    "P&L — Social Media" / "P&L — Media": entirely zero/null every month.
    //  - "P&L — B2C - GST" / "P&L — B2C - Non GST" / "P&L — B2C - Total": a
    //    sub-split of the GST sales / Non-GST Sales rows already shown.
    //  - "Balance as per bank" (correct spelling, only 3 populated months) vs
    //    "Balance asper bank" (source's own typo, 36 populated months) — same
    //    concept from two vintages; the fuller one is kept under its own
    //    (as-typed) label rather than merged, since no dashboard KPI depends
    //    on it either way.
    //  - "S & M %": stored as percentage-point values (e.g. 1.85 meaning
    //    1.85%), not the 0–1 fraction this dashboard's percentage formatter
    //    expects — showing it through fmtPct would misrender ~185%. Left out
    //    rather than risk a 100x-wrong display.
    //  - "Average Recovery period for invoices": unit is ambiguous from the
    //    data alone (tiny fractional values, not a clean day-count) — left
    //    out rather than guessed at.
    //  - "Receivables due over 60 days" / "Payables due over 60 days": zero
    //    in every populated month. "Unbilled (Yet to Invoice)" duplicates
    //    "Unbilled". "Contractors": only 3 populated months.
    excludeRowMatchers: [
      /^P&L — Particulars$/i, /^One-time$/i, /^MRR\*$/i, /^ARR \(MRR x 12\)$/i, /^Booked revenues$/i, /^Gross margins$/i,
      /^Saas Revenue$/i, /^Enterprise Revenue$/i, /^P&L — Enterprise India$/i, /^P&L — Enterprise Total$/i,
      /^P&L — Edtech$/i, /^P&L — Digital Media$/i, /^P&L — Others$/i, /^P&L — Social Media$/i, /^P&L — Media$/i,
      /^P&L — B2C - GST$/i, /^P&L — B2C - Non GST$/i, /^P&L — B2C - Total$/i, /^Balance as per bank$/i,
      /^S & M %$/i, /^Average Recovery period for invoices$/i, /^Receivables due over 60 days$/i,
      /^Payables due over 60 days$/i, /^Unbilled \(Yet to Invoice\)$/i, /^Contractors$/i,
    ],
    kpi: {
      rows: [
        { label: "No. of Customers", slug: "customers", matchers: [/^total\s*no\.?\s*of\s*customers?$/i, /^total\s*number\s*of\s*customers?$/i], fmt: fmtNum },
        { label: "ARPU (in Thousands)", slug: "arpu", matchers: [/^arpa$/i, /^arpu$/i, /average\s*revenue\s*per\s*account/i, /average\s*revenue\s*per\s*user/i], fmt: fmtRupeeThousands },
      ],
    },
    // "Total no. of customers" and "Active Customers" are point-in-time
    // counts (as of period end), not something that accrues within a period —
    // read as the last value in the period, same treatment as MULTIPL/
    // GrayQuest's own point-in-time counts.
    stockRowMatchers: [/^total\s*no\.?\s*of\s*customers?$/i, /^active\s*customers?$/i, /^cash\s*in\s*hand$/i, /^total\s*trade\s*receivables$/i, /^total\s*trade\s*payables$/i, /^beginning\s*cash\s*balance$/i, /^ending\s*cash\s*balance$/i, /^balance\s*asper\s*bank$/i],
    // ARPA (Average Revenue Per Account) and CAC (Customer Acquisition Cost)
    // are per-unit averages, not sums across months — averaged like
    // GrayQuest's take rate, formatted as a plain ₹ amount.
    rupeeAvgRowMatchers: [/^arpa$/i, /^cac$/i],
    countRowMatchers: [
      /^total\s*no\.?\s*of\s*customers?$/i, /^active\s*customers?$/i, /^new\s*customers?\s*-\s*b2b$/i, /^new\s*customers?\s*-\s*b2c$/i,
      /^total\s*employees$/i, /^no\.?\s*of\s*men\s*employees$/i, /^no\.?\s*of\s*women\s*employees$/i, /^new\s*employees$/i,
    ],
    showForecast: false,
    layout: "apexFutureLabs",
  },
  leegality: {
    id: "leegality",
    defaultDescription: {
      companyName: "Leegality",
      description: "Grey Swift Private Limited (brand: Leegality) is a digital-signature and e-stamping " +
        "infrastructure platform that lets businesses execute legally valid electronic signatures (Aadhaar and " +
        "non-Aadhaar eSign, digital signature certificates) and pay stamp duty online. Revenue is earned through " +
        "per-transaction eSign/stamp fees and recurring enterprise subscription plans.",
      tags: ["Digital Signature Infrastructure", "e-Stamping", "Transactional SaaS"],
      scaleMetrics: [],
      strategicNote: null,
    },
    // Distinctive Leegality row names — product-type eSign revenue lines and
    // its stamp-ordering row, none of which appear in the other companies'
    // sheets.
    signals: [/virtual\s*sign/i, /automated\s*sign/i, /total\s*for\s*transactional\s*esign/i, /number\s*of\s*stamps\s*ordered/i, /docsigner/i],
    revenueBaseKey: "Total Revenue",
    revenueLabel: "Net Revenue",
    // Canonical aliases onto this workbook's real row names — verified
    // numerically: Gross Profit (With Interest Income) = Total Revenue −
    // Total Cost of Goods Sold; EBITDA (Including Interest Income) = that
    // Gross Profit − Total Operating Expense; Net Profit/Loss = EBITDA
    // (Including) − (Depreciation + ESOP). The "With/Including Interest
    // Income" variants are the ones aliased since Total Revenue (the
    // dashboard's revenue base) itself includes Interest Income & Other
    // Charges — pairing them keeps every derived margin internally
    // consistent with its own numerator/denominator.
    keyAliases: {
      "Gross Profit": "Gross Profit (With Interest Income)",
      "EBITDA": "EBITDA (Including Interest Income)",
      "Net Profit": "Net Profit/Loss",
      "Direct Expenses": "Total Cost of Goods Sold",
    },
    opexKey: "Total Operating Expense",
    priorityOrder: [
      "Total Revenue", "Direct Expenses", "Gross Profit", "Total Operating Expense", "EBITDA", "Net Profit",
      "Total for Transactional eSign", "Transactional- Stamp Convenience", "Subscription Realisation", "Service Fee",
      "Sale of Doc Signer Certificate", "Interest Income & Other Charges",
    ],
    // Verified numerically: these seven rows sum exactly to Total Revenue
    // (the six operating-revenue lines sum to "Total for Operating Revenue",
    // which plus Interest Income & Other Charges equals Total Revenue). The
    // more granular per-product eSign rows (Aadhar, Virtual Sign, DSC, ...)
    // are themselves the components of "Total for Transactional eSign" and
    // are deliberately left out of this list — including them alongside
    // their own subtotal would double-count the revenue mix chart/P&L.
    // NOTE: "Others" is deliberately excluded from this list even though it
    // is a genuine (if tiny — a few tens of rupees/month) revenue-mix line.
    // The source workbook reuses the exact label "Others" for a second,
    // unrelated row further down (a near-zero expense-category line in the
    // Operating Expenses section) — parseWorkbook keys KPI rows by label, so
    // the second occurrence silently overwrites the first. Rather than have
    // the dashboard show a number that might actually be either row
    // depending on sheet order, both are left out entirely (immaterial
    // either way: well under ₹100 against tens of millions of revenue).
    pnlRevenueSubLines: [
      /^Total for Transactional eSign$/i, /^Transactional-\s*Stamp Convenience$/i, /^Subscription Realisation$/i,
      /^Service Fee$/i, /^Sale of Doc Signer Certificate$/i, /^Interest Income & Other Charges$/i,
    ],
    // "m" is a leftover month-index row (values 1..12, not a KPI); the
    // Without-Interest-Income Gross Profit variant and both % rows duplicate
    // the aliased With-Interest-Income figures already shown as canonical
    // Gross Profit/Gross Margin; "Total for Operating Revenue" and "Total"
    // (Depreciation+ESOP memo) are subtotals with no standalone meaning once
    // their components are shown; "EBITDA(Excluding Interest Income)"
    // duplicates the aliased canonical EBITDA; "Avg per day eSign count" is a
    // pre-averaged rate this dashboard has no clean cross-period aggregation
    // for (unlike take-rate-style % rows) and is left out rather than summed
    // into a meaningless quarter/year total; "Others" is excluded per the
    // label-collision note above. The per-product eSign revenue rows (Aadhar,
    // Virtual Sign, Automated Sign, DSC, ...) and the granular Operating
    // Expense sub-lines (Consultants Expense, HR Expenses, Salaries, ...) are
    // real data but are themselves the components of rows already shown
    // (Total for Transactional eSign; Total Operating Expense) — excluded
    // from the generic Key Metrics grid so the same rupee isn't shown twice,
    // once inside its subtotal and once as its own tiny card.
    excludeRowMatchers: [
      /^m$/i, /^Others$/i, /^Gross Profit \(Without Interest Income\)$/i, /^Gross profit % \(Without Interest Income\)$/i,
      /^Gross profit % \(With Interest Income\)$/i, /^Total for Operating Revenue$/i, /^Total$/i,
      /^EBITDA\(Excluding Interest Income\)$/i, /^Avg per day eSign count$/i, /^Employee Compensation -ESOP$/i,
      /^Aadhar$/i, /^Virtual Sign$/i, /^Automated Sign$/i, /^DSC$/i, /^Expiration Income$/i, /^Payment Collect$/i,
      /^DocSigner$/i, /^Doc Approval$/i, /^Face Match$/i, /^Quick Sign$/i, /^Visual Sign$/i, /^Smart Liveliness$/i,
      /^Whatsapp Session$/i, /^NESL$/i,
      /^Bad Debt and Other write offs$/i, /^Consultants Expense$/i, /^Finance Cost$/i, /^HR Expenses$/i,
      /^Legal Expenses-/i, /^Office Expenses$/i, /^Operational Technology & Testing Expenses$/i,
      /^Round off expenses$/i, /^Salaries and Employee Wages$/i, /^Sales and Marketing Expense$/i,
      /^Insurance Expense$/i, /^Gratuity$/i, /^Infrastructure Expenses$/i, /^Partner's Direct Commission$/i,
      /^DPDP Outsourcing$/i, /^Cost of esign$/i, /^Cost of Doc Signer Certificate$/i, /^Stamp Processing Expense$/i,
    ],
    // True percentage rows (already fractions) — averaged across the period
    // like GrayQuest's take rate, not summed.
    rateRowMatchers: [/%\s*of\s*aadhaar\s*variants/i, /%\s*of\s*non-aadhaar\s*variants/i],
    // Point-in-time snapshot rows — last reading in the period, not a sum
    // (an "active subscription wallets" count summed across 3 months would
    // triple-count the same customers still active in month 2 and 3).
    stockRowMatchers: [
      /active\s*subscription\s*wallets/i, /monthly\s*active\s*organisational\s*wallets/i, /^mrr$/i, /^arr$/i,
      /esigns?\s*-\s*till\s*date/i, /stamps-\s*till\s*date/i, /cash\s*in\s*the\s*bank/i, /amount\s*receivable/i,
      /unrealised\s*revenue/i,
    ],
    // Plain-count rows — formatted as a number, not ₹ Cr. ("Total number of
    // eSigns" and "Number of Stamps Ordered per period" are the per-period
    // flow counts this dashboard's e-sign/stamp KPI table sums per FY/quarter;
    // the "- till date" / "as on last day" variants above are their
    // cumulative/snapshot counterparts, shown as of period-end instead.)
    countRowMatchers: [
      /number\s*of\s*invoices/i, /number\s*of\s*new\s*wallets\s*invoiced/i, /number\s*of\s*new\s*subscription\s*wallets/i,
      /total\s*number\s*of\s*esigns/i, /number\s*of\s*stamps\s*ordered\s*per\s*period/i,
      /active\s*subscription\s*wallets/i, /monthly\s*active\s*organisational\s*wallets/i,
      /esigns?\s*-\s*till\s*date/i, /stamps-\s*till\s*date/i,
    ],
    kpi: {
      rows: [
        // Anchored (not a loose substring match): the workbook also carries
        // "Number of esigns - till date" (a cumulative snapshot, handled via
        // stockRowMatchers below) and an unanchored /number of e-?signs/
        // pattern would match that row too — alphabetically ahead of "Total
        // number of eSigns" in an unordered sort, so it would silently win
        // and this KPI would show the wrong (cumulative, not per-period) figure.
        { label: "No. of e-signs", slug: "esigns", growthSlug: "esignsGrowth", matchers: [/^total\s*number\s*of\s*esigns$/i, /^number\s*of\s*e-?signs$/i] },
        { label: "No. of stamps", slug: "stamps", growthSlug: "stampsGrowth", matchers: [/number\s*of\s*stamps\s*ordered\s*per\s*period/i, /^number\s*of\s*stamps$/i] },
      ],
      // Point-in-time, no YoY growth row in the reference layout.
      stockRow: { label: "No. of total subscription accounts", slug: "subscriptionAccounts", matchers: [/active\s*subscription\s*wallets/i, /total\s*subscription\s*accounts/i] },
    },
    showForecast: false,
    layout: "leegality",
  },
  finbox: {
    id: "finbox",
    defaultDescription: {
      companyName: "FinBox",
      description: "MOSHPIT Technologies Private Limited (brand: FinBox) is an API-first embedded-lending and " +
        "credit-infrastructure platform — risk intelligence, a low-code SDK, Device Connect, underwriting and " +
        "collections workflows and rule-engine applications that let banks, NBFCs and digital platforms embed " +
        "lending into their own products.",
      tags: ["Embedded Lending", "Credit Infrastructure", "API-First B2B SaaS"],
      scaleMetrics: [],
      strategicNote: null,
    },
    // Distinctive FinBox row names — none appear in any other company's sheet.
    signals: [/^aws\s*expenses$/i, /^software\s*expenses$/i, /-ebitda\s*excl\.?\s*exceptional\s*items/i, /^sg&a\s*expenses$/i, /^-adj\s*pbt$/i],
    // Per the workbook's own Read Me & spec: "Net Revenue" is the requested
    // label. Verified numerically: Net Revenue = Gross Revenue = Total
    // Revenue in every sampled month (three names for the same figure in
    // this sheet) — Net Revenue is used as the single canonical row; the
    // other two are excluded below rather than shown as duplicate cards.
    revenueBaseKey: "Net Revenue",
    revenueLabel: "Net Revenue",
    // "Gross Margin" is a literal row in this sheet (a currency figure
    // despite the name) — verified exactly: Gross Margin = Total Revenue −
    // AWS Expenses − Software Expenses, every sampled month. Aliased onto the
    // canonical "Gross Profit" the rest of the dashboard already knows how to
    // render (and margin against Net Revenue).
    // EBITDA: this sheet carries TWO vintages that genuinely diverge in their
    // overlap window (verified — e.g. Mar-2025: old "EBITDA" = -0.805 Cr vs
    // new "-EBITDA excl. exceptional items" = -13.337 Cr, an exceptional-item
    // swing). Per the Read Me's own consolidation rule ("For overlapping
    // periods, the latest/revised source is used when it contains a nonblank
    // value"), the newer, revised vintage is preferred whenever it has a
    // reading; the older "EBITDA" row only fills in the early months
    // (Apr-2020–Mar-2022) the newer vintage doesn't cover at all.
    keyAliases: {
      "Gross Profit": "Gross Margin",
      "EBITDA": ["-EBITDA excl. exceptional items", "EBITDA"],
    },
    priorityOrder: [
      "Net Revenue", "AWS Expenses", "Software Expenses", "Gross Profit", "S&M Expenses",
      "Employee Benefit Expenses", "SG&A Expenses", "Other Expenses", "EBITDA",
      "-Recurring Revenue", "-One time Revenue",
    ],
    // Verified: "-One time Revenue" + "-Recurring Revenue" sum exactly to Net
    // Revenue in every month both are populated (the newer-vintage months
    // only — older months have no revenue-mix split, left blank rather than
    // guessed).
    pnlRevenueSubLines: [/^-One time Revenue$/i, /^-Recurring Revenue$/i],
    // "Gross Revenue"/"Total Revenue" duplicate Net Revenue exactly (see
    // above). "(+) Other Income", "(-) Credits/Discounts", "(-) GST/TDS" are
    // near-always-zero adjustment lines already folded into Net Revenue,
    // not a standalone metric. "YoY Growth" is a literal precomputed row —
    // excluded so the dashboard's own consistently-defined GrowthBadge (used
    // by every other company) is the one YoY figure shown, not a second,
    // differently-defined number. "GM (%)" / "EBITDA (%)" / "-Adj PBT%" are
    // precomputed ratios redundant with this dashboard's own computed
    // margins (verified GM (%) matches Gross Profit/Net Revenue exactly).
    excludeRowMatchers: [
      /^Gross Revenue$/i, /^Total Revenue$/i, /^\(\+\) Other Income$/i, /^\(-\) Credits\/Discounts$/i,
      /^\(-\) GST\/TDS$/i, /^YoY Growth$/i, /^GM \(%\)$/i, /^EBITDA \(%\)$/i, /^-Adj PBT%$/i,
    ],
    // Source MIS is labelled INR Cr throughout (per its Read Me) — every
    // currency row must skip fmtCr's /1e7 (which assumes raw rupees).
    alreadyCrRowMatchers: [
      /^Net Revenue$/i, /^Gross Profit$/i, /^S&M Expenses$/i, /^Employee Benefit Expenses$/i, /^SG&A Expenses$/i,
      /^Other Expenses$/i, /^EBITDA$/i, /^Finance Costs$/i, /^Depreciation & Amortization Costs$/i,
      /^Profit before exceptional and extraordinary items and tax/i, /^Deferred tax$/i, /^Profit$/i,
      /^-One time Revenue$/i, /^-Recurring Revenue$/i, /^\(-\)Exceptional Items\/Credit notes$/i, /^\+ Other Income$/i,
      /^-Adj PBT$/i, /^Profit\/\(Loss\)$/i,
    ],
    // Operational KPIs the reference layout calls for (Embedded Finance %,
    // Device Connect %, Bank Connect %, Bureau Connect %, MarketX/Sentinel %)
    // — none of these exist as rows in this standardized MIS (confirmed
    // against every one of its 30 rows); matched semantically so a future
    // upload picks them up automatically, but every cell renders N/A today
    // rather than a guessed figure, per spec.
    kpi: {
      rows: [
        { label: "Embedded Finance %", slug: "embeddedFinance", matchers: [/embedded\s*finance/i], fmt: fmtPct },
        { label: "Device Connect %", slug: "deviceConnect", matchers: [/device\s*connect/i], fmt: fmtPct },
        { label: "Bank Connect %", slug: "bankConnect", matchers: [/bank\s*connect/i], fmt: fmtPct },
        { label: "Bureau Connect %", slug: "bureauConnect", matchers: [/bureau\s*connect/i], fmt: fmtPct },
        { label: "MarketX / Sentinel %", slug: "marketXSentinel", matchers: [/market\s*x/i, /sentinel/i], fmt: fmtPct },
      ],
    },
    showForecast: false,
    layout: "finbox",
  },
  fundamento: {
    id: "fundamento",
    defaultDescription: {
      companyName: "Fundamento",
      description: "Fundamento is an AI-powered Voice Agent solution for contact centres — an API-first, no-code, " +
        "multilingual platform deployable on cloud or on-prem, providing real-time agent guidance and conversational " +
        "intelligence. Revenue is earned on a per-pulse (usage-based) pricing model.",
      tags: ["Conversational AI", "Contact Centre Voice Agents", "Usage-Based SaaS"],
      scaleMetrics: [],
      strategicNote: null,
    },
    // Distinctive Fundamento row names — "pulse" doesn't appear in any other
    // company's sheet (note the source's own spelling, "Cummulative", with
    // two Ms — preserved verbatim; the regex below matches either spelling).
    signals: [/cum+ulative\s*pulses/i, /^pulse\s*per\s*month$/i, /^revenue\s*per\s*pulse$/i],
    // The source row is literally named "Total Revenue"; the reference spec
    // calls it "Gross Revenue" — same underlying figure, dashboard-facing
    // label only.
    revenueBaseKey: "Total Revenue",
    revenueLabel: "Gross Revenue",
    // This workbook's own header row has a data-quality bug: the year field
    // is incremented one year too early for each Oct/Nov/Dec run, then
    // corrected again the following January (verified: columns read
    // Jan..Sep-2025, Oct..Dec-2026, Jan..Mar-2026, Apr-2026.. — i.e. every
    // Oct/Nov/Dec is mislabeled a year ahead of its true position in the
    // sequence). Left uncorrected, FY/quarter grouping would scatter this
    // sheet's real monthly figures into the wrong fiscal years. See
    // `forceConsecutiveMonths` in buildDataset — every column's date is
    // reconstructed as a strict one-month step from the (correctly-labeled)
    // first column instead of trusting each column's own header cell.
    forceConsecutiveMonths: true,
    // This sheet has only 3 P&L rows: Total Revenue, Total Cost, EBITDA — no
    // separate COGS/Opex split. Verified numerically across every populated
    // month: EBITDA = Total Revenue − Total Cost, exactly. There is therefore
    // no reliable, distinct "Gross Profit" figure to show here (it would be
    // identical to EBITDA, falsely implying two different profitability
    // levels that don't actually exist in this MIS) — Gross Margin is
    // intentionally left as N/A rather than aliased to EBITDA. "Total Cost"
    // is instead wired as the Operating Expenses line so the Complete P&L
    // reflects this sheet's real (flatter) structure.
    opexKey: "Total Cost",
    priorityOrder: ["Total Revenue", "Total Cost", "EBITDA", "Cummulative Pulses", "Pulse per month", "Revenue Per Pulse"],
    pnlRevenueSubLines: [],
    // "Cummulative Pulses" resets to zero every April (verified: it's a
    // fiscal-year-to-date running total, not a lifetime counter) — so the
    // last reading in a period is exactly right for both views: for a full
    // FY it's the FY's total pulse count; for Q4 specifically it's also the
    // full-year total (since the FY-to-date counter has, by Q4, accumulated
    // the whole year) — which is why the reference layout asks for this KPI
    // "across FY25/FY26/Q4FY25/Q4FY26" and no other quarters. Read as the
    // last value in the period, not summed (summing would double-count).
    stockRowMatchers: [/cum+ulative\s*pulses/i],
    countRowMatchers: [/^pulse\s*per\s*month$/i, /cum+ulative\s*pulses/i],
    // Revenue Per Pulse is a per-unit average (verified: = Total Revenue ÷
    // Pulse per month, exactly, every populated month) — averaged across a
    // period like a rate, formatted as a plain ₹ amount, not summed.
    rupeeAvgRowMatchers: [/^revenue\s*per\s*pulse$/i],
    kpi: {
      rows: [
        { label: "Cumulative Pulses", slug: "cumulativePulses", matchers: [/cum+ulative\s*pulses/i], fmt: fmtNum },
      ],
    },
    showForecast: false,
    layout: "fundamento",
  },
};

function detectCompanyConfig(kpiKeys) {
  let best = null, bestScore = 0;
  Object.values(COMPANY_CONFIGS).forEach(cfg => {
    const score = cfg.signals.reduce((s, re) => s + (kpiKeys.some(k => re.test(k)) ? 1 : 0), 0);
    if (score > bestScore) { bestScore = score; best = cfg; }
  });
  // Require at least 2 distinct signal hits before committing to a company-specific
  // template — a single coincidental match (e.g. a generic "Revenue" row) shouldn't
  // be enough to mis-identify an unrelated company's sheet.
  return bestScore >= 2 ? best : COMPANY_CONFIGS.easyrewardz;
}

function buildFYGroups(months) {
  const order = [];
  const byEndYear = {};
  months.forEach((mo, i) => {
    const fyEnd = mo.m >= 4 ? mo.y + 1 : mo.y;
    if (!byEndYear[fyEnd]) { byEndYear[fyEnd] = []; order.push(fyEnd); }
    byEndYear[fyEnd].push(i);
  });
  return order.map(fyEnd => {
    const idxs = byEndYear[fyEnd];
    const first = months[idxs[0]], last = months[idxs[idxs.length - 1]];
    const partial = idxs.length < 12;
    return {
      key: `FY${String(fyEnd).slice(2)}`,
      label: `FY${String(fyEnd).slice(2)}${partial ? " (partial)" : ""}`,
      sub: `${first.label}–${last.label}`,
      fyEnd,
      partial,
      idxs,
    };
  });
}

/* Financial-year quarters: Q1=Apr–Jun, Q2=Jul–Sep, Q3=Oct–Dec, Q4=Jan–Mar */
function buildQuarterGroups(months) {
  const order = [];
  const byKey = {};
  months.forEach((mo, i) => {
    const fyEnd = mo.m >= 4 ? mo.y + 1 : mo.y;
    const monthInFY = mo.m >= 4 ? mo.m - 3 : mo.m + 9; // 1..12, Apr = 1
    const qNum = Math.ceil(monthInFY / 3);
    const k = `${fyEnd}-Q${qNum}`;
    if (!byKey[k]) { byKey[k] = []; order.push(k); }
    byKey[k].push(i);
  });
  return order.map(k => {
    const [fyEndStr, qStr] = k.split("-Q");
    const fyEnd = Number(fyEndStr), qNum = Number(qStr);
    const idxs = byKey[k];
    const first = months[idxs[0]], last = months[idxs[idxs.length - 1]];
    const complete = idxs.length === 3;
    return {
      key: `Q${qNum}FY${String(fyEnd).slice(2)}`,
      label: `Q${qNum}FY${String(fyEnd).slice(2)}${complete ? "" : " (partial)"}`,
      sub: `${first.label}–${last.label}`,
      fyEnd,
      qNum,
      complete,
      idxs,
    };
  });
}

/* Universal KPI ordering only — company-specific revenue sub-lines (e.g. a
   particular company's "Banking Revenue" row) aren't listed here on purpose,
   so any company's sheet sorts sensibly without code changes; they fall back
   to alphabetical order after these core lines. */
const PRIORITY_ORDER = [
  "Total Revenue", "Direct Expenses", "Gross Profit", "Indirect Expenses", "EBITDA", "Net Profit",
];

function buildDataset(parsed, companyInfo, companyConfig = COMPANY_CONFIGS.easyrewardz) {
  let { months, kpis, headcount } = parsed;
  // A small number of source workbooks have a genuine bug in their own
  // header row: the year field is incremented one year too early for a run
  // of Oct/Nov/Dec columns, then corrected again the following January
  // (verified in Fundamento's file: columns literally read Jan..Sep-2025,
  // Oct..Dec-2026, Jan..Mar-2026, Apr-2026.. — the Oct-Dec columns are
  // mislabeled a year ahead of where they actually sit in the sequence).
  // Rather than trust those broken header dates for FY/quarter grouping
  // (which would scatter real data into a nonsensical fiscal year), a
  // company config can opt into `forceConsecutiveMonths: true` — this
  // reconstructs every column's date as a strict one-month step from the
  // FIRST column (which is always correctly labeled in every file seen so
  // far), so the sheet's actual chronological, consecutive-monthly
  // structure is respected regardless of what a later column's header cell
  // happens to say. Values themselves are never touched — only which FY/
  // quarter bucket a column's values are grouped into.
  if (companyConfig.forceConsecutiveMonths && months.length) {
    const y0 = months[0].y, m0 = months[0].m;
    months = months.map((mo, i) => {
      const total = y0 * 12 + (m0 - 1) + i;
      const y = Math.floor(total / 12), m = (total % 12) + 1;
      return { y, m, label: monthLabel(y, m), key: `${y}-${String(m).padStart(2, "0")}` };
    });
  }
  const fyGroups = buildFYGroups(months);
  const qGroups = buildQuarterGroups(months);
  const revenueBaseKey = companyConfig.revenueBaseKey || "Total Revenue";
  const revenueLabel = companyConfig.revenueLabel || "Revenue";
  const priorityOrder = companyConfig.priorityOrder || PRIORITY_ORDER;
  // KEY ALIASES — a company's sheet may express a canonical dashboard concept
  // (Gross Profit, EBITDA, Net Profit, Direct Expenses) under its own real MIS
  // row name (e.g. Leegality's "Net Profit/Loss", FASTSURANCE's "Profit/(Loss)"
  // standing in for EBITDA since Depreciation is a separate line below it).
  // `keyAliases` maps { canonicalName: actualRowLabelInThisSheet } — when the
  // actual row exists, its values are copied onto the canonical key (never
  // invented, never re-derived — the exact same numbers, just addressable
  // under the name the rest of the dashboard already knows how to render) and
  // the original row name is hidden from the generic KPI-card grid so the
  // same figure never appears twice under two different labels.
  // A canonical key may also be given as an ORDERED ARRAY of candidate source
  // rows (e.g. FinBox's two-vintage EBITDA: newer "-EBITDA excl. exceptional
  // items" preferred, older "EBITDA" filling in only the months the newer
  // vintage doesn't cover) — this merges them month-by-month, taking the
  // first source in the list that has a real (non-null) reading for that
  // specific month, per the workbook's own Read Me rule ("For overlapping
  // periods, the latest/revised source is used when it contains a nonblank
  // value"). A plain string keeps working exactly as before (single source,
  // no merge) — fully backward-compatible with fastsurance/leegality's usage.
  const keyAliases = companyConfig.keyAliases || {};
  const aliasedOriginalKeys = new Set();
  const aliasCanonicalKeys = new Set();
  Object.entries(keyAliases).forEach(([canonical, actual]) => {
    const isMerge = Array.isArray(actual);
    const sources = isMerge ? actual : [actual];
    const validSources = sources.filter(s => kpis[s] !== undefined);
    if (!validSources.length) return;
    // A single-string alias only fires when the canonical name doesn't
    // already exist as its own raw row (the original, narrower behaviour —
    // unchanged for fastsurance/leegality/apexFutureLabs). A multi-source
    // MERGE always (re)computes, even when the canonical name happens to be
    // identical to one of its own source rows — e.g. FinBox's older MIS
    // vintage is itself literally named "EBITDA", the same name the merged,
    // newest-preferred series should be addressable under; without this, the
    // merge would silently no-op because "EBITDA" already "exists".
    if (isMerge || kpis[canonical] === undefined) {
      kpis[canonical] = months.map((_, i) => {
        for (const s of validSources) {
          const v = kpis[s][i];
          if (typeof v === "number") return v;
        }
        return null;
      });
      aliasCanonicalKeys.add(canonical);
    }
    // Hide every *other* source row from the generic grid (the merged
    // canonical key is what should surface) — but never hide the canonical
    // name itself, even when it's also one of the source rows.
    validSources.forEach(s => { if (s !== canonical) aliasedOriginalKeys.add(s); });
  });
  // Some workbooks carry redundant duplicate rows of a metric already shown
  // elsewhere, just in a different unit (e.g. MULTIPL's "Revenue (INR Lakhs)"
  // duplicating "Total Revenue", or "Monthly Burn (INR million)" duplicating
  // "Monthly Burn") — these are dropped entirely rather than surfaced as a
  // second, confusingly-scaled card for the same underlying figure.
  const excludeMatchers = companyConfig.excludeRowMatchers || [];
  // ALLOW-LIST — a small number of workbooks (e.g. FASTSURANCE's) interleave
  // genuine P&L subtotal rows with hundreds/thousands of ledger-detail rows
  // (per-agent commission lines, per-vendor spend lines) that are real data
  // but far too granular to ever be a dashboard KPI. `includeRowMatchers`,
  // when a company config defines it, keeps ONLY rows matching at least one
  // pattern — the opposite of excludeRowMatchers's blacklist — so a sheet
  // like that surfaces its ~15 meaningful subtotal rows instead of ~1,400
  // one-off "KPI cards", one per payee name.
  const includeMatchers = companyConfig.includeRowMatchers || null;
  const kpiKeys = Object.keys(kpis)
    .filter(k => !aliasedOriginalKeys.has(k))
    // Canonical alias targets (e.g. FASTSURANCE's "EBITDA" ← "Profit/(Loss)")
    // are deliberately synthesized by this company's own config, not raw
    // sheet noise — they always pass through, even when an include-list only
    // names the original row, or an exclude pattern would otherwise happen to
    // match the canonical name.
    .filter(k => aliasCanonicalKeys.has(k) || !excludeMatchers.some(re => re.test(k)))
    .filter(k => aliasCanonicalKeys.has(k) || !includeMatchers || includeMatchers.some(re => re.test(k)))
    .sort((a, b) => {
      const ia = priorityOrder.indexOf(a), ib = priorityOrder.indexOf(b);
      if (ia === -1 && ib === -1) return a.localeCompare(b);
      if (ia === -1) return 1;
      if (ib === -1) return -1;
      return ia - ib;
    });

  const hasRevenue = kpiKeys.includes(revenueBaseKey);
  const hasGP = kpiKeys.includes("Gross Profit");
  const hasEBITDA = kpiKeys.includes("EBITDA");
  const hasNet = kpiKeys.includes("Net Profit");

  // Rate/ratio rows (take rate, IRR, cost of funds, ...) must never be summed
  // across a quarter/year like a flow amount would be — three months of a ~20%
  // rate summed reads as ~60%, which is meaningless. These are averaged across
  // the period's non-null months instead (a plain average, not volume-weighted —
  // noted wherever it's surfaced). Only companies that actually have such rows
  // define rateRowMatchers; Easyrewardz's config leaves it empty, so this is a
  // no-op there.
  const rateMatchers = companyConfig.rateRowMatchers || [];
  const rateKeys = new Set(kpiKeys.filter(k => rateMatchers.some(re => re.test(k))));
  // Count rows (e.g. "Loan Disbursals") are a unit count, not a currency
  // amount — format as a plain number on Key Metrics cards instead of ₹ Cr,
  // which would otherwise divide a count like 170,646 by 1e7 and show a
  // meaningless "₹0.02 Cr".
  const countMatchers = companyConfig.countRowMatchers || [];
  const countKeys = new Set(kpiKeys.filter(k => countMatchers.some(re => re.test(k))));
  // Stock/balance rows (AUA, KYC-to-date, cumulative signups, cash on hand,
  // active partner count, ...) represent a point-in-time snapshot, not
  // something that accrues within a period — the correct "value for FY26" is
  // the last reading in FY26, not a sum or average of FY26's monthly readings.
  const stockMatchers = companyConfig.stockRowMatchers || [];
  const stockKeys = new Set(kpiKeys.filter(k => stockMatchers.some(re => re.test(k))));
  // Cumulative-running-total rows where the dashboard KPI actually wants the
  // period's *increase* (e.g. "goals created this FY", not "goals created
  // ever, as of this FY's end") — the period value is last-in-period minus
  // last-value-immediately-before-the-period, i.e. the delta of the running
  // total. Returns null (not a guess) when there's no reading before the
  // period to diff against, e.g. the sheet's very first period.
  const deltaMatchers = companyConfig.deltaRowMatchers || [];
  const deltaKeys = new Set(kpiKeys.filter(k => deltaMatchers.some(re => re.test(k))));
  // Rows already expressed in INR Cr by the source MIS itself (per the
  // workbook's own Read Me) — must NOT be run through fmtCr's /1e7, which
  // assumes raw-rupee input and would otherwise shrink a real "115 Cr" AUA
  // reading down to "0.00 Cr".
  const alreadyCrMatchers = companyConfig.alreadyCrRowMatchers || [];
  const alreadyCrKeys = new Set(kpiKeys.filter(k => alreadyCrMatchers.some(re => re.test(k))));
  // Per-unit rupee rows (cost per goal acquired, average value per goal, ...)
  // — averaged like a rate, but formatted as a plain ₹ amount, not a %.
  const rupeeAvgMatchers = companyConfig.rupeeAvgRowMatchers || [];
  const rupeeAvgKeys = new Set(kpiKeys.filter(k => rupeeAvgMatchers.some(re => re.test(k))));

  function sumFor(key, idxs) {
    const arr = kpis[key];
    let sum = 0, has = false;
    idxs.forEach(i => { const v = arr[i]; if (typeof v === "number") { sum += v; has = true; } });
    return has ? sum : null;
  }
  function avgFor(key, idxs) {
    const arr = kpis[key];
    const vals = idxs.map(i => arr[i]).filter(v => typeof v === "number");
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  }
  function lastNonNullInIdxs(arr, idxs) {
    for (let j = idxs.length - 1; j >= 0; j--) {
      const v = arr[idxs[j]];
      if (typeof v === "number") return v;
    }
    return null;
  }
  function stockFor(key, idxs) {
    return lastNonNullInIdxs(kpis[key], idxs);
  }
  function deltaFor(key, idxs) {
    const arr = kpis[key];
    const endVal = lastNonNullInIdxs(arr, idxs);
    if (endVal === null) return null;
    const minIdx = Math.min(...idxs);
    let beforeVal = null;
    for (let i = minIdx - 1; i >= 0; i--) {
      if (typeof arr[i] === "number") { beforeVal = arr[i]; break; }
    }
    // No reading exists before this period (e.g. the sheet's very first
    // period) — the "increase during the period" isn't reliably knowable
    // (it would either be undercounted or falsely assume a zero baseline),
    // so this is left null rather than guessed.
    if (beforeVal === null) return null;
    return endVal - beforeVal;
  }
  function avgHC(idxs) {
    if (!headcount) return null;
    const vals = idxs.map(i => headcount[i]).filter(v => typeof v === "number");
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  }
  function computeRow(groupMeta, idxs) {
    const row = { ...groupMeta };
    kpiKeys.forEach(k => {
      if (stockKeys.has(k)) row[k] = stockFor(k, idxs);
      else if (deltaKeys.has(k)) row[k] = deltaFor(k, idxs);
      else if (rateKeys.has(k) || rupeeAvgKeys.has(k)) row[k] = avgFor(k, idxs);
      else row[k] = sumFor(k, idxs);
    });
    row["Headcount"] = avgHC(idxs);
    // Guard every division on typeof === "number" for BOTH sides, not just
    // truthiness — `null / x` silently evaluates to 0 in JS, which would
    // otherwise render a misleading "0.0%" for a genuinely missing numerator
    // (e.g. a "Net Profit" row that exists but has no data yet) instead of N/A.
    const base = row[revenueBaseKey];
    const baseOk = typeof base === "number" && base !== 0;
    row["Gross Margin"] = (hasRevenue && hasGP && baseOk && typeof row["Gross Profit"] === "number") ? row["Gross Profit"] / base : null;
    row["EBITDA Margin"] = (hasRevenue && hasEBITDA && baseOk && typeof row["EBITDA"] === "number") ? row["EBITDA"] / base : null;
    row["Net Margin"] = (hasRevenue && hasNet && baseOk && typeof row["Net Profit"] === "number") ? row["Net Profit"] / base : null;
    row["Rev per Employee"] = (hasRevenue && headcount && baseOk && typeof row["Headcount"] === "number" && row["Headcount"] !== 0) ? base / row["Headcount"] : null;
    return row;
  }

  const fyData = fyGroups.map(fy => computeRow(fy, fy.idxs));
  const qData = qGroups.map(q => computeRow(q, q.idxs));

  const cardConfigs = [];
  kpiKeys.forEach(k => cardConfigs.push({
    key: k,
    label: k,
    fmt: countKeys.has(k) ? fmtNum
      : rupeeAvgKeys.has(k) ? fmtRupee
      : alreadyCrKeys.has(k) ? fmtCrAlready
      : rateKeys.has(k) ? fmtPct
      : fmtCr,
    good: "up",
    primary: k === revenueBaseKey || k === "EBITDA" || k === "Net Profit",
  }));
  if (hasRevenue && hasGP) cardConfigs.push({ key: "Gross Margin", label: "Gross Margin", fmt: fmtPct, good: "up", isMargin: true, marginOf: "Gross Profit" });
  if (hasRevenue && hasEBITDA) cardConfigs.push({ key: "EBITDA Margin", label: "EBITDA Margin", fmt: fmtPct, good: "up", isMargin: true, marginOf: "EBITDA" });
  if (hasRevenue && hasNet) cardConfigs.push({ key: "Net Margin", label: "Net Margin", fmt: fmtPct, good: "up", isMargin: true, marginOf: "Net Profit" });
  if (headcount) cardConfigs.push({ key: "Headcount", label: "Headcount (avg)", fmt: fmtNum, good: "neutral", isHeadcount: true });
  if (hasRevenue && headcount) cardConfigs.push({ key: "Rev per Employee", label: "Revenue per Employee", fmt: fmtCr, good: "up", isRevPerEmp: true });

  return {
    months, kpis, headcount, kpiKeys, fyGroups, fyData, qGroups, qData, cardConfigs, hasRevenue, hasGP, hasEBITDA, hasNet, companyInfo,
    companyId: companyConfig.id, companyConfig, revenueBaseKey, revenueLabel, pnlRevenueSubLines: companyConfig.pnlRevenueSubLines || null,
    opexKey: companyConfig.opexKey || null,
    rateKeys, stockKeys, deltaKeys, rupeeAvgKeys, alreadyCrKeys,
  };
}

function fmtCr(v) {
  if (v === null || v === undefined || Number.isNaN(v)) return "N/A";
  const cr = v / 1e7;
  const sign = cr < 0 ? "-" : "";
  const abs = Math.abs(cr);
  return `${sign}₹${abs.toFixed(abs >= 100 ? 0 : abs >= 10 ? 1 : 2)} Cr`;
}
function fmtPct(v) {
  if (v === null || v === undefined || Number.isNaN(v)) return "N/A";
  return `${(v * 100).toFixed(1)}%`;
}
function fmtNum(v) {
  if (v === null || v === undefined || Number.isNaN(v)) return "N/A";
  return Math.round(v).toLocaleString("en-IN");
}
function fmtPctSigned(v) {
  if (v === null || v === undefined || Number.isNaN(v)) return "N/A";
  return `${v >= 0 ? "+" : ""}${v.toFixed(1)}%`;
}
// Plain "value in Cr, no ₹ symbol" formatter — used where the unit is already
// spelled out in the label itself (e.g. "GWP (In Cr.)", "Policy Count (In Cr.)"),
// so repeating "₹" or "Cr" inside the cell would be redundant.
function fmtCrPlain(v, decimals = 1) {
  if (v === null || v === undefined || typeof v !== "number" || Number.isNaN(v)) return "N/A";
  const cr = v / 1e7;
  return cr.toLocaleString("en-IN", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}
// Rows a source MIS already expresses in INR Cr (per that workbook's own Read
// Me) — same ₹/Cr presentation as fmtCr, but WITHOUT the /1e7, since the raw
// value is already in crores.
function fmtCrAlready(v) {
  if (v === null || v === undefined || typeof v !== "number" || Number.isNaN(v)) return "N/A";
  const sign = v < 0 ? "-" : "";
  const abs = Math.abs(v);
  return `${sign}₹${abs.toFixed(abs >= 100 ? 0 : abs >= 10 ? 1 : 2)} Cr`;
}
// Plain per-unit rupee amount (e.g. cost per goal acquired, average value per
// goal) — too small to express in Cr, and not a percentage either.
function fmtRupee(v) {
  if (v === null || v === undefined || typeof v !== "number" || Number.isNaN(v)) return "N/A";
  return `₹${Math.round(v).toLocaleString("en-IN")}`;
}
// Same per-unit rupee value as fmtRupee, but scaled to thousands — used where
// the label itself says "(in Thousands)" (e.g. Vitra's ARPU), so the cell
// isn't showing a 4-5 digit rupee figure the label has already told the
// reader to read in '000s.
function fmtRupeeThousands(v) {
  if (v === null || v === undefined || typeof v !== "number" || Number.isNaN(v)) return "N/A";
  return `₹${(v / 1000).toLocaleString("en-IN", { minimumFractionDigits: 1, maximumFractionDigits: 1 })}`;
}

/* ============================================================
   TEMPORARY FALLBACK KPI DATA — a small number of KPIs (per
   company) aren't yet reliably derivable from the uploaded MIS.
   Each company config may carry a `fallbackKPIs: { slug: { FY24:
   n, ..., Q4FY26: n } }` map, keyed by the exact fyData/qData
   period `.key` (e.g. "FY24", "Q4FY26"). Priority is always:
   real MIS value first, this fallback only when the real value is
   null for that specific period — never the reverse. Values are
   stored in the SAME units/scale the relevant formatter expects
   (fractions for fmtPct, raw units for fmtCr/fmtCrPlain, Cr-native
   for fmtCrAlready, plain numbers for fmtNum) so a fallback cell
   renders through the exact same formatting call as a real one —
   nothing in the UI marks it as a fallback, per spec.
   This is explicitly temporary scaffolding for KPIs the MIS
   doesn't cover yet; removing a company's fallbackKPIs entry (once
   better source data exists) requires no other code change.
   ============================================================ */
function fallbackKPI(companyConfig, slug, periodKey) {
  if (!slug) return null;
  const map = companyConfig.fallbackKPIs?.[slug];
  return map && typeof map[periodKey] === "number" ? map[periodKey] : null;
}
function withFallback(companyConfig, slug, periodKey, value) {
  return typeof value === "number" ? value : fallbackKPI(companyConfig, slug, periodKey);
}

/* ============================================================
   PERIOD GROWTH — used by the Revenue & Profitability table.
   Yearly: sequential FY-over-FY (array order). Quarterly: the
   comparable quarter one FY back (Q2FY26 vs Q2FY25), not the
   prior sequential quarter — avoids misleading swings from
   seasonality. Returns a fraction (0.234 = +23.4%) or null when
   either side of the comparison is missing.
   ============================================================ */
function periodGrowth(list, idx, key, quarterly) {
  const curr = list[idx];
  if (!curr || typeof curr[key] !== "number") return null;
  const prev = quarterly
    ? list.find(o => o.qNum === curr.qNum && o.fyEnd === curr.fyEnd - 1)
    : (idx > 0 ? list[idx - 1] : null);
  if (!prev || typeof prev[key] !== "number" || prev[key] === 0) return null;
  return (curr[key] - prev[key]) / Math.abs(prev[key]);
}

function GrowthBadge({ value }) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return <span className="delta delta-flat"><Minus size={11} strokeWidth={2.5} />N/A</span>;
  }
  const isUp = value > 0.0005, isDown = value < -0.0005;
  const cls = isUp ? "delta-pos" : isDown ? "delta-neg" : "delta-flat";
  const Icon = isUp ? TrendingUp : isDown ? TrendingDown : Minus;
  return <span className={`delta ${cls}`}><Icon size={11} strokeWidth={2.5} />{fmtPctSigned(value * 100)}</span>;
}

/* ============================================================
   FORECASTING — simple least-squares trend on the most recent
   complete FY-quarters, projected forward. Used for the
   Outlook & Forecast section only; every other number on the
   page is an actual from the sheet.
   ============================================================ */
function linearRegression(points) {
  const n = points.length;
  const sumX = points.reduce((a, p) => a + p.x, 0);
  const sumY = points.reduce((a, p) => a + p.y, 0);
  const sumXY = points.reduce((a, p) => a + p.x * p.y, 0);
  const sumXX = points.reduce((a, p) => a + p.x * p.x, 0);
  const denom = n * sumXX - sumX * sumX;
  const slope = denom !== 0 ? (n * sumXY - sumX * sumY) / denom : 0;
  const intercept = (sumY - slope * sumX) / n;
  return { slope, intercept };
}

function buildQuarterlyForecast(qData, key, count = 2) {
  const complete = qData.filter(q => q.complete && typeof q[key] === "number");
  if (complete.length < 4) return null;
  const recent = complete.slice(-8);
  const points = recent.map((q, i) => ({ x: i, y: q[key] }));
  const { slope, intercept } = linearRegression(points);
  const lastX = points.length - 1;

  let cursor = { fyEnd: recent[recent.length - 1].fyEnd, qNum: recent[recent.length - 1].qNum };
  const forecastQuarters = [];
  for (let i = 1; i <= count; i++) {
    cursor = cursor.qNum === 4 ? { fyEnd: cursor.fyEnd + 1, qNum: 1 } : { fyEnd: cursor.fyEnd, qNum: cursor.qNum + 1 };
    forecastQuarters.push({
      key: `Q${cursor.qNum}FY${String(cursor.fyEnd).slice(2)}`,
      value: slope * (lastX + i) + intercept,
    });
  }

  const chartData = recent.map((q, i) => ({
    period: q.key,
    actual: q[key],
    projected: i === recent.length - 1 ? q[key] : null,
  }));
  forecastQuarters.forEach(f => chartData.push({ period: f.key, actual: null, projected: f.value }));

  return { chartData, slope, forecastQuarters };
}

function getExecStats(ds) {
  const completeQ = ds.qData.filter(q => q.complete);
  // The narrative is framed as a fiscal year-end (Q4) update, so anchor on the most
  // recent complete Q4 when one exists — falling back to the latest complete quarter
  // of any kind if the sheet hasn't reached a Q4 yet.
  const completeQ4s = completeQ.filter(q => q.qNum === 4);
  const latestQ = completeQ4s.length ? completeQ4s[completeQ4s.length - 1] : (completeQ.length ? completeQ[completeQ.length - 1] : null);
  const prevYearQ = latestQ ? ds.qData.find(q => q.qNum === latestQ.qNum && q.fyEnd === latestQ.fyEnd - 1) : null;
  const completeFY = ds.fyData.filter(f => !f.partial);
  const latestFY = completeFY.length ? completeFY[completeFY.length - 1] : null;
  const prevFYIdx = latestFY ? ds.fyData.findIndex(f => f.key === latestFY.key) - 1 : -1;
  const prevFY = prevFYIdx >= 0 ? ds.fyData[prevFYIdx] : null;
  return { latestQ, prevYearQ, latestFY, prevFY };
}

function Delta({ curr, prev, good = "up" }) {
  if (curr === null || prev === null || curr === undefined || prev === undefined || prev === 0) {
    return <span className="delta delta-flat"><Minus size={12} strokeWidth={2.5} /> N/A</span>;
  }
  const pct = ((curr - prev) / Math.abs(prev)) * 100;
  const isUp = pct > 0.05, isDown = pct < -0.05;
  const positive = good === "up" ? isUp : good === "down" ? isDown : null;
  const cls = positive === true ? "delta-pos" : positive === false ? "delta-neg" : "delta-flat";
  const Icon = isUp ? TrendingUp : isDown ? TrendingDown : Minus;
  return <span className={`delta ${cls}`}><Icon size={12} strokeWidth={2.5} />{Math.abs(pct).toFixed(1)}%</span>;
}

function MiniSpark({ data, dataKey, color }) {
  return (
    <ResponsiveContainer width="100%" height={36}>
      <LineChart data={data} margin={{ top: 2, right: 2, bottom: 2, left: 2 }}>
        <Line type="monotone" dataKey={dataKey} stroke={color} strokeWidth={1.75} dot={false} isAnimationActive={false} />
      </LineChart>
    </ResponsiveContainer>
  );
}

function monthlySeriesFor(ds, cfg, fyIdxs) {
  return fyIdxs.map(i => {
    let value = null;
    const rev = ds.kpis[ds.revenueBaseKey]?.[i];
    if (cfg.isHeadcount) value = ds.headcount?.[i] ?? null;
    else if (cfg.isRevPerEmp) value = ds.headcount?.[i] ? rev / ds.headcount[i] : null;
    else if (cfg.isMargin) value = rev ? (ds.kpis[cfg.marginOf]?.[i] ?? null) / rev : null;
    else value = ds.kpis[cfg.key]?.[i] ?? null;
    const mo = ds.months[i];
    return { month: mo.label, value };
  });
}

function KpiCard({ cfg, ds, fyIndex, expanded, onToggle }) {
  const fy = ds.fyData[fyIndex];
  const curr = fy[cfg.key];
  const prev = fyIndex > 0 ? ds.fyData[fyIndex - 1][cfg.key] : null;
  const sparkData = ds.fyData.map(f => ({ v: f[cfg.key] }));
  const isOpen = expanded === cfg.key;

  return (
    <button
      className={`kpi-card ${cfg.primary ? "kpi-card--primary" : ""} ${isOpen ? "kpi-card--active" : ""}`}
      onClick={() => onToggle(cfg.key)}
    >
      <div className="kpi-card__top">
        <span className="kpi-card__label">{cfg.label}</span>
        <ChevronDown size={16} className="kpi-card__chev" />
      </div>
      <div className="kpi-card__value">{cfg.fmt(curr)}</div>
      <div className="kpi-card__foot">
        <Delta curr={curr} prev={prev} good={cfg.good} />
        <span className="kpi-card__spark"><MiniSpark data={sparkData} dataKey="v" color="#1D4E4A" /></span>
      </div>
    </button>
  );
}

function DrillDownModal({ cfg, ds, fyIndex, onClose }) {
  useEffect(() => {
    const onKey = (e) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", onKey);
    document.body.style.overflow = "hidden";
    return () => { window.removeEventListener("keydown", onKey); document.body.style.overflow = ""; };
  }, [onClose]);

  const fy = ds.fyData[fyIndex];
  const curr = fy[cfg.key];
  const prev = fyIndex > 0 ? ds.fyData[fyIndex - 1][cfg.key] : null;
  const monthly = useMemo(() => monthlySeriesFor(ds, cfg, fy.idxs), [ds, cfg, fy]);
  const isPct = cfg.isMargin;

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-panel" onClick={(e) => e.stopPropagation()}>
        <div className="modal-panel__header">
          <div>
            <div className="modal-panel__eyebrow">{fy.label} · {fy.sub}</div>
            <div className="modal-panel__title">{cfg.label}</div>
          </div>
          <button className="modal-panel__close" onClick={onClose} aria-label="Close"><X size={18} /></button>
        </div>

        <div className="modal-panel__value-row">
          <div className="modal-panel__value">{cfg.fmt(curr)}</div>
          <Delta curr={curr} prev={prev} good={cfg.good} />
          <span className="modal-panel__vs">vs {fyIndex > 0 ? ds.fyData[fyIndex - 1].label : "—"} ({fyIndex > 0 ? cfg.fmt(prev) : "N/A"})</span>
        </div>

        <div className="drawer-subhead">History — {ds.fyData.map(f => f.label).join(" → ")}</div>
        <div className="fy-bars">
          {ds.fyData.map((f, i) => {
            const v = f[cfg.key];
            const max = Math.max(...ds.fyData.map(x => (typeof x[cfg.key] === "number" ? Math.abs(x[cfg.key]) : 0)));
            const h = v && max ? Math.max(4, (Math.abs(v) / max) * 56) : 2;
            const neg = v < 0;
            return (
              <div key={f.key} className={`fy-bar-col ${i === fyIndex ? "fy-bar-col--active" : ""}`}>
                <div className="fy-bar-track"><div className={`fy-bar ${neg ? "fy-bar--neg" : ""}`} style={{ height: `${h}px` }} /></div>
                <div className="fy-bar-label">{f.label}</div>
                <div className="fy-bar-value">{cfg.fmt(v)}</div>
              </div>
            );
          })}
        </div>

        <div className="drawer-subhead" style={{ marginTop: 20 }}>Monthly trend — {fy.label} ({fy.sub})</div>
        <ResponsiveContainer width="100%" height={180}>
          <LineChart data={monthly} margin={{ top: 8, right: 12, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#EEF0F3" vertical={false} />
            <XAxis dataKey="month" tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#E5E7EB" }} tickLine={false} />
            <YAxis tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false}
              tickFormatter={(v) => isPct ? `${(v * 100).toFixed(0)}%` : cfg.isHeadcount ? v : fmtCr(v)} width={isPct ? 34 : 56} />
            <Tooltip formatter={(v) => cfg.fmt(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #E5E7EB", borderRadius: 8 }} />
            <Line type="monotone" dataKey="value" stroke="#B08A3E" strokeWidth={2} dot={{ r: 2.5, fill: "#B08A3E" }} isAnimationActive={false} />
          </LineChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function RevenueTrendChart({ ds }) {
  if (!ds.hasRevenue) return null;
  const data = ds.months.map((mo, i) => ({ month: mo.label, revenue: ds.kpis[ds.revenueBaseKey][i] }));
  return (
    <ResponsiveContainer width="100%" height={230}>
      <LineChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#EEF0F3" vertical={false} />
        <XAxis dataKey="month" tick={{ fontSize: 9, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} interval={Math.max(0, Math.floor(data.length / 16))} axisLine={{ stroke: "#E5E7EB" }} tickLine={false} />
        <YAxis tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
        <Tooltip formatter={(v) => fmtCr(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #E5E7EB", borderRadius: 8 }} />
        <Line type="monotone" dataKey="revenue" stroke="#1D4E4A" strokeWidth={2} dot={false} isAnimationActive={false} />
      </LineChart>
    </ResponsiveContainer>
  );
}

function RevenueMixChart({ ds }) {
  const revenueLines = ds.pnlRevenueSubLines
    ? ds.kpiKeys.filter(k => ds.pnlRevenueSubLines.some(re => re.test(k)))
    : ds.kpiKeys.filter(k => /revenue/i.test(k) && k !== ds.revenueBaseKey);
  if (!revenueLines.length) return <div className="chart-empty">No revenue sub-lines to break down — only "{ds.revenueLabel}" is present.</div>;
  const colors = ["#1D4E4A", "#B08A3E", "#7C9885", "#8B95F2", "#B3492F"];
  const data = ds.fyData.map(f => {
    const row = { fy: f.label };
    revenueLines.forEach(k => { row[k] = f[k]; });
    return row;
  });
  return (
    <ResponsiveContainer width="100%" height={230}>
      <BarChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#EEF0F3" vertical={false} />
        <XAxis dataKey="fy" tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#E5E7EB" }} tickLine={false} />
        <YAxis tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
        <Tooltip formatter={(v) => fmtCr(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #E5E7EB", borderRadius: 8 }} />
        <Legend wrapperStyle={{ fontFamily: "Inter", fontSize: 11 }} />
        {revenueLines.map((k, i) => <Bar key={k} dataKey={k} stackId="a" fill={colors[i % colors.length]} radius={i === revenueLines.length - 1 ? [3, 3, 0, 0] : [0, 0, 0, 0]} />)}
      </BarChart>
    </ResponsiveContainer>
  );
}

function ProfitabilityChart({ ds }) {
  if (!ds.hasRevenue) return null;
  const data = ds.fyData.map(f => ({ fy: f.label, Revenue: f[ds.revenueBaseKey], EBITDA: f["EBITDA"], "Net Profit": f["Net Profit"] }));
  return (
    <ResponsiveContainer width="100%" height={230}>
      <ComposedChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#EEF0F3" vertical={false} />
        <XAxis dataKey="fy" tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#E5E7EB" }} tickLine={false} />
        <YAxis tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
        <Tooltip formatter={(v) => fmtCr(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #E5E7EB", borderRadius: 8 }} />
        <Legend wrapperStyle={{ fontFamily: "Inter", fontSize: 11 }} />
        <Bar dataKey="Revenue" fill="#E8E4DA" radius={[3, 3, 0, 0]} />
        {ds.hasEBITDA && <Line type="monotone" dataKey="EBITDA" stroke="#B08A3E" strokeWidth={2} dot={{ r: 3 }} />}
        {ds.hasNet && <Line type="monotone" dataKey="Net Profit" stroke="#B3492F" strokeWidth={2} dot={{ r: 3 }} />}
      </ComposedChart>
    </ResponsiveContainer>
  );
}

function MarginTrendChart({ ds }) {
  const lines = [];
  if (ds.hasRevenue && ds.hasGP) lines.push(["Gross Margin", "#1D4E4A"]);
  if (ds.hasRevenue && ds.hasEBITDA) lines.push(["EBITDA Margin", "#B08A3E"]);
  if (ds.hasRevenue && ds.hasNet) lines.push(["Net Margin", "#B3492F"]);
  if (!lines.length) return <div className="chart-empty">Need Total Revenue plus at least one of Gross Profit / EBITDA / Net Profit to compute margins.</div>;
  const data = ds.fyData.map(f => {
    const row = { fy: f.label };
    lines.forEach(([k]) => { row[k] = f[k]; });
    return row;
  });
  return (
    <ResponsiveContainer width="100%" height={230}>
      <LineChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#EEF0F3" vertical={false} />
        <XAxis dataKey="fy" tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#E5E7EB" }} tickLine={false} />
        <YAxis tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v * 100).toFixed(0)}%`} width={46} />
        <Tooltip formatter={(v) => fmtPct(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #E5E7EB", borderRadius: 8 }} />
        <Legend wrapperStyle={{ fontFamily: "Inter", fontSize: 11 }} />
        {lines.map(([k, c]) => <Line key={k} type="monotone" dataKey={k} stroke={c} strokeWidth={2} dot={{ r: 3 }} />)}
      </LineChart>
    </ResponsiveContainer>
  );
}

function HeadcountChart({ ds }) {
  if (!ds.headcount) return <div className="chart-empty">No "Headcount" row found in the sheet.</div>;
  const data = ds.fyData.map(f => ({ fy: f.label, Headcount: f["Headcount"], "Rev/Employee": f["Rev per Employee"] ? f["Rev per Employee"] / 1e5 : null }));
  return (
    <ResponsiveContainer width="100%" height={230}>
      <ComposedChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#EEF0F3" vertical={false} />
        <XAxis dataKey="fy" tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#E5E7EB" }} tickLine={false} />
        <YAxis yAxisId="left" tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} width={36} />
        <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} width={46} tickFormatter={(v) => `₹${v.toFixed(0)}L`} />
        <Tooltip formatter={(v, n) => n === "Rev/Employee" ? `₹${v.toFixed(1)} L` : Math.round(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #E5E7EB", borderRadius: 8 }} />
        <Legend wrapperStyle={{ fontFamily: "Inter", fontSize: 11 }} />
        <Bar yAxisId="left" dataKey="Headcount" fill="#E8E4DA" radius={[3, 3, 0, 0]} />
        {ds.hasRevenue && <Line yAxisId="right" type="monotone" dataKey="Rev/Employee" stroke="#1D4E4A" strokeWidth={2} dot={{ r: 3 }} />}
      </ComposedChart>
    </ResponsiveContainer>
  );
}

/* ============================================================
   NEW — Quarterly YoY chart, EBITDA turnaround + forecast, and
   Revenue trend + forecast. Actuals are solid, projections
   (simple linear trend on the last up-to-8 complete quarters)
   are dashed and clearly labelled.
   ============================================================ */
function QuarterlyRevenueChart({ ds }) {
  if (!ds.hasRevenue) return null;
  const quarters = ds.qData.slice(-8);
  if (!quarters.length) return <div className="chart-empty">No quarterly data available yet.</div>;
  const latestKey = quarters[quarters.length - 1].key;
  const data = quarters.map(q => ({
    period: q.label.replace(" (partial)", ""),
    value: q[ds.revenueBaseKey],
    isLatest: q.key === latestKey,
  }));
  return (
    <ResponsiveContainer width="100%" height={210}>
      <BarChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#EEF0F3" vertical={false} />
        <XAxis dataKey="period" tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#E5E7EB" }} tickLine={false} />
        <YAxis tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
        <Tooltip formatter={(v) => fmtCr(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #E5E7EB", borderRadius: 8 }} />
        <Bar dataKey="value" radius={[3, 3, 0, 0]}>
          {data.map((d, i) => <Cell key={i} fill={d.isLatest ? "#B08A3E" : "#1D4E4A"} />)}
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
}

function EbitdaTurnaroundChart({ ds }) {
  if (!ds.hasEBITDA) return <div className="chart-empty">No "EBITDA" row found in the sheet.</div>;
  const forecast = buildQuarterlyForecast(ds.qData, "EBITDA", 2);
  const fallback = ds.qData.filter(q => q.complete && typeof q["EBITDA"] === "number").slice(-8)
    .map(q => ({ period: q.key, actual: q["EBITDA"], projected: null }));
  const chartData = forecast ? forecast.chartData : fallback;
  if (!chartData.length) return <div className="chart-empty">Not enough complete quarters yet to chart EBITDA.</div>;
  return (
    <>
      <ResponsiveContainer width="100%" height={190}>
        <ComposedChart data={chartData} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#EEF0F3" vertical={false} />
          <XAxis dataKey="period" tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#E5E7EB" }} tickLine={false} />
          <YAxis tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
          <Tooltip formatter={(v) => (v === null ? "N/A" : fmtCr(v))} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #E5E7EB", borderRadius: 8 }} />
          <ReferenceLine y={0} stroke="#B3492F" strokeDasharray="2 2" />
          <Bar dataKey="actual" radius={[3, 3, 0, 0]}>
            {chartData.map((d, i) => <Cell key={i} fill={d.actual == null ? "transparent" : d.actual >= 0 ? "#16A34A" : "#B3492F"} />)}
          </Bar>
          {forecast && <Line type="monotone" dataKey="projected" stroke="#B08A3E" strokeWidth={2} strokeDasharray="5 5" dot={{ r: 3 }} />}
        </ComposedChart>
      </ResponsiveContainer>
      {forecast && (
        <div className="forecast-note">
          <span className="forecast-dot forecast-dot--proj" /> Dashed = linear-trend projection for {forecast.forecastQuarters.map(f => f.key).join(", ")}
        </div>
      )}
    </>
  );
}

function RevenueForecastChart({ ds }) {
  if (!ds.hasRevenue) return null;
  const forecast = buildQuarterlyForecast(ds.qData, ds.revenueBaseKey, 2);
  if (!forecast) return <div className="chart-empty">Need at least 4 complete quarters of {ds.revenueLabel} to project a trend.</div>;
  return (
    <>
      <ResponsiveContainer width="100%" height={190}>
        <ComposedChart data={forecast.chartData} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#EEF0F3" vertical={false} />
          <XAxis dataKey="period" tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#E5E7EB" }} tickLine={false} />
          <YAxis tick={{ fontSize: 10, fill: "#8891A3", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
          <Tooltip formatter={(v) => (v === null ? "N/A" : fmtCr(v))} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #E5E7EB", borderRadius: 8 }} />
          <Line type="monotone" dataKey="actual" stroke="#1D4E4A" strokeWidth={2} dot={{ r: 3 }} />
          <Line type="monotone" dataKey="projected" stroke="#B08A3E" strokeWidth={2} strokeDasharray="5 5" dot={{ r: 3 }} />
        </ComposedChart>
      </ResponsiveContainer>
      <div className="forecast-note">
        <span className="forecast-dot forecast-dot--actual" /> Actual &nbsp; <span className="forecast-dot forecast-dot--proj" /> Projected — trend line through {forecast.forecastQuarters.map(f => f.key).join(", ")}
      </div>
    </>
  );
}

/* ============================================================
   NEW — Financial tables (Revenue & Profitability, Complete
   P&L). Both reuse ds.qData / ds.fyData directly: those rows
   already sum the monthly values per period and compute margins
   from the aggregated numerator/denominator (see computeRow
   above), so no separate aggregation logic is needed here — this
   is presentation only. Both share a Quarterly/Yearly toggle.
   ============================================================ */
function PeriodToggle({ mode, onChange }) {
  return (
    <div className="period-toggle">
      <button
        className={`period-toggle__btn ${mode === "quarterly" ? "period-toggle__btn--active" : ""}`}
        onClick={() => onChange("quarterly")}
      >Quarterly</button>
      <button
        className={`period-toggle__btn ${mode === "yearly" ? "period-toggle__btn--active" : ""}`}
        onClick={() => onChange("yearly")}
      >Yearly</button>
    </div>
  );
}

function periodsFor(ds, mode) {
  return mode === "quarterly" ? ds.qData : ds.fyData;
}

/* ------------------------------------------------------------
   EXPORT TO EXCEL — every export below is built from an
   array-of-arrays of the exact strings already on screen (same
   fmtCr/fmtPct/fmtPctSigned calls, same row/period arrays), so the
   downloaded .xlsx can never disagree with what the table shows.
   Client-side only (SheetJS writeFile triggers a normal browser
   download) — no server involved, consistent with the rest of
   this app.
   ------------------------------------------------------------ */
function buildExportFilename(ds, suffix) {
  const company = (ds.companyInfo?.companyName || "Dashboard").replace(/[^a-z0-9]+/gi, "_");
  const stamp = new Date().toISOString().slice(0, 10);
  return `${company}_${suffix}_${stamp}.xlsx`;
}

function exportAoaToExcel(filename, sheetName, aoa) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const colCount = aoa.reduce((max, row) => Math.max(max, row.length), 0);
  ws["!cols"] = Array.from({ length: colCount }, (_, i) => ({ wch: i === 0 ? 26 : 15 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
  XLSX.writeFile(wb, filename);
}

function ExportButton({ onClick, label = "Export to Excel" }) {
  return (
    <button type="button" className="export-btn" onClick={onClick}>
      <FileSpreadsheet size={13} /> {label}
    </button>
  );
}

/* ============================================================
   KEY PERFORMANCE INDICATORS — revenue-mix table (business-line
   revenue as a % of Total Revenue), aggregated per FY/quarter the
   same way as every other table here (sum-of-period first, then
   divide — never an average of monthly percentages).

   Row → source-column matching is semantic, not a single literal
   string: each KPI concept carries a small synonym set (a company
   might call it "Platform Revenue", "SaaS Revenue", "Subscription
   Revenue", etc.) and is matched against every *revenue* line in
   the uploaded sheet (Total Revenue itself is excluded — it's the
   denominator, never a component). Every kpiKey can be claimed by
   at most one KPI row (first rule to match wins), so the same
   revenue line is never summed into two different KPIs.

   If, after that synonym search, a KPI still has zero matching
   revenue lines, it renders N/A rather than falling back to an
   unrelated line — e.g. this dashboard has been tested against a
   workbook that splits revenue by client segment ("Retail/B2B
   Revenue", "Banking Revenue") rather than by fee type, and there
   is no reliable way to know what fraction of a client-segment
   total is recurring platform/subscription fee vs. one-time setup
   fee vs. something else — collapsing "everything that isn't
   Campaign Management" into "Platform Revenue" would be a guess
   dressed up as a number, which is exactly what this table must
   never do.
   ============================================================ */
const KPI_SEMANTIC_RULES = [
  // Order matters: most specific / least ambiguous concept first,
  // so it claims its line(s) before broader concepts get a look.
  { label: "Campaign Management", slug: "campaignManagement", match: /campaign/i },
  { label: "Platform Revenue", slug: "platformRevenue", match: /platform|software\s*revenue|saas\s*revenue|technology\s*revenue|subscription/i },
  { label: "SetUp Revenue", slug: "setupRevenue", match: /set[\s-]?up|onboard|deploy(ment)?|implementation|integration\s*revenue/i },
];

function KeyPerformanceIndicatorsTable({ ds }) {
  const [mode, setMode] = useState("quarterly");
  if (!ds.hasRevenue) return null;
  const periods = periodsFor(ds, mode);
  const quarterly = mode === "quarterly";

  const revenueLines = ds.kpiKeys.filter(k => k !== ds.revenueBaseKey && /revenue/i.test(k));
  const claimed = new Set();
  const rows = KPI_SEMANTIC_RULES.map(rule => {
    const matchedKeys = revenueLines.filter(k => !claimed.has(k) && rule.match.test(k));
    matchedKeys.forEach(k => claimed.add(k));
    return { label: rule.label, slug: rule.slug, matchedKeys };
  });
  // A row only needs the "no matching line in this workbook" footnote if it
  // has neither a real match NOR a provided fallback value for any period.
  const unmatched = rows.filter(r => !r.matchedKeys.length && !ds.companyConfig.fallbackKPIs?.[r.slug]);

  function valueFor(p, keys) {
    let sum = 0, has = false;
    keys.forEach(k => { const v = p[k]; if (typeof v === "number") { sum += v; has = true; } });
    return has ? sum : null;
  }

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">Key Performance Indicators</div>
        <PeriodToggle mode={mode} onChange={setMode} />
      </div>
      <div className="fin-table-wrap fin-table-wrap--kpi">
        <table className="fin-table fin-table--kpi">
          <thead>
            <tr>
              <th className="fin-table__label-col">KPI</th>
              {periods.map(p => (
                <th key={p.key}>
                  {p.label.replace(" (partial)", "")}
                  {quarterly
                    ? (!p.complete ? <span className="fin-table__partial"> (partial)</span> : "")
                    : (p.partial ? <span className="fin-table__partial"> (partial)</span> : "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.label}>
                <td className="fin-table__label-col">{row.label}</td>
                {periods.map(p => {
                  let pct = null;
                  if (row.matchedKeys.length) {
                    const raw = valueFor(p, row.matchedKeys);
                    const total = p[ds.revenueBaseKey];
                    pct = (typeof raw === "number" && typeof total === "number" && total !== 0) ? raw / total : null;
                  }
                  const val = withFallback(ds.companyConfig, row.slug, p.key, pct);
                  return <td key={p.key}>{fmtPct(val)}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {unmatched.length > 0 && (
        <div className="fin-table__foot-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          {unmatched.map(r => r.label).join(" and ")} {unmatched.length > 1 ? "have" : "has"} no revenue line in this
          workbook that reliably matches that concept (by name or common synonym) — shown as N/A rather than guessed,
          since this workbook currently splits revenue by client segment, not by fee type.
        </div>
      )}
    </section>
  );
}

/* ============================================================
   GRAYQUEST KEY PERFORMANCE INDICATORS — structurally different
   from Easyrewardz's KPI table: one absolute row ("Total
   Disbursals", a loan count) plus three revenue-mix-style share
   rows (Schools / Colleges / Edtech Platforms, each a % of Total
   Disbursals). Row matching uses the semantic rules from
   COMPANY_CONFIGS.grayquest.kpi — e.g. "Colleges" resolves to
   whatever row matches "Higher Education Loan Disbursals" rather
   than requiring a row literally named "Colleges" — so a future
   GrayQuest export with slightly different row names still works,
   and if a concept genuinely has no match, it renders N/A instead
   of guessing.
   ============================================================ */
function GrayQuestKPITable({ ds }) {
  const [mode, setMode] = useState("quarterly");
  const cfg = ds.companyConfig.kpi;
  if (!cfg) return null;
  const periods = periodsFor(ds, mode);
  const quarterly = mode === "quarterly";

  const totalKey = ds.kpiKeys.find(k => cfg.totalMatch.test(k)) || null;
  const shareRows = cfg.shareRows.map(r => ({ ...r, matchedKey: ds.kpiKeys.find(k => r.match.test(k)) || null }));
  // The total row's provided fallback (if any) is a ₹-Cr disbursal-*amount*
  // figure, while this row displays a loan *count* — a genuine unit mismatch
  // with no clean fallback path, so it's deliberately left un-wired here. In
  // practice this never matters: every GrayQuest workbook seen so far has a
  // reliable "Loan Disbursals" count row, so this total is always MIS-derived.
  const unmatched = shareRows.filter(r => !r.matchedKey && !ds.companyConfig.fallbackKPIs?.[r.slug]);

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">Key Performance Indicators</div>
        <PeriodToggle mode={mode} onChange={setMode} />
      </div>
      <div className="fin-table-wrap fin-table-wrap--kpi">
        <table className="fin-table fin-table--kpi">
          <thead>
            <tr>
              <th className="fin-table__label-col">KPI</th>
              {periods.map(p => (
                <th key={p.key}>
                  {p.label.replace(" (partial)", "")}
                  {quarterly
                    ? (!p.complete ? <span className="fin-table__partial"> (partial)</span> : "")
                    : (p.partial ? <span className="fin-table__partial"> (partial)</span> : "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            <tr>
              <td className="fin-table__label-col">{cfg.totalLabel}</td>
              {periods.map(p => (
                <td key={p.key}>{totalKey ? fmtNum(p[totalKey]) : <span className="fin-table__na">N/A</span>}</td>
              ))}
            </tr>
            {shareRows.map(row => (
              <tr key={row.label}>
                <td className="fin-table__label-col">{row.label}</td>
                {periods.map(p => {
                  let pct = null;
                  if (row.matchedKey && totalKey) {
                    const raw = p[row.matchedKey];
                    const total = p[totalKey];
                    pct = (typeof raw === "number" && typeof total === "number" && total !== 0) ? raw / total : null;
                  }
                  const val = withFallback(ds.companyConfig, row.slug, p.key, pct);
                  return <td key={p.key}>{fmtPct(val)}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {!totalKey && (
        <div className="fin-table__foot-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          No row matching "Loan Disbursals" (the total-disbursal concept) was found in this workbook — every row above
          shows N/A rather than a guessed figure.
        </div>
      )}
      {totalKey && unmatched.length > 0 && (
        <div className="fin-table__foot-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          {unmatched.map(r => r.label).join(" and ")} {unmatched.length > 1 ? "have" : "has"} no matching disbursal line
          in this workbook — shown as N/A rather than guessed.
        </div>
      )}
    </section>
  );
}

/* ============================================================
   RISKCOVRY KEY PERFORMANCE INDICATORS — a 4-row table (Policy
   Count + its YoY growth, GWP + its YoY growth), distinct in shape
   from both the Easyrewardz mix-table and the GrayQuest
   total-plus-shares table, since Riskcovry's KPIs are two absolute
   metrics tracked over time rather than a total split into shares.
   Growth reuses the same periodGrowth() helper the Revenue &
   Profitability table uses (yearly = sequential FY-over-FY,
   quarterly = comparable quarter one FY back) — one consistent
   growth methodology across every company, per spec.
   ============================================================ */
function RiskcovryKPITable({ ds }) {
  const [mode, setMode] = useState("quarterly");
  const cfg = ds.companyConfig.kpi;
  if (!cfg || !cfg.rows) return null;
  const periods = periodsFor(ds, mode);
  const quarterly = mode === "quarterly";

  const rows = cfg.rows.map(r => ({
    ...r,
    matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null,
  }));
  const unmatched = rows.filter(r => !r.matchedKey && !ds.companyConfig.fallbackKPIs?.[r.slug]);

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">Key Performance Indicators</div>
        <PeriodToggle mode={mode} onChange={setMode} />
      </div>
      <div className="fin-table-wrap fin-table-wrap--kpi">
        <table className="fin-table fin-table--kpi">
          <thead>
            <tr>
              <th className="fin-table__label-col">KPI</th>
              {periods.map(p => (
                <th key={p.key}>
                  {p.label.replace(" (partial)", "")}
                  {quarterly
                    ? (!p.complete ? <span className="fin-table__partial"> (partial)</span> : "")
                    : (p.partial ? <span className="fin-table__partial"> (partial)</span> : "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <React.Fragment key={row.label}>
                <tr>
                  <td className="fin-table__label-col">{row.label}</td>
                  {periods.map(p => {
                    const raw = row.matchedKey ? p[row.matchedKey] : null;
                    const val = withFallback(ds.companyConfig, row.slug, p.key, raw);
                    return <td key={p.key}>{fmtCrPlain(val, row.decimals)}</td>;
                  })}
                </tr>
                <tr>
                  <td className="fin-table__label-col">% Growth YoY</td>
                  {periods.map((p, i) => {
                    const g = row.matchedKey ? periodGrowth(periods, i, row.matchedKey, quarterly) : null;
                    const val = withFallback(ds.companyConfig, row.growthSlug, p.key, g);
                    return <td key={p.key}><GrowthBadge value={val} /></td>;
                  })}
                </tr>
              </React.Fragment>
            ))}
          </tbody>
        </table>
      </div>
      {unmatched.length > 0 && (
        <div className="fin-table__foot-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          No row matching {unmatched.map(r => `"${r.label.replace(" (In Cr.)", "")}"`).join(" or ")} was found in this
          workbook — shown as N/A rather than a guessed figure.
        </div>
      )}
    </section>
  );
}

/* ============================================================
   MULTIPL KEY PERFORMANCE INDICATORS — three flat value rows
   (Total Downloads, Total Goals, Amt of Goals), no growth
   sub-rows (unlike Riskcovry's table). "Total Goals" reads off a
   delta-of-cumulative aggregation and "Amt of Goals" off a
   stock/last-value aggregation — both computed once in
   buildDataset (via deltaRowMatchers/stockRowMatchers/
   alreadyCrRowMatchers) so this component just displays whatever
   ds.fyData/ds.qData already computed, falling back to the
   company's provided reference data only where the MIS itself has
   no reading for that specific period.
   ============================================================ */
function MultiplKPITable({ ds }) {
  const [mode, setMode] = useState("quarterly");
  const cfg = ds.companyConfig.kpi;
  if (!cfg || !cfg.rows) return null;
  const periods = periodsFor(ds, mode);
  const quarterly = mode === "quarterly";

  const rows = cfg.rows.map(r => ({
    ...r,
    matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null,
  }));
  const unmatched = rows.filter(r => !r.matchedKey && !ds.companyConfig.fallbackKPIs?.[r.slug]);

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">Key Performance Indicators</div>
        <PeriodToggle mode={mode} onChange={setMode} />
      </div>
      <div className="fin-table-wrap fin-table-wrap--kpi">
        <table className="fin-table fin-table--kpi">
          <thead>
            <tr>
              <th className="fin-table__label-col">KPI</th>
              {periods.map(p => (
                <th key={p.key}>
                  {p.label.replace(" (partial)", "")}
                  {quarterly
                    ? (!p.complete ? <span className="fin-table__partial"> (partial)</span> : "")
                    : (p.partial ? <span className="fin-table__partial"> (partial)</span> : "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.label}>
                <td className="fin-table__label-col">{row.label}</td>
                {periods.map(p => {
                  const raw = row.matchedKey ? p[row.matchedKey] : null;
                  const val = withFallback(ds.companyConfig, row.slug, p.key, raw);
                  return <td key={p.key}>{row.fmt(val)}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {unmatched.length > 0 && (
        <div className="fin-table__foot-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          No row matching {unmatched.map(r => `"${r.label}"`).join(" or ")} was found in this workbook — shown as N/A
          rather than a guessed figure.
        </div>
      )}
    </section>
  );
}

/* ============================================================
   FASTSURANCE KEY PERFORMANCE INDICATORS — the reference layout's
   insurance-operations metrics (Total Registrations, Total
   Resolved Cases, Resolved Case Value, % Resolved). Same flat
   value-row shape as MultiplKPITable, matched semantically against
   ds.companyConfig.kpi.rows. As of the current standardized MIS,
   none of these four concepts have a matching row in the sheet
   (it currently carries only the financial P&L, no case-management
   data) — every cell therefore renders N/A with a footnote, which
   is the correct, honest behaviour per spec rather than a reason
   to omit the table. The day a registrations/resolved-case row is
   added to the sheet, this table picks it up with no code change.
   ============================================================ */
function FastsuranceKPITable({ ds }) {
  const [mode, setMode] = useState("quarterly");
  const cfg = ds.companyConfig.kpi;
  if (!cfg || !cfg.rows) return null;
  const periods = periodsFor(ds, mode);
  const quarterly = mode === "quarterly";

  const rows = cfg.rows.map(r => ({
    ...r,
    matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null,
  }));
  const unmatched = rows.filter(r => !r.matchedKey && !ds.companyConfig.fallbackKPIs?.[r.slug]);

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">Key Performance Indicators</div>
        <PeriodToggle mode={mode} onChange={setMode} />
      </div>
      <div className="fin-table-wrap fin-table-wrap--kpi">
        <table className="fin-table fin-table--kpi">
          <thead>
            <tr>
              <th className="fin-table__label-col">KPI</th>
              {periods.map(p => (
                <th key={p.key}>
                  {p.label.replace(" (partial)", "")}
                  {quarterly
                    ? (!p.complete ? <span className="fin-table__partial"> (partial)</span> : "")
                    : (p.partial ? <span className="fin-table__partial"> (partial)</span> : "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.label}>
                <td className="fin-table__label-col">{row.label}</td>
                {periods.map(p => {
                  const raw = row.matchedKey ? p[row.matchedKey] : null;
                  const val = withFallback(ds.companyConfig, row.slug, p.key, raw);
                  return <td key={p.key}>{row.fmt(val)}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {unmatched.length > 0 && (
        <div className="fin-table__foot-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          No row matching {unmatched.map(r => `"${r.label}"`).join(" or ")} was found in this workbook — shown as N/A
          rather than a guessed figure. This standardized MIS currently carries financial (P&amp;L) data only; add a
          registrations/resolved-case tracker to the sheet to unlock these.
        </div>
      )}
    </section>
  );
}

/* ============================================================
   VITRA (APEX FUTURE LABS) KEY PERFORMANCE INDICATORS — No. of
   Customers (a point-in-time count, read as the last value in the
   period — see stockRowMatchers) and ARPU (a per-customer average,
   not summed — see rupeeAvgRowMatchers). Same flat value-row shape
   as MultiplKPITable/FastsuranceKPITable. See the note on
   COMPANY_CONFIGS.apexFutureLabs — the currently-provided
   standardized workbook for this company has no usable row labels
   at all (a data-integrity issue in that file, not this dashboard),
   so both rows render N/A until a corrected file is uploaded.
   ============================================================ */
function ApexFutureLabsKPITable({ ds }) {
  const [mode, setMode] = useState("quarterly");
  const cfg = ds.companyConfig.kpi;
  if (!cfg || !cfg.rows) return null;
  const periods = periodsFor(ds, mode);
  const quarterly = mode === "quarterly";

  const rows = cfg.rows.map(r => ({
    ...r,
    matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null,
  }));
  const unmatched = rows.filter(r => !r.matchedKey && !ds.companyConfig.fallbackKPIs?.[r.slug]);

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">Key Performance Indicators</div>
        <PeriodToggle mode={mode} onChange={setMode} />
      </div>
      <div className="fin-table-wrap fin-table-wrap--kpi">
        <table className="fin-table fin-table--kpi">
          <thead>
            <tr>
              <th className="fin-table__label-col">KPI</th>
              {periods.map(p => (
                <th key={p.key}>
                  {p.label.replace(" (partial)", "")}
                  {quarterly
                    ? (!p.complete ? <span className="fin-table__partial"> (partial)</span> : "")
                    : (p.partial ? <span className="fin-table__partial"> (partial)</span> : "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.label}>
                <td className="fin-table__label-col">{row.label}</td>
                {periods.map(p => {
                  const raw = row.matchedKey ? p[row.matchedKey] : null;
                  const val = withFallback(ds.companyConfig, row.slug, p.key, raw);
                  return <td key={p.key}>{row.fmt(val)}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {unmatched.length > 0 && (
        <div className="fin-table__foot-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          No row matching {unmatched.map(r => `"${r.label}"`).join(" or ")} was found in this workbook — shown as N/A
          rather than a guessed figure.
        </div>
      )}
    </section>
  );
}

/* ============================================================
   FINBOX KEY PERFORMANCE INDICATORS — Embedded Finance % / Device
   Connect % / Bank Connect % / Bureau Connect % / MarketX-Sentinel
   %, per the reference layout. None of these exist as rows in the
   standardized MIS today (confirmed against all 30 rows) — every
   cell renders N/A with the footnote below rather than a guessed
   figure. Same flat value-row shape as MultiplKPITable/
   FastsuranceKPITable/ApexFutureLabsKPITable.
   ============================================================ */
function FinboxKPITable({ ds }) {
  const [mode, setMode] = useState("quarterly");
  const cfg = ds.companyConfig.kpi;
  if (!cfg || !cfg.rows) return null;
  const periods = periodsFor(ds, mode);
  const quarterly = mode === "quarterly";

  const rows = cfg.rows.map(r => ({
    ...r,
    matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null,
  }));
  const unmatched = rows.filter(r => !r.matchedKey && !ds.companyConfig.fallbackKPIs?.[r.slug]);

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">Key Performance Indicators</div>
        <PeriodToggle mode={mode} onChange={setMode} />
      </div>
      <div className="fin-table-wrap fin-table-wrap--kpi">
        <table className="fin-table fin-table--kpi">
          <thead>
            <tr>
              <th className="fin-table__label-col">KPI</th>
              {periods.map(p => (
                <th key={p.key}>
                  {p.label.replace(" (partial)", "")}
                  {quarterly
                    ? (!p.complete ? <span className="fin-table__partial"> (partial)</span> : "")
                    : (p.partial ? <span className="fin-table__partial"> (partial)</span> : "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.label}>
                <td className="fin-table__label-col">{row.label}</td>
                {periods.map(p => {
                  const raw = row.matchedKey ? p[row.matchedKey] : null;
                  const val = withFallback(ds.companyConfig, row.slug, p.key, raw);
                  return <td key={p.key}>{row.fmt(val)}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {unmatched.length > 0 && (
        <div className="fin-table__foot-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          No row matching {unmatched.map(r => `"${r.label}"`).join(", ")} was found in this workbook — this
          standardized MIS currently carries financial (revenue/cost/EBITDA) data only, so these operational mix
          metrics are shown as N/A above rather than estimated.
        </div>
      )}
    </section>
  );
}

/* ============================================================
   FUNDAMENTO KEY PERFORMANCE INDICATORS — Cumulative Pulses, read
   as the last value in the period (see stockRowMatchers: the
   source counter resets every April, so the FY-end/Q4-end reading
   is exactly the full-year pulse count). Same flat value-row shape
   as the other single/few-KPI companies.
   ============================================================ */
function FundamentoKPITable({ ds }) {
  const [mode, setMode] = useState("quarterly");
  const cfg = ds.companyConfig.kpi;
  if (!cfg || !cfg.rows) return null;
  const periods = periodsFor(ds, mode);
  const quarterly = mode === "quarterly";

  const rows = cfg.rows.map(r => ({
    ...r,
    matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null,
  }));
  const unmatched = rows.filter(r => !r.matchedKey && !ds.companyConfig.fallbackKPIs?.[r.slug]);

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">Key Performance Indicators</div>
        <PeriodToggle mode={mode} onChange={setMode} />
      </div>
      <div className="fin-table-wrap fin-table-wrap--kpi">
        <table className="fin-table fin-table--kpi">
          <thead>
            <tr>
              <th className="fin-table__label-col">KPI</th>
              {periods.map(p => (
                <th key={p.key}>
                  {p.label.replace(" (partial)", "")}
                  {quarterly
                    ? (!p.complete ? <span className="fin-table__partial"> (partial)</span> : "")
                    : (p.partial ? <span className="fin-table__partial"> (partial)</span> : "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.label}>
                <td className="fin-table__label-col">{row.label}</td>
                {periods.map(p => {
                  const raw = row.matchedKey ? p[row.matchedKey] : null;
                  const val = withFallback(ds.companyConfig, row.slug, p.key, raw);
                  return <td key={p.key}>{row.fmt(val)}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {unmatched.length > 0 && (
        <div className="fin-table__foot-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          No row matching {unmatched.map(r => `"${r.label}"`).join(" or ")} was found in this workbook — shown as N/A
          rather than a guessed figure.
        </div>
      )}
    </section>
  );
}

/* ============================================================
   LEEGALITY KEY PERFORMANCE INDICATORS — combines Riskcovry's
   value+YoY-growth row shape (for e-signs and stamps, both period
   flow counts) with one plain stock row (subscription accounts, a
   point-in-time snapshot with no growth row in the reference
   layout). Growth reuses the same periodGrowth() helper as every
   other company's table — sequential FY-over-FY yearly, comparable
   quarter one FY back quarterly.
   ============================================================ */
function LeegalityKPITable({ ds }) {
  const [mode, setMode] = useState("quarterly");
  const cfg = ds.companyConfig.kpi;
  if (!cfg || !cfg.rows) return null;
  const periods = periodsFor(ds, mode);
  const quarterly = mode === "quarterly";

  const rows = cfg.rows.map(r => ({
    ...r,
    matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null,
  }));
  const stockRow = cfg.stockRow ? { ...cfg.stockRow, matchedKey: ds.kpiKeys.find(k => cfg.stockRow.matchers.some(re => re.test(k))) || null } : null;
  const unmatched = [...rows, ...(stockRow ? [stockRow] : [])].filter(r => !r.matchedKey && !ds.companyConfig.fallbackKPIs?.[r.slug]);

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">Key Performance Indicators</div>
        <PeriodToggle mode={mode} onChange={setMode} />
      </div>
      <div className="fin-table-wrap fin-table-wrap--kpi">
        <table className="fin-table fin-table--kpi">
          <thead>
            <tr>
              <th className="fin-table__label-col">KPI</th>
              {periods.map(p => (
                <th key={p.key}>
                  {p.label.replace(" (partial)", "")}
                  {quarterly
                    ? (!p.complete ? <span className="fin-table__partial"> (partial)</span> : "")
                    : (p.partial ? <span className="fin-table__partial"> (partial)</span> : "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <React.Fragment key={row.label}>
                <tr>
                  <td className="fin-table__label-col">{row.label}</td>
                  {periods.map(p => {
                    const raw = row.matchedKey ? p[row.matchedKey] : null;
                    const val = withFallback(ds.companyConfig, row.slug, p.key, raw);
                    return <td key={p.key}>{fmtNum(val)}</td>;
                  })}
                </tr>
                <tr>
                  <td className="fin-table__label-col">Growth YoY</td>
                  {periods.map((p, i) => {
                    const g = row.matchedKey ? periodGrowth(periods, i, row.matchedKey, quarterly) : null;
                    const val = withFallback(ds.companyConfig, row.growthSlug, p.key, g);
                    return <td key={p.key}><GrowthBadge value={val} /></td>;
                  })}
                </tr>
              </React.Fragment>
            ))}
            {stockRow && (
              <tr>
                <td className="fin-table__label-col">{stockRow.label}</td>
                {periods.map(p => {
                  const raw = stockRow.matchedKey ? p[stockRow.matchedKey] : null;
                  const val = withFallback(ds.companyConfig, stockRow.slug, p.key, raw);
                  return <td key={p.key}>{fmtNum(val)}</td>;
                })}
              </tr>
            )}
          </tbody>
        </table>
      </div>
      {unmatched.length > 0 && (
        <div className="fin-table__foot-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          No row matching {unmatched.map(r => `"${r.label}"`).join(" or ")} was found in this workbook — shown as N/A
          rather than a guessed figure.
        </div>
      )}
    </section>
  );
}

function RevenueProfitabilityTable({ ds, title = "Revenue & Profitability" }) {
  const [mode, setMode] = useState("quarterly");
  if (!ds.hasRevenue) return null;
  const periods = periodsFor(ds, mode);
  const quarterly = mode === "quarterly";

  const rows = [
    { label: ds.revenueLabel || "Revenue", key: ds.revenueBaseKey, type: "currency" },
    { label: `${ds.revenueLabel || "Revenue"} Growth`, key: ds.revenueBaseKey, type: "growth" },
    ds.hasGP && { label: "Gross Profit", key: "Gross Profit", type: "currency" },
    ds.hasGP && { label: "Gross Margin", key: "Gross Margin", type: "percent" },
    ds.hasEBITDA && { label: "EBITDA", key: "EBITDA", type: "currency" },
    ds.hasEBITDA && { label: "EBITDA Margin", key: "EBITDA Margin", type: "percent" },
  ].filter(Boolean);

  // Cell values below are computed with the exact same helpers used to
  // render the table (fmtCr / fmtPct / periodGrowth+fmtPctSigned) — the
  // export can never drift from what's on screen because it reads the
  // same `rows`/`periods` this render pass already built.
  const handleExport = () => {
    const header = ["Metric", ...periods.map(p => p.label)];
    const aoa = [header, ...rows.map(row => [
      row.label,
      ...periods.map((p, i) => {
        if (row.type === "growth") {
          const g = periodGrowth(periods, i, row.key, quarterly);
          return g === null ? "N/A" : fmtPctSigned(g * 100);
        }
        const v = p[row.key];
        return row.type === "percent" ? fmtPct(v) : fmtCr(v);
      }),
    ])];
    exportAoaToExcel(
      buildExportFilename(ds, `${title.replace(/[^a-z0-9]+/gi, "_")}_${quarterly ? "Quarterly" : "Yearly"}`),
      title.slice(0, 31),
      aoa
    );
  };

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">{title}</div>
        <div className="fin-section__controls">
          <PeriodToggle mode={mode} onChange={setMode} />
          <ExportButton onClick={handleExport} />
        </div>
      </div>
      <div className="fin-table-wrap">
        <table className="fin-table">
          <thead>
            <tr>
              <th className="fin-table__label-col">Metric</th>
              {periods.map(p => <th key={p.key}>{p.label.replace(" (partial)", "")}{!quarterly && p.partial ? <span className="fin-table__partial"> (partial)</span> : ""}{quarterly && !p.complete ? <span className="fin-table__partial"> (partial)</span> : ""}</th>)}
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.label}>
                <td className="fin-table__label-col">{row.label}</td>
                {periods.map((p, i) => {
                  if (row.type === "growth") {
                    return <td key={p.key}><GrowthBadge value={periodGrowth(periods, i, row.key, quarterly)} /></td>;
                  }
                  const v = p[row.key];
                  return <td key={p.key}>{row.type === "percent" ? fmtPct(v) : fmtCr(v)}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function ProfitAndLossTable({ ds }) {
  const [mode, setMode] = useState("quarterly");
  if (!ds.hasRevenue) return null;
  const periods = periodsFor(ds, mode);
  const quarterly = mode === "quarterly";

  const revenueLines = ds.pnlRevenueSubLines
    ? ds.kpiKeys.filter(k => ds.pnlRevenueSubLines.some(re => re.test(k)))
    : ds.kpiKeys.filter(k => /revenue/i.test(k) && k !== ds.revenueBaseKey);
  const hasDirect = ds.kpiKeys.includes("Direct Expenses");
  // Most companies' operating-expense line is literally named "Indirect
  // Expenses"; a company whose MIS uses different terminology (e.g. MULTIPL's
  // "Total Operating Costs", which stands alone with no Gross Profit split)
  // can override via companyConfig.opexKey rather than forcing an
  // "Indirect Expenses" row that doesn't exist in its sheet.
  const opexKey = ds.opexKey || "Indirect Expenses";
  const hasIndirect = ds.kpiKeys.includes(opexKey);

  const sections = [
    {
      heading: ds.revenueLabel || "Revenue",
      rows: [
        ...revenueLines.map(k => ({ label: k, key: k, type: "currency" })),
        { label: ds.revenueLabel || "Total Revenue", key: ds.revenueBaseKey, type: "currency", subtotal: true },
      ],
    },
    ds.hasGP && {
      heading: "Cost of Revenue",
      rows: [
        hasDirect && { label: "Direct Expenses", key: "Direct Expenses", type: "currency" },
        { label: "Gross Profit", key: "Gross Profit", type: "currency", subtotal: true },
        { label: "Gross Margin %", key: "Gross Margin", type: "percent", subtotal: true },
      ].filter(Boolean),
    },
    ds.hasEBITDA && {
      heading: "Operating Expenses",
      rows: [
        hasIndirect && { label: opexKey, key: opexKey, type: "currency" },
        { label: "EBITDA", key: "EBITDA", type: "currency", subtotal: true },
        { label: "EBITDA Margin %", key: "EBITDA Margin", type: "percent", subtotal: true },
      ].filter(Boolean),
    },
    ds.hasNet && {
      heading: "Below EBITDA",
      rows: [
        { label: "Net Profit", key: "Net Profit", type: "currency", subtotal: true },
        { label: "Net Profit Margin %", key: "Net Margin", type: "percent", subtotal: true },
      ],
    },
  ].filter(Boolean);

  const handleExport = () => {
    const header = ["Line item", ...periods.map(p => p.label)];
    const aoa = [header];
    sections.forEach(sec => {
      aoa.push([sec.heading, ...periods.map(() => "")]);
      sec.rows.forEach(row => {
        aoa.push([row.label, ...periods.map(p => {
          const v = p[row.key];
          return row.type === "percent" ? fmtPct(v) : fmtCr(v);
        })]);
      });
    });
    exportAoaToExcel(
      buildExportFilename(ds, `Complete_PnL_${quarterly ? "Quarterly" : "Yearly"}`),
      "Profit & Loss",
      aoa
    );
  };

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">Profit &amp; Loss Statement</div>
        <div className="fin-section__controls">
          <PeriodToggle mode={mode} onChange={setMode} />
          <ExportButton onClick={handleExport} />
        </div>
      </div>
      <div className="fin-table-wrap">
        <table className="fin-table fin-table--pnl">
          <thead>
            <tr>
              <th className="fin-table__label-col">Line item</th>
              {periods.map(p => <th key={p.key}>{p.label.replace(" (partial)", "")}{quarterly ? (!p.complete ? <span className="fin-table__partial"> (partial)</span> : "") : (p.partial ? <span className="fin-table__partial"> (partial)</span> : "")}</th>)}
            </tr>
          </thead>
          {sections.map(sec => (
            <tbody key={sec.heading}>
              <tr className="fin-table__section-row">
                <td colSpan={periods.length + 1}>{sec.heading}</td>
              </tr>
              {sec.rows.map(row => (
                <tr key={row.label} className={row.subtotal ? "fin-table__subtotal-row" : ""}>
                  <td className="fin-table__label-col">{row.label}</td>
                  {periods.map(p => {
                    const v = p[row.key];
                    return <td key={p.key}>{row.type === "percent" ? fmtPct(v) : fmtCr(v)}</td>;
                  })}
                </tr>
              ))}
            </tbody>
          ))}
        </table>
      </div>
    </section>
  );
}

function PlaceholderSection({ title, note, eyebrow = "Coming next" }) {
  return (
    <section className="section">
      <div className="placeholder-page">
        <div className="placeholder-page__eyebrow">{eyebrow}</div>
        <div className="placeholder-page__title">{title}</div>
        <div className="placeholder-page__note">{note}</div>
      </div>
    </section>
  );
}

/* ============================================================
   PHASE 2 — Industry & Competitors, News & Updates. Both read
   from optional sheets ("Industry Data", "News Feed") parsed
   above, populated by an independent research pipeline (not the
   Excel financial pipeline — see parseNewsSheet/parseIndustrySheet).
   Every fact rendered here carries its own embedded source link;
   nothing is generated from model knowledge at render time.
   ============================================================ */
function fmtDate(d) {
  if (!d) return "N/A";
  const date = d instanceof Date ? d : new Date(d);
  if (isNaN(date)) return typeof d === "string" ? d : "N/A";
  return date.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function SourceLink({ name, url }) {
  if (!url || !name) return null;
  return (
    <a className="src-link" href={url} target="_blank" rel="noopener noreferrer">
      {name}<span className="src-link__arrow">↗</span>
    </a>
  );
}

function CapabilityMark({ value }) {
  const v = (value || "").trim();
  if (v === "✓" || /^y(es)?$/i.test(v)) return <span className="cap-mark cap-mark--yes">✓</span>;
  if (v === "—" || v === "-" || /^no?t?$/i.test(v)) return <span className="cap-mark cap-mark--no">—</span>;
  return <span className="cap-mark cap-mark--na">N/A</span>;
}

const NEWS_CATEGORY_ORDER = ["Company", "Product & Launches", "Partnerships", "Funding & Financial", "Leadership", "Events", "Industry"];
const NEWS_DATE_RANGES = [
  { key: "7d", label: "Last 7 Days", days: 7 },
  { key: "30d", label: "Last 30 Days", days: 30 },
  { key: "90d", label: "Last 90 Days", days: 90 },
  { key: "1y", label: "Last 1 Year", days: 365 },
  { key: "all", label: "All", days: Infinity },
];

/* ------------------------------------------------------------
   LIVE FETCH — used only when the uploaded workbook has no
   "News Feed" / "Industry Data" sheets. Calls a public news API
   directly from the browser (no server), in real time, on each
   page view. Nothing here is persisted: it lives in React state
   only and disappears on reload or re-upload, by design. Every
   item comes straight from the API's own url/source fields —
   nothing is invented or filled in.
   ------------------------------------------------------------ */
const NEWS_CATEGORY_KEYWORDS = [
  ["Funding & Financial", /\b(funding|raise[sd]?|series [a-e]|investment|valuation|ipo|acquir|acquisition|merger)\b/i],
  ["Partnerships", /\b(partner(ship|s)?|collaborat|tie-?up|alliance)\b/i],
  ["Product & Launches", /\b(launch(es|ed)?|unveil|introduc|rollout|release[sd]?|new (product|feature|platform))\b/i],
  ["Leadership", /\b(ceo|cfo|cto|appoint|hire[sd]?|joins as|steps down|resign|leadership)\b/i],
  ["Events", /\b(summit|conference|webinar|event|expo|forum)\b/i],
  ["Industry", /\b(industry|market( size| report| trend)?|sector|cagr)\b/i],
];
function guessNewsCategory(text) {
  const t = text || "";
  for (const [cat, re] of NEWS_CATEGORY_KEYWORDS) if (re.test(t)) return cat;
  return "Company";
}

function useLiveNews(query) {
  const [state, setState] = useState({ status: "idle", items: [] });
  useEffect(() => {
    if (!query) { setState({ status: "no-query", items: [] }); return; }
    const apiKey = import.meta.env.VITE_GNEWS_API_KEY;
    if (!apiKey) { setState({ status: "no-key", items: [] }); return; }

    let cancelled = false;
    setState({ status: "loading", items: [] });
    const url = `https://gnews.io/api/v4/search?q=${encodeURIComponent(query)}&lang=en&max=10&sortby=publishedAt&apikey=${apiKey}`;

    fetch(url)
      .then(res => {
        if (!res.ok) throw new Error(`API returned ${res.status}`);
        return res.json();
      })
      .then(data => {
        if (cancelled) return;
        const items = (data.articles || [])
          .filter(a => a && a.url && a.title)
          .map(a => ({
            title: a.title,
            summary: a.description || "",
            category: guessNewsCategory(`${a.title} ${a.description || ""}`),
            publishedAt: a.publishedAt ? new Date(a.publishedAt) : null,
            sourceName: a.source?.name || "Source",
            sourceUrl: a.url,
            secondarySourceName: null,
            secondarySourceUrl: null,
          }));
        setState({ status: "success", items });
      })
      .catch(err => {
        if (cancelled) return;
        setState({ status: "error", items: [], message: err.message });
      });

    return () => { cancelled = true; };
  }, [query]);
  return state;
}

function LiveFetchBanner() {
  return (
    <div className="news-refresh">
      <span className="news-refresh__stamp">● Live — fetched just now from the web</span>
      <span className="news-refresh__note">Not saved anywhere; this list disappears on reload or re-upload and is fetched fresh each time.</span>
    </div>
  );
}

function LiveNewsFeed({ query, emptyHint }) {
  const { status, items, message } = useLiveNews(query);

  if (!query) {
    return <div className="chart-empty">{emptyHint}</div>;
  }
  if (status === "no-key") {
    return <div className="chart-empty">Live news fetching isn't configured yet — no API key is set for this dashboard.</div>;
  }
  if (status === "loading" || status === "idle") {
    return <div className="chart-empty">Fetching live news for &quot;{query}&quot;…</div>;
  }
  if (status === "error") {
    return (
      <div className="chart-empty">
        Live news is unavailable right now{message ? ` (${message})` : ""}. This can happen if the request was blocked by your browser or network — try reloading.
      </div>
    );
  }
  if (!items.length) {
    return <div className="chart-empty">No live news found for &quot;{query}&quot; right now.</div>;
  }
  return (
    <>
      <LiveFetchBanner />
      <div className="news-grid">
        {items.map((item, i) => <NewsCard key={`${item.sourceUrl}-${i}`} item={item} />)}
      </div>
    </>
  );
}

function NewsCard({ item }) {
  return (
    <div className="news-card">
      <div className="news-card__top">
        <span className="news-card__category">{item.category}</span>
        <span className="news-card__date">{fmtDate(item.publishedAt)}</span>
      </div>
      <div className="news-card__title">{item.title}</div>
      {item.summary && <div className="news-card__summary">{item.summary}</div>}
      <div className="news-card__sources">
        <SourceLink name={item.sourceName} url={item.sourceUrl} />
        {item.secondarySourceUrl && <SourceLink name={item.secondarySourceName || "Additional source"} url={item.secondarySourceUrl} />}
      </div>
    </div>
  );
}

function NewsUpdatesPage({ ds }) {
  const [category, setCategory] = useState("All");
  const [range, setRange] = useState("90d");
  const feed = ds.newsFeed;

  if (feed === null) {
    const companyName = ds.companyInfo?.companyName;
    const query = companyName && !/your company/i.test(companyName) ? companyName : null;
    return (
      <section className="section">
        <div className="fin-section__head">
          <div className="section__title">News &amp; Updates</div>
        </div>
        <LiveNewsFeed
          query={query}
          emptyHint='Add a real "Company Name" in the Company Info sheet to fetch live news for your company.'
        />
      </section>
    );
  }

  const now = new Date();
  const rangeDef = NEWS_DATE_RANGES.find(r => r.key === range);
  const inRange = (list, days) => list.filter(f => {
    if (category !== "All" && f.category !== category) return false;
    if (!f.publishedAt) return true;
    return (now - f.publishedAt) / 86400000 <= days;
  });

  let filtered = inRange(feed, rangeDef.days);
  let expandedNote = null;
  if (filtered.length < 3 && rangeDef.key !== "all") {
    const wider = NEWS_DATE_RANGES.slice(NEWS_DATE_RANGES.findIndex(r => r.key === range) + 1);
    for (const w of wider) {
      const candidate = inRange(feed, w.days);
      if (candidate.length > filtered.length) {
        filtered = candidate;
        expandedNote = `Showing "${w.label}" instead — not enough items in the selected window.`;
      }
      if (candidate.length >= 3) break;
    }
  }

  const categoriesPresent = NEWS_CATEGORY_ORDER.filter(c => feed.some(f => f.category === c));
  const refreshedAt = ds.refreshMeta?.newsRefreshedAt;

  return (
    <section className="section">
      <div className="fin-section__head">
        <div className="section__title">News &amp; Updates</div>
        <div className="news-refresh">
          {refreshedAt && <span className="news-refresh__stamp">● Last refreshed: {refreshedAt}</span>}
          <span className="news-refresh__note">Refreshes automatically once a day; re-upload to pick up the latest run.</span>
        </div>
      </div>

      <div className="news-filters">
        <div className="news-filters__group">
          <button className={`chip ${category === "All" ? "chip--active" : ""}`} onClick={() => setCategory("All")}>All</button>
          {categoriesPresent.map(c => (
            <button key={c} className={`chip ${category === c ? "chip--active" : ""}`} onClick={() => setCategory(c)}>{c}</button>
          ))}
        </div>
        <div className="news-filters__group">
          {NEWS_DATE_RANGES.map(r => (
            <button key={r.key} className={`chip chip--mono ${range === r.key ? "chip--active" : ""}`} onClick={() => setRange(r.key)}>{r.label}</button>
          ))}
        </div>
      </div>

      {expandedNote && <div className="news-expanded-note">{expandedNote}</div>}

      {filtered.length ? (
        <div className="news-grid">
          {filtered.map((item, i) => <NewsCard key={`${item.title}-${i}`} item={item} />)}
        </div>
      ) : (
        <div className="chart-empty">No relevant verified updates found for this filter.</div>
      )}
    </section>
  );
}

function IndustryCompetitorsPage({ ds }) {
  const data = ds.industryData;

  if (data === null) {
    const companyName = ds.companyInfo?.companyName;
    const cleanName = companyName && !/your company/i.test(companyName) ? companyName : null;
    const tag = ds.companyInfo?.tags?.[0];
    const query = cleanName ? `${cleanName} industry` : (tag ? `${tag} industry India` : null);
    return (
      <section className="section">
        <div className="fin-section__head">
          <div className="section__title">Industry &amp; Competitors</div>
        </div>
        <div className="placeholder-page__note" style={{ marginBottom: 16 }}>
          A structured competitor/market breakdown needs the researched "Industry Data" sheet, which this upload doesn't have. Showing live industry news headlines instead, fetched directly from the web — not a substitute for verified competitor analysis.
        </div>
        <LiveNewsFeed
          query={query}
          emptyHint='Add a "Company Name" or a "Tag 1" in the Company Info sheet to fetch live industry news.'
        />
      </section>
    );
  }

  const capabilityCols = ["CRM", "Loyalty", "CDP", "Marketing Automation", "Conversational Commerce"];
  const refreshedAt = ds.refreshMeta?.industryRefreshedAt;

  return (
    <>
      <section className="section">
        <div className="fin-section__head">
          <div className="section__title">Industry &amp; Competitors</div>
          {refreshedAt && <span className="news-refresh__stamp">● Data refreshed: {refreshedAt}</span>}
        </div>

        {data.overviewDescription && <p className="biz-desc-text" style={{ marginBottom: 16 }}>{data.overviewDescription}</p>}

        {data.categories.length > 0 && (
          <div className="biz-chip-row" style={{ marginBottom: 8 }}>
            {data.categories.map(c => (
              <span className="biz-chip" key={c.Name}>
                {c.Name}
                {c.SourceUrl && <> · <SourceLink name={c.SourceName || "source"} url={c.SourceUrl} /></>}
              </span>
            ))}
          </div>
        )}
      </section>

      {data.snapshot.length > 0 && (
        <section className="section">
          <div className="section__title">Industry Snapshot</div>
          <div className="snapshot-grid">
            {data.snapshot.map(m => (
              <div className="snapshot-tile" key={m.Metric}>
                <div className="snapshot-tile__label">{m.Metric}</div>
                <div className="snapshot-tile__value">{m.Value}</div>
                <div className="snapshot-tile__meta">
                  <span className="snapshot-tile__period">{m.Period}</span>
                  <SourceLink name={m.SourceName} url={m.SourceUrl} />
                </div>
                {m.Note && <div className="snapshot-tile__note">{m.Note}</div>}
              </div>
            ))}
          </div>
        </section>
      )}

      {data.trends.length > 0 && (
        <section className="section">
          <div className="section__title">Industry Trends</div>
          <div className="chart-grid">
            {data.trends.map(t => (
              <div className="chart-card trend-card" key={t.Title}>
                <div className="chart-card__title">{t.Title}</div>
                {t.Description && <p className="trend-card__desc">{t.Description}</p>}
                {t.WhyItMatters && (
                  <div className="trend-card__why">
                    <span className="trend-card__why-label">Why it matters</span>
                    {t.WhyItMatters}
                  </div>
                )}
                <div className="trend-card__foot">
                  <SourceLink name={t.SourceName} url={t.SourceUrl} />
                  {t.PublishedAt && <span className="trend-card__date">{t.PublishedAt}</span>}
                </div>
              </div>
            ))}
          </div>
        </section>
      )}

      {data.competitors.length > 0 && (
        <section className="section">
          <div className="section__title">Competitive Landscape</div>
          <div className="fin-table-wrap">
            <table className="fin-table fin-table--competitors">
              <thead>
                <tr>
                  <th className="fin-table__label-col">Company</th>
                  <th>Primary Focus</th>
                  {capabilityCols.map(c => <th key={c}>{c}</th>)}
                  <th>Geography</th>
                  <th>Source</th>
                </tr>
              </thead>
              <tbody>
                {data.competitors.map(c => (
                  <tr key={c.name}>
                    <td className="fin-table__label-col">{c.name}</td>
                    <td style={{ textAlign: "left" }}>{c.PrimaryFocus || "N/A"}</td>
                    <td><CapabilityMark value={c.CRM} /></td>
                    <td><CapabilityMark value={c.Loyalty} /></td>
                    <td><CapabilityMark value={c.CDP} /></td>
                    <td><CapabilityMark value={c["MarketingAutomation"]} /></td>
                    <td><CapabilityMark value={c["ConversationalCommerce"]} /></td>
                    <td style={{ textAlign: "left" }}>{c.Geography || "N/A"}</td>
                    <td><SourceLink name={c.SourceName || "source"} url={c.SourceUrl} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {data.competitors.some(c => c.RelevanceNote) && (
            <div className="competitor-notes">
              {data.competitors.filter(c => c.RelevanceNote).map(c => (
                <div className="competitor-notes__item" key={c.name}><strong>{c.name}:</strong> {c.RelevanceNote}</div>
              ))}
            </div>
          )}
        </section>
      )}

      {data.analysis.length > 0 && (
        <section className="section">
          <div className="section__title">Company vs. Competitors <span className="section__sub">— dashboard analysis</span></div>
          <div className="analysis-list">
            {data.analysis.map((a, i) => (
              <div className="analysis-item" key={i}>
                <span className="analysis-item__badge">Dashboard analysis</span>
                <span>{a.Text}</span>
                {a.SourceUrl && <SourceLink name={a.SourceName || "supporting source"} url={a.SourceUrl} />}
              </div>
            ))}
          </div>
        </section>
      )}

      {data.methodology && (
        <div className="footnote">
          <Info size={13} style={{ flexShrink: 0, marginTop: 2 }} />
          <span>{data.methodology}</span>
        </div>
      )}
    </>
  );
}

/* ============================================================
   NEW — Executive summary (narrative, computed live off the
   latest complete quarter / FY in the sheet) and business
   description (static company context).
   ============================================================ */
/* Picks whichever margin the sheet actually supports, in order of how
   commonly it's tracked, so the third narrative bullet always has something
   meaningful to say regardless of which KPI rows a given company fills in. */
function bestMarginKey(ds) {
  if (ds.hasRevenue && ds.hasEBITDA) return ["EBITDA Margin", "EBITDA"];
  if (ds.hasRevenue && ds.hasNet) return ["Net Margin", "Net Profit"];
  if (ds.hasRevenue && ds.hasGP) return ["Gross Margin", "Gross Profit"];
  return null;
}

function ExecutiveSummary({ ds }) {
  const { latestQ, prevYearQ, latestFY, prevFY } = getExecStats(ds);
  const companyInfo = ds.companyInfo;

  const revKey = ds.revenueBaseKey;
  const revYoY = latestQ && prevYearQ && ds.hasRevenue && prevYearQ[revKey]
    ? ((latestQ[revKey] - prevYearQ[revKey]) / Math.abs(prevYearQ[revKey])) * 100
    : null;

  const ebitdaDelta = latestFY && prevFY && ds.hasEBITDA && typeof latestFY["EBITDA"] === "number" && typeof prevFY["EBITDA"] === "number"
    ? latestFY["EBITDA"] - prevFY["EBITDA"]
    : null;
  const ebitdaDirection = ebitdaDelta === null ? null : ebitdaDelta > 0 ? "improved" : ebitdaDelta < 0 ? "declined" : "held steady";

  const marginPick = bestMarginKey(ds);
  const marginFYIdx = latestFY ? ds.fyData.findIndex(f => f.key === latestFY.key) : -1;
  const marginPrevFY = marginFYIdx > 0 ? ds.fyData[marginFYIdx - 1] : null;
  const marginCurr = marginPick && latestFY ? latestFY[marginPick[0]] : null;
  const marginPrev = marginPick && marginPrevFY ? marginPrevFY[marginPick[0]] : null;

  return (
    <section className="section narrative-section">
      <div className="section__title">Performance Summary</div>
      <div className="narrative-grid">
        <div className="narrative-card">
          <div className="narrative-card__eyebrow">Quarterly &amp; FY momentum</div>
          <div className="narrative-bullets">
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><TrendingUp size={15} /></span>
              <span>
                {latestQ && prevYearQ && ds.hasRevenue && prevYearQ[revKey] ? (
                  <><strong>{latestQ.label}</strong> closed with {(ds.revenueLabel || "net revenue").toLowerCase()} of <strong>{fmtCr(latestQ[revKey])}</strong>,{" "}
                  {revYoY >= 0 ? "up" : "down"} {Math.abs(revYoY).toFixed(1)}% versus {fmtCr(prevYearQ[revKey])} in{" "}
                  {prevYearQ.label}.</>
                ) : (
                  <>Revenue data isn't complete enough yet to compute a like-for-like quarterly YoY comparison — add more
                  months to the sheet to unlock this.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Rocket size={15} /></span>
              <span>
                {latestFY && prevFY && ds.hasEBITDA && ebitdaDirection ? (
                  <>FY EBITDA {ebitdaDirection} from <strong>{fmtCr(prevFY["EBITDA"])}</strong> in {prevFY.label} to{" "}
                  <strong>{fmtCr(latestFY["EBITDA"])}</strong> in {latestFY.label}
                  {latestQ && typeof latestQ["EBITDA"] === "number" ? <>, with {latestQ.label} at <strong>{fmtCr(latestQ["EBITDA"])}</strong></> : ""}.</>
                ) : (
                  <>Not enough complete fiscal years of EBITDA yet to describe a trend — add more months to unlock this.</>
                )}
                {companyInfo?.strategicNote && (
                  <> <strong>{companyInfo.strategicNote.value}</strong>{companyInfo.strategicNote.sub ? ` — ${companyInfo.strategicNote.sub}` : ""}.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Target size={15} /></span>
              <span>
                {marginPick && latestFY && typeof marginCurr === "number" ? (
                  <>{marginPick[0]} for {latestFY.label} stood at <strong>{fmtPct(marginCurr)}</strong>
                  {typeof marginPrev === "number" ? <>, versus {fmtPct(marginPrev)} the prior FY</> : ""}.</>
                ) : (
                  <>Add {ds.revenueLabel || "Total Revenue"} plus Gross Profit, EBITDA, or Net Profit to the sheet to unlock a margin-trend read-out here.</>
                )}
              </span>
            </div>
          </div>
        </div>

        <div className="narrative-stat-col">
          <div className="stat-tile">
            <div className="stat-tile__label">Latest quarter vs prior year</div>
            <div className="stat-tile__value">{latestQ ? fmtCr(latestQ[revKey]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestQ ? latestQ.label : "—"} {(ds.revenueLabel || "revenue").toLowerCase()} {revYoY !== null && <Delta curr={latestQ?.[revKey]} prev={prevYearQ?.[revKey]} good="up" />}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">FY EBITDA</div>
            <div className="stat-tile__value">{latestFY ? fmtCr(latestFY["EBITDA"]) : "N/A"}</div>
            <div className="stat-tile__sub">{latestFY ? `${latestFY.label}, vs ${prevFY ? fmtCr(prevFY["EBITDA"]) : "N/A"} prior FY` : "—"}</div>
          </div>
          <div className="stat-tile">
            {companyInfo?.strategicNote ? (
              <>
                <div className="stat-tile__label">{companyInfo.strategicNote.label}</div>
                <div className="stat-tile__value" style={{ fontSize: 16 }}>{companyInfo.strategicNote.value}</div>
                {companyInfo.strategicNote.sub && <div className="stat-tile__sub">{companyInfo.strategicNote.sub}</div>}
              </>
            ) : (
              <>
                <div className="stat-tile__label">{marginPick ? marginPick[0] : "FY Margin"}</div>
                <div className="stat-tile__value">{marginPick && typeof marginCurr === "number" ? fmtPct(marginCurr) : "N/A"}</div>
                <div className="stat-tile__sub">{latestFY ? `${latestFY.label}${typeof marginPrev === "number" ? `, vs ${fmtPct(marginPrev)} prior FY` : ""}` : "—"}</div>
              </>
            )}
          </div>
        </div>
      </div>
    </section>
  );
}

/* ============================================================
   GRAYQUEST PERFORMANCE COMMENTARY — plays the same role as
   ExecutiveSummary but with GrayQuest-specific insights (disbursal
   growth, Schools/Colleges/Edtech mix, take rate, cost of funds)
   layered in per the spec. Every sentence below either renders a
   number computed straight from ds.fyData/ds.qData, or renders a
   plain "not enough data yet" fallback — never a qualitative claim
   without a number behind it.
   ============================================================ */
function describeEbitdaTrend(curr, prev) {
  if (typeof curr !== "number" || typeof prev !== "number") return null;
  if (prev < 0 && curr < 0) return curr > prev ? "loss narrowed" : curr < prev ? "loss widened" : "loss held steady";
  if (prev < 0 && curr >= 0) return "turned positive";
  if (prev >= 0 && curr < 0) return "turned negative";
  return curr > prev ? "grew" : curr < prev ? "declined" : "held steady";
}

function GrayQuestCommentary({ ds }) {
  const { latestQ, prevYearQ, latestFY, prevFY } = getExecStats(ds);
  const revKey = ds.revenueBaseKey;
  const kpiCfg = ds.companyConfig.kpi;
  const totalDisbKey = kpiCfg ? ds.kpiKeys.find(k => kpiCfg.totalMatch.test(k)) : null;

  const revYoY = latestQ && prevYearQ && typeof latestQ[revKey] === "number" && typeof prevYearQ[revKey] === "number" && prevYearQ[revKey]
    ? ((latestQ[revKey] - prevYearQ[revKey]) / Math.abs(prevYearQ[revKey])) * 100 : null;

  const disbYoY = totalDisbKey && latestFY && prevFY && typeof latestFY[totalDisbKey] === "number" && typeof prevFY[totalDisbKey] === "number" && prevFY[totalDisbKey]
    ? ((latestFY[totalDisbKey] - prevFY[totalDisbKey]) / Math.abs(prevFY[totalDisbKey])) * 100 : null;

  const ebitdaCurr = latestFY && typeof latestFY["EBITDA"] === "number" ? latestFY["EBITDA"] : null;
  const ebitdaPrev = prevFY && typeof prevFY["EBITDA"] === "number" ? prevFY["EBITDA"] : null;
  const ebitdaTrend = describeEbitdaTrend(ebitdaCurr, ebitdaPrev);

  const marginPick = bestMarginKey(ds);
  const marginFYIdx = latestFY ? ds.fyData.findIndex(f => f.key === latestFY.key) : -1;
  const marginPrevFY = marginFYIdx > 0 ? ds.fyData[marginFYIdx - 1] : null;
  const marginCurr = marginPick && latestFY ? latestFY[marginPick[0]] : null;
  const marginPrev = marginPick && marginPrevFY ? marginPrevFY[marginPick[0]] : null;

  function shareFor(fy, matchedKey) {
    if (!fy || !matchedKey || !totalDisbKey) return null;
    const num = fy[matchedKey], den = fy[totalDisbKey];
    return (typeof num === "number" && typeof den === "number" && den !== 0) ? (num / den) * 100 : null;
  }
  const shareRows = (kpiCfg?.shareRows || []).map(r => ({ label: r.label, matchedKey: ds.kpiKeys.find(k => r.match.test(k)) || null }));
  const shares = shareRows
    .map(r => ({ label: r.label, curr: shareFor(latestFY, r.matchedKey), prev: shareFor(prevFY, r.matchedKey) }))
    .filter(s => s.curr !== null);

  const takeRateKey = ds.kpiKeys.find(k => /take\s*rate/i.test(k)) || null;
  const takeRateCurr = takeRateKey && latestFY && typeof latestFY[takeRateKey] === "number" ? latestFY[takeRateKey] : null;
  const takeRatePrev = takeRateKey && prevFY && typeof prevFY[takeRateKey] === "number" ? prevFY[takeRateKey] : null;

  const cofKey = ds.kpiKeys.find(k => /cost\s*of\s*funds?/i.test(k)) || null;
  const cofCurr = cofKey && latestFY && typeof latestFY[cofKey] === "number" ? latestFY[cofKey] : null;
  const cofPrev = cofKey && prevFY && typeof prevFY[cofKey] === "number" ? prevFY[cofKey] : null;

  return (
    <section className="section narrative-section">
      <div className="section__title">Performance Commentary</div>
      <div className="narrative-grid">
        <div className="narrative-card">
          <div className="narrative-card__eyebrow">Growth &amp; mix</div>
          <div className="narrative-bullets">
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><TrendingUp size={15} /></span>
              <span>
                {totalDisbKey && latestFY && prevFY && disbYoY !== null ? (
                  <>Disbursals {disbYoY >= 0 ? "increased" : "decreased"} <strong>{Math.abs(disbYoY).toFixed(1)}%</strong> YoY to{" "}
                  <strong>{fmtNum(latestFY[totalDisbKey])}</strong> loans in {latestFY.label}, from {fmtNum(prevFY[totalDisbKey])} in {prevFY.label}.</>
                ) : (
                  <>Not enough complete fiscal years of disbursal data yet to compute YoY disbursal growth.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Rocket size={15} /></span>
              <span>
                {latestQ && prevYearQ && revYoY !== null ? (
                  <><strong>{latestQ.label}</strong> closed with {(ds.revenueLabel || "net revenue").toLowerCase()} of{" "}
                  <strong>{fmtCr(latestQ[revKey])}</strong>, {revYoY >= 0 ? "up" : "down"} {Math.abs(revYoY).toFixed(1)}% versus{" "}
                  {fmtCr(prevYearQ[revKey])} in {prevYearQ.label}.</>
                ) : (
                  <>{ds.revenueLabel || "Revenue"} data isn't complete enough yet to compute a like-for-like quarterly YoY comparison.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Users size={15} /></span>
              <span>
                {shares.length ? (
                  <>
                    {shares.map((s, i) => (
                      <React.Fragment key={s.label}>
                        {i > 0 ? "; " : ""}
                        <strong>{s.label}</strong> accounted for <strong>{s.curr.toFixed(1)}%</strong> of total disbursals in {latestFY.label}
                        {typeof s.prev === "number" ? <>, versus {s.prev.toFixed(1)}% in {prevFY.label}</> : ""}
                      </React.Fragment>
                    ))}.
                  </>
                ) : (
                  <>Segment-level disbursal mix (Schools / Colleges / Edtech Platforms) isn't available yet for this period.</>
                )}
              </span>
            </div>
          </div>
        </div>

        <div className="narrative-stat-col">
          <div className="stat-tile">
            <div className="stat-tile__label">Latest quarter vs prior year</div>
            <div className="stat-tile__value">{latestQ ? fmtCr(latestQ[revKey]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestQ ? latestQ.label : "—"} {(ds.revenueLabel || "revenue").toLowerCase()} {revYoY !== null && <Delta curr={latestQ?.[revKey]} prev={prevYearQ?.[revKey]} good="up" />}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">FY EBITDA</div>
            <div className="stat-tile__value">{latestFY ? fmtCr(latestFY["EBITDA"]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestFY ? `${latestFY.label}${ebitdaTrend ? `, ${ebitdaTrend} from ${fmtCr(ebitdaPrev)} prior FY` : ""}` : "—"}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">{marginPick ? marginPick[0] : "FY Margin"}</div>
            <div className="stat-tile__value">{marginPick && typeof marginCurr === "number" ? fmtPct(marginCurr) : "N/A"}</div>
            <div className="stat-tile__sub">{latestFY ? `${latestFY.label}${typeof marginPrev === "number" ? `, vs ${fmtPct(marginPrev)} prior FY` : ""}` : "—"}</div>
          </div>
        </div>
      </div>

      {(takeRateKey || cofKey) && (typeof takeRateCurr === "number" || typeof cofCurr === "number") && (
        <div className="narrative-extra-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          <span>
            {typeof takeRateCurr === "number" && (
              <>Take rate averaged <strong>{fmtPct(takeRateCurr)}</strong> in {latestFY?.label}
              {typeof takeRatePrev === "number" ? <>, versus {fmtPct(takeRatePrev)} the prior FY</> : ""}. </>
            )}
            {typeof cofCurr === "number" && (
              <>Cost of funds averaged <strong>{fmtPct(cofCurr)}</strong> in {latestFY?.label}
              {typeof cofPrev === "number" ? <>, versus {fmtPct(cofPrev)} the prior FY</> : ""}. </>
            )}
            <span style={{ opacity: 0.75 }}>(period figures are a simple average of the monthly rate, not volume-weighted)</span>
          </span>
        </div>
      )}
    </section>
  );
}

/* ============================================================
   RISKCOVRY PERFORMANCE SUMMARY — plays the same "Performance
   Summary / Commentary" role as ExecutiveSummary, but layers in
   Riskcovry-specific insights: revenue growth, revenue-mix (Platform
   Subscription / Product-Commission / Setup Fees contribution),
   EBITDA + margin movement, operating-expense trend, and headcount
   productivity where available. Titled "Performance Summary" (not
   "Performance Commentary") to match the section name in the
   Riskcovry page hierarchy — it sits right after Revenue &
   Profitability and before the KPI table, unlike GrayQuest's
   commentary which sits after both. Every sentence renders a number
   computed from ds.fyData/ds.qData, or a plain "not enough data yet"
   fallback — never a qualitative claim without a number behind it.
   ============================================================ */
function RiskcovryCommentary({ ds }) {
  const { latestQ, prevYearQ, latestFY, prevFY } = getExecStats(ds);
  const revKey = ds.revenueBaseKey;

  const revYoY = latestQ && prevYearQ && typeof latestQ[revKey] === "number" && typeof prevYearQ[revKey] === "number" && prevYearQ[revKey]
    ? ((latestQ[revKey] - prevYearQ[revKey]) / Math.abs(prevYearQ[revKey])) * 100 : null;

  const ebitdaCurr = latestFY && typeof latestFY["EBITDA"] === "number" ? latestFY["EBITDA"] : null;
  const ebitdaPrev = prevFY && typeof prevFY["EBITDA"] === "number" ? prevFY["EBITDA"] : null;
  const ebitdaTrend = describeEbitdaTrend(ebitdaCurr, ebitdaPrev);

  const marginPick = (ds.hasRevenue && ds.hasGP) ? ["Gross Margin", "Gross Profit"] : bestMarginKey(ds);
  const marginFYIdx = latestFY ? ds.fyData.findIndex(f => f.key === latestFY.key) : -1;
  const marginPrevFY = marginFYIdx > 0 ? ds.fyData[marginFYIdx - 1] : null;
  const marginCurr = marginPick && latestFY ? latestFY[marginPick[0]] : null;
  const marginPrev = marginPick && marginPrevFY ? marginPrevFY[marginPick[0]] : null;

  // Platform Subscription + Product/Commission + One-time Setup Fees sum to
  // "Gross Revenue" in this workbook, not to the Net Revenue base used
  // elsewhere on the page (see the pnlRevenueSubLines comment in
  // COMPANY_CONFIGS) — so the mix share below is computed against Gross
  // Revenue specifically, and labelled as such, rather than silently reusing
  // ds.revenueBaseKey and producing shares that could read as >100%.
  const mixBaseKey = ds.kpiKeys.includes("Gross Revenue") ? "Gross Revenue" : null;
  function shareFor(fy, matchedKey) {
    if (!fy || !matchedKey || !mixBaseKey || typeof fy[mixBaseKey] !== "number" || fy[mixBaseKey] === 0) return null;
    const num = fy[matchedKey];
    return typeof num === "number" ? (num / fy[mixBaseKey]) * 100 : null;
  }
  const mixCfg = ds.companyConfig.kpi?.mixRows || [];
  const mixRows = mixCfg.map(r => ({ label: r.label, matchedKey: ds.kpiKeys.find(k => r.match.test(k)) || null }));
  const mixes = mixRows
    .map(r => ({ label: r.label, curr: shareFor(latestFY, r.matchedKey), prev: shareFor(prevFY, r.matchedKey) }))
    .filter(m => m.curr !== null);

  const opexKey = ds.kpiKeys.includes("Indirect Expenses") ? "Indirect Expenses" : null;
  const opexCurr = opexKey && latestFY && typeof latestFY[opexKey] === "number" ? latestFY[opexKey] : null;
  const opexPrev = opexKey && prevFY && typeof prevFY[opexKey] === "number" ? prevFY[opexKey] : null;
  const opexGrowth = (typeof opexCurr === "number" && typeof opexPrev === "number" && opexPrev !== 0)
    ? ((opexCurr - opexPrev) / Math.abs(opexPrev)) * 100 : null;

  const headcountCurr = latestFY && typeof latestFY["Headcount"] === "number" ? latestFY["Headcount"] : null;
  const revPerEmpCurr = latestFY && typeof latestFY["Rev per Employee"] === "number" ? latestFY["Rev per Employee"] : null;
  const revPerEmpPrev = prevFY && typeof prevFY["Rev per Employee"] === "number" ? prevFY["Rev per Employee"] : null;

  return (
    <section className="section narrative-section">
      <div className="section__title">Performance Summary</div>
      <div className="narrative-grid">
        <div className="narrative-card">
          <div className="narrative-card__eyebrow">Growth &amp; mix</div>
          <div className="narrative-bullets">
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><TrendingUp size={15} /></span>
              <span>
                {latestQ && prevYearQ && revYoY !== null ? (
                  <><strong>{latestQ.label}</strong> closed with {(ds.revenueLabel || "revenue").toLowerCase()} of{" "}
                  <strong>{fmtCr(latestQ[revKey])}</strong>, {revYoY >= 0 ? "up" : "down"} {Math.abs(revYoY).toFixed(1)}% versus{" "}
                  {fmtCr(prevYearQ[revKey])} in {prevYearQ.label}.</>
                ) : (
                  <>{ds.revenueLabel || "Revenue"} data isn't complete enough yet to compute a like-for-like quarterly YoY comparison.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Rocket size={15} /></span>
              <span>
                {latestFY && prevFY && ebitdaTrend ? (
                  <>FY EBITDA {ebitdaTrend} from <strong>{fmtCr(ebitdaPrev)}</strong> in {prevFY.label} to{" "}
                  <strong>{fmtCr(ebitdaCurr)}</strong> in {latestFY.label}
                  {marginPick && typeof marginCurr === "number" ? <>, with {marginPick[0].toLowerCase()} at <strong>{fmtPct(marginCurr)}</strong>
                  {typeof marginPrev === "number" ? <> versus {fmtPct(marginPrev)} the prior FY</> : ""}</> : ""}.</>
                ) : (
                  <>Not enough complete fiscal years of EBITDA yet to describe a trend.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Layers size={15} /></span>
              <span>
                {mixes.length ? (
                  <>
                    {mixes.map((m, i) => (
                      <React.Fragment key={m.label}>
                        {i > 0 ? "; " : ""}
                        <strong>{m.label}</strong> contributed <strong>{m.curr.toFixed(1)}%</strong> of gross revenue in {latestFY.label}
                        {typeof m.prev === "number" ? <>, versus {m.prev.toFixed(1)}% in {prevFY.label}</> : ""}
                      </React.Fragment>
                    ))}.
                  </>
                ) : (
                  <>Revenue-mix data (Platform Subscription / Product-Commission / Setup Fees, as a share of Gross Revenue) isn't available yet for this period.</>
                )}
              </span>
            </div>
            {opexGrowth !== null && (
              <div className="narrative-bullet">
                <span className="narrative-bullet__icon"><Target size={15} /></span>
                <span>
                  Total operating expenses {opexGrowth >= 0 ? "increased" : "decreased"} <strong>{Math.abs(opexGrowth).toFixed(1)}%</strong> YoY to{" "}
                  <strong>{fmtCr(opexCurr)}</strong> in {latestFY.label}, from {fmtCr(opexPrev)} in {prevFY.label}.
                </span>
              </div>
            )}
          </div>
        </div>

        <div className="narrative-stat-col">
          <div className="stat-tile">
            <div className="stat-tile__label">Latest quarter vs prior year</div>
            <div className="stat-tile__value">{latestQ ? fmtCr(latestQ[revKey]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestQ ? latestQ.label : "—"} {(ds.revenueLabel || "revenue").toLowerCase()} {revYoY !== null && <Delta curr={latestQ?.[revKey]} prev={prevYearQ?.[revKey]} good="up" />}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">FY EBITDA</div>
            <div className="stat-tile__value">{latestFY ? fmtCr(latestFY["EBITDA"]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestFY ? `${latestFY.label}${ebitdaTrend ? `, ${ebitdaTrend} from ${fmtCr(ebitdaPrev)} prior FY` : ""}` : "—"}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">{marginPick ? marginPick[0] : "FY Margin"}</div>
            <div className="stat-tile__value">{marginPick && typeof marginCurr === "number" ? fmtPct(marginCurr) : "N/A"}</div>
            <div className="stat-tile__sub">{latestFY ? `${latestFY.label}${typeof marginPrev === "number" ? `, vs ${fmtPct(marginPrev)} prior FY` : ""}` : "—"}</div>
          </div>
        </div>
      </div>

      {typeof revPerEmpCurr === "number" && (
        <div className="narrative-extra-note">
          <Info size={12} style={{ flexShrink: 0, position: "relative", top: 1 }} />
          <span>
            Revenue per employee was <strong>{fmtCr(revPerEmpCurr)}</strong> in {latestFY?.label} (avg headcount {fmtNum(headcountCurr)})
            {typeof revPerEmpPrev === "number" ? <>, versus {fmtCr(revPerEmpPrev)} the prior FY</> : ""}.
          </span>
        </div>
      )}
    </section>
  );
}

/* ============================================================
   MULTIPL PERFORMANCE COMMENTARY — revenue trend/volatility, EBITDA
   / burn, goal-creation growth + average goal value, and adoption
   metrics (AUA / signups / brand partners) where available. Unlike
   GrayQuest/Riskcovry's commentary, this deliberately does NOT
   touch "Total Downloads" — that concept has no MIS row at all in
   this workbook, and the fallback KPI data is reserved for the
   dedicated KPI table (per spec), not narrative claims: making a
   qualitative "downloads grew" statement here would go beyond what
   the uploaded MIS itself can support.
   ============================================================ */
function MultiplCommentary({ ds }) {
  const { latestQ, prevYearQ, latestFY, prevFY } = getExecStats(ds);
  const revKey = ds.revenueBaseKey;

  const revYoY = latestQ && prevYearQ && typeof latestQ[revKey] === "number" && typeof prevYearQ[revKey] === "number" && prevYearQ[revKey]
    ? ((latestQ[revKey] - prevYearQ[revKey]) / Math.abs(prevYearQ[revKey])) * 100 : null;

  const ebitdaCurr = latestFY && typeof latestFY["EBITDA"] === "number" ? latestFY["EBITDA"] : null;
  const ebitdaPrev = prevFY && typeof prevFY["EBITDA"] === "number" ? prevFY["EBITDA"] : null;
  const ebitdaTrend = describeEbitdaTrend(ebitdaCurr, ebitdaPrev);

  // No Gross Profit or Net Profit rows in this workbook, so bestMarginKey
  // resolves to EBITDA Margin — the only margin concept this MIS supports.
  const marginPick = bestMarginKey(ds);
  const marginFYIdx = latestFY ? ds.fyData.findIndex(f => f.key === latestFY.key) : -1;
  const marginPrevFY = marginFYIdx > 0 ? ds.fyData[marginFYIdx - 1] : null;
  const marginCurr = marginPick && latestFY ? latestFY[marginPick[0]] : null;
  const marginPrev = marginPick && marginPrevFY ? marginPrevFY[marginPick[0]] : null;

  // "Goals Created" is aggregated as a period delta (deltaRowMatchers) — this
  // is genuinely "goals created within the FY", not a cumulative total.
  const goalsKey = ds.kpiKeys.find(k => /goals\s*created/i.test(k)) || null;
  const goalsCurr = goalsKey && latestFY && typeof latestFY[goalsKey] === "number" ? latestFY[goalsKey] : null;
  const goalsPrev = goalsKey && prevFY && typeof prevFY[goalsKey] === "number" ? prevFY[goalsKey] : null;
  const goalsGrowth = (typeof goalsCurr === "number" && typeof goalsPrev === "number" && goalsPrev !== 0)
    ? ((goalsCurr - goalsPrev) / Math.abs(goalsPrev)) * 100 : null;

  const avgGoalKey = ds.kpiKeys.find(k => /average\s*goal\s*value/i.test(k)) || null;
  const avgGoalCurr = avgGoalKey && latestFY && typeof latestFY[avgGoalKey] === "number" ? latestFY[avgGoalKey] : null;
  const avgGoalPrev = avgGoalKey && prevFY && typeof prevFY[avgGoalKey] === "number" ? prevFY[avgGoalKey] : null;

  const auaKey = ds.kpiKeys.find(k => /assets\s*under\s*advice/i.test(k)) || null;
  const auaCurr = auaKey && latestFY && typeof latestFY[auaKey] === "number" ? latestFY[auaKey] : null;
  const auaPrev = auaKey && prevFY && typeof prevFY[auaKey] === "number" ? prevFY[auaKey] : null;
  const auaGrowth = (typeof auaCurr === "number" && typeof auaPrev === "number" && auaPrev !== 0)
    ? ((auaCurr - auaPrev) / Math.abs(auaPrev)) * 100 : null;

  const signupsKey = ds.kpiKeys.find(k => /^signups$/i.test(k)) || null;
  const signupsCurr = signupsKey && latestFY && typeof latestFY[signupsKey] === "number" ? latestFY[signupsKey] : null;

  const brandKey = ds.kpiKeys.find(k => /brand\s*partners/i.test(k)) || null;
  const brandCurr = brandKey && latestFY && typeof latestFY[brandKey] === "number" ? latestFY[brandKey] : null;
  const brandPrev = brandKey && prevFY && typeof prevFY[brandKey] === "number" ? prevFY[brandKey] : null;

  return (
    <section className="section narrative-section">
      <div className="section__title">Performance Commentary</div>
      <div className="narrative-grid">
        <div className="narrative-card">
          <div className="narrative-card__eyebrow">Growth &amp; adoption</div>
          <div className="narrative-bullets">
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><TrendingUp size={15} /></span>
              <span>
                {latestQ && prevYearQ && revYoY !== null ? (
                  <><strong>{latestQ.label}</strong> closed with {(ds.revenueLabel || "revenue").toLowerCase()} of{" "}
                  <strong>{fmtCr(latestQ[revKey])}</strong>, {revYoY >= 0 ? "up" : "down"} {Math.abs(revYoY).toFixed(1)}% versus{" "}
                  {fmtCr(prevYearQ[revKey])} in {prevYearQ.label}{Math.abs(revYoY) > 40 ? " — quarterly revenue has shown meaningful volatility" : ""}.</>
                ) : (
                  <>{ds.revenueLabel || "Revenue"} data isn't complete enough yet to compute a like-for-like quarterly YoY comparison.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Rocket size={15} /></span>
              <span>
                {latestFY && prevFY && ebitdaTrend ? (
                  <>FY EBITDA {ebitdaTrend} from <strong>{fmtCr(ebitdaPrev)}</strong> in {prevFY.label} to{" "}
                  <strong>{fmtCr(ebitdaCurr)}</strong> in {latestFY.label}
                  {marginPick && typeof marginCurr === "number" ? <>, with {marginPick[0].toLowerCase()} at <strong>{fmtPct(marginCurr)}</strong>
                  {typeof marginPrev === "number" ? <> versus {fmtPct(marginPrev)} the prior FY</> : ""}</> : ""}
                  {" "}— consistent with continued investment ahead of monetisation.</>
                ) : (
                  <>Not enough complete fiscal years of EBITDA yet to describe a trend.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Layers size={15} /></span>
              <span>
                {goalsKey && latestFY && typeof goalsCurr === "number" ? (
                  <>Goals created {goalsGrowth !== null ? <>{goalsGrowth >= 0 ? "grew" : "declined"} <strong>{Math.abs(goalsGrowth).toFixed(1)}%</strong> YoY to</> : "reached"}{" "}
                  <strong>{fmtNum(goalsCurr)}</strong> in {latestFY.label}
                  {typeof goalsPrev === "number" ? <>, from {fmtNum(goalsPrev)} in {prevFY.label}</> : ""}
                  {typeof avgGoalCurr === "number" ? <>, while average goal value stood at <strong>{fmtRupee(avgGoalCurr)}</strong>
                  {typeof avgGoalPrev === "number" ? <> (versus {fmtRupee(avgGoalPrev)} the prior FY)</> : ""}</> : ""}.</>
                ) : (
                  <>Goal-creation data isn't complete enough yet to compute a YoY comparison.</>
                )}
              </span>
            </div>
            {(typeof auaCurr === "number" || typeof signupsCurr === "number" || typeof brandCurr === "number") && (
              <div className="narrative-bullet">
                <span className="narrative-bullet__icon"><Users size={15} /></span>
                <span>
                  {typeof auaCurr === "number" && (
                    <>Assets Under Advice reached <strong>{fmtCrAlready(auaCurr)}</strong> as of {latestFY?.label}
                    {auaGrowth !== null ? <>, up {auaGrowth.toFixed(1)}% versus {fmtCrAlready(auaPrev)} the prior FY-end</> : ""}. </>
                  )}
                  {typeof signupsCurr === "number" && <>Cumulative signups stood at <strong>{fmtNum(signupsCurr)}</strong> as of {latestFY?.label}. </>}
                  {typeof brandCurr === "number" && (
                    <>Active brand partners totalled <strong>{fmtNum(brandCurr)}</strong> as of {latestFY?.label}
                    {typeof brandPrev === "number" ? <>, versus {fmtNum(brandPrev)} the prior FY-end</> : ""}.</>
                  )}
                </span>
              </div>
            )}
          </div>
        </div>

        <div className="narrative-stat-col">
          <div className="stat-tile">
            <div className="stat-tile__label">Latest quarter vs prior year</div>
            <div className="stat-tile__value">{latestQ ? fmtCr(latestQ[revKey]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestQ ? latestQ.label : "—"} {(ds.revenueLabel || "revenue").toLowerCase()} {revYoY !== null && <Delta curr={latestQ?.[revKey]} prev={prevYearQ?.[revKey]} good="up" />}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">FY EBITDA</div>
            <div className="stat-tile__value">{latestFY ? fmtCr(latestFY["EBITDA"]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestFY ? `${latestFY.label}${ebitdaTrend ? `, ${ebitdaTrend} from ${fmtCr(ebitdaPrev)} prior FY` : ""}` : "—"}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">Goals Created</div>
            <div className="stat-tile__value">{typeof goalsCurr === "number" ? fmtNum(goalsCurr) : "N/A"}</div>
            <div className="stat-tile__sub">{latestFY ? `${latestFY.label}${goalsGrowth !== null ? `, ${goalsGrowth >= 0 ? "+" : ""}${goalsGrowth.toFixed(1)}% YoY` : ""}` : "—"}</div>
          </div>
        </div>
      </div>
    </section>
  );
}

/* ============================================================
   FASTSURANCE (INSURANCE SAMADHAN) PERFORMANCE COMMENTARY —
   revenue and EBITDA trend (same methodology as every other
   company's commentary), plus an operational-KPI bullet that
   reports registrations/resolved-cases/% resolved when the sheet
   has them and says plainly that it doesn't when it currently
   doesn't (this MIS is financial-P&amp;L-only today) — never a
   qualitative claim standing in for a missing number.
   ============================================================ */
function FastsuranceCommentary({ ds }) {
  const { latestQ, prevYearQ, latestFY, prevFY } = getExecStats(ds);
  const revKey = ds.revenueBaseKey;
  const kpiCfg = ds.companyConfig.kpi;

  const revYoY = latestQ && prevYearQ && typeof latestQ[revKey] === "number" && typeof prevYearQ[revKey] === "number" && prevYearQ[revKey]
    ? ((latestQ[revKey] - prevYearQ[revKey]) / Math.abs(prevYearQ[revKey])) * 100 : null;

  const ebitdaCurr = latestFY && typeof latestFY["EBITDA"] === "number" ? latestFY["EBITDA"] : null;
  const ebitdaPrev = prevFY && typeof prevFY["EBITDA"] === "number" ? prevFY["EBITDA"] : null;
  const ebitdaTrend = describeEbitdaTrend(ebitdaCurr, ebitdaPrev);

  const marginPick = bestMarginKey(ds);
  const marginFYIdx = latestFY ? ds.fyData.findIndex(f => f.key === latestFY.key) : -1;
  const marginPrevFY = marginFYIdx > 0 ? ds.fyData[marginFYIdx - 1] : null;
  const marginCurr = marginPick && latestFY ? latestFY[marginPick[0]] : null;
  const marginPrev = marginPick && marginPrevFY ? marginPrevFY[marginPick[0]] : null;

  const opRows = (kpiCfg?.rows || []).map(r => ({ ...r, matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null }));
  const opAvailable = opRows.filter(r => r.matchedKey);
  const pctResolvedRow = opRows.find(r => r.slug === "pctResolved");
  const pctResolvedCurr = pctResolvedRow?.matchedKey && latestFY && typeof latestFY[pctResolvedRow.matchedKey] === "number" ? latestFY[pctResolvedRow.matchedKey] : null;

  return (
    <section className="section narrative-section">
      <div className="section__title">Performance Commentary</div>
      <div className="narrative-grid">
        <div className="narrative-card">
          <div className="narrative-card__eyebrow">Revenue &amp; profitability</div>
          <div className="narrative-bullets">
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><TrendingUp size={15} /></span>
              <span>
                {latestQ && prevYearQ && revYoY !== null ? (
                  <><strong>{latestQ.label}</strong> closed with {(ds.revenueLabel || "net revenue").toLowerCase()} of{" "}
                  <strong>{fmtCr(latestQ[revKey])}</strong>, {revYoY >= 0 ? "up" : "down"} {Math.abs(revYoY).toFixed(1)}% versus{" "}
                  {fmtCr(prevYearQ[revKey])} in {prevYearQ.label}.</>
                ) : (
                  <>{ds.revenueLabel || "Revenue"} data isn't complete enough yet to compute a like-for-like quarterly YoY comparison.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Rocket size={15} /></span>
              <span>
                {latestFY && prevFY && ebitdaTrend ? (
                  <>FY EBITDA {ebitdaTrend} from <strong>{fmtCr(ebitdaPrev)}</strong> in {prevFY.label} to{" "}
                  <strong>{fmtCr(ebitdaCurr)}</strong> in {latestFY.label}
                  {marginPick && typeof marginCurr === "number" ? <>, with {marginPick[0].toLowerCase()} at <strong>{fmtPct(marginCurr)}</strong>
                  {typeof marginPrev === "number" ? <> versus {fmtPct(marginPrev)} the prior FY</> : ""}</> : ""}.</>
                ) : (
                  <>Not enough complete fiscal years of EBITDA yet to describe a trend.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><ShieldCheck size={15} /></span>
              <span>
                {opAvailable.length ? (
                  <>{opAvailable.map((r, i) => (
                    <React.Fragment key={r.label}>
                      {i > 0 ? "; " : ""}<strong>{r.label}</strong> stood at <strong>{r.fmt(latestFY?.[r.matchedKey])}</strong> in {latestFY?.label}
                    </React.Fragment>
                  ))}.</>
                ) : (
                  <>This standardized MIS currently carries financial (P&amp;L) data only — case-level operating metrics
                  (registrations, resolved cases, resolution rate) aren't present in the uploaded sheet yet, so they're
                  shown as N/A above rather than estimated.</>
                )}
              </span>
            </div>
          </div>
        </div>

        <div className="narrative-stat-col">
          <div className="stat-tile">
            <div className="stat-tile__label">Latest quarter vs prior year</div>
            <div className="stat-tile__value">{latestQ ? fmtCr(latestQ[revKey]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestQ ? latestQ.label : "—"} {(ds.revenueLabel || "revenue").toLowerCase()} {revYoY !== null && <Delta curr={latestQ?.[revKey]} prev={prevYearQ?.[revKey]} good="up" />}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">FY EBITDA</div>
            <div className="stat-tile__value">{latestFY ? fmtCr(latestFY["EBITDA"]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestFY ? `${latestFY.label}${ebitdaTrend ? `, ${ebitdaTrend} from ${fmtCr(ebitdaPrev)} prior FY` : ""}` : "—"}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">% Resolved</div>
            <div className="stat-tile__value">{typeof pctResolvedCurr === "number" ? fmtPct(pctResolvedCurr) : "N/A"}</div>
            <div className="stat-tile__sub">{latestFY ? latestFY.label : "—"}</div>
          </div>
        </div>
      </div>
    </section>
  );
}

/* ============================================================
   VITRA (APEX FUTURE LABS) PERFORMANCE COMMENTARY — same
   revenue/EBITDA-margin methodology as every other company, plus a
   Customers/ARPU bullet. See the note on COMPANY_CONFIGS.
   apexFutureLabs: the currently-provided standardized workbook has
   no usable row labels, so every figure below renders as "not
   enough data yet" until a corrected file is uploaded — this
   component is otherwise fully generic and will narrate real
   numbers the moment the sheet's labels are fixed.
   ============================================================ */
function VitraCommentary({ ds }) {
  const { latestQ, prevYearQ, latestFY, prevFY } = getExecStats(ds);
  const revKey = ds.revenueBaseKey;
  const kpiCfg = ds.companyConfig.kpi;

  const revYoY = latestQ && prevYearQ && typeof latestQ[revKey] === "number" && typeof prevYearQ[revKey] === "number" && prevYearQ[revKey]
    ? ((latestQ[revKey] - prevYearQ[revKey]) / Math.abs(prevYearQ[revKey])) * 100 : null;

  const ebitdaCurr = latestFY && typeof latestFY["EBITDA"] === "number" ? latestFY["EBITDA"] : null;
  const ebitdaPrev = prevFY && typeof prevFY["EBITDA"] === "number" ? prevFY["EBITDA"] : null;
  const ebitdaTrend = describeEbitdaTrend(ebitdaCurr, ebitdaPrev);

  const marginPick = bestMarginKey(ds);
  const marginCurr = marginPick && latestFY ? latestFY[marginPick[0]] : null;
  const marginPrevFY = latestFY ? ds.fyData[ds.fyData.findIndex(f => f.key === latestFY.key) - 1] : null;
  const marginPrev = marginPick && marginPrevFY ? marginPrevFY[marginPick[0]] : null;

  const custKey = ds.kpiKeys.find(k => (kpiCfg?.rows || []).find(r => r.slug === "customers")?.matchers.some(re => re.test(k))) || null;
  const arpuKey = ds.kpiKeys.find(k => (kpiCfg?.rows || []).find(r => r.slug === "arpu")?.matchers.some(re => re.test(k))) || null;
  const custCurr = custKey && latestFY && typeof latestFY[custKey] === "number" ? latestFY[custKey] : null;
  const custPrev = custKey && prevFY && typeof prevFY[custKey] === "number" ? prevFY[custKey] : null;
  const arpuCurr = arpuKey && latestFY && typeof latestFY[arpuKey] === "number" ? latestFY[arpuKey] : null;

  return (
    <section className="section narrative-section">
      <div className="section__title">Performance Commentary</div>
      <div className="narrative-grid">
        <div className="narrative-card">
          <div className="narrative-card__eyebrow">Revenue &amp; profitability</div>
          <div className="narrative-bullets">
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><TrendingUp size={15} /></span>
              <span>
                {latestQ && prevYearQ && revYoY !== null ? (
                  <><strong>{latestQ.label}</strong> closed with {(ds.revenueLabel || "net revenue").toLowerCase()} of{" "}
                  <strong>{fmtCr(latestQ[revKey])}</strong>, {revYoY >= 0 ? "up" : "down"} {Math.abs(revYoY).toFixed(1)}% versus{" "}
                  {fmtCr(prevYearQ[revKey])} in {prevYearQ.label}.</>
                ) : (
                  <>{ds.revenueLabel || "Revenue"} data isn't complete enough yet to compute a like-for-like quarterly YoY comparison.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Rocket size={15} /></span>
              <span>
                {latestFY && prevFY && ebitdaTrend ? (
                  <>FY EBITDA {ebitdaTrend} from <strong>{fmtCr(ebitdaPrev)}</strong> in {prevFY.label} to{" "}
                  <strong>{fmtCr(ebitdaCurr)}</strong> in {latestFY.label}
                  {marginPick && typeof marginCurr === "number" ? <>, with {marginPick[0].toLowerCase()} at <strong>{fmtPct(marginCurr)}</strong>
                  {typeof marginPrev === "number" ? <> versus {fmtPct(marginPrev)} the prior FY</> : ""}</> : ""}.</>
                ) : (
                  <>Not enough complete fiscal years of EBITDA yet to describe a trend.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Users size={15} /></span>
              <span>
                {typeof custCurr === "number" ? (
                  <>Customers stood at <strong>{fmtNum(custCurr)}</strong> as of {latestFY?.label}
                  {typeof custPrev === "number" ? <>, versus {fmtNum(custPrev)} the prior FY-end</> : ""}
                  {typeof arpuCurr === "number" ? <>, with ARPU (in thousands) of <strong>{fmtRupeeThousands(arpuCurr)}</strong></> : ""}.</>
                ) : (
                  <>Customer count and ARPU aren't available yet for this workbook — add rows for these to the sheet
                  (matched by name or common synonym) to unlock this.</>
                )}
              </span>
            </div>
          </div>
        </div>

        <div className="narrative-stat-col">
          <div className="stat-tile">
            <div className="stat-tile__label">Latest quarter vs prior year</div>
            <div className="stat-tile__value">{latestQ ? fmtCr(latestQ[revKey]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestQ ? latestQ.label : "—"} {(ds.revenueLabel || "revenue").toLowerCase()} {revYoY !== null && <Delta curr={latestQ?.[revKey]} prev={prevYearQ?.[revKey]} good="up" />}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">FY EBITDA</div>
            <div className="stat-tile__value">{latestFY ? fmtCr(latestFY["EBITDA"]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestFY ? `${latestFY.label}${ebitdaTrend ? `, ${ebitdaTrend} from ${fmtCr(ebitdaPrev)} prior FY` : ""}` : "—"}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">Customers</div>
            <div className="stat-tile__value">{typeof custCurr === "number" ? fmtNum(custCurr) : "N/A"}</div>
            <div className="stat-tile__sub">{latestFY ? latestFY.label : "—"}</div>
          </div>
        </div>
      </div>
    </section>
  );
}

/* ============================================================
   FINBOX PERFORMANCE COMMENTARY — revenue growth and gross-margin
   evolution, EBITDA trajectory (merged two-vintage figure — see
   keyAliases on COMPANY_CONFIGS.finbox), a real product-mix bullet
   (One-time vs Recurring revenue split, computed from the sheet's
   own "-One time Revenue"/"-Recurring Revenue" rows where
   populated), and an explicit N/A callout for the embedded-finance
   mix KPIs the reference layout asks for but this MIS doesn't carry.
   ============================================================ */
function FinboxCommentary({ ds }) {
  const { latestQ, prevYearQ, latestFY, prevFY } = getExecStats(ds);
  const revKey = ds.revenueBaseKey;
  const kpiCfg = ds.companyConfig.kpi;

  const revYoY = latestQ && prevYearQ && typeof latestQ[revKey] === "number" && typeof prevYearQ[revKey] === "number" && prevYearQ[revKey]
    ? ((latestQ[revKey] - prevYearQ[revKey]) / Math.abs(prevYearQ[revKey])) * 100 : null;

  const ebitdaCurr = latestFY && typeof latestFY["EBITDA"] === "number" ? latestFY["EBITDA"] : null;
  const ebitdaPrev = prevFY && typeof prevFY["EBITDA"] === "number" ? prevFY["EBITDA"] : null;
  const ebitdaTrend = describeEbitdaTrend(ebitdaCurr, ebitdaPrev);

  const gmCurr = latestFY && typeof latestFY["Gross Margin"] === "number" ? latestFY["Gross Margin"] : null;
  const gmPrev = prevFY && typeof prevFY["Gross Margin"] === "number" ? prevFY["Gross Margin"] : null;

  const oneTimeKey = ds.kpiKeys.find(k => /^-One time Revenue$/i.test(k)) || null;
  const recurringKey = ds.kpiKeys.find(k => /^-Recurring Revenue$/i.test(k)) || null;
  const recurringCurr = recurringKey && latestFY && typeof latestFY[recurringKey] === "number" ? latestFY[recurringKey] : null;
  const revCurr = latestFY && typeof latestFY[revKey] === "number" ? latestFY[revKey] : null;
  const recurringShare = (typeof recurringCurr === "number" && typeof revCurr === "number" && revCurr) ? recurringCurr / revCurr : null;

  const opRows = (kpiCfg?.rows || []).map(r => ({ ...r, matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null }));
  const opAvailable = opRows.filter(r => r.matchedKey);

  return (
    <section className="section narrative-section">
      <div className="section__title">Performance Commentary</div>
      <div className="narrative-grid">
        <div className="narrative-card">
          <div className="narrative-card__eyebrow">Revenue &amp; profitability</div>
          <div className="narrative-bullets">
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><TrendingUp size={15} /></span>
              <span>
                {latestQ && prevYearQ && revYoY !== null ? (
                  <><strong>{latestQ.label}</strong> closed with {(ds.revenueLabel || "net revenue").toLowerCase()} of{" "}
                  <strong>{fmtCr(latestQ[revKey])}</strong>, {revYoY >= 0 ? "up" : "down"} {Math.abs(revYoY).toFixed(1)}% versus{" "}
                  {fmtCr(prevYearQ[revKey])} in {prevYearQ.label}
                  {typeof gmCurr === "number" ? <>, with gross margin at <strong>{fmtPct(gmCurr)}</strong>
                  {typeof gmPrev === "number" ? <> versus {fmtPct(gmPrev)} the prior FY</> : ""}</> : ""}.</>
                ) : (
                  <>{ds.revenueLabel || "Revenue"} data isn't complete enough yet to compute a like-for-like quarterly YoY comparison.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Rocket size={15} /></span>
              <span>
                {latestFY && prevFY && ebitdaTrend ? (
                  <>FY EBITDA {ebitdaTrend} from <strong>{fmtCr(ebitdaPrev)}</strong> in {prevFY.label} to{" "}
                  <strong>{fmtCr(ebitdaCurr)}</strong> in {latestFY.label} (merging the sheet's two MIS vintages,
                  latest/revised source preferred for overlapping months, per its own Read Me).</>
                ) : (
                  <>Not enough complete fiscal years of EBITDA yet to describe a trend.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Target size={15} /></span>
              <span>
                {typeof recurringShare === "number" ? (
                  <>Recurring revenue made up <strong>{fmtPct(recurringShare)}</strong> of {(ds.revenueLabel || "revenue").toLowerCase()}{" "}
                  in {latestFY?.label}, the balance one-time/project revenue. </>
                ) : (
                  <>Revenue-mix (recurring vs one-time) data isn't populated for the latest FY yet. </>
                )}
                {opAvailable.length ? (
                  <>{opAvailable.map((r, i) => (
                    <React.Fragment key={r.label}>{i > 0 ? "; " : ""}<strong>{r.label}</strong> stood at{" "}
                    <strong>{r.fmt(latestFY?.[r.matchedKey])}</strong> in {latestFY?.label}</React.Fragment>
                  ))}.</>
                ) : (
                  <>Embedded-finance product-mix metrics (Embedded Finance %, Device Connect %, Bank Connect %, Bureau
                  Connect %, MarketX/Sentinel %) aren't present in the uploaded standardized MIS — shown as N/A above
                  rather than estimated.</>
                )}
              </span>
            </div>
          </div>
        </div>

        <div className="narrative-stat-col">
          <div className="stat-tile">
            <div className="stat-tile__label">Latest quarter vs prior year</div>
            <div className="stat-tile__value">{latestQ ? fmtCr(latestQ[revKey]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestQ ? latestQ.label : "—"} {(ds.revenueLabel || "revenue").toLowerCase()} {revYoY !== null && <Delta curr={latestQ?.[revKey]} prev={prevYearQ?.[revKey]} good="up" />}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">FY EBITDA</div>
            <div className="stat-tile__value">{latestFY ? fmtCr(latestFY["EBITDA"]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestFY ? `${latestFY.label}${ebitdaTrend ? `, ${ebitdaTrend} from ${fmtCr(ebitdaPrev)} prior FY` : ""}` : "—"}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">Gross Margin</div>
            <div className="stat-tile__value">{typeof gmCurr === "number" ? fmtPct(gmCurr) : "N/A"}</div>
            <div className="stat-tile__sub">{latestFY ? latestFY.label : "—"}</div>
          </div>
        </div>
      </div>
    </section>
  );
}

/* ============================================================
   FUNDAMENTO PERFORMANCE COMMENTARY — pulse growth (Cumulative
   Pulses), revenue growth, EBITDA trajectory, and an operating-
   leverage bullet computed from real per-pulse economics (cost per
   pulse vs revenue per pulse) since this MIS has no distinct Gross
   Profit line (see companyConfig note: EBITDA = Total Revenue −
   Total Cost exactly, so "Gross Margin" isn't a separable figure
   here and is never shown as a guessed duplicate of EBITDA).
   ============================================================ */
function FundamentoCommentary({ ds }) {
  const { latestQ, prevYearQ, latestFY, prevFY } = getExecStats(ds);
  const revKey = ds.revenueBaseKey;

  const revYoY = latestQ && prevYearQ && typeof latestQ[revKey] === "number" && typeof prevYearQ[revKey] === "number" && prevYearQ[revKey]
    ? ((latestQ[revKey] - prevYearQ[revKey]) / Math.abs(prevYearQ[revKey])) * 100 : null;

  const ebitdaCurr = latestFY && typeof latestFY["EBITDA"] === "number" ? latestFY["EBITDA"] : null;
  const ebitdaPrev = prevFY && typeof prevFY["EBITDA"] === "number" ? prevFY["EBITDA"] : null;
  const ebitdaTrend = describeEbitdaTrend(ebitdaCurr, ebitdaPrev);
  const ebitdaMarginCurr = latestFY && typeof latestFY["EBITDA Margin"] === "number" ? latestFY["EBITDA Margin"] : null;
  const ebitdaMarginPrev = prevFY && typeof prevFY["EBITDA Margin"] === "number" ? prevFY["EBITDA Margin"] : null;

  const pulsesKey = ds.kpiKeys.find(k => /cum+ulative\s*pulses/i.test(k)) || null;
  const pulsesCurr = pulsesKey && latestFY && typeof latestFY[pulsesKey] === "number" ? latestFY[pulsesKey] : null;
  const pulsesPrev = pulsesKey && prevFY && typeof prevFY[pulsesKey] === "number" ? prevFY[pulsesKey] : null;
  const pulsesGrowth = (typeof pulsesCurr === "number" && typeof pulsesPrev === "number" && pulsesPrev)
    ? ((pulsesCurr - pulsesPrev) / Math.abs(pulsesPrev)) * 100 : null;

  const rppKey = ds.kpiKeys.find(k => /^revenue\s*per\s*pulse$/i.test(k)) || null;
  const rppCurr = rppKey && latestFY && typeof latestFY[rppKey] === "number" ? latestFY[rppKey] : null;
  const rppPrev = rppKey && prevFY && typeof prevFY[rppKey] === "number" ? prevFY[rppKey] : null;
  const totalCostCurr = latestFY && typeof latestFY["Total Cost"] === "number" ? latestFY["Total Cost"] : null;
  const costPerPulseCurr = (typeof totalCostCurr === "number" && typeof pulsesCurr === "number" && pulsesCurr)
    ? totalCostCurr / pulsesCurr : null;
  const totalCostPrev = prevFY && typeof prevFY["Total Cost"] === "number" ? prevFY["Total Cost"] : null;
  const costPerPulsePrev = (typeof totalCostPrev === "number" && typeof pulsesPrev === "number" && pulsesPrev)
    ? totalCostPrev / pulsesPrev : null;

  return (
    <section className="section narrative-section">
      <div className="section__title">Performance Commentary</div>
      <div className="narrative-grid">
        <div className="narrative-card">
          <div className="narrative-card__eyebrow">Revenue &amp; profitability</div>
          <div className="narrative-bullets">
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><TrendingUp size={15} /></span>
              <span>
                {latestQ && prevYearQ && revYoY !== null ? (
                  <><strong>{latestQ.label}</strong> closed with {(ds.revenueLabel || "revenue").toLowerCase()} of{" "}
                  <strong>{fmtCr(latestQ[revKey])}</strong>, {revYoY >= 0 ? "up" : "down"} {Math.abs(revYoY).toFixed(1)}% versus{" "}
                  {fmtCr(prevYearQ[revKey])} in {prevYearQ.label}.</>
                ) : (
                  <>{ds.revenueLabel || "Revenue"} data isn't complete enough yet to compute a like-for-like quarterly YoY comparison.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Rocket size={15} /></span>
              <span>
                {typeof pulsesCurr === "number" ? (
                  <>Cumulative Pulses {pulsesGrowth !== null ? <>{pulsesGrowth >= 0 ? "grew" : "declined"} <strong>{Math.abs(pulsesGrowth).toFixed(1)}%</strong> YoY to</> : "reached"}{" "}
                  <strong>{fmtNum(pulsesCurr)}</strong> in {latestFY?.label}
                  {typeof pulsesPrev === "number" ? <>, from {fmtNum(pulsesPrev)} in {prevFY?.label}</> : ""}
                  {typeof rppCurr === "number" ? <>, at a realised revenue of <strong>{fmtRupee(rppCurr)}</strong> per pulse
                  {typeof rppPrev === "number" ? <> (versus {fmtRupee(rppPrev)} the prior FY)</> : ""}</> : ""}.</>
                ) : (
                  <>Pulse-volume data isn't complete enough yet to compute a YoY comparison.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><ShieldCheck size={15} /></span>
              <span>
                {latestFY && prevFY && ebitdaTrend ? (
                  <>FY EBITDA {ebitdaTrend} from <strong>{fmtCr(ebitdaPrev)}</strong> in {prevFY.label} to{" "}
                  <strong>{fmtCr(ebitdaCurr)}</strong> in {latestFY.label}
                  {typeof ebitdaMarginCurr === "number" ? <>, EBITDA margin at <strong>{fmtPct(ebitdaMarginCurr)}</strong>
                  {typeof ebitdaMarginPrev === "number" ? <> versus {fmtPct(ebitdaMarginPrev)} prior FY</> : ""}</> : ""}
                  {(typeof costPerPulseCurr === "number" && typeof costPerPulsePrev === "number") ? <>. Cost per pulse
                  {costPerPulseCurr < costPerPulsePrev ? " declined" : costPerPulseCurr > costPerPulsePrev ? " rose" : " held steady"} to{" "}
                  <strong>{fmtRupee(costPerPulseCurr)}</strong> from {fmtRupee(costPerPulsePrev)} the prior FY — the clearest read on operating
                  leverage this MIS supports (it carries Total Revenue, Total Cost and EBITDA only, with no separate
                  Gross Profit line, so Gross Margin is intentionally not shown as a distinct metric)</> : ""}.</>
                ) : (
                  <>Not enough complete fiscal years of EBITDA yet to describe a trend.</>
                )}
              </span>
            </div>
          </div>
        </div>

        <div className="narrative-stat-col">
          <div className="stat-tile">
            <div className="stat-tile__label">Latest quarter vs prior year</div>
            <div className="stat-tile__value">{latestQ ? fmtCr(latestQ[revKey]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestQ ? latestQ.label : "—"} {(ds.revenueLabel || "revenue").toLowerCase()} {revYoY !== null && <Delta curr={latestQ?.[revKey]} prev={prevYearQ?.[revKey]} good="up" />}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">FY EBITDA</div>
            <div className="stat-tile__value">{latestFY ? fmtCr(latestFY["EBITDA"]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestFY ? `${latestFY.label}${ebitdaTrend ? `, ${ebitdaTrend} from ${fmtCr(ebitdaPrev)} prior FY` : ""}` : "—"}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">Cumulative Pulses</div>
            <div className="stat-tile__value">{typeof pulsesCurr === "number" ? fmtNum(pulsesCurr) : "N/A"}</div>
            <div className="stat-tile__sub">{latestFY ? `${latestFY.label}${pulsesGrowth !== null ? `, ${pulsesGrowth >= 0 ? "+" : ""}${pulsesGrowth.toFixed(1)}% YoY` : ""}` : "—"}</div>
          </div>
        </div>
      </div>
    </section>
  );
}

/* ============================================================
   LEEGALITY PERFORMANCE COMMENTARY — revenue/EBITDA-margin
   methodology as every other company, plus an eSign/stamp/
   subscription-accounts bullet layered in per the spec.
   ============================================================ */
function LeegalityCommentary({ ds }) {
  const { latestQ, prevYearQ, latestFY, prevFY } = getExecStats(ds);
  const revKey = ds.revenueBaseKey;
  const kpiCfg = ds.companyConfig.kpi;

  const revYoY = latestQ && prevYearQ && typeof latestQ[revKey] === "number" && typeof prevYearQ[revKey] === "number" && prevYearQ[revKey]
    ? ((latestQ[revKey] - prevYearQ[revKey]) / Math.abs(prevYearQ[revKey])) * 100 : null;

  const ebitdaCurr = latestFY && typeof latestFY["EBITDA"] === "number" ? latestFY["EBITDA"] : null;
  const ebitdaPrev = prevFY && typeof prevFY["EBITDA"] === "number" ? prevFY["EBITDA"] : null;
  const ebitdaTrend = describeEbitdaTrend(ebitdaCurr, ebitdaPrev);

  const marginPick = bestMarginKey(ds);
  const marginFYIdx = latestFY ? ds.fyData.findIndex(f => f.key === latestFY.key) : -1;
  const marginPrevFY = marginFYIdx > 0 ? ds.fyData[marginFYIdx - 1] : null;
  const marginCurr = marginPick && latestFY ? latestFY[marginPick[0]] : null;
  const marginPrev = marginPick && marginPrevFY ? marginPrevFY[marginPick[0]] : null;

  const esignRow = (kpiCfg?.rows || []).find(r => r.slug === "esigns");
  const stampRow = (kpiCfg?.rows || []).find(r => r.slug === "stamps");
  const esignKey = esignRow ? ds.kpiKeys.find(k => esignRow.matchers.some(re => re.test(k))) : null;
  const stampKey = stampRow ? ds.kpiKeys.find(k => stampRow.matchers.some(re => re.test(k))) : null;
  const subKey = kpiCfg?.stockRow ? ds.kpiKeys.find(k => kpiCfg.stockRow.matchers.some(re => re.test(k))) : null;

  const esignGrowthIdx = ds.qData.length - 1;
  const esignGrowth = esignKey ? periodGrowth(ds.qData, esignGrowthIdx, esignKey, true) : null;
  const stampGrowth = stampKey ? periodGrowth(ds.qData, esignGrowthIdx, stampKey, true) : null;
  const latestQEsign = esignKey && latestQ ? latestQ[esignKey] : null;
  const latestQStamp = stampKey && latestQ ? latestQ[stampKey] : null;
  const subCurr = subKey && latestQ && typeof latestQ[subKey] === "number" ? latestQ[subKey] : null;

  return (
    <section className="section narrative-section">
      <div className="section__title">Performance Commentary</div>
      <div className="narrative-grid">
        <div className="narrative-card">
          <div className="narrative-card__eyebrow">Growth &amp; profitability</div>
          <div className="narrative-bullets">
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><TrendingUp size={15} /></span>
              <span>
                {latestQ && prevYearQ && revYoY !== null ? (
                  <><strong>{latestQ.label}</strong> closed with {(ds.revenueLabel || "net revenue").toLowerCase()} of{" "}
                  <strong>{fmtCr(latestQ[revKey])}</strong>, {revYoY >= 0 ? "up" : "down"} {Math.abs(revYoY).toFixed(1)}% versus{" "}
                  {fmtCr(prevYearQ[revKey])} in {prevYearQ.label}.</>
                ) : (
                  <>{ds.revenueLabel || "Revenue"} data isn't complete enough yet to compute a like-for-like quarterly YoY comparison.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Rocket size={15} /></span>
              <span>
                {latestFY && prevFY && ebitdaTrend ? (
                  <>FY EBITDA {ebitdaTrend} from <strong>{fmtCr(ebitdaPrev)}</strong> in {prevFY.label} to{" "}
                  <strong>{fmtCr(ebitdaCurr)}</strong> in {latestFY.label}
                  {marginPick && typeof marginCurr === "number" ? <>, with {marginPick[0].toLowerCase()} at <strong>{fmtPct(marginCurr)}</strong>
                  {typeof marginPrev === "number" ? <> versus {fmtPct(marginPrev)} the prior FY</> : ""}</> : ""}.</>
                ) : (
                  <>Not enough complete fiscal years of EBITDA yet to describe a trend.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Layers size={15} /></span>
              <span>
                {(typeof latestQEsign === "number" || typeof latestQStamp === "number" || typeof subCurr === "number") ? (
                  <>
                    {typeof latestQEsign === "number" && (
                      <>{latestQ.label} logged <strong>{fmtNum(latestQEsign)}</strong> eSigns
                      {esignGrowth !== null ? <> ({esignGrowth >= 0 ? "+" : ""}{(esignGrowth * 100).toFixed(1)}% YoY)</> : ""}. </>
                    )}
                    {typeof latestQStamp === "number" && (
                      <><strong>{fmtNum(latestQStamp)}</strong> stamps were ordered
                      {stampGrowth !== null ? <> ({stampGrowth >= 0 ? "+" : ""}{(stampGrowth * 100).toFixed(1)}% YoY)</> : ""}. </>
                    )}
                    {typeof subCurr === "number" && <>Active subscription accounts stood at <strong>{fmtNum(subCurr)}</strong> as of {latestQ.label}.</>}
                  </>
                ) : (
                  <>eSign / stamp / subscription-account data isn't complete enough yet for this period.</>
                )}
              </span>
            </div>
          </div>
        </div>

        <div className="narrative-stat-col">
          <div className="stat-tile">
            <div className="stat-tile__label">Latest quarter vs prior year</div>
            <div className="stat-tile__value">{latestQ ? fmtCr(latestQ[revKey]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestQ ? latestQ.label : "—"} {(ds.revenueLabel || "revenue").toLowerCase()} {revYoY !== null && <Delta curr={latestQ?.[revKey]} prev={prevYearQ?.[revKey]} good="up" />}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">FY EBITDA</div>
            <div className="stat-tile__value">{latestFY ? fmtCr(latestFY["EBITDA"]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestFY ? `${latestFY.label}${ebitdaTrend ? `, ${ebitdaTrend} from ${fmtCr(ebitdaPrev)} prior FY` : ""}` : "—"}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">No. of e-signs</div>
            <div className="stat-tile__value">{typeof latestQEsign === "number" ? fmtNum(latestQEsign) : "N/A"}</div>
            <div className="stat-tile__sub">{latestQ ? latestQ.label : "—"}</div>
          </div>
        </div>
      </div>
    </section>
  );
}

const CHIP_ICONS = [Layers, ShieldCheck, Target, Info];
const SCALE_ICONS = [Building2, Store, Users];

function BusinessDescription({ companyInfo }) {
  if (!companyInfo) return null;
  const { description, tags, scaleMetrics } = companyInfo;
  if (!description && !tags.length && !scaleMetrics.length) return null;

  return (
    <section className="section narrative-section">
      <div className="section__title">Business Description</div>
      <div className="biz-card">
        {description && <p className="biz-desc-text">{description}</p>}

        {tags.length > 0 && (
          <div className="biz-chip-row">
            {tags.map((tag, i) => {
              const Icon = CHIP_ICONS[i % CHIP_ICONS.length];
              return (
                <span className="biz-chip" key={tag}>
                  <Icon size={11} style={{ marginRight: 5, verticalAlign: -2 }} />{tag}
                </span>
              );
            })}
          </div>
        )}

        {scaleMetrics.length > 0 && (
          <div className="biz-scale-row">
            {scaleMetrics.map((sm, i) => {
              const Icon = SCALE_ICONS[i % SCALE_ICONS.length];
              return (
                <div className="biz-scale-tile" key={sm.label}>
                  <Icon size={22} className="biz-scale-icon" />
                  <div>
                    <div className="biz-scale-value">{sm.value}</div>
                    <div className="biz-scale-label">{sm.label}</div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </section>
  );
}

function EmptyState({ onFile, status }) {
  const [dragOver, setDragOver] = useState(false);
  return (
    <div className="empty-wrap">
      <div className="empty-eyebrow">Consolidated MIS</div>
      <div className="empty-title">MIS Dashboard</div>
      <div className="empty-sub">Upload your company's mastersheet to build the dashboard. Nothing renders until real data arrives.</div>

      <label
        className={`dropzone ${dragOver ? "dropzone--over" : ""}`}
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => { e.preventDefault(); setDragOver(false); if (e.dataTransfer.files?.[0]) onFile(e.dataTransfer.files[0]); }}
      >
        <Upload size={28} strokeWidth={1.5} />
        <div className="dropzone__title">Drop your mastersheet here, or click to browse</div>
        <div className="dropzone__sub">.xlsx — needs a "Monthly Data" sheet (KPI rows × month columns)</div>
        <input type="file" accept=".xlsx,.xls" onChange={(e) => e.target.files?.[0] && onFile(e.target.files[0])} />
      </label>

      {status && (
        <div className={`empty-status empty-status--${status.type}`}>
          {status.type === "error" && <AlertTriangle size={14} />}
          {status.text}
        </div>
      )}

      <div className="empty-hint">
        Don't have one yet? Duplicate the mastersheet template in the <code>excel_template</code> folder — fill in
        your company's KPIs on the "Monthly Data" sheet and (optionally) your company profile on the "Company Info"
        sheet, keeping row labels exactly as they are, then upload it here. Each new quarter, add columns and re-upload.
      </div>
    </div>
  );
}

export default function App() {
  const [dataset, setDataset] = useState(null);
  const [fyIndex, setFyIndex] = useState(null);
  const [expanded, setExpanded] = useState(null);
  const [status, setStatus] = useState(null);
  const [fileName, setFileName] = useState(null);
  const [section, setSection] = useState("performance");

  const onToggle = useCallback((key) => setExpanded(e => (e === key ? null : key)), []);

  const handleFile = useCallback(async (file) => {
    setStatus({ type: "info", text: `Reading ${file.name}…` });
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const parsed = parseWorkbook(wb);
      // Identify which company's template this is from the row names actually
      // present — never from the filename — so the right KPI/financial/commentary
      // mappings apply. An unrecognized sheet falls back to the original
      // (Easyrewardz-shaped) generic behaviour, unchanged from before GrayQuest.
      const companyConfig = detectCompanyConfig(Object.keys(parsed.kpis));
      const sheetCompanyInfo = parseCompanyInfo(wb);
      // If the workbook has no Company Info sheet, fall back to this company's
      // static profile text (if one is configured) rather than leaving the
      // Business Description section blank — this is user-supplied profile copy,
      // never derived from or mixed with the financial figures.
      const companyInfo = sheetCompanyInfo || companyConfig.defaultDescription || null;
      // News Feed / Industry Data are a separate, independent pipeline from the
      // financial Monthly Data parse above — a missing or malformed sheet here
      // must never affect Performance section parsing.
      let newsFeed = null, industryData = null, refreshMeta = null;
      try { newsFeed = parseNewsSheet(wb); } catch { newsFeed = null; }
      try { industryData = parseIndustrySheet(wb); } catch { industryData = null; }
      try { refreshMeta = parseRefreshMeta(wb); } catch { refreshMeta = null; }
      const ds = buildDataset(parsed, companyInfo, companyConfig);
      ds.newsFeed = newsFeed;
      ds.industryData = industryData;
      ds.refreshMeta = refreshMeta;
      setDataset(ds);
      setFyIndex(ds.fyData.length - 1);
      setFileName(file.name);
      setStatus({ type: "success", text: `Loaded ${ds.months.length} months, ${ds.kpiKeys.length} KPIs, ${ds.fyData.length} FY periods from "${parsed.sheetName}" (${companyConfig.displayName || companyInfo?.companyName || "company"} template).` });
    } catch (err) {
      setStatus({ type: "error", text: String(err.message || err) });
    }
  }, []);

  const handleReplace = useCallback(async (e) => {
    const file = e.target.files?.[0];
    if (file) await handleFile(file);
  }, [handleFile]);

  useEffect(() => {
    document.title = dataset?.companyInfo?.companyName ? `${dataset.companyInfo.companyName} — MIS Dashboard` : "MIS Dashboard";
  }, [dataset]);

  if (!dataset) {
    return (
      <div className="dash">
        <GlobalStyles />
        <EmptyState onFile={handleFile} status={status} />
      </div>
    );
  }

  const fy = dataset.fyData[fyIndex];

  return (
    <div className="dash">
      <GlobalStyles />

      <header className="masthead">
        <div>
          <div className="masthead__eyebrow">{dataset.companyInfo?.companyName ? `${dataset.companyInfo.companyName} · ` : ""}Consolidated MIS</div>
          <div className="masthead__title">MIS Dashboard</div>
          <div className="masthead__meta">
            {fileName} · {dataset.months[0].label} → {dataset.months[dataset.months.length - 1].label} · viewing {fy.label} ({fy.sub})
          </div>
        </div>
        <div className="masthead__actions">
          <label className="replace-btn">
            <RefreshCw size={13} /> New upload
            <input type="file" accept=".xlsx,.xls" onChange={handleReplace} />
          </label>
        </div>
      </header>

      <nav className="dash-nav">
        <button className={`dash-nav__tab ${section === "performance" ? "dash-nav__tab--active" : ""}`} onClick={() => setSection("performance")}>Performance</button>
        <button className={`dash-nav__tab ${section === "industry" ? "dash-nav__tab--active" : ""}`} onClick={() => setSection("industry")}>Industry &amp; Competitors</button>
        <button className={`dash-nav__tab ${section === "news" ? "dash-nav__tab--active" : ""}`} onClick={() => setSection("news")}>News &amp; Updates</button>
      </nav>

      {status && status.type !== "success" && (
        <div className={`upload-bar upload-bar--${status.type}`}>
          <FileSpreadsheet size={14} /><span>{status.text}</span>
        </div>
      )}

      {section === "performance" && (
        <>
          <BusinessDescription companyInfo={dataset.companyInfo} />

          {dataset.companyConfig.layout === "grayquest" ? (
            <>
              <GrayQuestKPITable ds={dataset} />
              <RevenueProfitabilityTable ds={dataset} title="Key Financial Highlights" />
              <GrayQuestCommentary ds={dataset} />
            </>
          ) : dataset.companyConfig.layout === "riskcovry" ? (
            <>
              <RevenueProfitabilityTable ds={dataset} />
              <RiskcovryCommentary ds={dataset} />
              <RiskcovryKPITable ds={dataset} />
            </>
          ) : dataset.companyConfig.layout === "multipl" ? (
            <>
              <MultiplKPITable ds={dataset} />
              <RevenueProfitabilityTable ds={dataset} title="Key Financial Highlights" />
              <MultiplCommentary ds={dataset} />
            </>
          ) : dataset.companyConfig.layout === "fastsurance" ? (
            <>
              <RevenueProfitabilityTable ds={dataset} title="Key Financial Highlights" />
              <FastsuranceCommentary ds={dataset} />
              <FastsuranceKPITable ds={dataset} />
            </>
          ) : dataset.companyConfig.layout === "apexFutureLabs" ? (
            <>
              <ApexFutureLabsKPITable ds={dataset} />
              <RevenueProfitabilityTable ds={dataset} title="Key Financial Highlights" />
              <VitraCommentary ds={dataset} />
            </>
          ) : dataset.companyConfig.layout === "leegality" ? (
            <>
              <LeegalityKPITable ds={dataset} />
              <RevenueProfitabilityTable ds={dataset} title="Key Financial Highlights" />
              <LeegalityCommentary ds={dataset} />
            </>
          ) : dataset.companyConfig.layout === "finbox" ? (
            <>
              <RevenueProfitabilityTable ds={dataset} title="Key Financial Highlights" />
              <FinboxCommentary ds={dataset} />
              <FinboxKPITable ds={dataset} />
            </>
          ) : dataset.companyConfig.layout === "fundamento" ? (
            <>
              <FundamentoKPITable ds={dataset} />
              <RevenueProfitabilityTable ds={dataset} title="Key Financial Highlights" />
              <FundamentoCommentary ds={dataset} />
            </>
          ) : (
            <>
              <RevenueProfitabilityTable ds={dataset} />
              <ExecutiveSummary ds={dataset} />
              <KeyPerformanceIndicatorsTable ds={dataset} />
            </>
          )}

          <section className="section">
            <div className="fin-section__head">
              <div className="section__title">Key Metrics <span className="section__sub">— {fy.label} vs {fyIndex > 0 ? dataset.fyData[fyIndex - 1].label : "—"}</span></div>
              <div className="fy-tabs">
                {dataset.fyData.map((f, i) => (
                  <button key={f.key} className={`fy-tab ${i === fyIndex ? "fy-tab--active" : ""}`} onClick={() => setFyIndex(i)}>{f.label}</button>
                ))}
              </div>
            </div>
            <div className="kpi-grid">
              {dataset.cardConfigs.map(cfg => (
                <KpiCard key={cfg.key} cfg={cfg} ds={dataset} fyIndex={fyIndex} expanded={expanded} onToggle={onToggle} />
              ))}
            </div>
          </section>

          {expanded && (
            <DrillDownModal
              cfg={dataset.cardConfigs.find(c => c.key === expanded)}
              ds={dataset}
              fyIndex={fyIndex}
              onClose={() => setExpanded(null)}
            />
          )}

          <section className="section">
            <div className="section__title">Performance &amp; Mix</div>
            <div className="chart-grid">
              <div className="chart-card">
                <div className="chart-card__title">Revenue Trend</div>
                <div className="chart-card__note">Monthly, {dataset.months[0].label} → {dataset.months[dataset.months.length - 1].label}</div>
                <RevenueTrendChart ds={dataset} />
              </div>
              <div className="chart-card">
                <div className="chart-card__title">Revenue Mix</div>
                <div className="chart-card__note">By revenue sub-line, per FY</div>
                <RevenueMixChart ds={dataset} />
              </div>
              <div className="chart-card">
                <div className="chart-card__title">Profitability Trend</div>
                <div className="chart-card__note">Revenue (bars) vs EBITDA &amp; Net Profit (lines), by FY</div>
                <ProfitabilityChart ds={dataset} />
              </div>
              <div className="chart-card">
                <div className="chart-card__title">Margin Trend</div>
                <div className="chart-card__note">Gross / EBITDA / Net margin, by FY</div>
                <MarginTrendChart ds={dataset} />
              </div>
            </div>
          </section>

          {dataset.companyConfig.showForecast && (
            <section className="section">
              <div className="section__title">Outlook &amp; Forecast <span className="section__sub">— quarterly, trend-projected</span></div>
              <div className="chart-grid">
                <div className="chart-card">
                  <div className="chart-card__title">Quarterly Revenue — YoY</div>
                  <div className="chart-card__note">Last 8 quarters; latest quarter highlighted in gold</div>
                  <QuarterlyRevenueChart ds={dataset} />
                </div>
                <div className="chart-card">
                  <div className="chart-card__title">EBITDA Turnaround &amp; Forecast</div>
                  <div className="chart-card__note">Green = positive, red = negative, dashed = projected</div>
                  <EbitdaTurnaroundChart ds={dataset} />
                </div>
                <div className="chart-card" style={{ gridColumn: "1 / -1" }}>
                  <div className="chart-card__title">Revenue Trend &amp; Forecast</div>
                  <div className="chart-card__note">Linear trend on the last up to 8 complete quarters, projected 2 quarters forward</div>
                  <RevenueForecastChart ds={dataset} />
                </div>
              </div>
            </section>
          )}

          <section className="section">
            <div className="section__title">Cash &amp; Operations</div>
            <div className="chart-grid">
              <div className="chart-card" style={{ gridColumn: "1 / -1" }}>
                <div className="chart-card__title">Headcount &amp; Productivity</div>
                <div className="chart-card__note">Average headcount (bars) vs revenue per employee (line), by FY</div>
                <HeadcountChart ds={dataset} />
              </div>
            </div>
          </section>

          <ProfitAndLossTable ds={dataset} />

          <div className="footnote">
            <Info size={13} style={{ flexShrink: 0, marginTop: 2 }} />
            <span>
              Every card, chart, table, narrative, and FY/quarter grouping above is computed live from the uploaded
              sheet — nothing is hardcoded. The Business Description section and company name come from the optional
              "Company Info" sheet, if one is present. FY periods are derived from the dates in row 1 (Apr–Mar), so
              adding a new quarter's columns and re-uploading is all a refresh needs. Forecast lines in the Outlook
              section are a simple linear trend over the most recent complete quarters, not a modeled projection —
              treat them as directional only.
            </span>
          </div>
        </>
      )}

      {section === "industry" && <IndustryCompetitorsPage ds={dataset} />}

      {section === "news" && <NewsUpdatesPage ds={dataset} />}
    </div>
  );
}

function GlobalStyles() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

      .dash { --ink:#14171F; --muted:#6B7280; --border:#E5E7EB; --surface:#F7F8FA; --bg:#FFFFFF;
              --brand:#1D4E4A; --gold:#B08A3E; --pos:#16A34A; --neg:#B3492F; --sage:#7C9885;
              font-family:'Inter',sans-serif; color:var(--ink); background:var(--bg);
              min-height:100vh; padding:0 0 64px 0; }
      .dash * { box-sizing:border-box; }

      .empty-wrap { max-width:640px; margin:0 auto; padding:96px 24px 40px; text-align:center; }
      .empty-eyebrow { font-family:'IBM Plex Mono',monospace; font-size:11px; letter-spacing:0.12em; text-transform:uppercase; color:var(--gold); margin-bottom:8px; }
      .empty-title { font-family:'Fraunces',serif; font-size:36px; font-weight:500; letter-spacing:-0.01em; margin-bottom:10px; }
      .empty-sub { font-size:14px; color:var(--muted); margin-bottom:36px; }
      .dropzone { display:flex; flex-direction:column; align-items:center; gap:8px; border:1.5px dashed var(--border); border-radius:16px;
                  padding:48px 24px; cursor:pointer; color:var(--brand); transition:border-color .15s, background .15s; }
      .dropzone:hover, .dropzone--over { border-color:var(--brand); background:var(--surface); }
      .dropzone input { display:none; }
      .dropzone__title { font-size:14.5px; font-weight:600; color:var(--ink); margin-top:6px; }
      .dropzone__sub { font-family:'IBM Plex Mono',monospace; font-size:11.5px; color:var(--muted); }
      .empty-status { margin-top:18px; font-family:'IBM Plex Mono',monospace; font-size:12px; padding:10px 14px; border-radius:8px; background:var(--surface); color:var(--muted); display:flex; align-items:center; gap:8px; justify-content:center; text-align:left; }
      .empty-status--error { color:var(--neg); background:#FBF0EC; }
      .empty-status--success { color:var(--pos); background:#EFF8F1; }
      .empty-hint { margin-top:28px; font-size:12px; color:var(--muted); line-height:1.6; }

      .masthead { border-bottom:1px solid var(--border); padding:28px 40px 22px; display:flex; justify-content:space-between; align-items:flex-end; flex-wrap:wrap; gap:16px; }
      .masthead__eyebrow { font-family:'IBM Plex Mono',monospace; font-size:11px; letter-spacing:0.12em; text-transform:uppercase; color:var(--gold); margin-bottom:6px; }
      .masthead__title { font-family:'Fraunces',serif; font-size:32px; font-weight:500; letter-spacing:-0.01em; line-height:1.1; }
      .masthead__meta { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--muted); margin-top:6px; }
      .masthead__actions { display:flex; align-items:center; gap:10px; flex-wrap:wrap; }

      .fy-tabs { display:flex; gap:4px; background:var(--surface); padding:4px; border-radius:10px; border:1px solid var(--border); }
      .fy-tab { font-family:'IBM Plex Mono',monospace; font-size:12px; padding:8px 14px; border-radius:7px; border:none; background:transparent; color:var(--muted); cursor:pointer; transition:all .15s; white-space:nowrap; }
      .fy-tab:hover { color:var(--ink); }
      .fy-tab--active { background:var(--brand); color:#fff; }

      .replace-btn { display:flex; align-items:center; gap:6px; font-family:'IBM Plex Mono',monospace; font-size:11.5px; color:var(--muted);
                     border:1px solid var(--border); border-radius:9px; padding:9px 12px; cursor:pointer; transition:border-color .15s, color .15s; }
      .replace-btn:hover { border-color:var(--brand); color:var(--brand); }
      .replace-btn input { display:none; }

      .upload-bar { margin:18px 40px 0; padding:10px 16px; border-radius:10px; display:flex; align-items:center; gap:10px; font-size:12.5px; }
      .upload-bar--info { background:var(--surface); color:var(--muted); }
      .upload-bar--error { background:#FBF0EC; color:var(--neg); }

      .section { padding:32px 40px 8px; }
      .section__title { font-family:'Fraunces',serif; font-size:19px; font-weight:500; margin-bottom:16px; display:flex; align-items:baseline; gap:10px; }
      .section__sub { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--muted); }

      .narrative-section { padding-top:28px; }
      .narrative-grid { display:grid; grid-template-columns:1.3fr 1fr; gap:20px; }
      @media (max-width:900px) { .narrative-grid { grid-template-columns:1fr; } }
      .narrative-card { border:1px solid #DDD3BC; border-radius:14px; padding:22px 24px; background:linear-gradient(180deg,#FFFDFA,#fff); }
      .narrative-card__eyebrow { font-family:'IBM Plex Mono',monospace; font-size:10.5px; text-transform:uppercase; letter-spacing:0.08em; color:var(--gold); margin-bottom:14px; }
      .narrative-bullets { display:flex; flex-direction:column; gap:16px; }
      .narrative-bullet { display:flex; gap:12px; align-items:flex-start; font-size:13.5px; line-height:1.65; color:var(--ink); }
      .narrative-bullet__icon { flex-shrink:0; width:28px; height:28px; border-radius:8px; display:flex; align-items:center; justify-content:center; background:var(--surface); color:var(--brand); margin-top:1px; }
      .narrative-extra-note { display:flex; align-items:flex-start; gap:6px; margin-top:16px; padding-top:16px; border-top:1px solid var(--border); font-size:12px; color:var(--muted); line-height:1.6; }
      .narrative-stat-col { display:flex; flex-direction:column; gap:14px; }
      .stat-tile { border:1px solid var(--border); border-radius:12px; padding:16px; background:#fff; }
      .stat-tile__label { font-family:'IBM Plex Mono',monospace; font-size:10.5px; text-transform:uppercase; letter-spacing:0.06em; color:var(--muted); margin-bottom:6px; }
      .stat-tile__value { font-family:'Fraunces',serif; font-size:22px; font-weight:500; }
      .stat-tile__sub { font-size:11px; color:var(--muted); margin-top:6px; display:flex; align-items:center; gap:6px; flex-wrap:wrap; }

      .biz-card { border:1px solid var(--border); border-radius:14px; padding:22px 24px; background:#fff; }
      .biz-desc-text { font-size:13.5px; line-height:1.75; color:var(--ink); margin:0 0 18px; }
      .biz-chip-row { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:20px; }
      .biz-chip { display:inline-flex; align-items:center; font-family:'IBM Plex Mono',monospace; font-size:11px; padding:7px 12px; border-radius:20px; border:1px solid var(--border); color:var(--brand); background:var(--surface); }
      .biz-scale-row { display:flex; gap:14px; flex-wrap:wrap; }
      .biz-scale-tile { flex:1; min-width:150px; border:1px solid #DDD3BC; border-radius:12px; padding:16px; background:linear-gradient(180deg,#FFFDFA,#fff); display:flex; align-items:center; gap:12px; }
      .biz-scale-icon { color:var(--gold); flex-shrink:0; }
      .biz-scale-value { font-family:'Fraunces',serif; font-size:19px; font-weight:500; line-height:1.2; }
      .biz-scale-label { font-family:'IBM Plex Mono',monospace; font-size:10px; color:var(--muted); text-transform:uppercase; letter-spacing:0.05em; margin-top:2px; }

      .forecast-note { display:flex; align-items:center; gap:6px; font-family:'IBM Plex Mono',monospace; font-size:10.5px; color:var(--muted); margin:6px 0 10px; flex-wrap:wrap; }
      .forecast-dot { width:8px; height:8px; border-radius:2px; display:inline-block; }
      .forecast-dot--actual { background:var(--brand); }
      .forecast-dot--proj { background:var(--gold); }

      .kpi-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; }
      @media (max-width:1100px) { .kpi-grid { grid-template-columns:repeat(2,1fr); } }
      @media (max-width:600px) { .kpi-grid { grid-template-columns:1fr; } }

      .kpi-card { all:unset; display:block; box-sizing:border-box; border:1px solid var(--border); border-radius:12px; background:#fff;
                  padding:16px 16px 14px; cursor:pointer; transition:box-shadow .15s, border-color .15s, transform .1s; }
      .kpi-card:hover { border-color:#D5D9E0; box-shadow:0 2px 10px rgba(20,23,31,0.05); transform:translateY(-1px); }
      .kpi-card:active { transform:translateY(0px) scale(0.99); }
      .kpi-card--primary { border-color:#DDD3BC; background:linear-gradient(180deg,#FFFDFA,#fff); }
      .kpi-card--active { border-color:var(--brand); box-shadow:0 0 0 3px rgba(29,78,74,0.12); }
      .kpi-card__top { display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; }
      .kpi-card__label { font-size:12px; color:var(--muted); font-weight:500; }
      .kpi-card__chev { color:var(--muted); transform:rotate(-90deg); }
      .kpi-card__value { font-family:'Fraunces',serif; font-size:26px; font-weight:500; letter-spacing:-0.01em; margin-bottom:10px; text-align:left; }
      .kpi-card__foot { display:flex; align-items:center; justify-content:space-between; gap:8px; }
      .kpi-card__spark { width:72px; height:36px; flex-shrink:0; opacity:0.8; }

      .delta { display:inline-flex; align-items:center; gap:3px; font-family:'IBM Plex Mono',monospace; font-size:11.5px; font-weight:500; padding:3px 7px; border-radius:6px; }
      .delta-pos { color:var(--pos); background:#EFF8F1; }
      .delta-neg { color:var(--neg); background:#FBF0EC; }
      .delta-flat { color:var(--muted); background:var(--surface); }

      .modal-backdrop { position:fixed; inset:0; background:rgba(20,23,31,0.45); backdrop-filter:blur(3px);
                         display:flex; align-items:center; justify-content:center; padding:24px; z-index:1000; animation:fadeIn .15s ease; }
      @keyframes fadeIn { from { opacity:0; } to { opacity:1; } }
      .modal-panel { background:#fff; border-radius:16px; width:100%; max-width:620px; max-height:88vh; overflow-y:auto;
                     padding:24px 28px 28px; box-shadow:0 24px 60px rgba(20,23,31,0.25); animation:popIn .18s cubic-bezier(.2,.9,.3,1); }
      @keyframes popIn { from { opacity:0; transform:scale(0.96) translateY(6px); } to { opacity:1; transform:scale(1) translateY(0); } }
      .modal-panel__header { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:18px; }
      .modal-panel__eyebrow { font-family:'IBM Plex Mono',monospace; font-size:10.5px; text-transform:uppercase; letter-spacing:0.08em; color:var(--gold); margin-bottom:4px; }
      .modal-panel__title { font-family:'Fraunces',serif; font-size:22px; font-weight:500; }
      .modal-panel__close { all:unset; cursor:pointer; color:var(--muted); padding:6px; border-radius:8px; }
      .modal-panel__close:hover { background:var(--surface); color:var(--ink); }
      .modal-panel__value-row { display:flex; align-items:center; gap:10px; margin-bottom:22px; flex-wrap:wrap; }
      .modal-panel__value { font-family:'Fraunces',serif; font-size:34px; font-weight:500; }
      .modal-panel__vs { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--muted); }

      .drawer-subhead { font-family:'IBM Plex Mono',monospace; font-size:10.5px; text-transform:uppercase; letter-spacing:0.06em; color:var(--muted); margin-bottom:10px; }
      .fy-bars { display:flex; gap:10px; align-items:flex-end; }
      .fy-bar-col { flex:1; text-align:center; opacity:0.55; }
      .fy-bar-col--active { opacity:1; }
      .fy-bar-track { height:60px; display:flex; align-items:flex-end; justify-content:center; }
      .fy-bar { width:22px; background:var(--brand); border-radius:3px 3px 0 0; }
      .fy-bar--neg { background:var(--neg); }
      .fy-bar-label { font-family:'IBM Plex Mono',monospace; font-size:10px; color:var(--muted); margin-top:6px; }
      .fy-bar-value { font-family:'IBM Plex Mono',monospace; font-size:9.5px; color:var(--ink); margin-top:2px; }

      .chart-grid { display:grid; grid-template-columns:1fr 1fr; gap:20px; margin-top:8px; }
      @media (max-width:900px) { .chart-grid { grid-template-columns:1fr; } }
      .chart-card { border:1px solid var(--border); border-radius:12px; padding:18px 18px 8px; background:#fff; }
      .chart-card__title { font-size:13px; font-weight:600; margin-bottom:2px; }
      .chart-card__note { font-size:11px; color:var(--muted); margin-bottom:10px; }
      .chart-empty { font-size:12.5px; color:var(--muted); padding:40px 8px; text-align:center; }

      .footnote { padding:24px 40px 0; font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--muted); display:flex; gap:8px; align-items:flex-start; line-height:1.6; }

      /* ---- top-level section nav ---- */
      .dash-nav { display:flex; gap:4px; padding:0 40px; border-bottom:1px solid var(--border); background:var(--surface); overflow-x:auto; }
      .dash-nav__tab { font-family:'IBM Plex Mono',monospace; font-size:12.5px; white-space:nowrap; padding:14px 18px; border:none; background:transparent;
                       color:var(--muted); cursor:pointer; border-bottom:2px solid transparent; margin-bottom:-1px; transition:color .15s, border-color .15s; }
      .dash-nav__tab:hover { color:var(--ink); }
      .dash-nav__tab--active { color:var(--brand); border-bottom-color:var(--brand); font-weight:600; }
      @media (max-width:600px) { .dash-nav { padding:0 20px; } }

      /* ---- placeholder pages (Industry & Competitors, News & Updates) ---- */
      .placeholder-page { max-width:520px; margin:48px auto; text-align:center; padding:48px 24px; border:1px dashed var(--border); border-radius:16px; }
      .placeholder-page__eyebrow { font-family:'IBM Plex Mono',monospace; font-size:11px; letter-spacing:0.12em; text-transform:uppercase; color:var(--gold); margin-bottom:10px; }
      .placeholder-page__title { font-family:'Fraunces',serif; font-size:26px; font-weight:500; margin-bottom:10px; }
      .placeholder-page__note { font-size:13px; color:var(--muted); line-height:1.6; }

      /* ---- financial tables (Revenue & Profitability, P&L) ---- */
      .fin-section__head { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; margin-bottom:16px; }
      .fin-section__head .section__title { margin-bottom:0; }
      .fin-section__controls { display:flex; align-items:center; gap:10px; flex-wrap:wrap; }

      .export-btn { display:flex; align-items:center; gap:6px; font-family:'IBM Plex Mono',monospace; font-size:11.5px; font-weight:500; color:var(--brand);
                    border:1px solid var(--brand); border-radius:9px; padding:8px 12px; cursor:pointer; background:#fff; transition:background .15s, color .15s; white-space:nowrap; }
      .export-btn:hover { background:var(--brand); color:#fff; }

      .period-toggle { display:flex; gap:4px; background:var(--surface); padding:4px; border-radius:10px; border:1px solid var(--border); flex-shrink:0; }
      .period-toggle__btn { font-family:'IBM Plex Mono',monospace; font-size:12px; padding:7px 14px; border-radius:7px; border:none; background:transparent; color:var(--muted); cursor:pointer; transition:all .15s; white-space:nowrap; }
      .period-toggle__btn:hover { color:var(--ink); }
      .period-toggle__btn--active { background:var(--brand); color:#fff; }

      .fin-table-wrap { overflow-x:auto; border:1px solid var(--border); border-radius:12px; }
      .fin-table { width:100%; border-collapse:collapse; font-family:'IBM Plex Mono',monospace; font-size:12.5px; white-space:nowrap; }
      .fin-table thead th { text-align:right; font-weight:500; color:var(--muted); font-size:10.5px; text-transform:uppercase; letter-spacing:0.04em; padding:12px 16px; border-bottom:1px solid var(--border); background:var(--surface); }
      .fin-table tbody td { text-align:right; padding:10px 16px; border-bottom:1px solid #F0F1F3; color:var(--ink); }
      .fin-table tbody tr:last-child td { border-bottom:none; }
      .fin-table tbody tr:not(.fin-table__section-row):hover td { background:#FAFAF7; }
      .fin-table tbody tr:not(.fin-table__section-row) td:last-child,
      .fin-table thead th:last-child { background:#F6F1E4; }
      .fin-table thead th:last-child { font-weight:600; }
      .fin-table tbody tr:not(.fin-table__section-row):hover td:last-child { background:#F1EBD9; }

      .fin-table__label-col { position:sticky; left:0; background:#fff; font-family:'Inter',sans-serif; font-size:12.5px; font-weight:500;
                               text-align:left !important; z-index:1; border-right:1px solid var(--border); min-width:170px; }
      .fin-table thead th.fin-table__label-col { background:var(--surface); z-index:2; }
      .fin-table__partial { color:var(--gold); font-weight:500; }

      .fin-table__section-row td { background:var(--surface); font-family:'Inter',sans-serif; font-size:11px; font-weight:600; text-transform:uppercase;
                                    letter-spacing:0.05em; color:var(--brand); text-align:left; padding:9px 16px; border-bottom:1px solid var(--border); }
      .fin-table__subtotal-row td { font-weight:600; border-top:1px solid var(--border); }
      .fin-table__subtotal-row td.fin-table__label-col { background:#fff; }
      .fin-table .delta { font-size:11px; padding:2px 6px; justify-content:flex-end; }

      /* ---- Key Performance Indicators table: compact, emphasized business-line rows ---- */
      .fin-table-wrap--kpi { border-color:var(--gold); border-width:1.5px; }
      .fin-table--kpi tbody tr td.fin-table__label-col { font-weight:600; color:var(--brand); }
      .fin-table--kpi tbody tr td:not(.fin-table__label-col) { font-family:'Inter',sans-serif; font-weight:500; font-variant-numeric:tabular-nums; }
      .fin-table--kpi tbody tr:not(:last-child) td { border-bottom:1px solid var(--border); }
      .fin-table__na { color:var(--muted); font-weight:400 !important; }
      .fin-table__foot-note { display:flex; align-items:flex-start; gap:6px; margin-top:10px; font-size:11.5px; color:var(--muted); line-height:1.6; }

      /* ---- source links (used throughout Industry & Competitors / News) ---- */
      .src-link { display:inline-flex; align-items:center; gap:2px; font-family:'IBM Plex Mono',monospace; font-size:11px;
                  color:var(--brand); text-decoration:none; border-bottom:1px dotted var(--brand); white-space:nowrap; }
      .src-link:hover { color:var(--gold); border-bottom-color:var(--gold); }
      .src-link__arrow { font-size:10px; }

      /* ---- News & Updates ---- */
      .news-refresh { display:flex; flex-direction:column; align-items:flex-end; gap:2px; font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--muted); }
      .news-refresh__stamp { color:var(--brand); font-weight:600; }
      .news-refresh__note { color:var(--muted); }
      .news-filters { display:flex; flex-direction:column; gap:8px; margin-bottom:18px; }
      .news-filters__group { display:flex; gap:6px; flex-wrap:wrap; }
      .chip { font-family:'Inter',sans-serif; font-size:12px; padding:6px 12px; border-radius:20px; border:1px solid var(--border);
              background:#fff; color:var(--muted); cursor:pointer; transition:all .15s; white-space:nowrap; }
      .chip--mono { font-family:'IBM Plex Mono',monospace; font-size:11px; }
      .chip:hover { border-color:var(--brand); color:var(--brand); }
      .chip--active { background:var(--brand); border-color:var(--brand); color:#fff; }
      .news-expanded-note { font-size:12px; color:var(--gold); margin-bottom:14px; font-style:italic; }
      .news-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(300px, 1fr)); gap:14px; }
      .news-card { border:1px solid var(--border); border-radius:12px; padding:16px 18px; background:#fff; display:flex; flex-direction:column; gap:8px; }
      .news-card__top { display:flex; justify-content:space-between; align-items:center; gap:8px; }
      .news-card__category { font-family:'IBM Plex Mono',monospace; font-size:10px; text-transform:uppercase; letter-spacing:0.05em;
                              color:var(--brand); background:var(--surface); padding:3px 8px; border-radius:6px; }
      .news-card__date { font-family:'IBM Plex Mono',monospace; font-size:10.5px; color:var(--muted); }
      .news-card__title { font-family:'Fraunces',serif; font-size:16px; font-weight:500; line-height:1.35; }
      .news-card__summary { font-size:12.5px; color:var(--muted); line-height:1.55; }
      .news-card__sources { display:flex; gap:12px; flex-wrap:wrap; margin-top:auto; padding-top:6px; }

      /* ---- Industry & Competitors ---- */
      .snapshot-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(220px, 1fr)); gap:14px; }
      .snapshot-tile { border:1px solid var(--border); border-radius:12px; padding:16px; background:#fff; }
      .snapshot-tile__label { font-size:12px; color:var(--muted); font-weight:500; margin-bottom:8px; }
      .snapshot-tile__value { font-family:'Fraunces',serif; font-size:22px; font-weight:500; margin-bottom:8px; }
      .snapshot-tile__meta { display:flex; align-items:center; justify-content:space-between; gap:8px; }
      .snapshot-tile__period { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--gold); background:#F6F1E4; padding:2px 7px; border-radius:6px; }
      .snapshot-tile__note { font-size:11.5px; color:var(--muted); margin-top:8px; line-height:1.5; }

      .trend-card { display:flex; flex-direction:column; }
      .trend-card__desc { font-size:12.5px; color:var(--ink); line-height:1.6; margin:4px 0 10px; }
      .trend-card__why { font-size:12px; color:var(--muted); line-height:1.55; background:var(--surface); border-radius:8px; padding:10px 12px; margin-bottom:10px; }
      .trend-card__why-label { display:block; font-family:'IBM Plex Mono',monospace; font-size:10px; text-transform:uppercase; letter-spacing:0.05em; color:var(--brand); margin-bottom:4px; }
      .trend-card__foot { display:flex; align-items:center; justify-content:space-between; gap:8px; margin-top:auto; }
      .trend-card__date { font-family:'IBM Plex Mono',monospace; font-size:10.5px; color:var(--muted); }

      .cap-mark { font-weight:600; }
      .cap-mark--yes { color:var(--pos); }
      .cap-mark--no { color:var(--muted); }
      .cap-mark--na { color:var(--muted); font-size:10px; font-family:'IBM Plex Mono',monospace; }
      .fin-table--competitors td, .fin-table--competitors th { text-align:center; }
      .fin-table--competitors td.fin-table__label-col, .fin-table--competitors th.fin-table__label-col { text-align:left; }

      .competitor-notes { margin-top:12px; display:flex; flex-direction:column; gap:6px; }
      .competitor-notes__item { font-size:12px; color:var(--muted); line-height:1.6; }

      .analysis-list { display:flex; flex-direction:column; gap:10px; }
      .analysis-item { display:flex; align-items:flex-start; gap:10px; flex-wrap:wrap; font-size:13px; line-height:1.6;
                        border:1px solid var(--border); border-radius:10px; padding:12px 14px; background:#fff; }
      .analysis-item__badge { flex-shrink:0; font-family:'IBM Plex Mono',monospace; font-size:9.5px; text-transform:uppercase; letter-spacing:0.05em;
                               color:var(--gold); background:#F6F1E4; padding:3px 8px; border-radius:6px; }
    `}</style>
  );
}