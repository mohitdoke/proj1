/* ============================================================
   SHARED MIS CALCULATION ENGINE
   ============================================================
   Extracted verbatim from the original single-file dashboard
   (MISDashboard.jsx) with NO logic changes, so the frontend and
   the backend (Vercel serverless functions, seed scripts) share
   one implementation of: Excel parsing, company detection,
   per-company configuration/semantic mapping, FY/quarter
   aggregation, and every derived metric (Gross Margin, EBITDA,
   growth, etc). This is what guarantees the backend-driven
   dashboard produces byte-identical output to the original
   client-only version for the same Excel input — it is quite
   literally the same code, not a reimplementation.

   Runs in both environments:
   - Browser (bundled by Vite into the frontend) — used to render
     the dashboard from data already fetched from the backend API,
     and (admin-only) to preview a file before uploading it.
   - Node (Vercel serverless functions, scripts/seed.mjs) — used
     to parse an uploaded workbook, detect its company, and
     validate it before persisting normalized rows to Supabase.
   ============================================================ */
// The `xlsx` package is CJS. Node's native ESM interop only exposes the
// statically-analyzable named exports (missing a couple, e.g. readFile) and
// puts the full module.exports on `.default`; Vite's esbuild-based bundling
// for the browser build is more permissive but also sets `.default` to the
// same full object. Normalizing to `.default` when present works correctly
// in both the Node (API routes, scripts) and browser (Vite) environments.
import * as XLSX_NS from "xlsx";
const XLSX = XLSX_NS.default || XLSX_NS;

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
    // TEMPORARY fallback — this workbook's test data only covers 4 months
    // (Apr–Jul 2026) of "Total number of eSigns" / "Number of Stamps Ordered
    // per period" / "Number of Active Subscription Wallets", so most
    // FY/quarter periods have no real aggregate yet. eSign/stamp counts are
    // stored as raw absolute counts (Cr value × 1e7, same convention as
    // Riskcovry's policyCount/gwp fallback above) to match fmtNum's plain
    // Math.round — the real "Total number of eSigns" row is itself tens of
    // millions per year, not a small Cr-scale decimal. Subscription accounts
    // is a point-in-time headcount already in the same raw units the real
    // "Active Subscription Wallets" row uses (low hundreds), so it's stored
    // as-is. Growth rows are fractions (GrowthBadge multiplies by 100), same
    // as every other growthSlug fallback.
    fallbackKPIs: {
      esigns: { FY24: 4.15e7, FY25: 5.23e7, FY26: 6.32e7, Q4FY25: 1.37e7, Q4FY26: 1.87e7 },
      esignsGrowth: { FY24: 1.00, FY25: 0.26, FY26: 0.208, Q4FY25: 0.08, Q4FY26: 0.365 },
      stamps: { FY24: 0.44e7, FY25: 0.62e7, FY26: 0.78e7, Q4FY25: 0.14e7, Q4FY26: 0.21e7 },
      stampsGrowth: { FY24: 1.35, FY25: 0.39, FY26: 0.258, Q4FY25: -0.01, Q4FY26: 0.50 },
      subscriptionAccounts: { FY24: 471, FY25: 524, FY26: 507, Q4FY25: 524, Q4FY26: 507 },
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
    // TEMPORARY fallback — none of the five operational KPI rows above exist
    // in this standardized MIS today (see the comment above `kpi:`), so every
    // cell is N/A from real data alone. Values are fractions (fmtPct
    // multiplies by 100), same convention as every other company's
    // fallbackKPIs. withFallback() only ever substitutes these when the real
    // MIS-matched value for that specific period is null — a future upload
    // that adds real Embedded Finance/Device Connect/etc rows immediately
    // takes priority over these figures with no code change required.
    fallbackKPIs: {
      embeddedFinance: { FY24: 0.459, FY25: 0.454, FY26: 0.574, Q4FY25: 0.432, Q4FY26: 0.644 },
      deviceConnect: { FY24: 0.390, FY25: 0.322, FY26: 0.190, Q4FY25: 0.337, Q4FY26: 0.108 },
      bankConnect: { FY24: 0.121, FY25: 0.135, FY26: 0.118, Q4FY25: 0.135, Q4FY26: 0.111 },
      bureauConnect: { FY24: 0.001, FY25: 0.008, FY26: 0.003, Q4FY25: 0.008, Q4FY26: 0.005 },
      marketXSentinel: { FY24: 0.028, FY25: 0.081, FY26: 0.114, Q4FY25: 0.089, Q4FY26: 0.131 },
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

function scoreCompanyConfigs(kpiKeys) {
  let best = null, bestScore = 0;
  Object.values(COMPANY_CONFIGS).forEach(cfg => {
    const score = cfg.signals.reduce((s, re) => s + (kpiKeys.some(k => re.test(k)) ? 1 : 0), 0);
    if (score > bestScore) { bestScore = score; best = cfg; }
  });
  return { best, bestScore };
}

function detectCompanyConfig(kpiKeys) {
  const { best, bestScore } = scoreCompanyConfigs(kpiKeys);
  // Require at least 2 distinct signal hits before committing to a company-specific
  // template — a single coincidental match (e.g. a generic "Revenue" row) shouldn't
  // be enough to mis-identify an unrelated company's sheet.
  return bestScore >= 2 ? best : COMPANY_CONFIGS.easyrewardz;
}

// Strict variant for the backend upload pipeline: NEVER silently falls back to
// Easyrewardz. Returns null when the sheet can't be confidently identified, so
// the caller can reject the upload with a clear error instead of mis-labelling
// an unrecognized workbook as an existing company's data.
function detectCompanyConfigStrict(kpiKeys) {
  const { best, bestScore } = scoreCompanyConfigs(kpiKeys);
  return bestScore >= 2 ? best : null;
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

/* A small number of source workbooks have a genuine bug in their own header
   row: the year field is incremented one year too early for a run of
   Oct/Nov/Dec columns, then corrected again the following January (verified
   in Fundamento's file: columns literally read Jan..Sep-2025, Oct..Dec-2026,
   Jan..Mar-2026, Apr-2026.. — the Oct-Dec columns are mislabeled a year
   ahead of where they actually sit in the sequence). Rather than trust
   those broken header dates for FY/quarter grouping (which would scatter
   real data into a nonsensical fiscal year) — or, just as importantly, for
   ANY month-keyed storage of this data (duplicate/out-of-order (y, m) pairs
   would collide with each other) — a company config can opt into
   `forceConsecutiveMonths: true`: this reconstructs every column's date as
   a strict one-month step from the FIRST column (which is always correctly
   labeled in every file seen so far), so the sheet's actual chronological,
   consecutive-monthly structure is respected regardless of what a later
   column's header cell happens to say. Values themselves are never
   touched — only which month/FY/quarter a column's values are attributed
   to. Exported so lib/misProcessing.js can apply the exact same correction
   before persisting monthly metrics, keeping the stored period_date for
   each column consistent with what buildDataset() itself will group by.
   Returns a NEW array; never mutates the one passed in. */
function resolveMonths(months, companyConfig) {
  if (!companyConfig.forceConsecutiveMonths || !months.length) return months;
  const y0 = months[0].y, m0 = months[0].m;
  return months.map((mo, i) => {
    const total = y0 * 12 + (m0 - 1) + i;
    const y = Math.floor(total / 12), m = (total % 12) + 1;
    return { y, m, label: monthLabel(y, m), key: `${y}-${String(m).padStart(2, "0")}` };
  });
}

function buildDataset(parsed, companyInfo, companyConfig = COMPANY_CONFIGS.easyrewardz) {
  const { headcount } = parsed;
  // Never mutate the caller's parsed.kpis — buildDataset synthesizes
  // canonical alias keys (e.g. FASTSURANCE's "EBITDA") onto `kpis` below,
  // and callers that build a dataset for VALIDATION before separately
  // persisting `parsed` verbatim (see lib/misProcessing.js) must still see
  // the original, un-synthesized rows afterwards — otherwise a computed
  // alias would get stored as if it were a genuine sheet row.
  const kpis = { ...parsed.kpis };
  const months = resolveMonths(parsed.months, companyConfig);
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

/* News-category classification — shared so a headline is tagged the same
   way whether it came from a manager-curated "News Feed" sheet, the old
   client-side live-fetch fallback, or the new server-side Tavily research
   service (lib/research.js). Categories match NEWS_CATEGORY_ORDER in
   MISDashboard.jsx exactly; this is pure text classification, not a
   financial calculation, but it still needs to be identical in both
   environments, so it lives here alongside the rest of the shared engine. */
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

export {
  excelSerialToDate,
  monthLabel,
  parseWorkbook,
  parseCompanyInfo,
  parseNewsSheet,
  parseIndustrySheet,
  parseRefreshMeta,
  COMPANY_CONFIGS,
  detectCompanyConfig,
  detectCompanyConfigStrict,
  buildFYGroups,
  buildQuarterGroups,
  resolveMonths,
  buildDataset,
  fmtCr,
  fmtPct,
  fmtNum,
  fmtPctSigned,
  fmtCrPlain,
  fmtCrAlready,
  fmtRupee,
  fmtRupeeThousands,
  fallbackKPI,
  withFallback,
  periodGrowth,
  guessNewsCategory,
};
