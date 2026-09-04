/* ============================================================
   REPORT CONTENT BUILDER — turns a built dataset (the exact object
   buildDataset() hands the dashboard) into the finished, already-
   formatted content of the one-company PPTX summary.

   WHY THIS EXISTS
   ---------------
   The report used to re-derive its own revenue/margin/EBITDA rows
   from the raw monthly sheet by generic keyword matching, over the
   last two *months* — independently of the website's own
   COMPANY_CONFIGS-driven resolution. That could (and did) put
   different figures in the deck than the dashboard showed for the
   same company. This module removes that whole class of drift: it
   reads the SAME `ds` the browser renders, calls the SAME formatters
   (fmtCr / fmtPct / fmtPctSigned / fmtNum / ...), the SAME
   periodGrowth()/withFallback() helpers, and the SAME period/margin/
   trend anchors (getExecStats / bestMarginKey / describeEbitdaTrend,
   which now live in misEngine.js precisely so both sides share one
   copy). Every number in the deck is therefore the identical string
   the website prints in that cell.

   WHAT IT PRODUCES — the three sections asked for, per company:
     1. kpiSection        — Key Performance Indicators (the company's
                            own KPI table, per its layout)
     2. financialSection  — Key Financial Highlights (revenue, growth,
                            gross profit/margin, EBITDA/margin)
     3. insights          — Business Description + the page's
                            Performance Summary / Commentary bullets
                            and stat tiles, as flat text
   ...each across FOUR period columns: the last two financial years
   and the last two quarters, taken as the literal last two columns of
   the website's own yearly and quarterly tables. A period still in
   progress is included exactly as the site shows it and flagged with
   an asterisk plus a footnote, rather than silently dropped.

   KEEPING IT IN STEP WITH THE PAGE
   --------------------------------
   Row selection and narrative wording below mirror the matching
   components in src/MISDashboard.jsx (RevenueProfitabilityTable, the
   per-company *KPITable, and the per-company *Commentary /
   ExecutiveSummary). The numbers cannot drift — they come from the
   shared engine — but the prose can, so a change to a commentary
   component on the page should be mirrored in the matching builder
   here. Each builder names the component it mirrors.
   ============================================================ */
import {
  fmtCr,
  fmtPct,
  fmtNum,
  fmtPctSigned,
  fmtCrPlain,
  fmtCrAlready,
  fmtRupee,
  fmtRupeeThousands,
  withFallback,
  periodGrowth,
  getExecStats,
  bestMarginKey,
  describeEbitdaTrend,
  KPI_SEMANTIC_RULES,
} from "./misEngine.js";

const PARTIAL_MARK = "*";

/* ------------------------------------------------------------
   PERIODS — the last two FY columns and the last two quarter
   columns, exactly as they appear (and in the order they appear)
   on the website's yearly/quarterly tables.

   Each entry keeps `list` + `index` (its position in the FULL
   ds.fyData / ds.qData array) because periodGrowth() looks its
   comparison period up inside the list it's handed — quarterly
   growth is "the same quarter one FY back", which simply isn't
   present in a 2-element slice. Passing the full list with the
   real index is what makes the deck's growth cells identical to
   the page's.
   ------------------------------------------------------------ */
function describePeriods(rows, list, quarterly) {
  return rows.map(row => {
    const partial = quarterly ? !row.complete : !!row.partial;
    return {
      row,
      list,
      index: list.indexOf(row),
      quarterly,
      key: row.key,
      label: `${row.label.replace(" (partial)", "")}${partial ? PARTIAL_MARK : ""}`,
      partial,
      // "Apr'25–Jun'25" — what the period actually covers, used by the
      // footnote that explains an asterisk instead of leaving it bare.
      sub: row.sub,
    };
  });
}

export function pickReportPeriods(ds) {
  const revKey = ds.revenueBaseKey;

  // Filter out future quarters that have no actual reported data (revenue 0 or null).
  const validQ = ds.hasRevenue
    ? ds.qData.filter(q => typeof q[revKey] === "number" && q[revKey] !== 0)
    : ds.qData.filter(q => ds.kpiKeys.some(k => typeof q[k] === "number" && q[k] !== 0));
  const lastQ = validQ.length ? validQ[validQ.length - 1] : ds.qData[ds.qData.length - 1];
  const lastQIdx = ds.qData.indexOf(lastQ);
  const qSlice = ds.qData.slice(Math.max(0, lastQIdx - 1), lastQIdx + 1);

  // A complete FY must have !f.partial AND its Q4 must have real non-zero revenue!
  const isFYComplete = (f) => {
    if (f.partial) return false;
    const q4 = ds.qData.find(q => q.fyEnd === f.fyEnd && q.qNum === 4);
    return q4 && typeof q4[revKey] === "number" && q4[revKey] !== 0;
  };

  const completeFYs = ds.fyData.filter(f => isFYComplete(f) && typeof f[revKey] === "number" && f[revKey] !== 0);
  let fySlice;
  if (completeFYs.length >= 2) {
    fySlice = completeFYs.slice(-2);
  } else {
    // If fewer than 2 complete FYs exist, take the FYs that have real revenue data
    const withData = ds.fyData.filter(f => typeof f[revKey] === "number" && f[revKey] !== 0);
    fySlice = withData.slice(-2);
  }

  return [
    ...describePeriods(fySlice, ds.fyData, false),
    ...describePeriods(qSlice, ds.qData, true),
  ];
}

function partialFootnote(periods) {
  const partials = periods.filter(p => p.partial);
  if (!partials.length) return null;
  const parts = partials.map(p => `${p.label} covers ${p.sub}`);
  return `${PARTIAL_MARK} Period still in progress — ${parts.join("; ")}.`;
}

/* ------------------------------------------------------------
   TABLE ROW HELPERS — one cell value per period, already
   formatted, so the renderer never does arithmetic.
   ------------------------------------------------------------ */
function row(label, values) {
  return { label, values };
}

function valueRow(periods, label, pick) {
  return row(label, periods.map(p => pick(p)));
}

/* ------------------------------------------------------------
   SECTION 2 — KEY FINANCIAL HIGHLIGHTS
   Mirrors RevenueProfitabilityTable in src/MISDashboard.jsx
   (rendered there as "Revenue & Profitability" for some layouts
   and "Key Financial Highlights" for others; the deck uses the
   newsletter's own single heading for every company).
   ------------------------------------------------------------ */
function buildFinancialSection(ds, periods) {
  if (!ds.hasRevenue) return null;
  const revLabel = ds.revenueLabel || "Revenue";
  const specs = [
    { label: revLabel, key: ds.revenueBaseKey, type: "currency" },
    { label: `${revLabel} Growth`, key: ds.revenueBaseKey, type: "growth" },
    ds.hasGP && { label: "Gross Profit", key: "Gross Profit", type: "currency" },
    ds.hasGP && { label: "Gross Margin", key: "Gross Margin", type: "percent" },
    ds.hasEBITDA && { label: "EBITDA", key: "EBITDA", type: "currency" },
    ds.hasEBITDA && { label: "EBITDA Margin", key: "EBITDA Margin", type: "percent" },
  ].filter(Boolean);

  const rows = specs.map(spec => valueRow(periods, spec.label, p => {
    if (spec.type === "growth") {
      const g = periodGrowth(p.list, p.index, spec.key, p.quarterly);
      return g === null ? "N/A" : fmtPctSigned(g * 100);
    }
    const v = p.row[spec.key];
    return spec.type === "percent" ? fmtPct(v) : fmtCr(v);
  }));

  return {
    title: "Key Financial Highlights",
    cornerLabel: "Metric",
    rows,
    notes: [],
  };
}

/* ------------------------------------------------------------
   SECTION 1 — KEY PERFORMANCE INDICATORS
   One builder per KPI-table layout in src/MISDashboard.jsx. Each
   returns the same { rows, notes } shape, so the renderer doesn't
   care which company it's drawing.
   ------------------------------------------------------------ */

/* Mirrors KeyPerformanceIndicatorsTable — revenue-mix shares of the
   revenue base, matched semantically. Used by the easyrewardz and
   "generic" layouts. */
function buildSemanticMixKpis(ds, periods) {
  const revenueLines = ds.kpiKeys.filter(k => k !== ds.revenueBaseKey && /revenue/i.test(k));
  const claimed = new Set();
  const specs = KPI_SEMANTIC_RULES.map(rule => {
    const matchedKeys = revenueLines.filter(k => !claimed.has(k) && rule.match.test(k));
    matchedKeys.forEach(k => claimed.add(k));
    return { label: rule.label, slug: rule.slug, matchedKeys };
  });

  const rows = specs.map(spec => valueRow(periods, spec.label, p => {
    let pct = null;
    if (spec.matchedKeys.length) {
      let sum = 0, has = false;
      spec.matchedKeys.forEach(k => { const v = p.row[k]; if (typeof v === "number") { sum += v; has = true; } });
      const total = p.row[ds.revenueBaseKey];
      pct = (has && typeof total === "number" && total !== 0) ? sum / total : null;
    }
    return fmtPct(withFallback(ds.companyConfig, spec.slug, p.key, pct));
  }));

  const unmatched = specs.filter(s => !s.matchedKeys.length && !ds.companyConfig.fallbackKPIs?.[s.slug]);
  const notes = unmatched.length
    ? [`${unmatched.map(r => r.label).join(" and ")} ${unmatched.length > 1 ? "have" : "has"} no revenue line in this workbook that reliably matches that concept — shown as N/A rather than guessed.`]
    : [];
  return { rows, notes };
}

/* Mirrors GrayQuestKPITable — one absolute total row plus share-of-total rows. */
function buildTotalAndSharesKpis(ds, periods) {
  const cfg = ds.companyConfig.kpi;
  const totalKey = ds.kpiKeys.find(k => cfg.totalMatch.test(k)) || null;
  const shareSpecs = cfg.shareRows.map(r => ({ ...r, matchedKey: ds.kpiKeys.find(k => r.match.test(k)) || null }));

  const rows = [
    valueRow(periods, cfg.totalLabel, p => (totalKey ? fmtNum(p.row[totalKey]) : "N/A")),
    ...shareSpecs.map(spec => valueRow(periods, spec.label, p => {
      let pct = null;
      if (spec.matchedKey && totalKey) {
        const raw = p.row[spec.matchedKey];
        const total = p.row[totalKey];
        pct = (typeof raw === "number" && typeof total === "number" && total !== 0) ? raw / total : null;
      }
      return fmtPct(withFallback(ds.companyConfig, spec.slug, p.key, pct));
    })),
  ];

  const notes = [];
  if (!totalKey) {
    notes.push(`No row matching "${cfg.totalLabel}" was found in this workbook — every KPI row above shows N/A rather than a guessed figure.`);
  } else {
    const unmatched = shareSpecs.filter(r => !r.matchedKey && !ds.companyConfig.fallbackKPIs?.[r.slug]);
    if (unmatched.length) {
      notes.push(`${unmatched.map(r => r.label).join(" and ")} ${unmatched.length > 1 ? "have" : "has"} no matching line in this workbook — shown as N/A rather than guessed.`);
    }
  }
  return { rows, notes };
}

/* Mirrors RiskcovryKPITable / LeegalityKPITable — a value row followed by
   its own YoY-growth row, plus (Leegality) one point-in-time stock row with
   no growth row. `valueFmt` and `growthLabel` are the only differences
   between the two components. */
function buildValueAndGrowthKpis(ds, periods, { valueFmt, growthLabel }) {
  const cfg = ds.companyConfig.kpi;
  const specs = (cfg.rows || []).map(r => ({ ...r, matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null }));
  const stockSpec = cfg.stockRow
    ? { ...cfg.stockRow, matchedKey: ds.kpiKeys.find(k => cfg.stockRow.matchers.some(re => re.test(k))) || null }
    : null;

  const rows = [];
  specs.forEach(spec => {
    rows.push(valueRow(periods, spec.label, p =>
      valueFmt(withFallback(ds.companyConfig, spec.slug, p.key, spec.matchedKey ? p.row[spec.matchedKey] : null), spec)));
    rows.push(valueRow(periods, growthLabel, p => {
      const g = spec.matchedKey ? periodGrowth(p.list, p.index, spec.matchedKey, p.quarterly) : null;
      const val = withFallback(ds.companyConfig, spec.growthSlug, p.key, g);
      return typeof val === "number" ? fmtPctSigned(val * 100) : "N/A";
    }));
  });
  if (stockSpec) {
    rows.push(valueRow(periods, stockSpec.label, p =>
      fmtNum(withFallback(ds.companyConfig, stockSpec.slug, p.key, stockSpec.matchedKey ? p.row[stockSpec.matchedKey] : null))));
  }

  const unmatched = [...specs, ...(stockSpec ? [stockSpec] : [])]
    .filter(r => !r.matchedKey && !ds.companyConfig.fallbackKPIs?.[r.slug]);
  const notes = unmatched.length
    ? [`No row matching ${unmatched.map(r => `"${r.label}"`).join(" or ")} was found in this workbook — shown as N/A rather than a guessed figure.`]
    : [];
  return { rows, notes };
}

/* Mirrors MultiplKPITable / FastsuranceKPITable / ApexFutureLabsKPITable /
   FinboxKPITable / FundamentoKPITable — flat value rows, each carrying its
   own formatter in the company config. */
function buildFlatKpis(ds, periods) {
  const specs = (ds.companyConfig.kpi.rows || []).map(r => ({ ...r, matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null }));
  const rows = specs.map(spec => valueRow(periods, spec.label, p =>
    spec.fmt(withFallback(ds.companyConfig, spec.slug, p.key, spec.matchedKey ? p.row[spec.matchedKey] : null))));
  const unmatched = specs.filter(r => !r.matchedKey && !ds.companyConfig.fallbackKPIs?.[r.slug]);
  const notes = unmatched.length
    ? [`No row matching ${unmatched.map(r => `"${r.label}"`).join(" or ")} was found in this workbook — shown as N/A rather than a guessed figure.`]
    : [];
  return { rows, notes };
}

/* Which KPI-table component the page would render for this layout. A layout
   with a company-specific table but no `kpi` config renders nothing on the
   page (each *KPITable bails on a missing config), so the deck drops the
   section too rather than silently falling back to a different table's rows
   for that company. */
function buildKpiSection(ds, periods) {
  const layout = ds.companyConfig.layout;
  const cfg = ds.companyConfig.kpi;
  let built = null;

  if (layout === "grayquest") {
    built = cfg ? buildTotalAndSharesKpis(ds, periods) : null;
  } else if (layout === "riskcovry") {
    built = cfg?.rows
      ? buildValueAndGrowthKpis(ds, periods, {
          valueFmt: (v, spec) => fmtCrPlain(v, spec.decimals),
          growthLabel: "% Growth YoY",
        })
      : null;
  } else if (layout === "leegality") {
    built = cfg?.rows ? buildValueAndGrowthKpis(ds, periods, { valueFmt: fmtNum, growthLabel: "Growth YoY" }) : null;
  } else if (cfg?.rows) {
    // multipl / fastsurance / apexFutureLabs / finbox / fundamento
    built = buildFlatKpis(ds, periods);
  } else if (ds.hasRevenue) {
    // easyrewardz and every "generic" layout
    built = buildSemanticMixKpis(ds, periods);
  }

  if (!built || !built.rows.length) return null;
  return { title: "Key Performance Indicators", cornerLabel: "KPI", ...built };
}

/* ------------------------------------------------------------
   SECTION 3 — BUSINESS INSIGHTS + PERFORMANCE SUMMARY
   One builder per commentary component in src/MISDashboard.jsx.
   Bullets and stat tiles are the same sentences the page renders,
   flattened to plain text (the page bolds figures inline; a table-
   free PPTX text box can't, so the wording is preserved and the
   emphasis dropped).
   ------------------------------------------------------------ */

/* The shared numbers every commentary opens with. */
function coreStats(ds) {
  const stats = getExecStats(ds);
  const { latestQ, prevYearQ, latestFY, prevFY } = stats;
  const revKey = ds.revenueBaseKey;

  const revYoY = latestQ && prevYearQ && typeof latestQ[revKey] === "number" && typeof prevYearQ[revKey] === "number" && prevYearQ[revKey]
    ? ((latestQ[revKey] - prevYearQ[revKey]) / Math.abs(prevYearQ[revKey])) * 100 : null;

  const ebitdaCurr = latestFY && typeof latestFY["EBITDA"] === "number" ? latestFY["EBITDA"] : null;
  const ebitdaPrev = prevFY && typeof prevFY["EBITDA"] === "number" ? prevFY["EBITDA"] : null;

  return { ...stats, revKey, revYoY, ebitdaCurr, ebitdaPrev, ebitdaTrend: describeEbitdaTrend(ebitdaCurr, ebitdaPrev) };
}

/* The margin concept the page's third bullet / third tile uses. Riskcovry
   overrides the pick (gross margin first); everyone else takes
   bestMarginKey()'s order. */
function marginStats(ds, latestFY, pick = bestMarginKey(ds)) {
  const idx = latestFY ? ds.fyData.findIndex(f => f.key === latestFY.key) : -1;
  const prevFY = idx > 0 ? ds.fyData[idx - 1] : null;
  return {
    pick,
    curr: pick && latestFY ? latestFY[pick[0]] : null,
    prev: pick && prevFY ? prevFY[pick[0]] : null,
  };
}

/* "Q4FY26 closed with net revenue of ₹23.3 Cr, up 19.9% versus ₹19.5 Cr in
   Q4FY25." — the opening bullet shared by every layout. `revenueWord`
   differs between components ("net revenue" vs "revenue"); `suffix` lets
   MULTIPL add its volatility clause. */
function revenueBullet(ds, s, { revenueWord = "net revenue", suffix = "" } = {}) {
  if (!(s.latestQ && s.prevYearQ && s.revYoY !== null)) {
    return `${ds.revenueLabel || "Revenue"} data isn't complete enough yet to compute a like-for-like quarterly YoY comparison.`;
  }
  return `${s.latestQ.label} closed with ${(ds.revenueLabel || revenueWord).toLowerCase()} of ${fmtCr(s.latestQ[s.revKey])}, ` +
    `${s.revYoY >= 0 ? "up" : "down"} ${Math.abs(s.revYoY).toFixed(1)}% versus ${fmtCr(s.prevYearQ[s.revKey])} in ${s.prevYearQ.label}${suffix}.`;
}

/* "FY EBITDA loss narrowed from ₹-1.36 Cr in FY25 to ₹0.22 Cr in FY26, with
   ebitda margin at 0.3% versus 1.8% the prior FY." */
function ebitdaBullet(s, margin, { suffix = "" } = {}) {
  if (!(s.latestFY && s.prevFY && s.ebitdaTrend)) {
    return "Not enough complete fiscal years of EBITDA yet to describe a trend.";
  }
  let text = `FY EBITDA ${s.ebitdaTrend} from ${fmtCr(s.ebitdaPrev)} in ${s.prevFY.label} to ${fmtCr(s.ebitdaCurr)} in ${s.latestFY.label}`;
  if (margin?.pick && typeof margin.curr === "number") {
    text += `, with ${margin.pick[0].toLowerCase()} at ${fmtPct(margin.curr)}`;
    if (typeof margin.prev === "number") text += ` versus ${fmtPct(margin.prev)} the prior FY`;
  }
  return `${text}${suffix}.`;
}

/* The two stat tiles every layout opens with. */
function coreTiles(ds, s) {
  const revWord = (ds.revenueLabel || "revenue").toLowerCase();
  const deltaSub = s.revYoY === null ? "" : `, ${fmtPctSigned(s.revYoY)} YoY`;
  return [
    {
      label: "Latest quarter vs prior year",
      value: s.latestQ ? fmtCr(s.latestQ[s.revKey]) : "N/A",
      sub: s.latestQ ? `${s.latestQ.label} ${revWord}${deltaSub}` : "—",
    },
    {
      label: "FY EBITDA",
      value: s.latestFY ? fmtCr(s.latestFY["EBITDA"]) : "N/A",
      sub: s.latestFY
        ? `${s.latestFY.label}${s.ebitdaTrend ? `, ${s.ebitdaTrend} from ${fmtCr(s.ebitdaPrev)} prior FY` : ""}`
        : "—",
    },
  ];
}

function marginTile(ds, s, margin) {
  return {
    label: margin.pick ? margin.pick[0] : "FY Margin",
    value: margin.pick && typeof margin.curr === "number" ? fmtPct(margin.curr) : "N/A",
    sub: s.latestFY ? `${s.latestFY.label}${typeof margin.prev === "number" ? `, vs ${fmtPct(margin.prev)} prior FY` : ""}` : "—",
  };
}

/* Mirrors ExecutiveSummary — easyrewardz and every "generic" layout. */
function summaryDefault(ds) {
  const s = coreStats(ds);
  const margin = marginStats(ds, s.latestFY);
  const note = ds.companyInfo?.strategicNote;

  const ebitdaDelta = s.latestFY && s.prevFY && ds.hasEBITDA && typeof s.ebitdaCurr === "number" && typeof s.ebitdaPrev === "number"
    ? s.ebitdaCurr - s.ebitdaPrev : null;
  const direction = ebitdaDelta === null ? null : ebitdaDelta > 0 ? "improved" : ebitdaDelta < 0 ? "declined" : "held steady";

  let ebitda;
  if (s.latestFY && s.prevFY && ds.hasEBITDA && direction) {
    ebitda = `FY EBITDA ${direction} from ${fmtCr(s.prevFY["EBITDA"])} in ${s.prevFY.label} to ${fmtCr(s.latestFY["EBITDA"])} in ${s.latestFY.label}` +
      `${s.latestQ && typeof s.latestQ["EBITDA"] === "number" ? `, with ${s.latestQ.label} at ${fmtCr(s.latestQ["EBITDA"])}` : ""}.`;
  } else {
    ebitda = "Not enough complete fiscal years of EBITDA yet to describe a trend — add more months to unlock this.";
  }
  if (note) ebitda += ` ${note.value}${note.sub ? ` — ${note.sub}` : ""}.`;

  const marginBullet = margin.pick && s.latestFY && typeof margin.curr === "number"
    ? `${margin.pick[0]} for ${s.latestFY.label} stood at ${fmtPct(margin.curr)}${typeof margin.prev === "number" ? `, versus ${fmtPct(margin.prev)} the prior FY` : ""}.`
    : `Add ${ds.revenueLabel || "Total Revenue"} plus Gross Profit, EBITDA, or Net Profit to the sheet to unlock a margin-trend read-out here.`;

  const noteTile = note
    ? { label: note.label, value: note.value, sub: note.sub || "" }
    : marginTile(ds, s, margin);

  return {
    title: "Performance Summary",
    eyebrow: "Quarterly & FY momentum",
    bullets: [revenueBullet(ds, s), ebitda, marginBullet],
    tiles: [...coreTiles(ds, s), noteTile],
    notes: [],
  };
}

/* Mirrors GrayQuestCommentary. */
function summaryGrayQuest(ds) {
  const s = coreStats(ds);
  const margin = marginStats(ds, s.latestFY);
  const kpiCfg = ds.companyConfig.kpi;
  const totalKey = kpiCfg ? ds.kpiKeys.find(k => kpiCfg.totalMatch.test(k)) : null;

  const disbYoY = totalKey && s.latestFY && s.prevFY && typeof s.latestFY[totalKey] === "number" && typeof s.prevFY[totalKey] === "number" && s.prevFY[totalKey]
    ? ((s.latestFY[totalKey] - s.prevFY[totalKey]) / Math.abs(s.prevFY[totalKey])) * 100 : null;

  const disbursals = disbYoY !== null
    ? `Disbursals ${disbYoY >= 0 ? "increased" : "decreased"} ${Math.abs(disbYoY).toFixed(1)}% YoY to ${fmtNum(s.latestFY[totalKey])} loans in ${s.latestFY.label}, from ${fmtNum(s.prevFY[totalKey])} in ${s.prevFY.label}.`
    : "Not enough complete fiscal years of disbursal data yet to compute YoY disbursal growth.";

  const shareOf = (fy, key) => {
    if (!fy || !key || !totalKey) return null;
    const num = fy[key], den = fy[totalKey];
    return (typeof num === "number" && typeof den === "number" && den !== 0) ? (num / den) * 100 : null;
  };
  const shares = (kpiCfg?.shareRows || [])
    .map(r => ({ label: r.label, matchedKey: ds.kpiKeys.find(k => r.match.test(k)) || null }))
    .map(r => ({ label: r.label, curr: shareOf(s.latestFY, r.matchedKey), prev: shareOf(s.prevFY, r.matchedKey) }))
    .filter(x => x.curr !== null);

  const mix = shares.length
    ? shares.map(x => `${x.label} accounted for ${x.curr.toFixed(1)}% of total disbursals in ${s.latestFY.label}${typeof x.prev === "number" ? `, versus ${x.prev.toFixed(1)}% in ${s.prevFY.label}` : ""}`).join("; ") + "."
    : "Segment-level disbursal mix (Schools / Colleges / Edtech Platforms) isn't available yet for this period.";

  const rateNote = (label, re) => {
    const key = ds.kpiKeys.find(k => re.test(k)) || null;
    const curr = key && s.latestFY && typeof s.latestFY[key] === "number" ? s.latestFY[key] : null;
    const prev = key && s.prevFY && typeof s.prevFY[key] === "number" ? s.prevFY[key] : null;
    if (typeof curr !== "number") return null;
    return `${label} averaged ${fmtPct(curr)} in ${s.latestFY?.label}${typeof prev === "number" ? `, versus ${fmtPct(prev)} the prior FY` : ""}.`;
  };
  const notes = [rateNote("Take rate", /take\s*rate/i), rateNote("Cost of funds", /cost\s*of\s*funds?/i)].filter(Boolean);
  if (notes.length) notes.push("(period figures are a simple average of the monthly rate, not volume-weighted)");

  return {
    title: "Performance Commentary",
    eyebrow: "Growth & mix",
    bullets: [disbursals, revenueBullet(ds, s), mix],
    tiles: [...coreTiles(ds, s), marginTile(ds, s, margin)],
    notes,
  };
}

/* Mirrors RiskcovryCommentary. */
function summaryRiskcovry(ds) {
  const s = coreStats(ds);
  // Riskcovry's commentary prefers gross margin over the generic pick.
  const margin = marginStats(ds, s.latestFY, (ds.hasRevenue && ds.hasGP) ? ["Gross Margin", "Gross Profit"] : bestMarginKey(ds));

  // The mix lines sum to Gross Revenue in this workbook, not to the Net
  // Revenue base used elsewhere — shares are computed against Gross Revenue
  // specifically and labelled as such, exactly as on the page.
  const mixBaseKey = ds.kpiKeys.includes("Gross Revenue") ? "Gross Revenue" : null;
  const shareOf = (fy, key) => {
    if (!fy || !key || !mixBaseKey || typeof fy[mixBaseKey] !== "number" || fy[mixBaseKey] === 0) return null;
    const num = fy[key];
    return typeof num === "number" ? (num / fy[mixBaseKey]) * 100 : null;
  };
  const mixes = (ds.companyConfig.kpi?.mixRows || [])
    .map(r => ({ label: r.label, matchedKey: ds.kpiKeys.find(k => r.match.test(k)) || null }))
    .map(r => ({ label: r.label, curr: shareOf(s.latestFY, r.matchedKey), prev: shareOf(s.prevFY, r.matchedKey) }))
    .filter(m => m.curr !== null);

  const mix = mixes.length
    ? mixes.map(m => `${m.label} contributed ${m.curr.toFixed(1)}% of gross revenue in ${s.latestFY.label}${typeof m.prev === "number" ? `, versus ${m.prev.toFixed(1)}% in ${s.prevFY.label}` : ""}`).join("; ") + "."
    : "Revenue-mix data (Platform Subscription / Product-Commission / Setup Fees, as a share of Gross Revenue) isn't available yet for this period.";

  const bullets = [revenueBullet(ds, s, { revenueWord: "revenue" }), ebitdaBullet(s, margin), mix];

  const opexKey = ds.kpiKeys.includes("Indirect Expenses") ? "Indirect Expenses" : null;
  const opexCurr = opexKey && s.latestFY && typeof s.latestFY[opexKey] === "number" ? s.latestFY[opexKey] : null;
  const opexPrev = opexKey && s.prevFY && typeof s.prevFY[opexKey] === "number" ? s.prevFY[opexKey] : null;
  const opexGrowth = (typeof opexCurr === "number" && typeof opexPrev === "number" && opexPrev !== 0)
    ? ((opexCurr - opexPrev) / Math.abs(opexPrev)) * 100 : null;
  if (opexGrowth !== null) {
    bullets.push(`Total operating expenses ${opexGrowth >= 0 ? "increased" : "decreased"} ${Math.abs(opexGrowth).toFixed(1)}% YoY to ${fmtCr(opexCurr)} in ${s.latestFY.label}, from ${fmtCr(opexPrev)} in ${s.prevFY.label}.`);
  }

  const revPerEmpCurr = s.latestFY && typeof s.latestFY["Rev per Employee"] === "number" ? s.latestFY["Rev per Employee"] : null;
  const revPerEmpPrev = s.prevFY && typeof s.prevFY["Rev per Employee"] === "number" ? s.prevFY["Rev per Employee"] : null;
  const headcountCurr = s.latestFY && typeof s.latestFY["Headcount"] === "number" ? s.latestFY["Headcount"] : null;
  const notes = typeof revPerEmpCurr === "number"
    ? [`Revenue per employee was ${fmtCr(revPerEmpCurr)} in ${s.latestFY?.label} (avg headcount ${fmtNum(headcountCurr)})${typeof revPerEmpPrev === "number" ? `, versus ${fmtCr(revPerEmpPrev)} the prior FY` : ""}.`]
    : [];

  return {
    title: "Performance Summary",
    eyebrow: "Growth & mix",
    bullets,
    tiles: [...coreTiles(ds, s), marginTile(ds, s, margin)],
    notes,
  };
}

/* Mirrors MultiplCommentary. */
function summaryMultipl(ds) {
  const s = coreStats(ds);
  const margin = marginStats(ds, s.latestFY);

  const pickFY = (re) => {
    const key = ds.kpiKeys.find(k => re.test(k)) || null;
    const curr = key && s.latestFY && typeof s.latestFY[key] === "number" ? s.latestFY[key] : null;
    const prev = key && s.prevFY && typeof s.prevFY[key] === "number" ? s.prevFY[key] : null;
    return { key, curr, prev };
  };
  const goals = pickFY(/goals\s*created/i);
  const avgGoal = pickFY(/average\s*goal\s*value/i);
  const aua = pickFY(/assets\s*under\s*advice/i);
  const signups = pickFY(/^signups$/i);
  const brand = pickFY(/brand\s*partners/i);

  const goalsGrowth = (typeof goals.curr === "number" && typeof goals.prev === "number" && goals.prev !== 0)
    ? ((goals.curr - goals.prev) / Math.abs(goals.prev)) * 100 : null;
  const auaGrowth = (typeof aua.curr === "number" && typeof aua.prev === "number" && aua.prev !== 0)
    ? ((aua.curr - aua.prev) / Math.abs(aua.prev)) * 100 : null;

  let goalsBullet;
  if (goals.key && s.latestFY && typeof goals.curr === "number") {
    goalsBullet = `Goals created ${goalsGrowth !== null ? `${goalsGrowth >= 0 ? "grew" : "declined"} ${Math.abs(goalsGrowth).toFixed(1)}% YoY to` : "reached"} ${fmtNum(goals.curr)} in ${s.latestFY.label}` +
      `${typeof goals.prev === "number" ? `, from ${fmtNum(goals.prev)} in ${s.prevFY.label}` : ""}` +
      `${typeof avgGoal.curr === "number" ? `, while average goal value stood at ${fmtRupee(avgGoal.curr)}${typeof avgGoal.prev === "number" ? ` (versus ${fmtRupee(avgGoal.prev)} the prior FY)` : ""}` : ""}.`;
  } else {
    goalsBullet = "Goal-creation data isn't complete enough yet to compute a YoY comparison.";
  }

  const bullets = [
    revenueBullet(ds, s, { revenueWord: "revenue", suffix: s.revYoY !== null && Math.abs(s.revYoY) > 40 ? " — quarterly revenue has shown meaningful volatility" : "" }),
    ebitdaBullet(s, margin, { suffix: " — consistent with continued investment ahead of monetisation" }),
    goalsBullet,
  ];

  const adoption = [];
  if (typeof aua.curr === "number") {
    adoption.push(`Assets Under Advice reached ${fmtCrAlready(aua.curr)} as of ${s.latestFY?.label}${auaGrowth !== null ? `, up ${auaGrowth.toFixed(1)}% versus ${fmtCrAlready(aua.prev)} the prior FY-end` : ""}.`);
  }
  if (typeof signups.curr === "number") adoption.push(`Cumulative signups stood at ${fmtNum(signups.curr)} as of ${s.latestFY?.label}.`);
  if (typeof brand.curr === "number") {
    adoption.push(`Active brand partners totalled ${fmtNum(brand.curr)} as of ${s.latestFY?.label}${typeof brand.prev === "number" ? `, versus ${fmtNum(brand.prev)} the prior FY-end` : ""}.`);
  }
  if (adoption.length) bullets.push(adoption.join(" "));

  return {
    title: "Performance Commentary",
    eyebrow: "Growth & adoption",
    bullets,
    tiles: [...coreTiles(ds, s), {
      label: "Goals Created",
      value: typeof goals.curr === "number" ? fmtNum(goals.curr) : "N/A",
      sub: s.latestFY ? `${s.latestFY.label}${goalsGrowth !== null ? `, ${fmtPctSigned(goalsGrowth)} YoY` : ""}` : "—",
    }],
    notes: [],
  };
}

/* Operational KPI rows resolved against the sheet — shared by the
   Fastsurance and FinBox commentaries. */
function matchedOpRows(ds) {
  return (ds.companyConfig.kpi?.rows || [])
    .map(r => ({ ...r, matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null }))
    .filter(r => r.matchedKey);
}

/* Mirrors FastsuranceCommentary. */
function summaryFastsurance(ds) {
  const s = coreStats(ds);
  const margin = marginStats(ds, s.latestFY);
  const opAvailable = matchedOpRows(ds);

  const ops = opAvailable.length
    ? opAvailable.map(r => `${r.label} stood at ${r.fmt(s.latestFY?.[r.matchedKey])} in ${s.latestFY?.label}`).join("; ") + "."
    : "This standardized MIS currently carries financial (P&L) data only — case-level operating metrics (registrations, resolved cases, resolution rate) aren't present in the uploaded sheet yet, so they're shown as N/A above rather than estimated.";

  const pctResolvedRow = (ds.companyConfig.kpi?.rows || [])
    .map(r => ({ ...r, matchedKey: ds.kpiKeys.find(k => r.matchers.some(re => re.test(k))) || null }))
    .find(r => r.slug === "pctResolved");
  const pctResolved = pctResolvedRow?.matchedKey && s.latestFY && typeof s.latestFY[pctResolvedRow.matchedKey] === "number"
    ? s.latestFY[pctResolvedRow.matchedKey] : null;

  return {
    title: "Performance Commentary",
    eyebrow: "Revenue & profitability",
    bullets: [revenueBullet(ds, s), ebitdaBullet(s, margin), ops],
    tiles: [...coreTiles(ds, s), {
      label: "% Resolved",
      value: typeof pctResolved === "number" ? fmtPct(pctResolved) : "N/A",
      sub: s.latestFY ? s.latestFY.label : "—",
    }],
    notes: [],
  };
}

/* Mirrors VitraCommentary (apexFutureLabs layout). */
function summaryVitra(ds) {
  const s = coreStats(ds);
  const margin = marginStats(ds, s.latestFY);
  const rows = ds.companyConfig.kpi?.rows || [];
  const keyFor = (slug) => {
    const spec = rows.find(r => r.slug === slug);
    return spec ? (ds.kpiKeys.find(k => spec.matchers.some(re => re.test(k))) || null) : null;
  };
  const custKey = keyFor("customers");
  const arpuKey = keyFor("arpu");
  const custCurr = custKey && s.latestFY && typeof s.latestFY[custKey] === "number" ? s.latestFY[custKey] : null;
  const custPrev = custKey && s.prevFY && typeof s.prevFY[custKey] === "number" ? s.prevFY[custKey] : null;
  const arpuCurr = arpuKey && s.latestFY && typeof s.latestFY[arpuKey] === "number" ? s.latestFY[arpuKey] : null;

  const customers = typeof custCurr === "number"
    ? `Customers stood at ${fmtNum(custCurr)} as of ${s.latestFY?.label}${typeof custPrev === "number" ? `, versus ${fmtNum(custPrev)} the prior FY-end` : ""}${typeof arpuCurr === "number" ? `, with ARPU (in thousands) of ${fmtRupeeThousands(arpuCurr)}` : ""}.`
    : "Customer count and ARPU aren't available yet for this workbook — add rows for these to the sheet (matched by name or common synonym) to unlock this.";

  return {
    title: "Performance Commentary",
    eyebrow: "Revenue & profitability",
    bullets: [revenueBullet(ds, s), ebitdaBullet(s, margin), customers],
    tiles: [...coreTiles(ds, s), {
      label: "Customers",
      value: typeof custCurr === "number" ? fmtNum(custCurr) : "N/A",
      sub: s.latestFY ? s.latestFY.label : "—",
    }],
    notes: [],
  };
}

/* Mirrors FinboxCommentary. */
function summaryFinbox(ds) {
  const s = coreStats(ds);
  const gmCurr = s.latestFY && typeof s.latestFY["Gross Margin"] === "number" ? s.latestFY["Gross Margin"] : null;
  const gmPrev = s.prevFY && typeof s.prevFY["Gross Margin"] === "number" ? s.prevFY["Gross Margin"] : null;

  let revenue;
  if (s.latestQ && s.prevYearQ && s.revYoY !== null) {
    revenue = `${s.latestQ.label} closed with ${(ds.revenueLabel || "net revenue").toLowerCase()} of ${fmtCr(s.latestQ[s.revKey])}, ` +
      `${s.revYoY >= 0 ? "up" : "down"} ${Math.abs(s.revYoY).toFixed(1)}% versus ${fmtCr(s.prevYearQ[s.revKey])} in ${s.prevYearQ.label}` +
      `${typeof gmCurr === "number" ? `, with gross margin at ${fmtPct(gmCurr)}${typeof gmPrev === "number" ? ` versus ${fmtPct(gmPrev)} the prior FY` : ""}` : ""}.`;
  } else {
    revenue = `${ds.revenueLabel || "Revenue"} data isn't complete enough yet to compute a like-for-like quarterly YoY comparison.`;
  }

  const ebitda = (s.latestFY && s.prevFY && s.ebitdaTrend)
    ? `FY EBITDA ${s.ebitdaTrend} from ${fmtCr(s.ebitdaPrev)} in ${s.prevFY.label} to ${fmtCr(s.ebitdaCurr)} in ${s.latestFY.label} (merging the sheet's two MIS vintages, latest/revised source preferred for overlapping months, per its own Read Me).`
    : "Not enough complete fiscal years of EBITDA yet to describe a trend.";

  const recurringKey = ds.kpiKeys.find(k => /^-Recurring Revenue$/i.test(k)) || null;
  const recurringCurr = recurringKey && s.latestFY && typeof s.latestFY[recurringKey] === "number" ? s.latestFY[recurringKey] : null;
  const revCurr = s.latestFY && typeof s.latestFY[s.revKey] === "number" ? s.latestFY[s.revKey] : null;
  const recurringShare = (typeof recurringCurr === "number" && typeof revCurr === "number" && revCurr) ? recurringCurr / revCurr : null;

  const opAvailable = matchedOpRows(ds);
  const mix = (typeof recurringShare === "number"
    ? `Recurring revenue made up ${fmtPct(recurringShare)} of ${(ds.revenueLabel || "revenue").toLowerCase()} in ${s.latestFY?.label}, the balance one-time/project revenue. `
    : "Revenue-mix (recurring vs one-time) data isn't populated for the latest FY yet. ") +
    (opAvailable.length
      ? opAvailable.map(r => `${r.label} stood at ${r.fmt(s.latestFY?.[r.matchedKey])} in ${s.latestFY?.label}`).join("; ") + "."
      : "Embedded-finance product-mix metrics (Embedded Finance %, Device Connect %, Bank Connect %, Bureau Connect %, MarketX/Sentinel %) aren't present in the uploaded standardized MIS — shown as N/A above rather than estimated.");

  return {
    title: "Performance Commentary",
    eyebrow: "Revenue & profitability",
    bullets: [revenue, ebitda, mix],
    tiles: [...coreTiles(ds, s), {
      label: "Gross Margin",
      value: typeof gmCurr === "number" ? fmtPct(gmCurr) : "N/A",
      sub: s.latestFY ? s.latestFY.label : "—",
    }],
    notes: [],
  };
}

/* Mirrors FundamentoCommentary. */
function summaryFundamento(ds) {
  const s = coreStats(ds);
  const ebitdaMarginCurr = s.latestFY && typeof s.latestFY["EBITDA Margin"] === "number" ? s.latestFY["EBITDA Margin"] : null;
  const ebitdaMarginPrev = s.prevFY && typeof s.prevFY["EBITDA Margin"] === "number" ? s.prevFY["EBITDA Margin"] : null;

  const pulsesKey = ds.kpiKeys.find(k => /cum+ulative\s*pulses/i.test(k)) || null;
  const pulsesCurr = pulsesKey && s.latestFY && typeof s.latestFY[pulsesKey] === "number" ? s.latestFY[pulsesKey] : null;
  const pulsesPrev = pulsesKey && s.prevFY && typeof s.prevFY[pulsesKey] === "number" ? s.prevFY[pulsesKey] : null;
  const pulsesGrowth = (typeof pulsesCurr === "number" && typeof pulsesPrev === "number" && pulsesPrev)
    ? ((pulsesCurr - pulsesPrev) / Math.abs(pulsesPrev)) * 100 : null;

  const rppKey = ds.kpiKeys.find(k => /^revenue\s*per\s*pulse$/i.test(k)) || null;
  const rppCurr = rppKey && s.latestFY && typeof s.latestFY[rppKey] === "number" ? s.latestFY[rppKey] : null;
  const rppPrev = rppKey && s.prevFY && typeof s.prevFY[rppKey] === "number" ? s.prevFY[rppKey] : null;

  const costPerPulse = (fy, pulses) => {
    const cost = fy && typeof fy["Total Cost"] === "number" ? fy["Total Cost"] : null;
    return (typeof cost === "number" && typeof pulses === "number" && pulses) ? cost / pulses : null;
  };
  const cppCurr = costPerPulse(s.latestFY, pulsesCurr);
  const cppPrev = costPerPulse(s.prevFY, pulsesPrev);

  const pulses = typeof pulsesCurr === "number"
    ? `Cumulative Pulses ${pulsesGrowth !== null ? `${pulsesGrowth >= 0 ? "grew" : "declined"} ${Math.abs(pulsesGrowth).toFixed(1)}% YoY to` : "reached"} ${fmtNum(pulsesCurr)} in ${s.latestFY?.label}` +
      `${typeof pulsesPrev === "number" ? `, from ${fmtNum(pulsesPrev)} in ${s.prevFY?.label}` : ""}` +
      `${typeof rppCurr === "number" ? `, at a realised revenue of ${fmtRupee(rppCurr)} per pulse${typeof rppPrev === "number" ? ` (versus ${fmtRupee(rppPrev)} the prior FY)` : ""}` : ""}.`
    : "Pulse-volume data isn't complete enough yet to compute a YoY comparison.";

  let ebitda;
  if (s.latestFY && s.prevFY && s.ebitdaTrend) {
    ebitda = `FY EBITDA ${s.ebitdaTrend} from ${fmtCr(s.ebitdaPrev)} in ${s.prevFY.label} to ${fmtCr(s.ebitdaCurr)} in ${s.latestFY.label}` +
      `${typeof ebitdaMarginCurr === "number" ? `, EBITDA margin at ${fmtPct(ebitdaMarginCurr)}${typeof ebitdaMarginPrev === "number" ? ` versus ${fmtPct(ebitdaMarginPrev)} prior FY` : ""}` : ""}` +
      `${(typeof cppCurr === "number" && typeof cppPrev === "number")
        ? `. Cost per pulse${cppCurr < cppPrev ? " declined" : cppCurr > cppPrev ? " rose" : " held steady"} to ${fmtRupee(cppCurr)} from ${fmtRupee(cppPrev)} the prior FY — the clearest read on operating leverage this MIS supports (it carries Total Revenue, Total Cost and EBITDA only, with no separate Gross Profit line, so Gross Margin is intentionally not shown as a distinct metric)`
        : ""}.`;
  } else {
    ebitda = "Not enough complete fiscal years of EBITDA yet to describe a trend.";
  }

  return {
    title: "Performance Commentary",
    eyebrow: "Revenue & profitability",
    bullets: [revenueBullet(ds, s, { revenueWord: "revenue" }), pulses, ebitda],
    tiles: [...coreTiles(ds, s), {
      label: "Cumulative Pulses",
      value: typeof pulsesCurr === "number" ? fmtNum(pulsesCurr) : "N/A",
      sub: s.latestFY ? `${s.latestFY.label}${pulsesGrowth !== null ? `, ${fmtPctSigned(pulsesGrowth)} YoY` : ""}` : "—",
    }],
    notes: [],
  };
}

/* Mirrors LeegalityCommentary. */
function summaryLeegality(ds) {
  const s = coreStats(ds);
  const margin = marginStats(ds, s.latestFY);
  const kpiCfg = ds.companyConfig.kpi;

  const keyForSlug = (slug) => {
    const spec = (kpiCfg?.rows || []).find(r => r.slug === slug);
    return spec ? (ds.kpiKeys.find(k => spec.matchers.some(re => re.test(k))) || null) : null;
  };
  const esignKey = keyForSlug("esigns");
  const stampKey = keyForSlug("stamps");
  const subKey = kpiCfg?.stockRow ? ds.kpiKeys.find(k => kpiCfg.stockRow.matchers.some(re => re.test(k))) || null : null;

  const lastQIdx = ds.qData.length - 1;
  const esignGrowth = esignKey ? periodGrowth(ds.qData, lastQIdx, esignKey, true) : null;
  const stampGrowth = stampKey ? periodGrowth(ds.qData, lastQIdx, stampKey, true) : null;
  const esigns = esignKey && s.latestQ ? s.latestQ[esignKey] : null;
  const stamps = stampKey && s.latestQ ? s.latestQ[stampKey] : null;
  const subs = subKey && s.latestQ && typeof s.latestQ[subKey] === "number" ? s.latestQ[subKey] : null;

  let ops = "";
  if (typeof esigns === "number") ops += `${s.latestQ.label} logged ${fmtNum(esigns)} eSigns${esignGrowth !== null ? ` (${fmtPctSigned(esignGrowth * 100)} YoY)` : ""}. `;
  if (typeof stamps === "number") ops += `${fmtNum(stamps)} stamps were ordered${stampGrowth !== null ? ` (${fmtPctSigned(stampGrowth * 100)} YoY)` : ""}. `;
  if (typeof subs === "number") ops += `Active subscription accounts stood at ${fmtNum(subs)} as of ${s.latestQ.label}.`;
  if (!ops) ops = "eSign / stamp / subscription-account data isn't complete enough yet for this period.";

  return {
    title: "Performance Commentary",
    eyebrow: "Growth & profitability",
    bullets: [revenueBullet(ds, s), ebitdaBullet(s, margin), ops.trim()],
    tiles: [...coreTiles(ds, s), {
      label: "No. of e-signs",
      value: typeof esigns === "number" ? fmtNum(esigns) : "N/A",
      sub: s.latestQ ? s.latestQ.label : "—",
    }],
    notes: [],
  };
}

const SUMMARY_BUILDERS = {
  grayquest: summaryGrayQuest,
  riskcovry: summaryRiskcovry,
  multipl: summaryMultipl,
  fastsurance: summaryFastsurance,
  apexFutureLabs: summaryVitra,
  finbox: summaryFinbox,
  fundamento: summaryFundamento,
  leegality: summaryLeegality,
};

function buildSummary(ds) {
  const builder = SUMMARY_BUILDERS[ds.companyConfig.layout] || summaryDefault;
  return builder(ds);
}

/* Mirrors BusinessDescription — the static company context above the
   narrative on the page. */
function buildBusinessDescription(ds) {
  const info = ds.companyInfo || ds.companyConfig?.defaultDescription;
  if (!info) return { title: "Business Description", description: "", tags: [], scaleMetrics: [] };
  return {
    title: "Business Description",
    description: info.description || "",
    tags: info.tags || [],
    scaleMetrics: (info.scaleMetrics || []).map(sm => ({ label: sm.label, value: sm.value })),
  };
}

/* ------------------------------------------------------------
   ENTRY POINT
   ------------------------------------------------------------ */
export function buildReportContent(ds, { legalEntityName, brand } = {}) {
  const info = ds.companyInfo || ds.companyConfig?.defaultDescription;
  const resolvedBrand = brand || info?.companyName || ds.companyInfo?.companyName || legalEntityName || "Company";
  const resolvedLegal = legalEntityName || ds.companyInfo?.companyName || info?.companyName || resolvedBrand;
  const periods = pickReportPeriods(ds);
  const summary = buildSummary(ds);
  const footnote = partialFootnote(periods);

  return {
    legal_entity_name: resolvedLegal,
    brand: resolvedBrand,
    period_labels: periods.map(p => p.label),
    business: buildBusinessDescription(ds),
    kpi_section: buildKpiSection(ds, periods),
    financial_section: buildFinancialSection(ds, periods),
    summary,
    // Anything that qualifies a figure rather than being one: the partial-
    // period marker, plus any per-section caveat the page shows as a
    // footnote under its own table.
    footnotes: [
      footnote,
      ...(summary.notes || []),
    ].filter(Boolean),
  };
}
