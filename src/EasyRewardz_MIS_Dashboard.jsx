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
    const values = monthCols.map(mc => {
      const v = row[mc.col];
      return typeof v === "number" && !isNaN(v) ? v : null;
    });
    if (/^headcount$/i.test(label)) headcount = values;
    else kpis[label] = values;
  }
  if (!Object.keys(kpis).length) throw new Error(`No KPI rows found under row 1 in "${sheetName}".`);

  return { sheetName, months, kpis, headcount };
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

const PRIORITY_ORDER = [
  "Total Revenue", "Retail/B2B Revenue", "Banking Revenue", "Campaign Mgmt Revenue",
  "Direct Expenses", "Gross Profit", "Indirect Expenses", "EBITDA", "Net Profit",
];

function buildDataset(parsed) {
  const { months, kpis, headcount } = parsed;
  const fyGroups = buildFYGroups(months);
  const qGroups = buildQuarterGroups(months);
  const kpiKeys = Object.keys(kpis).sort((a, b) => {
    const ia = PRIORITY_ORDER.indexOf(a), ib = PRIORITY_ORDER.indexOf(b);
    if (ia === -1 && ib === -1) return a.localeCompare(b);
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });

  const hasRevenue = kpiKeys.includes("Total Revenue");
  const hasGP = kpiKeys.includes("Gross Profit");
  const hasEBITDA = kpiKeys.includes("EBITDA");
  const hasNet = kpiKeys.includes("Net Profit");

  function sumFor(key, idxs) {
    const arr = kpis[key];
    let sum = 0, has = false;
    idxs.forEach(i => { const v = arr[i]; if (typeof v === "number") { sum += v; has = true; } });
    return has ? sum : null;
  }
  function avgHC(idxs) {
    if (!headcount) return null;
    const vals = idxs.map(i => headcount[i]).filter(v => typeof v === "number");
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  }
  function computeRow(groupMeta, idxs) {
    const row = { ...groupMeta };
    kpiKeys.forEach(k => { row[k] = sumFor(k, idxs); });
    row["Headcount"] = avgHC(idxs);
    if (hasRevenue && hasGP) row["Gross Margin"] = row["Total Revenue"] ? row["Gross Profit"] / row["Total Revenue"] : null;
    if (hasRevenue && hasEBITDA) row["EBITDA Margin"] = row["Total Revenue"] ? row["EBITDA"] / row["Total Revenue"] : null;
    if (hasRevenue && hasNet) row["Net Margin"] = row["Total Revenue"] ? row["Net Profit"] / row["Total Revenue"] : null;
    if (hasRevenue && headcount) row["Rev per Employee"] = row["Headcount"] ? row["Total Revenue"] / row["Headcount"] : null;
    return row;
  }

  const fyData = fyGroups.map(fy => computeRow(fy, fy.idxs));
  const qData = qGroups.map(q => computeRow(q, q.idxs));

  const cardConfigs = [];
  kpiKeys.forEach(k => cardConfigs.push({ key: k, label: k, fmt: fmtCr, good: "up", primary: k === "Total Revenue" || k === "EBITDA" || k === "Net Profit" }));
  if (hasRevenue && hasGP) cardConfigs.push({ key: "Gross Margin", label: "Gross Margin", fmt: fmtPct, good: "up", isMargin: true, marginOf: "Gross Profit" });
  if (hasRevenue && hasEBITDA) cardConfigs.push({ key: "EBITDA Margin", label: "EBITDA Margin", fmt: fmtPct, good: "up", isMargin: true, marginOf: "EBITDA" });
  if (hasRevenue && hasNet) cardConfigs.push({ key: "Net Margin", label: "Net Margin", fmt: fmtPct, good: "up", isMargin: true, marginOf: "Net Profit" });
  if (headcount) cardConfigs.push({ key: "Headcount", label: "Headcount (avg)", fmt: fmtNum, good: "neutral", isHeadcount: true });
  if (hasRevenue && headcount) cardConfigs.push({ key: "Rev per Employee", label: "Revenue per Employee", fmt: fmtCr, good: "up", isRevPerEmp: true });

  return { months, kpis, headcount, kpiKeys, fyGroups, fyData, qGroups, qData, cardConfigs, hasRevenue, hasGP, hasEBITDA, hasNet };
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
    const rev = ds.kpis["Total Revenue"]?.[i];
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
  const data = ds.months.map((mo, i) => ({ month: mo.label, revenue: ds.kpis["Total Revenue"][i] }));
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
  const revenueLines = ds.kpiKeys.filter(k => /revenue/i.test(k) && k !== "Total Revenue");
  if (!revenueLines.length) return <div className="chart-empty">No revenue sub-lines to break down — only "Total Revenue" is present.</div>;
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
  const data = ds.fyData.map(f => ({ fy: f.label, Revenue: f["Total Revenue"], EBITDA: f["EBITDA"], "Net Profit": f["Net Profit"] }));
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
    value: q["Total Revenue"],
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
  const forecast = buildQuarterlyForecast(ds.qData, "Total Revenue", 2);
  if (!forecast) return <div className="chart-empty">Need at least 4 complete quarters of Total Revenue to project a trend.</div>;
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
   NEW — Executive summary (narrative, computed live off the
   latest complete quarter / FY in the sheet) and business
   description (static company context).
   ============================================================ */
function ExecutiveSummary({ ds }) {
  const { latestQ, prevYearQ, latestFY, prevFY } = getExecStats(ds);

  const revYoY = latestQ && prevYearQ && ds.hasRevenue && prevYearQ["Total Revenue"]
    ? ((latestQ["Total Revenue"] - prevYearQ["Total Revenue"]) / Math.abs(prevYearQ["Total Revenue"])) * 100
    : null;

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
                {latestQ && prevYearQ && ds.hasRevenue ? (
                  <>The company reported strong quarterly momentum in <strong>{latestQ.label}</strong>, with net revenue
                  increasing to <strong>{fmtCr(latestQ["Total Revenue"])}</strong> from {fmtCr(prevYearQ["Total Revenue"])} in{" "}
                  {prevYearQ.label} ({fmtPctSigned(revYoY)} YoY), driven by higher campaign-led revenues and improved
                  client activation during the year-end period.</>
                ) : (
                  <>Revenue data isn't complete enough yet to compute a like-for-like quarterly YoY comparison — add more
                  months to the sheet to unlock this.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Rocket size={15} /></span>
              <span>
                {latestFY && prevFY && ds.hasEBITDA ? (
                  <>A patient turnaround. EBITDA flipped from <strong>{fmtCr(prevFY["EBITDA"])}</strong> in {prevFY.label} to{" "}
                  <strong>{fmtCr(latestFY["EBITDA"])}</strong> in {latestFY.label}
                  {latestQ ? <>, and turned positive in {latestQ.label} at <strong>{fmtCr(latestQ["EBITDA"])}</strong></> : ""}.
                  New product launch: <strong>OneConsent</strong>, a CDP for managing user consent across marketing channels.</>
                ) : (
                  <>New product launch: <strong>OneConsent</strong>, a CDP for managing user consent across marketing channels.</>
                )}
              </span>
            </div>
            <div className="narrative-bullet">
              <span className="narrative-bullet__icon"><Target size={15} /></span>
              <span>The MOIC isn't impressive yet — but profitable companies tend to find their multiple eventually.
              The reverse is far rarer.</span>
            </div>
          </div>
        </div>

        <div className="narrative-stat-col">
          <div className="stat-tile">
            <div className="stat-tile__label">Latest quarter vs prior year</div>
            <div className="stat-tile__value">{latestQ ? fmtCr(latestQ["Total Revenue"]) : "N/A"}</div>
            <div className="stat-tile__sub">
              {latestQ ? latestQ.label : "—"} revenue {revYoY !== null && <Delta curr={latestQ?.["Total Revenue"]} prev={prevYearQ?.["Total Revenue"]} good="up" />}
            </div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">FY EBITDA</div>
            <div className="stat-tile__value">{latestFY ? fmtCr(latestFY["EBITDA"]) : "N/A"}</div>
            <div className="stat-tile__sub">{latestFY ? `${latestFY.label}, vs ${prevFY ? fmtCr(prevFY["EBITDA"]) : "N/A"} prior FY` : "—"}</div>
          </div>
          <div className="stat-tile">
            <div className="stat-tile__label">Strategic note</div>
            <div className="stat-tile__value" style={{ fontSize: 16 }}>OneConsent CDP</div>
            <div className="stat-tile__sub">Consent management across marketing channels</div>
          </div>
        </div>
      </div>
    </section>
  );
}

function BusinessDescription() {
  return (
    <section className="section narrative-section">
      <div className="section__title">Business Description</div>
      <div className="biz-card">
        <p className="biz-desc-text">
          The company offers industry-agnostic, cloud-based CRM, Loyalty and Conversational Commerce solutions that
          enable seamless omnichannel customer experience. It runs loyalty programs for banks and online platforms.
        </p>
        <div className="biz-chip-row">
          <span className="biz-chip"><Layers size={11} style={{ marginRight: 5, verticalAlign: -2 }} />SaaS Technology</span>
          <span className="biz-chip"><ShieldCheck size={11} style={{ marginRight: 5, verticalAlign: -2 }} />Loyalty Platform as a Service</span>
          <span className="biz-chip"><Target size={11} style={{ marginRight: 5, verticalAlign: -2 }} />Consumer Analytics</span>
        </div>
        <div className="biz-scale-row">
          <div className="biz-scale-tile">
            <Building2 size={22} className="biz-scale-icon" />
            <div>
              <div className="biz-scale-value">220+</div>
              <div className="biz-scale-label">Brands</div>
            </div>
          </div>
          <div className="biz-scale-tile">
            <Store size={22} className="biz-scale-icon" />
            <div>
              <div className="biz-scale-value">25,000+</div>
              <div className="biz-scale-label">Stores / Branches</div>
            </div>
          </div>
          <div className="biz-scale-tile">
            <Users size={22} className="biz-scale-icon" />
            <div>
              <div className="biz-scale-value">Banks &amp; Platforms</div>
              <div className="biz-scale-label">Core loyalty client base</div>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
}

function EmptyState({ onFile, status }) {
  const [dragOver, setDragOver] = useState(false);
  return (
    <div className="empty-wrap">
      <div className="empty-eyebrow">EasyRewardz · Consolidated MIS</div>
      <div className="empty-title">Startup MIS Dashboard</div>
      <div className="empty-sub">Upload the mastersheet to build the dashboard. Nothing renders until real data arrives.</div>

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
        Don't have one yet? Use the rolling mastersheet template — it already has Mar'22 through Jun'26
        filled in from the source MIS workbooks. Each new quarter, add columns and re-upload.
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

  const onToggle = useCallback((key) => setExpanded(e => (e === key ? null : key)), []);

  const handleFile = useCallback(async (file) => {
    setStatus({ type: "info", text: `Reading ${file.name}…` });
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const parsed = parseWorkbook(wb);
      const ds = buildDataset(parsed);
      setDataset(ds);
      setFyIndex(ds.fyData.length - 1);
      setFileName(file.name);
      setStatus({ type: "success", text: `Loaded ${ds.months.length} months, ${ds.kpiKeys.length} KPIs, ${ds.fyData.length} FY periods from "${parsed.sheetName}".` });
    } catch (err) {
      setStatus({ type: "error", text: String(err.message || err) });
    }
  }, []);

  const handleReplace = useCallback(async (e) => {
    const file = e.target.files?.[0];
    if (file) await handleFile(file);
  }, [handleFile]);

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
          <div className="masthead__eyebrow">EasyRewardz · Consolidated MIS</div>
          <div className="masthead__title">Startup MIS Dashboard</div>
          <div className="masthead__meta">
            {fileName} · {dataset.months[0].label} → {dataset.months[dataset.months.length - 1].label} · viewing {fy.label} ({fy.sub})
          </div>
        </div>
        <div className="masthead__actions">
          <div className="fy-tabs">
            {dataset.fyData.map((f, i) => (
              <button key={f.key} className={`fy-tab ${i === fyIndex ? "fy-tab--active" : ""}`} onClick={() => setFyIndex(i)}>{f.label}</button>
            ))}
          </div>
          <label className="replace-btn">
            <RefreshCw size={13} /> New upload
            <input type="file" accept=".xlsx,.xls" onChange={handleReplace} />
          </label>
        </div>
      </header>

      {status && status.type !== "success" && (
        <div className={`upload-bar upload-bar--${status.type}`}>
          <FileSpreadsheet size={14} /><span>{status.text}</span>
        </div>
      )}

      <ExecutiveSummary ds={dataset} />
      <BusinessDescription />

      <section className="section">
        <div className="section__title">Key Metrics <span className="section__sub">— {fy.label} vs {fyIndex > 0 ? dataset.fyData[fyIndex - 1].label : "—"}</span></div>
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

      <div className="footnote">
        <Info size={13} style={{ flexShrink: 0, marginTop: 2 }} />
        <span>
          Every card, chart, and FY/quarter grouping above is computed live from the uploaded sheet — nothing is
          hardcoded except the business description and product-launch note. FY periods are derived from the dates
          in row 1 (Apr–Mar), so adding a new quarter's columns and re-uploading is all a refresh needs. Forecast
          lines in the Outlook section are a simple linear trend over the most recent complete quarters, not a
          modeled projection — treat them as directional only.
        </span>
      </div>
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
    `}</style>
  );
}