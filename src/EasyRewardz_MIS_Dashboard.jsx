import React, { useState, useMemo, useCallback, useEffect } from "react";
import {
  LineChart, Line, BarChart, Bar, ComposedChart,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend
} from "recharts";
import { ChevronDown, TrendingUp, TrendingDown, Minus, Upload, FileSpreadsheet, Info, X, RefreshCw, AlertTriangle } from "lucide-react";
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

  const fyData = fyGroups.map(fy => {
    const row = { ...fy };
    kpiKeys.forEach(k => { row[k] = sumFor(k, fy.idxs); });
    row["Headcount"] = avgHC(fy.idxs);
    if (hasRevenue && hasGP) row["Gross Margin"] = row["Total Revenue"] ? row["Gross Profit"] / row["Total Revenue"] : null;
    if (hasRevenue && hasEBITDA) row["EBITDA Margin"] = row["Total Revenue"] ? row["EBITDA"] / row["Total Revenue"] : null;
    if (hasRevenue && hasNet) row["Net Margin"] = row["Total Revenue"] ? row["Net Profit"] / row["Total Revenue"] : null;
    if (hasRevenue && headcount) row["Rev per Employee"] = row["Headcount"] ? row["Total Revenue"] / row["Headcount"] : null;
    return row;
  });

  const cardConfigs = [];
  kpiKeys.forEach(k => cardConfigs.push({ key: k, label: k, fmt: fmtCr, good: "up", primary: k === "Total Revenue" || k === "EBITDA" || k === "Net Profit" }));
  if (hasRevenue && hasGP) cardConfigs.push({ key: "Gross Margin", label: "Gross Margin", fmt: fmtPct, good: "up", isMargin: true, marginOf: "Gross Profit" });
  if (hasRevenue && hasEBITDA) cardConfigs.push({ key: "EBITDA Margin", label: "EBITDA Margin", fmt: fmtPct, good: "up", isMargin: true, marginOf: "EBITDA" });
  if (hasRevenue && hasNet) cardConfigs.push({ key: "Net Margin", label: "Net Margin", fmt: fmtPct, good: "up", isMargin: true, marginOf: "Net Profit" });
  if (headcount) cardConfigs.push({ key: "Headcount", label: "Headcount (avg)", fmt: fmtNum, good: "neutral", isHeadcount: true });
  if (hasRevenue && headcount) cardConfigs.push({ key: "Rev per Employee", label: "Revenue per Employee", fmt: fmtCr, good: "up", isRevPerEmp: true });

  return { months, kpis, headcount, kpiKeys, fyGroups, fyData, cardConfigs, hasRevenue, hasGP, hasEBITDA, hasNet };
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
          Every card, chart, and FY grouping above is computed live from the uploaded sheet — nothing is
          hardcoded. FY periods are derived from the dates in row 1 (Apr–Mar), so adding a new quarter's
          columns and re-uploading is all a refresh needs.
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