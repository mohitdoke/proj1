import React, { useState, useMemo, useCallback, useEffect, useRef } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  LineChart, Line, BarChart, Bar, ComposedChart, Cell, ReferenceLine,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend
} from "recharts";
import {
  ChevronDown, TrendingUp, TrendingDown, Minus, Upload, FileSpreadsheet, Info, X, RefreshCw,
  AlertTriangle, Rocket, Building2, Target, Users, Store, Layers, ShieldCheck
} from "lucide-react";
import * as XLSX_NS from "xlsx";
const XLSX = XLSX_NS.default || XLSX_NS;

// Excel parsing, company detection/configuration, and every calculation
// (FY/quarter aggregation, growth, margins, ...) now live in one module
// shared with the backend (see src/lib/misEngine.js) so the API-driven
// dashboard below produces exactly the same numbers the original
// client-only version did for the same Excel input — same code, not a
// reimplementation. This file only imports it; nothing here computes.
import {
  excelSerialToDate,
  parseWorkbook,
  parseCompanyInfo,
  parseNewsSheet,
  parseIndustrySheet,
  parseRefreshMeta,
  COMPANY_CONFIGS,
  detectCompanyConfig,
  buildFYGroups,
  buildQuarterGroups,
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
} from "./lib/misEngine.js";

/* ============================================================
   Reveal — thin motion.section wrapper used in place of every plain
   <section> in this file (mechanically swapped in below). Fades +
   lifts a section into place the first time it scrolls into view;
   `once:true` means it never re-triggers on scroll-back, so tables/
   charts underneath never get yanked around mid-read. Framer Motion
   itself already no-ops big transforms under prefers-reduced-motion
   at the browser level via its `useReducedMotion`-aware defaults, and
   GlobalStyles' own reduced-motion query collapses any leftover CSS
   transition/animation durations to ~0 as a second safety net.
   ============================================================ */
function Reveal({ className = "", children, ...rest }) {
  return (
    <motion.section
      className={className}
      initial={{ opacity: 0, y: 22 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true, margin: "-60px" }}
      transition={{ duration: 0.5, ease: [0.2, 0.65, 0.3, 1] }}
      {...rest}
    >
      {children}
    </motion.section>
  );
}

/* Small requestAnimationFrame-driven ease-out counter: interpolates from the
   previous rendered value to the new target whenever `target` changes, so
   KPI card values animate rather than snapping — purely presentational,
   `fmt` (the same formatter already used everywhere else) still owns the
   actual displayed text at every frame, so this can never show a number
   that formatting logic didn't produce. */
function useAnimatedNumber(target, duration = 700) {
  const [display, setDisplay] = useState(target);
  const fromRef = useRef(target);
  const rafRef = useRef(null);

  useEffect(() => {
    if (typeof target !== "number" || Number.isNaN(target)) { setDisplay(target); fromRef.current = target; return; }
    const from = typeof fromRef.current === "number" && !Number.isNaN(fromRef.current) ? fromRef.current : target;
    if (from === target) { setDisplay(target); return; }
    if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) { setDisplay(target); fromRef.current = target; return; }
    const start = performance.now();
    cancelAnimationFrame(rafRef.current);
    function tick(now) {
      const t = Math.min(1, (now - start) / duration);
      const eased = 1 - Math.pow(1 - t, 3); // easeOutCubic
      setDisplay(from + (target - from) * eased);
      if (t < 1) rafRef.current = requestAnimationFrame(tick);
      else fromRef.current = target;
    }
    rafRef.current = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(rafRef.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [target]);

  return display;
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

function KpiCard({ cfg, ds, fyIndex, expanded, onToggle, index = 0 }) {
  const fy = ds.fyData[fyIndex];
  const curr = fy[cfg.key];
  const prev = fyIndex > 0 ? ds.fyData[fyIndex - 1][cfg.key] : null;
  const sparkData = ds.fyData.map(f => ({ v: f[cfg.key] }));
  const isOpen = expanded === cfg.key;
  const animatedCurr = useAnimatedNumber(typeof curr === "number" ? curr : null);

  return (
    <motion.button
      className={`kpi-card ${cfg.primary ? "kpi-card--primary" : ""} ${isOpen ? "kpi-card--active" : ""}`}
      onClick={() => onToggle(cfg.key)}
      initial={{ opacity: 0, y: 16 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true, margin: "-40px" }}
      transition={{ duration: 0.35, delay: Math.min(index, 8) * 0.05, ease: "easeOut" }}
      whileHover={{ y: -3 }}
      whileTap={{ scale: 0.98 }}
    >
      <div className="kpi-card__top">
        <span className="kpi-card__label">{cfg.label}</span>
        <ChevronDown size={16} className="kpi-card__chev" />
      </div>
      <div className="kpi-card__value">{cfg.fmt(typeof curr === "number" ? animatedCurr : curr)}</div>
      <div className="kpi-card__foot">
        <Delta curr={curr} prev={prev} good={cfg.good} />
        <span className="kpi-card__spark"><MiniSpark data={sparkData} dataKey="v" color="#2DD4BF" /></span>
      </div>
    </motion.button>
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
    <motion.div
      className="modal-backdrop"
      onClick={onClose}
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      transition={{ duration: 0.18 }}
    >
      <motion.div
        className="modal-panel"
        onClick={(e) => e.stopPropagation()}
        initial={{ opacity: 0, scale: 0.95, y: 10 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.96, y: 8 }}
        transition={{ duration: 0.22, ease: [0.2, 0.9, 0.3, 1] }}
      >
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
            <CartesianGrid strokeDasharray="3 3" stroke="#1E2433" vertical={false} />
            <XAxis dataKey="month" tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#2A3142" }} tickLine={false} />
            <YAxis tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false}
              tickFormatter={(v) => isPct ? `${(v * 100).toFixed(0)}%` : cfg.isHeadcount ? v : fmtCr(v)} width={isPct ? 34 : 56} />
            <Tooltip formatter={(v) => cfg.fmt(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #2A3142", borderRadius: 8, background: "#0D1220", color: "#EAF0FA", boxShadow: "0 12px 32px -12px rgba(0,0,0,0.6)" }} />
            <Line type="monotone" dataKey="value" stroke="#F5B759" strokeWidth={2} dot={{ r: 2.5, fill: "#F5B759" }} isAnimationActive={false} />
          </LineChart>
        </ResponsiveContainer>
      </motion.div>
    </motion.div>
  );
}

function RevenueTrendChart({ ds }) {
  if (!ds.hasRevenue) return null;
  const data = ds.months.map((mo, i) => ({ month: mo.label, revenue: ds.kpis[ds.revenueBaseKey][i] }));
  return (
    <ResponsiveContainer width="100%" height={230}>
      <LineChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#1E2433" vertical={false} />
        <XAxis dataKey="month" tick={{ fontSize: 9, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} interval={Math.max(0, Math.floor(data.length / 16))} axisLine={{ stroke: "#2A3142" }} tickLine={false} />
        <YAxis tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
        <Tooltip formatter={(v) => fmtCr(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #2A3142", borderRadius: 8, background: "#0D1220", color: "#EAF0FA", boxShadow: "0 12px 32px -12px rgba(0,0,0,0.6)" }} />
        <Line type="monotone" dataKey="revenue" stroke="#2DD4BF" strokeWidth={2} dot={false} isAnimationActive={false} />
      </LineChart>
    </ResponsiveContainer>
  );
}

function RevenueMixChart({ ds }) {
  const revenueLines = ds.pnlRevenueSubLines
    ? ds.kpiKeys.filter(k => ds.pnlRevenueSubLines.some(re => re.test(k)))
    : ds.kpiKeys.filter(k => /revenue/i.test(k) && k !== ds.revenueBaseKey);
  if (!revenueLines.length) return <div className="chart-empty">No revenue sub-lines to break down — only "{ds.revenueLabel}" is present.</div>;
  const colors = ["#2DD4BF", "#F5B759", "#38BDF8", "#A78BFA", "#FB7185"];
  const data = ds.fyData.map(f => {
    const row = { fy: f.label };
    revenueLines.forEach(k => { row[k] = f[k]; });
    return row;
  });
  return (
    <ResponsiveContainer width="100%" height={230}>
      <BarChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#1E2433" vertical={false} />
        <XAxis dataKey="fy" tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#2A3142" }} tickLine={false} />
        <YAxis tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
        <Tooltip formatter={(v) => fmtCr(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #2A3142", borderRadius: 8, background: "#0D1220", color: "#EAF0FA", boxShadow: "0 12px 32px -12px rgba(0,0,0,0.6)" }} />
        <Legend wrapperStyle={{ fontFamily: "Inter", fontSize: 11, color: "#96A0BE" }} />
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
        <CartesianGrid strokeDasharray="3 3" stroke="#1E2433" vertical={false} />
        <XAxis dataKey="fy" tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#2A3142" }} tickLine={false} />
        <YAxis tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
        <Tooltip formatter={(v) => fmtCr(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #2A3142", borderRadius: 8, background: "#0D1220", color: "#EAF0FA", boxShadow: "0 12px 32px -12px rgba(0,0,0,0.6)" }} />
        <Legend wrapperStyle={{ fontFamily: "Inter", fontSize: 11, color: "#96A0BE" }} />
        <Bar dataKey="Revenue" fill="#333B4F" radius={[3, 3, 0, 0]} />
        {ds.hasEBITDA && <Line type="monotone" dataKey="EBITDA" stroke="#F5B759" strokeWidth={2} dot={{ r: 3 }} />}
        {ds.hasNet && <Line type="monotone" dataKey="Net Profit" stroke="#FB7185" strokeWidth={2} dot={{ r: 3 }} />}
      </ComposedChart>
    </ResponsiveContainer>
  );
}

function MarginTrendChart({ ds }) {
  const lines = [];
  if (ds.hasRevenue && ds.hasGP) lines.push(["Gross Margin", "#2DD4BF"]);
  if (ds.hasRevenue && ds.hasEBITDA) lines.push(["EBITDA Margin", "#F5B759"]);
  if (ds.hasRevenue && ds.hasNet) lines.push(["Net Margin", "#FB7185"]);
  if (!lines.length) return <div className="chart-empty">Need Total Revenue plus at least one of Gross Profit / EBITDA / Net Profit to compute margins.</div>;
  const data = ds.fyData.map(f => {
    const row = { fy: f.label };
    lines.forEach(([k]) => { row[k] = f[k]; });
    return row;
  });
  return (
    <ResponsiveContainer width="100%" height={230}>
      <LineChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#1E2433" vertical={false} />
        <XAxis dataKey="fy" tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#2A3142" }} tickLine={false} />
        <YAxis tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v * 100).toFixed(0)}%`} width={46} />
        <Tooltip formatter={(v) => fmtPct(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #2A3142", borderRadius: 8, background: "#0D1220", color: "#EAF0FA", boxShadow: "0 12px 32px -12px rgba(0,0,0,0.6)" }} />
        <Legend wrapperStyle={{ fontFamily: "Inter", fontSize: 11, color: "#96A0BE" }} />
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
        <CartesianGrid strokeDasharray="3 3" stroke="#1E2433" vertical={false} />
        <XAxis dataKey="fy" tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#2A3142" }} tickLine={false} />
        <YAxis yAxisId="left" tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} width={36} />
        <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} width={46} tickFormatter={(v) => `₹${v.toFixed(0)}L`} />
        <Tooltip formatter={(v, n) => n === "Rev/Employee" ? `₹${v.toFixed(1)} L` : Math.round(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #2A3142", borderRadius: 8, background: "#0D1220", color: "#EAF0FA", boxShadow: "0 12px 32px -12px rgba(0,0,0,0.6)" }} />
        <Legend wrapperStyle={{ fontFamily: "Inter", fontSize: 11, color: "#96A0BE" }} />
        <Bar yAxisId="left" dataKey="Headcount" fill="#333B4F" radius={[3, 3, 0, 0]} />
        {ds.hasRevenue && <Line yAxisId="right" type="monotone" dataKey="Rev/Employee" stroke="#2DD4BF" strokeWidth={2} dot={{ r: 3 }} />}
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
        <CartesianGrid strokeDasharray="3 3" stroke="#1E2433" vertical={false} />
        <XAxis dataKey="period" tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#2A3142" }} tickLine={false} />
        <YAxis tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
        <Tooltip formatter={(v) => fmtCr(v)} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #2A3142", borderRadius: 8, background: "#0D1220", color: "#EAF0FA", boxShadow: "0 12px 32px -12px rgba(0,0,0,0.6)" }} />
        <Bar dataKey="value" radius={[3, 3, 0, 0]}>
          {data.map((d, i) => <Cell key={i} fill={d.isLatest ? "#F5B759" : "#2DD4BF"} />)}
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
          <CartesianGrid strokeDasharray="3 3" stroke="#1E2433" vertical={false} />
          <XAxis dataKey="period" tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#2A3142" }} tickLine={false} />
          <YAxis tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
          <Tooltip formatter={(v) => (v === null ? "N/A" : fmtCr(v))} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #2A3142", borderRadius: 8, background: "#0D1220", color: "#EAF0FA", boxShadow: "0 12px 32px -12px rgba(0,0,0,0.6)" }} />
          <ReferenceLine y={0} stroke="#FB7185" strokeDasharray="2 2" />
          <Bar dataKey="actual" radius={[3, 3, 0, 0]}>
            {chartData.map((d, i) => <Cell key={i} fill={d.actual == null ? "transparent" : d.actual >= 0 ? "#4ADE80" : "#FB7185"} />)}
          </Bar>
          {forecast && <Line type="monotone" dataKey="projected" stroke="#F5B759" strokeWidth={2} strokeDasharray="5 5" dot={{ r: 3 }} />}
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
          <CartesianGrid strokeDasharray="3 3" stroke="#1E2433" vertical={false} />
          <XAxis dataKey="period" tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={{ stroke: "#2A3142" }} tickLine={false} />
          <YAxis tick={{ fontSize: 10, fill: "#96A0BE", fontFamily: "IBM Plex Mono" }} axisLine={false} tickLine={false} tickFormatter={fmtCr} width={62} />
          <Tooltip formatter={(v) => (v === null ? "N/A" : fmtCr(v))} contentStyle={{ fontFamily: "IBM Plex Mono", fontSize: 12, border: "1px solid #2A3142", borderRadius: 8, background: "#0D1220", color: "#EAF0FA", boxShadow: "0 12px 32px -12px rgba(0,0,0,0.6)" }} />
          <Line type="monotone" dataKey="actual" stroke="#2DD4BF" strokeWidth={2} dot={{ r: 3 }} />
          <Line type="monotone" dataKey="projected" stroke="#F5B759" strokeWidth={2} strokeDasharray="5 5" dot={{ r: 3 }} />
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
    <Reveal className="section">
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
    </Reveal>
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
    <Reveal className="section">
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
    </Reveal>
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
    <Reveal className="section">
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
    </Reveal>
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
    <Reveal className="section">
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
    </Reveal>
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
    <Reveal className="section">
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
    </Reveal>
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
    <Reveal className="section">
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
    </Reveal>
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
    <Reveal className="section">
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
    </Reveal>
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
    <Reveal className="section">
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
    </Reveal>
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
    <Reveal className="section">
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
    </Reveal>
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
    <Reveal className="section">
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
    </Reveal>
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
    <Reveal className="section">
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
    </Reveal>
  );
}

function PlaceholderSection({ title, note, eyebrow = "Coming next" }) {
  return (
    <Reveal className="section">
      <div className="placeholder-page">
        <div className="placeholder-page__eyebrow">{eyebrow}</div>
        <div className="placeholder-page__title">{title}</div>
        <div className="placeholder-page__note">{note}</div>
      </div>
    </Reveal>
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
// NEWS_CATEGORY_KEYWORDS / guessNewsCategory now live in ./lib/misEngine.js
// (imported in the header above) so the exact same classifier runs both here
// and in the server-side Tavily research service.

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
      <Reveal className="section">
        <div className="fin-section__head">
          <div className="section__title">News &amp; Updates</div>
        </div>
        <LiveNewsFeed
          query={query}
          emptyHint='Add a real "Company Name" in the Company Info sheet to fetch live news for your company.'
        />
      </Reveal>
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
    <Reveal className="section">
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
    </Reveal>
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
      <Reveal className="section">
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
      </Reveal>
    );
  }

  const capabilityCols = ["CRM", "Loyalty", "CDP", "Marketing Automation", "Conversational Commerce"];
  const refreshedAt = ds.refreshMeta?.industryRefreshedAt;

  return (
    <>
      <Reveal className="section">
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
      </Reveal>

      {data.snapshot.length > 0 && (
        <Reveal className="section">
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
        </Reveal>
      )}

      {data.trends.length > 0 && (
        <Reveal className="section">
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
        </Reveal>
      )}

      {data.competitors.length > 0 && (
        <Reveal className="section">
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
        </Reveal>
      )}

      {data.analysis.length > 0 && (
        <Reveal className="section">
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
        </Reveal>
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
    <Reveal className="section narrative-section">
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
    </Reveal>
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
    <Reveal className="section narrative-section">
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
    </Reveal>
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
    <Reveal className="section narrative-section">
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
    </Reveal>
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
    <Reveal className="section narrative-section">
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
    </Reveal>
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
    <Reveal className="section narrative-section">
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
    </Reveal>
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
    <Reveal className="section narrative-section">
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
    </Reveal>
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
    <Reveal className="section narrative-section">
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
    </Reveal>
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
    <Reveal className="section narrative-section">
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
    </Reveal>
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
    <Reveal className="section narrative-section">
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
    </Reveal>
  );
}

const CHIP_ICONS = [Layers, ShieldCheck, Target, Info];
const SCALE_ICONS = [Building2, Store, Users];

function BusinessDescription({ companyInfo }) {
  if (!companyInfo) return null;
  const { description, tags, scaleMetrics } = companyInfo;
  if (!description && !tags.length && !scaleMetrics.length) return null;

  return (
    <Reveal className="section narrative-section">
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
    </Reveal>
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

/* ============================================================
   DashboardView — everything BELOW the "have a dataset" line in the
   original single-file App(). Unchanged JSX/behaviour, just taking an
   already-built `ds` (from buildDataset()) as a prop instead of managing
   its own Excel-upload state, so the SAME rendering code serves both:
   - the original local-upload App() below (kept for admin preview /
     standalone use), and
   - the new backend-driven AppRoot (src/AppRoot.jsx), which fetches
     `parsed`/`companyInfo`/`configKey` from the API, runs buildDataset()
     itself, and passes the result here.
   `headerActions` lets the caller swap the masthead's right-hand control
   (a raw-file "New upload" input here vs. nothing/a fund-company switcher
   in the backend-driven app) without touching this component's markup.
   ============================================================ */
export function DashboardView({ ds: dataset, fileName, headerActions, banner }) {
  const [fyIndex, setFyIndex] = useState(dataset.fyData.length - 1);
  const [expanded, setExpanded] = useState(null);
  const [section, setSection] = useState("performance");

  const onToggle = useCallback((key) => setExpanded(e => (e === key ? null : key)), []);

  // A different company (or a new upload for the same one) means a whole
  // new `ds` object — reset view-local state so e.g. a stale fyIndex from a
  // company with fewer FY periods can't index out of range.
  useEffect(() => {
    setFyIndex(dataset.fyData.length - 1);
    setExpanded(null);
    setSection("performance");
  }, [dataset]);

  useEffect(() => {
    document.title = dataset?.companyInfo?.companyName ? `${dataset.companyInfo.companyName} — MIS Dashboard` : "MIS Dashboard";
  }, [dataset]);

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
          {headerActions}
        </div>
      </header>

      <nav className="dash-nav">
        {[
          { key: "performance", label: "Performance" },
          { key: "industry", label: "Industry & Competitors" },
          { key: "news", label: "News & Updates" },
        ].map(tab => (
          <button
            key={tab.key}
            className={`dash-nav__tab ${section === tab.key ? "dash-nav__tab--active" : ""}`}
            onClick={() => setSection(tab.key)}
          >
            {tab.label}
            {section === tab.key && (
              <motion.span className="dash-nav__underline" layoutId="dashNavUnderline" transition={{ type: "spring", stiffness: 500, damping: 40 }} />
            )}
          </button>
        ))}
      </nav>

      {banner}

      <AnimatePresence mode="wait">
      {section === "performance" && (
        <motion.div key="performance" initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -8 }} transition={{ duration: 0.22, ease: "easeOut" }}>
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

          <Reveal className="section">
            <div className="fin-section__head">
              <div className="section__title">Key Metrics <span className="section__sub">— {fy.label} vs {fyIndex > 0 ? dataset.fyData[fyIndex - 1].label : "—"}</span></div>
              <div className="fy-tabs">
                {dataset.fyData.map((f, i) => (
                  <button key={f.key} className={`fy-tab ${i === fyIndex ? "fy-tab--active" : ""}`} onClick={() => setFyIndex(i)}>{f.label}</button>
                ))}
              </div>
            </div>
            <div className="kpi-grid">
              {dataset.cardConfigs.map((cfg, i) => (
                <KpiCard key={cfg.key} cfg={cfg} ds={dataset} fyIndex={fyIndex} expanded={expanded} onToggle={onToggle} index={i} />
              ))}
            </div>
          </Reveal>

          <AnimatePresence>
            {expanded && (
              <DrillDownModal
                cfg={dataset.cardConfigs.find(c => c.key === expanded)}
                ds={dataset}
                fyIndex={fyIndex}
                onClose={() => setExpanded(null)}
              />
            )}
          </AnimatePresence>

          <Reveal className="section">
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
          </Reveal>

          {dataset.companyConfig.showForecast && (
            <Reveal className="section">
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
            </Reveal>
          )}

          <Reveal className="section">
            <div className="section__title">Cash &amp; Operations</div>
            <div className="chart-grid">
              <div className="chart-card" style={{ gridColumn: "1 / -1" }}>
                <div className="chart-card__title">Headcount &amp; Productivity</div>
                <div className="chart-card__note">Average headcount (bars) vs revenue per employee (line), by FY</div>
                <HeadcountChart ds={dataset} />
              </div>
            </div>
          </Reveal>

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
        </motion.div>
      )}

      {section === "industry" && (
        <motion.div key="industry" initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -8 }} transition={{ duration: 0.22, ease: "easeOut" }}>
          <IndustryCompetitorsPage ds={dataset} />
        </motion.div>
      )}

      {section === "news" && (
        <motion.div key="news" initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -8 }} transition={{ duration: 0.22, ease: "easeOut" }}>
          <NewsUpdatesPage ds={dataset} />
        </motion.div>
      )}
      </AnimatePresence>
    </div>
  );
}

/* ============================================================
   App — original standalone, local-upload-driven entry point. Kept as-is
   (still useful for an admin's "preview before uploading" check, or for
   running this file completely standalone) — parses an Excel file
   client-side with the exact same shared engine, then hands the result to
   DashboardView above. Normal end users reach the dashboard through
   src/AppRoot.jsx instead, which never asks them to upload anything.
   ============================================================ */
export function App() {
  const [dataset, setDataset] = useState(null);
  const [status, setStatus] = useState(null);
  const [fileName, setFileName] = useState(null);

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

  if (!dataset) {
    return (
      <div className="dash">
        <GlobalStyles />
        <EmptyState onFile={handleFile} status={status} />
      </div>
    );
  }

  return (
    <DashboardView
      ds={dataset}
      fileName={fileName}
      headerActions={
        <label className="replace-btn">
          <RefreshCw size={13} /> New upload
          <input type="file" accept=".xlsx,.xls" onChange={handleReplace} />
        </label>
      }
      banner={
        status && status.type !== "success" ? (
          <div className={`upload-bar upload-bar--${status.type}`}>
            <FileSpreadsheet size={14} /><span>{status.text}</span>
          </div>
        ) : null
      }
    />
  );
}

export function GlobalStyles() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

      :root { color-scheme: dark; }

      html, body { background:#070A11; }
      body {
        margin:0;
        overflow-x:hidden;
      }
      /* Fixed, slowly-drifting aurora glow behind the whole app. Pure CSS —
         no canvas/JS cost — and switched off entirely under reduced motion. */
      body::before {
        content:"";
        position:fixed;
        inset:-20vh -20vw;
        z-index:-1;
        background:
          radial-gradient(45% 40% at 14% 10%, rgba(45,212,191,0.09), transparent 70%),
          radial-gradient(40% 36% at 90% 15%, rgba(245,183,89,0.06), transparent 70%),
          radial-gradient(50% 46% at 78% 95%, rgba(167,139,250,0.07), transparent 70%),
          radial-gradient(36% 32% at 6% 90%, rgba(56,189,248,0.05), transparent 70%);
        filter: blur(100px);
        animation: auroraDrift 26s ease-in-out infinite alternate;
        pointer-events:none;
      }
      @keyframes auroraDrift {
        0%   { transform: translate3d(0,0,0) scale(1); }
        50%  { transform: translate3d(-2%,1.5%,0) scale(1.06); }
        100% { transform: translate3d(2%,-1%,0) scale(1.02); }
      }
      @media (prefers-reduced-motion: reduce) {
        body::before { animation:none; }
        * { animation-duration:0.001ms !important; animation-iteration-count:1 !important; transition-duration:0.001ms !important; scroll-behavior:auto !important; }
      }

      ::selection { background:rgba(45,212,191,0.35); color:#fff; }
      *:focus-visible { outline:2px solid var(--brand); outline-offset:2px; border-radius:4px; }

      ::-webkit-scrollbar { width:10px; height:10px; }
      ::-webkit-scrollbar-track { background:transparent; }
      ::-webkit-scrollbar-thumb { background:rgba(255,255,255,0.14); border-radius:8px; border:2px solid transparent; background-clip:padding-box; }
      ::-webkit-scrollbar-thumb:hover { background:rgba(255,255,255,0.22); background-clip:padding-box; }

      .dash { --ink:#EAF0FA; --muted:#96A0BE; --border:rgba(255,255,255,0.10); --border-strong:rgba(255,255,255,0.18);
              --surface:rgba(255,255,255,0.05); --bg:#070A11;
              --brand:#2DD4BF; --brand-rgb:45,212,191; --gold:#F5B759; --gold-rgb:245,183,89;
              --pos:#4ADE80; --pos-rgb:74,222,128; --neg:#FB7185; --neg-rgb:251,113,133; --sage:#38BDF8;
              --glass: linear-gradient(180deg, rgba(255,255,255,0.055), rgba(255,255,255,0.02));
              --glass-shadow: 0 1px 0 rgba(255,255,255,0.07) inset, 0 20px 44px -22px rgba(0,0,0,0.6);
              font-family:'Inter',sans-serif; color:var(--ink); background:var(--bg);
              min-height:100vh; padding:0 0 64px 0; position:relative; }
      .dash * { box-sizing:border-box; }
      .dash ::-webkit-scrollbar { width:10px; height:10px; }

      .empty-wrap { max-width:640px; margin:0 auto; padding:96px 24px 40px; text-align:center; }
      .empty-eyebrow { font-family:'IBM Plex Mono',monospace; font-size:11px; letter-spacing:0.12em; text-transform:uppercase; color:var(--gold); margin-bottom:8px; }
      .empty-title { font-family:'Space Grotesk',sans-serif; font-size:38px; font-weight:600; letter-spacing:-0.02em; margin-bottom:10px;
                     background:linear-gradient(135deg,#fff 20%,var(--brand) 65%,var(--gold) 100%); -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent; }
      .empty-sub { font-size:14px; color:var(--muted); margin-bottom:36px; }
      .dropzone { display:flex; flex-direction:column; align-items:center; gap:8px; border:1.5px dashed var(--border-strong); border-radius:16px;
                  padding:48px 24px; cursor:pointer; color:var(--brand); transition:border-color .2s, background .2s, box-shadow .2s, transform .2s;
                  background:var(--glass); backdrop-filter:blur(16px); }
      .dropzone:hover, .dropzone--over { border-color:var(--brand); box-shadow:0 0 0 1px rgba(var(--brand-rgb),0.3), 0 12px 32px -12px rgba(var(--brand-rgb),0.35); transform:translateY(-2px); }
      .dropzone input { display:none; }
      .dropzone__title { font-size:14.5px; font-weight:600; color:var(--ink); margin-top:6px; }
      .dropzone__sub { font-family:'IBM Plex Mono',monospace; font-size:11.5px; color:var(--muted); }
      .empty-status { margin-top:18px; font-family:'IBM Plex Mono',monospace; font-size:12px; padding:10px 14px; border-radius:8px; background:var(--surface); border:1px solid var(--border); color:var(--muted); display:flex; align-items:center; gap:8px; justify-content:center; text-align:left; }
      .empty-status--error { color:var(--neg); background:rgba(var(--neg-rgb),0.12); border-color:rgba(var(--neg-rgb),0.3); }
      .empty-status--success { color:var(--pos); background:rgba(var(--pos-rgb),0.12); border-color:rgba(var(--pos-rgb),0.3); }
      .empty-hint { margin-top:28px; font-size:12px; color:var(--muted); line-height:1.6; }

      .masthead { border-bottom:1px solid var(--border); padding:28px 40px 22px; display:flex; justify-content:space-between; align-items:flex-end; flex-wrap:wrap; gap:16px;
                  background:linear-gradient(180deg, rgba(255,255,255,0.03), transparent); }
      .masthead__eyebrow { font-family:'IBM Plex Mono',monospace; font-size:11px; letter-spacing:0.12em; text-transform:uppercase; color:var(--gold); margin-bottom:6px; }
      .masthead__title { font-family:'Space Grotesk',sans-serif; font-size:30px; font-weight:600; letter-spacing:-0.01em; line-height:1.1; color:var(--ink); }
      .masthead__meta { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--muted); margin-top:6px; }
      .masthead__actions { display:flex; align-items:center; gap:10px; flex-wrap:wrap; }

      .fy-tabs { display:flex; gap:4px; background:var(--surface); padding:4px; border-radius:10px; border:1px solid var(--border); backdrop-filter:blur(10px); }
      .fy-tab { font-family:'IBM Plex Mono',monospace; font-size:12px; padding:8px 14px; border-radius:7px; border:none; background:transparent; color:var(--muted); cursor:pointer; transition:all .18s; white-space:nowrap; }
      .fy-tab:hover { color:var(--ink); }
      .fy-tab--active { background:var(--brand); color:#052420; font-weight:600; box-shadow:0 0 16px -2px rgba(var(--brand-rgb),0.6); }

      .replace-btn { display:flex; align-items:center; gap:6px; font-family:'IBM Plex Mono',monospace; font-size:11.5px; color:var(--muted);
                     border:1px solid var(--border); border-radius:9px; padding:9px 12px; cursor:pointer; transition:border-color .18s, color .18s, box-shadow .18s;
                     background:var(--surface); backdrop-filter:blur(10px); }
      .replace-btn:hover { border-color:var(--brand); color:var(--brand); box-shadow:0 0 0 1px rgba(var(--brand-rgb),0.25); }
      .replace-btn input { display:none; }

      .upload-bar { margin:18px 40px 0; padding:10px 16px; border-radius:10px; display:flex; align-items:center; gap:10px; font-size:12.5px; border:1px solid var(--border); }
      .upload-bar--info { background:var(--surface); color:var(--muted); }
      .upload-bar--error { background:rgba(var(--neg-rgb),0.12); color:var(--neg); border-color:rgba(var(--neg-rgb),0.3); }

      .section { padding:32px 40px 8px; }
      .section__title { font-family:'Space Grotesk',sans-serif; font-size:19px; font-weight:600; margin-bottom:16px; display:flex; align-items:baseline; gap:10px; color:var(--ink); }
      .section__sub { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--muted); font-weight:400; }

      .narrative-section { padding-top:28px; }
      .narrative-grid { display:grid; grid-template-columns:1.3fr 1fr; gap:20px; }
      @media (max-width:900px) { .narrative-grid { grid-template-columns:1fr; } }
      .narrative-card, .biz-card, .stat-tile, .kpi-card, .chart-card, .news-card, .snapshot-tile, .analysis-item,
      .modal-panel, .admin-panel, .biz-scale-tile {
        background:var(--glass); backdrop-filter:blur(18px) saturate(140%); -webkit-backdrop-filter:blur(18px) saturate(140%);
        border:1px solid var(--border); box-shadow:var(--glass-shadow);
      }
      .narrative-card { border-radius:14px; padding:22px 24px; border-color:rgba(var(--gold-rgb),0.28); }
      .narrative-card__eyebrow { font-family:'IBM Plex Mono',monospace; font-size:10.5px; text-transform:uppercase; letter-spacing:0.08em; color:var(--gold); margin-bottom:14px; }
      .narrative-bullets { display:flex; flex-direction:column; gap:16px; }
      .narrative-bullet { display:flex; gap:12px; align-items:flex-start; font-size:13.5px; line-height:1.65; color:var(--ink); }
      .narrative-bullet__icon { flex-shrink:0; width:28px; height:28px; border-radius:8px; display:flex; align-items:center; justify-content:center; background:rgba(var(--brand-rgb),0.14); color:var(--brand); margin-top:1px; }
      .narrative-extra-note { display:flex; align-items:flex-start; gap:6px; margin-top:16px; padding-top:16px; border-top:1px solid var(--border); font-size:12px; color:var(--muted); line-height:1.6; }
      .narrative-stat-col { display:flex; flex-direction:column; gap:14px; }
      .stat-tile { border-radius:12px; padding:16px; transition:transform .2s, box-shadow .2s; }
      .stat-tile:hover { transform:translateY(-2px); }
      .stat-tile__label { font-family:'IBM Plex Mono',monospace; font-size:10.5px; text-transform:uppercase; letter-spacing:0.06em; color:var(--muted); margin-bottom:6px; }
      .stat-tile__value { font-family:'Space Grotesk',sans-serif; font-size:22px; font-weight:600; color:var(--ink); }
      .stat-tile__sub { font-size:11px; color:var(--muted); margin-top:6px; display:flex; align-items:center; gap:6px; flex-wrap:wrap; }

      .biz-card { border-radius:14px; padding:22px 24px; }
      .biz-desc-text { font-size:13.5px; line-height:1.75; color:var(--ink); margin:0 0 18px; }
      .biz-chip-row { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:20px; }
      .biz-chip { display:inline-flex; align-items:center; font-family:'IBM Plex Mono',monospace; font-size:11px; padding:7px 12px; border-radius:20px; border:1px solid rgba(var(--brand-rgb),0.3); color:var(--brand); background:rgba(var(--brand-rgb),0.1); }
      .biz-scale-row { display:flex; gap:14px; flex-wrap:wrap; }
      .biz-scale-tile { flex:1; min-width:150px; border-radius:12px; padding:16px; border-color:rgba(var(--gold-rgb),0.28); display:flex; align-items:center; gap:12px; }
      .biz-scale-icon { color:var(--gold); flex-shrink:0; }
      .biz-scale-value { font-family:'Space Grotesk',sans-serif; font-size:19px; font-weight:600; line-height:1.2; color:var(--ink); }
      .biz-scale-label { font-family:'IBM Plex Mono',monospace; font-size:10px; color:var(--muted); text-transform:uppercase; letter-spacing:0.05em; margin-top:2px; }

      .forecast-note { display:flex; align-items:center; gap:6px; font-family:'IBM Plex Mono',monospace; font-size:10.5px; color:var(--muted); margin:6px 0 10px; flex-wrap:wrap; }
      .forecast-dot { width:8px; height:8px; border-radius:2px; display:inline-block; }
      .forecast-dot--actual { background:var(--brand); box-shadow:0 0 6px rgba(var(--brand-rgb),0.7); }
      .forecast-dot--proj { background:var(--gold); box-shadow:0 0 6px rgba(var(--gold-rgb),0.7); }

      .kpi-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; }
      @media (max-width:1100px) { .kpi-grid { grid-template-columns:repeat(2,1fr); } }
      @media (max-width:600px) { .kpi-grid { grid-template-columns:1fr; } }

      .kpi-card { all:unset; display:block; box-sizing:border-box; border-radius:12px;
                  padding:16px 16px 14px; cursor:pointer; transition:box-shadow .2s, border-color .2s, transform .2s; }
      .kpi-card:hover { border-color:rgba(var(--brand-rgb),0.4); box-shadow:0 0 0 1px rgba(var(--brand-rgb),0.18), 0 16px 32px -16px rgba(var(--brand-rgb),0.4); transform:translateY(-3px); }
      .kpi-card:active { transform:translateY(-1px) scale(0.99); }
      .kpi-card--primary { border-color:rgba(var(--gold-rgb),0.32); }
      .kpi-card--active { border-color:var(--brand); box-shadow:0 0 0 3px rgba(var(--brand-rgb),0.22), 0 16px 32px -16px rgba(var(--brand-rgb),0.5); }
      .kpi-card__top { display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; }
      .kpi-card__label { font-size:12px; color:var(--muted); font-weight:500; }
      .kpi-card__chev { color:var(--muted); transform:rotate(-90deg); transition:transform .2s; }
      .kpi-card--active .kpi-card__chev { transform:rotate(0deg); color:var(--brand); }
      .kpi-card__value { font-family:'Space Grotesk',sans-serif; font-size:26px; font-weight:600; letter-spacing:-0.01em; margin-bottom:10px; text-align:left; color:var(--ink); font-variant-numeric:tabular-nums; }
      .kpi-card__foot { display:flex; align-items:center; justify-content:space-between; gap:8px; }
      .kpi-card__spark { width:72px; height:36px; flex-shrink:0; opacity:0.9; }

      .delta { display:inline-flex; align-items:center; gap:3px; font-family:'IBM Plex Mono',monospace; font-size:11.5px; font-weight:500; padding:3px 7px; border-radius:6px; }
      .delta-pos { color:var(--pos); background:rgba(var(--pos-rgb),0.14); }
      .delta-neg { color:var(--neg); background:rgba(var(--neg-rgb),0.14); }
      .delta-flat { color:var(--muted); background:var(--surface); }

      .modal-backdrop { position:fixed; inset:0; background:rgba(3,5,10,0.65); backdrop-filter:blur(6px);
                         display:flex; align-items:center; justify-content:center; padding:24px; z-index:1000; }
      .modal-panel { border-radius:16px; width:100%; max-width:620px; max-height:88vh; overflow-y:auto;
                     padding:24px 28px 28px; box-shadow:var(--glass-shadow), 0 0 80px -20px rgba(var(--brand-rgb),0.25); }
      .modal-panel__header { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:18px; }
      .modal-panel__eyebrow { font-family:'IBM Plex Mono',monospace; font-size:10.5px; text-transform:uppercase; letter-spacing:0.08em; color:var(--gold); margin-bottom:4px; }
      .modal-panel__title { font-family:'Space Grotesk',sans-serif; font-size:22px; font-weight:600; color:var(--ink); }
      .modal-panel__close { all:unset; cursor:pointer; color:var(--muted); padding:6px; border-radius:8px; transition:background .15s, color .15s; }
      .modal-panel__close:hover { background:var(--surface); color:var(--ink); }
      .modal-panel__value-row { display:flex; align-items:center; gap:10px; margin-bottom:22px; flex-wrap:wrap; }
      .modal-panel__value { font-family:'Space Grotesk',sans-serif; font-size:34px; font-weight:600; color:var(--ink); font-variant-numeric:tabular-nums; }
      .modal-panel__vs { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--muted); }

      .drawer-subhead { font-family:'IBM Plex Mono',monospace; font-size:10.5px; text-transform:uppercase; letter-spacing:0.06em; color:var(--muted); margin-bottom:10px; }
      .fy-bars { display:flex; gap:10px; align-items:flex-end; }
      .fy-bar-col { flex:1; text-align:center; opacity:0.5; transition:opacity .2s; }
      .fy-bar-col--active { opacity:1; }
      .fy-bar-track { height:60px; display:flex; align-items:flex-end; justify-content:center; }
      .fy-bar { width:22px; background:linear-gradient(180deg, var(--brand), rgba(var(--brand-rgb),0.55)); border-radius:3px 3px 0 0; transition:height .3s cubic-bezier(.2,.8,.3,1); }
      .fy-bar--neg { background:linear-gradient(180deg, var(--neg), rgba(var(--neg-rgb),0.55)); }
      .fy-bar-label { font-family:'IBM Plex Mono',monospace; font-size:10px; color:var(--muted); margin-top:6px; }
      .fy-bar-value { font-family:'IBM Plex Mono',monospace; font-size:9.5px; color:var(--ink); margin-top:2px; }

      .chart-grid { display:grid; grid-template-columns:1fr 1fr; gap:20px; margin-top:8px; }
      @media (max-width:900px) { .chart-grid { grid-template-columns:1fr; } }
      .chart-card { border-radius:12px; padding:18px 18px 8px; transition:box-shadow .2s, transform .2s, border-color .2s; }
      .chart-card:hover { border-color:var(--border-strong); box-shadow:var(--glass-shadow), 0 0 0 1px rgba(255,255,255,0.06); }
      .chart-card__title { font-size:13px; font-weight:600; margin-bottom:2px; color:var(--ink); }
      .chart-card__note { font-size:11px; color:var(--muted); margin-bottom:10px; }
      .chart-empty { font-size:12.5px; color:var(--muted); padding:40px 8px; text-align:center; }

      .footnote { padding:24px 40px 0; font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--muted); display:flex; gap:8px; align-items:flex-start; line-height:1.6; }

      /* ---- top-level section nav ---- */
      .dash-nav { display:flex; gap:4px; padding:0 40px; border-bottom:1px solid var(--border); background:#0B0F19; overflow-x:auto; }
      .dash-nav__tab { font-family:'IBM Plex Mono',monospace; font-size:12.5px; white-space:nowrap; padding:14px 18px; border:none; background:transparent;
                       color:var(--muted); cursor:pointer; border-bottom:2px solid transparent; margin-bottom:-1px; transition:color .18s; position:relative; }
      .dash-nav__tab:hover { color:var(--ink); }
      .dash-nav__tab--active { color:var(--brand); font-weight:600; text-shadow:0 0 18px rgba(var(--brand-rgb),0.5); }
      .dash-nav__underline { position:absolute; left:10px; right:10px; bottom:-2px; height:2px; border-radius:2px;
                              background:var(--brand); box-shadow:0 0 10px 1px rgba(var(--brand-rgb),0.7); }
      @media (max-width:600px) { .dash-nav { padding:0 20px; } }

      /* ---- placeholder pages (Industry & Competitors, News & Updates) ---- */
      .placeholder-page { max-width:520px; margin:48px auto; text-align:center; padding:48px 24px; border:1px dashed var(--border-strong); border-radius:16px; background:var(--surface); backdrop-filter:blur(12px); }
      .placeholder-page__eyebrow { font-family:'IBM Plex Mono',monospace; font-size:11px; letter-spacing:0.12em; text-transform:uppercase; color:var(--gold); margin-bottom:10px; }
      .placeholder-page__title { font-family:'Space Grotesk',sans-serif; font-size:26px; font-weight:600; margin-bottom:10px; color:var(--ink); }
      .placeholder-page__note { font-size:13px; color:var(--muted); line-height:1.6; }

      /* ---- financial tables (Revenue & Profitability, P&L) ---- */
      .fin-section__head { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; margin-bottom:16px; }
      .fin-section__head .section__title { margin-bottom:0; }
      .fin-section__controls { display:flex; align-items:center; gap:10px; flex-wrap:wrap; }

      .export-btn { display:flex; align-items:center; gap:6px; font-family:'IBM Plex Mono',monospace; font-size:11.5px; font-weight:500; color:var(--brand);
                    border:1px solid rgba(var(--brand-rgb),0.4); border-radius:9px; padding:8px 12px; cursor:pointer; background:rgba(var(--brand-rgb),0.08); transition:background .18s, color .18s, box-shadow .18s; white-space:nowrap; }
      .export-btn:hover { background:var(--brand); color:#052420; box-shadow:0 0 20px -4px rgba(var(--brand-rgb),0.6); }

      .period-toggle { display:flex; gap:4px; background:var(--surface); padding:4px; border-radius:10px; border:1px solid var(--border); flex-shrink:0; backdrop-filter:blur(10px); }
      .period-toggle__btn { font-family:'IBM Plex Mono',monospace; font-size:12px; padding:7px 14px; border-radius:7px; border:none; background:transparent; color:var(--muted); cursor:pointer; transition:all .18s; white-space:nowrap; }
      .period-toggle__btn:hover { color:var(--ink); }
      .period-toggle__btn--active { background:var(--brand); color:#052420; font-weight:600; box-shadow:0 0 16px -2px rgba(var(--brand-rgb),0.6); }

      .fin-table-wrap { overflow-x:auto; border:1px solid var(--border); border-radius:12px; background:var(--glass); backdrop-filter:blur(18px); box-shadow:var(--glass-shadow); }
      .fin-table { width:100%; border-collapse:collapse; font-family:'IBM Plex Mono',monospace; font-size:12.5px; white-space:nowrap; }
      .fin-table thead th { text-align:right; font-weight:500; color:var(--muted); font-size:10.5px; text-transform:uppercase; letter-spacing:0.04em; padding:12px 16px; border-bottom:1px solid var(--border); background:rgba(255,255,255,0.03); }
      .fin-table tbody td { text-align:right; padding:10px 16px; border-bottom:1px solid rgba(255,255,255,0.05); color:var(--ink); }
      .fin-table tbody tr:last-child td { border-bottom:none; }
      .fin-table tbody tr:not(.fin-table__section-row):hover td { background:rgba(255,255,255,0.035); }
      .fin-table tbody tr:not(.fin-table__section-row) td:last-child,
      .fin-table thead th:last-child { background:rgba(var(--gold-rgb),0.1); }
      .fin-table thead th:last-child { font-weight:600; }
      .fin-table tbody tr:not(.fin-table__section-row):hover td:last-child { background:rgba(var(--gold-rgb),0.16); }

      .fin-table__label-col { position:sticky; left:0; background:#0B0F19; font-family:'Inter',sans-serif; font-size:12.5px; font-weight:500;
                               text-align:left !important; z-index:1; border-right:1px solid var(--border); min-width:170px; }
      .fin-table thead th.fin-table__label-col { background:#0D1220; z-index:2; }
      .fin-table__partial { color:var(--gold); font-weight:500; }

      .fin-table__section-row td { background:rgba(255,255,255,0.03); font-family:'Inter',sans-serif; font-size:11px; font-weight:600; text-transform:uppercase;
                                    letter-spacing:0.05em; color:var(--brand); text-align:left; padding:9px 16px; border-bottom:1px solid var(--border); }
      .fin-table__subtotal-row td { font-weight:600; border-top:1px solid var(--border); }
      .fin-table__subtotal-row td.fin-table__label-col { background:#0B0F19; }
      .fin-table .delta { font-size:11px; padding:2px 6px; justify-content:flex-end; }

      /* ---- Key Performance Indicators table: compact, emphasized business-line rows ---- */
      .fin-table-wrap--kpi { border-color:rgba(var(--gold-rgb),0.4); border-width:1.5px; }
      .fin-table--kpi tbody tr td.fin-table__label-col { font-weight:600; color:var(--brand); }
      .fin-table--kpi tbody tr td:not(.fin-table__label-col) { font-family:'Inter',sans-serif; font-weight:500; font-variant-numeric:tabular-nums; }
      .fin-table--kpi tbody tr:not(:last-child) td { border-bottom:1px solid var(--border); }
      .fin-table__na { color:var(--muted); font-weight:400 !important; }
      .fin-table__foot-note { display:flex; align-items:flex-start; gap:6px; margin-top:10px; font-size:11.5px; color:var(--muted); line-height:1.6; }

      /* ---- source links (used throughout Industry & Competitors / News) ---- */
      .src-link { display:inline-flex; align-items:center; gap:2px; font-family:'IBM Plex Mono',monospace; font-size:11px;
                  color:var(--brand); text-decoration:none; border-bottom:1px dotted var(--brand); white-space:nowrap; transition:color .15s, border-color .15s; }
      .src-link:hover { color:var(--gold); border-bottom-color:var(--gold); }
      .src-link__arrow { font-size:10px; }

      /* ---- News & Updates ---- */
      .news-refresh { display:flex; flex-direction:column; align-items:flex-end; gap:2px; font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--muted); }
      .news-refresh__stamp { color:var(--brand); font-weight:600; }
      .news-refresh__note { color:var(--muted); }
      .news-filters { display:flex; flex-direction:column; gap:8px; margin-bottom:18px; }
      .news-filters__group { display:flex; gap:6px; flex-wrap:wrap; }
      .chip { font-family:'Inter',sans-serif; font-size:12px; padding:6px 12px; border-radius:20px; border:1px solid var(--border);
              background:var(--surface); color:var(--muted); cursor:pointer; transition:all .18s; white-space:nowrap; backdrop-filter:blur(10px); }
      .chip--mono { font-family:'IBM Plex Mono',monospace; font-size:11px; }
      .chip:hover { border-color:var(--brand); color:var(--brand); }
      .chip--active { background:var(--brand); border-color:var(--brand); color:#052420; font-weight:600; box-shadow:0 0 16px -2px rgba(var(--brand-rgb),0.6); }
      .news-expanded-note { font-size:12px; color:var(--gold); margin-bottom:14px; font-style:italic; }
      .news-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(300px, 1fr)); gap:14px; }
      .news-card { border-radius:12px; padding:16px 18px; display:flex; flex-direction:column; gap:8px; transition:transform .2s, box-shadow .2s; }
      .news-card:hover { transform:translateY(-2px); }
      .news-card__top { display:flex; justify-content:space-between; align-items:center; gap:8px; }
      .news-card__category { font-family:'IBM Plex Mono',monospace; font-size:10px; text-transform:uppercase; letter-spacing:0.05em;
                              color:var(--brand); background:rgba(var(--brand-rgb),0.12); padding:3px 8px; border-radius:6px; }
      .news-card__date { font-family:'IBM Plex Mono',monospace; font-size:10.5px; color:var(--muted); }
      .news-card__title { font-family:'Space Grotesk',sans-serif; font-size:16px; font-weight:600; line-height:1.35; color:var(--ink); }
      .news-card__summary { font-size:12.5px; color:var(--muted); line-height:1.55; }
      .news-card__sources { display:flex; gap:12px; flex-wrap:wrap; margin-top:auto; padding-top:6px; }

      /* ---- Industry & Competitors ---- */
      .snapshot-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(220px, 1fr)); gap:14px; }
      .snapshot-tile { border-radius:12px; padding:16px; transition:transform .2s; }
      .snapshot-tile:hover { transform:translateY(-2px); }
      .snapshot-tile__label { font-size:12px; color:var(--muted); font-weight:500; margin-bottom:8px; }
      .snapshot-tile__value { font-family:'Space Grotesk',sans-serif; font-size:22px; font-weight:600; margin-bottom:8px; color:var(--ink); }
      .snapshot-tile__meta { display:flex; align-items:center; justify-content:space-between; gap:8px; }
      .snapshot-tile__period { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--gold); background:rgba(var(--gold-rgb),0.14); padding:2px 7px; border-radius:6px; }
      .snapshot-tile__note { font-size:11.5px; color:var(--muted); margin-top:8px; line-height:1.5; }

      .trend-card { display:flex; flex-direction:column; }
      .trend-card__desc { font-size:12.5px; color:var(--ink); line-height:1.6; margin:4px 0 10px; }
      .trend-card__why { font-size:12px; color:var(--muted); line-height:1.55; background:rgba(255,255,255,0.03); border:1px solid var(--border); border-radius:8px; padding:10px 12px; margin-bottom:10px; }
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
                        border-radius:10px; padding:12px 14px; }
      .analysis-item__badge { flex-shrink:0; font-family:'IBM Plex Mono',monospace; font-size:9.5px; text-transform:uppercase; letter-spacing:0.05em;
                               color:var(--gold); background:rgba(var(--gold-rgb),0.14); padding:3px 8px; border-radius:6px; }

      /* ---- shared motion utilities (paired with Framer Motion in JSX) ---- */
      .skeleton-pulse { background:linear-gradient(90deg, rgba(255,255,255,0.04) 25%, rgba(255,255,255,0.09) 50%, rgba(255,255,255,0.04) 75%);
                         background-size:200% 100%; animation:skeletonShimmer 1.4s ease-in-out infinite; }
      @keyframes skeletonShimmer { 0% { background-position:200% 0; } 100% { background-position:-200% 0; } }
    `}</style>
  );
}