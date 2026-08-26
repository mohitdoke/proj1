import React, { useState, useEffect, useCallback, useMemo } from "react";
import { COMPANY_CONFIGS, buildDataset } from "./lib/misEngine.js";
import { fetchFunds, fetchCompaniesForFund, fetchCompanyDashboardInput, fetchCompanyResearch, uploadMisFile } from "./lib/apiClient.js";
import { DashboardView, GlobalStyles } from "./MISDashboard.jsx";

/* ============================================================
   AppRoot — the app normal users actually load. No Excel upload
   anywhere in this component: pick a fund, pick a company, the backend's
   already-processed data renders through the UNCHANGED DashboardView.

   Flow: fetch funds -> fetch companies for the selected fund -> fetch that
   company's current dataset (parsed + companyInfo + configKey) from
   /api/companies/:id/dashboard -> run the SAME buildDataset() the original
   client-only dashboard used, locally, on that data -> attach the
   separately-fetched News/Industry research -> render DashboardView.

   Selected fund/company are kept in the URL (?fund=...&company=...) so
   refresh, sharing a link, and browser back/forward all keep working
   without ever re-asking for an upload.
   ============================================================ */

function readParams() {
  const p = new URLSearchParams(window.location.search);
  return { fund: p.get("fund") || null, company: p.get("company") || null, admin: p.get("admin") === "1" };
}

function writeParams({ fund, company }) {
  const p = new URLSearchParams(window.location.search);
  if (fund) p.set("fund", fund); else p.delete("fund");
  if (company) p.set("company", company); else p.delete("company");
  const next = `${window.location.pathname}?${p.toString()}`;
  window.history.replaceState(null, "", next);
}

export default function AppRoot() {
  const initial = useMemo(readParams, []);
  const [funds, setFunds] = useState(null);
  const [fundsError, setFundsError] = useState(null);
  const [fundSlug, setFundSlug] = useState(initial.fund);
  const [companies, setCompanies] = useState(null);
  const [companySlug, setCompanySlug] = useState(initial.company);
  const [ds, setDs] = useState(null);
  const [dashError, setDashError] = useState(null);
  const [loadingDashboard, setLoadingDashboard] = useState(false);
  const [researchWarning, setResearchWarning] = useState(null);
  const [showAdmin, setShowAdmin] = useState(initial.admin);

  // 1. Load funds once.
  useEffect(() => {
    fetchFunds()
      .then(list => {
        setFunds(list);
        if (!fundSlug && list.length) setFundSlug(list[0].slug);
      })
      .catch(err => setFundsError(String(err.message || err)));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 2. Load this fund's companies whenever the selected fund changes.
  useEffect(() => {
    if (!fundSlug) return;
    let cancelled = false;
    fetchCompaniesForFund(fundSlug)
      .then(({ companies: list }) => {
        if (cancelled) return;
        setCompanies(list);
        // Keep the current company only if it actually belongs to this fund;
        // otherwise default to the fund's first company. Never guess across
        // funds — this only picks within the fund just confirmed to exist.
        setCompanySlug(prev => (prev && list.some(c => c.slug === prev)) ? prev : (list[0]?.slug || null));
      })
      .catch(err => setFundsError(String(err.message || err)));
    return () => { cancelled = true; };
  }, [fundSlug]);

  // 3. Load the selected company's dashboard data + research whenever it changes.
  useEffect(() => {
    if (!companySlug) return;
    let cancelled = false;
    setLoadingDashboard(true);
    setDashError(null);
    setResearchWarning(null);

    Promise.all([
      fetchCompanyDashboardInput(companySlug),
      fetchCompanyResearch(companySlug).catch(err => ({ __researchFailed: true, error: String(err.message || err) })),
    ])
      .then(([dashboardInput, research]) => {
        if (cancelled) return;
        const companyConfig = COMPANY_CONFIGS[dashboardInput.configKey];
        if (!companyConfig) throw new Error(`Unknown company config key "${dashboardInput.configKey}".`);
        // The exact same calculation engine the original client-only
        // dashboard used — this is what makes the output identical for the
        // same underlying data, not a reimplementation.
        const built = buildDataset(dashboardInput.parsed, dashboardInput.companyInfo, companyConfig);
        if (research.__researchFailed) {
          built.newsFeed = [];
          built.industryData = null;
          built.refreshMeta = { newsRefreshedAt: null, industryRefreshedAt: null };
          setResearchWarning(`News & Industry research is temporarily unavailable (${research.error}).`);
        } else {
          built.newsFeed = research.newsFeed;
          built.industryData = research.industryData;
          built.refreshMeta = research.refreshMeta;
        }
        setDs({ built, uploadMeta: dashboardInput.uploadMeta, company: dashboardInput.company });
      })
      .catch(err => { if (!cancelled) { setDs(null); setDashError(String(err.message || err)); } })
      .finally(() => { if (!cancelled) setLoadingDashboard(false); });

    return () => { cancelled = true; };
  }, [companySlug]);

  // Keep the URL in sync so refresh/back-forward/sharing a link all work.
  useEffect(() => {
    writeParams({ fund: fundSlug, company: companySlug });
  }, [fundSlug, companySlug]);

  const selectedFund = funds?.find(f => f.slug === fundSlug) || null;

  return (
    <div className="dash">
      {/* GlobalStyles defines the --ink/--brand/--border/... CSS variables
          RootStyles below (and the fund/company nav bar) rely on — render
          it unconditionally here so the nav bar is styled correctly even
          before a dashboard has loaded. DashboardView renders it again once
          a company is selected; a duplicate <style> tag is harmless. */}
      <GlobalStyles />
      <RootStyles />

      <div className="fund-nav">
        <div className="fund-nav__row">
          <span className="fund-nav__label">Fund</span>
          {fundsError && <span className="fund-nav__error">{fundsError}</span>}
          {(funds || []).map(f => (
            <button
              key={f.slug}
              className={`chip ${f.slug === fundSlug ? "chip--active" : ""}`}
              onClick={() => setFundSlug(f.slug)}
            >
              {f.name}
            </button>
          ))}
          <button className="fund-nav__admin-link" onClick={() => setShowAdmin(s => !s)}>
            {showAdmin ? "Hide manager upload" : "Manager upload"}
          </button>
        </div>
        <div className="fund-nav__row">
          <span className="fund-nav__label">Company</span>
          {(companies || []).map(c => (
            <button
              key={c.slug}
              className={`chip chip--mono ${c.slug === companySlug ? "chip--active" : ""}`}
              onClick={() => setCompanySlug(c.slug)}
            >
              {c.name}
            </button>
          ))}
          {companies && !companies.length && <span className="fund-nav__error">No companies configured for {selectedFund?.name || "this fund"} yet.</span>}
        </div>
      </div>

      {showAdmin && (
        <AdminUploadPanel
          defaultCompanySlug={companySlug}
          onUploaded={(uploadedSlug) => {
            if (uploadedSlug === companySlug) {
              // Re-trigger the dashboard fetch for the company just updated.
              setCompanySlug(null);
              setTimeout(() => setCompanySlug(uploadedSlug), 0);
            }
          }}
        />
      )}

      {loadingDashboard && <div className="fund-nav__status">Loading dashboard…</div>}
      {dashError && <div className="fund-nav__status fund-nav__status--error">{dashError}</div>}

      {ds && !loadingDashboard && !dashError && (
        <DashboardView
          ds={ds.built}
          fileName={ds.uploadMeta?.originalFilename}
          banner={researchWarning ? <div className="upload-bar upload-bar--info"><span>{researchWarning}</span></div> : null}
        />
      )}
    </div>
  );
}

function AdminUploadPanel({ defaultCompanySlug, onUploaded }) {
  const [file, setFile] = useState(null);
  const [companySlug, setCompanySlug] = useState(defaultCompanySlug || "");
  const [adminToken, setAdminToken] = useState("");
  const [status, setStatus] = useState(null);

  const handleSubmit = useCallback(async (e) => {
    e.preventDefault();
    if (!file) { setStatus({ type: "error", text: "Choose a .xlsx file first." }); return; }
    setStatus({ type: "info", text: `Uploading ${file.name}…` });
    try {
      const result = await uploadMisFile({ file, companySlug: companySlug || undefined, adminToken });
      if (result.ok) {
        setStatus({ type: "success", text: `Processed: ${result.companyName} — ${result.monthsCount} months, ${result.kpiCount} KPIs, ${result.fyCount} FY periods. This is now the current dataset.` });
        onUploaded?.(result.companySlug);
      } else {
        setStatus({ type: "error", text: result.error || "Upload failed." });
      }
    } catch (err) {
      setStatus({ type: "error", text: String(err.message || err) });
    }
  }, [file, companySlug, adminToken, onUploaded]);

  return (
    <form className="admin-panel" onSubmit={handleSubmit}>
      <div className="admin-panel__title">Manager MIS upload</div>
      <p className="admin-panel__note">
        Uploads are validated and processed server-side before they replace anything — if processing fails, the
        currently-live dataset is left untouched. Company is detected from the sheet's own row names, not the
        filename or this dropdown (this field is just a cross-check).
      </p>
      <div className="admin-panel__row">
        <input type="password" placeholder="Admin token" value={adminToken} onChange={e => setAdminToken(e.target.value)} />
        <input type="text" placeholder="Company slug (optional cross-check, e.g. easyrewardz)" value={companySlug} onChange={e => setCompanySlug(e.target.value)} />
        <input type="file" accept=".xlsx,.xls" onChange={e => setFile(e.target.files?.[0] || null)} />
        <button type="submit">Upload</button>
      </div>
      {status && <div className={`admin-panel__status admin-panel__status--${status.type}`}>{status.text}</div>}
    </form>
  );
}

function RootStyles() {
  return (
    <style>{`
      .fund-nav { padding:20px 40px 12px; border-bottom:1px solid var(--border); background:var(--surface); display:flex; flex-direction:column; gap:10px; }
      .fund-nav__row { display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
      .fund-nav__label { font-family:'IBM Plex Mono',monospace; font-size:10.5px; text-transform:uppercase; letter-spacing:0.06em; color:var(--muted); margin-right:4px; }
      .fund-nav__admin-link { margin-left:auto; font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--muted); background:none; border:none; text-decoration:underline; cursor:pointer; }
      .fund-nav__admin-link:hover { color:var(--brand); }
      .fund-nav__error { font-size:12px; color:var(--neg); }
      .fund-nav__status { padding:14px 40px; font-family:'IBM Plex Mono',monospace; font-size:12.5px; color:var(--muted); }
      .fund-nav__status--error { color:var(--neg); }

      .admin-panel { margin:16px 40px; padding:18px 20px; border:1px solid var(--border); border-radius:12px; background:#fff; }
      .admin-panel__title { font-family:'Fraunces',serif; font-size:16px; margin-bottom:6px; }
      .admin-panel__note { font-size:12px; color:var(--muted); line-height:1.6; margin-bottom:12px; }
      .admin-panel__row { display:flex; gap:8px; flex-wrap:wrap; align-items:center; }
      .admin-panel__row input[type="password"], .admin-panel__row input[type="text"] { font-size:12.5px; padding:8px 10px; border:1px solid var(--border); border-radius:8px; flex:1; min-width:160px; }
      .admin-panel__row button { font-size:12.5px; padding:8px 14px; border-radius:8px; border:1px solid var(--brand); background:var(--brand); color:#fff; cursor:pointer; }
      .admin-panel__status { margin-top:10px; font-family:'IBM Plex Mono',monospace; font-size:12px; padding:8px 12px; border-radius:8px; background:var(--surface); color:var(--muted); }
      .admin-panel__status--error { color:var(--neg); background:#FBF0EC; }
      .admin-panel__status--success { color:var(--pos); background:#EFF8F1; }
    `}</style>
  );
}
