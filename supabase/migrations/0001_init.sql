-- ============================================================
-- MIS Dashboard — core schema
-- ============================================================
-- Normalized data model shared by ALL companies (current and future).
-- No per-company tables. Business/calculation logic (Excel parsing,
-- company detection, semantic KPI mapping, derived metrics) stays in
-- application code (src/lib/misEngine.js) — this schema only stores:
--   - reference data (funds, companies, fund membership)
--   - upload/version metadata
--   - normalized monthly metric values
--   - cached web-research results (news / industry — separate from MIS)
-- No executable JavaScript, and no precomputed financial aggregates,
-- are stored here (see README "Database vs calculation logic").
-- ============================================================

create extension if not exists "pgcrypto"; -- gen_random_uuid()

-- ------------------------------------------------------------
-- FUNDS
-- ------------------------------------------------------------
create table if not exists funds (
  id          uuid primary key default gen_random_uuid(),
  name        text not null,
  slug        text not null unique,
  created_at  timestamptz not null default now()
);

-- ------------------------------------------------------------
-- COMPANIES
-- ------------------------------------------------------------
-- `slug` must match a key in COMPANY_CONFIGS (src/lib/misEngine.js),
-- e.g. "easyrewardz", "grayquest", "apexFutureLabs", "finbox".
-- That's the one link between a DB row and the code-side calculation
-- config for that company.
create table if not exists companies (
  id          uuid primary key default gen_random_uuid(),
  slug        text not null unique,
  name        text not null,          -- brand/display name, e.g. "Vitra.ai"
  legal_name  text,                   -- e.g. "Apex Future Labs Private Limited"
  active      boolean not null default true,
  created_at  timestamptz not null default now()
);

-- ------------------------------------------------------------
-- FUND_COMPANIES — which companies appear under which fund, and in
-- what order. A company can belong to more than one fund.
-- ------------------------------------------------------------
create table if not exists fund_companies (
  fund_id       uuid not null references funds(id) on delete cascade,
  company_id    uuid not null references companies(id) on delete cascade,
  display_order int not null default 0,
  active        boolean not null default true,
  primary key (fund_id, company_id)
);
create index if not exists idx_fund_companies_fund on fund_companies(fund_id) where active;

-- ------------------------------------------------------------
-- COMPANY_CONFIGS — data-safe configuration/overrides layered on top
-- of the code-side COMPANY_CONFIGS entry. `config_key` is required and
-- must match a key in src/lib/misEngine.js's COMPANY_CONFIGS; the
-- regex matchers, formatters and calculation logic for that key live
-- in code, never here. `overrides` may carry display-only metadata
-- (e.g. a company description/tags override) — never calculation
-- rules, and never executable code.
-- ------------------------------------------------------------
create table if not exists company_configs (
  company_id   uuid primary key references companies(id) on delete cascade,
  config_key   text not null,
  overrides    jsonb not null default '{}'::jsonb,
  updated_at   timestamptz not null default now()
);

-- ------------------------------------------------------------
-- MIS_UPLOADS — one row per manager upload attempt (a version).
-- ------------------------------------------------------------
create table if not exists mis_uploads (
  id                    uuid primary key default gen_random_uuid(),
  company_id            uuid not null references companies(id) on delete cascade,
  original_filename     text not null,
  storage_path          text not null,      -- path inside the private "mis-files" bucket
  uploaded_at           timestamptz not null default now(),
  uploaded_by           text,               -- free-text identifier (email/name) of the manager; see README on auth
  processing_status     text not null default 'pending'
                          check (processing_status in ('pending','processing','succeeded','failed')),
  processing_error      text,
  detected_config_key   text,               -- COMPANY_CONFIGS key detectCompanyConfig() resolved to
  months_count          int,
  -- The full ordered [{y, m, label, key}, ...] month axis for this upload
  -- (buildDataset()'s own forceConsecutiveMonths-corrected list — see
  -- misEngine.js resolveMonths()), stored explicitly rather than inferred
  -- from mis_metrics' distinct period_dates. A calendar month where every
  -- single KPI is genuinely blank (a real case in at least one source
  -- workbook, which pre-lists trailing not-yet-reported months in its
  -- header row) would otherwise leave literally zero metric rows for that
  -- month, making it impossible to tell "this month exists with no data
  -- yet" from "this month was never part of the sheet" — this column is
  -- the authoritative month axis either way.
  months                jsonb,
  kpi_count             int,
  is_current            boolean not null default false,
  created_at            timestamptz not null default now()
);
create index if not exists idx_mis_uploads_company on mis_uploads(company_id);
-- At most one CURRENT upload per company at a time.
create unique index if not exists idx_mis_uploads_current
  on mis_uploads(company_id) where is_current;

-- ------------------------------------------------------------
-- MIS_METRICS — normalized monthly values for one upload. One row per
-- (metric, month). This is the "Monthly Data" sheet, flattened.
-- ------------------------------------------------------------
create table if not exists mis_metrics (
  id                    bigint generated always as identity primary key,
  upload_id             uuid not null references mis_uploads(id) on delete cascade,
  company_id            uuid not null references companies(id) on delete cascade, -- denormalized for RLS/query convenience
  metric_key            text not null,   -- normalized (trimmed) form of the row label, for indexing/search
  original_metric_label text not null,   -- EXACT row label as it appeared in the sheet (case/spacing preserved —
                                          -- this is what the calculation engine's matchers key off of)
  period_date           date not null,   -- first day of the month this value covers
  value                 numeric,         -- null = genuinely blank in the source sheet (never coerced to 0)
  unit                  text,            -- informational only ("INR", "count", "percent", ...); never used in calc
  created_at            timestamptz not null default now()
);
create unique index if not exists idx_mis_metrics_unique
  on mis_metrics(upload_id, original_metric_label, period_date);
create index if not exists idx_mis_metrics_upload on mis_metrics(upload_id);
create index if not exists idx_mis_metrics_company on mis_metrics(company_id);

-- ------------------------------------------------------------
-- COMPANY_RESEARCH — cached Tavily-sourced News & Industry data, kept
-- completely separate from financial MIS data. One row per company;
-- refreshed periodically rather than on every page view (see README).
-- ------------------------------------------------------------
create table if not exists company_research (
  company_id    uuid primary key references companies(id) on delete cascade,
  news_data     jsonb,             -- array of {title, summary, category, publishedAt, sourceName, sourceUrl, ...}
  industry_data jsonb,             -- {overviewDescription, categories[], snapshot[], trends[], competitors[], analysis[], methodology}
  fetched_at    timestamptz,
  fetch_error   text,
  updated_at    timestamptz not null default now()
);

-- ============================================================
-- ROW LEVEL SECURITY
-- ============================================================
-- The frontend never talks to Supabase directly — every read/write goes
-- through this repo's own /api/* serverless functions using the
-- service-role key server-side (which bypasses RLS by design). RLS
-- below is defense-in-depth / safe-by-default for the anon/public key:
-- reference data and cached research are readable, raw MIS uploads and
-- metrics are not exposed to anon/public at all (no policy = default
-- deny once RLS is enabled).

alter table funds             enable row level security;
alter table companies         enable row level security;
alter table fund_companies    enable row level security;
alter table company_configs   enable row level security;
alter table mis_uploads       enable row level security;
alter table mis_metrics       enable row level security;
alter table company_research  enable row level security;

create policy "funds are publicly readable" on funds
  for select using (true);
create policy "companies are publicly readable" on companies
  for select using (true);
create policy "fund_companies are publicly readable" on fund_companies
  for select using (true);
create policy "company_configs are publicly readable" on company_configs
  for select using (true);
create policy "company_research is publicly readable" on company_research
  for select using (true);

-- Intentionally NO anon/authenticated policies on mis_uploads / mis_metrics —
-- only the service role (used exclusively server-side) can read/write them.
