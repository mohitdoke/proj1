# MIS Dashboard

A consolidated MIS (Monthly Information System) dashboard for IIFL Fintech Fund portfolio companies. Managers upload each company's monthly MIS workbook; everyone else just picks a fund, picks a company, and sees that company's dashboard — no upload required, ever, for a normal viewer.

## Architecture

```
Manager uploads .xlsx
        |
        v
POST /api/admin/mis/upload  ──>  lib/misProcessing.js
        |                          - parse workbook (src/lib/misEngine.js)
        |                          - detect company from the sheet's own
        |                            row names (never the filename)
        |                          - validate it produces a real dashboard
        |                          - store the raw file in Supabase Storage
        |                          - store normalized monthly values in
        |                            Postgres (mis_metrics)
        |                          - only then flip this company's
        |                            "current" version
        v
   Supabase (Postgres + Storage)
        ^
        |
GET /api/companies/:id/dashboard  ──>  lib/dashboardRead.js
        |                                - reconstructs the exact
        |                                  { months, kpis, headcount }
        |                                  shape a live Excel parse would
        |                                  have produced
        v
   Browser (React)
        |
        - runs the SAME buildDataset() (src/lib/misEngine.js) the original
          client-only dashboard used, on that data
        - renders the SAME dashboard UI (src/MISDashboard.jsx), unchanged
```

The key idea: **all the Excel-parsing, company-detection, and financial-calculation logic lives in one file, `src/lib/misEngine.js`**, imported by both the browser bundle and the server (Vercel functions, the seed script). The backend's job is only to get the *same raw inputs* (`parsed`, `companyInfo`, a company config key) to that function that a live Excel upload would have produced — it never recomputes or re-derives financial figures itself. That's what guarantees the backend-driven dashboard produces the same output as the original upload-driven one for the same underlying data; see `scripts/parity_check.mjs`, which proves this for all 9 companies' real MIS files.

News & Industry research (Tavily-backed) is a completely separate, independent path — it never touches or influences any financial figure, calculation, or the `mis_metrics` table.

## Project structure

```
src/
  lib/misEngine.js       Shared calculation engine (parsing, company config,
                          detection, FY/quarter math, formatting). Runs in
                          both the browser and Node.
  lib/misMetricsShape.js Pure parsed <-> mis_metrics-rows conversions,
                          shared by the write path and the read path.
  lib/apiClient.js        Frontend fetch wrappers around /api/*.
  MISDashboard.jsx         The dashboard UI itself — DashboardView (the
                          actual dashboard, unchanged from the original
                          single-file app) + App (a standalone local-upload
                          entry point, kept for admin preview / offline use).
  AppRoot.jsx              The app normal users load: fund/company picker +
                          manager upload panel, wired to the API.
  main.jsx                 Entry point — renders AppRoot.

lib/                       Server-only code (never bundled into the browser).
  supabaseAdmin.js          Service-role Supabase client.
  misProcessing.js          The upload pipeline (parse -> detect -> validate
                          -> store file -> store metrics -> version swap).
  dashboardRead.js          Read path: funds/companies listing, and
                          reconstructing dashboard input from stored metrics.
  research.js               Tavily-backed News/Industry research, cached.
  apiHelpers.js             Small response/auth helpers for api/*.js.

api/                        Vercel serverless functions (file-based routing).
  funds/index.js                    GET  /api/funds
  funds/[fundId]/companies.js       GET  /api/funds/:fundId/companies
  companies/[companyId]/index.js    GET  /api/companies/:companyId
  companies/[companyId]/dashboard.js GET /api/companies/:companyId/dashboard
  companies/[companyId]/research.js GET  /api/companies/:companyId/research
  admin/mis/upload.js               POST /api/admin/mis/upload

supabase/
  migrations/0001_init.sql   Full schema (funds, companies, fund_companies,
                          company_configs, mis_uploads, mis_metrics,
                          company_research) + RLS policies.
  seed/reference_data.sql    Funds, companies, fund membership, and each
                          company's config_key (NOT financial data).

scripts/
  seed.mjs                   Imports each company's real master .xlsx into
                          Supabase via the same pipeline a manager upload
                          uses. Run once after migrations + reference seed.
  parity_check.mjs           Proves old-dashboard-output === new-backend-
                          driven-output for every company, without needing
                          a live Supabase project.
```

## Local setup

```bash
npm install
cp .env.example .env   # fill in the values below
```

### 1. Create a Supabase project

Create a project at supabase.com, then from **Project Settings -> API** copy:
- the Project URL -> `SUPABASE_URL`
- the `anon` `public` key -> `SUPABASE_ANON_KEY`
- the `service_role` key -> `SUPABASE_SERVICE_ROLE_KEY` (keep this secret — it bypasses Row Level Security and must never reach the browser)

### 2. Apply the schema

In the Supabase SQL editor (or via the `supabase` CLI), run, in order:
1. `supabase/migrations/0001_init.sql` — creates every table, index, and RLS policy.
2. `supabase/seed/reference_data.sql` — inserts the two funds, the 9 companies, their `company_configs.config_key`, and starting fund membership. This is reference data only (names/slugs) — no financial numbers.

### 3. Create the Storage bucket

In Supabase Storage, create a bucket named exactly `mis-files`, and make it **private** (not public). The service-role key (used server-side only) reads and writes it; nothing else can.

### 4. Set the remaining env vars

- `ADMIN_UPLOAD_TOKEN` — any long random string; required to call the manager upload endpoint (see "Manager authentication" below).
- `TAVILY_API_KEY` — from tavily.com, for the News/Industry research service.

### 5. Import the currently-available MIS data

```bash
set -a; source .env; set +a   # or export the vars another way
npm run seed
```

This runs `scripts/seed.mjs`, which reads each company's real master `.xlsx` file already in this repo (`excel_template/`, `misdata/`) and imports it through the exact same pipeline (`lib/misProcessing.js`) a manager's HTTP upload uses — **never** by typing values into SQL. Safe to re-run; pass `--only=grayquest,riskcovry` to import a subset.

### 6. Run it

```bash
npm run dev      # frontend, proxying /api to nothing locally —
                  # use `vercel dev` instead to also run the API routes:
vercel dev
```

`vercel dev` runs both the Vite frontend and the `api/*.js` serverless functions together, which is what you want for anything beyond pure UI work.

## Manager MIS upload

Open the app, click **"Manager upload"** next to the fund selector, paste in `ADMIN_UPLOAD_TOKEN`, optionally type the company's slug as a cross-check, choose the `.xlsx` file, and submit. The server:
1. Parses the workbook and detects the company from its own row names (never the filename, never your dropdown choice — that's only a cross-check).
2. Refuses with a clear error if it can't confidently identify the company (`"Unable to confidently identify company from MIS."`), rather than silently defaulting to any existing company's template.
3. Validates the parse actually produces a usable dashboard (a recognizable revenue row, etc.) before storing anything.
4. Stores the raw file in the private `mis-files` bucket and the normalized monthly values in `mis_metrics`.
5. Only once all of that has succeeded does it flip this company's "current" version — a failed upload never touches or corrupts the previously-live dataset.

## Adding a new company

1. Add a `COMPANY_CONFIGS` entry for it in `src/lib/misEngine.js` (signals, revenue key, KPI table, etc.) — same as adding a company to the original dashboard.
2. Insert its row into `companies` and `company_configs` (with `config_key` matching the new `COMPANY_CONFIGS` key), e.g. by adding it to `supabase/seed/reference_data.sql` and re-running that file.
3. Upload its MIS (via the manager upload flow, or add it to `scripts/seed.mjs`'s `SOURCE_FILES` list and run `npm run seed`).

## Assigning a company to a fund

`fund_companies` is a plain join table (`fund_id`, `company_id`, `display_order`). Move a company between funds, or change display order, with a one-row `update`/`insert` — no code or deploy needed. See the comments in `supabase/seed/reference_data.sql` for the starting membership.

## Manager authentication

The current gate on `POST /api/admin/mis/upload` is a single shared secret (`ADMIN_UPLOAD_TOKEN`), checked in `lib/apiHelpers.js`. This is a deliberately simple starting point, not a full auth system. To upgrade it to real per-manager accounts without touching the upload pipeline itself:
1. Enable Supabase Auth (email/magic-link) and create accounts for managers.
2. Add a Postgres check (or an RLS policy plus a thin server-side check) restricting upload to a specific set of manager email addresses/roles.
3. Replace `requireAdmin()` in `lib/apiHelpers.js` with a check against the authenticated Supabase session instead of the shared token — `processMisUpload()` and everything else in the pipeline stays exactly the same.

## Tavily research (News & Industry)

`lib/research.js` runs at most two Tavily searches per company per refresh cycle (basic search depth, ~5 results each, no crawling/extraction), caches the result in `company_research`, and only refreshes once the cache is more than a day old (or when `?refresh=1` is passed) — never just because a user switched tabs. Results are reshaped into the exact data structures the existing `NewsUpdatesPage`/`IndustryCompetitorsPage` components already render, so no UI changed to support this. If `TAVILY_API_KEY` is unset or a request fails, the endpoint degrades to the last successful cache (or an empty result) rather than erroring the whole dashboard.

## Verifying nothing regressed

```bash
npm run parity-check
```

This proves, for every company this repo has a real MIS file for, that parsing it directly and running it through the full write-path-flatten / read-path-reconstruct round trip produce **identical** `buildDataset()` output — company detection, every FY/quarter figure, every KPI card. It needs no live Supabase project (it simulates the storage round trip in memory using the same `lib/misMetricsShape.js` functions the real pipeline uses), so it's safe to run in CI on every change to `src/lib/misEngine.js` or `lib/misProcessing.js`/`lib/dashboardRead.js`.

## Deploying to Vercel

1. `vercel link` this repo to a Vercel project.
2. Add every variable from `.env.example` to the project's Environment Variables (Production **and** Preview).
3. Deploy (`vercel --prod` or push to the connected Git branch). `vercel.json` configures the build and a couple of function timeouts; everything else uses Vercel's zero-config Vite + Node serverless function detection.
4. Run the Supabase migration/seed/reference-data steps above against the **production** Supabase project before (or right after) the first deploy — the app has nothing to show until at least one company has a successful `mis_uploads` row.
