// Thin fetch wrappers around this app's own /api/* backend. Nothing here
// talks to Supabase directly — the browser only ever calls our own API,
// which holds the service-role key server-side (see lib/supabaseAdmin.js).
async function getJson(url) {
  const res = await fetch(url);
  let body;
  try {
    body = await res.json();
  } catch {
    throw new Error(`Request to ${url} failed (${res.status}) and returned a non-JSON response.`);
  }
  if (!res.ok || body.ok === false) {
    throw new Error(body?.error || `Request to ${url} failed (${res.status}).`);
  }
  return body;
}

export function fetchFunds() {
  return getJson("/api/funds").then(b => b.funds);
}

export function fetchCompaniesForFund(fundSlug) {
  return getJson(`/api/funds/${encodeURIComponent(fundSlug)}/companies`).then(b => ({ fund: b.fund, companies: b.companies }));
}

export function fetchCompanyDashboardInput(companySlug) {
  return getJson(`/api/companies/${encodeURIComponent(companySlug)}/dashboard`);
}

/** Research items come back from the API as JSON (dates serialized as ISO
 * strings inside news_data); rehydrate `publishedAt` into real Date objects
 * so NewsUpdatesPage's date-range filtering (which does arithmetic on
 * `f.publishedAt`) behaves exactly as it did with dates parsed straight out
 * of an Excel sheet. */
export async function fetchCompanyResearch(companySlug, { refresh = false } = {}) {
  const qs = refresh ? "?refresh=1" : "";
  const body = await getJson(`/api/companies/${encodeURIComponent(companySlug)}/research${qs}`);
  return {
    newsFeed: (body.newsFeed || []).map(item => ({ ...item, publishedAt: item.publishedAt ? new Date(item.publishedAt) : null })),
    industryData: body.industryData || null,
    refreshMeta: body.refreshMeta || null,
  };
}

/** Downloads this company's auto-generated summary report (.pptx) and
 * triggers a browser save, same pattern as the existing Excel export button
 * uses (an in-memory blob + a throwaway <a download> click, no server-side
 * file persisted anywhere). companyId can be the slug or the UUID - the
 * backend route accepts either. */
export async function downloadCompanyReport(companyId) {
  const res = await fetch(`/api/companies/${encodeURIComponent(companyId)}/report`);
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body?.error || `Report generation failed (${res.status}).`);
  }
  const blob = await res.blob();
  const disposition = res.headers.get("content-disposition") || "";
  const match = disposition.match(/filename="([^"]+)"/);
  const filename = match ? match[1] : "Report.pptx";

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/** Manager-only upload. `adminToken` goes in the Authorization header (see
 * lib/apiHelpers.js requireAdmin) — never logged, never stored beyond the
 * caller's own component state. */
export async function uploadMisFile({ file, companySlug, adminToken, uploadedBy }) {
  const form = new FormData();
  form.append("file", file);
  if (companySlug) form.append("companySlug", companySlug);
  if (uploadedBy) form.append("uploadedBy", uploadedBy);

  const res = await fetch("/api/admin/mis/upload", {
    method: "POST",
    headers: { Authorization: `Bearer ${adminToken}` },
    body: form,
  });
  const body = await res.json().catch(() => ({ ok: false, error: `Upload failed (${res.status}).` }));
  return body;
}
