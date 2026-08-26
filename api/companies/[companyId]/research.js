// GET /api/companies/:companyId/research — cached News & Updates / Industry
// & Competitors data for one company, refreshed server-side via Tavily at
// most once/day (see lib/research.js). Completely separate data path from
// /dashboard above — this never touches or influences any financial figure.
// Never called just because the user switched tabs; the frontend fetches
// this once per company selection and the server itself decides whether the
// cache is stale enough to warrant a fresh (bounded, ~2-search) research
// cycle. Pass ?refresh=1 to force a refresh (e.g. a manual "Refresh" button,
// if one is added) — still capped by the same per-cycle search budget.
import { getCompanyResearch } from "../../../lib/research.js";
import { sendJson, methodNotAllowed, withErrorHandling } from "../../../lib/apiHelpers.js";

export default withErrorHandling(async (req, res) => {
  if (req.method !== "GET") return methodNotAllowed(res, ["GET"]);
  const { companyId, refresh } = req.query;
  const result = await getCompanyResearch(companyId, { force: refresh === "1" || refresh === "true" });
  if (!result.ok) return sendJson(res, 404, result);
  sendJson(res, 200, result);
});
