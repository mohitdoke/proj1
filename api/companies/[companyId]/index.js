// GET /api/companies/:companyId — basic company info (accepts UUID or
// slug), independent of any fund. Mainly useful for deep-linking straight
// to a company (e.g. /dashboard?fund=fund1&company=easyrewardz) before the
// fund's company list has loaded.
import { getCompany } from "../../../lib/dashboardRead.js";
import { sendJson, methodNotAllowed, withErrorHandling } from "../../../lib/apiHelpers.js";

export default withErrorHandling(async (req, res) => {
  if (req.method !== "GET") return methodNotAllowed(res, ["GET"]);
  const { companyId } = req.query;
  const company = await getCompany(companyId);
  if (!company) return sendJson(res, 404, { ok: false, error: `Company "${companyId}" not found.` });
  sendJson(res, 200, { ok: true, company });
});
