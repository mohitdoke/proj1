// GET /api/funds/:fundId/companies — companies belonging to one fund
// (accepts either the fund's UUID or its slug, e.g. "fund-1"), in display
// order. Used when the user switches funds in the top-level fund selector.
import { listCompaniesForFund } from "../../../lib/dashboardRead.js";
import { sendJson, methodNotAllowed, withErrorHandling } from "../../../lib/apiHelpers.js";

export default withErrorHandling(async (req, res) => {
  if (req.method !== "GET") return methodNotAllowed(res, ["GET"]);
  const { fundId } = req.query;
  const result = await listCompaniesForFund(fundId);
  if (!result) return sendJson(res, 404, { ok: false, error: `Fund "${fundId}" not found.` });
  sendJson(res, 200, { ok: true, fund: result.fund, companies: result.companies });
});
