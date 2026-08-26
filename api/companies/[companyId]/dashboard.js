// GET /api/companies/:companyId/dashboard — the core read endpoint. Returns
// { parsed, companyInfo, configKey, company, uploadMeta } reconstructed from
// this company's CURRENT normalized mis_metrics rows — i.e. exactly the raw
// shape parseWorkbook() would have produced from a live Excel file, NOT a
// precomputed dashboard. The frontend calls the shared buildDataset(parsed,
// companyInfo, COMPANY_CONFIGS[configKey]) locally with this response, which
// is what guarantees byte-identical output to the original client-only
// dashboard for the same underlying data (see src/lib/misEngine.js header).
import { getCompanyDashboardInput } from "../../../lib/dashboardRead.js";
import { sendJson, methodNotAllowed, withErrorHandling } from "../../../lib/apiHelpers.js";

export default withErrorHandling(async (req, res) => {
  if (req.method !== "GET") return methodNotAllowed(res, ["GET"]);
  const { companyId } = req.query;
  const result = await getCompanyDashboardInput(companyId);
  if (!result.ok) return sendJson(res, 404, result);
  sendJson(res, 200, result);
});
