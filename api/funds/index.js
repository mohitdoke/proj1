// GET /api/funds — list both funds, each with its ordered, active company
// list. This is the very first request the frontend makes on load.
import { listFunds } from "../../lib/dashboardRead.js";
import { sendJson, methodNotAllowed, withErrorHandling } from "../../lib/apiHelpers.js";

export default withErrorHandling(async (req, res) => {
  if (req.method !== "GET") return methodNotAllowed(res, ["GET"]);
  const funds = await listFunds();
  sendJson(res, 200, { ok: true, funds });
});
