// GET /api/companies/:companyId/report — generates and streams back a
// one-company summary PPTX.
//
// Fetches this company's current dashboard input the SAME way GET
// .../dashboard does (getCompanyDashboardInput — the Supabase-reconstructed
// { months, kpis } shape the browser's own dashboard renders from), then runs
// the SAME buildDataset() the browser runs on it and hands the result to
// buildReportContent() (src/lib/reportContent.js), which produces the deck's
// three sections — Key Performance Indicators, Key Financial Highlights, and
// the Business Description + Performance Summary narrative — already
// formatted, over the last two financial years and the last two quarters.
//
// That shared path is the point: the deck's figures are the identical strings
// the website prints for the same company, because they come from the same
// engine, the same company config and the same formatters. Previously this
// endpoint shipped raw rows to Python, which re-picked revenue/margin/EBITDA
// by generic keyword matching over the last two *months* — a real source of
// divergence between the deck and the dashboard, now removed.
//
// Node still owns all Supabase access here, per lib/supabaseAdmin.js's own
// rule; api/reports/generate.py never touches Supabase, it only renders a
// pptx from the JSON it's handed.
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { getCompanyDashboardInput } from "../../../lib/dashboardRead.js";
import { COMPANY_CONFIGS, buildDataset } from "../../../src/lib/misEngine.js";
import { buildReportContent } from "../../../src/lib/reportContent.js";
import { sendJson, methodNotAllowed, withErrorHandling } from "../../../lib/apiHelpers.js";

const execFileAsync = promisify(execFile);

async function renderReportLocally(payload) {
  const tmpDir = os.tmpdir();
  const jsonPath = path.join(tmpDir, `report_input_${Date.now()}_${Math.random().toString(36).slice(2)}.json`);
  const pptxPath = path.join(tmpDir, `report_output_${Date.now()}_${Math.random().toString(36).slice(2)}.pptx`);
  const scriptPath = path.resolve(process.cwd(), "api/reports/_generate_company_report.py");

  await fs.promises.writeFile(jsonPath, JSON.stringify(payload), "utf-8");

  try {
    try {
      await execFileAsync("py", [scriptPath, jsonPath, pptxPath]);
    } catch {
      await execFileAsync("python", [scriptPath, jsonPath, pptxPath]);
    }
    const data = await fs.promises.readFile(pptxPath);
    return data;
  } finally {
    fs.promises.unlink(jsonPath).catch(() => {});
    fs.promises.unlink(pptxPath).catch(() => {});
  }
}

export default withErrorHandling(async (req, res) => {
  if (req.method !== "GET") return methodNotAllowed(res, ["GET"]);
  const { companyId } = req.query;

  const result = await getCompanyDashboardInput(companyId);
  if (!result.ok) return sendJson(res, 404, result);

  const { parsed, companyInfo, configKey, company } = result;

  const companyConfig = COMPANY_CONFIGS[configKey];
  if (!companyConfig) {
    return sendJson(res, 500, { ok: false, error: `Unknown company config key "${configKey}".` });
  }

  const ds = buildDataset(parsed, companyInfo, companyConfig);
  const payload = buildReportContent(ds, {
    legalEntityName: company.legalName || company.name,
    brand: company.name,
  });

  const protocol = req.headers["x-forwarded-proto"] || "https";
  const origin = `${protocol}://${req.headers.host}`;

  let buffer;
  let disposition = `attachment; filename="${company.name}_Report.pptx"`;

  let pyRes;
  try {
    pyRes = await fetch(`${origin}/api/reports/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (pyRes.ok) {
      buffer = Buffer.from(await pyRes.arrayBuffer());
      if (pyRes.headers.get("content-disposition")) {
        disposition = pyRes.headers.get("content-disposition");
      }
    }
  } catch {
    // Network / local dev fallback
  }

  if (!buffer) {
    try {
      buffer = await renderReportLocally(payload);
    } catch (localErr) {
      return sendJson(res, 502, {
        ok: false,
        error: `Report rendering failed: ${pyRes ? `HTTP ${pyRes.status}` : ""}; fallback error: ${localErr.message || localErr}`,
      });
    }
  }

  res.status(200);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  res.setHeader("Content-Disposition", disposition);
  res.end(buffer);
});
