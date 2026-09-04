"""POST /api/reports/generate — renders a one-company summary PPTX.

Pure rendering function: takes the company's report content as JSON — already
computed and formatted by the Node side from the SAME buildDataset() the
website's own dashboard renders from (see src/lib/reportContent.js) — and
returns the generated .pptx as a binary response. Does not touch Supabase
itself; Node owns all DB/Storage access (see lib/supabaseAdmin.js's own "only
from api/*.js or scripts/*.mjs" rule).

It also does no arithmetic and no row-picking. Every figure arrives as a
finished string produced by the dashboard's own engine, company config and
formatters, so the deck cannot disagree with the page for the same company.
(This closes the divergence the previous version carried: it re-derived
revenue/margin/EBITDA rows here by generic keyword matching over the last two
months, independently of the website's COMPANY_CONFIGS-driven resolution.)

Expected POST body (application/json) — see buildReportContent():
{
  "legal_entity_name": str, "brand": str,
  "period_labels": [str, ...],           # last 2 FYs + last 2 quarters
  "business":       {"title", "description", "tags", "scaleMetrics"},
  "kpi_section":       {"title", "cornerLabel", "rows", "notes"} | null,
  "financial_section": {"title", "cornerLabel", "rows", "notes"} | null,
  "summary":        {"title", "eyebrow", "bullets", "tiles", "notes"},
  "footnotes":      [str, ...]
}
"""
import json
import os
import sys
import tempfile
from http.server import BaseHTTPRequestHandler

sys.path.insert(0, os.path.dirname(__file__))
from _generate_company_report import generate_report


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length) or b"{}")

            brand = body.get("brand") or body.get("legal_entity_name") or "Company"
            content = {
                "legal_entity_name": body.get("legal_entity_name") or brand,
                "brand": brand,
                "logo_path": body.get("logo_path"),
                "period_labels": body.get("period_labels") or [],
                "business": body.get("business") or {},
                "kpi_section": body.get("kpi_section"),
                "financial_section": body.get("financial_section"),
                "summary": body.get("summary") or {},
                "footnotes": body.get("footnotes") or [],
            }

            if not content["period_labels"]:
                self._send_json(400, {"error": "period_labels is required and must be non-empty."})
                return
            if not content["kpi_section"] and not content["financial_section"]:
                self._send_json(400, {"error": "At least one of kpi_section or financial_section is required."})
                return

            with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
                out_path = tmp.name
            try:
                generate_report(content, out_path)
                with open(out_path, "rb") as f:
                    data = f.read()
            finally:
                try:
                    os.remove(out_path)
                except OSError:
                    pass

            safe_name = "".join(c for c in brand if c.isalnum() or c in " _-").strip() or "Report"
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
            self.send_header("Content-Disposition", f'attachment; filename="{safe_name}_Report.pptx"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as e:  # noqa: BLE001 - surface any failure as JSON, not a bare 500
            self._send_json(500, {"error": f"{type(e).__name__}: {e}"})

    def _send_json(self, status, obj):
        payload = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)
