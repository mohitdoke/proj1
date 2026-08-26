// POST /api/admin/mis/upload — manager-only MIS upload endpoint.
// multipart/form-data with a single file field ("file") and optional
// "companySlug" field (a manager's dropdown choice, cross-checked against
// what's actually detected from the sheet — see lib/misProcessing.js).
//
// This route is deliberately thin: all the real work (parse, strict-detect,
// validate, store file + normalized rows, atomic version swap) lives in
// lib/misProcessing.js's processMisUpload(), the exact same function
// scripts/seed.mjs uses for the initial data import — one implementation
// for both paths, per the master spec.
import formidable from "formidable";
import fs from "node:fs/promises";
import { processMisUpload } from "../../../lib/misProcessing.js";
import { sendJson, methodNotAllowed, withErrorHandling, requireAdmin } from "../../../lib/apiHelpers.js";

// Vercel Node functions: disable the default body parser so formidable can
// read the raw multipart stream itself.
export const config = { api: { bodyParser: false } };

function parseMultipart(req) {
  return new Promise((resolve, reject) => {
    const form = formidable({ maxFileSize: 25 * 1024 * 1024, multiples: false });
    form.parse(req, (err, fields, files) => {
      if (err) reject(err);
      else resolve({ fields, files });
    });
  });
}

export default withErrorHandling(async (req, res) => {
  if (req.method !== "POST") return methodNotAllowed(res, ["POST"]);
  if (!requireAdmin(req, res)) return;

  const { fields, files } = await parseMultipart(req);
  const uploaded = Array.isArray(files.file) ? files.file[0] : files.file;
  if (!uploaded) {
    return sendJson(res, 400, { ok: false, error: 'No file received — send it as multipart/form-data with field name "file".' });
  }

  const buffer = await fs.readFile(uploaded.filepath);
  const originalFilename = uploaded.originalFilename || "upload.xlsx";
  const companySlugHint = Array.isArray(fields.companySlug) ? fields.companySlug[0] : fields.companySlug || undefined;
  const uploadedBy = Array.isArray(fields.uploadedBy) ? fields.uploadedBy[0] : fields.uploadedBy || undefined;

  const result = await processMisUpload({ buffer, originalFilename, uploadedBy, companySlugHint });
  await fs.unlink(uploaded.filepath).catch(() => {});

  sendJson(res, result.ok ? 200 : 422, result);
});
