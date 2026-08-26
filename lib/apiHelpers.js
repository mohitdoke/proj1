// Tiny shared helpers for api/*.js route handlers — just consistent JSON
// response shapes and method/error handling, nothing framework-specific.
export function sendJson(res, status, body) {
  res.status(status).setHeader("Content-Type", "application/json").end(JSON.stringify(body));
}

export function methodNotAllowed(res, allowed) {
  res.setHeader("Allow", allowed.join(", "));
  sendJson(res, 405, { ok: false, error: `Method not allowed. Use ${allowed.join(" or ")}.` });
}

/** Wrap a route handler so any thrown error becomes a clean 500 JSON body
 * instead of a raw Vercel crash/stack trace, and so ok:false results from
 * lib/dashboardRead.js consistently map to 404. */
export function withErrorHandling(fn) {
  return async (req, res) => {
    try {
      await fn(req, res);
    } catch (err) {
      console.error(err);
      sendJson(res, 500, { ok: false, error: String(err.message || err) });
    }
  };
}

/**
 * Manager-only gate for the upload route: requires a shared secret
 * (ADMIN_UPLOAD_TOKEN, set server-side only — see .env.example) in either
 * an `Authorization: Bearer <token>` or `x-admin-token` header. This is a
 * deliberately simple starting point so upload isn't a wide-open public
 * endpoint; see README "Manager authentication" for how to swap this for
 * real Supabase Auth (magic-link accounts restricted to manager emails)
 * without changing anything else in the upload pipeline. Returns true and
 * writes a 401/500 response itself if the check fails, so callers can just
 * `if (!requireAdmin(req, res)) return;`.
 */
export function requireAdmin(req, res) {
  const expected = process.env.ADMIN_UPLOAD_TOKEN;
  if (!expected) {
    sendJson(res, 500, { ok: false, error: "ADMIN_UPLOAD_TOKEN is not configured on the server — upload is disabled until it is set." });
    return false;
  }
  const header = req.headers["authorization"] || "";
  const bearer = header.startsWith("Bearer ") ? header.slice(7) : null;
  const provided = bearer || req.headers["x-admin-token"];
  if (provided !== expected) {
    sendJson(res, 401, { ok: false, error: "Missing or invalid admin credentials for this upload endpoint." });
    return false;
  }
  return true;
}
