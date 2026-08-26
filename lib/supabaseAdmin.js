// Server-only Supabase client, using the SERVICE ROLE key. Import this ONLY
// from api/*.js (Vercel serverless functions) or scripts/*.mjs — never from
// anything under src/, which is bundled into the browser. The service role
// key bypasses Row Level Security, which is exactly why it must never reach
// the client.
import { createClient } from "@supabase/supabase-js";

let cached = null;

export function getSupabaseAdmin() {
  if (cached) return cached;

  const url = process.env.SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!url || !key) {
    throw new Error(
      "SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY must be set (server-side environment variables — see .env.example)."
    );
  }
  cached = createClient(url, key, {
    auth: { persistSession: false, autoRefreshToken: false },
  });
  return cached;
}

export const MIS_BUCKET = "mis-files";
