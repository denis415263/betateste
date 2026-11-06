// Vívora — exige usuário logado e autorizado nas funções /api/*.
//
// O navegador manda o token da sessão do Supabase em "Authorization: Bearer ...". A consulta
// abaixo roda com esse token: o Supabase valida a assinatura e a validade, e o RLS de
// vivo_usuarios só devolve a linha do próprio usuário. Linha encontrada = autorizado.
// Qualquer falha nega o acesso. A URL e a chave anon são públicas (já estão no site).

const { sendJson } = require("./_google");

const SUPABASE_URL = process.env.SUPABASE_URL || "https://rartvdhsnvrnsczuleqq.supabase.co";
const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY || "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJhcnR2ZGhzbnZybnNjenVsZXFxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODcxNDg5ODgsImV4cCI6MjEwMjcyNDk4OH0.BCAxD7WHwk_NbbwhFZVJ3RgVS_Zr8O7UBH3oT6aWD4E";

async function isAuthorized(token) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/vivo_usuarios?select=user_id&limit=1`, {
    headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${token}` },
  });
  if (!res.ok) return false;
  const rows = await res.json();
  return Array.isArray(rows) && rows.length > 0;
}

// Devolve true se pode seguir; senão já respondeu 401.
async function requireUser(req, res) {
  const m = /^Bearer\s+(.+)$/i.exec(req.headers.authorization || "");
  let ok = false;
  if (m) {
    try { ok = await isAuthorized(m[1]); } catch (e) { ok = false; }
  }
  if (!ok) sendJson(res, 401, { error: "Faça login com um usuário autorizado." });
  return ok;
}

module.exports = { requireUser };
