// Vívora — acesso somente leitura ao Google (Drive + Sheets) com a conta de serviço.
// Compartilhado pelas funções /api/mc e /api/produtos. Arquivos com "_" no início
// não viram rota no Vercel.
//
// A conta de serviço fica na variável de ambiente GOOGLE_SERVICE_ACCOUNT_JSON e só
// pede escopos de leitura — mesmo com permissão maior por engano, não altera nada.

const crypto = require("crypto");

const SCOPES = [
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/spreadsheets.readonly",
].join(" ");

function b64url(input) {
  return Buffer.from(input).toString("base64").replace(/=+$/, "").replace(/\+/g, "-").replace(/\//g, "_");
}

let tokenCache = { token: null, exp: 0 };
async function getAccessToken() {
  if (tokenCache.token && Date.now() < tokenCache.exp - 60000) return tokenCache.token;
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!raw) throw new Error("Variável GOOGLE_SERVICE_ACCOUNT_JSON não configurada no Vercel.");
  const sa = JSON.parse(raw);
  const key = String(sa.private_key || "").replace(/\\n/g, "\n");
  const now = Math.floor(Date.now() / 1000);
  const head = b64url(JSON.stringify({ alg: "RS256", typ: "JWT" }));
  const claim = b64url(JSON.stringify({ iss: sa.client_email, scope: SCOPES, aud: "https://oauth2.googleapis.com/token", iat: now, exp: now + 3600 }));
  const sig = crypto.createSign("RSA-SHA256").update(`${head}.${claim}`).sign(key);
  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: `${head}.${claim}.${b64url(sig)}` }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error("Falha ao autenticar no Google: " + (data.error_description || data.error || res.status));
  tokenCache = { token: data.access_token, exp: Date.now() + (data.expires_in || 3600) * 1000 };
  return tokenCache.token;
}

async function gget(url, token) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error((data.error && data.error.message) || `HTTP ${res.status}`);
  return data;
}

// Só responde para o próprio site (bloqueia acesso direto pelo endereço).
// Devolve false (e já responde 403) quando a origem não confere.
function allowOrigin(req, res) {
  const host = req.headers.host || "";
  const ref = req.headers.origin || req.headers.referer || "";
  let refHost = "";
  try { refHost = new URL(ref).host; } catch (e) { /* sem referer */ }
  if (refHost && refHost === host) return true;
  sendJson(res, 403, { error: "Acesso não permitido." });
  return false;
}

function sendJson(res, status, body, cacheControl) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  if (cacheControl) res.setHeader("Cache-Control", cacheControl);
  res.end(JSON.stringify(body));
}

module.exports = { getAccessToken, gget, allowOrigin, sendJson };
