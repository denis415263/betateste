// Vívora — leitura da tabela de produtos no Google Sheets ("NOVA TABELA PRODUTOS 2.0").
//
// Roda como função do Vercel (rota /api/produtos), só leitura. Substitui o upload da
// Pasta1.xlsx na tela "Tabela de preços": mesma aba, mesmas colunas, mas localizadas
// pelo NOME do cabeçalho (não pela posição).
//
// Variáveis de ambiente opcionais: PRODUTOS_SHEET_ID e PRODUTOS_SHEET_TAB.
// A planilha precisa estar compartilhada (leitor) com o e-mail da conta de serviço.

const { getAccessToken, gget, allowOrigin, sendJson } = require("./_google");

const DEFAULT_SHEET_ID = "1zAh_1cNKSEUdTtbGijETG3EQNAqPCxLZT3vJkN7JI8Q";
const DEFAULT_TAB = "Tabela Regime Normal - Natueu";

// Campo → nomes aceitos no cabeçalho (já normalizados)
const COLUMNS = {
  sku:           ["REF"],
  codFabricante: ["CODFABRICANTE"],
  precoCusto:    ["CUSTOTOTAL+IPI", "CUSTOTOTALIPI"],   // custo de nota (antiga coluna K)
  custoLiquido:  ["TOTALCUSTOATON", "CUSTOATON"],      // custo líquido ATON (antiga coluna S)
  fabricante:    ["FABRICANTE"],
};

function norm(s) {
  return String(s ?? "")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[\s_.\-]+/g, "");
}

// Aceita número puro (UNFORMATTED_VALUE) ou texto pt-BR ("R$ 1.234,56")
function num(v) {
  if (typeof v === "number") return isFinite(v) ? v : null;
  let s = String(v ?? "").replace(/R\$/g, "").replace(/\s/g, "").trim();
  if (!s || s.startsWith("#")) return null;
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isFinite(n) ? n : null;
}

function parseProdutos(values) {
  const rows = values || [];
  const headerIdx = rows.findIndex((r) => (r || []).some((c) => norm(c) === "REF"));
  if (headerIdx < 0) throw new Error("cabeçalho com a coluna REF não encontrado na aba");
  const header = rows[headerIdx].map(norm);
  const col = {};
  for (const [field, names] of Object.entries(COLUMNS)) col[field] = header.findIndex((h) => names.includes(h));
  if (col.sku < 0) throw new Error("coluna REF não encontrada");

  const out = {};
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const sku = String(row[col.sku] ?? "").trim();
    if (!sku) continue;
    const rec = {};
    const pc = col.precoCusto >= 0 ? num(row[col.precoCusto]) : null;
    const cl = col.custoLiquido >= 0 ? num(row[col.custoLiquido]) : null;
    if (pc !== null) rec.precoCusto = Math.round(pc * 10000) / 10000;
    if (cl !== null) rec.custoLiquido = Math.round(cl * 10000) / 10000;
    if (col.codFabricante >= 0 && String(row[col.codFabricante] ?? "").trim()) rec.codFabricante = String(row[col.codFabricante]).trim();
    if (col.fabricante >= 0 && String(row[col.fabricante] ?? "").trim()) rec.fabricante = String(row[col.fabricante]).trim();
    out[sku] = rec;
  }
  return { rows: out, missing: Object.keys(COLUMNS).filter((k) => col[k] < 0) };
}

async function handler(req, res) {
  if (!allowOrigin(req, res)) return;
  try {
    const token = await getAccessToken();
    const id = process.env.PRODUTOS_SHEET_ID || DEFAULT_SHEET_ID;
    const tab = process.env.PRODUTOS_SHEET_TAB || DEFAULT_TAB;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(`'${tab}'!A1:AZ5000`)}?valueRenderOption=UNFORMATTED_VALUE`;
    const data = await gget(url, token);
    const parsed = parseProdutos(data.values);
    sendJson(res, 200, {
      generatedAt: new Date().toISOString(),
      sheetTab: tab,
      total: Object.keys(parsed.rows).length,
      missingColumns: parsed.missing,
      rows: parsed.rows,
    }, "no-store");
  } catch (e) {
    sendJson(res, 500, { error: String(e.message || e) });
  }
}

module.exports = handler;
module.exports.parseProdutos = parseProdutos;
