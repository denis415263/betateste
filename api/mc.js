// Vívora — leitura da Margem de Contribuição (MC) nas planilhas mensais do Google Drive.
//
// Roda como função do Vercel (rota /api/mc). A autenticação no Google (conta de
// serviço, só leitura) fica em _google.js.
//
// Para cada planilha "AAAA.MM MARGEM CONTRIBUIÇÃO" da pasta:
//   - aba "Dados"        → MC por marca (fabricante), por canal + total oficial
//   - aba "Dados Itens"  → MC por SKU, por canal
// E da planilha mais recente que tiver a aba "Carteira": fornecedores de cada carteira
// comercial (A e B; a lista AB, compartilhada, fica fora das duas no app).
// As colunas são localizadas pelo NOME do cabeçalho (a posição muda entre os meses).

const { getAccessToken, gget, allowOrigin, sendJson } = require("./_google");

const DEFAULT_FOLDER_ID = "1zOwKTPabTqqIIE6_Wu18cGD1hP2UCpPC";

// Canais conhecidos — nomes diferentes entre meses apontam para o mesmo canal
const CHANNELS = [
  { key: "shopee_natuway", label: "Shopee Natuway", match: ["SHOPEENATUWAY"] },
  { key: "shopee_natueu",  label: "Shopee Natueu",  match: ["SHOPEENATUEU"] },
  { key: "ml_weley",       label: "ML Weley",       match: ["MLBARUKWELEY", "MLBARUK", "WELEY", "MLNATUWAY"] },
  { key: "ml_natubr",      label: "ML Natubr",      match: ["MLNATUBR", "MLNATUEU"] },
  { key: "loja",           label: "Loja (Tray)",    match: ["TRAY", "LV"] },
  { key: "distribuicao",   label: "Distribuição",   match: ["DISTRIBUICAO"] },
  { key: "amazon",         label: "Amazon",         match: ["AMAZON"] },
];

// ── utilidades ──────────────────────────────────────────────────────────────
function norm(s) {
  return String(s ?? "")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[\s_.\-]+/g, "");
}

// Aceita número puro (UNFORMATTED_VALUE) ou texto pt-BR ("R$ 1.234,56", "-12,5%")
function num(v) {
  if (typeof v === "number") return isFinite(v) ? v : 0;
  if (v === null || v === undefined) return 0;
  let s = String(v).replace(/ /g, " ").trim();
  if (!s || s.startsWith("#")) return 0;
  const isPct = s.includes("%");
  s = s.replace(/R\$/g, "").replace(/%/g, "").replace(/\s/g, "");
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  if (!isFinite(n)) return 0;
  return isPct ? n / 100 : n;
}

function round2(n) { return Math.round(n * 100) / 100; }

function channelOf(label) {
  const n = norm(label);
  for (const c of CHANNELS) if (c.match.some((m) => n === m || n.startsWith(m))) return c;
  return null;
}

// ── leitura de uma aba (marca ou item) ──────────────────────────────────────
// kind: "brand" (aba Dados) | "item" (aba Dados Itens)
function parseSheet(rows, kind) {
  rows = rows || [];
  const isHeader = (r) => {
    const a = norm(r && r[0]);
    return kind === "brand" ? a === "FABRICANTE" : a.startsWith("CODINTERNO");
  };
  const headerIdx = rows.findIndex(isHeader);
  if (headerIdx < 0) throw new Error(`cabeçalho não encontrado na aba ${kind === "brand" ? "Dados" : "Dados Itens"}`);
  const header = rows[headerIdx].map(norm);

  // Linha com os nomes dos canais = a que contém "TOTAL", acima do cabeçalho
  let chanIdx = -1;
  for (let i = 0; i < headerIdx; i++) {
    if ((rows[i] || []).some((c) => norm(c) === "TOTAL")) { chanIdx = i; break; }
  }
  if (chanIdx < 0) throw new Error("linha de canais (com TOTAL) não encontrada");
  const chanRow = rows[chanIdx];
  const statusRow = rows.slice(0, headerIdx).find((r) => norm(r && r[0]) === "STATUS") || null;

  // Blocos de canal: cada coluna "FAT" do cabeçalho começa um bloco. O nome do canal
  // fica em qualquer célula da linha de canais dentro do bloco (às vezes deslocado).
  const totalCol = chanRow.findIndex((c) => norm(c) === "TOTAL");
  const fatCols = [];
  header.forEach((h, i) => { if (h === "FAT") fatCols.push(i); });
  const nextFat = (i) => { const n = fatCols.find((f) => f > i); return n === undefined ? header.length : n; };
  const isText = (c) => typeof c === "string" && /[A-Za-z]/.test(c);

  const blocks = fatCols
    .filter((s) => s < totalCol) // depois do TOTAL vêm agrupamentos (NATUWAY / NATUEU) — ignorados
    .map((s) => {
      const end = Math.min(nextFat(s), totalCol);
      let label = "";
      for (let i = s; i < end; i++) if (isText(chanRow[i])) { label = chanRow[i]; break; }
      return { start: s, end, label };
    });

  const findCol = (b, names) => {
    for (let i = b.start; i < b.end; i++) if (names.includes(header[i])) return i;
    return -1;
  };

  // Total oficial (aba Dados): coluna "MC$" do bloco TOTAL
  const total = { mc: findCol({ start: totalCol, end: nextFat(totalCol) }, ["MC$"]) };

  const channels = [];
  for (const b of blocks) {
    const ch = channelOf(b.label);
    if (!ch) continue;
    let status = "";
    if (statusRow) for (let i = b.start; i < b.end; i++) if (statusRow[i]) { status = norm(statusRow[i]); break; }
    channels.push({
      key: ch.key,
      fat: findCol(b, ["FAT"]),
      mc: findCol(b, ["MC$"]),
      un: findCol(b, ["UN", "U"]),
      ads: findCol(b, ["ADS"]),
      emAndamento: status.includes("ANDAMENTO"),
    });
  }

  const out = {};
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const idRaw = row[0];
    if (idRaw === null || idRaw === undefined || String(idRaw).trim() === "") continue;
    const id = String(idRaw).trim();
    if (norm(id) === "TOTAL") continue;
    const rec = { fat: 0, mc: 0, un: 0, ch: {} };
    for (const c of channels) {
      const fat = c.fat >= 0 ? num(row[c.fat]) : 0;
      const mc = c.mc >= 0 ? num(row[c.mc]) : 0;
      const un = c.un >= 0 ? num(row[c.un]) : 0;
      if (!fat && !mc && !un) continue;
      rec.ch[c.key] = kind === "item" ? [round2(fat), round2(mc), un] : [round2(fat), round2(mc)];
      rec.fat += fat; rec.mc += mc; rec.un += un;
    }
    // Marca: MC total oficial vem da coluna MC$ do bloco TOTAL (já desconta devoluções)
    if (kind === "brand" && total && total.mc >= 0) rec.mc = num(row[total.mc]);
    if (!rec.fat && !rec.mc) continue;
    rec.fat = round2(rec.fat); rec.mc = round2(rec.mc);
    if (kind === "brand") { delete rec.un; rec.name = id; }
    out[kind === "brand" ? norm(id) : id.toUpperCase()] = rec;
  }
  return { rows: out, emAndamento: channels.some((c) => c.emAndamento), canaisEmAndamento: channels.filter((c) => c.emAndamento).map((c) => c.key), channels: channels.map((c) => c.key) };
}

function parseMonth(dadosRows, itensRows) {
  const brands = parseSheet(dadosRows, "brand");
  const items = parseSheet(itensRows, "item");
  return {
    emAndamento: brands.emAndamento || items.emAndamento,
    canaisEmAndamento: Array.from(new Set([...brands.canaisEmAndamento, ...items.canaisEmAndamento])),
    channels: Array.from(new Set([...brands.channels, ...items.channels])),
    brands: brands.rows,
    items: items.rows,
  };
}

// ── aba "Carteira" ──────────────────────────────────────────────────────────
// Linha 1 tem os títulos "Carteira A", "Carteira B" e "AB"; abaixo de cada um, os
// fornecedores (a coluna ao lado traz a MC, ignorada). A lista AB termina na primeira
// célula vazia — mais abaixo, na mesma coluna, vêm totais.
function parseCarteira(rows) {
  rows = rows || [];
  const head = rows[0] || [];
  const colOf = (label) => head.findIndex((c) => norm(c) === label);
  const cols = { A: colOf("CARTEIRAA"), B: colOf("CARTEIRAB"), AB: colOf("AB") };
  if (cols.A < 0 && cols.B < 0) throw new Error("títulos Carteira A / Carteira B não encontrados");
  const out = { A: [], B: [], AB: [] };
  for (const [k, c] of Object.entries(cols)) {
    if (c < 0) continue;
    for (let r = 1; r < rows.length; r++) {
      const v = String((rows[r] || [])[c] ?? "").trim();
      if (!v) { if (k === "AB") break; continue; }
      out[k].push(v);
    }
  }
  return out;
}

// ── Google: leitura ─────────────────────────────────────────────────────────
async function listMonthFiles(folderId, token) {
  const q = `'${folderId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`;
  const url = "https://www.googleapis.com/drive/v3/files?" + new URLSearchParams({
    q, fields: "files(id,name,modifiedTime)", pageSize: "200",
    supportsAllDrives: "true", includeItemsFromAllDrives: "true",
  });
  const { files = [] } = await gget(url, token);
  // Só "AAAA.MM MARGEM CONTRIBUIÇÃO" — ignora quinzenas (AAAA.MM.DD) e cópias
  const byMonth = {};
  for (const f of files) {
    const m = String(f.name).trim().match(/^(\d{4})\.(\d{2})\s+MARGEM\s+CONTRIBUI/i);
    if (!m) continue;
    const key = `${m[1]}-${m[2]}`;
    if (!byMonth[key] || f.modifiedTime > byMonth[key].modifiedTime) byMonth[key] = { ...f, key };
  }
  return Object.values(byMonth).sort((a, b) => a.key.localeCompare(b.key));
}

async function readMonth(file, token) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${file.id}/values:batchGet?` +
    new URLSearchParams([
      ["ranges", "Dados!A1:AZ300"],
      ["ranges", "'Dados Itens'!A1:AZ5000"],
      ["valueRenderOption", "UNFORMATTED_VALUE"],
    ]);
  const data = await gget(url, token);
  const [dados, itens] = data.valueRanges || [];
  return parseMonth((dados && dados.values) || [], (itens && itens.values) || []);
}

// Cache por instância: só relê o mês cuja planilha mudou
const monthCache = {};

// Carteiras: tenta as 3 planilhas mais recentes até achar uma aba "Carteira" preenchida
async function readCarteira(files, token) {
  for (const f of files.slice(-3).reverse()) {
    const ck = `carteira:${f.id}:${f.modifiedTime}`;
    if (monthCache[ck] === undefined) {
      try {
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${f.id}/values/${encodeURIComponent("Carteira!A1:K300")}`;
        const c = parseCarteira((await gget(url, token)).values);
        monthCache[ck] = c.A.length + c.B.length + c.AB.length ? { month: f.key, fileName: f.name, ...c } : null;
      } catch (e) {
        monthCache[ck] = null; // mês sem a aba (ou fora do padrão) — tenta o anterior
      }
    }
    if (monthCache[ck]) return monthCache[ck];
  }
  return null;
}

async function handler(req, res) {
  if (!allowOrigin(req, res)) return;

  try {
    const token = await getAccessToken();
    const files = await listMonthFiles(process.env.MC_FOLDER_ID || DEFAULT_FOLDER_ID, token);
    const nowKey = new Date().toISOString().slice(0, 7);
    const errors = [];
    const carteirasPromise = readCarteira(files, token);
    const months = await Promise.all(files.map(async (f) => {
      const ck = `${f.id}:${f.modifiedTime}`;
      try {
        if (!monthCache[ck]) monthCache[ck] = await readMonth(f, token);
        const m = monthCache[ck];
        const totalFat = Object.values(m.brands).reduce((a, b) => a + b.fat, 0) + Object.values(m.items).reduce((a, b) => a + b.fat, 0);
        if (!totalFat) return null; // planilha do mês ainda vazia
        return { key: f.key, fileName: f.name, modifiedTime: f.modifiedTime, provisional: m.emAndamento || f.key >= nowKey, canaisEmAndamento: m.canaisEmAndamento, channels: m.channels, brands: m.brands, items: m.items };
      } catch (e) {
        errors.push({ month: f.key, fileName: f.name, error: String(e.message || e) });
        return null;
      }
    }));
    const body = {
      generatedAt: new Date().toISOString(),
      channelLabels: Object.fromEntries(CHANNELS.map((c) => [c.key, c.label])),
      months: months.filter(Boolean),
      carteiras: await carteirasPromise,
      errors,
    };
    sendJson(res, 200, body, "private, max-age=600");
  } catch (e) {
    sendJson(res, 500, { error: String(e.message || e) });
  }
}

module.exports = handler;
module.exports.parseMonth = parseMonth;
module.exports.parseSheet = parseSheet;
module.exports.parseCarteira = parseCarteira;
