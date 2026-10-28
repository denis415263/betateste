// Vívora — reembolsos da Loja Virtual (Natueu), lidos da lista "Reembolso" do ClickUp.
//
// Roda como função do Vercel (rota /api/reembolsos), só leitura. O token do ClickUp fica
// na variável de ambiente CLICKUP_TOKEN e nunca vai para o navegador.
//
// Regras:
//   - só tarefas com Canal de Venda = "Loja VIrtual" (assim mesmo, com o "I" maiúsculo
//     no ClickUp — a comparação ignora maiúsculas/acentos/espaços)
//   - status do tipo "closed" ("reembolso concluído") = reembolso feito; data = date_closed
//   - demais status ("aguardando", "divergência") = pendente

const { allowOrigin, sendJson } = require("./_google");

const LIST_ID = "901325395468";
const FIELDS = {
  pedido:           "03247528-59fc-447b-90ec-33fb777abd02",
  reembolso:        "15d1dd9a-d643-4502-abad-e8ed05b4e12e",
  prejuizoFrete:    "bd01c3c8-b26e-4f9a-bebb-c3c69ce6c3cd",
  prejuizoProduto:  "caeace7a-6654-4313-9ac9-9a7db74eefbb",
  motivo:           "a741af42-512a-4ca9-9115-828b1d68e829",
  responsabilidade: "a9e1b7bb-b0dc-438c-bea0-6d88a8efd50c",
  canal:            "6fb4ccc7-c40a-4874-8118-8c8ac735be7c",
  sku1:             "88cf2a51-200f-4e75-a833-2c6be92f28fe",
  sku2:             "ed1c52ad-87af-4c38-92cd-1388ce5fac18",
  nomeCliente:      "6b158228-60b0-497f-bc13-68dca57bf4eb",
};
const CANAL_LV = "LOJAVIRTUAL";

function norm(s) {
  return String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/[\s_.\-]+/g, "");
}

function money(v) {
  if (v === null || v === undefined || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(",", "."));
  return isFinite(n) ? n : 0;
}

// Lista suspensa: a API manda o orderindex (campos antigos) ou o id da opção (campos novos)
function dropdown(field) {
  if (!field || field.value === null || field.value === undefined || field.value === "") return null;
  const v = field.value;
  if (typeof v === "object") return { name: v.name || null, color: v.color || null };
  const options = (field.type_config && field.type_config.options) || [];
  const opt = options.find((o) => o.id === v) || options.find((o) => String(o.orderindex) === String(v));
  return opt ? { name: opt.name, color: opt.color || null } : null;
}

function parseTask(t) {
  const byId = {};
  for (const f of t.custom_fields || []) byId[f.id] = f;
  const text = (k) => {
    const f = byId[FIELDS[k]];
    const v = f && f.value;
    return v === null || v === undefined ? "" : String(v).trim();
  };
  const canal = dropdown(byId[FIELDS.canal]);
  if (!canal || norm(canal.name) !== CANAL_LV) return null;

  // Nome da tarefa = cliente; tarefas vindas de formulário usam o campo "Nome Cliente"
  let cliente = String(t.name || "").trim();
  if (/^form submission/i.test(cliente) && text("nomeCliente")) cliente = text("nomeCliente");

  const closed = (t.status && t.status.type) === "closed";
  const prejuizoFrete = money(byId[FIELDS.prejuizoFrete] && byId[FIELDS.prejuizoFrete].value);
  const prejuizoProduto = money(byId[FIELDS.prejuizoProduto] && byId[FIELDS.prejuizoProduto].value);
  return {
    id: t.id,
    url: t.url,
    cliente,
    pedido: text("pedido"),
    sku1: text("sku1"),
    sku2: text("sku2"),
    motivo: dropdown(byId[FIELDS.motivo]),
    responsabilidade: dropdown(byId[FIELDS.responsabilidade]),
    reembolso: money(byId[FIELDS.reembolso] && byId[FIELDS.reembolso].value),
    prejuizoFrete,
    prejuizoProduto,
    prejuizo: prejuizoFrete + prejuizoProduto,
    status: t.status ? t.status.status : "",
    statusColor: t.status ? t.status.color : null,
    closed,
    dateClosed: closed && t.date_closed ? Number(t.date_closed) : null,
    dateCreated: t.date_created ? Number(t.date_created) : null,
  };
}

async function fetchAllTasks(token) {
  const tasks = [];
  for (let page = 0; page < 50; page++) {
    const url = `https://api.clickup.com/api/v2/list/${LIST_ID}/task?` + new URLSearchParams({
      include_closed: "true", subtasks: "true", page: String(page),
    });
    const res = await fetch(url, { headers: { Authorization: token } });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error("ClickUp: " + (data.err || data.error || `HTTP ${res.status}`));
    const batch = data.tasks || [];
    tasks.push(...batch);
    if (data.last_page || batch.length === 0) break;
  }
  return tasks;
}

async function handler(req, res) {
  if (!allowOrigin(req, res)) return;
  try {
    const token = process.env.CLICKUP_TOKEN;
    if (!token) throw new Error("Variável CLICKUP_TOKEN não configurada no Vercel.");
    const items = (await fetchAllTasks(token)).map(parseTask).filter(Boolean);
    sendJson(res, 200, { generatedAt: new Date().toISOString(), items }, "private, max-age=300");
  } catch (e) {
    sendJson(res, 500, { error: String(e.message || e) });
  }
}

module.exports = handler;
module.exports.parseTask = parseTask;
