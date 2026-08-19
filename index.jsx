import React, { useState, useEffect, useMemo, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  Upload, Package, ShoppingCart, Search,
  ChevronDown, ChevronRight, AlertTriangle, CheckCircle2,
  Settings, FileSpreadsheet, History, Trash2, Box, TrendingUp, TrendingDown, Minus
} from "lucide-react";

// ---------- Storage keys ----------
const K_PRODUCTS = "vivo:products";       // map by codigo_interno -> product record
const K_HISTORY = "vivo:history";         // array of import events (lightweight log)
const K_SETTINGS = "vivo:settings";       // global settings (default coverage days etc)
const K_ORDERS = "vivo:orders";           // saved purchase orders

// ---------- Helpers ----------
function normalizeHeader(h) {
  return String(h || "")
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_|_$/g, "");
}

const HEADER_MAP = {
  fornecedor: ["fornecedor", "fabricante", "fornecedor_fabricante", "marca"],
  codigo: ["codigo_interno", "codigo", "cod_interno", "sku", "cod"],
  item: ["item", "produto", "descricao", "nome_produto", "nome"],
  estoque: ["estoque_atual", "estoque", "saldo_estoque", "qtd_estoque"],
  vendas30: ["vendas_30_dias", "vendas30", "qtd_vendida_30_dias", "vendas_ultimos_30_dias", "30_dias"],
  vendas90: ["vendas_90_dias", "vendas90", "qtd_vendida_90_dias", "vendas_ultimos_90_dias", "90_dias"],
  vendas180: ["vendas_180_dias", "vendas180", "qtd_vendida_180_dias", "vendas_ultimos_180_dias", "180_dias"],
  precoCusto: ["vlr_custo", "valor_custo", "preco_custo", "preco_de_custo", "custo_unitario", "custo_medio", "vlr_de_custo", "custo"],
};

function mapRowToFields(row) {
  const normRow = {};
  for (const key of Object.keys(row)) {
    normRow[normalizeHeader(key)] = row[key];
  }
  const out = {};
  for (const field of Object.keys(HEADER_MAP)) {
    for (const candidate of HEADER_MAP[field]) {
      if (normRow[candidate] !== undefined) {
        out[field] = normRow[candidate];
        break;
      }
    }
  }
  return out;
}

function toNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const cleaned = String(v).replace(/\./g, "").replace(",", ".").replace(/[^\d.-]/g, "");
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function fmtNumber(n) {
  return new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 }).format(Math.round(n || 0));
}

function fmtCurrency(n) {
  return new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(n || 0);
}

function fmtDate(ts) {
  return new Date(ts).toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit", year: "2-digit", hour: "2-digit", minute: "2-digit" });
}

function fmtDateShort(ts) {
  return new Date(ts).toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" });
}

function uid() {
  return Math.random().toString(36).slice(2, 10);
}

function resolveCoverageDays(product, settings) {
  if (product.coverageDaysOverride !== null && product.coverageDaysOverride !== undefined) {
    return product.coverageDaysOverride;
  }
  const supplierDays = settings.supplierCoverageDays?.[product.fornecedor];
  if (supplierDays !== undefined && supplierDays !== null) {
    return supplierDays;
  }
  return settings.defaultCoverageDays;
}

function coverageSource(product, settings) {
  if (product.coverageDaysOverride !== null && product.coverageDaysOverride !== undefined) return "produto";
  const supplierDays = settings.supplierCoverageDays?.[product.fornecedor];
  if (supplierDays !== undefined && supplierDays !== null) return "fornecedor";
  return "padrao";
}

const MS_PER_DAY = 1000 * 60 * 60 * 24;

// Retorna a chave de dia local (AAAA-MM-DD) de um timestamp, para agrupar
// importações por data calendário, não por hora exata da importação.
function dayKey(ts) {
  const d = new Date(ts);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

// Quando há mais de uma importação no mesmo dia (ex: erro na planilha, reenvio),
// mantém só a leitura mais recente daquele dia para fins de série histórica.
// As demais continuam guardadas em salesHistory (auditoria), só não entram no cálculo de tendência.
function dedupeByCalendarDay(points) {
  const byDay = {};
  for (const p of points) {
    const key = dayKey(p.ts);
    if (!byDay[key] || p.ts > byDay[key].ts) byDay[key] = p;
  }
  return Object.values(byDay).sort((a, b) => a.ts - b.ts);
}

// Para uma métrica (vendas30, vendas90 ou vendas180), retorna a série de intervalos
// entre DIAS consecutivos com dados (não entre importações — múltiplas importações
// no mesmo dia são reduzidas à mais recente daquele dia antes de calcular), com a
// estimativa de vendas diárias naquele intervalo: (delta de vendas) / (dias corridos).
// É uma estimativa, pois "vendas 30 dias" é uma janela móvel, não um contador zerado.
function buildIntervalSeries(salesHistory, metricKey) {
  const withMetric = (salesHistory || [])
    .filter((h) => h[metricKey] !== undefined && h[metricKey] !== null)
    .sort((a, b) => a.ts - b.ts);
  const points = dedupeByCalendarDay(withMetric);

  const intervals = [];
  for (let i = 1; i < points.length; i++) {
    const prev = points[i - 1];
    const curr = points[i];
    const days = (curr.ts - prev.ts) / MS_PER_DAY;
    if (days <= 0) continue; // mesmo dia (não deveria sobrar após o dedupe) ou relógio inconsistente; ignora
    const delta = curr[metricKey] - prev[metricKey];
    const estimatedDailyAvg = delta / days;
    intervals.push({
      from: prev.ts,
      to: curr.ts,
      days: Math.round(days * 10) / 10,
      delta,
      estimatedDailyAvg,
    });
  }
  return intervals;
}

// Retorna um resumo de tendência: compara a média diária do intervalo mais recente
// com a do intervalo anterior. "sem_dados" se não houver pelo menos 2 intervalos.
function getTrendSummary(salesHistory) {
  const series = buildIntervalSeries(salesHistory, "vendas30");
  if (series.length === 0) return { status: "sem_dados", lastAvg: null, prevAvg: null, deltaPct: null };
  const last = series[series.length - 1];
  if (series.length === 1) return { status: "primeira_leitura", lastAvg: last.estimatedDailyAvg, prevAvg: null, deltaPct: null };
  const prev = series[series.length - 2];
  const deltaPct = prev.estimatedDailyAvg > 0
    ? ((last.estimatedDailyAvg - prev.estimatedDailyAvg) / prev.estimatedDailyAvg) * 100
    : null;
  let status = "estavel";
  if (deltaPct !== null) {
    if (deltaPct > 8) status = "subindo";
    else if (deltaPct < -8) status = "caindo";
  }
  return { status, lastAvg: last.estimatedDailyAvg, prevAvg: prev.estimatedDailyAvg, deltaPct };
}

// Calcula tempo de estoque (em dias) e capital imobilizado de um produto.
// Regra: sem vendas nos últimos 30 dias = tempo de estoque "infinito" (Infinity),
// já que o item parado não tem previsão de giro pela métrica disponível.
const STOCK_QUALITY_THRESHOLD_DAYS = 90;

function getStockQuality(product) {
  const avgDay = (product.vendas30 || 0) / 30;
  const estoque = product.estoque || 0;
  const stockDays = avgDay > 0 ? estoque / avgDay : (estoque > 0 ? Infinity : 0);
  const capitalImobilizado = estoque * (product.precoCusto || 0);
  const isExcess = stockDays > STOCK_QUALITY_THRESHOLD_DAYS;
  return { avgDay, stockDays, capitalImobilizado, isExcess };
}

// Tolerância para considerar a tendência "Igual": variações pequenas (ruído) não
// devem aparecer como acelerando/caindo. Abaixo de 5% de variação = estável.
const MONITORING_STABLE_TOLERANCE_PCT = 5;

function getMonitoringStatus(product) {
  const startedAt = product.monitoringStartedAt;
  const startV30 = product.monitoringStartV30;
  const currentV30 = product.vendas30 ?? null;

  const days = startedAt ? Math.max(0, Math.round((Date.now() - startedAt) / MS_PER_DAY)) : null;

  if (startV30 === null || startV30 === undefined || currentV30 === null) {
    return { days, startV30, currentV30, status: "sem_dados", deltaPct: null };
  }

  let deltaPct;
  if (startV30 === 0) {
    deltaPct = currentV30 > 0 ? 100 : 0;
  } else {
    deltaPct = ((currentV30 - startV30) / startV30) * 100;
  }

  let status = "igual";
  if (deltaPct > MONITORING_STABLE_TOLERANCE_PCT) status = "acelerando";
  else if (deltaPct < -MONITORING_STABLE_TOLERANCE_PCT) status = "caindo";

  return { days, startV30, currentV30, status, deltaPct };
}

// Retorna os produtos cujo lastUpdated não coincide com a importação mais recente
// (ts do primeiro item do histórico, já que ele é guardado com o mais novo primeiro),
// ou seja, itens que não vieram em nenhuma das planilhas do último lote enviado.
function getMissingFromLastImport(products, history) {
  if (!history || history.length === 0) return { lastImportTs: null, missing: [] };
  const lastImportTs = history[0].ts;
  const missing = Object.values(products).filter(
    (p) => !p.inactive && p.lastUpdated !== lastImportTs
  );
  return { lastImportTs, missing };
}

const QUALITY_REASON_TAGS = ["Negociação Fornecedor", "Item Novo", "Ruptura ou Reposição", "Queda Vendas"];
const QUALITY_SITUATION_TAGS = ["Liquidação", "Ads", "Sem Ação"];
const QUALITY_MARGIN_OPTIONS = [15, 10, 5, 0, -5, -10, -15, -20, -25, -30];

// Define a classe de cor de cada select de classificação, conforme o valor escolhido,
// para dar leitura visual rápida na tabela densa de Qualidade de Estoque.
function tagColorClass(kind, value) {
  if (!value && value !== 0) return "";
  if (kind === "reason") {
    if (value === "Negociação Fornecedor") return " vivo-tag-blue";
    if (value === "Item Novo") return " vivo-tag-teal";
    if (value === "Ruptura ou Reposição") return " vivo-tag-amber";
    if (value === "Queda Vendas") return " vivo-tag-rust";
    return "";
  }
  if (kind === "situation") {
    if (value === "Liquidação") return " vivo-tag-rust";
    if (value === "Ads") return " vivo-tag-blue";
    if (value === "Sem Ação") return " vivo-tag-neutral";
    return "";
  }
  if (kind === "margin") {
    const n = Number(value);
    if (n >= 10) return " vivo-tag-olive";
    if (n >= 0) return " vivo-tag-amber";
    return " vivo-tag-rust";
  }
  return "";
}

// ---------- Supabase Storage ----------
// Dados guardados no Supabase — persistem entre publicações, atualizações de código
// e acessos de qualquer dispositivo. Substitui o window.storage do artifact.
const SUPA_URL = "https://rartvdhsnvrnsczuleqq.supabase.co";
const SUPA_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJhcnR2ZGhzbnZybnNjenVsZXFxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODcxNDg5ODgsImV4cCI6MjEwMjcyNDk4OH0.BCAxD7WHwk_NbbwhFZVJ3RgVS_Zr8O7UBH3oT6aWD4E";
const SUPA_TABLE = "vivo_storage";
const SUPA_HEADERS = {
  "Content-Type": "application/json",
  "apikey": SUPA_KEY,
  "Authorization": `Bearer ${SUPA_KEY}`,
  "Prefer": "resolution=merge-duplicates",
};

function withTimeout(promise, ms, fallback) {
  return new Promise((resolve) => {
    let done = false;
    const timer = setTimeout(() => {
      if (!done) { done = true; resolve(fallback); }
    }, ms);
    promise.then((v) => { if (!done) { done = true; clearTimeout(timer); resolve(v); } })
           .catch(() => { if (!done) { done = true; clearTimeout(timer); resolve(fallback); } });
  });
}

async function storageGet(key, fallback) {
  try {
    const res = await withTimeout(
      fetch(`${SUPA_URL}/rest/v1/${SUPA_TABLE}?key=eq.${encodeURIComponent(key)}&select=value`, {
        headers: SUPA_HEADERS,
      }),
      10000, null
    );
    if (!res || !res.ok) return fallback;
    const rows = await res.json();
    if (!rows || rows.length === 0) return fallback;
    return JSON.parse(rows[0].value);
  } catch (e) {
    console.error("storageGet error", key, e);
    return fallback;
  }
}

async function storageSet(key, value) {
  try {
    await withTimeout(
      fetch(`${SUPA_URL}/rest/v1/${SUPA_TABLE}`, {
        method: "POST",
        headers: SUPA_HEADERS,
        body: JSON.stringify({ key, value: JSON.stringify(value), updated_at: new Date().toISOString() }),
      }),
      10000, null
    );
  } catch (e) {
    console.error("storageSet error", key, e);
  }
}

// Mantida para compatibilidade com o painel de diagnóstico — agora não faz nada
// já que não há mais espaço privado vs compartilhado: tudo vai pro Supabase.
async function storageGetPrivate(key, fallback) {
  return fallback;
}

const APP_PASSWORD = "Natuweb123";
const APP_VERSION = "2026-06-25.2";

// Tela de senha simples — barreira contra acesso casual/acidental (ex: link
// encontrado em buscas), não uma proteção de segurança real, já que o código
// e a senha ficam visíveis a quem inspecionar o artifact.
function PasswordGate({ onSuccess }) {
  const [value, setValue] = useState("");
  const [error, setError] = useState(false);

  function checkPassword() {
    if (value.trim() === APP_PASSWORD) {
      setError(false);
      onSuccess();
    } else {
      setError(true);
    }
  }

  return (
    <div className="vivo-root vivo-loading">
      <style>{VIVO_CSS}</style>
      <div className="vivo-loading-mark">VIVO</div>
      <div className="vivo-password-form">
        <p className="vivo-password-label">Acesso restrito — digite a senha para continuar</p>
        <input
          type="password"
          className={"vivo-password-input" + (error ? " is-error" : "")}
          value={value}
          onChange={(e) => { setValue(e.target.value); setError(false); }}
          onKeyDown={(e) => { if (e.key === "Enter") checkPassword(); }}
          placeholder="Senha"
          autoFocus
        />
        {error && <p className="vivo-password-error">Senha incorreta. Tente novamente.</p>}
        <button type="button" className="vivo-btn vivo-btn-primary vivo-password-btn" onClick={checkPassword}>
          Entrar
        </button>
        <p className="vivo-password-version">versão {APP_VERSION}</p>
      </div>
    </div>
  );
}

// Aviso que aparece quando o espaço compartilhado está vazio mas há dados antigos
// no espaço privado desta conta — oferece migração manual com um clique.
function MigrationBanner({ migration, onMigrate }) {
  if (migration.done) {
    return (
      <div className="vivo-migration-banner vivo-migration-success">
        <CheckCircle2 size={16} />
        <span>Dados migrados com sucesso para o espaço compartilhado. Atualize a página em outros dispositivos para ver.</span>
      </div>
    );
  }
  return (
    <div className="vivo-migration-banner">
      <AlertTriangle size={16} />
      <span>
        Encontramos dados importados anteriormente que ainda não estão no espaço compartilhado
        (visível para todos com o link). Clique para copiá-los agora.
      </span>
      <button className="vivo-btn vivo-btn-primary" onClick={onMigrate} disabled={migration.running}>
        {migration.running ? "Migrando…" : "Migrar dados agora"}
      </button>
      {migration.error && <span className="vivo-migration-error">{migration.error}</span>}
    </div>
  );
}

export default function App() {
  const [products, setProducts] = useState({});
  const [history, setHistory] = useState([]);
  const [settings, setSettings] = useState({ defaultCoverageDays: 30, supplierCoverageDays: {} });
  const [orders, setOrders] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [mainTab, setMainTab] = useState("relatorios");
  const [subTab, setSubTab] = useState("import");
  const [authed, setAuthed] = useState(false);
  const [migration, setMigration] = useState({ checked: false, available: false, running: false, done: false, error: null });

  function goTo(mainId, subId) {
    setMainTab(mainId);
    setSubTab(subId);
  }

  // Load all data only after the password screen is passed
  useEffect(() => {
    if (!authed) return;
    let cancelled = false;
    (async () => {
      // Timeout de segurança: garante que a tela nunca fique travada em "carregando"
      // para sempre, mesmo se alguma chamada de storage não responder.
      const safetyTimer = setTimeout(() => {
        if (!cancelled) setLoaded(true);
      }, 10000);

      const [p, h, s, o] = await Promise.all([
        storageGet(K_PRODUCTS, {}),
        storageGet(K_HISTORY, []),
        storageGet(K_SETTINGS, { defaultCoverageDays: 30, supplierCoverageDays: {} }),
        storageGet(K_ORDERS, []),
      ]);
      if (cancelled) return;
      clearTimeout(safetyTimer);

      // Migração do campo renomeado: monitoringStartV90 → monitoringStartV30.
      // Produtos salvos antes desta mudança ainda têm o campo antigo — copia o valor
      // para o novo campo silenciosamente e persiste, sem perder nenhum dado.
      let needsMigration = false;
      const migratedProducts = { ...p };
      for (const codigo of Object.keys(migratedProducts)) {
        const prod = migratedProducts[codigo];
        if (prod.monitoringStartV90 !== undefined && prod.monitoringStartV30 === undefined) {
          migratedProducts[codigo] = { ...prod, monitoringStartV30: prod.monitoringStartV90 };
          needsMigration = true;
        }
      }
      if (needsMigration) {
        await withTimeout(window.storage.set(K_PRODUCTS, JSON.stringify(migratedProducts), true), 8000, null);
      }
      const finalProducts = needsMigration ? migratedProducts : p;

      setProducts(finalProducts);
      setHistory(h);
      setSettings({ defaultCoverageDays: 30, supplierCoverageDays: {}, ...s });
      setOrders(o);
      setLoaded(true);

      // Se o espaço compartilhado está vazio, checa se há dados antigos no espaço
      // privado desta conta (de antes da troca para armazenamento compartilhado).
      if (Object.keys(finalProducts).length === 0) {
        const privateProducts = await storageGetPrivate(K_PRODUCTS, {});
        if (!cancelled) {
          setMigration((m) => ({ ...m, checked: true, available: Object.keys(privateProducts).length > 0 }));
        }
      } else {
        setMigration((m) => ({ ...m, checked: true, available: false }));
      }
    })();
    return () => { cancelled = true; };
  }, [authed]);

  async function runMigration() {
    setMigration((m) => ({ ...m, running: true, error: null }));
    try {
      const [p, h, s, o] = await Promise.all([
        storageGetPrivate(K_PRODUCTS, {}),
        storageGetPrivate(K_HISTORY, []),
        storageGetPrivate(K_SETTINGS, { defaultCoverageDays: 30, supplierCoverageDays: {} }),
        storageGetPrivate(K_ORDERS, []),
      ]);
      await Promise.all([
        storageSet(K_PRODUCTS, p),
        storageSet(K_HISTORY, h),
        storageSet(K_SETTINGS, s),
        storageSet(K_ORDERS, o),
      ]);
      setProducts(p);
      setHistory(h);
      setSettings({ defaultCoverageDays: 30, supplierCoverageDays: {}, ...s });
      setOrders(o);
      setMigration((m) => ({ ...m, running: false, done: true, available: false }));
    } catch (e) {
      setMigration((m) => ({ ...m, running: false, error: "Não foi possível migrar os dados. Tente novamente." }));
    }
  }

  const persistProducts = useCallback(async (next) => {
    setProducts(next);
    await storageSet(K_PRODUCTS, next);
  }, []);
  const persistHistory = useCallback(async (next) => {
    setHistory(next);
    await storageSet(K_HISTORY, next);
  }, []);
  const persistSettings = useCallback(async (next) => {
    setSettings(next);
    await storageSet(K_SETTINGS, next);
  }, []);
  const persistOrders = useCallback(async (next) => {
    setOrders(next);
    await storageSet(K_ORDERS, next);
  }, []);

  if (!authed) {
    return <PasswordGate onSuccess={() => setAuthed(true)} />;
  }

  if (!loaded) {
    return (
      <div className="vivo-root vivo-loading">
        <style>{VIVO_CSS}</style>
        <div className="vivo-loading-mark">VIVO</div>
        <div className="vivo-loading-text">Carregando dados compartilhados…</div>
      </div>
    );
  }

  return (
    <div className="vivo-root">
      <style>{VIVO_CSS}</style>
      {migration.available && (
        <MigrationBanner migration={migration} onMigrate={runMigration} />
      )}
      <TopNav
        mainTab={mainTab}
        subTab={subTab}
        goTo={goTo}
        productCount={Object.values(products).filter((p) => !p.inactive).length}
        ordersCount={orders.length}
        monitoredCount={Object.values(products).filter((p) => p.monitoringActive).length}
      />
      <main className="vivo-main">
        {mainTab === "relatorios" && subTab === "import" && (
          <ImportTab
            products={products}
            persistProducts={persistProducts}
            history={history}
            persistHistory={persistHistory}
            goToProducts={() => goTo("compras", "products")}
          />
        )}
        {mainTab === "relatorios" && subTab === "history" && (
          <HistoryTab history={history} persistHistory={persistHistory} />
        )}
        {mainTab === "relatorios" && subTab === "cleanup" && (
          <CleanupTab
            products={products}
            persistProducts={persistProducts}
            history={history}
            settings={settings}
            orders={orders}
            persistHistory={persistHistory}
            persistSettings={persistSettings}
            persistOrders={persistOrders}
          />
        )}
        {mainTab === "compras" && subTab === "products" && (
          <ProductsTab products={products} persistProducts={persistProducts} settings={settings} />
        )}
        {mainTab === "compras" && subTab === "orders" && (
          <OrdersTab
            products={products}
            settings={settings}
            persistSettings={persistSettings}
            orders={orders}
            persistOrders={persistOrders}
          />
        )}
        {mainTab === "performance" && subTab === "tendencias" && (
          <PerformanceTab products={products} />
        )}
        {mainTab === "qualidade" && subTab === "analise" && (
          <StockQualityTab products={products} persistProducts={persistProducts} />
        )}
        {mainTab === "qualidade" && subTab === "monitorados" && (
          <MonitoredTab products={products} persistProducts={persistProducts} />
        )}
      </main>
    </div>
  );
}

// ====================== TOP NAV ======================
const NAV_STRUCTURE = [
  {
    id: "relatorios",
    label: "Relatórios",
    icon: Upload,
    subItems: [
      { id: "import", label: "Subir relatório" },
      { id: "history", label: "Histórico de importações" },
      { id: "cleanup", label: "Limpeza de itens" },
    ],
  },
  {
    id: "compras",
    label: "Compras",
    icon: ShoppingCart,
    subItems: [
      { id: "products", label: "Produtos" },
      { id: "orders", label: "Pedidos de compra" },
    ],
  },
  {
    id: "performance",
    label: "Performance",
    icon: TrendingUp,
    subItems: [
      { id: "tendencias", label: "Tendência de vendas" },
    ],
  },
  {
    id: "qualidade",
    label: "Qualidade de Estoque",
    icon: AlertTriangle,
    subItems: [
      { id: "analise", label: "Análise" },
      { id: "monitorados", label: "Monitorados" },
    ],
  },
];

function TopNav({ mainTab, subTab, goTo, productCount, ordersCount, monitoredCount }) {
  const countFor = (subId) => {
    if (subId === "products") return productCount;
    if (subId === "orders") return ordersCount;
    if (subId === "monitorados") return monitoredCount;
    return null;
  };

  const activeMain = NAV_STRUCTURE.find((m) => m.id === mainTab);

  return (
    <header className="vivo-topnav">
      <div className="vivo-topnav-row vivo-topnav-main">
        <div className="vivo-brand">
          <span className="vivo-brand-mark">●</span>
          <div>
            <div className="vivo-brand-name">Vívora</div>
            <div className="vivo-brand-sub">Compras &amp; Estoque</div>
          </div>
        </div>
        <nav className="vivo-main-tabs">
          {NAV_STRUCTURE.map((m) => (
            <button
              key={m.id}
              className={"vivo-main-tab-btn" + (mainTab === m.id ? " is-active" : "")}
              onClick={() => goTo(m.id, m.subItems[0].id)}
            >
              <m.icon size={16} strokeWidth={2} />
              <span>{m.label}</span>
            </button>
          ))}
        </nav>
        <div className="vivo-topnav-spacer" />
        <span className="vivo-topnav-hint">v{APP_VERSION} • Dados compartilhados</span>
      </div>

      {activeMain && (
        <div className="vivo-topnav-row vivo-sub-tabs">
          {activeMain.subItems.map((s) => {
            const count = countFor(s.id);
            return (
              <button
                key={s.id}
                className={"vivo-sub-tab-btn" + (subTab === s.id ? " is-active" : "")}
                onClick={() => goTo(mainTab, s.id)}
              >
                {s.label}
                {typeof count === "number" && count > 0 && (
                  <span className="vivo-sub-tab-count">{count}</span>
                )}
              </button>
            );
          })}
        </div>
      )}
    </header>
  );
}

// ====================== IMPORT TAB ======================
const PERIOD_LABELS = { 30: "30 dias", 90: "90 dias", 180: "180 dias" };

function detectPeriod(fileName, rows) {
  const name = fileName.toLowerCase();
  if (name.includes("180")) return 180;
  if (name.includes("90")) return 90;
  if (name.includes("30")) return 30;
  // fallback: check which vendas column actually has data in the headers
  if (rows && rows.length > 0) {
    const mapped = mapRowToFields(rows[0]);
    if (mapped.vendas180 !== undefined) return 180;
    if (mapped.vendas90 !== undefined) return 90;
    if (mapped.vendas30 !== undefined) return 30;
  }
  return 30;
}

function ImportTab({ products, persistProducts, history, persistHistory, goToProducts }) {
  const [pendingFiles, setPendingFiles] = useState([]); // [{id, fileName, period, rows, mapped, withSupplier, withoutSupplier}]
  const [error, setError] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [loading, setLoading] = useState(false);

  const readFile = useCallback((file) => {
    return new Promise((resolve, reject) => {
      const isCSV = /\.csv$/i.test(file.name);
      if (isCSV) {
        Papa.parse(file, {
          header: true,
          skipEmptyLines: true,
          complete: (res) => resolve(res.data),
          error: (err) => reject(err),
        });
      } else {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const wb = XLSX.read(e.target.result, { type: "array" });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            resolve(XLSX.utils.sheet_to_json(sheet, { defval: "" }));
          } catch (err) {
            reject(err);
          }
        };
        reader.onerror = () => reject(new Error("Não foi possível ler o arquivo."));
        reader.readAsArrayBuffer(file);
      }
    });
  }, []);

  const handleFiles = useCallback(async (fileList) => {
    setError("");
    setLoading(true);
    const files = Array.from(fileList).slice(0, 6); // safety cap
    try {
      const results = [];
      for (const file of files) {
        const rawRows = await readFile(file);
        if (!rawRows || rawRows.length === 0) {
          results.push({ id: uid(), fileName: file.name, period: 30, error: "Nenhuma linha encontrada." });
          continue;
        }
        const mapped = rawRows.map((r) => mapRowToFields(r)).filter((r) => r.codigo || r.item);
        const withSupplier = mapped.filter((r) => r.fornecedor && String(r.fornecedor).trim() !== "");
        const withoutSupplier = mapped.length - withSupplier.length;
        const period = detectPeriod(file.name, rawRows);
        results.push({
          id: uid(),
          fileName: file.name,
          period,
          total: mapped.length,
          withSupplier,
          withoutSupplierCount: withoutSupplier,
        });
      }
      setPendingFiles((prev) => [...prev, ...results]);
    } catch (err) {
      setError("Erro ao ler arquivo(s): " + (err.message || err));
    } finally {
      setLoading(false);
    }
  }, [readFile]);

  function onDrop(e) {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer.files?.length) handleFiles(e.dataTransfer.files);
  }

  function updatePeriod(id, period) {
    setPendingFiles((prev) => prev.map((f) => (f.id === id ? { ...f, period: Number(period) } : f)));
  }

  function removeFile(id) {
    setPendingFiles((prev) => prev.filter((f) => f.id !== id));
  }

  function clearAll() {
    setPendingFiles([]);
    setError("");
  }

  const periodsUsed = pendingFiles.filter((f) => !f.error).map((f) => f.period);
  const duplicatePeriods = periodsUsed.length !== new Set(periodsUsed).size;
  const validFiles = pendingFiles.filter((f) => !f.error && f.withSupplier && f.withSupplier.length > 0);

  async function confirmImport() {
    if (validFiles.length === 0) return;
    const now = Date.now();
    const next = { ...products };

    // Merge: build a map codigo -> { fornecedor, item, estoque, precoCusto, vendasByPeriod }
    const merged = {};
    for (const file of validFiles) {
      for (const row of file.withSupplier) {
        const codigo = String(row.codigo || "").trim() || String(row.item || "").trim();
        if (!codigo) continue;
        const fornecedor = String(row.fornecedor).trim();
        const item = String(row.item || "").trim() || codigo;
        const estoque = toNumber(row.estoque);
        const precoCusto = row.precoCusto !== undefined && row.precoCusto !== "" ? toNumber(row.precoCusto) : undefined;

        if (!merged[codigo]) {
          merged[codigo] = { codigo, fornecedor, item, estoque, precoCusto, vendas: {} };
        }
        // estoque e custo: mantém o valor lido mais recente do lote (deve coincidir entre os 3 arquivos)
        merged[codigo].estoque = estoque;
        merged[codigo].fornecedor = fornecedor || merged[codigo].fornecedor;
        merged[codigo].item = item || merged[codigo].item;
        if (precoCusto !== undefined) merged[codigo].precoCusto = precoCusto;

        // Pick the sales value matching this file's period, whichever column it landed in
        let salesValue;
        if (file.period === 30) salesValue = row.vendas30 !== undefined ? row.vendas30 : row.vendas90 !== undefined ? row.vendas90 : row.vendas180;
        else if (file.period === 90) salesValue = row.vendas90 !== undefined ? row.vendas90 : row.vendas30 !== undefined ? row.vendas30 : row.vendas180;
        else salesValue = row.vendas180 !== undefined ? row.vendas180 : row.vendas90 !== undefined ? row.vendas90 : row.vendas30;

        merged[codigo].vendas[file.period] = toNumber(salesValue);
      }
    }

    let newCount = 0;
    let updatedCount = 0;

    for (const codigo of Object.keys(merged)) {
      const m = merged[codigo];
      const existing = next[codigo];
      const v30 = m.vendas[30];
      const v90 = m.vendas[90];
      const v180 = m.vendas[180];
      const salesSnapshot = { ts: now, vendas30: v30, vendas90: v90, vendas180: v180, estoque: m.estoque, precoCusto: m.precoCusto };

      if (existing) {
        updatedCount++;
        next[codigo] = {
          ...existing,
          fornecedor: m.fornecedor || existing.fornecedor,
          item: m.item || existing.item,
          estoque: m.estoque,
          precoCusto: m.precoCusto !== undefined ? m.precoCusto : existing.precoCusto,
          vendas30: v30 !== undefined ? v30 : existing.vendas30,
          vendas90: v90 !== undefined ? v90 : existing.vendas90,
          vendas180: v180 !== undefined ? v180 : existing.vendas180,
          lastUpdated: now,
          salesHistory: [...(existing.salesHistory || []), salesSnapshot],
        };
      } else {
        newCount++;
        next[codigo] = {
          codigo,
          fornecedor: m.fornecedor,
          item: m.item,
          estoque: m.estoque,
          precoCusto: m.precoCusto,
          vendas30: v30,
          vendas90: v90,
          vendas180: v180,
          coverageDaysOverride: null,
          lastUpdated: now,
          firstSeen: now,
          salesHistory: [salesSnapshot],
        };
      }
    }

    await persistProducts(next);

    const totalSkipped = validFiles.reduce((acc, f) => acc + (f.withoutSupplierCount || 0), 0);
    const totalRows = validFiles.reduce((acc, f) => acc + (f.total || 0), 0);
    const event = {
      id: uid(),
      ts: now,
      fileName: validFiles.map((f) => `${f.fileName} (${f.period}d)`).join(" + "),
      totalRows,
      importedRows: Object.keys(merged).length,
      skippedNoSupplier: totalSkipped,
      newCount,
      updatedCount,
    };
    await persistHistory([event, ...history]);

    setPendingFiles([]);
    goToProducts();
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Importar relatório</h1>
        <p>Envie as planilhas exportadas do seu ERP (.xlsx ou .csv) — pode selecionar as 3 (30, 90 e 180 dias) de uma vez. Linhas sem fornecedor preenchido são descartadas automaticamente.</p>
      </header>

      <div
        className={"vivo-dropzone" + (dragOver ? " is-drag" : "")}
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={onDrop}
      >
        <FileSpreadsheet size={28} strokeWidth={1.5} />
        <p className="vivo-dropzone-title">Arraste os arquivos aqui</p>
        <p className="vivo-dropzone-sub">ou</p>
        <label className="vivo-btn vivo-btn-primary">
          Escolher arquivos
          <input
            type="file"
            accept=".csv,.xlsx,.xls"
            multiple
            style={{ display: "none" }}
            onChange={(e) => e.target.files?.length && handleFiles(e.target.files)}
          />
        </label>
        <p className="vivo-dropzone-hint">Cada planilha deve ter Fornecedor/Fabricante, Item, Código interno, Estoque atual e a coluna de vendas (30, 90 ou 180 dias).</p>
      </div>

      {loading && <p className="vivo-loading-text" style={{ marginTop: 12 }}>Lendo arquivos…</p>}

      {error && (
        <div className="vivo-alert vivo-alert-error">
          <AlertTriangle size={16} /> {error}
        </div>
      )}

      {pendingFiles.length > 0 && (
        <div className="vivo-card vivo-preview">
          <div className="vivo-preview-head">
            <div>
              <h3>{pendingFiles.length} arquivo(s) carregado(s)</h3>
              <p className="vivo-settings-hint" style={{ margin: 0 }}>
                Confirme o período de vendas de cada arquivo antes de importar. Os produtos serão combinados pelo código interno.
              </p>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button className="vivo-btn vivo-btn-ghost" onClick={clearAll}>Limpar tudo</button>
              <button
                className="vivo-btn vivo-btn-primary"
                onClick={confirmImport}
                disabled={validFiles.length === 0 || duplicatePeriods}
              >
                Confirmar importação
              </button>
            </div>
          </div>

          {duplicatePeriods && (
            <div className="vivo-alert vivo-alert-error" style={{ marginTop: 0, marginBottom: 14 }}>
              <AlertTriangle size={16} /> Dois arquivos estão marcados com o mesmo período. Ajuste antes de confirmar.
            </div>
          )}

          <div className="vivo-pending-files">
            {pendingFiles.map((f) => (
              <div className="vivo-pending-file" key={f.id}>
                <div className="vivo-pending-file-name">
                  <FileSpreadsheet size={15} />
                  <span>{f.fileName}</span>
                </div>
                {f.error ? (
                  <span className="vivo-pending-file-error"><AlertTriangle size={13} /> {f.error}</span>
                ) : (
                  <>
                    <div className="vivo-pending-file-stats">
                      <span><strong>{f.total}</strong> linhas</span>
                      <span className="ok"><CheckCircle2 size={13} /> <strong>{f.withSupplier.length}</strong> com fornecedor</span>
                      {f.withoutSupplierCount > 0 && (
                        <span className="warn"><AlertTriangle size={13} /> <strong>{f.withoutSupplierCount}</strong> ignoradas</span>
                      )}
                    </div>
                    <label className="vivo-period-select">
                      Período de vendas:
                      <select value={f.period} onChange={(e) => updatePeriod(f.id, e.target.value)} className="vivo-select">
                        <option value={30}>30 dias</option>
                        <option value={90}>90 dias</option>
                        <option value={180}>180 dias</option>
                      </select>
                    </label>
                  </>
                )}
                <button className="vivo-icon-btn vivo-icon-danger" onClick={() => removeFile(f.id)} title="Remover arquivo">
                  <Trash2 size={14} />
                </button>
              </div>
            ))}
          </div>

          {validFiles.length > 0 && (
            <div className="vivo-table-wrap" style={{ marginTop: 16 }}>
              <table className="vivo-table">
                <thead>
                  <tr>
                    <th>Fornecedor</th>
                    <th>Item</th>
                    <th>Código</th>
                    <th>Estoque</th>
                    <th>Vlr. Custo</th>
                    <th>Período</th>
                    <th>Vendas no período</th>
                  </tr>
                </thead>
                <tbody>
                  {validFiles.flatMap((f) =>
                    f.withSupplier.slice(0, 3).map((r, i) => (
                      <tr key={f.id + "-" + i}>
                        <td>{r.fornecedor}</td>
                        <td>{r.item}</td>
                        <td className="mono">{r.codigo}</td>
                        <td className="mono">{r.estoque ?? "—"}</td>
                        <td className="mono">{r.precoCusto !== undefined && r.precoCusto !== "" ? fmtCurrency(toNumber(r.precoCusto)) : "—"}</td>
                        <td>{PERIOD_LABELS[f.period]}</td>
                        <td className="mono">{r.vendas30 ?? r.vendas90 ?? r.vendas180 ?? "—"}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
              <p className="vivo-table-more">Pré-visualização de 3 linhas por arquivo. Os dados completos serão combinados na importação.</p>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ====================== PRODUCTS TAB ======================
function ProductsTab({ products, persistProducts, settings }) {
  const [search, setSearch] = useState("");
  const [supplierFilter, setSupplierFilter] = useState("all");
  const [sortBy, setSortBy] = useState("fornecedor_codigo");
  const [expanded, setExpanded] = useState(null);

  const list = useMemo(() => Object.values(products).filter((p) => !p.inactive), [products]);

  const suppliers = useMemo(() => {
    const s = new Set(list.map((p) => p.fornecedor));
    return Array.from(s).sort();
  }, [list]);

  const filtered = useMemo(() => {
    let out = list;
    if (supplierFilter !== "all") out = out.filter((p) => p.fornecedor === supplierFilter);
    if (search.trim()) {
      const q = search.toLowerCase();
      out = out.filter((p) => p.item.toLowerCase().includes(q) || p.codigo.toLowerCase().includes(q));
    }
    out = [...out].sort((a, b) => {
      if (sortBy === "fornecedor_codigo") {
        const f = a.fornecedor.localeCompare(b.fornecedor, "pt-BR");
        if (f !== 0) return f;
        return a.codigo.localeCompare(b.codigo, "pt-BR", { numeric: true });
      }
      if (sortBy === "item") return a.item.localeCompare(b.item, "pt-BR");
      if (sortBy === "estoque") return (a.estoque || 0) - (b.estoque || 0);
      if (sortBy === "vendas30") return (b.vendas30 || 0) - (a.vendas30 || 0);
      if (sortBy === "fornecedor") return a.fornecedor.localeCompare(b.fornecedor, "pt-BR");
      return 0;
    });
    return out;
  }, [list, supplierFilter, search, sortBy]);

  const { page, setPage, totalPages, pageItems, totalCount } = usePagination(filtered);

  async function removeProduct(codigo) {
    const next = { ...products };
    delete next[codigo];
    await persistProducts(next);
  }

  async function setCoverageOverride(codigo, value) {
    const next = { ...products };
    next[codigo] = { ...next[codigo], coverageDaysOverride: value === "" ? null : Number(value) };
    await persistProducts(next);
  }

  if (list.length === 0) {
    return (
      <div className="vivo-page">
        <EmptyState
          icon={Box}
          title="Nenhum produto importado ainda"
          text="Vá em “Importar relatório” e envie a primeira planilha do seu ERP."
        />
      </div>
    );
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Produtos</h1>
        <p>{list.length} produtos cadastrados com fornecedor • atualizados conforme os relatórios importados.</p>
      </header>

      <div className="vivo-toolbar">
        <div className="vivo-search">
          <Search size={15} />
          <input placeholder="Buscar por nome ou código…" value={search} onChange={(e) => setSearch(e.target.value)} />
        </div>
        <select className="vivo-select" value={supplierFilter} onChange={(e) => setSupplierFilter(e.target.value)}>
          <option value="all">Todos os fornecedores</option>
          {suppliers.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <select className="vivo-select" value={sortBy} onChange={(e) => setSortBy(e.target.value)}>
          <option value="fornecedor_codigo">Ordenar: Fornecedor + Código</option>
          <option value="item">Ordenar: Nome</option>
          <option value="fornecedor">Ordenar: Fornecedor</option>
          <option value="estoque">Ordenar: Estoque (menor primeiro)</option>
          <option value="vendas30">Ordenar: Vendas 30d (maior primeiro)</option>
        </select>
      </div>

      <div className="vivo-table-wrap vivo-card">
        <table className="vivo-table vivo-table-products">
          <thead>
            <tr>
              <th></th>
              <th>Item</th>
              <th>Fornecedor</th>
              <th className="mono">Código</th>
              <th className="num">Estoque</th>
              <th className="num">Vlr. Custo</th>
              <th className="num">Vendas 30d</th>
              <th className="num">Vendas 90d</th>
              <th className="num">Vendas 180d</th>
              <th className="num">Média/dia</th>
              <th>Cobertura (dias)</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((p) => {
              const avgDay = (p.vendas30 || 0) / 30;
              const coverageDays = resolveCoverageDays(p, settings);
              const source = coverageSource(p, settings);
              const daysOfStock = avgDay > 0 ? (p.estoque || 0) / avgDay : null;
              const low = daysOfStock !== null && daysOfStock < coverageDays;
              const isOpen = expanded === p.codigo;
              return (
                <React.Fragment key={p.codigo}>
                  <tr className={low ? "is-low" : ""}>
                    <td>
                      <button className="vivo-icon-btn" onClick={() => setExpanded(isOpen ? null : p.codigo)}>
                        {isOpen ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
                      </button>
                    </td>
                    <ItemCell name={p.item} />
                    <td>{p.fornecedor}</td>
                    <td className="mono">{p.codigo}</td>
                    <td className="num mono">{fmtNumber(p.estoque)}</td>
                    <td className="num mono">{p.precoCusto !== undefined ? fmtCurrency(p.precoCusto) : "—"}</td>
                    <td className="num mono">{p.vendas30 ?? "—"}</td>
                    <td className="num mono">{p.vendas90 ?? "—"}</td>
                    <td className="num mono">{p.vendas180 ?? "—"}</td>
                    <td className="num mono">{avgDay ? avgDay.toFixed(2) : "—"}</td>
                    <td>
                      <input
                        type="number"
                        min="1"
                        className="vivo-input-mini"
                        placeholder={String(coverageDays)}
                        value={p.coverageDaysOverride ?? ""}
                        onChange={(e) => setCoverageOverride(p.codigo, e.target.value)}
                        title={source === "fornecedor" ? `Usando cobertura do fornecedor (${coverageDays}d)` : source === "padrao" ? `Usando cobertura padrão (${coverageDays}d)` : "Cobertura específica deste produto"}
                      />
                      {source !== "produto" && <span className="vivo-coverage-tag">{source === "fornecedor" ? "fornecedor" : "padrão"}</span>}
                    </td>
                    <td>
                      <button className="vivo-icon-btn vivo-icon-danger" onClick={() => removeProduct(p.codigo)} title="Remover produto">
                        <Trash2 size={14} />
                      </button>
                    </td>
                  </tr>
                  {isOpen && (
                    <tr className="vivo-row-detail">
                      <td colSpan={12}>
                        <ProductDetail product={p} daysOfStock={daysOfStock} coverageDays={coverageDays} source={source} />
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
      <PaginationControls page={page} setPage={setPage} totalPages={totalPages} totalCount={totalCount} />
    </div>
  );
}

function ProductDetail({ product, daysOfStock, coverageDays, source }) {
  const history = product.salesHistory || [];
  const series30 = useMemo(() => buildIntervalSeries(history, "vendas30"), [history]);

  return (
    <div className="vivo-detail">
      <div className="vivo-detail-grid">
        <div>
          <span className="vivo-detail-label">Cobertura atual de estoque</span>
          <span className="vivo-detail-value">
            {daysOfStock !== null ? `${daysOfStock.toFixed(0)} dias` : "sem histórico de vendas"}
            <span className="vivo-detail-muted">
              {" "}(meta: {coverageDays} dias{source === "fornecedor" ? ", do fornecedor" : source === "padrao" ? ", padrão geral" : ", definida neste produto"})
            </span>
          </span>
        </div>
        <div>
          <span className="vivo-detail-label">Primeira vez visto</span>
          <span className="vivo-detail-value">{fmtDate(product.firstSeen)}</span>
        </div>
        <div>
          <span className="vivo-detail-label">Última atualização</span>
          <span className="vivo-detail-value">{fmtDate(product.lastUpdated)}</span>
        </div>
      </div>

      <div className="vivo-detail-history">
        <span className="vivo-detail-label">Histórico de importações deste item ({history.length})</span>
        <table className="vivo-mini-table">
          <thead>
            <tr><th>Data</th><th className="num">Estoque</th><th className="num">30d</th><th className="num">90d</th><th className="num">180d</th></tr>
          </thead>
          <tbody>
            {history.slice().reverse().map((h, i) => (
              <tr key={i}>
                <td>{fmtDate(h.ts)}</td>
                <td className="num mono">{h.estoque ?? "—"}</td>
                <td className="num mono">{h.vendas30 ?? "—"}</td>
                <td className="num mono">{h.vendas90 ?? "—"}</td>
                <td className="num mono">{h.vendas180 ?? "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {series30.length > 0 && (
        <div className="vivo-detail-history" style={{ marginTop: 18 }}>
          <span className="vivo-detail-label">
            Vendas diárias estimadas entre importações (base: vendas 30 dias)
          </span>
          <p className="vivo-settings-hint" style={{ margin: "2px 0 8px" }}>
            Estimativa: (variação da métrica de 30 dias) ÷ (dias entre as datas). Agrupado por dia — se houver mais de uma importação no mesmo dia, só a mais recente entra no cálculo. Como é uma janela móvel, é uma aproximação — quanto menor o intervalo, mais precisa.
          </p>
          <table className="vivo-mini-table">
            <thead>
              <tr>
                <th>Período</th>
                <th className="num">Dias</th>
                <th className="num">Variação na métrica</th>
                <th className="num">Média est. /dia</th>
              </tr>
            </thead>
            <tbody>
              {series30.slice().reverse().map((iv, i) => (
                <tr key={i}>
                  <td>{fmtDate(iv.from)} → {fmtDate(iv.to)}</td>
                  <td className="num mono">{iv.days}</td>
                  <td className={"num mono" + (iv.delta > 0 ? " vivo-trend-up" : iv.delta < 0 ? " vivo-trend-down" : "")}>
                    {iv.delta > 0 ? "+" : ""}{fmtNumber(iv.delta)}
                  </td>
                  <td className="num mono">{iv.estimatedDailyAvg.toFixed(2)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ====================== ORDERS TAB ======================
function OrdersTab({ products, settings, persistSettings, orders, persistOrders }) {
  const [supplierFilter, setSupplierFilter] = useState("all");
  const [search, setSearch] = useState("");
  const [onlyNeeded, setOnlyNeeded] = useState(true);
  const [qtyOverrides, setQtyOverrides] = useState({});
  const [selected, setSelected] = useState({});
  const [showSettings, setShowSettings] = useState(false);
  const [orderSortBy, setOrderSortBy] = useState("fornecedor_codigo");

  const list = useMemo(() => Object.values(products).filter((p) => !p.inactive), [products]);
  const suppliers = useMemo(() => {
    const s = new Set(list.map((p) => p.fornecedor));
    return Array.from(s).sort();
  }, [list]);

  const computed = useMemo(() => {
    return list.map((p) => {
      const avgDay = (p.vendas30 || 0) / 30; // base sempre vendas 30 dias
      const coverageDays = resolveCoverageDays(p, settings);
      const source = coverageSource(p, settings);
      const idealStock = avgDay * coverageDays;
      const suggested = Math.max(0, Math.ceil(idealStock - (p.estoque || 0)));
      const daysOfStock = avgDay > 0 ? (p.estoque || 0) / avgDay : null;
      return { ...p, avgDay, coverageDays, source, idealStock, suggested, daysOfStock };
    });
  }, [list, settings]);

  const filtered = useMemo(() => {
    let out = computed;
    if (supplierFilter !== "all") out = out.filter((p) => p.fornecedor === supplierFilter);
    if (search.trim()) {
      const q = search.toLowerCase();
      out = out.filter((p) => p.item.toLowerCase().includes(q) || p.codigo.toLowerCase().includes(q));
    }
    if (onlyNeeded) out = out.filter((p) => p.suggested > 0);
    const sorted = [...out];
    if (orderSortBy === "fornecedor_codigo") {
      sorted.sort((a, b) => {
        const f = a.fornecedor.localeCompare(b.fornecedor, "pt-BR");
        if (f !== 0) return f;
        return a.codigo.localeCompare(b.codigo, "pt-BR", { numeric: true });
      });
    } else {
      sorted.sort((a, b) => b.suggested - a.suggested);
    }
    return sorted;
  }, [computed, supplierFilter, search, onlyNeeded, orderSortBy]);

  // Paginação só afeta a RENDERIZAÇÃO da tabela — seleção, totais por fornecedor e
  // exportação continuam considerando toda a lista filtrada, não só a página visível.
  const { page, setPage, totalPages, pageItems, totalCount } = usePagination(filtered);

  function getQty(p) {
    return qtyOverrides[p.codigo] !== undefined ? qtyOverrides[p.codigo] : p.suggested;
  }

  function getSubtotal(p) {
    return getQty(p) * (p.precoCusto || 0);
  }

  function toggleSelect(codigo) {
    setSelected((s) => ({ ...s, [codigo]: !s[codigo] }));
  }

  function selectAllVisible() {
    const next = { ...selected };
    filtered.forEach((p) => { next[p.codigo] = true; });
    setSelected(next);
  }
  function clearSelection() {
    setSelected({});
  }

  const selectedItems = filtered.filter((p) => selected[p.codigo] && getQty(p) > 0);
  const totalUnits = selectedItems.reduce((acc, p) => acc + getQty(p), 0);
  const totalValue = selectedItems.reduce((acc, p) => acc + getSubtotal(p), 0);
  const itemsMissingPrice = selectedItems.filter((p) => !p.precoCusto).length;

  const totalsBySupplier = useMemo(() => {
    const map = {};
    for (const p of selectedItems) {
      if (!map[p.fornecedor]) map[p.fornecedor] = { fornecedor: p.fornecedor, items: 0, units: 0, value: 0, missingPrice: 0 };
      map[p.fornecedor].items += 1;
      map[p.fornecedor].units += getQty(p);
      map[p.fornecedor].value += getSubtotal(p);
      if (!p.precoCusto) map[p.fornecedor].missingPrice += 1;
    }
    return Object.values(map).sort((a, b) => a.fornecedor.localeCompare(b.fornecedor, "pt-BR"));
  }, [selectedItems, qtyOverrides]);

  async function saveOrder() {
    if (selectedItems.length === 0) return;
    const now = Date.now();
    const order = {
      id: uid(),
      ts: now,
      items: selectedItems.map((p) => ({
        codigo: p.codigo,
        item: p.item,
        fornecedor: p.fornecedor,
        estoque: p.estoque,
        vendas30: p.vendas30,
        precoCusto: p.precoCusto,
        quantidade: getQty(p),
        subtotal: getSubtotal(p),
      })),
      totalValue,
      totalsBySupplier,
    };
    await persistOrders([order, ...orders]);
    setSelected({});
    setQtyOverrides({});
  }

  function exportCSV() {
    if (selectedItems.length === 0) return;
    const rows = selectedItems.map((p) => ({
      Fornecedor: p.fornecedor,
      Item: p.item,
      Codigo: p.codigo,
      Estoque_Atual: p.estoque,
      Vendas_30d: p.vendas30 ?? "",
      Quantidade_Pedido: getQty(p),
      Vlr_Custo_Unitario: p.precoCusto ?? "",
      Subtotal: getSubtotal(p).toFixed(2),
    }));
    const csv = Papa.unparse(rows);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `pedido-compra-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Pedidos de compra</h1>
        <p>Quantidade sugerida = (vendas dos últimos 30 dias ÷ 30 × dias de cobertura) − estoque atual.</p>
      </header>

      <div className="vivo-toolbar">
        <div className="vivo-search">
          <Search size={15} />
          <input placeholder="Buscar por nome ou código…" value={search} onChange={(e) => setSearch(e.target.value)} />
        </div>
        <select className="vivo-select" value={supplierFilter} onChange={(e) => setSupplierFilter(e.target.value)}>
          <option value="all">Todos os fornecedores</option>
          {suppliers.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <select className="vivo-select" value={orderSortBy} onChange={(e) => setOrderSortBy(e.target.value)}>
          <option value="fornecedor_codigo">Ordenar: Fornecedor + Código</option>
          <option value="suggested">Ordenar: Maior necessidade primeiro</option>
        </select>
        <label className="vivo-checkbox">
          <input type="checkbox" checked={onlyNeeded} onChange={(e) => setOnlyNeeded(e.target.checked)} />
          Mostrar só itens com necessidade de compra
        </label>
        <button className="vivo-btn vivo-btn-ghost" onClick={() => setShowSettings((s) => !s)}>
          <Settings size={15} /> Cobertura: {settings.defaultCoverageDays}d padrão
        </button>
      </div>

      {showSettings && (
        <div className="vivo-card vivo-settings-panel">
          <label>
            Dias de cobertura padrão (aplica a todos os produtos sem regra própria)
            <input
              type="number"
              min="1"
              className="vivo-input"
              value={settings.defaultCoverageDays}
              onChange={(e) => persistSettings({ ...settings, defaultCoverageDays: Math.max(1, Number(e.target.value) || 1) })}
            />
          </label>

          <div className="vivo-supplier-coverage">
            <span className="vivo-detail-label" style={{ display: "block", marginTop: 16, marginBottom: 8 }}>
              Cobertura por fornecedor (sobrescreve o padrão para todos os itens dele)
            </span>
            {suppliers.map((s) => (
              <div className="vivo-supplier-coverage-row" key={s}>
                <span className="vivo-supplier-coverage-name">{s}</span>
                <input
                  type="number"
                  min="1"
                  className="vivo-input-mini"
                  placeholder={String(settings.defaultCoverageDays)}
                  value={settings.supplierCoverageDays?.[s] ?? ""}
                  onChange={(e) => {
                    const val = e.target.value === "" ? null : Math.max(1, Number(e.target.value) || 1);
                    const nextSupplierMap = { ...(settings.supplierCoverageDays || {}) };
                    if (val === null) delete nextSupplierMap[s];
                    else nextSupplierMap[s] = val;
                    persistSettings({ ...settings, supplierCoverageDays: nextSupplierMap });
                  }}
                />
                <span className="vivo-supplier-coverage-unit">dias</span>
              </div>
            ))}
          </div>

          <p className="vivo-settings-hint">Prioridade do cálculo: cobertura própria do produto (na aba “Produtos”) → cobertura do fornecedor → cobertura padrão.</p>
        </div>
      )}

      <div className="vivo-action-row">
        <button className="vivo-btn vivo-btn-ghost" onClick={selectAllVisible}>Selecionar todos visíveis</button>
        <button className="vivo-btn vivo-btn-ghost" onClick={clearSelection}>Limpar seleção</button>
        <div className="vivo-action-spacer" />
        <span className="vivo-action-summary">
          {selectedItems.length} itens selecionados • {fmtNumber(totalUnits)} unidades • <strong>{fmtCurrency(totalValue)}</strong>
          {itemsMissingPrice > 0 && <span className="vivo-action-warning"> ({itemsMissingPrice} sem preço de custo)</span>}
        </span>
        <button className="vivo-btn vivo-btn-secondary" onClick={exportCSV} disabled={!selectedItems.length}>Exportar CSV</button>
        <button className="vivo-btn vivo-btn-primary" onClick={saveOrder} disabled={!selectedItems.length}>Salvar pedido</button>
      </div>

      {totalsBySupplier.length > 0 && (
        <div className="vivo-card vivo-supplier-totals">
          <span className="vivo-detail-label" style={{ display: "block", marginBottom: 8 }}>Total por fornecedor (seleção atual)</span>
          <table className="vivo-mini-table">
            <thead>
              <tr>
                <th>Fornecedor</th>
                <th className="num">Itens</th>
                <th className="num">Unidades</th>
                <th className="num">Valor total</th>
              </tr>
            </thead>
            <tbody>
              {totalsBySupplier.map((t) => (
                <tr key={t.fornecedor}>
                  <td>{t.fornecedor}</td>
                  <td className="num mono">{t.items}</td>
                  <td className="num mono">{fmtNumber(t.units)}</td>
                  <td className="num mono">
                    {fmtCurrency(t.value)}
                    {t.missingPrice > 0 && <sup className="vivo-source-tag" title="itens sem preço de custo cadastrado"> {t.missingPrice} s/ preço</sup>}
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr>
                <td><strong>Total geral</strong></td>
                <td className="num mono"><strong>{selectedItems.length}</strong></td>
                <td className="num mono"><strong>{fmtNumber(totalUnits)}</strong></td>
                <td className="num mono"><strong>{fmtCurrency(totalValue)}</strong></td>
              </tr>
            </tfoot>
          </table>
        </div>
      )}

      <div className="vivo-table-wrap vivo-card">
        <table className="vivo-table">
          <thead>
            <tr>
              <th></th>
              <th>Item</th>
              <th>Fornecedor</th>
              <th className="mono">Código</th>
              <th className="num">Estoque</th>
              <th className="num">Vendas 30d</th>
              <th className="num">Cobertura</th>
              <th className="num">Sugerido</th>
              <th className="num">Quantidade a pedir</th>
              <th className="num">Vlr. Custo</th>
              <th className="num">Subtotal</th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((p) => (
              <tr key={p.codigo} className={selected[p.codigo] ? "is-selected" : ""}>
                <td>
                  <input type="checkbox" checked={!!selected[p.codigo]} onChange={() => toggleSelect(p.codigo)} />
                </td>
                <ItemCell name={p.item} />
                <td>{p.fornecedor}</td>
                <td className="mono">{p.codigo}</td>
                <td className="num mono">{fmtNumber(p.estoque)}</td>
                <td className="num mono">{p.vendas30 ?? "—"}</td>
                <td className="num mono" title={p.source === "fornecedor" ? "Cobertura do fornecedor" : p.source === "padrao" ? "Cobertura padrão" : "Cobertura própria do produto"}>
                  {p.coverageDays}d{p.source === "fornecedor" ? <sup className="vivo-source-tag">forn.</sup> : p.source === "produto" ? <sup className="vivo-source-tag">item</sup> : null}
                </td>
                <td className="num mono vivo-suggested">{fmtNumber(p.suggested)}</td>
                <td>
                  <input
                    type="number"
                    min="0"
                    className="vivo-input-mini"
                    value={getQty(p)}
                    onChange={(e) => setQtyOverrides((q) => ({ ...q, [p.codigo]: Math.max(0, Number(e.target.value) || 0) }))}
                  />
                </td>
                <td className="num mono">{p.precoCusto !== undefined ? fmtCurrency(p.precoCusto) : <span className="vivo-no-price">—</span>}</td>
                <td className="num mono">{fmtCurrency(getSubtotal(p))}</td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={11} className="vivo-table-empty">Nenhum produto encontrado com os filtros atuais.</td></tr>
            )}
          </tbody>
        </table>
      </div>
      <PaginationControls page={page} setPage={setPage} totalPages={totalPages} totalCount={totalCount} />

      {orders.length > 0 && (
        <div className="vivo-saved-orders">
          <h3>Pedidos salvos</h3>
          {orders.slice(0, 6).map((o) => (
            <div className="vivo-card vivo-order-card" key={o.id}>
              <div className="vivo-order-card-head">
                <span>{fmtDate(o.ts)}</span>
                <span>
                  {o.items.length} itens • {fmtNumber(o.items.reduce((a, i) => a + i.quantidade, 0))} unidades
                  {o.totalValue !== undefined && <> • <strong>{fmtCurrency(o.totalValue)}</strong></>}
                </span>
              </div>
              {o.totalsBySupplier && o.totalsBySupplier.length > 1 && (
                <div className="vivo-order-card-suppliers">
                  {o.totalsBySupplier.map((t) => (
                    <span key={t.fornecedor} className="vivo-pill vivo-pill-supplier">
                      {t.fornecedor}: {fmtCurrency(t.value)}
                    </span>
                  ))}
                </div>
              )}
              <div className="vivo-order-card-items">
                {o.items.slice(0, 4).map((i) => (
                  <span key={i.codigo} className="vivo-pill">{i.item} × {i.quantidade}</span>
                ))}
                {o.items.length > 4 && <span className="vivo-pill vivo-pill-muted">+ {o.items.length - 4} itens</span>}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ====================== PERFORMANCE TAB ======================
function PerformanceTab({ products }) {
  const [search, setSearch] = useState("");
  const [supplierFilter, setSupplierFilter] = useState("all");
  const [trendFilter, setTrendFilter] = useState("all");
  const [sortBy, setSortBy] = useState("deltaPct");

  const list = useMemo(() => Object.values(products).filter((p) => !p.inactive), [products]);
  const suppliers = useMemo(() => {
    const s = new Set(list.map((p) => p.fornecedor));
    return Array.from(s).sort();
  }, [list]);

  const computed = useMemo(() => {
    return list.map((p) => {
      const trend = getTrendSummary(p.salesHistory || []);
      return { ...p, trend };
    });
  }, [list]);

  const filtered = useMemo(() => {
    let out = computed;
    if (supplierFilter !== "all") out = out.filter((p) => p.fornecedor === supplierFilter);
    if (search.trim()) {
      const q = search.toLowerCase();
      out = out.filter((p) => p.item.toLowerCase().includes(q) || p.codigo.toLowerCase().includes(q));
    }
    if (trendFilter !== "all") out = out.filter((p) => p.trend.status === trendFilter);
    const sorted = [...out];
    if (sortBy === "deltaPct") {
      sorted.sort((a, b) => (b.trend.deltaPct ?? -Infinity) - (a.trend.deltaPct ?? -Infinity));
    } else if (sortBy === "deltaPctAsc") {
      sorted.sort((a, b) => (a.trend.deltaPct ?? Infinity) - (b.trend.deltaPct ?? Infinity));
    } else if (sortBy === "fornecedor_codigo") {
      sorted.sort((a, b) => {
        const f = a.fornecedor.localeCompare(b.fornecedor, "pt-BR");
        if (f !== 0) return f;
        return a.codigo.localeCompare(b.codigo, "pt-BR", { numeric: true });
      });
    } else if (sortBy === "lastAvg") {
      sorted.sort((a, b) => (b.trend.lastAvg ?? -Infinity) - (a.trend.lastAvg ?? -Infinity));
    }
    return sorted;
  }, [computed, supplierFilter, search, trendFilter, sortBy]);

  const { page, setPage, totalPages, pageItems, totalCount } = usePagination(filtered);

  const counts = useMemo(() => {
    const c = { subindo: 0, caindo: 0, estavel: 0, primeira_leitura: 0, sem_dados: 0 };
    computed.forEach((p) => { c[p.trend.status] = (c[p.trend.status] || 0) + 1; });
    return c;
  }, [computed]);

  if (list.length === 0) {
    return (
      <div className="vivo-page">
        <EmptyState icon={TrendingUp} title="Nenhum dado de performance ainda" text="Importe relatórios em pelo menos duas datas diferentes para começar a ver tendências de vendas." />
      </div>
    );
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Tendência de vendas</h1>
        <p>Estimativa de vendas diárias por item, comparando o intervalo mais recente com o anterior (base: vendas 30 dias).</p>
      </header>

      <div className="vivo-trend-summary">
        <button className={"vivo-trend-chip vivo-trend-chip-up" + (trendFilter === "subindo" ? " is-active" : "")} onClick={() => setTrendFilter(trendFilter === "subindo" ? "all" : "subindo")}>
          <TrendingUp size={14} /> {counts.subindo} subindo
        </button>
        <button className={"vivo-trend-chip vivo-trend-chip-down" + (trendFilter === "caindo" ? " is-active" : "")} onClick={() => setTrendFilter(trendFilter === "caindo" ? "all" : "caindo")}>
          <TrendingDown size={14} /> {counts.caindo} caindo
        </button>
        <button className={"vivo-trend-chip vivo-trend-chip-flat" + (trendFilter === "estavel" ? " is-active" : "")} onClick={() => setTrendFilter(trendFilter === "estavel" ? "all" : "estavel")}>
          <Minus size={14} /> {counts.estavel} estável
        </button>
        {(counts.primeira_leitura > 0 || counts.sem_dados > 0) && (
          <span className="vivo-trend-chip-muted">{counts.primeira_leitura + counts.sem_dados} sem comparação ainda</span>
        )}
      </div>

      <div className="vivo-toolbar">
        <div className="vivo-search">
          <Search size={15} />
          <input placeholder="Buscar por nome ou código…" value={search} onChange={(e) => setSearch(e.target.value)} />
        </div>
        <select className="vivo-select" value={supplierFilter} onChange={(e) => setSupplierFilter(e.target.value)}>
          <option value="all">Todos os fornecedores</option>
          {suppliers.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <select className="vivo-select" value={sortBy} onChange={(e) => setSortBy(e.target.value)}>
          <option value="deltaPct">Ordenar: Maior alta primeiro</option>
          <option value="deltaPctAsc">Ordenar: Maior queda primeiro</option>
          <option value="lastAvg">Ordenar: Maior média diária</option>
          <option value="fornecedor_codigo">Ordenar: Fornecedor + Código</option>
        </select>
      </div>

      <div className="vivo-table-wrap vivo-card">
        <table className="vivo-table">
          <thead>
            <tr>
              <th>Item</th>
              <th>Fornecedor</th>
              <th className="mono">Código</th>
              <th className="num">Média est. /dia (atual)</th>
              <th className="num">Média est. /dia (anterior)</th>
              <th className="num">Variação</th>
              <th>Tendência</th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((p) => (
              <tr key={p.codigo}>
                <ItemCell name={p.item} />
                <td>{p.fornecedor}</td>
                <td className="mono">{p.codigo}</td>
                <td className="num mono">{p.trend.lastAvg !== null ? p.trend.lastAvg.toFixed(2) : "—"}</td>
                <td className="num mono">{p.trend.prevAvg !== undefined ? p.trend.prevAvg?.toFixed(2) : "—"}</td>
                <td className={"num mono" + (p.trend.deltaPct > 0 ? " vivo-trend-up" : p.trend.deltaPct < 0 ? " vivo-trend-down" : "")}>
                  {p.trend.deltaPct !== null ? `${p.trend.deltaPct > 0 ? "+" : ""}${p.trend.deltaPct.toFixed(0)}%` : "—"}
                </td>
                <td><TrendBadge status={p.trend.status} /></td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={7} className="vivo-table-empty">Nenhum produto encontrado com os filtros atuais.</td></tr>
            )}
          </tbody>
        </table>
      </div>
      <PaginationControls page={page} setPage={setPage} totalPages={totalPages} totalCount={totalCount} />
    </div>
  );
}

function TrendBadge({ status }) {
  if (status === "subindo") return <span className="vivo-badge vivo-badge-up"><TrendingUp size={12} /> Subindo</span>;
  if (status === "caindo") return <span className="vivo-badge vivo-badge-down"><TrendingDown size={12} /> Caindo</span>;
  if (status === "estavel") return <span className="vivo-badge vivo-badge-flat"><Minus size={12} /> Estável</span>;
  if (status === "primeira_leitura") return <span className="vivo-badge vivo-badge-muted">1ª leitura</span>;
  return <span className="vivo-badge vivo-badge-muted">Sem dados</span>;
}

// ====================== STOCK QUALITY TAB ======================
// Constrói série temporal de capital imobilizado (estoque × custo) nos últimos N dias,
// considerando APENAS os produtos que estão em excesso de estoque (isExcess = true),
// para manter consistência com o valor exibido no filtro de fornecedor.
// O conjunto "excessCodes" é passado de fora para garantir que os dois cálculos
// usem exatamente os mesmos produtos.
function buildCapitalSeries(products, supplierFilter, excessCodes, days = 90) {
  const now = Date.now();
  const cutoff = now - days * MS_PER_DAY;
  const productList = Object.values(products).filter((p) =>
    !p.inactive &&
    excessCodes.has(p.codigo) &&
    (supplierFilter === "all" || p.fornecedor === supplierFilter)
  );

  // Coleta todos os snapshots dentro da janela, agrupados por dia
  const byDay = {};
  for (const p of productList) {
    const custo = p.precoCusto || 0;
    for (const snap of (p.salesHistory || [])) {
      if (snap.ts < cutoff) continue;
      const dk = dayKey(snap.ts);
      if (!byDay[dk]) byDay[dk] = { ts: snap.ts, items: {} };
      // Guarda o snapshot mais recente de cada produto naquele dia
      if (!byDay[dk].items[p.codigo] || snap.ts > byDay[dk].items[p.codigo].ts) {
        byDay[dk].items[p.codigo] = { ts: snap.ts, estoque: snap.estoque ?? 0, custo };
      }
      if (snap.ts > byDay[dk].ts) byDay[dk].ts = snap.ts;
    }
  }

  // Soma capital por dia
  const series = Object.entries(byDay)
    .map(([dk, { ts, items }]) => ({
      dk,
      ts,
      capital: Object.values(items).reduce((acc, i) => acc + (i.estoque * i.custo), 0),
    }))
    .sort((a, b) => a.ts - b.ts);

  return series;
}

function CapitalChart({ products, supplierFilter, excessCodes }) {
  const [hover, setHover] = useState(null);
  const svgRef = React.useRef(null);

  const series = useMemo(
    () => buildCapitalSeries(products, supplierFilter, excessCodes, 90),
    [products, supplierFilter, excessCodes]
  );

  if (series.length < 2) {
    return (
      <div className="vivo-capital-chart vivo-capital-chart-empty">
        <span>Dados insuficientes para o gráfico (mínimo 2 importações nos últimos 90 dias)</span>
      </div>
    );
  }

  const W = 420, H = 72, PAD = { top: 8, right: 12, bottom: 20, left: 12 };
  const innerW = W - PAD.left - PAD.right;
  const innerH = H - PAD.top - PAD.bottom;

  const values = series.map((s) => s.capital);
  const minV = Math.min(...values);
  const maxV = Math.max(...values);
  const range = maxV - minV || 1;

  const toX = (i) => PAD.left + (i / (series.length - 1)) * innerW;
  const toY = (v) => PAD.top + innerH - ((v - minV) / range) * innerH;

  const points = series.map((s, i) => `${toX(i)},${toY(s.capital)}`).join(" ");
  const areaPoints = `${toX(0)},${PAD.top + innerH} ${points} ${toX(series.length - 1)},${PAD.top + innerH}`;

  const last = series[series.length - 1];
  const first = series[0];
  const delta = last.capital - first.capital;
  const deltaPct = first.capital > 0 ? (delta / first.capital) * 100 : 0;
  const trend = Math.abs(deltaPct) < 3 ? "flat" : delta > 0 ? "up" : "down";
  const trendColor = trend === "up" ? "var(--rust)" : trend === "down" ? "var(--olive-dark)" : "var(--ink-soft)";

  // Encontra o ponto mais próximo do mouse baseado na posição X
  function handleMouseMove(e) {
    const svg = svgRef.current;
    if (!svg) return;
    const rect = svg.getBoundingClientRect();
    const mouseX = e.clientX - rect.left;
    // Converte posição do mouse para coordenadas do viewBox
    const scaleX = W / rect.width;
    const vbX = mouseX * scaleX;
    // Encontra o índice mais próximo
    let closest = 0;
    let minDist = Infinity;
    series.forEach((_, i) => {
      const dist = Math.abs(toX(i) - vbX);
      if (dist < minDist) { minDist = dist; closest = i; }
    });
    const s = series[closest];
    setHover({ i: closest, x: toX(closest), y: toY(s.capital), capital: s.capital, ts: s.ts });
  }

  // Decide se o tooltip vai pra esquerda ou direita do ponto para não sair da área
  const tooltipLeft = hover ? hover.i < series.length * 0.65 : false;

  // Variação do ponto em hover em relação ao ponto anterior
  const hoverDelta = hover && hover.i > 0
    ? hover.capital - series[hover.i - 1].capital
    : null;

  const displayPoint = hover || { capital: last.capital, ts: last.ts, x: toX(series.length - 1), y: toY(last.capital) };

  return (
    <div className="vivo-capital-chart">
      <div className="vivo-capital-chart-header">
        <div>
          <span className="vivo-capital-chart-label">Capital imobilizado — últimos 90 dias</span>
          {supplierFilter !== "all" && <span className="vivo-capital-chart-supplier"> · {supplierFilter}</span>}
        </div>
        <div className="vivo-capital-chart-summary">
          {hover ? (
            <>
              <span style={{ fontWeight: 600, color: "var(--ink)" }}>{fmtCurrency(hover.capital)}</span>
              <span className="vivo-capital-chart-delta" style={{ color: "var(--ink-soft)" }}>
                {fmtDateShort(hover.ts)}
                {hoverDelta !== null && (
                  <span style={{ color: hoverDelta > 0 ? "var(--rust)" : "var(--olive-dark)", marginLeft: 6 }}>
                    {hoverDelta > 0 ? "+" : ""}{fmtCurrency(hoverDelta)}
                  </span>
                )}
              </span>
            </>
          ) : (
            <>
              <span style={{ color: trendColor, fontWeight: 600 }}>
                {trend === "up" ? "▲" : trend === "down" ? "▼" : "●"} {fmtCurrency(last.capital)}
              </span>
              <span className="vivo-capital-chart-delta" style={{ color: trendColor }}>
                {deltaPct > 0 ? "+" : ""}{deltaPct.toFixed(1)}% desde {fmtDateShort(first.ts)}
              </span>
            </>
          )}
        </div>
      </div>

      <div style={{ position: "relative" }}>
        <svg
          ref={svgRef}
          viewBox={`0 0 ${W} ${H}`}
          className="vivo-capital-chart-svg"
          preserveAspectRatio="none"
          onMouseMove={handleMouseMove}
          onMouseLeave={() => setHover(null)}
          style={{ cursor: "crosshair" }}
        >
          <defs>
            <linearGradient id="capGrad" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor={trendColor} stopOpacity="0.18" />
              <stop offset="100%" stopColor={trendColor} stopOpacity="0.02" />
            </linearGradient>
          </defs>
          <polygon points={areaPoints} fill="url(#capGrad)" />
          <polyline points={points} fill="none" stroke={trendColor} strokeWidth="1.8" strokeLinejoin="round" strokeLinecap="round" />

          {/* Linha vertical de hover */}
          {hover && (
            <line
              x1={hover.x} y1={PAD.top}
              x2={hover.x} y2={PAD.top + innerH}
              stroke="var(--ink-soft)" strokeWidth="1" strokeDasharray="3,2"
            />
          )}

          {/* Ponto ativo (hover ou último) */}
          <circle
            cx={displayPoint.x} cy={displayPoint.y}
            r={hover ? 4 : 3}
            fill={hover ? "var(--ink)" : trendColor}
            stroke="var(--card)" strokeWidth="1.5"
          />

          {/* Labels de data no eixo X */}
          <text x={toX(0)} y={H - 4} fontSize="9" fill="var(--ink-soft)" textAnchor="start">{fmtDateShort(first.ts)}</text>
          <text x={toX(series.length - 1)} y={H - 4} fontSize="9" fill="var(--ink-soft)" textAnchor="end">{fmtDateShort(last.ts)}</text>
        </svg>

        {/* Tooltip flutuante */}
        {hover && (
          <div
            className="vivo-chart-tooltip"
            style={{
              left: tooltipLeft ? `calc(${(hover.x / W) * 100}% + 8px)` : "auto",
              right: tooltipLeft ? "auto" : `calc(${((W - hover.x) / W) * 100}% + 8px)`,
              top: "4px",
            }}
          >
            <div className="vivo-chart-tooltip-date">{new Date(hover.ts).toLocaleDateString("pt-BR", { day: "2-digit", month: "short" })}</div>
            <div className="vivo-chart-tooltip-value">{fmtCurrency(hover.capital)}</div>
            {hoverDelta !== null && (
              <div className="vivo-chart-tooltip-delta" style={{ color: hoverDelta > 0 ? "var(--rust)" : hoverDelta < 0 ? "var(--olive-dark)" : "var(--ink-soft)" }}>
                {hoverDelta > 0 ? "+" : ""}{fmtCurrency(hoverDelta)} vs dia anterior
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

function StockQualityTab({ products, persistProducts }) {
  const [search, setSearch] = useState("");
  const [supplierFilter, setSupplierFilter] = useState("all");
  const [reasonFilter, setReasonFilter] = useState("all");
  const [minDays, setMinDays] = useState("");
  const [maxDays, setMaxDays] = useState("");
  const [sortBy, setSortBy] = useState("stockDays_desc");
  const [scrollWidth, setScrollWidth] = useState(0);
  const [selected, setSelected] = useState({});
  const topScrollRef = React.useRef(null);
  const tableWrapRef = React.useRef(null);

  // Mede a largura real da tabela para que a barra de rolagem superior tenha o mesmo "tamanho de conteúdo"
  const measureScrollWidth = useCallback(() => {
    if (tableWrapRef.current) setScrollWidth(tableWrapRef.current.scrollWidth);
  }, []);

  useEffect(() => {
    measureScrollWidth();
    window.addEventListener("resize", measureScrollWidth);
    return () => window.removeEventListener("resize", measureScrollWidth);
  }, [measureScrollWidth]);

  function syncFromTop(e) {
    if (tableWrapRef.current) tableWrapRef.current.scrollLeft = e.target.scrollLeft;
  }
  function syncFromTable(e) {
    if (topScrollRef.current) topScrollRef.current.scrollLeft = e.target.scrollLeft;
  }

  const list = useMemo(() => Object.values(products).filter((p) => !p.inactive), [products]);

  const computed = useMemo(() => {
    return list
      .map((p) => ({ ...p, quality: getStockQuality(p) }))
      .filter((p) => p.quality.isExcess);
  }, [list]);

  // Set dos códigos de produtos em excesso — usado pelo gráfico para filtrar os
  // mesmos itens que o capitalBySupplier, garantindo consistência entre os dois valores.
  const excessCodes = useMemo(() => new Set(computed.map((p) => p.codigo)), [computed]);

  const suppliers = useMemo(() => {
    const s = new Set(computed.map((p) => p.fornecedor));
    return Array.from(s).sort();
  }, [computed]);

  const capitalBySupplier = useMemo(() => {
    const map = {};
    for (const p of computed) {
      map[p.fornecedor] = (map[p.fornecedor] || 0) + p.quality.capitalImobilizado;
    }
    return map;
  }, [computed]);

  const filtered = useMemo(() => {
    let out = computed;
    if (supplierFilter !== "all") out = out.filter((p) => p.fornecedor === supplierFilter);
    if (reasonFilter !== "all") out = out.filter((p) => p.qualityReason === reasonFilter);
    if (search.trim()) {
      const q = search.toLowerCase();
      out = out.filter((p) => p.item.toLowerCase().includes(q) || p.codigo.toLowerCase().includes(q));
    }
    const min = minDays !== "" ? Number(minDays) : null;
    const max = maxDays !== "" ? Number(maxDays) : null;
    if (min !== null) out = out.filter((p) => p.quality.stockDays >= min);
    if (max !== null) out = out.filter((p) => p.quality.stockDays !== Infinity && p.quality.stockDays <= max);

    const sorted = [...out];
    if (sortBy === "stockDays_desc") {
      sorted.sort((a, b) => b.quality.stockDays - a.quality.stockDays);
    } else if (sortBy === "stockDays_asc") {
      sorted.sort((a, b) => a.quality.stockDays - b.quality.stockDays);
    } else if (sortBy === "capital_desc") {
      sorted.sort((a, b) => b.quality.capitalImobilizado - a.quality.capitalImobilizado);
    } else if (sortBy === "fornecedor_codigo") {
      sorted.sort((a, b) => {
        const f = a.fornecedor.localeCompare(b.fornecedor, "pt-BR");
        if (f !== 0) return f;
        return a.codigo.localeCompare(b.codigo, "pt-BR", { numeric: true });
      });
    }
    return sorted;
  }, [computed, supplierFilter, reasonFilter, search, minDays, maxDays, sortBy]);

  // Paginação só afeta a renderização da tabela. totalCapital e seleção em massa
  // continuam considerando toda a lista filtrada, não só a página visível.
  const { page, setPage, totalPages, pageItems, totalCount } = usePagination(filtered);

  useEffect(() => {
    // Remede a largura após a tabela re-renderizar com novos dados/filtros
    const id = setTimeout(measureScrollWidth, 0);
    return () => clearTimeout(id);
  }, [pageItems, measureScrollWidth]);

  const totalCapital = filtered.reduce((acc, p) => acc + p.quality.capitalImobilizado, 0);

  async function updateClassification(codigo, field, value) {
    const next = { ...products };
    next[codigo] = { ...next[codigo], [field]: value === "" ? null : value };
    await persistProducts(next);
  }

  async function toggleMonitoring(codigo) {
    const next = { ...products };
    const p = next[codigo];
    const isOn = !!p.monitoringActive;
    if (isOn) {
      // Desativando: limpa o ponto de partida, mas mantém a classificação (motivo/situação/margem)
      next[codigo] = { ...p, monitoringActive: false, monitoringStartedAt: null, monitoringStartV30: null };
    } else {
      // Ativando agora: grava a data e o snapshot de vendas 90D como ponto de partida
      next[codigo] = {
        ...p,
        monitoringActive: true,
        monitoringStartedAt: Date.now(),
        monitoringStartV30: p.vendas30 ?? null,
      };
    }
    await persistProducts(next);
  }

  function toggleSelect(codigo) {
    setSelected((s) => ({ ...s, [codigo]: !s[codigo] }));
  }
  function selectAllVisible() {
    const next = { ...selected };
    filtered.forEach((p) => { next[p.codigo] = true; });
    setSelected(next);
  }
  function clearSelection() {
    setSelected({});
  }

  const selectedCodes = Object.keys(selected).filter((c) => selected[c]);
  const selectedCount = filtered.filter((p) => selected[p.codigo]).length;

  async function sendSelectedToMonitoring() {
    if (selectedCount === 0) return;
    const next = { ...products };
    const now = Date.now();
    for (const p of filtered) {
      if (!selected[p.codigo]) continue;
      const existing = next[p.codigo];
      if (existing.monitoringActive) continue; // já monitorado: não reinicia o ponto de partida
      next[p.codigo] = {
        ...existing,
        monitoringActive: true,
        monitoringStartedAt: now,
        monitoringStartV30: existing.vendas30 ?? null,
      };
    }
    await persistProducts(next);
    setSelected({});
  }

  if (list.length === 0) {
    return (
      <div className="vivo-page">
        <EmptyState icon={AlertTriangle} title="Nenhum produto importado ainda" text="Importe relatórios para começar a identificar excesso de estoque." />
      </div>
    );
  }

  if (computed.length === 0) {
    return (
      <div className="vivo-page">
        <header className="vivo-page-head">
          <h1>Qualidade de Estoque</h1>
          <p>Itens com tempo de estoque acima de {STOCK_QUALITY_THRESHOLD_DAYS} dias de previsão de vendas (ou sem vendas nos últimos 30 dias).</p>
        </header>
        <EmptyState icon={CheckCircle2} title="Nenhum item em excesso de estoque" text="Por enquanto, nenhum produto excede o limite configurado. Bom sinal." />
      </div>
    );
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Qualidade de Estoque</h1>
        <p>Itens com tempo de estoque acima de {STOCK_QUALITY_THRESHOLD_DAYS} dias de previsão de vendas (base: vendas 30 dias). Itens sem vendas entram com tempo de estoque infinito.</p>
      </header>

      <div className="vivo-toolbar">
        <div className="vivo-search">
          <Search size={15} />
          <input placeholder="Buscar por nome ou código…" value={search} onChange={(e) => setSearch(e.target.value)} />
        </div>
        <select className="vivo-select" value={supplierFilter} onChange={(e) => setSupplierFilter(e.target.value)}>
          <option value="all">Todos os fornecedores</option>
          {suppliers.map((s) => (
            <option key={s} value={s}>{s} — {fmtCurrency(capitalBySupplier[s])}</option>
          ))}
        </select>
        <select className="vivo-select" value={reasonFilter} onChange={(e) => setReasonFilter(e.target.value)}>
          <option value="all">Todos os motivos</option>
          {QUALITY_REASON_TAGS.map((t) => <option key={t} value={t}>{t}</option>)}
          <option value="">Sem motivo definido</option>
        </select>
        <div className="vivo-range-filter">
          <span>Tempo de estoque:</span>
          <input type="number" min="0" className="vivo-input-mini" placeholder="mín." value={minDays} onChange={(e) => setMinDays(e.target.value)} />
          <span>–</span>
          <input type="number" min="0" className="vivo-input-mini" placeholder="máx." value={maxDays} onChange={(e) => setMaxDays(e.target.value)} />
          <span>dias</span>
        </div>
        <select className="vivo-select" value={sortBy} onChange={(e) => setSortBy(e.target.value)}>
          <option value="stockDays_desc">Ordenar: Tempo de estoque (maior → menor)</option>
          <option value="stockDays_asc">Ordenar: Tempo de estoque (menor → maior)</option>
          <option value="capital_desc">Ordenar: Capital imobilizado (maior primeiro)</option>
          <option value="fornecedor_codigo">Ordenar: Fornecedor + Código</option>
        </select>
      </div>

      <div className="vivo-quality-summary">
        <span><strong>{filtered.length}</strong> itens em excesso</span>
        <span>•</span>
        <span>Capital imobilizado total: <strong className="vivo-capital-danger">{fmtCurrency(totalCapital)}</strong></span>
      </div>

      <CapitalChart products={products} supplierFilter={supplierFilter} excessCodes={excessCodes} />

      <div className="vivo-action-row">
        <button className="vivo-btn vivo-btn-ghost" onClick={selectAllVisible}>Selecionar todos visíveis</button>
        <button className="vivo-btn vivo-btn-ghost" onClick={clearSelection}>Limpar seleção</button>
        <div className="vivo-action-spacer" />
        <span className="vivo-action-summary">{selectedCount} selecionado(s)</span>
        <button className="vivo-btn vivo-btn-primary" onClick={sendSelectedToMonitoring} disabled={selectedCount === 0}>
          Enviar para Monitorados
        </button>
      </div>

      <div
        className="vivo-top-scrollbar"
        ref={topScrollRef}
        onScroll={syncFromTop}
      >
        <div style={{ width: scrollWidth, height: 1 }} />
      </div>

      <div className="vivo-table-wrap vivo-table-wrap-sticky vivo-card" ref={tableWrapRef} onScroll={syncFromTable}>
        <table className="vivo-table vivo-table-quality vivo-table-compact">
          <thead>
            <tr>
              <th></th>
              <th>Item</th>
              <th>Fornecedor</th>
              <th className="mono">Código</th>
              <th className="num">Estoque</th>
              <th className="num">Vendas 30D</th>
              <th className="num">Preço de Custo</th>
              <th className="num">Capital Imobilizado</th>
              <th className="num">Tempo de estoque</th>
              <th>Motivo</th>
              <th>Ação</th>
              <th>Margem MC</th>
              <th>Monitoramento</th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((p) => (
              <tr key={p.codigo} className={selected[p.codigo] ? "is-selected" : ""}>
                <td>
                  <input type="checkbox" checked={!!selected[p.codigo]} onChange={() => toggleSelect(p.codigo)} />
                </td>
                <ItemCell name={p.item} />
                <td>{p.fornecedor}</td>
                <td className="mono">{p.codigo}</td>
                <td className="num mono">{fmtNumber(p.estoque)}</td>
                <td className="num mono">{p.vendas30 ?? "—"}</td>
                <td className="num mono">{p.precoCusto !== undefined ? fmtCurrency(p.precoCusto) : <span className="vivo-no-price">—</span>}</td>
                <td className="num mono vivo-capital-danger">{fmtCurrency(p.quality.capitalImobilizado)}</td>
                <td className="num mono vivo-stockdays">
                  {p.quality.stockDays === Infinity ? "sem vendas" : `${Math.round(p.quality.stockDays)}d`}
                </td>
                <td>
                  <select
                    className={"vivo-select-tag" + tagColorClass("reason", p.qualityReason)}
                    value={p.qualityReason || ""}
                    onChange={(e) => updateClassification(p.codigo, "qualityReason", e.target.value)}
                  >
                    <option value="">—</option>
                    {QUALITY_REASON_TAGS.map((t) => <option key={t} value={t}>{t}</option>)}
                  </select>
                </td>
                <td>
                  <select
                    className={"vivo-select-tag" + tagColorClass("situation", p.qualitySituation)}
                    value={p.qualitySituation || ""}
                    onChange={(e) => updateClassification(p.codigo, "qualitySituation", e.target.value)}
                  >
                    <option value="">—</option>
                    {QUALITY_SITUATION_TAGS.map((t) => <option key={t} value={t}>{t}</option>)}
                  </select>
                </td>
                <td>
                  <select
                    className={"vivo-select-tag" + tagColorClass("margin", p.qualityMargin)}
                    value={p.qualityMargin ?? ""}
                    onChange={(e) => updateClassification(p.codigo, "qualityMargin", e.target.value === "" ? "" : Number(e.target.value))}
                  >
                    <option value="">—</option>
                    {QUALITY_MARGIN_OPTIONS.map((m) => <option key={m} value={m}>{m > 0 ? `+${m}` : m}%</option>)}
                  </select>
                </td>
                <td>
                  <button
                    className={"vivo-toggle-btn" + (p.monitoringActive ? " is-on" : "")}
                    onClick={() => toggleMonitoring(p.codigo)}
                  >
                    {p.monitoringActive ? "Ativo" : "Ativar"}
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <PaginationControls page={page} setPage={setPage} totalPages={totalPages} totalCount={totalCount} />
    </div>
  );
}

// ====================== MONITORED TAB (placeholder) ======================
function MonitoredTab({ products, persistProducts }) {
  const [supplierFilter, setSupplierFilter] = useState("all");
  const [selected, setSelected] = useState({});
  const [scrollWidth, setScrollWidth] = useState(0);
  const topScrollRef = React.useRef(null);
  const tableWrapRef = React.useRef(null);

  const measureScrollWidth = useCallback(() => {
    if (tableWrapRef.current) setScrollWidth(tableWrapRef.current.scrollWidth);
  }, []);

  useEffect(() => {
    measureScrollWidth();
    window.addEventListener("resize", measureScrollWidth);
    return () => window.removeEventListener("resize", measureScrollWidth);
  }, [measureScrollWidth]);

  function syncFromTop(e) {
    if (tableWrapRef.current) tableWrapRef.current.scrollLeft = e.target.scrollLeft;
  }
  function syncFromTable(e) {
    if (topScrollRef.current) topScrollRef.current.scrollLeft = e.target.scrollLeft;
  }

  const list = useMemo(() => Object.values(products), [products]);
  const allMonitored = useMemo(
    () => list
      .filter((p) => p.monitoringActive)
      .map((p) => ({ ...p, quality: getStockQuality(p), monitoring: getMonitoringStatus(p) }))
      .sort((a, b) => (b.monitoringStartedAt || 0) - (a.monitoringStartedAt || 0)),
    [list]
  );

  const suppliers = useMemo(() => {
    const s = new Set(allMonitored.map((p) => p.fornecedor));
    return Array.from(s).sort();
  }, [allMonitored]);

  const monitored = useMemo(() => {
    if (supplierFilter === "all") return allMonitored;
    return allMonitored.filter((p) => p.fornecedor === supplierFilter);
  }, [allMonitored, supplierFilter]);

  // Paginação só afeta a renderização. Seleção em massa continua sobre toda a
  // lista filtrada (monitored), não só a página visível.
  const { page, setPage, totalPages, pageItems, totalCount } = usePagination(monitored);

  useEffect(() => {
    const id = setTimeout(measureScrollWidth, 0);
    return () => clearTimeout(id);
  }, [pageItems, measureScrollWidth]);

  function toggleSelect(codigo) {
    setSelected((s) => ({ ...s, [codigo]: !s[codigo] }));
  }
  function selectAllVisible() {
    const next = { ...selected };
    monitored.forEach((p) => { next[p.codigo] = true; });
    setSelected(next);
  }
  function clearSelection() {
    setSelected({});
  }
  const selectedCount = monitored.filter((p) => selected[p.codigo]).length;

  async function deactivateMonitoring(codigo) {
    const next = { ...products };
    const p = next[codigo];
    next[codigo] = { ...p, monitoringActive: false, monitoringStartedAt: null, monitoringStartV30: null };
    await persistProducts(next);
  }

  async function saveNotes(codigo, text) {
    const next = { ...products };
    next[codigo] = { ...next[codigo], monitoringNotes: text };
    await persistProducts(next);
  }

  async function deactivateSelected() {
    if (selectedCount === 0) return;
    const next = { ...products };
    for (const p of monitored) {
      if (!selected[p.codigo]) continue;
      next[p.codigo] = { ...next[p.codigo], monitoringActive: false, monitoringStartedAt: null, monitoringStartV30: null };
    }
    await persistProducts(next);
    setSelected({});
  }

  if (allMonitored.length === 0) {
    return (
      <div className="vivo-page">
        <header className="vivo-page-head">
          <h1>Monitorados</h1>
          <p>Itens com o monitoramento ativado na aba Qualidade de Estoque.</p>
        </header>
        <EmptyState icon={AlertTriangle} title="Nenhum item em monitoramento" text="Vá em Qualidade de Estoque → Análise e clique em “Ativar” no item que quiser acompanhar." />
      </div>
    );
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Monitorados</h1>
        <p>{allMonitored.length} itens em acompanhamento. Comparação entre as vendas dos últimos 30 dias no início do monitoramento e o valor atual.</p>
      </header>

      <div className="vivo-toolbar">
        <select className="vivo-select" value={supplierFilter} onChange={(e) => setSupplierFilter(e.target.value)}>
          <option value="all">Todas as marcas/fornecedores</option>
          {suppliers.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      </div>

      <div className="vivo-action-row">
        <button className="vivo-btn vivo-btn-ghost" onClick={selectAllVisible}>Selecionar todos visíveis</button>
        <button className="vivo-btn vivo-btn-ghost" onClick={clearSelection}>Limpar seleção</button>
        <div className="vivo-action-spacer" />
        <span className="vivo-action-summary">{selectedCount} selecionado(s)</span>
        <button className="vivo-btn vivo-btn-danger" onClick={deactivateSelected} disabled={selectedCount === 0}>
          Desativar monitoramento
        </button>
      </div>

      <div className="vivo-top-scrollbar" ref={topScrollRef} onScroll={syncFromTop}>
        <div style={{ width: scrollWidth, height: 1 }} />
      </div>

      <div className="vivo-table-wrap vivo-table-wrap-sticky vivo-card" ref={tableWrapRef} onScroll={syncFromTable}>
        <table className="vivo-table vivo-table-compact">
          <thead>
            <tr>
              <th></th>
              <th>Item</th>
              <th>Fornecedor</th>
              <th className="mono">Código</th>
              <th className="num">Estoque</th>
              <th className="num">Tempo de estoque</th>
              <th>Desde</th>
              <th className="num">Dias</th>
              <th className="num">V30D Início</th>
              <th className="num">V30D Atual</th>
              <th>Situação</th>
              <th>Motivo</th>
              <th>Ação</th>
              <th>Margem MC</th>
              <th>Observações</th>
              <th>Desativar</th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((p) => (
              <tr key={p.codigo} className={selected[p.codigo] ? "is-selected" : ""}>
                <td>
                  <input type="checkbox" checked={!!selected[p.codigo]} onChange={() => toggleSelect(p.codigo)} />
                </td>
                <ItemCell name={p.item} />
                <td>{p.fornecedor}</td>
                <td className="mono">{p.codigo}</td>
                <td className="num mono">{fmtNumber(p.estoque)}</td>
                <td className="num mono">{p.quality.stockDays === Infinity ? "sem vendas" : `${Math.round(p.quality.stockDays)}d`}</td>
                <td className="mono">{p.monitoringStartedAt ? fmtDateShort(p.monitoringStartedAt) : "—"}</td>
                <td className="num mono">{p.monitoring.days !== null ? p.monitoring.days : "—"}</td>
                <td className="num mono">{p.monitoring.startV30 ?? "—"}</td>
                <td className="num mono">{p.monitoring.currentV30 ?? "—"}</td>
                <td><MonitoringStatusBadge status={p.monitoring.status} deltaPct={p.monitoring.deltaPct} /></td>
                <td>{p.qualityReason || "—"}</td>
                <td>{p.qualitySituation || "—"}</td>
                <td>{p.qualityMargin !== undefined && p.qualityMargin !== null && p.qualityMargin !== "" ? `${p.qualityMargin > 0 ? "+" : ""}${p.qualityMargin}%` : "—"}</td>
                <td>
                  <NotesField
                    value={p.monitoringNotes || ""}
                    onSave={(text) => saveNotes(p.codigo, text)}
                  />
                </td>
                <td>
                  <button className="vivo-toggle-btn vivo-toggle-btn-off" onClick={() => deactivateMonitoring(p.codigo)}>
                    Desativar
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <PaginationControls page={page} setPage={setPage} totalPages={totalPages} totalCount={totalCount} />
    </div>
  );
}

// Campo de observações com estado local: salva no blur (ou Enter), não a cada tecla digitada,
// pra digitação ficar fluida sem disparar persistência a cada caractere.
function NotesField({ value, onSave }) {
  const [draft, setDraft] = useState(value);

  useEffect(() => {
    setDraft(value);
  }, [value]);

  function commit() {
    if (draft !== value) onSave(draft);
  }

  return (
    <input
      type="text"
      className="vivo-notes-input"
      placeholder="Observação…"
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => { if (e.key === "Enter") { commit(); e.target.blur(); } }}
    />
  );
}


function MonitoringStatusBadge({ status, deltaPct }) {
  const pctLabel = deltaPct !== null && deltaPct !== undefined ? ` (${deltaPct > 0 ? "+" : ""}${deltaPct.toFixed(0)}%)` : "";
  if (status === "acelerando") return <span className="vivo-monitor-badge vivo-monitor-up"><TrendingUp size={13} /> Acelerando{pctLabel}</span>;
  if (status === "caindo") return <span className="vivo-monitor-badge vivo-monitor-down"><TrendingDown size={13} /> Caindo{pctLabel}</span>;
  if (status === "igual") return <span className="vivo-monitor-badge vivo-monitor-flat"><Minus size={13} /> Igual{pctLabel}</span>;
  return <span className="vivo-monitor-badge vivo-monitor-muted">Sem dados</span>;
}

// ====================== HISTORY TAB ======================
function HistoryTab({ history, persistHistory }) {
  if (history.length === 0) {
    return (
      <div className="vivo-page">
        <EmptyState icon={History} title="Nenhuma importação registrada" text="O histórico de cada relatório enviado aparecerá aqui." />
      </div>
    );
  }
  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Histórico de importações</h1>
        <p>Registro de cada relatório enviado, para acompanhar a evolução dos dados ao longo do tempo.</p>
      </header>
      <div className="vivo-table-wrap vivo-card">
        <table className="vivo-table">
          <thead>
            <tr>
              <th>Data</th>
              <th>Arquivo</th>
              <th className="num">Linhas lidas</th>
              <th className="num">Importadas</th>
              <th className="num">Ignoradas (sem fornecedor)</th>
              <th className="num">Novos produtos</th>
              <th className="num">Atualizados</th>
            </tr>
          </thead>
          <tbody>
            {history.map((h) => (
              <tr key={h.id}>
                <td>{fmtDate(h.ts)}</td>
                <td>{h.fileName}</td>
                <td className="num mono">{h.totalRows}</td>
                <td className="num mono">{h.importedRows}</td>
                <td className="num mono">{h.skippedNoSupplier}</td>
                <td className="num mono">{h.newCount}</td>
                <td className="num mono">{h.updatedCount}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// Painel de diagnóstico manual: mostra quantos produtos existem em cada espaço de
// armazenamento (compartilhado vs privado desta conta) e permite forçar a migração
// a qualquer momento, sem depender de detecção automática.
function StorageDiagnosticPanel({ diag, migrationState, onCheck, onMigrate, onExport, onImportFile, importState, onClearHistory }) {
  return (
    <div className="vivo-card vivo-diagnostic-panel">
      <div className="vivo-diagnostic-head">
        <span className="vivo-detail-label">Diagnóstico de armazenamento</span>
        <button className="vivo-btn vivo-btn-ghost" onClick={onCheck}>Verificar agora</button>
      </div>

      {diag.checked && (
        <div className="vivo-diagnostic-result">
          <span>Espaço <strong>compartilhado</strong> (visível no link publicado): <strong>{diag.sharedCount}</strong> produtos</span>
          <span>Espaço <strong>privado</strong> desta conta (aqui na conversa): <strong>{diag.privateCount}</strong> produtos</span>
        </div>
      )}

      {diag.checked && diag.privateCount > 0 && (
        <div className="vivo-diagnostic-action">
          <p className="vivo-settings-hint" style={{ margin: "8px 0" }}>
            {diag.sharedCount > 0
              ? "Atenção: já existem dados no espaço compartilhado. Migrar agora vai SOBRESCREVER esses dados com os do espaço privado."
              : "O espaço compartilhado está vazio. Migrar agora vai copiar os dados privados para lá."}
          </p>
          <button className="vivo-btn vivo-btn-primary" onClick={onMigrate} disabled={migrationState.running}>
            {migrationState.running ? "Migrando…" : "Forçar migração para o espaço compartilhado"}
          </button>
        </div>
      )}

      {migrationState.done && (
        <div className="vivo-diagnostic-success"><CheckCircle2 size={14} /> Migração concluída. Recarregue o link publicado para ver.</div>
      )}
      {migrationState.error && (
        <div className="vivo-diagnostic-error"><AlertTriangle size={14} /> {migrationState.error}</div>
      )}

      <div className="vivo-diagnostic-divider" />

      <div className="vivo-diagnostic-backup">
        <span className="vivo-detail-label" style={{ display: "block", marginBottom: 6 }}>
          Transferir dados entre esta conversa e o link publicado
        </span>
        <p className="vivo-settings-hint" style={{ margin: "0 0 10px" }}>
          Cada publicação tem seu próprio espaço de dados — migrar aqui não afeta o link já publicado.
          Exporte um arquivo aqui e importe-o dentro do link publicado para levar os dados até lá.
          Se a exportação completa travar (muitos produtos com histórico), use a resumida.
        </p>
        <div className="vivo-diagnostic-backup-actions">
          <button className="vivo-btn vivo-btn-secondary" onClick={() => onExport(false)}>
            Exportar resumido (recomendado)
          </button>
          <button className="vivo-btn vivo-btn-ghost" onClick={() => onExport(true)}>
            Exportar completo (com histórico)
          </button>
          <label className="vivo-btn vivo-btn-secondary">
            Importar arquivo de backup
            <input type="file" accept=".json" style={{ display: "none" }} onChange={onImportFile} />
          </label>
        </div>
        <p className="vivo-settings-hint" style={{ margin: "8px 0 0" }}>
          O resumido mantém estoque, custo, classificações e monitoramento — só não traz o histórico detalhado de cada importação anterior.
        </p>
        {importState?.progress && !importState?.done && (
          <div className="vivo-diagnostic-progress">{importState.progress}</div>
        )}
        {importState?.done && (
          <div className="vivo-diagnostic-success"><CheckCircle2 size={14} /> Backup importado com sucesso. A página será recarregada.</div>
        )}
        {importState?.error && (
          <div className="vivo-diagnostic-error"><AlertTriangle size={14} /> {importState.error}</div>
        )}
      </div>

      <div className="vivo-diagnostic-divider" />

      <div className="vivo-diagnostic-backup">
        <span className="vivo-detail-label" style={{ display: "block", marginBottom: 6 }}>
          Reduzir o tamanho dos dados
        </span>
        <p className="vivo-settings-hint" style={{ margin: "0 0 10px" }}>
          Remove o histórico detalhado de vendas acumulado de cada produto (usado nos gráficos de tendência em Performance).
          Estoque, custo, classificações e monitoramento não são afetados. Útil se o sistema estiver lento ou se exportações estiverem travando.
        </p>
        <button className="vivo-btn vivo-btn-danger" onClick={onClearHistory}>
          Limpar histórico de vendas de todos os produtos
        </button>
      </div>
    </div>
  );
}

// ====================== CLEANUP TAB ======================
function CleanupTab({ products, persistProducts, history, settings, orders, persistHistory, persistSettings, persistOrders }) {
  const [selected, setSelected] = useState({});
  const [showInactive, setShowInactive] = useState(false);
  const [diag, setDiag] = useState({ checked: false, sharedCount: 0, privateCount: 0 });
  const [migrationState, setMigrationState] = useState({ running: false, done: false, error: null });
  const [importState, setImportState] = useState({ done: false, error: null });

  async function checkStorage() {
    setDiag((d) => ({ ...d, checked: false }));
    const [sharedProducts, privateProducts] = await Promise.all([
      storageGet(K_PRODUCTS, {}),
      storageGetPrivate(K_PRODUCTS, {}),
    ]);
    setDiag({
      checked: true,
      sharedCount: Object.keys(sharedProducts).length,
      privateCount: Object.keys(privateProducts).length,
    });
  }

  async function forceMigration() {
    setMigrationState({ running: true, done: false, error: null });
    try {
      const [p, h, s, o] = await Promise.all([
        storageGetPrivate(K_PRODUCTS, {}),
        storageGetPrivate(K_HISTORY, []),
        storageGetPrivate(K_SETTINGS, { defaultCoverageDays: 30, supplierCoverageDays: {} }),
        storageGetPrivate(K_ORDERS, []),
      ]);
      await Promise.all([
        storageSet(K_PRODUCTS, p),
        storageSet(K_HISTORY, h),
        storageSet(K_SETTINGS, s),
        storageSet(K_ORDERS, o),
      ]);
      await persistProducts(p);
      setMigrationState({ running: false, done: true, error: null });
      await checkStorage();
    } catch (e) {
      setMigrationState({ running: false, done: false, error: "Não foi possível migrar. Tente novamente em alguns segundos." });
    }
  }

  function exportBackup(full) {
    try {
      let exportProducts = products;
      if (!full) {
        // Versão resumida: remove o histórico de vendas acumulado de cada produto,
        // que é o campo que mais cresce e mais pesa no arquivo. Mantém tudo o resto
        // (estoque, classificações, monitoramento) intacto.
        exportProducts = {};
        for (const codigo of Object.keys(products)) {
          const { salesHistory, ...rest } = products[codigo];
          exportProducts[codigo] = rest;
        }
      }
      const backup = {
        exportedAt: Date.now(),
        version: APP_VERSION,
        full,
        products: exportProducts,
        history,
        settings,
        orders,
      };
      const json = JSON.stringify(backup);
      const blob = new Blob([json], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `vivo-backup-${full ? "completo" : "resumido"}-${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (e) {
      console.error("Erro ao exportar backup", e);
      alert("Não foi possível gerar o arquivo de backup. Tente a opção resumida, que é mais leve.");
    }
  }

  function importBackupFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportState({ done: false, error: null, progress: "Lendo arquivo…" });
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const backup = JSON.parse(ev.target.result);
        if (!backup.products) throw new Error("Arquivo de backup inválido.");

        // Grava uma chave de cada vez (em vez de tudo simultâneo) para reduzir o
        // pico de carga no navegador/armazenamento com volumes grandes de dados.
        setImportState((s) => ({ ...s, progress: "Salvando produtos…" }));
        await persistProducts(backup.products || {});

        setImportState((s) => ({ ...s, progress: "Salvando histórico de importações…" }));
        await persistHistory(backup.history || []);

        setImportState((s) => ({ ...s, progress: "Salvando configurações…" }));
        await persistSettings(backup.settings || { defaultCoverageDays: 30, supplierCoverageDays: {} });

        setImportState((s) => ({ ...s, progress: "Salvando pedidos…" }));
        await persistOrders(backup.orders || []);

        setImportState({ done: true, error: null, progress: null });
        setTimeout(() => window.location.reload(), 1500);
      } catch (err) {
        setImportState({ done: false, error: "Não foi possível importar este arquivo. Verifique se é um backup válido.", progress: null });
      }
    };
    reader.onerror = () => setImportState({ done: false, error: "Não foi possível ler o arquivo.", progress: null });
    reader.readAsText(file);
  }

  function clearAllSalesHistory() {
    if (!window.confirm("Isso remove o histórico detalhado de vendas de todos os produtos (usado nos gráficos de tendência). Estoque, custo, classificações e monitoramento não são afetados. Continuar?")) return;
    const next = {};
    for (const codigo of Object.keys(products)) {
      const { salesHistory, ...rest } = products[codigo];
      next[codigo] = rest;
    }
    persistProducts(next);
  }

  const { lastImportTs, missing } = useMemo(() => getMissingFromLastImport(products, history), [products, history]);

  // Paginação só afeta a renderização da tabela de itens ausentes. Seleção em massa
  // continua sobre toda a lista "missing", não só a página visível.
  const { page, setPage, totalPages, pageItems, totalCount } = usePagination(missing);

  const inactiveList = useMemo(
    () => Object.values(products).filter((p) => p.inactive),
    [products]
  );

  function toggleSelect(codigo) {
    setSelected((s) => ({ ...s, [codigo]: !s[codigo] }));
  }
  function selectAll() {
    const next = {};
    missing.forEach((p) => { next[p.codigo] = true; });
    setSelected(next);
  }
  function clearSelection() {
    setSelected({});
  }
  const selectedCount = missing.filter((p) => selected[p.codigo]).length;

  async function markSelectedInactive() {
    if (selectedCount === 0) return;
    const next = { ...products };
    for (const p of missing) {
      if (!selected[p.codigo]) continue;
      next[p.codigo] = { ...next[p.codigo], inactive: true, inactivatedAt: Date.now() };
    }
    await persistProducts(next);
    setSelected({});
  }

  async function reactivate(codigo) {
    const next = { ...products };
    next[codigo] = { ...next[codigo], inactive: false, inactivatedAt: null };
    await persistProducts(next);
  }

  if (!history || history.length === 0) {
    return (
      <div className="vivo-page">
        <header className="vivo-page-head">
          <h1>Limpeza de itens</h1>
          <p>Identifica produtos que não vieram na importação mais recente, para marcá-los como inativos.</p>
        </header>
        <StorageDiagnosticPanel diag={diag} migrationState={migrationState} onCheck={checkStorage} onMigrate={forceMigration} onExport={exportBackup} onImportFile={importBackupFile} importState={importState} onClearHistory={clearAllSalesHistory} />
        <EmptyState icon={Trash2} title="Nenhuma importação registrada ainda" text="Importe ao menos um relatório para que o sistema tenha uma referência de comparação." />
      </div>
    );
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Limpeza de itens</h1>
        <p>
          Última importação: <strong>{fmtDate(lastImportTs)}</strong>. Os itens abaixo não apareceram em nenhuma das planilhas desse envio mais recente —
          provavelmente foram descontinuados ou inativados no seu ERP.
        </p>
      </header>

      <StorageDiagnosticPanel diag={diag} migrationState={migrationState} onCheck={checkStorage} onMigrate={forceMigration} onExport={exportBackup} onImportFile={importBackupFile} importState={importState} onClearHistory={clearAllSalesHistory} />

      {missing.length === 0 ? (
        <EmptyState icon={CheckCircle2} title="Nenhum item para limpar" text="Todos os produtos ativos apareceram na última importação. Tudo em dia." />
      ) : (
        <>
          <div className="vivo-alert vivo-alert-warning">
            <AlertTriangle size={16} />
            Marcar como inativo <strong>não exclui</strong> o produto: ele some das telas de Produtos, Pedidos, Performance e Qualidade de Estoque, mas o histórico continua guardado e pode ser reativado a qualquer momento.
          </div>

          <div className="vivo-action-row">
            <button className="vivo-btn vivo-btn-ghost" onClick={selectAll}>Selecionar todos ({missing.length})</button>
            <button className="vivo-btn vivo-btn-ghost" onClick={clearSelection}>Limpar seleção</button>
            <div className="vivo-action-spacer" />
            <span className="vivo-action-summary">{selectedCount} selecionado(s)</span>
            <button className="vivo-btn vivo-btn-danger" onClick={markSelectedInactive} disabled={selectedCount === 0}>
              Marcar como inativo
            </button>
          </div>

          <div className="vivo-table-wrap vivo-card">
            <table className="vivo-table vivo-table-compact">
              <thead>
                <tr>
                  <th></th>
                  <th>Item</th>
                  <th>Fornecedor</th>
                  <th className="mono">Código</th>
                  <th className="num">Estoque</th>
                  <th>Última atualização</th>
                </tr>
              </thead>
              <tbody>
                {pageItems.map((p) => (
                  <tr key={p.codigo} className={selected[p.codigo] ? "is-selected" : ""}>
                    <td>
                      <input type="checkbox" checked={!!selected[p.codigo]} onChange={() => toggleSelect(p.codigo)} />
                    </td>
                    <ItemCell name={p.item} />
                    <td>{p.fornecedor}</td>
                    <td className="mono">{p.codigo}</td>
                    <td className="num mono">{fmtNumber(p.estoque)}</td>
                    <td className="mono">{fmtDate(p.lastUpdated)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <PaginationControls page={page} setPage={setPage} totalPages={totalPages} totalCount={totalCount} />
        </>
      )}

      <div className="vivo-cleanup-inactive-section">
        <button className="vivo-btn vivo-btn-ghost" onClick={() => setShowInactive((s) => !s)}>
          {showInactive ? "Ocultar" : "Ver"} produtos já inativos ({inactiveList.length})
        </button>

        {showInactive && inactiveList.length > 0 && (
          <div className="vivo-table-wrap vivo-card" style={{ marginTop: 12 }}>
            <table className="vivo-table vivo-table-compact">
              <thead>
                <tr>
                  <th>Item</th>
                  <th>Fornecedor</th>
                  <th className="mono">Código</th>
                  <th>Inativado em</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {inactiveList.map((p) => (
                  <tr key={p.codigo}>
                    <ItemCell name={p.item} />
                    <td>{p.fornecedor}</td>
                    <td className="mono">{p.codigo}</td>
                    <td className="mono">{p.inactivatedAt ? fmtDate(p.inactivatedAt) : "—"}</td>
                    <td>
                      <button className="vivo-toggle-btn" onClick={() => reactivate(p.codigo)}>Reativar</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

// ====================== SHARED ======================
// ====================== PAGINATION ======================
const PAGE_SIZE = 50;

// Pagina uma lista já filtrada/ordenada. Reinicia para a página 1 sempre que o
// tamanho da lista muda (ex: aplicou um novo filtro) para não ficar numa página
// vazia. Renderizar só uma fatia por vez evita travar com listas de milhares de itens.
// Célula de item com truncamento e expansão ao clicar.
// Mostra "..." quando o texto não cabe, e expande/recolhe ao clicar.
function ItemCell({ name, className = "" }) {
  const [expanded, setExpanded] = useState(false);
  return (
    <td
      className={"vivo-item-cell" + (expanded ? " is-expanded" : "") + (className ? " " + className : "")}
      onClick={() => setExpanded((e) => !e)}
      title={expanded ? "Clique para recolher" : name}
      style={{ cursor: "pointer" }}
    >
      {name}
    </td>
  );
}

function usePagination(items, pageSize = PAGE_SIZE) {
  const [page, setPage] = useState(1);
  const totalPages = Math.max(1, Math.ceil(items.length / pageSize));

  useEffect(() => {
    setPage(1);
  }, [items.length]);

  const safePage = Math.min(page, totalPages);
  const start = (safePage - 1) * pageSize;
  const pageItems = useMemo(() => items.slice(start, start + pageSize), [items, start, pageSize]);

  return { page: safePage, setPage, totalPages, pageItems, totalCount: items.length };
}

function PaginationControls({ page, setPage, totalPages, totalCount, pageSize = PAGE_SIZE }) {
  if (totalCount === 0) return null;
  const start = (page - 1) * pageSize + 1;
  const end = Math.min(page * pageSize, totalCount);
  return (
    <div className="vivo-pagination">
      <span className="vivo-pagination-info">{start}–{end} de {totalCount}</span>
      <div className="vivo-pagination-buttons">
        <button className="vivo-btn vivo-btn-ghost" onClick={() => setPage(1)} disabled={page === 1}>«</button>
        <button className="vivo-btn vivo-btn-ghost" onClick={() => setPage((p) => Math.max(1, p - 1))} disabled={page === 1}>‹ Anterior</button>
        <span className="vivo-pagination-page">Página {page} de {totalPages}</span>
        <button className="vivo-btn vivo-btn-ghost" onClick={() => setPage((p) => Math.min(totalPages, p + 1))} disabled={page === totalPages}>Próxima ›</button>
        <button className="vivo-btn vivo-btn-ghost" onClick={() => setPage(totalPages)} disabled={page === totalPages}>»</button>
      </div>
    </div>
  );
}

// ====================== SHARED ======================
function EmptyState({ icon: Icon, title, text }) {
  return (
    <div className="vivo-empty">
      <Icon size={32} strokeWidth={1.5} />
      <h3>{title}</h3>
      <p>{text}</p>
    </div>
  );
}

// ====================== STYLES ======================
const VIVO_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

.vivo-root {
  --ink: #2b2620;
  --ink-soft: #6b6358;
  --paper: #faf7f0;
  --card: #ffffff;
  --line: #e6ddc9;
  --olive: #5b6b4f;
  --olive-dark: #41503a;
  --amber: #c1762f;
  --amber-soft: #f3e3cf;
  --rust: #a8472f;
  --rust-soft: #f7e4dd;
  --sidebar-bg: #232017;
  --sidebar-line: #38332a;

  display: flex;
  flex-direction: column;
  min-height: 100%;
  width: 100%;
  background: var(--paper);
  color: var(--ink);
  font-family: 'Inter', -apple-system, sans-serif;
  font-size: 13px;
  line-height: 1.4;
}

.vivo-root * { box-sizing: border-box; }

.vivo-loading {
  align-items: center;
  justify-content: center;
  flex-direction: column;
  gap: 10px;
  width: 100%;
  min-height: 480px;
}
.vivo-loading-mark {
  font-family: 'Fraunces', serif;
  font-size: 28px;
  letter-spacing: 0.08em;
  color: var(--olive-dark);
}
.vivo-loading-text { color: var(--ink-soft); font-size: 13px; }

.vivo-password-form {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 12px;
  margin-top: 18px;
  width: 260px;
}
.vivo-password-label { font-size: 13px; color: var(--ink-soft); text-align: center; margin: 0; }
.vivo-password-input {
  width: 100%;
  font-family: inherit;
  font-size: 14px;
  padding: 10px 14px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: var(--card);
  color: var(--ink);
  text-align: center;
}
.vivo-password-input:focus { outline: none; border-color: var(--olive-dark); }
.vivo-password-input.is-error { border-color: var(--rust); background: var(--rust-soft); }
.vivo-password-error { color: var(--rust); font-size: 12.5px; margin: 0; }
.vivo-password-btn { width: 100%; justify-content: center; }
.vivo-password-version { font-size: 10.5px; color: #b5ad9a; margin: 4px 0 0; }

/* ---------- Top Nav ---------- */
.vivo-topnav {
  background: var(--sidebar-bg);
  color: #e9e3d6;
  flex-shrink: 0;
}
.vivo-topnav-row {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 0 28px;
}
.vivo-topnav-main { height: 60px; border-bottom: 1px solid var(--sidebar-line); }

.vivo-brand {
  display: flex;
  align-items: center;
  gap: 9px;
  margin-right: 28px;
}
.vivo-brand-mark { color: var(--amber); font-size: 16px; }
.vivo-brand-name {
  font-family: 'Fraunces', serif;
  font-size: 16px;
  font-weight: 600;
  letter-spacing: 0.01em;
  color: #fbf8f0;
  line-height: 1.1;
}
.vivo-brand-sub { font-size: 10.5px; color: #a39c8a; margin-top: 1px; }

.vivo-main-tabs { display: flex; align-items: center; gap: 4px; height: 100%; }
.vivo-main-tab-btn {
  display: flex;
  align-items: center;
  gap: 8px;
  background: transparent;
  border: none;
  color: #c8c0ac;
  padding: 0 16px;
  height: 100%;
  font-size: 13.5px;
  font-weight: 500;
  font-family: inherit;
  cursor: pointer;
  position: relative;
  transition: color 0.15s;
}
.vivo-main-tab-btn:hover { color: #fbf8f0; }
.vivo-main-tab-btn.is-active { color: #fff; }
.vivo-main-tab-btn.is-active::after {
  content: "";
  position: absolute;
  left: 16px;
  right: 16px;
  bottom: 0;
  height: 2px;
  background: var(--amber);
}

.vivo-topnav-spacer { flex: 1; }
.vivo-topnav-hint { font-size: 11.5px; color: #847d6c; white-space: nowrap; }

.vivo-sub-tabs { height: 46px; gap: 6px; }
.vivo-sub-tab-btn {
  display: flex;
  align-items: center;
  gap: 7px;
  background: transparent;
  border: none;
  color: #b3ab97;
  padding: 7px 13px;
  border-radius: 7px;
  font-size: 13px;
  font-family: inherit;
  cursor: pointer;
  transition: background 0.15s, color 0.15s;
}
.vivo-sub-tab-btn:hover { background: #2e2a20; color: #fbf8f0; }
.vivo-sub-tab-btn.is-active { background: #383228; color: #fff; }
.vivo-sub-tab-count {
  background: rgba(255,255,255,0.14);
  border-radius: 100px;
  font-size: 10.5px;
  padding: 1px 6px;
  font-family: 'JetBrains Mono', monospace;
}

/* ---------- Main ---------- */
.vivo-main { flex: 1; min-width: 0; overflow-x: auto; }
.vivo-page { padding: 30px 36px 50px; max-width: 1200px; }
.vivo-page-head { margin-bottom: 22px; }
.vivo-page-head h1 {
  font-family: 'Fraunces', serif;
  font-size: 26px;
  font-weight: 600;
  margin: 0 0 5px;
  color: var(--ink);
}
.vivo-page-head p { color: var(--ink-soft); margin: 0; font-size: 13.5px; max-width: 640px; }

/* ---------- Dropzone ---------- */
.vivo-dropzone {
  border: 1.5px dashed var(--line);
  border-radius: 12px;
  background: var(--card);
  padding: 40px 24px;
  text-align: center;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 6px;
  color: var(--olive-dark);
  transition: border-color 0.15s, background 0.15s;
  max-width: 560px;
}
.vivo-dropzone.is-drag { border-color: var(--amber); background: var(--amber-soft); }
.vivo-dropzone-title { font-weight: 600; font-size: 15px; margin-top: 6px; color: var(--ink); }
.vivo-dropzone-sub { color: var(--ink-soft); font-size: 12px; margin: 0; }
.vivo-dropzone-hint { color: var(--ink-soft); font-size: 12px; margin-top: 10px; max-width: 380px; }

/* ---------- Buttons ---------- */
.vivo-btn {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  border-radius: 7px;
  padding: 8px 14px;
  font-size: 13px;
  font-weight: 500;
  font-family: inherit;
  cursor: pointer;
  border: 1px solid transparent;
  white-space: nowrap;
  transition: opacity 0.15s, background 0.15s;
}
.vivo-btn:disabled { opacity: 0.4; cursor: not-allowed; }
.vivo-btn-primary { background: var(--olive-dark); color: #fff; }
.vivo-btn-primary:hover:not(:disabled) { background: var(--olive); }
.vivo-btn-secondary { background: var(--amber-soft); color: #6b4419; border-color: #e6cda4; }
.vivo-btn-secondary:hover:not(:disabled) { background: #ecd6b4; }
.vivo-btn-ghost { background: transparent; color: var(--ink-soft); border-color: var(--line); }
.vivo-btn-ghost:hover:not(:disabled) { background: #f1ece0; color: var(--ink); }
.vivo-btn-danger { background: var(--rust-soft); color: var(--rust); border-color: #e3b4a3; }
.vivo-btn-danger:hover:not(:disabled) { background: #f3c4b3; }

/* ---------- Alerts ---------- */
.vivo-alert {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 10px 14px;
  border-radius: 8px;
  font-size: 13px;
  margin-top: 16px;
  max-width: 560px;
}
.vivo-alert-error { background: var(--rust-soft); color: #7a2c19; }
.vivo-alert-warning { background: var(--amber-soft); color: #6b4419; max-width: 720px; align-items: flex-start; }
.vivo-alert-warning strong { font-weight: 600; }

.vivo-migration-banner {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 12px 24px;
  background: var(--amber-soft);
  color: #6b4419;
  font-size: 13px;
  border-bottom: 1px solid #e6cda4;
}
.vivo-migration-banner span { flex: 1; }
.vivo-migration-success { background: #e7ede2; color: var(--olive-dark); }
.vivo-migration-success span { flex: initial; }
.vivo-migration-error { color: var(--rust); font-size: 12px; white-space: nowrap; }

.vivo-diagnostic-panel { padding: 16px 20px; margin-bottom: 18px; max-width: 640px; }
.vivo-diagnostic-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px; }
.vivo-diagnostic-result { display: flex; flex-direction: column; gap: 4px; margin-top: 10px; font-size: 13px; color: var(--ink); }
.vivo-diagnostic-action { margin-top: 6px; }
.vivo-diagnostic-success {
  display: flex; align-items: center; gap: 6px; margin-top: 10px;
  font-size: 12.5px; color: var(--olive-dark);
}
.vivo-diagnostic-error {
  display: flex; align-items: center; gap: 6px; margin-top: 10px;
  font-size: 12.5px; color: var(--rust);
}
.vivo-diagnostic-progress {
  margin-top: 10px;
  font-size: 12.5px;
  color: var(--ink-soft);
  font-family: 'JetBrains Mono', monospace;
}

.vivo-diagnostic-divider { height: 1px; background: var(--line); margin: 16px 0; }
.vivo-diagnostic-backup-actions { display: flex; gap: 10px; flex-wrap: wrap; }

.vivo-cleanup-inactive-section { margin-top: 28px; }

.vivo-pagination {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-top: 12px;
  flex-wrap: wrap;
}
.vivo-pagination-info { font-size: 12.5px; color: var(--ink-soft); font-family: 'JetBrains Mono', monospace; }
.vivo-pagination-buttons { display: flex; align-items: center; gap: 6px; }
.vivo-pagination-page { font-size: 12.5px; color: var(--ink-soft); padding: 0 4px; white-space: nowrap; }

/* ---------- Cards ---------- */
.vivo-card {
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 12px;
}
.vivo-preview { margin-top: 22px; padding: 18px 20px; max-width: 920px; }
.vivo-preview-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; margin-bottom: 14px; }
.vivo-preview-head h3 { font-size: 14px; font-weight: 600; margin: 0 0 5px; }
.vivo-preview-stats { display: flex; gap: 14px; font-size: 12.5px; color: var(--ink-soft); flex-wrap: wrap; }
.vivo-preview-stats .ok { color: var(--olive-dark); display: inline-flex; align-items: center; gap: 4px; }
.vivo-preview-stats .warn { color: var(--rust); display: inline-flex; align-items: center; gap: 4px; }

.vivo-pending-files { display: flex; flex-direction: column; gap: 8px; }
.vivo-pending-file {
  display: flex;
  align-items: center;
  gap: 16px;
  flex-wrap: wrap;
  background: #fbf8f0;
  border: 1px solid var(--line);
  border-radius: 8px;
  padding: 10px 14px;
}
.vivo-pending-file-name { display: flex; align-items: center; gap: 7px; font-size: 13px; font-weight: 500; color: var(--ink); min-width: 180px; }
.vivo-pending-file-stats { display: flex; gap: 12px; font-size: 12px; color: var(--ink-soft); flex-wrap: wrap; }
.vivo-pending-file-stats .ok { color: var(--olive-dark); display: inline-flex; align-items: center; gap: 4px; }
.vivo-pending-file-stats .warn { color: var(--rust); display: inline-flex; align-items: center; gap: 4px; }
.vivo-pending-file-error { color: var(--rust); font-size: 12.5px; display: inline-flex; align-items: center; gap: 5px; }
.vivo-period-select { display: flex; align-items: center; gap: 7px; font-size: 12.5px; color: var(--ink-soft); margin-left: auto; }
.vivo-period-select .vivo-select { padding: 5px 8px; font-size: 12.5px; }

/* ---------- Tables ---------- */
.vivo-table-wrap { overflow-x: auto; }
.vivo-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.vivo-table thead th {
  text-align: left;
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  color: var(--ink-soft);
  font-weight: 600;
  padding: 7px 8px;
  border-bottom: 1px solid var(--line);
  white-space: nowrap;
}
.vivo-table th.num, .vivo-table td.num { text-align: right; }
.vivo-table tbody td {
  padding: 5px 8px;
  border-bottom: 1px solid #f0ebdd;
  vertical-align: middle;
  font-size: 12px;
  white-space: nowrap;
}
.vivo-table tbody tr:hover { background: #fbf8f0; }
.vivo-table tbody tr.is-low { background: var(--rust-soft); }
.vivo-table tbody tr.is-low:hover { background: #f3d8cd; }
.vivo-table tbody tr.is-selected { background: var(--amber-soft); }
.vivo-table .mono { font-family: 'JetBrains Mono', monospace; font-size: 11px; color: var(--ink-soft); }
.vivo-item-cell {
  font-weight: 500;
  color: var(--ink);
  max-width: 220px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.vivo-item-cell.is-expanded {
  white-space: normal;
  overflow: visible;
  max-width: 320px;
  word-break: break-word;
}
.vivo-suggested { color: var(--amber); font-weight: 600; }
.vivo-trend-up { color: var(--olive-dark); font-weight: 600; }
.vivo-trend-down { color: var(--rust); font-weight: 600; }

.vivo-trend-summary { display: flex; align-items: center; gap: 8px; margin-bottom: 16px; flex-wrap: wrap; }
.vivo-trend-chip {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  border: 1px solid var(--line);
  background: var(--card);
  border-radius: 100px;
  padding: 6px 13px;
  font-size: 12.5px;
  font-family: inherit;
  cursor: pointer;
  color: var(--ink-soft);
  transition: background 0.15s, border-color 0.15s;
}
.vivo-trend-chip:hover { background: #f1ece0; }
.vivo-trend-chip-up.is-active { background: var(--olive-dark); color: #fff; border-color: var(--olive-dark); }
.vivo-trend-chip-down.is-active { background: var(--rust); color: #fff; border-color: var(--rust); }
.vivo-trend-chip-flat.is-active { background: var(--ink-soft); color: #fff; border-color: var(--ink-soft); }
.vivo-trend-chip-muted { font-size: 12px; color: var(--ink-soft); padding: 6px 4px; }

.vivo-badge {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  font-size: 12px;
  font-weight: 500;
  padding: 3px 9px;
  border-radius: 100px;
}
.vivo-badge-up { background: #e7ede2; color: var(--olive-dark); }
.vivo-badge-down { background: var(--rust-soft); color: var(--rust); }
.vivo-badge-flat { background: #f0ebdd; color: var(--ink-soft); }
.vivo-badge-muted { background: #f0ebdd; color: var(--ink-soft); font-style: italic; }

.vivo-range-filter {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12.5px;
  color: var(--ink-soft);
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 7px;
  padding: 6px 10px;
}
.vivo-range-filter .vivo-input-mini { width: 56px; }

.vivo-quality-summary {
  display: flex;
  align-items: center;
  gap: 10px;
  font-size: 13px;
  color: var(--ink-soft);
  margin-bottom: 14px;
}
.vivo-quality-summary strong { color: var(--ink); }

.vivo-capital-chart {
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 10px;
  padding: 12px 16px 8px;
  margin-bottom: 14px;
}
.vivo-capital-chart-empty {
  display: flex;
  align-items: center;
  font-size: 12px;
  color: var(--ink-soft);
  padding: 14px 16px;
  min-height: 56px;
}
.vivo-capital-chart-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  gap: 12px;
  margin-bottom: 6px;
  flex-wrap: wrap;
}
.vivo-capital-chart-label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-soft); font-weight: 600; }
.vivo-capital-chart-supplier { color: var(--amber); }
.vivo-capital-chart-summary { display: flex; align-items: baseline; gap: 10px; }
.vivo-capital-chart-delta { font-size: 12px; font-family: 'JetBrains Mono', monospace; }
.vivo-capital-chart-svg { width: 100%; height: 72px; display: block; }

.vivo-chart-tooltip {
  position: absolute;
  background: var(--sidebar-bg);
  color: #fbf8f0;
  border-radius: 7px;
  padding: 7px 10px;
  font-size: 12px;
  pointer-events: none;
  white-space: nowrap;
  z-index: 10;
  box-shadow: 0 2px 8px rgba(0,0,0,0.18);
}
.vivo-chart-tooltip-date { font-size: 10.5px; color: #a39c8a; margin-bottom: 2px; }
.vivo-chart-tooltip-value { font-weight: 600; font-size: 13px; font-family: 'JetBrains Mono', monospace; }
.vivo-chart-tooltip-delta { font-size: 11px; margin-top: 2px; font-family: 'JetBrains Mono', monospace; }

.vivo-stockdays { color: var(--rust); font-weight: 600; }

.vivo-select-tag { font-size: 12px; padding: 5px 7px; min-width: 132px; }
.vivo-table-quality td { vertical-align: middle; }

/* Tabela compacta: menos padding, fonte um pouco menor, pra caber mais colunas/linhas na tela */
.vivo-table-compact th { padding: 5px 7px; font-size: 9.5px; }
.vivo-table-compact td { padding: 4px 7px; font-size: 11.5px; }
.vivo-table-compact .vivo-select-tag { min-width: 100px; padding: 3px 5px; font-size: 11px; }
.vivo-table-compact .vivo-item-cell { max-width: 180px; font-size: 11.5px; }
.vivo-table-compact td:nth-child(2) { font-size: 11.5px; }

/* Wrapper com cabeçalho fixo (sticky) e barra de rolagem própria no topo e no fundo */
.vivo-table-wrap-sticky {
  max-height: 70vh;
  overflow: auto;
  position: relative;
}
.vivo-table-wrap-sticky table thead th {
  position: sticky;
  top: 0;
  background: var(--card);
  z-index: 2;
  box-shadow: 0 1px 0 var(--line);
}

.vivo-capital-danger { color: var(--rust); font-weight: 600; }

.vivo-top-scrollbar {
  overflow-x: auto;
  overflow-y: hidden;
  height: 14px;
  margin-bottom: 2px;
}
.vivo-top-scrollbar::-webkit-scrollbar { height: 10px; }
.vivo-top-scrollbar::-webkit-scrollbar-thumb { background: #d8cdb0; border-radius: 100px; }
.vivo-top-scrollbar::-webkit-scrollbar-track { background: #f0ebdd; border-radius: 100px; }

/* Selects de classificação coloridos por valor escolhido */
.vivo-select-tag {
  border: 1px solid var(--line);
  border-radius: 6px;
  background: var(--card);
  color: var(--ink);
  font-family: inherit;
  cursor: pointer;
}
.vivo-tag-blue { background: #e7eef6; border-color: #b9cee4; color: #2f5277; }
.vivo-tag-teal { background: #e3eee9; border-color: #aed4c4; color: #2c6650; }
.vivo-tag-amber { background: var(--amber-soft); border-color: #e6cda4; color: #6b4419; }
.vivo-tag-rust { background: var(--rust-soft); border-color: #e3b4a3; color: #7a2c19; }
.vivo-tag-olive { background: #eef1e9; border-color: #c7d4b8; color: var(--olive-dark); }
.vivo-tag-neutral { background: #f0ebdd; border-color: var(--line); color: var(--ink-soft); }

.vivo-toggle-btn {
  font-family: inherit;
  font-size: 11.5px;
  font-weight: 600;
  padding: 5px 14px;
  border-radius: 100px;
  border: 1px solid var(--line);
  background: var(--card);
  color: var(--ink-soft);
  cursor: pointer;
  transition: background 0.15s, border-color 0.15s, color 0.15s;
  white-space: nowrap;
}
.vivo-toggle-btn:hover { background: #f1ece0; }
.vivo-toggle-btn.is-on {
  background: var(--olive-dark);
  border-color: var(--olive-dark);
  color: #fff;
}
.vivo-toggle-btn.is-on:hover { background: var(--olive); }
.vivo-toggle-btn-off { color: var(--rust); border-color: #e3b4a3; }
.vivo-toggle-btn-off:hover { background: var(--rust-soft); }

.vivo-notes-input {
  width: 180px;
  font-family: inherit;
  font-size: 12px;
  padding: 5px 8px;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: var(--card);
  color: var(--ink);
}
.vivo-notes-input:focus {
  outline: none;
  border-color: var(--amber);
  background: #fffdf8;
}
.vivo-notes-input::placeholder { color: #b5ad9a; }

.vivo-monitor-badge {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  font-size: 12px;
  font-weight: 600;
  padding: 4px 10px;
  border-radius: 100px;
  white-space: nowrap;
}
.vivo-monitor-up { background: #e1ecd9; color: #3c6b2a; }
.vivo-monitor-down { background: var(--rust-soft); color: var(--rust); }
.vivo-monitor-flat { background: #f0ebdd; color: var(--ink-soft); }
.vivo-monitor-muted { background: #f0ebdd; color: var(--ink-soft); font-style: italic; font-size: 11.5px; font-weight: 400; padding: 4px 10px; border-radius: 100px; }
.vivo-table-more { padding: 10px 12px; font-size: 12px; color: var(--ink-soft); }
.vivo-table-empty { text-align: center; color: var(--ink-soft); padding: 30px; }

.vivo-icon-btn {
  background: transparent;
  border: none;
  color: var(--ink-soft);
  cursor: pointer;
  display: flex;
  align-items: center;
  padding: 2px;
  border-radius: 4px;
}
.vivo-icon-btn:hover { color: var(--ink); background: #f0ebdd; }
.vivo-icon-danger:hover { color: var(--rust); background: var(--rust-soft); }

.vivo-row-detail td { background: #fbf8f0; padding: 0; }
.vivo-detail { padding: 16px 20px 20px 44px; }
.vivo-detail-grid { display: flex; gap: 32px; margin-bottom: 16px; flex-wrap: wrap; }
.vivo-detail-label { display: block; font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-soft); margin-bottom: 3px; }
.vivo-detail-value { font-size: 14px; font-weight: 500; color: var(--ink); }
.vivo-detail-muted { font-size: 12px; color: var(--ink-soft); font-weight: 400; }
.vivo-detail-history { max-width: 480px; }
.vivo-mini-table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 8px; }
.vivo-mini-table th { text-align: left; color: var(--ink-soft); font-weight: 500; padding: 5px 8px; border-bottom: 1px solid var(--line); }
.vivo-mini-table th.num, .vivo-mini-table td.num { text-align: right; }
.vivo-mini-table td { padding: 5px 8px; font-family: 'JetBrains Mono', monospace; border-bottom: 1px solid #f0ebdd; }

/* ---------- Toolbar / filters ---------- */
.vivo-toolbar { display: flex; align-items: center; gap: 10px; margin-bottom: 16px; flex-wrap: wrap; }
.vivo-search {
  display: flex;
  align-items: center;
  gap: 7px;
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 7px;
  padding: 7px 11px;
  color: var(--ink-soft);
  min-width: 220px;
}
.vivo-search input { border: none; outline: none; background: transparent; font-size: 13px; flex: 1; color: var(--ink); font-family: inherit; }
.vivo-select {
  border: 1px solid var(--line);
  background: var(--card);
  border-radius: 7px;
  padding: 7px 10px;
  font-size: 13px;
  color: var(--ink);
  font-family: inherit;
}
.vivo-checkbox { display: inline-flex; align-items: center; gap: 6px; font-size: 13px; color: var(--ink-soft); cursor: pointer; }

.vivo-input, .vivo-input-mini {
  border: 1px solid var(--line);
  border-radius: 6px;
  padding: 6px 9px;
  font-size: 13px;
  font-family: 'JetBrains Mono', monospace;
  color: var(--ink);
  background: var(--card);
}
.vivo-input { display: block; margin-top: 6px; width: 140px; }
.vivo-input-mini { width: 68px; padding: 5px 7px; }

.vivo-settings-panel { padding: 14px 18px; margin-bottom: 16px; max-width: 460px; font-size: 13px; }
.vivo-settings-hint { color: var(--ink-soft); font-size: 12px; margin: 10px 0 0; }
.vivo-supplier-coverage { max-height: 260px; overflow-y: auto; padding-right: 4px; }
.vivo-supplier-coverage-row { display: flex; align-items: center; gap: 10px; padding: 5px 0; border-bottom: 1px solid #f0ebdd; }
.vivo-supplier-coverage-name { flex: 1; font-size: 13px; color: var(--ink); }
.vivo-supplier-coverage-unit { font-size: 12px; color: var(--ink-soft); }
.vivo-coverage-tag {
  display: inline-block;
  margin-top: 3px;
  font-size: 10px;
  color: var(--ink-soft);
  background: #f0ebdd;
  border-radius: 4px;
  padding: 1px 5px;
}
.vivo-source-tag {
  font-size: 9px;
  color: var(--olive-dark);
  margin-left: 3px;
  font-weight: 600;
  text-transform: uppercase;
}

.vivo-action-row { display: flex; align-items: center; gap: 10px; margin-bottom: 14px; flex-wrap: wrap; }
.vivo-action-spacer { flex: 1; }
.vivo-action-summary { font-size: 12.5px; color: var(--ink-soft); font-family: 'JetBrains Mono', monospace; }
.vivo-action-warning { color: var(--rust); }
.vivo-no-price { color: var(--rust); font-family: inherit; }

.vivo-supplier-totals { padding: 16px 20px; margin-bottom: 16px; max-width: 560px; }
.vivo-supplier-totals tfoot td { border-top: 1.5px solid var(--line); border-bottom: none; padding-top: 8px; }

.vivo-order-card-suppliers { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 8px; }
.vivo-pill-supplier { background: #eef1e9; color: var(--olive-dark); }

/* ---------- Saved orders ---------- */
.vivo-saved-orders { margin-top: 28px; }
.vivo-saved-orders h3 { font-family: 'Fraunces', serif; font-size: 16px; font-weight: 600; margin: 0 0 12px; }
.vivo-order-card { padding: 12px 16px; margin-bottom: 8px; }
.vivo-order-card-head { display: flex; justify-content: space-between; font-size: 12px; color: var(--ink-soft); margin-bottom: 8px; }
.vivo-order-card-items { display: flex; flex-wrap: wrap; gap: 6px; }
.vivo-pill { background: var(--amber-soft); color: #6b4419; font-size: 12px; padding: 3px 9px; border-radius: 100px; }
.vivo-pill-muted { background: #f0ebdd; color: var(--ink-soft); }

/* ---------- Empty state ---------- */
.vivo-empty {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 8px;
  text-align: center;
  padding: 70px 20px;
  color: var(--ink-soft);
  max-width: 360px;
  margin: 0 auto;
}
.vivo-empty h3 { font-family: 'Fraunces', serif; font-size: 18px; color: var(--ink); margin: 6px 0 0; font-weight: 600; }
.vivo-empty p { font-size: 13px; margin: 0; }

@media (max-width: 760px) {
  .vivo-topnav-row { padding: 0 14px; overflow-x: auto; }
  .vivo-topnav-hint { display: none; }
  .vivo-brand { margin-right: 16px; }
  .vivo-page { padding: 20px 16px 40px; }
}
`;

