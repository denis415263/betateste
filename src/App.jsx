import React, { useState, useEffect, useMemo, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  Upload, Package, ShoppingCart, Search,
  ChevronDown, ChevronRight, AlertTriangle, CheckCircle2,
  Settings, FileSpreadsheet, History, Trash2, Box, TrendingUp, TrendingDown, Minus
} from "lucide-react";

const K_PRODUCTS = "vivo:products";
const K_HISTORY = "vivo:history";
const K_SETTINGS = "vivo:settings";
const K_ORDERS = "vivo:orders";

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

// Custo de nota (col K) — usado em pedidos de compra e geração de pedido
function custoNota(p) { return p.precoCusto || 0; }

// Custo líquido (col S) — usado em relatórios financeiros, capital imobilizado, estoque R$
function custoLiq(p) { return p.custoLiquido || p.precoCusto || 0; }

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

function dayKey(ts) {
  const d = new Date(ts);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function dedupeByCalendarDay(points) {
  const byDay = {};
  for (const p of points) {
    const key = dayKey(p.ts);
    if (!byDay[key] || p.ts > byDay[key].ts) byDay[key] = p;
  }
  return Object.values(byDay).sort((a, b) => a.ts - b.ts);
}

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
    if (days <= 0) continue;
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

const STOCK_QUALITY_THRESHOLD_DAYS = 90;

function getStockQuality(product) {
  const avgDay = (product.vendas30 || 0) / 30;
  const estoque = product.estoque || 0;
  const stockDays = avgDay > 0 ? estoque / avgDay : (estoque > 0 ? Infinity : 0);
  const capitalImobilizado = estoque * custoLiq(product);
  const isExcess = stockDays > STOCK_QUALITY_THRESHOLD_DAYS;
  return { avgDay, stockDays, capitalImobilizado, isExcess };
}

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

const IS_LOCALHOST = typeof window !== "undefined" &&
  (window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1");

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
  if (IS_LOCALHOST) {
    try {
      const val = localStorage.getItem("vivo:" + key);
      return val ? JSON.parse(val) : fallback;
    } catch (e) {
      return fallback;
    }
  }
  try {
    const res = await withTimeout(window.storage.get(key, true), 5000, null);
    if (!res) return fallback;
    return JSON.parse(res.value);
  } catch (e) {
    return fallback;
  }
}

async function storageSet(key, value) {
  if (IS_LOCALHOST) {
    try {
      localStorage.setItem("vivo:" + key, JSON.stringify(value));
    } catch (e) {
      console.error("storageSet (localStorage) error", key, e);
    }
    return;
  }
  try {
    await withTimeout(window.storage.set(key, JSON.stringify(value), true), 5000, null);
  } catch (e) {
    console.error("storageSet error", key, e);
  }
}

async function storageGetPrivate(key, fallback) {
  return fallback;
}

const APP_VERSION = "2026-06-25.3";

function MigrationBanner({ migration, onMigrate }) {
  if (migration.done) {
    return (
      <div className="vivo-migration-banner vivo-migration-success">
        <CheckCircle2 size={16} />
        <span>Dados migrados com sucesso.</span>
      </div>
    );
  }
  return (
    <div className="vivo-migration-banner">
      <AlertTriangle size={16} />
      <span>Dados ainda não estão no espaço compartilhado. Clique para migrar.</span>
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
  const [generatedOrder, setGeneratedOrder] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [mainTab, setMainTab] = useState("relatorios");
  const [subTab, setSubTab] = useState("import");
  const [authed, setAuthed] = useState(true);
  const [migration, setMigration] = useState({ checked: false, available: false, running: false, done: false, error: null });

  function goTo(mainId, subId) {
    setMainTab(mainId);
    setSubTab(subId);
  }

  useEffect(() => {
    if (!authed) return;
    let cancelled = false;
    (async () => {
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

      let needsMigration = false;
      const migratedProducts = { ...p };
      for (const codigo of Object.keys(migratedProducts)) {
        const prod = migratedProducts[codigo];
        if (prod.monitoringStartV90 !== undefined && prod.monitoringStartV30 === undefined) {
          migratedProducts[codigo] = { ...prod, monitoringStartV30: prod.monitoringStartV90 };
          needsMigration = true;
        }
      }
      if (needsMigration && !IS_LOCALHOST) {
        await withTimeout(window.storage.set(K_PRODUCTS, JSON.stringify(migratedProducts), true), 8000, null);
      }
      const finalProducts = needsMigration ? migratedProducts : p;

      setProducts(finalProducts);
      setHistory(h);
      setSettings({ defaultCoverageDays: 30, supplierCoverageDays: {}, ...s });
      setOrders(o);
      setLoaded(true);

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
      setMigration((m) => ({ ...m, running: false, error: "Não foi possível migrar." }));
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

  if (!loaded) {
    return (
      <div className="vivo-root vivo-loading">
        <style>{VIVO_CSS}</style>
        <div className="vivo-loading-mark">VIVO</div>
        <div className="vivo-loading-text">Carregando…</div>
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
        monitoredCount={Object.values(products).filter((p) => p.monitoringActive).length}
      />
      <main className="vivo-main">
        {mainTab === "relatorios" && subTab === "import" && (
          <ImportTab
            products={products}
            persistProducts={persistProducts}
            history={history}
            persistHistory={persistHistory}
            goToProducts={() => goTo("compras", "suppliers")}
          />
        )}
        {mainTab === "relatorios" && subTab === "prices" && (
          <PriceTableTab products={products} persistProducts={persistProducts} />
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
        {mainTab === "compras" && subTab === "suppliers" && (
          <SuppliersTab
            products={products}
            settings={settings}
            persistSettings={persistSettings}
            persistProducts={persistProducts}
            onGenerateOrder={(order) => { setGeneratedOrder(order); goTo("compras", "generated-order"); }}
          />
        )}
        {mainTab === "compras" && subTab === "generated-order" && (
          <GeneratedOrderTab
            generatedOrder={generatedOrder}
            products={products}
            settings={settings}
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

const NAV_STRUCTURE = [
  {
    id: "relatorios",
    label: "Relatórios",
    icon: Upload,
    subItems: [
      { id: "import", label: "Subir relatório" },
      { id: "prices", label: "Tabela de preços" },
      { id: "history", label: "Histórico de importações" },
      { id: "cleanup", label: "Limpeza de itens" },
    ],
  },
  {
    id: "compras",
    label: "Compras",
    icon: ShoppingCart,
    subItems: [
      { id: "suppliers", label: "Fornecedores" },
      { id: "generated-order", label: "Pedido Gerado" },
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

function TopNav({ mainTab, subTab, goTo, monitoredCount }) {
  const countFor = (subId) => {
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
        <span className="vivo-topnav-hint">v{APP_VERSION} • custo líquido ativo</span>
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

function PriceTableTab({ products, persistProducts }) {
  const [status, setStatus] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);
    setStatus(null);
    setError(null);

    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

      const priceMap = {};
      for (let i = 2; i < rows.length; i++) {
        const row = rows[i];
        const sku = row[2] !== null && row[2] !== undefined ? String(row[2]).trim() : null;
        const costK = row[10];
        const costS = row[18];
        if (!sku) continue;
        priceMap[sku] = {};
        if (costK !== null && costK !== undefined && !isNaN(Number(costK)))
          priceMap[sku].precoCusto = Math.round(Number(costK) * 10000) / 10000;
        if (costS !== null && costS !== undefined && !isNaN(Number(costS)))
          priceMap[sku].custoLiquido = Math.round(Number(costS) * 10000) / 10000;
      }

      const totalTabela = Object.keys(priceMap).length;
      let updated = 0;
      let notFound = 0;

      const next = { ...products };
      for (const codigo of Object.keys(next)) {
        if (priceMap[codigo] !== undefined) {
          next[codigo] = { ...next[codigo], ...priceMap[codigo] };
          updated++;
        }
      }

      for (const sku of Object.keys(priceMap)) {
        if (!products[sku]) notFound++;
      }

      await persistProducts(next);
      setStatus({ updated, notFound, totalTabela, totalProdutos: Object.keys(products).length });
    } catch (err) {
      setError("Não foi possível ler o arquivo.");
      console.error(err);
    } finally {
      setLoading(false);
      e.target.value = "";
    }
  }

  const updatedCount = Object.values(products).filter((p) => p.precoCusto).length;
  const liquidoCount = Object.values(products).filter((p) => p.custoLiquido).length;
  const missingCount = Object.values(products).filter((p) => !p.precoCusto && !p.inactive).length;

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Tabela de preços de custo</h1>
        <p>
          Importa dois custos distintos por SKU:<br/>
          <strong>Custo de nota (K)</strong> — usado em pedidos de compra e geração de pedido.<br/>
          <strong>Custo líquido ATON (S)</strong> — usado em relatórios financeiros, capital imobilizado e estoque R$.
        </p>
      </header>

      <div className="vivo-price-status-cards">
        <div className="vivo-kpi-card">
          <span className="vivo-kpi-label">Com custo de nota (K)</span>
          <span className="vivo-kpi-value vivo-above-min">{updatedCount}</span>
        </div>
        <div className="vivo-kpi-card">
          <span className="vivo-kpi-label">Com custo líquido (S)</span>
          <span className="vivo-kpi-value vivo-above-min">{liquidoCount}</span>
        </div>
        <div className="vivo-kpi-card">
          <span className="vivo-kpi-label">Sem nenhum custo</span>
          <span className={"vivo-kpi-value" + (missingCount > 0 ? " vivo-kpi-warn" : "")}>{missingCount}</span>
        </div>
      </div>

      <div className="vivo-card vivo-price-upload-card">
        <p className="vivo-settings-hint" style={{ margin: "0 0 14px" }}>
          Selecione <strong>Pasta1.xlsx</strong> (ou versão atualizada). Os preços existentes serão sobrescritos.
        </p>
        <label className={"vivo-btn vivo-btn-primary" + (loading ? " vivo-btn-loading" : "")}>
          {loading ? "Importando…" : "Selecionar planilha de preços"}
          <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={handleFile} disabled={loading} />
        </label>
      </div>

      {status && (
        <div className="vivo-card vivo-price-result">
          <div className="vivo-price-result-item vivo-above-min">
            <CheckCircle2 size={16} />
            <strong>{status.updated}</strong> produtos atualizados (custo de nota + custo líquido)
          </div>
          <div className="vivo-price-result-item" style={{ color: "var(--ink-soft)" }}>
            <Package size={16} />
            {status.notFound} SKUs da tabela não encontrados nos produtos importados
          </div>
        </div>
      )}

      {error && (
        <div className="vivo-alert vivo-alert-warning" style={{ marginTop: 16 }}>
          <AlertTriangle size={16} /> {error}
        </div>
      )}
    </div>
  );
}

const PERIOD_LABELS = { 30: "30 dias", 90: "90 dias", 180: "180 dias" };

function detectPeriod(fileName, rows) {
  const name = fileName.toLowerCase();
  if (name.includes("180")) return 180;
  if (name.includes("90")) return 90;
  if (name.includes("30")) return 30;
  if (rows && rows.length > 0) {
    const mapped = mapRowToFields(rows[0]);
    if (mapped.vendas180 !== undefined) return 180;
    if (mapped.vendas90 !== undefined) return 90;
    if (mapped.vendas30 !== undefined) return 30;
  }
  return 30;
}

function ImportTab({ products, persistProducts, history, persistHistory, goToProducts }) {
  const [pendingFiles, setPendingFiles] = useState([]);
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
    const files = Array.from(fileList).slice(0, 6);
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
        merged[codigo].estoque = estoque;
        merged[codigo].fornecedor = fornecedor || merged[codigo].fornecedor;
        merged[codigo].item = item || merged[codigo].item;
        if (precoCusto !== undefined) merged[codigo].precoCusto = precoCusto;

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
        <p>Envie as planilhas exportadas do seu ERP (.xlsx ou .csv). Linhas sem fornecedor preenchido são descartadas automaticamente.</p>
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
        <p className="vivo-dropzone-hint">Cada planilha deve ter Fornecedor, Item, Código, Estoque e coluna de vendas (30, 90 ou 180 dias).</p>
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
                Confirme o período de vendas de cada arquivo antes de importar.
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
              <AlertTriangle size={16} /> Dois arquivos estão marcados com o mesmo período.
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
        </div>
      )}
    </div>
  );
}

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

  if (list.length === 0) {
    return (
      <div className="vivo-page">
        <EmptyState icon={Box} title="Nenhum produto importado ainda" text="Vá em Importar relatório e envie a primeira planilha." />
      </div>
    );
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Produtos</h1>
        <p>{list.length} produtos cadastrados.</p>
      </header>
      <div className="vivo-toolbar">
        <div className="vivo-search">
          <Search size={15} />
          <input placeholder="Buscar…" value={search} onChange={(e) => setSearch(e.target.value)} />
        </div>
        <select className="vivo-select" value={supplierFilter} onChange={(e) => setSupplierFilter(e.target.value)}>
          <option value="all">Todos os fornecedores</option>
          {suppliers.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      </div>
      <div className="vivo-table-wrap vivo-card">
        <table className="vivo-table">
          <thead>
            <tr>
              <th>Item</th>
              <th>Fornecedor</th>
              <th className="mono">Código</th>
              <th className="num">Estoque</th>
              <th className="num">Custo Nota</th>
              <th className="num">Custo Líq.</th>
              <th className="num">Vendas 30d</th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((p) => (
              <tr key={p.codigo}>
                <ItemCell name={p.item} />
                <td>{p.fornecedor}</td>
                <td className="mono">{p.codigo}</td>
                <td className="num mono">{fmtNumber(p.estoque)}</td>
                <td className="num mono">{p.precoCusto ? fmtCurrency(p.precoCusto) : "—"}</td>
                <td className="num mono">{p.custoLiquido ? fmtCurrency(p.custoLiquido) : "—"}</td>
                <td className="num mono">{p.vendas30 ?? "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <PaginationControls page={page} setPage={setPage} totalPages={totalPages} totalCount={totalCount} />
    </div>
  );
}

function PerformanceTab({ products }) {
  const [supplierFilter, setSupplierFilter] = useState("all");
  const [trendFilter, setTrendFilter] = useState("all");
  const [sortBy, setSortBy] = useState("deltaPct");

  const list = useMemo(() => Object.values(products).filter((p) => !p.inactive), [products]);
  const suppliers = useMemo(() => {
    const s = new Set(list.map((p) => p.fornecedor));
    return Array.from(s).sort();
  }, [list]);

  const computed = useMemo(() => list.map((p) => ({ ...p, trend: getTrendSummary(p.salesHistory || []) })), [list]);

  const filtered = useMemo(() => {
    let out = computed;
    if (supplierFilter !== "all") out = out.filter((p) => p.fornecedor === supplierFilter);
    if (trendFilter !== "all") out = out.filter((p) => p.trend.status === trendFilter);
    return [...out].sort((a, b) => (b.trend.deltaPct ?? -Infinity) - (a.trend.deltaPct ?? -Infinity));
  }, [computed, supplierFilter, trendFilter]);

  const { page, setPage, totalPages, pageItems, totalCount } = usePagination(filtered);

  const counts = useMemo(() => {
    const c = { subindo: 0, caindo: 0, estavel: 0, primeira_leitura: 0, sem_dados: 0 };
    computed.forEach((p) => { c[p.trend.status] = (c[p.trend.status] || 0) + 1; });
    return c;
  }, [computed]);

  if (list.length === 0) {
    return <div className="vivo-page"><EmptyState icon={TrendingUp} title="Nenhum dado ainda" text="Importe relatórios em pelo menos duas datas diferentes." /></div>;
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Tendência de vendas</h1>
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
      </div>
      <div className="vivo-toolbar">
        <select className="vivo-select" value={supplierFilter} onChange={(e) => setSupplierFilter(e.target.value)}>
          <option value="all">Todos os fornecedores</option>
          {suppliers.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      </div>
      <div className="vivo-table-wrap vivo-card">
        <table className="vivo-table">
          <thead>
            <tr>
              <th>Item</th>
              <th>Fornecedor</th>
              <th className="mono">Código</th>
              <th className="num">Média /dia (atual)</th>
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
                <td className={"num mono" + (p.trend.deltaPct > 0 ? " vivo-trend-up" : p.trend.deltaPct < 0 ? " vivo-trend-down" : "")}>
                  {p.trend.deltaPct !== null ? `${p.trend.deltaPct > 0 ? "+" : ""}${p.trend.deltaPct.toFixed(0)}%` : "—"}
                </td>
                <td><TrendBadge status={p.trend.status} /></td>
              </tr>
            ))}
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
  return <span className="vivo-badge vivo-badge-muted">Sem dados</span>;
}

function buildCapitalSeries(products, supplierFilter, excessCodes, days = 90) {
  const now = Date.now();
  const cutoff = now - days * MS_PER_DAY;
  const productList = Object.values(products).filter((p) =>
    !p.inactive && excessCodes.has(p.codigo) && (supplierFilter === "all" || p.fornecedor === supplierFilter)
  );
  const byDay = {};
  for (const p of productList) {
    const custo = custoLiq(p);
    for (const snap of (p.salesHistory || [])) {
      if (snap.ts < cutoff) continue;
      const dk = dayKey(snap.ts);
      if (!byDay[dk]) byDay[dk] = { ts: snap.ts, items: {} };
      if (!byDay[dk].items[p.codigo] || snap.ts > byDay[dk].items[p.codigo].ts) {
        byDay[dk].items[p.codigo] = { ts: snap.ts, estoque: snap.estoque ?? 0, custo };
      }
      if (snap.ts > byDay[dk].ts) byDay[dk].ts = snap.ts;
    }
  }
  return Object.entries(byDay)
    .map(([dk, { ts, items }]) => ({ dk, ts, capital: Object.values(items).reduce((acc, i) => acc + (i.estoque * i.custo), 0) }))
    .sort((a, b) => a.ts - b.ts);
}

function buildTotalStockSeries(products, supplierFilter, days = 90) {
  const now = Date.now();
  const cutoff = now - days * MS_PER_DAY;
  const productList = Object.values(products).filter((p) => !p.inactive && (supplierFilter === "all" || p.fornecedor === supplierFilter));
  const byDay = {};
  for (const p of productList) {
    const custo = custoLiq(p);
    for (const snap of (p.salesHistory || [])) {
      if (snap.ts < cutoff) continue;
      const dk = dayKey(snap.ts);
      if (!byDay[dk]) byDay[dk] = { ts: snap.ts, items: {} };
      if (!byDay[dk].items[p.codigo] || snap.ts > byDay[dk].items[p.codigo].ts) {
        byDay[dk].items[p.codigo] = { ts: snap.ts, estoque: snap.estoque ?? 0, custo };
      }
      if (snap.ts > byDay[dk].ts) byDay[dk].ts = snap.ts;
    }
  }
  return Object.entries(byDay)
    .map(([dk, { ts, items }]) => ({ dk, ts, capital: Object.values(items).reduce((acc, i) => acc + (i.estoque * i.custo), 0) }))
    .sort((a, b) => a.ts - b.ts);
}

function buildExcessStockSeries(products, supplierFilter, thresholdDays = 90, days = 90) {
  const now = Date.now();
  const cutoff = now - days * MS_PER_DAY;
  const productList = Object.values(products).filter((p) => !p.inactive && (supplierFilter === "all" || p.fornecedor === supplierFilter));
  const byDay = {};
  for (const p of productList) {
    const custo = custoLiq(p);
    for (const snap of (p.salesHistory || [])) {
      if (snap.ts < cutoff) continue;
      const dk = dayKey(snap.ts);
      if (!byDay[dk]) byDay[dk] = { ts: snap.ts, items: {} };
      if (!byDay[dk].items[p.codigo] || snap.ts > byDay[dk].items[p.codigo].ts) {
        byDay[dk].items[p.codigo] = { ts: snap.ts, estoque: snap.estoque ?? 0, vendas30: snap.vendas30 ?? 0, custo };
      }
      if (snap.ts > byDay[dk].ts) byDay[dk].ts = snap.ts;
    }
  }
  return Object.entries(byDay)
    .map(([dk, { ts, items }]) => {
      const excess = Object.values(items).reduce((acc, i) => {
        const avgDay = (i.vendas30 || 0) / 30;
        const idealMax = avgDay * thresholdDays;
        const excessUnits = Math.max(0, i.estoque - idealMax);
        return acc + excessUnits * i.custo;
      }, 0);
      return { dk, ts, capital: excess };
    })
    .sort((a, b) => a.ts - b.ts);
}

function CapitalChart({ products, supplierFilter, excessCodes }) {
  const [hover, setHover] = useState(null);
  const svgRef = React.useRef(null);

  const series = useMemo(() => buildCapitalSeries(products, supplierFilter, excessCodes, 90), [products, supplierFilter, excessCodes]);

  if (series.length < 2) {
    return <div className="vivo-capital-chart vivo-capital-chart-empty"><span>Dados insuficientes para o gráfico</span></div>;
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

  function handleMouseMove(e) {
    const svg = svgRef.current;
    if (!svg) return;
    const rect = svg.getBoundingClientRect();
    const vbX = (e.clientX - rect.left) * (W / rect.width);
    let closest = 0, minDist = Infinity;
    series.forEach((_, i) => { const dist = Math.abs(toX(i) - vbX); if (dist < minDist) { minDist = dist; closest = i; } });
    setHover({ i: closest, x: toX(closest), y: toY(series[closest].capital), capital: series[closest].capital, ts: series[closest].ts });
  }

  const tooltipLeft = hover ? hover.i < series.length * 0.65 : false;
  const hoverDelta = hover && hover.i > 0 ? hover.capital - series[hover.i - 1].capital : null;
  const displayPoint = hover || { capital: last.capital, ts: last.ts, x: toX(series.length - 1), y: toY(last.capital) };

  return (
    <div className="vivo-capital-chart">
      <div className="vivo-capital-chart-header">
        <div>
          <span className="vivo-capital-chart-label">Capital imobilizado (excesso, custo líquido)</span>
          {supplierFilter !== "all" && <span className="vivo-capital-chart-supplier"> · {supplierFilter}</span>}
        </div>
        <div className="vivo-capital-chart-summary">
          {hover ? (
            <>
              <span style={{ fontWeight: 600, color: "var(--ink)" }}>{fmtCurrency(hover.capital)}</span>
              <span className="vivo-capital-chart-delta" style={{ color: "var(--ink-soft)" }}>
                {fmtDateShort(hover.ts)}
                {hoverDelta !== null && <span style={{ color: hoverDelta > 0 ? "var(--rust)" : "var(--olive-dark)", marginLeft: 6 }}>{hoverDelta > 0 ? "+" : ""}{fmtCurrency(hoverDelta)}</span>}
              </span>
            </>
          ) : (
            <>
              <span style={{ color: trendColor, fontWeight: 600 }}>{trend === "up" ? "▲" : trend === "down" ? "▼" : "●"} {fmtCurrency(last.capital)}</span>
              <span className="vivo-capital-chart-delta" style={{ color: trendColor }}>{deltaPct > 0 ? "+" : ""}{deltaPct.toFixed(1)}% desde {fmtDateShort(first.ts)}</span>
            </>
          )}
        </div>
      </div>
      <div style={{ position: "relative" }}>
        <svg ref={svgRef} viewBox={`0 0 ${W} ${H}`} className="vivo-capital-chart-svg" preserveAspectRatio="none" onMouseMove={handleMouseMove} onMouseLeave={() => setHover(null)} style={{ cursor: "crosshair" }}>
          <defs>
            <linearGradient id="capGrad" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor={trendColor} stopOpacity="0.18" />
              <stop offset="100%" stopColor={trendColor} stopOpacity="0.02" />
            </linearGradient>
          </defs>
          <polygon points={areaPoints} fill="url(#capGrad)" />
          <polyline points={points} fill="none" stroke={trendColor} strokeWidth="1.8" strokeLinejoin="round" strokeLinecap="round" />
          {hover && <line x1={hover.x} y1={PAD.top} x2={hover.x} y2={PAD.top + innerH} stroke="var(--ink-soft)" strokeWidth="1" strokeDasharray="3,2" />}
          <circle cx={displayPoint.x} cy={displayPoint.y} r={hover ? 4 : 3} fill={hover ? "var(--ink)" : trendColor} stroke="var(--card)" strokeWidth="1.5" />
          <text x={toX(0)} y={H - 4} fontSize="9" fill="var(--ink-soft)" textAnchor="start">{fmtDateShort(first.ts)}</text>
          <text x={toX(series.length - 1)} y={H - 4} fontSize="9" fill="var(--ink-soft)" textAnchor="end">{fmtDateShort(last.ts)}</text>
        </svg>
        {hover && (
          <div className="vivo-chart-tooltip" style={{ left: tooltipLeft ? `calc(${(hover.x / W) * 100}% + 8px)` : "auto", right: tooltipLeft ? "auto" : `calc(${((W - hover.x) / W) * 100}% + 8px)`, top: "4px" }}>
            <div className="vivo-chart-tooltip-date">{new Date(hover.ts).toLocaleDateString("pt-BR", { day: "2-digit", month: "short" })}</div>
            <div className="vivo-chart-tooltip-value">{fmtCurrency(hover.capital)}</div>
            {hoverDelta !== null && <div className="vivo-chart-tooltip-delta" style={{ color: hoverDelta > 0 ? "var(--rust)" : hoverDelta < 0 ? "var(--olive-dark)" : "var(--ink-soft)" }}>{hoverDelta > 0 ? "+" : ""}{fmtCurrency(hoverDelta)} vs anterior</div>}
          </div>
        )}
      </div>
    </div>
  );
}

function StockQualityTab({ products, persistProducts }) {
  const [search, setSearch] = useState("");
  const [supplierFilter, setSupplierFilter] = useState("all");
  const [sortBy, setSortBy] = useState("stockDays_desc");
  const [scrollWidth, setScrollWidth] = useState(0);
  const [selected, setSelected] = useState({});
  const topScrollRef = React.useRef(null);
  const tableWrapRef = React.useRef(null);

  const measureScrollWidth = useCallback(() => { if (tableWrapRef.current) setScrollWidth(tableWrapRef.current.scrollWidth); }, []);
  useEffect(() => { measureScrollWidth(); window.addEventListener("resize", measureScrollWidth); return () => window.removeEventListener("resize", measureScrollWidth); }, [measureScrollWidth]);
  function syncFromTop(e) { if (tableWrapRef.current) tableWrapRef.current.scrollLeft = e.target.scrollLeft; }
  function syncFromTable(e) { if (topScrollRef.current) topScrollRef.current.scrollLeft = e.target.scrollLeft; }

  const list = useMemo(() => Object.values(products).filter((p) => !p.inactive), [products]);
  const computed = useMemo(() => list.map((p) => ({ ...p, quality: getStockQuality(p) })).filter((p) => p.quality.isExcess), [list]);
  const excessCodes = useMemo(() => new Set(computed.map((p) => p.codigo)), [computed]);
  const suppliers = useMemo(() => Array.from(new Set(computed.map((p) => p.fornecedor))).sort(), [computed]);
  const capitalBySupplier = useMemo(() => { const map = {}; for (const p of computed) map[p.fornecedor] = (map[p.fornecedor] || 0) + p.quality.capitalImobilizado; return map; }, [computed]);

  const filtered = useMemo(() => {
    let out = computed;
    if (supplierFilter !== "all") out = out.filter((p) => p.fornecedor === supplierFilter);
    if (search.trim()) { const q = search.toLowerCase(); out = out.filter((p) => p.item.toLowerCase().includes(q) || p.codigo.toLowerCase().includes(q)); }
    const sorted = [...out];
    if (sortBy === "stockDays_desc") sorted.sort((a, b) => b.quality.stockDays - a.quality.stockDays);
    else if (sortBy === "capital_desc") sorted.sort((a, b) => b.quality.capitalImobilizado - a.quality.capitalImobilizado);
    return sorted;
  }, [computed, supplierFilter, search, sortBy]);

  const { page, setPage, totalPages, pageItems, totalCount } = usePagination(filtered);
  useEffect(() => { const id = setTimeout(measureScrollWidth, 0); return () => clearTimeout(id); }, [pageItems, measureScrollWidth]);
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
      next[codigo] = { ...p, monitoringActive: false, monitoringStartedAt: null, monitoringStartV30: null };
    } else {
      next[codigo] = { ...p, monitoringActive: true, monitoringStartedAt: Date.now(), monitoringStartV30: p.vendas30 ?? null };
    }
    await persistProducts(next);
  }

  function toggleSelect(codigo) { setSelected((s) => ({ ...s, [codigo]: !s[codigo] })); }
  function selectAllVisible() { const next = { ...selected }; filtered.forEach((p) => { next[p.codigo] = true; }); setSelected(next); }
  function clearSelection() { setSelected({}); }
  const selectedCount = filtered.filter((p) => selected[p.codigo]).length;

  async function sendSelectedToMonitoring() {
    if (selectedCount === 0) return;
    const next = { ...products };
    const now = Date.now();
    for (const p of filtered) {
      if (!selected[p.codigo]) continue;
      const existing = next[p.codigo];
      if (existing.monitoringActive) continue;
      next[p.codigo] = { ...existing, monitoringActive: true, monitoringStartedAt: now, monitoringStartV30: existing.vendas30 ?? null };
    }
    await persistProducts(next);
    setSelected({});
  }

  if (list.length === 0) return <div className="vivo-page"><EmptyState icon={AlertTriangle} title="Nenhum produto importado" text="Importe relatórios para começar." /></div>;
  if (computed.length === 0) return <div className="vivo-page"><header className="vivo-page-head"><h1>Qualidade de Estoque</h1></header><EmptyState icon={CheckCircle2} title="Nenhum item em excesso" text="Bom sinal." /></div>;

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Qualidade de Estoque</h1>
        <p>Itens com mais de {STOCK_QUALITY_THRESHOLD_DAYS} dias de estoque. Capital imobilizado calculado pelo <strong>custo líquido</strong>.</p>
      </header>

      <div className="vivo-toolbar">
        <div className="vivo-search"><Search size={15} /><input placeholder="Buscar…" value={search} onChange={(e) => setSearch(e.target.value)} /></div>
        <select className="vivo-select" value={supplierFilter} onChange={(e) => setSupplierFilter(e.target.value)}>
          <option value="all">Todos os fornecedores</option>
          {suppliers.map((s) => <option key={s} value={s}>{s} — {fmtCurrency(capitalBySupplier[s])}</option>)}
        </select>
        <select className="vivo-select" value={sortBy} onChange={(e) => setSortBy(e.target.value)}>
          <option value="stockDays_desc">Tempo de estoque (maior)</option>
          <option value="capital_desc">Capital imobilizado (maior)</option>
        </select>
      </div>

      <div className="vivo-quality-summary">
        <span><strong>{filtered.length}</strong> itens em excesso</span>
        <span>•</span>
        <span>Capital imobilizado (custo líquido): <strong className="vivo-capital-danger">{fmtCurrency(totalCapital)}</strong></span>
      </div>

      <CapitalChart products={products} supplierFilter={supplierFilter} excessCodes={excessCodes} />

      <div className="vivo-action-row">
        <button className="vivo-btn vivo-btn-ghost" onClick={selectAllVisible}>Selecionar todos</button>
        <button className="vivo-btn vivo-btn-ghost" onClick={clearSelection}>Limpar</button>
        <div className="vivo-action-spacer" />
        <span className="vivo-action-summary">{selectedCount} selecionado(s)</span>
        <button className="vivo-btn vivo-btn-primary" onClick={sendSelectedToMonitoring} disabled={selectedCount === 0}>Enviar para Monitorados</button>
      </div>

      <div className="vivo-top-scrollbar" ref={topScrollRef} onScroll={syncFromTop}><div style={{ width: scrollWidth, height: 1 }} /></div>

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
              <th className="num">Custo Líq.</th>
              <th className="num">Capital Imob.</th>
              <th className="num">Tempo estoque</th>
              <th>Motivo</th>
              <th>Ação</th>
              <th>Margem MC</th>
              <th>Monitoramento</th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((p) => (
              <tr key={p.codigo} className={selected[p.codigo] ? "is-selected" : ""}>
                <td><input type="checkbox" checked={!!selected[p.codigo]} onChange={() => toggleSelect(p.codigo)} /></td>
                <ItemCell name={p.item} />
                <td>{p.fornecedor}</td>
                <td className="mono">{p.codigo}</td>
                <td className="num mono">{fmtNumber(p.estoque)}</td>
                <td className="num mono">{p.vendas30 ?? "—"}</td>
                <td className="num mono">{custoLiq(p) ? fmtCurrency(custoLiq(p)) : <span className="vivo-no-price">—</span>}</td>
                <td className="num mono vivo-capital-danger">{fmtCurrency(p.quality.capitalImobilizado)}</td>
                <td className="num mono vivo-stockdays">{p.quality.stockDays === Infinity ? "sem vendas" : `${Math.round(p.quality.stockDays)}d`}</td>
                <td>
                  <select className={"vivo-select-tag" + tagColorClass("reason", p.qualityReason)} value={p.qualityReason || ""} onChange={(e) => updateClassification(p.codigo, "qualityReason", e.target.value)}>
                    <option value="">—</option>
                    {QUALITY_REASON_TAGS.map((t) => <option key={t} value={t}>{t}</option>)}
                  </select>
                </td>
                <td>
                  <select className={"vivo-select-tag" + tagColorClass("situation", p.qualitySituation)} value={p.qualitySituation || ""} onChange={(e) => updateClassification(p.codigo, "qualitySituation", e.target.value)}>
                    <option value="">—</option>
                    {QUALITY_SITUATION_TAGS.map((t) => <option key={t} value={t}>{t}</option>)}
                  </select>
                </td>
                <td>
                  <select className={"vivo-select-tag" + tagColorClass("margin", p.qualityMargin)} value={p.qualityMargin ?? ""} onChange={(e) => updateClassification(p.codigo, "qualityMargin", e.target.value === "" ? "" : Number(e.target.value))}>
                    <option value="">—</option>
                    {QUALITY_MARGIN_OPTIONS.map((m) => <option key={m} value={m}>{m > 0 ? `+${m}` : m}%</option>)}
                  </select>
                </td>
                <td>
                  <button className={"vivo-toggle-btn" + (p.monitoringActive ? " is-on" : "")} onClick={() => toggleMonitoring(p.codigo)}>
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

function MonitoredTab({ products, persistProducts }) {
  const [supplierFilter, setSupplierFilter] = useState("all");
  const [selected, setSelected] = useState({});
  const [scrollWidth, setScrollWidth] = useState(0);
  const topScrollRef = React.useRef(null);
  const tableWrapRef = React.useRef(null);

  const measureScrollWidth = useCallback(() => { if (tableWrapRef.current) setScrollWidth(tableWrapRef.current.scrollWidth); }, []);
  useEffect(() => { measureScrollWidth(); window.addEventListener("resize", measureScrollWidth); return () => window.removeEventListener("resize", measureScrollWidth); }, [measureScrollWidth]);
  function syncFromTop(e) { if (tableWrapRef.current) tableWrapRef.current.scrollLeft = e.target.scrollLeft; }
  function syncFromTable(e) { if (topScrollRef.current) topScrollRef.current.scrollLeft = e.target.scrollLeft; }

  const list = useMemo(() => Object.values(products), [products]);
  const allMonitored = useMemo(() => list.filter((p) => p.monitoringActive).map((p) => ({ ...p, quality: getStockQuality(p), monitoring: getMonitoringStatus(p) })).sort((a, b) => (b.monitoringStartedAt || 0) - (a.monitoringStartedAt || 0)), [list]);
  const suppliers = useMemo(() => Array.from(new Set(allMonitored.map((p) => p.fornecedor))).sort(), [allMonitored]);
  const monitored = useMemo(() => supplierFilter === "all" ? allMonitored : allMonitored.filter((p) => p.fornecedor === supplierFilter), [allMonitored, supplierFilter]);
  const { page, setPage, totalPages, pageItems, totalCount } = usePagination(monitored);
  useEffect(() => { const id = setTimeout(measureScrollWidth, 0); return () => clearTimeout(id); }, [pageItems, measureScrollWidth]);

  function toggleSelect(codigo) { setSelected((s) => ({ ...s, [codigo]: !s[codigo] })); }
  function selectAllVisible() { const next = { ...selected }; monitored.forEach((p) => { next[p.codigo] = true; }); setSelected(next); }
  function clearSelection() { setSelected({}); }
  const selectedCount = monitored.filter((p) => selected[p.codigo]).length;

  async function deactivateMonitoring(codigo) {
    const next = { ...products };
    next[codigo] = { ...next[codigo], monitoringActive: false, monitoringStartedAt: null, monitoringStartV30: null };
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

  if (allMonitored.length === 0) return <div className="vivo-page"><header className="vivo-page-head"><h1>Monitorados</h1></header><EmptyState icon={AlertTriangle} title="Nenhum item em monitoramento" text='Ative o monitoramento na aba Qualidade de Estoque.' /></div>;

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Monitorados</h1>
        <p>{allMonitored.length} itens em acompanhamento.</p>
      </header>
      <div className="vivo-toolbar">
        <select className="vivo-select" value={supplierFilter} onChange={(e) => setSupplierFilter(e.target.value)}>
          <option value="all">Todas as marcas</option>
          {suppliers.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      </div>
      <div className="vivo-action-row">
        <button className="vivo-btn vivo-btn-ghost" onClick={selectAllVisible}>Selecionar todos</button>
        <button className="vivo-btn vivo-btn-ghost" onClick={clearSelection}>Limpar</button>
        <div className="vivo-action-spacer" />
        <span className="vivo-action-summary">{selectedCount} selecionado(s)</span>
        <button className="vivo-btn vivo-btn-danger" onClick={deactivateSelected} disabled={selectedCount === 0}>Desativar monitoramento</button>
      </div>
      <div className="vivo-top-scrollbar" ref={topScrollRef} onScroll={syncFromTop}><div style={{ width: scrollWidth, height: 1 }} /></div>
      <div className="vivo-table-wrap vivo-table-wrap-sticky vivo-card" ref={tableWrapRef} onScroll={syncFromTable}>
        <table className="vivo-table vivo-table-compact">
          <thead>
            <tr>
              <th></th>
              <th>Item</th>
              <th>Fornecedor</th>
              <th className="mono">Código</th>
              <th className="num">Estoque</th>
              <th>Desde</th>
              <th className="num">Dias</th>
              <th className="num">V30D Início</th>
              <th className="num">V30D Atual</th>
              <th>Situação</th>
              <th>Observações</th>
              <th>Desativar</th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((p) => (
              <tr key={p.codigo} className={selected[p.codigo] ? "is-selected" : ""}>
                <td><input type="checkbox" checked={!!selected[p.codigo]} onChange={() => toggleSelect(p.codigo)} /></td>
                <ItemCell name={p.item} />
                <td>{p.fornecedor}</td>
                <td className="mono">{p.codigo}</td>
                <td className="num mono">{fmtNumber(p.estoque)}</td>
                <td className="mono">{p.monitoringStartedAt ? fmtDateShort(p.monitoringStartedAt) : "—"}</td>
                <td className="num mono">{p.monitoring.days !== null ? p.monitoring.days : "—"}</td>
                <td className="num mono">{p.monitoring.startV30 ?? "—"}</td>
                <td className="num mono">{p.monitoring.currentV30 ?? "—"}</td>
                <td><MonitoringStatusBadge status={p.monitoring.status} deltaPct={p.monitoring.deltaPct} /></td>
                <td><NotesField value={p.monitoringNotes || ""} onSave={(text) => saveNotes(p.codigo, text)} /></td>
                <td><button className="vivo-toggle-btn vivo-toggle-btn-off" onClick={() => deactivateMonitoring(p.codigo)}>Desativar</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <PaginationControls page={page} setPage={setPage} totalPages={totalPages} totalCount={totalCount} />
    </div>
  );
}

function NotesField({ value, onSave }) {
  const [draft, setDraft] = useState(value);
  useEffect(() => { setDraft(value); }, [value]);
  function commit() { if (draft !== value) onSave(draft); }
  return (
    <input type="text" className="vivo-notes-input" placeholder="Observação…" value={draft}
      onChange={(e) => setDraft(e.target.value)} onBlur={commit}
      onKeyDown={(e) => { if (e.key === "Enter") { commit(); e.target.blur(); } }} />
  );
}

function MonitoringStatusBadge({ status, deltaPct }) {
  const pctLabel = deltaPct !== null && deltaPct !== undefined ? ` (${deltaPct > 0 ? "+" : ""}${deltaPct.toFixed(0)}%)` : "";
  if (status === "acelerando") return <span className="vivo-monitor-badge vivo-monitor-up"><TrendingUp size={13} /> Acelerando{pctLabel}</span>;
  if (status === "caindo") return <span className="vivo-monitor-badge vivo-monitor-down"><TrendingDown size={13} /> Caindo{pctLabel}</span>;
  if (status === "igual") return <span className="vivo-monitor-badge vivo-monitor-flat"><Minus size={13} /> Igual{pctLabel}</span>;
  return <span className="vivo-monitor-badge vivo-monitor-muted">Sem dados</span>;
}

function calcSupplierPurchase(products, fornecedor, days) {
  return Object.values(products)
    .filter((p) => !p.inactive && p.fornecedor === fornecedor)
    .reduce((acc, p) => {
      const avgDay = (p.vendas30 || 0) / 30;
      const idealStock = avgDay * days;
      const needed = Math.max(0, idealStock - (p.estoque || 0));
      return acc + needed * custoNota(p); // custo de nota para pedidos
    }, 0);
}

function GeneratedOrderTab({ generatedOrder, products, settings }) {
  const defaultDays = generatedOrder?.days || 30;
  const [skuDays, setSkuDays] = useState({});
  useEffect(() => { setSkuDays({}); }, [generatedOrder]);

  if (!generatedOrder) return <div className="vivo-page"><EmptyState icon={ShoppingCart} title="Nenhum pedido gerado" text="Vá em Fornecedores, defina a cobertura e clique em Gerar Pedido." /></div>;

  const { fornecedor, ts } = generatedOrder;

  const items = useMemo(() => {
    return Object.values(products)
      .filter((p) => !p.inactive && p.fornecedor === fornecedor)
      .map((p) => {
        const days = skuDays[p.codigo] !== undefined ? skuDays[p.codigo] : defaultDays;
        const avgDay = (p.vendas30 || 0) / 30;
        const needed = Math.max(0, avgDay * days - (p.estoque || 0));
        const valorCompra = needed * custoNota(p);
        return { ...p, days, avgDay, needed, valorCompra };
      })
      .sort((a, b) => b.valorCompra - a.valorCompra);
  }, [products, fornecedor, defaultDays, skuDays]);

  const totalPedido = items.reduce((acc, i) => acc + i.valorCompra, 0);
  const totalItensNecessidade = items.filter((i) => i.needed > 0).length;

  function setDaysForSku(codigo, value) { setSkuDays((prev) => ({ ...prev, [codigo]: Math.max(1, Number(value) || 1) })); }

  function exportToExcel() {
    const rows = items.map((p) => ({
      "Item": p.item, "Fornecedor": p.fornecedor, "SKU": p.codigo,
      "Estoque": p.estoque || 0, "Vendas 30D": p.vendas30 ?? 0,
      "Cobertura (dias)": p.days, "Qtd. Necessária": p.needed > 0 ? Math.ceil(p.needed) : 0,
      "Custo de Nota (R$)": custoNota(p), "Valor Compra (R$)": Number(p.valorCompra.toFixed(2)),
    }));
    rows.push({ "Item": "TOTAL", "Fornecedor": fornecedor, "SKU": "", "Estoque": "", "Vendas 30D": "", "Cobertura (dias)": "", "Qtd. Necessária": "", "Custo de Nota (R$)": "", "Valor Compra (R$)": Number(totalPedido.toFixed(2)) });
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 40 }, { wch: 20 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 16 }, { wch: 16 }, { wch: 18 }, { wch: 18 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pedido");
    XLSX.writeFile(wb, `pedido-${fornecedor.toLowerCase().replace(/\s+/g, "-")}-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Pedido Gerado — {fornecedor}</h1>
        <p>Gerado em {fmtDate(ts)}. Valores calculados pelo <strong>custo de nota</strong>. Cobertura padrão: <strong>{defaultDays} dias</strong> — ajuste por SKU abaixo.</p>
      </header>
      <div className="vivo-generated-order-summary">
        <div className="vivo-generated-order-total">
          <span>Valor total (custo de nota)</span>
          <strong>{fmtCurrency(totalPedido)}</strong>
        </div>
        <div className="vivo-generated-order-meta">
          <span>{items.length} itens · {totalItensNecessidade} precisam de reposição</span>
        </div>
        <button className="vivo-btn vivo-btn-secondary" onClick={exportToExcel} style={{ marginLeft: "auto" }}>Exportar Excel</button>
      </div>
      <div className="vivo-table-wrap vivo-card">
        <table className="vivo-table vivo-table-compact">
          <thead>
            <tr>
              <th>Item</th>
              <th>Fornecedor</th>
              <th className="mono">SKU</th>
              <th className="num">Estoque</th>
              <th className="num">Vendas 30D</th>
              <th className="num">Cobertura (dias)</th>
              <th className="num">Qtd. Necessária</th>
              <th className="num">Custo Nota</th>
              <th className="num">Valor Compra</th>
            </tr>
          </thead>
          <tbody>
            {items.map((p) => (
              <tr key={p.codigo} className={p.needed === 0 ? "vivo-row-zero" : ""}>
                <ItemCell name={p.item} />
                <td>{p.fornecedor}</td>
                <td className="mono">{p.codigo}</td>
                <td className="num mono">{fmtNumber(p.estoque || 0)}</td>
                <td className="num mono">{p.vendas30 ?? "—"}</td>
                <td className="num">
                  <input type="number" min="1" className={"vivo-input-mini" + (skuDays[p.codigo] !== undefined ? " vivo-sku-days-custom" : "")} value={p.days} onChange={(e) => setDaysForSku(p.codigo, e.target.value)} />
                </td>
                <td className="num mono">{p.needed > 0 ? fmtNumber(Math.ceil(p.needed)) : "—"}</td>
                <td className="num mono">{custoNota(p) ? fmtCurrency(custoNota(p)) : "—"}</td>
                <td className={"num mono" + (p.valorCompra > 0 ? " vivo-value-positive" : "")}>{fmtCurrency(p.valorCompra)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function SkuDrawer({ product, products, persistProducts, onClose }) {
  const [obsText, setObsText] = useState("");
  const [saving, setSaving] = useState(false);

  const currentProduct = products[product.codigo] || product;
  const observations = currentProduct.observations || [];

  const salesSeries = useMemo(() => (currentProduct.salesHistory || []).filter((s) => s.vendas30 !== undefined).sort((a, b) => a.ts - b.ts).map((s) => ({ ts: s.ts, total: s.vendas30 ?? 0 })), [currentProduct.salesHistory]);
  const stockSeries = useMemo(() => (currentProduct.salesHistory || []).filter((s) => s.estoque !== undefined).sort((a, b) => a.ts - b.ts).map((s) => ({ ts: s.ts, total: s.estoque ?? 0 })), [currentProduct.salesHistory]);

  async function saveObs() {
    const text = obsText.trim();
    if (!text) return;
    setSaving(true);
    const newObs = { ts: Date.now(), text };
    const next = { ...products, [product.codigo]: { ...products[product.codigo], observations: [newObs, ...observations] } };
    await persistProducts(next);
    setObsText("");
    setSaving(false);
  }

  return (
    <>
      <div className="vivo-drawer-overlay" onClick={onClose} />
      <div className="vivo-drawer">
        <div className="vivo-drawer-header">
          <div>
            <div className="vivo-drawer-title">{product.item}</div>
            <div className="vivo-drawer-sub">{product.codigo} · {product.fornecedor}</div>
          </div>
          <button className="vivo-drawer-close" onClick={onClose}>✕</button>
        </div>
        <div className="vivo-drawer-body">
          <div className="vivo-drawer-chart-block">
            {salesSeries.length >= 2
              ? <MiniLineChart series={salesSeries} valueKey="total" color="var(--olive-dark)" unit="un" title={`Vendas 30D — ${fmtDateShort(salesSeries[0].ts)} a ${fmtDateShort(salesSeries[salesSeries.length-1].ts)}`} />
              : <div className="vivo-capital-chart-empty"><span>Dados insuficientes para gráfico de vendas</span></div>
            }
          </div>
          <div className="vivo-drawer-chart-block">
            {stockSeries.length >= 2
              ? <MiniLineChart series={stockSeries} valueKey="total" color="var(--amber)" unit="un" title={`Estoque — ${fmtDateShort(stockSeries[0].ts)} a ${fmtDateShort(stockSeries[stockSeries.length-1].ts)}`} />
              : <div className="vivo-capital-chart-empty"><span>Dados insuficientes para gráfico de estoque</span></div>
            }
          </div>
          <div className="vivo-drawer-obs-section">
            <div className="vivo-capital-chart-label" style={{ marginBottom: 8 }}>Observações</div>
            <div className="vivo-drawer-obs-input-row">
              <textarea className="vivo-drawer-obs-input" placeholder="Digite uma observação…" value={obsText} onChange={(e) => setObsText(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); saveObs(); } }} rows={2} />
              <button className="vivo-btn vivo-btn-primary" onClick={saveObs} disabled={saving || !obsText.trim()}>{saving ? "…" : "Salvar"}</button>
            </div>
            <div className="vivo-drawer-obs-hint">Enter para salvar · Shift+Enter para nova linha</div>
            {observations.length > 0 ? (
              <div className="vivo-drawer-obs-list">
                {observations.map((obs, i) => (
                  <div key={i} className="vivo-drawer-obs-item">
                    <span className="vivo-drawer-obs-date">{new Date(obs.ts).toLocaleDateString("pt-BR", { day: "2-digit", month: "short", year: "numeric" })} {new Date(obs.ts).toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })}</span>
                    <span className="vivo-drawer-obs-text">{obs.text}</span>
                  </div>
                ))}
              </div>
            ) : <p className="vivo-drawer-obs-empty">Nenhuma observação ainda.</p>}
          </div>
        </div>
      </div>
    </>
  );
}

function SupplierDetail({ fornecedor, products, settings, persistProducts, onBack }) {
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState("desc");
  const [statusFilter, setStatusFilter] = useState("all");
  const [itemColWidth, setItemColWidth] = useState(220);
  const [drawerProduct, setDrawerProduct] = useState(null);

  const TAG_STYLES = {
    escalando: { background: "#e1ecd9", color: "#3c6b2a" },
    liquidacao: { background: "var(--rust-soft)", color: "var(--rust)" },
    inativo: { background: "#e8e8e8", color: "#333" },
  };
  const TAG_LABELS = { escalando: "Escalando", liquidacao: "Liquidação", inativo: "Inativo" };

  function toggleSort(col) {
    if (sortCol === col) setSortDir((d) => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("desc"); }
  }

  function SortBtn({ col }) {
    const active = sortCol === col;
    return <button className="vivo-sort-btn" onClick={() => toggleSort(col)}>{active ? (sortDir === "desc" ? "▼" : "▲") : "⇅"}</button>;
  }

  const tag = settings.supplierTags?.[fornecedor] || null;
  const minOrder = settings.supplierMinOrders?.[fornecedor] || null;
  const excessCodes = useMemo(() => new Set(Object.values(products).filter((p) => !p.inactive && getStockQuality(p).isExcess).map((p) => p.codigo)), [products]);

  const skus = useMemo(() =>
    Object.values(products)
      .filter((p) => !p.inactive && p.fornecedor === fornecedor)
      .map((p) => {
        const q = getStockQuality(p);
        const avgDay = (p.vendas30 || 0) / 30;
        const coverageDays = avgDay > 0 ? Math.floor((p.estoque || 0) / avgDay) : null;
        const status = q.isExcess ? "excesso" : avgDay > 0 && coverageDays !== null && coverageDays < 15 ? "critico" : "ok";
        return { ...p, quality: q, avgDay, coverageDays, status, estoqueR: (p.estoque || 0) * custoLiq(p), capitalExc: q.isExcess ? q.capitalImobilizado : 0 };
      })
      .sort((a, b) => (b.vendas30 || 0) - (a.vendas30 || 0)),
    [products, fornecedor]);

  const skusSorted = useMemo(() => {
    let list = statusFilter === "all" ? skus : skus.filter((p) => p.status === statusFilter);
    if (sortCol) {
      list = [...list].sort((a, b) => {
        const va = a[sortCol] ?? -Infinity;
        const vb = b[sortCol] ?? -Infinity;
        return sortDir === "desc" ? vb - va : va - vb;
      });
    }
    return list;
  }, [skus, statusFilter, sortCol, sortDir]);

  const estoqueTotal = skus.reduce((acc, p) => acc + (p.estoque || 0) * custoLiq(p), 0);
  const capitalExcesso = skus.filter((p) => p.quality.isExcess).reduce((acc, p) => acc + p.quality.capitalImobilizado, 0);
  const criticos = skus.filter((p) => p.status === "critico").length;
  const excesso = skus.filter((p) => p.status === "excesso").length;

  const salesSeries = useMemo(() => {
    const byDay = {};
    const cutoff = Date.now() - 90 * MS_PER_DAY;
    for (const p of skus) {
      for (const snap of (p.salesHistory || [])) {
        if (snap.ts < cutoff || snap.vendas30 === undefined) continue;
        const dk = dayKey(snap.ts);
        if (!byDay[dk]) byDay[dk] = { ts: snap.ts, total: 0 };
        byDay[dk].total += snap.vendas30 || 0;
        if (snap.ts > byDay[dk].ts) byDay[dk].ts = snap.ts;
      }
    }
    return Object.values(byDay).sort((a, b) => a.ts - b.ts);
  }, [skus]);

  const totalStockSeries = useMemo(() => buildTotalStockSeries(products, fornecedor, 90), [products, fornecedor]);
  const excessStockSeries = useMemo(() => buildExcessStockSeries(products, fornecedor, 90, 90), [products, fornecedor]);

  return (
    <div className="vivo-page">
      <div className="vivo-detail-back" onClick={onBack}>← Voltar para Fornecedores</div>
      <header className="vivo-page-head" style={{ marginTop: 8 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <h1 style={{ margin: 0 }}>{fornecedor}</h1>
          {tag && <span className="vivo-tag-chip" style={{ ...TAG_STYLES[tag], padding: "3px 10px", borderRadius: 100, fontSize: 12, fontWeight: 600 }}>{TAG_LABELS[tag]}</span>}
        </div>
      </header>

      <div className="vivo-supplier-kpis">
        <div className="vivo-kpi-card"><span className="vivo-kpi-label">Estoque R$ (custo líq.)</span><span className="vivo-kpi-value">{fmtCurrency(estoqueTotal)}</span></div>
        <div className="vivo-kpi-card"><span className="vivo-kpi-label">Capital em Excesso</span><span className="vivo-kpi-value vivo-kpi-danger">{fmtCurrency(capitalExcesso)}</span></div>
        <div className="vivo-kpi-card"><span className="vivo-kpi-label">SKUs Críticos (&lt;15d)</span><span className={"vivo-kpi-value" + (criticos > 0 ? " vivo-kpi-danger" : "")}>{criticos}</span></div>
        <div className="vivo-kpi-card"><span className="vivo-kpi-label">SKUs em Excesso</span><span className={"vivo-kpi-value" + (excesso > 0 ? " vivo-kpi-warn" : "")}>{excesso}</span></div>
        <div className="vivo-kpi-card"><span className="vivo-kpi-label">Total de SKUs</span><span className="vivo-kpi-value">{skus.length}</span></div>
        {minOrder && <div className="vivo-kpi-card"><span className="vivo-kpi-label">Pedido Mínimo</span><span className="vivo-kpi-value">{fmtCurrency(minOrder)}</span></div>}
      </div>

      <div className="vivo-supplier-charts">
        <div className="vivo-supplier-chart-block vivo-card">
          <MiniCurrencyChart series={totalStockSeries} color="var(--olive-dark)" title={totalStockSeries.length >= 2 ? `Estoque total (custo líq.) — ${fmtDateShort(totalStockSeries[0].ts)} a ${fmtDateShort(totalStockSeries[totalStockSeries.length-1].ts)}` : "Estoque total em R$"} />
        </div>
        <div className="vivo-supplier-chart-block vivo-card">
          <MiniCurrencyChart series={excessStockSeries} color="var(--rust)" title={excessStockSeries.length >= 2 ? `Capital excedente (>90d, custo líq.) — ${fmtDateShort(excessStockSeries[0].ts)} a ${fmtDateShort(excessStockSeries[excessStockSeries.length-1].ts)}` : "Capital excedente"} />
        </div>
        <div className="vivo-supplier-chart-block vivo-card">
          <MiniLineChart series={salesSeries} valueKey="total" color="#4a7c6f" unit="un" title={salesSeries.length >= 2 ? `Vendas 30D — ${fmtDateShort(salesSeries[0].ts)} a ${fmtDateShort(salesSeries[salesSeries.length-1].ts)}` : "Vendas 30D totais"} />
        </div>
      </div>

      <div style={{ marginTop: 18 }}>
        <div className="vivo-toolbar" style={{ marginBottom: 8 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12 }}>
            <span style={{ color: "var(--ink-soft)" }}>Largura do item:</span>
            <input type="range" min={120} max={500} step={10} value={itemColWidth} onChange={(e) => setItemColWidth(Number(e.target.value))} style={{ width: 100 }} />
            <span style={{ color: "var(--ink-soft)", fontFamily: "monospace" }}>{itemColWidth}px</span>
          </div>
          <div className="vivo-supplier-tag-chips">
            <span className="vivo-supplier-tag-label">Situação:</span>
            {["all","ok","critico","excesso"].map((s) => (
              <button key={s} className={"vivo-supplier-tag-chip" + (statusFilter === s ? (s === "ok" ? " vivo-stag-escalando is-active" : s !== "all" ? " vivo-stag-liquidacao is-active" : " is-active-neutral") : "")} onClick={() => setStatusFilter(s)}>
                {s === "all" ? "Todas" : s === "ok" ? "OK" : s === "critico" ? "Crítico" : "Excesso"}
              </button>
            ))}
          </div>
        </div>

        <div className="vivo-table-wrap vivo-card">
          <table className="vivo-table vivo-table-compact">
            <thead>
              <tr>
                <th style={{ width: itemColWidth, minWidth: itemColWidth, maxWidth: itemColWidth }}>Item</th>
                <th className="mono">SKU</th>
                <th className="num">Estoque <SortBtn col="estoque" /></th>
                <th className="num">Vendas 30D <SortBtn col="vendas30" /></th>
                <th className="num">Cobertura <SortBtn col="coverageDays" /></th>
                <th className="num" title="Custo líquido (ATON) — usado em relatórios">Custo Líq. <SortBtn col="custoLiquido" /></th>
                <th className="num">Estoque R$ <SortBtn col="estoqueR" /></th>
                <th className="num">Capital Exc. <SortBtn col="capitalExc" /></th>
                <th>Situação</th>
              </tr>
            </thead>
            <tbody>
              {skusSorted.map((p) => (
                <tr key={p.codigo} className={drawerProduct?.codigo === p.codigo ? "is-selected" : ""} style={{ cursor: "pointer" }} onClick={() => setDrawerProduct(drawerProduct?.codigo === p.codigo ? null : p)}>
                  <td style={{ width: itemColWidth, minWidth: itemColWidth, maxWidth: itemColWidth, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: 500 }} title={p.item}>{p.item}</td>
                  <td className="mono">{p.codigo}</td>
                  <td className="num mono">{fmtNumber(p.estoque || 0)}</td>
                  <td className="num mono">{p.vendas30 ?? "—"}</td>
                  <td className={"num mono" + (p.status === "critico" ? " vivo-below-min" : "")}>{p.coverageDays !== null ? `${p.coverageDays}d` : "—"}</td>
                  <td className="num mono" title={p.precoCusto ? `Custo de nota: ${fmtCurrency(p.precoCusto)}` : "Custo de nota não cadastrado"}>{custoLiq(p) ? fmtCurrency(custoLiq(p)) : "—"}</td>
                  <td className="num mono">{fmtCurrency((p.estoque || 0) * custoLiq(p))}</td>
                  <td className={"num mono" + (p.quality.isExcess ? " vivo-below-min" : "")}>{p.quality.isExcess ? fmtCurrency(p.quality.capitalImobilizado) : "—"}</td>
                  <td>
                    {p.status === "critico" && <span className="vivo-badge vivo-badge-danger">Crítico</span>}
                    {p.status === "excesso" && <span className="vivo-badge vivo-badge-warn">Excesso</span>}
                    {p.status === "ok" && <span className="vivo-badge vivo-badge-ok">OK</span>}
                  </td>
                </tr>
              ))}
              {skusSorted.length === 0 && <tr><td colSpan={9} className="vivo-table-empty">Nenhum item com essa situação.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {drawerProduct && <SkuDrawer product={drawerProduct} products={products} persistProducts={persistProducts} onClose={() => setDrawerProduct(null)} />}
    </div>
  );
}

function MiniCurrencyChart({ series, color, title }) {
  const [hover, setHover] = useState(null);
  const svgRef = React.useRef(null);

  if (!series || series.length < 2) {
    return <div><div className="vivo-minichart-header"><span className="vivo-capital-chart-label">{title}</span></div><div className="vivo-capital-chart vivo-capital-chart-empty"><span>Dados insuficientes</span></div></div>;
  }

  const W = 800, H = 140, PAD = { top: 16, right: 20, bottom: 32, left: 80 };
  const innerW = W - PAD.left - PAD.right, innerH = H - PAD.top - PAD.bottom;
  const values = series.map((s) => s.capital);
  const minV = Math.min(...values), maxV = Math.max(...values), range = maxV - minV || 1;
  const toX = (i) => PAD.left + (i / (series.length - 1)) * innerW;
  const toY = (v) => PAD.top + innerH - ((v - minV) / range) * innerH;
  const points = series.map((s, i) => `${toX(i)},${toY(s.capital)}`).join(" ");
  const areaPoints = `${toX(0)},${PAD.top + innerH} ${points} ${toX(series.length - 1)},${PAD.top + innerH}`;
  const last = series[series.length - 1], first = series[0];
  const delta = last.capital - first.capital;
  const deltaPct = first.capital > 0 ? ((delta / first.capital) * 100).toFixed(1) : null;
  const gridLines = [0, 0.5, 1].map((pct) => ({ y: PAD.top + innerH - pct * innerH, label: fmtCurrency(minV + pct * range) }));
  const xLabels = useMemo(() => {
    const total = series.length;
    if (total <= 7) return series.map((s, i) => ({ i, ts: s.ts }));
    const step = Math.floor(total / 6);
    const result = [];
    for (let i = 0; i < total; i += step) result.push({ i, ts: series[i].ts });
    if (result[result.length - 1].i !== total - 1) result.push({ i: total - 1, ts: series[total - 1].ts });
    return result;
  }, [series]);

  function handleMouseMove(e) {
    const svg = svgRef.current; if (!svg) return;
    const rect = svg.getBoundingClientRect();
    const vbX = (e.clientX - rect.left) * (W / rect.width);
    let closest = 0, minDist = Infinity;
    series.forEach((_, i) => { const dist = Math.abs(toX(i) - vbX); if (dist < minDist) { minDist = dist; closest = i; } });
    const prev = closest > 0 ? series[closest - 1] : null;
    setHover({ i: closest, x: toX(closest), y: toY(series[closest].capital), value: series[closest].capital, ts: series[closest].ts, diff: prev ? series[closest].capital - prev.capital : null });
  }

  const tooltipLeft = hover ? hover.i < series.length * 0.65 : false;
  const gradId = `cgrad-${title.slice(0, 8).replace(/\s/g, "")}`;

  return (
    <div>
      <div className="vivo-minichart-header">
        <span className="vivo-capital-chart-label">{title}</span>
        <span className="vivo-minichart-current" style={{ color }}>
          {hover ? <><strong>{fmtCurrency(hover.value)}</strong>{hover.diff !== null && <span style={{ fontSize: 11, marginLeft: 8, color: hover.diff > 0 ? color : "var(--olive-dark)" }}>{hover.diff > 0 ? "+" : ""}{fmtCurrency(hover.diff)} vs anterior</span>}</> : <><strong>{fmtCurrency(last.capital)}</strong>{deltaPct && <span style={{ fontSize: 11, marginLeft: 8, color: delta >= 0 ? color : "var(--olive-dark)" }}>{delta >= 0 ? "+" : ""}{deltaPct}% no período</span>}</>}
        </span>
      </div>
      <div style={{ position: "relative" }}>
        <svg ref={svgRef} viewBox={`0 0 ${W} ${H}`} className="vivo-supplier-chart-svg" preserveAspectRatio="none" onMouseMove={handleMouseMove} onMouseLeave={() => setHover(null)} style={{ cursor: "crosshair" }}>
          <defs><linearGradient id={gradId} x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={color} stopOpacity="0.2" /><stop offset="100%" stopColor={color} stopOpacity="0.02" /></linearGradient></defs>
          {gridLines.map((g, gi) => (<g key={gi}><line x1={PAD.left} y1={g.y} x2={PAD.left + innerW} y2={g.y} stroke="var(--line)" strokeWidth="1" strokeDasharray="4,3" /><text x={PAD.left - 6} y={g.y + 4} fontSize="8.5" fill="var(--ink-soft)" textAnchor="end">{g.label}</text></g>))}
          <polygon points={areaPoints} fill={`url(#${gradId})`} />
          <polyline points={points} fill="none" stroke={color} strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
          {hover && <line x1={hover.x} y1={PAD.top} x2={hover.x} y2={PAD.top + innerH} stroke="var(--ink-soft)" strokeWidth="1" strokeDasharray="3,2" />}
          <circle cx={hover ? hover.x : toX(series.length - 1)} cy={hover ? hover.y : toY(last.capital)} r={hover ? 5 : 4} fill={hover ? "var(--ink)" : color} stroke="var(--card)" strokeWidth="2" />
          {xLabels.map((lb) => (<text key={lb.i} x={toX(lb.i)} y={H - 6} fontSize="9" fill="var(--ink-soft)" textAnchor="middle">{new Date(lb.ts).toLocaleDateString("pt-BR", { day: "2-digit", month: "short" })}</text>))}
          <line x1={PAD.left} y1={PAD.top + innerH} x2={PAD.left + innerW} y2={PAD.top + innerH} stroke="var(--line)" strokeWidth="1" />
        </svg>
        {hover && (
          <div className="vivo-chart-tooltip" style={{ left: tooltipLeft ? `calc(${(hover.x / W) * 100}% + 8px)` : "auto", right: tooltipLeft ? "auto" : `calc(${((W - hover.x) / W) * 100}% + 8px)`, top: "4px" }}>
            <div className="vivo-chart-tooltip-date">{new Date(hover.ts).toLocaleDateString("pt-BR", { day: "2-digit", month: "long", year: "numeric" })}</div>
            <div className="vivo-chart-tooltip-value">{fmtCurrency(hover.value)}</div>
            {hover.diff !== null && <div className="vivo-chart-tooltip-delta" style={{ color: hover.diff >= 0 ? color : "var(--olive-dark)" }}>{hover.diff >= 0 ? "+" : ""}{fmtCurrency(hover.diff)} vs importação anterior</div>}
          </div>
        )}
      </div>
    </div>
  );
}

function MiniLineChart({ series, valueKey, color, unit, title }) {
  const [hover, setHover] = useState(null);
  const svgRef = React.useRef(null);

  if (!series || series.length < 2) return <div className="vivo-capital-chart vivo-capital-chart-empty"><span>Dados insuficientes</span></div>;

  const W = 800, H = 140, PAD = { top: 16, right: 20, bottom: 32, left: 56 };
  const innerW = W - PAD.left - PAD.right, innerH = H - PAD.top - PAD.bottom;
  const values = series.map((s) => s[valueKey]);
  const minV = Math.min(...values), maxV = Math.max(...values), range = maxV - minV || 1;
  const toX = (i) => PAD.left + (i / (series.length - 1)) * innerW;
  const toY = (v) => PAD.top + innerH - ((v - minV) / range) * innerH;
  const points = series.map((s, i) => `${toX(i)},${toY(s[valueKey])}`).join(" ");
  const areaPoints = `${toX(0)},${PAD.top + innerH} ${points} ${toX(series.length - 1)},${PAD.top + innerH}`;
  const last = series[series.length - 1], first = series[0];
  const delta = last[valueKey] - first[valueKey];
  const deltaPct = first[valueKey] > 0 ? ((delta / first[valueKey]) * 100).toFixed(1) : null;
  const gridLines = [0, 0.5, 1].map((pct) => ({ y: PAD.top + innerH - pct * innerH, label: Math.round(minV + pct * range) }));
  const xLabels = useMemo(() => {
    const total = series.length;
    if (total <= 7) return series.map((s, i) => ({ i, ts: s.ts }));
    const step = Math.floor(total / 6);
    const result = [];
    for (let i = 0; i < total; i += step) result.push({ i, ts: series[i].ts });
    if (result[result.length - 1].i !== total - 1) result.push({ i: total - 1, ts: series[total - 1].ts });
    return result;
  }, [series]);

  function handleMouseMove(e) {
    const svg = svgRef.current; if (!svg) return;
    const rect = svg.getBoundingClientRect();
    const vbX = (e.clientX - rect.left) * (W / rect.width);
    let closest = 0, minDist = Infinity;
    series.forEach((_, i) => { const dist = Math.abs(toX(i) - vbX); if (dist < minDist) { minDist = dist; closest = i; } });
    const prev = closest > 0 ? series[closest - 1] : null;
    setHover({ i: closest, x: toX(closest), y: toY(series[closest][valueKey]), value: series[closest][valueKey], ts: series[closest].ts, diff: prev ? series[closest][valueKey] - prev[valueKey] : null });
  }

  const tooltipLeft = hover ? hover.i < series.length * 0.65 : false;

  return (
    <div>
      <div className="vivo-minichart-header">
        <span className="vivo-capital-chart-label">{title}</span>
        <span className="vivo-minichart-current" style={{ color }}>
          {hover ? <><strong>{hover.value.toLocaleString("pt-BR")} {unit}</strong>{hover.diff !== null && <span style={{ fontSize: 11, marginLeft: 8, color: hover.diff > 0 ? color : "var(--rust)" }}>{hover.diff > 0 ? "+" : ""}{hover.diff.toLocaleString("pt-BR")} vs anterior</span>}</> : <><strong>{last[valueKey].toLocaleString("pt-BR")} {unit}</strong>{deltaPct && <span style={{ fontSize: 11, marginLeft: 8, color: delta >= 0 ? color : "var(--rust)" }}>{delta >= 0 ? "+" : ""}{deltaPct}% no período</span>}</>}
        </span>
      </div>
      <div style={{ position: "relative" }}>
        <svg ref={svgRef} viewBox={`0 0 ${W} ${H}`} className="vivo-supplier-chart-svg" preserveAspectRatio="none" onMouseMove={handleMouseMove} onMouseLeave={() => setHover(null)} style={{ cursor: "crosshair" }}>
          <defs><linearGradient id={`grad-${valueKey}`} x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={color} stopOpacity="0.2" /><stop offset="100%" stopColor={color} stopOpacity="0.02" /></linearGradient></defs>
          {gridLines.map((g) => (<g key={g.y}><line x1={PAD.left} y1={g.y} x2={PAD.left + innerW} y2={g.y} stroke="var(--line)" strokeWidth="1" strokeDasharray="4,3" /><text x={PAD.left - 6} y={g.y + 4} fontSize="9" fill="var(--ink-soft)" textAnchor="end">{g.label.toLocaleString("pt-BR")}</text></g>))}
          <polygon points={areaPoints} fill={`url(#grad-${valueKey})`} />
          <polyline points={points} fill="none" stroke={color} strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
          {hover && <line x1={hover.x} y1={PAD.top} x2={hover.x} y2={PAD.top + innerH} stroke="var(--ink-soft)" strokeWidth="1" strokeDasharray="3,2" />}
          <circle cx={hover ? hover.x : toX(series.length - 1)} cy={hover ? hover.y : toY(last[valueKey])} r={hover ? 5 : 4} fill={hover ? "var(--ink)" : color} stroke="var(--card)" strokeWidth="2" />
          {xLabels.map((lb) => (<text key={lb.i} x={toX(lb.i)} y={H - 6} fontSize="9" fill="var(--ink-soft)" textAnchor="middle">{new Date(lb.ts).toLocaleDateString("pt-BR", { day: "2-digit", month: "short" })}</text>))}
          <line x1={PAD.left} y1={PAD.top + innerH} x2={PAD.left + innerW} y2={PAD.top + innerH} stroke="var(--line)" strokeWidth="1" />
        </svg>
        {hover && (
          <div className="vivo-chart-tooltip" style={{ left: tooltipLeft ? `calc(${(hover.x / W) * 100}% + 8px)` : "auto", right: tooltipLeft ? "auto" : `calc(${((W - hover.x) / W) * 100}% + 8px)`, top: "4px" }}>
            <div className="vivo-chart-tooltip-date">{new Date(hover.ts).toLocaleDateString("pt-BR", { day: "2-digit", month: "long", year: "numeric" })}</div>
            <div className="vivo-chart-tooltip-value">{hover.value.toLocaleString("pt-BR")} {unit}</div>
            {hover.diff !== null && <div className="vivo-chart-tooltip-delta" style={{ color: hover.diff >= 0 ? color : "var(--rust)" }}>{hover.diff >= 0 ? "+" : ""}{hover.diff.toLocaleString("pt-BR")} vs importação anterior</div>}
          </div>
        )}
      </div>
    </div>
  );
}

function SuppliersTab({ products, settings, persistSettings, persistProducts, onGenerateOrder }) {
  const [sortCol, setSortCol] = useState("fornecedor");
  const [sortDir, setSortDir] = useState("asc");
  const [tagFilter, setTagFilter] = useState("all");
  const [minEstoque, setMinEstoque] = useState("");
  const [maxEstoque, setMaxEstoque] = useState("");
  const [minC30, setMinC30] = useState("");
  const [maxC30, setMaxC30] = useState("");
  const [scrollWidth, setScrollWidth] = useState(0);
  const [selectedSupplier, setSelectedSupplier] = useState(null);
  const topScrollRef = React.useRef(null);
  const tableWrapRef = React.useRef(null);

  const measureScrollWidth = useCallback(() => { if (tableWrapRef.current) setScrollWidth(tableWrapRef.current.scrollWidth); }, []);
  useEffect(() => { measureScrollWidth(); window.addEventListener("resize", measureScrollWidth); return () => window.removeEventListener("resize", measureScrollWidth); }, [measureScrollWidth]);
  function syncFromTop(e) { if (tableWrapRef.current) tableWrapRef.current.scrollLeft = e.target.scrollLeft; }
  function syncFromTable(e) { if (topScrollRef.current) topScrollRef.current.scrollLeft = e.target.scrollLeft; }

  const minOrders = settings.supplierMinOrders || {};
  const supplierTags = settings.supplierTags || {};
  const customDaysMap = settings.supplierCustomDays || {};

  function toggleSort(col) {
    if (sortCol === col) setSortDir((d) => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("asc"); }
  }

  function SortIcon({ col }) {
    if (sortCol !== col) return <span style={{ opacity: 0.3, fontSize: 9 }}>↕</span>;
    return <span style={{ fontSize: 9 }}>{sortDir === "asc" ? "▲" : "▼"}</span>;
  }

  async function setMinOrder(fornecedor, value) {
    const next = { ...minOrders };
    if (value === "" || value === null) delete next[fornecedor];
    else next[fornecedor] = Number(value);
    await persistSettings({ ...settings, supplierMinOrders: next });
  }

  async function setCustomDays(fornecedor, value) {
    const next = { ...customDaysMap };
    if (value === "" || value === null) delete next[fornecedor];
    else next[fornecedor] = Number(value);
    await persistSettings({ ...settings, supplierCustomDays: next });
  }

  async function setTag(fornecedor, value) {
    const next = { ...supplierTags, [fornecedor]: value || null };
    await persistSettings({ ...settings, supplierTags: next });
  }

  const TAG_OPTIONS = [{ value: "", label: "—" }, { value: "escalando", label: "Escalando" }, { value: "liquidacao", label: "Liquidação" }, { value: "inativo", label: "Inativo" }];
  const TAG_STYLES = {
    escalando: { background: "#e1ecd9", color: "#3c6b2a", fontWeight: 600 },
    liquidacao: { background: "var(--rust-soft)", color: "var(--rust)", fontWeight: 600 },
    inativo: { background: "#e8e8e8", color: "#333", fontWeight: 600 },
  };

  const suppliers = useMemo(() => {
    const seen = new Set();
    return Object.values(products).filter((p) => !p.inactive && p.fornecedor).map((p) => p.fornecedor).filter((f) => { if (seen.has(f)) return false; seen.add(f); return true; });
  }, [products]);

  const rows = useMemo(() => suppliers.map((f) => {
    const prods = Object.values(products).filter((p) => !p.inactive && p.fornecedor === f);
    const estoque = prods.reduce((acc, p) => acc + (p.estoque || 0) * custoLiq(p), 0);
    const customDays = customDaysMap[f] || null;
    return {
      fornecedor: f, minOrder: minOrders[f] ?? "", tag: supplierTags[f] || "",
      customDays: customDays ?? "", estoque,
      c30: calcSupplierPurchase(products, f, 30), c45: calcSupplierPurchase(products, f, 45),
      c60: calcSupplierPurchase(products, f, 60), c90: calcSupplierPurchase(products, f, 90),
      cCustom: customDays ? calcSupplierPurchase(products, f, customDays) : null,
    };
  }), [suppliers, products, minOrders, supplierTags, customDaysMap]);

  const filtered = useMemo(() => {
    let out = rows;
    if (tagFilter !== "inativo") out = out.filter((r) => r.tag !== "inativo");
    if (tagFilter !== "all") out = out.filter((r) => r.tag === tagFilter);
    if (minEstoque !== "") out = out.filter((r) => r.estoque >= Number(minEstoque));
    if (maxEstoque !== "") out = out.filter((r) => r.estoque <= Number(maxEstoque));
    if (minC30 !== "") out = out.filter((r) => r.c30 >= Number(minC30));
    if (maxC30 !== "") out = out.filter((r) => r.c30 <= Number(maxC30));
    return [...out].sort((a, b) => {
      let va = a[sortCol], vb = b[sortCol];
      if (sortCol === "fornecedor" || sortCol === "tag") { va = String(va || ""); vb = String(vb || ""); return sortDir === "asc" ? va.localeCompare(vb, "pt-BR") : vb.localeCompare(va, "pt-BR"); }
      va = Number(va) || 0; vb = Number(vb) || 0;
      return sortDir === "asc" ? va - vb : vb - va;
    });
  }, [rows, tagFilter, minEstoque, maxEstoque, minC30, maxC30, sortCol, sortDir]);

  if (suppliers.length === 0) return <div className="vivo-page"><EmptyState icon={Package} title="Nenhum fornecedor" text="Importe relatórios para que os fornecedores apareçam aqui." /></div>;

  if (selectedSupplier) {
    return <SupplierDetail fornecedor={selectedSupplier} products={products} settings={settings} persistProducts={persistProducts} onBack={() => setSelectedSupplier(null)} />;
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Fornecedores</h1>
        <p>Valores de compra pelo <strong>custo de nota</strong>. Estoque R$ pelo <strong>custo líquido</strong>.</p>
      </header>

      <div className="vivo-toolbar">
        <div className="vivo-supplier-tag-chips">
          <span className="vivo-supplier-tag-label">Tag:</span>
          <button className={"vivo-supplier-tag-chip" + (tagFilter === "all" ? " is-active-neutral" : "")} onClick={() => setTagFilter("all")}>Todas</button>
          <button className={"vivo-supplier-tag-chip" + (tagFilter === "" ? " is-active-neutral" : "")} onClick={() => setTagFilter(tagFilter === "" ? "all" : "")}>Sem tag</button>
          <button className={"vivo-supplier-tag-chip vivo-stag-escalando" + (tagFilter === "escalando" ? " is-active" : "")} onClick={() => setTagFilter(tagFilter === "escalando" ? "all" : "escalando")}>Escalando</button>
          <button className={"vivo-supplier-tag-chip vivo-stag-liquidacao" + (tagFilter === "liquidacao" ? " is-active" : "")} onClick={() => setTagFilter(tagFilter === "liquidacao" ? "all" : "liquidacao")}>Liquidação</button>
          <button className={"vivo-supplier-tag-chip vivo-stag-inativo" + (tagFilter === "inativo" ? " is-active" : "")} onClick={() => setTagFilter(tagFilter === "inativo" ? "all" : "inativo")}>Inativo</button>
        </div>
        <div className="vivo-range-filter">
          <span>Estoque R$:</span>
          <input type="number" className="vivo-input-mini" placeholder="mín." value={minEstoque} onChange={(e) => setMinEstoque(e.target.value)} />
          <span>–</span>
          <input type="number" className="vivo-input-mini" placeholder="máx." value={maxEstoque} onChange={(e) => setMaxEstoque(e.target.value)} />
        </div>
        <div className="vivo-range-filter">
          <span>Compra 30D:</span>
          <input type="number" className="vivo-input-mini" placeholder="mín." value={minC30} onChange={(e) => setMinC30(e.target.value)} />
          <span>–</span>
          <input type="number" className="vivo-input-mini" placeholder="máx." value={maxC30} onChange={(e) => setMaxC30(e.target.value)} />
        </div>
        <button className="vivo-btn vivo-btn-ghost" onClick={() => { setTagFilter("all"); setMinEstoque(""); setMaxEstoque(""); setMinC30(""); setMaxC30(""); }}>Limpar filtros</button>
      </div>

      <div className="vivo-top-scrollbar" ref={topScrollRef} onScroll={syncFromTop}><div style={{ width: scrollWidth, height: 1 }} /></div>

      <div className="vivo-table-wrap vivo-table-wrap-sticky vivo-card" ref={tableWrapRef} onScroll={syncFromTable}>
        <table className="vivo-table vivo-table-suppliers">
          <thead>
            <tr>
              <th onClick={() => toggleSort("fornecedor")} className="vivo-th-sortable">Fornecedor <SortIcon col="fornecedor" /></th>
              <th onClick={() => toggleSort("tag")} className="vivo-th-sortable">Tag <SortIcon col="tag" /></th>
              <th className="num vivo-th-custom">Cobertura (dias)</th>
              <th onClick={() => toggleSort("cCustom")} className="num vivo-th-sortable vivo-th-custom">Compra Personalizada <SortIcon col="cCustom" /></th>
              <th className="vivo-th-custom"></th>
              <th className="num">Pedido Mínimo (R$)</th>
              <th onClick={() => toggleSort("estoque")} className="num vivo-th-sortable">Estoque R$ <SortIcon col="estoque" /></th>
              <th onClick={() => toggleSort("c30")} className="num vivo-th-sortable">Compra 30D <SortIcon col="c30" /></th>
              <th onClick={() => toggleSort("c45")} className="num vivo-th-sortable">Compra 45D <SortIcon col="c45" /></th>
              <th onClick={() => toggleSort("c60")} className="num vivo-th-sortable">Compra 60D <SortIcon col="c60" /></th>
              <th onClick={() => toggleSort("c90")} className="num vivo-th-sortable">Compra 90D <SortIcon col="c90" /></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((r, i) => (
              <tr key={r.fornecedor} className={i % 2 === 0 ? "vivo-row-even" : "vivo-row-odd"}>
                <td className="vivo-supplier-name vivo-supplier-name-link" onClick={() => setSelectedSupplier(r.fornecedor)} title="Clique para ver detalhes">{r.fornecedor}</td>
                <td>
                  <select className="vivo-select-tag" style={r.tag ? { ...TAG_STYLES[r.tag], border: "none", borderRadius: 6, padding: "3px 8px" } : {}} value={r.tag} onChange={(e) => setTag(r.fornecedor, e.target.value)}>
                    {TAG_OPTIONS.map((t) => <option key={t.value} value={t.value}>{t.label}</option>)}
                  </select>
                </td>
                <td className="num vivo-td-custom">
                  <input type="number" min="1" className="vivo-input-mini" placeholder="—" value={r.customDays} onChange={(e) => setCustomDays(r.fornecedor, e.target.value)} />
                </td>
                <td className={"num mono vivo-td-custom" + (r.minOrder && r.cCustom !== null && r.cCustom < r.minOrder ? " vivo-below-min" : r.minOrder && r.cCustom !== null && r.cCustom >= r.minOrder ? " vivo-above-min" : "")}>
                  {r.cCustom !== null ? fmtCurrency(r.cCustom) : "—"}
                </td>
                <td className="vivo-td-custom">
                  <button className="vivo-btn vivo-btn-primary" disabled={!r.customDays} onClick={() => onGenerateOrder({ fornecedor: r.fornecedor, days: Number(r.customDays), ts: Date.now() })}>
                    Gerar Pedido
                  </button>
                </td>
                <td className="num">
                  <input type="number" min="0" className="vivo-input-mini vivo-input-min-order" placeholder="—" value={r.minOrder} onChange={(e) => setMinOrder(r.fornecedor, e.target.value)} />
                </td>
                <td className="num mono">{fmtCurrency(r.estoque)}</td>
                <td className={"num mono" + (r.minOrder && r.c30 < r.minOrder ? " vivo-below-min" : r.minOrder && r.c30 >= r.minOrder ? " vivo-above-min" : "")}>{fmtCurrency(r.c30)}</td>
                <td className={"num mono" + (r.minOrder && r.c45 < r.minOrder ? " vivo-below-min" : r.minOrder && r.c45 >= r.minOrder ? " vivo-above-min" : "")}>{fmtCurrency(r.c45)}</td>
                <td className={"num mono" + (r.minOrder && r.c60 < r.minOrder ? " vivo-below-min" : r.minOrder && r.c60 >= r.minOrder ? " vivo-above-min" : "")}>{fmtCurrency(r.c60)}</td>
                <td className={"num mono" + (r.minOrder && r.c90 < r.minOrder ? " vivo-below-min" : r.minOrder && r.c90 >= r.minOrder ? " vivo-above-min" : "")}>{fmtCurrency(r.c90)}</td>
              </tr>
            ))}
            {filtered.length === 0 && <tr><td colSpan={11} className="vivo-table-empty">Nenhum fornecedor encontrado.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function HistoryTab({ history, persistHistory }) {
  if (history.length === 0) return <div className="vivo-page"><EmptyState icon={History} title="Nenhuma importação" text="O histórico aparece aqui." /></div>;
  return (
    <div className="vivo-page">
      <header className="vivo-page-head"><h1>Histórico de importações</h1></header>
      <div className="vivo-table-wrap vivo-card">
        <table className="vivo-table">
          <thead><tr><th>Data</th><th>Arquivo</th><th className="num">Linhas</th><th className="num">Importadas</th><th className="num">Ignoradas</th><th className="num">Novos</th><th className="num">Atualizados</th></tr></thead>
          <tbody>
            {history.map((h) => (
              <tr key={h.id}>
                <td>{fmtDate(h.ts)}</td><td>{h.fileName}</td>
                <td className="num mono">{h.totalRows}</td><td className="num mono">{h.importedRows}</td>
                <td className="num mono">{h.skippedNoSupplier}</td><td className="num mono">{h.newCount}</td>
                <td className="num mono">{h.updatedCount}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function StorageDiagnosticPanel({ diag, migrationState, onCheck, onMigrate, onExport, onImportFile, importState, onClearHistory }) {
  return (
    <div className="vivo-card vivo-diagnostic-panel">
      <div className="vivo-diagnostic-head">
        <span className="vivo-detail-label">Diagnóstico de armazenamento</span>
        <button className="vivo-btn vivo-btn-ghost" onClick={onCheck}>Verificar agora</button>
      </div>
      {diag.checked && (
        <div className="vivo-diagnostic-result">
          <span>Compartilhado: <strong>{diag.sharedCount}</strong> produtos</span>
          <span>Privado: <strong>{diag.privateCount}</strong> produtos</span>
        </div>
      )}
      {diag.checked && diag.privateCount > 0 && (
        <div className="vivo-diagnostic-action">
          <button className="vivo-btn vivo-btn-primary" onClick={onMigrate} disabled={migrationState.running}>
            {migrationState.running ? "Migrando…" : "Forçar migração"}
          </button>
        </div>
      )}
      {migrationState.done && <div className="vivo-diagnostic-success"><CheckCircle2 size={14} /> Migração concluída.</div>}
      {migrationState.error && <div className="vivo-diagnostic-error"><AlertTriangle size={14} /> {migrationState.error}</div>}
      <div className="vivo-diagnostic-divider" />
      <div className="vivo-diagnostic-backup">
        <span className="vivo-detail-label" style={{ display: "block", marginBottom: 6 }}>Backup e restauração</span>
        <div className="vivo-diagnostic-backup-actions">
          <button className="vivo-btn vivo-btn-secondary" onClick={() => onExport(false)}>Exportar resumido</button>
          <button className="vivo-btn vivo-btn-ghost" onClick={() => onExport(true)}>Exportar completo</button>
          <label className="vivo-btn vivo-btn-secondary">Importar backup<input type="file" accept=".json" style={{ display: "none" }} onChange={onImportFile} /></label>
        </div>
        {importState?.progress && !importState?.done && <div className="vivo-diagnostic-progress">{importState.progress}</div>}
        {importState?.done && <div className="vivo-diagnostic-success"><CheckCircle2 size={14} /> Backup importado.</div>}
        {importState?.error && <div className="vivo-diagnostic-error"><AlertTriangle size={14} /> {importState.error}</div>}
      </div>
      <div className="vivo-diagnostic-divider" />
      <div className="vivo-diagnostic-backup">
        <span className="vivo-detail-label" style={{ display: "block", marginBottom: 6 }}>Reduzir tamanho dos dados</span>
        <button className="vivo-btn vivo-btn-danger" onClick={onClearHistory}>Limpar histórico de vendas</button>
      </div>
    </div>
  );
}

function CleanupTab({ products, persistProducts, history, settings, orders, persistHistory, persistSettings, persistOrders }) {
  const [selected, setSelected] = useState({});
  const [showInactive, setShowInactive] = useState(false);
  const [diag, setDiag] = useState({ checked: false, sharedCount: 0, privateCount: 0 });
  const [migrationState, setMigrationState] = useState({ running: false, done: false, error: null });
  const [importState, setImportState] = useState({ done: false, error: null });

  async function checkStorage() {
    setDiag((d) => ({ ...d, checked: false }));
    const [sharedProducts, privateProducts] = await Promise.all([storageGet(K_PRODUCTS, {}), storageGetPrivate(K_PRODUCTS, {})]);
    setDiag({ checked: true, sharedCount: Object.keys(sharedProducts).length, privateCount: Object.keys(privateProducts).length });
  }

  async function forceMigration() {
    setMigrationState({ running: true, done: false, error: null });
    try {
      const [p, h, s, o] = await Promise.all([storageGetPrivate(K_PRODUCTS, {}), storageGetPrivate(K_HISTORY, []), storageGetPrivate(K_SETTINGS, {}), storageGetPrivate(K_ORDERS, [])]);
      await Promise.all([storageSet(K_PRODUCTS, p), storageSet(K_HISTORY, h), storageSet(K_SETTINGS, s), storageSet(K_ORDERS, o)]);
      await persistProducts(p);
      setMigrationState({ running: false, done: true, error: null });
      await checkStorage();
    } catch (e) {
      setMigrationState({ running: false, done: false, error: "Não foi possível migrar." });
    }
  }

  function exportBackup(full) {
    try {
      let exportProducts = products;
      if (!full) {
        exportProducts = {};
        for (const codigo of Object.keys(products)) { const { salesHistory, ...rest } = products[codigo]; exportProducts[codigo] = rest; }
      }
      const json = JSON.stringify({ exportedAt: Date.now(), version: APP_VERSION, full, products: exportProducts, history, settings, orders });
      const blob = new Blob([json], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `vivo-backup-${full ? "completo" : "resumido"}-${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (e) { alert("Não foi possível gerar o backup. Tente a opção resumida."); }
  }

  function importBackupFile(e) {
    const file = e.target.files?.[0]; if (!file) return;
    setImportState({ done: false, error: null, progress: "Lendo arquivo…" });
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const backup = JSON.parse(ev.target.result);
        if (!backup.products) throw new Error("Inválido.");
        setImportState((s) => ({ ...s, progress: "Salvando produtos…" }));
        await persistProducts(backup.products || {});
        setImportState((s) => ({ ...s, progress: "Salvando histórico…" }));
        await persistHistory(backup.history || []);
        setImportState((s) => ({ ...s, progress: "Salvando configurações…" }));
        await persistSettings(backup.settings || {});
        setImportState((s) => ({ ...s, progress: "Salvando pedidos…" }));
        await persistOrders(backup.orders || []);
        setImportState({ done: true, error: null, progress: null });
        setTimeout(() => window.location.reload(), 1500);
      } catch (err) { setImportState({ done: false, error: "Arquivo inválido.", progress: null }); }
    };
    reader.onerror = () => setImportState({ done: false, error: "Não foi possível ler.", progress: null });
    reader.readAsText(file);
  }

  function clearAllSalesHistory() {
    if (!window.confirm("Remove o histórico de vendas de todos os produtos. Continuar?")) return;
    const next = {};
    for (const codigo of Object.keys(products)) { const { salesHistory, ...rest } = products[codigo]; next[codigo] = rest; }
    persistProducts(next);
  }

  const { lastImportTs, missing } = useMemo(() => getMissingFromLastImport(products, history), [products, history]);
  const { page, setPage, totalPages, pageItems, totalCount } = usePagination(missing);
  const inactiveList = useMemo(() => Object.values(products).filter((p) => p.inactive), [products]);

  function toggleSelect(codigo) { setSelected((s) => ({ ...s, [codigo]: !s[codigo] })); }
  function selectAll() { const next = {}; missing.forEach((p) => { next[p.codigo] = true; }); setSelected(next); }
  function clearSelection() { setSelected({}); }
  const selectedCount = missing.filter((p) => selected[p.codigo]).length;

  async function markSelectedInactive() {
    if (selectedCount === 0) return;
    const next = { ...products };
    for (const p of missing) { if (!selected[p.codigo]) continue; next[p.codigo] = { ...next[p.codigo], inactive: true, inactivatedAt: Date.now() }; }
    await persistProducts(next); setSelected({});
  }

  async function reactivate(codigo) {
    const next = { ...products }; next[codigo] = { ...next[codigo], inactive: false, inactivatedAt: null }; await persistProducts(next);
  }

  if (!history || history.length === 0) {
    return (
      <div className="vivo-page">
        <header className="vivo-page-head"><h1>Limpeza de itens</h1></header>
        <StorageDiagnosticPanel diag={diag} migrationState={migrationState} onCheck={checkStorage} onMigrate={forceMigration} onExport={exportBackup} onImportFile={importBackupFile} importState={importState} onClearHistory={clearAllSalesHistory} />
        <EmptyState icon={Trash2} title="Nenhuma importação ainda" text="Importe ao menos um relatório." />
      </div>
    );
  }

  return (
    <div className="vivo-page">
      <header className="vivo-page-head">
        <h1>Limpeza de itens</h1>
        <p>Última importação: <strong>{fmtDate(lastImportTs)}</strong>. Itens que não apareceram no último envio.</p>
      </header>
      <StorageDiagnosticPanel diag={diag} migrationState={migrationState} onCheck={checkStorage} onMigrate={forceMigration} onExport={exportBackup} onImportFile={importBackupFile} importState={importState} onClearHistory={clearAllSalesHistory} />

      {missing.length === 0 ? (
        <EmptyState icon={CheckCircle2} title="Nenhum item para limpar" text="Todos os produtos aparecem na última importação." />
      ) : (
        <>
          <div className="vivo-action-row">
            <button className="vivo-btn vivo-btn-ghost" onClick={selectAll}>Selecionar todos ({missing.length})</button>
            <button className="vivo-btn vivo-btn-ghost" onClick={clearSelection}>Limpar</button>
            <div className="vivo-action-spacer" />
            <span className="vivo-action-summary">{selectedCount} selecionado(s)</span>
            <button className="vivo-btn vivo-btn-danger" onClick={markSelectedInactive} disabled={selectedCount === 0}>Marcar como inativo</button>
          </div>
          <div className="vivo-table-wrap vivo-card">
            <table className="vivo-table vivo-table-compact">
              <thead><tr><th></th><th>Item</th><th>Fornecedor</th><th className="mono">Código</th><th className="num">Estoque</th><th>Última atualização</th></tr></thead>
              <tbody>
                {pageItems.map((p) => (
                  <tr key={p.codigo} className={selected[p.codigo] ? "is-selected" : ""}>
                    <td><input type="checkbox" checked={!!selected[p.codigo]} onChange={() => toggleSelect(p.codigo)} /></td>
                    <ItemCell name={p.item} />
                    <td>{p.fornecedor}</td><td className="mono">{p.codigo}</td>
                    <td className="num mono">{fmtNumber(p.estoque)}</td><td className="mono">{fmtDate(p.lastUpdated)}</td>
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
          {showInactive ? "Ocultar" : "Ver"} produtos inativos ({inactiveList.length})
        </button>
        {showInactive && inactiveList.length > 0 && (
          <div className="vivo-table-wrap vivo-card" style={{ marginTop: 12 }}>
            <table className="vivo-table vivo-table-compact">
              <thead><tr><th>Item</th><th>Fornecedor</th><th className="mono">Código</th><th>Inativado em</th><th></th></tr></thead>
              <tbody>
                {inactiveList.map((p) => (
                  <tr key={p.codigo}>
                    <ItemCell name={p.item} /><td>{p.fornecedor}</td><td className="mono">{p.codigo}</td>
                    <td className="mono">{p.inactivatedAt ? fmtDate(p.inactivatedAt) : "—"}</td>
                    <td><button className="vivo-toggle-btn" onClick={() => reactivate(p.codigo)}>Reativar</button></td>
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

function ItemCell({ name, className = "" }) {
  const [expanded, setExpanded] = useState(false);
  return (
    <td className={"vivo-item-cell" + (expanded ? " is-expanded" : "") + (className ? " " + className : "")} onClick={() => setExpanded((e) => !e)} title={expanded ? "Clique para recolher" : name} style={{ cursor: "pointer" }}>
      {name}
    </td>
  );
}

const PAGE_SIZE = 50;

function usePagination(items, pageSize = PAGE_SIZE) {
  const [page, setPage] = useState(1);
  const totalPages = Math.max(1, Math.ceil(items.length / pageSize));
  useEffect(() => { setPage(1); }, [items.length]);
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

function EmptyState({ icon: Icon, title, text }) {
  return (
    <div className="vivo-empty">
      <Icon size={32} strokeWidth={1.5} />
      <h3>{title}</h3>
      <p>{text}</p>
    </div>
  );
}

const VIVO_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

.vivo-root {
  --ink: #2b2620; --ink-soft: #6b6358; --paper: #faf7f0; --card: #ffffff;
  --line: #e6ddc9; --olive: #5b6b4f; --olive-dark: #41503a;
  --amber: #c1762f; --amber-soft: #f3e3cf; --rust: #a8472f; --rust-soft: #f7e4dd;
  --sidebar-bg: #232017; --sidebar-line: #38332a;
  display: flex; flex-direction: column; min-height: 100%; width: 100%;
  background: var(--paper); color: var(--ink);
  font-family: 'Inter', -apple-system, sans-serif; font-size: 13px; line-height: 1.4;
}
.vivo-root * { box-sizing: border-box; }
.vivo-loading { align-items: center; justify-content: center; flex-direction: column; gap: 10px; width: 100%; min-height: 480px; }
.vivo-loading-mark { font-family: 'Fraunces', serif; font-size: 28px; letter-spacing: 0.08em; color: var(--olive-dark); }
.vivo-loading-text { color: var(--ink-soft); font-size: 13px; }
.vivo-topnav { background: var(--sidebar-bg); color: #e9e3d6; flex-shrink: 0; }
.vivo-topnav-row { display: flex; align-items: center; gap: 4px; padding: 0 28px; }
.vivo-topnav-main { height: 60px; border-bottom: 1px solid var(--sidebar-line); }
.vivo-brand { display: flex; align-items: center; gap: 9px; margin-right: 28px; }
.vivo-brand-mark { color: var(--amber); font-size: 16px; }
.vivo-brand-name { font-family: 'Fraunces', serif; font-size: 16px; font-weight: 600; letter-spacing: 0.01em; color: #fbf8f0; line-height: 1.1; }
.vivo-brand-sub { font-size: 10.5px; color: #a39c8a; margin-top: 1px; }
.vivo-main-tabs { display: flex; align-items: center; gap: 4px; height: 100%; }
.vivo-main-tab-btn { display: flex; align-items: center; gap: 8px; background: transparent; border: none; color: #c8c0ac; padding: 0 16px; height: 100%; font-size: 13.5px; font-weight: 500; font-family: inherit; cursor: pointer; position: relative; transition: color 0.15s; }
.vivo-main-tab-btn:hover { color: #fbf8f0; }
.vivo-main-tab-btn.is-active { color: #fff; }
.vivo-main-tab-btn.is-active::after { content: ""; position: absolute; left: 16px; right: 16px; bottom: 0; height: 2px; background: var(--amber); }
.vivo-topnav-spacer { flex: 1; }
.vivo-topnav-hint { font-size: 11.5px; color: #847d6c; white-space: nowrap; }
.vivo-sub-tabs { height: 46px; gap: 6px; }
.vivo-sub-tab-btn { display: flex; align-items: center; gap: 7px; background: transparent; border: none; color: #b3ab97; padding: 7px 13px; border-radius: 7px; font-size: 13px; font-family: inherit; cursor: pointer; transition: background 0.15s, color 0.15s; }
.vivo-sub-tab-btn:hover { background: #2e2a20; color: #fbf8f0; }
.vivo-sub-tab-btn.is-active { background: #383228; color: #fff; }
.vivo-sub-tab-count { background: rgba(255,255,255,0.14); border-radius: 100px; font-size: 10.5px; padding: 1px 6px; font-family: 'JetBrains Mono', monospace; }
.vivo-main { flex: 1; min-width: 0; overflow-x: auto; }
.vivo-page { padding: 30px 36px 50px; max-width: 1200px; }
.vivo-page-head { margin-bottom: 22px; }
.vivo-page-head h1 { font-family: 'Fraunces', serif; font-size: 26px; font-weight: 600; margin: 0 0 5px; color: var(--ink); }
.vivo-page-head p { color: var(--ink-soft); margin: 0; font-size: 13.5px; max-width: 640px; }
.vivo-dropzone { border: 1.5px dashed var(--line); border-radius: 12px; background: var(--card); padding: 40px 24px; text-align: center; display: flex; flex-direction: column; align-items: center; gap: 6px; color: var(--olive-dark); transition: border-color 0.15s; max-width: 560px; }
.vivo-dropzone.is-drag { border-color: var(--amber); background: var(--amber-soft); }
.vivo-dropzone-title { font-weight: 600; font-size: 15px; margin-top: 6px; color: var(--ink); }
.vivo-dropzone-sub { color: var(--ink-soft); font-size: 12px; margin: 0; }
.vivo-dropzone-hint { color: var(--ink-soft); font-size: 12px; margin-top: 10px; max-width: 380px; }
.vivo-btn { display: inline-flex; align-items: center; gap: 7px; border-radius: 7px; padding: 8px 14px; font-size: 13px; font-weight: 500; font-family: inherit; cursor: pointer; border: 1px solid transparent; white-space: nowrap; transition: opacity 0.15s, background 0.15s; }
.vivo-btn:disabled { opacity: 0.4; cursor: not-allowed; }
.vivo-btn-primary { background: var(--olive-dark); color: #fff; }
.vivo-btn-primary:hover:not(:disabled) { background: var(--olive); }
.vivo-btn-secondary { background: var(--amber-soft); color: #6b4419; border-color: #e6cda4; }
.vivo-btn-secondary:hover:not(:disabled) { background: #ecd6b4; }
.vivo-btn-ghost { background: transparent; color: var(--ink-soft); border-color: var(--line); }
.vivo-btn-ghost:hover:not(:disabled) { background: #f1ece0; color: var(--ink); }
.vivo-btn-danger { background: var(--rust-soft); color: var(--rust); border-color: #e3b4a3; }
.vivo-btn-danger:hover:not(:disabled) { background: #f3c4b3; }
.vivo-btn-loading { opacity: 0.7; cursor: not-allowed; }
.vivo-alert { display: flex; align-items: center; gap: 8px; padding: 10px 14px; border-radius: 8px; font-size: 13px; margin-top: 16px; max-width: 560px; }
.vivo-alert-error { background: var(--rust-soft); color: #7a2c19; }
.vivo-alert-warning { background: var(--amber-soft); color: #6b4419; max-width: 720px; }
.vivo-migration-banner { display: flex; align-items: center; gap: 12px; padding: 12px 24px; background: var(--amber-soft); color: #6b4419; font-size: 13px; border-bottom: 1px solid #e6cda4; }
.vivo-migration-banner span { flex: 1; }
.vivo-migration-success { background: #e7ede2; color: var(--olive-dark); }
.vivo-migration-error { color: var(--rust); font-size: 12px; }
.vivo-diagnostic-panel { padding: 16px 20px; margin-bottom: 18px; max-width: 640px; }
.vivo-diagnostic-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px; }
.vivo-diagnostic-result { display: flex; flex-direction: column; gap: 4px; margin-top: 10px; font-size: 13px; }
.vivo-diagnostic-action { margin-top: 6px; }
.vivo-diagnostic-success { display: flex; align-items: center; gap: 6px; margin-top: 10px; font-size: 12.5px; color: var(--olive-dark); }
.vivo-diagnostic-error { display: flex; align-items: center; gap: 6px; margin-top: 10px; font-size: 12.5px; color: var(--rust); }
.vivo-diagnostic-progress { margin-top: 10px; font-size: 12.5px; color: var(--ink-soft); font-family: 'JetBrains Mono', monospace; }
.vivo-diagnostic-divider { height: 1px; background: var(--line); margin: 16px 0; }
.vivo-diagnostic-backup-actions { display: flex; gap: 10px; flex-wrap: wrap; }
.vivo-cleanup-inactive-section { margin-top: 28px; }
.vivo-table-suppliers .vivo-supplier-name { font-weight: 600; color: var(--ink); font-size: 12.5px; }
.vivo-row-even { background: var(--card); }
.vivo-row-odd { background: #f7f3ec; }
.vivo-row-even:hover, .vivo-row-odd:hover { background: #f1ece0; }
.vivo-input-min-order { width: 100px; text-align: right; }
.vivo-below-min { color: var(--rust); font-weight: 600; }
.vivo-above-min { color: var(--olive-dark); font-weight: 600; }
.vivo-th-sortable { cursor: pointer; user-select: none; white-space: nowrap; }
.vivo-th-sortable:hover { background: #f0ebdd; }
.vivo-th-custom { background: #f0f4ee !important; color: var(--olive-dark); border-left: 2px solid #c7d4b8; }
.vivo-td-custom { background: #f5f8f3; border-left: 2px solid #c7d4b8; }
.vivo-detail-back { display: inline-flex; align-items: center; gap: 6px; font-size: 12.5px; color: var(--ink-soft); cursor: pointer; padding: 4px 0; margin-bottom: 4px; }
.vivo-detail-back:hover { color: var(--olive-dark); }
.vivo-supplier-name-link { cursor: pointer; color: var(--olive-dark) !important; text-decoration: underline; text-underline-offset: 3px; }
.vivo-supplier-name-link:hover { color: var(--amber) !important; }
.vivo-supplier-kpis { display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 18px; }
.vivo-kpi-card { background: var(--card); border: 1px solid var(--line); border-radius: 10px; padding: 12px 18px; display: flex; flex-direction: column; gap: 4px; min-width: 130px; }
.vivo-kpi-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--ink-soft); font-weight: 600; }
.vivo-kpi-value { font-size: 18px; font-weight: 700; color: var(--ink); font-family: 'JetBrains Mono', monospace; }
.vivo-kpi-danger { color: var(--rust) !important; }
.vivo-kpi-warn { color: var(--amber) !important; }
.vivo-price-status-cards { display: flex; gap: 12px; margin-bottom: 18px; flex-wrap: wrap; }
.vivo-price-upload-card { padding: 20px 24px; margin-bottom: 16px; }
.vivo-price-result { padding: 16px 20px; display: flex; flex-direction: column; gap: 10px; }
.vivo-price-result-item { display: flex; align-items: center; gap: 8px; font-size: 13px; }
.vivo-supplier-charts { display: flex; flex-direction: column; gap: 16px; margin-bottom: 4px; }
.vivo-supplier-chart-block { padding: 18px 20px; }
.vivo-supplier-chart-svg { width: 100%; height: 140px; display: block; }
.vivo-minichart-header { display: flex; justify-content: space-between; align-items: baseline; flex-wrap: wrap; gap: 8px; margin-bottom: 10px; }
.vivo-minichart-current { font-size: 15px; font-family: 'JetBrains Mono', monospace; }
.vivo-badge { display: inline-block; padding: 2px 8px; border-radius: 100px; font-size: 11px; font-weight: 600; }
.vivo-badge-danger { background: var(--rust-soft); color: var(--rust); }
.vivo-badge-warn { background: var(--amber-soft); color: #7a4a10; }
.vivo-badge-ok { background: #e1ecd9; color: #3c6b2a; }
.vivo-sort-btn { display: inline-flex; align-items: center; justify-content: center; background: none; border: none; cursor: pointer; font-size: 10px; color: var(--ink-soft); padding: 1px 3px; border-radius: 3px; line-height: 1; margin-left: 2px; }
.vivo-sort-btn:hover { background: var(--line); color: var(--ink); }
.vivo-drawer-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.15); z-index: 100; }
.vivo-drawer { position: fixed; top: 0; right: 0; bottom: 0; width: min(520px, 95vw); background: var(--paper); box-shadow: -4px 0 24px rgba(0,0,0,0.12); z-index: 101; display: flex; flex-direction: column; overflow: hidden; }
.vivo-drawer-header { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; padding: 18px 20px 14px; border-bottom: 1px solid var(--line); background: var(--card); }
.vivo-drawer-title { font-size: 14px; font-weight: 600; color: var(--ink); line-height: 1.3; }
.vivo-drawer-sub { font-size: 11.5px; color: var(--ink-soft); margin-top: 3px; font-family: monospace; }
.vivo-drawer-close { flex-shrink: 0; background: none; border: none; cursor: pointer; font-size: 16px; color: var(--ink-soft); padding: 2px 6px; border-radius: 4px; }
.vivo-drawer-close:hover { background: var(--line); color: var(--ink); }
.vivo-drawer-body { flex: 1; overflow-y: auto; padding: 18px 20px; display: flex; flex-direction: column; gap: 20px; }
.vivo-drawer-chart-block { background: var(--card); border: 1px solid var(--line); border-radius: 10px; padding: 14px 16px; }
.vivo-drawer-obs-section { display: flex; flex-direction: column; gap: 10px; }
.vivo-drawer-obs-input-row { display: flex; gap: 8px; align-items: flex-end; }
.vivo-drawer-obs-input { flex: 1; font-family: inherit; font-size: 13px; padding: 8px 12px; border: 1px solid var(--line); border-radius: 8px; background: var(--card); color: var(--ink); resize: none; line-height: 1.4; }
.vivo-drawer-obs-input:focus { outline: none; border-color: var(--olive-dark); }
.vivo-drawer-obs-hint { font-size: 11px; color: var(--ink-soft); }
.vivo-drawer-obs-list { display: flex; flex-direction: column; gap: 8px; }
.vivo-drawer-obs-item { display: flex; flex-direction: column; gap: 2px; padding: 10px 12px; background: var(--card); border: 1px solid var(--line); border-radius: 8px; }
.vivo-drawer-obs-date { font-size: 10.5px; color: var(--ink-soft); font-family: monospace; }
.vivo-drawer-obs-text { font-size: 13px; color: var(--ink); line-height: 1.4; white-space: pre-wrap; }
.vivo-drawer-obs-empty { font-size: 12.5px; color: var(--ink-soft); font-style: italic; margin: 0; }
.vivo-supplier-tag-chips { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
.vivo-generated-order-summary { display: flex; align-items: center; gap: 24px; padding: 16px 20px; background: var(--card); border: 1px solid var(--line); border-radius: 10px; margin-bottom: 16px; flex-wrap: wrap; }
.vivo-generated-order-total { display: flex; flex-direction: column; gap: 2px; }
.vivo-generated-order-total span { font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--ink-soft); }
.vivo-generated-order-total strong { font-size: 22px; font-weight: 700; color: var(--olive-dark); font-family: 'JetBrains Mono', monospace; }
.vivo-generated-order-meta { font-size: 12.5px; color: var(--ink-soft); flex: 1; }
.vivo-row-zero { opacity: 0.45; }
.vivo-value-positive { color: var(--olive-dark); font-weight: 600; }
.vivo-sku-days-custom { border-color: var(--amber); background: var(--amber-soft); font-weight: 600; }
.vivo-supplier-tag-label { font-size: 12px; color: var(--ink-soft); white-space: nowrap; }
.vivo-supplier-tag-chip { font-family: inherit; font-size: 12px; font-weight: 500; padding: 5px 12px; border-radius: 100px; border: 1px solid var(--line); background: var(--card); color: var(--ink-soft); cursor: pointer; transition: background 0.12s; }
.vivo-supplier-tag-chip:hover { background: #f1ece0; }
.vivo-supplier-tag-chip.is-active-neutral { background: var(--ink-soft); color: #fff; border-color: var(--ink-soft); }
.vivo-stag-escalando { border-color: #aed4c4; color: #2c6650; }
.vivo-stag-escalando.is-active { background: #3c6b2a; color: #fff; border-color: #3c6b2a; }
.vivo-stag-liquidacao { border-color: #e3b4a3; color: var(--rust); }
.vivo-stag-liquidacao.is-active { background: var(--rust); color: #fff; border-color: var(--rust); }
.vivo-stag-inativo { border-color: #ccc; color: #444; }
.vivo-stag-inativo.is-active { background: #333; color: #fff; border-color: #333; }
.vivo-pagination { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-top: 12px; flex-wrap: wrap; }
.vivo-pagination-info { font-size: 12.5px; color: var(--ink-soft); font-family: 'JetBrains Mono', monospace; }
.vivo-pagination-buttons { display: flex; align-items: center; gap: 6px; }
.vivo-pagination-page { font-size: 12.5px; color: var(--ink-soft); padding: 0 4px; white-space: nowrap; }
.vivo-card { background: var(--card); border: 1px solid var(--line); border-radius: 12px; }
.vivo-preview { margin-top: 22px; padding: 18px 20px; max-width: 920px; }
.vivo-preview-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; margin-bottom: 14px; }
.vivo-preview-head h3 { font-size: 14px; font-weight: 600; margin: 0 0 5px; }
.vivo-pending-files { display: flex; flex-direction: column; gap: 8px; }
.vivo-pending-file { display: flex; align-items: center; gap: 16px; flex-wrap: wrap; background: #fbf8f0; border: 1px solid var(--line); border-radius: 8px; padding: 10px 14px; }
.vivo-pending-file-name { display: flex; align-items: center; gap: 7px; font-size: 13px; font-weight: 500; color: var(--ink); min-width: 180px; }
.vivo-pending-file-stats { display: flex; gap: 12px; font-size: 12px; color: var(--ink-soft); flex-wrap: wrap; }
.vivo-pending-file-stats .ok { color: var(--olive-dark); display: inline-flex; align-items: center; gap: 4px; }
.vivo-pending-file-stats .warn { color: var(--rust); display: inline-flex; align-items: center; gap: 4px; }
.vivo-pending-file-error { color: var(--rust); font-size: 12.5px; display: inline-flex; align-items: center; gap: 5px; }
.vivo-period-select { display: flex; align-items: center; gap: 7px; font-size: 12.5px; color: var(--ink-soft); margin-left: auto; }
.vivo-table-wrap { overflow-x: auto; }
.vivo-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.vivo-table thead th { text-align: left; font-size: 10px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-soft); font-weight: 600; padding: 7px 8px; border-bottom: 1px solid var(--line); white-space: nowrap; }
.vivo-table th.num, .vivo-table td.num { text-align: right; }
.vivo-table tbody td { padding: 5px 8px; border-bottom: 1px solid #f0ebdd; vertical-align: middle; font-size: 12px; white-space: nowrap; }
.vivo-table tbody tr:hover { background: #fbf8f0; }
.vivo-table tbody tr.is-low { background: var(--rust-soft); }
.vivo-table tbody tr.is-selected { background: var(--amber-soft); }
.vivo-table .mono { font-family: 'JetBrains Mono', monospace; font-size: 11px; color: var(--ink-soft); }
.vivo-item-cell { font-weight: 500; color: var(--ink); max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.vivo-item-cell.is-expanded { white-space: normal; overflow: visible; max-width: 320px; word-break: break-word; }
.vivo-suggested { color: var(--amber); font-weight: 600; }
.vivo-trend-up { color: var(--olive-dark); font-weight: 600; }
.vivo-trend-down { color: var(--rust); font-weight: 600; }
.vivo-trend-summary { display: flex; align-items: center; gap: 8px; margin-bottom: 16px; flex-wrap: wrap; }
.vivo-trend-chip { display: inline-flex; align-items: center; gap: 6px; border: 1px solid var(--line); background: var(--card); border-radius: 100px; padding: 6px 13px; font-size: 12.5px; font-family: inherit; cursor: pointer; color: var(--ink-soft); transition: background 0.15s; }
.vivo-trend-chip:hover { background: #f1ece0; }
.vivo-trend-chip-up.is-active { background: var(--olive-dark); color: #fff; border-color: var(--olive-dark); }
.vivo-trend-chip-down.is-active { background: var(--rust); color: #fff; border-color: var(--rust); }
.vivo-trend-chip-flat.is-active { background: var(--ink-soft); color: #fff; border-color: var(--ink-soft); }
.vivo-badge { display: inline-flex; align-items: center; gap: 5px; font-size: 12px; font-weight: 500; padding: 3px 9px; border-radius: 100px; }
.vivo-badge-up { background: #e7ede2; color: var(--olive-dark); }
.vivo-badge-down { background: var(--rust-soft); color: var(--rust); }
.vivo-badge-flat { background: #f0ebdd; color: var(--ink-soft); }
.vivo-badge-muted { background: #f0ebdd; color: var(--ink-soft); font-style: italic; }
.vivo-range-filter { display: flex; align-items: center; gap: 6px; font-size: 12.5px; color: var(--ink-soft); background: var(--card); border: 1px solid var(--line); border-radius: 7px; padding: 6px 10px; }
.vivo-range-filter .vivo-input-mini { width: 56px; }
.vivo-quality-summary { display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--ink-soft); margin-bottom: 14px; }
.vivo-quality-summary strong { color: var(--ink); }
.vivo-capital-chart { background: var(--card); border: 1px solid var(--line); border-radius: 10px; padding: 12px 16px 8px; margin-bottom: 14px; }
.vivo-capital-chart-empty { display: flex; align-items: center; font-size: 12px; color: var(--ink-soft); padding: 14px 16px; min-height: 56px; }
.vivo-capital-chart-header { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; margin-bottom: 6px; flex-wrap: wrap; }
.vivo-capital-chart-label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-soft); font-weight: 600; }
.vivo-capital-chart-supplier { color: var(--amber); }
.vivo-capital-chart-summary { display: flex; align-items: baseline; gap: 10px; }
.vivo-capital-chart-delta { font-size: 12px; font-family: 'JetBrains Mono', monospace; }
.vivo-capital-chart-svg { width: 100%; height: 72px; display: block; }
.vivo-chart-tooltip { position: absolute; background: var(--sidebar-bg); color: #fbf8f0; border-radius: 7px; padding: 7px 10px; font-size: 12px; pointer-events: none; white-space: nowrap; z-index: 10; box-shadow: 0 2px 8px rgba(0,0,0,0.18); }
.vivo-chart-tooltip-date { font-size: 10.5px; color: #a39c8a; margin-bottom: 2px; }
.vivo-chart-tooltip-value { font-weight: 600; font-size: 13px; font-family: 'JetBrains Mono', monospace; }
.vivo-chart-tooltip-delta { font-size: 11px; margin-top: 2px; font-family: 'JetBrains Mono', monospace; }
.vivo-stockdays { color: var(--rust); font-weight: 600; }
.vivo-select-tag { font-size: 12px; padding: 5px 7px; min-width: 132px; border: 1px solid var(--line); border-radius: 6px; background: var(--card); color: var(--ink); font-family: inherit; cursor: pointer; }
.vivo-table-quality td { vertical-align: middle; }
.vivo-table-compact th { padding: 5px 7px; font-size: 9.5px; }
.vivo-table-compact td { padding: 4px 7px; font-size: 11.5px; }
.vivo-table-compact .vivo-select-tag { min-width: 100px; padding: 3px 5px; font-size: 11px; }
.vivo-table-compact .vivo-item-cell { max-width: 180px; font-size: 11.5px; }
.vivo-table-wrap-sticky { max-height: 70vh; overflow: auto; position: relative; }
.vivo-table-wrap-sticky table thead th { position: sticky; top: 0; background: var(--card); z-index: 2; box-shadow: 0 1px 0 var(--line); }
.vivo-capital-danger { color: var(--rust); font-weight: 600; }
.vivo-top-scrollbar { overflow-x: auto; overflow-y: hidden; height: 14px; margin-bottom: 2px; }
.vivo-top-scrollbar::-webkit-scrollbar { height: 10px; }
.vivo-top-scrollbar::-webkit-scrollbar-thumb { background: #d8cdb0; border-radius: 100px; }
.vivo-top-scrollbar::-webkit-scrollbar-track { background: #f0ebdd; border-radius: 100px; }
.vivo-tag-blue { background: #e7eef6; border-color: #b9cee4; color: #2f5277; }
.vivo-tag-teal { background: #e3eee9; border-color: #aed4c4; color: #2c6650; }
.vivo-tag-amber { background: var(--amber-soft); border-color: #e6cda4; color: #6b4419; }
.vivo-tag-rust { background: var(--rust-soft); border-color: #e3b4a3; color: #7a2c19; }
.vivo-tag-olive { background: #eef1e9; border-color: #c7d4b8; color: var(--olive-dark); }
.vivo-tag-neutral { background: #f0ebdd; border-color: var(--line); color: var(--ink-soft); }
.vivo-toggle-btn { font-family: inherit; font-size: 11.5px; font-weight: 600; padding: 5px 14px; border-radius: 100px; border: 1px solid var(--line); background: var(--card); color: var(--ink-soft); cursor: pointer; transition: background 0.15s; white-space: nowrap; }
.vivo-toggle-btn:hover { background: #f1ece0; }
.vivo-toggle-btn.is-on { background: var(--olive-dark); border-color: var(--olive-dark); color: #fff; }
.vivo-toggle-btn.is-on:hover { background: var(--olive); }
.vivo-toggle-btn-off { color: var(--rust); border-color: #e3b4a3; }
.vivo-toggle-btn-off:hover { background: var(--rust-soft); }
.vivo-notes-input { width: 180px; font-family: inherit; font-size: 12px; padding: 5px 8px; border: 1px solid var(--line); border-radius: 6px; background: var(--card); color: var(--ink); }
.vivo-notes-input:focus { outline: none; border-color: var(--amber); }
.vivo-notes-input::placeholder { color: #b5ad9a; }
.vivo-monitor-badge { display: inline-flex; align-items: center; gap: 5px; font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 100px; white-space: nowrap; }
.vivo-monitor-up { background: #e1ecd9; color: #3c6b2a; }
.vivo-monitor-down { background: var(--rust-soft); color: var(--rust); }
.vivo-monitor-flat { background: #f0ebdd; color: var(--ink-soft); }
.vivo-monitor-muted { background: #f0ebdd; color: var(--ink-soft); font-style: italic; font-size: 11.5px; font-weight: 400; padding: 4px 10px; border-radius: 100px; }
.vivo-table-more { padding: 10px 12px; font-size: 12px; color: var(--ink-soft); }
.vivo-table-empty { text-align: center; color: var(--ink-soft); padding: 30px; }
.vivo-icon-btn { background: transparent; border: none; color: var(--ink-soft); cursor: pointer; display: flex; align-items: center; padding: 2px; border-radius: 4px; }
.vivo-icon-btn:hover { color: var(--ink); background: #f0ebdd; }
.vivo-icon-danger:hover { color: var(--rust); background: var(--rust-soft); }
.vivo-detail-label { display: block; font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-soft); margin-bottom: 3px; }
.vivo-detail-value { font-size: 14px; font-weight: 500; color: var(--ink); }
.vivo-mini-table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 8px; }
.vivo-mini-table th { text-align: left; color: var(--ink-soft); font-weight: 500; padding: 5px 8px; border-bottom: 1px solid var(--line); }
.vivo-mini-table th.num, .vivo-mini-table td.num { text-align: right; }
.vivo-mini-table td { padding: 5px 8px; font-family: 'JetBrains Mono', monospace; border-bottom: 1px solid #f0ebdd; }
.vivo-toolbar { display: flex; align-items: center; gap: 10px; margin-bottom: 16px; flex-wrap: wrap; }
.vivo-search { display: flex; align-items: center; gap: 7px; background: var(--card); border: 1px solid var(--line); border-radius: 7px; padding: 7px 11px; color: var(--ink-soft); min-width: 220px; }
.vivo-search input { border: none; outline: none; background: transparent; font-size: 13px; flex: 1; color: var(--ink); font-family: inherit; }
.vivo-select { border: 1px solid var(--line); background: var(--card); border-radius: 7px; padding: 7px 10px; font-size: 13px; color: var(--ink); font-family: inherit; }
.vivo-checkbox { display: inline-flex; align-items: center; gap: 6px; font-size: 13px; color: var(--ink-soft); cursor: pointer; }
.vivo-input, .vivo-input-mini { border: 1px solid var(--line); border-radius: 6px; padding: 6px 9px; font-size: 13px; font-family: 'JetBrains Mono', monospace; color: var(--ink); background: var(--card); }
.vivo-input { display: block; margin-top: 6px; width: 140px; }
.vivo-input-mini { width: 68px; padding: 5px 7px; }
.vivo-settings-hint { color: var(--ink-soft); font-size: 12px; margin: 10px 0 0; }
.vivo-source-tag { font-size: 9px; color: var(--olive-dark); margin-left: 3px; font-weight: 600; text-transform: uppercase; }
.vivo-action-row { display: flex; align-items: center; gap: 10px; margin-bottom: 14px; flex-wrap: wrap; }
.vivo-action-spacer { flex: 1; }
.vivo-action-summary { font-size: 12.5px; color: var(--ink-soft); font-family: 'JetBrains Mono', monospace; }
.vivo-action-warning { color: var(--rust); }
.vivo-no-price { color: var(--rust); font-family: inherit; }
.vivo-empty { display: flex; flex-direction: column; align-items: center; gap: 8px; text-align: center; padding: 70px 20px; color: var(--ink-soft); max-width: 360px; margin: 0 auto; }
.vivo-empty h3 { font-family: 'Fraunces', serif; font-size: 18px; color: var(--ink); margin: 6px 0 0; font-weight: 600; }
.vivo-empty p { font-size: 13px; margin: 0; }
@media (max-width: 760px) {
  .vivo-topnav-row { padding: 0 14px; overflow-x: auto; }
  .vivo-topnav-hint { display: none; }
  .vivo-page { padding: 20px 16px 40px; }
}
`;
