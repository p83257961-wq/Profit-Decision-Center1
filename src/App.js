import React, {
  useState,
  useMemo,
  useEffect,
  useRef,
  useCallback,
} from "react";
import {
  Trash2,
  BarChart3,
  Search,
  Package,
  FileSpreadsheet,
  TrendingUp,
  CheckCircle2,
  AlertTriangle,
  Settings,
  Download,
  UploadCloud,
  Wallet,
  CreditCard,
  AlertCircle,
  PieChart,
  RotateCcw,
  Lock,
  Loader2,
  Gift,
  Zap,
  Layers,
  Sun,
  Moon,
  ChevronDown,
  ChevronUp,
  Users,
  X,
  Info,
} from "lucide-react";
import {
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip as RTooltip,
  ResponsiveContainer,
  ComposedChart,
  Bar,
  Line,
  /* lucide 也有 PieChart（圖示），recharts 的另取名 */
  PieChart as RPieChart,
  Pie,
  Cell,
} from "recharts";
import { initializeApp, getApps, getApp } from "firebase/app";
import { getAuth, onAuthStateChanged, signInAnonymously } from "firebase/auth";
import {
  getFirestore,
  doc,
  collection,
  onSnapshot,
  setDoc,
  getDoc,
  deleteDoc,
  serverTimestamp,
} from "firebase/firestore";

/* ─── Firebase Config ─────────────────────────────────────────── */
const FBC = {
  apiKey: "AIzaSyBGWCKe1mw87g_KRtt-Ar3ffeAExOoYJrg",
  authDomain: "enterprise-e7704.firebaseapp.com",
  projectId: "enterprise-e7704",
  storageBucket: "enterprise-e7704.firebasestorage.app",
  messagingSenderId: "435446255525",
  appId: "1:435446255525:web:7f1727440a507d7add224c",
};
const FSPC = { collection: "warrooms", docId: "unified_profit_center_v1" };
/* 舊版單一 doc 的訂單來源：只給 runMigrationIfNeeded 與 forceLegacyRestore（救援用）讀 */
const FSPC_SL_ORD = { collection: "warrooms", docId: "unified_sl_orders_v1" };
const FSPC_SP_ORD = { collection: "warrooms", docId: "unified_sp_orders_v1" };

/* 新的按月拆分 collection */
const SL_MONTHLY_COLL = "sl_orders_monthly";
const SP_MONTHLY_COLL = "sp_orders_monthly";
const POS_MONTHLY_COLL = "pos_orders_monthly";

/* ─── 門市 POS：通路分類（吃交易明細「付款方式」欄） ─────────── */
const POS_CHANNELS = [
  { key: "retail", label: "現場零售" },
  { key: "dealer", label: "經銷·老客價" },
  { key: "phone", label: "電話訂購" },
  { key: "omnichat", label: "Omnichat" },
  { key: "partner", label: "合作通路" },
  { key: "corp", label: "企業採購" },
];
const posChannelOf = (payMethod) => {
  const p = String(payMethod || "");
  if (p.includes("經銷") || p.includes("老客")) return "dealer";
  if (p.includes("電話")) return "phone";
  if (/omnichat/i.test(p)) return "omnichat";
  if (p.includes("合作")) return "partner";
  if (p.includes("企業")) return "corp";
  return "retail";
};
const posChannelLabel = (key) =>
  POS_CHANNELS.find((c) => c.key === key)?.label || "現場零售";
/* 測試單門檻：金額 ≤ 此值視為測試（LINE Pay 1 元那類），預設排除計算 */
const POS_TEST_MAX = 10;
/* 門市發票判定（老闆 2026-08-18）：POS 只對電子支付自動帶發票號碼；企業客戶常走 SHOPLINE
   線上「新增獨立發票」自訂品名，POS 匯出裡沒有號碼——所以再看備註／統編／通路：
   ① 有發票號碼 ② 交易明細填了統編 ③ 備註寫到「發票／公司戶／統編／三聯／二聯」
   ④ 通路＝企業採購 或 合作通路（老闆 2026-08-18：這兩個通路都一定開發票）。
   備註若寫「不開／免開／不用發票」則一律視為未開。
   回傳 { has, src }，src 供畫面標示來源；另有逐單手動覆寫（invoiceOverride） */
/* 否定只認「發票」本身：「無統編／不用統編／不需三聯」在台灣零售通常是「開二聯、不打統編」，不是不開發票 */
const POS_INVOICE_NEG = /(不|免|無|沒|未)\s*(用|需|要|開|開立|需要)?\s*發票|發票\s*(不用|免|不開|不需)/;
const POS_INVOICE_POS = /發票|公司戶|統編|三聯|二聯|invoice/i;
const posInvoiceOf = ({ invoiceNo, taxId, remark, channel }) => {
  const r = String(remark || "");
  if (POS_INVOICE_NEG.test(r)) return { has: false, src: "備註註明不開" };
  if (invoiceNo) return { has: true, src: `號碼 ${invoiceNo}` };
  if (taxId) return { has: true, src: `統編 ${taxId}` };
  if (POS_INVOICE_POS.test(r)) return { has: true, src: "備註" };
  if (channel === "corp") return { has: true, src: "企業採購通路" };
  if (channel === "partner") return { has: true, src: "合作通路" };
  return { has: false, src: "" };
};

/* 依 YYYY-MM 把訂單分組 */
const groupOrdersByMonth = (orders) => {
  const byMonth = {};
  Object.entries(orders || {}).forEach(([id, o]) => {
    const ym = String(o?.date || "").substring(0, 7) || "unknown";
    if (!byMonth[ym]) byMonth[ym] = {};
    byMonth[ym][id] = o;
  });
  return byMonth;
};

/* ─── Constants ────────────────────────────────────────────────── */
const SL_PAYMENT_RATES = {
  信用卡: { rate: 0.022, flat: 0 },
  "LINE Pay": { rate: 0.023, flat: 0 },
  "7-11": { rate: 0, flat: 0 },
  全家: { rate: 0, flat: 0 },
  "宅配（貨到付款）": { rate: 0.01, flat: 0 },
  ApplePay: { rate: 0.022, flat: 0 },
  "Apple Pay": { rate: 0.022, flat: 0 },
  銀行轉帳: { rate: 0.01, flat: 0 },
  ATM: { rate: 0.01, flat: 0 },
  PayPal: { rate: 0.044, flat: 10 },
  WeChat: { rate: 0.0275, flat: 0 },
  /* 官網報表裡的 POS 單（付款方式「實體商店」）：零金流費 */
  實體商店: { rate: 0, flat: 0 },
};
const SL_SHIPPING_RATES = {
  "7-11": 65,
  全家: 65,
  宅配: 120,
  順豐: 250,
  SF: 250,
  /* 門市取貨／自取：零運費（要排在宅配之前無所謂，名稱不重疊） */
  實體商店: 0,
  自取: 0,
};
/* 國際單：運費收入已含在付款總金額，成本＝該筆運費收入（比對不分大小寫）。
   SHOPLINE 實際名稱：「國際快捷（Express Mail Service）」「FedEx 國際快遞（美國）」 */
const SL_INTL_METHODS = [
  "EMS",
  "EXPRESS MAIL",
  "FEDEX",
  "DHL",
  "UPS",
  "國際",
  "海外",
  "中國",
  "新加坡",
  "國外",
];
/* 送貨方式比不到對照表時的後備運費——只在真的比不到時用，並在明細標 ⚠ */
const SL_SHIP_FALLBACK = 120;
const slIntl = (dlv) => {
  const d = String(dlv || "").toUpperCase();
  return SL_INTL_METHODS.some((k) => d.includes(k.toUpperCase()));
};
const slShipRate = (dlv) => {
  const d = String(dlv || "").toUpperCase();
  for (const [k, v] of Object.entries(SL_SHIPPING_RATES))
    if (d.includes(k.toUpperCase())) return v;
  return null;
};
const slPayRate = (pay) => {
  const p = String(pay || "");
  for (const [k, v] of Object.entries(SL_PAYMENT_RATES)) if (p.includes(k)) return v;
  return null;
};

/* KPI 2026-08-25 老闆拍板（用儀表板現行口徑計價）：官網 15、蝦皮 10、門市統一 12、
   總覽固定目標 12（綠 ≥12／黃 10–12 成長油門帶／紅 <10）。營業費為老闆自調值勿改。
   2026-09-03 補述：10% 是地板不是目標，10–12 那 2 個百分點（≈84 萬/年）刻意留給
   加碼投放；煞車＝邊際 ROAS <3、或淨利真的碰到 10%。9 月起執行。 */
const DEFAULT_FP_SL = {
  platformFeeRate: "1.0",
  opExpense: "30.0",
  tax: "6.2",
  targetNet: "15.0",
  posTargetNet: "12.0",
};
const DEFAULT_FP_SP = { opExpense: "30.0", tax: "6.2", targetNet: "10.0" };

const SK = {
  platform: "upc_platform_v1",
  slFp: "upc_sl_fee_params_v1",
  spFp: "upc_sp_fee_params_v1",
  slCosts: "upc_sl_costs_v1",
  spCosts: "upc_sp_costs_v1",
  slOrders: "upc_sl_orders_v1",
  spOrders: "upc_sp_orders_v1",
  commissions: "upc_commissions_v1",
  components: "upc_components_v1",
  slRecipes: "upc_sl_recipes_v1",
  spRecipes: "upc_sp_recipes_v1",
  posOrders: "upc_pos_orders_v1",
  posCosts: "upc_pos_costs_v1",
  posRecipes: "upc_pos_recipes_v1",
  posRatios: "upc_pos_ratios_v1",
  posIncluded: "upc_pos_included_v1",
  theme: "upc_theme_v1",
};

/* ─── Utility Functions ────────────────────────────────────────── */
const fmt$ = (v) =>
  new Intl.NumberFormat("zh-TW", {
    style: "currency",
    currency: "TWD",
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(Number(v || 0));
const fmtP = (v) => `${((Number(v) || 0) * 100).toFixed(2)}%`;
/* 帶小數的金額（原料庫/配方用）：最多 2 位、整數不補零 */
const fmt$d = (v) =>
  new Intl.NumberFormat("zh-TW", {
    style: "currency",
    currency: "TWD",
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  }).format(Number(v || 0));
const numOrZero = (v) => {
  const n = parseFloat(String(v ?? "").replace(/,/g, ""));
  return Number.isFinite(n) ? n : 0;
};
const safeText = (v) => String(v ?? "").trim();
/* 今天所在的年／月（畫面預設期間用）。老闆 2026-09-03：進來就要停在當月，
   沒資料就顯示空白等匯入，不要自動退回上個月或整年 */
const nowYM = () => {
  const d = new Date();
  return { y: String(d.getFullYear()), m: String(d.getMonth() + 1).padStart(2, "0") };
};
/* 欄名比對：忽略所有空白字元。平台偶爾會把「退貨 / 退款狀態」改成「退貨/退款狀態」，
   精確比對會靜默失效（整欄變 0／空字串）而且不會有任何警告 */
const nzKey = (s) => safeText(s).replace(/\s+/g, "");
const nzIndex = (hdrs, name) => {
  const t = nzKey(name);
  return (hdrs || []).findIndex((h) => nzKey(h) === t);
};
/* 日期正規化：容忍 2026/1/5、2026-1-5、含時間字串，一律轉成 YYYY-MM-DD */
const normDate = (raw) => {
  const s = safeText(raw).split(" ")[0].split("T")[0].replace(/\//g, "-");
  const p = s.split("-");
  if (p.length === 3 && p[0].length === 4) {
    const d = `${p[0]}-${p[1].padStart(2, "0")}-${p[2].padStart(2, "0")}`;
    if (/^\d{4}-\d{2}-\d{2}$/.test(d)) return d;
  }
  /* 無法辨識的內容一律進 1970 保底桶（會出現在年份選單，異常看得見），
     不原樣存入以免訂單被所有年月篩選靜默吞掉 */
  return "1970-01-01";
};
const jp = (s, f) => {
  try {
    return JSON.parse(s);
  } catch {
    return f;
  }
};
const gl = (k, f) => {
  try {
    const r = window.localStorage.getItem(k);
    return r ? jp(r, f) : f;
  } catch {
    return f;
  }
};
const sl_s = (k, v) => {
  try {
    window.localStorage.setItem(k, JSON.stringify(v));
    return true;
  } catch {
    return false;
  }
};
const gcid = () => {
  const K = "upc_client_id_v1";
  /* localStorage 被封鎖（私密模式／企業政策）時不能讓整頁在 mount 就炸，退回記憶體 id */
  try {
    const e = window.localStorage.getItem(K);
    if (e) return e;
  } catch {}
  const n = `c_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  try {
    window.localStorage.setItem(K, n);
  } catch {}
  return n;
};
/* 鍵序無關的序列化：用來比對「內容有沒有變」（雲端 echo／兩台互寫時不被鍵序或時間戳騙） */
const metaCoreOf = (m) => {
  if (!m || typeof m !== "object") return null;
  const rest = { ...m };
  delete rest.updatedAtMs;
  delete rest.updatedBy;
  return stableStringify(deepClean(rest));
};
const stableStringify = (v) =>
  JSON.stringify(v, (k, val) =>
    val && typeof val === "object" && !Array.isArray(val)
      ? Object.keys(val)
          .sort()
          .reduce((a, key) => {
            a[key] = val[key];
            return a;
          }, {})
      : val
  );
const deepClean = (o) => {
  if (Array.isArray(o)) return o.map(deepClean).filter((v) => v !== undefined);
  if (o && typeof o === "object") {
    const c = {};
    Object.entries(o).forEach(([k, v]) => {
      if (v !== undefined) c[k] = deepClean(v);
    });
    return c;
  }
  return o === undefined ? undefined : o;
};
const parseCSV = (text) => {
  let p = "",
    row = [""],
    ret = [row],
    i = 0,
    r = 0,
    s = true,
    l;
  for (l of text) {
    if ('"' === l) {
      if (s && l === p) row[i] += l;
      s = !s;
    } else if ("," === l && s) {
      l = row[++i] = "";
    } else if ("\n" === l && s) {
      if ("\r" === p) row[i] = row[i].slice(0, -1);
      row = ret[++r] = [(l = "")];
      i = 0;
    } else {
      row[i] += l;
    }
    p = l;
  }
  if (row[i] === "") row.pop();
  if (ret[ret.length - 1].length <= 1 && ret[ret.length - 1][0] === "")
    ret.pop();
  return ret;
};
const commKey = (yr, mo) =>
  yr === "All" ? "All" : mo === "All" ? yr : `${yr}-${mo}`;

/* 月份是否與自訂區間重疊（自訂區間的月費用以整月計） */
const ymOverlaps = (ym, range) => {
  if (!range) return true;
  const start = `${ym}-01`,
    end = `${ym}-31`;
  if (range.from && end < range.from) return false;
  if (range.to && start > range.to) return false;
  return true;
};
/* 期間費用加總：分潤等按月費用用（key 為 YYYY-MM） */
const periodExpense = (map, y, m, range) => {
  const val = (v) =>
    v !== "" && v !== undefined && v !== null ? Number(v) || 0 : 0;
  if (y === "Custom")
    return Object.entries(map || {}).reduce(
      (s, [k, v]) => (ymOverlaps(k, range) ? s + val(v) : s),
      0
    );
  if (y === "All")
    return Object.values(map || {}).reduce((s, v) => s + val(v), 0);
  if (m === "All")
    return Object.entries(map || {}).reduce(
      (s, [k, v]) => (k.startsWith(y + "-") ? s + val(v) : s),
      0
    );
  return val((map || {})[commKey(y, m)]);
};
/* 數字或 null（NaN 不寫入快照；?? 攔不住 NaN） */
const numOrNull = (v) => (Number.isFinite(Number(v)) ? Number(v) : null);
/* 重新匯入同一張訂單時保留舊快照：參數整組沿用、成本按品項 key 比對。
   新品項若在舊快照中沒有對應成本，該單會呈現「部分鎖定」提醒使用者重鎖 */
const withOldSnapshot = (oldOrder, next) => {
  if (!oldOrder?.snapshotFeeParams) return next;
  const oldCost = {};
  (oldOrder.items || []).forEach((i) => {
    if (Object.prototype.hasOwnProperty.call(i, "snapshotCost"))
      oldCost[i.key] = { cost: i.snapshotCost, est: i.snapshotEst === true };
  });
  return {
    ...next,
    snapshotFeeParams: oldOrder.snapshotFeeParams,
    items: (next.items || []).map((i) =>
      Object.prototype.hasOwnProperty.call(oldCost, i.key)
        ? { ...i, snapshotCost: oldCost[i.key].cost, snapshotEst: oldCost[i.key].est }
        : i
    ),
  };
};
/* ─── 成本組件（配方制）───────────────────────────────────────
   原料庫 components: { compId: { name, price } }（兩平台共用單價）
   配方 recipes: { costKey: [{ compId, qty }] }（各平台各自掛）
   商品有效成本 = 配方存在 ? Σ(組件單價×用量) : 手填成本。
   改組件單價 → 所有掛配方的商品自動重算；已鎖定月份因快照凍結不受影響 */
const recipeCost = (lines, components) =>
  (lines || []).reduce(
    (s, l) =>
      s + (Number(components?.[l.compId]?.price) || 0) * (Number(l.qty) || 0),
    0
  );
const resolveCosts = (costs, recipes, components) => {
  const eff = { ...costs };
  Object.entries(recipes || {}).forEach(([k, lines]) => {
    if (Array.isArray(lines) && lines.length)
      eff[k] = recipeCost(lines, components);
  });
  return eff;
};
const newCompId = () =>
  `cmp_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`;
/* 全歷史用途索引：商品名對照（找回無本期銷售的成本條目名稱）＋最後售出日（清理提醒） */
const buildUsage = (orders) => {
  const nameMap = {};
  const lastSold = {};
  let maxDate = "";
  Object.values(orders || {}).forEach((o) => {
    const d = String(o.date || "");
    if (d > maxDate) maxDate = d;
    (o.items || []).forEach((i) => {
      nameMap[i.key] = { name: i.name, option: i.option };
      if (!lastSold[i.key] || d > lastSold[i.key]) lastSold[i.key] = d;
    });
  });
  return { nameMap, lastSold, maxDate };
};
/* 按月費用（分潤）更新：空值＝刪除該月 key */
const monthlyUpd = (setter, key, value) =>
  setter((prev) => {
    if (value === "" || value === null || value === undefined) {
      const n = { ...prev };
      delete n[key];
      return n;
    }
    return { ...prev, [key]: Number(value) };
  });
/* 輸入防抖：搜尋框用，避免每個按鍵觸發全表過濾 */
const useDebounced = (value, delay = 200) => {
  const [v, setV] = useState(value);
  useEffect(() => {
    const t = setTimeout(() => setV(value), delay);
    return () => clearTimeout(t);
  }, [value, delay]);
  return v;
};

/* 窄螢幕（手機）判定：只給「CSS 改不動」的東西用——圖表要吃 width/height 數字。
   桌機一律回 false，走的是和以前完全一樣的那條路徑（老闆 2026-09-06：電腦版不動）。
   版面壓縮一律寫在 CSS 的 @media(max-width:600px) 裡，不要在這裡長條件。 */
const useIsNarrow = (q = "(max-width:600px)") => {
  const [v, setV] = useState(
    () => typeof window !== "undefined" && !!window.matchMedia?.(q).matches
  );
  useEffect(() => {
    if (typeof window === "undefined" || !window.matchMedia) return;
    const m = window.matchMedia(q);
    const on = () => setV(m.matches);
    setV(m.matches);
    /* Safari 14 以前只有 addListener */
    if (m.addEventListener) m.addEventListener("change", on);
    else m.addListener(on);
    /* resize 是保險絲：手機轉向、或某些環境沒送出 change 事件時仍會校正
       （值相同時 React 自己會跳過重繪，不會多跑） */
    window.addEventListener("resize", on);
    return () => {
      if (m.removeEventListener) m.removeEventListener("change", on);
      else m.removeListener(on);
      window.removeEventListener("resize", on);
    };
  }, [q]);
  return v;
};

/* ─── CSS ─────────────────────────────────────────────────────── */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500;600;700;800&family=Noto+Sans+TC:wght@400;500;600;700;800;900&family=Inter:wght@400;500;600;700;800&display=swap');
[data-theme="light"]{
  --bg:#FFFFFF;--s1:#FFFFFF;--s2:#F8F8F6;--s3:#EAEAE6;--s4:#D8D8D2;
  --t1:#1A1A18;--t2:#5C5C54;--t3:#74746A;--t4:#96968C;
  --accent:#1A6B3C;--accent-dim:rgba(26,107,60,0.06);--accent-bdr:rgba(26,107,60,0.18);--accent-text:#1A6B3C;
  --up:#1A6B3C;--up-dim:rgba(26,107,60,0.06);--up-bdr:rgba(26,107,60,0.18);
  --dn:#C0392B;--dn-dim:rgba(192,57,43,0.05);--dn-bdr:rgba(192,57,43,0.18);
  --wn:#B7600A;--wn-dim:rgba(183,96,10,0.05);--wn-bdr:rgba(183,96,10,0.18);
  --blue:#2E6DA4;--purple:#7B5EA7;--orange:#D4820A;--gold:#8B6914;
  --header-bg:rgba(255,255,255,0.92);
  --bar-track:#EAEAE6;
  --row-loss:rgba(192,57,43,0.04);
  --sp-accent:#EE4D2D;--sp-accent-dim:rgba(238,77,45,0.06);--sp-accent-bdr:rgba(238,77,45,0.2);
  --pos-accent:#7B5EA7;--pos-accent-dim:rgba(123,94,167,0.06);--pos-accent-bdr:rgba(123,94,167,0.2);
}
[data-theme="dark"]{
  --bg:#080A0E;--s1:#0E1117;--s2:#151921;--s3:#1C212B;--s4:#282D38;
  --t1:#E8E6E1;--t2:#9DA0A8;--t3:#7E838E;--t4:#656A73;
  --accent:#2ECC71;--accent-dim:rgba(46,204,113,0.08);--accent-bdr:rgba(46,204,113,0.2);--accent-text:#2ECC71;
  --up:#2ECC71;--up-dim:rgba(46,204,113,0.08);--up-bdr:rgba(46,204,113,0.2);
  --dn:#E74C3C;--dn-dim:rgba(231,76,60,0.08);--dn-bdr:rgba(231,76,60,0.2);
  --wn:#E67E22;--wn-dim:rgba(230,126,34,0.08);--wn-bdr:rgba(230,126,34,0.2);
  --blue:#3498DB;--purple:#9B7FCA;--orange:#F0A030;--gold:#C9A84C;
  --header-bg:rgba(8,10,14,0.88);
  --bar-track:#1C212B;
  --row-loss:rgba(231,76,60,0.07);
  --sp-accent:#FF6533;--sp-accent-dim:rgba(255,101,51,0.08);--sp-accent-bdr:rgba(255,101,51,0.22);
  --pos-accent:#9B7FCA;--pos-accent-dim:rgba(155,127,202,0.08);--pos-accent-bdr:rgba(155,127,202,0.22);
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
::-webkit-scrollbar{width:4px;height:4px;}
::-webkit-scrollbar-track{background:transparent;}
::-webkit-scrollbar-thumb{background:var(--s4);border-radius:99px;}
@keyframes fadeUp{from{opacity:0;transform:translateY(16px);}to{opacity:1;transform:translateY(0);}}
@keyframes spin{to{transform:rotate(360deg);}}
@keyframes toastIn{from{opacity:0;transform:translateX(100%);}to{opacity:1;transform:translateX(0);}}
@keyframes toastOut{from{opacity:1;}to{opacity:0;transform:translateX(100%);}}
@keyframes dlgIn{from{opacity:0;transform:scale(.96);}to{opacity:1;transform:scale(1);}}
.spin{animation:spin 1s linear infinite;}
.f0{animation:fadeUp .42s cubic-bezier(.16,1,.3,1) both;}
.f1{animation:fadeUp .42s cubic-bezier(.16,1,.3,1) .06s both;}
.f2{animation:fadeUp .42s cubic-bezier(.16,1,.3,1) .12s both;}
.f3{animation:fadeUp .42s cubic-bezier(.16,1,.3,1) .18s both;}
.f4{animation:fadeUp .42s cubic-bezier(.16,1,.3,1) .24s both;}
.f5{animation:fadeUp .42s cubic-bezier(.16,1,.3,1) .30s both;}
.gm{display:grid;grid-template-columns:240px 1fr;gap:20px;align-items:start;}
.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;}
.g3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;}
@media(max-width:900px){.gm{grid-template-columns:1fr;}.side-sticky{position:static!important;}}
@media(max-width:1000px){.g4{grid-template-columns:repeat(2,1fr);}.g3{grid-template-columns:repeat(2,1fr);}}
@media(max-width:600px){.g4,.g3{grid-template-columns:1fr;}.app-header{position:static!important;}}
.gcmp{display:grid;grid-template-columns:1fr 1px 1fr 1px 1fr;border-top:1px solid var(--s3);padding-top:20px;}
.gcmp-c{padding:0 16px;min-width:0;}
.gcmp-c:first-child{padding-left:0;}
.gcmp-c:last-child{padding-right:0;}
.gcmp-div{background:var(--s3);}
@media(max-width:1000px){.gcmp{grid-template-columns:1fr;gap:20px;}.gcmp-div{display:none;}.gcmp-c{padding:0;}}
.hero-num{font-size:clamp(40px,8vw,72px);overflow-wrap:anywhere;}
.hero-num-md{font-size:clamp(36px,7vw,64px);overflow-wrap:anywhere;}
.hero-pct{font-size:clamp(30px,5.5vw,48px);}
.hero-pct-md{font-size:clamp(28px,5vw,44px);}
input,select,button{font-family:'Inter','Noto Sans TC',sans-serif;}
button{cursor:pointer;transition:all .12s;}
button:hover{filter:brightness(1.06);}
button:active{transform:scale(.97);}
tr{transition:background .1s;}
tr:hover td{background:var(--s2)!important;}
tr.rw:hover td{background:var(--wn-dim)!important;}
tr.rl:hover td{background:var(--row-loss)!important;}
.rw td{background:var(--wn-dim)!important;}
.rl td{background:var(--row-loss)!important;}
.iw{border-color:var(--wn)!important;}
.iok{border-color:var(--up-bdr)!important;}
@media(prefers-reduced-motion:reduce){*,*::before,*::after{animation-duration:.01ms!important;animation-iteration-count:1!important;transition-duration:.01ms!important;}}
/* ─── 行動版版面（≤600px）─────────────────────────────────────────
   老闆 2026-09-06：「手機不需要的可以省略，電腦版不動」。
   所有規則都關在這個 media query 裡，桌機（>600px）完全不受影響。
   五個原則：①匯入／重置手機用不到＝隱藏，參數側欄移到資料下面
   ②外框與卡片內距減半 ③大字降一級 ④圖表縮進容器（圓餅另由 narrow 判斷換小尺寸）
   ⑤寬表格隱藏次要欄。要恢復哪一項就刪掉對應那行。 */
@media(max-width:600px){
  .page-wrap{padding:12px 12px 48px!important;}
  .hd-wrap{padding:8px 12px!important;gap:6px!important;}
  .hd-sub{display:none!important;}
  /* 卡片內距 24→14。比對的是 inline style 字串，日後若把 padding 改成別的值，
     這幾行會靜靜失效（不會壞版面），那時同步改這裡的數字即可 */
  .page-wrap [style*="padding: 24px"]{padding:14px!important;}
  .page-wrap [style*="padding: 22px 24px"]{padding:14px!important;}
  .page-wrap [style*="padding: 32px 36px"]{padding:18px 14px!important;}
  .page-wrap [style*="padding: 28px 36px"]{padding:16px 14px!important;}
  .page-wrap [style*="padding: 20px 22px"]{padding:12px 14px!important;}
  .page-wrap [style*="padding: 18px 24px"]{padding:14px!important;}
  .hero-row{gap:14px!important;margin-bottom:14px!important;}
  .hero-num{font-size:clamp(30px,8.5vw,40px);}
  .hero-num-md{font-size:clamp(28px,8vw,34px);}
  .hero-pct{font-size:26px;}
  .hero-pct-md{font-size:24px;}
  /* 匯入報表、重置本期＝手機不做的事 */
  .imp-zone,.reset-row{display:none!important;}
  /* 參數側欄移到資料下面（手機一進來先看到數字，不是設定） */
  .side-col{order:2;margin-top:14px;}
  /* 圓餅置中；圖例改兩欄、金額省略（金額在下面各平台卡片都有） */
  .pie-box{margin:2px auto 0!important;align-self:center!important;}
  .ch-legend{display:grid!important;grid-template-columns:1fr 1fr;gap:5px 10px!important;}
  .lg-amt{display:none!important;}
  /* 趨勢圖圖例：允許換行，避免「淨利」被擠成一個字一行 */
  .trend-legend{flex-wrap:wrap!important;gap:4px 12px!important;}
  .trend-legend>span{white-space:nowrap;}
  /* 寬表格：手機隱藏次要欄並解除 minWidth，讓主要欄位擠進畫面 */
  .tb-ov th,.tb-ov td,.tb-ch th,.tb-ch td,.tb-ord th,.tb-ord td,.tb-ord-sl th,.tb-ord-sl td,.tb-ord-sp th,.tb-ord-sp td{padding-left:6px!important;padding-right:6px!important;}
  /* 六通路表只留 通路／營收／佔全公司／淨利率——不用右滑就看得到淨利率（09-06 手機實測補的） */
  .tb-ov{min-width:0!important;}
  .tb-ov th:nth-child(4),.tb-ov td:nth-child(4),.tb-ov th:nth-child(5),.tb-ov td:nth-child(5),.tb-ov th:nth-child(6),.tb-ov td:nth-child(6),.tb-ov th:nth-child(8),.tb-ov td:nth-child(8){display:none;}
  /* 通路拆解只留 勾選／通路／營收／佔比／淨利率 */
  .tb-ch{min-width:0!important;}
  .tb-ch th:nth-child(5),.tb-ch td:nth-child(5),.tb-ch th:nth-child(6),.tb-ch td:nth-child(6),.tb-ch th:nth-child(7),.tb-ch td:nth-child(7),.tb-ch th:nth-child(9),.tb-ch td:nth-child(9),.tb-ch th:nth-child(10),.tb-ch td:nth-child(10){display:none;}
  /* 成本資料庫只留 商品名稱／規格／單位成本／編輯 */
  .tb-cost{min-width:0!important;}
  .tb-cost th,.tb-cost td{padding-left:6px!important;padding-right:6px!important;}
  .tb-cost th:nth-child(3),.tb-cost td:nth-child(3),.tb-cost th:nth-child(4),.tb-cost td:nth-child(4),.tb-cost th:nth-child(5),.tb-cost td:nth-child(5){display:none;}
  .tb-ord{min-width:0!important;}
  .tb-ord th:nth-child(4),.tb-ord td:nth-child(4),.tb-ord th:nth-child(6),.tb-ord td:nth-child(6),.tb-ord th:nth-child(7),.tb-ord td:nth-child(7){display:none;}
  .tb-ord-sl{min-width:0!important;}
  .tb-ord-sl th:nth-child(3),.tb-ord-sl td:nth-child(3),.tb-ord-sl th:nth-child(4),.tb-ord-sl td:nth-child(4),.tb-ord-sl th:nth-child(5),.tb-ord-sl td:nth-child(5){display:none;}
  .tb-ord-sp{min-width:0!important;}
  .tb-ord-sp th:nth-child(4),.tb-ord-sp td:nth-child(4),.tb-ord-sp th:nth-child(5),.tb-ord-sp td:nth-child(5),.tb-ord-sp th:nth-child(6),.tb-ord-sp td:nth-child(6){display:none;}
}
`;

/* ─── Small UI Components ──────────────────────────────────────── */
const mono = "'JetBrains Mono',monospace";
const inp = {
  border: "1px solid var(--s3)",
  borderRadius: 6,
  padding: "6px 8px",
  fontSize: 13,
  fontWeight: 500,
  outline: "none",
  textAlign: "right",
  fontFamily: mono,
  background: "var(--s2)",
  color: "var(--t1)",
  transition: "border-color .15s",
};
const sel = {
  border: "1px solid var(--s3)",
  borderRadius: 8,
  padding: "6px 12px",
  fontSize: 12,
  fontWeight: 600,
  background: "var(--s1)",
  color: "var(--t1)",
  outline: "none",
  cursor: "pointer",
};
const th = {
  position: "sticky",
  top: 0,
  background: "var(--s2)",
  fontSize: 11,
  color: "var(--t3)",
  fontWeight: 700,
  padding: "10px 14px",
  borderBottom: "1px solid var(--s3)",
  zIndex: 1,
  userSelect: "none",
  cursor: "pointer",
};
const td2 = {
  padding: "10px 14px",
  borderBottom: "1px solid var(--s3)",
  fontSize: 13,
  verticalAlign: "middle",
  background: "var(--s1)",
};

const SyncDot = ({ status, last }) => {
  const m = {
    idle: { l: "離線", c: "var(--t3)" },
    connecting: { l: "連線中", c: "var(--wn)" },
    synced: { l: "已同步", c: "var(--up)" },
    pending: { l: "待同步", c: "var(--wn)" },
    saving: { l: "儲存中", c: "var(--orange)" },
    error: { l: "失敗", c: "var(--dn)" },
  };
  const v = m[status] || m.idle;
  const t = last
    ? new Date(last).toLocaleString("zh-TW", { hour12: false })
    : "—";
  return (
    <div
      role="status"
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        padding: "5px 12px",
        borderRadius: 99,
        fontSize: 11,
        fontWeight: 600,
        background: "var(--s2)",
        color: v.c,
        fontFamily: mono,
      }}
    >
      <div style={{ width: 6, height: 6, borderRadius: 99, background: v.c }} />
      {(status === "connecting" || status === "saving") && (
        <Loader2 size={11} className="spin" />
      )}
      <span>{v.l}</span>
      <span style={{ color: "var(--s4)" }}>·</span>
      <span style={{ color: "var(--t2)", fontSize: 10 }}>{t}</span>
    </div>
  );
};

const Tag = ({ children, v = "default", style: st = {}, onClick }) => {
  const vs = {
    default: { bg: "var(--s2)", c: "var(--t2)", bd: "var(--s3)" },
    ok: { bg: "var(--up-dim)", c: "var(--up)", bd: "var(--up-bdr)" },
    bad: { bg: "var(--dn-dim)", c: "var(--dn)", bd: "var(--dn-bdr)" },
    warn: { bg: "var(--wn-dim)", c: "var(--wn)", bd: "var(--wn-bdr)" },
  };
  const s = vs[v] || vs.default;
  /* 可點的 Tag（例如 Hero「未填成本 N」跳轉）要能用鍵盤到達，
     與檔內其他可點非按鈕元素同一套：role/tabIndex/onKeyDown */
  const clickable = typeof onClick === "function";
  return (
    <span
      onClick={onClick}
      role={clickable ? "button" : undefined}
      tabIndex={clickable ? 0 : undefined}
      onKeyDown={
        clickable
          ? (e) => {
              if (e.key === "Enter" || e.key === " ") {
                e.preventDefault();
                onClick(e);
              }
            }
          : undefined
      }
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 4,
        padding: "4px 10px",
        borderRadius: 6,
        fontSize: 11,
        fontWeight: 700,
        background: s.bg,
        color: s.c,
        border: `1px solid ${s.bd}`,
        ...st,
      }}
    >
      {children}
    </span>
  );
};

const Btn = ({ children, v = "default", style: st = {}, ...p }) => {
  const vs = {
    default: {
      background: "var(--s2)",
      color: "var(--t1)",
      border: "1px solid var(--s3)",
    },
    primary: {
      background: "var(--accent-dim)",
      color: "var(--accent-text)",
      border: "1px solid var(--accent-bdr)",
    },
    danger: {
      background: "var(--dn-dim)",
      color: "var(--dn)",
      border: "1px solid var(--dn-bdr)",
    },
    ghost: {
      background: "transparent",
      color: "var(--t3)",
      border: "1px solid transparent",
    },
    shopee: {
      background: "var(--sp-accent-dim)",
      color: "var(--sp-accent)",
      border: "1px solid var(--sp-accent-bdr)",
    },
  };
  const s = vs[v] || vs.default;
  return (
    <button
      {...p}
      style={{
        ...s,
        borderRadius: 8,
        padding: "7px 14px",
        fontSize: 11,
        fontWeight: 700,
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        whiteSpace: "nowrap",
        ...st,
      }}
    >
      {children}
    </button>
  );
};

const Lbl = ({ children }) => (
  <div
    style={{
      fontSize: 12,
      fontWeight: 600,
      color: "var(--t3)",
      marginBottom: 4,
    }}
  >
    {children}
  </div>
);

const SortTh = ({ children, sortKey, currentSort, onSort, align = "left" }) => {
  const isActive = currentSort.key === sortKey;
  const dir = isActive ? currentSort.dir : null;
  return (
    <th
      scope="col"
      tabIndex={0}
      aria-sort={
        isActive ? (dir === "asc" ? "ascending" : "descending") : "none"
      }
      onClick={() => onSort(sortKey)}
      onKeyDown={(e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          onSort(sortKey);
        }
      }}
      style={{ ...th, textAlign: align }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 4,
          justifyContent: align === "right" ? "flex-end" : "flex-start",
        }}
      >
        {children}
        {isActive ? (
          dir === "asc" ? (
            <ChevronUp size={10} />
          ) : (
            <ChevronDown size={10} />
          )
        ) : (
          <ChevronDown size={9} style={{ opacity: 0.3 }} />
        )}
      </div>
    </th>
  );
};

/* 成本輸入框：本地草稿、失焦才寫回，避免每個按鍵觸發全表重算與雲端寫入 */
const CostInput = React.memo(function CostInput({
  costKey,
  label,
  value,
  miss,
  onCommit,
}) {
  const norm = (v) =>
    v === undefined || v === null || v === "" || Number(v) === 0
      ? ""
      : String(v);
  const [draft, setDraft] = useState(() => norm(value));
  const focused = useRef(false);
  useEffect(() => {
    if (!focused.current) setDraft(norm(value));
    // eslint-disable-next-line
  }, [value]);
  return (
    <input
      type="number"
      value={draft}
      placeholder="—"
      aria-label={label}
      onFocus={() => {
        focused.current = true;
      }}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => {
        focused.current = false;
        const n = parseFloat(draft);
        onCommit(costKey, Number.isFinite(n) ? n : 0);
      }}
      onKeyDown={(e) => {
        if (e.key === "Enter") e.currentTarget.blur();
      }}
      className={miss ? "iw" : Number(value) > 0 ? "iok" : ""}
      style={{ ...inp, width: 80 }}
    />
  );
});

/* 財務參數輸入框：同樣失焦才提交，避免每個按鍵觸發全部訂單重算＋雲端寫入 */
const FpInput = React.memo(function FpInput({ field, label, value, onCommit }) {
  const [draft, setDraft] = useState(() => String(value ?? ""));
  const focused = useRef(false);
  useEffect(() => {
    if (!focused.current) setDraft(String(value ?? ""));
    // eslint-disable-next-line
  }, [value]);
  return (
    <input
      type="number"
      step="0.1"
      aria-label={label}
      value={draft}
      onFocus={() => {
        focused.current = true;
      }}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => {
        focused.current = false;
        onCommit(field, draft);
      }}
      onKeyDown={(e) => {
        if (e.key === "Enter") e.currentTarget.blur();
      }}
      style={{ ...inp, width: 60, fontSize: 12 }}
    />
  );
});

/* 配方組件選擇器：可搜尋、點選後保持焦點連續加入（取代長下拉） */
const CompPicker = React.memo(function CompPicker({ compGroups, onPick }) {
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const ql = q.trim().toLowerCase();
  const groups = compGroups
    .map(([cat, list]) => [
      cat,
      ql
        ? list.filter(
            ([, c]) =>
              c.name.toLowerCase().includes(ql) ||
              (c.cat || "").toLowerCase().includes(ql)
          )
        : list,
    ])
    .filter(([, l]) => l.length);
  const flat = groups.flatMap(([, l]) => l);
  const pick = (id) => {
    onPick(id);
    setQ("");
  };
  return (
    <div style={{ position: "relative", flex: "1 1 220px", maxWidth: 300 }}>
      <input
        value={q}
        placeholder="＋ 搜尋組件加入（Enter 加第一筆）…"
        aria-label="搜尋組件加入配方"
        onFocus={() => setOpen(true)}
        onBlur={() => setTimeout(() => setOpen(false), 120)}
        onChange={(e) => {
          setQ(e.target.value);
          setOpen(true);
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter" && flat.length) {
            e.preventDefault();
            pick(flat[0][0]);
          }
          if (e.key === "Escape") setOpen(false);
        }}
        style={{
          ...inp,
          width: "100%",
          textAlign: "left",
          fontSize: 12,
          padding: "7px 10px",
          fontFamily: "'Inter','Noto Sans TC',sans-serif",
        }}
      />
      {open && (
        <div
          style={{
            position: "absolute",
            top: "calc(100% + 4px)",
            left: 0,
            zIndex: 30,
            width: 320,
            maxHeight: 260,
            overflowY: "auto",
            background: "var(--s1)",
            border: "1px solid var(--s3)",
            borderRadius: 10,
            boxShadow: "0 12px 40px rgba(0,0,0,.18)",
            padding: 4,
          }}
        >
          {groups.length === 0 && (
            <div style={{ padding: 10, fontSize: 11, color: "var(--t3)" }}>
              找不到符合的組件
            </div>
          )}
          {groups.map(([cat, list]) => (
            <div key={cat}>
              <div
                style={{
                  fontSize: 9,
                  fontWeight: 800,
                  color: "var(--accent-text)",
                  padding: "6px 8px 2px",
                  letterSpacing: "0.05em",
                }}
              >
                {cat}
              </div>
              {list.map(([id, c]) => (
                <div
                  key={id}
                  role="button"
                  tabIndex={0}
                  onMouseDown={(e) => {
                    /* mousedown（blur 前）選取＋保留輸入框焦點，連續加入 */
                    e.preventDefault();
                    pick(id);
                  }}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") pick(id);
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.background = "var(--s2)")
                  }
                  onMouseLeave={(e) =>
                    (e.currentTarget.style.background = "transparent")
                  }
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                    gap: 8,
                    padding: "5px 8px",
                    borderRadius: 6,
                    cursor: "pointer",
                    fontSize: 12,
                  }}
                >
                  <span
                    style={{
                      color: "var(--t1)",
                      fontWeight: 600,
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {c.name}
                  </span>
                  <span
                    style={{
                      fontFamily: mono,
                      color: "var(--t3)",
                      fontSize: 11,
                      flexShrink: 0,
                    }}
                  >
                    {fmt$d(c.price)}
                  </span>
                </div>
              ))}
            </div>
          ))}
        </div>
      )}
    </div>
  );
});

/* ─── 門市 POS Dashboard ────────────────────────────────────── */
/* 視覺語彙與官網／蝦皮同一套：Hero（狀態列＋大字淨利＋淨利率＋Waterfall）→ KPI 卡 → 表格 */
const posCard = {
  background: "var(--s1)",
  border: "1px solid var(--s3)",
  borderRadius: 16,
  padding: 24,
};
/* 與官網 KPI 卡同規格：Lbl ＋ 30px mono 數字 ＋ 11px 說明 */
const PosKpi = ({ label, value, sub, color }) => (
  <div
    style={{
      background: "var(--s1)",
      border: "1px solid var(--s3)",
      borderRadius: 14,
      padding: "22px 24px",
    }}
  >
    <Lbl>{label}</Lbl>
    <div
      style={{
        fontSize: 30,
        fontWeight: 700,
        fontFamily: mono,
        letterSpacing: "-0.03em",
        color: color || "var(--t1)",
        marginTop: 6,
      }}
    >
      {value}
    </div>
    {sub && (
      <div style={{ fontSize: 11, color: "var(--t4)", marginTop: 8 }}>{sub}</div>
    )}
  </div>
);
/* 門市統一淨利目標（六通路一個數字，2026-08-25 老闆拍板 12%）：
   側欄可調（slFp.posTargetNet），此常數只是無存值時的後備 */
const POS_TARGET_DEFAULT = 0.12;
const posTargetOf = (fp) => (parseFloat(fp?.posTargetNet) || 12) / 100;
/* 門市三個指標的色彩門檻——Hero／KPI／通路拆解／總覽六通路表一律呼叫這三支，
   不各自寫死數字（2026-08-18 審查：三處門檻不一致導致同頁自相矛盾） */
const posNetColor = (nm, target = POS_TARGET_DEFAULT) =>
  nm >= target ? "var(--up)" : nm >= 0.05 ? "var(--wn)" : "var(--dn)";
const posGmColor = (gm) =>
  gm >= 0.45 ? "var(--up)" : gm >= 0.35 ? "var(--wn)" : "var(--dn)";
const posCovColor = (cov) =>
  cov >= 0.9 ? "var(--up)" : cov >= 0.6 ? "var(--wn)" : "var(--dn)";
/* 營業費率分級（總覽大磚與三平台卡共用）：0 在目標內／1 FY2026 可接受帶／
   2 活動月加碼帶／3 需檢討。先四捨五入到顯示精度（百分比兩位）再比，
   否則 35.00% 會因浮點誤差 0.35000000000000003 被判成超標 */
const opexBandOf = (r) => {
  const v = Math.round((Number(r) || 0) * 10000) / 10000;
  return v <= 0.3 ? 0 : v <= 0.33 ? 1 : v <= 0.35 ? 2 : 3;
};
/* CSV 匯出共用工具（門市頁與官網／蝦皮共用同一份，避免兩份複製貼上各自漂移）。
   公式注入防護：OWASP 觸發字元 = + - @ 與 tab/CR 開頭一律前綴單引號，
   但純數字與百分比（含負的 -3.21%）是我們自己算出來的值，不能被加引號變成髒值。
   換行判斷含 \r：備註裡的單獨 CR 沒加引號會把一列拆成兩列 */
const csvEscape = (v0) => {
  const v = String(v0 ?? "");
  const safe = /^-?\d+(\.\d+)?%?$/.test(v);
  const body = !safe && /^[=@+\-\t\r]/.test(v) ? "'" + v : v;
  return /[",\r\n]/.test(body) ? `"${body.replace(/"/g, '""')}"` : body;
};
const downloadCsv = (rows, filename) => {
  const csv = "﻿" + rows.map((r) => r.map(csvEscape).join(",")).join("\r\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(
    new Blob([csv], { type: "text/csv;charset=utf-8" })
  );
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
};
/* 門市通路彙總：門市頁與總覽各自 reduce 一份會分岔（要加欄位容易只加一邊），
   統一走這支；呼叫端自己挑要用的欄位 */
const POS_CH_SUM_KEYS = [
  "rev",
  "cost",
  "gp",
  "net",
  "op",
  "tax",
  "orders",
  "noCostRev",
  "noCostCount",
  "estRev",
  "estCost",
  "estCount",
  "invoiceRev",
  "coveredRev",
  "coveredGp",
  "coveredNet",
  "coveredOp",
  "coveredTax",
  "lossCount",
];
const sumPosChannels = (list) =>
  (list || []).reduce((a, c) => {
    POS_CH_SUM_KEYS.forEach((k) => {
      a[k] = (a[k] || 0) + (Number(c?.[k]) || 0);
    });
    return a;
  }, Object.fromEntries(POS_CH_SUM_KEYS.map((k) => [k, 0])));
/* 交接檔陷阱③：單價≤1 且數量>100 ＝ 疑似開單時單價／數量顛倒（總價打進數量欄） */
const posSwapSuspect = (items) =>
  (items || []).some((it) => (Number(it.price) || 0) <= 1 && (Number(it.qty) || 0) > 100);
function POSDashboard({
  data,
  included,
  onSetIncluded,
  accentColor,
  accentDim,
  accentBdr,
  opExpense,
  taxRate,
  posTarget,
  isLocked,
  snapParams,
  onToggleSnap,
  canLock,
  missN,
  onJumpMiss,
  costsEff,
  recipes,
  components,
  ratios,
  monthly,
  sY,
  sM,
  periodLabel,
  onToggleInvoice,
  onExported,
}) {
  const s = data.summary;
  const gapVal = s.netMargin - posTarget;
  const [q, setQ] = useState("");
  const dq = useDebounced(q);
  const [openId, setOpenId] = useState(null);
  const [lossOnly, setLossOnly] = useState(false);
  const [sort, setSort] = useState({ key: "date", dir: "desc" });
  const [page, setPage] = useState(0);
  const [pageSize, setPageSize] = useState(20);
  useEffect(() => {
    setPage(0);
    setOpenId(null);
  }, [dq, lossOnly, sort, sY, sM]);
  const allKeys = data.channels.map((c) => c.key);
  const excludedKeys = allKeys.filter((k) => !included.includes(k));
  const scopeLabel = data.channels
    .filter((c) => !c.excluded)
    .map((c) => c.label)
    .join("＋");
  const netColor = posNetColor(s.netMargin, posTarget);
  const covColor = posCovColor(s.costCoverage);
  const cth = {
    padding: "9px 10px",
    fontSize: 10,
    fontWeight: 700,
    color: "var(--t3)",
    textAlign: "right",
    borderBottom: "1px solid var(--s3)",
    whiteSpace: "nowrap",
  };
  const ctd = {
    padding: "9px 10px",
    fontSize: 12,
    fontFamily: mono,
    textAlign: "right",
    borderBottom: "1px solid var(--s3)",
    color: "var(--t2)",
    whiteSpace: "nowrap",
  };
  const ctdL = { ...ctd, textAlign: "left", fontFamily: "inherit" };
  /* 通路合計（只算勾選中的通路，與 KPI 同口徑） */
  const chTotal = sumPosChannels(data.channels.filter((c) => !c.excluded));
  /* 訂單表：跟著上方「通路分組」的勾選走（老闆 2026-09-03：勾哪個通路就只看那個
     通路的訂單）。以前是全通路都列、未計入的灰掉，但那樣表頭的虧損／缺成本筆數
     會跟 Hero 對不起來，逐筆看某個通路也要自己用眼睛挑 */
  const visibleOrders = data.orderList.filter((o) =>
    included.includes(o.channel)
  );
  /* 搜尋／只看虧損／排序／分頁（與官網蝦皮同一套操作） */
  const filtered = visibleOrders
    .filter((o) => !lossOnly || (!o.missCost && o.net < 0))
    .filter((o) => {
      if (!dq) return true;
      const t = dq.toLowerCase();
      return (
        String(o.orderId).toLowerCase().includes(t) ||
        String(o.remark || "").toLowerCase().includes(t) ||
        String(o.channelLabel || "").toLowerCase().includes(t) ||
        String(o.payMethod || "").toLowerCase().includes(t) ||
        (o.items || []).some((i) =>
          `${i.name} ${i.option || ""}`.toLowerCase().includes(t)
        )
      );
    })
    .sort((a, b) => {
      const m = sort.dir === "desc" ? -1 : 1;
      const k = sort.key;
      if (k === "date")
        return (
          m *
          (String(a.date).localeCompare(String(b.date)) ||
            String(a.orderId).localeCompare(String(b.orderId)))
        );
      /* 缺成本的單沒有毛利／淨利，排序時一律沉底 */
      const va = k === "revenue" ? a.revenue : a.missCost ? -Infinity : a[k === "cost" ? "oCost" : k];
      const vb = k === "revenue" ? b.revenue : b.missCost ? -Infinity : b[k === "cost" ? "oCost" : k];
      return m * ((va || 0) - (vb || 0));
    });
  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const curPage = Math.min(page, totalPages - 1);
  const paged = filtered.slice(curPage * pageSize, (curPage + 1) * pageSize);
  /* 表頭筆數與表身同口徑（都只算勾選中的通路），才不會兩個數字對不起來 */
  const lossAll = visibleOrders.filter((o) => !o.missCost && o.net < 0).length;
  const noCostAll = visibleOrders.filter((o) => o.missCost).length;
  /* 顯示用參數：已鎖定月份的淨利是用快照參數算的，KPI 卡副標不能印側欄現值，
     否則同一張 Hero 上面寫「快照 38.8%」下面寫「41.2%」 */
  const snapOne =
    isLocked && snapParams && !snapParams.mixed ? snapParams.list[0] : null;
  const shownOp = snapOne ? snapOne.opExpense : opExpense;
  const shownTax = snapOne ? snapOne.tax : taxRate;
  const shownFpLabel =
    isLocked && snapParams && snapParams.mixed
      ? "各月快照 %（見側欄）"
      : `營業費 ${shownOp}%・稅 ${shownTax}%`;
  const onSort = (k) =>
    setSort((p) => ({
      key: k,
      dir: p.key === k ? (p.dir === "desc" ? "asc" : "desc") : "desc",
    }));
  /* 匯出 CSV：彙總（全期間）＋明細（套用目前篩選），與官網同格式與公式注入防護 */
  const exportCsv = () => {
    const r0 = Math.round;
    const rows = [
      ["平台", "門市"],
      ["期間", periodLabel],
      ["計入通路", scopeLabel],
      ["有效訂單", s.valid],
      ["營收", r0(s.rev)],
      ["成本齊全訂單營收", r0(s.coveredRev)],
      ["商品成本（成本齊全單）", r0(s.coveredRev - s.coveredGp)],
      ["毛利（成本齊全單）", r0(s.coveredGp)],
      ["營業費（成本齊全單）", r0(s.coveredOp)],
      ["稅賦（成本齊全單・只課開票）", r0(s.coveredTax)],
      ["最終淨利（成本齊全單）", r0(s.coveredNet)],
      ["毛利率", (s.grossMargin * 100).toFixed(2) + "%"],
      ["淨利率", (s.netMargin * 100).toFixed(2) + "%"],
      ["成本覆蓋率", (s.costCoverage * 100).toFixed(2) + "%"],
      ["開票比例", (s.invoiceRate * 100).toFixed(2) + "%"],
      ["平均客單價", r0(s.aov)],
      [],
    ];
    const bits = [];
    if (lossOnly) bits.push("僅虧損單");
    if (dq) bits.push(`搜尋「${dq}」`);
    rows.push([
      "明細範圍",
      bits.length
        ? `${bits.join("、")}，共 ${filtered.length} 筆（上方彙總仍為全期間）`
        : `全期間共 ${filtered.length} 筆`,
    ]);
    rows.push([
      "日期",
      "單號",
      "通路",
      "付款方式",
      "備註",
      "狀態",
      "營收",
      "商品成本",
      "毛利",
      "營業費",
      "稅賦",
      "單筆淨利",
      "發票",
      "商品",
    ]);
    filtered.forEach((o) =>
      rows.push([
        o.date,
        o.orderId,
        o.channelLabel,
        o.payMethod,
        o.remark,
        o.status,
        r0(o.revenue),
        o.missCost ? "缺" : r0(o.oCost),
        o.missCost ? "" : r0(o.gp),
        r0(o.opx),
        r0(o.taxAmt),
        o.missCost ? "" : r0(o.net),
        o.hasInvoice ? "已開" : "",
        (o.items || [])
          .map((i) => `${i.name}${i.option ? `（${i.option}）` : ""}×${i.qty}`)
          .join("；"),
      ])
    );
    downloadCsv(rows, `門市損益報表_${periodLabel}.csv`);
    onExported?.();
  };
  return (
    <>
      {/* 本期零筆：分辨「沒匯報表」與「真的沒生意」，避免一排 0 被誤讀 */}
      {s.valid === 0 && (
        <div
          className="f0"
          style={{
            background: "var(--wn-dim)",
            border: "1px solid var(--wn-bdr)",
            borderRadius: 12,
            padding: "10px 14px",
            fontSize: 12,
            color: "var(--t2)",
            lineHeight: 1.6,
          }}
        >
          <b style={{ color: "var(--wn)" }}>本期沒有門市訂單。</b>{" "}
          {data.channels.length && !data.channels.some((c) => !c.excluded)
            ? "本期沒有現場零售單、其他通路又未勾選計入——到下方通路拆解勾選。"
            : s.rawTotal > 0
            ? "本期只有取消／測試單。"
            : "可能還沒匯入這個月份的兩份報表，或期間選錯了。"}
        </div>
      )}

      {/* ══ HERO：與官網／蝦皮同版型（狀態列 → 大字淨利＋淨利率 → Waterfall） ══ */}
      <div className="f1" style={{ ...posCard, padding: "32px 36px" }}>
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: 8,
            alignItems: "center",
            marginBottom: 16,
          }}
        >
          <Tag v={s.coveredRev > 0 ? (gapVal >= 0 ? "ok" : "bad") : "default"}>
            <Zap size={10} />{" "}
            {s.coveredRev > 0 ? (gapVal >= 0 ? "穩健" : "警告") : "無資料"}
          </Tag>
          {/* 不 disable：三平台一致由 toggleSnap 用 toast 說明為什麼不能鎖，
              disabled 按鈕無法聚焦，鍵盤／手機使用者會看不到任何原因 */}
          <Btn
            v={isLocked ? "danger" : "default"}
            onClick={onToggleSnap}
            title={canLock ? "" : "請先切到單一月份（各月營業費 % 不同）"}
          >
            <Lock size={11} /> {isLocked ? "解除快照" : "鎖定快照"}
          </Btn>
          {isLocked && snapParams ? (
            <Tag v="default">
              <Lock size={10} />{" "}
              {snapParams.mixed
                ? "快照參數各月不同（見側欄）"
                : `快照 營業費 ${snapParams.list[0].opExpense}%・稅 ${snapParams.list[0].tax}%`}
            </Tag>
          ) : snapParams ? (
            <Tag v="default">
              部分訂單帶快照（見側欄）・現值 營業費 {opExpense}%・稅 {taxRate}%
            </Tag>
          ) : (
            <Tag v="default">
              即時 營業費 {opExpense}%・稅 {taxRate}%（只課開票單）
            </Tag>
          )}
          <Tag v="default">計入 {scopeLabel || "—"}</Tag>
          {/* 已鎖定的期間成本已凍結，不再提醒補成本（老闆 2026-09-03：
              鎖定後還叫人回頭處理以前的沒必要） */}
          {missN > 0 && !isLocked && (
            <Tag v="warn" style={{ cursor: "pointer" }} onClick={onJumpMiss}>
              <AlertCircle size={10} /> 未填成本 {missN} 項
            </Tag>
          )}
          {s.noCostCount > 0 && (
            <Tag v="warn">
              {s.noCostCount} 筆算不出成本 · {fmt$(s.noCostRev)} 未計毛利
            </Tag>
          )}
          {s.lossCount > 0 && (
            <Tag v="bad">虧損 {s.lossCount} 筆</Tag>
          )}
          {s.estCount > 0 && (
            <Tag v="warn">
              含估算成本 {s.estCount} 筆 · 估 {fmt$(s.estCost)} ·{" "}
              {fmtP(s.estShare)} 營收用成本率估
            </Tag>
          )}
          {s.coveredRev > 0 && (
            <span
              style={{
                fontSize: 12,
                color: gapVal >= 0 ? "var(--t3)" : "var(--wn)",
                marginLeft: 8,
              }}
            >
              {gapVal >= 0
                ? `✓ 淨利率 ${fmtP(s.netMargin)}，高於門市目標 ${(gapVal * 100).toFixed(1)}%`
                : `⚠ 淨利率 ${fmtP(s.netMargin)}，距門市目標差 ${Math.abs(gapVal * 100).toFixed(1)}%`}
            </span>
          )}
        </div>
        <div
          style={{
            display: "flex",
            alignItems: "flex-end",
            justifyContent: "space-between",
            flexWrap: "wrap",
            gap: 24,
          }}
        >
          <div>
            <div
              style={{
                fontSize: 12,
                fontWeight: 600,
                color: "var(--t3)",
                marginBottom: 4,
                letterSpacing: "0.06em",
              }}
            >
              最終結算淨利 · NET PROFIT
            </div>
            <div
              className="hero-num"
              style={{
                lineHeight: 1,
                fontWeight: 700,
                letterSpacing: "-0.04em",
                fontFamily: mono,
                color: s.coveredNet >= 0 ? "var(--t1)" : "var(--dn)",
              }}
            >
              {fmt$(s.coveredNet)}
            </div>
            <div style={{ fontSize: 12, color: "var(--t3)", marginTop: 8 }}>
              營收：{fmt$(s.rev)}（{scopeLabel || "—"}） ｜ {s.valid} 筆 ｜ 客單{" "}
              {fmt$(s.aov)}
              {s.coveredRev > 0
                ? ` ｜ 單筆平均淨利 ${fmt$(
                    s.coveredNet / Math.max(1, s.valid - s.noCostCount)
                  )}`
                : ""}
              {s.cancelledTotal > 0 ? ` ｜ 取消：${fmt$(s.cancelledTotal)}` : ""}
            </div>
            <PeriodCompare monthly={monthly} sY={sY} sM={sM} />
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: "var(--t3)" }}>
              淨利率
            </div>
            <div
              className="hero-pct"
              style={{
                fontWeight: 700,
                fontFamily: mono,
                lineHeight: 1,
                color: s.coveredRev > 0 ? netColor : "var(--t4)",
              }}
            >
              {s.coveredRev > 0 ? fmtP(s.netMargin) : "—"}
            </div>
            <div style={{ fontSize: 11, color: "var(--t3)", marginTop: 4 }}>
              門市目標 {fmtP(posTarget)}　差距{" "}
              <span
                style={{ color: gapVal >= 0 ? "var(--up)" : "var(--dn)" }}
              >
                {s.coveredRev > 0
                  ? `${gapVal >= 0 ? "+" : ""}${(gapVal * 100).toFixed(1)}%`
                  : "—"}
              </span>
            </div>
          </div>
        </div>
        {/* Waterfall：毛利（成本齊全單）− 營業費 − 稅賦（只課開票單）= 淨利，逐項可驗算 */}
        <div
          style={{
            marginTop: 28,
            borderTop: "1px solid var(--s3)",
            paddingTop: 20,
          }}
        >
          <div
            style={{
              fontSize: 11,
              fontWeight: 700,
              color: "var(--t3)",
              marginBottom: 14,
              letterSpacing: "0.06em",
            }}
          >
            損益分解 · WATERFALL
            {s.noCostCount > 0 && (
              <span style={{ fontWeight: 500, color: "var(--t4)", marginLeft: 8 }}>
                （只含成本齊全的 {s.valid - s.noCostCount} 筆）
              </span>
            )}
          </div>
          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              alignItems: "flex-end",
              gap: 0,
            }}
          >
            {[
              { l: "毛利（無平台抽成）", v: s.coveredGp, c: "var(--t1)" },
              { l: "營業費", v: -s.coveredOp, c: "var(--dn)" },
              { l: "稅賦（開票單）", v: -s.coveredTax, c: "var(--dn)" },
              { l: "淨利", v: s.coveredNet, c: accentColor, bold: true },
            ].map((item, i, arr) => (
              <React.Fragment key={i}>
                <div
                  style={{
                    flex: "1 1 0",
                    minWidth: 90,
                    textAlign: "center",
                    padding: "0 8px",
                  }}
                >
                  <div
                    style={{
                      fontSize: 11,
                      color: "var(--t3)",
                      fontWeight: 600,
                      marginBottom: 4,
                    }}
                  >
                    {item.l}
                  </div>
                  <div
                    style={{
                      fontSize: 20,
                      fontWeight: item.bold ? 800 : 600,
                      fontFamily: mono,
                      color: item.c,
                      letterSpacing: "-0.02em",
                    }}
                  >
                    {fmt$(item.v)}
                  </div>
                </div>
                {i < arr.length - 1 && (
                  <div
                    style={{
                      color: "var(--s4)",
                      fontSize: 18,
                      padding: "0 2px",
                      alignSelf: "center",
                    }}
                  >
                    ›
                  </div>
                )}
              </React.Fragment>
            ))}
          </div>
        </div>
      </div>

      {/* KPI 三卡：％為主、金額為輔（老闆 2026-08-18：主要看毛利率／淨利率）
          成本覆蓋率降為毛利率卡的副標（<100% 才顯示、變色提醒），開票比例併進淨利率卡副標 */}
      <div className="g3 f2">
        <PosKpi
          label={`門市營收（${scopeLabel || "無通路"}）`}
          value={fmt$(s.rev)}
          sub={`${s.valid} 筆 ｜ 客單 ${fmt$(s.aov)}${
            s.cancelledTotal > 0 ? ` ｜ 取消 ${fmt$(s.cancelledTotal)}` : ""
          }`}
        />
        <PosKpi
          label="毛利率（無平台抽成）"
          value={s.coveredRev > 0 ? fmtP(s.grossMargin) : "—"}
          sub={
            <>
              毛利 {fmt$(s.coveredGp)}
              {s.rev > 0 && s.costCoverage < 0.999 ? (
                <>
                  {" ｜ 成本覆蓋 "}
                  <b style={{ color: covColor }}>{fmtP(s.costCoverage)}</b>
                  {`（${s.noCostCount} 筆算不出成本）`}
                </>
              ) : s.rev > 0 ? (
                " ｜ ✓ 成本全覆蓋"
              ) : null}
            </>
          }
          color={s.coveredRev <= 0 ? "var(--t4)" : posGmColor(s.grossMargin)}
        />
        <PosKpi
          label="稅後淨利率"
          value={s.coveredRev > 0 ? fmtP(s.netMargin) : "—"}
          sub={`淨利 ${fmt$(s.coveredNet)} ｜ ${shownFpLabel}（開票 ${
            s.rev > 0 ? fmtP(s.invoiceRate) : "—"
          } 才課）`}
          color={s.coveredRev > 0 ? netColor : "var(--t4)"}
        />
      </div>
      {/* 通路分組 */}
      <div className="f0" style={posCard}>
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            alignItems: "center",
            gap: 12,
            marginBottom: 14,
          }}
        >
          <div style={{ fontSize: 13, fontWeight: 700, color: "var(--t1)" }}>
            通路拆解
          </div>
          <div
            style={{
              marginLeft: "auto",
              display: "flex",
              gap: 6,
              alignItems: "center",
            }}
          >
            <span style={{ fontSize: 10, color: "var(--t4)" }}>
              勾選＝計入上方 KPI 與商品表（預設只算現場零售）
            </span>
            {excludedKeys.length > 0 && (
              <Btn onClick={() => onSetIncluded(allKeys)} style={{ fontSize: 10 }}>
                全部計入
              </Btn>
            )}
            {excludedKeys.length < allKeys.length - 1 && (
              <Btn
                onClick={() => onSetIncluded(["retail"])}
                style={{ fontSize: 10 }}
              >
                只看現場零售
              </Btn>
            )}
          </div>
        </div>
        <div style={{ overflowX: "auto" }}>
          {/* tb-ch：手機隱藏客單／成本覆蓋／開票三欄（見 CSS 行動版區塊） */}
          <table
            className="tb-ch"
            style={{ width: "100%", borderCollapse: "collapse", minWidth: 680 }}
          >
            <thead>
              <tr>
                <th style={{ ...cth, textAlign: "left", width: 30 }}></th>
                <th style={{ ...cth, textAlign: "left" }}>通路</th>
                <th style={cth}>營收</th>
                <th style={cth}>佔比</th>
                <th style={cth}>筆數</th>
                <th style={cth}>客單</th>
                <th style={cth}>毛利率</th>
                <th style={cth}>淨利率</th>
                <th style={cth}>成本覆蓋</th>
                <th style={cth}>開票</th>
              </tr>
            </thead>
            <tbody>
              {data.channels.map((c) => {
                const gm = c.coveredRev > 0 ? c.coveredGp / c.coveredRev : 0;
                const nm = c.coveredRev > 0 ? c.coveredNet / c.coveredRev : 0;
                const cov = c.rev > 0 ? (c.rev - c.noCostRev) / c.rev : 0;
                const isDealer = c.key === "dealer";
                return (
                  <tr
                    key={c.key}
                    style={{
                      background: isDealer ? accentDim : "transparent",
                      opacity: c.excluded ? 0.4 : 1,
                    }}
                  >
                    <td style={{ ...ctd, textAlign: "left", padding: "9px 4px" }}>
                      <input
                        type="checkbox"
                        checked={!c.excluded}
                        aria-label={`把${c.label}計入 KPI`}
                        onChange={(e) =>
                          onSetIncluded(
                            e.target.checked
                              ? [...included, c.key]
                              : included.filter((k) => k !== c.key)
                          )
                        }
                        style={{ accentColor, cursor: "pointer" }}
                      />
                    </td>
                    <td style={{ ...ctdL, fontWeight: 700, color: "var(--t1)" }}>
                      {c.label}
                      {isDealer && (
                        <span
                          style={{
                            marginLeft: 6,
                            fontSize: 9,
                            padding: "1px 5px",
                            borderRadius: 4,
                            border: `1px solid ${accentBdr}`,
                            color: accentColor,
                          }}
                        >
                          老闆關係單
                        </span>
                      )}
                    </td>
                    <td style={ctd}>{fmt$(c.rev)}</td>
                    <td style={ctd}>
                      {chTotal.rev > 0 && !c.excluded
                        ? fmtP(c.rev / chTotal.rev)
                        : "—"}
                    </td>
                    <td style={ctd}>{c.orders}</td>
                    <td style={ctd}>
                      {c.orders > 0 ? fmt$(c.rev / c.orders) : "—"}
                    </td>
                    <td
                      style={{
                        ...ctd,
                        color: posGmColor(gm),
                        fontWeight: 700,
                      }}
                    >
                      {c.coveredRev > 0 ? fmtP(gm) : "—"}
                    </td>
                    <td
                      style={{
                        ...ctd,
                        color: posNetColor(nm, posTarget),
                        fontWeight: 700,
                      }}
                    >
                      {c.coveredRev > 0 ? fmtP(nm) : "—"}
                    </td>
                    <td
                      style={{
                        ...ctd,
                        color: posCovColor(cov),
                      }}
                    >
                      {c.rev > 0 ? fmtP(cov) : "—"}
                    </td>
                    <td style={ctd}>
                      {c.rev > 0 ? fmtP(c.invoiceRev / c.rev) : "—"}
                    </td>
                  </tr>
                );
              })}
              {data.channels.length > 1 && (
                <tr>
                  <td style={ctd}></td>
                  <td style={{ ...ctdL, fontWeight: 700, color: "var(--t3)" }}>
                    合計（勾選中）
                  </td>
                  <td style={{ ...ctd, fontWeight: 700 }}>
                    {fmt$(chTotal.rev)}
                  </td>
                  <td style={ctd}>100%</td>
                  <td style={ctd}>{chTotal.orders}</td>
                  <td style={ctd}>
                    {chTotal.orders > 0
                      ? fmt$(chTotal.rev / chTotal.orders)
                      : "—"}
                  </td>
                  <td style={{ ...ctd, fontWeight: 700 }}>
                    {chTotal.coveredRev > 0
                      ? fmtP(chTotal.coveredGp / chTotal.coveredRev)
                      : "—"}
                  </td>
                  <td style={{ ...ctd, fontWeight: 700 }}>
                    {chTotal.coveredRev > 0
                      ? fmtP(chTotal.coveredNet / chTotal.coveredRev)
                      : "—"}
                  </td>
                  <td style={ctd}>
                    {chTotal.rev > 0
                      ? fmtP(chTotal.coveredRev / chTotal.rev)
                      : "—"}
                  </td>
                  <td style={ctd}>
                    {chTotal.rev > 0
                      ? fmtP(chTotal.invoiceRev / chTotal.rev)
                      : "—"}
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        {(s.noCostRev > 0 || s.testCount > 0 || s.cancelledTotal > 0) && (
          <div
            style={{
              marginTop: 12,
              fontSize: 10,
              color: "var(--t4)",
              lineHeight: 1.7,
            }}
          >
            {s.noCostRev > 0 && (
              <div>
                ※ {fmt$(s.noCostRev)}{" "}
                營收的訂單算不出成本（品項未選正規如「茶葉」泛稱、或缺商品明細），已計入營收但排除在毛利計算外
              </div>
            )}
            {s.testCount > 0 && (
              <div>
                ※ 已排除 {s.testCount} 筆測試單（≤{POS_TEST_MAX} 元，共{" "}
                {fmt$(s.testTotal)}）
              </div>
            )}
            {s.cancelledTotal > 0 && (
              <div>※ 已排除取消／退款 {fmt$(s.cancelledTotal)}</div>
            )}
          </div>
        )}
      </div>

      {/* ── 單筆訂單決策明細（版型同官網／蝦皮：標題＋虧損數、匯出、搜尋、只看虧損、可排序、展開、分頁） ── */}
      <div className="f5" style={posCard}>
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            justifyContent: "space-between",
            gap: 12,
            alignItems: "center",
            marginBottom: 14,
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <BarChart3 size={16} color="var(--t3)" />
            <span style={{ fontSize: 14, fontWeight: 700 }}>單筆訂單決策明細</span>
            {/* 表身已跟著勾選過濾，筆數與 Hero 同口徑，不用再標「本表」 */}
            <span style={{ fontSize: 11, color: "var(--dn)" }}>
              虧損 {lossAll} 筆
            </span>
            {noCostAll > 0 && (
              <span style={{ fontSize: 11, color: "var(--wn)" }}>
                缺成本 {noCostAll} 筆
              </span>
            )}
            {excludedKeys.length > 0 && (
              <span style={{ fontSize: 10, color: "var(--t4)" }}>
                {/* 用「本期有單且已勾選」的通路數，與上方「計入 …」標籤同一個口徑；
                    勾了但本期沒單的通路不算，否則數字對不上使用者看到的東西 */}
                （只顯示 {scopeLabel || "—"}，要看其他通路請在上方勾起來）
              </span>
            )}
          </div>
          <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
            <Btn onClick={exportCsv}>
              <Download size={12} /> 匯出報表
            </Btn>
            <div style={{ position: "relative" }}>
              <Search
                size={13}
                color="var(--t4)"
                style={{
                  position: "absolute",
                  left: 10,
                  top: "50%",
                  transform: "translateY(-50%)",
                }}
              />
              <input
                type="text"
                placeholder="搜尋單號／備註／商品..."
                aria-label="搜尋門市訂單"
                value={q}
                onChange={(e) => setQ(e.target.value)}
                style={{
                  ...inp,
                  width: 200,
                  textAlign: "left",
                  paddingLeft: 30,
                  borderRadius: 10,
                  padding: "7px 12px 7px 30px",
                  fontSize: 12,
                }}
              />
            </div>
            <label
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                fontSize: 11,
                fontWeight: 600,
                color: "var(--t3)",
                cursor: "pointer",
              }}
            >
              <input
                type="checkbox"
                checked={lossOnly}
                onChange={(e) => setLossOnly(e.target.checked)}
                style={{ accentColor: "var(--dn)" }}
              />{" "}
              只看虧損（本期）
            </label>
          </div>
        </div>
        <div
          style={{
            overflowX: "auto",
            overflowY: "auto",
            maxHeight: 500,
            border: "1px solid var(--s3)",
            borderRadius: 12,
          }}
        >
          {/* tb-ord：手機隱藏備註／成本／毛利三欄，留單號·通路·商品·營收·淨利·發票 */}
          <table
            className="tb-ord"
            style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}
          >
            <thead>
              <tr>
                <SortTh sortKey="date" currentSort={sort} onSort={onSort}>
                  單號
                </SortTh>
                <th scope="col" style={{ ...th, textAlign: "left" }}>
                  通路
                </th>
                <th scope="col" style={{ ...th, textAlign: "left" }}>
                  商品
                </th>
                <th scope="col" style={{ ...th, textAlign: "left" }}>
                  備註
                </th>
                <SortTh sortKey="revenue" currentSort={sort} onSort={onSort} align="right">
                  營收
                </SortTh>
                <SortTh sortKey="cost" currentSort={sort} onSort={onSort} align="right">
                  成本
                </SortTh>
                <SortTh sortKey="gp" currentSort={sort} onSort={onSort} align="right">
                  毛利
                </SortTh>
                <SortTh sortKey="net" currentSort={sort} onSort={onSort} align="right">
                  最終淨利
                </SortTh>
                <th scope="col" style={{ ...th, textAlign: "right" }}>
                  發票
                </th>
              </tr>
            </thead>
            <tbody>
              {paged.length > 0 ? (
                paged.map((o) => {
                  const isLoss = !o.missCost && o.net < 0;
                  const isOpen = openId === o.orderId;
                  const gm = o.revenue > 0 ? o.gp / o.revenue : 0;
                  /* 兩種開單異常：毛利率 >85%＝沒選克數規格（總價打在數量 1 上）；
                     單價≤1 且數量>100＝單價／數量顛倒（成本被放大上千倍、毛利極負） */
                  const suspect = (!o.missCost && gm > 0.85) || o.swapSuspect;
                  const suspectMsg = o.swapSuspect
                    ? "疑似單價／數量顛倒：有商品單價≤1 元、數量>100（總價打進數量欄），成本被放大，請到 POS 修單"
                    : "毛利率異常高：可能是開單沒選克數規格（總價打在數量 1 上）";
                  const items = o.items || [];
                  const summary =
                    items.length === 0
                      ? "（無明細）"
                      : items.length === 1
                      ? items[0].name
                      : `${items[0].name} 等 ${items.length} 件`;
                  return (
                    <React.Fragment key={o.orderId}>
                      <tr
                        className={isLoss ? "rl" : ""}
                        tabIndex={0}
                        aria-expanded={isOpen}
                        onClick={() => setOpenId(isOpen ? null : o.orderId)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter" || e.key === " ") {
                            e.preventDefault();
                            setOpenId(isOpen ? null : o.orderId);
                          }
                        }}
                        /* 表身已跟著勾選過濾，不會再出現未計入通路的列，
                           灰掉的樣式留著當保險（例如重匯後通路 key 對不上時） */
                        style={{ cursor: "pointer", opacity: o.excludedCh ? 0.5 : 1 }}
                      >
                        <td style={{ ...td2 }}>
                          <div
                            style={{
                              fontWeight: 600,
                              fontSize: 12,
                              display: "flex",
                              alignItems: "center",
                              gap: 4,
                            }}
                          >
                            {isOpen ? (
                              <ChevronUp size={11} color="var(--t3)" />
                            ) : (
                              <ChevronDown size={11} color="var(--t3)" />
                            )}
                            {o.date}
                          </div>
                          <div
                            style={{
                              fontSize: 10,
                              color: "var(--t3)",
                              marginTop: 2,
                              fontFamily: mono,
                            }}
                          >
                            {o.orderId}
                          </div>
                        </td>
                        <td style={{ ...td2 }}>
                          <div
                            style={{
                              fontSize: 12,
                              fontWeight: o.channel === "dealer" ? 700 : 600,
                              color: o.channel === "dealer" ? accentColor : "var(--t2)",
                            }}
                          >
                            {o.channelLabel}
                          </div>
                          <div style={{ fontSize: 10, color: "var(--t4)", marginTop: 2 }}>
                            {o.payMethod || "—"}
                          </div>
                        </td>
                        <td style={{ ...td2, maxWidth: 200 }}>
                          <div
                            style={{
                              fontSize: 11,
                              fontWeight: 600,
                              color: items.length ? "var(--t2)" : "var(--wn)",
                              whiteSpace: "nowrap",
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                              maxWidth: 190,
                            }}
                            title={items.map((i) => i.name).join("、")}
                          >
                            {summary}
                          </div>
                        </td>
                        <td
                          style={{
                            ...td2,
                            fontSize: 11,
                            color: "var(--t3)",
                            maxWidth: 120,
                            whiteSpace: "nowrap",
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                          }}
                        >
                          {o.remark || "—"}
                        </td>
                        <td
                          style={{ ...td2, textAlign: "right", fontFamily: mono, fontWeight: 600 }}
                        >
                          {fmt$(o.revenue)}
                        </td>
                        <td style={{ ...td2, textAlign: "right", fontFamily: mono, color: "var(--dn)" }}>
                          {o.missCost ? (
                            <span style={{ color: "var(--wn)", fontSize: 10, fontWeight: 700 }}>
                              {items.length ? "未填" : "無明細"}
                            </span>
                          ) : (
                            `-${fmt$(o.oCost)}`
                          )}
                        </td>
                        <td
                          style={{ ...td2, textAlign: "right", fontFamily: mono, fontWeight: 600 }}
                          title={suspect ? suspectMsg : undefined}
                        >
                          {o.missCost ? (
                            <span style={{ color: "var(--t4)" }}>
                              —{o.swapSuspect ? " ⚠" : ""}
                            </span>
                          ) : (
                            <>
                              {fmt$(o.gp)}
                              <span
                                style={{
                                  fontSize: 10,
                                  color: suspect ? "var(--wn)" : "var(--t4)",
                                  marginLeft: 4,
                                }}
                              >
                                {fmtP(gm)}
                                {suspect ? "⚠" : ""}
                              </span>
                            </>
                          )}
                        </td>
                        <td
                          style={{
                            ...td2,
                            textAlign: "right",
                            fontFamily: mono,
                            fontWeight: 800,
                            color: o.missCost
                              ? "var(--t4)"
                              : isLoss
                              ? "var(--dn)"
                              : accentColor,
                          }}
                        >
                          {o.missCost ? "—" : fmt$(o.net)}
                        </td>
                        <td
                          style={{ ...td2, textAlign: "right", fontSize: 10 }}
                          title={
                            o.hasInvoice
                              ? `發票判定來源：${
                                  typeof o.invoiceOverride === "boolean"
                                    ? "手動設定"
                                    : o.invoiceSrc || "—"
                                }`
                              : "未開發票（不課稅）"
                          }
                        >
                          {o.hasInvoice ? (
                            <span style={{ color: "var(--up)" }}>
                              已開
                              {typeof o.invoiceOverride === "boolean"
                                ? "·手"
                                : o.invoiceSrc?.startsWith("備註")
                                ? "·註"
                                : o.invoiceSrc?.startsWith("統編")
                                ? "·統"
                                : o.invoiceSrc?.startsWith("企業")
                                ? "·企"
                                : o.invoiceSrc?.startsWith("合作")
                                ? "·合"
                                : ""}
                            </span>
                          ) : (
                            <span style={{ color: "var(--t4)" }}>
                              {typeof o.invoiceOverride === "boolean" ? "未開·手" : "—"}
                            </span>
                          )}
                        </td>
                      </tr>
                      {isOpen && (
                        <tr onClick={(e) => e.stopPropagation()}>
                          <td
                            colSpan={9}
                            style={{ ...td2, background: "var(--s2)", padding: "16px 20px" }}
                          >
                            <POSOrderDetail
                              order={o}
                              costsEff={costsEff}
                              recipes={recipes}
                              components={components}
                              ratios={ratios}
                              onToggleInvoice={onToggleInvoice}
                            />
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })
              ) : (
                <tr>
                  <td
                    colSpan={9}
                    style={{ ...td2, textAlign: "center", color: "var(--t4)", padding: 40 }}
                  >
                    找不到符合條件的訂單
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        {/* Pagination（同官網） */}
        {filtered.length > pageSize && (
          <div
            style={{
              padding: "12px 4px 0",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              flexWrap: "wrap",
              gap: 8,
            }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontSize: 11, color: "var(--t3)", fontFamily: mono }}>
                {curPage * pageSize + 1}–
                {Math.min((curPage + 1) * pageSize, filtered.length)} / {filtered.length} 筆
              </span>
              <span style={{ fontSize: 10, color: "var(--t4)" }}>每頁</span>
              <select
                value={pageSize}
                aria-label="每頁筆數"
                onChange={(e) => {
                  setPageSize(Number(e.target.value));
                  setPage(0);
                }}
                style={{ ...sel, padding: "3px 8px", fontSize: 11, fontFamily: mono }}
              >
                {[20, 30, 50, 100].map((n) => (
                  <option key={n} value={n}>
                    {n}
                  </option>
                ))}
              </select>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
              {[
                { label: "«", aria: "第一頁", action: () => setPage(0) },
                { label: "‹", aria: "上一頁", action: () => setPage(Math.max(0, curPage - 1)) },
                null,
                {
                  label: "›",
                  aria: "下一頁",
                  action: () => setPage(Math.min(totalPages - 1, curPage + 1)),
                },
                { label: "»", aria: "最後一頁", action: () => setPage(totalPages - 1) },
              ].map((btn, i) =>
                btn === null ? (
                  <span
                    key={i}
                    style={{
                      fontSize: 12,
                      fontWeight: 800,
                      color: "var(--t1)",
                      fontFamily: mono,
                      padding: "0 10px",
                    }}
                  >
                    {curPage + 1} / {totalPages}
                  </span>
                ) : (
                  <Btn key={i} aria-label={btn.aria} onClick={btn.action} style={{ padding: "4px 10px" }}>
                    {btn.label}
                  </Btn>
                )
              )}
            </div>
          </div>
        )}
      </div>
    </>
  );
}

/* ─── Overview Dashboard ────────────────────────────────────── */
function OverviewDashboard({
  slData,
  spData,
  posData,
  posTarget,
  slOrders,
  spOrders,
  posOrders,
  slCosts,
  spCosts,
  allMonthly,
  monthlyByPlatform,
  theme,
  onNavigate,
  sY,
  sM,
  range,
}) {
  const slD = slData?.summary;
  const spS = spData?.s;
  const isDark = theme === "dark";
  const narrow = useIsNarrow();
  const greenC = isDark ? "#2ECC71" : "#1A6B3C";
  const spC = isDark ? "#FF6533" : "#EE4D2D";
  const posC = isDark ? "#9B7FCA" : "#7B5EA7";
  const goldC = isDark ? "#C9A84C" : "#8B6914";
  const gridC = isDark ? "rgba(255,255,255,0.05)" : "rgba(0,0,0,0.05)";

  /* 門市在總覽一律「六通路全算」（不受門市頁的計入勾選影響）；
     淨利只算成本齊全的單（算不出成本的單計營收、不計淨利，下方會標示） */
  const posCh = useMemo(() => posData?.channelsAll || [], [posData]);
  const posS = useMemo(() => {
    if (!posData) return null;
    const sum = sumPosChannels(posCh);
    /* 總覽的 net 語意＝成本齊全單的淨利（門市頁叫 coveredNet），沿用舊欄位名 */
    const t = { ...sum, net: sum.coveredNet };
    return {
      ...t,
      netMargin: t.coveredRev > 0 ? t.net / t.coveredRev : 0,
      grossMargin: t.coveredRev > 0 ? t.coveredGp / t.coveredRev : 0,
      aov: t.orders > 0 ? t.rev / t.orders : 0,
      coverage: t.rev > 0 ? (t.rev - t.noCostRev) / t.rev : 0,
    };
  }, [posData, posCh]);
  /* 六通路表的小條：改用「佔全公司」後數字會很小（十幾 %），
     條長改成相對最大通路，視覺才看得出差距；精確比例看旁邊那欄 */
  const posChMax = useMemo(
    () => Math.max(0, ...posCh.map((c) => Number(c.rev) || 0)),
    [posCh]
  );

  const hasAny = slD || spS || posS;
  /* 合計＝全公司（官網＋蝦皮＋門市六通路）；但呈現上門市不包成一包：
     「門市」＝現場零售，其他五個通路各自一項（老闆 2026-08-19 定） */
  const totalRev = (slD?.rev || 0) + (spS?.tG || 0) + (posS?.rev || 0);
  const totalNet = (slD?.net || 0) + (spS?.afterComm || 0) + (posS?.net || 0);
  const totalNetMargin = totalRev > 0 ? totalNet / totalRev : 0;
  /* 總覽固定目標 12%、成長油門帶 10–12、紅線 10（2026-08-25 老闆拍板；加權線提案已否決） */
  const OVERALL_TARGET = 0.12;
  const OVERALL_RED = 0.1;
  /* 範圍內平均營業費率：營收加權（各訂單已依所屬月份快照 % 計算），
     選整年即為 1-12 月的加權平均 */
  const totalOpex = (slD?.opExpTotal || 0) + (spS?.tOp || 0) + (posS?.op || 0);
  const totalOpexRate = totalRev > 0 ? totalOpex / totalRev : 0;
  const slRevShare = totalRev > 0 ? (slD?.rev || 0) / totalRev : 0;
  const spRevShare = totalRev > 0 ? (spS?.tG || 0) / totalRev : 0;
  const posRetail = posCh.find((c) => c.key === "retail") || null;
  const posOthers = posCh.filter((c) => c.key !== "retail");
  const chNet = (c) => (c && c.coveredRev > 0 ? c.coveredNet / c.coveredRev : 0);
  /* 通路色：門市零售＝紫；其他五通路各自一色（老闆 2026-09-05 要「分開列出來分色」）。
     五色用 dataviz 驗色工具做 OKLCH 搜尋挑出：照圓餅相鄰序（紫→經銷→電話→Omnichat→合作
     →企業→官網綠）在淺／深兩模式相鄰色弱 ΔE≥8、正常視力 ≥15 全過；避開狀態紅／琥珀與
     三平台色相；金＋橄欖綠在色弱模擬下同色（ΔE 2.6）故只留金。最接近的一對是 Omnichat 藍
     vs 企業天藍（非相鄰），靠圖例名字補；深色唯一不及格是官網綠 #2ECC71 自己的亮度帶
     （app 既有 --up，不動）。顏色跟著實體走：進度條、圓餅、圖例、下方五通路卡片同一張表 */
  const posChColor = isDark
    ? { dealer: "#AF8E2A", phone: "#3CA694", omnichat: "#4693F1", partner: "#DD628E", corp: "#1AA1C9" }
    : { dealer: "#967600", phone: "#289785", omnichat: "#1467C2", partner: "#AD3665", corp: "#089AC3" };
  const chColor = (c) => posChColor[c.key] || posC;

  const periodLabel =
    sY === "Custom"
      ? `${range?.from || "起"} ～ ${range?.to || "迄"}`
      : sY === "All"
      ? "歷年"
      : sM === "All"
      ? `${sY}年`
      : `${sY}/${sM}`;

  const alerts = useMemo(() => {
    const list = [];
    if (slD && slD.valid > 0 && slD.trueNetMargin < slD.targetNetRate)
      list.push({
        level: "warn",
        platform: "官網",
        msg: `淨利率 ${fmtP(slD.trueNetMargin)} 低於目標 ${fmtP(
          slD.targetNetRate
        )}，差距 ${Math.abs(slD.gapVal).toFixed(1)}%`,
      });
    if (spS && spS.validN > 0 && spS.netMargin < spS.targetNet)
      list.push({
        level: "warn",
        platform: "蝦皮",
        msg: `淨利率 ${fmtP(spS.netMargin)} 低於目標 ${fmtP(
          spS.targetNet
        )}，差距 ${((spS.targetNet - spS.netMargin) * 100).toFixed(1)}%`,
      });

    if (slD && slD.lossCount > 0)
      list.push({
        level: "info",
        platform: "官網",
        msg: `本期有 ${slD.lossCount} 筆虧損訂單，建議檢視運費與折扣設定`,
      });
    if (spS && spS.lossN > 0)
      list.push({
        level: "info",
        platform: "蝦皮",
        msg: `本期有 ${spS.lossN} 筆虧損訂單`,
      });
    if (slD && slD.returnRate > 0.05)
      list.push({
        level: "warn",
        platform: "官網",
        msg: `退貨率 ${fmtP(slD.returnRate)}（${
          slD.returnCount
        } 筆），高於 5% 警戒線`,
      });
    if (spS && spS.refundN > 0)
      list.push({
        level: "info",
        platform: "蝦皮",
        msg: `本期排除 ${spS.refundN} 筆退貨/退款訂單（${fmt$(
          spS.refundG
        )} 未計入營收）`,
      });
    if (
      slData?.matrixList?.some(
        (p) =>
          p.soldQty > 0 && (!slCosts[p.key] || Number(slCosts[p.key]) === 0)
      )
    )
      list.push({
        level: "error",
        platform: "官網",
        msg: "有商品成本未填，淨利計算可能偏高",
      });
    if (
      spData?.uniqueProducts?.some(
        (p) =>
          p.soldQty > 0 && (!spCosts[p.key] || Number(spCosts[p.key]) === 0)
      )
    )
      list.push({
        level: "error",
        platform: "蝦皮",
        msg: "有商品成本未填，淨利計算可能偏高",
      });
    /* 門市（現場零售）對門市統一目標；算不出成本的單、虧損單（六通路）、經銷單淨利 */
    {
      const r = posCh.find((c) => c.key === "retail");
      const rm = r && r.coveredRev > 0 ? r.coveredNet / r.coveredRev : null;
      if (rm !== null && rm < posTarget)
        list.push({
          level: "warn",
          platform: "門市",
          msg: `現場零售淨利率 ${fmtP(rm)} 低於門市目標 ${fmtP(
            posTarget
          )}，差距 ${((posTarget - rm) * 100).toFixed(1)}%`,
        });
    }
    if (posS && posS.noCostCount > 0)
      list.push({
        level: posS.coverage < 0.6 ? "error" : "warn",
        platform: "門市",
        msg: `${posS.noCostCount} 筆（${fmt$(
          posS.noCostRev
        )}）算不出成本：品項未選正規或缺商品明細——這部分只計營收、不計淨利，成本覆蓋率 ${fmtP(
          posS.coverage
        )}`,
      });
    if (posS && posS.lossCount > 0)
      list.push({
        level: "info",
        platform: "門市",
        msg: `本期有 ${posS.lossCount} 筆虧損訂單`,
      });
    {
      const dealer = posCh.find((c) => c.key === "dealer");
      if (dealer && dealer.coveredRev > 0 && dealer.coveredNet / dealer.coveredRev < 0.05)
        list.push({
          level: "warn",
          platform: "門市",
          msg: `經銷·老客價 淨利率僅 ${fmtP(
            dealer.coveredNet / dealer.coveredRev
          )}（${dealer.orders} 筆 ${fmt$(dealer.rev)}），接近打平`,
        });
    }
    /* 月中 run-rate 預警：檢視「本月」且已過 7 日，用上月同期進度對比，
       落後 10% 以上提前示警，不必等月底才發現 */
    if (sY !== "Custom" && sY !== "All" && sM !== "All") {
      const now = new Date();
      const curYm = `${now.getFullYear()}-${String(
        now.getMonth() + 1
      ).padStart(2, "0")}`;
      if (`${sY}-${sM}` === curYm) {
        /* 進度基準用「已匯入資料的最後一天」而非今天：
           報表是手動匯入的，晚幾天看不該被誤判為落後。
           三通路各自取匯入進度、只拿「本月已有資料的通路」來比，且分子分母同一組通路
           （門市通常月底才匯，若本月還沒匯就不納入比較，不會把 pace 壓到 0） */
        const lastOf = (src) => {
          let d0 = 0;
          Object.values(src || {}).forEach((o) => {
            const d = String(o.date || "");
            if (d.startsWith(curYm)) {
              const dd = Number(d.substring(8, 10)) || 0;
              if (dd > d0) d0 = dd;
            }
          });
          return d0;
        };
        const mNum = Number(sM);
        const prevYm =
          mNum === 1
            ? `${Number(sY) - 1}-12`
            : `${sY}-${String(mNum - 1).padStart(2, "0")}`;
        const plats = [
          { k: "sl", last: lastOf(slOrders), cur: slD?.rev || 0 },
          { k: "sp", last: lastOf(spOrders), cur: spS?.tG || 0 },
          { k: "pos", last: lastOf(posOrders), cur: posS?.rev || 0 },
        ].filter((p) => p.last > 0);
        const lastDay = plats.length ? Math.min(...plats.map((p) => p.last)) : 0;
        const effDay = Math.min(now.getDate(), lastDay);
        const prevRev = plats.reduce(
          (s, p) => s + (monthlyByPlatform?.[p.k]?.[prevYm]?.rev || 0),
          0
        );
        /* 分母用「上月」天數（prevRev 是上月的量），並封頂 1 */
        const prevDays = new Date(Number(sY), mNum - 1, 0).getDate();
        const frac = Math.min(1, effDay / prevDays);
        const curRev = plats.reduce((s, p) => s + p.cur, 0);
        if (effDay >= 7 && prevRev > 0 && frac > 0 && curRev > 0) {
          const pace = curRev / (prevRev * frac);
          if (pace < 0.9)
            list.push({
              level: "warn",
              platform: "全站",
              msg: `本月營收進度僅為上月同期的 ${(pace * 100).toFixed(
                0
              )}%（依已匯入資料，截至 ${effDay} 日），留意月底達成`,
            });
          else if (pace >= 1.1)
            list.push({
              level: "info",
              platform: "全站",
              msg: `本月營收進度為上月同期的 ${(pace * 100).toFixed(
                0
              )}%（截至 ${effDay} 日），超前上月步調`,
            });
        }
      }
    }
    return list;
  }, [
    posTarget,
    slD,
    spS,
    posS,
    posCh,
    slData,
    spData,
    slCosts,
    spCosts,
    sY,
    sM,
    monthlyByPlatform,
    slOrders,
    spOrders,
    posOrders,
  ]);

  /* 同比／環比有一端沒門市資料（門市 2026-08 才上線）→ 不是同口徑，標出來
     避免「新增一個通路」被讀成官網蝦皮成長 */
  const posCompareNote = useMemo(() => {
    const pm = monthlyByPlatform?.pos;
    if (!pm || sY === "All" || sY === "Custom" || sM === "All") return null;
    const cur = `${sY}-${sM}`;
    if (!pm[cur]) return null;
    const mNum = Number(sM);
    const prevYm =
      mNum === 1 ? `${Number(sY) - 1}-12` : `${sY}-${String(mNum - 1).padStart(2, "0")}`;
    const yoyYm = `${Number(sY) - 1}-${sM}`;
    const miss = [];
    if (!pm[prevYm]) miss.push("環比");
    if (!pm[yoyYm]) miss.push("同比");
    return miss.length
      ? `※ ${miss.join("／")}含門市但比較基期無門市資料（門市 2026-08 起才有），非同口徑`
      : null;
  }, [monthlyByPlatform, sY, sM]);

  /* 月度趨勢固定顯示整年（或歷年最近 12 個月），不受單月篩選影響；
     淨利線取自 allMonthly（已扣分潤）；自訂區間因非整月，不畫淨利線 */
  const trendData = useMemo(() => {
    const byMonth = {};
    const passPeriod = (d) => {
      if (sY === "Custom") {
        if (range?.from && d < range.from) return false;
        if (range?.to && d > range.to) return false;
        return true;
      }
      if (sY && sY !== "All" && !d.startsWith(sY)) return false;
      return true;
    };
    Object.values(slOrders || {}).forEach((o) => {
      const cx =
        (o.status || "").includes("取消") || (o.status || "").includes("刪除");
      if (cx) return;
      const d = String(o.date || "");
      if (!passPeriod(d)) return;
      const ym = d.substring(0, 7);
      if (!ym || ym.length < 7) return;
      if (!byMonth[ym])
        byMonth[ym] = { month: ym, slRev: 0, spRev: 0, posRev: 0, posOtherRev: 0 };
      byMonth[ym].slRev += o.revenue || 0;
    });
    Object.values(spOrders || {}).forEach((o) => {
      const st = String(o.status || ""),
        rf = String(o.refundStatus || "");
      const bad =
        st.includes("不成立") ||
        st.includes("取消") ||
        rf !== "" ||
        (st.includes("退貨") && !st.includes("已完成"));
      if (bad) return;
      const d = String(o.date || "");
      if (!passPeriod(d)) return;
      const ym = d.substring(0, 7);
      if (!ym || ym.length < 7) return;
      if (!byMonth[ym])
        byMonth[ym] = { month: ym, slRev: 0, spRev: 0, posRev: 0, posOtherRev: 0 };
      const gross = o.grossPrice || 0;
      byMonth[ym].spRev += gross;
    });
    /* 門市：取消／退款／測試單排除，六通路全算（與總覽口徑一致） */
    Object.values(posOrders || {}).forEach((o) => {
      const st = String(o.status || "");
      if (st.includes("取消") || st.includes("退款") || st.includes("已退")) return;
      const g = Number(o.revenue) || 0;
      if (g > 0 && g <= POS_TEST_MAX) return;
      const d = String(o.date || "");
      if (!passPeriod(d)) return;
      const ym = d.substring(0, 7);
      if (!ym || ym.length < 7) return;
      if (!byMonth[ym])
        byMonth[ym] = { month: ym, slRev: 0, spRev: 0, posRev: 0, posOtherRev: 0 };
      if (o.channel === "retail") byMonth[ym].posRev += g;
      else byMonth[ym].posOtherRev += g;
    });
    return Object.values(byMonth)
      .sort((a, b) => a.month.localeCompare(b.month))
      .slice(-12)
      .map((d) => ({
        ...d,
        posRev: d.posRev || 0,
        posOtherRev: d.posOtherRev || 0,
        label: d.month.substring(2).replace("-", "/"),
        total: d.slRev + d.spRev + (d.posRev || 0) + (d.posOtherRev || 0),
        net: sY === "Custom" ? undefined : allMonthly?.[d.month]?.net,
      }));
  }, [slOrders, spOrders, posOrders, sY, range, allMonthly]);
  const showNetLine = trendData.some((d) => d.net !== undefined);

  const crossProductRank = useMemo(() => {
    const map = {};
    (slData?.matrixList || []).forEach((p) => {
      if (!map[p.name]) map[p.name] = { name: p.name, slQty: 0, spQty: 0 };
      map[p.name].slQty += p.soldQty || 0;
    });
    (spData?.uniqueProducts || []).forEach((p) => {
      if (!map[p.name]) map[p.name] = { name: p.name, slQty: 0, spQty: 0 };
      map[p.name].spQty += p.soldQty || 0;
    });
    return Object.values(map)
      .filter((p) => p.slQty + p.spQty > 0)
      .sort((a, b) => b.slQty + b.spQty - (a.slQty + a.spQty))
      .slice(0, 8);
  }, [slData, spData]);

  if (!hasAny) {
    return (
      <div
        className="f0"
        style={{
          minHeight: 400,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          flexDirection: "column",
          gap: 14,
          background: "var(--s1)",
          border: "1px solid var(--s3)",
          borderRadius: 16,
        }}
      >
        <BarChart3 size={40} color="var(--s4)" />
        <div style={{ fontSize: 15, fontWeight: 700, color: "var(--t3)" }}>
          尚無任何資料
        </div>
        <div style={{ fontSize: 12, color: "var(--t4)" }}>
          請先上傳官網、蝦皮或門市報表
        </div>
      </div>
    );
  }

  /* 綠＝達目標 12；黃＝成長油門帶 10–12；紅＝跌破 10。
     10–12 這段不是警戒是刻意留的空間：2 個百分點 ≈ 84 萬/年，用來加碼投放
     （老闆 2026-09-03 拍板，9 月起執行；淨利 10% 是地板不是目標） */
  const overallStatus =
    totalNetMargin >= OVERALL_TARGET
      ? { label: "整體健康", c: "var(--up)" }
      : totalNetMargin >= OVERALL_RED
      ? { label: "成長油門帶", c: "var(--wn)" }
      : totalNetMargin > 0
      ? { label: "跌破紅線 10%", c: "var(--dn)" }
      : { label: "整體虧損", c: "var(--dn)" };

  /* 全通路營收圓餅（老闆 2026-09-05 要的）：八片＝官網／蝦皮／門市現場零售＋門市其他五通路
     各自一片一色（09-05 第二版，原本合成一片灰的「門市其他」拆開）。片序＝進度條與圖例序，
     色由 posChColor 統一。直接標籤只給 ≥4% 的片，其餘靠圖例；片與片之間留 2px 底色縫 */
  const pieShort = {
    官網: "官網",
    蝦皮: "蝦皮",
    "門市（現場零售）": "門市",
    "經銷·老客價": "經銷",
    電話訂購: "電話",
    Omnichat: "Omni",
    合作通路: "合作",
    企業採購: "企業",
  };
  const pieData = [
    { l: "官網", c: greenC, op: 1, v: slD?.rev || 0 },
    { l: "蝦皮", c: spC, op: 0.85, v: spS?.tG || 0 },
    { l: "門市（現場零售）", c: posC, op: 0.9, v: posRetail?.rev || 0 },
    ...posOthers.map((c) => ({ l: c.label, c: chColor(c), op: 1, v: c.rev || 0 })),
  ].filter((d) => d.v > 0);
  /* 手機把整個環縮一號，外圈標籤才不會被螢幕切掉（桌機維持 320×180 原尺寸） */
  const pieW = narrow ? 300 : 320;
  const pieH = narrow ? 168 : 180;
  const pieIR = narrow ? 42 : 50;
  const pieOR = narrow ? 62 : 74;
  const RAD = Math.PI / 180;
  const renderPieLabel = ({ cx, cy, midAngle, outerRadius, percent, name }) => {
    if (!(percent >= 0.04)) return null;
    const r = outerRadius + (narrow ? 10 : 12);
    const x = cx + Math.cos(-midAngle * RAD) * r;
    const y = cy + Math.sin(-midAngle * RAD) * r;
    return (
      <text
        x={x}
        y={y}
        fill="var(--t2)"
        fontSize={10}
        fontWeight={700}
        fontFamily={mono}
        textAnchor={x > cx ? "start" : "end"}
        dominantBaseline="central"
      >
        {pieShort[name] || name} {(percent * 100).toFixed(1)}%
      </text>
    );
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      {totalRev === 0 && (
        <div
          className="f0"
          style={{
            background: "var(--wn-dim)",
            border: "1px solid var(--wn-bdr)",
            borderRadius: 14,
            padding: "14px 20px",
            display: "flex",
            alignItems: "center",
            gap: 10,
          }}
        >
          <AlertTriangle size={16} color="var(--wn)" />
          <div
            style={{
              fontSize: 12,
              color: "var(--t1)",
              fontWeight: 600,
              lineHeight: 1.6,
            }}
          >
            此期間三個通路都沒有有效訂單，下方數字全為
            0——請先確認期間選擇是否正確、該期間報表是否已匯入。
          </div>
        </div>
      )}
      {/* ── 老闆月報 Hero ── */}
      <div
        className="f1"
        style={{
          background: "var(--s1)",
          border: "1px solid var(--s3)",
          borderRadius: 16,
          overflow: "hidden",
          position: "relative",
        }}
      >
        <div
          style={{
            height: 3,
            background: `linear-gradient(90deg, ${greenC}, ${spC}, ${posC})`,
          }}
        />
        <div style={{ padding: "28px 36px" }}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 12,
              marginBottom: 20,
              flexWrap: "wrap",
            }}
          >
            <div
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 6,
                padding: "5px 14px",
                borderRadius: 99,
                background: `${overallStatus.c}15`,
                border: `1px solid ${overallStatus.c}40`,
                fontSize: 12,
                fontWeight: 700,
                color: overallStatus.c,
              }}
            >
              <div
                style={{
                  width: 6,
                  height: 6,
                  borderRadius: 99,
                  background: overallStatus.c,
                }}
              />
              {overallStatus.label}
            </div>
            <span style={{ fontSize: 12, color: "var(--t3)", fontWeight: 600 }}>
              {periodLabel} 跨平台合計
            </span>
          </div>

          <div
            className="hero-row"
            style={{
              display: "flex",
              flexWrap: "wrap",
              gap: 40,
              alignItems: "flex-end",
              marginBottom: 28,
            }}
          >
            <div>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 600,
                  color: "var(--t3)",
                  marginBottom: 6,
                  letterSpacing: "0.08em",
                  textTransform: "uppercase",
                }}
              >
                合計淨利
              </div>
              <div
                className="hero-num-md"
                style={{
                  lineHeight: 1,
                  fontWeight: 700,
                  letterSpacing: "-0.04em",
                  fontFamily: mono,
                  color: totalNet >= 0 ? "var(--t1)" : "var(--dn)",
                }}
              >
                {fmt$(totalNet)}
              </div>
              <div style={{ fontSize: 12, color: "var(--t4)", marginTop: 8 }}>
                合計營收 {fmt$(totalRev)}
                {posS && posS.noCostCount > 0
                  ? `（門市 ${posS.noCostCount} 筆算不出成本，計營收不計淨利）`
                  : ""}
              </div>
              <PeriodCompare monthly={allMonthly} sY={sY} sM={sM} />
              {posCompareNote && (
                <div style={{ fontSize: 10, color: "var(--wn)", marginTop: 4 }}>
                  {posCompareNote}
                </div>
              )}
            </div>
            <div>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 600,
                  color: "var(--t3)",
                  marginBottom: 6,
                }}
              >
                綜合淨利率（目標 12%）
              </div>
              <div
                className="hero-pct-md"
                style={{
                  fontWeight: 700,
                  fontFamily: mono,
                  lineHeight: 1,
                  color:
                    totalNetMargin >= OVERALL_TARGET
                      ? "var(--up)"
                      : totalNetMargin >= OVERALL_RED
                      ? "var(--wn)"
                      : "var(--dn)",
                }}
              >
                {fmtP(totalNetMargin)}
              </div>
            </div>
            <div>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 600,
                  color: "var(--t3)",
                  marginBottom: 6,
                }}
              >
                平均營業費率
              </div>
              <div
                className="hero-pct-md"
                style={{
                  fontWeight: 700,
                  fontFamily: mono,
                  lineHeight: 1,
                  /* OPEX 治理線（老闆自己的 SOP）：年度預算目標 30%、FY2026 可接受帶 33%
                     （員旅 80 萬＋年終 1 個月＝營收 2.9%，硬 30% 今年數學上不可達）、
                     活動月安全上限 35%（中秋、年節這種要加碼投放的月份）、>35% 才需檢討。
                     注意：費用低不等於好——營收沒成長時的低費用是警訊不是成績 */
                  color:
                    totalRev === 0
                      ? "var(--t2)"
                      : opexBandOf(totalOpexRate) === 0
                      ? "var(--up)"
                      : opexBandOf(totalOpexRate) <= 2
                      ? "var(--wn)"
                      : "var(--dn)",
                }}
              >
                {fmtP(totalOpexRate)}
              </div>
              {totalRev > 0 && (
                <div
                  style={{
                    fontSize: 10,
                    fontWeight: 600,
                    marginTop: 6,
                    color:
                      opexBandOf(totalOpexRate) === 0
                        ? "var(--up)"
                        : opexBandOf(totalOpexRate) <= 2
                        ? "var(--wn)"
                        : "var(--dn)",
                  }}
                >
                  {
                    [
                      "✓ 在 OPEX 目標 30% 內（年度預算線）",
                      "⚠ 高於目標 30%，仍在 FY2026 可接受帶 33% 內",
                      "⚠ 33–35%：活動月加碼可接受，淡月要收回來",
                      "⚠ 超過活動月上限 35%，需檢討",
                    ][opexBandOf(totalOpexRate)]
                  }
                </div>
              )}
              <div style={{ fontSize: 10, color: "var(--t4)", marginTop: 2 }}>
                依各月快照％營收加權（含廣告）
              </div>
            </div>
            {/* 全通路營收圓餅：放 Hero 列右側空位（老闆 2026-09-05 指定）。寬 320 是給外側
                直接標籤留位（最長「門市 38.2%」），環在正中；視窗窄時自動換行到下一列 */}
            {pieData.length > 0 && (
              <div
                className="pie-box"
                style={{
                  position: "relative",
                  width: pieW,
                  height: pieH,
                  flex: "0 0 auto",
                  marginLeft: "auto",
                  alignSelf: "center",
                }}
                role="img"
                aria-label={`全通路營收佔比：${pieData
                  .map((d) => `${d.l} ${((d.v / totalRev) * 100).toFixed(1)}%`)
                  .join("、")}`}
              >
                <RPieChart width={pieW} height={pieH}>
                  <Pie
                    data={pieData}
                    dataKey="v"
                    nameKey="l"
                    cx="50%"
                    cy="50%"
                    innerRadius={pieIR}
                    outerRadius={pieOR}
                    paddingAngle={1.5}
                    stroke="var(--s1)"
                    strokeWidth={2}
                    isAnimationActive={false}
                    labelLine={false}
                    label={renderPieLabel}
                  >
                    {pieData.map((d) => (
                      <Cell key={d.l} fill={d.c} fillOpacity={d.op} />
                    ))}
                  </Pie>
                  <RTooltip
                    content={({ active, payload }) => {
                      if (!active || !payload?.length) return null;
                      const d = payload[0].payload;
                      return (
                        <div
                          style={{
                            background: "var(--s1)",
                            border: "1px solid var(--s3)",
                            borderRadius: 10,
                            padding: "10px 14px",
                            boxShadow: "0 4px 20px rgba(0,0,0,0.12)",
                            minWidth: 170,
                          }}
                        >
                          <div
                            style={{
                              display: "flex",
                              justifyContent: "space-between",
                              alignItems: "center",
                              gap: 14,
                              fontSize: 12,
                              fontWeight: 700,
                              color: "var(--t1)",
                            }}
                          >
                            <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                              <span
                                style={{
                                  width: 8,
                                  height: 8,
                                  borderRadius: 2,
                                  background: d.c,
                                  opacity: d.op,
                                  flex: "0 0 auto",
                                }}
                              />
                              {d.l}
                            </span>
                            <span style={{ fontFamily: mono }}>
                              {fmt$(d.v)} ·{" "}
                              {totalRev > 0 ? `${((d.v / totalRev) * 100).toFixed(1)}%` : "—"}
                            </span>
                          </div>
                        </div>
                      );
                    }}
                  />
                </RPieChart>
                {/* 環中心：合計營收（位置由圓餅尺寸算，手機縮圖也對得準） */}
                <div
                  style={{
                    position: "absolute",
                    left: pieW / 2 - 46,
                    top: pieH / 2 - 16,
                    width: 92,
                    textAlign: "center",
                    pointerEvents: "none",
                  }}
                >
                  <div style={{ fontSize: 9, color: "var(--t4)", fontWeight: 700 }}>合計</div>
                  <div
                    style={{
                      fontSize: 12,
                      fontWeight: 700,
                      fontFamily: mono,
                      color: "var(--t1)",
                      letterSpacing: "-0.02em",
                    }}
                  >
                    {fmt$(totalRev)}
                  </div>
                </div>
              </div>
            )}
          </div>

          <div style={{ marginBottom: 24 }}>
            <div
              style={{
                height: 8,
                borderRadius: 99,
                background: "var(--s3)",
                overflow: "hidden",
                display: "flex",
                marginBottom: 8,
              }}
            >
              <div
                style={{
                  width: `${slRevShare * 100}%`,
                  background: greenC,
                  transition: "width .6s",
                }}
              />
              <div
                style={{
                  width: `${spRevShare * 100}%`,
                  background: spC,
                  opacity: 0.85,
                  transition: "width .6s",
                }}
              />
              <div
                style={{
                  width: `${totalRev > 0 ? ((posRetail?.rev || 0) / totalRev) * 100 : 0}%`,
                  background: posC,
                  opacity: 0.9,
                  transition: "width .6s",
                }}
              />
              {posOthers.map((c) => (
                <div
                  key={c.key}
                  style={{
                    width: `${totalRev > 0 ? (c.rev / totalRev) * 100 : 0}%`,
                    background: chColor(c),
                    borderLeft: c.rev > 0 ? "1px solid var(--s1)" : "none",
                    transition: "width .6s",
                  }}
                />
              ))}
            </div>
            {/* 圖例：官網／蝦皮／門市（現場零售）／經銷／電話／Omnichat／合作／企業 ＝ 8 項全列，
                沒營收也列；色＝圓餅、進度條同一張表 */}
            <div
              className="ch-legend"
              style={{
                display: "flex",
                flexWrap: "wrap",
                gap: "6px 14px",
              }}
            >
              {[
                { l: "官網", c: greenC, op: 1, v: slD?.rev || 0 },
                { l: "蝦皮", c: spC, op: 0.85, v: spS?.tG || 0 },
                { l: "門市（現場零售）", c: posC, op: 0.9, v: posRetail?.rev || 0 },
                ...posOthers.map((c) => ({
                  l: c.label,
                  c: chColor(c),
                  op: 1,
                  v: c.rev,
                })),
              ].map((p) => {
                const share = totalRev > 0 ? p.v / totalRev : 0;
                const empty = p.v <= 0;
                return (
                  <div
                    key={p.l}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 6,
                      fontSize: 11,
                      fontWeight: 600,
                      opacity: empty ? 0.55 : 1,
                    }}
                  >
                    <div
                      style={{
                        width: 8,
                        height: 8,
                        borderRadius: 2,
                        background: p.c,
                        opacity: p.op,
                      }}
                    />
                    <span style={{ color: "var(--t2)" }}>{p.l}</span>
                    <span style={{ fontFamily: mono, color: empty ? "var(--t4)" : p.c }}>
                      {(share * 100).toFixed(1)}%
                    </span>
                    {/* 金額在手機隱藏（.lg-amt）：下面各平台卡片與通路卡片都看得到 */}
                    <span className="lg-amt" style={{ color: "var(--t4)" }}>
                      {empty ? "—" : fmt$(p.v)}
                    </span>
                  </div>
                );
              })}
            </div>
          </div>

          <div className="gcmp">
            {[
              {
                label: "官網",
                color: greenC,
                rev: slD?.rev || 0,
                net: slD?.net || 0,
                margin: slD?.trueNetMargin || 0,
                target: slD?.targetNetRate || 0.15,
                orders: slD?.valid || 0,
                opexRate:
                  slD && slD.rev > 0 ? slD.opExpTotal / slD.rev : 0,
                aov: slD && slD.valid > 0 ? slD.rev / slD.valid : 0,
                id: "shopline",
              },
              {
                label: "蝦皮",
                color: spC,
                rev: spS?.tG || 0,
                net: spS?.afterComm || 0,
                margin: spS?.netMargin || 0,
                target: spS?.targetNet || 0.1,
                orders: spS?.validN || 0,
                opexRate: spS && spS.tG > 0 ? spS.tOp / spS.tG : 0,
                aov: spS?.avgAOV || 0,
                id: "shopee",
              },
              {
                label: "門市（現場零售）",
                color: posC,
                rev: posRetail?.rev || 0,
                net: posRetail?.coveredNet || 0,
                margin: chNet(posRetail),
                target: posTarget,
                targetLabel: "目標",
                orders: posRetail?.orders || 0,
                opexRate:
                  posRetail && posRetail.rev > 0 ? posRetail.op / posRetail.rev : 0,
                aov:
                  posRetail && posRetail.orders > 0
                    ? posRetail.rev / posRetail.orders
                    : 0,
                note:
                  posRetail && posRetail.noCostCount > 0
                    ? `${posRetail.noCostCount} 筆算不出成本未計淨利`
                    : null,
                id: "pos",
              },
            ].map((p, i) => (
              <React.Fragment key={p.id}>
                {i > 0 && <div className="gcmp-div" />}
                <div className="gcmp-c">
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      marginBottom: 12,
                    }}
                  >
                    <div
                      style={{ display: "flex", alignItems: "center", gap: 6 }}
                    >
                      <div
                        style={{
                          width: 8,
                          height: 8,
                          borderRadius: 2,
                          background: p.color,
                        }}
                      />
                      <span
                        style={{
                          fontSize: 13,
                          fontWeight: 700,
                          color: "var(--t1)",
                        }}
                      >
                        {p.label}
                      </span>
                    </div>
                    <button
                      onClick={() => onNavigate(p.id)}
                      style={{
                        fontSize: 10,
                        fontWeight: 700,
                        color: p.color,
                        background: "transparent",
                        border: `1px solid ${p.color}44`,
                        borderRadius: 6,
                        padding: "3px 10px",
                        cursor: "pointer",
                      }}
                    >
                      詳細分析 →
                    </button>
                  </div>
                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: "1fr 1fr",
                      gap: 8,
                    }}
                  >
                    {[
                      {
                        l: "淨利",
                        v: fmt$(p.net),
                        c: p.net >= 0 ? p.color : "var(--dn)",
                      },
                      {
                        l: "淨利率",
                        v: fmtP(p.margin),
                        /* 三段色：達標＝平台色、未達標但仍賺＝黃、虧損＝紅。
                           原本兩段色會讓虧損顯示成黃色，與同頁下方六通路表的
                           posNetColor 自相矛盾 */
                        c:
                          p.margin >= p.target
                            ? p.color
                            : p.margin > 0
                            ? "var(--wn)"
                            : "var(--dn)",
                        sub:
                          p.margin >= p.target
                            ? `${p.targetLabel ? "高於" + p.targetLabel + " " : "超標 "}+${(
                                (p.margin - p.target) *
                                100
                              ).toFixed(1)}%`
                            : `${p.targetLabel ? "距" + p.targetLabel : "差"} ${(
                                (p.target - p.margin) *
                                100
                              ).toFixed(1)}%`,
                      },
                      { l: "營收", v: fmt$(p.rev), sub: p.note || undefined },
                      { l: "有效訂單", v: `${p.orders} 筆` },
                      {
                        l: "營業費率（快照加權）",
                        v: fmtP(p.opexRate),
                        /* 門檻與總覽同一組：30 目標／33 可接受帶／35 活動月上限 */
                        c:
                          p.rev === 0
                            ? "var(--t2)"
                            : opexBandOf(p.opexRate) === 0
                            ? p.color
                            : opexBandOf(p.opexRate) <= 2
                            ? "var(--wn)"
                            : "var(--dn)",
                      },
                      { l: "客單價 AOV", v: fmt$(p.aov) },
                    ].map((k, j) => (
                      <div
                        key={j}
                        style={{
                          background: "var(--s2)",
                          borderRadius: 8,
                          padding: "10px 12px",
                        }}
                      >
                        <div
                          style={{
                            fontSize: 10,
                            color: "var(--t4)",
                            fontWeight: 600,
                            marginBottom: 3,
                          }}
                        >
                          {k.l}
                        </div>
                        <div
                          style={{
                            fontFamily: mono,
                            fontSize: 13,
                            fontWeight: 700,
                            color: k.c || "var(--t1)",
                          }}
                        >
                          {k.v}
                        </div>
                        {k.sub && (
                          <div
                            style={{
                              fontSize: 9,
                              color: k.c || "var(--t4)",
                              marginTop: 2,
                              fontWeight: 600,
                            }}
                          >
                            {k.sub}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              </React.Fragment>
            ))}
          </div>
          {/* 門市其他五通路：各自一項（電話／Omnichat／經銷／合作／企業），沒單也列 */}
          {posS && (
            <div
              style={{
                borderTop: "1px solid var(--s3)",
                marginTop: 20,
                paddingTop: 16,
              }}
            >
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  color: "var(--t3)",
                  marginBottom: 10,
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                }}
              >
                門市其他通路（不計入「門市」，各自獨立；卡片左邊色條＝該通路在圓餅／圖例的顏色）
              </div>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))",
                  gap: 8,
                }}
              >
                {posOthers.map((c, i) => {
                  const empty = c.orders === 0;
                  const nm = chNet(c);
                  return (
                    <div
                      key={c.key}
                      style={{
                        background: "var(--s2)",
                        borderRadius: 8,
                        padding: "10px 12px",
                        opacity: empty ? 0.55 : 1,
                        borderLeft: `3px solid ${chColor(c)}`,
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                          marginBottom: 6,
                        }}
                      >
                        <span style={{ fontSize: 12, fontWeight: 700, color: "var(--t1)" }}>
                          {c.label}
                        </span>
                        <span style={{ fontSize: 10, color: "var(--t4)", fontFamily: mono }}>
                          {empty ? "無單" : `${c.orders} 筆`}
                        </span>
                      </div>
                      <div
                        style={{
                          fontFamily: mono,
                          fontSize: 15,
                          fontWeight: 700,
                          color: empty ? "var(--t4)" : "var(--t1)",
                        }}
                      >
                        {empty ? "—" : fmt$(c.rev)}
                      </div>
                      <div style={{ fontSize: 10, color: "var(--t4)", marginTop: 3 }}>
                        {empty ? (
                          "本期沒有訂單"
                        ) : (
                          <>
                            淨利 {fmt$(c.coveredNet)}・
                            <span style={{ color: c.coveredRev > 0 ? posNetColor(nm, posTarget) : "var(--t4)", fontWeight: 700 }}>
                              {c.coveredRev > 0 ? fmtP(nm) : "—"}
                            </span>
                            {c.noCostCount > 0 ? `　缺成本 ${c.noCostCount}` : ""}
                          </>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* ── 門市六通路一覽（老闆 2026-08-18：總覽要一眼看到六個通路；沒單的也列 0） ── */}
      {posS && (
        <div
          className="f2"
          style={{
            background: "var(--s1)",
            border: "1px solid var(--s3)",
            borderRadius: 16,
            padding: 24,
          }}
        >
          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              alignItems: "center",
              gap: 10,
              marginBottom: 12,
            }}
          >
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: "var(--t2)",
                display: "flex",
                alignItems: "center",
                gap: 6,
              }}
            >
              <div
                style={{ width: 8, height: 8, borderRadius: 2, background: posC }}
              />
              門市六通路 · {periodLabel}
            </div>
            <span style={{ fontSize: 11, color: "var(--t3)" }}>
              合計 {fmt$(posS.rev)} ｜ {posS.orders} 筆 ｜ 淨利率{" "}
              {posS.coveredRev > 0 ? fmtP(posS.netMargin) : "—"}
              {posS.noCostCount > 0
                ? `（${posS.noCostCount} 筆算不出成本未計淨利）`
                : ""}
            </span>
            <button
              onClick={() => onNavigate("pos")}
              style={{
                marginLeft: "auto",
                fontSize: 10,
                fontWeight: 700,
                color: posC,
                background: "transparent",
                border: `1px solid ${posC}44`,
                borderRadius: 6,
                padding: "3px 10px",
                cursor: "pointer",
              }}
            >
              詳細分析 →
            </button>
          </div>
          <div style={{ overflowX: "auto" }}>
            {/* tb-ov：手機隱藏客單／開票兩欄 */}
            <table
              className="tb-ov"
              style={{ width: "100%", borderCollapse: "collapse", minWidth: 640 }}
            >
              <thead>
                <tr>
                  {[
                    ["通路", "left"],
                    ["營收", "right"],
                    /* 佔比改用全公司當分母（老闆 2026-09-03：「佔門市」怪怪的）——
                       六通路只是都走 POS 記帳，不是一個事業體；Omnichat 佔門市 45%
                       這種數字會誤導。總覽本來就把它們當各自獨立的通路看 */
                    ["佔全公司", "right"],
                    ["筆數", "right"],
                    ["客單", "right"],
                    ["毛利率", "right"],
                    ["淨利率", "right"],
                    ["開票", "right"],
                  ].map(([h, al]) => (
                    <th
                      key={h}
                      style={{
                        padding: "8px 10px",
                        fontSize: 10,
                        fontWeight: 700,
                        color: "var(--t3)",
                        textAlign: al,
                        borderBottom: "1px solid var(--s3)",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {posCh.map((c) => {
                  const empty = c.orders === 0;
                  const gm = c.coveredRev > 0 ? c.coveredGp / c.coveredRev : null;
                  const nm = c.coveredRev > 0 ? c.coveredNet / c.coveredRev : null;
                  const share = totalRev > 0 ? c.rev / totalRev : 0;
                  const cellR = {
                    padding: "9px 10px",
                    fontSize: 12,
                    fontFamily: mono,
                    textAlign: "right",
                    borderBottom: "1px solid var(--s3)",
                    color: empty ? "var(--t4)" : "var(--t2)",
                    whiteSpace: "nowrap",
                  };
                  return (
                    <tr key={c.key} style={{ opacity: empty ? 0.55 : 1 }}>
                      <td
                        style={{
                          ...cellR,
                          textAlign: "left",
                          fontFamily: "inherit",
                          fontWeight: 700,
                          color: empty ? "var(--t4)" : "var(--t1)",
                          minWidth: 150,
                        }}
                      >
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <span>{c.label}</span>
                          {c.key === "dealer" && !empty && (
                            <span
                              style={{
                                fontSize: 9,
                                padding: "1px 5px",
                                borderRadius: 4,
                                border: `1px solid ${posC}55`,
                                color: posC,
                                fontWeight: 700,
                              }}
                            >
                              老闆關係單
                            </span>
                          )}
                        </div>
                        {/* 營收比較條：長度相對最大通路，不是絕對佔比 */}
                        <div
                          style={{
                            height: 4,
                            borderRadius: 99,
                            background: "var(--s3)",
                            marginTop: 5,
                            overflow: "hidden",
                            maxWidth: 160,
                          }}
                        >
                          <div
                            style={{
                              width: `${
                                posChMax > 0 ? (c.rev / posChMax) * 100 : 0
                              }%`,
                              height: "100%",
                              background: posC,
                              opacity: 0.85,
                            }}
                          />
                        </div>
                      </td>
                      <td style={{ ...cellR, fontWeight: empty ? 400 : 700 }}>
                        {empty ? "—" : fmt$(c.rev)}
                      </td>
                      <td style={cellR}>{empty ? "—" : fmtP(share)}</td>
                      <td style={cellR}>{empty ? "0" : c.orders}</td>
                      <td style={cellR}>{empty ? "—" : fmt$(c.rev / c.orders)}</td>
                      <td
                        style={{
                          ...cellR,
                          fontWeight: 700,
                          color: gm === null ? "var(--t4)" : posGmColor(gm),
                        }}
                      >
                        {gm === null ? "—" : fmtP(gm)}
                      </td>
                      <td
                        style={{
                          ...cellR,
                          fontWeight: 700,
                          color: nm === null ? "var(--t4)" : posNetColor(nm, posTarget),
                        }}
                      >
                        {nm === null ? "—" : fmtP(nm)}
                        {!empty && c.noCostCount > 0 && (
                          <span
                            style={{ fontSize: 9, color: "var(--wn)", marginLeft: 4 }}
                            title={`${c.noCostCount} 筆算不出成本，未計入此通路淨利率`}
                          >
                            缺{c.noCostCount}
                          </span>
                        )}
                      </td>
                      <td style={cellR}>
                        {empty || c.rev === 0 ? "—" : fmtP(c.invoiceRev / c.rev)}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ fontSize: 10, color: "var(--t4)", marginTop: 10, lineHeight: 1.6 }}>
            門市無平台抽成；毛利率／淨利率只算成本齊全的單；稅只課有開發票的單。淨利率綠＝達門市目標
            {fmtP(posTarget)}（六通路同一把尺）。
            {posCh.some((c) => c.orders === 0) &&
              "　灰色通路＝本期還沒有訂單（員工開單時選對應付款方式就會出現）。"}
          </div>
        </div>
      )}

      {/* ── 異常警示 ── */}
      {alerts.length > 0 && (
        <div
          className="f2"
          style={{
            background: "var(--s1)",
            border: "1px solid var(--s3)",
            borderRadius: 16,
            padding: "18px 24px",
          }}
        >
          <div
            style={{
              fontSize: 11,
              fontWeight: 700,
              color: "var(--t3)",
              marginBottom: 12,
              letterSpacing: "0.06em",
              textTransform: "uppercase",
              display: "flex",
              alignItems: "center",
              gap: 6,
            }}
          >
            <AlertTriangle size={13} color="var(--wn)" /> 需要注意
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {alerts.map((a, i) => {
              const col =
                a.level === "error"
                  ? "var(--dn)"
                  : a.level === "warn"
                  ? "var(--wn)"
                  : "var(--blue)";
              return (
                <div
                  key={i}
                  style={{
                    display: "flex",
                    alignItems: "flex-start",
                    gap: 10,
                    padding: "10px 14px",
                    background: "var(--s2)",
                    border: `1px solid var(--s3)`,
                    borderRadius: 10,
                    borderLeft: `3px solid ${col}`,
                  }}
                >
                  <div
                    style={{
                      fontSize: 10,
                      fontWeight: 800,
                      color: col,
                      background: "var(--s3)",
                      padding: "2px 8px",
                      borderRadius: 4,
                      flexShrink: 0,
                      marginTop: 1,
                    }}
                  >
                    {a.platform}
                  </div>
                  <div
                    style={{
                      fontSize: 12,
                      color: "var(--t2)",
                      fontWeight: 500,
                      lineHeight: 1.5,
                    }}
                  >
                    {a.msg}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* ── 月度趨勢 ── */}
      {trendData.length >= 1 && (
        <div
          className="f3"
          style={{
            background: "var(--s1)",
            border: "1px solid var(--s3)",
            borderRadius: 16,
            padding: 24,
          }}
        >
          <div
            style={{
              fontSize: 12,
              fontWeight: 700,
              color: "var(--t2)",
              marginBottom: 4,
              display: "flex",
              alignItems: "center",
              gap: 6,
            }}
          >
            <TrendingUp size={14} color="var(--t3)" /> 月度營收與淨利趨勢
          </div>
          <div
            className="trend-legend"
            style={{
              fontSize: 11,
              color: "var(--t3)",
              marginBottom: 16,
              display: "flex",
              gap: 16,
            }}
          >
            <span
              style={{ display: "inline-flex", alignItems: "center", gap: 4 }}
            >
              <span
                style={{
                  display: "inline-block",
                  width: 8,
                  height: 3,
                  borderRadius: 2,
                  background: greenC,
                }}
              />
              官網
            </span>
            <span
              style={{ display: "inline-flex", alignItems: "center", gap: 4 }}
            >
              <span
                style={{
                  display: "inline-block",
                  width: 8,
                  height: 3,
                  borderRadius: 2,
                  background: spC,
                }}
              />
              蝦皮
            </span>
            <span
              style={{ display: "inline-flex", alignItems: "center", gap: 4 }}
            >
              <span
                style={{
                  display: "inline-block",
                  width: 8,
                  height: 3,
                  borderRadius: 2,
                  background: posC,
                }}
              />
              門市（現場零售）
            </span>
            <span
              style={{ display: "inline-flex", alignItems: "center", gap: 4 }}
            >
              <span
                style={{
                  display: "inline-block",
                  width: 8,
                  height: 3,
                  borderRadius: 2,
                  background: posC,
                  opacity: 0.45,
                }}
              />
              門市其他通路
            </span>
            {showNetLine && (
              <span
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 4,
                }}
              >
                <span
                  style={{
                    display: "inline-block",
                    width: 8,
                    height: 3,
                    borderRadius: 2,
                    background: goldC,
                  }}
                />
                淨利
              </span>
            )}
          </div>
          <ResponsiveContainer width="100%" height={180}>
            <ComposedChart
              data={trendData}
              margin={{ top: 4, right: 16, left: -8, bottom: 4 }}
            >
              <CartesianGrid
                strokeDasharray="3 3"
                stroke={gridC}
                vertical={false}
              />
              <XAxis
                dataKey="label"
                tick={{
                  fontSize: 10,
                  fill: "var(--t3)",
                  fontFamily: mono,
                  fontWeight: 600,
                }}
                axisLine={{ stroke: gridC }}
                tickLine={false}
                dy={4}
              />
              <YAxis
                tick={{ fontSize: 9, fill: "var(--t3)", fontFamily: mono }}
                axisLine={false}
                tickLine={false}
                tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`}
              />
              <RTooltip
                content={({ active, payload, label }) => {
                  if (!active || !payload?.length) return null;
                  const bars = payload.filter((e) => e.dataKey !== "net");
                  const netE = payload.find((e) => e.dataKey === "net");
                  const total = bars.reduce((s, e) => s + (e.value || 0), 0);
                  return (
                    <div
                      style={{
                        background: "var(--s1)",
                        border: "1px solid var(--s3)",
                        borderRadius: 10,
                        padding: "10px 14px",
                        boxShadow: "0 4px 20px rgba(0,0,0,0.12)",
                        minWidth: 180,
                      }}
                    >
                      <div
                        style={{
                          fontSize: 10,
                          fontWeight: 700,
                          color: "var(--t3)",
                          marginBottom: 8,
                          fontFamily: mono,
                        }}
                      >
                        {label}
                      </div>
                      {bars.map((e, i) => {
                        const pct =
                          total > 0
                            ? (((e.value || 0) / total) * 100).toFixed(1)
                            : "0.0";
                        return (
                          <div
                            key={i}
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: 8,
                              padding: "3px 0",
                            }}
                          >
                            <div
                              style={{
                                width: 8,
                                height: 8,
                                borderRadius: 2,
                                background: e.color,
                                flexShrink: 0,
                              }}
                            />
                            <span
                              style={{
                                fontSize: 11,
                                color: "var(--t2)",
                                fontWeight: 600,
                                minWidth: 52,
                              }}
                            >
                              {e.name}
                            </span>
                            <span
                              style={{
                                fontSize: 10,
                                fontWeight: 700,
                                color: e.color,
                                fontFamily: mono,
                                background: "var(--s3)",
                                padding: "1px 6px",
                                borderRadius: 4,
                                width: 46,
                                textAlign: "center",
                              }}
                            >
                              {pct}%
                            </span>
                            <span
                              style={{
                                fontSize: 12,
                                fontWeight: 800,
                                color: "var(--t1)",
                                fontFamily: mono,
                                textAlign: "right",
                                flex: 1,
                              }}
                            >
                              {fmt$(e.value)}
                            </span>
                          </div>
                        );
                      })}
                      <div
                        style={{
                          borderTop: "1px solid var(--s3)",
                          marginTop: 8,
                          paddingTop: 6,
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <span
                          style={{
                            fontSize: 10,
                            color: "var(--t4)",
                            fontWeight: 600,
                          }}
                        >
                          合計營收
                        </span>
                        <span
                          style={{
                            fontSize: 13,
                            fontWeight: 800,
                            fontFamily: mono,
                            color: "var(--t1)",
                          }}
                        >
                          {fmt$(total)}
                        </span>
                      </div>
                      {netE && netE.value !== undefined && (
                        <div
                          style={{
                            marginTop: 4,
                            display: "flex",
                            justifyContent: "space-between",
                            alignItems: "center",
                          }}
                        >
                          <span
                            style={{
                              fontSize: 10,
                              color: goldC,
                              fontWeight: 700,
                            }}
                          >
                            淨利
                          </span>
                          <span
                            style={{
                              fontSize: 12,
                              fontWeight: 800,
                              fontFamily: mono,
                              color: netE.value >= 0 ? goldC : "var(--dn)",
                            }}
                          >
                            {fmt$(netE.value)}
                          </span>
                        </div>
                      )}
                    </div>
                  );
                }}
              />
              <Bar
                dataKey="slRev"
                name="官網"
                fill={greenC}
                opacity={0.85}
                radius={[0, 0, 0, 0]}
                maxBarSize={32}
                stackId="a"
              />
              <Bar
                dataKey="spRev"
                name="蝦皮"
                fill={spC}
                opacity={0.85}
                radius={[0, 0, 0, 0]}
                maxBarSize={32}
                stackId="a"
              />
              <Bar
                dataKey="posRev"
                name="門市零售"
                fill={posC}
                opacity={0.9}
                radius={[0, 0, 0, 0]}
                maxBarSize={32}
                stackId="a"
              />
              <Bar
                dataKey="posOtherRev"
                name="門市其他"
                fill={posC}
                opacity={0.45}
                radius={[3, 3, 0, 0]}
                maxBarSize={32}
                stackId="a"
              />
              {showNetLine && (
                <Line
                  type="monotone"
                  dataKey="net"
                  name="淨利"
                  stroke={goldC}
                  strokeWidth={2}
                  dot={{ r: 2, fill: goldC, strokeWidth: 0 }}
                  activeDot={{ r: 4 }}
                  connectNulls
                />
              )}
            </ComposedChart>
          </ResponsiveContainer>
        </div>
      )}

      {/* ── 跨平台商品排行 ── */}
      {crossProductRank.length > 0 && (
        <div
          className="f4"
          style={{
            background: "var(--s1)",
            border: "1px solid var(--s3)",
            borderRadius: 16,
            padding: 24,
          }}
        >
          <div
            style={{
              fontSize: 12,
              fontWeight: 700,
              color: "var(--t2)",
              display: "flex",
              alignItems: "center",
              gap: 6,
              marginBottom: 4,
            }}
          >
            <Package size={14} color="var(--t3)" /> 官網 × 蝦皮 銷售排行
          </div>
          <div style={{ fontSize: 11, color: "var(--t3)", marginBottom: 16 }}>
            綠 = 官網　橘 = 蝦皮（門市品名格式不同、未納入）
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {crossProductRank.map((p, i) => {
              const total = p.slQty + p.spQty;
              const slPct = total > 0 ? p.slQty / total : 0;
              const maxTotal =
                crossProductRank[0].slQty + crossProductRank[0].spQty;
              return (
                <div
                  key={p.name}
                  style={{ display: "flex", alignItems: "center", gap: 12 }}
                >
                  <div
                    style={{
                      fontSize: 11,
                      fontWeight: 700,
                      color: "var(--t4)",
                      fontFamily: mono,
                      width: 16,
                      textAlign: "right",
                      flexShrink: 0,
                    }}
                  >
                    {i + 1}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div
                      style={{
                        fontSize: 12,
                        fontWeight: 600,
                        color: "var(--t1)",
                        marginBottom: 5,
                        whiteSpace: "nowrap",
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                      }}
                    >
                      {p.name}
                    </div>
                    <div
                      style={{
                        height: 7,
                        borderRadius: 99,
                        background: "var(--s3)",
                        overflow: "hidden",
                        width: `${(total / maxTotal) * 100}%`,
                      }}
                    >
                      <div style={{ height: "100%", display: "flex" }}>
                        <div
                          style={{
                            width: `${slPct * 100}%`,
                            background: greenC,
                          }}
                        />
                        <div
                          style={{ flex: 1, background: spC, opacity: 0.8 }}
                        />
                      </div>
                    </div>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      gap: 6,
                      fontSize: 11,
                      fontFamily: mono,
                      flexShrink: 0,
                      alignItems: "center",
                    }}
                  >
                    <span style={{ color: greenC, fontWeight: 700 }}>
                      {p.slQty}
                    </span>
                    <span style={{ color: "var(--s4)" }}>+</span>
                    <span style={{ color: spC, fontWeight: 700 }}>
                      {p.spQty}
                    </span>
                    <span
                      style={{
                        color: "var(--t3)",
                        fontWeight: 600,
                        minWidth: 36,
                      }}
                    >
                      = {total}
                    </span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

/* ─── Monthly Expense Panel（分潤等按月費用）─────────────────── */
function MonthlyExpensePanel({
  title,
  icon,
  color = "var(--purple)",
  values,
  onUpdate,
  selYear,
  selMonth,
  range,
  hint,
}) {
  const key = commKey(selYear, selMonth);
  const isAggregated =
    selMonth === "All" || selYear === "All" || selYear === "Custom";
  const aggregatedVal = useMemo(
    () =>
      isAggregated ? periodExpense(values, selYear, selMonth, range) : null,
    [values, selYear, selMonth, range, isAggregated]
  );

  const [local, setLocal] = useState(String(values[key] ?? ""));
  const [focused, setFocused] = useState(false);
  useEffect(() => {
    /* 輸入中不被遠端同步覆蓋草稿（與 CostInput/FpInput 同一套防呆） */
    if (!focused) setLocal(String(values[key] ?? ""));
  }, [key, values, focused]);
  const handleBlur = () => {
    setFocused(false);
    const n = parseFloat(local);
    onUpdate(key, isNaN(n) ? "" : n);
  };
  const hasVal =
    values[key] !== undefined && values[key] !== "" && Number(values[key]) > 0;
  const label =
    selYear === "Custom"
      ? "自訂區間"
      : selYear === "All"
      ? "歷年"
      : selMonth === "All"
      ? `${selYear}`
      : `${selYear}/${selMonth}`;

  return (
    <div
      style={{
        background: "var(--s1)",
        border: "1px solid var(--s3)",
        borderRadius: 14,
        padding: 16,
      }}
    >
      <div
        style={{
          fontSize: 11,
          fontWeight: 700,
          color: "var(--t3)",
          display: "flex",
          alignItems: "center",
          gap: 6,
          marginBottom: 10,
        }}
      >
        {icon} {title}
      </div>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: 8,
        }}
      >
        <span
          style={{
            fontFamily: mono,
            fontSize: 11,
            fontWeight: 700,
            color,
            background: "var(--s2)",
            border: "1px solid var(--s3)",
            padding: "3px 8px",
            borderRadius: 5,
          }}
        >
          {label}
        </span>
        {!isAggregated && hasVal && (
          <button
            onClick={() => {
              setLocal("");
              onUpdate(key, "");
            }}
            aria-label={`清除此期間${title}`}
            style={{
              border: "none",
              background: "none",
              color: "var(--t4)",
              cursor: "pointer",
              display: "flex",
            }}
          >
            <X size={13} />
          </button>
        )}
      </div>
      {isAggregated ? (
        <div
          style={{
            background: "var(--s2)",
            borderRadius: 8,
            padding: "10px 12px",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <span style={{ fontSize: 10, color: "var(--t4)", fontWeight: 600 }}>
            各月合計
          </span>
          <span
            style={{
              fontFamily: mono,
              fontSize: 15,
              fontWeight: 700,
              color: aggregatedVal > 0 ? color : "var(--t3)",
            }}
          >
            {fmt$(aggregatedVal)}
          </span>
        </div>
      ) : (
        <div style={{ position: "relative" }}>
          <span
            style={{
              position: "absolute",
              left: 10,
              top: "50%",
              transform: "translateY(-50%)",
              fontSize: 10,
              fontWeight: 700,
              color: "var(--t3)",
              fontFamily: mono,
              pointerEvents: "none",
            }}
          >
            NT$
          </span>
          <input
            type="number"
            min="0"
            value={local}
            placeholder="0"
            aria-label={`此期間${title}`}
            onChange={(e) => setLocal(e.target.value)}
            onFocus={() => setFocused(true)}
            onBlur={handleBlur}
            style={{
              ...inp,
              width: "100%",
              textAlign: "right",
              paddingLeft: 36,
              borderColor: focused ? color : "var(--s3)",
            }}
          />
        </div>
      )}
      <p
        style={{
          fontSize: 10,
          color: "var(--t4)",
          marginTop: 6,
          lineHeight: 1.6,
        }}
      >
        {isAggregated
          ? selYear === "Custom"
            ? "自訂區間以整月合計；要編輯請於上方年份選單改選單一月份"
            : "各月合計；要編輯請於上方選單改選單一月份"
          : hint}
      </p>
    </div>
  );
}

/* 費率參數解析：快照有值用快照、否則用側欄現值。三個 OrderFin 共用同一條規則，
   免得日後改快照語意只改到其中一個。回傳百分比數值（要小數自己 /100） */
const fpPct = (ofp, fp, k) =>
  (ofp?.[k] != null ? Number(ofp[k]) : parseFloat(fp?.[k])) || 0;

/* ─── 單筆訂單損益計算（畫面與月度彙總共用，勿分岔） ─────────── */
const slOrderFin = (order, fp, costsMap) => {
  const ofp = order.snapshotFeeParams;
  const sfr = fpPct(ofp, fp, "platformFeeRate") / 100;
  const oer = fpPct(ofp, fp, "opExpense") / 100;
  const dtr = fpPct(ofp, fp, "tax") / 100;
  /* 金流／物流：查表命中才用表值；比不到不再無聲套預設——回傳 payKnown/dlvKnown 讓明細標 ⚠ */
  const prHit = slPayRate(order.paymentMethod);
  const payKnown = prHit !== null;
  const pr = prHit || { rate: 0.022, flat: 0 };
  const pf = order.revenue * pr.rate + pr.flat;
  let sc2;
  let dlvKnown = true;
  if (slIntl(order.deliveryMethod)) sc2 = numOrZero(order.shippingIncome);
  else {
    const hit = slShipRate(order.deliveryMethod);
    if (hit !== null) sc2 = hit;
    else {
      sc2 = SL_SHIP_FALLBACK;
      dlvKnown = false;
    }
  }
  const plf = order.revenue * sfr;
  let oc = 0;
  (order.items || []).forEach((item) => {
    const cv =
      Object.prototype.hasOwnProperty.call(item, "snapshotCost") &&
      item.snapshotCost !== null
        ? Number(item.snapshotCost) || 0
        : Number(costsMap[item.key]) || 0;
    oc += cv * item.qty;
  });
  const cm = order.revenue - oc - pf - sc2 - plf;
  const tax = order.isTaxExempt ? 0 : order.revenue * dtr;
  const opx = order.revenue * oer;
  return { pf, sc2, plf, oc, cm, tax, opx, net: cm - opx - tax, payKnown, dlvKnown };
};

const spOrderFin = (order, fp, costsMap) => {
  const st = safeText(order.status),
    rf = safeText(order.refundStatus);
  const isCanc = st.includes("不成立") || st.includes("取消");
  const isRef =
    !isCanc && (rf !== "" || (st.includes("退貨") && !st.includes("已完成")));
  const gross = numOrZero(order.grossPrice);
  const ofp = order.snapshotFeeParams;
  const opEx = fpPct(ofp, fp, "opExpense");
  const tx = fpPct(ofp, fp, "tax");
  let oCost = 0;
  (order.items || []).forEach((item) => {
    const ic =
      Object.prototype.hasOwnProperty.call(item, "snapshotCost") &&
      item.snapshotCost !== null
        ? Number(item.snapshotCost) || 0
        : Number(costsMap[item.key]) || 0;
    oCost += ic * (item.qty || 1);
  });
  /* 賣場優惠券：報表的「商品總價」已經是扣掉賣場券後的金額（2,693 筆實測 499/509 分毫不差），
     這裡不能再扣一次；voucher 只回傳供券率統計。賣家蝦幣回饋券是賣家實付成本，併入通路費 */
  const voucher = numOrZero(order.sellerVoucher);
  const fee =
    numOrZero(order.exactOrderFee) + numOrZero(order.sellerCoinCashback);
  const net = gross - fee;
  const gp = net - oCost;
  const opAmt = gross * (opEx / 100);
  const taxBase = numOrZero(order.buyerTotal) || gross;
  const txAmt = taxBase * (tx / 100);
  return {
    isCanc,
    isRef,
    gross,
    voucher,
    fee,
    oCost,
    net,
    gp,
    opAmt,
    txAmt,
    finalNet: gp - opAmt - txAmt,
    opEx,
    tx,
  };
};

/* ─── 門市 POS 單筆損益 ────────────────────────────────────────
   門市無平台抽成：毛利＝營收−商品成本；營業費沿用內部 %；
   稅只對「有開發票」的單課徵（老闆 2026-08-18 定案）。         */
/* ── 泛稱品項成本率（門市專用）──────────────────────────────────
   2026-08-18 門市商品目錄上線前，員工只點得到「茶葉」「茶葉禮盒」這種泛稱，
   沒有規格可回推單位成本，而且多數列是「數量 1、單價＝整筆金額」（少數顛倒）。
   這類鍵改用成本率：單位成本＝該列單價 × 率，總成本＝單價 × 數量 × 率
   ＝營收 × 率，與件數無關，單價/數量顛倒的列也算得對。
   posRatios: { costKey: 百分比 }。
   優先序：配方／手填成本永遠優先，率只在兩者都沒有時才生效——之後補了真成本，
   率自動失效（不必刪），畫面與帳上永遠是同一個數字。 */
const posRatioUnit = (item, costsMap, ratios) => {
  /* 有真成本就不套率 */
  if (Number(costsMap?.[item?.key]) > 0) return null;
  const r = Number(ratios?.[item?.key]);
  if (!(r > 0)) return null;
  const price = Number(item?.price) || 0;
  /* 單價 0 的列（結帳價全 0、營收靠等分攤）估不出來，仍算缺成本 */
  return price > 0 ? price * (r / 100) : null;
};
const posItemUnit = (item, costsMap, ratios) => {
  const real = Number(costsMap?.[item.key]) || 0;
  if (real > 0) return real;
  const rv = posRatioUnit(item, costsMap, ratios);
  return rv !== null ? rv : 0;
};
const posOrderFin = (order, fp, costsMap, ratios) => {
  const st = safeText(order.status);
  const isCanc = st.includes("取消") || st.includes("退款") || st.includes("已退");
  const gross = numOrZero(order.revenue);
  const ofp = order.snapshotFeeParams;
  const opEx = fpPct(ofp, fp, "opExpense");
  const tx = fpPct(ofp, fp, "tax");
  let oCost = 0;
  /* 估算成本金額（走成本率的部分）：讓畫面能誠實標示「這個數字含多少估的」 */
  let est = 0;
  /* 沒有商品明細（交易有、明細沒對到）＝算不出成本，一律視為缺成本 */
  let missCost = !(order.items || []).length;
  (order.items || []).forEach((item) => {
    const has =
      Object.prototype.hasOwnProperty.call(item, "snapshotCost") &&
      item.snapshotCost !== null;
    const real = Number(costsMap[item.key]) || 0;
    const rUnit = has || real > 0 ? null : posRatioUnit(item, costsMap, ratios);
    const unit = has
      ? Number(item.snapshotCost) || 0
      : real > 0
      ? real
      : rUnit !== null
      ? rUnit
      : 0;
    if (!has && real <= 0 && rUnit === null) missCost = true;
    /* 估算身分：未鎖＝正在走率；已鎖＝快照標了 snapshotEst（鎖定當下就是估算值），
       兩種都要算進 est，Hero 的「含估算」標籤與建議率的排除才不會在鎖定後失效 */
    if ((has && item.snapshotEst === true) || (!has && real <= 0 && rUnit !== null))
      est += unit * (item.qty || 1);
    oCost += unit * (item.qty || 1);
  });
  const gp = gross - oCost;
  const opAmt = gross * (opEx / 100);
  const txAmt = order.hasInvoice ? gross * (tx / 100) : 0;
  return {
    isCanc,
    isTest: gross > 0 && gross <= POS_TEST_MAX,
    gross,
    oCost,
    estCost: est,
    hasEst: est > 0,
    gp,
    opAmt,
    txAmt,
    missCost,
    finalNet: gp - opAmt - txAmt,
    opEx,
    tx,
  };
};

/* ─── 門市 POS：商品名＋選項 → 配方（自動對應規則引擎） ────────
   三條鐵則（見 利潤決策中心\門市POS成本對應-交接.md）：
   ① 名稱含「公版」→ 公版茶葉分類  ② 含「禮盒茶」→ 禮盒專用茶葉
   ③ 其餘散茶 → 官網茶葉分類。規格 二兩=75g／四兩=150g。      */
const POS_SHELL_RECIPES = {
  品牌茶葉禮盒: [
    ["禮盒體(品牌/福悠然)", 1], ["內襯(品牌/福悠然)", 1], ["提袋(品牌禮盒)", 1],
    ["吊牌(品牌)", 1], ["介紹卡(品牌)", 1], ["沖泡小卡", 1], ["茶罐(單罐)", 2],
  ],
  福悠然茶葉禮盒: [
    ["禮盒體(品牌/福悠然)", 1], ["內襯(品牌/福悠然)", 1], ["提袋(福悠然)", 1],
    ["吊牌(福悠然)", 1], ["介紹卡(福悠然)", 1], ["沖泡小卡", 1], ["茶罐(單罐)", 2],
  ],
  "拾光茶信禮盒-含原葉茶包": [
    ["禮盒體(茶信)", 1], ["提袋(茶信)", 1], ["介紹卡(茶信/朝霞)", 1],
    ["沖泡小卡", 1], ["茶罐(單罐)", 1],
    ["悠悠紅玉立體茶包(1包·含茶葉)", 2], ["奶香金萱立體茶包(1包·含茶葉)", 2],
    ["暮暮觀音立體茶包(1包·含茶葉)", 2], ["醇翠烏龍立體茶包(1包·含茶葉)", 2],
  ],
  暖心茶山禮盒: [
    ["禮盒體(暖心茶山小)", 1], ["提袋(暖心茶山小)", 1], ["介紹卡(暖心茶山)", 1],
    ["沖泡小卡", 1], ["茶罐(單罐)", 2],
  ],
  朝霞映春禮盒: [
    ["禮盒體(朝霞映春·含內襯)", 1], ["過年提袋(朝霞)", 1], ["紅茶外盒+標籤(朝霞)", 1],
    ["介紹卡(茶信/朝霞)", 1], ["沖泡小卡", 1], ["茶罐(單罐)", 1],
  ],
  "暖心茶山禮盒(大禹嶺茶包)": [
    ["禮盒體(暖心茶山小)", 1], ["提袋(暖心茶山小)", 1], ["介紹卡(暖心茶山)", 1],
    ["沖泡小卡", 1], ["茶罐(單罐)", 2], ["大禹嶺茶包(1包)", 10],
  ],
  "暖心茶山禮盒(福壽山、焙烏龍、紅茶包)": [
    ["禮盒體(暖心茶山小)", 1], ["提袋(暖心茶山小)", 1], ["介紹卡(暖心茶山)", 1],
    ["沖泡小卡", 1], ["茶罐(單罐)", 2],
    ["福壽山茶包(1包)", 4], ["焙烏龍茶包(1包)", 3], ["華崗紅茶茶包(1包·含茶葉)", 3],
  ],
};
const POS_GIFT_TEA = {
  福壽山: "福壽山150g(禮盒款)",
  大禹嶺: "大禹嶺150g(禮盒款)",
  梨山: "翠峰150g(禮盒款)",
  阿里山: "阿里山150g(禮盒款)",
};
const POS_COLD_TEA = {
  悠悠紅玉: "悠悠紅玉立體茶包(1包·含茶葉)",
  奶香金萱: "奶香金萱立體茶包(1包·含茶葉)",
  暮暮觀音: "暮暮觀音立體茶包(1包·含茶葉)",
  醇翠烏龍: "醇翠烏龍立體茶包(1包·含茶葉)",
};
/* 規格字串 → 克數（二兩=75、四兩=150；也吃「75g」「150g」直寫） */
const posGramOf = (option) => {
  const o = String(option || "");
  if (o.includes("二兩")) return 75;
  if (o.includes("四兩")) return 150;
  const m = o.match(/(\d+)\s*g/i);
  if (m) return parseInt(m[1], 10);
  return null;
};
const buildPosRecipe = (name, option, components) => {
  const compEntries = Object.entries(components || {});
  const byExact = (nm) => compEntries.find(([, c]) => c.name === nm)?.[0] || null;
  const lines = (pairs) => {
    const out = [];
    for (const [nm, q] of pairs) {
      const id = byExact(nm);
      if (!id) return null;
      out.push({ compId: id, qty: q });
    }
    return out;
  };
  const n = String(name || "").trim();
  const opt = String(option || "").trim();

  /* 禮盒包裝／茶葉禮盒加購（改名前後都吃） */
  if (n.includes("禮盒包裝") || n.includes("茶葉禮盒加購")) {
    const pairs = POS_SHELL_RECIPES[opt];
    return pairs ? lines(pairs) : null;
  }
  /* 禮盒茶（150g罐）：選項＝款式 */
  if (n.includes("禮盒茶")) {
    const nm = POS_GIFT_TEA[opt];
    const id = nm ? byExact(nm) : null;
    return id ? [{ compId: id, qty: 1 }] : null;
  }
  /* 冷泡茶：空罐＋立體茶包 */
  if (n.includes("冷泡茶")) {
    const can = byExact("冷泡茶罐(空罐)");
    const teaNm = POS_COLD_TEA[opt];
    const tea = teaNm ? byExact(teaNm) : null;
    return can && tea ? [{ compId: can, qty: 1 }, { compId: tea, qty: 1 }] : null;
  }
  /* 30 入盒裝茶包：「XX茶包｜經典好茶」＝平面單包×30＋對應外盒 */
  if (/茶包｜經典好茶/.test(n)) {
    const base = n.replace(/｜經典好茶.*$/, "").trim();
    const pack = compEntries.find(
      ([, c]) => c.cat === "茶包單包" && c.name.startsWith(base)
    );
    const box = compEntries.find(
      ([, c]) => c.cat === "外盒" && c.name.startsWith(base) && c.name.includes("外盒")
    );
    if (pack)
      return box
        ? [{ compId: pack[0], qty: 30 }, { compId: box[0], qty: 1 }]
        : [{ compId: pack[0], qty: 30 }];
    return null;
  }
  /* 公版散茶：完全同名優先；否則用「公版{品項}-」前綴＋規格克數比對
     （POS 可能只打「公版華崗」，庫裡是「公版華崗-邱高翠｜75g」） */
  if (n.includes("公版")) {
    const exact = byExact(n);
    if (exact) return [{ compId: exact, qty: 1 }];
    const base = n.replace(/｜.*$/, "").trim();
    const g = posGramOf(opt) || posGramOf(n);
    const pool = compEntries.filter(
      ([, c]) =>
        c.cat === "公版茶葉" &&
        (c.name === base ||
          c.name.startsWith(`${base}-`) ||
          c.name.startsWith(`${base}｜`))
    );
    if (!pool.length) return null;
    const hit = g ? pool.find(([, c]) => c.name.endsWith(`｜${g}g`)) : null;
    const pick = hit || (pool.length === 1 ? pool[0] : null);
    return pick ? [{ compId: pick[0], qty: 1 }] : null;
  }
  /* 品牌茶罐 */
  if (n.includes("茶罐")) {
    const id = byExact("茶罐(單罐)");
    return id ? [{ compId: id, qty: 1 }] : null;
  }
  /* 單包／大份量茶包：選項＝茶包款式 */
  if (n.includes("單包茶包") || n.includes("大份量茶包")) {
    const id =
      compEntries.find(
        ([, c]) => c.cat === "茶包單包" && opt && c.name.startsWith(opt)
      )?.[0] || null;
    return id ? [{ compId: id, qty: 1 }] : null;
  }
  /* 一般散茶：用規格克數找官網茶葉組件（排除禮盒款／公版） */
  const g = posGramOf(opt);
  if (g) {
    const base = n.replace(/^品牌｜/, "").replace(/｜.*$/, "").trim();
    const hit = compEntries.find(
      ([, c]) =>
        c.cat === "茶葉" &&
        c.name === `${base}${g}g`
    );
    if (hit) return [{ compId: hit[0], qty: 1 }];
    /* 寬鬆比對只在「候選唯一」時採用；同茶區有多支（阿里山／阿里山金萱／阿里山瑞里）
       就回 null 讓它進「無成本資料」清單，由老闆在成本資料庫手動指定——寧可漏配不可配錯 */
    const loose = compEntries.filter(
      ([, c]) =>
        c.cat === "茶葉" &&
        c.name.endsWith(`${g}g`) &&
        (c.name.startsWith(base) || base.includes(c.name.replace(`${g}g`, "")))
    );
    if (loose.length === 1) return [{ compId: loose[0][0], qty: 1 }];
  }
  return null;
};

/* ─── 期間比較（環比／同比） ─────────────────────────────────── */
const CmpVal = ({ label, cur, prev }) => {
  let txt, c;
  if (prev > 0) {
    const d = (cur - prev) / prev;
    txt = `${d >= 0 ? "+" : ""}${(d * 100).toFixed(1)}%`;
    c = d > 0 ? "var(--up)" : d < 0 ? "var(--dn)" : "var(--t3)";
  } else {
    const d = cur - prev;
    txt = `${d >= 0 ? "+" : "−"}${fmt$(Math.abs(d))}`;
    c = d > 0 ? "var(--up)" : d < 0 ? "var(--dn)" : "var(--t3)";
  }
  return (
    <span style={{ display: "inline-flex", gap: 4, alignItems: "center" }}>
      <span style={{ color: "var(--t3)" }}>{label}</span>
      <span style={{ fontFamily: mono, fontWeight: 700, color: c }}>
        {txt}
      </span>
    </span>
  );
};

function PeriodCompare({ monthly, sY, sM }) {
  const data = useMemo(() => {
    if (!monthly || sY === "All" || sY === "Custom") return null;
    const sumYear = (y) => {
      let rev = 0,
        net = 0,
        has = false;
      Object.entries(monthly).forEach(([k, v]) => {
        if (k.startsWith(y + "-")) {
          rev += v.rev;
          net += v.net;
          has = true;
        }
      });
      return has ? { rev, net } : null;
    };
    /* 節慶型電商季節性強：同比（去年同期）優先，環比僅供參考 */
    let cur;
    let missYoY = null;
    const groups = [];
    if (sM !== "All") {
      cur = monthly[`${sY}-${sM}`];
      const mNum = Number(sM);
      const pmKey =
        mNum === 1
          ? `${Number(sY) - 1}-12`
          : `${sY}-${String(mNum - 1).padStart(2, "0")}`;
      const pyKey = `${Number(sY) - 1}-${sM}`;
      if (monthly[pyKey])
        groups.push({
          label: `同比 ${pyKey.replace("-", "/")}`,
          prev: monthly[pyKey],
          primary: true,
        });
      else missYoY = `同比 ${pyKey.replace("-", "/")}`;
      if (monthly[pmKey])
        groups.push({
          label: `環比 ${pmKey.replace("-", "/")}`,
          prev: monthly[pmKey],
        });
    } else {
      cur = sumYear(sY);
      const py = sumYear(String(Number(sY) - 1));
      if (py)
        groups.push({
          label: `同比 ${Number(sY) - 1}年`,
          prev: py,
          primary: true,
        });
      else missYoY = `同比 ${Number(sY) - 1}年`;
    }
    if (!cur || (!groups.length && !missYoY)) return null;
    return { cur, groups, missYoY };
  }, [monthly, sY, sM]);
  if (!data) return null;
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 12 }}>
      {data.groups.map((g) => (
        <div
          key={g.label}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 10,
            padding: "5px 12px",
            borderRadius: 8,
            background: g.primary ? "var(--accent-dim)" : "var(--s2)",
            border: `1px solid ${g.primary ? "var(--accent-bdr)" : "var(--s3)"}`,
            fontSize: 11,
            fontWeight: 600,
          }}
        >
          <span
            style={{
              color: g.primary ? "var(--accent-text)" : "var(--t2)",
              fontWeight: 700,
            }}
          >
            {g.label}
          </span>
          <CmpVal label="營收" cur={data.cur.rev} prev={g.prev.rev} />
          <CmpVal label="淨利" cur={data.cur.net} prev={g.prev.net} />
        </div>
      ))}
      {data.missYoY && (
        <div
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 8,
            padding: "5px 12px",
            borderRadius: 8,
            background: "var(--s2)",
            border: "1px dashed var(--s4)",
            fontSize: 11,
            fontWeight: 600,
          }}
        >
          <span style={{ color: "var(--t3)", fontWeight: 700 }}>
            {data.missYoY}
          </span>
          <span style={{ color: "var(--t4)", fontWeight: 500 }}>
            匯入去年報表後顯示
          </span>
        </div>
      )}
    </div>
  );
}

/* ─── 訂單展開明細 ───────────────────────────────────────────── */
function OrderDetail({ order, isSL, slFp, slCosts, spCosts }) {
  const costOf = (item) =>
    Object.prototype.hasOwnProperty.call(item, "snapshotCost") &&
    item.snapshotCost !== null
      ? Number(item.snapshotCost) || 0
      : Number((isSL ? slCosts : spCosts)[item.key]) || 0;

  const lines = isSL
    ? (() => {
        const fin = slOrderFin(order, slFp, slCosts);
        return [
          { l: "訂單營收", v: order.revenue },
          { l: "商品成本", v: -fin.oc, neg: true },
          {
            l: `金流手續費（${order.paymentMethod || "—"}${
              fin.payKnown ? "" : " ⚠ 比不到費率表，套 2.2%"
            }）`,
            v: -fin.pf,
            neg: true,
          },
          {
            l: `物流成本（${order.deliveryMethod || "—"}${
              fin.dlvKnown ? "" : ` ⚠ 比不到運費表，套預設 ${SL_SHIP_FALLBACK}`
            }）`,
            v: -fin.sc2,
            neg: true,
          },
          { l: "系統服務費", v: -fin.plf, neg: true },
          { l: "通路後毛利", v: fin.cm, sub: true },
          { l: "內部營業費", v: -fin.opx, neg: true },
          {
            l: order.isTaxExempt ? "稅賦（免稅）" : "稅賦",
            v: -fin.tax,
            neg: true,
          },
          { l: "最終淨利", v: fin.net, bold: true },
        ];
      })()
    : [
        { l: "商品總價（含補貼還原・已扣賣場券）", v: order.localGross },
        { l: "平台手續費＋金流＋蝦幣回饋", v: -order.totalOrderFee, neg: true },
        { l: "商品成本", v: -order.orderCost, neg: true },
        { l: "通路後毛利", v: order.grossProfit, sub: true },
        { l: "內部營業費", v: -order.orderOpExpense, neg: true },
        { l: "稅賦（以買家支付計）", v: -order.orderTax, neg: true },
        { l: "最終淨利", v: order.finalNetProfit, bold: true },
      ];

  const metaBits = isSL
    ? [
        order.status && `狀態：${order.status}`,
        order.voucherAmount > 0 && `優惠折讓：${fmt$(order.voucherAmount)}`,
        Number(order.refunded) > 0 &&
          `⚠ 已退款 ${fmt$(order.refunded)}（營收已扣成淨額）`,
        order.hasReturn && !(Number(order.refunded) > 0) && "⚠ 此訂單有退貨單",
      ].filter(Boolean)
    : [
        order.status && `狀態：${order.status}`,
        order.refundStatus && `退貨/退款：${order.refundStatus}`,
        numOrZero(order.sellerVoucher) > 0 &&
          `賣場優惠券 ${fmt$(order.sellerVoucher)}（報表的商品總價已扣除，不重複扣）`,
        numOrZero(order.buyerTotal) > 0 &&
          `買家總支付：${fmt$(order.buyerTotal)}`,
      ].filter(Boolean);

  const secTitle = {
    fontSize: 11,
    fontWeight: 700,
    color: "var(--t3)",
    letterSpacing: "0.05em",
    marginBottom: 6,
  };

  return (
    <div
      style={{
        display: "grid",
        gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
        gap: 20,
      }}
    >
      <div>
        <div style={secTitle}>商品明細</div>
        {(order.items || []).map((it, i) => {
          const c = costOf(it);
          const price = isSL
            ? Number(it.price) || 0
            : numOrZero(it.activityPrice);
          return (
            <div
              key={i}
              style={{
                display: "flex",
                justifyContent: "space-between",
                gap: 10,
                padding: "5px 0",
                borderBottom: "1px dashed var(--s3)",
                fontSize: 12,
              }}
            >
              <span style={{ color: "var(--t1)", minWidth: 0 }}>
                {it.name}
                {it.option ? `（${it.option}）` : ""} × {it.qty}
                {it.isGift && (
                  <span style={{ color: "var(--purple)" }}>（贈品）</span>
                )}
              </span>
              <span
                style={{
                  fontFamily: mono,
                  whiteSpace: "nowrap",
                  color: "var(--t2)",
                }}
              >
                {fmt$(price * it.qty)} ／ 成本{" "}
                {c > 0 ? (
                  fmt$(c * it.qty)
                ) : (
                  <span style={{ color: "var(--wn)", fontWeight: 700 }}>
                    未填
                  </span>
                )}
              </span>
            </div>
          );
        })}
        {metaBits.length > 0 && (
          <div
            style={{
              fontSize: 11,
              color: "var(--t3)",
              marginTop: 8,
              lineHeight: 1.8,
            }}
          >
            {metaBits.join("　")}
          </div>
        )}
      </div>
      <div>
        <div style={secTitle}>損益拆解</div>
        {lines.map((r, i) => (
          <div
            key={i}
            style={{
              display: "flex",
              justifyContent: "space-between",
              padding: r.sub || r.bold ? "8px 0 4px" : "4px 0",
              fontSize: 12,
              borderTop: r.sub || r.bold ? "1px solid var(--s3)" : "none",
              marginTop: r.sub || r.bold ? 4 : 0,
            }}
          >
            <span
              style={{
                color: r.bold ? "var(--t1)" : "var(--t2)",
                fontWeight: r.bold || r.sub ? 700 : 500,
              }}
            >
              {r.l}
            </span>
            <span
              style={{
                fontFamily: mono,
                fontWeight: r.bold ? 800 : 600,
                color: r.bold
                  ? r.v >= 0
                    ? "var(--up)"
                    : "var(--dn)"
                  : r.neg
                  ? "var(--dn)"
                  : "var(--t1)",
              }}
            >
              {fmt$(r.v)}
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}

/* 門市單筆訂單決策明細：版型同 OrderDetail（左 商品明細／右 損益拆解），
   差別在門市無平台抽成、稅只課開票單、成本來源顯示配方 */
function POSOrderDetail({
  order,
  costsEff,
  recipes,
  components,
  ratios,
  onToggleInvoice,
}) {
  const has = (it) =>
    Object.prototype.hasOwnProperty.call(it, "snapshotCost") &&
    it.snapshotCost !== null;
  const unitOf = (it) =>
    has(it) ? Number(it.snapshotCost) || 0 : posItemUnit(it, costsEff, ratios);
  /* 標籤必須與 unitOf 同源：已鎖快照 → 配方 → 手填 → 率 */
  /* 已鎖定且鎖定當下是估算值 → 仍標估算（率可能已刪，能查到就順便顯示） */
  const lockedEst = (it) => has(it) && it.snapshotEst === true;
  const rateOf = (it) =>
    lockedEst(it)
      ? Number(ratios?.[it.key]) || null
      : has(it) || Number(costsEff?.[it.key]) > 0
      ? null
      : posRatioUnit(it, costsEff, ratios) !== null
      ? Number(ratios[it.key])
      : null;
  const recipeOf = (key) => {
    const ls = recipes?.[key];
    if (!ls || !ls.length) return null;
    return ls
      .map((l) => `${components?.[l.compId]?.name || "（組件已刪）"}×${l.qty}`)
      .join("＋");
  };
  const items = order.items || [];
  const lineSum = items.reduce(
    (s, it) => s + (Number(it.price) || 0) * (it.qty || 1),
    0
  );
  /* revenue 已是扣退貨後淨額，退貨金額不能再算成「全單折扣」 */
  const discount =
    lineSum > 0
      ? Math.max(0, lineSum - order.revenue - (Number(order.refundAmt) || 0))
      : 0;
  const lines = order.missCost
    ? [
        { l: "訂單營收", v: order.revenue },
        { l: "商品成本", v: null, neg: true, note: items.length ? "未填" : "無明細" },
        { l: "毛利（無平台抽成）", v: null, sub: true, note: "—" },
        { l: "內部營業費", v: -order.opx, neg: true },
        {
          l: order.hasInvoice ? "稅賦（已開票）" : "稅賦（未開票免課）",
          v: -order.taxAmt,
          neg: true,
        },
        { l: "最終淨利", v: null, bold: true, note: "—（缺成本）" },
      ]
    : [
        { l: "訂單營收", v: order.revenue },
        { l: "商品成本", v: -order.oCost, neg: true },
        { l: "毛利（無平台抽成）", v: order.gp, sub: true },
        { l: "內部營業費", v: -order.opx, neg: true },
        {
          l: order.hasInvoice ? "稅賦（已開票）" : "稅賦（未開票免課）",
          v: -order.taxAmt,
          neg: true,
        },
        { l: "最終淨利", v: order.net, bold: true },
      ];
  const metaBits = [
    order.status && `狀態：${order.status}`,
    order.payMethod && `付款：${order.payMethod}`,
    order.staff && `銷售人員：${order.staff}`,
    order.taxId && `統編：${order.taxId}`,
    order.refundAmt > 0 &&
      `⚠ 部分退貨 ${fmt$(order.refundAmt)}：營收已扣成淨額，成本仍以全部商品計（退的是哪件 POS 匯出沒寫）`,
    order.swapSuspect &&
      "⚠ 疑似單價／數量顛倒：有商品單價≤1 元、數量>100，成本被放大，請到 POS 修單",
    discount > 0 && `全單折扣：${fmt$(discount)}（已按比例攤入商品）`,
    order.remark && `備註：${order.remark}`,
  ].filter(Boolean);
  const invLabel = order.hasInvoice ? "已開立" : "未開立";
  const invSrc =
    typeof order.invoiceOverride === "boolean"
      ? "手動設定"
      : order.invoiceSrc || (order.hasInvoice ? "" : "無號碼／統編／備註");
  const secTitle = {
    fontSize: 11,
    fontWeight: 700,
    color: "var(--t3)",
    letterSpacing: "0.05em",
    marginBottom: 6,
  };
  return (
    <div
      style={{
        display: "grid",
        gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
        gap: 20,
      }}
    >
      <div>
        <div style={secTitle}>商品明細</div>
        {items.length ? (
          items.map((it, i) => {
            const u = unitOf(it);
            const rc = recipeOf(it.key);
            return (
              <div
                key={i}
                style={{
                  padding: "5px 0",
                  borderBottom: "1px dashed var(--s3)",
                  fontSize: 12,
                }}
              >
                <div
                  style={{ display: "flex", justifyContent: "space-between", gap: 10 }}
                >
                  <span style={{ color: "var(--t1)", minWidth: 0 }}>
                    {it.name}
                    {it.option ? `（${it.option}）` : ""} × {it.qty}
                  </span>
                  <span
                    style={{ fontFamily: mono, whiteSpace: "nowrap", color: "var(--t2)" }}
                  >
                    {fmt$((Number(it.price) || 0) * it.qty)} ／ 成本{" "}
                    {u > 0 ? (
                      <>
                        {fmt$(u * it.qty)}
                        {has(it) && (
                          <span
                            style={{ fontSize: 9, color: "var(--t4)", marginLeft: 3 }}
                            title="此筆成本已隨本期快照鎖定，不受之後改原料價影響"
                          >
                            鎖
                          </span>
                        )}
                      </>
                    ) : (
                      <span style={{ color: "var(--wn)", fontWeight: 700 }}>未填</span>
                    )}
                  </span>
                </div>
                <div
                  style={{
                    fontSize: 10,
                    color: rc ? "var(--accent-text)" : "var(--t4)",
                    marginTop: 2,
                  }}
                >
                  {rc
                    ? `配方：${rc}`
                    : lockedEst(it)
                    ? `⚠ 估算（已鎖定${rateOf(it) ? `・成本率 ${rateOf(it)}%` : ""}）`
                    : rateOf(it)
                    ? `⚠ 估算：成本率 ${rateOf(it)}%（無規格可查，按金額比例估）`
                    : u > 0
                    ? "手填成本"
                    : "無成本來源——到下方成本資料庫掛配方或填成本"}
                </div>
              </div>
            );
          })
        ) : (
          <div style={{ fontSize: 12, color: "var(--wn)", lineHeight: 1.6 }}>
            這張交易沒有對到商品明細——POS訂單明細的日期範圍往前多抓一個月再匯一次，會自動補齊。
          </div>
        )}
        {metaBits.length > 0 && (
          <div
            style={{ fontSize: 11, color: "var(--t3)", marginTop: 8, lineHeight: 1.8 }}
          >
            {metaBits.join("　")}
          </div>
        )}
        {/* 發票判定：來源＋手動覆寫（企業客走 SHOPLINE 線上獨立發票時 POS 沒號碼，
            老闆可在備註寫「公司戶發票」自動認，或在這裡直接改） */}
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            marginTop: 8,
            fontSize: 11,
            color: "var(--t3)",
            flexWrap: "wrap",
          }}
        >
          <span>
            發票：
            <b style={{ color: order.hasInvoice ? "var(--up)" : "var(--t2)" }}>
              {invLabel}
            </b>
            {invSrc ? `（${invSrc}）` : ""}
            {order.hasInvoice ? "→ 課稅" : "→ 不課稅"}
          </span>
          {onToggleInvoice && (
            <Btn
              onClick={(e) => {
                e.stopPropagation();
                onToggleInvoice(order.orderId);
              }}
              style={{ fontSize: 10, padding: "3px 8px" }}
            >
              改為{order.hasInvoice ? "未開立" : "已開立"}
            </Btn>
          )}
          {typeof order.invoiceOverride === "boolean" && onToggleInvoice && (
            <Btn
              onClick={(e) => {
                e.stopPropagation();
                onToggleInvoice(order.orderId, "reset");
              }}
              style={{ fontSize: 10, padding: "3px 8px" }}
            >
              恢復自動判定
            </Btn>
          )}
        </div>
      </div>
      <div>
        <div style={secTitle}>損益拆解</div>
        {lines.map((r, i) => (
          <div
            key={i}
            style={{
              display: "flex",
              justifyContent: "space-between",
              padding: r.sub || r.bold ? "8px 0 4px" : "4px 0",
              fontSize: 12,
              borderTop: r.sub || r.bold ? "1px solid var(--s3)" : "none",
              marginTop: r.sub || r.bold ? 4 : 0,
            }}
          >
            <span
              style={{
                color: r.bold ? "var(--t1)" : "var(--t2)",
                fontWeight: r.bold || r.sub ? 700 : 500,
              }}
            >
              {r.l}
            </span>
            <span
              style={{
                fontFamily: mono,
                fontWeight: r.bold ? 800 : 600,
                color:
                  r.v === null
                    ? "var(--wn)"
                    : r.bold
                    ? r.v >= 0
                      ? "var(--up)"
                      : "var(--dn)"
                    : r.neg
                    ? "var(--dn)"
                    : "var(--t1)",
              }}
            >
              {r.v === null ? r.note : fmt$(r.v)}
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ─── Main App ───────────────────────────────────────────────── */
function ProfitCenter() {
  const [theme, setTheme] = useState(() => gl(SK.theme, "light"));
  const [platform, setPlatform] = useState(() => gl(SK.platform, "overview"));
  /* 存值蓋過預設，但新欄位（posTargetNet）要從預設補進舊存檔 */
  const [slFp, setSlFp] = useState(() => ({ ...DEFAULT_FP_SL, ...gl(SK.slFp, {}) }));
  const [spFp, setSpFp] = useState(() => ({ ...DEFAULT_FP_SP, ...gl(SK.spFp, {}) }));
  /* 內部營業費／預估稅率＝全公司單一口徑（老闆 2026-08-31 明示「全通路一致」），
     唯一來源是 slFp，只在官網頁可編輯；蝦皮、門市讀同一份值、畫面上唯讀顯示。
     spFp 只留自己的淨利目標（仍同步寫入 opExpense/tax 讓備份檔向下相容）。 */
  const [slCosts, setSlCosts] = useState(() => gl(SK.slCosts, {}));
  const [spCosts, setSpCosts] = useState(() => gl(SK.spCosts, {}));
  const [slOrders, setSlOrders] = useState(() => gl(SK.slOrders, {}));
  const [spOrders, setSpOrders] = useState(() => gl(SK.spOrders, {}));
  const [commissions, setCommissions] = useState(() => gl(SK.commissions, {}));
  const [components, setComponents] = useState(() => gl(SK.components, {}));
  const [slRecipes, setSlRecipes] = useState(() => gl(SK.slRecipes, {}));
  const [spRecipes, setSpRecipes] = useState(() => gl(SK.spRecipes, {}));
  const [posOrders, setPosOrders] = useState(() => gl(SK.posOrders, {}));
  const [posCosts, setPosCosts] = useState(() => gl(SK.posCosts, {}));
  const [posRecipes, setPosRecipes] = useState(() => gl(SK.posRecipes, {}));
  /* 泛稱品項成本率 { costKey: % }：見 posRatioUnit */
  const [posRatios, setPosRatios] = useState(() => gl(SK.posRatios, {}));
  /* 門市匯入：兩份 xls 分次拖入，先進暫存區、湊齊再 join */
  const posStage = useRef({ trans: null, orders: null });
  /* 門市通路過濾：KPI 預設只算「現場零售」（老闆 2026-08-18 定）；
     其他通路（經銷／電話／Omnichat…）仍顯示在通路表上、數字照算，勾了才計入 KPI。
     用「計入清單」而非「排除清單」：之後新出現的通路自動不計入，不會偷偷灌進零售 KPI */
  const [posIncluded, setPosIncluded] = useState(() => {
    /* 空陣列＝合法（全部不計入），要跟雲端一致，不能重整後偷偷變回 retail */
    const v = gl(SK.posIncluded, ["retail"]);
    return Array.isArray(v) ? v : ["retail"];
  });
  /* 門市運費：老闆 2026-08-19 拍板「不另計，視為含在內部營業費 % 裡」——
     勿再提運費欄位／每筆扣運費的方案 */

  const [toasts, setToasts] = useState([]);
  const toastIdRef = useRef(0);
  const [confirmBox, setConfirmBox] = useState(null);
  /* 預設期間＝當月（不是資料裡的最新月）：進來就停在這個月，沒匯報表就顯示空白 */
  const [sY, setSY] = useState(() => nowYM().y);
  const [sM, setSM] = useState(() => nowYM().m);
  const [range, setRange] = useState({ from: "", to: "" });
  const [search, setSearch] = useState("");
  const [mSearch, setMSearch] = useState("");
  const dSearch = useDebounced(search);
  const dMSearch = useDebounced(mSearch);
  const [lossOnly, setLossOnly] = useState(false);
  const [sync, setSync] = useState("connecting");
  const [cReady, setCReady] = useState(false);
  const [aReady, setAReady] = useState(false);
  const [costSort, setCostSort] = useState({ key: "soldQty", dir: "desc" });
  const [orderSort, setOrderSort] = useState({ key: "date", dir: "desc" });
  const [page, setPage] = useState(0);
  const [pageSize, setPageSize] = useState(30);
  const [dragOver, setDragOver] = useState(false);
  const [expandedId, setExpandedId] = useState(null);
  const [recipeEditKey, setRecipeEditKey] = useState(null);
  const [newComp, setNewComp] = useState({ name: "", price: "", cat: "" });
  /* 原料庫空的時候預設展開（引導新手），有內容則收合省空間 */
  const [compPanelOpen, setCompPanelOpen] = useState(
    () => Object.keys(gl(SK.components, {})).length === 0
  );
  const [compCatOpen, setCompCatOpen] = useState({});
  const [compSearch, setCompSearch] = useState("");
  const dCompSearch = useDebounced(compSearch);
  const [cleanupOnly, setCleanupOnly] = useState(false);
  const [soldOnly, setSoldOnly] = useState(false);

  const fRef = useRef({});
  const cRef = useRef(null);
  const cDoc = useRef(null);
  const firstMissRef = useRef(null);
  const prevSlMonthlyHashes = useRef({});
  const prevSpMonthlyHashes = useRef({});
  const prevPosMonthlyHashes = useRef({});
  const migrating = useRef(false);
  const sTimer = useRef(null);
  /* 寫入失敗的自動重試：saveTick 變動會重新觸發存檔 effect；saveRetry 限次數 */
  const [saveTick, setSaveTick] = useState(0);
  const saveRetry = useRef(0);
  /* 同步改用「內容比對」而非時間戳：lastMetaCore＝最後一次套用／寫入的 meta 內容
     （去掉 updatedAtMs/updatedBy、鍵序無關），月份文件用 prevXMonthlyHashes 的 ordersJson 比對。
     自己的 echo 內容相同→不重套；別台的變動內容不同→一定套用——不再被 50ms 視窗、
     兩台時鐘快慢、同批 meta 先到等條件擋掉（2026-09-03 盤查 M22–M25） */
  const lastMetaCore = useRef(null);
  /* 訂單 state 是 immutable 更新：參照沒變＝內容沒變。存檔時先比參照，
     省掉三平台 5,400 筆訂單重新 groupOrdersByMonth＋逐月 JSON.stringify
     （只改一個淨利目標也會跑一次，主執行緒約 20–40ms） */
  const lastSavedOrders = useRef({ sl: null, sp: null, pos: null });
  /* 首載閘門：meta＋三個月份集合的第一份快照都到齊才 cReady；之前不寫雲端、也不收匯入／改成本 */
  const firstLoads = useRef({ meta: false, sl: false, sp: false, pos: false });
  const meta = useRef({
    clientId: typeof window !== "undefined" ? gcid() : "",
  });
  const [lastSyncAt, setLastSyncAt] = useState(0);

  /* toast：帶動作（復原）的通知給較長的 10 秒，時間到一樣自動消失 */
  const toast = useCallback((msg, opts = {}) => {
    const id = ++toastIdRef.current;
    const { type = "info", action, actionLabel } = opts;
    const duration = opts.duration ?? (action ? 10000 : 3500);
    setToasts((p) => [
      ...p,
      { id, msg, type, duration, action, actionLabel, removing: false },
    ]);
    setTimeout(() => {
      setToasts((p) =>
        p.map((t) => (t.id === id ? { ...t, removing: true } : t))
      );
      setTimeout(() => setToasts((p) => p.filter((t) => t.id !== id)), 350);
    }, duration);
  }, []);
  const removeToast = useCallback((id) => {
    setToasts((p) =>
      p.map((t) => (t.id === id ? { ...t, removing: true } : t))
    );
    setTimeout(() => setToasts((p) => p.filter((t) => t.id !== id)), 350);
  }, []);

  /* localStorage 寫入失敗（容量滿）時提醒一次，避免無聲遺失本地快取 */
  const storageWarned = useRef(false);
  const persist = useCallback(
    (k, v) => {
      if (!sl_s(k, v) && !storageWarned.current) {
        storageWarned.current = true;
        toast("瀏覽器本地儲存空間不足，離線快取可能不完整（雲端同步不受影響）", {
          type: "warning",
          duration: 8000,
        });
      }
    },
    [toast]
  );

  useEffect(() => {
    persist(SK.theme, theme);
  }, [theme, persist]);
  useEffect(() => {
    persist(SK.platform, platform);
  }, [platform, persist]);
  useEffect(() => {
    persist(SK.slFp, slFp);
  }, [slFp, persist]);
  useEffect(() => {
    persist(SK.spFp, spFp);
  }, [spFp, persist]);
  useEffect(() => {
    persist(SK.slCosts, slCosts);
  }, [slCosts, persist]);
  useEffect(() => {
    persist(SK.spCosts, spCosts);
  }, [spCosts, persist]);
  useEffect(() => {
    persist(SK.slOrders, slOrders);
  }, [slOrders, persist]);
  useEffect(() => {
    persist(SK.spOrders, spOrders);
  }, [spOrders, persist]);
  useEffect(() => {
    persist(SK.commissions, commissions);
  }, [commissions, persist]);
  useEffect(() => {
    persist(SK.components, components);
  }, [components, persist]);
  useEffect(() => {
    persist(SK.slRecipes, slRecipes);
  }, [slRecipes, persist]);
  useEffect(() => {
    persist(SK.spRecipes, spRecipes);
  }, [spRecipes, persist]);
  useEffect(() => {
    persist(SK.posOrders, posOrders);
  }, [posOrders, persist]);
  useEffect(() => {
    persist(SK.posCosts, posCosts);
  }, [posCosts, persist]);
  useEffect(() => {
    persist(SK.posRecipes, posRecipes);
  }, [posRecipes, persist]);
  useEffect(() => {
    persist(SK.posRatios, posRatios);
  }, [posRatios, persist]);
  useEffect(() => {
    persist(SK.posIncluded, posIncluded);
  }, [posIncluded, persist]);

  /* Firebase init */
  useEffect(() => {
    try {
      const app = getApps().length ? getApp() : initializeApp(FBC);
      const auth = getAuth(app),
        db = getFirestore(app);
      cDoc.current = doc(db, FSPC.collection, FSPC.docId);
      fRef.current._db = db;
      fRef.current._slOrdDoc = doc(
        db,
        FSPC_SL_ORD.collection,
        FSPC_SL_ORD.docId
      );
      fRef.current._spOrdDoc = doc(
        db,
        FSPC_SP_ORD.collection,
        FSPC_SP_ORD.docId
      );
      fRef.current._slMonthlyColl = collection(db, SL_MONTHLY_COLL);
      fRef.current._spMonthlyColl = collection(db, SP_MONTHLY_COLL);
      fRef.current._posMonthlyColl = collection(db, POS_MONTHLY_COLL);
      setSync("connecting");
      const un = onAuthStateChanged(auth, async (u) => {
        try {
          if (!u) {
            await signInAnonymously(auth);
            return;
          }
          setAReady(true);
        } catch (e) {
          console.error("[Auth Error]", e);
          setSync("error");
        }
      });
      return () => un();
    } catch (e) {
      console.error("[Firebase Init Error]", e);
      setSync("error");
    }
  }, []);

  /* Firebase load (meta + monthly collections) */
  useEffect(() => {
    if (!aReady || !cDoc.current) return;

    const parseMeta = (snap) => {
      if (!snap.exists()) return null;
      const d = snap.data();
      if (d?.payloadJson) {
        try {
          return JSON.parse(d.payloadJson);
        } catch {
          return null;
        }
      }
      if (d?.payload) return d.payload;
      return null;
    };

    const runMigrationIfNeeded = async (metaSnap) => {
      if (migrating.current) return;
      const d = metaSnap.exists() ? metaSnap.data() : {};
      if (d?.splitByMonth === true) return;
      migrating.current = true;
      try {
        console.log("[Migration] Starting old-doc → monthly migration");
        // 以本地 state 為主（避免舊 doc 因寫入失敗而落後）
        let slSource = slOrders;
        let spSource = spOrders;
        if (Object.keys(slSource).length === 0) {
          try {
            const oldSnap = await getDoc(fRef.current._slOrdDoc);
            if (oldSnap.exists() && oldSnap.data()?.ordersJson) {
              slSource = JSON.parse(oldSnap.data().ordersJson);
              setSlOrders(slSource);
            }
          } catch (e) {
            console.error("[Migration read SL]", e);
          }
        }
        if (Object.keys(spSource).length === 0) {
          try {
            const oldSnap = await getDoc(fRef.current._spOrdDoc);
            if (oldSnap.exists() && oldSnap.data()?.ordersJson) {
              spSource = JSON.parse(oldSnap.data().ordersJson);
              setSpOrders(spSource);
            }
          } catch (e) {
            console.error("[Migration read SP]", e);
          }
        }
        const ms = Date.now();
        const slByMonth = groupOrdersByMonth(slSource);
        const spByMonth = groupOrdersByMonth(spSource);
        const writes = [];
        Object.entries(slByMonth).forEach(([ym, orders]) => {
          const json = JSON.stringify(orders);
          prevSlMonthlyHashes.current[ym] = json;
          writes.push(
            setDoc(doc(fRef.current._db, SL_MONTHLY_COLL, ym), {
              ordersJson: json,
              count: Object.keys(orders).length,
              updatedAtMs: ms,
            })
          );
        });
        Object.entries(spByMonth).forEach(([ym, orders]) => {
          const json = JSON.stringify(orders);
          prevSpMonthlyHashes.current[ym] = json;
          writes.push(
            setDoc(doc(fRef.current._db, SP_MONTHLY_COLL, ym), {
              ordersJson: json,
              count: Object.keys(orders).length,
              updatedAtMs: ms,
            })
          );
        });
        await Promise.all(writes);
        await setDoc(
          cDoc.current,
          {
            splitByMonth: true,
            migratedAt: serverTimestamp(),
          },
          { merge: true }
        );
        const totalSl = Object.keys(slSource).length;
        const totalSp = Object.keys(spSource).length;
        console.log(`[Migration] Done. SL=${totalSl}, SP=${totalSp}`);
        toast(`✓ 已完成按月拆分：官網 ${totalSl} 筆 + 蝦皮 ${totalSp} 筆`, {
          type: "success",
          duration: 8000,
        });
      } catch (e) {
        console.error("[Migration Error]", e);
        toast("遷移失敗：" + e.message, { type: "error", duration: 8000 });
      } finally {
        migrating.current = false;
      }
    };

    // === LEGACY DATA RESTORE (manual via DevTools) ===
    // Usage in browser console:
    //   window.forceLegacyRestore()           -> inspect legacy docs only
    //   window.forceLegacyRestore("apply")   -> write legacy orders into monthly collections
    window.forceLegacyRestore = async (mode) => {
      try {
        console.log("[Restore] mode:", mode || "inspect");
        const slOldSnap = await getDoc(fRef.current._slOrdDoc);
        const spOldSnap = await getDoc(fRef.current._spOrdDoc);
        const slOldRaw = slOldSnap.exists() ? slOldSnap.data() : null;
        const spOldRaw = spOldSnap.exists() ? spOldSnap.data() : null;
        const parse = (raw) => {
          if (!raw) return null;
          if (raw.ordersJson) {
            try { return JSON.parse(raw.ordersJson); } catch { return null; }
          }
          return raw.orders || null;
        };
        const slLegacy = parse(slOldRaw) || {};
        const spLegacy = parse(spOldRaw) || {};
        const slCount = Object.keys(slLegacy).length;
        const spCount = Object.keys(spLegacy).length;
        const sample = (obj) =>
          Object.entries(obj)
            .slice(0, 2)
            .map(([k, v]) => ({
              k: k.length > 60 ? k.slice(0, 60) + "..." : k,
              vType: typeof v,
              vKeys: v && typeof v === "object" ? Object.keys(v).slice(0, 8) : null,
            }));
        const slSample = sample(slLegacy);
        const spSample = sample(spLegacy);
        console.log("[Restore] sl legacy count:", slCount, "sample:", slSample);
        console.log("[Restore] sp legacy count:", spCount, "sample:", spSample);
        if (mode !== "apply") {
          console.log("[Restore] inspect-only. Run forceLegacyRestore('apply') to write monthly docs.");
          return { slCount, spCount, slSample, spSample };
        }
        if (!window.confirm("Will write " + slCount + " SL + " + spCount + " SP orders into monthly collections. Proceed?")) return;
        const slByMonth = groupOrdersByMonth(slLegacy);
        const spByMonth = groupOrdersByMonth(spLegacy);
        const ms = Date.now();
        const writes = [];
        for (const [ym, orders] of Object.entries(slByMonth)) {
          writes.push(
            setDoc(doc(fRef.current._db, SL_MONTHLY_COLL, ym), {
              ordersJson: JSON.stringify(orders),
              count: Object.keys(orders).length,
              updatedAtMs: ms,
            })
          );
        }
        for (const [ym, orders] of Object.entries(spByMonth)) {
          writes.push(
            setDoc(doc(fRef.current._db, SP_MONTHLY_COLL, ym), {
              ordersJson: JSON.stringify(orders),
              count: Object.keys(orders).length,
              updatedAtMs: ms,
            })
          );
        }
        await Promise.all(writes);
        setSlOrders(slLegacy);
        setSpOrders(spLegacy);
        console.log("[Restore] DONE. Wrote", writes.length, "monthly docs.");
        return { wrote: writes.length, slCount, spCount };
      } catch (e) {
        console.error("[Restore] error:", e);
        return { error: e.message };
      }
    };
    // === END LEGACY DATA RESTORE ===

    const markFirst = (k) => {
      firstLoads.current[k] = true;
      if (Object.values(firstLoads.current).every(Boolean)) setCReady(true);
    };

    // meta 監聽（內容比對：自己寫入的 echo 內容相同→不重套；別台改了→一定套）
    const unMeta = onSnapshot(
      cDoc.current,
      async (snap) => {
        try {
          const metaData = parseMeta(snap);
          const core = metaCoreOf(metaData);
          /* 自己還沒被 server 確認的寫入（hasPendingWrites）不套用：lastMetaCore 改成
             「寫成功才推進」之後，這裡不擋會把自己 in-flight 的內容當成別台的改動重套 */
          const own = !!snap.metadata?.hasPendingWrites;
          if (metaData && !own && core !== lastMetaCore.current) {
            lastMetaCore.current = core;
            if (metaData.slFp) setSlFp({ ...DEFAULT_FP_SL, ...metaData.slFp });
            if (metaData.spFp) setSpFp({ ...DEFAULT_FP_SP, ...metaData.spFp });
            if (metaData.slCosts) setSlCosts(metaData.slCosts);
            if (metaData.spCosts) setSpCosts(metaData.spCosts);
            if (metaData.commissions) setCommissions(metaData.commissions);
            if (metaData.components) setComponents(metaData.components);
            if (metaData.slRecipes) setSlRecipes(metaData.slRecipes);
            if (metaData.spRecipes) setSpRecipes(metaData.spRecipes);
            if (metaData.posCosts) setPosCosts(metaData.posCosts);
            if (metaData.posRecipes) setPosRecipes(metaData.posRecipes);
            if (metaData.posRatios) setPosRatios(metaData.posRatios);
            /* 空陣列＝合法設定（全部不計入），照樣套用才會兩台一致 */
            if (Array.isArray(metaData.posIncluded))
              setPosIncluded(metaData.posIncluded);
            setLastSyncAt(Date.now());
          }
          await runMigrationIfNeeded(snap);
          markFirst("meta");
          /* 只把「連線中」轉成已同步；待同步／儲存中／失敗由存檔路徑自己管，
             別台的快照到了不代表自己的寫入成功了 */
          setSync((s) => (s === "connecting" ? "synced" : s));
        } catch (e) {
          console.error("[Meta Snapshot Error]", e);
          markFirst("meta");
          setSync("error");
        }
      },
      (err) => {
        console.error("[Meta Snapshot Error]", err);
        markFirst("meta");
        setSync("error");
      }
    );

    /* 三個月份集合共用同一支監聽（原本三段 108 行完全同構，修一份漏兩份）：
       以每個月份 doc 的 ordersJson 內容比對——新增／修改／刪除 doc 都會被偵測到；
       hash 只在真的套用時才整組換掉，自己剛寫入的內容本來就等於 hash → 不會重套自己 */
    const listenMonthly = (coll, setter, hashesRef, key, label) => {
      let first = true;
      const done = () => {
        if (first) {
          first = false;
          markFirst(key);
        }
      };
      return onSnapshot(
        coll,
        (snapshot) => {
          try {
            const prev = hashesRef.current;
            const next = {};
            snapshot.forEach((docSnap) => {
              /* 自己 in-flight 的寫入（還沒被 server 確認）：當作雲端沒變，
                 等 ack 到了 hash 也推進了，自然對得上 */
              if (docSnap.metadata?.hasPendingWrites) {
                if (prev[docSnap.id] !== undefined) next[docSnap.id] = prev[docSnap.id];
                return;
              }
              next[docSnap.id] = docSnap.data()?.ordersJson || "";
            });
            const changed =
              Object.keys(next).some((ym) => prev[ym] !== next[ym]) ||
              Object.keys(prev).some((ym) => !(ym in next));
            if (first) {
              /* 首載：雲端為準、整份取代——本機 localStorage 可能是別台改過之前的舊版，
                 不能拿來蓋雲端（2026-09-03 M26） */
              const all = {};
              Object.values(next).forEach((json) => {
                if (!json) return;
                try {
                  Object.assign(all, JSON.parse(json));
                } catch {}
              });
              hashesRef.current = next;
              setter(all);
              setLastSyncAt(Date.now());
            } else if (changed) {
              /* 之後的快照：逐月合併，不整包取代。本機在 900ms 防抖內還沒寫出去的改動
                 （匯入／鎖定／發票覆寫／重置）不能被別台的快照沖掉（2026-09-04 delta 審查 3/3）。
                 判「本機這個月有沒有改」＝本機該月 json 是否還等於上次同步點的 hash */
              setter((local) => {
                const localBy = groupOrdersByMonth(local);
                const merged = {};
                const newHash = {};
                let conflicts = 0;
                new Set([...Object.keys(next), ...Object.keys(localBy)]).forEach((ym) => {
                  const cloudJson = next[ym];
                  const oldHash = prev[ym];
                  const localMo = localBy[ym] || null;
                  const localJson = localMo ? JSON.stringify(localMo) : undefined;
                  const localDirty = localJson !== oldHash;
                  if (cloudJson === undefined) {
                    /* 雲端沒這個月（別台重置）：本機沒改→跟著刪；本機有改→留著，之後會寫上去 */
                    if (localDirty && localMo) {
                      Object.assign(merged, localMo);
                      if (oldHash !== undefined) newHash[ym] = oldHash;
                    }
                    return;
                  }
                  if (cloudJson === oldHash) {
                    /* 雲端沒變：保留本機（可能有待寫的改動） */
                    if (localMo) Object.assign(merged, localMo);
                    newHash[ym] = cloudJson;
                    return;
                  }
                  let cloudMo = {};
                  try {
                    cloudMo = JSON.parse(cloudJson);
                  } catch {}
                  if (!localDirty || localJson === cloudJson) {
                    /* 雲端變了、本機沒改（或改成一模一樣＝自己的 ack 先到）：吃雲端 */
                    Object.assign(merged, cloudMo);
                    newHash[ym] = cloudJson;
                    return;
                  }
                  /* 兩邊都變且不同：訂單級合併，本機改過的訂單優先；
                     hash 停在舊值 → 下次存檔會把合併結果寫出去 */
                  conflicts++;
                  Object.assign(merged, cloudMo, localMo);
                  if (oldHash !== undefined) newHash[ym] = oldHash;
                });
                hashesRef.current = newHash;
                if (conflicts)
                  setTimeout(
                    () =>
                      toast(
                        `另一台剛改過${label}的 ${conflicts} 個月份，已與本機未存檔的改動合併`,
                        { type: "warning", duration: 8000 }
                      ),
                    0
                  );
                return merged;
              });
              setLastSyncAt(Date.now());
            }
          } catch (e) {
            console.error(`[${label} Monthly Snapshot Error]`, e);
          }
          done();
        },
        (err) => {
          console.error(`[${label} Monthly Snapshot Error]`, err);
          done();
        }
      );
    };
    const unSl = listenMonthly(
      fRef.current._slMonthlyColl,
      setSlOrders,
      prevSlMonthlyHashes,
      "sl",
      "SL"
    );
    const unSp = listenMonthly(
      fRef.current._spMonthlyColl,
      setSpOrders,
      prevSpMonthlyHashes,
      "sp",
      "SP"
    );
    const unPos = listenMonthly(
      fRef.current._posMonthlyColl,
      setPosOrders,
      prevPosMonthlyHashes,
      "pos",
      "POS"
    );

    return () => {
      unMeta();
      unSl();
      unSp();
      unPos();
    };
    // 監聽器只在登入完成時建立一次；slOrders/spOrders 僅供一次性遷移讀取
    // eslint-disable-next-line
  }, [aReady]);

  /* Firebase save (meta + only changed months) */
  useEffect(() => {
    /* cReady＝四個監聽的首份快照都到齊；在那之前本機 state 還沒跟雲端對齊，不能寫。
       不再用 applying 早退（那會把 900ms 內被遠端快照打斷的本機改動永久丟掉）：
       改由「內容比對」決定要不要寫，echo 回來內容相同就不會再寫一次 */
    if (!aReady || !cReady || !cDoc.current) return;
    if (migrating.current) return;
    clearTimeout(sTimer.current);
    /* 有待寫改動時燈先轉「待同步」，關分頁會被 beforeunload 攔下；
       timer 跑完發現沒東西要寫會回 synced（遠端套用造成的 900ms 閃爍可接受） */
    setSync((s) => (s === "synced" ? "pending" : s));
    sTimer.current = setTimeout(async () => {
      sTimer.current = null;
      const onSuccess = [];
      try {
        const ms = Date.now();
        const db = fRef.current._db;

        const metaPl = deepClean({
          slFp,
          spFp,
          slCosts,
          spCosts,
          commissions,
          components,
          slRecipes,
          spRecipes,
          posCosts,
          posRecipes,
          posRatios,
          posIncluded,
          updatedAtMs: ms,
          updatedBy: meta.current.clientId,
        });

        const writes = [];
        /* meta 只在內容真的變了才寫（去掉時間戳比對）——從雲端套回來的內容不會被自己再寫一次 */
        /* hash／lastMetaCore／lastSavedOrders 一律「寫成功才推進」——寫入被 server 拒絕時
           記號留在舊值，下一輪存檔會自然重寫；以前在 await 之前就推進，失敗後永遠不重試
           且燈會變綠（2026-09-04 delta 審查 3/3） */
        const core = metaCoreOf(metaPl);
        if (core !== lastMetaCore.current) {
          writes.push(
            setDoc(
              cDoc.current,
              {
                payloadJson: JSON.stringify(metaPl),
                updatedAtMs: ms,
                updatedBy: metaPl.updatedBy,
                splitByMonth: true,
                updatedAtServer: serverTimestamp(),
              },
              { merge: true }
            ).then(() => {
              lastMetaCore.current = core;
            })
          );
        }

        /* 三平台共用：訂單參照沒變就整段跳過（不重算月份分組、不重新序列化） */
        const pushMonthlyWrites = (orders, coll, hashesRef, refKey) => {
          if (lastSavedOrders.current[refKey] === orders) return;
          const byMonth = groupOrdersByMonth(orders);
          Object.entries(byMonth).forEach(([ym, mo]) => {
            const json = JSON.stringify(mo);
            if (hashesRef.current[ym] === json) return;
            writes.push(
              setDoc(doc(db, coll, ym), {
                ordersJson: json,
                count: Object.keys(mo).length,
                updatedAtMs: ms,
              }).then(() => {
                hashesRef.current[ym] = json;
              })
            );
          });
          /* 偵測已刪除的月份（重置本期／清空最後一筆） */
          Object.keys(hashesRef.current).forEach((ym) => {
            if (!byMonth[ym]) {
              writes.push(
                deleteDoc(doc(db, coll, ym)).then(() => {
                  delete hashesRef.current[ym];
                })
              );
            }
          });
          /* 這個平台全部寫成功才記「已存過這個參照」 */
          onSuccess.push(() => {
            lastSavedOrders.current[refKey] = orders;
          });
        };
        pushMonthlyWrites(slOrders, SL_MONTHLY_COLL, prevSlMonthlyHashes, "sl");
        pushMonthlyWrites(spOrders, SP_MONTHLY_COLL, prevSpMonthlyHashes, "sp");
        pushMonthlyWrites(
          posOrders,
          POS_MONTHLY_COLL,
          prevPosMonthlyHashes,
          "pos"
        );

        if (!writes.length) {
          onSuccess.forEach((f) => f());
          setSync("synced");
          return;
        }
        setSync("saving");
        await Promise.all(writes);
        onSuccess.forEach((f) => f());
        saveRetry.current = 0;
        setLastSyncAt(Date.now());
        setSync("synced");
      } catch (e) {
        console.error("[Save Error]", e);
        /* 記號都沒推進，下一輪存檔會自然重寫。這裡只負責亮紅燈＋排一次自動重試
           （最多兩次；再失敗就等使用者下一個動作再觸發） */
        setSync("error");
        if (saveRetry.current < 2) {
          saveRetry.current += 1;
          setTimeout(() => setSaveTick((t) => t + 1), 5000);
        }
      }
    }, 900);
    return () => {
      clearTimeout(sTimer.current);
      sTimer.current = null;
    };
  }, [
    saveTick,
    slFp,
    spFp,
    slCosts,
    spCosts,
    slOrders,
    spOrders,
    commissions,
    components,
    slRecipes,
    spRecipes,
    posOrders,
    posCosts,
    posRecipes,
    posRatios,
    posIncluded,
    aReady,
    cReady,
  ]);

  /* 還有沒寫出去的改動（timer 待跑／儲存中／失敗）時關分頁要攔：以前 900ms 內
     關掉分頁＝雲端沒收到、重開又被首載快照蓋回去，燈卻一直顯示已同步 */
  const syncRef = useRef(sync);
  useEffect(() => {
    syncRef.current = sync;
  }, [sync]);
  useEffect(() => {
    const h = (e) => {
      const s = syncRef.current;
      if (sTimer.current || s === "pending" || s === "saving" || s === "error") {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", h);
    return () => window.removeEventListener("beforeunload", h);
  }, []);

  /* 匯入時的缺欄警告：平台改欄名時，缺席的金額欄會被靜默當 0，
     報表照樣說「已匯入 N 筆」。這支把缺席的關鍵欄一次講清楚 */
  const warnMissingCols = (pairs, hdrs) => {
    const miss = pairs.filter(([, i]) => i === undefined || i === -1).map(([n]) => n);
    if (!miss.length) return;
    console.warn("[匯入] 缺少關鍵欄位", miss, "實際檔頭：", hdrs);
    toast(
      `⚠ 這份報表缺少關鍵欄位：${miss.join("、")}——相關金額會以 0 計算，數字會失真。請確認匯出時有勾這些欄位`,
      { type: "warning", duration: 12000 }
    );
  };

  /* ─── Shopline CSV/XLSX Parser ─────────────────────────────── */
  const processSLParsed = (parsed) => {
    if (!Array.isArray(parsed) || parsed.length < 2) {
      toast("格式錯誤", { type: "error" });
      return;
    }
    const hdrs = parsed[0].map((h) => safeText(h).replace(/^\uFEFF/, ""));
    const idx = (n) => nzIndex(hdrs, n);
    const idxF = (a, b) => {
      const i = idx(a);
      return i !== -1 ? i : idx(b);
    };

    const im = {
      cartId: idx("購物車編號"),
      orderId: idx("訂單號碼"),
      date: idxF("訂單日期", "訂單成立於"),
      status: idx("訂單狀態"),
      payMethod: idx("付款方式"),
      delivery: idx("送貨方式"),
      subtotal: idx("訂單小計"),
      shippingFee: idx("運費"),
      discount: idx("優惠折扣"),
      creditOffset: idx("折抵購物金"),
      pointOffset: idx("點數折現"),
      total: idx("訂單合計"),
      paidTotal: idx("付款總金額"),
      refunded: idx("已退款金額"),
      prodName: idx("商品名稱"),
      option: idx("選項"),
      prodId: idx("商品貨號"),
      qty: idx("數量"),
      unitPrice: idxF("商品結帳價", "商品原價"),
      prodType: idx("商品類型"),
      addOnType: idx("加購品類型"),
      itemDiscount: idx("商品折扣金額"),
      orderShare: idx("全單折扣金額"),
      creditShare: idx("折抵購物金分攤"),
      pointShare: idx("點數折現分攤"),
      taxExempt: idx("發票稅別"),
      invoiceStatus: idx("發票狀態"),
      returnId: idx("退貨單編號"),
    };

    if (im.orderId === -1 || im.date === -1) {
      toast("找不到必要欄位（訂單號碼/訂單日期），請確認是 Shopline 標準報表", {
        type: "error",
      });
      return;
    }
    /* 檔頭指紋：官網報表一定有「送貨方式」或「購物車編號」；門市 POS 訂單明細沒有、
       蝦皮有「商品ID」——丟錯分頁不能靜默當官網單套費率 */
    if (idx("商品ID") > -1) {
      toast("這份是蝦皮報表（有「商品ID」欄），請切到蝦皮分頁再匯入", {
        type: "error",
        duration: 8000,
      });
      return;
    }
    if (im.delivery === -1 && im.cartId === -1) {
      toast(
        "這份缺「送貨方式／購物車編號」欄，不像官網報表——若是門市 POS 訂單明細請切到門市分頁",
        { type: "error", duration: 8000 }
      );
      return;
    }
    /* 金額關鍵欄缺席只會靜默當 0（營收 0、成本照扣＝整批顯示虧損），必須講出來 */
    warnMissingCols(
      [
        ["訂單狀態", im.status],
        ["付款總金額／訂單合計", im.paidTotal > -1 ? im.paidTotal : im.total],
        ["商品結帳價／商品原價", im.unitPrice],
        ["數量", im.qty],
      ],
      hdrs
    );

    const newOrders = {};

    for (let i = 1; i < parsed.length; i++) {
      const row = parsed[i];
      if (!row || row.length < 5) continue;

      const rawOrderId = safeText(row[im.orderId]);
      if (!rawOrderId) continue;

      const cartId = im.cartId > -1 ? safeText(row[im.cartId]) : rawOrderId;
      const groupKey = cartId || rawOrderId;

      const date = normDate(row[im.date]);

      const rowRevenue = () =>
        numOrZero(
          im.paidTotal > -1
            ? row[im.paidTotal]
            : im.total > -1
            ? row[im.total]
            : 0
        );
      const rowVoucher = () =>
        numOrZero(im.discount > -1 ? row[im.discount] : 0) +
        numOrZero(im.creditOffset > -1 ? row[im.creditOffset] : 0) +
        numOrZero(im.pointOffset > -1 ? row[im.pointOffset] : 0);
      const rowShipping = () =>
        numOrZero(im.shippingFee > -1 ? row[im.shippingFee] : 0);
      const rowRefunded = () =>
        numOrZero(im.refunded > -1 ? row[im.refunded] : 0);

      if (!newOrders[groupKey]) {
        const statusRaw = im.status > -1 ? safeText(row[im.status]) : "";
        const isTaxExempt =
          (im.taxExempt > -1 && safeText(row[im.taxExempt]) === "免稅") ||
          (im.invoiceStatus > -1 &&
            safeText(row[im.invoiceStatus]) === "待開立");
        const hasReturn = im.returnId > -1 && safeText(row[im.returnId]) !== "";
        const payMethod = im.payMethod > -1 ? safeText(row[im.payMethod]) : "";
        const delivMethod = im.delivery > -1 ? safeText(row[im.delivery]) : "";

        newOrders[groupKey] = {
          orderId: rawOrderId,
          date,
          status: statusRaw,
          revenue: rowRevenue(),
          voucherAmount: rowVoucher(),
          shippingIncome: rowShipping(),
          /* 已退款金額：全額退視同取消、部分退以淨額入帳（slData／slMonthly 同一判斷） */
          refunded: rowRefunded(),
          paymentMethod: payMethod,
          deliveryMethod: delivMethod,
          isTaxExempt,
          hasReturn,
          items: [],
          _orderIds: [rawOrderId],
        };
      } else if (!newOrders[groupKey]._orderIds.includes(rawOrderId)) {
        // 同一購物車內的另一張訂單：金額累加，避免只取第一張造成漏算。
        // 但這張自己是取消／刪除就不累加金額——否則整組會「非全對即全錯」
        // （狀態／免稅／退貨旗標仍只取第一張，至少讓金額是對的）
        newOrders[groupKey]._orderIds.push(rawOrderId);
        const rowStatus = im.status > -1 ? safeText(row[im.status]) : "";
        if (!rowStatus.includes("取消") && !rowStatus.includes("刪除")) {
          newOrders[groupKey].revenue += rowRevenue();
          newOrders[groupKey].voucherAmount += rowVoucher();
          newOrders[groupKey].shippingIncome += rowShipping();
          newOrders[groupKey].refunded += rowRefunded();
        }
      }

      const prodName =
        im.prodName > -1 ? safeText(row[im.prodName]) : "未知商品";
      const option = im.option > -1 ? safeText(row[im.option]) : "";
      const qty = parseInt(im.qty > -1 ? row[im.qty] || 1 : 1, 10) || 1;
      const price = numOrZero(im.unitPrice > -1 ? row[im.unitPrice] : 0);
      const prodType = im.prodType > -1 ? safeText(row[im.prodType]) : "商品";
      const addOnType = im.addOnType > -1 ? safeText(row[im.addOnType]) : "";
      const isGift = prodType === "贈品";

      if (!prodName) continue;

      const costKey = `${prodName}_${option}`.trim();

      /* 逐列分攤：全單折扣＋購物金＋點數（報表已按商品行攤好），商品表營收要扣掉才對得上訂單營收 */
      const share =
        numOrZero(im.orderShare > -1 ? row[im.orderShare] : 0) +
        numOrZero(im.creditShare > -1 ? row[im.creditShare] : 0) +
        numOrZero(im.pointShare > -1 ? row[im.pointShare] : 0);
      newOrders[groupKey].items.push({
        key: costKey,
        name: prodName,
        option,
        qty,
        price,
        share,
        isGift,
        isAddOn: prodType === "加購品",
        addOnType,
      });
    }

    /* 預掃部分重匯數量（供匯入提示用；實際判斷在 functional updater 內以最新狀態為準） */
    let partialKept = 0;
    {
      const idxPre = {};
      Object.entries(slOrders).forEach(([sk, o]) =>
        (o.memberIds || [o.orderId]).forEach((id) => {
          idxPre[id] = sk;
        })
      );
      Object.values(newOrders).forEach((order) => {
        let sk = null;
        for (const id of order._orderIds)
          if (idxPre[id]) {
            sk = idxPre[id];
            break;
          }
        const old = sk ? slOrders[sk] : null;
        if (
          old?.memberIds &&
          old.memberIds.some((id) => !order._orderIds.includes(id))
        )
          partialKept++;
      });
    }
    setSlOrders((p) => {
      const merged = { ...p };
      /* 成員單號索引：同一購物車的任何一張單號都指回同一筆儲存紀錄，
         避免之後的部分報表（只含 cart 內另一張單）被誤判為新訂單而重複入帳 */
      const memberIdx = {};
      Object.entries(merged).forEach(([sk, o]) => {
        (o.memberIds || [o.orderId]).forEach((id) => {
          memberIdx[id] = sk;
        });
      });
      Object.values(newOrders).forEach((order) => {
        const { _orderIds, ...clean } = order;
        let sk = clean.orderId;
        for (const id of _orderIds) {
          if (memberIdx[id]) {
            sk = memberIdx[id];
            break;
          }
        }
        const old = merged[sk];
        /* 部分重匯防護：本次檔案缺了該購物車既有成員單號時，
           整筆覆蓋會讓缺席成員的營收靜默消失——保留舊合併紀錄不動 */
        if (
          old?.memberIds &&
          old.memberIds.some((id) => !_orderIds.includes(id))
        ) {
          return;
        }
        /* 吸收同購物車先前被分開儲存的舊紀錄（歷史資料時期產生），
           避免同一張單同時存在於合併紀錄與獨立紀錄而重複入帳 */
        _orderIds.forEach((id) => {
          const otherKey = memberIdx[id];
          if (otherKey && otherKey !== sk && merged[otherKey]) {
            delete merged[otherKey];
          }
        });
        clean.orderId = sk;
        clean.memberIds = [
          ...new Set([...(old?.memberIds || []), ..._orderIds]),
        ];
        clean.memberIds.forEach((id) => {
          memberIdx[id] = sk;
        });
        /* 重匯同單：保留既有快照，避免鎖定的歷史參數被靜默清除 */
        merged[sk] = withOldSnapshot(old, clean);
      });
      return merged;
    });
    const dates = Object.values(newOrders)
      .map((o) => String(o.date))
      .filter(Boolean)
      .sort()
      .reverse();
    if (dates.length) {
      setSY(dates[0].substring(0, 4));
      setSM(dates[0].substring(5, 7));
    }
    /* 比不到費率表的送貨／付款方式：不再無聲套預設，匯入時點名一次（明細列也會標 ⚠） */
    const unkDlv = new Set();
    const unkPay = new Set();
    Object.values(newOrders).forEach((o) => {
      if (!slIntl(o.deliveryMethod) && slShipRate(o.deliveryMethod) === null)
        unkDlv.add(o.deliveryMethod || "（空白）");
      if (slPayRate(o.paymentMethod) === null)
        unkPay.add(o.paymentMethod || "（空白）");
    });
    const unkMsg =
      (unkDlv.size
        ? `；⚠ ${unkDlv.size} 種送貨方式比不到運費表（${[...unkDlv].join("、")}），暫套預設 ${SL_SHIP_FALLBACK} 元`
        : "") +
      (unkPay.size
        ? `；⚠ ${unkPay.size} 種付款方式比不到費率表（${[...unkPay].join("、")}），暫套 2.2%`
        : "");
    toast(
      /* count 是購物車數；訂單數＝各購物車的 _orderIds 總和（多半相等，同購物車拆單時才差） */
      `已匯入 ${Object.values(newOrders).reduce(
        (a, o) => a + (o._orderIds?.length || 1),
        0
      )} 筆官網訂單` +
        (partialKept > 0
          ? `（${partialKept} 筆購物車部分重匯，已保留原合併紀錄——如需更正請重匯含完整購物車的報表）`
          : "") +
        unkMsg,
      {
        type: unkMsg ? "warning" : "success",
        duration: unkMsg || partialKept > 0 ? 10000 : 3500,
      }
    );
  };

  /* ─── Shopee CSV/XLSX Parser ────────────────────────────────── */
  const processSPParsed = (parsed) => {
    if (!Array.isArray(parsed) || parsed.length < 2) {
      toast("格式錯誤", { type: "error" });
      return;
    }
    const hdrs = parsed[0].map((h) => safeText(h).replace(/^\uFEFF/, ""));
    const idx = (n) => nzIndex(hdrs, n);
    const idxF = (a, b) => {
      const i = idx(a);
      return i !== -1 ? i : idx(b);
    };

    const im = {
      orderId: idx("訂單編號"),
      date: idx("訂單成立日期"),
      status: idx("訂單狀態"),
      refundStatus: idx("退貨 / 退款狀態"),
      grossPrice: idx("商品總價"),
      buyerTotal: idx("買家總支付金額"),
      coinDiscount: idx("蝦幣折抵"),
      platformSubsidy: idx("蝦皮補貼金額"),
      platformShippingSubsidy: idx("蝦皮補助運費"),
      sellerVoucher: idxF("賣場優惠券", "賣家負擔優惠券"),
      platformVoucher: idxF("優惠券", "蝦皮負擔優惠券"),
      sellerCoinCashback: idxF("賣家蝦幣回饋券", "賣家負擔蝦幣回饋券"),
      txFee: idx("成交手續費"),
      otherFee: idx("其他服務費"),
      paymentFee: idx("金流與系統處理費"),
      prodName: idx("商品名稱"),
      optName: idx("商品選項名稱"),
      prodId: idx("商品ID"),
      optId: idx("規格ID"),
      qty: idx("數量"),
      activityPrice: idx("商品活動價格"),
    };

    if (im.orderId === -1 || im.date === -1) {
      toast("找不到必要欄位，請確認是蝦皮標準報表", { type: "error" });
      return;
    }
    /* 檔頭指紋：官網／門市報表丟到蝦皮分頁要擋下來 */
    if (idx("購物車編號") > -1 || idx("送貨方式") > -1) {
      toast("這份是官網報表（有「送貨方式／購物車編號」欄），請切到官網分頁再匯入", {
        type: "error",
        duration: 8000,
      });
      return;
    }
    if (idx("交易號碼") > -1) {
      toast("這份是門市交易明細（有「交易號碼」欄），請切到門市分頁再匯入", {
        type: "error",
        duration: 8000,
      });
      return;
    }
    /* 金額關鍵欄缺席＝算出來的錢一定錯（手續費少扣＝淨利虛高、退款狀態讀不到＝
       退款單全算營收），但程式只會靜默當 0。只警告這幾欄，避免警報疲勞 */
    warnMissingCols([
      ["訂單狀態", im.status],
      ["退貨 / 退款狀態", im.refundStatus],
      ["商品總價", im.grossPrice],
      ["成交手續費", im.txFee],
      ["金流與系統處理費", im.paymentFee],
      ["數量", im.qty],
    ], hdrs);

    const newOrders = {};
    let count = 0;

    for (let i = 1; i < parsed.length; i++) {
      const row = parsed[i];
      if (!row || row.length < 5) continue;
      const orderId = safeText(row[im.orderId]);
      if (!orderId) continue;
      const date = normDate(row[im.date]);

      if (!newOrders[orderId]) {
        const rawGross =
          numOrZero(im.grossPrice > -1 ? row[im.grossPrice] : 0) +
          numOrZero(im.coinDiscount > -1 ? row[im.coinDiscount] : 0) +
          numOrZero(im.platformSubsidy > -1 ? row[im.platformSubsidy] : 0) +
          numOrZero(im.platformVoucher > -1 ? row[im.platformVoucher] : 0);

        newOrders[orderId] = {
          orderId,
          date,
          status: im.status > -1 ? safeText(row[im.status]) : "",
          refundStatus:
            im.refundStatus > -1 ? safeText(row[im.refundStatus]) : "",
          grossPrice: rawGross,
          buyerTotal: numOrZero(im.buyerTotal > -1 ? row[im.buyerTotal] : 0),
          sellerVoucher: numOrZero(
            im.sellerVoucher > -1 ? row[im.sellerVoucher] : 0
          ),
          platformVoucher: 0,
          coinOffset: 0,
          sellerCoinCashback: numOrZero(
            im.sellerCoinCashback > -1 ? row[im.sellerCoinCashback] : 0
          ),
          exactOrderFee:
            numOrZero(im.txFee > -1 ? row[im.txFee] : 0) +
            numOrZero(im.otherFee > -1 ? row[im.otherFee] : 0) +
            numOrZero(im.paymentFee > -1 ? row[im.paymentFee] : 0),
          items: [],
        };
        count++;
      }

      newOrders[orderId].items.push({
        key: `${im.prodId > -1 ? safeText(row[im.prodId]) : ""}_${
          im.optId > -1 ? safeText(row[im.optId]) : ""
        }`,
        qty: parseInt(im.qty > -1 ? row[im.qty] || 1 : 1, 10) || 1,
        name: im.prodName > -1 ? safeText(row[im.prodName]) : "未知商品",
        option: im.optName > -1 ? safeText(row[im.optName]) : "",
        activityPrice: numOrZero(
          im.activityPrice > -1 ? row[im.activityPrice] : 0
        ),
      });
    }

    setSpOrders((p) => {
      const merged = { ...p };
      Object.values(newOrders).forEach((order) => {
        /* 重匯同單：保留既有快照，避免鎖定的歷史參數被靜默清除 */
        merged[order.orderId] = withOldSnapshot(merged[order.orderId], order);
      });
      return merged;
    });
    const dates = Object.values(newOrders)
      .map((o) => String(o.date))
      .filter(Boolean)
      .sort()
      .reverse();
    if (dates.length) {
      setSY(dates[0].substring(0, 4));
      setSM(dates[0].substring(5, 7));
    }
    toast(`已匯入 ${count} 筆蝦皮訂單`, { type: "success" });
  };

  /* ─── 門市 POS 解析（交易明細＋POS訂單明細，兩份 join） ────── */
  const posHeaderIdx = (head) => {
    const idx = {};
    head.forEach((h, i) => {
      const t = safeText(h).replace(/\s/g, "");
      if (t.includes("訂單號碼")) idx.orderId = i;
      else if (t.includes("交易號碼")) idx.txId = i;
      else if (t.includes("交易類別")) idx.txType = i;
      else if (t.includes("交易日期")) idx.txDate = i;
      else if (t.includes("訂單日期")) idx.orderDate = i;
      else if (t.includes("訂單狀態")) idx.status = i;
      else if (t.includes("付款方式")) idx.payMethod = i;
      else if (t.includes("訂單合計")) idx.total = i;
      else if (t.includes("交易備註")) idx.remark = i;
      else if (t.includes("銷售人員")) idx.staff = i;
      else if (t.includes("統一編號")) idx.taxId = i;
      else if (t.includes("發票號碼")) idx.invoiceNo = i;
      else if (t.includes("商品名稱")) idx.prodName = i;
      else if (t === "選項") idx.option = i;
      else if (t.includes("商品結帳價")) idx.price = i;
      else if (t === "數量") idx.qty = i;
    });
    return idx;
  };

  const processPOSParsed = (rows, fileName) => {
    if (!Array.isArray(rows) || rows.length < 2) {
      toast("門市報表格式錯誤", { type: "error" });
      return;
    }
    /* 檔頭指紋：官網（送貨方式／購物車編號）與蝦皮（商品ID）報表丟到門市分頁要擋下來，
       否則官網 CSV 會被當 POS 訂單明細暫存、甚至跟交易明細 join */
    const rawHdr = (rows[0] || []).map((h) => safeText(h).replace(/\s/g, ""));
    if (rawHdr.some((h) => h.includes("購物車編號") || h.includes("送貨方式"))) {
      toast("這份是官網報表（有「送貨方式／購物車編號」欄），請切到官網分頁再匯入", {
        type: "error",
        duration: 8000,
      });
      return;
    }
    if (rawHdr.some((h) => h === "商品ID")) {
      toast("這份是蝦皮報表（有「商品ID」欄），請切到蝦皮分頁再匯入", {
        type: "error",
        duration: 8000,
      });
      return;
    }
    const idx = posHeaderIdx(rows[0]);
    const isTrans = idx.payMethod > -1 && idx.prodName === undefined;
    const isOrders = idx.prodName > -1;
    if (!isTrans && !isOrders) {
      toast("認不出這份門市報表（需含「付款方式」或「商品名稱」欄）", {
        type: "error",
        duration: 7000,
      });
      return;
    }
    /* 暫存區跨分頁切換保留（切去官網看一眼再回來不會丟檔），但超過 30 分鐘的半份暫存視為過期：
       避免隔天拿新匯出的一份去跟舊的一份 join */
    const STAGE_TTL = 30 * 60 * 1000;
    const otherKey = isTrans ? "orders" : "trans";
    const other = posStage.current[otherKey];
    if (other && Date.now() - (other.at || 0) > STAGE_TTL) {
      posStage.current[otherKey] = null;
      toast(
        `先前暫存的${isTrans ? "POS訂單明細" : "交易明細"}已超過 30 分鐘，已捨棄——請重新拖入兩份`,
        { type: "warning", duration: 7000 }
      );
    }
    posStage.current[isTrans ? "trans" : "orders"] = { rows, idx, at: Date.now() };
    const have = posStage.current;
    if (!have.trans || !have.orders) {
      toast(
        `已讀取${isTrans ? "交易明細" : "POS訂單明細"}（暫存 30 分鐘），請再拖入另一份${
          isTrans ? "POS訂單明細" : "交易明細"
        }`,
        { type: "info", duration: 6000 }
      );
      return;
    }

    /* 1) 交易明細 → 每張訂單的通路/發票/備註（結清為主、退貨標記） */
    const tIdx = have.trans.idx;
    const head = {};
    for (let i = 1; i < have.trans.rows.length; i++) {
      const r = have.trans.rows[i];
      if (!r || !r.length) continue;
      const oid = safeText(r[tIdx.orderId]).replace(/^#/, "");
      if (!oid) continue;
      const txType = safeText(r[tIdx.txType]);
      const isRefund = txType.includes("退貨") || txType.includes("退款");
      const rec =
        head[oid] ||
        (head[oid] = {
          payMethod: "",
          date: "",
          remark: "",
          staff: "",
          taxId: "",
          invoiceNo: "",
          refunded: false,
          refundAmt: 0,
          status: "",
          total: 0,
        });
      if (isRefund) {
        /* 退貨列金額（匯出是負數）要留下來抵銷：部分退貨的單以淨額入帳，
           全額退才整單排除（交接檔陷阱⑤） */
        rec.refunded = true;
        rec.refundAmt += Math.abs(numOrZero(r[tIdx.total]));
        continue;
      }
      rec.payMethod = safeText(r[tIdx.payMethod]) || rec.payMethod;
      /* normDate 對空字串會回 1970-01-01（truthy），舊寫法 `|| rec.date` 是死碼：
         交易日期空白時會把已對到的日期蓋成 1970。空白就保留前一列的日期。
         註：rec.total 用「覆蓋」是對的——交易明細的「訂單合計」每一列都是整單總額，
         改成累加會讓預訂＋結清的單營收翻倍（2026-09-03 用 7 份真實檔驗過） */
      const txDateRaw = safeText(r[tIdx.txDate]);
      if (txDateRaw) rec.date = normDate(txDateRaw);
      rec.remark = safeText(r[tIdx.remark]) || rec.remark;
      rec.staff =
        (tIdx.staff !== undefined ? safeText(r[tIdx.staff]) : "") || rec.staff;
      rec.taxId = safeText(r[tIdx.taxId]) || rec.taxId;
      rec.status = safeText(r[tIdx.status]) || rec.status;
      rec.total = numOrZero(r[tIdx.total]) || rec.total;
      rec.invoiceNo = safeText(r[tIdx.invoiceNo]) || rec.invoiceNo;
    }

    /* 2) 訂單明細 → 商品行 */
    const oIdx = have.orders.idx;
    const built = {};
    for (let i = 1; i < have.orders.rows.length; i++) {
      const r = have.orders.rows[i];
      if (!r || !r.length) continue;
      const oid = safeText(r[oIdx.orderId]).replace(/^#/, "");
      const nm = safeText(r[oIdx.prodName]);
      if (!oid || !nm) continue;
      const opt = safeText(r[oIdx.option]);
      const total = numOrZero(r[oIdx.total]);
      if (!built[oid])
        built[oid] = {
          orderId: oid,
          revenue: 0,
          status: safeText(r[oIdx.status]),
          orderDate: normDate(safeText(r[oIdx.orderDate])),
          items: [],
        };
      if (total > 0 && !built[oid].revenue) built[oid].revenue = total;
      /* 數量：空白才回退 1；明確的 0／小數照實保留（parseInt||1 會把 0 和 1.5 都變 1） */
      const qRaw = safeText(r[oIdx.qty]);
      const qNum = parseFloat(qRaw);
      built[oid].items.push({
        key: `${nm}_${opt}`,
        name: nm,
        option: opt,
        qty: qRaw === "" || Number.isNaN(qNum) ? 1 : qNum,
        price: numOrZero(r[oIdx.price]),
      });
    }

    /* 3) join：以訂單號碼串起來。交易有、明細沒有（早開單晚付款、明細範圍沒抓到）
       的訂單不丟——先以「無商品明細」入帳，營收照算、毛利排除；下次補匯明細會自動補齊 */
    const orphanTrans = [];
    /* 本次窗口只出現「退貨列」（沒有結清列、明細也沒抓到）的舊單：不能把先前入帳的營收覆蓋成 0，
       merge 時要用舊單重算淨額（見下方 setPosOrders） */
    const refundOnlyIds = new Set();
    const newOrders = {};
    Object.entries(head).forEach(([oid, h]) => {
      const b = built[oid];
      if (!b) orphanTrans.push(oid);
      if (!b && !h.total && h.refunded) refundOnlyIds.add(oid);
      const channel = posChannelOf(h.payMethod);
      const inv = posInvoiceOf({
        invoiceNo: h.invoiceNo,
        taxId: h.taxId,
        remark: h.remark,
        channel,
      });
      /* 退貨抵銷：淨額 = 結清 − 退貨。淨額 ≤0 ＝全額退（整單排除）；>0 ＝部分退貨，以淨額入帳
         （成本仍以全部商品計，明細裡會標示供人工複核） */
      const grossTotal = b?.revenue || h.total || 0;
      const netTotal = grossTotal - (h.refundAmt || 0);
      const fullRefund = h.refunded && netTotal <= 0;
      const partialRefund = h.refunded && netTotal > 0;
      const st = `${b?.status || h.status}${
        fullRefund ? " 已退款" : partialRefund ? " 含部分退貨" : ""
      }`;
      newOrders[oid] = {
        orderId: oid,
        date: h.date || b?.orderDate || "",
        status: st,
        channel,
        payMethod: h.payMethod,
        staff: h.staff,
        hasInvoice: inv.has,
        invoiceSrc: inv.src,
        invoiceNo: h.invoiceNo,
        taxId: h.taxId,
        remark: h.remark,
        revenue: partialRefund ? netTotal : grossTotal,
        refundAmt: h.refundAmt || 0,
        items: b ? b.items : [],
      };
    });
    /* 反向孤兒：訂單明細有、交易明細沒有（多半是還沒付款、或交易明細範圍沒抓到）。
       不能當營收入帳（沒收到錢），但也不能靜默吞掉——列出來提醒 */
    const orphanBuilt = Object.keys(built).filter((oid) => !head[oid]);

    /* 4) 自動建配方：對沒有配方的商品鍵套規則引擎 */
    const autoR = {};
    const unmatched = [];
    Object.values(newOrders).forEach((o) =>
      o.items.forEach((it) => {
        if (posRecipes[it.key] || autoR[it.key]) return;
        const lines = buildPosRecipe(it.name, it.option, components);
        if (lines && lines.length) autoR[it.key] = lines;
        else if (!unmatched.includes(it.key)) unmatched.push(it.key);
      })
    );
    if (Object.keys(autoR).length)
      setPosRecipes((p) => ({ ...autoR, ...p }));

    /* 日期空白（只有退貨列又沒有舊單可沿用）：進 1970 保底桶而不是 "unknown" 月份 doc，並點名 */
    let noDateN = 0;
    setPosOrders((p) => {
      const merged = { ...p };
      Object.values(newOrders).forEach((o) => {
        const old = merged[o.orderId];
        if (refundOnlyIds.has(o.orderId) && old) {
          /* 舊單已入帳、本次只看到退貨：以舊營收重算淨額，日期／明細／通路全部沿用舊單。
             同一張退貨列可能在重疊的匯出窗口出現兩次 → 用 max 而不是累加 */
          const gross =
            (Number(old.revenue) || 0) + (Number(old.refundAmt) || 0);
          const refund = Math.max(Number(old.refundAmt) || 0, o.refundAmt || 0);
          const net = gross - refund;
          const base = String(old.status || "").replace(/ (已退款|含部分退貨)$/, "");
          merged[o.orderId] = {
            ...old,
            refundAmt: refund,
            revenue: net > 0 ? net : 0,
            status: `${base}${net <= 0 ? " 已退款" : " 含部分退貨"}`,
          };
          return;
        }
        if (!o.date) {
          o = { ...o, date: old?.date || "1970-01-01" };
          if (!old?.date) noDateN++;
        }
        /* 這次只有交易頭、沒有明細，但先前已匯過完整明細 → 保留舊明細，別被空陣列蓋掉 */
        let next =
          !o.items.length && old?.items?.length ? { ...o, items: old.items } : o;
        /* 老闆手動改過的發票判定（invoiceOverride）重匯不能被自動判定蓋掉 */
        if (old && typeof old.invoiceOverride === "boolean")
          next = {
            ...next,
            hasInvoice: old.invoiceOverride,
            invoiceOverride: old.invoiceOverride,
            invoiceSrc: "手動設定",
          };
        merged[o.orderId] = withOldSnapshot(old, next);
      });
      return merged;
    });

    const dates = Object.values(newOrders)
      .map((o) => String(o.date))
      .filter(Boolean)
      .sort()
      .reverse();
    if (dates.length) {
      setSY(dates[0].substring(0, 4));
      setSM(dates[0].substring(5, 7));
    }
    posStage.current = { trans: null, orders: null };
    const n = Object.keys(newOrders).length;
    let msg = `已匯入 ${n} 筆門市訂單`;
    if (Object.keys(autoR).length)
      msg += `，自動對應 ${Object.keys(autoR).length} 項成本`;
    if (unmatched.length) msg += `；${unmatched.length} 項無成本資料`;
    if (orphanTrans.length)
      msg += `；${orphanTrans.length} 張交易沒有對到商品明細（已先以營收入帳、不算毛利；訂單明細日期範圍往前多抓一個月再匯一次即可補齊）`;
    if (orphanBuilt.length)
      msg += `；${orphanBuilt.length} 張訂單有商品明細但對不到交易紀錄（多半尚未付款或交易明細範圍沒抓到）——未入帳，補匯較新的交易明細即可`;
    const partialN = Object.values(newOrders).filter((o) =>
      String(o.status).includes("含部分退貨")
    ).length;
    if (partialN) msg += `；${partialN} 筆含部分退貨已以淨額入帳`;
    if (refundOnlyIds.size)
      msg += `；${refundOnlyIds.size} 張只看到退貨列的舊單已用先前入帳的營收重算淨額`;
    if (noDateN)
      msg += `；${noDateN} 張交易沒有日期（暫歸 1970-01-01，補匯含結清列的交易明細可修正）`;
    toast(msg, {
      type:
        unmatched.length || orphanTrans.length || orphanBuilt.length
          ? "warning"
          : "success",
      duration: 10000,
    });
    if (unmatched.length)
      console.warn("[POS] 無法自動對應成本的商品：", unmatched);
    if (orphanBuilt.length)
      console.warn("[POS] 有明細但無交易紀錄（未入帳）的訂單：", orphanBuilt);
    if (orphanTrans.length) console.warn("[POS] join 不到的交易：", orphanTrans);
  };

  const processFile = (f) => {
    if (!f) return;
    /* 雲端首載還沒到齊前匯入，會被首份快照整包覆蓋掉（toast 卻已說匯入成功）——先擋 */
    if (!cReady) {
      toast("雲端資料還在同步中，請等右上角同步燈變綠再匯入", {
        type: "warning",
        duration: 6000,
      });
      return;
    }
    const fname = f.name.toLowerCase();
    const isX = fname.endsWith(".xlsx") || fname.endsWith(".xls");
    if (!isX && !fname.endsWith(".csv")) {
      toast("僅支援 CSV / XLSX / XLS 檔案", { type: "error" });
      return;
    }
    const rd = new FileReader();
    const exec = (d2, x) => {
      if (x) {
        const wb = window.XLSX.read(d2, { type: "array" });
        const j = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
          header: 1,
          defval: "",
          raw: false,
        });
        const rows = j.map((r) => r.map((c) => String(c)));
        if (platform === "pos") processPOSParsed(rows, f.name);
        else if (platform === "shopline") processSLParsed(rows);
        else processSPParsed(rows);
      } else {
        const rows = parseCSV(d2);
        if (platform === "pos") processPOSParsed(rows, f.name);
        else if (platform === "shopline") processSLParsed(rows);
        else processSPParsed(rows);
      }
    };
    rd.onload = (ev) => {
      try {
        exec(ev.target.result, isX);
      } catch (err) {
        console.error("[Parse Error]", err);
        toast("報表解析失敗：" + err.message, { type: "error", duration: 8000 });
      }
    };
    rd.onerror = () => toast("檔案讀取失敗，請重試", { type: "error" });
    if (isX) {
      if (typeof window.XLSX === "undefined") {
        /* 門市一次丟兩份 xls：解析器只載一次，第二份排隊等 onload */
        if (!window.__xlsxLoading) {
          window.__xlsxLoading = new Promise((res, rej) => {
            const s2 = document.createElement("script");
            s2.src =
              "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
            s2.onload = res;
            s2.onerror = rej;
            document.head.appendChild(s2);
          });
        }
        window.__xlsxLoading.then(
          () => rd.readAsArrayBuffer(f),
          () => {
            window.__xlsxLoading = null;
            toast("XLSX 解析器載入失敗，請確認網路後重試", { type: "error" });
          }
        );
      } else rd.readAsArrayBuffer(f);
    } else rd.readAsText(f);
  };

  const handleFile = (e) => {
    /* 門市要吃兩份 xls：一次選兩個也支援，依序解析 */
    const fs = Array.from(e.target.files || []);
    fs.forEach((f, i) => setTimeout(() => processFile(f), i * 60));
    e.target.value = "";
  };

  /* ─── 期間過濾（年月／自訂區間共用） ────────────────────── */
  /* 目前平台實際有資料的月份（切平台時保留期間，但該平台沒有那個月就視為全月份，
     避免切到門市看見一片 0——期間選擇本身不變，切回去照樣是原本的月份） */
  const monthsOfPlatform = useMemo(() => {
    const src =
      platform === "pos"
        ? posOrders
        : platform === "shopline"
        ? slOrders
        : platform === "shopee"
        ? spOrders
        : null;
    if (!src) return null;
    return new Set(
      Object.values(src)
        .map((o) => String(o.date || ""))
        .filter((d) => sY === "All" || d.startsWith(sY))
        .map((d) => d.substring(5, 7))
        .filter(Boolean)
    );
  }, [platform, posOrders, slOrders, spOrders, sY]);
  /* 該平台在所選月份沒資料時退成「全月份」，避免切平台看到一片 0（2026-08-18 定）。
     **例外：當月不退**——當月本來就可能還沒匯報表，退成全年會把上個月的數字端出來，
     老闆要的是空白等匯入（2026-09-03） */
  const curYM = nowYM();
  const isCurMonth = sY === curYM.y && sM === curYM.m;
  const effM =
    sM !== "All" && !isCurMonth && monthsOfPlatform && !monthsOfPlatform.has(sM)
      ? "All"
      : sM;

  const inPeriod = useCallback(
    (d) => {
      const s = String(d || "");
      if (sY === "Custom") {
        if (range.from && s < range.from) return false;
        if (range.to && s > range.to) return false;
        return true;
      }
      if (sY !== "All" && !s.startsWith(sY)) return false;
      if (effM !== "All" && !s.startsWith(`${sY}-${effM}`)) return false;
      return true;
    },
    [sY, effM, range]
  );

  /* ─── 有效成本表：配方（組件×用量）優先，無配方回退手填 ── */
  const slEffCosts = useMemo(
    () => resolveCosts(slCosts, slRecipes, components),
    [slCosts, slRecipes, components]
  );
  const spEffCosts = useMemo(
    () => resolveCosts(spCosts, spRecipes, components),
    [spCosts, spRecipes, components]
  );
  /* 蝦皮實際採用的費率參數：淨利目標用自己的，營業費／稅率吃全公司口徑（slFp） */
  const spFpEff = useMemo(
    () => ({ ...spFp, opExpense: slFp.opExpense, tax: slFp.tax }),
    [spFp, slFp.opExpense, slFp.tax]
  );
  const posEffCosts = useMemo(
    () => resolveCosts(posCosts, posRecipes, components),
    [posCosts, posRecipes, components]
  );

  /* ─── Shopline Data Processing ────────────────────────────── */
  /* ─── 門市 POS 期間彙總 ─────────────────────────────────────── */
  const posData = useMemo(() => {
    const all = Object.values(posOrders);
    if (!all.length) return null;
    const years = [...new Set(all.map((o) => String(o.date).substring(0, 4)))]
      .filter(Boolean)
      .sort()
      .reverse();
    const t = {
      rev: 0,
      cost: 0,
      gp: 0,
      net: 0,
      opExpTotal: 0,
      taxTotal: 0,
      valid: 0,
      rawTotal: 0,
      cancelledTotal: 0,
      testTotal: 0,
      testCount: 0,
      noCostRev: 0,
      noCostCount: 0,
      estRev: 0,
      estCost: 0,
      estCount: 0,
      invoiceRev: 0,
      coveredRev: 0,
      coveredGp: 0,
      coveredNet: 0,
      coveredOp: 0,
      coveredTax: 0,
      lossCount: 0,
    };
    const byChannel = {};
    POS_CHANNELS.forEach((c) => {
      byChannel[c.key] = {
        key: c.key,
        label: c.label,
        rev: 0,
        cost: 0,
        gp: 0,
        net: 0,
        op: 0,
        tax: 0,
        orders: 0,
        noCostRev: 0,
        noCostCount: 0,
        estRev: 0,
        estCost: 0,
        estCount: 0,
        invoiceRev: 0,
        coveredRev: 0,
        coveredGp: 0,
        coveredNet: 0,
        coveredOp: 0,
        coveredTax: 0,
        lossCount: 0,
      };
    });
    const mm = {};
    const ol = [];
    all.forEach((order) => {
      if (!inPeriod(order.date)) return;
      const fin = posOrderFin(order, slFp, posEffCosts, posRatios);
      t.rawTotal += fin.gross;
      if (fin.isCanc) {
        t.cancelledTotal += fin.gross;
        return;
      }
      if (fin.isTest) {
        t.testTotal += fin.gross;
        t.testCount++;
        return;
      }
      const ch = byChannel[order.channel] || byChannel.retail;
      const excluded = !posIncluded.includes(order.channel);
      /* 通路層一律照算（表上看得到）；總計 KPI／商品表只加未被排除的通路 */
      ch.rev += fin.gross;
      ch.cost += fin.oCost;
      ch.gp += fin.gp;
      ch.net += fin.finalNet;
      ch.op += fin.opAmt;
      ch.tax += fin.txAmt;
      ch.orders++;
      if (fin.missCost) {
        ch.noCostRev += fin.gross;
        ch.noCostCount++;
      } else {
        ch.coveredRev += fin.gross;
        ch.coveredGp += fin.gp;
        ch.coveredNet += fin.finalNet;
        ch.coveredOp += fin.opAmt;
        ch.coveredTax += fin.txAmt;
        if (fin.finalNet < 0) ch.lossCount++;
        /* 這筆的成本含估算（走成本率）——毛利率/淨利率會受影響，畫面要標明 */
        if (fin.hasEst) {
          ch.estRev += fin.gross;
          ch.estCost += fin.estCost;
          ch.estCount++;
        }
      }
      if (order.hasInvoice) ch.invoiceRev += fin.gross;
      /* 訂單明細表列出「所有」有效訂單（含未計入 KPI 的通路，畫面上會灰掉標示）——
         逐筆檢視經銷單賺不賺正是這張表的用途；KPI／商品表才受計入清單限制 */
      ol.push({
        ...order,
        oCost: fin.oCost,
        gp: fin.gp,
        opx: fin.opAmt,
        taxAmt: fin.txAmt,
        net: fin.finalNet,
        missCost: fin.missCost,
        hasEst: fin.hasEst,
        estCost: fin.estCost,
        excludedCh: excluded,
        swapSuspect: posSwapSuspect(order.items),
        channelLabel: posChannelLabel(order.channel),
      });
      if (excluded) return;
      /* 商品行營收＝結帳價×數量，是「全單折扣前」的價；把折扣按比例攤回各行，
         商品毛利率才會跟訂單毛利率對得起來（例：公版華崗 400×8=3200、全單折 200、合計 3000） */
      const lineSum = order.items.reduce(
        (s, it) => s + (Number(it.price) || 0) * (it.qty || 1),
        0
      );
      /* 合計 0（免費單／全額折讓）也要攤成 0，否則商品表按原價入營收、跟 KPI 對不上 */
      const scale = lineSum > 0 ? fin.gross / lineSum : 1;
      /* 商品行結帳價全空／全 0 但訂單合計有值（贈品／人工調整單）：
         按數量等權攤，別讓這筆營收在商品表消失 */
      const qtySum = order.items.reduce((s, it) => s + (it.qty || 1), 0);
      const evenSplit = lineSum <= 0 && fin.gross > 0 && qtySum > 0;
      order.items.forEach((it) => {
        const has =
          Object.prototype.hasOwnProperty.call(it, "snapshotCost") &&
          it.snapshotCost !== null;
        const unit = has
          ? Number(it.snapshotCost) || 0
          : posItemUnit(it, posEffCosts, posRatios);
        const ir = evenSplit
          ? (fin.gross * (it.qty || 1)) / qtySum
          : (Number(it.price) || 0) * (it.qty || 1) * scale;
        const ic = unit * (it.qty || 1);
        if (!mm[it.key])
          mm[it.key] = {
            key: it.key,
            name: it.name,
            option: it.option || "標準規格",
            soldQty: 0,
            profitContribution: 0,
            totalRevenue: 0,
            totalCost: 0,
          };
        mm[it.key].soldQty += it.qty || 1;
        mm[it.key].totalRevenue += ir;
        mm[it.key].totalCost += ic;
        mm[it.key].profitContribution += ir - ic;
      });
      t.rev += fin.gross;
      t.cost += fin.oCost;
      t.gp += fin.gp;
      t.net += fin.finalNet;
      t.opExpTotal += fin.opAmt;
      t.taxTotal += fin.txAmt;
      t.valid++;
      if (fin.missCost) {
        t.noCostRev += fin.gross;
        t.noCostCount++;
      } else {
        /* 毛利率只用「成本齊全」的訂單當分母，否則會被無成本單灌成虛高 */
        t.coveredRev += fin.gross;
        t.coveredGp += fin.gp;
        t.coveredNet += fin.finalNet;
        t.coveredOp += fin.opAmt;
        t.coveredTax += fin.txAmt;
        if (fin.finalNet < 0) t.lossCount++;
        if (fin.hasEst) {
          t.estRev += fin.gross;
          t.estCost += fin.estCost;
          t.estCount++;
        }
      }
      if (order.hasInvoice) t.invoiceRev += fin.gross;
    });
    ol.sort((a, b) => String(b.date).localeCompare(String(a.date)));
    const covered = t.rev - t.noCostRev;
    return {
      years,
      orderList: ol,
      matrixList: Object.values(mm).sort((a, b) => b.soldQty - a.soldQty),
      channels: POS_CHANNELS.map((c) => byChannel[c.key])
        .filter((c) => c.rev > 0 || c.orders > 0)
        .map((c) => ({ ...c, excluded: !posIncluded.includes(c.key) })),
      /* 總覽用：六個通路一律列出（沒單的也給 0 列），且不受「計入清單」影響 */
      channelsAll: POS_CHANNELS.map((c) => ({
        ...byChannel[c.key],
        excluded: !posIncluded.includes(c.key),
      })),
      cancelledTotal: t.cancelledTotal,
      testCount: t.testCount,
      summary: {
        ...t,
        /* 毛利／淨利率一律用「成本齊全」的訂單當基數 */
        grossMargin: t.coveredRev > 0 ? t.coveredGp / t.coveredRev : 0,
        netMargin: t.coveredRev > 0 ? t.coveredNet / t.coveredRev : 0,
        aov: t.valid > 0 ? t.rev / t.valid : 0,
        costCoverage: t.rev > 0 ? covered / t.rev : 0,
        /* 覆蓋率中有多少是「用成本率估的」——毛利率／淨利率的可信度指標 */
        estShare: t.coveredRev > 0 ? t.estRev / t.coveredRev : 0,
        invoiceRate: t.rev > 0 ? t.invoiceRev / t.rev : 0,
      },
    };
  }, [posOrders, posEffCosts, posRatios, slFp, inPeriod, posIncluded]);

  /* 門市每月營收／淨利（環比同比用；門市頁＝計入的通路；總覽＝全部通路）。
     淨利只算成本齊全的單 */
  const buildPosMonthly = useCallback(
    (channelFilter) => {
      const map = {};
      Object.values(posOrders).forEach((o) => {
        const ym = String(o.date || "").substring(0, 7);
        if (ym.length < 7) return;
        if (channelFilter && !channelFilter.includes(o.channel)) return;
        const fin = posOrderFin(o, slFp, posEffCosts, posRatios);
        if (fin.isCanc || fin.isTest) return;
        if (!map[ym]) map[ym] = { rev: 0, net: 0 };
        map[ym].rev += fin.gross;
        if (!fin.missCost) map[ym].net += fin.finalNet;
      });
      return map;
    },
    [posOrders, slFp, posEffCosts, posRatios]
  );
  const posMonthly = useMemo(
    () => buildPosMonthly(posIncluded),
    [buildPosMonthly, posIncluded]
  );
  const posMonthlyAll = useMemo(() => buildPosMonthly(null), [buildPosMonthly]);

  const slData = useMemo(() => {
    const all = Object.values(slOrders);
    if (!all.length) return null;
    const years = [...new Set(all.map((o) => String(o.date || "").substring(0, 4)))]
      .sort()
      .reverse();
    const months = [
      ...new Set(
        all
          .filter((o) => sY === "All" || String(o.date || "").startsWith(sY))
          .map((o) => String(o.date || "").substring(5, 7))
      ),
    ].sort();
    const tnr = (parseFloat(slFp.targetNet) || 15) / 100;
    const mm = {};
    Object.keys(slCosts).forEach((k) => {
      /* costKey＝商品名_選項，用最後一個底線切：有 5 支商品名本身含「 _ 」
         （熱銷茶包 _ 3入組 等），從第一個底線切會把名字砍半、後半段變成假規格 */
      const cut = k.lastIndexOf("_");
      mm[k] = {
        key: k,
        name: cut >= 0 ? k.slice(0, cut) : k,
        option: (cut >= 0 ? k.slice(cut + 1) : "").trim() || "標準規格",
        soldQty: 0,
        profitContribution: 0,
        totalRevenue: 0,
        totalCost: 0,
      };
    });
    const t = {
      rev: 0,
      inbound: 0,
      pFee: 0,
      sCost: 0,
      platformFee: 0,
      cost: 0,
      net: 0,
      valid: 0,
      voucher: 0,
      opExpTotal: 0,
      taxTotal: 0,
      rawTotal: 0,
      cancelledTotal: 0,
      contributionMargin: 0,
      giftCost: 0,
      giftQty: 0,
      totalQty: 0,
      returnCount: 0,
      returnRev: 0,
      addOnRev: 0,
      addOnQty: 0,
      addOnOrders: 0,
    };
    const fl = all.filter((o) => {
      if (!inPeriod(o.date)) return false;
      const st = String(o.status || "");
      const cx = st.includes("取消") || st.includes("刪除");
      t.rawTotal += o.revenue;
      /* 全額退款（已退款金額 ≥ 營收）視同取消；部分退款在下面以淨額入帳（與門市規則同口徑） */
      const refunded = Number(o.refunded) || 0;
      if (cx || (refunded > 0 && o.revenue - refunded <= 0)) {
        t.cancelledTotal += o.revenue;
        return false;
      }
      return true;
    });
    const ol = fl
      .map((order) => {
        const refunded = Number(order.refunded) || 0;
        const ord =
          refunded > 0 ? { ...order, revenue: order.revenue - refunded } : order;
        const fin = slOrderFin(ord, slFp, slEffCosts);
        let hasAddOn = false;
        order.items.forEach((item) => {
          const cv =
            Object.prototype.hasOwnProperty.call(item, "snapshotCost") &&
            item.snapshotCost !== null
              ? Number(item.snapshotCost) || 0
              : Number(slEffCosts[item.key]) || 0;
          t.totalQty += item.qty;
          /* 本期有賣的品項：訂單上的 name/option 才是真名，蓋掉預埋時從 key 拆出來的猜測 */
          if (mm[item.key]) {
            mm[item.key].name = item.name;
            mm[item.key].option = item.option?.trim() || "標準規格";
          }
          if (!mm[item.key])
            mm[item.key] = {
              key: item.key,
              name: item.name,
              option: item.option?.trim() || "標準規格",
              soldQty: 0,
              profitContribution: 0,
              totalRevenue: 0,
              totalCost: 0,
            };
          mm[item.key].soldQty += item.qty;
          /* 商品行營收＝結帳價×數量 − 該行分攤到的全單折扣／購物金／點數（舊資料無 share＝0） */
          const ir = Math.max(
              0,
              (Number(item.price) || 0) * item.qty - (Number(item.share) || 0)
            ),
            ic = cv * item.qty;
          mm[item.key].profitContribution += ir - ic;
          mm[item.key].totalRevenue += ir;
          mm[item.key].totalCost += ic;
          if (item.isGift === true || safeText(item.name).includes("贈品")) {
            t.giftCost += ic;
            t.giftQty += item.qty;
          }
          if (item.isAddOn === true) {
            t.addOnRev += ir;
            t.addOnQty += item.qty;
            hasAddOn = true;
          }
        });
        if (hasAddOn) t.addOnOrders++;
        const { pf, sc2, plf, oc, cm, tax, opx, net } = fin;
        t.rev += ord.revenue;
        t.pFee += pf;
        t.sCost += sc2;
        t.platformFee += plf;
        t.cost += oc;
        t.contributionMargin += cm;
        t.net += net;
        t.inbound += ord.revenue - pf - sc2 - plf;
        t.voucher += order.voucherAmount;
        t.opExpTotal += opx;
        t.taxTotal += tax;
        if (order.hasReturn || refunded > 0) {
          t.returnCount++;
          t.returnRev += refunded;
        }
        t.valid++;
        return {
          ...ord,
          pFee: pf,
          sCost: sc2,
          plFee: plf,
          channelFee: pf + sc2 + plf,
          opx,
          taxAmt: tax,
          net,
          oCost: oc,
          currentOrderContribution: cm,
        };
      })
      .sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));
    const tnm = t.rev > 0 ? t.net / t.rev : 0;
    return {
      years,
      months,
      orderList: ol,
      lossCount: ol.filter((o) => o.net < 0).length,
      matrixList: Object.values(mm).sort((a, b) => b.soldQty - a.soldQty),
      summary: {
        ...t,
        trueNetMargin: tnm,
        gapVal: (tnm - tnr) * 100,
        targetNetRate: tnr,
        grossMargin: t.rev > 0 ? (t.rev - t.cost) / t.rev : 0,
        realCommissionRate:
          t.rev > 0 ? (t.pFee + t.sCost + t.platformFee) / t.rev : 0,
        voucherRate: t.rev > 0 ? t.voucher / t.rev : 0,
        giftCostRate: t.rev > 0 ? t.giftCost / t.rev : 0,
        returnRate: t.valid > 0 ? t.returnCount / t.valid : 0,
      },
    };
  }, [slOrders, sY, slFp, slCosts, slEffCosts, inPeriod]);

  /* ─── Shopee Data Processing ──────────────────────────────── */
  const spData = useMemo(() => {
    const all = Object.values(spOrders);
    if (!all.length) return null;
    const years = [
      ...new Set(
        all.map((o) => String(o.date).substring(0, 4)).filter(Boolean)
      ),
    ]
      .sort()
      .reverse();
    const months =
      sY !== "All"
        ? [
            ...new Set(
              all
                .filter((o) => String(o.date).startsWith(sY))
                .map((o) => String(o.date).substring(5, 7))
                .filter(Boolean)
            ),
          ].sort()
        : [];
    const targetNet = (parseFloat(spFpEff.targetNet) || 10) / 100;
    const prods = {};
    let tG = 0,
      tV = 0,
      tF = 0,
      tC = 0,
      tOp = 0,
      tTx = 0,
      validN = 0,
      lossN = 0,
      refundN = 0,
      refundG = 0;
    const filtered = all.filter((o) => inPeriod(o.date));
    const orderList = filtered
      .map((order) => {
        const fin = spOrderFin(order, spFpEff, spEffCosts);
        if (fin.isCanc) return null;
        if (fin.isRef) {
          refundN++;
          refundG += fin.gross;
          return null;
        }
        /* 商品行＝活動價×數量（券前）；訂單營收＝商品總價（券後）。把差額按比例攤回各行，
           商品表營收合計才等於訂單營收（與門市 scale 同口徑） */
        const lineSum = (order.items || []).reduce(
          (s, it) => s + numOrZero(it.activityPrice) * (it.qty || 1),
          0
        );
        const scale = lineSum > 0 ? fin.gross / lineSum : 1;
        (order.items || []).forEach((item) => {
          const ic =
            Object.prototype.hasOwnProperty.call(item, "snapshotCost") &&
            item.snapshotCost !== null
              ? Number(item.snapshotCost) || 0
              : Number(spEffCosts[item.key]) || 0;
          if (!prods[item.key])
            prods[item.key] = {
              key: item.key,
              name: item.name,
              option: item.option,
              soldQty: 0,
              estProfit: 0,
              totalRevenue: 0,
              totalCost: 0,
            };
          prods[item.key].soldQty += item.qty || 1;
          const ir = numOrZero(item.activityPrice) * (item.qty || 1) * scale;
          prods[item.key].totalRevenue += ir;
          prods[item.key].totalCost += ic * (item.qty || 1);
          prods[item.key].estProfit +=
            ir -
            ic * (item.qty || 1) -
            ir * (fin.opEx / 100) -
            ir * (fin.tx / 100);
        });

        tG += fin.gross;
        tV += fin.voucher;
        tF += fin.fee;
        tC += fin.oCost;
        tOp += fin.opAmt;
        tTx += fin.txAmt;
        validN++;
        if (fin.finalNet < 0) lossN++;
        return {
          ...order,
          localGross: fin.gross,
          totalOrderFee: fin.fee,
          channelFee: fin.fee,
          orderCost: fin.oCost,
          netIncome: fin.net,
          grossProfit: fin.gp,
          finalNetProfit: fin.finalNet,
          orderOpExpense: fin.opAmt,
          orderTax: fin.txAmt,
        };
      })
      .filter(Boolean);
    /* 分潤為期間層級費用，從最終淨利實扣；期間用 effM（該平台沒那個月時＝全年，與訂單過濾同口徑） */
    const comm = periodExpense(commissions, sY, effM, range);
    /* tV（賣場券）已含在 tG 裡，不再重扣 */
    const tNetPro = tG - tF - tC - tOp - tTx;
    const afterComm = tNetPro - comm;
    const netMargin = tG > 0 ? afterComm / tG : 0;

    let badge = { label: "虧損", color: "var(--dn)" };
    if (netMargin >= targetNet) {
      badge = { label: "優秀", color: "var(--up)" };
    } else if (netMargin >= targetNet * 0.6) {
      badge = { label: "穩健", color: "var(--orange)" };
    } else if (netMargin > 0) {
      badge = { label: "偏弱", color: "var(--wn)" };
    }
    return {
      years,
      months,
      orderList,
      uniqueProducts: Object.values(prods).sort(
        (a, b) => b.soldQty - a.soldQty
      ),
      s: {
        tG,
        tV,
        tF,
        tC,
        tOp,
        tTx,
        tNetPro,
        comm,
        afterComm,
        netMargin,
        targetNet,
        validN,
        lossN,
        refundN,
        refundG,
        badge,
        avgAOV: validN > 0 ? tG / validN : 0,
        avgNetPer: validN > 0 ? afterComm / validN : 0,
        grossMargin: tG > 0 ? (tG - tF - tC) / tG : 0,
        feeRate: tG > 0 ? tF / tG : 0,
        voucherRate: tG > 0 ? tV / tG : 0,
      },
    };
  }, [spOrders, sY, effM, spFpEff, spEffCosts, commissions, range, inPeriod]);

  /* ─── 每月營收/淨利彙總（環比/同比用；已扣分潤） ────────── */
  const slMonthly = useMemo(() => {
    const map = {};
    Object.values(slOrders).forEach((o) => {
      const st = String(o.status || "");
      if (st.includes("取消") || st.includes("刪除")) return;
      const refunded = Number(o.refunded) || 0;
      if (refunded > 0 && (o.revenue || 0) - refunded <= 0) return;
      const ord = refunded > 0 ? { ...o, revenue: o.revenue - refunded } : o;
      const ym = String(o.date || "").substring(0, 7);
      if (ym.length < 7) return;
      if (!map[ym]) map[ym] = { rev: 0, net: 0 };
      map[ym].rev += ord.revenue || 0;
      map[ym].net += slOrderFin(ord, slFp, slEffCosts).net;
    });
    return map;
  }, [slOrders, slFp, slEffCosts]);

  const spMonthly = useMemo(() => {
    const map = {};
    Object.values(spOrders).forEach((o) => {
      const ym = String(o.date || "").substring(0, 7);
      if (ym.length < 7) return;
      const fin = spOrderFin(o, spFpEff, spEffCosts);
      if (fin.isCanc || fin.isRef) return;
      if (!map[ym]) map[ym] = { rev: 0, net: 0 };
      map[ym].rev += fin.gross;
      map[ym].net += fin.finalNet;
    });
    /* 有分潤但當月無有效訂單時也要入帳（建立負值月份），與主視圖 afterComm 口徑一致 */
    Object.entries(commissions).forEach(([k, v]) => {
      if (v === "" || v === undefined || !/^\d{4}-\d{2}$/.test(k)) return;
      if (!map[k]) map[k] = { rev: 0, net: 0 };
      map[k].net -= Number(v) || 0;
    });
    return map;
  }, [spOrders, spFpEff, spEffCosts, commissions]);

  /* 三通路合計月表（總覽環比同比／趨勢淨利線用；門市＝全部通路、淨利只算成本齊全單） */
  const allMonthly = useMemo(() => {
    const map = {};
    [slMonthly, spMonthly, posMonthlyAll].forEach((src) =>
      Object.entries(src).forEach(([k, v]) => {
        if (!map[k]) map[k] = { rev: 0, net: 0 };
        map[k].rev += v.rev;
        map[k].net += v.net;
      })
    );
    return map;
  }, [slMonthly, spMonthly, posMonthlyAll]);

  /* 總覽用的三平台月表：包成穩定參照，否則 OverviewDashboard 的 alerts/posCompareNote
     useMemo 每次重渲染都失效，會把三平台全部訂單重掃一遍 */
  const monthlyByPlatform = useMemo(
    () => ({ sl: slMonthly, sp: spMonthly, pos: posMonthlyAll }),
    [slMonthly, spMonthly, posMonthlyAll]
  );

  /* ─── Derived state ────────────────────────────────────────── */
  const isOverview = platform === "overview";
  const isSL = platform === "shopline";
  const isPOS = platform === "pos";
  const costs = isPOS ? posCosts : isSL ? slCosts : spCosts;
  const setCosts = isPOS ? setPosCosts : isSL ? setSlCosts : setSpCosts;
  const costsEff = isPOS ? posEffCosts : isSL ? slEffCosts : spEffCosts;
  const recipes = isPOS ? posRecipes : isSL ? slRecipes : spRecipes;
  const setRecipes = isPOS ? setPosRecipes : isSL ? setSlRecipes : setSpRecipes;
  const currentData = isPOS
    ? posData
    : isSL
    ? slData
    : isOverview
    ? null
    : spData;
  const aY0 = isOverview
    ? [
        ...new Set([
          ...(slData?.years || []),
          ...(spData?.years || []),
          ...(posData?.years || []),
        ]),
      ]
        .sort()
        .reverse()
    : (isPOS ? posData : isSL ? slData : spData)?.years || [];
  /* 當年／當月一定要在下拉裡，否則「停在當月」時選單會顯示不出目前選的是什麼 */
  const aY = aY0.includes(curYM.y)
    ? aY0
    : [...aY0, curYM.y].sort().reverse();
  const aM0 = isOverview
    ? sY !== "All" && sY !== "Custom"
      ? [
          ...new Set(
            [
              ...Object.values(slOrders)
                .filter((o) => String(o.date).startsWith(sY))
                .map((o) => String(o.date).substring(5, 7)),
              ...Object.values(spOrders)
                .filter((o) => String(o.date).startsWith(sY))
                .map((o) => String(o.date).substring(5, 7)),
              ...Object.values(posOrders)
                .filter((o) => String(o.date).startsWith(sY))
                .map((o) => String(o.date).substring(5, 7)),
            ].filter(Boolean)
          ),
        ].sort()
      : []
    : isPOS
    ? [
        ...new Set(
          Object.values(posOrders)
            .filter((o) => sY === "All" || String(o.date).startsWith(sY))
            .map((o) => String(o.date).substring(5, 7))
            .filter(Boolean)
        ),
      ].sort()
    : (isSL ? slData : spData)?.months || [];
  const aM =
    sY === curYM.y && !aM0.includes(curYM.m)
      ? [...aM0, curYM.m].sort()
      : aM0;

  useEffect(() => {
    setPage(0);
    setExpandedId(null);
    setRecipeEditKey(null);
  }, [lossOnly, dSearch, orderSort, sY, sM, platform, range]);

  /* 2026-09-03 老闆定：載入不再自動跳到「資料裡的最新月份」，一律停在當月。
     當月還沒匯報表就顯示空白＋「此期間沒有任何有效訂單」提示，等匯入。
     （匯入報表後仍會跳到該批資料的最新月份，見三個 parser 結尾） */

  const slUsage = useMemo(() => buildUsage(slOrders), [slOrders]);
  const spUsage = useMemo(() => buildUsage(spOrders), [spOrders]);
  const posUsage = useMemo(() => buildUsage(posOrders), [posOrders]);

  /* 未過濾的完整商品清單（本期有賣的＋孤兒條目）。
     「未填成本 N」必須從這裡算，不能從畫面篩選後的 matrixList 算——
     否則在成本表打字搜尋就會讓 Hero 的未填提醒整個消失 */
  const matrixAll = useMemo(() => {
    const source =
      (isPOS
        ? posData?.matrixList
        : isSL
        ? slData?.matrixList
        : spData?.uniqueProducts) || [];
    const usage = isPOS ? posUsage : isSL ? slUsage : spUsage;
    /* 孤兒條目：有成本/配方但本期無銷售的商品——蝦皮先前完全看不到、無法刪除。
       用全歷史訂單找回商品名；查無紀錄則顯示 key */
    const known = new Set(source.map((p) => p.key));
    const orphanKeys = [
      ...new Set([...Object.keys(costs), ...Object.keys(recipes)]),
    ].filter((k) => !known.has(k));
    /* 官網／門市的 key 都是「商品名_選項」，可直接拆；蝦皮是 id 對 */
    const splitKey = isSL || isPOS;
    const orphans = orphanKeys.map((k) => {
      const nm = usage.nameMap[k];
      /* 查無歷史紀錄才從 key 拆；用最後一個底線切（商品名本身可能含「 _ 」） */
      const cut = splitKey ? k.lastIndexOf("_") : -1;
      return {
        key: k,
        name:
          nm?.name ||
          (splitKey ? (cut >= 0 ? k.slice(0, cut) : k) : `（查無商品名）${k}`),
        option:
          nm?.option ??
          (splitKey ? (cut >= 0 ? k.slice(cut + 1) : "").trim() || "標準規格" : ""),
        soldQty: 0,
        profitContribution: 0,
        estProfit: 0,
        totalRevenue: 0,
        totalCost: 0,
      };
    });
    const dayMs = 86400000;
    const withUsage = [...source, ...orphans].map((p) => {
      const last = usage.lastSold[p.key] || null;
      const staleDays =
        last && usage.maxDate
          ? Math.floor(
              (Date.parse(usage.maxDate) - Date.parse(last)) / dayMs
            )
          : null;
      return {
        ...p,
        lastSold: last,
        staleDays,
        neverSold: !last,
        stale: !last || (staleDays !== null && staleDays >= 180),
      };
    });
    return withUsage;
  }, [
    isSL,
    isPOS,
    slData,
    spData,
    posData,
    costs,
    recipes,
    slUsage,
    spUsage,
    posUsage,
  ]);

  const matrixList = useMemo(() => {
    const gmOf = (p) =>
      p.totalRevenue > 0
        ? (p.totalRevenue - p.totalCost) / p.totalRevenue
        : -Infinity;
    return matrixAll
      .filter((p) => !cleanupOnly || p.stale)
      .filter((p) => !soldOnly || (p.soldQty || 0) > 0)
      .filter(
        (p) =>
          !dMSearch ||
          p.name.toLowerCase().includes(dMSearch.toLowerCase()) ||
          (p.option || "").toLowerCase().includes(dMSearch.toLowerCase())
      )
      .sort((a, b) => {
        const { key, dir } = costSort;
        const m = dir === "desc" ? -1 : 1;
        if (key === "name")
          return m * String(a.name).localeCompare(String(b.name));
        if (key === "soldQty") return m * ((a.soldQty || 0) - (b.soldQty || 0));
        if (key === "profit")
          return (
            m *
            ((a.profitContribution || a.estProfit || 0) -
              (b.profitContribution || b.estProfit || 0))
          );
        if (key === "margin") return m * (gmOf(a) - gmOf(b));
        if (key === "cost")
          return (
            m *
            ((Number(costsEff[a.key]) || 0) - (Number(costsEff[b.key]) || 0))
          );
        return 0;
      });
  }, [matrixAll, dMSearch, costSort, costsEff, cleanupOnly, soldOnly]);

  /* 一鍵清理：刪除目前清單（久未使用過濾後）的手填成本與配方，可復原 */
  const bulkCleanStale = () => {
    const keys = matrixList.map((p) => p.key);
    if (!keys.length) return;
    setConfirmBox({
      title: "清除久未使用條目",
      message: `將刪除目前清單中 ${keys.length} 項的手填成本與配方。\n訂單紀錄不受影響；之後若再售出，商品會重新出現在清單（成本需重填）。\n10 秒內可在右下角通知按「復原」。`,
      danger: true,
      onOk: () => {
        const removedCosts = {};
        const removedRecipes = {};
        setCosts((p) => {
          const n = { ...p };
          keys.forEach((k) => {
            if (n[k] !== undefined) {
              removedCosts[k] = n[k];
              delete n[k];
            }
          });
          return n;
        });
        setRecipes((p) => {
          const n = { ...p };
          keys.forEach((k) => {
            if (n[k] !== undefined) {
              removedRecipes[k] = n[k];
              delete n[k];
            }
          });
          return n;
        });
        toast(`已清除 ${keys.length} 項成本/配方`, {
          type: "info",
          action: () => {
            /* 只補回「目前還是空的」鍵：使用者在這 10 秒內重填的成本不能被舊值蓋掉。
               「空」的口徑要跟 missCost／CostInput 一樣＝undefined／null／""／0——
               清除後那格若被聚焦再失焦會 commit 0，用 undefined 判會誤以為使用者填過。
               略過的鍵用 Set 收（updater 在 StrictMode 會跑兩次，計數器會重複加），
               而且 updater 要到 React 下一次 render 才執行，提示要等它跑完再讀 */
            const skipped = new Set();
            const blankCost = (v) =>
              v === undefined || v === null || v === "" || !(Number(v) > 0);
            const blankRecipe = (v) => !Array.isArray(v) || v.length === 0;
            const restore = (setter, removed, blank) =>
              setter((p) => {
                const n = { ...p };
                Object.entries(removed).forEach(([k, v]) => {
                  if (blank(n[k])) n[k] = v;
                  else skipped.add(k);
                });
                return n;
              });
            restore(setCosts, removedCosts, blankCost);
            restore(setRecipes, removedRecipes, blankRecipe);
            setTimeout(() => {
              if (skipped.size > 0)
                toast(
                  `已復原，但略過 ${skipped.size} 項（清除後已重新填過，保留新值）`,
                  { type: "warning", duration: 8000 }
                );
            }, 0);
          },
          actionLabel: "復原",
        });
      },
    });
  };

  const missCost = useMemo(() => {
    /* 口徑＝本期真的有賣的商品（不含孤兒條目），且不受成本表的搜尋／篩選影響。
       以前用 matrixList 算，在搜尋框打字就會讓 Hero 的未填提醒整個消失 */
    const sold = matrixAll.filter((p) => (p.soldQty || 0) > 0);
    const miss = sold.filter((p) => {
      /* 門市掛了成本率的泛稱鍵＝有估算成本，不算未填 */
      if (isPOS && Number(posRatios[p.key]) > 0) return false;
      const v = costsEff[p.key];
      return v === undefined || v === null || v === "" || Number(v) === 0;
    });
    return {
      total: sold.length,
      n: miss.length,
      keys: new Set(miss.map((p) => p.key)),
    };
  }, [matrixAll, costsEff, isPOS, posRatios]);

  const filteredOrders = useMemo(() => {
    if (!currentData) return [];
    const list = currentData.orderList;
    return list
      .filter((o) => {
        if (lossOnly && (isSL ? o.net >= 0 : o.finalNetProfit >= 0))
          return false;
        if (dSearch) {
          const t = dSearch.toLowerCase();
          const oid = String(o.orderId).toLowerCase();
          if (
            !oid.includes(t) &&
            !(o.items || []).some((i) =>
              String(i.name || "")
                .toLowerCase()
                .includes(t)
            )
          )
            return false;
        }
        return true;
      })
      .sort((a, b) => {
        const { key, dir } = orderSort;
        const m = dir === "desc" ? -1 : 1;
        const gv = (o) =>
          isSL
            ? {
                date: o.date,
                revenue: o.revenue,
                fee: o.channelFee,
                cost: o.oCost,
                profit: o.currentOrderContribution,
                net: o.net,
              }
            : {
                date: o.date,
                revenue: o.localGross,
                fee: o.channelFee,
                cost: o.orderCost,
                profit: o.grossProfit,
                net: o.finalNetProfit,
              };
        const av = gv(a),
          bv = gv(b);
        if (key === "date")
          return (
            m * `${a.date}${a.orderId}`.localeCompare(`${b.date}${b.orderId}`)
          );
        if (key === "revenue")
          return m * ((av.revenue || 0) - (bv.revenue || 0));
        if (key === "fee") return m * ((av.fee || 0) - (bv.fee || 0));
        if (key === "cost") return m * ((av.cost || 0) - (bv.cost || 0));
        if (key === "profit") return m * ((av.profit || 0) - (bv.profit || 0));
        if (key === "net") return m * ((av.net || 0) - (bv.net || 0));
        return 0;
      });
  }, [currentData, isSL, lossOnly, dSearch, orderSort]);

  /* 分頁夾住：資料變少（如重置本期）時避免停在超出範圍的頁碼 */
  const totalPages = Math.max(1, Math.ceil(filteredOrders.length / pageSize));
  const curPage = Math.min(page, totalPages - 1);
  const pagedOrders = useMemo(
    () => filteredOrders.slice(curPage * pageSize, (curPage + 1) * pageSize),
    [filteredOrders, curPage, pageSize]
  );

  /* 目前平台的訂單集（快照鎖定／參數檢視共用；門市與官網蝦皮同一套機制） */
  const ordersOfPlatform = isPOS ? posOrders : isSL ? slOrders : spOrders;
  const isLocked = useMemo(() => {
    if (!currentData?.orderList?.length) return false;
    const src = ordersOfPlatform;
    return currentData.orderList.every((o) => {
      const t = src[o.orderId];
      if (!t?.snapshotFeeParams) return false;
      /* 門市「交易有、明細沒對到」的單沒有品項可凍結成本，只凍參數即視為已鎖 */
      if (!t.items?.length) return isPOS;
      return t.items.every((i) =>
        Object.prototype.hasOwnProperty.call(i, "snapshotCost")
      );
    });
  }, [currentData, ordersOfPlatform, isPOS]);

  /* 目前檢視期間鎖進快照的費率參數（回答「這個月當時用的是幾 %」）；
     期間內若有多組（各月不同）會分組列出 */
  const snapParams = useMemo(() => {
    if (!currentData?.orderList?.length) return null;
    const src = ordersOfPlatform;
    const norm = (v) =>
      v === null || v === undefined || Number.isNaN(Number(v))
        ? null
        : Number(v);
    const sets = new Map();
    /* 鎖定時成本未填（snapshotCost:null）的品項不受快照保護，
       會跟著之後的成本表變動——統計出來明白告知 */
    const nullCostKeys = new Set();
    currentData.orderList.forEach((o) => {
      const rec = src[o.orderId];
      (rec?.items || []).forEach((i) => {
        if (
          Object.prototype.hasOwnProperty.call(i, "snapshotCost") &&
          i.snapshotCost === null
        )
          nullCostKeys.add(i.key);
      });
      const sp = rec?.snapshotFeeParams;
      if (!sp) return;
      const p = {
        opExpense: norm(sp.opExpense),
        tax: norm(sp.tax),
        platformFeeRate: norm(sp.platformFeeRate),
      };
      const key = `${p.opExpense}|${p.tax}|${p.platformFeeRate}`;
      if (!sets.has(key)) sets.set(key, { ...p, count: 0 });
      sets.get(key).count++;
    });
    if (!sets.size) return null;
    const list = [...sets.values()].sort((a, b) => b.count - a.count);
    return { list, mixed: sets.size > 1, nullCostCount: nullCostKeys.size };
  }, [currentData, ordersOfPlatform]);
  const pct = (v) => (v === null ? "—" : `${v}%`);

  const toggleSnap = () => {
    if (!currentData?.orderList?.length) return;
    /* 各月營業費 % 不同：限制單一月份操作，避免跨月訂單被寫入同一組（今天的）參數
       （用 effM：該平台沒資料的月份會被視為全月份，不能拿來鎖） */
    if (sY === "All" || sY === "Custom" || effM === "All") {
      toast(
        "請先切換到「單一月份」再鎖定/解除快照——各月營業費 % 不同，跨月操作會把同一組參數寫進所有月份",
        { type: "warning", duration: 8000 }
      );
      return;
    }
    const wasLocked = isLocked;
    const apply = () => {
      /* 門市沿用官網的營業費／稅率參數（無平台抽成，platformFeeRate 不適用）；
         蝦皮的營業費／稅率也吃全公司口徑（spFpEff） */
      const fp = isSL || isPOS ? slFp : spFpEff;
      const setter = isPOS ? setPosOrders : isSL ? setSlOrders : setSpOrders;
      /* functional update：以「按下確定當下」的最新訂單集為基底，
         避免確認框開啟期間遠端同步進來的訂單被過期閉包覆蓋 */
      setter((src) => {
        const no = { ...src };
        currentData.orderList.forEach((o) => {
          const tg = no[o.orderId];
          if (!tg) return;
          /* 門市無明細單也要凍參數（營業費／稅率），否則該月永遠鎖不起來；官網蝦皮維持原邏輯 */
          if (!tg.items?.length && !isPOS) return;
          if (wasLocked) {
            const nx = { ...tg };
            nx.items = (tg.items || []).map((i) => {
              const ni = { ...i };
              delete ni.snapshotCost;
              delete ni.snapshotEst;
              return ni;
            });
            delete nx.snapshotFeeParams;
            no[o.orderId] = nx;
          } else {
            /* 鎖定＝「只補空缺」：已有快照參數的單沿用原參數（各月營業費不同，重匯後補鎖
               不能把歷史 % 蓋成今天的）；已凍結的品項不動，只補沒有快照的品項。
               要整月換參數走「解除 → 改 → 鎖定」。淨利目標是對照線，不進快照。 */
            no[o.orderId] = {
              ...tg,
              snapshotFeeParams: tg.snapshotFeeParams || {
                platformFeeRate: isPOS ? null : numOrNull(fp.platformFeeRate),
                opExpense: numOrNull(fp.opExpense),
                tax: numOrNull(fp.tax),
              },
              items: (tg.items || []).map((i) => {
                if (
                  Object.prototype.hasOwnProperty.call(i, "snapshotCost") &&
                  i.snapshotCost !== null
                )
                  return i;
                /* 凍結「有效成本」；泛稱鍵凍結「該列單價×成本率」並標 snapshotEst；
                   成本 0／空白一律凍 null（＝不受保護，走既有「未填」警示），不能凍成 0 當作有成本 */
                const eff = Number(costsEff[i.key]);
                const rUnit =
                  isPOS && !(eff > 0) ? posRatioUnit(i, costsEff, posRatios) : null;
                return {
                  ...i,
                  snapshotCost: eff > 0 ? eff : rUnit,
                  snapshotEst: !(eff > 0) && rUnit !== null,
                };
              }),
            };
          }
        });
        return no;
      });
      toast(wasLocked ? "已解除快照" : "已鎖定本期成本快照", {
        type: "success",
      });
    };
    const fp = isSL || isPOS ? slFp : spFpEff;
    const curLine = `營業費 ${fp.opExpense}%・稅率 ${fp.tax}%${
      isSL ? `・系統費 ${fp.platformFeeRate}%` : ""
    }${isPOS ? "（門市：稅只課有開發票的單）" : ""}`;
    const snapLine =
      snapParams && !snapParams.mixed
        ? `營業費 ${pct(snapParams.list[0].opExpense)}・稅率 ${pct(
            snapParams.list[0].tax
          )}${isSL ? `・系統費 ${pct(snapParams.list[0].platformFeeRate)}` : ""}`
        : snapParams
        ? "各月不同（見側欄明細）"
        : "—";
    setConfirmBox({
      title: wasLocked ? "解除快照" : "鎖定快照",
      message: wasLocked
        ? `原快照參數：${snapLine}\n\n解除後，本期訂單將改回以「目前」側欄參數即時計算（${curLine}）。\n\n若是要修正本期的 %：解除後先到側欄改好參數，再重新鎖定。`
        : `鎖定後，本期訂單將固定採用目前參數：\n${curLine}\n\n之後修改側欄參數不會影響本期（「淨利目標」僅為對照線，不隨快照鎖定）。若本期實際營業費 % 還不確定，可先鎖定，之後「解除 → 改 % → 重新鎖定」修正。${
          snapParams
            ? `\n\nℹ 本期已有 ${snapParams.list.reduce((s, p) => s + p.count, 0)} 筆帶快照（${snapLine}）——本次只補「還沒鎖」的訂單／品項，既有快照的參數與成本不會被覆蓋。`
            : ""
        }${
          missCost.n > 0
            ? `\n\n⚠ 目前有 ${missCost.n} 項商品成本未填：未填項不受快照保護，之後改成本表仍會影響本期數字，建議先補成本再鎖定。`
            : ""
        }`,
      danger: wasLocked,
      onOk: apply,
    });
  };

  const expC = () => {
    /* v2 備份：手填成本＋本平台配方＋共用原料庫一起打包 */
    const pfTag = isPOS ? "pos" : isSL ? "sl" : "sp";
    const bundle = {
      __v: 2,
      platform: pfTag,
      costs,
      recipes,
      components,
      ...(isPOS ? { ratios: posRatios } : {}),
    };
    const b = new Blob([JSON.stringify(bundle, null, 2)], {
      type: "application/json",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(b);
    a.download = `${pfTag}_costs_${
      new Date().toISOString().split("T")[0]
    }.json`;
    a.click();
  };
  const impC = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    /* 與報表匯入、成本編輯同一道閘門：首載四份快照到齊前不收寫入，
       否則首份 meta 快照一到就把剛還原的整張表蓋回雲端版 */
    if (!cReady) {
      toast("雲端同步中，請等右上角同步燈變綠再還原", { type: "warning" });
      e.target.value = "";
      return;
    }
    const r = new FileReader();
    r.onload = (ev) => {
      try {
        const parsed = JSON.parse(ev.target.result);
        const isPlainObj = (v) =>
          v && typeof v === "object" && !Array.isArray(v);
        if (!isPlainObj(parsed)) throw new Error("not-an-object");
        /* 版本守門：未來的 __v:3 用舊版工具還原會把 __v/platform/costs 當成商品鍵寫進成本表 */
        if (parsed.__v !== undefined && parsed.__v !== 2) {
          toast(
            `這份備份是 v${parsed.__v} 格式，這個版本的工具讀不懂，已中止（請用產生它的版本還原）`,
            { type: "error", duration: 9000 }
          );
          return;
        }
        const pfTag = isPOS ? "pos" : isSL ? "sl" : "sp";
        const pfName = (t) =>
          t === "pos" ? "門市" : t === "sl" ? "官網" : "蝦皮";
        if (parsed.__v === 2) {
          /* 形狀檢查：壞檔會產生垃圾成本鍵，或讓 compGroups 對 null 取屬性直接炸掉整頁 */
          const costsIn = isPlainObj(parsed.costs) ? parsed.costs : {};
          const recipesIn = isPlainObj(parsed.recipes) ? parsed.recipes : {};
          const compsIn = isPlainObj(parsed.components) ? parsed.components : {};
          const ratiosIn = isPlainObj(parsed.ratios) ? parsed.ratios : {};
          const comps = Object.fromEntries(
            Object.entries(compsIn).filter(
              ([, v]) => isPlainObj(v) && typeof v.name === "string"
            )
          );
          const recs = Object.fromEntries(
            Object.entries(recipesIn).filter(
              ([, v]) =>
                Array.isArray(v) && v.every((l) => isPlainObj(l) && l.compId)
            )
          );
          /* 組件是三平台共用：還原舊備份會讓另外兩個平台的配方成本一起變動。
             會變動的兩種：單價不同的、以及「本機已刪但備份裡還有、且還有配方在引用」的
             （還原會把它復活，引用它的配方成本從 0 跳回單價） */
          const refIds = new Set();
          [slRecipes, spRecipes, posRecipes].forEach((rs) =>
            Object.values(rs || {}).forEach((ls) =>
              (ls || []).forEach((l) => refIds.add(l.compId))
            )
          );
          const compChanged = Object.entries(comps).filter(([id, v]) =>
            components[id]
              ? Number(components[id].price) !== Number(v.price)
              : refIds.has(id)
          ).length;
          const has = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
          const doRestore = () => {
            /* 「匯入前」的值在按確定的當下、從最新 state 裡抓（functional updater），
               不是選檔那一輪 render 的閉包；而且只記「這次匯入會覆蓋的鍵」——
               復原時只退這些鍵，別的鍵（可能已被別台改過）不動 */
            const prev = { costs: {}, recipes: {}, components: {}, ratios: {} };
            const snap = (bucket, p, incoming) => {
              Object.keys(incoming).forEach((k) => {
                if (has(p, k)) bucket[k] = p[k];
              });
            };
            setCosts((p) => {
              snap(prev.costs, p, costsIn);
              return { ...p, ...costsIn };
            });
            setRecipes((p) => {
              snap(prev.recipes, p, recs);
              return { ...p, ...recs };
            });
            setComponents((p) => {
              snap(prev.components, p, comps);
              return { ...p, ...comps };
            });
            if (isPOS)
              setPosRatios((p) => {
                snap(prev.ratios, p, ratiosIn);
                return { ...p, ...ratiosIn };
              });
            const revert = (setter, incoming, before) =>
              setter((p) => {
                const n = { ...p };
                Object.keys(incoming).forEach((k) => {
                  if (has(before, k)) n[k] = before[k];
                  else delete n[k];
                });
                return n;
              });
            toast("成本＋配方＋原料庫還原成功", {
              type: "success",
              action: () => {
                revert(setCosts, costsIn, prev.costs);
                revert(setRecipes, recs, prev.recipes);
                revert(setComponents, comps, prev.components);
                if (isPOS) revert(setPosRatios, ratiosIn, prev.ratios);
                toast("已還原到匯入前的狀態（只退這次匯入的鍵）", { type: "info" });
              },
              actionLabel: "復原",
              duration: 12000,
            });
          };
          setConfirmBox({
            title: "還原成本備份",
            message:
              (parsed.platform && parsed.platform !== pfTag
                ? `⚠ 這是「${pfName(parsed.platform)}」的備份，但目前在「${pfName(
                    pfTag
                  )}」頁——還原會把別平台的商品鍵灌進這裡。\n\n`
                : "") +
              `將覆蓋：手填成本 ${Object.keys(costsIn).length} 項、配方 ${
                Object.keys(recs).length
              } 項、原料庫 ${Object.keys(comps).length} 項${
                isPOS ? `、成本率 ${Object.keys(ratiosIn).length} 項` : ""
              }。\n${
                compChanged > 0
                  ? `其中 ${compChanged} 個組件單價會變動——原料庫是三平台共用，官網／蝦皮／門市所有掛到這些組件的配方成本會一起改變。\n`
                  : ""
              }\n匯入後 12 秒內可在右下角通知按「復原」。`,
            danger: parsed.platform && parsed.platform !== pfTag,
            onOk: doRestore,
          });
        } else {
          /* 舊版備份＝純成本表：值必須全是數字，否則是誤選的無關 JSON */
          const vals = Object.values(parsed);
          if (!vals.length || !vals.every((v) => Number.isFinite(Number(v)))) {
            toast(
              "這份 JSON 不像成本備份（內容不是「商品鍵→數字」），已中止",
              { type: "error", duration: 9000 }
            );
            return;
          }
          setConfirmBox({
            title: "還原舊版成本表",
            message: `將覆蓋目前「${pfName(pfTag)}」的手填成本 ${
              vals.length
            } 項（舊版備份不含配方與原料庫）。`,
            danger: false,
            onOk: () => {
              /* 同 v2：匯入前的值從最新 state 抓、復原只退這次匯入的鍵 */
              const before = {};
              setCosts((p) => {
                Object.keys(parsed).forEach((k) => {
                  if (Object.prototype.hasOwnProperty.call(p, k)) before[k] = p[k];
                });
                return { ...p, ...parsed };
              });
              toast("成本資料匯入成功", {
                type: "success",
                action: () =>
                  setCosts((p) => {
                    const n = { ...p };
                    Object.keys(parsed).forEach((k) => {
                      if (Object.prototype.hasOwnProperty.call(before, k)) n[k] = before[k];
                      else delete n[k];
                    });
                    return n;
                  }),
                actionLabel: "復原",
                duration: 12000,
              });
            },
          });
        }
      } catch {
        toast("匯入失敗：請選擇本工具「備份」產生的 JSON 檔", {
          type: "error",
        });
      }
    };
    r.readAsText(f);
    e.target.value = "";
  };
  /* ─── 原料庫與配方操作 ────────────────────────────────────── */
  const compUsage = useMemo(() => {
    const u = {};
    /* 算「幾個配方用到」而不是「幾行用到」：同一配方重複加同一組件會有兩行，
       刪除警告若照行數會高報（「被 2 個配方使用中」但其實只有 1 個） */
    [slRecipes, spRecipes, posRecipes].forEach((rs) =>
      Object.values(rs || {}).forEach((lines) =>
        new Set((lines || []).map((l) => l.compId)).forEach((id) => {
          u[id] = (u[id] || 0) + 1;
        })
      )
    );
    return u;
  }, [slRecipes, spRecipes, posRecipes]);
  const addComponent = () => {
    const name = safeText(newComp.name);
    const price = Number(newComp.price);
    if (!name) {
      toast("請輸入組件名稱（例：高山茶 150g、禮盒A、提袋）", {
        type: "warning",
      });
      return;
    }
    setComponents((p) => ({
      ...p,
      [newCompId()]: {
        name,
        price: Number.isFinite(price) ? price : 0,
        cat: safeText(newComp.cat),
      },
    }));
    setNewComp({ name: "", price: "", cat: "" });
  };
  const commitCompField = useCallback((compId, field, value) => {
    setComponents((p) => {
      const c = p[compId];
      if (!c) return p;
      if (field === "price") {
        const n = Number(value);
        return {
          ...p,
          [compId]: { ...c, price: Number.isFinite(n) ? n : 0 },
        };
      }
      if (field === "cat") {
        return { ...p, [compId]: { ...c, cat: safeText(value) } };
      }
      const name = safeText(value);
      return name ? { ...p, [compId]: { ...c, name } } : p;
    });
  }, []);
  /* 原料庫分組（未分類排最後）與既有分類清單（datalist 建議用） */
  const compGroups = useMemo(() => {
    const g = {};
    Object.entries(components).forEach(([id, c]) => {
      const cat = safeText(c.cat) || "未分類";
      (g[cat] = g[cat] || []).push([id, c]);
    });
    return Object.entries(g).sort((a, b) =>
      a[0] === "未分類"
        ? 1
        : b[0] === "未分類"
        ? -1
        : a[0].localeCompare(b[0], "zh-Hant")
    );
  }, [components]);
  const compCats = useMemo(
    () =>
      [
        ...new Set(
          Object.values(components)
            .map((c) => safeText(c.cat))
            .filter(Boolean)
        ),
      ].sort((a, b) => a.localeCompare(b, "zh-Hant")),
    [components]
  );
  const deleteComponent = (compId) => {
    const used = compUsage[compId] || 0;
    setConfirmBox({
      title: "刪除成本組件",
      message:
        `確定刪除「${components[compId]?.name || compId}」？` +
        (used > 0
          ? `\n⚠ 它被 ${used} 個配方使用中——刪除後那些配方會少掉這個組件的成本（以 0 計），建議先到配方裡移除引用。`
          : ""),
      danger: true,
      onOk: () =>
        setComponents((p) => {
          const n = { ...p };
          delete n[compId];
          return n;
        }),
    });
  };
  const setRecipeLine = (key, idx, patch) =>
    setRecipes((p) => {
      const lines = [...(p[key] || [])];
      if (patch === null) lines.splice(idx, 1);
      else lines[idx] = { ...lines[idx], ...patch };
      const n = { ...p };
      if (lines.length) n[key] = lines;
      else delete n[key];
      return n;
    });
  const addRecipeLine = (key, compId) => {
    if (!compId) return;
    setRecipes((p) => ({
      ...p,
      [key]: [...(p[key] || []), { compId, qty: 1 }],
    }));
  };
  const removeRecipe = (key) =>
    setConfirmBox({
      title: "移除配方",
      message:
        "移除後此商品改回「手填單位成本」（右欄輸入框），原手填值若存在會恢復生效。",
      danger: false,
      onOk: () =>
        setRecipes((p) => {
          const n = { ...p };
          delete n[key];
          return n;
        }),
    });

  /* ─── 匯出本期損益報表（CSV，含 BOM 供 Excel 直開） ──────── */
  const expReport = () => {
    if (!currentData) return;
    const slD0 = slData?.summary;
    const spS0 = spData?.s;
    const pl =
      sY === "Custom"
        ? `${range.from || "起"}~${range.to || "迄"}`
        : sY === "All"
        ? "歷年"
        : effM === "All"
        ? `${sY}年`
        : `${sY}-${effM}`;
    const r0 = Math.round;
    const rows = [];
    /* 明細列跟畫面同步套用「虧損篩選／搜尋」；彙總列維持全期間口徑並明確標註 */
    const filterBits = [];
    if (lossOnly) filterBits.push("僅虧損單");
    if (dSearch) filterBits.push(`搜尋「${dSearch}」`);
    const detailNote = filterBits.length
      ? `${filterBits.join("、")}，共 ${filteredOrders.length} 筆（上方彙總仍為全期間）`
      : `全期間共 ${filteredOrders.length} 筆`;
    if (isSL && slD0) {
      rows.push(
        ["平台", "官網"],
        ["期間", pl],
        ["有效訂單", slD0.valid],
        ["營收", r0(slD0.rev)],
        ["商品成本", r0(slD0.cost)],
        ["通路費用（金流+物流+系統）", r0(slD0.pFee + slD0.sCost + slD0.platformFee)],
        ["營業費（含廣告）", r0(slD0.opExpTotal)],
        ["稅賦", r0(slD0.taxTotal)],
        ["最終淨利", r0(slD0.net)],
        ["淨利率", (slD0.trueNetMargin * 100).toFixed(2) + "%"],
        ["平均客單價", r0(slD0.rev / (slD0.valid || 1))],
        ["加購品營收（含於營收）", r0(slD0.addOnRev)],
        [
          "加購滲透率",
          ((slD0.valid > 0 ? slD0.addOnOrders / slD0.valid : 0) * 100).toFixed(
            1
          ) + "%",
        ],
        []
      );
      rows.push(["明細範圍", detailNote]);
      rows.push([
        "日期",
        "單號",
        "狀態",
        "營收",
        "通路費用",
        "商品成本",
        "通路後毛利",
        "營業費",
        "稅賦",
        "單筆淨利",
      ]);
      filteredOrders.forEach((o) =>
        rows.push([
          o.date,
          o.orderId,
          o.status,
          r0(o.revenue),
          r0(o.channelFee),
          r0(o.oCost),
          r0(o.currentOrderContribution),
          r0(o.opx),
          r0(o.taxAmt),
          r0(o.net),
        ])
      );
    } else if (spS0) {
      rows.push(
        ["平台", "蝦皮"],
        ["期間", pl],
        ["有效訂單", spS0.validN],
        ["營收（含補貼還原）", r0(spS0.tG)],
        ["商品成本", r0(spS0.tC)],
        ["通路費用（手續費+金流+蝦幣回饋）", r0(spS0.tF)],
        ["賣場優惠券（已含於營收，不另扣）", r0(spS0.tV)],
        ["營業費（含廣告）", r0(spS0.tOp)],
        ["稅賦", r0(spS0.tTx)],
        ["分潤", r0(spS0.comm)],
        ["最終淨利", r0(spS0.afterComm)],
        ["淨利率", (spS0.netMargin * 100).toFixed(2) + "%"],
        ["平均客單價", r0(spS0.avgAOV)],
        ["單筆平均淨利", r0(spS0.avgNetPer)],
        []
      );
      rows.push(["明細範圍", detailNote]);
      rows.push([
        "日期",
        "單號",
        "狀態",
        "營收",
        "通路費用",
        "商品成本",
        "通路後毛利",
        "營業費",
        "稅賦",
        "單筆淨利",
      ]);
      filteredOrders.forEach((o) =>
        rows.push([
          o.date,
          o.orderId,
          o.status,
          r0(o.localGross),
          r0(o.channelFee),
          r0(o.orderCost),
          r0(o.grossProfit),
          r0(o.orderOpExpense),
          r0(o.orderTax),
          r0(o.finalNetProfit),
        ])
      );
    }
    if (!rows.length) return;
    downloadCsv(rows, `${isSL ? "官網" : "蝦皮"}損益報表_${pl}.csv`);
    toast("報表已匯出", { type: "success" });
  };

  const handleComm = useCallback(
    (key, value) => monthlyUpd(setCommissions, key, value),
    []
  );
  const commitCost = useCallback(
    (key, n) => {
      if (!cReady) {
        toast("雲端同步中，請稍候再填成本", { type: "warning" });
        return;
      }
      /* 門市要寫 posCosts——原本只分 isSL/否，門市手填會寫進蝦皮成本表（打完就不見） */
      const setter = isPOS ? setPosCosts : isSL ? setSlCosts : setSpCosts;
      setter((pr) => ({ ...pr, [key]: n }));
    },
    [isSL, isPOS, cReady, toast]
  );
  /* 門市發票判定手動覆寫（逐單；reset＝恢復自動判定）。稅是即時算的，不受快照凍結 */
  const togglePosInvoice = useCallback((orderId, mode) => {
    setPosOrders((p) => {
      const o = p[orderId];
      if (!o) return p;
      if (mode === "reset") {
        const inv = posInvoiceOf({
          invoiceNo: o.invoiceNo,
          taxId: o.taxId,
          remark: o.remark,
          channel: o.channel,
        });
        const n = { ...o, hasInvoice: inv.has, invoiceSrc: inv.src };
        delete n.invoiceOverride;
        return { ...p, [orderId]: n };
      }
      const v = !o.hasInvoice;
      return {
        ...p,
        [orderId]: { ...o, hasInvoice: v, invoiceOverride: v, invoiceSrc: "手動設定" },
      };
    });
  }, []);

  const commitFp = useCallback(
    (field, v) => {
      if (!cReady) {
        toast("雲端同步中，請稍候再改參數", { type: "warning" });
        return;
      }
      /* 門市沿用官網那組參數（營業費／稅率是全公司共用口徑） */
      const setter = isSL || isPOS ? setSlFp : setSpFp;
      setter((p) => ({ ...p, [field]: v }));
      /* 內部營業費／預估稅率全通路一致（老闆 2026-08-31 明示）：唯一來源是 slFp，
         只有官網頁能改；這裡把值鏡射進 spFp，讓備份檔／舊版讀到的仍是同一組數字。
         已鎖定期間不受影響（吃快照）。 */
      if ((field === "opExpense" || field === "tax") && (isSL || isPOS)) {
        setSpFp((p) => (p[field] === v ? p : { ...p, [field]: v }));
      }
    },
    [isSL, isPOS, cReady, toast]
  );

  /* ── 未填成本跳轉 helper ── */
  const jumpToFirstMissCost = () => {
    /* 先清空所有篩選，避免缺成本項目被篩掉導致跳轉落空（延遲需大於搜尋 debounce） */
    setMSearch("");
    setSoldOnly(false);
    setCleanupOnly(false);
    setTimeout(() => {
      /* 記住當初上色的那一列：使用者 2 秒內把它填好後 ref 會移到下一個未填列，
         清框若重讀 ref 會把下一列的框拿掉、原本那列的框留著 */
      const el = firstMissRef.current;
      if (!el) return;
      el.scrollIntoView({ behavior: "smooth", block: "center" });
      el.style.outline = "2px solid var(--wn)";
      el.style.outlineOffset = "2px";
      setTimeout(() => {
        el.style.outline = "none";
      }, 2000);
    }, 320);
  };

  const slD = slData?.summary;
  const spS = spData?.s;

  const accentColor = isPOS
    ? "var(--pos-accent)"
    : isSL
    ? "var(--accent)"
    : "var(--sp-accent)";
  const accentDim = isPOS
    ? "var(--pos-accent-dim)"
    : isSL
    ? "var(--accent-dim)"
    : "var(--sp-accent-dim)";
  const accentBdr = isPOS
    ? "var(--pos-accent-bdr)"
    : isSL
    ? "var(--accent-bdr)"
    : "var(--sp-accent-bdr)";

  /* ─── 泛稱品項成本率卡（門市專用）─────────────────────────────
     8/18 商品目錄上線前員工只點得到「茶葉」這種泛稱，回推不出單位成本，
     改用「成本率」估：成本＝該列營收×率。建議值＝本期成本齊全訂單的實測平均。 */
  const posRatioSuggest = useMemo(() => {
    if (!isPOS || !posData) return null;
    /* 只用「成本全部來自配方／手填」的訂單當基準——把估算單算進來會自我循環：
       套了率的單變成 covered，下次算出來的「實測平均」就含自己的估算值。
       同理只取現場零售（經銷成本率天生高很多，混進來會拉高估算）。 */
    const agg = { retail: { rev: 0, cost: 0, n: 0 }, all: { rev: 0, cost: 0, n: 0 } };
    (posData.orderList || []).forEach((o) => {
      if (o.missCost || o.hasEst || !(o.revenue > 0)) return;
      const add = (a) => {
        a.rev += o.revenue;
        a.cost += o.oCost;
        a.n++;
      };
      add(agg.all);
      if (o.channel === "retail") add(agg.retail);
    });
    const pick =
      agg.retail.n >= 5 ? { ...agg.retail, scope: "現場零售" } : { ...agg.all, scope: "全部通路" };
    if (!(pick.rev > 0)) return null;
    const rate = (pick.cost / pick.rev) * 100;
    if (!(rate > 0 && rate < 100)) return null;
    /* 對照組：經銷單的實測成本率（頁尾說明用，不寫死數字） */
    const d = { rev: 0, cost: 0, n: 0 };
    (posData.orderList || []).forEach((o) => {
      if (o.missCost || o.hasEst || o.channel !== "dealer" || !(o.revenue > 0)) return;
      d.rev += o.revenue;
      d.cost += o.oCost;
      d.n++;
    });
    return {
      rate: Math.round(rate * 10) / 10,
      scope: pick.scope,
      n: pick.n,
      dealer: d.rev > 0 ? Math.round((d.cost / d.rev) * 1000) / 10 : null,
      dealerN: d.n,
    };
  }, [isPOS, posData]);
  const posRatioRows = useMemo(() => {
    if (!isPOS) return [];
    return matrixList
      .filter(
        (p) =>
          Number(posRatios[p.key]) > 0 ||
          (!(posRecipes[p.key] || []).length && !(Number(posCosts[p.key]) > 0))
      )
      .sort((a, b) => (b.totalRevenue || 0) - (a.totalRevenue || 0));
  }, [isPOS, matrixList, posRatios, posRecipes, posCosts]);
  const commitRatio = useCallback(
    (key, v) => {
      if (!cReady) {
        toast("雲端同步中，請稍候再設定成本率", { type: "warning" });
        return;
      }
      setPosRatios((p) => {
        const n = { ...p };
        if (!(v > 0)) delete n[key];
        else n[key] = v;
        return n;
      });
    },
    [cReady, toast]
  );
  const posRatioCard =
    isPOS && posRatioRows.length > 0 ? (
      <div
        className="f4"
        style={{
          background: "var(--s1)",
          border: "1px solid var(--s3)",
          borderRadius: 16,
          padding: 24,
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "flex-start",
            justifyContent: "space-between",
            gap: 12,
            flexWrap: "wrap",
            marginBottom: 12,
          }}
        >
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, color: "var(--t1)" }}>
              泛稱品項・成本率估算
            </div>
            <div
              style={{
                fontSize: 11,
                color: "var(--t3)",
                marginTop: 4,
                lineHeight: 1.7,
              }}
            >
              沒有規格可查的品項（例：「茶葉」）用成本率估：成本＝該筆營收 ×
              率。填了率就不算未填成本，明細會標「估」。
            </div>
          </div>
          {posRatioSuggest && (
            <Btn
              v="primary"
              onClick={() =>
                setPosRatios((p) => {
                  const n = { ...p };
                  let filled = 0;
                  /* 只填「還沒有率」的列——不覆蓋老闆手調過的數字 */
                  posRatioRows.forEach((r) => {
                    if (!(Number(n[r.key]) > 0)) {
                      n[r.key] = posRatioSuggest.rate;
                      filled++;
                    }
                  });
                  toast(
                    filled > 0
                      ? `已填入 ${filled} 項（已手調的不動）`
                      : "每一項都已經有率了，未變更",
                    { type: filled > 0 ? "success" : "info" }
                  );
                  return n;
                })
              }
              title={`基準＝本期「成本全部來自配方／手填」的${posRatioSuggest.scope}訂單 ${posRatioSuggest.n} 筆，不含任何估算值`}
            >
              空白處填入實測平均 {posRatioSuggest.rate}%
            </Btn>
          )}
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={{ ...th, textAlign: "left" }}>商品（本期無成本來源）</th>
                <th style={th}>本期營收</th>
                <th style={th}>成本率 %</th>
                <th style={th}>估算成本</th>
                <th style={th}>估算毛利率</th>
              </tr>
            </thead>
            <tbody>
              {posRatioRows.map((p) => {
                const rate = Number(posRatios[p.key]) || 0;
                const rev = p.totalRevenue || 0;
                return (
                  <tr key={p.key} style={{ borderTop: "1px solid var(--s3)" }}>
                    <td style={{ ...td2, textAlign: "left" }}>
                      <div style={{ fontWeight: 600, color: "var(--t1)" }}>
                        {p.name}
                      </div>
                      <div style={{ fontSize: 10, color: "var(--t4)" }}>
                        {p.option || "標準規格"}　{p.soldQty || 0} 件
                      </div>
                    </td>
                    <td style={{ ...td2, fontFamily: mono }}>{fmt$(rev)}</td>
                    <td style={td2}>
                      <CostInput
                        costKey={p.key}
                        label={`${p.name} 成本率（%）`}
                        value={posRatios[p.key]}
                        miss={!(rate > 0)}
                        onCommit={commitRatio}
                      />
                    </td>
                    {/* 估算成本／毛利率直接取引擎算出來的值（p.totalCost 已含折扣攤分口徑），
                        不在卡片裡另算一次，避免同一頁出現兩個不一樣的毛利率 */}
                    <td style={{ ...td2, fontFamily: mono }}>
                      {rate > 0 ? fmt$(p.totalCost || 0) : "—"}
                    </td>
                    <td
                      style={{
                        ...td2,
                        fontFamily: mono,
                        fontWeight: 700,
                        color:
                          rate > 0 && rev > 0
                            ? posGmColor((rev - (p.totalCost || 0)) / rev)
                            : "var(--t4)",
                      }}
                    >
                      {rate > 0 && rev > 0
                        ? fmtP((rev - (p.totalCost || 0)) / rev)
                        : "—"}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div
          style={{
            fontSize: 10,
            color: "var(--t4)",
            marginTop: 10,
            lineHeight: 1.7,
          }}
        >
          率只在「沒有配方也沒有手填成本」時生效——之後補了配方或成本，配方優先、率自動失效（不必刪）。
          估算成本會跟著鎖定快照凍結。
          {posRatioSuggest && (
            <>
              　建議值基準＝本期「成本全部來自配方／手填」的{posRatioSuggest.scope}訂單{" "}
              {posRatioSuggest.n} 筆，不含任何估算值。
              {posRatioSuggest.dealer !== null &&
                `　對照：經銷單實測成本率 ${posRatioSuggest.dealer}%（${posRatioSuggest.dealerN} 筆）——泛稱鍵若混了經銷單，估出來會偏樂觀。`}
            </>
          )}
          　有正規品名＋規格的品項請掛配方，不要用率——率是給「查不出賣了什麼」的泛稱鍵用的。
        </div>
      </div>
    ) : null;

  /* ─── 商品成本資料庫卡（官網／蝦皮／門市三處共用：手填成本＋配方編輯器＋原料庫＋備份還原） ── */
  const costMatrixCard = (
                <div
                  className="f4"
                  style={{
                    background: "var(--s1)",
                    border: "1px solid var(--s3)",
                    borderRadius: 16,
                    padding: 24,
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      flexWrap: "wrap",
                      justifyContent: "space-between",
                      gap: 12,
                      alignItems: "center",
                      marginBottom: 14,
                    }}
                  >
                    <div
                      style={{ display: "flex", alignItems: "center", gap: 8 }}
                    >
                      <Package size={16} color="var(--t3)" />
                      <span style={{ fontSize: 14, fontWeight: 700 }}>
                        商品成本資料庫
                      </span>
                      <span style={{ fontSize: 11, color: "var(--t3)" }}>
                        顯示 {matrixList.length} 項
                      </span>
                    </div>
                    <div
                      style={{ display: "flex", gap: 6, alignItems: "center" }}
                    >
                      <Btn onClick={expC}>
                        <Download size={12} /> 備份
                      </Btn>
                      <Btn v="primary" onClick={() => cRef.current?.click()}>
                        <UploadCloud size={12} /> 還原
                      </Btn>
                      <input
                        ref={cRef}
                        type="file"
                        accept=".json"
                        onChange={impC}
                        style={{ display: "none" }}
                      />
                    </div>
                  </div>
                  {/* ── 成本組件（原料庫）：兩平台共用單價 ── */}
                  <div
                    style={{
                      border: "1px solid var(--s3)",
                      borderRadius: 12,
                      padding: "12px 14px",
                      background: "var(--s2)",
                      marginBottom: 14,
                    }}
                  >
                    <div
                      role="button"
                      tabIndex={0}
                      aria-expanded={compPanelOpen}
                      onClick={() => setCompPanelOpen((v) => !v)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter" || e.key === " ") {
                          e.preventDefault();
                          setCompPanelOpen((v) => !v);
                        }
                      }}
                      style={{
                        fontSize: 11,
                        fontWeight: 700,
                        color: "var(--t2)",
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        cursor: "pointer",
                        userSelect: "none",
                      }}
                    >
                      {compPanelOpen ? (
                        <ChevronUp size={13} color="var(--t3)" />
                      ) : (
                        <ChevronDown size={13} color="var(--t3)" />
                      )}
                      <Layers size={13} color="var(--accent-text)" />
                      成本組件（原料庫・兩平台共用）
                      <span
                        style={{
                          fontFamily: mono,
                          color: "var(--t3)",
                          fontWeight: 600,
                        }}
                      >
                        {Object.keys(components).length} 顆・
                        {compGroups.length} 分類
                      </span>
                      {!compPanelOpen && (
                        <span
                          style={{
                            fontSize: 10,
                            color: "var(--t4)",
                            fontWeight: 500,
                            marginLeft: "auto",
                          }}
                        >
                          點擊展開管理
                        </span>
                      )}
                    </div>
                    {compPanelOpen && (
                      <>
                        <div
                          style={{
                            fontSize: 10,
                            color: "var(--t4)",
                            margin: "6px 0 10px",
                            lineHeight: 1.6,
                          }}
                        >
                          改這裡一個單價＝所有掛配方的商品自動重算（已鎖定月份不受影響）。商品掛配方：點商品列尾的
                          <Layers
                            size={10}
                            style={{ margin: "0 2px", verticalAlign: "-1px" }}
                          />
                          圖示。
                        </div>
                        <div style={{ position: "relative", marginBottom: 8 }}>
                          <Search
                            size={12}
                            color="var(--t4)"
                            style={{
                              position: "absolute",
                              left: 10,
                              top: "50%",
                              transform: "translateY(-50%)",
                            }}
                          />
                          <input
                            type="text"
                            value={compSearch}
                            placeholder="搜尋組件名稱或分類..."
                            aria-label="搜尋成本組件"
                            onChange={(e) => setCompSearch(e.target.value)}
                            style={{
                              ...inp,
                              width: "100%",
                              maxWidth: 300,
                              textAlign: "left",
                              paddingLeft: 30,
                              borderRadius: 8,
                              padding: "7px 10px 7px 30px",
                              fontSize: 12,
                            }}
                          />
                        </div>
                    <datalist id="comp-cats">
                      {compCats.map((c) => (
                        <option key={c} value={c} />
                      ))}
                    </datalist>
                    {compGroups
                      .map(([cat, list]) => {
                        const q = dCompSearch.trim().toLowerCase();
                        const shown = q
                          ? list.filter(
                              ([, c]) =>
                                c.name.toLowerCase().includes(q) ||
                                (c.cat || "").toLowerCase().includes(q)
                            )
                          : list;
                        return [cat, shown];
                      })
                      .filter(([, shown]) => shown.length > 0)
                      .map(([cat, list]) => {
                        const catOpen =
                          !!compCatOpen[cat] || !!dCompSearch.trim();
                        return (
                      <div
                        key={cat}
                        style={{
                          border: "1px solid var(--s3)",
                          borderRadius: 10,
                          marginBottom: 6,
                          background: "var(--s1)",
                          overflow: "hidden",
                        }}
                      >
                        <div
                          role="button"
                          tabIndex={0}
                          aria-expanded={catOpen}
                          onClick={() =>
                            setCompCatOpen((p) => ({ ...p, [cat]: !p[cat] }))
                          }
                          onKeyDown={(e) => {
                            if (e.key === "Enter" || e.key === " ") {
                              e.preventDefault();
                              setCompCatOpen((p) => ({
                                ...p,
                                [cat]: !p[cat],
                              }));
                            }
                          }}
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: 6,
                            padding: "8px 10px",
                            cursor: "pointer",
                            fontSize: 11,
                            fontWeight: 800,
                            userSelect: "none",
                          }}
                        >
                          {catOpen ? (
                            <ChevronUp size={12} color="var(--t3)" />
                          ) : (
                            <ChevronDown size={12} color="var(--t3)" />
                          )}
                          <span style={{ color: "var(--accent-text)" }}>
                            {cat}
                          </span>
                          <span
                            style={{
                              color: "var(--t3)",
                              fontWeight: 600,
                              fontFamily: mono,
                            }}
                          >
                            {list.length}
                          </span>
                        </div>
                        {catOpen && (
                        <div
                          style={{
                            display: "flex",
                            flexWrap: "wrap",
                            gap: 8,
                            alignItems: "center",
                            padding: "0 10px 10px",
                          }}
                        >
                          {list.map(([id, c]) => (
                            <div
                              key={id}
                              style={{
                                display: "inline-flex",
                                alignItems: "center",
                                gap: 4,
                                background: "var(--s1)",
                                border: "1px solid var(--s3)",
                                borderRadius: 8,
                                padding: "4px 6px",
                              }}
                            >
                              <input
                                key={id + "_" + c.name}
                                defaultValue={c.name}
                                aria-label="組件名稱"
                                onBlur={(e) =>
                                  commitCompField(id, "name", e.target.value)
                                }
                                onKeyDown={(e) => {
                                  if (e.key === "Enter")
                                    e.currentTarget.blur();
                                }}
                                style={{
                                  ...inp,
                                  width: 110,
                                  textAlign: "left",
                                  fontSize: 11,
                                  padding: "4px 6px",
                                  fontFamily:
                                    "'Inter','Noto Sans TC',sans-serif",
                                }}
                              />
                              <FpInput
                                field={id}
                                label={`${c.name} 單價`}
                                value={c.price}
                                onCommit={(cid, v) =>
                                  commitCompField(cid, "price", v)
                                }
                              />
                              <input
                                key={id + "_c_" + (c.cat || "")}
                                defaultValue={c.cat || ""}
                                list="comp-cats"
                                placeholder="分類"
                                aria-label={`${c.name} 分類`}
                                onBlur={(e) =>
                                  commitCompField(id, "cat", e.target.value)
                                }
                                onKeyDown={(e) => {
                                  if (e.key === "Enter")
                                    e.currentTarget.blur();
                                }}
                                style={{
                                  ...inp,
                                  width: 76,
                                  textAlign: "left",
                                  fontSize: 10,
                                  padding: "4px 6px",
                                  fontFamily:
                                    "'Inter','Noto Sans TC',sans-serif",
                                }}
                              />
                              {(compUsage[id] || 0) > 0 && (
                                <span
                                  style={{
                                    fontSize: 9,
                                    fontFamily: mono,
                                    color: "var(--t3)",
                                    fontWeight: 700,
                                  }}
                                  title="被幾個配方使用"
                                >
                                  ×{compUsage[id]}
                                </span>
                              )}
                              <Btn
                                v="ghost"
                                aria-label={`刪除組件 ${c.name}`}
                                onClick={() => deleteComponent(id)}
                                style={{ padding: 2 }}
                              >
                                <Trash2 size={11} color="var(--t4)" />
                              </Btn>
                            </div>
                          ))}
                        </div>
                        )}
                      </div>
                        );
                      })}
                    <div
                      style={{
                        display: "flex",
                        flexWrap: "wrap",
                        alignItems: "center",
                        gap: 4,
                        marginTop: compGroups.length ? 4 : 0,
                      }}
                    >
                      <input
                        value={newComp.name}
                        placeholder="新組件名稱（例：高山茶 150g）"
                        aria-label="新組件名稱"
                        onChange={(e) =>
                          setNewComp((p) => ({ ...p, name: e.target.value }))
                        }
                        onKeyDown={(e) => {
                          if (e.key === "Enter") addComponent();
                        }}
                        style={{
                          ...inp,
                          width: 170,
                          textAlign: "left",
                          fontSize: 11,
                          padding: "4px 6px",
                          fontFamily: "'Inter','Noto Sans TC',sans-serif",
                        }}
                      />
                      <input
                        type="number"
                        step="0.01"
                        value={newComp.price}
                        placeholder="單價"
                        aria-label="新組件單價"
                        onChange={(e) =>
                          setNewComp((p) => ({
                            ...p,
                            price: e.target.value,
                          }))
                        }
                        onKeyDown={(e) => {
                          if (e.key === "Enter") addComponent();
                        }}
                        style={{ ...inp, width: 72, fontSize: 11 }}
                      />
                      <input
                        value={newComp.cat}
                        list="comp-cats"
                        placeholder="分類（例：茶葉）"
                        aria-label="新組件分類"
                        onChange={(e) =>
                          setNewComp((p) => ({ ...p, cat: e.target.value }))
                        }
                        onKeyDown={(e) => {
                          if (e.key === "Enter") addComponent();
                        }}
                        style={{
                          ...inp,
                          width: 110,
                          textAlign: "left",
                          fontSize: 11,
                          padding: "4px 6px",
                          fontFamily: "'Inter','Noto Sans TC',sans-serif",
                        }}
                      />
                      <Btn v="primary" onClick={addComponent}>
                        新增
                      </Btn>
                    </div>
                      </>
                    )}
                  </div>
                  <div
                    style={{
                      display: "flex",
                      flexWrap: "wrap",
                      alignItems: "center",
                      gap: 10,
                      marginBottom: 14,
                    }}
                  >
                    <div style={{ position: "relative", flex: "1 1 240px", maxWidth: 360 }}>
                      <Search
                        size={14}
                        color="var(--t4)"
                        style={{
                          position: "absolute",
                          left: 12,
                          top: "50%",
                          transform: "translateY(-50%)",
                        }}
                      />
                      <input
                        type="text"
                        placeholder="搜尋商品名稱或規格 ..."
                        aria-label="搜尋商品名稱或規格"
                        value={mSearch}
                        onChange={(e) => setMSearch(e.target.value)}
                        style={{
                          ...inp,
                          width: "100%",
                          textAlign: "left",
                          paddingLeft: 36,
                          borderRadius: 10,
                          padding: "10px 12px 10px 36px",
                          fontSize: 13,
                        }}
                      />
                    </div>
                    <label
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        fontSize: 11,
                        fontWeight: 600,
                        color: "var(--t3)",
                        cursor: "pointer",
                        whiteSpace: "nowrap",
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={cleanupOnly}
                        onChange={(e) => setCleanupOnly(e.target.checked)}
                        style={{ accentColor: "var(--wn)" }}
                      />{" "}
                      只看可清理（180 天未售/無紀錄）
                    </label>
                    <label
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        fontSize: 11,
                        fontWeight: 600,
                        color: "var(--t3)",
                        cursor: "pointer",
                        whiteSpace: "nowrap",
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={soldOnly}
                        onChange={(e) => setSoldOnly(e.target.checked)}
                        style={{ accentColor }}
                      />{" "}
                      只顯示本期有銷售
                    </label>
                    {cleanupOnly && matrixList.length > 0 && (
                      <Btn v="danger" onClick={bulkCleanStale}>
                        <Trash2 size={11} /> 清除清單全部 {matrixList.length} 項
                      </Btn>
                    )}
                  </div>
                  <div
                    style={{
                      overflowX: "auto",
                      overflowY: "auto",
                      maxHeight: 480,
                      border: "1px solid var(--s3)",
                      borderRadius: 12,
                    }}
                  >
                    {/* 手機隱藏銷量／淨利貢獻／毛利率，留商品名稱·規格·單位成本（補成本才是手機會做的事） */}
                    <table
                      className="tb-cost"
                      style={{
                        width: "100%",
                        borderCollapse: "collapse",
                        minWidth: 760,
                      }}
                    >
                      <thead>
                        <tr>
                          <SortTh
                            sortKey="name"
                            currentSort={costSort}
                            onSort={(k) =>
                              setCostSort((p) => ({
                                key: k,
                                dir:
                                  p.key === k
                                    ? p.dir === "desc"
                                      ? "asc"
                                      : "desc"
                                    : "desc",
                              }))
                            }
                          >
                            商品名稱
                          </SortTh>
                          <th scope="col" style={{ ...th, textAlign: "left" }}>
                            規格
                          </th>
                          <SortTh
                            sortKey="soldQty"
                            currentSort={costSort}
                            onSort={(k) =>
                              setCostSort((p) => ({
                                key: k,
                                dir:
                                  p.key === k
                                    ? p.dir === "desc"
                                      ? "asc"
                                      : "desc"
                                    : "desc",
                              }))
                            }
                            align="right"
                          >
                            銷量
                          </SortTh>
                          <SortTh
                            sortKey="profit"
                            currentSort={costSort}
                            onSort={(k) =>
                              setCostSort((p) => ({
                                key: k,
                                dir:
                                  p.key === k
                                    ? p.dir === "desc"
                                      ? "asc"
                                      : "desc"
                                    : "desc",
                              }))
                            }
                            align="right"
                          >
                            淨利貢獻
                          </SortTh>
                          <SortTh
                            sortKey="margin"
                            currentSort={costSort}
                            onSort={(k) =>
                              setCostSort((p) => ({
                                key: k,
                                dir:
                                  p.key === k
                                    ? p.dir === "desc"
                                      ? "asc"
                                      : "desc"
                                    : "desc",
                              }))
                            }
                            align="right"
                          >
                            毛利率
                          </SortTh>
                          <SortTh
                            sortKey="cost"
                            currentSort={costSort}
                            onSort={(k) =>
                              setCostSort((p) => ({
                                key: k,
                                dir:
                                  p.key === k
                                    ? p.dir === "desc"
                                      ? "asc"
                                      : "desc"
                                    : "desc",
                              }))
                            }
                            align="right"
                          >
                            單位成本
                          </SortTh>
                          <th
                            scope="col"
                            style={{ ...th, textAlign: "center", width: 40 }}
                          ></th>
                        </tr>
                      </thead>
                      <tbody>
                        {!matrixList.length ? (
                          <tr>
                            <td
                              colSpan={7}
                              style={{
                                ...td2,
                                textAlign: "center",
                                color: "var(--t4)",
                                padding: 40,
                              }}
                            >
                              尚無商品數據
                            </td>
                          </tr>
                        ) : (
                          (() => {
                            let missFound = false;
                            return matrixList.map((p) => {
                              const miss = missCost.keys.has(p.key),
                                hs = p.soldQty > 0;
                              const isFirstMiss = miss && !missFound;
                              if (isFirstMiss) missFound = true;
                              const profitVal =
                                p.profitContribution ?? p.estProfit ?? 0;
                              const gmVal =
                                p.totalRevenue > 0
                                  ? (p.totalRevenue - p.totalCost) /
                                    p.totalRevenue
                                  : null;
                              const hasRecipe =
                                Array.isArray(recipes[p.key]) &&
                                recipes[p.key].length > 0;
                              return (
                                <React.Fragment key={p.key}>
                                <tr
                                  ref={isFirstMiss ? firstMissRef : undefined}
                                  className={miss && hs ? "rw" : ""}
                                >
                                  <td
                                    style={{
                                      ...td2,
                                      fontWeight: 600,
                                      maxWidth: 260,
                                      overflow: "hidden",
                                    }}
                                    title={p.name}
                                  >
                                    <div
                                      style={{
                                        whiteSpace: "nowrap",
                                        overflow: "hidden",
                                        textOverflow: "ellipsis",
                                      }}
                                    >
                                      {p.name}
                                    </div>
                                    {p.stale && (
                                      <div
                                        style={{
                                          fontSize: 9,
                                          fontWeight: 700,
                                          marginTop: 2,
                                          whiteSpace: "nowrap",
                                          color: p.neverSold
                                            ? "var(--t3)"
                                            : "var(--wn)",
                                        }}
                                      >
                                        {p.neverSold
                                          ? "無銷售紀錄"
                                          : `久未售出・最後 ${p.lastSold}（${p.staleDays} 天前）`}
                                      </div>
                                    )}
                                  </td>
                                  <td
                                    style={{
                                      ...td2,
                                      color: "var(--t3)",
                                      fontSize: 12,
                                    }}
                                  >
                                    {p.option}
                                  </td>
                                  <td
                                    style={{
                                      ...td2,
                                      textAlign: "right",
                                      fontWeight: 700,
                                      fontFamily: mono,
                                    }}
                                  >
                                    {p.soldQty}
                                  </td>
                                  <td
                                    style={{
                                      ...td2,
                                      textAlign: "right",
                                      fontWeight: 700,
                                      fontFamily: mono,
                                      color:
                                        profitVal >= 0
                                          ? "var(--up)"
                                          : "var(--dn)",
                                    }}
                                  >
                                    {fmt$(profitVal)}
                                  </td>
                                  <td
                                    style={{
                                      ...td2,
                                      textAlign: "right",
                                      fontWeight: 600,
                                      fontFamily: mono,
                                      color:
                                        gmVal === null
                                          ? "var(--t4)"
                                          : gmVal < 0
                                          ? "var(--dn)"
                                          : gmVal < 0.35
                                          ? "var(--wn)"
                                          : "var(--t1)",
                                    }}
                                  >
                                    {gmVal === null ? "—" : fmtP(gmVal)}
                                  </td>
                                  <td style={{ ...td2, textAlign: "right" }}>
                                    <div
                                      style={{
                                        display: "flex",
                                        alignItems: "center",
                                        justifyContent: "flex-end",
                                        gap: 4,
                                      }}
                                    >
                                      {miss && hs && !hasRecipe && (
                                        <span
                                          style={{
                                            fontSize: 10,
                                            color: "var(--wn)",
                                            fontWeight: 700,
                                          }}
                                        >
                                          —
                                        </span>
                                      )}
                                      {/* 掛了率的門市鍵：標「估 X%」但輸入框保留——填了手填成本，
                                          率自動失效（優先序：配方 > 手填 > 率） */}
                                      {isPOS &&
                                        Number(posRatios[p.key]) > 0 &&
                                        !hasRecipe &&
                                        !(Number(costs[p.key]) > 0) && (
                                          <span
                                            title="目前用成本率估算（上方「泛稱品項成本率」卡調整）。在右邊填入單位成本即可改用真成本，率自動失效"
                                            style={{
                                              fontFamily: mono,
                                              fontWeight: 700,
                                              fontSize: 11,
                                              color: "var(--wn)",
                                              whiteSpace: "nowrap",
                                            }}
                                          >
                                            估 {Number(posRatios[p.key])}%
                                          </span>
                                        )}
                                      {hasRecipe ? (
                                        <span
                                          role="button"
                                          tabIndex={0}
                                          title="配方計算值——點擊編輯配方"
                                          onClick={() =>
                                            setRecipeEditKey(
                                              recipeEditKey === p.key
                                                ? null
                                                : p.key
                                            )
                                          }
                                          onKeyDown={(e) => {
                                            if (
                                              e.key === "Enter" ||
                                              e.key === " "
                                            ) {
                                              e.preventDefault();
                                              setRecipeEditKey(
                                                recipeEditKey === p.key
                                                  ? null
                                                  : p.key
                                              );
                                            }
                                          }}
                                          style={{
                                            fontFamily: mono,
                                            fontWeight: 700,
                                            fontSize: 13,
                                            color: "var(--accent-text)",
                                            cursor: "pointer",
                                            whiteSpace: "nowrap",
                                          }}
                                        >
                                          {fmt$d(costsEff[p.key])}
                                          <span
                                            style={{
                                              fontSize: 9,
                                              marginLeft: 4,
                                              color: "var(--t3)",
                                              fontWeight: 600,
                                            }}
                                          >
                                            配方
                                          </span>
                                        </span>
                                      ) : (
                                        <CostInput
                                          costKey={p.key}
                                          label={`${p.name} ${
                                            p.option || ""
                                          } 單位成本`}
                                          value={costs[p.key]}
                                          miss={miss}
                                          onCommit={commitCost}
                                        />
                                      )}
                                    </div>
                                  </td>
                                  <td style={{ ...td2, textAlign: "center" }}>
                                    <div
                                      style={{
                                        display: "flex",
                                        gap: 2,
                                        justifyContent: "center",
                                      }}
                                    >
                                      <Btn
                                        v="ghost"
                                        aria-label={`編輯 ${p.name} 的配方`}
                                        title="配方（組件×用量）"
                                        onClick={() =>
                                          setRecipeEditKey(
                                            recipeEditKey === p.key
                                              ? null
                                              : p.key
                                          )
                                        }
                                        style={{ padding: "2px" }}
                                      >
                                        <Layers
                                          size={12}
                                          color={
                                            hasRecipe
                                              ? "var(--accent-text)"
                                              : "var(--t4)"
                                          }
                                        />
                                      </Btn>
                                      <Btn
                                        v="ghost"
                                        aria-label={`刪除 ${p.name} 的成本設定`}
                                        onClick={() =>
                                          setConfirmBox({
                                            title: "刪除成本設定",
                                            message: `確定刪除「${p.name}${
                                              p.option &&
                                              p.option !== "標準規格"
                                                ? `（${p.option}）`
                                                : ""
                                            }」的手填單位成本？${
                                              hasRecipe
                                                ? "\n（此商品掛有配方，配方不受影響、仍以配方計算）"
                                                : ""
                                            }`,
                                            danger: true,
                                            onOk: () => {
                                              /* functional update：確認框開著期間若有遠端同步／復原改動，不要用過期閉包蓋回去 */
                                              setCosts((prev) => {
                                                const n = { ...prev };
                                                delete n[p.key];
                                                return n;
                                              });
                                            },
                                          })
                                        }
                                        style={{ padding: "2px" }}
                                      >
                                        <Trash2 size={12} color="var(--t4)" />
                                      </Btn>
                                    </div>
                                  </td>
                                </tr>
                                {recipeEditKey === p.key && (
                                  <tr>
                                    <td
                                      colSpan={7}
                                      style={{
                                        ...td2,
                                        background: "var(--s2)",
                                        padding: "14px 18px",
                                      }}
                                    >
                                      <div
                                        style={{
                                          fontSize: 12,
                                          fontWeight: 700,
                                          color: "var(--t1)",
                                          marginBottom: 8,
                                        }}
                                      >
                                        配方：{p.name}
                                        {p.option && p.option !== "標準規格"
                                          ? `（${p.option}）`
                                          : ""}
                                        <span
                                          style={{
                                            color: "var(--t3)",
                                            fontWeight: 500,
                                            marginLeft: 8,
                                            fontSize: 11,
                                          }}
                                        >
                                          單位成本＝各組件單價×用量加總
                                        </span>
                                      </div>
                                      {(recipes[p.key] || []).map((l, li) => {
                                        const comp = components[l.compId];
                                        return (
                                          <div
                                            key={li}
                                            style={{
                                              display: "flex",
                                              alignItems: "center",
                                              gap: 8,
                                              padding: "3px 0",
                                              fontSize: 12,
                                            }}
                                          >
                                            <span
                                              style={{
                                                minWidth: 150,
                                                color: comp
                                                  ? "var(--t1)"
                                                  : "var(--dn)",
                                                fontWeight: 600,
                                              }}
                                            >
                                              {comp
                                                ? comp.name
                                                : "（組件已刪除，成本以 0 計）"}
                                            </span>
                                            <span
                                              style={{
                                                color: "var(--t4)",
                                                fontSize: 11,
                                              }}
                                            >
                                              ×
                                            </span>
                                            <FpInput
                                              field={li}
                                              label="用量"
                                              value={l.qty}
                                              onCommit={(idx, v) =>
                                                setRecipeLine(p.key, idx, {
                                                  qty: Number(v) || 0,
                                                })
                                              }
                                            />
                                            <span
                                              style={{
                                                fontFamily: mono,
                                                color: "var(--t2)",
                                                fontSize: 11,
                                                minWidth: 88,
                                                textAlign: "right",
                                              }}
                                            >
                                              ={" "}
                                              {fmt$d(
                                                (Number(comp?.price) || 0) *
                                                  (Number(l.qty) || 0)
                                              )}
                                            </span>
                                            <Btn
                                              v="ghost"
                                              aria-label="移除此組件"
                                              onClick={() =>
                                                setRecipeLine(p.key, li, null)
                                              }
                                              style={{ padding: 2 }}
                                            >
                                              <X
                                                size={12}
                                                color="var(--t4)"
                                              />
                                            </Btn>
                                          </div>
                                        );
                                      })}
                                      <div
                                        style={{
                                          display: "flex",
                                          alignItems: "center",
                                          gap: 8,
                                          marginTop: 8,
                                          flexWrap: "wrap",
                                        }}
                                      >
                                        <CompPicker
                                          compGroups={compGroups}
                                          onPick={(id) =>
                                            addRecipeLine(p.key, id)
                                          }
                                        />
                                        <span
                                          style={{
                                            fontFamily: mono,
                                            fontWeight: 800,
                                            fontSize: 13,
                                            color: "var(--accent-text)",
                                            marginLeft: "auto",
                                          }}
                                        >
                                          合計{" "}
                                          {fmt$d(
                                            recipeCost(
                                              recipes[p.key],
                                              components
                                            )
                                          )}
                                        </span>
                                        {(recipes[p.key] || []).length > 0 && (
                                          <Btn
                                            v="danger"
                                            onClick={() =>
                                              removeRecipe(p.key)
                                            }
                                            style={{ fontSize: 10 }}
                                          >
                                            移除配方
                                          </Btn>
                                        )}
                                        <Btn
                                          onClick={() =>
                                            setRecipeEditKey(null)
                                          }
                                          style={{ fontSize: 10 }}
                                        >
                                          完成
                                        </Btn>
                                      </div>
                                      {Object.keys(components).length ===
                                        0 && (
                                        <div
                                          style={{
                                            fontSize: 11,
                                            color: "var(--wn)",
                                            marginTop: 6,
                                          }}
                                        >
                                          原料庫還沒有組件——先在上方「成本組件」新增（例：高山茶
                                          150g、茶包袋、禮盒A、提袋）
                                        </div>
                                      )}
                                    </td>
                                  </tr>
                                )}
                                </React.Fragment>
                              );
                            });
                          })()
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
  );

  return (
    <div
      data-theme={theme}
      style={{
        minHeight: "100vh",
        background: "var(--bg)",
        color: "var(--t1)",
        fontFamily: "'Inter','Noto Sans TC',sans-serif",
        transition: "background .3s,color .3s",
      }}
    >
      <style>{CSS}</style>

      {/* Header */}
      <header
        className="app-header"
        style={{
          position: "sticky",
          top: 0,
          zIndex: 50,
          background: "var(--header-bg)",
          backdropFilter: "blur(24px)",
          borderBottom: "1px solid var(--s3)",
        }}
      >
        <div
          className="hd-wrap"
          style={{
            maxWidth: 1560,
            margin: "0 auto",
            padding: "10px 24px",
            display: "flex",
            flexWrap: "wrap",
            alignItems: "center",
            justifyContent: "space-between",
            gap: 10,
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div
              style={{
                width: 34,
                height: 34,
                borderRadius: 10,
                background: accentDim,
                border: `1px solid ${accentBdr}`,
                color: accentColor,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              <Layers size={18} />
            </div>
            <div>
              <h1
                style={{
                  fontSize: 15,
                  fontWeight: 800,
                  letterSpacing: "-0.01em",
                }}
              >
                {isOverview
                  ? "跨平台"
                  : isPOS
                  ? "門市"
                  : isSL
                  ? "官網"
                  : "蝦皮"}{" "}
                利潤決策中心
              </h1>
              <div
                className="hd-sub"
                style={{
                  fontSize: 10,
                  color: "var(--t3)",
                  fontFamily: mono,
                  letterSpacing: "0.06em",
                }}
              >
                PROFIT INTELLIGENCE · FIREBASE SYNC
              </div>
            </div>
          </div>
          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              gap: 8,
              alignItems: "center",
            }}
          >
            {/* Platform Toggle */}
            <div
              style={{
                display: "flex",
                border: "1px solid var(--s3)",
                borderRadius: 10,
                overflow: "hidden",
              }}
            >
              {[
                { id: "overview", label: "總覽" },
                { id: "shopline", label: "官網" },
                { id: "shopee", label: "蝦皮" },
                { id: "pos", label: "門市" },
              ].map((p) => (
                <button
                  key={p.id}
                  /* 切換平台保留目前選擇的期間（年/月/自訂區間都不動） */
                  onClick={() => setPlatform(p.id)}
                  style={{
                    padding: "6px 14px",
                    fontSize: 11,
                    fontWeight: 700,
                    border: "none",
                    cursor: "pointer",
                    background:
                      platform === p.id
                        ? p.id === "overview"
                          ? "var(--blue)"
                          : p.id === "pos"
                          ? "var(--pos-accent)"
                          : accentColor
                        : "var(--s1)",
                    color: platform === p.id ? "#fff" : "var(--t2)",
                    transition: "all .15s",
                  }}
                >
                  {p.label}
                </button>
              ))}
            </div>
            <SyncDot status={sync} last={lastSyncAt} />
            <button
              onClick={() => setTheme((t) => (t === "dark" ? "light" : "dark"))}
              aria-label={
                theme === "dark" ? "切換為淺色主題" : "切換為深色主題"
              }
              style={{
                width: 32,
                height: 32,
                borderRadius: 8,
                border: "1px solid var(--s3)",
                background: "var(--s2)",
                color: "var(--t2)",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              {theme === "dark" ? <Sun size={15} /> : <Moon size={15} />}
            </button>
            <select
              value={sY}
              onChange={(e) => {
                const v = e.target.value;
                /* 會計年度（公司 4 月起算）：換算成自訂區間 4/1～翌年 3/31 */
                if (v.startsWith("FY")) {
                  const y = Number(v.slice(2));
                  setRange({ from: `${y}-04-01`, to: `${y + 1}-03-31` });
                  setSY("Custom");
                  setSM("All");
                  return;
                }
                setSY(v);
                setSM("All");
              }}
              aria-label="選擇年份"
              style={sel}
            >
              <option value="All">歷年數據</option>
              {aY.map((y) => (
                <option key={y} value={y}>
                  {y} 年
                </option>
              ))}
              {[...new Set(aY.flatMap((y) => [Number(y) - 1, Number(y)]))]
                .sort((a, b) => b - a)
                .map((y) => (
                  <option key={`FY${y}`} value={`FY${y}`}>
                    FY{y}（{y}/4～{y + 1}/3）
                  </option>
                ))}
              <option value="Custom">自訂區間</option>
            </select>
            {sY === "Custom" ? (
              <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                <input
                  type="date"
                  value={range.from}
                  aria-label="起始日期"
                  onChange={(e) =>
                    setRange((r) => ({ ...r, from: e.target.value }))
                  }
                  style={{
                    ...sel,
                    fontFamily: mono,
                    fontSize: 11,
                    padding: "5px 8px",
                  }}
                />
                <span style={{ color: "var(--t4)", fontSize: 11 }}>～</span>
                <input
                  type="date"
                  value={range.to}
                  aria-label="結束日期"
                  onChange={(e) =>
                    setRange((r) => ({ ...r, to: e.target.value }))
                  }
                  style={{
                    ...sel,
                    fontFamily: mono,
                    fontSize: 11,
                    padding: "5px 8px",
                  }}
                />
              </div>
            ) : (
              <select
                value={effM}
                onChange={(e) => setSM(e.target.value)}
                aria-label="選擇月份"
                style={sel}
              >
                <option value="All">全月份</option>
                {aM.map((m) => (
                  <option key={m} value={m}>
                    {m} 月
                  </option>
                ))}
              </select>
            )}
          </div>
        </div>
      </header>

      <div
        className="page-wrap"
        style={{ maxWidth: 1560, margin: "0 auto", padding: "20px 24px 80px" }}
      >
        <div className={isOverview ? "" : "gm"}>
          {/* Sidebar（手機版靠 .side-col 移到資料下面） */}
          <aside
            className="f0 side-col"
            style={{ display: isOverview ? "none" : undefined }}
          >
            <div
              className="side-sticky"
              style={{
                background: "var(--s1)",
                border: "1px solid var(--s3)",
                borderRadius: 16,
                padding: 16,
                position: "sticky",
                top: 64,
                display: "flex",
                flexDirection: "column",
                gap: 12,
              }}
            >
              {/* Upload（手機隱藏：報表檔在電腦上，手機不匯入） */}
              <div
                className="imp-zone"
                role="button"
                tabIndex={0}
                aria-label={`匯入${
                  isPOS ? "門市" : isSL ? "官網" : "蝦皮"
                }報表檔案（點擊或拖曳）`}
                onClick={() => fRef.current._fileInput?.click()}
                onKeyDown={(e) => {
                  if (e.key === "Enter" || e.key === " ") {
                    e.preventDefault();
                    fRef.current._fileInput?.click();
                  }
                }}
                onDragOver={(e) => {
                  e.preventDefault();
                  setDragOver(true);
                }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setDragOver(false);
                  /* 門市可一次拖兩份（交易明細＋訂單明細），依序解析 */
                  const fs = Array.from(e.dataTransfer.files || []);
                  fs.forEach((f, i) => setTimeout(() => processFile(f), i * 60));
                }}
                style={{
                  border: `1.5px dashed ${
                    dragOver ? accentColor : "var(--s4)"
                  }`,
                  borderRadius: 12,
                  padding: "18px 12px",
                  textAlign: "center",
                  cursor: "pointer",
                  background: dragOver ? accentDim : "var(--s2)",
                  transition: "border-color .15s, background .15s",
                }}
              >
                <input
                  ref={(el) => {
                    if (fRef.current) fRef.current._fileInput = el;
                  }}
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  multiple={isPOS}
                  onChange={handleFile}
                  style={{ display: "none" }}
                />
                <FileSpreadsheet size={22} color="var(--t3)" />
                <div
                  style={{
                    marginTop: 6,
                    fontWeight: 700,
                    fontSize: 12,
                    color: "var(--t2)",
                  }}
                >
                  匯入{isPOS ? "門市" : isSL ? "官網" : "蝦皮"}報表
                </div>
                <div style={{ fontSize: 10, color: "var(--t4)", marginTop: 2 }}>
                  {isPOS ? "交易明細＋POS訂單明細（兩份）" : "CSV · XLSX · 拖曳"}
                </div>
              </div>
              {currentData && (
                <div
                  style={{
                    fontSize: 11,
                    color: "var(--t3)",
                    fontWeight: 600,
                    padding: "4px 2px",
                  }}
                >
                  {isPOS
                    ? posData?.summary?.valid
                    : isSL
                    ? slD?.valid
                    : spS?.validN}{" "}
                  筆 ·{" "}
                  {sY === "All" ? "歷年" : sY === "Custom" ? "自訂區間" : sY}
                  {effM !== "All" && sY !== "Custom" ? `/${effM}` : ""}
                </div>
              )}

              {/* Fee Params */}
              <div style={{ borderTop: "1px solid var(--s3)", paddingTop: 12 }}>
                <div
                  style={{
                    fontSize: 11,
                    fontWeight: 700,
                    color: "var(--t3)",
                    display: "flex",
                    alignItems: "center",
                    gap: 6,
                    marginBottom: 10,
                  }}
                >
                  <Settings size={12} /> 財務模型參數
                  {isSL && (
                    <span
                      style={{ fontWeight: 500, color: "var(--t4)", fontSize: 10 }}
                    >
                      （營業費／稅率＝全公司共用）
                    </span>
                  )}
                </div>
                {snapParams && (
                  <div
                    style={{
                      background: "var(--wn-dim)",
                      border: "1px solid var(--wn-bdr)",
                      borderRadius: 8,
                      padding: "8px 10px",
                      marginBottom: 10,
                    }}
                  >
                    <div
                      style={{
                        fontSize: 10,
                        fontWeight: 800,
                        color: "var(--wn)",
                        display: "flex",
                        alignItems: "center",
                        gap: 4,
                        marginBottom: 4,
                      }}
                    >
                      <Lock size={10} />
                      {isLocked ? "本期已鎖定快照" : "本期部分訂單帶快照"}
                      {snapParams.mixed ? "（各月參數不同）" : ""}
                    </div>
                    {snapParams.list.map((sp, i) => (
                      <div
                        key={i}
                        style={{
                          fontSize: 10,
                          color: "var(--t2)",
                          fontFamily: mono,
                          lineHeight: 1.8,
                        }}
                      >
                        營業費 {pct(sp.opExpense)}・稅 {pct(sp.tax)}
                        {isSL ? `・系統費 ${pct(sp.platformFeeRate)}` : ""}
                        {snapParams.mixed ? `（${sp.count} 筆）` : ""}
                      </div>
                    ))}
                    {snapParams.nullCostCount > 0 && (
                      <div
                        style={{
                          fontSize: 10,
                          color: "var(--dn)",
                          fontWeight: 700,
                          marginTop: 4,
                          lineHeight: 1.6,
                        }}
                      >
                        ⚠ {snapParams.nullCostCount}{" "}
                        個品項鎖定時成本未填、不受快照保護，仍會跟著成本表變動——補填成本後「解除
                        → 重新鎖定」
                      </div>
                    )}
                    <div
                      style={{
                        fontSize: 9,
                        color: "var(--t4)",
                        marginTop: 4,
                        lineHeight: 1.6,
                      }}
                    >
                      鎖定的訂單以上述參數計算，下方輸入僅影響未鎖定期間。
                      要修正本期 %：解除快照 → 改參數 → 重新鎖定。
                    </div>
                  </div>
                )}
                {/* 內部營業費／預估稅率＝全公司單一口徑，只在官網頁開放編輯，
                    其他兩頁顯示唯讀（老闆 2026-09-01：不用三個平台各出現一次） */}
                {(isSL
                  ? [
                      { l: "淨利目標", n: "targetNet" },
                      { l: "內部營業費", n: "opExpense" },
                      { l: "預估稅率", n: "tax" },
                      { l: "系統服務費率", n: "platformFeeRate" },
                    ]
                  : isPOS
                  ? /* 門市無平台抽成、無系統費；淨利目標＝六通路統一一個數字（2026-08-25 拍板 12） */
                    [{ l: "淨利目標", n: "posTargetNet" }]
                  : [{ l: "淨利目標", n: "targetNet" }]
                ).map((item) => (
                  <div
                    key={item.n}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "space-between",
                      padding: "5px 0",
                      borderBottom: "1px solid var(--s3)",
                    }}
                  >
                    <span
                      style={{
                        fontSize: 12,
                        fontWeight: 600,
                        color: "var(--t2)",
                      }}
                    >
                      {item.l}
                    </span>
                    <div
                      style={{ display: "flex", alignItems: "center", gap: 4 }}
                    >
                      <FpInput
                        field={item.n}
                        label={`${item.l}（%）`}
                        value={isSL || isPOS ? slFp[item.n] : spFp[item.n]}
                        onCommit={commitFp}
                      />
                      <span style={{ fontSize: 11, color: "var(--t4)" }}>
                        %
                      </span>
                    </div>
                  </div>
                ))}
                {!isSL && (
                  <div
                    style={{
                      fontSize: 11,
                      color: "var(--t3)",
                      lineHeight: 1.8,
                      padding: "7px 0 0",
                    }}
                  >
                    內部營業費{" "}
                    <b style={{ fontFamily: mono, color: "var(--t1)" }}>
                      {slFp.opExpense}%
                    </b>
                    ・預估稅率{" "}
                    <b style={{ fontFamily: mono, color: "var(--t1)" }}>
                      {slFp.tax}%
                    </b>
                    <div style={{ fontSize: 10, color: "var(--t4)" }}>
                      全公司共用一組，到「官網」頁調整
                    </div>
                  </div>
                )}
              </div>

              {/* Commission Panel (Shopee only) */}
              {!isSL && !isPOS && currentData && (
                <MonthlyExpensePanel
                  title="分潤費用"
                  icon={<Users size={12} color="var(--purple)" />}
                  color="var(--purple)"
                  values={commissions}
                  onUpdate={handleComm}
                  selYear={sY}
                  selMonth={effM}
                  range={range}
                  hint="此期間分潤費用將從最終淨利扣除"
                />
              )}

              {/* Reset（手機隱藏：清除訂單是破壞性操作，不在手機上做） */}
              <div className="reset-row" style={{ display: "flex", gap: 6 }}>
                <Btn
                  v="primary"
                  onClick={() => {
                    const src = isPOS ? posOrders : isSL ? slOrders : spOrders;
                    const toDelete = Object.keys(src).filter((k) =>
                      inPeriod(String(src[k].date || ""))
                    );
                    if (!toDelete.length) {
                      toast("本期無訂單可清除", { type: "warning" });
                      return;
                    }
                    /* 用 effM（實際生效的月份）而非 sM，否則會出現
                       「說要清 7 月、實際清全年」的誤導 */
                    const periodLabel =
                      sY === "Custom"
                        ? `${range.from || "起"} ～ ${range.to || "迄"}`
                        : sY === "All"
                        ? "歷年"
                        : effM === "All"
                        ? `${sY} 年`
                        : `${sY}/${effM}`;
                    setConfirmBox({
                      title: "重置本期訂單",
                      message: `將清除「${periodLabel}」的${
                        isPOS ? "門市" : isSL ? "官網" : "蝦皮"
                      }訂單共 ${
                        toDelete.length
                      } 筆。\n清除後 10 秒內可在右下角通知按「復原」，或重新匯入該期報表。${
                        isLocked
                          ? "\n注意：本期已鎖定快照，清除後快照設定將一併移除。"
                          : ""
                      }`,
                      danger: true,
                      onOk: () => {
                        const removed = {};
                        const setter = isPOS
                          ? setPosOrders
                          : isSL
                          ? setSlOrders
                          : setSpOrders;
                        /* functional update：以最新訂單集為基底刪除，避免過期閉包 */
                        setter((cur) => {
                          const updated = { ...cur };
                          toDelete.forEach((k) => {
                            if (updated[k] !== undefined) {
                              removed[k] = updated[k];
                              delete updated[k];
                            }
                          });
                          return updated;
                        });
                        setExpandedId(null);
                        toast(`已清除 ${toDelete.length} 筆訂單`, {
                          type: "info",
                          action: () => {
                            /* 只補回目前不存在的訂單：這 10 秒內重新匯入的新版本
                               不能被舊版本蓋回去（確認文案本來就叫人可以重匯）。
                               略過數用 Set 收、等 updater 跑完（下一個 tick）再提示 */
                            const skipped = new Set();
                            setter((p) => {
                              const n = { ...p };
                              Object.entries(removed).forEach(([k, v]) => {
                                if (n[k] === undefined) n[k] = v;
                                else skipped.add(k);
                              });
                              return n;
                            });
                            setTimeout(() => {
                              if (skipped.size > 0)
                                toast(
                                  `已復原，但略過 ${skipped.size} 筆（清除後已重新匯入，保留新版本）`,
                                  { type: "warning", duration: 8000 }
                                );
                            }, 0);
                          },
                          actionLabel: "復原",
                        });
                      },
                    });
                  }}
                  style={{ flex: 1, justifyContent: "center", fontSize: 10 }}
                >
                  <RotateCcw size={11} /> 重置本期
                </Btn>
              </div>
            </div>
          </aside>

          {/* Main */}
          <main
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 16,
              /* grid 子元素預設 min-width:auto，寬表格會撐開整頁：
                 歸零後讓內層 overflow 容器接手橫向捲動 */
              minWidth: 0,
            }}
          >
            {/* 主內容區單獨包一層邊界：切換平台時 key 變動＝自動重置錯誤狀態 */}
            <ErrorBoundary inner key={platform}>
            {isOverview ? (
              <OverviewDashboard
                slData={slData}
                spData={spData}
                posData={posData}
                posTarget={posTargetOf(slFp)}
                slOrders={slOrders}
                spOrders={spOrders}
                posOrders={posOrders}
                slCosts={slEffCosts}
                spCosts={spEffCosts}
                allMonthly={allMonthly}
                monthlyByPlatform={monthlyByPlatform}
                theme={theme}
                sY={sY}
                sM={effM}
                range={range}
                onNavigate={(id) => setPlatform(id)}
              />
            ) : !currentData ? (
              <div
                className="f0"
                style={{
                  minHeight: 400,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  flexDirection: "column",
                  gap: 14,
                  background: "var(--s1)",
                  border: "1px solid var(--s3)",
                  borderRadius: 16,
                }}
              >
                <FileSpreadsheet size={40} color="var(--s4)" />
                <div
                  style={{ fontSize: 15, fontWeight: 700, color: "var(--t3)" }}
                >
                  等待財務數據注入
                </div>
                <div style={{ fontSize: 12, color: "var(--t4)" }}>
                  {isPOS
                    ? "拖入門市 POS 的「交易明細」＋「POS訂單明細」兩份報表"
                    : `上傳${isSL ? "官網" : "蝦皮"}訂單報表以啟動分析`}
                </div>
              </div>
            ) : isPOS ? (
              <>
                <POSDashboard
                  data={posData}
                  included={posIncluded}
                  onSetIncluded={setPosIncluded}
                  accentColor={accentColor}
                  accentDim={accentDim}
                  accentBdr={accentBdr}
                  opExpense={parseFloat(slFp.opExpense) || 0}
                  taxRate={parseFloat(slFp.tax) || 0}
                  posTarget={posTargetOf(slFp)}
                  isLocked={isLocked}
                  snapParams={snapParams}
                  onToggleSnap={toggleSnap}
                  canLock={sY !== "All" && sY !== "Custom" && effM !== "All"}
                  missN={missCost.n}
                  onJumpMiss={jumpToFirstMissCost}
                  costsEff={posEffCosts}
                  recipes={posRecipes}
                  components={components}
                  ratios={posRatios}
                  monthly={posMonthly}
                  sY={sY}
                  sM={effM}
                  onToggleInvoice={togglePosInvoice}
                  onExported={() => toast("報表已匯出", { type: "success" })}
                  periodLabel={
                    sY === "Custom"
                      ? `${range.from || "起"}~${range.to || "迄"}`
                      : sY === "All"
                      ? "歷年"
                      : effM === "All"
                      ? `${sY}年`
                      : `${sY}-${effM}`
                  }
                />
                {posRatioCard}
                {costMatrixCard}
              </>
            ) : (
              <>
                {/* 本期零筆提示：區分「沒匯資料/期間選錯」與「業績歸零」，
                    避免全 0 的 KPI 被誤讀成真實警訊 */}
                {((isSL && slD && slD.valid === 0) ||
                  (!isSL && spS && spS.validN === 0)) && (
                  <div
                    className="f0"
                    style={{
                      background: "var(--wn-dim)",
                      border: "1px solid var(--wn-bdr)",
                      borderRadius: 14,
                      padding: "14px 20px",
                      display: "flex",
                      alignItems: "center",
                      gap: 10,
                    }}
                  >
                    <AlertTriangle size={16} color="var(--wn)" />
                    <div
                      style={{
                        fontSize: 12,
                        color: "var(--t1)",
                        fontWeight: 600,
                        lineHeight: 1.6,
                      }}
                    >
                      此期間沒有任何有效訂單，下方數字全為
                      0——請先確認：期間選擇是否正確（含自訂區間起訖日）、該期間的
                      {isSL ? "官網" : "蝦皮"}報表是否已匯入。
                    </div>
                  </div>
                )}
                {/* ══ SHOPLINE HERO + KPI ══ */}
                {isSL && slD && (
                  <>
                    <div
                      className="f1"
                      style={{
                        background: "var(--s1)",
                        border: "1px solid var(--s3)",
                        borderRadius: 16,
                        padding: "32px 36px",
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          flexWrap: "wrap",
                          gap: 8,
                          alignItems: "center",
                          marginBottom: 16,
                        }}
                      >
                        <Tag v={slD.gapVal >= 0 ? "ok" : "bad"}>
                          <Zap size={10} /> {slD.gapVal >= 0 ? "穩健" : "警告"}
                        </Tag>
                        {slD.returnCount > 0 && (
                          <Tag v={slD.returnRate > 0.05 ? "warn" : "default"}>
                            退貨 {slD.returnCount} 筆 ·{" "}
                            {fmtP(slD.returnRate)}
                          </Tag>
                        )}
                        {/* 已鎖定的期間成本已凍結，不再提醒補成本（老闆 2026-09-03） */}
                        {missCost.n > 0 && !isLocked && (
                          <Tag
                            v="warn"
                            style={{ cursor: "pointer" }}
                            onClick={jumpToFirstMissCost}
                          >
                            <AlertCircle size={10} /> 未填成本 {missCost.n}/
                            {missCost.total}
                          </Tag>
                        )}
                        <Btn
                          v={isLocked ? "danger" : "default"}
                          onClick={toggleSnap}
                        >
                          <Lock size={11} />{" "}
                          {isLocked ? "解除快照" : "鎖定快照"}
                        </Btn>
                        {isLocked && snapParams && (
                          <Tag v="default">
                            <Lock size={10} />{" "}
                            {snapParams.mixed
                              ? "快照參數各月不同（見側欄）"
                              : `快照 營業費 ${pct(
                                  snapParams.list[0].opExpense
                                )}・稅 ${pct(snapParams.list[0].tax)}`}
                          </Tag>
                        )}
                        <span
                          style={{
                            fontSize: 12,
                            color: slD.gapVal >= 0 ? "var(--t3)" : "var(--wn)",
                            marginLeft: 8,
                          }}
                        >
                          {slD.gapVal >= 0
                            ? `✓ 淨利率 ${fmtP(
                                slD.trueNetMargin
                              )}，超標 ${slD.gapVal.toFixed(1)}%`
                            : `⚠ 淨利率 ${fmtP(
                                slD.trueNetMargin
                              )}，距目標差 ${Math.abs(slD.gapVal).toFixed(1)}%`}
                        </span>
                      </div>
                      <div
                        style={{
                          display: "flex",
                          alignItems: "flex-end",
                          justifyContent: "space-between",
                          flexWrap: "wrap",
                          gap: 24,
                        }}
                      >
                        <div>
                          <div
                            style={{
                              fontSize: 12,
                              fontWeight: 600,
                              color: "var(--t3)",
                              marginBottom: 4,
                              letterSpacing: "0.06em",
                            }}
                          >
                            最終結算淨利 · NET PROFIT
                          </div>
                          <div
                            className="hero-num"
                            style={{
                              lineHeight: 1,
                              fontWeight: 700,
                              letterSpacing: "-0.04em",
                              fontFamily: mono,
                              color: slD.net >= 0 ? "var(--t1)" : "var(--dn)",
                            }}
                          >
                            {fmt$(slD.net)}
                          </div>
                          <div
                            style={{
                              fontSize: 12,
                              color: "var(--t3)",
                              marginTop: 8,
                            }}
                          >
                            原始營收：{fmt$(slD.rawTotal)} ｜ 取消：
                            {fmt$(slD.cancelledTotal)}
                          </div>
                          <PeriodCompare
                            monthly={slMonthly}
                            sY={sY}
                            sM={effM}
                          />
                        </div>
                        <div style={{ textAlign: "right" }}>
                          <div
                            style={{
                              fontSize: 12,
                              fontWeight: 600,
                              color: "var(--t3)",
                            }}
                          >
                            淨利率
                          </div>
                          <div
                            className="hero-pct"
                            style={{
                              fontWeight: 700,
                              fontFamily: mono,
                              lineHeight: 1,
                              color:
                                slD.gapVal >= 0
                                  ? "var(--up)"
                                  : slD.trueNetMargin >=
                                    slD.targetNetRate - 0.03
                                  ? "var(--wn)"
                                  : "var(--dn)",
                            }}
                          >
                            {fmtP(slD.trueNetMargin)}
                          </div>
                          <div
                            style={{
                              fontSize: 11,
                              color: "var(--t3)",
                              marginTop: 4,
                            }}
                          >
                            目標 {fmtP(slD.targetNetRate)}　差距{" "}
                            <span
                              style={{
                                color:
                                  slD.gapVal >= 0 ? "var(--up)" : "var(--dn)",
                              }}
                            >
                              {slD.gapVal >= 0 ? "+" : ""}
                              {slD.gapVal.toFixed(1)}%
                            </span>
                          </div>
                        </div>
                      </div>
                      {/* Waterfall：通路後毛利 − 營業費(含廣告) − 稅賦 = 淨利，逐項可驗算 */}
                      <div
                        style={{
                          marginTop: 28,
                          borderTop: "1px solid var(--s3)",
                          paddingTop: 20,
                        }}
                      >
                        <div
                          style={{
                            fontSize: 11,
                            fontWeight: 700,
                            color: "var(--t3)",
                            marginBottom: 14,
                            letterSpacing: "0.06em",
                          }}
                        >
                          損益分解 · WATERFALL
                        </div>
                        <div
                          style={{
                            display: "flex",
                            flexWrap: "wrap",
                            alignItems: "flex-end",
                            gap: 0,
                          }}
                        >
                          {[
                            {
                              l: "通路後毛利",
                              v: slD.contributionMargin,
                              c: "var(--t1)",
                            },
                            { l: "營業費", v: -slD.opExpTotal, c: "var(--dn)" },
                            { l: "稅賦", v: -slD.taxTotal, c: "var(--dn)" },
                            {
                              l: "淨利",
                              v: slD.net,
                              c: "var(--accent)",
                              bold: true,
                            },
                          ].map((item, i, arr) => (
                            <React.Fragment key={i}>
                              <div
                                style={{
                                  flex: "1 1 0",
                                  minWidth: 90,
                                  textAlign: "center",
                                  padding: "0 8px",
                                }}
                              >
                                <div
                                  style={{
                                    fontSize: 11,
                                    color: "var(--t3)",
                                    fontWeight: 600,
                                    marginBottom: 4,
                                  }}
                                >
                                  {item.l}
                                </div>
                                <div
                                  style={{
                                    fontSize: 20,
                                    fontWeight: item.bold ? 800 : 600,
                                    fontFamily: mono,
                                    color: item.c,
                                    letterSpacing: "-0.02em",
                                  }}
                                >
                                  {fmt$(item.v)}
                                </div>
                              </div>
                              {i < arr.length - 1 && (
                                <div
                                  style={{
                                    color: "var(--s4)",
                                    fontSize: 18,
                                    padding: "0 2px",
                                    alignSelf: "center",
                                  }}
                                >
                                  ›
                                </div>
                              )}
                            </React.Fragment>
                          ))}
                        </div>
                      </div>
                    </div>
                    {/* KPI 上排 4 */}
                    <div className="g4 f2">
                      {[
                        {
                          l: "精準營收基底",
                          v: fmt$(slD.rev),
                          c: "var(--t1)",
                          h: `原始 ${fmt$(slD.rawTotal)} ｜ 取消 ${fmt$(
                            slD.cancelledTotal
                          )}`,
                        },
                        {
                          l: "預估總入帳",
                          v: fmt$(slD.inbound),
                          c: "var(--blue)",
                          h: "扣除金流+物流+平台費",
                        },
                        {
                          l: "商品毛利率",
                          v: fmtP(slD.grossMargin),
                          c:
                            slD.grossMargin >= 0.62
                              ? "var(--up)"
                              : slD.grossMargin >= 0.58
                              ? "var(--accent)"
                              : "var(--dn)",
                          h:
                            slD.grossMargin >= 0.62
                              ? "✓ 超越目標 62%，表現優異"
                              : slD.grossMargin >= 0.58
                              ? "✓ 達標，目標帶 58~60%"
                              : "⚠ 低於警戒線 58%，檢視成本與定價",
                        },
                        {
                          l: "通路後毛利率",
                          v: fmtP(
                            slD.rev > 0 ? slD.contributionMargin / slD.rev : 0
                          ),
                          /* 實收口徑（2026-07-08 基準）：
                             54 優異／52~54 目標帶／50 調整／介入線 */
                          c: (() => {
                            const r =
                              slD.rev > 0
                                ? slD.contributionMargin / slD.rev
                                : 0;
                            return r >= 0.52
                              ? "var(--up)"
                              : r >= 0.5
                              ? "var(--wn)"
                              : "var(--dn)";
                          })(),
                          h: (() => {
                            const r =
                              slD.rev > 0
                                ? slD.contributionMargin / slD.rev
                                : 0;
                            return r >= 0.54
                              ? "✓ 表現優異（理論天花板 ≈54.5%）"
                              : r >= 0.52
                              ? "✓ 達標，目標帶 52~54%"
                              : r >= 0.5
                              ? "⚠ 低於 52%：先調商品結構與曝光，收斂折讓贈品"
                              : "⚠ 低於介入線 50%：檢視折讓、贈品、通路費";
                          })(),
                        },
                      ].map((k, i) => (
                        <div
                          key={i}
                          style={{
                            background: "var(--s1)",
                            border: "1px solid var(--s3)",
                            borderRadius: 14,
                            padding: "22px 24px",
                          }}
                        >
                          <Lbl>{k.l}</Lbl>
                          <div
                            style={{
                              fontSize: 30,
                              fontWeight: 700,
                              fontFamily: mono,
                              letterSpacing: "-0.03em",
                              color: k.c,
                              marginTop: 6,
                            }}
                          >
                            {k.v}
                          </div>
                          <div
                            style={{
                              fontSize: 11,
                              color: "var(--t4)",
                              marginTop: 8,
                            }}
                          >
                            {k.h}
                          </div>
                        </div>
                      ))}
                    </div>
                    {/* KPI 下排 4 */}
                    <div className="g4 f3">
                      {[
                        {
                          l: "單筆平均淨利",
                          v: fmt$(slD.net / (slD.valid || 1)),
                          c: slD.net >= 0 ? "var(--up)" : "var(--dn)",
                          ic: <PieChart size={13} />,
                          h: "平均每筆訂單實際貢獻盈餘",
                        },
                        {
                          l: "通路總成本佔比",
                          v: fmtP(slD.realCommissionRate),
                          c: "var(--blue)",
                          ic: <CreditCard size={13} />,
                          h: "金流＋物流＋系統費",
                        },
                        {
                          l: "營收折讓率",
                          v: fmtP(slD.voucherRate),
                          ic: <Wallet size={13} />,
                          c:
                            slD.voucherRate > 0.065
                              ? "var(--dn)"
                              : slD.voucherRate > 0.055
                              ? "var(--wn)"
                              : slD.voucherRate < 0.04
                              ? "var(--t3)"
                              : slD.voucherRate <= 0.045
                              ? "var(--up)"
                              : "var(--purple)",
                          h:
                            slD.voucherRate > 0.065
                              ? "⚠ 品牌警戒！單月 >6.5%，立即介入"
                              : slD.voucherRate > 0.055
                              ? "⚠ 超出警戒線 5.5%，啟動檢討"
                              : slD.voucherRate < 0.04
                              ? "低於目標下限 4%，折讓/贈品力度偏低"
                              : slD.voucherRate <= 0.045
                              ? "✓ 在目標範圍 4~4.5% 內"
                              : "注意：接近警戒線 5.5%",
                        },
                        {
                          l: "贈品成本佔比",
                          v: fmtP(slD.giftCostRate || 0),
                          ic: <Gift size={13} />,
                          c:
                            (slD.giftCostRate || 0) > 0.045
                              ? "var(--dn)"
                              : (slD.giftCostRate || 0) > 0.035
                              ? "var(--wn)"
                              : (slD.giftCostRate || 0) >= 0.018
                              ? "var(--up)"
                              : "var(--t3)",
                          h:
                            (slD.giftCostRate || 0) > 0.045
                              ? `⚠ 超上限 4.5%！成本 ${fmt$(slD.giftCost)}`
                              : (slD.giftCostRate || 0) > 0.035
                              ? `⚠ 超警戒 3.5%，共 ${slD.giftQty} 件`
                              : (slD.giftCostRate || 0) >= 0.018
                              ? `✓ 目標範圍內，共 ${slD.giftQty} 件`
                              : `低於正常值，共 ${slD.giftQty} 件`,
                        },
                      ].map((k, i) => (
                        <div
                          key={i}
                          style={{
                            background: "var(--s1)",
                            border: "1px solid var(--s3)",
                            borderRadius: 14,
                            padding: "20px 22px",
                          }}
                        >
                          <div
                            style={{
                              fontSize: 12,
                              fontWeight: 600,
                              color: "var(--t3)",
                              display: "flex",
                              alignItems: "center",
                              gap: 6,
                              marginBottom: 8,
                            }}
                          >
                            {k.ic} {k.l}
                          </div>
                          <div
                            style={{
                              fontSize: 26,
                              fontWeight: 700,
                              fontFamily: mono,
                              letterSpacing: "-0.03em",
                              color: k.c,
                            }}
                          >
                            {k.v}
                          </div>
                          <div
                            style={{
                              fontSize: 11,
                              color: "var(--t4)",
                              marginTop: 8,
                            }}
                          >
                            {k.h}
                          </div>
                        </div>
                      ))}
                    </div>
                    {/* KPI 第三排：客單價與每月加價購檔期成效 */}
                    <div className="g3 f3">
                      {[
                        {
                          l: "平均客單價 AOV",
                          v: fmt$(slD.rev / (slD.valid || 1)),
                          c: "var(--t1)",
                          h: "營收 ÷ 有效訂單（含運費收入）",
                        },
                        {
                          l: "加購品營收",
                          v: fmt$(slD.addOnRev),
                          c: slD.addOnRev > 0 ? "var(--purple)" : "var(--t3)",
                          h: `${slD.addOnQty} 件・佔營收 ${fmtP(
                            slD.rev > 0 ? slD.addOnRev / slD.rev : 0
                          )}`,
                        },
                        {
                          l: "加購滲透率",
                          v: fmtP(
                            slD.valid > 0 ? slD.addOnOrders / slD.valid : 0
                          ),
                          c: "var(--blue)",
                          h: `有加購的訂單 ${slD.addOnOrders}/${slD.valid} 筆・每月加價購檔期成效`,
                        },
                      ].map((k, i) => (
                        <div
                          key={i}
                          style={{
                            background: "var(--s1)",
                            border: "1px solid var(--s3)",
                            borderRadius: 14,
                            padding: "20px 22px",
                          }}
                        >
                          <div
                            style={{
                              fontSize: 12,
                              fontWeight: 600,
                              color: "var(--t3)",
                              marginBottom: 8,
                            }}
                          >
                            {k.l}
                          </div>
                          <div
                            style={{
                              fontSize: 26,
                              fontWeight: 700,
                              fontFamily: mono,
                              letterSpacing: "-0.03em",
                              color: k.c,
                            }}
                          >
                            {k.v}
                          </div>
                          <div
                            style={{
                              fontSize: 11,
                              color: "var(--t4)",
                              marginTop: 8,
                            }}
                          >
                            {k.h}
                          </div>
                        </div>
                      ))}
                    </div>
                  </>
                )}

                {/* ══ SHOPEE HERO + KPI ══ */}
                {!isSL && spS && (
                  <>
                    <div
                      className="f1"
                      style={{
                        background: "var(--s1)",
                        border: "1px solid var(--s3)",
                        borderRadius: 16,
                        padding: "32px 36px",
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          flexWrap: "wrap",
                          gap: 8,
                          alignItems: "center",
                          marginBottom: 16,
                        }}
                      >
                        <Tag
                          v={
                            spS.netMargin >= spS.targetNet
                              ? "ok"
                              : spS.netMargin > 0
                              ? "warn"
                              : "bad"
                          }
                        >
                          {spS.badge.label}
                        </Tag>
                        {/* 已鎖定的期間成本已凍結，不再提醒補成本（老闆 2026-09-03） */}
                        {missCost.n > 0 && !isLocked && (
                          <Tag
                            v="warn"
                            style={{ cursor: "pointer" }}
                            onClick={jumpToFirstMissCost}
                          >
                            <AlertCircle size={10} /> 未填成本 {missCost.n}/
                            {missCost.total}
                          </Tag>
                        )}
                        <Btn
                          v={isLocked ? "danger" : "default"}
                          onClick={toggleSnap}
                        >
                          <Lock size={11} />{" "}
                          {isLocked ? "解除快照" : "鎖定快照"}
                        </Btn>
                        {isLocked && snapParams && (
                          <Tag v="default">
                            <Lock size={10} />{" "}
                            {snapParams.mixed
                              ? "快照參數各月不同（見側欄）"
                              : `快照 營業費 ${pct(
                                  snapParams.list[0].opExpense
                                )}・稅 ${pct(snapParams.list[0].tax)}`}
                          </Tag>
                        )}
                        {spS.comm > 0 && (
                          <Tag v="warn">
                            <Users size={10} /> 已扣分潤 {fmt$(spS.comm)}
                          </Tag>
                        )}
                        {spS.refundN > 0 && (
                          <Tag v="default">
                            退貨/退款 {spS.refundN} 筆 · {fmt$(spS.refundG)}{" "}
                            未計入
                          </Tag>
                        )}
                      </div>
                      <div
                        style={{
                          display: "flex",
                          alignItems: "flex-end",
                          justifyContent: "space-between",
                          flexWrap: "wrap",
                          gap: 24,
                        }}
                      >
                        <div>
                          <div
                            style={{
                              fontSize: 12,
                              fontWeight: 600,
                              color: "var(--t3)",
                              marginBottom: 4,
                              letterSpacing: "0.06em",
                            }}
                          >
                            最終結算淨利 · NET PROFIT
                          </div>
                          <div
                            className="hero-num"
                            style={{
                              lineHeight: 1,
                              fontWeight: 700,
                              letterSpacing: "-0.04em",
                              fontFamily: mono,
                              color:
                                spS.afterComm >= 0 ? "var(--t1)" : "var(--dn)",
                            }}
                          >
                            {fmt$(spS.afterComm)}
                          </div>
                          {spS.comm > 0 && (
                            <div
                              style={{
                                fontSize: 12,
                                color: "var(--t3)",
                                marginTop: 6,
                              }}
                            >
                              分潤前：{fmt$(spS.tNetPro)}
                            </div>
                          )}
                          <PeriodCompare
                            monthly={spMonthly}
                            sY={sY}
                            sM={effM}
                          />
                        </div>
                        <div style={{ textAlign: "right" }}>
                          <div
                            style={{
                              fontSize: 12,
                              fontWeight: 600,
                              color: "var(--t3)",
                            }}
                          >
                            淨利率
                          </div>
                          <div
                            className="hero-pct"
                            style={{
                              fontWeight: 700,
                              fontFamily: mono,
                              lineHeight: 1,
                              color:
                                spS.netMargin >= spS.targetNet
                                  ? "var(--up)"
                                  : spS.netMargin >= spS.targetNet * 0.6
                                  ? "var(--orange)"
                                  : spS.netMargin > 0
                                  ? "var(--wn)"
                                  : "var(--dn)",
                            }}
                          >
                            {fmtP(spS.netMargin)}
                          </div>
                          <div
                            style={{
                              fontSize: 11,
                              color: "var(--t3)",
                              marginTop: 4,
                            }}
                          >
                            目標 {fmtP(spS.targetNet)}
                          </div>
                        </div>
                      </div>
                      {/* Waterfall：通路後毛利 − 營業費(含廣告) − 稅賦 −（分潤）= 淨利，逐項可驗算 */}
                      <div
                        style={{
                          marginTop: 28,
                          borderTop: "1px solid var(--s3)",
                          paddingTop: 20,
                        }}
                      >
                        <div
                          style={{
                            fontSize: 11,
                            fontWeight: 700,
                            color: "var(--t3)",
                            marginBottom: 14,
                            letterSpacing: "0.06em",
                          }}
                        >
                          損益分解 · WATERFALL
                        </div>
                        <div
                          style={{
                            display: "flex",
                            flexWrap: "wrap",
                            alignItems: "flex-end",
                            gap: 0,
                          }}
                        >
                          {[
                            {
                              l: "通路後毛利",
                              v: spS.tG - spS.tF - spS.tC,
                            },
                            { l: "營業費", v: -spS.tOp, neg: true },
                            { l: "稅賦", v: -spS.tTx, neg: true },
                            ...(spS.comm > 0
                              ? [{ l: "分潤", v: -spS.comm, neg: true }]
                              : []),
                            { l: "淨利", v: spS.afterComm, bold: true },
                          ].map((item, i, arr) => (
                            <React.Fragment key={i}>
                              <div
                                style={{
                                  flex: "1 1 0",
                                  minWidth: 80,
                                  textAlign: "center",
                                  padding: "0 6px",
                                }}
                              >
                                <div
                                  style={{
                                    fontSize: 11,
                                    color: "var(--t3)",
                                    fontWeight: 600,
                                    marginBottom: 4,
                                  }}
                                >
                                  {item.l}
                                </div>
                                <div
                                  style={{
                                    fontSize: 18,
                                    fontWeight: item.bold ? 800 : 600,
                                    fontFamily: mono,
                                    letterSpacing: "-0.02em",
                                    color: item.bold
                                      ? spS.afterComm >= 0
                                        ? "var(--up)"
                                        : "var(--dn)"
                                      : item.neg
                                      ? "var(--dn)"
                                      : "var(--t1)",
                                  }}
                                >
                                  {fmt$(item.v)}
                                </div>
                              </div>
                              {i < arr.length - 1 && (
                                <div
                                  style={{
                                    color: "var(--s4)",
                                    fontSize: 16,
                                    padding: "0 2px",
                                    alignSelf: "center",
                                  }}
                                >
                                  ›
                                </div>
                              )}
                            </React.Fragment>
                          ))}
                        </div>
                      </div>
                    </div>
                    {/* KPI 上排 4 */}
                    <div className="g4 f2">
                      {[
                        {
                          l: "精準營收基底",
                          v: fmt$(spS.tG),
                          c: "var(--t1)",
                          h: `賣場券 ${fmt$(spS.tV)} 已含於此 ｜ 手續費 -${fmt$(
                            spS.tF
                          )}`,
                        },
                        {
                          l: "預估總入帳",
                          v: fmt$(spS.tG - spS.tF),
                          c: "var(--blue)",
                          h: "扣除手續費（賣場券已在營收扣過，不重複扣）",
                        },
                        {
                          l: "商品毛利",
                          v: fmt$(spS.tG - spS.tC),
                          c: "var(--up)",
                          h: `毛利率 ${fmtP(
                            spS.tG > 0 ? (spS.tG - spS.tC) / spS.tG : 0
                          )}`,
                        },
                        {
                          l: "結算現金",
                          v: fmt$(spS.afterComm),
                          c:
                            spS.afterComm >= 0
                              ? "var(--sp-accent)"
                              : "var(--dn)",
                          h:
                            spS.comm > 0
                              ? `-${fmt$(spS.comm)} 分潤`
                              : `淨利率 ${fmtP(spS.netMargin)}`,
                          hc: spS.comm > 0 ? "var(--purple)" : "var(--t4)",
                          border: "var(--sp-accent)",
                        },
                      ].map((k, i) => (
                        <div
                          key={i}
                          style={{
                            background: "var(--s1)",
                            border: `1px solid ${k.border || "var(--s3)"}`,
                            borderRadius: 14,
                            padding: "22px 24px",
                            borderLeft: k.border
                              ? `3px solid ${k.border}`
                              : undefined,
                          }}
                        >
                          <Lbl>{k.l}</Lbl>
                          <div
                            style={{
                              fontSize: 30,
                              fontWeight: 700,
                              fontFamily: mono,
                              letterSpacing: "-0.03em",
                              color: k.c,
                              marginTop: 6,
                            }}
                          >
                            {k.v}
                          </div>
                          <div
                            style={{
                              fontSize: 11,
                              color: k.hc || "var(--t4)",
                              marginTop: 8,
                            }}
                          >
                            {k.h}
                          </div>
                        </div>
                      ))}
                    </div>
                    {/* KPI 下排 4 */}
                    <div className="g4 f3">
                      {[
                        {
                          l: "商品毛利率",
                          v: fmtP(spS.tG > 0 ? (spS.tG - spS.tC) / spS.tG : 0),
                          c: (() => {
                            const r =
                              spS.tG > 0 ? (spS.tG - spS.tC) / spS.tG : 0;
                            return r >= 0.68
                              ? "var(--up)"
                              : r >= 0.63
                              ? "var(--accent)"
                              : "var(--dn)";
                          })(),
                          note: (() => {
                            const r =
                              spS.tG > 0 ? (spS.tG - spS.tC) / spS.tG : 0;
                            return r >= 0.68
                              ? "✓ 超越目標 68%，表現優異"
                              : r >= 0.65
                              ? "✓ 達標，目標 65~68%"
                              : r >= 0.63
                              ? "⚠ 正常帶下緣，目標 65~68%"
                              : "⚠ 低於警戒線 63%，檢視定價";
                          })(),
                        },
                        {
                          l: "真實抽成率",
                          v: fmtP(spS.feeRate),
                          c: "var(--blue)",
                          note: "結算檔實際手續費＋蝦幣回饋 ÷ 營收（實測值，非固定費率）",
                        },
                        {
                          l: "優惠券發放率",
                          v: fmtP(spS.voucherRate),
                          c:
                            spS.voucherRate > 0.03
                              ? "var(--dn)"
                              : spS.voucherRate > 0.025
                              ? "var(--wn)"
                              : spS.voucherRate < 0.008
                              ? "var(--t3)"
                              : spS.voucherRate <= 0.015
                              ? "var(--up)"
                              : "var(--purple)",
                          note:
                            spS.voucherRate > 0.03
                              ? "⚠ 品牌警戒！單月 >3%，立即介入"
                              : spS.voucherRate > 0.025
                              ? "⚠ 超出警戒線 2.5%，啟動檢討"
                              : spS.voucherRate < 0.008
                              ? "低於下限 0.8%，本期幾乎未配券（確認檔期是否漏配）"
                              : spS.voucherRate <= 0.015
                              ? "✓ 在目標範圍 0.8~1.5% 內"
                              : "注意：接近警戒線 2.5%",
                        },
                        {
                          /* 階梯制（2026-07-08 雙平台淨利管理基準）：
                             50 優異／49.2 達標線／48 觀察／47 調整線／介入線 */
                          l: "通路後毛利率",
                          v: fmtP(spS.grossMargin),
                          c:
                            spS.grossMargin >= 0.492
                              ? "var(--up)"
                              : spS.grossMargin >= 0.48
                              ? "var(--wn)"
                              : "var(--dn)",
                          note:
                            spS.grossMargin >= 0.5
                              ? "✓ 超越 50%，表現優異（換算稅後 ≈13.8%）"
                              : spS.grossMargin >= 0.492
                              ? "✓ 達標，達標線 49.2%"
                              : spS.grossMargin >= 0.48
                              ? "⚠ 觀察帶，追蹤商品組合與讓利深度"
                              : spS.grossMargin >= 0.47
                              ? "⚠ 低於調整線 48%：調商品結構，原則不漲價"
                              : "⚠ 低於介入線 47%：立即檢視平台費用與優惠券",
                        },
                      ].map((k, i) => (
                        <div
                          key={i}
                          style={{
                            background: "var(--s1)",
                            border: "1px solid var(--s3)",
                            borderRadius: 14,
                            padding: "20px 22px",
                          }}
                        >
                          <div
                            style={{
                              fontSize: 12,
                              fontWeight: 600,
                              color: "var(--t3)",
                              marginBottom: 8,
                            }}
                          >
                            {k.l}
                          </div>
                          <div
                            style={{
                              fontSize: 26,
                              fontWeight: 700,
                              fontFamily: mono,
                              letterSpacing: "-0.03em",
                              color: k.c,
                            }}
                          >
                            {k.v}
                          </div>
                          <div
                            style={{
                              fontSize: 11,
                              color: "var(--t4)",
                              marginTop: 8,
                              display: "flex",
                              alignItems: "center",
                              gap: 4,
                            }}
                          >
                            <span
                              style={{
                                display: "inline-block",
                                width: 5,
                                height: 5,
                                borderRadius: "50%",
                                background: k.c,
                                flexShrink: 0,
                              }}
                            />
                            {k.note}
                          </div>
                        </div>
                      ))}
                    </div>
                    {/* KPI 第三排：券策略錨點與單筆貢獻 */}
                    <div className="g3 f3">
                      {[
                        {
                          l: "平均客單價 AOV",
                          v: fmt$(spS.avgAOV),
                          c: "var(--t1)",
                          h: "優惠券門檻設計的錨點指標（含補貼還原）",
                        },
                        {
                          l: "單筆平均淨利",
                          v: fmt$(spS.avgNetPer),
                          c: spS.avgNetPer >= 0 ? "var(--up)" : "var(--dn)",
                          h: "已扣分潤後平均每單實際貢獻",
                        },
                        {
                          l: "退貨/退款排除",
                          v: `${spS.refundN} 筆`,
                          c: spS.refundN > 0 ? "var(--wn)" : "var(--t3)",
                          h: `${fmt$(spS.refundG)} 未計入營收`,
                        },
                      ].map((k, i) => (
                        <div
                          key={i}
                          style={{
                            background: "var(--s1)",
                            border: "1px solid var(--s3)",
                            borderRadius: 14,
                            padding: "20px 22px",
                          }}
                        >
                          <div
                            style={{
                              fontSize: 12,
                              fontWeight: 600,
                              color: "var(--t3)",
                              marginBottom: 8,
                            }}
                          >
                            {k.l}
                          </div>
                          <div
                            style={{
                              fontSize: 26,
                              fontWeight: 700,
                              fontFamily: mono,
                              letterSpacing: "-0.03em",
                              color: k.c,
                            }}
                          >
                            {k.v}
                          </div>
                          <div
                            style={{
                              fontSize: 11,
                              color: "var(--t4)",
                              marginTop: 8,
                            }}
                          >
                            {k.h}
                          </div>
                        </div>
                      ))}
                    </div>
                  </>
                )}

                {/* ── Cost Matrix（共用卡） ── */}
                {costMatrixCard}

                {/* ── Order Table ── */}
                <div
                  className="f5"
                  style={{
                    background: "var(--s1)",
                    border: "1px solid var(--s3)",
                    borderRadius: 16,
                    padding: 24,
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      flexWrap: "wrap",
                      justifyContent: "space-between",
                      gap: 12,
                      alignItems: "center",
                      marginBottom: 14,
                    }}
                  >
                    <div
                      style={{ display: "flex", alignItems: "center", gap: 8 }}
                    >
                      <BarChart3 size={16} color="var(--t3)" />
                      <span style={{ fontSize: 14, fontWeight: 700 }}>
                        單筆訂單決策明細
                      </span>
                      <span style={{ fontSize: 11, color: "var(--dn)" }}>
                        虧損 {isSL ? slData?.lossCount : spS?.lossN} 筆
                      </span>
                    </div>
                    <div
                      style={{
                        display: "flex",
                        gap: 6,
                        alignItems: "center",
                        flexWrap: "wrap",
                      }}
                    >
                      <Btn onClick={expReport}>
                        <Download size={12} /> 匯出報表
                      </Btn>
                      <div style={{ position: "relative" }}>
                        <Search
                          size={13}
                          color="var(--t4)"
                          style={{
                            position: "absolute",
                            left: 10,
                            top: "50%",
                            transform: "translateY(-50%)",
                          }}
                        />
                        <input
                          type="text"
                          placeholder="搜尋單號或商品..."
                          aria-label="搜尋單號或商品"
                          value={search}
                          onChange={(e) => setSearch(e.target.value)}
                          style={{
                            ...inp,
                            width: 180,
                            textAlign: "left",
                            paddingLeft: 30,
                            borderRadius: 10,
                            padding: "7px 12px 7px 30px",
                            fontSize: 12,
                          }}
                        />
                      </div>
                      <label
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 6,
                          fontSize: 11,
                          fontWeight: 600,
                          color: "var(--t3)",
                          cursor: "pointer",
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={lossOnly}
                          onChange={(e) => setLossOnly(e.target.checked)}
                          style={{ accentColor: "var(--dn)" }}
                        />{" "}
                        只看虧損（本期）
                      </label>
                    </div>
                  </div>
                  <div
                    style={{
                      overflowX: "auto",
                      overflowY: "auto",
                      maxHeight: 500,
                      border: "1px solid var(--s3)",
                      borderRadius: 12,
                    }}
                  >
                    {/* 手機隱藏通路費用／成本／毛利三欄（官網 6 欄、蝦皮 7 欄，欄序不同用兩支 class） */}
                    <table
                      className={isSL ? "tb-ord-sl" : "tb-ord-sp"}
                      style={{
                        width: "100%",
                        borderCollapse: "collapse",
                        minWidth: 820,
                      }}
                    >
                      <thead>
                        <tr>
                          <SortTh
                            sortKey="date"
                            currentSort={orderSort}
                            onSort={(k) =>
                              setOrderSort((p) => ({
                                key: k,
                                dir:
                                  p.key === k
                                    ? p.dir === "desc"
                                      ? "asc"
                                      : "desc"
                                    : "desc",
                              }))
                            }
                          >
                            單號
                          </SortTh>
                          {!isSL && (
                            <th
                              scope="col"
                              style={{ ...th, textAlign: "left" }}
                            >
                              商品
                            </th>
                          )}
                          <SortTh
                            sortKey="revenue"
                            currentSort={orderSort}
                            onSort={(k) =>
                              setOrderSort((p) => ({
                                key: k,
                                dir:
                                  p.key === k
                                    ? p.dir === "desc"
                                      ? "asc"
                                      : "desc"
                                    : "desc",
                              }))
                            }
                            align="right"
                          >
                            營收
                          </SortTh>
                          <SortTh
                            sortKey="fee"
                            currentSort={orderSort}
                            onSort={(k) =>
                              setOrderSort((p) => ({
                                key: k,
                                dir:
                                  p.key === k
                                    ? p.dir === "desc"
                                      ? "asc"
                                      : "desc"
                                    : "desc",
                              }))
                            }
                            align="right"
                          >
                            通路費用
                          </SortTh>
                          <SortTh
                            sortKey="cost"
                            currentSort={orderSort}
                            onSort={(k) =>
                              setOrderSort((p) => ({
                                key: k,
                                dir:
                                  p.key === k
                                    ? p.dir === "desc"
                                      ? "asc"
                                      : "desc"
                                    : "desc",
                              }))
                            }
                            align="right"
                          >
                            成本
                          </SortTh>
                          <SortTh
                            sortKey="profit"
                            currentSort={orderSort}
                            onSort={(k) =>
                              setOrderSort((p) => ({
                                key: k,
                                dir:
                                  p.key === k
                                    ? p.dir === "desc"
                                      ? "asc"
                                      : "desc"
                                    : "desc",
                              }))
                            }
                            align="right"
                          >
                            毛利
                          </SortTh>
                          <SortTh
                            sortKey="net"
                            currentSort={orderSort}
                            onSort={(k) =>
                              setOrderSort((p) => ({
                                key: k,
                                dir:
                                  p.key === k
                                    ? p.dir === "desc"
                                      ? "asc"
                                      : "desc"
                                    : "desc",
                              }))
                            }
                            align="right"
                          >
                            最終淨利
                          </SortTh>
                        </tr>
                      </thead>
                      <tbody>
                        {pagedOrders.length > 0 ? (
                          pagedOrders.map((o) => {
                            const isLoss = isSL
                              ? o.net < 0
                              : o.finalNetProfit < 0;
                            const rev = isSL ? o.revenue : o.localGross;
                            const fee = o.channelFee;
                            const cost = isSL ? o.oCost : o.orderCost;
                            const gross = isSL
                              ? o.currentOrderContribution
                              : o.grossProfit;
                            const net = isSL ? o.net : o.finalNetProfit;
                            const isOpen = expandedId === o.orderId;
                            return (
                              <React.Fragment key={o.orderId}>
                              <tr
                                className={isLoss ? "rl" : ""}
                                tabIndex={0}
                                aria-expanded={isOpen}
                                onClick={() =>
                                  setExpandedId(isOpen ? null : o.orderId)
                                }
                                onKeyDown={(e) => {
                                  if (e.key === "Enter" || e.key === " ") {
                                    e.preventDefault();
                                    setExpandedId(isOpen ? null : o.orderId);
                                  }
                                }}
                                style={{ cursor: "pointer" }}
                              >
                                <td style={{ ...td2 }}>
                                  <div
                                    style={{
                                      fontWeight: 600,
                                      fontSize: 12,
                                      display: "flex",
                                      alignItems: "center",
                                      gap: 4,
                                    }}
                                  >
                                    {isOpen ? (
                                      <ChevronUp
                                        size={11}
                                        color="var(--t3)"
                                      />
                                    ) : (
                                      <ChevronDown
                                        size={11}
                                        color="var(--t3)"
                                      />
                                    )}
                                    {o.date}
                                  </div>
                                  <div
                                    style={{
                                      fontSize: 10,
                                      color: "var(--t3)",
                                      marginTop: 2,
                                      fontFamily: mono,
                                    }}
                                  >
                                    {o.orderId}
                                  </div>
                                </td>
                                {!isSL && (
                                  <td style={{ ...td2, maxWidth: 180 }}>
                                    <div
                                      style={{
                                        fontSize: 11,
                                        fontWeight: 600,
                                        color: "var(--t2)",
                                        whiteSpace: "nowrap",
                                        overflow: "hidden",
                                        textOverflow: "ellipsis",
                                        maxWidth: 170,
                                      }}
                                      title={(o.items || [])
                                        .map((i) => i.name)
                                        .join("、")}
                                    >
                                      {(o.items || []).length === 1
                                        ? o.items[0].name
                                        : `${o.items?.[0]?.name || "—"} 等 ${
                                            o.items?.length || 0
                                          } 件`}
                                    </div>
                                  </td>
                                )}
                                <td
                                  style={{
                                    ...td2,
                                    textAlign: "right",
                                    fontFamily: mono,
                                    fontWeight: 600,
                                  }}
                                >
                                  {fmt$(rev)}
                                </td>
                                <td
                                  style={{
                                    ...td2,
                                    textAlign: "right",
                                    fontFamily: mono,
                                    color: "var(--dn)",
                                  }}
                                >
                                  -{fmt$(fee)}
                                </td>
                                <td
                                  style={{
                                    ...td2,
                                    textAlign: "right",
                                    fontFamily: mono,
                                    color: "var(--dn)",
                                  }}
                                >
                                  -{fmt$(cost)}
                                </td>
                                <td
                                  style={{
                                    ...td2,
                                    textAlign: "right",
                                    fontFamily: mono,
                                    fontWeight: 600,
                                  }}
                                >
                                  {fmt$(gross)}
                                </td>
                                <td
                                  style={{
                                    ...td2,
                                    textAlign: "right",
                                    fontFamily: mono,
                                    fontWeight: 800,
                                    color: isLoss
                                      ? "var(--dn)"
                                      : "var(--accent)",
                                  }}
                                >
                                  {fmt$(net)}
                                </td>
                              </tr>
                              {isOpen && (
                                <tr>
                                  <td
                                    colSpan={isSL ? 6 : 7}
                                    style={{
                                      ...td2,
                                      background: "var(--s2)",
                                      padding: "16px 20px",
                                    }}
                                  >
                                    <OrderDetail
                                      order={o}
                                      isSL={isSL}
                                      slFp={slFp}
                                      slCosts={slEffCosts}
                                      spCosts={spEffCosts}
                                    />
                                  </td>
                                </tr>
                              )}
                              </React.Fragment>
                            );
                          })
                        ) : (
                          <tr>
                            <td
                              colSpan={isSL ? 6 : 7}
                              style={{
                                ...td2,
                                textAlign: "center",
                                color: "var(--t4)",
                                padding: 40,
                              }}
                            >
                              找不到符合條件的訂單
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                  {/* Pagination */}
                  {filteredOrders.length > pageSize && (
                    <div
                      style={{
                        padding: "12px 4px 0",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        flexWrap: "wrap",
                        gap: 8,
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 8,
                        }}
                      >
                        <span
                          style={{
                            fontSize: 11,
                            color: "var(--t3)",
                            fontFamily: mono,
                          }}
                        >
                          {curPage * pageSize + 1}–
                          {Math.min(
                            (curPage + 1) * pageSize,
                            filteredOrders.length
                          )}{" "}
                          / {filteredOrders.length} 筆
                        </span>
                        <span style={{ fontSize: 10, color: "var(--t4)" }}>
                          每頁
                        </span>
                        <select
                          value={pageSize}
                          aria-label="每頁筆數"
                          onChange={(e) => {
                            setPageSize(Number(e.target.value));
                            setPage(0);
                          }}
                          style={{
                            ...sel,
                            padding: "3px 8px",
                            fontSize: 11,
                            fontFamily: mono,
                          }}
                        >
                          {[20, 30, 50, 100].map((n) => (
                            <option key={n} value={n}>
                              {n}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 4,
                        }}
                      >
                        {[
                          {
                            label: "«",
                            aria: "第一頁",
                            action: () => setPage(0),
                          },
                          {
                            label: "‹",
                            aria: "上一頁",
                            action: () =>
                              setPage(Math.max(0, curPage - 1)),
                          },
                          null,
                          {
                            label: "›",
                            aria: "下一頁",
                            action: () =>
                              setPage(Math.min(totalPages - 1, curPage + 1)),
                          },
                          {
                            label: "»",
                            aria: "最後一頁",
                            action: () => setPage(totalPages - 1),
                          },
                        ].map((btn, i) =>
                          btn === null ? (
                            <span
                              key={i}
                              style={{
                                fontSize: 12,
                                fontWeight: 800,
                                color: "var(--t1)",
                                fontFamily: mono,
                                padding: "0 10px",
                              }}
                            >
                              {curPage + 1} / {totalPages}
                            </span>
                          ) : (
                            <Btn
                              key={i}
                              v="ghost"
                              aria-label={btn.aria}
                              onClick={btn.action}
                              style={{
                                padding: "4px 8px",
                                fontSize: 12,
                                minWidth: 32,
                                justifyContent: "center",
                              }}
                            >
                              {btn.label}
                            </Btn>
                          )
                        )}
                      </div>
                    </div>
                  )}
                </div>
              </>
            )}
            </ErrorBoundary>
          </main>
        </div>
      </div>

      {/* Confirm Dialog */}
      {confirmBox && (
        <div
          role="dialog"
          aria-modal="true"
          aria-label={confirmBox.title}
          onClick={() => setConfirmBox(null)}
          style={{
            position: "fixed",
            inset: 0,
            zIndex: 300,
            background: "rgba(0,0,0,0.45)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            padding: 20,
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            onKeyDown={(e) => {
              if (e.key === "Escape") setConfirmBox(null);
            }}
            style={{
              background: "var(--s1)",
              border: "1px solid var(--s3)",
              borderRadius: 14,
              padding: "22px 24px",
              maxWidth: 400,
              width: "100%",
              boxShadow: "0 24px 80px rgba(0,0,0,.35)",
              animation: "dlgIn .18s ease both",
            }}
          >
            <div
              style={{
                fontSize: 14,
                fontWeight: 800,
                marginBottom: 8,
                color: "var(--t1)",
              }}
            >
              {confirmBox.title}
            </div>
            <div
              style={{
                fontSize: 12,
                color: "var(--t2)",
                lineHeight: 1.7,
                whiteSpace: "pre-wrap",
                marginBottom: 18,
              }}
            >
              {confirmBox.message}
            </div>
            <div
              style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}
            >
              {/* 破壞性操作預設焦點給「取消」，避免殘留 Enter 誤觸清資料 */}
              <Btn
                autoFocus={!!confirmBox.danger}
                onClick={() => setConfirmBox(null)}
              >
                取消
              </Btn>
              <Btn
                v={confirmBox.danger ? "danger" : "primary"}
                autoFocus={!confirmBox.danger}
                onClick={() => {
                  const fn = confirmBox.onOk;
                  setConfirmBox(null);
                  fn?.();
                }}
              >
                確定
              </Btn>
            </div>
          </div>
        </div>
      )}

      {/* Toast Container */}
      <div
        style={{
          position: "fixed",
          bottom: 20,
          right: 20,
          zIndex: 9999,
          display: "flex",
          flexDirection: "column",
          gap: 8,
          maxWidth: 360,
          pointerEvents: "none",
        }}
      >
        {toasts.map((t) => {
          const borderCol =
            t.type === "success"
              ? "var(--up)"
              : t.type === "error"
              ? "var(--dn)"
              : t.type === "warning"
              ? "var(--wn)"
              : "var(--orange)";
          return (
            <div
              key={t.id}
              role="status"
              aria-live="polite"
              style={{
                pointerEvents: "auto",
                background: "var(--s1)",
                border: "1px solid var(--s3)",
                borderLeft: `3px solid ${borderCol}`,
                borderRadius: 10,
                padding: "12px 16px",
                boxShadow: "0 8px 30px rgba(0,0,0,.15)",
                display: "flex",
                alignItems: "center",
                gap: 10,
                animation: t.removing
                  ? "toastOut .3s ease forwards"
                  : "toastIn .3s ease",
                overflow: "hidden",
                position: "relative",
              }}
            >
              {t.type === "success" ? (
                <CheckCircle2 size={14} color="var(--up)" />
              ) : t.type === "error" ? (
                <AlertTriangle size={14} color="var(--dn)" />
              ) : (
                <Info size={14} color="var(--orange)" />
              )}
              <span
                style={{
                  fontSize: 12,
                  fontWeight: 700,
                  color: "var(--t1)",
                  flex: 1,
                }}
              >
                {t.msg}
              </span>
              {t.action && (
                <button
                  onClick={() => {
                    t.action();
                    removeToast(t.id);
                  }}
                  style={{
                    border: `1px solid ${borderCol}`,
                    background: "var(--s2)",
                    color: borderCol,
                    borderRadius: 6,
                    padding: "4px 10px",
                    fontSize: 10,
                    fontWeight: 800,
                    cursor: "pointer",
                  }}
                >
                  {t.actionLabel || "復原"}
                </button>
              )}
              <button
                onClick={() => removeToast(t.id)}
                aria-label="關閉通知"
                style={{
                  border: "none",
                  background: "none",
                  color: "var(--t4)",
                  cursor: "pointer",
                  padding: 2,
                  display: "flex",
                }}
              >
                <X size={12} />
              </button>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ─── Error Boundary：單筆資料異常不至於整頁白屏 ────────────── */
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { err: null };
  }
  static getDerivedStateFromError(err) {
    return { err };
  }
  componentDidCatch(err, info) {
    console.error("[ErrorBoundary]", err, info);
  }
  render() {
    if (this.state.err) {
      /* inner＝只包主內容區：側欄（匯入／重置本期／備份）保持可用，
         使用者才有自救路徑，不會被一筆壞資料鎖死整頁 */
      if (this.props.inner) {
        return (
          <div
            style={{
              border: "1px solid #E4B8B8",
              background: "rgba(200,64,64,0.05)",
              borderRadius: 12,
              padding: 20,
              display: "flex",
              flexDirection: "column",
              gap: 10,
              color: "var(--t1)",
            }}
          >
            <div style={{ fontSize: 14, fontWeight: 800 }}>
              ⚠️ 這個平台的畫面算不出來
            </div>
            <div style={{ fontSize: 12, color: "var(--t3)", lineHeight: 1.7 }}>
              資料仍在雲端與本機，沒有遺失。多半是某一筆訂單的欄位壞掉。
              可以到左側「重置本期」清掉這一期再重新匯入報表，或先切到別的平台。
            </div>
            <div
              style={{
                fontSize: 11,
                color: "var(--t4)",
                fontFamily: "monospace",
                wordBreak: "break-all",
              }}
            >
              {String(this.state.err?.message || this.state.err)}
            </div>
            <div>
              <Btn onClick={() => this.setState({ err: null })}>重試</Btn>
            </div>
          </div>
        );
      }
      return (
        <div
          style={{
            minHeight: "100vh",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
            gap: 14,
            background: "#F8F8F6",
            color: "#1A1A18",
            fontFamily: "'Inter','Noto Sans TC',sans-serif",
            padding: 24,
            textAlign: "center",
          }}
        >
          <div style={{ fontSize: 40 }}>⚠️</div>
          <div style={{ fontSize: 16, fontWeight: 800 }}>
            畫面發生錯誤，資料不受影響
          </div>
          <div
            style={{
              fontSize: 12,
              color: "#8E8E84",
              maxWidth: 480,
              lineHeight: 1.7,
              fontFamily: "monospace",
              wordBreak: "break-all",
            }}
          >
            {String(this.state.err?.message || this.state.err)}
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button
              onClick={() => this.setState({ err: null })}
              style={{
                border: "1px solid #D8D8D2",
                background: "#FFFFFF",
                borderRadius: 8,
                padding: "8px 18px",
                fontSize: 12,
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              重試
            </button>
            <button
              onClick={() => window.location.reload()}
              style={{
                border: "1px solid rgba(26,107,60,0.18)",
                background: "rgba(26,107,60,0.06)",
                color: "#1A6B3C",
                borderRadius: 8,
                padding: "8px 18px",
                fontSize: 12,
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              重新整理頁面
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

export default function App() {
  return (
    <ErrorBoundary>
      <ProfitCenter />
    </ErrorBoundary>
  );
}
