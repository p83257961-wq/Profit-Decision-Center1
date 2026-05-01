import React, {
  useState,
  useMemo,
  useEffect,
  useRef,
  useCallback,
} from "react";
import {
  Building2,
  Trash2,
  BarChart3,
  Search,
  Package,
  FileSpreadsheet,
  TrendingUp,
  CheckCircle2,
  AlertTriangle,
  Activity,
  Settings,
  Target,
  Download,
  UploadCloud,
  Wallet,
  Truck,
  CreditCard,
  ShoppingCart,
  AlertCircle,
  PieChart,
  RotateCcw,
  HelpCircle,
  Lock,
  Loader2,
  Check,
  Gift,
  Zap,
  Layers,
  Sun,
  Moon,
  Award,
  Star,
  ChevronDown,
  ChevronUp,
  ChevronRight,
  Filter,
  Users,
  X,
  Info,
  Menu,
  Store,
} from "lucide-react";
import {
  AreaChart,
  Area,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip as RTooltip,
  ResponsiveContainer,
  Cell,
  ComposedChart,
  Line,
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
  getDocs,
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
const FSPC_SL = { collection: "warrooms", docId: "shopline_profit_center_v16" };
const FSPC_SP = { collection: "warrooms", docId: "shopee_profit_center_v2" };
const FSPC_SL_ORD = { collection: "warrooms", docId: "unified_sl_orders_v1" };
const FSPC_SP_ORD = { collection: "warrooms", docId: "unified_sp_orders_v1" };

/* 新的按月拆分 collection */
const SL_MONTHLY_COLL = "sl_orders_monthly";
const SP_MONTHLY_COLL = "sp_orders_monthly";

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
const MONO = "'JetBrains Mono',monospace";

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
};
const SL_SHIPPING_RATES = {
  "7-11": 65,
  全家: 65,
  宅配: 120,
  順豐: 250,
  SF: 250,
};
const SL_INTL_METHODS = ["EMS", "FEDEX", "中國", "新加坡", "國外"];

const DEFAULT_FP_SL = {
  platformFeeRate: "1.0",
  opExpense: "30.0",
  tax: "6.2",
  targetNet: "17.0",
};
const DEFAULT_FP_SP = { opExpense: "30.0", tax: "6.2", targetNet: "14.0" };

const SK = {
  platform: "upc_platform_v1",
  slFp: "upc_sl_fee_params_v1",
  spFp: "upc_sp_fee_params_v1",
  slCosts: "upc_sl_costs_v1",
  spCosts: "upc_sp_costs_v1",
  slOrders: "upc_sl_orders_v1",
  spOrders: "upc_sp_orders_v1",
  commissions: "upc_commissions_v1",
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
const numOrZero = (v) => {
  const n = parseFloat(String(v ?? "").replace(/,/g, ""));
  return Number.isFinite(n) ? n : 0;
};
const safeText = (v) => String(v ?? "").trim();
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
  } catch {}
};
const gcid = () => {
  const K = "upc_client_id_v1",
    e = window.localStorage.getItem(K);
  if (e) return e;
  const n = `c_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  window.localStorage.setItem(K, n);
  return n;
};
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

/* ─── CSS ─────────────────────────────────────────────────────── */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500;600;700;800&family=Noto+Sans+TC:wght@400;500;600;700;800;900&family=Inter:wght@400;500;600;700;800&display=swap');
[data-theme="light"]{
  --bg:#FFFFFF;--s1:#FFFFFF;--s2:#F8F8F6;--s3:#EAEAE6;--s4:#D8D8D2;
  --t1:#1A1A18;--t2:#5C5C54;--t3:#8E8E84;--t4:#B8B8AE;
  --accent:#1A6B3C;--accent-dim:rgba(26,107,60,0.06);--accent-bdr:rgba(26,107,60,0.18);--accent-text:#1A6B3C;
  --up:#1A6B3C;--up-dim:rgba(26,107,60,0.06);--up-bdr:rgba(26,107,60,0.18);
  --dn:#C0392B;--dn-dim:rgba(192,57,43,0.05);--dn-bdr:rgba(192,57,43,0.18);
  --wn:#B7600A;--wn-dim:rgba(183,96,10,0.05);--wn-bdr:rgba(183,96,10,0.18);
  --blue:#2E6DA4;--purple:#7B5EA7;--orange:#D4820A;--gold:#8B6914;
  --header-bg:rgba(255,255,255,0.92);
  --bar-track:#EAEAE6;
  --row-loss:rgba(192,57,43,0.04);
  --sp-accent:#EE4D2D;--sp-accent-dim:rgba(238,77,45,0.06);--sp-accent-bdr:rgba(238,77,45,0.2);
}
[data-theme="dark"]{
  --bg:#080A0E;--s1:#0E1117;--s2:#151921;--s3:#1C212B;--s4:#282D38;
  --t1:#E8E6E1;--t2:#8B8D93;--t3:#505259;--t4:#35373D;
  --accent:#2ECC71;--accent-dim:rgba(46,204,113,0.08);--accent-bdr:rgba(46,204,113,0.2);--accent-text:#2ECC71;
  --up:#2ECC71;--up-dim:rgba(46,204,113,0.08);--up-bdr:rgba(46,204,113,0.2);
  --dn:#E74C3C;--dn-dim:rgba(231,76,60,0.08);--dn-bdr:rgba(231,76,60,0.2);
  --wn:#E67E22;--wn-dim:rgba(230,126,34,0.08);--wn-bdr:rgba(230,126,34,0.2);
  --blue:#3498DB;--purple:#9B7FCA;--orange:#F0A030;--gold:#C9A84C;
  --header-bg:rgba(8,10,14,0.88);
  --bar-track:#1C212B;
  --row-loss:rgba(231,76,60,0.07);
  --sp-accent:#FF6533;--sp-accent-dim:rgba(255,101,51,0.08);--sp-accent-bdr:rgba(255,101,51,0.22);
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
::-webkit-scrollbar{width:4px;height:4px;}
::-webkit-scrollbar-track{background:transparent;}
::-webkit-scrollbar-thumb{background:var(--s4);border-radius:99px;}
@keyframes fadeUp{from{opacity:0;transform:translateY(16px);}to{opacity:1;transform:translateY(0);}}
@keyframes fadeIn{from{opacity:0;}to{opacity:1;}}
@keyframes spin{to{transform:rotate(360deg);}}
@keyframes notifIn{from{opacity:0;transform:translate(-50%,-10px);}to{opacity:1;transform:translate(-50%,0);}}
@keyframes toastIn{from{opacity:0;transform:translateX(100%);}to{opacity:1;transform:translateX(0);}}
@keyframes toastOut{from{opacity:1;}to{opacity:0;transform:translateX(100%);}}
@keyframes timerBar{from{width:100%;}to{width:0%;}}
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
.g2{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;}
@media(max-width:900px){.gm{grid-template-columns:1fr;}}
@media(max-width:1000px){.g4{grid-template-columns:repeat(2,1fr);}.g3{grid-template-columns:repeat(2,1fr);}}
@media(max-width:600px){.g4,.g3,.g2{grid-template-columns:1fr;}}
input,select,button{font-family:'Inter','Noto Sans TC',sans-serif;}
button{cursor:pointer;transition:all .12s;}
button:hover{filter:brightness(1.06);}
button:active{transform:scale(.97);}
tr{transition:background .1s;}
tr:hover td{background:var(--s2)!important;}
.rw td{background:var(--wn-dim)!important;}
.rl{background:var(--row-loss)!important;}
.iw{border-color:var(--wn)!important;}
.iok{border-color:var(--up-bdr)!important;}
.chart-tip{background:var(--s2)!important;border:1px solid var(--s3)!important;border-radius:10px!important;padding:12px 16px!important;}
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
    saving: { l: "儲存中", c: "var(--orange)" },
    error: { l: "失敗", c: "var(--dn)" },
  };
  const v = m[status] || m.idle;
  const t = last
    ? new Date(last).toLocaleString("zh-TW", { hour12: false })
    : "—";
  return (
    <div
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
      <span style={{ opacity: 0.6, fontSize: 10 }}>{t}</span>
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
  return (
    <span
      onClick={onClick}
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

const Lbl = ({ children, tip }) => (
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
    <th onClick={() => onSort(sortKey)} style={{ ...th, textAlign: align }}>
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

/* ─── Chart Tooltip ──────────────────────────────────────────── */
const ChartTip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div
      style={{
        background: "var(--s1)",
        border: "1px solid var(--s3)",
        borderRadius: 10,
        padding: "10px 14px",
        boxShadow: "0 4px 20px rgba(0,0,0,0.12)",
      }}
    >
      <div
        style={{
          fontSize: 10,
          fontWeight: 700,
          color: "var(--t3)",
          marginBottom: 6,
          fontFamily: mono,
        }}
      >
        {label}
      </div>
      {payload.map((e, i) => (
        <div
          key={i}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 10,
            padding: "2px 0",
          }}
        >
          <div
            style={{
              width: 8,
              height: 8,
              borderRadius: 2,
              background: e.color,
            }}
          />
          <span style={{ fontSize: 11, color: "var(--t2)", fontWeight: 600 }}>
            {e.name}
          </span>
          <span
            style={{
              fontSize: 12,
              fontWeight: 800,
              color: "var(--t1)",
              fontFamily: mono,
              marginLeft: "auto",
            }}
          >
            {typeof e.value === "number" && Math.abs(e.value) >= 1
              ? fmt$(e.value)
              : fmtP(e.value)}
          </span>
        </div>
      ))}
    </div>
  );
};

/* ─── Overview Dashboard ────────────────────────────────────── */
function OverviewDashboard({
  slData,
  spData,
  slOrders,
  spOrders,
  slCosts,
  spCosts,
  theme,
  onNavigate,
  sY,
  sM,
}) {
  const slD = slData?.summary;
  const spS = spData?.s;
  const mono = "'JetBrains Mono',monospace";
  const fmt$ = (v) =>
    new Intl.NumberFormat("zh-TW", {
      style: "currency",
      currency: "TWD",
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    }).format(Number(v || 0));
  const fmtP = (v) => `${((Number(v) || 0) * 100).toFixed(2)}%`;
  const isDark = theme === "dark";
  const greenC = isDark ? "#2ECC71" : "#1A6B3C";
  const spC = isDark ? "#FF6533" : "#EE4D2D";
  const gridC = isDark ? "rgba(255,255,255,0.05)" : "rgba(0,0,0,0.05)";

  const hasAny = slD || spS;
  const totalRev = (slD?.rev || 0) + (spS?.tG || 0);
  const totalNet = (slD?.net || 0) + (spS?.afterComm || 0);
  const totalNetMargin = totalRev > 0 ? totalNet / totalRev : 0;
  const slRevShare = totalRev > 0 ? (slD?.rev || 0) / totalRev : 0;
  const spRevShare = totalRev > 0 ? (spS?.tG || 0) / totalRev : 0;

  const prevTotalRev =
    (slData?.prevMonth?.rev || 0) + (spData?.prevMonth?.rev || 0);
  const prevTotalNet =
    (slData?.prevMonth?.net || 0) + (spData?.prevMonth?.net || 0);
  const prevTotalMargin = prevTotalRev > 0 ? prevTotalNet / prevTotalRev : null;
  const hasPrev = prevTotalRev > 0;

  const periodLabel =
    sY === "All" ? "歷年" : sM === "All" ? `${sY}年` : `${sY}/${sM}`;

  const alerts = useMemo(() => {
    const list = [];
    if (slD && slD.trueNetMargin < slD.targetNetRate)
      list.push({
        level: "warn",
        platform: "官網",
        msg: `淨利率 ${fmtP(slD.trueNetMargin)} 低於目標 ${fmtP(
          slD.targetNetRate
        )}，差距 ${Math.abs(slD.gapVal).toFixed(1)}%`,
      });
    if (spS && spS.netMargin < spS.targetNet)
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
    return list;
  }, [slD, spS, slData, spData, slCosts, spCosts]);

  const trendData = useMemo(() => {
    const byMonth = {};
    Object.values(slOrders || {}).forEach((o) => {
      const cx =
        (o.status || "").includes("取消") || (o.status || "").includes("刪除");
      if (cx) return;
      if (sY && sY !== "All" && !String(o.date).startsWith(sY)) return;
      if (sM && sM !== "All" && String(o.date).substring(5, 7) !== sM) return;
      const ym = String(o.date || "").substring(0, 7);
      if (!ym || ym.length < 7) return;
      if (!byMonth[ym]) byMonth[ym] = { month: ym, slRev: 0, spRev: 0 };
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
      if (sY && sY !== "All" && !String(o.date).startsWith(sY)) return;
      if (sM && sM !== "All" && String(o.date).substring(5, 7) !== sM) return;
      const ym = String(o.date || "").substring(0, 7);
      if (!ym || ym.length < 7) return;
      if (!byMonth[ym]) byMonth[ym] = { month: ym, slRev: 0, spRev: 0 };
      const gross = o.grossPrice || 0;
      byMonth[ym].spRev += gross;
    });
    return Object.values(byMonth)
      .sort((a, b) => a.month.localeCompare(b.month))
      .slice(-12)
      .map((d) => ({
        ...d,
        label: d.month.substring(2).replace("-", "/"),
        total: d.slRev + d.spRev,
      }));
  }, [slOrders, spOrders, sY, sM]);

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
  }, [slData, spData, slCosts, spCosts]);

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
          請先上傳官網或蝦皮報表
        </div>
      </div>
    );
  }

  const overallStatus =
    totalNetMargin >= 0.15
      ? { label: "整體健康", c: "var(--up)" }
      : totalNetMargin >= 0.11
      ? { label: "需要關注", c: "var(--wn)" }
      : totalNetMargin > 0
      ? { label: "低於警戒", c: "var(--dn)" }
      : { label: "整體虧損", c: "var(--dn)" };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
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
            background: `linear-gradient(90deg, ${greenC}, ${spC})`,
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
                style={{
                  fontSize: 64,
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
                綜合淨利率
              </div>
              <div
                style={{
                  fontSize: 44,
                  fontWeight: 700,
                  fontFamily: mono,
                  lineHeight: 1,
                  color:
                    totalNetMargin >= 0.15
                      ? "var(--up)"
                      : totalNetMargin >= 0.11
                      ? "var(--wn)"
                      : totalNetMargin > 0
                      ? "var(--dn)"
                      : "var(--dn)",
                }}
              >
                {fmtP(totalNetMargin)}
              </div>
            </div>
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
              <div style={{ flex: 1, background: spC, opacity: 0.8 }} />
            </div>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  fontSize: 11,
                  fontWeight: 600,
                }}
              >
                <div
                  style={{
                    width: 8,
                    height: 8,
                    borderRadius: 2,
                    background: greenC,
                  }}
                />
                <span style={{ color: "var(--t2)" }}>官網</span>
                <span style={{ fontFamily: mono, color: greenC }}>
                  {(slRevShare * 100).toFixed(1)}%
                </span>
                <span style={{ color: "var(--t4)" }}>
                  {fmt$(slD?.rev || 0)}
                </span>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  fontSize: 11,
                  fontWeight: 600,
                }}
              >
                <span style={{ color: "var(--t4)" }}>{fmt$(spS?.tG || 0)}</span>
                <span style={{ fontFamily: mono, color: spC }}>
                  {(spRevShare * 100).toFixed(1)}%
                </span>
                <span style={{ color: "var(--t2)" }}>蝦皮</span>
                <div
                  style={{
                    width: 8,
                    height: 8,
                    borderRadius: 2,
                    background: spC,
                  }}
                />
              </div>
            </div>
          </div>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1px 1fr",
              borderTop: "1px solid var(--s3)",
              paddingTop: 20,
            }}
          >
            {[
              {
                label: "官網",
                color: greenC,
                rev: slD?.rev || 0,
                net: slD?.net || 0,
                margin: slD?.trueNetMargin || 0,
                target: slD?.targetNetRate || 0.15,
                orders: slD?.valid || 0,
                id: "shopline",
              },
              {
                label: "蝦皮",
                color: spC,
                rev: spS?.tG || 0,
                net: spS?.afterComm || 0,
                margin: spS?.netMargin || 0,
                target: spS?.targetNet || 0.14,
                orders: spS?.validN || 0,
                id: "shopee",
              },
            ].map((p, i) => (
              <React.Fragment key={p.id}>
                {i === 1 && (
                  <div style={{ background: "var(--s3)", margin: "0 20px" }} />
                )}
                <div style={{ padding: i === 0 ? "0 20px 0 0" : "0 0 0 20px" }}>
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
                        c: p.margin >= p.target ? p.color : "var(--wn)",
                        sub:
                          p.margin >= p.target
                            ? `超標 +${((p.margin - p.target) * 100).toFixed(
                                1
                              )}%`
                            : `差 ${((p.target - p.margin) * 100).toFixed(1)}%`,
                      },
                      { l: "營收", v: fmt$(p.rev) },
                      { l: "有效訂單", v: `${p.orders} 筆` },
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
        </div>
      </div>

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
            <TrendingUp size={14} color="var(--t3)" /> 月度營收趨勢
          </div>
          <div
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
                  const total = payload.reduce((s, e) => s + (e.value || 0), 0);
                  const fmt = (v) =>
                    new Intl.NumberFormat("zh-TW", {
                      style: "currency",
                      currency: "TWD",
                      minimumFractionDigits: 0,
                    }).format(v);
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
                      {payload.map((e, i) => {
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
                                width: 28,
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
                              {fmt(e.value)}
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
                          合計
                        </span>
                        <span
                          style={{
                            fontSize: 13,
                            fontWeight: 800,
                            fontFamily: mono,
                            color: "var(--t1)",
                          }}
                        >
                          {fmt(total)}
                        </span>
                      </div>
                    </div>
                  );
                }}
              />
              <Bar
                dataKey="slRev"
                name="官網"
                fill={greenC}
                opacity={0.85}
                radius={[3, 3, 0, 0]}
                maxBarSize={32}
                stackId="a"
              />
              <Bar
                dataKey="spRev"
                name="蝦皮"
                fill={spC}
                opacity={0.85}
                radius={[3, 3, 0, 0]}
                maxBarSize={32}
                stackId="a"
              />
            </ComposedChart>
          </ResponsiveContainer>
        </div>
      )}

      {/* ── AI 財務顧問 ── */}
      <AIAdvisor
        slData={slData}
        spData={spData}
        sY={sY}
        sM={sM}
        theme={theme}
      />

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
            <Package size={14} color="var(--t3)" /> 跨平台銷售排行
          </div>
          <div style={{ fontSize: 11, color: "var(--t3)", marginBottom: 16 }}>
            綠 = 官網　橘 = 蝦皮
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

/* ─── AI 財務顧問 ──────────────────────────────────────────── */
function AIAdvisor({ slData, spData, sY, sM, theme }) {
  const [apiKey, setApiKey] = React.useState(() => {
    try {
      return window.localStorage.getItem("bestea_anthropic_key") || "";
    } catch {
      return "";
    }
  });
  const [showKeyInput, setShowKeyInput] = React.useState(false);
  const [input, setInput] = React.useState("");
  const [messages, setMessages] = React.useState([]);
  const [loading, setLoading] = React.useState(false);
  const [keyDraft, setKeyDraft] = React.useState("");
  const bottomRef = React.useRef(null);
  const mono = "'JetBrains Mono',monospace";

  const isDark = theme === "dark";
  const accentC = isDark ? "#2ECC71" : "#1A6B3C";

  const fmt$ = (v) =>
    new Intl.NumberFormat("zh-TW", {
      style: "currency",
      currency: "TWD",
      minimumFractionDigits: 0,
    }).format(Number(v || 0));
  const fmtP = (v) => `${((Number(v) || 0) * 100).toFixed(2)}%`;

  const buildContext = () => {
    const slD = slData?.summary;
    const spS = spData?.s;
    const period =
      sY === "All" ? "歷年" : sM === "All" ? `${sY}年` : `${sY}/${sM}`;
    const lines = [
      `你是 BESTEA 天下第一好茶的財務顧問，以下是 ${period} 的跨平台經營數據：`,
      "",
      "【官網 Shopline】",
      slD
        ? [
            `  營收：${fmt$(slD.rev)}`,
            `  淨利：${fmt$(slD.net)}`,
            `  淨利率：${fmtP(slD.trueNetMargin)}（目標 ${fmtP(
              slD.targetNetRate
            )}）`,
            `  有效訂單：${slD.valid} 筆`,
            `  商品毛利率：${fmtP(slD.grossMargin)}`,
            `  通路總成本佔比：${fmtP(slD.realCommissionRate)}`,
            `  營收折讓率：${fmtP(slD.voucherRate)}`,
          ].join("\n")
        : "  無資料",
      "",
      "【蝦皮】",
      spS
        ? [
            `  營收：${fmt$(spS.tG)}`,
            `  淨利（扣分潤）：${fmt$(spS.afterComm)}`,
            `  淨利率：${fmtP(spS.netMargin)}（目標 ${fmtP(spS.targetNet)}）`,
            `  有效訂單：${spS.validN} 筆`,
            `  商品毛利率：${fmtP(
              spS.tG > 0 ? (spS.tG - spS.tC) / spS.tG : 0
            )}`,
            `  真實抽成率：${fmtP(spS.feeRate)}`,
            `  優惠券發放率：${fmtP(spS.voucherRate)}`,
            spS.comm > 0 ? `  分潤費用：${fmt$(spS.comm)}` : "",
          ]
            .filter(Boolean)
            .join("\n")
        : "  無資料",
      "",
      "請根據以上數據，用繁體中文簡潔回答使用者的問題。分析要具體、直接，提供可執行的建議。",
    ];
    return lines.join("\n");
  };

  const saveKey = () => {
    const k = keyDraft.trim();
    if (!k.startsWith("sk-ant-")) {
      alert("請輸入正確的 Anthropic API Key（以 sk-ant- 開頭）");
      return;
    }
    try {
      window.localStorage.setItem("bestea_anthropic_key", k);
    } catch {}
    setApiKey(k);
    setShowKeyInput(false);
    setKeyDraft("");
  };

  const send = async () => {
    if (!input.trim() || loading) return;
    if (!apiKey) {
      setShowKeyInput(true);
      return;
    }
    const userMsg = { role: "user", content: input.trim() };
    const newMsgs = [...messages, userMsg];
    setMessages(newMsgs);
    setInput("");
    setLoading(true);
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-5",
          max_tokens: 1024,
          system: buildContext(),
          messages: newMsgs,
        }),
      });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error?.message || "API 錯誤");
      }
      const data = await res.json();
      const reply = data.content?.[0]?.text || "（無回應）";
      setMessages((p) => [...p, { role: "assistant", content: reply }]);
    } catch (e) {
      setMessages((p) => [
        ...p,
        { role: "assistant", content: `⚠️ 錯誤：${e.message}` },
      ]);
    } finally {
      setLoading(false);
    }
  };

  React.useEffect(() => {
    if (bottomRef.current) {
      const container = bottomRef.current.closest("[data-chat-scroll]");
      if (container) container.scrollTop = container.scrollHeight;
    }
  }, [messages, loading]);

  const quickPrompts = [
    "這期整體表現如何？有哪些警訊？",
    "哪個平台利潤率更好？原因是什麼？",
    "下個月我應該重點優化什麼？",
    "通路費用是否過高？如何改善？",
  ];

  return (
    <div
      className="f5"
      style={{
        background: "var(--s1)",
        border: "1px solid var(--s3)",
        borderRadius: 16,
        overflow: "hidden",
      }}
    >
      {/* Header */}
      <div
        style={{
          padding: "16px 24px",
          borderBottom: "1px solid var(--s3)",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div
            style={{
              width: 28,
              height: 28,
              borderRadius: 8,
              background: "var(--s2)",
              border: "1px solid var(--s3)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: 14,
            }}
          >
            ✦
          </div>
          <div>
            <div style={{ fontSize: 12, fontWeight: 700, color: "var(--t1)" }}>
              AI 財務顧問
            </div>
            <div style={{ fontSize: 10, color: "var(--t3)" }}>
              基於本期數據 · Claude Sonnet
            </div>
          </div>
        </div>
        <button
          onClick={() => setShowKeyInput(!showKeyInput)}
          style={{
            fontSize: 10,
            fontWeight: 700,
            color: apiKey ? "var(--up)" : "var(--wn)",
            background: apiKey ? "var(--up-dim)" : "var(--wn-dim)",
            border: `1px solid ${apiKey ? "var(--up-bdr)" : "var(--wn-bdr)"}`,
            borderRadius: 6,
            padding: "4px 12px",
            cursor: "pointer",
          }}
        >
          {apiKey ? "✓ 已設定 API Key" : "⚠ 設定 API Key"}
        </button>
      </div>

      {/* API Key input */}
      {showKeyInput && (
        <div
          style={{
            padding: "14px 24px",
            background: "var(--s2)",
            borderBottom: "1px solid var(--s3)",
            display: "flex",
            gap: 8,
            alignItems: "center",
            flexWrap: "wrap",
          }}
        >
          <input
            type="password"
            placeholder="sk-ant-api03-..."
            value={keyDraft}
            onChange={(e) => setKeyDraft(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && saveKey()}
            style={{
              flex: 1,
              minWidth: 200,
              border: "1px solid var(--s3)",
              borderRadius: 8,
              padding: "8px 12px",
              fontSize: 12,
              fontFamily: mono,
              background: "var(--s1)",
              color: "var(--t1)",
              outline: "none",
            }}
          />
          <button
            onClick={saveKey}
            style={{
              background: accentC,
              color: "#fff",
              border: "none",
              borderRadius: 8,
              padding: "8px 16px",
              fontSize: 11,
              fontWeight: 700,
              cursor: "pointer",
            }}
          >
            儲存
          </button>
          {apiKey && (
            <button
              onClick={() => {
                setApiKey("");
                try {
                  window.localStorage.removeItem("bestea_anthropic_key");
                } catch {}
                setShowKeyInput(false);
              }}
              style={{
                background: "var(--dn-dim)",
                color: "var(--dn)",
                border: "1px solid var(--dn-bdr)",
                borderRadius: 8,
                padding: "8px 12px",
                fontSize: 11,
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              清除
            </button>
          )}
          <div style={{ fontSize: 10, color: "var(--t3)", width: "100%" }}>
            Key 僅存在瀏覽器 localStorage，不會上傳任何伺服器
          </div>
        </div>
      )}

      {/* Messages */}
      <div
        style={{
          height: 280,
          overflowY: "auto",
          padding: "16px 24px",
          display: "flex",
          flexDirection: "column",
          gap: 12,
        }}
      >
        {messages.length === 0 && (
          <div
            style={{
              flex: 1,
              display: "flex",
              flexDirection: "column",
              justifyContent: "center",
              gap: 12,
            }}
          >
            <div
              style={{
                fontSize: 12,
                color: "var(--t3)",
                textAlign: "center",
                marginBottom: 8,
              }}
            >
              快速提問
            </div>
            <div
              style={{
                display: "flex",
                flexWrap: "wrap",
                gap: 8,
                justifyContent: "center",
              }}
            >
              {quickPrompts.map((q, i) => (
                <button
                  key={i}
                  onClick={() => {
                    setInput(q);
                  }}
                  style={{
                    fontSize: 11,
                    color: "var(--t2)",
                    background: "var(--s2)",
                    border: "1px solid var(--s3)",
                    borderRadius: 20,
                    padding: "6px 14px",
                    cursor: "pointer",
                    transition: "all .15s",
                  }}
                >
                  {q}
                </button>
              ))}
            </div>
          </div>
        )}
        {messages.map((m, i) => (
          <div
            key={i}
            style={{
              display: "flex",
              justifyContent: m.role === "user" ? "flex-end" : "flex-start",
            }}
          >
            <div
              style={{
                maxWidth: "82%",
                padding: "10px 14px",
                borderRadius:
                  m.role === "user"
                    ? "12px 12px 2px 12px"
                    : "12px 12px 12px 2px",
                background: m.role === "user" ? accentC : "var(--s2)",
                color: m.role === "user" ? "#fff" : "var(--t1)",
                fontSize: 12,
                fontWeight: 500,
                lineHeight: 1.7,
                whiteSpace: "pre-wrap",
              }}
            >
              {m.content}
            </div>
          </div>
        ))}
        {loading && (
          <div style={{ display: "flex", justifyContent: "flex-start" }}>
            <div
              style={{
                padding: "10px 16px",
                borderRadius: "12px 12px 12px 2px",
                background: "var(--s2)",
                display: "flex",
                gap: 4,
                alignItems: "center",
              }}
            >
              {[0, 1, 2].map((i) => (
                <div
                  key={i}
                  style={{
                    width: 6,
                    height: 6,
                    borderRadius: "50%",
                    background: accentC,
                  }}
                />
              ))}
            </div>
          </div>
        )}
        <div ref={bottomRef} />
      </div>

      {/* Input */}
      <div
        style={{
          padding: "12px 24px",
          borderTop: "1px solid var(--s3)",
          display: "flex",
          gap: 8,
        }}
      >
        <input
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && !e.shiftKey && send()}
          placeholder={
            apiKey
              ? "輸入問題，例如：這個月為什麼利潤下降？"
              : "請先設定 API Key"
          }
          disabled={!apiKey || loading}
          style={{
            flex: 1,
            border: "1px solid var(--s3)",
            borderRadius: 10,
            padding: "10px 14px",
            fontSize: 12,
            background: "var(--s2)",
            color: "var(--t1)",
            outline: "none",
            fontFamily: "'Inter','Noto Sans TC',sans-serif",
          }}
        />
        <button
          onClick={send}
          disabled={!input.trim() || loading || !apiKey}
          style={{
            background: accentC,
            color: "#fff",
            border: "none",
            borderRadius: 10,
            padding: "10px 18px",
            fontSize: 12,
            fontWeight: 700,
            cursor: "pointer",
            opacity: !input.trim() || loading || !apiKey ? 0.4 : 1,
            transition: "opacity .15s",
          }}
        >
          送出
        </button>
        {messages.length > 0 && (
          <button
            onClick={() => setMessages([])}
            style={{
              background: "var(--s2)",
              color: "var(--t3)",
              border: "1px solid var(--s3)",
              borderRadius: 10,
              padding: "10px 12px",
              fontSize: 11,
              cursor: "pointer",
            }}
          >
            清除
          </button>
        )}
      </div>
    </div>
  );
}

/* ─── Commission Panel ──────────────────────────────────────── */
function CommissionPanel({ commissions, onUpdate, selYear, selMonth }) {
  const key = commKey(selYear, selMonth);
  const isAggregated = selMonth === "All" || selYear === "All";
  const aggregatedVal = useMemo(() => {
    if (selYear === "All") {
      return Object.values(commissions).reduce(
        (s, v) => s + (v !== "" && v !== undefined ? Number(v) || 0 : 0),
        0
      );
    }
    if (selMonth === "All") {
      return Object.entries(commissions).reduce((s, [k, v]) => {
        if (k.startsWith(selYear + "-") && v !== "" && v !== undefined)
          s += Number(v) || 0;
        return s;
      }, 0);
    }
    return null;
  }, [commissions, selYear, selMonth]);

  const [local, setLocal] = useState(String(commissions[key] ?? ""));
  const [focused, setFocused] = useState(false);
  useEffect(() => {
    setLocal(String(commissions[key] ?? ""));
  }, [key, commissions]);
  const handleBlur = () => {
    setFocused(false);
    const n = parseFloat(local);
    onUpdate(key, isNaN(n) ? "" : n);
  };
  const hasVal =
    commissions[key] !== undefined &&
    commissions[key] !== "" &&
    Number(commissions[key]) > 0;
  const label =
    selYear === "All"
      ? "歷年"
      : selMonth === "All"
      ? `${selYear}`
      : `${selYear}/${selMonth}`;
  const fmt$ = (v) =>
    new Intl.NumberFormat("zh-TW", {
      style: "currency",
      currency: "TWD",
      minimumFractionDigits: 0,
    }).format(Number(v || 0));

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
        <Users size={12} color="var(--purple)" /> 分潤費用
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
            color: "var(--purple)",
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
              color: aggregatedVal > 0 ? "var(--purple)" : "var(--t3)",
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
            onChange={(e) => setLocal(e.target.value)}
            onFocus={() => setFocused(true)}
            onBlur={handleBlur}
            style={{
              ...inp,
              width: "100%",
              textAlign: "right",
              paddingLeft: 36,
              borderColor: focused ? "var(--purple)" : "var(--s3)",
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
          ? "各月分潤加總，切換到特定月份可編輯"
          : "此期間分潤費用將從最終淨利扣除"}
      </p>
    </div>
  );
}

/* ─── Main App ───────────────────────────────────────────────── */
export default function App() {
  const [theme, setTheme] = useState(() => gl(SK.theme, "light"));
  const [platform, setPlatform] = useState(() => gl(SK.platform, "overview"));
  const [slFp, setSlFp] = useState(() => gl(SK.slFp, DEFAULT_FP_SL));
  const [spFp, setSpFp] = useState(() => gl(SK.spFp, DEFAULT_FP_SP));
  const [slCosts, setSlCosts] = useState(() => gl(SK.slCosts, {}));
  const [spCosts, setSpCosts] = useState(() => gl(SK.spCosts, {}));
  const [slOrders, setSlOrders] = useState(() => gl(SK.slOrders, {}));
  const [spOrders, setSpOrders] = useState(() => gl(SK.spOrders, {}));
  const [commissions, setCommissions] = useState(() => gl(SK.commissions, {}));

  const [notif, setNotif] = useState(null);
  const [toasts, setToasts] = useState([]);
  const toastIdRef = useRef(0);
  const [sY, setSY] = useState("All");
  const [sM, setSM] = useState("All");
  const [search, setSearch] = useState("");
  const [mSearch, setMSearch] = useState("");
  const [lossOnly, setLossOnly] = useState(false);
  const [sync, setSync] = useState("connecting");
  const [cReady, setCReady] = useState(false);
  const [aReady, setAReady] = useState(false);
  const [costSort, setCostSort] = useState({ key: "soldQty", dir: "desc" });
  const [orderSort, setOrderSort] = useState({ key: "date", dir: "desc" });
  const [page, setPage] = useState(0);
  const [pageSize, setPageSize] = useState(30);

  const fRef = useRef({});
  const cRef = useRef(null);
  const cDoc = useRef(null);
  const firstMissRef = useRef(null);
  const prevSlMonthlyHashes = useRef({});
  const prevSpMonthlyHashes = useRef({});
  const migrating = useRef(false);
  const applying = useRef(false);
  const sTimer = useRef(null);
  const lR = useRef(0);
  const lL = useRef(0);
  const meta = useRef({
    clientId: typeof window !== "undefined" ? gcid() : "",
  });
  const [lastSyncAt, setLastSyncAt] = useState(0);

  useEffect(() => {
    sl_s(SK.theme, theme);
  }, [theme]);
  useEffect(() => {
    sl_s(SK.platform, platform);
  }, [platform]);
  useEffect(() => {
    sl_s(SK.slFp, slFp);
  }, [slFp]);
  useEffect(() => {
    sl_s(SK.spFp, spFp);
  }, [spFp]);
  useEffect(() => {
    sl_s(SK.slCosts, slCosts);
  }, [slCosts]);
  useEffect(() => {
    sl_s(SK.spCosts, spCosts);
  }, [spCosts]);
  useEffect(() => {
    sl_s(SK.slOrders, slOrders);
  }, [slOrders]);
  useEffect(() => {
    sl_s(SK.spOrders, spOrders);
  }, [spOrders]);
  useEffect(() => {
    sl_s(SK.commissions, commissions);
  }, [commissions]);

  const toast = useCallback((msg, opts = {}) => {
    const id = ++toastIdRef.current;
    const { type = "info", duration = 3500, action, actionLabel } = opts;
    setToasts((p) => [
      ...p,
      { id, msg, type, duration, action, actionLabel, removing: false },
    ]);
    if (!action) {
      setTimeout(() => {
        setToasts((p) =>
          p.map((t) => (t.id === id ? { ...t, removing: true } : t))
        );
        setTimeout(() => setToasts((p) => p.filter((t) => t.id !== id)), 350);
      }, duration);
    }
  }, []);
  const removeToast = useCallback((id) => {
    setToasts((p) =>
      p.map((t) => (t.id === id ? { ...t, removing: true } : t))
    );
    setTimeout(() => setToasts((p) => p.filter((t) => t.id !== id)), 350);
  }, []);

  const msg = (m) => {
    setNotif(m);
    clearTimeout(msg._t);
    msg._t = setTimeout(() => setNotif(null), 3500);
  };

  /* Firebase init */
  useEffect(() => {
    try {
      const app = getApps().length ? getApp() : initializeApp(FBC);
      const auth = getAuth(app),
        db = getFirestore(app);
      cDoc.current = doc(db, FSPC.collection, FSPC.docId);
      fRef.current._db = db;
      fRef.current._slDoc = doc(db, FSPC_SL.collection, FSPC_SL.docId);
      fRef.current._spDoc = doc(db, FSPC_SP.collection, FSPC_SP.docId);
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

    // meta 監聽
    const unMeta = onSnapshot(
      cDoc.current,
      async (snap) => {
        try {
          const metaData = parseMeta(snap);
          const rMs = Number(metaData?.updatedAtMs || 0);
          if (metaData && rMs > lR.current && rMs > lL.current) {
            applying.current = true;
            if (metaData.slFp) setSlFp(metaData.slFp);
            if (metaData.spFp) setSpFp(metaData.spFp);
            if (metaData.slCosts) setSlCosts(metaData.slCosts);
            if (metaData.spCosts) setSpCosts(metaData.spCosts);
            if (metaData.commissions) setCommissions(metaData.commissions);
            lR.current = rMs;
            lL.current = rMs;
            setLastSyncAt(Date.now());
            setTimeout(() => {
              applying.current = false;
            }, 50);
          }
          await runMigrationIfNeeded(snap);
          setCReady(true);
          setSync("synced");
        } catch (e) {
          console.error("[Meta Snapshot Error]", e);
          setCReady(true);
          setSync("error");
        }
      },
      (err) => {
        console.error("[Meta Snapshot Error]", err);
        setCReady(true);
        setSync("error");
      }
    );

    // 官網月份 collection 監聽
    let slFirstLoad = true;
    const unSl = onSnapshot(
      fRef.current._slMonthlyColl,
      (snapshot) => {
        try {
          const all = {};
          let maxMs = 0;
          snapshot.forEach((docSnap) => {
            const d = docSnap.data();
            if (d?.ordersJson) {
              try {
                Object.assign(all, JSON.parse(d.ordersJson));
              } catch {}
            }
            const m = Number(d?.updatedAtMs || 0);
            if (m > maxMs) maxMs = m;
            prevSlMonthlyHashes.current[docSnap.id] = d?.ordersJson || "";
          });
          if ((slFirstLoad || maxMs > lR.current) && maxMs > lL.current && !applying.current) {
            slFirstLoad = false;
        applying.current = true;
            setSlOrders(all);
            if (maxMs > lR.current) lR.current = maxMs;
            setLastSyncAt(Date.now());
            setTimeout(() => {
              applying.current = false;
            }, 50);
          }
        } catch (e) {
          console.error("[SL Monthly Snapshot Error]", e);
        }
      },
      (err) => console.error("[SL Monthly Snapshot Error]", err)
    );

    // 蝦皮月份 collection 監聽
    let spFirstLoad = true;
    const unSp = onSnapshot(
      fRef.current._spMonthlyColl,
      (snapshot) => {
        try {
          const all = {};
          let maxMs = 0;
          snapshot.forEach((docSnap) => {
            const d = docSnap.data();
            if (d?.ordersJson) {
              try {
                Object.assign(all, JSON.parse(d.ordersJson));
              } catch {}
            }
            const m = Number(d?.updatedAtMs || 0);
            if (m > maxMs) maxMs = m;
            prevSpMonthlyHashes.current[docSnap.id] = d?.ordersJson || "";
          });
          if ((spFirstLoad || maxMs > lR.current) && maxMs > lL.current && !applying.current) {
            spFirstLoad = false;
        applying.current = true;
            setSpOrders(all);
            if (maxMs > lR.current) lR.current = maxMs;
            setLastSyncAt(Date.now());
            setTimeout(() => {
              applying.current = false;
            }, 50);
          }
        } catch (e) {
          console.error("[SP Monthly Snapshot Error]", e);
        }
      },
      (err) => console.error("[SP Monthly Snapshot Error]", err)
    );

    return () => {
      unMeta();
      unSl();
      unSp();
    };
  }, [aReady]);

  /* Firebase save (meta + only changed months) */
  useEffect(() => {
    if (!aReady || !cReady || !cDoc.current || applying.current) return;
    if (migrating.current) return;
    clearTimeout(sTimer.current);
    sTimer.current = setTimeout(async () => {
      try {
        setSync("saving");
        const ms = Date.now();
        const db = fRef.current._db;

        const metaPl = deepClean({
          slFp,
          spFp,
          slCosts,
          spCosts,
          commissions,
          updatedAtMs: ms,
          updatedBy: meta.current.clientId,
        });

        const slByMonth = groupOrdersByMonth(slOrders);
        const spByMonth = groupOrdersByMonth(spOrders);

        const writes = [
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
          ),
        ];

        // 官網：只寫有變動的月份
        Object.entries(slByMonth).forEach(([ym, orders]) => {
          const json = JSON.stringify(orders);
          if (prevSlMonthlyHashes.current[ym] === json) return;
          prevSlMonthlyHashes.current[ym] = json;
          writes.push(
            setDoc(doc(db, SL_MONTHLY_COLL, ym), {
              ordersJson: json,
              count: Object.keys(orders).length,
              updatedAtMs: ms,
            })
          );
        });
        // 官網：偵測已刪除的月份
        Object.keys(prevSlMonthlyHashes.current).forEach((ym) => {
          if (!slByMonth[ym]) {
            delete prevSlMonthlyHashes.current[ym];
            writes.push(deleteDoc(doc(db, SL_MONTHLY_COLL, ym)));
          }
        });

        // 蝦皮：只寫有變動的月份
        Object.entries(spByMonth).forEach(([ym, orders]) => {
          const json = JSON.stringify(orders);
          if (prevSpMonthlyHashes.current[ym] === json) return;
          prevSpMonthlyHashes.current[ym] = json;
          writes.push(
            setDoc(doc(db, SP_MONTHLY_COLL, ym), {
              ordersJson: json,
              count: Object.keys(orders).length,
              updatedAtMs: ms,
            })
          );
        });
        Object.keys(prevSpMonthlyHashes.current).forEach((ym) => {
          if (!spByMonth[ym]) {
            delete prevSpMonthlyHashes.current[ym];
            writes.push(deleteDoc(doc(db, SP_MONTHLY_COLL, ym)));
          }
        });

        await Promise.all(writes);
        lL.current = ms;
        setLastSyncAt(Date.now());
        setSync("synced");
      } catch (e) {
        console.error("[Save Error]", e);
        setSync("error");
      }
    }, 900);
    return () => clearTimeout(sTimer.current);
  }, [
    slFp,
    spFp,
    slCosts,
    spCosts,
    slOrders,
    spOrders,
    commissions,
    aReady,
    cReady,
  ]);

  /* ─── Shopline CSV/XLSX Parser ─────────────────────────────── */
  const processSLParsed = (parsed) => {
    if (!Array.isArray(parsed) || parsed.length < 2) {
      toast("格式錯誤", { type: "error" });
      return;
    }
    const hdrs = parsed[0].map((h) => safeText(h).replace(/^\uFEFF/, ""));
    const idx = (n) => hdrs.indexOf(n);
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

    const newOrders = {};
    let count = 0;

    for (let i = 1; i < parsed.length; i++) {
      const row = parsed[i];
      if (!row || row.length < 5) continue;

      const rawOrderId = safeText(row[im.orderId]);
      if (!rawOrderId) continue;

      const cartId = im.cartId > -1 ? safeText(row[im.cartId]) : rawOrderId;
      const groupKey = cartId || rawOrderId;

      const rawDate = safeText(row[im.date]);
      const date = rawDate
        ? rawDate.replace(/\//g, "-").split(" ")[0].split("T")[0]
        : "1970-01-01";

      if (!newOrders[groupKey]) {
        const revenue = numOrZero(
          im.paidTotal > -1
            ? row[im.paidTotal]
            : im.total > -1
            ? row[im.total]
            : 0
        );
        const voucherAmt =
          numOrZero(im.discount > -1 ? row[im.discount] : 0) +
          numOrZero(im.creditOffset > -1 ? row[im.creditOffset] : 0) +
          numOrZero(im.pointOffset > -1 ? row[im.pointOffset] : 0);

        const statusRaw = im.status > -1 ? safeText(row[im.status]) : "";
        const isTaxExempt =
          (im.taxExempt > -1 && safeText(row[im.taxExempt]) === "免稅") ||
          (im.invoiceStatus > -1 &&
            safeText(row[im.invoiceStatus]) === "待開立");
        const hasReturn = im.returnId > -1 && safeText(row[im.returnId]) !== "";
        const shippingIncome = numOrZero(
          im.shippingFee > -1 ? row[im.shippingFee] : 0
        );
        const payMethod = im.payMethod > -1 ? safeText(row[im.payMethod]) : "";
        const delivMethod = im.delivery > -1 ? safeText(row[im.delivery]) : "";

        newOrders[groupKey] = {
          orderId: rawOrderId,
          date,
          status: statusRaw,
          revenue,
          voucherAmount: voucherAmt,
          shippingIncome,
          paymentMethod: payMethod,
          deliveryMethod: delivMethod,
          isTaxExempt,
          hasReturn,
          items: [],
        };
        count++;
      }

      const prodName =
        im.prodName > -1 ? safeText(row[im.prodName]) : "未知商品";
      const option = im.option > -1 ? safeText(row[im.option]) : "";
      const prodId = im.prodId > -1 ? safeText(row[im.prodId]) : "";
      const qty = parseInt(im.qty > -1 ? row[im.qty] || 1 : 1, 10) || 1;
      const price = numOrZero(im.unitPrice > -1 ? row[im.unitPrice] : 0);
      const prodType = im.prodType > -1 ? safeText(row[im.prodType]) : "商品";
      const addOnType = im.addOnType > -1 ? safeText(row[im.addOnType]) : "";
      const isGift = prodType === "贈品";

      if (!prodName) continue;

      const costKey = `${prodName}_${option}`.trim();

      newOrders[groupKey].items.push({
        key: costKey,
        name: prodName,
        option,
        qty,
        price,
        isGift,
        isAddOn: prodType === "加購品",
        addOnType,
      });
    }

    setSlOrders((p) => {
      const merged = { ...p };
      Object.entries(newOrders).forEach(([gk, order]) => {
        merged[order.orderId] = order;
      });
      const dates = Object.values(merged)
        .map((o) => String(o.date))
        .filter(Boolean)
        .sort()
        .reverse();
      if (dates.length) {
        const ly = dates[0].substring(0, 4);
        const lm = dates[0].substring(5, 7);
        setSY(ly);
        setSM(lm);
      }
      toast(`已匯入 ${count} 筆官網訂單`, { type: "success" });
      return merged;
    });
  };

  /* ─── Shopee CSV/XLSX Parser ────────────────────────────────── */
  const processSPParsed = (parsed) => {
    if (!Array.isArray(parsed) || parsed.length < 2) {
      toast("格式錯誤", { type: "error" });
      return;
    }
    const hdrs = parsed[0].map((h) => safeText(h).replace(/^\uFEFF/, ""));
    const idx = (n) => hdrs.indexOf(n);
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

    const newOrders = {};
    let count = 0;

    for (let i = 1; i < parsed.length; i++) {
      const row = parsed[i];
      if (!row || row.length < 5) continue;
      const orderId = safeText(row[im.orderId]);
      if (!orderId) continue;
      const date = (
        safeText(row[im.date]).split(" ")[0] || "1970-01-01"
      ).replace(/\//g, "-");

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
          platformShippingFee: 0,
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
      const m = { ...p, ...newOrders };
      const dates = Object.values(m)
        .map((o) => String(o.date))
        .filter(Boolean)
        .sort()
        .reverse();
      if (dates.length) {
        const ly = dates[0].substring(0, 4);
        const lm = dates[0].substring(5, 7);
        setSY(ly);
        setSM(lm);
      }
      toast(`已匯入 ${count} 筆蝦皮訂單`, { type: "success" });
      return m;
    });
  };

  const handleFile = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    const isX = f.name.endsWith(".xlsx") || f.name.endsWith(".xls");
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
        if (platform === "shopline") processSLParsed(rows);
        else processSPParsed(rows);
      } else {
        if (platform === "shopline") processSLParsed(parseCSV(d2));
        else processSPParsed(parseCSV(d2));
      }
    };
    rd.onload = (ev) => exec(ev.target.result, isX);
    if (isX) {
      if (typeof window.XLSX === "undefined") {
        const s2 = document.createElement("script");
        s2.src =
          "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
        s2.onload = () => rd.readAsArrayBuffer(f);
        document.head.appendChild(s2);
      } else rd.readAsArrayBuffer(f);
    } else rd.readAsText(f);
    e.target.value = "";
  };

  /* ─── Shopline Data Processing ────────────────────────────── */
  const slData = useMemo(() => {
    const all = Object.values(slOrders);
    if (!all.length) return null;
    const years = [...new Set(all.map((o) => o.date.substring(0, 4)))]
      .sort()
      .reverse();
    const months = [
      ...new Set(
        all
          .filter((o) => sY === "All" || o.date.startsWith(sY))
          .map((o) => o.date.substring(5, 7))
      ),
    ].sort();
    const tnr = (parseFloat(slFp.targetNet) || 17) / 100;
    const mm = {};
    Object.keys(slCosts).forEach((k) => {
      const p = k.split("_");
      mm[k] = {
        key: k,
        name: p[0],
        option: p[1]?.trim() || "標準規格",
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
    };
    const fl = all.filter((o) => {
      const oy = o.date.substring(0, 4),
        om = o.date.substring(5, 7);
      if (sY !== "All" && oy !== sY) return false;
      if (sM !== "All" && om !== sM) return false;
      const cx = o.status.includes("取消") || o.status.includes("刪除");
      t.rawTotal += o.revenue;
      if (cx) {
        t.cancelledTotal += o.revenue;
        return false;
      }
      return true;
    });
    const ol = fl
      .map((order) => {
        const ofp = order.snapshotFeeParams;
        const sfr =
          ((ofp?.platformFeeRate != null
            ? Number(ofp.platformFeeRate)
            : parseFloat(slFp.platformFeeRate)) || 0) / 100;
        const oer =
          ((ofp?.opExpense != null
            ? Number(ofp.opExpense)
            : parseFloat(slFp.opExpense)) || 0) / 100;
        const dtr =
          ((ofp?.tax != null ? Number(ofp.tax) : parseFloat(slFp.tax)) || 0) /
          100;
        let pr = { rate: 0.022, flat: 0 };
        for (const [k, v] of Object.entries(SL_PAYMENT_RATES)) {
          if (order.paymentMethod.includes(k)) {
            pr = v;
            break;
          }
        }
        const pf = order.revenue * pr.rate + pr.flat;
        let sc2 = 0;
        if (SL_INTL_METHODS.some((k) => order.deliveryMethod.includes(k)))
          sc2 = order.shippingIncome;
        else {
          for (const [k, v] of Object.entries(SL_SHIPPING_RATES)) {
            if (order.deliveryMethod.includes(k)) {
              sc2 = v;
              break;
            }
          }
          if (sc2 === 0) sc2 = 120;
        }
        const plf = order.revenue * sfr;
        let oc = 0;
        order.items.forEach((item) => {
          const cv =
            Object.prototype.hasOwnProperty.call(item, "snapshotCost") &&
            item.snapshotCost !== null
              ? Number(item.snapshotCost) || 0
              : Number(slCosts[item.key]) || 0;
          oc += cv * item.qty;
          t.totalQty += item.qty;
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
          const ir = (Number(item.price) || 0) * item.qty,
            ic = cv * item.qty;
          mm[item.key].profitContribution += ir - ic;
          mm[item.key].totalRevenue += ir;
          mm[item.key].totalCost += ic;
          if (item.isGift === true || safeText(item.name).includes("贈品")) {
            t.giftCost += ic;
            t.giftQty += item.qty;
          }
        });
        const cm = order.revenue - oc - pf - sc2 - plf;
        const tax = order.isTaxExempt ? 0 : order.revenue * dtr;
        const opx = order.revenue * oer;
        const net = cm - opx - tax;
        t.rev += order.revenue;
        t.pFee += pf;
        t.sCost += sc2;
        t.platformFee += plf;
        t.cost += oc;
        t.contributionMargin += cm;
        t.net += net;
        t.inbound += order.revenue - pf - sc2 - plf;
        t.voucher += order.voucherAmount;
        t.opExpTotal += opx;
        t.taxTotal += tax;
        t.valid++;
        return {
          ...order,
          pFee: pf,
          sCost: sc2,
          net,
          oCost: oc,
          currentOrderContribution: cm,
        };
      })
      .sort((a, b) => b.date.localeCompare(a.date));
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
      },
    };
  }, [slOrders, sY, sM, slFp, slCosts]);

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
    const targetNet = (parseFloat(spFp.targetNet) || 14) / 100;
    const prods = {};
    let tG = 0,
      tV = 0,
      tF = 0,
      tC = 0,
      tOp = 0,
      tTx = 0,
      validN = 0,
      lossN = 0;
    const filtered = all.filter((o) => {
      if (sY !== "All" && !String(o.date).startsWith(sY)) return false;
      if (sM !== "All" && !String(o.date).startsWith(`${sY}-${sM}`))
        return false;
      return true;
    });
    const orderList = filtered
      .map((order) => {
        const st = safeText(order.status),
          rf = safeText(order.refundStatus);
        const isCanc = st.includes("不成立") || st.includes("取消");
        const isRef =
          !isCanc &&
          (rf !== "" || (st.includes("退貨") && !st.includes("已完成")));

        if (isCanc) return null;

        const gross = numOrZero(order.grossPrice);

        if (isRef) return null;

        const fee = numOrZero(order.exactOrderFee);
        const platformShipping = numOrZero(order.platformShippingFee);

        const ofp = order.snapshotFeeParams;
        const opEx =
          ofp?.opExpense != null
            ? Number(ofp.opExpense) || 0
            : parseFloat(spFp.opExpense) || 0;
        const tx =
          ofp?.tax != null ? Number(ofp.tax) || 0 : parseFloat(spFp.tax) || 0;
        let oCost = 0;
        (order.items || []).forEach((item) => {
          const ic =
            Object.prototype.hasOwnProperty.call(item, "snapshotCost") &&
            item.snapshotCost !== null
              ? Number(item.snapshotCost) || 0
              : Number(spCosts[item.key]) || 0;
          oCost += ic * (item.qty || 1);
          if (!prods[item.key])
            prods[item.key] = {
              key: item.key,
              name: item.name,
              option: item.option,
              soldQty: 0,
              estProfit: 0,
            };
          prods[item.key].soldQty += item.qty || 1;
          const ir = numOrZero(item.activityPrice) * (item.qty || 1);
          prods[item.key].estProfit +=
            ir - ic * (item.qty || 1) - ir * (opEx / 100) - ir * (tx / 100);
        });

        const net =
          gross - numOrZero(order.sellerVoucher) - fee - platformShipping;
        const gp = net - oCost;
        const opAmt = gross * (opEx / 100);
        const taxBase = numOrZero(order.buyerTotal) || gross;
        const txAmt = taxBase * (tx / 100);
        const finalNet = gp - opAmt - txAmt;
        tG += gross;
        tV += numOrZero(order.sellerVoucher);
        tF += fee + platformShipping;
        tC += oCost;
        tOp += opAmt;
        tTx += txAmt;
        validN++;
        if (finalNet < 0) lossN++;
        return {
          ...order,
          localGross: gross,
          totalOrderFee: fee + platformShipping,
          orderCost: oCost,
          netIncome: net,
          grossProfit: gp,
          finalNetProfit: finalNet,
          orderOpExpense: opAmt,
          orderTax: txAmt,
        };
      })
      .filter(Boolean);
    const ck = commKey(sY, sM);
    let comm = 0;
    if (commissions[ck] !== undefined && commissions[ck] !== "") {
      comm = Number(commissions[ck]);
    } else if (sM === "All" && sY !== "All") {
      Object.entries(commissions).forEach(([k, v]) => {
        if (k.startsWith(sY + "-") && v !== "" && v !== undefined)
          comm += Number(v) || 0;
      });
    } else if (sY === "All") {
      Object.values(commissions).forEach((v) => {
        if (v !== "" && v !== undefined) comm += Number(v) || 0;
      });
    }
    const tNetPro = tG - tV - tF - tC - tOp - tTx;
    const afterComm = tNetPro - comm;
    const netMargin = tG > 0 ? afterComm / tG : 0;

    let spPrevM = null;
    if (sM !== "All" && sY !== "All") {
      const pmNum2 = Number(sM) - 1;
      const pm2 = pmNum2 === 0 ? "12" : String(pmNum2).padStart(2, "0");
      const py2 = pmNum2 === 0 ? String(Number(sY) - 1) : sY;
      const prevKey = `${py2}-${pm2}`;
      const prevSPOrders = all.filter((o) =>
        String(o.date).startsWith(prevKey)
      );
      if (prevSPOrders.length > 0) {
        let ptG = 0,
          ptV = 0,
          ptF = 0,
          ptC = 0,
          ptOp = 0,
          ptTx = 0;
        prevSPOrders.forEach((order) => {
          const st2 = safeText(order.status),
            rf2 = safeText(order.refundStatus);
          if (st2.includes("不成立") || st2.includes("取消")) return;
          if (rf2 !== "" || (st2.includes("退貨") && !st2.includes("已完成")))
            return;
          const g2 = numOrZero(order.grossPrice);
          const fee2 = numOrZero(order.exactOrderFee);
          const platformShipping2 = numOrZero(order.platformShippingFee);
          const opEx2 = parseFloat(spFp.opExpense) || 0;
          const tx2 = parseFloat(spFp.tax) || 0;
          let oc3 = 0;
          (order.items || []).forEach((item) => {
            oc3 += (Number(spCosts[item.key]) || 0) * (item.qty || 1);
          });
          ptG += g2;
          ptV += numOrZero(order.sellerVoucher);
          ptF += fee2 + platformShipping2;
          ptC += oc3;
          ptOp += g2 * (opEx2 / 100);
          const taxBase2 = numOrZero(order.buyerTotal) || g2;
          ptTx += taxBase2 * (tx2 / 100);
        });
        const ptNet = ptG - ptV - ptF - ptC - ptOp - ptTx;
        spPrevM = {
          rev: ptG,
          net: ptNet,
          netMargin: ptG > 0 ? ptNet / ptG : 0,
          grossMargin: ptG > 0 ? (ptG - ptC) / ptG : 0,
          voucherRate: ptG > 0 ? ptV / ptG : 0,
          feeRate: ptG > 0 ? ptF / ptG : 0,
          channelMargin: ptG > 0 ? (ptG - ptC - ptF - ptV) / ptG : 0,
        };
      }
    }
    let badge = { label: "虧損", color: "var(--dn)" };
    let conclusion = "出現虧損，請立即檢視定價與費用。";
    if (netMargin >= targetNet) {
      badge = { label: "優秀", color: "var(--up)" };
      conclusion = `淨利率 ${fmtP(netMargin)} 超越 ${fmtP(targetNet)} 目標。`;
    } else if (netMargin >= targetNet * 0.6) {
      badge = { label: "穩健", color: "var(--orange)" };
      conclusion = `淨利率 ${fmtP(netMargin)}，距高標差 ${(
        (targetNet - netMargin) *
        100
      ).toFixed(1)}%。`;
    } else if (netMargin > 0) {
      badge = { label: "偏弱", color: "var(--wn)" };
      conclusion = `淨利率僅 ${fmtP(netMargin)}，建議優化。`;
    }
    const top = Object.values(prods).reduce(
      (m, o) => (o.estProfit > m.estProfit ? o : m),
      { estProfit: -Infinity, name: "—" }
    );
    return {
      years,
      months,
      orderList,
      uniqueProducts: Object.values(prods).sort(
        (a, b) => b.soldQty - a.soldQty
      ),
      prevMonth: spPrevM,
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
        badge,
        conclusion,
        top,
        topProfit: top.estProfit,
        avgAOV: validN > 0 ? tG / validN : 0,
        avgNetPer: validN > 0 ? afterComm / validN : 0,
        grossMargin: tG > 0 ? (tG - tF - tV - tC) / tG : 0,
        feeRate: tG > 0 ? tF / tG : 0,
        voucherRate: tG > 0 ? tV / tG : 0,
      },
    };
  }, [spOrders, sY, sM, spFp, spCosts, commissions]);

  /* ─── Derived state ────────────────────────────────────────── */
  const isOverview = platform === "overview";
  const isSL = platform === "shopline";
  const costs = isSL ? slCosts : spCosts;
  const setCosts = isSL ? setSlCosts : setSpCosts;
  const currentData = isSL ? slData : isOverview ? null : spData;
  const aY = isOverview
    ? [...new Set([...(slData?.years || []), ...(spData?.years || [])])]
        .sort()
        .reverse()
    : (isSL ? slData : spData)?.years || [];
  const aM = isOverview
    ? sY !== "All"
      ? [
          ...new Set(
            [
              ...Object.values(slOrders)
                .filter((o) => o.date.startsWith(sY))
                .map((o) => o.date.substring(5, 7)),
              ...Object.values(spOrders)
                .filter((o) => String(o.date).startsWith(sY))
                .map((o) => String(o.date).substring(5, 7)),
            ].filter(Boolean)
          ),
        ].sort()
      : []
    : (isSL ? slData : spData)?.months || [];

  const autoJumpedRef = useRef(false);
  useEffect(() => {
    if (!autoJumpedRef.current) setSM("All");
  }, [sY]);
  useEffect(() => {
    setPage(0);
  }, [lossOnly, search, orderSort, sY, sM, platform]);

  useEffect(() => {
    if (autoJumpedRef.current) return;
    const slVals = Object.values(slOrders);
    const spVals = Object.values(spOrders);
    if (!slVals.length && !spVals.length) return;
    const dates = [
      ...slVals.map((o) => o.date),
      ...spVals.map((o) => String(o.date)),
    ]
      .filter(Boolean)
      .sort()
      .reverse();
    if (!dates.length) return;
    autoJumpedRef.current = true;
    const ly = dates[0].substring(0, 4);
    const lm = dates[0].substring(5, 7);
    setSY(ly);
    setTimeout(() => {
      setSM(lm);
    }, 0);
  }, [slOrders, spOrders]);

  const matrixList = useMemo(() => {
    const source = isSL ? slData?.matrixList : spData?.uniqueProducts;
    if (!source) return [];
    return source
      .filter(
        (p) =>
          !mSearch ||
          p.name.toLowerCase().includes(mSearch.toLowerCase()) ||
          (p.option || "").toLowerCase().includes(mSearch.toLowerCase())
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
        if (key === "cost")
          return (
            m * ((Number(costs[a.key]) || 0) - (Number(costs[b.key]) || 0))
          );
        return 0;
      });
  }, [isSL, slData, spData, mSearch, costSort, slCosts, spCosts]);

  const missCost = useMemo(() => {
    const miss = matrixList.filter((p) => {
      const v = costs[p.key];
      return v === undefined || v === null || v === "" || Number(v) === 0;
    });
    return {
      total: matrixList.length,
      n: miss.length,
      keys: new Set(miss.map((p) => p.key)),
    };
  }, [matrixList, slCosts, spCosts]);

  const filteredOrders = useMemo(() => {
    if (!currentData) return [];
    const list = isSL ? currentData.orderList : currentData.orderList;
    return list
      .filter((o) => {
        if (lossOnly && (isSL ? o.net >= 0 : o.finalNetProfit >= 0))
          return false;
        if (search) {
          const t = search.toLowerCase();
          const oid = String(isSL ? o.orderId : o.orderId).toLowerCase();
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
                fee: o.pFee,
                cost: o.oCost,
                profit: o.currentOrderContribution,
                net: o.net,
              }
            : {
                date: o.date,
                revenue: o.localGross,
                fee: o.totalOrderFee,
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
  }, [currentData, isSL, lossOnly, search, orderSort]);

  const totalPages = Math.max(1, Math.ceil(filteredOrders.length / pageSize));
  const pagedOrders = useMemo(
    () => filteredOrders.slice(page * pageSize, (page + 1) * pageSize),
    [filteredOrders, page, pageSize]
  );

  const isLocked = useMemo(() => {
    if (!currentData?.orderList?.length) return false;
    const src = isSL ? slOrders : spOrders;
    return currentData.orderList.every((o) => {
      const t = src[o.orderId];
      return (
        t?.snapshotFeeParams &&
        t?.items?.length &&
        t.items.every((i) =>
          Object.prototype.hasOwnProperty.call(i, "snapshotCost")
        )
      );
    });
  }, [currentData, isSL, slOrders, spOrders]);

  const toggleSnap = () => {
    if (!currentData?.orderList?.length) return;
    if (!window.confirm(isLocked ? "解除快照？" : "鎖定本期成本？")) return;
    const fp = isSL ? slFp : spFp;
    const setter = isSL ? setSlOrders : setSpOrders;
    const src = isSL ? slOrders : spOrders;
    const no = { ...src };
    currentData.orderList.forEach((o) => {
      const tg = no[o.orderId];
      if (!tg?.items?.length) return;
      if (isLocked) {
        const nx = { ...tg };
        nx.items = tg.items.map((i) => {
          const ni = { ...i };
          delete ni.snapshotCost;
          return ni;
        });
        delete nx.snapshotFeeParams;
        no[o.orderId] = nx;
      } else {
        no[o.orderId] = {
          ...tg,
          snapshotFeeParams: {
            platformFeeRate: Number(fp.platformFeeRate) ?? null,
            opExpense: Number(fp.opExpense) ?? null,
            tax: Number(fp.tax) ?? null,
            targetNet: Number(fp.targetNet) ?? null,
          },
          items: tg.items.map((i) => ({
            ...i,
            snapshotCost:
              costs[i.key] === undefined ? null : Number(costs[i.key]),
          })),
        };
      }
    });
    setter(no);
    msg(isLocked ? "已解除快照" : "已鎖定快照");
  };

  const expC = () => {
    const b = new Blob([JSON.stringify(costs, null, 2)], {
      type: "application/json",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(b);
    a.download = `${isSL ? "sl" : "sp"}_costs_${
      new Date().toISOString().split("T")[0]
    }.json`;
    a.click();
  };
  const impC = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    const r = new FileReader();
    r.onload = (ev) => {
      try {
        setCosts((p) => ({ ...p, ...JSON.parse(ev.target.result) }));
        msg("匯入成功");
      } catch {
        msg("匯入失敗");
      }
    };
    r.readAsText(f);
    e.target.value = "";
  };

  const handleComm = (key, value) =>
    setCommissions((prev) => {
      if (value === "" || value === null || value === undefined) {
        const n = { ...prev };
        delete n[key];
        return n;
      }
      return { ...prev, [key]: Number(value) };
    });

  /* ── 未填成本跳轉 helper ── */
  const jumpToFirstMissCost = () => {
    setTimeout(() => {
      if (firstMissRef.current) {
        firstMissRef.current.scrollIntoView({
          behavior: "smooth",
          block: "center",
        });
        firstMissRef.current.style.outline = "2px solid var(--wn)";
        firstMissRef.current.style.outlineOffset = "2px";
        setTimeout(() => {
          if (firstMissRef.current) firstMissRef.current.style.outline = "none";
        }, 2000);
      }
    }, 50);
  };

  const slD = slData?.summary;
  const spS = spData?.s;

  const accentColor = isSL ? "var(--accent)" : "var(--sp-accent)";
  const accentDim = isSL ? "var(--accent-dim)" : "var(--sp-accent-dim)";
  const accentBdr = isSL ? "var(--accent-bdr)" : "var(--sp-accent-bdr)";

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

      {/* Notification */}
      {notif && (
        <div
          style={{
            position: "fixed",
            top: 12,
            left: "50%",
            transform: "translateX(-50%)",
            zIndex: 200,
            background: "var(--t1)",
            color: "var(--bg)",
            borderRadius: 10,
            padding: "10px 20px",
            boxShadow: "0 12px 40px rgba(0,0,0,.2)",
            display: "flex",
            alignItems: "center",
            gap: 8,
            fontWeight: 600,
            fontSize: 12,
            animation: "notifIn .25s ease both",
          }}
        >
          <Zap size={13} />
          {notif}
        </div>
      )}

      {/* Header */}
      <header
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
                {isOverview ? "跨平台" : isSL ? "官網" : "蝦皮"} 利潤決策中心
              </h1>
              <div
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
              ].map((p) => (
                <button
                  key={p.id}
                  onClick={() => {
                    setPlatform(p.id);
                    if (p.id === "shopline") {
                      const allOrders = Object.values(slOrders);
                      if (allOrders.length) {
                        const ly = allOrders
                          .map((o) => o.date.substring(0, 4))
                          .filter(Boolean)
                          .sort()
                          .reverse()[0];
                        const lm =
                          allOrders
                            .filter((o) => o.date.startsWith(ly))
                            .map((o) => o.date.substring(5, 7))
                            .filter(Boolean)
                            .sort()
                            .reverse()[0] || "All";
                        setSY(ly);
                        setSM(lm);
                      } else {
                        setSY("All");
                        setSM("All");
                      }
                    } else if (p.id === "shopee") {
                      const allOrders = Object.values(spOrders);
                      if (allOrders.length) {
                        const ly = allOrders
                          .map((o) => String(o.date).substring(0, 4))
                          .filter(Boolean)
                          .sort()
                          .reverse()[0];
                        const lm =
                          allOrders
                            .filter((o) => String(o.date).startsWith(ly))
                            .map((o) => String(o.date).substring(5, 7))
                            .filter(Boolean)
                            .sort()
                            .reverse()[0] || "All";
                        setSY(ly);
                        setSM(lm);
                      } else {
                        setSY("All");
                        setSM("All");
                      }
                    }
                  }}
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
              onChange={(e) => setSY(e.target.value)}
              style={sel}
            >
              <option value="All">歷年數據</option>
              {aY.map((y) => (
                <option key={y} value={y}>
                  {y} 年
                </option>
              ))}
            </select>
            <select
              value={sM}
              onChange={(e) => setSM(e.target.value)}
              style={sel}
            >
              <option value="All">全月份</option>
              {aM.map((m) => (
                <option key={m} value={m}>
                  {m} 月
                </option>
              ))}
            </select>
          </div>
        </div>
      </header>

      <div
        style={{ maxWidth: 1560, margin: "0 auto", padding: "20px 24px 80px" }}
      >
        <div className={isOverview ? "" : "gm"}>
          {/* Sidebar */}
          <aside
            className="f0"
            style={{ display: isOverview ? "none" : undefined }}
          >
            <div
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
              {/* Upload */}
              <div
                onClick={() => fRef.current._fileInput?.click()}
                style={{
                  border: "1.5px dashed var(--s4)",
                  borderRadius: 12,
                  padding: "18px 12px",
                  textAlign: "center",
                  cursor: "pointer",
                  background: "var(--s2)",
                }}
              >
                <input
                  ref={(el) => {
                    if (fRef.current) fRef.current._fileInput = el;
                  }}
                  type="file"
                  accept=".csv,.xlsx,.xls"
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
                  匯入{isSL ? "官網" : "蝦皮"}報表
                </div>
                <div style={{ fontSize: 10, color: "var(--t4)", marginTop: 2 }}>
                  CSV · XLSX · 拖曳
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
                  {isSL ? slD?.valid : spS?.validN} 筆 ·{" "}
                  {sY === "All" ? "歷年" : sY}
                  {sM !== "All" ? `/${sM}` : ""}
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
                </div>
                {(isSL
                  ? [
                      { l: "淨利目標", n: "targetNet" },
                      { l: "內部營業費", n: "opExpense" },
                      { l: "預估稅率", n: "tax" },
                      { l: "系統服務費率", n: "platformFeeRate" },
                    ]
                  : [
                      { l: "淨利目標", n: "targetNet" },
                      { l: "內部營業費", n: "opExpense" },
                      { l: "預估稅率", n: "tax" },
                    ]
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
                      <input
                        type="number"
                        step="0.1"
                        value={isSL ? slFp[item.n] : spFp[item.n]}
                        onChange={(e) =>
                          isSL
                            ? setSlFp((p) => ({
                                ...p,
                                [item.n]: e.target.value,
                              }))
                            : setSpFp((p) => ({
                                ...p,
                                [item.n]: e.target.value,
                              }))
                        }
                        style={{ ...inp, width: 60, fontSize: 12 }}
                      />
                      <span style={{ fontSize: 11, color: "var(--t4)" }}>
                        %
                      </span>
                    </div>
                  </div>
                ))}
              </div>

              {/* Commission Panel (Shopee only) */}
              {!isSL && currentData && (
                <CommissionPanel
                  commissions={commissions}
                  onUpdate={handleComm}
                  selYear={sY}
                  selMonth={sM}
                />
              )}

              {/* Reset */}
              <div style={{ display: "flex", gap: 6 }}>
                <Btn
                  v="primary"
                  onClick={() => {
                    const src = isSL ? slOrders : spOrders;
                    const toDelete = Object.keys(src).filter((k) => {
                      const d = String(src[k].date || "");
                      if (sY !== "All" && !d.startsWith(sY)) return false;
                      if (sM !== "All" && !d.startsWith(`${sY}-${sM}`))
                        return false;
                      return true;
                    });
                    if (!toDelete.length) {
                      msg("本期無訂單可清除");
                      return;
                    }
                    const periodLabel =
                      sY === "All"
                        ? "歷年"
                        : sM === "All"
                        ? `${sY} 年`
                        : `${sY}/${sM}`;
                    if (
                      !window.confirm(
                        `重置「${periodLabel}」${
                          isSL ? "官網" : "蝦皮"
                        }訂單（共 ${
                          toDelete.length
                        } 筆）？\n\n舊資料將被清除，請重新匯入該期報表。`
                      )
                    )
                      return;
                    const updated = { ...src };
                    toDelete.forEach((k) => delete updated[k]);
                    if (isSL) setSlOrders(updated);
                    else setSpOrders(updated);
                    toast(`已清除 ${toDelete.length} 筆，請重新上傳報表`, {
                      type: "info",
                      duration: 5000,
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
          <main style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            {isOverview ? (
              <OverviewDashboard
                slData={slData}
                spData={spData}
                slOrders={slOrders}
                spOrders={spOrders}
                slCosts={slCosts}
                spCosts={spCosts}
                theme={theme}
                sY={sY}
                sM={sM}
                onNavigate={(id) => {
                  setPlatform(id);
                  if (id === "shopline" && slData?.years?.length) {
                    const ly = slData.years[0];
                    setSY(ly);
                    setSM("All");
                  } else if (id === "shopee" && spData?.years?.length) {
                    const ly = spData.years[0];
                    const ms = [
                      ...new Set(
                        Object.values(spOrders || {})
                          .filter((o) => String(o.date).startsWith(ly))
                          .map((o) => String(o.date).substring(5, 7))
                          .filter(Boolean)
                      ),
                    ]
                      .sort()
                      .reverse();
                    setSY(ly);
                    setSM(ms[0] || "All");
                  } else {
                    setSY("All");
                    setSM("All");
                  }
                }}
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
                  上傳{isSL ? "官網" : "蝦皮"}訂單報表以啟動分析
                </div>
              </div>
            ) : (
              <>
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
                        {missCost.n > 0 && (
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
                            style={{
                              fontSize: 72,
                              lineHeight: 1,
                              fontWeight: 700,
                              letterSpacing: "-0.04em",
                              fontFamily: mono,
                              color: "var(--t1)",
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
                            style={{
                              fontSize: 48,
                              fontWeight: 700,
                              fontFamily: mono,
                              lineHeight: 1,
                              color:
                                slD.trueNetMargin >= 0.19
                                  ? "var(--up)"
                                  : slD.trueNetMargin >= 0.16
                                  ? "var(--accent)"
                                  : slD.trueNetMargin >= 0.14
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
                      {/* Waterfall */}
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
                              l: "毛利",
                              v:
                                slD.contributionMargin +
                                slD.opExpTotal +
                                slD.taxTotal,
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
                              : slD.grossMargin >= 0.6
                              ? "✓ 達標，目標 60%"
                              : slD.grossMargin >= 0.58
                              ? "⚠ 正常帶下緣，目標 60%"
                              : "⚠ 低於警戒線 58%，檢視成本與定價",
                        },
                        {
                          l: "通路後毛利率",
                          v: fmtP(
                            slD.rev > 0 ? slD.contributionMargin / slD.rev : 0
                          ),
                          c: (() => {
                            const r =
                              slD.rev > 0
                                ? slD.contributionMargin / slD.rev
                                : 0;
                            return r >= 0.565
                              ? "var(--up)"
                              : r >= 0.53
                              ? "var(--accent)"
                              : "var(--dn)";
                          })(),
                          h: (() => {
                            const r =
                              slD.rev > 0
                                ? slD.contributionMargin / slD.rev
                                : 0;
                            return r >= 0.565
                              ? "✓ 超越目標 56.5%，成本控管優秀"
                              : r >= 0.55
                              ? "✓ 達標，目標 55%"
                              : r >= 0.53
                              ? "⚠ 正常帶下緣，目標 55%"
                              : "⚠ 低於警戒線 53%，檢視通路費與折扣";
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
                              : slD.voucherRate <= 0.045
                              ? "var(--up)"
                              : "var(--purple)",
                          h:
                            slD.voucherRate > 0.065
                              ? "⚠ 品牌警戒！單月 >6.5%，立即介入"
                              : slD.voucherRate > 0.055
                              ? "⚠ 超出警戒線 5.5%，啟動檢討"
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
                        {missCost.n > 0 && (
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
                        {spS.comm > 0 && (
                          <Tag v="warn">
                            <Users size={10} /> 已扣分潤 {fmt$(spS.comm)}
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
                            style={{
                              fontSize: 72,
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
                            style={{
                              fontSize: 48,
                              fontWeight: 700,
                              fontFamily: mono,
                              lineHeight: 1,
                              color:
                                spS.netMargin >= 0.13
                                  ? "var(--up)"
                                  : spS.netMargin >= 0.09
                                  ? "var(--accent)"
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
                      {/* Waterfall */}
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
                              l: "毛利",
                              v:
                                spS.tG -
                                spS.tC -
                                spS.tF -
                                spS.tV +
                                spS.tOp +
                                spS.tTx,
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
                          h: `賣場券 -${fmt$(spS.tV)} ｜ 手續費 -${fmt$(
                            spS.tF
                          )}`,
                        },
                        {
                          l: "預估總入帳",
                          v: fmt$(spS.tG - spS.tV - spS.tF),
                          c: "var(--blue)",
                          h: "扣除賣場券與手續費",
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
                              ? "✓ 達標，目標 65~67%"
                              : r >= 0.63
                              ? "⚠ 正常帶下緣，目標 65~67%"
                              : "⚠ 低於警戒線 63%，檢視定價";
                          })(),
                        },
                        {
                          l: "真實抽成率",
                          v: fmtP(spS.feeRate),
                          c: "var(--orange)",
                          note: "蝦皮固定抽成費率",
                        },
                        {
                          l: "優惠券發放率",
                          v: fmtP(spS.voucherRate),
                          c:
                            spS.voucherRate > 0.03
                              ? "var(--dn)"
                              : spS.voucherRate > 0.025
                              ? "var(--wn)"
                              : spS.voucherRate <= 0.015
                              ? "var(--up)"
                              : "var(--purple)",
                          note:
                            spS.voucherRate > 0.03
                              ? "⚠ 品牌警戒！單月 >3%，立即介入"
                              : spS.voucherRate > 0.025
                              ? "⚠ 超出警戒線 2.5%，啟動檢討"
                              : spS.voucherRate <= 0.015
                              ? "✓ 在目標範圍 0.8~1.5% 內"
                              : "注意：接近警戒線 2.5%",
                        },
                        {
                          l: "通路後毛利率",
                          v: fmtP(spS.grossMargin),
                          c:
                            spS.grossMargin >= 0.49
                              ? "var(--up)"
                              : spS.grossMargin >= 0.44
                              ? "var(--accent)"
                              : "var(--dn)",
                          note:
                            spS.grossMargin >= 0.49
                              ? "✓ 超越目標 49%，費用控管優秀"
                              : spS.grossMargin >= 0.46
                              ? "✓ 達標，目標 46~48%"
                              : spS.grossMargin >= 0.44
                              ? "⚠ 正常帶下緣，目標 46~48%"
                              : "⚠ 低於警戒線 44%，檢視平台費用",
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
                  </>
                )}

                {/* ── Cost Matrix ── */}
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
                        共 {missCost.total} 項
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
                  <div style={{ position: "relative", marginBottom: 14 }}>
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
                      value={mSearch}
                      onChange={(e) => setMSearch(e.target.value)}
                      style={{
                        ...inp,
                        width: "100%",
                        maxWidth: 360,
                        textAlign: "left",
                        paddingLeft: 36,
                        borderRadius: 10,
                        padding: "10px 12px 10px 36px",
                        fontSize: 13,
                      }}
                    />
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
                    <table
                      style={{
                        width: "100%",
                        borderCollapse: "collapse",
                        minWidth: 680,
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
                          <th style={{ ...th, textAlign: "left" }}>規格</th>
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
                            style={{ ...th, textAlign: "center", width: 40 }}
                          ></th>
                        </tr>
                      </thead>
                      <tbody>
                        {!matrixList.length ? (
                          <tr>
                            <td
                              colSpan={6}
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
                              return (
                                <tr
                                  key={p.key}
                                  ref={isFirstMiss ? firstMissRef : undefined}
                                  className={miss && hs ? "rw" : ""}
                                >
                                  <td style={{ ...td2, fontWeight: 600 }}>
                                    {p.name}
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
                                  <td style={{ ...td2, textAlign: "right" }}>
                                    <div
                                      style={{
                                        display: "flex",
                                        alignItems: "center",
                                        justifyContent: "flex-end",
                                        gap: 4,
                                      }}
                                    >
                                      {miss && hs && (
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
                                      <input
                                        type="number"
                                        value={costs[p.key] || ""}
                                        onChange={(e) =>
                                          setCosts((pr) => ({
                                            ...pr,
                                            [p.key]:
                                              parseFloat(e.target.value) || 0,
                                          }))
                                        }
                                        className={
                                          miss
                                            ? "iw"
                                            : costs[p.key] > 0
                                            ? "iok"
                                            : ""
                                        }
                                        placeholder="—"
                                        style={{ ...inp, width: 80 }}
                                      />
                                    </div>
                                  </td>
                                  <td style={{ ...td2, textAlign: "center" }}>
                                    <Btn
                                      v="ghost"
                                      onClick={() => {
                                        if (window.confirm("確定刪除？")) {
                                          const n = { ...costs };
                                          delete n[p.key];
                                          setCosts(n);
                                        }
                                      }}
                                      style={{ padding: "2px" }}
                                    >
                                      <Trash2 size={12} color="var(--t4)" />
                                    </Btn>
                                  </td>
                                </tr>
                              );
                            });
                          })()
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

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
                      {isSL && (
                        <span style={{ fontSize: 11, color: "var(--dn)" }}>
                          虧損 {slData?.lossCount} 筆
                        </span>
                      )}
                      {!isSL && (
                        <span style={{ fontSize: 11, color: "var(--dn)" }}>
                          虧損 {spS?.lossN} 筆
                        </span>
                      )}
                    </div>
                    <div
                      style={{
                        display: "flex",
                        gap: 6,
                        alignItems: "center",
                        flexWrap: "wrap",
                      }}
                    >
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
                        虧損篩選
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
                    <table
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
                            <th style={{ ...th, textAlign: "left" }}>商品</th>
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
                            手續費
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
                            const fee = isSL
                              ? o.pFee +
                                o.revenue * ((slFp.platformFeeRate || 1) / 100)
                              : o.totalOrderFee;
                            const cost = isSL ? o.oCost : o.orderCost;
                            const gross = isSL
                              ? o.currentOrderContribution
                              : o.grossProfit;
                            const net = isSL ? o.net : o.finalNetProfit;
                            return (
                              <tr
                                key={o.orderId}
                                className={isLoss ? "rl" : ""}
                              >
                                <td style={{ ...td2 }}>
                                  <div
                                    style={{ fontWeight: 600, fontSize: 12 }}
                                  >
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
                          {page * pageSize + 1}–
                          {Math.min(
                            (page + 1) * pageSize,
                            filteredOrders.length
                          )}{" "}
                          / {filteredOrders.length} 筆
                        </span>
                        <span style={{ fontSize: 10, color: "var(--t4)" }}>
                          每頁
                        </span>
                        <select
                          value={pageSize}
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
                          { label: "«", action: () => setPage(0) },
                          {
                            label: "‹",
                            action: () => setPage((p) => Math.max(0, p - 1)),
                          },
                          null,
                          {
                            label: "›",
                            action: () =>
                              setPage((p) => Math.min(totalPages - 1, p + 1)),
                          },
                          { label: "»", action: () => setPage(totalPages - 1) },
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
                              {page + 1} / {totalPages}
                            </span>
                          ) : (
                            <Btn
                              key={i}
                              v="ghost"
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
          </main>
        </div>
      </div>

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
