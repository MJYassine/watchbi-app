import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  BarChart, Bar, XAxis, YAxis,
  CartesianGrid, Tooltip, ResponsiveContainer, Cell,
} from "recharts";
import {
  Upload, FileSpreadsheet, Boxes, TrendingUp, Sparkles,
  Check, X, AlertTriangle, Lock, Watch, Trash2, ArrowRight, Bell,
} from "lucide-react";
import * as XLSX from "xlsx";

/* =========================================================================
   THEME — refined "horology" look: warm near-black, champagne accent, cream.
   ========================================================================= */
const C = {
  bg: "#15120e",
  panel: "#1d1913",
  panel2: "#241f17",
  line: "#3a3225",
  text: "#efe7d6",
  dim: "#a99f88",
  faint: "#6f6754",
  gold: "#c8a96a",
  gold2: "#e2c98c",
  green: "#7faa6f",
  red: "#c1715f",
  blue: "#7c93b3",
};
const SERIF = "'Fraunces', Georgia, 'Times New Roman', serif";
const SANS = "'DM Sans', -apple-system, system-ui, sans-serif";

const TARGET_BRANDS = ["rolex", "patek philippe", "audemars piguet", "omega"];
const BRAND_ALIAS = {
  patek: "patek philippe", "patek phillipe": "patek philippe", pp: "patek philippe",
  ap: "audemars piguet", audemars: "audemars piguet", "audemars piguet": "audemars piguet",
  rolex: "rolex", omega: "omega",
};

const LINE_KEYWORDS = {
  rolex: ["sea-dweller", "sky-dweller", "yacht-master", "gmt-master", "day-date",
    "datejust", "submariner", "daytona", "explorer", "oyster perpetual", "air-king",
    "milgauss", "cellini"],
  "patek philippe": ["nautilus", "aquanaut", "calatrava", "grand complications",
    "complications", "twenty", "golden ellipse", "gondolo"],
  "audemars piguet": ["royal oak offshore", "royal oak concept", "royal oak",
    "code 11.59", "millenary"],
  omega: ["speedmaster", "moonwatch", "seamaster", "diver 300m", "aqua terra",
    "planet ocean", "constellation", "de ville", "railmaster", "ploprof"],
};
const LINE_NORMALIZE = {
  moonwatch: "Speedmaster", "diver 300m": "Seamaster",
  "aqua terra": "Seamaster", "planet ocean": "Seamaster",
};

/* Canonical fields per dataset role. */
const FIELDS = {
  inventory: [
    ["brand", "Brand"], ["modelName", "Model name"], ["modelNumber", "Model / reference #"],
    ["cost", "Cost (total)"], ["purchaseDate", "Purchase date"],
    ["targetWholesale", "Target wholesale price"], ["tagPrice", "Tag price"],
    ["condition", "Condition — New/Used (optional)"], ["status", "Status"],
    ["invType", "Inventory type — Owned/Consignment/Memo (optional)"],
    ["paymentStatus", "Payment status — Paid/Unpaid/Voided (optional)"],
    ["supplier", "Supplier / vendor (optional)"], ["serial", "Serial #"],
  ],
  sales: [
    ["saleDate", "Sale / invoice date"], ["purchaseDate", "Purchase date"],
    ["cost", "Cost (COGS)"], ["salePrice", "Sale / invoice price"],
    ["profit", "Profit $ (optional)"], ["brand", "Brand (optional)"],
    ["modelName", "Model name (optional)"], ["modelNumber", "Model / reference # (optional)"],
    ["inventoryType", "Inventory type (optional)"], ["condition", "Condition — New/Used (optional)"],
    ["shippingState", "Shipping state (optional)"], ["supplier", "Supplier / vendor (optional)"],
    ["salesperson", "Salesperson — created by (optional)"],
    ["amountOwed", "Amount owed / remaining balance (optional)"], ["customer", "Customer name (optional)"],
    ["serial", "Serial # (optional)"],
  ],
  messages: [
    ["intent", "Intent (buy/sell)"], ["brand", "Brand"], ["model", "Model"],
    ["reference", "Reference #"], ["price", "Price"], ["messageBody", "Message text"],
    ["timestamp", "Timestamp"], ["sender", "Sender"], ["chat", "Chat / group"],
  ],
};

/* ---- header auto-detect ---- */
function guessField(role, header) {
  const h = String(header).trim().toLowerCase();
  const has = (s) => h.includes(s);
  if (role === "messages") {
    if (h === "intent") return "intent";
    if (h === "fullreferencenumber" || h === "reference" || has("reference number")) return "reference";
    if (h === "messagebody" || has("message body") || h === "body") return "messageBody";
    if (h === "timestamp" || h === "date" || has("time")) return "timestamp";
    if (h === "sendername" || has("sender")) return "sender";
    if (h === "chatname" || has("chat") || has("group")) return "chat";
    if (h === "price") return "price";
    if (h === "model") return "model";
    if (h === "brand") return "brand";
    return null;
  }
  if (has("brand")) return "brand";
  if (has("model name") || has("title item")) return "modelName";
  if (has("model number") || has("reference") || h === "ref") return "modelNumber";
  if (has("condition")) return "condition";
  if (has("supplier") || has("vendor") || has("consignor") || has("bought from") || has("purchased from") || has("source")) return "supplier";
  if (has("payment") || has("paid")) return role === "inventory" ? "paymentStatus" : null;
  if (role === "sales") {
    if (has("created by") || has("salesperson") || has("sales person") || has("sold by") || has("sales rep")) return "salesperson";
    if (has("remaining balance") || has("amount owed") || has("balance due") || h === "balance") return "amountOwed";
    if (has("customer name") || has("to who") || h === "customer") return "customer";
    if (has("invoice date") || has("sale date") || has("sold date")) return "saleDate";
    if (has("invoice price") || has("sale price") || has("sold price")) return "salePrice";
    if (h === "profit") return "profit";
    if ((has("ship") && has("state")) || h === "state" || has("ship to state")) return "shippingState";
    if (has("inventory type") || h === "type") return "inventoryType";
  }
  if (role === "inventory") {
    if (has("inventory type") || h === "type") return "invType";
  }
  if (has("purchase date") || has("purchased date")) return "purchaseDate";
  if (has("cogs") || has("total cost") || h === "cost") return "cost";
  if (has("wholesale")) return "targetWholesale";
  if (has("tag price")) return "tagPrice";
  if (has("status") && !has("payment")) return "status";
  if (has("serial")) return "serial";
  return null;
}

function autoMap(role, columns) {
  const map = {};
  const taken = new Set();
  for (const col of columns) {
    const f = guessField(role, col);
    if (f && !taken.has(f)) { map[f] = col; taken.add(f); }
  }
  return map;
}

function guessRole(columns) {
  const lower = columns.map((c) => String(c).toLowerCase());
  const any = (s) => lower.some((c) => c.includes(s));
  if (any("intent") && (any("messagebody") || any("messageid") || any("chatname"))) return "messages";
  if (any("invoice date") || any("invoice price") || any("sale date")) return "sales";
  if (any("brand") || any("wholesale") || any("inventory status") || any("model number"))
    return "inventory";
  return "ignore";
}

/* ---- value coercion ---- */
function toNum(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? null : n;
}
function toDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === "number") { // excel serial
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isNaN(d) ? null : d;
  }
  const d = new Date(v);
  return isNaN(d) ? null : d;
}
function daysBetween(a, b) {
  if (!a || !b) return null;
  return Math.round((b - a) / 86400000);
}

/* normalize a free-text condition value to "New" / "Used" (falls back to the
   raw value, title-cased, if it's neither — e.g. "CPO" / "Pre-Owned") */
function normalizeCondition(v) {
  if (v == null || v === "") return "Unspecified";
  const s = String(v).trim().toLowerCase();
  if (!s) return "Unspecified";
  if (s.includes("new")) return "New";
  if (s.includes("used") || s.includes("pre-owned") || s.includes("preowned") || s.includes("pre owned") || s.includes("second")) return "Used";
  return String(v).trim().replace(/\b\w/g, (m) => m.toUpperCase());
}
/* Display name for a watch in the brand/line/model breakdowns.
   - has a name              -> the name
   - brand but no name        -> "Unknown" (bucketed under that brand)
   - no brand and no name     -> null  (fully unknown → hidden from charts) */
function modelDisplayName(brand, modelName) {
  const b = brand != null && String(brand).trim();
  const n = modelName != null && String(modelName).trim();
  if (n) return String(modelName).trim();
  if (b) return "Unknown";
  return null;
}

/* canonical key identifying a watch model for the exclusion filter — prefers
   the reference/model number, falls back to the model name */
function modelKeyOf(modelNumber, modelName) {
  const num = modelNumber != null ? String(modelNumber).trim() : "";
  const name = modelName != null ? String(modelName).trim() : "";
  return num || name || "";
}

/* rank conditions best -> worst so "New" sorts first, then down in quality */
function conditionRank(c) {
  const s = String(c || "").toLowerCase();
  if (s === "new") return 0;
  if (s.includes("mint") || s.includes("excellent") || s.includes("very good")) return 1;
  if (s.includes("good")) return 2;
  if (s.includes("used") || s.includes("fair")) return 3;
  if (!s || s === "unspecified") return 99;
  return 50;
}

/* normalize a free-text payment value to Paid / Unpaid / Voided */
function normalizePayment(v) {
  if (v == null || v === "") return "Unspecified";
  const s = String(v).trim().toLowerCase();
  if (!s) return "Unspecified";
  if (s.includes("void") || s.includes("cancel")) return "Voided";
  if (s.includes("unpaid") || s.includes("not paid") || s.includes("due") || s.includes("outstanding") || s === "no") return "Unpaid";
  if (s.includes("paid") || s === "yes") return "Paid";
  return String(v).trim().replace(/\b\w/g, (m) => m.toUpperCase());
}
const isMemoType = (t) => /memo/i.test(String(t || ""));
const isConsignmentType = (t) => /consign/i.test(String(t || ""));

/* these reference numbers are excluded from velocity / median-days-to-sell
   calculations (kept in unit counts / weekly average), per dealer request */
const VELOCITY_IGNORE_REFS = ["210.90.42.20.01.001", "310.30.42.50.01.002"];
function isVelocityIgnored(...vals) {
  const norm = vals.map((v) => String(v || "").replace(/\s+/g, ""));
  return VELOCITY_IGNORE_REFS.some((ref) => norm.some((v) => v.includes(ref)));
}

/* illustrative ("phony") sales-tax rules: rate applies to state sales above the
   threshold. Only these states have a rule; others collect no tax. */
const STATE_TAX_RULES = {
  "New York": { rate: 0.08, threshold: 500000 },
  "Florida": { rate: 0.07, threshold: 100000 },
  "California": { rate: 0.095, threshold: 500000 },
  "Georgia": { rate: 0.07, threshold: 100000 },
};

/* price tier buckets for sale price breakdowns */
const PRICE_TIERS = [
  { label: "$0-5k", lo: 0, hi: 5000 },
  { label: "$5-10k", lo: 5000, hi: 10000 },
  { label: "$10-20k", lo: 10000, hi: 20000 },
  { label: "$20-30k", lo: 20000, hi: 30000 },
  { label: "$30-40k", lo: 30000, hi: 40000 },
  { label: "$40-50k", lo: 40000, hi: 50000 },
  { label: "$50-100k", lo: 50000, hi: 100000 },
  { label: "$100k+", lo: 100000, hi: Infinity },
];
function priceTierLabel(price) {
  if (price == null) return null;
  for (const t of PRICE_TIERS) {
    if (price >= t.lo && price < t.hi) return t.label;
  }
  return PRICE_TIERS[PRICE_TIERS.length - 1].label;
}

/* inventory condition grade A (fresh) -> F (aged) based on days in stock */
const GRADE_BUCKETS = [
  { grade: "A", lo: 0, hi: 30 },
  { grade: "B", lo: 31, hi: 60 },
  { grade: "C", lo: 61, hi: 90 },
  { grade: "D", lo: 91, hi: 180 },
  { grade: "F", lo: 181, hi: Infinity },
];
function ageGrade(age) {
  if (age == null) return null;
  for (const b of GRADE_BUCKETS) {
    if (age >= b.lo && age <= b.hi) return b.grade;
  }
  return "F";
}
const GRADE_RANK = { A: 0, B: 1, C: 2, D: 3, F: 4 };

/* market-value grades vs WatchCharts median asking price (A best) */
// how well you SOLD: sale price as % of market. higher = better.
function salesGrade(price, market) {
  if (!market || price == null) return null;
  const pct = (price / market) * 100;
  if (pct >= 100) return "A"; if (pct >= 95) return "B"; if (pct >= 90) return "C"; return "D";
}
// Quality of Cost: what you PAID as % of market. lower = better.
function qocGrade(cost, market) {
  if (!market || !cost) return null;
  const pct = (cost / market) * 100;
  if (pct >= 100) return "D"; if (pct >= 95) return "C"; if (pct >= 90) return "B"; return "A";
}
// Quality of Target: your target price vs market. A = priced near market, then wider bands.
function qotGrade(target, market) {
  if (!market || !target || target <= 0) return null;
  const d = Math.abs((target / market) * 100 - 100);
  if (d <= 3) return "A"; if (d <= 8) return "B"; if (d <= 15) return "C"; return "D";
}

/* DJ's grade (single sale): a points system per DJ Allen — "each data point is 1 point."
   Scores Margin, Days-on-hand, and Gross-profit-$, each A=4..D=1, averaged → A/B/C/D.
   Averaging resolves conflicting factors (e.g. A-margin / D-days). All thresholds are
   configurable per dealer (this DJ_GRADE_CFG is DJ's default set — to be fine-tuned). */
const DJ_GRADE_CFG = {
  margin: { A: 30, B: 20, C: 10 },      // GP% (profit/cost), higher better
  days: { A: 30, B: 60, C: 120 },       // days on hand, lower better
  profit: { A: 5000, B: 2500, C: 1000 },// gross profit $, higher better
  weights: { margin: 1, days: 1, profit: 1 },
};
function factorPoints(value, bands, higherBetter) {
  if (value == null || isNaN(value)) return null;
  if (higherBetter) return value >= bands.A ? 4 : value >= bands.B ? 3 : value >= bands.C ? 2 : 1;
  return value <= bands.A ? 4 : value <= bands.B ? 3 : value <= bands.C ? 2 : 1;
}
function ptsToGrade(avg) {
  if (avg == null) return null;
  return avg >= 3.5 ? "A" : avg >= 2.5 ? "B" : avg >= 1.5 ? "C" : "D";
}
// sale = { marginPct, days, profit }; returns { grade, points, parts }
function djSaleGrade(sale, cfg = DJ_GRADE_CFG) {
  const parts = {
    margin: factorPoints(sale.marginPct, cfg.margin, true),
    days: factorPoints(sale.days, cfg.days, false),
    profit: factorPoints(sale.profit, cfg.profit, true),
  };
  let sum = 0, wsum = 0;
  for (const k of ["margin", "days", "profit"]) {
    if (parts[k] != null) { sum += parts[k] * cfg.weights[k]; wsum += cfg.weights[k]; }
  }
  if (!wsum) return { grade: null, points: null, parts };
  const avg = sum / wsum;
  return { grade: ptsToGrade(avg), points: +avg.toFixed(2), parts };
}

function normalizeBrand(b) {
  if (!b) return null;
  let s = String(b).trim().toLowerCase();
  if (BRAND_ALIAS[s]) s = BRAND_ALIAS[s];
  return s;
}
function findLine(brandNorm, modelName) {
  if (!brandNorm || !TARGET_BRANDS.includes(brandNorm)) return null;
  const name = String(modelName || "").toLowerCase();
  const kws = LINE_KEYWORDS[brandNorm] || [];
  for (const kw of kws) {
    if (name.includes(kw)) {
      const norm = LINE_NORMALIZE[kw];
      return norm || kw.replace(/\b\w/g, (m) => m.toUpperCase());
    }
  }
  return "Other";
}
/* "Other" is ambiguous on its own across brands — tie it to the brand name. */
function lineLabel(brand, line) {
  if (!line) return line;
  if (line === "Other") {
    const b = String(brand || "").trim();
    return b ? `${b} — Other` : "Other";
  }
  return line;
}

/* ---- stock health ----
   weeklyVelocity = units sold per week (over the observed sales window).
   weeksOfStock   = how many weeks current stock will last at that pace.
   red    = under 1 week of stock left (or zero stock with active sales) -> buy now
   yellow = under 2 weeks of stock left -> buy soon
   green  = healthy
*/
function stockHealth(stock, weeklyVelocity) {
  if (stock == null) return null;
  if (!weeklyVelocity || weeklyVelocity <= 0) return stock === 0 ? null : "green";
  if (stock === 0) return "red";
  const weeks = stock / weeklyVelocity;
  if (weeks < 1) return "red";
  if (weeks < 2) return "yellow";
  return "green";
}
/* Adds weeklyVelocity / weeksOfStock / health / buyScore to a list of
   { units, profit, avgProfit, medianDays, stock } rows. */
function enrichRanking(arr, weeks) {
  if (!arr || !arr.length) return arr || [];
  const maxV = Math.max(...arr.map((x) => x.units || 0)) || 1;
  const maxP = Math.max(...arr.map((x) => x.avgProfit || 0)) || 1;
  const maxD = Math.max(...arr.map((x) => x.medianDays || 0)) || 1;
  return arr.map((x) => {
    const weeklyVelocity = weeks > 0 ? x.units / weeks : 0;
    const weeksOfStock = (x.stock != null && weeklyVelocity > 0)
      ? +(x.stock / weeklyVelocity).toFixed(2) : null;
    const health = stockHealth(x.stock, weeklyVelocity);
    const vel = 1 - (x.medianDays ?? maxD) / maxD;
    const prof = maxP ? (x.avgProfit || 0) / maxP : 0;
    const vol = Math.log(1 + x.units) / Math.log(1 + maxV);
    const stockPenalty = x.stock ? 1 / (1 + x.stock) : 1.15;
    const buyScore = +((vel * 0.35 + prof * 0.35 + vol * 0.30) * stockPenalty).toFixed(3);
    return {
      ...x,
      weeklyVelocity: +weeklyVelocity.toFixed(2),
      weeksOfStock, health, buyScore,
    };
  });
}

const fmtMoney = (n) =>
  n == null ? "--" : "$" + Math.round(n).toLocaleString();
const fmtK = (n) =>
  n == null ? "--" : Math.abs(n) >= 1000 ? "$" + (n / 1000).toFixed(0) + "k" : "$" + Math.round(n);
const fmtPct = (n) => (n == null ? "--" : n.toFixed(1) + "%");
const median = (arr) => {
  const a = arr.filter((x) => x != null).sort((x, y) => x - y);
  if (!a.length) return null;
  const m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
};
const pctile = (arr, p) => {
  const a = arr.filter((x) => x != null).sort((x, y) => x - y);
  if (!a.length) return null;
  return a[Math.min(a.length - 1, Math.floor((p / 100) * a.length))];
};
/* normalize a reference number for matching (uppercase, strip punctuation except . and /) */
const normRef = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9./]/g, "").trim();

/* =========================================================================
   METRICS ENGINE
   ========================================================================= */
function computeMetrics(datasets, dateRange, includePresold, windowDays, invStatusFilter, excludedModels) {
  const inv = datasets.find((d) => d.role === "inventory");
  const sal = datasets.find((d) => d.role === "sales");
  const today = new Date();
  const out = { hasInv: !!inv, hasSales: !!sal };
  const hasExcl = excludedModels && excludedModels.size > 0;
  // accumulates the distinct watch models present (for the exclusion autocomplete)
  const modelOptionMap = {};
  const addModelOption = (brand, modelNumber, modelName) => {
    const key = modelKeyOf(modelNumber, modelName);
    if (!key) return;
    if (!modelOptionMap[key]) {
      const num = modelNumber != null ? String(modelNumber).trim() : "";
      const name = modelName != null ? String(modelName).trim() : "";
      const label = [brand, num, name].filter(Boolean).join(" · ");
      modelOptionMap[key] = { key, brand: brand || null, modelNumber: num || null, modelName: name || null, label, count: 0 };
    }
    modelOptionMap[key].count++;
  };

  /* ----- inventory ----- */
  if (inv) {
    const m = inv.mapping;
    let items = inv.rows.map((r, i) => {
      const cost = toNum(r[m.cost]);
      const tw = toNum(r[m.targetWholesale]);
      const tag = toNum(r[m.tagPrice]);
      const brandNorm = normalizeBrand(r[m.brand]);
      const pd = toDate(r[m.purchaseDate]);
      const age = pd ? daysBetween(pd, today) : null;
      const invType = (r[m.invType] && String(r[m.invType]).trim()) || null;
      addModelOption(r[m.brand], r[m.modelNumber], r[m.modelName]);
      return {
        _id: i,
        brand: r[m.brand] || "Unknown",
        brandNorm,
        line: findLine(brandNorm, r[m.modelName]),
        modelName: r[m.modelName],
        chartName: modelDisplayName(r[m.brand], r[m.modelName]),
        modelNumber: r[m.modelNumber] || null,
        modelKey: modelKeyOf(r[m.modelNumber], r[m.modelName]),
        cost: cost || 0,
        costMissing: cost == null || cost === 0,
        purchaseDate: pd,
        targetWholesale: tw,
        tagPrice: tag,
        age,
        grade: ageGrade(age),
        condition: normalizeCondition(r[m.condition]),
        status: r[m.status] || null,
        invType,
        invTypeLabel: invType || "Unspecified",
        paymentStatus: normalizePayment(r[m.paymentStatus]),
        supplier: (r[m.supplier] && String(r[m.supplier]).trim()) || null,
      };
    });
    // collect all statuses before filtering so the filter UI can show all options
    out.invStatuses = [...new Set(items.map((x) => x.status).filter(Boolean))].sort();
    // drop excluded models from every inventory total
    if (hasExcl) items = items.filter((x) => !excludedModels.has(x.modelKey));
    const filteredItems = (invStatusFilter && invStatusFilter.size)
      ? items.filter((x) => invStatusFilter.has(x.status))
      : items;

    out.invCount = filteredItems.length;
    out.invCost = filteredItems.reduce((s, x) => s + x.cost, 0);

    // fully-unknown items (no brand AND no name) are hidden from the breakdowns
    const namedItems = filteredItems.filter((x) => x.chartName != null);
    const byBrand = {};
    namedItems.forEach((x) => {
      const k = x.brand;
      byBrand[k] = byBrand[k] || { brand: k, count: 0, cost: 0 };
      byBrand[k].count++; byBrand[k].cost += x.cost;
    });
    out.invByBrand = Object.values(byBrand).sort((a, b) => b.cost - a.cost);
    out.invBrandCount = out.invByBrand.length;

    // by brand & condition (new / used)
    const bbc = {};
    namedItems.forEach((x) => {
      const cond = x.condition || "Unspecified";
      const k = x.brand + " · " + cond;
      bbc[k] = bbc[k] || { key: k, brand: x.brand, condition: cond, count: 0, cost: 0 };
      bbc[k].count++; bbc[k].cost += x.cost;
    });
    out.invByBrandCondition = Object.values(bbc)
      .sort((a, b) => a.brand.localeCompare(b.brand) || (conditionRank(a.condition) - conditionRank(b.condition)) || a.condition.localeCompare(b.condition));
    out.invHasCondition = filteredItems.some((x) => x.condition && x.condition !== "Unspecified");

    const byLine = {};
    namedItems.filter((x) => x.line).forEach((x) => {
      const lbl = lineLabel(x.brand, x.line);
      const k = x.brand + " · " + lbl;
      byLine[k] = byLine[k] || { key: k, brand: x.brand, line: lbl, count: 0, cost: 0 };
      byLine[k].count++; byLine[k].cost += x.cost;
    });
    out.invByLine = Object.values(byLine).sort((a, b) => b.cost - a.cost);

    const buckets = [
      { label: "0-30", lo: 0, hi: 30 }, { label: "31-60", lo: 31, hi: 60 },
      { label: "61-90", lo: 61, hi: 90 }, { label: "91-180", lo: 91, hi: 180 },
      { label: "180+", lo: 181, hi: 1e9 },
    ].map((b) => ({ ...b, count: 0, cost: 0, items: [] }));
    filteredItems.forEach((x) => {
      if (x.age == null) return;
      const b = buckets.find((q) => x.age >= q.lo && x.age <= q.hi);
      if (b) { b.count++; b.cost += x.cost; b.items.push(x); }
    });
    out.aging = buckets;
    out.agedValue = buckets.filter((b) => b.lo >= 91).reduce((s, b) => s + b.cost, 0);

    // ── inventory quality grade (A = fresh, F = aged) ──
    const byGrade = {};
    filteredItems.forEach((x) => {
      if (!x.grade) return;
      byGrade[x.grade] = byGrade[x.grade] || { grade: x.grade, count: 0, cost: 0, items: [] };
      byGrade[x.grade].count++; byGrade[x.grade].cost += x.cost; byGrade[x.grade].items.push(x);
    });
    out.invByGrade = GRADE_BUCKETS.map((b) => byGrade[b.grade] || { grade: b.grade, count: 0, cost: 0, items: [] });

    // ── watches to sell: worst grade first, then most cash tied up ──
    out.needToSellRanked = [...filteredItems]
      .filter((x) => x.grade)
      .sort((a, b) => (GRADE_RANK[b.grade] - GRADE_RANK[a.grade]) || (b.cost - a.cost));
    out.needToSell = out.needToSellRanked.slice(0, 10);

    // ── oldest 20 watches in stock (for market-value QoC / QoT grading) ──
    out.oldest20 = [...filteredItems]
      .filter((x) => x.age != null)
      .sort((a, b) => b.age - a.age)
      .slice(0, 20)
      .map((x) => ({
        brand: x.brand, model: x.chartName || x.modelName, reference: x.modelNumber,
        age: x.age, grade: x.grade, cost: x.cost, targetWholesale: x.targetWholesale,
      }));

    // ── data quality: how many items are missing each key field ──
    const dqDefs = [
      { key: "status", label: "Inventory status", missing: (x) => !x.status },
      { key: "invType", label: "Inventory type", missing: (x) => !x.invType },
      { key: "cost", label: "Purchase price", missing: (x) => x.costMissing },
      { key: "purchaseDate", label: "Purchase date", missing: (x) => x.purchaseDate == null },
      { key: "targetWholesale", label: "Wholesale target", missing: (x) => x.targetWholesale == null || x.targetWholesale <= 0 },
    ];
    out.dataQuality = dqDefs.map((d) => {
      const missingItems = filteredItems.filter(d.missing);
      return {
        key: d.key, field: d.label, missing: missingItems.length,
        present: filteredItems.length - missingItems.length, items: missingItems,
      };
    });

    // ── liabilities (computed on the FULL inventory, independent of status filter) ──
    // unpaid: payment marked unpaid, excluding memo items and voided items
    const unpaid = items.filter((x) => x.paymentStatus === "Unpaid" && !isMemoType(x.invType));
    out.unpaidItems = [...unpaid].sort((a, b) => b.cost - a.cost);
    out.unpaidLiability = unpaid.reduce((s, x) => s + x.cost, 0);
    out.unpaidCount = unpaid.length;

    // consignment: unsold consignment stock (everything in inventory is unsold), not voided
    const consign = items.filter((x) => isConsignmentType(x.invType) && x.paymentStatus !== "Voided");
    out.consignmentItems = [...consign].sort((a, b) => b.cost - a.cost);
    out.consignmentLiability = consign.reduce((s, x) => s + x.cost, 0);
    out.consignmentCount = consign.length;

    out.hasPaymentStatus = items.some((x) => x.paymentStatus && x.paymentStatus !== "Unspecified");
    out.hasInvType = items.some((x) => x.invType);

    // per-model & per-brand inventory aggregates for the projected-vs-historical join.
    // projected profit/margin is computed only on the priced (target-wholesale) portion.
    const invModel = {}, invBrand = {};
    namedItems.forEach((x) => {
      const brand = (x.brand && String(x.brand).trim()) || "Unknown";
      const priced = x.targetWholesale && x.targetWholesale > 0;
      const mk = x.modelKey || (brand + "|" + x.chartName);
      const bk = brand;
      if (!invModel[mk]) invModel[mk] = { key: mk, brand, model: x.chartName, stock: 0, pricedCost: 0, projProfit: 0 };
      if (!invBrand[bk]) invBrand[bk] = { key: bk, brand, stock: 0, pricedCost: 0, projProfit: 0 };
      for (const agg of [invModel[mk], invBrand[bk]]) {
        agg.stock++;
        if (priced) { agg.pricedCost += x.cost; agg.projProfit += (x.targetWholesale - x.cost); }
      }
    });
    out._invModelAgg = invModel;
    out._invBrandAgg = invBrand;

    // per-reference inventory position (stock + cost history) for the Alerts engine
    const refInv = {};
    filteredItems.forEach((x) => {
      const rf = normRef(x.modelNumber);
      if (!rf) return;
      refInv[rf] = refInv[rf] || { stock: 0, costs: [], brand: x.brand, model: x.chartName || x.modelName };
      refInv[rf].stock++;
      if (x.cost) refInv[rf].costs.push(x.cost);
    });
    out._refInv = refInv;

    // supplier ("purchased from") spend, from inventory purchase cost
    out._invSupplierAgg = {};
    filteredItems.forEach((x) => {
      if (!x.supplier) return;
      const k = x.supplier;
      out._invSupplierAgg[k] = out._invSupplierAgg[k] || { supplier: k, units: 0, spend: 0, profit: null };
      out._invSupplierAgg[k].units++; out._invSupplierAgg[k].spend += x.cost;
    });

    const withTarget = filteredItems.filter((x) => x.targetWholesale && x.targetWholesale > 0);
    out.projItems = withTarget.length;
    out.projProfit = withTarget.reduce((s, x) => s + (x.targetWholesale - x.cost), 0);
    const pbb = {};
    withTarget.forEach((x) => {
      pbb[x.brand] = pbb[x.brand] || { brand: x.brand, profit: 0, count: 0 };
      pbb[x.brand].profit += x.targetWholesale - x.cost; pbb[x.brand].count++;
    });
    out.projByBrand = Object.values(pbb).sort((a, b) => b.profit - a.profit);

    // projected profit by model
    const pbm = {};
    withTarget.forEach((x) => {
      const key = (x.brand || "Unknown") + "|" + (x.modelName || "Unknown");
      pbm[key] = pbm[key] || { brand: x.brand, model: x.modelName || "Unknown", profit: 0, cost: 0, tw: 0, count: 0 };
      pbm[key].profit += x.targetWholesale - x.cost;
      pbm[key].cost += x.cost;
      pbm[key].tw += x.targetWholesale;
      pbm[key].count++;
    });
    out.projByModel = Object.values(pbm).map((b) => ({
      ...b,
      marginPct: b.cost ? (b.profit / b.cost) * 100 : null,
    })).sort((a, b) => b.profit - a.profit);

    // current stock count per brand+model (for "are we out" once sales has model)
    const stock = {};
    filteredItems.forEach((x) => {
      const key = (x.brandNorm || "") + "|" + String(x.modelName || "").toLowerCase();
      stock[key] = (stock[key] || 0) + 1;
    });
    out._stock = stock;

    // current stock count per brand+product line
    const stockByLine = {};
    filteredItems.filter((x) => x.line).forEach((x) => {
      const lbl = lineLabel(x.brand, x.line);
      const key = (x.brandNorm || "") + "|" + lbl.toLowerCase();
      stockByLine[key] = (stockByLine[key] || 0) + 1;
    });
    out._stockByLine = stockByLine;
  }

  /* ----- sales ----- */
  if (sal) {
    const m = sal.mapping;
    let rows = sal.rows.map((r) => {
      const cost = toNum(r[m.cost]);
      const price = toNum(r[m.salePrice]);
      let profit = toNum(r[m.profit]);
      if (profit == null && price != null && cost != null) profit = price - cost;
      const pd = toDate(r[m.purchaseDate]);
      const sdt = toDate(r[m.saleDate]);
      let days = daysBetween(pd, sdt);
      if (days != null && (days < 0 || days > 2000)) days = null; // typo guard
      // some reference numbers are excluded from velocity/median math: drop their
      // days-to-sell so they never feed median/velocity, but keep the unit counted
      const velIgnored = isVelocityIgnored(r[m.modelNumber], r[m.modelName]);
      if (velIgnored) days = null;
      const presold = days != null && days >= 0 && days <= 2;
      const brandNorm = normalizeBrand(r[m.brand]);
      addModelOption(r[m.brand], r[m.modelNumber], r[m.modelName]);
      return {
        saleDate: sdt, days, presold, velIgnored, cost: cost || 0, price: price || 0,
        profit: profit || 0,
        marginPct: cost ? (profit / cost) * 100 : null,
        brand: r[m.brand] || null, brandNorm,
        line: findLine(brandNorm, r[m.modelName]),
        modelName: r[m.modelName] || null,
        chartName: modelDisplayName(r[m.brand], r[m.modelName]),
        modelNumber: r[m.modelNumber] || null,
        modelKey: modelKeyOf(r[m.modelNumber], r[m.modelName]),
        type: r[m.inventoryType] || "Unspecified",
        condition: normalizeCondition(r[m.condition]),
        priceTier: priceTierLabel(price),
        shippingState: (r[m.shippingState] && String(r[m.shippingState]).trim()) || "Unspecified",
        supplier: (r[m.supplier] && String(r[m.supplier]).trim()) || null,
        salesperson: (r[m.salesperson] && String(r[m.salesperson]).trim()) || null,
        amountOwed: toNum(r[m.amountOwed]),
        customer: (r[m.customer] && String(r[m.customer]).trim()) || null,
      };
    });
    // full mapped set (before the period window) — liabilities reflect all invoices
    const fullSalesRows = rows;

    // ── latest 20 sales (for market-value sales grading) ──
    out.latest20 = [...fullSalesRows]
      .filter((x) => x.saleDate)
      .sort((a, b) => b.saleDate - a.saleDate)
      .slice(0, 20)
      .map((x) => ({
        brand: x.brand, model: x.chartName || x.modelName, reference: x.modelNumber,
        saleDate: x.saleDate, price: x.price, cost: x.cost, profit: x.profit,
      }));

    // all-time historical aggregates by model & brand (for projected-vs-historical health)
    const histModel = {}, histBrand = {};
    fullSalesRows.filter((x) => x.chartName != null).forEach((x) => {
      const brand = (x.brand && String(x.brand).trim()) || "Unknown";
      const mk = x.modelKey || (brand + "|" + x.chartName);
      if (!histModel[mk]) histModel[mk] = { key: mk, brand, model: x.chartName, sold: 0, profit: 0, cost: 0 };
      if (!histBrand[brand]) histBrand[brand] = { key: brand, brand, sold: 0, profit: 0, cost: 0 };
      for (const agg of [histModel[mk], histBrand[brand]]) {
        agg.sold++; agg.profit += x.profit; agg.cost += x.cost;
      }
    });
    out._histModelAgg = histModel;
    out._histBrandAgg = histBrand;

    // ── overpaid invoices (negative amount owed) — a liability regardless of period ──
    out.hasAmountOwed = m.amountOwed != null && fullSalesRows.some((x) => x.amountOwed != null);
    const overpaid = fullSalesRows.filter((x) => x.amountOwed != null && x.amountOwed < 0);
    out.overpaidItems = overpaid.map((x) => ({
      customer: x.customer, brand: x.brand, model: x.chartName || x.modelName || x.modelNumber,
      saleDate: x.saleDate, overage: -x.amountOwed, price: x.price,
    })).sort((a, b) => b.overage - a.overage);
    out.overpaidLiability = overpaid.reduce((s, x) => s + (-x.amountOwed), 0);
    out.overpaidCount = overpaid.length;
    // drop excluded models from every sales total
    if (hasExcl) rows = rows.filter((x) => !excludedModels.has(x.modelKey));

    // overall date span of the data (used for the date-range picker bounds) — full, pre-window
    const allSaleDates = rows.map((x) => x.saleDate).filter(Boolean);
    out.salesDateMin = allSaleDates.length ? new Date(Math.min(...allSaleDates.map((d) => d.getTime()))) : null;
    out.salesDateMax = allSaleDates.length ? new Date(Math.max(...allSaleDates.map((d) => d.getTime()))) : null;

    // Sales period — applies to the ENTIRE sales analysis (KPIs, breakdowns, Buy
    // Signals, Salespeople, Tax) so every section reflects the same timeframe.
    // Precedence: explicit custom date range > "last N days from today" > all time.
    out.salesWindowDays = windowDays || null;
    if (dateRange && (dateRange.start || dateRange.end)) {
      const startT = dateRange.start ? new Date(dateRange.start + "T00:00:00").getTime() : -Infinity;
      const endT = dateRange.end ? new Date(dateRange.end + "T23:59:59").getTime() : Infinity;
      rows = rows.filter((x) => x.saleDate && x.saleDate.getTime() >= startT && x.saleDate.getTime() <= endT);
      out.windowMode = "custom";
    } else if (windowDays) {
      const windowStartT = today.getTime() - windowDays * 86400000;
      out.windowStart = new Date(windowStartT);
      rows = rows.filter((x) => x.saleDate && x.saleDate.getTime() >= windowStartT);
      out.windowMode = "window";
    } else {
      out.windowMode = "all";
    }
    // every downstream total (incl. tax & salesperson) uses the same windowed set
    const allSalesRows = rows;

    // "presold" = sold within 0-2 days of purchase — likely flipped before it ever
    // hit the floor, so it's excluded from sell-through analysis by default.
    out.presoldCount = rows.filter((x) => x.presold).length;
    if (!includePresold) {
      rows = rows.filter((x) => !x.presold);
    }

    out.salesUnits = rows.length;
    out.salesProfit = rows.reduce((s, x) => s + x.profit, 0);
    out.salesRevenue = rows.reduce((s, x) => s + x.price, 0);
    out.salesCOGS = rows.reduce((s, x) => s + x.cost, 0);
    out.salesProfitPct = out.salesCOGS ? (out.salesProfit / out.salesCOGS) * 100 : null;
    out.medianMargin = median(rows.map((x) => x.marginPct));
    out.medianDays = median(rows.map((x) => x.days));
    out.meanDays = (() => {
      const a = rows.map((x) => x.days).filter((x) => x != null);
      return a.length ? a.reduce((s, x) => s + x, 0) / a.length : null;
    })();

    // span of the sales data, in weeks — used for sell-through velocity.
    // when a "last N days" window is active we use N/7 as the denominator so
    // sparse or single-sale windows don't overstate velocity.
    out.salesWeeks = (() => {
      if (out.windowMode === "window") return windowDays / 7;
      const ds = rows.map((x) => x.saleDate).filter(Boolean).map((d) => d.getTime());
      if (ds.length < 2) return 1;
      const span = (Math.max(...ds) - Math.min(...ds)) / (7 * 86400000);
      return Math.max(span, 1);
    })();

    // per-reference sales position for the Alerts engine: all-time cost & sale-price
    // history, plus recent (windowed) units for a velocity / weeks-of-stock estimate
    const refSal = {};
    fullSalesRows.forEach((x) => {
      const rf = normRef(x.modelNumber);
      if (!rf) return;
      refSal[rf] = refSal[rf] || { costs: [], sales: [], recent: 0, lastSale: null, brand: x.brand, model: x.chartName || x.modelName };
      if (x.cost) refSal[rf].costs.push(x.cost);
      if (x.price) refSal[rf].sales.push(x.price);
      if (x.saleDate && (!refSal[rf].lastSale || x.saleDate > refSal[rf].lastSale)) refSal[rf].lastSale = x.saleDate;
    });
    rows.forEach((x) => { const rf = normRef(x.modelNumber); if (rf && refSal[rf]) refSal[rf].recent++; });
    out._refSal = refSal;

    // monthly
    const mo = {};
    rows.forEach((x) => {
      if (!x.saleDate) return;
      const k = x.saleDate.getFullYear() + "-" + String(x.saleDate.getMonth() + 1).padStart(2, "0");
      mo[k] = mo[k] || { month: k, units: 0, profit: 0 };
      mo[k].units++; mo[k].profit += x.profit;
    });
    out.monthly = Object.values(mo).sort((a, b) => a.month.localeCompare(b.month));

    // by salesperson ("created by") — profit, sales, velocity (median days to sell).
    // respects the active sales period like every other section.
    const bperson = {};
    allSalesRows.forEach((x) => {
      if (!x.salesperson) return;
      const k = x.salesperson;
      bperson[k] = bperson[k] || { salesperson: k, units: 0, profit: 0, revenue: 0, cost: 0, _days: [], _m: [], _pts: [], sales: [] };
      bperson[k].units++; bperson[k].profit += x.profit; bperson[k].revenue += x.price; bperson[k].cost += x.cost;
      if (x.days != null) bperson[k]._days.push(x.days);
      if (x.marginPct != null) bperson[k]._m.push(x.marginPct);
      const dj = djSaleGrade({ marginPct: x.marginPct, days: x.days, profit: x.profit });
      if (dj.points != null) bperson[k]._pts.push(dj.points);
      bperson[k].sales.push({
        saleDate: x.saleDate, brand: x.brand, model: x.chartName || x.modelName, reference: x.modelNumber,
        price: x.price, cost: x.cost, profit: x.profit, marginPct: x.marginPct, days: x.days,
        djGrade: dj.grade, djPoints: dj.points,
      });
    });
    out.salesByPerson = Object.values(bperson).map((b) => {
      const avgPts = b._pts.length ? b._pts.reduce((s, v) => s + v, 0) / b._pts.length : null;
      return {
        salesperson: b.salesperson, units: b.units, profit: b.profit, revenue: b.revenue,
        avgProfit: b.units ? b.profit / b.units : null,
        profitPct: b.cost ? (b.profit / b.cost) * 100 : null,
        medianDays: median(b._days), medianMargin: median(b._m),
        djGrade: ptsToGrade(avgPts), djPoints: avgPts != null ? +avgPts.toFixed(2) : null,
        sales: b.sales.sort((a, c) => (c.saleDate || 0) - (a.saleDate || 0)),
      };
    }).sort((a, b) => b.profit - a.profit);
    out.hasSalesperson = allSalesRows.some((x) => x.salesperson);

    // by type (always available)
    const bt = {};
    rows.forEach((x) => {
      bt[x.type] = bt[x.type] || { type: x.type, units: 0, profit: 0 };
      bt[x.type].units++; bt[x.type].profit += x.profit;
    });
    out.salesByType = Object.values(bt).sort((a, b) => b.units - a.units);

    // by price tier
    const bp = {};
    rows.forEach((x) => {
      const tier = x.priceTier || "Unspecified";
      bp[tier] = bp[tier] || { tier, units: 0, profit: 0, revenue: 0, cost: 0, _m: [] };
      bp[tier].units++; bp[tier].profit += x.profit; bp[tier].revenue += x.price; bp[tier].cost += x.cost;
      if (x.marginPct != null) bp[tier]._m.push(x.marginPct);
    });
    out.salesByPriceTier = PRICE_TIERS.map((t) => bp[t.label] || { tier: t.label, units: 0, profit: 0, revenue: 0, cost: 0, _m: [] })
      .map((b) => ({
        tier: b.tier, units: b.units, profit: b.profit, revenue: b.revenue,
        avgProfit: b.units ? b.profit / b.units : null,
        profitPct: b.cost ? (b.profit / b.cost) * 100 : null,
        medianMargin: median(b._m),
      }));

    // ── sales tax by shipping state — a liability, so ALL invoices (ignores the period) ──
    const bst = {};
    fullSalesRows.forEach((x) => {
      const st = x.shippingState || "Unspecified";
      bst[st] = bst[st] || { state: st, units: 0, revenue: 0 };
      bst[st].units++; bst[st].revenue += x.price;
    });
    // apply the illustrative per-state tax rules (tax on revenue above threshold)
    const taxed = Object.values(bst).map((s) => {
      const rule = STATE_TAX_RULES[s.state];
      const taxable = rule ? Math.max(0, s.revenue - rule.threshold) : 0;
      const tax = rule ? taxable * rule.rate : 0;
      return {
        ...s,
        taxRate: rule ? rule.rate : null,
        taxThreshold: rule ? rule.threshold : null,
        taxableBase: taxable,
        tax,
      };
    });
    // only the states that have a tax rule (top 4 by tax owed)
    out.salesTaxByState = taxed
      .filter((s) => STATE_TAX_RULES[s.state])
      .sort((a, b) => b.tax - a.tax || b.revenue - a.revenue)
      .slice(0, 4);
    out.salesTaxTotal = taxed.reduce((s, x) => s + x.tax, 0);
    out.hasShippingState = fullSalesRows.some((x) => x.shippingState && x.shippingState !== "Unspecified");

    // supplier ("purchased from") spend + profit, from sales rows
    out._salSupplierAgg = {};
    rows.forEach((x) => {
      if (!x.supplier) return;
      const k = x.supplier;
      out._salSupplierAgg[k] = out._salSupplierAgg[k] || { supplier: k, units: 0, spend: 0, profit: 0 };
      out._salSupplierAgg[k].units++; out._salSupplierAgg[k].spend += x.cost; out._salSupplierAgg[k].profit += x.profit;
    });

    // brand-level (only if brand present on sales)
    out.salesHasBrand = rows.some((x) => x.brand);
    // fully-unknown watches (no brand AND no name) are hidden from every breakdown
    const namedRows = rows.filter((x) => x.chartName != null);
    if (out.salesHasBrand) {
      const bb = {};
      namedRows.forEach((x) => {
        const k = x.brand || "Unknown";
        bb[k] = bb[k] || { brand: k, units: 0, profit: 0, revenue: 0, cost: 0, _days: [], _m: [] };
        bb[k].units++; bb[k].profit += x.profit; bb[k].revenue += x.price; bb[k].cost += x.cost;
        if (x.days != null) bb[k]._days.push(x.days);
        if (x.marginPct != null) bb[k]._m.push(x.marginPct);
      });
      out.salesByBrand = Object.values(bb).map((b) => ({
        brand: b.brand, units: b.units, profit: b.profit, revenue: b.revenue,
        profitPct: b.cost ? (b.profit / b.cost) * 100 : null,
        medianDays: median(b._days), medianMargin: median(b._m),
      })).sort((a, b) => b.profit - a.profit);

      // by brand & condition (new / used)
      const bc = {};
      namedRows.forEach((x) => {
        const brand = x.brand || "Unknown";
        const cond = x.condition || "Unspecified";
        const k = brand + " · " + cond;
        bc[k] = bc[k] || { key: k, brand, condition: cond, units: 0, profit: 0, revenue: 0, cost: 0, _m: [] };
        bc[k].units++; bc[k].profit += x.profit; bc[k].revenue += x.price; bc[k].cost += x.cost;
        if (x.marginPct != null) bc[k]._m.push(x.marginPct);
      });
      out.salesByBrandCondition = Object.values(bc).map((b) => ({
        key: b.key, brand: b.brand, condition: b.condition, units: b.units, profit: b.profit, revenue: b.revenue,
        profitPct: b.cost ? (b.profit / b.cost) * 100 : null,
        medianMargin: median(b._m),
      })).sort((a, b) => a.brand.localeCompare(b.brand) || (conditionRank(a.condition) - conditionRank(b.condition)) || a.condition.localeCompare(b.condition));
      out.hasCondition = rows.some((x) => x.condition && x.condition !== "Unspecified");

      // by product line
      const bl = {};
      namedRows.forEach((x) => {
        if (!x.line) return;
        const lbl = lineLabel(x.brand, x.line);
        const k = (x.brand || "Unknown") + " · " + lbl;
        bl[k] = bl[k] || { key: k, brand: x.brand, brandNorm: x.brandNorm, line: lbl, units: 0, profit: 0, cost: 0, _days: [], _m: [] };
        bl[k].units++; bl[k].profit += x.profit; bl[k].cost += x.cost;
        if (x.days != null) bl[k]._days.push(x.days);
        if (x.marginPct != null) bl[k]._m.push(x.marginPct);
      });
      let lines = Object.values(bl).map((b) => ({
        key: b.key, brand: b.brand, line: b.line, units: b.units, profit: b.profit,
        avgProfit: b.profit / b.units,
        profitPct: b.cost ? (b.profit / b.cost) * 100 : null,
        medianDays: median(b._days), medianMargin: median(b._m),
        stock: out._stockByLine ? (out._stockByLine[(b.brandNorm || "") + "|" + b.line.toLowerCase()] ?? null) : null,
      }));
      out.salesByLine = enrichRanking(lines, out.salesWeeks).sort((a, b) => b.profit - a.profit);

      // by model frequency + ranking + buy score
      const bm = {};
      namedRows.forEach((x) => {
        const name = x.chartName;
        const k = (x.brand || "") + "|" + name;
        bm[k] = bm[k] || { key: k, brand: x.brand, line: lineLabel(x.brand, x.line), model: name, brandNorm: x.brandNorm, units: 0, profit: 0, revenue: 0, cost: 0, _days: [], _m: [] };
        bm[k].units++; bm[k].profit += x.profit; bm[k].revenue += x.price; bm[k].cost += x.cost;
        if (x.days != null) bm[k]._days.push(x.days);
        if (x.marginPct != null) bm[k]._m.push(x.marginPct);
      });
      let models = Object.values(bm).map((b) => ({
        brand: b.brand, line: b.line, model: b.model, units: b.units, profit: b.profit, revenue: b.revenue,
        avgProfit: b.profit / b.units,
        avgCost: b.units ? b.cost / b.units : null,
        profitPct: b.cost ? (b.profit / b.cost) * 100 : null,
        medianDays: median(b._days), medianMargin: median(b._m),
        stock: out._stock ? (out._stock[(b.brandNorm || "") + "|" + String(b.model || "").toLowerCase()] ?? null) : null,
      }));
      models = enrichRanking(models, out.salesWeeks);
      out.salesByModel = [...models].sort((a, b) => b.units - a.units);
      out.byVelocity = [...models]
        .filter((x) => x.medianDays != null)
        .sort((a, b) => a.medianDays - b.medianDays);

      // buy signals / "most profitable" should ignore models that have only
      // sold once or twice — not enough history to trust the ranking
      const modelsForBuy = models.filter((x) => x.units >= 3);
      out.ranking = [...modelsForBuy].sort((a, b) => b.buyScore - a.buyScore);

      // buy-signals hierarchy: brand → product line → model, ranked by summed buy score
      const bh = {};
      modelsForBuy.forEach((mo) => {
        const brand = mo.brand || "Unknown";
        const line = mo.line || (brand + " — Other");
        bh[brand] = bh[brand] || { brand, score: 0, units: 0, profit: 0, projCost: 0, _lines: {} };
        const B = bh[brand];
        B.score += mo.buyScore || 0; B.units += mo.units; B.profit += mo.profit; B.projCost += (mo.avgCost || 0);
        B._lines[line] = B._lines[line] || { line, brand, score: 0, units: 0, profit: 0, projCost: 0, models: [] };
        const L = B._lines[line];
        L.score += mo.buyScore || 0; L.units += mo.units; L.profit += mo.profit; L.projCost += (mo.avgCost || 0);
        L.models.push(mo);
      });
      out.buyHierarchy = Object.values(bh).map((B) => ({
        brand: B.brand, score: +B.score.toFixed(3), units: B.units, profit: B.profit, projCost: B.projCost,
        lines: Object.values(B._lines).map((L) => ({
          line: L.line, brand: L.brand, score: +L.score.toFixed(3), units: L.units, profit: L.profit, projCost: L.projCost,
          models: [...L.models].sort((a, b) => b.buyScore - a.buyScore),
        })).sort((a, b) => b.score - a.score),
      })).sort((a, b) => b.score - a.score);

      // combined velocity × profit ranking (for Q3)
      const maxP2 = Math.max(...modelsForBuy.map((x) => x.avgProfit || 0)) || 1;
      const maxD2 = Math.max(...modelsForBuy.map((x) => x.medianDays || 0)) || 1;
      out.velProfRanking = [...modelsForBuy].sort((a, b) => {
        const scoreA = (1 - (a.medianDays ?? maxD2) / maxD2) * 0.5 + (maxP2 ? (a.avgProfit || 0) / maxP2 : 0) * 0.5;
        const scoreB = (1 - (b.medianDays ?? maxD2) / maxD2) * 0.5 + (maxP2 ? (b.avgProfit || 0) / maxP2 : 0) * 0.5;
        return scoreB - scoreA;
      });

      // ── Top-10 lists (model & product-line granularity) ──
      const top = (arr, n, sortFn) => [...arr].sort(sortFn).slice(0, n);
      out.fastestModels = top(models.filter((x) => x.medianDays != null), 10, (a, b) => a.medianDays - b.medianDays);
      out.fastestLines = top(out.salesByLine.filter((x) => x.medianDays != null), 10, (a, b) => a.medianDays - b.medianDays);
      // full (un-sliced) fastest lists so the UI can refill after an ignore
      out.fastestModelsAll = [...models].filter((x) => x.medianDays != null).sort((a, b) => a.medianDays - b.medianDays);
      out.fastestLinesAll = [...out.salesByLine].filter((x) => x.medianDays != null).sort((a, b) => a.medianDays - b.medianDays);
      out.bestProfitModels = top(modelsForBuy, 10, (a, b) => b.profit - a.profit);
      out.bestProfitLines = top(out.salesByLine, 10, (a, b) => b.profit - a.profit);
      out.bestScoreModels = top(modelsForBuy, 10, (a, b) => b.buyScore - a.buyScore);
      out.bestScoreLines = top(out.salesByLine, 10, (a, b) => b.buyScore - a.buyScore);

      // ── Stock health lists ──
      const healthRank = { red: 0, yellow: 1, green: 2 };
      out.healthByVelocityModels = modelsForBuy.filter((x) => x.health)
        .sort((a, b) => (a.weeksOfStock ?? 1e9) - (b.weeksOfStock ?? 1e9));
      out.healthByVelocityLines = out.salesByLine.filter((x) => x.health)
        .sort((a, b) => (a.weeksOfStock ?? 1e9) - (b.weeksOfStock ?? 1e9));
      out.healthByScoreModels = modelsForBuy.filter((x) => x.health)
        .sort((a, b) => (healthRank[a.health] - healthRank[b.health]) || (b.buyScore - a.buyScore));
      out.healthByScoreLines = out.salesByLine.filter((x) => x.health)
        .sort((a, b) => (healthRank[a.health] - healthRank[b.health]) || (b.buyScore - a.buyScore));
    }
  }

  // ── top suppliers ("purchased from") — spend from inventory, profit from sales ──
  const invSup = out._invSupplierAgg || {};
  const salSup = out._salSupplierAgg || {};
  // spend: prefer the inventory purchase record; fall back to sales cost
  const spendAgg = Object.keys(invSup).length ? invSup : salSup;
  out.suppliersBySpend = Object.values(spendAgg).sort((a, b) => b.spend - a.spend).slice(0, 10);
  out.hasSupplierSpend = out.suppliersBySpend.length > 0;
  // profit only available where sales carry the supplier column
  out.suppliersByProfit = Object.values(salSup).sort((a, b) => (b.profit || 0) - (a.profit || 0)).slice(0, 10);
  out.hasSupplierProfit = out.suppliersByProfit.length > 0;
  out.hasSupplier = out.hasSupplierSpend || out.hasSupplierProfit;
  delete out._invSupplierAgg; delete out._salSupplierAgg;

  // distinct models present in the data, for the exclusion autocomplete
  out.modelOptions = Object.values(modelOptionMap).sort((a, b) => a.label.localeCompare(b.label));

  // ── projected profit (current stock) vs historical profit (all past sales) ──
  const joinProjHist = (invAgg, histAgg) => {
    invAgg = invAgg || {}; histAgg = histAgg || {};
    const keys = new Set([...Object.keys(invAgg), ...Object.keys(histAgg)]);
    return [...keys].map((k) => {
      const iv = invAgg[k], hv = histAgg[k];
      const projMargin = iv && iv.pricedCost ? (iv.projProfit / iv.pricedCost) * 100 : null;
      const histMargin = hv && hv.cost ? (hv.profit / hv.cost) * 100 : null;
      return {
        brand: (iv && iv.brand) || (hv && hv.brand),
        model: (iv && iv.model) || (hv && hv.model),
        stock: (iv && iv.stock) || 0,
        sold: (hv && hv.sold) || 0,
        projProfit: (iv && iv.projProfit) || 0,
        histProfit: (hv && hv.profit) || 0,
        projMargin, histMargin,
        delta: (projMargin != null && histMargin != null) ? projMargin - histMargin : null,
      };
    }).sort((a, b) => b.projProfit - a.projProfit);
  };
  out.projVsHistModels = joinProjHist(out._invModelAgg, out._histModelAgg);
  out.projVsHistBrands = joinProjHist(out._invBrandAgg, out._histBrandAgg);
  out.hasProjVsHist = (out.projVsHistModels.length > 0) && (out.hasInv && out.hasSales);
  delete out._invModelAgg; delete out._invBrandAgg; delete out._histModelAgg; delete out._histBrandAgg;

  // ── reference positions (for the Alerts engine): merge inventory + sales by ref ──
  const refPos = {};
  const seed = (rf, o) => {
    refPos[rf] = refPos[rf] || { ref: rf, stock: 0, costs: [], sales: [], recent: 0, lastSale: null, brand: null, model: null };
    if (o.brand && !refPos[rf].brand) refPos[rf].brand = o.brand;
    if (o.model && !refPos[rf].model) refPos[rf].model = o.model;
    return refPos[rf];
  };
  Object.entries(out._refInv || {}).forEach(([rf, v]) => { const p = seed(rf, v); p.stock += v.stock; p.costs.push(...v.costs); });
  Object.entries(out._refSal || {}).forEach(([rf, v]) => { const p = seed(rf, v); p.costs.push(...v.costs); p.sales.push(...v.sales); p.recent += v.recent; if (v.lastSale && (!p.lastSale || v.lastSale > p.lastSale)) p.lastSale = v.lastSale; });
  const wks = out.salesWeeks || (45 / 7);
  out.refPositions = {};
  for (const rf in refPos) {
    const p = refPos[rf];
    const wv = p.recent / wks;
    out.refPositions[rf] = {
      ref: rf, brand: p.brand, model: p.model, stock: p.stock,
      weeklyVelocity: +wv.toFixed(2),
      weeksOfStock: wv > 0 ? +(p.stock / wv).toFixed(1) : null,
      medianCost: median(p.costs), loCost: pctile(p.costs, 10), hiCost: pctile(p.costs, 90),
      maxCost: p.costs.length ? Math.max(...p.costs) : null,
      medianSale: median(p.sales),
      lastSale: p.lastSale,
    };
  }
  delete out._refInv; delete out._refSal;

  // ── WhatsApp dealer-group messages (buy/sell intents) for the Alerts engine ──
  // supports several daily files dropped together
  const msgSets = datasets.filter((d) => d.role === "messages");
  const parsed = [];
  // build candidate reference keys, PREFERRING the structured fullReferenceNumber
  // over the loose `model` field (which the upstream parser sometimes fills with a
  // price), then resolve to whichever candidate we actually trade.
  const stripVariant = (s) => String(s || "").trim().replace(/-\d{3,4}$/, ""); // drop "-0013" style variant suffix
  const refCandidates = (rawRef, rawModel, rawBrand) => {
    // fullReferenceNumber is authoritative — when present, do NOT fall back to the
    // loose `model` field (it sometimes holds the price, e.g. "15200")
    const c = rawRef
      ? [normRef(stripVariant(rawRef)), normRef(rawRef)]
      : [normRef(rawModel), normRef(rawBrand)];
    return c.filter((x) => x && x.length >= 4);
  };
  msgSets.forEach((msg) => {
    const mm = msg.mapping;
    msg.rows.forEach((r) => {
      const intent = String(r[mm.intent] || "").toLowerCase().trim();
      if (intent !== "buy" && intent !== "sell") return;
      const rawModel = r[mm.model], rawBrand = r[mm.brand], rawRef = mm.reference ? r[mm.reference] : null;
      const cands = refCandidates(rawRef, rawModel, rawBrand);
      if (!cands.length) return;
      const ref = cands.find((c) => out.refPositions[c]) || cands[0]; // prefer one we trade
      parsed.push({
        intent, price: toNum(r[mm.price]), ts: toDate(r[mm.timestamp]),
        body: r[mm.messageBody] || null, sender: r[mm.sender] || null, chat: r[mm.chat] || null,
        ref, rawModel: rawModel || rawBrand || null,
      });
    });
  });
  out.messages = parsed;
  out.hasMessages = parsed.length > 0;
  const ts = parsed.map((x) => x.ts).filter(Boolean).map((d) => d.getTime());
  out.messagesMax = ts.length ? new Date(Math.max(...ts)) : null;

  return out;
}

/* =========================================================================
   UI PRIMITIVES
   ========================================================================= */
function Stat({ label, value, sub }) {
  return (
    <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 14 }}
      className="p-4 flex flex-col gap-1">
      <div style={{ color: C.dim, fontFamily: SANS, letterSpacing: ".06em" }}
        className="text-xs uppercase">{label}</div>
      <div style={{ color: C.text, fontFamily: SERIF }} className="text-2xl">{value}</div>
      {sub && <div style={{ color: C.faint, fontFamily: SANS }} className="text-xs">{sub}</div>}
    </div>
  );
}
/* side-by-side "with vs without exclusions" comparison of key totals */
function ExclusionSummary({ rows, excludedCount, scope }) {
  return (
    <div style={{ background: C.panel, border: `1px solid ${C.gold}`, borderRadius: 14 }} className="p-4">
      <div className="flex items-baseline justify-between mb-3 flex-wrap gap-2">
        <h3 style={{ color: C.gold, fontFamily: SERIF }} className="text-lg">
          With vs without {excludedCount} excluded {excludedCount === 1 ? "model" : "models"}
        </h3>
        <span style={{ color: C.faint, fontFamily: SANS }} className="text-xs">
          {scope === "dashboard" ? "exclusion applied dashboard-wide" : "exclusion applied to totals only"}
        </span>
      </div>
      <div className="overflow-auto">
        <table className="w-full" style={{ fontFamily: SANS, fontSize: 13 }}>
          <thead>
            <tr style={{ color: C.faint }}>
              <th className="text-left py-2 pr-4 font-normal text-xs uppercase tracking-wide">Metric</th>
              <th className="text-right py-2 pr-6 font-normal text-xs uppercase tracking-wide">Full</th>
              <th className="text-right py-2 pr-6 font-normal text-xs uppercase tracking-wide">Adjusted</th>
              <th className="text-right py-2 font-normal text-xs uppercase tracking-wide">Difference</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => {
              const diff = (r.full ?? 0) - (r.adj ?? 0);
              const fmt = r.fmt || ((v) => v);
              return (
                <tr key={r.label} style={{ borderTop: `1px solid ${C.line}`, color: C.text }}>
                  <td className="py-2 pr-4">{r.label}</td>
                  <td className="py-2 pr-6 text-right">{fmt(r.full)}</td>
                  <td className="py-2 pr-6 text-right" style={{ color: C.gold }}>{fmt(r.adj)}</td>
                  <td className="py-2 text-right" style={{ color: Math.abs(diff) < 1e-9 ? C.faint : C.dim }}>
                    {Math.abs(diff) < 1e-9 ? "—" : "−" + fmt(Math.abs(diff))}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
function Panel({ title, children, note }) {
  return (
    <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 14 }} className="p-4">
      <div className="flex items-baseline justify-between mb-3">
        <h3 style={{ color: C.text, fontFamily: SERIF }} className="text-lg">{title}</h3>
        {note && <span style={{ color: C.faint, fontFamily: SANS }} className="text-xs">{note}</span>}
      </div>
      {children}
    </div>
  );
}
function Locked({ msg }) {
  return (
    <div style={{ color: C.faint, fontFamily: SANS, border: `1px dashed ${C.line}`, borderRadius: 12 }}
      className="p-6 flex items-center gap-3 text-sm">
      <Lock size={18} /> <span>{msg}</span>
    </div>
  );
}
/* secondary tab row, sits under the main Inventory / Sales / Buy tabs */
function SubTabs({ tabs, value, onChange }) {
  return (
    <div className="flex gap-1 mb-4 flex-wrap" style={{ borderBottom: `1px solid ${C.line}`, paddingBottom: 8 }}>
      {tabs.map(([k, label]) => (
        <button key={k} onClick={() => onChange(k)}
          style={{
            fontFamily: SANS, borderRadius: 8, fontSize: 12.5,
            background: value === k ? C.panel2 : "transparent",
            color: value === k ? C.gold : C.dim,
            border: `1px solid ${value === k ? C.line : "transparent"}`,
          }} className="px-3 py-1.5">
          {label}
        </button>
      ))}
    </div>
  );
}
const chartTip = {
  contentStyle: { background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 10, fontFamily: SANS, color: C.text },
  labelStyle: { color: C.dim }, itemStyle: { color: C.text },
};

/* health badge: red = buy now, yellow = buy soon, green = healthy */
const HEALTH_CFG = {
  red: { bg: C.red, label: "BUY NOW" },
  yellow: { bg: "#c8863a", label: "BUY SOON" },
  green: { bg: C.green, label: "OK" },
};
function HealthBadge({ health, weeksOfStock }) {
  if (!health) return <span style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>—</span>;
  const cfg = HEALTH_CFG[health];
  return (
    <span style={{
      background: cfg.bg, color: "#1a1410", borderRadius: 6, fontFamily: SANS,
      fontWeight: 700, fontSize: 11, padding: "3px 8px", letterSpacing: ".03em", whiteSpace: "nowrap",
    }}>
      {cfg.label}{weeksOfStock != null ? ` · ${weeksOfStock}w` : ""}
    </span>
  );
}

const GRADE_CFG = {
  A: { bg: C.green, label: "A" },
  B: { bg: C.blue, label: "B" },
  C: { bg: C.gold, label: "C" },
  D: { bg: "#c8863a", label: "D" },
  F: { bg: C.red, label: "F" },
};
function GradeBadge({ grade }) {
  if (!grade) return <span style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>—</span>;
  const cfg = GRADE_CFG[grade] || { bg: C.faint, label: grade };
  return (
    <span style={{
      background: cfg.bg, color: "#1a1410", borderRadius: 6, fontFamily: SANS,
      fontWeight: 700, fontSize: 12, padding: "3px 9px", letterSpacing: ".03em", whiteSpace: "nowrap",
    }}>
      {cfg.label}
    </span>
  );
}

/* generic chip-based filter row — used for brand / line / health filters */
function ChipFilter({ label, items, selected, onToggle, onAll, render }) {
  if (!items || items.length < 2) return null;
  return (
    <div className="flex flex-wrap items-center gap-2 mb-1">
      <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".08em", minWidth: 50 }}>
        {label}
      </span>
      <button onClick={onAll}
        style={{
          fontFamily: SANS, borderRadius: 999, fontSize: 12,
          border: `1px solid ${selected.size === items.length ? C.gold : C.line}`,
          background: selected.size === items.length ? C.gold : "transparent",
          color: selected.size === items.length ? C.bg : C.dim,
        }} className="px-3 py-1">All</button>
      {items.map((it) => (
        <button key={it} onClick={() => onToggle(it)}
          style={{
            fontFamily: SANS, borderRadius: 999, fontSize: 12,
            border: `1px solid ${selected.has(it) ? C.gold : C.line}`,
            background: selected.has(it) ? C.gold : "transparent",
            color: selected.has(it) ? C.bg : C.dim,
          }} className="px-3 py-1">{render ? render(it) : it}</button>
      ))}
    </div>
  );
}

/* in-stock / out-of-stock segmented toggle */
function StockFilter({ value, onChange }) {
  const opts = [["all", "All"], ["in", "In stock"], ["out", "Out of stock"]];
  return (
    <div className="flex flex-wrap items-center gap-2 mb-1">
      <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".08em", minWidth: 50 }}>
        Stock
      </span>
      {opts.map(([k, lbl]) => (
        <button key={k} onClick={() => onChange(k)}
          style={{
            fontFamily: SANS, borderRadius: 999, fontSize: 12,
            border: `1px solid ${value === k ? C.gold : C.line}`,
            background: value === k ? C.gold : "transparent",
            color: value === k ? C.bg : C.dim,
          }} className="px-3 py-1">{lbl}</button>
      ))}
    </div>
  );
}

/* presold (0-2 day) sales toggle */
/* searchable multi-select that excludes specific watch models from the totals */
function ModelExclusionFilter({ options, excluded, onAdd, onRemove, onClear, scope, onScope }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const q = query.trim().toLowerCase();
  const matches = q
    ? options.filter((o) =>
        !excluded.has(o.key) &&
        (o.label.toLowerCase().includes(q) ||
         (o.modelNumber && o.modelNumber.toLowerCase().includes(q)) ||
         (o.modelName && o.modelName.toLowerCase().includes(q)) ||
         (o.brand && o.brand.toLowerCase().includes(q)))
      ).slice(0, 12)
    : [];
  const byKey = {};
  options.forEach((o) => { byKey[o.key] = o; });
  const excludedList = Array.from(excluded).map((k) => byKey[k] || { key: k, label: k });
  const scopeOpts = [["dashboard", "Dashboard-wide"], ["summary", "Summary totals only"]];

  return (
    <div className="flex flex-col gap-1.5 mb-1">
      <div className="flex flex-wrap items-start gap-2">
        <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".08em", minWidth: 50, marginTop: 6 }}>
          Exclude
        </span>
        <div style={{ position: "relative", minWidth: 260 }}>
          <input
            value={query}
            onChange={(e) => { setQuery(e.target.value); setOpen(true); }}
            onFocus={() => setOpen(true)}
            onBlur={() => setTimeout(() => setOpen(false), 150)}
            placeholder="Search a model # or name to exclude…"
            style={{
              fontFamily: SANS, fontSize: 12, color: C.text, background: C.bg,
              border: `1px solid ${C.line}`, borderRadius: 8, padding: "5px 10px", width: 280,
            }} />
          {open && matches.length > 0 && (
            <div style={{
              position: "absolute", zIndex: 30, top: "calc(100% + 4px)", left: 0, width: 320,
              background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 8,
              maxHeight: 240, overflowY: "auto", boxShadow: "0 8px 24px rgba(0,0,0,.4)",
            }}>
              {matches.map((o) => (
                <button key={o.key}
                  onMouseDown={(e) => { e.preventDefault(); onAdd(o.key); setQuery(""); }}
                  style={{
                    display: "block", width: "100%", textAlign: "left", fontFamily: SANS, fontSize: 12,
                    color: C.text, background: "transparent", border: "none", padding: "7px 10px", cursor: "pointer",
                  }}
                  className="hover:opacity-80">
                  {o.label}{o.count ? <span style={{ color: C.faint }}> · {o.count}</span> : null}
                </button>
              ))}
            </div>
          )}
        </div>
        {excludedList.length > 0 && (
          <button onClick={onClear}
            style={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${C.line}`, background: "transparent", color: C.dim, marginTop: 1 }}
            className="px-3 py-1">Clear all</button>
        )}
      </div>
      {excludedList.length > 0 && (
        <div className="flex flex-wrap items-center gap-2" style={{ marginLeft: 58 }}>
          {excludedList.map((o) => (
            <button key={o.key} onClick={() => onRemove(o.key)}
              style={{ fontFamily: SANS, fontSize: 12, borderRadius: 999, border: `1px solid ${C.gold}`, background: "transparent", color: C.gold }}
              className="px-3 py-1">{o.label} ✕</button>
          ))}
        </div>
      )}
      {excludedList.length > 0 && (
        <div className="flex flex-wrap items-center gap-2" style={{ marginLeft: 58 }}>
          <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11 }}>Scope:</span>
          {scopeOpts.map(([k, lbl]) => (
            <button key={k} onClick={() => onScope(k)}
              style={{
                fontFamily: SANS, fontSize: 12, borderRadius: 999,
                border: `1px solid ${scope === k ? C.gold : C.line}`,
                background: scope === k ? C.gold : "transparent",
                color: scope === k ? C.bg : C.dim,
              }} className="px-3 py-1">{lbl}</button>
          ))}
          <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11 }}>
            {scope === "dashboard" ? "excluded from the whole dashboard" : "excluded from summary totals only — still shown in detail tables"}
          </span>
        </div>
      )}
    </div>
  );
}

function PresoldFilter({ value, onChange, count }) {
  const opts = [[false, "Exclude presold (default)"], [true, "Include presold"]];
  return (
    <div className="flex flex-wrap items-center gap-2 mb-1">
      <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".08em", minWidth: 50 }}>
        Presold
      </span>
      {opts.map(([k, lbl]) => (
        <button key={String(k)} onClick={() => onChange(k)}
          style={{
            fontFamily: SANS, borderRadius: 999, fontSize: 12,
            border: `1px solid ${value === k ? C.gold : C.line}`,
            background: value === k ? C.gold : "transparent",
            color: value === k ? C.bg : C.dim,
          }} className="px-3 py-1">{lbl}</button>
      ))}
      <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11 }}>
        Sales completed within 0–2 days of purchase ({count ?? 0}) are considered "presold" and {value ? "are included in" : "are not part of"} this analysis.
      </span>
    </div>
  );
}

/* sales-period control: "last N days from today" window (adjustable) vs all time.
   applies to every sales-derived section so the whole dashboard shares one period. */
function SalesPeriodFilter({ on, onToggle, days, onDays }) {
  const opts = [[true, `Last ${days} days`], [false, "All time"]];
  return (
    <div className="flex flex-wrap items-center gap-2 mb-1">
      <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".08em", minWidth: 50 }}>
        Period
      </span>
      {opts.map(([k, lbl]) => (
        <button key={String(k)} onClick={() => onToggle(k)}
          style={{
            fontFamily: SANS, borderRadius: 999, fontSize: 12,
            border: `1px solid ${on === k ? C.gold : C.line}`,
            background: on === k ? C.gold : "transparent",
            color: on === k ? C.bg : C.dim,
          }} className="px-3 py-1">{lbl}</button>
      ))}
      {on && (
        <span className="flex items-center gap-1" style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>
          window:
          <input type="number" min={1} max={3650} value={days}
            onChange={(e) => { const n = parseInt(e.target.value, 10); if (!isNaN(n) && n > 0) onDays(n); }}
            style={{ width: 64, fontFamily: SANS, fontSize: 12, color: C.text, background: C.bg, border: `1px solid ${C.line}`, borderRadius: 8, padding: "3px 8px" }} />
          days from today
        </span>
      )}
      <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11 }}>
        {on ? "All sales figures reflect this window." : "Showing all sales on record."}
      </span>
    </div>
  );
}

/* sale-date range filter (only relevant when sales data is loaded) */
function DateRangeFilter({ value, onChange, min, max }) {
  const fmt = (d) => d ? new Date(d).toISOString().slice(0, 10) : undefined;
  return (
    <div className="flex flex-wrap items-center gap-2 mb-1">
      <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".08em", minWidth: 50 }}>
        Sale date
      </span>
      <input type="date" value={value.start} min={fmt(min)} max={fmt(max)}
        onChange={(e) => onChange({ ...value, start: e.target.value })}
        style={{ fontFamily: SANS, fontSize: 12, border: `1px solid ${C.line}`, borderRadius: 8, background: "transparent", color: C.dim }}
        className="px-2 py-1" />
      <span style={{ color: C.faint, fontSize: 12 }}>to</span>
      <input type="date" value={value.end} min={fmt(min)} max={fmt(max)}
        onChange={(e) => onChange({ ...value, end: e.target.value })}
        style={{ fontFamily: SANS, fontSize: 12, border: `1px solid ${C.line}`, borderRadius: 8, background: "transparent", color: C.dim }}
        className="px-2 py-1" />
      {(value.start || value.end) && (
        <button onClick={() => onChange({ start: "", end: "" })}
          style={{ fontFamily: SANS, fontSize: 12, color: C.dim, border: `1px solid ${C.line}`, borderRadius: 999, background: "transparent" }}
          className="px-3 py-1">Clear</button>
      )}
    </div>
  );
}

/* generic chip-selection state: clicking a chip while "all" are selected
   isolates that item; clicking further chips adds/removes from the set. */
function useChipFilter(allItems) {
  const [selected, setSelected] = useState(new Set(allItems));
  useEffect(() => { setSelected(new Set(allItems)); }, [allItems.join("|")]);
  const toggle = (item) => setSelected((prev) => {
    const allSelected = prev.size === allItems.length;
    if (allSelected) return new Set([item]);
    const next = new Set(prev);
    next.has(item) ? next.delete(item) : next.add(item);
    return next.size ? next : new Set(allItems);
  });
  const selectAll = () => setSelected(new Set(allItems));
  const isolate = (item) => setSelected(new Set([item]));
  const filterSet = selected.size === allItems.length ? null : selected;
  return [selected, toggle, selectAll, filterSet, isolate, setSelected];
}

/* =========================================================================
   MAIN
   ========================================================================= */
export default function WatchBI() {
  const [stage, setStage] = useState("upload"); // upload | map | dash
  const [datasets, setDatasets] = useState([]);
  const [err, setErr] = useState("");
  const fileRef = useRef();

  useEffect(() => {
    if (document.getElementById("wbi-fonts")) return;
    const l = document.createElement("link");
    l.id = "wbi-fonts"; l.rel = "stylesheet";
    l.href = "https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=DM+Sans:wght@400;500;600&display=swap";
    document.head.appendChild(l);
  }, []);

  async function handleFiles(fileList) {
    setErr("");
    try {
      const next = [];
      for (const file of Array.from(fileList)) {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array", cellDates: true });
        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
          if (!rows.length) continue;
          const columns = Object.keys(rows[0]);
          if (columns.length < 2) continue;
          const role = guessRole(columns);
          next.push({
            id: file.name + "::" + sheetName + "::" + Math.random().toString(36).slice(2, 6),
            fileName: file.name, sheetName, columns, rows,
            role, mapping: role === "ignore" ? {} : autoMap(role, columns),
          });
        }
      }
      if (!next.length) { setErr("No readable sheets found in that file."); return; }
      setDatasets((d) => [...d, ...next]);
      setStage("map");
    } catch (e) {
      setErr("Couldn't read that file: " + (e && e.message ? e.message : String(e)));
    }
  }

  function setRole(id, role) {
    setDatasets((ds) => ds.map((d) =>
      d.id === id ? { ...d, role, mapping: role === "ignore" ? {} : autoMap(role, d.columns) } : d));
  }
  function setMap(id, field, col) {
    setDatasets((ds) => ds.map((d) =>
      d.id === id ? { ...d, mapping: { ...d.mapping, [field]: col || undefined } } : d));
  }
  function removeDs(id) { setDatasets((ds) => ds.filter((d) => d.id !== id)); }

  const active = datasets.filter((d) => d.role !== "ignore");

  // pull the dealer's live inventory from the WatchOps API and inject as an inventory dataset
  const [invSync, setInvSync] = useState({ state: "idle", info: null });
  const INV_COLS = ["brand", "modelName", "modelNumber", "cost", "purchaseDate", "targetWholesale", "tagPrice", "condition", "status", "invType", "paymentStatus", "supplier", "serial"];
  async function loadWatchOpsInventory() {
    setInvSync({ state: "loading", info: null });
    try {
      const res = await fetch("/api/inventory");
      const data = await res.json();
      if (data.error) { setInvSync({ state: "error", info: data.error }); return; }
      const mapping = Object.fromEntries(INV_COLS.map((c) => [c, c]));
      const ds = {
        id: "watchops-inventory", fileName: "WatchOps API",
        sheetName: `live inventory · ${data.count} watches`, columns: INV_COLS, rows: data.rows || [],
        role: "inventory", mapping,
      };
      setDatasets((prev) => [...prev.filter((d) => d.id !== "watchops-inventory"), ds]);
      setInvSync({ state: "loaded", info: { count: data.count } });
      setStage("map");
    } catch (e) { setInvSync({ state: "error", info: String(e && e.message || e) }); }
  }

  // pull recent WhatsApp messages from Google Drive and inject as a messages dataset
  const [msgSync, setMsgSync] = useState({ state: "idle", info: null });
  const MSG_COLS = ["intent", "brand", "model", "reference", "price", "messageBody", "timestamp", "sender", "chat"];
  const MSG_MAP = Object.fromEntries(MSG_COLS.map((c) => [c, c]));
  async function syncGoogleMessages(days = 5) {
    setMsgSync({ state: "loading", info: null });
    try {
      const res = await fetch(`/api/messages?days=${days}`);
      const data = await res.json();
      if (data.error) { setMsgSync({ state: "error", info: data.error }); return; }
      const ds = {
        id: "google-messages", fileName: "Google Drive",
        sheetName: `messages · last ${data.days}d`, columns: MSG_COLS, rows: data.rows || [],
        role: "messages", mapping: MSG_MAP,
      };
      setDatasets((prev) => [...prev.filter((d) => d.id !== "google-messages"), ds]);
      setMsgSync({ state: "loaded", info: { count: (data.rows || []).length, latest: data.latest, files: (data.files || []).length } });
    } catch (e) {
      setMsgSync({ state: "error", info: String(e && e.message || e) });
    }
  }
  // auto-pull once when the dashboard opens (if not already loaded)
  useEffect(() => {
    if (stage === "dash" && msgSync.state === "idle" && !datasets.some((d) => d.id === "google-messages")) {
      syncGoogleMessages(5);
    }
  }, [stage]);

  const [dateRange, setDateRange] = useState({ start: "", end: "" });
  const [includePresold, setIncludePresold] = useState(false);
  // sales period: "last N days from today" window, on by default (45 days), adjustable / switchable
  const [salesWindowOn, setSalesWindowOn] = useState(true);
  const [salesWindowDays, setSalesWindowDays] = useState(45);
  const windowDays = salesWindowOn ? salesWindowDays : null;
  // starts null (all statuses included); Dashboard sets this to exclude
  // "On Hold / Reserved" after the first metrics render via useEffect.
  const [invStatusFilter, setInvStatusFilter] = useState(null);
  // model exclusion filter: a set of model keys to drop, plus scope of effect
  const [excludedModels, setExcludedModels] = useState(() => new Set());
  const [exclScope, setExclScope] = useState("dashboard"); // "dashboard" | "summary"
  const hasExcl = excludedModels.size > 0;

  // full (no exclusions) metrics — always computed; drives the model picker + "with" view
  const metricsFull = useMemo(() => (stage === "dash" ? computeMetrics(active, dateRange, includePresold, windowDays, invStatusFilter, null) : null), [stage, datasets, dateRange, includePresold, windowDays, invStatusFilter]);
  // adjusted metrics — same but with excluded models removed
  const metricsExcl = useMemo(() => (stage === "dash" && hasExcl ? computeMetrics(active, dateRange, includePresold, windowDays, invStatusFilter, excludedModels) : metricsFull), [stage, datasets, dateRange, includePresold, windowDays, invStatusFilter, excludedModels, hasExcl, metricsFull]);
  // which one drives the dashboard body: dashboard-wide exclusion → adjusted; summary-only → full
  const metrics = (hasExcl && exclScope === "dashboard") ? metricsExcl : metricsFull;

  return (
    <div style={{ background: C.bg, color: C.text, fontFamily: SANS, minHeight: 600 }} className="w-full">
      {/* header */}
      <div style={{ borderBottom: `1px solid ${C.line}` }} className="px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <Watch size={26} style={{ color: C.gold }} />
          <div>
            <div style={{ fontFamily: SERIF, color: C.text }} className="text-xl leading-none">Horometrics</div>
            <div style={{ color: C.faint }} className="text-xs tracking-wide">inventory & sales intelligence</div>
          </div>
        </div>
        {stage === "dash" && (
          <button onClick={() => { setDatasets([]); setStage("upload"); }}
            style={{ color: C.dim, border: `1px solid ${C.line}`, borderRadius: 10, fontFamily: SANS }}
            className="text-xs px-3 py-2 hover:opacity-80">Load different data</button>
        )}
      </div>

      {stage === "upload" && <UploadView fileRef={fileRef} onFiles={handleFiles} err={err} onLoadWatchOps={loadWatchOpsInventory} invSync={invSync} />}
      {stage === "map" && (
        <MapView datasets={datasets} setRole={setRole} setMap={setMap} removeDs={removeDs}
          onBuild={() => setStage("dash")} onAdd={() => fileRef.current?.click()} fileRef={fileRef} onFiles={handleFiles} />
      )}
      {stage === "dash" && metrics && <Dashboard M={metrics} MFull={metricsFull} MExcl={metricsExcl} dateRange={dateRange} setDateRange={setDateRange} includePresold={includePresold} setIncludePresold={setIncludePresold} salesWindowOn={salesWindowOn} setSalesWindowOn={setSalesWindowOn} salesWindowDays={salesWindowDays} setSalesWindowDays={setSalesWindowDays} invStatusFilter={invStatusFilter} setInvStatusFilter={setInvStatusFilter} excludedModels={excludedModels} setExcludedModels={setExcludedModels} exclScope={exclScope} setExclScope={setExclScope} msgSync={msgSync} onSyncMessages={syncGoogleMessages} />}
    </div>
  );
}

/* ---------- upload ---------- */
function UploadView({ fileRef, onFiles, err, onLoadWatchOps, invSync }) {
  const [drag, setDrag] = useState(false);
  const loadingInv = invSync && invSync.state === "loading";
  return (
    <div className="p-8 flex flex-col items-center" style={{ minHeight: 480, justifyContent: "center" }}>
      <input ref={fileRef} type="file" multiple accept=".xlsx,.xls,.csv" className="hidden"
        onChange={(e) => e.target.files.length && onFiles(e.target.files)} />
      {err && (
        <div style={{ background: "#3a201b", border: `1px solid ${C.red}`, color: C.text, borderRadius: 10, maxWidth: 560 }}
          className="px-4 py-3 mb-4 text-sm flex items-center gap-2 w-full">
          <AlertTriangle size={16} style={{ color: C.red }} /> {err}
        </div>
      )}
      <div
        onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => { e.preventDefault(); setDrag(false); e.dataTransfer.files.length && onFiles(e.dataTransfer.files); }}
        onClick={() => fileRef.current?.click()}
        style={{
          border: `1.5px dashed ${drag ? C.gold : C.line}`, borderRadius: 18,
          background: drag ? C.panel2 : C.panel, width: "100%", maxWidth: 560, cursor: "pointer",
        }}
        className="p-12 flex flex-col items-center gap-4 transition-colors">
        <Upload size={40} style={{ color: C.gold }} />
        <div style={{ fontFamily: SERIF }} className="text-2xl">Drop your spreadsheets</div>
        <div style={{ color: C.dim }} className="text-sm text-center max-w-sm">
          Inventory export, sales export, or both. Excel or CSV. Nothing is uploaded anywhere — it's read in your browser.
        </div>
        <div style={{ color: C.faint }} className="text-xs">.xlsx · .xls · .csv · multiple files & sheets supported</div>
      </div>

      {onLoadWatchOps && (
        <div className="flex flex-col items-center gap-2 mt-5" style={{ width: "100%", maxWidth: 560 }}>
          <div className="flex items-center gap-3 w-full">
            <div style={{ flex: 1, height: 1, background: C.line }} />
            <span style={{ color: C.faint, fontFamily: SANS }} className="text-xs">or pull it live</span>
            <div style={{ flex: 1, height: 1, background: C.line }} />
          </div>
          <button onClick={onLoadWatchOps} disabled={loadingInv}
            style={{ background: loadingInv ? C.line : C.panel2, color: loadingInv ? C.faint : C.gold, border: `1px solid ${C.gold}`, borderRadius: 12, fontFamily: SANS }}
            className="px-5 py-2.5 text-sm flex items-center gap-2">
            {loadingInv ? "Loading inventory…" : "Load inventory from WatchOps"}
          </button>
          {invSync && invSync.state === "error" && <span style={{ color: C.red, fontFamily: SANS }} className="text-xs">{String(invSync.info)}</span>}
          <span style={{ color: C.faint, fontFamily: SANS }} className="text-xs text-center">Pulls your current holdings from WatchOps; add a sales export on the next screen.</span>
        </div>
      )}

      <div style={{ color: C.faint }} className="text-xs mt-6 max-w-md text-center">
        The tool detects your columns and lets you confirm the mapping, so it works on any dealer's export — not just one fixed format.
      </div>
    </div>
  );
}

/* ---------- mapping ---------- */
function MapView({ datasets, setRole, setMap, removeDs, onBuild, onAdd }) {
  const usable = datasets.filter((d) => d.role !== "ignore");
  const canBuild = usable.length > 0 && usable.every((d) => {
    if (d.role === "inventory") return d.mapping.brand || d.mapping.cost;
    if (d.role === "sales") return d.mapping.saleDate || d.mapping.salePrice;
    return true;
  });
  return (
    <div className="p-6">
      <div className="flex items-baseline justify-between mb-4">
        <h2 style={{ fontFamily: SERIF }} className="text-2xl">Confirm the columns</h2>
        <button onClick={onAdd} style={{ color: C.dim, border: `1px solid ${C.line}`, borderRadius: 10 }}
          className="text-xs px-3 py-2">+ add file</button>
      </div>
      <p style={{ color: C.dim }} className="text-sm mb-5 max-w-2xl">
        I auto-detected each sheet. Set what it is, fix any mapping that's off, and ignore anything you don't need (like a yearly summary tab).
      </p>
      <div className="flex flex-col gap-4">
        {datasets.map((d) => (
          <div key={d.id} style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 14 }} className="p-4">
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-2">
                <FileSpreadsheet size={18} style={{ color: C.gold }} />
                <span style={{ fontFamily: SANS }} className="text-sm">{d.fileName}</span>
                <span style={{ color: C.faint }} className="text-xs">· {d.sheetName} · {d.rows.length} rows</span>
              </div>
              <div className="flex items-center gap-2">
                {["inventory", "sales", "messages", "ignore"].map((r) => (
                  <button key={r} onClick={() => setRole(d.id, r)}
                    style={{
                      fontFamily: SANS, borderRadius: 8,
                      border: `1px solid ${d.role === r ? C.gold : C.line}`,
                      background: d.role === r ? C.gold : "transparent",
                      color: d.role === r ? C.bg : C.dim,
                    }} className="text-xs px-3 py-1 capitalize">{r}</button>
                ))}
                <button onClick={() => removeDs(d.id)} style={{ color: C.faint }} className="ml-1"><Trash2 size={15} /></button>
              </div>
            </div>
            {d.role !== "ignore" && (
              <div className="grid grid-cols-1 gap-2" style={{ gridTemplateColumns: "repeat(auto-fill,minmax(230px,1fr))" }}>
                {FIELDS[d.role].map(([field, label]) => (
                  <div key={field} className="flex flex-col gap-1">
                    <label style={{ color: C.dim, fontFamily: SANS }} className="text-xs">{label}</label>
                    <select value={d.mapping[field] || ""} onChange={(e) => setMap(d.id, field, e.target.value)}
                      style={{ background: C.panel2, color: C.text, border: `1px solid ${C.line}`, borderRadius: 8, fontFamily: SANS }}
                      className="text-xs px-2 py-2">
                      <option value="">— not mapped —</option>
                      {d.columns.map((c) => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                ))}
              </div>
            )}
          </div>
        ))}
      </div>
      <div className="mt-6 flex items-center gap-4">
        <button disabled={!canBuild} onClick={onBuild}
          style={{
            background: canBuild ? C.gold : C.line, color: canBuild ? C.bg : C.faint,
            borderRadius: 12, fontFamily: SANS, cursor: canBuild ? "pointer" : "default",
          }} className="px-6 py-3 text-sm flex items-center gap-2 font-medium">
          Build dashboard <ArrowRight size={16} />
        </button>
        {!canBuild && <span style={{ color: C.faint }} className="text-xs">Map at least cost/brand for inventory or a sale date/price for sales.</span>}
      </div>
    </div>
  );
}

/* ---------- dashboard ---------- */
const ON_HOLD_STATUS = "On Hold / Reserved";

function Dashboard({ M, MFull, MExcl, dateRange, setDateRange, includePresold, setIncludePresold, salesWindowOn, setSalesWindowOn, salesWindowDays, setSalesWindowDays, invStatusFilter, setInvStatusFilter, excludedModels, setExcludedModels, exclScope, setExclScope, msgSync, onSyncMessages }) {
  const hasExcl = excludedModels && excludedModels.size > 0;
  const addExcl = (k) => setExcludedModels((prev) => new Set(prev).add(k));
  const removeExcl = (k) => setExcludedModels((prev) => { const n = new Set(prev); n.delete(k); return n; });
  const clearExcl = () => setExcludedModels(new Set());
  const [tab, setTab] = useState(M.hasInv ? "inventory" : "sales");
  const tabs = [
    M.hasInv && ["inventory", "Inventory", Boxes],
    M.hasSales && ["sales", "Sales", TrendingUp],
    M.hasSales && ["buy", "Buy Signals", Sparkles],
    (M.hasInv || M.hasSales) && ["liabilities", "Liabilities", AlertTriangle],
    (M.hasMessages || (msgSync && msgSync.state !== "idle")) && ["alerts", "Alerts", Bell],
  ].filter(Boolean);

  // filters shared across Inventory / Sales / Buy tabs
  const allBrands = useMemo(() => {
    const s = new Set();
    (M.salesByBrand || []).forEach((b) => s.add(b.brand));
    (M.invByBrand || []).forEach((b) => s.add(b.brand));
    return Array.from(s).sort();
  }, [M]);
  const allLines = useMemo(() => {
    const s = new Set();
    (M.salesByLine || []).forEach((l) => l.line && s.add(l.line));
    (M.invByLine || []).forEach((l) => l.line && s.add(l.line));
    (M.salesByModel || []).forEach((l) => l.line && s.add(l.line));
    return Array.from(s).sort();
  }, [M]);
  const allHealth = useMemo(() => ["red", "yellow", "green"], []);

  const [selectedBrands, toggleBrand, selectAllBrands, brandFilterSet] = useChipFilter(allBrands);
  const [selectedLines, toggleLine, selectAllLines, lineFilterSet, , setSelectedLines] = useChipFilter(allLines);

  // map of brand -> set of product lines that belong to it, so selecting a
  // brand can auto-select its lines in the Line filter
  const brandLineMap = useMemo(() => {
    const map = {};
    const addAll = (arr) => (arr || []).forEach((x) => {
      if (!x.line || !x.brand) return;
      if (!map[x.brand]) map[x.brand] = new Set();
      map[x.brand].add(x.line);
    });
    addAll(M.salesByLine); addAll(M.invByLine); addAll(M.salesByModel);
    return map;
  }, [M]);

  function syncLinesToBrands(newBrands) {
    if (newBrands.size === allBrands.length) {
      setSelectedLines(new Set(allLines));
      return;
    }
    const lines = new Set();
    newBrands.forEach((b) => (brandLineMap[b] || new Set()).forEach((l) => lines.add(l)));
    setSelectedLines(lines.size ? lines : new Set(allLines));
  }

  function handleToggleBrand(item) {
    toggleBrand(item);
    const wasAllSelected = selectedBrands.size === allBrands.length;
    let newBrands;
    if (wasAllSelected) {
      newBrands = new Set([item]);
    } else {
      newBrands = new Set(selectedBrands);
      newBrands.has(item) ? newBrands.delete(item) : newBrands.add(item);
      if (!newBrands.size) newBrands = new Set(allBrands);
    }
    syncLinesToBrands(newBrands);
  }

  function handleSelectAllBrands() {
    selectAllBrands();
    setSelectedLines(new Set(allLines));
  }
  const [selectedHealth, toggleHealth, selectAllHealth, healthFilterSet] = useChipFilter(allHealth);
  const [stockFilter, setStockFilter] = useState("all");
  const [filtersOpen, setFiltersOpen] = useState(false);

  // inventory status filter — default excludes "On Hold / Reserved"
  const allStatuses = M.invStatuses || [];
  const defaultStatuses = useMemo(
    () => new Set(allStatuses.filter((s) => s !== ON_HOLD_STATUS)),
    [allStatuses.join("|")]
  );
  const [selectedStatuses, setSelectedStatuses] = useState(() => defaultStatuses);
  // keep selectedStatuses in sync when the loaded dataset changes
  useEffect(() => { setSelectedStatuses(defaultStatuses); }, [allStatuses.join("|")]);
  // push the derived filter set up to App so computeMetrics re-runs
  useEffect(() => {
    const isAll = selectedStatuses.size === allStatuses.length;
    setInvStatusFilter(isAll || !allStatuses.length ? null : selectedStatuses);
  }, [selectedStatuses, allStatuses.join("|")]);

  function toggleStatus(s) {
    setSelectedStatuses((prev) => {
      const isAll = prev.size === allStatuses.length;
      if (isAll) return new Set([s]);
      const next = new Set(prev);
      next.has(s) ? next.delete(s) : next.add(s);
      return next.size ? next : new Set(allStatuses);
    });
  }
  function selectAllStatuses() { setSelectedStatuses(new Set(allStatuses)); }
  const statusFilterActive = allStatuses.length > 0 && selectedStatuses.size < allStatuses.length;

  const filters = {
    brands: brandFilterSet,
    lines: lineFilterSet,
    health: healthFilterSet,
    stock: stockFilter,
  };
  const dateActive = !!(dateRange && (dateRange.start || dateRange.end));
  const filtersActiveCount =
    (brandFilterSet ? 1 : 0) + (lineFilterSet ? 1 : 0) + (healthFilterSet ? 1 : 0) +
    (stockFilter !== "all" ? 1 : 0) + (dateActive ? 1 : 0) + (includePresold ? 1 : 0) +
    (salesWindowOn ? 1 : 0) + (statusFilterActive ? 1 : 0) + (hasExcl ? 1 : 0);

  return (
    <div style={{ minHeight: 520 }}>
      <div className="p-6">
        <div className="flex gap-1 mb-5">
          {tabs.map(([k, label, Icon]) => (
            <button key={k} onClick={() => setTab(k)}
              style={{
                fontFamily: SANS, borderRadius: 10,
                background: tab === k ? C.panel2 : "transparent",
                color: tab === k ? C.gold : C.dim,
                border: `1px solid ${tab === k ? C.line : "transparent"}`,
              }} className="px-4 py-2 text-sm flex items-center gap-2">
              <Icon size={15} /> {label}
            </button>
          ))}
        </div>
        {allBrands.length > 1 && (
          <div className="mb-3">
            <button onClick={() => setFiltersOpen((o) => !o)}
              style={{
                fontFamily: SANS, fontSize: 11, color: C.dim,
                border: `1px solid ${C.line}`, borderRadius: 8, background: "transparent",
              }} className="px-3 py-1.5 flex items-center gap-2">
              <span style={{ display: "inline-block", transform: filtersOpen ? "rotate(90deg)" : "none", transition: "transform .15s" }}>▸</span>
              Filters
              {filtersActiveCount > 0 ? (
                <span style={{ color: C.gold }}>({filtersActiveCount} active)</span>
              ) : null}
            </button>
            {filtersOpen && (
              <div className="mt-2 flex flex-col gap-1.5">
                <ChipFilter label="Brand" items={allBrands} selected={selectedBrands} onToggle={handleToggleBrand} onAll={handleSelectAllBrands} />
                <ChipFilter label="Line" items={allLines} selected={selectedLines} onToggle={toggleLine} onAll={selectAllLines} />
                <ChipFilter label="Health" items={allHealth} selected={selectedHealth} onToggle={toggleHealth} onAll={selectAllHealth}
                  render={(h) => HEALTH_CFG[h]?.label || h} />
                <StockFilter value={stockFilter} onChange={setStockFilter} />
                {allStatuses.length > 0 && (
                  <ChipFilter label="Status" items={allStatuses} selected={selectedStatuses} onToggle={toggleStatus} onAll={selectAllStatuses} />
                )}
                {M.hasSales && (
                  <SalesPeriodFilter on={salesWindowOn} onToggle={setSalesWindowOn} days={salesWindowDays} onDays={setSalesWindowDays} />
                )}
                {M.hasSales && (
                  <PresoldFilter value={includePresold} onChange={setIncludePresold} count={M.presoldCount} />
                )}
                {M.hasSales && (
                  <DateRangeFilter value={dateRange} onChange={setDateRange} min={M.salesDateMin} max={M.salesDateMax} />
                )}
                <ModelExclusionFilter
                  options={(MFull && MFull.modelOptions) || []}
                  excluded={excludedModels}
                  onAdd={addExcl} onRemove={removeExcl} onClear={clearExcl}
                  scope={exclScope} onScope={setExclScope} />
              </div>
            )}
          </div>
        )}
        {hasExcl && (tab === "inventory" || tab === "sales") && (
          <div className="mb-5">
            <ExclusionSummary
              excludedCount={excludedModels.size}
              scope={exclScope}
              rows={tab === "sales"
                ? [
                    { label: "Invoice total (revenue)", full: MFull.salesRevenue, adj: MExcl.salesRevenue, fmt: fmtMoney },
                    { label: "COGS", full: MFull.salesCOGS, adj: MExcl.salesCOGS, fmt: fmtMoney },
                    { label: "Profit", full: MFull.salesProfit, adj: MExcl.salesProfit, fmt: fmtMoney },
                    { label: "Profit %", full: MFull.salesProfitPct, adj: MExcl.salesProfitPct, fmt: fmtPct },
                    { label: "Units sold", full: MFull.salesUnits, adj: MExcl.salesUnits },
                  ]
                : [
                    { label: "Items in stock", full: MFull.invCount, adj: MExcl.invCount },
                    { label: "Capital tied up", full: MFull.invCost, adj: MExcl.invCost, fmt: fmtMoney },
                  ]} />
          </div>
        )}
        {tab === "inventory" && <InventoryTab M={M} filters={filters} />}
        {tab === "sales" && <SalesTab M={M} filters={filters} />}
        {tab === "buy" && <BuyTab M={M} filters={filters} includePresold={includePresold} />}
        {tab === "liabilities" && <LiabilitiesTab M={M} />}
        {tab === "alerts" && <AlertsTab M={M} msgSync={msgSync} onSyncMessages={onSyncMessages} />}
      </div>
    </div>
  );
}

/* apply the shared brand / line / health / stock filters to a metrics array.
   a check is skipped entirely if the row's dataset doesn't carry that field at
   all (e.g. brand-only aggregates have no `.line`/`.health`/`.stock` key, so
   those filters are a no-op there). but if the field DOES exist on the row
   (even as `null`, e.g. an item with no recognized product line / no health
   data), an active filter excludes it — null can't match any selected chip. */
function applyFilters(rows, filters) {
  if (!rows) return [];
  if (!filters) return rows;
  return rows.filter((r) => {
    if (filters.brands && !filters.brands.has(r.brand)) return false;
    if (filters.lines && "line" in r) {
      if (r.line == null || !filters.lines.has(r.line)) return false;
    }
    if (filters.health && "health" in r) {
      if (r.health == null || !filters.health.has(r.health)) return false;
    }
    if (filters.stock === "in" && r.stock === 0) return false;
    if (filters.stock === "out" && !(r.stock === 0)) return false;
    return true;
  });
}

/* dynamic height: perRow px per bar + 48px for axes, minimum min px */
function barH(n, perRow = 38, min = 180) {
  return Math.max(min, n * perRow + 48);
}
/* dynamic YAxis width: base on longest label */
function yAxisW(items, key = "brand", base = 80) {
  if (!items || !items.length) return base;
  const longest = Math.max(...items.map((x) => String(x[key] || "").length));
  return Math.min(Math.max(base, longest * 7), 210);
}

function SectionLabel({ children }) {
  return (
    <div style={{ color: C.gold, fontFamily: SANS, letterSpacing: ".1em", fontSize: 11, textTransform: "uppercase", fontWeight: 600, marginBottom: 12, marginTop: 4 }}>
      {children}
    </div>
  );
}

// columns used to list individual watches in drill-downs
const WATCH_COLS = [
  ["brand","Brand"],["modelName","Model"],["modelNumber","Ref #"],
  ["grade","Grade",(v) => <GradeBadge grade={v} />],
  ["age","Days in stock"],["cost","Cost",fmtMoney],
];

/* fetch WatchCharts median asking price for a set of {brand, reference} rows.
   Server-side cached, so repeated views don't re-spend data credits. */
function useMarketValues(rows) {
  const [map, setMap] = useState({});
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState(null);
  const valid = (rows || []).filter((r) => r.brand && r.reference);
  const depKey = valid.map((r) => r.brand + "|" + r.reference).join(",");
  useEffect(() => {
    if (!valid.length) { setMap({}); return; }
    let cancelled = false;
    setLoading(true); setErr(null);
    fetch("/api/market", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ items: valid.map((r) => ({ brand: r.brand, reference: r.reference })) }),
    })
      .then((r) => r.json())
      .then((d) => { if (!cancelled) { if (d.error) setErr(d.error); else setMap(d.results || {}); } })
      .catch(() => { if (!cancelled) setErr("Couldn't reach market data."); })
      .finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
  }, [depKey]);
  const lookup = (brand, reference) =>
    map[String(brand || "").toLowerCase().trim() + "|" + String(reference || "").toLowerCase().trim()] || null;
  return { lookup, loading, err };
}

function MarketStatus({ loading, err }) {
  if (err) return <div style={{ color: C.red, fontFamily: SANS, fontSize: 12, marginBottom: 8 }}>Market data unavailable: {err}</div>;
  if (loading) return <div style={{ color: C.gold, fontFamily: SANS, fontSize: 12, marginBottom: 8 }}>Loading WatchCharts market values…</div>;
  return null;
}

/* Oldest 20 in stock — target & cost vs market, with QoC + QoT grades */
function OldestMarketTable({ rows }) {
  const { lookup, loading, err } = useMarketValues(rows);
  const graded = rows.map((r) => {
    const mv = lookup(r.brand, r.reference);
    const m = mv?.medianAsking ?? null;
    return { ...r, median: m, qoc: qocGrade(r.cost, m), qot: qotGrade(r.targetWholesale, m) };
  });
  return (
    <>
      <MarketStatus loading={loading} err={err} />
      <ItemTable rows={graded} cols={[
        ["brand","Brand"],["model","Model"],["reference","Ref #"],["age","Days in stock"],
        ["cost","Cost",fmtMoney],["targetWholesale","Target",(v) => v ? fmtMoney(v) : "—"],
        ["median","Median asking",(v) => v ? fmtMoney(v) : "—"],
        ["qoc","QoC",(v) => <GradeBadge grade={v} />],
        ["qot","QoT",(v) => <GradeBadge grade={v} />],
      ]} />
    </>
  );
}

/* Latest 20 sales — actual sale price vs market, with a sales grade */
function LatestMarketTable({ rows }) {
  const { lookup, loading, err } = useMarketValues(rows);
  const fmtDate = (d) => d ? new Date(d).toISOString().slice(0, 10) : "—";
  const graded = rows.map((r) => {
    const mv = lookup(r.brand, r.reference);
    const m = mv?.medianAsking ?? null;
    return { ...r, median: m, salesGrade: salesGrade(r.price, m), pctOfMarket: m ? (r.price / m) * 100 : null };
  });
  return (
    <>
      <MarketStatus loading={loading} err={err} />
      <ItemTable rows={graded} cols={[
        ["saleDate","Date",fmtDate],["brand","Brand"],["model","Model"],["reference","Ref #"],
        ["price","Sold for",fmtMoney],
        ["median","Median asking",(v) => v ? fmtMoney(v) : "—"],
        ["pctOfMarket","% of market",(v) => v == null ? "—" : v.toFixed(0) + "%"],
        ["salesGrade","Grade",(v) => <GradeBadge grade={v} />],
      ]} />
    </>
  );
}

/* ---------- Alerts: buy/sell signals from WhatsApp dealer-group messages ---------- */
function AlertsTab({ M, msgSync, onSyncMessages }) {
  const [windowDays, setWindowDays] = useState(3);
  const [costTol, setCostTol] = useState(10);   // % above your median cost still "in range"
  const [lowStockOnly, setLowStockOnly] = useState(false);
  const [sellBand, setSellBand] = useState(15); // % under your median sale still worth a sell alert

  const anchor = M.messagesMax ? M.messagesMax.getTime() : Date.now();
  const startT = anchor - windowDays * 86400000;
  const pos = M.refPositions || {};
  const STALE_MS = 18 * 30.44 * 86400000;        // 18 months
  const staleCutoff = Date.now() - STALE_MS;

  // is the message's price actually present in its text? catches mis-parsed prices.
  // "ok" = found, "mismatch" = body has other numbers but not this one, "unverified" = can't tell
  const priceCheck = (price, body) => {
    if (!price) return "noprice";
    const nums = (String(body || "").match(/\d[\d,]{2,}/g) || []).map((n) => parseInt(n.replace(/,/g, ""), 10)).filter((v) => v >= 100);
    if (!nums.length) return "unverified";
    return nums.some((v) => Math.abs(v - price) <= Math.max(50, price * 0.03)) ? "ok" : "mismatch";
  };

  const inWindow = (M.messages || []).filter((m) => m.ts && m.ts.getTime() >= startT);

  // BUY alerts: someone is selling (intent="sell") a ref we trade, priced in our cost range
  const buySeen = new Set();
  const buyAlerts = [];
  inWindow.forEach((m) => {
    if (m.intent !== "sell" || !m.price) return;
    const p = pos[m.ref]; if (!p || !p.medianCost) return;
    // 18-month freshness: don't suggest buying something we haven't sold in 18 months
    if (!p.lastSale || p.lastSale.getTime() < staleCutoff) return;
    // sanity cap: price must be within a believable band of our cost
    if (m.price < p.medianCost * 0.4 || m.price > p.medianCost * 2.5) return;
    // double-check the price appears in the message text
    if (priceCheck(m.price, m.body) === "mismatch") return;
    if (m.price > p.medianCost * (1 + costTol / 100)) return;     // out of your buy range
    const low = p.weeksOfStock != null && p.weeksOfStock < 1;
    if (lowStockOnly && !low) return;
    const dk = m.ref + "|" + Math.round(m.price);
    if (buySeen.has(dk)) return; buySeen.add(dk);
    buyAlerts.push({
      grade: qocGrade(m.price, p.medianCost), ref: m.ref, brand: p.brand, model: p.model,
      price: m.price, medianCost: p.medianCost, loCost: p.loCost, hiCost: p.hiCost,
      stock: p.stock, low, verified: priceCheck(m.price, m.body) === "ok",
      sender: m.sender, chat: m.chat, body: m.body, ts: m.ts,
    });
  });
  const gradeOrder = { A: 0, B: 1, C: 2, D: 3, null: 4 };
  buyAlerts.sort((a, b) => (b.low - a.low) || (gradeOrder[a.grade] - gradeOrder[b.grade]) || (b.ts - a.ts));

  // SELL alerts: someone wants to buy (intent="buy" / WTB) a ref we hold
  const sellSeen = new Set();
  const sellAlerts = [];
  inWindow.forEach((m) => {
    if (m.intent !== "buy") return;
    const p = pos[m.ref]; if (!p || !p.stock) return;
    // sanity cap + price double-check: if the stated price is implausible or not in
    // the text, drop the price and treat it as an ungraded lead rather than a bad grade
    let price = m.price;
    const ref = p.medianSale || p.medianCost;
    if (price && ref && (price < ref * 0.3 || price > ref * 3)) price = null;
    if (price && priceCheck(price, m.body) === "mismatch") price = null;
    const dk = m.ref + "|" + (m.sender || "") + "|" + (price || "np");
    if (sellSeen.has(dk)) return; sellSeen.add(dk);
    const grade = (price && p.medianSale) ? salesGrade(price, p.medianSale) : null;
    sellAlerts.push({
      grade, ref: m.ref, brand: p.brand, model: p.model, price: price || null,
      medianSale: p.medianSale, stock: p.stock, sender: m.sender, chat: m.chat, body: m.body, ts: m.ts,
    });
  });
  sellAlerts.sort((a, b) => ((b.price ? 1 : 0) - (a.price ? 1 : 0)) || (gradeOrder[a.grade] - gradeOrder[b.grade]) || (b.ts - a.ts));

  const fmtWhen = (d) => d ? new Date(d).toISOString().slice(0, 10) : "—";
  const Body = (v) => <span style={{ color: C.dim, fontFamily: SANS, fontSize: 12 }}>{String(v || "").replace(/\s+/g, " ").slice(0, 90)}</span>;

  const syncing = msgSync && msgSync.state === "loading";
  const syncErr = msgSync && msgSync.state === "error";
  return (
    <div className="flex flex-col gap-5">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12, maxWidth: 640 }}>
          Buy & sell signals matched from dealer-group messages against your own cost & sale history — graded A–D vs <b>your</b> numbers,
          not the open market. Showing the last {windowDays} days (through {fmtWhen(M.messagesMax)}).
        </div>
        {onSyncMessages && (
          <button onClick={() => onSyncMessages(5)} disabled={syncing}
            style={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${C.line}`, background: "transparent", color: syncing ? C.faint : C.gold, whiteSpace: "nowrap" }}
            className="px-3 py-1.5">{syncing ? "Syncing…" : "↻ Sync from Google Drive"}</button>
        )}
      </div>
      {syncErr && <div style={{ color: C.red, fontFamily: SANS, fontSize: 12 }}>Google sync failed: {String(msgSync.info)}</div>}

      {!M.hasMessages ? (
        <Locked msg={syncing ? "Pulling recent messages from Google Drive…" : "No messages loaded. Click 'Sync from Google Drive' above, or upload message sheets."} />
      ) : (<>

      {/* controls */}
      <div className="flex flex-wrap items-center gap-4" style={{ fontFamily: SANS, fontSize: 12, color: C.dim }}>
        <label className="flex items-center gap-1">Window
          <input type="number" min={1} max={30} value={windowDays} onChange={(e) => setWindowDays(Math.max(1, parseInt(e.target.value, 10) || 1))}
            style={{ width: 56, background: C.bg, color: C.text, border: `1px solid ${C.line}`, borderRadius: 8, padding: "3px 8px" }} /> days</label>
        <label className="flex items-center gap-1">Buy range: within
          <input type="number" min={0} max={100} value={costTol} onChange={(e) => setCostTol(Math.max(0, parseInt(e.target.value, 10) || 0))}
            style={{ width: 56, background: C.bg, color: C.text, border: `1px solid ${C.line}`, borderRadius: 8, padding: "3px 8px" }} />% of your median cost</label>
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={lowStockOnly} onChange={(e) => setLowStockOnly(e.target.checked)} />
          low-stock only (&lt; 1 week)</label>
      </div>

      <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))" }}>
        <Stat label="Buy alerts" value={buyAlerts.length} sub="offers in your cost range" />
        <Stat label="Sell alerts" value={sellAlerts.length} sub="wanted: watches you hold" />
        <Stat label="Messages scanned" value={inWindow.length} sub={`last ${windowDays} days`} />
      </div>

      {/* BUY */}
      <Panel title="Buy alerts" note="someone is selling a watch you trade, in your price range">
        {buyAlerts.length === 0
          ? <Locked msg="No offers landed in your cost range for this window. Widen the buy range % or window above." />
          : <ItemTable rows={buyAlerts} cols={[
              ["grade","Grade",(v) => <GradeBadge grade={v} />],
              ["low","",(v) => v ? <span style={{ background: C.red, color: "#1a1410", borderRadius: 6, fontFamily: SANS, fontWeight: 700, fontSize: 10, padding: "2px 6px" }}>LOW STOCK</span> : ""],
              ["brand","Brand"],["ref","Ref #"],
              ["price","Offered",fmtMoney],
              ["medianCost","Your med cost",(v) => v ? fmtMoney(v) : "—"],
              ["stock","Stock"],
              ["sender","From"],["chat","Group"],
              ["body","Message",Body],["ts","When",fmtWhen],
            ]} />}
      </Panel>

      {/* SELL */}
      <Panel title="Sell alerts" note="someone is looking for a watch you hold">
        {sellAlerts.length === 0
          ? <Locked msg="No wanted-to-buy posts matched your current stock in this window." />
          : <ItemTable rows={sellAlerts} cols={[
              ["grade","Grade",(v) => v ? <GradeBadge grade={v} /> : <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11 }}>lead</span>],
              ["brand","Brand"],["ref","Ref #"],
              ["price","They'll pay",(v) => v ? fmtMoney(v) : "no price"],
              ["medianSale","Your med sale",(v) => v ? fmtMoney(v) : "—"],
              ["stock","You hold"],
              ["sender","From"],["chat","Group"],
              ["body","Message",Body],["ts","When",fmtWhen],
            ]} />}
      </Panel>
      </>)}
    </div>
  );
}

/* ---------- Liabilities (financial exposure — ignores brand/period filters) ---------- */
function LiabilitiesTab({ M }) {
  const fmtDate = (d) => d ? new Date(d).toISOString().slice(0, 10) : "--";
  return (
    <div className="flex flex-col gap-5">
      <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>
        These reflect real financial exposure across the full data set and are not affected by the brand, status, or period filters.
      </div>

      {/* KPI row */}
      <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))" }}>
        {M.hasInv && <Stat label="Unpaid inventory" value={fmtMoney(M.unpaidLiability)} sub={`${M.unpaidCount} watches · excl. memo & voided`} />}
        {M.hasInv && <Stat label="Consignment" value={fmtMoney(M.consignmentLiability)} sub={`${M.consignmentCount} unsold`} />}
        {M.hasSales && <Stat label="Overpaid invoices" value={fmtMoney(M.overpaidLiability)} sub={`${M.overpaidCount} invoices owed back`} />}
        {M.hasSales && <Stat label="Sales tax (est.)" value={fmtMoney(M.salesTaxTotal)} sub="4 ruled states" />}
      </div>

      {/* Unpaid inventory */}
      {M.hasInv && (
        <Panel title="Unpaid inventory liability" note="unpaid watches, excluding memo & voided">
          {!M.hasPaymentStatus
            ? <Locked msg="Add a 'Payment status' column (Paid / Unpaid / Voided) to your inventory export to track this." />
            : M.unpaidItems.length === 0
            ? <div style={{ color: C.green, fontFamily: SANS, fontSize: 13 }}>✓ No unpaid watches. Nothing owed.</div>
            : (<>
                <div style={{ color: C.gold, fontFamily: SERIF }} className="text-2xl mb-3">{fmtMoney(M.unpaidLiability)} owed</div>
                <ItemTable rows={M.unpaidItems} cols={[
                  ["brand","Brand"],["modelName","Model"],["modelNumber","Ref #"],
                  ["invTypeLabel","Type"],["cost","Cost owed",fmtMoney],
                ]} />
              </>)}
        </Panel>
      )}

      {/* Consignment */}
      {M.hasInv && (
        <Panel title="Consignment liability" note="unsold consignment inventory">
          {!M.hasInvType
            ? <Locked msg="Add an 'Inventory type' column (Owned / Consignment / Memo) to your inventory export to track this." />
            : M.consignmentItems.length === 0
            ? <div style={{ color: C.dim, fontFamily: SANS, fontSize: 13 }}>No consignment inventory on hand.</div>
            : (<>
                <div style={{ color: C.gold, fontFamily: SERIF }} className="text-2xl mb-3">{fmtMoney(M.consignmentLiability)} in consignment stock</div>
                <ItemTable rows={M.consignmentItems} cols={[
                  ["brand","Brand"],["modelName","Model"],["modelNumber","Ref #"],
                  ["age","Days in stock"],["cost","Cost",fmtMoney],
                ]} />
              </>)}
        </Panel>
      )}

      {/* Overpaid invoices */}
      {M.hasSales && (
        <Panel title="Overpaid invoices" note="invoices with a negative amount owed (overpayment / trade-in over balance)">
          {!M.hasAmountOwed
            ? <Locked msg="Add a 'Remaining balance' / 'Amount owed' column to your sales export to track overpayments." />
            : M.overpaidItems.length === 0
            ? <div style={{ color: C.green, fontFamily: SANS, fontSize: 13 }}>✓ No overpaid invoices.</div>
            : (<>
                <div style={{ color: C.gold, fontFamily: SERIF }} className="text-2xl mb-3">{fmtMoney(M.overpaidLiability)} owed back to customers</div>
                <ItemTable rows={M.overpaidItems} cols={[
                  ["customer","Customer"],["brand","Brand"],["model","Model"],
                  ["saleDate","Invoice date",fmtDate],["overage","Overpaid",fmtMoney],
                ]} />
              </>)}
        </Panel>
      )}

      {/* Sales tax */}
      {M.hasSales && (
        <Panel title="Sales tax by state" note="estimated · states with a tax rule">
          <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
            Estimated sales tax owed per shipping state. Tax applies to state sales above each state's threshold:
            New York 8% over $500k · California 9.5% over $500k · Florida 7% over $100k · Georgia 7% over $100k.
          </div>
          {!M.hasShippingState && (
            <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
              No shipping-state column detected. Map a "Shipping state" column on your sales export to break this down.
            </div>
          )}
          <div style={{ color: C.gold, fontFamily: SERIF }} className="text-2xl mb-3">{fmtMoney(M.salesTaxTotal)} estimated tax owed</div>
          <ResponsiveContainer width="100%" height={barH(M.salesTaxByState.length, 34, 140)}>
            <BarChart data={M.salesTaxByState} layout="vertical" margin={{ left: 8, right: 24, top: 4, bottom: 4 }}>
              <CartesianGrid stroke={C.line} horizontal={false} />
              <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} tickFormatter={fmtK} />
              <YAxis type="category" dataKey="state" width={yAxisW(M.salesTaxByState, "state", 90)} tick={{ fill: C.dim, fontSize: 11 }} />
              <Tooltip {...chartTip} formatter={(v) => fmtMoney(v)} />
              <Bar dataKey="tax" name="Tax owed" fill={C.gold} radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
          <div className="mt-3">
            <ItemTable rows={M.salesTaxByState} cols={[
              ["state","State"],
              ["taxRate","Rate",(v) => v == null ? "—" : (v * 100).toFixed(1) + "%"],
              ["revenue","Total sales",fmtMoney],
              ["taxableBase","Taxable base",fmtMoney],
              ["tax","Tax owed",fmtMoney],
            ]} />
          </div>
        </Panel>
      )}
    </div>
  );
}

function InventoryTab({ M, filters }) {
  const [sub, setSub] = useState("overview");
  const fByBrand = applyFilters(M.invByBrand, filters);
  const fByLine = applyFilters(M.invByLine, filters);
  const fProjByBrand = applyFilters(M.projByBrand, filters);
  const fProjByModel = applyFilters(M.projByModel, filters);
  const fInvCount = fByBrand.reduce((s, r) => s + (r.count || 0), 0);
  const fInvCost = fByBrand.reduce((s, r) => s + (r.cost || 0), 0);
  const fProjProfit = fProjByModel.reduce((s, r) => s + (r.profit || 0), 0);
  const anyFilterActive = !!(filters.brands || filters.lines || filters.health || filters.stock !== "all");

  // grade & data-quality drill-down state
  const [expandedGrade, setExpandedGrade] = useState(null);
  const [expandedDQ, setExpandedDQ] = useState(null);
  const [expandedAge, setExpandedAge] = useState(null);
  const [phView, setPhView] = useState("model"); // projected-vs-historical: model | brand

  // top watches to sell — exclude/refill
  const [excludedSell, setExcludedSell] = useState(() => new Set());
  const sellRanked = M.needToSellRanked || [];
  const sellRankedKey = sellRanked.map((x) => x._id).join(",");
  useEffect(() => { setExcludedSell(new Set()); }, [sellRankedKey]);
  const visibleSell = sellRanked.filter((x) => !excludedSell.has(x._id)).slice(0, 10);
  const excludedSellItems = sellRanked.filter((x) => excludedSell.has(x._id));

  const subTabs = [
    ["overview", "Overview"],
    ["grading", "Grading & Aging"],
    ["market", "Market Grades"],
    ["quality", "Data Quality"],
    ["suppliers", "Suppliers"],
    ["projections", "Projections"],
  ];

  return (
    <div className="flex flex-col gap-5">
      {/* ── KPIs (always visible) ── */}
      <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))" }}>
        <Stat label="Items in stock" value={fInvCount} />
        <Stat label="Capital tied up" value={fmtMoney(fInvCost)} />
        <Stat label="Brands" value={fByBrand.length} />
        <Stat label="Aged 91+ days" value={fmtMoney(M.agedValue)} sub={anyFilterActive ? "all brands · not filterable" : "cost sitting on the shelf"} />
        {M.projItems > 0 && <Stat label="Projected profit" value={fmtMoney(fProjProfit)} sub={`${fProjByModel.length} priced items`} />}
      </div>

      <SubTabs tabs={subTabs} value={sub} onChange={setSub} />

      {/* ════ OVERVIEW ════ */}
      {sub === "overview" && (<>
        <Panel title="Inventory by brand" note="cost on the shelf">
          <ResponsiveContainer width="100%" height={barH(fByBrand.length)}>
            <BarChart data={fByBrand} layout="vertical" margin={{ left: 8, right: 24, top: 4, bottom: 4 }}>
              <CartesianGrid stroke={C.line} horizontal={false} />
              <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} tickFormatter={fmtK} />
              <YAxis type="category" dataKey="brand" width={yAxisW(fByBrand)} tick={{ fill: C.dim, fontSize: 11 }} />
              <Tooltip {...chartTip} formatter={(v) => fmtMoney(v)} />
              <Bar dataKey="cost" name="Cost" fill={C.gold} radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Panel>

        <Panel title="By product line" note="Rolex · Patek Philippe · Audemars Piguet · Omega">
          {fByLine.length === 0
            ? <Locked msg="No items matched a known product line for the four focus brands." />
            : (<>
              <ResponsiveContainer width="100%" height={barH(fByLine.length, 34, 160)}>
                <BarChart data={fByLine} layout="vertical" margin={{ left: 8, right: 24, top: 4, bottom: 4 }}>
                  <CartesianGrid stroke={C.line} horizontal={false} />
                  <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} tickFormatter={fmtK} />
                  <YAxis type="category" dataKey="line" width={yAxisW(fByLine, "line")} tick={{ fill: C.dim, fontSize: 11 }} />
                  <Tooltip {...chartTip} formatter={(v, n) => n === "Cost" ? fmtMoney(v) : v} />
                  <Bar dataKey="cost" name="Cost" fill={C.blue} radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
              <div className="mt-3">
                <ItemTable rows={fByLine} cols={[["brand","Brand"],["line","Line"],["count","Items"],["cost","Cost",fmtMoney]]} />
              </div>
            </>)}
        </Panel>

        <Panel title="By brand & condition" note={M.invHasCondition ? "new vs. used" : "needs a condition column on inventory"}>
          {!M.invHasCondition
            ? <Locked msg="Add a 'Condition' column (New / Used) to your inventory export to break this down." />
            : <ItemTable rows={applyFilters(M.invByBrandCondition, filters)} cols={[
                ["brand","Brand"],["condition","Condition"],["count","Items"],["cost","Cost",fmtMoney],
              ]} />}
        </Panel>
      </>)}

      {/* ════ GRADING & AGING ════ */}
      {sub === "grading" && (<>
        <Panel title="Inventory quality grade (A → F)" note="A = just bought, F = sitting longest">
          <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
            Every item is graded by how long it's been in stock: A (0–30 days), B (31–60), C (61–90), D (91–180), F (180+).
            Grades naturally fall as a watch sits longer without selling. Click a grade to see the watches in it.
          </div>
          <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))" }}>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={M.invByGrade} margin={{ top: 4, right: 8, bottom: 4, left: 0 }}>
                <CartesianGrid stroke={C.line} vertical={false} />
                <XAxis dataKey="grade" tick={{ fill: C.dim, fontSize: 11 }} />
                <YAxis tick={{ fill: C.faint, fontSize: 11 }} />
                <Tooltip {...chartTip} formatter={(v, n) => n === "Cost" ? fmtMoney(v) : v} />
                <Bar dataKey="count" name="Items" radius={[4, 4, 0, 0]} cursor="pointer"
                  onClick={(d) => { const g = d?.payload?.grade ?? d?.grade; if (g) setExpandedGrade((p) => p === g ? null : g); }}>
                  {M.invByGrade.map((b, i) => <Cell key={i} fill={(GRADE_CFG[b.grade] || {}).bg || C.gold} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
            <ItemTable rows={M.invByGrade}
              getRowKey={(r) => r.grade}
              expandedKey={expandedGrade}
              onRowClick={(r) => setExpandedGrade((p) => p === r.grade ? null : r.grade)}
              renderExpanded={(r) => (
                r.items.length === 0
                  ? <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>No watches in grade {r.grade}.</div>
                  : <div>
                      <div style={{ color: C.gold, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }} className="mb-1">
                        Grade {r.grade} — {r.items.length} {r.items.length === 1 ? "watch" : "watches"}
                      </div>
                      <ItemTable rows={r.items} cols={WATCH_COLS} />
                    </div>
              )}
              cols={[
                ["grade","Grade",(v) => <GradeBadge grade={v} />],
                ["count","Items"],["cost","Cost tied up",fmtMoney],
              ]} />
          </div>
        </Panel>

        <Panel title="Age of inventory" note="days stock held · click a bucket to see its watches">
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={M.aging} margin={{ top: 4, right: 8, bottom: 4, left: 0 }}>
              <CartesianGrid stroke={C.line} vertical={false} />
              <XAxis dataKey="label" tick={{ fill: C.dim, fontSize: 11 }} />
              <YAxis tick={{ fill: C.faint, fontSize: 11 }} />
              <Tooltip {...chartTip} formatter={(v, n) => n === "Cost" ? fmtMoney(v) : v} />
              <Bar dataKey="count" name="Items" radius={[4, 4, 0, 0]} cursor="pointer"
                onClick={(d) => { const l = d?.payload?.label ?? d?.label; if (l) setExpandedAge((p) => p === l ? null : l); }}>
                {M.aging.map((b, i) => <Cell key={i} fill={b.lo >= 91 ? C.red : b.lo >= 61 ? "#c8863a" : C.gold} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
          <div className="mt-3">
            <ItemTable rows={M.aging}
              getRowKey={(r) => r.label}
              expandedKey={expandedAge}
              onRowClick={(r) => setExpandedAge((p) => p === r.label ? null : r.label)}
              renderExpanded={(r) => (
                r.items.length === 0
                  ? <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>No watches in this age range.</div>
                  : <div>
                      <div style={{ color: C.gold, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }} className="mb-1">
                        {r.label} days — {r.items.length} {r.items.length === 1 ? "watch" : "watches"}
                      </div>
                      <ItemTable rows={[...r.items].sort((a, b) => b.age - a.age)} cols={WATCH_COLS} />
                    </div>
              )}
              cols={[
                ["label","Age bucket",(v) => v + " days"],
                ["count","Items"],["cost","Cost tied up",fmtMoney],
              ]} />
            {M.agedValue > 0 && (
              <div style={{ color: C.red, fontFamily: SANS, fontSize: 12, marginTop: 10 }}>
                ⚠ {fmtMoney(M.agedValue)} sitting 91+ days
              </div>
            )}
          </div>
        </Panel>

        <Panel title="Top 10 watches to sell" note="worst grade + most cash tied up first">
          {visibleSell.length > 0 ? (
            <>
              <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
                These are holding the most cash for the longest. Moving them frees up capital for faster, more profitable watches.
                Click "Exclude" to drop one — the next-ranked watch takes its place.
              </div>
              <ItemTable rows={visibleSell} getRowKey={(r) => r._id} cols={[
                ["brand","Brand"],["modelName","Model"],["modelNumber","Ref #"],
                ["grade","Grade",(v) => <GradeBadge grade={v} />],
                ["age","Days in stock"],["cost","Cost",fmtMoney],
                ["_id","",(_, r) => (
                  <button onClick={() => setExcludedSell((prev) => new Set(prev).add(r._id))}
                    style={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${C.line}`, background: "transparent", color: C.dim }}
                    className="px-3 py-1">Exclude</button>
                )],
              ]} />
              {excludedSellItems.length > 0 && (
                <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }} className="mt-3">
                  <div className="mb-1">Excluded:</div>
                  <div className="flex flex-wrap gap-2">
                    {excludedSellItems.map((x) => (
                      <button key={x._id}
                        onClick={() => setExcludedSell((prev) => { const n = new Set(prev); n.delete(x._id); return n; })}
                        style={{ fontFamily: SANS, fontSize: 12, borderRadius: 999, border: `1px solid ${C.line}`, background: "transparent", color: C.dim }}
                        className="px-3 py-1">{x.brand} — {x.modelName || x.modelNumber || "watch"} ✕</button>
                    ))}
                  </div>
                </div>
              )}
            </>
          ) : <Locked msg="No graded inventory items available." />}
        </Panel>
      </>)}

      {/* ════ DATA QUALITY ════ */}
      {sub === "quality" && (
        <Panel title="Data quality" note="missing fields by watch">
          <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
            How complete is each watch record. Click a field to see which watches are missing it, so you can clean up the source data.
          </div>
          <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))" }}>
            <ResponsiveContainer width="100%" height={barH(M.dataQuality.length, 34, 180)}>
              <BarChart data={M.dataQuality} layout="vertical" margin={{ left: 8, right: 24, top: 4, bottom: 4 }}>
                <CartesianGrid stroke={C.line} horizontal={false} />
                <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} />
                <YAxis type="category" dataKey="field" width={yAxisW(M.dataQuality, "field", 120)} tick={{ fill: C.dim, fontSize: 11 }} />
                <Tooltip {...chartTip} />
                <Bar dataKey="missing" name="Missing" fill={C.red} radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
            <ItemTable rows={M.dataQuality}
              getRowKey={(r) => r.key}
              expandedKey={expandedDQ}
              onRowClick={(r) => setExpandedDQ((p) => p === r.key ? null : r.key)}
              renderExpanded={(r) => (
                r.items.length === 0
                  ? <div style={{ color: C.green, fontFamily: SANS, fontSize: 12 }}>✓ All watches have {r.field.toLowerCase()}.</div>
                  : <div>
                      <div style={{ color: C.gold, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }} className="mb-1">
                        Missing {r.field.toLowerCase()} — {r.items.length} {r.items.length === 1 ? "watch" : "watches"}
                      </div>
                      <ItemTable rows={r.items} cols={WATCH_COLS} />
                    </div>
              )}
              cols={[
                ["field","Field"],["missing","Missing"],["present","Present"],
              ]} />
          </div>
        </Panel>
      )}

      {/* ════ MARKET GRADES ════ */}
      {sub === "market" && (
        <Panel title="Oldest 20 in stock — vs WatchCharts market" note="target & cost graded against median asking price">
          <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
            Your 20 longest-held watches, compared to WatchCharts' <b>median asking price</b>.
            <b> QoC</b> (Quality of Cost) grades what you paid vs market — A = paid under 90% of market, D = paid 100%+.
            <b> QoT</b> (Quality of Target) grades how close your target price is to market — A = within 3%, down to D = off by more than 15%.
          </div>
          {(M.oldest20 || []).length === 0
            ? <Locked msg="No inventory with age data to grade." />
            : <OldestMarketTable rows={M.oldest20} />}
        </Panel>
      )}

      {/* ════ SUPPLIERS ════ */}
      {sub === "suppliers" && (
        <Panel title="Top suppliers" note={M.hasSupplier ? "by spend & by profit" : "needs a 'purchased from' / supplier column"}>
          {!M.hasSupplier
            ? <Locked msg="Add a 'Supplier / vendor' (or 'Purchased from') column to your export to rank who you buy from." />
            : (
              <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))" }}>
                <div>
                  <SectionLabel>By spend (purchase cost)</SectionLabel>
                  {M.hasSupplierSpend
                    ? <ItemTable rows={M.suppliersBySpend} cols={[
                        ["supplier","Supplier"],["units","Watches"],["spend","Spend",fmtMoney],
                      ]} />
                    : <Locked msg="No supplier spend data available." />}
                </div>
                <div>
                  <SectionLabel>By profit</SectionLabel>
                  {M.hasSupplierProfit
                    ? <ItemTable rows={M.suppliersByProfit} cols={[
                        ["supplier","Supplier"],["units","Watches"],["profit","Profit $",fmtMoney],
                      ]} />
                    : <Locked msg="Profit by supplier needs a supplier column on your sales export." />}
                </div>
              </div>
            )}
        </Panel>
      )}

      {/* ════ PROJECTIONS ════ */}
      {sub === "projections" && (<>
        {M.projItems === 0
          ? <Panel title="Projected profit"><Locked msg="No items have a target wholesale price yet. Add 'Target Wholesale Price' to your inventory export." /></Panel>
          : fProjByModel.length === 0
          ? <Panel title="Projected profit"><Locked msg="No priced items match the current brand filter." /></Panel>
          : (<>
            <Panel title="Projected profit by brand" note={`${fProjByModel.length} of ${M.projItems} priced items`}>
              <div style={{ color: C.gold, fontFamily: SERIF }} className="text-2xl mb-3">{fmtMoney(fProjProfit)} total</div>
              <ResponsiveContainer width="100%" height={barH(fProjByBrand.length, 38, 140)}>
                <BarChart data={fProjByBrand} layout="vertical" margin={{ left: 8, right: 24, top: 4, bottom: 4 }}>
                  <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} tickFormatter={fmtK} />
                  <YAxis type="category" dataKey="brand" width={yAxisW(fProjByBrand)} tick={{ fill: C.dim, fontSize: 11 }} />
                  <Tooltip {...chartTip} formatter={(v) => fmtMoney(v)} />
                  <Bar dataKey="profit" name="Projected profit" fill={C.green} radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </Panel>

            <Panel title="Projected profit by model" note="based on target wholesale price">
              <ItemTable rows={fProjByModel.slice(0, 30)} cols={[
                ["brand","Brand"],["model","Model"],["count","Items"],
                ["profit","Proj. Profit",fmtMoney],
                ["marginPct","Margin %",(v) => fmtPct(v)],
              ]} />
            </Panel>
          </>)}

        {/* Projected vs historical health */}
        <Panel title="Projected vs historical profit" note="how current stock's projected margin compares to what you've actually realized">
          {!M.hasProjVsHist
            ? <Locked msg="Needs both inventory (with target wholesale prices) and sales history loaded." />
            : (<>
                <div className="flex gap-1 mb-3">
                  {[["model","By model"],["brand","By brand"]].map(([k, lbl]) => (
                    <button key={k} onClick={() => setPhView(k)}
                      style={{ fontFamily: SANS, borderRadius: 8, fontSize: 12,
                        border: `1px solid ${phView === k ? C.gold : C.line}`,
                        background: phView === k ? C.gold : "transparent",
                        color: phView === k ? C.bg : C.dim }} className="px-3 py-1">{lbl}</button>
                  ))}
                </div>
                <ItemTable rows={(phView === "model" ? M.projVsHistModels : M.projVsHistBrands).slice(0, 40)} cols={[
                  ["brand","Brand"],
                  ...(phView === "model" ? [["model","Model"]] : []),
                  ["stock","In stock"],["sold","Sold"],
                  ["projProfit","Projected $",fmtMoney],["histProfit","Historical $",fmtMoney],
                  ["projMargin","Proj. margin",(v) => fmtPct(v)],["histMargin","Hist. margin",(v) => fmtPct(v)],
                  ["delta","Δ margin",(v) => v == null ? "—" : <span style={{ color: v >= 0 ? C.green : C.red }}>{(v >= 0 ? "+" : "") + v.toFixed(1) + "%"}</span>],
                ]} />
              </>)}
        </Panel>
      </>)}
    </div>
  );
}

function SalesTab({ M, filters }) {
  const needsBrand = !M.salesHasBrand;
  const fBrand = applyFilters(M.salesByBrand, filters);
  const fLine = applyFilters(M.salesByLine, filters);
  const fModel = applyFilters(M.salesByModel, filters);
  const allModels = applyFilters(M.salesByModel, filters);
  const [expandedLine, setExpandedLine] = useState(null);
  const [expandedBrand, setExpandedBrand] = useState(null);
  const [expandedVelBrand, setExpandedVelBrand] = useState(null);
  const [expandedPerson, setExpandedPerson] = useState(null);
  const lineTableRef = useRef(null);
  const handleLineBarClick = (data) => {
    const line = data?.payload?.line ?? data?.line;
    if (!line) return;
    setExpandedLine(line);
    lineTableRef.current?.scrollIntoView({ behavior: "smooth", block: "center" });
  };
  const modelsForBrand = (brand) => allModels.filter((mo) => (mo.brand || "Unknown") === brand);
  const brandModelCols = [
    ["model","Model"],["units","Units"],["profit","Profit $",fmtMoney],
    ["avgProfit","Avg $",fmtMoney],["profitPct","Margin %",fmtPct],["medianDays","Median days"],
  ];
  const renderBrandModels = (r) => {
    const models = modelsForBrand(r.brand);
    return models.length === 0
      ? <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>No model-level data for {r.brand}.</div>
      : <div>
          <div style={{ color: C.gold, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }} className="mb-1">
            {r.brand} — by model
          </div>
          <ItemTable rows={[...models].sort((a, b) => b.profit - a.profit)} cols={brandModelCols} />
        </div>;
  };
  const fVelocity = applyFilters(M.byVelocity, filters);
  const [sub, setSub] = useState("performance");
  const salesSubTabs = [
    ["performance", "Performance"],
    ["breakdowns", "Breakdowns"],
    ["salespeople", "Salespeople"],
    ["market", "Market Grades"],
  ];
  return (
    <div className="flex flex-col gap-5">
      {M.windowMode === "window" && (
        <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>
          Showing the last {M.salesWindowDays} days of sales (from today). Every section below — including Salespeople and Tax —
          reflects this window. Switch the Period filter to "All time" to see everything.
        </div>
      )}
      {/* ── KPIs ── */}
      <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))" }}>
        <Stat label="Units sold" value={M.salesUnits} />
        <Stat label="Total profit $" value={fmtMoney(M.salesProfit)} />
        <Stat label="Revenue" value={fmtMoney(M.salesRevenue)} />
        <Stat label="Median margin %" value={fmtPct(M.medianMargin)} sub="on cost" />
        <Stat label="Median days to sell" value={M.medianDays ?? "--"} sub={M.meanDays ? `avg ${Math.round(M.meanDays)} days` : ""} />
      </div>

      <SubTabs tabs={salesSubTabs} value={sub} onChange={setSub} />

      {/* ════ PERFORMANCE ════ */}
      {sub === "performance" && (<>
      {/* ── By Brand ── */}
      <Panel title="By brand" note={needsBrand ? "needs brand column on sales" : "profit $ and margin %"}>
        {needsBrand
          ? <Locked msg="Add a brand or model/reference column to your sales export and this section lights up." />
          : (<>
            <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))" }}>
              <div>
                <SectionLabel>Profit $</SectionLabel>
                <ResponsiveContainer width="100%" height={barH(fBrand.length, 36, 140)}>
                  <BarChart data={fBrand} layout="vertical" margin={{ left: 8, right: 24, top: 2, bottom: 2 }}>
                    <CartesianGrid stroke={C.line} horizontal={false} />
                    <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} tickFormatter={fmtK} />
                    <YAxis type="category" dataKey="brand" width={yAxisW(fBrand)} tick={{ fill: C.dim, fontSize: 11 }} />
                    <Tooltip {...chartTip} formatter={(v) => fmtMoney(v)} />
                    <Bar dataKey="profit" name="Profit" fill={C.green} radius={[0, 4, 4, 0]} cursor="pointer"
                      onClick={(d) => { const b = d?.payload?.brand ?? d?.brand; if (b) setExpandedBrand((p) => p === b ? null : b); }} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div>
                <SectionLabel>Margin %</SectionLabel>
                <ResponsiveContainer width="100%" height={barH(fBrand.length, 36, 140)}>
                  <BarChart data={[...fBrand].sort((a,b) => (b.profitPct||0)-(a.profitPct||0))}
                    layout="vertical" margin={{ left: 8, right: 24, top: 2, bottom: 2 }}>
                    <CartesianGrid stroke={C.line} horizontal={false} />
                    <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} tickFormatter={(v) => v.toFixed(0) + "%"} />
                    <YAxis type="category" dataKey="brand" width={yAxisW(fBrand)} tick={{ fill: C.dim, fontSize: 11 }} />
                    <Tooltip {...chartTip} formatter={(v) => fmtPct(v)} />
                    <Bar dataKey="profitPct" name="Margin %" fill={C.blue} radius={[0, 4, 4, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="mt-3">
              <div style={{ color: C.faint, fontFamily: SANS, fontSize: 11 }} className="mb-1">Click a brand (row or bar) to see its models</div>
              <ItemTable rows={fBrand}
                getRowKey={(r) => r.brand}
                expandedKey={expandedBrand}
                onRowClick={(r) => setExpandedBrand((p) => p === r.brand ? null : r.brand)}
                renderExpanded={renderBrandModels}
                cols={[
                  ["brand","Brand"],["units","Units"],["profit","Profit $",fmtMoney],
                  ["profitPct","Margin %",fmtPct],["medianDays","Median days"],
                ]} />
            </div>
          </>)}
      </Panel>
      </>)}

      {/* ════ SALESPEOPLE ════ */}
      {sub === "salespeople" && (
        <Panel title="By salesperson" note={M.hasSalesperson ? (M.windowMode === "window" ? `profit · sales · velocity · last ${M.salesWindowDays} days` : "profit · sales · velocity") : "needs a 'created by' / salesperson column"}>
        {!M.hasSalesperson
          ? <Locked msg="Add a 'Created by' (salesperson) column to your sales export to rank who sold what." />
          : (<>
            <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))" }}>
              <div>
                <SectionLabel>Profit $</SectionLabel>
                <ResponsiveContainer width="100%" height={barH(M.salesByPerson.length, 36, 120)}>
                  <BarChart data={[...M.salesByPerson].sort((a,b) => b.profit - a.profit)} layout="vertical" margin={{ left: 8, right: 24, top: 2, bottom: 2 }}>
                    <CartesianGrid stroke={C.line} horizontal={false} />
                    <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} tickFormatter={fmtK} />
                    <YAxis type="category" dataKey="salesperson" width={yAxisW(M.salesByPerson, "salesperson", 90)} tick={{ fill: C.dim, fontSize: 11 }} />
                    <Tooltip {...chartTip} formatter={(v) => fmtMoney(v)} />
                    <Bar dataKey="profit" name="Profit" fill={C.green} radius={[0, 4, 4, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div>
                <SectionLabel>Sales $</SectionLabel>
                <ResponsiveContainer width="100%" height={barH(M.salesByPerson.length, 36, 120)}>
                  <BarChart data={[...M.salesByPerson].sort((a,b) => b.revenue - a.revenue)} layout="vertical" margin={{ left: 8, right: 24, top: 2, bottom: 2 }}>
                    <CartesianGrid stroke={C.line} horizontal={false} />
                    <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} tickFormatter={fmtK} />
                    <YAxis type="category" dataKey="salesperson" width={yAxisW(M.salesByPerson, "salesperson", 90)} tick={{ fill: C.dim, fontSize: 11 }} />
                    <Tooltip {...chartTip} formatter={(v) => fmtMoney(v)} />
                    <Bar dataKey="revenue" name="Sales" fill={C.gold} radius={[0, 4, 4, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="mt-3">
              <div style={{ color: C.faint, fontFamily: SANS, fontSize: 11 }} className="mb-1">
                Click a salesperson to see each sale they made, graded by <b>DJ's grade</b> (margin + days-on-hand + gross profit). Their letter is the average across their sales.
              </div>
              <ItemTable rows={M.salesByPerson}
                getRowKey={(r) => r.salesperson}
                expandedKey={expandedPerson}
                onRowClick={(r) => setExpandedPerson((p) => p === r.salesperson ? null : r.salesperson)}
                renderExpanded={(r) => (
                  (r.sales && r.sales.length) ? (
                    <div>
                      <div style={{ color: C.gold, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }} className="mb-1">
                        {r.salesperson} — {r.sales.length} sales · avg DJ grade <GradeBadge grade={r.djGrade} /> {r.djPoints != null ? `(${r.djPoints})` : ""}
                      </div>
                      <ItemTable rows={r.sales.slice(0, 200)} cols={[
                        ["saleDate","Date",(v) => v ? new Date(v).toISOString().slice(0, 10) : "—"],
                        ["brand","Brand"],["model","Model"],["reference","Ref #"],
                        ["price","Sold for",fmtMoney],["profit","Profit $",fmtMoney],
                        ["marginPct","Margin %",(v) => fmtPct(v)],["days","Days on hand"],
                        ["djGrade","DJ grade",(v) => <GradeBadge grade={v} />],
                      ]} />
                    </div>
                  ) : <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>No individual sales recorded.</div>
                )}
                cols={[
                  ["salesperson","Salesperson"],
                  ["djGrade","DJ grade",(v) => <GradeBadge grade={v} />],
                  ["units","Units"],
                  ["revenue","Sales $",fmtMoney],["profit","Profit $",fmtMoney],
                  ["profitPct","Margin %",fmtPct],["medianDays","Velocity (median days)"],
                ]} />
            </div>
          </>)}
        </Panel>
      )}

      {/* ════ BREAKDOWNS ════ */}
      {sub === "breakdowns" && (<>
      {/* ── By Product Line ── */}
      <Panel title="By product line" note={needsBrand ? "needs brand column" : "Rolex · Patek · AP · Omega"}>
        {needsBrand
          ? <Locked msg="Requires brand and model columns on your sales export." />
          : (fLine.length > 0
            ? (<>
              <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))" }}>
                <div>
                  <SectionLabel>Profit $ by line</SectionLabel>
                  <ResponsiveContainer width="100%" height={barH(fLine.length, 34, 140)}>
                    <BarChart data={fLine} layout="vertical" margin={{ left: 8, right: 24, top: 2, bottom: 2 }}>
                      <CartesianGrid stroke={C.line} horizontal={false} />
                      <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} tickFormatter={fmtK} />
                      <YAxis type="category" dataKey="line" width={yAxisW(fLine, "line", 90)} tick={{ fill: C.dim, fontSize: 11 }} />
                      <Tooltip {...chartTip} formatter={(v) => fmtMoney(v)} />
                      <Bar dataKey="profit" name="Profit" fill={C.gold} radius={[0, 4, 4, 0]} cursor="pointer" onClick={handleLineBarClick} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                <div>
                  <SectionLabel>Velocity by line (median days)</SectionLabel>
                  <ResponsiveContainer width="100%" height={barH(fLine.length, 34, 140)}>
                    <BarChart data={[...fLine].filter(x=>x.medianDays!=null).sort((a,b)=>a.medianDays-b.medianDays)}
                      layout="vertical" margin={{ left: 8, right: 24, top: 2, bottom: 2 }}>
                      <CartesianGrid stroke={C.line} horizontal={false} />
                      <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} unit=" d" />
                      <YAxis type="category" dataKey="line" width={yAxisW(fLine, "line", 90)} tick={{ fill: C.dim, fontSize: 11 }} />
                      <Tooltip {...chartTip} formatter={(v) => v + " days"} />
                      <Bar dataKey="medianDays" name="Median days to sell" fill={C.blue} radius={[0, 4, 4, 0]} cursor="pointer" onClick={handleLineBarClick} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="mt-3" ref={lineTableRef}>
                <div style={{ color: C.faint, fontFamily: SANS, fontSize: 11 }} className="mb-1">
                  Click a bar above, or a row below, to see its model breakdown
                </div>
                <ItemTable rows={fLine}
                  getRowKey={(r) => r.line}
                  expandedKey={expandedLine}
                  onRowClick={(r) => setExpandedLine((prev) => (prev === r.line ? null : r.line))}
                  renderExpanded={(r) => {
                    const models = allModels.filter((mo) => mo.line === r.line);
                    return models.length === 0
                      ? <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>No model-level sales data for {r.line}.</div>
                      : (
                        <div>
                          <div style={{ color: C.gold, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }} className="mb-1">
                            {r.line} — by model
                          </div>
                          <ItemTable rows={models} cols={[
                            ["model","Model / Ref #"],["units","Units sold"],
                            ["profit","Total profit $",fmtMoney],["avgProfit","Avg profit $",fmtMoney],
                            ["profitPct","Margin %",fmtPct],["medianDays","Median days"],
                            ["health","Stock health",(v,r2) => <HealthBadge health={v} weeksOfStock={r2?.weeksOfStock} />],
                          ]} />
                        </div>
                      );
                  }}
                  cols={[
                    ["brand","Brand"],["line","Line"],["units","Units"],
                    ["profit","Profit $",fmtMoney],["profitPct","Margin %",fmtPct],["medianDays","Median days"],
                    ["health","Stock health",(v,r) => <HealthBadge health={v} weeksOfStock={r?.weeksOfStock} />],
                  ]} />
              </div>
            </>)
            : <Locked msg="No product-line matches found. Lines are detected from Rolex, Patek Philippe, AP, and Omega model names." />
          )}
      </Panel>

      {/* ── By Model ── */}
      <Panel title="By model number" note={needsBrand ? "needs model column" : `${fModel.length} models`}>
        {needsBrand
          ? <Locked msg="Add a model/reference column to your sales export." />
          : (<>
            {filters.lines && (
              <div className="mb-2 flex items-center gap-2 flex-wrap">
                <span style={{ color: C.gold, fontFamily: SANS, fontSize: 12 }}>
                  Showing models for: {Array.from(filters.lines).join(", ")}
                </span>
              </div>
            )}
            <ItemTable rows={fModel.slice(0, 40)} cols={[
              ["brand","Brand"],["model","Model / Ref #"],["units","Units sold"],
              ["profit","Total profit $",fmtMoney],["avgProfit","Avg profit $",fmtMoney],
              ["profitPct","Margin %",fmtPct],["medianDays","Median days"],
              ["health","Stock health",(v,r) => <HealthBadge health={v} weeksOfStock={r?.weeksOfStock} />],
            ]} />
          </>)}
      </Panel>
      </>)}

      {/* ════ PERFORMANCE (velocity) ════ */}
      {sub === "performance" && (<>
      {/* ── Velocity ── */}
      <Panel title="Velocity — how quickly do we sell?" note="median days to sell">
        {needsBrand
          ? <Locked msg="Requires brand and model columns on sales." />
          : (<>
            <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))" }}>
              <div>
                <SectionLabel>By brand (median days)</SectionLabel>
                <ResponsiveContainer width="100%" height={barH(fBrand.filter(x=>x.medianDays!=null).length, 36, 120)}>
                  <BarChart
                    data={[...fBrand].filter(x=>x.medianDays!=null).sort((a,b)=>a.medianDays-b.medianDays)}
                    layout="vertical" margin={{ left: 8, right: 24, top: 2, bottom: 2 }}>
                    <CartesianGrid stroke={C.line} horizontal={false} />
                    <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} unit=" d" />
                    <YAxis type="category" dataKey="brand" width={yAxisW(fBrand)} tick={{ fill: C.dim, fontSize: 11 }} />
                    <Tooltip {...chartTip} formatter={(v) => v + " days"} />
                    <Bar dataKey="medianDays" name="Median days" fill={C.gold} radius={[0, 4, 4, 0]} cursor="pointer"
                      onClick={(d) => { const b = d?.payload?.brand ?? d?.brand; if (b) setExpandedVelBrand((p) => p === b ? null : b); }} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div>
                <SectionLabel>Top 15 fastest models</SectionLabel>
                <ItemTable rows={fVelocity.slice(0, 15)} cols={[
                  ["brand","Brand"],["model","Model"],["medianDays","Days"],["units","Units"],
                ]} />
              </div>
            </div>
            <div className="mt-3">
              <div style={{ color: C.faint, fontFamily: SANS, fontSize: 11 }} className="mb-1">Click a brand (row or bar) to see its models, fastest first</div>
              <ItemTable rows={[...fBrand].filter(x=>x.medianDays!=null).sort((a,b)=>a.medianDays-b.medianDays)}
                getRowKey={(r) => r.brand}
                expandedKey={expandedVelBrand}
                onRowClick={(r) => setExpandedVelBrand((p) => p === r.brand ? null : r.brand)}
                renderExpanded={(r) => {
                  const models = modelsForBrand(r.brand).filter((mo) => mo.medianDays != null).sort((a, b) => a.medianDays - b.medianDays);
                  return models.length === 0
                    ? <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>No model velocity data for {r.brand}.</div>
                    : <div>
                        <div style={{ color: C.gold, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }} className="mb-1">
                          {r.brand} — models by velocity
                        </div>
                        <ItemTable rows={models} cols={[
                          ["model","Model"],["medianDays","Median days"],["units","Units"],["profit","Profit $",fmtMoney],
                        ]} />
                      </div>;
                }}
                cols={[
                  ["brand","Brand"],["medianDays","Median days"],["units","Units"],
                ]} />
            </div>
          </>)}
      </Panel>
      </>)}

      {/* ════ BREAKDOWNS (continued) ════ */}
      {sub === "breakdowns" && (<>
      {/* ── Inventory type breakdown ── */}
      <Panel title="By inventory type">
        <ItemTable rows={M.salesByType} cols={[["type","Type"],["units","Units"],["profit","Profit $",fmtMoney]]} />
      </Panel>

      {/* ── By brand & condition ── */}
      <Panel title="By brand & condition" note={M.hasCondition ? "new vs. used" : "needs a condition column on sales"}>
        {!M.hasCondition
          ? <Locked msg="Add a 'Condition' column (New / Used) to your sales export to break this down." />
          : <ItemTable rows={applyFilters(M.salesByBrandCondition, filters)} cols={[
              ["brand","Brand"],["condition","Condition"],["units","Units"],
              ["profit","Profit $",fmtMoney],["profitPct","Margin %",fmtPct],
            ]} />}
      </Panel>

      {/* ── By price tier ── */}
      <Panel title="By price tier" note="sale price ranges">
        <ResponsiveContainer width="100%" height={barH(M.salesByPriceTier.length, 30, 160)}>
          <BarChart data={M.salesByPriceTier} layout="vertical" margin={{ left: 8, right: 24, top: 4, bottom: 4 }}>
            <CartesianGrid stroke={C.line} horizontal={false} />
            <XAxis type="number" tick={{ fill: C.faint, fontSize: 11 }} />
            <YAxis type="category" dataKey="tier" width={yAxisW(M.salesByPriceTier, "tier", 80)} tick={{ fill: C.dim, fontSize: 11 }} />
            <Tooltip {...chartTip} formatter={(v, n) => n === "Profit $" ? fmtMoney(v) : v} />
            <Bar dataKey="units" name="Units" fill={C.gold} radius={[0, 4, 4, 0]} />
          </BarChart>
        </ResponsiveContainer>
        <div className="mt-3">
          <ItemTable rows={M.salesByPriceTier} cols={[
            ["tier","Price tier"],["units","Units"],["revenue","Revenue",fmtMoney],
            ["profit","Profit $",fmtMoney],["avgProfit","Avg profit $",fmtMoney],["profitPct","Margin %",fmtPct],
          ]} />
        </div>
      </Panel>

      </>)}

      {/* ════ MARKET GRADES ════ */}
      {sub === "market" && (
        <Panel title="Latest 20 sales — vs WatchCharts market" note="actual sale price graded against median asking price">
          <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
            Your 20 most recent sales, compared to WatchCharts' <b>median asking price</b>.
            Sales grade: <b>A</b> = sold at 100%+ of market, <b>B</b> = 95%+, <b>C</b> = 90%+, <b>D</b> = below.
          </div>
          {(M.latest20 || []).length === 0
            ? <Locked msg="No dated sales to grade." />
            : <LatestMarketTable rows={M.latest20} />}
        </Panel>
      )}

    </div>
  );
}

function QuestionCard({ num, question, children }) {
  return (
    <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 14 }} className="p-4">
      <div className="flex items-start gap-3 mb-3">
        <div style={{
          background: C.gold, color: C.bg, fontFamily: SANS, fontWeight: 700,
          borderRadius: "50%", width: 26, height: 26, display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 13, flexShrink: 0,
        }}>{num}</div>
        <div style={{ color: C.text, fontFamily: SERIF }} className="text-lg leading-snug">{question}</div>
      </div>
      {children}
    </div>
  );
}

/* generic ranked table for model/line rows */
function RankTable({ rows, nameKey, showScore, models, onIgnore }) {
  const [expanded, setExpanded] = useState(null);
  const cols = [
    ["brand","Brand"],
    [nameKey, nameKey === "model" ? "Model" : "Line"],
    ["units","Units"],
    ["profit","Profit $",fmtMoney],
    ["avgProfit","Avg $",fmtMoney],
    ["profitPct","Margin %",fmtPct],
    ["medianDays","Days"],
    ["stock","Stock",(v) => v == null ? "—" : v === 0 ? "⚡ OUT" : String(v)],
  ];
  if (showScore) cols.push(["buyScore","Score",(v) => v?.toFixed(3)]);
  cols.push(["health","Stock health",(v,r) => <HealthBadge health={v} weeksOfStock={r?.weeksOfStock} />]);
  if (onIgnore) cols.push(["__ignore","",(_, r) => (
    <button onClick={(e) => { e.stopPropagation(); onIgnore(r); }}
      style={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${C.line}`, background: "transparent", color: C.dim, whiteSpace: "nowrap" }}
      className="px-3 py-1">Ignore</button>
  )]);

  if (nameKey === "line" && models) {
    const subCols = [
      ["model","Model / Ref #"],["units","Units"],
      ["profit","Profit $",fmtMoney],["avgProfit","Avg $",fmtMoney],["profitPct","Margin %",fmtPct],
      ["medianDays","Days"],["stock","Stock",(v) => v == null ? "—" : v === 0 ? "⚡ OUT" : String(v)],
    ];
    if (showScore) subCols.push(["buyScore","Score",(v) => v?.toFixed(3)]);
    subCols.push(["health","Stock health",(v,r) => <HealthBadge health={v} weeksOfStock={r?.weeksOfStock} />]);
    return (
      <ItemTable rows={rows} cols={cols}
        getRowKey={(r) => r.line}
        expandedKey={expanded}
        onRowClick={(r) => setExpanded((p) => (p === r.line ? null : r.line))}
        renderExpanded={(r) => {
          const sub = models.filter((m) => m.line === r.line);
          return sub.length === 0
            ? <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>No model-level sales data for {r.line}.</div>
            : (
              <div>
                <div style={{ color: C.gold, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }} className="mb-1">
                  {r.line} — by model
                </div>
                <ItemTable rows={sub} cols={subCols} />
              </div>
            );
        }}
      />
    );
  }
  return <ItemTable rows={rows} cols={cols} />;
}

/* Model | Product Line toggle */
function GranularityToggle({ value, onChange }) {
  return (
    <div className="flex gap-1 mb-3">
      {["model","line"].map((g) => (
        <button key={g} onClick={() => onChange(g)}
          style={{
            fontFamily: SANS, borderRadius: 8, fontSize: 12,
            border: `1px solid ${value === g ? C.gold : C.line}`,
            background: value === g ? C.gold : "transparent",
            color: value === g ? C.bg : C.dim,
          }} className="px-3 py-1">
          {g === "model" ? "By model" : "By product line"}
        </button>
      ))}
    </div>
  );
}

/* Buy Signals hierarchy: Brand → Product line → Model (with projected cost) */
function BuyModelTable({ models }) {
  return <ItemTable rows={models} cols={[
    ["model","Model"],["units","Units"],["profit","Profit $",fmtMoney],
    ["avgProfit","Avg profit",fmtMoney],["avgCost","Proj. cost",fmtMoney],
    ["medianDays","Median days"],["buyScore","Score",(v) => v?.toFixed(3)],
    ["health","Stock health",(v,r) => <HealthBadge health={v} weeksOfStock={r?.weeksOfStock} />],
  ]} />;
}
function BuyLineTable({ lines }) {
  const [open, setOpen] = useState(null);
  return <ItemTable rows={lines}
    getRowKey={(r) => r.line} expandedKey={open}
    onRowClick={(r) => setOpen((p) => p === r.line ? null : r.line)}
    renderExpanded={(r) => (
      <div>
        <div style={{ color: C.gold, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }} className="mb-1">{r.line} — models</div>
        <BuyModelTable models={r.models} />
      </div>
    )}
    cols={[["line","Product line"],["units","Units"],["profit","Profit $",fmtMoney],["projCost","Proj. cost",fmtMoney],["score","Score",(v) => v?.toFixed(3)]]}
  />;
}
function BuyHierarchy({ brands }) {
  const [open, setOpen] = useState(null);
  if (!brands.length) return <Locked msg="No ranked data available for this filter." />;
  return <ItemTable rows={brands}
    getRowKey={(r) => r.brand} expandedKey={open}
    onRowClick={(r) => setOpen((p) => p === r.brand ? null : r.brand)}
    renderExpanded={(r) => (
      <div>
        <div style={{ color: C.gold, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }} className="mb-1">{r.brand} — product lines</div>
        <BuyLineTable lines={r.lines} />
      </div>
    )}
    cols={[["brand","Brand"],["units","Units"],["profit","Profit $",fmtMoney],["projCost","Proj. cost",fmtMoney],["score","Buy score",(v) => v?.toFixed(3)]]}
  />;
}

/* Funding Scenarios: given a budget, spread it across the top buy-signal models
   (one unit each, down the ranking) using projected cost = avg historical total cost */
function FundingScenarios({ ranking }) {
  const [budget, setBudget] = useState(50000);
  const priced = ranking.filter((m) => m.avgCost && m.avgCost > 0);
  let remaining = budget;
  const picks = [];
  for (const m of priced) {
    if (m.avgCost <= remaining) { picks.push(m); remaining -= m.avgCost; }
  }
  const spent = budget - remaining;
  const presets = [20000, 50000, 100000];
  return (
    <div className="flex flex-col gap-3">
      <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12 }}>
        Pick a budget and we'll spread it across the highest buy-score watches — one unit each down the ranking — using each
        model's projected cost (its average historical total cost). This avoids dumping the whole budget into one reference.
      </div>
      <div className="flex flex-wrap items-center gap-2">
        {presets.map((p) => (
          <button key={p} onClick={() => setBudget(p)}
            style={{ fontFamily: SANS, fontSize: 13, borderRadius: 999,
              border: `1px solid ${budget === p ? C.gold : C.line}`,
              background: budget === p ? C.gold : "transparent",
              color: budget === p ? C.bg : C.dim }} className="px-4 py-1.5">{fmtMoney(p)}</button>
        ))}
        <span className="flex items-center gap-1" style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>
          custom $
          <input type="number" min={0} step={1000} value={budget}
            onChange={(e) => { const n = parseInt(e.target.value, 10); setBudget(isNaN(n) ? 0 : n); }}
            style={{ width: 110, fontFamily: SANS, fontSize: 13, color: C.text, background: C.bg, border: `1px solid ${C.line}`, borderRadius: 8, padding: "4px 8px" }} />
        </span>
      </div>
      <div className="flex flex-wrap gap-3">
        <Stat label="Watches to buy" value={picks.length} />
        <Stat label="Projected spend" value={fmtMoney(spent)} sub={`of ${fmtMoney(budget)} budget`} />
        <Stat label="Left over" value={fmtMoney(remaining)} sub="too small for the next pick" />
        <Stat label="Projected profit" value={fmtMoney(picks.reduce((s, m) => s + (m.avgProfit || 0), 0))} sub="if each sells at avg profit" />
      </div>
      {picks.length === 0
        ? <Locked msg={priced.length === 0 ? "No models have a cost basis yet (need historical sales cost)." : "Budget too small for even the cheapest ranked watch."} />
        : <ItemTable rows={picks} cols={[
            ["brand","Brand"],["model","Model"],["avgCost","Proj. cost",fmtMoney],
            ["avgProfit","Avg profit",fmtMoney],["medianDays","Median days"],["buyScore","Score",(v) => v?.toFixed(3)],
          ]} />}
    </div>
  );
}

function BuyTab({ M, filters, includePresold }) {
  const [g1, setG1] = useState("model"); // fastest
  const [g2, setG2] = useState("model"); // best profit
  const [g3, setG3] = useState("model"); // best score
  const [g4, setG4] = useState("model"); // health by velocity
  const [g5, setG5] = useState("model"); // health by score

  if (!M.salesHasBrand)
    return (
      <div className="flex flex-col gap-4">
        <Locked msg="Everything below needs brand + model/reference on your sales export. Add those columns and this tab fills in completely." />
        <Panel title="What this tab answers">
          <ul style={{ color: C.dim, fontFamily: SANS }} className="text-sm flex flex-col gap-2 mt-1">
            <li><span style={{ color: C.gold }}>1.</span> What is our highest velocity, highest profit product — and are we out?</li>
            <li><span style={{ color: C.gold }}>2.</span> What products should I buy?</li>
            <li><span style={{ color: C.gold }}>3.</span> Rank products by velocity and profit combined.</li>
            <li><span style={{ color: C.gold }}>4.</span> Most frequently sold items — by model and by product line.</li>
          </ul>
        </Panel>
      </div>
    );

  const ranking = applyFilters(M.ranking, filters);
  // buy-signals hierarchy, filtered by the active brand/line chips
  const buyBrands = (M.buyHierarchy || [])
    .filter((b) => !filters.brands || filters.brands.has(b.brand))
    .map((b) => ({ ...b, lines: b.lines.filter((l) => !filters.lines || filters.lines.has(l.line)) }))
    .filter((b) => b.lines.length > 0);
  const allModels = applyFilters(M.salesByModel, filters);
  const velProf = applyFilters(M.velProfRanking, filters);

  const [excludedKeys, setExcludedKeys] = useState(() => new Set());
  const keyOf = (x) => (x.brand || "") + "|" + x.model;
  // clear exclusions whenever the underlying ranked list changes (e.g. brand filter applied)
  const velProfKey = velProf.map(keyOf).join(",");
  useEffect(() => { setExcludedKeys(new Set()); }, [velProfKey]);
  const availableVelProf = velProf.filter((x) => !excludedKeys.has(keyOf(x)));
  const top10VelProf = availableVelProf.slice(0, 10);
  const excludedVelProf = velProf.filter((x) => excludedKeys.has(keyOf(x)));
  function excludeVelProf(x) {
    setExcludedKeys((prev) => new Set(prev).add(keyOf(x)));
  }
  function restoreVelProf(x) {
    setExcludedKeys((prev) => { const next = new Set(prev); next.delete(keyOf(x)); return next; });
  }

  // fastest-10 with ignore/refill (uses the full ranked list so a removed watch is replaced)
  const fastestAll = { model: applyFilters(M.fastestModelsAll, filters), line: applyFilters(M.fastestLinesAll, filters) };
  const fastKeyOf = (r) => (r.brand || "") + "|" + (r.model ?? r.line);
  const [excludedFast, setExcludedFast] = useState(() => new Set());
  const fastListKey = fastestAll[g1].map(fastKeyOf).join(",");
  useEffect(() => { setExcludedFast(new Set()); }, [fastListKey]);
  const fastVisible = fastestAll[g1].filter((r) => !excludedFast.has(fastKeyOf(r))).slice(0, 10);
  const fastExcludedItems = fastestAll[g1].filter((r) => excludedFast.has(fastKeyOf(r)));

  const bestProfit = { model: applyFilters(M.bestProfitModels, filters), line: applyFilters(M.bestProfitLines, filters) };
  const bestScore = { model: applyFilters(M.bestScoreModels, filters), line: applyFilters(M.bestScoreLines, filters) };
  const healthVel = { model: applyFilters(M.healthByVelocityModels, filters), line: applyFilters(M.healthByVelocityLines, filters) };
  const healthScore = { model: applyFilters(M.healthByScoreModels, filters), line: applyFilters(M.healthByScoreLines, filters) };

  return (
    <div className="flex flex-col gap-5">
      {M.windowMode === "window" && (
        <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>
          Based on the last {M.salesWindowDays} days of sales (from today).
          Models sold only 1–2 times in that window are excluded from rankings below — not enough history to trust.
        </div>
      )}

      {/* Q1 */}
      <QuestionCard num="1" question="What is our highest velocity, highest profit product — and are we out?">
        {top10VelProf.length ? (
          <div className="flex flex-col gap-3">
            <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12 }}>
              Top 10 ranked by velocity + profit combined. Click "Exclude" to drop a watch from the list — the next-ranked watch will take its place.
            </div>
            <div className="flex flex-col gap-2">
              {top10VelProf.map((x, i) => (
                <div key={keyOf(x)} style={{ background: C.panel2, borderRadius: 12 }} className="p-4">
                  <div className="flex items-center justify-between flex-wrap gap-3">
                    <div className="flex items-center gap-3">
                      <div style={{ color: C.gold, fontFamily: SERIF, fontSize: 18, minWidth: 24 }}>#{i + 1}</div>
                      <div>
                        <div style={{ fontFamily: SERIF, color: C.gold }} className="text-lg">{x.brand} — {x.model}</div>
                        <div style={{ color: C.dim, fontFamily: SANS }} className="text-sm mt-1">
                          {x.units} sold · {fmtMoney(x.avgProfit)} avg profit · {x.medianDays != null ? x.medianDays + " days to sell" : "—"} · {fmtPct(x.profitPct)} margin
                        </div>
                      </div>
                    </div>
                    <div className="flex items-center gap-3">
                      <HealthBadge health={x.health || (x.stock === 0 || x.stock == null ? "red" : "green")} weeksOfStock={x.weeksOfStock} />
                      <button onClick={() => excludeVelProf(x)}
                        style={{
                          fontFamily: SANS, fontSize: 12, borderRadius: 8,
                          border: `1px solid ${C.line}`, background: "transparent", color: C.dim,
                        }} className="px-3 py-1">
                        Exclude
                      </button>
                    </div>
                  </div>
                </div>
              ))}
            </div>
            {excludedVelProf.length > 0 && (
              <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }} className="mt-1">
                <div className="mb-1">Excluded:</div>
                <div className="flex flex-wrap gap-2">
                  {excludedVelProf.map((x) => (
                    <button key={keyOf(x)} onClick={() => restoreVelProf(x)}
                      style={{
                        fontFamily: SANS, fontSize: 12, borderRadius: 999,
                        border: `1px solid ${C.line}`, background: "transparent", color: C.dim,
                      }} className="px-3 py-1">
                      {x.brand} — {x.model} ✕
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>
        ) : <Locked msg="No ranked data available." />}
      </QuestionCard>

      {/* Q2 */}
      <QuestionCard num="2" question="What products should I buy?">
        <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
          Ranked by <b>buy score</b> = velocity (35%) + avg profit (35%) + sales volume (30%), boosted when out of stock.
          Brands are ordered by their models' combined buy score. Click a <b>brand</b> to see its product lines, then a line to see models.
          Each model shows its <b>projected cost</b> (average historical total cost).
          {" "}{includePresold ? "Presold sales (0–2 days) are included." : "Presold sales (0–2 days) are excluded."}
        </div>
        <BuyHierarchy brands={buyBrands} />
      </QuestionCard>

      {/* Funding scenarios */}
      <QuestionCard num="$" question="Funding scenarios — what to buy with your budget">
        <FundingScenarios ranking={ranking} />
      </QuestionCard>

      {/* Fastest 10 */}
      <QuestionCard num="3" question="Fastest 10 watches to sell">
        <GranularityToggle value={g1} onChange={setG1} />
        <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
          Fastest sellers by median days to sell. Click "Ignore" to drop one — the next-fastest takes its place.
        </div>
        {fastVisible.length === 0
          ? <Locked msg="No median-days data available for this view." />
          : <RankTable rows={fastVisible} nameKey={g1} models={allModels}
              onIgnore={(r) => setExcludedFast((prev) => new Set(prev).add(fastKeyOf(r)))} />}
        {fastExcludedItems.length > 0 && (
          <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }} className="mt-3">
            <div className="mb-1">Ignored:</div>
            <div className="flex flex-wrap gap-2">
              {fastExcludedItems.map((r) => (
                <button key={fastKeyOf(r)}
                  onClick={() => setExcludedFast((prev) => { const n = new Set(prev); n.delete(fastKeyOf(r)); return n; })}
                  style={{ fontFamily: SANS, fontSize: 12, borderRadius: 999, border: `1px solid ${C.line}`, background: "transparent", color: C.dim }}
                  className="px-3 py-1">{r.brand} — {r.model ?? r.line} ✕</button>
              ))}
            </div>
          </div>
        )}
      </QuestionCard>

      {/* Best 10 by profit */}
      <QuestionCard num="4" question="Best 10 watches by profit">
        <GranularityToggle value={g2} onChange={setG2} />
        {bestProfit[g2].length === 0
          ? <Locked msg="No profit data available for this view." />
          : <RankTable rows={bestProfit[g2]} nameKey={g2} models={allModels} />}
      </QuestionCard>

      {/* Best 10 by score */}
      <QuestionCard num="5" question="Best 10 watches by score">
        <GranularityToggle value={g3} onChange={setG3} />
        <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
          Score blends velocity, profit, and volume — boosted when out of stock.
        </div>
        {bestScore[g3].length === 0
          ? <Locked msg="No scored data available for this view." />
          : <RankTable rows={bestScore[g3]} nameKey={g3} showScore models={allModels} />}
      </QuestionCard>

      {/* Stock health by velocity */}
      <QuestionCard num="6" question="Stock health by velocity — what do I need to buy soon?">
        <GranularityToggle value={g4} onChange={setG4} />
        <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
          Sorted by urgency: weeks of stock left at the current sell rate.
          <span style={{ color: C.red }}> Red</span> = under 1 week (buy now),
          <span style={{ color: "#c8863a" }}> yellow</span> = under 2 weeks (buy soon),
          <span style={{ color: C.green }}> green</span> = healthy.
        </div>
        {healthVel[g4].length === 0
          ? <Locked msg="Not enough sales velocity data to assess stock health for this view." />
          : <RankTable rows={healthVel[g4].slice(0, 30)} nameKey={g4} models={allModels} />}
      </QuestionCard>

      {/* Stock health by score */}
      <QuestionCard num="7" question="Stock health by score">
        <GranularityToggle value={g5} onChange={setG5} />
        <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
          Same urgency colors, but ordered by buy score within each health tier — your highest-priority restocks float to the top.
        </div>
        {healthScore[g5].length === 0
          ? <Locked msg="Not enough data to assess stock health for this view." />
          : <RankTable rows={healthScore[g5].slice(0, 30)} nameKey={g5} showScore models={allModels} />}
      </QuestionCard>

      {/* Q8: most frequently sold */}
      <QuestionCard num="8" question="Most frequently sold items — by model and by product line">
        <div className="grid gap-5" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))" }}>
          <div>
            <SectionLabel>By model (top 20)</SectionLabel>
            <ItemTable rows={applyFilters(M.salesByModel, filters).slice(0, 20)} cols={[
              ["model","Model"],["brand","Brand"],["units","Units"],
              ["profit","Profit $",fmtMoney],["profitPct","Margin %",fmtPct],
            ]} />
          </div>
          <div>
            <SectionLabel>By product line</SectionLabel>
            {applyFilters(M.salesByLine, filters).length === 0
              ? <Locked msg="No product-line matches in your sales data." />
              : <ItemTable rows={applyFilters(M.salesByLine, filters)} cols={[
                ["brand","Brand"],["line","Line"],["units","Units"],
                ["profit","Profit $",fmtMoney],["medianDays","Median days"],
              ]} />
            }
          </div>
        </div>
      </QuestionCard>

    </div>
  );
}

function ItemTable({ rows, cols, onRowClick, getRowKey, expandedKey, renderExpanded }) {
  return (
    <div className="overflow-auto">
      <table className="w-full" style={{ fontFamily: SANS, fontSize: 13 }}>
        <thead>
          <tr style={{ color: C.faint }}>
            {onRowClick && <th style={{ width: 22 }}></th>}
            {cols.map(([k, label]) => <th key={k} className="text-left py-2 pr-4 font-normal text-xs uppercase tracking-wide">{label}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => {
            const key = getRowKey ? getRowKey(r) : i;
            const isExpanded = renderExpanded && expandedKey != null && key === expandedKey;
            return (
              <React.Fragment key={key}>
                <tr onClick={onRowClick ? () => onRowClick(r) : undefined}
                  style={{
                    borderTop: `1px solid ${C.line}`, color: C.text,
                    cursor: onRowClick ? "pointer" : "default",
                  }}
                  className={onRowClick ? "hover:opacity-70" : undefined}>
                  {onRowClick && (
                    <td className="py-2 text-center" style={{ color: C.faint, fontSize: 11 }}>
                      {renderExpanded ? (isExpanded ? "▾" : "▸") : ""}
                    </td>
                  )}
                  {cols.map(([k, , fmt]) => (
                    <td key={k} className="py-2 pr-4">{fmt ? fmt(r[k], r) : (r[k] ?? "--")}</td>
                  ))}
                </tr>
                {isExpanded && (
                  <tr style={{ borderTop: `1px solid ${C.line}` }}>
                    <td colSpan={cols.length + 1} className="pt-2 pb-3 pl-6">{renderExpanded(r)}</td>
                  </tr>
                )}
              </React.Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
