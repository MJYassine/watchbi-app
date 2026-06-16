import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis,
  CartesianGrid, Tooltip, ResponsiveContainer, Cell, Legend,
} from "recharts";
import {
  Upload, FileSpreadsheet, Boxes, TrendingUp, Sparkles, Send,
  Check, X, AlertTriangle, Lock, Watch, Trash2, ArrowRight,
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
    ["status", "Status"], ["serial", "Serial #"],
  ],
  sales: [
    ["saleDate", "Sale / invoice date"], ["purchaseDate", "Purchase date"],
    ["cost", "Cost (COGS)"], ["salePrice", "Sale / invoice price"],
    ["profit", "Profit $ (optional)"], ["brand", "Brand (optional)"],
    ["modelName", "Model name (optional)"], ["modelNumber", "Model / reference # (optional)"],
    ["inventoryType", "Inventory type (optional)"], ["condition", "Condition — New/Used (optional)"],
    ["serial", "Serial # (optional)"],
  ],
};

/* ---- header auto-detect ---- */
function guessField(role, header) {
  const h = String(header).trim().toLowerCase();
  const has = (s) => h.includes(s);
  if (has("brand")) return "brand";
  if (has("model name") || has("title item")) return "modelName";
  if (has("model number") || has("reference") || h === "ref") return "modelNumber";
  if (role === "sales") {
    if (has("invoice date") || has("sale date") || has("sold date")) return "saleDate";
    if (has("invoice price") || has("sale price") || has("sold price")) return "salePrice";
    if (h === "profit") return "profit";
    if (has("condition")) return "condition";
    if (has("inventory type") || h === "type") return "inventoryType";
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

/* =========================================================================
   METRICS ENGINE
   ========================================================================= */
function computeMetrics(datasets, dateRange, includePresold, includeOlderSales, invStatusFilter) {
  const inv = datasets.find((d) => d.role === "inventory");
  const sal = datasets.find((d) => d.role === "sales");
  const today = new Date();
  const out = { hasInv: !!inv, hasSales: !!sal };

  /* ----- inventory ----- */
  if (inv) {
    const m = inv.mapping;
    const items = inv.rows.map((r) => {
      const cost = toNum(r[m.cost]);
      const tw = toNum(r[m.targetWholesale]);
      const tag = toNum(r[m.tagPrice]);
      const brandNorm = normalizeBrand(r[m.brand]);
      const pd = toDate(r[m.purchaseDate]);
      const age = pd ? daysBetween(pd, today) : null;
      return {
        brand: r[m.brand] || "Unknown",
        brandNorm,
        line: findLine(brandNorm, r[m.modelName]),
        modelName: r[m.modelName],
        modelNumber: r[m.modelNumber] || null,
        cost: cost || 0,
        targetWholesale: tw,
        tagPrice: tag,
        age,
        grade: ageGrade(age),
        status: r[m.status],
      };
    });
    // collect all statuses before filtering so the filter UI can show all options
    out.invStatuses = [...new Set(items.map((x) => x.status).filter(Boolean))].sort();
    const filteredItems = (invStatusFilter && invStatusFilter.size)
      ? items.filter((x) => invStatusFilter.has(x.status))
      : items;

    out.invCount = filteredItems.length;
    out.invCost = filteredItems.reduce((s, x) => s + x.cost, 0);

    const byBrand = {};
    filteredItems.forEach((x) => {
      const k = x.brand;
      byBrand[k] = byBrand[k] || { brand: k, count: 0, cost: 0 };
      byBrand[k].count++; byBrand[k].cost += x.cost;
    });
    out.invByBrand = Object.values(byBrand).sort((a, b) => b.cost - a.cost);
    out.invBrandCount = out.invByBrand.length;

    const byLine = {};
    filteredItems.filter((x) => x.line).forEach((x) => {
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
    ].map((b) => ({ ...b, count: 0, cost: 0 }));
    filteredItems.forEach((x) => {
      if (x.age == null) return;
      const b = buckets.find((q) => x.age >= q.lo && x.age <= q.hi);
      if (b) { b.count++; b.cost += x.cost; }
    });
    out.aging = buckets;
    out.agedValue = buckets.filter((b) => b.lo >= 91).reduce((s, b) => s + b.cost, 0);

    // ── condition grading (A = fresh, F = aged) ──
    const byGrade = {};
    filteredItems.forEach((x) => {
      if (!x.grade) return;
      byGrade[x.grade] = byGrade[x.grade] || { grade: x.grade, count: 0, cost: 0 };
      byGrade[x.grade].count++; byGrade[x.grade].cost += x.cost;
    });
    out.invByGrade = GRADE_BUCKETS.map((b) => byGrade[b.grade] || { grade: b.grade, count: 0, cost: 0 });

    // ── top 10 watches to sell: worst grade first, then most cash tied up ──
    out.needToSell = [...filteredItems]
      .filter((x) => x.grade)
      .sort((a, b) => (GRADE_RANK[b.grade] - GRADE_RANK[a.grade]) || (b.cost - a.cost))
      .slice(0, 10);

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
      const presold = days != null && days >= 0 && days <= 5;
      const brandNorm = normalizeBrand(r[m.brand]);
      return {
        saleDate: sdt, days, presold, cost: cost || 0, price: price || 0,
        profit: profit || 0,
        marginPct: cost ? (profit / cost) * 100 : null,
        brand: r[m.brand] || null, brandNorm,
        line: findLine(brandNorm, r[m.modelName]),
        modelName: r[m.modelName] || null,
        type: r[m.inventoryType] || "Unspecified",
        condition: normalizeCondition(r[m.condition]),
        priceTier: priceTierLabel(price),
      };
    });
    // overall date span of the sales data (used for the date-range picker bounds)
    const allSaleDates = rows.map((x) => x.saleDate).filter(Boolean);
    out.salesDateMin = allSaleDates.length ? new Date(Math.min(...allSaleDates.map((d) => d.getTime()))) : null;
    out.salesDateMax = allSaleDates.length ? new Date(Math.max(...allSaleDates.map((d) => d.getTime()))) : null;

    // apply optional date-range filter (rows with no sale date are kept either way);
    // if the user hasn't picked a custom range, default to the trailing 45-day
    // sales history — this also drives the Buy Signals tab.
    out.salesWindowDays = 45;
    if (dateRange && (dateRange.start || dateRange.end)) {
      const startT = dateRange.start ? new Date(dateRange.start + "T00:00:00").getTime() : -Infinity;
      const endT = dateRange.end ? new Date(dateRange.end + "T23:59:59").getTime() : Infinity;
      rows = rows.filter((x) => !x.saleDate || (x.saleDate.getTime() >= startT && x.saleDate.getTime() <= endT));
      out.usingDefaultWindow = false;
    } else if (out.salesDateMax && !includeOlderSales) {
      const windowStartT = out.salesDateMax.getTime() - out.salesWindowDays * 86400000;
      out.windowStart = new Date(windowStartT);
      rows = rows.filter((x) => x.saleDate && x.saleDate.getTime() >= windowStartT);
      out.usingDefaultWindow = true;
    } else {
      out.usingDefaultWindow = false;
    }

    // "presold" = sold within 0-5 days of purchase — likely flipped before it ever
    // hit the floor, so it's excluded from sell-through analysis by default.
    out.presoldCount = rows.filter((x) => x.presold).length;
    if (!includePresold) {
      rows = rows.filter((x) => !x.presold);
    }

    out.salesUnits = rows.length;
    out.salesProfit = rows.reduce((s, x) => s + x.profit, 0);
    out.salesRevenue = rows.reduce((s, x) => s + x.price, 0);
    out.medianMargin = median(rows.map((x) => x.marginPct));
    out.medianDays = median(rows.map((x) => x.days));
    out.meanDays = (() => {
      const a = rows.map((x) => x.days).filter((x) => x != null);
      return a.length ? a.reduce((s, x) => s + x, 0) / a.length : null;
    })();

    // span of the sales data, in weeks (min 1) — used for sell-through velocity
    out.salesWeeks = (() => {
      const ds = rows.map((x) => x.saleDate).filter(Boolean).map((d) => d.getTime());
      if (ds.length < 2) return 1;
      const span = (Math.max(...ds) - Math.min(...ds)) / (7 * 86400000);
      return Math.max(span, 1);
    })();

    // monthly
    const mo = {};
    rows.forEach((x) => {
      if (!x.saleDate) return;
      const k = x.saleDate.getFullYear() + "-" + String(x.saleDate.getMonth() + 1).padStart(2, "0");
      mo[k] = mo[k] || { month: k, units: 0, profit: 0 };
      mo[k].units++; mo[k].profit += x.profit;
    });
    out.monthly = Object.values(mo).sort((a, b) => a.month.localeCompare(b.month));

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

    // brand-level (only if brand present on sales)
    out.salesHasBrand = rows.some((x) => x.brand);
    if (out.salesHasBrand) {
      const bb = {};
      rows.forEach((x) => {
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
      rows.forEach((x) => {
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
      })).sort((a, b) => (a.brand === b.brand ? a.condition.localeCompare(b.condition) : b.profit - a.profit));
      out.hasCondition = rows.some((x) => x.condition && x.condition !== "Unspecified");

      // by product line
      const bl = {};
      rows.forEach((x) => {
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
      rows.forEach((x) => {
        const name = x.modelName || x.brand;
        const k = (x.brand || "") + "|" + name;
        bm[k] = bm[k] || { key: k, brand: x.brand, line: lineLabel(x.brand, x.line), model: name, brandNorm: x.brandNorm, units: 0, profit: 0, revenue: 0, cost: 0, _days: [], _m: [] };
        bm[k].units++; bm[k].profit += x.profit; bm[k].revenue += x.price; bm[k].cost += x.cost;
        if (x.days != null) bm[k]._days.push(x.days);
        if (x.marginPct != null) bm[k]._m.push(x.marginPct);
      });
      let models = Object.values(bm).map((b) => ({
        brand: b.brand, line: b.line, model: b.model, units: b.units, profit: b.profit, revenue: b.revenue,
        avgProfit: b.profit / b.units,
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
  return out;
}

/* Compact summary handed to the chatbot — computed aggregates only. */
function metricsForChat(M) {
  const o = {
    inventory_loaded: M.hasInv, sales_loaded: M.hasSales,
    sales_has_brand_or_model: !!M.salesHasBrand,
  };
  if (M.hasInv) {
    o.inventory = {
      items: M.invCount, total_cost: Math.round(M.invCost),
      brands: M.invBrandCount,
      by_brand: M.invByBrand.slice(0, 20).map((b) => ({ brand: b.brand, count: b.count, cost: Math.round(b.cost) })),
      by_line: M.invByLine.map((b) => ({ brand: b.brand, line: b.line, count: b.count })),
      aging: M.aging.map((b) => ({ bucket: b.label, items: b.count, cost: Math.round(b.cost) })),
      aged_capital_91plus: Math.round(M.agedValue),
      projected_profit_items: M.projItems,
      projected_profit_total: Math.round(M.projProfit),
      projected_profit_note: "only computed on items that have a target wholesale price",
      by_grade: M.invByGrade.map((b) => ({ grade: b.grade, items: b.count, cost: Math.round(b.cost) })),
      grade_note: "A = 0-30 days in stock, B = 31-60, C = 61-90, D = 91-180, F = 180+",
      top_to_sell: (M.needToSell || []).map((b) => ({
        brand: b.brand, model: b.modelName, ref: b.modelNumber,
        grade: b.grade, days_in_stock: b.age, cost: Math.round(b.cost),
      })),
    };
  }
  if (M.hasSales) {
    o.sales = {
      units_sold: M.salesUnits, total_profit: Math.round(M.salesProfit),
      revenue: Math.round(M.salesRevenue),
      median_margin_pct: M.medianMargin && +M.medianMargin.toFixed(1),
      median_days_to_sell: M.medianDays, mean_days_to_sell: M.meanDays && Math.round(M.meanDays),
      window_note: M.usingDefaultWindow
        ? `figures reflect the trailing ${M.salesWindowDays} days of sales history (through ${M.salesDateMax ? M.salesDateMax.toISOString().slice(0, 10) : "--"})`
        : "figures reflect the user-selected date range",
      by_inventory_type: M.salesByType,
      by_price_tier: M.salesByPriceTier.map((b) => ({
        tier: b.tier, units: b.units, profit: Math.round(b.profit), avg_profit: b.avgProfit != null ? Math.round(b.avgProfit) : null,
      })),
      by_brand_condition: (M.salesByBrandCondition || []).map((b) => ({
        brand: b.brand, condition: b.condition, units: b.units, profit: Math.round(b.profit), margin_pct: b.profitPct != null ? +b.profitPct.toFixed(1) : null,
      })),
      monthly: M.monthly,
    };
    if (M.salesHasBrand) {
      o.sales.by_brand = M.salesByBrand.map((b) => ({
        brand: b.brand, units: b.units, profit: Math.round(b.profit),
        median_days: b.medianDays, median_margin_pct: b.medianMargin && +b.medianMargin.toFixed(1),
      }));
      o.sales.top_models_by_frequency = M.salesByModel.slice(0, 15).map((b) => ({
        brand: b.brand, model: b.model, units: b.units,
        avg_profit: Math.round(b.avgProfit), median_days: b.medianDays, current_stock: b.stock,
      }));
      o.sales.ranked_buy_list = M.ranking.slice(0, 15).map((b) => ({
        brand: b.brand, model: b.model, buy_score: b.buyScore,
        units_sold: b.units, avg_profit: Math.round(b.avgProfit),
        median_days: b.medianDays, current_stock: b.stock,
        weekly_velocity: b.weeklyVelocity, weeks_of_stock: b.weeksOfStock, stock_health: b.health,
      }));
      o.sales.fastest_models = M.fastestModels.map((b) => ({ brand: b.brand, model: b.model, median_days: b.medianDays }));
      o.sales.fastest_lines = M.fastestLines.map((b) => ({ brand: b.brand, line: b.line, median_days: b.medianDays }));
      o.sales.best_profit_models = M.bestProfitModels.map((b) => ({ brand: b.brand, model: b.model, profit: Math.round(b.profit) }));
      o.sales.best_profit_lines = M.bestProfitLines.map((b) => ({ brand: b.brand, line: b.line, profit: Math.round(b.profit) }));
      o.sales.stock_health_urgent = M.healthByVelocityModels.filter((b) => b.health !== "green").slice(0, 15).map((b) => ({
        brand: b.brand, model: b.model, health: b.health, weeks_of_stock: b.weeksOfStock,
        current_stock: b.stock, weekly_velocity: b.weeklyVelocity,
      }));
    } else {
      o.sales.brand_model_breakdowns = "UNAVAILABLE — the sales data has no brand or model/reference column, so velocity/profit/frequency by brand, line, or model cannot be computed yet. Tell the user this plainly when relevant.";
    }
  }
  return o;
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

/* presold (0-5 day) sales toggle */
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
        Sales completed within 0–5 days of purchase ({count ?? 0}) are considered "presold" and {value ? "are included in" : "are not part of"} this analysis.
      </span>
    </div>
  );
}

/* 45-day sales window toggle */
function SalesWindowFilter({ value, onChange, windowDays, salesDateMax }) {
  const opts = [[false, `Last ${windowDays ?? 45} days (default)`], [true, "Include 45+ day sales"]];
  return (
    <div className="flex flex-wrap items-center gap-2 mb-1">
      <span style={{ color: C.faint, fontFamily: SANS, fontSize: 11, textTransform: "uppercase", letterSpacing: ".08em", minWidth: 50 }}>
        History
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
        By default, Sales and Buy Signals use the trailing {windowDays ?? 45} days{salesDateMax ? ` (through ${salesDateMax.toISOString().slice(0, 10)})` : ""}.
        {value ? " Now showing the full sales history instead." : " Toggle to include older sales."}
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
  const [dateRange, setDateRange] = useState({ start: "", end: "" });
  const [includePresold, setIncludePresold] = useState(false);
  const [includeOlderSales, setIncludeOlderSales] = useState(false);
  // invStatusFilter is derived from selectedInvStatuses after metrics are known;
  // we bootstrap with null (all) on first render, then the Dashboard passes back
  // the active set via setInvStatusFilter once the user picks statuses.
  const [invStatusFilter, setInvStatusFilter] = useState(null);
  const metrics = useMemo(() => (stage === "dash" ? computeMetrics(active, dateRange, includePresold, includeOlderSales, invStatusFilter) : null), [stage, datasets, dateRange, includePresold, includeOlderSales, invStatusFilter]);

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

      {stage === "upload" && <UploadView fileRef={fileRef} onFiles={handleFiles} err={err} />}
      {stage === "map" && (
        <MapView datasets={datasets} setRole={setRole} setMap={setMap} removeDs={removeDs}
          onBuild={() => setStage("dash")} onAdd={() => fileRef.current?.click()} fileRef={fileRef} onFiles={handleFiles} />
      )}
      {stage === "dash" && metrics && <Dashboard M={metrics} dateRange={dateRange} setDateRange={setDateRange} includePresold={includePresold} setIncludePresold={setIncludePresold} includeOlderSales={includeOlderSales} setIncludeOlderSales={setIncludeOlderSales} invStatusFilter={invStatusFilter} setInvStatusFilter={setInvStatusFilter} />}
    </div>
  );
}

/* ---------- upload ---------- */
function UploadView({ fileRef, onFiles, err }) {
  const [drag, setDrag] = useState(false);
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
                {["inventory", "sales", "ignore"].map((r) => (
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

function Dashboard({ M, dateRange, setDateRange, includePresold, setIncludePresold, includeOlderSales, setIncludeOlderSales, invStatusFilter, setInvStatusFilter }) {
  const [tab, setTab] = useState(M.hasInv ? "inventory" : "sales");
  const tabs = [
    M.hasInv && ["inventory", "Inventory", Boxes],
    M.hasSales && ["sales", "Sales", TrendingUp],
    M.hasSales && ["buy", "Buy Signals", Sparkles],
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
    (includeOlderSales ? 1 : 0) + (statusFilterActive ? 1 : 0);

  return (
    <div className="flex flex-col lg:flex-row" style={{ minHeight: 520 }}>
      <div className="flex-1 p-6">
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
                  <SalesWindowFilter value={includeOlderSales} onChange={setIncludeOlderSales} windowDays={M.salesWindowDays} salesDateMax={M.salesDateMax} />
                )}
                {M.hasSales && (
                  <PresoldFilter value={includePresold} onChange={setIncludePresold} count={M.presoldCount} />
                )}
                {M.hasSales && (
                  <DateRangeFilter value={dateRange} onChange={setDateRange} min={M.salesDateMin} max={M.salesDateMax} />
                )}
              </div>
            )}
          </div>
        )}
        {tab === "inventory" && <InventoryTab M={M} filters={filters} />}
        {tab === "sales" && <SalesTab M={M} filters={filters} />}
        {tab === "buy" && <BuyTab M={M} filters={filters} includePresold={includePresold} />}
      </div>
      <ChatPanel M={M} />
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

function InventoryTab({ M, filters }) {
  const fByBrand = applyFilters(M.invByBrand, filters);
  const fByLine = applyFilters(M.invByLine, filters);
  const fProjByBrand = applyFilters(M.projByBrand, filters);
  const fProjByModel = applyFilters(M.projByModel, filters);
  const fInvCount = fByBrand.reduce((s, r) => s + (r.count || 0), 0);
  const fInvCost = fByBrand.reduce((s, r) => s + (r.cost || 0), 0);
  const fProjProfit = fProjByModel.reduce((s, r) => s + (r.profit || 0), 0);
  const anyFilterActive = !!(filters.brands || filters.lines || filters.health || filters.stock !== "all");

  return (
    <div className="flex flex-col gap-5">
      {/* ── KPIs ── */}
      <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))" }}>
        <Stat label="Items in stock" value={fInvCount} />
        <Stat label="Capital tied up" value={fmtMoney(fInvCost)} />
        <Stat label="Brands" value={fByBrand.length} />
        <Stat label="Aged 91+ days" value={fmtMoney(M.agedValue)} sub={anyFilterActive ? "all brands · not filterable" : "cost sitting on the shelf"} />
        {M.projItems > 0 && <Stat label="Projected profit" value={fmtMoney(fProjProfit)} sub={`${fProjByModel.length} priced items`} />}
      </div>

      {/* ── By Brand ── */}
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

      {/* ── By Product Line ── */}
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

      {/* ── Aging ── */}
      <Panel title="Age of inventory" note="days stock held">
        <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))" }}>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={M.aging} margin={{ top: 4, right: 8, bottom: 4, left: 0 }}>
              <CartesianGrid stroke={C.line} vertical={false} />
              <XAxis dataKey="label" tick={{ fill: C.dim, fontSize: 11 }} />
              <YAxis tick={{ fill: C.faint, fontSize: 11 }} />
              <Tooltip {...chartTip} formatter={(v, n) => n === "Cost" ? fmtMoney(v) : v} />
              <Bar dataKey="count" name="Items" radius={[4, 4, 0, 0]}>
                {M.aging.map((b, i) => <Cell key={i} fill={b.lo >= 91 ? C.red : b.lo >= 61 ? "#c8863a" : C.gold} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
          <div>
            <ItemTable rows={M.aging.map(b => ({ ...b, label: b.label + " days" }))}
              cols={[["label","Age bucket"],["count","Items"],["cost","Cost tied up",fmtMoney]]} />
            {M.agedValue > 0 && (
              <div style={{ color: C.red, fontFamily: SANS, fontSize: 12, marginTop: 10 }}>
                ⚠ {fmtMoney(M.agedValue)} sitting 91+ days
              </div>
            )}
          </div>
        </div>
      </Panel>

      {/* ── Condition grading ── */}
      <Panel title="Condition grade (A → F)" note="A = just bought, F = sitting longest">
        <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
          Every item is graded by how long it's been in stock: A (0–30 days), B (31–60), C (61–90), D (91–180), F (180+).
          Grades naturally fall as a watch sits longer without selling.
        </div>
        <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))" }}>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={M.invByGrade} margin={{ top: 4, right: 8, bottom: 4, left: 0 }}>
              <CartesianGrid stroke={C.line} vertical={false} />
              <XAxis dataKey="grade" tick={{ fill: C.dim, fontSize: 11 }} />
              <YAxis tick={{ fill: C.faint, fontSize: 11 }} />
              <Tooltip {...chartTip} formatter={(v, n) => n === "Cost" ? fmtMoney(v) : v} />
              <Bar dataKey="count" name="Items" radius={[4, 4, 0, 0]}>
                {M.invByGrade.map((b, i) => <Cell key={i} fill={(GRADE_CFG[b.grade] || {}).bg || C.gold} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
          <ItemTable rows={M.invByGrade} cols={[
            ["grade","Grade",(v) => <GradeBadge grade={v} />],
            ["count","Items"],["cost","Cost tied up",fmtMoney],
          ]} />
        </div>
      </Panel>

      {/* ── Top 10 to sell ── */}
      <Panel title="Top 10 watches to sell" note="worst grade + most cash tied up first">
        {M.needToSell && M.needToSell.length > 0 ? (
          <>
            <div style={{ color: C.dim, fontFamily: SANS, fontSize: 12, marginBottom: 12 }}>
              These are holding the most cash for the longest. Moving these frees up capital to reinvest in faster, more profitable watches.
            </div>
            <ItemTable rows={M.needToSell} cols={[
              ["brand","Brand"],["modelName","Model"],["modelNumber","Ref #"],
              ["grade","Grade",(v) => <GradeBadge grade={v} />],
              ["age","Days in stock"],["cost","Cost",fmtMoney],
            ]} />
          </>
        ) : <Locked msg="No graded inventory items available." />}
      </Panel>

      {/* ── Projected Profit ── */}
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
  const lineTableRef = useRef(null);
  const handleLineBarClick = (data) => {
    const line = data?.payload?.line ?? data?.line;
    if (!line) return;
    setExpandedLine(line);
    lineTableRef.current?.scrollIntoView({ behavior: "smooth", block: "center" });
  };
  const fVelocity = applyFilters(M.byVelocity, filters);
  return (
    <div className="flex flex-col gap-5">
      {M.usingDefaultWindow && (
        <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>
          Showing the trailing {M.salesWindowDays}-day sales history (through {M.salesDateMax ? M.salesDateMax.toISOString().slice(0, 10) : "--"}) —
          this is also what Buy Signals is based on. Pick a custom date range above to override.
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

      {/* ── Over time ── */}
      <Panel title="Sales over time" note="units & profit by month">
        <ResponsiveContainer width="100%" height={270}>
          <LineChart data={M.monthly} margin={{ top: 4, right: 16, bottom: 20, left: 0 }}>
            <CartesianGrid stroke={C.line} vertical={false} />
            <XAxis dataKey="month" tick={{ fill: C.faint, fontSize: 10 }}
              angle={M.monthly.length > 12 ? -40 : 0}
              textAnchor={M.monthly.length > 12 ? "end" : "middle"}
              interval={M.monthly.length > 24 ? "preserveStartEnd" : 0}
              height={M.monthly.length > 12 ? 48 : 24} />
            <YAxis yAxisId="l" tick={{ fill: C.faint, fontSize: 11 }} />
            <YAxis yAxisId="r" orientation="right" tick={{ fill: C.faint, fontSize: 11 }} tickFormatter={fmtK} />
            <Tooltip {...chartTip} formatter={(v, n) => n === "Profit $" ? fmtMoney(v) : v} />
            <Legend wrapperStyle={{ fontFamily: SANS, fontSize: 12, paddingTop: 8 }}
              formatter={(val) => <span style={{ color: val === "Units" ? C.gold : C.green }}>{val}</span>} />
            <Line yAxisId="l" type="monotone" dataKey="units" name="Units" stroke={C.gold} strokeWidth={2} dot={M.monthly.length < 30} />
            <Line yAxisId="r" type="monotone" dataKey="profit" name="Profit $" stroke={C.green} strokeWidth={2} dot={M.monthly.length < 30} />
          </LineChart>
        </ResponsiveContainer>
      </Panel>

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
                    <Bar dataKey="profit" name="Profit" fill={C.green} radius={[0, 4, 4, 0]} />
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
              <ItemTable rows={fBrand} cols={[
                ["brand","Brand"],["units","Units"],["profit","Profit $",fmtMoney],
                ["profitPct","Margin %",fmtPct],["medianDays","Median days"],
              ]} />
            </div>
          </>)}
      </Panel>

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
                    <Bar dataKey="medianDays" name="Median days" fill={C.gold} radius={[0, 4, 4, 0]} />
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
          </>)}
      </Panel>

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
function RankTable({ rows, nameKey, showScore, models }) {
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
  const allModels = applyFilters(M.salesByModel, filters);
  const velProf = applyFilters(M.velProfRanking, filters);

  const [excludedKeys, setExcludedKeys] = useState(() => new Set());
  const keyOf = (x) => (x.brand || "") + "|" + x.model;
  const availableVelProf = velProf.filter((x) => !excludedKeys.has(keyOf(x)));
  const top10VelProf = availableVelProf.slice(0, 10);
  const excludedVelProf = velProf.filter((x) => excludedKeys.has(keyOf(x)));
  function excludeVelProf(x) {
    setExcludedKeys((prev) => new Set(prev).add(keyOf(x)));
  }
  function restoreVelProf(x) {
    setExcludedKeys((prev) => { const next = new Set(prev); next.delete(keyOf(x)); return next; });
  }

  const fastest = { model: applyFilters(M.fastestModels, filters), line: applyFilters(M.fastestLines, filters) };
  const bestProfit = { model: applyFilters(M.bestProfitModels, filters), line: applyFilters(M.bestProfitLines, filters) };
  const bestScore = { model: applyFilters(M.bestScoreModels, filters), line: applyFilters(M.bestScoreLines, filters) };
  const healthVel = { model: applyFilters(M.healthByVelocityModels, filters), line: applyFilters(M.healthByVelocityLines, filters) };
  const healthScore = { model: applyFilters(M.healthByScoreModels, filters), line: applyFilters(M.healthByScoreLines, filters) };

  return (
    <div className="flex flex-col gap-5">
      {M.usingDefaultWindow && (
        <div style={{ color: C.faint, fontFamily: SANS, fontSize: 12 }}>
          Based on the trailing {M.salesWindowDays}-day sales history (through {M.salesDateMax ? M.salesDateMax.toISOString().slice(0, 10) : "--"}).
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
          Scored by: velocity (35%) + avg profit (35%) + sales volume (30%). Out-of-stock items get a boost.
          Stock health = current stock ÷ weekly sell rate (red &lt; 1 week, yellow &lt; 2 weeks).
          {" "}{includePresold
            ? "Presold sales (0–5 days) are included in this ranking."
            : "Presold sales (0–5 days) are excluded from this ranking — toggle \"Include presold\" in the filter bar to factor them in."}
        </div>
        <RankTable rows={ranking.slice(0, 20)} nameKey="model" showScore />
      </QuestionCard>

      {/* Fastest 10 */}
      <QuestionCard num="3" question="Fastest 10 watches to sell">
        <GranularityToggle value={g1} onChange={setG1} />
        {fastest[g1].length === 0
          ? <Locked msg="No median-days data available for this view." />
          : <RankTable rows={fastest[g1]} nameKey={g1} models={allModels} />}
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

/* ---------- chatbot ---------- */
function ChatPanel({ M }) {
  const [msgs, setMsgs] = useState([
    { role: "assistant", text: "Ask me about your inventory and sales. Try: \"What should I buy?\" or \"Rank by velocity and profit.\"" },
  ]);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const scrollRef = useRef();
  useEffect(() => { scrollRef.current?.scrollTo(0, 1e9); }, [msgs, busy]);

  async function send() {
    const q = input.trim();
    if (!q || busy) return;
    const history = [...msgs, { role: "user", text: q }];
    setMsgs(history); setInput(""); setBusy(true);
    const summary = JSON.stringify(metricsForChat(M));
    const instructions =
      "You are a sharp analyst for a luxury watch dealership. Answer ONLY from the METRICS JSON below. " +
      "Use concrete numbers. Be concise and direct, no filler. If a question needs data marked UNAVAILABLE, " +
      "say so plainly and name the missing field instead of guessing. Never invent figures not in the JSON.";
    // Prior turns as plain text (skip the seeded greeting at index 0) so the
    // request is always a single, valid user message.
    const transcript = msgs
      .filter((_, i) => i !== 0)
      .map((m) => (m.role === "user" ? "Q: " : "A: ") + m.text)
      .join("\n");
    const userContent =
      instructions +
      "\n\nMETRICS JSON:\n" + summary +
      (transcript ? "\n\nEarlier in this chat:\n" + transcript : "") +
      "\n\nQuestion: " + q;
    try {
      const res = await fetch("/api/chat", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ message: userContent }),
      });
      const data = await res.json();
      const text = data.text || "";
      setMsgs((m) => [...m, { role: "assistant", text: text || "No response from the model." }]);
    } catch (e) {
      setMsgs((m) => [...m, { role: "assistant", text: "Couldn't reach the model just now. The dashboard numbers are still live." }]);
    } finally { setBusy(false); }
  }

  return (
    <div style={{
      background: C.panel, borderLeft: `1px solid ${C.line}`,
      width: 380, minWidth: 300, flexShrink: 0,
    }} className="flex flex-col">
      <div style={{ borderBottom: `1px solid ${C.line}` }} className="px-4 py-3 flex items-center gap-2">
        <Sparkles size={16} style={{ color: C.gold }} />
        <span style={{ fontFamily: SERIF }} className="text-base">Ask the data</span>
      </div>
      <div ref={scrollRef} className="flex-1 overflow-auto p-4 flex flex-col gap-3"
        style={{ minHeight: 320, maxHeight: "calc(100vh - 180px)" }}>
        {msgs.map((m, i) => (
          <div key={i} style={{
            alignSelf: m.role === "user" ? "flex-end" : "flex-start",
            background: m.role === "user" ? C.gold : C.panel2,
            color: m.role === "user" ? C.bg : C.text,
            borderRadius: 12, fontFamily: SANS, whiteSpace: "pre-wrap", maxWidth: "92%",
            lineHeight: 1.55,
          }} className="px-3 py-2 text-sm">{m.text}</div>
        ))}
        {busy && (
          <div style={{ color: C.faint, fontFamily: SANS, fontStyle: "italic" }} className="text-sm">
            thinking…
          </div>
        )}
      </div>
      <div style={{ borderTop: `1px solid ${C.line}` }} className="p-3 flex gap-2">
        <input value={input} onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && send()}
          placeholder="Ask about velocity, profit, what to buy…"
          style={{ background: C.panel2, color: C.text, border: `1px solid ${C.line}`, borderRadius: 10, fontFamily: SANS }}
          className="flex-1 text-sm px-3 py-2 outline-none" />
        <button onClick={send} disabled={busy}
          style={{
            background: busy ? C.line : C.gold,
            color: busy ? C.faint : C.bg,
            borderRadius: 10, transition: "background .2s",
          }} className="px-3 py-2">
          <Send size={16} />
        </button>
      </div>
    </div>
  );
}
