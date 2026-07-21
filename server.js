import "dotenv/config";
import express from "express";
import cors from "cors";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { JWT } from "google-auth-library";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const app = express();
app.use(cors());
app.use(express.json({ limit: "1mb" }));

/* ---------- password gate ----------
   Set APP_PASSWORD (and optionally APP_USER, default "admin") to lock the whole
   site — pages and APIs. Unset = wide open, so local dev needs no config. */
const APP_PASSWORD = process.env.APP_PASSWORD;
const APP_USER = process.env.APP_USER || "admin";
if (APP_PASSWORD) {
  app.use((req, res, next) => {
    const hdr = req.headers.authorization || "";
    if (hdr.startsWith("Basic ")) {
      const [user, ...rest] = Buffer.from(hdr.slice(6), "base64").toString().split(":");
      // constant-ish comparison; fine for a single shared password
      if (user === APP_USER && rest.join(":") === APP_PASSWORD) return next();
    }
    res.set("WWW-Authenticate", 'Basic realm="Horometrics", charset="UTF-8"');
    return res.status(401).send("Authentication required");
  });
}

/* ---------- Google Drive: pull recent WhatsApp message sheets ---------- */
const DRIVE_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;
// returns { creds } or { error } so the endpoint can report exactly what's wrong
function driveCreds() {
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    try { return { creds: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON) }; }
    catch (e) { return { error: "GOOGLE_SERVICE_ACCOUNT_JSON is set but is not valid JSON — re-paste the whole key file" }; }
  }
  const p = path.join(__dirname, "service-account.json");
  if (fs.existsSync(p)) {
    try { return { creds: JSON.parse(fs.readFileSync(p, "utf8")) }; }
    catch (e) { return { error: "service-account.json is not valid JSON" }; }
  }
  return { error: "No credentials — set the GOOGLE_SERVICE_ACCOUNT_JSON env var (paste the whole key file)" };
}
// lightweight auth: a JWT client (no giant googleapis lib). returns { client } or { error }
let _jwt = null;
function getDriveAuth() {
  if (_jwt) return { client: _jwt };
  if (!DRIVE_FOLDER_ID) return { error: "GOOGLE_DRIVE_FOLDER_ID env var is not set" };
  const c = driveCreds();
  if (c.error) return { error: c.error };
  _jwt = new JWT({ email: c.creds.client_email, key: c.creds.private_key, scopes: ["https://www.googleapis.com/auth/drive.readonly"] });
  return { client: _jwt };
}
async function driveToken(client) {
  const t = await client.getAccessToken();
  return t && t.token;
}
// minimal RFC-4180 CSV parser (handles quoted commas/newlines in messageBody)
function csvToObjects(text) {
  const rows = []; let row = [], field = "", inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false; }
      else field += c;
    } else if (c === '"') inQ = true;
    else if (c === ",") { row.push(field); field = ""; }
    else if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
    else if (c !== "\r") field += c;
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  if (!rows.length) return [];
  const headers = rows[0];
  return rows.slice(1).map((r) => { const o = {}; headers.forEach((h, i) => { o[h] = r[i] ?? null; }); return o; });
}
// map a raw message row (varied schemas) to canonical fields the dashboard reads
function normalizeMsgRow(r) {
  const g = (...keys) => { for (const k of keys) { if (r[k] != null && r[k] !== "") return r[k]; } return null; };
  return {
    intent: g("intent"),
    brand: g("brand"),
    model: g("model"),
    reference: g("fullReferenceNumber", "reference", "referenceNumber"),
    price: g("price"),
    // trim: the UI shows ~90 chars and the price check scans the first few numbers.
    // full bodies across thousands of messages are a large chunk of the heap.
    messageBody: String(g("messageBody", "body") ?? "").slice(0, 300) || null,
    timestamp: g("timestamp", "date"),
    sender: g("senderName", "sender"),
    chat: g("chatName", "chat"),
  };
}
// fetch one file's rows via the Drive REST API. Native Sheets are exported to CSV;
// .csv files are downloaded directly. (Rare .xlsx files are skipped — no xlsx lib.)
async function fetchFileRows(token, file) {
  let url;
  if (file.mimeType === "application/vnd.google-apps.spreadsheet") {
    url = `https://www.googleapis.com/drive/v3/files/${file.id}/export?mimeType=text/csv`;
  } else if (file.mimeType === "text/csv") {
    url = `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media&supportsAllDrives=true`;
  } else {
    return [];
  }
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) return [];
  return csvToObjects(await r.text());
}

// single-entry cache — one payload only. keying by `days` kept a full copy per
// distinct value, which grew the heap without bound.
let _msgCache = null; // { days, at, payload }
const MSG_TTL = 30 * 60 * 1000; // 30 min — the folder updates once a day
app.get("/api/messages", async (req, res) => {
  const g = getDriveAuth();
  if (g.error) return res.status(500).json({ error: g.error });
  const days = Math.min(30, Math.max(1, parseInt(req.query.days, 10) || 4));
  if (_msgCache && _msgCache.days === days && Date.now() - _msgCache.at < MSG_TTL && !req.query.force) return res.json(_msgCache.payload);
  try {
    const token = await driveToken(g.client);
    if (!token) return res.status(502).json({ error: "Google auth failed (check the service-account key)" });
    // list the folder via Drive REST
    let files = [], pageToken;
    do {
      const params = new URLSearchParams({
        q: `'${DRIVE_FOLDER_ID}' in parents and trashed=false`,
        fields: "nextPageToken,files(id,name,mimeType)", pageSize: "1000",
        supportsAllDrives: "true", includeItemsFromAllDrives: "true",
      });
      if (pageToken) params.set("pageToken", pageToken);
      const r = await fetch(`https://www.googleapis.com/drive/v3/files?${params}`, { headers: { Authorization: `Bearer ${token}` } });
      if (!r.ok) return res.status(502).json({ error: `Google Drive list failed (${r.status})` });
      const j = await r.json();
      files = files.concat(j.files || []); pageToken = j.nextPageToken;
    } while (pageToken);
    // date from filename; keep files within `days` of the newest date
    files.forEach((f) => { const m = String(f.name).match(/(\d{4}-\d{2}-\d{2})/); f.date = m ? m[1] : ""; });
    files = files.filter((f) => f.date).sort((a, b) => b.date.localeCompare(a.date));
    if (!files.length) return res.json({ rows: [], files: [], latest: null });
    const latest = files[0].date;
    const cutoff = new Date(new Date(latest + "T00:00:00").getTime() - days * 86400000);
    const recent = files.filter((f) => new Date(f.date + "T00:00:00") >= cutoff);
    const rows = [];
    for (const f of recent) {
      try {
        const raw = await fetchFileRows(token, f);
        for (const rr of raw) {
          const n = normalizeMsgRow(rr);
          const it = String(n.intent || "").toLowerCase();
          if (it === "buy" || it === "sell") rows.push(n);
        }
      } catch { /* skip a bad file */ }
    }
    const payload = { rows, files: recent.map((f) => f.name), latest, days };
    _msgCache = { days, at: Date.now(), payload }; // replaces any previous payload
    res.json(payload);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* ---------- WatchCharts market-value proxy (cached) ----------
   Two-step lookup per watch:
     1. /v3/search/watch?brand_name=&reference=  -> results[0].uuid
     2. /v3/watch/info?uuid=                       -> median_asking_price
   Results (including "not found" nulls) are cached to disk so we never
   spend a WatchCharts data credit on the same reference twice. */
const WC_KEY = process.env.WATCHCHARTS_API_KEY;
const WC_BASE = "https://api.watchcharts.com/v3";
const CACHE_FILE = path.join(__dirname, ".market-cache.json");

let marketCache = {};
try { marketCache = JSON.parse(fs.readFileSync(CACHE_FILE, "utf8")); } catch { /* no cache yet */ }
let cacheDirty = false;
function saveCache() {
  if (!cacheDirty) return;
  try { fs.writeFileSync(CACHE_FILE, JSON.stringify(marketCache)); cacheDirty = false; } catch { /* ignore */ }
}
setInterval(saveCache, 10000).unref?.();

const mkKey = (brand, reference) =>
  String(brand || "").toLowerCase().trim() + "|" + String(reference || "").toLowerCase().trim();

async function lookupMarket(brand, reference) {
  const key = mkKey(brand, reference);
  if (key in marketCache) return marketCache[key];
  let result = null;
  try {
    const s = await fetch(`${WC_BASE}/search/watch?brand_name=${encodeURIComponent(brand)}&reference=${encodeURIComponent(reference)}`,
      { headers: { "x-api-key": WC_KEY } });
    const sj = await s.json();
    const hit = sj?.results?.[0];
    if (hit?.uuid) {
      const i = await fetch(`${WC_BASE}/watch/info?uuid=${hit.uuid}`, { headers: { "x-api-key": WC_KEY } });
      const ij = await i.json();
      if (ij && ij.median_asking_price != null) {
        result = {
          medianAsking: ij.median_asking_price, marketPrice: ij.market_price ?? null,
          dealerPrice: ij.dealer_price ?? null, model: ij.model ?? hit.model ?? null,
          brand: ij.brand ?? null, uuid: hit.uuid, confidence: hit.confidence ?? null, updated: ij.updated ?? null,
        };
      }
    }
  } catch { result = null; }
  marketCache[key] = result; cacheDirty = true;
  return result;
}

app.post("/api/market", async (req, res) => {
  if (!WC_KEY) return res.status(500).json({ error: "WATCHCHARTS_API_KEY not set" });
  const items = Array.isArray(req.body?.items) ? req.body.items.slice(0, 60) : [];
  const results = {};
  for (const it of items) {
    const brand = String(it?.brand || "").trim();
    const reference = String(it?.reference || "").trim();
    if (!brand || !reference) continue;
    results[mkKey(brand, reference)] = await lookupMarket(brand, reference);
  }
  saveCache();
  res.json({ results });
});

/* ---------- WatchOps: pull the dealer's live inventory (holdings) ---------- */
const WO_TOKEN = process.env.WATCHOPS_TOKEN;
const WO_BASE = "https://api.watchops.com/v1";
const woNum = (v) => { const n = parseFloat(String(v ?? "").replace(/[^0-9.\-]/g, "")); return isNaN(n) ? null : n; };
// WatchOps condition codes -> the wording WatchOps' own xlsx export uses for the same item,
// so API-sourced and sheet-sourced inventory share one vocabulary. Established by joining a
// full export against this endpoint on INVENTORY ID (2,146/2,146 rows agreed, no ambiguity).
// Code 5 is never observed; 10 behaves as a duplicate of 1.
const WO_CONDITION = { 4: "New", 3: "Retail Ready/Like New", 2: "Mint/Very Good", 1: "Used/Good", 10: "Used/Good" };
// WatchOps inventorytype codes -> label (confirmed by joining a full xlsx export on
// INVENTORY ID): 1 Regular/Owned · 2 Memo · 3 Partnership · 4 Consignment
const WO_INVTYPE = { 1: "Owned", 2: "Memo", 3: "Partnership", 4: "Consignment" };
function mapWoItem(x) {
  const pa = x.prodattr || {};
  const paid = x.purchase_invoice_paid;
  return {
    brand: x.brandname || null,
    modelName: pa.model || x.inventory_title || null,
    modelNumber: pa.reference || null,
    cost: woNum(x.total_cost) ?? woNum(x.purchaseprice),
    purchaseDate: x.purchasedate || null,
    targetWholesale: woNum(x.targetwholesaleprice),
    targetEndCustomer: woNum(x.targetendcustomerprice),
    tagPrice: woNum(x.tag_price),
    condition: x.condition != null ? (WO_CONDITION[x.condition] || String(x.condition)) : null,
    status: x.status || null,
    invType: WO_INVTYPE[x.inventorytype] || "Owned",
    paymentStatus: paid === true ? "Paid" : paid === false ? "Unpaid" : null,
    supplier: x.purchasefrom || null,
    serial: x.serialno || null,
  };
}
let _invCache = null;
app.get("/api/inventory", async (req, res) => {
  if (!WO_TOKEN) return res.status(500).json({ error: "WATCHOPS_TOKEN not set" });
  if (_invCache && Date.now() - _invCache.at < 30 * 60 * 1000 && !req.query.force) return res.json(_invCache.payload);
  try {
    // exclude sold and voided items — keep only live, valid holdings
    const isVoided = (x) => /void/i.test(String(x.status || "")) || x.voidreason != null || x.purchase_voidreason != null;
    // map+filter each page as it arrives and drop the raw payload — holding all
    // 2,600 raw records (with prodattr/images) just to return ~180 costs ~60MB
    const limit = 250; let page = 1, fetched = 0, total = Infinity, totalCost = null;
    const rows = [];
    while (fetched < total && page <= 40) {
      const r = await fetch(`${WO_BASE}/inventories`, {
        method: "POST", headers: { Authorization: `Bearer ${WO_TOKEN}`, "Content-Type": "application/json" },
        body: JSON.stringify({ status: "ALL", type: "ALL", limit, page_no: page }),
        signal: AbortSignal.timeout(30000),
      });
      if (!r.ok) { if (page === 1) return res.status(502).json({ error: `WatchOps returned ${r.status} — token may be expired` }); break; }
      const j = await r.json();
      if (!j || j.success !== 1 || !Array.isArray(j.data)) { if (page === 1) return res.status(502).json({ error: "WatchOps returned no data (check the token)" }); break; }
      if (page === 1) { total = parseInt(j.totalInventory || "0", 10) || j.data.length; totalCost = woNum(j.totalCost); }
      const n = j.data.length;
      for (const x of j.data) { if (!x.sold && !isVoided(x)) rows.push(mapWoItem(x)); }
      fetched += n;
      if (n < limit) break;
      page++;
    }
    const payload = { rows, count: rows.length, total, totalCost };
    _invCache = { at: Date.now(), payload };
    res.json(payload);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// serve the built frontend (created by `npm run build`) in production
const distPath = path.join(__dirname, "dist");
app.use(express.static(distPath));
app.get(/^(?!\/api).*/, (req, res) => {
  res.sendFile(path.join(distPath, "index.html"));
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
