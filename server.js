import "dotenv/config";
import express from "express";
import cors from "cors";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { google } from "googleapis";
import * as XLSX from "xlsx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const app = express();
app.use(cors());
app.use(express.json({ limit: "1mb" }));

/* ---------- Google Drive: pull recent WhatsApp message sheets ---------- */
const DRIVE_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;
function driveCreds() {
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) return JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const p = path.join(__dirname, "service-account.json");
  if (fs.existsSync(p)) return JSON.parse(fs.readFileSync(p, "utf8"));
  return null;
}
let _drive = null;
function getDrive() {
  if (_drive) return _drive;
  const creds = driveCreds();
  if (!creds) return null;
  const auth = new google.auth.GoogleAuth({ credentials: creds, scopes: ["https://www.googleapis.com/auth/drive.readonly"] });
  _drive = google.drive({ version: "v3", auth });
  return _drive;
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
    messageBody: g("messageBody", "body"),
    timestamp: g("timestamp", "date"),
    sender: g("senderName", "sender"),
    chat: g("chatName", "chat"),
  };
}
async function fetchFileRows(drive, file) {
  let buf;
  if (file.mimeType === "application/vnd.google-apps.spreadsheet") {
    const res = await drive.files.export({ fileId: file.id, mimeType: "text/csv" }, { responseType: "arraybuffer" });
    return XLSX.utils.sheet_to_json(XLSX.read(Buffer.from(res.data), { type: "buffer" }).Sheets.Sheet1 || Object.values(XLSX.read(Buffer.from(res.data), { type: "buffer" }).Sheets)[0], { defval: null });
  }
  const res = await drive.files.get({ fileId: file.id, alt: "media" }, { responseType: "arraybuffer" });
  buf = Buffer.from(res.data);
  const wb = XLSX.read(buf, { type: "buffer" });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });
}

const _msgCache = {}; // days -> { at, payload }
const MSG_TTL = 30 * 60 * 1000; // 30 min — the folder updates once a day
app.get("/api/messages", async (req, res) => {
  const drive = getDrive();
  if (!drive || !DRIVE_FOLDER_ID) return res.status(500).json({ error: "Google Drive not configured" });
  const days = Math.min(30, Math.max(1, parseInt(req.query.days, 10) || 5));
  const cached = _msgCache[days];
  if (cached && Date.now() - cached.at < MSG_TTL && !req.query.force) return res.json(cached.payload);
  try {
    let files = [], token;
    do {
      const r = await drive.files.list({
        q: `'${DRIVE_FOLDER_ID}' in parents and trashed=false`,
        fields: "nextPageToken,files(id,name,mimeType)", pageSize: 1000, pageToken: token,
        supportsAllDrives: true, includeItemsFromAllDrives: true,
      });
      files = files.concat(r.data.files || []); token = r.data.nextPageToken;
    } while (token);
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
        const raw = await fetchFileRows(drive, f);
        raw.forEach((rr) => {
          const n = normalizeMsgRow(rr);
          const it = String(n.intent || "").toLowerCase();
          if (it === "buy" || it === "sell") rows.push(n);
        });
      } catch { /* skip a bad file */ }
    }
    const payload = { rows, files: recent.map((f) => f.name), latest, days };
    _msgCache[days] = { at: Date.now(), payload };
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

// serve the built frontend (created by `npm run build`) in production
const distPath = path.join(__dirname, "dist");
app.use(express.static(distPath));
app.get(/^(?!\/api).*/, (req, res) => {
  res.sendFile(path.join(distPath, "index.html"));
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
