import "dotenv/config";
import express from "express";
import cors from "cors";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const app = express();
app.use(cors());
app.use(express.json({ limit: "1mb" }));

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
