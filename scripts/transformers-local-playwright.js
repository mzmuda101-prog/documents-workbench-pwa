// Biblioteka AI (Transformers.js + onnxruntime-web) z lib/transformers zamiast esm.sh/jsdelivr
// (audyt bezpieczeństwa 2026-11-06): import w workerze modułowym, ścieżki wasm lokalne, żadnych
// żądań do CDN; workery apki nie odwołują się już do esm.sh.
//   node scripts/transformers-local-playwright.js            (bez sieci: import + ścieżki)
//   FULL=1 node scripts/transformers-local-playwright.js     (+ prawdziwy model z Hugging Face, wasm)
const playwright = require("playwright");
const fs = require("fs");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE || "chromium";
const FULL = !!process.env.FULL;

async function run() {
  const fails = [];
  const check = (name, ok, detail) => { console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail).slice(0, 600)}`}`); if (!ok) fails.push(name); };
  for (const f of ["app/local-ai-worker.js", "app/semantic-search-worker.js"]) {
    const src = fs.readFileSync(path.join(__dirname, "..", f), "utf8");
    check(`${f}: bez importu z CDN, lib/transformers`, !/import\(\s*["']https?:/.test(src) && !/"https:\/\/esm\.sh/.test(src) && src.includes("../lib/transformers/transformers.min.mjs"), f);
  }
  const browser = await playwright[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block" });
  const cdn = [];
  await context.route(/^https:\/\/(esm\.sh|cdn\.jsdelivr\.net)\//, (r) => { cdn.push(r.request().url()); r.abort(); });
  const page = await context.newPage();
  await page.goto(APP_URL + "scripts/fixtures/esm-import-worker.js", { waitUntil: "load" }); // ten sam origin, bez startu apki
  const out = await page.evaluate((full) => new Promise((res) => {
    const w = new Worker("/scripts/fixtures/transformers-local-worker.js", { type: "module" });
    w.onmessage = (e) => { res(e.data); w.terminate(); };
    w.onerror = (e) => res({ error: `onerror ${e.message || ""}` });
    setTimeout(() => res({ error: "timeout" }), full ? 240000 : 30000);
    w.postMessage({ full });
  }), FULL);
  check("worker: biblioteka z lib/ zaimportowana", out.imported, out);
  check("worker: wasm z lib/transformers", out.paths && /\/lib\/transformers\/ort-wasm-simd-threaded\.asyncify\.wasm$/.test(out.paths.wasm), out.paths || out);
  if (FULL) check("model (wasm): wektor 384, znormalizowany", out.dims && out.dims[1] === 384 && Math.abs(out.norm - 1) < 1e-3, out);
  check("żadnych żądań do esm.sh / jsdelivr", !cdn.length, cdn);
  await browser.close();
  if (fails.length) { console.log(`❌ transformers-local (${ENGINE}): ${fails.length} błędów`); process.exit(1); }
  console.log(`✅ transformers-local (${ENGINE}${FULL ? ", FULL" : ""}): wszystko OK`);
}
run().catch((e) => { console.error(e); process.exit(1); });
