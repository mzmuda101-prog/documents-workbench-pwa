// pwa-offline-playwright.js — zachowanie service workera (sw.js) na PRAWDZIWYM SW.
// Przeniesione z Sheet Workbench (2026-09-28) + krok aktualizacji apki.
//
// Własny serwer z logiem żądań i trybem „zawieszonej sieci", adres pwa.localhost (nie
// localhost — tam SW celowo zostawia stale-while-revalidate dla wygody developmentu).
//
// Co musi być prawdą:
//   1. drugie otwarcie NIE pyta serwera o pliki z ?v= (są niezmienne, idą z cache),
//   2. „jest kreska, a nic nie przechodzi" (serwer nie odpowiada na nawigację):
//      apka startuje z cache po ~3 s, a nie wisi na białym ekranie,
//   3. pełny offline: apka startuje i otwiera przykładowy .docx (biblioteki z cache),
//   4. nowa wersja sw.js → przycisk „Aktualizuj” → klik → strona wstaje pod NOWYM workerem.

const { chromium } = require("playwright");
const http = require("http");
const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const PORT = 4191;
const HOST = "pwa.localhost"; // *.localhost = bezpieczny kontekst (SW działa), ale nie „localhost" z sw.js
const ORIGIN = `http://${HOST}:${PORT}`;
const SEMANTIC_MODEL_CACHE = "dwb-embeddinggemma-2-q4-v1";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".css": "text/css", ".json": "application/json", ".png": "image/png", ".webp": "image/webp", ".woff2": "font/woff2", ".mp4": "video/mp4", ".svg": "image/svg+xml", ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" };
let log = [];
let hangNavigations = false;
let swOverride = null; // podmieniona treść sw.js (symulacja nowego wydania)
const hung = [];
const server = http.createServer((req, res) => {
  const url = new URL(req.url, ORIGIN);
  log.push(url.pathname + url.search);
  const isNav = url.pathname === "/" || url.pathname.endsWith(".html");
  if (hangNavigations && isNav) { hung.push(res); return; } // nigdy nie odpowiada
  if (url.pathname === "/sw.js" && swOverride) {
    res.writeHead(200, { "Content-Type": "text/javascript", "Cache-Control": "no-store" });
    res.end(swOverride);
    return;
  }
  let file = path.join(ROOT, decodeURIComponent(url.pathname));
  if (url.pathname === "/") file = path.join(ROOT, "index.html");
  if (!file.startsWith(ROOT) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream", "Cache-Control": "no-store" });
  fs.createReadStream(file).pipe(res);
});

async function run() {
  await new Promise((r) => server.listen(PORT, "127.0.0.1", r));
  const browser = await chromium.launch({
    headless: true,
    args: [`--host-resolver-rules=MAP ${HOST} 127.0.0.1`],
  });
  const context = await browser.newContext({ viewport: { width: 1100, height: 800 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));

  const results = [];
  const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
  const docRendered = () => page.evaluate(() => !!document.querySelector(".docx-preview-host p"));

  // Pierwsze wejście: instalacja SW + precache powłoki.
  await page.goto(ORIGIN + "/", { waitUntil: "load" });
  await page.evaluate(() => navigator.serviceWorker.ready);
  await page.reload({ waitUntil: "load" }); // teraz strona jest kontrolowana
  check("SW kontroluje stronę", await page.evaluate(() => !!navigator.serviceWorker.controller));
  // docx-preview/jszip + sample.docx trafiają do cache przy pierwszym użyciu
  await page.evaluate(() => loadSampleDocument());
  await page.waitForFunction(() => !!document.querySelector(".docx-preview-host p"), null, { timeout: 15000 });
  await sleep(1000);

  // 1) Drugie otwarcie: zero zapytań o pliki z ?v= (poza samym sw.js — to sprawdzanie aktualizacji)
  log = [];
  await page.reload({ waitUntil: "load" });
  await sleep(1500);
  const versioned = log.filter((u) => /[?&]v=/.test(u) && !u.startsWith("/sw.js"));
  check("ponowne otwarcie nie pyta sieci o pliki z ?v=", versioned.length === 0, versioned.slice(0, 5).join(", ") || `zapytań łącznie: ${log.length}`);

  // 2) Zawieszona sieć przy nawigacji → start z cache po limicie
  hangNavigations = true;
  const t0 = Date.now();
  await page.reload({ waitUntil: "load", timeout: 15000 });
  const waited = Date.now() - t0;
  const booted = await page.evaluate(() => !!document.getElementById("fileInput") && typeof ingestFile === "function");
  check("zawieszona sieć: start z cache w < 6 s", booted && waited < 6000, `${waited} ms`);
  hangNavigations = false;
  hung.splice(0).forEach((res) => { try { res.destroy(); } catch (_) {} });

  // 3) Pełny offline: start + otwarcie przykładowego dokumentu
  await context.setOffline(true);
  await page.reload({ waitUntil: "load" });
  check("offline: apka startuje", await page.evaluate(() => !!document.getElementById("fileInput")));
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => loadSampleDocument());
  await page.waitForFunction(() => !!document.querySelector(".docx-preview-host p"), null, { timeout: 15000 }).catch(() => {});
  check("offline: przykładowy .docx się otwiera", await docRendered());
  await context.setOffline(false);

  // 4) Nowe wydanie: inna treść sw.js pod tym samym adresem → przycisk → klik → nowy worker
  // Cache wag modelu nie jest cache'em powłoki i ma przetrwać aktualizację PWA.
  await page.evaluate(async (cacheName) => {
    const cache = await caches.open(cacheName);
    await cache.put("./__semantic-model-cache-test", new Response("kept"));
  }, SEMANTIC_MODEL_CACHE);
  const swSrc = fs.readFileSync(path.join(ROOT, "sw.js"), "utf8");
  swOverride = swSrc.replace(/CACHE_VERSION\s*=\s*"[^"]+"/, 'CACHE_VERSION = "29991231-99"');
  await page.reload({ waitUntil: "load" });
  await page.evaluate(() => navigator.serviceWorker.getRegistration().then((r) => r && r.update()));
  const btnShown = await page.waitForFunction(() => !document.getElementById("appUpdateBtn")?.classList.contains("hidden"), null, { timeout: 10000 }).then(() => true).catch(() => false);
  check("nowa wersja: przycisk „Aktualizuj” widoczny", btnShown);
  if (btnShown) {
    // Panel zostaje OTWARTY (od 1024 px stoi obok dokumentu) — przycisk w nagłówku
    // musi być klikalny bez zamykania czegokolwiek (dawniej zasłaniał go scrim).
    await page.evaluate(() => { window.__beforeUpdate = true; });
    const nav = page.waitForNavigation({ waitUntil: "load", timeout: 10000 }).then(() => true).catch(() => false);
    await page.click("#appUpdateBtn");
    const reloaded = await nav;
    await sleep(500);
    const state = await page.evaluate(async (cacheName) => ({
      fresh: !window.__beforeUpdate,
      controlled: !!navigator.serviceWorker.controller,
      newCache: (await caches.keys()).some((k) => k.includes("29991231-99")),
      modelCacheKept: !!(await (await caches.open(cacheName)).match("./__semantic-model-cache-test")),
    }), SEMANTIC_MODEL_CACHE);
    check("klik „Aktualizuj”: przeładowanie pod nowym workerem", reloaded && state.fresh && state.controlled && state.newCache, JSON.stringify(state));
    check("aktualizacja PWA zachowuje pobrany model", state.modelCacheKept, JSON.stringify(state));
  }
  swOverride = null;

  check("brak błędów strony", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();
  server.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ PWA: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); server.close(); process.exit(1); });
