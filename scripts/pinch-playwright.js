// pinch-playwright.js — przybliżanie dokumentu dwoma palcami (app/pinch-zoom.js).
//
//   node scripts/pinch-playwright.js                (Chromium, telefon z dotykiem)
//   ENGINE=webkit node scripts/pinch-playwright.js  (WebKit)
//
// Gest jest syntetyczny (zdarzenia touch* z ręcznie ustawionym `touches`) — działa
// w obu silnikach; prawdziwe palce sprawdza się na symulatorze/telefonie.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

// Dwa palce: rozstaw d, środek (cx, cy). type: touchstart | touchmove | touchend.
async function fingers(page, type, d, cx = 195, cy = 400) {
  await page.evaluate(([type, d, cx, cy]) => {
    const vp = document.getElementById("docViewport");
    const pt = (x, y) => ({ clientX: x, clientY: y });
    const touches = type === "touchend" ? [] : [pt(cx - d / 2, cy), pt(cx + d / 2, cy)];
    const ev = new Event(type, { bubbles: true, cancelable: true });
    Object.defineProperty(ev, "touches", { value: touches });
    vp.dispatchEvent(ev);
  }, [type, d, cx, cy]);
}
async function pinch(page, from, to, cx, cy) {
  await fingers(page, "touchstart", from, cx, cy);
  for (let i = 1; i <= 4; i++) await fingers(page, "touchmove", from + ((to - from) * i) / 4, cx, cy);
  await page.waitForTimeout(60);
  await fingers(page, "touchend", to, cx, cy);
  await page.waitForTimeout(60);
}
const state = (page) => page.evaluate(() => ({
  zoom: parseFloat(document.getElementById("docCanvas").style.getPropertyValue("--doc-zoom")) || 1,
  now: document.getElementById("zoomNow").textContent,
  mode: getZoomMode(),
  reflow: document.getElementById("docCanvas").classList.contains("doc-reflow-mode"),
  sw: document.getElementById("docViewport").scrollWidth - document.getElementById("docViewport").clientWidth,
  transform: document.getElementById("docZoomShell").style.transform,
}));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({
    serviceWorkers: "block",
    viewport: { width: 390, height: 844 },
    hasTouch: true,
    ...(ENGINE === "chromium" ? { isMobile: true } : {}),
  });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));

  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.waitForTimeout(400);

  const s0 = await state(page);
  check("start: tryb Dopasuj (reflow), 100%", s0.mode === "fit" && s0.reflow && s0.zoom === 1, JSON.stringify(s0));
  check("touch-action: pan-x pan-y na dokumencie (przeglądarka nie przybliża sama)",
    await page.evaluate(() => getComputedStyle(document.getElementById("docViewport")).touchAction === "pan-x pan-y"));

  await pinch(page, 100, 200);
  const s1 = await state(page);
  check("rozsunięcie palców przybliża (~2×) mimo Dopasuj", s1.zoom > 1.8 && s1.zoom <= 2.05, JSON.stringify(s1));
  check("zostaje w trybie Dopasuj (tekst zawija się)", s1.mode === "fit" && s1.reflow);
  check("bez poziomego przewijania po przybliżeniu", s1.sw <= 1, String(s1.sw));
  check("pasek pokazuje procent zamiast „Dopasuj”", /^\d+%$/.test(s1.now), s1.now);
  check("po geście transform tymczasowy zdjęty", s1.transform === "");

  await pinch(page, 200, 100);
  const s2 = await state(page);
  check("zsunięcie palców wraca do ~1×", Math.abs(s2.zoom - s1.zoom / 2) < 0.1, JSON.stringify(s2));

  await pinch(page, 300, 20);
  const s3 = await state(page);
  check("dół zakresu 0,35", s3.zoom >= 0.35 && s3.zoom <= 0.4, String(s3.zoom));
  await pinch(page, 20, 600);
  const s4 = await state(page);
  check("góra zakresu 3×", s4.zoom <= 3.001 && s4.zoom >= 2.9, String(s4.zoom));

  await page.click("#zoomFitBtn");
  await page.waitForTimeout(300);
  const s5 = await state(page);
  check("„Dopasuj” wraca do 100% i etykiety „Dopasuj”", s5.zoom === 1 && s5.mode === "fit" && !/%/.test(s5.now), JSON.stringify(s5));

  // jeden palec = zwykłe przewijanie, nic się nie zmienia
  await fingers(page, "touchstart", 0);
  const one = await page.evaluate(() => {
    const vp = document.getElementById("docViewport");
    const ev = new Event("touchstart", { bubbles: true, cancelable: true });
    Object.defineProperty(ev, "touches", { value: [{ clientX: 10, clientY: 10 }] });
    vp.dispatchEvent(ev);
    return document.getElementById("docZoomShell").style.transform;
  });
  check("jeden palec nie startuje gestu", one === "");

  // tablet: ręczny zoom, pinch przełącza fit → manual bez przebudowy
  await page.setViewportSize({ width: 820, height: 1000 });
  await page.waitForTimeout(400);
  await page.click("#zoomFitBtn");
  await page.waitForTimeout(300);
  const t0 = await state(page);
  await pinch(page, 100, 160, 410, 500);
  const t1 = await state(page);
  check("tablet: pinch zmienia zoom i przechodzi w tryb ręczny", t1.mode === "manual" && t1.zoom > t0.zoom * 1.4, JSON.stringify({ t0: t0.zoom, t1: t1.zoom, mode: t1.mode }));

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ pinch-zoom [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
