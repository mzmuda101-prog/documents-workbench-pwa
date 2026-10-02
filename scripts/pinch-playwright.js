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
  // to, co widać: wysokość linii tekstu na ekranie (nie tylko zmienna CSS)
  lineH: (() => { const p = [...document.querySelectorAll(".docx-preview-host p")].find((e) => e.textContent.trim().length > 40); const r = document.createRange(); r.selectNodeContents(p); return r.getClientRects()[0]?.height || 0; })(),
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

  // w trakcie gestu: plakietka z procentem nad dokumentem + ten sam procent na pasku (jak Word)
  await fingers(page, "touchstart", 100);
  await fingers(page, "touchmove", 150);
  await page.waitForTimeout(80);
  const mid = await page.evaluate(() => {
    const b = document.querySelector(".zoom-badge");
    return { on: b.classList.contains("is-on"), text: b.textContent, bar: document.getElementById("zoomNow").textContent, op: getComputedStyle(b).opacity };
  });
  check("w trakcie gestu: plakietka „150%” nad dokumentem", mid.on && mid.text === "150%", JSON.stringify(mid));
  check("w trakcie gestu: pasek też pokazuje 150%", mid.bar === "150%", mid.bar);
  await fingers(page, "touchmove", 104); // ~104% → przyciąga do 100%
  await page.waitForTimeout(60);
  const snap = await page.evaluate(() => ({ text: document.querySelector(".zoom-badge").textContent, snap: document.querySelector(".zoom-badge").classList.contains("is-snap") }));
  check("w pobliżu 100% przyciąga do równych 100% (wyróżnione)", snap.text === "100%" && snap.snap, JSON.stringify(snap));
  await fingers(page, "touchend", 104);
  await page.waitForTimeout(900);
  check("po puszczeniu palców plakietka gaśnie", await page.evaluate(() => !document.querySelector(".zoom-badge").classList.contains("is-on")));
  check("przyciągnięte 100% zatwierdzone", (await state(page)).zoom === 1);

  await pinch(page, 100, 200);
  const s1 = await state(page);
  check("rozsunięcie palców przybliża (~2×) mimo Dopasuj", s1.zoom > 1.8 && s1.zoom <= 2.05, JSON.stringify(s1));
  check("zostaje w trybie Dopasuj (tekst zawija się)", s1.mode === "fit" && s1.reflow);
  check("tekst NA EKRANIE jest ~2× większy (nie tylko zmienna)", s1.lineH > s0.lineH * 1.7, `${s0.lineH} → ${s1.lineH}`);
  check("bez poziomego przewijania po przybliżeniu", s1.sw <= 1, String(s1.sw));
  check("pasek pokazuje procent zamiast „Dopasuj”", /^\d+%$/.test(s1.now), s1.now);
  check("po geście transform tymczasowy zdjęty", s1.transform === "");

  await pinch(page, 200, 100);
  const s2 = await state(page);
  check("zsunięcie palców wraca do ~1×", Math.abs(s2.zoom - s1.zoom / 2) < 0.1, JSON.stringify(s2));
  check("tekst na ekranie wraca do pierwotnej wielkości", Math.abs(s2.lineH - s0.lineH) < s0.lineH * 0.15, `${s0.lineH} / ${s2.lineH}`);

  await pinch(page, 200, 110); // ~55%
  await pinch(page, 110, 97); // ~48% → dół Widoku mobilnego, ale jeszcze nie „na strony”
  const s3 = await state(page);
  check("dół zakresu Widoku mobilnego 50% (zostaje w Widoku mobilnym)", s3.zoom === 0.5 && s3.reflow, JSON.stringify(s3));
  await pinch(page, 20, 600);
  const s4 = await state(page);
  check("góra zakresu 3×", s4.zoom <= 3.001 && s4.zoom >= 2.9, String(s4.zoom));

  // miejsce pod palcami zostaje pod palcami (tekst w „Dopasuj” zawija się na nowo)
  await page.click("#zoomFitBtn");
  await page.waitForTimeout(300);
  await page.evaluate(() => { document.getElementById("docViewport").scrollTop = 400; });
  await page.waitForTimeout(100);
  const probe = () => page.evaluate(() => {
    const r = document.caretRangeFromPoint ? document.caretRangeFromPoint(195, 500) : null;
    if (r) return r.startContainer.textContent.slice(r.startOffset, r.startOffset + 12);
    const p = document.caretPositionFromPoint(195, 500);
    return p.offsetNode.textContent.slice(p.offset, p.offset + 12);
  });
  const wordBefore = await probe();
  const markBefore = await page.evaluate(() => {
    const r = document.caretRangeFromPoint ? document.caretRangeFromPoint(195, 500) : (() => { const p = document.caretPositionFromPoint(195, 500); const x = document.createRange(); x.setStart(p.offsetNode, p.offset); return x; })();
    window.__anchorNode = r.startContainer; window.__anchorOff = r.startOffset;
    return true;
  });
  await pinch(page, 100, 200, 195, 500);
  const drift = await page.evaluate(() => {
    const r = document.createRange();
    r.setStart(window.__anchorNode, window.__anchorOff);
    r.setEnd(window.__anchorNode, Math.min(window.__anchorOff + 1, window.__anchorNode.length));
    const b = r.getBoundingClientRect();
    return { dy: Math.round(b.top + b.height / 2 - 500), top: document.getElementById("docViewport").scrollTop };
  });
  check("tekst pod palcami zostaje pod palcami (bez skoku na początek)", markBefore && Math.abs(drift.dy) < 60 && drift.top > 400, `${wordBefore} ${JSON.stringify(drift)}`);

  await page.click("#zoomFitBtn");
  await page.waitForTimeout(300);
  const s5 = await state(page);
  check("„Dopasuj” wraca do 100% (etykieta „100%” z ikonką Widoku mobilnego)", s5.zoom === 1 && s5.mode === "fit" && s5.now === "100%"
    && await page.evaluate(() => document.getElementById("zoomNow").dataset.layout === "mobile"), JSON.stringify(s5));

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

  // zsunięcie palców wyraźnie poniżej 50% = „cała strona”: Widok desktopowy (jak w zdjęciach)
  await page.click("#zoomFitBtn");
  await page.waitForTimeout(300);
  await page.evaluate(() => { document.getElementById("docViewport").scrollTop = 1200; });
  await page.waitForTimeout(250);
  const topBefore = await page.evaluate(() => { const vp = document.getElementById("docViewport"); const t = vp.getBoundingClientRect().top; return collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).findIndex((p) => p.getBoundingClientRect().bottom > t + 1); });
  await fingers(page, "touchstart", 200);
  await fingers(page, "touchmove", 120);
  await fingers(page, "touchmove", 70); // 35% — poniżej 80% minimum
  await page.waitForTimeout(80);
  const hint = await page.evaluate(() => document.querySelector(".zoom-badge").textContent);
  check("plakietka uprzedza: „Puść — Widok desktopowy”", hint === "Puść — Widok desktopowy", hint);
  await fingers(page, "touchend", 70);
  await page.waitForFunction(() => document.getElementById("loadingOverlay").classList.contains("hidden") && getViewLayout() === "desktop", null, { timeout: 15000 });
  await page.waitForTimeout(500);
  const pg = await state(page);
  const topAfter = await page.evaluate(() => { const vp = document.getElementById("docViewport"); const t = vp.getBoundingClientRect().top; return collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).findIndex((p) => p.getBoundingClientRect().bottom > t + 1); });
  check("po puszczeniu: strony, cała kartka na szerokość, bez przewijania w bok", !pg.reflow && pg.zoom < 0.6 && pg.sw <= 1, JSON.stringify(pg));
  check("…w tym samym miejscu dokumentu", topBefore > 0 && Math.abs(topAfter - topBefore) <= 1, `${topBefore} → ${topAfter}`);
  check("dół zakresu stron 25%", await page.evaluate(() => getZoomLimits().min === 0.25));
  await page.evaluate(() => setViewLayoutPref("auto"));
  await page.waitForTimeout(800);

  // tablet w pionie (Auto): Widok mobilny — pinch skaluje tekst, który dalej się zawija
  await page.setViewportSize({ width: 820, height: 1000 });
  await page.waitForTimeout(600);
  await page.click("#zoomFitBtn");
  await page.waitForTimeout(300);
  const p0 = await state(page);
  check("tablet w pionie: Widok mobilny (tekst zawija się)", p0.mode === "fit" && p0.reflow, JSON.stringify(p0));
  await pinch(page, 100, 140, 410, 500);
  const p1 = await state(page);
  check("tablet w pionie: pinch powiększa tekst, bez poziomego przewijania", p1.reflow && p1.lineH > p0.lineH * 1.25 && p1.sw <= 1, JSON.stringify({ p0: p0.lineH, p1: p1.lineH, sw: p1.sw }));
  await page.click("#zoomFitBtn");

  // tablet w poziomie (Auto): Widok desktopowy — strony, pinch przybliża jak zdjęcie
  await page.setViewportSize({ width: 1180, height: 820 });
  await page.waitForTimeout(800);
  await page.waitForFunction(() => !document.getElementById("docCanvas").classList.contains("doc-reflow-mode"));
  await page.waitForTimeout(300);
  const t0 = await state(page);
  await pinch(page, 100, 160, 590, 500);
  const t1 = await state(page);
  check("tablet w poziomie: Widok desktopowy, pinch zmienia zoom stron", t1.mode === "manual" && !t1.reflow && t1.zoom > t0.zoom * 1.4, JSON.stringify({ t0: t0.zoom, t1: t1.zoom, mode: t1.mode }));
  check("tablet w poziomie: tekst na ekranie faktycznie większy", t1.lineH > t0.lineH * 1.4, `${t0.lineH} → ${t1.lineH}`);

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
