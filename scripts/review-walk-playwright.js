// review-walk-playwright.js — przegląd zmian po kolei (app/review-walk.js): licznik „n z N”, zamiana
// jako jedna decyzja, Enter = akceptuj, Backspace = odrzuć (w pliku, bez przerysowania), Cofnij,
// Esc; na telefonie przesunięcie karty palcem (w prawo akceptuj, w lewo odrzuć, krótko — nic).
//   node scripts/review-walk-playwright.js        ENGINE=webkit node scripts/review-walk-playwright.js
const fs = require("fs");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const SAMPLE = fs.readFileSync("docs/samples/review-sample.docx");
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(500));

async function open(page, name) {
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name, mimeType: DOCX_MIME, buffer: SAMPLE });
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);
  await page.evaluate(() => ensureLazyFeature("review"));
  await page.waitForFunction(() => typeof dwbWalk !== "undefined");
  await page.evaluate(() => {
    window.__renders = 0;
    const orig = window.renderCurrentDocument;
    window.renderCurrentDocument = function counted(...a) { window.__renders++; return orig.apply(this, a); };
  });
}
const state = (page) => page.evaluate(async () => {
  const bar = document.querySelector(".walk-bar");
  const scan = await scanDocxRevisions(originalFileBytes);
  return {
    shown: !!bar && !bar.hidden,
    count: bar?.querySelector(".walk-count")?.textContent || "",
    what: bar?.querySelector(".walk-what")?.textContent || "",
    units: dwbWalk._units().length,
    at: dwbWalk._at(),
    revs: scan.changes.length,
    focus: !!bar?.contains(document.activeElement),
    para: document.querySelector(".walk-para")?.textContent.slice(0, 20) || "",
    hl: typeof CSS !== "undefined" && CSS.highlights ? CSS.highlights.has("dwb-walk") : null,
    renders: window.__renders,
  };
});
const fileText = (page, start) => page.evaluate(async (start) => {
  const d = await getDocumentXmlDom(originalFileBytes);
  const p = collectParagraphElements(d.documentElement, "all").find((x) => x.textContent.startsWith(start));
  return p ? { text: p.textContent, ins: p.getElementsByTagNameNS(W_NS, "ins").length, del: p.getElementsByTagNameNS(W_NS, "del").length } : null;
}, start);

async function desktop(browser) {
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await open(page, "przeglad.docx");

  // start z panelu Recenzja: „Przejrzyj po kolei”
  await page.evaluate(() => document.getElementById("rvWalkBtn").click());
  await page.waitForSelector(".walk-bar:not([hidden])");
  let s = await state(page);
  check("start: pasek, „Zmiana 1 z N”, fokus w pasku", s.shown && s.count === `Zmiana 1 z ${s.units}` && s.focus, JSON.stringify(s));
  check("zamiana „14 → 30” i znak akapitu przy swojej treści to jedna decyzja (mniej decyzji niż poprawek)", s.units < s.revs, `${s.units} z ${s.revs}`);
  const N = s.units;
  // do zamiany „14” → „30” strzałką w prawo
  for (let k = 0; k < N && !/„14” → „30”/.test(s.what); k++) { await page.keyboard.press("ArrowRight"); s = await state(page); }
  check("strzałki: zamiana „14” → „30”, akapit wskazany, wstawka podświetlona", /„14” → „30”/.test(s.what) && s.para.startsWith("Termin") && s.hl !== false, JSON.stringify(s));
  const at0 = s.at;
  // Enter = akceptuj i dalej
  await page.keyboard.press("Enter");
  await idle(page);
  s = await state(page);
  const t1 = await fileText(page, "Termin");
  check("Enter: zamiana przyjęta w pliku (bez w:ins / w:del, „30”)", t1 && !t1.ins && !t1.del && /wynosi 30 dni/.test(t1.text), JSON.stringify(t1));
  check("…licznik N−1, ta sama pozycja = następna zmiana, fokus dalej w pasku", s.units === N - 1 && s.at === at0 && s.count === `Zmiana ${at0 + 1} z ${N - 1}` && s.focus, JSON.stringify(s));
  check("…bez przerysowania dokumentu", s.renders === 0, String(s.renders));
  // Backspace = odrzuć
  const before = s.what;
  await page.keyboard.press("Backspace");
  await idle(page);
  s = await state(page);
  check("Backspace: odrzucona, licznik N−2, tekst dokumentu nie dostał klawisza", s.units === N - 2 && s.what !== before && s.focus, JSON.stringify(s));
  // Cofnij — decyzja wraca, lista od nowa
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  await page.waitForTimeout(500);
  s = await state(page);
  check("Cofnij w trakcie przeglądu: zmiana wraca na listę (N−1)", s.units === N - 1 && s.shown, JSON.stringify(s));
  // Esc — koniec
  await page.evaluate(() => document.querySelector(".walk-bar").focus());
  await page.keyboard.press("Escape");
  s = await state(page);
  check("Esc: pasek zamknięty, bez obwódki i podświetlenia", !s.shown && !s.para && !s.hl, JSON.stringify(s));
  // przegląd do końca: same Entery → „Wszystkie zmiany przejrzane”, pasek znika
  await page.evaluate(() => dwbWalk.start());
  for (let k = 0; k < 20 && (await state(page)).shown; k++) { await page.keyboard.press("Enter"); await idle(page); }
  s = await state(page);
  check("do końca: wszystkie przyjęte, pasek znika", !s.shown && s.revs === 0, JSON.stringify(s));
  check("bez błędów strony (komputer)", !errors.length, errors.join(" | "));
  await ctx.close();
}

async function phone(browser) {
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 }, hasTouch: true, ...(ENGINE === "chromium" ? { isMobile: true } : {}) });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await open(page, "przeglad-telefon.docx");
  await page.evaluate(() => dwbWalk.start());
  await page.waitForSelector(".walk-bar:not([hidden])");
  const box = await page.evaluate(() => { const r = document.querySelector(".walk-bar").getBoundingClientRect(); const b = [...document.querySelectorAll(".walk-btn")].map((x) => x.getBoundingClientRect()); return { l: r.left, r: r.right, vw: innerWidth, minBtn: Math.min(...b.map((x) => Math.min(x.width, x.height))), scroll: document.documentElement.scrollWidth }; });
  check("telefon: pasek mieści się na ekranie, przyciski ≥ 44 px", box.l >= 0 && box.r <= box.vw && box.minBtn >= 44 && box.scroll <= box.vw, JSON.stringify(box));
  // przesunięcie karty palcem (zdarzenia wskaźnika „touch”)
  const swipe = (dx) => page.evaluate(async (dx) => {
    const card = document.querySelector(".walk-card");
    const r = card.getBoundingClientRect();
    const x = r.left + r.width / 2, y = r.top + r.height / 2;
    const ev = (type, cx) => card.dispatchEvent(new PointerEvent(type, { pointerType: "touch", pointerId: 7, clientX: cx, clientY: y, bubbles: true, isPrimary: true }));
    ev("pointerdown", x);
    for (let k = 1; k <= 6; k++) { ev("pointermove", x + (dx * k) / 6); await new Promise((r) => setTimeout(r, 16)); }
    ev("pointerup", x + dx);
  }, dx);
  let s = await state(page);
  const N = s.units, what0 = s.what;
  await swipe(40);
  await page.waitForTimeout(400);
  s = await state(page);
  check("krótkie przesunięcie (40 px): nic się nie dzieje, karta wraca", s.units === N && s.what === what0, JSON.stringify(s));
  await swipe(160);
  await idle(page);
  s = await state(page);
  const accepted = s.units === N - 1;
  await swipe(-160);
  await idle(page);
  s = await state(page);
  check("przesunięcie w prawo: akceptuj, w lewo: odrzuć (licznik −1 za każdym razem)", accepted && s.units === N - 2, JSON.stringify(s));
  const rev = await page.evaluate(async () => (await scanDocxRevisions(originalFileBytes)).changes.length);
  check("…decyzje są w pliku", rev === s.revs && s.units < N, String(rev));
  check("bez błędów strony (telefon)", !errors.length, errors.join(" | "));
  await ctx.close();
}

(async () => {
  const browser = await pw[ENGINE].launch({ headless: true });
  try {
    await desktop(browser);
    await phone(browser);
  } catch (e) {
    check("przebieg bez wyjątku", false, String(e?.stack || e).slice(0, 600));
  }
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})();
