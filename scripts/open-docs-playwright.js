// open-docs-playwright.js — kilka otwartych dokumentów naraz: karty (app/open-docs.js).
//
// Otwarcie przy otwartym dokumencie dokłada kartę (bez pytania o porzucenie zmian), wybór
// kilku plików = karty czekające na wczytanie, przełączanie zachowuje niezapisane zmiany
// (pisanie + operacja z panelu) i stan „niezapisane”, ten sam plik drugi raz = przełączenie,
// zamykanie kart, szkic odzyskiwania dla dokumentu w tle po zabiciu aplikacji, telefon.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const SAMPLES = path.join(__dirname, "..", "docs", "samples");
const S = (n) => path.join(SAMPLES, `${n}.docx`);
const errors = [];
const dialogs = [];

const idle = async (page) => {
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  await page.evaluate(() => dwbOpenDocs._idle());
  await sleep(350);
};

async function openPage(context, url = APP_URL) {
  const page = await context.newPage();
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => { dialogs.push(d.message()); d.accept(); });
  await page.goto(url, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await sleep(300);
  return page;
}

async function typeAtEnd(page, text) {
  await page.evaluate(() => {
    if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); }
  });
  await page.waitForFunction(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).some((p) => p.isContentEditable), null, { timeout: 10000 });
  await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.isContentEditable && p.textContent.trim().length > 3);
    placeCaret(ps[ps.length - 1], ps[ps.length - 1].textContent.length);
  });
  await page.keyboard.type(text, { delay: 5 });
  await sleep(200);
}

const tabs = (page) => page.evaluate(() => ({
  visible: !document.getElementById("docTabs").hidden,
  list: dwbOpenDocs.list(),
  dom: [...document.querySelectorAll("#docTabsRow .doc-tab")].map((el) => ({ name: el.querySelector(".doc-tab-name").textContent, active: el.classList.contains("is-active"), dirty: el.classList.contains("is-dirty"), waiting: el.classList.contains("is-waiting") })),
  file: currentFileName,
  dirty: hasUnsavedChanges,
  title: document.getElementById("heroTitle").textContent,
}));
const fileText = (page) => page.evaluate(async () => {
  const bytes = await buildDocumentForSave();
  const z = await JSZip.loadAsync(bytes);
  return (await z.file("word/document.xml").async("string")).replace(/<[^>]+>/g, "");
});
const clickTab = async (page, name) => {
  await page.click(`#docTabsRow .doc-tab:has(.doc-tab-name:text-is("${name}")) .doc-tab-btn`);
  await idle(page);
};

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));

  // ── 1. pierwszy dokument: bez kart ─────────────────────────────────────────
  let page = await openPage(context);
  await page.setInputFiles("#fileInput", S("headings-sample"));
  await idle(page);
  let st = await tabs(page);
  check("jeden dokument: karty schowane", !st.visible && st.list.length === 1, JSON.stringify(st));
  await typeAtEnd(page, " KARTA-A");
  await page.evaluate(() => applyDocumentEdit({ op: "replace", find: "Kaucja", replace: "Depozyt", regex: false, scope: "all" }));
  await idle(page);

  // ── 2. dwa kolejne pliki naraz (niezapisane zmiany w A — bez pytania) ─────
  dialogs.length = 0;
  await page.setInputFiles("#fileInput", [S("sample"), S("links-sample")]);
  await idle(page);
  st = await tabs(page);
  check("wybór 2 plików przy otwartym: 3 karty, pierwszy z wybranych na wierzchu, drugi czeka", st.visible && st.dom.length === 3 && st.file === "sample.docx" && st.dom[1].active && st.dom[2].waiting && st.dom[2].name === "links-sample", JSON.stringify(st.dom));
  check("otwieranie nie pyta o porzucenie zmian (A zostaje w karcie)", dialogs.length === 0, JSON.stringify(dialogs));
  check("karta A: kropka „niezapisane”", st.dom[0].name === "headings-sample" && st.dom[0].dirty && !st.dom[1].dirty, JSON.stringify(st.dom));
  check("nowy dokument czysty, nagłówek z jego nazwą", !st.dirty && st.title === "sample.docx", JSON.stringify({ dirty: st.dirty, title: st.title }));

  // ── 3. powrót do A: zmiany i „niezapisane” na miejscu ──────────────────────
  await clickTab(page, "headings-sample");
  st = await tabs(page);
  let txt = await fileText(page);
  check("powrót na kartę A: pisanie i operacja z panelu są w pliku", st.file === "headings-sample.docx" && txt.includes("KARTA-A") && txt.includes("Depozyt") && !txt.includes("Kaucja"), txt.slice(0, 80));
  check("powrót na kartę A: dalej niezapisane (Zapisz świeci)", st.dirty && (await page.evaluate(() => document.getElementById("heroSaveBtn").classList.contains("is-dirty"))), "");
  await typeAtEnd(page, " KARTA-A2");

  // ── 4. karta czekająca wczytuje się po stuknięciu; potem znów A ────────────
  await clickTab(page, "links-sample");
  st = await tabs(page);
  check("karta czekająca: wczytuje się po stuknięciu", st.file === "links-sample.docx" && !st.dom[2].waiting && st.dom[2].active, JSON.stringify(st.dom));
  await clickTab(page, "headings-sample");
  txt = await fileText(page);
  check("dwa przełączenia: wszystkie zmiany A zostają", txt.includes("KARTA-A") && txt.includes("KARTA-A2") && txt.includes("Depozyt"), "");
  const undo = await page.evaluate(() => ({ undoDisabled: document.getElementById("undoBtn")?.disabled }));
  check("po przełączeniu Cofnij zaczyna od nowa (bez cofania cudzych zmian)", undo.undoDisabled !== false, JSON.stringify(undo));

  // ── 5. ten sam plik drugi raz → przełączenie, bez duplikatu ────────────────
  await clickTab(page, "sample");
  await page.setInputFiles("#fileInput", S("links-sample"));
  await idle(page);
  st = await tabs(page);
  check("ten sam plik drugi raz: przełączenie na jego kartę, bez duplikatu", st.dom.length === 3 && st.file === "links-sample.docx", JSON.stringify(st.dom));

  // ── 6. zabicie aplikacji z niezapisanym A w tle → szkic do odzyskania ─────
  await page.evaluate(async () => {
    Object.defineProperty(document, "visibilityState", { configurable: true, get: () => "hidden" });
    document.dispatchEvent(new Event("visibilitychange"));
    await dwbDrafts._flush();
  });
  await page.close({ runBeforeUnload: false });
  page = await openPage(context);
  await sleep(700);
  const cards = await page.evaluate(() => [...document.querySelectorAll("#draftRecovery .draft-card .draft-name")].map((n) => n.textContent));
  check("po zabiciu: niezapisany dokument z karty W TLE do odzyskania", cards.length === 1 && cards[0] === "headings-sample.docx", JSON.stringify(cards));
  await page.click("#draftRecovery .btn.primary");
  await idle(page);
  txt = await fileText(page);
  check("odzyskany dokument z tła ma wszystkie zmiany", txt.includes("KARTA-A2") && txt.includes("Depozyt"), "");

  // ── 7. zamykanie kart ──────────────────────────────────────────────────────
  await page.setInputFiles("#fileInput", [S("sample"), S("review-sample")]);
  await idle(page);
  st = await tabs(page);
  check("odzyskany + 2 nowe = 3 karty", st.dom.length === 3, JSON.stringify(st.dom));
  dialogs.length = 0;
  await page.click('#docTabsRow .doc-tab:has(.doc-tab-name:text-is("headings-sample")) .doc-tab-close');
  await idle(page);
  st = await tabs(page);
  check("zamknięcie niezapisanej karty w tle: pyta, karta znika, bieżący bez zmian", dialogs.length === 1 && st.dom.length === 2 && st.file === "sample.docx", JSON.stringify({ dialogs, dom: st.dom, file: st.file }));
  await page.click('#docTabsRow .doc-tab.is-active .doc-tab-close');
  await idle(page);
  st = await tabs(page);
  check("zamknięcie karty na wierzchu: na wierzch wchodzi sąsiednia, karty chowają się przy 1", st.file === "review-sample.docx" && !st.visible && st.list.length === 1, JSON.stringify(st));
  const noDrafts = await page.evaluate(async () => (await dwbDrafts.recoverable()).length);
  check("zamknięte karty nie zostawiają szkiców", noDrafts === 0, String(noDrafts));

  // ── 8. telefon: rząd kart przewija się w bok z wygaszeniem ─────────────────
  await page.setViewportSize({ width: 390, height: 800 });
  await page.setInputFiles("#fileInput", ["sample", "links-sample", "forms-sample", "styled-sample", "przewodnik"].map(S));
  await idle(page);
  const phone = await page.evaluate(() => {
    const row = document.getElementById("docTabsRow");
    const r = document.getElementById("docTabs").getBoundingClientRect();
    return { n: row.children.length, scroll: row.scrollWidth > row.clientWidth, fade: row.classList.contains("more-r") || row.classList.contains("more-l"), fits: r.right <= innerWidth + 0.5, pageScrollX: document.documentElement.scrollWidth <= innerWidth };
  });
  check("telefon: 6 kart w jednym rzędzie, przewijanie w bok z wygaszeniem, bez poziomego przewijania strony", phone.n === 6 && phone.scroll && phone.fade && phone.fits && phone.pageScrollX, JSON.stringify(phone));

  await browser.close();
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Karty dokumentów [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
