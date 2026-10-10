// track-changes-playwright.js — „Śledź zmiany” i „Utwórz dokument z poprawkami” (2026-10-10).
//
// Jak w Wordzie: przełącznik w Recenzji (i Ctrl/⌘+Shift+E), imię jako podpis poprawek, paski
// „Prostych znaczników” przy zmienionych akapitach, zapis na dysk z w:ins/w:del i w:trackRevisions
// (w aplikacji czysty stan — pisze się dalej), „Pokaż poprawki w dokumencie”, wyłączenie (poprawki
// w dokumencie, przegląd w Recenzji), Odrzuć wszystkie = oryginał, plik z Worda ze śledzeniem —
// włącza się sam, stan śledzenia w kartach, dokument z poprawkami z Porównaj wersje. ENGINE=webkit.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

const xmlOf = (page, expr) => page.evaluate(async (expr) => {
  // eslint-disable-next-line no-eval
  const bytes = await eval(expr);
  const z = await JSZip.loadAsync(bytes);
  return { doc: await z.file("word/document.xml").async("string"), settings: z.file("word/settings.xml") ? await z.file("word/settings.xml").async("string") : "" };
}, expr);
const bodyTexts = (page, mode) => page.evaluate(async (mode) => {
  const z = await JSZip.loadAsync(originalFileBytes);
  const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
  if (mode) applyRevisionsToDoc(d, mode, null);
  return collectParagraphElements(d.documentElement, "all").map((p) => Array.from(p.getElementsByTagNameNS(W_NS, "t")).map((t) => t.textContent).join(""));
}, mode);
const selectWord = (page, needle, word) => page.evaluate(({ needle, word }) => {
  const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => x.textContent.includes(needle));
  const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  for (let n = w.nextNode(); n; n = w.nextNode()) {
    const i = n.data.indexOf(word);
    if (i < 0) continue;
    const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + word.length);
    p.closest(".docx-edit-root").focus({ preventScroll: true });
    getSelection().removeAllRanges(); getSelection().addRange(r);
    return true;
  }
  return false;
}, { needle, word });

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.removeItem("dwb.authorName"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => (d.type() === "prompt" ? d.accept("Jan Test") : d.accept()));
  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await idle(page);
  const original = await bodyTexts(page);
  await page.evaluate(() => { appFrame.setReadOnly(false); setSidebarOpen(true); document.getElementById("panel-review").open = true; });
  await idle(page);

  // ── włączenie ──────────────────────────────────────────────────────────────
  check("domyślnie wyłączone, bez znacznika na pasku stanu", await page.evaluate(() => !document.getElementById("rvTrack").checked && document.getElementById("statusTrack").hidden));
  await page.click("#rvTrack");
  await page.waitForFunction(() => dwbTrack.isOn(), null, { timeout: 8000 });
  check("Śledź zmiany: włączone, imię z pytania, znacznik „Śledzenie zmian” na pasku stanu", await page.evaluate(() => dwbTrack._state.author === "Jan Test" && !document.getElementById("statusTrack").hidden && document.getElementById("rvTrack").checked));

  // ── edycja: zamiana słowa, dopisek, nowy akapit ─────────────────────────────
  await selectWord(page, "Dokument testowy dla", "testowy");
  await page.keyboard.type("próbny");
  await selectWord(page, "§1 Strony umowy", "umowy");
  await page.evaluate(() => getSelection().collapseToEnd());
  await page.keyboard.type(" i aneksu");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Nowy akapit ze śledzeniem");
  await idle(page);
  await page.waitForTimeout(500);
  const marks = await page.evaluate(() => [...document.querySelectorAll(".docx-preview-host p.dwb-tracked")].map((p) => p.textContent.slice(0, 30)));
  check("paski przy zmienionych akapitach (3), reszta bez", marks.length === 3 && marks.some((x) => x.includes("próbny")) && marks.some((x) => x.includes("Nowy akapit")), JSON.stringify(marks));
  const bars = await page.evaluate(() => [...document.querySelectorAll(".dwb-track-bar")].map((b) => { const r = b.getBoundingClientRect(); return { w: Math.round(r.width), h: Math.round(r.height) }; }));
  check("czerwone paski na marginesie przy zmienionych akapitach (też jednowierszowych)", bars.length === 3 && bars.every((b) => b.w >= 1 && b.h >= 10), JSON.stringify(bars));
  check("informacja w Recenzji: liczba zmienionych akapitów", /3/.test(await page.textContent("#rvTrackInfo")), await page.textContent("#rvTrackInfo"));

  // ── zapis na dysk ───────────────────────────────────────────────────────────
  const disk = await xmlOf(page, "(async () => dwbTrack.bytesForDisk(await buildDocumentForSave()))()");
  check("zapis: w:del „testowy” + w:ins „próbny”, autor Jan Test", /<w:del [^>]*w:author="Jan Test"[^>]*>.*?<w:delText[^>]*>testowy<\/w:delText>/.test(disk.doc) && /<w:ins [^>]*w:author="Jan Test"[^>]*>.*?<w:t[^>]*>próbny<\/w:t>/.test(disk.doc), disk.doc.slice(disk.doc.indexOf("testowy") - 300, disk.doc.indexOf("testowy") + 100));
  check("zapis: nowy akapit ze znacznikiem akapitu wstawionym", /<w:rPr><w:ins [^>]*\/>.*?Nowy akapit ze śledzeniem|Nowy akapit ze śledzeniem/.test(disk.doc) && /<w:pPr>(?:(?!<\/w:pPr>).)*<w:rPr><w:ins /.test(disk.doc));
  check("zapis: w:trackRevisions w ustawieniach (Word dalej śledzi)", /<w:trackRevisions\/>/.test(disk.settings));
  const clean = await xmlOf(page, "buildDocumentForSave()");
  check("w aplikacji czysty stan (bez w:ins/w:del) — dalej da się pisać", !/<w:ins |<w:del /.test(clean.doc) && await page.evaluate(() => !collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => x.textContent.includes("próbny"))?.dataset.lock));

  // ── Pokaż poprawki w dokumencie (śledzenie trwa) ────────────────────────────
  await page.click("#rvTrackShowBtn");
  await idle(page);
  check("„Pokaż poprawki”: poprawki w dokumencie, śledzenie dalej włączone", /<w:ins /.test((await xmlOf(page, "Promise.resolve(originalFileBytes)")).doc) && await page.evaluate(() => dwbTrack.isOn()));
  await page.waitForTimeout(300);
  const rv = await page.evaluate(async () => { await ensureLazyFeature("review"); if (typeof renderReviewPanel === "function") await renderReviewPanel(); return document.getElementById("rvChanges").textContent; });
  check("Recenzja pokazuje poprawki Jana Test", /Jan Test/.test(rv), rv.slice(0, 200));
  check("po pokazaniu poprawek: bez pasków (nowy punkt odniesienia)", await page.evaluate(() => !document.querySelector(".docx-preview-host p.dwb-tracked")));

  // dalsza zmiana i wyłączenie skrótem
  await selectWord(page, "Dane Wynajmującego", "Wynajmującego");
  await page.keyboard.type("Właściciela");
  await idle(page);
  await page.keyboard.press(`${MOD}+Shift+KeyE`);
  await page.waitForFunction(() => !dwbTrack.isOn(), null, { timeout: 8000 });
  await idle(page);
  const off = await xmlOf(page, "Promise.resolve(originalFileBytes)");
  check("Ctrl/⌘+Shift+E wyłącza: obie tury poprawek w dokumencie, bez w:trackRevisions", /<w:delText[^>]*>Wynajmującego<\/w:delText>/.test(off.doc) && /<w:delText[^>]*>testowy<\/w:delText>/.test(off.doc) && !/trackRevisions/.test(off.settings));
  check("wyłączone: znacznik na pasku stanu znika", await page.evaluate(() => document.getElementById("statusTrack").hidden && !document.getElementById("rvTrack").checked));

  // ── Odrzuć wszystkie = oryginał; Cofnij ─────────────────────────────────────
  check("Odrzuć wszystkie = dokładnie tekst oryginału", JSON.stringify(await bodyTexts(page, "reject")) === JSON.stringify(original), JSON.stringify((await bodyTexts(page, "reject")).slice(0, 6)));
  const accepted = await bodyTexts(page, "accept");
  check("Akceptuj wszystkie = tekst po zmianach", accepted.some((x) => x.includes("Dokument próbny")) && accepted.includes("Nowy akapit ze śledzeniem") && accepted.some((x) => x.includes("Dane Właściciela")), JSON.stringify(accepted.slice(0, 8)));

  // ── plik z Worda ze śledzeniem — włącza się sam ─────────────────────────────
  await page.evaluate(async () => {
    const z = await JSZip.loadAsync(originalFileBytes);
    await rlSetTrackRevisions(z, true); // jak plik zapisany przez Worda ze śledzeniem
    const bytes = await z.generateAsync({ type: "uint8array" });
    await ingestFile(new File([bytes], "ze-sledzeniem.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }), { silent: true });
  });
  await idle(page);
  check("plik z w:trackRevisions: śledzenie włącza się samo", await page.evaluate(() => dwbTrack.isOn() && document.getElementById("rvTrack").checked));

  // ── karty: stan śledzenia jedzie z kartą ────────────────────────────────────
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);
  await selectWord(page, "Najemca zobowiązuje", "lokalu"); // zwykły akapit (z poprawkami są tylko do odczytu)
  await page.keyboard.type("mieszkania");
  await idle(page);
  await page.evaluate(async () => {
    const b = await (await fetch("docs/samples/sample.docx")).arrayBuffer();
    await openDocumentFiles([{ file: new File([b], "druga-karta.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }) }]);
  });
  await idle(page);
  check("druga karta: śledzenie wyłączone", await page.evaluate(() => !dwbTrack.isOn()));
  await page.evaluate(async () => { const first = dwbOpenDocs.list().find((d) => d.name === "ze-sledzeniem.docx"); await dwbOpenDocs.switchTo(first.id); });
  await idle(page);
  await page.waitForTimeout(500);
  check("powrót do pierwszej karty: śledzenie dalej włączone, zmiana dalej zaznaczona", await page.evaluate(() => dwbTrack.isOn() && !!document.querySelector(".docx-preview-host p.dwb-tracked")), JSON.stringify(await page.evaluate(() => ({ on: dwbTrack.isOn(), name: currentFileName, marks: document.querySelectorAll(".docx-preview-host p.dwb-tracked").length, tabs: dwbOpenDocs.list(), txt: collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).slice(0, 3).map((p) => p.textContent.slice(0, 30)) }))));

  // ── Porównaj wersje → dokument z poprawkami ─────────────────────────────────
  await page.evaluate(async () => {
    document.getElementById("panel-compare").open = true;
    await ensureLazyFeature("compare");
    const b = await (await fetch("docs/samples/headings-sample.docx")).arrayBuffer();
    const dt = new DataTransfer();
    dt.items.add(new File([b], "stara.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
    document.getElementById("compareFile").files = dt.files;
  });
  await page.click("#compareRunBtn");
  await page.waitForFunction(() => !document.getElementById("compareRedlineBtn").hidden, null, { timeout: 15000 });
  const tabsBefore = await page.evaluate(() => dwbOpenDocs.list().length);
  await page.click("#compareRedlineBtn");
  await page.waitForFunction((n) => dwbOpenDocs.list().length === n + 1, tabsBefore, { timeout: 15000 });
  await idle(page);
  const red = await xmlOf(page, "Promise.resolve(originalFileBytes)");
  check("„Utwórz dokument z poprawkami”: nowa karta z w:ins/w:del, porównywane pliki bez zmian", /<w:ins /.test(red.doc) && /<w:del /.test(red.doc) && await page.evaluate(() => /poprawki/.test(currentFileName)), await page.evaluate(() => currentFileName));

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ śledzenie zmian [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
