// undo-playwright.js — Cofnij / Ponów (paczka D, app/undo.js).
//
//   node scripts/undo-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/undo-playwright.js
//
// Musi być prawdą:
//   - operacja z panelu (zamiana) cofa się i ponawia; przycisk mówi, CO cofnie,
//   - ciągłe pisanie (z Enterem) = jeden krok; przerwa ≥ 1,5 s = nowy krok,
//   - po cofnięciu zapisany plik = podgląd (akapity i treść),
//   - „Zapisz”: licznik zmian maleje przy Cofnij, powrót do stanu zapisu gasi przycisk,
//   - nowa zmiana po Cofnij czyści Ponów,
//   - Ctrl+Z w polu szukania cofa PISANIE w polu, nie dokument,
//   - B (pogrubienie) to osobny krok,
//   - nowy dokument = pusta historia.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = "Control"; // Playwright: Control działa w obu silnikach (skrót łapie ctrlKey || metaKey)
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

const paraText = (page, needle) => page.evaluate((n) => [...document.querySelectorAll(".docx-preview-host p")].map((p) => p.textContent).find((t) => t.includes(n)) || null, needle);
const saveState = (page) => page.evaluate(() => ({ dirty: document.getElementById("heroSaveBtn").classList.contains("is-dirty"), count: document.getElementById("heroSaveCount").hidden ? "" : document.getElementById("heroSaveCount").textContent }));
const viewEqualsFile = (page) => page.evaluate(async () => {
  const bytes = await buildDocumentForSave();
  const dom = [...document.querySelectorAll(".docx-preview-host p")].map((x) => x.textContent);
  const xml = (await extractParagraphTextsFromDocx(bytes)).map(String);
  return JSON.stringify(dom) === JSON.stringify(xml) ? "ok" : `DOM ${dom.length} / plik ${xml.length}`;
});
const idle = (page) => page.waitForFunction(() => !document.getElementById("loadingOverlay") || document.getElementById("loadingOverlay").classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(250));

async function placeCaretEnd(page, needle) {
  await page.locator(".docx-preview-host p", { hasText: needle }).first().click();
  await page.evaluate((n) => {
    const el = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.includes(n));
    const r = document.createRange(); r.selectNodeContents(el); r.collapse(false);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  }, needle);
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await idle(page);

  check("start: Cofnij i Ponów nieaktywne", await page.evaluate(() => document.getElementById("undoBtn").disabled && document.getElementById("redoBtn").disabled));

  // 1) operacja z panelu
  await page.evaluate(() => applyDocumentEdit({ op: "replace", find: "Kaucja", replace: "Depozyt", regex: false, scope: "all" }));
  await idle(page);
  const afterReplace = await paraText(page, "§4");
  const label = await page.getAttribute("#undoBtn", "aria-label");
  check("zamiana z panelu zrobiona, „Cofnij: zamiana tekstu”", afterReplace === "§4 Depozyt" && /Cofnij: zamiana tekstu/.test(label), `${afterReplace} | ${label}`);
  check("licznik „Zapisz” = 1", (await saveState(page)).count === "1");
  await page.click("#undoBtn");
  await idle(page);
  check("Cofnij przywraca tekst", (await paraText(page, "§4")) === "§4 Kaucja");
  const clean = await saveState(page);
  check("powrót do stanu z wczytania gasi „Zapisz”", !clean.dirty, JSON.stringify(clean));
  await page.click("#redoBtn");
  await idle(page);
  check("Ponów przywraca zamianę", (await paraText(page, "§4")) === "§4 Depozyt");

  // 2) pisanie z Enterem = jeden krok
  await page.click('.mode-btn[data-mode="edit"]');
  await placeCaretEnd(page, "Dane Najemcy");
  await page.keyboard.type(" X", { delay: 20 });
  await page.keyboard.press("Enter");
  await page.keyboard.type("Nowy akapit", { delay: 20 });
  await page.waitForTimeout(300);
  check("pisanie + Enter: jest nowy akapit", !!(await paraText(page, "Nowy akapit")));
  check("licznik „Zapisz” = 2 (zamiana + pisanie)", (await saveState(page)).count === "2", JSON.stringify(await saveState(page)));
  await page.keyboard.press(`${MOD}+z`);
  await idle(page);
  const t1 = { nowy: await paraText(page, "Nowy akapit"), dane: await paraText(page, "Dane Najemcy") };
  check("Ctrl+Z cofa całe pisanie (z Enterem) jednym krokiem", !t1.nowy && t1.dane === "Dane Najemcy", JSON.stringify(t1));
  check("po cofnięciu pisania: zapis = podgląd", (await viewEqualsFile(page)) === "ok", await viewEqualsFile(page));
  await page.keyboard.press(`${MOD}+Shift+z`);
  await idle(page);
  check("Ctrl+Shift+Z ponawia pisanie", !!(await paraText(page, "Nowy akapit")) && (await paraText(page, "Dane Najemcy X")) === "Dane Najemcy X");

  // 3) przerwa = nowy krok
  await placeCaretEnd(page, "Nowy akapit");
  await page.keyboard.type(" raz", { delay: 20 });
  await page.waitForTimeout(1700);
  await page.keyboard.type(" dwa", { delay: 20 });
  await page.waitForTimeout(200);
  await page.keyboard.press(`${MOD}+z`);
  await idle(page);
  check("przerwa ≥ 1,5 s dzieli pisanie na kroki (cofa tylko „ dwa”)", (await paraText(page, "Nowy akapit")) === "Nowy akapit raz", await paraText(page, "Nowy akapit"));

  // 4) nowa zmiana czyści Ponów
  await placeCaretEnd(page, "Nowy akapit");
  await page.keyboard.type("!", { delay: 20 });
  await page.waitForTimeout(200);
  check("nowa zmiana po Cofnij czyści Ponów", await page.evaluate(() => document.getElementById("redoBtn").disabled));

  // 5) pogrubienie = osobny krok
  await page.waitForTimeout(1600);
  await page.evaluate(() => {
    const el = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.startsWith("Dokument testowy"));
    el.focus();
    const r = document.createRange(); r.selectNodeContents(el);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  });
  await page.click("#fmtBold");
  await page.waitForTimeout(200);
  const boldLabel = await page.getAttribute("#undoBtn", "aria-label");
  const bolded = await page.evaluate(() => !![...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.startsWith("Dokument testowy"))?.querySelector("b, strong, [style*='bold'], [style*='700']"));
  await page.click("#undoBtn");
  await idle(page);
  const unbolded = await page.evaluate(() => ![...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.startsWith("Dokument testowy"))?.querySelector("b, strong, [style*='bold'], [style*='700']"));
  check("B = osobny krok „formatowanie”, Cofnij zdejmuje pogrubienie", /formatowanie/.test(boldLabel) && bolded && unbolded, `${boldLabel} ${bolded} ${unbolded}`);

  // 5b) zapis = nowy punkt odniesienia: licznik 0; Cofnij zmiany sprzed zapisu = znów „zmieniony”
  await page.evaluate(() => setDirtyState(false)); // to samo, co robi udany zapis
  await page.waitForTimeout(100);
  const saved = await saveState(page);
  await page.click("#undoBtn");
  await idle(page);
  const afterUndoSaved = await saveState(page);
  await page.click("#redoBtn");
  await idle(page);
  const backToSaved = await saveState(page);
  check("po zapisie: licznik 0 → Cofnij = 1 zmiana → Ponów = znów zapisane", !saved.dirty && afterUndoSaved.dirty && afterUndoSaved.count === "1" && !backToSaved.dirty, JSON.stringify({ saved, afterUndoSaved, backToSaved }));

  // 5c) szybka ścieżka (paczka F): samo pisanie bez Entera cofa się BEZ przerysowania
  await page.waitForTimeout(1600);
  await placeCaretEnd(page, "Dane Wynajmującego");
  await page.keyboard.type(" AAA", { delay: 15 });
  await placeCaretEnd(page, "§5 Obowiązki stron");
  await page.keyboard.type(" BBB", { delay: 15 });
  await page.waitForTimeout(200);
  const fast = await page.evaluate(async () => {
    let rerendered = false;
    const orig = window.renderDocxPreview;
    window.renderDocxPreview = function (...a) { rerendered = true; return orig.apply(this, a); };
    const txt = (n) => [...document.querySelectorAll(".docx-preview-host p")].map((p) => p.textContent).find((x) => x.startsWith(n));
    const t0 = performance.now();
    await dwbUndo.undo(); // klik w drugi akapit zamknął krok → najpierw znika tylko „ BBB”
    const afterFirst = { a: txt("Dane Wynajmującego"), b: txt("§5 Obowiązki") };
    await dwbUndo.undo();
    const ms = performance.now() - t0;
    window.renderDocxPreview = orig;
    return { rerendered, ms: Math.round(ms), afterFirst, a: txt("Dane Wynajmującego"), b: txt("§5 Obowiązki") };
  });
  check("cofnięcie samego pisania bez przerysowania (2 kroki: „ BBB”, potem „ AAA”)", !fast.rerendered && fast.afterFirst.a === "Dane Wynajmującego AAA" && fast.afterFirst.b === "§5 Obowiązki stron" && fast.a === "Dane Wynajmującego" && fast.b === "§5 Obowiązki stron", JSON.stringify(fast));
  check("szybkie cofnięcie: zapis = podgląd", (await viewEqualsFile(page)) === "ok", await viewEqualsFile(page));
  await page.evaluate(() => dwbUndo.redo());
  await idle(page);
  check("ponowienie pisania: tekst wraca, zapis = podgląd", (await paraText(page, "Dane Wynajmującego")) === "Dane Wynajmującego AAA" && (await viewEqualsFile(page)) === "ok");

  // 5d) wstawienie {{pola}} z panelu (inna ścieżka zmiany DOM) — trafia do zapisu
  await page.evaluate(() => ensureLazyFeature("placeholders"));
  await placeCaretEnd(page, "§6 Wypowiedzenie");
  await page.evaluate(() => { document.getElementById("phInsertName").value = "termin"; insertPlaceholderAtCaret(); });
  await page.waitForTimeout(200);
  const phSaved = await page.evaluate(async () => (await extractParagraphTextsFromDocx(await buildDocumentForSave())).some((t) => /§6 Wypowiedzenie.*\{\{termin\}\}/.test(t)));
  check("wstawione {{pole}} trafia do zapisanego pliku", phSaved && (await viewEqualsFile(page)) === "ok");

  // 6) Ctrl+Z w polu szukania nie rusza dokumentu
  const before = await page.evaluate(() => dwbUndo._debug().undo.length);
  await page.fill("#searchQuery", "");
  await page.click("#searchQuery");
  await page.keyboard.type("abc");
  await page.keyboard.press(`${MOD}+z`);
  await page.waitForTimeout(300);
  const after = await page.evaluate(() => ({ steps: dwbUndo._debug().undo.length, loading: !document.getElementById("loadingOverlay").classList.contains("hidden") }));
  check("Ctrl+Z w polu szukania nie cofa dokumentu", after.steps === before && !after.loading, JSON.stringify({ before, after }));

  // 7) nowy dokument = pusta historia
  await page.evaluate(() => { hasUnsavedChanges = false; return loadSampleDocument("headings-sample"); });
  await idle(page);
  check("nowy dokument: historia pusta", await page.evaluate(() => document.getElementById("undoBtn").disabled && document.getElementById("redoBtn").disabled && dwbUndo._debug().undo.length === 0));

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ Cofnij/Ponów [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
