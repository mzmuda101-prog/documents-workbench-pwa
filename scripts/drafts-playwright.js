// drafts-playwright.js — szkic niezapisanej pracy i odzyskiwanie (app/drafts.js).
//
// Symulacja iOS: edycja → aplikacja w tle (visibilitychange) → karta ZABITA bez żadnego
// zdarzenia zamknięcia → nowe otwarcie: karta „Niezapisana praca” → „Przywróć” odtwarza
// wszystko (pisanie + operacja z panelu). Plus: „Odrzuć”, zapis kasuje szkic, cofnięcie do
// stanu z pliku kasuje szkic, świadome zamknięcie dokumentu, żyjące drugie okno (jego szkic
// nie jest „zgubioną pracą”), szkic starszy niż 7 dni znika. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => sleep(350));

async function openPage(context, url = APP_URL) {
  const page = await context.newPage();
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept(d.defaultValue()));
  await page.goto(url, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await sleep(300);
  return page;
}
const errors = [];

async function typeAtEnd(page, text) {
  await page.evaluate(() => {
    if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); }
  });
  // akapity dostają contentEditable chwilę po narysowaniu (oznaczanie blokad) — poczekaj
  await page.waitForFunction(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).some((p) => p.isContentEditable), null, { timeout: 10000 });
  await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.isContentEditable && p.textContent.trim().length > 3);
    placeCaret(ps[ps.length - 1], ps[ps.length - 1].textContent.length);
  });
  await page.keyboard.type(text, { delay: 5 });
}
// „system zabija aplikację w tle”: tło (jedyne zdarzenie, jakie dostaje) → zamknięcie bez unload
async function backgroundAndKill(page) {
  await page.evaluate(async () => {
    Object.defineProperty(document, "visibilityState", { configurable: true, get: () => "hidden" });
    document.dispatchEvent(new Event("visibilitychange"));
    await dwbDrafts._flush();
  });
  await page.close({ runBeforeUnload: false });
}
const cards = (page) => page.evaluate(async () => {
  await new Promise((r) => setTimeout(r, 600)); // odpowiedzi żyjących okien + IndexedDB
  return [...document.querySelectorAll("#draftRecovery .draft-card")].map((c) => ({ name: c.querySelector(".draft-name")?.textContent, meta: c.querySelector(".draft-meta")?.textContent }));
});
const dbCount = (page) => page.evaluate(() => new Promise((resolve) => {
  const r = indexedDB.open("dwb-drafts", 1);
  r.onsuccess = () => { const t = r.result.transaction("drafts", "readonly").objectStore("drafts").count(); t.onsuccess = () => { resolve(t.result); r.result.close(); }; };
  r.onerror = () => resolve(-1);
}));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));

  // ── 1. edycja → tło → zabicie → odzyskanie ─────────────────────────────────
  let page = await openPage(context, `${APP_URL}?sample=headings-sample`);
  await idle(page);
  await typeAtEnd(page, " SZKIC-A");
  await page.evaluate(() => applyDocumentEdit({ op: "replace", find: "Kaucja", replace: "Depozyt", regex: false, scope: "all" }));
  await idle(page);
  await typeAtEnd(page, " SZKIC-B");
  await backgroundAndKill(page);

  page = await openPage(context);
  let c = await cards(page);
  check("po zabiciu aplikacji: karta „Niezapisana praca” z nazwą pliku i liczbą zmian", c.length === 1 && c[0].name === "headings-sample.docx" && /dziś o \d\d:\d\d/.test(c[0].meta) && /zmiany: [1-9]/.test(c[0].meta), JSON.stringify(c));
  check("karta mówi, że szkic jest tylko na urządzeniu i znika po 7 dniach", /7 dniach/.test(await page.textContent("#draftRecovery")));
  await page.click("#draftRecovery .btn.primary");
  await idle(page);
  let st = await page.evaluate(async () => {
    const view = document.querySelector(".docx-preview-host")?.textContent || "";
    const file = (await extractParagraphTextsFromDocx(await buildDocumentForSave())).join("\n");
    return { name: currentFileName, dirty: hasUnsavedChanges, view: ["SZKIC-A", "SZKIC-B", "Depozyt"].every((m) => view.includes(m)) && !view.includes("Kaucja"), file: ["SZKIC-A", "SZKIC-B", "Depozyt"].every((m) => file.includes(m)), card: !document.getElementById("draftRecovery").hidden };
  });
  check("„Przywróć”: pisanie i zamiana z panelu wracają (ekran + zapis), „Zapisz” świeci", st.name === "headings-sample.docx" && st.dirty && st.view && st.file && !st.card, JSON.stringify(st));

  // dalsza praca po odzyskaniu nadpisuje TEN SAM szkic (nie mnoży wpisów)
  await typeAtEnd(page, " SZKIC-C");
  await backgroundAndKill(page);
  page = await openPage(context);
  c = await cards(page);
  check("praca po odzyskaniu → nadal jeden szkic, z nowymi zmianami", c.length === 1 && (await dbCount(page)) === 1, JSON.stringify(c));
  await page.click("#draftRecovery .btn.primary");
  await idle(page);
  check("drugie odzyskanie ma też zmiany zrobione po pierwszym", await page.evaluate(() => document.querySelector(".docx-preview-host").textContent.includes("SZKIC-C")));

  // ── 2. zapis kasuje szkic ──────────────────────────────────────────────────
  await page.evaluate(async () => {
    if (window.showSaveFilePicker) {
      window.showSaveFilePicker = async () => ({ kind: "file", name: "kopia.docx", createWritable: async () => ({ write: async () => {}, close: async () => {} }), queryPermission: async () => "granted", requestPermission: async () => "granted" });
    }
    await saveDocument();
  });
  await sleep(600);
  check("po zapisie: brak niezapisanych zmian i brak szkicu", !(await page.evaluate(() => hasUnsavedChanges)) && (await dbCount(page)) === 0, String(await dbCount(page)));

  // ── 3. „Odrzuć” ────────────────────────────────────────────────────────────
  await typeAtEnd(page, " DO-ODRZUCENIA");
  await backgroundAndKill(page);
  page = await openPage(context);
  c = await cards(page);
  await page.click("#draftRecovery .btn:not(.primary)");
  await sleep(500);
  check("„Odrzuć” (po potwierdzeniu) usuwa szkic i kartę", c.length === 1 && (await dbCount(page)) === 0 && (await page.evaluate(() => document.getElementById("draftRecovery").hidden)), JSON.stringify(c));

  // ── 4. cofnięcie do stanu z pliku i świadome zamknięcie kasują szkic ───────
  await page.evaluate(() => loadSampleDocument("headings-sample"));
  await idle(page);
  await typeAtEnd(page, " COFNIJ");
  await page.evaluate(async () => { document.dispatchEvent(new Event("visibilitychange")); await new Promise((r) => setTimeout(r, 1800)); await dwbDrafts._flush(); });
  const before = await dbCount(page);
  await page.click("#undoBtn");
  await idle(page);
  await sleep(300);
  check("Cofnij do stanu z pliku usuwa szkic (nie ma czego odzyskiwać)", before === 1 && (await dbCount(page)) === 0 && !(await page.evaluate(() => hasUnsavedChanges)), `${before} → ${await dbCount(page)}`);
  await typeAtEnd(page, " ZAMKNIJ");
  await sleep(1800);
  await page.evaluate(() => requestCloseDocument()); // pyta o porzucenie zmian → OK
  await sleep(500);
  check("świadome zamknięcie dokumentu (zgoda na porzucenie) usuwa szkic", (await dbCount(page)) === 0);

  // ── 5. drugie, ŻYJĄCE okno nie widzi szkicu pierwszego jako zgubionej pracy ─
  await page.evaluate(() => loadSampleDocument("headings-sample"));
  await idle(page);
  await typeAtEnd(page, " ZYJE");
  await sleep(1800);
  const other = await openPage(context);
  c = await cards(other);
  check("drugie okno: szkic żyjącego okna nie jest pokazywany do odzyskania", c.length === 0 && (await dbCount(other)) === 1, JSON.stringify(c));
  await other.close();

  // ── 6. starszy niż 7 dni znika ─────────────────────────────────────────────
  await backgroundAndKill(page);
  page = await openPage(context);
  await page.evaluate(() => new Promise((resolve) => {
    const r = indexedDB.open("dwb-drafts", 1);
    r.onsuccess = () => {
      const t = r.result.transaction("drafts", "readwrite");
      const s = t.objectStore("drafts");
      s.getAll().onsuccess = (e) => { e.target.result.forEach((d) => s.put({ ...d, savedAt: Date.now() - 8 * 86400000 })); };
      t.oncomplete = () => { r.result.close(); resolve(); };
    };
  }));
  await page.reload({ waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  c = await cards(page);
  check("szkic starszy niż 7 dni: brak karty i sam się usuwa", c.length === 0 && (await dbCount(page)) === 0, JSON.stringify(c));

  // ── 7. bez IndexedDB (np. tryb prywatny) aplikacja działa normalnie ───────
  const ctx2 = await browser.newContext({ serviceWorkers: "block" });
  await ctx2.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); Object.defineProperty(window, "indexedDB", { value: undefined }); });
  const p2 = await ctx2.newPage();
  const err2 = [];
  p2.on("pageerror", (e) => err2.push(e.message));
  await p2.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await sleep(2500);
  await typeAtEnd(p2, " BEZ-IDB");
  await p2.evaluate(async () => { document.dispatchEvent(new Event("visibilitychange")); await dwbDrafts._flush(); });
  check("bez IndexedDB: edycja działa, brak błędów", (await p2.evaluate(() => document.querySelector(".docx-preview-host").textContent.includes("BEZ-IDB"))) && !err2.length, err2.join(" | "));
  await ctx2.close();

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ szkic i odzyskiwanie [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
