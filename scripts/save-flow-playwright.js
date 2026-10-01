// save-flow-playwright.js — co robi „Zapisz” w zależności od tego, jak plik otwarto.
//
// Zgłoszenie 2026-10-01 (Windows, Chrome/Edge): plik otwarty bez uchwytu (przeciągnięty,
// „Przykład”, zwykłe okno wyboru) → „Zapisz” otwierało okno OTWIERANIA plików. Teraz:
//   - bez uchwytu / bez zgody na zapis → okno ZAPISU („Zapisz jako”), nigdy otwierania,
//   - z uchwytem → zapis w miejscu, pytanie „Zapisać w oryginale?” tylko raz na plik,
//   - po „Zapisz jako” kolejne „Zapisz” idzie do wybranego pliku bez pytań,
//   - Safari/iPhone (brak File System Access) → pobranie kopii.
// Okna wyboru plików są tu podstawione (Playwright ich nie obsłuży) — liczymy, które wołano.
// ENGINE=webkit sprawdza ścieżkę Safari (pobranie).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(250));

// Podstawione okna wyboru + udawany uchwyt pliku (zapis trafia do window.__fs.files[nazwa]).
const installFakeFs = () => {
  window.__fs = { calls: [], files: {}, confirms: 0, perm: "granted" };
  const handle = (name) => ({
    kind: "file", name,
    queryPermission: async () => window.__fs.perm,
    requestPermission: async () => { window.__fs.calls.push(`perm:${name}`); return window.__fs.perm; },
    getFile: async () => new File([window.__fs.files[name] || new Uint8Array()], name),
    createWritable: async () => {
      const chunks = [];
      return { write: async (b) => chunks.push(b), close: async () => { window.__fs.files[name] = chunks[0]; window.__fs.calls.push(`write:${name}`); } };
    },
  });
  window.__fsHandle = handle;
  window.showOpenFilePicker = async () => { window.__fs.calls.push("OPEN-PICKER"); throw new DOMException("x", "AbortError"); };
  window.showSaveFilePicker = async (opts) => { window.__fs.calls.push(`save-picker:${opts?.suggestedName}`); return handle("wybrany.docx"); };
  const origConfirm = window.confirm;
  window.confirm = (msg) => { if (/oryginalnym pliku/.test(msg)) window.__fs.confirms++; return true; };
  void origConfirm;
};

async function edit(page, text) {
  await page.evaluate((txt) => applyDocumentEdit({ op: "affix", prefix: "", suffix: txt, scope: "all" }), text);
  await idle(page);
}
async function clickSave(page) {
  await page.click("#heroSaveBtn");
  await page.waitForTimeout(400);
  await idle(page);
}
const fsState = (page) => page.evaluate(() => ({ calls: window.__fs.calls.join(" "), confirms: window.__fs.confirms, files: Object.keys(window.__fs.files), dirty: hasUnsavedChanges }));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", acceptDownloads: true, viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept(d.defaultValue()));
  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await idle(page);

  if (ENGINE === "webkit") {
    // Safari / iPhone: brak File System Access → „Zapisz” = pobranie kopii z podaną nazwą
    await edit(page, " X");
    const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 8000 }).catch(() => null), page.click("#heroSaveBtn")]);
    check("Safari: „Zapisz” pobiera kopię (bez okna otwierania)", !!dl && /_edited\.docx$/.test(dl.suggestedFilename()), dl?.suggestedFilename());
    await page.waitForTimeout(300);
    check("Safari: po pobraniu „Zapisz” gaśnie (zmiany zapisane)", !(await page.evaluate(() => hasUnsavedChanges)));
  } else {
    await page.evaluate(installFakeFs);

    // 1) plik bez uchwytu (Przykład / przeciągnięty / zwykłe okno) → okno ZAPISU
    await edit(page, " A");
    await clickSave(page);
    let s = await fsState(page);
    check("bez uchwytu: „Zapisz” otwiera okno ZAPISU, nie otwierania", /save-picker:headings-sample_edited\.docx/.test(s.calls) && !/OPEN-PICKER/.test(s.calls), s.calls);
    check("bez uchwytu: plik zapisany, zmiany zgaszone", s.files.includes("wybrany.docx") && !s.dirty, JSON.stringify(s));
    // 2) kolejne „Zapisz” → do wybranego pliku, bez okna i bez pytania
    await page.evaluate(() => { window.__fs.calls = []; });
    await edit(page, " B");
    await clickSave(page);
    s = await fsState(page);
    check("po „Zapisz jako”: następne „Zapisz” prosto do tego pliku, bez okna i pytań", s.calls === "write:wybrany.docx" && s.confirms === 0, JSON.stringify(s));
    const saved = await page.evaluate(async () => (await extractParagraphTextsFromDocx(new Uint8Array(await window.__fs.files["wybrany.docx"].arrayBuffer?.() ?? window.__fs.files["wybrany.docx"]))).some((p) => p.endsWith(" A B")));
    check("zapisany plik ma obie zmiany", saved);

    // 3) plik otwarty z uchwytem (Otwórz / przeciągnięty w Chrome) → zapis w miejscu, pytanie raz
    await page.evaluate(async () => {
      window.__fs.calls = [];
      const bytes = originalFileBytes;
      window.__fs.files["umowa.docx"] = bytes;
      await ingestFile(new File([bytes], "umowa.docx"), { handle: window.__fsHandle("umowa.docx") });
    });
    await idle(page);
    await edit(page, " C");
    await clickSave(page);
    await edit(page, " D");
    await clickSave(page);
    s = await fsState(page);
    check("z uchwytem: zapis w oryginale, pytanie tylko RAZ", s.calls === "write:umowa.docx write:umowa.docx" && s.confirms === 1, JSON.stringify(s));

    // 4) brak zgody na zapis do oryginału → okno ZAPISU (kopia), dalej bez okna otwierania
    await page.evaluate(async () => {
      window.__fs.calls = [];
      window.__fs.perm = "denied";
      await ingestFile(new File([originalFileBytes], "zablokowany.docx"), { handle: window.__fsHandle("zablokowany.docx") });
    });
    await idle(page);
    await edit(page, " E");
    await clickSave(page);
    s = await fsState(page);
    check("bez zgody: okno ZAPISU zamiast błędu / okna otwierania", /save-picker:zablokowany_edited\.docx/.test(s.calls) && !/OPEN-PICKER/.test(s.calls), s.calls);

    // 5) przeciągnięty plik: uchwyt z dataTransfer → późniejsze „Zapisz” w miejscu
    const dropped = await page.evaluate(async () => {
      window.__fs.perm = "granted";
      const file = new File([originalFileBytes], "przeciagniety.docx");
      const fake = { dataTransfer: { items: [{ kind: "file", getAsFileSystemHandle: async () => window.__fsHandle("przeciagniety.docx") }] } };
      await ingestDroppedFile(file, droppedFileHandle(fake));
      return { name: fileHandle?.name };
    });
    check("przeciągnięty plik (Chrome/Edge) dostaje uchwyt → „Zapisz” w miejscu", dropped.name === "przeciagniety.docx", JSON.stringify(dropped));
  }

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ zapis [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
