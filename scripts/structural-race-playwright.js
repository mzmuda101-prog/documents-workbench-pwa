// structural-race-playwright.js — Enter/Backspace w trakcie scalania, zapisu i po nieudanej operacji.
//
//   node scripts/structural-race-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/structural-race-playwright.js
//
// Kolejka Enter/Backspace (docx-inline-edit.js) przebudowuje plik w tle. Scalanie wpisanego tekstu
// (mergeInlineEditsIntoBytes) i zapis budowały plik OBOK tej kolejki: Enter wciśnięty w trakcie
// budował na starym pliku i wygrywał ten, kto skończył później (ginął podział albo wpisany tekst),
// a baza akapitów rozjeżdżała się z podglądem — zapis wpisywał treść akapitu w sąsiedni.
// Nieudana operacja (wyjątek przy budowie) zostawiała podgląd z innym podziałem niż plik.
// Musi być prawdą: zapis = podgląd, wpisany tekst jest w pliku, także w tych sytuacjach.
// Scalanie/zapis spowalniamy sztucznie (400 ms), żeby okno wyścigu było pewne.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function openDoc(browser) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.waitForTimeout(500);
  await page.click('.mode-btn[data-mode="edit"]');
  await page.waitForTimeout(300);
  await page.evaluate(() => {
    const orig = buildPatchedDocx;
    window.__slow = 0;
    window.__failSplit = 0;
    window.__failMerge = 0;
    window.buildPatchedDocx = async (bytes, edits, ...rest) => {
      if (window.__failSplit && edits?.some((e) => e.op === "splitParagraph")) {
        window.__failSplit--;
        throw new Error("test: nieudany podział");
      }
      if (window.__failMerge && edits?.some((e) => e.op === "mergeParagraph")) {
        window.__failMerge--;
        throw new Error("test: nieudane scalenie");
      }
      // spowalniamy tylko scalanie/zapis (paragraphBatch): Enter z kolejki kończy PRZED nimi —
      // najgorsza kolejność (dawniej scalanie nadpisywało wtedy plik i bazę bez podziału)
      if (window.__slow && edits?.some((e) => e.op === "paragraphBatch")) await new Promise((r) => setTimeout(r, window.__slow));
      return orig(bytes, edits, ...rest);
    };
    // dwa edytowalne akapity z tekstem, daleko od siebie
    const ps = collectPreviewParagraphElements(docCanvasEl.querySelector(".docx-preview-host"));
    const ok = ps.map((p, i) => [p, i]).filter(([p]) => !p.dataset.lock && p.classList.contains("docx-editable-p") && p.textContent.trim().length > 8);
    window.__pa = ok[1][1];
    window.__pb = ok[Math.min(ok.length - 1, 5)][1];
  });
  return { context, page, errors };
}

// kursor na końcu akapitu o danym indeksie (End w długim akapicie idzie tylko na koniec wiersza)
async function caretAtEnd(page, which) {
  await page.evaluate((w) => {
    const ps = collectPreviewParagraphElements(docCanvasEl.querySelector(".docx-preview-host"));
    const el = ps[window[w]];
    el.closest("[contenteditable=true]")?.focus();
    const r = document.createRange();
    r.selectNodeContents(el);
    r.collapse(false);
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  }, which);
}

async function settle(page) {
  await page.waitForFunction(() => !inlineStructuralPending && !inlineLocksPending && document.getElementById("loadingOverlay")?.classList.contains("hidden") !== false, null, { timeout: 20000 });
  await page.waitForTimeout(400);
  await page.waitForFunction(() => !inlineStructuralPending && !inlineLocksPending, null, { timeout: 20000 });
}

async function compare(page) {
  return page.evaluate(async () => {
    window.__slow = 0;
    const bytes = await buildDocumentForSave();
    const dom = collectPreviewParagraphElements(docCanvasEl.querySelector(".docx-preview-host")).map((x) => x.textContent);
    const xml = (await extractParagraphTextsFromDocx(bytes)).map(String);
    const diff = dom.findIndex((t, i) => t !== xml[i]);
    return { same: JSON.stringify(dom) === JSON.stringify(xml), nDom: dom.length, nXml: xml.length, diff, dom: dom[diff], xml: xml[diff], all: dom.join("\n") };
  });
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  try {
    // A) Enter w trakcie scalania wpisanego tekstu do pliku
    {
      const { context, page, errors } = await openDoc(browser);
      await caretAtEnd(page, "__pa");
      await page.keyboard.type(" WPISANE-A");
      await caretAtEnd(page, "__pb");
      await page.evaluate(() => { window.__slow = 400; window.__job = mergeInlineEditsIntoBytes(); });
      await page.keyboard.press("Enter");
      await page.keyboard.type("PO-ENTERZE-A");
      await page.evaluate(() => window.__job);
      await settle(page);
      const r = await compare(page);
      check("A: Enter w trakcie scalania — zapis = podgląd", r.same, `DOM ${r.nDom} / plik ${r.nXml}, różnica w ${r.diff}: ${JSON.stringify(r.dom)} ≠ ${JSON.stringify(r.xml)}`);
      check("A: wpisany tekst i nowy akapit w pliku", /WPISANE-A/.test(r.all) && /\nPO-ENTERZE-A/.test(r.all), "");
      if (errors.length) check("A: brak błędów strony", false, errors.slice(0, 2).join(" | "));
      await context.close();
    }
    // B) Enter i pisanie w trakcie zapisu (Zapisz w miejscu, uchwyt pliku podstawiony)
    {
      const { context, page, errors } = await openDoc(browser);
      await page.evaluate(() => {
        fileHandle = {
          name: "test.docx",
          queryPermission: async () => "granted",
          requestPermission: async () => "granted",
          createWritable: async () => ({ write: async (b) => { window.__written = b; }, close: async () => {} }),
        };
        overwriteConfirmedFor = fileHandle;
      });
      await caretAtEnd(page, "__pa");
      await page.keyboard.type(" WPISANE-B");
      await caretAtEnd(page, "__pb");
      await page.evaluate(() => { window.__slow = 400; window.__job = saveDocument(); });
      await page.keyboard.press("Enter");
      await page.keyboard.type("PO-ENTERZE-B");
      await page.evaluate(() => window.__job);
      await settle(page);
      const dirty = await page.evaluate(() => hasUnsavedChanges);
      const r = await compare(page);
      check("B: Enter w trakcie zapisu — zapis = podgląd", r.same, `DOM ${r.nDom} / plik ${r.nXml}, różnica w ${r.diff}: ${JSON.stringify(r.dom)} ≠ ${JSON.stringify(r.xml)}`);
      check("B: zmiany po zbudowaniu pliku dalej „niezapisane”", dirty, `hasUnsavedChanges=${dirty}`);
      check("B: wpisany tekst i nowy akapit w pliku", /WPISANE-B/.test(r.all) && /\nPO-ENTERZE-B/.test(r.all), "");
      if (errors.length) check("B: brak błędów strony", false, errors.slice(0, 2).join(" | "));
      await context.close();
    }
    // C) Enter, którego nie udało się nanieść na plik — podgląd wraca do pliku, wpisany tekst zostaje
    {
      const { context, page, errors } = await openDoc(browser);
      const before = await page.evaluate(() => collectPreviewParagraphElements(docCanvasEl.querySelector(".docx-preview-host")).length);
      await caretAtEnd(page, "__pa");
      await page.keyboard.type(" WPISANE-C");
      await caretAtEnd(page, "__pb");
      await page.evaluate(() => { window.__failSplit = 1; });
      await page.keyboard.press("Enter");
      await page.keyboard.type("PO-ENTERZE-C");
      await settle(page);
      const r = await compare(page);
      check("C: nieudany Enter — zapis = podgląd", r.same, `DOM ${r.nDom} / plik ${r.nXml}, różnica w ${r.diff}: ${JSON.stringify(r.dom)} ≠ ${JSON.stringify(r.xml)}`);
      check("C: wpisany wcześniej tekst zachowany", /WPISANE-C/.test(r.all), "");
      check("C: tekst po Enterze zachowany (w akapicie, którego podział się nie udał)", /PO-ENTERZE-C/.test(r.all) && r.nDom === before, `akapitów ${r.nDom} (przed: ${before})`);
      const unexpected = errors.filter((m) => !/test: nieudany podział/.test(m));
      if (unexpected.length) check("C: brak błędów strony", false, unexpected.slice(0, 2).join(" | "));
      await context.close();
    }
    // D) Backspace (scalenie akapitów), którego nie udało się nanieść — plik i podgląd znów zgodne,
    // tekst wpisany w INNYCH akapitach zostaje (scalonego akapitu nie da się rozdzielić z powrotem)
    {
      const { context, page, errors } = await openDoc(browser);
      const before = await page.evaluate(() => collectPreviewParagraphElements(docCanvasEl.querySelector(".docx-preview-host")).length);
      await caretAtEnd(page, "__pa");
      await page.keyboard.type(" WPISANE-D");
      await page.evaluate(() => {
        const ps = collectPreviewParagraphElements(docCanvasEl.querySelector(".docx-preview-host"));
        const el = ps[window.__pb];
        el.closest("[contenteditable=true]")?.focus();
        const r = document.createRange();
        r.setStart(el.firstChild || el, 0);
        r.collapse(true);
        getSelection().removeAllRanges();
        getSelection().addRange(r);
        window.__failMerge = 1;
      });
      await page.keyboard.press("Backspace");
      await settle(page);
      const r = await compare(page);
      check("D: nieudany Backspace — zapis = podgląd", r.same, `DOM ${r.nDom} / plik ${r.nXml}, różnica w ${r.diff}: ${JSON.stringify(r.dom)} ≠ ${JSON.stringify(r.xml)}`);
      check("D: tekst w innym akapicie zachowany, podział jak w pliku", /WPISANE-D/.test(r.all) && r.nDom === before, `akapitów ${r.nDom} (przed: ${before})`);
      const unexpected = errors.filter((m) => !/test: nieudane scalenie/.test(m));
      if (unexpected.length) check("D: brak błędów strony", false, unexpected.slice(0, 2).join(" | "));
      await context.close();
    }
  } finally {
    await browser.close();
  }
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ wyścig Enter/scalanie/zapis [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
