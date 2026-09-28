// user-flows-playwright.js — funkcje aplikacji klikane JAK UŻYTKOWNIK (mysz, klawiatura,
// przyciski w panelu), a nie przez wołanie funkcji. Łapie błędy typu „kliknięcie w przycisk
// zabiera kursor z dokumentu” (placeholder/snippet: „Ustaw kursor w akapicie”, choć był).
// Każdy krok sprawdza TREŚĆ PLIKU (document.xml), nie tylko podgląd.
//
//   node scripts/user-flows-playwright.js                (Chromium)
//   ENGINE=webkit node scripts/user-flows-playwright.js  (WebKit)

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = process.platform === "darwin" && ENGINE === "webkit" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1366, height: 900 }, acceptDownloads: true });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.clear(); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept(d.type() === "prompt" ? d.defaultValue() : undefined));
  const toasts = [];
  await page.exposeFunction("__toast", (m) => toasts.push(m));

  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.waitForTimeout(500);
  await page.evaluate(() => {
    const orig = window.toast;
    window.toast = (m, type) => { window.__toast(`${type}: ${m}`); return orig(m, type); };
  });

  // ── pomocnicze ─────────────────────────────────────────────────────────────
  const fileParas = () => page.evaluate(async () => {
    await waitInlineStructuralIdle?.();
    const bytes = await buildDocumentForSave(); // to, co trafiłoby do pliku przy zapisie
    const z = await JSZip.loadAsync(bytes);
    const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
    return collectParagraphElements(d.documentElement, "all").map((p) => ({ text: paragraphSearchText(p), xml: new XMLSerializer().serializeToString(p) }));
  });
  const para = (i) => page.evaluate((i) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    p.scrollIntoView({ block: "center" });
    return true;
  }, i);
  // klik na końcu tekstu akapitu i (jak człowiek)
  const clickEndOf = async (i) => {
    await para(i);
    const box = await page.evaluate((i) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
      const r = document.createRange(); r.selectNodeContents(p);
      const rects = r.getClientRects(); const last = rects[rects.length - 1];
      return { x: Math.min(last.right + 4, p.getBoundingClientRect().right - 1), y: last.top + last.height / 2 };
    }, i);
    await page.mouse.click(box.x, box.y); // tuż za końcem tekstu = kursor na końcu (End na Macu w WebKit nie działa)
  };
  // zaznacz słowo dwuklikiem
  const dblWord = async (i, word) => {
    await para(i);
    const box = await page.evaluate(([i, word]) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
      const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      let n;
      while ((n = w.nextNode())) {
        const at = n.textContent.indexOf(word);
        if (at >= 0) { const r = document.createRange(); r.setStart(n, at); r.setEnd(n, at + word.length); const b = r.getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + b.height / 2 }; }
      }
      return null;
    }, [i, word]);
    await page.mouse.dblclick(box.x, box.y);
  };
  const openPanel = async (id) => {
    await page.evaluate((id) => {
      if (!document.documentElement.classList.contains("sidebar-open")) document.getElementById("panelToggle").click();
      const d = document.getElementById(id);
      d.scrollIntoView({ block: "start" });
    }, id);
    await page.waitForTimeout(250);
    if (!(await page.evaluate((id) => document.getElementById(id).open, id))) await page.click(`#${id} > summary`);
    await page.waitForTimeout(400);
  };
  const editMode = async () => { await page.click('.mode-btn[data-mode="edit"]'); await page.waitForTimeout(400); };
  const readMode = async () => { await page.click('.mode-btn[data-mode="read"]'); await page.waitForTimeout(300); };
  const lastToast = () => toasts[toasts.length - 1] || "";

  // ── 1. placeholder: kursor w tekście → pole nazwy → „Wstaw” ─────────────────
  await editMode();
  await clickEndOf(1);
  await openPanel("panel-placeholders");
  await page.click("#phInsertName");
  await page.fill("#phInsertName", "klient");
  await page.click("#phInsertBtn");
  await page.waitForTimeout(200);
  let f = await fileParas();
  check("placeholder: „Wstaw” po kliknięciu w pole nazwy wstawia {{klient}} tam, gdzie był kursor",
    /Struktura\.\{\{klient\}\}$|Struktura\. ?\{\{klient\}\}$/.test(f[1].text), `${f[1].text} | toast: ${lastToast()}`);
  await page.click("#phScanBtn");
  await page.waitForSelector('#phForm input[data-ph-name="klient"]', { timeout: 4000 }).catch(() => {});
  const phField = await page.$('#phForm input[data-ph-name="klient"]');
  check("placeholder: „Skanuj” widzi nowe pole {{klient}}", !!phField, await page.textContent("#phStatus"));
  if (phField) {
    await phField.fill("ACME sp. z o.o.");
    await page.click("#phFillBtn");
    await page.waitForTimeout(700);
    f = await fileParas();
    check("placeholder: „Wypełnij” wpisuje wartość do pliku", f[1].text.includes("ACME sp. z o.o.") && !f.some((p) => p.text.includes("{{klient}}")), f[1].text);
  }

  // ── 2. snippety: zapisz → kursor w tekście → „Wstaw treść” / „Wstaw !trigger” → Rozwiń ──
  await editMode();
  await openPanel("panel-snippets");
  await page.fill("#snName", "podpis");
  await page.fill("#snBody", "Z poważaniem");
  await page.click("#snSaveBtn");
  await clickEndOf(3);
  await page.click("#snInsertBtn");
  await page.waitForTimeout(200);
  f = await fileParas();
  check("snippet: „Wstaw treść” wstawia w miejscu kursora", f[3].text.endsWith("Z poważaniem"), `${f[3].text.slice(-60)} | ${lastToast()}`);
  await clickEndOf(4);
  await page.click("#snInsertTriggerBtn");
  await page.waitForTimeout(200);
  f = await fileParas();
  check("snippet: „Wstaw !trigger” wstawia !podpis", f[4].text.endsWith("!podpis"), `${f[4].text.slice(-40)} | ${lastToast()}`);
  await page.click("#snExpandBtn");
  await page.waitForTimeout(700);
  f = await fileParas();
  check("snippet: „Rozwiń w dokumencie” zamienia !podpis na treść", f[4].text.endsWith("Z poważaniem") && !f.some((p) => p.text.includes("!podpis")), f[4].text.slice(-40));

  // ── 3. formatowanie: dwuklik słowa → B na pasku; rozmiar z listy ────────────
  await editMode();
  await dblWord(6, "Najemca");
  await page.click("#fmtBold");
  await page.waitForTimeout(200);
  f = await fileParas();
  check("B na pasku pogrubia zaznaczone słowo (w pliku)", /<w:b( w:val="(1|true)")?\/><\/w:rPr><w:t[^>]*>Najemca</.test(f[6].xml), f[6].xml.slice(0, 300));
  await dblWord(6, "lokalu");
  await page.selectOption("#fmtFontSize", "18");
  await page.waitForTimeout(200);
  f = await fileParas();
  check("rozmiar z listy (18) zmienia zaznaczone słowo (w pliku)", /<w:sz w:val="36"\/>/.test(f[6].xml), f[6].xml.slice(0, 400));

  // ── 4. cofnij / ponów przyciskami ───────────────────────────────────────────
  await clickEndOf(8);
  await page.keyboard.type(" DOPISEK");
  await page.waitForTimeout(300);
  await page.click("#undoBtn");
  await page.waitForTimeout(600);
  f = await fileParas();
  check("↶ cofa wpisany tekst", !f[8].text.includes("DOPISEK"), f[8].text.slice(-30));
  await page.click("#redoBtn");
  await page.waitForTimeout(600);
  f = await fileParas();
  check("↷ przywraca wpisany tekst", f[8].text.endsWith(" DOPISEK"), f[8].text.slice(-30));

  // ── 5. Znajdź i zamień z paska + panelu ─────────────────────────────────────
  await readMode();
  await page.click("#searchQuery");
  await page.fill("#searchQuery", "lokalu");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(500);
  const pos1 = await page.textContent("#searchPos");
  check("szukanie z paska: Enter pokazuje „1 / N”", /^1 \/ \d+$/.test(pos1.trim()), pos1);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(200);
  check("drugi Enter → następne trafienie", /^2 \/ \d+$/.test((await page.textContent("#searchPos")).trim()), await page.textContent("#searchPos"));
  await openPanel("panel-search");
  await page.fill("#frReplace", "LOKALU");
  const before = (await fileParas()).map((p) => p.text).join("\n").split("lokalu").length - 1;
  await page.click("#frReplaceOneBtn");
  await page.waitForTimeout(700);
  const after = (await fileParas()).map((p) => p.text).join("\n");
  check("„Zamień bieżące” zamienia jedno wystąpienie", after.split("LOKALU").length - 1 === 1 && after.split("lokalu").length - 1 === before - 1, `${before} → ${after.split("lokalu").length - 1}`);
  await page.click("#frReplaceAllBtn");
  await page.waitForTimeout(800);
  const afterAll = (await fileParas()).map((p) => p.text).join("\n");
  check("„Zamień wszystkie” zamienia resztę", !/lokalu/.test(afterAll) && afterAll.split("LOKALU").length - 1 === before, String(afterAll.split("LOKALU").length - 1));

  // ── 6. Narzędzia edycji: wielkie litery w nagłówkach? (cały dokument: prefiks) ─
  await openPanel("panel-edit-tools");
  await page.selectOption("#editOp", "affix").catch(() => {});
  await page.waitForTimeout(150);
  const affixVisible = await page.isVisible("#editPrefix");
  if (affixVisible) {
    await page.fill("#editPrefix", "» ");
    await page.selectOption("#editScope", { index: 0 }).catch(() => {});
    await page.click("#applyEditToolBtn");
    await page.waitForTimeout(800);
    f = await fileParas();
    check("Narzędzia edycji: prefiks „» ” dodany do akapitów", f.filter((p) => p.text.trim()).every((p) => p.text.startsWith("» ")), f.slice(0, 3).map((p) => p.text.slice(0, 20)).join(" | "));
    await page.click("#undoBtn");
    await page.waitForTimeout(800);
    f = await fileParas();
    check("…i ↶ to cofa", !f.some((p) => p.text.startsWith("» ")), f[0].text.slice(0, 20));
  } else check("Narzędzia edycji: pole prefiksu widoczne po wyborze operacji", false);

  // ── 7. Struktura: klik w nagłówek przewija; szybka edycja akapitu ───────────
  await openPanel("panel-structure");
  const item = await page.$("#structureOutline [data-id], #structureOutline button, #structureOutline li");
  check("Struktura: lista elementów jest", !!item);
  if (item) {
    await item.click();
    await page.waitForTimeout(400);
    check("Struktura: klik podświetla element w dokumencie", await page.evaluate(() => !!document.querySelector(".docx-preview-host .search-hit-active")));
  }

  // ── 8. Korekta ──────────────────────────────────────────────────────────────
  await editMode();
  await clickEndOf(9);
  await page.keyboard.type("  podwójna spacja .");
  await clickEndOf(6); // akapit z pogrubionym „Najemca” — poprawka nie może zdjąć formatu
  await page.keyboard.type("  koniec .");
  await readMode();
  await openPanel("panel-grammar");
  await page.click("#grammarScanBtn");
  await page.waitForTimeout(800);
  const gCount = await page.$$eval("#grammarSuggestions details, #grammarSuggestions .grammar-hit, #grammarSuggestions button", (e) => e.length);
  check("Korekta: skan znajduje wpisane „  ” i „ .”", gCount > 0, await page.textContent("#grammarStatus"));
  await page.click("#grammarApplyAllBtn").catch(() => {});
  await page.waitForTimeout(800);
  f = await fileParas();
  check("Korekta: „Popraw wszystko” usuwa podwójną spację i spację przed kropką", !/ {2}podwójna/.test(f[9].text) && !/spacja \./.test(f[9].text), f[9].text.slice(-40));
  check("Korekta: poprawiony akapit ZACHOWUJE formatowanie (pogrubienie z kroku 3 w akapicie 7)", /<w:b( w:val="(1|true)")?\/><\/w:rPr><w:t[^>]*>Najemca</.test(f[6].xml), f[6].xml.slice(0, 200));

  // ── 9. Metadane z formularza ────────────────────────────────────────────────
  await openPanel("panel-metadata");
  await page.fill("#metaTitle", "Tytuł z testu");
  await page.click("#metaApplyBtn");
  await page.waitForTimeout(600);
  const title = await page.evaluate(async () => (await extractCoreMetadataFromDocx(await buildDocumentForSave())).title);
  check("Metadane: „Zastosuj” zapisuje tytuł", title === "Tytuł z testu", title);

  // ── 10. Zapisz jako → pobranie pliku, który się otwiera ─────────────────────
  // bez File System Access (Safari/iPhone) „Zapisz jako” = pobranie; w Chromium wyłączamy okno wyboru
  await page.evaluate(() => { delete window.showSaveFilePicker; window.showSaveFilePicker = undefined; });
  const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 8000 }).catch(() => null), page.evaluate(() => document.getElementById("saveAsBtn").click())]);
  if (dl) {
    const path = await dl.path();
    const bytes = require("fs").readFileSync(path);
    const ok = await page.evaluate(async (arr) => {
      const z = await JSZip.loadAsync(new Uint8Array(arr));
      const x = await z.file("word/document.xml").async("string");
      return { p1: collectParagraphElements(new DOMParser().parseFromString(x, "application/xml").documentElement, "all").map(paragraphSearchText)[1], acme: /ACME sp\. z o\.o\./.test(x), lok: /LOKALU/.test(x), lower: /lokalu/.test(x), podpis: /Z poważaniem/.test(x) };
    }, Array.from(bytes));
    check("„Zapisz jako” daje plik ze wszystkimi zmianami", ok.acme && ok.lok && !ok.lower && ok.podpis, `${dl.suggestedFilename()} ${JSON.stringify(ok)}`);
  } else check("„Zapisz jako” daje plik do pobrania", false, "brak pobrania (może FSA)");

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ przepływy użytkownika [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
