// export-stats-playwright.js — paczka H: eksport (TXT/MD/HTML), statystyki, import/eksport
// snippetów i wartości pól, „Usuń dane osobowe”.

const fs = require("fs");
const { createTestPage, bootApp, loadDocxFile, assertNoErrors } = require("./docx-test-helpers");

const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function download(page, click) {
  const [dl] = await Promise.all([page.waitForEvent("download"), click()]);
  const path = await dl.path();
  return { name: dl.suggestedFilename(), text: fs.readFileSync(path, "utf8") };
}
async function openPanel(page, id) {
  await page.evaluate(async (id) => {
    setSidebarOpen(true);
    document.getElementById(id).open = true;
    await ensureLazyFeature(PANEL_LAZY_FEATURE[id]);
  }, id);
  // panel się wysuwa — kliknięcia dopiero, gdy stoi na miejscu
  await page.waitForFunction(() => document.getElementById("toolsPanel").getBoundingClientRect().left >= 0, null, { timeout: 5000 });
  await page.waitForTimeout(300);
}

async function run() {
  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await loadDocxFile(page, "docs/samples/headings-sample.docx");
  await page.waitForSelector(".docx-preview-host p");

  // ── eksport ────────────────────────────────────────────────────────────────
  await openPanel(page, "panel-export");
  const txt = await download(page, () => page.click("#exTxtBtn"));
  check("TXT: nazwa .txt i treść akapitów", txt.name.endsWith(".txt") && txt.text.includes("Umowa najmu lokalu") && txt.text.includes("Najemca zobowiązuje się"), txt.name);
  const md = await download(page, () => page.click("#exMdBtn"));
  check("Markdown: nagłówki jako #", /^# Umowa najmu/m.test(md.text) && /^#{2,3} Dane Wynajmującego/m.test(md.text), md.text.slice(0, 120));
  const html = await download(page, () => page.click("#exHtmlBtn"));
  check("HTML: dokument z <h1> i <p>", html.text.startsWith("<!doctype html>") && /<h1>Umowa najmu/.test(html.text) && html.text.includes("<p>"));
  check("HTML: znaki specjalne są escapowane", await page.evaluate(() => escHtmlEx('<a href="x">&</a>') === "&lt;a href=&quot;x&quot;&gt;&amp;&lt;/a&gt;"));

  // eksport widzi edycje (z DOM), tabela → MD
  const model = await page.evaluate(() => exportToMarkdown([
    { kind: "table", rows: [["A", "B|c"], ["1", "2"]] },
    { kind: "li", level: 0, text: "raz" }, { kind: "li", level: 1, text: "dwa" },
  ]));
  check("MD: tabela z nagłówkiem, | escapowane, lista zagnieżdżona", model.includes("| A | B\\|c |") && model.includes("| --- | --- |") && model.includes("- raz\n  - dwa"), model);
  await page.evaluate(() => { window.__printed = 0; });
  await page.click("#exPrintBtn");
  await page.waitForFunction(() => [...document.querySelectorAll("iframe")].some((f) => f.srcdoc.includes("docx-wrapper") || f.srcdoc.includes("<h1>")), null, { timeout: 5000 }).catch(() => {});
  check("Druk: ukryta ramka z treścią dokumentu", await page.evaluate(() => [...document.querySelectorAll("iframe")].some((f) => f.srcdoc.includes("Umowa najmu"))));

  // ── statystyki ─────────────────────────────────────────────────────────────
  await openPanel(page, "panel-stats");
  await page.waitForSelector("#stGrid .stat-cell");
  const cells = await page.$$eval("#stGrid .stat-cell", (els) => els.map((e) => e.textContent));
  check("Statystyki: 8 kafelków z liczbami", cells.length === 8 && /\d/.test(cells[0]), cells.join(" | "));
  const calc = await page.evaluate(() => stTextStats("Ala ma kota. Kot ma Alę! Czy tak?"));
  check("stTextStats: słowa/zdania/znaki", calc.words === 8 && calc.sentences === 3 && calc.chars === 33, JSON.stringify(calc));
  await page.fill("#stLongLimit", "10"); // fokus zostaje w polu — klik w wiersz musi zadziałać mimo to
  await page.waitForTimeout(400);
  const longRows = await page.$$eval("#stLongSentences .stat-row", (els) => els.length);
  check("Długie zdania (próg 10): są wiersze ze skokiem", longRows > 0, String(longRows));
  await page.waitForTimeout(200);
  await page.evaluate(() => document.querySelector("#stLongSentences .stat-row").scrollIntoView({ block: "center" }));
  await page.waitForTimeout(150);
  await page.click("#stLongSentences .stat-row");
  await page.waitForTimeout(100);
  check("Skok podświetla akapit", await page.evaluate(() => !!document.querySelector(".search-hit-active")), await page.evaluate(() => document.querySelector(".search-hit")?.outerHTML.slice(0, 80) || "brak .search-hit"));
  await page.evaluate(() => {
    const p = document.querySelector(".docx-preview-host p:nth-of-type(2)");
    const r = document.createRange(); r.selectNodeContents(p); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
    document.dispatchEvent(new Event("selectionchange"));
  });
  await page.waitForTimeout(100);
  check("Statystyki zaznaczenia", /Zaznaczenie: \d+ słów/.test(await page.textContent("#stSelection")), await page.textContent("#stSelection"));

  // ── snippety: eksport / import ─────────────────────────────────────────────
  await openPanel(page, "panel-snippets");
  await page.evaluate(() => { localStorage.removeItem(SNIPPETS_STORAGE_KEY); upsertSnippet("adres", "ul. Testowa 1, {{miasto}}"); upsertSnippet("stopka", "Pozdrawiam"); renderSnippetList(); });
  const sn = await download(page, () => page.click("#snExportBtn"));
  const snData = JSON.parse(sn.text);
  check("Eksport snippetów: JSON z 2 pozycjami", snData.type === "snippets" && snData.items.length === 2, sn.text.slice(0, 80));
  await page.evaluate(() => { localStorage.removeItem(SNIPPETS_STORAGE_KEY); upsertSnippet("adres", "STARE"); });
  fs.writeFileSync("/tmp/dwb-snippets.json", sn.text);
  await page.setInputFiles("#snImportFile", []);
  const importRes = await page.evaluate((json) => { const r = importSnippetsData(JSON.parse(json)); return { r, list: loadSnippets().map((s) => `${s.name}=${s.body}`) }; }, sn.text);
  check("Import snippetów: 1 nowy, 1 nadpisany", importRes.r.added === 1 && importRes.r.updated === 1 && importRes.list.includes("adres=ul. Testowa 1, {{miasto}}"), JSON.stringify(importRes));
  check("Import złego pliku odrzucony", await page.evaluate(() => importSnippetsData({ foo: 1 }) === null));

  // ── placeholdery: wartości ─────────────────────────────────────────────────
  await loadDocxFile(page, "docs/samples/sample.docx");
  await openPanel(page, "panel-placeholders");
  await page.click("#phScanBtn");
  await page.waitForSelector("#phForm input[data-ph-name]");
  const name = await page.$eval("#phForm input[data-ph-name]", (e) => e.dataset.phName);
  await page.fill("#phForm input[data-ph-name]", "Wartość 1");
  const ph = await download(page, () => page.click("#phExportBtn"));
  check("Eksport wartości pól: JSON z nazwą pola", JSON.parse(ph.text).values[name] === "Wartość 1", ph.text);
  await page.fill("#phForm input[data-ph-name]", "");
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.click("#phImportBtn")]);
  fs.writeFileSync("/tmp/dwb-ph.json", ph.text);
  await chooser.setFiles("/tmp/dwb-ph.json");
  await page.waitForTimeout(400);
  check("Import wartości wypełnia formularz", (await page.$eval("#phForm input[data-ph-name]", (e) => e.value)) === "Wartość 1");

  // ── metadane: usuń dane osobowe ────────────────────────────────────────────
  await openPanel(page, "panel-metadata");
  const scrub = await page.evaluate(async () => {
    await ensureDocLibs(false);
    await applyDocumentEdit({ op: "coreMetadata", fields: { creator: "Jan Kowalski", lastModifiedBy: "Anna Nowak", description: "poufne", subject: "temat", title: "Zostaje" } });
    const zip0 = await window.JSZip.loadAsync(originalFileBytes);
    const before = await zip0.file("docProps/core.xml").async("string");
    window.confirm = () => true;
    await scrubPersonalMetadata();
    const zip = await window.JSZip.loadAsync(originalFileBytes);
    const xml = await zip.file("docProps/core.xml").async("string");
    return { before, xml };
  });
  check("Przed czyszczeniem są dane osobowe", /Jan Kowalski/.test(scrub.before) && /Anna Nowak/.test(scrub.before));
  check("Po czyszczeniu: brak autora, ostatnio zmieniającego, opisu, tematu; tytuł zostaje",
    !/Jan Kowalski|Anna Nowak|poufne|temat/.test(scrub.xml) && /Zostaje/.test(scrub.xml), scrub.xml);

  assertNoErrors(errors, "export-stats");
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ eksport/statystyki/import: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
