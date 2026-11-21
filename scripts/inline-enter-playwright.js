// inline-enter-playwright.js — szybkie pisanie wokół Enter/Backspace w podglądzie.
//
//   node scripts/inline-enter-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/inline-enter-playwright.js
//
// Błąd z 2026-09-28: Enter dzielił akapit w podglądzie od razu, ale kursor przeskakiwał
// dopiero po przebudowie pliku — tekst pisany w tym czasie lądował w złym akapicie,
// a ZAPISANY plik różnił się od podglądu (34 akapity na ekranie, 33 w pliku).
// Musi być prawdą: to, co widać, to dokładnie to, co trafia do pliku — przy każdym tempie,
// także na dużym dokumencie (przebudowa trwa wtedy najdłużej).
// Błąd z 2026-11-21: WebKit (iPhone, iPad) — Enter na końcu punktu listy dawał pusty punkt bez
// miejsca na kursor; pierwsza wpisana litera szła na początek NASTĘPNEGO akapitu.

const fs = require("fs");
const path = require("path");
const os = require("os");
const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL, DOCX_FIXTURE } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function session(browser, loadDoc, anchorText, delay, label) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await loadDoc(page);
  await page.evaluate(() => document.querySelector('.mode-btn[data-mode="edit"]').click());
  await page.locator(".docx-preview-host p", { hasText: anchorText }).first().click();
  // kursor na KOŃCU akapitu (End w długim akapicie idzie tylko na koniec wiersza)
  await page.evaluate((anchor) => {
    const el = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.includes(anchor));
    const r = document.createRange();
    r.selectNodeContents(el);
    r.collapse(false);
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  }, anchorText);
  await page.keyboard.type(" A", { delay });
  await page.keyboard.press("Enter");
  await page.keyboard.type("Pierwszy nowy", { delay });
  await page.keyboard.press("Enter");
  await page.keyboard.type("Drugi nowy", { delay });
  await page.keyboard.press("Enter");
  await page.keyboard.press("Backspace");
  await page.keyboard.type(" koniec", { delay });
  const r = await page.evaluate(async (anchor) => {
    const bytes = await buildDocumentForSave(); // czeka na kolejkę przebudowy
    const dom = [...document.querySelectorAll(".docx-preview-host p")].map((x) => x.textContent);
    const xml = (await extractParagraphTextsFromDocx(bytes)).map(String);
    const i = dom.findIndex((t) => t.includes(anchor));
    return { same: JSON.stringify(dom) === JSON.stringify(xml), nDom: dom.length, nXml: xml.length, around: dom.slice(i, i + 3).map((t) => t.slice(-24)) };
  }, anchorText);
  const expected = r.around[1] === "Pierwszy nowy" && r.around[2].endsWith("Drugi nowy koniec") && r.around[0].endsWith(" A");
  check(`${label}, tempo ${delay} ms: tekst w dobrych akapitach`, expected, JSON.stringify(r.around));
  check(`${label}, tempo ${delay} ms: zapis = podgląd`, r.same, `DOM ${r.nDom} / plik ${r.nXml}`);
  if (errors.length) check(`${label}: brak błędów strony`, false, errors.slice(0, 2).join(" | "));
  await context.close();
}

// lista numerowana bez wcięcia wiszącego (numer i tabulator jak w Wordzie) + akapit pod nią
async function listDocx() {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const item = (t) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>');
  zip.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>');
  zip.file("word/numbering.xml", `<?xml version="1.0" encoding="UTF-8"?><w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${item("Punkt pierwszy.")}${item("Punkt drugi.")}<w:p><w:r><w:t>Akapit pod listą</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const file = path.join(os.tmpdir(), `dwb-inline-enter-list-${process.pid}.docx`);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer" }));
  return file;
}

async function listEnterAtEnd(browser) {
  const file = await listDocx();
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles(file);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.evaluate(() => document.querySelector('.mode-btn[data-mode="edit"]').click());
  await page.waitForSelector(".docx-edit-root");
  await page.waitForTimeout(400);
  // kursor na koniec klawiszem (tak jak użytkownik), nie przez Range — tylko tak WebKit gubił miejsce
  await page.locator(".docx-preview-host p", { hasText: "Punkt drugi" }).first().click({ position: { x: 60, y: 8 } });
  await page.keyboard.press(ENGINE === "webkit" ? "Meta+ArrowRight" : "End");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(300);
  await page.keyboard.type("Nowy punkt");
  const r = await page.evaluate(async () => {
    const bytes = await buildDocumentForSave();
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    return { dom: ps.map((p) => `${p.getAttribute("data-dwb-num") || ""}${p.textContent}`), xml: (await extractParagraphTextsFromDocx(bytes)).map(String) };
  });
  check("lista: Enter na końcu punktu → tekst w nowym punkcie", r.dom[2] === "3.\tNowy punkt" && r.dom[3] === "Akapit pod listą", JSON.stringify(r.dom));
  check("lista: zapis = podgląd (bez dodatkowego łamania wiersza)", JSON.stringify(r.xml) === JSON.stringify(r.dom.map((t) => t.replace(/^\d+\.\t/, ""))), JSON.stringify(r.xml));
  await context.close();
  fs.rmSync(file, { force: true });
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const small = async (page) => {
    await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
    await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
    await page.waitForTimeout(500);
  };
  const big = async (page) => {
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.locator("#fileInput").setInputFiles(DOCX_FIXTURE);
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 60000 });
    await page.waitForTimeout(500);
  };
  try {
    for (const delay of [0, 60]) await session(browser, small, "Dane Wynajmującego", delay, "mały plik");
    await listEnterAtEnd(browser);
    if (fs.existsSync(DOCX_FIXTURE)) await session(browser, big, "Phasellus dapibus", 0, "duży plik (3,8 MB)");
  } finally {
    await browser.close();
  }
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ Enter/Backspace przy szybkim pisaniu [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
