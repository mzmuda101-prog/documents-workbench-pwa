// page-gaps-edit-playwright.js — odstępy między stronami w Edycji przy Enter / Backspace
// i wyjściu z listy; krój pustego akapitu.
//
//   node scripts/page-gaps-edit-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/page-gaps-edit-playwright.js
//
// Zgłoszenie Mateusza 2026-10-05 (wykład, koniec dokumentu, Safari):
//  1) Enter / Backspace w pustym punkcie listy (koniec listy) — po przerysowaniu widok uciekał
//     o kilka stron w górę: obszar przewijania był jeszcze krótszy (odstępy stron dochodzą po
//     przerysowaniu), więc powrót na miejsce obcinało do jego dołu;
//  2) Enter / Backspace (bez „input”, wysokość ostatniej kartki stała) nie przeliczały granic
//     stron — tekst wjeżdżał pod pas „str. N” albo zostawała pusta dziura;
//  3) pusty / nowy akapit miał krój APLIKACJI (Space Grotesk 12 pt), nie dokumentu (Calibri 11);
//  4) (decyzja Mateusza) długi akapit dzieli się między strony jak w Wordzie — dawniej cały szedł
//     na następną stronę i zostawiał dziurę.
// Plik budowany w teście: docDefaults Calibri 11 pt, ~3 strony tekstu, na końcu lista punktowana.

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const LOREM = "Wiara opiera się na przekonującym dowodzie, a nie na ślepym zaufaniu. Badanie Biblii, obserwacja przyrody i własne doświadczenia to cegiełki, z których buduje się mocną wiarę.";

async function buildDocx() {
  const para = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  const item = (t) => `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  const body = [];
  for (let i = 1; i <= 16; i++) body.push(para(`Akapit ${i}. ${LOREM} ${LOREM} ${LOREM}`));
  body.push(item("Podsumuj główne punkty."), item("Zachęć do działania."), item("Zakończ wezwaniem, które Bóg nam udostępnia."));
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>');
  zip.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>');
  zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="pl-PL"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style></w:styles>`);
  zip.file("word/numbering.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body.join("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: "nodebuffer" });
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1400, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.setItem("dwb-page-gaps-v1", "1"); localStorage.setItem("dwb-panel-docked-open-v1", "0"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "wyklad-test.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await buildDocx() });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForFunction(() => !readOnlyMode && document.querySelector(".docx-edit-root"), null, { timeout: 10000 });
  await sleep(900);

  const idle = async () => {
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 });
    await page.evaluate(() => waitInlineStructuralIdle());
    await sleep(900); // przeliczenie granic stron (opóźnione, w wolnej chwili)
  };
  const caretAtEnd = (text) => page.evaluate((text) => {
    const el = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.includes(text));
    el.scrollIntoView({ block: "center" });
    docEditRoot().focus({ preventScroll: true });
    const r = document.createRange(); r.selectNodeContents(el); r.collapse(false);
    getSelection().removeAllRanges(); getSelection().addRange(r);
  }, text);
  const caretAtStart = (text) => page.evaluate((text) => {
    const el = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.startsWith(text));
    docEditRoot().focus({ preventScroll: true });
    const r = document.createRange(); r.selectNodeContents(el); r.collapse(true);
    getSelection().removeAllRanges(); getSelection().addRange(r);
  }, text);
  const spot = () => page.evaluate(() => {
    const p = docCaretParagraph(document.activeElement);
    return { scroll: Math.round(docViewportEl.scrollTop), top: p ? Math.round(p.getBoundingClientRect().top) : null, inList: !!p && /docx-num-/.test(p.className) };
  });
  // żadna linijka tekstu nie leży pod pasem przerwy między stronami; odstęp stoi tuż przed blokiem
  const gapsOk = () => page.evaluate(() => {
    const bad = [];
    const bands = [...document.querySelectorAll(".dwb-page-gap-band")].map((b) => ({ r: b.getBoundingClientRect(), label: b.dataset.label }));
    const range = document.createRange();
    document.querySelectorAll(".docx-preview-host section.docx > article p").forEach((p) => {
      range.selectNodeContents(p);
      const rects = [...range.getClientRects()].filter((q) => q.height > 0);
      for (const q of rects) for (const b of bands) if (q.bottom > b.r.top + 1 && q.top < b.r.bottom - 1) bad.push(`${b.label}: „${p.textContent.slice(0, 24)}”`);
    });
    const orphan = [...document.querySelectorAll(".dwb-page-gap")].filter((g) => !g.nextElementSibling?.matches("p, table")).length;
    return { ok: !bad.length && !orphan, bad: [...new Set(bad)].slice(0, 3), orphan, bands: bands.length };
  });

  // granice w środku akapitu: wiersz po granicy na górnym marginesie nowej strony, przed nią ≥ 2 wiersze
  const splitInfo = () => page.evaluate(() => {
    const out = [];
    document.querySelectorAll(".dwb-page-gap.dwb-gap-split").forEach((g) => {
      const sec = g.closest("section.docx");
      const sr = sec.getBoundingClientRect();
      const sc = sr.height / sec.offsetHeight || 1;
      const padT = parseFloat(getComputedStyle(sec).paddingTop);
      const bands = [...sec.querySelectorAll(".dwb-page-gap-band")].map((b) => ({ top: parseFloat(b.style.top), h: parseFloat(b.style.height) }));
      const p = g.nextElementSibling;
      const r = document.createRange(); r.selectNodeContents(p);
      const rects = [...r.getClientRects()].filter((q) => q.height > 0).sort((a, b) => a.top - b.top);
      const lines = [];
      rects.forEach((q) => { const l = lines[lines.length - 1]; if (l && q.top < l.bottom - 2) l.bottom = Math.max(l.bottom, q.bottom); else lines.push({ top: q.top, bottom: q.bottom }); });
      const k = +g.dataset.line;
      const loc = (y) => (y - sr.top) / sc;
      // pas przerwy tuż nad wierszem po granicy
      const band = lines[k] ? bands.filter((x) => x.top + x.h <= loc(lines[k].top)).sort((x, y) => y.top - x.top)[0] : null;
      out.push({ line: k, lines: lines.length, topDiff: band && lines[k] ? Math.round(loc(lines[k].top) - (band.top + band.h) - padT) : null, prevAbove: band && lines[k - 1] ? loc(lines[k - 1].bottom) <= band.top : null, dbg: band && lines[k - 1] && [Math.round(loc(lines[k - 1].bottom)), band.top, Math.round(loc(lines[k].top)), lines.map((l) => Math.round(loc(l.top)))] });
    });
    return out;
  });
  // ── 3) krój pustego akapitu = krój dokumentu ──
  const last = "Bóg nam udostępnia";
  await caretAtEnd(last);
  await page.keyboard.press("Enter"); // nowy pusty punkt listy
  await idle();
  const font1 = await page.evaluate(() => {
    const p = docCaretParagraph(document.activeElement);
    const cs = getComputedStyle(p);
    return { family: cs.fontFamily, size: cs.fontSize, field: document.getElementById("fmtFontFamily")?.value, empty: !p.textContent };
  });
  check("pusty nowy punkt listy: krój i rozmiar dokumentu (Calibri 11 pt), nie aplikacji", font1.empty && /^"?Calibri/.test(font1.family) && Math.abs(parseFloat(font1.size) - 14.667) < 0.1 && font1.field === "Calibri", JSON.stringify(font1));

  // ── 1) Enter w pustym punkcie na końcu dokumentu = koniec listy, widok stoi ──
  const before = await spot();
  await page.keyboard.press("Enter");
  await idle();
  const after = await spot();
  check("Enter w pustym punkcie (koniec dokumentu): koniec listy, akapit w tym samym miejscu ekranu", !after.inList && Math.abs(after.top - before.top) <= 3 && Math.abs(after.scroll - before.scroll) <= 3, JSON.stringify({ before, after }));
  const font2 = await page.evaluate(() => { const cs = getComputedStyle(docCaretParagraph(document.activeElement)); return { family: cs.fontFamily, size: cs.fontSize }; });
  check("akapit po liście: Calibri 11 pt", /^"?Calibri/.test(font2.family) && Math.abs(parseFloat(font2.size) - 14.667) < 0.1, JSON.stringify(font2));
  await page.evaluate(() => dwbUndo.undo());
  await idle();
  const b2 = await spot();
  await page.evaluate(() => { const p = docCaretParagraph(document.activeElement); const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await page.keyboard.press("Backspace");
  await idle();
  const a2 = await spot();
  check("Backspace w pustym punkcie (koniec dokumentu): koniec listy, widok stoi", b2.inList && !a2.inList && Math.abs(a2.top - b2.top) <= 3 && Math.abs(a2.scroll - b2.scroll) <= 3, JSON.stringify({ b2, a2 }));

  // ── 2) Enter / Backspace przy granicy strony: odstęp idzie za akapitami ──
  const g0 = await gapsOk();
  check("na starcie: odstępy między stronami są, tekst nie leży pod pasem", g0.ok && g0.bands >= 1, JSON.stringify(g0));
  const sp0 = await splitInfo();
  check("długi akapit dzieli się między strony jak w Wordzie: ≥ 2 wiersze przed granicą, dalszy wiersz na górnym marginesie następnej strony", sp0.length >= 1 && sp0.every((x) => x.line >= 2 && x.lines - x.line >= 2 && x.prevAbove && Math.abs(x.topDiff) <= 3), JSON.stringify(sp0));
  // pierwszy akapit strony 2 (ten, przed którym stoi odstęp)
  const startP2 = await page.evaluate(() => document.querySelector(".dwb-page-gap")?.nextElementSibling?.textContent.slice(0, 10));
  // kilka pustych akapitów na końcu strony 1 (Enter na początku pierwszego akapitu strony 2)
  await caretAtStart(startP2);
  for (let i = 0; i < 4; i++) { await page.keyboard.press("Enter"); await sleep(60); }
  await idle();
  const g1 = await gapsOk();
  check("Enter ×4 przed granicą strony: odstęp przesunięty, żaden tekst pod pasem „str. 2”", g1.ok, JSON.stringify(g1));
  // Backspace na początku tego akapitu — skleja go z pustym akapitem przed nim (przed odstępem)
  for (let i = 0; i < 4; i++) {
    await caretAtStart(startP2);
    await page.keyboard.press("Backspace");
    await sleep(60);
  }
  await idle();
  const g2 = await gapsOk();
  const empties = await page.evaluate(() => [...document.querySelectorAll(".docx-preview-host section.docx > article > p")].filter((p) => !p.textContent).length);
  check("Backspace przy granicy strony: puste akapity znikają, tekst nie wjeżdża pod pas", g2.ok && empties <= 1, JSON.stringify({ ...g2, empties }));
  // sklejenie dwóch długich akapitów przez granicę strony (Backspace na początku pierwszego akapitu
  // strony 2, przed nim tekst) — sklejony akapit nie może zostać pod pasem przerwy
  const firstOfPage = await page.evaluate(() => [...document.querySelectorAll(".dwb-page-gap")].map((g) => g.nextElementSibling?.textContent.slice(0, 10)).find((t) => t && t.startsWith("Akapit")));
  await caretAtStart(firstOfPage);
  await page.keyboard.press("Backspace");
  await idle();
  const g3 = await gapsOk();
  const merged = await page.evaluate((t) => [...document.querySelectorAll(".docx-preview-host p")].some((p) => p.textContent.indexOf(t) > 0), firstOfPage);
  check("Backspace skleja długie akapity przez granicę strony: granica przeliczona, tekst nie pod pasem", merged && g3.ok, JSON.stringify({ merged, ...g3 }));
  // koniec dokumentu (jak w zgłoszeniu): puste akapity aż do nowej strony, długi wpisany akapit,
  // Backspace na jego początku — wjeżdża na koniec poprzedniej strony
  const endText = await page.evaluate(() => { const ps = [...document.querySelectorAll(".docx-preview-host section.docx > article > p")]; const p = ps[ps.length - 1]; p.scrollIntoView({ block: "center" }); docEditRoot().focus({ preventScroll: true }); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); return p.textContent; });
  let pushed = false;
  for (let i = 0; i < 80 && !pushed; i++) {
    await page.keyboard.press("Enter");
    await sleep(40);
    if (i % 5 === 4) {
      await idle();
      pushed = await page.evaluate(() => { const p = docCaretParagraph(document.activeElement); const prev = p?.previousElementSibling; return !!prev?.matches(".dwb-page-gap"); });
    }
  }
  const nested = await page.evaluate(() => document.querySelectorAll(".docx-preview-host p p").length);
  check("Enter w pustym akapicie (po wyjściu z listy, puste z pliku): nowy akapit obok, nie w środku", pushed && nested === 0, JSON.stringify({ pushed, nested }));
  await page.keyboard.insertText(`Nr 54 ${LOREM} ${LOREM} ${LOREM}`);
  await idle();
  for (let i = 0; i < 2; i++) {
    await caretAtStart("Nr 54");
    await page.keyboard.press("Backspace");
    await idle();
  }
  const g4 = await gapsOk();
  check("koniec dokumentu: długi akapit po pustych, Backspace — granica stron przeliczona, tekst nie pod pasem", pushed && g4.ok, JSON.stringify({ pushed, endText: endText.slice(0, 20), ...g4 }));
  // plik = podgląd (odstępy to tylko wygląd)
  const same = await page.evaluate(async () => {
    const bytes = await buildDocumentForSave();
    const xml = (await extractParagraphTextsFromDocx(bytes)).map(String);
    const dom = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => previewRunsToPlainText(extractRunsFromPreviewParagraph(p)));
    let i = 0; while (i < xml.length && xml[i] === dom[i]) i++;
    return { same: JSON.stringify(xml) === JSON.stringify(dom), nXml: xml.length, nDom: dom.length, at: i, xml: xml.slice(i - 1, i + 3).map((t) => t.slice(0, 30)), dom: dom.slice(i - 1, i + 3).map((t) => t.slice(0, 30)) };
  });
  check("zapis = podgląd (akapity i tekst)", same.same, JSON.stringify(same));

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  let fail = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail})`}`);
    if (!r.ok) fail++;
  }
  console.log(`\n[${ENGINE}] ${fail ? `${fail} z ${results.length} nie przeszło` : `${results.length}/${results.length} OK`}`);
  process.exit(fail ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
