// run-fidelity-playwright.js — pisanie w akapicie z pliku nie zmienia formatu, którego model
// akapitu nie zna (2026-10-05, polowanie na błędy: zapis 52 prawdziwych plików).
//
// Znalezione: (1) krój z podglądu „"DM Sans", sans-serif” szedł do pliku jako „DM Sans"”;
// (2) w:rPr pisanego akapitu był budowany od zera — ginęły język, kapitaliki, odstęp liter,
// krój motywu; (3) <w:b w:val="0"/> czytane jako pogrubienie; (4) symbol w:sym („§” z kroju
// Symbol) po cichu znikał z pisanego akapitu; (5) pliki z (1) zapisane wcześniej — naprawa;
// (6) BOM w XML: Safari nie zapisywał wpisanego tekstu.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const t = (s) => `<w:t xml:space="preserve">${s}</w:t>`;

async function fixture() {
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  z.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Gruby"><w:name w:val="Gruby akapit"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style></w:styles>`);
  const body = [
    // 1. właściwości spoza modelu: język, kapitaliki, odstęp liter, krój motywu
    `<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:smallCaps/><w:spacing w:val="20"/><w:lang w:val="de-DE"/></w:rPr>${t("Kapitaliki z odstępem i językiem")}</w:r></w:p>`,
    // 2. w pogrubionym stylu akapitu fragment jawnie NIE pogrubiony
    `<w:p><w:pPr><w:pStyle w:val="Gruby"/></w:pPr><w:r>${t("Pogrubiony ze stylu, ")}</w:r><w:r><w:rPr><w:b w:val="0"/></w:rPr>${t("a ten nie")}</w:r></w:p>`,
    // 3. symbol w osobnym fragmencie („§” z kroju Symbol = F0A7)
    `<w:p><w:r>${t("Art. 29 ")}</w:r><w:r><w:sym w:font="Symbol" w:char="F0A7"/></w:r><w:r>${t(" 11 Kodeksu pracy")}</w:r></w:p>`,
    // 4. symbol i tekst w jednym fragmencie — akapit tylko do odczytu
    `<w:p><w:r><w:sym w:font="Wingdings" w:char="F0A8"/>${t(" Tak, zgadzam się")}</w:r></w:p>`,
    // 5. zepsuta nazwa kroju (starsza wersja aplikacji)
    `<w:p><w:r><w:rPr><w:rFonts w:ascii="Georgia&quot;" w:hAnsi="Georgia&quot;"/></w:rPr>${t("Tekst z zepsutą nazwą kroju")}</w:r></w:p>`,
  ].join("");
  // BOM na początku (tak zapisują niektóre programy .NET) — Safari odrzucał taki XML i zapis gubił tekst
  z.file("word/document.xml", `\uFEFF<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return z.generateAsync({ type: "nodebuffer" });
}

const savedParas = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  const doc = await z.file("word/document.xml").async("string");
  return doc.match(/<w:body>([\s\S]*)<w:sectPr/)[1].match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
});

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const file = path.join(os.tmpdir(), `dwb-wiernosc-${process.pid}.docx`);
  fs.writeFileSync(file, await fixture());
  await page.setInputFiles("#fileInput", file);
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  fs.rmSync(file, { force: true });
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await idle(page);

  const locks = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => p.dataset.lock || null));
  check("symbol w osobnym fragmencie: akapit edytowalny", !locks[2], JSON.stringify(locks));
  check("symbol razem z tekstem w jednym fragmencie: tylko do odczytu (lockSymbol)", locks[3] === "lockSymbol", JSON.stringify(locks));
  const same = await page.evaluate(async () => { inlineDirtyValid = false; const n = collectInlineParagraphEdits().length; inlineDirtyValid = true; return n; });
  check("Edycja bez pisania: żaden akapit nie wygląda na zmieniony (podgląd = plik)", same === 0, String(same));

  const typeEnd = async (i, text) => {
    await page.evaluate((i) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; placeCaret(p, p.textContent.length); }, i);
    await page.keyboard.type(text);
    await page.waitForTimeout(250);
  };
  await typeEnd(0, " plus");
  await typeEnd(1, " też");
  await typeEnd(2, " (zmiana)");
  await typeEnd(4, " i dopisek");
  const ps = await savedParas(page);
  check("pisanie: język, kapitaliki, odstęp liter i krój motywu zostają w pliku", /Kapitaliki z odstępem i językiem plus/.test(ps[0]) && /<w:smallCaps\/>/.test(ps[0]) && /<w:spacing w:val="20"\/>/.test(ps[0]) && /w:lang w:val="de-DE"/.test(ps[0]) && /w:asciiTheme="minorHAnsi"/.test(ps[0]) && !/var\(--/.test(ps[0]), ps[0].slice(0, 400));
  check("pisanie w jawnie NIE pogrubionym fragmencie pogrubionego akapitu: zostaje w:b w:val=\"0\"", /<w:b w:val="0"\/>(?:(?!<\/w:r>).)*a ten nie też/.test(ps[1]), ps[1].slice(0, 400));
  check("symbol „§” (w:sym) zostaje w pisanym akapicie", /<w:sym w:font="Symbol" w:char="F0A7"\/>/.test(ps[2]) && /Kodeksu pracy \(zmiana\)/.test(ps[2]), ps[2].slice(0, 400));
  check("zepsuta nazwa kroju naprawiona (Georgia, bez cudzysłowu)", /w:ascii="Georgia"/.test(ps[4]) && !/&quot;/.test(ps[4]) && /i dopisek/.test(ps[4]), ps[4].slice(0, 300));

  // Ctrl/⌘+B wyłączające pogrubienie w pogrubionym (ze stylu) akapicie → w:b w:val="0" w pliku
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1]; const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); const n = w.nextNode(); const r = document.createRange(); r.setStart(n, n.data.indexOf(",")); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
  await page.keyboard.press(process.platform === "darwin" ? "Meta+b" : "Control+b");
  await page.keyboard.type(" cienko");
  await page.waitForTimeout(300);
  const p1 = (await savedParas(page))[1];
  check("Ctrl/⌘+B w pogrubionym nagłówku: wpisany tekst zapisany z w:b w:val=\"0\"", /<w:b w:val="0"\/>(?:(?!<\/w:r>).)*<w:t[^>]*> cienko/.test(p1), p1.slice(0, 500));

  await browser.close();
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Wierność formatu przy pisaniu [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
