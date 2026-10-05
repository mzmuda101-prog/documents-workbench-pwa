// object-islands-playwright.js — rysunki w akapicie jako „wyspy” (2026-10-05).
//
// Zgłoszenie: tytuł dokumentu z PDF → DOCX był tylko do odczytu („Akapit zawiera obraz lub
// obiekt”), bo do pierwszego akapitu strony przypięta jest warstwa grafiki. To samo w pismach
// z Worda (logo przypięte do akapitu, obrazek w zdaniu). Teraz fragment z samym rysunkiem jest
// wyspą: pisanie obok go nie gubi, zapis oddaje go bez zmian. Blokada zostaje tam, gdzie wyspa
// mogłaby coś zgubić: pole tekstowe w rysunku, rysunek we wspólnym fragmencie z tekstem.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

// 1×1 px PNG (czerwony)
const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==", "base64");
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const pic = (id) => `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Logo ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
const anchorRun = (id) => `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>5000000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>400000</wp:posOffset></wp:positionV><wp:extent cx="457200" cy="228600"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="Logo ${id}"/><wp:cNvGraphicFramePr/>${pic(id)}</wp:anchor></w:drawing></w:r>`;
const inlineRun = (id) => `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="457200" cy="228600"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Obraz ${id}"/><wp:cNvGraphicFramePr/>${pic(id)}</wp:inline></w:drawing></w:r>`;
const textRun = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const textbox = `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="9" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1800000" cy="400000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="9" name="Pole tekstowe"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p>${textRun("Tekst w polu")}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><w:txbxContent><w:p>${textRun("Tekst w polu")}</w:p></w:txbxContent></w:pict></mc:Fallback></mc:AlternateContent></w:r>`; // Fallback jak w Wordzie (docx-preview bez niego się wywraca)

async function fixture() {
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/></Relationships>');
  z.file("word/media/logo.png", PNG);
  const body = [
    `<w:p>${anchorRun(1)}${textRun("Tytuł pisma z logo przypiętym do akapitu")}</w:p>`,
    `<w:p>${textRun("Przed obrazkiem ")}${inlineRun(2)}${textRun(" za obrazkiem.")}</w:p>`,
    `<w:p>${textbox}${textRun("Akapit z polem tekstowym")}</w:p>`,
    `<w:p><w:r><w:t xml:space="preserve">Tekst i rysunek w jednym fragmencie </w:t><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="457200" cy="228600"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="7" name="Obraz 7"/><wp:cNvGraphicFramePr/>${pic(7)}</wp:inline></w:drawing></w:r></w:p>`,
    `<w:p>${textRun("Zwykły akapit na końcu.")}</w:p>`,
  ].join("");
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS} mc:Ignorable=""><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return z.generateAsync({ type: "base64" });
}

const savedDoc = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  const doc = await z.file("word/document.xml").async("string");
  const ps = doc.match(/<w:body>([\s\S]*)<w:sectPr/)[1].match(/<w:p[ >][\s\S]*?<\/w:p>(?=<w:p[ >]|$)/g) || [];
  // kolejność w akapicie: T = tekst, D = rysunek
  return ps.map((p) => ({ order: (p.match(/<w:t[^>]*>[^<]*<|<w:drawing>/g) || []).map((m) => (m.startsWith("<w:drawing") ? "D" : `T:${m.replace(/<w:t[^>]*>/, "").replace(/<$/, "")}`)).join("|"), drawings: (p.match(/<w:drawing>/g) || []).length, embed: /r:embed="rIdImg"/.test(p) }));
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
  const b64 = await fixture();
  const file = require("path").join(require("os").tmpdir(), `dwb-wyspy-${process.pid}.docx`);
  require("fs").writeFileSync(file, Buffer.from(b64, "base64"));
  await page.setInputFiles("#fileInput", file);
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  require("fs").rmSync(file, { force: true });
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await idle(page);
  const st = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => ({ lock: p.dataset.lock || null, edit: p.isContentEditable, objIsland: !!p.querySelector('[data-cm-kind="obj"]'), mark: !!p.querySelector('[data-cm-kind="obj-mark"]') })));
  check("rysunek przypięty + tekst: akapit edytowalny (znacznik-wyspa)", !st[0].lock && st[0].mark, JSON.stringify(st[0]));
  check("obrazek w zdaniu: akapit edytowalny, sam obraz jest wyspą", !st[1].lock && st[1].objIsland, JSON.stringify(st[1]));
  check("pole tekstowe w rysunku: dalej tylko do odczytu", st[2].lock === "lockObject", JSON.stringify(st[2]));
  check("tekst i rysunek w jednym fragmencie: dalej tylko do odczytu", st[3].lock === "lockObject", JSON.stringify(st[3]));

  const typeAt = async (i, text, atEnd = true) => {
    await page.evaluate(([i, atEnd]) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; placeCaret(p, atEnd ? p.textContent.length : 0); }, [i, atEnd]);
    await page.keyboard.type(text);
    await page.waitForTimeout(300);
  };
  await typeAt(0, " — NOWE");
  await typeAt(1, "Na początku: ", false);
  await typeAt(1, " Koniec.");
  let saved = await savedDoc(page);
  check("pisanie obok przypiętego rysunku: tekst w pliku, rysunek zachowany (ten sam obraz)", saved[0].drawings === 1 && saved[0].embed && /Tytuł pisma z logo przypiętym do akapitu — NOWE/.test(saved[0].order), JSON.stringify(saved[0]));
  check("pisanie przed i za obrazkiem w zdaniu: obraz zostaje między tekstami", /^T:Na początku: Przed obrazkiem \|D\|T: za obrazkiem\. Koniec\.$/.test(saved[1].order), saved[1].order);
  check("zablokowane akapity bez zmian (pole tekstowe, wspólny fragment)", saved[2].drawings === 1 && saved[3].drawings === 1, JSON.stringify([saved[2], saved[3]]));

  // usunięcie obrazka z zdania (zaznacz go i Delete) — znika też z pliku, tekst zostaje
  await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1];
    const w = p.querySelector('[data-cm-kind="obj"]');
    const r = document.createRange();
    r.setStartBefore(w);
    r.setEndAfter(w);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  });
  await page.keyboard.press("Delete");
  await page.waitForTimeout(400);
  saved = await savedDoc(page);
  check("usunięty obrazek w zdaniu znika z pliku, tekst zostaje", saved[1].drawings === 0 && /Przed obrazkiem/.test(saved[1].order) && /za obrazkiem/.test(saved[1].order), saved[1].order);

  // Cofnij przywraca obrazek
  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  await idle(page);
  saved = await savedDoc(page);
  check("Cofnij przywraca usunięty obrazek", saved[1].drawings === 1, saved[1].order);

  await browser.close();
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Rysunki jako wyspy [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
