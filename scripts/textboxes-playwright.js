// textboxes-playwright.js — pola tekstowe jak w Wordzie (textboxes.js, 2026-10-28).
//
// Zgłoszenie: w karcie terenu z PDF nie dało się dopisać napisu z boku mapki. Sprawdzamy:
// Wstaw → Pole tekstowe (pisanie od razu), zapis zgodny z Wordem (wps + kopia VML, ta sama treść),
// przesuwanie myszą z przyciąganiem do marginesu, zmianę rozmiaru uchwytem, format z karty,
// usuwanie i Cofnij, pisanie w akapicie kotwicy (pole zostaje), pole z pliku Worda (dawniej akapit
// tylko do odczytu) — rysowane z tekstem, edycja w obu kopiach, przeniesienie na drugą stronę.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(350));
const mod = process.platform === "darwin" ? "Meta" : "Control";

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const tp = (s) => `<w:p><w:r><w:t xml:space="preserve">${s}</w:t></w:r></w:p>`;
// pole jak z Worda: kotwica względem marginesu (poziomo) i akapitu (pionowo), ramka, kopia VML
const wordBox = `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>3600000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>360000</wp:posOffset></wp:positionV><wp:extent cx="1440000" cy="540000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="Pole tekstowe 2"/><wp:cNvGraphicFramePr><a:graphicFrameLocks/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr><wps:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="1440000" cy="540000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Pole z Worda</w:t></w:r></w:p><w:p><w:r><w:t>druga linia</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="Pole tekstowe 2" o:spid="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:283.45pt;margin-top:28.35pt;width:113.4pt;height:42.5pt;z-index:251659264;mso-position-horizontal-relative:margin;mso-position-vertical-relative:text"><v:textbox><w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Pole z Worda</w:t></w:r></w:p><w:p><w:r><w:t>druga linia</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

async function fixture() {
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const body = [
    `<w:p><w:r><w:t xml:space="preserve">Akapit z polem tekstowym Worda.</w:t></w:r>${wordBox}</w:p>`,
    tp("Drugi akapit pierwszej strony, zwykły tekst do przyciągania linijek."),
    tp("Trzeci akapit."),
    `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`,
    tp("Pierwszy akapit drugiej strony."),
    tp("Drugi akapit drugiej strony."),
  ].join("");
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS} mc:Ignorable="w14"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return z.generateAsync({ type: "nodebuffer" });
}

// pola w zapisanym pliku: tekst w wps i w kopii VML, położenie, wymiary, akapit kotwicy
const savedBoxes = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  const xml = await z.file("word/document.xml").async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const out = [];
  collectParagraphElements(doc.documentElement, "all").forEach((p, index) => {
    tbBoxesOfParagraph(p).forEach((b, nth) => {
      const d = tbDescribe(b);
      const vText = b.vTxbx ? Array.from(b.vTxbx.getElementsByTagNameNS(W_NS, "p")).map((q) => tbParaPlain(q)).join("\n") : null;
      out.push({ index, nth, text: d.text, vText, x: d.h.off, y: d.v.off, relH: d.h.rel, relV: d.v.rel, w: d.w, h: d.hgt, line: !!d.line, fill: !!d.fill, autoFit: d.autoFit, bold: d.paras.every((q) => q.runs.every((r) => r.b)), sz: d.paras[0]?.runs[0]?.sz || null, jc: d.paras[0]?.jc || null, vStyle: b.vShape?.getAttribute("style") || "" });
    });
  });
  return { boxes: out, nsOk: /xmlns:wps=/.test(xml.slice(0, 3000)) && /xmlns:mc=/.test(xml.slice(0, 3000)), paraCount: collectParagraphElements(doc.documentElement, "all").length };
});
const previewBoxes = (page) => page.evaluate(() => [...document.querySelectorAll(".docx-preview-host .dwb-tb")].map((el) => {
  const sec = el.closest("section.docx");
  const secs = [...document.querySelectorAll(".docx-preview-host section.docx")];
  return { text: el.textContent, left: parseFloat(el.style.left), top: parseFloat(el.style.top), w: el.offsetWidth, h: el.offsetHeight, page: secs.indexOf(sec), sel: el.classList.contains("is-selected"), index: +el.dataset.tbIndex, nth: +el.dataset.tbNth };
}));
const boxCenter = (page, i = 0) => page.evaluate((i) => { const r = document.querySelectorAll(".docx-preview-host .dwb-tb")[i].getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2, r: { l: r.left, t: r.top, w: r.width, h: r.height } }; }, i);

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
  const file = require("path").join(require("os").tmpdir(), `dwb-pola-${process.pid}.docx`);
  require("fs").writeFileSync(file, await fixture());
  await page.setInputFiles("#fileInput", file);
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  require("fs").rmSync(file, { force: true });
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await idle(page);

  // 1. pole z pliku Worda: narysowane z tekstem, akapit kotwicy edytowalny, liczba akapitów zgodna
  let pv = await previewBoxes(page);
  const st0 = await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const sec = ps[0].closest("section.docx");
    const cs = getComputedStyle(sec);
    return { lock: ps[0].dataset.lock || null, edit: ps[0].isContentEditable, count: ps.length, padL: parseFloat(cs.paddingLeft), pTop: (ps[0].getBoundingClientRect().top - sec.getBoundingClientRect().top) / (sec.getBoundingClientRect().width / sec.offsetWidth) };
  });
  check("pole z Worda narysowane z tekstem obu akapitów", pv.length === 1 && /Pole z Worda/.test(pv[0].text) && /druga linia/.test(pv[0].text), JSON.stringify(pv));
  check("pole z Worda w miejscu z pliku (margines + 3,6 mln EMU, akapit + 0,36 mln EMU)", pv[0] && Math.abs(pv[0].left - (st0.padL + 3600000 / 9525)) < 1.5 && Math.abs(pv[0].top - (st0.pTop + 360000 / 9525)) < 1.5, JSON.stringify({ pv: pv[0], st0 }));
  check("akapit z polem tekstowym edytowalny (dawniej tylko do odczytu)", !st0.lock && st0.edit, JSON.stringify(st0));
  let saved = await savedBoxes(page);
  check("liczba akapitów podglądu = pliku (pola nie przesuwają numeracji)", saved.paraCount === st0.count, `${saved.paraCount} vs ${st0.count}`);

  // 2. pisanie w akapicie kotwicy — pole zostaje nietknięte
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0]; placeCaret(p, 6); });
  await page.keyboard.type("XYZ ");
  await page.waitForTimeout(300);
  saved = await savedBoxes(page);
  const anchorText = await page.evaluate(async () => { const z = await JSZip.loadAsync(await buildDocumentForSave()); const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml"); return getParagraphText(collectParagraphElements(d.documentElement, "all")[0]); });
  check("pisanie w akapicie kotwicy: tekst w pliku, pole bez zmian (obie kopie)", /AkapitXYZ /.test(anchorText.replace(/\s+/g, (m) => m)) && saved.boxes.length === 1 && saved.boxes[0].text === "Pole z Worda\ndruga linia" && saved.boxes[0].vText === "Pole z Worda\ndruga linia", JSON.stringify({ anchorText, b: saved.boxes[0] }));

  // 3. edycja tekstu pola z Worda: dwuklik, nowy tekst, Esc → obie kopie, pogrubienie 1. linii zostaje
  let c = await boxCenter(page);
  await page.mouse.dblclick(c.x, c.y);
  await page.waitForTimeout(250);
  const hasEditor = await page.evaluate(() => document.activeElement?.classList.contains("dwb-tb-editor"));
  check("dwuklik w pole: pisanie w polu (edytor ma fokus)", hasEditor);
  await page.keyboard.press(`${mod}+a`);
  await page.keyboard.type("Pole z Worda\nnowa linia");
  await page.keyboard.press("Escape");
  await idle(page);
  saved = await savedBoxes(page);
  check("edycja pola z Worda: nowy tekst w wps i w kopii VML", saved.boxes[0]?.text === "Pole z Worda\nnowa linia" && saved.boxes[0]?.vText === "Pole z Worda\nnowa linia", JSON.stringify(saved.boxes[0]));
  const firstBold = await page.evaluate(async () => { const z = await JSZip.loadAsync(await buildDocumentForSave()); const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml"); const b = tbBoxesOfParagraph(collectParagraphElements(d.documentElement, "all")[0])[0]; return tbDescribe(b).paras.map((p) => p.runs.map((r) => r.b)); });
  check("niezmieniona linia zachowuje format (pogrubienie), zmieniona bierze format z miejsca", JSON.stringify(firstBold) === "[[true],[false]]", JSON.stringify(firstBold));
  pv = await previewBoxes(page);
  check("po edycji pole dalej zaznaczone (karta narzędzi)", pv[0]?.sel && await page.evaluate(() => !!document.querySelector(".dwb-tb-card")), JSON.stringify(pv));

  // 4. Wstaw → Pole tekstowe przy kursorze, pisanie od razu
  await page.evaluate(() => { dwbTextbox.deselect(); const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[2]; placeCaret(p, 0); });
  await page.evaluate(() => dwbTextbox.insert());
  await idle(page);
  await page.waitForFunction(() => document.activeElement?.classList.contains("dwb-tb-editor"), null, { timeout: 5000 }).catch(() => {});
  await page.keyboard.type("ok 34 m.");
  await page.keyboard.press("Escape");
  await idle(page);
  saved = await savedBoxes(page);
  const nb = saved.boxes.find((b) => b.text === "ok 34 m.");
  check("nowe pole: tekst w pliku (wps + VML), względem strony, bez ramki i tła, dopasowane do tekstu", nb && nb.vText === "ok 34 m." && nb.relH === "page" && nb.relV === "page" && !nb.line && !nb.fill && nb.autoFit, JSON.stringify(nb));
  check("przestrzenie nazw na korzeniu dokumentu (wps, mc)", saved.nsOk);
  pv = await previewBoxes(page);
  const nbv = pv.find((b) => b.text === "ok 34 m.");
  check("nowe pole w podglądzie tam, gdzie w pliku", nbv && nb && Math.abs(nbv.left - nb.x) < 1 && Math.abs(nbv.top - nb.y) < 1, JSON.stringify({ nbv, nb }));

  // 5. przeciąganie z przyciąganiem do lewego marginesu
  const ni = pv.indexOf(nbv);
  c = await boxCenter(page, ni);
  const marg = await page.evaluate(() => { const sec = document.querySelector(".docx-preview-host section.docx"); return { padL: parseFloat(getComputedStyle(sec).paddingLeft), s: sec.getBoundingClientRect().width / sec.offsetWidth, left: sec.getBoundingClientRect().left }; });
  // cel: lewa krawędź pola 3 px (ekranu) od marginesu → przyciąga dokładnie na margines
  const targetLeft = marg.left + marg.padL * marg.s + 3;
  await page.mouse.move(c.x, c.y);
  await page.mouse.down();
  await page.mouse.move(c.x - 20, c.y + 5, { steps: 3 });
  await page.mouse.move(targetLeft + c.r.w / 2, c.y + 40, { steps: 6 });
  const guides = await page.evaluate(() => document.querySelectorAll(".dwb-tb-guide").length);
  await page.mouse.up();
  await idle(page);
  saved = await savedBoxes(page);
  const moved = saved.boxes.find((b) => b.text === "ok 34 m.");
  check("przeciąganie: linia pomocnicza widoczna w trakcie", guides > 0, `linii: ${guides}`);
  // krawędź pola albo początek jego tekstu (po wewnętrznym marginesie 0,1") na marginesie strony
  check("przeciąganie: pole przyciągnięte do lewego marginesu (w pliku)", moved && (Math.abs(moved.x - marg.padL) < 0.05 || Math.abs(moved.x + 9.6 - marg.padL) < 0.05), JSON.stringify({ moved, padL: marg.padL }));
  check("przeciąganie: kopia VML też przesunięta", moved && new RegExp(`margin-left:${Math.round(moved.x * 0.75 * 100) / 100}pt`).test(moved.vStyle), moved?.vStyle);

  // 6. zmiana szerokości uchwytem „e”
  const w0 = moved.w;
  const hE = await page.evaluate((txt) => { const el = [...document.querySelectorAll(".dwb-tb")].find((b) => b.textContent === txt); const r = el.querySelector('.dwb-tb-h[data-h="e"]')?.getBoundingClientRect(); return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; }, "ok 34 m.");
  check("zaznaczone pole ma uchwyty", !!hE);
  if (hE) {
    await page.mouse.move(hE.x, hE.y);
    await page.mouse.down();
    await page.mouse.move(hE.x + 30, hE.y, { steps: 3 });
    await page.mouse.move(hE.x + 61, hE.y + 1, { steps: 4 });
    await page.mouse.up();
    await idle(page);
    saved = await savedBoxes(page);
    const rs = saved.boxes.find((b) => b.text === "ok 34 m.");
    check("uchwyt: pole szersze (w pliku), dalej dopasowane do tekstu", rs && rs.w > w0 + 40 && rs.autoFit, JSON.stringify({ w0, rs }));
  }

  // 7. format z karty: większa czcionka, pogrubienie, do środka, ramka, tło
  const clickCard = async (cls) => { await page.click(`.dwb-tb-card .${cls}`); await idle(page); };
  await clickCard("tbc-bigger");
  await clickCard("tbc-bold");
  await page.click('.dwb-tb-card [data-align="center"]'); await idle(page);
  await clickCard("tbc-border");
  await clickCard("tbc-fill");
  saved = await savedBoxes(page);
  const fm = saved.boxes.find((b) => b.text === "ok 34 m.");
  check("karta: rozmiar w górę, pogrubienie, środek, ramka, tło — w pliku", fm && fm.sz > 20 && fm.bold && fm.jc === "center" && fm.line && fm.fill && /stroked="t"|strok/.test(JSON.stringify(fm)) !== null, JSON.stringify(fm));

  // 8. Delete usuwa pole, Cofnij przywraca
  const before = saved.boxes.length;
  await page.evaluate(() => document.querySelector(".dwb-tb.is-selected")?.focus());
  await page.keyboard.press("Delete");
  await idle(page);
  saved = await savedBoxes(page);
  check("Delete przy zaznaczonym polu usuwa je z pliku", saved.boxes.length === before - 1, `${before} → ${saved.boxes.length}`);
  await page.evaluate(() => document.querySelector(".docx-edit-root")?.focus());
  await page.keyboard.press(`${mod}+z`);
  await idle(page);
  saved = await savedBoxes(page);
  check("Cofnij przywraca usunięte pole", saved.boxes.length === before, `${saved.boxes.length}`);

  // 8b. linia bazowa (jak Affinity): pole z większą czcionką staje spodem liter na spodzie liter
  // wiersza obok; lepkość: po przyciągnięciu ruch o 9 px (tolerancja 6) dalej trzyma linię
  const baseOf = () => page.evaluate(() => {
    const desc = (el) => { const cs = getComputedStyle(el); const c = document.createElement("canvas").getContext("2d"); c.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`; return c.measureText("Hg").fontBoundingBoxDescent || parseFloat(cs.fontSize) * 0.21; };
    const sec = document.querySelector(".docx-preview-host section.docx");
    const s = sec.getBoundingClientRect().width / sec.offsetWidth;
    const lineText = [...sec.querySelectorAll("article p")].find((p) => /Drugi akapit pierwszej/.test(p.textContent));
    const tn = document.createTreeWalker(lineText, NodeFilter.SHOW_TEXT).nextNode();
    const r1 = document.createRange(); r1.selectNodeContents(tn); const lr = [...r1.getClientRects()][0];
    const box = [...document.querySelectorAll(".dwb-tb")].find((b) => b.textContent === "ok 34 m.");
    const bt = document.createTreeWalker(box, NodeFilter.SHOW_TEXT).nextNode();
    const r2 = document.createRange(); r2.selectNodeContents(bt); const br = r2.getClientRects()[0];
    const bb = box.getBoundingClientRect();
    return { lineBase: lr.bottom - desc(tn.parentElement) * s, lineRight: lr.right, boxBase: br.bottom - desc(bt.parentElement) * s, boxL: bb.left, boxT: bb.top, boxW: bb.width, boxH: bb.height, lineSize: getComputedStyle(tn.parentElement).fontSize, boxSize: getComputedStyle(bt.parentElement).fontSize };
  });
  await page.evaluate(() => { dwbTextbox.deselect(); document.querySelector(".docx-preview-host section.docx").scrollIntoView({ block: "start" }); });
  await page.waitForTimeout(300);
  let bl = await baseOf();
  const grab = { x: bl.boxL + bl.boxW / 2, y: bl.boxT + bl.boxH / 2 };
  const off = { x: grab.x - bl.boxL, base: bl.boxBase - grab.y };
  // cel: lewa krawędź 60 px za końcem wiersza, linia bazowa 3 px pod linią bazową wiersza
  const tx = bl.lineRight + 60 + off.x, ty = bl.lineBase + 3 - off.base;
  await page.mouse.move(grab.x, grab.y);
  await page.mouse.down();
  await page.mouse.move(grab.x + 15, grab.y + 15, { steps: 3 });
  await page.mouse.move(tx, ty, { steps: 10 });
  let mid = await baseOf();
  check("linia bazowa: większa czcionka pola stoi spodem liter na wierszu obok (w trakcie)", bl.boxSize !== bl.lineSize && Math.abs(mid.boxBase - mid.lineBase) < 0.75, JSON.stringify({ mid, rozmiary: [bl.boxSize, bl.lineSize] }));
  await page.mouse.move(tx, ty, { steps: 6 });
  await page.mouse.up();
  await idle(page);
  bl = await baseOf();
  check("po upuszczeniu (zapis i przerysowanie) linia bazowa dalej równa", Math.abs(bl.boxBase - bl.lineBase) < 0.75, JSON.stringify(bl));
  await page.evaluate(() => dwbTextbox.deselect()); // karta pola nie zasłania kolejnych kroków

  // 8b2. uchwyt: ta sama szerokość co inne pole (pole z Worda) — gdy krawędź nie trafia na linię
  {
    const info = await page.evaluate(() => {
      const ok = [...document.querySelectorAll(".dwb-tb")].find((b) => b.textContent === "ok 34 m.");
      const wb = [...document.querySelectorAll(".dwb-tb")].find((b) => /Pole z Worda/.test(b.textContent));
      ok.scrollIntoView({ block: "center" });
      const r = ok.getBoundingClientRect();
      return { okW: r.width, wbW: wb.getBoundingClientRect().width, l: r.left, y: r.top + r.height / 2 };
    });
    await page.waitForTimeout(200);
    const c0 = await boxCenter(page, (await previewBoxes(page)).findIndex((b) => b.text === "ok 34 m."));
    await page.mouse.click(c0.x, c0.y);
    await page.waitForTimeout(250);
    const he = await page.evaluate(() => { const el = document.querySelector(".dwb-tb.is-selected"); const r = el?.querySelector('.dwb-tb-h[data-h="w"]')?.getBoundingClientRect(); return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2, right: el.getBoundingClientRect().right } : null; });
    if (he) {
      const target = he.right - info.wbW - 3; // lewy uchwyt: 3 px ponad szerokość pola z Worda
      await page.mouse.move(he.x, he.y);
      await page.mouse.down();
      await page.mouse.move(he.x + 20, he.y, { steps: 3 });
      await page.mouse.move(target, he.y, { steps: 8 });
      const wNow = await page.evaluate(() => document.querySelector(".dwb-tb.is-selected").getBoundingClientRect().width);
      await page.mouse.up();
      await idle(page);
      check("uchwyt: ta sama szerokość co pole z Worda", Math.abs(wNow - info.wbW) < 0.75, JSON.stringify({ wNow, wbW: info.wbW }));
    } else check("uchwyt: ta sama szerokość — uchwyt dostępny", false, "");
    await page.evaluate(() => dwbTextbox.deselect());
  }

  // 8c. reguły przyciągania (jak Figma / Affinity): bliższe w tolerancji wygrywa, trzymana linia
  // puszcza dopiero po 1,8 × tolerancji, gdy nic innego nie jest blisko; dalekie obiekty pomijane
  const rules = await page.evaluate(() => {
    const sa = dwbTextbox._snapAxis;
    const box = { L: 100, R: 200, T: 100, B: 120 };
    const near = { v: 104, src: { L: 210, R: 300, T: 90, B: 130 } };
    const far = { v: 102, src: { L: 900, R: 950, T: 90, B: 130 } };
    const a = sa([{ f: 100, id: "t" }], [near, far], 6, box, 300, null);
    const held = { id: "t", v: 104, src: near.src };
    const b = sa([{ f: 95, id: "t" }], [], 6, box, 300, held); // 9 od trzymanej linii, nic innego
    const c = sa([{ f: 92, id: "t" }], [], 6, box, 300, held); // 12 — za daleko
    const d = sa([{ f: 95, id: "t" }], [{ v: 97, src: near.src }], 6, box, 300, held); // inna linia bliżej
    return { a: a && a.v, b: b && [b.v, !!b.held], c, d: d && [d.v, !!d.held] };
  });
  check("przyciąganie: daleki obiekt pominięty, bliski wygrywa", rules.a === 104, JSON.stringify(rules));
  check("lepkość: bez konkurencji trzyma linię do 1,8 × tolerancji, potem puszcza", JSON.stringify(rules.b) === "[104,true]" && rules.c === null, JSON.stringify(rules));
  check("lepkość nie blokuje bliższej linii w tolerancji", JSON.stringify(rules.d) === "[97,false]", JSON.stringify(rules));

  // 9. przeniesienie pola z Worda na drugą stronę: kotwica w akapicie drugiej strony
  pv = await previewBoxes(page);
  const wi = pv.findIndex((b) => /Pole z Worda/.test(b.text));
  c = await boxCenter(page, wi);
  const p2 = await page.evaluate(() => { const secs = document.querySelectorAll(".docx-preview-host section.docx"); const r = secs[1]?.getBoundingClientRect(); return r ? { x: r.left + r.width / 2, y: r.top + 260 } : null; });
  check("dokument ma drugą stronę", !!p2);
  if (p2) {
    await page.evaluate((y) => { const vp = docViewportEl; vp.scrollTop += 0; }, 0);
    await page.mouse.move(c.x, c.y);
    await page.mouse.down();
    await page.mouse.move(c.x + 10, c.y + 10, { steps: 2 });
    await page.mouse.move(p2.x, p2.y, { steps: 12 });
    await page.mouse.up();
    await idle(page);
    saved = await savedBoxes(page);
    const wb = saved.boxes.find((b) => /Pole z Worda/.test(b.text));
    pv = await previewBoxes(page);
    const wv = pv.find((b) => /Pole z Worda/.test(b.text));
    check("pole na drugiej stronie: kotwica w akapicie 2. strony, położenie względem strony", wb && wb.index >= 4 && wb.relH === "page" && wb.relV === "page", JSON.stringify(wb));
    check("pole na drugiej stronie w podglądzie po przerysowaniu", wv?.page === 1, JSON.stringify(wv));
  }

  // 10. zapis i ponowne otwarcie: pola w tych samych miejscach
  saved = await savedBoxes(page);
  const bytes = await page.evaluate(async () => Array.from(await buildDocumentForSave()));
  const f2 = require("path").join(require("os").tmpdir(), `dwb-pola2-${process.pid}.docx`);
  require("fs").writeFileSync(f2, Buffer.from(bytes));
  if (process.env.OUT) require("fs").writeFileSync(process.env.OUT, Buffer.from(bytes)); // podgląd w Wordzie / walidator
  await page.setInputFiles("#fileInput", f2);
  await page.waitForTimeout(800);
  await idle(page);
  require("fs").rmSync(f2, { force: true });
  const pv2 = await previewBoxes(page);
  check("po ponownym otwarciu: te same pola w tych samych miejscach", pv2.length === saved.boxes.length && saved.boxes.every((b) => pv2.some((q) => q.text.replace(/\s/g, "") === b.text.replace(/\s/g, "") && Math.abs(q.left - b.x) < 1 && Math.abs(q.top - b.y) < 1)), JSON.stringify({ pv2, s: saved.boxes }));

  // 11. Czytanie: pola widoczne, bez uchwytów i bez reakcji na klik
  await page.evaluate(() => { readModeEl.checked = true; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForTimeout(400);
  c = await boxCenter(page, 0);
  await page.mouse.click(c.x, c.y);
  await page.waitForTimeout(200);
  check("w Czytaniu klik w pole nic nie zaznacza", await page.evaluate(() => !document.querySelector(".dwb-tb.is-selected") && !document.querySelector(".dwb-tb-card")));

  await browser.close();
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Pola tekstowe [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => {
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  console.error(e);
  process.exit(1);
});
