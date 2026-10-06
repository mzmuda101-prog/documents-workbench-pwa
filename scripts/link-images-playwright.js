// link-images-playwright.js — każdy obraz jako cel linku „Miejsce w dokumencie” (2026-10-06).
//
// Zgłoszenie: 4 obrazy w pliku, a w okienku linku (Wstaw → Link → Miejsce w dokumencie → Obrazy)
// tylko 3. Przyczyna: lista brała PIERWSZY obraz każdego akapitu, a akapit przecięty podziałem
// strony liczyła bez drugiej połówki (data-dwb-cont) — obraz za łamaniem wiersza albo za
// podziałem strony w ogóle nie był celem. Sprawdzane: lista ma każdy obraz; link do drugiego
// obrazu akapitu i do obrazu za podziałem strony ma zakładkę tuż przed TYM obrazem (plik),
// podgląd przy linku pokazuje właściwy obraz, klik przewija do niego; okienko istniejącego
// linku pokazuje wybrany obraz. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const zlib = require("zlib");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(500));

// PNG jednego koloru (bez bibliotek) — każdy obraz inny, żeby było widać, który jest który
function png(w, h, [r, g, b]) {
  const crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (buf) => { let c = 0xffffffff; for (const x of buf) c = crcTable[(c ^ x) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const row = Buffer.concat([Buffer.from([0]), Buffer.from(Array.from({ length: w }, () => [r, g, b]).flat())]);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk("IHDR", ihdr), chunk("IDAT", zlib.deflateSync(Buffer.concat(Array(h).fill(row)))), chunk("IEND", Buffer.alloc(0))]);
}

const COLORS = [[200, 40, 40], [40, 160, 60], [40, 80, 200], [220, 160, 20], [140, 40, 160]];
const drawing = (n) => `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2880000" cy="1440000"/><wp:docPr id="${n}" name="Obraz ${n}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${n}" name="img${n}.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg${n}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2880000" cy="1440000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const t = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;

async function buildDocx(file) {
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>
<w:p>${t("Zobacz ilustrację poniżej.")}</w:p>
<w:p>${drawing(1)}</w:p>
<w:p>${t("Dwa obrazy w jednym akapicie:")}</w:p>
<w:p>${drawing(2)}<w:r><w:br/></w:r>${drawing(3)}</w:p>
<w:p>${drawing(4)}<w:r><w:br w:type="page"/></w:r>${drawing(5)}</w:p>
<w:p>${t("Koniec.")}</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${COLORS.map((_, i) => `<Relationship Id="rIdImg${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/img${i + 1}.png"/>`).join("")}</Relationships>`);
  zip.file("word/document.xml", doc);
  COLORS.forEach((c, i) => zip.file(`word/media/img${i + 1}.png`, png(40, 20, c)));
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

const savedXml = async (page) => {
  const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
  const zip = await JSZip.loadAsync(Buffer.from(b64, "base64"));
  return zip.file("word/document.xml").async("string");
};

// link z akapitu 0 (słowo „ilustrację”) do obrazu o etykiecie `label`
async function linkTo(page, label) {
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0]; const r = formDomRange(p, 7, 17); p.closest(".docx-edit-root").focus({ preventScroll: true }); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await page.evaluate(() => composeUi.openLinkForm(document.getElementById("insertMenuBtn")));
  await page.click(".lf-mode button[data-mode=doc]");
  await page.selectOption(".lf-target", { label });
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
}

// który obraz (0–4, kolejność w pliku) jest celem linku z akapitu 0
const linkedImage = (page) => page.evaluate(() => {
  const a = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0].querySelector("a[href^='#']");
  const hit = a && linkTargetImage(a.getAttribute("href").slice(1));
  return hit ? docImageTargets(document.querySelector(".docx-preview-host")).findIndex((x) => x.img === hit.img) : -2;
});

async function run() {
  const file = path.join(os.tmpdir(), `dwb-link-images-${process.pid}.docx`);
  await buildDocx(file);
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles(file);
  await page.waitForSelector(".docx-preview-host img", { timeout: 20000 });
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);

  const shape = await page.evaluate(() => ({ imgs: document.querySelectorAll(".docx-preview-host img").length, cont: !!document.querySelector(".docx-preview-host p[data-dwb-cont] img") }));
  check("plik: 5 obrazów w podglądzie, ostatni w drugiej połówce akapitu (za podziałem strony)", shape.imgs === 5 && shape.cont, JSON.stringify(shape));

  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0]; const r = formDomRange(p, 7, 17); p.closest(".docx-edit-root").focus({ preventScroll: true }); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await page.evaluate(() => composeUi.openLinkForm(document.getElementById("insertMenuBtn")));
  await page.click(".lf-mode button[data-mode=doc]");
  const opts = await page.evaluate(() => [...document.querySelectorAll(".lf-target optgroup")].find((g) => g.label === "Obrazy")?.querySelectorAll("option").length || 0);
  check("okienko linku: każdy obraz osobno na liście (5, nie 3)", opts === 5, String(opts));
  await page.click(".compose-pop-form .lf-cancel");

  // drugi obraz w tym samym akapicie (Obraz 3)
  await linkTo(page, "Obraz 3");
  let xml = await savedXml(page);
  let bm = (xml.match(/<w:hyperlink w:anchor="([^"]+)"/) || [])[1];
  const before3 = bm ? new RegExp(`<w:bookmarkStart w:id="\\d+" w:name="${bm}"/>${'<w:r><w:drawing>'}[\\s\\S]*?name="img3\\.png"`).test(xml.replace(/ xmlns:\w+="[^"]+"/g, "")) : false;
  check("plik: zakładka tuż przed drugim obrazem akapitu (nie na początku akapitu)", before3, bm);
  check("podgląd: link prowadzi do drugiego obrazu akapitu", (await linkedImage(page)) === 2, String(await linkedImage(page)));

  // okienko istniejącego linku pokazuje wybrany obraz
  await page.evaluate(() => { const a = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0].querySelector("a"); const r = document.createRange(); r.setStart(a.firstChild.firstChild || a.firstChild, 1); r.collapse(true); a.closest(".docx-edit-root").focus({ preventScroll: true }); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await page.evaluate(() => composeUi.openLinkForm(document.getElementById("insertMenuBtn")));
  const selLabel = await page.evaluate(() => document.querySelector(".lf-target").selectedOptions[0]?.textContent);
  check("okienko istniejącego linku: wybrany „Obraz 3”", selLabel === "Obraz 3", selLabel);
  await page.click(".compose-pop-form .lf-cancel");

  // obraz za podziałem strony (Obraz 5) — ten sam link, zmiana celu
  await linkTo(page, "Obraz 5");
  xml = await savedXml(page);
  bm = (xml.match(/<w:hyperlink w:anchor="([^"]+)"/) || [])[1];
  const before5 = bm ? new RegExp(`<w:bookmarkStart w:id="\\d+" w:name="${bm}"/><w:r><w:drawing>[\\s\\S]*?name="img5\\.png"`).test(xml.replace(/ xmlns:\w+="[^"]+"/g, "")) : false;
  check("plik: zakładka tuż przed obrazem za podziałem strony", before5, bm);
  check("podgląd: link prowadzi do obrazu za podziałem strony", (await linkedImage(page)) === 4, String(await linkedImage(page)));
  check("plik: dalej 5 obrazów, jeden link, zakładki sparowane", (xml.match(/<w:drawing>/g) || []).length === 5 && (xml.match(/<w:hyperlink /g) || []).length === 1
    && (xml.match(/<w:bookmarkStart /g) || []).length === (xml.match(/<w:bookmarkEnd /g) || []).length);

  // pierwszy obraz (sam w akapicie) — dalej działa
  await linkTo(page, "Obraz 1");
  check("podgląd: link do obrazu samego w akapicie", (await linkedImage(page)) === 0, String(await linkedImage(page)));

  // czytanie: podgląd przy linku i skok
  await linkTo(page, "Obraz 5");
  await page.evaluate(() => appFrame.setReadOnly(true));
  await idle(page);
  await page.evaluate(() => { docViewportEl.scrollTop = 0; });
  await page.hover(".docx-preview-host a[data-dwb-img-link]");
  await page.waitForSelector(".link-peek", { timeout: 3000 }).catch(() => {});
  const peek = await page.evaluate(() => {
    const el = document.querySelector(".link-peek");
    const want = docImageTargets(document.querySelector(".docx-preview-host"))[4].img;
    return el && { cap: el.querySelector(".link-peek-cap").textContent, same: el.querySelector("img").src === (want.currentSrc || want.src) };
  });
  check("najechanie: podgląd pokazuje obraz za podziałem strony, podpis „Obraz 5”", peek?.same && peek.cap === "Obraz 5", JSON.stringify(peek));
  await page.mouse.move(5, 5);
  await page.click(".docx-preview-host a[data-dwb-img-link]");
  await page.waitForTimeout(1200);
  const seen = await page.evaluate(() => {
    const r = docImageTargets(document.querySelector(".docx-preview-host"))[4].img.getBoundingClientRect();
    const v = docViewportEl.getBoundingClientRect();
    return r.top >= v.top - 1 && r.bottom <= v.bottom + 1;
  });
  check("klik w link: obraz za podziałem strony widoczny w oknie", seen);

  check("brak błędów strony", errors.length === 0, errors.join(" | "));
  await browser.close();
  fs.rmSync(file, { force: true });
}

run().then(() => {
  let failed = 0;
  results.forEach((r) => { if (!r.ok) failed++; console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` — ${r.detail}`}`); });
  console.log(`\n${ENGINE}: ${results.length - failed}/${results.length} OK`);
  process.exit(failed ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
