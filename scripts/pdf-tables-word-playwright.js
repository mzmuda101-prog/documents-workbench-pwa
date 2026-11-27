// pdf-tables-word-playwright.js — PDF → DOCX: tabele bez pionowych linii (jak w prawdziwym Wordzie).
// Sprawdza wzorce porównane z Wordem na macOS:
//   1) tylko poziome linie (nowoczesna tabela Worda, 5×3, bez obrazu linii w tle),
//   2) styl akademicki / booktabs (góra + linia pod nagłówkiem + dół, 5×3),
//   3) poziome linie + zawinięty opis w komórce (3×3, kontynuacja w tej samej komórce),
//   4) cieniowany nagłówek + poziome linie poniżej (5×3 jako jedna tabela, nagłówek podzielony na kolumny),
//   5) zebra bez pionowych linii (5×3 i parzysta 4×3 z końcowym białym wierszem, jedna tabela z cieniowaniem),
//   6) sama ramka zewnętrzna bez linii wewnętrznych (5×3),
//   7) strażniki: spis treści z kropkami oraz dwie ozdobne linie wokół cytatu NIE stają się tabelą,
//   8) edycja komórki + zapis + walidacja OOXML + widok telefonu 390 px bez przewijania w bok.
//   node scripts/pdf-tables-word-playwright.js / ENGINE=webkit node scripts/pdf-tables-word-playwright.js

const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const ooxml = require("./ooxml-validate");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "dwb-pdf-tblw-"));

function makePdf(name, content) {
  const objs = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.2 841.8] /Resources << /Font << /F1 4 0 R /F2 5 0 R >> >> /Contents 6 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>",
    `<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}\nendstream`,
  ];
  let out = "%PDF-1.4\n";
  const offs = [];
  objs.forEach((o, i) => { offs.push(Buffer.byteLength(out)); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const xref = Buffer.byteLength(out);
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n${offs.map((o) => `${String(o).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  const file = path.join(TMP, name);
  fs.writeFileSync(file, out);
  return file;
}

const txt = (x, y, s, f = "F1") => `BT /${f} 11 Tf 0 0 0 rg ${x} ${y} Td (${s}) Tj ET`;
const hline = (x0, x1, y, w = 0.75) => `0 0 0 RG ${w} w ${x0} ${y} m ${x1} ${y} l S`;
const vline = (x, y0, y1, w = 0.75) => `0 0 0 RG ${w} w ${x} ${y0} m ${x} ${y1} l S`;
const rect = (x, y, w, h, g = 0.88) => `${g} ${g} ${g} rg ${x} ${y} ${w} ${h} re f`;

const ROWS = [
  ["Item", "Quantity", "Price"],
  ["Notebook", "12", "45.00"],
  ["Keyboard", "4", "180.00"],
  ["Monitor", "2", "920.00"],
  ["Mouse", "15", "65.00"],
];
const XS = [60, 260, 410];
const YS = [700, 676, 652, 628, 604];
const drawRows = (arr, cols, yArr) => arr.flatMap((r, ri) => r.map((s, ci) => txt(cols[ci], yArr[ri], s, ri === 0 ? "F2" : "F1")));

async function convertAndInspect(page, file, marker) {
  await page.evaluate(() => {
    const s = document.querySelector(".docx-preview-host section");
    if (s) s.setAttribute("data-prev-doc", "1");
  });
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction((m) => document.getElementById("loadingOverlay")?.classList.contains("hidden")
    && !!document.querySelector(".docx-preview-host section:not([data-prev-doc])")
    && document.querySelector(".docx-preview-host").textContent.includes(m), marker, { timeout: 30000 });
  await page.waitForTimeout(150);
  return page.evaluate(async () => {
    const bytes = new Uint8Array(await buildDocumentForSave());
    const z = await JSZip.loadAsync(bytes);
    const xml = await z.file("word/document.xml").async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const tbls = [...doc.getElementsByTagNameNS(NS, "tbl")].map((t) => {
      const trs = [...t.children].filter((e) => e.localName === "tr");
      const rows = trs.map((tr) => [...tr.children].filter((e) => e.localName === "tc").map((tc) => {
        const s = new XMLSerializer().serializeToString(tc);
        return {
          text: [...tc.getElementsByTagNameNS(NS, "t")].map((x) => x.textContent).join(" ").replace(/\s+/g, " ").trim(),
          shd: /<w:shd\b[^>]*w:fill="([^"]+)"/.exec(s)?.[1] || null,
          top: /<w:top\b[^>]*w:val="single"/.test(s),
          bottom: /<w:bottom\b[^>]*w:val="single"/.test(s),
          left: /<w:left\b[^>]*w:val="single"/.test(s),
          right: /<w:right\b[^>]*w:val="single"/.test(s),
        };
      }));
      return { nR: rows.length, nC: Math.max(0, ...rows.map((r) => r.length)), rows };
    });
    const imgs = doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "anchor").length
      + doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "inline").length;
    return { tbls, imgs, bytes: Array.from(bytes) };
  });
}

(async () => {
  const fHoriz = makePdf("horiz.pdf", [
    txt(60, 750, "Horizontal Lines Only", "F2"),
    ...drawRows(ROWS, XS, YS),
    ...[716, 692, 668, 644, 620, 596].map((y) => hline(54, 510, y)),
  ].join("\n"));

  const fBooktabs = makePdf("booktabs.pdf", [
    txt(60, 750, "Academic Booktabs", "F2"),
    ...drawRows(ROWS, XS, YS),
    hline(54, 510, 716, 1.2),
    hline(54, 510, 692, 0.75),
    hline(54, 510, 596, 1.2),
  ].join("\n"));

  const fMulti = makePdf("multiline.pdf", [
    txt(60, 750, "Multiline Horizontal", "F2"),
    hline(54, 510, 718),
    txt(60, 702, "Product", "F2"), txt(250, 702, "Description", "F2"), txt(450, 702, "Price", "F2"),
    hline(54, 510, 694),
    txt(60, 678, "Widget A"), txt(250, 678, "Compact aluminum enclosure"), txt(450, 678, "120.00"),
    txt(250, 662, "with waterproof seal and clip"),
    hline(54, 510, 652),
    txt(60, 636, "Widget B"), txt(250, 636, "Standard plastic housing"), txt(450, 636, "45.00"),
    hline(54, 510, 626),
  ].join("\n"));

  const fShadedHead = makePdf("shaded-head.pdf", [
    txt(60, 750, "Shaded Header Plus Rules", "F2"),
    rect(54, 692, 456, 24, 0.82),
    ...drawRows(ROWS, XS, YS),
    ...[692, 668, 644, 620, 596].map((y) => hline(54, 510, y)),
  ].join("\n"));

  const fZebraOdd = makePdf("zebra-odd.pdf", [
    txt(60, 750, "Zebra Odd", "F2"),
    rect(54, 692, 456, 24, 0.80),
    rect(54, 644, 456, 24, 0.92),
    rect(54, 596, 456, 24, 0.92),
    ...drawRows(ROWS, XS, YS),
  ].join("\n"));

  const fZebraEven = makePdf("zebra-even.pdf", [
    txt(60, 750, "Zebra Even", "F2"),
    rect(54, 692, 456, 24, 0.80),
    rect(54, 644, 456, 24, 0.92),
    ...drawRows(ROWS.slice(0, 4), XS, YS.slice(0, 4)),
  ].join("\n"));

  const fOuterBox = makePdf("outer-box.pdf", [
    txt(60, 750, "Outer Box Only", "F2"),
    ...drawRows(ROWS, XS, YS),
    hline(54, 510, 716), hline(54, 510, 596),
    vline(54, 596, 716), vline(510, 596, 716),
  ].join("\n"));

  const fToc = makePdf("toc.pdf", [
    txt(60, 750, "Table of Contents", "F2"),
    hline(54, 510, 738),
    txt(60, 710, "1. Introduction ........................................ 3"),
    txt(60, 690, "2. Methodology ......................................... 7"),
    txt(60, 670, "3. Conclusion ......................................... 21"),
    hline(54, 510, 656),
  ].join("\n"));

  const fQuote = makePdf("quote.pdf", [
    txt(60, 750, "Regular Article", "F2"),
    hline(54, 510, 720),
    txt(100, 700, "A single centered pull quote across the page"),
    txt(120, 682, "continues cleanly without any columns here."),
    hline(54, 510, 668),
  ].join("\n"));

  const browser = await pw[ENGINE].launch({ headless: true });
  const errors = [];
  const docxBuffers = [];
  try {
    const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
    await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
    const page = await context.newPage();
    page.on("pageerror", (e) => errors.push(e.message));
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());

    // 1. Tylko poziome linie
    const r1 = await convertAndInspect(page, fHoriz, "Horizontal Lines Only");
    docxBuffers.push(["horiz.docx", Buffer.from(r1.bytes)]);
    check("tylko poziome linie: tabela 5×3, bez obrazu linii w tle, bez pionowych ramek",
      r1.tbls.length === 1 && r1.tbls[0].nR === 5 && r1.tbls[0].nC === 3 && r1.imgs === 0
      && r1.tbls[0].rows[0][0].text === "Item" && r1.tbls[0].rows[4][2].text === "65.00"
      && r1.tbls[0].rows[0][0].top && r1.tbls[0].rows[0][0].bottom && !r1.tbls[0].rows[0][0].left,
      JSON.stringify({ n: r1.tbls.length, dim: r1.tbls[0] && `${r1.tbls[0].nR}x${r1.tbls[0].nC}`, imgs: r1.imgs, r0: r1.tbls[0]?.rows[0] }));

    // 2. Styl akademicki (booktabs)
    const r2 = await convertAndInspect(page, fBooktabs, "Academic Booktabs");
    docxBuffers.push(["booktabs.docx", Buffer.from(r2.bytes)]);
    check("booktabs (góra + pod nagłówkiem + dół): tabela 5×3, bez obrazu linii, ramka pod nagłówkiem i bez wewnętrznych",
      r2.tbls.length === 1 && r2.tbls[0].nR === 5 && r2.tbls[0].nC === 3 && r2.imgs === 0
      && r2.tbls[0].rows[0][0].top && r2.tbls[0].rows[0][0].bottom && !r2.tbls[0].rows[1][0].bottom && r2.tbls[0].rows[4][0].bottom
      && r2.tbls[0].rows[2][0].text === "Keyboard" && r2.tbls[0].rows[2][2].text === "180.00",
      JSON.stringify({ n: r2.tbls.length, dim: r2.tbls[0] && `${r2.tbls[0].nR}x${r2.tbls[0].nC}`, imgs: r2.imgs }));

    // 3. Poziome linie + wielowierszowy opis
    const r3 = await convertAndInspect(page, fMulti, "Multiline Horizontal");
    docxBuffers.push(["multiline.docx", Buffer.from(r3.bytes)]);
    check("poziome linie + zawinięty opis: tabela 3×3, obie linijki opisu w komórce (1,1)",
      r3.tbls.length === 1 && r3.tbls[0].nR === 3 && r3.tbls[0].nC === 3 && r3.imgs === 0
      && /Compact aluminum enclosure/.test(r3.tbls[0].rows[1][1].text) && /waterproof seal/.test(r3.tbls[0].rows[1][1].text)
      && r3.tbls[0].rows[2][0].text === "Widget B",
      JSON.stringify({ n: r3.tbls.length, dim: r3.tbls[0] && `${r3.tbls[0].nR}x${r3.tbls[0].nC}`, cell11: r3.tbls[0]?.rows[1]?.[1]?.text }));

    // 4. Cieniowany nagłówek + poziome linie
    const r4 = await convertAndInspect(page, fShadedHead, "Shaded Header Plus Rules");
    docxBuffers.push(["shaded-head.docx", Buffer.from(r4.bytes)]);
    check("cieniowany nagłówek + poziome linie: jedna tabela 5×3, nagłówek podzielony na 3 cieniowane komórki",
      r4.tbls.length === 1 && r4.tbls[0].nR === 5 && r4.tbls[0].nC === 3 && r4.imgs === 0
      && r4.tbls[0].rows[0].length === 3 && r4.tbls[0].rows[0].every((c) => !!c.shd)
      && r4.tbls[0].rows[0][2].text === "Price" && r4.tbls[0].rows[3][0].text === "Monitor",
      JSON.stringify({ n: r4.tbls.length, dim: r4.tbls[0] && `${r4.tbls[0].nR}x${r4.tbls[0].nC}`, r0: r4.tbls[0]?.rows[0] }));

    // 5. Zebra (nieparzysta 5×3 i parzysta 4×3)
    const r5a = await convertAndInspect(page, fZebraOdd, "Zebra Odd");
    docxBuffers.push(["zebra-odd.docx", Buffer.from(r5a.bytes)]);
    check("zebra 5×3 (bez pionowych linii): jedna tabela z naprzemiennym cieniowaniem",
      r5a.tbls.length === 1 && r5a.tbls[0].nR === 5 && r5a.tbls[0].nC === 3 && r5a.imgs === 0
      && !!r5a.tbls[0].rows[0][0].shd && !r5a.tbls[0].rows[1][0].shd && !!r5a.tbls[0].rows[2][0].shd && !r5a.tbls[0].rows[3][0].shd && !!r5a.tbls[0].rows[4][0].shd
      && r5a.tbls[0].rows[1][0].text === "Notebook" && r5a.tbls[0].rows[3][2].text === "920.00",
      JSON.stringify({ n: r5a.tbls.length, dim: r5a.tbls[0] && `${r5a.tbls[0].nR}x${r5a.tbls[0].nC}`, shds: r5a.tbls[0]?.rows.map((r) => !!r[0].shd) }));

    const r5b = await convertAndInspect(page, fZebraEven, "Zebra Even");
    docxBuffers.push(["zebra-even.docx", Buffer.from(r5b.bytes)]);
    check("zebra 4×3 (ostatni wiersz biały): jedna tabela 4×3 obejmująca też dolny biały wiersz",
      r5b.tbls.length === 1 && r5b.tbls[0].nR === 4 && r5b.tbls[0].nC === 3 && r5b.imgs === 0
      && r5b.tbls[0].rows[3][0].text === "Monitor" && !r5b.tbls[0].rows[3][0].shd,
      JSON.stringify({ n: r5b.tbls.length, dim: r5b.tbls[0] && `${r5b.tbls[0].nR}x${r5b.tbls[0].nC}`, r3: r5b.tbls[0]?.rows[3] }));

    // 6. Sama ramka zewnętrzna
    const r6 = await convertAndInspect(page, fOuterBox, "Outer Box Only");
    docxBuffers.push(["outer-box.docx", Buffer.from(r6.bytes)]);
    check("sama ramka zewnętrzna: tabela 5×3 z zewnętrznymi ramkami i bez wewnętrznych",
      r6.tbls.length === 1 && r6.tbls[0].nR === 5 && r6.tbls[0].nC === 3 && r6.imgs === 0
      && r6.tbls[0].rows[0][0].top && r6.tbls[0].rows[0][0].left && !r6.tbls[0].rows[0][0].right && !r6.tbls[0].rows[0][0].bottom
      && r6.tbls[0].rows[4][2].bottom && r6.tbls[0].rows[4][2].right,
      JSON.stringify({ n: r6.tbls.length, dim: r6.tbls[0] && `${r6.tbls[0].nR}x${r6.tbls[0].nC}`, c00: r6.tbls[0]?.rows[0]?.[0], c42: r6.tbls[0]?.rows[4]?.[2] }));

    // 7. Strażniki: TOC z kropkami oraz dwie linie wokół cytatu NIE stają się tabelą
    const r7 = await convertAndInspect(page, fToc, "Table of Contents");
    check("strażnik: spis treści z kropkami między dwiema liniami NIE staje się tabelą", r7.tbls.length === 0, JSON.stringify(r7.tbls));
    const r8 = await convertAndInspect(page, fQuote, "Regular Article");
    check("strażnik: cytat między dwiema liniami bez kolumn NIE staje się tabelą", r8.tbls.length === 0, JSON.stringify(r8.tbls));

    // 8. Edycja komórki + zapis + ponowne otwarcie + telefon 390 px
    await convertAndInspect(page, fHoriz, "Horizontal Lines Only");
    await page.evaluate(() => appFrame.setReadOnly(false));
    await page.waitForFunction(() => !!document.querySelector(".docx-preview-host td p.docx-editable-p"), null, { timeout: 10000 });
    await page.evaluate(() => {
      const p = [...document.querySelectorAll(".docx-preview-host td p.docx-editable-p")].find((x) => x.textContent.includes("Notebook"));
      placeCaret(p, p.textContent.length);
    });
    await page.keyboard.type(" Pro");
    await page.waitForTimeout(400);
    const saved = await page.evaluate(async () => Array.from(new Uint8Array(await buildDocumentForSave())));
    const savedPath = path.join(TMP, "horiz-edited.docx");
    fs.writeFileSync(savedPath, Buffer.from(saved));
    docxBuffers.push(["horiz-edited.docx", Buffer.from(saved)]);
    await page.evaluate(() => {
      const s = document.querySelector(".docx-preview-host section");
      if (s) s.setAttribute("data-prev-doc", "1");
    });
    await page.setInputFiles("#fileInput", savedPath);
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden")
      && !!document.querySelector(".docx-preview-host section:not([data-prev-doc])")
      && document.querySelector(".docx-preview-host")?.textContent.includes("Notebook Pro"), null, { timeout: 20000 });
    const reopened = await page.evaluate(() => ({
      rows: document.querySelectorAll(".docx-preview-host table tr").length,
      hasEdit: document.querySelector(".docx-preview-host")?.textContent.includes("Notebook Pro"),
    }));
    check("edycja komórki tabeli z PDF → zapis → ponowne otwarcie zachowuje układ i zmianę", reopened.rows === 5 && reopened.hasEdit, JSON.stringify(reopened));

    await page.setViewportSize({ width: 390, height: 844 });
    await page.waitForTimeout(400);
    const mob = await page.evaluate(() => ({ sw: document.documentElement.scrollWidth, vw: window.innerWidth, rows: document.querySelectorAll(".docx-preview-host table tr").length }));
    check("telefon 390 px: tabela mieści się bez przewijania całej strony w bok", mob.sw <= mob.vw + 1 && mob.rows === 5, JSON.stringify(mob));

    await context.close();
  } finally {
    await browser.close();
  }

  // Walidacja OOXML dla wszystkich zapisanych plików (jeśli dotnet dostępny)
  if (ooxml.available()) {
    const outFiles = docxBuffers.map(([name, buf]) => {
      const p = path.join(TMP, name);
      fs.writeFileSync(p, buf);
      return p;
    });
    const rep = ooxml.validate(outFiles);
    const bad = rep.filter((x) => x.errors.length);
    check("walidacja OOXML (OpenXML SDK): wszystkie przekonwertowane tabele poprawne", bad.length === 0, bad.map((x) => `${path.basename(x.file)}: ${x.errors[0]}`).join(" | "));
  }
  fs.rmSync(TMP, { recursive: true, force: true });

  if (errors.length) check("bez błędów JS na stronie", false, errors.slice(0, 3).join(" | "));
  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? ` — ${r.detail}` : ""}`);
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
