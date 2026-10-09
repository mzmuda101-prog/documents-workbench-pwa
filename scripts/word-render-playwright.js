// word-render-playwright.js — podgląd składa dokument jak Word (2026-10-09, porównanie z PDF-ami
// z Worda: npm run word:compare na 41 plikach).
//
// (1) Styl domyślny akapitu („Normalny”) działa na TEKST, nie tylko na akapit: docx-preview
//     budował „.docx p, p.docx_normalny span {…}” — krój/rozmiar Normalnego przegrywały z
//     docDefaults (podanie z Liberation Serif 11 pt wyglądało jak Calibri 12 pt, 2 strony
//     zamiast 1). Style znakowe i style akapitów dalej wygrywają z Normalnym.
// (2) Sekcja „ciągła” zaczyna się na tej samej stronie (kolejny <article> tej samej kartki, np.
//     w 2 kolumnach); sekcja bez typu (= następna strona) i podział strony — nowa strona.
//     Dawniej każda sekcja = nowa kartka („ANALIZA POTRZEB…”: 5 stron zamiast 2).
// (3) „Bez odstępu między akapitami tego samego stylu” (w:contextualSpacing, styl „Akapit z listą”):
//     punkty listy bez 8 pt odstępu między sobą, ostatni z odstępem przed innym stylem.
// (4) Pusty akapit = jedna linijka liczona od znaku końca akapitu (w:pPr/w:rPr/w:sz) — nie
//     rozmiar czcionki stylu bez interlinii; kolor znaku akapitu dla punktora (--dwb-mark-color).
// (5) Indeks górny: 60 % rozmiaru i nie podnosi linijki (akapit tej samej wysokości co zwykły).
// (6) Tabele: linia pozioma doliczana do wysokości wiersza (17 pt + 0,5 pt), pierwszy wiersz
//     scalenia w pionie ≥ pierwsza linijka (reguła zmierzona w Wordzie), wiersz nagłówka
//     (w:tblHeader) powtarzany na następnej stronie podglądu wydruku.
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
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const t = (s) => `<w:t xml:space="preserve">${s}</w:t>`;
const P = (s, pPr = "", rPr = "") => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}${t(s)}</w:r></w:p>`;
const SECT = (extra = "") => `<w:sectPr>${extra}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`;

async function fixture() {
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  // docDefaults: Courier New 12 pt; Normalny: Times New Roman 11 pt (jak w podaniu: Normalny ≠ domyślne)
  z.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>
    <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normalny"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr></w:style>
    <w:style w:type="paragraph" w:styleId="Naglowek"><w:name w:val="Nagłówek"/><w:basedOn w:val="Normalny"/><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="32"/></w:rPr></w:style>
    <w:style w:type="paragraph" w:styleId="ListaP"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normalny"/><w:pPr><w:spacing w:after="160"/><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
    <w:style w:type="character" w:styleId="Link"><w:name w:val="Link"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:color w:val="0000FF"/></w:rPr></w:style>
  </w:styles>`);
  const body = [
    P("Zwykły akapit"), // 0
    P("Fragment ze stylem znakowym", "", '<w:rStyle w:val="Link"/>'), // 1
    P("Nagłówek ze stylu akapitu", '<w:pStyle w:val="Naglowek"/>'), // 2
    P("Fragment z własnym krojem", "", '<w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/>'), // 3
    P("Punkt pierwszy", '<w:pStyle w:val="ListaP"/>'), P("Punkt drugi", '<w:pStyle w:val="ListaP"/>'), P("Punkt ostatni", '<w:pStyle w:val="ListaP"/><w:rPr><w:color w:val="FF0000"/></w:rPr>'),
    P("Akapit po liście"),
    '<w:p><w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr></w:p>', // pusty, znak akapitu 20 pt
    P("Linijka zwykła"),
    `<w:p><w:r>${t("Linijka ze znacznikiem ")}</w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>${t("12:30")}</w:r></w:p>`,
    // koniec sekcji 1 (bez typu: pierwsza i tak)
    `<w:p><w:pPr>${SECT()}</w:pPr><w:r>${t("Koniec sekcji pierwszej")}</w:r></w:p>`,
    // sekcja 2: ciągła, 2 kolumny — ta sama strona
    P("Lewa kolumna A"), P("Lewa kolumna B"), P("Prawa kolumna C"),
    `<w:p><w:pPr>${SECT('<w:type w:val="continuous"/>').replace("<w:pgSz", '<w:cols w:num="2" w:space="708"/><w:pgSz')}</w:pPr><w:r>${t("Koniec sekcji kolumn")}</w:r></w:p>`,
    // sekcja 3: bez typu = następna strona; w niej jawny podział strony
    P("Sekcja trzecia — nowa strona"),
    `<w:p><w:r><w:br w:type="page"/></w:r><w:r>${t("Po podziale strony")}</w:r></w:p>`,
  ].join("");
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${SECT()}</w:body></w:document>`);
  return z.generateAsync({ type: "nodebuffer" });
}

// Drugi dokument: tabela z nagłówkiem na 2 strony (obramowana) i para wierszy scalonych w pionie.
async function fixtureTables() {
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const cell = (txt, extra = "", pPr = "") => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>${extra}</w:tcPr><w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${txt ? `<w:r>${t(txt)}</w:r>` : ""}</w:p></w:tc>`;
  const b = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>';
  const mar = '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="40" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tblCellMar>';
  const rows = [`<w:tr><w:trPr><w:tblHeader/><w:trHeight w:val="400"/></w:trPr>${cell("Nagłówek tabeli")}</w:tr>`];
  for (let i = 1; i <= 60; i++) rows.push(`<w:tr><w:trPr><w:trHeight w:val="340"/></w:trPr>${cell(`Wiersz ${i}`)}</w:tr>`);
  const big = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/>${b}${mar}</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${rows.join("")}</w:tbl>`;
  // para scalona w pionie (bez linii): linijka dokładnie 9,4 pt, odstęp przed 3,5 pt; minima 3,5 + 9,95 pt
  const pair = (i) => `<w:tr><w:trPr><w:trHeight w:val="70" w:hRule="atLeast"/></w:trPr>${cell(`Scalona ${i}`, '<w:vMerge w:val="restart"/>', '<w:spacing w:before="70" w:after="0" w:line="188" w:lineRule="exact"/><w:rPr><w:sz w:val="16"/></w:rPr>')}</w:tr><w:tr><w:trPr><w:trHeight w:val="199" w:hRule="atLeast"/></w:trPr>${cell("", "<w:vMerge/>", '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>')}</w:tr>`;
  const merged = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/>${mar}</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${pair(1)}${pair(2)}</w:tbl>`;
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${merged}${P("Między tabelami")}${big}${P("Koniec")}${SECT()}</w:body></w:document>`);
  return z.generateAsync({ type: "nodebuffer" });
}

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
  const file = path.join(os.tmpdir(), `dwb-word-render-${process.pid}.docx`);
  fs.writeFileSync(file, await fixture());
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 20000 });
  fs.rmSync(file, { force: true });
  await page.waitForTimeout(500);

  const info = await page.evaluate(() => {
    const host = document.querySelector(".docx-preview-host");
    const run = (txt) => {
      const sp = [...host.querySelectorAll("section.docx span")].find((s) => s.textContent.includes(txt));
      if (!sp) return null;
      const cs = getComputedStyle(sp);
      return { font: cs.fontFamily.split(",")[0].replace(/["']/g, "").trim(), size: Math.round(parseFloat(cs.fontSize) * 0.75 * 10) / 10, color: cs.color };
    };
    const secs = [...host.querySelectorAll(".docx-wrapper > section.docx")];
    return {
      plain: run("Zwykły akapit"), link: run("stylem znakowym"), head: run("Nagłówek ze stylu"), own: run("własnym krojem"),
      sections: secs.length,
      articles: secs.map((s) => s.querySelectorAll(":scope > article").length),
      cols: secs[0] ? [...secs[0].querySelectorAll(":scope > article")].map((a) => getComputedStyle(a).columnCount) : [],
      firstSecText: secs[0]?.textContent || "",
      list: ["Punkt pierwszy", "Punkt drugi", "Punkt ostatni"].map((x) => { const p = [...host.querySelectorAll("p")].find((q) => q.textContent === x); const cs = getComputedStyle(p); return { mt: parseFloat(cs.marginTop), mb: parseFloat(cs.marginBottom), mark: p.style.getPropertyValue("--dwb-mark-color") }; }),
      empty: (() => { const p = [...host.querySelectorAll("p")].find((q) => q.dataset.dwbMarkSize === "20.00pt" || q.dataset.dwbMarkSize === "20pt"); return p ? { h: p.getBoundingClientRect().height, mark: p.dataset.dwbMarkSize } : null; })(),
      sup: (() => { const plain = [...host.querySelectorAll("p")].find((q) => q.textContent === "Linijka zwykła"); const p = [...host.querySelectorAll("p")].find((q) => q.textContent.startsWith("Linijka ze znacznikiem")); const sup = p?.querySelector("sup"); return { plainH: plain?.getBoundingClientRect().height, supH: p?.getBoundingClientRect().height, supFs: sup && parseFloat(getComputedStyle(sup).fontSize), baseFs: sup && parseFloat(getComputedStyle(sup.parentElement).fontSize) }; })(),
    };
  });
  check("Normalny działa na tekst: Times New Roman 11 pt (nie Courier New 12 pt z docDefaults)", info.plain?.font === "Times New Roman" && info.plain?.size === 11, JSON.stringify(info.plain));
  check("styl znakowy wygrywa z Normalnym (Arial, niebieski)", info.link?.font === "Arial" && /0, 0, 255/.test(info.link?.color || ""), JSON.stringify(info.link));
  check("styl akapitu wygrywa z Normalnym (Georgia 16 pt)", info.head?.font === "Georgia" && info.head?.size === 16, JSON.stringify(info.head));
  check("krój wprost we fragmencie wygrywa (Verdana)", info.own?.font === "Verdana", JSON.stringify(info.own));
  const [l1, l2, l3] = info.list;
  check("lista (contextualSpacing): bez odstępu między punktami", l1.mb === 0 && l2.mt === 0 && l2.mb === 0 && l3.mt === 0, JSON.stringify(info.list));
  check("lista: ostatni punkt z odstępem 8 pt przed akapitem innego stylu", Math.abs(l3.mb - 160 / 15) < 0.2, JSON.stringify(l3));
  check("kolor znaku akapitu w zmiennej dla punktora (#FF0000)", /^#ff0000$/i.test(l3.mark) && !l1.mark, JSON.stringify(info.list.map((x) => x.mark)));
  // Times New Roman: wysokość linijki 1,149 × rozmiar (WORD_LINE_FACTORS); znak akapitu 20 pt
  check("pusty akapit = linijka ze znaku akapitu (20 pt × 1,149 ≈ 30,6 px)", info.empty && Math.abs(info.empty.h - 20 * 1.149 * 4 / 3) < 1, JSON.stringify(info.empty));
  check("indeks górny: 60 % rozmiaru tekstu", info.sup.supFs && Math.abs(info.sup.supFs / info.sup.baseFs - 0.6) < 0.02, JSON.stringify(info.sup));
  check("indeks górny nie podnosi linijki (wysokość jak zwykłego akapitu)", Math.abs(info.sup.supH - info.sup.plainH) < 0.6, JSON.stringify(info.sup));
  check("sekcja ciągła na tej samej kartce: 3 kartki (sekcje 1+2 | sekcja 3 | po podziale), nie 4", info.sections === 3, JSON.stringify({ sections: info.sections, articles: info.articles }));
  check("pierwsza kartka: 2 bloki treści, drugi w 2 kolumnach", info.articles[0] === 2 && info.cols[1] === "2", JSON.stringify({ articles: info.articles, cols: info.cols }));
  check("kolumny zostają na pierwszej kartce", /Prawa kolumna C/.test(info.firstSecText) && !/Sekcja trzecia/.test(info.firstSecText), info.firstSecText.slice(0, 200));

  // Podgląd wydruku: 3 strony (sekcje 1+2 | sekcja 3 | po podziale), kolumny obok siebie
  await page.evaluate(() => dwbPrint.open());
  await page.waitForSelector(".pp-sheet", { timeout: 20000 });
  await page.waitForTimeout(400);
  const pp = await page.evaluate(() => {
    const sheets = [...document.querySelectorAll(".pp-sheet")];
    const pos = (txt) => {
      const p = [...sheets[0].querySelectorAll("p")].find((x) => x.textContent.includes(txt));
      return p ? p.getBoundingClientRect() : null;
    };
    const a = pos("Lewa kolumna A"), c = pos("Prawa kolumna C");
    return { n: sheets.length, texts: sheets.map((s) => s.textContent.replace(/\s+/g, " ").slice(0, 80)), side: a && c ? c.left - a.left : null, rowDiff: a && c ? Math.abs(c.top - a.top) : null };
  });
  check("podgląd wydruku: 3 strony", pp.n === 3, JSON.stringify(pp.texts));
  check("podgląd wydruku: kolumna prawa obok lewej (ta sama wysokość, przesunięta w prawo)", pp.side > 150 && pp.rowDiff < 40, JSON.stringify(pp));

  // ── tabele (drugi dokument) ──
  await page.evaluate(() => { if (window.dwbPrint?.isOpen()) dwbPrint.close(); if (typeof setDirtyState === "function") setDirtyState(false); });
  const file2 = path.join(os.tmpdir(), `dwb-word-render-t-${process.pid}.docx`);
  fs.writeFileSync(file2, await fixtureTables());
  await page.setInputFiles("#fileInput", file2);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && [...document.querySelectorAll(".docx-preview-host td")].some((td) => td.textContent === "Wiersz 1"), null, { timeout: 20000 });
  fs.rmSync(file2, { force: true });
  await page.waitForTimeout(500);
  const tb = await page.evaluate(() => {
    const host = document.querySelector(".docx-preview-host");
    const tr = (txt) => [...host.querySelectorAll("tr")].find((r) => r.textContent === txt);
    const r1 = tr("Wiersz 1"), m1 = tr("Scalona 1"), m2 = tr("Scalona 2");
    return { row: r1.getBoundingClientRect().height, pair: m2.getBoundingClientRect().top - m1.getBoundingClientRect().top, head: !!tr("Nagłówek tabeli")?.dataset.dwbHeader };
  });
  check("wiersz tabeli: 17 pt + linia 0,5 pt (jak Word, 23,33 px)", Math.abs(tb.row - 17.5 * 4 / 3) < 0.4, JSON.stringify(tb));
  // WebKit wkłada całą treść komórki scalonej do PIERWSZEGO wiersza (17,2 px zamiast 12,5) — tak
  // było też przed regułą; tam tylko pilnujemy, żeby nie było gorzej (30,45 px)
  if (ENGINE === "webkit") check("scalenie w pionie (WebKit): nie wyżej niż przed regułą Worda", tb.pair < 31, JSON.stringify(tb));
  else check("scalenie w pionie: para wierszy = linijka 9,4 pt + minimum 9,95 pt (Word 19,4 pt ≈ 25,8 px)", Math.abs(tb.pair - 19.35 * 4 / 3) < 0.8, JSON.stringify(tb));
  check("wiersz nagłówka oznaczony (w:tblHeader bez wartości = tak)", tb.head, JSON.stringify(tb));
  await page.evaluate(() => dwbPrint.open());
  await page.waitForSelector('.pp-sheet[data-page="2"]', { timeout: 20000 });
  await page.waitForTimeout(400);
  const rep2 = await page.evaluate(() => {
    const s = document.querySelector('.pp-sheet[data-page="2"]');
    const h = s.querySelector(".pp-repeat-head");
    const clip = s.querySelector(".pp-clip").getBoundingClientRect();
    const rows = [...s.querySelectorAll(".pp-clip tr")].filter((r) => { const b = r.getBoundingClientRect(); return b.top >= clip.top - 1 && b.bottom <= clip.bottom + 1 && b.height; });
    return { head: h?.textContent || "", headBottom: h ? Math.round(h.getBoundingClientRect().bottom) : null, firstTop: rows[0] ? Math.round(rows[0].getBoundingClientRect().top) : null, first: rows[0]?.textContent };
  });
  check("podgląd wydruku: nagłówek tabeli powtórzony na stronie 2, wiersze pod nim", rep2.head === "Nagłówek tabeli" && rep2.firstTop >= rep2.headBottom - 1 && /^Wiersz \d+$/.test(rep2.first || ""), JSON.stringify(rep2));

  await browser.close();
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Skład jak w Wordzie [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
