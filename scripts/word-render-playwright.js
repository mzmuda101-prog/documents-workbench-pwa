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
// (7) Reguły zmierzone wyrocznią Worda 2026-10-25 (podania do dziekana): przebieg z samych spacji/
//     tabulatorów nie podnosi linijki (Calibri 28 pt obok Arial 8 pt — 12,33 px), akapit z samych
//     spacji = pusty (spacja 12 pt + znak 5 pt → 7,67 px), spacje na końcu akapitu do prawej wiszą,
//     interlinia w komórce: domyślne < styl tabeli < styl akapitu, stopka dosunięta do dołu
//     (w:footer od krawędzi), wcięcie stylu akapitu w stopce zostaje w podglądzie wydruku;
//     symbol w kroju zastępczym Worda (✓ — Segoe UI Symbol) podnosi swoją linijkę.
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
  // ta sama para, ale obok zwykła kolumna (kwota w 2. wierszu — jak tabele z konwersji PDF): Word NIE
  // podnosi 1. wiersza do linijki, minima 3,5 + 9,95 pt wystarczą (wyrocznia 2026-10-28: 13,25 pt)
  const side = (w, txt, extra, pPr) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${extra}</w:tcPr><w:p><w:pPr>${pPr}<w:rPr><w:sz w:val="16"/></w:rPr></w:pPr>${txt ? `<w:r><w:rPr><w:sz w:val="16"/></w:rPr>${t(txt)}</w:r>` : ""}</w:p></w:tc>`;
  const pairSide = (i) => `<w:tr><w:trPr><w:trHeight w:val="70" w:hRule="atLeast"/></w:trPr>${side(2700, `Obok ${i}`, '<w:vMerge w:val="restart"/>', '<w:spacing w:before="70" w:after="0" w:line="188" w:lineRule="exact"/>')}${side(1300, "", "", '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>')}</w:tr><w:tr><w:trPr><w:trHeight w:val="199" w:hRule="atLeast"/></w:trPr>${side(2700, "", "<w:vMerge/>", '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>')}${side(1300, `kwota ${i}`, "", '<w:spacing w:before="1" w:after="0" w:line="188" w:lineRule="exact"/>')}</w:tr>`;
  const mergedSide = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/>${mar}</w:tblPr><w:tblGrid><w:gridCol w:w="2700"/><w:gridCol w:w="1300"/></w:tblGrid>${pairSide(1)}${pairSide(2)}</w:tbl>`;
  // marginesy komórki ponad minimum wiersza, wiersz „dokładnie”, w:trPr/w:jc (położenie, nie tekst)
  const bc = '<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>';
  const mrow = (txt, trPr) => `<w:tr><w:trPr>${trPr}</w:trPr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>${bc}</w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>${t(txt)}</w:r></w:p></w:tc></w:tr>`;
  const marg = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblCellMar><w:top w:w="200" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${mrow("Min A", '<w:trHeight w:val="340"/><w:jc w:val="center"/>')}${mrow("Min B", '<w:trHeight w:val="340"/>')}${mrow("Exact A", '<w:trHeight w:val="1000" w:hRule="exact"/>')}${mrow("Exact B", '<w:trHeight w:val="1000" w:hRule="exact"/>')}</w:tbl>`;
  // wiersze bez minimum (wysokość z treści), Calibri 11 pt, po 8 pt, 1,08: Word 23 pt = 30,67 px na
  // wiersz (wyrocznia probe-row-bottom.js, Kolokwium 18.11) — linia 0,5 pt to 0,67 px, nie 1 px
  const crow = (txt) => `<w:tr>${cell("", "", "").replace("<w:p></w:p>", `<w:p><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>${t(txt)}</w:r></w:p>`)}</w:tr>`;
  const content = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/>${b}</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${crow("Treść 1")}${crow("Treść 2")}${crow("Treść 3")}</w:tbl>`;
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${merged}${P("Obok scalenia")}${mergedSide}${P("Przed marginesami")}${marg}${P("Przed treścią")}${content}${P("Między tabelami")}${big}${P("Koniec")}${SECT()}</w:body></w:document>`);
  return z.generateAsync({ type: "nodebuffer" });
}

// Trzeci dokument: reguły z wyroczni Worda (7). Strona A4, margines dolny 2 cm, stopka 0,5 cm od krawędzi.
async function fixtureRules() {
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>');
  z.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>
    <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
    <w:style w:type="paragraph" w:styleId="Podwojna"><w:name w:val="Podwojna"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="Wciety"><w:name w:val="Wciety"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0"/><w:ind w:left="720"/></w:pPr></w:style>
    <w:style w:type="table" w:styleId="TS"><w:name w:val="TS"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
  </w:styles>`);
  const A8 = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/>';
  const C28 = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="56"/>';
  const sp = '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>';
  const R = (rPr, inner) => `<w:r><w:rPr>${rPr}</w:rPr>${inner}</w:r>`;
  const cellTbl = (pPr, k) => `<w:tbl><w:tblPr><w:tblStyle w:val="TS"/><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${R('<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/>', `${t(k + "a")}<w:br/>${t(k + "b")}`)}</w:p></w:tc></w:tr></w:tbl>`;
  const body = [
    `<w:p><w:pPr>${sp}</w:pPr>${R(A8, t("Wzorzec 1"))}</w:p>`,
    `<w:p><w:pPr>${sp}</w:pPr>${R(C28, "<w:tab/>")}${R(A8, t("Po tabulatorze 28 pt"))}</w:p>`,
    `<w:p><w:pPr>${sp}</w:pPr>${R(C28, t(" "))}${R(A8, t("Po spacji 28 pt"))}</w:p>`,
    `<w:p><w:pPr>${sp}</w:pPr>${R(A8, t("Wzorzec 2"))}</w:p>`,
    `<w:p><w:pPr>${sp}<w:rPr><w:sz w:val="10"/></w:rPr></w:pPr>${R('<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/>', t(" "))}</w:p>`,
    `<w:p><w:pPr>${sp}</w:pPr>${R(A8, t("Wzorzec 3"))}</w:p>`,
    `<w:p><w:pPr>${sp}<w:jc w:val="right"/></w:pPr>${R('<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/>', t("Do prawej   "))}</w:p>`,
    // symbol w kroju zastępczym Worda (Segoe UI Symbol) podnosi linijkę, strzałka z Ariala nie
    ...["Wzorzec 4", "Znak ✓ tutaj", "Wzorzec 5", "Strzałka → tutaj", "Wzorzec 6"].map((x) => `<w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr>${R('<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>', t(x))}</w:p>`),
    cellTbl("", "T"), cellTbl('<w:pStyle w:val="Podwojna"/>', "S"),
    `<w:p><w:pPr>${sp}</w:pPr>${R(A8, t("Koniec"))}</w:p>`,
  ].join("");
  z.file("word/footer1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NS}><w:p><w:pPr><w:pStyle w:val="Wciety"/></w:pPr>${R(A8, t("Stopka wcięta"))}</w:p></w:ftr>`);
  const sect = '<w:sectPr><w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdF"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="283" w:gutter="0"/></w:sectPr>';
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}${sect}</w:body></w:document>`);
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
  await page.addInitScript(() => { window.__dwbAutoDialog = "accept"; }); // własne okna apki jak systemowe
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
  // Word: 2/3 rozmiaru zaokrąglone w dół do pół punktu (11 pt → 7 pt)
  check("indeks górny: 2/3 rozmiaru w dół do pół punktu (11 pt → 7 pt)", info.sup.supFs && Math.abs(info.sup.supFs * 0.75 - Math.floor(info.sup.baseFs * 0.75 * 4 / 3) / 2) < 0.05, JSON.stringify(info.sup));
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
  // Word (tools/word-oracle/probe-columns-balance.js): kolumny przed sekcją „następna strona” NIE są
  // wyrównywane — krótki tekst stoi cały w kolumnie 1 (C pod B); wyrównanie tylko przed sekcją ciągłą
  check("podgląd wydruku: kolumny jak w Wordzie — przed nową stroną bez wyrównania (C pod A i B, w kolumnie 1)", Math.abs(pp.side) < 10 && pp.rowDiff > 20, JSON.stringify(pp));

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
    const top = (x) => tr(x).getBoundingClientRect().top;
    const minA = tr("Min A");
    return { row: r1.getBoundingClientRect().height, pair: m2.getBoundingClientRect().top - m1.getBoundingClientRect().top, pairSide: top("Obok 2") - top("Obok 1"), head: !!tr("Nagłówek tabeli")?.dataset.dwbHeader,
      minFirst: top("Min B") - top("Min A"), min: top("Exact A") - top("Min B"), exact: top("Exact B") - top("Exact A"),
      rowAlign: getComputedStyle(minA.cells[0].querySelector("p")).textAlign, content: top("Treść 3") - top("Treść 2") };
  });
  check("wiersz tabeli: 17 pt + linia 0,5 pt (jak Word, 23,33 px)", Math.abs(tb.row - 17.5 * 4 / 3) < 0.4, JSON.stringify(tb));
  // WebKit wkłada całą treść komórki scalonej do PIERWSZEGO wiersza (17,2 px zamiast 12,5) — tak
  // było też przed regułą; tam tylko pilnujemy, żeby nie było gorzej (30,45 px)
  if (ENGINE === "webkit") check("scalenie w pionie (WebKit): nie wyżej niż przed regułą Worda", tb.pair < 31, JSON.stringify(tb));
  else check("scalenie w pionie: para wierszy = linijka 9,4 pt + minimum 9,95 pt (Word 19,4 pt ≈ 25,8 px)", Math.abs(tb.pair - 19.35 * 4 / 3) < 0.8, JSON.stringify(tb));
  check("scalenie w pionie ze zwykłą komórką obok: minima wierszy, bez linijki (Word 13,25 pt ≈ 17,7 px)", Math.abs(tb.pairSide - 13.45 * 4 / 3) < 0.8, JSON.stringify(tb));
  check("wiersz nagłówka oznaczony (w:tblHeader bez wartości = tak)", tb.head, JSON.stringify(tb));
  // wyrocznia 2026-10-25: minimum 17 pt + margines komórki 200 tw + linia 0,5 pt = 36,7 px; „dokładnie” 1000 tw = 66,7 px
  check("minimum wiersza bez marginesów komórki — marginesy ponad nie (Word 36,7 px)", Math.abs(tb.min - 36.67) < 0.6, JSON.stringify(tb));
  check("pierwszy wiersz bez górnej linii tabeli — tej samej wysokości co kolejny", Math.abs(tb.minFirst - tb.min) < 0.3, JSON.stringify(tb));
  check("wiersz „dokładnie” (hRule exact): linia i marginesy w środku (Word 66,7 px)", Math.abs(tb.exact - 66.67) < 0.6, JSON.stringify(tb));
  check("wiersz z treścią (bez minimum): linia 0,5 pt dolicza 0,67 px, nie 1 px (Word 30,67 px)", Math.abs(tb.content - 30.67) < 0.15, JSON.stringify(tb));
  check("w:trPr/w:jc to położenie wiersza — tekst komórki dalej do lewej", tb.rowAlign === "left" || tb.rowAlign === "start", JSON.stringify(tb));
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

  // ── reguły z wyroczni Worda (trzeci dokument) ──
  await page.evaluate(() => { if (window.dwbPrint?.isOpen()) dwbPrint.close(); if (typeof setDirtyState === "function") setDirtyState(false); });
  const file3 = path.join(os.tmpdir(), `dwb-word-render-r-${process.pid}.docx`);
  fs.writeFileSync(file3, await fixtureRules());
  await page.setInputFiles("#fileInput", file3);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && [...document.querySelectorAll(".docx-preview-host p")].some((p) => p.textContent.startsWith("Do prawej")), null, { timeout: 20000 }); // „Koniec” ma też poprzedni plik
  fs.rmSync(file3, { force: true });
  await page.waitForTimeout(600);
  const ru = await page.evaluate(() => {
    const host = document.querySelector(".docx-preview-host");
    const texts = []; const tw = document.createTreeWalker(host.querySelector("section.docx article"), NodeFilter.SHOW_TEXT); for (let n; (n = tw.nextNode());) texts.push(n);
    const top = (txt) => { const n = texts.find((x) => x.data === txt); if (!n) return NaN; const r = document.createRange(); r.selectNodeContents(n); return r.getBoundingClientRect().top; };
    const right = [...host.querySelectorAll("section.docx article p")].find((p) => p.textContent.startsWith("Do prawej"));
    const rr = document.createRange(); rr.setStart(right.querySelector("span").firstChild, 0); rr.setEnd(right.querySelector("span").firstChild, 9);
    const pr = right.getBoundingClientRect(), pcs = getComputedStyle(right);
    return {
      tab: top("Po tabulatorze 28 pt") - top("Wzorzec 1"), space: top("Po spacji 28 pt") - top("Po tabulatorze 28 pt"), next: top("Wzorzec 2") - top("Po spacji 28 pt"),
      wsOnly: top("Wzorzec 3") - top("Wzorzec 2") - 12.25,
      trail: (pr.right - parseFloat(pcs.paddingRight)) - rr.getBoundingClientRect().right,
      table: top("Tb") - top("Ta"), style: top("Sb") - top("Sa"),
      sym: (top("Wzorzec 5") - top("Wzorzec 4")) - (top("Wzorzec 6") - top("Wzorzec 5")), arrow: top("Wzorzec 6") - top("Wzorzec 5"),
    };
  });
  check("przebieg z samego tabulatora/spacji 28 pt nie podnosi linijki 8 pt (Word 12,33 px)", [ru.tab, ru.space, ru.next].every((v) => Math.abs(v - 12.3) < 0.7), JSON.stringify(ru));
  check("akapit z samych spacji = pusty: linijka znaku akapitu 5 pt (Word 7,67 px)", Math.abs(ru.wsOnly - 7.67) < 0.7, JSON.stringify(ru));
  check("spacje na końcu akapitu do prawej wiszą za marginesem (tekst przy prawej krawędzi)", Math.abs(ru.trail) < 1, JSON.stringify(ru));
  // wyrocznia: Arial 11 pt × 1,15 = 19,33 px na linijkę; ✓ (Segoe UI Symbol) +2,7 px, → bez zmian
  check("symbol ✓ (krój zastępczy Worda) podnosi swoją linijkę o ~2,7 px", Math.abs(ru.sym - 2.69) < 0.6, JSON.stringify(ru));
  check("strzałka → (jest w Arialu) nie podnosi linijki (2 × 19,33 px)", Math.abs(ru.arrow - 38.66) < 1, JSON.stringify(ru));
  check("komórka: styl tabeli wygrywa z domyślnymi (pojedyncza, Word 15,33 px)", Math.abs(ru.table - 15.33) < 0.7, JSON.stringify(ru));
  check("komórka: styl akapitu wygrywa ze stylem tabeli (podwójna, Word 30,7 px)", Math.abs(ru.style - 30.7) < 1, JSON.stringify(ru));
  await page.evaluate(() => dwbPrint.open());
  await page.waitForSelector(".pp-sheet", { timeout: 20000 });
  await page.waitForTimeout(500);
  const ft = await page.evaluate(() => {
    const sheet = document.querySelector(".pp-sheet");
    const S = sheet.getBoundingClientRect();
    const zoom = S.width / sheet.offsetWidth || 1;
    const p = [...sheet.querySelectorAll(".pp-hf p")].find((x) => x.textContent === "Stopka wcięta");
    const r = document.createRange(); r.selectNodeContents(p.querySelector("span").firstChild);
    const b = r.getBoundingClientRect();
    return { bottom: (b.bottom - S.top) / zoom, pageH: S.height / zoom, left: (b.left - S.left) / zoom, ml: parseFloat(getComputedStyle(p).marginLeft) };
  });
  // dół stopki ≈ strona − 0,5 cm (283 tw = 18,9 px); margines dolny 2 cm = 75,6 px
  check("stopka dosunięta do dołu (dół tekstu ~0,5 cm od krawędzi, nie pod marginesem 2 cm)", ft.bottom > ft.pageH - 30 && ft.bottom <= ft.pageH - 15, JSON.stringify(ft));
  check("podgląd wydruku: wcięcie stylu akapitu w stopce zostaje (36 pt = 48 px)", Math.abs(ft.ml - 48) < 0.6 && Math.abs(ft.left - (75.6 + 48)) < 2, JSON.stringify(ft));

  // ── początek strony i symbole (czwarty dokument; wyrocznia probe-before-top / probe-sym-pitch) ──
  await page.evaluate(() => { if (window.dwbPrint?.isOpen()) dwbPrint.close(); if (typeof setDirtyState === "function") setDirtyState(false); });
  const file4 = path.join(os.tmpdir(), `dwb-word-render-s-${process.pid}.docx`);
  const { buildDocx } = require("../tools/word-oracle/build-docx");
  const head = (txt, extra = "") => `<w:p><w:pPr>${extra}<w:spacing w:before="320" w:after="160"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="40"/></w:rPr>${t(txt)}</w:r></w:p>`;
  const arial = (txt) => `<w:p><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/></w:rPr>${t(txt)}</w:r></w:p>`;
  // ramki akapitu (wyrocznia probe-para-group.js): grupa = jedno pole; linia = grubość + odstęp
  const bside = (n, sz, sp) => `<w:${n} w:val="single" w:sz="${sz}" w:space="${sp}" w:color="000000"/>`;
  const bbox = (sz, sp, between = "") => `<w:pBdr>${bside("top", sz, sp)}${bside("left", sz, 4)}${bside("bottom", sz, sp)}${bside("right", sz, 4)}${between}</w:pBdr>`;
  const bp = (txt, bdr) => `<w:p><w:pPr>${bdr}<w:spacing w:after="0"/></w:pPr><w:r>${t(txt)}</w:r></w:p>`;
  const plainP = (txt) => `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r>${t(txt)}</w:r></w:p>`;
  const boxes = [plainP("Przed ramka"), ...["Ga1", "Ga2", "Ga3"].map((x) => bp(x, bbox(4, 1))), plainP("Miedzy ramkami"),
    ...["Gb1", "Gb2"].map((x) => bp(x, bbox(34, 1))), plainP("Przed miedzy"), ...["Gc1", "Gc2"].map((x) => bp(x, bbox(4, 1, bside("between", 4, 1)))), plainP("Po ramkach")];
  const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
  fs.writeFileSync(file4, await buildDocx({ body: [head("Start dokumentu"), arial("Zwykla linijka"), arial("Skrot Ctrl/⌘+F"), arial("Strzalka a ← b"),
    '<w:p><w:r><w:br w:type="page"/></w:r></w:p>', head("Po podziale strony"), head("Po podziale przed", "<w:pageBreakBefore/>"),
    `<w:p><w:pPr><w:spacing w:after="160"/>${sect}</w:pPr><w:r>${t("Koniec sekcji")}</w:r></w:p>`, head("Nowa sekcja"),
    ...boxes] }));
  await page.setInputFiles("#fileInput", file4);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && [...document.querySelectorAll(".docx-preview-host p")].some((p) => p.textContent === "Nowa sekcja"), null, { timeout: 20000 });
  fs.rmSync(file4, { force: true });
  await page.waitForTimeout(600);
  const st = await page.evaluate(() => {
    const ps = [...document.querySelectorAll(".docx-preview-host section.docx p")];
    const p = (txt) => ps.find((x) => x.textContent === txt);
    const mt = (txt) => parseFloat(getComputedStyle(p(txt)).marginTop);
    const h = (txt) => p(txt).getBoundingClientRect().height;
    const tt = (txt) => { const rg = document.createRange(); rg.selectNodeContents(p(txt)); return rg.getBoundingClientRect().top * 0.75; };
    const bx = { a0: tt("Ga1") - tt("Przed ramka"), a1: tt("Ga2") - tt("Ga1"), a3: tt("Miedzy ramkami") - tt("Ga3"), b0: tt("Gb1") - tt("Miedzy ramkami"), b1: tt("Gb2") - tt("Gb1"), c1: tt("Gc2") - tt("Gc1"),
      text: p("Ga1").getBoundingClientRect().left + parseFloat(getComputedStyle(p("Ga1")).paddingLeft) - p("Przed ramka").getBoundingClientRect().left };
    return { bx, first: mt("Start dokumentu"), brk: mt("Po podziale strony"), pbb: mt("Po podziale przed"), sect: mt("Nowa sekcja"), plain: h("Zwykla linijka"), cmd: h("Skrot Ctrl/⌘+F"), left: h("Strzalka a ← b") };
  });
  check("odstęp przed na początku dokumentu zostaje (16 pt = 21,33 px)", Math.abs(st.first - 21.33) < 0.1, JSON.stringify(st));
  check("strona po podziale (w:br, pageBreakBefore): odstęp przed pominięty jak w Wordzie", st.brk === 0 && st.pbb === 0, JSON.stringify(st));
  check("początek sekcji: odstęp przed minus po z poprzedniej strony (16 − 8 pt = 10,67 px)", Math.abs(st.sect - 10.67) < 0.1, JSON.stringify(st));
  const near = (v, w) => Math.abs(v - w) < 0.05;
  check("ramka akapitu: grupa = jedno pole (Word: 16 / 14,5 / 16 pt), tekst na swoim wcięciu", near(st.bx.a0, 16) && near(st.bx.a1, 14.5) && near(st.bx.a3, 16) && Math.abs(st.bx.text) < 0.5, JSON.stringify(st.bx));
  check("ramka akapitu 4,25 pt + odstęp 1 pt: bez zaokrąglania do pikseli (Word 19,75 pt)", near(st.bx.b0, 19.75) && near(st.bx.b1, 14.5), JSON.stringify(st.bx));
  check("linia „między” akapitami: grubość + 2 × odstęp (Word 17 pt)", near(st.bx.c1, 17), JSON.stringify(st.bx));
  check("⌘ (krój zastępczy 1,172 em) podnosi linijkę Arial 11 pt o ~0,38 px", Math.abs(st.cmd - st.plain - 0.383) < 0.15, JSON.stringify(st));
  check("← (jest w Arialu) nie podnosi linijki", Math.abs(st.left - st.plain) < 0.05, JSON.stringify(st));

  // ── wcięcie tabeli a tryb zgodności (wyrocznia probe-table-indent.js): tblInd 0 + margines komórki
  //    5,4 pt; tryb 15 — tekst komórki 7,2 px od marginesu, tryb 14 / bez ustawień — na marginesie ──
  const JSZipT = require("jszip");
  const tbB = ["top", "left", "bottom", "right", "insideH", "insideV"].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join("");
  const tblXml = (txt) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>${tbB}</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r>${t(txt)}</w:r></w:p></w:tc></w:tr></w:tbl>`;
  const indentOf = async (mode) => {
    const zip = await JSZipT.loadAsync(await buildDocx({ body: [`<w:p><w:r>${t(`Tekst ${mode}`)}</w:r></w:p>`, tblXml(`Komorka ${mode}`), "<w:p/>"] }));
    zip.file("word/settings.xml", (await zip.file("word/settings.xml").async("string")).replace(/w:val="15"/, `w:val="${mode}"`));
    const f = path.join(os.tmpdir(), `dwb-word-render-c${mode}-${process.pid}.docx`);
    fs.writeFileSync(f, await zip.generateAsync({ type: "nodebuffer" }));
    await page.evaluate(() => { if (typeof setDirtyState === "function") setDirtyState(false); });
    await page.setInputFiles("#fileInput", f);
    await page.waitForFunction((m) => document.getElementById("loadingOverlay")?.classList.contains("hidden") && [...document.querySelectorAll(".docx-preview-host p")].some((p) => p.textContent === `Komorka ${m}`), mode, { timeout: 20000 });
    fs.rmSync(f, { force: true });
    await page.waitForTimeout(400);
    return page.evaluate((m) => {
      const ps = [...document.querySelectorAll(".docx-preview-host section.docx p")];
      const x = (txt) => { const rg = document.createRange(); rg.selectNodeContents(ps.find((p) => p.textContent === txt)); return rg.getBoundingClientRect().left; };
      return Math.round((x(`Komorka ${m}`) - x(`Tekst ${m}`)) * 100) / 100;
    }, mode);
  };
  const ind15 = await indentOf("15"), ind14 = await indentOf("14");
  // Word tryb 15: tekst komórki 5,75 pt = 7,67 px od marginesu (margines komórki + linia); u nas linia 1 px
  check("wcięcie tabeli, tryb 15: tekst komórki o margines komórki od marginesu (Word 7,67 px)", Math.abs(ind15 - 7.67) < 0.6, JSON.stringify({ ind15, ind14 }));
  check("wcięcie tabeli, tryb 14: tekst komórki na marginesie, linia przed nim (jak Word)", Math.abs(ind14) < 0.3, JSON.stringify({ ind15, ind14 }));

  // ── odstęp automatyczny (probe-autospacing.js) i podział strony na końcu (probe-trailing-break.js) ──
  const asSty = `<w:style w:type="paragraph" w:styleId="NormalWeb"><w:name w:val="Normal (Web)"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="auto"/></w:pPr></w:style>`;
  const asT = (txt) => `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r>${t(txt)}</w:r></w:p>`;
  const asW = (txt) => `<w:p><w:pPr><w:pStyle w:val="NormalWeb"/></w:pPr><w:r>${t(txt)}</w:r></w:p>`;
  const asCell = (c) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${c}</w:tc>`;
  const asTbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${tbB}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${asCell(asT("Wa1"))}${asCell(asW("Wc1"))}</w:tr><w:tr>${asCell(asT("Wa2"))}${asCell(asT("x"))}</w:tr></w:tbl>`;
  const f6 = path.join(os.tmpdir(), `dwb-word-render-as-${process.pid}.docx`);
  fs.writeFileSync(f6, await buildDocx({ body: [asT("Wprzed"), asW("Wb1"), asW("Wb2"), asT("Wpo"), asTbl, asT("Ostatni"), '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'], styles: asSty }));
  await page.evaluate(() => { if (typeof setDirtyState === "function") setDirtyState(false); });
  await page.setInputFiles("#fileInput", f6);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && [...document.querySelectorAll(".docx-preview-host p")].some((p) => p.textContent === "Wc1"), null, { timeout: 20000 });
  fs.rmSync(f6, { force: true });
  await page.waitForTimeout(400);
  const as = await page.evaluate(() => {
    const ps = [...document.querySelectorAll(".docx-preview-host section.docx p")];
    const top = (txt) => { const rg = document.createRange(); rg.selectNodeContents(ps.find((p) => p.textContent === txt)); return rg.getBoundingClientRect().top * 0.75; };
    return { body: top("Wb2") - top("Wb1"), row: top("Wa2") - top("Wa1"), pages: document.querySelector(".docx-preview-host").querySelectorAll(".docx-wrapper > section.docx").length, hosts: document.querySelectorAll(".docx-preview-host").length };
  });
  check("odstęp automatyczny (Normal (Web)) w treści: 14 pt (Word 27,5 pt na akapit)", Math.abs(as.body - 27.5) < 0.3, JSON.stringify(as));
  check("odstęp automatyczny w komórce: zdjęty na brzegach komórki (wiersz jak zwykły tekst, Word 14 pt)", Math.abs(as.row - 14) < 0.4, JSON.stringify(as));
  check("podział strony na samym końcu dokumentu: bez pustej strony (Word 1)", as.pages === 1, JSON.stringify(as));

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
