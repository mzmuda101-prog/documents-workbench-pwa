// pdf-import-playwright.js — konwersja PDF → DOCX w aplikacji (app/pdf-*.js).
//
// PDF-y testowe NIE leżą w repo: generuje je Chromium (page.pdf) z HTML-a poniżej do katalogu
// tymczasowego — prawdziwe czcionki osadzone, polskie litery, tabela z tłem, lista, obraz,
// link, dwie kolumny, kropki spisu treści, stopka z numerem strony, 40-stronicowy „duży” plik.
// Sprawdza TREŚĆ pliku .docx (document.xml, rels, stopka), podgląd, zapis, Anuluj, płynność.
// Uruchom też: ENGINE=webkit (Safari/iPad) — PDF-y i tak robi Chromium.
//
// Opcjonalnie PDF_SAMPLES=katalog — dodatkowo konwertuje każdy *.pdf z katalogu (poza repo,
// np. próbki użytkownika) i sprawdza, że wychodzi dokument z tekstem/obrazem bez błędów.

const pw = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");
const ooxml = require("./ooxml-validate");
const converted = []; // DOCX-y z konwersji — na końcu walidator schematu Office (gdy jest .NET)

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "dwb-pdf-"));

const POLISH = "Zażółć gęślą jaźń — źdźbło łąki, ćma i żółw.";
const FIXTURE_HTML = `<!doctype html><html lang="pl"><head><meta charset="utf-8"><style>
  @page { size: A4; margin: 22mm 20mm 24mm 20mm; }
  body { font-family: Arial, Helvetica, sans-serif; font-size: 11pt; color: #111; }
  h1 { font-size: 22pt; margin: 0 0 10pt; }
  p.j { text-align: justify; line-height: 1.35; }
  table { border-collapse: collapse; width: 100%; margin: 10pt 0; }
  td, th { border: 0.8pt solid #333; padding: 3pt 5pt; font-size: 10pt; }
  th { background: #dbe9f7; text-align: left; }
  .cols { display: flex; gap: 28pt; margin-top: 12pt; }
  .cols > div { flex: 1; }
  .toc { display: flex; } .toc span.d { flex: 1; overflow: hidden; white-space: nowrap; }
  .pb { break-before: page; }
</style></head><body>
<h1>Protokół odbioru – próba</h1>
<p class="j">${POLISH} Ten akapit jest długi, żeby zawinął się na kilka wierszy i dał się rozpoznać jako jeden akapit wyjustowany: odbiór prac obejmuje sprawdzenie dokumentacji, pomiarów oraz zgodności z projektem wykonawczym, a wszelkie usterki zostają wpisane do załącznika numer jeden wraz z terminem usunięcia.</p>
<ul><li>Pierwszy punkt listy</li><li>Drugi punkt listy z polskimi znakami: ąę</li><li>Trzeci punkt</li></ul>
<table><tr><th>Pozycja</th><th>Ilość</th><th>Cena</th><th>Wartość</th></tr>
<tr><td>Kabel YDY 3×2,5</td><td>120 m</td><td>4,20 zł</td><td>504,00 zł</td></tr>
<tr><td>Puszka instalacyjna</td><td>15 szt.</td><td>1,10 zł</td><td>16,50 zł</td></tr>
<tr style="height: 30pt"><td></td><td></td><td></td><td></td></tr>
<tr><td colspan="3">Razem</td><td>520,50 zł</td></tr></table>
<p>Powierzchnia: 25 m<sup>2</sup>. Regulamin: <a href="https://example.com/regulamin">example.com/regulamin</a>.</p>
<p><img alt="logo" width="120" height="60" src="__IMG__"></p>
<div class="pb"></div>
<h1>Spis treści</h1>
<div class="toc"><span>Wprowadzenie</span><span class="d">${".".repeat(300)}</span><span>3</span></div>
<div class="toc"><span>Zakres prac</span><span class="d">${".".repeat(300)}</span><span>7</span></div>
<div class="cols"><div>${[1, 2, 3, 4, 5].map((i) => `<p>Kolumna A${i}: lewa kolumna tekstu, wiersz ${i}.</p>`).join("")}</div>
<div>${[1, 2, 3, 4, 5].map((i) => `<p>Kolumna B${i}: prawa kolumna tekstu, wiersz ${i}.</p>`).join("")}</div></div>
<h2>Wprowadzenie</h2>
<ol><li>Krok pierwszy<ul><li>szczegół kroku</li></ul></li><li>Krok drugi</li><li>Krok trzeci</li></ol>
<h2>Zakres prac</h2>
<p>Zakres obejmuje montaż i odbiór.</p>
</body></html>`;

const BIG_HTML = `<!doctype html><html lang="pl"><head><meta charset="utf-8"><style>
  @page { size: A4; margin: 20mm; } body { font-family: Georgia, serif; font-size: 11pt; } p { text-align: justify; }
</style></head><body>${Array.from({ length: 40 }, (_, k) => `<h2 style="break-before:page">Rozdział ${k + 1}</h2>${Array.from({ length: 9 }, (_, j) => `<p>Akapit ${j + 1} rozdziału ${k + 1}. ${POLISH} Treść wypełniająca stronę, żeby dokument miał realną objętość i sprawdzał płynność konwersji dłuższego pliku na urządzeniu.</p>`).join("")}`).join("")}</body></html>`;

async function makePdfs() {
  const browser = await pw.chromium.launch({ headless: true });
  const page = await browser.newPage();
  const img = await page.evaluate(() => {
    const c = document.createElement("canvas");
    c.width = 240; c.height = 120;
    const x = c.getContext("2d");
    const g = x.createLinearGradient(0, 0, 240, 0);
    g.addColorStop(0, "#1e6fd9"); g.addColorStop(1, "#f2a541");
    x.fillStyle = g; x.fillRect(0, 0, 240, 120);
    x.fillStyle = "#fff"; x.font = "bold 40px sans-serif"; x.fillText("DWB", 60, 75);
    return c.toDataURL("image/png");
  });
  await page.setContent(FIXTURE_HTML.replace("__IMG__", img), { waitUntil: "load" });
  const footer = '<div style="font-size:9pt;width:100%;text-align:center;font-family:Arial">Strona <span class="pageNumber"></span> z <span class="totalPages"></span></div>';
  await page.pdf({ path: path.join(TMP, "proba.pdf"), format: "A4", printBackground: true, displayHeaderFooter: true, headerTemplate: "<span></span>", footerTemplate: footer, margin: { top: "22mm", bottom: "24mm", left: "20mm", right: "20mm" } });
  await page.setContent(BIG_HTML, { waitUntil: "load" });
  await page.pdf({ path: path.join(TMP, "duzy.pdf"), format: "A4" });
  // lista tuż przy zdjęciu / zdjęcie lekko zachodzi na wiersze (karty terenu 101–103: mapka
  // zaczyna się 1 px za albo 3 px przed końcem najdłuższego wiersza)
  for (const [name, off] of [["styka", -1], ["zachodzi", 3]]) {
    await page.setContent(`<!doctype html><html lang="pl"><head><meta charset="utf-8"><style>body{margin:40px;font-family:Arial;font-size:16px}.row{display:flex;align-items:flex-start}.list{width:140px}.list p{margin:0 0 6px}img{width:260px;height:150px;display:block}</style></head><body><p>Karta testowa</p><div class="row"><div class="list"><p>KLATKI:</p><p>1/1 - 10</p><p>14/131 - 140</p><p>17/161 - 170</p></div><img src="__IMG__"></div><p>Akapit pod spodem.</p></body></html>`.replace("__IMG__", img), { waitUntil: "load" });
    await page.evaluate((off) => {
      const ps = [...document.querySelectorAll(".list p")];
      const r = document.createRange();
      const widest = Math.max(...ps.map((p) => { r.selectNodeContents(p); return r.getBoundingClientRect().right; }));
      const im = document.querySelector("img");
      im.style.marginLeft = `${widest - off - im.getBoundingClientRect().left}px`;
    }, off);
    await page.pdf({ path: path.join(TMP, `${name}.pdf`), width: "210mm", height: "148mm" });
  }
  // tekst zamieniony na krzywe (zgłoszenie „10. ZAGROŻENIE ATAKIEM…”): „o” = obrys + dziura w
  // JEDNEJ ścieżce (nonzero i evenodd), „i” = kreska + kwadracik kropki, kropka = sam kwadracik
  await page.setContent(`<!doctype html><html><body style="margin:0"><svg width="400" height="200" style="display:block">
    <path fill="#000" d="M50 50 a30 30 0 1 0 60 0 a30 30 0 1 0 -60 0 Z M65 50 a15 15 0 1 1 30 0 a15 15 0 1 1 -30 0 Z"/>
    <path fill="#000" fill-rule="evenodd" d="M50 140 a30 30 0 1 0 60 0 a30 30 0 1 0 -60 0 Z M65 140 a15 15 0 1 0 30 0 a15 15 0 1 0 -30 0 Z"/>
    <path fill="#000" d="M150 40 h3 v3 h-3 Z M150 50 h3 v30 h-3 Z"/>
    <path fill="#000" d="M170 77 h3 v3 h-3 Z"/>
    <path fill="#000" d="M200 50 C200 30 240 30 240 50 C240 70 200 70 200 50 Z M235 77 h3 v3 h-3 Z"/>
  </svg></body></html>`, { waitUntil: "load" });
  await page.pdf({ path: path.join(TMP, "krzywe.pdf"), width: "210mm", height: "297mm", margin: { top: 0, bottom: 0, left: 0, right: 0 } });
  // nieosadzone Times zwykły + pogrubiony (Word „bez osadzania”): pdf.js daje obu wspólne
  // loadedName „Times” — pogrubienie urywało się w środku zdania (informacje_ubezpieczenie…)
  fs.writeFileSync(path.join(TMP, "times-nieosadzony.pdf"), handPdf([
    "BT /F1 12 Tf 72 760 Td (Zwykly tekst na poczatku wiersza, ) Tj ET",
    "BT /F2 12 Tf 72 740 Td (POGRUBIONY srodek zdania ) Tj /F1 12 Tf (i znowu zwykly koniec.) Tj ET",
    "BT /F2 12 Tf 72 720 Td (Caly wiersz pogrubiony.) Tj ET",
  ], {
    F1: "<< /Type /Font /Subtype /TrueType /BaseFont /TimesNewRomanPSMT /Encoding /WinAnsiEncoding >>",
    F2: "<< /Type /Font /Subtype /TrueType /BaseFont /TimesNewRomanPS-BoldMT /Encoding /WinAnsiEncoding >>",
  }));
  // „skan”: zrzut strony z tekstem jako obraz, PDF z samym obrazem (bez warstwy tekstu)
  await page.setViewportSize({ width: 794, height: 1123 });
  await page.setContent(`<!doctype html><html lang="pl"><head><meta charset="utf-8"><style>body{margin:60px;font-family:Arial;font-size:22px;color:#111}h1{font-size:34px}.box{background:#f39a5b;padding:10px 16px}</style></head><body><div class="box"><b>KATEDRA AUTOMATYKI</b></div><h1>Lista obecności — lipiec</h1><p>Spóźnienia usprawiedliwione oraz nieusprawiedliwione.</p><p>Zażółć gęślą jaźń.</p></body></html>`, { waitUntil: "load" });
  const shot = await page.screenshot({ type: "jpeg", quality: 85 });
  await page.setContent(`<html><body style="margin:0"><img src="data:image/jpeg;base64,${shot.toString("base64")}" style="width:210mm;height:297mm;display:block"></body></html>`, { waitUntil: "load" });
  await page.pdf({ path: path.join(TMP, "skan.pdf"), width: "210mm", height: "297mm", margin: { top: 0, bottom: 0, left: 0, right: 0 }, printBackground: true });
  // Plan zajęć z arkusza: pozioma kreska nie dochodzi do wąskiej kolumny dni. PDF-owy
  // wykrywacz może więc potraktować dwa pola jako jedną komórkę. Pionowe nazwy dni muszą
  // mimo tego zostać osobnymi polami Worda, z natywnym kierunkiem tekstu (nie „czwartekpiątek”).
  await page.setContent(`<!doctype html><html lang="pl"><head><meta charset="utf-8"><style>
    @page { size: A4 landscape; margin: 14mm; }
    body { font-family: Arial, sans-serif; }
    table { border-collapse: collapse; width: 100%; table-layout: fixed; }
    td { border: 1.2pt solid #111; height: 90pt; text-align: center; font-size: 12pt; }
    .day { width: 20pt; padding: 0; writing-mode: vertical-rl; transform: rotate(180deg); font-weight: bold; }
    .day.top { border-bottom: 0; } .day.bottom { border-top: 0; }
  </style></head><body><table>
    <tr><td class="day top">czwartek</td><td>Laboratorium systemów energetycznych</td></tr>
    <tr><td class="day bottom">piątek</td><td>Projektowanie instalacji</td></tr>
  </table></body></html>`, { waitUntil: "load" });
  await page.pdf({ path: path.join(TMP, "plan-pionowe-dni.pdf"), format: "A4", landscape: true, printBackground: true });
  // skan tabeli obrócony o 180° + 3° (kartka odwrotnie i krzywo w skanerze) — prostowanie + tabela Worda
  await page.setViewportSize({ width: 794, height: 1123 });
  await page.setContent(`<!doctype html><html lang="pl"><head><meta charset="utf-8"><style>body{margin:70px;font-family:Arial;font-size:15px;color:#111}table{border-collapse:collapse;width:100%}td,th{border:1.5px solid #222;padding:8px 10px;text-align:left}th{background:#ddd}</style></head><body><h2>Protokół przekazania sprzętu</h2><table><tr><th>Nazwa</th><th>Numer seryjny</th><th>Stan</th></tr><tr><td>Wiertarka udarowa</td><td>WU-2231-778</td><td>sprawna</td></tr><tr><td>Szlifierka kątowa</td><td>SK-0912-114</td><td>uszkodzona osłona</td></tr><tr><td>Poziomica laserowa</td><td>PL-5520-031</td><td>sprawna</td></tr></table><p>Przekazał: Jan Kowalczyk. Odebrał: Zofia Wójcik.</p></body></html>`, { waitUntil: "load" });
  const tshot = (await page.screenshot({ type: "png" })).toString("base64");
  const rotated = await page.evaluate(async (src) => {
    const im = new Image();
    im.src = "data:image/png;base64," + src;
    await im.decode();
    const W = 1654, H = 2339; // ~200 dpi
    const c = Object.assign(document.createElement("canvas"), { width: W, height: H });
    const x = c.getContext("2d");
    x.fillStyle = "#fff"; x.fillRect(0, 0, W, H);
    x.translate(W / 2, H / 2); x.rotate(Math.PI + (3 * Math.PI) / 180);
    x.drawImage(im, -W / 2, -H / 2, W, H);
    return c.toDataURL("image/jpeg", 0.8);
  }, tshot);
  await page.setContent(`<html><body style="margin:0"><img src="${rotated}" style="width:210mm;height:297mm;display:block"></body></html>`, { waitUntil: "load" });
  await page.pdf({ path: path.join(TMP, "skan-tabela.pdf"), width: "210mm", height: "297mm", margin: { top: 0, bottom: 0, left: 0, right: 0 }, printBackground: true });
  // zdjęcia kartek telefonem: strona obrócona o kilka stopni na ciemnym blacie, ciepłe światło
  for (const [name, text, ang] of [["foto-1", "Zawiadomienie o zebraniu wspólnoty mieszkaniowej", 7], ["foto-2", "Porządek obrad i głosowanie nad uchwałami", -5]]) {
    await page.setViewportSize({ width: 794, height: 1123 });
    await page.setContent(`<!doctype html><html lang="pl"><head><meta charset="utf-8"><style>body{margin:80px;font-family:Arial;font-size:20px;color:#111}</style></head><body><h2>${text}</h2><p>Zebranie odbędzie się w środę o godzinie osiemnastej w świetlicy osiedlowej. Prosimy o punktualne przybycie.</p></body></html>`, { waitUntil: "load" });
    const shotP = (await page.screenshot({ type: "png" })).toString("base64");
    const photo = await page.evaluate(async ({ src, ang }) => {
      const im = new Image();
      im.src = "data:image/png;base64," + src;
      await im.decode();
      const c = Object.assign(document.createElement("canvas"), { width: 1500, height: 2000 });
      const x = c.getContext("2d");
      x.fillStyle = "#4a3424"; x.fillRect(0, 0, 1500, 2000);
      x.translate(750, 1000); x.rotate((ang * Math.PI) / 180);
      x.fillStyle = "#f3eee2"; x.fillRect(-560, -790, 1120, 1580);
      x.globalCompositeOperation = "multiply";
      x.drawImage(im, -560, -790, 1120, 1580);
      return c.toDataURL("image/jpeg", 0.85);
    }, { src: shotP, ang });
    fs.writeFileSync(path.join(TMP, `${name}.jpg`), Buffer.from(photo.split(",")[1], "base64"));
  }
  // tekst zamieniony na krzywe (Ghostscript -dNoOutputFonts) — gdy gs jest na komputerze
  await page.setContent(`<!doctype html><html lang="pl"><head><meta charset="utf-8"><style>body{margin:70px;font-family:Arial;font-size:16px}</style></head><body><h2>Zagrożenie atakiem terrorystycznym</h2><p>W razie zauważenia podejrzanego pakunku nie wolno go dotykać. Należy niezwłocznie powiadomić przełożonego i służby ochrony obiektu, a następnie oddalić się na bezpieczną odległość.</p></body></html>`, { waitUntil: "load" });
  await page.pdf({ path: path.join(TMP, "krzywe-zrodlo.pdf"), format: "A4" });
  try {
    require("child_process").execFileSync("gs", ["-q", "-o", path.join(TMP, "tekst-krzywe.pdf"), "-sDEVICE=pdfwrite", "-dNoOutputFonts", path.join(TMP, "krzywe-zrodlo.pdf")]);
  } catch (_) { /* brak gs — test pominięty */ }
  await browser.close();
  fs.writeFileSync(path.join(TMP, "zepsuty.pdf"), Buffer.from("%PDF-1.7\nto nie jest prawdziwy pdf\n%%EOF"));
}

// Najprostszy PDF pisany ręcznie (kroje NIEosadzone — Chromium zawsze osadza): treść strony
// z podanych linii, słownik krojów /F1, /F2… z podanych definicji.
function handPdf(lines, fonts) {
  const content = lines.join("\n");
  const objs = [];
  objs.push("<< /Type /Catalog /Pages 2 0 R >>");
  objs.push("<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
  const names = Object.keys(fonts);
  const fontRefs = names.map((n, i) => `/${n} ${5 + i} 0 R`).join(" ");
  objs.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << ${fontRefs} >> >> /Contents 4 0 R >>`);
  objs.push(`<< /Length ${Buffer.byteLength(content, "latin1")} >>\nstream\n${content}\nendstream`);
  for (const n of names) objs.push(fonts[n]);
  let out = "%PDF-1.4\n";
  const offs = [];
  objs.forEach((o, i) => { offs.push(Buffer.byteLength(out, "latin1")); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const xref = Buffer.byteLength(out, "latin1");
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n` + offs.map((o) => `${String(o).padStart(10, "0")} 00000 n \n`).join("");
  out += `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, "latin1");
}

const docxState = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(originalFileBytes);
  const doc = await z.file("word/document.xml").async("string");
  const rels = await z.file("word/_rels/document.xml.rels").async("string");
  const numbering = z.file("word/numbering.xml") ? await z.file("word/numbering.xml").async("string") : "";
  const files = Object.keys(z.files);
  const footerName = files.find((f) => /^word\/footer\d+\.xml$/.test(f));
  const footer = footerName ? await z.file(footerName).async("string") : "";
  const d = new DOMParser().parseFromString(doc, "application/xml");
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const paras = [...d.getElementsByTagNameNS(W, "p")].map((p) => [...p.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join(""));
  return {
    text: paras.join("\n"),
    paras,
    heading1: [...d.getElementsByTagNameNS(W, "pStyle")].filter((s) => s.getAttributeNS(W, "val") === "Heading1").length,
    tables: d.getElementsByTagNameNS(W, "tbl").length,
    // pusta rubryka tabeli: akapit na wysokość wiersza tekstu (nie 1 pt — kursor stawał na linii komórki)
    emptyCells: [...d.getElementsByTagNameNS(W, "tc")].filter((tc) => !tc.textContent.trim() && !tc.getElementsByTagNameNS(W, "tbl").length).map((tc) => Number(tc.getElementsByTagNameNS(W, "spacing")[0]?.getAttributeNS(W, "line") || 0)),
    shading: /<w:shd [^>]*w:fill="(?!auto)/.test(doc),
    media: files.filter((f) => f.startsWith("word/media/")).length,
    link: /example\.com\/regulamin/.test(rels) && /<w:hyperlink /.test(doc),
    pageBreaks: (doc.match(/<w:pageBreakBefore\/>/g) || []).length,
    a4: /<w:pgSz w:w="119\d\d" w:h="168\d\d"/.test(doc),
    footerPage: /PAGE/.test(footer) && /NUMPAGES/.test(footer),
    dotTab: /w:leader="dot"/.test(doc),
    sup: /<w:vertAlign w:val="superscript"\/>/.test(doc),
    hanging: /w:hanging="/.test(doc),
    // lista Worda (pdf-convert.js assignLists): numPr w akapicie, punktor w numbering.xml, w tekście bez „•”
    listDbg: (doc.match(/<w:p>(?:(?!<\/w:p>).)*(?:Krok|szczegół)(?:(?!<\/w:p>).)*<\/w:p>/g) || []).map((x) => `${(x.match(/<w:numPr>.*?<\/w:numPr>/) || ["-"])[0]} ${(x.match(/<w:t[^>]*>[^<]*/g) || []).map((y) => y.replace(/<w:t[^>]*>/, "")).join("|")}`).join(" ## ") + " || " + numbering.replace(/<w:rPr>.*?<\/w:rPr>/g, "").slice(-700),
    tocLinks: (doc.match(/<w:hyperlink w:anchor="_Toc\d+"/g) || []).length,
    tocBookmarks: /<w:bookmarkStart w:id="\d+" w:name="_Toc1"\/>(?:(?!<\/w:p>).)*Wprowadzenie/.test(doc),
    tocHeading: /<w:pStyle w:val="TOCHeading"\/>(?:(?!<\/w:p>).)*Spis treści/.test(doc),
    numberedLevels: /<w:numFmt w:val="decimal"\/><w:lvlText w:val="%1\."\/>/.test(numbering) && /<w:lvl w:ilvl="1"><w:start w:val="1"\/><w:numFmt w:val="bullet"\/>/.test(numbering) && /<w:ilvl w:val="1"\/>/.test(doc)
      && !(doc.match(/<w:p>(?:(?!<\/w:p>).)*(?:Krok|szczegół)(?:(?!<\/w:p>).)*<\/w:p>/g) || []).some((x) => /<w:t[^>]*>\s*(\d+\.|◦|•)/.test(x)),
    wordList: /<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="\d+"\/><\/w:numPr>/.test(doc) && /<w:numFmt w:val="bullet"\/><w:lvlText w:val="•"\/>/.test(numbering) && /numbering\.xml/.test(rels) && !/>•<\/w:t>/.test(doc),
  };
});

async function convert(page, file, { timeout = 60000 } = {}) {
  await page.evaluate(() => {
    window.__conv = null;
    window.__convErr = null;
    if (!window.__convHooked) {
      window.addEventListener("dwb:pdf-converted", (e) => (window.__conv = e.detail));
      window.__convHooked = true;
    }
  });
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction(() => window.__conv, null, { timeout });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.waitForTimeout(300);
  await keepDocx(page, path.basename(file, ".pdf"));
  return page.evaluate(() => window.__conv);
}

// bieżący dokument (wynik konwersji) do katalogu testu — na końcu walidator schematu
async function keepDocx(page, name) {
  const b64 = await page.evaluate(() => { const b = originalFileBytes; let s = ""; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000)); return btoa(s); });
  const out = path.join(TMP, `wynik-${converted.length + 1}-${name}.docx`);
  fs.writeFileSync(out, Buffer.from(b64, "base64"));
  converted.push(out);
}

async function run() {
  await makePdfs();
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());

  // --- 1. plik próbny
  const t0 = Date.now();
  const conv = await convert(page, path.join(TMP, "proba.pdf"));
  const ms = Date.now() - t0;
  check("konwersja kończy się otwarciem dokumentu", conv && conv.name === "proba.docx", JSON.stringify(conv?.name));
  check("czas konwersji 2 stron < 6 s", ms < 6000, `${ms} ms`);
  const ui = await page.evaluate(() => ({ dirty: hasUnsavedChanges, type: currentFileType, name: currentFileName, dialogOpen: !!document.getElementById("pdfConvDialog")?.open, handle: !!fileHandle }));
  check("nowy plik: niezapisany, bez uchwytu (Zapisz pyta o miejsce), okno postępu zamknięte", ui.dirty && ui.type === "docx" && !ui.handle && !ui.dialogOpen, JSON.stringify(ui));
  const st = await docxState(page);
  check("polskie litery bez strat", st.text.includes("Zażółć gęślą jaźń") && st.text.includes("źdźbło łąki"), st.text.slice(0, 120));
  check("tytuł jako Nagłówek 1 (nawigacja/struktura)", st.heading1 >= 1, `Heading1: ${st.heading1}`);
  check("akapit wyjustowany = jeden akapit (nie wiersz na akapit)", st.paras.some((p) => p.includes("Ten akapit jest długi") && p.includes("terminem usunięcia")), "");
  check("tabela z tłem nagłówka", st.tables >= 1 && st.shading, `tabele ${st.tables}`);
  check("pusty wiersz tabeli: 4 komórki z akapitem na wiersz tekstu (≥ 10 pt), nie 1 pt", st.emptyCells.filter((l) => l >= 200).length === 4, JSON.stringify(st.emptyCells));
  check("komórki tabeli: Ilość, 120 m, Razem, 520,50 zł", ["Ilość", "120 m", "Razem", "520,50 zł"].every((s) => st.paras.some((p) => p.trim() === s)), "");
  check("obraz w pliku", st.media >= 1, `media ${st.media}`);
  check("link zachowany jako hiperłącze", st.link, "");
  check("2 strony = podział strony + A4", st.pageBreaks >= 1 && st.a4, `pageBreakBefore ${st.pageBreaks}`);
  check("stopka „Strona N z M” jako pola PAGE/NUMPAGES", st.footerPage, "");
  check("stopka nie została w treści", !st.paras.some((p) => /^Strona \d+ z \d+$/.test(p.trim())), "");
  check("spis treści: tabulator z kropkami", st.dotTab, "");
  check("indeks górny (m²)", st.sup, "");
  check("spis treści: wpisy = linki do zakładek przy rozdziałach, „Spis treści” jako nagłówek spisu", st.tocLinks === 2 && st.tocBookmarks && st.tocHeading, JSON.stringify({ links: st.tocLinks, bm: st.tocBookmarks, head: st.tocHeading }));
  check("lista numerowana z zagnieżdżonym punktorem = numeracja Worda z poziomami (bez „1.” w tekście)", st.numberedLevels, st.listDbg);
  check("lista: prawdziwa lista Worda (punktor w definicji listy, nie w tekście) + wcięcie wiszące", st.wordList && st.hanging, "");
  const ia = st.paras.findIndex((p) => p.includes("Kolumna A5")), ib = st.paras.findIndex((p) => p.includes("Kolumna B1"));
  check("dwie kolumny czytane po kolei (cała lewa, potem prawa)", ia >= 0 && ib > ia, `A5@${ia} B1@${ib}`);
  const prev = await page.evaluate(() => ({
    tabs: document.querySelectorAll(".docx-tab-laid.tab-dot").length,
    imgs: document.querySelectorAll("section.docx img").length,
    sections: document.querySelectorAll("section.docx").length,
  }));
  check("podgląd: kropki spisu treści rozłożone, obraz widoczny, 2 kartki", prev.tabs >= 2 && prev.imgs >= 1 && prev.sections >= 2, JSON.stringify(prev));
  const saved = await page.evaluate(async () => {
    const bytes = await buildDocumentForSave();
    const z = await JSZip.loadAsync(bytes);
    return (await z.file("word/document.xml").async("string")).includes("Zażółć");
  });
  check("zapis .docx po konwersji działa", saved, "");

  // --- 1b. Excelowy plan: brak kreski w wąskiej kolumnie nie może złączyć pionowych dni
  await page.evaluate(() => { setDirtyState(false); });
  await convert(page, path.join(TMP, "plan-pionowe-dni.pdf"));
  const verticalDays = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(originalFileBytes);
    const doc = await z.file("word/document.xml").async("string");
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const xml = new DOMParser().parseFromString(doc, "application/xml");
    const cells = [...xml.getElementsByTagNameNS(W, "tc")].map((tc) => ({
      text: [...tc.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join(""),
      vertical: !!tc.getElementsByTagNameNS(W, "textDirection").length,
    }));
    return { cells, merged: /czwartek\s*piątek|piątek\s*czwartek/i.test(doc) };
  });
  const days = verticalDays.cells.filter((c) => /^(czwartek|piątek)$/i.test(c.text.trim()));
  check("plan zajęć: pionowe dni są osobnymi komórkami Worda, bez sklejania tekstu", !verticalDays.merged && days.length === 2 && days.every((c) => c.vertical), JSON.stringify(verticalDays));

  // --- 2. duży plik: płynność
  await page.evaluate(() => { setDirtyState(false); });
  await page.evaluate(() => {
    window.__maxGap = 0;
    window.__gapRun = true;
    let last = performance.now();
    const tick = () => {
      const now = performance.now();
      window.__maxGap = Math.max(window.__maxGap, now - last);
      last = now;
      if (window.__gapRun) requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  });
  const t1 = Date.now();
  const big = await convert(page, path.join(TMP, "duzy.pdf"), { timeout: 120000 });
  const bigMs = Date.now() - t1;
  const gap = await page.evaluate(() => { window.__gapRun = false; return Math.round(window.__maxGap); });
  check("40 stron: konwersja < 25 s", big && bigMs < 25000, `${bigMs} ms`);
  // najdłuższa przerwa obejmuje też render całego dokumentu przez docx-preview — sama konwersja oddaje wątek co stronę
  check("40 stron: apka nie zamarza (najdłuższa przerwa klatek < 1500 ms, z renderem dokumentu)", gap < 1500, `${gap} ms`);
  const bigSt = await docxState(page);
  check("40 stron: 40 rozdziałów jako nagłówki, 39 podziałów stron", bigSt.heading1 + (bigSt.text.match(/Rozdział \d+/g) || []).length >= 40 && bigSt.pageBreaks >= 39, `h1 ${bigSt.heading1}, pb ${bigSt.pageBreaks}`);

  // --- 3. Anuluj
  await page.evaluate(() => setDirtyState(false));
  const before = await page.evaluate(() => currentFileName);
  await page.evaluate(() => { window.__conv = null; });
  await page.setInputFiles("#fileInput", path.join(TMP, "duzy.pdf"));
  await page.waitForFunction(() => document.getElementById("pdfConvDialog")?.open, null, { timeout: 10000 }).catch(async (e) => {
    console.log("DBG", await page.evaluate(() => JSON.stringify({ open: document.getElementById("pdfConvDialog")?.open, phase: document.getElementById("pdfConvPhase")?.textContent, name: currentFileName, toasts: document.body.innerText.slice(-300) })));
    throw e;
  });
  await page.waitForTimeout(150);
  await page.click("#pdfConvCancel");
  await page.waitForFunction(() => !document.getElementById("pdfConvDialog")?.open, null, { timeout: 10000 });
  await page.waitForTimeout(500);
  const afterCancel = await page.evaluate(() => ({ conv: window.__conv, name: currentFileName, toast: document.body.innerText.includes("anulowano") }));
  check("Anuluj: przerywa, dokument bez zmian, komunikat", !afterCancel.conv && afterCancel.name === before && afterCancel.toast, JSON.stringify(afterCancel));

  // --- 4. zepsuty plik, potem znów działa
  await page.setInputFiles("#fileInput", path.join(TMP, "zepsuty.pdf"));
  await page.waitForFunction(() => !document.getElementById("pdfConvDialog")?.open && /poprawny plik PDF|przekonwertować/.test(document.body.innerText), null, { timeout: 15000 }).catch(() => {});
  const bad = await page.evaluate(() => ({ name: currentFileName, msg: /poprawny plik PDF|przekonwertować/.test(document.body.innerText) }));
  check("zepsuty PDF: czytelny błąd, dokument bez zmian", bad.msg && bad.name === before, JSON.stringify(bad));
  const again = await convert(page, path.join(TMP, "proba.pdf"));
  check("po błędzie kolejna konwersja działa", again && again.name === "proba.docx", "");

  // --- 4b. zdjęcie tuż przy liście / lekko na nią zachodzi: lista zostaje OBOK (nie spada pod nie)
  for (const name of ["styka", "zachodzi"]) {
    await page.evaluate(() => { setDirtyState(false); });
    await convert(page, path.join(TMP, `${name}.pdf`));
    const side = await page.evaluate(() => {
      const sec = document.querySelector("section.docx");
      const img = [...sec.querySelectorAll("img")].sort((a, b) => b.getBoundingClientRect().width - a.getBoundingClientRect().width)[0];
      const p = [...sec.querySelectorAll("p")].find((q) => q.textContent.includes("17/161 - 170"));
      if (!img || !p) return { img: !!img, paras: [...sec.querySelectorAll("p")].map((q) => q.textContent.trim()).filter(Boolean) };
      const a = img.getBoundingClientRect(), b = p.getBoundingClientRect();
      return { listBottom: Math.round(b.bottom), imgTop: Math.round(a.top), imgBottom: Math.round(a.bottom), listLeft: Math.round(b.left), imgLeft: Math.round(a.left) };
    });
    check(`zdjęcie ${name === "styka" ? "tuż przy liście" : "lekko zachodzi na listę"}: lista obok zdjęcia, nie pod nim`, side && side.listBottom <= side.imgBottom + 4 && side.imgLeft > side.listLeft, JSON.stringify(side));
  }

  // --- 4c. tekst jako krzywe: dziury w „o” zostają puste, kropki (kwadraciki) nie giną
  await page.evaluate(() => { setDirtyState(false); });
  await convert(page, path.join(TMP, "krzywe.pdf"));
  const curves = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(originalFileBytes);
    const name = Object.keys(z.files).find((f) => /^word\/media\/.+\.png$/.test(f));
    if (!name) return { media: false };
    const bmp = await createImageBitmap(await z.file(name).async("blob"));
    const c = document.createElement("canvas");
    c.width = bmp.width; c.height = bmp.height;
    const x = c.getContext("2d");
    x.drawImage(bmp, 0, 0);
    // punkty SVG w px → ułamek strony A4 (595,28 × 841,89 pt, 1 px = 0,75 pt)
    const ink = (px, py) => {
      const d = x.getImageData(Math.round((px * 0.75 / 595.28) * c.width), Math.round((py * 0.75 / 841.89) * c.height), 1, 1).data;
      return d[3] > 128 && d[0] < 128;
    };
    return {
      media: true,
      holeNonzero: !ink(80, 50), ringNonzero: ink(102, 50),
      holeEvenOdd: !ink(80, 140), ringEvenOdd: ink(102, 140),
      iDot: ink(151.5, 41.5), iStem: ink(151.5, 65), period: ink(171.5, 78.5), curveDot: ink(236.5, 78.5), curve: ink(220, 50),
    };
  });
  check("tekst jako krzywe: dziury w „o” puste (nonzero i evenodd), obrys widoczny", curves.media && curves.holeNonzero && curves.ringNonzero && curves.holeEvenOdd && curves.ringEvenOdd, JSON.stringify(curves));
  check("tekst jako krzywe: kropka nad „i”, kropka i kropka w ścieżce z krzywą nie giną", curves.iDot && curves.iStem && curves.period && curves.curveDot && curves.curve, JSON.stringify(curves));

  // --- 4d. nieosadzone Times zwykły + pogrubiony: pogrubienie dokładnie tam, gdzie w PDF
  await page.evaluate(() => { setDirtyState(false); });
  await convert(page, path.join(TMP, "times-nieosadzony.pdf"));
  const tb = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(originalFileBytes);
    const doc = await z.file("word/document.xml").async("string");
    const runs = [...doc.matchAll(/<w:r>(.*?)<\/w:r>/g)].map((m) => ({ b: m[1].includes("<w:b/>"), t: [...m[1].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((x) => x[1]).join("") })).filter((r) => r.t.trim());
    const isBold = (needle) => runs.filter((r) => r.t.includes(needle)).map((r) => r.b);
    return { zwykly: isBold("Zwykly tekst"), pogr: isBold("POGRUBIONY"), koniec: isBold("znowu zwykly"), caly: isBold("Caly wiersz"), runs: runs.map((r) => (r.b ? "**" : "") + r.t) };
  });
  check("nieosadzone kroje zwykły + pogrubiony: pogrubienie dokładnie tam, gdzie w PDF", tb.zwykly.join() === "false" && tb.pogr.join() === "true" && tb.koniec.join() === "false" && tb.caly.join() === "true", JSON.stringify(tb));
  // nazwy z Ghostscripta: „Wt” + nazwa + losowy ogon; „Demi” = pogrubienie
  const gsn = await page.evaluate(() => ["ELELGP+WtKnollTextDemi01cPX6!Bw", "WtUtSansMediumcPX6!Bw", "WtTimesBold01jcnDxdw", "NotoSansCJK-DemiLight"].map((n) => DWPdf.parseFontName(n)));
  check("nazwy krojów z Ghostscripta: rodzina bez śmieci, „Demi” pogrubione, „DemiLight” nie", gsn[0].family === "Knoll Text" && gsn[0].bold && gsn[1].family === "Ut Sans" && !gsn[1].bold && gsn[2].bold && !gsn[3].bold, JSON.stringify(gsn.map((g) => [g.family, g.bold])));

  // --- 5. skan: pytanie o OCR, rozpoznanie na urządzeniu, tekst + obraz tła bez słów
  await page.evaluate(() => { setDirtyState(false); window.__conv = null; });
  const t2 = Date.now();
  await page.setInputFiles("#fileInput", path.join(TMP, "skan.pdf"));
  await page.waitForSelector("#pdfConvOcrYes", { state: "visible", timeout: 30000 });
  check("skan: okno pyta o rozpoznanie tekstu", true, "");
  await page.click("#pdfConvOcrYes");
  await page.waitForFunction(() => window.__conv, null, { timeout: 120000 });
  await page.waitForTimeout(400);
  const ocrMs = Date.now() - t2;
  const ocr = await docxState(page);
  await keepDocx(page, "skan-ocr");
  const rep = await page.evaluate(() => window.__conv.report);
  check("skan: rozpoznany tekst z polskimi literami", /Lista obecności/.test(ocr.text) && /usprawiedliwione/.test(ocr.text) && /Zażółć/.test(ocr.text), ocr.text.replace(/\s+/g, " ").slice(0, 160));
  check("skan: obraz tła zostaje (kolory, linie), raport OCR", ocr.media >= 1 && rep.ocrPages === 1, JSON.stringify({ media: ocr.media, ocr: rep.ocrPages }));
  check("skan: OCR jednej strony < 25 s", ocrMs < 25000, `${ocrMs} ms`);
  // „Zostaw jako obraz”
  await page.evaluate(() => { setDirtyState(false); window.__conv = null; });
  await page.setInputFiles("#fileInput", path.join(TMP, "skan.pdf"));
  await page.waitForSelector("#pdfConvOcrNo", { state: "visible", timeout: 30000 });
  await page.click("#pdfConvOcrNo");
  await page.waitForFunction(() => window.__conv, null, { timeout: 60000 });
  await page.waitForTimeout(300);
  const noocr = await docxState(page);
  await keepDocx(page, "skan-obraz");
  check("skan: „Zostaw jako obraz” = sam obraz, bez tekstu", noocr.media >= 1 && !/Lista obecności/.test(noocr.text), noocr.text.slice(0, 80));

  // --- 5b. skan tabeli do góry nogami i krzywo: strona wyprostowana, tabela Worda z komórkami
  page.on("dialog", (d) => d.accept().catch(() => {}));
  const closeTabs = () => page.evaluate(() => { for (const d of dwbOpenDocs.list()) if (!d.active) dwbOpenDocs.closeTab(d.id); });
  await page.evaluate(() => { setDirtyState(false); window.__conv = null; });
  await closeTabs();
  await page.setInputFiles("#fileInput", path.join(TMP, "skan-tabela.pdf"));
  await page.waitForSelector("#pdfConvOcrYes", { state: "visible", timeout: 30000 });
  await page.click("#pdfConvOcrYes");
  await page.waitForFunction(() => window.__conv, null, { timeout: 120000 });
  await page.waitForTimeout(300);
  const st5 = await docxState(page);
  await keepDocx(page, "skan-tabela");
  const tbRep = await page.evaluate(() => window.__conv.report);
  check("skan do góry nogami i krzywo: strona wyprostowana, tekst rozpoznany", tbRep.straightened === 1 && /Protokół przekazania sprzętu/.test(st5.text) && /Szlifierka kątowa/.test(st5.text) && /Zofia Wójcik/.test(st5.text), JSON.stringify({ st: tbRep.straightened, t: st5.text.replace(/\s+/g, " ").slice(0, 160) }));
  // (numery seryjne mają pewność 87–92% — zależą od progu OCR, więc nie są tu sprawdzane)
  check("skan tabeli: prawdziwa tabela Worda, komórki osobno (Nazwa, Numer seryjny, uszkodzona osłona)", st5.tables >= 1 && ["Nazwa", "Numer seryjny", "uszkodzona osłona"].every((x) => st5.paras.some((p) => p.trim() === x)), JSON.stringify(st5.paras.filter((p) => p.trim()).slice(0, 14)));

  // --- 5c. tekst jako krzywe (bez znaków w PDF): pytanie o rozpoznanie, tekst zamiast kształtów
  if (fs.existsSync(path.join(TMP, "tekst-krzywe.pdf"))) {
    await page.evaluate(() => { setDirtyState(false); window.__conv = null; });
    await closeTabs();
    await page.setInputFiles("#fileInput", path.join(TMP, "tekst-krzywe.pdf"));
    await page.waitForSelector("#pdfConvOcrYes", { state: "visible", timeout: 30000 });
    const askTxt = await page.textContent("#pdfConvAskText");
    await page.click("#pdfConvOcrYes");
    await page.waitForFunction(() => window.__conv, null, { timeout: 120000 });
    await page.waitForTimeout(300);
    const cv = await docxState(page);
    await keepDocx(page, "tekst-krzywe");
    const cvRep = await page.evaluate(() => window.__conv.report);
    check("tekst-krzywe: pytanie mówi o tekście jako rysunku, rozpoznany tekst w dokumencie", /rysunek/.test(askTxt) && cvRep.curvePages === 1 && /Zagrożenie atakiem terrorystycznym/.test(cv.text) && /podejrzanego pakunku/.test(cv.text), JSON.stringify({ ask: askTxt.slice(0, 40), cp: cvRep.curvePages, t: cv.text.slice(0, 120) }));
    // kształty liter zastąpione tekstem: warstwa grafiki pusta (nic nie leży pod tekstem)
    check("tekst-krzywe: kształty liter usunięte (bez warstwy grafiki pod tekstem)", cv.media === 0, `obrazy: ${cv.media}`);
  }

  // --- 5d. dwa zdjęcia kartek naraz = JEDEN dokument, 2 strony (kartka wykryta i wyprostowana)
  await page.evaluate(() => { setDirtyState(false); window.__conv = null; });
  await closeTabs();
  const tabsBefore = await page.evaluate(() => dwbOpenDocs.list().length);
  await page.setInputFiles("#fileInput", [path.join(TMP, "foto-1.jpg"), path.join(TMP, "foto-2.jpg")]);
  await page.waitForSelector("#pdfConvOcrYes", { state: "visible", timeout: 60000 });
  const photoAsk = await page.textContent("#pdfConvAskText");
  await page.click("#pdfConvOcrYes");
  await page.waitForFunction(() => window.__conv, null, { timeout: 120000 });
  await page.waitForTimeout(400);
  const ph = await docxState(page);
  await keepDocx(page, "zdjecia");
  const phInfo = await page.evaluate(() => ({ name: currentFileName, tabs: dwbOpenDocs.list().length }));
  check("zdjęcia: pytanie mówi o zdjęciu dokumentu", /zdjęcie dokumentu/.test(photoAsk), photoAsk);
  check("dwa zdjęcia naraz: jeden dokument (jedna nowa karta), 2 strony, nazwa „foto-1 (2 str.).docx”", phInfo.tabs === tabsBefore + 1 && ph.pageBreaks >= 1 && /^foto-1 \(2 str\.\)\.docx$/.test(phInfo.name), JSON.stringify({ ...phInfo, pb: ph.pageBreaks, before: tabsBefore }));
  const phText = ph.text.replace(/\s+/g, " ");
  check("zdjęcia: tekst obu kartek rozpoznany (obrót i blat nie przeszkadzają)", /Zawiadomienie o zebraniu/.test(phText) && /Porządek obrad/.test(phText) && /świetlicy osiedlowej/.test(phText), phText.slice(0, 200));

  // --- 6. opcjonalnie: próbki spoza repo
  const dir = process.env.PDF_SAMPLES;
  if (dir && fs.existsSync(dir)) {
    // każda konwersja to nowa karta (open-docs.js) — zamknij poprzednie, limit to 12
    const closeOthers = () => page.evaluate(() => { for (const d of dwbOpenDocs.list()) if (!d.active) dwbOpenDocs.closeTab(d.id); });
    await closeOthers();
    for (const f of fs.readdirSync(dir).filter((n) => /\.pdf$/i.test(n))) {
      await page.evaluate(() => setDirtyState(false));
      await closeOthers();
      try {
        const auto = setInterval(() => page.click("#pdfConvOcrYes", { timeout: 200 }).catch(() => {}), 400);
        const r = await convert(page, path.join(dir, f), { timeout: 120000 }).finally(() => clearInterval(auto));
        const s = await docxState(page);
        check(`próbka ${f}: dokument z treścią`, r && (s.text.trim().length > 0 || s.media > 0), `${s.text.length} znaków, ${s.media} obrazów`);
      } catch (e) {
        check(`próbka ${f}`, false, String(e.message || e).slice(0, 160));
      }
    }
  }

  await browser.close();
  // DOCX z PDF-a tworzy w całości aplikacja — ma być zgodny ze schematem Office co do joty
  if (ooxml.available() && converted.length) {
    for (const v of ooxml.validate(converted)) check(`schemat Office: ${path.basename(v.file)}`, v.ok, v.error || `${v.errors.length}: ${v.errors.slice(0, 4).map(ooxml.short).join(" | ")}`);
  } else console.log("ℹ️  bez sprawdzania schematu (brak .NET — brew install dotnet)");
  if (!process.env.KEEP) fs.rmSync(TMP, { recursive: true, force: true }); // KEEP=1: PDF-y testu zostają (diagnoza)
  // ostrzeżenia pdf.js o brakujących czcionkach systemowych nie są błędami aplikacji
  const realErrors = errors.filter((e) => !/Warning|standardFontDataUrl|font/i.test(e));
  if (realErrors.length) check("bez błędów w konsoli", false, realErrors.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail && (!r.ok || process.env.V) ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ PDF → DOCX [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => {
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`));
  console.error(e);
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (_) { /* nic */ }
  process.exit(1);
});
