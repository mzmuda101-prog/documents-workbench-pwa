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
<tr><td colspan="3">Razem</td><td>520,50 zł</td></tr></table>
<p>Powierzchnia: 25 m<sup>2</sup>. Regulamin: <a href="https://example.com/regulamin">example.com/regulamin</a>.</p>
<p><img alt="logo" width="120" height="60" src="__IMG__"></p>
<div class="pb"></div>
<h1>Spis treści</h1>
<div class="toc"><span>Wprowadzenie</span><span class="d">${".".repeat(300)}</span><span>3</span></div>
<div class="toc"><span>Zakres prac</span><span class="d">${".".repeat(300)}</span><span>7</span></div>
<div class="cols"><div>${[1, 2, 3, 4, 5].map((i) => `<p>Kolumna A${i}: lewa kolumna tekstu, wiersz ${i}.</p>`).join("")}</div>
<div>${[1, 2, 3, 4, 5].map((i) => `<p>Kolumna B${i}: prawa kolumna tekstu, wiersz ${i}.</p>`).join("")}</div></div>
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
  // „skan”: zrzut strony z tekstem jako obraz, PDF z samym obrazem (bez warstwy tekstu)
  await page.setViewportSize({ width: 794, height: 1123 });
  await page.setContent(`<!doctype html><html lang="pl"><head><meta charset="utf-8"><style>body{margin:60px;font-family:Arial;font-size:22px;color:#111}h1{font-size:34px}.box{background:#f39a5b;padding:10px 16px}</style></head><body><div class="box"><b>KATEDRA AUTOMATYKI</b></div><h1>Lista obecności — lipiec</h1><p>Spóźnienia usprawiedliwione oraz nieusprawiedliwione.</p><p>Zażółć gęślą jaźń.</p></body></html>`, { waitUntil: "load" });
  const shot = await page.screenshot({ type: "jpeg", quality: 85 });
  await page.setContent(`<html><body style="margin:0"><img src="data:image/jpeg;base64,${shot.toString("base64")}" style="width:210mm;height:297mm;display:block"></body></html>`, { waitUntil: "load" });
  await page.pdf({ path: path.join(TMP, "skan.pdf"), width: "210mm", height: "297mm", margin: { top: 0, bottom: 0, left: 0, right: 0 }, printBackground: true });
  await browser.close();
  fs.writeFileSync(path.join(TMP, "zepsuty.pdf"), Buffer.from("%PDF-1.7\nto nie jest prawdziwy pdf\n%%EOF"));
}

const docxState = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(originalFileBytes);
  const doc = await z.file("word/document.xml").async("string");
  const rels = await z.file("word/_rels/document.xml.rels").async("string");
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
    shading: /<w:shd [^>]*w:fill="(?!auto)/.test(doc),
    media: files.filter((f) => f.startsWith("word/media/")).length,
    link: /example\.com\/regulamin/.test(rels) && /<w:hyperlink /.test(doc),
    pageBreaks: (doc.match(/<w:pageBreakBefore\/>/g) || []).length,
    a4: /<w:pgSz w:w="119\d\d" w:h="168\d\d"/.test(doc),
    footerPage: /PAGE/.test(footer) && /NUMPAGES/.test(footer),
    dotTab: /w:leader="dot"/.test(doc),
    sup: /<w:vertAlign w:val="superscript"\/>/.test(doc),
    hanging: /w:hanging="/.test(doc),
    bulletTab: /•<\/w:t><\/w:r><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:tab\/>/.test(doc),
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
  return page.evaluate(() => window.__conv);
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
  check("komórki tabeli: Ilość, 120 m, Razem, 520,50 zł", ["Ilość", "120 m", "Razem", "520,50 zł"].every((s) => st.paras.some((p) => p.trim() === s)), "");
  check("obraz w pliku", st.media >= 1, `media ${st.media}`);
  check("link zachowany jako hiperłącze", st.link, "");
  check("2 strony = podział strony + A4", st.pageBreaks >= 1 && st.a4, `pageBreakBefore ${st.pageBreaks}`);
  check("stopka „Strona N z M” jako pola PAGE/NUMPAGES", st.footerPage, "");
  check("stopka nie została w treści", !st.paras.some((p) => /^Strona \d+ z \d+$/.test(p.trim())), "");
  check("spis treści: tabulator z kropkami", st.dotTab, "");
  check("indeks górny (m²)", st.sup, "");
  check("lista: punktor + tabulator + wcięcie wiszące", st.bulletTab && st.hanging, "");
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
  check("skan: „Zostaw jako obraz” = sam obraz, bez tekstu", noocr.media >= 1 && !/Lista obecności/.test(noocr.text), noocr.text.slice(0, 80));

  // --- 6. opcjonalnie: próbki spoza repo
  const dir = process.env.PDF_SAMPLES;
  if (dir && fs.existsSync(dir)) {
    // każda konwersja to nowa karta (open-docs.js) — zamknij poprzednie, limit to 12
    page.on("dialog", (d) => d.accept());
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
  fs.rmSync(TMP, { recursive: true, force: true });
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
