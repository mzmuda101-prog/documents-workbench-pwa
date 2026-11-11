// ocr-layout-playwright.js — strażnik poprawek z 2026-11-11 (infografiki, paragony, długie zrzuty):
//   1. układ strony na sztucznych znakach (szybko, bez OCR) — DWPdf.layoutPage + treeToBlocks:
//      - kreska pod tytułem ramki / krawędź wiersza dłuższa od tekstu ≠ podkreślenie,
//        kreska dokładnie pod słowem = podkreślenie,
//      - ramki po bokach dużego tytułu nie stają się jego indeksem górnym/dolnym
//        (litery „LEGENDA” i „Pełna klatka” przeplatały się: „FeLlnEaG”),
//      - lista „towar … cena” = wiersze z tabulatorem, nie dwie kolumny (najpierw wszystkie
//        nazwy, potem wszystkie ceny); dwie kolumny artykułu zostają kolumnami;
//   2. długa strona (PDF 210 × 1400 mm i zrzut ekranu PNG 1080 × ~5800 px z OCR):
//      strony w pliku nie wyższe niż 22″ (limit Worda), każdy akapit dokładnie raz,
//      obraz zrzutu bez zmniejszania (dawniej 3200 px dłuższego boku → 490 px szerokości).
// Uruchom też: ENGINE=webkit.

const pw = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "dwb-ocr-layout-"));

async function makeFiles() {
  const b = await pw.chromium.launch({ headless: true });
  const page = await b.newPage();
  const paras = Array.from({ length: 60 }, (_, i) => `<p>Akapit ${i + 1}: jesienią warto zadbać o ogród, bo od tego zależy, jak przetrwa mrozy i wiosnę.</p>`).join("");
  await page.setContent(`<!doctype html><html lang="pl"><head><meta charset="utf-8"><style>@page{size:210mm 1400mm;margin:15mm} body{font-family:Arial;font-size:12pt} p{margin:0 0 9pt}</style></head><body>${paras}</body></html>`);
  await page.pdf({ path: path.join(TMP, "dlugi.pdf"), width: "210mm", height: "1400mm", printBackground: true });
  // zrzut ekranu telefonu: 1080 px szerokości, tekst 30 px, 80 akapitów (~5800 px wysokości)
  await page.setViewportSize({ width: 540, height: 800 });
  await page.setContent(`<!doctype html><html lang="pl"><head><meta charset="utf-8"><style>body{margin:0;padding:20px;font-family:Arial;font-size:15px;line-height:1.4;background:#fff;color:#111} p{margin:0 0 14px}</style></head><body>${Array.from({ length: 80 }, (_, i) => `<p>Akapit ${i + 1}: pamiętaj o rękawicach i ostrym sekatorze w ogrodzie.</p>`).join("")}</body></html>`);
  const ctx2 = await b.newContext({ viewport: { width: 540, height: 800 }, deviceScaleFactor: 2 });
  const p2 = await ctx2.newPage();
  await p2.setContent(await page.content());
  await p2.screenshot({ path: path.join(TMP, "zrzut.png"), fullPage: true });
  await b.close();
}

async function run() {
  await makeFiles();
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept().catch(() => {}));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  // moduły konwersji ładują się dopiero przy pierwszym PDF-ie — tu wprost, po kolei
  await page.evaluate(async () => {
    for (const src of ["app/pdf-extract.js", "app/pdf-layout.js", "app/pdf-docx.js", "app/pdf-fonts.js", "app/pdf-convert.js"]) {
      if (document.querySelector(`script[src^="${src}"]`)) continue;
      await new Promise((ok, bad) => { const s = document.createElement("script"); s.src = src; s.onload = ok; s.onerror = bad; document.head.appendChild(s); });
    }
  });
  await page.waitForFunction(() => window.DWPdf?.layoutPage && window.DWPdf?.treeToBlocks, null, { timeout: 30000 });

  // ── 1. układ strony na sztucznych znakach ──
  const lay = await page.evaluate(() => {
    const NS = window.DWPdf;
    const font = { id: "t", family: "Arial", bold: false, italic: false, mono: false, serif: false, symbolic: false, asc: 0.72, desc: 0.21 };
    let order = 0;
    const text = (s, x, y, size) => {
      const out = [];
      for (const u of s) {
        const w = NS.helvWidth(u) * size;
        out.push({ u, x, x1: x + w, y, adv: w, size, asc: size * 0.72, desc: size * 0.21, angle: 0, font, fontRef: "F1", isSpace: u === " ", color: "000000", bold: false, italic: false, order: order++ });
        x += w;
      }
      return out;
    };
    const line = (x0, x1, y, width = 0.8) => ({ x0, x1, y0: y, y1: y, width, color: "000000", order: order++ });
    const end = (gl) => gl[gl.length - 1].x1;
    const layout = (glyphs, lines = [], W = 595, H = 842) => {
      const L = NS.layoutPage({ width: W, height: H, glyphs, lines, rects: [], graphics: [], images: [], annotations: [] }, {});
      const paras = [];
      // akapity także w kolumnach („columns”: cells[].blocks) i tabelach (rows[].cells[].blocks)
      const walk = (bs) => { for (const b of bs || []) { if (b.type === "para") paras.push({ ...b, col: b.col }); for (const c of b.cells || []) walk(c.blocks); for (const r of b.rows || []) for (const c of r.cells || []) walk(c.blocks); if (b.blocks) walk(b.blocks); } };
      walk(NS.treeToBlocks(L.tree, { x0: 20, x1: W - 20, top: 20 }));
      return paras.map((p) => ({ text: p.tokens.map((t) => (t.type === "text" ? t.text : t.type === "tab" ? "\t" : "")).join(""), under: p.tokens.filter((t) => t.type === "text" && t.underline).map((t) => t.text).join("|"), vert: p.tokens.some((t) => t.vert) }));
    };
    const res = {};
    // (a) tytuł ramki z kreską na całą szerokość ramki tuż pod linią bazową, (b) „Model … 2001”
    // z kreską pod całym wierszem, (c) prawdziwe podkreślenie dokładnie pod słowem
    const a = text("Lustrzanki profesjonalne", 60, 100, 12);
    const b1 = text("Model pierwszy", 60, 140, 10), b2 = text("2001", 300, 140, 10);
    const c1 = text("Wybierz ", 60, 200, 11), c2 = text("tańszą serię", end(c1), 200, 11), c3 = text(" albo inną.", 0, 0, 11);
    const c3x = text(" albo inną.", end(c2), 200, 11);
    void c3;
    res.under = layout([...a, ...b1, ...b2, ...c1, ...c2, ...c3x], [line(60, 330, 101.5), line(60, 330, 141.5), line(c2[0].x, end(c2), 201.5)]);
    // duży tytuł na środku, po bokach mniejsze ramki trochę wyżej i trochę niżej
    order = 0;
    const title = text("MAPA RODZINY APARATÓW", 200, 30, 17);
    const leg = text("LEGENDA IKONEK", 30, 22, 7.5), legR = text("OZNACZENIA", 500, 22, 7.5);
    const item = text("Pełna klatka", 30, 37, 6.5), itemR = text("Nowy model", 500, 37, 6.5);
    res.sides = layout([...title, ...leg, ...legR, ...item, ...itemR]);
    // paragon: towar … cena w 8 wierszach (prawa strona dużo węższa)
    order = 0;
    const rec = [];
    const names = ["Chleb żytni", "Masło ekstra", "Mleko świeże", "Jabłka polskie", "Ser żółty", "Kawa ziarnista", "Herbata czarna", "Woda mineralna"];
    names.forEach((n, i) => rec.push(...text(n, 20, 40 + i * 16, 10), ...text(((i + 3) * 1.37).toFixed(2).replace(".", ","), 190, 40 + i * 16, 10)));
    res.receipt = layout(rec, [], 227, 400);
    // dwie kolumny artykułu na tej samej siatce wierszy — mają zostać kolumnami
    order = 0;
    const cols = [];
    for (let i = 0; i < 8; i++) {
      cols.push(...text(`Lewa kolumna artykułu, wiersz numer ${i + 1} tekstu`, 40, 100 + i * 14, 10));
      cols.push(...text(`Prawa kolumna artykułu, wiersz numer ${i + 1} tekstu`, 320, 100 + i * 14, 10));
    }
    res.cols = layout(cols);
    return res;
  });
  const u = lay.under;
  check("kreska pod tytułem ramki (dłuższa od tekstu) — bez podkreślenia", !u.some((p) => /Lustrzanki/.test(p.under)), JSON.stringify(u));
  check("kreska pod wierszem „Model … 2001” (dziura w środku) — bez podkreślenia", !u.some((p) => /Model|2001/.test(p.under)), JSON.stringify(u));
  check("kreska dokładnie pod „tańszą serię” — podkreślenie tylko tych słów", u.some((p) => p.under.replace(/\|/g, "").trim() === "tańszą serię"), JSON.stringify(u.map((p) => p.under)));
  const s = lay.sides;
  check("ramki po bokach dużego tytułu: bez indeksów, litery się nie przeplatają", !s.some((p) => p.vert) && s.some((p) => /LEGENDA IKONEK/.test(p.text)) && s.some((p) => /Pełna klatka/.test(p.text)) && s.some((p) => /MAPA RODZINY APARATÓW/.test(p.text)), JSON.stringify(s));
  const r = lay.receipt;
  check("paragon: towar i cena w jednym wierszu (tabulator), nie kolumna nazw i kolumna cen", ["Chleb żytni\t4,11", "Woda mineralna\t13,70"].every((x) => r.some((p) => p.text.includes(x))), JSON.stringify(r.map((p) => p.text)));
  const c = lay.cols;
  check("dwie kolumny artykułu (podobna szerokość) zostają kolumnami", c.some((p) => /Lewa kolumna/.test(p.text) && !/Prawa/.test(p.text)) && !c.some((p) => /Lewa.*\t.*Prawa/.test(p.text)), JSON.stringify(c.map((p) => p.text).slice(0, 4)));

  // ── 2. długie strony ──
  const convert = async (file) => {
    await page.evaluate(() => { if (typeof setDirtyState === "function") setDirtyState(false); window.__conv = null; if (!window.__h) { window.addEventListener("dwb:pdf-converted", (e) => (window.__conv = e.detail)); window.__h = 1; } for (const d of dwbOpenDocs.list()) if (!d.active) dwbOpenDocs.closeTab(d.id); });
    await page.setInputFiles("#fileInput", file);
    const auto = setInterval(() => page.click("#pdfConvOcrYes", { timeout: 200 }).catch(() => {}), 300);
    try { await page.waitForFunction(() => window.__conv, null, { timeout: 180000 }); } finally { clearInterval(auto); }
    return page.evaluate(async () => {
      const z = await JSZip.loadAsync(originalFileBytes);
      const doc = await z.file("word/document.xml").async("string");
      const heights = [...doc.matchAll(/<w:pgSz [^>]*w:h="(\d+)"/g)].map((m) => +m[1]);
      const text = doc.replace(/<w:tab\/>/g, " ").replace(/<\/w:p>/g, "\n").replace(/<[^>]+>/g, "");
      const nums = [...text.matchAll(/Akapit (\d+)/g)].map((m) => +m[1]);
      const widths = [];
      for (const f of Object.keys(z.files).filter((n) => /^word\/media\/.+\.(png|jpe?g)$/.test(n))) {
        const bmp = await createImageBitmap(new Blob([await z.file(f).async("uint8array")]));
        widths.push(bmp.width);
        bmp.close();
      }
      return { heights, nums, widths, report: window.__conv.report };
    });
  };
  const counts = (nums, n) => Array.from({ length: n }, (_, i) => nums.filter((v) => v === i + 1).length);
  const d = await convert(path.join(TMP, "dlugi.pdf"));
  // strony o tym samym rozmiarze dzielą sekcję Worda — liczba stron z raportu konwersji
  check("PDF 1400 mm: kilka stron, każda ≤ 22″ (31680 tw)", d.report.pages >= 2 && d.heights.every((h) => h <= 31680), JSON.stringify({ pages: d.report.pages, h: d.heights }));
  const dc = counts(d.nums, 60);
  check("PDF 1400 mm: każdy z 60 akapitów dokładnie raz (cięcie nie gubi i nie dubluje)", dc.every((k) => k === 1), JSON.stringify(dc));
  const z = await convert(path.join(TMP, "zrzut.png"));
  check("zrzut 1080 px: kilka stron, każda ≤ 22″", z.report.pages >= 2 && z.heights.every((h) => h <= 31680), JSON.stringify({ pages: z.report.pages, h: z.heights }));
  const zc = counts(z.nums, 80);
  check("zrzut: rozpoznane akapity (≥ 72 z 80), żaden dwa razy", zc.filter((k) => k === 1).length >= 72 && zc.every((k) => k <= 1), JSON.stringify(zc));
  check("zrzut: obraz bez zmniejszania (≥ 1000 px szerokości)", z.widths.length && Math.min(...z.widths) >= 1000, JSON.stringify(z.widths));

  await browser.close();
  fs.rmSync(TMP, { recursive: true, force: true });
  if (errors.length) check("bez błędów w konsoli", false, errors.join(" | ").slice(0, 400));
  let failed = 0;
  for (const x of results) {
    console.log(`${x.ok ? "✅" : "❌"} ${x.name}${x.detail && (!x.ok || process.env.V) ? `  (${x.detail})` : ""}`);
    if (!x.ok) failed++;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ OCR / układ / długie strony [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => {
  results.forEach((x) => console.log(`${x.ok ? "✅" : "❌"} ${x.name}${!x.ok && x.detail ? `  (${x.detail})` : ""}`));
  console.error(e);
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (_) { /* nic */ }
  process.exit(1);
});
