// doc-fonts-playwright.js — zamienniki krojów Office (assets/fonts/doc, scripts/gen-doc-fonts.py).
//
// Dokument w Calibri / Cambria / Arial / Times New Roman / Courier New / Georgia (każda odmiana:
// zwykła, pogrubiona, kursywa, pogrubiona kursywa) ma w podglądzie DOKŁADNIE szerokości z Worda —
// czy to z oryginału na urządzeniu, czy z naszego pliku. Wzorce policzone z plików Office
// (hmtx, kerning wyłączony jak w Wordzie). Bez zamienników Calibri szedł na Arial (+13%),
// a Cambria u Mateusza na Caladeę z Google Fonts (inne cyfry). Do tego: plik „ext” (greka)
// dociąga się tylko przy użyciu, pliki z listy precache w sw.js istnieją, a raz pobrany plik
// działa offline (osobny, niewersjonowany cache). ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const fs = require("fs");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ROOT = path.join(__dirname, "..");
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const TEXT = "Zażółć gęślą jaźń 0123456789 WAVE fi";
// szerokość TEXT przy 100 px: zwykła, pogrubiona, kursywa, pogrubiona kursywa (z plików Office)
const EXPECT = {
  Calibri: [1550.39, 1573.97, 1559.91, 1585.3],
  Cambria: [1664.11, 1772.36, 1622.95, 1736.18],
  Arial: [1756.59, 1817.53, 1756.59, 1817.53],
  "Times New Roman": [1630.03, 1674.85, 1591.6, 1630.66],
  "Courier New": [2160.35, 2160.35, 2160.35, 2160.35],
  Georgia: [1714.21, 1968.12, 1739.5, 1997.61],
};
const STYLES = [["", "zwykła"], ["<w:b/>", "pogrubiona"], ["<w:i/>", "kursywa"], ["<w:b/><w:i/>", "pogr. kursywa"]];

(async () => {
  // ── pliki z listy precache w sw.js istnieją, a każdy plik z CSS jest na dysku ──
  const sw = fs.readFileSync(path.join(ROOT, "sw.js"), "utf8");
  const famM = sw.match(/const DOC_FONT_CORE = \[([^\]]+)\]/);
  const fams = famM ? [...famM[1].matchAll(/"([a-z]+)"/g)].map((m) => m[1]) : [];
  const core = fams.flatMap((f) => ["r", "b", "i", "bi"].map((s) => `assets/fonts/doc/${f}-${s}-core-1.woff2`));
  check("sw.js: lista precache zamienników (24 pliki)", core.length === 24, core.length);
  const missing = core.filter((f) => !fs.existsSync(path.join(ROOT, f)));
  check("sw.js: wszystkie pliki z precache istnieją", !missing.length, missing.join(", "));
  const css = fs.readFileSync(path.join(ROOT, "styles", "app.css"), "utf8");
  const urls = [...new Set([...css.matchAll(/url\("\.\.\/(assets\/fonts\/doc\/[^"]+)"\)/g)].map((m) => m[1]))];
  const missingCss = urls.filter((u) => !fs.existsSync(path.join(ROOT, u)));
  check(`app.css: pliki z @font-face istnieją (${urls.length})`, urls.length >= 40 && !missingCss.length, missingCss.join(", "));

  const browser = await pw[ENGINE].launch();
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const errors = [];
  const page = await context.newPage();
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForFunction(() => !!originalFileBytes && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });

  // ── szerokości w podglądzie dokumentu ──
  const widths = await page.evaluate(async ({ TEXT, fams, styles }) => {
    const z = await JSZip.loadAsync(originalFileBytes);
    let doc = await z.file("word/document.xml").async("string");
    const esc = (s) => s.replace(/&/g, "&amp;");
    const run = (font, pr, txt) => `<w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>${pr}<w:sz w:val="40"/></w:rPr><w:t xml:space="preserve">${esc(txt)}</w:t></w:r>`;
    let body = "";
    for (const f of fams) for (const [pr] of styles) body += `<w:p>${run(f, pr, TEXT)}</w:p>`;
    body += `<w:p>${run("Calibri", "", "αβγ Ωμπ")}</w:p>`;
    doc = doc.replace("<w:body>", `<w:body>${body}`);
    z.file("word/document.xml", doc);
    await ingestFile(new File([await z.generateAsync({ type: "uint8array" })], "kroje.docx"), { silent: true });
    await document.fonts.ready;
    await new Promise((r) => setTimeout(r, 300));
    await document.fonts.ready;
    const spans = [...document.querySelectorAll(".docx-preview-host section.docx span")].filter((sp) => sp.textContent === TEXT);
    const greek = [...document.querySelectorAll(".docx-preview-host section.docx span")].find((sp) => sp.textContent.startsWith("αβγ"));
    return {
      list: spans.map((sp) => {
        const cs = getComputedStyle(sp);
        return { family: cs.fontFamily, w: (sp.getBoundingClientRect().width / parseFloat(cs.fontSize)) * 100 };
      }),
      greekFamily: greek && getComputedStyle(greek).fontFamily,
      extLoaded: [...document.fonts].some((f) => f.family.replace(/"/g, "") === "Calibri" && f.status === "loaded" && /-4FF/i.test(f.unicodeRange)),
    };
  }, { TEXT, fams: Object.keys(EXPECT), styles: STYLES });

  const famNames = Object.keys(EXPECT);
  check("podgląd: wszystkie fragmenty testowe są", widths.list.length === famNames.length * 4, widths.list.length);
  famNames.forEach((fam, fi) => {
    STYLES.forEach(([, label], si) => {
      const got = widths.list[fi * 4 + si];
      const want = EXPECT[fam][si];
      const rel = got ? Math.abs(got.w - want) / want : 1;
      check(`${fam} (${label}): szerokość jak w Wordzie`, rel < 0.003, got ? `${got.w.toFixed(2)} vs ${want} (${(rel * 100).toFixed(2)}%) · ${got.family}` : "brak");
    });
  });
  check("Calibri: rodzaj zapasowy dalej dopisany", /^"?Calibri"?, sans-serif$/.test(widths.list[0]?.family || ""), widths.list[0]?.family);
  check("greka w Calibri dociąga plik „ext”", widths.extLoaded, widths.greekFamily);

  // ── offline: raz pobrany zamiennik zostaje w osobnym cache i działa bez sieci ──
  await page.reload({ waitUntil: "load" });
  await page.evaluate(() => navigator.serviceWorker.ready);
  if (!(await page.evaluate(() => !!navigator.serviceWorker.controller))) {
    await page.reload({ waitUntil: "load" });
    await page.evaluate(() => navigator.serviceWorker.ready);
  }
  const url = "./assets/fonts/doc/tinos-i-core-1.woff2";
  const first = await page.evaluate(async (u) => {
    const r = await fetch(u);
    await r.arrayBuffer();
    await new Promise((res) => setTimeout(res, 300));
    const keys = await caches.keys();
    return { ok: r.ok, keys, cached: !!(await (await caches.open("docs-wb-docfonts-1")).match(u)) };
  }, url);
  check("SW: zamiennik trafia do niewersjonowanego cache", first.ok && first.cached, JSON.stringify(first.keys));
  // WebKit w Playwright: tryb offline blokuje też odpowiedzi service workera (nawet manifest.json
  // z cache) — tam wystarcza sprawdzenie cache wyżej.
  if (ENGINE === "chromium") {
    await context.setOffline(true);
    const offline = await page.evaluate(async (u) => {
      try {
        const ff = new FontFace("T_offline", `url(${u})`, { style: "italic" });
        await ff.load();
        return ff.status;
      } catch (e) { return String(e); }
    }, url);
    check("SW: zamiennik działa offline", offline === "loaded", offline);
    await context.setOffline(false);
  }

  check("brak błędów JS", !errors.length, errors.slice(0, 3).join(" | "));
  await browser.close();

  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`  ${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`);
  console.log(`\n${ENGINE}: ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
