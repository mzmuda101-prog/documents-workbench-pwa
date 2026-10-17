// word-stats-playwright.js — statystyka wyrazów jak w Wordzie (analysis.js documentWordStats).
// Oczekiwane liczby = wynik prawdziwego Worda 16.114 dla tych samych plików
// (tools/word-oracle/run-stats.js → tools/word-oracle/expected-stats.json, 2026-10-17):
// myślniki dzielą słowa, numery list się liczą, ukryty tekst nie, emoji = 1 znak, łamanie wiersza
// dzieli słowa, puste akapity się nie liczą, tabele tak. Plus pasek stanu i panel Statystyki.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");
const CASES = require("../tools/word-oracle/cases-stats");
const EXPECTED = require("../tools/word-oracle/expected-stats.json");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const errors = [];
  for (const c of CASES) {
    const page = await context.newPage();
    page.on("pageerror", (e) => errors.push(e.message));
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.locator("#fileInput").setInputFiles({ name: `${c.name}.docx`, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await buildDocx(c) });
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p") && typeof documentStructure !== "undefined" && documentStructure, null, { timeout: 30000 });
    await page.waitForTimeout(200);
    const a = await page.evaluate(async () => {
      const s = documentWordStats(docCanvasEl);
      await ensureLazyFeature("stats");
      const panel = computeDocumentStats().totals;
      return { ...s, bar: analyzeDocumentDom(docCanvasEl).words, panelWords: panel.words, panelChars: panel.chars, panelNoSp: panel.charsNoSpaces };
    });
    const w = EXPECTED[c.name];
    const same = ["words", "chars", "charsWithSpaces", "paragraphs"].every((k) => a[k] === w[k]) && a.bar === w.words && a.panelWords === w.words && a.panelChars === w.charsWithSpaces && a.panelNoSp === w.chars;
    check(`jak Word: ${c.name} (${w.words} sł., ${w.chars}/${w.charsWithSpaces} zn., ${w.paragraphs} akap.)`, same, JSON.stringify(a));
    await page.close();
  }
  check("brak błędów strony", errors.length === 0, errors.join(" | "));
  await browser.close();
  const bad = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`));
  console.log(`\n[${ENGINE}] ${results.length - bad.length}/${results.length} OK`);
  process.exit(bad.length ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
