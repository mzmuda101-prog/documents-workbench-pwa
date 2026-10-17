// table-word-playwright.js — scalanie / dzielenie komórek / wiersz nagłówka jak w Wordzie.
// Oczekiwana tabela po operacji = to, co zapisuje Word 16.114 (tools/word-oracle/run-tables.js →
// tools/word-oracle/expected-tables.json, 2026-10-17): struktura (gridSpan, vMerge, tblHeader),
// treść (pusta komórka nic nie wnosi, niepusta — wszystkie akapity), szerokości siatki i komórek
// (Word usuwa z siatki granice kolumn, których żaden wiersz już nie używa). ENGINE=webkit.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");
const { CASES, normalizeTable } = require("../tools/word-oracle/cases-tables");
const { appEdit, docOf } = require("../tools/word-oracle/run-tables");
const { savedParts } = require("../tools/word-oracle/app");
const EXPECTED = require("../tools/word-oracle/expected-tables.json");

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
    await page.locator("#fileInput").setInputFiles({ name: `${c.name}.docx`, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await buildDocx(docOf(c)) });
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host td p"), null, { timeout: 30000 });
    await page.evaluate(() => { if (readOnlyMode) appFrame.setReadOnly(false); });
    const n = await page.evaluate((e) => applyDocumentEdit(e), appEdit(c));
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 });
    const a = await page.evaluate(normalizeTable, await savedParts(page));
    const w = EXPECTED[c.name];
    check(`jak Word: ${c.name} → ${w.rows.map((r) => `${r.h ? "⟦H⟧" : ""}${r.cells.join(" | ")}`).join(" / ")}`, n > 0 && JSON.stringify(a) === JSON.stringify(w), JSON.stringify(a));
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
