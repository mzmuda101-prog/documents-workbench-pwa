// find-word-playwright.js — panel Znajdź i zamień („Zamień wszystko”) jak w Wordzie.
// Oczekiwany tekst i format po zamianie = wynik prawdziwego Worda 16.114 dla tych samych plików
// (tools/word-oracle/run-find.js → tools/word-oracle/expected-find.json, 2026-10-17): bez
// „Uwzględnij wielkość liter” zamiana przejmuje wielkość liter trafienia (KOT → PIES, Kot → Pies),
// „całe wyrazy”: „_” rozdziela, cyfry nie, przy spacji w szukanym tekście opcja nie działa; prosty
// cudzysłów/apostrof znajduje drukarskie, spacja — twardą; polskie litery; zamiana bierze format
// pierwszego znaku trafienia. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");
const { CASES } = require("../tools/word-oracle/cases-find");
const { normalizeParas } = require("../tools/word-oracle/cases-quote");
const { savedParts } = require("../tools/word-oracle/app");
const EXPECTED = require("../tools/word-oracle/expected-find.json");

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
    page.on("dialog", (d) => d.accept());
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.locator("#fileInput").setInputFiles({ name: `${c.name}.docx`, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await buildDocx({ body: c.body }) });
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p") && typeof documentStructure !== "undefined" && documentStructure, null, { timeout: 30000 });
    await page.evaluate(async (c) => {
      await ensureLazyFeature("find-replace");
      searchQueryEl.value = c.find;
      frReplaceEl.value = c.replace;
      frRegexEl.checked = false;
      frMatchCaseEl.checked = !!c.mc;
      frWholeWordEl.checked = !!c.ww;
      await runFindReplaceScan();
      await replaceFrAll();
    }, c);
    const x = await savedParts(page);
    const runs = (await page.evaluate(normalizeParas, x)).map((p) => p.runs);
    const w = EXPECTED[c.name];
    check(`jak Word: ${c.name} [${c.find} → ${c.replace}${c.mc ? ", wielkość" : ""}${c.ww ? ", całe" : ""}] „${w.join(" ¶ ")}”`, JSON.stringify(runs) === JSON.stringify(w), runs.join(" ¶ "));
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
