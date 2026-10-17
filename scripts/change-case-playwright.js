// change-case-playwright.js — „Zmień wielkość liter” (Aa) jak w Wordzie (changeTextCase).
// Oczekiwane teksty = wynik prawdziwego Worda 16.114 (tools/word-oracle/run-case.js →
// tools/word-oracle/expected-case.json, 218 przypadków, 2026-10-17): 5 trybów, całe akapity i
// zaznaczenia od środka słowa, skrótowce (USA, 3D) zostają, granica formatowania nie dzieli słowa.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");
const CASES = require("../tools/word-oracle/expected-case.json");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const RUNS = { raw: '<w:r><w:t>red</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold text</w:t></w:r>' };

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const buffer = await buildDocx({ body: CASES.map((c) => (c.runs ? RUNS : c.text)) });
  await page.locator("#fileInput").setInputFiles({ name: "case.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.evaluate(() => { if (readOnlyMode) appFrame.setReadOnly(false); });
  await page.waitForTimeout(400);
  const got = await page.evaluate((list) => list.map((o, i) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    const r = formDomRange(p, o.from, o.to == null ? p.textContent.length : o.to);
    getSelection().removeAllRanges(); getSelection().addRange(r);
    changeTextCase(o.mode);
    return { text: p.textContent.replace(/﻿/g, ""), bold: p.querySelector("strong, b, [style*='font-weight']")?.textContent || "" };
  }), CASES);
  const bad = CASES.map((c, i) => [c, got[i]]).filter(([c, g]) => c.word.replace(/\u0009/g, "\t") !== g.text);
  bad.slice(0, 15).forEach(([c, g]) => console.log(`❌ ${c.mode} [${c.from}…${c.to ?? "koniec"}] „${c.text}”\n   Word: „${c.word}”\n   apka: „${g.text}”`));
  const runsKept = got.filter((g, i) => CASES[i].runs).every((g) => g.bold && g.bold.length === "bold text".length);
  console.log(`${runsKept ? "✅" : "❌"} pogrubiony fragment zostaje pogrubiony (zmienia się tylko tekst)`);
  console.log(`${errors.length ? "❌" : "✅"} brak błędów strony ${errors.join(" | ")}`);
  console.log(`\n[${ENGINE}] ${CASES.length - bad.length}/${CASES.length} jak w Wordzie`);
  await browser.close();
  process.exit(bad.length || !runsKept || errors.length ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
