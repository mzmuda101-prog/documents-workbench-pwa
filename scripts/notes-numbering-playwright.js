// notes-numbering-playwright.js — numeracja przypisów dolnych/końcowych jak w Wordzie (doc-notes.js).
// Oczekiwane numery = to, co rysuje prawdziwy Word 16.114 dla tych samych plików (PDF z Worda,
// tools/word-oracle/run-notes.js → tools/word-oracle/expected-notes.json, 2026-10-17):
// liczy się tylko footnotePr/endnotePr w sectPr sekcji odnośnika (settings.xml Word pomija),
// sekcja bez ustawień nie dziedziczy formatu, numStart + licznik w zakresie (dokument / sekcja /
// strona), przy „co stronę” numStart pominięty, słownie/porządkowe po angielsku. ENGINE=webkit.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");
const { CASES, toDocx } = require("../tools/word-oracle/cases-notes");
const EXPECTED = require("../tools/word-oracle/expected-notes.json");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const errors = [];
  for (const c of CASES) {
    const w = EXPECTED[c.name];
    if (!w) { check(`oczekiwane z Worda: ${c.name}`, false, "brak w expected-notes.json (WRITE=1 node tools/word-oracle/run-notes.js)"); continue; }
    const page = await context.newPage();
    page.on("pageerror", (e) => errors.push(e.message));
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.locator("#fileInput").setInputFiles({ name: `${c.name}.docx`, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await buildDocx(toDocx(c)) });
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host ol.dwb-notes p sup.note-mark"), null, { timeout: 30000 });
    await page.waitForTimeout(150);
    const a = await page.evaluate(() => {
      const out = { refs: {}, notes: {} };
      const clean = (s) => s.replace(/[﻿​]/g, "").replace(/[\s ]+/g, " ").trim();
      document.querySelectorAll(".docx-preview-host p").forEach((p) => {
        const t = clean(p.textContent);
        const ol = p.closest("ol.dwb-notes");
        let m;
        if (!ol && (m = /^([RE]\d\d):(.*)$/.exec(t))) out.refs[m[1]] = m[2].trim();
        if (ol && (m = /^(.+?)\s*([NM]\d\d)$/.exec(t))) out.notes[m[2]] = m[1];
      });
      return out;
    });
    const bad = [...Object.keys(w.refs).filter((k) => a.refs[k] !== w.refs[k]).map((k) => `${k}: ${a.refs[k]}≠${w.refs[k]}`),
      ...Object.keys(w.notes).filter((k) => a.notes[k] !== w.notes[k]).map((k) => `${k}: ${a.notes[k]}≠${w.notes[k]}`)];
    check(`jak Word: ${c.name} (${Object.values(w.refs).join(" ")})`, !bad.length, bad.join(", "));
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
