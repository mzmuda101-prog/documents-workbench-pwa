// run-case.js — „Zmień wielkość liter” (Aa): Word kontra aplikacja (changeTextCase)
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { buildDocx } = require("./build-docx");
const word = require("./word");
const app = require("./app");

const TEXTS = [
  "ala MA kota. kot ma alę! czy? tak... np. to jest test",
  "biało-czerwony o'neil e-mail iPhone mcdonald's 3d ąę ŁÓDŹ żółw",
  "ALA MA KOTA. KOT MA ALĘ",
  "„cytat. drugi” (nawias. tu) - myślnik. 1. punkt",
  "Dr. Smith i prof. Kowalski w USA. mr. x",
  "tekst\tpo tabulatorze. i dalej",
  "to JEST A test 3D ABC123 DON'T zażółć X. y I",
  "aBC Abc ABc ABC-def ABC-DEF ŻÓŁW. ąą",
  "PIERWSZE zdanie. drugie ZDANIE",
];
// akapit z dwoma fragmentami (red + pogrubione „bold text”) — granica formatowania to nie granica słowa
const RUNS = { raw: '<w:r><w:t>red</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold text</w:t></w:r>' };
const MODES = ["sentence", "lower", "upper", "title", "toggle"];
const PARTIAL = [[4, 12], [8, 20], [0, 5], [10, 30]];

(async () => {
  const ops = [];
  TEXTS.forEach((t, ti) => MODES.forEach((m) => { ops.push({ text: t, mode: m, from: 0, to: null }); PARTIAL.forEach(([a, b]) => { if (b <= t.length) ops.push({ text: t, mode: m, from: a, to: b }); }); }));
  ["title", "sentence", "upper"].forEach((m) => ops.push({ text: "redbold text", raw: RUNS, mode: m, from: 0, to: null }));
  const bytes = await buildDocx({ body: ops.map((o) => o.raw || o.text) });
  const key = crypto.createHash("sha1").update(bytes).update(JSON.stringify(ops)).digest("hex").slice(0, 16);
  const cacheF = path.join(word.DIR, "cache", `case-${key}.json`);
  let w;
  if (fs.existsSync(cacheF) && !process.env.FRESH) w = JSON.parse(fs.readFileSync(cacheF, "utf8"));
  else { w = word.changeCase(bytes, ops.map((o, i) => ({ ...o, para: i + 1 }))); fs.mkdirSync(path.dirname(cacheF), { recursive: true }); fs.writeFileSync(cacheF, JSON.stringify(w)); }
  const { browser, page } = await app.openApp();
  await app.loadBytes(page, bytes, "case.docx");
  await page.evaluate(() => { if (readOnlyMode) appFrame.setReadOnly(false); });
  await page.waitForTimeout(300);
  const a = await page.evaluate((list) => list.map((o, i) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    const len = p.textContent.length;
    const r = formDomRange(p, o.from, o.to == null ? len : o.to);
    p.closest(".docx-edit-root")?.focus?.();
    getSelection().removeAllRanges(); getSelection().addRange(r);
    changeTextCase(o.mode);
    return p.textContent.replace(/﻿/g, "");
  }), ops);
  await browser.close();
  let bad = 0;
  // WRITE=1 — oczekiwane wyniki (Word) dla scripts/change-case-playwright.js
  if (process.env.WRITE) fs.writeFileSync(path.join(__dirname, "expected-case.json"), `${JSON.stringify(ops.map((o, i) => ({ text: o.text, runs: !!o.raw, mode: o.mode, from: o.from, to: o.to, word: w[i] })), null, 1)}\n`);
  ops.forEach((o, i) => {
    const ww = w[i].replace(/\u0009/g, "\t");
    if (ww !== a[i]) { bad++; console.log(`❌ ${o.mode} [${o.from}…${o.to ?? "koniec"}] „${o.text}”\n   Word: „${ww}”\n   apka: „${a[i]}”`); }
  });
  console.log(`\n${ops.length - bad}/${ops.length} jak w Wordzie`);
})();
