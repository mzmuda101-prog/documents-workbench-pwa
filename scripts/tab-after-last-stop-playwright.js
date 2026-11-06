// tab-after-last-stop-playwright.js — tabulator ZA ostatnim własnym tab-stopem akapitu idzie na
// najbliższy domyślny tabulator liczony od MARGINESU (co defaultTabStop), jak w Wordzie (lista 3.11).
// Pozycje oczekiwane z prawdziwego Worda (tools/word-oracle/probe-tab-after-last.js, 2026-11-06),
// w pt od lewego marginesu; tolerancja 1 pt.

const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const tab = (pos) => `<w:tabs><w:tab w:val="left" w:pos="${pos}"/></w:tabs>`;
const T = "<w:r><w:tab/></w:r>";
const r = (t) => `<w:r><w:t>${t}</w:t></w:r>`;
const CASES = [
  { id: "c1", pPr: tab(1985), raw: `${r("A")}${T}${r("BBBB")}${T}${r("Cx")}`, word: { BBBB: 99.23, Cx: 141.5 } },
  { id: "c2", pPr: `${tab(1985)}<w:ind w:left="567"/>`, raw: `${r("A")}${T}${r("BBBB")}${T}${r("Cx")}`, word: { BBBB: 99.23, Cx: 141.5 } },
  { id: "c3", pPr: '<w:tabs><w:tab w:val="left" w:pos="1000"/><w:tab w:val="left" w:pos="3000"/></w:tabs>', raw: `${r("A")}${T}${r("B")}${T}${r("D")}${T}${r("Cx")}`, word: { B: 49.95, D: 150, Cx: 177 } },
];

(async () => {
  const bytes = await buildDocx({ body: CASES.map((c) => `<w:p><w:pPr>${c.pPr}</w:pPr>${c.raw}</w:p>`) });
  const file = path.join(os.tmpdir(), `dwb-tab-after-last-${process.pid}.docx`);
  fs.writeFileSync(file, bytes);
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 1000 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles(file);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelectorAll("section.docx span.docx-tab-laid").length >= 7, null, { timeout: 30000 });
  await page.waitForTimeout(500);
  const got = await page.evaluate(() => {
    const out = [];
    const ps = [...document.querySelectorAll("section.docx p")].filter((p) => p.querySelector("span.docx-tab"));
    ps.forEach((p) => {
      const pr = p.getBoundingClientRect();
      const scale = p.offsetWidth ? pr.width / p.offsetWidth : 1;
      const origin = pr.left - (parseFloat(getComputedStyle(p).marginLeft) || 0) * scale;
      const pos = {};
      const tw = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      const range = document.createRange();
      for (let t = tw.nextNode(); t; t = tw.nextNode()) {
        const w = t.data.trim();
        if (!w || /^\s| /.test(t.data) && !w.replace(/ /g, "")) continue;
        const i = t.data.indexOf(w[0]);
        range.setStart(t, i); range.setEnd(t, i + 1);
        pos[w] = +(((range.getBoundingClientRect().left - origin) / scale) * 0.75).toFixed(2);
      }
      out.push(pos);
    });
    return out;
  });
  const fails = [];
  CASES.forEach((c, i) => {
    for (const [w, want] of Object.entries(c.word)) {
      const g = got[i] && got[i][w];
      if (!(typeof g === "number" && Math.abs(g - want) <= 1)) fails.push(`${c.id}: „${w}” na ${g} pt, Word: ${want} pt`);
    }
  });
  if (errors.length) fails.push("błędy strony: " + errors.join(" | "));
  await browser.close();
  fs.unlinkSync(file);
  if (fails.length) { console.log(`❌ tab-after-last-stop (${ENGINE}): ${fails.length} błędów`); fails.forEach((f) => console.log("  ✗", f)); process.exit(1); }
  console.log(`✅ tab-after-last-stop (${ENGINE}): wszystko przeszło`);
})().catch((e) => { console.error(e); process.exit(1); });
