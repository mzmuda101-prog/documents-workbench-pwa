// pdf-centered-paragraph-playwright.js — PDF → DOCX: wielowierszowy WYŚRODKOWANY akapit zostaje
// jednym akapitem, także gdy ma wcięcia po bokach (zawija się na własnej szerokości, nie na
// szerokości obszaru — lista 3.11), a tytuł z krótkich wyśrodkowanych wierszy zostaje podzielony.
// PDF-y z prawdziwego Worda: tools/word-oracle/probe-center-pdf.js → scripts/fixtures/center-*.pdf.

const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const FIX = path.join(__dirname, "fixtures");

async function convert(page, file) {
  await page.evaluate(() => { window.__conv = null; if (!window.__convHooked) { window.addEventListener("dwb:pdf-converted", (e) => (window.__conv = e.detail)); window.__convHooked = true; } });
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction(() => window.__conv, null, { timeout: 60000 });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.waitForTimeout(300);
  return page.evaluate(() => [...document.querySelectorAll("section.docx p")].map((p) => ({ t: p.textContent.replace(/\s+/g, " ").trim(), al: getComputedStyle(p).textAlign })).filter((x) => x.t));
}

(async () => {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const fails = [];
  const one = (ps, start) => ps.filter((p) => /Niniejsze zarządzenie|okres kadencji|tryb składania/.test(p.t));
  for (const name of ["full", "indented"]) {
    const ps = await convert(page, path.join(FIX, `center-${name}.pdf`));
    const hit = one(ps);
    if (hit.length !== 1 || !/^Niniejsze.*kadencji\.$/.test(hit[0].t)) fails.push(`${name}: akapit podzielony na ${hit.length}: ${JSON.stringify(hit.map((h) => h.t))}`);
    else if (!/center/.test(hit[0].al)) fails.push(`${name}: akapit nie jest wyśrodkowany (${hit[0].al})`);
  }
  const t = await convert(page, path.join(FIX, "center-title.pdf"));
  const titles = t.filter((p) => /ZARZĄDZENIE|Rektora|z dnia/.test(p.t));
  if (titles.length !== 3) fails.push(`tytuł: ${titles.length} akapitów zamiast 3: ${JSON.stringify(titles.map((x) => x.t))}`);
  if (errors.length) fails.push("błędy strony: " + errors.join(" | "));
  await browser.close();
  if (fails.length) { console.log(`❌ pdf-centered-paragraph (${ENGINE}): ${fails.length} błędów`); fails.forEach((f) => console.log("  ✗", f)); process.exit(1); }
  console.log(`✅ pdf-centered-paragraph (${ENGINE}): wszystko przeszło`);
})().catch((e) => { console.error(e); process.exit(1); });
