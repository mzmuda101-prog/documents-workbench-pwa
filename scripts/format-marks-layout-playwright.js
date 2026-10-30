// format-marks-layout-playwright.js — ¶ nie rusza układu (2026-10-30).
//
// Zgłoszenie Mateusza (nagranie): po włączeniu ¶ linijki się rozjeżdżały, strona robiła się niższa,
// a linia w nagłówku strony 2 przecinała tekst. Przyczyna: znak ¶ był zwykłym tekstem w kroju
// interfejsu (inna wysokość linijki) i zabierał miejsce w wierszu. W Wordzie ¶ tylko się dorysowuje —
// tu też: każdy akapit i każda strona stoją co do piksela tam, gdzie bez ¶. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const path = require("path");
const { APP_URL, DOCX_FIXTURE, fixtureExists } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const SAMPLES = ["przewodnik", "sample", "styled-sample", "forms-sample", "notes-sample", "headings-sample"]
  .map((n) => path.resolve(__dirname, `../docs/samples/${n}.docx`));
if (fixtureExists(DOCX_FIXTURE)) SAMPLES.push(DOCX_FIXTURE);

// pozycje akapitów (także nagłówków/stopek stron) i wysokości stron
const snap = (page) => page.evaluate(() => {
  const host = document.querySelector(".docx-preview-host");
  const r = (el) => { const b = el.getBoundingClientRect(); return [b.left, b.top, b.width, b.height].map((v) => Math.round(v * 10) / 10); };
  return {
    pages: [...host.querySelectorAll("section.docx")].map(r),
    paras: [...host.querySelectorAll("section.docx p")].map(r),
  };
});
const diff = (a, b) => {
  const out = [];
  for (const k of ["pages", "paras"]) {
    if (a[k].length !== b[k].length) { out.push(`${k}: ${a[k].length} → ${b[k].length}`); continue; }
    a[k].forEach((v, i) => { if (v.some((x, j) => Math.abs(x - b[k][i][j]) > 0.6)) out.push(`${k}[${i}] ${v.join(",")} → ${b[k][i].join(",")}`); });
  }
  return out;
};
// czeka, aż układ stanie (WebKit dociąga obrazy/kroje jeszcze po zdjęciu nakładki wczytywania)
const settle = async (page) => {
  await page.evaluate(() => document.fonts.ready);
  let prev = null;
  for (let i = 0; i < 20; i++) {
    await page.waitForTimeout(300);
    const cur = await snap(page);
    if (prev && diff(prev, cur).length === 0) return cur;
    prev = cur;
  }
  return prev;
};

(async () => {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); try { localStorage.removeItem("dwb-show-formatting-marks-v1"); } catch (_) {} });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); if (typeof setSidebarOpen === "function") setSidebarOpen(false); });

  // Czytanie, potem Edycja (tam granice stron liczy page-breaks.js — też nie mogą drgnąć)
  for (const mode of ["read", "edit"]) for (const file of SAMPLES) {
    const name = `${path.basename(file)} [${mode === "read" ? "Czytanie" : "Edycja"}]`;
    await page.locator("#fileInput").setInputFiles(file);
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host section.docx p"), null, { timeout: 60000 });
    await page.evaluate((read) => { if (readModeEl.checked !== read) { readModeEl.checked = read; readModeEl.dispatchEvent(new Event("change")); } }, mode === "read");
    if (await page.evaluate(() => docCanvasEl.classList.contains("show-formatting-marks"))) await page.click("#formatMarksBtn");
    await page.waitForTimeout(3000); // przewodnik w WebKit dokłada ostatni piksel jeszcze ~2–3 s po wczytaniu
    const off = await settle(page);
    await page.click("#formatMarksBtn");
    const on = await settle(page);
    const shown = await page.evaluate(() => { const p = [...document.querySelectorAll(".docx-preview-host section.docx p")].find((x) => x.textContent.trim()); return p ? getComputedStyle(p, "::after").content : ""; });
    await page.click("#formatMarksBtn");
    const back = await settle(page);
    const d = diff(off, on);
    check(`${name}: ¶ widać, a ${off.paras.length} akapitów i ${off.pages.length} stron stoi jak bez ¶`, /¶/.test(shown) && d.length === 0, `${shown} ${d.slice(0, 6).join(" | ")}${d.length > 6 ? ` … (+${d.length - 6})` : ""}`);
    check(`${name}: po wyłączeniu ¶ układ wraca do wyjściowego`, diff(off, back).length === 0, diff(off, back).slice(0, 4).join(" | "));
  }
  check("brak błędów strony", errors.length === 0, errors.join("\n"));

  await browser.close();
  const bad = results.filter((r) => !r.ok);
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `\n   ${r.detail}`}`);
  console.log(`\n[${ENGINE}] ${results.length - bad.length}/${results.length}`);
  process.exit(bad.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
