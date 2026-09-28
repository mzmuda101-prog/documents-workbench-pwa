// inline-enter-playwright.js — szybkie pisanie wokół Enter/Backspace w podglądzie.
//
//   node scripts/inline-enter-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/inline-enter-playwright.js
//
// Błąd z 2026-09-28: Enter dzielił akapit w podglądzie od razu, ale kursor przeskakiwał
// dopiero po przebudowie pliku — tekst pisany w tym czasie lądował w złym akapicie,
// a ZAPISANY plik różnił się od podglądu (34 akapity na ekranie, 33 w pliku).
// Musi być prawdą: to, co widać, to dokładnie to, co trafia do pliku — przy każdym tempie,
// także na dużym dokumencie (przebudowa trwa wtedy najdłużej).

const fs = require("fs");
const pw = require("playwright");
const { APP_URL, DOCX_FIXTURE } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function session(browser, loadDoc, anchorText, delay, label) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await loadDoc(page);
  await page.click('.mode-btn[data-mode="edit"]');
  await page.locator(".docx-preview-host p", { hasText: anchorText }).first().click();
  // kursor na KOŃCU akapitu (End w długim akapicie idzie tylko na koniec wiersza)
  await page.evaluate((anchor) => {
    const el = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.includes(anchor));
    const r = document.createRange();
    r.selectNodeContents(el);
    r.collapse(false);
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  }, anchorText);
  await page.keyboard.type(" A", { delay });
  await page.keyboard.press("Enter");
  await page.keyboard.type("Pierwszy nowy", { delay });
  await page.keyboard.press("Enter");
  await page.keyboard.type("Drugi nowy", { delay });
  await page.keyboard.press("Enter");
  await page.keyboard.press("Backspace");
  await page.keyboard.type(" koniec", { delay });
  const r = await page.evaluate(async (anchor) => {
    const bytes = await buildDocumentForSave(); // czeka na kolejkę przebudowy
    const dom = [...document.querySelectorAll(".docx-preview-host p")].map((x) => x.textContent);
    const xml = (await extractParagraphTextsFromDocx(bytes)).map(String);
    const i = dom.findIndex((t) => t.includes(anchor));
    return { same: JSON.stringify(dom) === JSON.stringify(xml), nDom: dom.length, nXml: xml.length, around: dom.slice(i, i + 3).map((t) => t.slice(-24)) };
  }, anchorText);
  const expected = r.around[1] === "Pierwszy nowy" && r.around[2].endsWith("Drugi nowy koniec") && r.around[0].endsWith(" A");
  check(`${label}, tempo ${delay} ms: tekst w dobrych akapitach`, expected, JSON.stringify(r.around));
  check(`${label}, tempo ${delay} ms: zapis = podgląd`, r.same, `DOM ${r.nDom} / plik ${r.nXml}`);
  if (errors.length) check(`${label}: brak błędów strony`, false, errors.slice(0, 2).join(" | "));
  await context.close();
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const small = async (page) => {
    await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
    await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
    await page.waitForTimeout(500);
  };
  const big = async (page) => {
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.locator("#fileInput").setInputFiles(DOCX_FIXTURE);
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 60000 });
    await page.waitForTimeout(500);
  };
  try {
    for (const delay of [0, 60]) await session(browser, small, "Dane Wynajmującego", delay, "mały plik");
    if (fs.existsSync(DOCX_FIXTURE)) await session(browser, big, "Phasellus dapibus", 0, "duży plik (3,8 MB)");
  } finally {
    await browser.close();
  }
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ Enter/Backspace przy szybkim pisaniu [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
