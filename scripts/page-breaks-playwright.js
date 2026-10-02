// page-breaks-playwright.js — granice stron w Widoku desktopowym (app/page-breaks.js) i odstępy
// wierszy jak w Wordzie (applyWordLineMetrics w app/docx-render-fixes.js).
//
//   node scripts/page-breaks-playwright.js                (Chromium)
//   ENGINE=webkit node scripts/page-breaks-playwright.js  (WebKit)
//
// Wzorzec z prawdziwego Worda (Microsoft Word dla Maca, 2026-10-02, AppleScript „active end
// page number” dla każdego akapitu docs/samples/przewodnik.docx): 5 stron; strona 2 zaczyna się
// od „3. Placeholdery…”, strona 3 od „8. Przypisy, tabele i listy”. Akapit „1. Czytanie
// i edycja” (rozdział) Word stawia 452,6 pt od góry strony 1.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(700));

// tekst zaczynający każdą wyliczoną stronę
const pageStarts = (page) => page.evaluate(() => {
  const out = [];
  document.querySelectorAll(".docx-wrapper > section.docx").forEach((sec) => {
    const sr = sec.getBoundingClientRect();
    const scale = sr.height / sec.offsetHeight || 1;
    sec.querySelectorAll(".dwb-page-break").forEach((m) => {
      const y = sr.top + parseFloat(m.style.top) * scale;
      const walker = document.createTreeWalker(sec.querySelector("article"), NodeFilter.SHOW_TEXT);
      let found = null;
      while (!found && walker.nextNode()) {
        const n = walker.currentNode;
        if (!n.textContent.trim()) continue;
        const r = document.createRange();
        for (let i = 0; i < n.length; i++) {
          r.setStart(n, i); r.setEnd(n, i + 1);
          const q = r.getBoundingClientRect();
          if (q.height && q.top >= y - 1) { found = n.textContent.slice(i, i + 40); break; }
        }
      }
      out.push({ label: m.dataset.label, top: m.style.top, text: found });
    });
  });
  return out;
});

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=przewodnik`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await idle(page);

  // odstępy wierszy jak w Wordzie: nagłówek rozdziału 1 tam, gdzie w Wordzie (±6 pt)
  const y19 = await page.evaluate(() => {
    const sec = document.querySelector(".docx-preview-host section.docx");
    const p = [...sec.querySelectorAll("article p")].filter((x) => !x.closest("table"))[18];
    return { text: p.textContent.slice(0, 20), pt: (p.getBoundingClientRect().top - sec.getBoundingClientRect().top) * 0.75 };
  });
  check("odstępy jak w Wordzie: „1. Czytanie i edycja” ~452,6 pt od góry (±6)", /^1\. Czytanie/.test(y19.text) && Math.abs(y19.pt - 452.6) < 6, JSON.stringify(y19));

  const s1 = await pageStarts(page);
  check("domyślnie włączone: 4 granice (5 stron jak w Wordzie)", s1.length === 4 && s1.map((s) => s.label).join(",") === "str. 2,str. 3,str. 4,str. 5", JSON.stringify(s1));
  check("str. 2 zaczyna się jak w Wordzie: „3. Placeholdery”", /^3\. Placeholdery/.test(s1[0]?.text || ""), JSON.stringify(s1[0]));
  check("str. 3 zaczyna się jak w Wordzie: „8. Przypisy, tabele i listy”", /^8\. Przypisy/.test(s1[1]?.text || ""), JSON.stringify(s1[1]));
  check("etykiety „str. N” nie są tekstem dokumentu (słowa, szukanie, zapis)", await page.evaluate(() => !document.querySelector(".docx-preview-host").textContent.includes("str. 2")));
  check("granica nie przecina tekstu (zaczyna się od początku linijki)", s1.every((s) => s.text && s.text.length > 0));

  await page.evaluate(() => { zoomLevelEl.value = "0.5"; zoomLevelEl.dispatchEvent(new Event("input", { bubbles: true })); });
  await page.waitForTimeout(600);
  const s2 = await pageStarts(page);
  check("zoom 50%: te same granice (zoom nie zmienia stron)", JSON.stringify(s2) === JSON.stringify(s1), JSON.stringify(s2));
  const label = await page.evaluate(() => { const m = document.querySelector(".dwb-page-break"); return getComputedStyle(m, "::after").content; });
  check("etykieta pokazuje numer strony", /str\. 2/.test(label), label);
  await page.click("#zoomFitBtn");
  await page.waitForTimeout(400);

  // pisanie przesuwa granice: dopisany długi tekst na stronie 1 → str. 2 zaczyna się wcześniej w treści
  await page.click('.mode-btn[data-mode="edit"]');
  await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.isContentEditable && p.textContent.trim().length > 40);
    placeCaret(ps[0], ps[0].textContent.length);
  });
  await page.keyboard.type(" Dopisany tekst, który wydłuża pierwszą stronę o kilka linijek i przesuwa resztę dokumentu w dół. ".repeat(4), { delay: 1 });
  await page.waitForTimeout(1200);
  const s3 = await pageStarts(page);
  // (dopisane linijki zjadają wolne miejsce na dole str. 1, które zostawił nagłówek przeniesiony
  //  „razem z następnym” — dlatego sprawdzamy położenie granicy, nie tekst)
  check("po dopisaniu tekstu granice przeliczają się (str. 2 niżej w kartce)", s3[0] && parseFloat(s3[0].top) > parseFloat(s1[0].top), JSON.stringify({ przed: s1[0], po: s3[0] }));
  check("dalej nie w tekście dokumentu (także w trybie Edycja)", await page.evaluate(() => !collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).some((p) => p.querySelector(".dwb-page-break"))));

  // opcja w panelu Widok: wyłączenie chowa granice i jest pamiętane
  await page.evaluate(() => { setSidebarOpen(true); document.getElementById("panel-view").open = true; });
  await page.click("#showPageBreaks");
  await page.waitForTimeout(300);
  check("wyłączona opcja: brak granic", await page.evaluate(() => document.querySelectorAll(".dwb-page-break").length === 0));
  check("wyłączenie zapamiętane", await page.evaluate(() => localStorage.getItem("dwb-page-breaks-v1") === "0"));
  await page.click("#showPageBreaks");
  await page.waitForTimeout(500);
  check("włączona z powrotem: granice są", await page.evaluate(() => document.querySelectorAll(".dwb-page-break").length >= 4));
  await page.evaluate(() => setSidebarOpen(false));

  // Widok mobilny: bez granic (tekst przełożony na szerokość ekranu)
  await page.evaluate(() => setViewLayoutPref("mobile"));
  await idle(page);
  check("Widok mobilny: bez granic stron", await page.evaluate(() => document.querySelectorAll(".dwb-page-break").length === 0 && !document.documentElement.classList.contains("page-breaks-on")));
  await page.evaluate(() => setViewLayoutPref("auto"));
  await idle(page);
  check("z powrotem Widok desktopowy: granice wracają", await page.evaluate(() => document.querySelectorAll(".dwb-page-break").length >= 4));
  check("niezapisany tekst przetrwał zmiany widoku", await page.evaluate(() => document.querySelector(".docx-preview-host").textContent.includes("Dopisany tekst") && hasUnsavedChanges));

  // druk: znaczniki ukryte
  check("druk/PDF: granice ukryte", await page.evaluate(() => {
    const css = [...document.styleSheets].flatMap((s) => { try { return [...s.cssRules]; } catch (_) { return []; } })
      .some((r) => r.media && /print/.test(r.media.mediaText) && /dwb-page-break/.test(r.cssText));
    return css;
  }));

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ granice stron [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
