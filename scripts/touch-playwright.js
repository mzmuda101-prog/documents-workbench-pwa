// touch-playwright.js — strażnik lekcji z iPhone'a (paczka C). Kontekst dotykowy (pointer: coarse).
//
//   node scripts/touch-playwright.js                (Chromium, telefon z dotykiem)
//   ENGINE=webkit node scripts/touch-playwright.js  (WebKit, telefon z dotykiem)
//
// Każdy punkt to błąd odtworzony na symulatorze iPhone'a albo lekcja z Sheet Workbench:
//   - pola < 16 px → iOS przybliża CAŁĄ stronę przy fokusie (i zostaje przybliżona),
//   - backdrop-filter na dotyku → drogie rozmycia (w Sheet: lag przewijania),
//   - content-visibility w wysuwanym panelu → białe sekcje na iOS,
//   - position:sticky w przewijanym panelu → tap trafia ~35 pt obok po przewinięciu,
//   - toasty u dołu telefonu → zasłonięte (pasek stanu, Safari, „Schowaj panel”),
//   - „przeciągnij plik w dowolne miejsce okna” → bez sensu na dotyku.
// Czego Playwright NIE sprawdzi (tylko symulator/telefon): zachowanie przy prawdziwej
// klawiaturze ekranowej (kb-open w touch.js) — opis w pamięci projektu.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({
    serviceWorkers: "block",
    viewport: { width: 390, height: 844 },
    hasTouch: true,
    ...(ENGINE === "chromium" ? { isMobile: true } : {}),
  });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const coarse = await page.evaluate(() => matchMedia("(pointer: coarse)").matches);
  check("kontekst testu to dotyk (pointer: coarse)", coarse);

  check("pusty start: bez podpowiedzi o przeciąganiu pliku", await page.evaluate(() => getComputedStyle(document.getElementById("emptyDropHint")).display === "none"));

  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.waitForTimeout(400);

  const small = await page.evaluate(() => [...document.querySelectorAll("input:not([type=checkbox]):not([type=range]):not([type=file]):not([hidden]), select, textarea")]
    .filter((el) => parseFloat(getComputedStyle(el).fontSize) < 16)
    .map((el) => `${el.id || el.tagName} ${getComputedStyle(el).fontSize}`));
  check("pola ≥ 16 px (iOS nie przybliża strony przy fokusie)", small.length === 0, small.slice(0, 5).join(", "));

  const blurred = await page.evaluate(() => [...document.querySelectorAll("body *")]
    .filter((el) => { const cs = getComputedStyle(el); const v = cs.backdropFilter || cs.webkitBackdropFilter; return v && v !== "none"; })
    .map((el) => el.className || el.tagName).slice(0, 5));
  check("brak backdrop-filter na dotyku", blurred.length === 0, blurred.join(", "));

  const sidebar = await page.evaluate(() => {
    const bad = [];
    document.querySelectorAll(".sidebar, .sidebar *").forEach((el) => {
      const cs = getComputedStyle(el);
      if (cs.contentVisibility && cs.contentVisibility !== "visible") bad.push(`cv:${el.id || el.className}`);
      if (cs.position === "sticky") bad.push(`sticky:${el.id || el.className}`);
    });
    return bad;
  });
  check("panel: bez content-visibility i sticky (iOS: białe sekcje / tap obok)", sidebar.length === 0, sidebar.slice(0, 5).join(", "));

  // bez rozmycia półprzezroczyste tła to „dziury” — pod panelem prześwitywał dokument
  await page.evaluate(() => setSidebarOpen(true));
  await page.waitForTimeout(350);
  const see = await page.evaluate(() => {
    const alpha = (el) => { const m = getComputedStyle(el).backgroundColor.match(/[\d.]+(?=\s*\)$)/); const c = getComputedStyle(el).backgroundColor; return /rgba|\/ /.test(c) && m ? parseFloat(m[0]) : 1; };
    return { sidebar: alpha(document.querySelector(".sidebar")), panel: alpha(document.querySelector("#panel-structure")) };
  });
  const lastPanel = await page.evaluate(() => {
    const scroll = document.querySelector(".sidebar-scroll");
    scroll.scrollTop = scroll.scrollHeight;
    const last = [...document.querySelectorAll(".sidebar details.panel")].pop().getBoundingClientRect();
    const btn = document.getElementById("panelHandle").getBoundingClientRect();
    return { lastBottom: Math.round(last.bottom), btnTop: Math.round(btn.top) };
  });
  await page.evaluate(() => setSidebarOpen(false));
  check("panel na telefonie: pełne tło (panel i karty)", see.sidebar === 1 && see.panel === 1, JSON.stringify(see));
  check("„Schowaj panel” nie zasłania ostatniej sekcji", lastPanel.lastBottom <= lastPanel.btnTop, JSON.stringify(lastPanel));

  const targets = await page.evaluate(() => [...document.querySelectorAll(".doc-toolbar .tb-btn, .mode-btn, #appMenuBtn, #heroSaveBtn")]
    .filter((el) => el.offsetParent)
    .map((el) => ({ id: el.id || el.className, h: Math.round(el.getBoundingClientRect().height) }))
    .filter((x) => x.h < 34));
  check("cele dotykowe paska ≥ 34 px wysokości", targets.length === 0, JSON.stringify(targets.slice(0, 4)));

  await page.evaluate(() => toast("test", "info"));
  const toastPos = await page.evaluate(() => {
    const r = document.querySelector(".toast").getBoundingClientRect();
    return { top: Math.round(r.top), bottom: Math.round(window.innerHeight - r.bottom) };
  });
  check("toast na telefonie u góry ekranu", toastPos.top < 80, JSON.stringify(toastPos));

  // Enter w szukaniu chowa klawiaturę (blur) — na dotyku wynik ma być widać
  await page.fill("#searchQuery", "Najemca");
  await page.focus("#searchQuery");
  await page.press("#searchQuery", "Enter");
  await page.waitForFunction(() => /\d+ \/ \d+/.test(document.getElementById("searchPos").textContent), null, { timeout: 8000 }).catch(() => {});
  await page.waitForTimeout(150);
  const afterEnter = await page.evaluate(() => ({ pos: document.getElementById("searchPos").textContent, focused: document.activeElement?.id }));
  check("Enter w szukaniu: wynik + klawiatura schowana (pole bez fokusu)", /1 \/ 21/.test(afterEnter.pos) && afterEnter.focused !== "searchQuery", JSON.stringify(afterEnter));

  const overflow = await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth + 1);
  check("nic nie wychodzi poza ekran w bok", !overflow);

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ dotyk [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
