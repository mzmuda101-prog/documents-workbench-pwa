// app-frame-playwright.js — strażnik ramy ekranu (paczka B, app/app-frame.js).
//
//   node scripts/app-frame-playwright.js              (Chromium)
//   ENGINE=webkit node scripts/app-frame-playwright.js (Safari/iPhone)
//
// Co musi być prawdą:
//   - pusty start ma przyciski, nagłówek = nazwa apki; po wczytaniu = nazwa pliku + słowa,
//   - „Zapisz” w nagłówku: szary bez zmian, akcent + liczba zmian po edycji,
//   - menu ⋯: otwiera się, ma język/motyw/odśwież/zamknij, Esc zamyka,
//   - ≥1024 px panel stoi OBOK dokumentu (bez scrimu), « chowa, wybór zapamiętany,
//   - < 1024 px panel wysuwany, nagłówek nie jest zasłonięty, gdy panel zamknięty,
//   - „Znajdź ustawienie…” zostawia tylko pasujące sekcje,
//   - skróty sekcji: nagłówki z styles.xml (też polski Word), klik = skok, × chowa na stałe,
//   - Czytanie/Edycja przełącza edycję w podglądzie, B/I/U tylko w Edycji,
//   - szukanie na pasku: licznik „1 / N” bez podwójnego liczenia,
//   - telefon: nagłówek zwija się przy przewijaniu dokumentu i wraca na górze.

const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const HEADINGS_DOCX = path.resolve(__dirname, "../docs/samples/headings-sample.docx");

const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function newPage(browser, viewport) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  return { context, page, errors };
}

async function loadHeadingsDoc(page, name = "Umowa-najmu.docx") {
  await page.locator("#fileInput").setInputFiles({ name, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: require("fs").readFileSync(HEADINGS_DOCX) });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 20000 });
  await page.waitForTimeout(300);
}

// Czekanie na WARUNEK, nie na zegar: runner w tłoku (JOBS=4) skraca waitForTimeout ×0,5, a płynne
// przewijanie (skok do sekcji), przejście panelu (--t-slow) i rAF-y apki (zwijanie nagłówka,
// syncDocViewportHeight) trwają tyle, ile trwają. Timeout = margines; check niżej i tak mierzy.
// Predykaty TYLKO jako funkcje — string idzie przez eval, a CSP apki (bez 'unsafe-eval') go blokuje.
const frames = (page, n = 3) => page.evaluate((n) => new Promise((resolve) => {
  const step = () => (n-- > 0 ? requestAnimationFrame(step) : resolve());
  step();
}), n);
const isCollapsed = () => document.body.classList.contains("hero-collapsed");
const isExpanded = () => !document.body.classList.contains("hero-collapsed");
const pageOver = () => document.documentElement.scrollHeight - window.innerHeight;
const pageFits = () => document.documentElement.scrollHeight - window.innerHeight <= 0;
const until = (page, fn, arg) => page.waitForFunction(fn, arg, { timeout: 3000 })
  .catch((e) => { if (!(e instanceof pw.errors.TimeoutError)) throw e; }); // timeout → rozstrzyga check

async function desktop(browser) {
  const { context, page, errors } = await newPage(browser, { width: 1280, height: 860 });

  // pusty start
  const empty = await page.evaluate(() => ({
    title: document.getElementById("heroTitle").textContent,
    buttons: !!document.getElementById("emptyOpenBtn")?.offsetParent && !!document.getElementById("emptySampleBtn")?.offsetParent,
    saveHidden: document.getElementById("heroSaveBtn").hidden,
    toolbarHidden: getComputedStyle(document.getElementById("docToolbar")).display === "none",
  }));
  check("pusty start: nazwa apki, przyciski, brak „Zapisz” i paska", empty.title === "Documents Workbench" && empty.buttons && empty.saveHidden && empty.toolbarHidden, JSON.stringify(empty));

  // panel obok dokumentu
  const dock = await page.evaluate(() => {
    const side = document.querySelector(".sidebar").getBoundingClientRect();
    const main = document.querySelector(".main").getBoundingClientRect();
    return {
      docked: document.documentElement.classList.contains("sidebar-docked"),
      open: document.documentElement.classList.contains("sidebar-open"),
      scrim: !document.getElementById("sidebarScrim").classList.contains("hidden") && getComputedStyle(document.getElementById("sidebarScrim")).display !== "none",
      sideRight: Math.round(side.right), mainLeft: Math.round(main.left),
    };
  });
  check("≥1024: panel otwarty OBOK dokumentu, bez scrimu", dock.docked && dock.open && !dock.scrim && dock.sideRight <= dock.mainLeft, JSON.stringify(dock));

  await loadHeadingsDoc(page);
  const hero = await page.evaluate(() => ({
    title: document.getElementById("heroTitle").textContent,
    meta: document.getElementById("heroMeta").textContent,
    save: !document.getElementById("heroSaveBtn").hidden && !document.getElementById("heroSaveBtn").classList.contains("is-dirty"),
  }));
  check("po wczytaniu: nazwa pliku + słowa w nagłówku, „Zapisz” szary", hero.title === "Umowa-najmu.docx" && /\d/.test(hero.meta) && hero.save, JSON.stringify(hero));

  // „Aktualizuj”/„Zapisz” klikalne przy otwartym panelu — nic ich nie zasłania
  const heroClickable = await page.evaluate(() => {
    const b = document.getElementById("heroSaveBtn").getBoundingClientRect();
    const top = document.elementFromPoint(b.left + b.width / 2, b.top + b.height / 2);
    return !!top && !!top.closest("#heroSaveBtn");
  });
  check("nagłówek klikalny przy otwartym panelu", heroClickable);

  // skróty sekcji
  const chips = await page.evaluate(() => [...document.querySelectorAll(".section-chip")].map((c) => c.textContent));
  check("skróty sekcji z nagłówków pliku (EN/PL/outlineLvl/basedOn)", chips.length === 8 && chips.includes("§3 Czynsz i opłaty") && chips.includes("§1 Strony umowy"), chips.join(" | "));
  await page.locator(".section-chip", { hasText: "§6 Wypowiedzenie" }).click();
  const jumpState = () => {
    const vp = document.getElementById("docViewport");
    const h = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.includes("§6 Wypowiedzenie"));
    const top = h.getBoundingClientRect().top - vp.getBoundingClientRect().top;
    return { top: Math.round(top), cur: document.querySelector(".section-chip.is-current")?.textContent };
  };
  // scrollTo({ behavior: "smooth" }) — w tłoku dłużej niż sztywne 450 ms
  await until(page, () => {
    const vp = document.getElementById("docViewport");
    const h = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.includes("§6 Wypowiedzenie"));
    const top = h.getBoundingClientRect().top - vp.getBoundingClientRect().top;
    return top >= -2 && top < 120 && document.querySelector(".section-chip.is-current")?.textContent === "§6 Wypowiedzenie";
  });
  const jump = await page.evaluate(jumpState);
  check("klik w skrót = skok do sekcji + podświetlenie", jump.top >= -2 && jump.top < 120 && jump.cur === "§6 Wypowiedzenie", JSON.stringify(jump));

  // szukanie na pasku (moduł ładuje się leniwie przy pierwszym Enter)
  await page.fill("#searchQuery", "Najemca");
  await page.press("#searchQuery", "Enter");
  await page.waitForFunction(() => /\/\s*\d+/.test(document.getElementById("searchPos").textContent), null, { timeout: 8000 }).catch(() => {});
  const pos = await page.evaluate(() => ({ pos: document.getElementById("searchPos").textContent, badge: document.querySelector("#panel-search .panel-count")?.textContent }));
  check("szukanie na pasku: „1 / 21” bez podwójnego liczenia + licznik przy sekcji", pos.pos === "1 / 21" && pos.badge === "21", JSON.stringify(pos));
  await page.click("#frNextBtn");
  check("↓ przechodzi do następnego trafienia", (await page.textContent("#searchPos")) === "2 / 21");

  // Czytanie / Edycja
  const ro = await page.evaluate(() => ({ fmtHidden: document.getElementById("formatToolbar").classList.contains("hidden"), editable: document.querySelector(".docx-preview-host p")?.isContentEditable }));
  await page.click('.mode-btn[data-mode="edit"]');
  const ed = await page.evaluate(() => ({ fmtHidden: document.getElementById("formatToolbar").classList.contains("hidden"), editable: document.querySelector(".docx-preview-host p")?.isContentEditable, pressed: document.querySelector('.mode-btn[data-mode="edit"]').getAttribute("aria-pressed") }));
  check("Czytanie → Edycja: edycja w podglądzie, B/I/U widoczne", ro.fmtHidden && !ro.editable && !ed.fmtHidden && ed.editable && ed.pressed === "true", JSON.stringify({ ro, ed }));

  // edycja → „Zapisz” z liczbą zmian
  const para = page.locator(".docx-preview-host p", { hasText: "Dokument testowy" }).first();
  const wordsBeforeEdit = await page.textContent("#statusWords");
  await para.click();
  await page.keyboard.press("End");
  await page.keyboard.type(" (poprawka)");
  await until(page, () => document.getElementById("heroSaveCount").textContent === "1");
  const dirty = await page.evaluate(() => ({ dirty: document.getElementById("heroSaveBtn").classList.contains("is-dirty"), count: document.getElementById("heroSaveCount").textContent, label: document.getElementById("heroSaveBtn").getAttribute("aria-label"), status: document.getElementById("statusDirty").hidden ? "" : document.getElementById("statusDirty").textContent }));
  check("po edycji „Zapisz” niebieski z liczbą zmian (1) + pigułka w pasku stanu", dirty.dirty && dirty.count === "1" && /1 zmiana/.test(dirty.label) && /niezapisane: 1 zmiana/.test(dirty.status), JSON.stringify(dirty));
  await until(page, (before) => document.getElementById("statusWords").textContent !== before, wordsBeforeEdit);
  const liveWords = await page.evaluate(() => ({
    status: document.getElementById("statusWords").textContent,
    hero: document.getElementById("heroMeta").textContent,
    expected: analyzeDocumentDom(document.getElementById("docCanvas")).words,
  }));
  check("po wpisaniu licznik słów w nagłówku i pasku stanu jest aktualny", liveWords.status.includes(String(liveWords.expected)) && liveWords.hero.includes(String(liveWords.expected)), JSON.stringify(liveWords));

  // Ctrl/⌘+S = Zapisz (przechwycone przez apkę, nie przeglądarkę)
  const saveHit = await page.evaluate(() => new Promise((resolve) => {
    const btn = document.getElementById("saveBtn");
    const orig = btn.click.bind(btn);
    btn.click = () => { btn.click = orig; resolve(true); };
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "s", code: "KeyS", ctrlKey: true, bubbles: true, cancelable: true }));
    setTimeout(() => resolve(false), 500);
  }));
  check("Ctrl+S uruchamia zapis", saveHit);

  // menu ⋯
  await page.click("#appMenuBtn");
  const menu = await page.evaluate(() => {
    const m = document.getElementById("appMenu");
    return { open: !m.hidden, lang: m.contains(document.getElementById("langSwitch")), theme: m.contains(document.getElementById("themeToggle")), close: !document.getElementById("closeDocBtn").classList.contains("hidden"), inBody: m.parentElement === document.body };
  });
  await page.keyboard.press("Escape");
  const menuClosed = await page.evaluate(() => document.getElementById("appMenu").hidden);
  check("menu ⋯: język, motyw, zamknij dokument; Esc zamyka", menu.open && menu.lang && menu.theme && menu.close && menu.inBody && menuClosed, JSON.stringify(menu));

  // Znajdź ustawienie…
  await page.fill("#sidebarFinder", "kluczowe"); // „autor” jest też w Recenzji („Autor zmian”)
  const finder = await page.evaluate(() => [...document.querySelectorAll(".sidebar details.panel")].filter((d) => !d.hidden).map((d) => d.id));
  await page.fill("#sidebarFinder", "");
  const finderReset = await page.evaluate(() => [...document.querySelectorAll(".sidebar details.panel")].every((d) => !d.hidden));
  check("„Znajdź ustawienie…”: „kluczowe” → tylko Metadane, po wyczyszczeniu wszystko", finder.length === 1 && finder[0] === "panel-metadata" && finderReset, finder.join(","));

  // × chowa skróty sekcji — na stałe (opcja w Widok)
  await page.click("#sectionStripHide");
  const hidden = await page.evaluate(() => ({ strip: document.getElementById("sectionStrip").hidden, opt: document.getElementById("showSectionChips").checked, saved: localStorage.getItem("dwb-section-strip-v1") }));
  check("× chowa skróty sekcji i zapamiętuje (opcja w Widok odznaczona)", hidden.strip && !hidden.opt && hidden.saved === "0", JSON.stringify(hidden));
  await page.evaluate(() => { const o = document.getElementById("showSectionChips"); o.checked = true; o.dispatchEvent(new Event("change")); });
  check("opcja w Widok przywraca skróty", await page.evaluate(() => !document.getElementById("sectionStrip").hidden));

  // « chowa panel; zapamiętane
  await page.click("#sidebarCloseBtn");
  // przejście marginesu (--t-slow 250 ms; sztywne 450 ms ×0,5 kończyło się przed nim)
  await until(page, () => document.querySelector(".main").getBoundingClientRect().left < 60);
  const closed = await page.evaluate(() => ({ open: document.documentElement.classList.contains("sidebar-open"), saved: localStorage.getItem("dwb-panel-docked-open-v1"), mainLeft: Math.round(document.querySelector(".main").getBoundingClientRect().left) }));
  check("« chowa panel (dokument zajmuje miejsce), wybór zapamiętany", !closed.open && closed.saved === "0" && closed.mainLeft < 60, JSON.stringify(closed));

  if (errors.length) check("desktop: brak błędów w konsoli", false, errors.slice(0, 3).join(" | "));
  await context.close();
}

async function tablet(browser) {
  const { context, page, errors } = await newPage(browser, { width: 900, height: 1100 });
  const s = await page.evaluate(() => ({ docked: document.documentElement.classList.contains("sidebar-docked"), open: document.documentElement.classList.contains("sidebar-open") }));
  await loadHeadingsDoc(page);
  await page.click("#panelToggle");
  const opened = await page.evaluate(() => document.documentElement.classList.contains("sidebar-open"));
  check("< 1024: panel wysuwany, startuje schowany, ⚙ otwiera", !s.docked && !s.open && opened, JSON.stringify(s));
  if (errors.length) check("tablet: brak błędów w konsoli", false, errors.slice(0, 3).join(" | "));
  await context.close();
}

async function phone(browser) {
  const { context, page, errors } = await newPage(browser, { width: 390, height: 844 });
  await page.goto(APP_URL, { waitUntil: "load" }).catch(() => {});
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await until(page, pageFits);
  const emptyOver = await page.evaluate(pageOver);
  check("telefon, pusty start: strona mieści się w oknie (bez paska przewijania)", emptyOver <= 0, `wystaje o ${emptyOver}px`);
  await loadHeadingsDoc(page);
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth + 1);
  check("telefon: nic nie wychodzi poza ekran w bok", !overflow);
  await page.evaluate(() => { document.getElementById("docViewport").scrollTop = 500; });
  await until(page, isCollapsed);
  const collapsed = await page.evaluate(isCollapsed);
  await page.evaluate(() => { document.getElementById("docViewport").scrollTop = 0; });
  await frames(page); // scroll → rAF w apce musiał się odbyć, inaczej „NIE rozwija” sprawdzałoby nic
  const stillCollapsed = await page.evaluate(isCollapsed);
  check("telefon: nagłówek zwija się przy przewijaniu, samo dojechanie do góry go NIE rozwija", collapsed && stillCollapsed, JSON.stringify({ collapsed, stillCollapsed }));

  // pociągnięcie palcem na samej górze: 40 px za mało, 100 px rozwija (jak Sheet)
  const pull = (dy) => page.evaluate((dy) => {
    const vp = document.getElementById("docViewport");
    const ev = (type, y) => { const e = new Event(type, { bubbles: true }); Object.defineProperty(e, "touches", { value: y == null ? [] : [{ clientX: 200, clientY: y }] }); vp.dispatchEvent(e); };
    ev("touchstart", 300);
    for (let i = 1; i <= 5; i++) ev("touchmove", 300 + (dy * i) / 5);
    const pulling = document.getElementById("heroGrip").classList.contains("pulling") || !document.body.classList.contains("hero-collapsed");
    ev("touchend", null);
    return { pulling, collapsed: document.body.classList.contains("hero-collapsed") };
  }, dy);
  const small = await pull(40);
  check("telefon: krótkie pociągnięcie (40 px) — uchwyt reaguje, nagłówek dalej schowany", small.pulling && small.collapsed, JSON.stringify(small));
  const big = await pull(100);
  check("telefon: pociągnięcie „na siłę” (100 px) na górze rozwija nagłówek", !big.collapsed, JSON.stringify(big));
  // nagłówek wrócił → wysokość dokumentu przelicza się w następnej klatce
  await until(page, pageFits);
  const fits = await page.evaluate(pageOver);
  check("telefon: strona mieści się w oknie (brak przewijania całej strony)", fits <= 0, `wystaje o ${fits}px`);
  const css = await page.evaluate(() => ({ ob: getComputedStyle(document.getElementById("docViewport")).overscrollBehaviorY }));
  check("telefon: pociągnięcie dokumentu nie przechodzi na stronę (overscroll-behavior: contain — bez odświeżenia strony w Safari)", css.ob === "contain", JSON.stringify(css));
  await page.evaluate(() => { document.getElementById("docViewport").scrollTop = 500; });
  // tap dopiero PO zwinięciu — tap w rozwinięty nagłówek go zwija (przełącznik)
  await until(page, isCollapsed);
  const beforeTap = await page.evaluate(isCollapsed);
  await page.click("#heroGrip");
  await until(page, isExpanded);
  const afterTap = await page.evaluate(isCollapsed);
  check("telefon: tap w uchwyt rozwija nagłówek", beforeTap && !afterTap, JSON.stringify({ beforeTap, afterTap }));
  if (errors.length) check("telefon: brak błędów w konsoli", false, errors.slice(0, 3).join(" | "));
  await context.close();
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  try {
    await desktop(browser);
    await tablet(browser);
    await phone(browser);
  } finally {
    await browser.close();
  }
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail && !r.ok ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ rama ekranu [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
