// keyboard-playwright.js — klawiatura, podpowiedzi, tryb skupienia, miejsce na dokument (paczka E).
//
//   node scripts/keyboard-playwright.js
//   ENGINE=webkit node scripts/keyboard-playwright.js
//
// Musi być prawdą (bez myszy):
//   - pierwszy Tab = „Przejdź do dokumentu”, Enter → fokus w dokumencie,
//   - Ctrl/⌘+F → pole szukania; Enter = szukaj, kolejny Enter = następne, Shift+Enter = poprzednie, F3,
//   - Ctrl/⌘+Alt+E → Czytanie ⇄ Edycja; Esc w akapicie → wyjście z pisania do dokumentu,
//   - Ctrl/⌘+Alt+1 → panel otwarty, fokus w „Znajdź ustawienie…”; Ctrl/⌘+Alt+3 → dokument,
//   - schowany panel nie łapie Tab (inert),
//   - skróty sekcji: jeden przystanek Tab, strzałki w bok,
//   - Ctrl/⌘+Alt+F → tryb skupienia: bez nagłówka / paska stanu / panelu, „Zapisz” na pasku,
//     dokument WYŻSZY; Esc wychodzi, wszystko wraca na miejsce,
//   - (Chromium) najechanie myszą → podpowiedź z tekstem z data-hint-pl.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const active = (page) => page.evaluate(() => {
  const a = document.activeElement;
  return a ? (a.id || (a.closest(".docx-editable-p") ? "paragraph" : a.className || a.tagName)) : "";
});

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.setItem("dwb-panel-docked-open-v1", "0"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.waitForTimeout(600);

  // skip link
  await page.evaluate(() => document.activeElement?.blur());
  await page.keyboard.press("Tab");
  const first = await active(page);
  if (ENGINE === "webkit") {
    // Safari domyślnie Tabem chodzi tylko po polach formularzy (linki pomija, dopóki
    // nie włączysz tego w ustawieniach) — tam pierwszy przystanek to szukanie, a do
    // dokumentu prowadzi Ctrl/⌘+Alt+3.
    check("Safari: 1. Tab = pole szukania", first === "searchQuery", first);
    await page.keyboard.press("Control+Alt+Digit3");
    check("Ctrl+Alt+3 → dokument", (await active(page)) === "docViewport", await active(page));
  } else {
    await page.keyboard.press("Enter");
    check("1. Tab = „Przejdź do dokumentu”, Enter → dokument", first === "skipToDoc" && (await active(page)) === "docViewport", `${first} → ${await active(page)}`);
  }

  // PageDown przewija dokument z fokusem
  const st0 = await page.evaluate(() => document.getElementById("docViewport").scrollTop);
  await page.keyboard.press("PageDown");
  await page.waitForTimeout(300);
  check("PageDown przewija dokument", (await page.evaluate(() => document.getElementById("docViewport").scrollTop)) > st0);

  // schowany panel poza Tab
  check("schowany panel: inert (Tab go omija)", await page.evaluate(() => document.querySelector(".sidebar").inert === true));

  // Ctrl+F i szukanie z klawiatury
  await page.keyboard.press("Control+f");
  check("Ctrl+F → pole szukania", (await active(page)) === "searchQuery");
  await page.keyboard.type("Najemca");
  await page.keyboard.press("Enter");
  await page.waitForFunction(() => document.getElementById("searchPos").textContent === "1 / 21", null, { timeout: 8000 }).catch(() => {});
  await page.keyboard.press("Enter");
  await page.waitForTimeout(150);
  const p2 = await page.textContent("#searchPos");
  await page.keyboard.press("Shift+Enter");
  await page.waitForTimeout(150);
  const p1 = await page.textContent("#searchPos");
  await page.keyboard.press("F3");
  await page.waitForTimeout(150);
  const p3 = await page.textContent("#searchPos");
  check("Enter = następne, Shift+Enter = poprzednie, F3 = następne", p2 === "2 / 21" && p1 === "1 / 21" && p3 === "2 / 21", `${p2} ${p1} ${p3}`);

  // Czytanie ⇄ Edycja
  await page.keyboard.press("Control+Alt+KeyE");
  await page.waitForTimeout(150);
  const editOn = await page.evaluate(() => !readOnlyMode && document.querySelector('.mode-btn[data-mode="edit"]').getAttribute("aria-pressed") === "true");
  check("Ctrl+Alt+E → Edycja", editOn);
  await page.keyboard.press("Control+Alt+Digit3");
  await page.waitForTimeout(150);
  check("Ctrl+Alt+3 w Edycji → kursor w akapicie", (await active(page)) === "paragraph", await active(page));
  await page.keyboard.press("Escape");
  check("Esc w akapicie → wyjście z pisania do dokumentu", (await active(page)) === "docViewport", await active(page));
  await page.keyboard.press("Control+Alt+KeyE");
  await page.waitForTimeout(100);
  check("Ctrl+Alt+E ponownie → Czytanie", await page.evaluate(() => readOnlyMode));

  // Esc czyści podświetlenia szukania
  await page.keyboard.press("Escape");
  check("Esc (dokument) czyści podświetlenia szukania", await page.evaluate(() => !document.querySelector(".search-hit") && document.getElementById("searchPos").textContent === ""));

  // panel
  await page.keyboard.press("Control+Alt+Digit1");
  await page.waitForTimeout(400);
  check("Ctrl+Alt+1 → panel otwarty, fokus w „Znajdź ustawienie…”", (await page.evaluate(() => isSidebarOpen() && !document.querySelector(".sidebar").inert)) && (await active(page)) === "sidebarFinder", await active(page));

  // skróty sekcji
  const chipTabStops = await page.evaluate(() => [...document.querySelectorAll(".section-chip")].filter((c) => c.tabIndex === 0).length);
  await page.evaluate(() => document.querySelector('.section-chip[tabindex="0"]').focus());
  const c0 = await page.evaluate(() => document.activeElement.textContent);
  await page.keyboard.press("ArrowRight");
  const c1 = await page.evaluate(() => document.activeElement.textContent);
  check("skróty sekcji: 1 przystanek Tab, → przechodzi dalej", chipTabStops === 1 && c0 !== c1, `${c0} → ${c1}`);

  // tryb skupienia
  const before = await page.evaluate(() => Math.round(document.getElementById("docViewport").getBoundingClientRect().height));
  await page.keyboard.press("Control+Alt+KeyF");
  await page.waitForTimeout(400);
  const fm = await page.evaluate(() => ({
    on: document.documentElement.classList.contains("focus-mode"),
    heroHidden: getComputedStyle(document.querySelector(".hero")).display === "none",
    statusHidden: getComputedStyle(document.querySelector(".status-bar")).display === "none",
    panelClosed: !isSidebarOpen(),
    saveInToolbar: !!document.querySelector("#docToolbar #heroSaveBtn"),
    h: Math.round(document.getElementById("docViewport").getBoundingClientRect().height),
  }));
  check("Ctrl+Alt+F → tryb skupienia: bez nagłówka, paska stanu i panelu; „Zapisz” na pasku", fm.on && fm.heroHidden && fm.statusHidden && fm.panelClosed && fm.saveInToolbar, JSON.stringify(fm));
  check(`tryb skupienia: dokument wyższy (${before} → ${fm.h} px)`, fm.h >= before + 80, `${before} → ${fm.h}`);
  await page.evaluate(() => document.getElementById("docViewport").focus());
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  const back = await page.evaluate(() => ({
    on: document.documentElement.classList.contains("focus-mode"),
    saveInHero: !!document.querySelector(".hero #heroSaveBtn") && !!document.querySelector(".hero #appMenuBtn"),
    panelOpen: isSidebarOpen(),
  }));
  check("Esc wychodzi z trybu skupienia; „Zapisz”/⋯ wracają do nagłówka, panel wraca", !back.on && back.saveInHero && back.panelOpen, JSON.stringify(back));

  // podpowiedź (tylko mysz; WebKit z Playwrighta nie emuluje hover tak samo — Chromium)
  if (ENGINE === "chromium") {
    await page.waitForFunction(() => !!window.MateuszCursorHint, null, { timeout: 6000 }).catch(() => {});
    await page.hover("#focusModeBtn");
    await page.mouse.move(0, 0);
    await page.hover("#focusModeBtn", { position: { x: 10, y: 10 } });
    await page.waitForTimeout(1200);
    const hint = await page.evaluate(() => ({ vis: document.getElementById("cursorHint").classList.contains("is-visible"), text: document.getElementById("cursorHintText").textContent }));
    check("najechanie myszą → podpowiedź z opisem i skrótem", hint.vis && /Tryb skupienia.*Ctrl\/⌘\+Alt\+F/.test(hint.text), JSON.stringify(hint));
  }

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ klawiatura i tryb skupienia [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
