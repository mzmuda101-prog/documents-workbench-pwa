// state-survival-playwright.js — czy NIEZAPISANE zmiany przeżywają „zwykłe” rzeczy wokół.
//
// Zgłoszenie 2026-10-01 (Windows): okno przeciągnięte na inny monitor → dokument narysował się
// od nowa, wpisany tekst zniknął z ekranu, „Zapisz” dalej świeciło, a zapis byłby BEZ niego.
// Tu: wpisujemy znacznik, robimy jedną czynność, i sprawdzamy, że WSZYSTKIE dotąd wpisane
// znaczniki są i na ekranie, i w pliku do zapisu (buildDocumentForSave), a „Zapisz” świeci.
// Czynności: zmiana szerokości okna (telefon ⇄ komputer), obrót, język, motyw, Czytanie/Edycja,
// zoom, Dopasuj, tryb skupienia, panel, panele z leniwym kodem, szukanie, skok ze Struktury,
// karta w tle / powrót, utrata fokusu okna, druk, pinch, Enter tuż przed zmianą układu,
// Cofnij/Ponów. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(350));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 800 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await idle(page);
  await page.click('.mode-btn[data-mode="edit"]');

  const markers = [];
  let n = 0;
  // wpisz znacznik na końcu ostatniego zwykłego akapitu (zawsze edytowalny, także po przerysowaniu)
  const typeMarker = async () => {
    const m = `ZN${++n}X`;
    await page.evaluate(() => {
      if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); }
      const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.isContentEditable && p.textContent.trim().length > 3);
      const el = ps[ps.length - 1];
      placeCaret(el, el.textContent.length);
    });
    await page.keyboard.type(` ${m}`, { delay: 5 });
    markers.push(m);
  };
  const verify = async (label) => {
    await idle(page);
    const r = await page.evaluate(async (ms) => {
      const view = document.querySelector(".docx-preview-host")?.textContent || "";
      const file = (await extractParagraphTextsFromDocx(await buildDocumentForSave())).join("\n");
      return { lostView: ms.filter((m) => !view.includes(m)), lostFile: ms.filter((m) => !file.includes(m)), dirty: hasUnsavedChanges };
    }, markers);
    check(`${label}: wpisany tekst zostaje na ekranie i w zapisie`, !r.lostView.length && !r.lostFile.length && r.dirty, JSON.stringify(r));
  };
  const step = async (label, action) => {
    await typeMarker();
    await action();
    await verify(label);
  };

  await step("szerokość: komputer → telefon → komputer (inny monitor)", async () => {
    await page.setViewportSize({ width: 700, height: 800 }); await page.waitForTimeout(700);
    await page.setViewportSize({ width: 1280, height: 800 });
  });
  await step("szybkie szarpanie rozmiarem okna (5×)", async () => {
    for (let i = 0; i < 5; i++) { await page.setViewportSize({ width: i % 2 ? 1280 : 700, height: 800 }); await page.waitForTimeout(60); }
    await page.setViewportSize({ width: 1280, height: 800 });
    await page.waitForTimeout(800);
  });
  await step("szerokość: zostaje telefon", async () => { await page.setViewportSize({ width: 600, height: 900 }); });
  await step("obrót (orientationchange) na wąskim ekranie", async () => { await page.evaluate(() => window.dispatchEvent(new Event("orientationchange"))); await page.waitForTimeout(400); });
  await step("telefon: „Dopasuj” ⇄ 100%", async () => {
    await page.evaluate(() => resetZoomTo100()); await idle(page);
    await page.evaluate(() => applyFitToWidth());
  });
  await step("telefon → komputer", async () => { await page.setViewportSize({ width: 1280, height: 800 }); });
  await step("Enter tuż przed zmianą układu", async () => {
    await page.keyboard.press("Enter");
    await page.keyboard.type("nowy akapit", { delay: 5 });
    await page.setViewportSize({ width: 700, height: 800 });
    await page.waitForTimeout(500);
    await page.setViewportSize({ width: 1280, height: 800 });
  });
  await step("język PL → EN → PL", async () => { await page.evaluate(() => { setLanguage("en"); setLanguage("pl"); }); });
  await step("motyw jasny ⇄ ciemny", async () => { await page.evaluate(() => { document.getElementById("themeToggle").click(); document.getElementById("themeToggle").click(); }); });
  await step("Czytanie → Edycja", async () => { await page.click('.mode-btn[data-mode="read"]'); await page.waitForTimeout(200); await page.click('.mode-btn[data-mode="edit"]'); });
  await step("zoom +, −, Dopasuj (komputer)", async () => { await page.click("#zoomInBtn"); await page.click("#zoomOutBtn"); await page.evaluate(() => applyFitToWidth()); });
  await step("pinch (zatwierdzenie zoomu)", async () => { await page.evaluate(() => { commitDocZoom(1.4); commitDocZoom(1); }); });
  await step("tryb skupienia wł./wył.", async () => { await page.click("#focusModeBtn"); await page.waitForTimeout(300); await page.keyboard.press("Escape"); await page.waitForTimeout(300); });
  await step("panel otwarty / zamknięty", async () => { await page.evaluate(() => { setSidebarOpen(true); setSidebarOpen(false); setSidebarOpen(true); }); });
  await step("panele z leniwym kodem (Statystyki, Recenzja, Formularz, Eksport, Metadane)", async () => {
    await page.evaluate(async () => {
      for (const id of ["panel-stats", "panel-review", "panel-forms", "panel-export", "panel-metadata"]) { document.getElementById(id).open = true; }
      for (const k of ["stats", "review", "forms", "export", "metadata"]) await ensureLazyFeature(k);
    });
    await page.waitForTimeout(600);
  });
  await step("szukanie + przejście do trafienia", async () => {
    await page.fill("#searchQuery", "Najemca"); await page.press("#searchQuery", "Enter"); await page.waitForTimeout(400);
    await page.fill("#searchQuery", ""); await page.press("#searchQuery", "Escape");
  });
  await step("skok ze Struktury / skrótu sekcji", async () => { await page.locator(".section-chip").nth(2).click(); await page.waitForTimeout(600); });
  await step("karta w tle i z powrotem", async () => {
    await page.evaluate(() => {
      Object.defineProperty(document, "visibilityState", { configurable: true, get: () => "hidden" });
      document.dispatchEvent(new Event("visibilitychange"));
      window.dispatchEvent(new Event("pagehide"));
      Object.defineProperty(document, "visibilityState", { configurable: true, get: () => "visible" });
      document.dispatchEvent(new Event("visibilitychange"));
      window.dispatchEvent(new Event("pageshow"));
    });
    await page.waitForTimeout(300);
  });
  await step("okno traci i odzyskuje fokus", async () => { await page.evaluate(() => { window.dispatchEvent(new Event("blur")); window.dispatchEvent(new Event("focus")); }); });
  await step("druk (beforeprint / afterprint)", async () => { await page.evaluate(() => { window.dispatchEvent(new Event("beforeprint")); window.dispatchEvent(new Event("afterprint")); }); });
  await step("operacja z panelu (zamiana tekstu)", async () => { await page.evaluate(() => applyDocumentEdit({ op: "replace", find: "Kaucja", replace: "Kaucja", regex: false, scope: "all" })); });

  // „Aktualizuj” / „Odśwież aplikację” przy niezapisanych zmianach: pyta, „Anuluj” = nic się nie dzieje
  await typeMarker();
  const upd = await page.evaluate(async () => {
    window.__stay = "ta sama strona";
    let asked = "";
    window.confirm = (msg) => { asked = msg; return false; };
    await hardRefreshApp();
    const btn = document.getElementById("appUpdateBtn");
    btn.classList.remove("hidden");
    btn.click();
    await new Promise((r) => setTimeout(r, 300));
    return { asked: /niezapisane zmiany/.test(asked), busy: btn.classList.contains("is-busy") };
  });
  await page.waitForTimeout(800);
  const still = await page.evaluate(() => window.__stay);
  check("„Aktualizuj” przy niezapisanych zmianach pyta, „Anuluj” nie przeładowuje", upd.asked && !upd.busy && still === "ta sama strona", JSON.stringify({ ...upd, still }));
  await page.evaluate(() => { window.confirm = () => true; document.getElementById("appUpdateBtn").classList.add("hidden"); });
  await verify("po anulowanej aktualizacji");

  // Cofnij / Ponów: po cofnięciu ostatniego kroku poprzednie zostają, Ponów przywraca ostatni
  await page.waitForTimeout(1700); // pisanie bez przerwy = jeden krok Cofnij (jak w Wordzie) — oddziel
  await typeMarker();
  await page.waitForTimeout(1700); // koniec kroku „pisanie”
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  const last = markers.pop();
  const u = await page.evaluate(async (args) => {
    const view = document.querySelector(".docx-preview-host").textContent;
    return { lastGone: !view.includes(args.last), othersKept: args.ms.every((m) => view.includes(m)) };
  }, { last, ms: markers });
  check("Cofnij cofa tylko ostatnie pisanie, wcześniejsze zostają", u.lastGone && u.othersKept, JSON.stringify(u));
  await page.evaluate(() => document.getElementById("redoBtn").click());
  markers.push(last);
  await verify("Ponów");

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ niezapisane zmiany przeżywają [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
