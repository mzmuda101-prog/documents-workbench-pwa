// launch-files-playwright.js — „Otwórz za pomocą” (file_handlers + launchQueue) i
// „Udostępnij” (share_target + sw.js) z app/launch-files.js.
//
// Systemowej integracji (menu Windows / Finder / Android) Playwright nie kliknie — sprawdzamy
// to, co od nas zależy: manifest bez błędów w Chromium, odbiór pliku z launchQueue (z uchwytem
// → zapis do oryginału), prawdziwy POST na ./share-target przez Service Workera, a w WebKit
// (Safari/iPhone — brak tych API) zwykły start bez zmian. ENGINE=webkit = ścieżka Safari.

const pw = require("playwright");
const fs = require("fs");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(300));
const SAMPLE = fs.readFileSync(path.join(__dirname, "../docs/samples/forms-sample.docx"));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const errors = [];

  if (ENGINE === "webkit") {
    const ctx = await browser.newContext({ serviceWorkers: "block" });
    await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
    const page = await ctx.newPage();
    page.on("pageerror", (e) => errors.push(e.message));
    await page.goto(`${APP_URL}?open=file`, { waitUntil: "load" });
    await page.waitForTimeout(1500);
    const st = await page.evaluate(() => ({ launchQueue: "launchQueue" in window, search: location.search, empty: !originalFileBytes }));
    check("Safari: brak launchQueue → zwykły start, adres wyczyszczony, bez błędów", !st.launchQueue && st.search === "" && st.empty, JSON.stringify(st));
  } else {
    // 1) manifest — Chromium go parsuje bez błędów
    const ctx1 = await browser.newContext({ serviceWorkers: "block" });
    await ctx1.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
    const p1 = await ctx1.newPage();
    await p1.goto(APP_URL, { waitUntil: "load" });
    const cdp = await ctx1.newCDPSession(p1);
    const man = await cdp.send("Page.getAppManifest");
    const json = JSON.parse(man.data || "{}");
    check("manifest: Chromium bez błędów parsowania", !man.errors?.length, JSON.stringify(man.errors));
    check("manifest: .docx w „Otwórz za pomocą”, każde otwarcie w nowym oknie, „Udostępnij” POST",
      json.file_handlers?.[0]?.accept?.["application/vnd.openxmlformats-officedocument.wordprocessingml.document"]?.includes(".docx")
      && json.launch_handler?.client_mode === "navigate-new" && json.share_target?.method === "POST" && json.share_target?.params?.files?.[0]?.name === "file");
    await ctx1.close();

    // 2) „Otwórz za pomocą”: plik przychodzi z launchQueue jako uchwyt
    const ctx2 = await browser.newContext({ serviceWorkers: "block" });
    await ctx2.addInitScript(() => {
      sessionStorage.setItem("introPlayed", "true");
      // atrapa API (Chrome dostarcza je tylko zainstalowanej aplikacji uruchomionej z pliku)
      // Chromium ma te API natywnie (tylko do odczytu) — przesłaniamy je na obiekcie window
      Object.defineProperty(window, "LaunchParams", { configurable: true, value: class { get files() { return []; } } });
      Object.defineProperty(window, "launchQueue", { configurable: true, value: { setConsumer(fn) { window.__launchConsumer = fn; } } });
    });
    const p2 = await ctx2.newPage();
    p2.on("pageerror", (e) => errors.push(e.message));
    await p2.goto(`${APP_URL}?open=file`, { waitUntil: "load" });
    await p2.waitForFunction(() => typeof window.__launchConsumer === "function", null, { timeout: 10000 });
    await p2.evaluate(async (arr) => {
      const bytes = new Uint8Array(arr);
      const handle = {
        kind: "file", name: "Wniosek.docx",
        getFile: async () => new File([bytes], "Wniosek.docx"),
        queryPermission: async () => "granted", requestPermission: async () => "granted",
        createWritable: async () => ({ write: async (b) => { window.__written = b; }, close: async () => {} }),
      };
      window.__handle = handle;
      await window.__launchConsumer({ files: [handle] });
    }, Array.from(SAMPLE));
    await idle(p2);
    let st = await p2.evaluate(() => ({ name: currentFileName, handle: fileHandle === window.__handle, search: location.search, fields: typeof docFormCounts !== "undefined" ? docFormCounts.fields : -1 }));
    check("Otwórz za pomocą: plik otwarty, z uchwytem (zapis do oryginału), adres wyczyszczony",
      st.name === "Wniosek.docx" && st.handle && st.search === "", JSON.stringify(st));
    check("Otwórz za pomocą: moduły widzą plik (Formularz: 12 pól)", st.fields === 12, String(st.fields));
    await p2.evaluate(() => { window.confirm = () => true; return applyDocumentEdit({ op: "affix", prefix: "", suffix: "!", scope: "all" }); });
    await idle(p2);
    await p2.click("#heroSaveBtn");
    await p2.waitForTimeout(800);
    st = await p2.evaluate(() => ({ written: !!window.__written, dirty: hasUnsavedChanges }));
    check("Otwórz za pomocą → „Zapisz” zapisuje prosto do oryginału", st.written && !st.dirty, JSON.stringify(st));
    await ctx2.close();

    // 3) „Udostępnij” (Android): prawdziwy POST przez Service Workera
    const ctx3 = await browser.newContext();
    await ctx3.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
    const p3 = await ctx3.newPage();
    p3.on("pageerror", (e) => errors.push(e.message));
    await p3.goto(APP_URL, { waitUntil: "load" });
    await p3.evaluate(() => navigator.serviceWorker.ready);
    await p3.reload({ waitUntil: "load" });
    await p3.waitForFunction(() => !!navigator.serviceWorker.controller, null, { timeout: 15000 });
    const shared = await p3.evaluate(async (arr) => {
      const fd = new FormData();
      fd.append("file", new File([new Uint8Array(arr)], "Udostępniony wniosek.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
      const res = await fetch("./share-target", { method: "POST", body: fd });
      return { redirected: res.redirected, url: res.url, ok: res.ok };
    }, Array.from(SAMPLE));
    check("Udostępnij: Service Worker przyjmuje plik i kieruje na ?open=shared", shared.redirected && /\?open=shared$/.test(shared.url) && shared.ok, JSON.stringify(shared));
    await p3.goto(`${APP_URL}?open=shared`, { waitUntil: "load" });
    await p3.waitForFunction(() => !!originalFileBytes, null, { timeout: 20000 });
    await idle(p3);
    st = await p3.evaluate(async () => ({ name: currentFileName, handle: !!fileHandle, search: location.search, left: !!(await (await caches.open("docs-wb-share")).match("./__shared-file")) }));
    check("Udostępnij: plik otwarty z nazwą, bez uchwytu (zapis = kopia), schowek po nim wyczyszczony",
      st.name === "Udostępniony wniosek.docx" && !st.handle && st.search === "" && !st.left, JSON.stringify(st));
    await p3.reload({ waitUntil: "load" });
    await p3.waitForTimeout(1200);
    check("odświeżenie strony nie otwiera udostępnionego pliku drugi raz", await p3.evaluate(() => !originalFileBytes));
    const offline = await p3.evaluate(async () => (await fetch("./index.html")).ok);
    check("zwykłe zapytania SW działają jak dawniej", offline);
    await ctx3.close();
  }

  await browser.close();
  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ Otwórz za pomocą / Udostępnij [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
