// ux-contract-playwright.js — wspólny kontrakt UX Sheet + Docs (v1.00.05) i wygody z audytu QoL Docs.
//
//   node scripts/ux-contract-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/ux-contract-playwright.js (Safari/iPhone)
//
// Co musi być prawdą:
//   - motyw: Systemowy / Jasny / Ciemny; Systemowy idzie za systemem na żywo, Jasny/Ciemny zostaje,
//   - dymki: kolejne info scalają się w jeden, błąd = role="alert" + ✕, dymek z akcją ma przycisk,
//   - zero systemowych confirm()/prompt(): zamykanie z niezapisanymi = Zapisz / Nie zapisuj / Anuluj
//     (Esc = Anuluj, dokument zostaje), Zamień wszystkie bez pytania + „Cofnij” w dymku cofa,
//   - Ctrl/⌘+/ otwiera listę skrótów (z rejestru), Ctrl/⌘+H — panel zamiany z kursorem w „Zamień na”,
//     menu ⋯ ma Otwórz / Zapisz jako / Skróty / „Wersja vX.YY.ZZ”,
//   - Czytanie: pisanie pokazuje podpowiedź „Edytuj”, przycisk przełącza na Edycję,
//   - pasek stanu: „N z M słów” przy zaznaczeniu, „Strona X z Y” w Widoku desktopowym,
//   - ekran startowy: przycisk zdjęcia, lista „Ostatnio otwierane” (otwiera z listy, ✕ usuwa),
//   - ponowne otwarcie pliku: dymek „wrócić do miejsca…”.

const path = require("path");
const fs = require("fs");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const HEADINGS = path.resolve(__dirname, "../docs/samples/headings-sample.docx");
const PRZEWODNIK = path.resolve(__dirname, "../docs/samples/przewodnik.docx");
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const until = (page, fn, arg, timeout = 4000) => page.waitForFunction(fn, arg, { timeout })
  .catch((e) => { if (!(e instanceof pw.errors.TimeoutError)) throw e; });

async function newPage(browser, opts = {}) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 }, ...opts });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  const natives = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
  page.on("dialog", (d) => { if (d.type() !== "beforeunload") natives.push(`${d.type()}: ${d.message()}`); d.dismiss().catch(() => {}); });
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  return { context, page, errors, natives };
}

async function openFile(page, file, name) {
  await page.locator("#fileInput").setInputFiles({ name, mimeType: DOCX_MIME, buffer: fs.readFileSync(file) });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.waitForTimeout(300);
}

async function main() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const { context, page, errors, natives } = await newPage(browser);

  // ── motyw ────────────────────────────────────────────────────────────────
  await page.emulateMedia({ colorScheme: "dark" });
  await page.click("#appMenuBtn");
  const menu = await page.evaluate(() => ({
    segs: [...document.querySelectorAll("#themeSeg [data-theme-pref]")].map((b) => b.textContent.trim()),
    system: document.querySelector('#themeSeg [data-theme-pref="system"]').getAttribute("aria-pressed"),
    theme: document.documentElement.dataset.theme,
    version: document.getElementById("appMenuBuild").textContent,
    items: ["openMenuItem", "saveAsMenuItem", "shortcutsMenuItem", "guideMenuItem", "photoMenuItem", "exportMenuItem"].every((id) => !!document.getElementById(id)),
    saveAsDisabled: document.getElementById("saveAsMenuItem").disabled,
  }));
  check("menu ⋯: Systemowy / Jasny / Ciemny, domyślnie Systemowy", menu.segs.join("|") === "Systemowy|Jasny|Ciemny" && menu.system === "true", JSON.stringify(menu));
  check("Systemowy: ciemny system → ciemny motyw", menu.theme === "dark", menu.theme);
  check("menu ⋯: „Wersja vX.YY.ZZ” widoczna", /^Wersja v\d+\.\d{2}\.\d{2}$/.test(menu.version), menu.version);
  check("menu ⋯: grupa Plik + Pomoc; Zapisz jako wyszarzone bez dokumentu", menu.items && menu.saveAsDisabled, JSON.stringify(menu));
  await page.emulateMedia({ colorScheme: "light" });
  await until(page, () => document.documentElement.dataset.theme === "light");
  check("Systemowy: zmiana systemu na żywo", (await page.evaluate(() => document.documentElement.dataset.theme)) === "light");
  await page.click('#themeSeg [data-theme-pref="dark"]');
  await page.emulateMedia({ colorScheme: "light" });
  const dark = await page.evaluate(() => ({ t: document.documentElement.dataset.theme, k: localStorage.getItem("documents-workbench-theme") }));
  check("Ciemny: zapisany i niezależny od systemu", dark.t === "dark" && dark.k === "dark", JSON.stringify(dark));
  await page.click('#themeSeg [data-theme-pref="system"]');
  const sys = await page.evaluate(() => ({ t: document.documentElement.dataset.theme, k: localStorage.getItem("documents-workbench-theme") }));
  check("Systemowy: klucz usunięty, motyw z systemu", sys.t === "light" && sys.k === null, JSON.stringify(sys));
  await page.keyboard.press("Escape");

  // ── dymki ────────────────────────────────────────────────────────────────
  const toasts = await page.evaluate(async () => {
    document.querySelectorAll(".toast").forEach((x) => x.remove());
    toast("pierwszy", "info"); toast("drugi", "info"); toast("trzeci", "info");
    const merged = document.querySelectorAll(".toast.info").length;
    const text = document.querySelector(".toast.info .toast-msg").textContent;
    const err = toast("błąd", "error");
    let acted = 0;
    const act = toast("akcja", "success", { action: "Cofnij", onAction: () => { acted++; } });
    act.querySelector(".toast-action").click();
    err.querySelector(".toast-close").click();
    await new Promise((r) => setTimeout(r, 260));
    return { merged, text, role: err.getAttribute("role"), infoRole: document.querySelector(".toast.info")?.getAttribute("role"), acted, errGone: !err.isConnected, actGone: !act.isConnected };
  });
  check("dymki: trzy info pod rząd = jeden dymek z ostatnią treścią", toasts.merged === 1 && toasts.text === "trzeci", JSON.stringify(toasts));
  check("dymki: błąd role=alert, info role=status, ✕ zamyka", toasts.role === "alert" && toasts.infoRole === "status" && toasts.errGone, JSON.stringify(toasts));
  check("dymki: przycisk akcji woła onAction i zamyka dymek", toasts.acted === 1 && toasts.actGone, JSON.stringify(toasts));

  // ── ekran startowy: zdjęcie ──────────────────────────────────────────────
  check("ekran startowy: przycisk „Zdjęcie dokumentu”", await page.evaluate(() => !!document.getElementById("emptyPhotoBtn")?.offsetParent));

  // ── Ctrl/⌘+/ — lista skrótów ─────────────────────────────────────────────
  await page.keyboard.press("Control+Slash");
  await until(page, () => !!document.querySelector("#shortcutsDialog[open]"));
  const sc = await page.evaluate(() => ({
    open: !!document.querySelector("#shortcutsDialog[open]"),
    rows: document.querySelectorAll("#shortcutsDialog .shortcuts-row").length,
    ids: ["replace", "link", "comment", "footnote", "painterCopy", "saveAs", "open", "goToPage"].every((id) => document.querySelector(`#shortcutsDialog [data-shortcut-id="${id}"]`)),
  }));
  check("Ctrl/⌘+/: okno skrótów z rejestru (m.in. Ctrl+H, K, przypis, malarz)", sc.open && sc.rows >= 30 && sc.ids, JSON.stringify(sc));
  await page.keyboard.press("Escape");
  await until(page, () => !document.querySelector("#shortcutsDialog[open]"));
  check("Esc zamyka listę skrótów", await page.evaluate(() => !document.querySelector("#shortcutsDialog[open]")));

  // ── dokument ─────────────────────────────────────────────────────────────
  await openFile(page, HEADINGS, "Umowa-najmu.docx");

  // Czytanie: pisanie → podpowiedź „Edytuj”
  await page.locator(".docx-preview-host p").nth(2).click();
  await page.keyboard.type("x");
  await until(page, () => !!document.querySelector(".read-hint"));
  const hint = await page.evaluate(() => ({ shown: !!document.querySelector(".read-hint"), text: document.querySelector(".read-hint")?.textContent || "", ro: readOnlyMode }));
  check("Czytanie: pisanie pokazuje podpowiedź z „Edytuj” i skrótem", hint.shown && /Edytuj/.test(hint.text) && /E/.test(hint.text) && hint.ro, JSON.stringify(hint));
  await page.click(".read-hint-btn");
  await until(page, () => !readOnlyMode && !document.querySelector(".read-hint"));
  const afterHint = await page.evaluate(() => ({ ro: readOnlyMode, hint: !!document.querySelector(".read-hint"), mode: JSON.parse(localStorage.getItem("dwb-read-mode-v1") || "{}") }));
  check("„Edytuj” przełącza na Edycję i zapamiętuje tryb", !afterHint.ro && !afterHint.hint && afterHint.mode.desktop === "edit", JSON.stringify(afterHint));

  // pasek stanu: słowa zaznaczenia (prawdziwe zaznaczenie myszą — dwuklik w słowo)
  await page.locator(".docx-preview-host p").filter({ hasText: /\S+\s+\S+\s+\S+/ }).first().dblclick({ position: { x: 12, y: 6 } });
  await until(page, () => / z /.test(document.getElementById("statusWords").textContent));
  const words = await page.evaluate(() => document.getElementById("statusWords").textContent);
  check("pasek stanu: „N z M słów” przy zaznaczeniu", /^\d+ z [\d\s  ]+ słów$/.test(words), words);
  await page.evaluate(() => getSelection().removeAllRanges());

  // strona X z Y
  await until(page, () => !document.getElementById("statusPage").hidden);
  const pg = await page.evaluate(() => document.getElementById("statusPage").textContent);
  check("pasek stanu: „Strona X z Y” w Widoku desktopowym", /^Strona \d+ z \d+$/.test(pg), pg);

  // Ctrl/⌘+H — panel zamiany, kursor w „Zamień na”
  await page.fill("#searchQuery", "Najemca");
  await page.locator(".docx-preview-host p").first().click();
  await page.keyboard.press("Control+KeyH");
  await until(page, () => document.activeElement?.id === "frReplace");
  const h = await page.evaluate(() => ({ open: document.getElementById("panel-search").open, focus: document.activeElement?.id }));
  check("Ctrl/⌘+H: Znajdź i zamień, kursor w „Zamień na”", h.open && h.focus === "frReplace", JSON.stringify(h));

  // Zamień wszystkie: bez pytania, „Cofnij” w dymku cofa
  await page.fill("#frReplace", "NAJEMCA-X");
  const before = await page.evaluate(() => document.querySelector(".docx-preview-host").textContent.includes("NAJEMCA-X"));
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((x) => x.remove()));
  await page.click("#frReplaceAllBtn");
  await until(page, () => !!document.querySelector(".toast .toast-action"), null, 8000);
  const replaced = await page.evaluate(() => ({ has: document.querySelector(".docx-preview-host").textContent.includes("NAJEMCA-X"), dialog: !!document.querySelector("dialog.app-dialog[open]") }));
  check("Zamień wszystkie: od razu, bez okna, dymek z „Cofnij”", !before && replaced.has && !replaced.dialog, JSON.stringify(replaced));
  await page.click(".toast .toast-action");
  await until(page, () => !document.querySelector(".docx-preview-host").textContent.includes("NAJEMCA-X"), null, 8000);
  check("„Cofnij” w dymku cofa zamianę", await page.evaluate(() => !document.querySelector(".docx-preview-host").textContent.includes("NAJEMCA-X")));

  // zamykanie z niezapisanymi zmianami: Zapisz / Nie zapisuj / Anuluj
  await until(page, () => !dwbUndo.canUndo() && document.getElementById("loadingOverlay").classList.contains("hidden"), null, 8000); // Cofnij skończone
  await page.waitForTimeout(300);
  await page.evaluate(() => setDirtyState(true));
  await page.evaluate(() => { requestCloseDocument(); });
  await until(page, () => !!document.querySelector("dialog.app-dialog[open]"));
  const dlg = await page.evaluate(() => [...document.querySelectorAll("dialog.app-dialog[open] .app-dialog-actions button")].map((b) => b.textContent));
  check("zamykanie: okno Zapisz / Nie zapisuj / Anuluj (w tej kolejności)", dlg.join("|") === "Zapisz|Nie zapisuj|Anuluj", dlg.join("|"));
  await page.keyboard.press("Escape");
  await until(page, () => !document.querySelector("dialog.app-dialog[open]"));
  check("Esc = Anuluj: dokument zostaje", await page.evaluate(() => !!originalFileBytes && !document.querySelector("dialog.app-dialog[open]")));
  await page.evaluate(() => { requestCloseDocument(); });
  await until(page, () => !!document.querySelector("dialog.app-dialog[open]"));
  await page.click("dialog.app-dialog[open] .app-dialog-actions button:nth-child(2)");
  await until(page, () => !originalFileBytes);
  check("„Nie zapisuj” zamyka dokument", await page.evaluate(() => !originalFileBytes));

  // ostatnio otwierane + powrót do miejsca
  await until(page, () => !!document.querySelector("#recentFiles .recent-open"));
  const rec = await page.evaluate(() => [...document.querySelectorAll("#recentFiles .recent-name")].map((x) => x.textContent));
  check("ekran startowy: „Ostatnio otwierane” z zamkniętym plikiem", rec.includes("Umowa-najmu.docx"), JSON.stringify(rec));
  await openFile(page, PRZEWODNIK, "Przewodnik-test.docx");
  await page.evaluate(() => {
    const ps = docBodyParagraphs(docCanvasEl);
    ps[Math.min(40, ps.length - 1)].scrollIntoView({ block: "start" });
    dwbComfort.rememberPlace();
  });
  await page.evaluate(() => { setDirtyState(false); requestCloseDocument(); });
  await until(page, () => !originalFileBytes);
  await until(page, () => [...document.querySelectorAll("#recentFiles .recent-name")].some((x) => x.textContent === "Przewodnik-test.docx"));
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((x) => x.remove()));
  await page.evaluate(() => [...document.querySelectorAll("#recentFiles .recent-open")].find((b) => /Przewodnik-test/.test(b.textContent)).click());
  await page.waitForFunction(() => !!originalFileBytes && currentFileName === "Przewodnik-test.docx", null, { timeout: 30000 });
  await until(page, () => [...document.querySelectorAll(".toast")].some((x) => /skończyłeś/.test(x.textContent)), null, 6000);
  const resume = await page.evaluate(() => [...document.querySelectorAll(".toast")].some((x) => /skończyłeś/.test(x.textContent) && x.querySelector(".toast-action")));
  check("otwarcie z listy + dymek „wrócić do miejsca…” z „Przejdź”", resume);
  if (resume) {
    const saved = await page.evaluate(() => JSON.parse(localStorage.getItem("dwb-resume-v1") || "[]").find((x) => x.id.startsWith("Przewodnik-test.docx|"))?.index);
    await page.evaluate(() => [...document.querySelectorAll(".toast")].find((x) => /skończyłeś/.test(x.textContent)).querySelector(".toast-action").click());
    await page.waitForTimeout(1500);
    const now = await page.evaluate(() => captureDocScrollAnchor()?.index);
    check("„Przejdź” przewija do zapamiętanego akapitu", saved > 3 && Math.abs(now - saved) <= 1, `${saved} → ${now}`);
  }

  check("żadnych systemowych confirm()/prompt()/alert()", natives.length === 0, natives.join(" | "));
  check("bez błędów w konsoli", errors.length === 0, errors.slice(0, 3).join(" | "));
  await context.close();

  // ── telefon: uchwyt panelu nie zasłania tekstu ────────────────────────────
  const ph = await newPage(browser, { viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: ENGINE === "chromium", deviceScaleFactor: 2 });
  await openFile(ph.page, HEADINGS, "Umowa-najmu.docx");
  const handle = await ph.page.evaluate(() => getComputedStyle(document.getElementById("panelHandle")).display);
  check("telefon z dokumentem: zamknięty uchwyt panelu schowany (nie zasłania liter)", handle === "none", handle);
  check("telefon: bez błędów w konsoli", ph.errors.length === 0, ph.errors.slice(0, 3).join(" | "));
  await ph.context.close();
  await browser.close();

  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail || ""}`}`);
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}

main().catch((e) => {
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail || ""}`}`);
  console.error(e);
  process.exit(1);
});
