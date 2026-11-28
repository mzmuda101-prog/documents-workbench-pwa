// phone-bar-playwright.js — górny pasek telefonu (app/phone-bar.js, etap 7 wstążki, makieta v4).
// ← zamyka do ekranu startowego, Cofnij w Edycji, „N niezapisane · tryb”, menu trybu, Zapisz w ⋯,
// w Czytaniu bez pływającego „Edytuj”, szeroko nic z tego nie widać. ENGINE=webkit.

const pw = require("playwright");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const errors = [];

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 800 }, hasTouch: true, ...(ENGINE === "chromium" ? { isMobile: true } : {}) });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
  const page = await context.newPage();
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", path.join(__dirname, "..", "docs", "samples", "sample.docx"));
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  await sleep(600);
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  const vis = (sel) => page.evaluate((s) => { const e = document.querySelector(s); return !!e && getComputedStyle(e).display !== "none" && e.getBoundingClientRect().width > 0; }, sel);
  const st = () => page.evaluate(() => document.getElementById("phStatus").textContent);

  check("Czytanie: ←, tryb i ⋯ u góry, bez Cofnij i bez Zapisz", (await vis("#phBackBtn")) && (await vis("#phModeBtn")) && !(await vis("#phUndoBtn")) && !(await vis("#heroSaveBtn")) && (await st()) === "Czytanie");
  // Czytanie: stuknięcie w tekst nie pokazuje pływającego „Edytuj”
  await page.locator(".docx-preview-host p").first().tap();
  await sleep(500);
  check("Czytanie: bez pływającego „Edytuj”", !(await page.evaluate(() => !!document.querySelector(".read-hint"))));

  await page.tap("#phModeBtn");
  await sleep(250);
  const items = await page.evaluate(() => Array.from(document.querySelectorAll(".compose-pop [role=menuitemradio]")).map((b) => `${b.textContent.trim().slice(0, 12)}:${b.getAttribute("aria-checked")}`));
  check("menu trybu: Edycja / Sugerowanie / Czytanie (zaznaczone)", items.length === 3 && /^Edycja/.test(items[0]) && /^Czytanie.*:true$/.test(items[2]), JSON.stringify(items));
  await page.locator(".compose-pop [role=menuitemradio]").first().tap();
  await sleep(500);
  check("Edycja: Cofnij u góry (nieaktywne), status „Edycja”", (await vis("#phUndoBtn")) && (await page.evaluate(() => document.getElementById("phUndoBtn").disabled)) && (await st()) === "Edycja");

  await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.isContentEditable && p.textContent.trim().length > 3);
    placeCaret(ps[0], ps[0].textContent.length);
  });
  await page.keyboard.type(" XYZ", { delay: 10 });
  await sleep(600);
  const dirty = await page.evaluate(() => ({ st: document.getElementById("phStatus").textContent, undo: document.getElementById("phUndoBtn").disabled }));
  check("po pisaniu: „1 niezapisane · Edycja”, Cofnij aktywne", dirty.st === "1 niezapisane · Edycja" && !dirty.undo, JSON.stringify(dirty));
  const isl = await page.evaluate(() => ({ save: !!document.getElementById("phIslandSave").offsetParent, n: document.querySelector("#phIslandSave .ph-isl-n").textContent, dirty: document.getElementById("phIslandSave").classList.contains("is-dirty"), mode: document.getElementById("phIslandMode").dataset.mode, focus: !!document.querySelector(".doc-island #focusModeBtn")?.offsetParent, pills: getComputedStyle(document.querySelector(".doc-island-status")).display }));
  check("wyspa: Zapisz z liczbą zmian, tryb Edycja, bez skupienia i bez pigułek", isl.save && isl.n === "1" && isl.dirty && isl.mode === "edit" && !isl.focus && isl.pills === "none", JSON.stringify(isl));
  await page.tap("#phUndoBtn");
  await sleep(600);
  const undone = await page.evaluate(() => ({ text: document.querySelector(".docx-preview-host").textContent.includes("XYZ"), st: document.getElementById("phStatus").textContent }));
  check("Cofnij u góry cofa pisanie", !undone.text && undone.st === "Edycja", JSON.stringify(undone));

  // ── 7b: rząd poleceń i arkusz kart ─────────────────────────────────────────
  check("Edycja: rząd poleceń u dołu, karta Główne, dawny pasek nad dokumentem schowany",
    (await vis("#phCmd")) && (await page.evaluate(() => document.querySelector(".ph-tsw-t").textContent)) === "Główne" && !(await vis("#formatToolbar")) && !(await vis(".doc-toolbar-row-2")));
  // #14: przy otwartej i schowanej klawiaturze (.is-kb) przyciski rzędu trzymają ten sam X
  const kbX = await page.evaluate(() => {
    const cmd = document.getElementById("phCmd");
    const tsw = document.querySelector(".ph-tsw");
    const b0 = document.querySelector(".ph-strip button");
    const xClosed = { tsw: Math.round(tsw.getBoundingClientRect().left * 10) / 10, b0: Math.round(b0.getBoundingClientRect().left * 10) / 10 };
    cmd.classList.add("is-kb");
    const xOpen = { tsw: Math.round(tsw.getBoundingClientRect().left * 10) / 10, b0: Math.round(b0.getBoundingClientRect().left * 10) / 10 };
    cmd.classList.remove("is-kb");
    return { xClosed, xOpen };
  });
  check("rząd: przyciski nie skaczą w bok przy chowaniu klawiatury (.is-kb)",
    Math.abs(kbX.xClosed.tsw - kbX.xOpen.tsw) <= 0.5 && Math.abs(kbX.xClosed.b0 - kbX.xOpen.b0) <= 0.5,
    JSON.stringify(kbX));
  const selectWord = () => page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => x.isContentEditable && /\w{6,}/.test(x.textContent));
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode()) && !/\w{6,}/.test(n.textContent));
    const m = n.textContent.match(/\w{6,}/); const r = document.createRange(); r.setStart(n, m.index); r.setEnd(n, m.index + m[0].length);
    p.focus(); const s = getSelection(); s.removeAllRanges(); s.addRange(r); return m[0];
  });
  const boldNow = () => page.evaluate(() => { const s = getSelection(); return s.rangeCount ? Number(getComputedStyle(s.getRangeAt(0).startContainer.parentElement).fontWeight) : 0; });
  const word = await selectWord();
  await sleep(200);
  await page.tap('.ph-strip [data-act="click:#fmtBold"]');
  await sleep(600);
  const b1 = { w: await boldNow(), pressed: await page.evaluate(() => document.querySelector('.ph-strip [data-act="click:#fmtBold"]').getAttribute("aria-pressed")) };
  check("rząd: B pogrubia zaznaczenie i jest wciśnięte", b1.w >= 600 && b1.pressed === "true", JSON.stringify({ word, ...b1 }));
  await page.tap(".ph-tsw");
  await sleep(400);
  const sh = await page.evaluate(() => ({ sheet: !document.querySelector(".ph-sheet").hidden, row: !document.querySelector(".ph-row").hidden, chips: [...document.querySelectorAll(".ph-chip")].map((c) => c.textContent), items: document.querySelectorAll(".ph-sheet-b .ph-item").length }));
  check("arkusz: karty jako chipy, polecenia z podpisami, rząd schowany", sh.sheet && !sh.row && sh.chips.join("|") === "Główne|Wstaw|Układ|Recenzja|Widok|Asystent" && sh.items >= 10, JSON.stringify(sh));
  await page.tap('.ph-sheet-b [data-act="click:#fmtBold"]');
  await sleep(600);
  check("arkusz: Pogrubienie działa na tym samym zaznaczeniu (zdejmuje)", (await boldNow()) < 600 && (await page.evaluate(() => getSelection().toString())) === word);
  // Cofnij obok Ponów (górny pasek z „Cofnij” bywa schowany przy otwartym arkuszu)
  const firstTwo = await page.evaluate(() => [...document.querySelectorAll(".ph-sheet-b .ph-item span")].slice(0, 2).map((x) => x.textContent));
  await page.tap('.ph-sheet-b [data-act="click:#undoBtn"]');
  await sleep(600);
  check("arkusz Główne: Cofnij i Ponów na początku, Cofnij cofa (pogrubienie wraca)", firstTwo.join("|") === "Cofnij|Ponów" && (await boldNow()) >= 600, JSON.stringify({ firstTwo, w: await boldNow() }));
  await page.tap('.ph-chip[data-tab="insert"]');
  await sleep(200);
  check("arkusz: zmiana karty pokazuje jej polecenia", (await page.evaluate(() => document.querySelector(".ph-sheet-b .ph-item span").textContent)) === "Tabela");
  await page.tap(".ph-done");
  await sleep(300);
  check("Gotowe: z powrotem rząd (karta Wstaw)", await page.evaluate(() => document.querySelector(".ph-sheet").hidden && !document.querySelector(".ph-row").hidden && document.querySelector(".ph-tsw-t").textContent === "Wstaw"));
  // kursor w tabeli → arkusz: karta Tabela (ostatnia, za Asystentem) widoczna bez przewijania
  // (testerzy AI 2026-11-27, AUDYT-MOBILE #4: dwie osoby nie znalazły dodawania wiersza)
  await page.evaluate(() => composeUi.insertTable(2, 2));
  await page.waitForFunction(() => !document.getElementById("tableToolsBtn")?.hidden, null, { timeout: 10000 });
  await sleep(300);
  await page.tap(".ph-tsw");
  await sleep(400);
  const ctx = await page.evaluate(() => {
    const row = document.querySelector(".ph-chips").getBoundingClientRect();
    const c = document.querySelector('.ph-chip[data-tab="table"]')?.getBoundingClientRect();
    return c ? { left: Math.round(c.left), right: Math.round(c.right), rowL: Math.round(row.left), rowR: Math.round(row.right) } : null;
  });
  check("arkusz z kursorem w tabeli: karta Tabela w całości widoczna w rzędzie kart", !!ctx && ctx.left >= ctx.rowL - 1 && ctx.right <= ctx.rowR + 1, JSON.stringify(ctx));
  await page.tap(".ph-done");
  await sleep(300);
  // wyspa wyłączona (jak w makiecie): rząd stoi przy dolnej krawędzi
  await page.evaluate(() => { const b = document.getElementById("phIslandOpt"); b.checked = false; b.dispatchEvent(new Event("change")); });
  await sleep(500);
  const noIsl = await page.evaluate(() => ({ island: getComputedStyle(document.getElementById("docIsland")).display, gap: Math.round(document.querySelector(".doc-panel").getBoundingClientRect().bottom - document.getElementById("phCmd").getBoundingClientRect().bottom) }));
  check("bez wyspy: wyspa schowana, rząd przy dole", noIsl.island === "none" && noIsl.gap <= 16, JSON.stringify(noIsl));
  await page.evaluate(() => { const b = document.getElementById("phIslandOpt"); b.checked = true; b.dispatchEvent(new Event("change")); });
  await sleep(300);

  await page.tap("#appMenuBtn");
  await sleep(300);
  check("⋯: jest „Zapisz”", await vis("#phSaveMenuItem"));
  await page.evaluate(() => appFrame.setMenuOpen(false));
  await sleep(200);

  await page.tap("#phIslandMode"); // tryb też z wyspy (górny pasek zwija się przy przewijaniu)
  await sleep(250);
  await page.locator(".compose-pop [role=menuitemradio]").nth(2).tap();
  await sleep(500);
  check("Czytanie: bez rzędu poleceń", !(await vis("#phCmd")));
  await page.tap("#phBackBtn");
  await sleep(800);
  check("←: dokument zamknięty, ekran startowy", await page.evaluate(() => !originalFileBytes && !document.querySelector(".hero").classList.contains("has-file")));

  await page.setViewportSize({ width: 1280, height: 800 });
  await sleep(400);
  check("szeroko: bez paska telefonu", !(await vis("#phBackBtn")) && !(await vis("#phModeBtn")));

  await browser.close();
  if (errors.length) check("bez błędów w konsoli", false, errors.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) { if (!r.ok) failed++; console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`); }
  console.log(`\n${failed ? "❌" : "✅"} Pasek telefonu [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}
run().catch((e) => { console.error(e); process.exit(1); });
