// hrule2-playwright.js — linia pozioma i pusta linijka pod nią, linia „pod tekstem”,
// pogrubienie wyłączone przed Enterem.
//
//   node scripts/hrule2-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/hrule2-playwright.js (Safari/iPhone)
//
// Zgłoszenia 2026-11-08:
//  1. Backspace w pustej linijce POD linią usuwał linię, a pusta linijka zostawała. Word: znika
//     pusta linijka (pusty akapit), linia zostaje — kolejny Backspace usuwa dopiero linię.
//  2. Tekst wpisany w akapicie linii (linia „pod napisem”) — linii nie dało się zdjąć bez
//     kasowania tekstu. Teraz menu kontekstowe ma „Usuń linię poziomą”; Backspace na początku
//     skleja akapity jak w Wordzie (linii nie zdejmuje).
//  3. Pogrubienie wyłączone (⌘/Ctrl+B) na końcu pogrubionego tekstu, potem Enter — w nowej
//     linijce pisało się znów pogrubionym. Wyłączenie obowiązuje też w nowym akapicie (Word).
const fs = require("fs");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const SAMPLE = path.resolve(__dirname, "../docs/samples/headings-sample.docx");
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(350));

// plik: liczba akapitów, teksty akapitów z krawędzią, teksty wszystkich akapitów
const fileState = (page) => page.evaluate(async () => {
  const zip = await JSZip.loadAsync(await buildDocumentForSave());
  const xml = await zip.file("word/document.xml").async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const ps = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
  return {
    n: ps.length,
    rules: ps.filter((p) => p.getElementsByTagNameNS(W_NS, "pBdr").length).map((p) => p.textContent),
    texts: ps.map((p) => p.textContent),
  };
});
const runBold = (page, text) => page.evaluate(async (txt) => {
  const zip = await JSZip.loadAsync(await buildDocumentForSave());
  const xml = await zip.file("word/document.xml").async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const r = Array.from(doc.getElementsByTagNameNS(W_NS, "r")).find((x) => x.textContent === txt || x.textContent.includes(txt));
  if (!r) return "brak";
  const b = Array.from(r.getElementsByTagNameNS(W_NS, "b"))[0];
  return !!b && !/^(0|false|off)$/.test(b.getAttributeNS(W_NS, "val") || "");
}, text);
const viewParas = (page) => page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))
  .map((p) => ({ t: p.textContent.replace(/﻿/g, ""), rule: getComputedStyle(p).borderBottomStyle !== "none" })));
const caretAt = (page, i, end = false) => page.evaluate(([idx, e]) => {
  const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
  placeCaret(ps[idx], e ? ps[idx].textContent.length : 0);
}, [i, end]);

async function main() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "linia2.docx", mimeType: DOCX_MIME, buffer: fs.readFileSync(SAMPLE) });
  await idle(page);
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await page.waitForTimeout(400);
  const base = await fileState(page);

  // ── 1. linia + pusta linijka pod nią: Backspace w pustej linijce usuwa ją, linia zostaje
  await caretAt(page, 1, true);
  await page.evaluate(() => composeUi.insertHrule());
  await idle(page);
  let v = await viewParas(page);
  check("wstawiona linia, pod nią pusta linijka", v[2]?.rule && !v[3]?.rule && !v[3]?.t, JSON.stringify(v.slice(1, 5)));
  await caretAt(page, 3);
  await page.keyboard.press("Backspace");
  await idle(page);
  let f = await fileState(page);
  v = await viewParas(page);
  check("Backspace w pustej linijce pod linią: linia zostaje", f.rules.length === base.rules.length + 1 && v[2]?.rule, JSON.stringify({ rules: f.rules, v: v.slice(1, 4) }));
  check("…a pusta linijka znika (plik i ekran)", f.n === base.n + 1 && v.length === base.n + 1, `${f.n} / ${v.length} / bazowo ${base.n}`);
  // drugi Backspace (kursor na końcu linii) — linia znika
  await page.keyboard.press("Backspace");
  await idle(page);
  f = await fileState(page);
  check("kolejny Backspace usuwa linię", f.rules.length === base.rules.length && f.n === base.n, JSON.stringify({ rules: f.rules, n: f.n }));

  // ── 2. tekst w akapicie linii — Backspace na początku skleja jak w Wordzie, menu zdejmuje linię
  await caretAt(page, 1, true);
  await page.evaluate(() => composeUi.insertHrule());
  await idle(page);
  await caretAt(page, 2);
  await page.keyboard.type("Podpis");
  await idle(page);
  f = await fileState(page);
  check("tekst wpisany w linii: akapit z tekstem i krawędzią", f.rules.includes("Podpis"), JSON.stringify(f.rules));
  // kliknięcie w dokument zamyka krok pisania (undo.js) — sklejenie to osobny krok; dawniej pauza
  // 1,7 s, którą runner skracał o połowę (SLEEP_SCALE=0.5) poniżej progu 1,5 s
  await page.locator(".docx-preview-host p", { hasText: "Podpis" }).first().click();
  await caretAt(page, 2);
  await page.keyboard.press("Backspace");
  await idle(page);
  f = await fileState(page);
  check("Backspace na początku tekstu z linią: skleja z akapitem wyżej (jak w Wordzie)", !f.texts.includes("Podpis") && f.texts.some((x) => x.endsWith("Podpis")), JSON.stringify(f.texts.slice(0, 4)));
  await page.click("#undoBtn");
  await idle(page);
  f = await fileState(page);
  check("Cofnij przywraca linię pod tekstem", f.rules.includes("Podpis"), JSON.stringify(f.rules));
  const para = page.locator(".docx-preview-host p", { hasText: "Podpis" }).first();
  await para.click({ button: "right" });
  const shown = await page.evaluate(() => [...document.querySelectorAll("[role=menuitem]")].some((b) => /Usuń linię poziomą|Delete horizontal line/.test(b.textContent)));
  check("menu kontekstowe przy tekście z linią: „Usuń linię poziomą”", shown);
  if (shown) {
    await page.evaluate(() => [...document.querySelectorAll("[role=menuitem]")].find((b) => /Usuń linię poziomą|Delete horizontal line/.test(b.textContent)).click());
    await idle(page);
    f = await fileState(page);
    check("…usuwa linię, tekst zostaje", !f.rules.includes("Podpis") && f.texts.includes("Podpis"), JSON.stringify(f.rules));
  }

  // ── 3. pogrubienie wyłączone przed Enterem
  await caretAt(page, 1, true);
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.type(" Grube");
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.press("Enter");
  await page.keyboard.type("Cienkie");
  await idle(page);
  check("pogrubiony tekst przed Enterem", (await runBold(page, "Grube")) === true, String(await runBold(page, "Grube")));
  const thin = await runBold(page, "Cienkie");
  const thinView = await page.evaluate(() => {
    const p = [...document.querySelectorAll(".docx-preview-host p")].find((x) => x.textContent.includes("Cienkie"));
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let max = -1; // najgrubszy fragment tekstu akapitu (litery mogą być w kilku węzłach)
    for (let n = w.nextNode(); n; n = w.nextNode()) if (n.data.replace(/\uFEFF/g, "").trim()) max = Math.max(max, parseInt(getComputedStyle(n.parentElement).fontWeight, 10) || 400);
    return max;
  });
  check("wyłączone pogrubienie + Enter: nowa linijka bez pogrubienia (plik)", thin === false, String(thin));
  check("…i na ekranie", thinView > 0 && thinView < 600, String(thinView));
  // włączone pogrubienie w pustym miejscu + Enter → dalej pogrubione
  await page.keyboard.press("Enter");
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.press("Enter");
  await page.keyboard.type("Znowu");
  await idle(page);
  check("włączone pogrubienie + Enter: nowa linijka pogrubiona", (await runBold(page, "Znowu")) === true, String(await runBold(page, "Znowu")));

  check("bez błędów w konsoli", errors.length === 0, errors.slice(0, 3).join(" | "));
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
