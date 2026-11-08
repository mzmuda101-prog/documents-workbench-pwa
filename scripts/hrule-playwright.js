// hrule-playwright.js — linia pozioma (Wstaw → Linia pozioma): da się ją usunąć i nie dubluje się.
//
//   node scripts/hrule-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/hrule-playwright.js (Safari/iPhone)
//
// Zgłoszenie 2026-11-08: wstawionej linii nie dało się usunąć — Delete/Backspace sklejał akapit
// z treścią z pustym akapitem linii, a wynik brał krawędź pustego, więc linia przeskakiwała pod
// następny akapit; Enter w linii robił drugą linię. Teraz jak w Wordzie: pusty akapit znika
// razem ze swoim formatowaniem, akapit z treścią zostaje ze swoim; Enter w linii — nowy akapit bez krawędzi.
const fs = require("fs");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const SAMPLE = path.resolve(__dirname, "../docs/samples/headings-sample.docx");
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(350));

// stan pliku: liczba akapitów z krawędzią + styl i krawędź akapitu z danym tekstem
const fileState = (page, text) => page.evaluate(async (txt) => {
  const zip = await JSZip.loadAsync(await buildDocumentForSave());
  const xml = await zip.file("word/document.xml").async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const ps = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
  const target = ps.find((p) => p.textContent.includes(txt));
  const pPr = target && Array.from(target.childNodes).find((n) => n.localName === "pPr");
  const style = pPr && Array.from(pPr.childNodes).find((n) => n.localName === "pStyle");
  return {
    rules: ps.filter((p) => p.getElementsByTagNameNS(W_NS, "pBdr").length).length,
    style: style ? style.getAttributeNS(W_NS, "val") : "",
    targetBorder: !!(target && target.getElementsByTagNameNS(W_NS, "pBdr").length),
  };
}, text);
// ekran: indeks akapitu z dolną krawędzią (linia)
const viewRuleIdx = (page) => page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))
  .findIndex((p) => getComputedStyle(p).borderBottomStyle !== "none"));
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
  await page.locator("#fileInput").setInputFiles({ name: "linia.docx", mimeType: DOCX_MIME, buffer: fs.readFileSync(SAMPLE) });
  await idle(page);
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await page.waitForTimeout(400);
  const NEXT = "§1 Strony umowy";
  const base = await fileState(page, NEXT);

  // 1. wstawienie
  await caretAt(page, 1, true);
  await page.evaluate(() => composeUi.insertHrule());
  await idle(page);
  let f = await fileState(page, NEXT);
  check("Wstaw → Linia pozioma: jedna linia w pliku i na ekranie", f.rules === base.rules + 1 && (await viewRuleIdx(page)) === 2, JSON.stringify(f));

  // 2. Enter w linii nie robi drugiej linii
  await caretAt(page, 2);
  await page.keyboard.press("Enter");
  await idle(page);
  f = await fileState(page, NEXT);
  check("Enter w linii: dalej jedna linia (nowy akapit bez krawędzi)", f.rules === base.rules + 1, JSON.stringify(f));

  // 3. Delete w linii: linia znika, następny akapit zostaje ze swoim stylem i bez krawędzi
  const li = await viewRuleIdx(page);
  await caretAt(page, li);
  await page.keyboard.press("Delete");
  await idle(page);
  // po Enterze za linią jest pusty akapit — Delete skleja linię z nim; drugi Delete z akapitem z treścią
  if ((await fileState(page, NEXT)).rules > base.rules) { await caretAt(page, await viewRuleIdx(page)); await page.keyboard.press("Delete"); await idle(page); }
  f = await fileState(page, NEXT);
  check("Delete w linii: linia znika z pliku", f.rules === base.rules, JSON.stringify(f));
  check("…następny akapit bez krawędzi i z własnym stylem", !f.targetBorder && f.style === base.style, JSON.stringify({ f, base }));
  check("…na ekranie też bez linii", (await viewRuleIdx(page)) === -1, String(await viewRuleIdx(page)));

  // 4. Cofnij przywraca linię
  await page.click("#undoBtn");
  await idle(page);
  if ((await fileState(page, NEXT)).rules === base.rules) { await page.click("#undoBtn"); await idle(page); }
  f = await fileState(page, NEXT);
  check("Cofnij przywraca linię", f.rules === base.rules + 1, JSON.stringify(f));

  // 5. Backspace na początku linii (za akapitem z treścią): linia znika, poprzedni akapit bez zmian
  const lj = await viewRuleIdx(page);
  await caretAt(page, lj);
  await page.keyboard.press("Backspace");
  await idle(page);
  f = await fileState(page, NEXT);
  check("Backspace na początku linii: linia znika", f.rules === base.rules && (await viewRuleIdx(page)) === -1, JSON.stringify(f));

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
