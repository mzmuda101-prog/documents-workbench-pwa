// caret-nav-playwright.js — poruszanie się po dokumencie w Edycji jak w Wordzie (app/doc-caret-nav.js)
// + znacznik bieżącego akapitu poza tekstem.
//   node scripts/caret-nav-playwright.js / ENGINE=webkit node scripts/caret-nav-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1366, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.waitForTimeout(500);
  await page.click('.mode-btn[data-mode="edit"]');
  await page.waitForTimeout(400);

  const where = () => page.evaluate(() => {
    const s = getSelection(); const n = s.anchorNode;
    const el = (n?.nodeType === 1 ? n : n?.parentElement)?.closest(".docx-editable-p");
    return { i: collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).indexOf(el), o: el ? getCaretOffset(el) : -1 };
  });
  const put = (i, end) => page.evaluate(([i, end]) => {
    const el = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    el.focus(); const r = document.createRange(); r.selectNodeContents(el); r.collapse(!end);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  }, [i, end]);
  const fileTexts = () => page.evaluate(async () => {
    await waitInlineStructuralIdle();
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
    return collectParagraphElements(d.documentElement, "all").map(paragraphSearchText);
  });

  let w;
  await put(1, true); await page.keyboard.press("ArrowDown"); w = await where();
  check("↓ z ostatniego wiersza → następny akapit", w.i === 2, JSON.stringify(w));
  await put(2, false); await page.keyboard.press("ArrowUp"); w = await where();
  check("↑ z pierwszego wiersza → poprzedni akapit", w.i === 1, JSON.stringify(w));
  await put(1, true); await page.keyboard.press("ArrowRight"); w = await where();
  check("→ na końcu → początek następnego", w.i === 2 && w.o === 0, JSON.stringify(w));
  await put(2, false); await page.keyboard.press("ArrowLeft"); w = await where();
  check("← na początku → koniec poprzedniego", w.i === 1 && w.o > 0, JSON.stringify(w));

  // długi akapit (kilka wierszy): ↓ z pierwszego wiersza zostaje w akapicie
  await put(3, false); await page.keyboard.press("ArrowDown"); w = await where();
  check("↓ w środku długiego akapitu przesuwa o wiersz, nie skacze", w.i === 3 && w.o > 0, JSON.stringify(w));

  const before = await fileTexts();
  await put(1, true); await page.keyboard.press("Delete"); await page.waitForTimeout(600);
  let after = await fileTexts();
  w = await where();
  check("Delete na końcu akapitu dołącza następny (w pliku)", after.length === before.length - 1 && after[1] === before[1] + before[2], `${before.length}→${after.length}: ${after[1]}`);
  check("…kursor w miejscu połączenia", w.i === 1 && w.o === before[1].length, JSON.stringify(w));
  await page.click("#undoBtn"); await page.waitForTimeout(800);
  after = await fileTexts();
  check("↶ cofa połączenie", after.length === before.length && after[2] === before[2], String(after.length));

  // klik na marginesie strony
  const box = await page.evaluate(() => {
    const el = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[3];
    el.scrollIntoView({ block: "center" });
    const r = el.getBoundingClientRect(); const sec = el.closest("section").getBoundingClientRect();
    return { l: sec.left + 8, r: sec.right - 8, y: r.top + 6 };
  });
  await page.mouse.click(box.l, box.y); w = await where();
  check("klik na lewym marginesie → kursor na początku wiersza akapitu", w.i === 3 && w.o === 0, JSON.stringify(w));
  await page.mouse.click(box.r, box.y); w = await where();
  check("klik na prawym marginesie → kursor na końcu wiersza", w.i === 3 && w.o > 0, JSON.stringify(w));
  await page.keyboard.type("X");
  await page.waitForTimeout(200);
  check("po kliknięciu na marginesie można od razu pisać", (await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[3].textContent)).includes("X"));

  // Ctrl+Home / Ctrl+End (⌘↑ / ⌘↓ na Macu)
  const mac = await page.evaluate(() => /Mac/.test(navigator.platform));
  await put(5, false);
  await page.keyboard.press(mac ? "Meta+ArrowDown" : "Control+End"); w = await where();
  const last = await page.evaluate(() => [...collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))].filter((p) => p.isContentEditable).length - 1);
  check("koniec dokumentu (Ctrl+End / ⌘↓)", w.i >= last - 1, `${w.i} vs ${last}`);
  await page.keyboard.press(mac ? "Meta+ArrowUp" : "Control+Home"); w = await where();
  check("początek dokumentu (Ctrl+Home / ⌘↑)", w.i === 0 && w.o === 0, JSON.stringify(w));

  // znacznik bieżącego akapitu nie leży na tekście
  const bar = await page.evaluate(() => {
    const el = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[3];
    el.focus();
    return getComputedStyle(el).boxShadow;
  });
  check("pasek bieżącego akapitu jest poza tekstem (bez „inset”)", bar && bar !== "none" && !/inset/.test(bar), bar);

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`); if (!r.ok) failed++; }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ nawigacja kursora [${ENGINE}]: ${results.length}/${results.length}`);
}
run().catch((e) => { console.error(e); process.exit(1); });
