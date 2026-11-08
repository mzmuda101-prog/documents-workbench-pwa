// doc-menu-playwright.js — menu kontekstowe dokumentu (doc-context-menu.js), jak w Wordzie / menu komórki Sheet.
//
//   node scripts/doc-menu-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/doc-menu-playwright.js (Safari/iPhone)
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
const menuItems = (page) => page.evaluate(() => {
  const m = document.querySelector(".doc-menu:not(.hidden)");
  return m ? { ids: [...m.querySelectorAll(".doc-menu-item")].map((b) => [...b.classList].find((c) => c.startsWith("dmi-")).slice(4)), fmt: m.querySelectorAll(".doc-menu-fmt").length, focus: m.contains(document.activeElement) } : null;
});
const paraBox = (page, i) => page.evaluate((idx) => { const r = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[idx].getBoundingClientRect(); return { x: r.left + 30, y: r.top + r.height / 2 }; }, i);
const fileRules = (page) => page.evaluate(async () => {
  const xml = await (await JSZip.loadAsync(await buildDocumentForSave())).file("word/document.xml").async("string");
  return { rules: (xml.match(/<w:pBdr>/g) || []).length, rows: (xml.match(/<w:tr[ >]/g) || []).length };
});
const clickItem = (page, id) => page.click(`.doc-menu:not(.hidden) .dmi-${id}`);

async function main() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  if (ENGINE === "chromium") await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "menu.docx", mimeType: DOCX_MIME, buffer: fs.readFileSync(SAMPLE) });
  await idle(page);

  // 1. Czytanie: Kopiuj / Zaznacz wszystko / Edytuj
  await page.evaluate(() => { readModeEl.checked = true; readModeEl.dispatchEvent(new Event("change")); });
  let b = await paraBox(page, 1);
  await page.mouse.click(b.x, b.y, { button: "right" });
  let m = await menuItems(page);
  check("Czytanie: menu z „Zaznacz wszystko” i „Edytuj dokument”, bez wklejania", m && m.ids.includes("selAll") && m.ids.includes("edit") && !m.ids.includes("paste"), JSON.stringify(m));
  await clickItem(page, "edit");
  await page.waitForTimeout(400);
  check("„Edytuj dokument” przełącza na Edycję i zamyka menu", await page.evaluate(() => !readOnlyMode) && !(await menuItems(page)));

  // 2. Edycja bez zaznaczenia
  await page.evaluate(() => { const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")); placeCaret(ps[1], 3); });
  b = await paraBox(page, 1);
  await page.mouse.click(b.x, b.y, { button: "right" });
  m = await menuItems(page);
  check("Edycja: Wklej, Link, Komentarz, style, Wstaw tabelę; bez B/I/U i Wytnij bez zaznaczenia", m && ["paste", "link", "comment", "h1", "table", "hrule"].every((id) => m.ids.includes(id)) && (ENGINE === "webkit" || (!m.ids.includes("cut") && m.fmt === 0)), JSON.stringify(m)); // Safari: prawy klik zaznacza słowo (jak system)
  await page.keyboard.press("Escape");
  check("Esc zamyka menu", !(await menuItems(page)));

  // 3. zaznaczenie: B/I/U, pogrubienie z menu, Kopiuj
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1]; const tn = document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode(); const r = document.createRange(); r.setStart(tn, 0); r.setEnd(tn, 8); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
  await page.mouse.click(b.x, b.y, { button: "right" });
  m = await menuItems(page);
  check("zaznaczenie: B/I/U na górze + Wytnij/Kopiuj + ✦ Przepisz", m && m.fmt === 3 && m.ids.includes("cut") && m.ids.includes("copy") && m.ids.includes("find"), JSON.stringify(m));
  if (ENGINE === "chromium") {
    await clickItem(page, "copy");
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    check("Kopiuj z menu: zaznaczony tekst w schowku", clip.length === 8, JSON.stringify(clip));
    await page.mouse.click(b.x, b.y, { button: "right" });
  }
  await page.click(".doc-menu:not(.hidden) .dmf-bold");
  await idle(page);
  const bold = await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1]; const tn = document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode(); return parseInt(getComputedStyle(tn.parentElement).fontWeight, 10); });
  check("B z menu: zaznaczenie pogrubione", bold >= 600, String(bold));

  // 4. Shift+F10: menu przy kursorze, fokus na pierwszej pozycji
  await page.evaluate(() => { const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")); placeCaret(ps[1], 2); });
  await page.keyboard.press("Shift+F10");
  m = await menuItems(page);
  check("Shift+F10: menu otwarte, fokus w menu", m && m.focus, JSON.stringify(m));
  await page.keyboard.press("Escape");

  // 5. linia pozioma: wstaw z menu, usuń z menu
  const base = await fileRules(page);
  await page.evaluate(() => { const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")); placeCaret(ps[1], ps[1].textContent.length); });
  b = await paraBox(page, 1);
  await page.mouse.click(b.x, b.y, { button: "right" });
  await clickItem(page, "hrule");
  await idle(page);
  const li = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).findIndex((p) => getComputedStyle(p).borderBottomStyle !== "none"));
  check("Wstaw linię poziomą z menu", li > 0 && (await fileRules(page)).rules === base.rules + 1, String(li));
  await page.evaluate((i) => { const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")); placeCaret(ps[i], 0); }, li);
  b = await paraBox(page, li);
  await page.mouse.click(b.x, b.y, { button: "right" });
  m = await menuItems(page);
  check("na linii: „Usuń linię poziomą” na górze", m && m.ids[0] === "delRule", JSON.stringify(m?.ids?.slice(0, 3)));
  await clickItem(page, "delRule");
  await idle(page);
  check("„Usuń linię poziomą”: linii nie ma w pliku", (await fileRules(page)).rules === base.rules, JSON.stringify(await fileRules(page)));

  // 6. tabela: wstaw z menu, w komórce „Wiersz poniżej”
  await page.evaluate(() => { const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")); placeCaret(ps[1], ps[1].textContent.length); });
  b = await paraBox(page, 1);
  await page.mouse.click(b.x, b.y, { button: "right" });
  await clickItem(page, "table");
  await idle(page);
  const rows0 = (await fileRules(page)).rows;
  const cell = await page.evaluate(() => { const td = document.querySelector(".docx-preview-host td p"); placeCaret(td, 0); const r = td.getBoundingClientRect(); return { x: r.left + 5, y: r.top + r.height / 2 }; });
  await page.mouse.click(cell.x, cell.y, { button: "right" });
  m = await menuItems(page);
  check("w komórce: wiersze/kolumny/usuń tabelę, bez „Wstaw tabelę”", m && ["rowAbove", "rowBelow", "colLeft", "delTable"].every((id) => m.ids.includes(id)) && !m.ids.includes("table"), JSON.stringify(m));
  await clickItem(page, "rowBelow");
  await idle(page);
  check("„Wiersz poniżej” z menu: tabela ma wiersz więcej", (await fileRules(page)).rows === rows0 + 1, `${rows0} → ${(await fileRules(page)).rows}`);

  // 7. dotyk: „⋯” przy zaznaczeniu otwiera to samo menu
  const ctx2 = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: ENGINE === "chromium" });
  await ctx2.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const ph = await ctx2.newPage();
  await ph.goto(APP_URL, { waitUntil: "load" });
  await ph.evaluate(() => document.getElementById("heroSplash")?.remove());
  await ph.locator("#fileInput").setInputFiles({ name: "menu.docx", mimeType: DOCX_MIME, buffer: fs.readFileSync(SAMPLE) });
  await idle(ph);
  await ph.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await ph.waitForTimeout(300);
  const pb = await paraBox(ph, 1);
  await ph.touchscreen.tap(pb.x, pb.y);
  await ph.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1]; const tn = document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode(); const r = document.createRange(); r.setStart(tn, 0); r.setEnd(tn, 6); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
  await ph.waitForTimeout(300);
  const chip = await ph.evaluate(() => { const c = document.querySelector(".doc-menu-chip:not([hidden])"); if (!c) return null; const r = c.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2, inView: r.right <= innerWidth && r.left >= 0 }; });
  check("dotyk: przy zaznaczeniu pojawia się „⋯” (w obrębie ekranu)", chip && chip.inView, JSON.stringify(chip));
  if (chip) {
    await ph.touchscreen.tap(chip.x, chip.y);
    const mm = await menuItems(ph);
    const fits = await ph.evaluate(() => { const r = document.querySelector(".doc-menu:not(.hidden)")?.getBoundingClientRect(); return r && r.left >= 0 && r.right <= innerWidth && r.bottom <= innerHeight + 1; });
    check("„⋯” otwiera menu z B/I/U, mieści się na ekranie telefonu", mm && mm.fmt === 3 && fits, JSON.stringify({ mm, fits }));
  }
  await ctx2.close();

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
