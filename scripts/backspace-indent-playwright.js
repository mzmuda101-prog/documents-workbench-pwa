// backspace-indent-playwright.js — Backspace na początku punktu listy i wciętego akapitu jak w Wordzie
// (sprawdzone w Word 16.114, 2026-10-17):
//   1. Backspace w (pustym) punkcie — numer/punktor znika, akapit ZOSTAJE wcięty tam, gdzie zaczynał
//      się tekst punktu,
//   2. Backspace — wcięcie do zera (kursor tam, gdzie pisze się bez listy), akapit nadal osobny,
//   3. Backspace — dopiero teraz sklejenie z poprzednim akapitem.
// Zwykły akapit z lewym wcięciem (też podwójnym): Backspace na początku → od razu margines.
// Enter w pustym punkcie — od razu margines (bez kroku z wcięciem). Plik: w:ind po każdym kroku.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(300));
// stan akapitów: tekst, lista, lewa krawędź tekstu względem pierwszego akapitu (px)
const state = (page) => page.evaluate(() => {
  const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
  const left0 = ps[0].getBoundingClientRect().left + (parseFloat(getComputedStyle(ps[0]).paddingLeft) || 0);
  const caretP = docCaretParagraph(document.activeElement);
  return ps.map((p) => {
    const cs = getComputedStyle(p);
    const textLeft = p.getBoundingClientRect().left + (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.textIndent) || 0);
    return { text: p.textContent.replace(/﻿/g, ""), list: isListParagraph(p), indent: Math.round(textLeft - left0), caret: p === caretP };
  });
});
const savedInd = async (page) => {
  const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
  const zip = await JSZip.loadAsync(Buffer.from(b64, "base64"));
  const xml = await zip.file("word/document.xml").async("string");
  return (xml.match(/<w:p[ >].*?<\/w:p>/gs) || []).map((p) => ({ text: (p.match(/<w:t[^>]*>([^<]*)/g) || []).map((t) => t.replace(/<[^>]+>/, "")).join(""), numPr: /<w:numPr>/.test(p), ind: p.match(/<w:ind [^>]*\/>/)?.[0] || "" }));
};

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);

  await page.keyboard.type("Zwykly akapit");
  await page.keyboard.press("Enter");
  await page.keyboard.type("pierwszy");
  await idle(page);
  await page.evaluate(() => composeUi.applyList("bullet"));
  await idle(page);
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.type("drugi");
  await page.keyboard.press("Enter"); // pusty punkt
  await idle(page);
  let st = await state(page);
  const listIndent = st[1].indent;
  check("start: pusty punkt listy", st[3]?.list && st[3].text === "" && st[3].caret && listIndent > 10, JSON.stringify(st));

  // 1. Backspace — punktor znika, wcięcie tekstu punktu zostaje
  await page.keyboard.press("Backspace");
  await idle(page);
  st = await state(page);
  check("1. Backspace: nie lista, wcięty jak tekst punktu, kursor w nim", st.length === 4 && !st[3].list && Math.abs(st[3].indent - listIndent) <= 2 && st[3].caret, JSON.stringify(st));
  let ind = await savedInd(page);
  check("1. Backspace w pliku: bez w:numPr, w:ind left > 0", !ind[3].numPr && /w:left="[1-9]/.test(ind[3].ind), JSON.stringify(ind[3]));

  // 2. Backspace — wcięcie do zera, akapit nadal osobny
  await page.keyboard.press("Backspace");
  await idle(page);
  st = await state(page);
  check("2. Backspace: margines, akapit osobny, kursor w nim", st.length === 4 && !st[3].list && Math.abs(st[3].indent) <= 2 && st[3].caret, JSON.stringify(st));
  ind = await savedInd(page);
  check("2. Backspace w pliku: w:ind left=0", /w:left="0"/.test(ind[3].ind), JSON.stringify(ind[3]));

  // 3. Backspace — sklejenie z „drugi”
  await page.keyboard.press("Backspace");
  await idle(page);
  st = await state(page);
  check("3. Backspace: sklejony z poprzednim punktem", st.length === 3 && st[2].text === "Drugi" || st[2]?.text === "drugi", JSON.stringify(st));

  // pisanie zaraz po Backspace (w trakcie przerysowania) nie ginie
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.press("Backspace");
  await page.keyboard.type("tekst", { delay: 10 });
  await idle(page);
  st = await state(page);
  const last = st[st.length - 1];
  check("pisanie od razu po Backspace nie gubi liter", last.text === "tekst" && !last.list && Math.abs(last.indent - listIndent) <= 2, JSON.stringify(last));

  // zwykły akapit z podwójnym wcięciem — Backspace na początku → margines
  await page.keyboard.press("Enter");
  await page.keyboard.type("wciety");
  await idle(page);
  const wi = (await state(page)).length - 1;
  await page.evaluate((i) => applyDocumentEdit({ op: "paraFormat", indices: [i], ind: { left: 1440 } }), wi);
  await idle(page);
  await page.evaluate((i) => placeCaret(collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i], 0), wi);
  st = await state(page);
  check("akapit wcięty o 2 kroki", st[wi].indent > 60, JSON.stringify(st[wi]));
  await page.keyboard.press("Backspace");
  await idle(page);
  st = await state(page);
  check("Backspace na początku wciętego akapitu → margines, bez sklejania", st.length === wi + 1 && Math.abs(st[wi].indent) <= 2 && st[wi].text === "wciety", JSON.stringify(st.slice(-2)));
  await page.keyboard.press("Backspace");
  await idle(page);
  st = await state(page);
  check("kolejny Backspace skleja", st.length === wi && st[wi - 1].text.endsWith("wciety"), JSON.stringify(st.slice(-2)));

  // Enter w pustym punkcie — od razu margines
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.type("punkt");
  await idle(page);
  await page.evaluate(() => composeUi.applyList("bullet"));
  await idle(page);
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.press("Enter");
  await idle(page);
  st = await state(page);
  const e = st[st.length - 1];
  check("Enter w pustym punkcie → margines (bez wcięcia)", !e.list && Math.abs(e.indent) <= 2 && e.caret, JSON.stringify(st.slice(-2)));

  check("brak błędów w konsoli", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  if (failed.length) process.exit(1);
}

run().catch((err) => { console.error(err); process.exit(1); });
