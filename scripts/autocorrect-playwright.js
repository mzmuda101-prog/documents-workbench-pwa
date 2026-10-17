// autocorrect-playwright.js — Autokorekta i Autoformatowanie podczas pisania (autocorrect.js, 2026-10-17).
//
// Reguły jak zaobserwowane w Microsoft Word 16.114 (język polski): poprawka słowa po spacji,
// wielka litera na początku zdania (nie po „np.”), „PIerwszy” → „Pierwszy”, cudzysłowy „ ” i ‘ ’,
// „ - ” / „ -- ” → półpauza, „a--b” → pauza, „1/2” → „½”, „...” → „…”, „(c)” → „©”,
// „1. ” / „* ” na początku akapitu → lista. Ctrl/⌘+Z zaraz po poprawce cofa SAMĄ poprawkę,
// a po liście — wraca do wpisanego „1. ”. Pisanie w tekście z formatem (Ctrl/⌘+B) też poprawia.
// Wyłącznik działa. Zapisany plik ma poprawiony tekst. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(300));
const paras = (page) => page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => ({ text: p.textContent.replace(/﻿/g, ""), list: isListParagraph(p) })));
const lastPara = async (page) => (await paras(page)).filter((p) => p.text || p.list).pop() || { text: "" };
const caretPara = (page) => page.evaluate(() => {
  const p = docCaretParagraph(document.activeElement);
  return p ? { text: p.textContent.replace(/﻿/g, ""), list: isListParagraph(p) } : null;
});

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => {
    sessionStorage.setItem("introPlayed", "true");
    localStorage.setItem("dwb-autocorrect-v1", JSON.stringify({ on: true }));
  });
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

  // 1) słowa, wielka litera, cudzysłowy, półpauza
  await page.keyboard.type(`ktory jest "cytat" i 'apostrof' -- test `, { delay: 5 });
  let t = (await caretPara(page))?.text;
  check("słowo + wielka litera + cudzysłowy + półpauza", t === "Który jest „cytat” i ‘apostrof’ – test ", t);

  // 2) DWie WIelkie, ułamek, skrót bez wielkiej litery, nowe zdanie, pauza, ogonki
  await page.keyboard.press("Enter");
  await page.keyboard.type("PIerwszy 1/2 np. tak. to raz - dwa a--b (c) i tak... juz ", { delay: 5 });
  t = (await caretPara(page))?.text;
  check("PIerwszy, ½, np. tak, To, –, —, ©, …, już", t === "Pierwszy ½ np. tak. To raz – dwa a—b © i tak… już ", t);

  // 3) Cofnij zaraz po poprawce — wraca samo słowo
  await page.keyboard.press(`${MOD}+z`);
  await idle(page);
  t = (await lastPara(page)).text;
  check("Ctrl/⌘+Z cofa samą autokorektę („juz” wraca)", t.endsWith("i tak… juz "), t);
  const label = await page.evaluate(() => dwbUndo._debug().redo.slice(-1)[0]);
  check("krok cofania nazywa się „autokorekta”", label === "undoOpAutoCorrect", label);

  // 4) środek istniejącego słowa — bez poprawek
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Zostaw ", { delay: 5 });
  await page.keyboard.type("ktory", { delay: 5 });
  await page.keyboard.press("ArrowLeft");
  await page.keyboard.press("ArrowLeft");
  await page.keyboard.type(" ", { delay: 5 });
  t = (await caretPara(page))?.text;
  check("spacja w środku słowa nie poprawia", t === "Zostaw kto ry", t);

  // 5) lista numerowana z „1. ” — i Cofnij wraca do „1. ”
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  // pisane w biegu, bez czekania na przerysowanie (dawniej litery w tym czasie ginęły)
  await page.keyboard.type("1. numer jeden ", { delay: 15 });
  await idle(page);
  let cp = await caretPara(page);
  check("„1. ” w biegu → lista numerowana, nic nie zginęło", cp?.list && cp.text === "Numer jeden ", JSON.stringify(cp));
  await page.keyboard.press("Enter");
  await page.keyboard.press("Enter"); // pusty punkt kończy listę
  await idle(page);
  await page.keyboard.type("1. ", { delay: 5 });
  await idle(page);
  await page.keyboard.press(`${MOD}+z`);
  await idle(page);
  cp = await caretPara(page);
  const all = await paras(page);
  const back = all.find((p) => p.text === "1. ");
  check("Cofnij po liście wraca do wpisanego „1. ”", back && !back.list, JSON.stringify(all.slice(-3)));

  // 6) lista punktowana z „* ”
  await page.evaluate(() => { const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")); const p = ps.find((x) => x.textContent === "1. "); placeCaret(p, 3); });
  await page.keyboard.press("Backspace");
  await page.keyboard.press("Backspace");
  await page.keyboard.press("Backspace");
  await page.keyboard.type("* punkt", { delay: 15 });
  await page.keyboard.press("Enter"); // Enter też w biegu — trzymany i wykonany po przerysowaniu
  await idle(page);
  const ps = await paras(page);
  const bi = ps.findIndex((x) => x.text === "Punkt");
  check("„* ” → lista punktowana, Enter w biegu dał nowy punkt", bi >= 0 && ps[bi].list && ps[bi + 1]?.list && ps[bi + 1].text === "", JSON.stringify(ps.slice(-3)));

  // 7) pisanie z formatem (Ctrl/⌘+B) — znak wstawia aplikacja, autokorekta i tak działa
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.type("bede gruby ", { delay: 5 });
  t = (await caretPara(page))?.text;
  const bold = await page.evaluate(() => { const p = docCaretParagraph(document.activeElement); return !!p?.querySelector("b, strong, span[style*='bold'], span[style*='700']"); });
  check("pisanie pogrubione też poprawiane („Będę”)", t === "Będę gruby " && bold, `${t} bold=${bold}`);
  await page.keyboard.press(`${MOD}+b`);

  // 7b) Enter w pustym punkcie kończy listę (przerysowanie) — litery wpisane zaraz potem nie giną
  await page.keyboard.press("Enter");
  await page.keyboard.press("Enter");
  await page.keyboard.type("dalej zwykly tekst ", { delay: 10 });
  await idle(page);
  cp = await caretPara(page);
  check("po końcu listy szybkie pisanie nie gubi liter", cp && !cp.list && cp.text === "Dalej zwykly tekst ", JSON.stringify(cp));

  // 8) wyłącznik
  await page.keyboard.press("Enter");
  await page.evaluate(() => dwbAutoCorrect.set("on", false));
  await page.keyboard.type(`juz "x" `, { delay: 5 });
  t = (await caretPara(page))?.text;
  check("wyłączona autokorekta nic nie zmienia", t === `juz "x" `, t);
  const boxDisabled = await page.evaluate(() => document.querySelector('[data-autocorrect="caps"]')?.disabled);
  check("wyłącznik główny wygasza opcje w panelu", boxDisabled === true, String(boxDisabled));
  await page.evaluate(() => dwbAutoCorrect.set("on", true));

  // 9) zapis
  await idle(page);
  const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
  const zip = await JSZip.loadAsync(Buffer.from(b64, "base64"));
  const xml = await zip.file("word/document.xml").async("string");
  if (process.env.DUMP) console.log(xml.match(/<w:p[ >].*?<\/w:p>/gs).map((p) => (p.match(/numPr.{0,120}/s)?.[0] || "-") + " :: " + (p.match(/<w:t[^>]*>[^<]*/g) || []).join("")).join("\n"));
  check("zapis: poprawiony tekst w pliku", xml.includes("Który jest „cytat” i ‘apostrof’ – test") && xml.includes("Pierwszy ½ np. tak. To raz – dwa a—b ©") && xml.includes("Numer jeden"), "");
  check("zapis: listy w pliku (w:numPr)", (xml.match(/<w:numPr>/g) || []).length >= 3, String((xml.match(/<w:numPr>/g) || []).length));

  check("brak błędów w konsoli", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();

  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  if (failed.length) process.exit(1);
}

run().catch((err) => { console.error(err); process.exit(1); });
