// linebreak-playwright.js — Shift+Enter i znaczniki za łamaniem wiersza (2026-10-06, z testu chaos).
//
// 1) Shift+Enter na końcu akapitu i tuż przed znacznikiem zakładki = JEDNO <w:br/> w pliku.
//    Dawniej przeglądarka (pre-wrap) wstawiała dwa „\n” — w Wordzie pusty wiersz za dużo.
//    Kursor ma stać w NOWYM wierszu (widocznym), a pisanie dalej — bez dodatkowych łamań.
// 2) Zakładka za łamaniem wiersza zostaje na miejscu po zapisie, ponownym otwarciu i kolejnym
//    zapisie (dawniej przesuwała się o liczbę łamań przed nią: podgląd liczył tylko tekst).
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 })
  .then(() => page.evaluate(() => waitInlineStructuralIdle())).then(() => page.waitForTimeout(300));
const savedB64 = (page) => page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
const xmlOf = async (b64) => (await (await JSZip.loadAsync(Buffer.from(b64, "base64"))).file("word/document.xml").async("string")).replace(/ w14:paraId="[^"]*"| w14:textId="[^"]*"/g, "");
const firstPara = (xml) => (xml.match(/<w:body><w:p[ >](?:(?!<\/w:p>).)*<\/w:p>/) || [""])[0];

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);

  // ── 1) Shift+Enter na końcu ──
  await page.keyboard.type("Pierwszy wiersz");
  const y0 = await page.evaluate(() => getSelection().getRangeAt(0).getClientRects()[0]?.top ?? 0);
  await page.keyboard.press("Shift+Enter");
  // gdzie stoi kursor: chwilowy znak zerowej szerokości w miejscu kursora (prostokąt pustego
  // zakresu tuż za „\n” przeglądarki podają jeszcze na poprzednim wierszu)
  const caret = await page.evaluate(() => {
    const sel = getSelection(); const r = sel.getRangeAt(0).cloneRange();
    const m = document.createElement("span"); m.textContent = "\u200b"; r.insertNode(m);
    const top = m.getBoundingClientRect().top; const parent = m.parentNode; m.remove(); parent.normalize();
    return top;
  });
  await idle(page);
  let xml = await xmlOf(await savedB64(page));
  check("Shift+Enter na końcu akapitu = jedno łamanie w pliku", (firstPara(xml).match(/<w:br\/>/g) || []).length === 1, firstPara(xml));
  check("…kursor w nowym wierszu (niżej niż tekst)", caret > y0 + 5, `${y0} → ${caret}`);
  await page.keyboard.type("Drugi wiersz");
  await idle(page);
  xml = await xmlOf(await savedB64(page));
  check("pisanie po Shift+Enter: dalej jedno łamanie, tekst w drugim wierszu", (firstPara(xml).match(/<w:br\/>/g) || []).length === 1 && /<w:br\/><\/w:r><w:r><w:t[^>]*>Drugi wiersz/.test(firstPara(xml)), firstPara(xml));
  await page.keyboard.press("Shift+Enter");
  await page.keyboard.press("Shift+Enter");
  await idle(page);
  xml = await xmlOf(await savedB64(page));
  check("dwa Shift+Enter na końcu = dwa łamania (nie cztery)", (firstPara(xml).match(/<w:br\/>/g) || []).length === 3, firstPara(xml));

  // ── 2) zakładka za łamaniem, potem Shift+Enter tuż przed nią ──
  await page.keyboard.press("Enter");
  await page.keyboard.type("Wprowadzenie");
  await idle(page);
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1]; placeCaret(p, 2); });
  await page.evaluate(() => composeUi.openBookmarkForm(document.getElementById("insertMenuBtn")));
  await page.waitForSelector(".bm-name");
  await page.fill(".bm-name", "Znacznik");
  await page.click(".bm-add");
  await idle(page);
  await page.keyboard.press("Shift+Enter");
  await idle(page);
  const b1 = await savedB64(page);
  xml = await xmlOf(b1);
  const p2 = (xml.match(/<w:p>(?:(?!<w:p>).)*Znacznik(?:(?!<\/w:p>).)*<\/w:p>/) || [""])[0].replace(/<w:rPr>.*?<\/w:rPr>/g, "");
  check("Shift+Enter tuż przed zakładką = jedno łamanie, zakładka za nim", /<w:t>Wp<\/w:t><\/w:r><w:r><w:br\/><\/w:r><w:bookmarkStart [^>]*w:name="Znacznik"\/>/.test(p2) && (p2.match(/<w:br\/>/g) || []).length === 1, p2);
  // zapis → otwarcie → zapis: zakładka w tym samym miejscu (dawniej przesuwała się o łamania)
  await page.locator("#fileInput").setInputFiles({ name: "lb.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: Buffer.from(b1, "base64") });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  await page.evaluate(() => { if (readOnlyMode) appFrame.setReadOnly(false); });
  await idle(page);
  // pisanie w tym akapicie (zapis przepisuje akapit ze znacznikami w miejscach z podglądu)
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1]; placeCaret(p, p.textContent.length); });
  await page.keyboard.type("!");
  await idle(page);
  const x2 = await xmlOf(await savedB64(page));
  const q2 = (x2.match(/<w:p>(?:(?!<w:p>).)*Znacznik(?:(?!<\/w:p>).)*<\/w:p>/) || [""])[0].replace(/<w:rPr>.*?<\/w:rPr>/g, "");
  check("po ponownym otwarciu i pisaniu: zakładka dalej tuż za łamaniem", /<w:br\/><\/w:r><w:bookmarkStart [^>]*w:name="Znacznik"\/><w:bookmarkEnd [^>]*\/><w:r><w:t[^>]*>rowadzenie!<\/w:t>/.test(q2), q2);

  check("brak błędów strony", errors.length === 0, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  let failed = 0;
  results.forEach((r) => { if (!r.ok) failed++; console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` — ${r.detail}`}`); });
  console.log(`\n${ENGINE}: ${results.length - failed}/${results.length} OK`);
  process.exit(failed ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
