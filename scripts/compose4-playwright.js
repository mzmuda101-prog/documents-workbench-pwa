// compose4-playwright.js — tworzenie dokumentu, Etap 4: kolor czcionki i wyróżnienie, komentarze,
// nagłówek i stopka z numerem strony.
//
// Jak użytkownik: przycisk „A” (paleta jak w Wordzie: kolor, wyróżnienie, automatyczny), Ctrl/⌘+Alt+M
// (komentarz do zaznaczenia), karta komentarza w komentowanym tekście (odpowiedz / rozwiąż / usuń),
// „Wstaw → Nagłówek, stopka, numer strony…” (+ klik w nagłówek w podglądzie), „inna pierwsza strona”.
// Do tego: pisanie w akapicie z komentarzem z pliku Worda (review-sample) nie gubi komentarza,
// a wyróżnienie przetrwa edycję akapitu. Na końcu ZAPISANY plik + telefon z dotykiem.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));
const select = (page, i, a, z) => page.evaluate(([i, a, z]) => {
  const el = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
  const r = formDomRange(el, a, z); el.focus(); getSelection().removeAllRanges(); getSelection().addRange(r);
}, [i, a, z]);
const paraHtml = (page, i) => page.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].innerHTML, i);
const savedZip = async (page) => {
  const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
  return JSZip.loadAsync(Buffer.from(b64, "base64"));
};
const docXml = async (page) => (await savedZip(page)).file("word/document.xml").async("string");
const sections = (page) => page.evaluate(() => [...document.querySelectorAll(".docx-preview-host section.docx")].map((s) => ({ h: s.querySelector(":scope>header")?.textContent || "", f: s.querySelector(":scope>footer")?.textContent || "" })));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); try { localStorage.removeItem("dwb.authorName"); localStorage.removeItem("dwb.lastFontColor"); } catch (_) {} });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);

  // ── kolor czcionki i wyróżnienie ───────────────────────────────────────────
  await page.keyboard.type("Ala ma kota, a kot ma Alę.");
  await select(page, 0, 4, 6);
  await page.click("#fmtColorBtn");
  const pal = await page.evaluate(() => ({ theme: document.querySelectorAll(".color-grid-theme .color-swatch").length, std: document.querySelectorAll(".compose-pop-color .color-grid:not(.color-grid-theme):not(.color-grid-hl) .color-swatch").length, hl: document.querySelectorAll(".color-grid-hl .color-swatch").length, more: !!document.querySelector(".color-more input[type=color]") }));
  check("paleta jak w Wordzie: 60 kolorów motywu, 10 standardowych, wyróżnienia, „Więcej kolorów”", pal.theme === 60 && pal.std === 10 && pal.hl >= 14 && pal.more, JSON.stringify(pal));
  await page.click('.compose-pop-color .color-swatch[aria-label="Czerwony"] >> nth=0');
  await idle(page);
  let html = await paraHtml(page, 0);
  check("kolor czerwony na zaznaczeniu „ma”, reszta bez zmian", /color: rgb\(255, 0, 0\);?">ma</.test(html) && (html.match(/rgb\(255, 0, 0\)/g) || []).length === 1, html);
  check("zaznaczenie zostaje po zmianie koloru", (await page.evaluate(() => getSelection().toString())) === "ma");
  check("pasek pod „A” pokazuje ostatni kolor", (await page.evaluate(() => getComputedStyle(document.getElementById("fmtColorBar")).backgroundColor)) === "rgb(255, 0, 0)");
  await page.click("#fmtColorBtn");
  await page.click('.color-grid-hl .color-swatch[aria-label="Żółty"]');
  await idle(page);
  html = await paraHtml(page, 0);
  check("wyróżnienie żółte dokłada się do koloru", /color: rgb\(255, 0, 0\); background-color: yellow/.test(html), html);
  let xml = await docXml(page);
  check("plik: w:color FF0000 + w:highlight yellow (kolejność w:rPr wg schematu)", /<w:rPr><w:color w:val="FF0000"\/><w:highlight w:val="yellow"\/><\/w:rPr><w:t>ma<\/w:t>/.test(xml), (xml.match(/<w:rPr>.*?<\/w:rPr>/) || [""])[0]);
  // pisanie w akapicie nie gubi wyróżnienia (model akapitu je zna)
  await page.evaluate(() => focusParagraphAtOffset(0, 26));
  await page.keyboard.type(" Dopisane.");
  await page.waitForTimeout(300);
  xml = await docXml(page);
  check("pisanie w akapicie: wyróżnienie i kolor zostają w pliku", /<w:highlight w:val="yellow"\/>/.test(xml) && /Dopisane\./.test(xml), "");
  await select(page, 0, 4, 6);
  await page.click("#fmtColorBtn");
  await page.click(".color-auto");
  await idle(page);
  await page.click("#fmtColorBtn");
  await page.click(".color-grid-hl .color-swatch.is-none");
  await idle(page);
  xml = await docXml(page);
  check("„Automatyczny” i „Bez wyróżnienia” zdejmują kolor i wyróżnienie", !/FF0000/.test(xml) && !/w:highlight/.test(xml), "");
  // bez zaznaczenia — kolor dla dalszego pisania
  await page.evaluate(() => focusParagraphAtOffset(0, 0));
  await page.click("#fmtColorBtn");
  await page.click('.compose-pop-color .color-swatch[aria-label="Niebieski"] >> nth=0');
  await page.keyboard.type("Nowe ");
  await page.waitForTimeout(250);
  html = await paraHtml(page, 0);
  check("bez zaznaczenia: dalsze pisanie w wybranym kolorze", /color: ?(rgb\(68, 114, 196\)|#4472C4);?">Nowe /i.test(html), html.slice(0, 200));

  // ── komentarze ─────────────────────────────────────────────────────────────
  await select(page, 0, 9, 11); // „ma” w „Nowe Ala ma kota”
  await page.keyboard.press(`${MOD}+Alt+KeyM`);
  check("Ctrl/⌘+Alt+M: okienko komentarza", await page.evaluate(() => !!document.querySelector(".compose-pop-form .cf-text")));
  await page.fill(".cf-text", "Czy na pewno?");
  await page.click(".compose-pop-form .lf-ok");
  await page.waitForTimeout(200);
  check("bez imienia — prośba o imię (podpis komentarza w Wordzie)", await page.evaluate(() => !!document.querySelector(".compose-pop-form")));
  await page.fill(".cf-author", "Jan Test");
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
  const marks = await page.evaluate(() => [...document.querySelectorAll(".docx-preview-host span[data-cm]")].map((s) => s.dataset.cmKind).join(","));
  check("komentarz: znaczniki w podglądzie (początek, koniec, odwołanie)", marks === "start,end,ref", marks);
  check("komentowany tekst podświetlony (CSS Highlight)", await page.evaluate(() => !CSS.highlights || (CSS.highlights.get("dwb-comment")?.size || 0) === 1));
  await page.evaluate(() => focusParagraphAtOffset(0, 10));
  await page.waitForTimeout(700);
  let card = await page.evaluate(() => document.querySelector(".comment-card")?.textContent || "");
  check("kursor w komentowanym tekście → karta z autorem i treścią", /Jan Test/.test(card) && /Czy na pewno\?/.test(card), card);
  await page.keyboard.type("x");
  await page.waitForTimeout(250);
  check("pisanie chowa kartę komentarza", !(await page.$(".comment-card")));
  await page.keyboard.press("Backspace");
  await page.keyboard.press("ArrowLeft"); // ruch kursora — karta wraca
  await page.waitForTimeout(500);
  check("ruch kursora pokazuje kartę z powrotem", !!(await page.$(".comment-card")));
  await page.click(".comment-card .btn:has-text('Odpowiedz')");
  check("odpowiedź: imię zapamiętane", (await page.inputValue(".cf-author")) === "Jan Test");
  await page.fill(".cf-text", "Tak.");
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
  await page.evaluate(() => focusParagraphAtOffset(0, 10));
  await page.waitForTimeout(700);
  card = await page.evaluate(() => document.querySelector(".comment-card")?.textContent || "");
  check("karta pokazuje odpowiedź", /Tak\./.test(card), card);
  await page.click(".comment-card .btn:has-text('Rozwiąż')");
  await idle(page);
  await page.evaluate(() => focusParagraphAtOffset(0, 10));
  await page.waitForTimeout(700);
  check("„Rozwiąż”: karta oznaczona jako rozwiązany", await page.evaluate(() => document.querySelector(".comment-card")?.classList.contains("is-done")));
  let zip = await savedZip(page);
  xml = await zip.file("word/document.xml").async("string");
  const ex = await zip.file("word/commentsExtended.xml")?.async("string");
  check("plik: 2 komentarze (z odpowiedzią), znaczniki obu, commentsExtended z rodzicem i done=1",
    ((await zip.file("word/comments.xml").async("string")).match(/<w:comment /g) || []).length === 2 && (xml.match(/<w:commentReference/g) || []).length === 2
    && /w15:paraIdParent/.test(ex || "") && /w15:done="1"/.test(ex || "") && /relationships\/comments"/.test(await zip.file("word/_rels/document.xml.rels").async("string")), ex);
  // pisanie w akapicie z komentarzem — znaczniki zostają
  await page.evaluate(() => focusParagraphAtOffset(0, 0));
  await page.keyboard.type("Start ");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(400);
  xml = await docXml(page);
  check("pisanie i Enter w akapicie z komentarzem: znaczniki zostają w pliku", (xml.match(/<w:commentRangeStart/g) || []).length === 2 && (xml.match(/<w:commentReference/g) || []).length === 2, "");
  await page.evaluate(() => focusParagraphAtOffset(1, 10));
  await page.waitForTimeout(600);
  await page.click(".comment-card .btn.cc-del");
  await idle(page);
  xml = await docXml(page);
  check("„Usuń” zdejmuje komentarz z odpowiedzią", !/commentReference/.test(xml), "");

  // ── nagłówek i stopka ──────────────────────────────────────────────────────
  await page.evaluate(() => focusParagraphAtOffset(1, 0));
  await page.evaluate(() => composeUi.insertPageBreak());
  await idle(page);
  await page.click("#insertMenuBtn");
  await page.click(".compose-pop-insert .compose-item-label:text-is('Nagłówek, stopka, numer strony…')");
  await page.fill(".hf-h", "Firma XYZ");
  await page.click('[data-seg="h"] button[data-v="right"]');
  await page.fill(".hf-f", "Poufne");
  await page.selectOption(".hf-n", "pageOf");
  await page.click(".compose-pop-hf .lf-ok");
  await idle(page);
  let secs = await sections(page);
  check("nagłówek i stopka na każdej stronie, numer „Strona X z N” w podglądzie", secs.length === 2 && secs.every((s) => s.h === "Firma XYZ") && secs[0].f === "PoufneStrona 1 z 2" && secs[1].f === "PoufneStrona 2 z 2", JSON.stringify(secs));
  check("nagłówek wyrównany do prawej", await page.evaluate(() => getComputedStyle(document.querySelector(".docx-preview-host section.docx > header p")).textAlign) === "right");
  await page.click(".docx-preview-host section.docx > footer >> nth=1");
  await page.waitForTimeout(400);
  const st = await page.evaluate(() => ({ h: document.querySelector(".hf-h")?.value, f: document.querySelector(".hf-f")?.value, n: document.querySelector(".hf-n")?.value }));
  check("klik w stopkę w Edycji otwiera okienko z obecnym stanem", st.h === "Firma XYZ" && st.f === "Poufne" && st.n === "pageOf", JSON.stringify(st));
  await page.click(".hf-first");
  await page.click(".compose-pop-hf .lf-ok");
  await idle(page);
  secs = await sections(page);
  check("„Inna pierwsza strona”: 1. strona bez nagłówka i stopki", secs[0].h === "" && secs[0].f === "" && secs[1].h === "Firma XYZ", JSON.stringify(secs));
  zip = await savedZip(page);
  xml = await zip.file("word/document.xml").async("string");
  const parts = Object.keys(zip.files).filter((f) => /word\/(header|footer)\d+\.xml/.test(f));
  const sect = (xml.match(/<w:sectPr>.*?<\/w:sectPr>/) || [""])[0];
  check("plik: odwołania nagłówka/stopki na początku sectPr, titlePg, 4 części (bez sierot)", /^<w:sectPr><w:headerReference w:type="default"/.test(sect) && /<w:titlePg\/>/.test(sect) && parts.length === 4, JSON.stringify({ sect: sect.slice(0, 160), parts }));
  const footer = await zip.file(parts.find((f) => f.includes("footer") && !/<w:p\/>/.test(""))).async("string");
  check("plik: numer strony to pola PAGE i NUMPAGES", /instrText[^>]*> PAGE </.test(footer + (await Promise.all(parts.map((f) => zip.file(f).async("string")))).join("")) && /NUMPAGES/.test((await Promise.all(parts.map((f) => zip.file(f).async("string")))).join("")), "");
  // usunięcie: puste pola, bez numeru, bez pierwszej
  await page.click("#insertMenuBtn");
  await page.click(".compose-pop-insert .compose-item-label:text-is('Nagłówek, stopka, numer strony…')");
  await page.fill(".hf-h", ""); await page.fill(".hf-f", ""); await page.selectOption(".hf-n", ""); await page.click(".hf-first");
  await page.click(".compose-pop-hf .lf-ok");
  await idle(page);
  zip = await savedZip(page);
  check("puste nagłówek/stopka: odwołania i części usunięte z pliku", !/headerReference|footerReference/.test(await zip.file("word/document.xml").async("string")) && !Object.keys(zip.files).some((f) => /word\/(header|footer)\d+\.xml/.test(f)), Object.keys(zip.files).join(","));

  // ── komentarz z pliku Worda: edycja akapitu go nie gubi ────────────────────
  await page.evaluate(() => loadSampleDocument("review-sample"));
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);
  const kara = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).findIndex((p) => p.textContent.startsWith("Kara umowna")));
  const locked = await page.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].dataset.lock || "", kara);
  await page.evaluate((i) => focusParagraphAtOffset(i, 0), kara);
  await page.keyboard.type("UWAGA: ");
  await page.waitForTimeout(300);
  xml = await docXml(page);
  const para = (xml.match(/<w:p[ >](?:(?!<\/w:p>).)*UWAGA: (?:(?!<\/w:p>).)*<\/w:p>/) || [""])[0];
  check("plik z Worda: akapit z komentarzem edytowalny, po pisaniu znaczniki komentarza zostają", !locked && /commentRangeStart/.test(para) && /commentRangeEnd/.test(para) && /commentReference/.test(para), locked || para.slice(0, 300));

  // ── telefon z dotykiem ─────────────────────────────────────────────────────
  const touch = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 800 }, isMobile: true, hasTouch: true });
  await touch.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const tp = await touch.newPage();
  tp.on("pageerror", (e) => errors.push(`touch: ${e.message}`));
  await tp.goto(APP_URL, { waitUntil: "load" });
  await tp.evaluate(() => document.getElementById("heroSplash")?.remove());
  await tp.evaluate(() => composeUi.createNew("note"));
  await tp.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(tp);
  await select(tp, 3, 0, 6);
  await tp.evaluate(() => document.getElementById("formatToolbar").scrollLeft = 9999);
  await tp.tap("#fmtColorBtn");
  const cp = await tp.evaluate(() => { const r = document.querySelector(".compose-pop-color")?.getBoundingClientRect(); return r && { l: r.left, r: r.right, w: innerWidth }; });
  check("telefon: paleta kolorów mieści się na ekranie", cp && cp.l >= 0 && cp.r <= cp.w, JSON.stringify(cp));
  await tp.tap('.compose-pop-color .color-swatch[aria-label="Zielony"] >> nth=1');
  await idle(tp);
  check("telefon: tap w kolor koloruje zaznaczenie", /rgb\(0, 176, 80\)/.test(await paraHtml(tp, 3)), (await paraHtml(tp, 3)).slice(0, 200));
  await tp.evaluate(() => document.getElementById("formatToolbar").scrollLeft = 0);
  await tp.tap("#insertMenuBtn");
  await tp.tap(".compose-pop-insert .compose-item-label:text-is('Nagłówek, stopka, numer strony…')");
  await tp.waitForSelector(".compose-pop-hf", { timeout: 5000 }); // okienko czyta najpierw stan z pliku
  const hfBox = await tp.evaluate(() => { const r = document.querySelector(".compose-pop-hf")?.getBoundingClientRect(); return r && { l: r.left, r: r.right, w: innerWidth }; });
  check("telefon: okienko nagłówka i stopki mieści się na ekranie", hfBox && hfBox.l >= 0 && hfBox.r <= hfBox.w, JSON.stringify(hfBox));
  await tp.selectOption(".hf-n", "n");
  await tp.tap(".compose-pop-hf .lf-ok");
  await idle(tp);
  const z2 = await savedZip(tp);
  check("telefon: numer strony zapisany", Object.keys(z2.files).some((f) => /footer\d+\.xml/.test(f)));
  await touch.close();

  check("brak błędów strony", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || r.detail == null ? "" : ` — ${String(r.detail).slice(0, 400)}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}

run().catch((e) => {
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || r.detail == null ? "" : ` — ${String(r.detail).slice(0, 400)}`}`));
  console.error(e);
  process.exit(1);
});
