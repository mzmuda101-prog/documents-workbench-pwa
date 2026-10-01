// compose-playwright.js — tworzenie dokumentu od zera (Etap 1: docx-compose.js + compose-ui.js).
//
// Jak użytkownik: „Nowy dokument” → szablon → pisanie; styl akapitu z listy, Enter po nagłówku
// (dalej zwykły tekst), wyrównanie, „＋ Wstaw” (podział strony, linia, data, znak), Ctrl+Enter,
// Cofnij. Na końcu zapisany plik: kolejność w:pPr zgodna ze schematem, style dopisane do
// styles.xml (także w pliku BEZ styles.xml), sekcja nie zdublowana. Plus pasek na telefonie:
// jeden rząd, przewija się w bok, okienko „Wstaw” mieści się na ekranie.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(350));

const paras = (page) => page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => ({
  text: p.textContent, style: composeStyleKeyOf(p), align: getComputedStyle(p).textAlign, editable: p.classList.contains("docx-editable-p"),
})));
const fileXml = async (page, part = "word/document.xml") => {
  const b64 = await page.evaluate(async () => {
    const bytes = await buildDocumentForSave();
    let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s);
  });
  const zip = await JSZip.loadAsync(Buffer.from(b64, "base64"));
  return zip.file(part) ? zip.file(part).async("string") : null;
};
const caretTo = (page, idx, atEnd = true) => page.evaluate(([i, end]) => {
  const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
  p.focus();
  const r = document.createRange(); r.selectNodeContents(p); r.collapse(!end);
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
}, [idx, atEnd]);
const caretAt = (page, idx, offset) => page.evaluate(([i, o]) => focusParagraphAtOffset(i, o), [idx, offset]);

// Kolejność dzieci w:pPr wg schematu (fragment istotny dla nas).
const ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "numPr", "pBdr", "spacing", "ind", "jc", "outlineLvl", "rPr", "sectPr"];
function pPrOrderOk(xml) {
  const bad = [];
  for (const m of xml.matchAll(/<w:pPr>(.*?)<\/w:pPr>/g)) {
    const names = [...m[1].matchAll(/<w:(\w+)[\s/>]/g)].map((x) => x[1]).filter((n) => ORDER.includes(n));
    // tylko bezpośrednie dzieci: pomijamy zagnieżdżone w rPr/sectPr/pBdr
    const top = [];
    let depth = 0;
    for (const t of m[1].matchAll(/<(\/?)w:(\w+)([^>]*?)(\/?)>/g)) {
      if (t[1]) { depth--; continue; }
      if (depth === 0) top.push(t[2]);
      if (!t[4]) depth++;
    }
    const ranks = top.filter((n) => ORDER.includes(n)).map((n) => ORDER.indexOf(n));
    if (ranks.some((r, i) => i && r < ranks[i - 1])) bad.push(top.join(","));
    void names;
  }
  return bad;
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => { errors.push(`dialog: ${d.message()}`); d.dismiss(); });
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForTimeout(400);

  // ── Nowy dokument: pusty ────────────────────────────────────────────────────
  await page.click("#emptyNewBtn");
  check("„Nowy dokument” na ekranie startowym otwiera okno szablonów", await page.evaluate(() => document.getElementById("newDocDialog").open));
  await page.click('.newdoc-card[data-template="blank"]');
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);
  const st0 = await page.evaluate(() => ({ name: currentFileName, ro: readOnlyMode, focus: !!document.activeElement?.closest?.(".docx-editable-p"), handle: !!fileHandle }));
  check("pusty dokument: nazwa „Nowy dokument.docx”, tryb Edycja, kursor w tekście, bez uchwytu pliku", st0.name === "Nowy dokument.docx" && !st0.ro && st0.focus && !st0.handle, JSON.stringify(st0));
  const h0 = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0].getBoundingClientRect().height);
  check("pusty akapit ma wysokość (da się w niego kliknąć)", h0 >= 12, h0);

  // pisanie → styl Tytuł z listy → Enter → zwykły tekst
  await page.keyboard.type("Mój raport");
  await page.selectOption("#fmtParaStyle", "title");
  await idle(page);
  let ps = await paras(page);
  check("lista „Styl” → Tytuł: akapit ma styl Tytuł, tekst zostaje", ps[0].style === "title" && ps[0].text === "Mój raport", JSON.stringify(ps[0]));
  const focusAfterStyle = await page.evaluate(() => { const p = document.activeElement?.closest?.(".docx-editable-p"); return p ? resolveParaIndex(p) : -1; });
  check("po zmianie stylu kursor wraca do tego akapitu", focusAfterStyle === 0, focusAfterStyle);
  await caretTo(page, 0, true);
  await page.keyboard.press("Enter");
  await page.keyboard.type("Wstęp do raportu.");
  await page.waitForTimeout(400);
  ps = await paras(page);
  check("Enter na końcu tytułu → nowy akapit zwykłym tekstem (podgląd)", ps.length === 2 && ps[1].style === "normal" && ps[1].text === "Wstęp do raportu.", JSON.stringify(ps));
  let xml = await fileXml(page);
  check("… i w pliku: drugi akapit bez w:pStyle", /<w:p>(?:<w:pPr>(?:(?!pStyle).)*?<\/w:pPr>)?<w:r>.*?Wstęp do raportu/.test(xml.replace(/<w:p [^>]*>/g, "<w:p>")), xml.slice(0, 600));

  // Nagłówek 1 i 2
  await page.keyboard.press("Enter");
  await page.keyboard.type("Rozdział pierwszy");
  await page.selectOption("#fmtParaStyle", "h1");
  await idle(page);
  await caretTo(page, 2, true);
  await page.keyboard.press("Enter");
  await page.keyboard.type("Treść rozdziału, która jest dość długa, żeby wyjustowanie miało znaczenie w podglądzie dokumentu.");
  await page.waitForTimeout(300);
  ps = await paras(page);
  check("Nagłówek 1 z listy + Enter → zwykły tekst", ps[2].style === "h1" && ps[3].style === "normal", JSON.stringify(ps.map((p) => p.style)));
  const chips = await page.evaluate(() => Array.from(document.querySelectorAll("#sectionChips .section-chip")).map((c) => c.textContent.trim()));
  check("nagłówek trafia do skrótów sekcji", chips.some((c) => /Rozdział pierwszy/.test(c)), JSON.stringify(chips));
  // stan listy podąża za kursorem
  await caretTo(page, 2, true);
  await page.waitForTimeout(150);
  const selVal = await page.evaluate(() => document.getElementById("fmtParaStyle").value);
  check("lista „Styl” pokazuje styl akapitu z kursorem", selVal === "h1", selVal);

  // ── wyrównanie ──────────────────────────────────────────────────────────────
  await caretTo(page, 3, true);
  await page.click("#fmtAlignBtn");
  const alignPop = await page.evaluate(() => { const el = document.querySelector(".compose-pop-align"); if (!el) return null; const r = el.getBoundingClientRect(); return { n: el.querySelectorAll("button").length, top: r.top, left: r.left, right: r.right, w: innerWidth }; });
  check("wyrównanie: okienko z 4 przyciskami pod przyciskiem", alignPop && alignPop.n === 4, JSON.stringify(alignPop));
  await page.click('.compose-pop-align button[aria-label="Wyjustuj"]');
  await idle(page);
  ps = await paras(page);
  check("wyjustuj: akapit z kursorem wyjustowany, okienko zamknięte", ps[3].align === "justify" && !(await page.$(".compose-pop")), ps[3].align);
  const alignIcon = await page.evaluate(() => document.querySelector("#fmtAlignBtn svg line:nth-child(2)")?.getAttribute("x2"));
  check("ikona wyrównania na pasku pokazuje stan (wyjustowany)", alignIcon === "20", alignIcon);
  // zaznaczenie przez dwa akapity → wyśrodkuj oba
  await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    ps[0].focus();
    const r = document.createRange(); r.setStart(ps[0].firstChild, 0); r.setEnd(ps[1].lastChild, 1);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  });
  await page.click("#fmtAlignBtn");
  await page.click('.compose-pop-align button[aria-label="Do środka"]');
  await idle(page);
  ps = await paras(page);
  check("zaznaczenie przez 2 akapity → oba wyśrodkowane, reszta bez zmian", ps[0].align === "center" && ps[1].align === "center" && ps[3].align === "justify", JSON.stringify(ps.map((p) => p.align)));

  // ── Wstaw: data, znak, linia, podział strony ────────────────────────────────
  await caretTo(page, 1, true);
  await page.click("#insertMenuBtn");
  const menu = await page.evaluate(() => { const el = document.querySelector(".compose-pop-insert"); return el && { items: Array.from(el.querySelectorAll(".compose-item-label")).map((x) => x.textContent), chars: el.querySelectorAll(".compose-char").length }; });
  check("menu „Wstaw”: podział strony, linia, link, data, spis treści, 4 pola formularza + znaki", menu && ["Podział strony", "Linia pozioma", "Link…", "Dzisiejsza data", "Spis treści", "Tekst", "Lista wyboru", "Data", "Pole wyboru"].every((x) => menu.items.includes(x)) && menu.chars >= 16, JSON.stringify(menu));
  await page.click('.compose-char[aria-label="§"]');
  await page.waitForTimeout(200);
  ps = await paras(page);
  check("znak specjalny trafia w miejsce kursora", ps[1].text === "Wstęp do raportu.§", ps[1].text);
  await page.click("#insertMenuBtn");
  await page.click(".compose-item-label:text-is('Dzisiejsza data')");
  await page.waitForTimeout(200);
  ps = await paras(page);
  const today = await page.evaluate(() => new Date().toLocaleDateString("pl-PL", { day: "numeric", month: "long", year: "numeric" }));
  check("dzisiejsza data wstawiona", ps[1].text.endsWith(today), ps[1].text);
  const undoLabel = await page.evaluate(() => document.getElementById("undoBtn").getAttribute("aria-label"));
  check("wstawienie to osobny krok cofania", /wstawienie/.test(undoLabel), undoLabel);
  await page.click("#undoBtn");
  await idle(page);
  ps = await paras(page);
  check("Cofnij zabiera tylko datę", ps[1].text === "Wstęp do raportu.§", ps[1].text);

  // linia pozioma za akapitem 1
  await caretTo(page, 1, true);
  await page.click("#insertMenuBtn");
  await page.click(".compose-item-label:text-is('Linia pozioma')");
  await idle(page);
  ps = await paras(page);
  const hr = await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[2]; const cs = getComputedStyle(p); return { style: cs.borderBottomStyle, w: cs.borderBottomWidth }; });
  check("linia pozioma: akapit z dolną krawędzią + pusty akapit do pisania", ps.length === 6 && hr.style === "solid" && ps[2].text === "" && ps[3].text === "", JSON.stringify({ n: ps.length, hr }));
  const focusHr = await page.evaluate(() => { const p = document.activeElement?.closest?.(".docx-editable-p"); return p ? resolveParaIndex(p) : -1; });
  check("po linii kursor stoi w akapicie pod nią", focusHr === 3, focusHr);
  await page.keyboard.type("Pod linią");
  await page.waitForTimeout(200);

  // podział strony w środku akapitu (Ctrl+Enter)
  const pagesBefore = await page.evaluate(() => document.querySelectorAll(".docx-preview-host section.docx").length);
  await caretAt(page, 5, 6); // „Treść |rozdziału…”
  await page.keyboard.press(`${MOD}+Enter`);
  await idle(page);
  ps = await paras(page);
  const pagesAfter = await page.evaluate(() => document.querySelectorAll(".docx-preview-host section.docx").length);
  check("Ctrl+Enter w środku akapitu: tekst za kursorem na nowej stronie", pagesAfter === pagesBefore + 1 && ps[5].text === "Treść " && ps[6].text.startsWith("rozdziału"), JSON.stringify({ pagesBefore, pagesAfter, a: ps[5]?.text, b: ps[6]?.text?.slice(0, 20) }));
  check("obie części zachowują wyjustowanie", ps[5].align === "justify" && ps[6].align === "justify", `${ps[5].align}/${ps[6].align}`);
  // podział z menu na końcu ostatniego akapitu → nowa pusta strona, kursor tam
  await caretTo(page, 6, true);
  await page.click("#insertMenuBtn");
  await page.click(".compose-item-label:text-is('Podział strony')");
  await idle(page);
  const pages3 = await page.evaluate(() => document.querySelectorAll(".docx-preview-host section.docx").length);
  const focusPb = await page.evaluate(() => { const p = document.activeElement?.closest?.(".docx-editable-p"); return p ? resolveParaIndex(p) : -1; });
  check("podział strony na końcu: nowa strona, kursor w pustym akapicie na niej", pages3 === pagesAfter + 1 && focusPb === 7, `${pages3} / ${focusPb}`);
  await page.keyboard.type("Strona trzecia");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Dalej ta sama strona");
  await idle(page);
  const pages4 = await page.evaluate(() => document.querySelectorAll(".docx-preview-host section.docx").length);
  await page.evaluate(() => rerenderKeepingEdits());
  await idle(page);
  const pages5 = await page.evaluate(() => document.querySelectorAll(".docx-preview-host section.docx").length);
  check("Enter za podziałem NIE robi kolejnej strony (po przerysowaniu też)", pages4 === pages3 && pages5 === pages3, `${pages3} → ${pages4} → ${pages5}`);
  const undo2 = await page.evaluate(() => dwbUndo._debug().undo.map((l) => t(l)));
  check("historia cofania nazywa kroki (podział strony, linia, formatowanie)", ["podział strony", "linia pozioma", "formatowanie akapitu"].every((k) => undo2.includes(k)), JSON.stringify(undo2));

  // ── zapisany plik ───────────────────────────────────────────────────────────
  xml = await fileXml(page);
  const styles = await fileXml(page, "word/styles.xml");
  check("plik: kolejność w:pPr zgodna ze schematem Worda", pPrOrderOk(xml).length === 0, JSON.stringify(pPrOrderOk(xml)));
  check("plik: tytuł/nagłówek przez style (Title, Heading1), 2× pageBreakBefore, 1 linia (pBdr)", /w:pStyle w:val="Title"/.test(xml) && /w:pStyle w:val="Heading1"/.test(xml) && (xml.match(/<w:pageBreakBefore\/>/g) || []).length === 2 && (xml.match(/<w:pBdr>/g) || []).length === 1, "");
  check("plik: jeden w:sectPr (sekcja nie zdublowana przy dzieleniu)", (xml.match(/<w:sectPr/g) || []).length === 1);
  check("plik: style.xml ma nazwy wbudowane (Word rozpozna „Nagłówek 1”)", /w:name w:val="heading 1"/.test(styles) && /w:name w:val="Title"/i.test(styles));

  // ── szablony ────────────────────────────────────────────────────────────────
  page.removeAllListeners("dialog");
  page.on("dialog", (d) => d.accept()); // „porzucić niezapisane zmiany?”
  await page.keyboard.press(`${MOD}+Alt+KeyN`);
  check("Ctrl/⌘+Alt+N otwiera okno szablonów", await page.evaluate(() => document.getElementById("newDocDialog").open));
  await page.click('.newdoc-card[data-template="note"]');
  await page.waitForFunction(() => currentFileName === "Notatka.docx", null, { timeout: 15000 });
  await idle(page);
  ps = await paras(page);
  const structure = await page.evaluate(() => (documentStructure?.headings || []).map((h) => h.text || h.title));
  check("szablon Notatka: tytuł + nagłówki widoczne w strukturze", ps[0].style === "title" && ps.filter((p) => p.style === "h1").length === 3 && structure.length >= 3, JSON.stringify({ s: ps.map((p) => p.style), structure }));
  await page.click("#appMenuBtn");
  await page.click("#newDocMenuItem");
  await page.click('.newdoc-card[data-template="letter"]');
  await page.waitForFunction(() => currentFileName === "Pismo.docx", null, { timeout: 15000 });
  await idle(page);
  const fields = await page.evaluate(() => (docCanvasEl.textContent.match(/\{\{[^}]+\}\}/g) || []).length);
  ps = await paras(page);
  check("szablon Pismo (z menu ⋯): pola {{…}} do wypełnienia, data do prawej", fields >= 6 && ps.some((p) => p.align === "right"), `${fields}`);

  // ── plik bez styles.xml (z generatora) → styl dopisany ──────────────────────
  await page.evaluate(() => loadSampleDocument("sample"));
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);
  await caretTo(page, 1, true);
  await page.selectOption("#fmtParaStyle", "quote");
  await idle(page);
  const q = await fileXml(page, "word/styles.xml");
  const ct = await fileXml(page, "[Content_Types].xml");
  const rels = await fileXml(page, "word/_rels/document.xml.rels");
  ps = await paras(page);
  check("plik bez styles.xml: Cytat dopisany (część, typ, powiązanie) i widoczny", q && /w:name w:val="quote"/.test(q) && /styles\.xml/.test(ct) && /relationships\/styles/.test(rels) && ps[1].style === "quote", JSON.stringify(ps[1]));

  // ── telefon: jeden rząd, przewijanie w bok, okienko na ekranie ──────────────
  await page.setViewportSize({ width: 390, height: 780 });
  await page.waitForTimeout(500);
  await idle(page);
  const bar = await page.evaluate(() => {
    const b = document.getElementById("formatToolbar");
    const kids = Array.from(b.children).filter((c) => c.offsetParent);
    const tops = new Set(kids.map((c) => { const r = c.getBoundingClientRect(); return Math.round((r.top + r.bottom) / 2 / 6); }));
    return { rows: tops.size, scroll: b.scrollWidth > b.clientWidth, fade: b.classList.contains("more-r"), h: b.getBoundingClientRect().height, pageScroll: document.documentElement.scrollWidth > innerWidth };
  });
  check("telefon: pasek Edycji w jednym rzędzie, przewija się w bok z wygaszeniem, strona się nie rozjeżdża", bar.rows === 1 && bar.h < 48 && !bar.pageScroll && (!bar.scroll || bar.fade), JSON.stringify(bar));
  await caretTo(page, 0, true);
  await page.click("#insertMenuBtn");
  const popBox = await page.evaluate(() => { const r = document.querySelector(".compose-pop").getBoundingClientRect(); return { l: r.left, r: r.right, b: r.bottom, w: innerWidth, h: innerHeight }; });
  check("telefon: okienko „Wstaw” mieści się na ekranie", popBox.l >= 0 && popBox.r <= popBox.w && popBox.b <= popBox.h, JSON.stringify(popBox));
  await page.mouse.click(200, 700);
  check("klik obok zamyka okienko", !(await page.$(".compose-pop")));
  await page.setViewportSize({ width: 1280, height: 860 });
  // tryb Czytanie: nic z tego nie widać
  await page.evaluate(() => appFrame.setReadOnly(true));
  await page.waitForTimeout(200);
  check("tryb Czytanie: przyciski tworzenia schowane", await page.evaluate(() => getComputedStyle(document.getElementById("insertMenuBtn")).display === "none" || !document.getElementById("insertMenuBtn").offsetParent));

  check("brak błędów strony", errors.length === 0, errors.join(" | "));
  await browser.close();

  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || r.detail == null ? "" : ` — ${String(r.detail).slice(0, 400)}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
