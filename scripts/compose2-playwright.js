// compose2-playwright.js — tworzenie dokumentu, Etap 2: listy, linki, pola formularza, spis treści.
//
// Jak użytkownik: przycisk „Lista” (punktowana / numerowana / poziom), Enter w pustym punkcie
// i Backspace na początku punktu (koniec listy), Ctrl/⌘+K (adres WWW i miejsce w dokumencie),
// karta linku pod kursorem (zmień / usuń), „＋ Wstaw” → pola formularza w zdaniu (akapit dalej
// edytowalny, pole przetrwa pisanie obok i Enter), spis treści (wstaw + aktualizuj).
// Na końcu ZAPISANY plik: hiperłącza z powiązaniem, numbering.xml, kontrolki, pole TOC.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(350));

const paras = (page) => page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => ({
  text: p.textContent, list: isListParagraph(p), level: +((p.className.match(/docx-num-\d+-(\d+)/) || [])[1] || 0), lock: p.dataset.lock || "",
  editable: p.isContentEditable, links: p.querySelectorAll("a[href]").length, fields: p.querySelectorAll(".ff-field").length,
})));
const idxOf = (page, start) => page.evaluate((s) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).findIndex((q) => q.textContent.startsWith(s)), start);
const caretEnd = (page, i) => page.evaluate((x) => {
  const el = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[x];
  const r = document.createRange(); r.selectNodeContents(el); r.collapse(false);
  el.focus(); getSelection().removeAllRanges(); getSelection().addRange(r);
}, i);
const caretAt = (page, i, o) => page.evaluate(([x, off]) => focusParagraphAtOffset(x, off), [i, o]);
const selectText = (page, i, from, to) => page.evaluate(([x, a, b]) => {
  const el = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[x];
  const s = formDomRange(el, a, b);
  el.focus(); getSelection().removeAllRanges(); getSelection().addRange(s);
}, [i, from, to]);
const focusedIndex = (page) => page.evaluate(() => { const q = docCaretParagraph(document.activeElement); return q ? resolveParaIndex(q) : -1; });
const savedZip = async (page) => {
  const b64 = await page.evaluate(async () => {
    const bytes = await buildDocumentForSave();
    let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s);
  });
  return JSZip.loadAsync(Buffer.from(b64, "base64"));
};
const insertMenu = async (page, label) => {
  await page.click("#insertMenuBtn");
  await page.click(`.compose-pop-insert .compose-item-label:text-is("${label}")`);
};

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => {
    sessionStorage.setItem("introPlayed", "true");
    window.__opened = [];
    window.open = (...args) => { window.__opened.push(args); return null; };
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);

  // ── listy ───────────────────────────────────────────────────────────────────
  await page.keyboard.type("Zakupy:");
  await page.keyboard.press("Enter");
  await page.keyboard.type("mleko");
  await page.click("#fmtListBtn");
  await page.click(".compose-pop-list .compose-item-label:text-is('Lista punktowana')");
  await idle(page);
  await caretEnd(page, 1);
  await page.keyboard.press("Enter"); await page.keyboard.type("chleb");
  await page.keyboard.press("Enter"); await page.keyboard.press("Tab"); await page.keyboard.type("razowy");
  await page.waitForTimeout(400);
  let ps = await paras(page);
  check("lista punktowana z przycisku „Lista”, Enter = kolejny punkt, Tab = głębiej", ps[1].list && ps[2].list && ps[3].list && ps[3].level === 1 && !ps[0].list, JSON.stringify(ps.map((p) => [p.text, p.list, p.level])));
  // poziom z okienka (tablet bez Tab)
  await page.click("#fmtListBtn");
  await page.click(".compose-pop-list .compose-row-btn:has-text('Płycej')");
  await page.waitForTimeout(400);
  ps = await paras(page);
  check("okienko listy: „Płycej” (dotyk, bez klawisza Tab)", ps[3].level === 0, ps[3].level);
  await caretEnd(page, 3);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(250);
  await page.keyboard.press("Enter"); // pusty punkt → koniec listy
  await idle(page);
  await page.keyboard.type("Kroki:");
  ps = await paras(page);
  check("Enter w pustym punkcie kończy listę (dalej zwykły akapit)", ps.length === 5 && !ps[4].list && ps[4].text === "Kroki:", JSON.stringify(ps.map((p) => [p.text, p.list])));
  await page.keyboard.press("Enter"); await page.keyboard.type("Krok pierwszy");
  await page.click("#fmtListBtn");
  await page.click(".compose-pop-list .compose-item-label:text-is('Lista numerowana')");
  await idle(page);
  await caretEnd(page, 5);
  await page.keyboard.press("Enter"); await page.keyboard.type("Krok drugi");
  await page.waitForTimeout(400);
  const nums = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).slice(5, 7).map((q) => getComputedStyle(q, "::before").content));
  check("lista numerowana: znaczniki z licznikiem (1., 2.)", nums.every((c) => /counter\(/.test(c)), JSON.stringify(nums));
  await caretAt(page, 6, 0);
  await page.keyboard.press("Backspace");
  await idle(page);
  ps = await paras(page);
  check("Backspace na początku punktu zdejmuje numerację, tekst zostaje", !ps[6].list && ps[6].text === "Krok drugi" && ps[5].list, JSON.stringify(ps.slice(5).map((p) => [p.text, p.list])));
  await caretEnd(page, 1);
  await page.click("#fmtListBtn");
  await page.click(".compose-pop-list .compose-item-label:text-is('Lista punktowana')"); // drugi raz = zdejmij
  await idle(page);
  ps = await paras(page);
  check("ta sama lista drugi raz zdejmuje ją z akapitu", !ps[1].list && ps[2].list, JSON.stringify(ps.slice(0, 4).map((p) => [p.text, p.list])));

  // ── linki ──────────────────────────────────────────────────────────────────
  await caretEnd(page, 6);
  await page.keyboard.press("Enter");
  await page.keyboard.type("Zobacz stronę firmy teraz.");
  await page.waitForTimeout(250);
  let i = await idxOf(page, "Zobacz");
  await selectText(page, i, 7, 19);
  await page.keyboard.press(`${MOD}+KeyK`);
  const form = await page.evaluate(() => ({ open: !!document.querySelector(".compose-pop-form"), text: document.querySelector(".lf-text")?.value, focus: document.activeElement?.className }));
  check("Ctrl/⌘+K: okienko linku z zaznaczonym tekstem, kursor w polu adresu", form.open && form.text === "stronę firmy" && form.focus === "lf-url", JSON.stringify(form));
  await page.keyboard.type("example.com/oferta");
  await page.keyboard.press("Enter");
  await idle(page);
  ps = await paras(page);
  check("link wstawiony, akapit dalej edytowalny", ps[i].links === 1 && ps[i].editable && !ps[i].lock, JSON.stringify(ps[i]));
  await page.keyboard.type("!");
  await page.waitForTimeout(250);
  const linkText = await page.evaluate((x) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[x].querySelector("a").textContent, i);
  check("pisanie zaraz za linkiem nie wydłuża linku", linkText === "stronę firmy", linkText);
  await caretAt(page, i, 10);
  await page.waitForTimeout(300);
  const card = await page.evaluate(() => document.querySelector(".link-card")?.textContent || "");
  check("kursor w linku → karta linku z adresem", /example\.com\/oferta/.test(card), card);
  await page.click(".link-card .lc-open");
  const opened = await page.evaluate(() => window.__opened.map((a) => a[0]));
  check("karta: „Otwórz” → nowa karta", opened.includes("https://example.com/oferta"), JSON.stringify(opened));
  await caretAt(page, i, 10);
  await page.waitForTimeout(200);
  await page.click(".link-card .lc-edit");
  await page.fill(".lf-url", "example.org");
  await page.click(".lf-ok");
  await idle(page);
  const href2 = await page.evaluate((x) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[x].querySelector("a").getAttribute("href"), i);
  check("karta: „Zmień” podmienia adres", href2 === "https://example.org", href2);
  // link wewnętrzny
  await caretEnd(page, i);
  await page.keyboard.press("Enter");
  await page.keyboard.type("Rozdział");
  await page.selectOption("#fmtParaStyle", "h1");
  await idle(page);
  await caretEnd(page, i);
  await page.keyboard.press(`${MOD}+KeyK`);
  await page.click(".lf-mode button[data-mode=doc]");
  await page.selectOption(".lf-target", { label: "Rozdział" });
  await page.fill(".lf-text", " (do rozdziału)");
  await page.click(".lf-ok");
  await idle(page);
  const internal = await page.evaluate((x) => [...collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[x].querySelectorAll("a")].map((a) => a.getAttribute("href")), i);
  check("link do miejsca w dokumencie (nagłówek dostaje zakładkę)", internal.length === 2 && /^#_Ref\d+/.test(internal[1]), JSON.stringify(internal));
  await caretAt(page, i, ps[i]?.text.length ? (await paras(page))[i].text.length - 4 : 0); // w „(do rozdziału)”
  await page.waitForTimeout(200);
  await page.click(".link-card .lc-remove");
  await idle(page);
  ps = await paras(page);
  check("karta: „Usuń link” zdejmuje link, tekst zostaje", ps[i].links === 1 && ps[i].text.endsWith("(do rozdziału)"), JSON.stringify(ps[i]));

  // ── pola formularza ────────────────────────────────────────────────────────
  await caretEnd(page, i);
  await page.keyboard.press("Enter");
  await page.keyboard.type("Imię: ");
  await insertMenu(page, "Tekst");
  await idle(page);
  i = await idxOf(page, "Imię");
  ps = await paras(page);
  check("pole tekstowe wstawione w zdanie, akapit dalej edytowalny", ps[i].fields === 1 && ps[i].editable && !ps[i].lock, JSON.stringify(ps[i]));
  await caretEnd(page, i);
  await page.keyboard.type(" Dział: ");
  await insertMenu(page, "Lista wyboru");
  await page.fill(".lf-options", "Sprzedaż\nIT");
  await page.click(".lf-ok");
  await idle(page);
  await caretEnd(page, i);
  await page.keyboard.type(" Od: ");
  await insertMenu(page, "Data");
  await idle(page);
  await caretEnd(page, i);
  await page.keyboard.type(" Zgoda: ");
  await insertMenu(page, "Pole wyboru");
  await idle(page);
  await caretEnd(page, i);
  await page.keyboard.type(" koniec.");
  await page.waitForTimeout(300);
  const kinds = await page.evaluate(() => formScan?.fields.map((f) => f.kind).join(","));
  check("4 pola: tekst, lista, data, ☐ (panel Formularz je widzi)", kinds === "text,dropdown,date,checkbox", kinds);
  let zip = await savedZip(page);
  let xml = await zip.file("word/document.xml").async("string");
  check("pisanie obok pól: w pliku 4 kontrolki i dopisany tekst", (xml.match(/<w:sdt[ >]/g) || []).length === 4 && /Zgoda: <\/w:t>.*koniec\./.test(xml.replace(/<w:sdt[ >].*?<\/w:sdt>/g, "")), (xml.match(/<w:sdt[ >]/g) || []).length);
  // wypełnienie pola po wpisaniu tekstu obok (kolejność: zapis akapitu, potem pole)
  await page.evaluate(() => {
    const k = formScan.fields.find((f) => f.kind === "checkbox").key;
    const r = formUi.ranges.get(k); const el = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
    el.closest(".ff-field").click();
  });
  await idle(page);
  zip = await savedZip(page);
  xml = await zip.file("word/document.xml").async("string");
  check("klik w ☐ zaznacza, tekst dopisany obok przetrwał", /w14:checked w14:val="1"/.test(xml) && /koniec\./.test(xml), "");
  // Enter w akapicie z polami: pola idą z tekstem do właściwych akapitów
  await caretAt(page, i, 5); // „Imię:|”
  await page.keyboard.press("Enter");
  await idle(page);
  zip = await savedZip(page);
  xml = await zip.file("word/document.xml").async("string");
  check("Enter w akapicie z polami: dalej 4 kontrolki", (xml.match(/<w:sdt[ >]/g) || []).length === 4, (xml.match(/<w:sdt[ >]/g) || []).length);

  // ── spis treści ────────────────────────────────────────────────────────────
  await caretAt(page, 0, 0);
  await insertMenu(page, "Spis treści");
  await idle(page);
  ps = await paras(page);
  check("spis treści na początku: tytuł + wpis „Rozdział” (zablokowany, z linkiem)", ps[0].text === "Spis treści" && /^Rozdział\t?\s*\d*$/.test(ps[1].text.replace(/\s+$/, "")) && ps[1].lock === "lockField" && ps[1].links === 1, JSON.stringify(ps.slice(0, 3)));
  check("kursor wraca do tekstu za spisem", (await focusedIndex(page)) === 2, await focusedIndex(page));
  const last = (await paras(page)).length - 1;
  await caretEnd(page, last);
  await page.keyboard.press("Enter");
  await page.keyboard.type("Zakończenie");
  await page.selectOption("#fmtParaStyle", "h2");
  await idle(page);
  await page.click("#insertMenuBtn");
  const tocLabel = await page.evaluate(() => [...document.querySelectorAll(".compose-pop-insert .compose-item-label")].map((x) => x.textContent).find((x) => /spis/i.test(x)));
  check("gdy spis jest — w menu „Zaktualizuj spis treści”", tocLabel === "Zaktualizuj spis treści", tocLabel);
  await page.click(`.compose-pop-insert .compose-item-label:text-is("Zaktualizuj spis treści")`);
  await idle(page);
  ps = await paras(page);
  check("aktualizacja dopisuje nowy nagłówek, spis nie podwaja się", ps.filter((p) => p.text === "Spis treści").length === 1 && ps.slice(1, 4).some((p) => p.text.startsWith("Zakończenie")), JSON.stringify(ps.slice(0, 4).map((p) => p.text)));
  await page.click(".docx-preview-host a[href^='#_Ref']");
  await page.waitForTimeout(900);
  check("klik we wpis spisu przewija do nagłówka", await page.evaluate(() => !document.querySelector(".link-back")?.hidden), "");

  // ── zapisany plik ──────────────────────────────────────────────────────────
  zip = await savedZip(page);
  xml = await zip.file("word/document.xml").async("string");
  const rels = await zip.file("word/_rels/document.xml.rels").async("string");
  const ct = await zip.file("[Content_Types].xml").async("string");
  const numbering = await zip.file("word/numbering.xml")?.async("string");
  const styles = await zip.file("word/styles.xml").async("string");
  check("plik: link WWW przez powiązanie (rId → example.org, External)", /<w:hyperlink r:id="rId\d+"/.test(xml) && /Target="https:\/\/example\.org" TargetMode="External"/.test(rels), "");
  check("plik: numbering.xml dopisany (część, typ, powiązanie), definicje przed numerami", numbering && /numbering\+xml/.test(ct) && /relationships\/numbering/.test(rels) && numbering.indexOf("<w:abstractNum ") < numbering.indexOf("<w:num "), "");
  check("plik: pole TOC (\\o 1-3 \\h) + style Spis treści 1/2 i Hiperłącze, Tekst zastępczy", /TOC \\o "1-3" \\h/.test(xml) && /w:name w:val="toc 1"/.test(styles) && /w:name w:val="Hyperlink"/.test(styles) && /w:name w:val="Placeholder Text"/.test(styles), "");
  check("plik: kolejne sklejenia nie zostawiają pustych <w:hyperlink>", !/<w:hyperlink[^>]*\/>|<w:hyperlink[^>]*><\/w:hyperlink>/.test(xml), "");

  // ── dotyk (iPhone/iPad): tap w przyciski i pozycje okienek ──────────────────
  // Błąd z 2026-10-01: preventDefault na pointerdown w Safari kasował kliknięcie palcem.
  const touch = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 800 }, isMobile: ENGINE !== "firefox", hasTouch: true });
  await touch.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const tp = await touch.newPage();
  tp.on("pageerror", (e) => errors.push(`touch: ${e.message}`));
  await tp.goto(APP_URL, { waitUntil: "load" });
  await tp.evaluate(() => document.getElementById("heroSplash")?.remove());
  await tp.evaluate(() => composeUi.createNew("note"));
  await tp.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(tp);
  await tp.evaluate(() => focusParagraphAtOffset(3, 0));
  await tp.tap("#insertMenuBtn");
  check("dotyk: tap „＋ Wstaw” otwiera menu", await tp.evaluate(() => !!document.querySelector(".compose-pop-insert")));
  await tp.tap(".compose-pop-insert .compose-char[aria-label='§']");
  await tp.waitForTimeout(250);
  check("dotyk: tap w znak wstawia go w tekst", (await tp.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[3].textContent)).startsWith("§"));
  await tp.tap("#fmtListBtn");
  await tp.tap(".compose-pop-list .compose-item-label:text-is('Lista numerowana')");
  await idle(tp);
  check("dotyk: tap „Lista” → „Lista numerowana” robi listę", await tp.evaluate(() => isListParagraph(collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[3])));
  await tp.evaluate(() => focusParagraphAtOffset(3, 1));
  await tp.tap("#fmtAlignBtn");
  await tp.tap(".compose-pop-align button[aria-label='Do prawej']");
  await idle(tp);
  check("dotyk: tap wyrównania działa", await tp.evaluate(() => getComputedStyle(collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[3]).textAlign) === "right");
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
