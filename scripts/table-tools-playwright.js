// table-tools-playwright.js — scalanie, dzielenie, nagłówek, kolor, krawędzie tabeli (2026-10-10).
//
// Jak Układ / Projekt tabeli w Wordzie, z menu „Tabela” na pasku Edycji: Scal zaznaczone komórki,
// Scal z komórką po prawej / poniżej (wygodne na dotyku), Podziel komórkę (scaloną — rozdziel,
// zwykłą — na dwie kolumny), Powtarzaj jako wiersz nagłówka (w:tblHeader), krawędzie (wszystkie /
// zewnętrzne / brak), równe kolumny, kolor komórki (w:shd, zakres komórka/wiersz/kolumna/tabela).
// Sprawdzamy podgląd (colspan/rowspan) i zapisany XML (gridSpan, vMerge, tblGrid). ENGINE=webkit.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

// tabela z pliku: wiersze → komórki { t: tekst, span, vm, shd }, siatka, nagłówki, krawędzie
const savedTable = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  const doc = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const tbl = doc.getElementsByTagNameNS(W, "tbl")[0];
  if (!tbl) return null;
  const kid = (el, name) => el && [...el.childNodes].find((n) => n.localName === name);
  const val = (el, a = "val") => el ? (el.getAttributeNS(W, a) || el.getAttribute(`w:${a}`) || "") : null;
  const rows = [...tbl.childNodes].filter((n) => n.localName === "tr").map((tr) => ({
    header: !!kid(kid(tr, "trPr"), "tblHeader"),
    cells: [...tr.childNodes].filter((n) => n.localName === "tc").map((tc) => {
      const pr = kid(tc, "tcPr");
      return { t: [...tc.childNodes].filter((n) => n.localName === "p").map((p) => p.textContent).filter(Boolean).join("+"), span: parseInt(val(kid(pr, "gridSpan")) || "1", 10), vm: kid(pr, "vMerge") ? (val(kid(pr, "vMerge")) || "continue") : "", shd: val(kid(pr, "shd"), "fill") || "" };
    }),
  }));
  const grid = [...(kid(tbl, "tblGrid")?.childNodes || [])].filter((n) => n.localName === "gridCol").map((g) => parseInt(val(g, "w"), 10));
  const borders = kid(kid(tbl, "tblPr"), "tblBorders");
  return { rows, grid, borders: borders ? [...borders.childNodes].map((b) => `${b.localName}:${val(b)}`).join(",") : "" };
});
const text = (t) => t.rows.map((r) => r.cells.map((c) => c.t + (c.span > 1 ? `^${c.span}` : "") + (c.vm ? `|${c.vm}` : "")).join(" ; ")).join(" / ");

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.addInitScript(() => { window.__dwbAutoDialog = "accept"; }); // własne okna apki jak systemowe
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);
  await page.keyboard.type("Tabela:");
  await page.evaluate(() => composeUi.insertTable(3, 3));
  await idle(page);
  for (const w of ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"]) { await page.keyboard.type(w); if (w !== "C3") await page.keyboard.press("Tab"); await page.waitForTimeout(100); }
  await idle(page);

  const settle = () => page.evaluate(() => new Promise((res) => {
    const vp = document.getElementById("docViewport");
    let last = vp.scrollTop, still = 0;
    const tick = () => { if (vp.scrollTop === last) still++; else { still = 0; last = vp.scrollTop; } if (still >= 6) res(); else requestAnimationFrame(tick); };
    tick();
  }));
  const inCell = (word) => page.evaluate((word) => {
    const p = [...document.querySelectorAll(".docx-preview-host td .docx-editable-p")].find((x) => x.textContent.trim() === word || x.textContent.includes(word));
    const n = document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode();
    const r = document.createRange();
    r.setStart(n || p, n ? 1 : 0);
    r.collapse(true);
    p.closest(".docx-edit-root").focus({ preventScroll: true });
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  }, word);
  const menu = async (key) => {
    await settle();
    await page.waitForTimeout(150);
    await page.waitForFunction(() => !document.getElementById("tableToolsBtn").hidden).then(() => page.evaluate(() => typeof dwbRibbon !== "undefined" && dwbRibbon?.isOn() && dwbRibbon.select("table"))).then(() => page.evaluate(() => document.getElementById("tableToolsBtn").click()));
    await page.waitForSelector(`.compose-pop-tabletools [data-tbl="${key}"]`, { timeout: 8000 });
    const box = await page.evaluate((key) => { const b = document.querySelector(`.compose-pop-tabletools [data-tbl="${key}"]`); b.scrollIntoView({ block: "nearest" }); const r = b.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; }, key);
    await page.mouse.click(box.x, box.y);
    await idle(page);
  };

  let t = await savedTable(page);
  check("tabela 3×3 wypełniona", text(t) === "A1 ; B1 ; C1 / A2 ; B2 ; C2 / A3 ; B3 ; C3", text(t));
  const grid0 = t.grid;

  // ── scal z prawą ────────────────────────────────────────────────────────────
  await inCell("A1");
  await menu("mergeRight");
  t = await savedTable(page);
  check("Scal z prawą: A1+B1 w jednej komórce (gridSpan=2), szerokość = suma kolumn", t.rows[0].cells.length === 2 && t.rows[0].cells[0].t === "A1+B1" && t.rows[0].cells[0].span === 2, text(t));
  check("podgląd: komórka z colspan=2", await page.evaluate(() => document.querySelector(".docx-preview-host table tr td")?.colSpan === 2));

  // ── podziel (rozdziel scalenie) ────────────────────────────────────────────
  await inCell("A1");
  await menu("split");
  t = await savedTable(page);
  check("Podziel scaloną: znów 3 komórki, treść w pierwszej", t.rows[0].cells.length === 3 && t.rows[0].cells.every((c) => c.span === 1) && t.rows[0].cells[0].t === "A1+B1", text(t));

  // ── scal z dolną ────────────────────────────────────────────────────────────
  await inCell("A2");
  await menu("mergeDown");
  t = await savedTable(page);
  check("Scal z dolną: A2 restart, A3 continue, treść A2+A3 u góry", t.rows[1].cells[0].vm === "restart" && t.rows[2].cells[0].vm === "continue" && t.rows[1].cells[0].t === "A2+A3" && t.rows[2].cells[0].t === "", text(t));
  check("podgląd: komórka z rowspan=2", await page.evaluate(() => [...document.querySelectorAll(".docx-preview-host table td")].some((td) => td.rowSpan === 2)));

  // ── scal zaznaczone (prostokąt B2..C3) ──────────────────────────────────────
  await page.evaluate(() => {
    const ps = [...document.querySelectorAll(".docx-preview-host td .docx-editable-p")];
    const a = ps.find((x) => x.textContent.includes("B2")), b = ps.find((x) => x.textContent.includes("C3"));
    const w = (p) => document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode();
    const r = document.createRange();
    r.setStart(w(a), 0);
    r.setEnd(w(b), 2);
    a.closest(".docx-edit-root").focus({ preventScroll: true });
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  });
  await menu("mergeSel");
  t = await savedTable(page);
  check("Scal zaznaczone B2..C3: jedna komórka 2×2 z treścią B2+C2+B3+C3", t.rows[1].cells[1]?.span === 2 && t.rows[1].cells[1].vm === "restart" && t.rows[1].cells[1].t === "B2+C2+B3+C3" && t.rows[2].cells[1]?.vm === "continue" && t.rows[2].cells.length === 2, text(t));

  // ── podziel zwykłą komórkę na dwie kolumny ─────────────────────────────────
  await inCell("C1");
  await menu("split");
  t = await savedTable(page);
  check("Podziel zwykłą C1: siatka 4 kolumny, wiersz 1 ma 4 komórki, szerokość tabeli bez zmian", t.grid.length === 4 && t.rows[0].cells.length === 4 && Math.abs(t.grid.reduce((a, b) => a + b, 0) - grid0.reduce((a, b) => a + b, 0)) <= 2, JSON.stringify({ grid: t.grid, t: text(t) }));
  check("…w innych wierszach komórka pod C1 szersza (gridSpan+1)", t.rows[1].cells[1].span === 3 && t.rows[2].cells[1].span === 3, text(t));

  // ── równe kolumny ───────────────────────────────────────────────────────────
  await inCell("A1");
  await menu("evenCols");
  t = await savedTable(page);
  check("Rozłóż kolumny równomiernie: wszystkie kolumny siatki tej samej szerokości", new Set(t.grid).size === 1, JSON.stringify(t.grid));

  // ── wiersz nagłówka ─────────────────────────────────────────────────────────
  await inCell("A1");
  await menu("headerRow");
  t = await savedTable(page);
  check("Powtarzaj jako nagłówek: w:tblHeader tylko w 1. wierszu", t.rows[0].header && !t.rows[1].header && !t.rows[2].header, JSON.stringify(t.rows.map((r) => r.header)));
  await settle();
  await inCell("A1");
  await page.waitForFunction(() => !document.getElementById("tableToolsBtn").hidden).then(() => page.evaluate(() => typeof dwbRibbon !== "undefined" && dwbRibbon?.isOn() && dwbRibbon.select("table"))).then(() => page.evaluate(() => document.getElementById("tableToolsBtn").click()));
  await page.waitForSelector('.compose-pop-tabletools [data-tbl="headerRow"]');
  check("menu pokazuje włączony nagłówek (aria-checked)", await page.evaluate(() => document.querySelector('.compose-pop-tabletools [data-tbl="headerRow"]').getAttribute("aria-checked") === "true"));
  await page.keyboard.press("Escape");
  await menu("headerRow");
  t = await savedTable(page);
  check("drugi raz: nagłówek wyłączony", t.rows.every((r) => !r.header), JSON.stringify(t.rows.map((r) => r.header)));

  // ── kolor komórki (zakres: wiersz) ──────────────────────────────────────────
  await settle();
  await inCell("A1");
  await page.waitForFunction(() => !document.getElementById("tableToolsBtn").hidden).then(() => page.evaluate(() => typeof dwbRibbon !== "undefined" && dwbRibbon?.isOn() && dwbRibbon.select("table"))).then(() => page.evaluate(() => document.getElementById("tableToolsBtn").click()));
  await page.waitForSelector(".compose-pop-tabletools .table-shades .color-swatch");
  await page.evaluate(() => {
    const pop = document.querySelector(".compose-pop-tabletools");
    const seg = pop.querySelector(".compose-shade-scope");
    seg.querySelector('button[data-v="row"]').click();
    pop.querySelector(".table-shades .color-swatch").click();
  });
  await idle(page);
  t = await savedTable(page);
  check("Kolor komórki dla wiersza: każda komórka 1. wiersza ma w:shd fill", t.rows[0].cells.every((c) => /^[0-9A-F]{6}$/.test(c.shd)) && t.rows[1].cells.every((c) => !c.shd), JSON.stringify(t.rows.map((r) => r.cells.map((c) => c.shd))));
  check("podgląd: tło komórki", await page.evaluate(() => { const bg = getComputedStyle(document.querySelector(".docx-preview-host table tr td")).backgroundColor; return bg && bg !== "rgba(0, 0, 0, 0)" && bg !== "transparent"; }));

  // ── krawędzie ───────────────────────────────────────────────────────────────
  await inCell("A1");
  await menu("borders-none");
  t = await savedTable(page);
  check("Bez krawędzi: tblBorders wszystkie „none”", /^top:none,left:none,bottom:none,right:none,insideH:none,insideV:none$/.test(t.borders), t.borders);
  check("podgląd: komórki bez linii", await page.evaluate(() => { const cs = getComputedStyle(document.querySelector(".docx-preview-host table tr td")); return parseFloat(cs.borderTopWidth) === 0 || cs.borderTopStyle === "none"; }));
  await inCell("A1");
  await menu("borders-outside");
  t = await savedTable(page);
  check("Tylko zewnętrzne: ramka single, w środku none", t.borders === "top:single,left:single,bottom:single,right:single,insideH:none,insideV:none", t.borders);

  // ── Cofnij ──────────────────────────────────────────────────────────────────
  await settle();
  await page.evaluate((s) => { if (typeof dwbRibbon !== "undefined" && dwbRibbon?.isOn()) { if (document.body.classList.contains("rb-reading")) document.querySelector('.mode-btn[data-mode="edit"]').click(); dwbRibbon.reveal(s); } }, "#undoBtn").then(async () => page.click("#undoBtn"));
  await idle(page);
  t = await savedTable(page);
  check("Cofnij: krawędzie wracają do „bez krawędzi”", /insideH:none/.test(t.borders) && /^top:none/.test(t.borders), t.borders);

  // ── Sortuj (kolejność jak w Wordzie: tools/word-oracle/run-tables.js „sortuj-*”) ─────────────
  // pogrubiony pierwszy wiersz = nagłówek, zostaje na górze; kolumna z kursorem to liczby → 41, 30, 25, 9
  const tc = (s, b) => `<w:tc><w:p><w:r>${b ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t>${s}</w:t></w:r></w:p></w:tc>`;
  const sortRows = [["Imię", "Wiek"], ["Ola", "30"], ["Ewa", "25"], ["Adam", "41"], ["Ewa", "9"]];
  const sortDoc = await buildDocx({ body: ["Lista:", `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${sortRows.map((r, i) => `<w:tr>${r.map((s) => tc(s, i === 0)).join("")}</w:tr>`).join("")}</w:tbl>`, "Koniec."] });
  await page.locator("#fileInput").setInputFiles({ name: "sort.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: sortDoc });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && [...document.querySelectorAll(".docx-preview-host td p")].some((p) => p.textContent === "Adam"), null, { timeout: 30000 });
  await page.evaluate(() => { if (readOnlyMode) appFrame.setReadOnly(false); });
  await idle(page);
  await inCell("30");
  await menu("sort-desc");
  t = await savedTable(page);
  check("Sortuj malejąco wg kolumny z kursorem: nagłówek na górze, liczby 41, 30, 25, 9", text(t) === "Imię ; Wiek / Adam ; 41 / Ola ; 30 / Ewa ; 25 / Ewa ; 9", text(t));
  check("podgląd w nowej kolejności", await page.evaluate(() => [...document.querySelectorAll(".docx-preview-host table tr")].map((tr) => tr.cells[0].textContent.trim()).join(",")) === "Imię,Adam,Ola,Ewa,Ewa");
  const toastTxt = await page.evaluate(() => [...document.querySelectorAll(".toast")].map((x) => x.textContent).join(" | "));
  check("komunikat mówi, jak posortowano (kolumna z nagłówka, liczby) i ma „Zmień…”", /„Wiek” jako liczby, malejąco/.test(toastTxt) && /Zmień…/.test(toastTxt), toastTxt);
  // „Zmień…” z komunikatu otwiera okno z tym, co rozpoznano
  await page.locator(".toast .toast-action", { hasText: "Zmień…" }).last().click();
  await page.waitForSelector(".compose-pop-sort .srt-col option", { state: "attached", timeout: 8000 });
  const form = () => page.evaluate(() => {
    const el = document.querySelector(".compose-pop-sort");
    const k = [...el.querySelectorAll(".srt-key")].map((b) => ({ col: b.querySelector(".srt-col").selectedOptions[0]?.text, type: b.querySelector(".srt-type").value, order: b.querySelector(".srt-order .is-on")?.dataset.v, typeOff: b.querySelector(".srt-type").disabled }));
    return { k, head: el.querySelector(".srt-head").checked, cols: [...el.querySelector(".srt-col").options].map((o) => o.text).join(",") };
  });
  let f = await form();
  check("okno z „Zmień…”: Wiek, liczby, malejąco, nagłówek zaznaczony, drugi klucz pusty", f.k[0].col === "Wiek" && f.k[0].type === "num" && f.k[0].order === "desc" && f.head && f.k[1].col === "(brak)" && f.k[1].typeOff, JSON.stringify(f));
  await page.click(".compose-pop-sort .lf-cancel");
  // ── okno z menu: dwa klucze — Imię rosnąco, potem Wiek rosnąco (Ewa 9 przed Ewa 25) ──
  await inCell("Ola");
  await menu("sort-form");
  await page.waitForSelector(".compose-pop-sort .srt-col option", { state: "attached", timeout: 8000 });
  f = await form();
  check("okno z menu: kolumna z kursorem (Imię) jako tekst, nazwy kolumn z nagłówka", f.k[0].col === "Imię" && f.k[0].type === "text" && f.cols === "Imię,Wiek", JSON.stringify(f));
  await page.locator(".compose-pop-sort .srt-head").uncheck();
  f = await form();
  check("bez nagłówka: kolumny „Kolumna 1, 2”", f.cols === "Kolumna 1,Kolumna 2" && !f.head, JSON.stringify(f));
  await page.locator(".compose-pop-sort .srt-head").check();
  await page.locator('.compose-pop-sort .srt-key[data-k="1"] .srt-col').selectOption("1");
  f = await form();
  check("drugi klucz Wiek: typ sam ustawiony na liczby", f.k[1].col === "Wiek" && f.k[1].type === "num" && !f.k[1].typeOff, JSON.stringify(f));
  await page.click(".compose-pop-sort .lf-ok");
  await idle(page);
  t = await savedTable(page);
  check("Sortuj… dwa klucze: Adam, Ewa 9, Ewa 25, Ola", text(t) === "Imię ; Wiek / Adam ; 41 / Ewa ; 9 / Ewa ; 25 / Ola ; 30", text(t));
  // typ wybrany ręcznie: Wiek jako tekst malejąco
  await inCell("Ola");
  await menu("sort-form");
  await page.waitForSelector(".compose-pop-sort .srt-col option", { state: "attached", timeout: 8000 });
  await page.locator('.compose-pop-sort .srt-key[data-k="0"] .srt-col').selectOption("1");
  await page.locator('.compose-pop-sort .srt-key[data-k="0"] .srt-type').selectOption("text");
  await page.click('.compose-pop-sort .srt-key[data-k="0"] .srt-order [data-v="desc"]');
  await page.click(".compose-pop-sort .lf-ok");
  await idle(page);
  t = await savedTable(page);
  check("Sortuj… typ ręcznie (tekst, malejąco): „9” przed „41” jak w Wordzie", text(t) === "Imię ; Wiek / Ewa ; 9 / Adam ; 41 / Ola ; 30 / Ewa ; 25", text(t));
  await settle();
  await page.evaluate((s) => { if (typeof dwbRibbon !== "undefined" && dwbRibbon?.isOn()) { if (document.body.classList.contains("rb-reading")) document.querySelector('.mode-btn[data-mode="edit"]').click(); dwbRibbon.reveal(s); } }, "#undoBtn").then(async () => page.click("#undoBtn"));
  await idle(page);
  t = await savedTable(page);
  check("Cofnij sortowanie: poprzednia kolejność", text(t) === "Imię ; Wiek / Adam ; 41 / Ewa ; 9 / Ewa ; 25 / Ola ; 30", text(t));

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ narzędzia tabeli [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
