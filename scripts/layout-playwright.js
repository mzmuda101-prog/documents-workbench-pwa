// layout-playwright.js — Układ strony (marginesy jak w Wordzie), Podgląd wydruku i druk, nowe przypisy.
//
//   Wstaw → Układ strony: gotowe marginesy (Normalne/Wąskie/Umiarkowane/Szerokie), niestandardowe
//   (przecinek, ostrzeżenie przy marginesie poza zasięgiem drukarki, odmowa przy za dużych),
//   orientacja (obraca też marginesy), rozmiar papieru — w pliku (w:pgSz / w:pgMar) i na kartce;
//   Podgląd wydruku (Ctrl/⌘+P): tyle kartek, ile stron w Edycji, kartka o rozmiarze z pliku,
//   nagłówek na każdej, numer strony, ostrzeżenie o strefie poza zasięgiem drukarki, Esc zamyka;
//   druk (PDF z Chromium) = te same strony; w Widoku mobilnym podgląd dalej ma strony A4;
//   Wstaw → Przypis dolny/końcowy (Ctrl/⌘+Alt+F / D): numeracja wg kolejności w tekście, część
//   footnotes.xml tworzona w pliku bez przypisów, usunięcie odnośnika usuwa przypis i przenumerowuje.
// Fixture: docs/samples/selection-sample.docx. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1400, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.setItem("dwb-panel-docked-open-v1", "0"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=selection-sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForFunction(() => !readOnlyMode && document.querySelector(".docx-edit-root"), null, { timeout: 10000 });
  await sleep(500);
  await page.waitForFunction(() => document.querySelector(".dwb-page-margin-guide"), null, { timeout: 10000 });
  const guide = await page.evaluate(() => {
    const g = document.querySelector(".dwb-page-margin-guide");
    const s = document.querySelector("#docCanvas section.docx");
    const cs = getComputedStyle(s);
    return { top: g?.style.top, left: g?.style.left, right: g?.style.right, height: g?.style.height, pad: [cs.paddingTop, cs.paddingLeft, cs.paddingRight, cs.paddingBottom] };
  });
  check("Przy każdej kartce widać dyskretną przerywaną granicę bieżących marginesów", [guide.top, guide.left, guide.right].every((v, i) => Math.abs(parseFloat(v) - parseFloat(guide.pad[i])) < 0.1) && parseFloat(guide.height) > 900, JSON.stringify(guide));

  const undo = async () => { await page.evaluate(() => dwbUndo.undo()); await idle(); };
  const idle = async () => {
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 });
    await sleep(500);
  };
  const zipOf = async () => {
    const b64 = await page.evaluate(async () => {
      const bytes = await buildDocumentForSave();
      let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s);
    });
    return JSZip.loadAsync(Buffer.from(b64, "base64"));
  };
  const sectPrs = async () => ((await (await zipOf()).file("word/document.xml").async("string")).match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g) || []);
  const openLayout = async () => {
    await page.evaluate(() => focusParagraphAtOffset(1, 2));
    await page.click("#pageLayoutBtn");
    await page.waitForSelector('.compose-pop [data-preset="normal"]');
    await sleep(150);
  };

  // ── marginesy: gotowe ──
  await openLayout();
  const menu = await page.evaluate(() => [...document.querySelectorAll(".compose-pop .compose-item")].filter((b) => b.getAttribute("aria-checked") === "true").map((b) => b.dataset.preset || b.dataset.orient || b.dataset.size || b.dataset.v));
  check("„Układ” na pasku: zaznaczone obecne ustawienia (Normalne, pionowa, A4, do góry)", JSON.stringify(menu) === JSON.stringify(["normal", "portrait", "A4", "top"]), JSON.stringify(menu));
  await page.keyboard.press("Escape");
  await page.click("#insertMenuBtn");
  const insertItems = await page.evaluate(() => [...document.querySelectorAll(".compose-pop-insert .compose-item-label")].map((x) => x.textContent));
  check("„＋ Wstaw” bez marginesów i układu strony (to nic nie wstawia)", insertItems.length > 5 && !insertItems.some((x) => /Układ strony|pionie|Marginesy/.test(x)), JSON.stringify(insertItems));
  await page.keyboard.press("Escape");
  await openLayout();
  await page.click('.compose-pop [data-preset="narrow"]');
  await idle();
  let sp = await sectPrs();
  const pad = await page.evaluate(() => getComputedStyle(document.querySelector("#docCanvas section.docx")).paddingLeft);
  check("Wąskie: 1,27 cm w pliku (720 tw) i na kartce (48 px)", /w:top="720" w:right="720" w:bottom="720" w:left="720"/.test(sp[0]) && /w:header="709"/.test(sp[0]) && pad === "48px", `${sp[0]} | ${pad}`);
  const narrowGuide = await page.evaluate(() => {
    const g = document.querySelector(".dwb-page-margin-guide");
    return { top: g?.style.top, left: g?.style.left, right: g?.style.right };
  });
  check("Prowadnica marginesów aktualizuje się po zmianie układu", narrowGuide.top === "48px" && narrowGuide.left === "48px" && narrowGuide.right === "48px", JSON.stringify(narrowGuide));
  check("…Cofnij zdejmuje zmianę marginesów", await page.evaluate(() => dwbUndo._debug().undo.slice(-1)[0] === "undoOpPageSetup"));

  // ── orientacja obraca kartkę i marginesy ──
  await page.evaluate(() => { const zip = null; return zip; });
  await openLayout();
  await page.click('.compose-pop [data-preset="moderate"]');
  await idle();
  await openLayout();
  await page.click('.compose-pop [data-orient="landscape"]');
  await idle();
  sp = await sectPrs();
  const secW = await page.evaluate(() => document.querySelector("#docCanvas section.docx").offsetWidth);
  check("Pozioma: kartka 29,7 × 21, marginesy obrócone (G/D 1,91, L/P 2,54)", /w:w="16838" w:h="11906" w:orient="landscape"/.test(sp[0]) && /w:top="1080" w:right="1440" w:bottom="1080" w:left="1440"/.test(sp[0]) && secW > 1100, `${sp[0]} | ${secW}`);
  await openLayout();
  await page.click('.compose-pop [data-size="A5"]');
  await idle();
  sp = await sectPrs();
  check("Rozmiar A5 przy poziomej: 21 × 14,8 cm", /w:w="11906" w:h="8391" w:orient="landscape"/.test(sp[0]), sp[0]);
  await openLayout();
  await page.click('.compose-pop [data-orient="portrait"]');
  await idle();
  await openLayout();
  await page.click('.compose-pop [data-size="A4"]');
  await idle();

  // ── marginesy niestandardowe ──
  await openLayout();
  await page.click('.compose-pop [data-preset="custom"]');
  await page.waitForSelector(".compose-pop .mf-left");
  await page.fill(".compose-pop .mf-left", "3,5");
  await page.fill(".compose-pop .mf-top", "0,4");
  await sleep(100);
  const warn = await page.evaluate(() => { const w = document.querySelector(".compose-pop .mf-warn"); return w.hidden ? "" : w.textContent; });
  check("niestandardowe: margines 0,4 cm → ostrzeżenie o drukarce", /drukar/.test(warn), warn);
  await page.fill(".compose-pop .mf-right", "19");
  await sleep(100);
  const tooBig = await page.evaluate(() => { const w = document.querySelector(".compose-pop .mf-warn"); return w.classList.contains("is-error") ? w.textContent : ""; });
  check("niestandardowe: za duże marginesy → błąd, bez zmiany", /Za duże/.test(tooBig), tooBig);
  await page.fill(".compose-pop .mf-right", "2,5");
  await page.click(".compose-pop .lf-ok");
  await idle();
  sp = await sectPrs();
  check("niestandardowe zapisane: lewy 3,5 cm (1984 tw), górny 0,4 cm (227 tw)", /w:top="227"/.test(sp[0]) && /w:left="1984"/.test(sp[0]) && /w:right="1417"/.test(sp[0]), sp[0]);

  // ── Podgląd wydruku ──
  const livePages = await page.evaluate(() => document.querySelectorAll("#docCanvas section.docx").length + document.querySelectorAll("#docCanvas .dwb-page-break").length);
  await page.keyboard.press(`${MOD}+KeyP`);
  await page.waitForSelector(".pp-overlay", { timeout: 20000 });
  await sleep(500);
  const pp = await page.evaluate(() => ({
    sheets: document.querySelectorAll(".pp-sheet").length,
    size: [...document.querySelectorAll(".pp-sheet")].map((s) => `${s.offsetWidth}x${Math.round(s.offsetHeight)}`),
    warn: document.querySelector(".pp-warn").hidden ? "" : document.querySelector(".pp-warn").textContent,
    zone: document.querySelector(".pp-overlay").classList.contains("pp-show-zone"),
    srcGone: !document.querySelector(".dwb-pp-src"),
    firstText: document.querySelector(".pp-sheet").textContent.slice(0, 40),
    lastText: [...document.querySelectorAll(".pp-sheet")].pop().textContent,
    liveOk: document.querySelectorAll("#docCanvas .docx-preview-host p").length > 5,
  }));
  check("Ctrl/⌘+P: podgląd — tyle kartek, ile stron w Edycji, A4 (794×1123)", pp.sheets === livePages && pp.size.every((s) => s === "794x1123"), JSON.stringify({ livePages, ...pp }));
  check("podgląd: górny margines 0,4 cm → ostrzeżenie o drukarce i widoczna strefa", /Str\. 1/.test(pp.warn) && pp.zone, pp.warn);
  check("podgląd: treść od nagłówka do ostatniego akapitu, edytor nietknięty", pp.firstText.startsWith("Rozdział pierwszy") && /Eta ostatni akapit/.test(pp.lastText) && pp.srcGone && pp.liveOk, JSON.stringify(pp).slice(0, 300));
  if (ENGINE === "chromium") {
    await page.emulateMedia({ media: "print" });
    const pdf = await page.pdf({ preferCSSPageSize: true });
    await page.emulateMedia({ media: "screen" });
    const pages = (pdf.toString("latin1").match(/\/Type\s*\/Page[^s]/g) || []).length;
    check("druk (PDF): tyle stron, ile kartek w podglądzie", pages === pp.sheets, `${pages} vs ${pp.sheets}`);
  }
  // Marginesy prosto z podglądu (jak w Wordzie przy drukowaniu) — kartki układają się od nowa
  await page.click(".pp-margins");
  await page.waitForSelector('.compose-pop [data-preset="normal"]');
  await sleep(150);
  await page.click('.compose-pop [data-preset="normal"]');
  await page.waitForFunction(() => document.querySelector(".pp-overlay") && !document.querySelector(".pp-warn:not([hidden])"), null, { timeout: 20000 }).catch(() => {});
  await sleep(400);
  const ppAfter = await page.evaluate(() => ({ warn: !document.querySelector(".pp-warn").hidden, pad: getComputedStyle(document.querySelector(".pp-sheet")).paddingTop }));
  check("podgląd → Marginesy → Normalne: kartki od nowa, bez ostrzeżenia (2,5 cm)", !ppAfter.warn && Math.abs(parseFloat(ppAfter.pad) - 94.53) < 0.1, JSON.stringify(ppAfter));
  await page.keyboard.press("Escape");
  await sleep(200);
  check("Esc zamyka podgląd", !(await page.evaluate(() => !!document.querySelector(".pp-overlay"))));
  // menu ⋯ → Marginesy i układ strony (także z trybu Czytanie)
  await page.evaluate(() => appFrame.setReadOnly(true));
  await sleep(300);
  await page.click("#appMenuBtn");
  await page.click("#pageSetupMenuItem");
  await page.waitForSelector('.compose-pop [data-preset="narrow"]', { timeout: 5000 }).catch(() => {});
  check("menu ⋯ → Marginesy i układ strony: okienko z marginesami (z Czytania przełącza na Edycję)", await page.evaluate(() => !!document.querySelector('.compose-pop [data-preset="narrow"]') && !readOnlyMode));
  await page.keyboard.press("Escape");
  await sleep(200);

  // ── odstępy między stronami w Edycji (domyślnie) i „Ukryj biały obszar” ──
  await page.evaluate(() => applyDocumentEdit({ op: "pageSetup", index: 0, scope: "all", margins: { top: 1134, bottom: 1134, left: 1418, right: 1418 } }));
  await idle();
  // dużo tekstu → kilka stron
  await page.evaluate(async () => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const p = ps[1];
    p.textContent = "Długi akapit testowy. ".repeat(30);
    for (let i = 0; i < 40; i++) { focusParagraphAtOffset(1, 0); await handleInlineEnter(collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1], 1, { preventDefault() {}, shiftKey: false }); }
  });
  await page.evaluate(() => waitInlineStructuralIdle());
  await page.evaluate(() => dwbPageBreaks.compute());
  await sleep(600);
  const gapInfo = await page.evaluate(() => {
    const sec = document.querySelector("#docCanvas section.docx");
    const sr = sec.getBoundingClientRect();
    const sc = sr.height / sec.offsetHeight;
    const padT = parseFloat(getComputedStyle(sec).paddingTop);
    const bands = [...sec.querySelectorAll(".dwb-page-gap-band")];
    const firsts = [...sec.querySelectorAll(".dwb-page-gap")].map((g) => {
      const n = g.nextElementSibling;
      const r = document.createRange(); r.selectNodeContents(n);
      const rr = [...r.getClientRects()].find((x) => x.height) || n.getBoundingClientRect();
      return (rr.top - sr.top) / sc;
    });
    return { n: bands.length, padT, diffs: bands.map((b, i) => Math.round(firsts[i] - (parseFloat(b.style.top) + parseFloat(b.style.height)) - padT)), labels: bands.map((b) => b.dataset.label).join(","), paras: collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).length };
  });
  check("odstępy między stronami: przerwa + górny margines (pierwsza linijka strony = góra kartki + margines)", gapInfo.n >= 1 && gapInfo.diffs.every((d) => Math.abs(d) <= 2) && /str\. 2/.test(gapInfo.labels), JSON.stringify(gapInfo));
  const fileParas = await page.evaluate(async () => (await extractParagraphTextsFromDocx(await buildDocumentForSave())).length);
  check("odstępy to tylko wygląd: liczba akapitów podglądu = pliku", fileParas === gapInfo.paras, `${fileParas} vs ${gapInfo.paras}`);
  await page.dblclick(".dwb-page-gap-band", { force: true });
  await sleep(500);
  const hidden = await page.evaluate(() => ({ gaps: document.querySelectorAll(".dwb-page-gap, .dwb-page-gap-band").length, seams: document.querySelectorAll(".dwb-page-break").length, box: document.getElementById("hideWhiteSpace").checked }));
  check("dwuklik w przerwę = Ukryj biały obszar (kreska, opcja w panelu zaznaczona)", hidden.gaps === 0 && hidden.seams >= 1 && hidden.box, JSON.stringify(hidden));
  await page.evaluate(() => dwbPageBreaks.setGaps(true));
  await sleep(300);
  await undo();

  // ── przypisy ──
  await page.evaluate(() => focusParagraphAtOffset(2, 4));
  await page.keyboard.press(`${MOD}+Alt+KeyF`);
  await idle();
  const caretInNote = await page.evaluate(() => !!getSelection().anchorNode?.parentElement?.closest?.("ol.dwb-notes"));
  check("Ctrl/⌘+Alt+F: przypis dolny, kursor w jego tekście", caretInNote);
  await page.keyboard.type("Drugi w tekście.");
  await page.evaluate(() => focusParagraphAtOffset(1, 4));
  await page.click("#insertMenuBtn");
  await page.click(".compose-pop .compose-item:has-text('Przypis dolny')");
  await idle();
  await page.keyboard.type("Pierwszy w tekście.");
  await page.evaluate(() => focusParagraphAtOffset(3, 5));
  await page.keyboard.press(`${MOD}+Alt+KeyD`);
  await idle();
  await page.keyboard.type("Końcowy.");
  await sleep(300);
  const st = () => page.evaluate(() => ({
    refs: [...document.querySelectorAll("article sup[data-dwb-note]")].map((s) => s.textContent).join(","),
    notes: [...document.querySelectorAll("ol.dwb-notes > li")].map((li) => li.textContent.trim()).join(" | "),
  }));
  let s = await st();
  check("numeracja jak w Wordzie: wg kolejności w tekście (1, 2) + końcowy „i”", s.refs === "1,2,i" && s.notes === "1 Pierwszy w tekście. | 2 Drugi w tekście. | i Końcowy.", JSON.stringify(s));
  let zip = await zipOf();
  const fx = await zip.file("word/footnotes.xml")?.async("string");
  const ct = await zip.file("[Content_Types].xml").async("string");
  const rels = await zip.file("word/_rels/document.xml.rels").async("string");
  check("plik bez przypisów: footnotes.xml/endnotes.xml z separatorami, typy i powiązania, style", /w:type="separator"/.test(fx || "") && /footnotes\+xml/.test(ct) && /endnotes\+xml/.test(ct) && /relationships\/footnotes"/.test(rels)
    && /Pierwszy w tekście/.test(fx) && /footnote text/.test(await zip.file("word/styles.xml").async("string")), (fx || "").slice(0, 200));
  // usunięcie odnośnika (Backspace za nim) → przypis znika, numery dalej
  await page.evaluate(() => {
    const isl = document.querySelector("article sup[data-dwb-note]").closest("[data-cm-kind=note]");
    const r = document.createRange(); r.setStartAfter(isl); r.collapse(true);
    docEditRoot().focus(); getSelection().removeAllRanges(); getSelection().addRange(r);
  });
  await page.keyboard.press("Backspace");
  await sleep(700);
  await idle();
  s = await st();
  zip = await zipOf();
  const fx2 = await zip.file("word/footnotes.xml").async("string");
  check("usunięty odnośnik: przypis znika (podgląd i plik), numery przesunięte", s.refs === "1,i" && s.notes === "1 Drugi w tekście. | i Końcowy." && !/Pierwszy w tekście/.test(fx2) && /Drugi w tekście/.test(fx2), JSON.stringify(s));
  await page.evaluate(() => dwbUndo.undo());
  await idle();
  s = await st();
  check("Cofnij przywraca odnośnik i przypis", s.refs === "1,2,i", JSON.stringify(s));

  // ── Widok mobilny: podgląd wydruku dalej ze stronami A4 ──
  await page.setViewportSize({ width: 390, height: 800 });
  await page.evaluate(() => typeof setViewLayoutPref === "function" && setViewLayoutPref("mobile"));
  await sleep(1200);
  const reflow = await page.evaluate(() => typeof shouldUseMobileReflow === "function" && shouldUseMobileReflow());
  await page.evaluate(() => dwbPrint.open());
  await page.waitForSelector(".pp-overlay", { timeout: 20000 });
  await sleep(400);
  const mob = await page.evaluate(() => ({ w: document.querySelector(".pp-sheet").offsetWidth, zoom: getComputedStyle(document.querySelector(".pp-pages")).getPropertyValue("--pp-zoom"), scrollW: document.querySelector(".pp-scroll").scrollWidth - document.querySelector(".pp-scroll").clientWidth }));
  check("Widok mobilny: podgląd wydruku ze stroną A4 dopasowaną do ekranu", reflow && mob.w === 794 && parseFloat(mob.zoom) < 0.5 && mob.scrollW <= 1, JSON.stringify({ reflow, ...mob }));
  await page.screenshot({ path: process.env.SHOT || "/dev/null" }).catch(() => {});
  await page.click(".pp-close");

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  let fail = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail})`}`);
    if (!r.ok) fail++;
  }
  console.log(`\n[${ENGINE}] ${fail ? `${fail} z ${results.length} nie przeszło` : `${results.length}/${results.length} OK`}`);
  process.exit(fail ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
