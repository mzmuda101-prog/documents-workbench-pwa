// image-wrap-edge-playwright.js — zawijanie i karta obrazu w nietypowych miejscach (przegląd błędów
// 2026-11-14): dwa obrazy w jednym akapicie (zmienia się ten kliknięty), pole tekstowe przed obrazem
// w akapicie (nietknięte), rozmiar i wyrównanie obrazu pływającego, Cofnij / Ponów, zapis → otwarcie,
// usuwanie pływającego, obraz w komórce tabeli, obraz z pliku Worda za tekstem (z PDF), tryb Czytanie
// (bez „Zaznacz obraz”), telefon (karta i menu w ekranie). Zawsze: podgląd = plik (liczba obrazów).
//   node scripts/image-wrap-edge-playwright.js        ENGINE=webkit node scripts/image-wrap-edge-playwright.js
const fs = require("fs");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 }).then(() => page.waitForTimeout(500));

async function newPage(browser, opts = {}) {
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 }, ...opts });
  await ctx.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
  const page = await ctx.newPage();
  page.errors = [];
  page.on("pageerror", (e) => page.errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  return page;
}
const makeImg = (page, color, w = 300, h = 200) => page.evaluateHandle(async ([color, w, h]) => {
  const c = document.createElement("canvas"); c.width = w; c.height = h;
  const g = c.getContext("2d"); g.fillStyle = color; g.fillRect(0, 0, w, h);
  const blob = await new Promise((r) => c.toBlob(r, "image/png"));
  return new File([blob], `${color.slice(1)}.png`, { type: "image/png" });
}, [color, w, h]);
// obrazy akapitu w pliku: [rodzaj (inline / anchor + zawijanie), szerokość cx]
const fileImgs = (page) => page.evaluate(async () => {
  const xml = await getDocumentXmlString(await buildDocumentForSave());
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  return collectParagraphElements(doc.documentElement, "all").map((p) => composeImageRuns(p).map((r) => {
    const box = [...r.getElementsByTagName("*")].find((n) => n.localName === "inline" || n.localName === "anchor");
    const wrap = box ? [...box.childNodes].find((n) => /^wrap/.test(n.localName || ""))?.localName || "" : "";
    const ext = box ? [...box.childNodes].find((n) => n.localName === "extent") : null;
    const ph = box ? [...box.childNodes].find((n) => n.localName === "positionH") : null;
    return `${box?.localName}${wrap ? `:${wrap}` : ""}${box?.getAttribute("behindDoc") === "1" ? ":behind" : ""}|${ext?.getAttribute("cx")}${ph ? `|${ph.textContent}` : ""}`;
  })).filter((x) => x.length);
});
const viewIsFile = (page) => page.evaluate(async () => {
  const xml = await getDocumentXmlString(await buildDocumentForSave());
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const inFile = collectParagraphElements(doc.documentElement, "all").reduce((n, p) => n + composeImageRuns(p).length, 0);
  const inView = docImageTargets(document.querySelector(".docx-preview-host")).length;
  return inFile === inView;
});
// karta obrazu: klik w obraz, a gdy leży pod tekstem — prawy przycisk → „Zaznacz obraz”
async function openCard(page, sel = ".docx-preview-host img", n = 0) {
  await page.evaluate(() => composeUi.hideImageCard()); // otwarta karta mogła leżeć nad drugim obrazem
  const box = await page.evaluate(([sel, n]) => { const i = document.querySelectorAll(sel)[n]; i.scrollIntoView({ block: "center" }); const r = i.getBoundingClientRect(); return { x: r.left + Math.min(r.width / 2, 20), y: r.top + Math.min(r.height / 2, 20), onTop: document.elementFromPoint(r.left + Math.min(r.width / 2, 20), r.top + Math.min(r.height / 2, 20)) === i }; }, [sel, n]);
  if (box.onTop) await page.mouse.click(box.x, box.y);
  else {
    // pod tekstem: prawy przycisk w środku obrazu (tam, gdzie leży tekst, nie na marginesie kartki)
    const mid = await page.evaluate(([sel, n]) => { const r = document.querySelectorAll(sel)[n].getBoundingClientRect(); return { x: Math.min(r.left + r.width / 2, innerWidth - 40), y: Math.min(Math.max(r.top + r.height / 2, 120), innerHeight - 60) }; }, [sel, n]);
    await page.mouse.click(mid.x, mid.y, { button: "right" });
    await page.waitForSelector(".doc-menu");
    await page.locator(".doc-menu button", { hasText: /Zaznacz obraz/ }).click();
  }
  await page.waitForSelector(".image-card");
  await page.waitForTimeout(250);
}
async function wrap(page, mode) {
  await page.click(".image-card .ic-wrap");
  await page.waitForSelector(".compose-pop-wrap");
  await page.click(`.compose-pop-wrap [data-wrap="${mode}"]`);
  await idle(page);
}
// akapity pliku przebudowane w przeglądarce (XML), potem przerysowanie — jak dokument z Worda
// (strona ma CSP bez eval — dwie stałe przebudowy: „twoImages”, „textboxFirst”)
const rebuild = (page, what) => page.evaluate(async (what) => {
  const zip = await JSZip.loadAsync(originalFileBytes);
  const doc = new DOMParser().parseFromString(await zip.file("word/document.xml").async("string"), "application/xml");
  const ps = collectParagraphElements(doc.documentElement, "all");
  if (what === "twoImages") {
    const withImg = ps.filter((p) => composeImageRuns(p).length);
    withImg[0].appendChild(composeImageRuns(withImg[1])[0]);
    withImg[1].parentNode.removeChild(withImg[1]);
  } else {
    const tb = [...doc.getElementsByTagNameNS(W_NS, "r")].find((r) => r.getElementsByTagNameNS(W_NS, "txbxContent").length);
    const imgP = ps.find((p) => composeImageRuns(p).length);
    if (tb && imgP) imgP.insertBefore(tb, imgP.firstChild);
  }
  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
  await reloadFromBytes(new Uint8Array(await zip.generateAsync({ type: "uint8array" })));
}, what);

(async () => {
  const browser = await pw[ENGINE].launch({ headless: true });

  // ── A. dwa obrazy w jednym akapicie + pole tekstowe przed obrazem ─────────────
  {
    const page = await newPage(browser);
    await page.evaluate(() => composeUi.createNew("blank"));
    await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
    await idle(page);
    await page.keyboard.type("Pierwszy akapit z tekstem.");
    await page.evaluate((f) => composeUi.insertImageFile(f), await makeImg(page, "#c0392b"));
    await idle(page);
    await page.evaluate((f) => composeUi.insertImageFile(f), await makeImg(page, "#2e86c1"));
    await idle(page);
    for (let i = 0; i < 3; i++) { await page.keyboard.type(`Tekst dalszy ${i + 1}, który może opływać obraz z boku i ma kilka słów.`); await page.keyboard.press("Enter"); }
    await idle(page);
    // oba obrazy do jednego akapitu (drugi run do akapitu pierwszego)
    await rebuild(page, "twoImages");
    await idle(page);
    await page.evaluate(() => appFrame.setReadOnly(false));
    await idle(page);
    let f = await fileImgs(page);
    check("A0: dwa obrazy w jednym akapicie (plik)", f.length === 1 && f[0].length === 2, JSON.stringify(f));
    await openCard(page, ".docx-preview-host img", 1);
    await wrap(page, "square");
    f = await fileImgs(page);
    check("A1: zawijanie na DRUGIM obrazie — zmienia się drugi, pierwszy w linii", /^inline/.test(f[0][0]) && /^anchor:wrapSquare/.test(f[0][1]), JSON.stringify(f));
    // rozmiar pierwszego (karta na pierwszym) — drugi nietknięty
    const before = f[0][1];
    await openCard(page, ".docx-preview-host img", 0);
    await page.evaluate(() => { const r = document.querySelector(".image-card input[type=range]"); r.value = "30"; r.dispatchEvent(new Event("input", { bubbles: true })); r.dispatchEvent(new Event("change", { bubbles: true })); r.blur(); });
    await page.waitForTimeout(700);
    await idle(page);
    f = await fileImgs(page);
    check("A2: rozmiar pierwszego obrazu — drugi (pływający) bez zmian", f[0][1] === before && f[0][0] !== `inline|${before.split("|")[1]}`, JSON.stringify(f));
    check("A3: podgląd = plik (liczba obrazów)", await viewIsFile(page));
    check("A3b: po zawijaniu akapit nie jest „zmieniony” (zapis nie przestawia obrazów)", (await page.evaluate(() => collectInlineParagraphEdits().length)) === 0, String(await page.evaluate(() => collectInlineParagraphEdits().map((e) => e.index))));
    // pole tekstowe przed obrazem w tym samym akapicie
    await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0]; p.closest(".docx-edit-root").focus(); placeCaret(p, 3); });
    await page.evaluate(() => dwbTextbox.insert());
    await idle(page);
    await rebuild(page, "textboxFirst");
    await idle(page);
    const tbBefore = await page.evaluate(async () => { const x = await getDocumentXmlString(originalFileBytes); return (x.match(/<wp:anchor[^>]*>(?:(?!<\/wp:anchor>)[\s\S])*?txbxContent[\s\S]*?<\/wp:anchor>/) || [""])[0].replace(/<w:txbxContent>[\s\S]*<\/w:txbxContent>/, ""); });
    await openCard(page, ".docx-preview-host img", 0);
    await wrap(page, "behind");
    const tbAfter = await page.evaluate(async () => { const x = await getDocumentXmlString(originalFileBytes); return (x.match(/<wp:anchor[^>]*>(?:(?!<\/wp:anchor>)[\s\S])*?txbxContent[\s\S]*?<\/wp:anchor>/) || [""])[0].replace(/<w:txbxContent>[\s\S]*<\/w:txbxContent>/, ""); });
    f = await fileImgs(page);
    check("A4: pole tekstowe przed obrazem — zawijanie trafia w obraz, pole nietknięte", tbBefore && tbBefore === tbAfter && /^anchor:wrapNone:behind/.test(f[0][0]), JSON.stringify({ f, same: tbBefore === tbAfter, tb: !!tbBefore }));
    // wyrównanie obrazu pływającego → w kolumnie
    await openCard(page, ".docx-preview-host img", 1);
    await page.click('.image-card [data-align="right"]');
    await idle(page);
    f = await fileImgs(page);
    check("A5: wyrównanie do prawej obrazu pływającego = positionH align right", /\|right$/.test(f[0][1]), JSON.stringify(f));
    // Cofnij / Ponów
    await page.evaluate(() => document.getElementById("undoBtn").click());
    await idle(page);
    const u = await fileImgs(page);
    await page.evaluate(() => document.getElementById("redoBtn").click());
    await idle(page);
    const r = await fileImgs(page);
    check("A6: Cofnij / Ponów wyrównania obrazu pływającego", !/\|right$/.test(u[0][1]) && /\|right$/.test(r[0][1]), JSON.stringify({ u, r }));
    // zapis → otwarcie: tryby zostają, karta pokazuje bieżący
    const bytes = await page.evaluate(async () => { const b = await buildDocumentForSave(); let s = ""; b.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
    await page.locator("#fileInput").setInputFiles({ name: "po-zapisie.docx", mimeType: DOCX_MIME, buffer: Buffer.from(bytes, "base64") });
    await idle(page);
    await page.evaluate(() => appFrame.setReadOnly(false));
    await idle(page);
    const re = await fileImgs(page);
    check("A7: zapis → otwarcie: te same obrazy i tryby", JSON.stringify(re) === JSON.stringify(r), JSON.stringify({ re, r }));
    await openCard(page, ".docx-preview-host img", 1);
    await page.click(".image-card .ic-wrap");
    await page.waitForSelector(".compose-pop-wrap");
    check("A8: po otwarciu karta pokazuje bieżący tryb (przyległe)", await page.evaluate(() => document.querySelector('.compose-pop-wrap [data-wrap="square"]')?.getAttribute("aria-checked") === "true"));
    await page.keyboard.press("Escape");
    // usuwanie obrazu pływającego z karty
    await openCard(page, ".docx-preview-host img", 1);
    await page.click(".image-card .ic-del");
    await idle(page);
    const d = await fileImgs(page);
    check("A9: usunięcie obrazu pływającego — zostaje jeden obraz, podgląd = plik", d.flat().length === 1 && await viewIsFile(page), JSON.stringify(d));
    // tryb Czytanie: bez „Zaznacz obraz”
    await page.evaluate(() => appFrame.setReadOnly(true));
    await idle(page);
    const pt = await page.evaluate(() => { const i = document.querySelector(".docx-preview-host img"); const b = i.getBoundingClientRect(); return { x: b.left + 5, y: b.top + 5 }; });
    await page.mouse.click(pt.x, pt.y, { button: "right" });
    await page.waitForTimeout(300);
    const items = await page.evaluate(() => [...document.querySelectorAll(".doc-menu button")].map((b) => b.textContent));
    check("A10: Czytanie — w menu bez „Zaznacz obraz”", !items.some((x) => /Zaznacz obraz/.test(x)), items.join(" | "));
    check("A: bez błędów strony", !page.errors.length, page.errors.join(" | "));
    await page.context().close();
  }

  // ── B. obraz w komórce tabeli ───────────────────────────────────────────────
  {
    const page = await newPage(browser);
    await page.evaluate(() => composeUi.createNew("blank"));
    await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
    await idle(page);
    await page.evaluate(() => composeUi.insertTable(2, 2));
    await idle(page);
    await page.evaluate((f) => { const td = document.querySelector(".docx-preview-host td p"); td.closest(".docx-edit-root").focus(); placeCaret(td, 0); return composeUi.insertImageFile(f, td); }, await makeImg(page, "#27ae60", 200, 120));
    await idle(page);
    const inCell = await page.evaluate(() => !!document.querySelector(".docx-preview-host td img"));
    if (inCell) {
      await openCard(page, ".docx-preview-host td img", 0);
      await wrap(page, "square");
      const f = await fileImgs(page);
      check("B1: obraz w komórce — przyległe zapisane, podgląd = plik", f.flat().some((x) => /^anchor:wrapSquare/.test(x)) && await viewIsFile(page), JSON.stringify(f));
      await openCard(page, ".docx-preview-host td img", 0);
      await wrap(page, "inline");
      check("B2: …i z powrotem w linii", (await fileImgs(page)).flat().every((x) => /^inline/.test(x)), JSON.stringify(await fileImgs(page)));
    } else check("B0: obraz w komórce wstawiony", false, "brak obrazu w komórce");
    check("B: bez błędów strony", !page.errors.length, page.errors.join(" | "));
    await page.context().close();
  }

  // ── C. obraz z pliku Worda za tekstem: tło zakotwiczone do strony (skan z PDF) ────────
  // Podgląd rysuje je poza akapitami — karta obrazu go nie obsługuje (znane ograniczenie). Menu
  // pod prawym przyciskiem nie może proponować „Zaznacz obraz”, które nic nie zrobi.
  {
    const file = "output/korpus/Lekcja 06.11.2026r..docx";
    if (fs.existsSync(file)) {
      const page = await newPage(browser);
      await page.locator("#fileInput").setInputFiles(file);
      await idle(page);
      await page.evaluate(() => appFrame.setReadOnly(false));
      await idle(page);
      const pt = await page.evaluate(() => { const i = document.querySelector(".docx-preview-host img"); i.scrollIntoView({ block: "center" }); const r = i.getBoundingClientRect(); return { x: r.left + r.width / 2, y: Math.min(r.top + r.height / 2, innerHeight - 80), inP: !!i.closest("p") }; });
      await page.mouse.click(pt.x, pt.y, { button: "right" });
      await page.waitForTimeout(400);
      const items = await page.evaluate(() => [...document.querySelectorAll(".doc-menu button")].map((b) => b.textContent));
      check("C1: tło strony z PDF (poza akapitem): menu bez martwego „Zaznacz obraz”", pt.inP || !items.some((x) => /Zaznacz obraz/.test(x)), JSON.stringify({ inP: pt.inP, items: items.slice(0, 4) }));
      check("C: bez błędów strony", !page.errors.length, page.errors.join(" | "));
      await page.context().close();
    }
  }

  // ── D. telefon: karta i menu zawijania w ekranie ────────────────────────────
  {
    const page = await newPage(browser, { viewport: { width: 390, height: 844 }, hasTouch: true, ...(ENGINE === "chromium" ? { isMobile: true } : {}) });
    await page.evaluate(() => composeUi.createNew("blank"));
    await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
    await idle(page);
    await page.keyboard.type("Tekst.");
    await page.evaluate((f) => composeUi.insertImageFile(f), await makeImg(page, "#8e44ad"));
    await idle(page);
    await page.evaluate(() => { const i = document.querySelector(".docx-preview-host img"); i.scrollIntoView({ block: "center" }); });
    await page.evaluate(() => composeUi.showImageCard(document.querySelector(".docx-preview-host img")));
    await page.waitForSelector(".image-card");
    await page.waitForTimeout(300);
    await page.evaluate(() => document.querySelector(".image-card .ic-wrap").click());
    await page.waitForSelector(".compose-pop-wrap");
    const fit = await page.evaluate(() => [".image-card", ".compose-pop-wrap"].map((s) => { const r = document.querySelector(s).getBoundingClientRect(); return r.left >= 0 && r.right <= innerWidth + 0.5 && r.top >= 0 && r.bottom <= innerHeight + 0.5; }));
    check("D1: telefon — karta obrazu i menu zawijania w ekranie", fit.every(Boolean), JSON.stringify(fit));
    check("D: bez błędów strony", !page.errors.length, page.errors.join(" | "));
    await page.context().close();
  }

  await browser.close();
  report();
})().catch((e) => { check("przebieg bez wyjątku", false, String(e?.stack || e).slice(0, 500)); report(); });
function report() {
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
}
