// compose3-playwright.js — tworzenie dokumentu, Etap 3: tabele i obrazy.
//
// Jak użytkownik: „＋ Wstaw → Tabela” (siatka rozmiaru), Tab po komórkach (w ostatniej — nowy
// wiersz), przycisk „Tabela” na pasku TYLKO w tabeli (wiersze/kolumny/usuwanie), Backspace nie
// skleja komórek; „＋ Wstaw → Obraz” (plik), wklejenie obrazu, karta obrazu (suwak rozmiaru,
// wyrównanie, tekst alternatywny, usuń). Na końcu ZAPISANY plik + telefon z dotykiem.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(350));
const tables = (page) => page.evaluate(() => [...document.querySelectorAll(".docx-preview-host table")].map((t) => [...t.rows].map((r) => [...r.cells].map((c) => c.textContent).join("|")).join(" / ")));
const caret = (page) => page.evaluate(() => {
  const q = document.activeElement?.closest?.(".docx-editable-p");
  if (!q) return null;
  const td = q.closest("td");
  return { text: q.textContent, inTable: !!td, r: td ? [...td.closest("table").rows].indexOf(td.parentElement) : -1, c: td ? [...td.parentElement.cells].indexOf(td) : -1 };
});
const savedZip = async (page) => {
  const b64 = await page.evaluate(async () => {
    const bytes = await buildDocumentForSave();
    let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s);
  });
  return JSZip.loadAsync(Buffer.from(b64, "base64"));
};
const tableTool = async (page, label) => {
  await page.click("#tableToolsBtn");
  await page.click(`.compose-pop-tabletools .compose-item-label:text-is("${label}")`);
  await idle(page);
};
// obraz z canvasa (bez plików na dysku)
const makeImage = (page, w, h, type = "image/png") => page.evaluateHandle(async ([w, h, type]) => {
  const cv = document.createElement("canvas"); cv.width = w; cv.height = h;
  const g = cv.getContext("2d"); g.fillStyle = "#2f6fdf"; g.fillRect(0, 0, w, h); g.fillStyle = "#fff"; g.fillRect(w / 4, h / 4, w / 2, h / 2);
  const blob = await new Promise((r) => cv.toBlob(r, type, 0.9));
  return new File([blob], `test.${type.split("/")[1]}`, { type });
}, [w, h, type]);

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

  // ── tabela ─────────────────────────────────────────────────────────────────
  await page.keyboard.type("Tabela poniżej:");
  check("poza tabelą przycisku „Tabela” nie ma", await page.evaluate(() => document.getElementById("tableToolsBtn").hidden));
  await page.click("#insertMenuBtn");
  await page.click(".compose-pop-insert .compose-item-label:text-is('Tabela')");
  await page.hover(".table-pick button[data-r='2'][data-c='3']");
  const label = await page.textContent(".table-pick-size");
  check("siatka: najechanie pokazuje rozmiar „3 × 2”", label === "3 × 2", label);
  await page.click(".table-pick button[data-r='2'][data-c='3']");
  await idle(page);
  let c = await caret(page);
  check("tabela 3×2 pod akapitem, kursor w pierwszej komórce", JSON.stringify(await tables(page)) === JSON.stringify(["|| / ||"]) && c?.inTable && c.r === 0 && c.c === 0, JSON.stringify({ t: await tables(page), c }));
  check("kursor w tabeli → na pasku przycisk „Tabela”", await page.evaluate(() => !document.getElementById("tableToolsBtn").hidden));
  for (const w of ["A1", "B1", "C1", "A2", "B2", "C2"]) { await page.keyboard.type(w); await page.keyboard.press("Tab"); await page.waitForTimeout(120); }
  await idle(page);
  c = await caret(page);
  check("Tab po komórkach, Tab w ostatniej = nowy wiersz z kursorem w 1. komórce", JSON.stringify(await tables(page)) === JSON.stringify(["A1|B1|C1 / A2|B2|C2 / ||"]) && c?.r === 2 && c.c === 0, JSON.stringify({ t: await tables(page), c }));
  await page.keyboard.press("Shift+Tab");
  c = await caret(page);
  check("Shift+Tab = poprzednia komórka (koniec tekstu)", c?.r === 1 && c.c === 2 && c.text === "C2", JSON.stringify(c));
  await page.evaluate(() => { const td = document.querySelector(".docx-preview-host table").rows[1].cells[1]; const p = td.querySelector("p"); focusParagraphAtOffset(resolveParaIndex(p), 0); });
  await page.keyboard.press("Backspace");
  await page.waitForTimeout(250);
  check("Backspace na początku komórki nie skleja jej z sąsiednią", JSON.stringify(await tables(page)) === JSON.stringify(["A1|B1|C1 / A2|B2|C2 / ||"]), JSON.stringify(await tables(page)));
  await tableTool(page, "Kolumna z prawej");
  c = await caret(page);
  check("„Kolumna z prawej”: nowa kolumna za bieżącą, kursor w niej", JSON.stringify(await tables(page)) === JSON.stringify(["A1|B1||C1 / A2|B2||C2 / |||"]) && c?.r === 1 && c.c === 2, JSON.stringify({ t: await tables(page), c }));
  await page.keyboard.type("X");
  await tableTool(page, "Wiersz powyżej");
  c = await caret(page);
  check("„Wiersz powyżej”: pusty wiersz nad bieżącym, kursor w nim", JSON.stringify(await tables(page)) === JSON.stringify(["A1|B1||C1 / ||| / A2|B2|X|C2 / |||"]) && c?.r === 1, JSON.stringify({ t: await tables(page), c }));
  await tableTool(page, "Usuń wiersz");
  await page.evaluate(() => { const td = document.querySelector(".docx-preview-host table").rows[1].cells[2]; focusParagraphAtOffset(resolveParaIndex(td.querySelector("p")), 0); });
  await tableTool(page, "Usuń kolumnę");
  check("„Usuń wiersz”, „Usuń kolumnę”", JSON.stringify(await tables(page)) === JSON.stringify(["A1|B1|C1 / A2|B2|C2 / ||"]), JSON.stringify(await tables(page)));
  const undoLabel = await page.getAttribute("#undoBtn", "aria-label");
  check("Cofnij nazywa krok „tabela”", /tabela/.test(undoLabel || ""), undoLabel);
  await page.click("#undoBtn");
  await idle(page);
  check("Cofnij przywraca kolumnę", JSON.stringify(await tables(page)) === JSON.stringify(["A1|B1|X|C1 / A2|B2||C2 / |||"]) || (await tables(page))[0].split(" / ")[0].split("|").length === 4, JSON.stringify(await tables(page)));
  let zip = await savedZip(page);
  let xml = await zip.file("word/document.xml").async("string");
  const styles = await zip.file("word/styles.xml").async("string");
  check("plik: tabela ze stylem „Table Grid”, siatka 4 kolumn, akapit po tabeli", /<w:tblStyle w:val="TableGrid"\/>/.test(styles + xml) && /w:name w:val="Table Grid"/.test(styles) && (xml.match(/<w:gridCol /g) || []).length === 4 && /<\/w:tbl><w:p/.test(xml), "");
  await page.evaluate(() => { const td = document.querySelector(".docx-preview-host table").rows[0].cells[0]; focusParagraphAtOffset(resolveParaIndex(td.querySelector("p")), 0); });
  await tableTool(page, "Usuń tabelę");
  c = await caret(page);
  check("„Usuń tabelę”: tabeli nie ma, kursor poza tabelą, przycisk znika", (await tables(page)).length === 0 && c && !c.inTable && await page.evaluate(() => document.getElementById("tableToolsBtn").hidden), JSON.stringify(c));

  // ── obraz ──────────────────────────────────────────────────────────────────
  await page.evaluate(() => { const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")); focusParagraphAtOffset(0, ps[0].textContent.length); });
  const big = await makeImage(page, 3000, 1500, "image/jpeg");
  await page.evaluate((f) => composeUi.insertImageFile(f), big);
  await idle(page);
  const img = await page.evaluate(() => {
    const i = document.querySelector(".docx-preview-host img");
    const p = i?.closest("p");
    return i && { w: i.getBoundingClientRect().width, pw: p.clientWidth, align: getComputedStyle(p).textAlign, idx: resolveParaIndex(p), caret: resolveParaIndex(document.activeElement?.closest?.(".docx-editable-p")) };
  });
  check("obraz pod akapitem: na szerokość tekstu (większy zmniejszony), wyśrodkowany, kursor pod nim", img && Math.abs(img.w - img.pw) < 4 && img.align === "center" && img.caret === img.idx + 1, JSON.stringify(img));
  zip = await savedZip(page);
  const media = Object.keys(zip.files).filter((f) => /^word\/media\/.+/.test(f));
  const mediaSize = media.length ? (await zip.file(media[0]).async("uint8array")).length : 0;
  const imgDims = await page.evaluate(() => { const i = document.querySelector(".docx-preview-host img"); return [i.naturalWidth, i.naturalHeight]; });
  check("zdjęcie 3000 px zmniejszone do 2400 px (JPEG), plik w word/media", media.length === 1 && /\.jpe?g$/.test(media[0]) && imgDims[0] === 2400 && mediaSize > 0, JSON.stringify({ media, imgDims, mediaSize }));
  check("akapit z obrazem bez ramki „tylko do odczytu”", await page.evaluate(() => getComputedStyle(document.querySelector(".docx-preview-host img").closest("p")).outlineColor === "rgba(0, 0, 0, 0)"));
  await page.click(".docx-preview-host img");
  const card = await page.evaluate(() => ({ open: !!document.querySelector(".image-card"), val: document.querySelector(".image-size-val")?.textContent, sel: document.querySelector(".docx-preview-host img").classList.contains("img-selected") }));
  check("klik w obraz → karta obrazu (100%), obraz zaznaczony", card.open && card.val === "100%" && card.sel, JSON.stringify(card));
  await page.evaluate(() => { const r = document.querySelector(".image-card input"); r.value = "40"; r.dispatchEvent(new Event("input")); r.dispatchEvent(new Event("change")); });
  await idle(page);
  const sized = await page.evaluate(() => { const i = document.querySelector(".docx-preview-host img"); return { r: i.getBoundingClientRect().width / i.closest("p").clientWidth, val: document.querySelector(".image-size-val")?.textContent }; });
  check("suwak 40% → obraz 40% szerokości tekstu, karta zostaje", Math.abs(sized.r - 0.4) < 0.03 && sized.val === "40%", JSON.stringify(sized));
  await page.click(".image-card [data-align='left']");
  await idle(page);
  check("wyrównanie obrazu do lewej", await page.evaluate(() => getComputedStyle(document.querySelector(".docx-preview-host img").closest("p")).textAlign) === "left");
  await page.click(".image-card .ic-alt");
  await page.fill(".compose-pop-form textarea", "Niebieski prostokąt");
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
  zip = await savedZip(page);
  xml = await zip.file("word/document.xml").async("string");
  check("tekst alternatywny w pliku (docPr descr)", /<wp:docPr [^>]*descr="Niebieski prostokąt"/.test(xml), (xml.match(/<wp:docPr [^>]*>/) || [""])[0]);
  check("plik: powiązanie obrazu + typ JPEG w [Content_Types]", /relationships\/image" Target="media\/image1\.jpe?g"/.test(await zip.file("word/_rels/document.xml.rels").async("string")) && /Extension="jpe?g"/.test(await zip.file("[Content_Types].xml").async("string")), "");
  // wklejenie
  const lastIdx = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).length - 1);
  await page.evaluate((i) => focusParagraphAtOffset(i, 0), lastIdx);
  await page.keyboard.type("Zrzut:");
  const shot = await makeImage(page, 400, 200);
  await page.evaluate((f) => {
    const dt = new DataTransfer(); dt.items.add(f);
    document.activeElement.closest(".docx-editable-p").dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  }, shot);
  await page.waitForTimeout(400);
  await idle(page);
  check("wklejony obraz (zrzut ekranu) trafia do dokumentu", await page.evaluate(() => document.querySelectorAll(".docx-preview-host img").length) === 2);
  await page.click(".docx-preview-host img >> nth=1");
  await page.keyboard.press("Delete");
  await idle(page);
  check("Delete przy zaznaczonym obrazie usuwa go", await page.evaluate(() => document.querySelectorAll(".docx-preview-host img").length) === 1);
  await page.evaluate(() => appFrame.setReadOnly(true));
  await page.click(".docx-preview-host img");
  check("Czytanie: klik w obraz nie pokazuje karty", !(await page.$(".image-card")));
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);

  // ── telefon z dotykiem ─────────────────────────────────────────────────────
  const touch = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 800 }, isMobile: true, hasTouch: true });
  await touch.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const tp = await touch.newPage();
  tp.on("pageerror", (e) => errors.push(`touch: ${e.message}`));
  await tp.goto(APP_URL, { waitUntil: "load" });
  await tp.evaluate(() => document.getElementById("heroSplash")?.remove());
  await tp.evaluate(() => composeUi.createNew("blank"));
  await tp.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(tp);
  await tp.evaluate(() => focusParagraphAtOffset(0, 0));
  await tp.tap("#insertMenuBtn");
  await tp.tap(".compose-pop-insert .compose-item-label:text-is('Tabela')");
  const pick = await tp.evaluate(() => { const r = document.querySelector(".compose-pop-table").getBoundingClientRect(); return { r: r.right, w: innerWidth, btn: document.querySelector(".table-pick button").getBoundingClientRect().width }; });
  check("telefon: siatka tabeli mieści się, pola ≥ 30 px do palca", pick.r <= pick.w && pick.btn >= 30, JSON.stringify(pick));
  await tp.tap(".table-pick button[data-r='2'][data-c='2']");
  await idle(tp);
  check("telefon: tap w siatkę wstawia tabelę 2×2", JSON.stringify(await tables(tp)) === JSON.stringify(["| / |"]), JSON.stringify(await tables(tp)));
  await tp.tap("#tableToolsBtn");
  await tp.tap(".compose-pop-tabletools .compose-item-label:text-is('Wiersz poniżej')");
  await idle(tp);
  check("telefon: tap „Tabela → Wiersz poniżej”", (await tables(tp))[0].split(" / ").length === 3, JSON.stringify(await tables(tp)));
  const f2 = await makeImage(tp, 800, 400);
  await tp.evaluate((f) => { focusParagraphAtOffset(collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).length - 1, 0); return composeUi.insertImageFile(f); }, f2);
  await idle(tp);
  await tp.tap(".docx-preview-host img");
  const tcard = await tp.evaluate(() => { const c = document.querySelector(".image-card"); if (!c) return null; const r = c.getBoundingClientRect(); return { l: r.left, r: r.right, w: innerWidth, b: r.bottom, h: innerHeight }; });
  check("telefon: tap w obraz → karta mieści się na ekranie", tcard && tcard.l >= 0 && tcard.r <= tcard.w && tcard.b <= tcard.h, JSON.stringify(tcard));
  await tp.tap(".image-card [data-align='right']");
  await idle(tp);
  check("telefon: tap wyrównania obrazu działa", await tp.evaluate(() => getComputedStyle(document.querySelector(".docx-preview-host img").closest("p")).textAlign) === "right");
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
