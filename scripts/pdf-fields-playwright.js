// pdf-fields-playwright.js — pola formularza z PDF jako pola Worda (2026-10-28, etap 3 edycji kart).
//
// Karty terenu: „Miejscowość: ……”, „Teren nr: ……” i lista klatek to pola formularza PDF (AcroForm).
// Konwersja robi z nich kontrolki tekstowe Worda (w:sdt), tytuł z napisu przed polem; kropki
// zostają tabulatorem, puste pole = pusta kontrolka między kropkami (wydruk jak w PDF). W apce:
// pola rozpoznane (panel Formularz), klik → wpis, Tab → następne pole. Przełącznik w panelu Widok
// wyłącza to (dawny wynik: zwykły tekst). PDF budowany ręcznie (Helvetica, pola bez wyglądu,
// NeedAppearances — jak karty z Ghostscripta). ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const path = require("path");
const fs = require("fs");
const os = require("os");
const { APP_URL } = require("./docx-test-helpers");
const ooxml = require("./ooxml-validate");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "dwb-pdf-pola-"));
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

function formPdf() {
  const content = [
    "BT /F1 11 Tf 72 700 Td (Miasto:) Tj ET",
    `BT /F1 11 Tf 115 700 Td (${".".repeat(50)}) Tj ET`,
    "BT /F1 11 Tf 275 700 Td (Teren nr:) Tj ET",
    `BT /F1 11 Tf 325 700 Td (${".".repeat(30)}) Tj ET`,
    "BT /F1 11 Tf 72 480 Td (Koniec karty.) Tj ET",
  ].join("\n");
  const helv = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>";
  const objs = [];
  objs.push("<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R 7 0 R 8 0 R] /NeedAppearances true /DA (/Helv 0 Tf 0 g) /DR << /Font << /Helv 5 0 R >> >> >> >>");
  objs.push("<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
  objs.push("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /Annots [6 0 R 7 0 R 8 0 R] >>");
  objs.push(`<< /Length ${Buffer.byteLength(content, "latin1")} >>\nstream\n${content}\nendstream`);
  objs.push(helv);
  objs.push("<< /Type /Annot /Subtype /Widget /FT /Tx /T (Text1) /V (PILECKIEGO 23) /Rect [118 696 266 711] /F 4 /DA (/Helv 13 Tf 0 g) /P 3 0 R >>");
  objs.push("<< /Type /Annot /Subtype /Widget /FT /Tx /T (Text2) /V () /Rect [328 696 415 711] /F 4 /DA (/Helv 10 Tf 0 g) /P 3 0 R >>");
  objs.push("<< /Type /Annot /Subtype /Widget /FT /Tx /T (Text3) /Ff 4096 /V (BLOK 3\rKLATKA 1) /Rect [72 560 300 640] /F 4 /DA (/Helv 11 Tf 0 g) /P 3 0 R >>");
  let out = "%PDF-1.4\n";
  const offs = [];
  objs.forEach((o, i) => { offs.push(Buffer.byteLength(out, "latin1")); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const xref = Buffer.byteLength(out, "latin1");
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n` + offs.map((o) => `${String(o).padStart(10, "0")} 00000 n \n`).join("");
  out += `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, "latin1");
}

async function convert(page, file) {
  await page.evaluate(() => {
    window.__conv = null;
    if (!window.__convHooked) { window.addEventListener("dwb:pdf-converted", (e) => (window.__conv = e.detail)); window.__convHooked = true; }
  });
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction(() => window.__conv, null, { timeout: 60000 });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p") && !inlineLocksPending, null, { timeout: 30000 });
  await page.waitForTimeout(500);
}

const fileFields = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(originalFileBytes);
  const d = await z.file("word/document.xml").async("string");
  const sdts = [...d.matchAll(/<w:sdt>([\s\S]*?)<\/w:sdt>/g)].map((m) => ({
    alias: (m[1].match(/w:alias w:val="([^"]*)"/) || [])[1] || "",
    text: [...m[1].matchAll(/<w:t[^>]*>([^<]*)/g)].map((x) => x[1]).join(""),
    multi: /w:multiLine="1"/.test(m[1]),
    textKind: /<w:text\b/.test(m[1]),
  }));
  const ids = [...d.matchAll(/<w:sdtPr>[\s\S]*?<w:id w:val="(-?\d+)"/g)].map((m) => m[1]);
  const plain = [...d.matchAll(/<w:t[^>]*>([^<]*)/g)].map((x) => x[1]).join("");
  return { sdts, idsUnique: new Set(ids).size === ids.length, plain, dotTabs: (d.match(/w:leader="dot"/g) || []).length };
});

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const pdf = path.join(TMP, "karta-pola.pdf");
  fs.writeFileSync(pdf, formPdf());

  // 1. domyślnie: pola PDF → kontrolki Worda
  check("przełącznik w panelu Widok domyślnie włączony", await page.evaluate(() => document.getElementById("pdfFormFields")?.checked === true));
  await convert(page, pdf);
  let ff = await fileFields(page);
  const city = ff.sdts.find((s) => s.text === "PILECKIEGO 23");
  check("pole z wpisem → kontrolka tekstowa z wpisem i tytułem z napisu przed nią", city && city.alias === "Miasto" && city.textKind, JSON.stringify(ff.sdts));
  const empty = ff.sdts.find((s) => s.alias === "Teren nr");
  check("puste pole nad kropkami → pusta kontrolka „Teren nr”", empty && empty.text === "", JSON.stringify(ff.sdts));
  const multi = ff.sdts.filter((s) => s.multi);
  check("pole wielowierszowe → kontrolka na wiersz, tytuł z 1. wiersza", multi.length === 2 && multi.map((s) => s.text).join("|") === "BLOK 3|KLATKA 1" && multi.every((s) => s.alias === "BLOK 3"), JSON.stringify(multi));
  check("kropki zostają tabulatorami z kropkami (wydruk jak w PDF)", ff.dotTabs >= 2, `${ff.dotTabs}`);
  check("numery kontrolek (w:id) jedyne", ff.idsUnique);
  const on = path.join(TMP, "pola-wlaczone.docx");
  fs.writeFileSync(on, Buffer.from(await page.evaluate(async () => Array.from(originalFileBytes))));

  // 2. apka rozpoznaje pola; akapity edytowalne
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await page.waitForTimeout(800);
  const app = await page.evaluate(() => ({
    counts: typeof docFormCounts !== "undefined" ? docFormCounts : null,
    shown: document.querySelectorAll(".docx-preview-host .ff-field").length,
    locks: collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.dataset.lock).map((p) => p.dataset.lock),
    dirty: collectInlineParagraphEdits().map((e) => e.index), // podgląd = plik (puste pole za kropkami stawało przed nimi)
    oneLine: (() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((q) => /Miasto/.test(q.textContent)); const lh = parseFloat(getComputedStyle(p).lineHeight) || 16; return p.getBoundingClientRect().height < lh * 1.8; })(),
  }));
  check("pola rozpoznane jako formularz (4, w tym 1 puste)", app.counts?.fields === 4 && app.counts?.empty === 1 && app.shown === 4, JSON.stringify(app));
  check("akapity z polami edytowalne (bez blokad)", !app.locks.length, JSON.stringify(app.locks));
  check("po otwarciu podgląd = plik (żaden akapit „zmieniony”, puste pole za kropkami na swoim miejscu)", !app.dirty.length, JSON.stringify(app.dirty));
  check("wiersz „Miasto … Teren nr …” w jednej linijce (pole nie łamie wiersza)", app.oneLine);

  // 3. klik w pole → wpis → Tab → następne pole (puste „Teren nr”) → wpis → Enter
  const r = await page.evaluate(() => { const w = [...document.querySelectorAll(".ff-field")].find((e) => e.textContent.includes("PILECKIEGO")); const b = w.getBoundingClientRect(); return { x: b.left + 8, y: b.top + b.height / 2 }; });
  await page.mouse.click(r.x, r.y);
  await page.waitForSelector(".ff-pop .ff-pop-input", { timeout: 5000 });
  check("okienko pola ma tytuł z karty („Miasto”)", await page.evaluate(() => document.querySelector(".ff-pop-title")?.textContent) === "Miasto");
  await page.fill(".ff-pop .ff-pop-input", "NIEPODLEGŁOŚCI 11");
  await page.keyboard.press("Tab");
  await page.waitForFunction(() => document.querySelector(".ff-pop-title")?.textContent === "Teren nr" && document.activeElement?.classList.contains("ff-pop-input"), null, { timeout: 8000 }).catch(() => {});
  check("Tab w okienku: następne pole („Teren nr”) gotowe do pisania", await page.evaluate(() => document.querySelector(".ff-pop-title")?.textContent === "Teren nr" && document.activeElement?.classList.contains("ff-pop-input")));
  await page.keyboard.type("301");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(800);
  await page.waitForFunction(() => !inlineLocksPending, null, { timeout: 10000 });
  const saved = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const d = await z.file("word/document.xml").async("string");
    return [...d.matchAll(/<w:sdt>([\s\S]*?)<\/w:sdt>/g)].map((m) => [...m[1].matchAll(/<w:t[^>]*>([^<]*)/g)].map((x) => x[1]).join(""));
  });
  check("wpisy w kontrolkach w pliku (Miasto, Teren nr)", saved[0] === "NIEPODLEGŁOŚCI 11" && saved[1] === "301", JSON.stringify(saved));
  const filled = path.join(TMP, "pola-wypelnione.docx");
  fs.writeFileSync(filled, Buffer.from(await page.evaluate(async () => Array.from(await buildDocumentForSave()))));

  // 3b. (chaos 2026-10-28) pole wielowierszowe z łamaniem wiersza: opakowanie z całą treścią po
  // przerysowaniu, Cofnij → Ponów bez podwojenia; link do akapitu z samą kontrolką = pełna zakładka
  const bm = await page.evaluate(async (target) => {
    await applyDocumentEdit({ op: "link", index: 0, start: 0, end: 0, targetIndex: target, text: "do klatek" });
    await waitInlineStructuralIdle();
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const x = await z.file("word/document.xml").async("string");
    return { s: (x.match(/<w:bookmarkStart\b/g) || []).length, e: (x.match(/<w:bookmarkEnd\b/g) || []).length };
  }, await page.evaluate(() => formScan.fields.find((x) => x.display === "KLATKA 1").paraIndex));
  check("link do akapitu z samą kontrolką: zakładka z początkiem i końcem", bm.s >= 1 && bm.s === bm.e, JSON.stringify(bm));
  const ml = await page.evaluate(async () => {
    const f = formScan.fields.find((x) => x.display === "KLATKA 1");
    await fillFormFields({ [f.key]: "KLATKA 1\nKLATKA 2" });
    await waitInlineStructuralIdle();
    const fileText = async () => { const z = await JSZip.loadAsync(await buildDocumentForSave()); const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml"); return collectParagraphElements(d.documentElement, "all")[f.paraIndex].textContent; };
    const a = await fileText();
    await dwbUndo.undo(); await dwbUndo.redo();
    await new Promise((r) => setTimeout(r, 400));
    const w = document.querySelector(`.docx-preview-host .ff-field[data-ff="${f.key}"]`);
    const ck = (await dwbSelfCheck()).problems.filter((x) => x.level === "error").map((x) => `${x.code}: ${x.msg}`);
    return { a, r: await fileText(), wrap: w?.textContent || "", glyph: !!w?.classList.contains("ff-glyph"), inline: collectInlineParagraphEdits().length, idx: f.paraIndex, ck };
  });
  check("pole wielowierszowe: po Cofnij → Ponów treść raz, opakowanie z oboma wierszami", ml.a === "KLATKA 1KLATKA 2" && ml.r === ml.a && ml.wrap === "KLATKA 1KLATKA 2" && !ml.glyph && !ml.inline && !ml.ck.length, JSON.stringify(ml));

  // 3c. (chaos 2026-10-28) Enter przed akapitem z samą kontrolką i Enter na zaznaczeniu z polem:
  // liczba kontrolek się zgadza, inne pola nietknięte, podgląd = plik
  const sdtState = () => page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const x = await z.file("word/document.xml").async("string");
    const vals = [...x.matchAll(/<w:sdt>([\s\S]*?)<\/w:sdt>/g)].map((m) => [...m[1].matchAll(/<w:t[^>]*>([^<]*)/g)].map((q) => q[1]).join(""));
    const ck = (await dwbSelfCheck()).problems.filter((q) => q.level === "error").map((q) => `${q.code}: ${q.msg}`);
    return { vals, ck };
  });
  const s0 = await sdtState();
  const kl = await page.evaluate(() => { const f = formScan.fields.find((x) => /^KLATKA/.test(x.display)); const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[f.paraIndex]; p.closest(".docx-edit-root").focus(); const r = document.createRange(); r.setStart(p, 0); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); return f.paraIndex; });
  await page.keyboard.press("Enter");
  await page.waitForTimeout(700);
  await page.waitForFunction(() => !inlineLocksPending, null, { timeout: 10000 });
  const s1 = await sdtState();
  check("Enter przed akapitem z kontrolką: kontrolki te same, podgląd = plik", JSON.stringify(s1.vals) === JSON.stringify(s0.vals) && !s1.ck.length, JSON.stringify({ kl, s0: s0.vals, s1 }));
  await page.evaluate(() => { const f = formScan.fields.find((x) => x.display === "NIEPODLEGŁOŚCI 11"); const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[f.paraIndex]; const w = p.querySelector(`.ff-field[data-ff="${f.key}"]`); p.closest(".docx-edit-root").focus(); const r = document.createRange(); r.setStartBefore(w); r.setEndAfter(w.nextSibling || w); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await page.keyboard.press("Enter");
  await page.waitForTimeout(700);
  await page.waitForFunction(() => !inlineLocksPending, null, { timeout: 10000 });
  const s2 = await sdtState();
  check("zaznaczenie z polem + Enter: pole usunięte, pozostałe kontrolki nietknięte, podgląd = plik", JSON.stringify(s2.vals) === JSON.stringify(s1.vals.filter((v) => v !== "NIEPODLEGŁOŚCI 11")) && !s2.ck.length, JSON.stringify({ s1: s1.vals, s2 }));

  // 3d. (chaos 2026-10-28, runda 4) Enter na części tekstu pola; Backspace za polem będącym całą
  // treścią akapitu; Backspace na końcu wiersza z polem + Delete (zastępczy <br> przeglądarki)
  const settle = async () => { await page.waitForTimeout(600); await page.waitForFunction(() => !inlineLocksPending, null, { timeout: 10000 }); };
  const errsNow = () => page.evaluate(async () => (await dwbSelfCheck()).problems.filter((q) => q.level === "error").map((q) => `${q.code}: ${q.msg}`));
  const caretIn = (find, at) => page.evaluate(([find, at]) => {
    const f = formScan.fields.find((x) => new RegExp(find).test(x.display));
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[f.paraIndex];
    const w = p.querySelector(`.ff-field[data-ff="${f.key}"]`);
    p.closest(".docx-edit-root").focus();
    const r = document.createRange();
    if (at === "inside") { const t = document.createTreeWalker(w, NodeFilter.SHOW_TEXT).nextNode(); r.setStart(t, 1); r.setEnd(t, 3); }
    else if (at === "after") { r.setStartAfter(w); r.collapse(true); }
    else { r.selectNodeContents(p); r.collapse(false); }
    getSelection().removeAllRanges(); getSelection().addRange(r);
    return f.key;
  }, [find, at]);
  await caretIn("^KLATKA", "inside");
  await page.keyboard.press("Enter");
  await settle();
  let ek = await errsNow();
  check("Enter na części tekstu pola: pole całe, podgląd = plik", !ek.length, ek.join(" | "));
  await caretIn("^KLATKA", "after");
  await page.keyboard.press("Backspace");
  await settle();
  ek = await errsNow();
  check("Backspace za polem (cała treść akapitu): akapit zostaje, podgląd = plik", !ek.length, ek.join(" | "));
  await caretIn("^301$", "end");
  for (let i = 0; i < 3; i++) await page.keyboard.press("Backspace");
  await settle();
  await page.keyboard.press("Delete");
  await page.keyboard.press("Delete");
  await settle();
  ek = await errsNow();
  check("Backspace na końcu wiersza z polem, potem Delete: podgląd = plik", !ek.length, ek.join(" | "));

  // 4. przełącznik wyłączony: dawny wynik (zwykły tekst, bez kontrolek)
  await page.evaluate(() => { const b = document.getElementById("pdfFormFields"); b.checked = false; b.dispatchEvent(new Event("change")); });
  check("przełącznik zapamiętany", await page.evaluate(() => typeof pdfFormFieldsOn === "function" && pdfFormFieldsOn() === false));
  await convert(page, pdf);
  ff = await fileFields(page);
  check("wyłączone: bez kontrolek, wpis z PDF jako zwykły tekst", ff.sdts.length === 0 && /PILECKIEGO 23/.test(ff.plain) && /KLATKA 1/.test(ff.plain), JSON.stringify(ff));
  // kropki za wpisem większą czcionką (13 pt przy tekście 11 pt) — na tej samej wysokości co
  // pozostałe (2026-10-28: przypięte do góry linijki, ~2,5 px wyżej; zeroWhitespaceRun przed layoutTabStops)
  for (const ro of [true, false]) {
    await page.evaluate((ro) => { if (readOnlyMode !== ro) { readModeEl.checked = ro; readModeEl.dispatchEvent(new Event("change")); } }, ro);
    await page.waitForTimeout(700);
    const tops = await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((q) => /Miasto/.test(q.textContent)); return [...p.querySelectorAll(".docx-tab.tab-dot")].map((t) => Math.round(t.getBoundingClientRect().bottom * 10) / 10); });
    check(`kropki w wierszu na jednej wysokości (${ro ? "Czytanie" : "Edycja"})`, tops.length >= 2 && Math.max(...tops) - Math.min(...tops) < 0.6, JSON.stringify(tops));
  }
  await page.evaluate(() => { const b = document.getElementById("pdfFormFields"); b.checked = true; b.dispatchEvent(new Event("change")); });

  // 5. schemat Office (gdy jest .NET)
  if (ooxml.available()) {
    for (const v of ooxml.validate([on, filled])) check(`schemat Office: ${path.basename(v.file)}`, v.ok, v.error || `${v.errors.length}: ${v.errors.slice(0, 4).map(ooxml.short).join(" | ")}`);
  }

  await browser.close();
  fs.rmSync(TMP, { recursive: true, force: true });
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Pola formularza z PDF [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => {
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  console.error(e);
  process.exit(1);
});
