// insertables-playwright.js — pełny cykl życia elementów wstawianych do dokumentu (lista kontrolna 2026-11-09).
//
//   node scripts/insertables-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/insertables-playwright.js (Safari/iPhone)
//
// Dla każdego elementu (podział strony, tabela, przypis, pole tekstowe i daty formularza, spis treści,
// pole tekstowe-ramka, linia pozioma) sprawdza w PLIKU: wstawienie, Backspace tuż po wstawieniu + Cofnij,
// Enter (bez duplikatów), Delete przed elementem (jak w Wordzie), zapis → ponowne otwarcie, Ctrl/⌘+A
// + litera, Kopiuj/Wklej (jedyne numery kontrolek, kształtów, przypisy z treścią) i walidację Open XML.
// Zgłoszenia, z których wyrosło: linia pozioma nieusuwalna (2026-11-08), Kopiuj/Wklej pola formularza
// dublował w:sdt/w:id, przypis na wydruku bez numeru (2026-11-09).
const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
let ooxml = null;
try { ooxml = require("./ooxml-validate"); } catch (_) { ooxml = null; }

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const SAMPLE = fs.readFileSync(path.resolve(__dirname, "../docs/samples/headings-sample.docx"));
const GUIDE = path.resolve(__dirname, "../docs/samples/przewodnik.docx");
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const mod = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

// re: jak policzyć element w document.xml; delBefore: ile zostaje po Delete na końcu akapitu przed (Word)
const KINDS = {
  "podział strony": { ins: () => composeUi.insertPageBreak(), re: "<w:pageBreakBefore\\b|w:type=\"page\"", delBefore: 0 },
  tabela: { ins: () => composeUi.insertTable(2, 2), re: "<w:tbl[ >]", delBefore: 1 },
  przypis: { ins: () => composeUi.insertNote("footnote"), re: "<w:footnoteReference\\b", delBefore: 1 },
  "pole formularza": { ins: () => composeUi.insertFormField("text"), re: "<w:sdt>(?![\\s\\S]{0,300}docPartGallery)", delBefore: 1 },
  "pole daty": { ins: () => composeUi.insertFormField("date"), re: "<w:date\\b", delBefore: 1 },
  "spis treści": { ins: () => composeUi.insertToc(), re: "TOC \\\\|Table of Contents", delBefore: 1 },
  "pole tekstowe": { ins: () => dwbTextbox.insert(), re: "<wps:txbx>", delBefore: 1 },
  "linia pozioma": { ins: () => composeUi.insertHrule(), re: "<w:pBdr>", delBefore: 0 },
};

async function newPage(browser, opts = {}) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 }, ...opts });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  return { context, page, errors };
}
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(450));
async function openSample(page, buf = SAMPLE) {
  await page.locator("#fileInput").setInputFiles({ name: "elementy.docx", mimeType: DOCX_MIME, buffer: buf });
  await idle(page);
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await page.waitForTimeout(250);
}
const count = (page, re) => page.evaluate(async (src) => {
  const xml = await (await JSZip.loadAsync(await buildDocumentForSave())).file("word/document.xml").async("string");
  return (xml.match(new RegExp(src, "g")) || []).length;
}, re);
const caretEnd = (page, i) => page.evaluate((idx) => { const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")); const p = ps[Math.min(idx, ps.length - 1)]; placeCaret(p, p.textContent.length); }, i);
const saveBytes = (page) => page.evaluate(async () => Array.from(await buildDocumentForSave()));

async function insertOne(page, name, k) {
  await caretEnd(page, 1);
  await page.evaluate(k.ins);
  await idle(page);
  if (name === "pole tekstowe") { await page.keyboard.type("ramka"); await page.keyboard.press("Escape"); await idle(page); }
}

async function main() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const allErrors = [];
  for (const [name, k] of Object.entries(KINDS)) {
    // A. wstaw → Backspace → Cofnij → Enter → zapis/otwarcie → Ctrl/⌘+A + litera
    const a = await newPage(browser);
    await openSample(a.page);
    const base = await count(a.page, k.re);
    await insertOne(a.page, name, k);
    const ins = await count(a.page, k.re);
    check(`${name}: wstawienie — jeden element w pliku`, ins === base + 1, `${base} → ${ins}`);
    await a.page.keyboard.press("Backspace"); await idle(a.page);
    const afterBs = await count(a.page, k.re);
    if (afterBs < ins) {
      await a.page.keyboard.press(`${mod}+z`); await idle(a.page);
      check(`${name}: Backspace usunął element, Cofnij go przywraca`, (await count(a.page, k.re)) === ins);
    }
    await a.page.keyboard.press("Enter"); await idle(a.page);
    check(`${name}: Enter nie dubluje elementu`, (await count(a.page, k.re)) <= ins, String(await count(a.page, k.re)));
    const cur = await count(a.page, k.re);
    const f = path.join(os.tmpdir(), `dwb-el-${process.pid}.docx`);
    fs.writeFileSync(f, Buffer.from(await saveBytes(a.page)));
    await a.page.locator("#fileInput").setInputFiles(f); await idle(a.page);
    fs.rmSync(f, { force: true });
    check(`${name}: zapis → ponowne otwarcie — element zostaje`, (await count(a.page, k.re)) === cur, `${cur} → ${await count(a.page, k.re)}`);
    await a.page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
    await a.page.evaluate(() => { const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")); placeCaret(ps[0], 0); });
    await a.page.keyboard.press(`${mod}+a`); await a.page.keyboard.type("X"); await idle(a.page);
    check(`${name}: Ctrl/⌘+A + litera — elementu nie ma w pliku`, (await count(a.page, k.re)) === 0);
    allErrors.push(...a.errors);
    await a.context.close();

    // B. Delete na końcu akapitu przed elementem (zachowanie jak w Wordzie) + Cofnij
    const b = await newPage(browser);
    await openSample(b.page);
    await insertOne(b.page, name, k);
    const n0 = await count(b.page, k.re);
    if (name === "pole tekstowe") await b.page.evaluate(() => dwbTextbox.deselect());
    await caretEnd(b.page, 1);
    await b.page.keyboard.press("Delete"); await idle(b.page);
    const nd = await count(b.page, k.re);
    check(`${name}: Delete przed elementem — ${k.delBefore ? "element zostaje" : "element znika"} (jak w Wordzie)`, nd === (k.delBefore ? n0 : n0 - 1), `${n0} → ${nd}`);
    if (nd !== n0) { await b.page.keyboard.press(`${mod}+z`); await idle(b.page); check(`${name}: Cofnij po Delete przywraca`, (await count(b.page, k.re)) === n0); }
    allErrors.push(...b.errors);
    await b.context.close();
  }

  // C. wszystko w jednym dokumencie + Kopiuj/Wklej fragmentu: jedyne numery, przypisy z treścią, walidacja
  const c = await newPage(browser);
  await openSample(c.page);
  for (const [name, k] of Object.entries(KINDS)) {
    if (name === "spis treści") { await c.page.evaluate(() => { const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")); placeCaret(ps[0], 0); }); await c.page.evaluate(k.ins); await idle(c.page); continue; }
    await insertOne(c.page, name, k);
  }
  await c.page.evaluate(() => { const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")); const r = document.createRange(); r.setStartBefore(ps[2].firstChild || ps[2]); r.setEndAfter(ps[Math.min(14, ps.length - 2)].lastChild || ps[14]); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await c.page.keyboard.press(`${mod}+c`);
  await caretEnd(c.page, 999);
  await c.page.keyboard.press(`${mod}+v`); await idle(c.page);
  const ids = await c.page.evaluate(async () => {
    const zip = await JSZip.loadAsync(await buildDocumentForSave());
    const xml = await zip.file("word/document.xml").async("string");
    const fn = zip.file("word/footnotes.xml") ? await zip.file("word/footnotes.xml").async("string") : "";
    const dup = (re) => { const v = [...xml.matchAll(re)].map((m) => m[1]); return v.length - new Set(v).size; };
    const refs = [...xml.matchAll(/<w:footnoteReference[^>]*w:id="(\d+)"/g)].map((m) => m[1]);
    const bodies = new Set([...fn.matchAll(/<w:footnote\b[^>]*w:id="(-?\d+)"/g)].map((m) => m[1]));
    return { sdt: dup(/<w:sdtPr>[\s\S]*?<w:id w:val="(-?\d+)"/g), docPr: dup(/<wp:docPr\b[^>]*?\sid="(\d+)"/g), fnRefDup: refs.length - new Set(refs).size, fnOrphan: refs.filter((x) => !bodies.has(x)).length, fields: (xml.match(/<w:sdt>/g) || []).length };
  });
  check("Kopiuj/Wklej: jedyne numery kontrolek (w:sdt/w:id)", ids.sdt === 0, JSON.stringify(ids));
  check("Kopiuj/Wklej: jedyne numery kształtów (wp:docPr)", ids.docPr === 0, JSON.stringify(ids));
  check("Kopiuj/Wklej: każdy przypis ma własną treść", ids.fnRefDup === 0 && ids.fnOrphan === 0, JSON.stringify(ids));
  if (ooxml?.available?.()) {
    const f = path.join(os.tmpdir(), `dwb-el-all-${process.pid}.docx`);
    fs.writeFileSync(f, Buffer.from(await saveBytes(c.page)));
    const [res] = ooxml.validate([f]); // [{ file, ok, errors }]
    fs.rmSync(f, { force: true });
    const errs = res?.ok ? [] : (res?.errors?.length ? res.errors : [res?.error || "?"]);
    check("Open XML SDK: dokument ze wszystkimi elementami i kopiami bez błędów", errs.length === 0, JSON.stringify(errs).slice(0, 300));
  } else console.log("⚠️  walidator Open XML niedostępny — pominięto");
  allErrors.push(...c.errors);
  await c.context.close();

  // D. przypis na wydruku: numer na początku treści przypisu (jak w Edycji i w Wordzie)
  const d = await newPage(browser);
  await openSample(d.page, fs.readFileSync(GUIDE));
  await d.page.evaluate(() => dwbPrint.open());
  await d.page.waitForFunction(() => document.querySelector(".pp-sheet"), null, { timeout: 30000 });
  await d.page.waitForTimeout(500);
  const pn = await d.page.evaluate(() => [...document.querySelectorAll(".pp-page-notes li")].map((li) => li.textContent.trim().slice(0, 30)));
  check("wydruk: przypis na dole kartki zaczyna się od numeru", pn.length > 0 && pn.every((t) => /^\S{1,3}\s/.test(t)), JSON.stringify(pn));
  allErrors.push(...d.errors);
  await d.context.close();

  check("bez błędów strony", allErrors.length === 0, allErrors.slice(0, 3).join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail || ""}`}`);
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}

main().catch((e) => {
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail || ""}`}`);
  console.error(e);
  process.exit(1);
});
