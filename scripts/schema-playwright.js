// schema-playwright.js — zapis zgodny ze schematem Office (Open XML SDK, scripts/ooxml-validate.js).
//
//   node scripts/schema-playwright.js                  (Chromium, pliki z docs/samples + nowe dokumenty)
//   ENGINE=webkit node scripts/schema-playwright.js
//   FILES="a.docx:b.docx" node scripts/schema-playwright.js   (własne pliki zamiast przykładów)
//   KEEP=1 — zapisane pliki zostają w output/schema/ (do obejrzenia w Wordzie)
//
// Każdy plik: otwarcie → Edycja → szeroki zestaw edycji (pisanie, Enter, Backspace, B/I/U, kolor,
// wyróżnienie, krój, rozmiar, style akapitów, wyrównanie, listy i poziomy, tabela, komentarz,
// przypisy, link, podział strony, linia, spis treści, pola formularza, wklejki, obraz, nagłówek
// i stopka, marginesy, metadane) → zapis → walidator. Liczą się tylko błędy, których NIE było
// w oryginale (część pliku + opis). Do tego nowe dokumenty z szablonów (pusty, pismo, notatka).
// Walidator: tools/ooxml-validator (.NET — brew install dotnet); bez .NET test się pomija.
// Dlaczego: poprawny XML to za mało — element w złej kolejności (np. w w:rPr) to dla Worda
// „nieczytelna zawartość” (2026-10-05: pierwszy przebieg znalazł w:themeFontLang przed w:compat
// w settings.xml każdego nowego dokumentu).

const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");
const ooxml = require("./ooxml-validate");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const OUT = process.env.KEEP ? path.resolve(__dirname, "../output/schema") : fs.mkdtempSync(path.join(os.tmpdir(), "dwb-schema-"));
fs.mkdirSync(OUT, { recursive: true });

const SAMPLES = path.resolve(__dirname, "../docs/samples");
const files = process.env.FILES ? process.env.FILES.split(":").filter(Boolean) : fs.readdirSync(SAMPLES).filter((f) => f.endsWith(".docx")).map((f) => path.join(SAMPLES, f));

async function openPage(browser) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1300, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.setItem("dwb-comment-author", "Test"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  return { context, page, errors };
}
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 })
  .then(() => page.evaluate(() => waitInlineStructuralIdle())).then(() => page.waitForTimeout(50)); // gotowość = nakładka + kolejka; dawniej 250 ms × 36 operacji × 9 plików ≈ 80 s samego czekania
const toEdit = async (page) => {
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await page.waitForFunction(() => !readOnlyMode && document.querySelector(".docx-edit-root"), null, { timeout: 15000 });
  await idle(page);
};
const saveTo = async (page, file) => {
  const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000)); return btoa(s); });
  fs.writeFileSync(file, Buffer.from(b64, "base64"));
};

// Seria edycji na otwartym dokumencie. Zwraca etykiety wykonanych operacji.
async function editAll(page) {
  let n = 0;
  // akapit z tekstem do edycji (kolejne operacje na różnych akapitach)
  const target = (k = n++) => page.evaluate((k) => {
    const all = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.isContentEditable && !p.dataset.lock && !p.dataset.noteKey && !p.closest("td, th"));
    // akapit z tekstem; w małym pliku (gdy edycje zjadły dłuższe) — jakikolwiek z kilkoma literami
    let ps = all.filter((p) => p.textContent.trim().length > 20);
    if (ps.length < 3) ps = all.filter((p) => /\p{L}{4,}/u.test(p.textContent));
    if (!ps.length) throw new Error("brak akapitu do edycji");
    const p = ps[(k * 7) % ps.length];
    p.scrollIntoView({ block: "center" });
    docEditRoot()?.focus({ preventScroll: true });
    return resolveParaIndex(p);
  }, k);
  const caret = (i, where) => page.evaluate(([i, where]) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    const len = p.textContent.length;
    placeCaret(p, where === "end" ? len : where === "start" ? 0 : Math.floor(len / 2));
  }, [i, where]);
  const word = (i) => page.evaluate((i) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = w.nextNode(); t; t = w.nextNode()) {
      const m = t.data.match(/\p{L}{4,}/u);
      if (m && !t.parentElement.closest('[contenteditable="false"], a')) { const r = document.createRange(); r.setStart(t, m.index); r.setEnd(t, m.index + m[0].length); const s = getSelection(); s.removeAllRanges(); s.addRange(r); return { text: m[0] }; }
    }
    return null;
  }, i);
  const offsets = (i, text) => page.evaluate(([i, text]) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    const full = previewRunsToPlainText(extractRunsFromPreviewParagraph(p));
    const s = full.indexOf(text);
    return s < 0 ? null : { start: s, end: s + text.length };
  }, [i, text]);
  const paste = (data) => page.evaluate((data) => { const dt = new DataTransfer(); Object.entries(data).forEach(([k, v]) => dt.setData(k, v)); docCanvasEl.querySelector(".docx-edit-root")?.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true })); }, data);
  const edit = (e) => page.evaluate((e) => applyDocumentEdit(e), e);

  const ops = [
    ["pisanie", async () => { const i = await target(); await caret(i, "end"); await page.keyboard.type(" dopisek ąę"); }],
    ["Enter w środku", async () => { const i = await target(); await caret(i, "mid"); await page.keyboard.press("Enter"); }],
    ["Backspace skleja", async () => { const i = await target(); await caret(i, "start"); await page.keyboard.press("Backspace"); }],
    ["B / I / U słowa", async () => { const i = await target(); if (!(await word(i))) return false; await page.keyboard.press(`${MOD}+b`); await idle(page); await word(i); await page.keyboard.press(`${MOD}+i`); await idle(page); await word(i); await page.keyboard.press(`${MOD}+u`); }],
    ["kolor tekstu", async () => { const i = await target(); if (!(await word(i))) return false; await page.evaluate(() => composeUi.applyColor("color", "C00000")); }],
    ["wyróżnienie", async () => { const i = await target(); if (!(await word(i))) return false; await page.evaluate(() => composeUi.applyColor("highlight", "yellow")); }],
    ["krój", async () => { const i = await target(); if (!(await word(i))) return false; await page.selectOption("#fmtFontFamily", await page.evaluate(() => [...document.querySelectorAll("#fmtFontFamily option")].map((o) => o.value).find((v) => /^(Georgia|Arial|Times New Roman)$/.test(v)))); }],
    ["rozmiar", async () => { const i = await target(); if (!(await word(i))) return false; await page.selectOption("#fmtFontSize", "14"); }],
    ["A+", async () => { const i = await target(); if (!(await word(i))) return false; await page.click("#fmtGrow"); }],
    ["Nagłówek 1", async () => { const i = await target(); await caret(i, "mid"); await page.selectOption("#fmtParaStyle", "h1"); }],
    ["Nagłówek 2", async () => { const i = await target(); await caret(i, "mid"); await page.selectOption("#fmtParaStyle", "h2"); }],
    ["cytat", async () => { const i = await target(); await caret(i, "mid"); await page.selectOption("#fmtParaStyle", "quote"); }],
    ["ramka", async () => { const i = await target(); await caret(i, "mid"); await page.selectOption("#fmtParaStyle", "callout"); }],
    ["indeks górny i dolny", async () => { const i = await target(); if (!(await word(i))) return false; await page.click("#fmtSuper"); await idle(page); await word(i); await page.click("#fmtSub"); }],
    ["Aa: WIELKIE", async () => { const i = await target(); if (!(await word(i))) return false; await page.evaluate(() => changeTextCase("upper")); }],
    ["wyczyść formatowanie", async () => { const i = await target(); await caret(i, "mid"); await page.evaluate(() => clearTextFormatting()); }],
    ["odstępy i wcięcia", async () => { const i = await target(); await edit({ op: "paraFormat", indices: [i], spacing: { before: 120, after: 0, line: 360, lineRule: "auto" }, ind: { left: 567, right: 283, hanging: 283 } }); await edit({ op: "paraFormat", indices: [i], indentDelta: 709 }); await edit({ op: "paraFormat", indices: [i], spacing: { line: 300, lineRule: "exact" }, ind: { firstLine: 400, hanging: 0 } }); }],
    ["wyśrodkowanie", async () => { const i = await target(); await caret(i, "mid"); await page.evaluate(() => composeUi.applyAlign("center")); }],
    ["wyjustowanie", async () => { const i = await target(); await caret(i, "mid"); await page.evaluate(() => composeUi.applyAlign("justify")); }],
    ["lista punktowana + poziom", async () => { const i = await target(); await caret(i, "mid"); await page.evaluate(() => composeUi.applyList("bullet")); await idle(page); await caret(i, "start"); await page.keyboard.press("Tab"); }],
    ["lista numerowana", async () => { const i = await target(); await caret(i, "mid"); await page.evaluate(() => composeUi.applyList("number")); }],
    ["tabela 2×3", async () => { const i = await target(); await caret(i, "end"); await page.evaluate(() => composeUi.insertTable(2, 3)); }],
    ["komentarz", async () => { const i = await target(); const w = await word(i); if (!w) return false; const o = await offsets(i, w.text); await edit({ op: "commentAdd", index: i, start: o.start, end: o.end, rich: [[{ text: "Uwaga " }, { text: "testowa", b: true }]], author: "Test", initials: "T" }); }],
    ["przypis dolny", async () => { const i = await target(); await caret(i, "end"); await page.evaluate(() => composeUi.insertNote("footnote")); }],
    ["przypis końcowy", async () => { const i = await target(); await caret(i, "mid"); await page.evaluate(() => composeUi.insertNote("endnote")); }],
    ["link", async () => { const i = await target(); const w = await word(i); if (!w) return false; const o = await offsets(i, w.text); await edit({ op: "link", index: i, start: o.start, end: o.end, href: "https://example.com/a?b=1&c=2" }); }],
    ["zakładka + link do niej (etykietka)", async () => { const i = await target(); const w = await word(i); if (!w) return false; const o = await offsets(i, w.text); await edit({ op: "bookmark", action: "add", name: "Test_Zakł", from: { index: i, offset: o.start }, to: { index: i, offset: o.end } }); await edit({ op: "link", index: i, start: o.start, end: o.end, anchor: "Test_Zakł", tooltip: "Etykietka & <test>" }); }],
    ["zakładka na 2 akapity + link na początek", async () => { const i = await target(); const w = await word(i); if (!w) return false; await edit({ op: "bookmark", action: "add", name: "Zakres", from: { index: i, offset: 0 }, to: { index: i + 1, offset: 1 } }); const o = await offsets(i, w.text); await edit({ op: "link", index: i, start: o.start, end: o.end, anchor: "_top" }); }],
    ["link e-mail z tematem", async () => { const i = await target(); const w = await word(i); if (!w) return false; const o = await offsets(i, w.text); await edit({ op: "link", index: i, start: o.start, end: o.end, href: "mailto:a@example.com?subject=Temat%20%C4%85" }); }],
    ["usunięcie zakładki", async () => { await edit({ op: "bookmark", action: "delete", name: "Test_Zakł" }); }],
    ["podział strony", async () => { const i = await target(); await caret(i, "end"); await page.evaluate(() => composeUi.insertPageBreak()); }],
    ["linia pozioma", async () => { const i = await target(); await caret(i, "end"); await page.evaluate(() => composeUi.insertHrule()); }],
    ["spis treści", async () => { const i = await target(); await caret(i, "start"); await page.evaluate(() => composeUi.insertToc()); }],
    ["pole: pole wyboru", async () => { const i = await target(); await caret(i, "end"); await page.evaluate(() => composeUi.insertFormField("checkbox")); }],
    ["pole: tekst", async () => { const i = await target(); await caret(i, "end"); await page.evaluate(() => composeUi.insertFormField("text")); }],
    ["pole: data", async () => { const i = await target(); await caret(i, "end"); await page.evaluate(() => composeUi.insertFormField("date")); }],
    ["pole: lista", async () => { const i = await target(); await caret(i, "end"); await page.evaluate(() => composeUi.insertFormField("dropdown", ["Tak", "Nie"])); }],
    ["wklejka Markdown", async () => { const i = await target(); await caret(i, "mid"); await paste({ "text/plain": "## Nagłówek z wklejki\n\n- punkt **gruby**\n  - pod-punkt\n1. pierwszy\n- [x] zrobione\n\n| A | B |\n|---|---|\n| 1 | 2 |" }); }],
    ["wklejka HTML", async () => { const i = await target(); await caret(i, "mid"); await paste({ "text/html": "<p><b>Gruby</b> i <i>pochyły</i> <a href='https://x.pl'>link</a></p><ul><li>raz</li><li>dwa</li></ul>", "text/plain": "Gruby i pochyły link\nraz\ndwa" }); }],
    ["wklejka wierszy", async () => { const i = await target(); await caret(i, "mid"); await paste({ "text/plain": "wiersz pierwszy\n\nwiersz\tz tabulatorem" }); }],
    ["obraz", async () => {
      const i = await target(); await caret(i, "end");
      await page.evaluate(async () => {
        const c = Object.assign(document.createElement("canvas"), { width: 120, height: 60 });
        const g = c.getContext("2d"); g.fillStyle = "#2b6cb0"; g.fillRect(0, 0, 120, 60);
        const blob = await new Promise((r) => c.toBlob(r, "image/png"));
        const p = docCaretParagraph(document.activeElement) || restoreDocCaret();
        await composeUi.insertImageFile(new File([blob], "obraz.png", { type: "image/png" }), p);
      });
    }],
    ["nagłówek i stopka z numerem", async () => { await edit({ op: "headerFooter", lang: "pl", total: 3, header: { text: "Nagłówek testu", align: "right" }, footer: { text: "Stopka", align: "left" }, number: { fmt: "n", align: "center" }, firstDifferent: false }); }],
    ["marginesy + pozioma", async () => { await edit({ op: "pageSetup", index: 0, scope: "all", margins: { top: 1134, bottom: 1134, left: 1418, right: 1418 }, orient: "landscape" }); }],
    ["metadane", async () => { await edit({ op: "coreMetadata", fields: { title: "Tytuł & <test>", creator: "Test", subject: "schemat" } }); }],
  ];
  const done = [];
  const failed = [];
  for (const [label, fn] of ops) {
    try {
      const r = await fn();
      await idle(page);
      if (r === false) continue;
      done.push(label);
    } catch (e) {
      failed.push(`${label}: ${String(e.message).split("\n")[0].slice(0, 120)}`);
      await idle(page).catch(() => {});
    }
  }
  return { done, failed };
}

async function run() {
  if (!ooxml.available()) {
    console.log("⚠️  Pominięto: brak .NET albo walidatora (brew install dotnet; tools/ooxml-validator).");
    process.exit(0);
  }
  const browser = await pw[ENGINE].launch({ headless: true });
  try {
    // ── pliki: oryginał vs po edycjach ──
    for (const file of files) {
      const name = path.basename(file, ".docx");
      const { context, page, errors } = await openPage(browser);
      await page.locator("#fileInput").setInputFiles(file);
      await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 60000 });
      await toEdit(page);
      const plain = path.join(OUT, `${name}.bez-zmian.${ENGINE}.docx`);
      await saveTo(page, plain); // zapis bez edycji (oryginał przechodzi przez naprawy przy otwarciu)
      const { done, failed } = await editAll(page);
      const edited = path.join(OUT, `${name}.po-edycji.${ENGINE}.docx`);
      await saveTo(page, edited);
      const [vo, vp, ve] = ooxml.validate([file, plain, edited]);
      const np = ooxml.newErrors(vo, vp);
      const ne = ooxml.newErrors(vo, ve);
      check(`${name}: zapis bez zmian — bez nowych błędów schematu`, !vp.error && !np.length, vp.error || np.slice(0, 4).map(ooxml.short).join(" | "));
      check(`${name}: po ${done.length} edycjach — bez nowych błędów schematu`, !ve.error && !ne.length, ve.error || `${ne.length} nowych: ${ne.slice(0, 5).map(ooxml.short).join(" | ")}`);
      if (failed.length) console.log(`   ℹ️  ${name}: operacje niewykonane (nie błąd schematu): ${failed.join("; ")}`);
      const real = errors.filter((e) => !/ResizeObserver/.test(e));
      check(`${name}: bez błędów strony`, !real.length, real.slice(0, 3).join(" | "));
      await context.close();
    }
    // ── nowe dokumenty z szablonów (+ edycje na pustym) ──
    for (const kind of ["blank", "letter", "note"]) {
      const { context, page, errors } = await openPage(browser);
      await page.evaluate((k) => composeUi.createNew(k), kind);
      await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
      await toEdit(page);
      const fresh = path.join(OUT, `nowy-${kind}.${ENGINE}.docx`);
      await saveTo(page, fresh);
      let edited = null;
      if (kind === "letter") {
        await editAll(page);
        edited = path.join(OUT, `nowy-${kind}.po-edycji.${ENGINE}.docx`);
        await saveTo(page, edited);
      }
      if (kind === "blank") {
        // plik utworzony starszą wersją aplikacji: w:themeFontLang przed w:compat — otwarcie i zapis naprawiają
        const zip = await JSZip.loadAsync(fs.readFileSync(fresh));
        const st = await zip.file("word/settings.xml").async("string");
        const lang = st.match(/<w:themeFontLang\b[^>]*\/>/)[0];
        zip.file("word/settings.xml", st.replace(lang, "").replace("<w:compat>", `${lang}<w:compat>`));
        const old = path.join(OUT, `stary-z-aplikacji.${ENGINE}.docx`);
        fs.writeFileSync(old, await zip.generateAsync({ type: "nodebuffer" }));
        const o = await openPage(browser);
        await o.page.locator("#fileInput").setInputFiles(old);
        await o.page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
        const resaved = path.join(OUT, `stary-z-aplikacji.zapis.${ENGINE}.docx`);
        await saveTo(o.page, resaved);
        const [vOld, vNew] = ooxml.validate([old, resaved]);
        check("stary plik z aplikacji (themeFontLang przed compat): zapis naprawia kolejność", !vOld.ok && vNew.ok, `${vOld.errors.length} → ${vNew.error || vNew.errors.map(ooxml.short).join(" | ")}`);
        await o.context.close();
      }
      const [vf, ve] = ooxml.validate(edited ? [fresh, edited] : [fresh]);
      check(`nowy dokument „${kind}”: zgodny ze schematem`, vf.ok, vf.error || vf.errors.slice(0, 4).map(ooxml.short).join(" | "));
      if (ve) check(`nowy dokument „${kind}” po edycjach: zgodny ze schematem`, ve.ok, ve.error || `${ve.errors.length}: ${ve.errors.slice(0, 5).map(ooxml.short).join(" | ")}`);
      const real = errors.filter((e) => !/ResizeObserver/.test(e));
      check(`nowy „${kind}”: bez błędów strony`, !real.length, real.slice(0, 3).join(" | "));
      await context.close();
    }
  } finally {
    await browser.close();
    if (!process.env.KEEP) fs.rmSync(OUT, { recursive: true, force: true });
  }
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
