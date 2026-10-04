// notes-playwright.js — przypisy dolne i końcowe (app/doc-notes.js).
//
// Fixture: docs/samples/notes-sample.docx (node scripts/gen-notes-docx.mjs). Numeracja jak w
// Wordzie (ciągła przez podział strony, przypis końcowy „i”), dymek z treścią, skok do przypisu,
// akapity z odnośnikiem edytowalne (dawniej tylko do odczytu), pisanie przed odnośnikiem i za
// odnośnikiem na końcu akapitu, edycja tekstu przypisu (dopisanie, Enter, Backspace nie kasuje
// numeru), przypis z tabelą tylko do odczytu, zapis (odnośniki, numery, tabulator, separator
// zostają), Cofnij/Ponów, przerysowanie bez utraty zmian. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const count = (s, re) => (s.match(re) || []).length;

function noteTexts(xml, tag, id) {
  const m = xml.match(new RegExp(`<w:${tag} w:id="${id}">([\\s\\S]*?)</w:${tag}>`));
  if (!m) return null;
  return m[1].split("</w:p>").filter((x) => x.includes("<w:p")).map((p) => (/w:(footnoteRef|endnoteRef)\/>/.test(p) ? "[nr]" : "") + (p.match(/<w:t[^>]*>[^<]*/g) || []).map((x) => x.replace(/<w:t[^>]*>/, "")).join(""));
}

(async () => {
  const browser = await pw[ENGINE].launch();
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
  await page.goto(`${APP_URL}?sample=notes-sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForFunction(() => document.querySelectorAll("sup[data-dwb-note]").length === 4 && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  await sleep(600);

  // ── Czytanie: numery, dymek, skok ──
  const read = await page.evaluate(() => ({
    labels: [...document.querySelectorAll("sup[data-dwb-note]")].map((s) => `${s.dataset.dwbNote}=${s.textContent}`),
    marks: [...document.querySelectorAll("ol.dwb-notes > li")].map((li) => `${li.dataset.dwbNote}=${li.querySelector("sup.note-mark")?.textContent || "?"}`),
    hint1: document.querySelector('sup[data-dwb-note="footnote:1"]').dataset.hintPl || "",
    readOnly: readOnlyMode,
  }));
  check("numeracja ciągła przez podział strony (1, 2, 3) i przypis końcowy „i”", read.labels.join() === "footnote:1=1,footnote:2=2,footnote:3=3,endnote:1=i", read.labels.join());
  check("numer na początku tekstu przypisu jak w Wordzie", read.marks.join() === "footnote:1=1,footnote:2=2,footnote:3=3,endnote:1=i", read.marks.join());
  check("dymek odnośnika = treść przypisu", read.hint1.startsWith("Brak ujęcia zakupu w Planie nie oznacza odmowy."), read.hint1);
  await page.locator('sup[data-dwb-note="footnote:2"]').click();
  await sleep(250);
  const jumped = await page.evaluate(() => ({
    flash: document.querySelector('ol.dwb-notes > li[data-dwb-note="footnote:2"]').classList.contains("note-flash"),
    back: !!document.querySelector(".link-back:not([hidden])"),
  }));
  check("Czytanie: klik w odnośnik → skok do przypisu + „Wróć”", jumped.flash && jumped.back, JSON.stringify(jumped));

  // ── Edycja: co jest edytowalne ──
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForFunction(() => !readOnlyMode && document.querySelector('ol.dwb-notes > li[data-dwb-note="footnote:1"] > p')?.isContentEditable, null, { timeout: 10000 });
  await sleep(300);
  const ed = await page.evaluate(() => ({
    refParas: [...document.querySelectorAll("sup[data-dwb-note]")].map((s) => s.closest("p").isContentEditable && !s.closest("p").dataset.lock),
    notes: Object.fromEntries([...document.querySelectorAll("ol.dwb-notes > li")].map((li) => [li.dataset.dwbNote, [...li.children].filter((c) => c.tagName === "P").map((p) => p.isContentEditable)])),
    lock3: document.querySelector('ol.dwb-notes > li[data-dwb-note="footnote:3"] > p')?.dataset.hintPl || "",
    edits: collectInlineParagraphEdits().length,
    split: [...document.querySelectorAll("section.docx > article p")].filter((p) => /Koniec pierwszej|Początek drugiej/.test(p.textContent)).map((p) => `${p.isContentEditable}|${!!p.dataset.dwbCont}|${(p.dataset.hintPl || "").includes("podział strony")}`),
  }));
  check("akapit z podziałem strony w środku: obie połówki tylko do odczytu, z wyjaśnieniem", ed.split.join() === "false|false|true,false|true|true", ed.split.join());
  check("akapity z odnośnikiem są edytowalne (dawniej tylko do odczytu)", ed.refParas.every(Boolean), JSON.stringify(ed.refParas));
  check("przypisy 1, 2 i końcowy edytowalne", ed.notes["footnote:1"]?.every(Boolean) && ed.notes["footnote:2"]?.every(Boolean) && ed.notes["endnote:1"]?.every(Boolean), JSON.stringify(ed.notes));
  check("przypis z tabelą tylko do odczytu, z wyjaśnieniem", ed.notes["footnote:3"] && !ed.notes["footnote:3"].some(Boolean) && ed.lock3.includes("Word"), JSON.stringify(ed.notes["footnote:3"]) + " " + ed.lock3);
  check("bez zmian zaraz po otwarciu (tabulator i odnośniki nie są „zmianą”)", ed.edits === 0, ed.edits);

  // ── pisanie ──
  const para1 = page.locator('sup[data-dwb-note="footnote:1"]').locator("xpath=ancestor::p[1]");
  await para1.evaluate((el) => placeCaret(el, "Plan zakupów".length));
  await page.keyboard.type(" rocznych");
  const para2 = page.locator('sup[data-dwb-note="footnote:2"]').locator("xpath=ancestor::p[1]");
  await para2.evaluate((el) => placeCaret(el, el.textContent.length));
  await page.keyboard.type(" (wpisz)");
  // klik myszą za odnośnikiem na końcu akapitu (Chrome stawiał kursor W numerze / na początku)
  const lastLine = await para2.evaluate((x) => { const r = document.createRange(); r.selectNodeContents(x); const rects = [...r.getClientRects()]; const l = rects[rects.length - 1]; return { x: l.right, y: l.top + l.height / 2 }; });
  await page.mouse.click(lastLine.x + 6, lastLine.y);
  await page.keyboard.type("!");
  const n1 = page.locator('ol.dwb-notes > li[data-dwb-note="footnote:1"] > p').first();
  await n1.evaluate((el) => placeCaret(el, el.textContent.length));
  await page.keyboard.type(" Uzupełnij uzasadnienie.");
  await n1.evaluate((el) => placeCaret(el, el.querySelector("sup.note-mark").textContent.length));
  await page.keyboard.press("Backspace");
  await page.keyboard.press("Backspace");
  const n2 = page.locator('ol.dwb-notes > li[data-dwb-note="footnote:2"] > p').first();
  await n2.evaluate((el) => placeCaret(el, el.textContent.indexOf("wpisz")));
  await page.keyboard.press("Enter");
  await page.keyboard.type("Nowy akapit: ");
  const para3 = page.locator('sup[data-dwb-note="endnote:1"]').locator("xpath=ancestor::p[1]");
  await para3.evaluate((el) => placeCaret(el, el.textContent.length));
  await page.keyboard.type(" Koniec.");
  const en = page.locator('ol.dwb-notes > li[data-dwb-note="endnote:1"] > p').first();
  await en.evaluate((el) => placeCaret(el, el.textContent.length));
  await page.keyboard.type(" Dopisek.");
  await sleep(300);
  const after = await page.evaluate(() => ({
    p2: document.querySelector('sup[data-dwb-note="footnote:2"]').closest("p").textContent.replace(/﻿/g, ""),
    mark1: !!document.querySelector('li[data-dwb-note="footnote:1"] sup.note-mark'),
    hint1: document.querySelector('sup[data-dwb-note="footnote:1"]').dataset.hintPl,
    n2paras: document.querySelectorAll('li[data-dwb-note="footnote:2"] > p').length,
    dirty: hasUnsavedChanges,
    edits: collectInlineParagraphEdits().map((e) => e.note || `p${e.index}`).sort().join(),
  }));
  check("pisanie za odnośnikiem na końcu akapitu (kursor i klik myszą)", after.p2.endsWith("z planu2 (wpisz)!"), after.p2);
  check("Backspace na początku przypisu nie kasuje numeru", after.mark1);
  check("dymek odnośnika pokazuje poprawiony tekst", after.hint1.includes("Uzupełnij uzasadnienie."), after.hint1);
  check("Enter w przypisie = nowy akapit przypisu", after.n2paras === 3, after.n2paras);
  check("zmiany w treści i w przypisach widoczne do zapisu", after.dirty && /endnote:1/.test(after.edits) && /footnote:1/.test(after.edits) && /footnote:2/.test(after.edits) && /p\d/.test(after.edits), after.edits);

  // ── zapis ──
  const saved = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    return { doc: await z.file("word/document.xml").async("string"), fn: await z.file("word/footnotes.xml").async("string"), en: await z.file("word/endnotes.xml").async("string") };
  });
  check("zapis: odnośniki w treści zostają (3 + 1)", count(saved.doc, /<w:footnoteReference /g) === 3 && count(saved.doc, /<w:endnoteReference /g) === 1);
  check("zapis: tabulator w akapicie zostaje <w:tab/>", count(saved.doc, /<w:r>(<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>)?<w:tab\/>/g) === 1, count(saved.doc, /<w:tab\/>/g));
  check("zapis: tekst dopisany przed odnośnikiem", /Plan zakupów rocznych<\/w:t><\/w:r><w:r[^>]*><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference w:id="1"\/>/.test(saved.doc), saved.doc.slice(saved.doc.indexOf("Plan zakupów"), saved.doc.indexOf("Plan zakupów") + 260));
  check("zapis: tekst za odnośnikiem na końcu akapitu", /<w:footnoteReference w:id="2"\/><\/w:r><w:r><w:t xml:space="preserve"> \(wpisz\)!<\/w:t>/.test(saved.doc));
  const f1 = noteTexts(saved.fn, "footnote", 1);
  const f2 = noteTexts(saved.fn, "footnote", 2);
  const f3 = noteTexts(saved.fn, "footnote", 3);
  const e1 = noteTexts(saved.en, "endnote", 1);
  check("zapis: akapit za podziałem strony dostaje SWÓJ tekst (numeracja akapitów się nie rozjeżdża)", /<w:endnoteReference w:id="1"\/><\/w:r><w:r><w:t(?: xml:space="preserve")?>\. Koniec\.<\/w:t>/.test(saved.doc) && saved.doc.includes('Koniec pierwszej części.</w:t></w:r><w:r><w:br w:type="page"/>'), saved.doc.slice(saved.doc.indexOf("Druga strona"), saved.doc.indexOf("Druga strona") + 700));
  check("zapis: przypis 1 z numerem i dopiskiem", JSON.stringify(f1) === JSON.stringify(["[nr] Brak ujęcia zakupu w Planie nie oznacza odmowy. Uzupełnij uzasadnienie."]), JSON.stringify(f1));
  check("zapis: przypis 2 podzielony Enterem (3 akapity)", JSON.stringify(f2) === JSON.stringify(["[nr] Jeżeli pozycji nie ma w Planie — ", "Nowy akapit: wpisz „brak”.", "Drugi akapit przypisu."]), JSON.stringify(f2));
  check("zapis: przypis z tabelą bez zmian", f3 && saved.fn.includes("<w:tbl>") && f3.join("|").includes("Kody finansowania"), JSON.stringify(f3));
  check("zapis: przypis końcowy z dopiskiem", JSON.stringify(e1) === JSON.stringify(["[nr] Uwaga końcowa dokumentu. Dopisek."]), JSON.stringify(e1));
  check("zapis: separatory przypisów nietknięte", count(saved.fn, /<w:separator\/>/g) === 1 && count(saved.fn, /<w:continuationSeparator\/>/g) === 1);
  check("zapis: nowy akapit przypisu ma styl przypisu", count(saved.fn.match(/<w:footnote w:id="2">[\s\S]*?<\/w:footnote>/)[0], /<w:pStyle w:val="FootnoteText"\/>/g) === 3);

  // ── przerysowanie (zmiana układu) — zmiany nie giną ──
  await page.evaluate(() => rerenderKeepingEdits());
  await page.waitForFunction(() => document.querySelector('li[data-dwb-note="footnote:1"]')?.textContent.includes("Uzupełnij"), null, { timeout: 15000 }).catch(() => {});
  await sleep(500);
  const rer = await page.evaluate(() => ({
    n1: document.querySelector('li[data-dwb-note="footnote:1"]')?.textContent || "",
    n2: document.querySelectorAll('li[data-dwb-note="footnote:2"] > p').length,
    body: document.querySelector('sup[data-dwb-note="footnote:1"]')?.closest("p").textContent || "",
    edits: collectInlineParagraphEdits().length,
  }));
  check("przerysowanie zachowuje zmiany w przypisach i treści", rer.n1.includes("Uzupełnij uzasadnienie.") && rer.n2 === 3 && rer.body.includes("rocznych") && rer.edits === 0, JSON.stringify(rer));

  // ── Cofnij / Ponów ──
  const n1b = page.locator('ol.dwb-notes > li[data-dwb-note="footnote:1"] > p').first();
  await n1b.evaluate((el) => placeCaret(el, el.textContent.length));
  await sleep(1300); // osobny krok cofania
  await page.keyboard.type(" XYZ");
  await sleep(300);
  await page.evaluate(() => dwbUndo.undo());
  await page.waitForFunction(() => !document.querySelector('li[data-dwb-note="footnote:1"]')?.textContent.includes("XYZ"), null, { timeout: 15000 }).catch(() => {});
  const undone = await page.evaluate(() => document.querySelector('li[data-dwb-note="footnote:1"]')?.textContent || "");
  await page.evaluate(() => dwbUndo.redo());
  await page.waitForFunction(() => document.querySelector('li[data-dwb-note="footnote:1"]')?.textContent.includes("XYZ"), null, { timeout: 15000 }).catch(() => {});
  const redone = await page.evaluate(() => document.querySelector('li[data-dwb-note="footnote:1"]')?.textContent || "");
  check("Cofnij zdejmuje pisanie w przypisie, Ponów przywraca", !undone.includes("XYZ") && undone.includes("Uzupełnij") && redone.includes("XYZ"), `${undone} | ${redone}`);

  // ── dwuklik w odnośnik (Edycja) → kursor w przypisie ──
  // numer „i” ma 3 px szerokości — kropka obok zachodzi na jego pole (WebKit); klik jak palec/mysz w środek
  await page.locator('sup[data-dwb-note="endnote:1"]').dblclick({ force: true });
  await sleep(400);
  const caretIn = await page.evaluate(() => document.activeElement?.closest?.("li")?.dataset.dwbNote || "");
  check("Edycja: dwuklik w odnośnik → kursor w tekście przypisu", caretIn === "endnote:1", caretIn);

  check("brak błędów JS", !errors.length, errors.slice(0, 3).join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`  ${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`);
  console.log(`\n${ENGINE}: ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
