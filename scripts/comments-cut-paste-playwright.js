// comments-cut-paste-playwright.js — wycięty cały komentowany fragment ma przenieść kotwicę
// komentarza razem z tekstem, zamiast zostawić po wklejeniu sam tekst.

const path = require("path");
const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1120, height: 800 } });
  await context.addInitScript(() => {
    sessionStorage.setItem("introPlayed", "true");
    localStorage.setItem("dwb.authorName", "Test komentarza");
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });

  await page.keyboard.type("Początek zaznaczony fragment koniec");
  await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0];
    const from = p.textContent.indexOf("zaznaczony");
    const r = formDomRange(p, from, from + "zaznaczony".length);
    p.closest(".docx-edit-root").focus();
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  });
  await page.keyboard.press(`${MOD}+Alt+KeyM`);
  await page.waitForSelector(".compose-pop-comment .cf-rich");
  await page.keyboard.type("Przenoszony komentarz");
  await page.click(".compose-pop-form .lf-ok");
  await page.waitForTimeout(500);

  const result = await page.evaluate(async () => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0];
    // Kopiujemy CAŁE zdanie, chociaż komentarz był tylko na jednym słowie.
    // To odtwarza przypadek, w którym zakres komentarza rozszerzał się po wklejeniu.
    const r = formDomRange(p, 0, p.textContent.length);
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    const clip = new DataTransfer();
    docEditRoot().dispatchEvent(new ClipboardEvent("copy", { clipboardData: clip, bubbles: true, cancelable: true }));
    const atEnd = document.createRange();
    atEnd.selectNodeContents(p);
    atEnd.collapse(false);
    getSelection().removeAllRanges();
    getSelection().addRange(atEnd);
    docEditRoot().dispatchEvent(new ClipboardEvent("paste", { clipboardData: clip, bubbles: true, cancelable: true }));
    await new Promise((r) => setTimeout(r, 400)); // nowy numer kopii komentarza (w tle)
    await waitInlineStructuralIdle();
    const bytes = await buildDocumentForSave();
    window.__savedCopy = bytes;
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("word/document.xml").async("string");
    const comments = await zip.file("word/comments.xml").async("string");
    const ranges = [...document.querySelectorAll("span[data-cm-kind='start']")].map((start) => {
      const end = [...document.querySelectorAll(`span[data-cm-kind="end"][data-cm-id="${CSS.escape(start.dataset.cmId)}"]`)]
        .find((candidate) => !!(start.compareDocumentPosition(candidate) & Node.DOCUMENT_POSITION_FOLLOWING));
      if (!end) return { id: start.dataset.cmId, text: "" };
      const range = document.createRange();
      range.setStartAfter(start);
      range.setEndBefore(end);
      return { id: start.dataset.cmId, text: range.toString() };
    });
    const highlight = CSS.highlights?.get("dwb-comment");
    return {
      text: p.textContent,
      starts: (xml.match(/<w:commentRangeStart\b/g) || []).length,
      ends: (xml.match(/<w:commentRangeEnd\b/g) || []).length,
      refs: (xml.match(/<w:commentReference\b/g) || []).length,
      refIds: [...xml.matchAll(/<w:commentReference\b[^>]*w:id="(\d+)"/g)].map((m) => m[1]),
      copies: (comments.match(/Przenoszony komentarz/g) || []).length,
      comment: comments.includes("Przenoszony komentarz"),
      ranges,
      highlightRanges: highlight ? highlight.size : null,
    };
  });

  const moved = result.ranges.length === 2 && result.ranges.every((range) => range.text === "zaznaczony");
  if (result.starts !== 2 || result.ends !== 2 || result.refs !== 2 || !result.comment || !moved
    || (result.highlightRanges !== null && result.highlightRanges < 1)) {
    throw new Error(`Komentarz zgubiony po Wytnij/Wklej: ${JSON.stringify(result)}`);
  }
  // Kopiuj/Wklej: kopia komentarza z NOWYM numerem i własnym wpisem (ten sam numer dwa razy =
  // „Word odnalazł zawartość, której nie można odczytać”, zgłoszenie 2026-10-12)
  if (new Set(result.refIds).size !== 2 || result.copies !== 2 || new Set(result.ranges.map((r) => r.id)).size !== 2) {
    throw new Error(`Kopia komentarza z tym samym numerem: ${JSON.stringify(result)}`);
  }
  await validateSaved(page, "kopia komentarza");

  // Przerysowanie dokumentu między Wytnij a Wklej: klucze znaczników z wycinka wskazywały potem
  // znaczniki INNEGO komentarza — wklejone zdanie dostawało cudzy komentarz, a plik powtórzony numer.
  const rer = await rerenderCase(page);
  if (rer.refIds.length !== 2 || new Set(rer.refIds).size !== 2 || rer.onPasted !== "Komentarz B" || rer.onOther !== "Komentarz A") {
    throw new Error(`Po przerysowaniu między Wytnij a Wklej komentarz pomylony: ${JSON.stringify(rer)}`);
  }
  await validateSaved(page, "przerysowanie między Wytnij a Wklej");

  // Plik, który już ma powtórzony numer komentarza: pierwszy zapis po zmianie go naprawia
  const dd = await duplicateFileCase(page);
  if (new Set(dd.refIds).size !== dd.refIds.length || dd.comments !== dd.refIds.length) {
    throw new Error(`Powtórzony numer komentarza nie naprawiony przy zapisie: ${JSON.stringify(dd)}`);
  }
  await validateSaved(page, "naprawa powtórzonego numeru");

  // Zgłoszenie 2026-10-11: komentarz do KOŃCA akapitu („…broszury "Czy życie…", strony 12 i 13”).
  // Koniec i odwołanie komentarza stoją za ostatnią literą, a zaznaczenie kończy się w tekście
  // tuż przed nimi. Wycinek brał sam początek komentarza; wklejony wyżej łączył się z końcem
  // w starym miejscu i podświetlenie szło przez kilka akapitów.
  const whole = await cutPasteCase(page, (text) => [0, text.length], "Czy życie stworzone");
  if (whole.starts !== 1 || whole.ends !== 1 || whole.ranges.length !== 1 || whole.ranges[0].text !== "Czy życie stworzone"
    || whole.ranges[0].para !== 0 || whole.second !== "") {
    throw new Error(`Komentarz do końca akapitu nie przeszedł w całości przy Wytnij/Wklej: ${JSON.stringify(whole)}`);
  }

  // Wycięcie tylko części komentowanego tekstu: komentarz zostaje w starym miejscu (krótszy),
  // a wklejony tekst nie niesie połowy znaczników.
  const part = await cutPasteCase(page, (text) => [0, text.indexOf(" stworzone")], "Czy życie stworzone");
  if (part.starts !== 1 || part.ends !== 1 || part.ranges.length !== 1 || part.ranges[0].text !== " stworzone"
    || part.ranges[0].para !== 1 || !part.first.endsWith("Wskazówka z broszury Czy życie")) {
    throw new Error(`Częściowe wycięcie komentarza rozjechało zakres: ${JSON.stringify(part)}`);
  }

  // Przebudowa pliku między Wytnij a Wklej: komentarz ma przejść z treścią (dawniej zostawało samo
  // podświetlenie — kliknięcie nic nie otwierało, a w zapisanym pliku komentarza nie było).
  const merged = await cutPasteCase(page, (text) => [0, text.length], "Czy życie stworzone", { mergeBetween: true });
  const mergedId = merged.ranges[0]?.id;
  if (merged.ranges.length !== 1 || merged.ranges[0].text !== "Czy życie stworzone" || !merged.known.includes(mergedId) || !merged.savedComment) {
    throw new Error(`Komentarz zgubiony, gdy plik przebudował się między Wytnij a Wklej: ${JSON.stringify(merged)}`);
  }
  // Wytnij bez wklejenia + zapis: komentarza nie ma w pliku (jak w Wordzie), żadnych resztek
  const dropped = await cutPasteCase(page, (text) => [0, text.length], "Czy życie stworzone", { mergeBetween: true, noPaste: true });
  if (dropped.starts !== 0 || dropped.ends !== 0 || dropped.savedComment) {
    throw new Error(`Wycięty i niewklejony komentarz został w zapisanym pliku: ${JSON.stringify(dropped)}`);
  }

  // To samo na pliku z Worda: tekst siedzi w <span> przebiegu, więc koniec zaznaczenia w DOM wypada
  // wewnątrz spanu, przed znacznikiem końca komentarza (stary kod uznawał komentarz za częściowy).
  const file = await wordFileCase(page);
  if (file.length !== 1 || file[0].text !== "Ostatni akapit." || file[0].para !== file[0].endPara || file[0].para !== file[0].dst) {
    throw new Error(`Komentarz z pliku Worda rozjechał się po Wytnij/Wklej: ${JSON.stringify(file)}`);
  }

  await browser.close();
  if (errors.length) throw new Error(errors.join("\n"));
  console.log(`✅ comments-cut-paste-playwright passed [${ENGINE}]`);
}

// Nowy dokument: „Pierwszy akapit” + „Wskazówka z broszury <komentarz>” (komentarz do końca).
// Wycina zakres [od, do) z drugiego akapitu (koniec zaznaczenia w tekście, jak przy przeciąganiu)
// i wkleja na końcu pierwszego. Zwraca znaczniki w pliku i zakresy komentarzy w podglądzie.
async function cutPasteCase(page, pick, commented, opts = {}) {
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await page.waitForTimeout(300);
  await page.keyboard.type("Pierwszy akapit");
  await page.keyboard.press("Enter");
  await page.keyboard.type(`Wskazówka z broszury ${commented}`);
  await page.evaluate((word) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1];
    const from = p.textContent.indexOf(word);
    const r = formDomRange(p, from, from + word.length);
    p.closest(".docx-edit-root").focus();
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  }, commented);
  await page.keyboard.press(`${process.env.ENGINE === "webkit" ? "Meta" : "Control"}+Alt+KeyM`);
  await page.waitForSelector(".compose-pop-comment .cf-rich");
  await page.keyboard.type("Tajemnice płetw");
  await page.click(".compose-pop-form .lf-ok");
  await page.waitForTimeout(500);

  return page.evaluate(async ([pickSrc, opts]) => {
    const paras = () => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const p = paras()[1];
    // eslint-disable-next-line no-new-func
    const [from, to] = new Function(`return (${pickSrc})`)()(p.textContent);
    const texts = [];
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) if (n.data.length) texts.push(n);
    // punkt w tekście: koniec zaznaczenia zostaje W węźle tekstowym (przed znacznikiem końca)
    const point = (at, preferEnd) => {
      let left = at;
      for (const n of texts) {
        if (left < n.data.length || (preferEnd && left === n.data.length)) return [n, left];
        left -= n.data.length;
      }
      return [texts[texts.length - 1], texts[texts.length - 1].data.length];
    };
    const r = document.createRange();
    r.setStart(...point(from, false));
    r.setEnd(...point(to, true));
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    const clip = new DataTransfer();
    docEditRoot().dispatchEvent(new ClipboardEvent("cut", { clipboardData: clip, bubbles: true, cancelable: true }));
    // przebudowa pliku między Wytnij a Wklej (Enter, wycięcie przez akapity, panel) — kasowała
    // komentarz „bez kotwicy”, choć był w schowku (zgłoszenie 2026-10-12)
    if (opts.mergeBetween) await mergeInlineEditsIntoBytes();
    const first = paras()[0];
    const atEnd = document.createRange();
    atEnd.selectNodeContents(first);
    atEnd.collapse(false);
    getSelection().removeAllRanges();
    getSelection().addRange(atEnd);
    if (!opts.noPaste) docEditRoot().dispatchEvent(new ClipboardEvent("paste", { clipboardData: clip, bubbles: true, cancelable: true }));
    composeUi.paintCommentHighlights();
    const known = [...(await composeUi.loadComments()).keys()];
    const bytes = await buildDocumentForSave();
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("word/document.xml").async("string");
    const commentsXml = await zip.file("word/comments.xml")?.async("string") || "";
    const list = paras();
    const ranges = [...document.querySelectorAll("span[data-cm-kind='start']")].map((start) => {
      const end = [...document.querySelectorAll(`span[data-cm-kind="end"][data-cm-id="${CSS.escape(start.dataset.cmId)}"]`)]
        .find((candidate) => !!(start.compareDocumentPosition(candidate) & Node.DOCUMENT_POSITION_FOLLOWING));
      if (!end) return { id: start.dataset.cmId, text: null };
      const range = document.createRange();
      range.setStartAfter(start);
      range.setEndBefore(end);
      return { id: start.dataset.cmId, text: range.toString(), para: list.findIndex((x) => x.contains(start)), endPara: list.findIndex((x) => x.contains(end)) };
    });
    return {
      first: list[0]?.textContent,
      second: list[1]?.textContent,
      starts: (xml.match(/<w:commentRangeStart\b/g) || []).length,
      ends: (xml.match(/<w:commentRangeEnd\b/g) || []).length,
      ranges,
      known,
      savedComment: commentsXml.includes("Tajemnice płetw"),
    };
  }, [pick.toString(), opts]);
}

// review-sample.docx: „Ostatni akapit.” z komentarzem 20 na całym tekście (koniec + odwołanie za
// tekstem). Wycinamy sam tekst akapitu i wklejamy na końcu akapitu „Kara umowna…”.
async function wordFileCase(page) {
  await page.locator("#fileInput").setInputFiles(path.join(__dirname, "..", "docs", "samples", "review-sample.docx"));
  await page.waitForFunction(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host") || document.body)
    .some((p) => p.textContent.startsWith("Ostatni akapit")), null, { timeout: 30000 });
  await page.evaluate(() => {
    if (!readModeEl.checked) return;
    readModeEl.checked = false;
    readModeEl.dispatchEvent(new Event("change", { bubbles: true }));
  });
  await page.waitForSelector(".docx-edit-root");
  await page.waitForTimeout(300);
  return page.evaluate(() => {
    const paras = () => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const src = paras().find((p) => p.textContent.startsWith("Ostatni akapit"));
    const dst = paras().find((p) => p.textContent.startsWith("Kara umowna"));
    const texts = [];
    const walker = document.createTreeWalker(src, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) if (n.data.replace(/\uFEFF/g, "").length) texts.push(n);
    const r = document.createRange();
    r.setStart(texts[0], 0);
    r.setEnd(texts[texts.length - 1], texts[texts.length - 1].data.length);
    src.closest(".docx-edit-root").focus();
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    const clip = new DataTransfer();
    docEditRoot().dispatchEvent(new ClipboardEvent("cut", { clipboardData: clip, bubbles: true, cancelable: true }));
    const atEnd = document.createRange();
    atEnd.selectNodeContents(dst);
    atEnd.collapse(false);
    getSelection().removeAllRanges();
    getSelection().addRange(atEnd);
    docEditRoot().dispatchEvent(new ClipboardEvent("paste", { clipboardData: clip, bubbles: true, cancelable: true }));
    const list = paras();
    return [...document.querySelectorAll("span[data-cm-kind='start'][data-cm-id='20']")].map((start) => {
      const end = [...document.querySelectorAll("span[data-cm-kind='end'][data-cm-id='20']")]
        .find((candidate) => !!(start.compareDocumentPosition(candidate) & Node.DOCUMENT_POSITION_FOLLOWING));
      const range = document.createRange();
      range.setStartAfter(start);
      if (end) range.setEndBefore(end);
      return { text: end ? range.toString() : null, para: list.indexOf(start.closest("p")), endPara: end ? list.indexOf(end.closest("p")) : -1, dst: list.indexOf(dst) };
    });
  });
}

// Zapisany plik (window.__savedCopy) przez walidator Open XML SDK, gdy jest .NET (test:schema).
async function validateSaved(page, label) {
  const v = require("./ooxml-validate");
  if (!v.available()) return;
  const b64 = await page.evaluate(() => { let s = ""; window.__savedCopy.forEach((x) => (s += String.fromCharCode(x))); return btoa(s); });
  const file = require("path").join(require("os").tmpdir(), `dwb-cmt-${process.pid}.docx`);
  require("fs").writeFileSync(file, Buffer.from(b64, "base64"));
  const r = v.validate([file])[0];
  require("fs").rmSync(file, { force: true });
  if (!r.ok) throw new Error(`Plik po „${label}” nie przechodzi walidacji: ${JSON.stringify(r.errors).slice(0, 600)}`);
}

// Dwa komentarze: B (wycinany, wcześniej w dokumencie) i A. Wytnij B → zapis do pliku i
// przerysowanie (klucze znaczników od nowa) → Wklej. B ma zostać B, A — A.
async function rerenderCase(page) {
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await page.waitForTimeout(300);
  await page.keyboard.type("Alfa beta");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Gamma delta");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Cel:");
  for (const [pi, word, text] of [[0, "beta", "Komentarz B"], [1, "delta", "Komentarz A"]]) {
    await page.evaluate(([pi, word]) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[pi];
      const from = p.textContent.indexOf(word);
      const r = formDomRange(p, from, from + word.length);
      p.closest(".docx-edit-root").focus();
      getSelection().removeAllRanges();
      getSelection().addRange(r);
    }, [pi, word]);
    await page.keyboard.press(`${MOD}+Alt+KeyM`);
    await page.waitForSelector(".compose-pop-comment .cf-rich");
    await page.keyboard.type(text);
    await page.click(".compose-pop-form .lf-ok");
    await page.waitForTimeout(600);
  }
  await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0];
    const from = p.textContent.indexOf("beta");
    const r = formDomRange(p, from, from + 4);
    // całe słowo razem ze znacznikami komentarza
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    window.__clip = new DataTransfer();
    docEditRoot().dispatchEvent(new ClipboardEvent("cut", { clipboardData: window.__clip, bubbles: true, cancelable: true }));
  });
  await page.evaluate(async () => { await mergeInlineEditsIntoBytes(); await reloadFromBytes(originalFileBytes); });
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await page.waitForTimeout(500);
  return page.evaluate(async () => {
    const list = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const target = list[list.length - 1];
    target.closest(".docx-edit-root").focus();
    placeCaret(target, Number.MAX_SAFE_INTEGER);
    docEditRoot().dispatchEvent(new ClipboardEvent("paste", { clipboardData: window.__clip, bubbles: true, cancelable: true }));
    await new Promise((r) => setTimeout(r, 400));
    await waitInlineStructuralIdle();
    const bytes = await buildDocumentForSave();
    window.__savedCopy = bytes;
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("word/document.xml").async("string");
    const cx = await zip.file("word/comments.xml").async("string");
    const textOf = (id) => { const m = cx.match(new RegExp(`<w:comment\\b[^>]*w:id="${id}"[\\s\\S]*?</w:comment>`)); return m ? m[0].replace(/<[^>]+>/g, "") : null; };
    const paras = [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map((m) => m[0]);
    const idIn = (needle) => { const p = paras.find((q) => q.replace(/<[^>]+>/g, "").includes(needle)); return p?.match(/commentReference w:id="(\d+)"/)?.[1]; };
    return {
      refIds: [...xml.matchAll(/<w:commentReference\b[^>]*w:id="(\d+)"/g)].map((m) => m[1]),
      onPasted: textOf(idIn("Cel:")),
      onOther: textOf(idIn("Gamma")),
    };
  });
}

// review-sample.docx z komentarzem 10 powtórzonym w drugim akapicie (tak wyglądał plik ze
// zgłoszenia). Po zmianie i zapisie: numery jedyne, kopia z własnym wpisem.
async function duplicateFileCase(page) {
  const fs = require("fs");
  const os = require("os");
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(__dirname, "..", "docs", "samples", "review-sample.docx")));
  let xml = await zip.file("word/document.xml").async("string");
  const dupPara = '<w:p><w:r><w:t xml:space="preserve">Kopia: </w:t></w:r><w:commentRangeStart w:id="10"/><w:r><w:t>skopiowany fragment</w:t></w:r><w:commentRangeEnd w:id="10"/><w:r><w:commentReference w:id="10"/></w:r></w:p>';
  xml = xml.replace(/(<w:sectPr)/, `${dupPara}$1`);
  zip.file("word/document.xml", xml);
  const file = path.join(os.tmpdir(), `dwb-dup-comment-${process.pid}.docx`);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer" }));
  await page.locator("#fileInput").setInputFiles(file);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && collectPreviewParagraphElements(document.querySelector(".docx-preview-host") || document.body).some((p) => p.textContent.startsWith("Kopia:")), null, { timeout: 30000 });
  await page.evaluate(() => { if (readModeEl.checked) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change", { bubbles: true })); } });
  await page.waitForSelector(".docx-edit-root");
  await page.waitForTimeout(300);
  const out = await page.evaluate(async () => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((q) => q.textContent.startsWith("Kopia:"));
    p.closest(".docx-edit-root").focus();
    placeCaret(p, 0);
    document.execCommand("insertText", false, "X");
    const bytes = await buildDocumentForSave();
    window.__savedCopy = bytes;
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("word/document.xml").async("string");
    const cx = await zip.file("word/comments.xml").async("string");
    return { refIds: [...xml.matchAll(/<w:commentReference\b[^>]*w:id="(\d+)"/g)].map((m) => m[1]), comments: (cx.match(/<w:comment\b/g) || []).length };
  });
  fs.rmSync(file, { force: true });
  return out;
}

run().catch((error) => {
  console.error("❌ comments-cut-paste-playwright failed:", error.message || error);
  process.exit(1);
});
