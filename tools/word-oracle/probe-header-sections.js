// probe-header-sections.js — czy PRAWDZIWY Word widzi nagłówki sekcji z aplikacji tak jak jej podgląd:
// sekcja 2 rozłączona z własnym nagłówkiem, sekcja 1 z nagłówkiem stron parzystych (inne strony
// parzyste i nieparzyste — ustawienie dokumentu), „Połącz z poprzednim” sekcji 2 wyłączone.
//   node tools/word-oracle/probe-header-sections.js
const pw = require("playwright");
const word = require("./word");
const { buildDocx } = require("./build-docx");
const { APP_URL } = require("../../scripts/docx-test-helpers");
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

(async () => {
  const bytes = await buildDocx({ body: ["Pierwszy.", "Drugi do podziału.", "Trzeci."], sectPr: '<w:headerReference w:type="default" r:id="rIdX0"/>', extraParts: { "header1.xml": { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", rel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p><w:r><w:t>NAGŁÓWEK FIRMY</w:t></w:r></w:p></w:hdr>` } } });
  const b = await pw.chromium.launch();
  const c = await b.newContext({ serviceWorkers: "block" });
  await c.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const p = await c.newPage();
  await p.goto(APP_URL);
  await p.evaluate(() => document.getElementById("heroSplash")?.remove());
  await p.locator("#fileInput").setInputFiles({ name: "s.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: bytes });
  const idle = () => p.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending).then(() => p.waitForTimeout(500));
  await idle();
  await p.evaluate(() => appFrame.setReadOnly(false));
  await idle();
  await p.evaluate(() => applyDocumentEdit({ op: "sectionBreak", index: 1, kind: "nextPage", mode: "split", beforeRuns: [{ text: "Drugi " }], afterRuns: [{ text: "do podziału." }] }));
  await idle();
  // sekcja 2: rozłączona, własny nagłówek; sekcja 1: inne strony parzyste
  await p.evaluate(() => applyDocumentEdit({ op: "headerFooter", lang: "pl", total: 2, section: 1, linked: false, header: { text: "ROZDZIAŁ 2", align: "left" }, footer: { text: "", align: "left" }, number: { fmt: null }, firstDifferent: false, evenOdd: false }));
  await idle();
  await p.evaluate(() => applyDocumentEdit({ op: "headerFooter", lang: "pl", total: 2, section: 0, linked: false, header: { text: "NAGŁÓWEK FIRMY", align: "left" }, footer: { text: "", align: "left" }, number: { fmt: null }, firstDifferent: false, evenOdd: true, even: { header: { text: "PARZYSTE", align: "right" }, footer: { text: "", align: "left" } } }));
  await idle();
  const b64 = await p.evaluate(async () => { const x = await buildDocumentForSave(); let s = ""; x.forEach((v) => { s += String.fromCharCode(v); }); return btoa(s); });
  await b.close();
  console.log(word.withDocument(Buffer.from(b64, "base64"), `set n to count of sections
    set res to "sekcje=" & n & " parzyste=" & (odd and even pages header footer of page setup of section 1 as string)
    repeat with i from 1 to n
      set sc to section i
      set res to res & " | " & i & ": nagłówek=" & (content of text object of (get header sc index header footer primary)) & " parzysty=" & (content of text object of (get header sc index header footer even pages))
      if i > 1 then set res to res & " połączony=" & (link to previous of (get header sc index header footer primary))
    end repeat`));
})();
