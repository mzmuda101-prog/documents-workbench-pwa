// probe-section-app.js — czy PRAWDZIWY Word widzi plik z podziałem sekcji z aplikacji tak jak jej podgląd:
// liczba sekcji, orientacja, nagłówek drugiej sekcji (dziedziczony), „Połącz z poprzednim”.
//   node tools/word-oracle/probe-section-app.js
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
  await p.evaluate(() => applyDocumentEdit({ op: "pageSetup", index: 2, scope: "section", orient: "landscape" }));
  await idle();
  const b64 = await p.evaluate(async () => { const x = await buildDocumentForSave(); let s = ""; x.forEach((v) => { s += String.fromCharCode(v); }); return btoa(s); });
  await b.close();
  console.log(word.withDocument(Buffer.from(b64, "base64"), `set n to count of sections
    set res to "sekcje=" & n
    repeat with i from 1 to n
      set sc to section i
      set res to res & " | " & i & ": orient=" & (orientation of page setup of sc as string) & " nagłówek=" & (content of text object of (get header sc index header footer primary))
      if i > 1 then set res to res & " połączony=" & (link to previous of (get header sc index header footer primary))
    end repeat`));
})();
