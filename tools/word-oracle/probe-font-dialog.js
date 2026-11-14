// probe-font-dialog.js — czy PRAWDZIWY Word czyta cechy z okna „Czcionka…” aplikacji: kapitaliki,
// wersaliki, odstęp znaków, położenie, podwójne podkreślenie.   node tools/word-oracle/probe-font-dialog.js
const pw = require("playwright");
const word = require("./word");
const { buildDocx } = require("./build-docx");
const { APP_URL } = require("../../scripts/docx-test-helpers");

(async () => {
  const b = await pw.chromium.launch();
  const c = await b.newContext({ serviceWorkers: "block" });
  await c.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const p = await c.newPage();
  await p.goto(APP_URL);
  await p.evaluate(() => document.getElementById("heroSplash")?.remove());
  await p.locator("#fileInput").setInputFiles({ name: "f.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await buildDocx({ body: ["Alfa Beta Gamma"] }) });
  const idle = () => p.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending).then(() => p.waitForTimeout(400));
  await idle();
  await p.evaluate(() => appFrame.setReadOnly(false));
  await idle();
  await p.evaluate(() => applyDocumentEdit({ op: "runStyle", ranges: [{ index: 0, start: 0, end: 4 }], font: { smallCaps: true, spacing: 1.5, underline: true, uStyle: "double" } }));
  await idle();
  await p.evaluate(() => applyDocumentEdit({ op: "runStyle", ranges: [{ index: 0, start: 5, end: 9 }], font: { caps: true, position: 3 } }));
  await idle();
  const b64 = await p.evaluate(async () => { const x = await buildDocumentForSave(); let s = ""; x.forEach((v) => { s += String.fromCharCode(v); }); return btoa(s); });
  await b.close();
  const bytes = Buffer.from(b64, "base64");
  console.log(word.withDocument(bytes, `repeat with w in {1, 2}
    set f to font object of (word w of text object of paragraph 1)
    set res to res & (content of word w of text object of paragraph 1) & ": kapitaliki=" & (small caps of f) & " wersaliki=" & (all caps of f) & " odstęp=" & (spacing of f) & " podkreślenie=" & (underline of f as string) & linefeed
  end repeat`));
  // położenie (w:position): AppleScript go nie podaje — z pliku zapisanego przez Worda
  const { parts } = word.saveDocx(bytes, "", ["word/document.xml"]);
  console.log("Word po zapisie (rPr):", (parts["word/document.xml"].match(/<w:rPr>[\s\S]*?<\/w:rPr>/g) || []).join(" "));
})();
