// probe-list-format.js — czy PRAWDZIWY Word numeruje plik z aplikacji tak jak jej podgląd: format z
// biblioteki (I.), lista po przerwie, „Kontynuuj numerację”, „Numeruj od nowa od 1”, 1.1.1, a).
//   node tools/word-oracle/probe-list-format.js
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
  await p.locator("#fileInput").setInputFiles({ name: "l.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await buildDocx({ body: ["Alfa", "Beta", "Gamma", "Przerwa", "Delta", "Epsilon", "Zeta", "Eta", "Theta", "Przerwa 2", "Iota", "Kappa"] }) });
  const idle = () => p.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending).then(() => p.waitForTimeout(400));
  await idle();
  await p.evaluate(() => appFrame.setReadOnly(false));
  await idle();
  const ed = async (e) => { await p.evaluate((e) => applyDocumentEdit(e), e); await idle(); };
  await ed({ op: "list", indices: [0, 1, 2], kind: "number", toggle: false, variant: { fmt: "upperRoman", punct: "." } });
  await ed({ op: "list", indices: [4, 5, 6], kind: "number", toggle: false, variant: { fmt: "upperRoman", punct: "." } });
  await ed({ op: "listContinue", index: 4 });
  await ed({ op: "listRestart", index: 6 });
  await ed({ op: "list", indices: [7, 8], kind: "number", toggle: false, variant: { legal: true } });
  await p.evaluate(() => { const q = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[8]; q.closest(".docx-edit-root").focus(); placeCaret(q, 0); });
  await p.keyboard.press("Tab");
  await idle();
  await ed({ op: "list", indices: [10, 11], kind: "number", toggle: false, variant: { fmt: "lowerLetter", punct: ")" } });
  const app = await p.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((q) => (q.getAttribute("data-dwb-num") || "").trim()));
  const b64 = await p.evaluate(async () => { const x = await buildDocumentForSave(); let s = ""; x.forEach((v) => { s += String.fromCharCode(v); }); return btoa(s); });
  await b.close();
  const res = word.withDocument(Buffer.from(b64, "base64"), `repeat with i from 1 to 12
    set res to res & (list string of list format of text object of paragraph i) & "|"
  end repeat`);
  const wd = res.split("|").slice(0, 12).map((s) => s.trim());
  console.log("aplikacja:", JSON.stringify(app));
  console.log("Word:     ", JSON.stringify(wd));
  console.log(JSON.stringify(app) === JSON.stringify(wd) ? "ZGODNE" : "RÓŻNICA");
})();
