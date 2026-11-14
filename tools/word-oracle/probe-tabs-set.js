// probe-tabs-set.js — czy PRAWDZIWY Word widzi tabulatory zapisane oknem „Tabulatory…” aplikacji
// (nowy prawy z kropkami; tabulator stylu wyłączony „clear” + nowy środkowy). node tools/word-oracle/probe-tabs-set.js
const pw = require("playwright");
const word = require("./word");
const { buildDocx } = require("./build-docx");
const { APP_URL } = require("../../scripts/docx-test-helpers");

(async () => {
  const styles = `<w:style w:type="paragraph" w:customStyle="1" w:styleId="Formularz"><w:name w:val="Formularz"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="left" w:leader="dot" w:pos="5000"/></w:tabs></w:pPr></w:style>`;
  const b = await pw.chromium.launch();
  const c = await b.newContext({ serviceWorkers: "block" });
  await c.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const p = await c.newPage();
  await p.goto(APP_URL);
  await p.evaluate(() => document.getElementById("heroSplash")?.remove());
  await p.locator("#fileInput").setInputFiles({ name: "t.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await buildDocx({ styles, body: ["Imię", { style: "Formularz", text: "Pole" }] }) });
  const idle = () => p.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending).then(() => p.waitForTimeout(400));
  await idle();
  await p.evaluate(() => appFrame.setReadOnly(false));
  await idle();
  await p.evaluate(() => applyDocumentEdit({ op: "tabsSet", indices: [0], stops: [{ pos: 7087, align: "right", leader: "dot" }] }));
  await idle();
  await p.evaluate(() => applyDocumentEdit({ op: "tabsSet", indices: [1], stops: [{ pos: 1701, align: "center", leader: "none" }] }));
  await idle();
  const b64 = await p.evaluate(async () => { const x = await buildDocumentForSave(); let s = ""; x.forEach((v) => { s += String.fromCharCode(v); }); return btoa(s); });
  await b.close();
  const { parts } = word.saveDocx(Buffer.from(b64, "base64"), "", ["word/document.xml"]);
  const paras = (parts["word/document.xml"].match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).map((x) => (x.match(/<w:tabs>[\s\S]*?<\/w:tabs>/) || ["(bez w:tabs)"])[0]);
  console.log("Word po zapisie:\n akapit 1: " + paras[0] + "\n akapit 2: " + paras[1]);
})();
