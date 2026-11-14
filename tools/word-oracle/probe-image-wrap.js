// probe-image-wrap.js — co zapisuje PRAWDZIWY Word, gdy obraz „W linii z tekstem” dostaje zawijanie
// (Przyległe / Za tekstem / Przed tekstem / Góra i dół): wp:anchor, położenie, odstępy, behindDoc.
//   node tools/word-oracle/probe-image-wrap.js
const fs = require("fs");
const path = require("path");
const word = require("./word");

const orig = fs.readFileSync(path.join(__dirname, "../../docs/samples/przewodnik.docx"));
for (const wrap of ["wrap square", "wrap behind", "wrap front", "wrap top bottom", "wrap tight"]) {
  try {
    const { parts } = word.saveDocx(orig, `
      convert to shape (inline picture 1)
      set n to count of shapes
      set wrap type of wrap format of shape n to ${wrap}`, ["word/document.xml"]);
    const seg = (parts["word/document.xml"].split("<wp:anchor").slice(1).find((x) => x.split("</wp:anchor>")[0].includes("pic:pic")) || "");
    const a = seg ? `<wp:anchor${seg.split("<wp:docPr")[0]}` : "(brak wp:anchor z obrazem)";
    console.log(`── ${wrap}\n${a}\n`);
  } catch (e) { console.log(`── ${wrap}: ${e.message}\n`); }
}
