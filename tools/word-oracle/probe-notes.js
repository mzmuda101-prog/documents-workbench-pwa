// probe-notes.js — co Word mówi o numerach przypisów (note reference / PDF)
const { buildDocx, W } = require("./build-docx");
const word = require("./word");
const fs = require("fs");
const path = require("path");
const FN = (kind, id, text) => `<w:${kind} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="${kind === "footnote" ? "FootnoteText" : "EndnoteText"}"/></w:pPr><w:r><w:rPr><w:rStyle w:val="${kind === "footnote" ? "FootnoteReference" : "EndnoteReference"}"/></w:rPr><w:${kind}Ref/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:${kind}>`;
const part = (kind, notes) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind}s ${W}><w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>${notes}</w:${kind}s>`;
const ref = (kind, id, text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${kind}Reference w:id="${id}"/></w:r></w:p>`;
(async () => {
  const bytes = await buildDocx({
    body: [ref("footnote", 1, "A"), ref("footnote", 2, "B"), `<w:p><w:pPr><w:sectPr><w:footnotePr><w:numFmt w:val="lowerRoman"/></w:footnotePr><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr></w:p>`, ref("footnote", 3, "C"), ref("endnote", 1, "D")],
    sectPr: '<w:footnotePr><w:numFmt w:val="upperLetter"/><w:numRestart w:val="eachSect"/></w:footnotePr>',
    extraParts: {
      "footnotes.xml": { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", rel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes", xml: part("footnote", FN("footnote", 1, "n1") + FN("footnote", 2, "n2") + FN("footnote", 3, "n3")) },
      "endnotes.xml": { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml", rel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes", xml: part("endnote", FN("endnote", 1, "e1")) },
    },
  });
  fs.writeFileSync(path.join(word.DIR, "notes-probe.docx"), bytes);
  const pdf = path.join(word.DIR, "_out.pdf");
  try { fs.unlinkSync(pdf); } catch (_) {}
  const out = word.withDocument(bytes, `
    set n to count of footnotes
    repeat with i from 1 to n
      set res to res & "F" & i & "=" & (content of note reference of footnote i) & "|"
    end repeat
    set res to res & (footnote number style of footnote options of text object) & "|"
    save as it file name "${pdf}" file format format PDF
  `);
  console.log(JSON.stringify(out));
  console.log(fs.existsSync(pdf));
})();
