// probe-trailing-break.js — podział strony na samym końcu dokumentu: czy Word robi pustą stronę?
// (Sprawozdanie_zadanie_3: ostatni akapit = sam w:br page; Word 4 strony, aplikacja 5).
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const { buildDocx } = require("./build-docx");
const { withDocument, DIR } = require("./word");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
const T = (t) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
const BR = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
const CASES = {
  brOwnLast: [T("Strona 1"), BR],
  brInLast: [`<w:p><w:r><w:t>Strona 1</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>`],
  brThenEmpty: [T("Strona 1"), BR, "<w:p/>"],
  brThenText: [T("Strona 1"), BR, T("Strona 2")],
  pbbEmptyLast: [T("Strona 1"), '<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'],
};
(async () => {
  for (const [k, body] of Object.entries(CASES)) {
    const pdf = path.join(DIR, `tbr-${k}.pdf`);
    try { fs.unlinkSync(pdf); } catch (_) {}
    withDocument(await buildDocx({ body }), `save as it file name "${esc(pdf)}" file format format PDF`);
    console.log(k.padEnd(14), "stron:", execFileSync("pdfinfo", [pdf]).toString().match(/Pages:\s+(\d+)/)[1]);
  }
})();
