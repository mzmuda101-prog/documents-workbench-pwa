// probe-center-pdf.js — PDF-y z Worda z wyśrodkowanymi akapitami (lista 3.11: wielowierszowy
// wyśrodkowany akapit z PDF dzielony na akapity). Zapisuje output/word-oracle/center-*.pdf.
const fs = require("fs");
const path = require("path");
const { buildDocx } = require("./build-docx");
const { withDocument, DIR } = require("./word");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
const LONG = "Niniejsze zarządzenie określa zasady przyznawania świadczeń dla pracowników uczelni, tryb składania wniosków oraz sposób ich rozpatrywania przez komisję powołaną przez rektora na okres kadencji.";
const para = (pPr, t) => `<w:p><w:pPr><w:jc w:val="center"/>${pPr}</w:pPr><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const CASES = {
  full: [para("", LONG), para("", "Tekst po akapicie.")],
  indented: [para('<w:ind w:left="1700" w:right="1700"/>', LONG), para("", "Tekst po akapicie.")],
  title: [para("", "ZARZĄDZENIE NR 15/2026"), para("", "Rektora Uniwersytetu"), para("", "z dnia 3 listopada 2026 r."), `<w:p><w:r><w:t>Treść zwykła do lewej.</w:t></w:r></w:p>`],
};
(async () => {
  for (const [name, body] of Object.entries(CASES)) {
    const bytes = await buildDocx({ body });
    const pdf = path.join(DIR, `center-${name}.pdf`);
    try { fs.unlinkSync(pdf); } catch (_) {}
    withDocument(bytes, `save as it file name "${esc(pdf)}" file format format PDF`);
    console.log(pdf, fs.statSync(pdf).size);
  }
})();
