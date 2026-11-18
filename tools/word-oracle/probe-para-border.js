// probe-para-border.js — ile ramka akapitu (w:pBdr góra/dół) dokłada do wysokości w Wordzie.
// Blok kodu z laboratorium: Consolas 8,5 pt, cieniowanie, linie góra/dół sz=34 (4,25 pt) space=0.
// Wynik: odległość „przed” → „po” (góra liter, pt) z ramką i bez; różnica = to, co dokłada ramka.
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const { buildDocx } = require("./build-docx");
const { withDocument, DIR } = require("./word");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
const T = (t) => `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
const code = (bdr) => `<w:p><w:pPr>${bdr}<w:shd w:val="clear" w:color="auto" w:fill="F2F5F7"/><w:spacing w:after="2" w:line="259" w:lineRule="auto"/><w:ind w:left="165"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="17"/></w:rPr><w:t>systeminfo</w:t></w:r></w:p>`;
const b = (sz, space = 0, color = "F2F5F7") => `<w:pBdr><w:top w:val="single" w:sz="${sz}" w:space="${space}" w:color="${color}"/><w:bottom w:val="single" w:sz="${sz}" w:space="${space}" w:color="${color}"/></w:pBdr>`;
const CASES = { none: "", sz34: b(34), sz4: b(4), sz34sp4: b(34, 4), sz34black: b(34, 0, "000000") };
(async () => {
  for (const [k, bdr] of Object.entries(CASES)) {
    const pdf = path.join(DIR, `pbdr-${k}.pdf`);
    try { fs.unlinkSync(pdf); } catch (_) {}
    withDocument(await buildDocx({ body: [T("PRZED"), code(bdr), T("PO")] }), `save as it file name "${esc(pdf)}" file format format PDF`);
    const xml = execFileSync("pdftotext", ["-f", "1", "-l", "1", "-bbox", pdf, "-"]).toString();
    const y = (w) => +xml.match(new RegExp(`yMin="([\\d.]+)"[^>]*>${w}<`))[1];
    console.log(k.padEnd(10), "PRZED→PO", (y("PO") - y("PRZED")).toFixed(2), "pt   PRZED→kod", (y("systeminfo") - y("PRZED")).toFixed(2));
  }
})();
