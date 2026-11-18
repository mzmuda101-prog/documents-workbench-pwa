// probe-autospacing.js — odstęp „automatyczny” (w:beforeAutospacing / w:afterAutospacing, styl
// „Normal (Web)”) w Wordzie: w treści i w komórce tabeli (Sprawozdanie_zadanie_3: wiersze tabeli
// o ~7,5 pt wyższe niż w Wordzie). Wynik: odległości kolejnych linijek (góra liter, pt).
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const { buildDocx } = require("./build-docx");
const { withDocument, DIR } = require("./word");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
const STY = `<w:style w:type="paragraph" w:styleId="NormalWeb"><w:name w:val="Normal (Web)"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="auto"/></w:pPr></w:style>`;
const T = (t) => `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
const W = (t) => `<w:p><w:pPr><w:pStyle w:val="NormalWeb"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
const B = ["top", "left", "bottom", "right", "insideH", "insideV"].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join("");
const row = (...cells) => `<w:tr>${cells.map((c) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${c}</w:tc>`).join("")}</w:tr>`;
const tbl = (...rows) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${B}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${rows.join("")}</w:tbl>`;
const CASES = {
  body: [T("PRZED"), W("A1"), W("A2"), T("PO")],
  cellOne: [T("PRZED"), tbl(row(T("R1"), W("A1")), row(T("R2"), T("x")), row(T("R3"), T("x"))), T("PO")],
  cellPlain: [T("PRZED"), tbl(row(T("R1"), T("A1")), row(T("R2"), T("x")), row(T("R3"), T("x"))), T("PO")],
  cellTwo: [T("PRZED"), tbl(row(T("R1"), W("A1") + W("A2")), row(T("R2"), T("x")), row(T("R3"), T("x"))), T("PO")],
};
(async () => {
  for (const [k, body] of Object.entries(CASES)) {
    const pdf = path.join(DIR, `autosp-${k}.pdf`);
    try { fs.unlinkSync(pdf); } catch (_) {}
    withDocument(await buildDocx({ body, styles: STY }), `save as it file name "${esc(pdf)}" file format format PDF`);
    const xml = execFileSync("pdftotext", ["-f", "1", "-l", "1", "-bbox", pdf, "-"]).toString();
    const words = [...xml.matchAll(/yMin="([\d.]+)"[^>]*>([^<]+)</g)].map((m) => [m[2], +m[1]]).filter(([w]) => /^(PRZED|PO|R\d|A\d)$/.test(w));
    console.log(k.padEnd(10), words.map(([w, y]) => `${w}@${y.toFixed(2)}`).join("  "));
  }
})();
