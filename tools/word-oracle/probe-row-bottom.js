// probe-row-bottom.js — czy wiersz tabeli u dołu strony mieści się bez odstępu „po” ostatniego
// akapitu w komórce (Kolokwium 18.11: Word zostawia wiersz, który z odstępem wystaje 2 pt za margines).
// Akapit odsuwający tabelę o B twipów, potem 40 wierszy jednolinijkowych (Calibri 11, po 8 pt).
// Wynik: ostatni wiersz strony 1 i jego góra liter (pt). Zapisuje output/word-oracle/rowb-*.pdf.
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const { buildDocx } = require("./build-docx");
const { withDocument, DIR } = require("./word");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
const AFTER = +(process.env.AFTER ?? 160);
const tc = (t) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="${AFTER}"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
const B = "<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${B}</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${Array.from({ length: +(process.env.ROWS || 40) }, (_, i) => `<w:tr>${tc(`R${i + 1}`)}${tc("x")}</w:tr>`).join("")}</w:tbl>`;
const sectPr = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
(async () => {
  const steps = (process.env.STEPS || "0,40,80,120,160,200,240,280,320,360,400,440").split(",").map(Number);
  for (const b of steps) {
    const body = [`<w:p><w:pPr><w:spacing w:before="${b}" w:after="0"/></w:pPr><w:r><w:t>X</w:t></w:r></w:p>${table}<w:p/>`];
    const bytes = await buildDocx({ body, sectPr });
    const pdf = path.join(DIR, `rowb-${b}.pdf`);
    try { fs.unlinkSync(pdf); } catch (_) {}
    withDocument(bytes, `save as it file name "${esc(pdf)}" file format format PDF`);
    const xml = execFileSync("pdftotext", ["-f", "1", "-l", "1", "-bbox", pdf, "-"]).toString();
    const rows = [...xml.matchAll(/yMin="([\d.]+)"[^>]*>(R\d+)</g)].map((m) => [m[2], +m[1]]);
    const last = rows[rows.length - 1];
    console.log(`B=${b} tw  ostatni na str. 1: ${last[0]} góra liter ${last[1].toFixed(2)} pt`);
  }
})();
