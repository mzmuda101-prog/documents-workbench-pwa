// probe-table-indent.js — gdzie Word stawia lewą krawędź tabeli (w:tblInd 0) w zależności od trybu
// zgodności: 15 (Word 2013+), 14, 12 i plik bez settings.xml (Kolokwium 18.11: linia 7,5 px przed
// marginesem). VARIANT=plain | ind0 | styled (Table Grid na Normal Table jak w pliku). Wynik: x lewej linii tabeli i x tekstu komórki (pt; margines 70,85 pt).
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const JSZip = require("jszip");
const { buildDocx } = require("./build-docx");
const { withDocument, DIR } = require("./word");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
const B = ["top", "left", "bottom", "right", "insideH", "insideV"].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join("");
const tbl = (ind = "") => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${ind}<w:tblBorders>${B}</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>KOMORKA</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
const STYLES = `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders>${B}</w:tblBorders></w:tblPr></w:style>`;
const styled = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>KOMORKA</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
const VARIANT = process.env.VARIANT || "plain";
const body = [`<w:p><w:r><w:t>TEKST</w:t></w:r></w:p>`, VARIANT === "styled" ? styled : VARIANT === "ind0" ? tbl('<w:tblInd w:w="0" w:type="dxa"/>') : tbl(), `<w:p/>`];
(async () => {
  for (const mode of ["15", "14", "12", "brak"]) {
    const zip = await JSZip.loadAsync(await buildDocx({ body, styles: VARIANT === "styled" ? STYLES : "" }));
    if (mode === "brak") { zip.remove("word/settings.xml"); zip.file("word/_rels/document.xml.rels", (await zip.file("word/_rels/document.xml.rels").async("string")).replace(/<Relationship Id="rId2"[^>]*\/>/, "")); zip.file("[Content_Types].xml", (await zip.file("[Content_Types].xml").async("string")).replace(/<Override PartName="\/word\/settings.xml"[^>]*\/>/, "")); }
    else zip.file("word/settings.xml", (await zip.file("word/settings.xml").async("string")).replace(/w:val="15"/, `w:val="${mode}"`));
    const bytes = await zip.generateAsync({ type: "nodebuffer" });
    const pdf = path.join(DIR, `tind-${mode}.pdf`);
    try { fs.unlinkSync(pdf); } catch (_) {}
    withDocument(bytes, `save as it file name "${esc(pdf)}" file format format PDF`);
    const xml = execFileSync("pdftotext", ["-f", "1", "-l", "1", "-bbox", pdf, "-"]).toString();
    const x = (w) => +xml.match(new RegExp(`xMin="([\\d.]+)"[^>]*>${w}<`))[1];
    // lewa linia: najjaśniej z obrazu — pdftoppm 288 dpi, pierwsza ciemna kolumna w wierszu tabeli
    const png = path.join(DIR, `tind-${mode}`);
    execFileSync("pdftoppm", ["-r", "288", "-f", "1", "-l", "1", "-gray", "-png", pdf, png]);
    console.log(`${VARIANT.padEnd(7)} tryb ${mode.padEnd(4)} tekst akapitu x=${x("TEKST").toFixed(2)}  tekst komórki x=${x("KOMORKA").toFixed(2)} pt`);
  }
})();
