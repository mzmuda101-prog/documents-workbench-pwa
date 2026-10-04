#!/usr/bin/env node
/**
 * Fixture zaznaczania przez akapity i czcionek: docs/samples/selection-sample.docx
 *
 * Nagłówek 1 + akapity Normal (Arial 11 pt ze stylu), akapit z fragmentami 9 pt i 14 pt i krojem
 * Georgia, tabela 2×2, akapit z samym podziałem strony (nowa kartka w podglądzie — tylko do
 * odczytu) i tekst na drugiej stronie. Oczekiwania: scripts/selection-playwright.js.
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import JSZip from "jszip";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "..", "docs", "samples", "selection-sample.docx");
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const t = (s, rPr = "") => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${s}</w:t></w:r>`;
const p = (body, pPr = "") => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${body}</w:p>`;
const cell = (s) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(t(s))}</w:tc>`;

const DOC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
${p(t("Rozdział pierwszy"), '<w:pStyle w:val="Heading1"/>')}
${p(t("Alfa pierwszy akapit zwykłego tekstu."))}
${p(t("Beta drugi akapit zwykłego tekstu."))}
${p(t("Gamma trzeci akapit zwykłego tekstu."))}
${p(t("Mały ") + t("dziewiątka", '<w:sz w:val="18"/>') + t(" i ") + t("czternastka", '<w:sz w:val="28"/>') + t(" oraz ") + t("Georgia", '<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>') + t("."))}
${p(t("Delta przed tabelą."))}
<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
<w:tr>${cell("A1 komórka")}${cell("B1 komórka")}</w:tr><w:tr>${cell("A2 komórka")}${cell("B2 komórka")}</w:tr></w:tbl>
${p(t("Epsilon po tabeli."))}
${p('<w:r><w:br w:type="page"/></w:r>')}
${p(t("Zeta na drugiej stronie."))}
${p(t("Eta ostatni akapit."))}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;
const FONTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts ${W}><w:font w:name="Arial"/><w:font w:name="Georgia"/><w:font w:name="Symbol"/></w:fonts>`;

const zip = new JSZip();
const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml";
zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="${CT}.styles+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="${CT}.fontTable+xml"/>
</Types>`);
zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${REL}/fontTable" Target="fontTable.xml"/>
</Relationships>`);
zip.file("word/document.xml", DOC);
zip.file("word/styles.xml", STYLES);
zip.file("word/fontTable.xml", FONTS);
fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
console.log(`✅  ${path.relative(process.cwd(), OUT)}`);
