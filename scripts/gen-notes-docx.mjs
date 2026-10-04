#!/usr/bin/env node
/**
 * Fixture przypisów: docs/samples/notes-sample.docx
 *
 * Jak z Worda: styl znakowy „Odwołanie przypisu dolnego” (indeks górny), „Tekst przypisu dolnego”,
 * separator / kontynuacja. Akapit z odnośnikiem w środku, akapit z tabulatorem i odnośnikiem na
 * KOŃCU, podział strony (numeracja ma iść dalej: 3, nie 1), przypis z dwoma akapitami, przypis
 * z tabelą (tylko do odczytu) i przypis końcowy (Word: i, ii…). Oczekiwania:
 * scripts/notes-playwright.js.
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import JSZip from "jszip";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "..", "docs", "samples", "notes-sample.docx");
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const t = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const fnRef = (id) => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const enRef = (id) => `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
const fnMark = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>';
const enMark = '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>';
const fnPara = (body, mark = "") => `<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>${mark}${body}</w:p>`;

const DOC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr>${t("Zapotrzebowanie")}</w:p>
<w:p>${t("Plan zakupów")}${fnRef(1)}${t(" — pozycja z planu na bieżący rok.")}</w:p>
<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="4536"/></w:tabs></w:pPr>${t("Numer pozycji")}<w:r><w:tab/></w:r>${t("z planu")}${fnRef(2)}</w:p>
<w:p>${t("Akapit bez przypisu, zwykły tekst.")}</w:p>
<w:p>${t("Koniec pierwszej części.")}<w:r><w:br w:type="page"/></w:r>${t("Początek drugiej strony w tym samym akapicie.")}</w:p>
<w:p>${t("Druga strona: źródło finansowania")}${fnRef(3)}${t(" i uwagi")}${enRef(1)}${t(".")}</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;

const SEP = `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>`;
const FOOTNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes ${W}>${SEP}
<w:footnote w:id="1">${fnPara(t(" Brak ujęcia zakupu w Planie nie oznacza odmowy."), fnMark)}</w:footnote>
<w:footnote w:id="2">${fnPara(t(" Jeżeli pozycji nie ma w Planie — wpisz „brak”."), fnMark)}${fnPara(t("Drugi akapit przypisu."))}</w:footnote>
<w:footnote w:id="3">${fnPara(t(" Kody finansowania:"), fnMark)}<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p>${t("kod 8")}</w:p></w:tc></w:tr></w:tbl>${fnPara(t("Koniec."))}</w:footnote>
</w:footnotes>`;
const ENDNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes ${W}><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
<w:endnote w:id="1"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>${enMark}${t(" Uwaga końcowa dokumentu.")}</w:p></w:endnote>
</w:endnotes>`;
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
</w:styles>`;
const SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings ${W}><w:defaultTabStop w:val="708"/><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr></w:settings>`;

const zip = new JSZip();
const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml";
zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="${CT}.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="${CT}.settings+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="${CT}.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="${CT}.endnotes+xml"/>
</Types>`);
zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${REL}/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="${REL}/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId4" Type="${REL}/endnotes" Target="endnotes.xml"/>
</Relationships>`);
zip.file("word/document.xml", DOC);
zip.file("word/styles.xml", STYLES);
zip.file("word/settings.xml", SETTINGS);
zip.file("word/footnotes.xml", FOOTNOTES);
zip.file("word/endnotes.xml", ENDNOTES);
fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
console.log(`✅  ${path.relative(process.cwd(), OUT)}`);
