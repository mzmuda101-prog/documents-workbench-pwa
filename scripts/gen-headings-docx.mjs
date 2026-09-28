#!/usr/bin/env node
/**
 * Fixture z nagłówkami (skróty sekcji, Struktura): docs/samples/headings-sample.docx
 *
 * Celowo różne sposoby oznaczania nagłówków, jak w prawdziwych plikach:
 *   - „Heading1” (angielski Word, nazwa „heading 1”),
 *   - „Nagwek2” (polski Word: styleId zlokalizowany, nazwa nadal „heading 2”),
 *   - „Rozdzial” (własny styl z w:outlineLvl=0),
 *   - „MojNaglowek” (własny styl dziedziczący po Heading1 przez basedOn).
 * Plus dużo treści, żeby dało się przewijać i sprawdzić podświetlanie bieżącej sekcji.
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import JSZip from "jszip";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "..", "docs", "samples", "headings-sample.docx");

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Nagwek2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
    <w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="Rozdzial"><w:name w:val="Rozdział"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="MojNaglowek"><w:name w:val="Mój nagłówek"/><w:basedOn w:val="Heading1"/></w:style>
</w:styles>`;

const LOREM = "Najemca zobowiązuje się używać lokalu zgodnie z jego przeznaczeniem i dbać o jego stan techniczny. Wszelkie zmiany wymagają pisemnej zgody Wynajmującego. Strony ustalają, że korespondencja będzie prowadzona drogą elektroniczną.";
const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
const p = (text, style) => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;

const body = [
  p("Umowa najmu lokalu — przykład z nagłówkami", "Heading1"),
  p("Dokument testowy dla skrótów sekcji i panelu Struktura."),
  p("§1 Strony umowy", "Rozdzial"), p(LOREM), p(LOREM),
  p("Dane Wynajmującego", "Nagwek2"), p(LOREM),
  p("Dane Najemcy", "Nagwek2"), p(LOREM),
  p("§2 Przedmiot najmu", "Heading1"), p(LOREM), p(LOREM), p(LOREM),
  p("§3 Czynsz i opłaty", "MojNaglowek"), p(LOREM), p(LOREM), p(LOREM),
  p("§4 Kaucja", "Heading1"), p(LOREM), p(LOREM),
  p("§5 Obowiązki stron", "Heading1"), p(LOREM), p(LOREM), p(LOREM),
  p("§6 Wypowiedzenie", "Heading1"), p(LOREM), p(LOREM),
  p("§7 Postanowienia końcowe", "Heading1"), p(LOREM), p(LOREM), p(LOREM), p(LOREM),
].join("\n    ");

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${body}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>`;

async function main() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", RELS);
  zip.file("word/_rels/document.xml.rels", DOC_RELS);
  zip.file("word/styles.xml", STYLES);
  zip.file("word/document.xml", DOCUMENT);
  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, buf);
  console.log(`✅  ${OUT} (${buf.length} bytes)`);
}

main();
