#!/usr/bin/env node
/**
 * Fixture linków: docs/samples/links-sample.docx
 *
 * Spis treści jak z Worda (pole TOC \h: wpisy = w:hyperlink do zakładek _Toc… + PAGEREF \h),
 * rozdziały z zakładkami i długą treścią (żeby skok przewijał), akapit z wysuniętym pierwszym
 * wierszem i łamaniami (jak „Nagłe zdarzenia 15-17” ze zgłoszenia), link do strony WWW,
 * link javascript: (ma być zablokowany), odsyłacz REF \h, pole HYPERLINK \l, link do brakującej
 * zakładki. Oczekiwania: scripts/links-playwright.js.
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import JSZip from "jszip";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "..", "docs", "samples", "links-sample.docx");
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const t = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const fld = (instr, result) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${t(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const BLUE = '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>';
const link = (attrs, text) => `<w:hyperlink ${attrs}><w:r>${BLUE}<w:t xml:space="preserve">${text}</w:t></w:r></w:hyperlink>`;
const tocEntry = (n, title, page, first, last) => `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>${first ? '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' : ""}<w:hyperlink w:anchor="_Toc${n}" w:history="1">${t(title)}<w:r><w:tab/></w:r>${fld(`PAGEREF _Toc${n} \\h`, page)}</w:hyperlink>${last ? '<w:r><w:fldChar w:fldCharType="end"/></w:r>' : ""}</w:p>`;
const heading = (n, title, extraBm = "") => `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="${n}" w:name="_Toc${n}"/>${extraBm}${t(title)}<w:bookmarkEnd w:id="${n}"/></w:p>`;
const filler = (n) => Array.from({ length: n }, (_, i) => `<w:p>${t(`Akapit wypełniający ${i + 1}: tekst, żeby rozdziały były daleko od siebie i skok musiał przewinąć dokument.`)}</w:p>`).join("");

const DOC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr>${t("Instrukcja bezpieczeństwa")}</w:p>
${tocEntry(1, "Bezpieczeństwo", "2", true, false)}
${tocEntry(2, "Nagłe zdarzenia", "3", false, false)}
${tocEntry(3, "Zagrożenia bombowe", "4", false, true)}
<w:p><w:pPr><w:ind w:left="1440" w:hanging="1080"/></w:pPr>${t("Nagłe zdarzenia 15-17")}<w:r><w:br/></w:r>${t("Niemedyczne 15")}<w:r><w:br/></w:r>${t("Medyczne 16")}<w:r><w:br/></w:r>${t("Zgłaszanie 17")}</w:p>
<w:p>${t("Strona firmy: ")}${link('r:id="rIdWeb"', "www.example.com")}${t(".")}</w:p>
<w:p>${t("Podejrzany link: ")}${link('r:id="rIdJs"', "kliknij mnie")}</w:p>
<w:p>${t("Zobacz też ")}${fld("REF _Ref9 \\h", "Zagrożenia bombowe")}${t(" na końcu.")}</w:p>
<w:p>${t("Albo ")}${fld('HYPERLINK \\l "_Toc2"', "skok polem HYPERLINK")}${t(".")}</w:p>
<w:p>${t("Zepsuty: ")}${link('w:anchor="_TocBrak"', "nie ma takiej zakładki")}</w:p>
${heading(1, "Bezpieczeństwo")}
${filler(40)}
${heading(2, "Nagłe zdarzenia")}
${filler(40)}
${heading(3, "Zagrożenia bombowe", '<w:bookmarkStart w:id="9" w:name="_Ref9"/><w:bookmarkEnd w:id="9"/>')}
${filler(10)}
<w:sectPr/>
</w:body></w:document>`;

const zip = new JSZip();
zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWeb" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://www.example.com/" TargetMode="External"/>
  <Relationship Id="rIdJs" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>
</Relationships>`);
zip.file("word/document.xml", DOC);
const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
fs.writeFileSync(OUT, buf);
console.log(`OK ${path.relative(process.cwd(), OUT)} (${buf.length} B)`);
