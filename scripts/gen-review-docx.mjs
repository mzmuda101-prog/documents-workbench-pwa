#!/usr/bin/env node
/**
 * Fixture recenzji: docs/samples/review-sample.docx
 *
 * Śledzone zmiany dwóch autorów (wstawienie, usunięcie, przeniesienie, zmiana formatowania,
 * zmiana akapitu, wstawiony akapit, usunięty znak akapitu, wstawienie w tabeli i w nagłówku),
 * komentarze z odpowiedzią, przypis dolny i końcowy. Oczekiwany tekst po „Akceptuj wszystkie”
 * i po „Odrzuć wszystkie” jest w scripts/review-playwright.js.
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import JSZip from "jszip";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "..", "docs", "samples", "review-sample.docx");
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const A = 'w:author="Anna Nowak" w:date="2026-09-20T10:00:00Z"';
const J = 'w:author="Jan Kowalski" w:date="2026-09-21T12:30:00Z"';
const t = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const dt = (s) => `<w:r><w:delText xml:space="preserve">${s}</w:delText></w:r>`;

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>`;
const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`;

const DOC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
<w:p>${t("Umowa serwisowa — wersja do recenzji")}</w:p>
<w:p>${t("Termin płatności wynosi ")}<w:del w:id="1" ${A}>${dt("14")}</w:del><w:ins w:id="2" ${A}>${t("30")}</w:ins>${t(" dni.")}</w:p>
<w:p><w:commentRangeStart w:id="10"/>${t("Kara umowna")}<w:commentRangeEnd w:id="10"/><w:r><w:commentReference w:id="10"/></w:r>${t(" wynosi 5% wartości zlecenia.")}</w:p>
<w:p>${t("Zleceniodawca ")}<w:ins w:id="3" ${J}>${t("niezwłocznie ")}</w:ins>${t("zgłasza usterki.")}<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>
<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="4" ${J}><w:rPr/></w:rPrChange></w:rPr><w:t>Pogrubione przez recenzenta.</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="5" ${A}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>${t("Wyśrodkowany akapit.")}</w:p>
<w:p><w:pPr><w:rPr><w:ins w:id="6" ${J}/></w:rPr></w:pPr><w:ins w:id="7" ${J}>${t("Cały nowy akapit od Jana.")}</w:ins></w:p>
<w:p><w:pPr><w:rPr><w:del w:id="8" ${A}/></w:rPr></w:pPr>${t("Początek zdania")}</w:p>
<w:p>${t(" i jego koniec.")}<w:r><w:endnoteReference w:id="1"/></w:r></w:p>
<w:p><w:moveFrom w:id="9" ${A}>${t("Przeniesiony fragment. ")}</w:moveFrom>${t("Stały tekst.")}</w:p>
<w:p>${t("Tu wraca: ")}<w:moveTo w:id="11" ${A}>${t("Przeniesiony fragment.")}</w:moveTo></w:p>
<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p>${t("Komórka ")}<w:ins w:id="12" ${J}>${t("uzupełniona")}</w:ins></w:p></w:tc></w:tr></w:tbl>
<w:p><w:commentRangeStart w:id="20"/>${t("Ostatni akapit.")}<w:commentRangeEnd w:id="20"/><w:r><w:commentReference w:id="20"/></w:r><w:r><w:commentReference w:id="21"/></w:r></w:p>
<w:sectPr><w:headerReference w:type="default" r:id="rId5"/></w:sectPr>
</w:body></w:document>`;

const COMMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments ${W}>
<w:comment w:id="10" w:author="Jan Kowalski" w:initials="JK" w:date="2026-09-21T09:00:00Z"><w:p w14:paraId="0000000A">${t("Czy 5% to nie za dużo?")}</w:p></w:comment>
<w:comment w:id="20" w:author="Anna Nowak" w:initials="AN" w:date="2026-09-22T09:00:00Z"><w:p w14:paraId="00000014">${t("Dopisać datę podpisania.")}</w:p></w:comment>
<w:comment w:id="21" w:author="Jan Kowalski" w:initials="JK" w:date="2026-09-22T10:00:00Z"><w:p w14:paraId="00000015">${t("Zgoda, dopiszę.")}</w:p></w:comment>
</w:comments>`;
const COMMENTS_EX = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx ${W}>
<w15:commentEx w15:paraId="0000000A" w15:done="0"/>
<w15:commentEx w15:paraId="00000014" w15:done="0"/>
<w15:commentEx w15:paraId="00000015" w15:paraIdParent="00000014" w15:done="0"/>
</w15:commentsEx>`;
const FOOTNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes ${W}>
<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
<w:footnote w:id="1"><w:p>${t("W ciągu 3 dni roboczych.")}</w:p></w:footnote>
</w:footnotes>`;
const ENDNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes ${W}>
<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
<w:endnote w:id="1"><w:p>${t("Zob. załącznik nr 2.")}</w:p></w:endnote>
</w:endnotes>`;
const HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${W}><w:p>${t("Poufne")}<w:ins w:id="13" ${J}>${t(" — projekt")}</w:ins></w:p></w:hdr>`;

const zip = new JSZip();
zip.file("[Content_Types].xml", CONTENT_TYPES);
zip.file("_rels/.rels", RELS);
zip.file("word/_rels/document.xml.rels", DOC_RELS);
zip.file("word/document.xml", DOC);
zip.file("word/comments.xml", COMMENTS);
zip.file("word/commentsExtended.xml", COMMENTS_EX);
zip.file("word/footnotes.xml", FOOTNOTES);
zip.file("word/endnotes.xml", ENDNOTES);
zip.file("word/header1.xml", HEADER);
fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer" }));
console.log(`✅ ${path.relative(process.cwd(), OUT)}`);
