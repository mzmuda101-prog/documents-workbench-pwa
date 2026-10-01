#!/usr/bin/env node
/**
 * Fixture formularza Worda: docs/samples/forms-sample.docx
 *
 * Kontrolki zawartości (w:sdt): tekst z tekstem zastępczym, lista rozwijana, pole kombi,
 * data (d MMMM yyyy, pl-PL), pole wyboru (puste i zaznaczone), lista na cały akapit,
 * pole zablokowane, dwa pola powiązane z tym samym miejscem w customXml (wpisujesz raz →
 * zmienia się wszędzie), blok „bogaty tekst” (zwykła edycja) i spis treści (docPartObj —
 * to nie pole). Do tego stare pola formularza: FORMTEXT, FORMCHECKBOX, FORMDROPDOWN
 * i ochrona „tylko wypełnianie formularzy” w settings.xml.
 * Oczekiwania: scripts/forms-playwright.js.
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import JSZip from "jszip";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "..", "docs", "samples", "forms-sample.docx");
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const STORE = "{6B1E2A8C-3F44-4C0B-9D7E-2A51C0FFEE01}";
const t = (s, rPr = "") => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${s}</w:t></w:r>`;
const PH = '<w:rStyle w:val="PlaceholderText"/>';
const GOTHIC = '<w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/>';
const checkbox = (id, alias, checked) => `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:id w:val="${id}"/><w14:checkbox><w14:checked w14:val="${checked ? 1 : 0}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent>${t(checked ? "☒" : "☐", GOTHIC)}</w:sdtContent></w:sdt>`;
const items = (list) => list.map(([d, v]) => `<w:listItem w:displayText="${d}" w:value="${v}"/>`).join("");
const binding = (xpath) => `<w:dataBinding w:prefixMappings="xmlns:ns0='http://example.com/umowa'" w:xpath="${xpath}" w:storeItemID="${STORE}"/>`;
const fld = (ffData, instr, result) => `<w:r><w:fldChar w:fldCharType="begin"><w:ffData>${ffData}</w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>${result == null ? "" : `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${result ? t(result) : ""}`}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const DOC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${t("Wniosek o dostęp")}</w:p>
<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:p>${t("Spis treści")}</w:p></w:sdtContent></w:sdt>
<w:p>${t("Imię i nazwisko: ")}<w:sdt><w:sdtPr><w:alias w:val="Imię i nazwisko"/><w:tag w:val="imie"/><w:id w:val="11"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent>${t("Kliknij lub naciśnij tutaj, aby wprowadzić tekst.", PH)}</w:sdtContent></w:sdt></w:p>
<w:p>${t("Dział: ")}<w:sdt><w:sdtPr><w:alias w:val="Dział"/><w:id w:val="12"/><w:dropDownList w:lastValue="IT">${items([["Wybierz element.", ""], ["Sprzedaż", "SPR"], ["Marketing", "MKT"], ["IT", "IT"]])}</w:dropDownList></w:sdtPr><w:sdtContent>${t("IT")}</w:sdtContent></w:sdt>${t(" (zmiana wymaga zgody kierownika)")}</w:p>
<w:p>${t("Miasto: ")}<w:sdt><w:sdtPr><w:id w:val="13"/><w:showingPlcHdr/><w:comboBox>${items([["Warszawa", "WAW"], ["Kraków", "KRK"]])}</w:comboBox></w:sdtPr><w:sdtContent>${t("Wybierz element.", PH)}</w:sdtContent></w:sdt></w:p>
<w:p>${t("Data rozpoczęcia: ")}<w:sdt><w:sdtPr><w:alias w:val="Data rozpoczęcia"/><w:id w:val="14"/><w:date w:fullDate="2026-09-15T00:00:00Z"><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="pl-PL"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr><w:sdtContent>${t("15 września 2026")}</w:sdtContent></w:sdt></w:p>
<w:p>${checkbox(15, "Regulamin", false)}${t(" Akceptuję regulamin")}</w:p>
<w:p>${checkbox(16, "Newsletter", true)}${t(" Chcę dostawać newsletter")}</w:p>
<w:sdt><w:sdtPr><w:alias w:val="Priorytet"/><w:id w:val="17"/><w:dropDownList w:lastValue="N">${items([["Normalny", "N"], ["Pilny", "P"]])}</w:dropDownList></w:sdtPr><w:sdtContent><w:p>${t("Normalny")}</w:p></w:sdtContent></w:sdt>
<w:p>${t("Numer umowy: ")}<w:sdt><w:sdtPr><w:alias w:val="Numer umowy"/><w:id w:val="18"/><w:lock w:val="contentLocked"/><w:text/></w:sdtPr><w:sdtContent>${t("UM/2026/001")}</w:sdtContent></w:sdt></w:p>
<w:p>${t("Klient: ")}<w:sdt><w:sdtPr><w:alias w:val="Klient"/><w:id w:val="19"/>${binding("/ns0:umowa[1]/ns0:klient[1]")}<w:text/></w:sdtPr><w:sdtContent>${t("ACME sp. z o.o.")}</w:sdtContent></w:sdt></w:p>
<w:p>${t("Podpis klienta (")}<w:sdt><w:sdtPr><w:alias w:val="Klient"/><w:id w:val="20"/>${binding("/ns0:umowa[1]/ns0:klient[1]")}<w:text/></w:sdtPr><w:sdtContent>${t("ACME sp. z o.o.")}</w:sdtContent></w:sdt>${t(")")}</w:p>
<w:sdt><w:sdtPr><w:alias w:val="Uwagi"/><w:id w:val="21"/></w:sdtPr><w:sdtContent><w:p>${t("Uwagi można pisać tu zwyczajnie.")}</w:p></w:sdtContent></w:sdt>
<w:p>${t("Telefon: ")}${fld('<w:name w:val="Telefon"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput><w:default w:val="000 000 000"/><w:maxLength w:val="15"/></w:textInput>', "FORMTEXT", "000 000 000")}</w:p>
<w:p>${fld('<w:name w:val="Pelnoletni"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox>', "FORMCHECKBOX", null)}${t(" Mam ukończone 18 lat")}</w:p>
<w:p>${t("Rozmiar: ")}${fld('<w:name w:val="Rozmiar"/><w:enabled/><w:calcOnExit w:val="0"/><w:ddList><w:result w:val="1"/><w:listEntry w:val="S"/><w:listEntry w:val="M"/><w:listEntry w:val="L"/></w:ddList>', "FORMDROPDOWN", "M")}</w:p>
<w:p>${t("Koniec formularza.")}</w:p>
<w:sectPr/>
</w:body></w:document>`;

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
</Types>`;
const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
</Relationships>`;
const SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings ${W}><w:documentProtection w:edit="forms" w:enforcement="1"/></w:settings>`;
const ITEM = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<umowa xmlns="http://example.com/umowa"><klient>ACME sp. z o.o.</klient></umowa>`;
const ITEM_PROPS = `<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<ds:datastoreItem ds:itemID="${STORE}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="http://example.com/umowa"/></ds:schemaRefs></ds:datastoreItem>`;
const ITEM_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>
</Relationships>`;

const zip = new JSZip();
zip.file("[Content_Types].xml", CONTENT_TYPES);
zip.file("_rels/.rels", RELS);
zip.file("word/_rels/document.xml.rels", DOC_RELS);
zip.file("word/document.xml", DOC);
zip.file("word/settings.xml", SETTINGS);
zip.file("customXml/item1.xml", ITEM);
zip.file("customXml/itemProps1.xml", ITEM_PROPS);
zip.file("customXml/_rels/item1.xml.rels", ITEM_RELS);
const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
fs.writeFileSync(OUT, buf);
console.log(`OK ${path.relative(process.cwd(), OUT)} (${buf.length} B)`);
