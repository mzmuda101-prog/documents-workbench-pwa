#!/usr/bin/env node
/**
 * Przewodnik po aplikacji: docs/samples/przewodnik.docx (przyciski „Przykład” / „Wypróbuj”).
 *
 * Dokument-samouczek: każda sekcja pokazuje jedną funkcję i mówi „Spróbuj: …”. Zawiera
 * celowo materiał dla paneli — nagłówki (skróty sekcji, Struktura), pola {{…}}, trigger
 * !podpis, błędy typograficzne (Korekta), powtórzone słowa (Znajdź i zamień, całe słowa,
 * wielkość liter), śledzone zmiany + komentarz (Recenzja), przypis, tabelę, listy,
 * długie zdanie (Statystyki) i metadane z autorem („Usuń dane osobowe”).
 *
 *   node scripts/gen-guide-docx.mjs
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import JSZip from "jszip";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "..", "docs", "samples", "przewodnik.docx");

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const REV = 'w:author="Anna Recenzentka" w:date="2026-09-20T10:00:00Z"';

// run: string = zwykły tekst; { t, b, i, u, color } = sformatowany; { raw } = gotowy XML
function run(x) {
  if (typeof x === "string") return `<w:r><w:t xml:space="preserve">${esc(x)}</w:t></w:r>`;
  if (x.raw) return x.raw;
  const pr = `${x.b ? "<w:b/>" : ""}${x.i ? "<w:i/>" : ""}${x.u ? '<w:u w:val="single"/>' : ""}${x.color ? `<w:color w:val="${x.color}"/>` : ""}${x.hl ? `<w:highlight w:val="${x.hl}"/>` : ""}`;
  return `<w:r>${pr ? `<w:rPr>${pr}</w:rPr>` : ""}<w:t xml:space="preserve">${esc(x.t)}</w:t></w:r>`;
}
const p = (runs, style, extraPPr = "") => {
  const list = Array.isArray(runs) ? runs : [runs];
  const pPr = style || extraPPr ? `<w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ""}${extraPPr}</w:pPr>` : "";
  return `<w:p>${pPr}${list.map(run).join("")}</w:p>`;
};
const h1 = (t) => p(t, "Heading1");
const h2 = (t) => p(t, "Heading2");
const tip = (t) => p([{ t: "Spróbuj: ", b: true, color: "1F5FBF" }, t], "Tip");
const bullet = (runs) => p(runs, "ListParagraph", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
const num = (runs) => p(runs, "ListParagraph", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>');

const cell = (t, head) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${head ? '<w:shd w:val="clear" w:color="auto" w:fill="E8EEF8"/>' : ""}</w:tcPr>${p(head ? { t, b: true } : t)}</w:tc>`;
const table = (rows) => `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${rows.map((r, i) => `<w:tr>${r.map((c) => cell(c, i === 0)).join("")}</w:tr>`).join("")}</w:tbl>`;

const body = [
  p("Przewodnik po Documents Workbench", "Title"),
  p([{ t: "Ten plik to jednocześnie instrukcja i poligon. ", b: true }, "Wszystko, co tu zmienisz, zostaje w Twojej przeglądarce — dokument nie jest nigdzie wysyłany. Możesz śmiało psuć: przycisk ↶ cofa każdą zmianę, a oryginalny przykład wczytasz ponownie z panelu Plik."]),
  tip("przewiń w dół albo użyj skrótów sekcji nad dokumentem (to nagłówki tego pliku). Panel narzędzi otwiera przycisk z kluczem obok pola szukania."),

  h1("1. Czytanie i edycja"),
  p("Aplikacja startuje w trybie Czytanie — nic przypadkiem się nie zmieni. Przełącz na Edycja na pasku nad dokumentem i kliknij w dowolny akapit, żeby pisać jak w Wordzie."),
  tip("przełącz na Edycja, kliknij na końcu tego zdania i dopisz kilka słów. Potem zaznacz jedno słowo i kliknij B, I albo U. Na końcu kliknij ↶ — zmiany się cofną."),
  p(["Formatowanie z Worda zostaje: ", { t: "pogrubienie", b: true }, ", ", { t: "kursywa", i: true }, ", ", { t: "podkreślenie", u: true }, ", ", { t: "kolor", color: "C0392B" }, " i ", { t: "wyróżnienie", hl: "yellow" }, "."]),
  p("Enter dzieli akapit, Shift+Enter łamie wiersz w tym samym akapicie, Tab na liście zmienia poziom punktu. Strzałki przechodzą między akapitami, Delete na końcu akapitu dołącza następny, a kliknięcie obok tekstu stawia kursor w najbliższym wierszu."),

  h1("2. Szukanie oraz Znajdź i zamień"),
  p("Najemca płaci czynsz do dziesiątego dnia miesiąca. Najemca dba o lokal, a najemca pokrywa drobne naprawy. Najemcami mogą być też osoby prawne."),
  tip("wpisz „najemca” w polu szukania nad dokumentem i naciśnij Enter (kolejne trafienia: Enter albo ↑ ↓). W panelu Znajdź i zamień zaznacz „Tylko całe słowa” — „Najemcami” przestanie być trafieniem; „Rozróżniaj wielkość liter” zostawi tylko to, co napisano małą literą — dokładnie tak, jak w polu szukania."),

  h1("3. Placeholdery — szablon do wypełnienia"),
  p("Umowę zawarto dnia {{data_podpisania}} w {{miasto}} pomiędzy {{wynajmujacy}} a {{najemca}}. Czynsz wynosi {{kwota}} zł miesięcznie."),
  tip("otwórz panel Placeholdery i kliknij „Skanuj pola” — pojawi się formularz z pięcioma polami. Wpisz wartości i „Wypełnij”. Wartości zapiszesz do pliku JSON, żeby użyć ich przy kolejnej umowie."),

  h1("4. Snippety — gotowe kawałki tekstu"),
  p("Snippet to zapisany fragment, który wstawiasz wpisując wykrzyknik i nazwę. Mogą zawierać dzisiejszą datę, pola do uzupełnienia i miejsce na kursor."),
  tip("w panelu Snippety kliknij „Dodaj przykładowe snippety”. Potem w trybie Edycja wpisz tutaj „!” — pojawi się lista; wybierz np. !dzis albo !pozdrawiam (zapyta o imię i nazwisko)."),
  p("Podpis na końcu pisma: !podpis"),
  p("Ten trigger jeszcze nie ma definicji — zapisz w panelu snippet o nazwie „podpis” i kliknij „Rozwiń w dokumencie”."),

  h1("5. Korekta typografii"),
  p("W tym akapicie są błędy  do poprawienia : podwójne spacje, spacja przed dwukropkiem, trzy kropki zamiast wielokropka... oraz \"proste cudzysłowy\" zamiast polskich. nowe zdanie od małej litery też się znajdzie, ale skrót „sp. z o.o.” zostanie w spokoju."),
  tip("otwórz panel Korekta i kliknij Skanuj. Każdą sugestię poprawisz osobno albo wszystkie naraz — formatowanie akapitu zostaje."),

  h1("6. Recenzja: zmiany i komentarze"),
  p([
    "Termin płatności wynosi ",
    { raw: `<w:del w:id="1" ${REV}><w:r><w:delText>14</w:delText></w:r></w:del>` },
    { raw: `<w:ins w:id="2" ${REV}><w:r><w:t>30</w:t></w:r></w:ins>` },
    " dni od doręczenia faktury.",
  ]),
  p([
    { raw: '<w:commentRangeStart w:id="10"/>' },
    "Kara umowna wynosi 5% wartości zlecenia.",
    { raw: '<w:commentRangeEnd w:id="10"/><w:r><w:commentReference w:id="10"/></w:r>' },
  ]),
  tip("otwórz panel Recenzja — zobaczysz zmianę recenzentki (14 → 30) i komentarz do kary umownej. Kliknij ✓ albo ✗ przy zmianie, albo „Akceptuj wszystkie”. Akapit ze śledzoną zmianą da się edytować dopiero po jej rozstrzygnięciu."),

  h1("7. Przypisy, tabele i listy"),
  p([
    "Aplikacja pokazuje przypisy tak jak Word",
    { raw: '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>' },
    " — ich lista jest też w panelu Recenzja.",
  ]),
  table([
    ["Funkcja", "Gdzie", "Skrót"],
    ["Szukaj", "pasek nad dokumentem", "Ctrl/⌘+F"],
    ["Cofnij / Ponów", "↶ ↷ na pasku", "Ctrl/⌘+Z / Shift+Z"],
    ["Czytanie ⇄ Edycja", "przełącznik na pasku", "Ctrl/⌘+Alt+E"],
    ["Tryb skupienia", "przycisk ⛶", "Ctrl/⌘+Alt+F"],
  ]),
  bullet("Lista punktowana — w trybie Edycja Enter dodaje kolejny punkt."),
  bullet("Tab przesuwa punkt poziom niżej, Shift+Tab wyżej."),
  num("Lista numerowana — numeracja przelicza się sama."),
  num("Drugi punkt listy numerowanej."),

  h1("8. Statystyki, eksport i metadane"),
  p("To zdanie jest celowo bardzo długie, bo panel Statystyki pokazuje najdłuższe zdania w dokumencie, a długie zdania, choć czasem potrzebne w umowach i pismach urzędowych, zwykle czyta się trudniej, więc warto je od czasu do czasu podzielić na krótsze i sprawdzić, czy wciąż mówią to samo."),
  tip("panel Statystyki pokaże liczbę słów, czas czytania i to długie zdanie (kliknij je, żeby do niego przejść). Eksport zapisze tekst jako TXT, Markdown, HTML albo PDF. W Metadanych jest autor „Jan Przykładowy” — przycisk „Usuń dane osobowe” wyczyści go przed wysłaniem pliku."),

  h1("9. Na telefonie i tablecie"),
  bullet("Dwa palce przybliżają i oddalają tekst — procent widać na żywo."),
  bullet("Nagłówek aplikacji chowa się przy przewijaniu; pociągnij w dół na samej górze, żeby go wysunąć."),
  bullet("„Zapisz” w Safari zapisuje kopię pliku (Pliki → Pobrane)."),

  h1("10. Zapis"),
  p("„Zapisz” nadpisuje otwarty plik (w Chrome i Edge na komputerze), a „Zapisz jako” tworzy kopię. Licznik przy „Zapisz” pokazuje, ile zmian czeka na zapis. Plik zostaje zwykłym .docx — otworzysz go w Wordzie, Pages i LibreOffice."),
  tip("zmień coś, zobacz licznik przy „Zapisz”, a potem użyj „Zapisz jako”, żeby mieć własną kopię tego przewodnika."),
];

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${body.join("\n")}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS}>
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:lang w:val="pl-PL"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="200"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3B63"/><w:sz w:val="44"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="320" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F5FBF"/><w:sz w:val="30"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="Tip"><w:name w:val="Wskazówka"/><w:basedOn w:val="Normal"/><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="1F5FBF"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="EEF3FB"/><w:ind w:left="240"/><w:spacing w:before="60" w:after="180"/></w:pPr><w:rPr><w:sz w:val="20"/><w:color w:val="33415C"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/><w:ind w:left="720"/></w:pPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders>
    <w:top w:val="single" w:sz="4" w:color="B8C4D6"/><w:left w:val="single" w:sz="4" w:color="B8C4D6"/><w:bottom w:val="single" w:sz="4" w:color="B8C4D6"/><w:right w:val="single" w:sz="4" w:color="B8C4D6"/><w:insideH w:val="single" w:sz="4" w:color="B8C4D6"/><w:insideV w:val="single" w:sz="4" w:color="B8C4D6"/>
  </w:tblBorders><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
</w:styles>`;

const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${NS}>
  <w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="–"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

const COMMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments ${NS}>
  <w:comment w:id="10" w:author="Anna Recenzentka" w:initials="AR" w:date="2026-09-20T10:05:00Z"><w:p w14:paraId="0000000A"><w:r><w:t>Czy 5% nie jest za wysokie? Proponuję 3%.</w:t></w:r></w:p></w:comment>
</w:comments>`;

const FOOTNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes ${NS}>
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> To jest przykładowy przypis dolny.</w:t></w:r></w:p></w:footnote>
</w:footnotes>`;

const CORE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Przewodnik po Documents Workbench</dc:title>
  <dc:creator>Jan Przykładowy</dc:creator>
  <cp:lastModifiedBy>Anna Recenzentka</cp:lastModifiedBy>
  <cp:keywords>przewodnik, przykład</cp:keywords>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-09-20T09:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-20T10:05:00Z</dcterms:modified>
</cp:coreProperties>`;

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;
const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;
const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>`;

const zip = new JSZip();
zip.file("[Content_Types].xml", CONTENT_TYPES);
zip.file("_rels/.rels", RELS);
zip.file("word/_rels/document.xml.rels", DOC_RELS);
zip.file("word/document.xml", DOCUMENT);
zip.file("word/styles.xml", STYLES);
zip.file("word/numbering.xml", NUMBERING);
zip.file("word/comments.xml", COMMENTS);
zip.file("word/footnotes.xml", FOOTNOTES);
zip.file("docProps/core.xml", CORE);
fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
console.log(`✅ ${path.relative(process.cwd(), OUT)}`);
