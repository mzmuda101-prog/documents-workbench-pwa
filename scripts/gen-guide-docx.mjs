#!/usr/bin/env node
/**
 * Przewodnik po aplikacji: docs/samples/przewodnik.docx (przyciski „Przykład” / „Wypróbuj”).
 *
 * Dokument-samouczek: każda sekcja pokazuje jedną funkcję i mówi „Spróbuj: …”. Zawiera
 * celowo materiał dla paneli — nagłówki (skróty sekcji, Struktura), pola {{…}}, trigger
 * !podpis, błędy typograficzne (Korekta), powtórzone słowa (Znajdź i zamień, całe słowa,
 * wielkość liter), śledzone zmiany + komentarz (Recenzja), przypis, tabelę, listy,
 * długie zdanie (Statystyki) i metadane z autorem („Usuń dane osobowe”), klikalny spis treści
 * (linki do zakładek przy nagłówkach) i odsyłacz REF \h (Linki), pola formularza Worda —
 * tekst z tekstem zastępczym, lista, data, pole wyboru (Formularz).
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
// Rozdziały mają zakładki _GuideN — do nich prowadzi spis treści i odsyłacz (jak w Wordzie).
const chapters = [];
const h1 = (t) => {
  const n = chapters.push(t);
  return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="${100 + n}" w:name="_Guide${n}"/>${run(t)}<w:bookmarkEnd w:id="${100 + n}"/></w:p>`;
};
const tocEntry = (n) => `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Guide${n}" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">${esc(chapters[n - 1])}</w:t></w:r></w:hyperlink></w:p>`;
// odsyłacz Worda (Wstaw → Odsyłacz, „Wstaw jako hiperłącze”) = pole REF z \h
const xref = (bookmark, text) => ({ raw: `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF ${bookmark} \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` });
// kontrolki zawartości Worda (Deweloper → Formanty)
let sdtId = 500;
const sdtText = (alias) => ({ raw: `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:id w:val="${++sdtId}"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Kliknij tutaj, aby wpisać.</w:t></w:r></w:sdtContent></w:sdt>` });
const sdtList = (alias, items, current) => ({ raw: `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:id w:val="${++sdtId}"/><w:dropDownList w:lastValue="${esc(current)}">${items.map((i) => `<w:listItem w:displayText="${esc(i)}" w:value="${esc(i)}"/>`).join("")}</w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>${esc(current)}</w:t></w:r></w:sdtContent></w:sdt>` });
const sdtDate = (alias) => ({ raw: `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:id w:val="${++sdtId}"/><w:showingPlcHdr/><w:date><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="pl-PL"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Wybierz datę.</w:t></w:r></w:sdtContent></w:sdt>` });
const sdtCheck = (alias) => ({ raw: `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:id w:val="${++sdtId}"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/></w:rPr><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>` });
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
  "@@TOC@@",

  h1("1. Czytanie i edycja"),
  p("Aplikacja startuje w trybie Czytanie — nic przypadkiem się nie zmieni. Przełącz na Edycja na pasku nad dokumentem i kliknij w dowolny akapit, żeby pisać jak w Wordzie."),
  tip("przełącz na Edycja, kliknij na końcu tego zdania i dopisz kilka słów. Potem zaznacz jedno słowo i kliknij B, I albo U. Na końcu kliknij ↶ — zmiany się cofną."),
  p(["Formatowanie z Worda zostaje: ", { t: "pogrubienie", b: true }, ", ", { t: "kursywa", i: true }, ", ", { t: "podkreślenie", u: true }, ", ", { t: "kolor", color: "C0392B" }, " i ", { t: "wyróżnienie", hl: "yellow" }, "."]),
  p("Enter dzieli akapit, Shift+Enter łamie wiersz w tym samym akapicie, Tab na liście zmienia poziom punktu. Strzałki przechodzą między akapitami, Delete na końcu akapitu dołącza następny, a kliknięcie obok tekstu stawia kursor w najbliższym wierszu."),
  p("Zaznaczanie jak w Wordzie, także przez wiele akapitów: przeciągnij myszą (na telefonie — uchwytami), Shift+klik, Shift+strzałki albo Ctrl/⌘+A (cała treść). Pisanie, Delete, Wytnij i Wklej zastępują całe zaznaczenie, a B, I, U, kolor, krój i rozmiar zmieniają je w każdym akapicie."),
  p("Krój i rozmiar czcionki wybierasz na pasku (Ctrl/⌘+Shift+> / < — większa / mniejsza). Pola pokazują to, co jest w miejscu kursora; przy zaznaczeniu z różnymi rozmiarami są puste, a pasek stanu na dole pokazuje np. „Arial · 9–14 pt”. Rozmiar wybrany bez zaznaczenia dotyczy tylko tekstu, który wpiszesz w tym miejscu."),

  h1("2. Szukanie oraz Znajdź i zamień"),
  p("Najemca płaci czynsz do dziesiątego dnia miesiąca. Najemca dba o lokal, a najemca pokrywa drobne naprawy. Najemcami mogą być też osoby prawne."),
  tip("wpisz „najemca” w polu szukania nad dokumentem i naciśnij Enter (kolejne trafienia: Enter albo ↑ ↓). W panelu Znajdź i zamień zaznacz „Tylko całe słowa” — „Najemcami” przestanie być trafieniem; „Rozróżniaj wielkość liter” zostawi tylko to, co napisano małą literą — dokładnie tak, jak w polu szukania."),

  h1("3. Placeholdery — szablon do wypełnienia"),
  p("Umowę zawarto dnia {{data_podpisania}} w {{miasto}} pomiędzy {{wynajmujacy}} a {{najemca}}. Czynsz wynosi {{kwota}} zł miesięcznie."),
  tip("otwórz panel Placeholdery i kliknij „Skanuj pola” — pojawi się formularz z pięcioma polami. Wpisz wartości i „Wypełnij”. Wartości zapiszesz do pliku JSON, żeby użyć ich przy kolejnej umowie."),

  h1("4. Formularz Worda — pola do klikania"),
  p("Word ma gotowe pola formularza: listę do wyboru, datę z kalendarza, pole wyboru i pole tekstowe z szarą podpowiedzią. Tutaj działają tak samo — i po zapisie dalej działają w Wordzie."),
  p(["Imię i nazwisko: ", sdtText("Imię i nazwisko")]),
  p(["Dział: ", sdtList("Dział", ["Sprzedaż", "Marketing", "IT"], "Sprzedaż")]),
  p(["Data rozpoczęcia: ", sdtDate("Data rozpoczęcia")]),
  p([sdtCheck("Zgoda"), " Akceptuję warunki"]),
  tip("kliknij „Sprzedaż” i wybierz inny dział, kliknij datę i wybierz dzień z kalendarza, kliknij ☐ — zaznaczy się od razu. Wszystkie pola widać też w panelu Formularz (plakietka pokazuje ich liczbę)."),

  h1("5. Snippety — gotowe kawałki tekstu"),
  p("Snippet to zapisany fragment, który wstawiasz wpisując wykrzyknik i nazwę. Mogą zawierać dzisiejszą datę, pola do uzupełnienia i miejsce na kursor."),
  tip("w panelu Snippety kliknij „Dodaj przykładowe snippety”. Potem w trybie Edycja wpisz tutaj „!” — pojawi się lista; wybierz np. !dzis albo !pozdrawiam (zapyta o imię i nazwisko)."),
  p("Podpis na końcu pisma: !podpis"),
  p("Ten trigger jeszcze nie ma definicji — zapisz w panelu snippet o nazwie „podpis” i kliknij „Rozwiń w dokumencie”."),

  h1("6. Korekta typografii"),
  p("W tym akapicie są błędy  do poprawienia : podwójne spacje, spacja przed dwukropkiem, trzy kropki zamiast wielokropka... oraz \"proste cudzysłowy\" zamiast polskich. nowe zdanie od małej litery też się znajdzie, ale skrót „sp. z o.o.” zostanie w spokoju."),
  tip("otwórz panel Korekta i kliknij Skanuj. Każdą sugestię poprawisz osobno albo wszystkie naraz — formatowanie akapitu zostaje."),

  h1("7. Recenzja: zmiany i komentarze"),
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

  h1("8. Przypisy, tabele i listy"),
  p([
    "Aplikacja pokazuje przypisy tak jak Word",
    { raw: '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>' },
    ". Najedź na numer przypisu, żeby zobaczyć jego treść; w Edycji dwuklik przenosi do przypisu na dole strony — jego tekst poprawisz tam jak zwykły akapit (Enter = nowy akapit przypisu), a klik w numer przypisu wraca do tekstu. Lista przypisów jest też w panelu Recenzja.",
  ]),
  p("Nowy przypis: „＋ Wstaw → Przypis dolny” (Ctrl/⌘+Alt+F) albo „Przypis końcowy” (Ctrl/⌘+Alt+D) — numer pojawia się w miejscu kursora, a kursor przechodzi do tekstu przypisu. Numery idą po kolei w tekście, a skasowanie numeru w tekście usuwa cały przypis (jak w Wordzie)."),
  table([
    ["Funkcja", "Gdzie", "Skrót"],
    ["Szukaj", "pasek nad dokumentem", "Ctrl/⌘+F"],
    ["Cofnij / Ponów", "↶ ↷ na pasku", "Ctrl/⌘+Z / Shift+Z"],
    ["Czytanie ⇄ Edycja", "przełącznik na pasku", "Ctrl/⌘+Alt+E"],
    ["Tryb skupienia", "przycisk ⛶", "Ctrl/⌘+Alt+F (poza tekstem)"],
    ["Podgląd wydruku", "menu ⋯", "Ctrl/⌘+P"],
  ]),
  bullet("Lista punktowana — w trybie Edycja Enter dodaje kolejny punkt."),
  bullet("Tab przesuwa punkt poziom niżej, Shift+Tab wyżej."),
  num("Lista numerowana — numeracja przelicza się sama."),
  num("Drugi punkt listy numerowanej."),

  h1("9. Statystyki, eksport i metadane"),
  p("To zdanie jest celowo bardzo długie, bo panel Statystyki pokazuje najdłuższe zdania w dokumencie, a długie zdania, choć czasem potrzebne w umowach i pismach urzędowych, zwykle czyta się trudniej, więc warto je od czasu do czasu podzielić na krótsze i sprawdzić, czy wciąż mówią to samo."),
  tip("panel Statystyki pokaże liczbę słów, czas czytania i to długie zdanie (kliknij je, żeby do niego przejść). Eksport zapisze tekst jako TXT, Markdown, HTML albo PDF. W Metadanych jest autor „Jan Przykładowy” — przycisk „Usuń dane osobowe” wyczyści go przed wysłaniem pliku."),

  p("„Drukuj / zapisz jako PDF” (albo Ctrl/⌘+P) otwiera Podgląd wydruku: prawdziwe kartki z marginesami, nagłówkiem i numerem na każdej stronie — dokładnie tak wyjdą z drukarki. Gdy tekst leży bliżej niż 6,35 mm od krawędzi kartki, podgląd ostrzega: wiele drukarek tam nie drukuje."),

  h1("10. Linki i odsyłacze"),
  p(["Spis treści na początku tego pliku to linki do rozdziałów — jak w Wordzie. Ten akapit ma też odsyłacz: szczegóły zapisu są w rozdziale ", xref("_Guide12", "12. Zapis i otwieranie plików"), "."]),
  tip("kliknij odsyłacz w zdaniu wyżej — dokument przeskoczy do rozdziału o zapisie. Na dole pojawi się „↩ Wróć” (albo Alt+←): wróci dokładnie tutaj. Link do strony WWW otworzy się w nowej karcie, a aplikacja zostanie otwarta."),

  h1("11. Na telefonie i tablecie"),
  bullet("Dwa palce przybliżają i oddalają tekst — procent widać na żywo."),
  bullet("Nagłówek aplikacji chowa się przy przewijaniu; pociągnij w dół na samej górze, żeby go wysunąć."),
  bullet("„Zapisz” w Safari zapisuje kopię pliku (Pliki → Pobrane)."),

  h1("12. Zapis i otwieranie plików"),
  p("„Zapisz” zapisuje zmiany w oryginalnym pliku, jeśli otworzyłeś go przyciskiem „Otwórz”, przeciągnięciem do okna albo przez „Otwórz za pomocą” (Chrome i Edge na komputerze) — o zgodę zapyta raz. W innym wypadku zapyta, gdzie zapisać kopię. „Zapisz jako” zawsze tworzy kopię. Licznik przy „Zapisz” pokazuje, ile zmian czeka na zapis. Plik zostaje zwykłym .docx — otworzysz go w Wordzie, Pages i LibreOffice."),
  p("Zainstalowana aplikacja (Chrome / Edge → „Zainstaluj”) pojawia się w menu „Otwórz za pomocą” przy plikach .docx na Windowsie i Macu, a na Androidzie w „Udostępnij”. Na iPhonie i iPadzie pliki otwierasz z wnętrza aplikacji."),
  tip("zmień coś, zobacz licznik przy „Zapisz”, a potem użyj „Zapisz jako”, żeby mieć własną kopię tego przewodnika."),

  h1("13. Tworzenie dokumentu od zera"),
  p("„Nowy dokument” (ekran startowy, menu ⋯, panel Plik albo Ctrl/⌘+Alt+N) tworzy plik od razu w przeglądarce: pusta kartka A4, pismo z polami {{…}} albo notatka z tytułem i nagłówkami. Pierwsze „Zapisz” zapyta o nazwę i miejsce."),
  p("W trybie Edycja na pasku są: „＋ Wstaw” (podział strony, linia pozioma, dzisiejsza data, znaki specjalne), lista stylu akapitu (Normalny, Tytuł, Podtytuł, Nagłówek 1–3, Cytat) i wyrównanie. Nagłówki od razu trafiają do skrótów sekcji, Struktury i spisu treści w Wordzie. Enter na końcu nagłówka zaczyna zwykły tekst, a Ctrl/⌘+Enter przenosi dalszy tekst na nową stronę."),
  tip("kliknij na końcu tego akapitu i wybierz z listy stylu „Nagłówek 2” — pojawi się w skrótach sekcji. Potem „＋ Wstaw” → „Dzisiejsza data”, a na koniec ↶ cofnie jedno i drugie."),
  p("Przycisk listy robi listę punktowaną albo numerowaną (drugi raz — zdejmuje). Tab / Shift+Tab albo „Głębiej / Płycej” w tym samym okienku zmienia poziom punktu; Enter w pustym punkcie kończy listę, a Backspace na początku punktu zdejmuje numerację."),
  p("Ctrl/⌘+K (albo „＋ Wstaw” → „Link…”) robi link z zaznaczonego tekstu — do strony WWW albo do nagłówka w tym dokumencie. Gdy kursor stoi w linku, pod nim pojawia się mała karta: otwórz, zmień, usuń. W „＋ Wstaw” są też pola formularza (tekst, lista wyboru, data, pole wyboru) — wstawiasz je w środek zdania i dalej piszesz obok — oraz spis treści z nagłówków 1–3, który później aktualizujesz tym samym przyciskiem."),
  tip("zaznacz słowo „spis treści” w zdaniu wyżej, naciśnij Ctrl/⌘+K, przełącz na „Miejsce w dokumencie” i wybierz rozdział 1 — kliknięcie w link (albo „Otwórz” na karcie) przeniesie Cię tam."),
  p("„＋ Wstaw → Tabela” pokazuje siatkę — wybierasz liczbę kolumn i wierszy. Tab przechodzi do następnej komórki, a w ostatniej dokłada wiersz. Gdy kursor stoi w tabeli, na pasku pojawia się przycisk „Tabela”: wiersz powyżej / poniżej, kolumna z lewej / z prawej, usuwanie wiersza, kolumny albo całej tabeli."),
  p("„＋ Wstaw → Obraz…” wstawia zdjęcie z dysku (na telefonie także z aparatu), a Ctrl/⌘+V wkleja zrzut ekranu. Duże zdjęcia są zmniejszane, żeby plik nie puchł. Kliknięcie obrazu pokazuje kartę: suwak szerokości, wyrównanie, tekst alternatywny (opis dla czytników ekranu) i usuwanie — działa też z obrazami z plików Worda."),
  tip("kliknij w komórkę tabeli z rozdziału 8 — na pasku pojawi się „Tabela”; dodaj wiersz poniżej i wpisz coś, przechodząc Tabem."),
  p("Przycisk „A” na pasku to kolor czcionki (paleta jak w Wordzie, „Automatyczny”, „Więcej kolorów…”) i wyróżnienie tekstu (zakreślacz). Bez zaznaczenia kolor obowiązuje dla dalszego pisania."),
  p("Komentarz: zaznacz tekst i Ctrl/⌘+Alt+M (albo „＋ Wstaw → Komentarz”). Komentowany tekst jest podświetlony; gdy kursor w nim stoi, karta pozwala odpowiedzieć, oznaczyć jako rozwiązany albo usunąć. Wszystkie komentarze są też w panelu Recenzja."),
  p("„＋ Wstaw → Układ strony…”: marginesy jak w Wordzie (Normalne 2,5 cm, Wąskie, Umiarkowane, Szerokie albo własne w cm), orientacja pionowa/pozioma i rozmiar papieru (A4, A5, Letter…). Zmiany widać na kartkach w Widoku desktopowym i w Podglądzie wydruku."),
  p("„＋ Wstaw → Nagłówek, stopka, numer strony…” (albo kliknięcie w nagłówek/stopkę w Edycji): tekst u góry i na dole każdej strony, numer strony („1”, „Strona 1”, „Strona 1 z 5”, „– 1 –”) i inna pierwsza strona, np. tytułowa bez numeru."),
  tip("zaznacz słowo w tym akapicie, naciśnij Ctrl/⌘+Alt+M i dodaj komentarz — potem kliknij w podświetlony tekst i odpowiedz na niego."),

  h1("14. PDF → Word (.docx)"),
  p("Otwórz plik .pdf tak jak .docx (Otwórz, przeciągnięcie do okna, „Otwórz za pomocą”) — aplikacja zamieni go w edytowalny dokument Worda. Wszystko dzieje się na tym urządzeniu: plik nigdzie nie jest wysyłany, działa też bez internetu. Okienko pokazuje postęp strona po stronie; „Anuluj” przerywa w każdej chwili."),
  bullet("Zostają: akapity z wyrównaniem i wcięciami, nagłówki (trafiają do skrótów sekcji i Struktury), punktory, tabele z liniami i tłem komórek, obrazy, linki, spis treści z kropkami, numery stron w stopce, kolumny i układ kilku kart na stronie."),
  bullet("Czcionki z PDF są osadzane w pliku, gdy na komputerze może ich brakować — tekst zawija się jak w oryginale."),
  bullet("Wypełnione pola formularza PDF trafiają do tekstu z polskimi literami, nawet gdy PDF rysował je bez „Ł” czy „Ś”."),
  p("Skan (zdjęcie strony) — aplikacja zapyta, czy rozpoznać tekst (OCR). Rozpoznany druk staje się zwykłym tekstem, a skan zostaje pod spodem jako tło: kolory, linie tabel i pieczątki widać jak w oryginale. Na tekst zamieniane są tylko słowa rozpoznane z dużą pewnością; pismo odręczne i nieczytelne fragmenty zostają na obrazie bez zmian (rozpoznawanie pisma odręcznego — do poprawy w przyszłości)."),
  p("Po konwersji dokument jest nowy i niezapisany — „Zapisz” zapyta o nazwę i miejsce pliku .docx. Sprawdź wynik, zwłaszcza po OCR."),
  tip("zapisz dowolną stronę WWW jako PDF (Drukuj → Zapisz jako PDF) i otwórz ją tutaj — zobaczysz, jak wygląda po zamianie na Worda."),
];

const tocBlock = [
  `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>${run("Spis treści")}</w:p>`,
  ...chapters.map((_, i) => tocEntry(i + 1)),
  tip("kliknij dowolny rozdział — przeskoczysz do niego. „↩ Wróć” na dole (albo Alt+←) przywróci to miejsce."),
].join("\n");
const bodyXml = body.map((x) => (x === "@@TOC@@" ? tocBlock : x));

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${bodyXml.join("\n")}
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
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="80"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3B63"/><w:sz w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="20"/><w:ind w:left="240"/></w:pPr></w:style>
  <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="1F5FBF"/><w:u w:val="single"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style>
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
