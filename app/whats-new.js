// @ts-check
// „Co nowego”: lista zmian w wydaniach. Po aktualizacji okno pokazuje się samo (raz na wersję),
// a całą listę da się zawsze otworzyć z menu ⋯ → Pomoc → Co nowego — treść nie znika razem z dymkiem.
// Przy wydaniu (npm run release) dopisz NA GÓRZE wpis z nową wersją; skrypt przypomina, gdy go brak.

/* global APP_BUILD_VERSION, appDialog, t, currentLang */

const WHATS_NEW = [
  {
    v: "v1.01.48",
    pl: [
      "Konwersja PDF na Worda: szary pas na marginesie kartki (np. miejsce na komentarze) nie skleja już ramek i tekstu w jedną wielką tabelę na całą stronę, a jasne tło w ramce z kodem nie spycha tekstu w dół o całą swoją wysokość. Szeroki odstęp po wzorze (np. „KMgF3     fluorek…”) zostaje odstępem, zamiast zwijać się do jednej spacji.",
    ],
    en: [
      "PDF to Word conversion: a grey strip in the page margin (such as a comment area) no longer merges frames and text into one page-sized table, and a light background behind a code box no longer pushes the text down by its full height. A wide gap after a formula (like \"KMgF3     fluoride…\") stays a gap instead of collapsing to a single space.",
    ],
  },
  {
    v: "v1.01.47",
    pl: [
      "Konwersja PDF na Worda (tabele bez linii i harmonogramy): aplikacja rozpoznaje teraz jak Word także tabele całkowicie bez ramek i cieniowania (np. cenniki 3-kolumnowe, zestawienia klucz–wartość i małe tabelki), nie psując 2-kolumnowych artykułów ani paragonów.",
      "Harmonogramy z szarymi pasami (jak z Excela): naprawiono rozstrzelenie wierszy przy wyśrodkowanej w pionie trzeciej kolumnie, zachowano pełny pas nagłówka (nie zamienia się w zakreślenie tekstu), rozdzielono krótkie linijki w komórkach (np. nazwisko i adres) oraz zatrzymano wystawanie podświetlenia aktywnego akapitu poza krawędź tabeli.",
    ],
    en: [
      "PDF to Word conversion (borderless tables & schedules): the app now detects completely borderless, unshaded tables just like Word (3-column grids, key–value summaries, and small tables), while keeping two-column articles and receipts untouched.",
      "Shaded schedule bands (such as Excel exports): fixed oversized vertical gaps when a right-hand column is vertically centered, kept the header bar as a real table row instead of a text highlight, kept short stacked lines (like name and street) as separate paragraphs, and prevented the active-paragraph focus bar from spilling outside shaded cells.",
    ],
  },
  {
    v: "v1.01.46",
    pl: [
      "Konwersja PDF na Worda: aplikacja rozpoznaje teraz jak Word także tabele bez pionowych linii — nowoczesne tabele z samymi poziomymi liniami, styl akademicki (linia pod nagłówkiem i na dole), tabele z cieniowanym nagłówkiem, tabele zebrowe (naprzemienne pasy tła) oraz tabele z samą ramką zewnętrzną. Linie stają się prawdziwymi ramkami komórek zamiast obrazka w tle.",
    ],
    en: [
      "PDF to Word conversion: the app now detects tables without vertical lines just like Word — modern horizontal-rule tables, academic booktabs tables, shaded-header tables, zebra-striped tables, and outer-box-only grids. Rules become real cell borders instead of a background image.",
    ],
  },
  {
    v: "v1.01.45",
    pl: [
      "Tabele: nowy wiersz lub kolumna dziedziczą krój, rozmiar i pogrubienie tekstu z sąsiedniego wiersza (jak w Wordzie), a wklejenie tabeli HTML w prostym dokumencie bez własnych stylów zachowuje domyślny Aptos 12 zamiast przełączać na Calibri 11.",
      "iPhone i iPad: wyższy uchwyt zwiniętego nagłówka na telefonie (łatwiej w niego trafić bez przypadkowego kliknięcia paska sekcji pod spodem), podświetlenia z dokumentu nie nachodzą na tytuły okienek (np. „Zapisz jako”), a przeciągnięcie po tekście przy otwartym dymku komentarza od razu zaznacza treść.",
    ],
    en: [
      "Tables: inserting a new row or column now inherits the font, size, and bold styling from the adjacent row (like Word), and pasting an HTML table into a minimal document without styles keeps default Aptos 12 instead of switching to Calibri 11.",
      "iPhone and iPad: the collapsed header grip on phones is taller so tapping it doesn't hit the section bar underneath, document highlights no longer bleed onto dialog titles (like Save As), and dragging across text while a comment bubble is open selects immediately.",
    ],
  },
  {
    v: "v1.01.44",
    pl: [
      "Telefon (iPhone): otwarty krótki 1-stronicowy dokument nie ciągnie już całej aplikacji w górę i w dół, nie ucina góry kartki po zwinięciu nagłówka ani nie rysuje niebieskiej ramki wokół ekranu po stuknięciu tła. W plikach bez nagłówków zniknął też pusty pasek u góry ekranu.",
    ],
    en: [
      "Phone (iPhone): short one-page documents no longer drag the whole app up and down, clip the top of the page when collapsing the header, or show a blue focus outline around the screen after tapping the background. The empty top bar strip is also gone for documents without headings.",
    ],
  },
  {
    v: "v1.01.43",
    pl: [
      "Tabele: kliknięcie w puste miejsce wysokiej lub pustej komórki (np. w jej dolnej połowie) stawia kursor w tej komórce zamiast przeskakiwać do sąsiedniej kolumny albo kolejnego wiersza.",
    ],
    en: [
      "Tables: clicking the empty space inside a tall or empty cell (for example its lower half) now places the cursor in that cell instead of jumping to a neighbouring column or the next row.",
    ],
  },
  {
    v: "v1.01.42",
    pl: [
      "Edycja jak wydruk: przypisy dolne stoją na dole swojej strony (nad stopką), tak samo jak w podglądzie wydruku i w Wordzie, i można je od razu edytować na miejscu. Liczba stron w Edycji zgadza się z wydrukiem.",
      "iPhone i Safari: skasowanie całego tekstu punktu listy klawiszem Backspace zostawia pusty punkt z kursorem — można od razu pisać dalej (dawniej znikał punktor i kursor).",
      "Komentarze i zablokowane akapity: kliknięcie Komentarza w przypisie nie powoduje już błędu, a zaznaczenie fragmentu w akapicie tylko do odczytu nie przeskakuje po cichu do poprzedniego słowa — pojawia się wyjaśnienie blokady.",
      "Drobne poprawki: numer „Str. N/M” aktualizuje się po przewinięciu na kolejną stronę, otwarcie innego pliku czyści starą frazę i podświetlenia szukania, pole Szukaj na tablecie jest szersze, a podpowiedź po wstawieniu spisu treści wskazuje „Wstaw → Spis treści → Aktualizuj”.",
    ],
    en: [
      "Editing looks like print: footnotes now sit at the bottom of their own page (above the footer), just like in Print Preview and Word, and stay directly editable in place. Page count in Editing matches print.",
      "iPhone and Safari: clearing all text of a list item with Backspace keeps the empty item and cursor so you can keep typing right away (previously the bullet and cursor disappeared).",
      "Comments and locked paragraphs: clicking Comment while inside a footnote no longer throws an error, and selecting text in a read-only paragraph no longer jumps back to an earlier word — it explains why the paragraph is locked.",
      "Polish: the “Page N/M” status updates reliably as you scroll to the next page, opening a new file clears the previous search query and highlights, the tablet search box is wider, and the table-of-contents hint points to “Insert → Table of contents → Update”.",
    ],
  },
  {
    v: "v1.01.41",
    pl: [
      "Telefon: tabele z PDF-a są teraz czytelne — cienkie linie komórek nie znikają, wcięcia w wąskich komórkach nie łamią słów litera po literze, a wąskie kolumny dostają szerokość na najdłuższe słowo kosztem szerszych obok.",
    ],
    en: [
      "Phone: tables from PDFs are now readable — thin cell borders no longer vanish, paragraph indents in narrow cells no longer break words letter by letter, and narrow columns get enough room for their longest word from wider neighbours.",
    ],
  },
  {
    v: "v1.01.40",
    pl: [
      "Edycja wygląda jak wydruk: nagłówek i stopka są na każdej stronie, a „Strona N z M” pokazuje właściwe numery (dawniej nagłówek był tylko na pierwszej stronie, a stopka tylko na ostatniej, z „Strona 1 z 1”). Wysoki nagłówek spycha tekst na każdej stronie, jak w Wordzie.",
      "Spis treści: przy wpisach są prawdziwe numery stron (dawniej same „1”), a wpisy trzeciego poziomu mają numer w tej samej linii.",
      "Telefon: pod zdjęciem zmniejszonym do szerokości ekranu nie zostaje już puste miejsce.",
    ],
    en: [
      "Editing looks like print: the header and footer appear on every page and “Page N of M” shows the right numbers (the header used to appear only on page one and the footer only on the last page, with “Page 1 of 1”). A tall header pushes the text down on every page, like in Word.",
      "Table of contents: entries show real page numbers (it used to be all “1”), and level-three entries keep the number on the same line.",
      "Phone: no more empty space under a photo scaled down to the screen width.",
    ],
  },
  {
    v: "v1.01.39",
    pl: [
      "Naprawione: po skasowaniu całego tekstu akapitu klawisz Delete w pustej linijce mógł po zapisie poprzestawiać tekst dalszej części dokumentu (np. treść tabeli w złych komórkach). Teraz plik zapisuje się dokładnie tak, jak widać.",
      "Telefon: gdy kursor stoi w tabeli, arkusz kart od razu pokazuje kartę Tabela (wcześniej była schowana na końcu rzędu).",
    ],
    en: [
      "Fixed: after clearing all text of a paragraph, pressing Delete on the empty line could shuffle text further down the document once saved (e.g. table content in the wrong cells). The file now saves exactly as shown.",
      "Phone: with the cursor in a table, the tab sheet now shows the Table tab straight away (it used to be hidden at the end of the row).",
    ],
  },
  {
    v: "v1.01.38",
    pl: [
      "PDF → Word: punkt listy w tabeli zawinięty na dwa wiersze jest jednym akapitem — po dopisaniu słów tekst przelewa się jak w Wordzie (dawniej każdy wiersz był osobnym akapitem).",
      "PDF → Word: tekst w komórce z ramką nie wchodzi już na linię i sąsiednią kolumnę.",
    ],
    en: [
      "PDF → Word: a list item in a table that wraps onto two lines is one paragraph — added words flow on like in Word (each line used to be a separate paragraph).",
      "PDF → Word: text in a bordered cell no longer runs over the line into the next column.",
    ],
  },
  {
    v: "v1.01.37",
    pl: [
      "Listy numerowane: tekst punktu zaczyna się w tym samym miejscu w Czytaniu, Edycji i na wydruku — jak w Wordzie (dawniej każdy tryb miał inny odstęp po numerze).",
      "iPhone i iPad: Enter na końcu punktu listy — wpisany tekst trafia do nowego punktu, a nie do akapitu pod listą.",
      "Wklejanie: po zamianie wklejonego tekstu na listę lub pogrubienie przycisk „Zachowaj tylko tekst” wkleja go jak Word — dokładnie tak, jak był.",
      "PDF → Word: harmonogramy i tabele z Excela — puste szare wiersze są wierszami tabeli (dawniej obrazek), odstępy w tekście stoją jak w PDF, a na telefonie szare pasy mają całą szerokość.",
    ],
    en: [
      "Numbered lists: item text starts at the same place in Reading, Editing and print — like in Word (each mode used to leave a different gap after the number).",
      "iPhone and iPad: pressing Enter at the end of a list item — what you type goes into the new item, not the paragraph below the list.",
      "Paste: after pasted text is turned into a list or bold, “Keep text only” pastes it like Word — exactly as it was.",
      "PDF → Word: schedules and tables from Excel — empty grey rows are table rows (used to be a picture), gaps in text sit where they are in the PDF, and grey bands span the full width on phones.",
    ],
  },
  {
    v: "v1.01.36",
    pl: ["PDF → Word: skan zapisany przez skaner na ogromnej stronie jest zmniejszany do A4 w całości — dawniej każda strona szła na 3 poziome kartki z tekstem wielkości 49 pt."],
    en: ["PDF → Word: a scan saved on a huge page is scaled down to A4 as a whole — each page used to become 3 landscape sheets with 49 pt text."],
  },
  {
    v: "v1.01.35",
    pl: ["PDF → Word: wykłady i notatki z LaTeX-a — polskie litery są poprawne („Możemy”, a nie „Mo»emy”), słowa mają spacje, a wiersze stoją w dobrej kolejności."],
    en: ["PDF → Word: LaTeX lectures and notes — Polish letters are correct, words keep their spaces and lines stay in the right order."],
  },
  {
    v: "v1.01.34",
    pl: ["Telefon: w arkuszu „Główne” jest teraz „Cofnij” obok „Ponów” — nie trzeba go szukać w górnym pasku."],
    en: ["Phone: the “Home” sheet now has “Undo” next to “Redo” — no need to look for it in the top bar."],
  },
  {
    v: "v1.01.33",
    pl: [
      "PDF → Word: zakreślenia z PDF-a są tłem tekstu, jak w Wordzie — tekst nie przesuwa się, nie zmienia kolejności i nie ucina się na telefonie.",
      "Linie do wpisywania („recover ______ energy”) widać na całą długość, ze spacjami po obu stronach — także w widoku mobilnym.",
      "Wydruk zachowuje zakreślenia, tło komórek i linie tabulatorów.",
    ],
    en: [
      "PDF → Word: highlights from a PDF become text shading, like in Word — text no longer shifts, changes order or gets cut off on phones.",
      "Fill-in lines (“recover ______ energy”) show at full length with spaces on both sides — in mobile view too.",
      "Printing keeps highlights, cell shading and tab lines.",
    ],
  },
  {
    v: "v1.01.32",
    pl: [
      "Nowy ekran startowy: krótszy tytuł, równe przyciski, a na telefonie ostatnie pliki od razu widoczne.",
      "Ostatnio otwierane: pliki o tej samej nazwie jako jedna pozycja — starsze wersje rozwiniesz pod „+3 starsze wersje”. Przy każdym pliku widać, czy to Word, czy PDF.",
      "Bez otwartego pliku nie ma narzędzi do dokumentu (lewy pasek, tryb, powiększenie), a szukanie jest wyszarzone.",
    ],
    en: [
      "New start screen: a shorter title, even buttons, and on phones your recent files are visible right away.",
      "Recently opened: files with the same name are one entry — expand older versions under “+3 older versions”. Each file shows whether it is Word or PDF.",
      "With no file open, document tools (left bar, mode, zoom) are hidden and search is greyed out.",
    ],
  },
  {
    v: "v1.01.31",
    pl: ["Lepszy kontrast: szare podpisy są ciemniejsze, a niebieskie przyciski w ciemnym motywie czytelniejsze."],
    en: ["Better contrast: grey captions are darker and blue buttons in dark mode are easier to read."],
  },
  {
    v: "v1.01.30",
    pl: [
      "Nowa wersja aplikacji włącza się dopiero po kliknięciu „Aktualizuj” — nie wchodzi już „po części” sama przy otwarciu.",
      "Okno „Co nowego” po każdej aktualizacji. Otworzysz je też w każdej chwili: menu ⋯ → Co nowego.",
    ],
    en: [
      "A new app version starts only after you tap “Update” — it no longer half-loads by itself on launch.",
      "A “What's new” window after every update. Open it any time from the ⋯ menu → What's new.",
    ],
  },
  {
    v: "v1.01.29",
    pl: ["Szukanie: po wpisaniu nowej frazy strzałki ↑ ↓ szukają nowej — bez klikania „Odznacz”."],
    en: ["Search: after typing a new phrase, the ↑ ↓ arrows search for it — no need to tap “Clear” first."],
  },
  {
    v: "v1.01.28",
    pl: ["Tryb skupienia: wyjście na tablecie i przycisk na dolnym pasku."],
    en: ["Focus mode: exit on tablets and a button on the bottom bar."],
  },
  {
    v: "v1.01.27",
    pl: ["Esc w polu Szukaj wraca do dokumentu."],
    en: ["Esc in the Search field returns to the document."],
  },
  {
    v: "v1.01.26",
    pl: ["Wstążka przyklejona do górnej krawędzi okna."],
    en: ["The ribbon sticks to the top edge of the window."],
  },
];

const WHATS_NEW_SEEN_KEY = "documents-workbench-seen-version";

/** @param {string} [since] tylko wydania nowsze niż ta wersja; bez — cała lista */
function showWhatsNew(since) {
  // vMAJOR.MM.PP z zerami — porównanie napisów = porównanie wersji
  const items = since ? WHATS_NEW.filter((e) => e.v > since) : WHATS_NEW;
  if (!items.length) return;
  const body = document.createElement("div");
  body.className = "whats-new";
  for (const e of items) {
    const h = document.createElement("h3");
    h.textContent = e.v;
    const ul = document.createElement("ul");
    for (const line of currentLang === "en" ? e.en : e.pl) {
      const li = document.createElement("li");
      li.textContent = line;
      ul.append(li);
    }
    body.append(h, ul);
  }
  const title = since && items.length === 1 ? t("whatsNewTitle", { v: items[0].v }) : t("whatsNewAllTitle");
  appDialog({ title, body, actions: [{ label: t("whatsNewOk"), value: true, kind: "primary" }] });
}

(function initWhatsNew() {
  document.getElementById("whatsNewMenuItem")?.addEventListener("click", () => showWhatsNew());
  let seen = null;
  try {
    seen = localStorage.getItem(WHATS_NEW_SEEN_KEY);
    if (seen === APP_BUILD_VERSION) return;
    localStorage.setItem(WHATS_NEW_SEEN_KEY, APP_BUILD_VERSION);
  } catch (_) {
    return; // bez pamięci przeglądarki okno wyskakiwałoby przy każdym starcie
  }
  // Pierwsze wejście w ogóle (brak workera) — nic. Pierwsza aktualizacja po dodaniu „Co nowego”
  // (worker jest, zapisanej wersji jeszcze nie) — tylko ta wersja.
  const returning = !!navigator.serviceWorker?.controller;
  if (!seen && !returning) return;
  const since = seen || WHATS_NEW[1]?.v || "";
  const show = () => setTimeout(() => showWhatsNew(since), 600);
  if (document.readyState === "complete") show();
  else window.addEventListener("load", show, { once: true });
})();
