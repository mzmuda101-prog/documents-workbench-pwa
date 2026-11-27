# Osoby dla testerów AI

Katalog osób, w które wcielają się testerzy AI (`/testerzy`). Skrypt `scripts/tester/pick.mjs`
losuje z niego skład według wag: zawsze co najmniej 2 telefony, 1 iPad i 1 komputer.

Jak dopisać osobę:

- Nowa sekcja zaczyna się od nagłówka drugiego poziomu z identyfikatorem (kebab-case, same małe
  litery ascii, cyfry i myślniki, np. `organizator-iphone`).
- Pod nim dokładnie pięć linii, w tej kolejności, każda zaczyna się od „- ”:
  - `urządzenie:` jedno z `iphone`, `iphone-landscape`, `ipad`, `desktop`;
  - `waga:` jedno z `częste`, `średnie`, `rzadkie` (jak często ktoś tak naprawdę używa apki);
  - `obszary:` tagi po przecinku, np. pdf, tabele, obrazy, pisanie, formatowanie, listy,
    wstawianie, recenzja, szukaj, pisownia, zapis, wydruk, wklejanie, przerwania, otwieranie,
    pierwsze-uruchomienie, offline, uklad (po nich działa `--area=`);
  - `kto:` jedna linia: kim jest, jak sprawnie używa komputera, jakie ma nawyki z Worda;
  - `zadanie:` cel człowieka, nie lista kliknięć. Może ciągnąć się w kolejnych liniach.
- Pliki: tylko z `docs/samples/` po nazwie, „PDF z puli” (ścieżkę poda komenda) albo
  „nowy dokument”. Żadnych prywatnych nazw plików.
- Po zmianie: `node scripts/tester-pick-test.mjs` (sprawdza, że każda sekcja się parsuje).

Ścieżki z `docs/AUDYT-MOBILE.md`: 1 pierwsze uruchomienie, 2 otwarcie pliku, 3 pisanie,
4 formatowanie, 5 wstawianie, 6 tabele na wąskim ekranie, 7 recenzja, 8 szukaj i pisownia,
9 zapis i wydruk, 10 przerwania. Przy każdej osobie w nawiasie na końcu zadania numery ścieżek.

## organizator-iphone
- urządzenie: iphone
- waga: częste
- obszary: pdf, tabele, obrazy, pisanie, zapis, otwieranie
- kto: organizuje prace w stowarzyszeniu, wszystko robi na telefonie w biegu; w Wordzie zna tylko podstawy i przyzwyczaił się, że Enter w tabeli robi nowy akapit w komórce.
- zadanie: dostał plan prac na sobotę jako PDF z puli. Chce zamienić go na dokument Worda, dopisać jedno zadanie
  do listy, dodać wiersz w tabeli z nową osobą i godziną, wstawić zdjęcie z miejsca prac i wysłać gotowy plik
  dalej. Na koniec sprawdza, czy zapisany plik otwiera się z tymi samymi zmianami. (2, 3, 5, 6, 9)

## studentka-ipad-klawiatura
- urządzenie: ipad
- waga: częste
- obszary: pisanie, formatowanie, listy, tabele, wstawianie, wydruk
- kto: studentka chemii z iPadem i klawiaturą, pisze szybko, w Wordzie używa stylów nagłówków i skrótów Cmd+B, Cmd+Z.
- zadanie: w nowym dokumencie pisze sprawozdanie z laboratorium: tytuł, nagłówki Cel, Wykonanie, Wyniki,
  Wnioski, numerowaną listę kroków, tabelę wyników z kilkoma pomiarami i przypis ze źródłem. Chce oddać je jako
  PDF, więc drukuje do PDF i sprawdza, czy wydruk wygląda jak to, co widziała w edycji. (3, 4, 5, 9)

## nauczyciel-ocr-desktop
- urządzenie: desktop
- waga: średnie
- obszary: pdf, szukaj, pisanie, pisownia, otwieranie, zapis
- kto: nauczyciel historii, pracuje na komputerze z myszą, w Wordzie często używa Ctrl+H.
- zadanie: ma zeskanowany sprawdzian jako PDF z puli (skan). Chce zamienić go na edytowalny dokument, poprawić
  błędy rozpoznawania tekstu, zamienić we wszystkich pytaniach starą datę na nową przez Szukaj i zamień,
  przejrzeć podkreślenia pisowni i zapisać wersję dla nowej klasy. (2, 3, 8, 9)

## recenzent-telefon-poziomo
- urządzenie: iphone-landscape
- waga: rzadkie
- obszary: recenzja, pisanie, zapis
- kto: kierownik, który poprawia cudze pisma w pociągu, telefon trzyma poziomo; w Wordzie zawsze włącza Śledzenie zmian, zanim cokolwiek ruszy.
- zadanie: dostał cudze pismo docs/samples/review-sample.docx. Chce przejrzeć istniejące zmiany i komentarze,
  część przyjąć, jedną odrzucić, dopisać własny komentarz i poprawić jedno zdanie z włączonym śledzeniem
  zmian, a potem odesłać plik tak, żeby autor zobaczył jego poprawki w Wordzie. (2, 7, 9)

## niecierpliwy-przerwania
- urządzenie: iphone
- waga: rzadkie
- obszary: przerwania, pisanie, zapis, offline
- kto: zabiegany przedsiębiorca, ciągle odbiera telefony i przełącza apki, nie czyta komunikatów, w Wordzie liczy na autozapis.
- zadanie: otwiera docs/samples/sample.docx i dopisuje kilka zdań do oferty. W trakcie przerywa: obraca telefon,
  przechodzi do innej karty i wraca, traci sieć, w końcu zamyka kartę bez zapisywania i otwiera apkę od nowa.
  Chce, żeby to, co napisał, nadal tam było, i żeby dokończył ofertę bez zgadywania, czy coś zginęło. (3, 10)

## wklejacz-excel-www-mail
- urządzenie: desktop
- waga: częste
- obszary: wklejanie, tabele, formatowanie, listy, zapis
- kto: pracownica biura, składa raporty z kawałków, wszystko kopiuje i wkleja; w Wordzie zna „Zachowaj tylko tekst”.
- zadanie: w nowym dokumencie składa tygodniowy raport: wkleja tabelę skopiowaną z arkusza kalkulacyjnego,
  fragment artykułu ze strony WWW z linkami i listą, oraz kilka akapitów z maila z podpisem. Chce, żeby tabela
  została tabelą, a tekst z WWW nie przyniósł obcych czcionek, potem wyrównuje formatowanie i zapisuje. (4, 5, 6, 9)

## pierwszy-raz-iphone
- urządzenie: iphone
- waga: średnie
- obszary: pierwsze-uruchomienie, offline, otwieranie, pisanie
- kto: emerytka, która dostała link od wnuczki; telefonem posługuje się ostrożnie, Worda używała kiedyś w pracy.
- zadanie: otwiera apkę pierwszy raz. Chce zrozumieć, do czego służy, dodać ją do ekranu początkowego, otworzyć
  docs/samples/przewodnik.docx i przeczytać go, a potem napisać krótką notatkę w nowym dokumencie. Sprawdza też,
  czy apka działa, gdy włączy tryb samolotowy. (1, 2, 3)

## kierowniczka-budowy-tabela
- urządzenie: iphone
- waga: częste
- obszary: pdf, tabele, obrazy, zapis
- kto: kierowniczka budowy, rękawice zdjęte na chwilę, pisze jednym kciukiem; w Wordzie na komputerze przeciąga krawędzie kolumn.
- zadanie: harmonogram robót przyszedł jako PDF z puli z szeroką tabelą. Chce go zamienić na Worda, przewinąć
  tabelę w bok, zmienić termin w jednej komórce, dodać dwa wiersze na nowe etapy i zdjęcie usterki pod tabelą,
  a potem zapisać i wysłać inspektorowi. (2, 5, 6, 9)

## urzednik-pismo-desktop
- urządzenie: desktop
- waga: średnie
- obszary: formatowanie, uklad, pisanie, wydruk, zapis
- kto: urzędnik gminy, pisze pisma na szablonach, w Wordzie zna style, tabulatory i nagłówki z numerem strony.
- zadanie: na bazie docs/samples/styled-sample.docx przygotowuje pismo urzędowe: zmienia nagłówek z danymi
  urzędu, ustawia numer strony w stopce, wyrównuje datę do prawej tabulatorem, pisze treść dwóch stron
  i drukuje. Wydruk ma mieć marginesy i podział stron jak w Wordzie. (4, 9)

## uczen-lista-zakupow
- urządzenie: iphone
- waga: częste
- obszary: listy, pisanie, formatowanie
- kto: licealista, pisze szybko na telefonie, korzysta z autokorekty i dyktowania; Worda zna ze szkoły.
- zadanie: w nowym dokumencie robi listę rzeczy na wycieczkę klasową: punkty z podpunktami, kilka pogrubionych
  ważnych pozycji, przestawia kolejność, usuwa punkt i cofa usunięcie. Część wpisuje dyktowaniem. Chce
  wysłać to wychowawczyni. (3, 4, 9)

## tlumaczka-ipad-szukaj
- urządzenie: ipad
- waga: średnie
- obszary: szukaj, pisownia, pisanie, recenzja
- kto: tłumaczka na zleceniach, iPad bez klawiatury, w Wordzie korzysta z wyszukiwania i komentarzy dla klienta.
- zadanie: dostała docs/samples/notes-sample.docx do korekty. Chce znaleźć wszystkie wystąpienia jednego terminu
  i zamienić na poprawny, poprawić literówki podkreślone przez sprawdzanie pisowni i zostawić klientowi dwa
  komentarze z pytaniami w miejscach, których nie jest pewna. (7, 8)

## grafik-ulotka-obrazy
- urządzenie: ipad
- waga: rzadkie
- obszary: obrazy, wstawianie, uklad, formatowanie, wydruk
- kto: grafik amator, przyzwyczajony do przeciągania rzeczy palcem lub rysikiem; w Wordzie ustawia zawijanie tekstu wokół zdjęć.
- zadanie: w nowym dokumencie robi ulotkę o kiermaszu: duży tytuł, dwa zdjęcia, jedno opływane tekstem, pole
  tekstowe z godzinami, linia pozioma i strona w poziomie. Chce wydrukować ją na A4 tak, jak wygląda na
  ekranie. (4, 5, 9)

## sekretarka-formularz
- urządzenie: desktop
- waga: rzadkie
- obszary: wstawianie, formatowanie, zapis, otwieranie
- kto: sekretarka szkoły, tworzy formularze zgód, w Wordzie używa pól formularza i pól wyboru.
- zadanie: otwiera docs/samples/forms-sample.docx i przerabia go na zgodę rodzica: zmienia treść, dodaje pole
  wyboru „wyrażam zgodę”, pole daty i listę rozwijaną z klasami. Zapisuje i otwiera zapisany plik, żeby
  sprawdzić, czy pola nadal działają. (2, 5, 9)

## doktorant-spis-tresci
- urządzenie: desktop
- waga: średnie
- obszary: formatowanie, wstawianie, pisanie, szukaj, zapis, wydruk
- kto: doktorant, pisze długie teksty, w Wordzie generuje spis treści z nagłówków i używa przypisów oraz linków.
- zadanie: na bazie docs/samples/headings-sample.docx porządkuje rozdział pracy: poprawia poziomy nagłówków,
  wstawia spis treści na początku i aktualizuje go po zmianach, dodaje dwa przypisy i link do źródła,
  przeskakuje po dokumencie przez wyszukiwanie. Na koniec zapisuje i sprawdza podgląd wydruku. (4, 5, 8, 9)

## ksiegowa-wklejanie-telefon
- urządzenie: iphone-landscape
- waga: średnie
- obszary: wklejanie, tabele, pisanie, zapis
- kto: księgowa w drodze, telefon poziomo, żeby tabela się mieściła; w Wordzie wkleja dane z arkusza i poprawia liczby.
- zadanie: otwiera docs/samples/sample.docx, wkleja do niego tabelę z zestawieniem kosztów skopiowaną z maila,
  poprawia dwie kwoty, dopisuje wiersz z sumą i krótki komentarz pod tabelą. Chce wysłać plik szefowi
  i mieć pewność, że tabela nie rozjedzie się w jego Wordzie. (5, 6, 9)

## dziadek-duza-czcionka
- urządzenie: ipad
- waga: rzadkie
- obszary: pierwsze-uruchomienie, otwieranie, pisanie, formatowanie, wydruk
- kto: starszy pan z większą czcionką systemową i ciemnym motywem na iPadzie, słabo widzi małe przyciski; w Wordzie powiększał widok.
- zadanie: otwiera docs/samples/links-sample.docx, który przysłał mu syn, chce go przeczytać w powiększeniu,
  kliknąć link, potem napisać list do urzędu dużą czcionką w nowym dokumencie i go wydrukować. (1, 2, 3, 4, 9)

## pdf-umowa-poprawki
- urządzenie: iphone
- waga: częste
- obszary: pdf, pisanie, formatowanie, listy, zapis, otwieranie
- kto: freelancerka, dostaje umowy jako PDF i chce je poprawić bez komputera; w Wordzie zaznacza tekst dwuklikiem i poprawia listy punktów.
- zadanie: dostała umowę jako PDF z puli. Chce zamienić ją na Worda, zmienić swoje dane i kwotę, dopisać jeden
  punkt do listy obowiązków, usunąć akapit, który jej nie pasuje, i odesłać plik. Sprawdza, czy po zapisie
  i ponownym otwarciu numeracja punktów i układ strony są takie jak w PDF. (2, 3, 4, 9)

## redaktor-zaznaczanie-ipad
- urządzenie: ipad
- waga: średnie
- obszary: pisanie, formatowanie, wklejanie, przerwania
- kto: redaktor lokalnej gazetki, iPad trzymany raz pionowo, raz poziomo; w Wordzie zaznacza całe akapity i przenosi je kopiuj-wklej.
- zadanie: w docs/samples/selection-sample.docx przestawia kolejność akapitów, zaznaczając je palcem,
  wycinając i wklejając w innym miejscu, zmienia krój i rozmiar tytułu, w trakcie obraca iPada i przełącza się
  do innej apki po zdjęcie. Chce, żeby po powrocie kursor i przewinięcie były tam, gdzie je zostawił. (3, 4, 10)
