# Audyt UX i bug bash — nacisk na telefon i iPada

Plan na jeden dzień pracy (zaplanowane 2026-11-21, do zrobienia w najbliższy wolny dzień).
Cel: przejść całą apkę jak użytkownik, znaleźć błędy **i** zachowania, które działają,
ale są słabe (tarcie, „papercuts”), spisać je, a potem poprawiać od najważniejszych.

Zasada: najpierw **szukamy i spisujemy**, dopiero potem poprawiamy. Nic nie poprawiamy
„przy okazji” w trakcie przeglądu — inaczej ginie obraz całości.

## Urządzenia i widoki (kolejność ważności)

1. iPhone, pionowo, 390 px (WebKit, DPR 3) — główny cel.
2. iPhone poziomo (844 × 390) — mało wysokości, klawiatura zasłania pół ekranu.
3. iPad pionowo 820 px i poziomo 1180 px, także z klawiaturą zewnętrzną.
4. Mały telefon 360 px i iPhone Duo (382 / 594 / 669 px).
5. Komputer (Chromium) — kontrolnie, na końcu.

Na koniec dnia: to samo na **prawdziwym iPhonie** (PWA dodana do ekranu początkowego),
bo emulacja nie pokazuje klawiatury, gumowego przewijania ani notcha.

## Skille do użycia

| Etap | Skill | Po co |
|------|-------|-------|
| Start | `superpowers:brainstorming` (krótko) | ustalić, które ścieżki są najważniejsze |
| Przegląd telefonu | `mobile-native` | 100vh, zoom przy polach, sticky hover, notch, pull-to-refresh, długie przytrzymanie |
| Przegląd ogólny | `web-design-guidelines`, `emil-design-eng` | stany przycisków, fokus, drobiazgi wykończenia |
| Gesty | `gestures`, `apple-design` | przeciąganie, szczypanie, arkusze od dołu |
| Skrajne dane | `break-ui` | długie nazwy, puste dokumenty, wielkie tabele, sto komentarzy |
| Dostępność | `accessibility-engineering` | VoiceOver, kontrast, cele dotyku ≥ 44 px |
| Wydajność | `performance-optimization`, `animation-performance` | duży .docx/PDF na telefonie, przycinanie przewijania |
| Każdy znaleziony błąd | `superpowers:systematic-debugging` | przyczyna, nie łatka |
| Przed wydaniem | `verification-before-completion`, `code-review` | sprawdzić, nie zgadywać |

Narzędzia: Playwright MCP / chrome-devtools MCP (emulacja, zrzuty), istniejące testy
`mobile-playwright.js`, `touch-playwright.js`, `phone-bar-playwright.js`, `pinch-playwright.js`,
`npm run test:monkey` (ma przebieg MOBILE=1 WebKit), `npm run word:compare` z `VIEW=edit ENGINE=webkit DPR=2`.

## Ścieżki do przejścia (scenariusze)

Każdą ścieżkę przejść na iPhonie 390 px, potem poziomo i na iPadzie.

1. **Pierwsze uruchomienie** — instalacja PWA, okno powitalne, „Co nowego”, tryb offline.
2. **Otwarcie pliku** — .docx z Plików, PDF, skan, plik uszkodzony, plik bardzo duży.
3. **Pisanie** — kursor, zaznaczanie palcem, Enter, Backspace, autokorekta, klawiatura
   zasłaniająca kursor, przewijanie do kursora, polskie znaki, dyktowanie.
4. **Formatowanie** — pasek telefonu, wstążka, mini-pasek, style, listy, wcięcia.
5. **Wstawianie** — tabela, obraz, linia, pole, przypis; pełny cykl (wstaw, usuń, cofnij,
   kopiuj/wklej, zapis → otwarcie).
6. **Tabele na wąskim ekranie** — przewijanie w bok, zmiana szerokości, dodawanie wierszy.
7. **Recenzja** — komentarze, śledzenie zmian, przechodzenie po zmianach.
8. **Szukaj i zamień**, sprawdzanie pisowni.
9. **Zapis, udostępnianie, wydruk/PDF** — arkusz udostępniania iOS, nazwa pliku, podgląd wydruku.
10. **Przerwania** — przejście do innej apki i powrót, obrót ekranu w trakcie edycji,
    brak sieci, zamknięcie karty (czy szkic przetrwał).

## Na co patrzeć (lista kontrolna)

- Coś zasłonięte przez klawiaturę, notch, pasek Safari lub dolny pasek apki.
- Coś ucięte zamiast przewijane; poziome przewijanie całej strony.
- Cele dotyku mniejsze niż 44 px albo za blisko siebie.
- Zoom strony po dotknięciu pola, przypadkowe zaznaczenie tekstu przycisku.
- Skoki i przeładowania widoku, utrata pozycji przewijania lub kursora.
- Opóźnienie po dotknięciu, brak informacji zwrotnej, „czy to zadziałało?”.
- Dwa miejsca na tę samą funkcję albo ta sama ikona dla różnych funkcji.
- Komunikaty niejasne albo kłamiące (toast mówi „zapisano”, a nie zapisano).
- Różnice Chromium kontra WebKit; Edycja kontra wydruk kontra Word.
- Ciemny motyw, większa czcionka systemowa, „Ogranicz ruch”.

## Jak spisywać znaleziska

Jedna tabela w tym pliku (sekcja niżej), jeden wiersz na problem:

| # | Waga | Gdzie (urządzenie, ścieżka) | Co się dzieje | Czego oczekujemy | Zrzut | Źródło | Stan |
|---|------|-----------------------------|---------------|------------------|-------|--------|------|

Wagi: **Krytyczny** (utrata danych, nie da się czegoś zrobić) → **Ważny** (wyraźnie
przeszkadza) → **Drobny** (papercut) → **Pomysł** (na później).

Na koniec dnia: ocena wg 6 wymiarów z `../.claude/rules/review-and-research.md`
i heurystyk Nielsena, potem kolejność poprawek (Krytyczne najpierw, każda z testem).

## Znaleziska

Źródło: `ręcznie` albo `testerzy <przebieg>` (komenda `/testerzy`, raporty w `output/testerzy/`).
Stan: `nowe` → `naprawione vX.YY.ZZ` / `odrzucone (powód)`.

| # | Waga | Gdzie (urządzenie, ścieżka) | Co się dzieje | Czego oczekujemy | Zrzut | Źródło | Stan |
|---|------|-----------------------------|---------------|------------------|-------|--------|------|
| 1 | Ważny | iPhone 390, Czytanie i Edycja; tabela z PDF-a (ścieżki 2, 6) | Wiersze tabeli tracą ramki (na komputerze są); wąska kolumna łamie słowa litera po literze („Ilość uczestników” w 15 liniach), nagłówek tabeli na pół ekranu | Ramki jak w pliku; kolumny tak, żeby słowa się nie rwały, a za szeroka tabela przewija się w bok | output/testerzy/2026-11-27-1301/spr-2/shot-001.png, shot-004.png | testerzy 2026-11-27-1301 | naprawione v1.01.41: wcięcia akapitów w komórce przycięte do 10% szerokości komórki (z miejscem na numer listy), ramki komórek ≥ 1 px, kolumny dzielą zapas szerokości pod najdłuższe słowa |
| 2 | Ważny | Komputer, Edycja; dokument 2-stronicowy z nagłówkiem i „Strona X z Y” (ścieżki 4, 9) | Nagłówek tylko nad stroną 1, stopka tylko na końcu ostatniej strony i brzmi „Strona 1 z 1”; dół strony 1 i góra strony 2 puste. Podgląd wydruku poprawny | Nagłówek i stopka na każdej stronie z numerem jak w Wordzie i w podglądzie wydruku | output/testerzy/2026-11-27-1301/t1-urzednik/shot-027.png, shot-028.png | testerzy 2026-11-27-1301 | naprawione v1.01.40: kopie nagłówka i stopki na każdej stronie Edycji, numery stron z odstępów (sectionPages liczyło zawsze 1) |
| 3 | Krytyczny | iPhone, Edycja; PDF z tabelą scaloną (umowa) → zmiana daty, punkt listy, Cofnij/Ponów, usunięcie akapitu → zapis (ścieżki 2, 3, 9) | W zapisanym pliku treść komórek tabeli przesunięta o komórkę (nagłówek „ZAKRES UBEZPIECZENIA” znika, opisy i kwoty w złych wierszach). Świeża konwersja i pojedyncze edycje na komputerze zapisują się dobrze — najkrótszej sekwencji jeszcze brak | Zapis nie zmienia komórek, których nie ruszano | output/testerzy/2026-11-27-1315/t6-umowa/shot-007.png → shot-022.png; plik output/testerzy/2026-11-27-1315/t6-umowa/pobrane/umowa.docx | testerzy 2026-11-27-1315 | naprawione v1.01.39 (przyczyna: <br> przeglądarki w opróżnionym akapicie, Delete omijał Edycję) |
| 4 | Ważny | iPhone, arkusz kart z kursorem w tabeli (ścieżki 5, 6) | Karta „Tabela” stoi na końcu rzędu kart, poza ekranem; jedyną wskazówką jest wygaszenie pod „Gotowe”. (Sprostowanie: rząd przewija się palcem — testerzy nie mogli, bo sterownik przewijał tylko w pionie; dodane --dx) | Karta Tabela widoczna od razu, gdy kursor jest w tabeli | output/testerzy/2026-11-27-1315/t5-kierowniczka/shot-016.png | testerzy 2026-11-27-1315 | naprawione v1.01.39 (rząd przewija się do karty Tabela) |
| 5 | Ważny | iPhone, Wstaw → Obraz (ścieżka 5) | Po wstawieniu obrazu w środku strony zostaje drugi pasek szukania i arkusz (duch) | Jeden pasek, arkusz zamknięty albo na swoim miejscu | output/testerzy/2026-11-27-1315/t5-kierowniczka/shot-040.png, shot-041.png | testerzy 2026-11-27-1315 | odrzucone: tester wstawił jako zdjęcie zrzut ekranu apki — „duch” paska to treść obrazka; testerzy dostali zdjęcie testowe |
| 6 | Ważny | Komputer, Wstaw → Spis treści → Aktualizuj (ścieżka 4) | Po aktualizacji przy każdym wpisie strona „1” w dokumencie 2-stronicowym, choć dymek mówi „zaktualizowany”; wpisy 3. poziomu zawijają numer strony do nowej linii | Numery stron jak w Wordzie, numer w tym samym wierszu | output/testerzy/2026-11-27-1315/t4-doktorant/shot-027.png | testerzy 2026-11-27-1315 | naprawione v1.01.40: numery z granic stron Edycji (to samo min-height co w #2); zawijanie 3. poziomu w #13 |
| 7 | Ważny | Komputer, Edycja z przypisami (ścieżki 5, 9) | Edycja: przypisy razem na końcu ostatniej strony, 2 strony; podgląd wydruku: 3 strony, przypis 1 na stronie 1. Ta sama rodzina co #2 (Edycja ≠ wydruk) | Edycja = wydruk = Word | output/testerzy/2026-11-27-1315/t4-doktorant/shot-019.png, shot-032.png | testerzy 2026-11-27-1315 | naprawione v1.01.42: page-breaks.js w Edycji liczy strony z includeNotes i stawia żywe, edytowalne <li> przypisów dolnych w .dwb-page-notes nad dolnym marginesem każdej strony (poza <article>, z przywracaniem kursora i porządku w clear) |
| 8 | Ważny | iPhone, wstawienie zdjęcia przy kursorze (ścieżka 5) | Pod zdjęciem duża pusta przestrzeń | Zdjęcie w wierszu tekstu bez pustego obszaru | output/testerzy/2026-11-27-1315/t1-organizator/shot-028.png, shot-030.png | testerzy 2026-11-27-1315 | naprawione v1.01.40: w widoku telefonu ramka obrazu miała wysokość z pliku, a obraz zmniejszony do ekranu zostawiał pod sobą pustkę |
| 9 | Ważny | iPad, komentarze i stuknięcia po nich (ścieżka 7) | Błąd JS „undefined is not an object (evaluating 'p.classList')”; drugi komentarz przypięty do innego słowa niż zaznaczone; „Komentarz” kilka razy mówi „Kliknij w tekst…” mimo kursora — do odtworzenia razem | Komentarz na zaznaczonym tekście, bez błędu | output/testerzy/2026-11-27-1315/t3-tlumaczka/shot-023.png…026.png, shot-031.png | testerzy 2026-11-27-1315 | naprawione v1.01.42: selectionInParagraph sprawdza ia/ib < 0 (przypisy nie wywołują list[-1].classList); lastDocCaret/restoreDocCaret i mini-bar nie cofają zaznaczenia z zablokowanego akapitu do poprzedniego słowa (komunikat o blokadzie) |
| 10 | Ważny | iPhone, PDF → zapis → ponowne otwarcie (ścieżki 2, 9) | Po ponownym otwarciu zmienia się układ akapitów („PAKIET WSPARCIE…” w osobnej linii, znika odstęp, tabela niżej) | Ten sam układ po otwarciu zapisanego pliku | output/testerzy/2026-11-27-1315/t6-umowa/shot-016.png → shot-021.png | testerzy 2026-11-27-1315 | naprawione v1.01.39 (skutek #3 — po naprawie #3 wszystkie 518 akapitów po usunięciu akapitu, zapisie i ponownym otwarciu mają ten sam układ; potwierdzone też w 2026-11-27-1505 t1) |
| 11 | Ważny | iPhone, lista: zaznaczenie całego tekstu punktu → Backspace (ścieżka 3) | Kursor znika, pisanie nie działa, wyjście tylko Cofnij. Na komputerze nie odtworzone | Pusty punkt z kursorem, jak w Wordzie | output/testerzy/2026-11-27-1301/t2-uczen/shot-040.png, shot-043.png | testerzy 2026-11-27-1301 | naprawione v1.01.42: doc-selection.js przejmuje usunięcie całego tekstu punktu listy przed WebKitem (WebKit usuwał element <p> i zostawiał gołe <br> w <article>), zachowuje pusty punkt z <br data-dwb-ph> i kursorem |
| 12 | Drobny | różne | Nowy wiersz tabeli ma inną czcionkę (iPhone, t1); Backspace w pustym akapicie po usunięciu zaznaczenia nic nie robi (iPhone, t6); pole „Szukaj” bardzo wąskie i panel „Znajdź i zamień” zasłania dokument (iPad, t3); podświetlenie trafienia zostaje po zmianie dokumentu, Link bez zaznaczenia nic nie robi (komputer, t4); tabulator prawy działa dopiero po przerysowaniu (komputer, 1301 t1) | — | raporty w output/testerzy/2026-11-27-1315/ | testerzy 2026-11-27-1315 | naprawione v1.01.42 + v1.01.45 + v1.01.52 (v1.01.52 domyka ostatnie dwa punkty: tabulator prawy/środkowy/dziesiętny wstawiany klawiszem Tab działa od razu na żywo przez span.docx-tab + layoutTabStops; panel „Znajdź i zamień” na tablecie 769–1023 px oraz przypięcie pinezką dokują się obok dokumentu, bez zasłaniania tekstu) |

| 13 | Ważny | Komputer, spis treści z nagłówkiem 3. poziomu (ścieżka 4) | Wpis 3. poziomu zawijał numer strony do nowej linii (tabulator prawy kończył się co do ułamka piksela na krawędzi) | Numer w tej samej linii | output/testerzy/2026-11-27-1505/t4-doktorant/shot-007.png | testerzy 2026-11-27-1505 | naprawione v1.01.40 |
| 14 | Drobny | różne | iPhone: po schowaniu klawiatury rząd poleceń zmienia pozycje przycisków (t1, t2); po przewinięciu górny pasek zwija się do kreski, stuknięcie w nią trafia w skrót sekcji i skacze na początek (t2); po „Pobierz” brak potwierdzenia zapisu (t2). iPad: pierwsze przeciągnięcie po zamknięciu dymka komentarza nic nie zaznacza (t3); podświetlenie wyszukiwania zostaje na tytule okna „Zapisz jako” (t3). Komputer: dymek po wstawieniu spisu odsyła do „prawy klik → Aktualizuj pole” (to opis Worda, w apce tego nie ma) (t4); „Str. 1/2” na pasku nie zmienia się po przewinięciu na stronę 2 (t4). Wklejka tabeli HTML bez kroju zmienia krój na Calibri 11 w dokumencie w Aptos 12 (t5) | — | raporty w output/testerzy/2026-11-27-1505/ | testerzy 2026-11-27-1505 | prawie zamknięte v1.01.42 + v1.01.45 (komunikat spisu treści, próg „Str. N/M”, wyższy uchwyt nagłówka 28 px, brak przeciekania ::highlight na <dialog>, niepochłanianie pierwszego przeciągnięcia po dymku komentarza, Aptos 12 w domyślnym styles.xml przy wklejce tabeli; toast po Pobierz już istnieje w ui-controls.js); zostaje ewentualnie stabilny układ rzędu po schowaniu klawiatury |
| 15 | Ważny | iPhone, PDF → Word z numerowanymi punktami, dopisanie / usunięcie punktu (ścieżki 2, 3) | Każdy punkt to osobna lista (numId 1, 2, 3…), więc dopisany lub usunięty punkt nie przenumerowuje reszty (po usunięciu 2. zostaje 1., 3.); punkt jest rozcięty na dwa akapity: pierwsze dwie linie z wcięciem wiszącym, reszta osobny akapit bez wcięcia, więc Enter i End działają „na kawałku” | Jedna lista z ciągłą numeracją, punkt = jeden akapit | output/testerzy/2026-11-27-1856/t6-umowa/shot-017.png, shot-020.png; output/testerzy/2026-11-27-1856/spr-2/shot-003.png | testerzy 2026-11-27-1856 | naprawione v1.01.49: pdf-layout.js — wiersz zawinięty pod sam numer to dalszy ciąg punktu (wcięcie od numeru, spacja za numerem w:suff jak w PDF); pdf-convert.js — kolejne punkty po nagłówku (także z podpunktami) dołączają do tej samej listy Worda, gdy numeracja idzie dalej; sprawdzone w Wordzie (1.–4. ciągle, 7 stron jak PDF), na iPhonie (usunięcie przenumerowuje, Enter daje nowy punkt, zapis → otwarcie), pdf:score 87,1 → 87,2 %. Dywizy łamania („zachowy- wać”) zostają — patrz #16 |
| 16 | Drobny | różne | iPhone: w PDF → Word zostają wyrazy z dywizem łamania („zachowy- wać”, „gościn- ności”) (t6, sprawdzić, co robi Word). iPad: Wytnij całego akapitu zaznaczonego palcem zostawia pusty akapit (t5; Word z inteligentnym zaznaczaniem akapitu zabiera znak akapitu — sprawdzić wyrocznią); okno komentarza otwiera się w lewym górnym rogu i zasłania komentowany tekst (t3); po „Zamień wszystkie” (5 zamian) panel pokazuje dwa razy „Brak trafień” (t3). iPhone poziomo: czerwone paski zmian z marginesu przebijają przez panel „Zmiany i komentarze” (t1). Komputer: tło komórki nagłówka przepada przy wklejce tabeli z arkusza (t4) | — | output/testerzy/2026-11-27-1856/t6-umowa/shot-004.png, t5-redaktor/shot-005.png, spr-3/shot-009.png, t3-tlumaczka/shot-006.png, t1-recenzent/shot-003.png, t4-wklejacz/shot-004.png | testerzy 2026-11-27-1856 | prawie zamknięte v1.01.50 (paski zmian .dwb-markup pod panelem bocznym z-index 20; jeden napis „Brak trafień” w panelu Znajdź i zamień; okienko nowego komentarza z ukrytej/znikającej kotwicy przy zaznaczonym tekście zamiast rogu (8, 6); wklejka tabeli HTML z arkusza zachowuje tło komórek z inline style i <style>; Wytnij całego akapitu usuwa pusty akapit jak prawdziwy Word; dodatkowo naprawiony test pdf-import-playwright.js — zamykanie kart w tle przed skanem); dywizy łamania z PDF na razie bez zmian (patrz #15) |
Odczucie 2026-11-27-1315 (6 testerów, 281 czynności): reakcja mediana iPhone 14 ms, iPad 9 ms, komputer 3 ms, czynności > 300 ms: 0. Martwe stuknięcia: iPhone 17/87, iPad 6/35, komputer 1/30 — zawyżone (stuknięcie ustawiające tylko kursor liczone jako martwe; poprawka sterownika po przeglądzie). Oceny 1–5 brak (agent z instrukcją sprzed sekcji „Odczucie”). Skoki widoku: 13, część to zamierzone zwinięcie paneli.

Odrzucone w 2026-11-27-1315: Tab w ostatniej komórce nie dodaje wiersza (działa na komputerze i iPhonie), Ctrl+X na Apple (właściwy jest Cmd+X), „Zapisz” na iPhonie otwiera „Zapisz jako” (brak dostępu do plików w Safari), pytanie o zapis po świeżej konwersji (to nowy, niezapisany dokument), przypadkowa zmiana orientacji (chybione stuknięcie testera), brak ostrzeżenia o piśmie odręcznym w skanie (pomysł, nie błąd).

Odczucie 2026-11-27-1505 (6 testerów, 188 czynności, pomiary po poprawkach sterownika): oceny 1–5 średnio — wysiłek 3,8, pewność 3,7, płynność 4,8, jak Word 3,7. Martwe stuknięcia: iPhone 0/51, iPad 1/27 (w tło), komputer 0/21. Reakcja mediana / p90: iPhone 23/52 ms, iPad 10/26 ms, komputer 4/18 ms; > 300 ms: 1 (iPad, „Skanuj dokument”). Skoki: 8, w większości zamierzone (arkusz kart, skrót sekcji).

Odrzucone w 2026-11-27-1505: duble #1 (wąskie kolumny tabeli z PDF na telefonie, t6, t1 „zł”), #7 (Edycja 2 strony, wydruk 3 — przypisy, t4), #12 (wąskie pole Szukaj, t3); Home/End w komórce idą do początku/końca linijki ekranowej (tak robi Word); nowe wiersze tabeli bez numeracji (Word też nie numeruje); wklejona tabela bez krawędzi (źródło HTML bez krawędzi — do sprawdzenia w Wordzie); stuknięcie w samą krawędź przycisku Zapisz; brak dwukliku w sterowniku (ograniczenie narzędzia). Potwierdzone w przebiegu: naprawy #3 (umowa po edycji i zapisie — tabela dobra), #4 (karta Tabela w arkuszu), #8 (zdjęcie bez pustki).

Odczucie 2026-11-27-1856 (6 testerów, 196 czynności): oceny 1–5 średnio — wysiłek 3,3, pewność 3,5, płynność 4,5, jak Word 3,2 (najniżej t6, PDF z punktami: 2/2/3/2). Martwe stuknięcia: iPhone 0/125 (5 zapisanych to Cmd+Z bez czego cofać, End i stuknięcie poza ekran), iPad 1/56 (Wklej z paska — ograniczenie schowka, patrz niżej), komputer 0/15. Reakcja mediana / p90: iPhone 22/36 ms, iPad 11/22 ms, komputer 4/6 ms; > 300 ms: 0. Skoki: 4 (przewinięcie do kursora nad paskiem przy pisaniu na iPhonie, rozwinięcie zwiniętego nagłówka) — zamierzone.

Odrzucone w 2026-11-27-1856: brak ostatniej linii podpisu z maila (tester podał schowkowi tekst bez tej linii; HTML bez formatowania wkleja się jako tekst — dane testu), stuknięcie w akapit z PDF stawia kursor gdzie indziej (odtworzone 4 razy — kursor trafia dokładnie; tester stukał w współrzędne sprzed przewinięcia, a podświetlenie akapitu z kursorem to celowe .dwb-caret-p), okno komentarza znika po stuknięciu w pole imienia (stuknięcie poza oknem; odtworzone — imię i komentarz zapisują się), pływający pasek zmienia miejsce (stare współrzędne testera), podwójna spacja „pozycji  z” (jest w pliku źródłowym), podwójny odstęp z pustego akapitu maila i brak ikony opcji wklejania (jest „Wklej jako zwykły tekst” w menu), panel Wstaw zostaje po wstawieniu obrazu (ma „Gotowe”), „Zapisz” na iPadzie otwiera „Zapisz jako” (Safari nie nadpisuje plików), kursor po Cofnij (niepewne, pomiar sterownika). Do sprawdzenia na prawdziwym iPadzie: przycisk Wklej — w emulacji schowek odmawia i jest komunikat „użyj ⌘V”; prawdziwe Safari powinno pokazać dymek „Wklej” po stuknięciu. Uwaga do sterownika: `caret.offset` po stuknięciu w długi akapit z PDF bywa nieaktualny (pokazał „si|ę” zamiast „miejscac|h”), co myli testerów.