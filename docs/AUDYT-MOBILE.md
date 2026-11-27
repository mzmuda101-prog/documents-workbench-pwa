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
| 1 | Ważny | iPhone 390, Czytanie i Edycja; tabela z PDF-a (ścieżki 2, 6) | Wiersze tabeli tracą ramki (na komputerze są); wąska kolumna łamie słowa litera po literze („Ilość uczestników” w 15 liniach), nagłówek tabeli na pół ekranu | Ramki jak w pliku; kolumny tak, żeby słowa się nie rwały, a za szeroka tabela przewija się w bok | output/testerzy/2026-11-27-1301/spr-2/shot-001.png, shot-004.png | testerzy 2026-11-27-1301 | nowe |
| 2 | Ważny | Komputer, Edycja; dokument 2-stronicowy z nagłówkiem i „Strona X z Y” (ścieżki 4, 9) | Nagłówek tylko nad stroną 1, stopka tylko na końcu ostatniej strony i brzmi „Strona 1 z 1”; dół strony 1 i góra strony 2 puste. Podgląd wydruku poprawny | Nagłówek i stopka na każdej stronie z numerem jak w Wordzie i w podglądzie wydruku | output/testerzy/2026-11-27-1301/t1-urzednik/shot-027.png, shot-028.png | testerzy 2026-11-27-1301 | nowe |
| 3 | Krytyczny | iPhone, Edycja; PDF z tabelą scaloną (umowa) → zmiana daty, punkt listy, Cofnij/Ponów, usunięcie akapitu → zapis (ścieżki 2, 3, 9) | W zapisanym pliku treść komórek tabeli przesunięta o komórkę (nagłówek „ZAKRES UBEZPIECZENIA” znika, opisy i kwoty w złych wierszach). Świeża konwersja i pojedyncze edycje na komputerze zapisują się dobrze — najkrótszej sekwencji jeszcze brak | Zapis nie zmienia komórek, których nie ruszano | output/testerzy/2026-11-27-1315/t6-umowa/shot-007.png → shot-022.png; plik output/testerzy/2026-11-27-1315/t6-umowa/pobrane/umowa.docx | testerzy 2026-11-27-1315 | nowe |
| 4 | Ważny | iPhone, arkusz kart (Główne/Wstaw/Układ/Re…) z kursorem w tabeli (ścieżki 5, 6) | Ostatnie karty ucięte pod „Gotowe”, rząd nie przewija się; karty „Tabela” nie da się dosięgnąć — dwaj testerzy nie znaleźli dodawania wiersza (tylko Tab) | Rząd kart przewija się (zasada „nic nie ucina się po cichu”), karta Tabela widoczna przy kursorze w tabeli | output/testerzy/2026-11-27-1315/t5-kierowniczka/shot-016.png, output/testerzy/2026-11-27-1315/t1-organizator/shot-016.png | testerzy 2026-11-27-1315 | nowe |
| 5 | Ważny | iPhone, Wstaw → Obraz (ścieżka 5) | Po wstawieniu obrazu w środku strony zostaje drugi pasek szukania i arkusz (duch) | Jeden pasek, arkusz zamknięty albo na swoim miejscu | output/testerzy/2026-11-27-1315/t5-kierowniczka/shot-040.png, shot-041.png | testerzy 2026-11-27-1315 | nowe |
| 6 | Ważny | Komputer, Wstaw → Spis treści → Aktualizuj (ścieżka 4) | Po aktualizacji przy każdym wpisie strona „1” w dokumencie 2-stronicowym, choć dymek mówi „zaktualizowany”; wpisy 3. poziomu zawijają numer strony do nowej linii | Numery stron jak w Wordzie, numer w tym samym wierszu | output/testerzy/2026-11-27-1315/t4-doktorant/shot-027.png | testerzy 2026-11-27-1315 | nowe |
| 7 | Ważny | Komputer, Edycja z przypisami (ścieżki 5, 9) | Edycja: przypisy razem na końcu ostatniej strony, 2 strony; podgląd wydruku: 3 strony, przypis 1 na stronie 1. Ta sama rodzina co #2 (Edycja ≠ wydruk) | Edycja = wydruk = Word | output/testerzy/2026-11-27-1315/t4-doktorant/shot-019.png, shot-032.png | testerzy 2026-11-27-1315 | nowe |
| 8 | Ważny | iPhone, wstawienie zdjęcia przy kursorze (ścieżka 5) | Pod zdjęciem duża pusta przestrzeń | Zdjęcie w wierszu tekstu bez pustego obszaru | output/testerzy/2026-11-27-1315/t1-organizator/shot-028.png, shot-030.png | testerzy 2026-11-27-1315 | nowe |
| 9 | Ważny | iPad, komentarze i stuknięcia po nich (ścieżka 7) | Błąd JS „undefined is not an object (evaluating 'p.classList')”; drugi komentarz przypięty do innego słowa niż zaznaczone; „Komentarz” kilka razy mówi „Kliknij w tekst…” mimo kursora — do odtworzenia razem | Komentarz na zaznaczonym tekście, bez błędu | output/testerzy/2026-11-27-1315/t3-tlumaczka/shot-023.png…026.png, shot-031.png | testerzy 2026-11-27-1315 | do odtworzenia |
| 10 | Ważny | iPhone, PDF → zapis → ponowne otwarcie (ścieżki 2, 9) | Po ponownym otwarciu zmienia się układ akapitów („PAKIET WSPARCIE…” w osobnej linii, znika odstęp, tabela niżej) | Ten sam układ po otwarciu zapisanego pliku | output/testerzy/2026-11-27-1315/t6-umowa/shot-016.png → shot-021.png | testerzy 2026-11-27-1315 | nowe |
| 11 | Ważny | iPhone, lista: zaznaczenie całego tekstu punktu → Backspace (ścieżka 3) | Kursor znika, pisanie nie działa, wyjście tylko Cofnij. Na komputerze nie odtworzone | Pusty punkt z kursorem, jak w Wordzie | output/testerzy/2026-11-27-1301/t2-uczen/shot-040.png, shot-043.png | testerzy 2026-11-27-1301 | do odtworzenia (telefon) |
| 12 | Drobny | różne | Nowy wiersz tabeli ma inną czcionkę (iPhone, t1); Backspace w pustym akapicie po usunięciu zaznaczenia nic nie robi (iPhone, t6); pole „Szukaj” bardzo wąskie i panel „Znajdź i zamień” zasłania dokument (iPad, t3); podświetlenie trafienia zostaje po zmianie dokumentu, Link bez zaznaczenia nic nie robi (komputer, t4); tabulator prawy działa dopiero po przerysowaniu (komputer, 1301 t1) | — | raporty w output/testerzy/2026-11-27-1315/ | testerzy 2026-11-27-1315 | nowe |

Odczucie 2026-11-27-1315 (6 testerów, 281 czynności): reakcja mediana iPhone 14 ms, iPad 9 ms, komputer 3 ms, czynności > 300 ms: 0. Martwe stuknięcia: iPhone 17/87, iPad 6/35, komputer 1/30 — zawyżone (stuknięcie ustawiające tylko kursor liczone jako martwe; poprawka sterownika po przeglądzie). Oceny 1–5 brak (agent z instrukcją sprzed sekcji „Odczucie”). Skoki widoku: 13, część to zamierzone zwinięcie paneli.

Odrzucone w 2026-11-27-1315: Tab w ostatniej komórce nie dodaje wiersza (działa na komputerze i iPhonie), Ctrl+X na Apple (właściwy jest Cmd+X), „Zapisz” na iPhonie otwiera „Zapisz jako” (brak dostępu do plików w Safari), pytanie o zapis po świeżej konwersji (to nowy, niezapisany dokument), przypadkowa zmiana orientacji (chybione stuknięcie testera), brak ostrzeżenia o piśmie odręcznym w skanie (pomysł, nie błąd).
