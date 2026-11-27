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
