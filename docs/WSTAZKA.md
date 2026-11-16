# Wstążka — wdrożenie makiety v4

Wzór: `docs/makiety/wstazka-v4.html` (zatwierdzona 2026-11-16). Zasada: **układ, rozmieszczenie,
teksty i zachowania 1:1 z makietą**. Animacje i przejścia mogą się różnić. Każde konieczne
odstępstwo najpierw do decyzji Mateusza.

Krój pisma i kolory zostają z aplikacji (tokeny w `styles/app.css`): makieta pokazuje układ,
nie zmienia motywu.

## Etapy (każdy to osobne wydanie)

| # | Etap | Co wchodzi | Stan |
|---|------|------------|------|
| 1 | Rama: dwa rzędy u góry (1a: v1.01.00) | Wiersz 1: logo, Plik, karty, Szukaj, tryb (menu), Zapisz, ⋯. Wiersz 2: Cofnij/Ponów + polecenia karty w jednej linii (wstążka uproszczona). Karty: Narzędzia główne, Wstaw, Układ, Recenzja, Korespondencja, Widok, Asystent, kontekstowa Tabela. Przewijany rząd z podpowiedzią (wygaszenie, strzałka „+N”, jednorazowe drgnięcie). Komputer (≥ 769 px). | w toku |
| 2 | Plik za kulisami (v1.01.02) | Ekran Plik: Nowy, Otwórz, Zdjęcie, Informacje, Zapisz jako, Eksportuj, Drukuj, Porównaj, Zamknij, Opcje, Pomoc, Zgłoś błąd. Menu ⋯ odchudzone (aktualizacja, wersja, inne aplikacje). | — |
| 3 | Szyna i panele | Szyna 44 px (Nawigacja, Znajdź i zamień, Komentarze, Korekta, Asystent) z licznikami; panel wysuwany na margines albo przypięty. Zastępuje stały panel boczny. | — |
| 4 | Dolny pasek | Karty otwartych dokumentów (przewijane) + strona, wyrazy, błędy, język, powiększenie, „Sugerowanie”. | — |
| 5 | Mini pasek | Nad zaznaczeniem, pod nim gdy brak miejsca, gaśnie przy przewijaniu. | — |
| 6 | Opcje wstążki | Klasyczna (duże przyciski, nazwy grup) / Uproszczona; Zawsze / Chowaj przy przewijaniu / Tylko karty (Ctrl+F1); gęstość Mysz/Dotyk. | — |
| 7 | Telefon | Pasek u góry (wstecz, cofnij, nazwa, tryb, ⋯); rząd poleceń nad klawiaturą; arkusz kart w miejscu klawiatury; Czytanie bez dodatkowych przycisków. | — |
| 8 | Znaczniki recenzji | Licznik „N zmian · M komentarzy” nad kolumną dymków, poprzedni/następny, Akceptuj wszystkie; Recenzja → Znaczniki: Wszystkie / Proste / Bez. Reszta jak dziś (`review-margin.js`). | — |

## Etap 1a — co weszło (2026-11-16)

`app/ribbon.js` + znaczniki `#ribbon` w `index.html` + sekcja „Wstążka” w `styles/app.css`.
Od 769 px: rząd 1 (logo, Plik, karty, Szukaj, tryb w menu, Zapisz, ⋯), rząd 2 (Cofnij/Ponów +
polecenia karty), przewijane rzędy z podpowiedzią, karta kontekstowa Tabela, w Czytaniu nazwa
pliku i „Edytuj” zamiast kart i „Zapisz”. Powiększenie na pasku stanu (jak dolny pasek makiety).
Telefon bez zmian do etapu 7.

Tymczasowe odstępstwa (znikają w kolejnych etapach):
- 🔧 panelu bocznego stoi w rzędzie 1 — do etapu 3 (szyna);
- „Plik” otwiera sekcję Plik w panelu — do etapu 2 (ekran Plik);
- „Szukaj” szuka tekstu w dokumencie, jeszcze nie poleceń;
- karta Układ: menu „Układ ▾” + „Marginesy i układ strony…” zamiast osobnych list Marginesy /
  Orientacja / Rozmiar / Kolumny / Podziały / Wyrównanie w pionie — etap 1b;
- brak jeszcze: Wklej / Wytnij / Kopiuj, osobno Punktory i Numerowanie, Zmniejsz/Zwiększ wcięcie,
  osobno Wyróżnienie; w Recenzji Akceptuj / Odrzuć / Poprzednia / Następna, Język, Dostępność — 1b.

Rozstrzygnięcia, których makieta nie pokazywała:
- tryb skupienia chowa rząd poleceń (rząd 1 zostaje, żeby był Zapisz i wyjście);
- karta Wstaw ma też istniejące polecenia spoza listy makiety (Nagłówek i stopka, Zakładka,
  Komentarz, Spis treści, pełne menu „Wstaw”) — żeby nic, co już działało, nie zniknęło.

## Etap 1b — co weszło

- Narzędzia główne: Schowek (Wklej, Wytnij, Kopiuj, Malarz formatów), osobno Wyróżnienie (żółte,
  jak domyślne w Wordzie) obok koloru czcionki, Punktory ▾ (▾ = dawne menu listy: format i poziom)
  i Numerowanie, Zmniejsz / Zwiększ wcięcie (w liście = poziom punktu, jak w Wordzie).
- Wklej: `navigator.clipboard.read()` → to samo zdarzenie „paste” co Ctrl/⌘+V (ta sama obsługa
  formatowania i obrazów). Chrome pyta raz o zgodę, Safari pokazuje dymek; odmowa = komunikat
  „użyj Ctrl/⌘+V”.
- Układ: osobne listy Marginesy, Orientacja, Rozmiar, Kolumny, Podziały, Wyrównanie w pionie
  (części tego samego okienka), „Układ” = całe okienko, Odstępy i wcięcia, Tabulatory….
- Recenzja: Sugerowanie, Znaczniki ▾ (Wszystkie / Proste / Bez), Akceptuj ▾ / Odrzuć ▾
  (wszystkie lub po kolei), Poprzednia / Następna (przegląd po kolei), Komentarz, Usuń wszystkie.
- Skróty w dymkach nowych przycisków (Ctrl/⌘+V/X/C, Enter, K, H, Alt+M, Alt+F, Alt+D).

Wciąż brakuje: Język i Dostępność w Recenzji (osobny etap — wymagają nowych okienek).

## Etap 2 — ekran Plik

`app/backstage.js` + `#backstage` w `index.html`. Wjeżdża pod rzędem kart tym samym ruchem co dawny
panel boczny (280 ms, `--ease-drawer`). Sekcje: Nowy (kafelki + ostatnio otwierane), Informacje
(panele Statystyki i Metadane), Eksportuj, Porównaj wersje, Opcje (motyw, język, Asystent AI),
Pomoc i skróty (Skróty, Przewodnik, panel informacji). Akcje (Otwórz, Zdjęcie, Zapisz jako, Drukuj,
Zamknij, Zgłoś błąd) klikają te same przyciski co dawne menu ⋯. `openToolPanel()` dla panelu,
który mieszka na ekranie Plik, otwiera właściwą sekcję. Menu ⋯: Odśwież, inne aplikacje, stan,
wersja (+ „Aktualizuj” w rzędzie 1).

Odstępstwo: czwarty kafelek „Nowy” to „Przewodnik po aplikacji” (istniejąca funkcja) zamiast
„Pismo z placeholderami” z makiety — takiego szablonu jeszcze nie ma; szablony są w oknie
„Pusty dokument”.

## Poprawki po przeglądzie (decyzje Mateusza, 2026-11-16)

- Widok bez „Czytanie” (tryb jest w przycisku u góry) i bez grupy Powiększenie (zoom jest na
  dolnym pasku) — odstępstwo od makiety na prośbę Mateusza: bez dublowania.
- Pełny ekran ma własną ikonę (inną niż Tryb skupienia i Dopasuj); przy samych ikonach podpisy
  (Skupienie, Pełny ekran, Struktura, Pasek sekcji, Widok, Statystyki, Recenzja).
- Snippety tylko na karcie Wstaw (nie dublują się w Korespondencji).
- „Przepisz” z ikoną zamiast znaku ✦; pole Szukaj szersze (napis się nie ucina).
- Dziennik techniczny („Log”) przy wstążce schowany (trafia do zgłoszenia: Plik → Zgłoś błąd),
  widać go z `?debug`; powiadomienia stoją nad dolnym paskiem.
- Druga runda (bez dubli w menu): menu Odstępy na Narzędziach głównych = interlinia i odstęp
  przed/po (wcięcia ± są przyciskami, „Odstępy i wcięcia…” i Tabulatory… na karcie Układ);
  Punktory ▾ bez „Punktory/Numerowanie” (są obok); pełne menu „Wstaw” i całe okienko „Układ”
  schowane (każda pozycja ma przycisk; pola formularza → Wstaw › Pole formularza ▾);
  Recenzja: Akceptuj wszystkie / Odrzuć wszystkie / Przejrzyj po kolei (zamiast trzech dróg do
  przeglądu), bez Sugerowania (tryb u góry); Plik bez Porównaj (jest w Recenzji) i bez Asystent AI
  (karta Asystent); panel Widok → Plik › Opcje bez dubli (powiększenie, wybór widoku, skróty
  sekcji); pigułka „niezapisane” schowana (licznik na Zapisz); obszar dokumentu dociąga do
  dolnego paska po zmianie wysokości wstążki (Czytanie).
- Zasada na dalej: jedna funkcja = jedno miejsce na wstążce (wyjątek jak w Wordzie: Komentarz
  na kartach Wstaw i Recenzja).

## Zasady techniczne

- Identyfikatory istniejących przycisków zostają (moduły i testy z nich korzystają); element
  zmienia tylko miejsce w DOM.
- Polecenia na kartach wołają te same funkcje co dziś (`composeUi.*`, `dwbShortcuts.openToolPanel`),
  bez kopii logiki.
- Rząd, który się nie mieści, przewija się w bok — nic nie chowa się do list ani pod „⋯”.
