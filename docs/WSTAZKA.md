# Wstążka — wdrożenie makiety v4

Wzór: `docs/makiety/wstazka-v4.html` (zatwierdzona 2026-11-16). Zasada: **układ, rozmieszczenie,
teksty i zachowania 1:1 z makietą**. Animacje i przejścia mogą się różnić. Każde konieczne
odstępstwo najpierw do decyzji Mateusza.

Krój pisma i kolory zostają z aplikacji (tokeny w `styles/app.css`): makieta pokazuje układ,
nie zmienia motywu.

## Etapy (każdy to osobne wydanie)

| # | Etap | Co wchodzi | Stan |
|---|------|------------|------|
| 1 | Rama: dwa rzędy u góry (1a v1.01.00, 1b v1.01.01) | Wiersz 1: logo, Plik, karty, Szukaj, tryb (menu), Zapisz, ⋯. Wiersz 2: Cofnij/Ponów + polecenia karty w jednej linii (wstążka uproszczona). Karty: Narzędzia główne, Wstaw, Układ, Recenzja, Korespondencja, Widok, Asystent, kontekstowa Tabela. Przewijany rząd z podpowiedzią (wygaszenie, strzałka „+N”, jednorazowe drgnięcie). Komputer (≥ 769 px). | w toku |
| 2 | Plik za kulisami (v1.01.02) | Ekran Plik: Nowy, Otwórz, Zdjęcie, Informacje, Zapisz jako, Eksportuj, Drukuj, Porównaj, Zamknij, Opcje, Pomoc, Zgłoś błąd. Menu ⋯ odchudzone (aktualizacja, wersja, inne aplikacje). | — |
| 3 | Szyna i panele (v1.01.04) | Szyna 44 px (Nawigacja, Znajdź i zamień, Komentarze, Korekta, Asystent) z licznikami; panel wysuwany na margines albo przypięty. Zastępuje stały panel boczny. | — |
| 4 | Dolny pasek (v1.01.06) | Karty otwartych dokumentów (przewijane) + strona, wyrazy, błędy, język, powiększenie, „Sugerowanie”. | — |
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

## Etap 3 — szyna i wysuwany panel (+ karta Tabela)

`app/rail.js` + `#rbRail`. Szyna po lewej: Nawigacja (Struktura), Znajdź i zamień, Zmiany i
komentarze (Recenzja), Korekta, Asystent (Zapytaj dokument + Znajdź sens) — z licznikami z
nagłówków sekcji. Dawny panel boczny = wysuwany panel: jeden panel naraz, tytuł, pinezka, «; ten
sam ruch; bez pinezki leży nad dokumentem, z pinezką stoi obok (dawne `sidebar-docked`, wybór
zapamiętany). Na starcie bez panelu. Każda droga do panelu (`openToolPanel`, skróty, przyciski kart,
`details.open`) pokazuje go w tym panelu. Ctrl/⌘+Alt+1 = fokus na szynie. Znikają: 🔧, „WYSUŃ”,
„Znajdź ustawienie…”, sekcja Plik w panelu.

Bez dubli: z kart znikły przyciski otwierające panele z szyny (Recenzja: Korekta, Recenzja;
Widok: Struktura; Asystent: Zapytaj dokument, Znajdź sens; Narzędzia główne: Zamień → zostaje
Ctrl/⌘+H i szyna, na karcie „Narzędzia edycji”). W panelu Zmiany i komentarze schowane: Śledź
zmiany, Znaczniki, Akceptuj/Odrzuć wszystkie, Przejrzyj po kolei, Usuń komentarze (są na karcie).

Karta Tabela jak w makiecie: Wstaw ▾ (wiersz nad/pod, kolumna z lewej/z prawej — razem, uwaga
Mateusza), Usuń ▾, Scal i podziel ▾, Wiersz nagłówka, Obramowanie ▾, Kolor komórki ▾, Wyrównanie w
komórce ▾, Kierunek tekstu ▾, Rozłóż kolumny, Sortuj ▾ — części menu tabeli; pełne „Tabela ▾”
schowane.

## Etap 4 — dolny pasek (v1.01.06)

- Karty dokumentów przy wstążce zawsze (też jeden plik), z ikoną pliku i rozszerzeniem jak w
  makiecie; rząd przewija się ze strzałkami „+N” tym samym kodem co rzędy wstążki
  (`ribbon.js` makeStrip, `data-strip="docs"`). Telefon bez zmian: karty od drugiego pliku.
- Kolejność: karty … Sugerowanie · Str. X/Y · N wyrazów · błędy (Korekta) · język · powiększenie.
  „Str. X/Y” i „N wyrazy/wyrazów” jak w makiecie; „Śledzenie zmian” nazywa się „Sugerowanie”
  (jak tryb) — też w pigułce na telefonie.
- Błędy: liczba z Korekty (pojawia się po sprawdzeniu, wcześniej sama ikona); klik → Korekta.
  Język: język sprawdzania z Korekty (PL/EN); klik → Korekta z wyborem języka.
- Schowane przy wstążce (duble): „Dokument wczytany” (gdy plik jest otwarty), krój/rozmiar przy
  kursorze (jest we wstążce), „niezapisane: N zmian” (licznik na Zapisz — już wcześniej).
- Poniżej 1000 px chowają się wyrazy i język.
- Odstępstwa: × zamknięcia karty pokazuje się pod kursorem (w makiecie go nie ma, a karty trzeba
  zamykać; na dotyku stoi zawsze); przycisk „Dopasuj” zostaje przy powiększeniu (Widok nie ma
  grupy Powiększenie — decyzja z przeglądu); bez ikonki widoku mobilnego/desktopowego przy „100%”.
- Wysuwany panel kończy się nad dolnym paskiem (wcześniej zasłaniał karty dokumentów).

## Zasady techniczne

- Identyfikatory istniejących przycisków zostają (moduły i testy z nich korzystają); element
  zmienia tylko miejsce w DOM.
- Polecenia na kartach wołają te same funkcje co dziś (`composeUi.*`, `dwbShortcuts.openToolPanel`),
  bez kopii logiki.
- Rząd, który się nie mieści, przewija się w bok — nic nie chowa się do list ani pod „⋯”.

## Dalej — etapy 4–8 (stan po v1.01.04; do wzięcia w nowym czacie)

Zrobione i wydane: 1a, 1b, 2, 3 (+ karta Tabela, sprzątanie dubli). Zostało:

**4. Dolny pasek** — zrobione (v1.01.06, wyżej). (`.status-bar`, `app/open-docs.js`, `app/ribbon.js`)
- Karty otwartych dokumentów widoczne ZAWSZE przy wstążce (też jeden plik) — dziś `#docTabs`
  pokazuje się od drugiego pliku; przewijane z podpowiedzią jak rzędy wstążki (`.rb-strip`).
- Kolejność jak w makiecie: [karty dokumentów] … „Sugerowanie” (gdy tryb) · Str. 1/4 · wyrazy ·
  błędy (klik → szyna Korekta) · język · zoom.
- Bez dubli: „Dokument wczytany” (`#status`) — sprawdzić, czy potrzebne; licznik zmian jest na Zapisz.
- Wąsko (< 1000 px): chowają się wyrazy i język (jak w makiecie, `.wide`).

**5. Mini pasek przy zaznaczeniu** (nowy moduł, wzór: makieta v3/v4 `.mini`)
- Nad zaznaczeniem w Edycji: krój, rozmiar, B/I/U, wyróżnienie, kolor, punktory, styl,
  komentarz, Przepisz — wołają te same funkcje co wstążka (`composeUi`, `#fmt*` kliknięte).
- Gdy nad zaznaczeniem brak miejsca → pod nim; przy przewijaniu gaśnie; zaznaczenie poza
  ekranem → znika; nie wchodzi na wysuwany panel. Tylko mysz (na dotyku systemowe menu).
- Uwaga na dubel: to celowy skrót przy tekście (powód jak w Wordzie) — reszta jednego miejsca.

**6. Opcje wstążki** (ostatni przycisk rzędu poleceń, też Plik → Opcje)
- Układ: Uproszczona (dziś) / Klasyczna (duże przyciski + nazwy grup, wys. ~98 px — makieta v1/v4).
- Pokazywanie: Zawsze / Chowaj przy przewijaniu (rząd 2 zwija się przy przewijaniu w dół,
  wraca w górę) / Tylko karty (Ctrl/⌘+F1; rząd poleceń wysuwa się po kliknięciu karty).
- Gęstość Mysz / Dotyk (większe cele); zapamiętane w localStorage.

**7. Telefon** (≤ 768 px; dziś stary układ — `ribbon.js` oddaje kontrolki na miejsca)
- Góra 44 px: ← Pliki, Cofnij, nazwa + „N niezapisane · tryb”, przycisk trybu, ⋯ (Plik).
- Edycja: jeden przewijany rząd poleceń bieżącej karty nad klawiaturą; przycisk z nazwą karty
  otwiera arkusz w miejscu klawiatury (karty jako „chipy”, polecenia z podpisami, cele 44–48 px).
- Czytanie: cały ekran dokument, BEZ pływającego „Edytuj” (tryb zmienia się u góry).
- Zachować zachowania telefonu, które są dziś (zwijanie nagłówka, wyspa szukania, pinch).

**8. Znaczniki recenzji** (`app/review-margin.js`)
- Nad kolumną dymków (szeroko): licznik „N zmian · M komentarzy”, ↑ ↓ (poprzedni/następny),
  ✓ Akceptuj wszystkie. Reszta jak dziś (dymki obok kartki, wąsko ikonki + karta pod wierszem).
- Recenzja → Znaczniki ▾ już jest (Wszystkie / Proste / Bez).

### Jak pracować (sprawdzone w tym projekcie)
- Wzór wyglądu: `docs/makiety/wstazka-v4.html` (otwórz w przeglądarce, suwak szerokości).
- Zasady: 1:1 z makietą (animacje mogą się różnić); jedna funkcja w jednym miejscu — przenosząc,
  usuń ze starego (też z menu, okienek, Pliku, paska stanu); różne funkcje = różne ikony; samotna
  ikona → podpis; po zmianie układu sprawdź Czytanie i Edycję (wysokości, puste pasy).
- Kontrolki przenosi się gniazdami `[data-rb="#id"]` (te same id; telefon dostaje je z powrotem).
- Testy: `JOBS=6 node scripts/run-tests.mjs` (sam pomija nietknięte; pełny ~45 min). Testy klikające
  kontrolki z kart: `dwbRibbon.reveal("#id")`; panele: `dwbShortcuts.openToolPanel(id)`.
  Znany czerwony: `ENGINE=webkit schema` (crash WebKit, pada też na starym kodzie).
- Wydanie: commit `feat/fix(docs): …`, potem `npm run release`, commit `chore(docs): wydanie vX — …`,
  `git push`. Bez stopki Co-Authored-By.
