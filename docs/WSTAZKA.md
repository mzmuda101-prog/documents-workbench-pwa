# Wstążka — wdrożenie makiety v4

Wzór: `docs/makiety/wstazka-v4.html` (zatwierdzona 2026-11-16). Zasada: **układ, rozmieszczenie,
teksty i zachowania 1:1 z makietą**. Animacje i przejścia mogą się różnić. Każde konieczne
odstępstwo najpierw do decyzji Mateusza.

Krój pisma i kolory zostają z aplikacji (tokeny w `styles/app.css`): makieta pokazuje układ,
nie zmienia motywu.

## Etapy (każdy to osobne wydanie)

| # | Etap | Co wchodzi | Stan |
|---|------|------------|------|
| 1 | Rama: dwa rzędy u góry | Wiersz 1: logo, Plik, karty, Szukaj, tryb (menu), Zapisz, ⋯. Wiersz 2: Cofnij/Ponów + polecenia karty w jednej linii (wstążka uproszczona). Karty: Narzędzia główne, Wstaw, Układ, Recenzja, Korespondencja, Widok, Asystent, kontekstowa Tabela. Przewijany rząd z podpowiedzią (wygaszenie, strzałka „+N”, jednorazowe drgnięcie). Komputer (≥ 769 px). | w toku |
| 2 | Plik za kulisami | Ekran Plik: Nowy, Otwórz, Zdjęcie, Informacje, Zapisz jako, Eksportuj, Drukuj, Porównaj, Zamknij, Opcje, Pomoc, Zgłoś błąd. Menu ⋯ odchudzone (aktualizacja, wersja, inne aplikacje). | — |
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

## Zasady techniczne

- Identyfikatory istniejących przycisków zostają (moduły i testy z nich korzystają); element
  zmienia tylko miejsce w DOM.
- Polecenia na kartach wołają te same funkcje co dziś (`composeUi.*`, `dwbShortcuts.openToolPanel`),
  bez kopii logiki.
- Rząd, który się nie mieści, przewija się w bok — nic nie chowa się do list ani pod „⋯”.
