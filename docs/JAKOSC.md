# Jakość w pojedynkę — jak łapać więcej błędów i pisać mniej

Wspólne zasady dla Documents Workbench, Sheet Workbench i Smart Kalkulatora (2026-11-08).
Wnioski z sesji, w której nowe narzędzia w jeden wieczór znalazły 5 błędów, których nie złapało
224 istniejących testów. Część „Ta aplikacja” na końcu — co tu już jest i co dalej.

## Dlaczego testy przykładowe nie wystarczają

Test przykładowy sprawdza to, co ktoś WYMYŚLIŁ („w tym pliku na wydruku jest obrazek”). Błędy
siedzą w przypadkach, o których nikt nie pomyślał. Jedna osoba nie wymyśli wszystkich, więc
potrzebne są narzędzia, które same szukają przypadków. Trzy rodzaje błędów z tej sesji:

1. **Druga ścieżka nie dostała poprawek pierwszej.** Wydruk renderuje dokument osobno — pominął
   obrazy zakotwiczone do strony, numerację przypisów i znaczki pól wyboru, które Edycja miała
   od dawna. Każda funkcja z dwiema drogami (widok/wydruk, ekran/plik, mysz/klawiatura) to ryzyko.
2. **Brakujące pole w tabeli przypadków.** „Backspace w pustej linijce pod linią” — nikt nie
   zapisał, co ma się stać, więc stało się coś przypadkowego.
3. **Stan tymczasowy gubiony przy przejściu.** Wyłączone pogrubienie (bez zaznaczenia) ginęło
   po Enterze — format „dla dalszego pisania” nie przechodził do nowego akapitu.

## Zasady

**1. Błąd → najpierw test, który go odtwarza.** Zanim poprawka: test, który PADA. Po poprawce
przechodzi. Sprawdzić, że bez poprawki rzeczywiście pada (cofnąć ją na chwilę) — inaczej test
niczego nie pilnuje. Każdy zgłoszony błąd zostaje testem na zawsze.

**2. Niezmienniki zamiast przykładów.** Reguła ogólna sprawdzana na KAŻDYM wejściu, np.:
- to, co widać = to, co się zapisze (podgląd = plik),
- zapis → ponowne otwarcie niczego nie zmienia,
- druga ścieżka = pierwsza (wydruk = Edycja, eksport = widok),
- Cofnij N razy = stan początkowy,
- operacja i jej odwrotność = tożsamość (konwersja tam i z powrotem),
- zero błędów w konsoli.
Po napisaniu niezmiennika: sprawdzić, że ZŁAPAŁBY znany, naprawiony błąd (przywrócić na chwilę
starą wersję pliku). To jedyny dowód, że reguła ma zęby.

**3. Korpus prawdziwych danych.** Folder z prawdziwymi plikami/wejściami użytkownika (poza repo —
dane prywatne). Każdy nowy dziwny plik = wrzucić do folderu, od tej nocy jest testowany zawsze.
Na nim niezmienniki + losowe operacje z ziarnem (ziarno z daty → co noc inne ścieżki; raport
podaje polecenie do powtórzenia). Uruchamiane nocą, rano raport.

**4. Specyfikacja jako tabela danych.** Zachowania typu „klawisz × miejsce → wynik” zapisane jako
JEDNA tabela, z której generują się testy. Puste pole od razu widać. Pola „?” = decyzja do
podjęcia: raport pokazuje, co aplikacja robi dziś, a decyzję wpisuje się do tabeli.

**5. Wyrocznia zamiast pamięci.** „Jak robi Word/Excel?” — nie zgadywać: zapytać prawdziwy
program (na Macu AppleScript: otworzyć plik, nacisnąć klawisz, zapisać, porównać). Pamięć
i intuicja mylą się w szczegółach (np. styl Cytat w Wordzie ma „następny: Normalny”).

**6. Jedna reguła w jednym miejscu.** Ta sama logika w dwóch kopiach (ekran i plik, widok
i wydruk) prędzej czy później się rozjedzie. Wydzielić funkcję i wołać ją z obu miejsc
(np. numeracja przypisów: jedna funkcja dla Edycji i wydruku).

**7. Fałszywe alarmy zabijają zaufanie.** Narzędzie, które krzyczy bez powodu, przestaje być
czytane. Każdy alarm kończy się albo poprawką aplikacji, albo poprawką narzędzia (normalizacja
różnic zapisu: tabulator, łamanie wiersza, symbole). Ale ostrożnie — normalizacja nie może
osłabić reguły (pomijać tabulatory, nie wszystkie białe znaki).

**8. Tanie zgłaszanie.** Nagranie ekranu wystarcza (klatki wyciąga ffmpeg). Docelowo: skrót
w aplikacji „zgłoś błąd” zapisujący ostatnie akcje + stan → plik, z którego powstaje test.

## Kolejność wdrażania (największa dźwignia najpierw)

1. Niezmienniki + korpus + nocny przebieg (kilka godzin, potem działa samo).
2. Tabela zachowań dla obszaru, w którym było najwięcej błędów.
3. Wyrocznia dla pytań „jak robi oryginał”.
4. Zgłaszanie jednym klikiem.

## Ta aplikacja: Documents Workbench

**Jest (2026-11-08):**
- niezmienniki na korpusie — `scripts/corpus-invariants-playwright.js` (podgląd = plik, wydruk =
  Edycja, zapis → otwarcie, losowe edycje z ziarnem); `npm run test:corpus`; korpus `docs/samples`
  w `test:serial`; nocny przebieg `tools/nightly-corpus.sh` (+ `tools/com.dwb.nightly-corpus.plist`
  — włącza się dopiero po instalacji w `~/Library/LaunchAgents`)
- tabela klawiszy — `scripts/keys-table.js` (dane) + `scripts/keys-table-playwright.js`
- wyrocznia Word — `tools/word-oracle/` (AppleScript; `probe-keys.js` dla klawiszy)

**Dalej:**
1. Tabela klawiszy: tabela, podział strony, pole formularza, przypis, zaznaczenie przez akapity.
2. Niezmiennik „Cofnij wszystko = plik początkowy” w korpusie.
3. Korpus z `OCR=1` raz w tygodniu (skany PDF).
4. „Zgłoś błąd” jednym skrótem → plik akcji → przypadek w `scripts/chaos-cases`.
