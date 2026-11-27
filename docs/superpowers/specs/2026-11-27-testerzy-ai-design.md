# Testerzy AI — projekt

Data: 2026-11-27. Status: do akceptacji.

## Po co

Mateusz sam znajduje błędy, używając apki do prawdziwej pracy. Przykład: konwertuje PDF, potem
dodaje zdjęcia i własny tekst. Najwięcej problemów to **niewygodna obsługa w porównaniu z Wordem**
oraz **ukryte błędy konwersji**, które wychodzą dopiero przy edycji. Istniejące testy losowe
(`chaos`, `monkey`, `undo-fuzz`) łapią wywrotki i niespójność danych, ale nie widzą rzeczy, które
„źle wyglądają” albo „męczą”. Przykład: tekst wchodzący na ramkę komórki (v1.01.38) przeszedł
przez nie bez słowa. Mateusz nie ma czasu na ręczne testy. Chce zespołu agentów, którzy robią to
za niego: szeroko, po ludzku, z różnymi osobami i zadaniami, częstymi i rzadkimi.

## Cele

- Komenda `/testerzy` uruchamia kilka (domyślnie 6) testerów AI naraz. Każdy ma osobę (kto,
  na czym, jak sprawny) i prawdziwe zadanie do wykonania od początku do końca.
- Tester naprawdę używa apki w przeglądarce (dotyk na telefonie, mysz na komputerze), ogląda
  zrzuty ekranu i ocenia jak człowiek przyzwyczajony do Worda.
- Znaleziska trafiają do jednej tabeli, zebrane, bez dubli i z wagami. Każde ma kroki
  do powtórzenia i zrzut.
- Potem segregacja (ja): potwierdzam każde znalezisko, odrzucam przesadzone, łączę duble i
  pokazuję Mateuszowi krótką listę od najważniejszych.
- Naprawa po jego decyzji. Każdy naprawiony błąd dostaje test: w istniejących testach, w
  `scripts/chaos-cases` albo w nowym automatycznym sprawdzeniu „źle wygląda”. To część B
  podejścia C, robiona przy okazji każdej poprawki, nie osobny projekt na start.

## Poza zakresem (na teraz)

- Uruchamianie samo w nocy lub w chmurze. Na razie odpala Mateusz, ręcznie.
- Testowanie na prawdziwym iPhonie. Emulacja WebKit nie zastępuje telefonu; to zostaje jako
  ręczny krok audytu.
- Samodzielne naprawianie przez agentów bez decyzji Mateusza.

## Czego nie chcemy (porażka, nawet jeśli „działa”)

- Zalew znalezisk w stylu „przycisk mógłby być bardziej niebieski”. Lepiej 8 prawdziwych niż 80
  szumu. Tester musi pokazać zrzut i kroki, inaczej znaleziska nie ma.
- Duble tego, co już jest na liście „świadomie inaczej niż Word” (np. `sort-dialog-word-diff`),
  zgłaszane co przebieg na nowo.
- Prywatne nazwy plików użytkownika w repo albo w commitach.
- Przebieg, który kosztuje dużo, a kończy się „nic nie znaleziono”, bo testerzy utknęli
  na klikaniu.

## Podejścia (rozważone)

| | Podejście | Plusy | Minusy |
|---|---|---|---|
| A | Testerzy AI z oczami, osoby + zadania | łapie niewygodę i „źle wygląda”, to, czego brakuje | wolniej, tokeny, ocena wymaga przesiania |
| B | Skryptowe osoby w chaos + reguły „źle wygląda” | tanie, powtarzalne, co noc | tylko to, co da się opisać regułą |
| **C** | **A szuka, każdy naprawiony błąd zamieniany w test B** | **szerokie szukanie + nic nie wraca** | dwa elementy do utrzymania |

Wybrane: **C, zaczynając od A** (decyzja Mateusza, 2026-11-27).

## Jak to działa

### 1. Sterownik przeglądarki dla agentów: `scripts/tester/drive.mjs`

Każdy tester ma **własną** przeglądarkę. Playwright MCP to jedna wspólna przeglądarka, więc
6 testerów by się w niej pobiło. Mały sterownik wywoływany z Basha:

- `start <imię> --device=iphone|iphone-landscape|ipad|desktop [--engine=webkit|chromium]`
  uruchamia przeglądarkę w tle (proces trzymający stronę, sterowany przez lokalny port lub
  gniazdo) i otwiera apkę z serwera testów `:7823`, z pominiętym ekranem powitalnym jak w
  testach.
- `look <imię>`: zrzut ekranu (plik PNG do obejrzenia przez Read) + krótka lista widocznych
  elementów, w które da się stuknąć (tekst/etykieta, współrzędne).
- `tap <imię> "<tekst lub etykieta>" | --xy=x,y`, `type <imię> "<tekst>"`, `key <imię> Enter`,
  `drag`, `scroll`, `open-file <imię> <ścieżka>` (wybór pliku jak użytkownik), `save <imię>`
  (pobranie zapisanego pliku do katalogu testera).
- `errors <imię>`: błędy JS i konsoli od ostatniego wywołania.
- `stop <imię>`.

Na telefonie stuknięcia są prawdziwym dotykiem (jak `VIEW=mobile` w chaos). Najpierw sprawdzić,
czy istnieje gotowe narzędzie (np. `@playwright/cli` z sesjami nazwanymi). Jeśli działa z
WebKit i dotykiem, użyć go zamiast własnego sterownika. To decyzja budującego.

### 2. Agent „tester”: `.claude/agents/tester-apki.md`

Definicja agenta Claude Code. Dostaje w poleceniu osobę i zadanie. Narzędzia: Bash (sterownik),
Read (zrzuty), Write (raport). Zasady w definicji:

- Najpierw **wykonać zadanie jak ta osoba**, nie szukać błędów na siłę. Niewygoda, którą
  poczuł po drodze, to znalezisko.
- Porównywać z Wordem (komputer, iPad, telefon), bo do niego przyzwyczajeni są użytkownicy.
- Lista „na co patrzeć” z `docs/AUDYT-MOBILE.md`.
- Na koniec zadania: zapis → ponowne otwarcie zapisanego pliku i sprawdzenie, czy wszystko jest.
- Limit kroków (np. 60) i czasu. Utknięcie na 3 próbach tej samej czynności to też znalezisko
  („nie da się zrobić X”), potem iść dalej.
- Raport: `output/testerzy/<przebieg>/<imię>/raport.md`. Każde znalezisko: waga (Krytyczny / Ważny /
  Drobny / Pomysł, jak w AUDYT-MOBILE), urządzenie, kroki, co się dzieje, czego oczekuje
  (jak w Wordzie), zrzut. Bez zrzutu i kroków nie ma znaleziska.
- Model: **Sonnet** (szybszy i tańszy, wystarcza do klikania i oceny). Segregacja i naprawy
  zostają u mnie.

### 3. Osoby i zadania: `docs/testerzy/osoby.md`

Katalog około 15–20 osób z wagą (częste / rzadkie), urządzeniem i zadaniem. Zbudowany z
10 ścieżek `AUDYT-MOBILE.md` i z tego, jak Mateusz faktycznie używa apki. Przykłady:

- Organizator na iPhonie: PDF z planem prac → dopisuje zadanie i wiersz tabeli → dodaje zdjęcie
  → zapisuje → wysyła. (częste)
- Studentka na iPadzie z klawiaturą: sprawozdanie z laboratorium → nagłówki, lista, tabela
  wyników, przypis → wydruk do PDF. (częste)
- Nauczyciel na komputerze: skan po OCR → poprawia tekst → szuka i zamienia. (średnie)
- Recenzent na telefonie poziomo: komentarze i śledzenie zmian w cudzym dokumencie. (rzadkie)
- Niecierpliwy: przerywa w trakcie, obraca ekran, wraca, zamyka kartę i otwiera ponownie
  (czy szkic przetrwał). (rzadkie)
- Ktoś, kto wkleja wszystko z Excela, stron WWW i maili. (częste)

Pliki do zadań: `docs/samples/` (w repo) i pula PDF z `~/.dwb-pdf-samples.txt` (poza repo, nazw
nie wpisujemy do raportów w repo). Testerzy mogą też tworzyć własne dokumenty od zera.

### 4. Komenda `/testerzy`: `.claude/commands/testerzy.md` (lub skill projektu)

`/testerzy [liczba=6] [obszar, np. „pdf”, „telefon”, „tabele”]`:

1. Uruchamia serwer testów, jeśli nie działa.
2. Losuje osoby z katalogu według wag (zawsze ≥ 2 telefon, ≥ 1 iPad, ≥ 1 komputer) albo
   bierze te z podanego obszaru.
3. Odpala testerów równolegle (agent `tester-apki`, w tle).
4. Po powrocie: **segregacja**. Odtworzyć każde znalezisko sterownikiem, odrzucić
   nieodtwarzalne i przesadzone, połączyć duble, sprawdzić z listą „świadomie inaczej niż Word”.
   W razie wątpliwości „jak robi Word” użyć `tools/word-oracle`.
5. Wynik: zebrane znaleziska do tabeli „Znaleziska” w `docs/AUDYT-MOBILE.md` (bez prywatnych
   nazw plików) + krótka lista dla Mateusza od najważniejszych z propozycją, co naprawić.

### 5. Naprawa (po decyzji Mateusza)

Każde znalezisko: `systematic-debugging` → poprawka → test (część B) → `npm test` dla obszaru.
Skill dobierany do rodzaju błędu, tylko gdy coś wnosi:

| Rodzaj | Skill |
|---|---|
| telefon, klawiatura, notch, dotyk | `mobile-native`, `gestures` |
| wykończenie, stany przycisków | `emil-design-eng`, `web-design-guidelines` |
| skrajne dane (długie nazwy, wielkie tabele) | `break-ui` |
| dostępność | `accessibility-engineering` |
| wolne działanie | `performance-optimization`, `animation-performance` |
| ruch, animacje | `review-animations` / `animate` |
| przed wydaniem | `verification-before-completion`, `code-review` |

## Ustalone

- Podejście C, start od testerów AI.
- Odpala Mateusz ręcznie; agenci tylko szukają i raportują, naprawy po jego decyzji.

## Moje decyzje (do sprawdzenia przez Mateusza)

- Testerzy na Sonnecie, 6 naraz, około 60 kroków każdy.
- Znaleziska do istniejącej tabeli w `docs/AUDYT-MOBILE.md` (ten plik trafia do repo).
- Raporty surowe i zrzuty w `output/testerzy/` (poza repo, już w `.gitignore`).
- Pierwszy przebieg próbny: 2 testerów, żeby sprawdzić, czy sterownik i raporty działają,
  zanim pójdzie pełne 6.

## Do decyzji budującego

- Własny sterownik czy gotowe narzędzie (warunek: WebKit + dotyk + równoległe sesje).
- Dokładna liczba kroków i limit czasu, format listy elementów w `look`.
- Komenda czy skill projektu dla `/testerzy`.
