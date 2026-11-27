# Testerzy AI — plan wykonania

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `/testerzy` uruchamia kilku agentów-testerów, z których każdy jako konkretna osoba wykonuje prawdziwe zadanie w apce we własnej przeglądarce, a znaleziska po segregacji trafiają do `docs/AUDYT-MOBILE.md`.

**Architecture:** Sterownik przeglądarki `scripts/tester/drive.mjs`: jeden proces w tle na testera trzyma stronę Playwrighta, a polecenia idą przez lokalny HTTP. Agent `tester-apki` (Sonnet) steruje nim z Basha i ogląda zrzuty. Katalog osób w Markdownie i losowanie z wagami w `scripts/tester/pick.mjs`. Komenda projektu `/testerzy` łączy wszystko i opisuje segregację.

**Tech Stack:** Node (ESM), Playwright 1.60 (już w devDependencies), definicje agentów i komend Claude Code (`.claude/agents`, `.claude/commands`).

**Spec:** `docs/superpowers/specs/2026-11-27-testerzy-ai-design.md`

## Global Constraints

- Nowe pliki JS zaczynają się od `// @ts-check` (AGENTS.md); `npm run check` i `npm run lint` czyste.
- Bez nowych zależności npm.
- Każdy tester ma własną przeglądarkę; serwer apki to serwer testów `http://127.0.0.1:7823/` (`python3 scripts/test-server.py 7823`).
- Surowe raporty, zrzuty i pobrane pliki tylko w `output/testerzy/<przebieg>/<imię>/` (poza repo).
- Prywatne nazwy plików użytkownika nie trafiają do repo ani commitów.
- Bez emoji w plikach, kodzie i commitach.
- Wagi znalezisk: Krytyczny / Ważny / Drobny / Pomysł (jak w `docs/AUDYT-MOBILE.md`).
- Model testerów: `sonnet`; domyślnie 6 testerów, limit 60 kroków każdy; skład co najmniej 2 telefon, 1 iPad, 1 komputer.
- Commity: `typ(docs): tytuł po polsku`, punkty „- …”, bez stopki Co-Authored-By.
- Obciążenie laptopa (M1 Pro, 32 GB): `drive.mjs start` odmawia, gdy żyje już 6 sesji (`MAX_TESTERS`, domyślnie 6); przeglądarka bez okna (headless); serwer testera kończy się sam po 30 min bez poleceń; `/testerzy` przed startem czyta `memory_pressure` i przy wolnej pamięci < 25% zmniejsza liczbę testerów (min. 2) i mówi o tym.

## Review Focus

- Tester padł albo przekroczył czas, a jego przeglądarka dalej działa w tle. `stop` i `stop-all` muszą zamykać procesy, a `start` na zajętym imieniu ma zgłosić błąd, a nie uruchamiać drugą przeglądarkę.
- Dwóch testerów naraz nie może dzielić portu ani katalogu. Test uruchamia dwóch jednocześnie.
- Pobrany (zapisany) plik musi dać się otworzyć z powrotem przez `open-file`. Test robi pełny cykl zapis → otwarcie.
- `look` na telefonie z otwartą klawiaturą lub oknem dialogowym ma pokazywać elementy w oknie, nie zasłonięte pod nim. Lista elementów pomija niewidoczne i zasłonięte (`elementFromPoint`).
- Serwer apki nie działa: `start` kończy się jasnym komunikatem, a nie pustą stroną.

---

### Task 1: Sterownik przeglądarki

**Files:**
- Create: `scripts/tester/drive.mjs`
- Create: `scripts/tester-drive-playwright.js` (test)
- Modify: `package.json` (skrypt `"test:tester": "node scripts/tester-drive-playwright.js"`)

**Interfaces:**
- Produces (CLI, wynik zawsze jako jeden JSON na stdout, kod wyjścia ≠ 0 przy błędzie z `{ "error": "…" }`):
  - `node scripts/tester/drive.mjs start <imię> --device=iphone|iphone-landscape|ipad|desktop [--run=<przebieg>]` zwraca `{ name, device, engine, dir, port }`. Uruchamia w tle (`spawn` z `detached`, `unref`) ten sam plik w trybie `serve`, czeka do 20 s na `GET /ping`. Stan zapisuje w `output/testerzy/<przebieg>/<imię>/state.json` oraz `output/testerzy/.sessions/<imię>.json` (port, pid, dir). Domyślny przebieg to `YYYY-MM-DD-HHMM`.
  - `look <imię>` zwraca `{ shot, url, dialog, downloads: [path], elements: [{ i, role, name, x, y, w, h }] }`. `shot` to ścieżka PNG (`shot-<n>.png`, n rośnie). `elements` to widoczne, niezasłonięte elementy do stuknięcia (button, a, input, select, textarea, [role=button|menuitem|tab|checkbox|option], [contenteditable]), najwyżej 120, `x`/`y` to środek.
  - `tap <imię> --el=<i>` (indeks z ostatniego `look`) | `--xy=<x>,<y>` | `"<tekst>"` (nazwa: najpierw równa, potem zawiera; niejednoznaczne → błąd z listą kandydatów). Na urządzeniach dotykowych `page.touchscreen.tap`, na komputerze `page.mouse.click`. Zwraca `{ ok: true }`.
  - `type <imię> "<tekst>"`, `key <imię> <Klawisz>` (składnia Playwrighta, np. `Shift+ArrowLeft`), `drag <imię> <x1>,<y1> <x2>,<y2>`, `scroll <imię> <dy> [--xy=x,y]`, `open-file <imię> <ścieżka>` (`setInputFiles('#fileInput')`), `wait <imię> <ms≤5000>`.
  - `errors <imię>` zwraca `{ errors: [string] }` (pageerror + console error od ostatniego wywołania).
  - `stop <imię>`, `stop-all`.
- Urządzenia: `iphone` = WebKit 390×844, DPR 3, `isMobile`, `hasTouch`; `iphone-landscape` = WebKit 844×390 jak wyżej; `ipad` = WebKit 820×1180, DPR 2, `isMobile`, `hasTouch`; `desktop` = Chromium 1440×900.
- Strona startowa jak w testach: `serviceWorkers: "block"`, init script `sessionStorage.setItem("introPlayed","true")`, po wczytaniu usunięcie `#heroSplash`. Pobrania (`page.on("download")`) zapisywane do `<dir>/pobrane/` pod sugerowaną nazwą.

- [ ] **Step 1: Write the failing test** `scripts/tester-drive-playwright.js`, z `check()` jak w innych testach `*-playwright.js`. Kolejno: `start` dwóch testerów naraz (`t-iphone --device=iphone`, `t-desk --device=desktop`, `--run=selftest`). Asercje: różne porty; `look` na każdym ma plik PNG istniejący i > 10 kB oraz `elements.length > 3`; `open-file t-desk docs/samples/przewodnik.docx`, po `wait 1500` `look` zawiera element `contenteditable` albo tekst dokumentu; `tap "<tekst>"` z nieistniejącym tekstem daje kod ≠ 0 i `error`; `start t-desk` drugi raz daje błąd „już działa”; `errors` zwraca tablicę; cykl zapisu: `key t-desk Meta+s` (na Linuksie `Control+s`) → w ciągu 5 s `look` ma `downloads.length ≥ 1` → `open-file` na tym pliku działa bez błędu w `errors`; `stop-all` → po 2 s procesy z `.sessions/*.json` nie żyją (`process.kill(pid, 0)` rzuca) i katalog `.sessions` jest pusty. Jeśli skrót zapisu w apce działa inaczej (okno z nazwą), test stuka przycisk zapisu przez `tap` i potwierdza okno; sprawdzić w `app/` przed pisaniem asercji.

- [ ] **Step 2: Run test to verify it fails**

Run: `node scripts/tester-drive-playwright.js` (serwer testów działa)
Expected: FAIL, `drive.mjs` nie istnieje.

- [ ] **Step 3: Implement `scripts/tester/drive.mjs`**: jeden plik, dwa tryby (CLI-klient i `serve`). Serwer to `http.createServer` na `127.0.0.1:0` (wolny port), trasa `POST /cmd` z `{ cmd, args }`. Wyłącza się sam po 30 min bez poleceń. `start` sprawdza najpierw `GET http://127.0.0.1:7823/`; gdy serwer nie odpowiada, zwraca błąd „uruchom: python3 scripts/test-server.py 7823”.

- [ ] **Step 4: Run test to verify it passes**

Run: `node scripts/tester-drive-playwright.js && npm run check && npm run lint`
Expected: wszystkie `check` zielone, check/lint czyste.

- [ ] **Step 5: Commit**

```bash
git add scripts/tester/drive.mjs scripts/tester-drive-playwright.js package.json
git commit -m "feat(docs): sterownik przeglądarki dla testerów AI"
```

### Task 2: Osoby i losowanie składu

**Files:**
- Create: `docs/testerzy/osoby.md`
- Create: `scripts/tester/pick.mjs`
- Create: `scripts/tester-pick-test.mjs` (test bez przeglądarki)

**Interfaces:**
- Consumes: nazwy urządzeń z Task 1.
- Produces: `node scripts/tester/pick.mjs [--n=6] [--area=<słowo>] [--seed=<liczba>]` zwraca na stdout JSON `[{ id, name, device, weight, areas: [string], task }]`. Format osoby w `osoby.md`: sekcja `## <id>` i pod nią linie `- urządzenie: iphone`, `- waga: częste|średnie|rzadkie`, `- obszary: pdf, tabele, …`, `- kto: …`, `- zadanie: …` (wielolinijkowe, do następnej sekcji). Wagi: częste 3, średnie 2, rzadkie 1.

- [ ] **Step 1: Write the failing test** `scripts/tester-pick-test.mjs` (asercje `node:assert`): dla `--n=6` i seedów 1..20: 6 różnych `id`, co najmniej 2 osoby z `device` iphone lub iphone-landscape, co najmniej 1 ipad, co najmniej 1 desktop; ten sam seed daje ten sam wynik; `--area=pdf` daje tylko osoby z `pdf` w `areas` (gdy jest ich mniej niż n, to wszystkie takie, bez wymogu składu); w 2000 losowań z `--n=1` osoba „częste” wypada częściej niż „rzadkie” (stosunek > 2); każda osoba w `osoby.md` ma wszystkie pięć pól (parser rzuca z nazwą sekcji przy braku).

- [ ] **Step 2: Run test to verify it fails**

Run: `node scripts/tester-pick-test.mjs`
Expected: FAIL, brak `pick.mjs`.

- [ ] **Step 3: Write `docs/testerzy/osoby.md`.** 16–20 osób pokrywających 10 ścieżek z `docs/AUDYT-MOBILE.md` („Ścieżki do przejścia”). Obowiązkowo z przykładami ze specu (organizator na iPhonie: PDF z planem → wiersz tabeli → zdjęcie → zapis; studentka na iPadzie: sprawozdanie, tabela, przypis, wydruk; nauczyciel na komputerze: skan po OCR, szukaj i zamień; recenzent poziomo: komentarze i śledzenie zmian; niecierpliwy: przerwania, obrót, zamknięcie karty; wklejający z Excela/WWW/maili). Zadanie pisane jako cel człowieka („przygotuj plan na sobotę i wyślij”), nie jako lista kliknięć. Pliki tylko z `docs/samples/` albo „PDF z puli” (tester dostaje ścieżkę od komendy, nazwy nie ma w tym pliku). Nagłówek pliku: jak dopisywać osoby.

- [ ] **Step 4: Implement `scripts/tester/pick.mjs`** z parserem `parseOsoby(md: string): Osoba[]` (eksport do testu) i losowaniem ważonym bez powtórzeń z ziarnem (mulberry32). Skład: najpierw wylosować brakujące urządzenia (2 telefon, 1 iPad, 1 komputer), resztę z całej puli.

- [ ] **Step 5: Run test**

Run: `node scripts/tester-pick-test.mjs && npm run check && npm run lint`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add docs/testerzy/osoby.md scripts/tester/pick.mjs scripts/tester-pick-test.mjs
git commit -m "feat(docs): katalog osób i losowanie składu testerów AI"
```

### Task 3: Agent `tester-apki`

**Files:**
- Create: `.claude/agents/tester-apki.md`

**Interfaces:**
- Consumes: CLI z Task 1; format osoby z Task 2.
- Produces: raport `output/testerzy/<przebieg>/<imię>/raport.md`. Na górze blok `wynik: ukończone|utknął|przerwane`, `kroki: N`. Potem znaleziska, każde jako sekcja `### <waga>: <krótki tytuł>` z polami `- urządzenie:`, `- kroki:` (lista numerowana), `- co się dzieje:`, `- czego oczekuję (jak w Wordzie):`, `- zrzut:` (ścieżka PNG z katalogu testera). Agent w ostatniej wiadomości podaje tylko ścieżkę raportu i liczbę znalezisk wg wagi.

- [ ] **Step 1: Write `.claude/agents/tester-apki.md`.** Frontmatter: `name: tester-apki`, `description` (kiedy używać: testowanie apki jako osoba z zadaniem), `model: sonnet`, `tools: Bash, Read, Write, Glob`. Treść po polsku, krótko:
  - wcielić się w osobę i wykonać zadanie jak ona, nie polować na błędy na siłę;
  - pętla: `look` → obejrzeć zrzut (Read) → jedna czynność → co kilka kroków `errors`;
  - lista „Na co patrzeć” z `docs/AUDYT-MOBILE.md`, plus porównanie z Wordem;
  - limit 60 czynności; 3 nieudane próby tej samej czynności to znalezisko „nie da się…”, potem dalej;
  - na koniec zapis → `open-file` pobranego pliku → sprawdzić, czy treść przetrwała;
  - znalezisko bez zrzutu i kroków nie wchodzi do raportu; bez rad typu „mogłoby być ładniej” bez konkretnej przeszkody; nie zgłaszać pozycji z `docs/testerzy/swiadomie-inaczej.md`;
  - zawsze `stop <imię>` na końcu, także po błędzie;
  - nie edytować plików apki, nie commitować.
- Utworzyć też `docs/testerzy/swiadomie-inaczej.md`: lista zachowań świadomie innych niż w Wordzie, zebrana z `docs/OD-ZERA.md`, `docs/JAKOSC.md` i pamięci projektu (np. „Sortuj…: kolumna z kursorem, brak pamięci ustawień”; „wiersz tabeli w Edycji nie dzieli się między strony”), jedna linia na pozycję.

- [ ] **Step 2: Próba ręczna jednego testera.** Wywołać agenta `tester-apki` z osobą „organizator na iPhonie” i krótkim limitem (20 czynności). Expected: istnieje `raport.md` w formacie z Interfaces, zrzuty istnieją, po zakończeniu nie ma żywego procesu z `.sessions`. Poprawić treść agenta, jeśli utknął na sterowaniu (np. nie trafia w elementy), a nie na apce.

- [ ] **Step 3: Commit**

```bash
git add .claude/agents/tester-apki.md docs/testerzy/swiadomie-inaczej.md
git commit -m "feat(docs): agent tester-apki"
```

### Task 4: Komenda `/testerzy` i segregacja

**Files:**
- Create: `.claude/commands/testerzy.md`
- Modify: `docs/AUDYT-MOBILE.md` (sekcja „Znaleziska”: dopisać kolumnę lub znacznik źródła `testerzy <przebieg>`; zatwierdzić plik, który dziś nie jest w repo)

**Interfaces:**
- Consumes: `pick.mjs` (Task 2), agent `tester-apki` (Task 3), `drive.mjs stop-all` (Task 1).
- Produces: `/testerzy [liczba] [obszar]`, a na końcu wiersze w tabeli „Znaleziska” i podsumowanie w czacie.

- [ ] **Step 1: Write `.claude/commands/testerzy.md`** (argumenty `$ARGUMENTS`: liczba domyślnie 6, opcjonalny obszar). Kroki dla mnie:
  1. Serwer testów: sprawdzić `:7823`, w razie potrzeby uruchomić w tle.
  2. `node scripts/tester/pick.mjs --n=<liczba> [--area=…]`; dla osób z „PDF z puli” wylosować plik z `~/.dwb-pdf-samples.txt` (ścieżkę podać testerowi tylko w poleceniu).
  3. Odpalić równolegle agentów `tester-apki` w tle, każdemu: imię, osobę, zadanie, urządzenie, przebieg, ścieżkę pliku.
  4. Po powrocie wszystkich: `drive.mjs stop-all`.
  5. Segregacja: przeczytać raporty. Każde znalezisko Krytyczne i Ważne odtworzyć sterownikiem (to samo urządzenie, te same kroki). Nieodtwarzalne oznaczyć „nie odtworzone” i nie wpisywać. Duble połączyć. Pozycje z `swiadomie-inaczej.md` pominąć. Przy sporze „jak robi Word” sprawdzić `tools/word-oracle`.
  6. Dopisać potwierdzone do tabeli „Znaleziska” w `docs/AUDYT-MOBILE.md` (bez nazw plików z puli; zrzut jako ścieżka w `output/`).
  7. W czacie: liczby wg wag, lista od najważniejszych (1 linia każde) z propozycją, co naprawić najpierw, oraz koszt/czas przebiegu. Nic nie naprawiać bez decyzji Mateusza.
  8. Przy naprawie: `systematic-debugging`, test regresji (część B), skill wg tabeli z sekcji 5 specu.

- [ ] **Step 2: Przebieg próbny `/testerzy 2`.** Expected: 2 raporty, segregacja wykonana, wiersze w tabeli albo jawne „nic potwierdzonego”, brak żywych procesów po zakończeniu. Zanotować czas i zużycie.

- [ ] **Step 3: Commit**

```bash
git add .claude/commands/testerzy.md docs/AUDYT-MOBILE.md
git commit -m "feat(docs): komenda /testerzy — zespół testerów AI z segregacją znalezisk"
```

### Task 5: Pierwszy pełny przebieg

- [ ] **Step 1:** `/testerzy` (6 osób). Expected: 6 raportów, segregacja, tabela uzupełniona, podsumowanie dla Mateusza z kosztem.
- [ ] **Step 2:** Jeśli testerzy masowo utykają na sterowaniu (nie na apce), poprawić `drive.mjs` lub treść agenta (osobny commit z testem w `tester-drive-playwright.js`) i powtórzyć tylko ten fragment.
