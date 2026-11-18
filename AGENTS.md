# Documents Workbench — instrukcje dla agentów AI

PWA do pracy z dokumentami Worda (.docx) i PDF w przeglądarce, także offline i na telefonie/iPadzie.
Czysty JavaScript + HTML + CSS, bez frameworka. esbuild tylko skleja pliki w `dist/`.

## Zanim zaczniesz
- W tym folderze może równolegle pracować inna sesja lub agent. Najpierw `git status` i `git log -3`;
  nie nadpisuj cudzych niezacommitowanych zmian.
- Zachowanie wzorowane na Wordzie sprawdzaj w prawdziwym Wordzie (`npm run word:compare`,
  skrypty `word:oracle:*`, walidator `npm run ooxml:validate`), nie zgaduj.

## Mapa
- `index.html` — wejście; `app/*.js` — moduły apki (`core.js`, `document.js`, `page-breaks.js`, `undo.js`…).
- `styles/` — CSS; `sw.js` — service worker (offline); `lib/` — biblioteki obce (nie edytuj).
- `scripts/*-playwright.js` — testy; `scripts/run-tests.mjs` — runner.

## Uruchomienie
- `npm run dev` → http://127.0.0.1:3000 · serwer testów: port 7823 (WebKit nie wpuszcza na niektóre
  porty, np. 4190 — pusta strona = sprawdź konsolę).

## Testy
- `npm test` — sam pomija kroki, których zmiana nie dotyczy (`scripts/test-impact.mjs`).
  `npm test -- słowo` = tylko pasujące; `-- --dry` = podgląd; `-- --why` = powody.
- Mała zmiana → testy powiązane. `npm run test:all` (z `test:stress`) tylko przed dużym wydaniem.
  Nie puszczaj nic dłuższego niż kilka minut bez uprzedzenia. Chaos-testy równolegle (`SEEDS=1,7,13`), nie `RUNS=30`.
- Coś, co da się wstawić do dokumentu (tabela, linia, pole, obraz…), testuj od razu w pełnym cyklu:
  wstaw, usuń (Delete/Backspace/Wytnij), Enter obok, Cofnij/Ponów, kopiuj/wklej, zapis → otwarcie
  (także w Wordzie), telefon 390 px, podgląd wydruku. Chromium + WebKit.
- W testach `__dwbAutoDialog` wymaga `await`; pauz nie skracaj przed sprawdzeniem „nic się nie stało”.

## Wydanie
- Wersja `vMAJOR.MM.PP`: `npm run release` (+1), `-- skok` (do okrągłej …0/…5), `-- minor`, `-- major`.
- Przed podaniem tekstu commita sprawdź `git log -3` — jeśli ten numer już poszedł, podbij wersję.
- Commit: `typ(docs): krótki tytuł po polsku (vX.YY.ZZ)`, potem punkty „- …”. Bez nazw prywatnych plików użytkownika.

## Zasady UI (skrót; pełne: `../.claude/rules/`)
- Każde `:hover` w `@media (hover: hover)`; ramki fokusu przycisków przez `:focus-visible` / `.kbd-focus`.
- Litery nie jako skróty klawiszowe (klawiatury tabletowe); F1–F12 tylko jako alias.
- Wydruk: treść co najmniej 12,7 mm od krawędzi kartki.
- To, co może się nie zmieścić, przewija się — nic nie ucina się po cichu.
- Bez emoji w UI, kodzie i commitach.

## Sprawdzanie kodu
- `npm run check` — TypeScript sprawdza pliki z `// @ts-check` na górze (bez zmiany budowania).
  Nowe pliki zaczynaj od `// @ts-check`.
- `npm run lint` — Biome, tylko reguły na prawdziwe błędy.
