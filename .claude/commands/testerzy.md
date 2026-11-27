---
description: Zespół testerów AI — kilka osób naraz używa apki do prawdziwych zadań, potem segregacja znalezisk
argument-hint: "[liczba=6] [obszar, np. pdf | tabele | telefon]"
---

Uruchom przebieg testerów AI dla Documents Workbench. Argumenty: `$ARGUMENTS`. Pierwsza liczba
to liczba testerów (domyślnie 6), słowo to obszar (`--area`, np. pdf, tabele, wklejanie;
puste = wszystko). Projekt i zasady: `docs/superpowers/specs/2026-11-27-testerzy-ai-design.md`.
Pracuj w `/Users/mateuszuser/Programistyka/documents-workbench-pwa`.

## 1. Przygotowanie

- Serwer testów: `curl -s -o /dev/null -w "%{http_code}" http://127.0.0.1:7823/`; gdy nie 200,
  uruchom w tle `python3 scripts/test-server.py 7823`.
- Obciążenie: `memory_pressure | tail -1`. Przy wolnej pamięci < 25% zmniejsz liczbę testerów
  (najmniej 2) i powiedz o tym. Sterownik i tak nie wpuści więcej niż 6 przeglądarek naraz.
- Sprzątanie po poprzednim razie: `node scripts/tester/drive.mjs stop-all`.
- Przebieg: `R=$(date +%Y-%m-%d-%H%M)`, katalog `output/testerzy/$R/`.

## 2. Skład

- `node scripts/tester/pick.mjs --n=<liczba> [--area=<obszar>]` daje osoby (JSON).
- Imię testera: krótkie, ascii, unikalne w przebiegu (np. `t1-organizator`).
- Osoby z „PDF z puli”: wylosuj plik z `~/.dwb-pdf-samples.txt` (pasujący do zadania, jeśli się
  da: skan dla OCR, tabela dla tabel) i skopiuj go do `output/testerzy/$R/pliki/<krótka-nazwa>.pdf`.
  Testerowi podaj tylko tę kopię. Prawdziwej nazwy pliku nie wpisuj nigdzie w repo.
- Pliki z `docs/samples/` podawaj ścieżką względną.

## 3. Testerzy (równolegle, w tle)

Dla każdej osoby jeden agent, wszystkie w jednej wiadomości, `run_in_background: true`.
Typ agenta `tester-apki`. Jeśli ten typ nie jest dostępny (nowa definicja, sesja sprzed niej),
użyj `general-purpose` z `model: "sonnet"` i zacznij polecenie od: „Najpierw przeczytaj w całości
`.claude/agents/tester-apki.md` i postępuj dokładnie według niego”. Polecenie dla testera:

```
imię: <imię>
przebieg: <R>
urządzenie: <device>
osoba: <id> — <kto>
zadanie: <zadanie>
plik: <ścieżka albo „nowy dokument”>
limit: 60 czynności
```

Czekaj na powiadomienia o zakończeniu (nie odpytuj w pętli). Po ostatnim:
`node scripts/tester/drive.mjs stop-all`.

## 4. Segregacja (Ty, nie testerzy)

Przeczytaj wszystkie `output/testerzy/$R/*/raport.md`. Dla każdego znaleziska:

1. **Krytyczne i Ważne odtwórz** własną sesją sterownika (to samo urządzenie, te same kroki,
   imię `spr-<n>`), oglądając zrzuty. Nie da się odtworzyć: „nie odtworzone” i nie wpisuj.
   Drobne: odtwórz, jeśli tanie; inaczej oceń po zrzucie.
2. **Odrzuć**:
   - pozycje z `docs/testerzy/swiadomie-inaczej.md`;
   - ograniczenia samej emulacji (brak prawdziwej klawiatury iOS, schowka systemu, gestów
     wieloma palcami). Te zapisz osobno jako „do sprawdzenia na prawdziwym iPhonie”;
   - wady pliku źródłowego (np. PDF tak wygląda) i gust bez konkretnej przeszkody.
3. **Połącz duble** (ten sam problem u kilku testerów to jedno znalezisko, wyżej w kolejce).
4. Spór „jak robi Word”: sprawdź w prawdziwym Wordzie (`tools/word-oracle`, AGENTS.md), nie zgaduj.

Dopisz potwierdzone do tabeli „Znaleziska” w `docs/AUDYT-MOBILE.md`, kolejne numery, Źródło
`testerzy <R>`, Stan `nowe`. Zrzut jako ścieżka w `output/` (zostaje lokalnie). Bez nazw plików
z puli.

## 5. Raport dla Mateusza (w czacie, po polsku, krótko)

- Ilu testerów, ile ukończyło zadanie, czas przebiegu, zużycie (suma `subagent_tokens`).
- Liczby: zgłoszone → potwierdzone → odrzucone (z jednym zdaniem dlaczego).
- Potwierdzone od najważniejszych, jedna linia każde, z numerem z tabeli.
- Propozycja, co naprawić najpierw. **Nic nie naprawiaj bez jego decyzji.**

## 6. Naprawa (po decyzji)

Każde znalezisko osobno: `superpowers:systematic-debugging` → poprawka → test, który bez
poprawki pada (istniejący `scripts/*-playwright.js`, `scripts/chaos-cases`, albo nowe sprawdzenie)
→ `npm test -- <obszar>`. Skill dobierz do rodzaju, tylko gdy coś wnosi: telefon i dotyk
`mobile-native`, `gestures`; wykończenie `emil-design-eng`, `web-design-guidelines`; skrajne dane
`break-ui`; dostępność `accessibility-engineering`; wolne działanie `performance-optimization`;
ruch `review-animations`. W tabeli Stan → `naprawione vX.YY.ZZ`. Wydanie jak w AGENTS.md.
