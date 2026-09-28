# Documents Workbench PWA

Lokalny, offline-first warsztat do przeglądania, edycji i zapisywania plików `.docx` w przeglądarce — bez wysyłania plików na serwer.

Siostrzana aplikacja [Sheet Workbench PWA](../sheet-workbench-pwa/README.md).

## Funkcje (v0.2)

- Otwarcie `.docx` (drag & drop, picker, File System Access API)
- Podgląd dokumentu (`docx-preview`, leniwe ładowanie bibliotek)
- **Edycja inline (WYSIWYG)** — kliknij akapit w trybie edycji, zapis przez ZIP-patch
- Inspektor struktury: słowa, akapity, tabele, nawigacja po nagłówkach
- Wyszukiwanie z podświetleniem
- Narzędzia edycji: znajdź i zamień, wielkość liter, trim/spacje, prefiks/sufiks
- Zapis: w miejscu (FSA) lub „Zapisz jako” / pobranie
- PWA: service worker, tryb offline po pierwszym załadowaniu
- PL / EN, jasny / ciemny motyw
- **Rama ekranu jak w Sheet Workbench**: nazwa pliku i „Zapisz” z liczbą zmian w nagłówku
  (Ctrl/⌘+S), menu ⋯ (język, motyw, odświeżenie, inne aplikacje), panel obok dokumentu
  od 1024 px z „Znajdź ustawienie…”, pasek nad dokumentem (szukanie ↑↓, Czytanie/Edycja,
  B/I/U, zoom), skróty sekcji z nagłówków Worda (też polskich stylów; da się schować),
  zwijany nagłówek na telefonie, upuszczanie pliku w dowolnym miejscu okna
- **Cofnij / Ponów** (Ctrl/⌘+Z, Ctrl/⌘+Shift+Z, Ctrl+Y, przyciski ↶ ↷): operacje z panelu,
  formatowanie i pisanie w podglądzie; licznik przy „Zapisz” = kroki od ostatniego zapisu
- **Klawiatura i podpowiedzi**: Ctrl/⌘+F szukaj, Enter / Shift+Enter / F3 po trafieniach,
  Ctrl/⌘+Alt+E Czytanie ⇄ Edycja, Ctrl/⌘+Alt+1/2/3 panel / pasek / dokument (F6), Esc krok wstecz,
  „Przejdź do dokumentu” (Tab); podpowiedzi po najechaniu (cursor-hint z Sheet)
- **Tryb skupienia i pełny ekran** (Ctrl/⌘+Alt+F, przycisk na pasku, menu ⋯): sam dokument
  z paskiem — 85–87% ekranu zamiast 43–69%
- **Dotyk / iPhone** (sprawdzone na symulatorze iOS): pisanie z klawiaturą ekranową bez
  uciekania paska, brak przybliżania strony przy polach, „Zapisz” w Safari = zapis kopii,
  lżejsze efekty na dotyku; `?sample=nazwa` otwiera `docs/samples/nazwa.docx`

## Start lokalnie

```bash
npm install
node scripts/gen-sample-docx.mjs
npm run serve
```

Otwórz `http://127.0.0.1:7823/`.

## Build produkcyjny

```bash
npm run build
# serwuj katalog dist/
```

## Testy

```bash
npm test              # wszystkie testy, 2 naraz (serwer testów startuje sam, port 7823)
npm run test:fast     # 5 naraz — szybciej, komputer mocniej pracuje
npm test -- pwa find  # tylko kroki, których komenda zawiera któreś słowo
npm run test:serial   # po kolei, jak dawniej
npm run test:pwa      # service worker: offline, zawieszona sieć, aktualizacja
npm run test:frame    # rama ekranu (Chromium + WebKit)
npm run test:touch    # lekcje z iPhone'a (dotyk)
npm run test:enter    # szybkie pisanie wokół Enter/Backspace: zapis = podgląd
npm run test:undo     # Cofnij / Ponów (Chromium + WebKit)
npm run test:keys     # klawiatura, tryb skupienia, podpowiedzi
npm run test:stress   # pełny przebieg na dużym .docx (wymaga npm run serve)
```

Runner nie zatrzymuje się na pierwszym błędzie; to, co padło, powtarza raz pojedynczo
(przejście za drugim razem = „niestabilny”, wypisany osobno).

## Wydanie nowej wersji

```bash
npm run release              # podbija wersję (YYYYMMDD-NN) w sw.js, core.js i ?v= w index.html
npm run release 20261001-02  # albo konkretna wersja
```

Bez tego telefon z zainstalowaną PWA nie dostanie zmian — pliki z `?v=` są brane
wprost z cache, a nowy `sw.js` pojawia się tylko przy nowej wersji. Po wydaniu apka
pokazuje przycisk „Aktualizuj”.

## Deploy (GitHub + Vercel)

```bash
# 1. Repozytorium GitHub (jednorazowo)
gh repo create documents-workbench-pwa --public --source=. --remote=origin --push

# 2. Vercel (build → dist/)
npx vercel --prod
# Build Command: npm run build
# Output Directory: dist
```

Konfiguracja jest już w `vercel.json`.

## Author

Mateusz Zmuda

## License

Source code: MIT — see [LICENSE](./LICENSE).

Branding, logo i zrzuty ekranu pozostają własnością autora (jak w Sheet Workbench).
