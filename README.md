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
  lżejsze efekty na dotyku, **przybliżanie dwoma palcami** (w Widoku mobilnym tekst skaluje się
  i dalej zawija, w desktopowym strony przybliżają się jak zdjęcie; „Dopasuj” wraca do 100% /
  szerokości strony; zakres: tekst 50–300%, strony 25–300%; zsunięcie palców poniżej 50% w
  Widoku mobilnym przechodzi na strony); `?sample=nazwa` otwiera `docs/samples/nazwa.docx`
- **Granice stron** (Widok desktopowy, opcja w panelu Widok, domyślnie wł.): przerywana linia
  i „str. N” tam, gdzie Word zaczyna nową stronę — wysokość strony i marginesy z pliku, zasady
  Worda („razem z następnym”, wiersze razem, wdowy/sieroty) i odstępy wierszy jak w Wordzie;
  sprawdzone z prawdziwym Wordem (`npm run test:pages`). Przybliżone przy czcionkach, których
  przeglądarka nie ma i dla których nie ma zamiennika (niżej)
- **Wklejanie z zachowaniem stylów** (`app/paste-rich.js`): z Notatek Apple, stron WWW, Google
  Docs, Worda (HTML) i z Markdownu (Obsidian, Bear…): akapity, nagłówki, listy punktowane i
  numerowane z poziomami, listy kontrolne jako pola wyboru Worda, cytaty, kod, tabele, linia;
  pogrubienie, kursywa, podkreślenie, przekreślenie, wyróżnienie, linki (tylko http/https/mailto/tel).
  Kroje, rozmiary i kolory ze źródła NIE — tekst przyjmuje styl dokumentu. Jedno zdanie wkleja się
  w miejscu kursora; więcej — jedną operacją (jeden krok Cofnij). Zwykły tekst jak dawniej.
  `npm run test:paste`
- **Pola w snippetach**: `{{termin:data}}` (kalendarz), `{{status:lista=A|B}}`, `{{uwagi:długi}}`,
  `{{ilość:liczba}}`, `{{zgoda:zaznacz}}` / `taknie` — okienko przy wstawianiu, do dokumentu idzie
  wartość; przedrostek `formularz-` wstawia prawdziwe pole formularza Worda. „＋ Pole” w panelu
  buduje zapis. `npm run test:snfields`
- **Przypisy dolne i końcowe** (`app/doc-notes.js`): numer jak w Wordzie (ciągły przez cały
  dokument, format z ustawień pliku — 1, i, a, *), dymek z treścią przypisu po najechaniu na
  odnośnik (na dotyku przytrzymanie), klik (Czytanie) / dwuklik (Edycja) → skok do przypisu i
  „↩ Wróć”, klik w numer przypisu wraca do tekstu. Akapity z odnośnikiem są edytowalne (odnośnik
  to nienaruszalna „wyspa”), tekst przypisu poprawia się na dole strony (Enter = nowy akapit
  przypisu, numeru nie da się skasować) — zapis do `footnotes.xml` / `endnotes.xml`, razem z
  Cofnij, szkicem i kartami. Przypis z tabelą/polem — tylko do odczytu. `npm run test:notes`
- **Czcionki jak w Wordzie bez Office** (iPhone, Android, Mac bez Office): brakujące Calibri,
  Cambria, Arial, Times New Roman, Courier New i Georgia zastępują darmowe kroje o IDENTYCZNYCH
  szerokościach liter (Carlito, Caladea, Arimo, Tinos, Cousine, Gelasio) — wiersze i strony
  łamią się tak jak w Wordzie, nazwy krojów w pliku zostają bez zmian. Oryginał na urządzeniu
  ma pierwszeństwo. Łacina z polskimi znakami jest zapisana do pracy offline (~460 KB, raz —
  nie przy każdej aktualizacji), greka/cyrylica pobierają się przy pierwszym użyciu.
  Pliki: `assets/fonts/doc` (generuje `scripts/gen-doc-fonts.py`), `npm run test:fonts`
- **Komputer jak Word**: Ctrl + kółko (co 10%) i szczypanie na gładziku przybliżają sam
  dokument w miejscu kursora; −/+, suwak i „Dopasuj” zostają w tym samym miejscu dokumentu
- **Widok mobilny / Widok desktopowy** (panel Widok albo stuknięcie w procent na pasku): Auto =
  mobilny na dotyku w pionie i w wąskim oknie (≤768 px), desktopowy poza tym; ręczny wybór
  pamiętany osobno dla dotyku w pionie, w poziomie i komputera. Zmiana widoku i obrót wracają
  do tego samego akapitu, niezapisane zmiany zostają (`npm run test:view`)

- **Eksport**: TXT, Markdown, HTML (z Twoimi edycjami: nagłówki, listy, tabele) oraz
  „Drukuj / zapisz jako PDF” z układem stron jak w podglądzie
- **Statystyki**: słowa, znaki (ze spacjami i bez), zdania, akapity, tabele, czas czytania,
  słów na zdanie, najczęstsze słowa, najdłuższe zdania ze skokiem do miejsca, statystyki zaznaczenia
- **Kopie JSON**: eksport / import snippetów (przeniesienie na inne urządzenie) i wartości pól
  `{{…}}` (ten sam szablon wypełniany wiele razy); **„Usuń dane osobowe”** w Metadanych
  (autor, ostatnio zmieniający, opis, temat — przed wysłaniem pliku dalej)

- **Recenzja**: śledzone zmiany (wstawienia, usunięcia, przeniesienia, formatowanie, akapity,
  wiersze — też w nagłówkach, stopkach i przypisach), komentarze z odpowiedziami i zakomentowanym
  tekstem, przypisy; skok do miejsca, filtr autora, ✓ / ✗ przy każdej zmianie, „Akceptuj / Odrzuć
  wszystkie” (lub tylko wybranego autora), „Usuń wszystkie komentarze”; komunikat i licznik po
  otwarciu pliku. Akapity z przypisem, obrazem, polem, linkiem lub śledzoną zmianą są w trybie
  Edycja tylko do odczytu (edycja tutaj by je zgubiła) — z wyjaśnieniem po najechaniu

- **Znajdź i zamień v2**: wielkość liter, całe słowa, wyrażenia regularne z `$1` w zamianie,
  „tylko nagłówki sekcji”, „Zamień wszystkie” także w nagłówkach/stopkach/przypisach, podświetlenie
  samego trafienia, lista trafień z kontekstem, historia fraz; znajduje i zamienia też słowa
  rozcięte między fragmenty tekstu (formatowanie w środku słowa)
- **Pinch-zoom z procentem na żywo** (jak Word) i przyciąganiem do 100%
- **Wklejanie** jako czysty tekst w formacie miejsca kursora (wiersze = łamania wiersza)

- **Przewodnik po aplikacji** (przycisk „Przykład” / „Wypróbuj”): dokument-samouczek, w którym każda
  sekcja pokazuje jedną funkcję z gotowym materiałem do wypróbowania (`node scripts/gen-guide-docx.mjs`)
- **Poruszanie się jak w Wordzie** w trybie Edycja: strzałki między akapitami, Delete na końcu akapitu
  dołącza następny, Ctrl+Home/End (⌘↑/⌘↓), klik obok tekstu stawia kursor; znacznik bieżącego akapitu
  stoi przed tekstem (nie zasłania pierwszych znaków)

## Start lokalnie

```bash
npm install
node scripts/gen-sample-docx.mjs
node scripts/gen-review-docx.mjs   # próbka recenzji (?sample=review-sample)
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
npm run test:pinch    # przybliżanie dwoma palcami (Chromium + WebKit)
npm run test:review   # recenzja: skan, Akceptuj/Odrzuć, komentarze, blokady akapitów, mapowanie akapitów
npm run test:guide    # przewodnik: każda wskazówka „Spróbuj” działa jak opisano
npm run test:caret    # poruszanie się kursorem między akapitami (Chromium + WebKit)
npm run test:flows    # funkcje klikane jak użytkownik (Chromium + WebKit)
npm run test:export   # eksport, statystyki, import/eksport JSON, usuwanie danych osobowych
npm run test:keys     # klawiatura, tryb skupienia, podpowiedzi
npm run test:paste    # wklejanie z Notatek / Markdown / HTML (Chromium + WebKit)
npm run test:snfields # pola w snippetach: kalendarz, lista, pola formularza Worda (Chromium + WebKit)
npm run test:notes    # przypisy: numery, dymek, skok, edycja treści i przypisów, zapis (Chromium + WebKit)
npm run test:fonts    # zamienniki krojów Office: szerokości jak w Wordzie, offline (Chromium + WebKit)
npm run bench         # pomiar wydajności na dużych dokumentach (~100 i ~300 stron, CPU ×4)
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

Czcionki zastępcze w `assets/fonts/doc` mają własne licencje (SIL OFL 1.1 / Apache 2.0) —
teksty i źródła w tym katalogu (`README.txt`).

Branding, logo i zrzuty ekranu pozostają własnością autora (jak w Sheet Workbench).
