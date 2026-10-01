# Documents Workbench PWA — Roadmap

Stan na **v0.2** (lipiec 2026). Ostatnia wersja cache: `20260702-03`.

---

## Zrobione

| Obszar | Status |
|--------|--------|
| Scaffold PWA (shell, build, i18n, SW) | ✅ |
| Otwarcie `.docx` (drop, picker, FSA) | ✅ |
| Podgląd `docx-preview` + lazy libs | ✅ |
| Inspektor struktury + wyszukiwanie | ✅ |
| Edycja inline (contenteditable → ZIP-patch) | ✅ |
| Narzędzia edycji: replace, case, trim, affix | ✅ |
| Zapis w miejscu / Zapisz jako | ✅ |
| Deploy GitHub + Vercel | ✅ |
| Bateria testów Playwright (`npm test`, `test:stress`, `test:visual`) | ✅ |
| Poprawki wizualne list (kropki, numeracja, litery) | ✅ |
| Audyt wizualny bulletów / wcięć | ✅ |
| **Faza 2.5 — Mobile UX** (fit-width, bottom sheet, safe-area, kompaktowy hero) | ✅ |
| **Faza 3.1 — Find/Replace workbench** (skan, nawigacja, podgląd XML, zamiana 1/wszystkie) | ✅ |
| **Faza 4 v1 — Korekta typografii** (skan offline, reguły 1–6, panel, apply 1/reguła/wszystkie) | ✅ |
| **Faza 3.2 — Placeholdery** `{{pole}}` (skan, formularz, podmiana w XML) | ✅ |
| **Faza 3.3 — Snippety / klauzule** (`!nazwa`, localStorage, rozwinięcie + współpraca z {{}}) | ✅ |
| **Faza 3.4 — Inspektor z akcjami** (outline, filtr, skok, szybka edycja) | ✅ |
| **Faza 3.5 — Metadane** (`docProps/core.xml`: tytuł, autor, słowa kluczowe) | ✅ |

| **Paczka A — fundament z Sheet Workbench** (SW: limit nawigacji 3 s, `?v=` z cache, ciężkie zasoby po aktywacji; przycisk „Aktualizuj” czeka na nowy SW; `npm run release`; równoległy `npm test`; strażnik `test:pwa`) | ✅ 2026-09-28 |

| **Paczka B — rama ekranu z Sheet Workbench** (`app/app-frame.js`: nagłówek z plikiem i „Zapisz” + licznik, menu ⋯, panel obok ≥1024 px, „Znajdź ustawienie…”, pasek nad dokumentem, Czytanie/Edycja, skróty sekcji ze stylów nagłówków, zwijany nagłówek na telefonie; strażnik `test:frame` Chromium + WebKit) | ✅ 2026-09-28 |

| **Paczka C — dotyk / iOS** (symulator iPhone 18 Pro: pola 16 px bez przybliżania, strona wraca po zamknięciu klawiatury, pisanie z klawiaturą — pasek zostaje, Enter chowa klawiaturę w szukaniu, „Zapisz” bez FSA = zapis kopii, pełne tła i bez blurów na dotyku, toasty u góry, hover tylko dla myszy; **fix utraty tekstu przy szybkim pisaniu po Enterze** — kolejka przebudowy pliku; strażnicy `test:touch`, `test:enter`) | ✅ 2026-09-28 |

| **Paczka D — Cofnij / Ponów** (`app/undo.js`: migawki bajtów + zmienionych akapitów, kroki z nazwą dla operacji z panelu, formatowanie osobno, pisanie grupowane jak w Wordzie; ↶ ↷ na pasku, Ctrl/⌘+Z / Shift+Z / Ctrl+Y, historyUndo z menu/iOS; licznik „Zapisz” = kroki od zapisu; strażnik `test:undo`) | ✅ 2026-09-28 |

| **Paczka E — klawiatura, podpowiedzi, miejsce na dokument** (`app/keyboard.js`, `app/view-mode.js`, cursor-hint z Sheet; skróty Ctrl/⌘+F, Enter/F3, Ctrl/⌘+Alt+1/2/3/E/F, F6, stopniowe Esc, skip-link, inert schowanego panelu; tryb skupienia + pełny ekran: dokument 85–87% ekranu; ciaśniejsze odstępy; strażnik `test:keys`) | ✅ 2026-09-28 |

| **Paczka F — wydajność dużych dokumentów** (`npm run bench`; ~300 stron, CPU ×4: cofnięcie pisania 5,1 s → 0,06 s, szukanie 322 → 135 ms, 1. klawisz 88 → 46 ms, Enter widoczny 243 → 88 ms, otwarcie 5,0 → 4,1 s; wspólna pamięć rozpakowanego pliku, śledzenie zmienionych akapitów (MutationObserver), szybka ścieżka Cofnij, biblioteki dociągane po starcie) | ✅ 2026-09-28 |

| **Paczka G — pinch-zoom** (`app/pinch-zoom.js`: dwa palce przybliżają/oddalają dokument także przy „Dopasuj” — tekst skaluje się i dalej zawija; transform w trakcie gestu + zatwierdzenie zoomu po puszczeniu; miejsce w TEKŚCIE pod palcami zostaje pod palcami; zakres 35–300%; sprawdzone gestem na symulatorze iPhone'a. GOTCHA: `zoom: 1 !important` w trybie Dopasuj po cichu zerował gest — strażnik `test:pinch` mierzy wielkość tekstu na ekranie, nie zmienną) | ✅ 2026-09-28 |

| **Paczka J — Znajdź i zamień v2 + przegląd funkcji** (jeden silnik szukania/liczenia/zamiany po tekście całego akapitu: wielkość liter, całe słowa, `$1`, tylko nagłówki, nagłówki/stopki/przypisy, podświetlenie trafienia CSS Highlight, historia; procent na żywo przy pinch. **Naprawione przy przeglądzie:** placeholder/snippet „Wstaw” gubił kursor (pamięć kursora), rozmiar czcionki z listy nie działał na zaznaczeniu, skany (placeholdery, snippety, korekta) nie widziały tekstu wpisanego w podglądzie, Korekta nadpisywała świeżo dopisany tekst i zdejmowała formatowanie akapitu, „wielka po kropce” psuła „sp. z o.o.”, zamykający cudzysłów był prosty, „sierota i” przestawiała słowa, Shift+Enter zapisywał spację zamiast łamania wiersza, wklejanie wstawiało surowy HTML, ↶ po wstawieniu z panelu cofało też wcześniejsze pisanie, „Otwórz”/„Przykładowy” bez pytania o niezapisane zmiany, brak tekstu dla PDF/.doc. Strażnicy `test:find`, `test:flows`) | ✅ 2026-09-29 |

| **Paczka I — Recenzja** (`app/docx-revisions.js` + `app/review-panel.js`: śledzone zmiany we wszystkich częściach pliku, komentarze z odpowiedziami, przypisy; Akceptuj/Odrzuć pojedynczo, wszystkie lub autora — z łączeniem akapitów i przywracaniem właściwości jak w Wordzie; usuwanie komentarzy; licznik + komunikat po otwarciu. **Naprawione przy okazji (starsze błędy):** podgląd liczył akapity nagłówka/przypisów i tylko PIERWSZĄ sekcję → edycje w plikach z nagłówkiem lub podziałem strony trafiały w cudze akapity; akapity z przypisem/obrazem/polem/linkiem/śledzoną zmianą są teraz tylko do odczytu w Edycji. Strażnik `test:review`) | ✅ 2026-09-28 |

| **Paczka H — eksport i statystyki** (`app/export-panel.js`: TXT/MD/HTML + druk/PDF; `app/stats-panel.js`: liczby, czas czytania, najczęstsze słowa, najdłuższe zdania ze skokiem, zaznaczenie; eksport/import JSON snippetów i wartości pól; „Usuń dane osobowe” w metadanych; strażnik `test:export`) | ✅ 2026-09-28 |

| **Paczka K — formularze Worda** (`app/docx-forms.js` + `app/forms-panel.js`: kontrolki zawartości — lista, lista lub tekst, data w formacie Worda, pole wyboru ☐/☒, tekst z tekstem zastępczym — oraz stare pola FORMTEXT/FORMCHECKBOX/FORMDROPDOWN; klik w pole w dokumencie = okienko z listą/kalendarzem/tekstem, ☐ przełącza się od razu; panel „Formularz” z listą pól; pola powiązane (`w:dataBinding`) zmieniają wszystkie kopie i źródło w customXml/docProps; blokady z Worda i ochrona „tylko formularze” widoczne; zapis zmienia TYLKO zawartość kontrolki. **Naprawione przy okazji:** edycja akapitu z kontrolką rozbijała ją (pusta kontrolka + tekst obok) — akapity z polem są tylko do odczytu, a zapis z podglądu pomija zablokowane akapity; Backspace sklejał akapit z poprzednim zablokowanym (gubił pole/link/przypis); Wielkość liter/Przytnij/Dopisz przepisywały akapity z polem, linkiem czy przypisem — teraz je omijają. Strażnik `test:forms` Chromium + WebKit, fixture `scripts/gen-forms-docx.mjs`) | ✅ 2026-10-01 |

| **Poprawka zapisu** (zgłoszenie z Windows: plik otwarty przeciągnięciem / „Przykład” → „Zapisz” otwierało okno OTWIERANIA plików). Teraz bez prawa do nadpisania „Zapisz” = okno ZAPISU z podpowiedzianą nazwą; przeciągnięty plik w Chrome/Edge dostaje uchwyt (zapis w miejscu); zgoda na zapis pytana przed budowaniem pliku; „Zapisać w oryginale?” raz na plik; na iPhonie komunikat „pobrany — Pliki → Pobrane”. Sprawdzone na symulatorze iPhone'a. Strażnik `test:save` Chromium + WebKit | ✅ 2026-10-01 |

| **„Otwórz za pomocą” / „Udostępnij”** (`app/launch-files.js`, manifest `file_handlers` + `launch_handler: navigate-new` + `share_target`, odbiór POST w `sw.js`): zainstalowana aplikacja w Chrome/Edge na Windows/macOS pojawia się w „Otwórz za pomocą” dla .docx (plik z uchwytem → „Zapisz” do oryginału, każdy plik we własnym oknie); Android — „Udostępnij → Documents Workbench” (kopia). iPhone/iPad/Safari — nie obsługują, bez zmian. Strażnik `test:launch` Chromium + WebKit | ✅ 2026-10-01 |

| **Linki w dokumencie** (`app/doc-links.js`): spis treści / zakładki → płynny skok w dokumencie + „↩ Wróć” (Alt+←), adres strony bez #…; adres WWW → nowa karta (dawniej ZASTĘPOWAŁ aplikację); javascript:/plik na dysku → zablokowane z komunikatem; odsyłacze Worda (REF/PAGEREF/NOTEREF \h, HYPERLINK) klikalne; podpowiedź dokąd prowadzi link. **Fix:** podświetlenie edytowanego akapitu / wyniku szukania przy wysuniętym 1. wierszu zaczynało się w środku słowa (`fixHangingBox`). Strażnik `test:links` Chromium + WebKit | ✅ 2026-10-01 |

| **Przewodnik + intro** (przewodnik.docx: klikalny spis treści, rozdział „Formularz Worda” z 4 polami, „Linki i odsyłacze” z odsyłaczem REF, zapis / „Otwórz za pomocą” / „Udostępnij”; strażnik `test:guide` klika nowe „Spróbuj:”. Intro: film 2 s (tempo z długości filmu), znika dopiero gdy film się skończy I aplikacja jest gotowa; blokada autoodtwarzania nie odsłania ładowania. Komunikat „to formularz Worda” nie dla wbudowanych przykładów) | ✅ 2026-10-01 |

| **Niezapisane zmiany nie giną** (zgłoszenie: okno przeciągnięte na inny monitor → dokument narysował się od nowa, wpisany tekst zniknął z ekranu, „Zapisz” świeciło, a zapis byłby BEZ niego). Przerysowanie przy zmianie układu (telefon ⇄ komputer, obrót iPada, „Dopasuj”) najpierw przenosi wpisane zmiany do pliku (`rerenderKeepingEdits`). „Aktualizuj” / „Odśwież aplikację” pytają przy niezapisanych zmianach (iOS nie pokazuje „Opuścić stronę?”); aktualizacja z innego okna nie przeładowuje okna z niezapisanymi zmianami. Strażnik `test:survival` (25 czynności, Chromium + WebKit) | ✅ 2026-10-01 |

| **Szkic i odzyskiwanie niezapisanej pracy** (`app/drafts.js`): iOS/Android zamykają aplikację w tle bez zdarzenia „zamykam się” → szkic w IndexedDB (tylko na urządzeniu): bajty pliku tylko gdy się zmieniły, lista zmian ~1,5 s po zmianie i NATYCHMIAST przy przejściu w tło. Przy następnym otwarciu karta „Niezapisana praca — Przywróć / Odrzuć”. Znika po zapisie, Odrzuć, Cofnij do stanu z pliku, świadomym zamknięciu i sam po 7 dniach; żyjące okna (BroadcastChannel) nie pokazują się jako zgubiona praca; uchwyt pliku (Chrome/Edge) wraca — chyba że oryginał zmienił się od szkicu (wtedy zapis = kopia). Sprawdzone na symulatorze iPhone'a (zabicie Safari). Strażnik `test:drafts` Chromium + WebKit | ✅ 2026-10-01 |

| **Tworzenie od zera — Etap 1** (`app/docx-compose.js` operacje na pliku, `app/compose-ui.js` interfejs): „Nowy dokument” z szablonu (pusty A4 / pismo z polami {{…}} / notatka z nagłówkami; ekran startowy, menu ⋯, panel Plik, Ctrl/⌘+Alt+N); w Edycji na pasku „＋ Wstaw” (podział strony, linia pozioma, dzisiejsza data, znaki specjalne), lista stylu akapitu (Normalny/Tytuł/Podtytuł/Nagłówek 1–3/Cytat — styl szukany po nazwie wbudowanej, brakujący dopisywany do styles.xml, także gdy pliku stylów nie ma) i wyrównanie (też dla kilku zaznaczonych akapitów); Enter na końcu nagłówka → zwykły tekst; Ctrl/⌘+Enter = podział strony (w:pageBreakBefore + łatka docx-preview w `vendor-libs.mjs`, bo biblioteka brała go tylko ze stylu). Pasek Edycji = jeden rząd przewijany w bok z wygaszeniem, okienka kotwiczone pod przyciskiem. Przy okazji: sekcja nie dubluje się przy Enterze w akapicie z w:sectPr, Enter za podziałem strony nie robi kolejnej strony, zamienniki czcionek Office (Calibri → Carlito/Arial, Cambria → Georgia) zamiast Times na Macu/iOS. Sprawdzone w prawdziwym Wordzie (bez naprawy, polskie nazwy stylów, 3 strony jak w podglądzie). Strażnik `test:compose` (39 sprawdzeń, Chromium + WebKit) | ✅ 2026-10-01 |

**Plan przenoszenia z Sheet Workbench (A–F) zakończony.** Otwarte tematy: cofnięcie kroku z Enterem nadal rysuje dokument od nowa (~4 s przy 300 stronach); pierwszy układ bardzo długiego dokumentu jest kosztem docx-preview (content-visibility odrzucone — psuje numerację list przez izolację stylów). Później: „Ostatnio otwierane” (tylko Chrome/Edge, uchwyty plików w IndexedDB — lokalnie).

**Poza pierwotnym planem (ale wartościowe):** … edycja inline ze stylami runów + rozmiar czcionki, zoom 0.5–2.

---

## Faza 2.5 — Mobile UX ✅ (zrobione)

Telefon / tablet to kluczowy scenariusz PWA. Word mobile daje tu dobry wzorzec — my nie kopiujemy reflow 1:1 (mamy `docx-preview` = układ strony), ale **dopasowanie do ekranu** i **czytelny panel** muszą działać tak samy dobrze.

### 2.5.1 Smart skalowanie dokumentu (jak Word mobile)

Word na telefonie ma dwa tryby ([Mobile view vs Print Layout](https://support.microsoft.com/en-us/office/use-word-views-to-read-or-edit-your-document-bb918d39-fa40-461c-b503-873eddbdda43)):

| Word mobile | Nasz odpowiednik (plan) |
|-------------|-------------------------|
| **Print Layout** — strona jak na papierze, pinch zoom | Domyślny podgląd `docx-preview`; strona skalowana do szerokości ekranu |
| **Mobile view** — tekst od krawędzi do krawędzi, bez poziomego scrolla | Opcjonalnie później: „Tryb czytania” (CSS reflow / węższa kolumna) |
| Auto dopasowanie przy obrocie / zoom | `resize` + `orientationchange` → przelicz skalę |
| Brak poziomego scrolla przy czytaniu | **Fit-to-width** jako domyślne na `max-width: 768px` |

**Implementacja v1 (Print Layout + fit width):**

- `app/mobile-doc-zoom.js` — `computeFitZoom(viewport, pageEl)` → ustawia `--doc-zoom` (obecny slider 0.7–1.4 to za mało na wąskich ekranach)
- Po `renderDocxPreview` + przy zmianie rozmiaru okna: **auto „Dopasuj do szerokości”** na mobile
- Przycisk / toggle w panelu Podgląd: `Dopasuj szerokość` | `100%` | ręczny zoom (zachować slider)
- `touch-action` + opcjonalny pinch-to-zoom w `doc-viewport` (jak Word — po pinch tekst/strona przelicza skalę)
- `docx-preview-host` (`max-width: 820px`) skalować przez `transform: scale()` na `.doc-canvas` — już jest `--doc-zoom`, rozszerzyć zakres np. `0.35–1.4` na mobile
- Test Playwright: viewport 390×844, brak `scrollWidth > clientWidth` na `.doc-viewport` po załadowaniu fixture

**Implementacja v2 (opcjonalnie):**

- Tryb „Mobile view” — uproszczony układ bez sztywnej szerokości strony A4 (większa ingerencja; rozważyć po v1)

### 2.5.2 Panel wysuwany na mobilkach (naprawa)

Znane problemy (stan obecny):

- Sidebar `top: auto` + minimalne reguły `@media 768px` — panel „pływa”, zachodzi na dokument
- Uchwyt `sidebar-handle` zasłania treść / mylące pozycjonowanie przy otwarciu
- Scrim vs z-index — trudności z zamknięciem / kliknięciem w dokument
- Brak `safe-area-inset` (notch, home indicator)
- Brak blokady scrolla `body` gdy panel otwarty

**Plan naprawy (wzorować się na sheet-workbench-pwa mobile):**

- **<768px:** panel jako **bottom sheet** lub pełnoekranowy drawer (nie lewy flyout)
- Pełna szerokość, `max-height: 85dvh`, zaokrąglony górny róg, chwytak „drag”
- Scrim pod panelem; tap poza = zamknij; `overscroll-behavior: contain`
- Uchwyt boczny ukryty na mobile — zostaje tylko przycisk **Panel** w hero
- `env(safe-area-inset-*)` na sidebarze i hero
- Po akcji (wczytaj plik, zastosuj edycję) — auto-zamknij panel na mobile
- Test Playwright mobile context: otwórz/zamknij panel, brak elementów poza viewportem, dokument klikalny gdy panel zamknięty

---

## Kolejność prac (ustalona)

| # | Faza | Moduł | Status |
|---|------|--------|--------|
| 1 | **4** | Korekta językowa i typografia (offline) | ✅ v1 |
| 2 | **3.2** | Placeholdery `{{pole}}` | ✅ v1 |
| 3 | **3.3** | Snippety / klauzule | ✅ v1 |
| 4 | **3.4** | Inspektor z akcjami | ✅ |
| 5 | **3.5** | Metadane (`docProps`) | ✅ |
| 6 | **3.6** | Eksport TXT / HTML / MD + druk PDF | ✅ |

---

## Faza 3: Moduły na edycji

**3.1 zrobione.** Kolejność po Fazie 4: **3.2 → 3.3 → 3.4** (patrz tabela powyżej).

1. ~~**Find/Replace workbench**~~ ✅ — podgląd trafień przed „zamień wszystkie”, licznik, przejście trafienie po trafieniu
2. ~~**Placeholdery** `{{pole}}`~~ ✅ — wykrywanie, formularz wypełniania, podmiana w XML
3. ~~**Snippety / klauzule**~~ ✅ — `!nazwa`, localStorage, rozwinięcie; snippet może zawierać `{{placeholdery}}`
4. ~~**Inspektor z akcjami**~~ ✅ — outline w kolejności dokumentu, filtr, skok + podświetlenie, kopiuj, edycja w podglądzie, szybka edycja akapitu
5. ~~**Metadane**~~ ✅ — `docProps/core.xml` (tytuł, autor, słowa kluczowe); tworzy brakujące części pakietu
6. ~~**Eksport**~~ ✅ — TXT / HTML / Markdown + druk / PDF (bez backendu)

---

## Faza 4: Korekta językowa i typografia — ✅ v1

Lokalny, **offline-first** moduł — bez wysyłania tekstu na serwer.

### Zrobione (v1)

- `app/grammar-style.js` — reguły offline + `scanDocument` / `scanParagraph`
- `app/grammar-panel.js` — panel „Korekta”, skan, grupy `<details>`, apply 1 / reguła / wszystkie
- Reguły: podwójne spacje, spacja przed interpunkcją, wielokropek, cudzysłowy PL, wielka po kropce, sierota „i”, NBSP (opcjonalnie)
- Zapis przez `paragraphBatch` + auto-rescan
- Test: `scripts/grammar-playwright.js` (`npm run test:grammar`)

### v2 (później)

- Integracja z lokalnym słownikiem (np. Hunspell w WASM) — **ortografia**
- Heurystyki gramatyczne (np. zgodność przyimków) — ostrożnie, bez halucynacji
- Podświetlenie w podglądzie jak przy wyszukiwaniu (`search-hit`) — per-trafienie w DOM

### Implementacja v1 (zrealizowane)

- `app/grammar-style.js` — reguły jako `{ id, scan, fixAll }`
- `app/grammar-panel.js` — skan → panel sugestii → `applyDocumentEdit` / `paragraphBatch`
- PL domyślnie; EN — uproszczony zestaw reguł (bez cudzysłowów PL / sieroty „i”)

---

## Faza 5: PDF (opcjonalnie)

- Podgląd `.pdf` (PDF.js), strony, zoom
- Bez OCR / edycji PDF w v1

---

## Faza 6: Jakość i produkt

- Więcej reguł w `auditDocxVisualIssues` (tabele, łamanie, fonty)
- Screenshot diff w CI (opcjonalnie)
- README / screenshots publiczne
- Ostrzeżenie przy bardzo dużych plikach (>50 MB)

---

## Rekomendowany następny krok

Edycja akapitów z linkami/przypisami bez gubienia ich (zamiast blokady), „Ostatnio otwierane”, tabele, porównanie dwóch wersji.

---

## Author

Mateusz Zmuda
