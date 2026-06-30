# Documents Workbench PWA — Roadmap

Stan na **v0.2** (czerwiec 2026). Ostatnia wersja cache: `20260630-12`.

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

**Poza pierwotnym planem (ale wartościowe):** `test-fixtures/`, `docx-render-fixes.js`, `visual-audit-playwright.js`, fix wyszukiwania (treść vs CSS).

---

## Faza 2.5 — Mobile UX (PRIORYTET #1)

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

## Faza 3: Moduły na edycji

Priorytet po **Fazie 2.5 (mobile)**:

1. **Find/Replace workbench** — podgląd trafień przed „zamień wszystkie”, licznik, przejście trafienie po trafieniu
2. **Placeholdery** `{{pole}}` — wykrywanie, formularz wypełniania, podmiana w XML
3. **Snippety / klauzule** — localStorage, wstawianie bloku tekstu
4. **Inspektor z akcjami** — skok + szybka edycja z panelu struktury
5. **Metadane** — `docProps/core.xml` (tytuł, autor, słowa kluczowe)
6. **Eksport** — pobranie TXT / HTML (bez backendu)

---

## Faza 4: Korekta językowa i typografia (NOWE)

Lokalny, **offline-first** moduł — bez wysyłania tekstu na serwer (zgodnie z filozofią warsztatu).

### Zakres v1 (reguły deterministyczne)

Panel **„Korekta”** w sidebarze: skan dokumentu → lista sugestii → zastosuj pojedynczo lub wszystkie.

| Reguła | Przykład problemu | Propozycja |
|--------|-------------------|------------|
| Sierota „i” | `… coś i.` na końcu zdania / linii | Przenieś „i” przed ostatni wyraz lub połącz zdania |
| Wielka litera po kropce | `coś. i następne` | `coś. I następne` (z kontekstem) |
| Spacja przed interpunkcją | `słowo ,` | `słowo,` |
| Podwójne spacje | `słowo  słowo` | pojedyncza spacja |
| Twarde spacje (PL) | `w Polsce`, `i tak` | NBSP po jednoliterowych przyimkach/spójnikach (opcjonalna reguła) |
| Cudzysłowy polskie | `"cytat"` | `„cytat”` |
| Wielokropek / myślnik | `...` | `…` (opcjonalnie) |

### Zakres v2 (później)

- Integracja z lokalnym słownikiem (np. Hunspell w WASM) — **ortografia**
- Heurystyki gramatyczne (np. zgodność przyimków) — ostrożnie, bez halucynacji
- Podświetlenie w podglądzie jak przy wyszukiwaniu (`search-hit`)
- Testy Playwright na fixture z celowo zepsutą typografią

### Implementacja (szkic)

- `app/grammar-style.js` — reguły jako `{ id, scan(text), suggest(text) }`
- Pipeline jak `edit-tools`: skan → panel sugestii → `applyDocumentEdit` / batch patch
- PL domyślnie; EN — uproszczony zestaw reguł

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

**Faza 2.5 — Mobile UX** (na pierwszym miejscu):

1. Fit-to-width + auto zoom na telefonie  
2. Naprawa panelu wysuwanego (bottom sheet, safe-area, scrim)

Potem **Faza 3.1 — Find/Replace workbench**, a równolegle lub zaraz po — **Faza 4.1 — Korekta typografii**.

---

## Author

Mateusz Zmuda
