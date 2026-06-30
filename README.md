# Documents Workbench PWA

Lokalny, offline-first warsztat do przeglądania, edycji i zapisywania plików `.docx` w przeglądarce — bez wysyłania plików na serwer.

Siostrzana aplikacja [Sheet Workbench PWA](../sheet-workbench-pwa/README.md).

## Funkcje (v0.1)

- Otwarcie `.docx` (drag & drop, picker, File System Access API)
- Podgląd dokumentu (`docx-preview`, leniwe ładowanie bibliotek)
- Inspektor struktury: słowa, akapity, tabele, nawigacja po nagłówkach
- Wyszukiwanie z podświetleniem
- Narzędzia edycji: znajdź i zamień (ZIP-patch na `word/document.xml`)
- Zapis: w miejscu (FSA) lub „Zapisz jako” / pobranie
- PWA: service worker, tryb offline po pierwszym załadowaniu
- PL / EN, jasny / ciemny motyw

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
npm run serve   # w osobnym terminalu
npm test
```

## Author

Mateusz Zmuda

## License

Source code: MIT — see [LICENSE](./LICENSE).

Branding, logo i zrzuty ekranu pozostają własnością autora (jak w Sheet Workbench).
