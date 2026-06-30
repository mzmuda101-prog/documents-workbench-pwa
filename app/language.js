// Language dictionaries and translation helpers.

let currentLang = localStorage.getItem(typeof LANG_KEY !== "undefined" ? LANG_KEY : "documents-workbench-lang") === "en" ? "en" : "pl";
let BASE_TITLE = document.title || "Documents Workbench";

const I18N = {
  pl: {
    locale: "pl-PL",
    title: "Documents Workbench",
    description: "Lokalne przeglądanie, edycja i analiza dokumentów Word",
    online: "Online",
    offline: "Offline",
    panelOpen: "Zamknij panel",
    panelClosed: "Panel",
    noFile: "Brak pliku",
    loadingGeneric: "Ładowanie…",
    loadingFile: "Wczytywanie dokumentu…",
    renderingDoc: "Renderowanie dokumentu…",
    savingFile: "Zapisywanie…",
    docLoaded: "Dokument wczytany",
    sampleLoaded: "Przykładowy dokument wczytany",
    unsupportedType: "Obsługiwane formaty: .docx (PDF wkrótce)",
    libsMissingStatus: "Brak bibliotek dokumentu",
    libsMissingToast: "Nie udało się wczytać bibliotek DOCX",
    noFileToSave: "Najpierw wczytaj dokument",
    saveDone: "Zapisano dokument",
    saveFailed: "Nie udało się zapisać pliku",
    saveAsPrompt: "Nazwa pliku",
    saveInPlaceWarn: "Nadpisać oryginalny plik?",
    savePermissionDenied: "Brak uprawnień do zapisu w tym miejscu",
    editApplied: "Zastosowano zmiany: {count}",
    editNothing: "Brak zmian do zastosowania",
    editErrNoFind: "Podaj szukany tekst",
    editErrBadRegex: "Niepoprawne wyrażenie regularne",
    editErrNoAffix: "Podaj prefiks lub sufiks",
    editToolNoChange: "Brak zmian do zastosowania",
    editOpLabel: "Operacja",
    editOpReplace: "Znajdź i zamień",
    editOpCase: "Zmiana wielkości",
    editOpTrim: "Przytnij / spacje",
    editOpAffix: "Prefiks / sufiks",
    editCaseLabel: "Wielkość liter",
    editCaseUpper: "WIELKIE",
    editCaseLower: "małe",
    editCaseTitle: "Jak Nazwa Własna",
    editTrimLabel: "Tryb",
    editTrimEnds: "Przytnij końce",
    editTrimCollapse: "Zwiń wielokrotne spacje",
    editTrimHard: "Zamień twarde spacje",
    editPrefixLabel: "Prefiks (z przodu)",
    editSuffixLabel: "Sufiks (z tyłu)",
    inlineEditHint: "Kliknij akapit w dokumencie, aby edytować tekst",
    searchNoMatches: "Brak trafień",
    searchMatches: "{count} trafień",
    readModeOn: "Tylko odczyt",
    readModeOff: "Tryb edycji",
    words: "Słowa",
    chars: "Znaki",
    headings: "Nagłówki",
    paragraphs: "Akapity",
    tables: "Tabele",
    editToolsPanel: "Narzędzia edycji",
    editFindLabel: "Szukaj",
    editReplaceLabel: "Zamień na",
    editRegexLabel: "Wyrażenie regularne",
    editScopeLabel: "Zakres",
    editScopeAll: "Cały dokument",
    editScopeBody: "Treść główna",
    applyEdit: "Zastosuj",
    searchPanel: "Wyszukiwanie",
    searchQueryLabel: "Fraza",
    searchScopeLabel: "Zakres",
    searchHighlight: "Podświetl",
    structurePanel: "Struktura",
    filePanel: "Plik",
    viewPanel: "Podgląd",
    zoomLabel: "Powiększenie",
    dropText: "Przeciągnij plik",
    dropOr: "lub kliknij tutaj",
    dropBtn: "Wybierz plik",
    openFile: "Otwórz plik",
    emptyTitle: "Wczytaj dokument .docx",
    emptySub: "Plik zostaje w przeglądarce — bez wysyłania na serwer.",
    loadSample: "Przykład",
    save: "Zapisz",
    saveAs: "Zapisz jako",
    updateApp: "Aktualizuj",
    pdfSoon: "Podgląd PDF — w przygotowaniu",
    introAria: "Intro Mateusza",
    groupData: "Dane",
    groupTools: "Analiza i edycja",
    searchScopeAll: "Wszystko",
    searchScopeHeadings: "Nagłówki",
  },
  en: {
    locale: "en-US",
    title: "Documents Workbench",
    description: "Local viewing, editing, and analysis of Word documents",
    online: "Online",
    offline: "Offline",
    panelOpen: "Close panel",
    panelClosed: "Panel",
    noFile: "No file",
    loadingGeneric: "Loading…",
    loadingFile: "Loading document…",
    renderingDoc: "Rendering document…",
    savingFile: "Saving…",
    docLoaded: "Document loaded",
    sampleLoaded: "Sample document loaded",
    unsupportedType: "Supported formats: .docx (PDF coming soon)",
    libsMissingStatus: "Document libraries missing",
    libsMissingToast: "Could not load DOCX libraries",
    noFileToSave: "Load a document first",
    saveDone: "Document saved",
    saveFailed: "Could not save file",
    saveAsPrompt: "File name",
    saveInPlaceWarn: "Overwrite the original file?",
    savePermissionDenied: "No permission to save in place",
    editApplied: "Changes applied: {count}",
    editNothing: "No changes to apply",
    editErrNoFind: "Enter search text",
    editErrBadRegex: "Invalid regular expression",
    editErrNoAffix: "Enter a prefix or suffix",
    editToolNoChange: "No changes to apply",
    editOpLabel: "Operation",
    editOpReplace: "Find & replace",
    editOpCase: "Change case",
    editOpTrim: "Trim / spaces",
    editOpAffix: "Prefix / suffix",
    editCaseLabel: "Letter case",
    editCaseUpper: "UPPERCASE",
    editCaseLower: "lowercase",
    editCaseTitle: "Title Case",
    editTrimLabel: "Mode",
    editTrimEnds: "Trim ends",
    editTrimCollapse: "Collapse spaces",
    editTrimHard: "Replace hard spaces",
    editPrefixLabel: "Prefix (prepend)",
    editSuffixLabel: "Suffix (append)",
    inlineEditHint: "Click a paragraph in the document to edit text",
    searchNoMatches: "No matches",
    searchMatches: "{count} matches",
    readModeOn: "Read-only",
    readModeOff: "Edit mode",
    words: "Words",
    chars: "Characters",
    headings: "Headings",
    paragraphs: "Paragraphs",
    tables: "Tables",
    editToolsPanel: "Edit tools",
    editFindLabel: "Find",
    editReplaceLabel: "Replace with",
    editRegexLabel: "Regular expression",
    editScopeLabel: "Scope",
    editScopeAll: "Whole document",
    editScopeBody: "Main body",
    applyEdit: "Apply",
    searchPanel: "Search",
    searchQueryLabel: "Query",
    searchScopeLabel: "Scope",
    searchHighlight: "Highlight",
    structurePanel: "Structure",
    filePanel: "File",
    viewPanel: "View",
    zoomLabel: "Zoom",
    dropText: "Drop a file",
    dropOr: "or click here",
    dropBtn: "Choose file",
    openFile: "Open file",
    emptyTitle: "Load a .docx document",
    emptySub: "Your file stays in the browser — nothing is uploaded to a server.",
    loadSample: "Sample",
    save: "Save",
    saveAs: "Save as",
    updateApp: "Update",
    pdfSoon: "PDF preview — coming soon",
    introAria: "Mateusz intro",
    groupData: "Data",
    groupTools: "Analysis & editing",
    searchScopeAll: "Everything",
    searchScopeHeadings: "Headings",
  },
};

function t(key, vars) {
  const copy = I18N[currentLang][key] ?? I18N.pl[key] ?? key;
  if (!vars) return copy;
  return copy.replace(/\{(\w+)\}/g, (_, k) => (vars[k] != null ? String(vars[k]) : ""));
}

function applyLanguage() {
  const copy = I18N[currentLang];
  BASE_TITLE = copy.title;
  document.documentElement.lang = currentLang;
  const desc = document.getElementById("pageDescription");
  if (desc) desc.setAttribute("content", copy.description);
  if (!hasUnsavedChanges) document.title = copy.title;
  if (panelToggle) panelToggle.textContent = isSidebarOpen() ? copy.panelOpen : copy.panelClosed;
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    const key = el.dataset.i18n;
    if (copy[key]) el.textContent = copy[key];
  });
  const introVid = document.getElementById("introVideo");
  if (introVid && copy.introAria) introVid.setAttribute("aria-label", copy.introAria);
  document.querySelectorAll(".lang-button").forEach((btn) => {
    const on = btn.dataset.lang === currentLang;
    btn.classList.toggle("is-active", on);
    btn.setAttribute("aria-pressed", on ? "true" : "false");
  });
  if (typeof syncLangSwitchPill === "function") syncLangSwitchPill();
}

function setLanguage(lang) {
  currentLang = lang === "en" ? "en" : "pl";
  localStorage.setItem(LANG_KEY, currentLang);
  applyLanguage();
}
