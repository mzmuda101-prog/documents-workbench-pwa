// keyboard.js — obsługa z klawiatury + podpowiedzi (paczka E). Wzór: Sheet Workbench
// (tura klawiaturowa 2026-09-02) i zasady Mateusza dla klawiatur tabletowych:
//   - główne skróty na Ctrl/⌘ (+Alt) z cyfrą albo literą; F-klawisze tylko jako alias
//     (na klawiaturach typu Logitech rzędu F często nie ma albo siedzi pod Fn),
//   - NIGDY gołe Alt+litera (polskie AltGr: ą ę ś ł ó ż ź ć ń),
//   - przy Alt rozpoznajemy po e.code (macOS: Option+litera daje inny znak w e.key).
//
// Skróty (poza polami formularzy — tam zostaje zwykłe pisanie):
//   Ctrl/⌘+F            szukaj w dokumencie (pole na pasku)
//   Enter / Shift+Enter  w polu szukania: następne / poprzednie trafienie
//   F3 / Shift+F3, Ctrl/⌘+G / Ctrl/⌘+Shift+G  następne / poprzednie trafienie (wszędzie)
//   Ctrl/⌘+Alt+1 / 2 / 3 panel / pasek / dokument   (F6 / Shift+F6 — po kolei)
//   Ctrl/⌘+Alt+E        Czytanie ⇄ Edycja
//   Ctrl/⌘+Alt+F        tryb skupienia
//   Ctrl/⌘+Alt+N        nowy dokument (okno szablonów)
//   Ctrl/⌘+Enter        podział strony (w tekście, compose-ui.js)
//   Esc (stopniowo)     wyjście z pisania → panel-nakładka → podświetlenia szukania → tryb skupienia
//   (Ctrl/⌘+S — app-frame.js, Ctrl/⌘+Z / Shift+Z / Y — undo.js)

(() => {
  const isMac = /Mac|iPhone|iPad/.test(navigator.platform || "") || /Mac OS X/.test(navigator.userAgent);
  const sidebar = document.querySelector(".sidebar");
  const chipsEl = document.getElementById("sectionChips");

  function inFormField(el) {
    const tag = String(el?.tagName || "").toLowerCase();
    if (el?.closest?.(".docx-editable-p, .docx-edit-root")) return false;
    return !!el && (el.isContentEditable || tag === "textarea" || tag === "select"
      || (tag === "input" && !["checkbox", "radio", "button", "range", "submit"].includes(el.type)));
  }

  // ── obszary ekranu ─────────────────────────────────────────────────────────
  function focusPanel() {
    if (!isSidebarOpen()) toggleSidebar();
    setTimeout(() => document.getElementById("sidebarFinder")?.focus(), 60);
  }
  function focusToolbar() {
    const s = document.getElementById("searchQuery");
    if (s && s.offsetParent) { s.focus(); s.select(); return; }
    document.getElementById("appMenuBtn")?.focus();
  }
  function focusDocument() {
    if (!originalFileBytes) { document.getElementById("emptyOpenBtn")?.focus(); return; }
    if (!readOnlyMode) {
      // w Edycji: kursor w pierwszym widocznym akapicie
      const vp = docViewportEl.getBoundingClientRect();
      const p = [...docCanvasEl.querySelectorAll(".docx-editable-p")].find((el) => el.getBoundingClientRect().bottom > vp.top + 8);
      if (p) { placeCaret(p, 0); return; }
    }
    docViewportEl.focus({ preventScroll: true });
  }
  const REGIONS = [
    { el: () => document.querySelector(".hero"), focus: () => document.getElementById("appMenuBtn")?.focus() },
    { el: () => document.getElementById("docToolbar"), focus: focusToolbar },
    { el: () => docViewportEl, focus: focusDocument },
    { el: () => sidebar, focus: focusPanel, skip: () => !originalFileBytes && !isSidebarOpen() },
  ];
  function cycleRegion(dir) {
    const a = document.activeElement;
    let i = REGIONS.findIndex((r) => r.el()?.contains(a));
    for (let n = 0; n < REGIONS.length; n++) {
      i = (i + dir + REGIONS.length) % REGIONS.length;
      const r = REGIONS[i];
      const el = r.el();
      if (r.skip?.() || !el || (el !== sidebar && !el.offsetParent)) continue;
      r.focus();
      return;
    }
  }

  // ── szukanie ───────────────────────────────────────────────────────────────
  function stepSearch(dir) {
    const q = searchQueryEl?.value.trim();
    if (!q) { focusToolbar(); return; }
    ensureLazyFeature("find-replace").then(() => {
      if (typeof frMatches !== "undefined" && frMatches.length) stepFrMatch(dir);
      else runFindReplaceScan();
    }).catch(() => {});
  }
  // Enter w polu: po zmianie frazy — nowe szukanie; ta sama fraza — następne trafienie
  // (dawniej każdy Enter szukał od nowa i stał na 1. trafieniu). Shift+Enter — poprzednie.
  let lastScanQuery = "";
  searchQueryEl?.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" || e.isComposing) return;
    const q = searchQueryEl.value.trim();
    const loaded = typeof lazyFeatureLoaded !== "undefined" && lazyFeatureLoaded.has("find-replace");
    const sameQuery = q && q === lastScanQuery && loaded && typeof frMatches !== "undefined" && frMatches.length;
    if (sameQuery || e.shiftKey) {
      e.preventDefault();
      e.stopImmediatePropagation();
      stepSearch(e.shiftKey ? -1 : 1);
      return;
    }
    lastScanQuery = q; // pierwsze Enter — obsłuży moduł Znajdź i zamień (skan)
  }, true);
  searchQueryEl?.addEventListener("input", () => { lastScanQuery = ""; });

  function clearSearchHighlights() {
    const hl = docCanvasEl?.querySelectorAll(".search-hit");
    const precise = !!window.CSS?.highlights?.has?.("dwb-find"); // dokładne podświetlenie trafień (Znajdź i zamień)
    if (!hl?.length && !precise) return false;
    hl?.forEach((el) => el.classList.remove("search-hit", "search-hit-active"));
    if (precise) { CSS.highlights.delete("dwb-find"); CSS.highlights.delete("dwb-find-active"); }
    rootEl.classList.remove("find-precise");
    if (typeof frMatches !== "undefined") { frMatches = []; frActiveIndex = -1; }
    if (typeof appFrame !== "undefined") appFrame.syncPanelCounts();
    // fraza zostaje w polu, ale bez „0” — to nie brak wyników, tylko schowane podświetlenia
    const pos = document.getElementById("searchPos");
    if (pos) { pos.textContent = ""; pos.classList.remove("is-empty"); }
    lastScanQuery = "";
    return true;
  }

  // Escape is a genuine exit from a result: it must work even when focus is still in the
  // search field or inside an editable paragraph. CSS Highlights are not a native text
  // selection, so clear both kinds of visual selection here.
  function dismissDocumentSelection() {
    const clearedSearch = clearSearchHighlights();
    const selection = window.getSelection?.();
    const hasDocumentSelection = !!selection?.rangeCount && !selection.isCollapsed && !!docCanvasEl?.contains(selection.anchorNode);
    if (hasDocumentSelection) selection.removeAllRanges();
    return clearedSearch || hasDocumentSelection;
  }

  // ── główna obsługa klawiszy ────────────────────────────────────────────────
  document.addEventListener("keydown", (e) => {
    if (e.defaultPrevented || e.isComposing) return;
    const mod = e.ctrlKey || e.metaKey;
    const a = document.activeElement;

    // Ctrl/⌘+Alt+… (po e.code)
    if (mod && e.altKey && !e.shiftKey) {
      const map = {
        Digit1: focusPanel, Numpad1: focusPanel,
        Digit2: focusToolbar, Numpad2: focusToolbar,
        Digit3: focusDocument, Numpad3: focusDocument,
        KeyE: () => originalFileBytes && appFrame.setReadOnly(!readOnlyMode),
        KeyF: () => originalFileBytes && dwbView.toggle(),
        KeyN: () => typeof composeUi !== "undefined" && composeUi.openNewDialog(),
      };
      const fn = map[e.code];
      if (fn) { e.preventDefault(); fn(); }
      return;
    }
    if (e.key === "F6") { e.preventDefault(); cycleRegion(e.shiftKey ? -1 : 1); return; }
    if (!originalFileBytes) return;

    if (mod && !e.altKey && e.code === "KeyF") {
      e.preventDefault();
      focusToolbar();
      return;
    }
    if (e.key === "F3" || (mod && !e.altKey && e.code === "KeyG")) {
      e.preventDefault();
      stepSearch(e.shiftKey ? -1 : 1);
      return;
    }
    if (e.key === "Escape") {
      if (document.querySelector(".app-menu:not([hidden])")) return; // menu ⋯ ma własne Esc
      if (typeof clearSemanticHighlights === "function" && clearSemanticHighlights()) { e.preventDefault(); return; }
      if (a?.closest?.(".docx-editable-p, .docx-edit-root")) { e.preventDefault(); docViewportEl.focus({ preventScroll: true }); return; }
      if (dismissDocumentSelection()) { e.preventDefault(); return; }
      if (inFormField(a)) return; // pola: Esc należy do nich (np. czyszczenie szukajki ustawień)
      if (!rootEl.classList.contains("sidebar-docked") && isSidebarOpen()) { e.preventDefault(); setSidebarOpen(false); panelToggle?.focus(); return; }
      if (dwbView.isActive()) { e.preventDefault(); dwbView.set(false); }
    }
  });

  // ── skróty sekcji: strzałki (jeden przystanek Tab na cały rząd) ─────────────
  function chipList() { return chipsEl ? [...chipsEl.querySelectorAll(".section-chip")] : []; }
  function syncChipTabindex() {
    const chips = chipList();
    const cur = chips.find((c) => c.classList.contains("is-current")) || chips[0];
    chips.forEach((c) => { c.tabIndex = c === cur ? 0 : -1; });
  }
  chipsEl?.addEventListener("keydown", (e) => {
    const chips = chipList();
    const i = chips.indexOf(document.activeElement);
    if (i < 0) return;
    let next = -1;
    if (e.key === "ArrowRight" || e.key === "ArrowDown") next = Math.min(chips.length - 1, i + 1);
    else if (e.key === "ArrowLeft" || e.key === "ArrowUp") next = Math.max(0, i - 1);
    else if (e.key === "Home") next = 0;
    else if (e.key === "End") next = chips.length - 1;
    if (next < 0) return;
    e.preventDefault();
    chips.forEach((c, k) => { c.tabIndex = k === next ? 0 : -1; });
    chips[next].focus();
  });
  if (chipsEl && typeof MutationObserver === "function") {
    new MutationObserver(syncChipTabindex).observe(chipsEl, { childList: true, subtree: true, attributes: true, attributeFilter: ["class"] });
  }

  // ── panel schowany = poza kolejnością Tab (inert) ───────────────────────────
  // Zamknięty panel jest tylko odsunięty transformem — bez inert Tab wędrował po
  // kilkudziesięciu niewidocznych kontrolkach (ta sama lekcja co w Sheet).
  function syncInert() {
    if (!sidebar) return;
    const open = isSidebarOpen();
    sidebar.inert = !open;
    sidebar.setAttribute("aria-hidden", open ? "false" : "true");
  }
  const origSetOpen = window.setSidebarOpen;
  window.setSidebarOpen = function setSidebarOpenInert(...args) {
    const r = origSetOpen.apply(this, args);
    syncInert();
    return r;
  };

  // skip-link: Enter → dokument (a nie samo przewinięcie do kotwicy)
  document.getElementById("skipToDoc")?.addEventListener("click", (e) => { e.preventDefault(); focusDocument(); });

  // ── podpowiedzi (cursor-hint z Sheet) — dogrywane w wolnej chwili ───────────
  // Opisy z data-hint-pl / data-hint-en; na dotyku tylko elementy z data-hint-touch
  // (przytrzymanie). Ładowane leniwie — to 25 KB, które nie są potrzebne do startu.
  const loadHints = () => loadLazyScript("app/cursor-hint.js")
    .then(() => window.MateuszCursorHint?.initCursorHints({ fallbackHint: "" }))
    .catch(() => {});
  (window.requestIdleCallback ? (fn) => requestIdleCallback(fn, { timeout: 2500 }) : (fn) => setTimeout(fn, 1200))(loadHints);

  syncInert();
  syncChipTabindex();
  window.dwbKeyboard = { focusPanel, focusToolbar, focusDocument, cycleRegion, isMac };
})();
