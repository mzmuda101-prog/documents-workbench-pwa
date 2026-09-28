// touch.js — dotyk i klawiatura ekranowa (paczka C, lekcje z Sheet Workbench).
//
// Wszystko za bramką (pointer: coarse): mysz i gładzik zachowują się jak dotąd.
//
// 1) Enter w polu szukania chowa klawiaturę — inaczej zasłania wynik, a do kolejnych
//    trafień są ↑ ↓ na pasku.
// 2) Kursor przy edycji dokumentu nie może zniknąć pod klawiaturą. Dokument przewija się
//    WEWNĄTRZ #docViewport (strona stoi), więc iOS nie zawsze sam go dosuwa.
//    GOTCHA (Sheet, 2026-09-19): getBoundingClientRect() jest liczone względem WIDOCZNEGO
//    obszaru (visual viewport) — widać 0 … visualViewport.height, a nie offsetTop … .
//    Część ruchów (animacja klawiatury) nie daje żadnego zdarzenia → sprawdzamy jeszcze
//    raz w kolejnych klatkach.

(() => {
  const coarse = matchMedia("(pointer: coarse)");
  const vv = window.visualViewport;

  // ── 1) Enter w szukaniu → schowaj klawiaturę ──────────────────────────────
  searchQueryEl?.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" || !coarse.matches || e.isComposing) return;
    // po tym, jak moduł Znajdź i zamień obsłuży Enter (ten sam event, później w kolejce)
    setTimeout(() => searchQueryEl.blur(), 0);
  });

  // ── 2) kursor nad klawiaturą ──────────────────────────────────────────────
  const MARGIN = 16;
  function caretRect() {
    const sel = window.getSelection?.();
    if (!sel || !sel.rangeCount) return null;
    const range = sel.getRangeAt(0);
    const node = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
    if (!node || !docCanvasEl?.contains(node) || !node.closest(".docx-editable-p")) return null;
    const rects = range.getClientRects();
    const r = rects.length ? rects[rects.length - 1] : range.getBoundingClientRect();
    if (r && (r.height || r.width)) return r;
    return node.closest(".docx-editable-p").getBoundingClientRect();
  }
  function keepCaretVisible() {
    if (!coarse.matches || !docViewportEl || readOnlyMode) return;
    const r = caretRect();
    if (!r) return;
    const visibleBottom = vv ? vv.height : window.innerHeight;
    const vpRect = docViewportEl.getBoundingClientRect();
    // dolna granica = niższa z: dół okna dokumentu, górna krawędź klawiatury
    const limit = Math.min(vpRect.bottom, visibleBottom) - MARGIN;
    if (r.bottom > limit) {
      docViewportEl.scrollTop += r.bottom - limit + 24;
    } else if (r.top < vpRect.top + MARGIN) {
      docViewportEl.scrollTop -= vpRect.top + MARGIN - r.top + 24;
    }
  }
  let queued = false;
  function queueCheck() {
    if (queued) return;
    queued = true;
    requestAnimationFrame(() => {
      queued = false;
      keepCaretVisible();
    });
  }
  // animacja klawiatury trwa ~250–400 ms i nie zawsze daje resize — dopilnuj kilka razy
  function checkSoon() {
    queueCheck();
    setTimeout(queueCheck, 120);
    setTimeout(queueCheck, 400);
  }
  vv?.addEventListener("resize", checkSoon);
  document.addEventListener("focusin", (e) => { if (e.target.closest?.(".docx-editable-p")) checkSoon(); });
  docCanvasEl?.addEventListener("input", queueCheck);
  document.addEventListener("selectionchange", () => { if (!readOnlyMode && document.activeElement?.closest?.(".docx-editable-p")) queueCheck(); });

  // ── 3) po zamknięciu klawiatury strona wraca na miejsce ────────────────────
  // iOS przy otwieraniu klawiatury przewija CAŁĄ stronę, żeby pokazać pole, i po jej
  // zamknięciu tak zostaje (odtworzone na symulatorze: nagłówek pod paskiem statusu,
  // pusty pas na dole). Z dokumentem układ mieści się w oknie — strona nie ma czego
  // przewijać, więc wracamy na górę.
  const EDITABLE = "input, textarea, select, [contenteditable='true']";
  function restorePageScroll() {
    if (!coarse.matches || !document.body.classList.contains("has-document")) return;
    if (document.activeElement?.matches?.(EDITABLE)) return; // klawiatura dalej potrzebna
    if ((window.scrollY || document.documentElement.scrollTop) > 0) window.scrollTo(0, 0);
  }
  document.addEventListener("focusout", () => setTimeout(restorePageScroll, 80));
  vv?.addEventListener("resize", () => {
    // klawiatura schowana = widoczny obszar znów prawie tak wysoki jak okno
    if (vv.height > window.innerHeight - 40) setTimeout(restorePageScroll, 80);
  });

  // ── 4) pisanie w dokumencie z otwartą klawiaturą ───────────────────────────
  // Odtworzone na symulatorze: iOS dosuwał kursor, przesuwając CAŁĄ stronę — nagłówek
  // i pasek (B/I/U, Zapisz) uciekały za ekran na cały czas pisania. Teraz: klasa
  // kb-open, obszar dokumentu = widoczna część ekranu (core.js syncDocViewportHeight),
  // strona wraca na górę, a kursor dosuwa przewijanie WEWNĄTRZ dokumentu. Nagłówek zwija
  // się do uchwytu (tap = rozwiń), skróty sekcji i pasek stanu chowają się na ten czas.
  let heroWasCollapsed = null;
  function editingInDoc() {
    return !!document.activeElement?.closest?.(".docx-editable-p");
  }
  function syncKeyboard() {
    if (!coarse.matches || !vv) return;
    const open = vv.height < window.innerHeight - 120 && editingInDoc();
    const was = document.body.classList.contains("kb-open");
    if (open === was) { if (open) checkSoon(); return; }
    document.body.classList.toggle("kb-open", open);
    if (open) {
      heroWasCollapsed = appFrame.isHeroCollapsed();
      appFrame.setHeroCollapsed(true);
    } else if (heroWasCollapsed === false) {
      appFrame.setHeroCollapsed(false);
    }
    syncDocViewportHeight();
    requestAnimationFrame(() => {
      if (open && (window.scrollY || document.documentElement.scrollTop) > 0) window.scrollTo(0, 0);
      checkSoon();
    });
  }
  vv?.addEventListener("resize", syncKeyboard);
  vv?.addEventListener("scroll", () => {
    // iOS potrafi przesunąć widok także bez zmiany rozmiaru — przy pisaniu cofnij to
    if (document.body.classList.contains("kb-open") && ((window.scrollY || document.documentElement.scrollTop) > 0)) window.scrollTo(0, 0);
  });
  document.addEventListener("focusin", () => setTimeout(syncKeyboard, 60));
  document.addEventListener("focusout", () => setTimeout(syncKeyboard, 120));

  window.dwbTouch = { keepCaretVisible, restorePageScroll, syncKeyboard };
})();
