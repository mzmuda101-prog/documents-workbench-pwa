// app-frame.js — „rama ekranu” wokół dokumentu (paczka B, makieta zaakceptowana 2026-09-28).
// Wzór: sheet-workbench-pwa/app/app-frame.js — ta sama rodzina zachowań.
//
// Co tu jest i dlaczego:
// - Nagłówek pokazuje, CO jest otwarte: nazwę pliku + jedną linię o nim (słowa, akapity,
//   czy zapis idzie w miejscu, czy jako nowy plik).
// - „Zapisz” na wierzchu: szary bez zmian, akcent + liczba zmian po edycji. Klik robi
//   dokładnie to, co „Zapisz” w panelu (albo „Zapisz jako”, gdy „Zapisz” jest niedostępne).
// - Menu ⋯: język, motyw, odświeżenie, zamknięcie dokumentu, inne aplikacje, stan PWA.
// - Panel OBOK dokumentu od 1024 px (bez nakładki i rozmytego tła), niżej — wysuwany.
// - „Znajdź ustawienie…” filtruje sekcje panelu.
// - Pasek nad dokumentem: szukanie, Czytanie/Edycja, B/I/U (tylko w Edycji), zoom.
// - Skróty sekcji z nagłówków dokumentu (da się schować — prośba Mateusza).
// - Telefon: nagłówek zwija się przy przewijaniu dokumentu.
// - Bezpiecznik przepełnienia: rzędy o zmiennej długości przewijają się w bok, a krawędź
//   miękko wygasa TYLKO wtedy, gdy faktycznie jest co przewinąć (zasada Mateusza).

const DOCK_KEY = "dwb-panel-docked-open-v1";
const SECTION_STRIP_KEY = "dwb-section-strip-v1";
const SECTION_CHIPS_MAX = 40;

// ── bezpiecznik przepełnienia ─────────────────────────────────────────────────
function attachOverflowFade(el) {
  if (!el || el._dwbFade) return;
  el._dwbFade = true;
  el.classList.add("overflow-fade");
  const sync = () => {
    const max = el.scrollWidth - el.clientWidth;
    el.classList.toggle("more-l", max > 1 && el.scrollLeft > 2);
    el.classList.toggle("more-r", max > 1 && el.scrollLeft < max - 2);
  };
  // Pionowe kółko myszy przewija rząd w bok (Shift+kółko działa natywnie).
  el.addEventListener("wheel", (e) => {
    if (Math.abs(e.deltaY) <= Math.abs(e.deltaX)) return;
    if (el.scrollWidth <= el.clientWidth + 1) return;
    el.scrollLeft += e.deltaY;
    e.preventDefault();
  }, { passive: false });
  el.addEventListener("scroll", sync, { passive: true });
  if (typeof ResizeObserver === "function") new ResizeObserver(sync).observe(el);
  if (typeof MutationObserver === "function") new MutationObserver(sync).observe(el, { childList: true, subtree: true, characterData: true });
  el._dwbFadeSync = sync;
  sync();
}

function formatChangesCount(n) {
  if (n === 1) return t("changesOne");
  const mod10 = n % 10;
  const mod100 = n % 100;
  const few = mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14);
  return t(few ? "changesFew" : "changesMany", { n });
}

const appFrame = (() => {
  const heroEl = document.querySelector(".hero");
  const heroTitleEl = document.getElementById("heroTitle");
  const heroMetaEl = document.getElementById("heroMeta");
  const heroSaveBtn = document.getElementById("heroSaveBtn");
  const heroSaveCount = document.getElementById("heroSaveCount");
  const heroGrip = document.getElementById("heroGrip");
  const menuBtn = document.getElementById("appMenuBtn");
  const menuEl = document.getElementById("appMenu");
  const buildEl = document.getElementById("appMenuBuild");
  const panelCountEl = document.getElementById("panelToggleCount");
  const sidebarNode = document.querySelector(".sidebar");
  const mainNode = document.querySelector(".main");
  const statusWordsEl = document.getElementById("statusWords");
  const statusDirtyEl = document.getElementById("statusDirty");
  const searchPosEl = document.getElementById("searchPos");
  const stripEl = document.getElementById("sectionStrip");
  const chipsEl = document.getElementById("sectionChips");
  const stripHideBtn = document.getElementById("sectionStripHide");
  const stripOpt = document.getElementById("showSectionChips");
  const narrowMq = matchMedia("(max-width: 768px)");
  const dockMq = matchMedia("(min-width: 1024px)");

  // ── nagłówek: plik ─────────────────────────────────────────────────────────
  function syncFile() {
    const hasFile = !!originalFileBytes;
    heroEl?.classList.toggle("has-file", hasFile);
    if (heroTitleEl) {
      heroTitleEl.textContent = hasFile ? currentFileName : "Documents Workbench";
      if (hasFile) heroTitleEl.setAttribute("title", currentFileName); else heroTitleEl.removeAttribute("title");
    }
    const s = documentStructure;
    const words = hasFile && s ? t("metaWords", { n: s.words.toLocaleString(I18N[currentLang].locale) }) : "";
    if (heroMetaEl) {
      const parts = [];
      if (hasFile && s) {
        parts.push(words, t("metaParas", { n: s.paragraphs.toLocaleString(I18N[currentLang].locale) }));
      }
      if (hasFile) parts.push(fileHandle ? t("metaInPlace") : t("metaSaveAs"));
      heroMetaEl.replaceChildren(...parts.map((txt, i) => {
        const span = document.createElement("span");
        span.textContent = txt;
        if (i) span.className = "hero-meta-sep";
        return span;
      }));
    }
    if (statusWordsEl) statusWordsEl.textContent = words;
    const closeDoc = document.getElementById("closeDocBtn");
    if (closeDoc) closeDoc.classList.toggle("hidden", !hasFile);
    syncSave();
  }

  // ── Zapisz + liczba zmian ──────────────────────────────────────────────────
  // Zmiana = jedna operacja (zamień, korekta, placeholdery, Enter/Backspace w podglądzie…)
  // albo jeden akapit poprawiony wpisywaniem. Liczymy tanio (bez porównywania całego
  // dokumentu przy każdym klawiszu): zbiór dotkniętych akapitów + licznik operacji.
  let opCount = 0;
  const touchedParas = new Set();
  function pendingCount() {
    // Z historią cofania (undo.js) liczymy DOKŁADNIE kroki od ostatniego zapisu — po
    // Cofnij licznik maleje, a powrót do zapisanego stanu gasi „Zapisz”.
    if (typeof dwbUndo !== "undefined") return dwbUndo.changesSinceSave();
    return opCount + touchedParas.size;
  }
  let saveSyncQueued = false;
  function syncSave() {
    if (saveSyncQueued) return;
    saveSyncQueued = true;
    requestAnimationFrame(() => {
      saveSyncQueued = false;
      if (!heroSaveBtn) return;
      const hasFile = !!originalFileBytes;
      heroSaveBtn.hidden = !hasFile;
      const dirty = hasFile && !!hasUnsavedChanges;
      const n = dirty ? Math.max(1, pendingCount()) : 0;
      heroSaveBtn.classList.toggle("is-dirty", dirty);
      if (heroSaveCount) {
        heroSaveCount.hidden = !dirty;
        heroSaveCount.textContent = n > 99 ? "99+" : String(n);
      }
      heroSaveBtn.setAttribute("aria-label", dirty ? t("heroSaveDirty", { changes: formatChangesCount(n) }) : t("heroSave"));
      if (statusDirtyEl) {
        statusDirtyEl.hidden = !dirty;
        statusDirtyEl.textContent = dirty ? `● ${t("statusUnsaved", { changes: formatChangesCount(n) })}` : "";
      }
    });
  }
  function clickSave() {
    if (saveBtn && !saveBtn.disabled) saveBtn.click();
    else if (saveAsBtn && !saveAsBtn.disabled) saveAsBtn.click();
  }
  heroSaveBtn?.addEventListener("click", clickSave);

  // Wpisywanie w podglądzie → dotknięty akapit (delegacja, bez ruszania docx-inline-edit.js).
  docCanvasEl?.addEventListener("input", (e) => {
    const p = e.target.closest?.(".docx-editable-p");
    if (p && p.dataset.paraIndex != null) { touchedParas.add(p.dataset.paraIndex); syncSave(); }
  }, true);

  // Jedno miejsce „stan zmian”: setDirtyState. false = zapisane / nowy plik → zeruj licznik.
  const origSetDirty = window.setDirtyState;
  window.setDirtyState = function setDirtyStateAndCount(isDirty, ...rest) {
    if (!isDirty) { opCount = 0; touchedParas.clear(); }
    const r = origSetDirty.call(this, isDirty, ...rest);
    syncSave();
    return r;
  };
  // Operacje na dokumencie przechodzą przez applyDocumentEdit albo (Enter/Backspace/Tab
  // w podglądzie) applyInlineStructuralEdit — tam liczymy +1 za udaną operację.
  ["applyDocumentEdit", "applyInlineStructuralEdit"].forEach((name) => {
    const orig = window[name];
    if (typeof orig !== "function") return;
    window[name] = async function countedEdit(...args) {
      const n = await orig.apply(this, args);
      if (n > 0) { opCount += 1; syncSave(); }
      return n;
    };
  });

  // Ctrl/⌘+S = Zapisz (jak w Sheet). Tylko gdy jest dokument — inaczej zostaw przeglądarce.
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;
    if (e.code !== "KeyS") return;
    if (!originalFileBytes) return;
    e.preventDefault();
    clickSave();
  });

  // ── menu ⋯ ─────────────────────────────────────────────────────────────────
  function menuItems() {
    return Array.from(menuEl.querySelectorAll("button, a[href]")).filter((el) => !el.disabled && el.offsetParent !== null);
  }
  // Menu mieszka w <body>, nie w nagłówku: nagłówek ma backdrop-filter (to tworzy własny
  // układ odniesienia dla position:fixed) — menu w środku byłoby przycięte.
  if (menuEl && menuEl.parentElement !== document.body) document.body.appendChild(menuEl);
  function placeMenu() {
    const r = menuBtn.getBoundingClientRect();
    const vw = document.documentElement.clientWidth;
    const width = Math.min(268, vw - 16);
    menuEl.style.width = `${width}px`;
    menuEl.style.top = `${Math.round(r.bottom + 8)}px`;
    menuEl.style.left = `${Math.round(Math.max(8, Math.min(vw - width - 8, r.right - width)))}px`;
    menuEl.style.maxHeight = `${Math.max(160, window.innerHeight - r.bottom - 20)}px`;
  }
  function setMenuOpen(open, opts = {}) {
    if (!menuEl || !menuBtn) return;
    menuEl.hidden = !open;
    menuBtn.setAttribute("aria-expanded", String(open));
    if (open) {
      placeMenu();
      // wskaźnik PL/EN liczy się z rozmiarów przycisków — w ukrytym menu było 0
      if (typeof syncLangSwitchPill === "function") syncLangSwitchPill();
      if (buildEl && typeof APP_BUILD_VERSION !== "undefined") buildEl.textContent = APP_BUILD_VERSION;
      if (opts.focusFirst) { const first = menuItems()[0]; if (first) first.focus(); }
    } else if (opts.returnFocus) {
      menuBtn.focus();
    }
  }
  if (menuBtn && menuEl) {
    menuBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      setMenuOpen(menuEl.hidden, { focusFirst: e.detail === 0 }); // klawiatura → fokus w menu
    });
    document.addEventListener("pointerdown", (e) => {
      if (menuEl.hidden || menuEl.contains(e.target) || menuBtn.contains(e.target)) return;
      setMenuOpen(false);
    });
    // Esc zamyka menu zawsze, gdy jest otwarte (faza capture — przed innymi Esc).
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape" || menuEl.hidden) return;
      e.preventDefault();
      e.stopPropagation();
      setMenuOpen(false, { returnFocus: menuEl.contains(document.activeElement) || document.activeElement === menuBtn });
    }, true);
    menuEl.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
      const items = menuItems();
      const i = items.indexOf(document.activeElement);
      const next = items[(i + (e.key === "ArrowDown" ? 1 : -1) + items.length) % items.length];
      if (next) { e.preventDefault(); next.focus(); }
    });
    window.addEventListener("resize", () => { if (!menuEl.hidden) placeMenu(); }, { passive: true });
    // Akcje w menu (odśwież, zamknij dokument, link) — menu niech się zamknie.
    menuEl.querySelectorAll(".app-menu-item").forEach((el) => el.addEventListener("click", () => setMenuOpen(false)));
  }

  // ── Czytanie / Edycja ──────────────────────────────────────────────────────
  const modeBtns = Array.from(document.querySelectorAll(".mode-switch .mode-btn"));
  function syncModeSwitch() {
    const ro = !!readModeEl?.checked;
    modeBtns.forEach((b) => {
      const on = (b.dataset.mode === "read") === ro;
      b.classList.toggle("is-active", on);
      b.setAttribute("aria-pressed", String(on));
    });
  }
  function setReadOnly(ro) {
    if (!readModeEl || readModeEl.checked === !!ro) { syncModeSwitch(); return; }
    readModeEl.checked = !!ro;
    readModeEl.dispatchEvent(new Event("change", { bubbles: true }));
  }
  modeBtns.forEach((b) => b.addEventListener("click", () => setReadOnly(b.dataset.mode === "read")));
  readModeEl?.addEventListener("change", syncModeSwitch);

  // ── zoom na pasku ──────────────────────────────────────────────────────────
  const zoomNowEl = document.getElementById("zoomNow");
  function syncZoomNow() {
    if (!zoomNowEl) return;
    const reflow = typeof shouldUseMobileReflow === "function" && shouldUseMobileReflow();
    zoomNowEl.textContent = reflow && !(typeof isReflowScaled === "function" && isReflowScaled()) ? t("zoomFitShort") : `${Math.round((parseFloat(zoomLevelEl?.value) || 1) * 100)}%`;
  }
  function stepZoom(dir) {
    if (!zoomLevelEl) return;
    const { min, max } = typeof getZoomLimits === "function" ? getZoomLimits() : { min: 0.5, max: 2 };
    const cur = parseFloat(zoomLevelEl.value) || 1;
    // kroki jak w Wordzie: co 10 %, z „przyciąganiem” do okrągłych wartości
    const next = Math.round((cur + dir * 0.1) * 10) / 10;
    zoomLevelEl.value = String(Math.max(min, Math.min(max, next)));
    zoomLevelEl.dispatchEvent(new Event("input", { bubbles: true }));
  }
  document.getElementById("zoomInBtn")?.addEventListener("click", () => stepZoom(1));
  document.getElementById("zoomOutBtn")?.addEventListener("click", () => stepZoom(-1));
  const origApplyZoom = window.applyZoom;
  window.applyZoom = function applyZoomAndShow(...args) {
    const r = origApplyZoom.apply(this, args);
    syncZoomNow();
    return r;
  };
  if (typeof window.setZoomMode === "function") {
    const origMode = window.setZoomMode;
    window.setZoomMode = function setZoomModeAndShow(...args) {
      const r = origMode.apply(this, args);
      syncZoomNow();
      return r;
    };
  }

  // ── szukanie na pasku (moduł Znajdź i zamień ładuje się leniwie) ────────────
  // Pole i ↑↓ są na wierzchu od razu, a find-replace-workbench.js dociąga się przy
  // pierwszym użyciu. Pierwszy Enter/↑↓ przed załadowaniem nie może przepaść.
  const FR = "find-replace";
  const frLoaded = () => typeof lazyFeatureLoaded !== "undefined" && lazyFeatureLoaded.has(FR);
  searchQueryEl?.addEventListener("focus", () => { ensureLazyFeature(FR).catch(() => {}); }, { once: true });
  searchQueryEl?.addEventListener("keydown", (e) => {
    if (frLoaded() || (e.key !== "Enter" && e.key !== "F3")) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    ensureLazyFeature(FR).then(() => runFindReplaceScan()).catch(() => {});
  });
  [["frPrevBtn", -1], ["frNextBtn", 1]].forEach(([id, d]) => {
    document.getElementById(id)?.addEventListener("click", (e) => {
      if (frLoaded()) return;
      e.stopImmediatePropagation();
      ensureLazyFeature(FR).then(() => stepFrMatch(d)).catch(() => {});
    });
  });
  // Telefon: ↑ ↓ tylko przy wpisanej frazie — puste pole szukania dostaje ich miejsce.
  const docToolbarEl = document.getElementById("docToolbar");
  const syncHasQuery = () => docToolbarEl?.classList.toggle("has-query", !!searchQueryEl?.value.trim());
  searchQueryEl?.addEventListener("input", () => {
    if (!searchQueryEl.value.trim() && searchPosEl) searchPosEl.textContent = "";
    syncHasQuery();
  });
  syncHasQuery();

  // ── liczniki: przy sekcjach panelu i na ⚙ ──────────────────────────────────
  function setSummaryCount(panelId, n) {
    const summary = document.querySelector(`#${panelId} > summary`);
    if (!summary) return;
    let badge = summary.querySelector(".panel-count");
    if (!n) { badge?.remove(); return; }
    if (!badge) {
      badge = document.createElement("span");
      badge.className = "panel-count";
      summary.querySelector(".chevron")?.before(badge);
    }
    badge.textContent = n > 999 ? "999+" : String(n);
  }
  function syncPanelCounts() {
    const hits = typeof frMatches !== "undefined" ? frMatches.length : 0;
    const pos = typeof frActiveIndex !== "undefined" ? frActiveIndex : -1;
    const grammar = typeof grammarScan !== "undefined" && grammarScan?.hits ? grammarScan.hits.length : 0;
    const fields = typeof placeholderScan !== "undefined" && placeholderScan?.fields ? placeholderScan.fields.length : 0;
    setSummaryCount("panel-search", hits);
    setSummaryCount("panel-grammar", grammar);
    setSummaryCount("panel-placeholders", fields);
    const review = typeof docReviewCounts !== "undefined" && originalFileBytes ? docReviewCounts.changes + docReviewCounts.comments : 0;
    setSummaryCount("panel-review", review);
    if (searchPosEl) {
      const q = searchQueryEl?.value.trim();
      searchPosEl.textContent = !q ? "" : hits ? `${pos + 1} / ${hits}` : "0";
      searchPosEl.classList.toggle("is-empty", !!q && !hits);
    }
    // ⚙: to, co czeka w panelu (sugestie korekty + pola do wypełnienia); trafienia są na pasku
    const waiting = grammar + fields;
    if (panelCountEl) {
      panelCountEl.hidden = !waiting;
      panelCountEl.textContent = waiting > 99 ? "99+" : String(waiting);
    }
  }

  // ── panel OBOK dokumentu (≥1024 px) ────────────────────────────────────────
  // Poniżej 1024 px panel zostaje wysuwany (nakładka / arkusz od dołu na telefonie).
  // Od 1024 px staje obok: .main dostaje lewy margines = szerokość panelu, a panel
  // (dalej position:fixed — własne przewijanie) stoi na wysokości obszaru dokumentu.
  // Otwarty/zamknięty — zapamiętane (tylko świadome kliknięcia, nie start aplikacji).
  function isDocked() { return dockMq.matches; }
  function syncDock() {
    const docked = isDocked();
    rootEl.classList.toggle("sidebar-docked", docked);
    if (!sidebarNode) return;
    if (!docked) { rootEl.style.removeProperty("--dock-top"); return; }
    rootEl.style.setProperty("--sidebar-w", `${Math.round(sidebarNode.getBoundingClientRect().width) || 320}px`);
    if (mainNode) {
      const top = Math.round(mainNode.getBoundingClientRect().top + (window.scrollY || 0));
      rootEl.style.setProperty("--dock-top", `${top}px`);
    }
  }
  function dockedInitialOpen() {
    if (!isDocked()) return false;
    try { return localStorage.getItem(DOCK_KEY) !== "0"; } catch (_) { return true; }
  }
  const origToggle = window.toggleSidebar;
  window.toggleSidebar = function toggleSidebarAndRemember(...args) {
    const r = origToggle.apply(this, args);
    if (isDocked()) { try { localStorage.setItem(DOCK_KEY, isSidebarOpen() ? "1" : "0"); } catch (_) {} }
    return r;
  };
  const origSetOpen = window.setSidebarOpen;
  window.setSidebarOpen = function setSidebarOpenDocked(open, ...rest) {
    syncDock();
    const r = origSetOpen.call(this, open, ...rest);
    panelToggle?.setAttribute("aria-expanded", String(!!open));
    panelToggle?.classList.toggle("is-on", !!open);
    // Szerokość dokumentu zmienia się bez zmiany okna — przelicz dopasowanie po przejściu.
    if (isDocked()) setTimeout(() => { if (typeof onViewportChange === "function") onViewportChange(); }, 320);
    return r;
  };
  document.getElementById("sidebarCloseBtn")?.addEventListener("click", () => {
    if (isSidebarOpen()) window.toggleSidebar();
    panelToggle?.focus();
  });
  const onDockChange = () => { syncDock(); if (!isDocked() && isSidebarOpen()) origSetOpen(false); };
  dockMq.addEventListener("change", onDockChange);
  window.addEventListener("resize", () => syncDock(), { passive: true });
  if (typeof ResizeObserver === "function" && heroEl) new ResizeObserver(() => syncDock()).observe(heroEl);

  // ── „Znajdź ustawienie…” ───────────────────────────────────────────────────
  // Filtruje sekcje panelu po tekście (tytuł, etykiety pól, przyciski, opisy). Pasujące
  // rozwija, resztę chowa; po wyczyszczeniu przywraca poprzedni stan otwarcia.
  const finder = document.getElementById("sidebarFinder");
  const finderClear = document.getElementById("sidebarFinderClear");
  const finderEmpty = document.getElementById("sidebarFinderEmpty");
  const normTxt = (x) => String(x || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/ł/g, "l");
  let finderSnapshot = null; // Map<details, open> sprzed szukania
  function sectionText(det) {
    if (det._dwbFindText) return det._dwbFindText;
    const parts = [det.textContent];
    det.querySelectorAll("[placeholder],[aria-label]").forEach((el) => {
      parts.push(el.getAttribute("placeholder") || "", el.getAttribute("aria-label") || "");
    });
    det._dwbFindText = normTxt(parts.join(" "));
    return det._dwbFindText;
  }
  function runFinder() {
    if (!finder || !sidebarNode) return;
    const q = normTxt(finder.value.trim());
    const panels = Array.from(sidebarNode.querySelectorAll("details.panel"));
    if (finderClear) finderClear.hidden = !q;
    sidebarNode.querySelectorAll(".finder-hit").forEach((el) => el.classList.remove("finder-hit"));
    if (!q) {
      panels.forEach((d) => { d.hidden = false; });
      sidebarNode.querySelectorAll(".sidebar-group").forEach((g) => { g.hidden = false; });
      if (finderSnapshot) { finderSnapshot.forEach((open, d) => { d.open = open; }); finderSnapshot = null; }
      if (finderEmpty) finderEmpty.hidden = true;
      return;
    }
    if (!finderSnapshot) finderSnapshot = new Map(panels.map((d) => [d, d.open]));
    const words = q.split(/\s+/).filter(Boolean);
    let any = 0;
    panels.forEach((d) => {
      const hit = words.every((w) => sectionText(d).includes(w));
      d.hidden = !hit;
      if (!hit) return;
      any++;
      d.open = true;
      d.querySelectorAll("label, button, .hint").forEach((el) => {
        if (words.every((w) => normTxt(`${el.textContent} ${el.getAttribute("aria-label") || ""}`).includes(w))) el.classList.add("finder-hit");
      });
    });
    sidebarNode.querySelectorAll(".sidebar-group").forEach((g) => {
      g.hidden = !g.querySelector("details.panel:not([hidden])");
    });
    if (finderEmpty) finderEmpty.hidden = any > 0;
  }
  if (finder) {
    finder.addEventListener("input", runFinder);
    finder.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && finder.value) { e.preventDefault(); e.stopPropagation(); finder.value = ""; runFinder(); }
      if (e.key === "Enter") {
        e.preventDefault();
        const hit = sidebarNode.querySelector(".finder-hit") || sidebarNode.querySelector("details.panel:not([hidden]) summary");
        if (!hit) return;
        hit.scrollIntoView({ block: "center" });
        const f = hit.matches("label") ? hit.querySelector("input,select,textarea,button") : hit;
        if (f) f.focus();
      }
    });
  }
  finderClear?.addEventListener("click", () => { finder.value = ""; runFinder(); finder.focus(); });

  // ── skróty sekcji (z nagłówków dokumentu) ──────────────────────────────────
  let sectionItems = []; // [{ el, chip }]
  function stripEnabled() {
    try { return localStorage.getItem(SECTION_STRIP_KEY) !== "0"; } catch (_) { return true; }
  }
  function setStripEnabled(on, opts = {}) {
    try { localStorage.setItem(SECTION_STRIP_KEY, on ? "1" : "0"); } catch (_) {}
    if (stripOpt) stripOpt.checked = on;
    renderSectionChips();
    if (!on && opts.announce) toast(t("sectionStripHidden"), "info");
  }
  function pickSectionHeadings(headings) {
    // tylko prawdziwe nagłówki (styl z pliku / znacznik) — zgadywane akapity to szum
    const list = (headings || []).filter((h) => h.el && h.label && h.source !== "guess");
    if (!list.length) return [];
    // Poziom najwyższy w dokumencie; jeśli sekcji jest mało — dołóż jeden poziom niżej.
    const top = Math.min(...list.map((h) => h.level || 2));
    let picked = list.filter((h) => (h.level || 2) === top);
    if (picked.length < 3) picked = list.filter((h) => (h.level || 2) <= top + 1);
    return picked.slice(0, SECTION_CHIPS_MAX);
  }
  function renderSectionChips() {
    if (!stripEl || !chipsEl) return;
    const picked = originalFileBytes && stripEnabled() ? pickSectionHeadings(documentStructure?.headings) : [];
    sectionItems = [];
    // Jedna sekcja to żadna nawigacja — pasek pokazujemy od dwóch.
    stripEl.hidden = picked.length < 2;
    if (stripEl.hidden) { chipsEl.replaceChildren(); syncDocViewportHeight(); return; }
    chipsEl.replaceChildren(...picked.map((h) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "section-chip";
      const label = h.label.replace(/\s+/g, " ").trim();
      chip.textContent = label.length > 42 ? `${label.slice(0, 41)}…` : label;
      if (label.length > 42) chip.setAttribute("aria-label", label);
      chip.addEventListener("click", () => scrollDocTo(h.el));
      sectionItems.push({ el: h.el, chip });
      return chip;
    }));
    attachOverflowFade(chipsEl);
    syncCurrentSection();
    syncDocViewportHeight();
  }
  function scrollDocTo(el) {
    if (!docViewportEl || !el) return;
    const top = el.getBoundingClientRect().top - docViewportEl.getBoundingClientRect().top + docViewportEl.scrollTop - 12;
    docViewportEl.scrollTo({ top: Math.max(0, top), behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth" });
  }
  let currentChip = null;
  function syncCurrentSection() {
    if (!sectionItems.length || !docViewportEl) return;
    const limit = docViewportEl.getBoundingClientRect().top + 90;
    let cur = sectionItems[0];
    for (const it of sectionItems) {
      if (it.el.getBoundingClientRect().top <= limit) cur = it; else break;
    }
    if (cur.chip === currentChip) return;
    currentChip?.classList.remove("is-current");
    currentChip?.removeAttribute("aria-current");
    currentChip = cur.chip;
    currentChip.classList.add("is-current");
    currentChip.setAttribute("aria-current", "true");
    // widoczny w rzędzie, ale bez scrollIntoView (ruszałby też stronę)
    const c = currentChip;
    const left = c.offsetLeft - chipsEl.offsetLeft;
    if (left < chipsEl.scrollLeft + 24 || left + c.offsetWidth > chipsEl.scrollLeft + chipsEl.clientWidth - 24) {
      chipsEl.scrollTo({ left: Math.max(0, left - 32), behavior: "smooth" });
    }
  }
  stripHideBtn?.addEventListener("click", () => setStripEnabled(false, { announce: true }));
  if (stripOpt) {
    stripOpt.checked = stripEnabled();
    stripOpt.addEventListener("change", () => setStripEnabled(stripOpt.checked));
  }

  // Każdy render dokumentu kończy się renderStructurePanel(documentStructure) — tu
  // odświeżamy nagłówek (słowa, akapity) i skróty sekcji.
  const origRenderStructure = window.renderStructurePanel;
  window.renderStructurePanel = function renderStructureAndFrame(...args) {
    const r = origRenderStructure.apply(this, args);
    renderSectionChips();
    syncFile();
    if (typeof refreshReviewCounts === "function") refreshReviewCounts();
    return r;
  };
  const origClear = window.clearDocumentState;
  window.clearDocumentState = function clearDocumentStateAndFrame(...args) {
    const r = origClear.apply(this, args);
    renderSectionChips();
    syncFile();
    syncPanelCounts();
    return r;
  };

  // ── telefon: nagłówek zwija się przy przewijaniu dokumentu ─────────────────
  // Rozwija się TYLKO na samej górze (scrollTop ≤ 6) albo po tapnięciu uchwytu — bez
  // tego progu zmiana wysokości nagłówka oscylowała z przewijaniem (lekcja z Sheet).
  let heroCollapsed = false;
  let manualExpandAt = null; // scrollTop, przy którym ktoś rozwinął ręcznie
  function setHeroCollapsed(on) {
    if (heroCollapsed === on) return;
    heroCollapsed = on;
    document.body.classList.toggle("hero-collapsed", on);
    heroGrip?.setAttribute("aria-expanded", String(!on));
    syncDocViewportHeight();
  }
  let scrollQueued = false;
  docViewportEl?.addEventListener("scroll", () => {
    if (scrollQueued) return;
    scrollQueued = true;
    requestAnimationFrame(() => {
      scrollQueued = false;
      syncCurrentSection();
      if (!narrowMq.matches || !originalFileBytes) return;
      const y = docViewportEl.scrollTop;
      if (manualExpandAt != null) {
        if (Math.abs(y - manualExpandAt) < 80) return;
        manualExpandAt = null;
      }
      if (!heroCollapsed && y > 48) setHeroCollapsed(true);
      else if (heroCollapsed && y <= 6) setHeroCollapsed(false);
    });
  }, { passive: true });
  heroGrip?.addEventListener("click", () => {
    const expand = heroCollapsed;
    setHeroCollapsed(!expand);
    manualExpandAt = expand ? docViewportEl?.scrollTop || 0 : null;
  });
  narrowMq.addEventListener("change", () => { if (!narrowMq.matches) setHeroCollapsed(false); });

  // ── upuszczenie pliku w dowolnym miejscu okna ──────────────────────────────
  let dragDepth = 0;
  const hasFiles = (e) => Array.from(e.dataTransfer?.types || []).includes("Files");
  window.addEventListener("dragenter", (e) => {
    if (!hasFiles(e)) return;
    dragDepth++;
    document.body.classList.add("drag-active");
  });
  window.addEventListener("dragleave", () => {
    dragDepth = Math.max(0, dragDepth - 1);
    if (!dragDepth) document.body.classList.remove("drag-active");
  });
  window.addEventListener("dragover", (e) => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener("drop", (e) => {
    dragDepth = 0;
    document.body.classList.remove("drag-active");
    if (e.defaultPrevented || !hasFiles(e)) return; // pole w panelu obsłużyło już upuszczenie
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (!file) return;
    if (hasUnsavedChanges && !window.confirm(t("closeDocWarn"))) return;
    ingestFile(file);
  });

  // ── pusty start ────────────────────────────────────────────────────────────
  document.getElementById("emptyOpenBtn")?.addEventListener("click", () => openFilePicker());
  document.getElementById("emptySampleBtn")?.addEventListener("click", () => loadSampleDocument());

  function onLanguageChange() {
    sidebarNode?.querySelectorAll("details.panel").forEach((d) => { d._dwbFindText = null; });
    syncFile();
    syncZoomNow();
    syncPanelCounts();
    renderSectionChips();
  }

  // ── start ──────────────────────────────────────────────────────────────────
  syncDock();
  syncModeSwitch();
  syncZoomNow();
  syncFile();

  return { syncFile, syncSave, syncPanelCounts, setMenuOpen, syncDock, isDocked, dockedInitialOpen, runFinder, renderSectionChips, setReadOnly, onLanguageChange, setHeroCollapsed, isHeroCollapsed: () => heroCollapsed };
})();
