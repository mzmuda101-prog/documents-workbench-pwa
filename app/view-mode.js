// view-mode.js — tryb skupienia i pełny ekran (paczka E, prośba Mateusza 2026-09-28:
// „żeby główna rzecz, czyli dokument, zajmowała jak najwięcej miejsca + opcja
// fullscreena lub semi-fullscreena”).
//
// Tryb skupienia (semi-fullscreen) — przycisk na pasku, Ctrl/⌘+Alt+F, pozycja w menu ⋯:
//   znika nagłówek, pasek stanu, skróty sekcji i panel; zostaje pasek nad dokumentem.
//   „Zapisz” i ⋯ przeprowadzają się z nagłówka na pasek (te same elementy — cała ich
//   obsługa działa bez zmian), więc zapis i menu są pod ręką. Esc wychodzi.
// Pełny ekran — menu ⋯ (tam, gdzie przeglądarka pozwala; iPhone nie pozwala): tryb
//   skupienia + Fullscreen API. Wyjście z pełnego ekranu (Esc przeglądarki) kończy też
//   tryb skupienia, jeśli to pełny ekran go włączył.
// Stan nie jest zapamiętywany — to tryb „na chwilę”, nie ustawienie.

const dwbView = (() => {
  const btn = document.getElementById("focusModeBtn");
  const menuItem = document.getElementById("focusModeMenuItem");
  const fsItem = document.getElementById("fullscreenMenuItem");
  const slot = document.getElementById("fmHeroSlot");
  const heroRight = document.querySelector(".hero-right");
  const saveBtn = document.getElementById("heroSaveBtn");
  const menuBtn = document.getElementById("appMenuBtn");
  const fsSupported = !!(document.fullscreenEnabled || document.webkitFullscreenEnabled);
  let active = false;
  let panelWasOpen = false;
  let fromFullscreen = false;

  function fsElement() { return document.fullscreenElement || document.webkitFullscreenElement || null; }

  function moveHeroControls(intoToolbar) {
    if (!slot || !heroRight || !saveBtn || !menuBtn) return;
    if (intoToolbar) {
      slot.append(saveBtn, menuBtn);
    } else {
      // z powrotem na swoje miejsca w nagłówku (menu ⋯ samo mieszka w <body>)
      heroRight.append(saveBtn, menuBtn);
    }
    if (typeof appFrame !== "undefined") appFrame.setMenuOpen(false);
  }

  function set(on, opts = {}) {
    on = !!on && !!originalFileBytes;
    if (on === active) return;
    active = on;
    rootEl.classList.toggle("focus-mode", on);
    btn?.setAttribute("aria-pressed", String(on));
    btn?.classList.toggle("is-on", on);
    if (btn) {
      btn.setAttribute("aria-label", t(on ? "focusModeExitAria" : "focusModeAria"));
      btn.dataset.hintPl = on ? "Wyjdź z trybu skupienia (Esc)" : "Tryb skupienia — tylko dokument (Ctrl/⌘+Alt+F · Esc wychodzi)";
      btn.dataset.hintEn = on ? "Exit focus mode (Esc)" : "Focus mode — just the document (Ctrl/⌘+Alt+F · Esc exits)";
    }
    if (on) {
      panelWasOpen = isSidebarOpen();
      if (panelWasOpen) setSidebarOpen(false); // bez zapamiętywania wyboru (to nie jest świadome zamknięcie)
      moveHeroControls(true);
      // na dotyku nie ma Esc — mówimy o przycisku
      if (!opts.silent) toast(t(matchMedia("(pointer: coarse)").matches ? "focusModeOnTouch" : "focusModeOn"), "info");
    } else {
      moveHeroControls(false);
      if (panelWasOpen) setSidebarOpen(true);
      if (fsElement() && !opts.keepFullscreen) exitFullscreen();
      fromFullscreen = false;
    }
    syncMenu();
    syncDocViewportHeight();
    if (typeof onViewportChange === "function") setTimeout(onViewportChange, 60);
  }

  function toggle() { set(!active); }

  async function enterFullscreen() {
    if (!fsSupported || !originalFileBytes) return;
    set(true, { silent: true });
    fromFullscreen = true;
    try {
      const el = document.documentElement;
      await (el.requestFullscreen ? el.requestFullscreen({ navigationUI: "hide" }) : el.webkitRequestFullscreen());
    } catch (_) {
      toast(t("fullscreenFailed"), "warning");
    }
  }
  function exitFullscreen() {
    try { (document.exitFullscreen || document.webkitExitFullscreen)?.call(document); } catch (_) {}
  }
  const onFsChange = () => {
    // Esc przeglądarki kończy pełny ekran — kończymy też tryb skupienia, jeśli to on go włączył
    if (!fsElement() && fromFullscreen && active) set(false, { keepFullscreen: true });
    syncMenu();
  };
  document.addEventListener("fullscreenchange", onFsChange);
  document.addEventListener("webkitfullscreenchange", onFsChange);

  function syncMenu() {
    if (menuItem) {
      menuItem.hidden = !originalFileBytes;
      const label = menuItem.querySelector(".app-menu-item-text");
      if (label) label.textContent = t(active ? "focusModeMenuExit" : "focusModeMenu");
    }
    if (fsItem) {
      fsItem.hidden = !fsSupported || !originalFileBytes;
      const label = fsItem.querySelector(".app-menu-item-text");
      if (label) label.textContent = t(fsElement() ? "fullscreenMenuExit" : "fullscreenMenu");
    }
  }

  btn?.addEventListener("click", toggle);
  menuItem?.addEventListener("click", toggle);
  fsItem?.addEventListener("click", () => (fsElement() ? exitFullscreen() : enterFullscreen()));

  // Zamknięcie dokumentu = koniec trybu.
  const origClear = window.clearDocumentState;
  window.clearDocumentState = function clearDocumentStateView(...args) {
    set(false, { silent: true });
    const r = origClear.apply(this, args);
    syncMenu();
    return r;
  };
  const origSyncShell = window.syncDocumentShellClass;
  window.syncDocumentShellClass = function syncDocumentShellClassView(...args) {
    const r = origSyncShell.apply(this, args);
    syncMenu();
    return r;
  };

  syncMenu();
  return { set, toggle, isActive: () => active, enterFullscreen, syncMenu };
})();
