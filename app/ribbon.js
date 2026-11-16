// ribbon.js — wstążka jak w Wordzie (docs/WSTAZKA.md; makieta docs/makiety/wstazka-v4.html).
//
// Od 769 px nagłówek i pasek nad dokumentem zamieniają się w dwa rzędy:
//   1. logo · Plik · karty · Szukaj · tryb (menu) · Zapisz · ⋯
//   2. Cofnij/Ponów · polecenia wybranej karty w jednej linii.
// Kontrolki to TE SAME elementy (te same id) — przenoszą się do gniazd [data-rb] wstążki, a na
// telefonie wracają na swoje miejsca (jak wyspa u dołu w app-frame.js), więc moduły i testy ich
// nie gubią. Nowe przyciski kart wołają istniejące funkcje (composeUi, dwbShortcuts) — bez kopii.
//
// Rząd, który się nie mieści, przewija się w bok (zasada Mateusza: nic nie chowa się do list ani
// pod „⋯”): krawędź wygasa, strzałka pokazuje, ile jest dalej („+3”), a za pierwszym razem w sesji
// rząd lekko drgnie w bok i wróci, żeby było widać, że się da (bez ruchu przy „ogranicz ruch”).

const dwbRibbon = (() => {
  const root = document.getElementById("ribbon");
  if (!root) return null;
  const deskMq = matchMedia("(min-width: 769px)");
  const reduceMq = matchMedia("(prefers-reduced-motion: reduce)");
  const HINT_KEY = "dwb-rb-hinted-v1";

  // ── przenoszenie kontrolek ────────────────────────────────────────────────
  const moves = Array.from(document.querySelectorAll("[data-rb]")).map((slot) => {
    const el = document.querySelector(slot.dataset.rb);
    return el ? { el, slot, mark: document.createComment("dwb-rb") } : null;
  }).filter(Boolean);
  let on = false;
  function place(desk) {
    if (desk === on) return;
    on = desk;
    for (const m of moves) {
      if (desk) { m.el.before(m.mark); m.slot.append(m.el); }
      else if (m.mark.parentNode) { m.mark.replaceWith(m.el); }
    }
    root.hidden = !desk;
    document.body.classList.toggle("rb-on", desk);
    closeModeMenu();
    if (desk) requestAnimationFrame(() => strips.forEach((s) => s.sync(true)));
    requestAnimationFrame(() => {
      if (typeof appFrame !== "undefined") appFrame.syncDock();
      if (typeof onViewportChange === "function") onViewportChange();
    });
  }
  // panel przypięty obok dokumentu liczy swoją górę od obszaru dokumentu — wstążka zmienia wysokość
  // (karta kontekstowa, Czytanie, wczytanie pliku), więc przelicz przy każdej zmianie
  if (typeof ResizeObserver === "function") {
    let rq = 0; // w następnej klatce — zmiana układu w samym wywołaniu to „ResizeObserver loop” w Safari
    new ResizeObserver(() => {
      cancelAnimationFrame(rq);
      rq = requestAnimationFrame(() => { if (on && typeof appFrame !== "undefined") appFrame.syncDock(); });
    }).observe(root);
  }

  // ── karty ─────────────────────────────────────────────────────────────────
  const tabs = Array.from(root.querySelectorAll(".rb-tab"));
  const panels = Array.from(root.querySelectorAll(".rb-panel"));
  let current = "home";
  function select(id, { focus = false } = {}) {
    const tab = tabs.find((t) => t.dataset.rtab === id && !t.hidden);
    if (!tab) return;
    current = id;
    tabs.forEach((t) => {
      const sel = t === tab;
      t.setAttribute("aria-selected", String(sel));
      t.tabIndex = sel ? 0 : -1;
    });
    panels.forEach((p) => { p.hidden = p.dataset.rtab !== id; });
    const cmds = strips.find((s) => s.key === "cmds");
    if (cmds) { cmds.sc.scrollLeft = 0; cmds.sync(true); }
    tabStrip?.reveal(tab);
    if (focus) tab.focus();
  }
  tabs.forEach((t) => t.addEventListener("click", () => select(t.dataset.rtab)));
  root.querySelector('[role="tablist"]')?.addEventListener("keydown", (e) => {
    if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(e.key)) return;
    const vis = tabs.filter((t) => !t.hidden);
    const i = vis.indexOf(document.activeElement);
    if (i < 0) return;
    e.preventDefault();
    const n = e.key === "Home" ? 0 : e.key === "End" ? vis.length - 1 : (i + (e.key === "ArrowRight" ? 1 : -1) + vis.length) % vis.length;
    select(vis[n].dataset.rtab, { focus: true });
  });

  // Karta kontekstowa „Tabela” (jak w Wordzie): jest, gdy kursor stoi w tabeli — compose-ui
  // pokazuje wtedy przycisk #tableToolsBtn. Nie przełącza się sama; za pojawieniem się mignie.
  const tableTab = tabs.find((t) => t.dataset.rtab === "table");
  const tableBtn = document.getElementById("tableToolsBtn");
  function syncTableTab() {
    if (!tableTab || !tableBtn) return;
    const show = !tableBtn.hidden;
    if (show === !tableTab.hidden) return;
    tableTab.hidden = !show;
    if (show && !reduceMq.matches) {
      tableTab.classList.remove("is-fresh");
      void tableTab.offsetWidth;
      tableTab.classList.add("is-fresh");
    }
    if (!show && current === "table") select("home");
    tabStrip?.sync();
  }
  if (tableBtn) new MutationObserver(syncTableTab).observe(tableBtn, { attributes: true, attributeFilter: ["hidden"] });

  // ── przewijany rząd z podpowiedzią ───────────────────────────────────────
  function hinted() { try { return JSON.parse(sessionStorage.getItem(HINT_KEY) || "[]"); } catch (_) { return []; } }
  function markHinted(key) { try { sessionStorage.setItem(HINT_KEY, JSON.stringify([...hinted(), key])); } catch (_) {} }
  function makeStrip(el) {
    const sc = el.querySelector(".rb-sc");
    const left = el.querySelector(".rb-nudge-l");
    const right = el.querySelector(".rb-nudge-r");
    const cnt = right?.querySelector(".rb-cnt");
    const key = el.dataset.strip;
    const items = () => Array.from(sc.querySelectorAll(key === "tabs" ? ".rb-tab:not([hidden])" : ".rb-panel:not([hidden]) > .rb-g > *"));
    function sync(first) {
      if (!on) return;
      const max = sc.scrollWidth - sc.clientWidth;
      const L = max > 1 && sc.scrollLeft > 2;
      const R = max > 1 && sc.scrollLeft < max - 2;
      el.classList.toggle("has-l", L);
      el.classList.toggle("has-r", R);
      const edge = sc.getBoundingClientRect().right;
      const n = R ? items().filter((c) => c.getBoundingClientRect().right > edge + 1 && c.offsetParent).length : 0;
      if (cnt) cnt.textContent = n ? `+${n}` : "";
      if (right) right.dataset.hintPl = right.dataset.hintEn = t("rbHidden", { n });
      if (first && R && !hinted().includes(key)) {
        markHinted(key);
        right.classList.add("is-hint");
        setTimeout(() => right.classList.remove("is-hint"), 1100);
        if (!reduceMq.matches) {
          sc.scrollBy({ left: 36, behavior: "smooth" });
          setTimeout(() => sc.scrollBy({ left: -36, behavior: "smooth" }), 420);
        }
      }
    }
    function reveal(node) {
      if (!node) return;
      const a = node.offsetLeft, b = a + node.offsetWidth;
      if (a < sc.scrollLeft) sc.scrollLeft = a - 24;
      else if (b > sc.scrollLeft + sc.clientWidth) sc.scrollLeft = b - sc.clientWidth + 40;
    }
    sc.addEventListener("scroll", () => sync(), { passive: true });
    // pionowe kółko przewija rząd w bok (Shift+kółko działa natywnie)
    sc.addEventListener("wheel", (e) => {
      if (Math.abs(e.deltaY) <= Math.abs(e.deltaX) || sc.scrollWidth <= sc.clientWidth + 1) return;
      sc.scrollLeft += e.deltaY;
      e.preventDefault();
    }, { passive: false });
    left?.addEventListener("click", () => sc.scrollBy({ left: -sc.clientWidth * 0.7, behavior: "smooth" }));
    right?.addEventListener("click", () => sc.scrollBy({ left: sc.clientWidth * 0.7, behavior: "smooth" }));
    // fokus z klawiatury na schowanym przycisku — rząd sam do niego dojeżdża
    sc.addEventListener("focusin", (e) => { const n = e.target.closest(".rb-tab, .rb-g > *"); if (n) reveal(n); });
    let rq = 0;
    if (typeof ResizeObserver === "function") new ResizeObserver(() => { cancelAnimationFrame(rq); rq = requestAnimationFrame(() => sync()); }).observe(sc);
    return { key, sc, sync, reveal };
  }
  const strips = Array.from(root.querySelectorAll(".rb-strip")).map(makeStrip);
  const tabStrip = strips.find((s) => s.key === "tabs");

  // ── tryb: Czytanie / Edycja / Sugerowanie ──────────────────────────────────
  const modeBtn = document.getElementById("rbModeBtn");
  const modeMenu = document.getElementById("rbModeMenu");
  const readModeEl = document.getElementById("readMode");
  const MODE_ICONS = {
    read: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>',
    edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
    suggest: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><path d="M12 6.5v4M10 8.5h4M10 13h4"/>',
  };
  const MODE_KEYS = { read: "modeRead", edit: "modeEdit", suggest: "modeSuggest" };
  function currentMode() {
    if (readModeEl?.checked) return "read";
    return typeof dwbTrack !== "undefined" && dwbTrack.isOn() ? "suggest" : "edit";
  }
  function syncMode() {
    const m = currentMode();
    if (modeBtn) {
      modeBtn.dataset.mode = m;
      modeBtn.querySelector(".rb-mode-ico").innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${MODE_ICONS[m]}</svg>`;
      modeBtn.querySelector(".rb-mode-txt").textContent = t(MODE_KEYS[m]);
      modeBtn.setAttribute("aria-label", t("rbModeAria", { mode: t(MODE_KEYS[m]) }));
    }
    document.body.classList.toggle("rb-reading", m === "read");
    root.querySelector('[data-rb-act="suggest"]')?.setAttribute("aria-pressed", String(m === "suggest"));
  }
  function closeModeMenu(refocus) {
    if (!modeMenu || modeMenu.hidden) return;
    modeMenu.hidden = true;
    modeBtn?.setAttribute("aria-expanded", "false");
    if (refocus) modeBtn?.focus();
  }
  function openModeMenu() {
    if (!modeMenu || !modeBtn) return;
    modeMenu.hidden = false;
    modeBtn.setAttribute("aria-expanded", "true");
    const r = modeBtn.getBoundingClientRect();
    const R = root.getBoundingClientRect();
    modeMenu.style.top = `${r.bottom - R.top + 6}px`;
    modeMenu.style.right = `${Math.max(8, R.right - r.right)}px`;
    (modeMenu.querySelector(".mode-btn.is-active") || modeMenu.querySelector(".mode-btn"))?.focus();
  }
  modeBtn?.addEventListener("click", () => (modeMenu.hidden ? openModeMenu() : closeModeMenu()));
  modeMenu?.addEventListener("click", (e) => { if (e.target.closest(".mode-btn")) setTimeout(() => closeModeMenu(true), 0); });
  modeMenu?.addEventListener("keydown", (e) => {
    const items = Array.from(modeMenu.querySelectorAll(".mode-btn"));
    const i = items.indexOf(document.activeElement);
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeModeMenu(true); }
    else if (e.key === "ArrowDown") { e.preventDefault(); items[(i + 1) % items.length]?.focus(); }
    else if (e.key === "ArrowUp") { e.preventDefault(); items[(i - 1 + items.length) % items.length]?.focus(); }
    else if (e.key === "Tab") closeModeMenu();
  });
  document.addEventListener("pointerdown", (e) => {
    if (!modeMenu?.hidden && !e.target.closest("#rbModeMenu, #rbModeBtn")) closeModeMenu();
  }, true);
  readModeEl?.addEventListener("change", syncMode);
  document.addEventListener("dwb:track", syncMode);
  const clickMode = (m) => document.querySelector(`.mode-btn[data-mode="${m}"]`)?.click();

  // ── polecenia kart ────────────────────────────────────────────────────────
  const cu = () => (typeof composeUi !== "undefined" ? composeUi : null);
  const MARKUP_ICONS = {
    all: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
    simple: '<path d="M4 6h16M4 12h9M4 18h12"/><path d="M2 4v16"/>',
    none: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>',
  };
  const ACT = {
    file: () => dwbShortcuts?.openToolPanel("panel-file"),
    replace: () => dwbShortcuts?.openReplace(),
    pageBreak: () => cu()?.insertPageBreak(),
    hf: () => cu()?.openHeaderFooterForm(),
    hrule: () => cu()?.insertHrule(),
    table: (b) => cu()?.openTablePicker(b),
    image: () => cu()?.pickImage(),
    textbox: () => { if (typeof dwbTextbox !== "undefined") dwbTextbox.insert(); },
    link: (b) => cu()?.openLinkForm(b),
    bookmark: (b) => cu()?.openBookmarkForm(b),
    comment: (b) => cu()?.openCommentForm(b),
    footnote: () => cu()?.insertNote("footnote"),
    endnote: () => cu()?.insertNote("endnote"),
    date: () => cu()?.insertToday(),
    symbol: (b) => cu()?.openSymbolPicker(b),
    toc: () => cu()?.insertToc(),
    pageSetup: () => cu()?.openPageSetup(),
    suggest: () => clickMode(currentMode() === "suggest" ? "edit" : "suggest"),
    read: () => clickMode("read"),
    edit: () => clickMode("edit"),
    fullscreen: () => document.getElementById("fullscreenMenuItem")?.click(),
    aiSettings: () => document.getElementById("aiMenuItem")?.click(),
    sectionStrip: () => document.getElementById("showSectionChips")?.click(),
    zoomOut: () => document.getElementById("zoomOutBtn")?.click(),
    paste: () => pasteFromClipboard(),
    cut: () => document.execCommand("cut"),
    copy: () => document.execCommand("copy"),
    highlight: () => cu()?.applyColor("highlight", "yellow"),
    bullets: () => cu()?.applyList("bullet"),
    numbering: () => cu()?.applyList("number"),
    outdent: () => cu()?.stepIndentAtCaret(-1),
    indent: () => cu()?.stepIndentAtCaret(1),
    spacingMenu: (b) => cu()?.openSpacingMenu(b),
    tabsForm: (b) => cu()?.openTabsForm(b),
    markupMenu: (b) => {
      // te same przyciski co w panelu Recenzja (#rvMarkup: Wszystkie / Proste / Bez)
      const opt = (v) => document.querySelector(`#rvMarkup [data-markup="${v}"]`);
      cu()?.openMenu(b, ["all", "simple", "none"].map((v) => ({
        label: t({ all: "markupAll", simple: "markupSimple", none: "markupNone" }[v]),
        desc: t({ all: "rbMarkupAllDesc", simple: "rbMarkupSimpleDesc", none: "rbMarkupNoneDesc" }[v]),
        icon: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${MARKUP_ICONS[v]}</svg>`,
        checked: opt(v)?.getAttribute("aria-checked") === "true",
        onPick: () => opt(v)?.click(),
      })));
    },
    acceptMenu: (b) => cu()?.openMenu(b, [
      { label: t("reviewAcceptAllBtn"), onPick: () => proxy("rvAcceptAllBtn") },
      { label: t("walkStart"), onPick: () => proxy("rvWalkBtn") },
    ]),
    rejectMenu: (b) => cu()?.openMenu(b, [
      { label: t("reviewRejectAllBtn"), onPick: () => proxy("rvRejectAllBtn") },
      { label: t("walkStart"), onPick: () => proxy("rvWalkBtn") },
    ]),
    walk: () => proxy("rvWalkBtn"),
    removeComments: () => proxy("rvRemoveCommentsBtn"),
    zoomIn: () => document.getElementById("zoomInBtn")?.click(),
    zoomFit: () => document.getElementById("zoomFitBtn")?.click(),
  };
  // przycisk z panelu Recenzja (ta sama akcja); nieaktywny = nie ma czego przeglądać
  function proxy(id) {
    const el = document.getElementById(id);
    if (!el || el.disabled || el.hidden) { toast(t("rbNoChanges"), "info"); return; }
    el.click();
  }
  // Wklej: schowek z formatowaniem trafia do TEJ SAMEJ obsługi co Ctrl/⌘+V (zdarzenie „paste” na
  // dokumencie — compose-ui / paste-rich), więc przycisk wkleja dokładnie jak skrót. Chrome pyta
  // raz o zgodę, Safari pokazuje dymek „Wklej”; odmowa lub brak API = podpowiedź o skrócie.
  async function pasteFromClipboard() {
    const target = document.activeElement?.closest?.("[contenteditable='true'], .docx-edit-root") || docCanvasEl;
    let dt = null;
    try {
      dt = new DataTransfer();
      if (navigator.clipboard?.read) {
        for (const item of await navigator.clipboard.read()) {
          for (const type of item.types) {
            const blob = await item.getType(type);
            if (type === "text/html" || type === "text/plain") dt.setData(type, await blob.text());
            else if (type.startsWith("image/")) dt.items.add(new File([blob], `obraz.${type.split("/")[1] || "png"}`, { type }));
          }
        }
      } else if (navigator.clipboard?.readText) {
        dt.setData("text/plain", await navigator.clipboard.readText());
      } else throw new Error("no clipboard API");
    } catch (_) {
      toast(t("rbPasteDenied"), "info");
      return;
    }
    if (!dt.types.length) return;
    const ev = new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true });
    target?.dispatchEvent(ev);
    // nikt nie obsłużył (zwykły tekst w polu) — wstaw jak przeglądarka
    if (!ev.defaultPrevented) document.execCommand("insertText", false, dt.getData("text/plain"));
  }
  root.addEventListener("click", (e) => {
    const b = e.target.closest("[data-rb-act]");
    if (!b || b.disabled) return;
    const a = b.dataset.rbAct;
    if (a.startsWith("panel:")) { dwbShortcuts?.openToolPanel(a.slice(6)); return; }
    if (a.startsWith("ps:")) { cu()?.openPageSetupPart(b, a.slice(3)); return; }
    ACT[a]?.(b);
  });
  // przyciski kart nie zabierają kursora z dokumentu (jak pasek formatowania)
  root.addEventListener("mousedown", (e) => {
    if (e.target.closest(".rb-btn, .rb-tab, .rb-nudge, .rb-file, .rb-edit")) e.preventDefault();
  });
  const fsItem = document.getElementById("fullscreenMenuItem");
  const fsBtn = root.querySelector('[data-rb-act="fullscreen"]');
  function syncExtras() {
    if (fsBtn && fsItem) fsBtn.hidden = fsItem.hidden;
    const strip = document.getElementById("showSectionChips");
    root.querySelector('[data-rb-act="sectionStrip"]')?.setAttribute("aria-pressed", String(!!strip?.checked));
  }
  if (fsItem) new MutationObserver(syncExtras).observe(fsItem, { attributes: true, attributeFilter: ["hidden"] });
  document.getElementById("showSectionChips")?.addEventListener("change", syncExtras);

  // ── start ─────────────────────────────────────────────────────────────────
  place(deskMq.matches);
  deskMq.addEventListener("change", (e) => place(e.matches));
  syncMode();
  syncTableTab();
  syncExtras();
  // zmiana języka: napis trybu i podpowiedzi strzałek (jak doc-comfort.js — owinięcie applyLanguage)
  if (typeof window.applyLanguage === "function") {
    const orig = window.applyLanguage;
    window.applyLanguage = function applyLanguageRibbon(...args) {
      const r = orig.apply(this, args);
      syncMode();
      strips.forEach((s) => s.sync());
      return r;
    };
  }

  // przełącza na kartę, na której stoi kontrolka (np. skrót klawiszowy otwiera jej okienko)
  function reveal(target) {
    const el = typeof target === "string" ? document.querySelector(target) : target;
    const p = el?.closest?.(".rb-panel");
    if (p) select(p.dataset.rtab);
    return !!p;
  }

  return { select, reveal, isOn: () => on, current: () => current };
})();
