// phone-bar.js — górny pasek telefonu jak w makiecie v4 (etap 7 wstążki; docs/WSTAZKA.md).
//
// Do 768 px, z otwartym plikiem: ← (zamknij do ekranu startowego) · Cofnij (w Edycji) · nazwa
// pliku i „N niezapisane · tryb” · przycisk trybu (menu: Edycja / Sugerowanie / Czytanie) · ⋯.
// Wszystko klika istniejące kontrolki (Zamknij z menu ⋯, #undoBtn, przyciski trybu, Zapisz), więc
// pytania o niezapisane zmiany, historia Cofnij i przełączanie trybu działają jak dotąd.

const dwbPhoneBar = (() => {
  const $ = (id) => document.getElementById(id);
  const back = $("phBackBtn"), undo = $("phUndoBtn"), status = $("phStatus"), modeBtn = $("phModeBtn");
  if (!back || !undo || !status || !modeBtn) return null;
  const readModeEl = $("readMode");
  const ICONS = {
    read: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>',
    edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
    suggest: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><path d="M12 6.5v4M10 8.5h4M10 13h4"/>',
  };
  const svg = (d, s = 18) => `<svg width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
  const NAME = { read: "modeRead", edit: "modeEdit", suggest: "modeSuggest" };
  const DESC = { edit: "phModeEditDesc", suggest: "phModeSuggestDesc", read: "phModeReadDesc" };

  function mode() {
    if (readModeEl?.checked) return "read";
    return typeof dwbTrack !== "undefined" && dwbTrack.isOn() ? "suggest" : "edit";
  }
  function sync() {
    const m = mode();
    document.body.classList.toggle("ph-reading", m === "read");
    modeBtn.dataset.mode = m;
    modeBtn.querySelector(".ph-mode-ico").innerHTML = svg(ICONS[m]);
    modeBtn.setAttribute("aria-label", t("phModeAria", { mode: t(NAME[m]) }));
    const saveBtn = $("heroSaveBtn");
    const dirty = !!saveBtn?.classList.contains("is-dirty");
    const n = dirty ? $("heroSaveCount")?.textContent || "1" : "";
    status.textContent = m === "read" ? t(NAME.read) : [n ? t("phUnsaved", { n }) : "", t(NAME[m])].filter(Boolean).join(" · ");
    const src = $("undoBtn");
    undo.disabled = !src || src.disabled;
    if (islMode) {
      islMode.dataset.mode = m;
      islMode.querySelector(".ph-mode-ico").innerHTML = svg(ICONS[m]);
      islMode.setAttribute("aria-label", t("phModeAria", { mode: t(NAME[m]) }));
    }
    if (islSave) {
      islSave.classList.toggle("is-dirty", dirty);
      const badge = islSave.querySelector(".ph-isl-n");
      badge.hidden = !n;
      badge.textContent = n;
      islSave.setAttribute("aria-label", saveBtn?.getAttribute("aria-label") || t("save"));
    }
  }
  let q = 0;
  const soon = () => { cancelAnimationFrame(q); q = requestAnimationFrame(sync); };

  back.addEventListener("click", () => $("closeDocBtn")?.click());
  undo.addEventListener("mousedown", (e) => e.preventDefault()); // kursor zostaje (pointerdown w WebKit zjadał klik)
  undo.addEventListener("click", () => $("undoBtn")?.click());
  $("phSaveMenuItem")?.addEventListener("click", () => {
    if (typeof appFrame !== "undefined") appFrame.setMenuOpen?.(false);
    $("heroSaveBtn")?.click();
  });
  const islSave = $("phIslandSave"), islMode = $("phIslandMode");
  function openModeMenu(anchor) {
    if (typeof composeUi === "undefined") return;
    const cur = mode();
    composeUi.openMenu(anchor, ["edit", "suggest", "read"].map((m) => ({
      label: t(NAME[m]), desc: t(DESC[m]), icon: svg(ICONS[m], 16), checked: m === cur,
      onPick: () => document.querySelector(`.mode-btn[data-mode="${m}"]`)?.click(),
    })));
  }
  // wyspa u dołu (prośba Mateusza 2026-11-17): Zapisz i tryb pod kciukiem — górny pasek zwija się przy przewijaniu
  islMode?.addEventListener("click", () => openModeMenu(islMode));
  islSave?.addEventListener("mousedown", (e) => e.preventDefault());
  islSave?.addEventListener("click", () => $("heroSaveBtn")?.click());
  modeBtn.addEventListener("click", () => {
    if (typeof composeUi === "undefined") return;
    const cur = mode();
    modeBtn.setAttribute("aria-expanded", "true");
    composeUi.openMenu(modeBtn, ["edit", "suggest", "read"].map((m) => ({
      label: t(NAME[m]), desc: t(DESC[m]), icon: svg(ICONS[m], 16), checked: m === cur,
      onPick: () => document.querySelector(`.mode-btn[data-mode="${m}"]`)?.click(),
    })));
    // menu zamyka się samo (compose-ui) — stan przycisku wraca przy najbliższej zmianie
    setTimeout(() => modeBtn.setAttribute("aria-expanded", "false"), 0);
  });

  readModeEl?.addEventListener("change", soon);
  document.addEventListener("dwb:track", soon);
  for (const id of ["heroSaveBtn", "heroSaveCount", "undoBtn"]) {
    const el = $(id);
    if (el) new MutationObserver(soon).observe(el, { attributes: true, childList: true, characterData: true, subtree: true });
  }
  if (typeof window.applyLanguage === "function") {
    const orig = window.applyLanguage;
    window.applyLanguage = function applyLanguagePhone(...args) { const r = orig.apply(this, args); sync(); return r; };
  }
  sync();

  // ── 7b: rząd poleceń karty u dołu (nad klawiaturą) i arkusz kart (makieta v4 .acc / .sheet) ──
  // W Edycji na telefonie. Polecenia to te same akcje co wstążka (dwbRibbon.run) albo kliknięcie
  // istniejącej kontrolki (click:#id); menu formatowania otwierają się przy przycisku telefonu
  // (composeUi.openMenuFor). Ikona: z przycisku wstążki, gdy ta sama funkcja tam jest.
  const panel = document.querySelector(".doc-panel");
  const ribbonRoot = $("ribbon");
  const IC = {
    undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-1"/>',
    redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h1"/>',
    bold: '<path d="M6 12h9a4 4 0 0 1 0 8H6V4h8a4 4 0 0 1 0 8"/>',
    italic: '<path d="M19 4h-9M14 20H5M15 4 9 20"/>',
    under: '<path d="M6 4v6a6 6 0 0 0 12 0V4M4 20h16"/>',
    strike: '<path d="M16 4H9a3 3 0 0 0-2.83 4M14 12a4 4 0 0 1 0 8H6M4 12h16"/>',
    color: '<path d="M4 20h16"/><path d="m6 16 6-12 6 12M8 12h8"/>',
    pil: '<path d="M13 4v16M17 4v16M19 4H9.5a4.5 4.5 0 0 0 0 9H13"/>',
    al: '<path d="M21 6H3M15 12H3M17 18H3"/>',
    spacing: '<path d="M5 4v16M2.5 6.5 5 4l2.5 2.5M2.5 17.5 5 20l2.5-2.5M11 6h10M11 12h10M11 18h10"/>',
    caseI: '<path d="m3 15 4-8 4 8M4 13h6"/><circle cx="18" cy="12" r="3"/><path d="M21 9v6"/>',
    clear: '<path d="M4 7V4h16v3M5 20h6M13 4 8 20M15 15l5 5M20 15l-5 5"/>',
    listlv: '<path d="M9 6h11M13 12h7M13 18h7"/><circle cx="4.5" cy="6" r="1"/><circle cx="8.5" cy="12" r="1"/><circle cx="8.5" cy="18" r="1"/>',
    spell: '<path d="m3 15 4-8 4 8M4 13h6M15 11h4.5a2 2 0 0 1 0 4H15V7h4a2 2 0 0 1 0 4"/><path d="m14 20 2 2 4-4"/>',
    cmt: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
    track: ICONS.suggest,
    navp: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18M13 8h4M13 12h4"/>',
    search: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>',
    focus: '<path d="M8 3H5a2 2 0 0 0-2 2v3M21 8V5a2 2 0 0 0-2-2h-3M3 16v3a2 2 0 0 0 2 2h3M16 21h3a2 2 0 0 0 2-2v-3"/>',
    fit: '<path d="M3 5v14"/><path d="M21 5v14"/><path d="m7 12h10"/><path d="m10 9-3 3 3 3"/><path d="m14 9 3 3-3 3"/>',
    marks: '<path d="M13 4v16M17 4v16M19 4H9.5a4.5 4.5 0 0 0 0 9H13"/>',
    view: '<rect x="2.5" y="4" width="19" height="12" rx="1.5"/><path d="M8 20h8M12 16v4"/>',
    wand: '<path d="m21.64 3.64-1.28-1.28a1.21 1.21 0 0 0-1.72 0L2.36 18.64a1.21 1.21 0 0 0 0 1.72l1.28 1.28a1.2 1.2 0 0 0 1.72 0L21.64 5.36a1.2 1.2 0 0 0 0-1.72M14 7l3 3M5 6v4M19 14v4M10 2v2M7 8H3M21 16h-4M11 3H9"/>',
    ask: '<circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3M12 17h.01"/>',
    sense: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3M8 11h.01M11 11h.01M14 11h.01"/>',
    rowA: '<rect x="3" y="12" width="18" height="8" rx="1"/><line x1="12" y1="3" x2="12" y2="9"/><line x1="9" y1="6" x2="15" y2="6"/>',
    rowB: '<rect x="3" y="4" width="18" height="8" rx="1"/><line x1="12" y1="15" x2="12" y2="21"/><line x1="9" y1="18" x2="15" y2="18"/>',
    colL: '<rect x="12" y="3" width="8" height="18" rx="1"/><line x1="3" y1="12" x2="9" y2="12"/><line x1="6" y1="9" x2="6" y2="15"/>',
    colR: '<rect x="4" y="3" width="8" height="18" rx="1"/><line x1="15" y1="12" x2="21" y2="12"/><line x1="18" y1="9" x2="18" y2="15"/>',
    down: '<path d="m6 9 6 6 6-6"/>',
  };
  // [akcja, klucz podpisu, ikona (gdy wstążka nie ma swojej), id kontrolki do stanu „wciśnięty”]
  const TABS = {
    // Cofnij obok Ponów: górny pasek z „Cofnij” chowa się przy otwartym arkuszu — sam „Ponów” wyglądał jak brak cofania
    home: [["click:#undoBtn", "phUndo", IC.undo], ["click:#redoBtn", "phRedo", IC.redo], ["click:#fmtBold", "phBold", IC.bold, "fmtBold"], ["click:#fmtItalic", "phItalic", IC.italic, "fmtItalic"], ["click:#fmtUnderline", "phUnderline", IC.under, "fmtUnderline"],
      ["click:#fmtStrike", "phStrike", IC.strike, "fmtStrike"], ["menu:fmtColorBtn", "phColor", IC.color], ["highlight", "rbHighlight"], ["style", "paraStyleLabel", IC.pil],
      ["bullets", "rbBullets"], ["numbering", "rbNumbering"], ["menu:fmtListBtn", "listLabel", IC.listlv], ["menu:fmtAlignBtn", "phAlign", IC.al], ["menu:fmtSpacingBtn", "spacingLabel", IC.spacing],
      ["indent", "indentMore"], ["outdent", "indentLess"], ["menu:fmtCaseBtn", "phCase", IC.caseI], ["click:#fmtClearBtn", "phClear", IC.clear]],
    insert: [["table", "insertTable"], ["image", "insertImage"], ["link", "insertLink"], ["footnote", "insertFootnote"], ["pageBreak", "insertPageBreak"], ["hrule", "insertHrule"],
      ["date", "insertDate"], ["symbol", "insertSymbol"], ["panel:panel-snippets", "rbSnippets"], ["hf", "hfInsert"], ["textbox", "insertTextbox"], ["bookmark", "insertBookmark"],
      ["endnote", "insertEndnote"], ["toc", "tocInsert"], ["formFields", "rbFormField"]],
    layout: [["ps:margins", "pageMargins"], ["ps:orient", "pageOrient"], ["ps:size", "pageSize"], ["ps:cols", "colsCap"], ["ps:breaks", "breaksCap"], ["tabsForm", "tabsMore"],
      ["ps:valign", "rbVAlign"], ["layoutForm", "spacingMore"]],
    review: [["panel:panel-grammar", "railProof", IC.spell], ["comment", "rbNewComment"], ["panel:panel-review", "phComments", IC.cmt], ["suggest", "modeSuggest", IC.track],
      ["acceptAll", "reviewAcceptAllBtn"], ["rejectAll", "reviewRejectAllBtn"], ["panel:panel-compare", "comparePanel"], ["panel:panel-stats", "rbStats"], ["walk", "walkStart"],
      ["markupMenu", "markupLabel"], ["removeComments", "reviewRemoveComments"]],
    view: [["panel:panel-structure", "railNav", IC.navp], ["panel:panel-search", "railFind", IC.search], ["click:#focusModeBtn", "focusModeAria", IC.focus], ["click:#zoomFitBtn", "zoomFit", IC.fit],
      ["click:#formatMarksBtn", "formatMarksLabel", IC.marks, "formatMarksBtn"], ["panel:panel-view", "phViewOpts", IC.view], ["sectionStrip", "rbSectionStrip"]],
    ai: [["click:#rewriteBtn", "rewriteBtn", IC.wand], ["panel:panel-ask", "askTitle", IC.ask], ["panel:panel-semantic-search", "semanticPanel", IC.sense], ["aiSettings", "aiMenu"]],
    table: [["tbl:rowAbove", "tableRowAbove", IC.rowA], ["tbl:rowBelow", "tableRowBelow", IC.rowB], ["tbl:colLeft", "tableColLeft", IC.colL], ["tbl:colRight", "tableColRight", IC.colR],
      ["tblp:merge", "tableMergeGroup"], ["tbl:headerRow", "rbHeaderRow"], ["tblp:borders", "rbBorders"], ["tblp:shade", "tableShadeGroup"], ["tblp:sort", "tableSortGroup"], ["tblp:delete", "tableDeleteGroup"]],
  };
  const TAB_KEYS = { home: "phTabHome", insert: "rbTabInsert", layout: "rbTabLayout", review: "rbTabReview", view: "rbTabView", ai: "rbTabAi", table: "rbTabTable" };
  const ribbonIcon = (act) => ribbonRoot?.querySelector(`.rb-panel:not(.rb-cl) [data-rb-act="${CSS.escape(act)}"] svg:not(.rb-dd)`)?.outerHTML;
  const iconFor = ([act, , icon]) => (icon ? svg(icon, 20) : ribbonIcon(act)?.replace(/width="\d+" height="\d+"/, 'width="20" height="20"') || svg(IC.cmt, 20));

  const bar = document.createElement("div");
  bar.className = "ph-cmd";
  bar.id = "phCmd";
  bar.hidden = true;
  bar.innerHTML = '<div class="ph-row" role="toolbar"><button type="button" class="ph-tsw" aria-expanded="false"><span class="ph-tsw-t"></span>' +
    '<svg class="rb-dd" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" aria-hidden="true"><path d="m6 9 6 6 6-6"/></svg></button>' +
    '<div class="ph-strip"></div><button type="button" class="ph-hidekb"></button></div>' +
    '<div class="ph-sheet" hidden><div class="ph-sheet-h"><div class="ph-chips" role="tablist"></div><button type="button" class="ph-done"></button></div><div class="ph-sheet-b" role="tabpanel"></div></div>';
  panel?.append(bar);
  const row = bar.querySelector(".ph-row"), tsw = bar.querySelector(".ph-tsw"), strip = bar.querySelector(".ph-strip"), hideKb = bar.querySelector(".ph-hidekb");
  const sheet = bar.querySelector(".ph-sheet"), chips = bar.querySelector(".ph-chips"), done = bar.querySelector(".ph-done"), body = bar.querySelector(".ph-sheet-b");
  hideKb.innerHTML = svg(IC.down, 20);
  if (typeof attachOverflowFade === "function") { attachOverflowFade(strip); attachOverflowFade(chips); }
  let tab = "home";
  let savedRange = null;
  const tableOn = () => { const b = $("tableToolsBtn"); return !!b && !b.hidden; };
  const tabs = () => Object.keys(TABS).filter((k) => k !== "table" || tableOn());

  function cmdButton(it, withLabel) {
    const [act, key, , mirror] = it;
    const b = document.createElement("button");
    b.type = "button";
    b.className = withLabel ? "ph-item" : "ph-ic";
    b.dataset.act = act;
    if (mirror) b.dataset.mirror = mirror;
    b.setAttribute("aria-label", t(key));
    b.innerHTML = iconFor(it);
    if (withLabel) { const l = document.createElement("span"); l.textContent = t(key); b.append(l); }
    return b;
  }
  function render() {
    if (!tabs().includes(tab)) tab = "home";
    bar.querySelector(".ph-tsw-t").textContent = t(TAB_KEYS[tab]);
    tsw.classList.toggle("is-ctx", tab === "table");
    tsw.setAttribute("aria-label", t("phTabAria", { tab: t(TAB_KEYS[tab]) }));
    row.setAttribute("aria-label", t(TAB_KEYS[tab]));
    hideKb.setAttribute("aria-label", t("phHideKb"));
    done.textContent = t("phDone");
    strip.replaceChildren(...TABS[tab].map((it) => cmdButton(it, false)));
    chips.replaceChildren(...tabs().map((k) => {
      const c = document.createElement("button");
      c.type = "button";
      c.className = "ph-chip" + (k === "table" ? " is-ctx" : "");
      c.setAttribute("role", "tab");
      c.setAttribute("aria-selected", String(k === tab));
      c.dataset.tab = k;
      c.textContent = t(TAB_KEYS[k]);
      return c;
    }));
    body.replaceChildren(...TABS[tab].map((it) => cmdButton(it, true)));
    body.setAttribute("aria-label", t(TAB_KEYS[tab]));
    syncPressed();
  }
  function syncPressed() {
    bar.querySelectorAll("[data-mirror]").forEach((b) => {
      const src = $(b.dataset.mirror);
      const on = !!src && (src.classList.contains("is-on") || src.getAttribute("aria-pressed") === "true");
      b.setAttribute("aria-pressed", String(on));
    });
    bar.querySelectorAll('[data-act="suggest"]').forEach((b) => b.setAttribute("aria-pressed", String(mode() === "suggest")));
  }

  // arkusz w miejscu klawiatury: zaznaczenie zostaje (zapamiętane), klawiatura się chowa
  function openSheet(open) {
    if (open) {
      const sel = getSelection();
      savedRange = sel?.rangeCount && docCanvasEl?.contains(sel.anchorNode) ? sel.getRangeAt(0).cloneRange() : null;
      const a = document.activeElement;
      if (a && docCanvasEl?.contains(a)) a.blur();
      if (savedRange) { sel.removeAllRanges(); sel.addRange(savedRange); }
    }
    sheet.hidden = !open;
    row.hidden = open;
    tsw.setAttribute("aria-expanded", String(open));
    bar.classList.toggle("is-sheet", open);
    place();
  }
  function restoreSel() {
    if (!savedRange) return;
    const sel = getSelection();
    if (sel && !(sel.rangeCount && docCanvasEl?.contains(sel.anchorNode))) { sel.removeAllRanges(); sel.addRange(savedRange); }
  }
  function styleMenu(anchor) {
    const sel = $("fmtParaStyle");
    if (!sel || typeof composeUi === "undefined") return;
    const opts = Array.from(sel.options).filter((o) => o.value && !o.value.startsWith("act:") && !o.value.startsWith("plus:") && !o.dataset.plus && !o.dataset.cur && !o.disabled && !o.hidden);
    composeUi.openMenu(anchor, opts.map((o) => ({ label: o.textContent.trim(), checked: o.value === sel.value, onPick: () => composeUi.applyStyle(o.value) })));
  }
  function runCmd(act, btn) {
    restoreSel();
    if (act === "style") { styleMenu(btn); return; }
    if (act.startsWith("menu:")) { composeUi?.openMenuFor(act.slice(5), btn); return; }
    if (typeof dwbRibbon !== "undefined" && dwbRibbon) dwbRibbon.run(act, btn);
    setTimeout(syncPressed, 120);
  }
  bar.addEventListener("mousedown", (e) => { if (e.target.closest("button")) e.preventDefault(); }); // kursor zostaje
  bar.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    if (b === tsw) { openSheet(true); return; }
    if (b === done) { openSheet(false); return; }
    if (b === hideKb) { $("docDismissBtn")?.click(); placeSoon(); return; } // to samo co „Gotowe” (touch.js)
    if (b.classList.contains("ph-chip")) { tab = b.dataset.tab; render(); return; }
    if (b.dataset.act) runCmd(b.dataset.act, b);
  });

  // ── kiedy i gdzie ─────────────────────────────────────────────────────────
  const narrow = matchMedia("(max-width: 768px)");
  const origInset = window.dwbDocBottomInset || (() => 0);
  let barH = 0;
  function shown() { return narrow.matches && document.body.classList.contains("has-document") && mode() !== "read" && !document.documentElement.classList.contains("focus-mode"); }
  function place() {
    const on = shown();
    if (bar.hidden === on) { bar.hidden = !on; if (!on) openSheet(false); }
    document.body.classList.toggle("ph-bar-on", on);
    if (!on) { barH = 0; panel?.style.removeProperty("--ph-bar-h"); return; }
    const kb = document.body.classList.contains("kb-open");
    bar.classList.toggle("is-kb", kb);
    // ⌄ = koniec pisania (dawne „Gotowe”): przy klawiaturze ekranowej i przy pisaniu z zewnętrznej
    const a = document.activeElement;
    hideKb.hidden = !(kb || (a && docCanvasEl?.contains(a) && a.isContentEditable));
    // nad wyspą, gdy jest; przy pisaniu (wyspa schowana) — na dole obszaru dokumentu = nad klawiaturą
    const below = panel && docViewportEl ? Math.max(0, panel.getBoundingClientRect().bottom - docViewportEl.getBoundingClientRect().bottom) : 0;
    const inset = origInset();
    bar.style.bottom = `${Math.round(below + (inset ? inset + 4 : kb ? 0 : 10))}px`;
    const h = bar.offsetHeight;
    if (h !== barH) { barH = h; panel?.style.setProperty("--ph-bar-h", `${h}px`); }
  }
  window.dwbDocBottomInset = () => origInset() + (shown() && !bar.hidden ? barH + 4 : 0);
  let pq = 0;
  const placeSoon = () => { cancelAnimationFrame(pq); pq = requestAnimationFrame(place); };
  new MutationObserver(placeSoon).observe(document.body, { attributes: true, attributeFilter: ["class"] });
  new MutationObserver(placeSoon).observe(document.documentElement, { attributes: true, attributeFilter: ["class"] });
  const island = $("docIsland");
  if (typeof ResizeObserver === "function") { const ro = new ResizeObserver(placeSoon); if (island) ro.observe(island); ro.observe(bar); if (panel) ro.observe(panel); }
  window.visualViewport?.addEventListener("resize", placeSoon);
  document.addEventListener("focusin", placeSoon);
  document.addEventListener("focusout", () => setTimeout(placeSoon, 0));
  narrow.addEventListener("change", placeSoon);
  readModeEl?.addEventListener("change", placeSoon);
  const tblBtn = $("tableToolsBtn");
  if (tblBtn) new MutationObserver(() => { render(); placeSoon(); }).observe(tblBtn, { attributes: true, attributeFilter: ["hidden"] });
  let sq = 0;
  document.addEventListener("selectionchange", () => { if (bar.hidden) return; cancelAnimationFrame(sq); sq = requestAnimationFrame(syncPressed); });
  document.addEventListener("dwb:track", syncPressed);

  // ── wyspa szukania: ustawienie (domyślnie jak dotąd; wyłączona = jak w makiecie) ────────────
  const ISLAND_KEY = "dwb-ph-island-v1";
  const islandBox = $("phIslandOpt");
  const islandOn = () => { try { return localStorage.getItem(ISLAND_KEY) !== "0"; } catch (_) { return true; } };
  function syncIsland() { document.body.classList.toggle("ph-no-island", !islandOn()); if (islandBox) islandBox.checked = islandOn(); }
  islandBox?.addEventListener("change", () => {
    try { localStorage.setItem(ISLAND_KEY, islandBox.checked ? "1" : "0"); } catch (_) {}
    syncIsland();
  });
  syncIsland();

  if (typeof window.applyLanguage === "function") {
    const orig2 = window.applyLanguage;
    window.applyLanguage = function applyLanguagePhoneBar(...args) { const r = orig2.apply(this, args); render(); return r; };
  }
  render();
  place();
  return { sync, place, openSheet, tab: () => tab, el: bar };
})();
