// shortcuts.js — JEDEN rejestr skrótów klawiszowych + okno „Skróty klawiszowe” (Ctrl/⌘+/) +
// akcje pliku w menu ⋯ (kontrakt UX §4–5, audyt D5–D7, D25).
//
// Z rejestru powstają: lista w oknie skrótów, napisy przy pozycjach menu ([data-shortcut="id"])
// i przycisk w panelu „Skróty i informacje”. Zapis zależy od platformy: na Macu/iPadzie ⌘ ⌥ ⇧ ⌃,
// gdzie indziej Ctrl / Alt / Shift. Dawniej lista żyła w trzech miejscach i rozjeżdżała się z kodem
// (brakowało w niej m.in. Ctrl/⌘+K, przypisów, malarza formatów).
//
// Klawisze w rejestrze: „Mod” = Ctrl (Windows) / ⌘ (Mac), „Ctrl” = zawsze Ctrl (np. ⌃H na Macu,
// bo ⌘H chowa okno), „/” między wariantami = „albo”. Obsługa samych klawiszy siedzi tam, gdzie
// działa funkcja (keyboard.js, compose-ui.js, docx-inline-edit.js…); tu tylko nowe: Ctrl/⌘+O,
// Ctrl/⌘+Shift+S, Ctrl/⌘+H, Ctrl/⌘+/. Test shortcuts-ux pilnuje, że wpis ma działający skrót.

const DWB_SHORTCUTS = [
  { group: "file", id: "open", keys: ["Mod+O"], pl: "Otwórz plik…", en: "Open a file…" },
  { group: "file", id: "save", keys: ["Mod+S"], pl: "Zapisz", en: "Save" },
  { group: "file", id: "saveAs", keys: ["Mod+Shift+S"], pl: "Zapisz jako…", en: "Save as…" },
  { group: "file", id: "newDoc", keys: ["Mod+Alt+N"], pl: "Nowy dokument…", en: "New document…" },
  { group: "file", id: "print", keys: ["Mod+P"], pl: "Podgląd wydruku / PDF", en: "Print preview / PDF" },

  { group: "find", id: "find", keys: ["Mod+F"], pl: "Szukaj w dokumencie", en: "Find in the document" },
  { group: "find", id: "replace", keys: ["Mod+H"], mac: ["Ctrl+H"], pl: "Znajdź i zamień", en: "Find and replace" },
  { group: "find", id: "findNext", keys: ["Enter", "Mod+G", "F3"], pl: "Następne trafienie", en: "Next match" },
  { group: "find", id: "findPrev", keys: ["Shift+Enter", "Mod+Shift+G", "Shift+F3"], pl: "Poprzednie trafienie", en: "Previous match" },

  { group: "edit", id: "mode", keys: ["Mod+Alt+E"], pl: "Czytanie ⇄ Edycja", en: "Reading ⇄ Editing" },
  { group: "edit", id: "undo", keys: ["Mod+Z"], pl: "Cofnij", en: "Undo" },
  { group: "edit", id: "redo", keys: ["Mod+Shift+Z", "Ctrl+Y"], pl: "Ponów", en: "Redo" },
  { group: "edit", id: "selectAll", keys: ["Mod+A"], pl: "Zaznacz cały dokument", en: "Select the whole document" },
  { group: "edit", id: "newPara", keys: ["Enter"], pl: "Nowy akapit", en: "New paragraph" },
  { group: "edit", id: "lineBreak", keys: ["Shift+Enter"], pl: "Nowa linia w akapicie", en: "Line break" },
  { group: "edit", id: "pageBreak", keys: ["Mod+Enter"], pl: "Podział strony", en: "Page break" },
  { group: "edit", id: "listIndent", keys: ["Tab", "Shift+Tab"], pl: "Lista: poziom niżej / wyżej", en: "List: demote / promote" },
  { group: "edit", id: "track", keys: ["Mod+Shift+E"], pl: "Śledzenie zmian", en: "Track changes" },

  { group: "format", id: "bold", keys: ["Mod+B"], pl: "Pogrubienie", en: "Bold" },
  { group: "format", id: "italic", keys: ["Mod+I"], pl: "Kursywa", en: "Italic" },
  { group: "format", id: "underline", keys: ["Mod+U"], pl: "Podkreślenie", en: "Underline" },
  { group: "format", id: "alignKeys", keys: ["Mod+L", "Mod+E", "Mod+R", "Mod+J"], pl: "Wyrównaj do lewej / do środka / do prawej / wyjustuj", en: "Align left / center / right / justify" },
  { group: "format", id: "headingKeys", keys: ["Mod+Alt+1", "Mod+Alt+2", "Mod+Alt+3"], pl: "Nagłówek 1 / 2 / 3 (kursor w tekście)", en: "Heading 1 / 2 / 3 (cursor in text)" },
  { group: "format", id: "sizeUp", keys: ["Mod+]", "Mod+Shift+>"], pl: "Większa czcionka", en: "Bigger font" },
  { group: "format", id: "sizeDown", keys: ["Mod+[", "Mod+Shift+<"], pl: "Mniejsza czcionka", en: "Smaller font" },
  { group: "format", id: "superscript", keys: ["Mod+Shift+="], pl: "Indeks górny", en: "Superscript" },
  { group: "format", id: "subscript", keys: ["Mod+="], pl: "Indeks dolny", en: "Subscript" },
  { group: "format", id: "painterCopy", keys: ["Mod+Shift+C"], pl: "Malarz formatów: kopiuj format", en: "Format painter: copy format" },
  { group: "format", id: "painterPaste", keys: ["Mod+Shift+V"], pl: "Malarz formatów: wklej format", en: "Format painter: paste format" },

  { group: "insert", id: "link", keys: ["Mod+K"], pl: "Link", en: "Link" },
  { group: "insert", id: "comment", keys: ["Mod+Alt+M"], pl: "Komentarz", en: "Comment" },
  { group: "insert", id: "footnote", keys: ["Mod+Alt+F"], pl: "Przypis dolny (kursor w tekście)", en: "Footnote (cursor in text)" },
  { group: "insert", id: "endnote", keys: ["Mod+Alt+D"], pl: "Przypis końcowy", en: "Endnote" },
  { group: "insert", id: "bookmark", keys: ["Mod+Shift+F5"], pl: "Zakładka", en: "Bookmark" },

  { group: "view", id: "focusMode", keys: ["Mod+Alt+F"], pl: "Tryb skupienia (kursor poza tekstem)", en: "Focus mode (cursor outside text)" },
  { group: "view", id: "regions", keys: ["Mod+Alt+1", "Mod+Alt+2", "Mod+Alt+3"], pl: "Panel / pasek / dokument (kursor poza tekstem)", en: "Panel / toolbar / document (cursor outside text)" },
  { group: "view", id: "regionCycle", keys: ["F6", "Shift+F6"], pl: "Obszary po kolei", en: "Cycle screen areas" },
  { group: "view", id: "nextHeading", keys: ["Mod+Tab"], pl: "Następny nagłówek (w tekście)", en: "Next heading (in text)" },
  { group: "view", id: "back", keys: ["Alt+←"], pl: "Wróć po kliknięciu linku", en: "Back after following a link" },
  { group: "view", id: "goToPage", keys: ["Mod+Alt+G"], pl: "Przejdź do strony…", en: "Go to page…" },
  { group: "view", id: "shortcuts", keys: ["Mod+/"], pl: "Ta lista skrótów", en: "This shortcut list" },
  { group: "view", id: "bugReport", keys: ["Mod+Alt+Shift+B"], pl: "Zgłoś błąd (zapisz dziennik)", en: "Report a bug (save log)" },
  { group: "view", id: "escape", keys: ["Esc"], pl: "Krok wstecz (pisanie → panel → podświetlenia → skupienie)", en: "Step back (typing → panel → highlights → focus)" },
];

const DWB_SHORTCUT_GROUPS = {
  file: { pl: "Plik", en: "File" },
  find: { pl: "Szukanie", en: "Find" },
  edit: { pl: "Edycja", en: "Editing" },
  format: { pl: "Formatowanie", en: "Formatting" },
  insert: { pl: "Wstawianie", en: "Insert" },
  view: { pl: "Widok i nawigacja", en: "View and navigation" },
};

const dwbShortcuts = (() => {
  const IS_APPLE = /Mac|iPhone|iPad/.test(navigator.platform || "") || /Mac OS X/.test(navigator.userAgent || "");
  const MAC_SYM = { Mod: "⌘", Alt: "⌥", Shift: "⇧", Ctrl: "⌃" };

  // „Mod+Shift+S” → ["⌘", "⇧", "S"] (Mac) albo ["Ctrl", "Shift", "S"]
  function keyParts(combo) {
    const parts = combo === "Mod++" ? ["Mod", "+"] : combo.split("+");
    return parts.map((p) => (IS_APPLE ? MAC_SYM[p] || p : p === "Mod" ? "Ctrl" : p));
  }
  function comboText(combo) {
    const parts = keyParts(combo);
    return IS_APPLE ? parts.join("") : parts.join("+");
  }
  function entryKeys(e) { return (IS_APPLE && e.mac) || e.keys; }
  function label(id) {
    const e = DWB_SHORTCUTS.find((x) => x.id === id);
    return e ? entryKeys(e).map(comboText).join(" / ") : "";
  }

  // napisy przy pozycjach menu i podpowiedziach: <span data-shortcut="saveAs">
  function syncLabels(root = document) {
    root.querySelectorAll("[data-shortcut]").forEach((el) => {
      const txt = label(el.dataset.shortcut);
      if (txt) el.textContent = txt;
    });
  }

  // ── okno „Skróty klawiszowe” ────────────────────────────────────────────
  let dlg = null;
  function build() {
    const lang = typeof currentLang !== "undefined" && currentLang === "en" ? "en" : "pl";
    const d = document.createElement("dialog");
    d.className = "app-dialog shortcuts-dialog";
    d.id = "shortcutsDialog";
    d.setAttribute("aria-labelledby", "shortcutsTitle");
    const head = document.createElement("div");
    head.className = "shortcuts-head";
    const h = document.createElement("h2");
    h.id = "shortcutsTitle";
    h.className = "app-dialog-title";
    h.textContent = t("shortcutsTitle");
    const close = document.createElement("button");
    close.type = "button";
    close.className = "newdoc-close";
    close.textContent = "×";
    close.setAttribute("aria-label", t("toastClose"));
    close.addEventListener("click", () => d.close());
    head.append(h, close);
    const body = document.createElement("div");
    body.className = "shortcuts-body";
    for (const [g, names] of Object.entries(DWB_SHORTCUT_GROUPS)) {
      const sec = document.createElement("section");
      sec.className = "shortcuts-group";
      const gh = document.createElement("h3");
      gh.textContent = names[lang];
      const dl = document.createElement("dl");
      for (const e of DWB_SHORTCUTS.filter((x) => x.group === g)) {
        const row = document.createElement("div");
        row.className = "shortcuts-row";
        row.dataset.shortcutId = e.id;
        const dt = document.createElement("dt");
        dt.textContent = e[lang];
        const dd = document.createElement("dd");
        entryKeys(e).forEach((combo, i) => {
          if (i) { const or = document.createElement("span"); or.className = "shortcuts-or"; or.textContent = "/"; dd.append(or); }
          const k = document.createElement("kbd");
          k.textContent = comboText(combo);
          dd.append(k);
        });
        row.append(dt, dd);
        dl.append(row);
      }
      sec.append(gh, dl);
      body.append(sec);
    }
    const foot = document.createElement("p");
    foot.className = "shortcuts-foot";
    foot.textContent = t("shortcutsFoot");
    d.append(head, body, foot);
    d.addEventListener("click", (e) => { if (e.target === d) d.close(); });
    d.addEventListener("close", () => { d.remove(); if (dlg === d) dlg = null; });
    return d;
  }
  function open() {
    if (typeof appFrame !== "undefined") appFrame.setMenuOpen?.(false);
    if (dlg?.open) return;
    dlg = build();
    // fokus na ✕ tylko z klawiatury (na dotyku bez skoku)
    if (lastInputModality === "keyboard") dlg.querySelector(".newdoc-close")?.setAttribute("autofocus", "");
    else { dlg.tabIndex = -1; dlg.setAttribute("autofocus", ""); }
    document.body.append(dlg);
    if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open", "");
  }

  // ── akcje pliku (menu ⋯ i skróty) ───────────────────────────────────────
  function openToolPanel(panelId, focusEl) {
    if (typeof setSidebarOpen === "function") setSidebarOpen(true);
    const panel = document.getElementById(panelId);
    if (!panel) return;
    panel.open = true;
    requestAnimationFrame(() => panel.scrollIntoView({ block: "start", behavior: "smooth" }));
    if (focusEl && lastInputModality === "keyboard") setTimeout(() => { focusEl.focus(); focusEl.select?.(); }, 120);
  }
  // Ctrl/⌘+H jak w Wordzie: panel „Znajdź i zamień”, kursor w „Zamień na”. Zaznaczony w dokumencie
  // krótki tekst trafia do szukania (gdy pole jest puste).
  function openReplace() {
    if (!originalFileBytes) return;
    const sel = window.getSelection?.();
    const picked = sel && !sel.isCollapsed && docCanvasEl?.contains(sel.anchorNode) ? String(sel).trim() : "";
    if (picked && picked.length <= 120 && !picked.includes("\n") && searchQueryEl && !searchQueryEl.value.trim()) {
      searchQueryEl.value = picked;
      searchQueryEl.dispatchEvent(new Event("input", { bubbles: true }));
    }
    const go = () => openToolPanel("panel-search", document.getElementById("frReplace"));
    if (typeof ensureLazyFeature === "function") ensureLazyFeature("find-replace").then(go, go); else go();
  }
  function openExport() {
    if (!originalFileBytes) { toast(t("noFileToSave"), "warning"); return; }
    openToolPanel("panel-export");
  }
  function takePhoto() { document.getElementById("photoDocBtn")?.click(); }

  async function copyVersion() {
    const txt = `Documents Workbench ${APP_BUILD_VERSION}`;
    try { await navigator.clipboard.writeText(txt); toast(t("versionCopied", { v: APP_BUILD_VERSION }), "success"); }
    catch (_) { toast(txt, "info"); }
  }
  function syncVersion() {
    const el = document.getElementById("appMenuBuild");
    if (el && typeof APP_BUILD_VERSION !== "undefined") el.textContent = t("versionLabel", { v: APP_BUILD_VERSION });
  }
  function syncNeedsDoc() {
    document.querySelectorAll("#appMenu .needs-doc").forEach((b) => { b.disabled = !originalFileBytes; });
  }

  function onKey(e) {
    if (e.defaultPrevented || e.isComposing) return;
    const mod = e.ctrlKey || e.metaKey;
    if (!mod) return;
    if (document.querySelector("dialog.app-dialog[open]")) return; // okno na wierzchu — jego klawisze
    if (!e.altKey && !e.shiftKey && e.code === "Slash") { e.preventDefault(); open(); return; }
    if (!e.altKey && !e.shiftKey && e.code === "KeyO") { e.preventDefault(); if (typeof openFilePicker === "function") openFilePicker(); return; }
    if (!e.altKey && e.shiftKey && e.code === "KeyS") {
      if (!originalFileBytes) return;
      e.preventDefault();
      if (typeof saveDocumentAs === "function") saveDocumentAs();
      return;
    }
    // ⌘H na Macu chowa okno (system) — tam działa ⌃H, jak w Wordzie dla Maca
    if (!e.altKey && !e.shiftKey && e.code === "KeyH" && (e.ctrlKey || !IS_APPLE)) {
      if (!originalFileBytes) return;
      e.preventDefault();
      openReplace();
    }
  }

  function init() {
    syncLabels();
    syncVersion();
    syncNeedsDoc();
    document.addEventListener("keydown", onKey);
    const on = (id, fn) => document.getElementById(id)?.addEventListener("click", fn);
    on("shortcutsMenuItem", () => setTimeout(open, 60));
    on("shortcutsHelpBtn", open);
    on("openMenuItem", () => { if (typeof openFilePicker === "function") openFilePicker(); });
    on("photoMenuItem", takePhoto);
    on("emptyPhotoBtn", takePhoto);
    on("saveAsMenuItem", () => { if (typeof saveDocumentAs === "function") saveDocumentAs(); });
    on("exportMenuItem", openExport);
    on("guideMenuItem", () => { if (typeof loadSampleDocument === "function") loadSampleDocument("przewodnik"); });
    on("appMenuBuild", copyVersion);
    // menu ⋯ otwiera się → aktualne „wymaga dokumentu”; język → nowe napisy
    document.getElementById("appMenuBtn")?.addEventListener("click", () => { syncNeedsDoc(); syncVersion(); }, true);
    if (typeof applyLanguage === "function") {
      const orig = window.applyLanguage;
      window.applyLanguage = function applyLanguageShortcuts(...args) {
        const r = orig.apply(this, args);
        syncVersion();
        syncLabels();
        return r;
      };
    }
  }
  init();

  return { open, label, comboText, syncLabels, openReplace, openToolPanel, isApple: IS_APPLE };
})();
