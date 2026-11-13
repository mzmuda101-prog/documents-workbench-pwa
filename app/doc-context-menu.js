// doc-context-menu.js — menu kontekstowe dokumentu (jak w Wordzie; wygląd i obsługa jak menu komórki w Sheet).
//
// Otwieranie: prawy klik w dokumencie, klawisz menu / Shift+F10 (przy kursorze), na dotyku przycisk
// „⋯” przy zaznaczeniu — przytrzymania palcem w tekście nie przejmujemy, bo iOS zaznacza nim tekst.
// Treść zależy od miejsca: zaznaczenie (B/I/U, Wytnij, Kopiuj), tabela (wiersze, kolumny), linia
// pozioma (Usuń linię), tryb Czytanie (Kopiuj, Znajdź, Edytuj). Każda pozycja woła tę samą funkcję co
// pasek narzędzi albo skrót — menu nie ma własnej logiki edycji (jeden krok Cofnij jak zawsze).
const dwbDocMenu = (() => {
  let menuEl = null;
  let chipEl = null;
  let returnFocus = null;
  let savedRange = null;
  let lastTouch = 0;
  let swallowClickUntil = 0; // klik kończący stuknięcie w „⋯” nie może trafić w pozycję menu pod palcem

  const MOD = /Mac|iPhone|iPad|iPod/.test(navigator.platform || "") || /Mac OS X/.test(navigator.userAgent || "") ? "⌘" : "Ctrl";
  const kbd = (s) => s.replace(/Mod/g, MOD);

  function host() { return docCanvasEl?.querySelector(".docx-preview-host"); }
  function selRange() {
    const sel = window.getSelection?.();
    if (!sel?.rangeCount) return null;
    const r = sel.getRangeAt(0);
    return docCanvasEl?.contains(r.commonAncestorContainer) ? r : null;
  }
  function paraOf(node) {
    const el = node?.nodeType === 1 ? node : node?.parentElement;
    return el?.closest?.(".docx-preview-host p") || null;
  }
  // linia pozioma = pusty akapit z dolną krawędzią (op "hrule")
  function isRule(p) {
    return !!p && !p.textContent.replace(/﻿/g, "") && !p.querySelector("img, svg, canvas") && getComputedStyle(p).borderBottomStyle !== "none";
  }
  function restoreRange() {
    if (!savedRange) return;
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(savedRange);
  }

  // ── akcje ──────────────────────────────────────────────────────────────────
  function execClip(cmd) {
    restoreRange();
    let ok = false;
    try { ok = document.execCommand(cmd); } catch (_) { ok = false; }
    if (!ok) toast(t("ctxClipBlocked", { key: kbd(cmd === "cut" ? "Mod+X" : "Mod+C") }), "info");
  }
  // Wklejanie przez to samo zdarzenie „paste” co Ctrl/⌘+V (onDocPaste: style, linki, pola…)
  async function pasteFromClipboard(plain) {
    restoreRange();
    const target = paraOf(window.getSelection()?.anchorNode) || host();
    try {
      const dt = new DataTransfer();
      let html = "";
      let text = "";
      if (!plain && navigator.clipboard?.read) {
        const items = await navigator.clipboard.read();
        for (const it of items) {
          if (!html && it.types.includes("text/html")) html = await (await it.getType("text/html")).text();
          if (!text && it.types.includes("text/plain")) text = await (await it.getType("text/plain")).text();
        }
      } else if (navigator.clipboard?.readText) {
        text = await navigator.clipboard.readText();
      }
      if (!html && !text) throw new Error("pusty schowek");
      restoreRange();
      if (html) dt.setData("text/html", html);
      dt.setData("text/plain", text);
      target.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
    } catch (_) {
      // przeglądarka nie dała dostępu do schowka (bez zgody / Safari bez gestu) — skrót działa zawsze
      toast(t("ctxPasteBlocked", { key: kbd("Mod+V") }), "info");
    }
  }
  function findSelection() {
    const text = String(savedRange || "").trim();
    if (!text || !searchQueryEl) return;
    searchQueryEl.value = text.slice(0, 120);
    searchQueryEl.dispatchEvent(new Event("input", { bubbles: true }));
    searchQueryEl.focus();
  }
  async function deleteRule(p) {
    const ps = collectPreviewParagraphElements(host());
    const i = ps.indexOf(p);
    if (i < 0) return;
    const index = resolveParaIndex(p);
    const next = ps[i + 1];
    const prev = ps[i - 1];
    // linia + następny akapit z treścią → zostaje następny ze swoim formatowaniem (mergeParagraphInXml);
    // za linią pusta linijka albo koniec dokumentu → linia wsiąka w poprzedni akapit (pusty doklejany
    // akapit znika, poprzedni zostaje ze swoim formatowaniem)
    const usable = (q) => q && !q.dataset.lock && (q.dataset.box || "") === (p.dataset.box || "");
    if (usable(next) && next.textContent.replace(/\uFEFF/g, "")) {
      await applyDocumentEdit({ op: "mergeParagraph", index: index + 1, mergedRuns: extractRunsFromPreviewParagraph(next) });
    } else if (usable(prev)) {
      await applyDocumentEdit({ op: "mergeParagraph", index, mergedRuns: extractRunsFromPreviewParagraph(prev) });
    } else {
      await applyDocumentEdit({ op: "paraFormat", indices: [index], clearRule: true });
    }
  }
  // Pusta strona za tabelą na końcu dokumentu: Word MUSI mieć akapit za tabelą, a gdy tabela kończy
  // stronę, ten akapit spada na nową kartkę. Sposób z porad Microsoftu — akapit 1 pt bez odstępów.
  // Dawniej robił to Backspace (zgłoszenie 2026-10-11); od 2026-11-09 Backspace jak w Wordzie (nic),
  // a ten sposób jest tutaj, w menu.
  function blankAfterTable(p) {
    if (!p || p.classList.contains("dwb-tiny-para") || p.textContent.replace(/\uFEFF/g, "") || p.querySelector("img, svg, canvas, [contenteditable=false]")) return null;
    let prev = p.previousElementSibling;
    while (prev?.matches?.(".dwb-page-gap, .dwb-page-gap-spacer, .dwb-page-gap-band, .dwb-page-break")) prev = prev.previousElementSibling;
    if (prev?.tagName !== "TABLE") return null;
    const ps = collectPreviewParagraphElements(host());
    return ps.indexOf(p) === ps.length - 1 ? p : null;
  }
  async function removeBlankPage(p) {
    const index = resolveParaIndex(p);
    if (index < 0) return;
    await applyDocumentEdit({ op: "paraFormat", indices: [index], spacing: { before: 0, after: 0, line: 20, lineRule: "exact" } });
    toast(t("blankPageRemoved"), "info");
  }
  function inlineFormat(cmd) { restoreRange(); if (typeof execInlineFormat === "function") execInlineFormat(cmd); }

  // ── pozycje menu wg miejsca ──────────────────────────────────────────────────
  function items(ctx) {
    const out = [];
    const sep = () => { if (out.length && out[out.length - 1] !== "sep") out.push("sep"); };
    const add = (id, key, run, opts = {}) => out.push({ id, label: t(key), run, ...opts });
    if (ctx.read) {
      if (ctx.hasSel) add("copy", "ctxCopy", () => execClip("copy"), { kbd: "Mod+C" });
      add("selAll", "ctxSelectAll", () => { const h = host(); if (h) window.getSelection().selectAllChildren(h); }, { kbd: "Mod+A" });
      if (ctx.hasSel) add("find", "ctxFind", findSelection, { kbd: "Mod+F" });
      sep();
      add("edit", "ctxEdit", () => appFrame.setReadOnly(false), { kbd: "Mod+Alt+E", primary: true });
      return out;
    }
    if (ctx.rule) {
      add("delRule", "ctxDeleteRule", () => deleteRule(ctx.rule), { danger: true });
      sep();
    } else if (ctx.textRule) { // linia pod tekstem — znika sama linia, tekst zostaje
      add("delRule", "ctxDeleteRule", () => composeUi.clearRuleAt(ctx.textRule), { danger: true });
      sep();
    }
    if (ctx.ocrMarks) {
      add("ocrMarks", "ctxClearOcrMarks", () => clearOcrMarks());
      sep();
    }
    if (ctx.blankPage) {
      add("blankPage", "ctxRemoveBlankPage", () => removeBlankPage(ctx.blankPage));
      sep();
    }
    // ── recenzja na górze menu (jak w Wordzie): zmiana w tym akapicie, komentarz pod kursorem ──
    if (ctx.revisions.length) {
      const ids = ctx.revisions.map((c) => c.id);
      add("revAccept", ctx.revisions.length > 1 ? "ctxAcceptChanges" : "ctxAcceptChange", () => dwbMargin.revisions("accept", ids), { primary: true });
      add("revReject", ctx.revisions.length > 1 ? "ctxRejectChanges" : "ctxRejectChange", () => dwbMargin.revisions("reject", ids));
      sep();
    }
    if (ctx.commentId) {
      const c = composeUi.commentById(ctx.commentId);
      add("cmReply", "ctxCommentReply", () => { restoreRange(); composeUi.commentAction(ctx.commentId, "reply"); });
      add("cmResolve", c?.done ? "ctxCommentReopen" : "ctxCommentResolve", () => composeUi.commentAction(ctx.commentId, "resolve"));
      add("cmDelete", "ctxCommentDelete", () => composeUi.commentAction(ctx.commentId, "delete"), { danger: true });
      sep();
    }
    if (ctx.hasSel) {
      add("cut", "ctxCut", () => execClip("cut"), { kbd: "Mod+X" });
      add("copy", "ctxCopy", () => execClip("copy"), { kbd: "Mod+C" });
    }
    add("paste", "ctxPaste", () => pasteFromClipboard(false), { kbd: "Mod+V" });
    add("pastePlain", "ctxPastePlain", () => pasteFromClipboard(true)); // Ctrl/⌘+Shift+V = malarz formatów, nie podpowiadamy
    sep();
    add("link", "ctxLink", () => { restoreRange(); composeUi.openLinkForm(); }, { kbd: "Mod+K" });
    add("comment", "ctxComment", () => { restoreRange(); composeUi.openCommentForm(); }, { kbd: "Mod+Alt+M" });
    // propozycja zmiany zaznaczonego tekstu (bez włączania trybu Sugerowanie) — obok komentarza
    if (ctx.hasSel && typeof dwbTrack !== "undefined") add("suggest", "ctxSuggest", () => { const r = savedRange?.cloneRange(); if (r) dwbTrack.suggest(r); });
    add("note", "ctxFootnote", () => { restoreRange(); composeUi.insertNote("footnote"); });
    if (ctx.cell) {
      sep();
      add("rowAbove", "tableRowAbove", () => { restoreRange(); composeUi.tableAction("rowAbove"); });
      add("rowBelow", "tableRowBelow", () => { restoreRange(); composeUi.tableAction("rowBelow"); });
      add("colLeft", "tableColLeft", () => { restoreRange(); composeUi.tableAction("colLeft"); });
      add("colRight", "tableColRight", () => { restoreRange(); composeUi.tableAction("colRight"); });
      add("delRow", "tableDelRow", () => { restoreRange(); composeUi.tableAction("delRow"); }, { danger: true });
      add("delCol", "tableDelCol", () => { restoreRange(); composeUi.tableAction("delCol"); }, { danger: true });
      add("delTable", "tableDelTable", () => { restoreRange(); composeUi.tableAction("delTable"); }, { danger: true });
    }
    sep();
    add("h1", "ctxHeading1", () => { restoreRange(); composeUi.applyStyle("h1"); });
    add("h2", "ctxHeading2", () => { restoreRange(); composeUi.applyStyle("h2"); });
    add("normal", "ctxNormal", () => { restoreRange(); composeUi.applyStyle("normal"); });
    add("bullets", "ctxBullets", () => { restoreRange(); composeUi.applyList("bullet"); });
    add("numbers", "ctxNumbers", () => { restoreRange(); composeUi.applyList("number"); });
    sep();
    if (!ctx.cell) add("table", "ctxInsertTable", () => { restoreRange(); composeUi.insertTable(3, 3); });
    if (typeof dwbTextbox !== "undefined") add("textbox", "ctxInsertTextbox", () => { restoreRange(); dwbTextbox.insert(); });
    add("hrule", "ctxInsertHrule", () => { restoreRange(); composeUi.insertHrule(); });
    add("pageBreak", "ctxPageBreak", () => { restoreRange(); composeUi.insertPageBreak(); }, { kbd: "Mod+Enter" });
    if (ctx.hasSel) {
      sep();
      add("find", "ctxFind", findSelection, { kbd: "Mod+F" });
      if (typeof dwbRewrite !== "undefined") add("rewrite", "ctxRewrite", () => { restoreRange(); dwbRewrite.open(); });
    }
    return out;
  }

  // ── rysowanie, pozycja, klawiatura ───────────────────────────────────────────
  function ensureMenu() {
    if (menuEl) return menuEl;
    menuEl = document.createElement("div");
    menuEl.className = "doc-menu hidden";
    menuEl.setAttribute("role", "menu");
    menuEl.setAttribute("aria-label", t("ctxMenuAria"));
    menuEl.addEventListener("mousedown", (e) => e.preventDefault()); // zaznaczenie w dokumencie zostaje
    menuEl.addEventListener("keydown", onMenuKey);
    document.body.appendChild(menuEl);
    document.addEventListener("pointerdown", (e) => { if (menuEl && !menuEl.classList.contains("hidden") && !menuEl.contains(e.target)) close(false); }, true);
    window.addEventListener("blur", () => close(false));
    window.addEventListener("click", (e) => {
      if (Date.now() > swallowClickUntil || !menuEl.contains(e.target)) return;
      swallowClickUntil = 0;
      e.preventDefault();
      e.stopPropagation();
    }, true);
    // menu otwarte myszą: fokus zostaje w dokumencie, więc Esc łapiemy na całej stronie
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape" || !menuEl || menuEl.classList.contains("hidden") || menuEl.contains(document.activeElement)) return;
      e.preventDefault();
      e.stopPropagation();
      close();
    }, true);
    docViewportEl?.addEventListener("scroll", () => close(false), { passive: true });
    return menuEl;
  }
  function commentAt(r) {
    if (!r || typeof composeUi === "undefined" || !composeUi.commentRanges) return null;
    // kursor w komentarzu albo zaznaczenie zachodzące na komentowany tekst (Safari zaznacza słowo
    // pod prawym kliknięciem — początek wypada wtedy dokładnie na granicy komentarza)
    // (zakresy zachodzą na siebie albo się stykają: początek r ≤ koniec komentarza i koniec r ≥ jego początek)
    const overlaps = (cr) => cr.comparePoint(r.startContainer, r.startOffset) <= 0 && cr.comparePoint(r.endContainer, r.endOffset) >= 0;
    const hit = composeUi.commentRanges().find(({ id, range }) => { try { return overlaps(range) && !composeUi.commentById(id)?.parentId; } catch (_) { return false; } });
    return hit ? hit.id : null;
  }
  function context() {
    const r = selRange();
    const p = paraOf(r?.startContainer);
    return {
      read: !!readOnlyMode,
      hasSel: !!r && !r.collapsed && !!String(r).trim(),
      rule: isRule(p) ? p : null,
      textRule: composeUi.hasTextRule?.(p) ? p : null,
      blankPage: blankAfterTable(p),
      ocrMarks: typeof ocrMarksPresent !== "undefined" && ocrMarksPresent,
      cell: !!p?.closest("td, th"),
      // recenzja: komentarz pod kursorem i śledzone zmiany w akapicie (review-margin.js)
      commentId: commentAt(r),
      revisions: typeof dwbMargin !== "undefined" && p ? dwbMargin.revisionsAt(p) : [],
    };
  }
  function open(point, { viaKeyboard = false } = {}) {
    if (!originalFileBytes) return false;
    const r = selRange();
    savedRange = r ? r.cloneRange() : null;
    const ctx = context();
    const list = items(ctx);
    const el = ensureMenu();
    const frag = document.createDocumentFragment();
    if (ctx.hasSel && !ctx.read) { // mini pasek jak w Wordzie: najczęstsze formatowanie od razu
      const row = document.createElement("div");
      row.className = "doc-menu-row";
      [["bold", "B", "ctxBold", "Mod+B"], ["italic", "I", "ctxItalic", "Mod+I"], ["underline", "U", "ctxUnderline", "Mod+U"]].forEach(([cmd, txt, key, k]) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = `doc-menu-fmt dmf-${cmd}`;
        b.setAttribute("role", "menuitem");
        b.textContent = txt;
        b.setAttribute("aria-label", `${t(key)} (${kbd(k)})`);
        b.dataset.hint = "";
        b.dataset.hintPl = `${I18N.pl[key]} (${kbd(k)})`;
        b.dataset.hintEn = `${I18N.en[key]} (${kbd(k)})`;
        b.addEventListener("click", () => { close(false); inlineFormat(cmd); });
        row.appendChild(b);
      });
      frag.appendChild(row);
    }
    list.forEach((it) => {
      if (it === "sep") { const s = document.createElement("div"); s.className = "doc-menu-sep"; s.setAttribute("role", "separator"); frag.appendChild(s); return; }
      const b = document.createElement("button");
      b.type = "button";
      b.className = `doc-menu-item dmi-${it.id}${it.danger ? " is-danger" : ""}${it.primary ? " is-primary" : ""}`;
      b.setAttribute("role", "menuitem");
      const lab = document.createElement("span");
      lab.textContent = it.label;
      b.appendChild(lab);
      if (it.kbd) { const k = document.createElement("kbd"); k.textContent = kbd(it.kbd); b.appendChild(k); }
      b.addEventListener("click", () => { close(false); Promise.resolve(it.run()).catch((err) => log?.(`Menu: ${err.message || err}`, "error")); });
      frag.appendChild(b);
    });
    el.replaceChildren(frag);
    el.classList.remove("hidden");
    hideChip();
    // pozycja: przy punkcie; z klawiatury przy kursorze; zawsze w oknie (przewijanie w menu, gdy za wysokie)
    const vv = window.visualViewport;
    const vw = vv ? vv.width : window.innerWidth;
    const vh = vv ? vv.height : window.innerHeight;
    el.style.maxHeight = `${Math.max(160, vh - 16)}px`;
    let pt = point;
    if (!pt) {
      const rr = savedRange?.getBoundingClientRect();
      pt = rr && (rr.width || rr.height) ? { x: rr.left, y: rr.bottom + 4 } : { x: vw / 2 - 110, y: vh / 3 };
    }
    let x = pt.x;
    let y = pt.y;
    const mw = el.offsetWidth;
    const mh = el.offsetHeight;
    if (x + mw > vw - 8) x = Math.max(8, vw - mw - 8);
    if (y + mh > vh - 8) y = Math.max(8, vh - mh - 8);
    el.style.left = `${Math.round(x + (vv ? vv.offsetLeft : 0))}px`;
    el.style.top = `${Math.round(y + (vv ? vv.offsetTop : 0))}px`;
    returnFocus = viaKeyboard ? document.activeElement : null;
    if (viaKeyboard) el.querySelector("button")?.focus();
    return true;
  }
  function close(restore = true) {
    if (!menuEl || menuEl.classList.contains("hidden")) return;
    menuEl.classList.add("hidden");
    menuEl.replaceChildren();
    if (restore) restoreRange();
    if (restore && returnFocus && document.contains(returnFocus)) returnFocus.focus({ preventScroll: true });
    returnFocus = null;
  }
  function onMenuKey(e) {
    const btns = Array.from(menuEl.querySelectorAll("button"));
    const i = btns.indexOf(document.activeElement);
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); close(); return; }
    if (e.key === "Tab") { e.preventDefault(); close(); return; }
    let next = null;
    if (e.key === "ArrowDown" || e.key === "ArrowRight") next = (i + 1) % btns.length;
    if (e.key === "ArrowUp" || e.key === "ArrowLeft") next = (i - 1 + btns.length) % btns.length;
    if (e.key === "Home") next = 0;
    if (e.key === "End") next = btns.length - 1;
    if (next === null) return;
    e.preventDefault();
    e.stopPropagation();
    btns[next].focus();
  }

  // ── dotyk: „⋯” przy zaznaczeniu ─────────────────────────────────────────────
  function hideChip() { if (chipEl) chipEl.hidden = true; }
  function syncChip() {
    if (Date.now() - lastTouch > 4000) { hideChip(); return; } // tylko po dotyku — myszą jest prawy klik
    const r = selRange();
    if (!r || r.collapsed || !String(r).trim() || (menuEl && !menuEl.classList.contains("hidden"))) { hideChip(); return; }
    if (!chipEl) {
      chipEl = document.createElement("button");
      chipEl.type = "button";
      chipEl.className = "doc-menu-chip";
      chipEl.textContent = "⋯";
      chipEl.setAttribute("aria-label", t("ctxMore"));
      // otwieramy już przy dotknięciu: na iOS stuknięcie najpierw czyści zaznaczenie, „⋯” się chowa
      // i klik trafiał w próżnię; klik zostaje dla klawiatury (Enter na przycisku)
      const fromChip = (e) => {
        e.preventDefault();
        const b = chipEl.getBoundingClientRect();
        if (open({ x: b.left, y: b.bottom + 6 }) && e.type === "pointerdown") swallowClickUntil = Date.now() + 700;
      };
      chipEl.addEventListener("pointerdown", fromChip);
      chipEl.addEventListener("click", (e) => { if (e.detail === 0) fromChip(e); });
      document.body.appendChild(chipEl);
    }
    const rects = r.getClientRects();
    const last = rects[rects.length - 1] || r.getBoundingClientRect();
    const vv = window.visualViewport;
    // pod końcem zaznaczenia — systemowy dymek iOS (Kopiuj/Wklej) stoi nad nim, więc się nie zasłaniają
    chipEl.style.left = `${Math.round(Math.min(last.right + 6, (vv ? vv.width : innerWidth) - 46) + (vv ? vv.offsetLeft : 0))}px`;
    chipEl.style.top = `${Math.round(last.bottom + 10 + (vv ? vv.offsetTop : 0))}px`;
    chipEl.hidden = false;
  }

  function init() {
    if (!docCanvasEl) return;
    docCanvasEl.addEventListener("contextmenu", (e) => {
      if (!e.target.closest?.(".docx-preview-host") || e.defaultPrevented) return;
      if (e.target.closest(".dwb-tb, a[href], img, .ff-field, .dwb-tb-editor")) return; // swoje karty/menu albo systemowe (link, obraz)
      if (Date.now() - lastTouch < 800) { e.preventDefault(); return; } // Android: przytrzymanie = zaznaczanie tekstu, nie menu
      e.preventDefault();
      open({ x: e.clientX, y: e.clientY });
    });
    docCanvasEl.addEventListener("keydown", (e) => {
      if (!(e.key === "ContextMenu" || (e.key === "F10" && e.shiftKey))) return;
      if (!e.target.closest?.(".docx-preview-host, .docx-edit-root")) return;
      e.preventDefault();
      e.stopPropagation();
      open(null, { viaKeyboard: true });
    });
    docCanvasEl.addEventListener("touchstart", () => { lastTouch = Date.now(); }, { passive: true, capture: true });
    document.addEventListener("selectionchange", () => requestAnimationFrame(syncChip));
    docViewportEl?.addEventListener("scroll", hideChip, { passive: true });
  }

  return { init, open, close, _isOpen: () => !!menuEl && !menuEl.classList.contains("hidden") };
})();

document.addEventListener("DOMContentLoaded", () => dwbDocMenu.init());
