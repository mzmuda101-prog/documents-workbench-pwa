// doc-selection.js — zaznaczenie przez kilka akapitów w trybie Edycja, jak w Wordzie.
//
// Cały dokument to jedno pole edycji (docx-inline-edit.js, docEditRoot), więc przeglądarka sama
// zaznacza przez akapity: przeciąganie, Shift+klik, Shift+strzałki, dwu-/trzykrotne kliknięcie,
// uchwyty zaznaczenia na telefonie i tablecie. Tu jest to, czego przeglądarce NIE wolno zrobić
// sama — jej domyślne usuwanie przez granicę akapitów przestawiłoby węzły podglądu i rozjechało
// numerację akapitów z plikiem. Każda taka zmiana idzie tą samą drogą co Enter/Backspace:
// podgląd od razu, plik w kolejce (applyInlineStructuralEdit, op „deleteRange”).
//
//   pisanie / Enter / Backspace / Delete / Wytnij / Wklej / Tab w miejsce zaznaczenia,
//   Backspace na początku i Delete na końcu akapitu także z klawiatury ekranowej (Android
//   nie podaje klawisza — tylko „beforeinput”),
//   Ctrl/⌘+A = cała treść dokumentu (bez nagłówków/stopek i przypisów — jak w Wordzie),
//   zaznaczenie przez komórki jednej tabeli: Delete czyści tekst w komórkach (tabela zostaje).
//
// Zasady Worda: zaznaczenie zaczęte na początku akapitu obejmuje cały akapit z jego znacznikiem —
// usunięcie zostawia formatowanie OSTATNIEGO akapitu (np. usunięty nagłówek nie zamienia
// następnego tekstu w nagłówek); zaczęte w środku — wynik ma formatowanie pierwszego.
// Pisanie w miejsce zaznaczenia kończącego się na samym początku następnego akapitu (trzykrotne
// kliknięcie) zostawia znacznik akapitu — zastępuje tylko tekst. Nowy tekst dostaje format
// pierwszego zaznaczonego znaku.

const dwbSel = (() => {
  const host = () => docCanvasEl?.querySelector(".docx-preview-host");
  const root = () => docEditRoot();

  function selRange() {
    const sel = window.getSelection?.();
    if (!sel?.rangeCount) return null;
    const r = sel.getRangeAt(0);
    const rt = root();
    return rt && rt.classList.contains("docx-edit-root") && rt.contains(r.commonAncestorContainer) ? r : null;
  }

  function paraOf(node, paras) {
    const el = node?.nodeType === 1 ? node : node?.parentElement;
    let p = el?.closest?.("p");
    while (p && !paras.has(p)) p = p.parentElement?.closest?.("p") || null;
    return p;
  }

  // Tekst przed pozycją w akapicie (bez znaków-pomocników kursora) — czy pozycja to początek akapitu.
  function atParaStart(p, node, offset) {
    const r = document.createRange();
    r.setStart(p, 0);
    try { r.setEnd(node, offset); } catch (_) { return false; }
    return !r.toString().replace(/﻿/g, "").length && !r.cloneContents().querySelector?.("[data-cm-kind='note'], .ff-field, br, .docx-tab");
  }
  function atParaEnd(p, node, offset) {
    const r = document.createRange();
    try { r.setStart(node, offset); } catch (_) { return false; }
    r.setEnd(p, p.childNodes.length);
    return !r.toString().replace(/﻿/g, "").length && !r.cloneContents().querySelector?.("[data-cm-kind='note'], .ff-field, br, .docx-tab");
  }

  // Analiza zaznaczenia: akapity początku/końca (indeksy jak w pliku) i czy przekracza akapit.
  // forInsert: zaznaczenie kończące się na początku następnego akapitu = do końca poprzedniego.
  function analyze(range, { forInsert = false } = {}) {
    if (!range || range.collapsed) return null;
    const list = collectPreviewParagraphElements(host());
    const set = new Set(list);
    let a = paraOf(range.startContainer, set);
    let b = paraOf(range.endContainer, set);
    if (!a || !b) {
      const hit = list.filter((p) => range.intersectsNode(p));
      a = a || hit[0];
      b = b || hit[hit.length - 1];
    }
    if (!a || !b) return null;
    const start = a.contains(range.startContainer) ? { node: range.startContainer, offset: range.startOffset } : { node: a, offset: 0 };
    let end = b.contains(range.endContainer) ? { node: range.endContainer, offset: range.endOffset } : { node: b, offset: b.childNodes.length };
    let ia = list.indexOf(a);
    let ib = list.indexOf(b);
    const rawCross = a !== b;
    if (forInsert && ib > ia && atParaStart(b, end.node, end.offset)) {
      ib -= 1;
      b = list[ib];
      end = { node: b, offset: b.childNodes.length };
    }
    return { list, a, b, ia, ib, start, end, rawCross, cross: ib > ia };
  }

  // Format pierwszego zaznaczonego znaku (tekst wpisany w miejsce zaznaczenia go dostaje).
  function styleOfFirst(info) {
    const r = document.createRange();
    r.setStart(info.start.node, info.start.offset);
    r.setEnd(info.a, info.a.childNodes.length);
    const walker = document.createTreeWalker(info.a, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (!r.intersectsNode(n) || !/[^﻿]/.test(n.data) || n.parentElement?.closest('[contenteditable="false"]')) continue;
      if (n === info.start.node && info.start.offset >= n.length) continue;
      const chain = [];
      for (let el = n.parentElement; el && el !== info.a; el = el.parentElement) chain.push(el);
      const style = {};
      for (let i = chain.length - 1; i >= 0; i--) accumulateElementStyle(chain[i], style);
      return style;
    }
    return null;
  }

  function setCaret(node, offset) {
    const sel = window.getSelection();
    const r = document.createRange();
    r.setStart(node, offset);
    r.collapse(true);
    sel.removeAllRanges();
    sel.addRange(r);
  }

  // Usunięcie zaznaczonego; zwraca akapit z kursorem albo null (odmowa — z wyjaśnieniem).
  function deleteSpan(info) {
    const { a, b, ia, ib, start, end } = info;
    if (ia === ib) {
      const r = document.createRange();
      r.setStart(start.node, start.offset);
      r.setEnd(end.node, end.offset);
      r.deleteContents();
      setCaret(r.startContainer, r.startOffset);
      onInlineParagraphInput();
      return a;
    }
    // komórki jednej tabeli: tekst znika, tabela zostaje (Word: Delete na zaznaczonych komórkach)
    const ta = a.closest("td, th")?.closest("table");
    const tb = b.closest("td, th")?.closest("table");
    if ((a.dataset.box || "") !== (b.dataset.box || "")) {
      const sameTable = ta && tb && (ta === tb || ta.contains(tb) || tb.contains(ta));
      if (!sameTable) { toast(t("selPartialTable"), "info"); return null; }
      return clearCells(info);
    }
    const lock = a.dataset.lock || b.dataset.lock;
    const wholeFirst = atParaStart(a, start.node, start.offset);
    if (lock && !(wholeFirst && !b.dataset.lock)) { toast(t(lock), "info"); return null; }
    const crossSection = a.closest("section") !== b.closest("section");

    if (wholeFirst) {
      // cały pierwszy akapit (z jego znacznikiem) — zostaje ostatni, z SWOIM formatowaniem
      const head = document.createRange();
      head.setStart(b, 0);
      head.setEnd(end.node, end.offset);
      head.deleteContents();
      const mid = document.createRange();
      mid.setStartBefore(a);
      mid.setEndBefore(b);
      mid.deleteContents();
      pristineParas.delete(b);
      setCaret(b, 0);
      const mergedRuns = extractRunsFromPreviewParagraph(b);
      mirrorBaseline(ia, ib - ia + 1, mergedRuns);
      queue({ op: "deleteRange", from: ia, to: ib, keep: "last", mergedRuns }, crossSection);
      return b;
    }
    const tail = document.createRange();
    tail.setStart(end.node, end.offset);
    tail.setEnd(b, b.childNodes.length);
    const frag = tail.extractContents();
    const headA = document.createRange();
    headA.setStart(start.node, start.offset);
    headA.setEnd(a, a.childNodes.length);
    headA.deleteContents();
    const mid = document.createRange();
    mid.setStartAfter(a);
    mid.setEndBefore(b);
    mid.deleteContents();
    b.remove();
    const join = a.childNodes.length;
    a.appendChild(frag);
    pristineParas.delete(a);
    setCaret(a, join);
    const mergedRuns = extractRunsFromPreviewParagraph(a);
    mirrorBaseline(ia, ib - ia + 1, mergedRuns);
    queue({ op: "deleteRange", from: ia, to: ib, keep: "first", mergedRuns }, crossSection);
    return a;
  }

  function clearCells(info) {
    const { list, ia, ib, start, end } = info;
    for (let i = ia; i <= ib; i++) {
      const p = list[i];
      if (!p || p.dataset.lock) continue;
      const r = document.createRange();
      if (i === ia) r.setStart(start.node, start.offset); else r.setStart(p, 0);
      if (i === ib) r.setEnd(end.node, end.offset); else r.setEnd(p, p.childNodes.length);
      r.deleteContents();
    }
    setCaret(start.node.isConnected ? start.node : info.a, start.node.isConnected ? Math.min(start.offset, start.node.nodeType === 3 ? start.node.length : start.node.childNodes.length) : 0);
    onInlineParagraphInput();
    return info.a;
  }

  // Zmiana w pliku w kolejce; usunięcie przez granicę strony (sekcje podglądu) zostawia w podglądzie
  // pustą kartkę — po zapisie rysujemy dokument od nowa (kursor i przewinięcie wracają).
  function queue(edit, rerender) {
    if (typeof dwbNotes !== "undefined" && !rerender) dwbNotes.watchRefs(); // usunięty odnośnik → przypis znika, numery dalej
    applyInlineStructuralEdit(edit).then(async () => {
      if (!rerender || inlineStructuralPending) return;
      const p = restoreDocCaret();
      const caret = p ? { paraIndex: resolveParaIndex(p), offset: getCaretOffset(p) } : null;
      const top = docViewportEl?.scrollTop || 0;
      await mergeInlineEditsIntoBytes();
      if (caret) pendingInlineCursor = caret;
      await reloadFromBytes(originalFileBytes);
      if (docViewportEl) docViewportEl.scrollTop = top;
    }).catch((err) => log(`Usuwanie zaznaczenia: ${err.message || err}`, "error"));
  }

  // Akapit z samym podziałem strony (lockPageBreakOnly): usunięcie razem z drugą połówką w podglądzie
  // (docx-preview dzieli go na koniec jednej kartki i początek następnej — data-dwb-cont).
  // keep: akapit obok, który zostaje („after” = następny — Backspace, „before” = poprzedni — Delete).
  function removeBreakParagraph(brIndex, keep, side) {
    const list = collectPreviewParagraphElements(host());
    const br = list[brIndex];
    if (!br || !keep) return false;
    const cont = [...host().querySelectorAll("p[data-dwb-cont]")].find((c) => br.compareDocumentPosition(c) & Node.DOCUMENT_POSITION_FOLLOWING);
    const caretAtEnd = side === "before";
    br.remove();
    if (cont && cont.dataset.lock === "lockPageBreakOnly") cont.remove();
    pristineParas.delete(keep);
    if (caretAtEnd) placeCaret(keep, Number.MAX_SAFE_INTEGER); else placeCaret(keep, 0);
    const mergedRuns = extractRunsFromPreviewParagraph(keep);
    const from = side === "before" ? brIndex - 1 : brIndex;
    mirrorBaseline(from, 2, mergedRuns);
    queue({ op: "deleteRange", from, to: from + 1, keep: side === "before" ? "first" : "last", mergedRuns }, true);
    return true;
  }

  function insertAfterDelete(info, text) {
    const style = styleOfFirst(info);
    const p = deleteSpan(info);
    if (!p) return;
    if (text) {
      insertStyledTextAtCaret(text, mergeRunStyles(style || {}, currentTypingStyle()), p);
      onInlineParagraphInput();
    }
  }

  // ── klawiatura ────────────────────────────────────────────────────────────
  function selectAll() {
    const list = collectPreviewParagraphElements(host());
    const rt = root();
    if (!list.length || !rt) return false;
    const r = document.createRange();
    r.setStart(list[0], 0);
    const last = list[list.length - 1];
    r.setEnd(last, last.childNodes.length);
    if (document.activeElement !== rt) rt.focus({ preventScroll: true });
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
    return true;
  }

  function keydown(e) {
    const rt = root();
    if (!rt?.classList.contains("docx-edit-root") || !rt.contains(e.target)) return false;
    const mod = e.ctrlKey || e.metaKey;
    if (mod && !e.altKey && !e.shiftKey && e.code === "KeyA") {
      e.preventDefault();
      return selectAll();
    }
    if (mod && e.key !== "Backspace" && e.key !== "Delete") return false;
    if (!["Enter", "Backspace", "Delete", "Tab"].includes(e.key)) return false;
    const range = selRange();
    if (!range || range.collapsed) return false;
    const info = analyze(range, { forInsert: e.key === "Enter" || e.key === "Tab" });
    if (!info?.rawCross) return false; // w jednym akapicie — jak dotąd (docx-inline-edit.js)
    e.preventDefault();
    if (e.key === "Tab" && info.cross) {
      // kilka punktów listy: Tab / Shift+Tab zmienia poziom każdego (jak w Wordzie)
      const paras = info.list.slice(info.ia, info.ib + 1).filter((p) => !p.dataset.lock);
      if (paras.length && paras.every((p) => isListParagraph(p))) {
        const delta = e.shiftKey ? -1 : 1;
        paras.forEach((p) => {
          const index = info.list.indexOf(p);
          applyDomListLevel({ index, delta });
          applyInlineStructuralEdit({ op: "listLevel", index, delta });
        });
        return true;
      }
      if (e.shiftKey) return true;
    }
    const p = deleteSpan(info);
    if (!p) return true;
    if (e.key === "Enter") handleInlineEnter(p, resolveParaIndex(p), { preventDefault() {}, shiftKey: e.shiftKey });
    else if (e.key === "Tab") { insertTextAtCaret("\t"); onInlineParagraphInput(); }
    return true;
  }

  // ── beforeinput: wszystko, co przeglądarka zrobiłaby przez granicę akapitów ───
  const INSERTS = new Set(["insertText", "insertReplacementText", "insertFromYank", "insertLineBreak", "insertParagraph", "insertTranspose", "insertFromDrop"]);
  function beforeInput(e) {
    const rt = root();
    if (!rt?.classList.contains("docx-edit-root") || e.target !== rt) return false;
    const type = e.inputType || "";
    if (type === "historyUndo" || type === "historyRedo" || type.startsWith("format") || type === "insertCompositionText") return false;
    const range = selRange();
    if (!range) return false;
    if (!range.collapsed) {
      const info = analyze(range, { forInsert: INSERTS.has(type) });
      if (!info?.rawCross) return false; // w jednym akapicie — przeglądarka (i docx-inline-edit.js)
      e.preventDefault();
      if (type === "insertFromDrop") return true; // przeciąganie kilku akapitów — nie (dragstart też blokuje)
      if (type === "insertParagraph" || type === "insertLineBreak") {
        const p = deleteSpan(info);
        if (p) handleInlineEnter(p, resolveParaIndex(p), { preventDefault() {}, shiftKey: type === "insertLineBreak" });
        return true;
      }
      insertAfterDelete(info, type.startsWith("insert") ? (e.data ?? e.dataTransfer?.getData("text/plain") ?? "") : "");
      return true;
    }
    // kursor (bez zaznaczenia): granica akapitu = nasza obsługa, nie scalanie węzłów przez przeglądarkę
    const node = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
    const own = node?.closest?.(".docx-editable-p");
    if (!own || !rt.contains(own)) {
      if (type.startsWith("insert") || type.startsWith("delete")) e.preventDefault(); // poza akapitem tekst by przepadł
      return true;
    }
    // insertLineBreak (Shift+Enter = <br> w akapicie) zostaje przeglądarce — execCommand wysyła je
    // też sam (WebKit), więc przejęcie go tutaj zapętlało Shift+Enter
    const boundary = /^(delete|insertParagraph$)/.test(type);
    if (!boundary) return false; // zwykłe pisanie — bez liczenia akapitów przy każdym znaku
    const list = collectPreviewParagraphElements(host());
    const p = paraOf(range.startContainer, new Set(list));
    if (!p || p !== own) return false;
    const index = list.indexOf(p);
    if (type.startsWith("deleteContentBackward") || type.startsWith("deleteWordBackward") || type.startsWith("deleteSoftLineBackward") || type.startsWith("deleteHardLineBackward")) {
      if (!atParaStart(p, range.startContainer, range.startOffset)) return false;
      e.preventDefault();
      handleInlineBackspace(p, index, { preventDefault() {}, key: "Backspace" });
      return true;
    }
    if (type.startsWith("deleteContentForward") || type.startsWith("deleteWordForward") || type.startsWith("deleteSoftLineForward") || type.startsWith("deleteHardLineForward")) {
      if (!atParaEnd(p, range.startContainer, range.startOffset)) return false;
      e.preventDefault();
      const next = list[index + 1];
      if (next?.dataset.lock === "lockPageBreakOnly") { removeBreakParagraph(index + 1, p, "before"); return true; }
      if (!next || !next.classList.contains("docx-editable-p")) return true;
      placeCaret(next, 0);
      handleInlineBackspace(next, index + 1, { preventDefault() {}, key: "Backspace", fromDelete: true });
      return true;
    }
    if (type === "insertParagraph") {
      // Enter z klawiatury ekranowej bez zdarzenia klawisza (Android) — jak Enter z klawiatury
      e.preventDefault();
      handleInlineEnter(p, index, { preventDefault() {}, shiftKey: false });
      return true;
    }
    return false;
  }

  // Wklejanie / pisanie złożone (IME): najpierw usuwamy zaznaczenie kilku akapitów.
  function collapseForInsert() {
    const range = selRange();
    const info = range && !range.collapsed ? analyze(range, { forInsert: true }) : null;
    if (!info?.rawCross) return true;
    return !!asUndoStep("undoOpTyping", () => deleteSpan(info));
  }

  function onCut(e) {
    const range = selRange();
    const info = range && !range.collapsed ? analyze(range) : null;
    if (!info?.rawCross || readOnlyMode) return;
    e.preventDefault();
    const box = document.createElement("div");
    box.appendChild(range.cloneContents());
    box.querySelectorAll('[contenteditable], .dwb-page-break').forEach((el) => { if (el.classList.contains("dwb-page-break")) el.remove(); else el.removeAttribute("contenteditable"); });
    e.clipboardData?.setData("text/plain", window.getSelection().toString());
    e.clipboardData?.setData("text/html", box.innerHTML);
    asUndoStep("undoOpCut", () => deleteSpan(info));
  }

  docCanvasEl?.addEventListener("cut", onCut);
  docCanvasEl?.addEventListener("compositionstart", () => { if (!readOnlyMode) collapseForInsert(); });
  // przeciągnięcie zaznaczenia kilku akapitów przeniosłoby węzły podglądu — nie pozwalamy
  docCanvasEl?.addEventListener("dragstart", (e) => {
    const range = selRange();
    if (range && !range.collapsed && analyze(range)?.rawCross) e.preventDefault();
  });

  return { keydown, beforeInput, collapseForInsert, selectAll, removeBreakParagraph, analyze: () => analyze(selRange()) };
})();
