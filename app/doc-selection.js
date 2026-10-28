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
    // <br data-dwb-ph> = pusta linijka „dla oka” na końcu akapitu, nie znak — za nią już koniec
    return !r.toString().replace(/﻿/g, "").length && !r.cloneContents().querySelector?.("[data-cm-kind='note'], .ff-field, br:not([data-dwb-ph]), .docx-tab");
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

  // Dlaczego zaznaczenia kilku akapitów nie da się usunąć (klucz komunikatu) albo null.
  function spanRefusal(info) {
    const { a, b, ia, ib, start } = info;
    if (ia === ib) return null;
    if ((a.dataset.box || "") !== (b.dataset.box || "")) {
      const ta = a.closest("td, th")?.closest("table");
      const tb = b.closest("td, th")?.closest("table");
      return ta && tb && (ta === tb || ta.contains(tb) || tb.contains(ta)) ? null : "selPartialTable";
    }
    const lock = a.dataset.lock || b.dataset.lock;
    if (lock && !(atParaStart(a, start.node, start.offset) && !b.dataset.lock)) return lock;
    return null;
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
    const why = spanRefusal(info);
    if (why) { toast(t(why), "info"); return null; }
    // komórki jednej tabeli: tekst znika, tabela zostaje (Word: Delete na zaznaczonych komórkach)
    if ((a.dataset.box || "") !== (b.dataset.box || "")) return clearCells(info);
    const wholeFirst = atParaStart(a, start.node, start.offset);
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
      if (typeof dwbTextbox !== "undefined") info.list.slice(ia, ib).forEach((q) => dwbTextbox.dropAnchoredTo?.(q));
      mid.deleteContents();
      pristineParas.delete(b);
      setCaret(b, 0);
      const mergedRuns = extractRunsFromPreviewParagraph(b);
      mirrorBaseline(ia, ib - ia + 1, mergedRuns);
      // początek akapitu b był zaznaczony — pola tekstowe zakotwiczone w nim znikają (Word: kotwica
      // w zaznaczeniu = kształt usunięty; Ctrl/⌘+A + pisanie zostawiało pole z ostatniego akapitu)
      if (typeof dwbTextbox !== "undefined") dwbTextbox.dropAnchoredTo?.(b);
      queue({ op: "deleteRange", from: ia, to: ib, keep: "last", mergedRuns, dropBoxes: true }, crossSection);
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
    // akapity znikające w całości (i b — jego początek był zaznaczony): ich pola tekstowe też
    // (plik usuwa je razem z akapitami; podgląd zostawiał je „w powietrzu”)
    if (typeof dwbTextbox !== "undefined") info.list.slice(ia + 1, ib + 1).forEach((q) => dwbTextbox.dropAnchoredTo?.(q));
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
      await mergeInlineEditsIntoBytes();
      if (caret) pendingInlineCursor = caret;
      await reloadFromBytes(originalFileBytes);
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

  // Pusty akapit za tabelą (Backspace, docx-inline-edit.js): znika, następny zostaje bez zmian,
  // kursor na końcu ostatniej komórki tabeli — jak w Wordzie.
  function removeEmptyParagraph(index, p, next, caretAt) {
    p.remove();
    pristineParas.delete(next);
    placeCaret(caretAt, Number.MAX_SAFE_INTEGER);
    const mergedRuns = extractRunsFromPreviewParagraph(next);
    mirrorBaseline(index, 2, mergedRuns);
    queue({ op: "deleteRange", from: index, to: index + 1, keep: "last", mergedRuns }, true);
    return true;
  }

  function insertAfterDelete(info, text) {
    const style = styleOfFirst(info);
    const p = deleteSpan(info);
    if (!p) return;
    if (text) {
      insertStyledTextAtCaret(text, mergeRunStyles(style || {}, currentTypingStyle()), p, { turnOff: currentTypingStyle() });
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
      if (!info?.rawCross) {
        // Backspace/Delete na zaznaczonym słowie: inteligentne wycinanie (Word) — bez podwójnej
        // spacji i bez spacji przed kropką (docx-inline-edit.js smartCutRange)
        const wider = /^delete(Content|Word|SoftLine|HardLine)/.test(type) && typeof smartCutRange === "function" ? smartCutRange(range) : null;
        const wi = wider && analyze(wider);
        if (!wi || wi.rawCross) return false; // w jednym akapicie — przeglądarka (i docx-inline-edit.js)
        e.preventDefault();
        asUndoStep("undoOpTyping", () => deleteSpan(wi));
        return true;
      }
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
    // odmowa przed zapisem kroku Cofnij — inaczej w historii zostawał pusty krok
    const why = spanRefusal(info);
    if (why) { toast(t(why), "info"); return false; }
    return !!asUndoStep("undoOpTyping", () => deleteSpan(info));
  }

  function onCut(e) {
    const range = selRange();
    const info = range && !range.collapsed ? analyze(range) : null;
    if (!info || readOnlyMode) return;
    // Znaczniki komentarza są niewidocznymi wyspami, więc zwykły HTML schowka kopiował tylko
    // zaznaczony tekst. Gdy wycinany jest CAŁY komentowany fragment, rozszerzamy wycinek o jego
    // start, koniec i odwołanie. Wklejanie w tym samym dokumencie odtworzy je z własnego formatu
    // schowka — komentarz przechodzi wraz z tekstem zamiast znikać przy następnym zapisie.
    const cutRange = expandRangeToWholeComments(range);
    const cutInfo = analyze(cutRange);
    if (!cutInfo) return;
    e.preventDefault();
    const box = clipboardBox(cutRange);
    // Połówki komentarzy wyciętych tylko częściowo (np. początek bez końca) zostają w dokumencie,
    // w miejscu wycięcia — jak w Wordzie zakres komentarza się skraca. Dawniej szły do schowka:
    // wklejony wyżej początek łączył się z końcem w starym miejscu i podświetlał kilka akapitów.
    const kept = orphanMarks(cutRange);
    const clipId = rememberClip(cutRange, box, true);
    // Natywny Selection zachowuje separator między akapitami; Range#toString go pomija.
    e.clipboardData?.setData("text/plain", window.getSelection().toString());
    e.clipboardData?.setData("text/html", box.innerHTML);
    if (clipId) e.clipboardData?.setData("application/x-dwb-clip", clipId);
    if (box.querySelector(COMMENT_MARKS)) {
      // Typ działa tylko w naszej aplikacji; zewnętrzne programy dalej dostają zwykły HTML/tekst.
      e.clipboardData?.setData("application/x-dwb-comment-fragment", fragmentHtml(box));
      // treść komentarza czeka w pliku na wklejenie (docx-inline-edit.js, dwbClipComments)
      box.querySelectorAll("[data-cm-kind='ref']").forEach((ref) => dwbClipComments.add(ref.dataset.cmId));
    }
    // przypis wycięty razem z odnośnikiem czeka na wklejenie (doc-notes.js pruneOrphanNotesInZip)
    noteIdsIn(box).forEach((key) => dwbClipNotes.add(key));
    // inteligentne wycinanie (Word): całe słowo znika razem z jedną spacją obok (smartCutRange);
    // do schowka idzie samo zaznaczenie
    const smartRange = typeof smartCutRange === "function" ? smartCutRange(cutRange) : null;
    const delInfo = (smartRange && analyze(smartRange)) || cutInfo;
    asUndoStep("undoOpCut", () => {
      const p = deleteSpan(delInfo);
      if (!p || !kept.length) return p;
      const at = selRange();
      if (!at) return p;
      const frag = document.createDocumentFragment();
      kept.forEach((mark) => frag.appendChild(mark));
      at.insertNode(frag); // zwinięty zakres po insertNode obejmuje wstawione znaczniki
      setCaret(at.endContainer, at.endOffset);
      onInlineParagraphInput();
      return p;
    });
    if (box.querySelector(COMMENT_MARKS) || kept.length) composeUi?.paintCommentHighlights?.();
  }

  // Kopiuj: zwykły tekst/HTML dla innych programów + nasz rekord schowka (dwbClip) — wklejenie
  // w tym dokumencie zachowuje formatowanie źródła jak w Wordzie (punktory, wcięcia, styl akapitu,
  // krój, rozmiar, kolor). Dawniej wklejka szła filtrem „jak ze strony WWW”: kilka punktów listy
  // wklejało się jako jeden sklejony akapit bez punktorów i kolorów (pomiar 2026-10-12).
  function onCopy(e) {
    const range = selRange();
    if (!range || range.collapsed) return;
    const copyRange = expandRangeToWholeComments(range);
    const box = clipboardBox(copyRange);
    const clipId = rememberClip(copyRange, box, false);
    if (!clipId && !box.querySelector(COMMENT_MARKS)) return;
    e.preventDefault();
    e.clipboardData?.setData("text/plain", window.getSelection().toString());
    e.clipboardData?.setData("text/html", box.innerHTML);
    if (clipId) e.clipboardData?.setData("application/x-dwb-clip", clipId);
    if (box.querySelector(COMMENT_MARKS)) e.clipboardData?.setData("application/x-dwb-comment-fragment", fragmentHtml(box));
  }

  // Rekord schowka: akapity zaznaczenia z formatowaniem. Jeden akapit bez jego znacznika (kawałek
  // tekstu) — fragment HTML z formatowaniem fragmentów; akapity ze znacznikiem (cały akapit,
  // kilka akapitów) — fragmenty + właściwości akapitów z pliku (w:pPr), wklejane jak w Wordzie
  // (docx-inline-edit.js pasteClipParagraphs, op „pasteParas”). Tylko w obrębie jednego „pojemnika”
  // (treść główna albo jedna komórka) — przez komórki tabeli zostaje zwykłe kopiowanie.
  function rememberClip(range, box, cut) {
    try {
      let info = analyze(range);
      if (!info) return null;
      let { b, ib, end } = info;
      const { list, start, a, ia } = info;
      let lastMark = false;
      if (ib > ia && atParaStart(b, end.node, end.offset)) {
        ib -= 1;
        b = list[ib];
        end = { node: b, offset: b.childNodes.length };
        lastMark = true;
      }
      // (dopiero po odcięciu: cały akapit tuż nad tabelą kończy zaznaczenie na początku komórki)
      if ((a.dataset.box || "") !== (b.dataset.box || "")) return null;
      const paras = [];
      for (let i = ia; i <= ib; i++) {
        const p = list[i];
        if (p.dataset.lock || p.dataset.noteKey) return null;
        const r = document.createRange();
        if (i === ia) r.setStart(start.node, start.offset); else r.setStart(p, 0);
        if (i === ib) r.setEnd(end.node, end.offset); else r.setEnd(p, p.childNodes.length);
        const tmp = document.createElement("p");
        tmp.appendChild(cloneWithFormat(r));
        tmp.querySelectorAll(".dwb-page-break, .dwb-page-gap, .dwb-page-gap-spacer, .dwb-page-gap-band").forEach((x) => x.remove());
        paras.push({ index: i, runs: extractRunsFromPreviewParagraph(tmp), mark: i < ib || lastMark });
      }
      // inteligentne wklejanie (Word): skopiowano całe słowa? — wtedy przy wklejeniu obok słowa spacja
      const WORD = /[\p{L}\p{N}]/u;
      const before = document.createRange();
      before.setStart(a, 0);
      before.setEnd(start.node, start.offset);
      const after = document.createRange();
      after.setStart(end.node, end.offset);
      after.setEnd(b, b.childNodes.length);
      const sb = !WORD.test(before.toString().replace(/\uFEFF/g, "").slice(-1));
      const eb = !WORD.test(after.toString().replace(/\uFEFF/g, "").slice(0, 1));
      const rec = { doc: dwbDocSerial, cut, paras, sb, eb };
      if (paras.length === 1 && !paras[0].mark) rec.html = fragmentHtml(box);
      else {
        // w:pPr z pliku w stanie z chwili kopiowania: po Enterach jeszcze w kolejce, przed
        // usunięciem tego zaznaczenia (Wytnij) — nasz then rejestruje się przed nim
        rec.pPrs = inlineStructuralChain.then(async () => {
          const bytes = originalFileBytes;
          const zip = await JSZip.loadAsync(bytes);
          const doc = new DOMParser().parseFromString(await zip.file("word/document.xml").async("string"), "application/xml");
          const xps = collectParagraphElements(doc.documentElement, "all");
          return paras.map((q) => {
            const pPr = Array.from(xps[q.index]?.childNodes || []).find((n) => n.nodeType === 1 && n.localName === "pPr");
            return pPr ? new XMLSerializer().serializeToString(pPr) : "";
          });
        });
        rec.pPrs.catch(() => {});
      }
      const id = `clip${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
      dwbClip.set(id, rec);
      if (dwbClip.size > 8) dwbClip.delete(dwbClip.keys().next().value);
      return id;
    } catch (err) {
      log(`Schowek: ${err.message || err}`, "error");
      return null;
    }
  }

  const COMMENT_MARKS = "[data-cm-kind='start'], [data-cm-kind='end'], [data-cm-kind='ref']";

  // HTML do schowka: kopia zaznaczenia bez łamań stron i bez połówek komentarzy. Komentarz
  // przechodzi tylko w komplecie (początek + koniec + odwołanie); sam początek albo sam koniec
  // po wklejeniu sparowałby się z drugą połówką w zupełnie innym miejscu dokumentu.
  // Kopia zaznaczenia razem z formatowaniem: gdy zaznaczenie leży w JEDNYM fragmencie (np. wewnątrz
  // pogrubionego), cloneContents daje sam tekst — opakowujemy go płytkimi kopiami przodków do
  // poziomu akapitu (Word: skopiowany kawałek pogrubionego tekstu wkleja się pogrubiony).
  function cloneWithFormat(range) {
    let frag = range.cloneContents();
    let anc = range.commonAncestorContainer;
    if (anc.nodeType !== 1) anc = anc.parentNode;
    const p = anc?.closest?.(".docx-editable-p");
    while (p && anc && anc !== p && p.contains(anc)) {
      const wrap = anc.cloneNode(false);
      wrap.removeAttribute?.("contenteditable");
      wrap.appendChild(frag);
      frag = document.createDocumentFragment();
      frag.appendChild(wrap);
      anc = anc.parentNode;
    }
    return frag;
  }

  function clipboardBox(range) {
    const box = document.createElement("div");
    box.appendChild(cloneWithFormat(range));
    box.querySelectorAll('[contenteditable], .dwb-page-break').forEach((el) => { if (el.classList.contains("dwb-page-break")) el.remove(); else el.removeAttribute("contenteditable"); });
    const whole = wholeCommentIds(box.querySelectorAll(COMMENT_MARKS));
    box.querySelectorAll(COMMENT_MARKS).forEach((mark) => { if (!whole.has(mark.dataset.cmId)) mark.remove(); });
    return box;
  }

  // Własny fragment schowka: znaczniki niosą swój XML. Klucze data-cm są ważne tylko do
  // następnego przerysowania dokumentu — po nim ten sam klucz wskazywał znacznik INNEGO
  // komentarza (wklejone zdanie dostawało cudzy komentarz, plik z powtórzonym numerem).
  function fragmentHtml(box) {
    const copy = box.cloneNode(true);
    copy.querySelectorAll("[data-cm]").forEach((el) => {
      const xml = docIslandXml.get(el.dataset.cm);
      if (xml) el.dataset.cmXml = xml;
    });
    return copy.innerHTML;
  }

  // „footnote:N” / „endnote:N” odnośników przypisów we fragmencie (wyspy z XML-em)
  function noteIdsIn(box) {
    const out = [];
    box.querySelectorAll("[data-cm]").forEach((el) => {
      const xml = docIslandXml.get(el.dataset.cm) || el.dataset.cmXml || "";
      for (const m of xml.matchAll(/<w:(footnote|endnote)Reference\b[^>]*\bw:id="(\d+)"/g)) out.push(`${m[1]}:${m[2]}`);
    });
    return out;
  }

  function wholeCommentIds(marks) {
    const kinds = new Map();
    marks.forEach((mark) => {
      if (!kinds.has(mark.dataset.cmId)) kinds.set(mark.dataset.cmId, new Set());
      kinds.get(mark.dataset.cmId).add(mark.dataset.cmKind);
    });
    return new Set([...kinds].filter(([, k]) => k.has("start") && k.has("end") && k.has("ref")).map(([id]) => id));
  }

  // Znaczniki komentarzy w zakresie, których komentarz NIE jest w nim cały (oryginalne węzły).
  function orphanMarks(range) {
    const h = host();
    if (!h) return [];
    const inside = [...h.querySelectorAll(COMMENT_MARKS)].filter((mark) => range.intersectsNode(mark)
      && range.comparePoint(mark, 0) === 0 && range.comparePoint(mark, mark.childNodes.length) === 0);
    const whole = wholeCommentIds(inside);
    return inside.filter((mark) => !whole.has(mark.dataset.cmId));
  }

  // Między dwoma miejscami są tylko niewidoczne znaczniki (bez znaków, obrazów, pól, tabulatorów).
  function onlyMarksBetween(aNode, aOffset, bNode, bOffset) {
    const r = document.createRange();
    try { r.setStart(aNode, aOffset); r.setEnd(bNode, bOffset); } catch (_) { return false; }
    return !r.toString().replace(/﻿/g, "").length
      && !r.cloneContents().querySelector?.("img, svg, canvas, [data-cm-kind='note'], [data-cm-kind='obj'], .ff-field, br:not([data-dwb-ph]), .docx-tab");
  }

  function expandRangeToWholeComments(range) {
    const h = host();
    if (!h) return range;
    const expanded = range.cloneRange();
    h.querySelectorAll("span[data-cm-kind='start']").forEach((start) => {
      const id = start.dataset.cmId;
      const end = h.querySelector(`span[data-cm-kind="end"][data-cm-id="${CSS.escape(id)}"]`);
      const ref = h.querySelector(`span[data-cm-kind="ref"][data-cm-id="${CSS.escape(id)}"]`);
      if (!end || !ref) return;
      const content = document.createRange();
      content.setStartAfter(start);
      content.setEndBefore(end);
      // Tylko pełny zakres komentarza: wycięcie części słowa zachowuje standardowe zachowanie
      // Worda, zamiast nieoczekiwanie rozszerzać zaznaczenie o resztę komentarza.
      // Zaznaczenie kończące się na ostatniej literze komentowanego tekstu obejmuje go całego,
      // chociaż w DOM koniec zaznaczenia wypada PRZED niewidocznym znacznikiem końca komentarza
      // (Word zapisuje koniec i odwołanie za tekstem — np. komentarz do końca akapitu).
      if (range.compareBoundaryPoints(Range.START_TO_START, content) > 0
        && !onlyMarksBetween(content.startContainer, content.startOffset, range.startContainer, range.startOffset)) return;
      if (range.compareBoundaryPoints(Range.END_TO_END, content) < 0
        && !onlyMarksBetween(range.endContainer, range.endOffset, content.endContainer, content.endOffset)) return;
      const markers = document.createRange();
      markers.setStartBefore(start);
      markers.setEndAfter(ref);
      if (markers.compareBoundaryPoints(Range.START_TO_START, expanded) < 0) expanded.setStartBefore(start);
      if (markers.compareBoundaryPoints(Range.END_TO_END, expanded) > 0) expanded.setEndAfter(ref);
    });
    return expanded;
  }

  docCanvasEl?.addEventListener("cut", onCut);
  docCanvasEl?.addEventListener("copy", onCopy);
  docCanvasEl?.addEventListener("compositionstart", () => { if (!readOnlyMode) collapseForInsert(); });
  // przeciągnięcie zaznaczenia kilku akapitów przeniosłoby węzły podglądu — nie pozwalamy
  docCanvasEl?.addEventListener("dragstart", (e) => {
    const range = selRange();
    if (range && !range.collapsed && analyze(range)?.rawCross) e.preventDefault();
  });

  return { keydown, beforeInput, collapseForInsert, selectAll, removeBreakParagraph, removeEmptyParagraph, analyze: () => analyze(selRange()) };
})();
