// doc-caret-nav.js — poruszanie się po dokumencie w trybie Edycja jak w Wordzie.
//
// Każdy akapit podglądu to osobne pole edycji (contenteditable), więc przeglądarka sama nie
// przechodzi między nimi. Tu dokładamy to, czego brakowało:
//   ↓ w ostatnim wierszu / ↑ w pierwszym  → sąsiedni akapit, w tej samej kolumnie (x),
//   → na końcu / ← na początku            → początek następnego / koniec poprzedniego,
//   Delete na końcu akapitu               → dołącza następny (jak Backspace na początku),
//   Ctrl+Home / Ctrl+End (⌘↑ / ⌘↓ na Macu) → początek / koniec dokumentu,
//   klik/stuknięcie obok tekstu (margines, pusty obszar strony) → kursor w najbliższym akapicie.
// Akapity zablokowane (przypis, obraz, śledzona zmiana — data-lock) są przeskakiwane.

(() => {
  const host = () => docCanvasEl?.querySelector(".docx-preview-host");
  const editable = (p) => p?.isConnected && p.classList.contains("docx-editable-p") && p.isContentEditable;
  const paras = () => collectPreviewParagraphElements(host()).filter(editable);

  // Położenie kursora; null gdy przeglądarka nie podaje prostokąta (np. pusty akapit, koniec
  // węzła) — wtedy decyzja po pozycji w tekście (moveVertical).
  function caretRect() {
    const sel = window.getSelection();
    if (!sel?.rangeCount) return null;
    const r = sel.getRangeAt(0).cloneRange();
    r.collapse(true);
    const rect = r.getClientRects()[0];
    return rect && (rect.width || rect.height) ? rect : null;
  }

  function lineRects(p) {
    const r = document.createRange();
    r.selectNodeContents(p);
    const rects = [...r.getClientRects()].filter((x) => x.height > 0);
    return rects.length ? rects : [p.getBoundingClientRect()];
  }

  function textLength(p) {
    return (p.textContent || "").length;
  }

  function placeAt(p, offset) {
    const len = textLength(p);
    const o = Math.max(0, Math.min(len, offset));
    p.focus({ preventScroll: true });
    if (typeof placeCaret === "function") placeCaret(p, o);
    else {
      const r = document.createRange();
      r.selectNodeContents(p);
      r.collapse(o === 0);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
    }
    p.scrollIntoView({ block: "nearest" });
  }

  // Kursor w akapicie p w poziomie x na wysokości y (pierwszy/ostatni wiersz).
  function placeAtPoint(p, x, y) {
    p.scrollIntoView({ block: "nearest" });
    const rect = p.getBoundingClientRect();
    const cx = Math.max(rect.left + 1, Math.min(rect.right - 1, x));
    let range = null;
    try {
      if (document.caretRangeFromPoint) range = document.caretRangeFromPoint(cx, y);
      else {
        const pos = document.caretPositionFromPoint?.(cx, y);
        if (pos) { range = document.createRange(); range.setStart(pos.offsetNode, pos.offset); }
      }
    } catch (_) { range = null; }
    p.focus({ preventScroll: true });
    if (range && p.contains(range.startContainer)) {
      range.collapse(true);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      return true;
    }
    return false;
  }

  function neighbour(p, dir) {
    const list = paras();
    const i = list.indexOf(p);
    if (i < 0) return null;
    return list[i + dir] || null;
  }

  function moveVertical(p, dir, e) {
    const lines = lineRects(p);
    const c = caretRect();
    const offset = typeof getCaretOffset === "function" ? getCaretOffset(p) : 0;
    const onEdge = c
      ? (dir > 0 ? c.top >= lines[lines.length - 1].top - 2 : c.bottom <= lines[0].bottom + 2)
      : (dir > 0 ? offset >= textLength(p) : offset === 0);
    if (!onEdge) return false; // w środku akapitu — przeglądarka przesunie o wiersz
    const next = neighbour(p, dir);
    if (!next) return false;
    e.preventDefault();
    next.scrollIntoView({ block: "nearest" });
    const nl = lineRects(next);
    const line = dir > 0 ? nl[0] : nl[nl.length - 1];
    const x = c ? c.left : (dir > 0 ? lines[lines.length - 1].right : lines[0].left);
    if (!placeAtPoint(next, x, line.top + line.height / 2)) placeAt(next, dir > 0 ? 0 : textLength(next));
    return true;
  }

  function onKeydown(e) {
    if (readOnlyMode || e.isComposing || e.defaultPrevented) return;
    const p = e.target.closest?.(".docx-editable-p");
    if (!p) return;
    const sel = window.getSelection();
    const collapsed = !!sel?.isCollapsed;
    const mod = e.ctrlKey || e.metaKey;
    const mac = /Mac|iPhone|iPad/.test(navigator.platform || "");

    // początek / koniec dokumentu
    if ((e.ctrlKey && !e.metaKey && (e.key === "Home" || e.key === "End")) || (mac && e.metaKey && (e.key === "ArrowUp" || e.key === "ArrowDown"))) {
      if (e.shiftKey || e.altKey) return;
      const list = paras();
      if (!list.length) return;
      e.preventDefault();
      const toEnd = e.key === "End" || e.key === "ArrowDown";
      const target = toEnd ? list[list.length - 1] : list[0];
      placeAt(target, toEnd ? textLength(target) : 0);
      return;
    }
    if (mod || e.altKey || e.shiftKey || !collapsed) return;

    const offset = typeof getCaretOffset === "function" ? getCaretOffset(p) : -1;
    const len = textLength(p);
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      moveVertical(p, e.key === "ArrowDown" ? 1 : -1, e);
    } else if (e.key === "ArrowRight" && offset >= len) {
      const next = neighbour(p, 1);
      if (next) { e.preventDefault(); placeAt(next, 0); }
    } else if (e.key === "ArrowLeft" && offset === 0) {
      const prev = neighbour(p, -1);
      if (prev) { e.preventDefault(); placeAt(prev, textLength(prev)); }
    } else if (e.key === "Delete" && offset >= len) {
      // jak Backspace na początku następnego akapitu (ta sama ścieżka zapisu i cofania)
      const next = neighbour(p, 1);
      const all = collectPreviewParagraphElements(host());
      if (!next || all.indexOf(next) !== all.indexOf(p) + 1 || typeof handleInlineBackspace !== "function") return;
      e.preventDefault();
      placeAt(next, 0);
      handleInlineBackspace(next, all.indexOf(next), { preventDefault() {}, key: "Backspace" });
    }
  }

  // Klik obok tekstu: najbliższy edytowalny akapit na tej wysokości (albo najbliższy pionowo).
  function onClick(e) {
    if (readOnlyMode || e.button !== 0 || e.detail > 1) return;
    const h = host();
    if (!h || !h.contains(e.target) || e.target.closest(".docx-editable-p, .docx-locked-p, a, button, input, textarea, select")) return;
    if (!window.getSelection()?.isCollapsed) return;
    const list = paras();
    if (!list.length) return;
    let best = null;
    let bestDist = Infinity;
    for (const p of list) {
      const r = p.getBoundingClientRect();
      if (r.bottom < e.clientY - 400 || r.top > e.clientY + 400) continue; // tylko w okolicy
      const dy = e.clientY < r.top ? r.top - e.clientY : e.clientY > r.bottom ? e.clientY - r.bottom : 0;
      if (dy < bestDist) { bestDist = dy; best = p; if (!dy) break; }
    }
    if (!best || bestDist > 60) return;
    const r = best.getBoundingClientRect();
    const y = Math.max(r.top + 2, Math.min(r.bottom - 2, e.clientY));
    if (!placeAtPoint(best, e.clientX, y)) placeAt(best, e.clientX < r.left + r.width / 2 ? 0 : textLength(best));
  }

  docCanvasEl?.addEventListener("keydown", onKeydown);
  docCanvasEl?.addEventListener("click", onClick);
})();
