// doc-caret-nav.js — poruszanie się i zaznaczanie klawiaturą w trybie Edycja, jak w Wordzie.
//
// Cały dokument to jedno pole edycji, ale ruch kursora przeglądarki gubił się na rzeczach, których
// Word nie ma w tekście (pomiar 2026-10-11 na pliku z komentarzami: ↑ stał, ↓ krążył, przeskoki
// o 2–3 wiersze, → potrzebował kilku naciśnięć przy komentarzu). Dlatego sami liczymy:
//   ↑ / ↓ (z Shift)            → wiersz wyżej/niżej w tej samej kolumnie (pamiętanej przez kolejne
//                                 naciśnięcia), w tabeli komórka pod/nad, z geometrii wierszy,
//   ← / → (z Shift), Alt (Mac) / Ctrl → znak / słowo; niewidoczne znaczniki komentarzy nie zjadają
//                                 naciśnięcia; na brzegu akapitu — sąsiedni (zablokowane pomijamy),
//   Home / End (z Shift)       → początek / koniec wiersza (Chrome na Macu tylko przewijał),
//   PageUp / PageDown (z Shift) → kursor o ekran (nie samo przewinięcie),
//   Alt+↑/↓ (Mac) / Ctrl+↑/↓   → początek akapitu / poprzedni / następny,
//   Ctrl+Home / Ctrl+End (⌘↑ / ⌘↓ na Macu), z Shift → początek / koniec dokumentu,
//   Delete na końcu akapitu    → dołącza następny (jak Backspace na początku),
//   klik/stuknięcie obok tekstu (margines, pusty obszar strony) → kursor w najbliższym akapicie.
// Akapity zablokowane (przypis, obraz, śledzona zmiana — data-lock) są przeskakiwane.
// Pomiary: scripts/nav-keys-playwright.js.

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

  // Kursor na końcu akapitu liczonym jak zapis: za nim nie ma tekstu ani niczego, co w pliku jest
  // znakiem (łamanie wiersza <br> z pliku, wyspa). Sam textContent <br> nie widzi — Delete przed
  // łamaniem doklejał następny akapit ZA łamaniem, a kolejny Delete zjadał jego pierwszą literę
  // (lustro Backspace przed „Wstęp”, 2026-10-09).
  function atModelEnd(p) {
    const sel = getSelection();
    if (!sel?.rangeCount || !p.contains(sel.getRangeAt(0).endContainer)) return false;
    const r = document.createRange();
    r.selectNodeContents(p);
    r.setStart(sel.getRangeAt(0).endContainer, sel.getRangeAt(0).endOffset);
    const frag = r.cloneContents();
    return !frag.textContent.replace(/\uFEFF/g, "") && !frag.querySelector("br:not([data-dwb-ph]), img, svg, canvas, [contenteditable=false]");
  }

  function placeAt(p, offset) {
    const len = textLength(p);
    const o = Math.max(0, Math.min(len, offset));
    focusDocParagraph(p);
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
    focusDocParagraph(p);
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

  // ── ↑ / ↓ (i z Shift) — liczone z geometrii wierszy, jak w Wordzie ─────────────────────
  // Przeglądarka sama robiła to źle w trzech miejscach (pomiar 2026-10-11, plik z komentarzami):
  //   • przy niewidocznych znacznikach komentarza (span contenteditable=false) ↑ stał w miejscu,
  //     a ↓ krążył między dwiema pozycjami,
  //   • na granicy zawinięcia wiersza prostokąt kursora wypada na KOŃCU poprzedniego wiersza —
  //     ↑ z początku 2. wiersza uznawał, że to 1. wiersz, i skakał o dwa wiersze,
  //   • w tabeli ↓ z ostatniego wiersza komórki szedł do komórki obok (kolejność w DOM).
  // Tu: bieżący wiersz z pamięci ostatniego ruchu (albo ze znaku ZA kursorem), cel = najbliższy
  // wiersz niżej/wyżej w kolumnie kursora (szerokość akapitu/komórki), kolumna pamiętana przez
  // kolejne ↑/↓ (jak w Wordzie: krótki wiersz nie gubi pozycji w poziomie).
  const MARK_SEL = '[contenteditable="false"]';
  // położenie po naszym ostatnim ruchu w pionie / kliknięciu / End: wysokość względem akapitu
  // (przewinięcie między naciśnięciami jej nie psuje), kolumna względem dokumentu
  let nav = null; // { node, offset, p, dy, gx }
  const hostLeft = () => host()?.getBoundingClientRect().left || 0;
  function remember(p, node, offset, y, x) {
    nav = { node, offset, p, dy: y - p.getBoundingClientRect().top, gx: x - hostLeft() };
  }

  // Wyspa = element, przez który kursor przechodzi w całości: znacznik/obraz/pole (contenteditable
  // false), obraz, łamanie wiersza. Ma pole na ekranie (obraz, numer przypisu) albo nie (znacznik).
  const atomic = (n) => n.nodeType === 1 && (n.matches(MARK_SEL) || /^(IMG|SVG|CANVAS|BR)$/i.test(n.tagName));
  const boxOf = (el) => {
    const img = el.matches("img, svg, canvas") ? el : el.querySelector?.("img, svg, canvas");
    const c = (img || el).getBoundingClientRect();
    return c.width > 1 && c.height > 1 ? c : null;
  };

  // Wiersze akapitu: prostokąty tekstu i wysp, łączone, gdy nachodzą na siebie w pionie (obraz
  // w wierszu z tekstem to jeden wiersz, nie dwa).
  function lineBoxes(p) {
    const r = document.createRange();
    r.selectNodeContents(p);
    const rects = [...r.getClientRects()].filter((x) => x.height > 1);
    p.querySelectorAll(`img, svg, canvas, ${MARK_SEL}`).forEach((el) => { const c = boxOf(el); if (c) rects.push(c); });
    if (!rects.length) {
      const b = p.getBoundingClientRect();
      return [{ top: b.top, bottom: b.top + Math.min(b.height || 16, 24), left: b.left, right: b.left }];
    }
    const lines = [];
    rects.sort((a, b) => a.top - b.top).forEach((x) => {
      const l = lines.find((l) => Math.min(l.bottom, x.bottom) - Math.max(l.top, x.top) >= 0.4 * Math.min(l.bottom - l.top, x.height));
      if (l) {
        l.top = Math.min(l.top, x.top); l.bottom = Math.max(l.bottom, x.bottom);
        if (x.width) { l.left = Math.min(l.left, x.left); l.right = Math.max(l.right, x.right); }
      } else lines.push({ top: x.top, bottom: x.bottom, left: x.width ? x.left : Infinity, right: x.width ? x.right : -Infinity });
    });
    const b = p.getBoundingClientRect();
    lines.forEach((l) => { if (!Number.isFinite(l.left)) { l.left = b.left; l.right = b.left; } });
    return lines.sort((x, y) => x.top - y.top);
  }

  const stepOut = (n, root, dir) => {
    for (; n && n !== root; n = n.parentNode) {
      const sib = dir > 0 ? n.nextSibling : n.previousSibling;
      if (sib) return sib;
    }
    return null;
  };
  // Prostokąt pierwszej rzeczy zajmującej miejsce od pozycji w stronę dir (znak albo wyspa
  // z polem — obraz, numer przypisu); niewidoczne znaczniki i łamania wiersza pomijamy.
  function charRectFrom(p, node, offset, dir) {
    const r = document.createRange();
    let cur;
    let idx = null;
    if (node.nodeType === 3) { cur = node; idx = dir > 0 ? offset : offset - 1; }
    else cur = node.childNodes[dir > 0 ? offset : offset - 1] || stepOut(node, p, dir);
    for (let guard = 0; cur && guard < 2000; guard++) {
      if (cur.nodeType === 3) {
        for (let i = idx ?? (dir > 0 ? 0 : cur.length - 1); i >= 0 && i < cur.length; i += dir) {
          const ch = cur.data[i];
          if (ch === "\uFEFF" || ch === "\n") continue;
          r.setStart(cur, i); r.setEnd(cur, i + 1);
          const c = [...r.getClientRects()].find((x) => x.height > 1);
          if (c) return c;
        }
      } else if (cur.nodeType === 1 && atomic(cur)) {
        if (cur.tagName !== "BR") { const c = boxOf(cur); if (c) return c; }
      } else if (cur.nodeType === 1 && (dir > 0 ? cur.firstChild : cur.lastChild)) {
        cur = dir > 0 ? cur.firstChild : cur.lastChild;
        idx = null;
        continue;
      }
      cur = stepOut(cur, p, dir);
      idx = null;
    }
    return null;
  }

  // Wysokość (y) i poziom (x) końca zaznaczenia (focus) — „w dół” (znak za nim), jak rysuje
  // przeglądarka po ruchu strzałką; po End/kliknięciu — z pamięci (nav).
  function focusPoint(p) {
    const sel = getSelection();
    const node = sel.focusNode;
    const offset = sel.focusOffset;
    if (nav && nav.node === node && nav.offset === offset && nav.p === p) {
      return { y: p.getBoundingClientRect().top + nav.dy, x: nav.gx + hostLeft() };
    }
    // znak ZA kursorem: na granicy zawinięcia przeglądarka po strzałce rysuje kursor na
    // początku nowego wiersza (prostokąt samego kursora wypadałby na końcu poprzedniego)
    const after = charRectFrom(p, node, offset, 1);
    if (after) return { y: (after.top + after.bottom) / 2, x: after.left };
    const before = charRectFrom(p, node, offset, -1);
    if (before) return { y: (before.top + before.bottom) / 2, x: before.right };
    const r = document.createRange();
    r.setStart(node, offset);
    const c = [...r.getClientRects()].find((x) => x.height > 0);
    if (c) return { y: (c.top + c.bottom) / 2, x: c.left };
    const l = lineBoxes(p)[0];
    return { y: (l.top + l.bottom) / 2, x: l.left };
  }

  // przewinięcie do kursora (nie do całego akapitu — długi akapit nie „skacze”)
  function revealFocus() {
    const sel = getSelection();
    const fm = focusModel();
    if (!fm.p || !sel?.focusNode) return;
    const c = charRectFrom(fm.p, sel.focusNode, sel.focusOffset, -1) || charRectFrom(fm.p, sel.focusNode, sel.focusOffset, 1);
    if (c) revealLine(fm.p, c); else fm.p.scrollIntoView({ block: "nearest" });
  }

  function scroller(el) {
    for (let n = el?.parentElement; n; n = n.parentElement) {
      const cs = getComputedStyle(n);
      if (/(auto|scroll)/.test(cs.overflowY) && n.scrollHeight > n.clientHeight + 1) return n;
    }
    return document.scrollingElement;
  }
  // wiersz ma być widoczny (z zapasem) — caretRangeFromPoint działa tylko na ekranie
  function revealLine(p, line) {
    const sc = scroller(p);
    const box = sc === document.scrollingElement ? { top: 0, bottom: window.innerHeight } : sc.getBoundingClientRect();
    const vv = window.visualViewport;
    const bottom = Math.min(box.bottom, vv ? vv.offsetTop + vv.height : box.bottom);
    const pad = 40;
    let dy = 0;
    if (line.top < box.top + pad) dy = line.top - box.top - pad;
    else if (line.bottom > bottom - pad) dy = line.bottom - bottom + pad;
    if (dy) sc.scrollTop += dy;
    return dy;
  }

  // Najbliższy wiersz w stronę dir w kolumnie x (w tabeli: komórka pod/nad kursorem), wiersz
  // najbliżej w pionie. Bez trafienia w kolumnie — najbliższy w ogóle.
  function nextLine(from, dir, x) {
    let best = null;
    let bestAny = null;
    const reach = window.innerHeight * 3;
    for (const p of paras()) {
      const b = p.getBoundingClientRect();
      if (!b.height && !b.width) continue;
      if (dir > 0 ? (b.bottom <= from.y + 1 || b.top > from.y + reach) : (b.top >= from.y - 1 || b.bottom < from.y - reach)) continue;
      // kolumna = komórka tabeli (↓ w tabeli idzie do komórki pod spodem, nie obok); poza tabelą
      // każdy akapit (wcięty punkt listy też — kursor przy lewym marginesie go nie „omija”)
      const cell = p.closest("td, th")?.getBoundingClientRect();
      const inCol = !cell || (x >= cell.left - 1 && x <= cell.right + 1);
      lineBoxes(p).forEach((l, index) => {
        const cy = (l.top + l.bottom) / 2;
        if (dir > 0 ? cy <= from.bottom + 1 : cy >= from.top - 1) return;
        const d = Math.abs(cy - from.y);
        const cand = { p, line: l, index, d };
        if (inCol && (!best || d < best.d - 0.5)) best = cand;
        if (!bestAny || d < bestAny.d - 0.5 || (Math.abs(d - bestAny.d) <= 0.5 && Math.abs(x - (b.left + b.right) / 2) < Math.abs(x - bestAny.cx))) bestAny = { ...cand, cx: (b.left + b.right) / 2 };
      });
    }
    // w kolumnie, chyba że to dużo dalej niż wiersz obok (np. koniec tabeli → tekst pod nią)
    return best || bestAny;
  }

  // Pozycja w akapicie p w wierszu line najbliżej x (bez lądowania w znaczniku komentarza).
  // caretRangeFromPoint bywa zawodny (lewy brzeg 2. wiersza punktu listy z wysunięciem dawał
  // początek 1. wiersza) — wynik spoza wiersza zastępuje przejście po znakach wiersza.
  // eol: koniec zawiniętego wiersza (ta sama pozycja co początek następnego — moveVertical).
  function positionAt(p, line, x) {
    const cx = Math.max(line.left, Math.min(line.right, x));
    const cy = (line.top + line.bottom) / 2;
    let pos = null;
    try {
      let node = null, offset = 0;
      if (document.caretRangeFromPoint) {
        const r = document.caretRangeFromPoint(cx, cy);
        if (r) { node = r.startContainer; offset = r.startOffset; }
      } else {
        const q = document.caretPositionFromPoint?.(cx, cy);
        if (q) { node = q.offsetNode; offset = q.offset; }
      }
      if (node && p.contains(node)) {
        const mark = (node.nodeType === 1 ? node : node.parentElement)?.closest(MARK_SEL);
        if (mark && p.contains(mark) && mark !== p) {
          const mb = mark.getBoundingClientRect();
          pos = cx > mb.left + mb.width / 2 ? afterNode(mark) : beforeNode(mark);
        } else pos = { node, offset };
      }
    } catch (_) { pos = null; }
    const inLine = (c) => !!c && (c.top + c.bottom) / 2 >= line.top - 1 && (c.top + c.bottom) / 2 <= line.bottom + 1;
    const sides = (q) => [charRectFrom(p, q.node, q.offset, 1), charRectFrom(p, q.node, q.offset, -1)];
    if (pos) {
      const [a, b] = sides(pos);
      if (!inLine(a) && !inLine(b) && (a || b)) pos = null;
    }
    if (!pos) pos = scanLine(p, line, cx);
    const [a, b] = sides(pos);
    pos.eol = !!a && !inLine(a) && inLine(b);
    return pos;
  }
  const beforeNode = (n) => ({ node: n.parentNode, offset: [...n.parentNode.childNodes].indexOf(n) });
  const afterNode = (n) => ({ node: n.parentNode, offset: [...n.parentNode.childNodes].indexOf(n) + 1 });

  // Zapas: przejście po znakach (i wyspach) wiersza — gdy punkt trafia w coś innego.
  function scanLine(p, line, x) {
    const onLine = (c) => (c.top + c.bottom) / 2 >= line.top - 1 && (c.top + c.bottom) / 2 <= line.bottom + 1;
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    const r = document.createRange();
    let best = null;
    const consider = (edge, pos) => { const d = Math.abs(edge - x); if (!best || d < best.d) best = { d, ...pos }; };
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      const mark = n.parentElement?.closest(MARK_SEL);
      if (mark && mark !== p && p.contains(mark)) continue;
      for (let i = 0; i < n.length; i++) {
        r.setStart(n, i); r.setEnd(n, i + 1);
        const c = [...r.getClientRects()].find((q) => q.height > 1);
        if (!c || !onLine(c)) continue;
        consider(c.left, { node: n, offset: i });
        consider(c.right, { node: n, offset: i + 1 });
      }
    }
    p.querySelectorAll(`img, svg, canvas, ${MARK_SEL}`).forEach((el) => {
      const isl = el.matches(MARK_SEL) ? el : el.closest(MARK_SEL) || el;
      if (!p.contains(isl) || isl === p) return;
      const c = boxOf(isl);
      if (!c || !onLine(c)) return;
      consider(c.left, beforeNode(isl));
      consider(c.right, afterNode(isl));
    });
    if (best) return { node: best.node, offset: best.offset };
    // wiersz bez znaków — pusta linijka po łamaniu wiersza (<br>): przed <br> w tym wierszu,
    // inaczej za ostatnim <br> nad nim
    const brs = [...p.querySelectorAll("br")];
    const here = brs.find((br) => { const c = br.getClientRects()[0]; return c && onLine(c); });
    if (here) return beforeNode(here);
    const above = brs.filter((br) => { const c = br.getClientRects()[0]; return c && (c.top + c.bottom) / 2 < line.top; }).pop();
    if (above) return afterNode(above);
    const r2 = document.createRange();
    r2.selectNodeContents(p);
    r2.collapse(line.top <= p.getBoundingClientRect().top + 2);
    return { node: r2.startContainer, offset: r2.startOffset };
  }

  function moveVertical(p, dir, e) {
    const sel = getSelection();
    if (!sel?.rangeCount || !sel.focusNode) return false;
    const fp = focusPoint(p);
    const lines = lineBoxes(p);
    const cur = lines.find((l) => fp.y >= l.top - 1 && fp.y <= l.bottom + 1)
      || lines.reduce((a, l) => (Math.abs((l.top + l.bottom) / 2 - fp.y) < Math.abs((a.top + a.bottom) / 2 - fp.y) ? l : a), lines[0]);
    const goalX = fp.x;
    const from = { y: (cur.top + cur.bottom) / 2, top: cur.top, bottom: cur.bottom };
    const target = nextLine(from, dir, goalX);
    e.preventDefault();
    let pos;
    let y;
    if (!target) {
      // pierwszy/ostatni wiersz dokumentu: na początek/koniec wiersza (jak Word)
      const all = paras();
      const edge = dir > 0 ? all[all.length - 1] : all[0];
      const r = document.createRange();
      r.selectNodeContents(edge || p);
      r.collapse(dir < 0);
      pos = { node: r.startContainer, offset: r.startOffset };
      y = from.y;
    } else {
      // po przewinięciu te same wiersze są gdzie indziej — bierzemy go po kolejności
      const line = revealLine(target.p, target.line) ? lineBoxes(target.p)[target.index] || target.line : target.line;
      pos = positionAt(target.p, line, goalX);
      y = (line.top + line.bottom) / 2;
    }
    const put = (q) => {
      if (e.shiftKey) { sel.extend(q.node, q.offset); return; }
      const r = document.createRange();
      r.setStart(q.node, q.offset);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
    };
    if (pos.eol && sel.modify && target) {
      // Koniec zawiniętego wiersza: pozycja ustawiona wprost rysuje się na początku NASTĘPNEGO
      // wiersza — początek docelowego wiersza i „do końca wiersza” (jak End) daje właściwy.
      const line = lineBoxes(target.p)[target.index] || target.line;
      put(scanLine(target.p, line, line.left - 1000));
      sel.modify(e.shiftKey ? "extend" : "move", "forward", "lineboundary");
    } else put(pos);
    const fp2 = (sel.focusNode.nodeType === 1 ? sel.focusNode : sel.focusNode.parentElement)?.closest(".docx-editable-p");
    if (fp2) remember(fp2, sel.focusNode, sel.focusOffset, y, goalX);
    return true;
  }

  // Model końca zaznaczenia (focus): akapit i liczba znaków od jego początku.
  function focusModel() {
    const sel = getSelection();
    const node = sel?.focusNode;
    const p = (node?.nodeType === 1 ? node : node?.parentElement)?.closest?.(".docx-editable-p") || null;
    if (!p) return { p: null, off: -1 };
    const r = document.createRange();
    r.selectNodeContents(p);
    try { r.setEnd(node, sel.focusOffset); } catch (_) { return { p, off: -1 }; }
    return { p, off: r.toString().length };
  }

  // ← / → (z Shift; Alt na Macu / Ctrl gdzie indziej = o słowo). Niewidoczny znacznik komentarza
  // to dla przeglądarki osobna pozycja — naciśnięcie „przechodziło” przez znacznik, a kursor stał
  // (przy „(Jak 2:19, 20)” trzy naciśnięcia bez ruchu). Powtarzamy krok, aż kursor się przesunie.
  function stepHorizontal(e, dir, unit) {
    const sel = getSelection();
    if (!sel?.modify) return false;
    const before = focusModel();
    if (!before.p) return false;
    e.preventDefault();
    for (let i = 0; i < 12; i++) {
      sel.modify(e.shiftKey ? "extend" : "move", dir > 0 ? "forward" : "backward", unit);
      const now = focusModel();
      if (!now.p || now.p !== before.p || now.off !== before.off) break;
    }
    nav = null;
    revealFocus();
    return true;
  }

  // Ctrl+↑/↓ (Alt+↑/↓ na Macu) — początek akapitu / poprzedni / następny (jak Word).
  function moveParagraph(p, dir, e) {
    const sel = getSelection();
    const fm = focusModel();
    let target = p;
    let offset = 0;
    if (dir < 0 && fm.off === 0) target = neighbour(p, -1);
    if (dir > 0) target = neighbour(p, 1);
    if (!target) { target = p; offset = dir > 0 ? textLength(p) : 0; }
    e.preventDefault();
    if (e.shiftKey) {
      const pos = boundaryOf(target, offset);
      sel.extend(pos.node, pos.offset);
      target.scrollIntoView({ block: "nearest" });
    } else placeAt(target, offset);
    nav = null;
  }
  function boundaryOf(p, offset) {
    const r = document.createRange();
    r.selectNodeContents(p);
    r.collapse(offset === 0);
    if (offset === 0) {
      const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      const t = w.nextNode();
      if (t && !t.parentElement.closest(MARK_SEL)) return { node: t, offset: 0 };
    }
    return { node: r.startContainer, offset: r.startOffset };
  }

  // Po End (⌘→ na Macu) kursor stoi na KOŃCU wiersza — przy zawinięciu ta sama pozycja co początek
  // następnego. Zapamiętujemy, że to koniec tego wiersza (znak przed kursorem), inaczej ↓ ominąłby wiersz.
  function rememberLineEnd(now = false) {
    const run = () => {
      const sel = getSelection();
      const fm = focusModel();
      if (!fm.p || !sel?.focusNode) return;
      const c = charRectFrom(fm.p, sel.focusNode, sel.focusOffset, -1);
      if (c) remember(fm.p, sel.focusNode, sel.focusOffset, (c.top + c.bottom) / 2, c.right);
    };
    if (now) run(); else requestAnimationFrame(run); // ⌘→ robi przeglądarka — po jej ruchu
  }
  // Klik: kursor w wierszu, w który kliknięto (koniec krótkiego wiersza ≠ początek następnego).
  function rememberClick(e) {
    if (readOnlyMode || e.button !== 0) return;
    requestAnimationFrame(() => {
      const sel = getSelection();
      const fm = focusModel();
      if (!fm.p || !sel?.isCollapsed) return;
      const b = charRectFrom(fm.p, sel.focusNode, sel.focusOffset, -1);
      const a = charRectFrom(fm.p, sel.focusNode, sel.focusOffset, 1);
      const near = (c) => (c ? Math.abs((c.top + c.bottom) / 2 - e.clientY) : Infinity);
      if (!a && !b) return;
      if (near(b) < near(a)) remember(fm.p, sel.focusNode, sel.focusOffset, (b.top + b.bottom) / 2, b.right);
      else remember(fm.p, sel.focusNode, sel.focusOffset, (a.top + a.bottom) / 2, a.left);
    });
  }

  // Home / End — początek / koniec wiersza (Word). Na Macu Chrome tylko przewijał stronę, a Shift+End
  // zaznaczał do końca całego pola edycji (poza akapit).
  function moveLineBoundary(e, dir) {
    const sel = getSelection();
    if (!sel?.modify) return;
    e.preventDefault();
    sel.modify(e.shiftKey ? "extend" : "move", dir > 0 ? "forward" : "backward", "lineboundary");
    nav = null;
    revealFocus();
    if (dir > 0) rememberLineEnd(true);
  }

  // PageUp / PageDown — kursor o ekran (Word), nie samo przewinięcie: treść przesuwa się o ekran,
  // kursor zostaje w tym samym miejscu okna (ta sama kolumna).
  function movePage(p, dir, e) {
    const sel = getSelection();
    if (!sel?.rangeCount) return;
    const fp = focusPoint(p);
    const sc = scroller(p);
    const view = sc === document.scrollingElement ? window.innerHeight : sc.clientHeight;
    e.preventDefault();
    const before = sc.scrollTop;
    sc.scrollTop += dir * Math.max(80, view * 0.85);
    const moved = sc.scrollTop - before;
    let y = fp.y;
    if (!moved) y += dir * view; // już przy brzegu dokumentu — sam kursor idzie dalej
    let target = null;
    for (const q of paras()) {
      const b = q.getBoundingClientRect();
      if (b.bottom < y - 400 || b.top > y + 400) continue;
      lineBoxes(q).forEach((l, index) => {
        const d = y < l.top ? l.top - y : y > l.bottom ? y - l.bottom : 0;
        if (!target || d < target.d) target = { p: q, line: l, index, d };
      });
    }
    if (!target) {
      const all = paras();
      const edge = dir > 0 ? all[all.length - 1] : all[0];
      if (!edge) return;
      const pos = boundaryOf(edge, dir > 0 ? textLength(edge) : 0);
      if (e.shiftKey) sel.extend(pos.node, pos.offset); else placeAt(edge, dir > 0 ? textLength(edge) : 0);
      nav = null;
      return;
    }
    const pos = positionAt(target.p, target.line, fp.x);
    if (e.shiftKey) sel.extend(pos.node, pos.offset);
    else {
      const r = document.createRange();
      r.setStart(pos.node, pos.offset);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
    }
    remember(target.p, sel.focusNode, sel.focusOffset, (target.line.top + target.line.bottom) / 2, fp.x);
  }

  function onKeydown(e) {
    if (readOnlyMode || e.isComposing || e.defaultPrevented) return;
    const p = docCaretParagraph(e.target);
    if (!p) return;
    const sel = window.getSelection();
    const collapsed = !!sel?.isCollapsed;
    const mac = /Mac|iPhone|iPad/.test(navigator.platform || "");
    const plain = !e.ctrlKey && !e.metaKey && !e.altKey;

    // początek / koniec dokumentu (z Shift — zaznaczenie do początku / końca, jak Word)
    if ((e.ctrlKey && !e.metaKey && !e.altKey && (e.key === "Home" || e.key === "End")) || (mac && e.metaKey && !e.altKey && !e.ctrlKey && (e.key === "ArrowUp" || e.key === "ArrowDown"))) {
      const list = paras();
      if (!list.length) return;
      e.preventDefault();
      const toEnd = e.key === "End" || e.key === "ArrowDown";
      const target = toEnd ? list[list.length - 1] : list[0];
      nav = null;
      if (e.shiftKey) {
        const pos = boundaryOf(target, toEnd ? textLength(target) : 0);
        sel.extend(pos.node, pos.offset);
        target.scrollIntoView({ block: "nearest" });
      } else placeAt(target, toEnd ? textLength(target) : 0);
      return;
    }
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      const dir = e.key === "ArrowDown" ? 1 : -1;
      if (plain) { moveVertical(p, dir, e); return; }
      if ((mac ? e.altKey && !e.ctrlKey : e.ctrlKey && !e.altKey) && !e.metaKey) { moveParagraph(p, dir, e); return; }
      return;
    }
    if (e.key === "ArrowLeft" || e.key === "ArrowRight") {
      const dir = e.key === "ArrowRight" ? 1 : -1;
      const word = (mac ? e.altKey && !e.ctrlKey : e.ctrlKey && !e.altKey) && !e.metaKey;
      if (!plain && !word) {
        if (mac && e.metaKey && dir > 0 && !e.altKey && !e.ctrlKey) rememberLineEnd();
        return;
      }
      if (!e.shiftKey && !collapsed) { nav = null; return; } // zaznaczenie → przeglądarka zwija je do brzegu
      if (!e.shiftKey && !word) {
        const offset = typeof getCaretOffset === "function" ? getCaretOffset(p) : -1;
        // na brzegu akapitu — sąsiedni edytowalny (zablokowane przeskakujemy)
        if (dir > 0 && offset >= textLength(p)) {
          const next = neighbour(p, 1);
          if (next) { e.preventDefault(); nav = null; placeAt(next, 0); }
          return;
        }
        if (dir < 0 && offset === 0) {
          const prev = neighbour(p, -1);
          if (prev) { e.preventDefault(); nav = null; placeAt(prev, textLength(prev)); }
          return;
        }
      }
      stepHorizontal(e, dir, word ? "word" : "character");
      return;
    }
    if ((e.key === "Home" || e.key === "End") && !e.ctrlKey && !e.metaKey && !e.altKey) { moveLineBoundary(e, e.key === "End" ? 1 : -1); return; }
    if ((e.key === "PageDown" || e.key === "PageUp") && !e.ctrlKey && !e.metaKey && !e.altKey) { movePage(p, e.key === "PageDown" ? 1 : -1, e); return; }
    if (!plain || e.shiftKey || !collapsed) return;
    if (e.key === "Delete" && atModelEnd(p)) {
      // jak Backspace na początku następnego akapitu (ta sama ścieżka zapisu i cofania)
      const next = neighbour(p, 1);
      const all = collectPreviewParagraphElements(host());
      // następny to sam podział strony — Delete go usuwa (jak w Wordzie)
      const after = all[all.indexOf(p) + 1];
      if (after?.dataset.lock === "lockPageBreakOnly" && typeof dwbSel !== "undefined") {
        e.preventDefault();
        dwbSel.removeBreakParagraph(all.indexOf(after), p, "before");
        return;
      }
      if (!next || all.indexOf(next) !== all.indexOf(p) + 1 || typeof handleInlineBackspace !== "function") return;
      e.preventDefault();
      placeAt(next, 0);
      handleInlineBackspace(next, all.indexOf(next), { preventDefault() {}, key: "Backspace", fromDelete: true });
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
  docCanvasEl?.addEventListener("mouseup", rememberClick);
})();
