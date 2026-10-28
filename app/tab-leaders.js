// tab-leaders.js — długość kropek („Miejscowość: ………… Teren nr:”) w Edycji (2026-10-28).
//
// Zgłoszenie (karty z PDF): kropek nie da się skrócić — kursor nie wchodzi między nie. Kropki to
// w Wordzie JEDEN znak tabulatora z wypełnieniem, a ich koniec to pozycja tabulatora akapitu
// (w Wordzie przeciąga się ją na linijce). Tutaj: klik w kropki zaznacza je i pokazuje uchwyt na
// ich końcu (przeciągnij) oraz kartę ◀ krótsze / dłuższe ▶. W trakcie podgląd na żywo (tylko ten
// akapit), po puszczeniu jedna operacja „tabStop” na pliku = jeden krok Cofnij, zgodnie z Wordem.
// Kursor po kliknięciu stoi za kropkami jak dawniej — Backspace dalej usuwa je w całości.

const dwbTabLeaders = (() => {
  const PX2PT = 0.75;
  const STEP_PT = 3; // ◀ ▶ — mniej więcej jedna kropka
  let sel = null; // { sp, p, stop, start, prevStops }
  let ui = null; // { handle, card }
  let commitTimer = 0;

  const editable = () => typeof readOnlyMode !== "undefined" && !readOnlyMode;
  const leaderTab = (el) => el?.closest?.(".docx-edit-root span.docx-tab.tab-dot, .docx-edit-root span.docx-tab.tab-line") || null;

  // Początek tabulatora i koniec (pozycja tabulatora) w pt od lewej krawędzi kolumny — jak
  // layoutTabStops (docx-render-fixes.js)
  function geometry(sp) {
    const p = sp.closest("p");
    if (!p) return null;
    const pr = p.getBoundingClientRect();
    const scale = p.offsetWidth ? pr.width / p.offsetWidth : 1;
    const cs = getComputedStyle(p);
    const originPx = pr.left - (parseFloat(cs.marginLeft) || 0) * scale;
    const r = sp.getBoundingClientRect();
    const start = ((r.left - originPx) / scale) * PX2PT;
    const end = start + (parseFloat(sp.style.width) || (r.width / scale) * PX2PT);
    let stops = [];
    try { stops = JSON.parse(sp.dataset.stops || "[]"); } catch (_) { stops = []; }
    return { p, scale, originPx, start, end, stops };
  }
  function stopOf(g) {
    let best = null;
    g.stops.forEach((s, i) => { if (s[2] !== "clear" && Math.abs(s[0] - g.end) < 1.2 && (!best || Math.abs(s[0] - g.end) < Math.abs(best.pos - g.end))) best = { i, pos: s[0], leader: s[1], align: s[2] }; });
    return best;
  }

  // nowa pozycja tabulatora na podglądzie: dane wszystkich tabulatorów akapitu + przeliczenie
  function preview(pos) {
    if (!sel) return;
    const stops = sel.prevStops.map((s, i) => (i === sel.stop.i ? [pos, s[1], s[2]] : s));
    const json = JSON.stringify(stops);
    sel.p.querySelectorAll("span.docx-tab").forEach((t) => { if (t.dataset.stops) t.dataset.stops = json; });
    if (typeof layoutTabStops === "function") layoutTabStops(sel.p);
    sel.pos = pos;
    place();
  }
  function clampPos(pos) {
    const others = sel.prevStops.map((s) => s[0]).filter((v, i) => i !== sel.stop.i);
    const next = Math.min(...others.filter((v) => v > sel.stop.pos), Infinity);
    const prev = Math.max(...others.filter((v) => v < sel.stop.pos), -Infinity);
    const min = Math.max(sel.start + 1, prev + 1);
    const max = Math.min(next - 1, sel.maxPos);
    return Math.round(Math.max(min, Math.min(max, pos)) * 20) / 20;
  }

  function select(sp) {
    if (sel?.sp === sp) return;
    if (typeof shouldUseMobileReflow === "function" && shouldUseMobileReflow()) return; // Widok mobilny: kropki łamane inaczej niż na kartce
    clear();
    const g = geometry(sp);
    const stop = g && stopOf(g);
    if (!stop) return; // tabulator domyślny (bez pozycji w akapicie) — nie ma czego przesuwać
    const sec = sp.closest("section.docx");
    const secCs = sec ? getComputedStyle(sec) : null;
    // nie dalej niż prawy margines strony / komórki
    const box = sp.closest("td, th") || sec;
    const boxR = box ? box.getBoundingClientRect() : null;
    const maxPos = boxR ? ((boxR.right - g.originPx) / g.scale) * PX2PT - (box === sec ? (parseFloat(secCs.paddingRight) || 0) * PX2PT : 2) : stop.pos + 200;
    sel = { sp, p: g.p, stop, start: g.start, prevStops: g.stops.map((s) => s.slice()), pos: stop.pos, maxPos: Math.max(maxPos, stop.pos) };
    build(); // podświetlenie = nakładka; zmiana klasy w dokumencie budziła obserwatory edytora (przebudowa fragmentu, Backspace za kropkami przestawał działać)
  }
  function clear() {
    clearTimeout(commitTimer);
    if (sel && sel.pos !== sel.stop.pos) commit(); // niezapisana zmiana z ◀ ▶
    sel = null;
    ui?.hl.remove();
    ui?.handle.remove();
    ui?.card.remove();
    ui = null;
  }

  function build() {
    const hl = document.createElement("div");
    hl.className = "dwb-tabstop-hl";
    const handle = document.createElement("div");
    handle.className = "dwb-tabstop-handle";
    handle.setAttribute("role", "slider");
    handle.setAttribute("aria-label", t("tabLeaderHandle"));
    handle.dataset.hint = "";
    handle.dataset.hintPl = I18N.pl.tabLeaderHandle;
    handle.dataset.hintEn = I18N.en.tabLeaderHandle;
    handle.dataset.hintDelay = "0.5";
    handle.tabIndex = -1;
    const card = document.createElement("div");
    card.className = "image-card dwb-tabstop-card";
    card.setAttribute("role", "toolbar");
    card.setAttribute("aria-label", t("tabLeaderTools"));
    card.innerHTML = `<span class="dwb-tabstop-label"></span><button type="button" class="tb-btn" data-d="-1">◀</button><button type="button" class="tb-btn" data-d="1">▶</button>`;
    card.querySelector(".dwb-tabstop-label").textContent = t("tabLeaderLabel");
    card.querySelectorAll("button").forEach((b) => {
      const key = b.dataset.d === "1" ? "tabLeaderLonger" : "tabLeaderShorter";
      b.setAttribute("aria-label", t(key));
      b.dataset.hint = "";
      b.dataset.hintPl = I18N.pl[key];
      b.dataset.hintEn = I18N.en[key];
      b.dataset.hintDelay = "0.4";
    });
    card.addEventListener("pointerdown", (e) => e.preventDefault()); // kursor zostaje za kropkami
    card.addEventListener("click", (e) => {
      const b = e.target.closest("button[data-d]");
      if (!b || !sel) return;
      preview(clampPos(sel.pos + Number(b.dataset.d) * STEP_PT * (e.shiftKey ? 4 : 1)));
      clearTimeout(commitTimer);
      commitTimer = setTimeout(commit, 600); // kilka kliknięć = jedna zmiana w pliku
    });
    handle.addEventListener("pointerdown", onHandleDown);
    document.body.append(hl, handle, card);
    ui = { hl, handle, card };
    place();
  }
  function place() {
    if (!sel || !ui) return;
    const r = sel.sp.getBoundingClientRect();
    const vp = docViewportEl?.getBoundingClientRect();
    const hidden = !sel.sp.isConnected || (vp && (r.bottom < vp.top || r.top > vp.bottom));
    ui.hl.style.visibility = ui.handle.style.visibility = ui.card.style.visibility = hidden ? "hidden" : "";
    if (hidden) return;
    Object.assign(ui.hl.style, { left: `${r.left + window.scrollX}px`, top: `${r.top + window.scrollY}px`, width: `${r.width}px`, height: `${r.height}px` });
    ui.handle.style.left = `${r.right + window.scrollX}px`;
    ui.handle.style.top = `${r.top + window.scrollY}px`;
    ui.handle.style.height = `${Math.max(14, r.height)}px`;
    const w = ui.card.offsetWidth, h = ui.card.offsetHeight;
    const left = Math.max(8, Math.min(r.right - w / 2, window.innerWidth - w - 8));
    const top = r.top - h - 10 >= 8 ? r.top - h - 10 : r.bottom + 12;
    ui.card.style.left = `${left + window.scrollX}px`;
    ui.card.style.top = `${top + window.scrollY}px`;
  }

  // ── przyciąganie z ⌘ (Mac) / Ctrl (2026-10-29, prośba Mateusza) ──
  // Bez klawisza — swobodnie (jak dotąd). Z klawiszem koniec kropek staje na: końcu kropek w innym
  // wierszu tej samej karty / kolumny (równe „Teren nr:” jeden pod drugim), pozycji sprzed
  // przeciągania, prawym brzegu. W pliku to dalej zwykła pozycja tabulatora — zgodność z Wordem.
  const SNAP_PX = 6;
  const snapKey = (e) => (/Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent) ? e.metaKey : e.ctrlKey);
  function snapTargets() { // w pt od początku kolumny akapitu
    const box = sel.sp.closest("td, th") || sel.sp.closest("section.docx");
    const out = [{ pos: sel.stop0, kind: "start" }, { pos: sel.maxPos, kind: "edge" }];
    box?.querySelectorAll("span.docx-tab.tab-dot, span.docx-tab.tab-line").forEach((t) => {
      if (t === sel.sp || t.closest("p") === sel.p || (t.closest("td, th") || null) !== (sel.sp.closest("td, th") || null)) return;
      const r = t.getBoundingClientRect();
      if (r.width < 2) return;
      out.push({ pos: ((r.right - drag.originPx) / drag.scale) * PX2PT, kind: "tab", el: t });
    });
    return out;
  }
  let guide = null;
  function showGuide(pos) {
    if (pos == null) { guide?.remove(); guide = null; return; }
    const box = (sel.sp.closest("td, th") || sel.sp.closest("section.docx")).getBoundingClientRect();
    if (!guide) { guide = document.createElement("div"); guide.className = "dwb-tabstop-guide"; document.body.append(guide); }
    const x = drag.originPx + (pos / PX2PT) * drag.scale;
    Object.assign(guide.style, { left: `${x + window.scrollX}px`, top: `${box.top + window.scrollY}px`, height: `${box.height}px` });
  }
  function dragTo(clientX, snap) {
    let pos = clampPos((((clientX - drag.dx) - drag.originPx) / drag.scale) * PX2PT);
    let hit = null;
    if (snap) {
      const tol = (SNAP_PX * (drag.touch ? 2 : 1) / drag.scale) * PX2PT;
      snapTargets().forEach((t) => { const d = Math.abs(t.pos - pos); if (d <= tol && (!hit || d < Math.abs(hit.pos - pos))) hit = t; });
      if (hit) pos = clampPos(hit.pos);
    }
    preview(pos);
    showGuide(hit && Math.abs(hit.pos - pos) < 0.5 ? pos : null);
  }
  function onSnapKey(e) {
    if (!drag || !/^(Meta|Control)$/.test(e.key)) return;
    drag.snap = e.type === "keydown";
    dragTo(drag.lastX, drag.snap);
  }

  let drag = null;
  function onHandleDown(e) {
    if (!sel || e.button > 0) return;
    e.preventDefault();
    e.stopPropagation();
    clearTimeout(commitTimer);
    const g = geometry(sel.sp);
    sel.stop0 = sel.pos;
    drag = { id: e.pointerId, originPx: g.originPx, scale: g.scale, dx: e.clientX - sel.sp.getBoundingClientRect().right, lastX: e.clientX, snap: snapKey(e), touch: e.pointerType === "touch" };
    window.addEventListener("keydown", onSnapKey, true);
    window.addEventListener("keyup", onSnapKey, true);
    ui.handle.classList.add("is-dragging");
    try { ui.handle.setPointerCapture(e.pointerId); } catch (_) { /* — */ }
    ui.handle.addEventListener("pointermove", onHandleMove);
    ui.handle.addEventListener("pointerup", onHandleUp);
    ui.handle.addEventListener("pointercancel", onHandleUp);
  }
  function onHandleMove(e) {
    if (!drag || e.pointerId !== drag.id) return;
    e.preventDefault();
    drag.lastX = e.clientX;
    drag.snap = snapKey(e);
    dragTo(e.clientX, drag.snap);
  }
  function onHandleUp(e) {
    if (!drag || e.pointerId !== drag.id) return;
    drag = null;
    showGuide(null);
    window.removeEventListener("keydown", onSnapKey, true);
    window.removeEventListener("keyup", onSnapKey, true);
    ui?.handle.classList.remove("is-dragging");
    ui?.handle.removeEventListener("pointermove", onHandleMove);
    commit();
  }

  async function commit() {
    clearTimeout(commitTimer);
    if (!sel || Math.abs(sel.pos - sel.stop.pos) < 0.05) return;
    const s = sel;
    const index = typeof resolveParaIndex === "function" ? resolveParaIndex(s.p) : -1;
    const from = Math.round(s.stop.pos * 20), to = Math.round(s.pos * 20);
    s.stop = { ...s.stop, pos: s.pos }; // kolejna zmiana liczy się od nowej pozycji
    s.prevStops = s.prevStops.map((v, i) => (i === s.stop.i ? [s.pos, v[1], v[2]] : v));
    if (index < 0) return;
    const tabIdx = [...s.p.querySelectorAll("span.docx-tab")].indexOf(s.sp);
    // kursor zostaje, gdzie był (klik w kropki stawia go za nimi) — Enter / pisanie dalej działa
    const cp = docCaretParagraph(document.activeElement);
    if (cp && docCanvasEl.contains(cp) && typeof getCaretOffset === "function") pendingInlineCursor = { paraIndex: resolveParaIndex(cp), offset: getCaretOffset(cp) };
    const n = await applyDocumentEdit({ op: "tabStop", index, from, to, leader: s.stop.leader, align: s.stop.align })
      .catch((err) => { if (typeof log === "function") log(`Kropki: ${err.message || err}`, "error"); return 0; });
    if (!n) return;
    // po przerysowaniu: te same kropki dalej zaznaczone
    const t0 = performance.now();
    const again = () => {
      if (typeof inlineLocksPending !== "undefined" && inlineLocksPending && performance.now() - t0 < 4000) { requestAnimationFrame(again); return; }
      sel = null;
      ui?.hl.remove(); ui?.handle.remove(); ui?.card.remove(); ui = null;
      const p = collectPreviewParagraphElements(docCanvasEl.querySelector(".docx-preview-host"))[index];
      const sp = p?.querySelectorAll("span.docx-tab")[tabIdx];
      if (sp && leaderTab(sp)) select(sp);
    };
    requestAnimationFrame(again);
  }

  // Tabulatory akapitu od nowa (raz na klatkę) — po zmianie tekstu przed nimi.
  const pending = new Set();
  function relayout(p) {
    if (!p || typeof layoutTabStops !== "function") return;
    if (!pending.size) requestAnimationFrame(() => { pending.forEach((q) => { if (q.isConnected) layoutTabStops(q); }); pending.clear(); if (sel) place(); });
    pending.add(p);
  }

  // ── kropki w zaznaczeniu (jak w Wordzie) ─────────────────────────────────────────
  // Zaznaczenie przez tabulator działało (Delete usuwał go z tekstem), ale kropek w nim nie było
  // widać — przeglądarka podświetla tylko znak w środku, nie całe pudełko z kropkami. Rysujemy
  // podświetlenie systemowym kolorem zaznaczenia nad każdym tabulatorem w zaznaczeniu (nakładka
  // w body; zmiana klas w dokumencie budzi obserwatory edytora).
  let marks = [];
  let markRaf = 0;
  function paintSelectedTabs() {
    markRaf = 0;
    marks.forEach((m) => m.remove());
    marks = [];
    const s = window.getSelection?.();
    if (!s?.rangeCount || s.isCollapsed) { syncHl(); return; }
    const r = s.getRangeAt(0);
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    let common = r.commonAncestorContainer;
    if (common.nodeType !== 1) common = common.parentElement;
    if (!host || !common || !(host.contains(common) || common.contains(host))) return;
    const scope = host.contains(common) ? common : host;
    const vp = docViewportEl?.getBoundingClientRect();
    const probe = document.createRange();
    for (const t of scope.querySelectorAll("span.docx-tab")) {
      if (marks.length > 400 || !r.intersectsNode(t)) continue;
      const b = t.getBoundingClientRect();
      if (b.width < 1 || (vp && (b.bottom < vp.top || b.top > vp.bottom))) continue;
      // wysokość pasa jak zaznaczenie liter tuż obok kropek (sąsiad z lewej / z prawej w tej linijce)
      const p = t.closest("p");
      let top = null, bottom = null;
      if (p) {
        const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        for (let n = w.nextNode(); n; n = w.nextNode()) {
          if (!n.data.trim() || t.contains(n)) continue;
          probe.selectNodeContents(n);
          for (const x of probe.getClientRects()) {
            if (x.bottom < b.top - 4 || x.top > b.bottom + 4) continue;
            if (Math.abs(x.right - b.left) > 3 && Math.abs(x.left - b.right) > 3) continue;
            top = top == null ? x.top : Math.min(top, x.top);
            bottom = bottom == null ? x.bottom : Math.max(bottom, x.bottom);
          }
        }
      }
      if (top == null) { probe.selectNodeContents(t); const tr = probe.getBoundingClientRect(); top = tr.height ? tr.top : b.top; bottom = top + (tr.height || b.height); }
      const h = bottom - top;
      const m = document.createElement("div");
      m.className = "dwb-tab-selmark";
      Object.assign(m.style, { left: `${b.left + window.scrollX}px`, top: `${top + window.scrollY}px`, width: `${b.width}px`, height: `${h}px` });
      document.body.appendChild(m);
      m.__tab = t;
      marks.push(m);
    }
    syncHl();
  }
  // kropki zaznaczone i zarazem wybrane do zmiany długości — jedno podświetlenie, nie dwa
  function syncHl() {
    if (ui?.hl) ui.hl.style.display = sel && marks.some((m) => m.__tab === sel.sp) ? "none" : "";
  }
  const repaintMarks = () => { if (!markRaf) markRaf = requestAnimationFrame(paintSelectedTabs); };

  // Dwuklik w tabulator zaznacza go w całości (Word); przeglądarka brała słowo obok.
  // Potem Delete / Backspace usuwa kropki, pisanie wpisuje tekst w ich miejsce.
  function selectTab(sp) {
    const s = window.getSelection?.();
    if (!s) return;
    const r = document.createRange();
    r.setStartBefore(sp);
    r.setEndAfter(sp);
    s.removeAllRanges();
    s.addRange(r);
  }

  function init() {
    if (!docCanvasEl) return;
    document.addEventListener("selectionchange", repaintMarks);
    docCanvasEl.addEventListener("dblclick", (e) => {
      if (!editable()) return;
      const sp = e.target.closest?.(".docx-edit-root span.docx-tab");
      if (!sp) return;
      e.preventDefault();
      selectTab(sp);
    });
    docCanvasEl.addEventListener("click", (e) => {
      if (!editable()) return;
      const sp = leaderTab(e.target);
      if (sp) select(sp);
      else if (sel && !e.target.closest?.(".dwb-tabstop-card, .dwb-tabstop-handle")) clear();
    });
    document.addEventListener("pointerdown", (e) => {
      if (sel && !e.target.closest?.(".dwb-tabstop-card, .dwb-tabstop-handle") && !leaderTab(e.target)) clear();
    }, true);
    document.addEventListener("keydown", (e) => {
      if (!sel || drag) return;
      if (e.key === "Escape") { clear(); return; }
      if (!e.ctrlKey && !e.metaKey && !e.altKey && !/^(Shift|Control|Meta|Alt)$/.test(e.key)) {
        clear(); // pisanie / Backspace — dalej jak zwykły tekst
        // niezapisana zmiana z ◀ ▶ właśnie idzie do pliku (przerysowanie): ten klawisz po nim — Enter
        // przed zapisem dzielił akapit tylko w pliku (test chaos: ◀ ×3 → Enter)
        if (typeof holdKeyEvent === "function") holdKeyEvent(e);
      }
    }, true);
    // Pisanie przed tabulatorem: koniec kropek stoi na pozycji tabulatora (jak w Wordzie), a nie
    // przesuwa się z tekstem — inaczej „Teren nr: 401” wyjeżdżał za kolumnę.
    docCanvasEl.addEventListener("input", (e) => {
      if (e.isComposing) return;
      const n = getSelection()?.focusNode;
      const p = (n?.nodeType === 1 ? n : n?.parentElement)?.closest?.(".docx-edit-root p");
      if (p?.querySelector("span.docx-tab")) relayout(p);
    });
    docViewportEl?.addEventListener("scroll", () => { place(); repaintMarks(); }, { passive: true });
    window.addEventListener("resize", () => { place(); repaintMarks(); });
    if (typeof readModeEl !== "undefined") readModeEl?.addEventListener("change", () => clear());
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();

  // niezapisana zmiana z ◀ ▶ (zapis odroczony o chwilę) — od razu do pliku; Cofnij czeka na nią
  // (dawniej zapis odpalał w środku cofania i kasował „Ponów”; test chaos)
  let flushing = null;
  function flush() {
    if (!sel || sel.pos === sel.stop.pos) return flushing;
    clearTimeout(commitTimer);
    flushing = commit().finally(() => { flushing = null; });
    return flushing;
  }
  return { select, clear, relayout, flush, isActive: () => !!sel };
})();
