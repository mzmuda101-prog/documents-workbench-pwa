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

  let drag = null;
  function onHandleDown(e) {
    if (!sel || e.button > 0) return;
    e.preventDefault();
    e.stopPropagation();
    clearTimeout(commitTimer);
    const g = geometry(sel.sp);
    drag = { id: e.pointerId, originPx: g.originPx, scale: g.scale, dx: e.clientX - sel.sp.getBoundingClientRect().right };
    ui.handle.classList.add("is-dragging");
    try { ui.handle.setPointerCapture(e.pointerId); } catch (_) { /* — */ }
    ui.handle.addEventListener("pointermove", onHandleMove);
    ui.handle.addEventListener("pointerup", onHandleUp);
    ui.handle.addEventListener("pointercancel", onHandleUp);
  }
  function onHandleMove(e) {
    if (!drag || e.pointerId !== drag.id) return;
    e.preventDefault();
    preview(clampPos((((e.clientX - drag.dx) - drag.originPx) / drag.scale) * PX2PT));
  }
  function onHandleUp(e) {
    if (!drag || e.pointerId !== drag.id) return;
    drag = null;
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

  function init() {
    if (!docCanvasEl) return;
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
      if (!e.ctrlKey && !e.metaKey && !e.altKey && !/^(Shift|Control|Meta|Alt)$/.test(e.key)) clear(); // pisanie / Backspace — dalej jak zwykły tekst
    }, true);
    // Pisanie przed tabulatorem: koniec kropek stoi na pozycji tabulatora (jak w Wordzie), a nie
    // przesuwa się z tekstem — inaczej „Teren nr: 401” wyjeżdżał za kolumnę.
    docCanvasEl.addEventListener("input", (e) => {
      if (e.isComposing) return;
      const n = getSelection()?.focusNode;
      const p = (n?.nodeType === 1 ? n : n?.parentElement)?.closest?.(".docx-edit-root p");
      if (p?.querySelector("span.docx-tab")) relayout(p);
    });
    docViewportEl?.addEventListener("scroll", () => place(), { passive: true });
    window.addEventListener("resize", () => place());
    if (typeof readModeEl !== "undefined") readModeEl?.addEventListener("change", () => clear());
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();

  return { select, clear, relayout, isActive: () => !!sel };
})();
