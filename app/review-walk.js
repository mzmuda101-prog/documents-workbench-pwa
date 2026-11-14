// review-walk.js — przegląd zmian po kolei (jak Recenzja → Następna / Poprzednia w Wordzie).
//
// Pasek na dole: „Zmiana 3 z 12”, autor, co się zmieniło, ‹ › ✓ ✗. Zamiana („14” → „30”) to jedna
// decyzja, jak w dymku (review-margin.js pairRows). Klawisze, gdy fokus jest w pasku: Enter —
// akceptuj i dalej, Backspace / Delete — odrzuć i dalej, ← → — poprzednia / następna, Esc — koniec.
// Na dotyku: kartę przesuwa się palcem — w prawo akceptuje, w lewo odrzuca.
// Decyzja to zwykłe Akceptuj / Odrzuć wskazanych zmian (jeden krok Cofnij, bez przerysowania).

const dwbWalk = (() => {
  const PART_ORDER = { body: 0, header: 1, footer: 2, footnote: 3, endnote: 4 };
  const SWIPE = 90; // px — od tylu przesunięcie palcem to decyzja
  let bar = null, card = null;
  let units = [];
  let at = 0;
  let busy = false;
  let walkPara = null;

  const isDel = (k) => k === "del" || k === "moveFrom";
  const isIns = (k) => k === "ins" || k === "moveTo";
  const reduced = () => matchMedia("(prefers-reduced-motion: reduce)").matches;

  // zmiany pliku → decyzje w kolejności dokumentu: zamiana (usunięcie + wstawka w tym samym miejscu,
  // ten sam autor) i znak akapitu przy swojej wstawce / usunięciu w tym akapicie — razem
  function buildUnits(scan) {
    const rows = (scan?.changes || []).filter((c) => c.id != null).map((c, i) => ({ ...c, i }));
    rows.sort((a, b) => (PART_ORDER[a.part] ?? 9) - (PART_ORDER[b.part] ?? 9)
      || (a.paraIndex ?? 1e9) - (b.paraIndex ?? 1e9)
      || (Number.isFinite(a.at) ? a.at : 1e9) - (Number.isFinite(b.at) ? b.at : 1e9)
      || (isDel(a.kind) ? 0 : 1) - (isDel(b.kind) ? 0 : 1) || a.i - b.i);
    const out = [];
    rows.forEach((c) => {
      const prev = out[out.length - 1];
      const same = prev && prev.part === c.part && prev.paraIndex === c.paraIndex && prev.author === c.author;
      if (same && prev.rows.length === 1 && isDel(prev.rows[0].kind) && isIns(c.kind) && Number.isFinite(c.at) && c.at === prev.rows[0].at) { prev.rows.push(c); return; }
      if (same && (c.kind === "paraIns" || c.kind === "paraDel") && prev.rows.some((x) => (c.kind === "paraIns" ? isIns : isDel)(x.kind))) { prev.rows.push(c); return; }
      out.push({ part: c.part, paraIndex: c.paraIndex, author: c.author, date: c.date, rows: [c] });
    });
    out.forEach((u) => { u.ids = u.rows.map((x) => x.id); u.key = u.ids.join(","); });
    return out;
  }

  function describe(u) {
    const del = u.rows.find((x) => isDel(x.kind));
    const ins = u.rows.find((x) => isIns(x.kind));
    const q = (s) => `„${rvShort(s, 60)}”`;
    if (del && ins) return t("walkReplace", { old: q(del.text), text: q(ins.text || ins.shown) });
    if (ins) return t("walkIns", { text: q(ins.text || ins.shown) });
    if (del) return t("walkDel", { text: q(del.text) });
    const c = u.rows[0];
    return `${t(RV_KIND_KEY[c.kind] || "reviewKindOtherFormat")}${c.context ? ` · ${q(c.context)}` : ""}`;
  }

  // ── pasek ────────────────────────────────────────────────────────────────
  const ICON = {
    prev: '<polyline points="15 18 9 12 15 6"/>',
    next: '<polyline points="9 18 15 12 9 6"/>',
    ok: '<polyline points="20 6 9 17 4 12"/>',
    no: '<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>',
  };
  const svg = (k) => `<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[k]}</svg>`;
  function button(cls, icon, label, fn) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = `walk-btn ${cls}`;
    b.innerHTML = svg(icon);
    b.setAttribute("aria-label", label);
    b.title = label;
    b.addEventListener("click", fn);
    return b;
  }
  function ensureBar() {
    if (bar) return;
    bar = document.createElement("div");
    bar.className = "walk-bar";
    bar.setAttribute("role", "toolbar");
    bar.setAttribute("aria-label", t("walkAria"));
    bar.tabIndex = -1;
    card = document.createElement("div");
    card.className = "walk-card";
    card.innerHTML = '<div class="walk-top"><span class="walk-count" aria-live="polite"></span><span class="walk-who"></span></div><div class="walk-what"></div><span class="walk-swipe" aria-hidden="true"></span>';
    const nav = document.createElement("div");
    nav.className = "walk-nav";
    nav.append(
      button("walk-prev", "prev", t("walkPrev"), () => go(at - 1)),
      button("walk-next", "next", t("walkNext"), () => go(at + 1)),
      button("walk-ok", "ok", t("walkAccept"), () => decide("accept")),
      button("walk-no", "no", t("walkReject"), () => decide("reject")),
    );
    const close = document.createElement("button");
    close.type = "button";
    close.className = "walk-close";
    close.textContent = t("walkDone");
    close.addEventListener("click", stop);
    bar.append(card, nav, close);
    bar.addEventListener("keydown", onKey);
    bindSwipe();
    document.body.appendChild(bar);
  }

  function onKey(e) {
    if (e.target.closest?.(".walk-close") && (e.key === "Enter" || e.key === " ")) return; // „Gotowe” jak przycisk
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    const k = e.key;
    let act = null;
    if (k === "Enter") act = () => decide("accept");
    else if (k === "Backspace" || k === "Delete") act = () => decide("reject");
    else if (k === "ArrowRight" || k === "ArrowDown") act = () => go(at + 1);
    else if (k === "ArrowLeft" || k === "ArrowUp") act = () => go(at - 1);
    else if (k === "Escape") act = stop;
    if (!act) return;
    e.preventDefault();
    e.stopPropagation();
    act();
  }

  // przesunięcie palcem: karta idzie za palcem 1:1, po puszczeniu — decyzja albo powrót
  function bindSwipe() {
    let x0 = null, y0 = 0, dx = 0, id = null, horizontal = null;
    const hint = () => card.querySelector(".walk-swipe");
    const reset = () => {
      card.style.transition = reduced() ? "" : "transform 200ms cubic-bezier(0.22, 1, 0.36, 1)";
      card.style.transform = "";
      card.classList.remove("is-ok", "is-no");
      hint().textContent = "";
      x0 = null; id = null; horizontal = null;
    };
    card.addEventListener("pointerdown", (e) => {
      if (e.pointerType === "mouse" || busy) return;
      x0 = e.clientX; y0 = e.clientY; dx = 0; id = e.pointerId; horizontal = null;
      card.style.transition = "none";
    });
    card.addEventListener("pointermove", (e) => {
      if (x0 == null || e.pointerId !== id) return;
      dx = e.clientX - x0;
      if (horizontal == null && (Math.abs(dx) > 8 || Math.abs(e.clientY - y0) > 8)) {
        horizontal = Math.abs(dx) > Math.abs(e.clientY - y0);
        if (horizontal) try { card.setPointerCapture?.(id); } catch (_) { /* wskaźnik już zniknął — gest idzie dalej bez przechwycenia */ }
      }
      if (!horizontal) return;
      card.style.transform = `translateX(${dx}px)`;
      const far = Math.abs(dx) >= SWIPE;
      card.classList.toggle("is-ok", far && dx > 0);
      card.classList.toggle("is-no", far && dx < 0);
      hint().textContent = far ? t(dx > 0 ? "walkAccept" : "walkReject") : "";
    });
    const end = (e) => {
      if (x0 == null || e.pointerId !== id) return;
      const decision = horizontal && Math.abs(dx) >= SWIPE ? (dx > 0 ? "accept" : "reject") : null;
      if (!decision) { reset(); return; }
      card.style.transition = reduced() ? "" : "transform 160ms ease-in, opacity 160ms ease-in";
      card.style.transform = `translateX(${dx > 0 ? "" : "-"}110%)`;
      card.style.opacity = "0";
      x0 = null;
      decide(decision).finally(() => { card.style.transition = "none"; card.style.opacity = ""; reset(); });
    };
    card.addEventListener("pointerup", end);
    card.addEventListener("pointercancel", () => reset());
  }

  // ── stan ─────────────────────────────────────────────────────────────────
  async function rescan() {
    const scan = originalFileBytes ? await scanDocxRevisions(originalFileBytes) : null;
    units = buildUnits(scan);
  }

  function paint() {
    const u = units[at];
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    walkPara?.classList.remove("walk-para");
    walkPara = null;
    if (typeof CSS !== "undefined" && CSS.highlights) CSS.highlights.delete("dwb-walk");
    if (!u) return;
    bar.querySelector(".walk-count").textContent = t("walkCount", { n: at + 1, total: units.length });
    bar.querySelector(".walk-who").textContent = [u.author, rvDate(u.date), RV_PART_KEY[u.part] ? t(RV_PART_KEY[u.part]) : ""].filter(Boolean).join(" · ");
    bar.querySelector(".walk-what").textContent = describe(u);
    bar.querySelector(".walk-prev").disabled = at <= 0;
    bar.querySelector(".walk-next").disabled = at >= units.length - 1;
    if (typeof dwbMargin !== "undefined") card.dataset.color = String(dwbMargin.colorOf(u.author));
    // miejsce w dokumencie: akapit z obwódką, wstawiony tekst podświetlony; widok na środku
    if (u.part !== "body" || !Number.isFinite(u.paraIndex) || !host) return;
    const p = collectPreviewParagraphElements(host)[u.paraIndex];
    if (!p) return;
    walkPara = p;
    p.classList.add("walk-para");
    const ins = u.rows.find((x) => isIns(x.kind) && x.shown);
    const r = ins && typeof dwbMargin !== "undefined" ? dwbMargin.rangeIn(p, ins.at, ins.shown) : null;
    if (r && !r.collapsed && typeof Highlight !== "undefined") CSS.highlights.set("dwb-walk", new Highlight(r));
    const rect = (r && !r.collapsed ? r : p).getBoundingClientRect();
    const v = docViewportEl;
    if (v) {
      const vr = v.getBoundingClientRect();
      const room = vr.height - (bar.offsetHeight || 0);
      if (rect.top < vr.top + 40 || rect.bottom > vr.top + room - 24) {
        v.scrollTo({ top: Math.max(0, v.scrollTop + rect.top - vr.top - room / 3), behavior: reduced() ? "auto" : "smooth" });
      }
    }
    // bez aktywowania dymka z marginesu — na wąskim ekranie otwierał drugie „Zaakceptować czy odrzucić?”
    // nad tekstem, a pasek już o to pyta
    if (typeof dwbMargin !== "undefined") dwbMargin.focusItem(null);
  }

  function go(i) {
    if (!units.length) return;
    at = Math.max(0, Math.min(units.length - 1, i));
    paint();
  }

  async function decide(action) {
    const u = units[at];
    if (!u || busy) return;
    busy = true;
    bar.classList.add("is-busy");
    try {
      await applyDocumentEdit({ op: "revisions", action, ids: u.ids });
      if (typeof runReviewScan === "function") runReviewScan().catch(() => {});
      await rescan();
      if (!units.length) {
        toastUndoable(t("walkFinished"));
        stop();
        return;
      }
      at = Math.min(at, units.length - 1); // następna stoi teraz na tym samym miejscu listy
      paint();
      if (!bar.contains(document.activeElement)) bar.focus({ preventScroll: true });
    } finally {
      busy = false;
      bar?.classList.remove("is-busy");
    }
  }

  async function start() {
    if (!originalFileBytes) return false;
    await rescan();
    if (!units.length) { toast(t("reviewNoChanges"), "info"); return false; }
    ensureBar();
    // od zmiany przy kursorze (jak „Następna” w Wordzie), inaczej od pierwszej
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    const sel = window.getSelection();
    const p = sel?.rangeCount && host?.contains(sel.anchorNode) ? (sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement)?.closest("p") : null;
    const pi = p ? collectPreviewParagraphElements(host).indexOf(p) : -1;
    const from = pi >= 0 ? units.findIndex((u) => u.part === "body" && u.paraIndex >= pi) : -1;
    at = from >= 0 ? from : 0;
    if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
    bar.hidden = false;
    document.documentElement.classList.add("walk-on");
    paint();
    bar.focus({ preventScroll: true });
    return true;
  }

  function stop() {
    if (!bar) return;
    bar.hidden = true;
    document.documentElement.classList.remove("walk-on");
    walkPara?.classList.remove("walk-para");
    walkPara = null;
    if (typeof CSS !== "undefined" && CSS.highlights) CSS.highlights.delete("dwb-walk");
  }

  // Cofnij / Ponów w trakcie przeglądu: lista zmian od nowa, ta sama pozycja
  document.addEventListener("dwb:undo-step", () => {
    if (!bar || bar.hidden) return;
    setTimeout(async () => { await rescan(); if (!units.length) { stop(); return; } at = Math.min(at, units.length - 1); paint(); }, 300);
  });

  return { start, stop, isOn: () => !!bar && !bar.hidden, _units: () => units, _at: () => at, _build: buildUnits };
})();

document.getElementById("rvWalkBtn")?.addEventListener("click", () => dwbWalk.start());
