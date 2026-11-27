// mini-bar.js — mini pasek przy zaznaczeniu (etap 5 wstążki; docs/WSTAZKA.md, makieta v4 .mini).
//
// Jak w Wordzie: po zaznaczeniu tekstu MYSZĄ w Edycji nad zaznaczeniem pojawia się mały pasek:
// krój, rozmiar · B I U · wyróżnienie, kolor · punktory, styl · komentarz, Przepisz. To celowy
// skrót przy tekście (ta sama funkcja co na karcie Narzędzia główne) — przyciski wołają TE SAME
// kontrolki i funkcje co wstążka (#fmt*, composeUi, dwbRewrite), nic nie jest kopią logiki.
// Ikony są klonami ikon z wstążki (ta sama funkcja = ta sama ikona).
//
// Miejsce: nad zaznaczeniem, a gdy tam brak miejsca — pod nim; nie wchodzi na wysuwany panel.
// Przy przewijaniu gaśnie i wraca po zatrzymaniu; zaznaczenie poza ekranem = brak paska.
// Znika po pisaniu, kliknięciu w dokument, zwinięciu zaznaczenia i w Czytaniu. Tylko przy wstążce
// i tylko mysz — na dotyku zostaje systemowe menu zaznaczenia.

const dwbMiniBar = (() => {
  const vp = document.getElementById("docViewport");
  const canvas = document.getElementById("docCanvas");
  if (!vp || !canvas) return null;
  const $ = (id) => document.getElementById(id);
  const cu = () => (typeof composeUi !== "undefined" ? composeUi : null);
  const GAP = 8;

  const bar = document.createElement("div");
  bar.className = "mini-bar";
  bar.hidden = true;
  bar.setAttribute("role", "toolbar");
  bar.dataset.i18nAria = "miniBarAria";
  bar.setAttribute("aria-label", t("miniBarAria"));

  // ── kontrolki ─────────────────────────────────────────────────────────────
  // lista kroju / rozmiaru: lustro listy ze wstążki (opcje kopiowane przy pokazaniu), wybór
  // przechodzi do oryginału — jego obsługa (też „Inny…”) robi resztę
  function mirrorSelect(srcId, cls) {
    const s = document.createElement("select");
    s.className = `tb-select ${cls}`;
    s.addEventListener("change", () => {
      const src = $(srcId);
      if (!src) return;
      src.value = s.value;
      src.dispatchEvent(new Event("change", { bubbles: true }));
    });
    return { el: s, srcId };
  }
  function iconFrom(src) {
    const frag = document.createDocumentFragment();
    if (!src) return frag;
    for (const n of src.childNodes) {
      if (n.nodeType !== 1 || n.matches("[data-i18n]:not(svg)")) continue; // bez podpisu karty
      const c = n.cloneNode(true);
      c.removeAttribute("id");
      c.querySelectorAll("[id]").forEach((x) => x.removeAttribute("id"));
      frag.append(c);
    }
    return frag;
  }
  const buttons = [];
  function button(srcSel, act, { mirror = false, menu = false, cls = "" } = {}) {
    const src = document.querySelector(srcSel);
    const b = document.createElement("button");
    b.type = "button";
    b.className = `tb-btn mini-btn ${cls}`.trim();
    b.append(iconFrom(src));
    if (menu) b.setAttribute("aria-haspopup", "menu");
    b.addEventListener("mousedown", (e) => e.preventDefault()); // zaznaczenie w tekście zostaje
    b.addEventListener("click", () => act(b));
    buttons.push({ b, src, srcSel, mirror });
    return b;
  }
  const sep = () => { const s = document.createElement("span"); s.className = "mini-sep"; s.setAttribute("aria-hidden", "true"); return s; };
  const click = (id) => () => $(id)?.click();

  const fam = mirrorSelect("fmtFontFamily", "mini-font");
  const size = mirrorSelect("fmtFontSize", "mini-size");
  const styleBtn = button(".style-pick", (b) => openStyleMenu(b), { menu: true });
  const styleIco = document.querySelector(".style-pick-ico")?.cloneNode(true);
  styleBtn.replaceChildren(...[styleIco, ddIcon()].filter(Boolean));
  const ai = document.createElement("span");
  ai.className = "mini-ai";
  ai.append(button("#rewriteBtn", () => { if (typeof dwbRewrite !== "undefined") dwbRewrite.open(); }));
  bar.append(
    fam.el, size.el, sep(),
    button("#fmtBold", click("fmtBold"), { mirror: true }),
    button("#fmtItalic", click("fmtItalic"), { mirror: true }),
    button("#fmtUnderline", click("fmtUnderline"), { mirror: true }),
    button('[data-rb-act="highlight"]', () => cu()?.applyColor("highlight", "yellow"), { cls: "rb-btn" }),
    button("#fmtColorBtn", (b) => cu()?.openColorMenu(b), { menu: true, cls: "color-btn" }),
    sep(),
    button('[data-rb-act="bullets"]', () => cu()?.applyList("bullet")),
    styleBtn,
    sep(),
    button('[data-rb-act="comment"]', (b) => cu()?.openCommentForm(b)),
    ai,
  );
  document.body.append(bar);

  function ddIcon() {
    const s = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    s.setAttribute("class", "rb-dd");
    s.setAttribute("width", "10"); s.setAttribute("height", "10"); s.setAttribute("viewBox", "0 0 24 24");
    s.setAttribute("fill", "none"); s.setAttribute("stroke", "currentColor"); s.setAttribute("stroke-width", "2.6");
    s.setAttribute("stroke-linecap", "round"); s.setAttribute("aria-hidden", "true");
    s.innerHTML = '<path d="m6 9 6 6 6-6"/>';
    return s;
  }

  // styl akapitu: te same style co lista na wstążce (bez „Nowy styl…” — to jest na wstążce)
  function openStyleMenu(b) {
    const sel = $("fmtParaStyle");
    if (!sel) return;
    const items = Array.from(sel.options)
      .filter((o) => o.value && !o.value.startsWith("act:") && !o.value.startsWith("plus:") && !o.dataset.plus && !o.dataset.cur && !o.hidden && !o.disabled)
      .map((o) => ({ label: o.textContent.trim(), checked: o.value === sel.value, onPick: () => cu()?.applyStyle(o.value) }));
    cu()?.openMenu(b, items);
  }

  // napisy i stan: z oryginałów (język, „wciśnięte”, kolor paska pod „A”)
  function syncState() {
    for (const s of [fam, size]) {
      const src = $(s.srcId);
      if (!src) continue;
      s.el.replaceChildren(...Array.from(src.children).map((c) => c.cloneNode(true)));
      s.el.value = src.value;
      s.el.setAttribute("aria-label", src.getAttribute("aria-label") || "");
      s.el.style.fontFamily = s.srcId === "fmtFontFamily" ? src.style.fontFamily : "";
    }
    for (const { b, src, mirror } of buttons) {
      if (!src) continue;
      const label = src.getAttribute("aria-label") || "";
      b.setAttribute("aria-label", label);
      for (const k of ["hint", "hintPl", "hintEn"]) if (src.dataset[k] != null) b.dataset[k] = src.dataset[k];
      b.dataset.hintDelay = "0.7";
      if (mirror) {
        b.classList.toggle("is-on", src.classList.contains("is-on") || src.getAttribute("aria-pressed") === "true");
        b.setAttribute("aria-pressed", String(b.classList.contains("is-on")));
      }
    }
    const barSrc = $("fmtColorBar");
    const barDst = bar.querySelector(".color-btn .color-btn-bar");
    if (barSrc && barDst) barDst.style.background = barSrc.style.background;
    styleBtn.setAttribute("aria-label", $("fmtParaStyle")?.getAttribute("aria-label") || "");
  }

  // ── kiedy i gdzie ─────────────────────────────────────────────────────────
  let wanted = false; // pasek „należy się” temu zaznaczeniu (mysz w Edycji)
  const on = () => document.body.classList.contains("rb-on") && typeof readOnlyMode !== "undefined" && !readOnlyMode;
  function selRange() {
    const sel = getSelection();
    if (!sel || sel.isCollapsed || !sel.rangeCount) return null;
    const a = sel.anchorNode, f = sel.focusNode;
    if (!canvas.contains(a) || !canvas.contains(f)) return null;
    const ed = (n) => {
      const el = n.nodeType === 1 ? n : n.parentElement;
      return !!el?.isContentEditable && !!el.closest("[contenteditable='true']");
    };
    if (!ed(a) || !ed(f)) return null;
    return sel.getRangeAt(0);
  }
  function hide() {
    wanted = false;
    bar.hidden = true;
    bar.classList.remove("is-off");
  }
  function place() {
    const r = wanted && on() && selRange();
    if (!r) { hide(); return; }
    const rects = Array.from(r.getClientRects()).filter((x) => x.width || x.height);
    if (!rects.length) { hide(); return; }
    const first = rects[0], last = rects[rects.length - 1];
    const V = vp.getBoundingClientRect();
    // wysuwany panel na marginesie (bez pinezki) — pasek nie wchodzi na niego
    const panel = $("toolsPanel");
    const over = panel && typeof isSidebarOpen === "function" && isSidebarOpen() && !document.documentElement.classList.contains("sidebar-docked");
    const minL = Math.max(V.left, over ? panel.getBoundingClientRect().right : V.left);
    const top = Math.min(...rects.map((x) => x.top)), bottom = Math.max(...rects.map((x) => x.bottom));
    if (bottom < V.top + 4 || top > V.bottom - 4 || last.right < minL + 40) { bar.classList.add("is-off"); return; }
    if (bar.hidden) { syncState(); bar.hidden = false; }
    const w = bar.offsetWidth, h = bar.offsetHeight;
    let y = first.top - h - GAP;
    if (y < V.top + 4) y = last.bottom + GAP; // nad zaznaczeniem brak miejsca — pod nim
    const x = Math.max(minL + GAP, Math.min(first.left - GAP, V.right - w - GAP));
    bar.style.top = `${Math.round(y)}px`;
    bar.style.left = `${Math.round(x)}px`;
    bar.classList.remove("is-off");
  }

  // pokazanie: puszczenie przycisku myszy po zaznaczeniu w dokumencie (też dwuklik, trzyklik)
  canvas.addEventListener("pointerup", (e) => {
    if (e.pointerType !== "mouse" || e.button !== 0) return;
    setTimeout(() => {
      if (!on() || !selRange()) return;
      wanted = true;
      place();
    }, 0);
  });
  // nowe kliknięcie w dokumencie, pisanie, zmiana zaznaczenia na pustą — pasek znika
  vp.addEventListener("pointerdown", (e) => { if (!bar.contains(e.target)) hide(); });
  document.addEventListener("keydown", (e) => {
    if (bar.hidden || bar.contains(e.target) || ["Shift", "Control", "Alt", "Meta"].includes(e.key)) return;
    hide();
  }, true);
  let sq = 0;
  document.addEventListener("selectionchange", () => {
    if (!wanted) return;
    cancelAnimationFrame(sq);
    sq = requestAnimationFrame(() => { if (!selRange()) hide(); else if (!bar.hidden) syncState(); });
  });
  // przewijanie: gaśnie, wraca na nowe miejsce po zatrzymaniu
  let st = 0;
  vp.addEventListener("scroll", () => {
    if (!wanted) return;
    bar.classList.add("is-off");
    clearTimeout(st);
    st = setTimeout(place, 160);
  }, { passive: true });
  window.addEventListener("resize", () => { if (wanted) place(); }, { passive: true });
  $("readMode")?.addEventListener("change", hide);
  // B / I / U i inne: stan przycisków nadąża za formatem po poleceniu
  bar.addEventListener("click", () => setTimeout(() => { if (!bar.hidden) syncState(); }, 60));

  return { hide, place, isShown: () => !bar.hidden && !bar.classList.contains("is-off"), el: bar };
})();
