// image-viewer.js — podgląd obrazu na cały ekran (2026-10-06).
//
// Otwierany z karty obrazu (przycisk z lupą), dwuklikiem obrazu w Edycji i zwykłym kliknięciem /
// stuknięciem w Czytaniu. Tylko OGLĄDANIE — rozmiar obrazu w dokumencie się nie zmienia.
// Dwa układy (zapamiętywane): „Pół ekranu” — panel obok dokumentu (w poziomie z prawej, w pionie
// na dole), dokument dalej da się czytać i przewijać, klik w inny obraz podmienia podgląd;
// Granicę panelu da się przeciągnąć (uchwyt na krawędzi, też strzałkami): 25–80% ekranu,
// zapamiętane osobno dla układu obok i na dole. „Całe okno” — nad aplikacją, z przyciemnionym tłem. Do tego prawdziwy pełny ekran (Fullscreen
// API; iPhone go nie ma — przycisk znika). Obraz w pełnej rozdzielczości z pliku (nie przycięty do
// strony): przybliżanie i oddalanie kółkiem / szczypaniem / przyciskami / klawiszami + −,
// przesuwanie przeciąganiem, dwuklik = przybliż w tym miejscu ↔ dopasuj, „Dopasuj” (0),
// „Rzeczywisty rozmiar” (1), poprzedni / następny obraz dokumentu (← →), pobranie obrazu.
// Procent na pasku = względem prawdziwego rozmiaru obrazu (100% = piksel obrazu na piksel ekranu).
//
// Rysujemy przez width/height + translate (nie scale): po przybliżeniu obraz jest ostry także
// w Safari (warstwa ze scale bywa rozmyta), a jeden element to tani układ na klatkę.

const imageViewer = (() => {
  let ui = null; // { root, stage, img, zoomVal, cap, count, prev, next, hint }
  let list = [];
  let idx = 0;
  let startIdx = 0;
  let scale = 1;
  let fit = 1;
  let tx = 0;
  let ty = 0;
  let nw = 0;
  let nh = 0;
  let returnFocus = null;
  let hintTimer = 0;
  const pointers = new Map(); // pointerId → { x, y }
  let gesture = null; // { ix, iy, s0, d0 } — punkt obrazu pod palcami + skala na starcie
  let lastTap = null; // { t, x, y } — dwukrotne stuknięcie (dblclick na dotyku bywa zawodny)
  // Pół ekranu: czy ostatni klik był w podglądzie (Safari nie daje fokusu klikniętym przyciskom,
  // więc sam fokus nie mówi, gdzie użytkownik „jest”).
  let lastPointerInside = false;
  const MODE_KEY = "dwb.ivMode";
  let mode = "full"; // "half" | "full"
  try { if (localStorage.getItem(MODE_KEY) === "half") mode = "half"; } catch (_) { /* tryb prywatny */ }
  const fsEl = () => document.fullscreenElement || document.webkitFullscreenElement || null;
  const fsSupported = () => !!(document.fullscreenEnabled || document.webkitFullscreenEnabled);

  const isOpen = () => !!ui && !ui.root.hidden;
  const minScale = () => Math.min(fit, 1) / 2;
  const maxScale = () => Math.max(fit, 1) * 8;
  const clampScale = (s) => Math.max(minScale(), Math.min(maxScale(), s));
  const stageRect = () => ui.stage.getBoundingClientRect();

  const ICON = (path) => `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${path}</svg>`;
  const BUTTONS = [
    ["prev", "ivPrev", '<polyline points="15 18 9 12 15 6"/>'],
    ["next", "ivNext", '<polyline points="9 18 15 12 9 6"/>'],
    ["out", "ivZoomOut", '<circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16" y2="16"/><line x1="8" y1="11" x2="14" y2="11"/>'],
    ["in", "ivZoomIn", '<circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16" y2="16"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/>'],
    ["mode", null, null],
    ["fullscreen", null, null],
    ["fit", "ivFit", '<polyline points="4 9 4 4 9 4"/><polyline points="20 9 20 4 15 4"/><polyline points="4 15 4 20 9 20"/><polyline points="20 15 20 20 15 20"/>'],
    ["actual", "ivActual", null],
    ["download", "ivDownload", '<path d="M12 3v12"/><polyline points="7 10 12 15 17 10"/><path d="M5 21h14"/>'],
    ["close", "ivClose", '<line x1="6" y1="6" x2="18" y2="18"/><line x1="18" y1="6" x2="6" y2="18"/>'],
  ];

  function build() {
    const root = document.createElement("div");
    root.className = "image-viewer";
    root.hidden = true;
    root.setAttribute("role", "dialog");
    root.setAttribute("aria-modal", "true");
    root.tabIndex = -1;
    root.innerHTML = `<div class="iv-bar">
        <div class="iv-info"><span class="iv-count"></span><span class="iv-cap"></span></div>
        <div class="iv-tools">
          <button type="button" class="iv-btn" data-act="out"></button>
          <button type="button" class="iv-zoom-val" data-act="fit"></button>
          <button type="button" class="iv-btn" data-act="in"></button>
          <span class="iv-sep" aria-hidden="true"></span>
          <button type="button" class="iv-btn" data-act="fit"></button>
          <button type="button" class="iv-btn iv-btn-text" data-act="actual">1:1</button>
          <span class="iv-sep" aria-hidden="true"></span>
          <button type="button" class="iv-btn" data-act="mode"></button>
          <button type="button" class="iv-btn" data-act="fullscreen"></button>
          <button type="button" class="iv-btn" data-act="download"></button>
        </div>
        <button type="button" class="iv-btn iv-close" data-act="close"></button>
      </div>
      <div class="iv-split" role="separator" tabindex="0"><span class="iv-split-grip" aria-hidden="true"></span></div>
      <div class="iv-stage"><img class="iv-img" alt="" draggable="false"></div>
      <button type="button" class="iv-btn iv-nav iv-nav-prev" data-act="prev"></button>
      <button type="button" class="iv-btn iv-nav iv-nav-next" data-act="next"></button>
      <div class="iv-hint" aria-hidden="true"></div>`;
    document.body.appendChild(root);
    ui = {
      root,
      stage: root.querySelector(".iv-stage"),
      img: root.querySelector(".iv-img"),
      zoomVal: root.querySelector(".iv-zoom-val"),
      cap: root.querySelector(".iv-cap"),
      count: root.querySelector(".iv-count"),
      prev: root.querySelector(".iv-nav-prev"),
      next: root.querySelector(".iv-nav-next"),
      hint: root.querySelector(".iv-hint"),
    };
    root.addEventListener("click", (e) => {
      const b = e.target.closest("[data-act]");
      if (b) act(b.dataset.act);
    });
    ui.img.addEventListener("load", () => { if (isOpen()) measureAndFit(); });
    bindGestures();
    bindSplit();
    window.addEventListener("resize", refit, { passive: true });
    SIDE_MQ.addEventListener?.("change", () => { if (isOpen()) { syncModeUi(); requestAnimationFrame(refit); } });
    ["fullscreenchange", "webkitfullscreenchange"].forEach((n) => document.addEventListener(n, () => { if (isOpen()) { syncModeUi(); requestAnimationFrame(refit); } }));
  }
  // Zmiana rozmiaru sceny (okno, układ, pełny ekran): dopasowany zostaje dopasowany.
  function refit() {
    if (!isOpen() || !nw) return;
    const wasFit = Math.abs(scale - fit) < 1e-3;
    computeFit();
    if (wasFit) { scale = fit; apply(); } else apply();
  }
  function setHint(b, key) {
    b.setAttribute("aria-label", t(key));
    b.dataset.hint = "";
    b.dataset.hintPl = I18N.pl[key];
    b.dataset.hintEn = I18N.en[key];
    b.dataset.hintDelay = "0.4";
  }
  const HALF_SVG = '<rect x="3" y="4" width="18" height="16" rx="2"/><line x1="12" y1="4" x2="12" y2="20"/><rect x="12" y="4" width="9" height="16" fill="currentColor" stroke="none" opacity=".35"/>';
  const FULL_SVG = '<rect x="3" y="4" width="18" height="16" rx="2"/><rect x="3" y="4" width="18" height="16" rx="2" fill="currentColor" stroke="none" opacity=".35"/>';
  const FS_ON_SVG = '<polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/>';
  const FS_OFF_SVG = '<polyline points="4 14 10 14 10 20"/><polyline points="20 10 14 10 14 4"/><line x1="14" y1="10" x2="21" y2="3"/><line x1="3" y1="21" x2="10" y2="14"/>';
  // Pół ekranu w poziomie: aplikacja zwęża się do lewej połowy (dokument obok obrazu, nic nie
  // jest zasłonięte); w pionie / na wąsko panel zajmuje dolną połowę nad dokumentem.
  const SIDE_MQ = matchMedia("(orientation: landscape) and (min-width: 701px)");
  function syncModeUi() {
    const fs = !!fsEl();
    const half = mode === "half" && !fs;
    const side = half && SIDE_MQ.matches;
    const sideBefore = document.documentElement.classList.contains("iv-side");
    if (half) applySplit(SIDE_MQ.matches ? "side" : "bottom", readSplit(SIDE_MQ.matches ? "side" : "bottom"));
    ui.root.classList.toggle("iv-half", half);
    ui.root.setAttribute("aria-modal", half ? "false" : "true");
    document.documentElement.classList.toggle("iv-open-full", !half);
    document.documentElement.classList.toggle("iv-side", side);
    if (side !== sideBefore) relayoutApp(side);
    const mb = ui.root.querySelector('[data-act="mode"]');
    mb.innerHTML = ICON(half ? FULL_SVG : HALF_SVG);
    setHint(mb, half ? "ivModeFull" : "ivModeHalf");
    mb.hidden = fs;
    const fb = ui.root.querySelector('[data-act="fullscreen"]');
    fb.hidden = !fsSupported();
    fb.innerHTML = ICON(fs ? FS_OFF_SVG : FS_ON_SVG);
    setHint(fb, fs ? "ivFullscreenExit" : "ivFullscreen");
  }
  // Szerokość aplikacji się zmieniła: dopasowania (szerokość strony, karty) liczą się od nowa,
  // a w trybie obok dokument pokazuje miejsce oglądanego obrazu.
  function relayoutApp(side, scroll = true) {
    if (side && typeof composeUi !== "undefined") composeUi.hideImageCard?.();
    requestAnimationFrame(() => {
      window.dispatchEvent(new Event("resize"));
      const shown = list[idx];
      if (side && scroll && shown?.isConnected) shown.scrollIntoView({ block: "center", behavior: "instant" });
    });
  }
  // ── granica panelu „Pół ekranu” ─────────────────────────────────────────────
  const SPLIT_KEY = { side: "dwb.ivSplitSide", bottom: "dwb.ivSplitBottom" };
  const clampSplit = (f) => Math.max(0.25, Math.min(0.8, f));
  const splitKind = () => (SIDE_MQ.matches ? "side" : "bottom");
  function readSplit(kind) {
    let f = 0.5;
    try { f = parseFloat(localStorage.getItem(SPLIT_KEY[kind])) || 0.5; } catch (_) { /* tryb prywatny */ }
    return clampSplit(f);
  }
  // withApp = false w trakcie przeciągania: przesuwa się sam panel (tanio), aplikacja obok
  // układa się raz — po puszczeniu.
  function applySplit(kind, f, withApp = true) {
    const st = document.documentElement.style;
    if (kind === "side") {
      st.setProperty("--iv-side-w", `${f * 100}vw`);
      if (withApp) st.setProperty("--iv-side-app", `${f * 100}vw`);
    } else st.setProperty("--iv-bottom-h", `${f * 100}dvh`);
    const sep = ui?.root.querySelector(".iv-split");
    if (sep) {
      sep.setAttribute("aria-orientation", kind === "side" ? "vertical" : "horizontal");
      sep.setAttribute("aria-valuenow", String(Math.round(f * 100)));
      sep.setAttribute("aria-valuemin", "25");
      sep.setAttribute("aria-valuemax", "80");
      sep.setAttribute("aria-label", t("ivSplit"));
    }
  }
  function saveSplit(kind, f) {
    try { localStorage.setItem(SPLIT_KEY[kind], String(Math.round(f * 1000) / 1000)); } catch (_) { /* tryb prywatny */ }
  }
  function bindSplit() {
    const sep = ui.root.querySelector(".iv-split");
    let drag = null; // { kind, f }
    const fracAt = (e, kind) => clampSplit(kind === "side" ? (innerWidth - e.clientX) / innerWidth : (innerHeight - e.clientY) / innerHeight);
    sep.addEventListener("pointerdown", (e) => {
      if (e.button > 0) return;
      e.preventDefault();
      sep.setPointerCapture?.(e.pointerId);
      drag = { kind: splitKind(), f: readSplit(splitKind()) };
      ui.root.classList.add("iv-resizing");
    });
    sep.addEventListener("pointermove", (e) => {
      if (!drag) return;
      drag.f = fracAt(e, drag.kind);
      applySplit(drag.kind, drag.f, false);
      refit();
    });
    const end = () => {
      if (!drag) return;
      const { kind, f } = drag;
      drag = null;
      ui.root.classList.remove("iv-resizing");
      saveSplit(kind, f);
      applySplit(kind, f);
      if (kind === "side") relayoutApp(true, false);
      requestAnimationFrame(refit);
    };
    sep.addEventListener("pointerup", end);
    sep.addEventListener("pointercancel", end);
    sep.addEventListener("lostpointercapture", end);
    // strzałki / Home / End na uchwycie (panel rośnie w stronę dokumentu)
    sep.addEventListener("keydown", (e) => {
      const kind = splitKind();
      const grow = kind === "side" ? ["ArrowLeft", "ArrowRight"] : ["ArrowUp", "ArrowDown"];
      let f = readSplit(kind);
      if (e.key === grow[0]) f += 0.05; else if (e.key === grow[1]) f -= 0.05;
      else if (e.key === "Home") f = 0.25; else if (e.key === "End") f = 0.8;
      else return;
      e.preventDefault();
      e.stopPropagation();
      f = clampSplit(f);
      saveSplit(kind, f);
      applySplit(kind, f);
      if (kind === "side") relayoutApp(true, false);
      requestAnimationFrame(refit);
    });
  }
  function setMode(m) {
    mode = m;
    try { localStorage.setItem(MODE_KEY, m); } catch (_) { /* tryb prywatny */ }
    syncModeUi();
    requestAnimationFrame(refit);
  }
  function toggleFullscreen() {
    if (fsEl()) { (document.exitFullscreen || document.webkitExitFullscreen)?.call(document); return; }
    const r = ui.root;
    const req = r.requestFullscreen || r.webkitRequestFullscreen;
    try { const p = req?.call(r); if (p?.catch) p.catch(() => {}); } catch (_) { /* przeglądarka odmówiła */ }
  }

  function labels() {
    ui.root.setAttribute("aria-label", t("ivTitle"));
    BUTTONS.forEach(([act, key, svg]) => {
      if (!key) return; // układ i pełny ekran — syncModeUi
      ui.root.querySelectorAll(`.iv-btn[data-act="${act}"]`).forEach((b) => {
        if (svg) b.innerHTML = ICON(svg);
        setHint(b, key);
      });
    });
    ui.zoomVal.setAttribute("aria-label", t("ivFit"));
    ui.hint.textContent = t(matchMedia("(pointer: coarse)").matches ? "ivHintTouch" : "ivHintMouse");
  }

  // ── stan widoku ────────────────────────────────────────────────────────────
  function computeFit() {
    const r = stageRect();
    const pad = r.width < 600 ? 8 : 32;
    fit = nw && nh ? Math.min((r.width - pad * 2) / nw, (r.height - pad * 2) / nh) : 1;
    if (!(fit > 0)) fit = 1;
  }
  // Mniejszy niż ekran — wyśrodkowany; większy — nie da się go odsunąć poza krawędź.
  function clampPan() {
    const r = stageRect();
    const w = nw * scale;
    const h = nh * scale;
    tx = w <= r.width ? (r.width - w) / 2 : Math.min(0, Math.max(r.width - w, tx));
    ty = h <= r.height ? (r.height - h) / 2 : Math.min(0, Math.max(r.height - h, ty));
  }
  function apply() {
    clampPan();
    const s = ui.img.style;
    s.width = `${nw * scale}px`;
    s.height = `${nh * scale}px`;
    s.transform = `translate(${tx}px, ${ty}px)`;
    ui.zoomVal.textContent = `${Math.round(scale * 100)}%`;
    ui.root.classList.toggle("iv-zoomed", nw * scale > stageRect().width + 1 || nh * scale > stageRect().height + 1);
  }
  // Zmiana skali z punktem (x, y na scenie) w miejscu — domyślnie środek sceny.
  function setScale(ns, x, y) {
    const r = stageRect();
    const px = x ?? r.width / 2;
    const py = y ?? r.height / 2;
    ns = clampScale(ns);
    tx = px - (px - tx) * (ns / scale);
    ty = py - (py - ty) * (ns / scale);
    scale = ns;
    apply();
  }
  function zoomBy(f, x, y) { setScale(scale * f, x, y); }
  function measureAndFit() {
    nw = ui.img.naturalWidth || 1;
    nh = ui.img.naturalHeight || 1;
    computeFit();
    scale = fit;
    tx = 0; ty = 0;
    apply();
  }

  // ── lista obrazów dokumentu ────────────────────────────────────────────────
  function collect(img) {
    const hostEl = img.closest(".docx-preview-host") || document.querySelector(".docx-preview-host");
    const all = hostEl ? Array.from(hostEl.querySelectorAll("img")).filter((i) => i.src && !i.closest(".image-viewer")) : [];
    return all.includes(img) ? all : [img];
  }
  function show(i) {
    idx = (i + list.length) % list.length;
    const src = list[idx];
    ui.root.classList.remove("iv-ready");
    ui.img.alt = src.getAttribute("alt") || "";
    ui.cap.textContent = src.getAttribute("alt") || "";
    ui.cap.hidden = !ui.cap.textContent;
    const many = list.length > 1;
    // jeden obraz bez opisu — pasek oddaje całe miejsce przyciskom (wąski telefon)
    ui.count.textContent = many ? t("ivCount", { n: idx + 1, total: list.length }) : "";
    ui.count.hidden = !many;
    ui.root.querySelector(".iv-info").hidden = !many && ui.cap.hidden;
    ui.prev.hidden = !many;
    ui.next.hidden = !many;
    if (ui.img.getAttribute("src") === src.currentSrc || ui.img.getAttribute("src") === src.src) {
      if (ui.img.complete) measureAndFit();
    } else {
      ui.img.src = src.currentSrc || src.src;
      if (ui.img.complete && ui.img.naturalWidth) measureAndFit();
    }
    requestAnimationFrame(() => ui.root.classList.add("iv-ready"));
  }

  function open(img) {
    if (!img?.src) return;
    if (!ui) build();
    labels();
    const wasOpen = isOpen();
    list = collect(img);
    startIdx = list.indexOf(img);
    if (!wasOpen) returnFocus = document.activeElement;
    lastPointerInside = !wasOpen;
    ui.root.hidden = false;
    syncModeUi();
    show(startIdx);
    // Pół ekranu, klik w inny obraz dokumentu: podgląd się podmienia, fokus zostaje w dokumencie
    if (!wasOpen || !ui.root.classList.contains("iv-half")) ui.root.focus({ preventScroll: true });
    if (wasOpen) return;
    // podpowiedź gestów — chwilę, tylko przy pierwszych otwarciach
    let shown = 0;
    try { shown = +localStorage.getItem("dwb.ivHint") || 0; } catch (_) { /* tryb prywatny */ }
    clearTimeout(hintTimer);
    ui.hint.classList.toggle("on", shown < 3);
    if (shown < 3) {
      try { localStorage.setItem("dwb.ivHint", String(shown + 1)); } catch (_) { /* tryb prywatny */ }
      hintTimer = setTimeout(() => ui.hint.classList.remove("on"), 3200);
    }
  }

  function close() {
    if (!isOpen()) return;
    if (fsEl() === ui.root) (document.exitFullscreen || document.webkitExitFullscreen)?.call(document);
    ui.root.hidden = true;
    ui.root.classList.remove("iv-half");
    document.documentElement.classList.remove("iv-open-full");
    if (document.documentElement.classList.contains("iv-side")) {
      document.documentElement.classList.remove("iv-side");
      relayoutApp(false);
    }
    pointers.clear();
    gesture = null;
    const shownImg = list[idx];
    // przejrzany dalej inny obraz — dokument przewija się do niego
    if (idx !== startIdx && shownImg?.isConnected) shownImg.scrollIntoView({ block: "center", behavior: "instant" });
    if (returnFocus?.isConnected) returnFocus.focus?.({ preventScroll: true });
    returnFocus = null;
    list = [];
  }

  async function download() {
    const src = list[idx];
    if (!src) return;
    try {
      const blob = await (await fetch(src.currentSrc || src.src)).blob();
      const ext = (blob.type.split("/")[1] || "png").replace("jpeg", "jpg").replace(/\+.*$/, "");
      const base = (typeof currentFileName === "string" && currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : "obraz");
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `${base}-${t("ivFileWord")}-${idx + 1}.${ext}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    } catch (err) {
      if (typeof log === "function") log(`Obraz: ${err.message || err}`, "error");
    }
  }

  function act(a) {
    if (a === "close") close();
    else if (a === "in") zoomBy(1.5);
    else if (a === "out") zoomBy(1 / 1.5);
    else if (a === "fit") { scale = fit; apply(); }
    else if (a === "actual") setScale(1);
    else if (a === "prev") show(idx - 1);
    else if (a === "next") show(idx + 1);
    else if (a === "download") download();
    else if (a === "mode") setMode(mode === "half" ? "full" : "half");
    else if (a === "fullscreen") toggleFullscreen();
  }

  // ── gesty: przeciąganie, szczypanie, kółko, dwuklik ───────────────────────
  function centroid() {
    const pts = Array.from(pointers.values());
    const r = stageRect();
    const x = pts.reduce((a, p) => a + p.x, 0) / pts.length - r.left;
    const y = pts.reduce((a, p) => a + p.y, 0) / pts.length - r.top;
    const d = pts.length > 1 ? Math.hypot(pts[0].x - pts[1].x, pts[0].y - pts[1].y) : 0;
    return { x, y, d };
  }
  // Przy każdej zmianie liczby palców: zapamiętaj punkt obrazu pod palcami — ruch i szczypanie
  // trzymają go pod palcami (jak zdjęcia w telefonie).
  function anchor() {
    if (!pointers.size) { gesture = null; return; }
    const c = centroid();
    gesture = { ix: (c.x - tx) / scale, iy: (c.y - ty) / scale, s0: scale, d0: c.d };
  }
  function bindGestures() {
    const st = ui.stage;
    st.addEventListener("pointerdown", (e) => {
      if (e.button > 0) return;
      st.setPointerCapture?.(e.pointerId);
      pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
      ui.root.classList.add("iv-grabbing");
      anchor();
      if (pointers.size === 1) gesture.moved = false;
    });
    st.addEventListener("pointermove", (e) => {
      if (!pointers.has(e.pointerId)) return;
      pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
      if (!gesture) return;
      const c = centroid();
      if (pointers.size > 1 && gesture.d0 > 0) scale = clampScale(gesture.s0 * c.d / gesture.d0);
      const nx = c.x - gesture.ix * scale;
      const ny = c.y - gesture.iy * scale;
      if (Math.abs(nx - tx) + Math.abs(ny - ty) > 2 || pointers.size > 1) gesture.moved = true;
      tx = nx; ty = ny;
      apply();
    });
    const lift = (e) => {
      if (!pointers.has(e.pointerId)) return;
      const moved = gesture?.moved;
      const single = pointers.size === 1;
      pointers.delete(e.pointerId);
      anchor();
      if (!pointers.size) ui.root.classList.remove("iv-grabbing");
      if (e.type !== "pointerup" || !single || moved) { lastTap = null; return; }
      // dwukrotne stuknięcie / kliknięcie
      const now = performance.now();
      if (lastTap && now - lastTap.t < 320 && Math.hypot(e.clientX - lastTap.x, e.clientY - lastTap.y) < 30) {
        lastTap = null;
        const r = stageRect();
        if (scale > fit * 1.05) { scale = fit; apply(); } else setScale(Math.max(fit * 2.5, Math.min(1, fit * 4)), e.clientX - r.left, e.clientY - r.top);
      } else {
        lastTap = { t: now, x: e.clientX, y: e.clientY };
        // stuknięcie w tło obok obrazu (nie przybliżonego) zamyka podgląd
        const onImg = e.target === ui.img;
        if (!onImg && !ui.root.classList.contains("iv-zoomed") && !ui.root.classList.contains("iv-half")) {
          const t0 = now;
          setTimeout(() => { if (lastTap?.t === t0) { lastTap = null; close(); } }, 330);
        }
      }
    };
    st.addEventListener("pointerup", lift);
    st.addEventListener("pointercancel", lift);
    st.addEventListener("lostpointercapture", lift);
    // kółko: przybliża w miejscu kursora (szczypanie gładzika przychodzi jako kółko z Ctrl —
    // drobniejsze kroki); z Shift / poziome — przesuwa przybliżony obraz
    st.addEventListener("wheel", (e) => {
      e.preventDefault();
      const r = stageRect();
      const dy = e.deltaMode === 1 ? e.deltaY * 16 : e.deltaY;
      if (!e.ctrlKey && (e.shiftKey || Math.abs(e.deltaX) > Math.abs(e.deltaY))) {
        tx -= e.shiftKey ? dy : e.deltaX;
        if (!e.shiftKey) ty -= dy;
        apply();
        return;
      }
      zoomBy(Math.exp(-dy * (e.ctrlKey ? 0.01 : 0.0025)), e.clientX - r.left, e.clientY - r.top);
    }, { passive: false });
    // Safari (Mac/iPad): bez tego gest szczypania przybliża całą stronę zamiast obrazu
    ["gesturestart", "gesturechange", "gestureend"].forEach((n) => ui.root.addEventListener(n, (e) => e.preventDefault()));
    ui.root.addEventListener("contextmenu", (e) => { if (e.target === ui.img && pointers.size) e.preventDefault(); });
  }

  // Klawisze — na oknie w fazie przechwytywania, przed skrótami aplikacji (Delete nie usuwa
  // obrazu pod spodem, Ctrl+Z nie cofa niewidocznie zmian dokumentu).
  document.addEventListener("pointerdown", (e) => { if (isOpen()) lastPointerInside = ui.root.contains(e.target); }, true);
  window.addEventListener("keydown", (e) => {
    if (!isOpen()) return;
    // Pół ekranu nie blokuje dokumentu: klawisze tylko, gdy fokus jest w podglądzie
    if (e.target?.closest?.(".iv-split")) return; // strzałki na uchwycie zmieniają rozmiar panelu
    if (ui.root.classList.contains("iv-half") && !ui.root.contains(document.activeElement) && !(lastPointerInside && !document.activeElement?.closest?.("input, textarea, [contenteditable]"))) return;
    e.stopPropagation();
    const k = e.key;
    const zoomed = ui.root.classList.contains("iv-zoomed");
    let handled = true;
    if (k === "Escape") close();
    else if (k === "+" || k === "=" || (e.code === "NumpadAdd")) zoomBy(1.25);
    else if (k === "-" || k === "_" || (e.code === "NumpadSubtract")) zoomBy(1 / 1.25);
    else if (k === "0") act("fit");
    else if (k === "f" || k === "F") { if (!(e.ctrlKey || e.metaKey) && fsSupported()) toggleFullscreen(); else handled = false; }
    else if (k === "1") act("actual");
    else if ((k === "ArrowLeft" || k === "ArrowRight") && !zoomed) act(k === "ArrowLeft" ? "prev" : "next");
    else if (k.startsWith("Arrow") && zoomed) {
      const step = e.shiftKey ? 200 : 60;
      if (k === "ArrowLeft") tx += step; else if (k === "ArrowRight") tx -= step;
      else if (k === "ArrowUp") ty += step; else ty -= step;
      apply();
    } else if (k === "PageUp" || k === "PageDown") act(k === "PageUp" ? "prev" : "next");
    else if (k === "Tab") {
      // fokus zostaje w podglądzie
      const items = Array.from(ui.root.querySelectorAll("button")).filter((b) => !b.hidden && b.offsetParent);
      const i = items.indexOf(document.activeElement);
      const nextI = e.shiftKey ? (i <= 0 ? items.length - 1 : i - 1) : (i + 1) % items.length;
      items[nextI]?.focus();
    } else if (k === "Enter" || k === " ") handled = false; // przyciski działają jak zwykle
    else if (!(e.ctrlKey || e.metaKey)) handled = true;
    else handled = false;
    if (handled) e.preventDefault();
  }, true);

  return { open, close, isOpen };
})();
