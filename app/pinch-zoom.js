// pinch-zoom.js — przybliżanie i oddalanie dokumentu gestem (jak Word).
//
// Dotyk: dwa palce. Komputer: Ctrl + kółko myszy, szczypanie na gładziku (Chrome/Edge/
// Firefox wysyłają je jako kółko z Ctrl) i gest Safari na Macu (gesturestart/change) —
// jak Ctrl + kółko w Wordzie: przybliża dokument, nie całą aplikację, w miejscu kursora.
// W Widoku mobilnym skaluje tekst (dalej zawija się do szerokości ekranu), w Widoku
// desktopowym przybliża strony jak zdjęcie (mobile-doc-zoom.js).
//
// W trakcie gestu NIE zmieniamy `zoom` (przy dużym dokumencie każda klatka to pełny
// układ strony) — tylko tanie transform: translate/scale na #docZoomShell. Po
// puszczeniu palców zatwierdzamy prawdziwy zoom i dosuwamy przewinięcie tak, żeby punkt
// pod palcami został pod palcami. Używamy zdarzeń touch*, bo pointer events kończą się
// pointercancel, gdy przeglądarka zacznie przewijać jednym palcem.
//
// Jak w Wordzie: w trakcie gestu nad dokumentem widać plakietkę z aktualnym procentem (i ten
// sam procent na pasku), a w pobliżu 100% zoom „przyciąga się” do 100%, żeby łatwo trafić
// w naturalną wielkość. Plakietka gaśnie chwilę po puszczeniu palców.
//
// Widok mobilny kończy się na 50% (niżej to ściana drobnego tekstu). Dalsze zsuwanie
// palców (o 20% poza dół) = „cała strona”: plakietka mówi „Puść — Widok desktopowy”, a po
// puszczeniu przechodzimy na strony dopasowane do ekranu (jak oddalanie zdjęć do siatki).

(() => {
  const shell = docZoomShellEl;
  const vp = docViewportEl;
  if (!shell || !vp) return;

  let g = null; // { kind, z0, z, ratio, ox, oy, fx0, fy0, dist0, mx, my, anchor, adx, ady, toPages }
  let raf = 0;
  let wheelTimer = 0;
  const SNAP_TO_100 = 0.05; // ±5 punktów procentowych wokół 100% = dokładnie 100%
  const TO_PAGES_PAST_MIN = 0.8; // Widok mobilny: zsunięcie do 80% minimum = przejście na strony

  // ── plakietka z procentem ──────────────────────────────────────────────────
  const badge = document.createElement("div");
  badge.className = "zoom-badge";
  badge.setAttribute("aria-hidden", "true");
  document.body.appendChild(badge);
  const zoomNowEl = document.getElementById("zoomNowText");
  let badgeTimer = 0;
  function showBadge(z, toPages = false) {
    const pct = `${Math.round(z * 100)}%`;
    const text = toPages ? t("pinchToPages") : pct;
    if (badge.textContent !== text) badge.textContent = text;
    badge.classList.toggle("is-snap", toPages || Math.abs(z - 1) < 0.001);
    if (!badge.classList.contains("is-on")) {
      const r = vp.getBoundingClientRect();
      badge.style.left = `${r.left + r.width / 2}px`;
      badge.style.top = `${Math.max(r.top, 0) + 14}px`;
      badge.classList.add("is-on");
    }
    if (zoomNowEl) zoomNowEl.textContent = pct; // pasek też na żywo
    clearTimeout(badgeTimer);
  }
  function hideBadgeSoon() {
    clearTimeout(badgeTimer);
    badgeTimer = setTimeout(() => badge.classList.remove("is-on"), 700);
  }

  const dist = (a, b) => Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY);
  const mid = (a, b) => ({ x: (a.clientX + b.clientX) / 2, y: (a.clientY + b.clientY) / 2 });

  function paint() {
    raf = 0;
    if (!g) return;
    shell.style.transform = `translate(${g.mx - g.fx0}px, ${g.my - g.fy0}px) scale(${g.ratio})`;
  }

  // Miejsce w TEKŚCIE pod palcami. W Widoku mobilnym tekst po zmianie zoomu zawija się
  // inaczej, więc proste przeliczenie proporcji przewinięcia gubi miejsce (skok na początek) —
  // po zatwierdzeniu szukamy tego samego znaku i dosuwamy go pod palce.
  function anchorAt(x, y) {
    try {
      if (document.caretRangeFromPoint) return document.caretRangeFromPoint(x, y);
      const pos = document.caretPositionFromPoint?.(x, y);
      if (!pos) return null;
      const r = document.createRange();
      r.setStart(pos.offsetNode, pos.offset);
      return r;
    } catch (_) { return null; }
  }
  function anchorRect(range) {
    if (!range || !shell.contains(range.startContainer)) return null;
    const rects = range.getClientRects();
    const r = rects.length ? rects[0] : range.getBoundingClientRect();
    if (r && (r.height || r.width)) return r;
    const el = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
    return el?.getBoundingClientRect() || null;
  }

  function begin(kind, x, y, dist0 = 1) {
    const r = shell.getBoundingClientRect();
    const anchor = anchorAt(x, y);
    const ar = anchorRect(anchor);
    const z0 = getDocZoom();
    g = { kind, z0, z: z0, ratio: 1, fx0: x, fy0: y, mx: x, my: y, ox: x - r.left, oy: y - r.top, dist0,
      anchor, adx: ar ? x - ar.left : 0, ady: ar ? y - ar.top : 0, toPages: false };
    shell.style.transformOrigin = `${g.ox}px ${g.oy}px`;
    shell.style.willChange = "transform";
    rootEl.classList.add("is-pinching");
    showBadge(z0);
  }

  // raw = zoom wynikający z gestu, jeszcze bez ograniczeń
  function update(raw, x, y) {
    const { min, max } = getZoomLimits();
    let z = Math.max(min, Math.min(max, raw));
    if (Math.abs(z - 1) < SNAP_TO_100) {
      if (!g.snapped && navigator.vibrate) navigator.vibrate(8); // Android: lekkie „tyk” (iOS nie ma API)
      z = 1;
      g.snapped = true;
    } else g.snapped = false;
    g.toPages = g.kind === "touch" && shouldUseMobileReflow() && raw < min * TO_PAGES_PAST_MIN;
    g.z = z;
    g.ratio = z / g.z0;
    showBadge(z, g.toPages);
    g.mx = x;
    g.my = y;
    if (!raf) raf = requestAnimationFrame(paint);
  }

  function end() {
    if (!g) return;
    const s = g;
    g = null;
    clearTimeout(wheelTimer);
    if (raf) cancelAnimationFrame(raf), (raf = 0);
    shell.style.transform = "";
    shell.style.transformOrigin = "";
    shell.style.willChange = "";
    rootEl.classList.remove("is-pinching");
    hideBadgeSoon();
    if (s.toPages && typeof switchToPagesFromPinch === "function") {
      switchToPagesFromPinch(s.anchor);
      return;
    }
    if (Math.abs(s.ratio - 1) < 0.01 && Math.abs(s.mx - s.fx0) + Math.abs(s.my - s.fy0) < 2) {
      if (typeof applyZoom === "function") applyZoom(); // pasek z powrotem na prawdziwą wartość
      return;
    }
    const z1 = commitDocZoom(s.z);
    const u = z1 / s.z0;
    // znak, który był pod palcami, ląduje pod ich ostatnim położeniem
    const ar = anchorRect(s.anchor);
    if (ar) {
      vp.scrollLeft += ar.left + s.adx * u - s.mx;
      vp.scrollTop += ar.top + s.ady * u - s.my;
      return;
    }
    const r = shell.getBoundingClientRect(); // brak tekstu pod palcami (np. obraz) — proporcja
    vp.scrollLeft += r.left + s.ox * u - s.mx;
    vp.scrollTop += r.top + s.oy * u - s.my;
  }

  // ── dotyk: dwa palce ────────────────────────────────────────────────────────
  vp.addEventListener("touchstart", (e) => {
    if (e.touches.length !== 2) { if (e.touches.length > 2) end(); return; }
    if (g) end();
    const [a, b] = [e.touches[0], e.touches[1]];
    const m = mid(a, b);
    begin("touch", m.x, m.y, dist(a, b) || 1);
  }, { passive: true });

  vp.addEventListener("touchmove", (e) => {
    if (!g || g.kind !== "touch") return;
    if (e.touches.length !== 2) return;
    if (e.cancelable) e.preventDefault(); // dwa palce = zoom, nie przewijanie
    const [a, b] = [e.touches[0], e.touches[1]];
    const m = mid(a, b);
    update(g.z0 * (dist(a, b) / g.dist0), m.x, m.y);
  }, { passive: false });

  const onLift = (e) => { if (g?.kind === "touch" && e.touches.length < 2) end(); };
  vp.addEventListener("touchend", onLift, { passive: true });
  vp.addEventListener("touchcancel", onLift, { passive: true });

  // ── komputer: Ctrl + kółko / szczypanie na gładziku ─────────────────────────
  // Kółko myszy: ząbek = 10% (jak Word), szczypanie gładzika: płynnie. Koniec gestu =
  // chwila bez zdarzeń.
  vp.addEventListener("wheel", (e) => {
    if (!e.ctrlKey) return;
    e.preventDefault(); // bez tego przeglądarka przybliżyłaby całą aplikację
    if (g && g.kind !== "wheel") return;
    if (!g) begin("wheel", e.clientX, e.clientY);
    const dy = e.deltaY * (e.deltaMode === 1 ? 33 : 1);
    const notch = e.deltaMode === 1 || (Math.abs(dy) >= 50 && Number.isInteger(dy));
    const cur = g.wz ?? g.z;
    let raw = notch ? Math.round((cur - Math.sign(dy) * 0.1) * 10) / 10 : cur * Math.exp(-dy / 100);
    const { min, max } = getZoomLimits();
    raw = Math.max(min, Math.min(max, raw));
    g.wz = raw; // przy przyciąganiu do 100% dalsze ząbki liczą się od prawdziwej wartości
    update(raw, e.clientX, e.clientY);
    clearTimeout(wheelTimer);
    wheelTimer = setTimeout(end, notch ? 260 : 160);
  }, { passive: false });

  // Safari: na iPhonie/iPadzie gest przybliżania strony nie może się nałożyć na nasz
  // (dotyk obsługują touch*), na Macu to szczypanie na gładziku — obsługujemy je tu.
  vp.addEventListener("gesturestart", (e) => {
    e.preventDefault();
    if (g) return;
    begin("gesture", e.clientX, e.clientY);
  }, { passive: false });
  vp.addEventListener("gesturechange", (e) => {
    e.preventDefault();
    if (g?.kind === "gesture") update(g.z0 * e.scale, g.fx0, g.fy0);
  }, { passive: false });
  vp.addEventListener("gestureend", (e) => {
    e.preventDefault();
    if (g?.kind === "gesture") end();
  }, { passive: false });
})();
