// pinch-zoom.js — przybliżanie i oddalanie dokumentu dwoma palcami (jak Word mobile).
//
// Działa też przy włączonym „Dopasuj do ekranu”: na telefonie skaluje tekst (dalej
// zawija się do szerokości ekranu), na tablecie przechodzi w zoom ręczny. „Dopasuj”
// wraca do dopasowania. Tylko dotyk (pointer: coarse) — mysz i gładzik bez zmian.
//
// W trakcie gestu NIE zmieniamy `zoom` (przy dużym dokumencie każda klatka to pełny
// układ strony) — tylko tanie transform: translate/scale na #docZoomShell. Po
// puszczeniu palców zatwierdzamy prawdziwy zoom i dosuwamy przewinięcie tak, żeby punkt
// pod palcami został pod palcami. Używamy zdarzeń touch*, bo pointer events kończą się
// pointercancel, gdy przeglądarka zacznie przewijać jednym palcem.

(() => {
  const coarse = matchMedia("(pointer: coarse)");
  const shell = docZoomShellEl;
  const vp = docViewportEl;
  if (!shell || !vp) return;

  let g = null; // { z0, ratio, ox, oy, fx0, fy0, dist0, mx, my }
  let raf = 0;

  const dist = (a, b) => Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY);
  const mid = (a, b) => ({ x: (a.clientX + b.clientX) / 2, y: (a.clientY + b.clientY) / 2 });

  function paint() {
    raf = 0;
    if (!g) return;
    shell.style.transform = `translate(${g.mx - g.fx0}px, ${g.my - g.fy0}px) scale(${g.ratio})`;
  }

  function begin(e) {
    const [a, b] = [e.touches[0], e.touches[1]];
    const m = mid(a, b);
    const r = shell.getBoundingClientRect();
    g = { z0: getDocZoom(), ratio: 1, fx0: m.x, fy0: m.y, mx: m.x, my: m.y, ox: m.x - r.left, oy: m.y - r.top, dist0: dist(a, b) || 1 };
    shell.style.transformOrigin = `${g.ox}px ${g.oy}px`;
    shell.style.willChange = "transform";
    rootEl.classList.add("is-pinching");
  }

  function end() {
    if (!g) return;
    const s = g;
    g = null;
    if (raf) cancelAnimationFrame(raf), (raf = 0);
    shell.style.transform = "";
    shell.style.transformOrigin = "";
    shell.style.willChange = "";
    rootEl.classList.remove("is-pinching");
    if (Math.abs(s.ratio - 1) < 0.01 && Math.abs(s.mx - s.fx0) + Math.abs(s.my - s.fy0) < 2) return;
    const z1 = commitDocZoom(s.z0 * s.ratio);
    // punkt dokumentu, który był pod palcami, ma wylądować pod ich ostatnim położeniem
    const r = shell.getBoundingClientRect();
    const u = z1 / s.z0;
    vp.scrollLeft += r.left + s.ox * u - s.mx;
    vp.scrollTop += r.top + s.oy * u - s.my;
  }

  vp.addEventListener("touchstart", (e) => {
    if (!coarse.matches || e.touches.length !== 2) { if (e.touches.length > 2) end(); return; }
    if (g) end();
    begin(e);
  }, { passive: true });

  vp.addEventListener("touchmove", (e) => {
    if (!g) return;
    if (e.touches.length !== 2) return;
    if (e.cancelable) e.preventDefault(); // dwa palce = zoom, nie przewijanie
    const [a, b] = [e.touches[0], e.touches[1]];
    const { min, max } = getZoomLimits();
    const z = Math.max(min, Math.min(max, g.z0 * (dist(a, b) / g.dist0)));
    g.ratio = z / g.z0;
    const m = mid(a, b);
    g.mx = m.x;
    g.my = m.y;
    if (!raf) raf = requestAnimationFrame(paint);
  }, { passive: false });

  const onLift = (e) => { if (g && e.touches.length < 2) end(); };
  vp.addEventListener("touchend", onLift, { passive: true });
  vp.addEventListener("touchcancel", onLift, { passive: true });

  // Safari: własny gest przybliżania strony nie może się nałożyć na nasz.
  ["gesturestart", "gesturechange", "gestureend"].forEach((n) =>
    vp.addEventListener(n, (e) => e.preventDefault(), { passive: false }));
})();
