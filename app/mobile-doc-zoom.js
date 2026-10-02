// Układ podglądu dokumentu — dwa widoki jak w Wordzie (sam podgląd, zapis pliku bez zmian):
//
// • Widok mobilny — tekst zawija się do szerokości ekranu (bez podziału na strony,
//   nagłówków i stopek); dwa palce / − + zmieniają wielkość liter („reflowScale”).
// • Widok desktopowy — prawdziwe strony; przybliżanie jak zdjęcie („pageZoom”). Na starcie
//   strona mieści się w szerokości (najwyżej 100%), dopóki ktoś nie ustawi zoomu ręcznie.
//
// Przełącznik ma trzy pozycje: Auto / Widok mobilny / Widok desktopowy. Wybór pamiętany
// OSOBNO dla trzech sytuacji: ekran dotykowy w pionie, w poziomie i komputer (mysz) —
// np. „desktopowy” ustawiony na iPadzie w pionie nie zmienia poziomu. „Auto”: mobilny na
// dotyku w pionie i na każdym wąskim oknie (≤768 px, też komputer i iPad z dwiema apkami
// obok siebie); poza tym desktopowy.
//
// Zmiana widoku = ponowne ułożenie podglądu przez rerenderKeepingEdits (najpierw wkłada do
// pliku tekst wpisany w podglądzie — niezapisane zmiany nie giną) i powrót do tego samego
// akapitu (captureDocScrollAnchor), a nie do proporcji przewinięcia.

const MOBILE_MQ = window.matchMedia("(max-width: 768px)");
const COARSE_MQ = window.matchMedia("(pointer: coarse)");
const PORTRAIT_MQ = window.matchMedia("(orientation: portrait)");
const ZOOM_MIN = 0.5;
const ZOOM_MIN_MOBILE = 0.35; // [PL] telefon: pinch może zejść niżej (cała strona na ekranie)
const ZOOM_MAX = 3;
const VIEW_LAYOUT_KEY = "dwb-view-layout-v1"; // { "touch-portrait": "mobile", ... }
const VIEW_AUTO_TIP_KEY = "dwb-view-auto-tip-v1";
const VIEW_LAYOUTS = ["auto", "mobile", "desktop"];

const docZoomShellEl = document.getElementById("docZoomShell");
const zoomFitBtnEl = document.getElementById("zoomFitBtn");
const zoomResetBtnEl = document.getElementById("zoomResetBtn");
const zoomSliderFieldEl = document.getElementById("zoomSliderField");
const zoomReflowHintEl = document.getElementById("zoomReflowHint");
const viewLayoutPanelEl = document.getElementById("viewLayoutPanel");

let reflowScale = 1; // [PL] skala tekstu w Widoku mobilnym (tekst dalej zawija się do szerokości)
let pageZoom = 1; // [PL] zoom stron w Widoku desktopowym
let pageFitMode = "auto"; // auto = dopasuj, najwyżej 100% · fit = „Dopasuj” (też >100%) · manual
let lastFitWidth = 0;
let appliedLayout = null; // widok, w którym podgląd jest teraz ułożony
let lastScrollAnchor = null; // akapit u góry po ostatnim przewinięciu (nie liczą się te od zmiany rozmiaru)
let scrollAnchorBeforeChange = null;
let lastResizeAt = 0;

function isMobileViewport() {
  return MOBILE_MQ.matches;
}

// ── wybór widoku ─────────────────────────────────────────────────────────────
function viewContext() {
  if (!COARSE_MQ.matches) return "desktop";
  return PORTRAIT_MQ.matches ? "touch-portrait" : "touch-landscape";
}

function readViewLayoutPrefs() {
  try {
    const v = JSON.parse(localStorage.getItem(VIEW_LAYOUT_KEY) || "{}");
    return v && typeof v === "object" ? v : {};
  } catch (_) { return {}; }
}

function getViewLayoutPref(ctx = viewContext()) {
  const v = readViewLayoutPrefs()[ctx];
  return v === "mobile" || v === "desktop" ? v : "auto";
}

function autoViewLayout() {
  return isMobileViewport() || viewContext() === "touch-portrait" ? "mobile" : "desktop";
}

function getViewLayout() {
  const pref = getViewLayoutPref();
  return pref === "auto" ? autoViewLayout() : pref;
}

function shouldUseMobileReflow() {
  return getViewLayout() === "mobile";
}

// zgodność ze starszym API (testy, app-frame): fit = Widok mobilny, manual = desktopowy
function getZoomMode() {
  return shouldUseMobileReflow() ? "fit" : "manual";
}

function setViewLayoutPref(pref) {
  const value = VIEW_LAYOUTS.includes(pref) ? pref : "auto";
  const prefs = readViewLayoutPrefs();
  if (value === "auto") delete prefs[viewContext()];
  else prefs[viewContext()] = value;
  try { localStorage.setItem(VIEW_LAYOUT_KEY, JSON.stringify(prefs)); } catch (_) {}
  return applyViewLayout();
}

// Ułóż podgląd w bieżącym widoku. Zwraca Promise (przebudowa podglądu trwa chwilę).
function applyViewLayout(opts = {}) {
  const layout = getViewLayout();
  const changed = appliedLayout !== null && appliedLayout !== layout;
  // Klasa widoku od razu przekłada tekst (i przewinięcie), więc miejsce bierzemy PRZED nią.
  // Po obrocie / zmianie okna przeglądarka już przełożyła tekst — wtedy ostatnie miejsce
  // zapamiętane przy przewijaniu.
  scrollAnchorBeforeChange = changed ? (opts.auto ? lastScrollAnchor : captureScrollAnchorNow()) : null;
  syncViewportClass();
  syncZoomSliderLimits();
  syncMobileZoomUi();
  if (changed && opts.auto && originalFileBytes) maybeShowAutoViewTip(layout);
  if (changed && originalFileBytes && typeof rerenderKeepingEdits === "function") {
    appliedLayout = layout;
    // nie gub tekstu wpisanego w podglądzie; wróć do akapitu sprzed zmiany układu
    // (po przywróceniu zapisz miejsce od razu — przewinięcie tuż po obrocie jest pomijane)
    return rerenderKeepingEdits({ anchor: scrollAnchorBeforeChange }).then(() => { lastScrollAnchor = captureScrollAnchorNow(); });
  }
  appliedLayout = layout;
  applyZoomForLayout();
  return Promise.resolve();
}

function maybeShowAutoViewTip(layout) {
  try {
    if (localStorage.getItem(VIEW_AUTO_TIP_KEY)) return;
    localStorage.setItem(VIEW_AUTO_TIP_KEY, "1");
  } catch (_) { return; }
  toast(t("viewAutoTip", { view: t(layout === "mobile" ? "viewLayoutMobile" : "viewLayoutDesktop") }));
}

// ── zoom w obu widokach ─────────────────────────────────────────────────────
function syncViewportClass() {
  rootEl.classList.toggle("is-mobile", isMobileViewport());
  docCanvasEl?.classList.toggle("doc-reflow-mode", shouldUseMobileReflow());
}

function getZoomLimits() {
  return { min: isMobileViewport() ? ZOOM_MIN_MOBILE : ZOOM_MIN, max: ZOOM_MAX };
}

function clampZoom(z) {
  const { min, max } = getZoomLimits();
  return Math.round(Math.max(min, Math.min(max, z)) * 100) / 100;
}

function syncZoomSliderLimits() {
  if (!zoomLevelEl) return;
  const { min, max } = getZoomLimits();
  zoomLevelEl.min = String(min);
  zoomLevelEl.max = String(max);
  const val = parseFloat(zoomLevelEl.value) || 1;
  if (val < min) zoomLevelEl.value = String(min);
  if (val > max) zoomLevelEl.value = String(max);
}

function applyZoomForLayout() {
  if (shouldUseMobileReflow()) {
    if (zoomLevelEl) zoomLevelEl.value = String(reflowScale);
    applyZoom();
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    if (host) applyMobileReflowLayout(host);
    if (docZoomShellEl) docZoomShellEl.style.height = "";
    return;
  }
  if (pageFitMode !== "manual" && docCanvasEl?.querySelector(".docx-preview-host section.docx")) {
    const fit = computeFitZoom();
    pageZoom = pageFitMode === "fit" ? fit : Math.min(1, fit);
    lastFitWidth = docViewportEl?.clientWidth || 0;
  }
  if (zoomLevelEl) zoomLevelEl.value = String(clampZoom(pageZoom));
  applyZoom();
}

function syncMobileZoomUi() {
  const reflow = shouldUseMobileReflow();
  if (zoomReflowHintEl) {
    zoomReflowHintEl.hidden = !reflow;
    if (reflow) zoomReflowHintEl.textContent = t("zoomReflowHint");
  }
  zoomSliderFieldEl?.classList.remove("hidden"); // suwak działa w obu widokach
  syncZoomFitButtonState();
  syncViewLayoutUi();
}

function applyMobileReflowLayout(host) {
  if (!host || !shouldUseMobileReflow()) return;
  const section = host.querySelector("section.docx") || host.querySelector(".docx");
  if (!section) return;
  const vpW = docViewportEl?.clientWidth || window.innerWidth;

  [host, section, ...host.querySelectorAll(".docx-wrapper, article")].forEach((el) => {
    el.style.width = "100%";
    el.style.maxWidth = "100%";
    el.style.minWidth = "0";
    el.style.margin = "0";
    el.style.boxSizing = "border-box";
  });

  host.querySelectorAll("[style]").forEach((el) => {
    const w = el.style.width;
    const mw = el.style.maxWidth;
    const ml = el.style.marginLeft;
    const mr = el.style.marginRight;
    if (w && w.endsWith("px") && parseFloat(w) > vpW * 0.5) el.style.width = "100%";
    if (mw && mw.endsWith("px")) el.style.maxWidth = "100%";
    if (ml === "auto" || mr === "auto") {
      el.style.marginLeft = "0";
      el.style.marginRight = "0";
    }
  });

  const readPad = Math.max(10, Math.min(14, Math.round(vpW * 0.028)));
  section.style.paddingLeft = `${readPad}px`;
  section.style.paddingRight = `${readPad}px`;
  section.style.paddingTop = "10px";
  section.style.paddingBottom = "16px";

  host.querySelectorAll("table").forEach((table) => {
    table.style.width = "100%";
    table.style.maxWidth = "100%";
    table.style.tableLayout = "fixed";
  });

  host.querySelectorAll("img").forEach((img) => {
    img.style.maxWidth = "100%";
    img.style.height = "auto";
  });

  host.querySelectorAll("p, div, span, td").forEach((el) => {
    const w = el.style.width;
    if (!w || w.endsWith("%")) return;
    if (parseFloat(w) > 120) {
      el.style.width = "";
      el.style.maxWidth = "100%";
    }
  });
}

// ── miejsce w dokumencie (akapit u góry) — na zmianę widoku i obrót ────────────
function captureScrollAnchorNow() {
  const vp = docViewportEl;
  if (!vp || vp.scrollTop < 1 || typeof captureDocScrollAnchor !== "function") return null;
  return captureDocScrollAnchor();
}

function trackScrollAnchor() {
  if (!docViewportEl) return;
  let timer = 0;
  window.addEventListener("resize", () => { lastResizeAt = performance.now(); }, { capture: true, passive: true });
  docViewportEl.addEventListener("scroll", () => {
    // przewinięcie wymuszone przez nowy układ (obrót, węższe okno) to nie wybór użytkownika
    if (performance.now() - lastResizeAt < 500 || rootEl.classList.contains("is-pinching")) return;
    clearTimeout(timer);
    timer = setTimeout(() => {
      if (performance.now() - lastResizeAt < 500) return;
      lastScrollAnchor = captureScrollAnchorNow();
    }, 120);
  }, { passive: true });
}

let _docReflowObserver = null;
let _refitRaf = 0;
function ensureMobileReflowObserver() {
  if (!docViewportEl || _docReflowObserver) return;
  _docReflowObserver = new ResizeObserver(() => {
    if (shouldUseMobileReflow()) {
      const host = docCanvasEl?.querySelector(".docx-preview-host");
      if (host) applyMobileReflowLayout(host);
      return;
    }
    // Widok desktopowy z dopasowaniem: panel obok / węższe okno → strona dalej się mieści.
    // Próg 20 px: pasek przewijania (pojawia się po zmianie zoomu) nie rozkręca pętli.
    if (pageFitMode === "manual" || _refitRaf) return;
    if (Math.abs((docViewportEl.clientWidth || 0) - lastFitWidth) < 20) return;
    _refitRaf = requestAnimationFrame(() => { _refitRaf = 0; applyZoomForLayout(); });
  });
  _docReflowObserver.observe(docViewportEl);
}

// Szerokość najszerszej strony + boczne odstępy obszaru, przy zoomie 100%.
function measureNaturalPageWidth() {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host) return 0;
  const prevZoom = docCanvasEl.style.getPropertyValue("--doc-zoom") || "1";
  docCanvasEl.style.setProperty("--doc-zoom", "1");
  let w = 0;
  host.querySelectorAll(".docx-wrapper > section.docx").forEach((s) => { w = Math.max(w, s.getBoundingClientRect().width); });
  if (!w) w = host.getBoundingClientRect().width || host.offsetWidth;
  const cs = getComputedStyle(docCanvasEl);
  const padX = (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.paddingRight) || 0);
  docCanvasEl.style.setProperty("--doc-zoom", prevZoom);
  return w + padX;
}

function computeFitZoom() {
  if (!docViewportEl || !docCanvasEl || shouldUseMobileReflow()) return 1;
  const naturalW = measureNaturalPageWidth();
  if (!naturalW) return 1;
  const vpStyle = getComputedStyle(docViewportEl);
  const padX = (parseFloat(vpStyle.paddingLeft) || 0) + (parseFloat(vpStyle.paddingRight) || 0);
  const avail = Math.max(120, docViewportEl.clientWidth - padX - 4);
  return clampZoom(Math.floor((avail / naturalW) * 100) / 100);
}

function updateZoomShellHeight(_zoom) {
  if (docZoomShellEl) docZoomShellEl.style.height = ""; // [EN] CSS zoom expands layout — no manual height hack
}

function syncZoomFitButtonState() {
  if (!zoomFitBtnEl) return;
  const mobile = shouldUseMobileReflow();
  const active = mobile ? Math.abs(reflowScale - 1) < 0.005 : pageFitMode === "fit";
  zoomFitBtnEl.classList.toggle("primary", active);
  zoomFitBtnEl.setAttribute("aria-pressed", active ? "true" : "false");
  zoomFitBtnEl.dataset.hintPl = mobile ? "Tekst w zwykłej wielkości (dopasowany do ekranu)" : "Dopasuj stronę do szerokości";
  zoomFitBtnEl.dataset.hintEn = mobile ? "Text at normal size (fitted to the screen)" : "Fit the page to the width";
}

// Widok mobilny: tekst wraca do 100%. Widok desktopowy: strona dokładnie na szerokość.
function applyFitToWidth() {
  if (shouldUseMobileReflow()) reflowScale = 1;
  else pageFitMode = "fit";
  applyZoomForLayout();
  syncZoomFitButtonState();
}

function resetZoomTo100() {
  if (shouldUseMobileReflow()) reflowScale = 1;
  else { pageFitMode = "manual"; pageZoom = 1; }
  applyZoomForLayout();
  syncZoomFitButtonState();
}

function syncMobileDocZoomAfterRender() {
  appliedLayout = getViewLayout();
  lastScrollAnchor = null; // nowy podgląd (też inny plik) — miejsce zapisze następne przewinięcie
  syncZoomSliderLimits();
  syncViewportClass();
  syncMobileZoomUi();
  applyZoomForLayout();
}

function onZoomSliderInput() {
  const z = clampZoom(parseFloat(zoomLevelEl?.value) || 1);
  if (shouldUseMobileReflow()) reflowScale = z; // [PL] +/− w Widoku mobilnym skalują tekst
  else { pageZoom = z; pageFitMode = "manual"; }
  applyZoom();
  syncZoomFitButtonState();
}

function onViewportChange() {
  syncDocViewportHeight();
  applyViewLayout({ auto: true });
  if (typeof syncSidebarHandle === "function") syncSidebarHandle();
}

function closeMobileSidebarIfOpen() {
  if (isMobileViewport() && isSidebarOpen()) setSidebarOpen(false);
}

// [PL] Dla pinch-zoom.js: jedno miejsce, gdzie gest zatwierdza nowy zoom.
function getDocZoom() {
  return shouldUseMobileReflow() ? reflowScale : parseFloat(zoomLevelEl?.value) || 1;
}

function commitDocZoom(zoom) {
  const z = clampZoom(zoom);
  if (shouldUseMobileReflow()) reflowScale = z;
  else { pageZoom = z; pageFitMode = "manual"; }
  if (zoomLevelEl) zoomLevelEl.value = String(z);
  applyZoom();
  syncZoomFitButtonState();
  return z;
}

function isReflowScaled() {
  return shouldUseMobileReflow() && Math.abs(reflowScale - 1) > 0.005;
}

// ── przełącznik: panel Widok + dymek z paska (stuknięcie w procent) ──────────
const VIEW_LAYOUT_ICONS = {
  auto: '<path d="M12 3v3M12 18v3M3 12h3M18 12h3"/><circle cx="12" cy="12" r="4"/>',
  mobile: '<rect x="7" y="2.5" width="10" height="19" rx="2"/><path d="M11 18.5h2"/>',
  desktop: '<rect x="2.5" y="4" width="19" height="12" rx="1.5"/><path d="M8 20h8M12 16v4"/>',
};
const VIEW_CONTEXT_KEYS = { "touch-portrait": "viewCtxTouchPortrait", "touch-landscape": "viewCtxTouchLandscape", desktop: "viewCtxDesktop" };

function viewLayoutIcon(key) {
  return `<svg class="vl-icon" viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${VIEW_LAYOUT_ICONS[key]}</svg>`;
}

function renderViewLayoutChoices(host) {
  if (!host) return;
  const pref = getViewLayoutPref();
  const autoNow = t(autoViewLayout() === "mobile" ? "viewLayoutMobile" : "viewLayoutDesktop");
  const desc = { auto: t("viewLayoutAutoDesc", { view: autoNow }), mobile: t("viewLayoutMobileDesc"), desktop: t("viewLayoutDesktopDesc") };
  const name = { auto: t("viewLayoutAuto"), mobile: t("viewLayoutMobile"), desktop: t("viewLayoutDesktop") };
  const hadFocus = host.contains(document.activeElement);
  host.innerHTML = "";
  const group = document.createElement("div");
  group.className = "view-layout-opts";
  group.setAttribute("role", "radiogroup");
  group.setAttribute("aria-label", t("viewLayoutLabel"));
  for (const key of VIEW_LAYOUTS) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "view-layout-opt";
    b.dataset.layout = key;
    b.setAttribute("role", "radio");
    b.setAttribute("aria-checked", String(key === pref));
    b.tabIndex = key === pref ? 0 : -1;
    b.innerHTML = `${viewLayoutIcon(key)}<span class="vl-text"><span class="vl-name"></span><span class="vl-desc"></span></span>`;
    b.querySelector(".vl-name").textContent = name[key];
    b.querySelector(".vl-desc").textContent = desc[key];
    group.append(b);
  }
  const note = document.createElement("p");
  note.className = "hint view-layout-note";
  note.textContent = t("viewLayoutRemembered", { ctx: t(VIEW_CONTEXT_KEYS[viewContext()]) });
  host.append(group, note);
  if (hadFocus) group.querySelector('[aria-checked="true"]')?.focus();
}

function bindViewLayoutChoices(host) {
  if (!host || host.dataset.bound) return;
  host.dataset.bound = "1";
  host.addEventListener("click", (e) => {
    const b = e.target.closest(".view-layout-opt");
    if (!b || !host.contains(b)) return;
    setViewLayoutPref(b.dataset.layout);
    if (viewPopEl?.contains(host)) setViewPopOpen(false); // dymek jak menu: wybór zamyka
  });
  host.addEventListener("keydown", (e) => {
    if (!["ArrowDown", "ArrowUp", "ArrowLeft", "ArrowRight"].includes(e.key)) return;
    const opts = [...host.querySelectorAll(".view-layout-opt")];
    const i = opts.indexOf(document.activeElement);
    if (i < 0) return;
    e.preventDefault();
    const next = opts[(i + (e.key === "ArrowDown" || e.key === "ArrowRight" ? 1 : -1) + opts.length) % opts.length];
    setViewLayoutPref(next.dataset.layout);
    host.querySelector(`[data-layout="${next.dataset.layout}"]`)?.focus();
  });
}

let viewPopEl = null;
function ensureViewPop() {
  if (viewPopEl) return viewPopEl;
  viewPopEl = document.createElement("div");
  viewPopEl.className = "view-pop";
  viewPopEl.id = "viewPop";
  viewPopEl.setAttribute("role", "dialog");
  viewPopEl.hidden = true;
  viewPopEl.innerHTML = '<div class="view-pop-title"></div><div class="view-pop-body"></div>';
  document.body.append(viewPopEl);
  bindViewLayoutChoices(viewPopEl.querySelector(".view-pop-body"));
  return viewPopEl;
}

function placeViewPop() {
  const btn = document.getElementById("zoomNow");
  if (!viewPopEl || !btn) return;
  const r = btn.getBoundingClientRect();
  const vw = document.documentElement.clientWidth;
  const width = Math.min(320, vw - 16);
  viewPopEl.style.width = `${width}px`;
  viewPopEl.style.top = `${Math.round(r.bottom + 6)}px`;
  viewPopEl.style.left = `${Math.round(Math.max(8, Math.min(vw - width - 8, r.left + r.width / 2 - width / 2)))}px`;
}

function setViewPopOpen(open, opts = {}) {
  const btn = document.getElementById("zoomNow");
  const pop = ensureViewPop();
  pop.hidden = !open;
  btn?.setAttribute("aria-expanded", String(open));
  if (open) {
    syncViewLayoutUi();
    placeViewPop();
    if (opts.focus) pop.querySelector('.view-layout-opt[aria-checked="true"]')?.focus();
  } else if (opts.returnFocus) btn?.focus();
}

function syncViewLayoutUi() {
  if (viewLayoutPanelEl) renderViewLayoutChoices(viewLayoutPanelEl);
  if (viewPopEl) {
    viewPopEl.setAttribute("aria-label", t("viewLayoutLabel"));
    viewPopEl.querySelector(".view-pop-title").textContent = t("viewLayoutLabel");
    if (!viewPopEl.hidden) renderViewLayoutChoices(viewPopEl.querySelector(".view-pop-body"));
  }
  const btn = document.getElementById("zoomNow");
  if (btn) btn.dataset.layout = getViewLayout();
}

function initViewLayoutSwitch() {
  bindViewLayoutChoices(viewLayoutPanelEl);
  const btn = document.getElementById("zoomNow");
  if (!btn) return;
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    setViewPopOpen(ensureViewPop().hidden, { focus: e.detail === 0 });
  });
  document.addEventListener("pointerdown", (e) => {
    if (!viewPopEl || viewPopEl.hidden || viewPopEl.contains(e.target) || btn.contains(e.target)) return;
    setViewPopOpen(false);
  });
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape" || !viewPopEl || viewPopEl.hidden) return;
    e.preventDefault();
    e.stopPropagation();
    setViewPopOpen(false, { returnFocus: true });
  }, true);
  window.addEventListener("resize", () => { if (viewPopEl && !viewPopEl.hidden) placeViewPop(); }, { passive: true });
}

function initMobileDocZoom() {
  appliedLayout = getViewLayout();
  syncViewportClass();
  syncZoomSliderLimits();
  initViewLayoutSwitch();
  syncMobileZoomUi();
  if (typeof syncSidebarHandle === "function") syncSidebarHandle();

  if (zoomFitBtnEl) zoomFitBtnEl.addEventListener("click", applyFitToWidth);
  if (zoomResetBtnEl) zoomResetBtnEl.addEventListener("click", resetZoomTo100);

  ensureMobileReflowObserver();
  trackScrollAnchor();
  for (const mq of [MOBILE_MQ, COARSE_MQ, PORTRAIT_MQ]) mq.addEventListener("change", onViewportChange);
  window.addEventListener("orientationchange", () => setTimeout(onViewportChange, 120));
}
