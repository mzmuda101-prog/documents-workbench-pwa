// Mobile doc layout: Word-like reflow on small screens (preview only — save unchanged).

const MOBILE_MQ = window.matchMedia("(max-width: 768px)");
const ZOOM_MIN_DESKTOP = 0.7;
const ZOOM_MAX_DESKTOP = 1.4;
const ZOOM_MIN_MOBILE = 0.35;
const ZOOM_MAX_MOBILE = 1.4;

const docZoomShellEl = document.getElementById("docZoomShell");
const zoomFitBtnEl = document.getElementById("zoomFitBtn");
const zoomResetBtnEl = document.getElementById("zoomResetBtn");
const zoomSliderFieldEl = document.getElementById("zoomSliderField");
const zoomReflowHintEl = document.getElementById("zoomReflowHint");

let zoomMode = "fit"; // [EN] fit = mobile reflow; manual = print layout + zoom slider

function isMobileViewport() {
  return MOBILE_MQ.matches;
}

function getZoomMode() {
  return zoomMode;
}

function shouldUseMobileReflow() {
  return isMobileViewport() && zoomMode === "fit";
}

function syncViewportClass() {
  rootEl.classList.toggle("is-mobile", isMobileViewport());
  docCanvasEl?.classList.toggle("doc-reflow-mode", shouldUseMobileReflow());
}

function getZoomLimits() {
  return isMobileViewport()
    ? { min: ZOOM_MIN_MOBILE, max: ZOOM_MAX_MOBILE }
    : { min: ZOOM_MIN_DESKTOP, max: ZOOM_MAX_DESKTOP };
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

function syncMobileZoomUi() {
  const reflow = shouldUseMobileReflow();
  zoomSliderFieldEl?.classList.toggle("hidden", reflow);
  if (zoomReflowHintEl) {
    zoomReflowHintEl.hidden = !reflow;
    if (reflow) zoomReflowHintEl.textContent = t("zoomReflowHint");
  }
  syncZoomFitButtonState();
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

let _docReflowObserver = null;
function ensureMobileReflowObserver() {
  if (!docViewportEl || _docReflowObserver) return;
  _docReflowObserver = new ResizeObserver(() => {
    if (!shouldUseMobileReflow()) return;
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    if (host) applyMobileReflowLayout(host);
  });
  _docReflowObserver.observe(docViewportEl);
}

function measureNaturalPageWidth() {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host) return 0;
  const prevZoom = docCanvasEl.style.getPropertyValue("--doc-zoom") || "1";
  docCanvasEl.style.setProperty("--doc-zoom", "1");
  const w = host.getBoundingClientRect().width || host.offsetWidth;
  docCanvasEl.style.setProperty("--doc-zoom", prevZoom);
  return w;
}

function computeFitZoom() {
  if (!docViewportEl || !docCanvasEl || shouldUseMobileReflow()) return 1;
  const naturalW = measureNaturalPageWidth();
  if (!naturalW) return 1;
  const vpStyle = getComputedStyle(docViewportEl);
  const padX = (parseFloat(vpStyle.paddingLeft) || 0) + (parseFloat(vpStyle.paddingRight) || 0);
  const avail = Math.max(120, docViewportEl.clientWidth - padX - 12);
  const { min, max } = getZoomLimits();
  const fit = avail / naturalW;
  return Math.max(min, Math.min(max, Math.round(fit * 100) / 100));
}

function updateZoomShellHeight(zoom) {
  if (!docZoomShellEl || !docCanvasEl || shouldUseMobileReflow()) {
    if (docZoomShellEl) docZoomShellEl.style.height = "";
    return;
  }
  if (!isMobileViewport()) {
    docZoomShellEl.style.height = "";
    return;
  }
  const host = docCanvasEl.querySelector(".docx-preview-host");
  if (!host) {
    docZoomShellEl.style.height = "";
    return;
  }
  const canvasStyle = getComputedStyle(docCanvasEl);
  const padY = (parseFloat(canvasStyle.paddingTop) || 0) + (parseFloat(canvasStyle.paddingBottom) || 0);
  const scaled = (host.offsetHeight + padY) * zoom;
  docZoomShellEl.style.height = `${Math.ceil(scaled)}px`;
}

function syncZoomFitButtonState() {
  if (!zoomFitBtnEl) return;
  const active = shouldUseMobileReflow();
  zoomFitBtnEl.classList.toggle("primary", active);
  zoomFitBtnEl.setAttribute("aria-pressed", active ? "true" : "false");
}

function setZoomMode(mode, options = {}) {
  const prev = zoomMode;
  zoomMode = mode === "manual" ? "manual" : "fit";
  syncViewportClass();
  syncMobileZoomUi();
  if (prev !== zoomMode && originalFileBytes && isMobileViewport() && !options.skipRerender) {
    renderCurrentDocument();
    return;
  }
  if (shouldUseMobileReflow()) {
    if (zoomLevelEl) zoomLevelEl.value = "1";
    applyZoom();
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    if (host) applyMobileReflowLayout(host);
  }
}

function applyFitToWidth() {
  setZoomMode("fit");
}

function resetZoomTo100() {
  setZoomMode("manual", { skipRerender: !originalFileBytes });
  if (zoomLevelEl) zoomLevelEl.value = "1";
  applyZoom();
  updateZoomShellHeight(1);
}

function syncMobileDocZoomAfterRender() {
  syncZoomSliderLimits();
  syncViewportClass();
  syncMobileZoomUi();
  if (!isMobileViewport()) return;
  if (shouldUseMobileReflow()) {
    if (zoomLevelEl) zoomLevelEl.value = "1";
    applyZoom();
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    if (host) applyMobileReflowLayout(host);
    if (docZoomShellEl) docZoomShellEl.style.height = "";
    return;
  }
  const fit = computeFitZoom();
  if (zoomLevelEl && zoomMode === "fit") zoomLevelEl.value = String(fit);
  updateZoomShellHeight(parseFloat(zoomLevelEl?.value) || 1);
}

function onZoomSliderInput() {
  setZoomMode("manual", { skipRerender: true });
  applyZoom();
  updateZoomShellHeight(parseFloat(zoomLevelEl?.value) || 1);
}

function onViewportChange() {
  const prevMobile = rootEl.classList.contains("is-mobile");
  syncViewportClass();
  const nowMobile = isMobileViewport();
  if (prevMobile !== nowMobile && originalFileBytes) {
    syncZoomSliderLimits();
    syncMobileZoomUi();
    syncDocViewportHeight();
    renderCurrentDocument();
    return;
  }
  syncZoomSliderLimits();
  syncMobileZoomUi();
  syncDocViewportHeight();
  if (!isMobileViewport()) {
    docCanvasEl?.classList.remove("doc-reflow-mode");
    updateZoomShellHeight(1);
    if (typeof syncSidebarHandle === "function") syncSidebarHandle();
    return;
  }
  if (shouldUseMobileReflow()) {
    if (zoomLevelEl) zoomLevelEl.value = "1";
    applyZoom();
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    if (host) applyMobileReflowLayout(host);
    if (docZoomShellEl) docZoomShellEl.style.height = "";
  } else {
    updateZoomShellHeight(parseFloat(zoomLevelEl?.value) || 1);
  }
  if (typeof syncSidebarHandle === "function") syncSidebarHandle();
}

function closeMobileSidebarIfOpen() {
  if (isMobileViewport() && isSidebarOpen()) setSidebarOpen(false);
}

function initMobileDocZoom() {
  syncViewportClass();
  syncZoomSliderLimits();
  syncMobileZoomUi();

  if (zoomFitBtnEl) zoomFitBtnEl.addEventListener("click", applyFitToWidth);
  if (zoomResetBtnEl) zoomResetBtnEl.addEventListener("click", resetZoomTo100);

  ensureMobileReflowObserver();
  MOBILE_MQ.addEventListener("change", onViewportChange);
  window.addEventListener("orientationchange", () => setTimeout(onViewportChange, 120));
}
