// Mobile doc zoom: fit-to-width, wider zoom range, viewport hooks.

const MOBILE_MQ = window.matchMedia("(max-width: 768px)");
const ZOOM_MIN_DESKTOP = 0.7;
const ZOOM_MAX_DESKTOP = 1.4;
const ZOOM_MIN_MOBILE = 0.35;
const ZOOM_MAX_MOBILE = 1.4;

const docZoomShellEl = document.getElementById("docZoomShell");
const zoomFitBtnEl = document.getElementById("zoomFitBtn");
const zoomResetBtnEl = document.getElementById("zoomResetBtn");

let zoomMode = "fit"; // [EN] fit = auto width; manual = user slider

function isMobileViewport() {
  return MOBILE_MQ.matches;
}

function syncViewportClass() {
  rootEl.classList.toggle("is-mobile", isMobileViewport());
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
  if (!docViewportEl || !docCanvasEl) return 1;
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
  if (!docZoomShellEl || !docCanvasEl) return;
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
  const active = isMobileViewport() && zoomMode === "fit";
  zoomFitBtnEl.classList.toggle("primary", active);
  zoomFitBtnEl.setAttribute("aria-pressed", active ? "true" : "false");
}

function setZoomMode(mode) {
  zoomMode = mode === "manual" ? "manual" : "fit";
  syncZoomFitButtonState();
}

function applyFitToWidth() {
  setZoomMode("fit");
  const fit = computeFitZoom();
  if (zoomLevelEl) zoomLevelEl.value = String(fit);
  applyZoom();
  updateZoomShellHeight(fit);
}

function resetZoomTo100() {
  setZoomMode("manual");
  if (zoomLevelEl) zoomLevelEl.value = "1";
  applyZoom();
  updateZoomShellHeight(1);
}

function syncMobileDocZoomAfterRender() {
  syncZoomSliderLimits();
  if (!isMobileViewport()) return;
  if (zoomMode === "fit") applyFitToWidth();
  else updateZoomShellHeight(parseFloat(zoomLevelEl?.value) || 1);
}

function onZoomSliderInput() {
  setZoomMode("manual");
  applyZoom();
  updateZoomShellHeight(parseFloat(zoomLevelEl?.value) || 1);
}

function onViewportChange() {
  syncViewportClass();
  syncZoomSliderLimits();
  syncDocViewportHeight();
  if (isMobileViewport()) {
    if (zoomMode === "fit") applyFitToWidth();
    else updateZoomShellHeight(parseFloat(zoomLevelEl?.value) || 1);
  } else {
    updateZoomShellHeight(1);
  }
  if (typeof syncSidebarHandle === "function") syncSidebarHandle();
}

function closeMobileSidebarIfOpen() {
  if (isMobileViewport() && isSidebarOpen()) setSidebarOpen(false);
}

function initMobileDocZoom() {
  syncViewportClass();
  syncZoomSliderLimits();
  syncZoomFitButtonState();

  if (zoomFitBtnEl) zoomFitBtnEl.addEventListener("click", applyFitToWidth);
  if (zoomResetBtnEl) zoomResetBtnEl.addEventListener("click", resetZoomTo100);

  MOBILE_MQ.addEventListener("change", onViewportChange);
  window.addEventListener("orientationchange", () => setTimeout(onViewportChange, 120));
}
