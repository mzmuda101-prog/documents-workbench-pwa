// Core runtime: DOM refs, shared state, and base UI helpers.

const IS_LOW_POWER = (() => {
  try {
    const cores = navigator.hardwareConcurrency || 8;
    const mem = navigator.deviceMemory || 8;
    const coarse = !!(window.matchMedia && window.matchMedia("(pointer: coarse)").matches);
    return cores <= 4 || mem <= 4 || (coarse && cores <= 6);
  } catch (_) {
    return false;
  }
})();
document.documentElement.classList.toggle("low-power", IS_LOW_POWER);

const rootEl = document.documentElement;
const appShellEl = document.querySelector(".app");
const logEl = document.getElementById("log");
const statusEl = document.getElementById("status");
const docViewportEl = document.getElementById("docViewport");
const docCanvasEl = document.getElementById("docCanvas");
const emptyStateEl = document.getElementById("emptyState");
const emptyTitleEl = document.getElementById("emptyTitle");
const emptySubEl = document.getElementById("emptySub");
const DEFAULT_EMPTY_TITLE = emptyTitleEl?.textContent || "";
const DEFAULT_EMPTY_SUB = emptySubEl?.textContent || "";
const docPanelEl = document.querySelector(".doc-panel");

const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const fileNameEl = document.getElementById("fileName");
const fileNameTextEl = document.getElementById("fileNameText");
const loadBtn = document.getElementById("loadBtn");
const loadSampleBtn = document.getElementById("loadSampleBtn");
const saveBtn = document.getElementById("saveBtn");
const saveAsBtn = document.getElementById("saveAsBtn");

const searchQueryEl = document.getElementById("searchQuery");
const searchScopeEl = document.getElementById("searchScope");
const searchHighlightBtn = document.getElementById("searchHighlightBtn");
const searchCountEl = document.getElementById("searchCount");

const structureSummaryEl = document.getElementById("structureSummary");
const headingNavEl = document.getElementById("headingNav");

const editOpEl = document.getElementById("editOp");
const editFindEl = document.getElementById("editFind");
const editReplaceEl = document.getElementById("editReplace");
const editRegexEl = document.getElementById("editRegex");
const editScopeEl = document.getElementById("editScope");
const editReplaceFieldsEl = document.getElementById("editReplaceFields");
const editCaseFieldsEl = document.getElementById("editCaseFields");
const editCaseModeEl = document.getElementById("editCaseMode");
const editTrimFieldsEl = document.getElementById("editTrimFields");
const editTrimModeEl = document.getElementById("editTrimMode");
const editAffixFieldsEl = document.getElementById("editAffixFields");
const editPrefixEl = document.getElementById("editPrefix");
const editSuffixEl = document.getElementById("editSuffix");
const applyEditToolBtnEl = document.getElementById("applyEditToolBtn");

const panelToggle = document.getElementById("panelToggle");
const panelHandle = document.getElementById("panelHandle");
const sidebarScrim = document.getElementById("sidebarScrim");
const themeToggle = document.getElementById("themeToggle");
const brandRefresh = document.getElementById("brandRefresh");
const closeDocBtn = document.getElementById("closeDocBtn");
const appUpdateBtn = document.getElementById("appUpdateBtn");
const networkBadge = document.getElementById("networkBadge");
const loadingOverlayEl = document.getElementById("loadingOverlay");
const loadingTextEl = document.getElementById("loadingText");
const toastContainerEl = document.getElementById("toastContainer");
const zoomLevelEl = document.getElementById("zoomLevel");
const zoomValueEl = document.getElementById("zoomValue");
const readModeEl = document.getElementById("readMode");

const THEME_KEY = "documents-workbench-theme";
const LANG_KEY = "documents-workbench-lang";
const INTRO_PLAYED_KEY = "introPlayed";

let currentFileName = "";
let currentFileType = "";
let originalFileBytes = null;
let pendingDocEdits = [];
let fileHandle = null;
let hasUnsavedChanges = false;
let documentStructure = null;
let searchMatches = [];
let activeSearchIndex = -1;
let readOnlyMode = true;

function log(msg, type = "info") {
  if (!logEl) return;
  const line = document.createElement("div");
  line.className = `log-line log-${type}`;
  const locale = typeof currentLang !== "undefined" ? I18N[currentLang].locale : "pl-PL";
  line.textContent = `${new Date().toLocaleTimeString(locale)} ${msg}`;
  logEl.prepend(line);
}

function toast(msg, type = "info") {
  const toastEl = document.createElement("div");
  toastEl.className = `toast ${type}`;
  const icon = document.createElement("div");
  icon.className = "toast-icon";
  icon.textContent = type === "success" ? "✓" : type === "error" ? "!" : type === "warning" ? "!" : "i";
  const label = document.createElement("div");
  label.textContent = msg;
  toastEl.appendChild(icon);
  toastEl.appendChild(label);
  toastContainerEl.appendChild(toastEl);
  setTimeout(() => {
    toastEl.classList.add("out");
    setTimeout(() => toastEl.remove(), 200);
  }, 2800);
}

function setLoading(isLoading, text) {
  loadingTextEl.textContent = text || t("loadingGeneric");
  loadingOverlayEl.classList.toggle("hidden", !isLoading);
}

function setStatus(msg) {
  statusEl.textContent = msg;
}

function setDirtyState(isDirty) {
  hasUnsavedChanges = !!isDirty;
  statusEl.classList.toggle("unsaved", hasUnsavedChanges);
  document.title = hasUnsavedChanges ? `* ${BASE_TITLE}` : BASE_TITLE;
  if (typeof syncActionButtons === "function") syncActionButtons();
}

function showEmptyState(title, sub) {
  emptyStateEl.classList.remove("hidden");
  emptyTitleEl.textContent = title || DEFAULT_EMPTY_TITLE;
  emptySubEl.textContent = sub || DEFAULT_EMPTY_SUB;
  docCanvasEl.classList.add("hidden");
}

function hideEmptyState() {
  emptyStateEl.classList.add("hidden");
  docCanvasEl.classList.remove("hidden");
}

function isSidebarOpen() {
  return rootEl.classList.contains("sidebar-open");
}

function setSidebarOpen(open) {
  rootEl.classList.toggle("sidebar-open", !!open);
  if (sidebarScrim) sidebarScrim.classList.toggle("hidden", !open);
  if (panelToggle) panelToggle.textContent = open ? t("panelOpen") : t("panelClosed");
  const mobile = window.matchMedia("(max-width: 768px)").matches;
  document.body.style.overflow = open && mobile ? "hidden" : "";
  if (typeof syncSidebarHandle === "function") syncSidebarHandle();
}

function toggleSidebar() {
  setSidebarOpen(!isSidebarOpen());
}

function applyZoom() {
  if (!docCanvasEl || !zoomLevelEl) return;
  const { min, max } = typeof getZoomLimits === "function" ? getZoomLimits() : { min: 0.7, max: 1.4 };
  let zoom = parseFloat(zoomLevelEl.value) || 1;
  zoom = Math.max(min, Math.min(max, zoom));
  zoomLevelEl.value = String(zoom);
  docCanvasEl.style.setProperty("--doc-zoom", String(zoom));
  if (zoomValueEl) zoomValueEl.textContent = `${Math.round(zoom * 100)}%`;
  if (typeof updateZoomShellHeight === "function") updateZoomShellHeight(zoom);
}

function syncDocumentShellClass() {
  document.body.classList.toggle("has-document", !!originalFileBytes);
  if (closeDocBtn) closeDocBtn.classList.toggle("hidden", !originalFileBytes);
  const closePanelBtn = document.getElementById("closeDocPanelBtn");
  if (closePanelBtn) closePanelBtn.classList.toggle("hidden", !originalFileBytes);
}

function syncDocViewportHeight() {
  if (!docPanelEl) return;
  requestAnimationFrame(() => {
    const rect = docPanelEl.getBoundingClientRect();
    const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 720;
    const safeBottom = parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--safe-bottom")) || 0;
    const bottomGap = window.matchMedia("(max-width: 768px)").matches ? 8 + safeBottom : 24;
    const available = Math.floor(viewportHeight - rect.top - bottomGap);
    const minHeight = window.matchMedia("(max-width: 768px)").matches ? 280 : 420;
    docPanelEl.style.setProperty("--doc-panel-height", `${Math.max(minHeight, available)}px`);
  });
}

function formatFileSize(bytes) {
  if (!bytes || bytes < 1024) return `${bytes || 0} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function detectFileType(name, mime) {
  const lower = (name || "").toLowerCase();
  if (lower.endsWith(".docx") || mime === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") return "docx";
  if (lower.endsWith(".pdf") || mime === "application/pdf") return "pdf";
  return "";
}

function clearDocumentState() {
  currentFileName = "";
  currentFileType = "";
  originalFileBytes = null;
  pendingDocEdits = [];
  fileHandle = null;
  documentStructure = null;
  searchMatches = [];
  activeSearchIndex = -1;
  if (docCanvasEl) docCanvasEl.replaceChildren();
  if (structureSummaryEl) structureSummaryEl.replaceChildren();
  if (headingNavEl) headingNavEl.replaceChildren();
  if (searchCountEl) searchCountEl.textContent = "";
  if (fileNameEl) fileNameEl.classList.add("hidden");
  if (fileNameTextEl) fileNameTextEl.textContent = t("noFile");
  baselineParagraphTexts = [];
  setDirtyState(false);
  showEmptyState();
  syncDocumentShellClass();
  if (typeof syncActionButtons === "function") syncActionButtons();
}
