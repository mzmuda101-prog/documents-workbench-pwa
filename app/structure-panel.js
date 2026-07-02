// Structure inspector — outline navigation, jump/highlight, quick paragraph edit.

const STRUCTURE_PAGE_SIZE_KEY = "documents-workbench-structure-page-size";

let structureFilter = "all";
let structureSearchQuery = "";
let structurePageSize = Number(localStorage.getItem(STRUCTURE_PAGE_SIZE_KEY)) || 20;
let structurePage = 0;
let structureSelectionId = null;

const structureFilterEl = document.getElementById("structureFilter");
const structureSearchEl = document.getElementById("structureSearch");
const structurePageSizeEl = document.getElementById("structurePageSize");
const structurePageInfoEl = document.getElementById("structurePageInfo");
const structurePrevPageBtn = document.getElementById("structurePrevPage");
const structureNextPageBtn = document.getElementById("structureNextPage");
const structureOutlineEl = document.getElementById("structureOutline");
const structureQuickEditEl = document.getElementById("structureQuickEdit");
const structureQuickEditTextEl = document.getElementById("structureQuickEditText");
const structureQuickApplyBtn = document.getElementById("structureQuickApplyBtn");
const structureQuickFocusBtn = document.getElementById("structureQuickFocusBtn");

if (structurePageSizeEl) structurePageSizeEl.value = String(structurePageSize);

function getStructureSelection() {
  if (!structureSelectionId || !documentStructure?.outline) return null;
  return documentStructure.outline.find((item) => item.id === structureSelectionId) || null;
}

function clearStructureHighlights() {
  docCanvasEl?.querySelectorAll(".search-hit, .search-hit-active").forEach((el) => {
    el.classList.remove("search-hit", "search-hit-active");
  });
}

function structureTypeLabel(type) {
  if (type === "heading") return t("structureTypeHeading");
  if (type === "table") return t("structureTypeTable");
  return t("structureTypeParagraph");
}

function filterStructureOutline(structure) {
  const items = structure?.outline || [];
  if (structureFilter === "headings") return items.filter((i) => i.type === "heading");
  if (structureFilter === "paragraphs") return items.filter((i) => i.type === "paragraph");
  if (structureFilter === "tables") return items.filter((i) => i.type === "table");
  return items;
}

function searchStructureOutline(items) {
  const q = structureSearchQuery.trim().toLowerCase();
  if (!q) return items;
  return items.filter((item) => {
    const hay = `${item.label || ""} ${item.preview || ""}`.toLowerCase();
    return hay.includes(q);
  });
}

function getVisibleStructureItems(structure) {
  return searchStructureOutline(filterStructureOutline(structure));
}

function resetStructurePageIfNeeded(totalItems) {
  const totalPages = Math.max(1, Math.ceil(totalItems / structurePageSize));
  if (structurePage >= totalPages) structurePage = totalPages - 1;
  if (structurePage < 0) structurePage = 0;
}

function syncStructurePager(filteredCount) {
  const totalPages = Math.max(1, Math.ceil(filteredCount / structurePageSize));
  resetStructurePageIfNeeded(filteredCount);
  const start = filteredCount ? structurePage * structurePageSize + 1 : 0;
  const end = Math.min(filteredCount, (structurePage + 1) * structurePageSize);
  if (structurePageInfoEl) {
    structurePageInfoEl.textContent = filteredCount
      ? t("structurePageInfo", { start, end, total: filteredCount, page: structurePage + 1, pages: totalPages })
      : (structureSearchQuery.trim() ? t("structureSearchNoMatch") : t("structureNoFilterMatch"));
  }
  if (structurePrevPageBtn) structurePrevPageBtn.disabled = structurePage <= 0;
  if (structureNextPageBtn) structureNextPageBtn.disabled = structurePage >= totalPages - 1;
}

function syncStructureQuickEdit(item) {
  if (!structureQuickEditEl || !structureQuickEditTextEl) return;
  if (!item || item.type === "table" || !Number.isFinite(item.paraIndex)) {
    structureQuickEditEl.classList.add("hidden");
    structureQuickEditTextEl.value = "";
    if (structureQuickApplyBtn) structureQuickApplyBtn.disabled = true;
    if (structureQuickFocusBtn) structureQuickFocusBtn.disabled = item?.type === "table";
    return;
  }
  structureQuickEditEl.classList.remove("hidden");
  const emptyLabel = t("structureEmptyPara");
  structureQuickEditTextEl.value = item.label === emptyLabel ? "" : item.label;
  if (structureQuickApplyBtn) structureQuickApplyBtn.disabled = false;
  if (structureQuickFocusBtn) structureQuickFocusBtn.disabled = false;
}

function jumpToStructureItem(item, options = {}) {
  if (!item?.el) return;
  clearStructureHighlights();
  structureSelectionId = item.id;
  item.el.classList.add("search-hit", "search-hit-active");
  item.el.scrollIntoView({ behavior: "smooth", block: "center" });
  if (!options.silentSelect) {
    syncStructureQuickEdit(item);
    renderStructureOutline(documentStructure);
  }
}

function enableEditModeForStructure() {
  if (!readOnlyMode || !readModeEl) return;
  readModeEl.checked = false;
  readOnlyMode = false;
  const label = readModeEl.closest("label")?.querySelector("span");
  if (label) {
    label.dataset.i18n = "readModeOff";
    label.textContent = t("readModeOff");
  }
  syncInlineEditMode();
}

function focusStructureItemInPreview(item) {
  if (!item) return;
  jumpToStructureItem(item);
  if (item.type === "table" || !Number.isFinite(item.paraIndex)) return;
  enableEditModeForStructure();
  focusParagraphStart(item.paraIndex);
}

async function copyStructureItemText(item) {
  if (!item?.label) return;
  const emptyLabel = t("structureEmptyPara");
  const text = item.label === emptyLabel ? "" : item.label;
  try {
    await navigator.clipboard.writeText(text);
    toast(t("structureCopied"), "success");
  } catch (_) {
    toast(t("structureCopyFailed"), "error");
  }
}

async function applyStructureQuickEdit() {
  const item = getStructureSelection();
  if (!item || item.type === "table" || !Number.isFinite(item.paraIndex)) return;
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  const newText = structureQuickEditTextEl?.value ?? "";
  const count = await applyDocumentEdit({
    op: "paragraphBatch",
    items: [{ index: item.paraIndex, text: newText }],
  });
  if (count > 0) {
    toast(t("editApplied", { count }), "success");
    structureSelectionId = item.id;
    if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
  } else {
    toast(t("editNothing"), "info");
  }
}

function renderStructureStats(structure) {
  if (!structureSummaryEl) return;
  structureSummaryEl.replaceChildren();
  if (!structure) return;

  const stats = [
    [t("words"), structure.words],
    [t("chars"), structure.chars],
    [t("headings"), structure.headings.length],
    [t("paragraphs"), structure.paragraphs],
    [t("tables"), structure.tables],
  ];
  const grid = document.createElement("div");
  grid.className = "structure-stats";
  stats.forEach(([label, value]) => {
    const item = document.createElement("div");
    item.className = "structure-stat";
    item.innerHTML = `<span class="structure-stat-label">${label}</span><strong>${value}</strong>`;
    grid.appendChild(item);
  });
  structureSummaryEl.appendChild(grid);
}

function renderStructureOutline(structure) {
  if (!structureOutlineEl) return;
  structureOutlineEl.replaceChildren();
  if (!structure?.outline?.length) {
    const empty = document.createElement("p");
    empty.className = "hint";
    empty.textContent = t("structureNoOutline");
    syncStructurePager(0);
    return;
  }

  const filtered = getVisibleStructureItems(structure);
  syncStructurePager(filtered.length);

  if (!filtered.length) {
    const empty = document.createElement("p");
    empty.className = "hint";
    empty.textContent = structureSearchQuery.trim() ? t("structureSearchNoMatch") : t("structureNoFilterMatch");
    structureOutlineEl.appendChild(empty);
    return;
  }

  const pageStart = structurePage * structurePageSize;
  const slice = filtered.slice(pageStart, pageStart + structurePageSize);

  const list = document.createElement("div");
  list.className = "structure-outline-list";
  list.style.setProperty("--structure-page-size", String(structurePageSize));

  slice.forEach((item) => {
    const row = document.createElement("div");
    row.className = "structure-outline-item";
    if (item.id === structureSelectionId) row.classList.add("is-active");
    if (item.type === "heading" && item.level) row.style.setProperty("--struct-indent", `${Math.max(0, item.level - 1) * 10}px`);

    const mainBtn = document.createElement("button");
    mainBtn.type = "button";
    mainBtn.className = "structure-outline-main";
    const badge = document.createElement("span");
    badge.className = `structure-type-badge structure-type-${item.type}`;
    badge.textContent = structureTypeLabel(item.type);
    const text = document.createElement("span");
    text.className = "structure-outline-text";
    text.textContent = item.preview || item.label;
    mainBtn.append(badge, text);
    mainBtn.addEventListener("click", () => jumpToStructureItem(item));

    const actions = document.createElement("div");
    actions.className = "structure-outline-actions";

    const copyBtn = document.createElement("button");
    copyBtn.type = "button";
    copyBtn.className = "btn structure-action-btn";
    copyBtn.textContent = t("structureCopy");
    copyBtn.title = t("structureCopy");
    copyBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      copyStructureItemText(item);
    });

    const focusBtn = document.createElement("button");
    focusBtn.type = "button";
    focusBtn.className = "btn structure-action-btn";
    focusBtn.textContent = t("structureEdit");
    focusBtn.title = t("structureEdit");
    focusBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      focusStructureItemInPreview(item);
    });

    actions.append(copyBtn, focusBtn);
    row.append(mainBtn, actions);
    list.appendChild(row);
  });

  structureOutlineEl.appendChild(list);
}

function renderStructurePanel(structure) {
  renderStructureStats(structure);
  if (!structure) {
    structureSelectionId = null;
    structurePage = 0;
    if (structureOutlineEl) structureOutlineEl.replaceChildren();
    syncStructurePager(0);
    syncStructureQuickEdit(null);
    return;
  }
  if (structureSelectionId && !structure.outline?.some((i) => i.id === structureSelectionId)) {
    structureSelectionId = null;
  }
  renderStructureOutline(structure);
  syncStructureQuickEdit(getStructureSelection());
}

function resetStructurePanelUi() {
  structureSelectionId = null;
  structurePage = 0;
  structureSearchQuery = "";
  if (structureSearchEl) structureSearchEl.value = "";
  clearStructureHighlights();
  if (structureOutlineEl) structureOutlineEl.replaceChildren();
  if (structureSummaryEl) structureSummaryEl.replaceChildren();
  syncStructurePager(0);
  syncStructureQuickEdit(null);
}

function onStructureListControlsChange() {
  structurePage = 0;
  renderStructureOutline(documentStructure);
}

function wireStructurePanel() {
  structureFilterEl?.addEventListener("change", () => {
    structureFilter = structureFilterEl.value || "all";
    onStructureListControlsChange();
  });
  structureSearchEl?.addEventListener("input", () => {
    structureSearchQuery = structureSearchEl.value || "";
    onStructureListControlsChange();
  });
  structurePageSizeEl?.addEventListener("change", () => {
    structurePageSize = Math.max(10, Math.min(40, Number(structurePageSizeEl.value) || 20));
    localStorage.setItem(STRUCTURE_PAGE_SIZE_KEY, String(structurePageSize));
    structurePage = 0;
    renderStructureOutline(documentStructure);
  });
  structurePrevPageBtn?.addEventListener("click", () => {
    if (structurePage <= 0) return;
    structurePage--;
    renderStructureOutline(documentStructure);
  });
  structureNextPageBtn?.addEventListener("click", () => {
    const filtered = getVisibleStructureItems(documentStructure);
    const totalPages = Math.max(1, Math.ceil(filtered.length / structurePageSize));
    if (structurePage >= totalPages - 1) return;
    structurePage++;
    renderStructureOutline(documentStructure);
  });
  structureQuickApplyBtn?.addEventListener("click", () => { applyStructureQuickEdit().catch(() => {}); });
  structureQuickFocusBtn?.addEventListener("click", () => focusStructureItemInPreview(getStructureSelection()));
}

wireStructurePanel();
