// Metadata panel — read/edit docProps/core.xml (title, author, keywords).

const metaTitleEl = document.getElementById("metaTitle");
const metaCreatorEl = document.getElementById("metaCreator");
const metaKeywordsEl = document.getElementById("metaKeywords");
const metaReloadBtn = document.getElementById("metaReloadBtn");
const metaApplyBtn = document.getElementById("metaApplyBtn");
const metaStatusEl = document.getElementById("metaStatus");

let metadataBaseline = null;

function collectMetadataForm() {
  return {
    title: metaTitleEl?.value?.trim() ?? "",
    creator: metaCreatorEl?.value?.trim() ?? "",
    keywords: metaKeywordsEl?.value?.trim() ?? "",
  };
}

function metadataEquals(a, b) {
  if (!a || !b) return false;
  return a.title === b.title && a.creator === b.creator && a.keywords === b.keywords;
}

function fillMetadataForm(data) {
  if (metaTitleEl) metaTitleEl.value = data?.title ?? "";
  if (metaCreatorEl) metaCreatorEl.value = data?.creator ?? "";
  if (metaKeywordsEl) metaKeywordsEl.value = data?.keywords ?? "";
}

function syncMetadataStatus() {
  if (!metaStatusEl) return;
  if (!originalFileBytes) {
    metaStatusEl.textContent = "";
    return;
  }
  const current = collectMetadataForm();
  if (!metadataBaseline) {
    metaStatusEl.textContent = t("metadataHint");
    return;
  }
  metaStatusEl.textContent = metadataEquals(current, metadataBaseline)
    ? t("metadataSynced")
    : t("metadataDirty");
}

async function loadMetadataFromDocument() {
  if (!originalFileBytes) {
    metadataBaseline = null;
    fillMetadataForm({ title: "", creator: "", keywords: "" });
    syncMetadataStatus();
    return;
  }
  const data = await extractCoreMetadataFromDocx(originalFileBytes);
  metadataBaseline = { ...data };
  fillMetadataForm(data);
  syncMetadataStatus();
}

async function applyMetadataEdit() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  const fields = collectMetadataForm();
  if (metadataBaseline && metadataEquals(fields, metadataBaseline)) {
    toast(t("editNothing"), "info");
    return;
  }
  const count = await applyDocumentEdit({ op: "coreMetadata", fields });
  if (count > 0) {
    metadataBaseline = { ...fields };
    syncMetadataStatus();
    toast(t("metadataApplied"), "success");
    if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
  } else {
    metadataBaseline = { ...fields };
    syncMetadataStatus();
    toast(t("editNothing"), "info");
  }
}

function wireMetadataPanel() {
  metaReloadBtn?.addEventListener("click", () => { loadMetadataFromDocument().catch(() => {}); });
  metaApplyBtn?.addEventListener("click", () => { applyMetadataEdit().catch(() => {}); });
  [metaTitleEl, metaCreatorEl, metaKeywordsEl].forEach((el) => {
    el?.addEventListener("input", syncMetadataStatus);
  });
}

wireMetadataPanel();
