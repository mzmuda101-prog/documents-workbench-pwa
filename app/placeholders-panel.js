// Placeholders panel — scan {{field}} tokens, form fill, apply to XML.

let placeholderScan = null;

const phScanBtn = document.getElementById("phScanBtn");
const phFillBtn = document.getElementById("phFillBtn");
const phClearBtn = document.getElementById("phClearBtn");
const phScopeEl = document.getElementById("phScope");
const phStatusEl = document.getElementById("phStatus");
const phFormEl = document.getElementById("phForm");

function escapePhHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function syncPlaceholderStatus() {
  if (!phStatusEl) return;
  if (!placeholderScan?.fields?.length) {
    phStatusEl.textContent = placeholderScan ? t("placeholdersNoFields") : "";
    return;
  }
  phStatusEl.textContent = t("placeholdersFound", {
    count: placeholderScan.fields.length,
    total: placeholderScan.total,
  });
}

function renderPlaceholderForm() {
  if (!phFormEl) return;
  phFormEl.replaceChildren();
  if (!placeholderScan?.fields?.length) return;

  placeholderScan.fields.forEach((field) => {
    const label = document.createElement("label");
    label.className = "field";
    const title = document.createElement("span");
    title.textContent = `${t("placeholdersFieldLabel")}: ${placeholderFieldLabel(field.name)}`;
    const input = document.createElement("input");
    input.type = "text";
    input.dataset.phName = field.name;
    input.placeholder = field.samples[0]?.token || `{{${field.name}}}`;
    input.autocomplete = "off";
    label.append(title, input);
    phFormEl.appendChild(label);

    if (field.samples.length) {
      const hint = document.createElement("p");
      hint.className = "hint ph-field-hint";
      hint.textContent = `${field.count}× · ${field.samples.map((s) => escapePhHtml(s.token)).join(", ")}`;
      phFormEl.appendChild(hint);
    }
  });
}

function collectPlaceholderValues() {
  const values = {};
  phFormEl?.querySelectorAll("input[data-ph-name]").forEach((input) => {
    const name = input.dataset.phName;
    const val = input.value.trim();
    if (name && val) values[name] = val;
  });
  return values;
}

async function runPlaceholderScan() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "error");
    return;
  }
  placeholderScan = await scanPlaceholders(originalFileBytes);
  syncPlaceholderStatus();
  renderPlaceholderForm();
  if (phFillBtn) phFillBtn.disabled = !placeholderScan.fields.length;
}

function clearPlaceholderPanel() {
  placeholderScan = null;
  syncPlaceholderStatus();
  if (phFormEl) phFormEl.replaceChildren();
  if (phFillBtn) phFillBtn.disabled = true;
}

async function applyPlaceholderFill() {
  if (!originalFileBytes || !placeholderScan?.fields?.length) return;
  const values = collectPlaceholderValues();
  const keys = Object.keys(values);
  if (!keys.length) {
    toast(t("placeholdersEmptyValues"), "error");
    return;
  }
  if (!confirm(t("placeholdersFillConfirm", { count: keys.length }))) return;

  const count = await applyDocumentEdit({
    op: "placeholderFill",
    values,
    scope: phScopeEl?.value || "all",
  });
  if (count > 0) {
    toast(t("placeholdersApplied", { count }), "success");
    await runPlaceholderScan();
    if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
  } else {
    toast(t("editNothing"), "info");
  }
}

function wirePlaceholdersPanel() {
  phScanBtn?.addEventListener("click", runPlaceholderScan);
  phFillBtn?.addEventListener("click", applyPlaceholderFill);
  phClearBtn?.addEventListener("click", clearPlaceholderPanel);
  if (phFillBtn) phFillBtn.disabled = true;
}

wirePlaceholdersPanel();
