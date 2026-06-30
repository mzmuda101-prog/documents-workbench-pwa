// Snippets panel — !name triggers, localStorage, expand in document.

let snippetScan = null;

const snNameEl = document.getElementById("snName");
const snBodyEl = document.getElementById("snBody");
const snSaveBtn = document.getElementById("snSaveBtn");
const snScanBtn = document.getElementById("snScanBtn");
const snExpandBtn = document.getElementById("snExpandBtn");
const snInsertBtn = document.getElementById("snInsertBtn");
const snInsertTriggerBtn = document.getElementById("snInsertTriggerBtn");
const snExpandModeEl = document.getElementById("snExpandMode");
const snScopeEl = document.getElementById("snScope");
const snStatusEl = document.getElementById("snStatus");
const snListEl = document.getElementById("snList");

function escapeSnHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function syncSnippetStatus() {
  if (!snStatusEl) return;
  if (!snippetScan?.triggers?.length) {
    snStatusEl.textContent = snippetScan ? t("snippetsNoTriggers") : "";
    return;
  }
  const missing = snippetScan.triggers.filter((tr) => !tr.hasDefinition).length;
  snStatusEl.textContent = missing
    ? t("snippetsFoundMissing", { count: snippetScan.triggers.length, total: snippetScan.total, missing })
    : t("snippetsFound", { count: snippetScan.triggers.length, total: snippetScan.total });
}

function renderSnippetList() {
  if (!snListEl) return;
  snListEl.replaceChildren();
  const list = loadSnippets();
  if (!list.length) {
    const empty = document.createElement("p");
    empty.className = "hint";
    empty.textContent = t("snippetsEmpty");
    snListEl.appendChild(empty);
    return;
  }
  list.forEach((sn) => {
    const row = document.createElement("div");
    row.className = "snippet-row";
    row.innerHTML = `<span class="snippet-name">${escapeSnHtml(formatSnippetTrigger(sn.name))}</span>`;
    const useBtn = document.createElement("button");
    useBtn.type = "button";
    useBtn.className = "btn snippet-use-btn";
    useBtn.textContent = t("snippetsUse");
    useBtn.addEventListener("click", () => {
      if (snNameEl) snNameEl.value = sn.name;
      if (snBodyEl) snBodyEl.value = sn.body;
    });
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "btn snippet-del-btn";
    delBtn.textContent = t("snippetsDelete");
    delBtn.addEventListener("click", () => {
      if (!confirm(t("snippetsDeleteConfirm", { name: formatSnippetTrigger(sn.name) }))) return;
      deleteSnippet(sn.name);
      renderSnippetList();
    });
    row.append(useBtn, delBtn);
    snListEl.appendChild(row);
  });
}

async function runSnippetScan() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "error");
    return;
  }
  snippetScan = await scanSnippetTriggers(originalFileBytes);
  syncSnippetStatus();
  if (snExpandBtn) snExpandBtn.disabled = !snippetScan.triggers.some((tr) => tr.hasDefinition);
}

function saveSnippetFromForm() {
  const entry = upsertSnippet(snNameEl?.value, snBodyEl?.value);
  if (!entry) {
    toast(t("snippetsSaveInvalid"), "error");
    return;
  }
  toast(t("snippetsSaved", { name: formatSnippetTrigger(entry.name) }), "success");
  renderSnippetList();
  if (originalFileBytes) runSnippetScan();
}

async function expandSnippetsInDocument() {
  if (!originalFileBytes) return;
  if (!snippetScan) await runSnippetScan();
  const expandMap = buildSnippetExpandMap(snippetScan?.triggers || [], snippetScan?.stored || snippetsToMap(loadSnippets()));
  const keys = Object.keys(expandMap);
  if (!keys.length) {
    toast(t("snippetsNothingToExpand"), "info");
    return;
  }
  if (!confirm(t("snippetsExpandConfirm", { count: keys.length }))) return;
  const count = await applyDocumentEdit({
    op: "snippetExpand",
    snippets: expandMap,
    scope: snScopeEl?.value || "all",
  });
  if (count > 0) {
    toast(t("snippetsExpanded", { count }), "success");
    await runSnippetScan();
    if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
  } else {
    toast(t("editNothing"), "info");
  }
}

function insertSnippetAtCaret() {
  if (readOnlyMode) {
    toast(t("readModeOn"), "info");
    return;
  }
  const name = normalizeSnippetName(snNameEl?.value);
  const list = loadSnippets();
  const sn = list.find((s) => s.name === name);
  const body = resolveSnippetBody(sn?.body || snBodyEl?.value);
  if (!body?.trim()) {
    toast(t("snippetsInsertEmpty"), "error");
    return;
  }
  const p = document.activeElement?.closest?.(".docx-editable-p");
  if (!p) {
    toast(t("snippetsInsertNoCaret"), "info");
    return;
  }
  p.focus();
  const style = mergeRunStyles(getInheritedRunStyleAtCaret(p), activeTypingStyle);
  if (runStyleHasProps(style)) insertStyledTextAtCaret(body, style, p);
  else insertTextAtCaret(body);
  onInlineParagraphInput();
  toast(t("snippetsInserted"), "success");
}

function insertSnippetTriggerAtCaret() {
  if (readOnlyMode) {
    toast(t("readModeOn"), "info");
    return;
  }
  const name = normalizeSnippetName(snNameEl?.value);
  if (!name) {
    toast(t("snippetsSaveInvalid"), "error");
    return;
  }
  const p = document.activeElement?.closest?.(".docx-editable-p");
  if (!p) {
    toast(t("snippetsInsertNoCaret"), "info");
    return;
  }
  p.focus();
  const trigger = formatSnippetTrigger(name);
  const style = mergeRunStyles(getInheritedRunStyleAtCaret(p), activeTypingStyle);
  if (runStyleHasProps(style)) insertStyledTextAtCaret(trigger, style, p);
  else insertTextAtCaret(trigger);
  onInlineParagraphInput();
  toast(t("snippetsTriggerInserted", { name: trigger }), "success");
}

function syncSnippetExpandModeSelect() {
  if (!snExpandModeEl) return;
  snExpandModeEl.value = getSnippetExpandMode();
}

function wireSnippetsPanel() {
  snSaveBtn?.addEventListener("click", saveSnippetFromForm);
  snScanBtn?.addEventListener("click", runSnippetScan);
  snExpandBtn?.addEventListener("click", expandSnippetsInDocument);
  snInsertBtn?.addEventListener("click", insertSnippetAtCaret);
  snInsertTriggerBtn?.addEventListener("click", insertSnippetTriggerAtCaret);
  snExpandModeEl?.addEventListener("change", () => {
    setSnippetExpandMode(snExpandModeEl.value);
    toast(t("snippetsExpandModeSaved"), "success");
  });
  if (snExpandBtn) snExpandBtn.disabled = true;
  syncSnippetExpandModeSelect();
  renderSnippetList();
}

wireSnippetsPanel();
