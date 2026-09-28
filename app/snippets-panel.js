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
const snExportBtn = document.getElementById("snExportBtn");
const snImportBtn = document.getElementById("snImportBtn");
const snImportFile = document.getElementById("snImportFile");
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

let snFilter = "";

// Lista: nazwa + podgląd treści, „Wstaw” (w miejscu kursora), „Edytuj” (do formularza), „Usuń”.
// Filtr przy dłuższej liście; pusta lista proponuje zestaw startowy.
function renderSnippetList() {
  if (!snListEl) return;
  snListEl.replaceChildren();
  const list = loadSnippets();
  if (!list.length) {
    const empty = document.createElement("p");
    empty.className = "hint";
    empty.textContent = t("snippetsEmpty");
    const starter = Object.assign(document.createElement("button"), { type: "button", className: "btn", id: "snStarterBtn", textContent: t("snippetsStarter") });
    starter.addEventListener("click", () => {
      const n = addStarterSnippets();
      renderSnippetList();
      toast(t("snippetsStarterAdded", { count: n }), "success");
    });
    snListEl.append(empty, starter);
    return;
  }
  if (list.length > 5) {
    const filter = Object.assign(document.createElement("input"), { type: "search", className: "sn-filter", placeholder: t("snippetsFilter"), value: snFilter });
    filter.addEventListener("input", () => { snFilter = filter.value; renderRows(); });
    snListEl.append(filter);
  }
  const rows = document.createElement("div");
  rows.className = "stack stack-tight";
  snListEl.append(rows);
  function renderRows() {
    rows.replaceChildren();
    const q = snFilter.trim().toLocaleLowerCase("pl-PL");
    list.filter((sn) => !q || sn.name.toLocaleLowerCase("pl-PL").includes(q) || sn.body.toLocaleLowerCase("pl-PL").includes(q)).forEach((sn) => {
      const row = document.createElement("div");
      row.className = "snippet-row";
      const info = document.createElement("div");
      info.className = "snippet-info";
      const name = Object.assign(document.createElement("span"), { className: "snippet-name", textContent: formatSnippetTrigger(sn.name) });
      const pv = Object.assign(document.createElement("span"), { className: "snippet-preview", textContent: sn.body.replace(/\s*\n\s*/g, " ⏎ ") });
      info.append(name, pv);
      const mk = (label, cls, fn) => {
        const b = Object.assign(document.createElement("button"), { type: "button", className: `btn ${cls}`, textContent: label });
        // przycisk nie zabiera fokusu z tekstu (kursor zostaje tam, gdzie był)
        b.addEventListener("mousedown", (e) => e.preventDefault());
        b.addEventListener("click", fn);
        return b;
      };
      const insertBtn = mk(t("snippetsInsertShort"), "snippet-insert-btn", () => insertSnippetAtCaret(sn));
      const editBtn = mk(t("snippetsEdit"), "snippet-use-btn", () => {
        if (snNameEl) snNameEl.value = sn.name;
        if (snBodyEl) snBodyEl.value = sn.body;
        snBodyEl?.focus();
      });
      const delBtn = mk("✕", "snippet-del-btn", () => {
        if (!confirm(t("snippetsDeleteConfirm", { name: formatSnippetTrigger(sn.name) }))) return;
        deleteSnippet(sn.name);
        renderSnippetList();
      });
      delBtn.setAttribute("aria-label", t("snippetsDelete"));
      const actions = document.createElement("div");
      actions.className = "snippet-actions";
      actions.append(insertBtn, editBtn, delBtn);
      row.append(info, actions);
      rows.appendChild(row);
    });
  }
  renderRows();
}

async function runSnippetScan() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "error");
    return;
  }
  // tekst wpisany w podglądzie (np. właśnie wstawione !nazwa) musi być w pliku, zanim go przeczytamy
  if (typeof mergeInlineEditsIntoBytes === "function") await mergeInlineEditsIntoBytes();
  snippetScan = await scanSnippetTriggers(originalFileBytes);
  syncSnippetStatus();
}

function saveSnippetFromForm() {
  const entry = upsertSnippet(snNameEl?.value, snBodyEl?.value);
  if (!entry) {
    toast(t("snippetsSaveInvalid"), "error");
    return;
  }
  toast(t(entry.existed ? "snippetsUpdated" : "snippetsSaved", { name: formatSnippetTrigger(entry.name) }), "success");
  renderSnippetList();
  if (originalFileBytes) runSnippetScan();
}

async function expandSnippetsInDocument() {
  if (!originalFileBytes) return;
  await runSnippetScan(); // zawsze świeżo — dawniej przycisk był wyszarzony do ręcznego „Skanuj”
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

// sn — z listy; bez sn: z formularza (zapisany pod tą nazwą albo sama wpisana treść).
async function insertSnippetAtCaret(sn) {
  if (readOnlyMode) {
    toast(t("readModeOn"), "info");
    return;
  }
  if (!sn || !sn.name) {
    const name = normalizeSnippetName(snNameEl?.value);
    const stored = loadSnippets().find((s) => s.name === name);
    sn = stored || { name: name || "snippet", body: snBodyEl?.value || "" };
  }
  if (!String(sn.body || "").trim()) {
    toast(t("snippetsInsertEmpty"), "error");
    return;
  }
  const p = restoreDocCaret(); // kliknięcie w pole/przycisk zabrało fokus — wracamy do kursora
  if (!p) {
    toast(t("snippetsInsertNoCaret"), "info");
    return;
  }
  if (await expandSnippetAtCaret(p, sn, 0)) toast(t("snippetsInserted"), "success");
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
  const p = restoreDocCaret(); // kliknięcie w pole/przycisk zabrało fokus — wracamy do kursora
  if (!p) {
    toast(t("snippetsInsertNoCaret"), "info");
    return;
  }
  const trigger = formatSnippetTrigger(name);
  const style = mergeRunStyles(getInheritedRunStyleAtCaret(p), activeTypingStyle);
  asUndoStep("undoOpInsert", () => {
    if (runStyleHasProps(style)) insertStyledTextAtCaret(trigger, style, p);
    else insertTextAtCaret(trigger);
  });
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
  snInsertBtn?.addEventListener("click", () => insertSnippetAtCaret());
  snInsertTriggerBtn?.addEventListener("click", insertSnippetTriggerAtCaret);
  snExpandModeEl?.addEventListener("change", () => {
    setSnippetExpandMode(snExpandModeEl.value);
    toast(t("snippetsExpandModeSaved"), "success");
  });
  syncSnippetExpandModeSelect();
  renderSnippetList();
}

wireSnippetsPanel();

snExportBtn?.addEventListener("click", () => {
  const data = buildSnippetsExport();
  if (!data.items.length) { toast(t("snippetsEmpty"), "info"); return; }
  downloadJsonFile(data, "snippety-documents-workbench.json");
  toast(t("snExported", { count: data.items.length }), "success");
});

snImportBtn?.addEventListener("click", async () => {
  const data = await pickJsonFile(snImportFile);
  if (data === null) return;
  const res = data === undefined ? null : importSnippetsData(data);
  if (!res) { toast(t("importBadFile"), "error"); return; }
  renderSnippetList();
  toast(t("snImported", res), "success");
});
