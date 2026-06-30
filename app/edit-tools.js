// Bulk edit tools — find/replace, case, trim, affix on document paragraphs (ZIP-patch).

function syncEditToolFields() {
  if (!editOpEl) return;
  const op = editOpEl.value;
  if (editReplaceFieldsEl) editReplaceFieldsEl.classList.toggle("hidden", op !== "replace");
  if (editCaseFieldsEl) editCaseFieldsEl.classList.toggle("hidden", op !== "case");
  if (editTrimFieldsEl) editTrimFieldsEl.classList.toggle("hidden", op !== "trim");
  if (editAffixFieldsEl) editAffixFieldsEl.classList.toggle("hidden", op !== "affix");
}

function buildEditTransform() {
  const op = editOpEl?.value;
  if (op === "replace") {
    const find = editFindEl?.value ?? "";
    if (!find) return { ok: false, err: "editErrNoFind" };
    const repl = editReplaceEl?.value ?? "";
    if (editRegexEl?.checked) {
      try { new RegExp(find, "g"); } catch { return { ok: false, err: "editErrBadRegex" }; }
    }
    return { ok: true, edit: { op: "replace", find, replace: repl, regex: !!editRegexEl?.checked, scope: editScopeEl?.value || "all" } };
  }
  if (op === "case") {
    return { ok: true, edit: { op: "case", mode: editCaseModeEl?.value || "upper", scope: editScopeEl?.value || "all" } };
  }
  if (op === "trim") {
    return { ok: true, edit: { op: "trim", mode: editTrimModeEl?.value || "ends", scope: editScopeEl?.value || "all" } };
  }
  if (op === "affix") {
    const prefix = editPrefixEl?.value ?? "";
    const suffix = editSuffixEl?.value ?? "";
    if (!prefix && !suffix) return { ok: false, err: "editErrNoAffix" };
    return { ok: true, edit: { op: "affix", prefix, suffix, scope: editScopeEl?.value || "all" } };
  }
  return { ok: false, err: "editToolNoChange" };
}

async function applyEditTool() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  const tr = buildEditTransform();
  if (!tr.ok) {
    toast(t(tr.err), "warning");
    return;
  }
  const count = await applyDocumentEdit(tr.edit);
  if (count > 0) toast(t("editApplied", { count }), "success");
  else toast(t("editNothing"), "info");
}

if (applyEditToolBtnEl) {
  applyEditToolBtnEl.addEventListener("click", () => applyEditTool());
}
if (editOpEl) {
  editOpEl.addEventListener("change", syncEditToolFields);
  syncEditToolFields();
}
