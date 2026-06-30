// Bulk edit tools — find/replace on document XML (ZIP-patch pipeline).

async function applyEditTool() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  const find = (editFindEl?.value || "").trim();
  if (!find) {
    toast(t("editErrNoFind"), "warning");
    return;
  }
  const replace = editReplaceEl?.value ?? "";
  const scope = editScopeEl?.value || "all";
  const regex = !!editRegexEl?.checked;

  const count = await applyDocumentEdit({ find, replace, regex, scope });
  if (count > 0) toast(t("editApplied", { count }), "success");
  else toast(t("editNothing"), "info");
}

if (applyEditToolBtnEl) {
  applyEditToolBtnEl.addEventListener("click", () => applyEditTool());
}
