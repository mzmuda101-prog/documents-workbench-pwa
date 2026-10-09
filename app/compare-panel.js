// Local, read-only counterpart of Word's Compare Documents.  The opened document is the
// current side; a picked .docx is the reference side.  Nothing is uploaded or patched.

const compareFileEl = document.getElementById("compareFile");
const compareRunBtn = document.getElementById("compareRunBtn");
const compareStatusEl = document.getElementById("compareStatus");
const compareSummaryEl = document.getElementById("compareSummary");
const compareResultsEl = document.getElementById("compareResults");
const COMPARE_RENDER_LIMIT = 220;

function compareSetStatus(message) { if (compareStatusEl) compareStatusEl.textContent = message || ""; }
function compareCell(label, value, tone) {
  const el = document.createElement("div"); el.className = `compare-cell ${tone || ""}`;
  el.append(Object.assign(document.createElement("strong"), { textContent: value }), Object.assign(document.createElement("span"), { textContent: label }));
  return el;
}
function compareTextLine(text, kind) {
  const el = document.createElement("div"); el.className = `compare-line compare-${kind}`;
  el.textContent = text; return el;
}
function compareChangedLine(before, after) {
  const el = document.createElement("div"); el.className = "compare-line compare-changed";
  dwbCompareWords(before, after).forEach((part) => {
    const node = document.createElement(part.kind === "same" ? "span" : "mark");
    node.className = `compare-word compare-${part.kind}`; node.textContent = part.text; el.append(node);
  });
  return el;
}

function renderComparison(current, reference, result) {
  compareSummaryEl?.replaceChildren(); compareResultsEl?.replaceChildren();
  const counts = result.rows.reduce((out, r) => { out[r.kind] = (out[r.kind] || 0) + 1; return out; }, {});
  compareSummaryEl?.append(
    compareCell(t("compareAdded"), String(counts.added || 0), "added"),
    compareCell(t("compareRemoved"), String(counts.removed || 0), "removed"),
    compareCell(t("compareChanged"), String(counts.changed || 0), "changed")
  );
  const shown = result.rows.filter((r) => r.kind !== "same");
  if (!shown.length) {
    compareResultsEl?.append(Object.assign(document.createElement("p"), { className: "hint", textContent: t("compareSame") }));
    return;
  }
  shown.slice(0, COMPARE_RENDER_LIMIT).forEach((row) => {
    const item = document.createElement("article"); item.className = `compare-item compare-item-${row.kind}`;
    item.append(Object.assign(document.createElement("span"), { className: "compare-kind", textContent: t(`compareKind${row.kind[0].toUpperCase()}${row.kind.slice(1)}`) }));
    if (row.kind === "changed") item.append(compareChangedLine(reference[row.right], current[row.left]));
    else if (row.kind === "added") item.append(compareTextLine(current[row.left], "added"));
    else item.append(compareTextLine(reference[row.right], "removed"));
    compareResultsEl?.append(item);
  });
  if (shown.length > COMPARE_RENDER_LIMIT) compareResultsEl?.append(Object.assign(document.createElement("p"), { className: "hint", textContent: t("compareTruncated", { shown: COMPARE_RENDER_LIMIT, total: shown.length }) }));
}

async function runDocumentComparison() {
  if (!originalFileBytes) { compareSetStatus(t("noFileToSave")); return; }
  const file = compareFileEl?.files?.[0];
  if (!file) { compareSetStatus(t("comparePickFile")); return; }
  if (!/\.docx$/i.test(file.name)) { compareSetStatus(t("compareBadFile")); return; }
  compareRunBtn.disabled = true; compareSetStatus(t("compareWorking"));
  try {
    // Include local, unsaved edits without mutating app state or the original package.
    const source = pendingDocEdits.length ? (await buildPatchedDocx(originalFileBytes, pendingDocEdits)).bytes : originalFileBytes;
    const [current, reference] = await Promise.all([dwbCompareParagraphs(source), dwbCompareParagraphs(await file.arrayBuffer())]);
    const result = dwbCompareAlign(current, reference);
    // Alignment names additions relative to its right side (the picked reference). The UI
    // describes what changed in the open document, so flip one-sided rows here.
    result.rows = result.rows.map((row) => row.kind === "added" ? { ...row, kind: "removed" } : row.kind === "removed" ? { ...row, kind: "added" } : row);
    renderComparison(current, reference, result);
    compareSetStatus(t(result.approximate ? "compareDoneApprox" : "compareDone", { current: current.length, reference: reference.length }));
  } catch (err) {
    compareSetStatus(t("compareFailed"));
  } finally { compareRunBtn.disabled = false; }
}

compareRunBtn?.addEventListener("click", runDocumentComparison);
