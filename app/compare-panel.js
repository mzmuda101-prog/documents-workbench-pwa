// Local, read-only counterpart of Word's Compare Documents.  The opened document is the
// current side; a picked .docx is the reference side.  Nothing is uploaded or patched.
// Which side is newer is the user's choice (#compareNewer, like Word's original → revised):
// „Dodane” always means text of the newer version, „Usunięte” — of the older one.

const compareFileEl = document.getElementById("compareFile");
const compareRunBtn = document.getElementById("compareRunBtn");
const compareStatusEl = document.getElementById("compareStatus");
const compareSummaryEl = document.getElementById("compareSummary");
const compareResultsEl = document.getElementById("compareResults");
const compareNewerEl = document.getElementById("compareNewer");
const compareRedlineBtn = document.getElementById("compareRedlineBtn");
const COMPARE_RENDER_LIMIT = 220;
let compareLast = null; // { current, reference, rows } — wyrównanie po stronie „otwarty = nowszy”
function compareNewerSide() {
  return compareNewerEl?.querySelector("button.is-on")?.dataset.newer === "reference" ? "reference" : "current";
}

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
  // starsza → nowsza: przy „nowszy = drugi plik” zmiana idzie od otwartego do drugiego
  const newerIsCurrent = compareNewerSide() === "current";
  const textOf = (row, side) => (side === "current" ? current[row.left] : reference[row.right]);
  shown.slice(0, COMPARE_RENDER_LIMIT).forEach((row) => {
    const item = document.createElement("article"); item.className = `compare-item compare-item-${row.kind}`;
    item.append(Object.assign(document.createElement("span"), { className: "compare-kind", textContent: t(`compareKind${row.kind[0].toUpperCase()}${row.kind.slice(1)}`) }));
    if (row.kind === "changed") item.append(newerIsCurrent ? compareChangedLine(reference[row.right], current[row.left]) : compareChangedLine(current[row.left], reference[row.right]));
    else item.append(compareTextLine(textOf(row, row.side), row.kind));
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
    const refBytes = new Uint8Array(await file.arrayBuffer());
    const [current, reference] = await Promise.all([dwbCompareParagraphs(source), dwbCompareParagraphs(refBytes)]);
    const result = dwbCompareAlign(current, reference);
    compareLast = { current, reference, result, source, refBytes, refName: file.name };
    renderCompareLast();
  } catch (err) {
    compareSetStatus(t("compareFailed"));
  } finally { compareRunBtn.disabled = false; }
}

// Alignment names additions relative to its right side (the picked reference): „added” = only in
// the open document. Each one-sided row remembers its side; the kind follows the chosen direction.
function renderCompareLast() {
  if (!compareLast) return;
  const { current, reference, result } = compareLast;
  const newer = compareNewerSide();
  const rows = result.rows.map((row) => {
    if (row.kind !== "added" && row.kind !== "removed") return row;
    const side = row.kind === "added" ? "reference" : "current";
    return { ...row, side, kind: side === newer ? "added" : "removed" };
  });
  renderComparison(current, reference, { ...result, rows });
  if (compareRedlineBtn) compareRedlineBtn.hidden = !rows.some((r) => r.kind !== "same");
  compareSetStatus(t(result.approximate ? "compareDoneApprox" : "compareDone", { current: current.length, reference: reference.length }));
}

// Dokument z poprawkami (w:ins / w:del) — starsza → nowsza, w nowej karcie: przegląd w Recenzji
// (Akceptuj/Odrzuć), zapis jak zwykłego pliku. Żaden z porównywanych plików się nie zmienia.
async function createRedlineDocument() {
  if (!compareLast?.source || !compareLast?.refBytes) return;
  compareRedlineBtn.disabled = true;
  compareSetStatus(t("compareRedlineWorking"));
  try {
    const newerIsCurrent = compareNewerSide() === "current";
    const older = newerIsCurrent ? compareLast.refBytes : compareLast.source;
    const newer = newerIsCurrent ? compareLast.source : compareLast.refBytes;
    let author = "";
    try { author = localStorage.getItem("dwb.authorName") || ""; } catch (_) {}
    const { bytes, stats } = await dwbRedline(older, newer, { author: author || t("compareRedlineAuthor") });
    const base = (newerIsCurrent ? currentFileName : compareLast.refName || "dokument").replace(/\.docx$/i, "");
    const name = `${base} (${t("compareRedlineSuffix")}).docx`;
    const ok = await openDocumentFiles([{ file: new File([bytes], name, { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }) }]);
    compareSetStatus(ok ? t("compareRedlineDone", { n: stats.changed + stats.added + stats.removed }) : t("compareFailed"));
  } catch (err) {
    log(`Porównanie: ${err.message || err}`, "error");
    compareSetStatus(t("compareFailed"));
  } finally { compareRedlineBtn.disabled = false; }
}

compareRunBtn?.addEventListener("click", runDocumentComparison);
compareRedlineBtn?.addEventListener("click", createRedlineDocument);
compareNewerEl?.addEventListener("click", (e) => {
  const b = e.target.closest("button[data-newer]");
  if (!b || b.classList.contains("is-on")) return;
  compareNewerEl.querySelectorAll("button[data-newer]").forEach((x) => { const on = x === b; x.classList.toggle("is-on", on); x.setAttribute("aria-checked", String(on)); });
  renderCompareLast();
});
