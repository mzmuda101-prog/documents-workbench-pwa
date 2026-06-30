// Find/Replace workbench — scan, navigate hits, preview, replace one / all.

let frMatches = [];
let frActiveIndex = -1;
let frPreview = null;

const frReplaceEl = document.getElementById("frReplace");
const frRegexEl = document.getElementById("frRegex");
const frScopeEl = document.getElementById("frScope");
const frPrevBtn = document.getElementById("frPrevBtn");
const frNextBtn = document.getElementById("frNextBtn");
const frScanBtn = document.getElementById("frScanBtn");
const frReplaceOneBtn = document.getElementById("frReplaceOneBtn");
const frReplaceAllBtn = document.getElementById("frReplaceAllBtn");
const frPreviewEl = document.getElementById("frPreview");
const frStatusEl = document.getElementById("frStatus");

function buildFrEdit() {
  const find = (searchQueryEl?.value || "").trim();
  const replace = frReplaceEl?.value ?? "";
  const regex = !!frRegexEl?.checked;
  const scope = frScopeEl?.value || "all";
  return { op: "replace", find, replace, regex, scope };
}

function validateFrFind(edit) {
  if (!edit.find) return { ok: false, err: "editErrNoFind" };
  if (edit.regex) {
    try { new RegExp(edit.find, "g"); } catch { return { ok: false, err: "editErrBadRegex" }; }
  }
  return { ok: true };
}

function collectFrDomMatches(find, scope, regex) {
  if (!find || !docCanvasEl) return [];
  const root = getDocContentRoot(docCanvasEl);
  if (!root) return [];
  const nodes = scope === "headings"
    ? root.querySelectorAll("h1, h2, h3, h4, h5, h6, p")
    : root.querySelectorAll("p, span, td, th, li");
  const matches = [];
  nodes.forEach((el) => {
    const text = el.textContent || "";
    if (!text) return;
    if (regex) {
      try {
        const re = new RegExp(find, "gi");
        let m;
        while ((m = re.exec(text)) !== null) {
          matches.push({ el, index: m.index, length: m[0].length, sample: m[0] });
          if (!m[0].length) re.lastIndex++;
        }
      } catch (_) { /* invalid */ }
      return;
    }
    const lower = text.toLowerCase();
    const q = find.toLowerCase();
    let idx = 0;
    while ((idx = lower.indexOf(q, idx)) !== -1) {
      matches.push({ el, index: idx, length: find.length, sample: text.slice(idx, idx + find.length) });
      idx += find.length || 1;
    }
  });
  return matches;
}

function clearFrHighlights() {
  docCanvasEl?.querySelectorAll(".search-hit").forEach((el) => {
    el.classList.remove("search-hit", "search-hit-active");
  });
}

function renderFrPreview(preview) {
  if (!frPreviewEl) return;
  frPreviewEl.replaceChildren();
  if (!preview?.samples?.length) return;
  const list = document.createElement("div");
  list.className = "fr-preview-list";
  preview.samples.forEach((s) => {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "fr-preview-item";
    row.innerHTML = `<span class="fr-preview-meta">${s.count}×</span><span class="fr-preview-snippet">${escapeHtml(s.snippet)}</span>`;
    row.addEventListener("click", () => {
      const hit = frMatches.find((m) => m.el && (m.el.textContent || "").includes(s.snippet.slice(0, 12)));
      if (hit) focusFrMatch(frMatches.indexOf(hit));
    });
    list.appendChild(row);
  });
  frPreviewEl.appendChild(list);
}

function escapeHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function syncFrStatus() {
  if (!frStatusEl) return;
  if (!frMatches.length) {
    frStatusEl.textContent = frPreview?.hits
      ? t("frPreviewOnly", { hits: frPreview.hits, paras: frPreview.paras })
      : (searchQueryEl?.value.trim() ? t("searchNoMatches") : "");
    return;
  }
  const pos = frActiveIndex >= 0 ? frActiveIndex + 1 : 0;
  frStatusEl.textContent = t("frMatchPos", { pos, total: frMatches.length });
}

function focusFrMatch(index) {
  if (!frMatches.length) return;
  const i = ((index % frMatches.length) + frMatches.length) % frMatches.length;
  clearFrHighlights();
  frMatches.forEach((m) => m.el?.classList.add("search-hit"));
  frActiveIndex = i;
  activeSearchIndex = i;
  searchMatches = frMatches.map((m) => m.el);
  const m = frMatches[i];
  m.el?.classList.add("search-hit-active");
  m.el?.scrollIntoView({ behavior: "smooth", block: "center" });
  syncFrStatus();
  if (searchCountEl) searchCountEl.textContent = t("searchMatches", { count: frMatches.length });
}

async function refreshFrPreview(edit) {
  if (!originalFileBytes || !edit.find) {
    frPreview = null;
    renderFrPreview(null);
    return;
  }
  await ensureDocLibs(false);
  const texts = await extractParagraphTextsFromDocx(originalFileBytes);
  frPreview = countReplacePreview(texts, edit);
  renderFrPreview(frPreview);
}

async function runFindReplaceScan() {
  clearFrHighlights();
  frMatches = [];
  frActiveIndex = -1;
  searchMatches = [];
  activeSearchIndex = -1;
  frPreview = null;
  if (frPreviewEl) frPreviewEl.replaceChildren();

  const edit = buildFrEdit();
  const valid = validateFrFind(edit);
  if (!valid.ok) {
    if (searchCountEl) searchCountEl.textContent = "";
    syncFrStatus();
    if (edit.find || searchQueryEl?.value.trim()) toast(t(valid.err), "warning");
    return;
  }

  const scope = searchScopeEl?.value || "all";
  frMatches = collectFrDomMatches(edit.find, scope, edit.regex);
  await refreshFrPreview(edit);

  if (searchCountEl) {
    searchCountEl.textContent = frMatches.length
      ? t("searchMatches", { count: frMatches.length })
      : t("searchNoMatches");
  }
  if (frMatches.length) focusFrMatch(0);
  else syncFrStatus();
}

function runDocumentSearch() {
  return runFindReplaceScan();
}

function stepFrMatch(delta) {
  if (!frMatches.length) {
    runFindReplaceScan();
    return;
  }
  focusFrMatch(frActiveIndex + delta);
}

async function replaceFrOne() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  const edit = buildFrEdit();
  const valid = validateFrFind(edit);
  if (!valid.ok) { toast(t(valid.err), "warning"); return; }
  if (!frMatches.length) {
    await runFindReplaceScan();
    if (!frMatches.length) return;
  }
  const skip = frActiveIndex >= 0 ? frActiveIndex : 0;
  const count = await applyDocumentEdit(edit, { maxReplacements: 1, skipReplacements: skip });
  if (count > 0) {
    toast(t("frReplacedOne"), "success");
    await runFindReplaceScan();
    if (frMatches.length) focusFrMatch(Math.min(skip, frMatches.length - 1));
  } else toast(t("editNothing"), "info");
}

async function replaceFrAll() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  const edit = buildFrEdit();
  const valid = validateFrFind(edit);
  if (!valid.ok) { toast(t(valid.err), "warning"); return; }

  await refreshFrPreview(edit);
  const hits = frPreview?.hits || 0;
  if (!hits) {
    toast(t("searchNoMatches"), "info");
    return;
  }
  const msg = t("frReplaceAllConfirm", { hits, paras: frPreview.paras });
  if (!window.confirm(msg)) return;

  const count = await applyDocumentEdit(edit);
  if (count > 0) toast(t("editApplied", { count }), "success");
  else toast(t("editNothing"), "info");
  await runFindReplaceScan();
}

function wireFindReplaceWorkbench() {
  if (frScanBtn) frScanBtn.addEventListener("click", () => runFindReplaceScan());
  if (frPrevBtn) frPrevBtn.addEventListener("click", () => stepFrMatch(-1));
  if (frNextBtn) frNextBtn.addEventListener("click", () => stepFrMatch(1));
  if (frReplaceOneBtn) frReplaceOneBtn.addEventListener("click", () => replaceFrOne());
  if (frReplaceAllBtn) frReplaceAllBtn.addEventListener("click", () => replaceFrAll());
  if (searchQueryEl) {
    searchQueryEl.addEventListener("keydown", (e) => {
      if (e.key === "Enter") runFindReplaceScan();
      if (e.key === "F3") { e.preventDefault(); stepFrMatch(e.shiftKey ? -1 : 1); }
    });
  }
}

wireFindReplaceWorkbench();
