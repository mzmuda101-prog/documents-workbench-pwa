// Grammar panel — scan, grouped suggestions, apply one / rule / all.

let grammarScan = null;
let grammarActiveHitId = null;

const grammarLangEl = document.getElementById("grammarLang");
const grammarNbspEl = document.getElementById("grammarNbsp");
const grammarScanBtn = document.getElementById("grammarScanBtn");
const grammarApplyAllBtn = document.getElementById("grammarApplyAllBtn");
const grammarClearBtn = document.getElementById("grammarClearBtn");
const grammarStatusEl = document.getElementById("grammarStatus");
const grammarSuggestionsEl = document.getElementById("grammarSuggestions");

function grammarScanOpts() {
  return {
    lang: grammarLangEl?.value || "pl",
    nbspPl: !!grammarNbspEl?.checked,
  };
}

function escapeGrammarHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function ruleLabel(ruleId) {
  const key = `grammarRule${ruleId.split("-").map((p) => p.charAt(0).toUpperCase() + p.slice(1)).join("")}`;
  return t(key);
}

function syncGrammarStatus() {
  if (!grammarStatusEl) return;
  if (!grammarScan?.hits?.length) {
    grammarStatusEl.textContent = grammarScan ? t("grammarNoHits") : "";
    return;
  }
  const ruleCount = Object.keys(grammarScan.byRule || {}).length;
  grammarStatusEl.textContent = t("grammarFound", {
    hits: grammarScan.hits.length,
    rules: ruleCount,
  });
}

function renderGrammarSuggestions() {
  if (!grammarSuggestionsEl) return;
  grammarSuggestionsEl.replaceChildren();
  if (!grammarScan?.hits?.length) return;

  Object.entries(grammarScan.byRule).forEach(([ruleId, hits]) => {
    const details = document.createElement("details");
    details.className = "grammar-rule-group";
    details.open = hits.length <= 4;

    const summary = document.createElement("summary");
    summary.className = "grammar-rule-summary";
    summary.innerHTML = `<span class="grammar-rule-name">${escapeGrammarHtml(ruleLabel(ruleId))}</span><span class="grammar-rule-count">(${hits.length})</span>`;
    details.appendChild(summary);

    const applyRuleBtn = document.createElement("button");
    applyRuleBtn.type = "button";
    applyRuleBtn.className = "btn grammar-apply-rule";
    applyRuleBtn.textContent = t("grammarApplyRule");
    applyRuleBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      applyGrammarByRule(ruleId);
    });
    summary.appendChild(applyRuleBtn);

    const list = document.createElement("div");
    list.className = "fr-preview-list grammar-hit-list";

    hits.forEach((hit) => {
      const row = document.createElement("div");
      row.className = "grammar-hit-row";
      if (hit.id === grammarActiveHitId) row.classList.add("is-active");

      const snippet = document.createElement("button");
      snippet.type = "button";
      snippet.className = "fr-preview-item grammar-hit-item";
      snippet.innerHTML = `<span class="fr-preview-snippet">${escapeGrammarHtml(hit.snippet)}</span>`;
      snippet.addEventListener("click", () => {
        grammarActiveHitId = hit.id;
        renderGrammarSuggestions();
        scrollToGrammarParagraph(hit.paraIndex);
      });
      row.appendChild(snippet);

      const applyBtn = document.createElement("button");
      applyBtn.type = "button";
      applyBtn.className = "btn grammar-apply-one";
      applyBtn.textContent = t("grammarApplyOne");
      applyBtn.addEventListener("click", () => applyGrammarOne(hit));
      row.appendChild(applyBtn);

      list.appendChild(row);
    });

    details.appendChild(list);
    grammarSuggestionsEl.appendChild(details);
  });
}

function scrollToGrammarParagraph(paraIndex) {
  if (!docCanvasEl) return;
  const root = getDocContentRoot(docCanvasEl);
  if (!root) return;
  const paras = root.querySelectorAll("p");
  const el = paras[paraIndex];
  if (!el) return;
  docCanvasEl.querySelectorAll(".search-hit-active").forEach((n) => n.classList.remove("search-hit-active"));
  el.classList.add("search-hit", "search-hit-active");
  el.scrollIntoView({ behavior: "smooth", block: "center" });
}

function clearGrammarState() {
  grammarScan = null;
  grammarActiveHitId = null;
  if (grammarSuggestionsEl) grammarSuggestionsEl.replaceChildren();
  docCanvasEl?.querySelectorAll(".search-hit, .search-hit-active").forEach((el) => {
    el.classList.remove("search-hit", "search-hit-active");
  });
  syncGrammarStatus();
}

async function runGrammarScan() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  grammarActiveHitId = null;
  grammarScan = await scanDocument(originalFileBytes, grammarScanOpts());
  renderGrammarSuggestions();
  syncGrammarStatus();
  if (grammarScan.hits.length) {
    grammarActiveHitId = grammarScan.hits[0].id;
    renderGrammarSuggestions();
    scrollToGrammarParagraph(grammarScan.hits[0].paraIndex);
  }
}

async function applyGrammarOne(hit) {
  if (!originalFileBytes || !hit) return;
  const count = await applyDocumentEdit({
    op: "paragraphBatch",
    items: [{ index: hit.paraIndex, text: hit.fixedParagraph }],
  });
  if (count > 0) {
    toast(t("grammarAppliedOne"), "success");
    await runGrammarScan();
  } else toast(t("editNothing"), "info");
}

async function applyGrammarBatch(filterRuleId = null) {
  if (!originalFileBytes || !grammarScan?.hits?.length) {
    toast(t("grammarNoHits"), "info");
    return;
  }
  const opts = grammarScanOpts();
  await ensureDocLibs(false);
  const texts = await extractParagraphTextsFromDocx(originalFileBytes);
  const hits = filterRuleId
    ? grammarScan.hits.filter((h) => h.ruleId === filterRuleId)
    : grammarScan.hits;
  if (!hits.length) {
    toast(t("grammarNoHits"), "info");
    return;
  }
  const items = buildGrammarBatchItems(texts, hits, opts, filterRuleId);
  if (!items.length) {
    toast(t("editNothing"), "info");
    return;
  }
  if (!filterRuleId) {
    const msg = t("grammarApplyAllConfirm", { count: items.length });
    if (!window.confirm(msg)) return;
  }
  const count = await applyDocumentEdit({ op: "paragraphBatch", items });
  if (count > 0) toast(t("editApplied", { count }), "success");
  else toast(t("editNothing"), "info");
  await runGrammarScan();
}

async function applyGrammarByRule(ruleId) {
  await applyGrammarBatch(ruleId);
}

async function applyGrammarAll() {
  await applyGrammarBatch(null);
}

function wireGrammarPanel() {
  if (grammarScanBtn) grammarScanBtn.addEventListener("click", () => runGrammarScan());
  if (grammarApplyAllBtn) grammarApplyAllBtn.addEventListener("click", () => applyGrammarAll());
  if (grammarClearBtn) grammarClearBtn.addEventListener("click", () => clearGrammarState());
  if (grammarLangEl) grammarLangEl.addEventListener("change", () => { if (grammarScan) runGrammarScan(); });
  if (grammarNbspEl) grammarNbspEl.addEventListener("change", () => { if (grammarScan) runGrammarScan(); });
}

wireGrammarPanel();
