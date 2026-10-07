// Optional local semantic search over the current document.
// Keeps a 256d in-memory index and invalidates it whenever the preview is rendered again.

const semanticEnableBtn = document.getElementById("semanticEnableBtn");
const semanticQueryEl = document.getElementById("semanticQuery");
const semanticSearchBtn = document.getElementById("semanticSearchBtn");
const semanticReleaseBtn = document.getElementById("semanticReleaseBtn");
const semanticClearSelectionBtn = document.getElementById("semanticClearSelectionBtn");
const semanticStatusEl = document.getElementById("semanticStatus");
const semanticResultsEl = document.getElementById("semanticResults");
const semanticAvailabilityEl = document.getElementById("semanticAvailability");
const semanticModelVariantEl = document.getElementById("semanticModelVariant");
const semanticModelInfoEl = document.getElementById("semanticModelInfo");
const semanticDimensionsEl = document.getElementById("semanticDimensions");
const semanticDimensionsInfoEl = document.getElementById("semanticDimensionsInfo");
const semanticPanelEl = document.getElementById("panel-semantic-search");

const SEMANTIC_BATCH_SIZE = 6;
const SEMANTIC_MIN_CHARS = 24;
const SEMANTIC_VARIANT_STORAGE_KEY = "dwbSemanticModelVariant";
const SEMANTIC_DIMENSIONS_STORAGE_KEY = "dwbSemanticDimensions";
const SEMANTIC_VARIANTS = new Set(["q4", "q8"]);
const SEMANTIC_DIMENSIONS = new Set([256, 512]);
let semanticWorker = null;
let semanticReady = false;
let semanticIndex = null;
let semanticRequest = 0;
let semanticPendingQuery = "";
let semanticVariant = SEMANTIC_VARIANTS.has(localStorage.getItem(SEMANTIC_VARIANT_STORAGE_KEY)) ? localStorage.getItem(SEMANTIC_VARIANT_STORAGE_KEY) : "q4";
let semanticDimensions = SEMANTIC_DIMENSIONS.has(Number(localStorage.getItem(SEMANTIC_DIMENSIONS_STORAGE_KEY))) ? Number(localStorage.getItem(SEMANTIC_DIMENSIONS_STORAGE_KEY)) : 256;
let semanticVariantCached = null;

function semanticText(key, vars) { return typeof t === "function" ? t(key, vars) : key; }
function setSemanticStatus(key, vars) { if (semanticStatusEl) semanticStatusEl.textContent = key ? semanticText(key, vars) : ""; }

function semanticCanRun() {
  return !!navigator.gpu;
}

function syncSemanticUi() {
  const available = semanticCanRun();
  if (semanticAvailabilityEl) {
    semanticAvailabilityEl.className = `semantic-ready ${available ? "is-ready" : "is-unavailable"}`;
    semanticAvailabilityEl.textContent = semanticText(available ? "semanticWebgpuReady" : "semanticWebgpuMissing");
  }
  if (!available) {
    semanticEnableBtn.disabled = true;
    semanticEnableBtn.textContent = semanticText("semanticUnsupported");
  } else if (semanticReady) {
    semanticEnableBtn.disabled = true;
    semanticEnableBtn.textContent = semanticText("semanticRunning");
  } else if (semanticVariantCached === null) {
    semanticEnableBtn.disabled = true;
    semanticEnableBtn.textContent = semanticText("semanticChecking");
  } else {
    semanticEnableBtn.disabled = false;
    semanticEnableBtn.textContent = semanticText(semanticVariantCached ? "semanticStart" : "semanticDownloadStart");
  }
  if (semanticModelVariantEl) {
    semanticModelVariantEl.value = semanticVariant;
    semanticModelVariantEl.disabled = !available;
  }
  if (semanticModelInfoEl) semanticModelInfoEl.textContent = semanticText(semanticVariant === "q8" ? "semanticModelQ8Info" : "semanticModelQ4Info");
  if (semanticDimensionsEl) {
    semanticDimensionsEl.value = String(semanticDimensions);
    semanticDimensionsEl.disabled = !available;
  }
  if (semanticDimensionsInfoEl) semanticDimensionsInfoEl.textContent = semanticText(semanticDimensions === 512 ? "semanticDimensions512Info" : "semanticDimensions256Info");
  semanticQueryEl.disabled = !semanticReady;
  semanticSearchBtn.disabled = !semanticReady || !originalFileBytes;
  semanticReleaseBtn.disabled = !semanticWorker;
}

function normalizeSemanticText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function semanticDocumentTitle() {
  return String(currentFileName || "document").replace(/\.docx$/i, "").replace(/\s+/g, " ").trim() || "document";
}

function buildSemanticPassages() {
  const structure = documentStructure || (docCanvasEl ? analyzeDocumentDom(docCanvasEl) : null);
  const source = (structure?.outline || []).filter((item) => item.type !== "table" || item.label);
  const passages = [];
  const headingPath = [];
  for (const item of source) {
    const text = normalizeSemanticText(item.type === "table" ? item.label : item.el?.textContent || item.label);
    if (item.type === "heading") {
      headingPath.length = Math.max(0, (item.level || 1) - 1);
      headingPath.push(text);
    }
    if (text.length >= SEMANTIC_MIN_CHARS) {
      const section = headingPath.join(" › ");
      passages.push({ item, text, title: section ? `${semanticDocumentTitle()} — ${section}` : semanticDocumentTitle() });
    }
  }
  return passages;
}

function dot(a, b) {
  let total = 0;
  for (let i = 0; i < a.length; i++) total += a[i] * b[i];
  return total;
}

function lexicalBoost(text, query) {
  const haystack = normalizeSemanticText(text).toLocaleLowerCase();
  const needle = normalizeSemanticText(query).toLocaleLowerCase();
  if (!needle) return 0;
  if (haystack.includes(needle)) return 0.12;
  const terms = needle.match(/[\p{L}\p{N}]{3,}/gu) || [];
  if (!terms.length) return 0;
  const present = terms.filter((term) => haystack.includes(term)).length;
  return (present / terms.length) * 0.08;
}

function clearSemanticHighlights() {
  const highlights = docCanvasEl?.querySelectorAll(".semantic-hit, .semantic-hit-active");
  const hadHighlights = !!highlights?.length;
  highlights?.forEach((el) => el.classList.remove("semantic-hit", "semantic-hit-active"));
  if (semanticClearSelectionBtn) semanticClearSelectionBtn.hidden = true;
  return hadHighlights;
}

function jumpToSemanticResult(result) {
  clearSemanticHighlights();
  result.item.el?.classList.add("semantic-hit", "semantic-hit-active");
  if (semanticClearSelectionBtn) semanticClearSelectionBtn.hidden = false;
  result.item.el?.scrollIntoView({ behavior: "smooth", block: "center" });
}

function renderSemanticResults(results) {
  if (!semanticResultsEl) return;
  semanticResultsEl.replaceChildren();
  clearSemanticHighlights();
  if (!results.length) {
    semanticResultsEl.textContent = semanticText("semanticNoResults");
    return;
  }
  results.forEach((result, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "semantic-result";
    const score = document.createElement("span");
    score.className = "semantic-result-score";
    score.textContent = semanticText("semanticResultScore", { score: result.score.toFixed(2) });
    const text = document.createElement("span");
    text.className = "semantic-result-text";
    text.textContent = result.text;
    // Karta pokazuje krótki fragment, ale pełna treść pozostaje pod kursorem.
    text.title = result.text;
    button.append(score, text);
    button.addEventListener("click", () => jumpToSemanticResult(result));
    semanticResultsEl.append(button);
    if (index === 0) {
      result.item.el?.classList.add("semantic-hit");
      if (semanticClearSelectionBtn) semanticClearSelectionBtn.hidden = false;
    }
  });
}

function workerMessage(data) {
  if (data.type === "progress") {
    const pct = data.total ? Math.round((data.loaded / data.total) * 100) : null;
    setSemanticStatus("semanticDownloading", { progress: pct == null ? "…" : `${pct}%` });
    return;
  }
  if (data.type === "status") { setSemanticStatus("semanticLoading"); return; }
  if (data.type === "ready") {
    semanticReady = true;
    semanticVariantCached = true;
    setSemanticStatus("semanticReady");
    syncSemanticUi();
    return;
  }
  if (data.type === "dimensions-set") return;
  if (data.type === "released") {
    semanticReady = false;
    semanticIndex = null;
    semanticWorker?.terminate();
    semanticWorker = null;
    clearSemanticHighlights();
    semanticResultsEl?.replaceChildren();
    setSemanticStatus("semanticReleased");
    syncSemanticUi();
    return;
  }
  if (data.type === "cache-status" && data.variant === semanticVariant) {
    semanticVariantCached = !!data.cached;
    syncSemanticUi();
    return;
  }
  if (data.type === "error") {
    setSemanticStatus("semanticError", { message: data.message });
    semanticEnableBtn.disabled = false;
    return;
  }
  if (data.type === "document-vectors" && semanticIndex?.requestId === data.requestId) {
    const vectors = data.vectors.map((buffer) => new Float32Array(buffer));
    semanticIndex.vectors.push(...vectors);
    semanticIndex.offset += vectors.length;
    if (semanticIndex.offset < semanticIndex.passages.length) {
      embedSemanticNextBatch();
    } else {
      semanticIndex.ready = true;
      setSemanticStatus("semanticIndexed", { count: semanticIndex.passages.length });
      if (semanticPendingQuery) {
        const query = semanticPendingQuery;
        semanticPendingQuery = "";
        semanticIndex.query = query;
        const requestId = ++semanticRequest;
        setSemanticStatus("semanticSearching");
        semanticWorker.postMessage({ type: "embed-query", requestId, query });
      }
    }
    return;
  }
  if (data.type === "query-vector" && data.requestId === semanticRequest && semanticIndex?.ready) {
    const query = new Float32Array(data.vector);
    const results = semanticIndex.passages.map((passage, index) => {
      const semanticScore = dot(query, semanticIndex.vectors[index]);
      return { ...passage, score: semanticScore + lexicalBoost(passage.text, semanticIndex.query), semanticScore };
    })
      .sort((a, b) => b.score - a.score).slice(0, 8);
    renderSemanticResults(results);
    setSemanticStatus("semanticResults", { count: results.length });
  }
}

function ensureSemanticWorker() {
  if (semanticWorker) return semanticWorker;
  semanticWorker = new Worker("app/semantic-search-worker.js?v=20261007-06", { type: "module", name: "dwb-semantic-search" });
  semanticWorker.addEventListener("message", ({ data }) => workerMessage(data));
  semanticWorker.addEventListener("error", (event) => {
    setSemanticStatus("semanticError", { message: event.message || semanticText("semanticWorkerError") });
    semanticReady = false;
    syncSemanticUi();
  });
  return semanticWorker;
}

function refreshSemanticCacheStatus() {
  if (!semanticCanRun() || semanticReady) return;
  ensureSemanticWorker().postMessage({ type: "cache-status", variant: semanticVariant });
}

function switchSemanticVariant(nextVariant) {
  if (!SEMANTIC_VARIANTS.has(nextVariant) || nextVariant === semanticVariant) return;
  semanticVariant = nextVariant;
  localStorage.setItem(SEMANTIC_VARIANT_STORAGE_KEY, semanticVariant);
  semanticRequest += 1;
  semanticWorker?.terminate();
  semanticWorker = null;
  semanticReady = false;
  semanticVariantCached = null;
  semanticIndex = null;
  semanticPendingQuery = "";
  clearSemanticHighlights();
  semanticResultsEl?.replaceChildren();
  setSemanticStatus("semanticVariantChanged");
  syncSemanticUi();
  refreshSemanticCacheStatus();
}

function embedSemanticNextBatch() {
  const index = semanticIndex;
  if (!index || !semanticWorker) return;
  const batch = index.passages.slice(index.offset, index.offset + SEMANTIC_BATCH_SIZE)
    .map(({ text, title }) => ({ text, title }));
  setSemanticStatus("semanticIndexing", { done: index.offset, total: index.passages.length });
  semanticWorker.postMessage({ type: "embed-documents", requestId: index.requestId, passages: batch });
}

function ensureSemanticIndex() {
  if (semanticIndex?.ready) return true;
  const passages = buildSemanticPassages();
  if (!passages.length) { setSemanticStatus("semanticNothingToIndex"); return false; }
  semanticIndex = { requestId: ++semanticRequest, passages, vectors: [], offset: 0, ready: false, query: semanticPendingQuery };
  embedSemanticNextBatch();
  return false;
}

function semanticDocumentRendered() {
  // A DOCX re-render may replace paragraph nodes or content; old vectors are never reused.
  semanticIndex = null;
  semanticPendingQuery = "";
  clearSemanticHighlights();
  semanticResultsEl?.replaceChildren();
  if (semanticReady) setSemanticStatus("semanticNeedsIndex");
  syncSemanticUi();
}

semanticEnableBtn?.addEventListener("click", () => {
  if (!semanticCanRun()) return;
  semanticEnableBtn.disabled = true;
  setSemanticStatus("semanticLoading");
  ensureSemanticWorker().postMessage({ type: "warmup", variant: semanticVariant, dimensions: semanticDimensions });
  syncSemanticUi();
});

semanticSearchBtn?.addEventListener("click", () => {
  const query = normalizeSemanticText(semanticQueryEl?.value);
  if (!query) { setSemanticStatus("semanticQueryMissing"); semanticQueryEl?.focus(); return; }
  if (!ensureSemanticIndex()) { semanticPendingQuery = query; return; }
  semanticIndex.query = query;
  const requestId = ++semanticRequest;
  setSemanticStatus("semanticSearching");
  semanticWorker.postMessage({ type: "embed-query", requestId, query });
});
semanticQueryEl?.addEventListener("keydown", (event) => { if (event.key === "Enter") semanticSearchBtn?.click(); });
semanticReleaseBtn?.addEventListener("click", () => semanticWorker?.postMessage({ type: "release" }));
semanticClearSelectionBtn?.addEventListener("click", () => clearSemanticHighlights());
semanticModelVariantEl?.addEventListener("change", () => switchSemanticVariant(semanticModelVariantEl.value));
semanticDimensionsEl?.addEventListener("change", () => {
  const next = Number(semanticDimensionsEl.value);
  if (!SEMANTIC_DIMENSIONS.has(next) || next === semanticDimensions) return;
  semanticDimensions = next;
  localStorage.setItem(SEMANTIC_DIMENSIONS_STORAGE_KEY, String(next));
  semanticRequest += 1;
  semanticIndex = null;
  semanticPendingQuery = "";
  clearSemanticHighlights();
  semanticResultsEl?.replaceChildren();
  if (semanticWorker) semanticWorker.postMessage({ type: "set-dimensions", dimensions: next });
  if (semanticReady) setSemanticStatus("semanticDimensionsChanged");
  syncSemanticUi();
});
semanticPanelEl?.addEventListener("toggle", () => { if (semanticPanelEl.open) refreshSemanticCacheStatus(); });

syncSemanticUi();
