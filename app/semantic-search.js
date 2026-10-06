// Optional local semantic search over the current document.
// Keeps a 256d in-memory index and invalidates it whenever the preview is rendered again.

const semanticEnableBtn = document.getElementById("semanticEnableBtn");
const semanticQueryEl = document.getElementById("semanticQuery");
const semanticSearchBtn = document.getElementById("semanticSearchBtn");
const semanticReleaseBtn = document.getElementById("semanticReleaseBtn");
const semanticStatusEl = document.getElementById("semanticStatus");
const semanticResultsEl = document.getElementById("semanticResults");
const semanticAvailabilityEl = document.getElementById("semanticAvailability");

const SEMANTIC_BATCH_SIZE = 6;
const SEMANTIC_MIN_CHARS = 24;
let semanticWorker = null;
let semanticReady = false;
let semanticIndex = null;
let semanticRequest = 0;
let semanticPendingQuery = "";

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
  } else if (!semanticReady) {
    semanticEnableBtn.disabled = false;
    semanticEnableBtn.textContent = semanticText("semanticEnable");
  }
  semanticQueryEl.disabled = !semanticReady;
  semanticSearchBtn.disabled = !semanticReady || !originalFileBytes;
  semanticReleaseBtn.disabled = !semanticWorker;
}

function normalizeSemanticText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function buildSemanticPassages() {
  const structure = documentStructure || (docCanvasEl ? analyzeDocumentDom(docCanvasEl) : null);
  const source = (structure?.outline || []).filter((item) => item.type !== "table" || item.label);
  const passages = [];
  for (const item of source) {
    const text = normalizeSemanticText(item.type === "table" ? item.label : item.el?.textContent || item.label);
    if (text.length >= SEMANTIC_MIN_CHARS) passages.push({ item, text });
  }
  return passages;
}

function dot(a, b) {
  let total = 0;
  for (let i = 0; i < a.length; i++) total += a[i] * b[i];
  return total;
}

function clearSemanticHighlights() {
  docCanvasEl?.querySelectorAll(".semantic-hit, .semantic-hit-active").forEach((el) => el.classList.remove("semantic-hit", "semantic-hit-active"));
}

function jumpToSemanticResult(result) {
  clearSemanticHighlights();
  result.item.el?.classList.add("semantic-hit", "semantic-hit-active");
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
    score.textContent = `${Math.round(Math.max(0, result.score) * 100)}%`;
    const text = document.createElement("span");
    text.className = "semantic-result-text";
    text.textContent = result.text;
    button.append(score, text);
    button.addEventListener("click", () => jumpToSemanticResult(result));
    semanticResultsEl.append(button);
    if (index === 0) result.item.el?.classList.add("semantic-hit");
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
    setSemanticStatus("semanticReady");
    syncSemanticUi();
    return;
  }
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
        const requestId = ++semanticRequest;
        setSemanticStatus("semanticSearching");
        semanticWorker.postMessage({ type: "embed-query", requestId, query });
      }
    }
    return;
  }
  if (data.type === "query-vector" && data.requestId === semanticRequest && semanticIndex?.ready) {
    const query = new Float32Array(data.vector);
    const results = semanticIndex.passages.map((passage, index) => ({ ...passage, score: dot(query, semanticIndex.vectors[index]) }))
      .sort((a, b) => b.score - a.score).slice(0, 8);
    renderSemanticResults(results);
    setSemanticStatus("semanticResults", { count: results.length });
  }
}

function ensureSemanticWorker() {
  if (semanticWorker) return semanticWorker;
  semanticWorker = new Worker("app/semantic-search-worker.js?v=20261006-09", { type: "module", name: "dwb-semantic-search" });
  semanticWorker.addEventListener("message", ({ data }) => workerMessage(data));
  semanticWorker.addEventListener("error", (event) => {
    setSemanticStatus("semanticError", { message: event.message || semanticText("semanticWorkerError") });
    semanticReady = false;
    syncSemanticUi();
  });
  return semanticWorker;
}

function embedSemanticNextBatch() {
  const index = semanticIndex;
  if (!index || !semanticWorker) return;
  const batch = index.passages.slice(index.offset, index.offset + SEMANTIC_BATCH_SIZE).map((p) => p.text);
  setSemanticStatus("semanticIndexing", { done: index.offset, total: index.passages.length });
  semanticWorker.postMessage({ type: "embed-documents", requestId: index.requestId, texts: batch });
}

function ensureSemanticIndex() {
  if (semanticIndex?.ready) return true;
  const passages = buildSemanticPassages();
  if (!passages.length) { setSemanticStatus("semanticNothingToIndex"); return false; }
  semanticIndex = { requestId: ++semanticRequest, passages, vectors: [], offset: 0, ready: false };
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
  ensureSemanticWorker().postMessage({ type: "warmup" });
  syncSemanticUi();
});

semanticSearchBtn?.addEventListener("click", () => {
  const query = normalizeSemanticText(semanticQueryEl?.value);
  if (!query) { setSemanticStatus("semanticQueryMissing"); semanticQueryEl?.focus(); return; }
  if (!ensureSemanticIndex()) { semanticPendingQuery = query; return; }
  const requestId = ++semanticRequest;
  setSemanticStatus("semanticSearching");
  semanticWorker.postMessage({ type: "embed-query", requestId, query });
});
semanticQueryEl?.addEventListener("keydown", (event) => { if (event.key === "Enter") semanticSearchBtn?.click(); });
semanticReleaseBtn?.addEventListener("click", () => semanticWorker?.postMessage({ type: "release" }));

syncSemanticUi();
