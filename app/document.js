// Document IO: lazy libs, load/save orchestration.

let _docLibsPromise = null;

function _loadScriptOnce(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.async = false;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("Failed to load: " + src));
    document.head.appendChild(s);
  });
}

async function ensureDocLibs(showFeedback = false) {
  const needPreview = !window.docx || typeof window.docx.renderAsync !== "function";
  const needZip = !window.JSZip;
  if (!needPreview && !needZip) return true;

  if (!_docLibsPromise) {
    const tasks = [];
    if (needZip) tasks.push(_loadScriptOnce("lib/jszip.min.js"));
    if (needPreview) tasks.push(_loadScriptOnce("lib/docx-preview.bundle.js"));
    _docLibsPromise = Promise.all(tasks).catch((e) => {
      _docLibsPromise = null;
      throw e;
    });
  }

  try {
    await _docLibsPromise;
  } catch (_) { /* offline without cache */ }

  const ok = !!(window.JSZip && window.docx && window.docx.renderAsync);
  if (!ok && showFeedback) {
    setStatus(t("libsMissingStatus"));
    toast(t("libsMissingToast"), "error");
    log("Brak bibliotek DOCX (JSZip / docx-preview).", "error");
  }
  return ok;
}

function setFileUi(name, size) {
  if (!fileNameEl || !fileNameTextEl) return;
  fileNameEl.classList.remove("hidden");
  fileNameTextEl.textContent = `${name} (${formatFileSize(size)})`;
  if (typeof syncActionButtons === "function") syncActionButtons();
}

async function ingestFile(file, options = {}) {
  if (!file) return false;
  const type = detectFileType(file.name, file.type);
  if (type !== "docx") {
    if (type === "pdf") toast(t("pdfSoon"), "info");
    else toast(t("unsupportedType"), "warning");
    return false;
  }

  setLoading(true, t("loadingFile"));
  try {
    const ok = await ensureDocLibs(true);
    if (!ok) return false;

    const buf = await file.arrayBuffer();
    originalFileBytes = new Uint8Array(buf);
    pendingDocEdits = [];
    currentFileName = file.name || "document.docx";
    currentFileType = "docx";
    fileHandle = options.handle || null;

    setFileUi(currentFileName, originalFileBytes.byteLength);
    setDirtyState(false);

    await renderCurrentDocument();
    setStatus(t("docLoaded"));
    if (!options.silent) toast(t("docLoaded"), "success");
    if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
    return true;
  } catch (e) {
    log(String(e.message || e), "error");
    toast(t("libsMissingToast"), "error");
    return false;
  } finally {
    setLoading(false);
  }
}

async function loadSampleDocument() {
  setLoading(true, t("loadingFile"));
  try {
    const res = await fetch("docs/samples/sample.docx");
    if (!res.ok) throw new Error("sample missing");
    const buf = await res.arrayBuffer();
    const file = new File([buf], "sample.docx", {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    await ingestFile(file, { silent: true });
    toast(t("sampleLoaded"), "success");
  } catch (e) {
    log(String(e.message || e), "error");
    toast(t("libsMissingToast"), "error");
  } finally {
    setLoading(false);
  }
}

async function reloadFromBytes(bytes) {
  originalFileBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  await renderCurrentDocument();
}

async function renderCurrentDocument() {
  if (!originalFileBytes || currentFileType !== "docx") return;
  setLoading(true, t("renderingDoc"));
  try {
    hideEmptyState();
    await renderDocxPreview(originalFileBytes, docCanvasEl);
    documentStructure = analyzeDocumentDom(docCanvasEl);
    renderStructurePanel(documentStructure);
    if (searchQueryEl?.value.trim()) runDocumentSearch();
    setupInlineEditingAfterRender();
    if (typeof syncMobileDocZoomAfterRender === "function") syncMobileDocZoomAfterRender();
  } finally {
    setLoading(false);
    syncDocViewportHeight();
  }
}

async function buildDocumentForSave() {
  if (!originalFileBytes) return null;
  const inlineEdits = collectInlineParagraphEdits();
  const edits = [...pendingDocEdits];
  if (inlineEdits.length) edits.push({ op: "paragraphBatch", items: inlineEdits });
  if (!edits.length) return originalFileBytes;
  const { bytes } = await buildPatchedDocx(originalFileBytes, edits);
  return bytes;
}

async function applyDocumentEdit(edit) {
  if (!originalFileBytes) return 0;
  await mergeInlineEditsIntoBytes();
  const normalized = edit.op ? edit : { ...edit, op: "replace" };
  recordPendingEdit(normalized);
  const { bytes, changeCount } = await buildPatchedDocx(originalFileBytes, pendingDocEdits);
  pendingDocEdits = [];
  originalFileBytes = bytes;
  await reloadFromBytes(bytes);
  if (changeCount > 0) setDirtyState(true);
  return changeCount;
}
