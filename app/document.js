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
    else if (/\.doc$/i.test(file.name || "")) toast(t("docOldFormat"), "warning");
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
    if (typeof loadMetadataFromDocument === "function") loadMetadataFromDocument().catch(() => {});
    if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
    syncDocumentShellClass();
    return true;
  } catch (e) {
    log(String(e.message || e), "error");
    toast(t("libsMissingToast"), "error");
    return false;
  } finally {
    setLoading(false);
  }
}

// name: plik z docs/samples/ (bez rozszerzenia). Tylko [a-z0-9-] — adres ?sample=… nie może
// wskazać niczego poza tym katalogiem.
async function loadSampleDocument(name = "sample") {
  if (typeof confirmDiscardChanges === "function" && !confirmDiscardChanges()) return false;
  const safe = /^[a-z0-9-]{1,40}$/.test(String(name)) ? name : "sample";
  setLoading(true, t("loadingFile"));
  try {
    const res = await fetch(`docs/samples/${safe}.docx`);
    if (!res.ok) throw new Error("sample missing");
    const buf = await res.arrayBuffer();
    const file = new File([buf], `${safe}.docx`, {
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

// Ciche przerysowanie (np. po kliknięciu pola formularza): nakładka „Renderowanie…” tylko,
// gdy trwa to dłużej — przy małym dokumencie to ~10 ms i rozmyty ekran byłby samym mignięciem.
let quietRenderOnce = false;

async function renderCurrentDocument() {
  if (!originalFileBytes || currentFileType !== "docx") return;
  const quiet = quietRenderOnce;
  quietRenderOnce = false;
  const overlayTimer = quiet ? setTimeout(() => setLoading(true, t("renderingDoc")), 400) : null;
  if (!quiet) setLoading(true, t("renderingDoc"));
  try {
    hideEmptyState();
    const [, headingStyles, composeStyles] = await Promise.all([
      renderDocxPreview(originalFileBytes, docCanvasEl),
      typeof readHeadingStyleClasses === "function" ? readHeadingStyleClasses(originalFileBytes) : new Map(),
      typeof readComposeStyleClasses === "function" ? readComposeStyleClasses(originalFileBytes) : new Map(),
    ]);
    docHeadingStyleClasses = headingStyles;
    if (typeof docComposeStyleClasses !== "undefined") docComposeStyleClasses = composeStyles;
    // podgląd = plik: od teraz śledzimy tylko akapity, które coś zmieni (paczka F)
    if (typeof resetInlineDirtyAfterRender === "function") resetInlineDirtyAfterRender();
    documentStructure = analyzeDocumentDom(docCanvasEl);
    renderStructurePanel(documentStructure);
    if (searchQueryEl?.value.trim()) runDocumentSearch();
    setupInlineEditingAfterRender();
    if (typeof syncMobileDocZoomAfterRender === "function") syncMobileDocZoomAfterRender();
  } finally {
    clearTimeout(overlayTimer);
    setLoading(false);
    syncDocViewportHeight();
  }
}

// Przerysowanie TEGO SAMEGO pliku, bo zmienił się układ (telefon ⇄ komputer przy przeciąganiu
// okna na inny monitor, obrót iPada, „Dopasuj” na telefonie). Tekst wpisany w podglądzie żyje
// tylko w podglądzie, dopóki nie trafi do pliku — zwykłe renderCurrentDocument() rysowało plik
// bez niego: zmiany znikały z ekranu, „Zapisz” dalej świeciło, a zapis byłby BEZ nich
// (zgłoszenie Mateusza, Windows, 2026-10-01). Najpierw wpisane zmiany do pliku, potem rysowanie;
// kolejne wywołania czekają na poprzednie (szybkie zmiany rozmiaru okna).
let relayoutJob = Promise.resolve();
function rerenderKeepingEdits() {
  relayoutJob = relayoutJob.then(async () => {
    if (!originalFileBytes || currentFileType !== "docx") return;
    if (typeof mergeInlineEditsIntoBytes === "function") await mergeInlineEditsIntoBytes();
    const vp = docViewportEl;
    const ratio = vp && vp.scrollHeight > vp.clientHeight ? vp.scrollTop / (vp.scrollHeight - vp.clientHeight) : 0;
    await renderCurrentDocument();
    if (vp && ratio) vp.scrollTop = ratio * Math.max(0, vp.scrollHeight - vp.clientHeight); // to samo miejsce w dokumencie
  }).catch((e) => log(String(e?.message || e), "error"));
  return relayoutJob;
}

async function buildDocumentForSave() {
  if (!originalFileBytes) return null;
  // Enter/Backspace w podglądzie mogą jeszcze przebudowywać plik w tle
  if (typeof waitInlineStructuralIdle === "function") await waitInlineStructuralIdle();
  const inlineEdits = collectInlineParagraphEdits();
  const edits = [...pendingDocEdits];
  if (inlineEdits.length) edits.push({ op: "paragraphBatch", items: inlineEdits });
  if (!edits.length) return originalFileBytes;
  const { bytes } = await buildPatchedDocx(originalFileBytes, edits);
  return bytes;
}

async function applyDocumentEdit(edit, xmlOpts = {}) {
  if (!originalFileBytes) return 0;
  await mergeInlineEditsIntoBytes();
  const normalized = edit.op ? edit : { ...edit, op: "replace" };
  recordPendingEdit(normalized);
  const { bytes, changeCount } = await buildPatchedDocx(originalFileBytes, pendingDocEdits, xmlOpts);
  pendingDocEdits = [];
  originalFileBytes = bytes;
  await reloadFromBytes(bytes);
  if (changeCount > 0) setDirtyState(true);
  return changeCount;
}
