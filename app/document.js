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
  if (type === "pdf") {
    // PDF → DOCX na urządzeniu (app/pdf-import.js, ładowany dopiero teraz)
    try {
      await loadLazyScript("app/pdf-import.js");
    } catch (_) {
      toast(t("libsMissingToast"), "error");
      return false;
    }
    return window.dwbPdfImport.convertFile(file, options);
  }
  if (type === "image") {
    // Zdjęcie(a) dokumentu: kartka wykryta, perspektywa wyprostowana, tło wybielone → PDF w pamięci
    // → ta sama konwersja co PDF (z pytaniem o OCR). file._photos = kilka zdjęć jako jeden dokument.
    const photos = file._photos || [file];
    try {
      await Promise.all([loadLazyScript("app/pdf-import.js"), loadLazyScript("app/photo-import.js")]);
    } catch (_) {
      toast(t("libsMissingToast"), "error");
      return false;
    }
    setLoading(true, t("photoPreparing", { n: photos.length }));
    let made;
    try {
      const base = (file.name || "zdjecie").replace(/\.[^.]+$/, "");
      made = await window.dwbPhotoImport.toPdf(photos, `${base}.pdf`);
    } catch (e) {
      log(String(e?.message || e), "error");
      toast(t(/heic|heif/i.test(photos.map((f) => f.name).join(" ")) ? "photoHeicUnsupported" : "photoDecodeFailed"), "error");
      return false;
    } finally {
      setLoading(false);
    }
    return window.dwbPdfImport.convertFile(made.file, { ...options, handle: null, photo: { pages: made.pages, cropped: made.cropped, name: file.name } });
  }
  if (type !== "docx") {
    if (/\.doc$/i.test(file.name || "")) toast(t("docOldFormat"), "warning");
    else toast(t("unsupportedType"), "warning");
    return false;
  }

  setLoading(true, t("loadingFile"));
  try {
    const ok = await ensureDocLibs(true);
    if (!ok) return false;

    const buf = await file.arrayBuffer();
    originalFileBytes = new Uint8Array(buf);
    // zepsute nazwy krojów ze starszych wersji (naprawa w pamięci; trafi do pliku przy zapisie)
    if (typeof repairDocxFontNames === "function") originalFileBytes = await repairDocxFontNames(originalFileBytes);
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

// Przerysowanie po zmianie w pliku (komentarz, tabela, kolor, Cofnij…): na ekranie zostaje to samo
// miejsce — akapit u góry obszaru dokumentu, nie surowe scrollTop. Świeży render nie ma jeszcze
// odstępów między stronami (page-breaks.js dokłada je ~200 ms później), więc ta sama liczba
// pikseli pokazywała tekst o sumę odstępów niżej, a page-breaks zakotwiczał już przesunięty widok
// (zgłoszenie Mateusza: po dodaniu komentarza strona „podjeżdża” i trzeba wrócić).
// opts.anchor — miejsce zapamiętane wcześniej (np. stan Cofnij); null = bez przywracania.
let renderScrollAnchor = null; // { anchor, until } — dla pierwszego przeliczenia odstępów stron
async function reloadFromBytes(bytes, opts = {}) {
  originalFileBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const shown = !!docCanvasEl?.querySelector(".docx-preview-host p");
  const anchor = opts.anchor !== undefined ? opts.anchor : shown ? captureDocScrollAnchor() : null;
  await renderCurrentDocument();
  if (!anchor || !restoreDocScrollAnchor(anchor)) return;
  renderScrollAnchor = { anchor, until: performance.now() + 2500 };
}

// Zmiana ramy nad dokumentem (pasek formatowania Edycji znika / wraca przy Czytanie ⇄ Edycja):
// tekst zostaje w tym samym miejscu ekranu — przesunięcie obszaru oddaje przewinięcie.
function keepDocTextInPlace(fn) {
  const vp = docViewportEl;
  const a = vp && vp.scrollTop > 0 ? captureDocScrollAnchor() : null;
  const p = a ? docBodyParagraphs(docCanvasEl)[a.index] : null;
  const before = p ? p.getBoundingClientRect().top : 0;
  fn();
  if (!p?.isConnected) return;
  const d = p.getBoundingClientRect().top - before;
  if (Math.abs(d) > 0.5) vp.scrollTop += d;
}

// page-breaks.js: odstępy stron po przerysowaniu mają zostawić na miejscu akapit sprzed
// przerysowania (nie ten, który chwilowo — bez odstępów — wypadł u góry). Raz; własne
// przewinięcie użytkownika kasuje zapamiętane miejsce.
function takeRenderScrollAnchor() {
  const r = renderScrollAnchor;
  renderScrollAnchor = null;
  return r && performance.now() < r.until ? r.anchor : null;
}
["wheel", "touchstart", "keydown", "pointerdown"].forEach((type) => {
  docViewportEl?.addEventListener(type, () => { renderScrollAnchor = null; }, { passive: true, capture: true });
});

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
// opts.anchor — miejsce zapamiętane wcześniej (zmiana widoku: zanim zmieni się układ)
function rerenderKeepingEdits(opts = {}) {
  relayoutJob = relayoutJob.then(async () => {
    if (!originalFileBytes || currentFileType !== "docx") return;
    if (typeof mergeInlineEditsIntoBytes === "function") await mergeInlineEditsIntoBytes();
    const vp = docViewportEl;
    const ratio = vp && vp.scrollHeight > vp.clientHeight ? vp.scrollTop / (vp.scrollHeight - vp.clientHeight) : 0;
    const anchor = opts.anchor !== undefined ? opts.anchor : ratio ? captureDocScrollAnchor() : null;
    await renderCurrentDocument();
    if (!vp || (!ratio && !anchor)) return;
    // to samo miejsce w dokumencie: ten sam akapit u góry (Widok mobilny ⇄ desktopowy
    // zawija tekst inaczej, więc sama proporcja przewinięcia potrafiła odjechać o strony)
    if (!restoreDocScrollAnchor(anchor)) vp.scrollTop = ratio * Math.max(0, vp.scrollHeight - vp.clientHeight);
  }).catch((e) => log(String(e?.message || e), "error"));
  return relayoutJob;
}

// Pierwszy akapit treści widoczny u góry obszaru dokumentu + jaka jego część jest już nad
// krawędzią. Akapity idą z góry na dół, więc wystarczy wyszukiwanie binarne.
function captureDocScrollAnchor() {
  const vp = docViewportEl;
  if (!vp || typeof docBodyParagraphs !== "function") return null;
  const paras = docBodyParagraphs(docCanvasEl);
  if (!paras.length) return null;
  const top = vp.getBoundingClientRect().top;
  let lo = 0;
  let hi = paras.length - 1;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (paras[mid].getBoundingClientRect().bottom <= top + 1) lo = mid + 1;
    else hi = mid;
  }
  const r = paras[lo].getBoundingClientRect();
  return { index: lo, frac: r.height ? Math.max(0, Math.min(1, (top - r.top) / r.height)) : 0, gap: Math.max(0, r.top - top) };
}

function restoreDocScrollAnchor(anchor) {
  const vp = docViewportEl;
  if (!vp || !anchor || typeof docBodyParagraphs !== "function") return false;
  const p = docBodyParagraphs(docCanvasEl)[anchor.index];
  if (!p) return false;
  const r = p.getBoundingClientRect();
  vp.scrollTop += r.top + anchor.frac * r.height - anchor.gap - vp.getBoundingClientRect().top;
  return true;
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
