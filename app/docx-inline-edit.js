// Inline WYSIWYG — contenteditable paragraphs synced back to word/document.xml on save.

let baselineParagraphTexts = [];

function normalizePreviewText(s) {
  return String(s || "").replace(/\s+/g, " ").trim();
}

function collectPreviewParagraphElements(host) {
  if (!host) return [];
  const docxRoot = host.querySelector(".docx") || host;
  return Array.from(docxRoot.querySelectorAll("p"));
}

async function refreshInlineEditBaseline(bytes) {
  if (!bytes) {
    baselineParagraphTexts = [];
    return;
  }
  baselineParagraphTexts = await extractParagraphTextsFromDocx(bytes);
}

function collectInlineParagraphEdits() {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host || !baselineParagraphTexts.length) return [];
  const previews = collectPreviewParagraphElements(host);
  const edits = [];
  const len = Math.min(previews.length, baselineParagraphTexts.length);
  for (let i = 0; i < len; i++) {
    const raw = previews[i].innerText ?? previews[i].textContent ?? "";
    const next = normalizePreviewText(raw);
    const prev = normalizePreviewText(baselineParagraphTexts[i]);
    if (next !== prev) edits.push({ index: i, text: raw.replace(/\r\n/g, "\n") });
  }
  return edits;
}

function onInlineParagraphInput() {
  if (!readOnlyMode) setDirtyState(true);
}

function syncInlineEditMode() {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host) return;
  const editable = !readOnlyMode && !!originalFileBytes;
  const paragraphs = collectPreviewParagraphElements(host);
  paragraphs.forEach((p, i) => {
    p.dataset.paraIndex = String(i);
    p.contentEditable = editable ? "true" : "false";
    p.classList.toggle("docx-editable-p", editable);
    p.classList.toggle("docx-editable-list", editable && (p.classList.contains("docx-list-numbered-fixed") || p.classList.contains("docx-bullet-fixed")));
    p.spellcheck = editable;
    if (editable && !p.dataset.inlineBound) {
      p.dataset.inlineBound = "1";
      p.addEventListener("input", onInlineParagraphInput);
    }
  });
  docCanvasEl?.classList.toggle("edit-mode", editable);
  docCanvasEl?.classList.toggle("read-only", readOnlyMode);
}

function setupInlineEditingAfterRender() {
  if (!originalFileBytes) return;
  refreshInlineEditBaseline(originalFileBytes).then(() => syncInlineEditMode());
}

async function mergeInlineEditsIntoBytes() {
  const inlineEdits = collectInlineParagraphEdits();
  if (!inlineEdits.length) return 0;
  const { bytes, changeCount } = await buildPatchedDocx(originalFileBytes, [{ op: "paragraphBatch", items: inlineEdits }]);
  originalFileBytes = bytes;
  await refreshInlineEditBaseline(bytes);
  if (changeCount > 0) setDirtyState(true);
  return changeCount;
}
