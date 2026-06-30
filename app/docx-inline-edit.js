// Inline WYSIWYG — contenteditable paragraphs synced back to word/document.xml on save.

let baselineParagraphTexts = [];
let pendingInlineCursor = null;
let inlineKeyboardBound = false;

function normalizePreviewText(s) {
  return String(s || "").replace(/\r\n/g, "\n");
}

function collapseSpacesKeepNewlines(s) {
  return normalizePreviewText(s).replace(/[^\S\n]+/g, " ").trim();
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
    const next = collapseSpacesKeepNewlines(raw);
    const prev = collapseSpacesKeepNewlines(baselineParagraphTexts[i]);
    if (next !== prev) edits.push({ index: i, text: normalizePreviewText(raw) });
  }
  return edits;
}

function onInlineParagraphInput() {
  if (!readOnlyMode) setDirtyState(true);
}

function resolveParaIndex(p) {
  let idx = Number(p?.dataset?.paraIndex);
  if (Number.isFinite(idx) && idx >= 0) return idx;
  return collectPreviewParagraphElements(p?.closest(".docx-preview-host")).indexOf(p);
}

function isListParagraph(p) {
  if (!p) return false;
  return p.classList.contains("docx-bullet-fixed")
    || p.classList.contains("docx-list-numbered-fixed")
    || /docx-num-\d+-\d+/.test(p.className || "");
}

function getCaretOffset(el) {
  const sel = window.getSelection();
  if (!sel?.rangeCount) return (el.innerText || "").length;
  const range = sel.getRangeAt(0);
  if (!el.contains(range.startContainer)) return (el.innerText || "").length;
  const pre = range.cloneRange();
  pre.selectNodeContents(el);
  pre.setEnd(range.startContainer, range.startOffset);
  return pre.toString().length;
}

function insertTextAtCaret(text) {
  const sel = window.getSelection();
  if (!sel?.rangeCount) return false;
  const range = sel.getRangeAt(0);
  range.deleteContents();
  range.insertNode(document.createTextNode(text));
  range.collapse(false);
  sel.removeAllRanges();
  sel.addRange(range);
  return true;
}

function focusParagraphAtOffset(paraIndex, offset) {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  const el = collectPreviewParagraphElements(host)[paraIndex];
  if (!el) return;
  el.focus();
  const range = document.createRange();
  const sel = window.getSelection();
  let remaining = Math.max(0, offset);
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let node = walker.nextNode();
  while (node) {
    const len = node.length;
    if (remaining <= len) {
      range.setStart(node, remaining);
      range.collapse(true);
      sel?.removeAllRanges();
      sel?.addRange(range);
      return;
    }
    remaining -= len;
    node = walker.nextNode();
  }
  range.selectNodeContents(el);
  range.collapse(offset <= 0);
  sel?.removeAllRanges();
  sel?.addRange(range);
}

function focusParagraphStart(paraIndex) {
  focusParagraphAtOffset(paraIndex, 0);
}

function getHeadingParaIndices() {
  const structure = documentStructure || (docCanvasEl ? analyzeDocumentDom(docCanvasEl) : null);
  if (!structure?.headings?.length) return [];
  const indices = structure.headings
    .map((h) => (Number.isFinite(h.paraIndex) ? h.paraIndex : collectPreviewParagraphElements(docCanvasEl?.querySelector(".docx-preview-host")).indexOf(h.el)))
    .filter((i) => i >= 0);
  return [...new Set(indices)].sort((a, b) => a - b);
}

function jumpToHeading(delta) {
  const headings = getHeadingParaIndices();
  if (!headings.length) return false;
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  const paras = collectPreviewParagraphElements(host);
  const active = document.activeElement?.closest?.("p");
  let current = active ? resolveParaIndex(active) : -1;
  if (current < 0) current = 0;
  let targetIdx = 0;
  if (delta > 0) {
    targetIdx = headings.find((i) => i > current) ?? headings[0];
  } else {
    const prev = headings.filter((i) => i < current);
    targetIdx = prev.length ? prev[prev.length - 1] : headings[headings.length - 1];
  }
  const el = paras[targetIdx];
  if (!el) return false;
  focusParagraphStart(targetIdx);
  el.scrollIntoView({ behavior: "smooth", block: "center" });
  return true;
}

async function handleInlineEnter(p, paraIndex, e) {
  if (e.shiftKey) {
    e.preventDefault();
    if (!insertTextAtCaret("\n")) {
      const text = p.innerText || "";
      const at = getCaretOffset(p);
      p.innerText = text.slice(0, at) + "\n" + text.slice(at);
      focusParagraphAtOffset(paraIndex, at + 1);
    }
    onInlineParagraphInput();
    return;
  }
  e.preventDefault();
  const text = normalizePreviewText(p.innerText || p.textContent || "");
  const at = getCaretOffset(p);
  const before = text.slice(0, at);
  const after = text.slice(at).replace(/^\n/, "");
  pendingInlineCursor = { paraIndex: paraIndex + 1, offset: 0 };
  await applyInlineStructuralEdit({ op: "splitParagraph", index: paraIndex, before, after });
}

async function handleInlineBackspace(p, paraIndex, e) {
  if (getCaretOffset(p) !== 0 || paraIndex <= 0) return;
  e.preventDefault();
  const prevLen = normalizePreviewText(
    collectPreviewParagraphElements(p.closest(".docx-preview-host"))[paraIndex - 1]?.innerText || ""
  ).length;
  pendingInlineCursor = { paraIndex: paraIndex - 1, offset: prevLen };
  await applyInlineStructuralEdit({ op: "mergeParagraph", index: paraIndex });
}

async function handleInlineTab(p, paraIndex, e) {
  if (e.shiftKey) {
    if (!isListParagraph(p)) return;
    e.preventDefault();
    const at = getCaretOffset(p);
    pendingInlineCursor = { paraIndex, offset: at };
    await applyInlineStructuralEdit({ op: "listLevel", index: paraIndex, delta: -1 });
    return;
  }
  if (isListParagraph(p)) {
    e.preventDefault();
    const at = getCaretOffset(p);
    pendingInlineCursor = { paraIndex, offset: at };
    await applyInlineStructuralEdit({ op: "listLevel", index: paraIndex, delta: 1 });
    return;
  }
  e.preventDefault();
  insertTextAtCaret("\t");
  onInlineParagraphInput();
}

function getListLevelFromDom(p) {
  const m = (p.className || "").match(/docx-num-\d+-(\d+)/);
  return m ? parseInt(m[1], 10) : 0;
}

function setListLevelOnDom(p, level) {
  const cls = p.className || "";
  const idMatch = cls.match(/docx-num-(\d+)-\d+/);
  if (!idMatch) return;
  const numId = idMatch[1];
  p.className = cls.replace(/docx-num-\d+-\d+/, `docx-num-${numId}-${level}`);
  if (p.classList.contains("docx-bullet-fixed")) {
    p.classList.remove("docx-bullet-l0", "docx-bullet-l1", "docx-bullet-l2");
    p.classList.add(`docx-bullet-l${level % 3}`);
  }
}

function applyDomParagraphSplit(edit) {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  const paras = collectPreviewParagraphElements(host);
  const p = paras[edit.index];
  if (!p) return;
  p.textContent = edit.before;
  const newP = p.cloneNode(false);
  newP.className = p.className;
  newP.removeAttribute("data-inline-bound");
  newP.removeAttribute("data-para-index");
  newP.textContent = edit.after;
  p.parentNode.insertBefore(newP, p.nextSibling);
}

function applyDomParagraphMerge(edit) {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  const paras = collectPreviewParagraphElements(host);
  const curr = paras[edit.index];
  const prev = paras[edit.index - 1];
  if (!curr || !prev) return;
  prev.textContent = normalizePreviewText(prev.textContent || "") + normalizePreviewText(curr.textContent || "");
  curr.remove();
}

function applyDomListLevel(edit) {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  const p = collectPreviewParagraphElements(host)[edit.index];
  if (!p) return;
  const next = Math.max(0, Math.min(8, getListLevelFromDom(p) + edit.delta));
  setListLevelOnDom(p, next);
}

async function applyInlineStructuralEdit(edit) {
  if (!originalFileBytes) return 0;
  await mergeInlineEditsIntoBytes();
  const { bytes, changeCount } = await buildPatchedDocx(originalFileBytes, [edit]);
  if (!changeCount) return 0;
  originalFileBytes = bytes;
  await refreshInlineEditBaseline(bytes);
  if (edit.op === "splitParagraph") applyDomParagraphSplit(edit);
  else if (edit.op === "mergeParagraph") applyDomParagraphMerge(edit);
  else if (edit.op === "listLevel") applyDomListLevel(edit);
  syncInlineEditMode();
  if (docCanvasEl) {
    documentStructure = analyzeDocumentDom(docCanvasEl);
    renderStructurePanel(documentStructure);
  }
  setDirtyState(true);
  return changeCount;
}

function handleInlineHeadingNav(e) {
  if (!e.ctrlKey && !e.metaKey) return false;
  if (e.key !== "Tab") return false;
  if (!docCanvasEl?.classList.contains("edit-mode")) return false;
  e.preventDefault();
  jumpToHeading(e.shiftKey ? -1 : 1);
  return true;
}

async function onDocCanvasKeydown(e) {
  if (readOnlyMode || e.isComposing) return;
  if (handleInlineHeadingNav(e)) return;

  const p = e.target.closest?.(".docx-editable-p");
  if (!p) return;

  const paraIndex = resolveParaIndex(p);
  if (paraIndex < 0) return;

  if (e.key === "Enter") {
    await handleInlineEnter(p, paraIndex, e);
    return;
  }
  if (e.key === "Backspace") {
    await handleInlineBackspace(p, paraIndex, e);
    return;
  }
  if (e.key === "Tab") {
    await handleInlineTab(p, paraIndex, e);
  }
}

function bindInlineEditKeyboard() {
  if (!docCanvasEl || inlineKeyboardBound) return;
  inlineKeyboardBound = true;
  docCanvasEl.addEventListener("keydown", onDocCanvasKeydown);
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
    p.classList.toggle("docx-editable-list", editable && isListParagraph(p));
    p.spellcheck = editable;
    if (editable && !p.dataset.inlineBound) {
      p.dataset.inlineBound = "1";
      p.addEventListener("input", onInlineParagraphInput);
    }
  });
  if (editable) bindInlineEditKeyboard();
  docCanvasEl?.classList.toggle("edit-mode", editable);
  docCanvasEl?.classList.toggle("read-only", readOnlyMode);
  if (pendingInlineCursor) {
    focusParagraphAtOffset(pendingInlineCursor.paraIndex, pendingInlineCursor.offset);
    pendingInlineCursor = null;
  }
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

bindInlineEditKeyboard();
