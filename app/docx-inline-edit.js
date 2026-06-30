// Inline WYSIWYG — contenteditable paragraphs synced back to word/document.xml on save.

let baselineParagraphTexts = [];
let baselineParagraphRuns = [];
let pendingInlineCursor = null;
let inlineKeyboardBound = false;
let activeTypingStyle = null;

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
    baselineParagraphRuns = [];
    return;
  }
  baselineParagraphTexts = await extractParagraphTextsFromDocx(bytes);
  baselineParagraphRuns = await extractParagraphRunsFromDocx(bytes);
}

function collectInlineParagraphEdits() {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host || !baselineParagraphRuns.length) return [];
  const previews = collectPreviewParagraphElements(host);
  const edits = [];
  const len = Math.min(previews.length, baselineParagraphRuns.length);
  for (let i = 0; i < len; i++) {
    const domRuns = extractRunsFromPreviewParagraph(previews[i]);
    if (!runsEqual(domRuns, baselineParagraphRuns[i])) edits.push({ index: i, runs: domRuns });
  }
  return edits;
}

function onInlineParagraphInput() {
  if (!readOnlyMode) setDirtyState(true);
}

function onParagraphBeforeInput(e) {
  if (readOnlyMode || e.isComposing) return;
  if (e.inputType !== "insertText" && e.inputType !== "insertReplacementText") return;
  const p = e.target.closest?.(".docx-editable-p");
  if (!p) return;
  const ch = e.data || "";

  if (ch && getSnippetExpandMode() === "auto" && SNIPPET_EXPAND_DELIMITER_RE.test(ch)) {
    const before = getTextBeforeCaret(p);
    const m = before.match(SNIPPET_TRIGGER_AT_END_RE);
    const sn = m ? getSnippetByName(m[1]) : null;
    if (sn) {
      e.preventDefault();
      const body = resolveSnippetBody(sn.body);
      const style = mergeRunStyles(getInheritedRunStyleAtCaret(p), activeTypingStyle);
      replaceTextEndingBeforeCaret(p, m[0].length, body + ch, style);
      onInlineParagraphInput();
      toast(t("snippetsAutoExpanded", { name: formatSnippetTrigger(sn.name) }), "success");
      return;
    }
  }

  const inherited = getInheritedRunStyleAtCaret(p);
  const style = mergeRunStyles(inherited, activeTypingStyle);
  if (!runStyleHasProps(style)) return;
  e.preventDefault();
  insertStyledTextAtCaret(ch, style, p);
  onInlineParagraphInput();
}

function onFormatSelectionChange() {
  if (readOnlyMode) return;
  const p = document.activeElement?.closest?.(".docx-editable-p");
  const fmtFontSize = document.getElementById("fmtFontSize");
  if (!fmtFontSize || !p) return;
  const pt = fontSizePtFromStyle(getInheritedRunStyleAtCaret(p));
  if (pt) fmtFontSize.value = pt;
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

function splitParagraphDomAtCaret(p) {
  const sel = window.getSelection();
  if (!sel?.rangeCount) return null;
  const range = sel.getRangeAt(0);
  if (!range.collapsed) range.deleteContents();
  const tailRange = document.createRange();
  tailRange.setStart(range.startContainer, range.startOffset);
  tailRange.setEndAfter(p.lastChild || p);
  const tail = tailRange.extractContents();
  const newP = p.cloneNode(false);
  newP.className = p.className;
  newP.removeAttribute("data-inline-bound");
  newP.removeAttribute("data-para-index");
  newP.appendChild(tail);
  p.parentNode.insertBefore(newP, p.nextSibling);
  return newP;
}

function mergeParagraphDom(prev, curr) {
  while (curr.firstChild) prev.appendChild(curr.firstChild);
  curr.remove();
}

async function handleInlineEnter(p, paraIndex, e) {
  if (e.shiftKey) {
    e.preventDefault();
    if (document.queryCommandSupported?.("insertLineBreak")) {
      document.execCommand("insertLineBreak");
    } else {
      insertTextAtCaret("\n");
    }
    onInlineParagraphInput();
    return;
  }
  e.preventDefault();
  splitParagraphDomAtCaret(p);
  pendingInlineCursor = { paraIndex: paraIndex + 1, offset: 0 };
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  const paras = collectPreviewParagraphElements(host);
  const beforeRuns = extractRunsFromPreviewParagraph(paras[paraIndex]);
  const afterRuns = extractRunsFromPreviewParagraph(paras[paraIndex + 1]);
  await applyInlineStructuralEdit({
    op: "splitParagraph",
    index: paraIndex,
    before: previewRunsToPlainText(beforeRuns),
    after: previewRunsToPlainText(afterRuns),
    beforeRuns,
    afterRuns,
  });
}

async function handleInlineBackspace(p, paraIndex, e) {
  if (getCaretOffset(p) !== 0 || paraIndex <= 0) return;
  e.preventDefault();
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  const paras = collectPreviewParagraphElements(host);
  const prev = paras[paraIndex - 1];
  const joinAt = previewRunsToPlainText(extractRunsFromPreviewParagraph(prev)).length;
  mergeParagraphDom(prev, p);
  const mergedRuns = extractRunsFromPreviewParagraph(prev);
  pendingInlineCursor = { paraIndex: paraIndex - 1, offset: joinAt };
  await applyInlineStructuralEdit({ op: "mergeParagraph", index: paraIndex, mergedRuns });
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
  if (edit.op === "listLevel") applyDomListLevel(edit);
  syncInlineEditMode();
  if (docCanvasEl) {
    documentStructure = analyzeDocumentDom(docCanvasEl);
    renderStructurePanel(documentStructure);
  }
  setDirtyState(true);
  return changeCount;
}

function execInlineFormat(command) {
  if (readOnlyMode) return;
  document.execCommand(command, false, null);
  onInlineParagraphInput();
}

function applyFontSizePt(pt) {
  if (readOnlyMode) return;
  const raw = String(pt || "").trim();
  if (!raw) {
    if (activeTypingStyle) delete activeTypingStyle.fontSize;
    return;
  }
  const sizeStyle = { fontSize: `${raw}pt` };
  activeTypingStyle = mergeRunStyles(activeTypingStyle || {}, sizeStyle);
  const p = document.activeElement?.closest?.(".docx-editable-p");
  const sel = window.getSelection();
  if (p && sel?.rangeCount && !sel.getRangeAt(0).collapsed) {
    const merged = mergeRunStyles(getInheritedRunStyleAtCaret(p), sizeStyle);
    if (applyRunStyleToSelection(merged, p)) onInlineParagraphInput();
  }
}

function handleFormatShortcut(e) {
  if (!(e.ctrlKey || e.metaKey) || e.altKey) return false;
  const cmd = { b: "bold", i: "italic", u: "underline" }[e.key.toLowerCase()];
  if (!cmd || !e.target.closest?.(".docx-editable-p")) return false;
  e.preventDefault();
  execInlineFormat(cmd);
  return true;
}

function syncFormatToolbar() {
  const bar = document.getElementById("formatToolbar");
  if (!bar) return;
  const editable = !readOnlyMode && !!originalFileBytes;
  bar.classList.toggle("hidden", !editable);
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
  if (handleFormatShortcut(e)) return;
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
      p.addEventListener("beforeinput", onParagraphBeforeInput);
    }
  });
  if (editable) {
    bindInlineEditKeyboard();
    if (!docCanvasEl.dataset.formatSelBound) {
      docCanvasEl.dataset.formatSelBound = "1";
      document.addEventListener("selectionchange", onFormatSelectionChange);
    }
  }
  docCanvasEl?.classList.toggle("edit-mode", editable);
  docCanvasEl?.classList.toggle("read-only", readOnlyMode);
  syncFormatToolbar();
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

function wireFormatToolbar() {
  [
    ["fmtBold", "bold"],
    ["fmtItalic", "italic"],
    ["fmtUnderline", "underline"],
  ].forEach(([id, cmd]) => {
    document.getElementById(id)?.addEventListener("click", () => execInlineFormat(cmd));
  });
  document.getElementById("fmtFontSize")?.addEventListener("change", (e) => applyFontSizePt(e.target.value));
}
wireFormatToolbar();
