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

// Akapity TREŚCI w kolejności pliku (= collectParagraphElements w XML). docx-preview rysuje
// każdą stronę/sekcję jako <section class="docx"> z <header>, <article> (treść), <footer> i
// <ol> przypisów. Dawniej brane było każde <p> z PIERWSZEJ sekcji — z nagłówkiem i
// przypisami, bez dalszych sekcji — więc numery akapitów przesuwały się względem pliku
// i zapis edycji trafiał w cudze akapity (np. nagłówek „Poufne” wpisany w treść).
function collectPreviewParagraphElements(host) {
  if (!host) return [];
  const body = host.querySelectorAll("section.docx > article p");
  if (body.length) return Array.from(body);
  const docxRoot = host.querySelector(".docx") || host;
  return Array.from(docxRoot.querySelectorAll("p"));
}

// Akapity, których edycja w podglądzie zgubiłaby coś z pliku: zapis akapitu przepisuje jego
// fragmenty tekstu od nowa (applyRunsToParagraphXml), więc przypis, obraz, pole, link czy
// śledzona zmiana w środku by przepadły. Takie akapity są tylko do odczytu, z wyjaśnieniem.
const INLINE_LOCK_TAGS = [
  ["ins", "lockTracked"], ["del", "lockTracked"], ["moveFrom", "lockTracked"], ["moveTo", "lockTracked"], ["rPrChange", "lockTracked"],
  ["footnoteReference", "lockNote"], ["endnoteReference", "lockNote"],
  ["drawing", "lockObject"], ["pict", "lockObject"], ["object", "lockObject"],
  ["fldChar", "lockField"], ["fldSimple", "lockField"], ["hyperlink", "lockLink"],
];
async function markLockedParagraphs(bytes) {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host || !bytes) return;
  const doc = await getDocumentXmlDom(bytes);
  if (!doc || bytes !== originalFileBytes) return; // w międzyczasie inny plik/wersja
  const previews = collectPreviewParagraphElements(host);
  collectParagraphElements(doc.documentElement, "all").forEach((xp, i) => {
    const el = previews[i];
    if (!el) return;
    const hit = INLINE_LOCK_TAGS.find(([tag]) => xp.getElementsByTagNameNS(W_NS, tag).length);
    if (hit) el.dataset.lock = hit[1];
    else delete el.dataset.lock;
  });
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

// ── które akapity mogły się zmienić (paczka F) ─────────────────────────────────
// Dawniej każde „co się zmieniło?” (początek pisania = migawka cofania, zapis, operacje
// z panelu) porównywało WSZYSTKIE akapity z bazą — przy ~6000 akapitach dziesiątki ms
// przy klawiszu. Teraz obserwator zmian DOM zapisuje akapity, których cokolwiek dotknęło
// (tekst, węzły, atrybut style — dokładnie to, z czego czytamy formatowanie), bez względu
// na to, która ścieżka to zrobiła: pisanie, wklejanie, B/I/U, snippety, placeholdery.
// Zbiór jest ZACHOWAWCZY: nadmiar tylko kosztuje porównanie, nigdy nie gubi zmiany.
// Zerujemy go wyłącznie wtedy, gdy wiemy, że podgląd = plik (świeży render).
const inlineDirtyParas = new Set();
let inlineDirtyValid = false; // false = nie wiemy → pełne porównanie
let pristineParas = new WeakMap(); // akapit → kopia dzieci z renderu (dokładne cofnięcie)

function inlineDirtyAdd(node) {
  const el = node?.nodeType === 1 ? node : node?.parentElement;
  const p = el?.closest?.("p");
  if (p && docCanvasEl?.contains(p)) inlineDirtyAdd._set.add(p);
}
inlineDirtyAdd._set = inlineDirtyParas;
const inlineDirtyObserver = typeof MutationObserver === "function"
  ? new MutationObserver((records) => records.forEach((r) => inlineDirtyAdd(r.target)))
  : null;
function flushInlineDirty() {
  if (inlineDirtyObserver) inlineDirtyObserver.takeRecords().forEach((r) => inlineDirtyAdd(r.target));
}
// Wołane zaraz po renderze (podgląd dokładnie = plik) — przed jakimkolwiek pisaniem.
function resetInlineDirtyAfterRender() {
  if (!inlineDirtyObserver || !docCanvasEl) { inlineDirtyValid = false; return; }
  inlineDirtyObserver.disconnect();
  inlineDirtyParas.clear();
  pristineParas = new WeakMap();
  const host = docCanvasEl.querySelector(".docx-preview-host");
  collectPreviewParagraphElements(host).forEach((p) => {
    const frag = document.createDocumentFragment();
    p.childNodes.forEach((n) => frag.appendChild(n.cloneNode(true)));
    pristineParas.set(p, frag);
  });
  inlineDirtyObserver.observe(docCanvasEl, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ["style"] });
  inlineDirtyValid = true;
}

function collectInlineParagraphEdits() {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host || !baselineParagraphRuns.length) return [];
  const previews = collectPreviewParagraphElements(host);
  const edits = [];
  if (inlineDirtyValid && previews.length === baselineParagraphRuns.length) {
    flushInlineDirty();
    const indexOf = new Map(previews.map((p, i) => [p, i]));
    inlineDirtyParas.forEach((p) => {
      const i = indexOf.get(p);
      if (i === undefined) { inlineDirtyParas.delete(p); return; } // odłączony (np. scalony)
      const domRuns = extractRunsFromPreviewParagraph(p);
      if (runsEqual(domRuns, baselineParagraphRuns[i])) inlineDirtyParas.delete(p); // wrócił do stanu z pliku
      else edits.push({ index: i, runs: domRuns });
    });
    return edits.sort((a, b) => a.index - b.index);
  }
  const len = Math.min(previews.length, baselineParagraphRuns.length);
  for (let i = 0; i < len; i++) {
    const domRuns = extractRunsFromPreviewParagraph(previews[i]);
    if (!runsEqual(domRuns, baselineParagraphRuns[i])) edits.push({ index: i, runs: domRuns });
  }
  return edits;
}

// Cofnięcie samego pisania BEZ przerysowania dokumentu (undo.js): gdy plik się nie
// zmienił, wystarczy podmienić treść akapitów. Akapit z migawki → jej formatowanie;
// akapit, który w migawce był jak w pliku → dokładna kopia z renderu (albo baza).
// Zwraca false, gdy szybka ścieżka nie ma sensu (wtedy undo rysuje od nowa).
function restoreInlineParagraphs(snapshotEdits) {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host || !inlineDirtyValid || inlineStructuralPending) return false;
  const previews = collectPreviewParagraphElements(host);
  if (previews.length !== baselineParagraphRuns.length) return false;
  const want = new Map((snapshotEdits || []).map((e) => [e.index, e.runs]));
  const touch = new Set([...collectInlineParagraphEdits().map((e) => e.index), ...want.keys()]);
  touch.forEach((i) => {
    const p = previews[i];
    if (!p) return;
    if (want.has(i)) { applyRunsToPreviewParagraph(p, want.get(i)); return; }
    const pristine = pristineParas.get(p);
    if (pristine) p.replaceChildren(...[...pristine.childNodes].map((n) => n.cloneNode(true)));
    else applyRunsToPreviewParagraph(p, baselineParagraphRuns[i]);
  });
  return true;
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
  // Kolejność w DOM, nie data-para-index: po Enterze/Backspace (kolejka niżej) indeksy
  // w atrybutach są nieaktualne aż do następnego syncInlineEditMode.
  const idx = collectPreviewParagraphElements(p?.closest(".docx-preview-host")).indexOf(p);
  if (idx >= 0) return idx;
  const stored = Number(p?.dataset?.paraIndex);
  return Number.isFinite(stored) && stored >= 0 ? stored : -1;
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

// ── pamięć kursora w dokumencie ──────────────────────────────────────────────
// Przyciski i pola w panelu („Wstaw placeholder”, „Wstaw treść” snippetu, rozmiar czcionki)
// zabierają fokus — w chwili kliknięcia activeElement to już pole/przycisk, a na iPhonie
// zaznaczenie w tekście znika. Dawniej wynik: „Ustaw kursor w akapicie”, choć kursor był.
// Pamiętamy ostatnie położenie kursora/zaznaczenia w edytowalnym akapicie i przywracamy je.
let lastDocCaret = null; // { p, range }
document.addEventListener("selectionchange", () => {
  const sel = window.getSelection?.();
  if (!sel?.rangeCount) return;
  const range = sel.getRangeAt(0);
  const node = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
  const p = node?.closest?.(".docx-editable-p");
  if (p && docCanvasEl?.contains(p)) lastDocCaret = { p, range: range.cloneRange() };
});

// Akapit z kursorem: bieżący, a gdy fokus uciekł do panelu — ostatni zapamiętany (przywrócony).
function restoreDocCaret() {
  if (readOnlyMode) return null;
  const active = document.activeElement?.closest?.(".docx-editable-p");
  if (active) return active;
  const saved = lastDocCaret;
  if (!saved || !saved.p.isConnected || !saved.p.classList.contains("docx-editable-p")) return null;
  saved.p.focus({ preventScroll: true });
  const sel = window.getSelection();
  sel.removeAllRanges();
  try { sel.addRange(saved.range); } catch (_) {
    const r = document.createRange(); r.selectNodeContents(saved.p); r.collapse(false); sel.addRange(r);
  }
  return saved.p;
}

// Wstawienie z panelu jako osobny krok cofania (undo.js ładuje się później — sprawdzamy w chwili użycia).
function asUndoStep(label, fn) {
  return typeof dwbUndo !== "undefined" && dwbUndo.record ? dwbUndo.record(label, fn) : fn();
}

// Wklejanie: czysty tekst w formacie miejsca kursora (jak „Wklej tylko tekst” w Wordzie).
// Bez tego przeglądarka wklejała surowy HTML z Worda/strony — obce style, a akapity z
// wklejanego tekstu jako <p> w środku akapitu (rozjeżdżało numerację akapitów z plikiem).
// Kolejne wiersze = łamania wiersza (jak Shift+Enter), całość = jeden krok cofania.
function onDocPaste(e) {
  const p = e.target?.closest?.(".docx-editable-p");
  if (!p || readOnlyMode) return;
  const text = e.clipboardData?.getData("text/plain");
  if (text == null) return;
  e.preventDefault();
  const lines = text.replace(/\r\n?/g, "\n").replace(/\n+$/, "").split("\n");
  asUndoStep("undoOpPaste", () => {
    lines.forEach((line, i) => {
      if (i) document.execCommand("insertLineBreak");
      if (line) document.execCommand("insertText", false, line);
    });
  });
  onInlineParagraphInput();
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

function placeCaret(el, offset) {
  if (!el) return;
  el.focus({ preventScroll: true });
  const sel = window.getSelection();
  const range = document.createRange();
  let remaining = Math.max(0, offset);
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let node = walker.nextNode();
  while (node) {
    if (remaining <= node.length) {
      range.setStart(node, remaining);
      range.collapse(true);
      sel?.removeAllRanges();
      sel?.addRange(range);
      return;
    }
    remaining -= node.length;
    node = walker.nextNode();
  }
  range.selectNodeContents(el);
  range.collapse(offset <= 0);
  sel?.removeAllRanges();
  sel?.addRange(range);
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
  const newP = splitParagraphDomAtCaret(p);
  if (!newP) return;
  prepareEditableParagraph(newP);
  // Kursor w nowym akapicie OD RAZU — dawniej przeskakiwał dopiero po przebudowie pliku,
  // a wszystko wpisane w tym czasie trafiało do starego akapitu.
  placeCaret(newP, 0);
  const beforeRuns = extractRunsFromPreviewParagraph(p);
  const afterRuns = extractRunsFromPreviewParagraph(newP);
  pristineParas.delete(p); // w pliku ten akapit będzie już inny niż przy renderze
  mirrorBaseline(paraIndex, 1, beforeRuns, afterRuns);
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
  placeCaret(prev, joinAt);
  const mergedRuns = extractRunsFromPreviewParagraph(prev);
  pristineParas.delete(prev);
  mirrorBaseline(paraIndex - 1, 2, mergedRuns);
  await applyInlineStructuralEdit({ op: "mergeParagraph", index: paraIndex, mergedRuns });
}

async function handleInlineTab(p, paraIndex, e) {
  if (e.shiftKey) {
    if (!isListParagraph(p)) return;
    e.preventDefault();
    applyDomListLevel({ index: paraIndex, delta: -1 });
    await applyInlineStructuralEdit({ op: "listLevel", index: paraIndex, delta: -1 });
    return;
  }
  if (isListParagraph(p)) {
    e.preventDefault();
    applyDomListLevel({ index: paraIndex, delta: 1 });
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

// ── Enter / Backspace / Tab: kolejka operacji na pliku ──────────────────────
// Podgląd i kursor zmieniają się OD RAZU (synchronicznie w keydown), a przebudowa
// pliku (buildPatchedDocx — rozpakowanie i spakowanie całego .docx) idzie w tle,
// JEDNA PO DRUGIEJ. Dwa warunki poprawności:
//  1) baza porównań (baselineParagraphRuns = „co będzie w pliku”) zmienia się w tej
//     samej chwili co podgląd (mirrorBaseline), więc akapity podglądu i bazy zawsze
//     stoją na tych samych indeksach — collectInlineParagraphEdits nie może wpisać
//     treści akapitu w sąsiedni;
//  2) operacje trafiają do pliku w kolejności wciśnięć, a indeks każdej liczony był
//     po wszystkich poprzednich — więc zgadza się z plikiem w chwili wykonania.
// Dawniej (2026-09-28, odtworzone): szybkie pisanie po Enterze → tekst w złym akapicie,
// a zapisany plik różnił się od podglądu (34 akapity na ekranie, 33 w pliku).
let inlineStructuralChain = Promise.resolve();
let inlineStructuralPending = 0;

function mirrorBaseline(index, removeCount, ...runsList) {
  if (!baselineParagraphRuns.length) return;
  baselineParagraphRuns.splice(index, removeCount, ...runsList);
  baselineParagraphTexts.splice(index, removeCount, ...runsList.map((r) => previewRunsToPlainText(r)));
}

function waitInlineStructuralIdle() {
  return inlineStructuralChain;
}

function applyInlineStructuralEdit(edit) {
  if (!originalFileBytes) return Promise.resolve(0);
  inlineStructuralPending++;
  setDirtyState(true);
  // (nowy akapit przygotowuje handleInlineEnter; indeksy liczymy z kolejności w DOM)
  const job = inlineStructuralChain.then(async () => {
    try {
      const { bytes, changeCount } = await buildPatchedDocx(originalFileBytes, [edit]);
      if (changeCount) originalFileBytes = bytes;
      return changeCount;
    } catch (err) {
      log(`Enter/Backspace: ${err.message || err}`, "error");
      return 0;
    } finally {
      inlineStructuralPending--;
      if (!inlineStructuralPending) {
        // kolejka pusta — baza z pliku (kontrola spójności) i odświeżona struktura.
        // Przypisujemy tylko, gdy w trakcie odczytu nie padł kolejny Enter/Backspace —
        // inaczej nadpisalibyśmy lustro bazy sprzed tej operacji.
        const snapshot = originalFileBytes;
        const [texts, runs] = await Promise.all([extractParagraphTextsFromDocx(snapshot), extractParagraphRunsFromDocx(snapshot)]);
        if (!inlineStructuralPending && originalFileBytes === snapshot) {
          baselineParagraphTexts = texts;
          baselineParagraphRuns = runs;
        }
        if (docCanvasEl && !inlineStructuralPending) {
          documentStructure = analyzeDocumentDom(docCanvasEl);
          renderStructurePanel(documentStructure);
        }
      }
    }
  });
  inlineStructuralChain = job.then(() => {}, () => {});
  return job;
}

function execInlineFormat(command) {
  if (readOnlyMode) return;
  if (!restoreDocCaret()) return; // przycisk na pasku / w panelu — wracamy do zaznaczenia w tekście
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
  const p = restoreDocCaret(); // lista rozmiarów zabrała fokus — wracamy do zaznaczenia
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
  docCanvasEl.addEventListener("paste", onDocPaste);
}

// Jeden akapit (nowy po Enterze) — zamiast syncInlineEditMode() na WSZYSTKICH akapitach
// (przy ~6000 akapitach to ~150 ms opóźnienia widocznego Entera, paczka F).
function prepareEditableParagraph(p) {
  if (!p) return;
  p.contentEditable = "true";
  p.classList.add("docx-editable-p");
  p.classList.toggle("docx-editable-list", isListParagraph(p));
  p.spellcheck = true;
  if (!p.dataset.inlineBound) {
    p.dataset.inlineBound = "1";
    p.addEventListener("input", onInlineParagraphInput);
    p.addEventListener("beforeinput", onParagraphBeforeInput);
  }
}

function syncInlineEditMode() {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host) return;
  const editable = !readOnlyMode && !!originalFileBytes;
  const paragraphs = collectPreviewParagraphElements(host);
  paragraphs.forEach((p, i) => {
    p.dataset.paraIndex = String(i);
    const locked = editable && !!p.dataset.lock;
    p.classList.toggle("docx-locked-p", locked);
    if (locked) {
      p.contentEditable = "false";
      p.classList.remove("docx-editable-p", "docx-editable-list");
      p.dataset.hint = "";
      p.dataset.hintPl = I18N.pl[p.dataset.lock];
      p.dataset.hintEn = I18N.en[p.dataset.lock];
      p.dataset.hintTouch = "on";
      p.dataset.lockHint = "1";
      return;
    }
    if (p.dataset.lockHint) { // blokada zdjęta (np. zmiany zaakceptowane) albo tryb Czytanie
      ["hint", "hintPl", "hintEn", "hintTouch", "lockHint"].forEach((k) => delete p.dataset[k]);
    }
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
  const bytes = originalFileBytes;
  refreshInlineEditBaseline(bytes).then(() => markLockedParagraphs(bytes)).then(() => syncInlineEditMode());
}

async function mergeInlineEditsIntoBytes() {
  await waitInlineStructuralIdle();
  const inlineEdits = collectInlineParagraphEdits();
  // te akapity trafią do pliku — ich kopia „z renderu” przestaje być stanem pliku
  {
    const previews = collectPreviewParagraphElements(docCanvasEl?.querySelector(".docx-preview-host"));
    inlineEdits.forEach((e) => { if (previews[e.index]) pristineParas.delete(previews[e.index]); });
  }
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
    const btn = document.getElementById(id);
    // jak w Wordzie: przycisk formatu nie zabiera fokusu (zaznaczenie w tekście zostaje,
    // na iPhonie klawiatura się nie chowa)
    btn?.addEventListener("mousedown", (e) => e.preventDefault());
    btn?.addEventListener("click", () => execInlineFormat(cmd));
  });
  document.getElementById("fmtFontSize")?.addEventListener("change", (e) => applyFontSizePt(e.target.value));
}
wireFormatToolbar();
