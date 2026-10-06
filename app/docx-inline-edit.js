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
// Druga połówka akapitu podzielonego podziałem strony (data-dwb-cont, łatka nr 6 docx-preview)
// to dalej TEN SAM akapit pliku — nie liczy się jako osobny.
function collectPreviewParagraphElements(host) {
  if (!host) return [];
  const body = host.querySelectorAll("section.docx > article p:not([data-dwb-cont])");
  if (body.length) return Array.from(body);
  const docxRoot = host.querySelector(".docx") || host;
  return Array.from(docxRoot.querySelectorAll("p"));
}

// ── jedno pole edycji na cały dokument ──────────────────────────────────────────
// W Edycji edytowalny jest pojemnik CAŁEGO dokumentu (.docx-wrapper), a nie każdy akapit osobno —
// zaznaczanie przez wiele akapitów (przeciąganie, Shift+klik, Shift+strzałki, Ctrl/⌘+A, uchwyty
// zaznaczenia na dotyku) działa jak w Wordzie. Dawniej każdy akapit był osobnym polem edycji:
// zaznaczenie zatrzymywało się na granicy akapitu, a Ctrl+A nie zaznaczał nic.
// Fokus (document.activeElement) i cel zdarzeń klawiatury to teraz pojemnik — akapit z kursorem
// wyznacza zaznaczenie (docCaretParagraph). Nieedytowalne w środku: akapity tylko do odczytu,
// nagłówki i stopki stron, znaczniki granic stron. Tekst przypisów to osobne pola (w Wordzie też
// osobna część — zaznaczenie nie przechodzi z treści do przypisów). Operacje na zaznaczeniu
// kilku akapitów (usuwanie, pisanie w miejsce, wycinanie, wklejanie) — doc-selection.js.
function docEditRoot(host = docCanvasEl?.querySelector(".docx-preview-host")) {
  return host?.querySelector(":scope > .docx-wrapper") || host || null;
}

// Akapit z kursorem dla elementu z fokusem albo celu zdarzenia: pole przypisu to dalej akapit,
// a w pojemniku dokumentu — akapit, w którym stoi ruchomy koniec zaznaczenia (kursor).
function docCaretParagraph(el) {
  if (!el?.closest) return null;
  const direct = el.closest(".docx-editable-p");
  if (direct) return direct;
  if (!el.classList?.contains("docx-edit-root")) return null;
  const sel = window.getSelection?.();
  if (!sel?.rangeCount) return null;
  const n = sel.focusNode;
  const node = n?.nodeType === 1 ? n : n?.parentElement;
  const p = node?.closest?.(".docx-editable-p");
  return p && el.contains(p) ? p : null;
}

// Fokus do pola, w którym leży akapit (pojemnik dokumentu albo akapit przypisu).
function focusDocParagraph(p) {
  if (!p) return;
  const host = p.getAttribute("contenteditable") === "true" ? p : p.closest(".docx-edit-root");
  if (host && document.activeElement !== host) host.focus({ preventScroll: true });
}

// Akapity, których edycja w podglądzie zgubiłaby coś z pliku: zapis akapitu przepisuje jego
// fragmenty tekstu od nowa (applyRunsToParagraphXml), więc przypis, obraz, pole, link czy
// śledzona zmiana w środku by przepadły. Takie akapity są tylko do odczytu, z wyjaśnieniem.
// Odnośnik do przypisu NIE blokuje — jest „wyspą” (doc-notes.js), jak znacznik komentarza.
const INLINE_LOCK_TAGS = [
  ["ins", "lockTracked"], ["del", "lockTracked"], ["moveFrom", "lockTracked"], ["moveTo", "lockTracked"], ["rPrChange", "lockTracked"],
  ["drawing", "lockObject"], ["pict", "lockObject"], ["object", "lockObject"],
  ["fldChar", "lockField"], ["fldSimple", "lockField"],
];
// Linki (w:hyperlink) z samym tekstem są edytowalne — model akapitu je zachowuje
// (docx-run-styles.js). Fragment tekstu GŁĘBIEJ niż akapit/link (np. w smartTag, customXml,
// kontrolce) zapis akapitu by zgubił — taki akapit zostaje tylko do odczytu.
// Znaczniki komentarza głębiej niż w akapicie (np. w linku) — zapis akapitu by je zgubił.
function paragraphCommentLock(xp) {
  const deep = ["commentRangeStart", "commentRangeEnd"].some((tag) => Array.from(xp.getElementsByTagNameNS(W_NS, tag)).some((m) => m.parentNode !== xp))
    // zakładka w linku — zapis akapitu buduje linki od nowa (zgubiłby ją)
    || ["bookmarkStart", "bookmarkEnd"].some((tag) => Array.from(xp.getElementsByTagNameNS(W_NS, tag)).some((m) => m.parentNode?.localName === "hyperlink"))
    || Array.from(xp.getElementsByTagNameNS(W_NS, "commentReference")).some((m) => m.parentNode?.parentNode !== xp);
  return deep ? "lockField" : null;
}

// Rysunki w akapicie jako „wyspy” (jak przypisy i komentarze): fragment z samym rysunkiem wraca
// przy zapisie akapitu dokładnie taki, jak w pliku — pisanie obok go nie gubi. Dawniej każdy akapit
// z rysunkiem był tylko do odczytu (np. tytuł strony z PDF → DOCX, bo warstwa grafiki strony jest
// przypięta do pierwszego akapitu; logo przypięte do akapitu w piśmie z Worda).
// Blokada zostaje, gdy wyspa mogłaby coś zgubić: rysunek we wspólnym fragmencie z tekstem, w linku
// albo polu (fragment nie jest bezpośrednio w akapicie), z polem tekstowym (jego akapity mają
// własną edycję — stara kopia z wyspy nadpisałaby zmiany).
const OBJECT_TAGS = ["drawing", "pict", "object"];
const OBJECT_RUN_KIDS = new Set(["rPr", "drawing", "pict", "object", "AlternateContent", "lastRenderedPageBreak"]);
function objectRunOf(o) {
  let n = o.parentNode;
  while (n && !(n.localName === "r" && n.namespaceURI === W_NS)) n = n.parentNode;
  return n;
}
function isObjectRun(r) {
  return r?.localName === "r" && OBJECT_TAGS.some((tag) => r.getElementsByTagNameNS(W_NS, tag).length);
}
function paragraphObjectLock(xp) {
  const objs = OBJECT_TAGS.flatMap((tag) => Array.from(xp.getElementsByTagNameNS(W_NS, tag)));
  if (!objs.length) return null;
  for (const o of objs) {
    const r = objectRunOf(o);
    if (!r || r.parentNode !== xp) return "lockObject";
    if (Array.from(r.childNodes).some((n) => n.nodeType === 1 && !OBJECT_RUN_KIDS.has(n.localName))) return "lockObject";
    if (r.getElementsByTagNameNS(W_NS, "txbxContent").length) return "lockObject";
  }
  return null;
}
// Rysunek „w linii tekstu” (wp:inline) — w podglądzie leży w akapicie i płynie z tekstem.
function isInlineObjectRun(r) {
  return Array.from(r.getElementsByTagNameNS("*", "inline")).some((n) => /wordprocessingDrawing/.test(n.namespaceURI || ""))
    && !Array.from(r.getElementsByTagNameNS("*", "anchor")).some((n) => /wordprocessingDrawing/.test(n.namespaceURI || ""));
}
// Rysunki w linii w podglądzie (po kolei): najwyższy element akapitu z obrazem, bez pływających.
function previewInlineObjects(el) {
  const out = [];
  el.querySelectorAll("img, svg, canvas").forEach((m) => {
    if (m.closest("svg") && m.tagName.toLowerCase() !== "svg") return; // wnętrze rysunku SVG
    let w = m;
    while (w.parentElement && w.parentElement !== el) w = w.parentElement;
    if (w.parentElement !== el || out.includes(w)) return;
    for (let a = m; a && a !== el; a = a.parentElement) if (/absolute|fixed/.test(getComputedStyle(a).position)) return;
    out.push(w);
  });
  return out;
}

// Symbole (w:sym) jako wyspy: <span> ze znakiem w podglądzie dostaje XML fragmentu z pliku.
// Kolejność i znaki muszą się zgadzać — inaczej akapit tylko do odczytu (bezpiecznie).
function stampSymbolIslands(xp, el, nextKey) {
  const runs = paragraphXmlParts(xp).filter(isSymRun);
  if (!runs.length) return null;
  const spans = [];
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const sp = n.parentElement;
    if (sp && sp !== el && sp.childNodes.length === 1 && sp.style?.fontFamily && !sp.dataset.cm) spans.push(sp);
  }
  let at = 0;
  for (const r of runs) {
    const ch = symRunText(r);
    while (at < spans.length && spans[at].textContent !== ch) at++;
    if (at >= spans.length) return "lockSymbol";
    const sp = spans[at++];
    const key = nextKey();
    docIslandXml.set(key, new XMLSerializer().serializeToString(r));
    sp.dataset.cm = key;
    sp.dataset.cmKind = "sym";
    sp.contentEditable = "false";
  }
  return null;
}

// Miejsce w podglądzie dla położenia w tekście akapitu LICZONEGO JAK ZAPIS (model akapitu:
// łamanie wiersza = 1 znak). formDomRange liczy tylko węzły tekstu, a docx-preview rysuje
// <w:br/> z pliku jako <br> (bez tekstu) — znacznik zakładki/komentarza za łamaniem wiersza lądował
// o tyle znaków dalej, ile łamań było przed nim, a następny zapis przenosił go w pliku
// (znalezione testem chaos: klik w słowo → zakładka → Shift+Enter → Cofnij/Ponów).
function modelDomRange(el, at) {
  let pos = 0;
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.nodeType === 1) {
      if (n.tagName !== "BR" || n.dataset.dwbPh) continue; // znacznik Shift+Enter tylko dla oka
      if (pos === at) { const r = document.createRange(); r.setStartBefore(n); r.collapse(true); return r; }
      pos += 1;
      continue;
    }
    const len = n.data.length;
    if (at <= pos + len) { const r = document.createRange(); r.setStart(n, at - pos); r.collapse(true); return r; }
    pos += len;
  }
  return null;
}

// Znaczniki komentarza w podglądzie: puste <span data-cm> w tych samych miejscach tekstu co w pliku
// (podgląd ich nie rysuje). Zapis akapitu oddaje je jako „wyspy” — komentarz nie ginie przy pisaniu.
// noteLabels — długości numerów przypisów (odnośnik / numer na początku przypisu) w kolejności:
// w pliku mają zerową długość, a w podglądzie numer to tekst — przesuwa położenia za nim.
function stampCommentMarks(xp, el, nextKey, noteLabels = []) {
  const parts = paragraphXmlParts(xp);
  if (!parts.some((n) => n.localName !== "r" || isCommentReferenceRun(n) || isObjectRun(n))) return;
  el.querySelectorAll('span[data-cm]:not([data-cm-kind="note"]):not([data-cm-kind="obj"]):not([data-cm-kind="sym"])').forEach((x) => x.remove());
  // rysunki w linii: wyspą jest sam obraz w podglądzie (usunięcie go usuwa go z pliku, jak w Wordzie),
  // o ile liczba się zgadza; inaczej pusty znacznik w miejscu fragmentu (rysunek zostaje w pliku)
  const inlineRuns = parts.filter((n) => isObjectRun(n) && isInlineObjectRun(n));
  const inlineEls = inlineRuns.length ? previewInlineObjects(el) : [];
  const bindInline = inlineRuns.length && inlineEls.length === inlineRuns.length;
  let inlineAt = 0;
  let offset = 0;
  let noteAt = 0;
  const groups = new Map(); // przesunięcie → [znaczniki po kolei]
  // Znaczniki o zerowej długości tuż przy obrazie w linii (zakładka-cel linku „do obrazu”, zakres
  // komentarza na obrazie) idą obok TEGO obrazu, nie na położenie w tekście — przy obrazach po
  // łamaniu wiersza położenie wypadało po drugiej stronie obrazu, akapit wychodził „zmieniony”
  // i zapis przenosił zakładkę za obraz (link do obrazu prowadził w złe miejsce).
  let afterObj = null; // { el: ostatni węzeł za obrazem, offset }
  parts.forEach((n) => {
    if (n.localName === "sdt") { offset += ffText(ffKid(n, "sdtContent")).length; return; }
    if (noteRunKind(n)) { offset += noteLabels[noteAt++] || 0; return; }
    if (isSymRun(n)) { offset += symRunText(n).length; return; } // wyspa-symbol (stampSymbolIslands)
    if (isObjectRun(n)) {
      const key = nextKey();
      docIslandXml.set(key, new XMLSerializer().serializeToString(n));
      if (bindInline && isInlineObjectRun(n)) {
        const w = inlineEls[inlineAt++];
        w.dataset.cm = key;
        w.dataset.cmKind = "obj";
        w.contentEditable = "false";
        (groups.get(offset) || []).forEach((sp) => w.before(sp)); // znaczniki przed obrazem w pliku
        groups.delete(offset);
        afterObj = { el: w, offset };
        return;
      }
      const span = document.createElement("span");
      span.dataset.cm = key;
      span.dataset.cmKind = "obj-mark";
      span.contentEditable = "false";
      span.className = "cm-mark";
      if (!groups.has(offset)) groups.set(offset, []);
      groups.get(offset).push(span);
      return;
    }
    if (n.localName !== "r" || isCommentReferenceRun(n)) {
      const key = nextKey();
      docIslandXml.set(key, new XMLSerializer().serializeToString(n));
      const span = document.createElement("span");
      span.dataset.cm = key;
      span.dataset.cmId = n.getAttributeNS(W_NS, "id") || n.getElementsByTagNameNS(W_NS, "commentReference")[0]?.getAttributeNS(W_NS, "id") || "";
      span.dataset.cmKind = n.localName === "r" ? "ref" : n.localName === "commentRangeStart" ? "start" : n.localName === "commentRangeEnd" ? "end"
        : n.localName === "bookmarkStart" ? "bm-start" : "bm-end"; // zakładka — osobne numery id niż komentarze
      span.contentEditable = "false";
      span.className = "cm-mark";
      if (afterObj?.offset === offset) { afterObj.el.after(span); afterObj.el = span; return; }
      if (!groups.has(offset)) groups.set(offset, []);
      groups.get(offset).push(span);
      return;
    }
    afterObj = null;
    Array.from(n.childNodes).forEach((c) => {
      if (c.localName === "t") offset += (c.textContent || "").length;
      else if (c.localName === "br" || c.localName === "tab") offset += 1;
    });
  });
  groups.forEach((spans, at) => {
    const frag = document.createDocumentFragment();
    spans.forEach((sp) => frag.appendChild(sp));
    const range = at > 0 ? modelDomRange(el, at) : null;
    // tuż za numerem przypisu położenie wypada W JEGO tekście — znacznik idzie za wyspę
    const inNote = range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest?.('[data-cm-kind="note"]');
    if (inNote && el.contains(inNote)) { range.setStartAfter(inNote); range.collapse(true); }
    if (range) range.insertNode(frag);
    else if (at === 0) el.insertBefore(frag, el.firstChild);
    else el.appendChild(frag);
  });
}

// Podział strony / kolumny W akapicie: model akapitu zna tylko złamanie wiersza — zapis zamieniłby
// podział strony na zwykłe złamanie (a podgląd dzieli taki akapit na dwa). Tylko do odczytu.
// Akapit z SAMYM podziałem strony (bez tekstu) — jak „Podział strony” w Wordzie: nie pisze się
// w nim, ale usuwa się go Backspace na początku następnego / Delete na końcu poprzedniego akapitu
// (doc-selection.js removeBreakParagraph).
function paragraphPageBreakLock(xp) {
  const hasBreak = Array.from(xp.getElementsByTagNameNS(W_NS, "br")).some((br) => /^(page|column)$/.test(br.getAttributeNS(W_NS, "type") || br.getAttribute("w:type") || ""));
  if (!hasBreak) return null;
  const onlyBreak = !Array.from(xp.getElementsByTagNameNS(W_NS, "t")).some((t) => t.textContent)
    && !["drawing", "pict", "object", "tab", "sym", "fldChar", "footnoteReference", "endnoteReference"].some((tag) => xp.getElementsByTagNameNS(W_NS, tag).length);
  return onlyBreak ? "lockPageBreakOnly" : "lockPageBreak";
}

function paragraphNestedRunLock(xp) {
  const runs = xp.getElementsByTagNameNS(W_NS, "r");
  for (let i = 0; i < runs.length; i++) {
    const parent = runs[i].parentNode;
    if (parent === xp) continue;
    if (parent.localName === "hyperlink" && parent.parentNode === xp) continue;
    const sdt = parent.localName === "sdtContent" ? parent.parentNode : null;
    if (sdt && sdt.parentNode === xp && typeof formIslandSdt === "function" && formIslandSdt(sdt)) continue; // pole-wyspa
    return parent.localName === "hyperlink" ? "lockLink" : "lockField";
  }
  return null;
}

// Etykietka ekranowa linku (w:tooltip, Word: „Etykietka ekranowa…”) — podgląd jej nie rysuje;
// <a> dostaje ją do podpowiedzi po najechaniu i do okienka „Zmień link”. Każdy akapit, też tylko do odczytu.
function stampLinkTips(xp, el) {
  const links = Array.from(xp.getElementsByTagNameNS(W_NS, "hyperlink")).filter((n) => n.getElementsByTagNameNS(W_NS, "t").length);
  const anchors = Array.from(el.querySelectorAll("a[href]:not(.doc-xref)"));
  if (!links.length || links.length !== anchors.length) return;
  links.forEach((h, i) => {
    const tip = h.getAttributeNS(W_NS, "tooltip") || "";
    const a = anchors[i];
    if (!tip) { delete a.dataset.dwbTip; return; }
    a.dataset.dwbTip = tip;
    if (!a.dataset.dwbImgLink) { a.dataset.hint = ""; a.dataset.hintPl = tip; a.dataset.hintEn = tip; }
  });
}

// Linki w podglądzie ↔ w pliku (w tej samej kolejności): <a> dostaje odwołanie z pliku,
// żeby zapis akapitu odtworzył DOKŁADNIE ten link. Gdy się nie zgadzają — tylko do odczytu.
function stampParagraphLinks(xp, el) {
  const links = Array.from(xp.childNodes).filter((n) => n.localName === "hyperlink" && n.namespaceURI === W_NS && n.getElementsByTagNameNS(W_NS, "t").length);
  const anchors = Array.from(el.querySelectorAll("a[href]:not(.doc-xref)"));
  if (!links.length && !anchors.length) return true;
  if (links.length !== anchors.length) return false;
  links.forEach((h, i) => {
    const ref = hyperlinkRef(h);
    anchors[i].dataset.dwbLink = ref;
    if (ref.startsWith("rel:")) docLinkHrefs.set(ref, anchors[i].getAttribute("href") || "");
  });
  return true;
}
async function markLockedParagraphs(bytes) {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host || !bytes) return;
  const doc = await getDocumentXmlDom(bytes);
  if (!doc || bytes !== originalFileBytes) return; // w międzyczasie inny plik/wersja
  // pola-wyspy: XML każdej kontrolki pod jej kluczem („s<N>” = N-te w:sdt, jak w docx-forms.js),
  // a opakowania pól w podglądzie muszą już być (paintFormFields) — inaczej akapit tylko do odczytu
  docIslandXml.clear();
  const sdts = Array.from(doc.getElementsByTagNameNS(W_NS, "sdt"));
  const islandKey = new Map();
  sdts.forEach((sdt, n) => {
    if (typeof formIslandSdt !== "function" || !formIslandSdt(sdt)) return;
    docIslandXml.set(`s${n}`, new XMLSerializer().serializeToString(sdt));
    islandKey.set(sdt, `s${n}`);
  });
  if (islandKey.size && typeof refreshFormScan === "function") {
    await refreshFormScan().catch(() => {});
    if (bytes !== originalFileBytes) return;
  }
  // przypisy: numery jak w Wordzie i treść do dymków — PRZED oznaczaniem (długości numerów)
  if (typeof dwbNotes !== "undefined") {
    await dwbNotes.prepare(bytes).catch((e) => log(`Przypisy: ${e.message || e}`, "error"));
    if (bytes !== originalFileBytes) return;
  }
  const previews = collectPreviewParagraphElements(host);
  const boxes = new Map(); // rodzic akapitu w pliku → numer „pojemnika”
  let commentKey = 0;
  const nextKey = () => `c${commentKey++}`;
  collectParagraphElements(doc.documentElement, "all").forEach((xp, i) => {
    const el = previews[i];
    if (!el) return;
    if (!boxes.has(xp.parentNode)) boxes.set(xp.parentNode, String(boxes.size));
    el.dataset.box = boxes.get(xp.parentNode);
    // rysunki: wyspy, gdy się da (paragraphObjectLock) — wtedy blokują tylko inne powody
    const objLock = paragraphObjectLock(xp);
    const hit = INLINE_LOCK_TAGS.find(([tag, why]) => (why !== "lockObject" || objLock) && xp.getElementsByTagNameNS(W_NS, tag).length);
    let lock = formParagraphLock(xp) || hit?.[1] || paragraphSymbolLock(xp) || paragraphPageBreakLock(xp) || paragraphNestedRunLock(xp) || paragraphCommentLock(xp); // pole formularza Worda: docx-forms.js
    stampLinkTips(xp, el);
    if (!lock && !stampParagraphLinks(xp, el)) lock = "lockLink";
    if (!lock) {
      const keys = Array.from(xp.childNodes).filter((n) => islandKey.has(n)).map((n) => islandKey.get(n));
      if (keys.some((k) => !el.querySelector(`.ff-field[data-ff="${k}"]`))) lock = "lockForm";
    }
    const hasRef = xp.getElementsByTagNameNS(W_NS, "footnoteReference").length || xp.getElementsByTagNameNS(W_NS, "endnoteReference").length;
    if (!lock && hasRef) lock = typeof dwbNotes !== "undefined" ? dwbNotes.stampRefs(xp, el, nextKey) : "lockNote";
    if (!lock) lock = stampSymbolIslands(xp, el, nextKey);
    if (!lock) stampCommentMarks(xp, el, nextKey, hasRef ? dwbNotes.refLabelLengths(el) : []);
    if (lock) el.dataset.lock = lock;
    else delete el.dataset.lock;
  });
  // druga połówka akapitu z podziałem strony — ta sama blokada co pierwsza
  let owner = null;
  host.querySelectorAll("section.docx > article p").forEach((p) => {
    if (!p.dataset.dwbCont) { owner = p; return; }
    p.dataset.lock = owner?.dataset.lock || "lockPageBreak";
  });
  if (typeof dwbNotes !== "undefined") dwbNotes.setupList(nextKey); // tekst przypisów na dole strony
  if (typeof composeUi !== "undefined") {
    composeUi.paintCommentHighlights(); // podświetlenie komentowanego tekstu
    composeUi.fixPreviewPageNumbers(); // numery stron w nagłówkach/stopkach podglądu
  }
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

// Zmiany do zapisu: akapity treści ({ index, runs }) + przypisy ({ note, paras }, doc-notes.js).
// Wszyscy odbiorcy (Zapisz, Cofnij, szkic, karty, przerysowanie) podają listę dalej jako
// op „paragraphBatch” — buildPatchedDocx kieruje przypisy do footnotes.xml / endnotes.xml.
function collectInlineParagraphEdits() {
  const body = collectBodyParagraphEdits();
  if (inlineLocksPending || typeof dwbNotes === "undefined") return body;
  const notes = dwbNotes.collectEdits();
  return notes.length ? [...body, ...notes] : body;
}

function collectBodyParagraphEdits() {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host || !baselineParagraphRuns.length || inlineLocksPending) return [];
  const previews = collectPreviewParagraphElements(host);
  const edits = [];
  if (inlineDirtyValid && previews.length === baselineParagraphRuns.length) {
    flushInlineDirty();
    const indexOf = new Map(previews.map((p, i) => [p, i]));
    inlineDirtyParas.forEach((p) => {
      const i = indexOf.get(p);
      if (i === undefined) { inlineDirtyParas.delete(p); return; } // odłączony (np. scalony)
      if (p.dataset.lock) { inlineDirtyParas.delete(p); return; } // tylko do odczytu (znaczek pola itp.)
      const domRuns = extractRunsFromPreviewParagraph(p);
      if (runsEqual(domRuns, baselineParagraphRuns[i])) inlineDirtyParas.delete(p); // wrócił do stanu z pliku
      else edits.push({ index: i, runs: domRuns });
    });
    return edits.sort((a, b) => a.index - b.index);
  }
  const len = Math.min(previews.length, baselineParagraphRuns.length);
  for (let i = 0; i < len; i++) {
    // Akapit tylko do odczytu: podgląd nie ma w nim pełnej treści pliku (np. tekst kontrolki
    // formularza), więc porównanie z bazą zawsze wyszłoby „zmieniony” — i zapis by go rozbił.
    if (previews[i].dataset.lock) continue;
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
  // przypisy — pełna ścieżka (przebudowa pliku), szybka zna tylko akapity treści
  if ((snapshotEdits || []).some((e) => e.note) || (typeof dwbNotes !== "undefined" && dwbNotes.collectEdits().length)) return false;
  const want = new Map((snapshotEdits || []).map((e) => [e.index, e.runs]));
  const touch = new Set([...collectInlineParagraphEdits().map((e) => e.index), ...want.keys()]);
  // akapit z polem formularza: odbudowa z fragmentów zgubiłaby opakowanie pola — rysujemy od nowa
  for (const i of touch) {
    const runs = want.has(i) ? want.get(i) : pristineParas.get(previews[i]) ? null : baselineParagraphRuns[i];
    // akapit z wyspą (pole, komentarz, przypis): kopia z renderu nie ma oznaczeń wysp — rysujemy od nowa
    if (runs?.some((r) => r.island || r.tab) || baselineParagraphRuns[i]?.some((r) => r.island || r.tab)) return false;
  }
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

// Zamiana słowa przez przeglądarkę: podpowiedź z paska klawiatury (iPhone, Android), poprawka
// pisowni z menu pod prawym przyciskiem, autokorekta. Nowe słowo jest w e.dataTransfer (e.data
// bywa puste), a zamieniany zakres w e.getTargetRanges() — kursor może stać gdzie indziej.
// Dawniej brany był e.data → pusty tekst wstawiany w miejsce zaznaczonego słowa: słowo znikało,
// zostawała dziura (zgłoszenie Mateusza 2026-10-05, telefon i komputer). Teraz: zakres z
// przeglądarki, stare słowo usunięte, nowe w formacie tego miejsca (jak przy pisaniu).
function replaceWordFromInput(e, p) {
  const text = e.data || e.dataTransfer?.getData("text/plain") || "";
  if (!text) return; // nie wiemy, co wstawić — niech zrobi to przeglądarka
  const target = e.getTargetRanges?.()?.[0];
  const sel = window.getSelection();
  let range = null;
  if (target && p.contains(target.startContainer) && p.contains(target.endContainer)) {
    range = document.createRange();
    range.setStart(target.startContainer, target.startOffset);
    range.setEnd(target.endContainer, target.endOffset);
  } else if (sel?.rangeCount && !sel.isCollapsed && p.contains(sel.getRangeAt(0).startContainer)) {
    range = sel.getRangeAt(0).cloneRange(); // bez zakresu od przeglądarki — zaznaczone słowo
  }
  // bez zakresu i bez zaznaczenia nie wiemy, które słowo zamienić — robi to przeglądarka (wie)
  if (!range) return;
  e.preventDefault();
  sel.removeAllRanges();
  sel.addRange(range);
  // format jak u początku zamienianego słowa (nie spacji przed nim)
  const at = range.cloneRange();
  at.collapse(true);
  sel.removeAllRanges();
  sel.addRange(at);
  const style = getInheritedRunStyleAtCaret(p);
  sel.removeAllRanges();
  sel.addRange(range);
  if (runStyleHasProps(style)) insertStyledTextAtCaret(text, style, p);
  else insertTextAtCaret(text);
  onInlineParagraphInput();
  p.dispatchEvent(new Event("input", { bubbles: true })); // jak po zwykłym pisaniu (granice stron, szkic)
}

function onParagraphBeforeInput(e) {
  if (readOnlyMode || e.defaultPrevented) return; // już obsłużone (np. pisanie na krawędzi linku — compose-ui.js)
  // zaznaczenie przez kilka akapitów, Backspace na granicy akapitów itp. (doc-selection.js)
  if (typeof dwbSel !== "undefined" && dwbSel.beforeInput(e)) return;
  if (e.isComposing) return;
  if (e.inputType !== "insertText" && e.inputType !== "insertReplacementText") return;
  const p = docCaretParagraph(e.target);
  if (!p) return;
  if (e.inputType === "insertReplacementText") { replaceWordFromInput(e, p); return; }
  const ch = e.data || "";

  if (ch && getSnippetExpandMode() === "auto" && SNIPPET_EXPAND_DELIMITER_RE.test(ch)) {
    const before = getTextBeforeCaret(p);
    const m = before.match(SNIPPET_TRIGGER_AT_END_RE);
    const sn = m ? getSnippetByName(m[1]) : null;
    if (sn) {
      e.preventDefault();
      // wspólna droga (snippet-suggest.js): pola {{…}}, {cursor}, osobny krok cofania
      expandSnippetAtCaret(p, sn, m[0].length, ch);
      return;
    }
    // iOS zjadł spację przed „!” (snippet-suggest.js) — wyzwalacz mimo to, spacja wraca
    const eaten = !m && typeof window.snippetEatenSpaceQuery === "function" ? window.snippetEatenSpaceQuery(p, before) : null;
    const sn2 = eaten ? getSnippetByName(eaten) : null;
    if (sn2) {
      e.preventDefault();
      expandSnippetAtCaret(p, sn2, eaten.length + 1, ch, " ");
      return;
    }
  }

  // litera wpisana z kursorem w środku pola/symbolu — obok wyspy, nie w niej
  const selNow = window.getSelection();
  if (selNow?.rangeCount && selNow.isCollapsed && caretOutOfIsland(selNow.getRangeAt(0), p) && !runStyleHasProps(mergeRunStyles(getInheritedRunStyleAtCaret(p), currentTypingStyle()))) {
    e.preventDefault();
    insertTextAtCaret(ch);
    onInlineParagraphInput();
    return;
  }
  const inherited = getInheritedRunStyleAtCaret(p);
  const typing = currentTypingStyle();
  const style = mergeRunStyles(inherited, typing);
  if (!runStyleHasProps(style)) return;
  e.preventDefault();
  insertStyledTextAtCaret(ch, style, p, { turnOff: typing });
  if (activeTypingStyle) typingStyleAt = caretPoint(); // kursor poszedł za wpisaną literą — format dalej obowiązuje
  onInlineParagraphInput();
}

// ── format „dla dalszego pisania” ────────────────────────────────────────────
// Rozmiar/krój/kolor wybrany BEZ zaznaczenia obowiązuje tylko w tym miejscu kursora (jak
// w Wordzie): wpisany tam tekst go dostaje, a przestawienie kursora go porzuca. Dawniej zostawał
// na zawsze — rozmiar 9 ustawiony w jednym zdaniu „przyklejał się” do pisania w każdym innym
// akapicie, choć lista rozmiarów pokazywała już rozmiar nowego miejsca.
let typingStyleAt = null; // { node, offset } — miejsce, w którym obowiązuje activeTypingStyle

function caretPoint() {
  const sel = window.getSelection?.();
  if (!sel?.rangeCount) return null;
  const r = sel.getRangeAt(0);
  return { node: r.startContainer, offset: r.startOffset, range: !r.collapsed };
}

function setTypingStyle(props) {
  activeTypingStyle = mergeRunStyles(currentTypingStyle() || {}, props); // z innego miejsca — porzucony
  typingStyleAt = caretPoint();
}

function clearTypingStyle() {
  activeTypingStyle = null;
  typingStyleAt = null;
}

function onTypingStyleSelectionChange() {
  if (!activeTypingStyle || !typingStyleAt) return;
  const now = caretPoint();
  if (!now || !docCanvasEl?.contains(now.node)) return; // fokus w pasku/liście — kursor w tekście ten sam
  if (now.range || now.node !== typingStyleAt.node || now.offset !== typingStyleAt.offset) clearTypingStyle();
}

// Format dla pisania w BIEŻĄCYM miejscu — sprawdzany w chwili użycia: „selectionchange” przychodzi
// z opóźnieniem, a litera wpisana zaraz po strzałce dostawała jeszcze format starego miejsca.
function currentTypingStyle() {
  onTypingStyleSelectionChange();
  return activeTypingStyle;
}

// ── krój i rozmiar w miejscu kursora (pasek formatu + pasek stanu) ─────────────
// Jak pola kroju i rozmiaru na wstążce Worda: zawsze to, co naprawdę jest w miejscu kursora
// (z dziedziczeniem po stylu akapitu), puste przy zaznaczeniu z kilkoma różnymi.
// Dawniej lista rozmiarów zmieniała się tylko przy fragmencie z WŁASNYM rozmiarem — w akapicie
// z rozmiarem ze stylu zostawała poprzednia wartość (np. 9 z innego zdania).
const WORD_FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const COMMON_FONT_FAMILIES = ["Aptos", "Arial", "Calibri", "Calibri Light", "Cambria", "Courier New", "Garamond", "Georgia", "Segoe UI", "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana"];

function docPreviewHost() {
  return docCanvasEl?.querySelector(".docx-preview-host") || null;
}

// Zakres zaznaczenia/kursora w dokumencie; gdy fokus uciekł do paska — ostatni zapamiętany.
function docSelectionRange({ remembered = true } = {}) {
  const host = docPreviewHost();
  if (!host) return null;
  const sel = window.getSelection?.();
  if (sel?.rangeCount) {
    const r = sel.getRangeAt(0);
    if (host.contains(r.startContainer)) return r;
  }
  if (remembered && lastDocCaret?.p?.isConnected && host.contains(lastDocCaret.p)) return lastDocCaret.range;
  return null;
}

function docSelectionIsRange() {
  const r = docSelectionRange();
  return !!r && !r.collapsed;
}

// Tekst treści (bez nagłówka/stopki strony i znaczków list) — liczy się do kroju/rozmiaru.
function isDocTextNode(n) {
  const el = n?.parentElement;
  return !!el && !!el.closest("p") && !el.closest("header, footer, .dwb-page-break");
}

function currentTextFormat() {
  const range = docSelectionRange();
  if (!range) return null;
  const fmt = textFormatOfRange(range, isDocTextNode);
  if (!fmt) return null;
  const pending = range.collapsed ? currentTypingStyle() : null;
  if (pending) {
    if (pending.fontSize) fmt.sizePt = cssFontSizeToPt(pending.fontSize) || fmt.sizePt;
    if (pending.fontFamily) fmt.family = pending.fontFamily;
    fmt.families = [fmt.family];
    fmt.sizes = [fmt.sizePt];
  }
  return fmt;
}

function fmtPt(pt) {
  return String(pt).replace(".", currentLang === "pl" ? "," : ".");
}

// Wyróżnienie akapitu z kursorem (dawniej :focus akapitu-pola; teraz polem jest cały dokument).
// Przy zaznaczeniu przez kilka akapitów bez wyróżnienia — widać samo zaznaczenie, jak w Wordzie.
let caretParaEl = null;
function markCaretParagraph() {
  const sel = window.getSelection?.();
  let p = null;
  if (sel?.rangeCount && !readOnlyMode) {
    const r = sel.getRangeAt(0);
    const at = (n) => (n?.nodeType === 1 ? n : n?.parentElement)?.closest?.(".docx-editable-p");
    const a = at(r.startContainer);
    if (a && a === at(r.endContainer) && a.closest(".docx-edit-root")) p = a;
  }
  if (p === caretParaEl) return;
  caretParaEl?.classList.remove("dwb-caret-p");
  caretParaEl = p;
  p?.classList.add("dwb-caret-p");
  if (p) fixHangingBox(p); // ramka akapitu z wysuniętym 1. wierszem (dawniej przy fokusie akapitu)
}

let formatSyncRaf = 0;
function scheduleFormatSync() {
  if (formatSyncRaf) return;
  formatSyncRaf = requestAnimationFrame(() => { formatSyncRaf = 0; syncFormatIndicators(); });
}

function syncFormatIndicators() {
  syncBiuButtons();
  const fmt = originalFileBytes ? currentTextFormat() : null;
  const sizeSel = document.getElementById("fmtFontSize");
  const famSel = document.getElementById("fmtFontFamily");
  if (sizeSel && document.activeElement !== sizeSel) {
    const pt = fmt?.sizePt || 0;
    let cur = sizeSel.querySelector('option[data-cur]');
    if (pt && !WORD_FONT_SIZES.includes(pt)) {
      if (!cur) { cur = document.createElement("option"); cur.dataset.cur = "1"; sizeSel.insertBefore(cur, sizeSel.querySelector('option[value="__custom"]')); }
      cur.value = String(pt); cur.textContent = fmtPt(pt);
    } else cur?.remove();
    sizeSel.value = pt ? String(pt) : "";
  }
  if (famSel && document.activeElement !== famSel) {
    const fam = fmt?.family || "";
    let cur = famSel.querySelector('option[data-cur]');
    if (fam && ![...famSel.options].some((o) => o.value === fam && !o.dataset.cur)) {
      // tuż za pustą pozycją (pozostałe są w grupach <optgroup> — nie dzieci listy)
      if (!cur) { cur = document.createElement("option"); cur.dataset.cur = "1"; famSel.insertBefore(cur, famSel.firstElementChild?.nextSibling || null); }
      cur.value = fam; cur.textContent = fam;
    } else cur?.remove();
    famSel.value = fam;
    famSel.style.fontFamily = fam ? `"${fam}", var(--font-ui, system-ui)` : "";
  }
  const status = document.getElementById("statusFont");
  if (status) {
    if (!fmt) { status.textContent = ""; status.hidden = true; return; }
    const fams = fmt.families.filter(Boolean);
    const famText = fmt.family || (fams.length > 1 ? t("fontMixed", { n: fams.length }) : fams[0] || "");
    const sizes = fmt.sizes.filter(Boolean);
    const sizeText = fmt.sizePt ? `${fmtPt(fmt.sizePt)} pt` : sizes.length > 1 ? `${fmtPt(sizes[0])}–${fmtPt(sizes[sizes.length - 1])} pt` : "";
    status.textContent = [famText, sizeText].filter(Boolean).join(" · ");
    status.hidden = !status.textContent;
  }
}

// Lista krojów: najpierw te z dokumentu (word/fontTable.xml i użyte w podglądzie), potem popularne.
let fontListKey = "";
async function refreshFontFamilyList() {
  const famSel = document.getElementById("fmtFontFamily");
  if (!famSel || !originalFileBytes) return;
  const docFonts = new Set();
  try {
    const zip = await loadDocxZipCached(originalFileBytes);
    const xml = await zip.file("word/fontTable.xml")?.async("string");
    if (xml) for (const m of xml.matchAll(/<w:font\b[^>]*\bw:name="([^"]+)"/g)) docFonts.add(m[1].replace(/&amp;/g, "&").replace(/&quot;/g, '"'));
  } catch (_) { /* bez listy z pliku */ }
  // kroje symboli i dalekowschodnie z tabeli Worda nie są do pisania tekstu
  const skip = /^(symbol|wingdings.*|webdings|ms mincho|ms gothic|simsun|mangal|batang|mingliu.*|ms ?pmincho|yu .*|dengxian.*|times new roman cyr)$/i;
  const docList = [...docFonts].filter((f) => !skip.test(f)).sort((a, b) => a.localeCompare(b));
  const key = `${currentLang}|${docList.join("|")}`;
  if (key === fontListKey) return;
  fontListKey = key;
  famSel.replaceChildren();
  const blank = new Option("", "");
  blank.hidden = true;
  famSel.appendChild(blank);
  const group = (label, list) => {
    if (!list.length) return;
    const g = document.createElement("optgroup");
    g.label = label;
    list.forEach((f) => { const o = new Option(f, f); o.style.fontFamily = `"${f}"`; g.appendChild(o); });
    famSel.appendChild(g);
  };
  group(t("fontGroupDoc"), docList);
  group(t("fontGroupCommon"), COMMON_FONT_FAMILIES.filter((f) => !docFonts.has(f)));
  famSel.appendChild(new Option(t("fontOther"), "__custom"));
  syncFormatIndicators();
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
  const active = docCaretParagraph(document.activeElement);
  if (active) return active;
  const saved = lastDocCaret;
  if (!saved || !saved.p.isConnected || !saved.p.classList.contains("docx-editable-p")) return null;
  focusDocParagraph(saved.p);
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

// Wklejanie: czysty tekst w formacie miejsca kursora (jak „Zachowaj tylko tekst” w Wordzie).
// Bez tego przeglądarka wklejała surowy HTML z Worda/strony — obce style, a akapity z
// wklejanego tekstu jako <p> w środku akapitu (rozjeżdżało numerację akapitów z plikiem).
// Jeden wiersz — w miejscu kursora, bez przerysowania. Kilka wierszy — jak w Wordzie każdy to
// osobny akapit z formatem akapitu, w który wklejamy (w liście: kolejne punkty), puste wiersze
// = puste akapity (decyzja Mateusza 2026-10-05; dawniej łamania wiersza w jednym akapicie).
// Całość = jeden krok cofania.
function onDocPaste(e) {
  if (readOnlyMode) return;
  // wklejenie w miejsce zaznaczenia kilku akapitów: najpierw usuwamy zaznaczenie (doc-selection.js)
  if (typeof dwbSel !== "undefined") dwbSel.collapseForInsert();
  const p = docCaretParagraph(e.target);
  if (!p) return;
  const commentFragment = e.clipboardData?.getData("application/x-dwb-comment-fragment");
  if (commentFragment && insertInternalCommentFragment(commentFragment)) {
    e.preventDefault();
    onInlineParagraphInput();
    return;
  }
  // Apple Notes, Markdown, strony WWW, Google Docs, Word: struktura i proste style (paste-rich.js).
  // Zwykły tekst bez formatowania — dalej niżej, jak dawniej.
  const rich = typeof dwbPaste !== "undefined" ? dwbPaste.parse(e.clipboardData) : null;
  if (rich) {
    e.preventDefault();
    dwbPaste.apply(p, rich).catch((err) => log(`Wklejanie: ${err.message || err}`, "error"));
    return;
  }
  const text = e.clipboardData?.getData("text/plain");
  if (text == null) return;
  e.preventDefault();
  const lines = text.replace(/\r\n?/g, "\n").replace(/\n+$/, "").split("\n");
  if (lines.length > 1 && typeof dwbPaste !== "undefined") {
    const blocks = lines.map((line) => ({ type: "p", runs: line ? [{ text: line }] : [] }));
    dwbPaste.apply(p, blocks, { keepPara: true }).catch((err) => log(`Wklejanie: ${err.message || err}`, "error"));
    return;
  }
  asUndoStep("undoOpPaste", () => {
    lines.forEach((line, i) => {
      if (i) document.execCommand("insertLineBreak");
      if (line) document.execCommand("insertText", false, line);
    });
  });
  onInlineParagraphInput();
}

// Własny fragment ze schowka po Wytnij: markerom komentarza odpowiada XML w docIslandXml,
// więc można je bezpiecznie przenieść wyłącznie wewnątrz tego samego otwartego dokumentu.
// Nie używamy go dla zwykłego HTML-a z systemowego schowka — zewnętrzna treść nadal przechodzi
// przez sanitizację w paste-rich.js.
function insertInternalCommentFragment(html) {
  const sel = window.getSelection();
  if (!sel?.rangeCount) return false;
  const template = document.createElement("template");
  template.innerHTML = html;
  const marks = [...template.content.querySelectorAll("[data-cm]")];
  if (!marks.length || marks.some((mark) => !docIslandXml.has(mark.dataset.cm))) return false;
  marks.forEach((mark) => {
    mark.contentEditable = "false";
    mark.className = "cm-mark";
  });
  const range = sel.getRangeAt(0);
  range.deleteContents();
  const last = template.content.lastChild;
  range.insertNode(template.content);
  if (last) {
    range.setStartAfter(last);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
  }
  return true;
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

// Shift+Enter = JEDNO łamanie wiersza. execCommand("insertLineBreak") przy white-space: pre-wrap
// wstawiał na końcu akapitu (i tuż przed znacznikiem zakładki/komentarza) DWA znaki „\n” — żeby
// nowy wiersz był widoczny — i oba szły do pliku: w Wordzie pusty wiersz za dużo (test chaos).
// Tu: jeden „\n”, a gdy za nim nie ma nic widocznego — znacznik <br data-dwb-ph> tylko dla oka
// (pomijany przy zapisie: extractRunsFromPreviewParagraph, modelDomRange).
function insertLineBreakAtCaret(p) {
  const sel = window.getSelection();
  if (!sel?.rangeCount) return false;
  const range = sel.getRangeAt(0);
  if (!range.collapsed) range.deleteContents();
  caretOutOfIsland(range, p);
  let tn;
  let at;
  if (range.startContainer.nodeType === 3) {
    tn = range.startContainer;
    at = range.startOffset;
    tn.insertData(at, "\n");
  } else {
    tn = document.createTextNode("\n");
    at = 0;
    range.insertNode(tn);
  }
  const caret = document.createRange();
  caret.setStart(tn, at + 1);
  caret.collapse(true);
  const tail = document.createRange();
  tail.setStart(tn, at + 1);
  tail.setEnd(p, p.childNodes.length);
  const frag = tail.cloneContents();
  const visibleAfter = (frag.textContent || "").length > 0 || !!frag.querySelector?.("img, svg, canvas, br:not([data-dwb-ph]), .docx-tab, [data-cm-kind='obj']");
  if (!visibleAfter && !frag.querySelector?.("br[data-dwb-ph]")) {
    const ph = document.createElement("br");
    ph.dataset.dwbPh = "1";
    if (at + 1 < tn.length) tn.splitText(at + 1);
    tn.parentNode.insertBefore(ph, tn.nextSibling);
  }
  sel.removeAllRanges();
  sel.addRange(caret);
  return true;
}

function focusParagraphAtOffset(paraIndex, offset) {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  const el = collectPreviewParagraphElements(host)[paraIndex];
  if (!el) return;
  focusDocParagraph(el);
  const range = document.createRange();
  const sel = window.getSelection();
  let remaining = Math.max(0, offset);
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let node = walker.nextNode();
  while (node) {
    const len = node.length;
    // pole formularza (nieedytowalna „wyspa”) — kursor przed albo za nim, nigdy w środku
    const island = node.parentElement?.closest('[contenteditable="false"]');
    if (island && el.contains(island) && remaining <= len) {
      if (remaining === 0) range.setStartBefore(island); else range.setStartAfter(island);
      range.collapse(true); sel?.removeAllRanges(); sel?.addRange(range); return;
    }
    // Na granicy końca linku kursor stoi ZA linkiem (jak w Wordzie: pisanie nie wydłuża linku)
    if (remaining === len && node.parentElement?.closest("a")) {
      const next = walker.nextNode();
      if (next && !next.parentElement?.closest("a")) { range.setStart(next, 0); range.collapse(true); sel?.removeAllRanges(); sel?.addRange(range); return; }
      if (next) { remaining -= len; node = next; continue; }
      walker.previousNode();
    }
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
  const active = docCaretParagraph(document.activeElement);
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

// Kursor (albo początek zaznaczenia) W ŚRODKU wyspy — pola formularza, symbolu, odnośnika
// przypisu, rysunku (contenteditable=false) — wychodzi przed nią (na samym początku jej tekstu)
// albo za nią. Strzałkami da się wejść kursorem w tekst pola; Enter rozcinał wtedy pole na dwa
// (zdublowana kontrolka, przesunięta numeracja pól i cudze pola w następnych akapitach).
function caretOutOfIsland(range, p) {
  const n = range.startContainer;
  const el = n.nodeType === 1 ? n : n.parentElement;
  let isl = el?.closest?.('[contenteditable="false"]');
  if (!isl || isl === p || !p.contains(isl)) return false;
  // najbardziej zewnętrzna wyspa w akapicie (pole w polu, znaczek w opakowaniu)
  for (let up = isl.parentElement?.closest('[contenteditable="false"]'); up && up !== p && p.contains(up); up = up.parentElement?.closest('[contenteditable="false"]')) isl = up;
  const atStart = document.createRange();
  atStart.setStart(isl, 0);
  atStart.setEnd(range.startContainer, range.startOffset);
  if (atStart.toString().replace(/\uFEFF/g, "").length === 0) range.setStartBefore(isl);
  else range.setStartAfter(isl);
  range.collapse(true);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  return true;
}

function splitParagraphDomAtCaret(p) {
  const sel = window.getSelection();
  if (!sel?.rangeCount) return null;
  const range = sel.getRangeAt(0);
  if (!range.collapsed) range.deleteContents();
  caretOutOfIsland(range, p);
  const tailRange = document.createRange();
  tailRange.setStart(range.startContainer, range.startOffset);
  // koniec TREŚCI akapitu. Dawniej setEndAfter(p.lastChild || p): w pustym akapicie (bez węzłów —
  // puste akapity z pliku, akapit po przerysowaniu) zakres kończył się ZA akapitem i wycinał jego
  // kopię → akapit w akapicie; podgląd miał o akapit więcej niż plik, a wpisany dalej tekst
  // był w podglądzie dwa razy (zgłoszenie 2026-10-05: Enter po wyjściu z listy na końcu dokumentu)
  tailRange.setEnd(p, p.childNodes.length);
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
  focusDocParagraph(el);
  const sel = window.getSelection();
  const range = document.createRange();
  let remaining = Math.max(0, offset);
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let node = walker.nextNode();
  while (node) {
    // wyspa (numer przypisu, pole) — kursor przed albo za nią, nigdy w środku (tam nie da się pisać)
    const island = node.parentElement?.closest('[contenteditable="false"]');
    if (island && island !== el && el.contains(island) && remaining <= node.length) {
      if (remaining === 0) range.setStartBefore(island); else range.setStartAfter(island);
      range.collapse(true);
      sel?.removeAllRanges();
      sel?.addRange(range);
      return;
    }
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
    insertLineBreakAtCaret(p);
    onInlineParagraphInput();
    return;
  }
  e.preventDefault();
  // Enter w PUSTYM punkcie listy kończy listę (jak w Wordzie), zamiast robić kolejny pusty punkt
  if (isListParagraph(p) && !previewRunsToPlainText(extractRunsFromPreviewParagraph(p)) && typeof composeUi !== "undefined" && composeUi.endListAt(p)) return;
  // Enter w PUSTYM akapicie „wyróżnionym” (ramka, cytat) — akapit wraca do Normalnego (wyjście
  // jak z listy); dawniej powstawał kolejny pusty w ramce i nie było jak z niej wyjść
  if (!previewRunsToPlainText(extractRunsFromPreviewParagraph(p)) && typeof composeUi !== "undefined" && composeUi.plainStyleAt?.(p)) return;
  const newP = splitParagraphDomAtCaret(p);
  if (!newP) return;
  prepareEditableParagraph(newP);
  // Kursor w nowym akapicie OD RAZU — dawniej przeskakiwał dopiero po przebudowie pliku,
  // a wszystko wpisane w tym czasie trafiało do starego akapitu.
  placeCaret(newP, 0);
  const beforeRuns = extractRunsFromPreviewParagraph(p);
  const afterRuns = extractRunsFromPreviewParagraph(newP);
  // Enter na końcu nagłówka → dalej zwykły tekst (docx-compose.js)
  const nextNormal = !previewRunsToPlainText(afterRuns) && typeof composeStripNextStyle === "function" && composeStripNextStyle(newP);
  pristineParas.delete(p); // w pliku ten akapit będzie już inny niż przy renderze
  mirrorBaseline(paraIndex, 1, beforeRuns, afterRuns);
  await applyInlineStructuralEdit({
    op: "splitParagraph",
    index: paraIndex,
    before: previewRunsToPlainText(beforeRuns),
    after: previewRunsToPlainText(afterRuns),
    beforeRuns,
    afterRuns,
    nextNormal,
  });
}

async function handleInlineBackspace(p, paraIndex, e) {
  const sel = window.getSelection();
  // Zaznaczenie (np. od początku akapitu myszą) — Backspace kasuje ZAZNACZENIE, nie znak akapitu.
  // Dawniej liczył się tylko początek zaznaczenia: „kursor na początku” → sklejenie z poprzednim
  // akapitem, a zaznaczony tekst zostawał (zgłoszenie 2026-10-04; zaznaczanie od końca działało).
  if (sel?.rangeCount && !sel.isCollapsed && !e.fromDelete) return;
  if (getCaretOffset(p) !== 0) return;
  // Backspace na początku punktu listy najpierw zdejmuje numerację (jak w Wordzie); kolejny skleja
  if (isListParagraph(p) && sel?.isCollapsed && !e.fromDelete && typeof composeUi !== "undefined") {
    e.preventDefault();
    composeUi.endListAt(p);
    return;
  }
  // pusty akapit „wyróżniony” (ramka, cytat) — najpierw zwykły tekst, kolejny Backspace skleja
  if (sel?.isCollapsed && !e.fromDelete && !previewRunsToPlainText(extractRunsFromPreviewParagraph(p)) && typeof composeUi !== "undefined" && composeUi.isBoxParagraph?.(p)) {
    e.preventDefault();
    composeUi.plainStyleAt(p);
    return;
  }
  if (paraIndex <= 0) return;
  e.preventDefault();
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  const paras = collectPreviewParagraphElements(host);
  const prev = paras[paraIndex - 1];
  // poprzedni akapit to sam podział strony — Backspace go usuwa (jak w Wordzie)
  if (prev?.dataset.lock === "lockPageBreakOnly" && typeof dwbSel !== "undefined") {
    dwbSel.removeBreakParagraph(paraIndex - 1, p, "after");
    return;
  }
  // poprzedni akapit tylko do odczytu (pole, link, przypis…) — sklejenie przepisałoby go i zgubiło zawartość
  if (!prev || prev.dataset.lock) {
    if (prev) toast(t(prev.dataset.lock), "info");
    return;
  }
  // różne „pojemniki” w pliku (komórka tabeli, blok kontrolki, treść główna): sklejenie
  // przeniosłoby akapit z jednego do drugiego i zostawiło pusty pojemnik (Word też tu nie skleja)
  if ((prev.dataset.box || "") !== (p.dataset.box || "")) {
    toast(t("mergeAcross"), "info");
    return;
  }
  const joinAt = previewRunsToPlainText(extractRunsFromPreviewParagraph(prev)).length;
  mergeParagraphDom(prev, p);
  placeCaret(prev, joinAt);
  const mergedRuns = extractRunsFromPreviewParagraph(prev);
  pristineParas.delete(prev);
  mirrorBaseline(paraIndex - 1, 2, mergedRuns);
  await applyInlineStructuralEdit({ op: "mergeParagraph", index: paraIndex, mergedRuns });
}

async function handleInlineTab(p, paraIndex, e) {
  // w tabeli Tab / Shift+Tab chodzi po komórkach (compose-ui.js), jak w Wordzie
  if (p.closest("td, th") && typeof composeUi !== "undefined") {
    e.preventDefault();
    composeUi.tableTab(p, e.shiftKey);
    return;
  }
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
  // W mobilnym reflow wcięcie jest przeliczone na styl inline. Po zmianie klasy poziomu
  // odświeżamy ten pojedynczy punkt, aby Tab / „Głębiej” było widoczne bez czekania na render.
  if (typeof refreshMobileReflowIndent === "function") refreshMobileReflowIndent(p);
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
  // Bez zaznaczenia (jak w Wordzie): przełącza format DALSZEGO pisania w tym miejscu. Dawniej szło
  // przez execCommand, którego stan przeglądarki nasze wstawianie liter ignorowało — za pogrubionym
  // słowem nie dało się zacząć pisać bez pogrubienia (Ctrl/⌘+B „nic nie robiło”).
  const range = docSelectionRange({ remembered: false });
  if (range?.collapsed && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest?.(".docx-editable-p")) {
    const on = formatFlagsAt(range)[command];
    setTypingStyle({ [command]: !on });
    syncFormatIndicators();
    return;
  }
  document.execCommand(command, false, null);
  onInlineParagraphInput();
  syncFormatIndicators();
}

// Pogrubienie/kursywa/podkreślenie tak, jak WIDAĆ (styl obliczony — też ze stylu akapitu, np.
// nagłówek): przy kursorze — znak, którego format dostanie pisanie, + format dalszego pisania;
// przy zaznaczeniu — włączone, gdy obejmuje CAŁY zaznaczony tekst (jak przyciski w Wordzie).
function underlinedEl(el) {
  for (let a = el; a && !a.classList?.contains("docx-editable-p"); a = a.parentElement) {
    if (/underline/.test(getComputedStyle(a).textDecorationLine || "")) return true;
    if (a.localName === "p") break;
  }
  return false;
}
function flagsOfEl(el) {
  if (!el) return { bold: false, italic: false, underline: false };
  const cs = getComputedStyle(el);
  const w = cs.fontWeight === "bold" ? 700 : parseInt(cs.fontWeight, 10) || 400;
  return { bold: w >= 600, italic: /italic|oblique/.test(cs.fontStyle), underline: underlinedEl(el) };
}
function formatFlagsAt(range) {
  if (!range) return { bold: false, italic: false, underline: false };
  const elOf = (n) => (n?.nodeType === 1 ? n : n?.parentElement);
  if (range.collapsed) {
    const p = elOf(range.startContainer)?.closest(".docx-editable-p");
    const src = p ? caretStyleSourceNode(p, range.startContainer, range.startOffset) : range.startContainer;
    const f = flagsOfEl(elOf(src));
    const pending = currentTypingStyle();
    ["bold", "italic", "underline"].forEach((k) => { if (pending && pending[k] != null) f[k] = !!pending[k]; });
    return f;
  }
  const out = { bold: true, italic: true, underline: true };
  let any = false;
  const root = range.commonAncestorContainer.nodeType === 1 ? range.commonAncestorContainer : range.commonAncestorContainer.parentElement;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!range.intersectsNode(n) || !/\S/.test(n.data) || !isDocTextNode(n)) continue;
    if (n === range.endContainer && range.endOffset === 0) continue;
    if (n === range.startContainer && range.startOffset >= n.length) continue;
    any = true;
    const f = flagsOfEl(n.parentElement);
    ["bold", "italic", "underline"].forEach((k) => { if (!f[k]) out[k] = false; });
  }
  return any ? out : { bold: false, italic: false, underline: false };
}

// Przyciski B / I / U „wciśnięte”, gdy format jest włączony (kursor, zaznaczenie, Ctrl/⌘+B).
// W karcie komentarza — stan edytora komentarza (execCommand).
function syncBiuButtons() {
  const ids = { bold: "fmtBold", italic: "fmtItalic", underline: "fmtUnderline" };
  const ed = document.activeElement?.closest?.(".cf-rich");
  let flags = null;
  if (ed) {
    flags = {};
    Object.keys(ids).forEach((k) => { try { flags[k] = document.queryCommandState(k); } catch (_) { flags[k] = false; } });
  } else if (originalFileBytes && !readOnlyMode) flags = formatFlagsAt(docSelectionRange());
  Object.entries(ids).forEach(([k, id]) => {
    const b = document.getElementById(id);
    if (!b) return;
    const on = !!flags?.[k];
    b.classList.toggle("is-on", on);
    b.setAttribute("aria-pressed", on ? "true" : "false");
  });
}

// Format znakowy dla zaznaczenia (też przez kilka akapitów) albo — bez zaznaczenia — dla dalszego
// pisania w miejscu kursora. props: { fontSize } / { fontFamily } albo funkcja (węzeł tekstu →
// cechy) — Powiększ/Pomniejsz czcionkę zmienia każdy rozmiar w zaznaczeniu osobno, jak Word.
function applyRunFormat(props) {
  if (readOnlyMode) return false;
  restoreDocCaret(); // lista/przycisk zabrał fokus — wracamy do zaznaczenia w tekście
  const range = docSelectionRange({ remembered: false });
  if (!range) return false;
  if (range.collapsed) {
    const p = (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest?.(".docx-editable-p");
    if (!p) return false;
    setTypingStyle(typeof props === "function" ? props(null) : props);
    syncFormatIndicators();
    return true;
  }
  const editableText = (n) => !!n.parentElement?.closest(".docx-editable-p") && isDocTextNode(n);
  const out = applyRunPropsToRange(range, props, editableText);
  if (!out) return false;
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(out);
  onInlineParagraphInput();
  syncFormatIndicators();
  return true;
}

function applyFontSizePt(pt) {
  if (readOnlyMode) return;
  const n = parseFloat(String(pt || "").replace(",", "."));
  if (!(n >= 1 && n <= 1638)) { syncFormatIndicators(); return; } // Word: 1–1638 pt
  applyRunFormat({ fontSize: `${Math.round(n * 2) / 2}pt` });
}

function applyFontFamily(name) {
  if (readOnlyMode) return;
  const fam = String(name || "").trim().replace(/["';{}<>]/g, "");
  if (!fam) { syncFormatIndicators(); return; }
  applyRunFormat({ fontFamily: fam });
}

// Ctrl/⌘+Shift+> / < — następny / poprzedni rozmiar z listy Worda; Ctrl/⌘+] / [ — o 1 pt.
function stepFontSize(dir, byPoint) {
  const next = (pt) => {
    if (byPoint) return Math.max(1, Math.min(1638, Math.round(pt) + dir));
    if (dir > 0) return WORD_FONT_SIZES.find((s) => s > pt) || Math.min(1638, Math.ceil((pt + 1) / 10) * 10);
    return [...WORD_FONT_SIZES].reverse().find((s) => s < pt) || Math.max(1, Math.floor(pt - 1));
  };
  const fmt = currentTextFormat();
  const base = fmt?.sizePt || fmt?.sizes?.[0] || 11;
  applyRunFormat((n) => {
    const pt = n ? textFormatOfElement(n.parentElement)?.sizePt || base : base;
    return { fontSize: `${next(pt)}pt` };
  });
}

function handleFormatShortcut(e) {
  if (!(e.ctrlKey || e.metaKey) || e.altKey) return false;
  // rozmiar czcionki jak w Wordzie (e.code — niezależnie od układu klawiatury)
  if (docCaretParagraph(e.target)) {
    const grow = e.shiftKey && e.code === "Period" ? [1, false] : e.shiftKey && e.code === "Comma" ? [-1, false]
      : !e.shiftKey && e.code === "BracketRight" ? [1, true] : !e.shiftKey && e.code === "BracketLeft" ? [-1, true] : null;
    if (grow) { e.preventDefault(); stepFontSize(...grow); return true; }
  }
  const cmd = { b: "bold", i: "italic", u: "underline" }[e.key.toLowerCase()];
  if (!cmd || !docCaretParagraph(e.target)) return false;
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
  // Ctrl/⌘+A, Enter/Backspace/Delete/Tab przy zaznaczeniu kilku akapitów (doc-selection.js)
  if (typeof dwbSel !== "undefined" && dwbSel.keydown(e)) return;

  const p = docCaretParagraph(e.target);
  if (!p) return;
  if (p.dataset.noteKey) { dwbNotes.keydown(p, e); return; } // tekst przypisu (doc-notes.js)

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

// Akapit z wysuniętym pierwszym wierszem (wcięcie wiszące — np. spis treści): ramka akapitu
// zaczyna się tam, gdzie KOLEJNE wiersze, a pierwszy wystaje w lewo. Podświetlenie edytowanego
// akapitu, pasek i obrys wyniku szukania zaczynały się więc w środku pierwszego słowa
// („N|agłe zdarzenia”). Ramkę wydłużamy w lewo o wysunięcie (margines −x, odstęp +x) — tekst
// zostaje co do piksela na miejscu. Leniwie (pierwsze najechanie / fokus / skok), bez list.
function fixHangingBox(p) {
  if (!p || p.tagName !== "P" || p.dataset.hang) return;
  p.dataset.hang = "0";
  if (isListParagraph(p)) return; // listy mają znacznik w wysunięciu — osobny styl
  // Widok mobilny: wcięcia akapitu skalowane zmiennymi (softenReflowIndents) — przestawiamy
  // oryginały, żeby wcięcie dalej reagowało na zoom
  const ti0 = parseFloat(p.style.getPropertyValue("--dwb-ti0"));
  if (ti0 < -1) {
    const ml0 = parseFloat(p.style.getPropertyValue("--dwb-ml0")) || 0;
    const pl0 = parseFloat(p.style.getPropertyValue("--dwb-pl0")) || 0;
    p.style.setProperty("--dwb-ml0", `${ml0 + ti0}px`);
    p.style.setProperty("--dwb-pl0", `${pl0 - ti0}px`);
    p.style.setProperty("margin-left", "calc(var(--dwb-ml0) * var(--dwb-ik, 1))");
    p.style.setProperty("padding-left", "calc(var(--dwb-pl0) * var(--dwb-ik, 1))");
    p.dataset.hang = "1";
    return;
  }
  const cs = getComputedStyle(p);
  const hang = -(parseFloat(cs.textIndent) || 0);
  if (!(hang > 1)) return;
  p.style.marginLeft = `${(parseFloat(cs.marginLeft) || 0) - hang}px`;
  p.style.paddingLeft = `${(parseFloat(cs.paddingLeft) || 0) + hang}px`;
  p.dataset.hang = "1";
}
function onHangingProbe(e) {
  const p = e.target.closest?.("p") || docCaretParagraph(e.target);
  if (p && docCanvasEl.contains(p)) fixHangingBox(p);
}

function bindInlineEditKeyboard() {
  if (!docCanvasEl || inlineKeyboardBound) return;
  inlineKeyboardBound = true;
  docCanvasEl.addEventListener("keydown", onDocCanvasKeydown);
  docCanvasEl.addEventListener("paste", onDocPaste);
  // input/beforeinput przychodzą do POLA edycji (pojemnik dokumentu albo akapit przypisu) — delegacja
  docCanvasEl.addEventListener("input", (e) => { if (docCaretParagraph(e.target) || e.target?.classList?.contains("docx-edit-root")) onInlineParagraphInput(); });
  docCanvasEl.addEventListener("beforeinput", onParagraphBeforeInput);
  docCanvasEl.addEventListener("pointerover", onHangingProbe, { passive: true });
  docCanvasEl.addEventListener("focusin", onHangingProbe);
}

// Jeden akapit (nowy po Enterze) — zamiast syncInlineEditMode() na WSZYSTKICH akapitach
// (przy ~6000 akapitach to ~150 ms opóźnienia widocznego Entera, paczka F).
function prepareEditableParagraph(p) {
  if (!p) return;
  // w pojemniku dokumentu akapit dziedziczy edytowalność (bez własnego pola — inaczej zaznaczenie
  // znów kończyłoby się na nim); akapit przypisu (w nieedytowalnej liście) jest osobnym polem
  if (p.parentElement?.isContentEditable) p.removeAttribute("contenteditable");
  else p.contentEditable = "true";
  p.classList.add("docx-editable-p");
  p.classList.toggle("docx-editable-list", isListParagraph(p));
  p.spellcheck = true;
}

function syncInlineEditMode() {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host) return;
  const editable = !readOnlyMode && !!originalFileBytes;
  const root = docEditRoot(host);
  if (root) {
    root.classList.toggle("docx-edit-root", editable);
    if (editable) { root.contentEditable = "true"; root.spellcheck = true; } else root.removeAttribute("contenteditable");
    // części, w których się nie pisze: nagłówek/stopka strony (osobne okienko), listy przypisów
    // (ich akapity to osobne pola — doc-notes.js), znaczniki granic stron
    root.querySelectorAll(":scope > section.docx > header, :scope > section.docx > footer, :scope > section.docx > ol, .dwb-page-break").forEach((el) => {
      if (editable) el.contentEditable = "false"; else el.removeAttribute("contenteditable");
    });
  }
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
      p.dataset.hintTouch = p.dataset.lock === "lockForm" ? "off" : "on";
      if (p.dataset.lock === "lockPageBreakOnly") p.dataset.pbLabel = t("pageBreakOnlyLabel");
      p.dataset.lockHint = "1";
      return;
    }
    if (p.dataset.lockHint) { // blokada zdjęta (np. zmiany zaakceptowane) albo tryb Czytanie
      ["hint", "hintPl", "hintEn", "hintTouch", "lockHint"].forEach((k) => delete p.dataset[k]);
    }
    p.removeAttribute("contenteditable"); // edytowalność z pojemnika dokumentu
    p.classList.toggle("docx-editable-p", editable);
    p.classList.toggle("docx-editable-list", editable && isListParagraph(p));
    p.spellcheck = editable;
  });
  if (editable) {
    bindInlineEditKeyboard();
  }
  // druga połówka akapitu z podziałem strony: nigdy edytowalna, wyjaśnienie jak przy blokadzie
  host.querySelectorAll("p[data-dwb-cont]").forEach((p) => {
    p.contentEditable = "false";
    p.classList.toggle("docx-locked-p", editable);
    if (editable && p.dataset.lock) {
      p.dataset.hint = "";
      p.dataset.hintPl = I18N.pl[p.dataset.lock];
      p.dataset.hintEn = I18N.en[p.dataset.lock];
      p.dataset.hintTouch = "on";
    } else ["hint", "hintPl", "hintEn", "hintTouch"].forEach((k) => delete p.dataset[k]);
  });
  if (typeof dwbNotes !== "undefined") dwbNotes.sync(editable);
  docCanvasEl?.classList.toggle("edit-mode", editable);
  docCanvasEl?.classList.toggle("read-only", readOnlyMode);
  syncFormatToolbar();
  if (pendingInlineCursor) {
    focusParagraphAtOffset(pendingInlineCursor.paraIndex, pendingInlineCursor.offset);
    pendingInlineCursor = null;
  }
}

// Od świeżego renderu do oznaczenia akapitów tylko do odczytu podgląd nie jest edytowalny,
// więc nie ma w nim zmian — a porównanie w tej chwili wzięłoby np. znaczek ☐ starego pola
// (docx-forms.js dorysowuje go do podglądu) za zmianę i przepisało akapit, gubiąc pole
// (odtworzone 2026-10-01: trzy szybkie kliknięcia w ☐ kasowały pole FORMTEXT obok).
let inlineLocksPending = false;
let inlineSetupSeq = 0;

function setupInlineEditingAfterRender() {
  if (!originalFileBytes) return;
  const bytes = originalFileBytes;
  const seq = ++inlineSetupSeq;
  inlineLocksPending = true;
  refreshInlineEditBaseline(bytes)
    .then(() => markLockedParagraphs(bytes))
    .finally(() => { if (seq === inlineSetupSeq) inlineLocksPending = false; }) // tylko ostatni render zdejmuje
    .then(() => { syncInlineEditMode(); refreshFontFamilyList(); });
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
  if (typeof dwbNotes !== "undefined") dwbNotes.rebase();
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
  const askOther = (sel, key, apply) => {
    // „Inny…” — dowolna wartość jak w polu Worda (np. 10,5 pt albo krój spoza listy)
    const v = window.prompt(t(key), "");
    sel.blur();
    if (v && v.trim()) apply(v.trim()); else syncFormatIndicators();
  };
  // A+ / A− — o stopień z listy Worda, każdy rozmiar w zaznaczeniu osobno (stepFontSize)
  [["fmtGrow", 1], ["fmtShrink", -1]].forEach(([id, dir]) => {
    const b = document.getElementById(id);
    b?.addEventListener("mousedown", (e) => e.preventDefault()); // zaznaczenie w tekście zostaje
    b?.addEventListener("click", () => { if (!readOnlyMode) stepFontSize(dir, false); });
  });
  const sizeSel = document.getElementById("fmtFontSize");
  sizeSel?.addEventListener("change", (e) => {
    if (e.target.value === "__custom") askOther(sizeSel, "fontSizeOther", applyFontSizePt);
    else { applyFontSizePt(e.target.value); sizeSel.blur(); }
  });
  const famSel = document.getElementById("fmtFontFamily");
  famSel?.addEventListener("change", (e) => {
    if (e.target.value === "__custom") askOther(famSel, "fontFamilyOther", applyFontFamily);
    else { applyFontFamily(e.target.value); famSel.blur(); }
  });
  // pasek formatu i pasek stanu nadążają za kursorem; format „dla dalszego pisania” znika po przestawieniu
  document.addEventListener("selectionchange", () => { onTypingStyleSelectionChange(); scheduleFormatSync(); markCaretParagraph(); });
}
wireFormatToolbar();
