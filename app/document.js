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
    if (typeof dwbClipComments !== "undefined") dwbClipComments.clear(); // schowek komentarzy dotyczył poprzedniego pliku
    if (typeof dwbClipNotes !== "undefined") dwbClipNotes.clear();
    if (typeof dwbDocSerial !== "undefined") dwbDocSerial++; // wklejka z formatowaniem tylko w tym samym dokumencie
    // zepsute nazwy krojów ze starszych wersji (naprawa w pamięci; trafi do pliku przy zapisie)
    if (typeof repairDocxFontNames === "function") originalFileBytes = await repairDocxFontNames(originalFileBytes);
    pendingDocEdits = [];
    currentFileName = file.name || "document.docx";
    currentFileFresh = !!options.fresh;
    currentFileType = "docx";
    fileHandle = options.handle || null;

    setFileUi(currentFileName, originalFileBytes.byteLength);
    setDirtyState(false);

    await renderCurrentDocument({ newFile: true });
    if (typeof dwbBugReport !== "undefined") dwbBugReport.startSession(originalFileBytes, currentFileName); // „Zgłoś błąd”: początek sesji
    // śledzenie zmian: stan z karty (open-docs) albo z ustawień pliku (Word ze śledzeniem).
    // Bez await: wczytanie nie może się wydłużać — kursor dla nowego dokumentu (createNew) jest
    // ustawiany tuż po powrocie i zużywany przez przygotowanie akapitów do edycji.
    if (typeof dwbTrack !== "undefined") dwbTrack.onDocumentLoaded(options.track || null).catch(() => {});
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
// Tak samo każda zmiana w już pokazanym dokumencie (lista, styl, wyrównanie, „1.⇥”, Enter w
// pustym punkcie — nagranie Mateusza 2026-10-30: ekran rozmywał się przy każdym kroku; Word nie
// miga). Nakładka jest od razu (łapie kliknięcia w trakcie przerysowania, a „hidden” nadal
// znaczy „gotowe”), ale przez pierwsze 0,4 s niewidoczna (.loading-quiet).
let quietRenderOnce = false;

// opts.newFile — pierwszy render świeżo otwartego pliku: na ekranie jest jeszcze POPRZEDNI dokument,
// więc nakładka „Wczytywanie…” ma być widoczna od razu (dawniej znikała na 0,4 s bez śladu).
// Żółte „do sprawdzenia” z OCR w dokumencie? (menu pod prawym przyciskiem: „Usuń żółte zaznaczenia
// OCR”) — liczone po każdym przerysowaniu z pliku w pamięci, menu czyta gotową odpowiedź.
let ocrMarksPresent = false;
function refreshOcrMarksState() {
  const bytes = originalFileBytes;
  getDocumentXmlString(bytes).then((xml) => { if (bytes === originalFileBytes) ocrMarksPresent = documentHasOcrMarks(xml); }).catch(() => {});
}
// Zdejmuje tylko zaznaczenia z OCR (styl „Do sprawdzenia (OCR)”), własne wyróżnienia zostają; jeden krok Cofnij.
async function clearOcrMarks() {
  const n = await applyDocumentEdit({ op: "clearOcrMarks" });
  if (n) toastUndoable(t("ocrMarksCleared", { n }));
  else toast(t("ocrMarksNone"), "info");
  return n;
}

async function renderCurrentDocument(opts = {}) {
  if (!originalFileBytes || currentFileType !== "docx") return;
  const quiet = !opts.newFile && (quietRenderOnce || !!docCanvasEl?.querySelector(".docx-preview-host p"));
  quietRenderOnce = false;
  loadingOverlayEl.classList.toggle("loading-quiet", quiet);
  setLoading(true, t("renderingDoc"));
  const overlayTimer = quiet ? setTimeout(() => loadingOverlayEl.classList.remove("loading-quiet"), 400) : null;
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
    if (typeof semanticDocumentRendered === "function") semanticDocumentRendered();
    // find-replace-workbench.js jest leniwy — bez niego wyjątek zatrzymałby włączenie edycji poniżej
    if (searchQueryEl?.value.trim() && typeof runDocumentSearch === "function") runDocumentSearch();
    setupInlineEditingAfterRender();
    refreshOcrMarksState();
    if (typeof restoreFormatMarks === "function") restoreFormatMarks();
    if (typeof syncMobileDocZoomAfterRender === "function") syncMobileDocZoomAfterRender();
  } finally {
    clearTimeout(overlayTimer);
    setLoading(false);
    loadingOverlayEl.classList.remove("loading-quiet");
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
    if (opts.quiet) quietRenderOnce = true; // nakładka tylko przy długim renderze
    await renderCurrentDocument();
    if (!vp || (!ratio && !anchor)) return;
    // to samo miejsce w dokumencie: ten sam akapit u góry (Widok mobilny ⇄ desktopowy
    // zawija tekst inaczej, więc sama proporcja przewinięcia potrafiła odjechać o strony)
    if (!restoreDocScrollAnchor(anchor)) vp.scrollTop = ratio * Math.max(0, vp.scrollHeight - vp.clientHeight);
    // świeży render nie ma jeszcze odstępów między stronami — bez nich krótki dokument (albo
    // miejsce przy końcu) obcinał przewinięcie; page-breaks wraca do tego akapitu po ich dołożeniu
    // (jak po reloadFromBytes; zmiana widoku gestem szczypania na krótkim dokumencie, 2026-10-05)
    else renderScrollAnchor = { anchor, until: performance.now() + 2500 };
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
  if (typeof syncPageScaleHeightNow === "function") syncPageScaleHeightNow(); // obszar już tak wysoki jak płótno
  const r = p.getBoundingClientRect();
  vp.scrollTop += r.top + anchor.frac * r.height - anchor.gap - vp.getBoundingClientRect().top;
  return true;
}

// Plik do zapisu = plik + zmiany z panelu + wpisany tekst. Pod wyłącznym dostępem do pliku
// (withQuietFile): Enter/Backspace w trakcie nie przebudowuje pliku obok tej budowy.
// Zapis tekstu pola tekstowego w toku (klik w Zapisz/Eksport kończy pisanie przez blur tuż przed
// tym wywołaniem) — najpierw on, inaczej plik szedł bez ostatnio wpisanego tekstu pola.
async function buildDocumentForSave() {
  const tb = typeof dwbTextbox !== "undefined" ? dwbTextbox.pending?.() : null;
  if (tb) await tb.catch(() => {});
  return typeof withQuietFile === "function" ? withQuietFile(buildDocumentForSaveNow) : buildDocumentForSaveNow();
}

// tylko pod withQuietFile — zmiany zbierane synchronicznie na starcie (indeksy = plik)
async function buildDocumentForSaveNow() {
  if (!originalFileBytes) return null;
  const inlineEdits = collectInlineParagraphEdits();
  const edits = [...pendingDocEdits];
  if (inlineEdits.length) edits.push({ op: "paragraphBatch", items: inlineEdits });
  // komentarze wycięte i niewklejone (docx-inline-edit.js, dwbClipComments) — w zapisanym pliku
  // ich nie ma (jak w Wordzie); w otwartym dokumencie dalej czekają na wklejenie
  const clipIds = typeof dwbClipComments !== "undefined" ? [...dwbClipComments] : [];
  // przypisy wycięte i niewklejone (dwbClipNotes) — przy zapisie znikają (forcePrune)
  const clipNotes = typeof dwbClipNotes !== "undefined" && dwbClipNotes.size;
  if (!edits.length && clipNotes) edits.push({ op: "paragraphBatch", items: [] });
  let bytes = edits.length ? (await buildPatchedDocx(originalFileBytes, edits, { forcePrune: true })).bytes : originalFileBytes;
  if (clipIds.length) {
    const refs = commentRefIds(await getDocumentXmlString(bytes));
    const gone = clipIds.filter((id) => !refs.has(id));
    if (gone.length) bytes = (await buildPatchedDocx(bytes, [{ op: "revisions", action: "removeComments", ids: gone }])).bytes;
  }
  return bytes;
}

// Scalenie wpisanego tekstu i zmiana pliku pod wyłącznym dostępem (jedna przebudowa naraz —
// dawniej Enter w trakcie budował na starym pliku obok); odrysowanie już poza nim.
// Klawisze wciśnięte w trakcie zmiany pliku i przerysowania (np. Enter tuż po ◀ ▶ kropek, pisanie
// zaraz po zmianie pola tekstowego) — trzymane i dopisane po przerysowaniu (holdTypingWhile).
// Dawniej trafiały do starego widoku: akapit podzielony w pliku, a w podglądzie nie (test chaos).
function applyDocumentEdit(edit, xmlOpts = {}) {
  if (!originalFileBytes) return Promise.resolve(0);
  return typeof holdTypingWhile === "function" ? holdTypingWhile(() => applyDocumentEditHeld(edit, xmlOpts), -1) : applyDocumentEditHeld(edit, xmlOpts);
}
async function applyDocumentEditHeld(edit, xmlOpts) {
  const res = typeof withQuietFile === "function" ? await withQuietFile(() => applyDocumentEditNow(edit, xmlOpts)) : await applyDocumentEditNow(edit, xmlOpts);
  if (!res.reload) {
    if (res.quiet && res.n > 0) {
      setDirtyState(true); // zmiana zapisana bez przerysowania (ekran już ją pokazuje)
      // miejsce kursora dla przerysowania (runFileEdit) — przerysowania nie było, kursor stoi, gdzie stał;
      // zostawione czekałoby i przestawiło kursor przy następnym, niezwiązanym przerysowaniu
      if (typeof pendingInlineCursor !== "undefined") pendingInlineCursor = null;
      afterQuietDocumentEdit(edit);
    }
    return res.n;
  }
  if (res.n > 0 && paraPatchable(edit) && await patchParagraphsInPlace(edit.indices).catch(() => false)) {
    setDirtyState(true);
    if (typeof pendingInlineCursor !== "undefined") pendingInlineCursor = null; // kursor nietknięty
    return res.n;
  }
  await reloadFromBytes(originalFileBytes);
  if (res.n > 0) setDirtyState(true);
  return res.n;
}

// ── przerysowanie tylko zmienionych akapitów ─────────────────────────────────
// Wyrównanie, wcięcie, odstępy nie zmieniają TREŚCI akapitu, tylko jego wygląd. Dawniej każda taka
// zmiana przerysowywała cały dokument (0,6 s / 8 stron, 1,6 s / 100 stron przy CPU ×4; chwilowy
// przeskok widoku — scripts/reload-bench.js). Teraz: te akapity (z sąsiadami — odstępy zależą od
// sąsiadów) jako mały plik, render w tle TYM SAMYM torem (docx-preview + poprawki), a do żywego
// akapitu przechodzi tylko wygląd: styl, klasy z renderu, atrybuty data-dwb-*. Treść akapitu
// zostaje nietknięta — kursor, zaznaczenie, pola, komentarze i przypisy nie mają jak zginąć.
// Pomiary zależne od miejsca (spacje na końcu, justowanie, ostatnia linijka) — od nowa na miejscu.
// Wszystko inne (styl akapitu — nowe reguły CSS, listy, tabele, kolumny) — pełne przerysowanie.
// Widok mobilny też szybko: jego poprawki (applyMobileReflowLayout, skalowanie wcięć) działają
// wewnątrz renderDocxPreview, więc render w tle ma je tak samo. Pilnuje test
// scripts/para-patch-playwright.js (łatka = pełny render; MOBILE=1 — telefon).
const PARA_PATCH_KEYS = new Set(["op", "indices", "align", "ind", "indentDelta", "spacing"]);
function paraPatchable(edit) {
  return edit?.op === "paraFormat" && Array.isArray(edit.indices) && edit.indices.length > 0 && edit.indices.length <= 50
    && Object.keys(edit).every((k) => PARA_PATCH_KEYS.has(k));
}
// klasy, które dokłada Edycja (nie render) — zostają na żywym akapicie
const PARA_APP_CLASSES = new Set(["docx-editable-p", "dwb-caret-p", "ff-para", "docx-locked-p", "dwb-tiny-active"]);

async function patchParagraphsInPlace(indices) {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host) return false;
  const previews = collectPreviewParagraphElements(host);
  const sorted = [...new Set(indices)].sort((a, b) => a - b);
  const live = sorted.map((i) => previews[i]);
  if (live.some((p) => !p || p.dataset.lock || p.dataset.dwbCont || p.nextElementSibling?.dataset?.dwbCont
    || p.closest("td, th, .dwb-tb, header, footer") || isListParagraph(p) || p.classList.contains("dwb-tiny-para"))) return false;
  // akapity w pliku: wprost w treści, jedna sekcja, kolejne po sobie
  const doc = new DOMParser().parseFromString(await getDocumentXmlString(originalFileBytes), "application/xml");
  const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
  const all = collectParagraphElements(doc.documentElement, "all");
  const xps = sorted.map((i) => all[i]);
  if (!body || xps.some((x) => !x || x.parentNode !== body)) return false;
  if (body.getElementsByTagNameNS(W_NS, "sectPr").length !== 1) return false; // kilka sekcji (kolumny, strony)
  const kids = [...body.childNodes].filter((n) => n.nodeType === 1);
  const lo = kids.indexOf(xps[0]);
  const hi = kids.indexOf(xps[xps.length - 1]);
  if (hi - lo !== xps.length - 1) return false; // nie po kolei (między nimi tabela albo inny akapit)
  const isP = (n) => n?.localName === "p" && n.namespaceURI === W_NS;
  const prev = isP(kids[lo - 1]) ? kids[lo - 1] : null;
  const next = isP(kids[hi + 1]) ? kids[hi + 1] : null;
  const sect = kids.find((n) => n.localName === "sectPr");
  const ser = new XMLSerializer();
  const part = [prev, ...xps, next].filter(Boolean).map((n) => ser.serializeToString(n)).join("") + (sect ? ser.serializeToString(sect) : "");
  const xml = await getDocumentXmlString(originalFileBytes);
  const open = xml.indexOf(">", xml.indexOf("<w:body")) + 1;
  const close = xml.lastIndexOf("</w:body>");
  if (open <= 0 || close < open) return false;
  const zip = await JSZip.loadAsync(originalFileBytes);
  zip.file("word/document.xml", xml.slice(0, open) + part + xml.slice(close));
  const mini = new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
  // render w tle: te same klasy płótna (tryb, zoom), ta sama szerokość
  const src = document.createElement("div");
  src.className = `${docCanvasEl.className} dwb-pp-src`;
  src.style.width = `${docCanvasEl.clientWidth}px`;
  src.setAttribute("aria-hidden", "true");
  document.body.appendChild(src);
  try {
    await renderDocxPreview(mini, src);
    const rendered = collectPreviewParagraphElements(src.querySelector(".docx-preview-host"));
    const off = prev ? 1 : 0;
    if (rendered.length !== xps.length + off + (next ? 1 : 0)) return false;
    live.forEach((p, k) => copyParagraphLook(p, rendered[off + k]));
  } finally {
    src.remove();
  }
  // poprawki zależne od sąsiadów i od miejsca — na żywym dokumencie
  if (typeof applyContextualSpacing === "function") applyContextualSpacing(host);
  if (typeof applyJustifySpaceCompression === "function") applyJustifySpaceCompression(host);
  live.forEach((p) => { if (typeof hangTrailingSpace === "function") hangTrailingSpace(p); });
  if (typeof applyScreenLastLines === "function") applyScreenLastLines(host, live.filter((p) => p.classList.contains("dwb-justify")));
  if (typeof dwbPageBreaks !== "undefined") dwbPageBreaks.schedule(0); // wysokość akapitu mogła się zmienić
  if (typeof composeUi !== "undefined") composeUi.syncState?.(); // przyciski wyrównania / wcięcia
  return true;
}

function copyParagraphLook(p, from) {
  p.setAttribute("style", from.getAttribute("style") || "");
  [...p.classList].forEach((c) => { if (!PARA_APP_CLASSES.has(c)) p.classList.remove(c); });
  from.classList.forEach((c) => { if (!PARA_APP_CLASSES.has(c)) p.classList.add(c); });
  Object.keys(p.dataset).forEach((k) => { if (k.startsWith("dwb") && !(k in from.dataset)) delete p.dataset[k]; });
  Object.keys(from.dataset).forEach((k) => { if (k.startsWith("dwb")) p.dataset[k] = from.dataset[k]; });
  delete p.dataset.hang; // wysunięcie (fixHangingBox) policzy się od nowa przy najechaniu
}

// Operacje, po których STRONA wygląda tak samo — zmieniają tylko części pliku poza treścią
// (komentarze: comments.xml, metadane: docProps/core.xml). Dawniej każda przerysowywała cały
// dokument: 0,6 s na 8 stronach, 1,6 s na 100 (CPU ×4), mignięcie i zgubiony kursor (pomiar
// scripts/reload-bench.js, 2026-11-10). Teraz: zapis do pliku + odświeżenie tego, co od nich zależy.
// Bez odpowiedzi na komentarz: odpowiedź ma w treści WŁASNE znaczniki (commentRangeStart,
// commentReference) — bez przerysowania widok ich nie miał i pisanie w akapicie je gubiło
// (compose4: „pisanie i Enter w akapicie z komentarzem”). Bezpiecznik niżej: cicho tylko wtedy,
// gdy word/document.xml się nie zmienił — każda operacja, która jednak dotknie treści, przerysowuje.
const QUIET_DOC_OPS = new Set(["commentEdit", "commentDone"]);
function docEditIsQuiet(edit) {
  if (QUIET_DOC_OPS.has(edit.op)) return true;
  // metadane — chyba że pole formularza w treści jest powiązane z nimi (pokazuje np. tytuł)
  if (edit.op === "coreMetadata") {
    const core = typeof FORM_STORE_CORE !== "undefined" ? FORM_STORE_CORE.toUpperCase() : null;
    return !(typeof formScan !== "undefined" && formScan?.fields?.some((f) => String(f.binding?.store || "").toUpperCase() === core));
  }
  return false;
}
function afterQuietDocumentEdit(edit) {
  if (QUIET_DOC_OPS.has(edit.op)) {
    if (typeof composeUi !== "undefined") composeUi.paintCommentHighlights?.(); // nowe dane komentarzy → podświetlenia i karty
    if (typeof refreshReviewCounts === "function") refreshReviewCounts(); // licznik i otwarty panel Recenzja
  }
}

async function applyDocumentEditNow(edit, xmlOpts) {
  if (typeof mergeInlineEditsIntoBytesNow === "function") await mergeInlineEditsIntoBytesNow();
  const normalized = edit.op ? edit : { ...edit, op: "replace" };
  // poprawki z różnicy „przed” → teraz (track-changes.js) — cały plik naraz, nie operacja na XML
  if (normalized.op === "redline") {
    if (pendingDocEdits.length) {
      originalFileBytes = (await buildPatchedDocx(originalFileBytes, pendingDocEdits)).bytes;
      pendingDocEdits = [];
    }
    const { bytes, stats } = await dwbRedline(normalized.baseline, originalFileBytes, { author: normalized.author, trackOn: normalized.trackOn });
    originalFileBytes = bytes;
    return { n: stats.changed + stats.added + stats.removed, reload: true };
  }
  // „✦ Przepisz” (doc-rewrite.js): nowy tekst akapitu (albo kilku) jako poprawka (w:ins / w:del) — zmiana
  // i poprawka w jednym kroku Cofnij. Przy włączonym śledzeniu zmiana idzie zwykłą drogą
  // (track-changes.js sam zrobi z niej poprawkę przy zapisie / „Pokaż poprawki”).
  if (normalized.op === "rewriteTracked") {
    if (pendingDocEdits.length) {
      originalFileBytes = (await buildPatchedDocx(originalFileBytes, pendingDocEdits)).bytes;
      pendingDocEdits = [];
    }
    const before = originalFileBytes;
    const items = normalized.items || [{ index: normalized.index, text: normalized.text }]; // kilka akapitów naraz
    const batch = { op: "paragraphBatch", items };
    const { bytes: after, changeCount: n } = await buildPatchedDocx(before, [batch]);
    if (!n) return { n: 0, reload: false };
    const tracking = typeof dwbTrack !== "undefined" && dwbTrack.isOn();
    originalFileBytes = tracking ? after : (await dwbRedline(before, after, { author: normalized.author })).bytes;
    return { n, reload: true };
  }
  recordPendingEdit(normalized);
  const beforeEdit = originalFileBytes;
  const { bytes, changeCount } = await buildPatchedDocx(originalFileBytes, pendingDocEdits, xmlOpts);
  pendingDocEdits = [];
  originalFileBytes = bytes;
  // Akceptuj/Odrzuć podczas śledzenia — to samo w pliku „przed” (bez nowych poprawek)
  if (normalized.op === "revisions" && typeof dwbTrack !== "undefined") await dwbTrack.mirrorRevisions(normalized);
  // quiet: operacja, którą ekran już pokazuje (np. przesunięta ramka pola tekstowego) — bez przerysowania
  if (normalized.quiet && changeCount > 0) return { n: changeCount, reload: false, quiet: true };
  if (docEditIsQuiet(normalized) && changeCount > 0
    && (await getDocumentXmlString(beforeEdit)) === (await getDocumentXmlString(bytes))) return { n: changeCount, reload: false, quiet: true };
  return { n: changeCount, reload: true };
}
