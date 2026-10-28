// undo.js — Cofnij / Ponów dla dokumentu (Ctrl/⌘+Z, Ctrl/⌘+Shift+Z, Ctrl+Y, przyciski ↶ ↷).
// Paczka D (2026-09-28), ta sama rodzina co sheet-workbench-pwa/app/undo.js — ale inny model.
//
// W arkuszu każda zmiana to komórka, więc Sheet zapisuje różnice komórek. Tu dokument to
// plik .docx (originalFileBytes) + niezapisane poprawki akapitów w podglądzie. Każda
// operacja na pliku tworzy NOWE bajty (stare zostają nietknięte), więc krok cofania to
// MIGAWKA: odwołanie do bajtów sprzed zmiany (bez kopiowania) + lista akapitów zmienionych
// w podglądzie względem tych bajtów. Cofnięcie = wróć do tych bajtów, nałóż te akapity,
// narysuj dokument od nowa, oddaj kursor i miejsce przewinięcia.
//
// Kroki:
//   - każda operacja z panelu (applyDocumentEdit: zamiana, korekta, placeholdery, snippety,
//     metadane, narzędzia edycji, szybka edycja w Strukturze) — jeden krok, z nazwą,
//   - formatowanie B / I / U / rozmiar — jeden krok,
//   - pisanie w podglądzie (litery, Enter, Backspace, Tab) — grupowane jak w Wordzie:
//     krok zamyka przerwa ≥ 1,5 s, kliknięcie/tapnięcie w dokument albo inna operacja.
//
// Spójność z kolejką Enter/Backspace (docx-inline-edit.js): migawka czyta bajty dopiero
// po operacjach zakolejkowanych PRZED nią (then na ogonie kolejki — wywołania na tej samej
// obietnicy idą w kolejności rejestracji), a listę akapitów liczy od razu, względem
// lustra bazy, które już te operacje uwzględnia. Czyli bajty i akapity opisują ten sam stan.

const UNDO_TYPING_PAUSE_MS = 1500;
const UNDO_OP_LABEL = {
  replace: "undoOpReplace",
  case: "undoOpCase",
  trim: "undoOpTrim",
  affix: "undoOpAffix",
  paragraphBatch: "undoOpParagraphs",
  textbox: "undoOpTextbox",
  coreMetadata: "undoOpMetadata",
  revisions: "undoOpReview",
  formFill: "undoOpForm",
  placeholderFill: "undoOpPlaceholders",
  snippetExpand: "undoOpSnippets",
  paraFormat: "undoOpParaFormat",
  pageBreak: "undoOpPageBreak",
  hrule: "undoOpHrule",
  link: "undoOpLink",
  bookmark: "undoOpBookmark",
  list: "undoOpList",
  toc: "undoOpToc",
  formInsert: "undoOpFormInsert",
  snippetInsert: "undoOpSnippet",
  pasteBlocks: "undoOpPaste",
  pasteParas: "undoOpPaste",
  pageVAlign: "undoOpFormat",
  pageSetup: "undoOpPageSetup",
  noteInsert: "undoOpNote",
  tableInsert: "undoOpTable",
  table: "undoOpTable",
  imageInsert: "undoOpImage",
  imageReplace: "undoOpImage",
  image: "undoOpImage",
  runStyle: "undoOpFormat",
  commentAdd: "undoOpComment",
  commentReply: "undoOpComment",
  commentDone: "undoOpComment",
  commentEdit: "undoOpComment",
  headerFooter: "undoOpHeaderFooter",
  redline: "undoOpTrack",
  rewriteTracked: "undoOpRewrite",
};

const dwbUndo = (() => {
  const undoBtn = document.getElementById("undoBtn");
  const redoBtn = document.getElementById("redoBtn");
  let undoStack = []; // { state, label }
  let redoStack = [];
  let savedPos = 0; // długość undoStack w chwili zapisu/wczytania; -1 = stan zapisany nieosiągalny
  let applying = false;
  let busy = null; // trwa cofanie/ponawianie — kolejne kliknięcia czekają
  let burst = null; // { lastAt } — trwający krok „Pisanie”

  function limit() {
    // Każdy krok trzyma odwołanie do bajtów sprzed zmiany. Przy dużych plikach mniej kroków.
    const size = originalFileBytes?.byteLength || 0;
    return size > 20 * 1024 * 1024 ? 10 : size > 5 * 1024 * 1024 ? 25 : 60;
  }

  function currentCaret() {
    const p = docCaretParagraph(document.activeElement);
    if (!p || !docCanvasEl?.contains(p)) return null;
    return { paraIndex: resolveParaIndex(p), offset: getCaretOffset(p) };
  }

  function captureState() {
    const tail = typeof waitInlineStructuralIdle === "function" ? waitInlineStructuralIdle() : Promise.resolve();
    return {
      bytesP: tail.then(() => originalFileBytes),
      inline: typeof collectInlineParagraphEdits === "function" ? collectInlineParagraphEdits() : [],
      caret: currentCaret(),
      scrollTop: docViewportEl?.scrollTop || 0,
      anchor: typeof captureDocScrollAnchor === "function" ? captureDocScrollAnchor() : null, // akapit u góry (scrollTop zależy od odstępów stron)
    };
  }

  let grouped = 0; // trwa group(): operacje w środku nie dokładają własnych kroków
  function push(label) {
    if (applying || grouped || !originalFileBytes) return null;
    const entry = { state: captureState(), label };
    undoStack.push(entry);
    if (savedPos >= undoStack.length) savedPos = -1; // zapisany stan był na porzuconej gałęzi (po Cofnij)
    redoStack = [];
    const max = limit();
    while (undoStack.length > max) {
      undoStack.shift();
      savedPos = savedPos > 0 ? savedPos - 1 : -1;
    }
    sync();
    return entry;
  }

  function drop(entry) {
    const i = undoStack.lastIndexOf(entry);
    if (i < 0) return;
    undoStack.splice(i, 1);
    if (savedPos > i) savedPos -= 1;
    sync();
  }

  function endBurst() { burst = null; }

  // Operacja w podglądzie spoza klawiatury (wstawienie z panelu, wklejenie) = JEDEN osobny
  // krok. Dawniej doklejała się do poprzedniego pisania i ↶ zabierało oba naraz.
  let opInProgress = false;
  function record(label, fn) {
    endBurst();
    push(label);
    opInProgress = true;
    try { return fn(); } finally { opInProgress = false; endBurst(); }
  }

  // Kilka zmian (też asynchronicznych, z przerysowaniem) jako JEDEN krok — np. autoformatowanie
  // „1. ” → lista (autocorrect.js): skasowanie znacznika + zmiana w pliku, a Cofnij wraca do „1. ”.
  async function group(label, fn) {
    endBurst();
    push(label);
    grouped += 1;
    try { return await fn(); } finally { grouped -= 1; endBurst(); }
  }

  function typing() {
    if (applying || opInProgress || !originalFileBytes || readOnlyMode) return;
    const now = Date.now();
    if (!burst || now - burst.lastAt > UNDO_TYPING_PAUSE_MS) {
      push("undoOpTyping");
      burst = { lastAt: now };
    } else {
      burst.lastAt = now;
    }
  }

  async function restore(state) {
    if (typeof waitInlineStructuralIdle === "function") await waitInlineStructuralIdle();
    const bytes = await state.bytesP;
    // Szybka ścieżka (paczka F): plik ten sam, zmieniło się tylko pisanie → podmieniamy
    // treść akapitów zamiast rysować cały dokument (przy ~300 stronach 5 s → ułamek).
    if (bytes === originalFileBytes && typeof restoreInlineParagraphs === "function" && restoreInlineParagraphs(state.inline)) {
      if (state.caret && !readOnlyMode) focusParagraphAtOffset(state.caret.paraIndex, state.caret.offset);
      if (!restoreDocScrollAnchor(state.anchor) && docViewportEl) docViewportEl.scrollTop = state.scrollTop;
      return;
    }
    let target = bytes;
    if (state.inline.length) {
      target = (await buildPatchedDocx(bytes, [{ op: "paragraphBatch", items: state.inline }])).bytes;
    }
    pendingDocEdits = [];
    // syncInlineEditMode (po odświeżeniu bazy) ustawi kursor tam, gdzie był
    if (state.caret && !readOnlyMode) pendingInlineCursor = state.caret;
    await reloadFromBytes(target, { anchor: state.anchor });
    if (!state.anchor && docViewportEl) docViewportEl.scrollTop = state.scrollTop;
  }

  function changesSinceSave() {
    return savedPos < 0 ? Math.max(1, undoStack.length) : Math.abs(undoStack.length - savedPos);
  }

  async function step(direction) {
    if (busy) return busy;
    const from = direction === "undo" ? undoStack : redoStack;
    if (!originalFileBytes || !from.length) {
      toast(t(direction === "undo" ? "undoNothing" : "redoNothing"), "info");
      return false;
    }
    endBurst();
    const entry = from.pop();
    // stan „teraz” — trafia na drugi stos, żeby dało się wrócić
    const now = { state: captureState(), label: entry.label };
    (direction === "undo" ? redoStack : undoStack).push(now);
    applying = true;
    busy = (async () => {
      try {
        // nakładka „Cofam…” tylko gdy trwa dłużej (szybka ścieżka nie powinna mrugać)
        const slowTimer = setTimeout(() => setLoading(true, t(direction === "undo" ? "undoWorking" : "redoWorking")), 150);
        try { await restore(entry.state); } finally { clearTimeout(slowTimer); }
        const dirty = changesSinceSave() > 0;
        // przez window.setDirtyState: inne moduły (szkic w drafts.js) też muszą wiedzieć, że
        // cofnięcie wróciło do stanu z pliku; własna obsługa tu pomija zapis (applying=true)
        window.setDirtyState(dirty);
        toast(t(direction === "undo" ? "undoDone" : "redoDone", { what: t(entry.label) }), "info");
        return true;
      } catch (err) {
        log(`Cofnij/Ponów: ${err.message || err}`, "error");
        toast(t("undoFailed"), "error");
        return false;
      } finally {
        setLoading(false);
        applying = false;
        busy = null;
        sync();
      }
    })();
    return busy;
  }

  function sync() {
    const has = !!originalFileBytes;
    const nextUndo = undoStack[undoStack.length - 1];
    const nextRedo = redoStack[redoStack.length - 1];
    if (undoBtn) {
      undoBtn.disabled = !has || !undoStack.length;
      undoBtn.setAttribute("aria-label", nextUndo ? t("undoWhat", { what: t(nextUndo.label) }) : t("undoLabel"));
      undoBtn.dataset.hint = undoBtn.getAttribute("aria-label"); // cursor-hint: ten sam, żywy opis
    }
    if (redoBtn) {
      redoBtn.disabled = !has || !redoStack.length;
      redoBtn.setAttribute("aria-label", nextRedo ? t("redoWhat", { what: t(nextRedo.label) }) : t("redoLabel"));
      redoBtn.dataset.hint = redoBtn.getAttribute("aria-label");
    }
    if (typeof appFrame !== "undefined") appFrame.syncSave();
  }

  function clear() {
    undoStack = [];
    redoStack = [];
    savedPos = 0;
    burst = null;
    sync();
  }

  // ── punkty zaczepienia ─────────────────────────────────────────────────────
  // Operacje z panelu: jeden krok z nazwą; brak zmian = krok znika.
  const origApplyEdit = window.applyDocumentEdit;
  window.applyDocumentEdit = async function applyDocumentEditUndoable(edit, ...rest) {
    endBurst();
    const entry = push(UNDO_OP_LABEL[edit?.op || "replace"] || "undoOpEdit");
    const n = await origApplyEdit.call(this, edit, ...rest);
    if (!(n > 0) && entry) drop(entry);
    return n;
  };
  // Formatowanie zaznaczenia: osobny krok. Bez zaznaczenia (format dla dalszego pisania) nic
  // się w dokumencie nie zmienia — kroku nie ma (dawniej pusty krok „Formatowanie”).
  ["execInlineFormat", "applyFontSizePt", "applyFontFamily", "stepFontSize", "toggleVertAlign", "clearTextFormatting", "changeTextCase"].forEach((name) => {
    const orig = window[name];
    if (typeof orig !== "function") return;
    window[name] = function formatUndoable(...args) {
      // wyczyść formatowanie i wielkość liter zmieniają tekst także bez zaznaczenia (akapit / słowo)
      const changes = name === "execInlineFormat" || name === "clearTextFormatting" || name === "changeTextCase" || typeof docSelectionIsRange !== "function" || docSelectionIsRange();
      if (!readOnlyMode && changes) { endBurst(); push("undoOpFormat"); endBurst(); }
      return orig.apply(this, args);
    };
  });
  // Pisanie: PRZED zmianą w DOM. Enter/Backspace/Tab/Delete mają własną obsługę w keydown
  // (preventDefault), więc łapiemy je w keydown w fazie capture; resztę — beforeinput.
  docCanvasEl?.addEventListener("keydown", (e) => {
    if (e.isComposing || e.ctrlKey || e.metaKey || e.altKey) return;
    if (!docCaretParagraph(e.target)) return;
    if (["Enter", "Backspace", "Delete", "Tab"].includes(e.key)) typing();
  }, true);
  docCanvasEl?.addEventListener("beforeinput", (e) => {
    if (!docCaretParagraph(e.target)) return;
    if (e.inputType === "historyUndo" || e.inputType === "historyRedo") {
      // menu Edycja → Cofnij, potrząśnięcie iPhone'em: cofanie aplikacji, nie przeglądarki
      e.preventDefault();
      step(e.inputType === "historyUndo" ? "undo" : "redo");
      return;
    }
    if (e.inputType.startsWith("format")) return; // execCommand B/I/U — krok już jest
    typing();
  }, true);
  // Kliknięcie/tapnięcie w dokument = nowy krok przy następnym pisaniu.
  docCanvasEl?.addEventListener("pointerdown", endBurst, true);

  // Zapis/wczytanie: setDirtyState(false) poza cofaniem = „ten stan jest zapisany”.
  let cleanBytes = null; // bajty pliku w chwili zapisu/wczytania (bez poprawek w podglądzie)
  const origSetDirty = window.setDirtyState;
  window.setDirtyState = function setDirtyStateUndo(isDirty, ...rest) {
    if (!isDirty && !applying) { savedPos = undoStack.length; endBurst(); cleanBytes = originalFileBytes; }
    const r = origSetDirty.call(this, isDirty, ...rest);
    sync();
    return r;
  };

  // Powrót do zapisanej treści BEZ Cofnij (zgłoszenie 2026-10-17): wpisane i skasowane słowo,
  // Enter + Backspace, pogrubienie włączone i wyłączone — dawniej licznik pokazywał np.
  // „13 zmian”, bo liczy kroki cofania od zapisu, a „Zapisz” świeciło, choć plik byłby taki sam.
  // Teraz po pauzie sprawdzamy, czy dokument = zapisany plik; jeśli tak, stan jest czysty (jak po
  // cofnięciu wszystkiego ↶). Najpierw tanio: te same bajty i brak zmian w podglądzie; po
  // operacjach na pliku (nowe bajty) — porównanie zawartości części .docx, też w tle.
  async function sameZipContent(a, b) {
    if (a === b) return true;
    if (!window.JSZip) return false; // biblioteka dociągana z plikiem .docx — bez niej nie porównamy
    const [za, zb] = await Promise.all([window.JSZip.loadAsync(a), window.JSZip.loadAsync(b)]);
    const names = (z) => Object.keys(z.files).filter((n) => !z.files[n].dir).sort();
    const na = names(za);
    const nb = names(zb);
    if (na.length !== nb.length || na.some((n, i) => n !== nb[i])) return false;
    for (const n of na) {
      const [x, y] = await Promise.all([za.file(n).async("uint8array"), zb.file(n).async("uint8array")]);
      if (x.length !== y.length) return false;
      for (let i = 0; i < x.length; i++) if (x[i] !== y[i]) return false;
    }
    return true;
  }
  let cleanTimer = 0;
  let cleanRun = 0;
  function scheduleCleanCheck() {
    clearTimeout(cleanTimer);
    cleanTimer = setTimeout(checkBackToSaved, 600);
  }
  async function checkBackToSaved() {
    if (!hasUnsavedChanges || applying || busy || !cleanBytes || !originalFileBytes || readOnlyMode) return;
    const run = ++cleanRun;
    if (typeof waitInlineStructuralIdle === "function") await waitInlineStructuralIdle();
    const bytes = originalFileBytes;
    const quiet = () => run === cleanRun && !applying && !busy && hasUnsavedChanges && originalFileBytes === bytes
      && !(typeof pendingDocEdits !== "undefined" && pendingDocEdits.length)
      && !(typeof collectInlineParagraphEdits === "function" && collectInlineParagraphEdits().length);
    if (!quiet()) return;
    if (bytes !== cleanBytes) {
      let same = false;
      try { same = await sameZipContent(bytes, cleanBytes); } catch (_) { same = false; }
      if (!same || !quiet()) return;
    }
    window.setDirtyState(false); // = zapisany stan (licznik 0, szkic niepotrzebny)
  }
  docCanvasEl?.addEventListener("input", scheduleCleanCheck, true);
  docCanvasEl?.addEventListener("keyup", (e) => { if (["Backspace", "Delete", "Enter"].includes(e.key)) scheduleCleanCheck(); }, true);
  // Bez sprawdzania zaraz po operacjach z panelu (applyDocumentEdit): w teście chaosu (ziarno
  // 102) sprawdzanie tuż po przerysowaniu dawało niestabilne „Execution context was destroyed”.
  // B / I / U, rozmiar itd. — zmiana w podglądzie bez zdarzenia „input”
  document.addEventListener("selectionchange", () => { if (hasUnsavedChanges) scheduleCleanCheck(); });
  // Nowy dokument = nowa historia.
  const origIngest = window.ingestFile;
  window.ingestFile = async function ingestFileUndo(...args) {
    const ok = await origIngest.apply(this, args);
    if (ok) clear();
    return ok;
  };
  const origClear = window.clearDocumentState;
  window.clearDocumentState = function clearDocumentStateUndo(...args) {
    const r = origClear.apply(this, args);
    clear();
    return r;
  };

  undoBtn?.addEventListener("click", () => step("undo"));
  redoBtn?.addEventListener("click", () => step("redo"));

  // Skróty. W polach formularzy (szukanie, panel) zostaje natywne cofanie PISANIA.
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
    const key = String(e.key).toLowerCase();
    const isZ = e.code === "KeyZ" || key === "z";
    const isY = (e.code === "KeyY" || key === "y") && !e.metaKey;
    if (!isZ && !isY) return;
    if (!originalFileBytes) return;
    const a = document.activeElement;
    const tag = String(a?.tagName || "").toLowerCase();
    const inDoc = !!a?.closest?.(".docx-editable-p, .docx-edit-root");
    if (!inDoc && (a?.isContentEditable || tag === "textarea" || (tag === "input" && !["checkbox", "radio", "button", "range"].includes(a.type)) || tag === "select")) return;
    e.preventDefault();
    step(isY || e.shiftKey ? "redo" : "undo");
  }, true);

  sync();
  return { record, group, undo: () => step("undo"), redo: () => step("redo"), clear, changesSinceSave, sync, canUndo: () => undoStack.length > 0, canRedo: () => redoStack.length > 0, _debug: () => ({ undo: undoStack.map((e) => e.label), redo: redoStack.map((e) => e.label), savedPos }) };
})();
