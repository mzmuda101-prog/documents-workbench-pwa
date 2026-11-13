// Core runtime: DOM refs, shared state, and base UI helpers.

// Pamięć przeglądarki zablokowana (Safari „Blokuj wszystkie pliki cookie”, część trybów
// prywatnych): samo sięgnięcie po localStorage rzuca SecurityError, a odczyt ustawień
// (język, motyw…) na starcie modułów wywracał cały interfejs. Wtedy podstawiamy pamięć
// tylko na czas sesji (ustawienia nie przetrwają zamknięcia, ale aplikacja działa).
(() => {
  let ok = false;
  try { window.localStorage.getItem("__wb_probe__"); ok = true; } catch (_) {}
  if (!ok) {
    const mem = new Map();
    const shim = {
      getItem: (k) => (mem.has(String(k)) ? mem.get(String(k)) : null),
      setItem: (k, v) => { mem.set(String(k), String(v)); },
      removeItem: (k) => { mem.delete(String(k)); },
      clear: () => mem.clear(),
      key: (i) => Array.from(mem.keys())[i] ?? null,
      get length() { return mem.size; },
    };
    try { Object.defineProperty(window, "localStorage", { value: shim, configurable: true }); } catch (_) {}
  }
})();

const APP_BUILD_VERSION = "v1.00.35";

const IS_LOW_POWER = (() => {
  try {
    const cores = navigator.hardwareConcurrency || 8;
    const mem = navigator.deviceMemory || 8;
    const coarse = !!(window.matchMedia && window.matchMedia("(pointer: coarse)").matches);
    return cores <= 4 || mem <= 4 || (coarse && cores <= 6);
  } catch (_) {
    return false;
  }
})();
document.documentElement.classList.toggle("low-power", IS_LOW_POWER);

const rootEl = document.documentElement;
const appShellEl = document.querySelector(".app");
const logEl = document.getElementById("log");
const statusEl = document.getElementById("status");
const docViewportEl = document.getElementById("docViewport");
const docCanvasEl = document.getElementById("docCanvas");
const emptyStateEl = document.getElementById("emptyState");
const emptyTitleEl = document.getElementById("emptyTitle");
const emptySubEl = document.getElementById("emptySub");
const DEFAULT_EMPTY_TITLE = emptyTitleEl?.textContent || "";
const DEFAULT_EMPTY_SUB = emptySubEl?.textContent || "";
const docPanelEl = document.querySelector(".doc-panel");

const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const fileNameEl = document.getElementById("fileName");
const fileNameTextEl = document.getElementById("fileNameText");
const loadBtn = document.getElementById("loadBtn");
const loadSampleBtn = document.getElementById("loadSampleBtn");
const saveBtn = document.getElementById("saveBtn");
const saveAsBtn = document.getElementById("saveAsBtn");

const searchQueryEl = document.getElementById("searchQuery");
const searchScopeEl = document.getElementById("searchScope");
const searchHighlightBtn = document.getElementById("searchHighlightBtn");
const searchCountEl = document.getElementById("searchCount");

const structureSummaryEl = document.getElementById("structureSummary");

const editOpEl = document.getElementById("editOp");
const editFindEl = document.getElementById("editFind");
const editReplaceEl = document.getElementById("editReplace");
const editRegexEl = document.getElementById("editRegex");
const editScopeEl = document.getElementById("editScope");
const editReplaceFieldsEl = document.getElementById("editReplaceFields");
const editCaseFieldsEl = document.getElementById("editCaseFields");
const editCaseModeEl = document.getElementById("editCaseMode");
const editTrimFieldsEl = document.getElementById("editTrimFields");
const editTrimModeEl = document.getElementById("editTrimMode");
const editAffixFieldsEl = document.getElementById("editAffixFields");
const editPrefixEl = document.getElementById("editPrefix");
const editSuffixEl = document.getElementById("editSuffix");
const applyEditToolBtnEl = document.getElementById("applyEditToolBtn");

const panelToggle = document.getElementById("panelToggle");
const panelHandle = document.getElementById("panelHandle");
const sidebarScrim = document.getElementById("sidebarScrim");
const themeToggle = document.getElementById("themeToggle");
const brandRefresh = document.getElementById("brandRefresh");
const closeDocBtn = document.getElementById("closeDocBtn");
const appUpdateBtn = document.getElementById("appUpdateBtn");
const networkBadge = document.getElementById("networkBadge");
const loadingOverlayEl = document.getElementById("loadingOverlay");
const loadingTextEl = document.getElementById("loadingText");
const toastContainerEl = document.getElementById("toastContainer");
const zoomLevelEl = document.getElementById("zoomLevel");
const zoomValueEl = document.getElementById("zoomValue");
const readModeEl = document.getElementById("readMode");

const THEME_KEY = "documents-workbench-theme";
const LANG_KEY = "documents-workbench-lang";
const INTRO_PLAYED_KEY = "introPlayed";

let currentFileName = "";
let currentFileType = "";
let originalFileBytes = null;
let pendingDocEdits = [];
let fileHandle = null;
// Dokument, którego nie ma jeszcze na dysku (nowy, z PDF-a/zdjęcia, wynik porównania) — „Zapisz
// jako” proponuje samą nazwę, bez „_edited” (to tylko ochrona oryginału przy edycji pliku z dysku).
let currentFileFresh = false;
let hasUnsavedChanges = false;
let documentStructure = null;
let searchMatches = [];
let activeSearchIndex = -1;
let readOnlyMode = true;

function log(msg, type = "info") {
  if (!logEl) return;
  const line = document.createElement("div");
  line.className = `log-line log-${type}`;
  const locale = typeof currentLang !== "undefined" ? I18N[currentLang].locale : "pl-PL";
  line.textContent = `${new Date().toLocaleTimeString(locale)} ${msg}`;
  logEl.prepend(line);
}

// ── dymki: toast(message, type, opts) — wspólny kontrakt UX Sheet + Docs (2026-11-07) ──────────
// type: info | success | warning | error. info/success 2,8 s, warning/error 5 s (dłuższy tekst — dłużej).
// opts.action + opts.onAction → przycisk w dymku (np. „Cofnij”); taki dymek wisi ≥ 5 s.
// ✕ przy ostrzeżeniach, błędach i dymkach z akcją. Kolejne info/success tuż po sobie (np. trzy
// „Cofnięto…” pod rząd) PODMIENIAJĄ treść poprzedniego zamiast się piętrzyć — na telefonie stos
// dymków u góry zasłaniał pasek (audyt D2; wzór: Sheet app/core.js). Ostrzeżenia i błędy zostają
// osobno — każde niesie inną informację. Najechanie myszą wstrzymuje znikanie.
const TOAST_MERGE_TYPES = new Set(["success", "info"]);
const TOAST_MERGE_MS = 2000;
let _lastMergeToast = null; // { el, label, type, timer, at }

function toastDuration(type, msg, hasAction) {
  const base = type === "error" || type === "warning" ? 5000 : 2800;
  const extra = Math.max(0, String(msg || "").length - 70) * 35; // długi komunikat — czas na przeczytanie
  return Math.min(10000, Math.max(hasAction ? 5000 : 0, base + extra));
}

function scheduleToastOut(toastEl, ms) {
  toastEl._dwbMs = ms;
  return setTimeout(() => dismissToast(toastEl), ms);
}

function dismissToast(toastEl) {
  if (!toastEl || toastEl.classList.contains("out")) return;
  clearTimeout(toastEl._dwbTimer);
  toastEl.classList.add("out");
  if (_lastMergeToast?.el === toastEl) _lastMergeToast = null;
  setTimeout(() => toastEl.remove(), 200);
}

function toast(msg, type = "info", opts = {}) {
  const hasAction = !!(opts && opts.action && typeof opts.onAction === "function");
  const last = _lastMergeToast;
  if (!hasAction && TOAST_MERGE_TYPES.has(type) && last && last.type === type && last.el.isConnected
    && !last.el.classList.contains("out") && performance.now() - last.at < TOAST_MERGE_MS) {
    last.label.textContent = msg;
    last.at = performance.now();
    clearTimeout(last.el._dwbTimer);
    last.el._dwbTimer = scheduleToastOut(last.el, toastDuration(type, msg, false));
    return last.el;
  }
  const toastEl = document.createElement("div");
  toastEl.className = `toast ${type}`;
  // błąd — od razu (alert); reszta grzecznie w kolejce czytnika ekranu
  toastEl.setAttribute("role", type === "error" ? "alert" : "status");
  const icon = document.createElement("div");
  icon.className = "toast-icon";
  icon.setAttribute("aria-hidden", "true");
  icon.textContent = type === "success" ? "✓" : type === "error" ? "!" : type === "warning" ? "!" : "i";
  const label = document.createElement("div");
  label.className = "toast-msg";
  label.textContent = msg;
  toastEl.append(icon, label);
  if (hasAction) {
    const act = document.createElement("button");
    act.type = "button";
    act.className = "toast-action";
    act.textContent = opts.action;
    act.addEventListener("click", (e) => {
      e.stopPropagation();
      dismissToast(toastEl);
      opts.onAction();
    }, { once: true });
    toastEl.append(act);
    toastEl.classList.add("has-action");
  }
  if (hasAction || type === "warning" || type === "error") {
    const close = document.createElement("button");
    close.type = "button";
    close.className = "toast-close";
    close.textContent = "✕";
    close.setAttribute("aria-label", typeof t === "function" ? t("toastClose") : "Zamknij");
    close.addEventListener("click", (e) => { e.stopPropagation(); dismissToast(toastEl); });
    toastEl.append(close);
  }
  // mysz nad dymkiem = czytam / sięgam po „Cofnij” — nie znikaj spod kursora
  toastEl.addEventListener("pointerenter", (e) => { if (e.pointerType === "mouse") clearTimeout(toastEl._dwbTimer); });
  toastEl.addEventListener("pointerleave", (e) => {
    if (e.pointerType !== "mouse" || toastEl.classList.contains("out")) return;
    toastEl._dwbTimer = scheduleToastOut(toastEl, 1600);
  });
  toastContainerEl.appendChild(toastEl);
  toastEl._dwbTimer = scheduleToastOut(toastEl, opts.duration || toastDuration(type, msg, hasAction));
  _lastMergeToast = !hasAction && TOAST_MERGE_TYPES.has(type) ? { el: toastEl, label, type, at: performance.now() } : null;
  return toastEl;
}

// Operacja z jednym krokiem Cofnij (Zamień wszystkie, Akceptuj wszystkie, Wypełnij pola…): bez
// pytania przed — dymek z „Cofnij” po (kontrakt UX §3). „Cofnij” cofa tylko TĘ operację: gdy
// od tamtej pory doszedł inny krok, nie cofamy cudzego, tylko mówimy, gdzie jest Cofnij.
function toastUndoable(msg, type = "success") {
  const mark = typeof dwbUndo !== "undefined" && dwbUndo.top ? dwbUndo.top() : null;
  if (!mark) return toast(msg, type);
  return toast(msg, type, {
    action: t("undoAction"),
    onAction: () => {
      if (dwbUndo.top() === mark) dwbUndo.undo();
      else toast(t("undoStale"), "info");
    },
  });
}

// ── własne okna zamiast confirm() / prompt() (kontrakt UX §3) ─────────────────────────────────
// appDialog({ title, message, actions: [{ label, value, kind }], input }) → Promise<value | null>
//   kind: primary | danger | secondary; Esc / tło / ✕ = null (Anuluj).
//   Z polem (input: { value, placeholder, label }) — wynik to wpisany tekst (albo null).
// Fokus: TYLKO przy obsłudze klawiaturą trafia do pola albo na bezpieczną akcję (na iPhonie fokus
// w polu wysunąłby klawiaturę); Tab krąży w oknie; po zamknięciu fokus wraca, skąd przyszedł.
let lastInputModality = "pointer";
document.addEventListener("keydown", (e) => { if (!["Shift", "Control", "Alt", "Meta"].includes(e.key)) lastInputModality = "keyboard"; }, true);
document.addEventListener("pointerdown", () => { lastInputModality = "pointer"; }, true);

const appDialogQueue = [];
let appDialogBusy = false;

function appDialog(opts = {}) {
  // Testy (scripts/*-playwright.js) dawniej zatwierdzały systemowe confirm()/prompt() przez
  // page.on("dialog") — window.__dwbAutoDialog daje to samo dla własnych okien: „accept” = pierwsza
  // główna akcja (z wartością pola), „dismiss” = Anuluj. Użytkownik nigdy tego nie ustawia.
  const auto = typeof window !== "undefined" && window.__dwbAutoDialog;
  // test może zbierać treść pytań po stronie Node (page.exposeFunction) — jak dawniej page.on("dialog")
  if (auto) { try { window.__dwbDialogSeen?.(`${opts.title ? `${opts.title} ` : ""}${opts.message || ""}`.trim()); } catch (_) { /* test bez zbierania */ } }
  // funkcja = odpowiedź testu na konkretne okno (jak dawne podmienione window.confirm(msg)):
  // true → główna akcja, false/null → Anuluj, inna wartość → ta wartość
  if (typeof auto === "function") {
    const ans = auto({ title: opts.title || "", message: opts.message || "", actions: opts.actions || null, input: !!opts.input });
    if (ans === true) {
      const acts = opts.actions || [{ value: true, kind: "primary" }];
      const main = acts.find((x) => x.value === "discard") || acts.find((x) => x.kind === "primary" || x.kind === "danger") || acts[0];
      return Promise.resolve(opts.input ? (opts.input.value ?? "") : main.value);
    }
    return Promise.resolve(ans === false ? null : ans);
  }
  if (auto === "accept" || auto === "dismiss") {
    if (auto === "dismiss") return Promise.resolve(null);
    const acts = opts.actions || [{ value: true, kind: "primary" }];
    // niezapisana praca: dawne systemowe „OK” znaczyło „kontynuuj bez zapisywania” = Nie zapisuj
    const main = acts.find((x) => x.value === "discard") || acts.find((x) => x.kind === "primary") || acts[0];
    return Promise.resolve(opts.input ? (opts.input.value ?? "") : main.value);
  }
  return new Promise((resolve) => {
    appDialogQueue.push({ opts, resolve });
    if (!appDialogBusy) runNextAppDialog();
  });
}

function runNextAppDialog() {
  const job = appDialogQueue.shift();
  if (!job) { appDialogBusy = false; return; }
  appDialogBusy = true;
  const { opts, resolve } = job;
  const returnTo = document.activeElement;
  const keyboard = lastInputModality === "keyboard";
  const dlg = document.createElement("dialog");
  dlg.className = "app-dialog";
  dlg.setAttribute("aria-modal", "true");
  const titleId = `appDlgT${Date.now().toString(36)}`;
  const box = document.createElement("form");
  box.method = "dialog";
  box.className = "app-dialog-inner";
  if (opts.title) {
    const h = document.createElement("h2");
    h.id = titleId;
    h.className = "app-dialog-title";
    h.textContent = opts.title;
    box.append(h);
    dlg.setAttribute("aria-labelledby", titleId);
  }
  if (opts.message) {
    const p = document.createElement("p");
    p.className = "app-dialog-msg";
    p.textContent = opts.message;
    box.append(p);
  }
  let inputEl = null;
  if (opts.input) {
    const lab = document.createElement("label");
    lab.className = "field app-dialog-field";
    if (opts.input.label) {
      const s = document.createElement("span");
      s.textContent = opts.input.label;
      lab.append(s);
    }
    inputEl = document.createElement("input");
    inputEl.type = "text";
    inputEl.autocomplete = "off";
    inputEl.enterKeyHint = "done";
    inputEl.value = opts.input.value || "";
    if (opts.input.placeholder) inputEl.placeholder = opts.input.placeholder;
    if (!opts.input.label && opts.title) inputEl.setAttribute("aria-labelledby", titleId);
    lab.append(inputEl);
    box.append(lab);
  }
  const actions = Array.isArray(opts.actions) && opts.actions.length ? opts.actions
    : [{ label: t("dlgOk"), value: true, kind: "primary" }, { label: t("dlgCancel"), value: null, kind: "secondary" }];
  const row = document.createElement("div");
  row.className = "app-dialog-actions";
  if (actions.length > 2) row.classList.add("is-three");
  let done = false;
  const finish = (value) => {
    if (done) return;
    done = true;
    document.removeEventListener("keydown", trapTab, true);
    try { dlg.close(); } catch (_) {}
    dlg.remove();
    // fokus wraca tylko przy klawiaturze (na dotyku powrót do akapitu wysunąłby klawiaturę bez potrzeby)
    if (lastInputModality === "keyboard" && returnTo?.isConnected && typeof returnTo.focus === "function") {
      try { returnTo.focus({ preventScroll: true }); } catch (_) {}
    }
    resolve(value);
    setTimeout(runNextAppDialog, 0);
  };
  const btns = actions.map((a) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = `btn app-dialog-btn${a.kind === "primary" ? " primary" : a.kind === "danger" ? " danger" : ""}`;
    b.textContent = a.label;
    b.addEventListener("click", () => {
      if (a.value === null || a.value === undefined) { finish(null); return; }
      finish(inputEl ? inputEl.value : a.value);
    });
    row.append(b);
    return { a, b };
  });
  box.append(row);
  // Enter w polu = główna akcja
  box.addEventListener("submit", (e) => {
    e.preventDefault();
    const main = btns.find((x) => x.a.kind === "primary") || btns[0];
    main?.b.click();
  });
  dlg.append(box);
  dlg.addEventListener("cancel", (e) => { e.preventDefault(); finish(null); }); // Esc
  // Esc także wtedy, gdy obsługa Esc dokumentu (keyboard.js) zablokuje domyślne „cancel” okna
  dlg.addEventListener("keydown", (e) => { if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); finish(null); } });
  dlg.addEventListener("click", (e) => { if (e.target === dlg) finish(null); }); // stuknięcie w tło
  // bezpieczna akcja: przy groźnej (danger) — Anuluj; inaczej główna
  const hasDanger = actions.some((a) => a.kind === "danger");
  const safe = (hasDanger ? btns.find((x) => x.a.value == null) : btns.find((x) => x.a.kind === "primary")) || btns[0];
  const initial = inputEl || safe?.b;
  if (keyboard && initial) initial.setAttribute("autofocus", "");
  else { dlg.tabIndex = -1; dlg.setAttribute("autofocus", ""); } // fokus na samym oknie — bez klawiatury ekranowej
  const trapTab = (e) => {
    if (e.key !== "Tab" || !dlg.open) return;
    const list = [...dlg.querySelectorAll("button, input")].filter((x) => !x.disabled);
    if (!list.length) return;
    const first = list[0];
    const lastEl = list[list.length - 1];
    if (!dlg.contains(document.activeElement) || document.activeElement === dlg) { e.preventDefault(); (e.shiftKey ? lastEl : first).focus(); }
    else if (e.shiftKey && document.activeElement === first) { e.preventDefault(); lastEl.focus(); }
    else if (!e.shiftKey && document.activeElement === lastEl) { e.preventDefault(); first.focus(); }
  };
  document.addEventListener("keydown", trapTab, true);
  document.body.append(dlg);
  if (typeof dlg.showModal === "function") dlg.showModal();
  else dlg.setAttribute("open", "");
  if (keyboard && inputEl) inputEl.select();
}

function appConfirm(message, { title = "", ok, cancel, danger = false } = {}) {
  return appDialog({
    title, message,
    actions: [
      { label: ok || t("dlgOk"), value: true, kind: danger ? "danger" : "primary" },
      { label: cancel || t("dlgCancel"), value: null, kind: "secondary" },
    ],
  }).then((v) => v === true);
}

function appPrompt(message, value = "", { title = "", placeholder = "", ok } = {}) {
  return appDialog({
    title: title || message,
    message: title ? message : "",
    input: { value, placeholder },
    actions: [
      { label: ok || t("dlgOk"), value: true, kind: "primary" },
      { label: t("dlgCancel"), value: null, kind: "secondary" },
    ],
  });
}

// Niezapisana praca — zawsze Zapisz / Nie zapisuj / Anuluj (jak Word). → "save" | "discard" | null
function askSaveChanges(name) {
  return appDialog({
    title: t("dlgSaveTitle", { name: name || currentFileName || t("dlgThisDoc") }),
    message: t("dlgSaveMsg"),
    actions: [
      { label: t("save"), value: "save", kind: "primary" },
      { label: t("dlgDontSave"), value: "discard", kind: "secondary" },
      { label: t("dlgCancel"), value: null, kind: "secondary" },
    ],
  });
}

function setLoading(isLoading, text) {
  loadingTextEl.textContent = text || t("loadingGeneric");
  loadingOverlayEl.classList.toggle("hidden", !isLoading);
}

function setStatus(msg) {
  statusEl.textContent = msg;
}

function setDirtyState(isDirty) {
  hasUnsavedChanges = !!isDirty;
  statusEl.classList.toggle("unsaved", hasUnsavedChanges);
  document.title = hasUnsavedChanges ? `* ${BASE_TITLE}` : BASE_TITLE;
  if (typeof syncActionButtons === "function") syncActionButtons();
}

function showEmptyState(title, sub) {
  emptyStateEl.classList.remove("hidden");
  emptyTitleEl.textContent = title || DEFAULT_EMPTY_TITLE;
  emptySubEl.textContent = sub || DEFAULT_EMPTY_SUB;
  docCanvasEl.classList.add("hidden");
}

function hideEmptyState() {
  emptyStateEl.classList.add("hidden");
  docCanvasEl.classList.remove("hidden");
}

// ── przejście „zdjęcie starego → nowy widok” (View Transitions API) ───────────
// Przy ciężkiej przebudowie (np. Widok mobilny ⇄ desktopowy) przeglądarka trzyma na ekranie
// zdjęcie starego widoku, w tym czasie raz układa nowy, a potem tylko przenika obrazy
// (opacity na kompozytorze — bez przeliczania strony w każdej klatce). Zamiast mignięcia
// nakładki „Renderowanie…” i skoku jest płynne przejście.
// update może zwrócić Promise; czekamy najwyżej maxWait ms — przy bardzo dużym pliku
// lepiej pokazać nakładkę z postępem niż zamrożony obraz bez znaku życia.
let activeViewTransition = null;
function withViewTransition(update, { maxWait = 700 } = {}) {
  if (typeof document.startViewTransition !== "function" || activeViewTransition
    || matchMedia("(prefers-reduced-motion: reduce)").matches || document.visibilityState !== "visible") {
    return Promise.resolve(update());
  }
  let job;
  const vt = document.startViewTransition(() => {
    job = Promise.resolve(update());
    return Promise.race([job.catch(() => {}), new Promise((r) => setTimeout(r, maxWait))]);
  });
  activeViewTransition = vt;
  rootEl.classList.add("dwb-vt");
  // skipTransition() odrzuca ready/finished — bez catch to „unhandled rejection”
  vt.ready.catch(() => {});
  const done = () => {
    if (activeViewTransition !== vt) return;
    activeViewTransition = null;
    rootEl.classList.remove("dwb-vt");
  };
  vt.finished.catch(() => {}).finally(done);
  // Bezpiecznik (lekcja z Sheet): iOS potrafi uśpić PWA w trakcie przejścia — zamrożona
  // nakładka ::view-transition łapała wtedy dotyk aż do ubicia aplikacji.
  setTimeout(() => { if (activeViewTransition === vt) { try { vt.skipTransition(); } catch (_) {} done(); } }, maxWait + 1200);
  return (vt.updateCallbackDone || vt.finished).catch(() => {}).then(() => job);
}
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden" && activeViewTransition) {
    try { activeViewTransition.skipTransition(); } catch (_) {}
  }
});

function isSidebarOpen() {
  return rootEl.classList.contains("sidebar-open");
}

function setSidebarOpen(open) {
  rootEl.classList.toggle("sidebar-open", !!open);
  if (sidebarScrim) sidebarScrim.classList.toggle("hidden", !open);
  const mobile = window.matchMedia("(max-width: 768px)").matches;
  document.body.style.overflow = open && mobile ? "hidden" : "";
  if (typeof syncSidebarHandle === "function") {
    syncSidebarHandle();
    requestAnimationFrame(() => syncSidebarHandle());
    window.setTimeout(() => syncSidebarHandle(), 270);
  }
}

function toggleSidebar() {
  setSidebarOpen(!isSidebarOpen());
}

function applyZoom() {
  if (!docCanvasEl || !zoomLevelEl) return;
  const { min, max } = typeof getZoomLimits === "function" ? getZoomLimits() : { min: 0.5, max: 2 };
  let zoom = parseFloat(zoomLevelEl.value) || 1;
  zoom = Math.max(min, Math.min(max, zoom));
  zoomLevelEl.value = String(zoom);
  docCanvasEl.style.setProperty("--doc-zoom", String(zoom));
  // Widok mobilny: wcięcia akapitów rosną jak √zoom, nie jak litery (softenReflowIndents)
  docCanvasEl.style.setProperty("--dwb-ik", zoom > 1 ? String(Math.round((1 / Math.sqrt(zoom)) * 1000) / 1000) : "1");
  if (zoomValueEl) zoomValueEl.textContent = `${Math.round(zoom * 100)}%`;
  if (typeof updateZoomShellHeight === "function") updateZoomShellHeight(zoom);
  docCanvasEl.dispatchEvent(new CustomEvent("dwb-zoom")); // np. karta komentarza rośnie razem z tekstem
}

function syncDocumentShellClass() {
  document.body.classList.toggle("has-document", !!originalFileBytes);
  if (closeDocBtn) closeDocBtn.classList.toggle("hidden", !originalFileBytes);
  const closePanelBtn = document.getElementById("closeDocPanelBtn");
  if (closePanelBtn) closePanelBtn.classList.toggle("hidden", !originalFileBytes);
}

function syncDocViewportHeight() {
  if (!docPanelEl) return;
  requestAnimationFrame(() => {
    const rect = docPanelEl.getBoundingClientRect();
    const kbOpen = document.body.classList.contains("kb-open");
    // Klawiatura ekranowa (touch.js): liczymy do WIDOCZNEGO obszaru, nie do okna — inaczej
    // dół dokumentu leży pod klawiaturą, a iOS przesuwa całą stronę (i pasek ucieka).
    const viewportHeight = (kbOpen && window.visualViewport ? window.visualViewport.height : window.innerHeight)
      || document.documentElement.clientHeight || 720;
    const mobile = window.matchMedia("(max-width: 768px)").matches;
    // Pod dokumentem stoi pasek stanu — obszar dokumentu kończy się nad nim, żeby
    // całość mieściła się w oknie bez przewijania strony.
    const statusBar = document.querySelector(".status-bar");
    const statusH = statusBar && !kbOpen ? statusBar.getBoundingClientRect().height + (mobile ? 0 : 12) : 0;
    const bottomGap = mobile ? 0 : 16;
    const available = Math.floor(viewportHeight - rect.top - statusH - bottomGap);
    const minHeight = kbOpen ? 120 : mobile ? 280 : 420;
    const height = Math.max(minHeight, available);
    // zapis tej samej wartości też unieważnia układ panelu z dokumentem — a zaraz niżej
    // scrollHeight by go wymusił (drugi pełny układ przy każdym Enterze/zmianie paska)
    const setPanelHeight = (v) => { if (docPanelEl.style.getPropertyValue("--doc-panel-height") !== v) docPanelEl.style.setProperty("--doc-panel-height", v); };
    setPanelHeight(`${height}px`);
    if (kbOpen) return;
    // Korekta po pomiarze: wyliczenie wyżej nie zna odstępów między panelem a paskiem stanu,
    // marginesu strony ani miejsca na pasek Home (safe-area) — na telefonie strona wystawała
    // o kilkadziesiąt px i dało się ją przewinąć (pasek przewijania, ucięty pasek stanu).
    // Mierzymy, o ile strona jest wyższa od okna, i tyle odejmujemy.
    const over = document.documentElement.scrollHeight - window.innerHeight;
    if (over > 0 && over < 400 && height - over >= minHeight) {
      setPanelHeight(`${height - over}px`);
      if (window.scrollY) window.scrollTo(0, 0);
    }
  });
}

function formatFileSize(bytes) {
  if (!bytes || bytes < 1024) return `${bytes || 0} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function detectFileType(name, mime) {
  const lower = (name || "").toLowerCase();
  if (lower.endsWith(".docx") || mime === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") return "docx";
  if (lower.endsWith(".pdf") || mime === "application/pdf") return "pdf";
  // zdjęcie dokumentu (app/photo-import.js → PDF → OCR); HEIC odczyta tylko Safari
  if (/\.(jpe?g|png|webp|heic|heif)$/.test(lower) || /^image\/(jpeg|png|webp|heic|heif)$/.test(mime || "")) return "image";
  return "";
}

function clearDocumentState() {
  currentFileName = "";
  currentFileFresh = false;
  currentFileType = "";
  originalFileBytes = null;
  pendingDocEdits = [];
  if (typeof dwbClipComments !== "undefined") dwbClipComments.clear();
  if (typeof dwbClipNotes !== "undefined") dwbClipNotes.clear();
  fileHandle = null;
  documentStructure = null;
  searchMatches = [];
  activeSearchIndex = -1;
  if (docCanvasEl) docCanvasEl.replaceChildren();
  if (structureSummaryEl) structureSummaryEl.replaceChildren();
  if (typeof resetStructurePanelUi === "function") resetStructurePanelUi();
  if (typeof loadMetadataFromDocument === "function") loadMetadataFromDocument().catch(() => {});
  if (searchCountEl) searchCountEl.textContent = "";
  if (fileNameEl) fileNameEl.classList.add("hidden");
  if (fileNameTextEl) fileNameTextEl.textContent = t("noFile");
  baselineParagraphTexts = [];
  setDirtyState(false);
  showEmptyState();
  syncDocumentShellClass();
  if (typeof syncActionButtons === "function") syncActionButtons();
}
