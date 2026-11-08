// doc-comfort.js — wygody wokół dokumentu (audyt QoL Docs 2026-11-06):
//   D1  Czytanie: pierwsza próba pisania (albo dwuklik w akapit) → podpowiedź „Edytuj (Ctrl/⌘+Alt+E)”
//       przy tekście; przycisk przełącza tryb i stawia kursor tam, gdzie kliknięto (jak „Edytuj
//       dokument” w Wordzie na telefonie). Dawniej pisanie w Czytaniu nie robiło nic, bez sygnału.
//   D24 tryb Czytanie/Edycja pamiętany osobno dla dotyku i komputera.
//   D4  pasek stanu przy zaznaczeniu: „57 z 3415 słów” (jak Word).
//   D16 „Strona 3 z 12” w pasku stanu (z granic stron page-breaks.js) + Przejdź do strony (klik, Ctrl/⌘+Alt+G).
//   D10 ponowne otwarcie tego samego pliku: dymek „Wróć do miejsca, w którym skończyłeś” + Przejdź.
//   D17 „Ostatnio otwierane” na ekranie startowym (kopie ostatnich plików w IndexedDB, tylko na tym urządzeniu).
//   D23 rozwinięte sekcje panelu pamiętane między uruchomieniami.

const dwbComfort = (() => {
  const statusWordsEl = document.getElementById("statusWords");
  const statusPageEl = document.getElementById("statusPage");
  const coarse = () => !!window.matchMedia?.("(pointer: coarse)").matches;
  const lsGet = (k) => { try { return localStorage.getItem(k); } catch (_) { return null; } };
  const lsSet = (k, v) => { try { localStorage.setItem(k, v); } catch (_) {} };

  // ── D1: podpowiedź w Czytaniu ──────────────────────────────────────────────
  let hintEl = null;
  let hintTimer = 0;
  let lastPoint = null; // { x, y } ostatniego stuknięcia/kliknięcia w tekst
  let dblHintShown = false; // dwuklik = zwykle zaznaczanie słowa; podpowiedź tylko raz na sesję

  function hideReadHint() {
    clearTimeout(hintTimer);
    hintEl?.remove();
    hintEl = null;
  }
  function caretPointFallback() {
    const s = window.getSelection?.();
    if (s?.rangeCount && docCanvasEl?.contains(s.anchorNode)) {
      const r = s.getRangeAt(0).getClientRects()[0] || s.getRangeAt(0).getBoundingClientRect();
      if (r && (r.width || r.height)) return { x: r.left, y: r.top + r.height / 2 };
    }
    return null;
  }
  function showReadHint(point) {
    if (!originalFileBytes || !readOnlyMode) return;
    hideReadHint();
    const at = point || lastPoint || caretPointFallback();
    hintEl = document.createElement("div");
    hintEl.className = "read-hint";
    hintEl.setAttribute("role", "status");
    const txt = document.createElement("span");
    txt.className = "read-hint-text";
    txt.textContent = t("readHintText");
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "read-hint-btn";
    btn.textContent = t("readHintEdit");
    const kbd = document.createElement("span");
    kbd.className = "read-hint-kbd";
    kbd.textContent = typeof dwbShortcuts !== "undefined" ? dwbShortcuts.label("mode") : "Ctrl/⌘+Alt+E";
    btn.append(kbd);
    btn.addEventListener("pointerdown", (e) => e.preventDefault()); // nie zabieraj zaznaczenia
    btn.addEventListener("click", () => { const p = at; hideReadHint(); enterEditAt(p); });
    const close = document.createElement("button");
    close.type = "button";
    close.className = "read-hint-close";
    close.textContent = "✕";
    close.setAttribute("aria-label", t("toastClose"));
    close.addEventListener("click", hideReadHint);
    hintEl.append(txt, btn, close);
    document.body.append(hintEl);
    // przy tekście (nad punktem), a bez punktu — u dołu obszaru dokumentu
    const vp = docViewportEl.getBoundingClientRect();
    const w = hintEl.offsetWidth;
    const h = hintEl.offsetHeight;
    const vw = document.documentElement.clientWidth;
    let x = at ? at.x - w / 2 : vp.left + vp.width / 2 - w / 2;
    let y = at ? at.y - h - 14 : vp.bottom - h - 16;
    if (at && y < vp.top + 6) y = at.y + 18;
    x = Math.max(8, Math.min(vw - w - 8, x));
    y = Math.max(8, Math.min(window.innerHeight - h - 8, y));
    hintEl.style.left = `${Math.round(x)}px`;
    hintEl.style.top = `${Math.round(y)}px`;
    if (lastInputModality === "keyboard") btn.focus({ preventScroll: true });
    hintTimer = setTimeout(hideReadHint, 7000);
  }

  // Edycja + kursor w miejscu kliknięcia (albo w zaznaczeniu / pierwszym widocznym akapicie)
  function enterEditAt(point) {
    const s = window.getSelection?.();
    const keep = s?.rangeCount && docCanvasEl?.contains(s.anchorNode) ? s.getRangeAt(0).cloneRange() : null;
    appFrame.setReadOnly(false);
    const place = () => {
      let range = null;
      if (point) {
        if (document.caretRangeFromPoint) range = document.caretRangeFromPoint(point.x, point.y);
        else if (document.caretPositionFromPoint) {
          const pos = document.caretPositionFromPoint(point.x, point.y);
          if (pos) { range = document.createRange(); range.setStart(pos.offsetNode, pos.offset); range.collapse(true); }
        }
      }
      if (!range || !docCanvasEl.contains(range.startContainer)) range = keep;
      const host = range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest?.(".docx-editable-p, .docx-edit-root");
      if (!host) { window.dwbKeyboard?.focusDocument(); return; }
      try { host.focus({ preventScroll: true }); } catch (_) {}
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    };
    // przełączenie trybu przebudowuje akapity do edycji (synchronicznie albo w następnej klatce)
    requestAnimationFrame(() => requestAnimationFrame(place));
  }

  const PRINTABLE = (e) => e.key && e.key.length === 1 && !e.ctrlKey && !e.metaKey;
  function inField(el) {
    const tag = String(el?.tagName || "").toLowerCase();
    return tag === "input" || tag === "textarea" || tag === "select" || (el?.isContentEditable && !docCanvasEl?.contains(el));
  }
  document.addEventListener("keydown", (e) => {
    if (!readOnlyMode || !originalFileBytes || e.isComposing || e.defaultPrevented) return;
    if (e.key === "Escape") { hideReadHint(); return; }
    if (!PRINTABLE(e) && e.key !== "Enter" && e.key !== "Backspace" && e.key !== "Delete") return;
    if (e.altKey && !PRINTABLE(e)) return;
    const a = document.activeElement;
    if (inField(a) || document.querySelector("dialog[open]")) return;
    if (!(docViewportEl?.contains(a) || a === document.body || a === document.documentElement)) return;
    if (hintEl) return;
    // spacja w Czytaniu to przewijanie albo szybki podgląd linku (jak na Macu) — nie pisanie;
    // podpowiedź pokaże pierwsza litera
    if (e.key === " ") return;
    showReadHint(null);
  });
  docCanvasEl?.addEventListener("pointerdown", (e) => { lastPoint = { x: e.clientX, y: e.clientY }; }, true);
  docCanvasEl?.addEventListener("dblclick", (e) => {
    if (!readOnlyMode || dblHintShown || !e.target.closest?.("p, li, td")) return;
    dblHintShown = true;
    showReadHint({ x: e.clientX, y: e.clientY });
  });
  docViewportEl?.addEventListener("scroll", () => { if (hintEl) hideReadHint(); }, { passive: true });
  readModeEl?.addEventListener("change", hideReadHint);

  // ── D24: tryb pamiętany osobno dla dotyku i komputera ──────────────────────
  const MODE_KEY = "dwb-read-mode-v1";
  function modeStore() { try { return JSON.parse(lsGet(MODE_KEY) || "{}") || {}; } catch (_) { return {}; } }
  readModeEl?.addEventListener("change", () => {
    const m = modeStore();
    m[coarse() ? "touch" : "desktop"] = readModeEl.checked ? "read" : "edit";
    lsSet(MODE_KEY, JSON.stringify(m));
  });
  function applySavedMode() {
    const want = modeStore()[coarse() ? "touch" : "desktop"];
    if (want === "edit" && readOnlyMode) appFrame.setReadOnly(false);
  }

  // ── D4: słowa zaznaczenia ─────────────────────────────────────────────────
  let selQueued = false;
  function syncSelectionWords() {
    selQueued = false;
    if (!statusWordsEl || !originalFileBytes || !documentStructure) return;
    const sel = window.getSelection?.();
    const total = documentStructure.words.toLocaleString(I18N[currentLang].locale);
    const inDoc = sel && !sel.isCollapsed && sel.rangeCount && docCanvasEl?.contains(sel.anchorNode);
    const n = inDoc && typeof wordTextStats === "function" ? wordTextStats(sel.toString().split("\n")).words : 0;
    statusWordsEl.textContent = n ? t("statusWordsSel", { n: n.toLocaleString(I18N[currentLang].locale), total }) : t("metaWords", { n: total });
  }
  document.addEventListener("selectionchange", () => {
    if (selQueued) return;
    selQueued = true;
    requestAnimationFrame(syncSelectionWords);
  });

  // ── D16: strona X z Y ─────────────────────────────────────────────────────
  // Strony = kartki (section.docx, jawne podziały) + granice z page-breaks.js (pas przerwy albo kreska).
  function pageStarts() {
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    if (!host || docCanvasEl.classList.contains("doc-reflow-mode")) return [];
    const out = [];
    host.querySelectorAll("section.docx").forEach((sec) => {
      const r = sec.getBoundingClientRect();
      if (!r.height) return;
      out.push({ y: r.top, el: sec });
      sec.querySelectorAll(":scope > .dwb-page-gap-band, :scope > .dwb-page-break").forEach((m) => {
        const mr = m.getBoundingClientRect();
        out.push({ y: m.classList.contains("dwb-page-gap-band") ? mr.bottom : mr.top, el: m });
      });
    });
    return out.sort((a, b) => a.y - b.y);
  }
  let pageQueued = false;
  function syncPage() {
    pageQueued = false;
    if (!statusPageEl) return;
    const starts = originalFileBytes ? pageStarts() : [];
    if (starts.length < 1) { statusPageEl.hidden = true; return; }
    const ref = docViewportEl.getBoundingClientRect().top + 48;
    let cur = 1;
    starts.forEach((s, i) => { if (s.y <= ref) cur = i + 1; });
    statusPageEl.hidden = false;
    statusPageEl.textContent = t("statusPage", { n: cur, total: starts.length });
    statusPageEl.dataset.total = String(starts.length);
  }
  function schedulePage() {
    if (pageQueued) return;
    pageQueued = true;
    requestAnimationFrame(syncPage);
  }
  async function goToPage() {
    if (!originalFileBytes) return;
    const starts = pageStarts();
    if (!starts.length) { toast(t("goToPageNoPages"), "info"); return; }
    const v = await appPrompt(t("goToPageAsk", { total: starts.length }), "", { title: t("goToPageTitle"), ok: t("goToPageOk") });
    if (v == null) return;
    const n = parseInt(String(v).trim(), 10);
    if (!Number.isFinite(n) || n < 1 || n > starts.length) { toast(t("goToPageBad", { total: starts.length }), "warning"); return; }
    const target = pageStarts()[n - 1];
    docViewportEl.scrollTop += target.y - docViewportEl.getBoundingClientRect().top - 8;
    schedulePage();
  }
  docViewportEl?.addEventListener("scroll", schedulePage, { passive: true });
  window.addEventListener("resize", schedulePage);
  statusPageEl?.addEventListener("click", goToPage);
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || !e.altKey || e.shiftKey || e.code !== "KeyG" || !originalFileBytes) return;
    e.preventDefault();
    goToPage();
  });
  // granice stron rysują się/zmieniają po przebudowie — patrzymy tylko na znaczniki stron
  if (docCanvasEl && typeof MutationObserver === "function") {
    const isMark = (n) => n.nodeType === 1 && (n.classList.contains("dwb-page-gap-band") || n.classList.contains("dwb-page-break") || n.matches("section.docx, .docx-preview-host, .docx-wrapper"));
    new MutationObserver((recs) => {
      for (const r of recs) {
        if ([...r.addedNodes].some(isMark) || [...r.removedNodes].some(isMark)) { schedulePage(); return; }
      }
    }).observe(docCanvasEl, { childList: true, subtree: true });
  }

  // ── D10: wróć do miejsca, w którym skończyłeś ─────────────────────────────
  const RESUME_KEY = "dwb-resume-v1";
  const RESUME_MAX = 30;
  let resumeId = null;
  function resumeList() { try { return JSON.parse(lsGet(RESUME_KEY) || "[]") || []; } catch (_) { return []; } }
  let resumeTimer = 0;
  function rememberPlace() {
    clearTimeout(resumeTimer);
    resumeTimer = 0;
    if (!resumeId || !originalFileBytes) return;
    const a = typeof captureDocScrollAnchor === "function" ? captureDocScrollAnchor() : null;
    if (!a) return;
    const list = resumeList().filter((x) => x.id !== resumeId);
    list.unshift({ id: resumeId, index: a.index, frac: Math.round(a.frac * 1000) / 1000, at: Date.now() });
    lsSet(RESUME_KEY, JSON.stringify(list.slice(0, RESUME_MAX)));
  }
  docViewportEl?.addEventListener("scroll", () => {
    if (!resumeId) return;
    clearTimeout(resumeTimer);
    resumeTimer = setTimeout(rememberPlace, 800);
  }, { passive: true });
  window.addEventListener("pagehide", rememberPlace);
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") rememberPlace(); });
  function offerResume(file) {
    resumeId = `${file.name}|${file.size}`;
    const saved = resumeList().find((x) => x.id === resumeId);
    if (!saved || saved.index < 3) return;
    toast(t("resumeAsk"), "info", {
      action: t("resumeGo"),
      duration: 8000,
      onAction: () => {
        if (typeof restoreDocScrollAnchor !== "function") return;
        const anchor = { index: saved.index, frac: saved.frac || 0, gap: 0 };
        restoreDocScrollAnchor(anchor);
        // tuż po otwarciu układ jeszcze dojeżdża (granice stron, kroje, obrazy, płynne przewinięcie
        // po wczytaniu) — dwie poprawki, o ile użytkownik sam w tym czasie nie przewijał
        let userMoved = false;
        const mark = () => { userMoved = true; };
        const evs = ["wheel", "touchstart", "keydown", "pointerdown"];
        evs.forEach((ev) => docViewportEl.addEventListener(ev, mark, { passive: true, once: true }));
        const fix = () => {
          if (userMoved) return;
          const cur = captureDocScrollAnchor();
          if (!cur || Math.abs(cur.index - anchor.index) > 1) restoreDocScrollAnchor(anchor);
        };
        setTimeout(fix, 450);
        setTimeout(() => { fix(); evs.forEach((ev) => docViewportEl.removeEventListener(ev, mark)); }, 1200);
      },
    });
  }

  // ── D17: ostatnio otwierane ───────────────────────────────────────────────
  // Kopia pliku (do 15 MB) + uchwyt (Chrome/Edge — wtedy otwieramy świeżą wersję z dysku).
  const RECENT_DB = "dwb-recent";
  const RECENT_MAX = 6;
  const RECENT_MAX_BYTES = 15 * 1024 * 1024;
  let recentDb = null;
  function rdb() {
    if (!recentDb) {
      recentDb = new Promise((resolve, reject) => {
        if (typeof indexedDB === "undefined") { reject(new Error("no indexedDB")); return; }
        const req = indexedDB.open(RECENT_DB, 1);
        req.onupgradeneeded = () => req.result.createObjectStore("files", { keyPath: "id" });
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      }).catch((e) => { recentDb = null; throw e; });
    }
    return recentDb;
  }
  async function rtx(mode, fn) {
    const d = await rdb();
    return new Promise((resolve, reject) => {
      const tr = d.transaction("files", mode);
      let out;
      const r = fn(tr.objectStore("files"));
      if (r && "onsuccess" in r) r.onsuccess = () => { out = r.result; };
      tr.oncomplete = () => resolve(out);
      tr.onerror = () => reject(tr.error);
      tr.onabort = () => reject(tr.error);
    });
  }
  const recentAll = () => rtx("readonly", (s) => s.getAll()).then((l) => (l || []).sort((a, b) => b.openedAt - a.openedAt)).catch(() => []);
  async function rememberRecent(file, handle) {
    if (!file || file.size > RECENT_MAX_BYTES) return;
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const id = `${file.name}|${file.size}`;
      const rec = { id, name: file.name, size: file.size, type: file.type || "", openedAt: Date.now(), bytes, handle: handle || null };
      await rtx("readwrite", (s) => { try { s.put(rec); } catch (_) { s.put({ ...rec, handle: null }); } });
      const all = await recentAll();
      const extra = all.slice(RECENT_MAX);
      if (extra.length) await rtx("readwrite", (s) => { extra.forEach((x) => s.delete(x.id)); });
    } catch (_) { /* tryb prywatny / brak miejsca — bez listy */ }
  }
  async function openRecent(rec) {
    let file = null;
    let handle = null;
    if (rec.handle && typeof rec.handle.requestPermission === "function") {
      try {
        if ((await rec.handle.requestPermission({ mode: "readwrite" })) === "granted") { file = await rec.handle.getFile(); handle = rec.handle; }
      } catch (_) { file = null; }
    }
    if (!file) file = new File([rec.bytes], rec.name, { type: rec.type || (/\.pdf$/i.test(rec.name) ? "application/pdf" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document") });
    const item = handle && /\.docx$/i.test(rec.name) ? { file, handle } : { file };
    if (typeof openDocumentFiles === "function") await openDocumentFiles([item]);
    else await ingestFile(item.file, item.handle ? { handle: item.handle } : {});
  }
  function whenText(ts) {
    const d = new Date(ts);
    const loc = I18N[currentLang]?.locale || "pl-PL";
    const today = new Date();
    if (d.toDateString() === today.toDateString()) return d.toLocaleTimeString(loc, { hour: "2-digit", minute: "2-digit" });
    return d.toLocaleDateString(loc, { day: "numeric", month: "short" });
  }
  let recentJob = 0;
  async function renderRecent() {
    const host = document.getElementById("recentFiles");
    if (!host) return;
    const job = ++recentJob;
    if (originalFileBytes) { host.hidden = true; return; }
    const list = await recentAll();
    if (job !== recentJob || originalFileBytes) return;
    host.replaceChildren();
    host.hidden = !list.length;
    if (!list.length) return;
    const h = document.createElement("h3");
    h.className = "recent-title";
    h.textContent = t("recentTitle");
    const ul = document.createElement("ul");
    ul.className = "recent-list";
    list.forEach((rec) => {
      const li = document.createElement("li");
      li.className = "recent-item";
      const open = document.createElement("button");
      open.type = "button";
      open.className = "recent-open";
      const name = document.createElement("span");
      name.className = "recent-name";
      name.textContent = rec.name;
      const meta = document.createElement("span");
      meta.className = "recent-meta";
      meta.textContent = `${whenText(rec.openedAt)} · ${formatFileSize(rec.size)}`;
      open.append(name, meta);
      open.addEventListener("click", async () => {
        open.disabled = true;
        try { await openRecent(rec); } catch (e) { log(String(e?.message || e), "error"); toast(t("recentOpenFailed"), "error"); }
        open.disabled = false;
      });
      const del = document.createElement("button");
      del.type = "button";
      del.className = "recent-del";
      del.textContent = "✕";
      del.setAttribute("aria-label", t("recentRemove", { name: rec.name }));
      del.dataset.hint = "";
      del.dataset.hintPl = "Usuń z listy (plik na dysku zostaje)";
      del.dataset.hintEn = "Remove from the list (the file on disk stays)";
      del.addEventListener("click", async () => {
        await rtx("readwrite", (s) => s.delete(rec.id)).catch(() => {});
        renderRecent();
        toast(t("recentRemoved", { name: rec.name }), "success", {
          action: t("undoAction"),
          onAction: () => rtx("readwrite", (s) => s.put(rec)).catch(() => {}).then(renderRecent),
        });
      });
      li.append(open, del);
      ul.append(li);
    });
    host.append(h, ul);
  }

  // wczytanie pliku przez użytkownika (nie przełączenie karty, nie szkic, nie przykład — te są „silent”)
  const origIngest = window.ingestFile;
  window.ingestFile = async function ingestFileComfort(file, options = {}) {
    rememberPlace(); // poprzedni dokument: zapamiętaj, gdzie skończyłeś
    const ok = await origIngest.call(this, file, options);
    if (ok && !options.silent && file) {
      const type = detectFileType(file.name, file.type);
      if (type === "docx") {
        offerResume(file);
        rememberRecent(file, options.handle);
      } else if (type === "pdf") rememberRecent(file, null);
    } else if (ok && options.silent) {
      resumeId = file ? `${file.name}|${file.size}` : null;
    }
    schedulePage();
    return ok;
  };
  const origClear = window.clearDocumentState;
  window.clearDocumentState = function clearDocumentStateComfort(...args) {
    rememberPlace();
    resumeId = null;
    hideReadHint();
    const r = origClear.apply(this, args);
    if (statusPageEl) statusPageEl.hidden = true;
    renderRecent();
    return r;
  };

  // ── D23: rozwinięte sekcje panelu ──────────────────────────────────────────
  const PANELS_KEY = "dwb-panel-open-v1";
  function restorePanels() {
    let ids = null;
    try { ids = JSON.parse(lsGet(PANELS_KEY) || "null"); } catch (_) {}
    if (!Array.isArray(ids)) return;
    document.querySelectorAll(".sidebar details.panel").forEach((d) => {
      const want = ids.includes(d.id);
      if (d.open !== want) d.open = want;
    });
  }
  function watchPanels() {
    document.querySelectorAll(".sidebar details.panel").forEach((d) => d.addEventListener("toggle", () => {
      const open = [...document.querySelectorAll(".sidebar details.panel")].filter((x) => x.open).map((x) => x.id);
      lsSet(PANELS_KEY, JSON.stringify(open));
    }));
  }

  function init() {
    restorePanels();
    watchPanels();
    applySavedMode();
    renderRecent();
    if (typeof applyLanguage === "function") {
      const orig = window.applyLanguage;
      window.applyLanguage = function applyLanguageComfort(...args) {
        const r = orig.apply(this, args);
        syncSelectionWords();
        schedulePage();
        renderRecent();
        return r;
      };
    }
  }
  document.addEventListener("DOMContentLoaded", init);

  return { showReadHint, hideReadHint, goToPage, syncPage, renderRecent, rememberPlace, enterEditAt };
})();
