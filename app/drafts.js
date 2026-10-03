// drafts.js — automatyczny szkic niezapisanej pracy + odzyskiwanie po zamknięciu aplikacji.
//
// iOS / Android potrafią zamknąć aplikację w tle (brak pamięci) bez żadnego zdarzenia
// „zamykam się”; na komputerze — awaria przeglądarki, prąd, restart. Wtedy niezapisane
// zmiany przepadały. Tu: szkic w IndexedDB (tylko na tym urządzeniu), przy następnym
// otwarciu karta „Niezapisana praca … [Przywróć] [Odrzuć]” na ekranie startowym.
//
// Co i kiedy:
//   - bajty pliku (originalFileBytes) — tylko gdy się zmieniły (otwarcie, operacja z panelu),
//   - lista zmian z podglądu (akapity) + oczekujące operacje — mała, ~1,5 s po zmianie,
//   - NATYCHMIAST przy przejściu w tło (visibilitychange → hidden, pagehide) — to ostatnia
//     chwila przed ewentualnym zabiciem karty przez system.
// Szkic znika: po zapisie (brak niezapisanych zmian), po „Odrzuć”, po świadomym zamknięciu
// dokumentu / otwarciu innego (wtedy aplikacja już zapytała o porzucenie zmian) i sam po 7 dniach.
// Każde okno ma własny szkic (id sesji). Żyjące okna odpowiadają na BroadcastChannel — ich
// szkice nie są pokazywane do odzyskania (to nie jest zgubiona praca).
// Uchwyt pliku (Chrome/Edge) też trafia do szkicu: po odzyskaniu „Zapisz” może zapisać do
// oryginału — chyba że plik na dysku zmienił się od czasu szkicu (wtedy zapis = kopia).

const DRAFT_DB = "dwb-drafts";
const DRAFT_STORE = "drafts"; // opis + lista zmian (małe, zapisywane często)
const DRAFT_BYTES = "bytes"; // bajty pliku (duże, zapisywane tylko gdy się zmieniły)
const DRAFT_MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000; // prośba Mateusza: 7 dni, nie 14
const DRAFT_DEBOUNCE_MS = 1500;
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const dwbDrafts = (() => {
  let sessionId = newDraftId();
  let timer = 0;
  let saving = null;
  let storedBytes = null; // bajty, które już leżą w szkicu tej sesji
  let stamp = null; // { size, lastModified } oryginału na dysku, gdy szkic powstał
  let persistAsked = false;
  let suspended = 0; // >0 = trwa przywracanie (ingestFile nie kasuje szkicu)
  const channel = typeof BroadcastChannel === "function" ? new BroadcastChannel("dwb-drafts") : null;

  function newDraftId() {
    return `d${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  }

  // ── IndexedDB ──────────────────────────────────────────────────────────────
  let dbPromise = null;
  function db() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        if (typeof indexedDB === "undefined") { reject(new Error("no indexedDB")); return; }
        const req = indexedDB.open(DRAFT_DB, 1);
        req.onupgradeneeded = () => {
          req.result.createObjectStore(DRAFT_STORE, { keyPath: "id" });
          req.result.createObjectStore(DRAFT_BYTES);
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      }).catch((e) => { dbPromise = null; throw e; });
    }
    return dbPromise;
  }
  async function tx(mode, fn) {
    const d = await db();
    return new Promise((resolve, reject) => {
      const t = d.transaction([DRAFT_STORE, DRAFT_BYTES], mode);
      const store = t.objectStore(DRAFT_STORE);
      let result;
      Promise.resolve(fn(store, t.objectStore(DRAFT_BYTES))).then((r) => { result = r; });
      t.oncomplete = () => resolve(result);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    });
  }
  const reqP = (r) => new Promise((resolve, reject) => { r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error); });
  const getRec = (id) => tx("readonly", async (s, b) => {
    const rec = await reqP(s.get(id));
    if (rec) rec.bytes = await reqP(b.get(id));
    return rec;
  });
  const allRecs = () => tx("readonly", (s) => reqP(s.getAll()));
  const delRec = (id) => tx("readwrite", (s, b) => { s.delete(id); b.delete(id); }).catch(() => {});

  // ── zapis szkicu ───────────────────────────────────────────────────────────
  async function collect() {
    if (typeof waitInlineStructuralIdle === "function") await waitInlineStructuralIdle();
    return {
      inline: typeof collectInlineParagraphEdits === "function" ? collectInlineParagraphEdits() : [],
      pending: Array.isArray(pendingDocEdits) ? pendingDocEdits.map((e) => ({ ...e })) : [],
      changes: typeof dwbUndo !== "undefined" && dwbUndo.changesSinceSave ? dwbUndo.changesSinceSave() : 0,
    };
  }

  async function writeNow() {
    clearTimeout(timer);
    timer = 0;
    if (!originalFileBytes || currentFileType !== "docx" || !hasUnsavedChanges) return;
    const bytes = originalFileBytes;
    const id = sessionId;
    const data = await collect();
    if (!stamp && fileHandle) {
      try { const f = await fileHandle.getFile(); stamp = { size: f.size, lastModified: f.lastModified }; } catch (_) { /* brak zgody — bez odcisku */ }
    }
    const rec = {
      id, fileName: currentFileName || "dokument.docx", savedAt: Date.now(), data, stamp,
      handle: fileHandle || null,
    };
    const writeBytes = storedBytes !== bytes; // plik (MB) tylko gdy się zmienił; lista zmian zawsze
    try {
      await tx("readwrite", (s, b) => {
        if (writeBytes) b.put(bytes, id);
        try { s.put(rec); } catch (_) { s.put({ ...rec, handle: null }); } // uchwytu nie da się zapisać (np. Safari) — bez niego
      });
      storedBytes = bytes;
      if (!persistAsked) {
        persistAsked = true;
        // „trwałe” miejsce: Safari / Chrome nie czyszczą wtedy danych przy braku miejsca
        navigator.storage?.persisted?.().then((p) => { if (!p) navigator.storage.persist?.(); }).catch(() => {});
      }
    } catch (e) {
      // tryb prywatny / brak miejsca — aplikacja działa dalej, tylko bez szkicu
      if (typeof log === "function") log(`Szkic niezapisany: ${e?.message || e}`, "warn");
    }
  }

  function save(delay = DRAFT_DEBOUNCE_MS) {
    if (suspended) return;
    clearTimeout(timer);
    timer = setTimeout(() => { saving = writeNow().finally(() => { saving = null; }); }, delay);
  }
  function saveImmediately() {
    if (suspended || !hasUnsavedChanges || !originalFileBytes) return;
    saving = writeNow().finally(() => { saving = null; });
  }

  async function dropOwn() {
    const id = sessionId; // karta dokumentu mogła już przejąć inny szkic (unpark) — kasujemy TEN
    clearTimeout(timer);
    timer = 0;
    if (saving) await saving.catch(() => {});
    if (id !== sessionId) return;
    storedBytes = null;
    stamp = null;
    await delRec(id);
  }

  // ── kilka otwartych dokumentów (open-docs.js) ──────────────────────────────
  // Przełączenie na inną kartę: szkic bieżącego dokumentu zapisany TERAZ i odłożony (zostaje
  // w bazie pod swoim id), sesja dostaje nowe id. Powrót na kartę = unpark (dalsze zmiany
  // nadpisują ten sam wpis). Czysty dokument nie zostawia szkicu.
  let heldIds = () => [];
  async function park() {
    clearTimeout(timer);
    timer = 0;
    if (saving) await saving.catch(() => {});
    let out = null;
    if (originalFileBytes && hasUnsavedChanges && currentFileType === "docx") {
      await writeNow();
      out = { id: sessionId, stamp };
    } else await dropOwn();
    sessionId = newDraftId();
    storedBytes = null;
    stamp = null;
    return out;
  }
  function unpark(p) {
    if (!p?.id) return;
    sessionId = p.id;
    storedBytes = null;
    stamp = p.stamp || null;
  }
  const drop = (p) => (p?.id ? delRec(p.id) : Promise.resolve());

  // ── odzyskiwanie ───────────────────────────────────────────────────────────
  // Okna, które żyją (np. drugi plik otwarty w osobnym oknie), odpowiadają — ich szkice
  // to nie jest zgubiona praca.
  function aliveIds(timeoutMs = 250) {
    if (!channel) return Promise.resolve(new Set());
    const ids = new Set();
    const onMsg = (e) => { if (e.data?.type === "alive" && e.data.id) ids.add(e.data.id); };
    channel.addEventListener("message", onMsg);
    channel.postMessage({ type: "who" });
    return new Promise((r) => setTimeout(() => { channel.removeEventListener("message", onMsg); r(ids); }, timeoutMs));
  }
  channel?.addEventListener("message", (e) => {
    if (e.data?.type !== "who") return;
    if (originalFileBytes) channel.postMessage({ type: "alive", id: sessionId });
    for (const id of heldIds()) channel.postMessage({ type: "alive", id }); // karty w tle tego okna
  });

  async function recoverable() {
    let recs;
    try { recs = await allRecs(); } catch (_) { return []; }
    const now = Date.now();
    const expired = recs.filter((r) => !r.savedAt || now - r.savedAt > DRAFT_MAX_AGE_MS);
    expired.forEach((r) => delRec(r.id));
    const alive = await aliveIds();
    return recs
      .filter((r) => !expired.includes(r) && r.id !== sessionId && !alive.has(r.id) && !heldIds().includes(r.id))
      .sort((a, b) => b.savedAt - a.savedAt);
  }

  async function restore(id) {
    const rec = await getRec(id);
    if (!rec?.bytes) { toast(t("draftGone"), "info"); renderRecovery(); return false; }
    // uchwyt oryginału: zgoda (to kliknięcie „Przywróć” = gest użytkownika) + czy plik się nie zmienił
    let handle = null;
    if (rec.handle && typeof rec.handle.requestPermission === "function") {
      try {
        const perm = await rec.handle.requestPermission({ mode: "readwrite" });
        if (perm === "granted") {
          const f = await rec.handle.getFile();
          if (rec.stamp && (f.size !== rec.stamp.size || f.lastModified !== rec.stamp.lastModified)) toast(t("draftOriginalChanged", { name: rec.fileName }), "warning");
          else handle = rec.handle;
        }
      } catch (_) { handle = null; }
    }
    // biblioteki .docx dociągają się po starcie — szybkie „Przywróć” nie może ich wyprzedzić
    if (typeof ensureDocLibs === "function" && !(await ensureDocLibs(true))) return false;
    const edits = [...(rec.data?.pending || [])];
    if (rec.data?.inline?.length) edits.push({ op: "paragraphBatch", items: rec.data.inline });
    let bytes = rec.bytes instanceof Uint8Array ? rec.bytes : new Uint8Array(rec.bytes);
    if (edits.length) bytes = (await buildPatchedDocx(bytes, edits)).bytes;
    suspended++;
    try {
      const ok = await ingestFile(new File([bytes], rec.fileName, { type: DOCX_MIME }), { silent: true, ...(handle ? { handle } : {}) });
      if (!ok) return false;
    } finally { suspended--; }
    // ta sesja przejmuje szkic (dalsze zmiany nadpisują ten sam wpis)
    sessionId = rec.id;
    storedBytes = null;
    stamp = handle ? rec.stamp : null;
    setDirtyState(true);
    toast(t("draftRestored", { name: rec.fileName }), "success");
    saveImmediately();
    renderRecovery();
    return true;
  }

  async function discard(id, name) {
    if (!window.confirm(t("draftDiscardConfirm", { name }))) return;
    await delRec(id);
    renderRecovery();
  }

  // ── karta na ekranie startowym ─────────────────────────────────────────────
  function when(ts) {
    const d = new Date(ts);
    const loc = (typeof I18N !== "undefined" && I18N[currentLang]?.locale) || "pl-PL";
    const time = d.toLocaleTimeString(loc, { hour: "2-digit", minute: "2-digit" });
    const today = new Date();
    const y = new Date(); y.setDate(today.getDate() - 1);
    if (d.toDateString() === today.toDateString()) return t("draftToday", { time });
    if (d.toDateString() === y.toDateString()) return t("draftYesterday", { time });
    return `${d.toLocaleDateString(loc, { day: "numeric", month: "long" })}, ${time}`;
  }

  let renderJob = 0;
  async function renderRecovery() {
    const job = ++renderJob;
    const host = document.getElementById("draftRecovery");
    if (!host) return;
    if (originalFileBytes) { host.replaceChildren(); host.hidden = true; return; }
    const recs = await recoverable();
    if (job !== renderJob || originalFileBytes) return;
    host.replaceChildren();
    host.hidden = !recs.length;
    recs.slice(0, 5).forEach((r) => {
      const card = document.createElement("div");
      card.className = "draft-card";
      card.setAttribute("role", "group");
      const head = document.createElement("div");
      head.className = "draft-head";
      head.textContent = t("draftTitle");
      const name = document.createElement("div");
      name.className = "draft-name";
      name.textContent = r.fileName;
      const meta = document.createElement("div");
      meta.className = "draft-meta";
      const n = r.data?.changes || 0;
      meta.textContent = [t("draftChangedAt", { when: when(r.savedAt) }), n ? t("draftChanges", { count: n }) : ""].filter(Boolean).join(" · ");
      const actions = document.createElement("div");
      actions.className = "draft-actions";
      const ok = Object.assign(document.createElement("button"), { type: "button", className: "btn primary", textContent: t("draftRestore") });
      ok.addEventListener("click", async () => {
        ok.disabled = true;
        try { await restore(r.id); } catch (e) { log(String(e?.message || e), "error"); toast(t("draftRestoreFailed"), "error"); ok.disabled = false; }
      });
      const no = Object.assign(document.createElement("button"), { type: "button", className: "btn", textContent: t("draftDiscard") });
      no.addEventListener("click", () => discard(r.id, r.fileName));
      actions.append(ok, no);
      card.append(head, name, meta, actions);
      host.append(card);
    });
    if (recs.length) {
      const note = document.createElement("p");
      note.className = "draft-note";
      note.textContent = t("draftNote");
      host.append(note);
    }
  }

  // ── podpięcie pod aplikację ────────────────────────────────────────────────
  function init() {
    const origSetDirty = window.setDirtyState;
    window.setDirtyState = function setDirtyStateDrafts(isDirty, ...rest) {
      const r = origSetDirty.call(this, isDirty, ...rest);
      if (isDirty) save();
      else if (!suspended) dropOwn(); // zapisane / cofnięte do stanu z pliku / nowy dokument
      return r;
    };
    // nowy plik albo zamknięcie: aplikacja już zapytała o porzucenie zmian → szkic tej sesji precz
    const origIngest = window.ingestFile;
    window.ingestFile = async function ingestFileDrafts(...args) {
      if (!suspended) { await dropOwn(); sessionId = newDraftId(); }
      return origIngest.apply(this, args);
    };
    const origClear = window.clearDocumentState;
    window.clearDocumentState = function clearDocumentStateDrafts(...args) {
      const r = origClear.apply(this, args);
      dropOwn().then(() => { sessionId = newDraftId(); renderRecovery(); });
      return r;
    };
    // pisanie w podglądzie (bez setDirtyState przy każdej literze)
    docCanvasEl?.addEventListener("input", () => { if (hasUnsavedChanges) save(); }, true);
    // ostatnia chwila przed ewentualnym zabiciem karty przez system
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") saveImmediately(); });
    window.addEventListener("pagehide", saveImmediately);
    renderRecovery();
  }

  return {
    init, save, saveImmediately, renderRecovery, restore, recoverable, park, unpark, drop,
    setHeldIds(fn) { heldIds = typeof fn === "function" ? fn : () => []; },
    _id: () => sessionId, _flush: () => (saving || Promise.resolve()),
  };
})();

document.addEventListener("DOMContentLoaded", () => dwbDrafts.init());
