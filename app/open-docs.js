// open-docs.js — kilka otwartych dokumentów naraz: karty nad dokumentem (prośba Mateusza
// 2026-10-03: „otwieranie wielu dokumentów”).
//
// Dlaczego karty („oddzielnie”), a nie sklejanie w jeden plik: każdy .docx ma własne style,
// numerację list, sekcje i obrazy — sklejony plik zapisywałby się jako NOWY dokument, a zmiany
// nie wracałyby do oryginałów. Karty: każdy plik zostaje sobą, „Zapisz” zapisuje ten na wierzchu.
//
// Jak to działa:
//   - Na wierzchu zawsze JEDEN dokument (cała aplikacja tak działa: podgląd, panele, Cofnij).
//     Dokumenty w tle to bajty pliku zbudowane jak przy „Zapisz” (z niezapisanymi zmianami),
//     nazwa, uchwyt pliku (zapis do oryginału), stan „niezapisane” i miejsce w dokumencie.
//   - Przełączenie = odłożenie bieżącego (buildDocumentForSave) + wczytanie wybranego. Historia
//     Cofnij zaczyna się od nowa po przełączeniu (zmiany zostają, tylko bez cofania sprzed).
//   - Otwarcie pliku, gdy coś jest otwarte, DOKŁADA kartę (dawniej podmieniało dokument po
//     pytaniu o porzucenie zmian). Kilka plików naraz (wybór wielu, upuszczenie, „Otwórz za
//     pomocą”): pierwszy się otwiera, pozostałe czekają jako karty i wczytują się po stuknięciu
//     (PDF konwertuje się dopiero wtedy — telefon nie mieli wszystkiego naraz).
//   - Szkic odzyskiwania (drafts.js): każdy niezapisany dokument w tle ma własny wpis — gdy
//     system zabije aplikację, karta „Niezapisana praca” pokaże każdy z nich osobno.
//   - Ten sam .docx drugi raz → przełączenie na jego kartę, bez duplikatu. PDF drugi raz =
//     konwersja od nowa w nowej karcie (np. ponownie, tym razem z rozpoznawaniem tekstu).

const OPEN_DOCS_MAX = 12;
const OPEN_DOCS_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const dwbOpenDocs = (() => {
  // { id, name, handle, stamp:{name,size,lastModified}|null, file (czeka na wczytanie) | null,
  //   bytes (w tle) | null, dirty, draft:{id,stamp}|null, anchor, pdf }
  const docs = [];
  let activeId = null;
  let seq = 0;
  let inner = 0; // >0 = trwa nasze wczytywanie (zagnieżdżone ingestFile przechodzą bez kart)
  let job = Promise.resolve();
  let tabsEl = null, rowEl = null, addBtn = null;

  const active = () => docs.find((d) => d.id === activeId) || null;
  const byId = (id) => docs.find((d) => d.id === id) || null;
  const fileStamp = (f) => (f ? { name: f.name, size: f.size, lastModified: f.lastModified } : null);
  const sameStamp = (a, b) => !!(a && b && a.name === b.name && a.size === b.size && a.lastModified === b.lastModified);
  const queue = (fn) => (job = job.then(fn, fn));

  async function sameHandle(a, b) {
    if (!a || !b || typeof a.isSameEntry !== "function") return false;
    try { return await a.isSameEntry(b); } catch (_) { return false; }
  }
  async function findOpen(file, handle) {
    if (detectFileType(file?.name, file?.type) === "pdf") return null;
    for (const d of docs) {
      if (d.pdf) continue;
      if (handle && (await sameHandle(handle, d.handle))) return d;
      if (!handle && !d.handle && sameStamp(fileStamp(file), d.stamp)) return d;
    }
    return null;
  }

  // ── odłożenie dokumentu z wierzchu ─────────────────────────────────────────
  async function park(d) {
    if (!d || !originalFileBytes) return;
    d.anchor = typeof captureDocScrollAnchor === "function" ? captureDocScrollAnchor() : null;
    d.bytes = await buildDocumentForSave();
    d.name = currentFileName || d.name;
    d.handle = fileHandle || null;
    d.dirty = !!hasUnsavedChanges;
    d.draft = typeof dwbDrafts !== "undefined" ? await dwbDrafts.park() : null;
    d.file = null;
  }
  function unparkFailed(d) {
    // nowy plik się nie otworzył (Anuluj przy PDF, zły plik) — bieżący zostaje na wierzchu
    if (typeof dwbDrafts !== "undefined") dwbDrafts.unpark(d.draft);
    d.bytes = null;
    d.draft = null;
  }

  // Wczytanie karty na wierzch (z bajtów albo z pliku, który jeszcze czekał).
  async function load(d) {
    inner++;
    try {
      let ok;
      if (d.bytes) {
        ok = await window.ingestFile(new File([d.bytes], d.name, { type: OPEN_DOCS_MIME }), { silent: true, ...(d.handle ? { handle: d.handle } : {}) });
        if (ok) {
          if (d.dirty) {
            if (typeof dwbDrafts !== "undefined") dwbDrafts.unpark(d.draft);
            setDirtyState(true);
          }
          if (d.anchor && typeof restoreDocScrollAnchor === "function") restoreDocScrollAnchor(d.anchor);
        }
      } else if (d.file) {
        ok = await window.ingestFile(d.file, d.handle && !d.pdf ? { handle: d.handle } : {});
      }
      if (ok) {
        d.bytes = null;
        d.file = null;
        d.draft = null;
        d.name = currentFileName || d.name;
        d.handle = fileHandle || null;
        activeId = d.id;
      }
      return !!ok;
    } finally {
      inner--;
    }
  }

  function switchTo(id) {
    return queue(async () => {
      const target = byId(id);
      const cur = active();
      if (!target || target === cur) return false;
      if (cur) await park(cur);
      const ok = await load(target);
      if (!ok) {
        toast(t("docTabSwitchFailed", { name: target.name }), "error");
        if (!target.bytes) docs.splice(docs.indexOf(target), 1); // plik nie do otwarcia — bez martwej karty
        if (cur) await load(cur);
      }
      render();
      return ok;
    });
  }

  // ── nowe pliki ─────────────────────────────────────────────────────────────
  // items: [{ file, handle? }] — pierwszy otwiera się od razu, reszta czeka jako karty.
  async function openFiles(items) {
    const list = (items || []).filter((it) => it && it.file);
    if (!list.length) return false;
    if (list.length === 1) {
      const type = detectFileType(list[0].file.name, list[0].file.type);
      if (type !== "docx" && type !== "pdf") return window.ingestFile(list[0].file); // komunikat „nieobsługiwany”
    }
    let first = null;
    const waiting = [];
    let room = OPEN_DOCS_MAX - docs.length;
    for (const it of list) {
      const known = await findOpen(it.file, it.handle);
      if (known) {
        if (!first) first = known;
        if (list.length === 1) toast(t("docTabAlready", { name: known.name }), "info");
        continue;
      }
      const type = detectFileType(it.file.name, it.file.type);
      if (type !== "docx" && type !== "pdf") continue;
      if (room <= 0) {
        toast(t("docTabsLimit", { n: OPEN_DOCS_MAX }), "warning");
        break;
      }
      room--;
      const d = { id: `t${++seq}`, name: it.file.name, handle: it.handle || null, stamp: fileStamp(it.file), file: it.file, bytes: null, dirty: false, draft: null, anchor: null, pdf: type === "pdf" };
      if (!first) first = d;
      else waiting.push(d);
    }
    if (!first) { render(); return false; }
    let ok = true;
    if (first.file && !docs.includes(first)) {
      // pierwszy nowy plik: przez zwykłe ingestFile (karta powstaje w opakowaniu niżej)
      ok = await window.ingestFile(first.file, first.handle && !first.pdf ? { handle: first.handle } : {});
    } else if (first.id !== activeId) {
      ok = await switchTo(first.id);
    }
    await job;
    const added = waiting.length;
    docs.push(...waiting);
    render();
    if (added && docs.length > 1) {
      const n = docs.length;
      const few = n % 10 >= 2 && n % 10 <= 4 && (n % 100 < 12 || n % 100 > 14);
      toast(t(few ? "docTabsAdded" : "docTabsAddedMany", { n }), "info");
    }
    return ok;
  }

  // ── zamykanie ──────────────────────────────────────────────────────────────
  function closeTab(id) {
    const d = byId(id);
    if (!d) return;
    if (d.id === activeId) {
      if (typeof requestCloseDocument === "function") requestCloseDocument();
      return;
    }
    if (d.dirty && !window.confirm(t("closeDocWarn"))) return;
    if (typeof dwbDrafts !== "undefined") dwbDrafts.drop(d.draft);
    docs.splice(docs.indexOf(d), 1);
    render();
  }

  // ── karty (UI) ─────────────────────────────────────────────────────────────
  function render() {
    if (!tabsEl) return;
    const cur = active();
    // W trakcie wczytywania na ekranie jest już NOWY plik, a karta na wierzchu to jeszcze
    // odłożony (bytes) — nie przepisuj jej nazwy ani „niezapisane” ze stanu ekranu.
    if (cur && !inner && !cur.bytes) {
      cur.name = currentFileName || cur.name;
      cur.dirty = !!hasUnsavedChanges;
    }
    const show = docs.length > 1;
    if (tabsEl.hidden === show) {
      tabsEl.hidden = !show;
      document.body.classList.toggle("has-doc-tabs", show);
      // rząd kart przesuwa obszar dokumentu: panel obok (≥1024 px) i wysokość dokumentu w dół
      if (typeof appFrame !== "undefined") appFrame.syncDock?.();
      if (typeof syncDocViewportHeight === "function") syncDocViewportHeight();
    }
    if (!show) { rowEl.replaceChildren(); return; }
    const frag = document.createDocumentFragment();
    for (const d of docs) {
      const on = d.id === activeId;
      const wrap = document.createElement("div");
      wrap.className = "doc-tab" + (on ? " is-active" : "") + (d.dirty ? " is-dirty" : "") + (d.file ? " is-waiting" : "");
      wrap.dataset.tab = d.id;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "doc-tab-btn";
      btn.setAttribute("role", "tab");
      btn.setAttribute("aria-selected", on ? "true" : "false");
      btn.tabIndex = on ? 0 : -1;
      const extra = [d.dirty ? t("docTabUnsaved") : "", d.file ? t("docTabNotLoaded") : ""].filter(Boolean).join(", ");
      btn.setAttribute("aria-label", extra ? `${d.name} — ${extra}` : d.name);
      btn.setAttribute("data-hint", "");
      btn.setAttribute("data-hint-pl", extra ? `${d.name} — ${extra}` : d.name);
      btn.setAttribute("data-hint-en", extra ? `${d.name} — ${extra}` : d.name);
      btn.setAttribute("data-hint-delay", "0.6");
      const dot = document.createElement("span");
      dot.className = "doc-tab-dot";
      dot.setAttribute("aria-hidden", "true");
      const name = document.createElement("span");
      name.className = "doc-tab-name";
      name.textContent = d.name.replace(/\.(docx|pdf)$/i, "");
      btn.append(dot, name);
      if (d.pdf && d.file) {
        const badge = document.createElement("span");
        badge.className = "doc-tab-kind";
        badge.textContent = "PDF";
        badge.setAttribute("aria-hidden", "true");
        btn.append(badge);
      }
      const x = document.createElement("button");
      x.type = "button";
      x.className = "doc-tab-close";
      x.tabIndex = -1;
      x.setAttribute("aria-label", t("docTabClose", { name: d.name }));
      x.innerHTML = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>';
      wrap.append(btn, x);
      frag.append(wrap);
    }
    rowEl.replaceChildren(frag);
    rowEl.setAttribute("aria-label", t("docTabsAria"));
    addBtn?.setAttribute("aria-label", t("docTabAdd"));
    // karta na wierzchu widoczna w przewijanym rzędzie
    const onEl = rowEl.querySelector(".doc-tab.is-active");
    if (onEl) {
      const r = onEl.getBoundingClientRect(), rr = rowEl.getBoundingClientRect();
      if (r.left < rr.left || r.right > rr.right) rowEl.scrollLeft += r.left < rr.left ? r.left - rr.left - 12 : r.right - rr.right + 12;
    }
    rowEl._dwbFadeSync?.();
  }
  let renderQueued = false;
  function renderSoon() {
    if (renderQueued) return;
    renderQueued = true;
    requestAnimationFrame(() => { renderQueued = false; render(); });
  }

  function initUi() {
    tabsEl = document.getElementById("docTabs");
    rowEl = document.getElementById("docTabsRow");
    addBtn = document.getElementById("docTabAdd");
    if (!tabsEl || !rowEl) return;
    if (typeof attachOverflowFade === "function") attachOverflowFade(rowEl);
    rowEl.addEventListener("click", (e) => {
      const wrap = e.target.closest(".doc-tab");
      if (!wrap) return;
      if (e.target.closest(".doc-tab-close")) closeTab(wrap.dataset.tab);
      else if (wrap.dataset.tab !== activeId) switchTo(wrap.dataset.tab);
    });
    // środkowy przycisk myszy zamyka kartę (jak w przeglądarce)
    rowEl.addEventListener("auxclick", (e) => {
      const wrap = e.button === 1 && e.target.closest(".doc-tab");
      if (wrap) { e.preventDefault(); closeTab(wrap.dataset.tab); }
    });
    // ←/→ między kartami (role=tablist), Delete zamyka
    rowEl.addEventListener("keydown", (e) => {
      const wrap = e.target.closest(".doc-tab");
      if (!wrap) return;
      const tabs = [...rowEl.querySelectorAll(".doc-tab")];
      const i = tabs.indexOf(wrap);
      let next = null;
      if (e.key === "ArrowRight") next = tabs[(i + 1) % tabs.length];
      else if (e.key === "ArrowLeft") next = tabs[(i - 1 + tabs.length) % tabs.length];
      else if (e.key === "Home") next = tabs[0];
      else if (e.key === "End") next = tabs[tabs.length - 1];
      else if (e.key === "Delete") { e.preventDefault(); closeTab(wrap.dataset.tab); return; }
      if (!next) return;
      e.preventDefault();
      next.querySelector(".doc-tab-btn")?.focus();
    });
    addBtn?.addEventListener("click", () => { if (typeof openFilePicker === "function") openFilePicker(); });
  }

  // ── podpięcie pod aplikację ────────────────────────────────────────────────
  function init() {
    initUi();
    if (typeof dwbDrafts !== "undefined") dwbDrafts.setHeldIds(() => docs.filter((d) => d.id !== activeId && d.draft).map((d) => d.draft.id));

    // Otwieranie nie porzuca już bieżącego dokumentu (zostaje w karcie) — bez pytania.
    window.confirmDiscardChanges = function confirmDiscardChangesTabs() {
      if (docs.length >= OPEN_DOCS_MAX && originalFileBytes) {
        toast(t("docTabsLimit", { n: OPEN_DOCS_MAX }), "warning");
        return false;
      }
      return true;
    };

    // Każdy nowy dokument (Otwórz, upuszczenie, przykład, nowy z szablonu, odzyskany szkic,
    // PDF po konwersji) przechodzi przez ingestFile — tu bieżący idzie do karty w tle.
    const origIngest = window.ingestFile;
    window.ingestFile = async function ingestFileTabs(file, options = {}) {
      if (inner) return origIngest.call(this, file, options);
      if (originalFileBytes && docs.length >= OPEN_DOCS_MAX) {
        toast(t("docTabsLimit", { n: OPEN_DOCS_MAX }), "warning");
        return false;
      }
      const cur = originalFileBytes ? active() : null;
      if (originalFileBytes && !cur) {
        // dokument otwarty zanim moduł ruszył (nie powinno się zdarzyć) — dopisz go
        docs.push({ id: `t${++seq}`, name: currentFileName, handle: fileHandle || null, stamp: null, file: null, bytes: null, dirty: !!hasUnsavedChanges, draft: null, anchor: null, pdf: false });
        activeId = docs[docs.length - 1].id;
      }
      const prev = active();
      return queue(async () => {
        if (prev) await park(prev);
        let ok = false;
        inner++;
        try {
          ok = await origIngest.call(this, file, options);
        } finally {
          inner--;
        }
        if (ok) {
          const type = detectFileType(file?.name, file?.type);
          const d = { id: `t${++seq}`, name: currentFileName, handle: fileHandle || null, stamp: options.handle || type === "pdf" ? null : fileStamp(file), file: null, bytes: null, dirty: false, draft: null, anchor: null, pdf: type === "pdf" };
          const at = prev ? docs.indexOf(prev) + 1 : docs.length;
          docs.splice(at, 0, d);
          activeId = d.id;
        } else if (prev) {
          unparkFailed(prev);
        }
        render();
        return ok;
      });
    };

    // Zamknięcie dokumentu na wierzchu: jego karta znika, na wierzch wchodzi poprzednia.
    const origClear = window.clearDocumentState;
    window.clearDocumentState = function clearDocumentStateTabs(...args) {
      const r = origClear.apply(this, args);
      const cur = active();
      if (cur) {
        const i = docs.indexOf(cur);
        docs.splice(i, 1);
        activeId = null;
        const next = docs[Math.max(0, i - 1)];
        if (next) queue(() => load(next).then((ok) => { if (!ok) docs.splice(docs.indexOf(next), 1); render(); }));
      }
      render();
      return r;
    };

    // „niezapisane” na karcie na bieżąco
    const origSetDirty = window.setDirtyState;
    window.setDirtyState = function setDirtyStateTabs(...args) {
      const r = origSetDirty.apply(this, args);
      renderSoon();
      return r;
    };
    if (typeof applyLanguage === "function") {
      const origLang = window.applyLanguage;
      window.applyLanguage = function applyLanguageTabs(...args) {
        const r = origLang.apply(this, args);
        renderSoon();
        return r;
      };
    }
  }

  return { init, openFiles, switchTo, closeTab, list: () => docs.map((d) => ({ id: d.id, name: d.name, dirty: d.id === activeId ? !!hasUnsavedChanges : d.dirty, active: d.id === activeId, waiting: !!d.file })), _idle: () => job };
})();

// Wiele plików naraz (wybór wielu, upuszczenie, „Otwórz za pomocą”).
function openDocumentFiles(items) {
  return dwbOpenDocs.openFiles(items);
}

document.addEventListener("DOMContentLoaded", () => dwbOpenDocs.init());
