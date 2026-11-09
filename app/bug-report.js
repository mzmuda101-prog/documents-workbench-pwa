// bug-report.js — „Zgłoś błąd”: plik z dokumentem z początku sesji, dziennikiem działań i stanem
// końcowym — z niego powstaje powtarzalny test (scripts/replay-report.js). Dawniej zgłoszenie
// to było nagranie ekranu, z którego kroki trzeba było odgadywać klatka po klatce (2026-11-08).
//
// Dziennik jest BIERNY (tylko słucha, niczego nie zmienia) i ograniczony do ostatnich 400 działań:
//   - klawisz w dokumencie — z miejscem kursora (akapit:znak, zaznaczenie) PRZED naciśnięciem,
//     więc odtworzenie nie zależy od współrzędnych myszy ani wielkości okna,
//   - klik w tekst — miejsce kursora po kliknięciu,
//   - wklejenie — tekst i HTML ze schowka (przycięte),
//   - operacja z paska / menu / panelu (applyDocumentEdit) — sama operacja,
//   - Cofnij / Ponów (każdą drogą — zdarzenie „dwb:undo-step” z undo.js), Czytanie ⇄ Edycja.
// Plik zostaje u użytkownika (pobranie) — zawiera jego dokument; wysyła go tylko on sam.

const dwbBugReport = (() => {
  const MAX = 400;
  const CUT = 20000; // dłuższe napisy (obraz w base64, wielka wklejka) — przycięte
  let log = [];
  let start = null; // { bytes, name } — dokument z początku sesji (otwarcie / nowy)
  const t0 = Date.now();

  const host = () => docCanvasEl?.querySelector(".docx-preview-host");
  const inDoc = (el) => !!(el && docCanvasEl?.contains(el));
  const cut = (v) => {
    if (typeof v === "string") return v.length > CUT ? `${v.slice(0, CUT)}…[przycięte ${v.length}]` : v;
    if (Array.isArray(v)) return v.map(cut);
    if (v && typeof v === "object") return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, cut(x)]));
    return v;
  };
  function push(entry) {
    log.push({ ms: Date.now() - t0, ...entry });
    if (log.length > MAX) log.shift();
  }
  // miejsce w dokumencie: akapit (jak w pliku) i znak; zaznaczenie — oba końce
  function spot() {
    const sel = window.getSelection?.();
    const h = host();
    if (!sel?.rangeCount || !h || typeof collectPreviewParagraphElements !== "function") return {};
    const ps = collectPreviewParagraphElements(h);
    const at = (node, off) => {
      const i = ps.findIndex((p) => p.contains(node));
      if (i < 0) return null;
      const r = document.createRange();
      r.selectNodeContents(ps[i]);
      try { r.setEnd(node, off); } catch (_) { return [i, 0]; }
      return [i, r.toString().replace(/﻿/g, "").length];
    };
    const a = at(sel.anchorNode, sel.anchorOffset);
    if (!a) return {};
    if (sel.isCollapsed) return { at: a };
    return { at: a, to: at(sel.focusNode, sel.focusOffset) };
  }
  // Klawisz / wklejenie / klik w tekst same wywołują operacje (Backspace w liście → applyDocumentEdit,
  // Ctrl/⌘+Z → dwbUndo.undo). Takie operacje to SKUTEK — oznaczone fromInput, odtworzenie je pomija
  // (inaczej zrobiłoby je dwa razy). Znacznik trwa do końca bieżącego zadania (też po await).
  let inInput = false;
  const markInput = () => { inInput = true; setTimeout(() => { inInput = false; }, 0); };
  // pisanie, które aplikacja sama wstawia przy wklejaniu — nie jest osobnym działaniem (zwykłe
  // pisanie, też klawiatura ekranowa / polskie litery bez zdarzenia klawisza — zawsze w dzienniku)
  let inPaste = false;
  const mods = (e) => [e.ctrlKey && "Control", e.metaKey && "Meta", e.altKey && "Alt", e.shiftKey && "Shift"].filter(Boolean).join("+");

  document.addEventListener("keydown", (e) => {
    if (!inDoc(e.target) || ["Shift", "Control", "Alt", "Meta", "CapsLock", "Fn"].includes(e.key)) return;
    push({ a: "key", key: e.key, mods: mods(e), ...spot() });
    markInput();
  }, true);
  // klawiatura ekranowa / IME: litery bez zwykłego zdarzenia klawisza
  document.addEventListener("beforeinput", (e) => {
    if (!inDoc(e.target) || (e.inputType !== "insertText" && e.inputType !== "insertCompositionText")) return;
    if (inPaste) return; // wstawione przez aplikację w trakcie wklejania (WebKit: wklejka przez insertText)
    const last = log[log.length - 1];
    if (last?.a === "key" && last.key === e.data) return; // zwykły klawisz — już w dzienniku
    push({ a: "text", data: e.data, ...spot() });
    markInput();
  }, true);
  document.addEventListener("mouseup", (e) => {
    if (!inDoc(e.target)) return;
    setTimeout(() => push({ a: "click", ...spot() }), 0); // miejsce po kliknięciu
  }, true);
  document.addEventListener("paste", (e) => {
    if (!inDoc(e.target)) return;
    push({ a: "paste", text: cut(e.clipboardData?.getData("text/plain") || ""), html: cut(e.clipboardData?.getData("text/html") || ""), ...spot() });
    markInput();
    inPaste = true;
    setTimeout(() => { inPaste = false; }, 0);
  }, true);

  // operacje z paska / menu / paneli i Cofnij / Ponów — owinięte po załadowaniu wszystkich plików
  window.addEventListener("load", () => {
    readModeEl?.addEventListener("change", () => push({ a: "mode", mode: readModeEl.checked ? "read" : "edit" }));
    const orig = window.applyDocumentEdit;
    if (typeof orig === "function") {
      window.applyDocumentEdit = function recorded(edit, ...rest) {
        push({ a: "edit", edit: cut(edit), ...(inInput ? { fromInput: true } : {}), ...spot() });
        return orig.call(this, edit, ...rest);
      };
    }
    // Cofnij / Ponów każdą drogą (przycisk, skrót, menu) — zdarzenie z undo.js
    document.addEventListener("dwb:undo-step", (e) => push({ a: e.detail?.direction === "redo" ? "redo" : "undo", ...(inInput ? { fromInput: true } : {}) }));
  });

  const b64 = (bytes) => {
    let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  };

  // nowy dokument (otwarcie, „Nowy”) — początek sesji
  function startSession(bytes, name) {
    start = { bytes, name };
    log = [];
    push({ a: "open", name, mode: readOnlyMode ? "read" : "edit" });
  }

  async function save() {
    if (!start || !originalFileBytes) { toast(t("bugReportNoDoc"), "info"); return; }
    const end = await buildDocumentForSave();
    const report = {
      kind: "dwb-bug-report",
      version: 1,
      app: typeof APP_BUILD_VERSION !== "undefined" ? APP_BUILD_VERSION : "",
      at: new Date().toISOString(),
      ua: navigator.userAgent,
      viewport: { w: window.innerWidth, h: window.innerHeight, mobile: typeof shouldUseMobileReflow === "function" ? shouldUseMobileReflow() : false },
      mode: readOnlyMode ? "read" : "edit",
      file: start.name,
      start: b64(start.bytes),
      end: b64(end),
      actions: log,
      errors: (window.__dwbErrors || []).slice(-20),
    };
    const stamp = report.at.slice(0, 19).replace(/[:T]/g, "-");
    const blob = new Blob([JSON.stringify(report)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `zgloszenie-${stamp}.dwbbug.json`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    toast(t("bugReportSaved", { n: log.length }), "success");
  }

  // błędy strony — do zgłoszenia (ostatnie)
  window.__dwbErrors = window.__dwbErrors || [];
  window.addEventListener("error", (e) => { window.__dwbErrors.push(String(e.message || e.error || "").slice(0, 300)); });
  window.addEventListener("unhandledrejection", (e) => { window.__dwbErrors.push(String(e.reason?.message || e.reason || "").slice(0, 300)); });

  document.getElementById("bugReportMenuItem")?.addEventListener("click", () => save());
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || !e.altKey || !e.shiftKey || e.code !== "KeyB") return;
    e.preventDefault();
    save();
  });

  return { startSession, save, _log: () => log };
})();
