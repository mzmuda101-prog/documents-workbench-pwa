// track-changes.js — „Śledź zmiany” jak w Wordzie (Recenzja → Śledź zmiany, Ctrl/⌘+Shift+E).
//
// Model: po włączeniu zapamiętujemy plik „przed” (baseline). Pisze się normalnie — w aplikacji
// stan zostaje czysty (każdy akapit edytowalny), a zmienione akapity mają pasek na marginesie
// jak „Proste znaczniki” w Wordzie. Poprawki (w:ins / w:del) powstają z różnicy (docx-redline.js):
//   - przy zapisie do pliku — na dysk idzie dokument z poprawkami i w:trackRevisions (Word po
//     otwarciu dalej śledzi), a w aplikacji zostaje czysty stan i śledzenie trwa dalej,
//   - przy wyłączeniu albo „Pokaż poprawki w dokumencie” — poprawki trafiają do dokumentu w
//     aplikacji (jeden krok Cofnij); takie akapity są do przejrzenia w Recenzji (Akceptuj/Odrzuć).
// Akceptuj/Odrzuć w trakcie śledzenia dotyczy też pliku „przed” — nie tworzy nowych poprawek.
// Plik zapisany z w:trackRevisions (Word ze śledzeniem) — po otwarciu śledzenie włącza się samo.

const dwbTrack = (() => {
  const AUTHOR_KEY = "dwb.authorName";
  const state = { on: false, baseline: null, author: "", baseKeys: null };
  const toggleEl = () => document.getElementById("rvTrack");
  const infoEl = () => document.getElementById("rvTrackInfo");
  const showBtn = () => document.getElementById("rvTrackShowBtn");
  const statusEl = () => document.getElementById("statusTrack");

  function authorName() {
    let name = "";
    try { name = localStorage.getItem(AUTHOR_KEY) || ""; } catch (_) {}
    if (!name) {
      name = (window.prompt(t("trackAuthorAsk"), "") || "").trim();
      if (name) try { localStorage.setItem(AUTHOR_KEY, name); } catch (_) {}
    }
    return name || t("trackAuthorDefault");
  }

  function sync() {
    const tg = toggleEl();
    if (tg) tg.checked = state.on;
    const st = statusEl();
    if (st) st.hidden = !state.on;
    const btn = showBtn();
    if (btn) btn.hidden = !state.on;
    document.documentElement.classList.toggle("track-on", state.on);
    if (!state.on) {
      if (infoEl()) infoEl().textContent = t("trackOffHint");
      clearMarks();
    } else scheduleMarks();
  }

  async function start({ silent = false } = {}) {
    if (!originalFileBytes || state.on) return false;
    await mergeInlineEditsIntoBytes();
    if (pendingDocEdits?.length) {
      const { bytes } = await buildPatchedDocx(originalFileBytes, pendingDocEdits);
      pendingDocEdits = [];
      originalFileBytes = bytes;
    }
    state.author = authorName();
    state.baseline = originalFileBytes;
    state.baseKeys = await baselineKeys(state.baseline);
    state.on = true;
    sync();
    if (!silent) toast(t("trackOnToast", { name: state.author }), "info");
    return true;
  }

  // Poprawki do dokumentu w aplikacji (jeden krok Cofnij). keepOn — śledzenie trwa dalej od tego stanu.
  async function bake({ keepOn = false } = {}) {
    if (!state.on || !state.baseline) return 0;
    const baseline = state.baseline;
    state.on = false; // na czas wpisywania — zapis nie liczy różnicy drugi raz
    const n = await applyDocumentEdit({ op: "redline", baseline, author: state.author, trackOn: keepOn });
    state.baseline = null;
    state.baseKeys = null;
    if (keepOn) await start({ silent: true });
    else sync();
    return n;
  }

  async function stop() {
    if (!state.on) return 0;
    const n = await bake({ keepOn: false });
    toast(n ? t("trackOffToast", { n }) : t("trackOffNone"), "info");
    return n;
  }

  function toggle() { return state.on ? stop() : start(); }

  // Do zapisu na dysk: czysty stan z aplikacji → plik z poprawkami (gdy śledzenie włączone).
  async function bytesForDisk(clean) {
    if (!state.on || !state.baseline || !clean) return clean;
    const { bytes } = await dwbRedline(state.baseline, clean, { author: state.author, trackOn: true });
    return bytes;
  }

  // Akceptuj/Odrzuć w trakcie śledzenia — to samo na pliku „przed”
  async function mirrorRevisions(edit) {
    if (!state.on || !state.baseline || edit?.op !== "revisions") return;
    const { bytes } = await buildPatchedDocx(state.baseline, [edit]);
    state.baseline = bytes;
    state.baseKeys = await baselineKeys(bytes);
  }

  // ── „Proste znaczniki”: pasek przy zmienionych akapitach ──────────────────
  // Tekst akapitów „jak po przyjęciu” — tak rysuje go podgląd: wstawione (w:ins) się liczy,
  // usunięte (w:delText) nie; łamanie wiersza = „\n”, tabulator = „\t” (jak klucze z podglądu).
  function acceptedText(p) {
    let out = "";
    const walk = (el) => {
      for (const n of Array.from(el.childNodes)) {
        if (n.nodeType !== 1 || n.namespaceURI !== W_NS) continue;
        if (n.localName === "pPr" || n.localName === "rPr" || n.localName === "del" || n.localName === "moveFrom") continue;
        if (n.localName === "t") out += n.textContent;
        else if (n.localName === "tab") out += "\t";
        else if (n.localName === "br") out += "\n";
        else walk(n);
      }
    };
    walk(p);
    return out;
  }
  async function baselineKeys(bytes) {
    const doc = await getDocumentXmlDom(bytes).catch(() => null);
    return doc ? collectParagraphElements(doc.documentElement, "all").map(acceptedText) : [];
  }
  function clearMarks() {
    docCanvasEl?.querySelectorAll(".dwb-tracked, .dwb-tracked-gap").forEach((p) => p.classList.remove("dwb-tracked", "dwb-tracked-gap"));
    docCanvasEl?.querySelectorAll(".dwb-track-bar").forEach((b) => b.remove());
  }
  // Pasek na marginesie strony obok akapitu (jak w Wordzie): osobny element w kartce — cień ani
  // pseudo-element akapitu tu nie pasują (pseudo-elementy mają listy i znaki ¶, cień zanika przy
  // jednowierszowym akapicie). Nieedytowalny, niewidoczny dla czytników ekranu, bez druku.
  function placeBar(p) {
    const sec = p.closest("section.docx") || p.parentElement;
    if (!sec) return;
    if (getComputedStyle(sec).position === "static") sec.style.position = "relative";
    const sr = sec.getBoundingClientRect(), pr = p.getBoundingClientRect();
    const k = sec.offsetWidth ? sr.width / sec.offsetWidth : 1; // powiększenie widoku (transform)
    const bar = document.createElement("div");
    bar.className = "dwb-track-bar";
    bar.contentEditable = "false";
    bar.setAttribute("aria-hidden", "true");
    bar.style.top = `${(pr.top - sr.top) / k}px`;
    bar.style.height = `${Math.max(4, pr.height / k)}px`;
    bar.style.left = `${Math.max(2, (pr.left - sr.left) / k - 14)}px`;
    sec.appendChild(bar);
  }
  let markTimer = null;
  function scheduleMarks() {
    clearTimeout(markTimer);
    markTimer = setTimeout(paintMarks, 350);
  }
  function paintMarks() {
    if (!state.on || !state.baseKeys) return;
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    const paras = host ? collectPreviewParagraphElements(host) : [];
    if ((paras.length + 1) * (state.baseKeys.length + 1) > 6_000_000) return; // bardzo długi dokument — bez pasków
    const keys = paras.map((p) => previewRunsToPlainText(extractRunsFromPreviewParagraph(p)));
    const rows = rlAlign(state.baseKeys, keys);
    clearMarks();
    let changed = 0, gapNext = false;
    rows.forEach((r) => {
      if (r.kind === "removed") { gapNext = true; changed++; return; }
      const p = r.n != null ? paras[r.n] : null;
      if (!p) return;
      if (r.kind === "changed" || r.kind === "added") { p.classList.add("dwb-tracked"); changed++; }
      if (gapNext) { p.classList.add("dwb-tracked-gap"); gapNext = false; }
    });
    docCanvasEl?.querySelectorAll("p.dwb-tracked").forEach(placeBar);
    if (infoEl()) infoEl().textContent = t("trackOnInfo", { name: state.author, n: changed });
  }

  // Po otwarciu pliku: stan z karty (restore) albo — Word zapisał plik ze śledzeniem — włącz.
  async function onDocumentLoaded(restore) {
    state.on = false; state.baseline = null; state.baseKeys = null;
    if (restore?.baseline) {
      // od razu (zapis tuż po przełączeniu karty też ma poprawki); klucze pasków — w tle
      state.baseline = restore.baseline;
      state.author = restore.author || t("trackAuthorDefault");
      state.on = true;
      sync();
      const keys = await baselineKeys(state.baseline);
      if (state.baseline === restore.baseline) { state.baseKeys = keys; scheduleMarks(); }
      return;
    }
    sync();
    if (!originalFileBytes) return;
    try {
      const zip = await JSZip.loadAsync(originalFileBytes);
      const f = zip.file("word/settings.xml");
      const s = f ? new DOMParser().parseFromString(await f.async("string"), "application/xml") : null;
      const el = s && Array.from(s.documentElement.childNodes).find((n) => n.localName === "trackRevisions");
      const val = el && (el.getAttributeNS(W_NS, "val") || el.getAttribute("w:val") || "");
      if (el && !/^(0|false|off)$/i.test(val)) {
        await start({ silent: true });
        toast(t("trackAutoOn"), "info");
      }
    } catch (_) {}
  }
  // Odłożenie karty (open-docs.js): plik „przed” i autor jadą razem z kartą.
  function snapshot() { return state.on && state.baseline ? { baseline: state.baseline, author: state.author } : null; }

  function init() {
    toggleEl()?.addEventListener("change", (e) => { if (e.target.checked !== state.on) toggle(); });
    showBtn()?.addEventListener("click", () => bake({ keepOn: true }));
    statusEl()?.addEventListener("click", () => {
      if (typeof setSidebarOpen === "function") setSidebarOpen(true);
      const panel = document.getElementById("panel-review");
      if (panel) { panel.open = true; panel.scrollIntoView({ block: "start", behavior: "smooth" }); }
    });
    // Ctrl/⌘+Shift+E — jak w Wordzie
    document.addEventListener("keydown", (e) => {
      if (!(e.ctrlKey || e.metaKey) || !e.shiftKey || e.altKey || e.code !== "KeyE" || !originalFileBytes || readOnlyMode) return;
      e.preventDefault();
      toggle();
    });
    docCanvasEl?.addEventListener("input", () => { if (state.on) scheduleMarks(); });
    window.addEventListener("resize", () => { if (state.on) scheduleMarks(); });
    docCanvasEl?.addEventListener("dwb-zoom", () => { if (state.on) scheduleMarks(); });
    // przerysowanie dokumentu (Enter, operacje z paska) — paski od nowa
    if (docCanvasEl && typeof MutationObserver === "function") new MutationObserver(() => { if (state.on) scheduleMarks(); }).observe(docCanvasEl, { childList: true });
    sync();
  }

  return { init, start, stop, toggle, bake, bytesForDisk, mirrorRevisions, onDocumentLoaded, snapshot, scheduleMarks, isOn: () => state.on, _state: state };
})();

document.addEventListener("DOMContentLoaded", () => dwbTrack.init());
