// Znajdź i zamień v2 — trafienia liczone z PLIKU (ten sam silnik co zamiana, docx-patch.js),
// podświetlane dokładnie (samo trafienie, nie cały akapit), opcje: wielkość liter, całe
// słowa, wyrażenia ($1 w zamianie), tylko nagłówki sekcji, też nagłówki/stopki/przypisy;
// historia ostatnich fraz.

let frMatches = []; // [{ paraIndex, occurrence, start, end, text, context, el }]
let frActiveIndex = -1;
let frOtherCount = 0; // trafienia poza treścią (nagłówki, stopki, przypisy)

const frReplaceEl = document.getElementById("frReplace");
const frRegexEl = document.getElementById("frRegex");
const frMatchCaseEl = document.getElementById("frMatchCase");
const frWholeWordEl = document.getElementById("frWholeWord");
const frOtherPartsEl = document.getElementById("frOtherParts");
const frPrevBtn = document.getElementById("frPrevBtn");
const frNextBtn = document.getElementById("frNextBtn");
const frScanBtn = document.getElementById("frScanBtn");
const frReplaceOneBtn = document.getElementById("frReplaceOneBtn");
const frReplaceAllBtn = document.getElementById("frReplaceAllBtn");
const frPreviewEl = document.getElementById("frPreview");
const frStatusEl = document.getElementById("frStatus");

const FR_HISTORY_KEY = "documents-workbench-fr-history";
const FR_HISTORY_MAX = 10;
const FR_LIST_MAX = 60;
const FR_HIGHLIGHT_MAX = 3000;
const frHighlightApi = typeof CSS !== "undefined" && CSS.highlights && typeof Highlight === "function";

function buildFrEdit() {
  return {
    op: "replace",
    find: (searchQueryEl?.value || "").trim(),
    replace: frReplaceEl?.value ?? "",
    regex: !!frRegexEl?.checked,
    matchCase: !!frMatchCaseEl?.checked,
    wholeWord: !!frWholeWordEl?.checked,
    scope: searchScopeEl?.value === "headings" ? "headings" : "all",
    otherParts: !!frOtherPartsEl?.checked,
  };
}

function validateFrFind(edit) {
  if (!edit.find) return { ok: false, err: "editErrNoFind" };
  if (!buildFindRegex(edit)) return { ok: false, err: "editErrBadRegex" };
  return { ok: true };
}

// ── historia fraz (lokalnie, podpowiedzi w polach) ────────────────────────────
function frLoadHistory() {
  try {
    const h = JSON.parse(localStorage.getItem(FR_HISTORY_KEY) || "{}");
    return { find: Array.isArray(h.find) ? h.find : [], replace: Array.isArray(h.replace) ? h.replace : [] };
  } catch (_) { return { find: [], replace: [] }; }
}
function frRemember(kind, value) {
  const v = String(value || "").trim();
  if (!v) return;
  const h = frLoadHistory();
  h[kind] = [v, ...h[kind].filter((x) => x !== v)].slice(0, FR_HISTORY_MAX);
  try { localStorage.setItem(FR_HISTORY_KEY, JSON.stringify(h)); } catch (_) { /* prywatne okno */ }
  frRenderHistory();
}
function frRenderHistory() {
  const h = frLoadHistory();
  [["frFindHistory", h.find, searchQueryEl], ["frReplaceHistory", h.replace, frReplaceEl]].forEach(([id, list, input]) => {
    let dl = document.getElementById(id);
    if (!dl) {
      dl = document.createElement("datalist");
      dl.id = id;
      document.body.appendChild(dl);
      input?.setAttribute("list", id);
    }
    dl.replaceChildren(...list.map((v) => Object.assign(document.createElement("option"), { value: v })));
  });
}

// ── podświetlenie ────────────────────────────────────────────────────────────
function clearFrHighlights() {
  docCanvasEl?.querySelectorAll(".search-hit").forEach((el) => {
    el.classList.remove("search-hit", "search-hit-active");
  });
  if (frHighlightApi) { CSS.highlights.delete("dwb-find"); CSS.highlights.delete("dwb-find-active"); }
  rootEl.classList.remove("find-precise");
}

// Zakres DOM trafienia: szukamy w tekście akapitu z podglądu tym samym wyrażeniem i bierzemy
// to samo wystąpienie. Gdy podgląd ma inny tekst niż plik (np. numer przypisu) — null,
// wtedy podświetlamy cały akapit.
function frDomRange(match, re) {
  const el = match.el;
  if (!el) return null;
  const text = el.textContent || "";
  const dom = findAllMatches(text, re);
  let hit = dom[match.occurrence];
  if (!hit || text.slice(hit.start, hit.end) !== match.text) hit = dom.find((d) => text.slice(d.start, d.end) === match.text && Math.abs(d.start - match.start) < 8);
  if (!hit) return null;
  const range = document.createRange();
  let pos = 0;
  let startSet = false;
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let n = walker.nextNode();
  while (n) {
    const len = n.textContent.length;
    if (!startSet && hit.start < pos + len) { range.setStart(n, hit.start - pos); startSet = true; }
    if (startSet && hit.end <= pos + len) { range.setEnd(n, hit.end - pos); return range; }
    pos += len;
    n = walker.nextNode();
  }
  return null;
}

function paintFrHighlights() {
  clearFrHighlights();
  if (!frMatches.length) return;
  const re = buildFindRegex(buildFrEdit());
  if (frHighlightApi && re) {
    const all = new Highlight();
    frMatches.slice(0, FR_HIGHLIGHT_MAX).forEach((m) => {
      m.range = frDomRange(m, re);
      if (m.range) all.add(m.range);
      else m.el?.classList.add("search-hit");
    });
    CSS.highlights.set("dwb-find", all);
    rootEl.classList.add("find-precise");
  } else {
    frMatches.forEach((m) => m.el?.classList.add("search-hit"));
  }
}

function focusFrMatch(index) {
  if (!frMatches.length) return;
  const i = ((index % frMatches.length) + frMatches.length) % frMatches.length;
  docCanvasEl?.querySelectorAll(".search-hit-active").forEach((el) => el.classList.remove("search-hit-active"));
  frActiveIndex = i;
  activeSearchIndex = i;
  searchMatches = frMatches.map((m) => m.el);
  const m = frMatches[i];
  if (frHighlightApi && m.range) {
    CSS.highlights.set("dwb-find-active", new Highlight(m.range));
    const r = m.range.getBoundingClientRect();
    const vr = docViewportEl.getBoundingClientRect();
    if (r.top < vr.top + 40 || r.bottom > vr.bottom - 40) {
      docViewportEl.scrollTo({ top: docViewportEl.scrollTop + r.top - vr.top - vr.height / 2, behavior: "smooth" });
    }
  } else {
    if (frHighlightApi) CSS.highlights.delete("dwb-find-active");
    m.el?.classList.add("search-hit", "search-hit-active");
    m.el?.scrollIntoView({ behavior: "smooth", block: "center" });
  }
  frPreviewEl?.querySelectorAll(".fr-preview-item").forEach((b) => b.classList.toggle("is-active", Number(b.dataset.index) === i));
  syncFrStatus();
  if (searchCountEl) searchCountEl.textContent = t("searchMatches", { count: frMatches.length });
}

// ── lista trafień z kontekstem ───────────────────────────────────────────────
function renderFrPreview() {
  if (!frPreviewEl) return;
  frPreviewEl.replaceChildren();
  if (!frMatches.length) return;
  const list = document.createElement("div");
  list.className = "fr-preview-list";
  frMatches.slice(0, FR_LIST_MAX).forEach((m, i) => {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "fr-preview-item";
    row.dataset.index = String(i);
    const snip = document.createElement("span");
    snip.className = "fr-preview-snippet";
    const a = Math.max(0, m.start - 32);
    const b = Math.min(m.context.length, m.end + 40);
    const mark = document.createElement("mark");
    mark.textContent = m.text;
    snip.append(`${a > 0 ? "…" : ""}${m.context.slice(a, m.start)}`, mark, `${m.context.slice(m.end, b)}${b < m.context.length ? "…" : ""}`);
    row.append(snip);
    row.addEventListener("click", () => focusFrMatch(i));
    list.appendChild(row);
  });
  if (frMatches.length > FR_LIST_MAX) {
    const more = document.createElement("p");
    more.className = "hint";
    more.textContent = t("frListMore", { count: frMatches.length - FR_LIST_MAX });
    list.appendChild(more);
  }
  frPreviewEl.appendChild(list);
}

function syncFrStatus() {
  if (typeof appFrame !== "undefined") appFrame.syncPanelCounts(); // licznik na pasku i przy sekcji
  if (!frStatusEl) return;
  const parts = [];
  if (frMatches.length) parts.push(t("frMatchPos", { pos: frActiveIndex >= 0 ? frActiveIndex + 1 : 0, total: frMatches.length }));
  else if (searchQueryEl?.value.trim()) parts.push(t("searchNoMatches"));
  if (frOtherCount) parts.push(t(frOtherPartsEl?.checked ? "frOtherFoundOn" : "frOtherFoundOff", { count: frOtherCount }));
  frStatusEl.textContent = parts.join(" · ");
}

async function runFindReplaceScan() {
  clearFrHighlights();
  frMatches = [];
  frActiveIndex = -1;
  frOtherCount = 0;
  searchMatches = [];
  activeSearchIndex = -1;
  if (frPreviewEl) frPreviewEl.replaceChildren();

  const edit = buildFrEdit();
  const valid = validateFrFind(edit);
  if (!valid.ok || !originalFileBytes) {
    if (searchCountEl) searchCountEl.textContent = "";
    syncFrStatus();
    if (edit.find && !valid.ok) toast(t(valid.err), "warning");
    return;
  }
  await ensureDocLibs(false);
  // Tekst wpisany w podglądzie, a jeszcze nie w pliku: te akapity liczymy z podglądu.
  // (Scalanie z plikiem przed każdym szukaniem kosztowało ~1 s przy dużym dokumencie;
  // zamiana i tak scala najpierw — applyDocumentEdit — więc numeracja się zgadza.)
  if (typeof waitInlineStructuralIdle === "function") await waitInlineStructuralIdle();
  const pending = typeof collectInlineParagraphEdits === "function" ? collectInlineParagraphEdits() : [];
  const override = new Map(pending.map((e) => [e.index, previewRunsToPlainText(e.runs).replace(/\n/g, "")]));
  const doc = await getDocumentXmlDom(originalFileBytes);
  const paras = docBodyParagraphs(docCanvasEl);
  frMatches = doc ? scanFindMatchesInDoc(doc, edit, edit.scope, override).map((m) => ({ ...m, el: paras[m.paraIndex] })) : [];
  frOtherCount = edit.scope === "all" ? await countInOtherParts(originalFileBytes, edit) : 0;
  frRemember("find", edit.find);

  if (searchCountEl) searchCountEl.textContent = frMatches.length ? t("searchMatches", { count: frMatches.length }) : t("searchNoMatches");
  paintFrHighlights();
  renderFrPreview();
  if (frMatches.length) focusFrMatch(0);
  else syncFrStatus();
}

function runDocumentSearch() {
  return runFindReplaceScan();
}

function stepFrMatch(delta) {
  if (!frMatches.length) { runFindReplaceScan(); return; }
  focusFrMatch(frActiveIndex + delta);
}

async function replaceFrOne() {
  if (!originalFileBytes) { toast(t("noFileToSave"), "warning"); return; }
  const edit = buildFrEdit();
  const valid = validateFrFind(edit);
  if (!valid.ok) { toast(t(valid.err), "warning"); return; }
  if (!frMatches.length) {
    await runFindReplaceScan();
    if (!frMatches.length) return;
  }
  const idx = Math.max(0, frActiveIndex);
  const m = frMatches[idx];
  const count = await applyDocumentEdit(edit, { target: { paraIndex: m.paraIndex, occurrence: m.occurrence } });
  if (count > 0) {
    frRemember("replace", edit.replace);
    toast(t("frReplacedOne"), "success");
    await runFindReplaceScan();
    if (frMatches.length) focusFrMatch(Math.min(idx, frMatches.length - 1)); // następne w kolejności
  } else toast(t("editNothing"), "info");
}

async function replaceFrAll() {
  if (!originalFileBytes) { toast(t("noFileToSave"), "warning"); return; }
  const edit = buildFrEdit();
  const valid = validateFrFind(edit);
  if (!valid.ok) { toast(t(valid.err), "warning"); return; }
  if (!frMatches.length && !frOtherCount) await runFindReplaceScan();
  const other = edit.otherParts ? frOtherCount : 0;
  const hits = frMatches.length + other;
  if (!hits) { toast(t("searchNoMatches"), "info"); return; }
  const paras = new Set(frMatches.map((m) => m.paraIndex)).size;
  const msg = t("frReplaceAllConfirm", { hits: frMatches.length, paras }) + (other ? ` ${t("frReplaceAllOther", { count: other })}` : "");
  if (!window.confirm(msg)) return;
  const count = await applyDocumentEdit(edit);
  if (count > 0) {
    frRemember("replace", edit.replace);
    toast(t("editApplied", { count }), "success");
  } else toast(t("editNothing"), "info");
  await runFindReplaceScan();
}

function wireFindReplaceWorkbench() {
  frScanBtn?.addEventListener("click", () => runFindReplaceScan());
  frPrevBtn?.addEventListener("click", () => stepFrMatch(-1));
  frNextBtn?.addEventListener("click", () => stepFrMatch(1));
  frReplaceOneBtn?.addEventListener("click", () => replaceFrOne());
  frReplaceAllBtn?.addEventListener("click", () => replaceFrAll());
  // zmiana opcji = od razu nowe trafienia (gdy jest czego szukać)
  [frMatchCaseEl, frWholeWordEl, frRegexEl, searchScopeEl].forEach((el) => el?.addEventListener("change", () => {
    if (searchQueryEl?.value.trim() && originalFileBytes) runFindReplaceScan();
  }));
  frOtherPartsEl?.addEventListener("change", syncFrStatus);
  searchQueryEl?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") runFindReplaceScan();
    if (e.key === "F3") { e.preventDefault(); stepFrMatch(e.shiftKey ? -1 : 1); }
  });
  frRenderHistory();
}

wireFindReplaceWorkbench();
