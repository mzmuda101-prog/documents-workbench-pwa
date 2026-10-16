// Grammar panel — scan, grouped suggestions, apply one / rule / all.

let grammarScan = null;
let grammarActiveHitId = null;

const grammarLangEl = document.getElementById("grammarLang");
const grammarNbspEl = document.getElementById("grammarNbsp");
const grammarScanBtn = document.getElementById("grammarScanBtn");
const grammarApplyAllBtn = document.getElementById("grammarApplyAllBtn");
const grammarClearBtn = document.getElementById("grammarClearBtn");
const grammarStatusEl = document.getElementById("grammarStatus");
const grammarSuggestionsEl = document.getElementById("grammarSuggestions");

const grammarSpellEl = document.getElementById("grammarSpell");
const GRAMMAR_SPELL_KEY = "dwb.grammarSpell";
const GRAMMAR_SPELL_SUGGEST_BATCH = 40;
try { if (grammarSpellEl && localStorage.getItem(GRAMMAR_SPELL_KEY) === "0") grammarSpellEl.checked = false; } catch (_) {}

function grammarScanOpts() {
  return {
    lang: grammarLangEl?.value || "pl",
    nbspPl: !!grammarNbspEl?.checked,
  };
}

// Podgląd poprawki: kontekst + skreślone (przed) → wstawione (po). Spacje w zmianie są widoczne
// („·”) — inaczej poprawki podwójnej spacji czy spacji przed przecinkiem nie dało się zobaczyć.
function grammarVisibleSpaces(s) { return escapeGrammarHtml(s).replace(/ /g, '<span class="gr-space">·</span>'); }
function grammarPreviewHtml(hit) {
  if (hit.ctxBefore == null) return escapeGrammarHtml(hit.snippet);
  const del = hit.before ? `<del class="gr-del">${grammarVisibleSpaces(hit.before)}</del>` : "";
  const ins = hit.after ? `<ins class="gr-ins">${grammarVisibleSpaces(hit.after)}</ins>` : "";
  return `${escapeGrammarHtml(hit.ctxBefore)}${del}${ins}${escapeGrammarHtml(hit.ctxAfter)}`;
}

function escapeGrammarHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function ruleLabel(ruleId) {
  const key = `grammarRule${ruleId.split("-").map((p) => p.charAt(0).toUpperCase() + p.slice(1)).join("")}`;
  return t(key);
}

function syncGrammarStatus() {
  if (typeof appFrame !== "undefined") appFrame.syncPanelCounts();
  if (!grammarStatusEl) return;
  const spelling = grammarScan?.spelling?.length || 0;
  const spellNote = spelling ? ` ${t("spellFound", { n: spelling })}` : grammarScan?.spellError ? ` ${t("spellFailed")}` : "";
  if (!grammarScan?.hits?.length) {
    grammarStatusEl.textContent = grammarScan ? (spelling ? spellNote.trim() : t("grammarNoHits") + spellNote) : "";
    return;
  }
  const ruleCount = Object.keys(grammarScan.byRule || {}).length;
  grammarStatusEl.textContent = t("grammarFound", {
    hits: grammarScan.hits.length,
    rules: ruleCount,
  }) + spellNote;
}

// ── Ortografia (spell-check.js) — osobna grupa: każdy błędny wyraz raz, z podpowiedziami.
// „Zastosuj wszystkie” jej NIE obejmuje — wybór poprawnego wyrazu należy do człowieka.
function renderSpellingGroup() {
  const list = grammarScan?.spelling;
  if (!list?.length) return null;
  const details = document.createElement("details");
  details.className = "grammar-rule-group spell-group";
  details.open = true;
  const summary = document.createElement("summary");
  summary.className = "grammar-rule-summary";
  summary.innerHTML = `<span class="grammar-rule-name">${escapeGrammarHtml(t("spellGroup"))}</span><span class="grammar-rule-count">(${list.length})</span>`;
  details.appendChild(summary);
  const box = document.createElement("div");
  box.className = "fr-preview-list grammar-hit-list";
  list.forEach((entry) => {
    const row = document.createElement("div");
    row.className = "spell-row";
    const word = document.createElement("button");
    word.type = "button";
    word.className = "spell-word";
    word.textContent = entry.word;
    if (entry.count > 1) word.append(Object.assign(document.createElement("span"), { className: "spell-count", textContent: ` ×${entry.count}` }));
    word.setAttribute("aria-label", t("spellJump", { word: entry.word }));
    word.addEventListener("click", () => scrollToGrammarParagraph(entry.occurrences[0].paraIndex));
    row.appendChild(word);
    const actions = document.createElement("div");
    actions.className = "spell-actions";
    if (entry.suggestions == null) {
      actions.append(Object.assign(document.createElement("span"), { className: "hint", textContent: "…" }));
    } else if (!entry.suggestions.length) {
      actions.append(Object.assign(document.createElement("span"), { className: "hint", textContent: t("spellNoSuggestions") }));
    }
    (entry.suggestions || []).slice(0, 3).forEach((sug) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "btn spell-suggestion";
      b.textContent = dwbSpell.matchCase(entry.word, sug);
      b.addEventListener("click", () => applySpellingFix(entry, sug));
      actions.appendChild(b);
    });
    const ignoreBtn = document.createElement("button");
    ignoreBtn.type = "button";
    ignoreBtn.className = "btn spell-ignore";
    ignoreBtn.textContent = t("spellIgnore");
    ignoreBtn.addEventListener("click", () => {
      dwbSpell.ignore(entry.word);
      grammarScan.spelling = grammarScan.spelling.filter((x) => x !== entry);
      renderGrammarSuggestions();
      syncGrammarStatus();
    });
    const addBtn = document.createElement("button");
    addBtn.type = "button";
    addBtn.className = "btn spell-add";
    addBtn.textContent = t("spellAdd");
    addBtn.addEventListener("click", () => {
      dwbSpell.addToDictionary(entry.word);
      grammarScan.spelling = grammarScan.spelling.filter((x) => x !== entry);
      renderGrammarSuggestions();
      syncGrammarStatus();
      toast(t("spellAdded", { word: entry.word }), "success");
    });
    actions.append(ignoreBtn, addBtn);
    row.appendChild(actions);
    box.appendChild(row);
  });
  details.appendChild(box);
  return details;
}

// Podpowiedzi partiami — pierwsze wyrazy od razu, reszta w tle (bez blokowania panelu).
async function fillSpellingSuggestions(scan) {
  const lang = grammarScanOpts().lang;
  while (grammarScan === scan) {
    const batch = (scan.spelling || []).filter((e) => e.suggestions == null).slice(0, GRAMMAR_SPELL_SUGGEST_BATCH);
    if (!batch.length) return;
    let got = {};
    try { got = await dwbSpell.suggest(batch.map((e) => e.word), lang); } catch (_) { got = {}; }
    batch.forEach((e) => { e.suggestions = got[e.word] || []; });
    if (grammarScan === scan) renderGrammarSuggestions();
  }
}

async function applySpellingFix(entry, suggestion) {
  if (!originalFileBytes) return;
  if (typeof mergeInlineEditsIntoBytes === "function" && (await mergeInlineEditsIntoBytes()) > 0) {
    toast(t("grammarRescanned"), "info"); // tekst zmienił się od skanu — miejsca wyrazów są nieaktualne
    await runGrammarScan();
    return;
  }
  const texts = await extractParagraphTextsFromDocx(originalFileBytes);
  // tylko miejsca, gdzie dalej stoi ten sam wyraz (ochrona przed przesunięciem tekstu)
  const live = entry.occurrences.filter((o) => texts[o.paraIndex]?.slice(o.start, o.end) === entry.word);
  const items = dwbSpell.replaceIn(texts, live, suggestion);
  if (!items.length) { toast(t("editNothing"), "info"); return; }
  const count = await applyDocumentEdit({ op: "paragraphBatch", items });
  if (count > 0) toast(t("spellReplaced", { word: entry.word, fix: dwbSpell.matchCase(entry.word, suggestion), n: live.length }), "success");
  await runGrammarScan();
}

function renderGrammarSuggestions() {
  if (!grammarSuggestionsEl) return;
  grammarSuggestionsEl.replaceChildren();
  const spellGroup = renderSpellingGroup();
  if (spellGroup) grammarSuggestionsEl.appendChild(spellGroup);
  if (!grammarScan?.hits?.length) return;

  Object.entries(grammarScan.byRule).forEach(([ruleId, hits]) => {
    const details = document.createElement("details");
    details.className = "grammar-rule-group";
    details.open = hits.length <= 4;

    const summary = document.createElement("summary");
    summary.className = "grammar-rule-summary";
    summary.innerHTML = `<span class="grammar-rule-name">${escapeGrammarHtml(ruleLabel(ruleId))}</span><span class="grammar-rule-count">(${hits.length})</span>`;
    details.appendChild(summary);

    const applyRuleBtn = document.createElement("button");
    applyRuleBtn.type = "button";
    applyRuleBtn.className = "btn grammar-apply-rule";
    applyRuleBtn.textContent = t("grammarApplyRule");
    applyRuleBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      applyGrammarByRule(ruleId);
    });
    summary.appendChild(applyRuleBtn);

    const list = document.createElement("div");
    list.className = "fr-preview-list grammar-hit-list";

    hits.forEach((hit) => {
      const row = document.createElement("div");
      row.className = "grammar-hit-row";
      if (hit.id === grammarActiveHitId) row.classList.add("is-active");

      const snippet = document.createElement("button");
      snippet.type = "button";
      snippet.className = "fr-preview-item grammar-hit-item";
      snippet.innerHTML = `<span class="fr-preview-snippet">${grammarPreviewHtml(hit)}</span>`;
      snippet.setAttribute("aria-label", hit.snippet);
      snippet.addEventListener("click", () => {
        grammarActiveHitId = hit.id;
        renderGrammarSuggestions();
        scrollToGrammarParagraph(hit.paraIndex);
      });
      row.appendChild(snippet);

      const applyBtn = document.createElement("button");
      applyBtn.type = "button";
      applyBtn.className = "btn grammar-apply-one";
      applyBtn.textContent = t("grammarApplyOne");
      applyBtn.addEventListener("click", () => applyGrammarOne(hit));
      row.appendChild(applyBtn);

      list.appendChild(row);
    });

    details.appendChild(list);
    grammarSuggestionsEl.appendChild(details);
  });
}

function scrollToGrammarParagraph(paraIndex) {
  if (!docCanvasEl) return;
  const el = docBodyParagraphs(docCanvasEl)[paraIndex]; // numer akapitu z pliku = akapit treści
  if (!el) return;
  docCanvasEl.querySelectorAll(".search-hit-active").forEach((n) => n.classList.remove("search-hit-active"));
  el.classList.add("search-hit", "search-hit-active");
  el.scrollIntoView({ behavior: "smooth", block: "center" });
}

function clearGrammarState() {
  grammarScan = null;
  grammarActiveHitId = null;
  if (grammarSuggestionsEl) grammarSuggestionsEl.replaceChildren();
  docCanvasEl?.querySelectorAll(".search-hit, .search-hit-active").forEach((el) => {
    el.classList.remove("search-hit", "search-hit-active");
  });
  syncGrammarStatus();
}

async function runGrammarScan() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  grammarActiveHitId = null;
  // tekst wpisany w podglądzie musi być w pliku — inaczej „Popraw” liczyłoby poprawkę ze
  // starej wersji akapitu i nadpisało nią świeżo dopisany tekst
  if (typeof mergeInlineEditsIntoBytes === "function") await mergeInlineEditsIntoBytes();
  const scan = await scanDocument(originalFileBytes, grammarScanOpts());
  if (grammarSpellEl?.checked && typeof dwbSpell !== "undefined") {
    try {
      const texts = await extractParagraphTextsFromDocx(originalFileBytes);
      scan.spelling = await dwbSpell.scan(texts, grammarScanOpts().lang);
    } catch (err) {
      scan.spelling = [];
      scan.spellError = true;
      if (typeof log === "function") log(`Ortografia: ${err.message || err}`, "error");
    }
  }
  grammarScan = scan;
  renderGrammarSuggestions();
  syncGrammarStatus();
  if (scan.spelling?.length) fillSpellingSuggestions(scan);
  if (!grammarScan.hits.length && scan.spelling?.length) scrollToGrammarParagraph(scan.spelling[0].occurrences[0].paraIndex);
  if (grammarScan.hits.length) {
    grammarActiveHitId = grammarScan.hits[0].id;
    renderGrammarSuggestions();
    scrollToGrammarParagraph(grammarScan.hits[0].paraIndex);
  }
}

async function applyGrammarOne(hit) {
  if (!originalFileBytes || !hit) return;
  if (typeof mergeInlineEditsIntoBytes === "function" && (await mergeInlineEditsIntoBytes()) > 0) {
    toast(t("grammarRescanned"), "info"); // akapit zmienił się od skanu — nie nadpisujemy go starą wersją
    await runGrammarScan();
    return;
  }
  const count = await applyDocumentEdit({
    op: "paragraphBatch",
    items: [{ index: hit.paraIndex, text: hit.fixedParagraph }],
  });
  if (count > 0) {
    toast(t("grammarAppliedOne"), "success");
    await runGrammarScan();
  } else toast(t("editNothing"), "info");
}

async function applyGrammarBatch(filterRuleId = null) {
  if (!originalFileBytes || !grammarScan?.hits?.length) {
    toast(t("grammarNoHits"), "info");
    return;
  }
  const opts = grammarScanOpts();
  await ensureDocLibs(false);
  if (typeof mergeInlineEditsIntoBytes === "function" && (await mergeInlineEditsIntoBytes()) > 0) {
    await runGrammarScan(); // coś dopisano od skanu — poprawki od nowa, na aktualnym tekście
    if (!grammarScan?.hits?.length) return;
  }
  const texts = await extractParagraphTextsFromDocx(originalFileBytes);
  const hits = filterRuleId
    ? grammarScan.hits.filter((h) => h.ruleId === filterRuleId)
    : grammarScan.hits;
  if (!hits.length) {
    toast(t("grammarNoHits"), "info");
    return;
  }
  const items = buildGrammarBatchItems(texts, hits, opts, filterRuleId);
  if (!items.length) {
    toast(t("editNothing"), "info");
    return;
  }
  if (!filterRuleId) {
    const msg = t("grammarApplyAllConfirm", { count: items.length });
    if (!window.confirm(msg)) return;
  }
  const count = await applyDocumentEdit({ op: "paragraphBatch", items });
  if (count > 0) toast(t("editApplied", { count }), "success");
  else toast(t("editNothing"), "info");
  await runGrammarScan();
}

async function applyGrammarByRule(ruleId) {
  await applyGrammarBatch(ruleId);
}

async function applyGrammarAll() {
  await applyGrammarBatch(null);
}

function wireGrammarPanel() {
  if (grammarScanBtn) grammarScanBtn.addEventListener("click", () => runGrammarScan());
  if (grammarApplyAllBtn) grammarApplyAllBtn.addEventListener("click", () => applyGrammarAll());
  if (grammarClearBtn) grammarClearBtn.addEventListener("click", () => clearGrammarState());
  if (grammarLangEl) grammarLangEl.addEventListener("change", () => { if (grammarScan) runGrammarScan(); });
  if (grammarNbspEl) grammarNbspEl.addEventListener("change", () => { if (grammarScan) runGrammarScan(); });
  if (grammarSpellEl) grammarSpellEl.addEventListener("change", () => {
    try { localStorage.setItem(GRAMMAR_SPELL_KEY, grammarSpellEl.checked ? "1" : "0"); } catch (_) {}
    if (grammarScan) runGrammarScan();
  });
}

wireGrammarPanel();
