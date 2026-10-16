// doc-ask.js — „Zapytaj dokument” (panel Znajdź sens): pytanie zwykłymi słowami → odpowiedź
// asystenta AI zbudowana WYŁĄCZNIE z fragmentów tego dokumentu, z odnośnikami [1] [2] do akapitów.
//
// 1. Fragmenty: gdy model „Znajdź sens” działa — wyszukiwanie znaczeniowe (semantic-search.js),
//    inaczej dopasowanie słów z polskimi odmianami (rdzeń = początek wyrazu, waga rzadkości).
// 2. Odpowiedź: local-ai.js. Każde zdanie ma wskazać fragment; odpowiedź bez poprawnego
//    odnośnika jest oznaczona jako niepewna. Bez asystenta — same najlepsze fragmenty.

const docAskInputEl = document.getElementById("docAskQuery");
const docAskBtn = document.getElementById("docAskBtn");
const docAskAnswerEl = document.getElementById("docAskAnswer");
const DOC_ASK_TOP = 6;
const DOC_ASK_FRAGMENT_CHARS = 700;
const DOC_ASK_STOP = new Set("oraz albo lub czy jak jaki jaka jakie który która które się jest są był była było będzie przez przy dla od do na po za ze we nie tak też to ten ta te tym tego jego jej ich ile kiedy gdzie co kto czym czego dlaczego jakim jaką może można musi należy the and for are was what when where which who how does".split(" "));

function docAskStem(word) {
  const w = word.toLocaleLowerCase();
  return w.length > 5 ? w.slice(0, Math.max(5, w.length - 3)) : w;
}
function docAskTerms(text) {
  return (String(text).toLocaleLowerCase().match(/[\p{L}\p{N}]{3,}/gu) || []).filter((w) => !DOC_ASK_STOP.has(w)).map(docAskStem);
}

// Dopasowanie słów (BM25 w uproszczeniu) — działa zawsze, bez żadnego modelu.
function docAskLexical(passages, query, k = DOC_ASK_TOP) {
  const q = [...new Set(docAskTerms(query))];
  if (!q.length) return [];
  const docs = passages.map((p) => docAskTerms(p.text));
  const avg = docs.reduce((s, d) => s + d.length, 0) / (docs.length || 1) || 1;
  const df = new Map();
  docs.forEach((d) => new Set(d).forEach((term) => df.set(term, (df.get(term) || 0) + 1)));
  const N = docs.length;
  return passages.map((p, i) => {
    const tf = new Map();
    docs[i].forEach((term) => tf.set(term, (tf.get(term) || 0) + 1));
    let score = 0;
    for (const term of q) {
      const f = tf.get(term);
      if (!f) continue;
      const idf = Math.log(1 + (N - (df.get(term) || 0) + 0.5) / ((df.get(term) || 0) + 0.5));
      score += idf * ((f * 2.2) / (f + 1.2 * (0.25 + 0.75 * (docs[i].length / avg))));
    }
    return { ...p, score };
  }).filter((r) => r.score > 0).sort((a, b) => b.score - a.score).slice(0, k);
}

async function docAskRetrieve(query) {
  const passages = typeof buildSemanticPassages === "function" ? buildSemanticPassages() : [];
  if (!passages.length) return { results: [], mode: "none" };
  if (typeof semanticRetrieve === "function") {
    const semantic = await semanticRetrieve(query, DOC_ASK_TOP).catch(() => null);
    if (semantic?.length) return { results: semantic, mode: "semantic" };
  }
  return { results: docAskLexical(passages, query), mode: "lexical" };
}

const DOC_ASK_SYSTEM =
  "Odpowiadasz na pytanie wyłącznie na podstawie ponumerowanych fragmentów dokumentu. Odpowiadasz krótko, " +
  "w języku pytania. Po każdym zdaniu podajesz numer fragmentu, z którego pochodzi informacja, w nawiasie " +
  "kwadratowym, np. [2]. Nie dodajesz wiedzy spoza fragmentów. Jeśli fragmenty nie zawierają odpowiedzi, " +
  "piszesz tylko: Brak w dokumencie.";

function docAskPrompt(question, results) {
  const blocks = results.map((r, i) => `[${i + 1}] ${String(r.text).slice(0, DOC_ASK_FRAGMENT_CHARS)}`);
  return `Fragmenty dokumentu:\n${blocks.join("\n\n")}\n\nPytanie: ${question}`;
}

// Odnośniki [n] w odpowiedzi → numery fragmentów (tylko istniejące)
function docAskCitations(answer, count) {
  const out = new Set();
  for (const m of String(answer).matchAll(/\[(\d{1,2})(?:\s*[,–-]\s*(\d{1,2}))?\]/g)) {
    [m[1], m[2]].filter(Boolean).map(Number).forEach((n) => { if (n >= 1 && n <= count) out.add(n); });
  }
  return [...out];
}

function docAskJump(result) {
  if (typeof jumpToSemanticResult === "function") jumpToSemanticResult(result);
  else result.item?.el?.scrollIntoView({ behavior: "smooth", block: "center" });
}

function docAskRenderAnswer(text, results, { ai = true, pending = false } = {}) {
  docAskAnswerEl.replaceChildren();
  docAskAnswerEl.hidden = false;
  if (ai) {
    const head = document.createElement("div");
    head.className = "doc-ask-head";
    head.append(Object.assign(document.createElement("span"), { className: "ai-badge", textContent: t("aiBadge") }));
    docAskAnswerEl.append(head);
    const p = document.createElement("p");
    p.className = "doc-ask-text";
    // [n] → przycisk skoku do fragmentu
    let last = 0;
    for (const m of String(text).matchAll(/\[(\d{1,2})\]/g)) {
      p.append(document.createTextNode(text.slice(last, m.index)));
      const n = Number(m[1]);
      if (results[n - 1]) {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "doc-ask-cite";
        b.textContent = n;
        b.setAttribute("aria-label", t("askCiteAria", { n }));
        b.addEventListener("click", () => docAskJump(results[n - 1]));
        p.append(b);
      } else p.append(document.createTextNode(m[0]));
      last = m.index + m[0].length;
    }
    p.append(document.createTextNode(text.slice(last)));
    docAskAnswerEl.append(p);
    if (!pending) {
      const cites = docAskCitations(text, results.length);
      const none = /brak w dokumencie|not in the document/i.test(text);
      if (!cites.length && !none) docAskAnswerEl.append(Object.assign(document.createElement("p"), { className: "doc-ask-warn", textContent: t("askNoCitation") }));
    }
  } else {
    docAskAnswerEl.append(Object.assign(document.createElement("p"), { className: "hint", textContent: text }));
  }
  if (!results.length || pending) return;
  const src = document.createElement("div");
  src.className = "doc-ask-sources";
  src.append(Object.assign(document.createElement("div"), { className: "doc-ask-sources-title", textContent: t("askSources") }));
  results.forEach((r, i) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "semantic-result doc-ask-source";
    b.append(
      Object.assign(document.createElement("span"), { className: "semantic-result-score", textContent: `[${i + 1}]` }),
      Object.assign(document.createElement("span"), { className: "semantic-result-text", textContent: r.text, title: r.text })
    );
    b.addEventListener("click", () => docAskJump(r));
    src.append(b);
  });
  docAskAnswerEl.append(src);
}

async function docAskRun() {
  const question = String(docAskInputEl?.value || "").replace(/\s+/g, " ").trim();
  if (!question) { docAskInputEl?.focus(); return; }
  if (!originalFileBytes) { toast(t("noFileToSave"), "warning"); return; }
  docAskBtn.disabled = true;
  try {
    const { results } = await docAskRetrieve(question);
    if (!results.length) {
      docAskRenderAnswer(t("askNothing"), [], { ai: false });
      return;
    }
    const useAi = typeof LocalAI !== "undefined" && (await LocalAI.ensure({ reason: t("askAiReason") }));
    if (!useAi) {
      docAskRenderAnswer(t("askNoAi"), results, { ai: false });
      return;
    }
    docAskRenderAnswer("…", results, { pending: true });
    const answer = await LocalAI.generate({
      system: DOC_ASK_SYSTEM,
      user: docAskPrompt(question, results),
      maxTokens: 260,
      onPartial: (s) => docAskRenderAnswer(s, results, { pending: true }),
    });
    docAskRenderAnswer(answer.trim() || t("askEmpty"), results);
  } catch (err) {
    docAskRenderAnswer(t("aiFailed", { msg: err.message || err }), [], { ai: false });
  } finally {
    docAskBtn.disabled = false;
  }
}

docAskBtn?.addEventListener("click", docAskRun);
docAskInputEl?.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); docAskRun(); } });
