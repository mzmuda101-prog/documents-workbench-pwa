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

// Polecenie dostrojone pod MAŁE modele (docs/llm-bench/prompts.html, 2026-10-17: Qwen3.5 0.8B
// 2/3 → 3/3): krótka instrukcja, jeden przykład i podpowiedź z kodu aplikacji — numer
// najtrafniejszego zdania (docAskBestSentence). Model tylko je przeredagowuje w odpowiedź.
const DOC_ASK_SYSTEM = "Odpowiadasz jednym zdaniem po polsku, tylko na podstawie fragmentów. Na końcu numer fragmentu w nawiasie, np. [2]. Jeśli fragmenty nie odpowiadają na pytanie, piszesz: Brak w dokumencie.";
const DOC_ASK_SHOT = [
  "Fragmenty:\n[1] Umowa obowiązuje od 1 marca 2026 r.\n[2] Opłata wynosi 200 zł miesięcznie.\n\nPytanie: Ile wynosi opłata?\nNajtrafniejsze zdanie: [2] „Opłata wynosi 200 zł miesięcznie.”",
  "Opłata wynosi 200 zł miesięcznie [2].",
];

// Krótki kontekst = szybka odpowiedź na telefonie: 3 fragmenty, przycięte (numery jak na liście).
const DOC_ASK_AI_FRAGMENTS = 3;
const DOC_ASK_AI_CHARS = 450;
const DOC_ASK_SYSTEM_WITH_HINT = "Odpowiadasz jednym zdaniem po polsku, na podstawie najtrafniejszego zdania i fragmentów. Na końcu numer fragmentu w nawiasie, np. [2].";

function docAskPrompt(question, results, best = null) {
  const bestIdx = best ? results.indexOf(best.result) : -1;
  const order = [...new Set([bestIdx, ...results.keys()].filter((i) => i >= 0))].slice(0, DOC_ASK_AI_FRAGMENTS);
  const blocks = order.map((i) => {
    const text = String(results[i].text);
    // z długiego akapitu — okolica najtrafniejszego zdania
    const at = i === bestIdx ? Math.max(0, text.indexOf(best.sentence) - 80) : 0;
    return `[${i + 1}] ${text.slice(at, at + DOC_ASK_AI_CHARS)}`;
  });
  // treść zdania, nie sam numer: przy długich akapitach mały model sam go nie znajdował
  // i odpowiadał „Brak w dokumencie” (sprawdzone na headings-sample, 0.8B, 2026-10-17)
  const hint = best ? `\nNajtrafniejsze zdanie: [${results.indexOf(best.result) + 1}] „${best.sentence}”` : "";
  return `Fragmenty:\n${blocks.join("\n")}\n\nPytanie: ${question}${hint}`;
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

// ── Odpowiedź bez dużego modelu: najlepsze ZDANIE z dokumentu ────────────────
// Pytanie → typ oczekiwanej odpowiedzi (liczba, data, kwota, osoba…) + słowa pytania.
// Zdania z najlepszych fragmentów punktujemy: wspólne rdzenie słów (waga rzadkości), pasujący
// typ (np. „ile dni” → zdanie z liczbą i „dni”), miejsce fragmentu na liście. Natychmiast,
// także na telefonie bez pobierania czegokolwiek.
const DOC_ASK_TYPES = [
  { q: /\bile\s+(dni|tygodni|miesięcy|miesiecy|lat|godzin)\b|\bjak długo\b|\bhow (long|many days)\b/i, a: /\d+\s*(dni|dnia|tygodni|miesięcy|miesiąca|lat|roku|godzin)|\b(dni|dnia|miesięcy|lat)\b/i },
  { q: /\bile\s+(kosztuje|wynosi|zł|złotych|pln)\b|\bjaka\s+(kwota|cena|opłata|stawka|kaucja)\b|\bhow much\b/i, a: /\d[\d\s.,]*\s*(zł|złotych|pln|eur|euro|%)|kwot|wynagrodzen|czynsz|kaucj/i },
  { q: /\bile\b|\bhow many\b/i, a: /\d/ },
  { q: /\bkiedy\b|\bdo kiedy\b|\bod kiedy\b|\bw jakim terminie\b|\bwhen\b/i, a: /\d{1,2}[.\-/ ]\d{1,2}|\d{4}|\b(stycz|lut|mar|kwie|maj|czerw|lip|sierp|wrze|paźdz|pazdz|listop|grud)\w*|termin|dni|dnia/i },
  { q: /\bkto\b|\bkogo\b|\bwho\b/i, a: /\b\p{Lu}\p{Ll}+\s+\p{Lu}\p{Ll}+|najemc|wynajmując|zleceniodawc|wykonawc|stron[ay]/iu },
  { q: /\bgdzie\b|\bwhere\b/i, a: /\bul\.|\bw\s+\p{Lu}|adres|siedzib|lokal/iu },
];

function docAskSentences(text) {
  // Kropka po skrócie („ust.”, „r.”, „pkt.”) nie kończy zdania.
  return String(text).split(/(?<=[.!?…])\s+(?=[\p{Lu}„"(\d§])/u).map((x) => x.trim()).filter((x) => x.length > 12);
}

function docAskBestSentence(question, results) {
  const qTerms = [...new Set(docAskTerms(question))];
  const type = DOC_ASK_TYPES.find((tp) => tp.q.test(question));
  const all = [];
  results.slice(0, 4).forEach((r, rank) => docAskSentences(r.text).forEach((sentence) => all.push({ sentence, rank, result: r })));
  if (!all.length) return null;
  const df = new Map();
  all.forEach((x) => new Set(docAskTerms(x.sentence)).forEach((term) => df.set(term, (df.get(term) || 0) + 1)));
  let best = null;
  for (const x of all) {
    const terms = new Set(docAskTerms(x.sentence));
    let score = 0;
    qTerms.forEach((term) => { if (terms.has(term)) score += Math.log(1 + all.length / (df.get(term) || 1)); });
    if (type && type.a.test(x.sentence)) score += 1.6;
    score -= x.rank * 0.35;
    score -= Math.max(0, x.sentence.length - 260) / 400; // długie zdanie = mniej konkretna odpowiedź
    if (!best || score > best.score) best = { ...x, score };
  }
  return best && best.score > 0.4 ? best : null;
}

function docAskEl(tag, cls, text) {
  return Object.assign(document.createElement(tag), { className: cls || "", textContent: text ?? "" });
}

function docAskCiteButton(n, result) {
  const b = docAskEl("button", "doc-ask-cite", String(n));
  b.type = "button";
  b.setAttribute("aria-label", t("askCiteAria", { n }));
  b.addEventListener("click", () => docAskJump(result));
  return b;
}

// state: { best, results, ai: { text, pending } | null, notice }
function docAskRender(state) {
  docAskAnswerEl.replaceChildren();
  docAskAnswerEl.hidden = false;
  const { best, results } = state;
  if (state.notice) docAskAnswerEl.append(docAskEl("p", "hint", state.notice));
  if (best) {
    const n = results.indexOf(best.result) + 1;
    const box = docAskEl("div", "doc-ask-quote");
    box.append(docAskEl("div", "doc-ask-label", t("askBestSentence")));
    const q = docAskEl("p", "doc-ask-quote-text");
    q.append(document.createTextNode(`„${best.sentence}” `), docAskCiteButton(n, best.result));
    box.append(q);
    docAskAnswerEl.append(box);
  }
  if (state.ai) {
    const box = docAskEl("div", "doc-ask-ai");
    box.append(docAskEl("span", "ai-badge", t("aiBadge")));
    const p = docAskEl("p", "doc-ask-text");
    const text = state.ai.text;
    let last = 0;
    for (const m of String(text).matchAll(/\[(\d{1,2})\]/g)) {
      p.append(document.createTextNode(text.slice(last, m.index)));
      const n = Number(m[1]);
      p.append(results[n - 1] ? docAskCiteButton(n, results[n - 1]) : document.createTextNode(m[0]));
      last = m.index + m[0].length;
    }
    p.append(document.createTextNode(text.slice(last)));
    box.append(p);
    if (!state.ai.pending) {
      const none = /brak w dokumencie|not in the document/i.test(text);
      if (!docAskCitations(text, results.length).length && !none) box.append(docAskEl("p", "doc-ask-warn", t("askNoCitation")));
    }
    docAskAnswerEl.append(box);
  }
  if (!results.length || state.ai?.pending) return;
  const src = docAskEl("details", "doc-ask-sources");
  src.append(docAskEl("summary", "doc-ask-sources-title", t("askSources", { n: results.length })));
  results.forEach((r, i) => {
    const b = docAskEl("button", "semantic-result doc-ask-source");
    b.type = "button";
    b.append(docAskEl("span", "semantic-result-score", `[${i + 1}]`), Object.assign(docAskEl("span", "semantic-result-text", r.text), { title: r.text }));
    b.addEventListener("click", () => docAskJump(r));
    src.append(b);
  });
  docAskAnswerEl.append(src);
}

// Linijka pod odpowiedzią: czym odpowiadamy i jak włączyć więcej (bez okien zgody przy każdym pytaniu).
function docAskSyncEngine() {
  const el = document.getElementById("docAskEngine");
  if (!el || typeof LocalAI === "undefined") return;
  el.replaceChildren();
  if (LocalAI.isReady()) {
    el.append(docAskEl("span", "hint", t("askEngineAi", { name: LocalAI.readyName?.() || "" })));
    return;
  }
  el.append(docAskEl("span", "hint", t("askEngineNone")));
  const b = docAskEl("button", "btn doc-ask-setup", t("askEngineSetup"));
  b.type = "button";
  b.dataset.localAiSettings = "";
  el.append(b);
}

async function docAskRun() {
  const question = String(docAskInputEl?.value || "").replace(/\s+/g, " ").trim();
  if (!question) { docAskInputEl?.focus(); return; }
  if (!originalFileBytes) { toast(t("noFileToSave"), "warning"); return; }
  docAskBtn.disabled = true;
  try {
    const { results } = await docAskRetrieve(question);
    if (!results.length) {
      docAskRender({ best: null, results: [], notice: t("askNothing") });
      return;
    }
    const best = docAskBestSentence(question, results);
    const state = { best, results, ai: null, notice: best ? "" : t("askNoSentence") };
    docAskRender(state);
    if (best) docAskJump(best.result);
    if (typeof LocalAI === "undefined" || !LocalAI.isReady()) return;
    state.ai = { text: "…", pending: true };
    docAskRender(state);
    const answer = await LocalAI.generate({
      // ze zdaniem z aplikacji — bez furtki „Brak w dokumencie” (mały model z niej nadużywał)
      system: best ? DOC_ASK_SYSTEM_WITH_HINT : DOC_ASK_SYSTEM,
      shots: [DOC_ASK_SHOT],
      user: docAskPrompt(question, results, best),
      maxTokens: 160,
      onPartial: (txt) => { state.ai = { text: txt, pending: true }; docAskRender(state); },
    });
    const text = answer.trim() || t("askEmpty");
    // model zaprzecza zdaniu, które aplikacja znalazła — zostaje samo zdanie z dokumentu
    if (best && /^brak w dokumencie|not in the document/i.test(text)) state.ai = null;
    else state.ai = { text, pending: false };
    docAskRender(state);
  } catch (err) {
    docAskRender({ best: null, results: [], notice: t("aiFailed", { msg: err.message || err }) });
  } finally {
    docAskBtn.disabled = false;
  }
}

if (typeof LocalAI !== "undefined") LocalAI.onChange(docAskSyncEngine);
document.getElementById("panel-ask")?.addEventListener("toggle", (e) => { if (e.target.open) docAskSyncEngine(); });
document.addEventListener("DOMContentLoaded", docAskSyncEngine);

docAskBtn?.addEventListener("click", docAskRun);
docAskInputEl?.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); docAskRun(); } });
