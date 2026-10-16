// spell-check.js — ortografia w panelu „Korekta” (słownik Hunspell, offline, bez AI).
//
// Skan: wyrazy z akapitów dokumentu → unikalne → spell-worker.js (Hunspell) → lista błędnych
// wyrazów z miejscami wystąpień. Podpowiedzi liczone partiami (bywają wolne: do ~0,2 s na wyraz).
// Pomijane: skróty WIELKIMI literami, wyrazy z cyframi, MieszaneWielkości (iPhone), adresy WWW
// i e-mail, słownik użytkownika (localStorage) i „Pomiń” z bieżącej sesji.

const dwbSpell = (() => {
  const USER_DICT_KEY = "dwb.userDictionary.v1";
  const WORD_RE = /\p{L}[\p{L}\p{M}\p{N}'’-]*/gu;
  const LINKISH_RE = /\b(?:https?:\/\/|www\.)\S+|[\w.+-]+@[\w-]+\.[\w.-]+/giu;
  let worker = null;
  let seq = 0;
  const pending = new Map();
  const ignored = new Set();

  function ensureWorker() {
    if (worker) return worker;
    const v = typeof lazyAssetVersion === "function" ? lazyAssetVersion() : "1";
    worker = new Worker(`app/spell-worker.js?v=${encodeURIComponent(v)}`, { type: "module", name: "dwb-spell" });
    worker.addEventListener("message", ({ data }) => {
      const p = pending.get(data.id);
      if (!p) return;
      pending.delete(data.id);
      if (data.error) p.reject(new Error(data.error));
      else p.resolve(data);
    });
    worker.addEventListener("error", (e) => {
      pending.forEach((p) => p.reject(new Error(e.message || "spell worker")));
      pending.clear();
      worker = null;
    });
    return worker;
  }

  function call(message) {
    const id = ++seq;
    return new Promise((resolve, reject) => {
      pending.set(id, { resolve, reject });
      ensureWorker().postMessage({ ...message, id });
    });
  }

  function userWords() {
    try { return new Set(JSON.parse(localStorage.getItem(USER_DICT_KEY) || "[]")); } catch { return new Set(); }
  }
  function addToDictionary(word) {
    const set = userWords();
    set.add(word.toLocaleLowerCase());
    try { localStorage.setItem(USER_DICT_KEY, JSON.stringify([...set].sort())); } catch {}
  }
  function ignore(word) { ignored.add(word.toLocaleLowerCase()); }

  function skipWord(word) {
    if (word.length < 2) return true;
    if (/\p{N}/u.test(word)) return true;
    const letters = word.replace(/[^\p{L}]/gu, "");
    if (letters.length >= 2 && letters === letters.toLocaleUpperCase() && letters.length <= 8) return true; // RODO, PESEL
    if (/\p{Ll}\p{Lu}/u.test(word)) return true; // iPhone, ChatGPT, eDowód
    return false;
  }

  // [{ word, start, end }] — wyrazy do sprawdzenia w jednym akapicie
  function tokenize(text) {
    const s = String(text || "");
    const blocked = [];
    for (const m of s.matchAll(LINKISH_RE)) blocked.push([m.index, m.index + m[0].length]);
    const out = [];
    for (const m of s.matchAll(WORD_RE)) {
      let word = m[0].replace(/['’-]+$/u, "");
      const start = m.index;
      const end = start + word.length;
      if (blocked.some(([a, b]) => start < b && end > a)) continue;
      if (skipWord(word)) continue;
      out.push({ word, start, end });
    }
    return out;
  }

  // texts: tekst akapitów (indeks = akapit pliku). Zwraca [{ word, count, occurrences: [{ paraIndex, start, end }] }]
  async function scan(texts, lang = "pl") {
    const occurrences = new Map();
    const user = userWords();
    texts.forEach((text, paraIndex) => {
      if (!text) return;
      for (const tok of tokenize(text)) {
        const low = tok.word.toLocaleLowerCase();
        if (user.has(low) || ignored.has(low)) continue;
        if (!occurrences.has(tok.word)) occurrences.set(tok.word, []);
        occurrences.get(tok.word).push({ paraIndex, start: tok.start, end: tok.end });
      }
    });
    if (!occurrences.size) return [];
    const { bad } = await call({ type: "check", lang, words: [...occurrences.keys()] });
    return bad.map((word) => ({ word, count: occurrences.get(word).length, occurrences: occurrences.get(word), suggestions: null }))
      .sort((a, b) => a.occurrences[0].paraIndex - b.occurrences[0].paraIndex || a.occurrences[0].start - b.occurrences[0].start);
  }

  async function suggest(words, lang = "pl") {
    if (!words.length) return {};
    const { suggestions } = await call({ type: "suggest", lang, words, limit: 4 });
    return suggestions;
  }

  // Wielka litera jak w oryginale: „Dokumętów” → „Dokumentów”.
  function matchCase(original, replacement) {
    if (original === original.toLocaleUpperCase() && original.length > 1) return replacement.toLocaleUpperCase();
    if (original[0] === original[0].toLocaleUpperCase()) return replacement[0].toLocaleUpperCase() + replacement.slice(1);
    return replacement;
  }

  // Nowy tekst akapitów po zamianie wyrazu we wskazanych miejscach (od końca — przesunięcia).
  function replaceIn(texts, occurrences, replacement) {
    const byPara = new Map();
    for (const o of occurrences) {
      if (!byPara.has(o.paraIndex)) byPara.set(o.paraIndex, []);
      byPara.get(o.paraIndex).push(o);
    }
    const items = [];
    byPara.forEach((list, paraIndex) => {
      let text = texts[paraIndex];
      if (text == null) return;
      list.sort((a, b) => b.start - a.start).forEach((o) => {
        const original = text.slice(o.start, o.end);
        text = text.slice(0, o.start) + matchCase(original, replacement) + text.slice(o.end);
      });
      if (text !== texts[paraIndex]) items.push({ index: paraIndex, text });
    });
    return items;
  }

  return { scan, suggest, tokenize, replaceIn, matchCase, addToDictionary, ignore, userWords };
})();
