// spell-worker.js — ortografia offline (Hunspell w WebAssembly) w osobnym wątku.
// Słowniki: lib/hunspell/dict (PL: sjp.pl 2026-09-01, EN: dictionary-en) — ładowane przy
// pierwszym skanie, potem z cache service workera. Nic nie wychodzi z urządzenia.

import { createHunspellFromStrings } from "../lib/hunspell/js/Hunspell.js";

const instances = new Map();

function dictionary(lang) {
  const key = lang === "en" ? "en" : "pl";
  if (!instances.has(key)) {
    instances.set(key, (async () => {
      const base = new URL("../lib/hunspell/dict/", import.meta.url);
      const [aff, dic] = await Promise.all(["aff", "dic"].map(async (ext) => {
        const res = await fetch(new URL(`${key}.${ext}`, base));
        if (!res.ok) throw new Error(`dictionary ${key}.${ext}: ${res.status}`);
        return res.text();
      }));
      return createHunspellFromStrings(aff, dic);
    })());
  }
  return instances.get(key);
}

// Wyraz z łącznikiem („biało-czerwony”, „e-mailem”): całość albo każda część osobno.
function correct(h, word) {
  if (h.testSpelling(word)) return true;
  if (word.includes("-")) {
    const parts = word.split("-").filter(Boolean);
    return parts.length > 1 && parts.every((p) => h.testSpelling(p));
  }
  return false;
}

self.onmessage = async ({ data }) => {
  const { id, type, lang } = data;
  try {
    const h = await dictionary(lang);
    if (type === "check") {
      const bad = data.words.filter((w) => !correct(h, w));
      self.postMessage({ id, bad });
    } else if (type === "suggest") {
      const out = {};
      for (const w of data.words) out[w] = h.getSpellingSuggestions(w).slice(0, data.limit || 4);
      self.postMessage({ id, suggestions: out });
    }
  } catch (err) {
    self.postMessage({ id, error: String(err?.message || err) });
  }
};
