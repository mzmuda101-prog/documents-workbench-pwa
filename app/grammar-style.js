// Grammar / typography rules — offline scan on paragraph text from DOCX XML.

const NBSP = "\u00A0";

const GRAMMAR_RULES = [
  {
    id: "double-space",
    langs: ["pl", "en"],
    scan(text) {
      const hits = [];
      const re = / {2,}/g;
      let m;
      while ((m = re.exec(text)) !== null) {
        hits.push({ start: m.index, end: m.index + m[0].length, before: m[0], after: " " });
      }
      return hits;
    },
    fixAll(text) { return text.replace(/ {2,}/g, " "); },
  },
  {
    id: "space-before-punct",
    langs: ["pl", "en"],
    scan(text) {
      const hits = [];
      const re = /\s+([,.;:!?])/g;
      let m;
      while ((m = re.exec(text)) !== null) {
        hits.push({ start: m.index, end: m.index + m[0].length, before: m[0], after: m[1] });
      }
      return hits;
    },
    fixAll(text) { return text.replace(/\s+([,.;:!?])/g, "$1"); },
  },
  {
    id: "ellipsis",
    langs: ["pl", "en"],
    scan(text) {
      const hits = [];
      const re = /\.{3}/g;
      let m;
      while ((m = re.exec(text)) !== null) {
        hits.push({ start: m.index, end: m.index + m[0].length, before: m[0], after: "…" });
      }
      return hits;
    },
    fixAll(text) { return text.replace(/\.{3}/g, "…"); },
  },
  {
    id: "quotes-pl",
    langs: ["pl"],
    scan(text) {
      const hits = [];
      const re = /"([^"\n]+)"/g;
      let m;
      while ((m = re.exec(text)) !== null) {
        const inner = m[1];
        hits.push({
          start: m.index,
          end: m.index + m[0].length,
          before: m[0],
          after: `„${inner}"`,
        });
      }
      return hits;
    },
    fixAll(text) { return text.replace(/"([^"\n]+)"/g, "„$1\""); },
  },
  {
    id: "cap-after-period",
    langs: ["pl", "en"],
    scan(text) {
      const hits = [];
      const re = /(^|[.!?]\s+)([a-ząćęłńóśźż])/gu;
      let m;
      while ((m = re.exec(text)) !== null) {
        const prefix = m[1];
        const letter = m[2];
        const start = m.index + prefix.length;
        if (/\w\.\w/.test(text.slice(Math.max(0, start - 4), start + 2))) continue; // skip abbreviations
        hits.push({
          start,
          end: start + 1,
          before: letter,
          after: letter.toLocaleUpperCase(currentLang === "en" ? "en-US" : "pl-PL"),
        });
      }
      return hits;
    },
    fixAll(text) {
      const locale = currentLang === "en" ? "en-US" : "pl-PL";
      return text.replace(/(^|[.!?]\s+)([a-ząćęłńóśźż])/gu, (full, prefix, letter, offset, src) => {
        const pos = offset + prefix.length;
        if (/\w\.\w/.test(src.slice(Math.max(0, pos - 4), pos + 2))) return full;
        return prefix + letter.toLocaleUpperCase(locale);
      });
    },
  },
  {
    id: "orphan-i",
    langs: ["pl"],
    scan(text) {
      const hits = [];
      const re = /(\S+)\s+(i)\.\s*$/u;
      const m = text.match(re);
      if (!m) return hits;
      const start = m.index;
      const prefix = text.slice(0, start);
      const lead = prefix.length && !/\s$/.test(prefix) ? " " : "";
      hits.push({
        start,
        end: text.length,
        before: m[0],
        after: `${lead}${m[2]} ${m[1]}.`,
      });
      return hits;
    },
    fixAll(text) {
      return text.replace(/(\S+)\s+(i)\.\s*$/u, (full, word, conj, offset, src) => {
        const prefix = src.slice(0, offset);
        const lead = prefix.length && !/\s$/.test(prefix) ? " " : "";
        return `${lead}${conj} ${word}.`;
      });
    },
  },
  {
    id: "nbsp-pl",
    langs: ["pl"],
    optional: true,
    scan(text, opts) {
      if (!opts?.nbspPl) return [];
      const hits = [];
      const re = /\b([wzioua])\s+/giu;
      let m;
      while ((m = re.exec(text)) !== null) {
        hits.push({
          start: m.index,
          end: m.index + m[0].length,
          before: m[0],
          after: m[1] + NBSP,
        });
      }
      return hits;
    },
    fixAll(text, opts) {
      if (!opts?.nbspPl) return text;
      return text.replace(/\b([wzioua])\s+/giu, (_, ch) => ch + NBSP);
    },
  },
];

function getEnabledGrammarRules(opts = {}) {
  const { lang = "pl", rules: ruleIds, nbspPl = false } = opts;
  return GRAMMAR_RULES.filter((rule) => {
    if (ruleIds?.length && !ruleIds.includes(rule.id)) return false;
    if (rule.optional && rule.id === "nbsp-pl" && !nbspPl) return false;
    if (rule.langs && !rule.langs.includes(lang)) return false;
    return true;
  });
}

function applyHitToParagraph(text, hit) {
  return text.slice(0, hit.start) + hit.after + text.slice(hit.end);
}

function buildGrammarSnippet(text, hit, maxLen = 80) {
  const ctx = 18;
  const from = Math.max(0, hit.start - ctx);
  const to = Math.min(text.length, hit.end + ctx);
  const slice = text.slice(from, to);
  const relStart = hit.start - from;
  const relEnd = hit.end - from;
  const shown = slice.slice(0, relStart) + hit.before + slice.slice(relEnd);
  const fixed = slice.slice(0, relStart) + hit.after + slice.slice(relEnd);
  let out = `"${shown.trim()}" → "${fixed.trim()}"`;
  if (out.length > maxLen) out = out.slice(0, maxLen - 1) + "…";
  return out;
}

function scanParagraph(text, lang, opts = {}) {
  const hits = [];
  const rules = getEnabledGrammarRules({ ...opts, lang });
  rules.forEach((rule) => {
    (rule.scan(text, opts) || []).forEach((m, i) => {
      hits.push({
        id: `${rule.id}-${i}-${m.start}`,
        ruleId: rule.id,
        start: m.start,
        end: m.end,
        before: m.before,
        after: m.after,
        snippet: buildGrammarSnippet(text, m),
        fixedParagraph: applyHitToParagraph(text, m),
      });
    });
  });
  return hits;
}

function fixParagraphWithRules(text, rules, opts = {}) {
  let out = text;
  rules.forEach((rule) => { out = rule.fixAll(out, opts); });
  return out;
}

async function scanDocument(bytes, opts = {}) {
  if (!bytes?.length) return { hits: [], byRule: {}, paragraphCount: 0 };
  await ensureDocLibs(false);
  const texts = await extractParagraphTextsFromDocx(bytes);
  const rules = getEnabledGrammarRules(opts);
  const hits = [];
  const byRule = {};

  texts.forEach((text, paraIndex) => {
    if (!text) return;
    rules.forEach((rule) => {
      (rule.scan(text, opts) || []).forEach((m, i) => {
        const hit = {
          id: `${rule.id}-${paraIndex}-${i}-${m.start}`,
          ruleId: rule.id,
          paraIndex,
          start: m.start,
          end: m.end,
          before: m.before,
          after: m.after,
          snippet: buildGrammarSnippet(text, m),
          fixedParagraph: applyHitToParagraph(text, m),
        };
        hits.push(hit);
        if (!byRule[rule.id]) byRule[rule.id] = [];
        byRule[rule.id].push(hit);
      });
    });
  });

  return { hits, byRule, paragraphCount: texts.length };
}

function buildGrammarBatchItems(texts, hits, opts = {}, filterRuleId = null) {
  const rules = getEnabledGrammarRules(opts);
  const ruleFilter = filterRuleId
    ? rules.filter((r) => r.id === filterRuleId)
    : rules;
  if (!ruleFilter.length) return [];

  const paraSet = new Set(hits.map((h) => h.paraIndex));
  const items = [];
  paraSet.forEach((paraIndex) => {
    const raw = texts[paraIndex];
    const fixed = fixParagraphWithRules(raw, ruleFilter, opts);
    if (fixed !== raw) items.push({ index: paraIndex, text: fixed });
  });
  return items;
}
