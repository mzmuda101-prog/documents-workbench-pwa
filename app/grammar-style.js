// Grammar / typography rules — offline scan on paragraph text from DOCX XML.

const NBSP = "\u00A0";

function grammarLocale(opts = {}) {
  return opts.lang === "en" ? "en-US" : "pl-PL";
}

// Skróty, po których kropka NIE kończy zdania („sp. z o.o.”, „np. w”, „art. 5 ust. 2 pkt. a”).
// Dawniej reguła wielkiej litery robiła z „ACME sp. z o.o.” → „sp. Z o.o.”.
const ABBREVIATIONS = new Set(`sp np m.in in ul al pl os tj tzn tzw itd itp jw ok godz min max zł gr r rr w ww wg tel
 nr poz pkt ust art par str s t tys mln mld dr prof mgr inż hab lek im św gen płk kpt por ks bp pn woj pow gm m st
 ang łac niem zob cd cdn dot ew wyd oprac red tłum zw mies tyg kw sek pon wt śr czw pt sob niedz sty lut mar kwi
 cze lip sie wrz paź lis gru vs etc e.g i.e mr mrs ms jr sr inc ltd co no vol fig approx dept z o.o p.n.e n.e
 n.p.m j.w tj r.p b.r ub.r bm b.m`.split(/\s+/).filter(Boolean));

// Czy małej litery na pozycji `letterPos` NIE podnosić: kropka po skrócie, po inicjale albo
// w środku skrótu z kropkami („o.o.”, „m.in.”); na początku akapitu: punkt listy „a) …”.
function isCapAfterPeriodException(text, matchStart, letterPos) {
  if (letterPos === 0 || !/[.!?]/.test(text[matchStart] || "")) {
    return /^[a-ząćęłńóśźż][).]/u.test(text.slice(letterPos, letterPos + 2)); // „a) …”, „b. …”
  }
  if (text[matchStart] !== ".") return false; // po ! i ? zawsze nowe zdanie
  const before = text.slice(0, matchStart);
  const word = (before.match(/([\p{L}.]+)$/u) || [])[1] || "";
  const bare = word.toLocaleLowerCase("pl-PL").replace(/^\.+/, "");
  if (!bare) return false;
  if (ABBREVIATIONS.has(bare) || ABBREVIATIONS.has(bare.replace(/\./g, ""))) return true;
  if (/\./.test(bare)) return true; // „o.o”, „m.in”, „p.n.e” — kropki w środku = skrót
  if (/^\p{L}$/u.test(bare)) return true; // inicjał / pojedyncza litera („J. kowalski”, „pkt a. b”)
  return /\w\.\w/.test(text.slice(Math.max(0, letterPos - 4), letterPos + 2));
}

// domeny i rozszerzenia plików („firma.pl”, „raport.docx”) — kropka w nich to nie koniec zdania
const ADDRESS_TLDS = /^(pl|com|org|net|eu|io|gov|edu|info|de|uk|us|fr|it|es|cz|sk|ua|app|dev|ai|docx?|xlsx?|pptx?|pdf|txt|md|html?|csv|odt|rtf|jpe?g|png|gif|heic|zip|json|js|xml)(?![\p{L}])/iu;
function isAddressLike(text, pos) {
  const startTok = text.lastIndexOf(" ", pos) + 1;
  const endRel = text.slice(pos).search(/\s/);
  const token = text.slice(startTok, endRel < 0 ? text.length : pos + endRel);
  if (/@|:\/\/|^www\./i.test(token)) return true;
  return text[pos] === "." && ADDRESS_TLDS.test(text.slice(pos + 1));
}

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
    // „koniec.nowe”, „tak,jak” — brak spacji po znaku. Oba sąsiednie wyrazy z co najmniej 2 liter:
    // skróty z kropkami w środku („o.o.”, „m.in.”) i liczby („3,5”) zostają. Pomijamy adresy:
    // token z „@”, „://”, „www.” albo kończący się domeną („firma.pl”, „example.com/x”).
    id: "space-after-punct",
    langs: ["pl", "en"],
    scan(text) {
      const hits = [];
      const re = /(?<=[\p{L}]{2})([.,;:!?])(?=[\p{L}]{2})/gu;
      let m;
      while ((m = re.exec(text)) !== null) {
        if (isAddressLike(text, m.index)) continue;
        hits.push({ start: m.index, end: m.index + 1, before: m[1], after: `${m[1]} ` });
      }
      return hits;
    },
    fixAll(text) {
      return text.replace(/(?<=[\p{L}]{2})([.,;:!?])(?=[\p{L}]{2})/gu, (full, punct, offset, src) => (isAddressLike(src, offset) ? full : `${punct} `));
    },
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
          after: `„${inner}”`, // polski zamykający to ” (U+201D), nie prosty "
        });
      }
      return hits;
    },
    fixAll(text) { return text.replace(/"([^"\n]+)"/g, "„$1”"); },
  },
  {
    id: "cap-after-period",
    langs: ["pl", "en"],
    scan(text, opts = {}) {
      const hits = [];
      const re = /(^|[.!?]\s+)([a-ząćęłńóśźż])/gu;
      let m;
      while ((m = re.exec(text)) !== null) {
        const prefix = m[1];
        const letter = m[2];
        const start = m.index + prefix.length;
        if (isCapAfterPeriodException(text, m.index, start)) continue;
        hits.push({
          start,
          end: start + 1,
          before: letter,
          after: letter.toLocaleUpperCase(grammarLocale(opts)),
        });
      }
      return hits;
    },
    fixAll(text, opts = {}) {
      const locale = grammarLocale(opts);
      return text.replace(/(^|[.!?]\s+)([a-ząćęłńóśźż])/gu, (full, prefix, letter, offset, src) => {
        const pos = offset + prefix.length;
        if (isCapAfterPeriodException(src, offset, pos)) return full;
        return prefix + letter.toLocaleUpperCase(locale);
      });
    },
  },
  {
    // Tylko wyrazy stojące bezpośrednio obok siebie. Nie zgadujemy semantyki
    // ani nie dotykamy fraz rozdzielonych interpunkcją.
    id: "repeated-word",
    langs: ["pl", "en"],
    scan(text) {
      const hits = [];
      const re = /(?<![\p{L}\p{N}_])([\p{L}]+(?:['’\-][\p{L}]+)*)[ \t]+\1(?![\p{L}\p{N}_])/giu;
      let m;
      while ((m = re.exec(text)) !== null) {
        hits.push({ start: m.index, end: m.index + m[0].length, before: m[0], after: m[1] });
      }
      return hits;
    },
    fixAll(text) {
      return text.replace(/(?<![\p{L}\p{N}_])([\p{L}]+(?:['’\-][\p{L}]+)*)[ \t]+\1(?![\p{L}\p{N}_])/giu, "$1");
    },
  },
  {
    id: "nbsp-pl",
    langs: ["pl"],
    optional: true,
    scan(text, opts) {
      if (!opts?.nbspPl) return [];
      const hits = [];
      const re = /(?<![\p{L}\p{N}])([wzioua])[ \t]+/giu; // \b nie zna polskich liter („cała sprawa”)
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
      return text.replace(/(?<![\p{L}\p{N}])([wzioua])[ \t]+/giu, (_, ch) => ch + NBSP);
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

// Kontekst wokół zmiany do podglądu w panelu (sama zmiana pokazana osobno: skreślone → wstawione)
function grammarContext(text, hit, ctx = 22) {
  const from = Math.max(0, hit.start - ctx);
  const to = Math.min(text.length, hit.end + ctx);
  return {
    ctxBefore: (from > 0 ? "…" : "") + text.slice(from, hit.start),
    ctxAfter: text.slice(hit.end, to) + (to < text.length ? "…" : ""),
  };
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
        ...grammarContext(text, m),
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
          ...grammarContext(text, m),
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
