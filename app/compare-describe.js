// compare-describe.js — „Opis zmian” w panelu Porównaj wersje: co się zmieniło, po ludzku.
//
// Rdzeń bez AI: z wyrównania akapitów (docx-compare.js) i różnicy słów wyciągamy pary
// „było → jest”, a liczby, kwoty, daty i terminy idą na początek — to zwykle te zmiany,
// które w umowie czy regulaminie naprawdę mają znaczenie. Wynik da się skopiować do maila.
// Opcjonalnie asystent AI (local-ai.js) pisze z tej listy 2–4 zdania; każda liczba w jego
// tekście musi występować w liście zmian, inaczej opis jest odrzucany.

const DWB_DESCRIBE_UNIT_RE = /^(dni|dnia|dzień|tygodni|tygodnie|tydzień|miesięcy|miesiące|miesiąc|lat|lata|rok|roku|r\.|godzin|godziny|godz\.|minut|zł|złotych|pln|eur|euro|usd|%|proc\.|procent|szt\.|sztuk|km|m2|m²|kg)(?=[\s.,;:!?)]|$)/iu;
const DWB_DESCRIBE_MAX = 24; // tyle pozycji listy — dalej „i N innych”

function dwbDescribeExcerpt(text, max = 70) {
  const s = String(text || "").replace(/\s+/g, " ").trim();
  return s.length > max ? `${s.slice(0, max - 1).trimEnd()}…` : s;
}

// „§ 3”, „Art. 5”, „2.1.” — numer jednostki z początku akapitu, inaczej jego początek
function dwbDescribeLabel(text) {
  const s = String(text || "").trim();
  const m = s.match(/^(§\s*\d+[a-z]?|art\.?\s*\d+[a-z]?|rozdział\s+[\divxlc]+|\d+(?:\.\d+)*)\.?(?=\s|$)/iu);
  return m ? m[1].replace(/\s+/g, " ") : `„${dwbDescribeExcerpt(s, 32)}”`;
}

// Pary zmian w jednym akapicie: kolejne usunięte + dodane słowa między fragmentami wspólnymi.
// Jednostka tuż za liczbą („14” → „30” + „dni”) dołącza do pary, żeby było wiadomo, co to za liczba.
function dwbDescribePairs(before, after) {
  const parts = dwbCompareWords(before, after);
  const pairs = [];
  let cur = null;
  let prevSame = "";
  for (const p of parts) {
    if (p.kind === "same") {
      if (cur) {
        const unit = p.text.trim().match(DWB_DESCRIBE_UNIT_RE);
        if (unit && /\d\s*$/.test(cur.removed.trim() + cur.added.trim())) {
          cur.removed = cur.removed.trimEnd() + (cur.removed.trim() ? ` ${unit[0]}` : "");
          cur.added = cur.added.trimEnd() + (cur.added.trim() ? ` ${unit[0]}` : "");
        }
        pairs.push(cur);
        cur = null;
      }
      prevSame = p.text;
      continue;
    }
    cur ||= { removed: "", added: "", prevSame };
    if (p.kind === "removed") cur.removed += p.text;
    else cur.added += p.text;
  }
  if (cur) pairs.push(cur);
  return pairs.map((x) => ({ before: x.removed.trim(), after: x.added.trim() })).filter((x) => x.before || x.after);
}

// older/newer: tekst akapitów; rows: [{ kind, o, n }] w kierunku starsza → nowsza.
function dwbDescribeChanges(older, newer, rows) {
  const stats = { changed: 0, added: 0, removed: 0 };
  const numbers = [];
  const wording = [];
  const added = [];
  const removed = [];
  for (const row of rows) {
    if (row.kind === "same") continue;
    stats[row.kind] = (stats[row.kind] || 0) + 1;
    if (row.kind === "added") { added.push({ type: "added", text: dwbDescribeExcerpt(newer[row.n], 90) }); continue; }
    if (row.kind === "removed") { removed.push({ type: "removed", text: dwbDescribeExcerpt(older[row.o], 90) }); continue; }
    const label = dwbDescribeLabel(newer[row.n]);
    const pairs = dwbDescribePairs(older[row.o], newer[row.n]);
    const totalLen = pairs.reduce((sum, p) => sum + p.before.length + p.after.length, 0);
    const base = Math.max(older[row.o].length, newer[row.n].length) || 1;
    if (!pairs.length) continue;
    // przeredagowany akapit: zmiana obejmuje większość tekstu — lista par byłaby nieczytelna
    if (totalLen / base > 0.9 && pairs.length > 3) {
      const nums = pairs.filter((p) => /\d/.test(p.before + p.after));
      nums.forEach((p) => numbers.push({ type: "number", label, before: dwbDescribeExcerpt(p.before, 40), after: dwbDescribeExcerpt(p.after, 40) }));
      wording.push({ type: "rewritten", label });
      continue;
    }
    for (const p of pairs) {
      const item = { label, before: dwbDescribeExcerpt(p.before, 60), after: dwbDescribeExcerpt(p.after, 60) };
      if (/\d/.test(p.before + p.after)) numbers.push({ type: "number", ...item });
      else wording.push({ type: "wording", ...item });
    }
  }
  return { stats, numbers, wording, added, removed };
}

function dwbDescribePairText(item) {
  if (item.type === "rewritten") return `${item.label}: ${t("describeRewritten")}`;
  if (!item.before) return `${item.label}: ${t("describeInserted", { text: item.after })}`;
  if (!item.after) return `${item.label}: ${t("describeDeleted", { text: item.before })}`;
  return `${item.label}: „${item.before}” → „${item.after}”`;
}

// Ten sam opis jako zwykły tekst (do schowka / maila i jako materiał dla AI).
function dwbDescribeText(desc, max = DWB_DESCRIBE_MAX) {
  const lines = [t("describeStats", desc.stats)];
  let left = max;
  const section = (title, items, fmt) => {
    if (!items.length) return;
    lines.push("", title);
    const shown = items.slice(0, Math.max(0, left));
    shown.forEach((it) => lines.push(`• ${fmt(it)}`));
    left -= shown.length;
    if (items.length > shown.length) lines.push(`• ${t("describeMore", { n: items.length - shown.length })}`);
  };
  section(t("describeNumbers"), desc.numbers, dwbDescribePairText);
  section(t("describeWording"), desc.wording, dwbDescribePairText);
  section(t("describeAddedTitle"), desc.added, (it) => `„${it.text}”`);
  section(t("describeRemovedTitle"), desc.removed, (it) => `„${it.text}”`);
  return lines.join("\n");
}

// Liczby z tekstu w postaci porównywalnej („4 500,00 zł” → „4500”, „30.” → „30”)
function dwbDescribeNumbers(text) {
  return (String(text).match(/\d[\d\s .,]*\d|\d/g) || [])
    .map((n) => n.replace(/[\s ]/g, "").replace(/[.,]0+$/, "").replace(/[.,]$/, ""))
    .filter(Boolean);
}

// Opis AI przechodzi tylko, gdy nie wymyślił żadnej liczby.
function dwbDescribeCheckAi(aiText, factsText) {
  const known = new Set(dwbDescribeNumbers(factsText));
  const unknown = dwbDescribeNumbers(aiText).filter((n) => !known.has(n));
  return { ok: unknown.length === 0, unknown };
}

// Polecenie dostrojone pod małe modele (docs/llm-bench/prompts.html, 2026-10-17: 0.8B 1/2 → 2/2):
// krótko, „przepisujesz liczby dokładnie”, bez przykładu (przykład przeciekał do odpowiedzi).
const DWB_DESCRIBE_SYSTEM = "Streszczasz podane zmiany w 1–3 zdaniach po polsku. Najpierw zmiany liczb, kwot, dat i terminów. Przepisujesz liczby dokładnie tak, jak w faktach; nie dodajesz innych.";

async function dwbDescribeWithAi(factsText, { onPartial } = {}) {
  const text = await LocalAI.generate({
    system: DWB_DESCRIBE_SYSTEM,
    user: `Fakty:\n${factsText}\n\nStreszczenie:`,
    maxTokens: 220,
    onPartial,
  });
  return { text: text.trim(), check: dwbDescribeCheckAi(text, factsText) };
}
