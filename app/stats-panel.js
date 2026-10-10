// Panel „Statystyki” — liczby, czas czytania, najczęstsze słowa, długie zdania (ze skokiem) i
// statystyki zaznaczenia. Wszystko liczone lokalnie z DOM podglądu (widać też Twoje edycje).

const stGridEl = document.getElementById("stGrid");
const stSelEl = document.getElementById("stSelection");
const stTopEl = document.getElementById("stTopWords");
const stLongEl = document.getElementById("stLongSentences");
const stRefreshBtn = document.getElementById("stRefreshBtn");
const stLongLimitEl = document.getElementById("stLongLimit");

const ST_WORD_RE = /[\p{L}\p{N}][\p{L}\p{N}'’-]*/gu;
const ST_READ_WPM = 200;

function stWords(text) {
  return text.match(ST_WORD_RE) || [];
}

function stSentences(text) {
  return (text.match(/[^.!?…]+(?:[.!?…]+|$)/g) || []).map((s) => s.trim()).filter((s) => /[\p{L}\p{N}]/u.test(s));
}

function stTextStats(text) {
  const words = stWords(text);
  return {
    words: words.length,
    chars: text.length,
    charsNoSpaces: text.replace(/\s/g, "").length,
    sentences: stSentences(text).length,
  };
}

function stFmtMinutes(words) {
  const min = words / ST_READ_WPM;
  return min < 1 ? t("statsUnderMinute") : t("statsMinutes", { n: Math.max(1, Math.round(min)) });
}

function stStatCell(label, value) {
  const cell = document.createElement("div");
  cell.className = "stat-cell";
  const v = document.createElement("strong");
  v.textContent = value;
  const l = document.createElement("span");
  l.textContent = label;
  cell.append(v, l);
  return cell;
}

function stListRow(text, count, onClick) {
  const row = document.createElement(onClick ? "button" : "div");
  row.className = "stat-row";
  if (onClick) { row.type = "button"; row.addEventListener("click", onClick); }
  const a = document.createElement("span");
  a.textContent = text;
  const b = document.createElement("em");
  b.textContent = count;
  row.append(a, b);
  return row;
}

function computeDocumentStats() {
  const a = analyzeDocumentDom(docCanvasEl);
  const text = a.outline.filter((o) => o.type !== "table").map((o) => o.el.textContent || "").join("\n");
  const totals = stTextStats(text);
  const freq = new Map();
  stWords(text.toLowerCase()).forEach((w) => { if (w.length >= 4) freq.set(w, (freq.get(w) || 0) + 1); });
  const top = [...freq.entries()].filter(([, n]) => n > 1).sort((x, y) => y[1] - x[1] || x[0].localeCompare(y[0], "pl")).slice(0, 8);
  const limit = Math.max(10, Number(stLongLimitEl?.value) || 35);
  const long = [];
  a.outline.forEach((o) => {
    if (o.type === "table") return;
    stSentences(o.el.textContent || "").forEach((s) => {
      const n = stWords(s).length;
      if (n > limit) long.push({ el: o.el, n, text: s });
    });
  });
  long.sort((x, y) => y.n - x.n);
  return { a, totals, top, long, limit };
}

function renderDocumentStats() {
  if (!stGridEl) return;
  stGridEl.replaceChildren();
  stTopEl?.replaceChildren();
  stLongEl?.replaceChildren();
  if (!originalFileBytes) { syncSelectionStats(); return; }
  const { a, totals, top, long } = computeDocumentStats();
  const avg = totals.sentences ? Math.round((totals.words / totals.sentences) * 10) / 10 : 0;
  stGridEl.append(
    stStatCell(t("statsWords"), fmtNum(totals.words)),
    stStatCell(t("statsChars"), fmtNum(totals.chars)),
    stStatCell(t("statsCharsNoSpaces"), fmtNum(totals.charsNoSpaces)),
    stStatCell(t("statsSentences"), fmtNum(totals.sentences)),
    stStatCell(t("statsParagraphs"), fmtNum(a.paragraphs)),
    stStatCell(t("statsTables"), fmtNum(a.tables)),
    stStatCell(t("statsReading"), stFmtMinutes(totals.words)),
    stStatCell(t("statsAvgSentence"), fmtNum(avg, { maximumFractionDigits: 1 })),
  );
  if (stTopEl) {
    if (!top.length) stTopEl.appendChild(Object.assign(document.createElement("p"), { className: "hint", textContent: t("statsNoRepeats") }));
    top.forEach(([w, n]) => stTopEl.appendChild(stListRow(w, `${n}×`)));
  }
  if (stLongEl) {
    if (!long.length) stLongEl.appendChild(Object.assign(document.createElement("p"), { className: "hint", textContent: t("statsNoLong") }));
    long.slice(0, 12).forEach((s) => {
      const label = s.text.length > 90 ? `${s.text.slice(0, 89)}…` : s.text;
      stLongEl.appendChild(stListRow(label, t("statsWordsShort", { n: s.n }), () => {
        jumpToStructureItem({ el: s.el, id: "stats" }, { silentSelect: true });
        if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
      }));
    });
  }
  syncSelectionStats();
}

function syncSelectionStats() {
  if (!stSelEl) return;
  const sel = window.getSelection?.();
  const inDoc = sel && !sel.isCollapsed && sel.rangeCount && docCanvasEl?.contains(sel.anchorNode);
  if (!inDoc) { stSelEl.textContent = t("statsSelectionHint"); return; }
  const s = stTextStats(sel.toString());
  stSelEl.textContent = t("statsSelection", { words: s.words, chars: s.chars, min: stFmtMinutes(s.words) });
}

stRefreshBtn?.addEventListener("click", renderDocumentStats);
// „input” z opóźnieniem, nie „change”: change odpala się przy utracie fokusu — czyli w chwili
// dotknięcia wiersza listy, który wtedy rysuje się od nowa pod palcem i kliknięcie przepada.
let stLimitTimer = 0;
stLongLimitEl?.addEventListener("input", () => {
  clearTimeout(stLimitTimer);
  stLimitTimer = setTimeout(renderDocumentStats, 300);
});
document.addEventListener("selectionchange", () => {
  if (document.getElementById("panel-stats")?.open) syncSelectionStats();
});
