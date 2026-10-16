// DOCX comparison primitives.  Deliberately read-only: both packages stay untouched.
// The panel uses paragraph alignment first, then a word-level diff for changed paragraphs.

const DWB_COMPARE_CELL_LIMIT = 1_500_000;

function dwbCompareText(el) {
  return Array.from(el.getElementsByTagNameNS(W_NS, "t")).map((n) => n.textContent || "").join("").replace(/\s+/g, " ").trim();
}

async function dwbCompareParagraphs(bytes) {
  if (typeof ensureDocLibs === "function" && !(await ensureDocLibs(true))) throw new Error("libraries");
  const zip = await window.JSZip.loadAsync(bytes);
  const file = zip.file("word/document.xml");
  if (!file) throw new Error("documentXml");
  const xml = await file.async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("xml");
  return Array.from(doc.getElementsByTagNameNS(W_NS, "p"))
    .map(dwbCompareText)
    .filter((text) => text.length);
}

function dwbCompareKey(text) {
  return text.normalize("NFKC").toLocaleLowerCase().replace(/\s+/g, " ").trim();
}

// LCS makes moved/inserted paragraphs visible as distinct items.  A capped fallback prevents
// a huge pair of long documents from consuming the tab's memory just to build a comparison.
function dwbCompareAlign(left, right) {
  const a = left.map(dwbCompareKey), b = right.map(dwbCompareKey);
  if (a.length * b.length > DWB_COMPARE_CELL_LIMIT) {
    const rows = [];
    const n = Math.max(a.length, b.length);
    for (let i = 0; i < n; i++) {
      if (a[i] === b[i]) rows.push({ kind: "same", left: i, right: i });
      else if (i >= a.length) rows.push({ kind: "added", right: i });
      else if (i >= b.length) rows.push({ kind: "removed", left: i });
      else rows.push({ kind: "changed", left: i, right: i });
    }
    return { rows, approximate: true };
  }
  const width = b.length + 1;
  const matrix = new Uint16Array((a.length + 1) * width);
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) {
    matrix[i * width + j] = a[i - 1] === b[j - 1]
      ? matrix[(i - 1) * width + j - 1] + 1
      : Math.max(matrix[(i - 1) * width + j], matrix[i * width + j - 1]);
  }
  const rows = [];
  let i = a.length, j = b.length;
  while (i || j) {
    if (i && j && a[i - 1] === b[j - 1]) { rows.push({ kind: "same", left: --i, right: --j }); continue; }
    if (j && (!i || matrix[i * width + j - 1] >= matrix[(i - 1) * width + j])) rows.push({ kind: "added", right: --j });
    else rows.push({ kind: "removed", left: --i });
  }
  rows.reverse();
  return { rows: dwbComparePairRuns(rows, left, right), approximate: false };
}

// Akapity w tej samej „dziurze” między wspólnymi fragmentami: usunięty + dodany o podobnych
// słowach to ten sam akapit po redakcji („changed”), a nie dwa niezależne zdarzenia. Dawniej
// łączyło się tylko sąsiadów, więc trzy poprawione z rzędu akapity parowały się na krzyż
// (§ 3 z „Strony ustalają…”). Kolejność zachowana: para idzie w miejscu akapitu nowszego.
function dwbCompareWordSet(text) {
  return new Set((dwbCompareKey(text).match(/[\p{L}\p{N}]{2,}/gu) || []));
}
function dwbCompareSimilarity(a, b) {
  if (!a.size || !b.size) return 0;
  let common = 0;
  a.forEach((w) => { if (b.has(w)) common++; });
  return common / (a.size + b.size - common);
}
function dwbComparePairRuns(rows, left, right) {
  const out = [];
  let k = 0;
  while (k < rows.length) {
    if (rows[k].kind === "same") { out.push(rows[k++]); continue; }
    const run = [];
    while (k < rows.length && rows[k].kind !== "same") run.push(rows[k++]);
    const removed = run.filter((r) => r.kind === "removed");
    const added = run.filter((r) => r.kind === "added");
    const sets = new Map();
    const setOf = (side, i) => {
      const key = `${side}${i}`;
      if (!sets.has(key)) sets.set(key, dwbCompareWordSet(side === "l" ? left[i] : right[i]));
      return sets.get(key);
    };
    const pairOf = new Map(); // dodany → usunięty
    // jeden usunięty i jeden dodany w tym samym miejscu — ten sam akapit, nawet całkiem przepisany
    if (removed.length === 1 && added.length === 1) pairOf.set(added[0], removed[0]);
    let from = pairOf.size ? added.length : 0;
    removed.forEach((r) => {
      for (let j = from; j < added.length; j++) {
        if (dwbCompareSimilarity(setOf("l", r.left), setOf("r", added[j].right)) >= 0.35) {
          pairOf.set(added[j], r);
          from = j + 1;
          return;
        }
      }
    });
    const paired = new Set(pairOf.values());
    run.forEach((r) => {
      if (r.kind === "removed" && paired.has(r)) return;
      if (r.kind === "added" && pairOf.has(r)) out.push({ kind: "changed", left: pairOf.get(r).left, right: r.right });
      else out.push(r);
    });
  }
  return out;
}

function dwbCompareWords(before, after) {
  const a = before.match(/\S+\s*/g) || [], b = after.match(/\S+\s*/g) || [];
  if (a.length * b.length > 80_000) return [{ kind: "removed", text: before }, { kind: "added", text: after }];
  const w = b.length + 1, grid = new Uint16Array((a.length + 1) * w);
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) grid[i * w + j] = a[i - 1] === b[j - 1] ? grid[(i - 1) * w + j - 1] + 1 : Math.max(grid[(i - 1) * w + j], grid[i * w + j - 1]);
  const out = []; let i = a.length, j = b.length;
  while (i || j) {
    // wspólne słowo — cofamy się w OBU tekstach (dawniej tylko w starszym: każde wspólne słowo
    // wychodziło drugi raz jako „dodane”)
    if (i && j && a[i - 1] === b[j - 1]) { out.push({ kind: "same", text: a[--i] }); j--; }
    else if (j && (!i || grid[i * w + j - 1] >= grid[(i - 1) * w + j])) out.push({ kind: "added", text: b[--j] });
    else out.push({ kind: "removed", text: a[--i] });
  }
  return out.reverse();
}
