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
  // An adjacent deletion+addition is an edited paragraph, rather than two unrelated events.
  const compact = [];
  for (let k = 0; k < rows.length; k++) {
    const x = rows[k], y = rows[k + 1];
    if (x?.kind === "removed" && y?.kind === "added") { compact.push({ kind: "changed", left: x.left, right: y.right }); k++; }
    else compact.push(x);
  }
  return { rows: compact, approximate: false };
}

function dwbCompareWords(before, after) {
  const a = before.match(/\S+\s*/g) || [], b = after.match(/\S+\s*/g) || [];
  if (a.length * b.length > 80_000) return [{ kind: "removed", text: before }, { kind: "added", text: after }];
  const w = b.length + 1, grid = new Uint16Array((a.length + 1) * w);
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) grid[i * w + j] = a[i - 1] === b[j - 1] ? grid[(i - 1) * w + j - 1] + 1 : Math.max(grid[(i - 1) * w + j], grid[i * w + j - 1]);
  const out = []; let i = a.length, j = b.length;
  while (i || j) {
    if (i && j && a[i - 1] === b[j - 1]) out.push({ kind: "same", text: a[--i] });
    else if (j && (!i || grid[i * w + j - 1] >= grid[(i - 1) * w + j])) out.push({ kind: "added", text: b[--j] });
    else out.push({ kind: "removed", text: a[--i] });
  }
  return out.reverse();
}
