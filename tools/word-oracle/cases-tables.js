// Narzędzia tabel: scalanie, dzielenie, wiersz nagłówka. Każdy przypadek = tabela (komórki tekstem,
// "" = pusta) + jedna operacja; porównujemy strukturę tabeli po zapisie (normalizeTable).
// Komórka: tekst albo { t, span, v: "restart"|"continue" } (scalenia już w pliku).
const { esc } = require("./build-docx");
const COLW = 2000;
function cellXml(c) {
  const o = typeof c === "string" ? { t: c } : c;
  const span = o.span || 1;
  const paras = (o.paras || [o.t || ""]).map((t) => `<w:p>${t ? `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>` : ""}</w:p>`).join("");
  return `<w:tc><w:tcPr><w:tcW w:w="${COLW * span}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ""}${o.v ? `<w:vMerge${o.v === "restart" ? ' w:val="restart"' : ""}/>` : ""}</w:tcPr>${paras}</w:tc>`;
}
function tableXml(rows, { header = [] } = {}) {
  const cols = Math.max(...rows.map((r) => r.reduce((n, c) => n + ((typeof c === "object" && c.span) || 1), 0)));
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${Array(cols).fill(`<w:gridCol w:w="${COLW}"/>`).join("")}</w:tblGrid>${rows.map((r, i) => `<w:tr>${header.includes(i) ? "<w:trPr><w:tblHeader/></w:trPr>" : ""}${r.map(cellXml).join("")}</w:tr>`).join("")}</w:tbl>`;
}
const STYLES = '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>';
const G3 = [["A1", "B1", "C1"], ["A2", "B2", "C2"], ["A3", "B3", "C3"]];
// op: { merge: [r0, c0, r1, c1] } (komórki 1…, jak w Wordzie: wiersz, kolumna siatki) | { split: [r, c], cols, rows } | { header: r }
const CASES = [
  { name: "scal-w-prawo", rows: G3, op: { merge: [1, 1, 1, 2] } },
  { name: "scal-w-dol", rows: G3, op: { merge: [1, 1, 2, 1] } },
  { name: "scal-prostokat-2x2", rows: G3, op: { merge: [2, 2, 3, 3] } },
  { name: "scal-caly-wiersz", rows: G3, op: { merge: [3, 1, 3, 3] } },
  { name: "scal-z-pustymi", rows: [["A1", "", "C1"], ["", "B2", ""], ["A3", "B3", "C3"]], op: { merge: [1, 1, 2, 2] } },
  { name: "scal-wszystkie-puste", rows: [["", "", "x"], ["", "", "y"]], op: { merge: [1, 1, 2, 2] } },
  { name: "scal-kilka-akapitow", rows: [[{ paras: ["a", "b"] }, { paras: ["c", "", "d"] }], ["x", "y"]], op: { merge: [1, 1, 1, 2] } },
  // (scalenie komórki już scalonej poziomo z węższą poniżej: AppleScript Worda „merge cell” tego nie umie)
  { name: "scal-pionowo-dwie-kolumny", rows: [["A1", "B1", "C1"], ["A2", "B2", "C2"], ["A3", "", "C3"]], op: { merge: [1, 2, 3, 2] } },
  { name: "scal-pelne-kolumny", rows: [["A1", "B1", "C1"], ["A2", "B2", "C2"]], op: { merge: [1, 2, 2, 3] } },
  { name: "podziel-zwykla-na-2", rows: G3, op: { split: [2, 2], cols: 2, rows: 1 } },
  { name: "podziel-z-tekstem-na-2", rows: [["Ala ma kota", "B1"], ["A2", "B2"]], op: { split: [1, 1], cols: 2, rows: 1 } },
  { name: "podziel-scalona-poziomo", rows: [[{ t: "AB", span: 2 }, "C1"], ["A2", "B2", "C2"]], op: { split: [1, 1], cols: 2, rows: 1 } },
  { name: "podziel-scalona-pionowo", rows: [[{ t: "A", v: "restart" }, "B1"], [{ t: "", v: "continue" }, "B2"]], op: { split: [1, 1], cols: 1, rows: 2 } },
  { name: "naglowek-pierwszy", rows: G3, op: { header: 1 } },
  { name: "naglowek-wylacz", rows: G3, header: [0, 1], op: { header: 1 } },
];
module.exports = { CASES, tableXml, STYLES };

// pierwszy akapit komórki (wiersz r, kolumna siatki c; od 1) jako indeks akapitu w dokumencie
// (akapit „Przed.” = 0, potem akapity komórek po kolei)
function cellParaIndex(rows, r, c) {
  let idx = 1;
  for (let i = 0; i < rows.length; i++) {
    let col = 1;
    for (const cell of rows[i]) {
      const o = typeof cell === "string" ? { t: cell } : cell;
      const span = o.span || 1;
      if (i === r - 1 && c >= col && c < col + span) return idx;
      idx += (o.paras || [0]).length;
      col += span;
    }
  }
  return -1;
}

// tabela po zapisie → { grid: [szerokości], rows: [{ h: nagłówek, cells: ["tekst¶tekst|s2|vR", …], w: [tcW] }] }
// (w przeglądarce: page.evaluate(normalizeTable, { doc }))
function normalizeTable({ doc }) {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const kids = (el, name) => Array.from(el?.childNodes || []).filter((n) => n.nodeType === 1 && n.namespaceURI === W && (!name || n.localName === name));
  const tbl = new DOMParser().parseFromString(doc, "application/xml").getElementsByTagNameNS(W, "tbl")[0];
  const val = (el) => el?.getAttributeNS(W, "val");
  return {
    grid: kids(kids(tbl, "tblGrid")[0], "gridCol").map((g) => Number(g.getAttributeNS(W, "w"))),
    rows: kids(tbl, "tr").map((tr) => {
      const h = kids(kids(tr, "trPr")[0], "tblHeader")[0];
      const cells = kids(tr, "tc");
      return {
        h: !!h && !/^(0|false|off)$/.test(val(h) || ""),
        cells: cells.map((tc) => {
          const pr = kids(tc, "tcPr")[0];
          const span = Number(val(kids(pr, "gridSpan")[0]) || 1);
          const vm = kids(pr, "vMerge")[0];
          const text = kids(tc, "p").map((p) => Array.from(p.getElementsByTagNameNS(W, "t")).map((t) => t.textContent).join("")).join("¶");
          return `${text}${span > 1 ? `|s${span}` : ""}${vm ? (val(vm) === "restart" ? "|vR" : "|vC") : ""}`;
        }),
        w: cells.map((tc) => Number(kids(kids(tc, "tcPr")[0], "tcW")[0]?.getAttributeNS(W, "w") || 0)),
      };
    }),
  };
}
module.exports.cellParaIndex = cellParaIndex;
module.exports.normalizeTable = normalizeTable;
