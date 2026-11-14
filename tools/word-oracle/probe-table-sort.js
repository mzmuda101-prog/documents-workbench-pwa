// probe-table-sort.js — jak Word sortuje tabelę (Tabela → Sortuj): kolejność tekstu, liczb, dat, pustych.
// node tools/word-oracle/probe-table-sort.js  (wypisuje kolumnę po sortowaniu)
const { buildDocx } = require("./build-docx");
const word = require("./word");
const { tableXml, STYLES } = require("./cases-tables");

const TYPE = { text: "sort field alphanumeric", num: "sort field numeric", date: "sort field date" };
async function run(name, col, type, order = "ascending", header = false) {
  const rows = col.map((t) => [t, String(col.indexOf(t))]);
  const bytes = await buildDocx({ styles: STYLES, body: ["Przed.", tableXml(rows), "Po."] });
  const SEP = "\u0001";
  const out = word.withDocument(bytes, `
    sort (text object of table 1) exclude header ${header} field number 1 sort field type ${TYPE[type]} sort order sort order ${order}
    repeat with i from 1 to (count of rows of table 1)
      set res to res & (content of text object of (get cell from table (table 1) row i column 1)) & "${SEP}"
    end repeat`);
  console.log(`${name} [${type} ${order}${header ? " +nagłówek" : ""}]:`, JSON.stringify(out.split(SEP).slice(0, -1).map((s) => s.replace(/[\r\u0007]+$/, ""))));
}
(async () => {
  const only = process.argv[2];
  const C = [
    ["wielkosc", ["b", "B", "a", "A", "c"], "text"],
    ["polskie", ["Żaba", "zebra", "łoś", "lis", "ćma", "cel", "ą", "a", "Ąb", "Ab"], "text"],
    ["puste", ["b", "", "a", " ", "c"], "text"],
    ["puste-malejaco", ["b", "", "a", "c"], "text", "descending"],
    ["cyfry-jako-tekst", ["10", "9", "100", "a", "2"], "text"],
    ["liczby", ["10", "9", "100", "-3", "2,5", "1 000", "abc", "", "12 zł", "x7"], "num"],
    ["liczby-malejaco", ["10", "9", "abc", "", "2"], "num", "descending"],
    ["daty", ["15.11.2026", "1.02.2025", "2024-03-05", "abc", "", "3/4/2026"], "date"],
    ["naglowek", ["Imię", "c", "a", "b"], "text", "ascending", true],
    ["stabilnosc", ["x", "X", "x", "X"], "text"],
    ["liczby2", ["2,7", "2,5", "2.6", "3", "+4", "(5)", "0,5", " 1", "1.000", "1e3", "3a", "-", "−2"], "num"],
    ["stab-malejaco", ["x", "X", "a", "x", "X"], "text", "descending"],
    ["daty2", ["5.3.2026", "4.3.2026", "2026-03-04", "3 marca 2026", "12.2025", "1/2/2026", "13/1/2026"], "date"],
  ];
  for (const [n, col, type, order, header] of C) if (!only || n.includes(only)) {
    try { await run(n, col, type, order, header); } catch (e) { console.log(n, "BŁĄD:", e.message); }
  }
})();
