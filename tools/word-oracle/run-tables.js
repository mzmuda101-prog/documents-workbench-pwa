// run-tables.js — scalanie / dzielenie komórek / wiersz nagłówka: Word kontra aplikacja (tabela po zapisie)
// node tools/word-oracle/run-tables.js [filtr]; WRITE=1 → expected-tables.json (scripts/table-word-playwright.js)
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { buildDocx } = require("./build-docx");
const word = require("./word");
const app = require("./app");
const { CASES, tableXml, STYLES, cellParaIndex, normalizeTable } = require("./cases-tables");

const docOf = (c) => ({ styles: STYLES, body: ["Przed.", tableXml(c.rows, { header: c.header }), "Po."] });
function wordOp(bytes, c) {
  const key = crypto.createHash("sha1").update(bytes).update(JSON.stringify(c.op)).digest("hex").slice(0, 16);
  const f = path.join(word.DIR, "cache", `tables-${key}.json`);
  if (fs.existsSync(f) && !process.env.FRESH) return JSON.parse(fs.readFileSync(f, "utf8"));
  const o = c.op;
  const cell = (r, k) => `(get cell from table (table 1) row ${r} column ${k})`;
  const body = o.merge ? `merge cell ${cell(o.merge[0], o.merge[1])} with ${cell(o.merge[2], o.merge[3])}`
    : o.split ? `split cell ${cell(o.split[0], o.split[1])} number of rows ${o.rows} number of columns ${o.cols}`
    : `set heading format of row ${o.header} of table 1 to (not (heading format of row ${o.header} of table 1))`;
  const out = { doc: word.saveDocx(bytes, body).parts["word/document.xml"] };
  fs.mkdirSync(path.dirname(f), { recursive: true });
  fs.writeFileSync(f, JSON.stringify(out));
  return out;
}
// ta sama operacja w aplikacji (jak przyciski w menu Tabela)
function appEdit(c) {
  const { cellParaIndex } = require("./cases-tables");
  const o = c.op;
  if (o.merge) return { op: "table", index: cellParaIndex(c.rows, o.merge[0], o.merge[1]), action: "merge", toIndex: cellParaIndex(c.rows, o.merge[2], o.merge[3]) };
  if (o.split) return { op: "table", index: cellParaIndex(c.rows, o.split[0], o.split[1]), action: "split" };
  return { op: "table", index: cellParaIndex(c.rows, o.header, 1), action: "headerRow" };
}

module.exports = { appEdit, docOf };
if (require.main === module) (async () => {
  const filter = process.argv[2];
  let bad = 0;
  const expected = {};
  for (const c of CASES.filter((x) => !filter || x.name.includes(filter))) {
    const bytes = await buildDocx(docOf(c));
    const wx = wordOp(bytes, c);
    const { browser, page } = await app.openApp();
    await app.loadBytes(page, bytes, `tables-${c.name}.docx`);
    await page.evaluate(() => { if (readOnlyMode) appFrame.setReadOnly(false); });
    const n = await page.evaluate((e) => applyDocumentEdit(e), appEdit(c));
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 });
    const a = await page.evaluate(normalizeTable, await app.savedParts(page));
    const w = await page.evaluate(normalizeTable, wx);
    await browser.close();
    expected[c.name] = w;
    const sa = JSON.stringify(a.rows.map((r) => [r.h, r.cells]));
    const sw = JSON.stringify(w.rows.map((r) => [r.h, r.cells]));
    const widths = JSON.stringify([a.grid, a.rows.map((r) => r.w)]) === JSON.stringify([w.grid, w.rows.map((r) => r.w)]);
    console.log(`${sa === sw ? (widths ? "✅" : "🟡") : "❌"} ${c.name}${n ? "" : " (apka: 0 zmian)"}`);
    if (sa !== sw) { bad++; console.log(`   Word: ${sw}\n   apka: ${sa}`); }
    if (!widths) console.log(`   szerokości Word: ${JSON.stringify([w.grid, w.rows.map((r) => r.w)])}\n   szerokości apka: ${JSON.stringify([a.grid, a.rows.map((r) => r.w)])}`);
  }
  if (process.env.WRITE && !filter) fs.writeFileSync(path.join(__dirname, "expected-tables.json"), `${JSON.stringify(expected, null, 1)}\n`);
  console.log(bad ? `\n${bad} przypadków różni się od Worda` : "\nwszystko jak w Wordzie");
})();
