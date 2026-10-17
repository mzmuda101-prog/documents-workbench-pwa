// probe-tables.js — jak Word odpowiada na scalanie/dzielenie przez AppleScript (szybki podgląd XML)
const { buildDocx } = require("./build-docx");
const word = require("./word");
const { CASES, tableXml, STYLES } = require("./cases-tables");
(async () => {
  for (const c of CASES.filter((x) => !process.argv[2] || x.name.includes(process.argv[2]))) {
    const bytes = await buildDocx({ styles: STYLES, body: ["Przed.", tableXml(c.rows, { header: c.header }), "Po."] });
    const o = c.op;
    const body = o.merge ? `merge cell (get cell from table (table 1) row ${o.merge[0]} column ${o.merge[1]}) with (get cell from table (table 1) row ${o.merge[2]} column ${o.merge[3]})`
      : o.split ? `split cell (get cell from table (table 1) row ${o.split[0]} column ${o.split[1]}) number of rows ${o.rows} number of columns ${o.cols}`
      : `set heading format of row ${o.header} of table 1 to (not (heading format of row ${o.header} of table 1))`;
    try {
      const r = word.saveDocx(bytes, body);
      console.log(c.name, r.parts["word/document.xml"].match(/<w:tbl>[\s\S]*<\/w:tbl>/)[0].replace(/ w14:\w+="[^"]*"| w:rsid\w*="[^"]*"/g, "").replace(/<w:tr[ >]/g, "\n  $&").slice(0, 3000));
    } catch (e) { console.log(c.name, "BŁĄD", e.message); }
  }
})();
