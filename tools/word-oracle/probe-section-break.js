// probe-section-break.js — co zapisuje PRAWDZIWY Word przy Układ → Podziały → Podział sekcji (następna strona /
// ciągły) wstawionym w środku akapitu: gdzie trafia w:sectPr, co zawiera (rozmiar, marginesy, nagłówki),
// jak dzieli tekst.   node tools/word-oracle/probe-section-break.js
const { buildDocx } = require("./build-docx");
const word = require("./word");

(async () => {
  const sectPr = '<w:titlePg/>';
  const body = ["Pierwszy akapit.", "Drugi akapit do podziału w środku.", "Trzeci akapit."];
  const bytes = await buildDocx({ body, sectPr });
  for (const kind of ["section break next page", "section break continuous"]) {
    const { parts } = word.saveDocx(bytes, `
      set s to (start of content of text object of paragraph 2) + 6
      set rg to create range it start s end s
      insert break at rg break type ${kind}`, ["word/document.xml"]);
    const x = parts["word/document.xml"].replace(/ w:rsid\w*="[^"]*"| w14:\w+="[^"]*"/g, "");
    const paras = (x.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).map((p) => p.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, ""));
    console.log(`── ${kind}`);
    paras.forEach((p, i) => console.log(` p${i + 1}: ${p.slice(0, 600)}`));
    console.log(" body sectPr:", (x.match(/<w:sectPr>(?:(?!<\/w:sectPr>)[\s\S])*<\/w:sectPr><\/w:body>/) || [""])[0].slice(0, 400), "\n");
  }
})();
