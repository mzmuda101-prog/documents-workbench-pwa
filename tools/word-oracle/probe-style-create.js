// probe-style-create.js — co zapisuje PRAWDZIWY Word w nowym stylu tworzonym z zaznaczonego akapitu
// (Style → Nowy styl; AppleScript „make new Word style” przy zaznaczeniu): XML stylu (bazowy,
// następny, pPr, rPr) i czy akapit go dostaje / traci formatowanie wprost.
//   node tools/word-oracle/probe-style-create.js
const { buildDocx } = require("./build-docx");
const word = require("./word");

const pick = (xml, re) => (xml.match(re) || [""])[0].replace(/ w:rsid\w*="[^"]*"/g, "");
(async () => {
  const styles = `<w:style w:type="paragraph" w:styleId="Wskazowka"><w:name w:val="Wskazówka"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="567"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>`;
  const body = [
    { pPr: '<w:jc w:val="center"/><w:spacing w:before="240" w:after="120"/>', raw: '<w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:color w:val="C00000"/><w:sz w:val="32"/></w:rPr><w:t>Wzorcowy akapit</w:t></w:r>' },
    { text: "Zwykły akapit" },
    { style: "Wskazowka", pPr: '<w:jc w:val="right"/>', raw: '<w:r><w:rPr><w:color w:val="0070C0"/></w:rPr><w:t>Wskazówka z formatem</w:t></w:r>' },
  ];
  for (const [label, n] of [["akapit Normalny + format wprost", 1], ["akapit „Wskazówka” + format wprost", 3]]) {
    const { parts } = word.saveDocx(await buildDocx({ body, styles }), `
      select text object of paragraph ${n}
      make new Word style at it with properties {name local:"Mój styl", style type:style type paragraph}
      set style of paragraph 2 to "Mój styl"`, ["word/styles.xml", "word/document.xml"]);
    const st = pick(parts["word/styles.xml"] || "", /<w:style [^>]*>(?:(?!<\/w:style>).)*?<w:name w:val="Mój styl"\/>[\s\S]*?<\/w:style>/);
    const paras = (parts["word/document.xml"].match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).map((p) => p.replace(/ w:rsid\w*="[^"]*"| w14:\w+="[^"]*"/g, ""));
    console.log(`── ${label}\nSTYL: ${st}\nAKAPIT ${n}: ${paras[n - 1]}\nAKAPIT 2 (nadany styl): ${paras[1]}\n`);
  }
})();
