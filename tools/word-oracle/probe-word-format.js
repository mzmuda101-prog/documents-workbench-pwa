// probe-word-format.js — Word: pogrubienie / kursywa przy samym kursorze (bez zaznaczenia) w środku
// słowa, na jego początku, za nim i na spacji — co dostaje format. node tools/word-oracle/probe-word-format.js
const { buildDocx } = require("./build-docx");
const word = require("./word");
const SEL = '(get selection of application "Microsoft Word")';
const runs = (xml) => (xml.match(/<w:r>[\s\S]*?<\/w:r>|<w:r [\s\S]*?<\/w:r>/g) || []).map((r) => (/<w:b\/>|<w:b w:val="1"\/>/.test(r) ? "**" : "") + (r.match(/<w:t[^>]*>([^<]*)<\/w:t>/) || [])[1]).join("|");
(async () => {
  const bytes = await buildDocx({ body: ["Alfa Beta Gamma"] });
  for (const [name, at] of [["w środku „Beta”", 7], ["na początku „Beta”", 5], ["tuż za „Beta”", 9], ["na końcu akapitu", 15]]) {
    const r = word.saveDocx(bytes, `select (create range it start ${at} end ${at})
    set bold of font object of ${SEL} to true
    type text ${SEL} text "X"`);
    console.log(`${name}: ${runs(r.parts["word/document.xml"])}`);
  }
})();
