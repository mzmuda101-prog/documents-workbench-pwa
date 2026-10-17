// probe-style-runs.js — które cechy znaków na CAŁYM akapicie Word zdejmuje przy nadaniu stylu
const { buildDocx } = require("./build-docx");
const word = require("./word");
const PROPS = [
  ["b", "<w:b/>"], ["i", "<w:i/>"], ["u", '<w:u w:val="single"/>'], ["color", '<w:color w:val="FF0000"/>'],
  ["sz", '<w:sz w:val="32"/>'], ["rFonts", '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'], ["lang", '<w:lang w:val="en-US"/>'],
  ["highlight", '<w:highlight w:val="yellow"/>'], ["rStyle", '<w:rStyle w:val="Strong"/>'], ["shd", '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>'],
  ["caps", "<w:caps/>"], ["vanish-no", "<w:spacing w:val=\"20\"/>"],
];
const ST = '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>';
(async () => {
  const body = [];
  PROPS.forEach(([k, x]) => body.push({ raw: `<w:r><w:rPr>${x}</w:rPr><w:t xml:space="preserve">${k} cały </w:t></w:r><w:r><w:rPr>${x}</w:rPr><w:t>akapit</w:t></w:r>` }));
  // większość (3/4) i mniejszość (1/4) tekstu
  body.push({ raw: '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">większość pogrubiona aaa </w:t></w:r><w:r><w:t>b</w:t></w:r>' });
  body.push({ raw: '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">b </w:t></w:r><w:r><w:t>mniejszość pogrubiona aaa</w:t></w:r>' });
  const bytes = await buildDocx({ styles: ST, body });
  const ops = body.map((_, i) => `set style of paragraph ${i + 1} to style normal`).join("\n");
  const r = word.saveDocx(bytes, ops);
  r.parts["word/document.xml"].match(/<w:p[ >][\s\S]*?<\/w:p>/g).forEach((p) => console.log(p.replace(/ w14:\w+="[^"]*"| w:rsid\w*="[^"]*"/g, "")));
})();
