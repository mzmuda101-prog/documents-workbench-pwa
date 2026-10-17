// probe-style-apply.js — co Word zdejmuje z akapitu przy nadaniu stylu (lista stylów / galeria)
const { buildDocx } = require("./build-docx");
const word = require("./word");
const { QUOTE_STYLES, BOX } = require("./cases-quote");
const H1 = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:style>';
const DIRECT = [
  ["jc", '<w:jc w:val="center"/>'],
  ["spacing", '<w:spacing w:before="480" w:after="480"/>'],
  ["ind", '<w:ind w:left="1440" w:firstLine="360"/>'],
  ["box", BOX],
  ["numPr", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'],
  ["keepNext", "<w:keepNext/>"],
  ["tabs", '<w:tabs><w:tab w:val="left" w:pos="3000"/></w:tabs>'],
  ["outline", '<w:outlineLvl w:val="2"/>'],
  ["pageBreakBefore", "<w:pageBreakBefore/>"],
  ["widow", '<w:widowControl w:val="0"/>'],
];
const RUNS = '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">całe pogrubione </w:t></w:r>';
const PART = '<w:r><w:t xml:space="preserve">zwykły </w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr><w:t>kawałek</w:t></w:r>';
(async () => {
  const body = [];
  for (const target of ["Normal", "H1"]) {
    DIRECT.forEach(([k, x]) => body.push({ pPr: x, text: `${target}:${k}` }));
    body.push({ raw: RUNS, text: "" }, { raw: PART, text: "" });
  }
  const bytes = await buildDocx({ styles: QUOTE_STYLES + H1, abstractNums: [{ levels: [{ fmt: "decimal" }] }], nums: [{ id: 1, abs: 0 }], body });
  const n = DIRECT.length + 2;
  const ops = [];
  for (let i = 1; i <= n; i++) ops.push(`set style of paragraph ${i} to style normal`);
  for (let i = n + 1; i <= 2 * n; i++) ops.push(`set style of paragraph ${i} to style heading1`);
  const r = word.saveDocx(bytes, ops.join("\n"));
  const ps = r.parts["word/document.xml"].match(/<w:p[ >][\s\S]*?<\/w:p>/g).map((p) => p.replace(/ w14:\w+="[^"]*"| w:rsid\w*="[^"]*"/g, ""));
  ps.forEach((p, i) => console.log(i < n ? "N" : "H", p.replace(/<w:p>|<\/w:p>/g, "")));
})();
