const { buildDocx } = require("./build-docx");
const word = require("./word");
(async () => {
  const texts = ["to JEST A test 3D ABC123 DON'T zażółć X. y I", "aBC Abc ABc ABC-def ABC-DEF ŻÓŁW. ąą", "PIERWSZE zdanie. drugie ZDANIE", "x <b>Kot</b>"];
  const ops = [];
  texts.forEach((t) => ["sentence", "title"].forEach((m) => ops.push({ text: t, mode: m })));
  const bytes = await buildDocx({ body: ops.map((o) => o.text) });
  const w = word.changeCase(bytes, ops.map((o, i) => ({ ...o, para: i + 1, from: 0, to: null })));
  ops.forEach((o, i) => console.log(o.mode.padEnd(8), JSON.stringify(o.text), "→", JSON.stringify(w[i])));
  // pogrubione słowo w środku: Red<b>bold</b> — tytuł całości
  const b2 = await buildDocx({ body: [{ raw: '<w:r><w:t>red</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold text</w:t></w:r>' }] });
  console.log("runy:", word.changeCase(b2, [{ para: 1, mode: "title" }]));
})();
