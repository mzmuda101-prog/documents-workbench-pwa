const { buildDocx } = require("./build-docx");
const word = require("./word");
(async () => {
  const cases = {
    lead: '<w:r><w:t>A</w:t></w:r><w:r><w:t> b</w:t></w:r>',
    trail: '<w:r><w:t>A </w:t></w:r><w:r><w:t>b</w:t></w:r>',
    inner2: '<w:r><w:t>A  b</w:t></w:r>',
    onlySpace: '<w:r><w:t>A</w:t></w:r><w:r><w:t> </w:t></w:r><w:r><w:t>b</w:t></w:r>',
    nbspEdge: '<w:r><w:t>A</w:t></w:r><w:r><w:t> b</w:t></w:r>',
    tabChar: '<w:r><w:t>A\tb</w:t></w:r>',
    newline: '<w:r><w:t>A\nb</w:t></w:r>',
    preserve: '<w:r><w:t xml:space="preserve">A </w:t></w:r><w:r><w:t>b</w:t></w:r>',
    paraPreserve: '<w:r><w:t>A </w:t></w:r><w:r><w:t>b</w:t></w:r>',
  };
  for (const [k, raw] of Object.entries(cases)) {
    const bytes = await buildDocx({ body: [{ raw }] });
    const s = word.statistics(bytes);
    const t = word.listStrings(bytes)[0].text;
    console.log(k.padEnd(12), JSON.stringify(t), s.words, s.charsWithSpaces);
  }
})();
