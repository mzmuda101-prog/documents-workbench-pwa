const { buildDocx } = require("./build-docx");
const word = require("./word");
const FN = { "footnotes.xml": { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", rel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes", xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Treść przypisu tu.</w:t></w:r></w:p></w:footnote></w:footnotes>' } };
(async () => {
  const cases = {
    pageBreak: { body: [{ raw: '<w:r><w:t>a.</w:t><w:br w:type="page"/><w:t>b</w:t></w:r>' }] },
    colBreak: { body: [{ raw: '<w:r><w:t>a</w:t><w:br w:type="column"/><w:t>b</w:t></w:r>' }] },
    fnInWord: { body: [{ raw: '<w:r><w:t>ab</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:t>cd</w:t></w:r>' }], extraParts: FN },
    fnAfterWord: { body: [{ raw: '<w:r><w:t>ab</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> cd</w:t></w:r>' }], extraParts: FN },
  };
  for (const [k, c] of Object.entries(cases)) {
    const bytes = await buildDocx(c);
    console.log(k.padEnd(12), JSON.stringify(word.statistics(bytes)), JSON.stringify(word.statistics(bytes, true)));
  }
})();
