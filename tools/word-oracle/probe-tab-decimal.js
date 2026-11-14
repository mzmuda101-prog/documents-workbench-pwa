// probe-tab-decimal.js — gdzie PRAWDZIWY Word (polski) stawia „12,50” / „12.50” / „1.234,50” / „125” przy
// tabulatorze dziesiętnym na 10 cm od marginesu: położenie przecinka/kropki (albo końca) w pt.
//   node tools/word-oracle/probe-tab-decimal.js
const { buildDocx } = require("./build-docx");
const word = require("./word");

(async () => {
  const texts = ["12,50", "12.50", "1.234,50", "125"];
  const body = texts.map((t) => ({ pPr: '<w:tabs><w:tab w:val="decimal" w:pos="5669"/></w:tabs>', raw: `<w:r><w:t>A</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${t}</w:t></w:r>` }));
  const bytes = await buildDocx({ body });
  const res = word.withDocument(bytes, `repeat with i from 1 to ${texts.length}
    set pa to text object of paragraph i
    set s to start of content of pa
    set n to (end of content of pa) - s
    set res to res & i & ":"
    repeat with k from 0 to n - 1
      set rg to create range it start (s + k) end (s + k + 1)
      set res to res & (content of rg) & "@" & (get range information rg information type horizontal position relative to text boundary) & " "
    end repeat
    set res to res & linefeed
  end repeat`);
  console.log("cel: 10 cm = 283,46 pt od marginesu\n" + res);
})();
