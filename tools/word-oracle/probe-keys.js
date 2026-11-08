// probe-keys.js — co robi PRAWDZIWY Word z klawiszami w miejscach „do decyzji” tabeli klawiszy
// (scripts/keys-table.js): Cytat + Backspace na początku / Enter na końcu, tekst z dolną krawędzią
// + Enter na początku / końcu. node tools/word-oracle/probe-keys.js
const { buildDocx } = require("./build-docx");
const word = require("./word");

const BDR = '<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="auto"/></w:pBdr>';
// stan akapitów z zapisanego przez Worda pliku: styl, wcięcie, krawędź, tekst
function paras(xml) {
  return (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).map((p) => {
    const style = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || "";
    const ind = (p.match(/<w:ind [^>]*>/) || [])[0] || "";
    const bdr = /<w:pBdr>[\s\S]*?<w:bottom/.test(p) ? "krawędź" : "";
    const text = (p.match(/<w:t[^>]*>[^<]*<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, "")).join("");
    return [style, ind, bdr, text || "∅"].filter(Boolean).join(" ");
  });
}
const at = (para, end) => `set r to text object of paragraph ${para}
    select (create range it start (${end ? "(end of content of r) - 1" : "start of content of r"}) end (${end ? "(end of content of r) - 1" : "start of content of r"}))`;
const SEL = "(get selection of application \"Microsoft Word\")";

const CASES = [
  // Cytat: styl wbudowany Worda (nazwa lokalna „Cytat” albo „Quote”)
  { name: "Cytat · Backspace na początku", quote: true, body: `${at(2, false)}
    type backspace ${SEL}` },
  { name: "Cytat · Enter na końcu + litera", quote: true, body: `${at(2, true)}
    type paragraph ${SEL}
    type text ${SEL} text "X"` },
  { name: "Tekst z krawędzią · Enter na początku", border: true, body: `${at(2, false)}
    type paragraph ${SEL}` },
  { name: "Tekst z krawędzią · Enter na końcu", border: true, body: `${at(2, true)}
    type paragraph ${SEL}` },
];

(async () => {
  for (const c of CASES) {
    const bytes = await buildDocx({ body: ["Alfa", { text: "Beta", pPr: c.border ? BDR : "" }, "Gamma"] });
    const setQuote = c.quote ? `try
      set style of paragraph 2 to "Cytat"
    on error
      set style of paragraph 2 to "Quote"
    end try
    ` : "";
    try {
      const before = c.quote ? paras(word.saveDocx(bytes, setQuote).parts["word/document.xml"]) : null;
      const r = word.saveDocx(bytes, `${setQuote}${c.body}`, ["word/document.xml", "word/styles.xml"]);
      console.log(`\n■ ${c.name}`);
      if (before) console.log(`  przed: ${before.join("  ¶  ")}`);
      console.log(`  po:    ${paras(r.parts["word/document.xml"]).join("  ¶  ")}`);
      if (c.quote) {
        const st = (r.parts["word/styles.xml"].match(/<w:style [^>]*w:styleId="(Quote|Cytat)"[\s\S]*?<\/w:style>/) || [])[0] || "";
        console.log(`  styl Cytat: next=${(st.match(/<w:next w:val="([^"]+)"/) || [])[1] || "(brak)"} ind=${(st.match(/<w:ind [^>]*>/) || [])[0] || "(brak)"}`);
      }
    } catch (e) {
      console.log(`\n■ ${c.name}\n  BŁĄD: ${e.message}`);
    }
  }
})();
