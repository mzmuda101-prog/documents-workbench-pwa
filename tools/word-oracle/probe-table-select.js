// probe-table-select.js — Word: zaznaczenie przez kilka komórek tabeli + pisanie / Backspace / Delete.
// Zakres od początku A1 do środka B1 (jak Shift+→ w aplikacji), potem klawisz; stan tabeli po zapisie.
// Uruchom: node tools/word-oracle/probe-table-select.js
const { buildDocx } = require("./build-docx");
const word = require("./word");
const SEL = '(get selection of application "Microsoft Word")';
const cell = (t) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p></w:tc>`;
const TBL = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell("A1") + cell("B1")}</w:tr><w:tr>${cell("A2") + cell("B2")}</w:tr></w:tbl>`;
const texts = (xml) => (xml.match(/<w:p\b[\s\S]*?<\/w:p>|<w:p\/>/g) || []).map((p) => (p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, "")).join("") || "∅").join(" ¶ ");
(async () => {
  const bytes = await buildDocx({ body: ["Przed", TBL, "Po"] });
  // A1 zaczyna się po „Przed\r” (6 znaków); B1 = A1 + „A1\x07” (3)
  const cases = {
    "A1→środek B1 + pisz X": `select (create range it start 6 end 10)
    type text ${SEL} text "X"`,
    "A1→środek B1 + Backspace": `select (create range it start 6 end 10)
    type backspace ${SEL}`,
    "A1→środek B1 + Delete (clear)": `select (create range it start 6 end 10)
    clear ${SEL}`,
    "cały wiersz 1 + Backspace": `select (create range it start 6 end 12)
    type backspace ${SEL}`,
    "zakres: zaznaczone komórki": `select (create range it start 6 end 10)
    set res to (start of content of text object of ${SEL} as string) & "-" & (end of content of text object of ${SEL} as string)`,
  };
  for (const [name, body] of Object.entries(cases)) {
    try {
      const r = word.saveDocx(bytes, body);
      console.log(`${name}\n   ${texts(r.parts["word/document.xml"] || "")}  ${r.res || ""}`);
    } catch (e) { console.log(`${name}\n   BŁĄD ${e.message}`); }
  }
})();
