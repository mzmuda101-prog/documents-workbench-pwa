// probe-columns-balance.js — czy Word wyrównuje kolumny sekcji: (a) po sekcji w kolumnach następna
// sekcja „następna strona”, (b) następna sekcja „ciągła”. Wypisuje kolumnę (1/2) akapitów A, B, C.
//   node tools/word-oracle/probe-columns-balance.js
const word = require("./word");
const { buildDocx } = require("./build-docx");

const SECT = (type, cols) => `<w:sectPr>${type ? `<w:type w:val="${type}"/>` : ""}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>${cols ? '<w:cols w:num="2" w:space="708"/>' : ""}</w:sectPr>`;
const where = `set res to ""
    repeat with i from 1 to count of paragraphs
      set pp to paragraph i
      set tx to content of text object of pp
      if tx starts with "Kol" then set res to res & tx & "@x=" & (get range information (text object of pp) information type horizontal position relative to page) & " "
    end repeat`;
(async () => {
  for (const nextType of ["", "continuous"]) {
    const body = [
      `<w:p><w:pPr>${SECT("", false)}</w:pPr><w:r><w:t>Sekcja pierwsza</w:t></w:r></w:p>`,
      "Kol A", "Kol B", "Kol C",
      `<w:p><w:pPr>${SECT("continuous", true)}</w:pPr><w:r><w:t>Kol koniec</w:t></w:r></w:p>`,
      "Sekcja trzecia",
    ];
    const bytes = await buildDocx({ body, sectPr: nextType ? `<w:type w:val="${nextType}"/>` : "" });
    console.log(`następna sekcja: ${nextType || "następna strona"} →`, word.withDocument(bytes, where));
  }
})();
