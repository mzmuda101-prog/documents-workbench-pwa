const { buildDocx } = require("./build-docx");
const word = require("./word");
(async () => {
  for (const t of ["a—b", "a–b", "a-b", "a/b", "a…b", "a,b", "a.b", "a b", "a​b", "a b", "a　b", "a­b", "a‑b", "a—", "—b", "a — b", "a b", "pierwszy—drugi—trzeci"]) {
    const bytes = await buildDocx({ body: [t] });
    const s = word.statistics(bytes);
    console.log(JSON.stringify(t), s.words, s.chars, s.charsWithSpaces);
  }
  const b = await buildDocx({ abstractNums: [{ levels: [{ fmt: "decimal", text: "%1." }] }], nums: [{ id: 1, abs: 0 }], body: [{ text: "", numId: 1 }, { text: "x", numId: 1 }] });
  console.log("pusty punkt listy", word.statistics(b));
})();
