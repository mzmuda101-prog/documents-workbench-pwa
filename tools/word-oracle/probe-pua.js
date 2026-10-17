const { buildDocx } = require("./build-docx");
const word = require("./word");
(async () => {
  const res = [];
  for (const c of ["", "", "", "", "", "", "▪", "●", "‣", "◦", "⁃", "·", "o", ""]) {
    const s = word.statistics(await buildDocx({ body: [`x ${c} y`] }));
    res.push(`${c.codePointAt(0).toString(16)}=${s.words}`);
  }
  console.log(res.join("  "));
  const b = await buildDocx({ abstractNums: [{ levels: [{ fmt: "bullet", text: "o" }] }, { levels: [{ fmt: "bullet", text: "" }] }], nums: [{ id: 1, abs: 0 }, { id: 2, abs: 1 }], body: [{ text: "a", numId: 1 }, { text: "b", numId: 2 }] });
  console.log("listy o / F0B7", word.statistics(b));
})();
