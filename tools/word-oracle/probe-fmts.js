const { buildDocx } = require("./build-docx");
const word = require("./word");
(async () => {
  for (const fmt of ["chicago", "decimalEnclosedCircle", "decimalEnclosedParen", "decimalEnclosedFullstop", "decimalFullWidth", "hex", "russianLower", "numberInDash"]) {
    const bytes = await buildDocx({ abstractNums: [{ levels: [{ fmt, text: "%1" }] }], nums: [{ id: 1, abs: 0 }], body: Array.from({ length: 22 }, (_, i) => ({ text: `p${i + 1}`, numId: 1 })) });
    console.log(fmt, word.listStrings(bytes).map((x) => x.label).join(" "));
  }
})();
