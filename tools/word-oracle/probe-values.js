// Word: jak wygląda wartość N w danym formacie (lista po 1 punkcie na wartość, startOverride)
const fs = require("fs");
const { buildDocx } = require("./build-docx");
const word = require("./word");
const VALUES = process.env.EXTRA ? [6000, 7000, 8000, 9000, 11000, 12000, 13000, 14000, 15000, 16000, 17000, 18000, 19000, 20000, 23000, 24000, 30000, 31000, 1500, 2500, 12000, 1012, 2012, 5012, 22022, 14014, 112, 1013, 3333, 4444, 6666, 7777, 8888, 9999] : [0, ...Array.from({ length: 125 }, (_, i) => i + 1), 140, 150, 199, 200, 211, 222, 300, 400, 500, 600, 700, 800, 900, 999, 1000, 1001, 1002, 1100, 1999, 2000, 2001, 2024, 3000, 3999, 4000, 5000, 5555, 10000, 12345, 21000, 22000, 25000, 100000, 101000, 200000, 999999, 1000000, 32767, 32768];
(async () => {
  const res = {};
  for (const [fmt, lang] of (process.env.EXTRA ? [["cardinalText", "pl-PL"], ["ordinalText", "pl-PL"], ["ordinalText", "en-US"]] : [["cardinalText", "pl-PL"], ["ordinalText", "pl-PL"], ["cardinalText", "en-US"], ["ordinalText", "en-US"], ["ordinal", "en-US"], ["ordinal", "pl-PL"], ["upperRoman", "pl-PL"], ["lowerLetter", "pl-PL"], ["decimalZero", "pl-PL"], ["upperLetter", "pl-PL"]])) {
    const bytes = await buildDocx({
      lang,
      abstractNums: [{ levels: [{ fmt, text: "%1" }] }],
      nums: VALUES.map((v, i) => ({ id: i + 1, abs: 0, overrides: [{ ilvl: 0, start: v }] })),
      body: VALUES.flatMap((v, i) => [{ text: `v${v}`, numId: i + 1 }, { text: "-" }]),
    });
    const w = word.listStrings(bytes).filter((x) => x.text !== "-");
    res[`${fmt}/${lang}`] = Object.fromEntries(w.map((x) => [x.text.slice(1), x.label]));
    console.log(fmt, lang, w.slice(0, 3).map((x) => x.label), w.slice(-6).map((x) => x.label));
  }
  fs.writeFileSync(process.env.EXTRA ? "output/word-oracle/values-extra.json" : "output/word-oracle/values.json", JSON.stringify(res, null, 1));
})();
