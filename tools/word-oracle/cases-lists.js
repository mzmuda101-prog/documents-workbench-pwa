// Przypadki numeracji list: to, co Word pokazuje jako numer (list string), kontra aplikacja.
const items = (n, numId, ilvl = 0, pre = "poz") => Array.from({ length: n }, (_, i) => ({ text: `${pre} ${i + 1}`, numId, ilvl }));
const one = (fmt, text, extra = {}) => ({ abstractNums: [{ levels: [{ fmt, text, ...extra }] }], nums: [{ id: 1, abs: 0 }] });

module.exports = [
  { name: "litery-ponad-26", ...one("lowerLetter", "%1)"), body: items(56, 1) },
  { name: "WIELKIE-litery-ponad-26", ...one("upperLetter", "%1."), body: items(30, 1) },
  { name: "rzymskie-duze", ...one("upperRoman", "%1.", { start: 45 }), body: items(8, 1) },
  { name: "rzymskie-male-start", ...one("lowerRoman", "(%1)", { start: 1 }), body: items(12, 1) },
  { name: "zero-wiodace", ...one("decimalZero", "%1."), body: items(12, 1) },
  { name: "start-zero", ...one("decimal", "%1.", { start: 0 }), body: items(3, 1) },
  { name: "litera-start-zero", ...one("lowerLetter", "%1)", { start: 0 }), body: items(3, 1) },
  { name: "rzymskie-start-zero", ...one("upperRoman", "%1.", { start: 0 }), body: items(3, 1) },
  { name: "porzadkowe-pl", ...one("ordinal", "%1"), body: items(12, 1).concat([{ text: "x" }]).concat(items(3, 1, 0, "en").map((p) => ({ ...p, lang: "en-US" }))) },
  { name: "slownie-pl", ...one("cardinalText", "%1"), body: items(25, 1) },
  { name: "slownie-porzadkowe-pl", ...one("ordinalText", "%1"), body: items(25, 1) },
  { name: "slownie-en", lang: "en-US", ...one("cardinalText", "%1"), body: items(22, 1) },
  { name: "porzadkowe-en", lang: "en-US", ...one("ordinal", "%1"), body: items(24, 1) },
  { name: "tekst-staly", ...one("decimal", "Art. %1 –"), body: items(3, 1) },
  { name: "bez-numeru", ...one("decimal", ""), body: items(2, 1) },
  {
    name: "wielopoziomowa",
    abstractNums: [{ multi: "multilevel", levels: [{ ilvl: 0, fmt: "decimal", text: "%1." }, { ilvl: 1, fmt: "decimal", text: "%1.%2." }, { ilvl: 2, fmt: "decimal", text: "%1.%2.%3." }] }],
    nums: [{ id: 1, abs: 0 }],
    body: [
      { text: "A", numId: 1, ilvl: 0 }, { text: "A1", numId: 1, ilvl: 1 }, { text: "A1a", numId: 1, ilvl: 2 }, { text: "A1b", numId: 1, ilvl: 2 },
      { text: "A2", numId: 1, ilvl: 1 }, { text: "A2a", numId: 1, ilvl: 2 }, { text: "B", numId: 1, ilvl: 0 }, { text: "B-skok", numId: 1, ilvl: 2 },
      { text: "B1", numId: 1, ilvl: 1 }, { text: "C", numId: 1, ilvl: 0 },
    ],
  },
  {
    name: "wielopoziomowa-bez-poziomu-0",
    abstractNums: [{ multi: "multilevel", levels: [{ ilvl: 0, fmt: "decimal", text: "%1." }, { ilvl: 1, fmt: "lowerLetter", text: "%1.%2)" }] }],
    nums: [{ id: 1, abs: 0 }],
    body: [{ text: "a?", numId: 1, ilvl: 1 }, { text: "b?", numId: 1, ilvl: 1 }, { text: "1", numId: 1, ilvl: 0 }, { text: "1a", numId: 1, ilvl: 1 }],
  },
  {
    name: "numeracja-prawnicza-isLgl",
    abstractNums: [{ multi: "multilevel", levels: [{ ilvl: 0, fmt: "upperRoman", text: "%1." }, { ilvl: 1, fmt: "lowerLetter", text: "%1.%2.", isLgl: true }] }],
    nums: [{ id: 1, abs: 0 }],
    body: [{ text: "I", numId: 1, ilvl: 0 }, { text: "I.a", numId: 1, ilvl: 1 }, { text: "I.b", numId: 1, ilvl: 1 }, { text: "II", numId: 1, ilvl: 0 }, { text: "II.a", numId: 1, ilvl: 1 }],
  },
  {
    name: "bez-restartu-lvlRestart0",
    abstractNums: [{ multi: "multilevel", levels: [{ ilvl: 0, fmt: "decimal", text: "%1." }, { ilvl: 1, fmt: "lowerLetter", text: "%2)", restart: 0 }] }],
    nums: [{ id: 1, abs: 0 }],
    body: [{ text: "1", numId: 1, ilvl: 0 }, { text: "a", numId: 1, ilvl: 1 }, { text: "b", numId: 1, ilvl: 1 }, { text: "2", numId: 1, ilvl: 0 }, { text: "c?", numId: 1, ilvl: 1 }],
  },
  {
    name: "restart-po-poziomie-1",
    abstractNums: [{ multi: "multilevel", levels: [{ ilvl: 0, fmt: "decimal", text: "%1." }, { ilvl: 1, fmt: "decimal", text: "%1.%2." }, { ilvl: 2, fmt: "lowerRoman", text: "%3.", restart: 1 }] }],
    nums: [{ id: 1, abs: 0 }],
    body: [{ text: "1", numId: 1, ilvl: 0 }, { text: "1.1", numId: 1, ilvl: 1 }, { text: "i", numId: 1, ilvl: 2 }, { text: "ii", numId: 1, ilvl: 2 }, { text: "1.2", numId: 1, ilvl: 1 }, { text: "i?", numId: 1, ilvl: 2 }, { text: "2", numId: 1, ilvl: 0 }, { text: "?", numId: 1, ilvl: 2 }],
  },
  {
    name: "dwa-num-jedna-definicja",
    abstractNums: [{ levels: [{ fmt: "decimal", text: "%1." }] }],
    nums: [{ id: 1, abs: 0 }, { id: 2, abs: 0 }, { id: 3, abs: 0, overrides: [{ ilvl: 0, start: 1 }] }, { id: 4, abs: 0, overrides: [{ ilvl: 0, start: 7 }] }],
    body: [...items(2, 1), { text: "przerwa" }, ...items(2, 2, 0, "num2"), ...items(2, 1, 0, "znów1"), ...items(2, 3, 0, "num3-restart"), ...items(2, 4, 0, "num4-od7"), ...items(1, 2, 0, "num2-znów")],
  },
  {
    name: "numId-0-wylacza",
    abstractNums: [{ levels: [{ fmt: "decimal", text: "%1." }] }],
    nums: [{ id: 1, abs: 0 }],
    body: [...items(2, 1), { text: "bez", numId: 0 }, ...items(2, 1, 0, "dalej")],
  },
  {
    name: "przerwa-zwykly-akapit",
    abstractNums: [{ levels: [{ fmt: "decimal", text: "%1)" }] }],
    nums: [{ id: 1, abs: 0 }],
    body: [...items(2, 1), { text: "Akapit między" }, { text: "" }, ...items(2, 1, 0, "dalej")],
  },
  {
    name: "nadpisany-poziom-lvlOverride",
    abstractNums: [{ levels: [{ fmt: "decimal", text: "%1." }] }],
    nums: [{ id: 1, abs: 0 }, { id: 2, abs: 0, overrides: [{ ilvl: 0, lvl: { fmt: "upperLetter", text: "%1)", start: 3 } }] }],
    body: [...items(2, 1), ...items(3, 2, 0, "nadp")],
  },
  {
    name: "naglowki-numerowane-stylem",
    styles: '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:numPr><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>',
    abstractNums: [{ multi: "multilevel", levels: [{ ilvl: 0, fmt: "decimal", text: "%1.", pStyle: "Heading1" }, { ilvl: 1, fmt: "decimal", text: "%1.%2.", pStyle: "Heading2" }] }],
    nums: [{ id: 1, abs: 0 }],
    body: [{ text: "Rozdział", style: "Heading1" }, { text: "Pod", style: "Heading2" }, { text: "Treść" }, { text: "Pod 2", style: "Heading2" }, { text: "Rozdział 2", style: "Heading1" }, { text: "Pod", style: "Heading2" }, { text: "Bez numeru", style: "Heading1", pPr: "" }],
  },
  {
    name: "litery-dwupoziomowe",
    abstractNums: [{ multi: "multilevel", levels: [{ ilvl: 0, fmt: "upperLetter", text: "%1." }, { ilvl: 1, fmt: "lowerRoman", text: "%1.%2" }] }],
    nums: [{ id: 1, abs: 0 }],
    body: [{ text: "A", numId: 1 }, { text: "A.i", numId: 1, ilvl: 1 }, { text: "A.ii", numId: 1, ilvl: 1 }, { text: "B", numId: 1 }, { text: "B.i", numId: 1, ilvl: 1 }],
  },
];
