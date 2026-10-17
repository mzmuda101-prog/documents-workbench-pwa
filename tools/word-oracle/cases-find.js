// Znajdź i zamień: przypadki (każdy = osobny dokument, „Zamień wszystko”).
// body: akapity (tekst albo { raw } z fragmentami), find/replace, mc = uwzględnij wielkość liter,
// ww = tylko całe wyrazy.
const R = (rpr, text) => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const CASES = [
  { name: "wielkosc-liter-wyl", body: ["Kot i kot oraz KOT, koty."], find: "kot", replace: "pies" },
  { name: "wielkosc-liter-wl", body: ["Kot i kot oraz KOT, koty."], find: "kot", replace: "pies", mc: true },
  { name: "mieszana-wielkosc", body: ["kOT KoT kot"], find: "kot", replace: "pies" },
  { name: "zamiana-z-wielka", body: ["ala Ala ALA"], find: "ala", replace: "Ewa" },
  { name: "szukane-z-wielkiej", body: ["ala Ala ALA"], find: "Ala", replace: "ewa" },
  { name: "dwa-slowa", body: ["Dzień dobry, DZIEŃ DOBRY i dzień Dobry."], find: "dzień dobry", replace: "cześć" },
  { name: "jedna-litera", body: ["I tak i nie, I."], find: "i", replace: "oraz", ww: true },
  { name: "cale-wyrazy", body: ["kot kotek kot-ek kot_ek kot's 2kot kot2 (kot) „kot” kot."], find: "kot", replace: "pies", ww: true },
  { name: "cale-wyrazy-ze-spacja", body: ["ma kota, ma kot."], find: "ma kot", replace: "X", ww: true },
  { name: "polskie-litery", body: ["Łódź łódź ŁÓDŹ łodzi"], find: "łódź", replace: "miasto" },
  { name: "polskie-cale-wyrazy", body: ["żółw żółwik pół-żółw ŻÓŁW Żółw."], find: "żółw", replace: "jeż", ww: true },
  { name: "polskie-wielkosc-wl", body: ["Ćma ćma ĆMA"], find: "ćma", replace: "motyl", mc: true },
  { name: "cudzyslowy-drukarskie", body: ["„cytat” \"cytat\" “cytat”"], find: "\"cytat\"", replace: "X" },
  { name: "apostrof-drukarski", body: ["don’t don't"], find: "don't", replace: "do" },
  { name: "twarda-spacja", body: ["10 000 i 10 000"], find: "10 000", replace: "dużo" },
  { name: "miekki-dywiz", body: ["ko­t i kot"], find: "kot", replace: "pies" },
  { name: "usuwanie", body: ["a  b  c"], find: "  ", replace: "" },
  { name: "format-poczatku", body: [{ raw: R("", "Ala ma ") + R("<w:b/>", "kota") + R("", " i psa.") }], find: "ma kota", replace: "lubi psa" },
  { name: "format-w-srodku", body: [{ raw: R("<w:b/>", "Ala") + R("", " ma kota.") }], find: "a ma", replace: "X" },
  { name: "format-calego", body: [{ raw: R("", "Tekst ") + R("<w:i/>", "kot") + R("", " tu.") }], find: "kot", replace: "pies" },
  { name: "format-przez-trzy", body: [{ raw: R("<w:b/>", "ab") + R("<w:i/>", "cd") + R("<w:u w:val=\"single\"/>", "ef") }], find: "bcde", replace: "XY" },
  { name: "dluzsza-zamiana-wielka", body: ["KOT."], find: "kot", replace: "pies burek" },
];
module.exports = { CASES };
