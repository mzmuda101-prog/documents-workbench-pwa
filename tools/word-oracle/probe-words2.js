const { buildDocx } = require("./build-docx");
const word = require("./word");
(async () => {
  const toks = [",", ".", "!", "?", "...", "…", "(", ")", "„", "”", "«", "»", "*", "•", "·", "#", "@", "&", "§", "|", "/", "-", "–", "—", "―", "‒", "‐", "−", ":", ";", "\"", "'", "~", "=", "+", "%", "0", "a—", "—a", "–a", "a––b", "a—-b"];
  // każdy token osobno: „x T y” — ile słów?
  const res = [];
  for (const t of toks) {
    const s = word.statistics(await buildDocx({ body: [`x ${t} y`] }));
    res.push(`${JSON.stringify(t)}=${s.words}`);
  }
  console.log(res.join("  "));
})();
