// run-keys.js — tabela klawiszy (scripts/keys-table.js) w PRAWDZIWYM Wordzie: ten sam plik
// (konteksty z `body`), kursor, klawisz, zapis — i ten sam odczyt stanu co test aplikacji
// (scripts/keys-state.js), więc wynik Worda wpisuje się wprost w tabelę.
//   node tools/word-oracle/run-keys.js [kontekst]
// Klawisze przez AppleScript: Backspace = type backspace, Enter = type paragraph; Delete — zaznaczenie
// jednego znaku za kursorem + Backspace (przybliżenie klawisza Delete).
// Word liczy znacznik końca wiersza tabeli jako osobny „akapit” — pomijamy go przy liczeniu
// (akapit „\r\x07” zaraz po akapicie zakończonym \x07 = koniec wiersza).
const pw = require("playwright");
const { buildDocx } = require("./build-docx");
const word = require("./word");
const { contexts, expect } = require("../../scripts/keys-table");
const { keysStateIn } = require("../../scripts/keys-state");

const SEL = '(get selection of application "Microsoft Word")';
// AppleScript: treść każdego akapitu Worda (do pominięcia końców wierszy) — „początek|treść” na linię
const LIST = `
    repeat with i from 1 to count of paragraphs
      set res to res & (start of content of text object of paragraph i) & "|" & (content of text object of paragraph i) & (ASCII character 30)
    end repeat`;
const wordParas = (out) => out.split("\x1e").filter((x) => x.includes("|")).map((x) => { const k = x.indexOf("|"); return { start: Number(x.slice(0, k)), text: x.slice(k + 1) }; });
// akapity w zapisie stanu: bez końców wierszy (indeksy Worda od 1)
function realParas(list) {
  const out = [];
  list.forEach((p, i) => {
    const rowEnd = p.text === "\r\x07" && i > 0 && list[i - 1].text.endsWith("\x07");
    if (!rowEnd) out.push({ ...p, word: i + 1 });
  });
  return out;
}
const KEY = {
  bs: [false, `type backspace ${SEL}`],
  enterStart: [false, `type paragraph ${SEL}`],
  enterEnd: [true, `type paragraph ${SEL}`],
  enter: [false, `type paragraph ${SEL}`],
  del: [true, `set s to selection start of ${SEL}
    select (create range it start s end s + 1)
    type backspace ${SEL}`],
};

(async () => {
  const only = process.argv[2];
  const browser = await pw.chromium.launch();
  const page = await browser.newPage();
  for (const [name, ctx] of Object.entries(contexts)) {
    if (!ctx.body || (only && name !== only)) continue;
    const bytes = await buildDocx({ body: ctx.body });
    for (const key of Object.keys(expect[name] || {})) {
      const [atEnd, press] = KEY[key];
      try {
        const before = realParas(wordParas(word.withDocument(bytes, LIST)));
        const target = before[ctx.at];
        const r = word.saveDocx(bytes, `set r to text object of paragraph ${target.word}
    set s to ${atEnd ? "(end of content of r) - 1" : "start of content of r"}
    select (create range it start s end s)
    ${press}
    set res to (selection start of ${SEL} as string) & (ASCII character 31)
    ${LIST}`);
        const [sel, rest] = r.res.split("\x1f");
        const after = realParas(wordParas(rest));
        let pi = 0;
        after.forEach((p, i) => { if (p.start <= Number(sel)) pi = i; });
        const st = after.map((p) => p.start);
        const state = await keysStateIn(page, r.parts["word/document.xml"]);
        const app = expect[name][key];
        console.log(`${name} · ${key}\n  Word: ${state} @${pi}:${Number(sel) - st[pi]}\n  tabela: ${app}`);
      } catch (e) {
        console.log(`${name} · ${key}\n  BŁĄD Worda: ${e.message}`);
      }
    }
  }
  await browser.close();
})();
