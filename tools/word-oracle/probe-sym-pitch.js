// probe-sym-pitch.js — o ile symbol z kroju zastępczego podnosi linijkę (Arial 11 pt, interlinia 1,15,
// po 6 pt — jak „przewodnik”; FONT=Calibri itd.). 20 akapitów z symbolem; skok akapitów z PDF Worda vs bez symbolu.
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const { buildDocx } = require("./build-docx");
const { withDocument, DIR } = require("./word");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
const FONT = process.env.FONT || "Arial";
const CASES = { plain: "Ctrl+F", cmd: "Ctrl/⌘+F", undo: "↶ ↷ na pasku", swap: "Czytanie ⇄ Edycja", arrow: "a → b", focus: "przycisk ⛶", dots: "menu ⋯", left: "← wróć", home: "⌂ dom" };
const para = (t) => `<w:p><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
(async () => {
  for (const [k, txt] of Object.entries(CASES)) {
    const body = Array.from({ length: 21 }, (_, i) => para(`L${i} ${txt}`));
    const pdf = path.join(DIR, `sym-${k}.pdf`);
    try { fs.unlinkSync(pdf); } catch (_) {}
    withDocument(await buildDocx({ body }), `save as it file name "${esc(pdf)}" file format format PDF`);
    const xml = execFileSync("pdftotext", ["-f", "1", "-l", "1", "-bbox", pdf, "-"]).toString();
    const y = Object.fromEntries([...xml.matchAll(/yMin="([\d.]+)"[^>]*>(L\d+)</g)].map((m) => [m[2], +m[1]]));
    console.log(k.padEnd(6), "skok akapitu", ((y.L20 - y.L0) / 20).toFixed(3), "pt");
  }
})();
