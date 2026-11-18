// probe-para-group.js — sąsiednie akapity z tą samą ramką (w:pBdr) w Wordzie: jedno pole czy osobne?
// Trzy akapity „A1..A3” z ramką; wynik: odległości PRZED→A1, A1→A2, A2→A3, A3→PO (góra liter, pt).
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const { buildDocx } = require("./build-docx");
const { withDocument, DIR } = require("./word");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
const T = (t) => `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
const side = (n, sz, sp) => `<w:${n} w:val="single" w:sz="${sz}" w:space="${sp}" w:color="000000"/>`;
const box = (sz, sp, between = "") => `<w:pBdr>${side("top", sz, sp)}${side("left", sz, 4)}${side("bottom", sz, sp)}${side("right", sz, 4)}${between}</w:pBdr>`;
const P = (t, bdr, ind = "") => `<w:p><w:pPr>${bdr}<w:spacing w:after="0"/>${ind}</w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
const CASES = {
  plain: () => [T("PRZED"), P("A1", ""), P("A2", ""), P("A3", ""), T("PO")],
  group: () => { const b = box(4, 1); return [T("PRZED"), P("A1", b), P("A2", b), P("A3", b), T("PO")]; },
  groupThick: () => { const b = box(34, 1); return [T("PRZED"), P("A1", b), P("A2", b), P("A3", b), T("PO")]; },
  groupBetween: () => { const b = box(4, 1, side("between", 4, 1)); return [T("PRZED"), P("A1", b), P("A2", b), P("A3", b), T("PO")]; },
  diffIndent: () => { const b = box(4, 1); return [T("PRZED"), P("A1", b), P("A2", b, '<w:ind w:left="720"/>'), P("A3", b), T("PO")]; },
  topOnly: () => { const b = `<w:pBdr>${side("top", 4, 1)}</w:pBdr>`; return [T("PRZED"), P("A1", b), P("A2", b), P("A3", b), T("PO")]; },
};
(async () => {
  for (const [k, mk] of Object.entries(CASES)) {
    const pdf = path.join(DIR, `pgrp-${k}.pdf`);
    try { fs.unlinkSync(pdf); } catch (_) {}
    withDocument(await buildDocx({ body: mk() }), `save as it file name "${esc(pdf)}" file format format PDF`);
    const xml = execFileSync("pdftotext", ["-f", "1", "-l", "1", "-bbox", pdf, "-"]).toString();
    const y = (w) => +xml.match(new RegExp(`yMin="([\\d.]+)"[^>]*>${w}<`))[1];
    const s = ["PRZED", "A1", "A2", "A3", "PO"].map(y);
    console.log(k.padEnd(13), s.slice(1).map((v, i) => (v - s[i]).toFixed(2)).join("  "));
  }
})();
