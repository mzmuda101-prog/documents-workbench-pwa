// probe-before-top.js — odstęp „przed” akapitu na górze strony: kiedy Word go rysuje?
// Nagłówek z odstępem przed 320 tw (16 pt) na początku strony 2 w kilku wariantach; wynik: góra liter
// nagłówka na str. 2 (pt) i różnica względem wariantu bez odstępu (A4, margines górny 70,85 pt).
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const { buildDocx } = require("./build-docx");
const { withDocument, DIR } = require("./word");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
const H = (before, extra = "") => `<w:p><w:pPr>${extra}<w:spacing w:before="${before}" w:after="160"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="40"/></w:rPr><w:t>NAGLOWEK</w:t></w:r></w:p>`;
const T = (t) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
const fill = Array.from({ length: +(process.env.FILL || 31) }, (_, i) => T(`Wiersz wypelnienia ${i}`)).join("");
const CASES = {
  brOwnPara: (b) => [T("Strona 1"), '<w:p><w:r><w:br w:type="page"/></w:r></w:p>', H(b)],
  brInPara: (b) => [T("Strona 1"), `<w:p><w:pPr><w:spacing w:before="${b}" w:after="160"/></w:pPr><w:r><w:br w:type="page"/></w:r><w:r><w:rPr><w:b/><w:sz w:val="40"/></w:rPr><w:t>NAGLOWEK</w:t></w:r></w:p>`],
  pageBreakBefore: (b) => [T("Strona 1"), H(b, "<w:pageBreakBefore/>")],
  natural: (b) => [fill, H(b)],
  naturalSame: (b) => [fill, H(b).replace("<w:b/><w:sz w:val=\"40\"/>", "")],
  sectNextAfter0: (b) => [`<w:p><w:pPr><w:spacing w:after="0"/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr><w:r><w:t>Strona 1</w:t></w:r></w:p>`, H(b)],
  sectNextAfter20: (b) => [`<w:p><w:pPr><w:spacing w:after="400"/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr><w:r><w:t>Strona 1</w:t></w:r></w:p>`, H(b)],
  sectNext: (b) => [`<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr><w:r><w:t>Strona 1</w:t></w:r></w:p>`, H(b)],
};
if (!process.env.FIRST) (async () => {
  for (const [k, mk] of Object.entries(CASES)) {
    const y = {};
    for (const b of [0, 320]) {
      const pdf = path.join(DIR, `btop-${k}-${b}.pdf`);
      try { fs.unlinkSync(pdf); } catch (_) {}
      withDocument(await buildDocx({ body: mk(b) }), `save as it file name "${esc(pdf)}" file format format PDF`);
      const n = +execFileSync("pdfinfo", [pdf]).toString().match(/Pages:\s+(\d+)/)[1];
      for (let p = 1; p <= n; p++) {
        const xml = execFileSync("pdftotext", ["-f", String(p), "-l", String(p), "-bbox", pdf, "-"]).toString();
        const m = xml.match(/yMin="([\d.]+)"[^>]*>NAGLOWEK</);
        if (m) y[b] = `${p}:${(+m[1]).toFixed(2)}`;
      }
    }
    const d = (+y[320].split(":")[1] - +y[0].split(":")[1]).toFixed(2);
    console.log(k.padEnd(16), "bez odstępu", y[0], " z 16 pt", y[320], " różnica", d, "pt");
  }
})();
// FIRST=1: pierwszy akapit dokumentu (strona 1)
if (process.env.FIRST) (async () => {
  for (const b of [0, 320]) {
    const pdf = path.join(DIR, `btop-first-${b}.pdf`);
    try { fs.unlinkSync(pdf); } catch (_) {}
    withDocument(await buildDocx({ body: [H(b), T("dalej")] }), `save as it file name "${esc(pdf)}" file format format PDF`);
    const xml = execFileSync("pdftotext", ["-f", "1", "-l", "1", "-bbox", pdf, "-"]).toString();
    console.log("first", b, xml.match(/yMin="([\d.]+)"[^>]*>NAGLOWEK</)[1]);
  }
})();
