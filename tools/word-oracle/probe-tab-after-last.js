// probe-tab-after-last.js — gdzie Word stawia tabulator ZA ostatnim własnym tab-stopem akapitu
// (lista 3.11). PDF z Worda → pdftotext -bbox → x litery „C” od lewego marginesu.
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const { buildDocx } = require("./build-docx");
const { withDocument, DIR } = require("./word");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
const tab = (pos) => `<w:tabs><w:tab w:val="left" w:pos="${pos}"/></w:tabs>`;
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r><w:r><w:tab/></w:r>`;
const CASES = [
  { name: "stop 1985 tw, potem B4 C", pPr: tab(1985), raw: `<w:r><w:t>A</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>BBBB</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Cx</w:t></w:r>` },
  { name: "stop 1985 tw, wcięcie 567, potem B4 C", pPr: `${tab(1985)}<w:ind w:left="567"/>`, raw: `<w:r><w:t>A</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>BBBB</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Cx</w:t></w:r>` },
  { name: "dwa stopy 1000/3000, za drugim", pPr: `<w:tabs><w:tab w:val="left" w:pos="1000"/><w:tab w:val="left" w:pos="3000"/></w:tabs>`, raw: `<w:r><w:t>A</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>B</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>D</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Cx</w:t></w:r>` },
];
(async () => {
  for (const c of CASES) {
    const bytes = await buildDocx({ body: [`<w:p><w:pPr>${c.pPr}</w:pPr>${c.raw}</w:p>`] });
    const pdf = path.join(DIR, "_tab.pdf");
    try { fs.unlinkSync(pdf); } catch (_) {}
    withDocument(bytes, `save as it file name "${esc(pdf)}" file format format PDF`);
    const html = execFileSync("pdftotext", ["-bbox", pdf, "-"], { encoding: "utf8" });
    const words = [...html.matchAll(/<word xMin="([\d.]+)"[^>]*>([^<]+)<\/word>/g)].map((m) => [m[2], (+m[1] - 70.85).toFixed(2)]);
    console.log(c.name, JSON.stringify(words));
  }
})();
