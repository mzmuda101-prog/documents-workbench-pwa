// probe-tab-wrap.js — tabulator za ostatnim (prawym, z kropkami) tab-stopem, blisko prawego wcięcia:
// gdzie Word stawia następny znak (lista 3.11 / print-pages-word przypadek 4)
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const { buildDocx } = require("./build-docx");
const { withDocument, DIR } = require("./word");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
const t = (s) => `<w:t xml:space="preserve">${s}</w:t>`;
const CASES = [
  ["dt 708, prawy 5000 kropki, wcięcie P 4000", 708, `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="5000"/></w:tabs><w:ind w:right="4000"/>`],
  ["dt 720, prawy 5000 kropki, wcięcie P 4000", 720, `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="5000"/></w:tabs><w:ind w:right="4000"/>`],
  ["dt 720, prawy 5000 kropki, bez wcięcia", 720, `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="5000"/></w:tabs>`],
  ["dt 720, lewy 2000, tekst do 4980, bez wcięcia", 720, `<w:tabs><w:tab w:val="left" w:pos="2000"/></w:tabs>`],
  ["BEZ X: dt 720, prawy 5000 kropki, wcięcie P 4000", 720, `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="5000"/></w:tabs><w:ind w:right="4000"/>`],
  ["BEZ X: dt 708, prawy 5000 kropki, wcięcie P 4000", 708, `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="5000"/></w:tabs><w:ind w:right="4000"/>`],
];
(async () => {
  for (const [name, dt, pPr] of CASES) {
    const raw = name.includes("lewy") ? `<w:r>${t("A")}</w:r><w:r><w:tab/></w:r><w:r>${t("BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB")}</w:r><w:r><w:tab/></w:r><w:r>${t("X")}</w:r>`
      : `<w:r>${t("Miejscowość:")}</w:r><w:r><w:tab/></w:r><w:r>${t("Teren nr:")}</w:r><w:r><w:tab/></w:r>${name.startsWith("BEZ X") ? "" : `<w:r>${t("X")}</w:r>`}`;
    const bytes = await buildDocx({ body: [`<w:p><w:pPr>${pPr}</w:pPr>${raw}</w:p>`], settings: "" });
    // defaultTabStop w build-docx = 708; dla 720 podmieniamy w settings
    const JSZip = require("jszip");
    const z = await JSZip.loadAsync(bytes);
    const st = (await z.file("word/settings.xml").async("string")).replace('w:defaultTabStop w:val="708"', `w:defaultTabStop w:val="${dt}"`);
    z.file("word/settings.xml", st, { date: new Date("2026-01-01T00:00:00Z") });
    const b2 = await z.generateAsync({ type: "nodebuffer" });
    const pdf = path.join(DIR, "_tabw.pdf");
    try { fs.unlinkSync(pdf); } catch (_) {}
    withDocument(b2, `save as it file name "${esc(pdf)}" file format format PDF`);
    const html = execFileSync("pdftotext", ["-bbox", pdf, "-"], { encoding: "utf8" });
    const words = [...html.matchAll(/<word xMin="([\d.]+)" yMin="([\d.]+)"[^>]*xMax="([\d.]+)"[^>]*>([^<]+)<\/word>/g)].map((m) => `${m[4].slice(0, 12)}@${(+m[1] - 70.85).toFixed(1)}..${(+m[3] - 70.85).toFixed(1)},y${(+m[2]).toFixed(0)}`);
    console.log(name, "|", words.join("  "));
  }
})();
