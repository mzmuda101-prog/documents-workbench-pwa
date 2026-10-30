// print-pages-word-playwright.js — strony wydruku jak w Wordzie (2026-10-30, korpus word:compare:
// „ubezp”, „karty”, „S-99”).
//
// (1) Pola w akapicie (vendor-libs łatka 21): wynik PAGE / NUMPAGES liczony na każdej kartce
//     (dawniej wartość zapisana w pliku — „1 / 4” na każdej stronie), <w:fldSimple> w akapicie
//     nie znika.
// (2) Inna pierwsza strona (w:titlePg, łatka 23): pierwsza kartka ma stopkę „first”, dalsze —
//     domyślną (dawniej „first” na wszystkich: brak numeru stron w „S-99”).
// (3) Wiersz tabeli wyższy niż strona: Word dzieli go KAŻDĄ KOMÓRKĘ OSOBNO między linijkami,
//     dalszy ciąg każdej komórki zaczyna się u góry następnej strony (dawniej jedna kreska przez
//     cały wiersz — „karty”: druga kolumna 70 px wyżej niż w Wordzie, linijki przecięte).
// (4) Tabulator za ostatnim tabulatorem akapitu, gdy domyślny wypada za prawym wcięciem: nowa
//     linijka z kropkami do tego samego tabulatora („Miejscowość: …… Teren nr:⇥”).
// (5) w:widowControl / keepNext / keepLines wpisane w akapit trafiają do podglądu (łatka 22).
// (6) Wyjustowany akapit łamie się jak w Wordzie: spacja może się ścisnąć o ¼ (wyrocznia Worda,
//     output/word-oracle/jx-arial.docx: Arial 10,5 pt, wcięcie z prawej 25 pt — „feugiat” jeszcze
//     w 1. linijce; ten sam akapit do lewej — już w 2.).
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const t = (s) => `<w:t xml:space="preserve">${s}</w:t>`;
const P = (s, pPr = "") => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}<w:r>${t(s)}</w:r></w:p>`;
const fld = (instr, cached) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r>${t(cached)}</w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

async function fixture() {
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
  z.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
  z.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rIdF1" Type="${REL}/footer" Target="footer1.xml"/><Relationship Id="rIdF2" Type="${REL}/footer" Target="footer2.xml"/></Relationships>`);
  z.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`);
  z.file("word/footer1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NS}><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${t("Strona ")}</w:r>${fld("PAGE", "9")}<w:r>${t(" z ")}</w:r>${fld("NUMPAGES", "9")}</w:p></w:ftr>`);
  z.file("word/footer2.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NS}><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${t("Pierwsza strona")}</w:r></w:p></w:ftr>`);
  // wiersz tabeli wyższy niż strona: lewa komórka 70 akapitów, prawa 50 (inne wysokości linijek)
  const cell = (n, w, sz, tag) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${Array.from({ length: n }, (_, i) => `<w:p><w:r><w:rPr><w:sz w:val="${sz}"/></w:rPr>${t(`${tag}${i + 1}`)}</w:r></w:p>`).join("")}</w:tc>`;
  const table = `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr>${cell(70, 4500, 20, "L")}${cell(50, 4500, 28, "R")}</w:tr></w:tbl>`;
  const body = [
    P("Pierwsza strona dokumentu"),
    `<w:p><w:r>${t("Pole proste: ")}</w:r><w:fldSimple w:instr=" PAGE "><w:r>${t("7")}</w:r></w:fldSimple></w:p>`,
    P("Akapit bez kontroli bękartów", '<w:widowControl w:val="0"/><w:keepNext/>'),
    `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="5000"/></w:tabs><w:ind w:right="4000"/></w:pPr><w:r>${t("Miejscowość:")}</w:r><w:r><w:tab/></w:r><w:r>${t("Teren nr:")}</w:r><w:r><w:tab/></w:r></w:p>`,
    ...["both", "left"].map((jc) => `<w:p><w:pPr><w:ind w:right="500"/><w:jc w:val="${jc}"/></w:pPr><w:r><w:rPr><w:sz w:val="21"/></w:rPr>${t(`${jc[0]}25 fringilla magna at augue maximus finibus. Donec mi dui, volutpat ut lectus sit amet, feugiat tincidunt ex. Sed id justo lobortis, eleifend lorem vitae.`)}</w:r></w:p>`),
    `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`,
    table,
    P("Za tabelą"),
  ].join("");
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr><w:footerReference w:type="default" r:id="rIdF1"/><w:footerReference w:type="first" r:id="rIdF2"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/><w:titlePg/></w:sectPr></w:body></w:document>`);
  return z.generateAsync({ type: "nodebuffer" });
}

(async () => {
  const file = path.join(os.tmpdir(), `dwb-print-pages-${process.pid}.docx`);
  fs.writeFileSync(file, await fixture());
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles(file);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host section.docx p"), null, { timeout: 60000 });
  await page.waitForTimeout(800);

  const view = await page.evaluate(() => {
    const host = document.querySelector(".docx-preview-host");
    const simple = [...host.querySelectorAll("p")].find((p) => p.textContent.startsWith("Pole proste"));
    const widow = [...host.querySelectorAll("p")].find((p) => p.textContent.startsWith("Akapit bez kontroli"));
    const dots = [...host.querySelectorAll("p")].find((p) => p.textContent.startsWith("Miejscowość"));
    const tabs = dots ? [...dots.querySelectorAll("span.docx-tab")].map((s) => { const r = s.getBoundingClientRect(); return { top: r.top, w: r.width }; }) : [];
    return {
      simple: simple && simple.querySelector("[data-dwb-field='PAGE']")?.textContent,
      widow: widow && [widow.dataset.dwbWidow, widow.dataset.dwbKeepNext],
      tabs,
      alts: host.querySelectorAll(".dwb-hf-alt").length,
      // ostatnie słowo 1. linijki akapitów „b25 …” (wyjustowany) i „l25 …” (do lewej)
      firstLine: ["b25 ", "l25 "].map((tag) => {
        const p = [...host.querySelectorAll("p")].find((x) => x.textContent.startsWith(tag));
        if (!p) return null;
        const r = document.createRange();
        const tn = p.querySelector("span")?.firstChild;
        if (!tn) return null;
        let top = null, last = "";
        for (const m of tn.data.matchAll(/\S+/g)) {
          r.setStart(tn, m.index); r.setEnd(tn, m.index + m[0].length);
          const y = r.getBoundingClientRect().top;
          if (top == null) top = y;
          if (y > top + 3) break;
          last = m[0];
        }
        return last;
      }),
    };
  });
  check("pole proste (w:fldSimple) w akapicie widać, z oznaczeniem PAGE", view.simple === "7", JSON.stringify(view.simple));
  check("w:widowControl=0 i w:keepNext akapitu w podglądzie (data-dwb-widow / -keep-next)", JSON.stringify(view.widow) === '["0","1"]', JSON.stringify(view.widow));
  check("tabulator za ostatnim: nowa linijka kropek (niżej, szeroka)", view.tabs.length === 2 && view.tabs[1].top > view.tabs[0].top + 5 && view.tabs[1].w > 150, JSON.stringify(view.tabs));
  check("podgląd (nie wydruk) bez ukrytych wersji stopki", view.alts === 0, String(view.alts));
  check("wyjustowany akapit łamie się jak Word (spacje ściśnięte: „feugiat” w 1. linijce), do lewej — bez ściskania", JSON.stringify(view.firstLine) === '["feugiat","amet,"]', JSON.stringify(view.firstLine));

  await page.evaluate(() => dwbPrint.open());
  await page.waitForSelector(".pp-sheet", { timeout: 30000 });
  await page.evaluate(async () => { const st = dwbPrint._state(); if (st) { st.fit = false; st.zoom = 1; } document.querySelector(".pp-pages").style.setProperty("--pp-zoom", "1"); await document.fonts.ready; });
  await page.waitForTimeout(300);
  const pp = await page.evaluate(() => {
    const sheets = [...document.querySelectorAll(".pp-sheet")];
    const lines = (sh) => {
      const s0 = sh.getBoundingClientRect();
      const clip = sh.querySelector(".pp-clip").getBoundingClientRect();
      return [...sh.querySelectorAll(".pp-clip td p")].map((p) => { const r = p.getBoundingClientRect(); return { t: p.textContent, top: r.top - s0.top, bottom: r.bottom - s0.top }; })
        .filter((x) => x.bottom > clip.top - s0.top + 0.5 && x.top < clip.bottom - s0.top - 0.5).map((x) => ({ ...x, cut: x.top < clip.top - s0.top - 1.5 || x.bottom > clip.bottom - s0.top + 1.5, clipTop: clip.top - s0.top }));
    };
    return {
      n: sheets.length,
      footers: sheets.map((sh) => [...sh.querySelectorAll("footer.pp-hf")].map((f) => f.textContent.trim()).join("|")),
      table: sheets.slice(1, 3).map(lines),
    };
  });
  const n = pp.n;
  check(`wydruk: pierwsza kartka ze stopką „first”, dalsze z numerem (${n} kartek)`, n >= 3 && pp.footers[0] === "Pierwsza strona" && pp.footers.slice(1).every((f, i) => f === `Strona ${i + 2} z ${n}`), JSON.stringify(pp.footers));
  const [s2, s3] = pp.table;
  const cut = [...(s2 || []), ...(s3 || [])].filter((x) => x.cut);
  check("wiersz wyższy niż strona: żadna linijka komórki nie przecięta granicą kartki", s2?.length && s3?.length && !cut.length, JSON.stringify(cut.slice(0, 3)));
  const firstL = s3?.find((x) => /^L/.test(x.t)), firstR = s3?.find((x) => /^R/.test(x.t));
  const lastL2 = [...(s2 || [])].reverse().find((x) => /^L/.test(x.t)), lastR2 = [...(s2 || [])].reverse().find((x) => /^R/.test(x.t));
  check("dalszy ciąg obu komórek u góry następnej kartki (jak w Wordzie)", firstL && firstR && Math.abs(firstL.top - firstL.clipTop) < 4 && Math.abs(firstR.top - firstR.clipTop) < 4, JSON.stringify({ firstL, firstR }));
  check("komórki ciągną się bez przerwy (kolejne numery na granicy)", lastL2 && firstL && +firstL.t.slice(1) === +lastL2.t.slice(1) + 1 && lastR2 && firstR && +firstR.t.slice(1) === +lastR2.t.slice(1) + 1, JSON.stringify({ lastL2: lastL2?.t, firstL: firstL?.t, lastR2: lastR2?.t, firstR: firstR?.t }));
  if (process.env.DEBUG) console.log(JSON.stringify(pp.table.map((s) => s.map((x) => x.t + "@" + Math.round(x.top) + (x.cut ? "!" : ""))), null, 0));
  check("brak błędów strony", errors.length === 0, errors.join("\n"));

  await browser.close();
  fs.rmSync(file, { force: true });
  const bad = results.filter((r) => !r.ok);
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `\n   ${r.detail}`}`);
  console.log(`\n[${ENGINE}] ${results.length - bad.length}/${results.length}`);
  process.exit(bad.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
