// pdf-bands-playwright.js — PDF → DOCX: harmonogram z Excela (szare pasy-wiersze).
// 1. „SOBOTA      godz.” — szeroki odstęp (kilka spacji ze źródła) staje się tabulatorem na pozycji
//    z PDF: „godz.” dokładnie tam, gdzie w PDF; dawniej zwijał się do jednej spacji i kolumna „godz.”
//    się rozjeżdżała (odstęp jak w źródle: 15 pt przy 11 pt = 1,4 wysokości litery — poniżej
//    dawnego progu tabulatora 1,6).
// 2. Puste szare pasy w serii (ostatnie wiersze bez wpisów) — puste wiersze tabeli z tym samym
//    tłem, nie obraz pod tekstem (dawniej nieedytowalne i znikały w widoku telefonu).
// 3. Widok telefonu: pasy na całą szerokość (dawniej kurczyły się do tekstu, pusty = kreska).
// PDF budowany tu, w teście (pliki .pdf są w .gitignore).
//   node scripts/pdf-bands-playwright.js / ENGINE=webkit node scripts/pdf-bands-playwright.js

const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";

function buildPdf(file) {
  const band = (y) => `0.878 0.878 0.878 rg 52 ${y} 490.8 37.8 re f`;
  const text = (x, y, s, font) => `BT /${font} 11 Tf 0 0 0 rg ${x} ${y} Td (${s}) Tj ET`;
  const parts = [];
  [660, 603.6, 547.2, 490.8].forEach((y, i) => { // dwa z tekstem, dwa puste
    parts.push(band(y));
    if (i < 2) parts.push(text(54, y + 22, `${7 + i * 7} listopad 2026`, "F1"), text(54, y + 6, "SOBOTA ", "F2"), text(116.4, y + 6, "godz. 10.30", "F2"));
  });
  const content = parts.join("\n");
  const objs = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.2 841.8] /Resources << /Font << /F1 4 0 R /F2 5 0 R >> >> /Contents 6 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>",
    `<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}\nendstream`,
  ];
  let out = "%PDF-1.4\n";
  const offs = [];
  objs.forEach((o, i) => { offs.push(Buffer.byteLength(out)); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const xref = Buffer.byteLength(out);
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n${offs.map((o) => `${String(o).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  fs.writeFileSync(file, out);
}

(async () => {
  const file = path.join(os.tmpdir(), `dwb-pdf-bands-${process.pid}.pdf`);
  buildPdf(file);
  const browser = await pw[ENGINE].launch({ headless: true });
  const fails = [];
  const errors = [];
  try {
    for (const width of [1280, 390]) {
      const context = await browser.newContext({ serviceWorkers: "block", viewport: { width, height: 900 } });
      await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
      const page = await context.newPage();
      page.on("pageerror", (e) => errors.push(e.message));
      await page.goto(APP_URL, { waitUntil: "load" });
      await page.evaluate(() => document.getElementById("heroSplash")?.remove());
      await page.setInputFiles("#fileInput", file);
      await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host table"), null, { timeout: 60000 });
      await page.waitForTimeout(500);
      const r = await page.evaluate(() => {
        const host = document.querySelector(".docx-preview-host");
        const sec = host.querySelector("section.docx").getBoundingClientRect();
        const tables = [...host.querySelectorAll("section.docx table")].map((t) => {
          const td = t.querySelector("td");
          // położenie „godz.” w pt od lewej krawędzi kartki (kartka = 595,2 pt)
          let godz = null;
          const tw = document.createTreeWalker(t, NodeFilter.SHOW_TEXT);
          for (let n = tw.nextNode(); n && godz == null; n = tw.nextNode()) {
            const i = n.data.indexOf("godz.");
            if (i < 0) continue;
            const rg = document.createRange();
            rg.setStart(n, i);
            rg.setEnd(n, i + 1);
            godz = ((rg.getBoundingClientRect().left - sec.left) * 595.2) / sec.width;
          }
          return { text: t.textContent, godz, w: t.getBoundingClientRect().width / sec.width, h: t.getBoundingClientRect().height, bg: getComputedStyle(td).backgroundColor };
        });
        return { tables, images: host.querySelectorAll("section.docx img").length, scrollW: document.documentElement.scrollWidth, vw: innerWidth };
      });
      const tag = `[${width} px]`;
      // „godz.” jak w PDF (116,4 pt) — tylko widok strony; na telefonie wiersz się przelicza
      const at = r.tables.filter((t) => t.godz != null).map((t) => t.godz);
      if (at.length !== 2 || (width > 600 && at.some((x) => Math.abs(x - 116.4) > 1.5)) || at.some((x) => x < 100)) fails.push(`${tag} „godz.” na pozycji z PDF (116,4 pt) w obu wierszach, jest: ${at.map((x) => x.toFixed(1)).join(", ")} — ${JSON.stringify(r.tables.map((t) => t.text))}`);
      const grey = r.tables.filter((t) => t.bg === "rgb(224, 224, 224)");
      const empty = grey.filter((t) => !t.text.trim());
      if (grey.length !== 4 || empty.length !== 2) fails.push(`${tag} 4 szare wiersze (2 puste) jako tabele, jest: ${grey.length} (${empty.length} pustych)`);
      if (r.images) fails.push(`${tag} pasy jako obraz pod tekstem (${r.images} obraz)`);
      if (grey.some((t) => t.w < 0.8)) fails.push(`${tag} pasy węższe niż strona: ${grey.map((t) => t.w.toFixed(2)).join(", ")}`);
      if (empty.some((t) => t.h < 20)) fails.push(`${tag} pusty wiersz bez wysokości: ${empty.map((t) => Math.round(t.h)).join(", ")}`);
      if (r.scrollW > r.vw + 1) fails.push(`${tag} strona przewija się w bok (${r.scrollW} > ${r.vw})`);
      await context.close();
    }
  } finally {
    await browser.close();
    fs.rmSync(file, { force: true });
  }
  if (errors.length) fails.push(`błędy strony: ${errors.slice(0, 2).join(" | ")}`);
  if (fails.length) {
    console.error(`❌ pdf-bands [${ENGINE}]:\n  ${fails.join("\n  ")}`);
    process.exit(1);
  }
  console.log(`✅ pdf-bands [${ENGINE}]: „godz.” na pozycji z PDF, puste szare wiersze jako tabela, pełna szerokość na telefonie`);
})().catch((e) => { console.error(e); process.exit(1); });
