// @ts-check
// mobile-table-cells-playwright.js — tabela (np. z PDF) w Widoku mobilnym telefonu: wcięcia akapitu
// w wąskiej komórce nie zjadają jej szerokości (dawniej „Ilość uczestników” litera pod literą —
// wcięcia 41 + 28 px w komórce 87 px), linie komórek co najmniej 1 px (0,33 px znikały).
// Testerzy AI 2026-11-27, AUDYT-MOBILE #1.
//   node scripts/mobile-table-cells-playwright.js        ENGINE=webkit node scripts/mobile-table-cells-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

const B = '<w:tcBorders><w:top w:val="single" w:sz="2" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="2" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="2" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="2" w:space="0" w:color="000000"/></w:tcBorders>';
const cell = (w, text, ind = "") => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${B}</w:tcPr><w:p><w:pPr>${ind}<w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (cells) => `<w:tr>${cells}</w:tr>`;

(async () => {
  // jak w pliku z PDF (plan prac): kolumny 5245 / 2696 / 2405, nagłówek 3. kolumny z wcięciami
  const table = `<w:tbl><w:tblPr><w:tblW w:w="10346" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="5245"/><w:gridCol w:w="2696"/><w:gridCol w:w="2405"/></w:tblGrid>`
    + row(cell(5245, "ZADANIE (wg. Planu rocznego)") + cell(2696, "BRYGADZISTA") + cell(2405, "Ilość uczestników (sugerowana)", '<w:ind w:left="611" w:right="414" w:hanging="178"/>'))
    + row(cell(5245, "Pomiar rezystancji instalacji odgromowej") + cell(2696, "") + cell(2405, "2"))
    // punkt listy w komórce (jak w pliku z PDF): numer w ::before przesunięty o wysunięcie
    + row(cell(5245, "Wymiana spłuczki sala nr 2, łazienka damska", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="527" w:hanging="425"/>') + cell(2696, "") + cell(2405, "1-2"))
    + "</w:tbl>";
  const bytes = await buildDocx({ body: ["Plan prac na sobotę.", table, "Koniec."], abstractNums: [{ levels: [{ ilvl: 0 }] }], nums: [{ id: 1, abs: 0 }] });
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 }, deviceScaleFactor: 3, hasTouch: true, ...(ENGINE === "chromium" ? { isMobile: true } : {}) });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "plan.docx", mimeType: DOCX_MIME, buffer: bytes });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host table"), null, { timeout: 30000 });
  await page.waitForTimeout(800);
  const st = await page.evaluate(() => {
    const t = document.querySelector(".docx-preview-host table");
    const c = t.rows[0].cells[2];
    const p = c.querySelector("p");
    const pr = p.getBoundingClientRect(), cr = c.getBoundingClientRect();
    const cs = getComputedStyle(c);
    const sides = ["Top", "Right", "Bottom", "Left"].map((s) => parseFloat(cs[`border${s}Width`]));
    // słowo w jednej linijce: wysokość tekstu komórki = jedna linijka
    const oneLine = (td) => { const q = td.querySelector("p"); const lh = parseFloat(getComputedStyle(q).lineHeight) || 16; const rg = document.createRange(); rg.selectNodeContents(q); const tops = new Set([...rg.getClientRects()].filter((r) => r.width > 1).map((r) => Math.round(r.top))); return tops.size === 1 && lh > 0; };
    const words = { brygadzista: oneLine(t.rows[0].cells[1]) };
    const lc = t.rows[2].cells[0], lp = lc.querySelector("p");
    const before = getComputedStyle(lp, "::before");
    const marker = { left: Math.round(lp.getBoundingClientRect().left + (parseFloat(before.left) || 0)), cell: Math.round(lc.getBoundingClientRect().left + (parseFloat(getComputedStyle(lc).borderLeftWidth) || 0)), content: before.content };
    return { words, marker, reflow: shouldUseMobileReflow(), cell: Math.round(cr.width), text: Math.round(pr.width - (parseFloat(getComputedStyle(p).marginLeft) || 0) * 0), lines: Math.round(pr.height / parseFloat(getComputedStyle(p).lineHeight || "16")), pW: Math.round(pr.width), borders: sides };
  });
  check("Widok mobilny (telefon) włączony", st.reflow, JSON.stringify(st));
  check("wąska komórka: akapit zajmuje ≥ 70% szerokości komórki (wcięcia nie zjadają miejsca)", st.pW >= st.cell * 0.7, JSON.stringify(st));
  check("„Ilość uczestników (sugerowana)” w kilku linijkach, nie litera pod literą", st.lines <= 5, JSON.stringify(st)); // było 14 (litera pod literą); w kolumnie 85 px samo „uczestników” pogrubione bywa szersze niż kolumna
  check("wąska kolumna dostaje szerokość najdłuższego słowa: „BRYGADZISTA” bez łamania", st.words.brygadzista, JSON.stringify(st));
  check("numer punktu listy w komórce nie wychodzi na ramkę", st.marker.content !== "none" && st.marker.left >= st.marker.cell, JSON.stringify(st.marker));
  check("linie komórki co najmniej 1 px", st.borders.every((w) => w >= 1), JSON.stringify(st.borders));
  check("bez błędów na stronie", !errors.length, errors.join(" | "));
  await browser.close();
  const bad = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  ${String(r.detail).slice(0, 400)}`}`));
  console.log(`${bad.length ? "❌" : "✅"} Tabele w Widoku mobilnym [${ENGINE}]: ${results.length - bad.length}/${results.length}`);
  process.exit(bad.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
