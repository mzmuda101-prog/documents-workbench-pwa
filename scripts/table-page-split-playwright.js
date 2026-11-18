// table-page-split-playwright.js — tabela na granicy stron w Edycji dzieli się między wierszami, jak
// w Wordzie i w Podglądzie wydruku (2026-11-18, zgłoszenie: Kolokwium 18.11 — w Edycji całe tabele
// przeskakiwały na następną stronę, od strony 3 każda strona inna niż w Wordzie).
// Sprawdzamy: tabela zaczyna się na stronie 1, odstęp strony stoi między wierszami (element w roli
// wiersza, nie <tr> — table.rows bez zmian), wiersz nagłówka powtórzony u góry strony 2 (kopia do
// oglądania, niewidoczna dla tekstu dokumentu), pierwszy wiersz strony 2 pod nim, odstęp stabilny
// przy kolejnych przeliczeniach, Podgląd wydruku zaczyna stronę 2 tym samym wierszem.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

const t = (s) => `<w:t xml:space="preserve">${s}</w:t>`;
const P = (s) => `<w:p><w:r>${t(s)}</w:r></w:p>`;
const B = ["top", "left", "bottom", "right", "insideH", "insideV"].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join("");
const cell = (s) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r>${t(s)}</w:r></w:p></w:tc>`;
const rows = [`<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell("Naglowek A")}${cell("Naglowek B")}</w:tr>`];
for (let i = 1; i <= 30; i++) rows.push(`<w:tr>${cell(`Wiersz ${i}`)}${cell("tresc")}</w:tr>`);
const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${B}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${rows.join("")}</w:tbl>`;

async function run() {
  const file = path.join(os.tmpdir(), `dwb-table-split-${process.pid}.docx`);
  fs.writeFileSync(file, await buildDocx({ body: [...Array.from({ length: 22 }, (_, i) => P(`Akapit przed tabela ${i + 1}`)), table, P("Po tabeli")] }));
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; window.__dwbOpenShadow = true; try { localStorage.removeItem("dwb-page-gaps-v1"); } catch (_) {} });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && [...document.querySelectorAll(".docx-preview-host p")].some((p) => p.textContent === "Po tabeli"), null, { timeout: 30000 });
  fs.rmSync(file, { force: true });
  await page.evaluate(() => document.querySelector('.mode-btn[data-mode="edit"]')?.click());
  await page.waitForFunction(() => document.querySelector(".dwb-gap-row"), null, { timeout: 15000 }).catch(() => {});
  await page.waitForTimeout(1500);

  const st = await page.evaluate(() => {
    const host = document.querySelector(".docx-preview-host");
    const sec = host.querySelector("section.docx");
    const table = sec.querySelector("article > table");
    const gap = sec.querySelector(".dwb-gap-row");
    const band = sec.querySelector(".dwb-page-gap-band");
    const next = gap?.nextElementSibling;
    const rep = sec.querySelector(".dwb-page-repeat");
    const repText = rep?.shadowRoot?.textContent || "";
    const scale = sec.getBoundingClientRect().height / sec.offsetHeight || 1;
    const y = (el) => (el.getBoundingClientRect().top - sec.getBoundingClientRect().top) / scale;
    return {
      tableBeforeBand: table && band ? y(table) < y(band) : false,
      gapInTable: !!gap && gap.closest("table") === table && gap.tagName === "DIV",
      rows: table.rows.length,
      nextRow: next?.tagName === "TR" ? next.textContent : null,
      repText,
      repBelowBand: rep && band ? y(rep) > y(band) : false,
      rowBelowRep: rep && next ? y(next) >= y(rep) + rep.getBoundingClientRect().height / scale - 1.5 : false,
      docText: host.textContent.includes("Naglowek ANaglowek B") ? host.textContent.split("Naglowek A").length - 1 : -1,
      gapEl: !!(window.__gap0 = gap),
    };
  });
  check("tabela zaczyna się na stronie 1 (nie przeskakuje w całości na stronę 2)", st.tableBeforeBand, JSON.stringify(st));
  check("odstęp strony między wierszami tabeli — element w roli wiersza, nie <tr>", st.gapInTable, JSON.stringify(st));
  check("table.rows bez zmian (31 wierszy: nagłówek + 30) — narzędzia tabeli liczą jak dawniej", st.rows === 31, JSON.stringify(st));
  check("strona 2 zaczyna się wierszem tabeli", /^Wiersz \d+tresc$/.test(st.nextRow || ""), JSON.stringify(st));
  check("wiersz nagłówka powtórzony u góry strony 2 (jak w Wordzie)", /Naglowek A/.test(st.repText) && st.repBelowBand, JSON.stringify(st));
  check("pierwszy wiersz strony 2 pod powtórzonym nagłówkiem", st.rowBelowRep, JSON.stringify(st));
  check("kopia nagłówka nie jest tekstem dokumentu (słowa, szukanie)", st.docText === 1, JSON.stringify(st));

  await page.evaluate(() => dwbPageBreaks.compute());
  await page.waitForTimeout(600);
  await page.evaluate(() => dwbPageBreaks.compute());
  await page.waitForTimeout(600);
  check("odstęp stabilny przy kolejnych przeliczeniach (bez przebudowy)", await page.evaluate(() => window.__gap0?.isConnected === true));

  // Podgląd wydruku: strona 2 zaczyna się tym samym wierszem
  const editFirst = st.nextRow?.replace(/tresc$/, "");
  await page.evaluate(() => dwbPrint.open());
  await page.waitForSelector(".pp-sheet", { timeout: 20000 });
  await page.waitForTimeout(800);
  const ppFirst = await page.evaluate(() => {
    const sh = document.querySelectorAll(".pp-sheet")[1];
    if (!sh) return null;
    // widoczne = wewnątrz okna kartki i wszystkich przycinających przodków (kopie bloków są przycięte)
    const visible = (el) => {
      const r = el.getBoundingClientRect();
      if (!r.height) return false;
      for (let a = el.parentElement; a && a !== sh.parentElement; a = a.parentElement) {
        const o = getComputedStyle(a);
        if (o.overflow === "visible" && o.overflowY === "visible" && o.clipPath === "none") continue;
        const c = a.getBoundingClientRect();
        if (r.top < c.top - 0.5 || r.bottom > c.bottom + 0.5) return false;
      }
      return true;
    };
    const rows = [...sh.querySelectorAll("tr")].filter((tr) => /^Wiersz/.test(tr.textContent) && visible(tr));
    rows.sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top);
    return rows[0]?.cells[0].textContent || null;
  });
  check("Podgląd wydruku zaczyna stronę 2 tym samym wierszem co Edycja", ppFirst && ppFirst === editFirst, JSON.stringify({ editFirst, ppFirst }));

  if (errors.length) check("bez błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Tabela na granicy stron [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
