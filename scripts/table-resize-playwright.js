// table-resize-playwright.js — szerokość kolumn przeciąganiem krawędzi komórki (2026-11-13).
//
// Jak w Wordzie: krawędź wewnątrz tabeli przesuwa granicę dwóch kolumn (suma bez zmian), prawa
// krawędź zmienia ostatnią kolumnę (tabela szersza). W trakcie — przerywana linia z szerokością;
// przyciąganie do równego podziału, do brzegu tekstu strony i do kolumn innych tabel; Alt — bez.
// Sprawdza siatkę tabeli, szerokości komórek i tabeli W PLIKU. Prawdziwa mysz. ENGINE=webkit.

const pw = require("playwright");
const { buildDocx } = require("../tools/word-oracle/build-docx");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const cell = (t, w) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
const B = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>';
const TBL = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${B}</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell("A1", 2000) + cell("B1", 4000)}</w:tr><w:tr>${cell("A2", 2000) + cell("B2", 4000)}</w:tr></w:tbl>`;

async function run() {
  const bytes = await buildDocx({ body: ["Przed", TBL, "Po"] });
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); setSidebarOpen(false); });
  await page.evaluate(async (b) => { const bin = atob(b); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); await ingestFile(new File([u], "kolumny.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" })); }, Buffer.from(bytes).toString("base64"));
  const idle = () => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending && !!document.querySelector(".docx-preview-host table"), null, { timeout: 15000 }).then(() => page.waitForTimeout(600));
  await idle();
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle();
  const grid = () => page.evaluate(async () => {
    const z = await JSZip.loadAsync(originalFileBytes);
    const x = await z.file("word/document.xml").async("string");
    const g = [...x.matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => +m[1]);
    const tc = [...x.matchAll(/<w:tcW w:w="(\d+)"/g)].map((m) => +m[1]);
    const tw = +(x.match(/<w:tblW w:w="(\d+)"/) || [])[1];
    return { g, tc, tw };
  });
  // krawędź prawa komórki nr `col` w pierwszym wierszu (ekran)
  const edge = (col) => page.evaluate((col) => { const td = document.querySelector(".docx-preview-host table").rows[0].cells[col]; const r = td.getBoundingClientRect(); return { x: r.right, y: r.top + r.height / 2 }; }, col);
  const drag = async (col, dx, { alt = false } = {}) => {
    const e = await edge(col);
    await page.mouse.move(e.x - 15, e.y);
    await page.mouse.move(e.x, e.y, { steps: 3 });
    await page.waitForSelector(".tbl-col-hit", { timeout: 3000 });
    await page.mouse.down();
    if (alt) await page.keyboard.down("Alt");
    for (let i = 1; i <= 8; i++) await page.mouse.move(e.x + dx * i / 8, e.y);
    const during = await page.evaluate(() => { const g = document.querySelector(".tbl-col-guide"); return g ? { text: g.textContent, snap: g.classList.contains("is-snap") } : null; });
    await page.mouse.up();
    if (alt) await page.keyboard.up("Alt");
    await idle();
    return during;
  };

  // ── krawędź wewnętrzna: granica przesunięta, suma bez zmian ────────────────
  const d1 = await drag(0, 40);
  const g1 = await grid();
  check("w trakcie: przerywana linia z szerokością kolumny w cm", d1 && /cm/.test(d1.text), JSON.stringify(d1));
  check("krawędź wewnętrzna: lewa kolumna szersza, prawa węższa, suma bez zmian", g1.g[0] > 2000 && g1.g[1] < 4000 && g1.g[0] + g1.g[1] === 6000, JSON.stringify(g1));
  check("szerokości komórek = siatka (oba wiersze)", JSON.stringify(g1.tc) === JSON.stringify([g1.g[0], g1.g[1], g1.g[0], g1.g[1]]), JSON.stringify(g1));
  check("szerokość tabeli = suma kolumn", g1.tw === 6000, JSON.stringify(g1));

  // ── przyciąganie: równy podział dwóch kolumn ──────────────────────────────
  const half = await page.evaluate(() => { const t = document.querySelector(".docx-preview-host table"); const a = t.rows[0].cells[0].getBoundingClientRect(), b = t.rows[0].cells[1].getBoundingClientRect(); return (a.left + b.right) / 2 - a.right; });
  const d2 = await drag(0, half + 3);
  const g2 = await grid();
  check("przyciąganie do równego podziału (etykieta „równe”)", d2?.snap && /równe|equal/.test(d2.text), JSON.stringify(d2));
  check("…kolumny równe w pliku", Math.abs(g2.g[0] - g2.g[1]) <= 2, JSON.stringify(g2));

  // ── prawa krawędź do brzegu tekstu strony ─────────────────────────────────
  const toText = await page.evaluate(() => { const t = document.querySelector(".docx-preview-host table"); const sec = t.closest("section.docx"); const sr = sec.getBoundingClientRect(); const k = sr.width / sec.offsetWidth; const right = sr.right - parseFloat(getComputedStyle(sec).paddingRight) * k; return right - t.rows[0].cells[1].getBoundingClientRect().right; });
  const d3 = await drag(1, toText - 4);
  const g3 = await grid();
  const textTw = await page.evaluate(() => { const d = new DOMParser(); return 11906 - 1417 - 1417; });
  check("prawa krawędź: przyciąga do brzegu tekstu strony", d3?.snap, JSON.stringify(d3));
  check("…tabela na całą szerokość tekstu, zmieniona tylko ostatnia kolumna", Math.abs(g3.tw - textTw) <= 30 && // ±30 tw ≈ 1 px przeliczenia
    g3.g[0] === g2.g[0], JSON.stringify({ g3, textTw }));

  // ── Alt: bez przyciągania ─────────────────────────────────────────────────
  const d4 = await drag(0, half > 0 ? 3 : -3, { alt: true });
  check("Alt: bez przyciągania", d4 && !d4.snap, JSON.stringify(d4));

  // ── Cofnij ─────────────────────────────────────────────────────────────────
  for (let i = 0; i < 4; i++) { await page.evaluate(() => dwbUndo.undo()); await idle(); }
  const g0 = await grid();
  check("Cofnij: szerokości jak w oryginale", JSON.stringify(g0.g) === "[2000,4000]", JSON.stringify(g0));

  // ── Czytanie: bez uchwytu ─────────────────────────────────────────────────
  await page.evaluate(() => appFrame.setReadOnly(true));
  await idle();
  const e5 = await edge(0);
  await page.mouse.move(e5.x - 15, e5.y);
  await page.mouse.move(e5.x, e5.y, { steps: 3 });
  check("Czytanie: krawędzi nie da się przeciągać", !(await page.$(".tbl-col-hit")));

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`);
    if (!r.ok) failed++;
  }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  if (failed) process.exit(1);
}

run().catch((e) => { console.error(e); process.exit(1); });
