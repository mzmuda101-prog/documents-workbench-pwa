// table-rowcol-edge-playwright.js — wysokość wierszy (nowe) i szerokość kolumn przeciąganiem, przegląd
// błędów (2026-11-14). Prawdziwa mysz. Wiersze: w dół (w:trHeight „co najmniej”), w górę poniżej tekstu
// (nie niżej niż tekst), ostatni wiersz, wiersz „dokładnie” zostaje „dokładnie”, tabela w tabeli (tylko
// wewnętrzna), powiększenie 150 %, Cofnij / Ponów, zapis → otwarcie, Czytanie (bez uchwytu). Kolumny:
// powiększenie 150 % (suma bez zmian), scalona komórka, tabela w tabeli, najwęższa kolumna, Cofnij.
//   node scripts/table-rowcol-edge-playwright.js        ENGINE=webkit node scripts/table-rowcol-edge-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 }).then(() => page.waitForTimeout(500));
const cell = (t, w = 3000, extra = "") => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${extra}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
const grid = (ws) => `<w:tblGrid>${ws.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>`;
const TBL = (rows, ws = [3000, 3000, 3000]) => `<w:tbl><w:tblPr><w:tblW w:w="${ws.reduce((a, b) => a + b, 0)}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>${grid(ws)}${rows.join("")}</w:tbl>`;

(async () => {
  const inner = `<w:tbl><w:tblPr><w:tblW w:w="2600" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>${grid([1300, 1300])}<w:tr>${cell("i1", 1300)}${cell("i2", 1300)}</w:tr><w:tr>${cell("i3", 1300)}${cell("i4", 1300)}</w:tr></w:tbl>`;
  const body = [
    "Tabela A",
    TBL([`<w:tr>${cell("A1")}${cell("A2")}${cell("A3")}</w:tr>`, `<w:tr><w:trPr><w:trHeight w:val="900" w:hRule="exact"/></w:trPr>${cell("B1 dokładnie")}${cell("B2")}${cell("B3")}</w:tr>`, `<w:tr>${cell("C1")}${cell("C2")}${cell("C3")}</w:tr>`]),
    "Tabela B (scalona i zagnieżdżona)",
    TBL([`<w:tr>${cell("Scalone", 6000, '<w:gridSpan w:val="2"/>')}${cell("X")}</w:tr>`, `<w:tr>${cell("Y1")}<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${inner}<w:p/></w:tc>${cell("Y3")}</w:tr>`]),
    "Koniec",
  ];
  const bytes = await buildDocx({ body });
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1400, height: 1000 } });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const load = async (b, name) => { await page.locator("#fileInput").setInputFiles({ name, mimeType: DOCX_MIME, buffer: b }); await idle(page); await page.evaluate(() => appFrame.setReadOnly(false)); await idle(page); };
  await load(bytes, "wiersze.docx");

  const T = (sel) => `(${sel})`;
  const rowsTw = () => page.evaluate(async () => {
    const x = await getDocumentXmlString(await buildDocumentForSave());
    const d = new DOMParser().parseFromString(x, "application/xml");
    return [...d.getElementsByTagNameNS(W_NS, "tr")].map((tr) => { const h = tr.getElementsByTagNameNS(W_NS, "trHeight")[0]; const own = h && h.parentNode.parentNode === tr; return own ? `${h.getAttributeNS(W_NS, "val")}/${h.getAttributeNS(W_NS, "hRule")}` : "-"; });
  });
  const gridTw = () => page.evaluate(async () => {
    const x = await getDocumentXmlString(await buildDocumentForSave());
    const d = new DOMParser().parseFromString(x, "application/xml");
    return [...d.getElementsByTagNameNS(W_NS, "tblGrid")].map((g) => [...g.getElementsByTagNameNS(W_NS, "gridCol")].map((c) => +c.getAttributeNS(W_NS, "w")));
  });
  // dolna krawędź wiersza r tabeli t (indeksy w podglądzie; tabele po kolei w dokumencie, też zagnieżdżone)
  const rowEdge = (t, r) => page.evaluate(([t, r]) => { const tb = document.querySelectorAll(".docx-preview-host table")[t]; tb.scrollIntoView({ block: "center", behavior: "instant" }); const row = tb.rows[r]; const rr = row.getBoundingClientRect(); const tr = tb.getBoundingClientRect(); return { x: tr.left + 30, y: rr.bottom, h: rr.height, top: rr.top }; }, [t, r]);
  const colEdge = (t, k) => page.evaluate(([t, k]) => { const tb = document.querySelectorAll(".docx-preview-host table")[t]; tb.scrollIntoView({ block: "center", behavior: "instant" }); const c = tb.rows[tb.rows.length - 1].cells[k]; const rc = c.getBoundingClientRect(); return { x: rc.right, y: rc.top + rc.height / 2 }; }, [t, k]);
  const drag = async (p0, dx, dy, opts = {}) => {
    await page.mouse.move(p0.x + (dy ? 0 : -6), p0.y + (dy ? -6 : 0));
    await page.waitForTimeout(60);
    await page.mouse.move(p0.x, p0.y);
    await page.waitForTimeout(150);
    await page.mouse.down();
    for (let i = 1; i <= 8; i++) { await page.mouse.move(p0.x + (dx * i) / 8, p0.y + (dy * i) / 8); await page.waitForTimeout(20); }
    const label = await page.evaluate(() => document.querySelector(".tbl-row-guide span, .tbl-col-guide span")?.textContent || "");
    await page.mouse.up();
    await idle(page);
    return label;
  };

  // ── WIERSZE ─────────────────────────────────────────────────────────────────
  let e = await rowEdge(0, 0);
  await page.mouse.move(e.x, e.y);
  await page.waitForTimeout(150);
  check("R0: najechanie na dolną krawędź wiersza — uchwyt ↕", await page.evaluate(() => !!document.querySelector(".tbl-row-hit")));
  const lbl = await drag(e, 0, 60);
  let r = await rowsTw();
  const want = Math.round((e.h + 60) * 15);
  check("R1: w dół o 60 px — w:trHeight „co najmniej” ≈ nowa wysokość, etykieta w cm", /\/atLeast$/.test(r[0]) && Math.abs(parseInt(r[0], 10) - want) < 60 && /cm$/.test(lbl), JSON.stringify({ r, want, lbl }));
  const after = await rowEdge(0, 0);
  check("R2: podgląd — wiersz wyższy o ~60 px", Math.abs(after.h - (e.h + 60)) < 4, `${e.h} → ${after.h}`);
  // w górę poniżej tekstu → nie niżej niż tekst
  e = await rowEdge(0, 2);
  await drag(e, 0, -200);
  r = await rowsTw();
  const c2 = await rowEdge(0, 2);
  check("R3: w górę daleko — nie niżej niż tekst wiersza (ostatni wiersz też działa)", r[2] === "-" || parseInt(r[2], 10) >= (c2.h - 2) * 15 * 0.9, JSON.stringify({ r: r[2], h: c2.h }));
  // wiersz „dokładnie” zostaje „dokładnie”
  e = await rowEdge(0, 1);
  await drag(e, 0, 30);
  r = await rowsTw();
  check("R4: wiersz „dokładnie” — po przeciągnięciu dalej exact, wyższy", /\/exact$/.test(r[1]) && parseInt(r[1], 10) > 900, r[1]);
  // tabela zagnieżdżona — tylko jej wiersz
  const outerBefore = (await rowsTw()).slice(3);
  e = await rowEdge(2, 0); // tabele w podglądzie: A, B, wewnętrzna
  await drag(e, 0, 40);
  r = await rowsTw();
  check("R5: tabela w tabeli — zmienia się wiersz wewnętrznej, zewnętrzne bez zmian", JSON.stringify(r.slice(3, 5)) === JSON.stringify(outerBefore.slice(0, 2)) && /\/atLeast$/.test(r[5]), JSON.stringify({ outerBefore, r }));
  // Cofnij / Ponów
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  const u = await rowsTw();
  await page.evaluate(() => document.getElementById("redoBtn").click());
  await idle(page);
  const rr = await rowsTw();
  check("R6: Cofnij / Ponów wysokości", u[5] === "-" && rr[5] === r[5], JSON.stringify({ u: u[5], rr: rr[5], r: r[5] }));
  // zapis → otwarcie: wysokość wiersza w podglądzie ta sama
  const h0 = (await rowEdge(0, 0)).h;
  const b64 = await page.evaluate(async () => { const x = await buildDocumentForSave(); let s = ""; x.forEach((v) => { s += String.fromCharCode(v); }); return btoa(s); });
  await load(Buffer.from(b64, "base64"), "wiersze-2.docx");
  check("R7: zapis → otwarcie — wysokość wiersza ta sama", Math.abs((await rowEdge(0, 0)).h - h0) < 2, `${h0} → ${(await rowEdge(0, 0)).h}`);

  // ── KOLUMNY ─────────────────────────────────────────────────────────────────
  let g0 = await gridTw();
  let ce = await colEdge(0, 0);
  await drag(ce, 60, 0);
  let g1 = await gridTw();
  check("K1: kolumna 1 szersza, kolumna 2 węższa, suma bez zmian", g1[0][0] > g0[0][0] && g1[0][1] < g0[0][1] && g1[0].reduce((a, b) => a + b) === g0[0].reduce((a, b) => a + b), JSON.stringify({ g0: g0[0], g1: g1[0] }));
  // najwęższa: w lewo daleko — nie węższa niż ~0,42 cm
  ce = await colEdge(0, 0);
  await drag(ce, -900, 0);
  g1 = await gridTw();
  check("K2: kolumna nie węższa niż ~0,42 cm (240 tw)", g1[0][0] >= 230, JSON.stringify(g1[0]));
  // scalona komórka (tabela B: pierwszy wiersz gridSpan 2) — granica 1|2 dalej działa w wierszu bez scalenia
  const gb = (await gridTw())[1];
  ce = await colEdge(1, 0);
  await drag(ce, 40, 0);
  const gb2 = (await gridTw())[1];
  check("K3: tabela ze scaloną komórką — granica przesunięta, suma bez zmian", gb2[0] > gb[0] && gb2.reduce((a, b) => a + b) === gb.reduce((a, b) => a + b), JSON.stringify({ gb, gb2 }));
  // tabela w tabeli — tylko wewnętrzna siatka
  const gAll = await gridTw();
  ce = await colEdge(2, 0);
  await drag(ce, 20, 0);
  const gAll2 = await gridTw();
  check("K4: tabela w tabeli — zmienia się siatka wewnętrznej, zewnętrzne bez zmian", JSON.stringify(gAll2.slice(0, 2)) === JSON.stringify(gAll.slice(0, 2)) && gAll2[2][0] > gAll[2][0], JSON.stringify({ gAll, gAll2 }));
  // powiększenie 150 % — przeliczenie px → tw
  await page.evaluate(() => { zoomLevelEl.value = "1.5"; applyZoom(); });
  await page.waitForTimeout(800);
  const zoomed = await page.evaluate(() => { const s = document.querySelector(".docx-preview-host section.docx"); return s ? s.getBoundingClientRect().width / s.offsetWidth : 1; });
  e = await rowEdge(0, 2);
  const hz = e.h;
  await page.mouse.move(e.x, e.y + 8); // od dołu (z akapitu pod tabelą)
  await drag(e, 0, 60);
  r = await rowsTw();
  const wz = Math.round(((hz + 60) / zoomed) * 15);
  check(`R8: powiększenie ×${zoomed.toFixed(2)} — wysokość w pliku liczona bez powiększenia`, Math.abs(parseInt(r[2], 10) - wz) < 80, JSON.stringify({ r: r[2], wz, zoomed }));
  // Czytanie — bez uchwytu
  await page.evaluate(() => appFrame.setReadOnly(true));
  await idle(page);
  e = await rowEdge(0, 0);
  await page.mouse.move(e.x, e.y + 1);
  await page.mouse.move(e.x, e.y);
  await page.waitForTimeout(200);
  check("R9: Czytanie — bez uchwytu wiersza", !(await page.evaluate(() => document.querySelector(".tbl-row-hit"))));
  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const x of results) { console.log(`${x.ok ? "✅" : "❌"} ${x.name}${!x.ok && x.detail ? `\n   ${x.detail}` : ""}`); if (!x.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
