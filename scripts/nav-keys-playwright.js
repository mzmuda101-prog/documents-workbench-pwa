// nav-keys-playwright.js — poruszanie się i zaznaczanie samą klawiaturą (app/doc-caret-nav.js).
// Zgłoszenie 2026-10-11: „strzałką o linijkę wyżej, a przychodzi dwie”, „↓ nie działa”. Pomiar na
// pliku z komentarzami: ↑ stał w miejscu, ↓ krążył, → potrzebował kilku naciśnięć przy komentarzu,
// ↑ z początku 2. wiersza skakał do poprzedniego akapitu, End na Macu nic nie robił.
// Plik budujemy tu (przewodnik.docx: listy i komentarze) — przypadki, na których to padało.
//   node scripts/nav-keys-playwright.js / ENGINE=webkit node scripts/nav-keys-playwright.js

const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const LONG = "Podkreśl, że prawdziwa wiara nie jest bierna. Jeśli wierzymy, że Bóg jest dobry, nasze życie powinno to odzwierciedlać, a wiara musi prowadzić do zmiany stylu życia, odwagi w podejmowaniu decyzji i posłuszeństwa w codziennych sprawach.";

async function buildDocx() {
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(__dirname, "..", "docs", "samples", "przewodnik.docx")));
  const xml = await zip.file("word/document.xml").async("string");
  const head = xml.slice(0, xml.indexOf("<w:body>") + 8);
  const sect = xml.match(/<w:sectPr[\s\S]*<\/w:sectPr>/)[0];
  const t = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
  const comment = (s) => `<w:commentRangeStart w:id="10"/>${t(s)}<w:commentRangeEnd w:id="10"/><w:r><w:commentReference w:id="10"/></w:r>`;
  const list = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
  const cell = (s) => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr><w:p>${t(s)}</w:p></w:tc>`;
  const body = [
    `<w:p>${t("Krótki.")}</w:p>`,
    `<w:p>${t("Wiara, która działa ")}${comment("(Jak 2:19, 20)")}${t(`: ${LONG}`)}</w:p>`,
    `<w:p>${list}${t("Spełnione proroctwa: przypomnij, jak precyzyjnie spełniły się proroctwa dotyczące narodów i królów ")}${comment("(Joz 23:14)")}${t(" oraz jak dokładnie opisano upadek Babilonu i Tyru w dawnych zapiskach.")}</w:p>`,
    `<w:p>${t(`Akapit z łamaniem na końcu. ${LONG}`)}<w:r><w:br/></w:r></w:p>`,
    '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>',
    ...[1, 2, 3].map((i) => `<w:tr>${cell(`Wiersz ${i}`)}${cell(`Wartość ${i}`)}</w:tr>`),
    "</w:tbl>",
    `<w:p>${t("Koniec dokumentu.")}</w:p>`,
  ].join("");
  zip.file("word/document.xml", `${head}${body}${sect}</w:body></w:document>`);
  const file = path.join(os.tmpdir(), `dwb-nav-keys-${process.pid}.docx`);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer" }));
  return file;
}

async function run() {
  const file = await buildDocx();
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles(file);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.click('.mode-btn[data-mode="edit"]');
  await page.waitForSelector(".docx-edit-root");
  await page.waitForTimeout(800);
  const fails = [];
  const check = (name, ok, detail = "") => { if (!ok) fails.push(`${name} ${detail}`); console.log(`${ok ? "✓" : "✗"} ${name}${ok ? "" : ` ${detail}`}`); };

  // Pomocnicze w stronie: akapity, wiersze (prostokąty tekstu), położenie końca zaznaczenia.
  await page.evaluate(() => {
    const list = () => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    window.__nav = {
      list,
      lines(i) {
        const r = document.createRange();
        r.selectNodeContents(list()[i]);
        const out = [];
        [...r.getClientRects()].filter((x) => x.height > 1).sort((a, b) => a.top - b.top).forEach((x) => {
          const l = out.find((l) => Math.min(l.bottom, x.bottom) - Math.max(l.top, x.top) > 2);
          if (l) { l.top = Math.min(l.top, x.top); l.bottom = Math.max(l.bottom, x.bottom); } else out.push({ top: x.top, bottom: x.bottom, left: x.left });
        });
        return out;
      },
      put(i, off) { const p = list()[i]; p.closest(".docx-edit-root").focus(); p.scrollIntoView({ block: "center" }); placeCaret(p, off); },
      // początek wiersza li akapitu i (pozycja na granicy zawinięcia — tak stawia ją strzałka)
      putLineStart(i, li) {
        const p = list()[i];
        p.closest(".docx-edit-root").focus();
        p.scrollIntoView({ block: "center" });
        const l = this.lines(i)[li];
        const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        const r = document.createRange();
        for (let n = w.nextNode(); n; n = w.nextNode()) {
          for (let k = 0; k < n.length; k++) {
            r.setStart(n, k); r.setEnd(n, k + 1);
            const c = r.getClientRects()[0];
            if (c && c.top + c.height / 2 > l.top && c.top + c.height / 2 < l.bottom) {
              const s = getSelection(); const q = document.createRange(); q.setStart(n, k); s.removeAllRanges(); s.addRange(q);
              return;
            }
          }
        }
      },
      // przesunięcia, od których zaczynają się kolejne wiersze akapitu i
      lineStarts(i) {
        const p = list()[i];
        const ls = this.lines(i);
        const out = [];
        const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        const r = document.createRange();
        let off = 0;
        for (let n = w.nextNode(); n; n = w.nextNode()) {
          for (let k = 0; k < n.length; k++, off++) {
            r.setStart(n, k); r.setEnd(n, k + 1);
            const c = r.getClientRects()[0];
            if (!c) continue;
            const li = ls.findIndex((l) => c.top + c.height / 2 >= l.top && c.top + c.height / 2 <= l.bottom);
            if (li >= 0 && out[li] == null) out[li] = off;
          }
        }
        return out;
      },
      where() {
        const s = getSelection();
        const n = s.focusNode;
        const p = (n?.nodeType === 1 ? n : n?.parentElement)?.closest(".docx-editable-p");
        const i = list().indexOf(p);
        const pre = document.createRange();
        if (p) { pre.selectNodeContents(p); pre.setEnd(n, s.focusOffset); }
        // wiersz: znak ZA kursorem (tak rysuje go przeglądarka po strzałce), na końcu — znak przed
        let li = -1;
        if (p) {
          const ls = this.lines(i);
          const r = document.createRange();
          let c = null;
          if (n.nodeType === 3 && s.focusOffset < n.length) { r.setStart(n, s.focusOffset); r.setEnd(n, s.focusOffset + 1); c = r.getClientRects()[0]; }
          if (!c && n.nodeType === 3 && s.focusOffset > 0) { r.setStart(n, s.focusOffset - 1); r.setEnd(n, s.focusOffset); c = r.getClientRects()[0]; }
          if (c) li = ls.findIndex((l) => c.top + c.height / 2 >= l.top && c.top + c.height / 2 <= l.bottom);
        }
        const cell = p?.closest("td");
        return { i, off: p ? pre.toString().length : -1, li, len: p?.textContent.length, cell: cell ? `${cell.parentElement.rowIndex},${cell.cellIndex}` : null, collapsed: s.isCollapsed, afterBr: n?.nodeType === 1 && n.childNodes[s.focusOffset - 1]?.nodeName === "BR", anchor: `${s.anchorNode?.nodeName}@${s.anchorOffset}` };
      },
    };
  });
  const where = () => page.evaluate(() => __nav.where());
  const press = async (key, n = 1) => { for (let k = 0; k < n; k++) await page.keyboard.press(key); return where(); };

  const P_SHORT = 0, P_CMT = 1, P_LIST = 2, P_BR = 3;
  const linesCmt = await page.evaluate((i) => __nav.lines(i).length, P_CMT);
  check("przygotowanie: akapit z komentarzem ma kilka wierszy", linesCmt >= 2, String(linesCmt));

  // ← / → przez znaczniki komentarza: każde naciśnięcie = jeden znak
  await page.evaluate((i) => __nav.put(i, 0), P_CMT);
  let stalls = 0;
  let w = await where();
  for (let k = 0; k < 40; k++) {
    const w2 = await press("ArrowRight");
    if (w2.off !== w.off + 1) stalls++;
    w = w2;
  }
  check("→ przez komentarz „(Jak 2:19, 20)”: 40 naciśnięć = 40 znaków", stalls === 0 && w.off === 40, JSON.stringify(w));
  for (let k = 0; k < 40; k++) {
    const w2 = await press("ArrowLeft");
    if (w2.off !== w.off - 1) stalls++;
    w = w2;
  }
  check("← z powrotem przez komentarz: 40 naciśnięć = 40 znaków", stalls === 0 && w.off === 0, JSON.stringify(w));
  await page.evaluate((i) => __nav.put(i, 14), P_CMT);
  w = await press("Shift+ArrowRight", 22);
  check("Shift+→ przez komentarz zaznacza znak po znaku", !w.collapsed && w.off === 36, JSON.stringify(w));

  // ↓ / ↑ wiersz po wierszu w akapicie z komentarzem (dawniej stał / krążył)
  await page.evaluate((i) => __nav.put(i, 0), P_CMT);
  const seq = [await where()];
  for (let k = 0; k < linesCmt; k++) seq.push(await press("ArrowDown"));
  const downOk = seq.slice(0, linesCmt).every((s, k) => s.i === P_CMT && s.li === k) && seq[linesCmt].i === P_LIST && seq[linesCmt].li === 0;
  check("↓ w akapicie z komentarzem: każdy wiersz po kolei, potem następny akapit", downOk, JSON.stringify(seq.map((s) => [s.i, s.li, s.off])));
  // ↑ z początku 2. wiersza (granica zawinięcia) — 1. wiersz, nie poprzedni akapit
  await page.evaluate((i) => __nav.putLineStart(i, 1), P_CMT);
  w = await press("ArrowUp");
  check("↑ z początku 2. wiersza → 1. wiersz tego akapitu (nie skok o dwa)", w.i === P_CMT && w.li === 0 && w.off === 0, JSON.stringify(w));
  w = await press("ArrowUp");
  check("…kolejne ↑ → akapit wyżej", w.i === P_SHORT, JSON.stringify(w));

  // punkt listy z wysunięciem: ↓ z początku 1. wiersza → 2. wiersz (dawniej kursor stał)
  await page.evaluate((i) => __nav.put(i, 0), P_LIST);
  w = await press("ArrowDown");
  check("punkt listy: ↓ z początku → 2. wiersz", w.i === P_LIST && w.li === 1, JSON.stringify(w));
  w = await press("ArrowUp");
  check("punkt listy: ↑ wraca na początek 1. wiersza", w.i === P_LIST && w.li === 0 && w.off === 0, JSON.stringify(w));

  // akapit kończący się łamaniem wiersza: pusta linijka po <br> jest osobnym przystankiem
  const textLines = (await page.evaluate((i) => __nav.lineStarts(i), P_BR)).filter((x) => x != null).length;
  await page.evaluate((i) => __nav.put(i, 0), P_BR);
  w = await press("ArrowDown", textLines - 1);
  check("akapit z łamaniem: ↓ do ostatniego wiersza tekstu", w.i === P_BR && w.li === textLines - 1 && !w.afterBr, JSON.stringify(w));
  w = await press("ArrowDown");
  check("…↓ → pusta linijka po łamaniu (nie początek akapitu)", w.i === P_BR && w.afterBr, JSON.stringify(w));
  w = await press("ArrowDown");
  check("…↓ → tabela, komórka pod kursorem (1. kolumna)", w.cell === "0,0", JSON.stringify(w));

  // tabela: ↓ w tej samej kolumnie, wyjście pod tabelę
  const cellPara = await page.evaluate(() => __nav.list().findIndex((p) => p.textContent === "Wartość 1"));
  await page.evaluate((i) => __nav.put(i, 3), cellPara);
  const cells = [];
  for (let k = 0; k < 3; k++) cells.push((await press("ArrowDown")).cell);
  w = await where();
  check("tabela: ↓ z 2. kolumny → komórki pod spodem, potem tekst pod tabelą", JSON.stringify(cells) === JSON.stringify(["1,1", "2,1", null]) && /Koniec/.test(await page.evaluate((i) => __nav.list()[i].textContent, w.i)), JSON.stringify(cells));
  w = await press("ArrowUp");
  check("tabela: ↑ spod tabeli → ostatni wiersz, ta sama kolumna", w.cell === "2,1", JSON.stringify(w));

  // End / Home i kolumna pamiętana przez krótki wiersz (jak Word). Koniec wiersza = ta sama
  // pozycja co początek następnego, więc sprawdzamy przesunięcia, nie wysokość.
  const starts = await page.evaluate((i) => __nav.lineStarts(i), P_CMT);
  await page.evaluate((i) => { __nav.put(i, 0); }, P_CMT);
  w = await press("End");
  const endOff = w.off;
  check("End → koniec 1. wiersza (nie koniec akapitu, nie bez ruchu)", w.i === P_CMT && endOff === starts[1], `${JSON.stringify(w)} starts=${starts}`);
  w = await press("ArrowDown");
  check("…↓ po End → koniec 2. wiersza (nie 3.)", w.i === P_CMT && w.off > starts[1] && w.off <= starts[2], `${JSON.stringify(w)} starts=${starts}`);
  await page.evaluate((i) => __nav.put(i, 0), P_CMT);
  await press("End");
  w = await press("ArrowUp");
  check("kolumna: ↑ z końca długiego wiersza → koniec krótkiego akapitu", w.i === P_SHORT && w.off === 7, JSON.stringify(w));
  w = await press("ArrowDown");
  check("…↓ wraca do tej samej kolumny (daleko w wierszu)", w.i === P_CMT && w.off > starts[1] - 6 && w.off <= starts[1], `${JSON.stringify(w)} vs ${endOff}`);
  w = await press("Home");
  check("Home → początek wiersza", w.i === P_CMT && w.off === 0, JSON.stringify(w));
  w = await press("Shift+End");
  check("Shift+End zaznacza do końca wiersza (w tym akapicie)", w.i === P_CMT && !w.collapsed && w.off === endOff, JSON.stringify(w));

  // Shift+↓ — kotwica stoi, koniec idzie wiersz po wierszu
  await page.evaluate((i) => __nav.put(i, 0), P_SHORT);
  const a0 = (await where()).anchor;
  const s1 = await press("Shift+ArrowDown");
  const s2 = await press("Shift+ArrowDown");
  check("Shift+↓ ×2: kotwica w miejscu, koniec w 2. wierszu następnego akapitu", s1.anchor === a0 && s2.anchor === a0 && s1.i === P_CMT && s1.li === 0 && s2.i === P_CMT && s2.li === 1 && !s2.collapsed, JSON.stringify([s1, s2]));

  await browser.close();
  fs.rmSync(file, { force: true });
  if (errors.length) throw new Error(errors.join("\n"));
  if (fails.length) throw new Error(fails.join("\n"));
  console.log(`✅ nav-keys-playwright passed [${ENGINE}]`);
}

run().catch((error) => {
  console.error("❌ nav-keys-playwright failed:", error.message || error);
  process.exit(1);
});
