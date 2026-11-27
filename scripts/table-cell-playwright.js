// table-cell-playwright.js — komórka tabeli jak w Wordzie: kierunek tekstu (Układ tabeli → Kierunek tekstu:
// poziomo / 90° / 270°, zakres komórka / wiersz), zwykły tekst „lrTb” z pliku Worda poziomo (docx-preview
// rysował go pionowo), tabela wstawiona w komórce (tabela w tabeli) — plik, podgląd,
// pisanie w niej, Cofnij.
//   node scripts/table-cell-playwright.js        ENGINE=webkit node scripts/table-cell-playwright.js
const fs = require("fs");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 }).then(() => page.waitForTimeout(500));

(async () => {
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);
  await page.keyboard.type("Tabela poniżej.");
  await page.evaluate(() => composeUi.insertTable(2, 2));
  await idle(page);
  const cellCaret = (r, c, sel = ".docx-preview-host table") => page.evaluate(([r, c, sel]) => {
    const p = document.querySelector(sel).rows[r].cells[c].querySelector("p");
    p.closest(".docx-edit-root").focus(); placeCaret(p, 0); document.dispatchEvent(new Event("selectionchange"));
  }, [r, c, sel]);
  await cellCaret(0, 0);
  await page.keyboard.type("Nagłówek");
  await idle(page);
  const tcDirs = () => page.evaluate(async () => {
    const x = await getDocumentXmlString(await buildDocumentForSave());
    return (x.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || []).map((tc) => (tc.match(/<w:textDirection w:val="(\w+)"\/>/) || [])[1] || "-");
  });
  const wm = () => page.evaluate(() => [...document.querySelector(".docx-preview-host table").rows].map((r) => [...r.cells].map((c) => getComputedStyle(c).writingMode)));
  const openTableMenu = async () => {
    await page.waitForFunction(() => !document.getElementById("tableToolsBtn").hidden);
    await page.waitForFunction(() => !document.getElementById("tableToolsBtn").hidden).then(() => page.evaluate(() => typeof dwbRibbon !== "undefined" && dwbRibbon?.isOn() && dwbRibbon.select("table"))).then(() => page.evaluate(() => document.getElementById("tableToolsBtn").click()));
    await page.waitForSelector(".compose-textdir");
  };

  // 1. komórka: 90° w dół
  await cellCaret(0, 0);
  await openTableMenu();
  await page.click('.compose-textdir [data-dir="tbRl"]');
  await idle(page);
  check("1: kierunek 90° — plik tbRl tylko w tej komórce, podgląd pionowo", JSON.stringify(await tcDirs()) === JSON.stringify(["tbRl", "-", "-", "-"]) && (await wm())[0][0] === "vertical-rl", JSON.stringify({ f: await tcDirs(), w: await wm() }));
  // 2. menu pokazuje bieżący kierunek; wiersz: 270°
  await cellCaret(0, 0);
  await openTableMenu();
  const cur = await page.evaluate(() => document.querySelector('.compose-textdir [data-dir="tbRl"]').getAttribute("aria-pressed"));
  check("2: menu zaznacza bieżący kierunek (90°)", cur === "true", cur);
  await page.click('.compose-cellalign-scope [data-v="row"]');
  await page.click('.compose-textdir [data-dir="btLr"]');
  await idle(page);
  check("3: zakres „Wiersz”, 270° — obie komórki wiersza btLr", JSON.stringify(await tcDirs()) === JSON.stringify(["btLr", "btLr", "-", "-"]), JSON.stringify(await tcDirs()));
  // 4. poziomo — znika z pliku
  await cellCaret(0, 1);
  await openTableMenu();
  await page.click('.compose-textdir [data-dir="lrTb"]');
  await idle(page);
  check("4: „Poziomo” — w pliku bez w:textDirection (jak Word)", JSON.stringify(await tcDirs()) === JSON.stringify(["btLr", "-", "-", "-"]) && (await wm())[0][1] === "horizontal-tb", JSON.stringify(await tcDirs()));
  // 5. Cofnij
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  check("5: Cofnij — znów obie komórki btLr", JSON.stringify(await tcDirs()) === JSON.stringify(["btLr", "btLr", "-", "-"]), JSON.stringify(await tcDirs()));

  // 6. tabela w tabeli: kursor w komórce (1,1) → Wstaw tabelę 2×2
  await cellCaret(1, 1);
  await page.evaluate(() => composeUi.insertTable(2, 2));
  await idle(page);
  const nested = await page.evaluate(async () => {
    const x = await getDocumentXmlString(await buildDocumentForSave());
    const doc = new DOMParser().parseFromString(x, "application/xml");
    const outer = doc.getElementsByTagNameNS(W_NS, "tbl")[0];
    const inner = [...outer.getElementsByTagNameNS(W_NS, "tbl")];
    const tc = inner[0]?.parentNode;
    const kids = tc ? [...tc.childNodes].filter((n) => n.nodeType === 1).map((n) => n.localName) : [];
    return { inner: inner.length, kids, view: document.querySelectorAll(".docx-preview-host table table").length };
  });
  check("6: tabela w komórce — w pliku w:tbl w w:tc, komórka kończy się akapitem, w podglądzie tabela w tabeli", nested.inner === 1 && nested.kids[nested.kids.length - 1] === "p" && nested.view === 1, JSON.stringify(nested));
  // 7. pisanie w zagnieżdżonej komórce trafia do pliku
  if (nested.view) {
    await cellCaret(0, 0, ".docx-preview-host table table");
    await page.keyboard.type("Wewnątrz");
    await idle(page);
    const inside = await page.evaluate(async () => { const x = await getDocumentXmlString(await buildDocumentForSave()); const d = new DOMParser().parseFromString(x, "application/xml"); return d.getElementsByTagNameNS(W_NS, "tbl")[0].getElementsByTagNameNS(W_NS, "tbl")[0]?.textContent || ""; });
    check("7: pisanie w komórce tabeli zagnieżdżonej — tekst w pliku", inside.includes("Wewnątrz"), inside);
  }
  check("bez błędów strony", !errors.length, errors.join(" | "));

  // 8. komórka z „lrTb” z pliku (OOXML: zwykły tekst poziomo) — docx-preview rysował ją pionowo
  {
    const { buildDocx } = require("../tools/word-oracle/build-docx");
    const tc = (dir, t) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${dir ? `<w:textDirection w:val="${dir}"/>` : ""}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
    const bytes = await buildDocx({ body: [`<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${tc("lrTb", "Poziomo")}${tc("tbRl", "W dół")}${tc("btLr", "W górę")}</w:tr></w:tbl>`, "Koniec"] });
    await page.locator("#fileInput").setInputFiles({ name: "kierunki.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: bytes });
    await idle(page);
    const w = await page.evaluate(() => [...document.querySelector(".docx-preview-host table").rows[0].cells].map((c) => getComputedStyle(c).writingMode));
    check("8: komórka „lrTb” z pliku poziomo, „tbRl” / „btLr” pionowo", JSON.stringify(w) === JSON.stringify(["horizontal-tb", "vertical-rl", "vertical-rl"]), JSON.stringify(w));
  }

  // 9. klik w wolne miejsce (dolna połowa) wysokiej pustej komórki zostawia kursor w TEJ komórce
  {
    const { buildDocx } = require("../tools/word-oracle/build-docx");
    const B = '<w:tcBorders><w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/></w:tcBorders>';
    const tc = (txt) => `<w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/>${B}</w:tcPr><w:p>${txt ? `<w:r><w:t>${txt}</w:t></w:r>` : ""}</w:p></w:tc>`;
    const tr = (a, b, c) => `<w:tr><w:trPr><w:trHeight w:val="720" w:hRule="atLeast"/></w:trPr>${tc(a)}${tc(b)}${tc(c)}</w:tr>`;
    const bytes = await buildDocx({ body: [`<w:tbl><w:tblPr><w:tblW w:w="8400" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2800"/><w:gridCol w:w="2800"/><w:gridCol w:w="2800"/></w:tblGrid>${tr("Pierwszy host", "", "")}${tr("Ostatni host", "", "")}${tr("Liczba hostów", "", "")}</w:tbl>`] });
    await page.locator("#fileInput").setInputFiles({ name: "puste-komorki.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: bytes });
    await idle(page);
    if (await page.evaluate(() => readOnlyMode)) {
      await page.evaluate(() => appFrame.setReadOnly(false));
      await idle(page);
    }
    const r12 = await page.evaluate(() => document.querySelector(".docx-preview-host table").rows[1].cells[2].getBoundingClientRect());
    await page.mouse.click(r12.x + r12.width / 2, r12.bottom - 5);
    await page.waitForTimeout(150);
    await page.keyboard.type("192.168.1.254");
    await page.waitForTimeout(150);
    const hit = await page.evaluate(() => {
      const t = document.querySelector(".docx-preview-host table");
      const s = window.getSelection();
      const n = s?.rangeCount ? s.anchorNode : null;
      const td = (n?.nodeType === 1 ? n : n?.parentElement)?.closest("td, th");
      return {
        row: td?.parentElement?.rowIndex ?? -1,
        col: td?.cellIndex ?? -1,
        cellText: t.rows[1].cells[2].textContent.trim(),
        col0Text: t.rows[1].cells[0].textContent.trim(),
        nextRowText: t.rows[2].cells[0].textContent.trim(),
      };
    });
    check("9: klik w dolną połowę wysokiej pustej komórki stawia kursor i wpisuje tekst w TEJ komórce",
      hit.row === 1 && hit.col === 2 && hit.cellText === "192.168.1.254" && hit.col0Text === "Ostatni host" && hit.nextRowText === "Liczba hostów",
      JSON.stringify(hit));
  }
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
