// align-playwright.js — wyrównanie jak w Wordzie: w komórce tabeli (siatka 3×3: poziomo + pionowo,
// zakres komórka / wiersz / kolumna / tabela) i strony w pionie (sekcja: góra / środek / wyjustuj /
// dół). Sprawdza plik (w:vAlign, w:jc) i podgląd (vertical-align, położenie treści na stronie).
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  const browser = await pw[ENGINE].launch();
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); try { localStorage.setItem("dwb-view-layout", JSON.stringify({ desktop: "desktop" })); } catch (_) {} });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForFunction(() => !!originalFileBytes && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  // dokument: krótki tekst + tabela 3×3 z wysokimi wierszami
  await page.evaluate(async () => {
    const z = await JSZip.loadAsync(originalFileBytes);
    const p = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
    const cell = (t) => `<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;
    const row = (r) => `<w:tr><w:trPr><w:trHeight w:val="1200" w:hRule="atLeast"/></w:trPr>${cell(`A${r}`)}${cell(`B${r}`)}${cell(`C${r}`)}</w:tr>`;
    const tbl = `<w:tbl><w:tblPr><w:tblW w:w="7500" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>${row(1)}${row(2)}${row(3)}</w:tbl>`;
    let doc = await z.file("word/document.xml").async("string");
    doc = doc.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${p("Strona tytułowa")}${tbl}${p("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body>`);
    z.file("word/document.xml", doc);
    await ingestFile(new File([await z.generateAsync({ type: "uint8array" })], "wyrownanie.docx"), { silent: true });
  });
  await page.waitForFunction(() => document.body.innerText.includes("B2") && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 });
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForFunction(() => document.querySelector(".docx-editable-p"), null, { timeout: 10000 });
  await sleep(300);

  const caretIn = (txt) => page.evaluate((txt) => { const p = [...document.querySelectorAll(".docx-editable-p")].find((x) => x.textContent === txt); placeCaret(p, 1); }, txt);
  const fileXml = () => page.evaluate(async () => (await JSZip.loadAsync(await buildDocumentForSave())).file("word/document.xml").async("string"));
  const cellXml = (xml, txt) => (xml.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || []).find((c) => c.includes(`>${txt}<`)) || "";

  // ── komórka: środek-środek ──
  await caretIn("B2");
  await page.click("#tableToolsBtn");
  await page.waitForSelector(".compose-cellalign .compose-align");
  const gridN = await page.evaluate(() => document.querySelectorAll(".compose-cellalign .compose-align").length);
  check("menu tabeli: siatka 3×3 wyrównania + zakres", gridN === 9 && await page.evaluate(() => document.querySelectorAll(".compose-cellalign-scope button").length === 4), gridN);
  await page.click(".compose-cellalign .compose-align:nth-child(5)"); // środek / środek
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 15000 });
  await sleep(500);
  let xml = await fileXml();
  check("komórka B2: w:vAlign center + w:jc center", /<w:vAlign w:val="center"\/>/.test(cellXml(xml, "B2")) && /<w:jc w:val="center"\/>/.test(cellXml(xml, "B2")) && !/vAlign/.test(cellXml(xml, "B1")), cellXml(xml, "B2").slice(0, 300));
  const prev = await page.evaluate(() => { const p = [...document.querySelectorAll("td p")].find((x) => x.textContent === "B2"); const td = p.closest("td"); const a = p.getBoundingClientRect(), b = td.getBoundingClientRect(); return { va: getComputedStyle(td).verticalAlign, ta: getComputedStyle(p).textAlign, mid: Math.abs((a.top + a.bottom) / 2 - (b.top + b.bottom) / 2) }; });
  check("podgląd: tekst w środku komórki (pionowo i poziomo)", prev.va === "middle" && prev.ta === "center" && prev.mid < 4, JSON.stringify(prev));
  const caretBack = await page.evaluate(() => docCaretParagraph(document.activeElement)?.textContent);
  check("kursor wraca do tej komórki", caretBack === "B2", caretBack);

  // ── kolumna: dół-prawo ──
  await caretIn("C1");
  await page.click("#tableToolsBtn");
  await page.waitForSelector(".compose-cellalign-scope button");
  await page.click('.compose-cellalign-scope button[data-v="col"]');
  await page.click(".compose-cellalign .compose-align:nth-child(9)"); // dół / prawo
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 15000 });
  await sleep(500);
  xml = await fileXml();
  check("zakres kolumna: C1–C3 dół + prawo, inne kolumny bez zmian", ["C1", "C2", "C3"].every((c) => /<w:vAlign w:val="bottom"\/>/.test(cellXml(xml, c)) && /<w:jc w:val="right"\/>/.test(cellXml(xml, c))) && !/bottom/.test(cellXml(xml, "A3")));
  // stan w menu: zaznaczony właściwy przycisk
  await caretIn("C2");
  await page.click("#tableToolsBtn");
  await page.waitForSelector(".compose-cellalign .compose-align");
  const onIdx = await page.evaluate(() => [...document.querySelectorAll(".compose-cellalign .compose-align")].findIndex((b) => b.classList.contains("is-on")));
  check("menu pokazuje obecne wyrównanie komórki (dół-prawo)", onIdx === 8, onIdx);
  await page.keyboard.press("Escape");
  await page.evaluate(() => document.querySelector(".compose-pop") && document.getElementById("tableToolsBtn").click());

  // ── tabela: góra-lewo zdejmuje vAlign (domyślne) ──
  await caretIn("A1");
  await page.click("#tableToolsBtn");
  await page.waitForSelector(".compose-cellalign-scope button");
  await page.click('.compose-cellalign-scope button[data-v="table"]');
  await page.click(".compose-cellalign .compose-align:nth-child(1)");
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 15000 });
  await sleep(500);
  xml = await fileXml();
  check("zakres tabela: góra-lewo = bez w:vAlign (domyślne Worda), jc left", !/<w:vAlign/.test(xml) && (xml.match(/<w:jc w:val="left"\/>/g) || []).length >= 9);

  // ── strona w pionie ──
  await caretIn("Strona tytułowa");
  await page.click("#pageLayoutBtn"); // „Układ” na pasku (wyrównanie w pionie — razem z marginesami)
  await page.waitForSelector(".compose-pop .compose-item[data-v='center']");
  await sleep(200);
  const curTop = await page.evaluate(() => document.querySelector(".compose-pop .compose-item[data-v='top']")?.getAttribute("aria-checked"));
  await page.click(".compose-pop .compose-item[data-v='center']");
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 15000 });
  await sleep(600);
  xml = await fileXml();
  check("strona: menu pokazuje obecne (góra), zapis w:vAlign center w sekcji", curTop === "true" && /<w:sectPr[^>]*>(?:(?!<\/w:sectPr>)[\s\S])*<w:vAlign w:val="center"\/>/.test(xml), curTop);
  const pg = await page.evaluate(() => {
    const sec = document.querySelector("section.docx");
    const art = sec.querySelector(":scope > article");
    const S = sec.getBoundingClientRect(), A = art.getBoundingClientRect();
    const cs = getComputedStyle(sec);
    const inTop = S.top + parseFloat(cs.paddingTop), inBottom = S.bottom - parseFloat(cs.paddingBottom);
    return { attr: sec.dataset.dwbVAlign, above: Math.round(A.top - inTop), below: Math.round(inBottom - A.bottom) };
  });
  check("podgląd: treść strony wyśrodkowana w pionie", pg.attr === "center" && pg.above > 50 && Math.abs(pg.above - pg.below) < 30, JSON.stringify(pg));

  if (process.env.OUT) { // ręczna kontrola w Wordzie
    const b64 = await page.evaluate(async () => btoa(Array.from(await buildDocumentForSave(), (x) => String.fromCharCode(x)).join("")));
    require("fs").writeFileSync(process.env.OUT, Buffer.from(b64, "base64"));
  }
  check("brak błędów JS", !errors.length, errors.slice(0, 3).join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`  ${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`);
  console.log(`\n${ENGINE}: ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
