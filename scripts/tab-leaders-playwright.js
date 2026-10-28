// tab-leaders-playwright.js — długość kropek w Edycji (tab-leaders.js, 2026-10-28).
//
// Zgłoszenie: w karcie z PDF nie da się skrócić kropek „Miejscowość: ……… Teren nr:” — to jeden
// tabulator z wypełnieniem. Klik w kropki: uchwyt na końcu + karta ◀ ▶; przeciągnięcie / ◀ zmienia
// pozycję tabulatora akapitu w pliku (jak linijka Worda), Cofnij wraca. Tabulator ze stylu:
// „wyczyść” + nowy (jak Word). Kropki nie skracają się za tekst. Backspace za kropkami dalej usuwa
// cały tabulator. W Czytaniu nic się nie pokazuje. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(350));
const mod = process.platform === "darwin" ? "Meta" : "Control";

const tab = '<w:r><w:tab/></w:r>';
const t = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const sdt = (id, name, v) => `<w:sdt><w:sdtPr><w:alias w:val="${name}"/><w:tag w:val="${name}"/><w:id w:val="${id}"/><w:text/></w:sdtPr><w:sdtContent>${t(v)}</w:sdtContent></w:sdt>`;
async function fixture() {
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  // styl „Formularz” z tabulatorem z kropkami na 4 cm
  z.file("word/styles.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Formularz"><w:name w:val="Formularz"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="left" w:leader="dot" w:pos="5000"/></w:tabs></w:pPr></w:style></w:styles>');
  const body = [
    `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:leader="dot" w:pos="2000"/><w:tab w:val="left" w:leader="dot" w:pos="6000"/><w:tab w:val="left" w:leader="dot" w:pos="8000"/></w:tabs></w:pPr>${t("Miejscowość:")}${tab}${t("PILECKIEGO 23")}${tab}${t("Teren nr:")}${tab}</w:p>`,
    `<w:p><w:pPr><w:pStyle w:val="Formularz"/></w:pPr>${t("Imię:")}${tab}</w:p>`,
    `<w:p>${t("Koniec.")}</w:p>`,
    // karta z PDF: kontrolki Worda (Miejscowość, Teren nr) między tabulatorami z kropkami
    `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:leader="dot" w:pos="2000"/><w:tab w:val="left" w:leader="dot" w:pos="5000"/><w:tab w:val="left" w:leader="dot" w:pos="6600"/><w:tab w:val="left" w:leader="dot" w:pos="8500"/></w:tabs></w:pPr>${t("Miejscowość:")}${tab}${sdt(11, "Miejscowość", "PILECKIEGO 23")}${tab}${t("Teren nr:")}${tab}${sdt(12, "Teren nr", "401")}${tab}</w:p>`,
  ].join("");
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return z.generateAsync({ type: "nodebuffer" });
}

const tabsOf = (page, i) => page.evaluate(async (i) => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
  const p = collectParagraphElements(d.documentElement, "all")[i];
  return { tabs: Array.from(p.getElementsByTagNameNS(W_NS, "tab")).filter((x) => x.parentNode.localName === "tabs").map((x) => `${x.getAttributeNS(W_NS, "val")}${x.getAttributeNS(W_NS, "leader") ? `/${x.getAttributeNS(W_NS, "leader")}` : ""}@${x.getAttributeNS(W_NS, "pos")}`), tabChars: Array.from(p.getElementsByTagNameNS(W_NS, "tab")).filter((x) => x.parentNode.localName === "r").length };
}, i);
const tabRect = (page, i, k) => page.evaluate(([i, k]) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; const r = p.querySelectorAll(".docx-tab")[k].getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2, l: r.left, r: r.right }; }, [i, k]);
const handle = (page) => page.evaluate(() => { const e = document.querySelector(".dwb-tabstop-handle"); if (!e || e.style.visibility === "hidden") return null; const r = e.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const file = require("path").join(require("os").tmpdir(), `dwb-kropki-${process.pid}.docx`);
  require("fs").writeFileSync(file, await fixture());
  await page.setInputFiles("#fileInput", file);
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  require("fs").rmSync(file, { force: true });
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await idle(page);

  // 1. klik w kropki przed „Teren nr:” → zaznaczone, uchwyt na końcu, karta ◀ ▶
  let r = await tabRect(page, 0, 1);
  await page.mouse.click(r.x, r.y);
  await page.waitForTimeout(250);
  let h = await handle(page);
  check("klik w kropki: uchwyt na końcu kropek i karta ◀ ▶", h && Math.abs(h.x - r.r) < 3 && await page.evaluate(() => !!document.querySelector(".dwb-tabstop-card") && !!document.querySelector(".dwb-tabstop-hl")), JSON.stringify({ h, r }));

  // 2. przeciągnięcie uchwytu w lewo: tabulator 6000 → mniej (w pliku), reszta bez zmian
  await page.mouse.move(h.x, h.y);
  await page.mouse.down();
  await page.mouse.move(h.x - 20, h.y, { steps: 3 });
  await page.mouse.move(h.x - 60, h.y, { steps: 4 });
  await page.mouse.up();
  await idle(page);
  let st = await tabsOf(page, 0);
  const moved = st.tabs[1] && parseInt(st.tabs[1].split("@")[1], 10);
  check("przeciągnięcie: koniec kropek bliżej (pozycja tabulatora w pliku)", st.tabs.length === 3 && st.tabs[0] === "left/dot@2000" && st.tabs[2] === "left/dot@8000" && moved < 5500 && moved > 3000, JSON.stringify(st));
  check("po zmianie kropki dalej zaznaczone (można dalej poprawiać)", await page.evaluate(() => !!document.querySelector(".dwb-tabstop-hl") && !!document.querySelector(".dwb-tabstop-handle")));

  // 3. ◀ dwa razy = 2 × 3 pt (60 twipów) krócej, jedna zmiana w pliku
  await page.click('.dwb-tabstop-card [data-d="-1"]');
  await page.click('.dwb-tabstop-card [data-d="-1"]');
  await page.waitForTimeout(900);
  await idle(page);
  st = await tabsOf(page, 0);
  const after2 = parseInt(st.tabs[1].split("@")[1], 10);
  check("◀ ◀: krócej o 2 × 3 pt", Math.abs(after2 - (moved - 120)) <= 2, `${moved} → ${after2}`);

  // 4. Cofnij: z powrotem do pozycji po przeciągnięciu
  await page.keyboard.press("Escape");
  await page.evaluate(() => document.querySelector(".docx-edit-root")?.focus());
  await page.keyboard.press(`${mod}+z`);
  await idle(page);
  st = await tabsOf(page, 0);
  check("Cofnij: pozycja sprzed ◀ ◀", parseInt(st.tabs[1].split("@")[1], 10) === moved, JSON.stringify(st));

  // 5. kropki nie skracają się za tekst (koniec „PILECKIEGO 23”)
  r = await tabRect(page, 0, 1);
  await page.mouse.click(r.x, r.y);
  await page.waitForTimeout(250);
  h = await handle(page);
  await page.mouse.move(h.x, h.y);
  await page.mouse.down();
  await page.mouse.move(h.x - 400, h.y, { steps: 6 });
  await page.mouse.up();
  await idle(page);
  const textEnd = await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0]; const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); let pil = null, ter = null; for (let n = w.nextNode(); n; n = w.nextNode()) { if (/PILECKIEGO/.test(n.data)) pil = n; if (/Teren/.test(n.data)) ter = n; } const rr = (n) => { const x = document.createRange(); x.selectNodeContents(n); return x.getBoundingClientRect(); }; return { pil: rr(pil).right, ter: rr(ter).left, sameLine: Math.abs(rr(pil).top - rr(ter).top) < 4 }; });
  check("skrócenie do oporu: „Teren nr” zaraz za tekstem, w tej samej linijce", textEnd.sameLine && textEnd.ter >= textEnd.pil - 0.5 && textEnd.ter - textEnd.pil < 12, JSON.stringify(textEnd));

  // 6. tabulator ze stylu: w akapicie „wyczyść” + nowy (jak Word)
  await page.keyboard.press("Escape");
  r = await tabRect(page, 1, 0);
  await page.mouse.click(r.x, r.y);
  await page.waitForTimeout(250);
  h = await handle(page);
  check("tabulator ze stylu: uchwyt jest", !!h);
  if (h) {
    await page.mouse.move(h.x, h.y);
    await page.mouse.down();
    await page.mouse.move(h.x + 40, h.y, { steps: 4 });
    await page.mouse.up();
    await idle(page);
    st = await tabsOf(page, 1);
    check("tabulator ze stylu: w akapicie wyczyść @5000 + nowy dalej", st.tabs.length === 2 && st.tabs[0] === "clear@5000" && /^left\/dot@\d+$/.test(st.tabs[1]) && parseInt(st.tabs[1].split("@")[1], 10) > 5300, JSON.stringify(st));
  }

  // 7. Backspace za kropkami dalej usuwa cały tabulator (zachowanie sprzed zmiany)
  await page.keyboard.press("Escape");
  r = await tabRect(page, 1, 0);
  await page.mouse.click(r.r - 2, r.y);
  await page.waitForTimeout(200);
  await page.keyboard.press("Backspace");
  await page.waitForTimeout(400);
  st = await tabsOf(page, 1);
  check("Backspace za kropkami usuwa tabulator (znak), karta znika", st.tabChars === 0 && await page.evaluate(() => !document.querySelector(".dwb-tabstop-card")), JSON.stringify(st));

  // 7b. Dłuższy tekst przed kropkami (nagranie 2026-10-28: „PILECKIEGO 23” → „PILECKIEGO 233450”):
  // kropki kończą się na pozycji tabulatora, więc „Teren nr:” zostaje w miejscu — dawniej szerokości
  // tabulatorów liczone tylko przy renderze i reszta wiersza wyjeżdżała za kolumnę.
  const terenX = () => page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[3];
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      const k = n.data.indexOf("Teren");
      if (k >= 0) { const q = document.createRange(); q.setStart(n, k); q.setEnd(n, k + 1); return q.getBoundingClientRect().left - p.getBoundingClientRect().left; }
    }
    return null;
  });
  const x0 = await terenX();
  const fkey = await page.evaluate(() => [...formUi.ranges.keys()].find((k) => formUi.ranges.get(k).toString().includes("PILECKIEGO")));
  const fbox = await page.evaluate((k) => { const q = formUi.ranges.get(k); const c = [...q.getClientRects()].find((x) => x.width > 0); return { x: c.left + 10, y: c.top + c.height / 2 }; }, fkey);
  await page.mouse.click(fbox.x, fbox.y);
  await page.waitForSelector(".ff-pop .ff-pop-input");
  await page.fill(".ff-pop .ff-pop-input", "PILECKIEGO 233450");
  await page.keyboard.press("Enter");
  await idle(page);
  await page.waitForTimeout(300);
  let x1 = await terenX();
  check("dłuższy tekst w kontrolce: „Teren nr:” w tym samym miejscu (kropki krótsze)", x0 != null && Math.abs(x1 - x0) < 1, `${x0} → ${x1}`);
  await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[3];
    p.closest(".docx-edit-root").focus();
    const q = document.createRange(); q.setStart(document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode(), 5);
    getSelection().removeAllRanges(); getSelection().addRange(q);
  });
  await page.keyboard.type("ab");
  await page.waitForTimeout(300);
  x1 = await terenX();
  const savedText = await page.evaluate(async () => (await extractParagraphTextsFromDocx(await buildDocumentForSave())).map(String)[3]);
  check("pisanie przed kropkami: „Teren nr:” w tym samym miejscu, tekst w pliku", Math.abs(x1 - x0) < 1 && /^Miejsabcowość:PILECKIEGO 233450Teren nr:401$/.test(savedText), `${x0} → ${x1} | ${JSON.stringify(savedText)}`);

  // 8. Czytanie: klik w kropki nic nie pokazuje
  await page.evaluate(() => { readModeEl.checked = true; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForTimeout(300);
  r = await tabRect(page, 0, 0);
  await page.mouse.click(r.x, r.y);
  await page.waitForTimeout(200);
  check("w Czytaniu klik w kropki nic nie pokazuje", await page.evaluate(() => !document.querySelector(".dwb-tabstop-handle")));

  await browser.close();
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const x of results) {
    if (!x.ok) failed++;
    console.log(`${x.ok ? "✅" : "❌"} ${x.name}${x.ok ? "" : `  (${x.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Długość kropek [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => {
  for (const x of results) console.log(`${x.ok ? "✅" : "❌"} ${x.name}${x.ok ? "" : `  (${x.detail || ""})`}`);
  console.error(e);
  process.exit(1);
});
