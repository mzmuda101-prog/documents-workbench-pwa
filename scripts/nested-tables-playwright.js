// nested-tables-playwright.js — tabela w tabeli i wcięcie tabeli (2026-10-04, DC-85).
//
// Dawniej: odczyt akapitów pliku liczył akapity tabeli zagnieżdżonej dwa razy — numery akapitów
// pliku rozjeżdżały się z podglądem, a zapis wpisywał tekst w cudze akapity; wyczyszczony akapit
// zapisywał się jako „undefined”; podgląd ignorował wcięcie tabeli (w:tblInd). Tu: dokument z
// tabelą w tabeli, akapitami pod nią i wciętą tabelą — po otwarciu żadnych „zmian”, edycje trafiają
// w swoje akapity, pusty akapit zostaje pusty, wcięcie widać w podglądzie. ENGINE=webkit.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  const browser = await pw[ENGINE].launch();
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForFunction(() => !!originalFileBytes && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });

  // dokument: Nagłówek / tabela [ komórka: „Lewa” + tabela { „W środku A” | „W środku B” } | „Prawa” ] / „Pod tabelą 1” / „Pod tabelą 2” / tabela wcięta 1 cal
  await page.evaluate(async () => {
    const z = await JSZip.loadAsync(originalFileBytes);
    const p = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
    const cell = (inner) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
    const inner = `<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr>${cell(p("W środku A"))}${cell(p("W środku B"))}</w:tr></w:tbl>`;
    const outer = `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell(p("Lewa") + inner + p(""))}${cell(p("Prawa"))}</w:tr></w:tbl>`;
    const indented = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblInd w:w="1440" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell(p("Wcięta tabela"))}</w:tr></w:tbl>`;
    const body = `<w:body>${p("Nagłówek")}${outer}${p("Pod tabelą 1")}${p("Pod tabelą 2")}${indented}${p("Koniec")}`;
    let doc = await z.file("word/document.xml").async("string");
    doc = doc.replace(/<w:body>[\s\S]*?(<w:sectPr[\s\S]*<\/w:body>)/, `${body}$1`);
    z.file("word/document.xml", doc);
    await ingestFile(new File([await z.generateAsync({ type: "uint8array" })], "zagniezdzone.docx"), { silent: true });
  });
  await page.waitForFunction(() => document.body.innerText.includes("Pod tabelą 2") && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 });
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForFunction(() => document.querySelector(".docx-editable-p"), null, { timeout: 10000 });
  await sleep(400);

  const st = await page.evaluate(async () => ({
    xml: (await extractParagraphTextsFromDocx(originalFileBytes)).join("|"),
    prev: collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => p.textContent).join("|"),
    edits: collectInlineParagraphEdits().length,
  }));
  check("tabela w tabeli: akapity pliku = akapity podglądu (każdy raz)", st.xml === st.prev && st.xml === "Nagłówek|Lewa|W środku A|W środku B||Prawa|Pod tabelą 1|Pod tabelą 2|Wcięta tabela|Koniec", `${st.xml}\n${st.prev}`);
  check("po otwarciu bez fałszywych zmian", st.edits === 0, st.edits);

  // edycje: dopisz w „Pod tabelą 2”, wyczyść „Pod tabelą 1”, dopisz w „W środku B”
  await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const by = (t) => ps.find((p) => p.textContent === t);
    by("Pod tabelą 2").append(" EDYTOWANE");
    by("Pod tabelą 1").replaceChildren();
    by("W środku B").append(" (zmiana)");
    ps.forEach((p) => p.dispatchEvent(new Event("input", { bubbles: true })));
  });
  const saved = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const doc = await z.file("word/document.xml").async("string");
    return { doc, texts: (doc.match(/<w:p[ >][\s\S]*?<\/w:p>|<w:p\/>/g) || []).map((p) => (p.match(/<w:t[^>]*>[^<]*/g) || []).map((t) => t.replace(/<w:t[^>]*>/, "")).join("")).join("|") };
  });
  check("zapis: każda zmiana w swoim akapicie, wyczyszczony akapit pusty (nie „undefined”)",
    saved.texts === "Nagłówek|Lewa|W środku A|W środku B (zmiana)||Prawa||Pod tabelą 2 EDYTOWANE|Wcięta tabela|Koniec" && !saved.doc.includes(">undefined<"), saved.texts);

  const ind = await page.evaluate(() => {
    const t = [...document.querySelectorAll("section.docx table")].find((x) => x.textContent.includes("Wcięta tabela"));
    const art = t.closest("article").getBoundingClientRect();
    return Math.round((t.getBoundingClientRect().left - art.left) * 0.75);
  });
  check("podgląd: wcięcie tabeli (w:tblInd 1 cal = 72 pt)", Math.abs(ind - 72) <= 1, ind);

  check("brak błędów JS", !errors.length, errors.slice(0, 3).join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`  ${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`);
  console.log(`\n${ENGINE}: ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
