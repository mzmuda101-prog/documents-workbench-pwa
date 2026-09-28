// find-replace-v2-playwright.js — Znajdź i zamień v2 (app/find-replace-workbench.js + silnik
// w docx-patch.js). Dokument budowany w teście, z przypadkami, które dawniej się psuły:
//   - słowo rozcięte między fragmenty tekstu (<w:r>…Naj</w:r><w:r>emca…) — znajdowane, nie zamieniane,
//   - wielkość liter: szukanie ignorowało, zamiana nie → „Zamień bieżące” trafiało w inne wystąpienie,
//   - nagłówek strony, całe słowa, $1 w wyrażeniach, „$100” w wartości pola {{…}}.
//
//   node scripts/find-replace-v2-playwright.js                (Chromium)
//   ENGINE=webkit node scripts/find-replace-v2-playwright.js  (WebKit)

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  page.on("dialog", (d) => d.accept());

  // ── dokument testowy ───────────────────────────────────────────────────────
  await page.evaluate(async () => {
    await ensureDocLibs(false);
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const r = (t, b) => `<w:r>${b ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${t}</w:t></w:r>`;
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${r("Najemca i najemca")}</w:p>
<w:p>${r("Naj")}${r("emca", true)}${r(" płaci czynsz. Najemca dba o lokal.")}</w:p>
<w:p>${r("najemca pisany małą; Najemcami się nie liczy w całych słowach.")}</w:p>
<w:p>${r("Data 2026-09-28, kwota {{kwota}}.")}</w:p>
<w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr></w:body></w:document>`;
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`);
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`);
    zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style></w:styles>`);
    zip.file("word/document.xml", doc);
    zip.file("word/header1.xml", `<?xml version="1.0" encoding="UTF-8"?><w:hdr ${W}><w:p>${r("Najemca — umowa")}</w:p></w:hdr>`);
    const blob = await zip.generateAsync({ type: "blob" });
    await ingestFile(new File([blob], "fr-v2.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
    setSidebarOpen(true);
    document.getElementById("panel-search").open = true;
    await ensureLazyFeature("find-replace");
  });
  await page.waitForTimeout(500);

  const body = () => page.evaluate(async () => {
    const z = await JSZip.loadAsync(originalFileBytes);
    const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
    const h = new DOMParser().parseFromString(await z.file("word/header1.xml").async("string"), "application/xml");
    return { paras: collectParagraphElements(d.documentElement, "all").map(paragraphSearchText), header: paragraphSearchText(h.getElementsByTagNameNS(W_NS, "p")[0]) };
  });
  const scan = (opts) => page.evaluate(async (o) => {
    document.getElementById("searchQuery").value = o.find;
    document.getElementById("frReplace").value = o.replace ?? "";
    document.getElementById("frMatchCase").checked = !!o.matchCase;
    document.getElementById("frWholeWord").checked = !!o.wholeWord;
    document.getElementById("frRegex").checked = !!o.regex;
    document.getElementById("frOtherParts").checked = !!o.otherParts;
    document.getElementById("searchScope").value = o.scope || "all";
    await runFindReplaceScan();
    return { n: frMatches.length, other: frOtherCount, hl: window.CSS?.highlights?.get("dwb-find")?.size ?? -1, status: document.getElementById("frStatus").textContent };
  }, opts);

  // ── liczenie ───────────────────────────────────────────────────────────────
  let s = await scan({ find: "najemca" });
  check("bez wielkości liter: 6 trafień w treści (w tym rozcięte „Naj|emca” i „Najemcami”)", s.n === 6, JSON.stringify(s));
  check("+1 w nagłówku strony pokazane osobno", s.other === 1 && /nagłówkach/.test(s.status), s.status);
  check("dokładne podświetlenie: tyle zakresów, ile trafień (nie całe akapity)", s.hl === 6, String(s.hl));
  s = await scan({ find: "Najemca", matchCase: true });
  check("wielkość liter: 4 (bez „najemca”)", s.n === 4, JSON.stringify(s));
  s = await scan({ find: "najemca", wholeWord: true });
  check("całe słowa: 5 (bez „Najemcami”)", s.n === 5, JSON.stringify(s));
  s = await scan({ find: "najemca", scope: "headings" });
  check("tylko nagłówki sekcji: 2 (w tytule)", s.n === 2, JSON.stringify(s));
  s = await scan({ find: "(", regex: true });
  check("błędne wyrażenie: 0, bez wyjątku", s.n === 0);

  // ── „Zamień bieżące” trafia w PODŚWIETLONE wystąpienie ─────────────────────
  await scan({ find: "najemca" });
  await page.evaluate(() => focusFrMatch(2)); // 3. trafienie = rozcięte „Naj|emca” w akapicie 2
  const active = await page.evaluate(() => ({ p: frMatches[frActiveIndex].paraIndex, t: frMatches[frActiveIndex].text }));
  await page.evaluate(() => { document.getElementById("frReplace").value = "X"; });
  await page.evaluate(() => replaceFrOne());
  await page.waitForTimeout(400);
  let b = await body();
  check("„Zamień bieżące” zamienia podświetlone (rozcięte między fragmenty) wystąpienie",
    active.p === 1 && b.paras[1].startsWith("X płaci czynsz. Najemca") && b.paras[0] === "Najemca i najemca", JSON.stringify({ active, p: b.paras.slice(0, 2) }));
  check("po zamianie bieżącego: licznik 5 i aktywne jest następne", await page.evaluate(() => frMatches.length === 5 && frActiveIndex === 2));

  // ── Zamień wszystkie: całe słowa, z nagłówkiem strony ──────────────────────
  await scan({ find: "najemca", replace: "Lokator", wholeWord: true, otherParts: true });
  await page.evaluate(() => replaceFrAll());
  await page.waitForTimeout(500);
  b = await body();
  check("zamień wszystkie (całe słowa): „Najemcami” zostaje, reszta zamieniona",
    b.paras[0] === "Lokator i Lokator" && b.paras[1] === "X płaci czynsz. Lokator dba o lokal." && /^Lokator pisany małą; Najemcami/.test(b.paras[2]), JSON.stringify(b.paras));
  check("…i w nagłówku strony (opcja zaznaczona)", b.header === "Lokator — umowa", b.header);

  // ── wyrażenia: $1 ──────────────────────────────────────────────────────────
  await scan({ find: "(\\d{4})-(\\d{2})-(\\d{2})", replace: "$3.$2.$1", regex: true });
  await page.evaluate(() => replaceFrAll());
  await page.waitForTimeout(400);
  b = await body();
  check("wyrażenia: $3.$2.$1 zamienia datę na 28.09.2026", /Data 28\.09\.2026,/.test(b.paras[3]), b.paras[3]);

  // ── wartość pola z „$” zostaje dosłowna ─────────────────────────────────────
  await page.evaluate(() => applyDocumentEdit({ op: "placeholderFill", values: { kwota: "$100 i $& i $1" }, scope: "all" }));
  await page.waitForTimeout(400);
  b = await body();
  check("pole {{kwota}} = „$100 i $& i $1” wstawione dosłownie", /kwota \$100 i \$& i \$1\./.test(b.paras[3]), b.paras[3]);

  // ── tekst wpisany w podglądzie (jeszcze nie w pliku) też jest szukany i zamieniany ──
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change", { bubbles: true })); });
  await page.waitForTimeout(400);
  await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[2];
    p.focus();
    const r = document.createRange(); r.selectNodeContents(p); r.collapse(false);
    const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
  });
  await page.keyboard.type(" Zwierzak");
  await page.waitForTimeout(200);
  s = await scan({ find: "zwierzak" });
  check("świeżo wpisane (nie w pliku) — znalezione i podświetlone", s.n === 1 && s.hl === 1, JSON.stringify(s));
  await page.evaluate(() => { document.getElementById("frReplace").value = "Kot"; });
  await page.evaluate(() => replaceFrOne());
  await page.waitForTimeout(500);
  b = await body();
  check("…i „Zamień bieżące” zamienia je w pliku", /Najemcami.* Kot$/.test(b.paras[2]) && !/Zwierzak/.test(b.paras[2]), b.paras[2]);
  await page.evaluate(() => { readModeEl.checked = true; readModeEl.dispatchEvent(new Event("change", { bubbles: true })); });

  // ── historia fraz ──────────────────────────────────────────────────────────
  const hist = await page.evaluate(() => [...document.querySelectorAll("#frFindHistory option")].map((o) => o.value));
  check("historia wyszukiwań podpowiada ostatnie frazy", hist[0] === "zwierzak" && hist[1] === "(\\d{4})-(\\d{2})-(\\d{2})" && hist.includes("najemca"), hist.join(" | "));

  // ── Esc czyści dokładne podświetlenie ──────────────────────────────────────
  await scan({ find: "Lokator" });
  await page.evaluate(() => document.getElementById("searchQuery").blur());
  await page.keyboard.press("Escape");
  await page.waitForTimeout(100);
  check("Esc zdejmuje podświetlenie trafień", await page.evaluate(() => !(window.CSS?.highlights?.has?.("dwb-find"))));

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ znajdź i zamień v2 [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
