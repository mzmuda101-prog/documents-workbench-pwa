// @ts-check
// edit-footnotes-pages-playwright.js — Edycja jak wydruk (#7): przypisy dolne na dole SWOJEJ
// strony (zamiast wszystkie na końcu ostatniej strony), liczba stron w Edycji = podgląd wydruku,
// przypisy pozostają edytowalne na swojej stronie, zapis zachowuje zmiany, bez pętli przeliczania.
//   node scripts/edit-footnotes-pages-playwright.js        ENGINE=webkit node scripts/edit-footnotes-pages-playwright.js
const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const FN_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const FN_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

(async () => {
  const lorem = "W niniejszym rozdziale omówiono zasady sporządzania dokumentacji projektowej oraz sposób weryfikacji jej zgodności z wymaganiami technicznymi.";
  // 2-stronicowy tekst z 3 przypisami: przypis 1 na stronie 1, przypis 2 pod koniec strony 1/2, przypis 3 dalej
  const body = Array.from({ length: 46 }, (_, i) => {
    if (i === 4) return `<w:p><w:r><w:t xml:space="preserve">${i + 1}. ${lorem}</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>`;
    if (i === 24) return `<w:p><w:r><w:t xml:space="preserve">${i + 1}. ${lorem}</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p>`;
    if (i === 40) return `<w:p><w:r><w:t xml:space="preserve">${i + 1}. ${lorem}</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="3"/></w:r></w:p>`;
    return `${i + 1}. ${lorem}`;
  });
  const fnXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W}>`
    + `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>`
    + `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>`
    + `<w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Pierwszy przypis ze strony pierwszej — dość długi opis, aby zająć miejsce na dole kartki.</w:t></w:r></w:p></w:footnote>`
    + `<w:footnote w:id="2"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Drugi przypis ze strony drugiej — również zajmuje pełny wiersz u dołu.</w:t></w:r></w:p></w:footnote>`
    + `<w:footnote w:id="3"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Trzeci przypis pod koniec dokumentu.</w:t></w:r></w:p></w:footnote>`
    + `</w:footnotes>`;
  const bytes = await buildDocx({
    body,
    extraParts: { "footnotes.xml": { type: FN_TYPE, rel: FN_REL, xml: fnXml } },
  });

  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "przypisy.docx", mimeType: DOCX_MIME, buffer: bytes });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 });
  await page.evaluate(() => appFrame.setReadOnly(false));
  await page.waitForFunction(() => document.querySelectorAll(".docx-preview-host .dwb-page-gap-band").length >= 1, null, { timeout: 15000 });
  await page.waitForTimeout(1000);

  // 1) Położenie przypisów w Edycji: przypis 1 na stronie 1 (nad pierwszym pasem przerwy),
  //    przypis 3 na ostatniej stronie (pod pasem przerwy), bez nakładania się na tekst treści
  const st = await page.evaluate(() => {
    const sec = document.querySelector(".docx-preview-host section.docx");
    const sr = sec.getBoundingClientRect();
    const scale = sr.height / sec.offsetHeight || 1;
    const Y = (v) => Math.round(((v - sr.top) / scale) * 10) / 10;
    const bands = [...sec.querySelectorAll(":scope > .dwb-page-gap-band")].map((b) => parseFloat(b.style.top));
    const notePos = (id) => {
      const li = sec.querySelector(`ol.dwb-notes > li[data-dwb-note="${id}"]`);
      if (!li) return null;
      const r = li.getBoundingClientRect();
      return { top: Y(r.top), bottom: Y(r.bottom), editable: !!li.querySelector("p")?.isContentEditable };
    };
    const refPos = (id) => {
      const sup = sec.querySelector(`sup[data-dwb-note="${id}"]`);
      if (!sup) return null;
      const r = sup.getBoundingClientRect();
      return { top: Y(r.top), bottom: Y(r.bottom) };
    };
    const pageOf = (y) => { let p = 0; bands.forEach((b) => { if (y > b) p++; }); return p; };
    const n1 = notePos("footnote:1"), n2 = notePos("footnote:2"), n3 = notePos("footnote:3");
    const r1 = refPos("footnote:1"), r2 = refPos("footnote:2"), r3 = refPos("footnote:3");
    // czy przypis 1 nie nachodzi na ostatni akapit strony 1
    const bodyLinesP1 = [...sec.querySelectorAll(":scope > article p")]
      .map((p) => { const r = p.getBoundingClientRect(); return { top: Y(r.top), bottom: Y(r.bottom) }; })
      .filter((x) => x.bottom < bands[0] + 1 && x.top < bands[0]);
    const maxBodyP1 = bodyLinesP1.length ? Math.max(...bodyLinesP1.map((x) => x.bottom)) : 0;
    return {
      editPages: bands.length + 1,
      bands,
      n1, n2, n3,
      samePage1: n1 && r1 && pageOf(n1.top) === pageOf(r1.top) && pageOf(n1.top) === 0,
      samePage2: n2 && r2 && pageOf(n2.top) === pageOf(r2.top),
      samePage3: n3 && r3 && pageOf(n3.top) === pageOf(r3.top),
      noOverlap1: n1 && n1.top >= maxBodyP1 - 1,
      maxBodyP1,
    };
  });
  check("Edycja ma ≥ 2 strony", st.editPages >= 2, JSON.stringify(st));
  check("przypis 1 stoi na stronie 1 (tej samej co odnośnik 1), nad pasem przerwy", st.samePage1 && st.n1.bottom <= st.bands[0], JSON.stringify(st));
  check("przypisy 2 i 3 stoją na swoich stronach (zgodnych z odnośnikami)", st.samePage2 && st.samePage3, JSON.stringify(st));
  check("przypis 1 nie nachodzi na treść strony 1", st.noOverlap1, JSON.stringify({ n1: st.n1, maxBodyP1: st.maxBodyP1 }));
  check("przypisy pozostają edytowalne na swoich stronach", st.n1?.editable && st.n2?.editable && st.n3?.editable, JSON.stringify({ n1: st.n1, n2: st.n2, n3: st.n3 }));

  // 2) Edycja przypisu 1 na stronie 1 + zapis do pliku + sprawdzenie spójności
  await page.evaluate(() => {
    const p = document.querySelector('ol.dwb-notes > li[data-dwb-note="footnote:1"] > p');
    placeCaret(p, p.textContent.length);
  });
  await page.keyboard.type(" DOPISANO");
  await page.waitForTimeout(500);
  const saved = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const fn = await z.file("word/footnotes.xml").async("string");
    const ck = (await dwbSelfCheck()).problems.filter((x) => x.level === "error").map((x) => x.msg);
    return { hasEdit: fn.includes("DOPISANO"), ck };
  });
  check("edycja przypisu 1 na stronie 1 trafia do footnotes.xml, dwbSelfCheck czysty", saved.hasEdit && !saved.ck.length, JSON.stringify(saved));

  // 3) Liczba stron w Edycji zgadza się z Podglądem wydruku
  await page.evaluate(async () => {
    if (typeof ensureLazyFeature === "function") await ensureLazyFeature("print-preview");
    await dwbPrint.open();
  });
  await page.waitForFunction(() => document.querySelectorAll(".pp-overlay .pp-sheet").length >= 1, null, { timeout: 30000 });
  const ppPages = await page.evaluate(() => document.querySelectorAll(".pp-overlay .pp-sheet").length);
  await page.evaluate(() => document.querySelector(".pp-overlay .pp-close")?.click());
  check("liczba stron w Edycji = liczba stron w Podglądzie wydruku", st.editPages === ppPages, `edit=${st.editPages} print=${ppPages}`);

  // 4) Brak pętli przeliczania w spoczynku
  await page.waitForTimeout(600);
  const idleMuts = await page.evaluate(async () => {
    const sec = document.querySelector(".docx-preview-host section.docx");
    let n = 0;
    const mo = new MutationObserver((ms) => { n += ms.length; });
    mo.observe(sec, { childList: true, subtree: true, characterData: true });
    await new Promise((r) => setTimeout(r, 1200));
    mo.disconnect();
    return n;
  });
  check("w spoczynku żadnych zmian w kartce (bez pętli przeliczania)", idleMuts === 0, String(idleMuts));
  check("bez błędów na stronie", !errors.length, errors.join(" | "));

  await browser.close();
  const bad = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  ${String(r.detail).slice(0, 350)}`}`));
  console.log(`${bad.length ? "❌" : "✅"} Przypisy na stronach w Edycji [${ENGINE}]: ${results.length - bad.length}/${results.length}`);
  process.exit(bad.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
