// @ts-check
// edit-hf-pages-playwright.js — Edycja jak wydruk: nagłówek i stopka na KAŻDEJ stronie dokumentu
// (kopie tylko do oglądania, page-breaks.js), numery „Strona N z M” jak w Wordzie i w podglądzie
// wydruku. Dawniej nagłówek był tylko nad stroną 1, a stopka tylko na końcu ostatniej strony
// z „Strona 1 z 1” (testerzy AI 2026-11-27, AUDYT-MOBILE #2).
//   node scripts/edit-hf-pages-playwright.js        ENGINE=webkit node scripts/edit-hf-pages-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const HDR = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FTR = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

(async () => {
  const lorem = "Najemca zobowiązuje się używać lokalu zgodnie z jego przeznaczeniem i dbać o jego stan techniczny, a wszelkie zmiany wymagają pisemnej zgody.";
  const bytes = await buildDocx({
    body: Array.from({ length: 70 }, (_, i) => `${i + 1}. ${lorem}`),
    sectPr: '<w:headerReference w:type="default" r:id="rIdX0"/><w:footerReference w:type="default" r:id="rIdX1"/>',
    extraParts: {
      "header1.xml": { type: HDR, rel: `${REL}header`, xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p><w:r><w:t>URZĄD GMINY KOWALEWO</w:t></w:r></w:p></w:hdr>` },
      "footer1.xml": { type: FTR, rel: `${REL}footer`, xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W}><w:p><w:r><w:t xml:space="preserve">Strona </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> z </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>` },
    },
  });
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbOpenShadow = true; });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "pismo.docx", mimeType: DOCX_MIME, buffer: bytes });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 });
  await page.evaluate(() => appFrame.setReadOnly(false));
  await page.waitForFunction(() => document.querySelectorAll(".docx-preview-host .dwb-page-gap-band").length >= 2, null, { timeout: 15000 });
  await page.waitForTimeout(1200);

  const st = await page.evaluate(() => {
    const sec = document.querySelector(".docx-preview-host section.docx");
    const bands = [...sec.querySelectorAll(":scope > .dwb-page-gap-band")].map((b) => parseFloat(b.style.top));
    const copies = [...sec.querySelectorAll(":scope > .dwb-page-hf")].map((c) => ({ kind: c.dataset.kind, page: Number(c.dataset.page), top: parseFloat(c.style.top), text: (c.shadowRoot?.textContent || "").replace(/\s+/g, " ").trim() }));
    const txt = (sel) => (sec.querySelector(sel)?.textContent || "").replace(/\s+/g, " ").trim();
    const secTop = sec.getBoundingClientRect().top;
    const scale = sec.getBoundingClientRect().height / sec.offsetHeight || 1;
    const hdr = sec.querySelector(":scope > header:not(.dwb-hf-alt)");
    return { bands, copies, header: txt(":scope > header:not(.dwb-hf-alt)"), footer: txt(":scope > footer:not(.dwb-hf-alt)"), headerTop: hdr ? (hdr.getBoundingClientRect().top - secTop) / scale : null, pageH: parseFloat(sec.dataset.dwbPageH) };
  });
  const pages = st.bands.length + 1;
  const heads = st.copies.filter((c) => c.kind === "header").sort((a, b) => a.page - b.page);
  const foots = st.copies.filter((c) => c.kind === "footer").sort((a, b) => a.page - b.page);
  check("dokument ma ≥ 3 strony w Edycji", pages >= 3, JSON.stringify(st.bands));
  check("nagłówek na każdej stronie od 2.", heads.length === pages - 1 && heads.every((h, i) => h.page === i + 2 && h.text === "URZĄD GMINY KOWALEWO"), JSON.stringify(heads));
  check("nagłówek strony k+1 tyle samo pod górą strony co na stronie 1", heads.every((h, i) => Math.abs(h.top - (st.bands[i] + 22 + st.headerTop)) < 2), JSON.stringify({ heads: heads.map((h) => h.top), bands: st.bands, headerTop: st.headerTop }));
  check("stopka na każdej stronie przed ostatnią, z jej numerem", foots.length === pages - 1 && foots.every((f, i) => f.page === i + 1 && f.text === `Strona ${i + 1} z ${pages}`), JSON.stringify(foots));
  check("stopki kopii na dole swojej strony (nad pasem przerwy)", foots.every((f, i) => f.top < st.bands[i] && f.top > st.bands[i] - st.pageH / 4), JSON.stringify({ foots: foots.map((f) => f.top), bands: st.bands }));
  check("oryginalna stopka (ostatnia strona) ma ostatni numer i liczbę stron", st.footer === `Strona ${pages} z ${pages}`, st.footer);
  check("oryginalny nagłówek bez zmian", st.header === "URZĄD GMINY KOWALEWO", st.header);
  // kopie nie są treścią dokumentu: wyszukiwanie i akapity ich nie widzą
  const leak = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => /URZĄD GMINY|Strona \d+ z/.test(p.textContent)).length);
  check("kopie poza akapitami dokumentu", leak === 0, String(leak));
  // dopisanie strony: liczby się aktualizują
  await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const p = ps[ps.length - 1];
    p.closest(".docx-edit-root").focus();
    placeCaret(p, p.textContent.length);
  });
  for (let i = 0; i < 25; i++) await page.keyboard.press("Enter");
  await page.waitForTimeout(1500);
  const after = await page.evaluate(() => {
    const sec = document.querySelector(".docx-preview-host section.docx");
    return { pages: sec.querySelectorAll(":scope > .dwb-page-gap-band").length + 1, foots: [...sec.querySelectorAll(":scope > .dwb-page-hf[data-kind=footer]")].map((c) => (c.shadowRoot?.textContent || "").replace(/\s+/g, " ").trim()), footer: (sec.querySelector(":scope > footer:not(.dwb-hf-alt)")?.textContent || "").replace(/\s+/g, " ").trim() };
  });
  check("po dopisaniu strony: numery i liczba stron aktualne", after.pages > pages && after.foots.length === after.pages - 1 && after.foots.every((t, i) => t === `Strona ${i + 1} z ${after.pages}`) && after.footer === `Strona ${after.pages} z ${after.pages}`, JSON.stringify(after));
  // spokój: bez zmian w treści przeliczanie stron nie przerysowuje kopii w kółko
  const churn = await page.evaluate(async () => {
    const sec = document.querySelector(".docx-preview-host section.docx");
    let n = 0;
    const mo = new MutationObserver((ms) => { for (const m of ms) n += [...m.addedNodes].filter((x) => x.nodeType === 1 && x.matches(".dwb-page-hf, .dwb-page-gap-band")).length; });
    mo.observe(sec, { childList: true });
    window.dispatchEvent(new Event("resize"));
    await new Promise((r) => setTimeout(r, 1500));
    mo.disconnect();
    return n;
  });
  check("bez zmian treści: kopie i pasy nie są wstawiane od nowa", churn === 0, String(churn));
  // w spoczynku kartka stoi: żadnych zapisów (numery stron wpisywane w kółko budziły przeliczanie
  // stron co 150 ms — psuło m.in. kursor w pustym akapicie pod tabelą)
  const idleMuts = await page.evaluate(async () => {
    const sec = document.querySelector(".docx-preview-host section.docx");
    let n = 0;
    const mo = new MutationObserver((ms) => { n += ms.length; });
    mo.observe(sec, { childList: true, subtree: true, characterData: true });
    await new Promise((r) => setTimeout(r, 1500));
    mo.disconnect();
    return n;
  });
  check("w spoczynku żadnych zmian w kartce (bez pętli przeliczania)", idleMuts === 0, String(idleMuts));
  // nagłówek wyższy niż górny margines (w:header daleko od krawędzi): Word spycha tekst KAŻDEJ strony
  // pod nagłówek; kopia nagłówka nie może zakrywać tekstu strony (dawniej leżała na pierwszej linijce)
  const tall = await buildDocx({
    body: Array.from({ length: 70 }, (_, i) => `${i + 1}. ${lorem}`),
    sectPr: '<w:headerReference w:type="default" r:id="rIdX0"/>',
    extraParts: { "header1.xml": { type: HDR, rel: `${REL}header`, xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p><w:r><w:t>NAGŁÓWEK</w:t></w:r></w:p><w:p><w:r><w:t>druga linia nagłówka</w:t></w:r></w:p><w:p><w:r><w:t>trzecia linia nagłówka</w:t></w:r></w:p><w:p><w:r><w:t>czwarta linia nagłówka</w:t></w:r></w:p></w:hdr>` } },
  });
  await page.locator("#fileInput").setInputFiles({ name: "wysoki-naglowek.docx", mimeType: DOCX_MIME, buffer: tall });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 });
  await page.evaluate(() => appFrame.setReadOnly(false));
  await page.waitForFunction(() => document.querySelectorAll(".docx-preview-host .dwb-page-hf").length >= 1, null, { timeout: 15000 });
  await page.waitForTimeout(1000);
  const over = await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    return [...document.querySelectorAll(".docx-preview-host .dwb-page-hf[data-kind=header]")].map((c) => {
      const r = c.getBoundingClientRect();
      // linijki tekstu (akapit podzielony między strony ma prostokąt przez odstęp — liczą się linijki)
      const hit = ps.find((p) => { const rg = document.createRange(); rg.selectNodeContents(p); return [...rg.getClientRects()].some((q) => q.height > 2 && q.bottom > r.top + 1 && q.top < r.bottom - 1); });
      return { top: Math.round(r.top), bottom: Math.round(r.bottom), hit: hit ? hit.textContent.slice(0, 30) : null };
    });
  });
  check("wysoki nagłówek: kopia nagłówka nie zakrywa tekstu strony (tekst pod nagłówkiem, jak w Wordzie)", over.length >= 1 && over.every((o) => !o.hit), JSON.stringify(over));
  check("bez błędów na stronie", !errors.length, errors.join(" | "));

  await browser.close();
  const bad = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  ${String(r.detail).slice(0, 400)}`}`));
  console.log(`${bad.length ? "❌" : "✅"} Nagłówki i stopki stron w Edycji [${ENGINE}]: ${results.length - bad.length}/${results.length}`);
  process.exit(bad.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
