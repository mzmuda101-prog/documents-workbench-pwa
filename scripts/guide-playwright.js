// guide-playwright.js — przewodnik (docs/samples/przewodnik.docx, przycisk „Przykład”):
// każda wskazówka „Spróbuj: …” musi dawać dokładnie to, co obiecuje tekst.
//   node scripts/gen-guide-docx.mjs   (po zmianie treści)
const { createTestPage, bootApp, assertNoErrors } = require("./docx-test-helpers");
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function run() {
  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await page.click("#emptySampleBtn");
  await page.waitForFunction(() => currentFileName === "przewodnik.docx", null, { timeout: 15000 });
  await page.waitForTimeout(800);
  const r = await page.evaluate(async () => {
    const out = {};
    out.chips = document.querySelectorAll("#sectionChips .section-chip").length;
    out.words = documentStructure.words;
    out.review = docReviewCounts;
    await ensureLazyFeature("find-replace");
    const scan = async (o) => {
      searchQueryEl.value = "najemca";
      document.getElementById("frMatchCase").checked = !!o.mc;
      document.getElementById("frWholeWord").checked = !!o.ww;
      await runFindReplaceScan();
      return frMatches.map((m) => m.context.slice(m.start, m.start + m.text.length + 2));
    };
    out.find = [await scan({}), await scan({ ww: true }), await scan({ mc: true })];
    await ensureLazyFeature("placeholders");
    await runPlaceholderScan();
    out.fields = placeholderScan.fields.map((f) => f.name);
    await ensureLazyFeature("grammar");
    await runGrammarScan();
    out.grammar = Object.keys(grammarScan.byRule || {}).sort();
    out.grammarTexts = grammarScan.hits.map((h) => h.before);
    await ensureLazyFeature("snippets-panel");
    await runSnippetScan();
    out.triggers = snippetScan.triggers.map((t) => `${t.name}:${t.hasDefinition}`);
    await ensureLazyFeature("stats");
    document.getElementById("panel-stats").open = true;
    renderDocumentStats();
    out.longest = document.querySelector("#stLongSentences .stat-row span")?.textContent || "";
    await ensureLazyFeature("review");
    await runReviewScan();
    out.notes = reviewScan.notes.length;
    await ensureLazyFeature("metadata");
    await loadMetadataFromDocument();
    out.author = document.getElementById("metaCreator").value;
    out.tables = documentStructure.tables;
    const host = document.querySelector(".docx-preview-host");
    out.lists = host.querySelectorAll("[class*='docx-num-']").length;
    out.locked = collectPreviewParagraphElements(host).filter((p) => p.dataset.lock).reduce((m, p) => { m[p.dataset.lock] = (m[p.dataset.lock] || 0) + 1; return m; }, {});
    out.forms = formScan.fields.map((f) => f.kind).join(",");
    out.toc = host.querySelectorAll('a[href^="#_Guide"]:not(.doc-xref)').length;
    out.xref = host.querySelector("a.doc-xref")?.getAttribute("href");
    return out;
  });
  check("skróty sekcji: tytuł + 12 rozdziałów", r.chips === 13, String(r.chips));
  const [all, ww, mc] = r.find;
  check("szukanie „najemca” znajduje też „Najemcami”", all.some((t) => /^Najemcami/.test(t)), all.join("|"));
  check("„Tylko całe słowa” wyklucza „Najemcami”", ww.length === all.length - all.filter((t) => /^Najemcami/i.test(t)).length && !ww.some((t) => /^najemcami/i.test(t)), ww.join("|"));
  check("„Rozróżniaj wielkość liter” zostawia tylko pisane jak w polu (małą literą)", mc.length > 0 && mc.every((t) => t.startsWith("najemc")) && mc.length < all.length, mc.join("|"));
  check("placeholdery: 5 pól", r.fields.length === 5 && r.fields.includes("kwota"), r.fields.join());
  check("korekta: podwójna spacja, spacja przed :, wielokropek, cudzysłowy, wielka litera", ["cap-after-period", "double-space", "ellipsis", "quotes-pl", "space-before-punct"].every((k) => r.grammar.includes(k)), r.grammar.join());
  check("korekta nie rusza „sp. z o.o.”", !r.grammarTexts.some((t) => /^z$/.test(t)), r.grammarTexts.join("|"));
  check("snippety: trigger !podpis bez definicji", r.triggers.includes("podpis:false"), r.triggers.join());
  check("recenzja: 2 zmiany, 1 komentarz, 1 przypis", r.review.changes === 2 && r.review.comments === 1 && r.notes === 1, JSON.stringify(r.review));
  check("statystyki: najdłuższe zdanie to celowo długie", /celowo bardzo długie/.test(r.longest), r.longest);
  check("metadane: autor „Jan Przykładowy”", r.author === "Jan Przykładowy", r.author);
  check("tabela i listy są", r.tables === 1 && r.lists >= 4, `${r.tables} tab, ${r.lists} list`);
  check("zablokowane: zmiana, przypis, 12 wpisów spisu (linki), 4 pola formularza, odsyłacz",
    r.locked.lockLink === 12 && r.locked.lockForm === 4 && r.locked.lockField === 1 && r.locked.lockTracked === 1 && r.locked.lockNote === 1 && Object.keys(r.locked).length === 5, JSON.stringify(r.locked));
  check("formularz: tekst, lista, data, pole wyboru", r.forms === "text,dropdown,date,checkbox", r.forms);
  check("spis treści: 12 linków do rozdziałów + odsyłacz do „Zapisu”", r.toc === 12 && r.xref === "#_Guide12", `${r.toc} ${r.xref}`);

  // ── „Spróbuj:” z nowych rozdziałów robi to, co obiecuje ────────────────────
  const clickIn = async (sel) => {
    await page.evaluate((s) => document.querySelector(s).scrollIntoView({ block: "center" }), sel);
    await page.waitForTimeout(150);
    await page.click(sel);
    await page.waitForTimeout(900);
  };
  const headingInView = (n) => page.evaluate((n) => {
    const h = document.querySelector(`.docx-preview-host [id="_Guide${n}"]`)?.closest("p");
    const vp = docViewportEl.getBoundingClientRect(); const r = h.getBoundingClientRect();
    return r.top >= vp.top - 2 && r.bottom <= vp.bottom + 2;
  }, n);
  await clickIn('.docx-preview-host a[href="#_Guide4"]:not(.doc-xref)');
  check("spis treści: klik w „4. Formularz Worda” przewija do rozdziału, jest „↩ Wróć”", await headingInView(4) && await page.evaluate(() => !document.querySelector(".link-back")?.hidden));
  await clickIn('.docx-preview-host a.doc-xref');
  check("odsyłacz z rozdziału 10 przewija do „12. Zapis…”", await headingInView(12));
  const ck = await page.evaluate(() => formScan.fields.find((f) => f.kind === "checkbox").key);
  await clickIn(`.ff-field[data-ff="${ck}"]`);
  await page.waitForFunction(() => formScan?.bytes === originalFileBytes, null, { timeout: 10000 });
  check("formularz: klik w ☐ zaznacza pole", await page.evaluate((k) => formScan.fields.find((f) => f.key === k).value === true, ck));
  assertNoErrors(errors, "guide");
  await browser.close();
  let failed = 0;
  for (const x of results) { console.log(`${x.ok ? "✅" : "❌"} ${x.name}${!x.ok && x.detail ? `  (${x.detail})` : ""}`); if (!x.ok) failed++; }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ przewodnik: ${results.length}/${results.length}`);
}
run().catch((e) => { console.error(e); process.exit(1); });
