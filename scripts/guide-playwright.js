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
  check("skróty sekcji: tytuł + 17 rozdziałów", r.chips === 18, String(r.chips));
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
  // Etap 2: linki i proste pola formularza w zdaniu nie blokują akapitu (model akapitu je zachowuje)
  // od 2026-10-04 akapit z odnośnikiem do przypisu też edytowalny (odnośnik = „wyspa”, doc-notes.js)
  // od 2026-10-05 także akapit z obrazem (obraz = wyspa, docx-inline-edit.js paragraphObjectLock)
  check("zablokowane: zmiana, odsyłacz (linki, pola w zdaniu, przypisy i obrazy edytowalne)",
    !r.locked.lockLink && !r.locked.lockForm && !r.locked.lockNote && !r.locked.lockObject && r.locked.lockField === 1 && r.locked.lockTracked === 1 && Object.keys(r.locked).length === 2, JSON.stringify(r.locked));
  check("formularz: tekst, lista, data, pole wyboru", r.forms === "text,dropdown,date,checkbox", r.forms);
  check("spis treści: 17 linków do rozdziałów (+ link do rozdz. 3 z rozdz. 16) + odsyłacz do „Zapisu”", r.toc === 18 && r.xref === "#_Guide12", `${r.toc} ${r.xref}`);

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

  // rozdział 10: najechanie na „Rysunek 1” = podgląd obrazu bez przewijania, klik = skok do obrazu
  await page.evaluate(() => document.querySelector(".docx-preview-host a[data-dwb-img-link]").scrollIntoView({ block: "center" }));
  await page.waitForTimeout(300);
  const st0 = await page.evaluate(() => docViewportEl.scrollTop);
  await page.hover(".docx-preview-host a[data-dwb-img-link]");
  await page.waitForTimeout(600);
  const peek = await page.evaluate(() => ({ cap: document.querySelector(".link-peek .link-peek-cap")?.textContent, ok: !!document.querySelector(".link-peek img")?.naturalWidth, st: docViewportEl.scrollTop }));
  check("rozdz. 10: najechanie na „Rysunek 1” pokazuje obraz z podpisem, dokument stoi", peek.ok && /^Rysunek 1\./.test(peek.cap || "") && peek.st === st0, JSON.stringify(peek));
  await page.click(".docx-preview-host a[data-dwb-img-link]");
  await page.waitForTimeout(900);
  check("rozdz. 10: klik w „Rysunek 1” przenosi do obrazu, jest „↩ Wróć”", await page.evaluate(() => { const i = document.querySelector(".docx-preview-host img").getBoundingClientRect(); const vp = docViewportEl.getBoundingClientRect(); return i.top >= vp.top - 2 && i.bottom <= vp.bottom + 2 && !document.querySelector(".link-back")?.hidden; }));
  // Edycja: klik w obraz = karta z suwakiem
  await page.evaluate(() => appFrame.setReadOnly(false));
  await page.waitForFunction(() => !inlineLocksPending, null, { timeout: 10000 });
  await page.waitForTimeout(400);
  await page.evaluate(() => document.querySelector(".docx-preview-host img").scrollIntoView({ block: "center" }));
  await page.click(".docx-preview-host img");
  check("rozdz. 10: w Edycji klik w obraz pokazuje kartę z suwakiem", await page.evaluate(() => !!document.querySelector(".image-card input[type=range]")));
  await page.keyboard.press("Escape");
  // rozdział 13: komentarz z pogrubieniem przyciskiem B z paska, ołówek, „Układ” → „Wąskie”
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => /^Przycisk „A” na pasku/.test(x.textContent)); p.scrollIntoView({ block: "center" }); const r = formDomRange(p, 10, 11); p.closest(".docx-edit-root").focus({ preventScroll: true }); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await page.waitForTimeout(400); // przewinięcie do akapitu (scroll zamyka okienka)
  await page.keyboard.press("Control+Alt+KeyM");
  await page.keyboard.type("Ważne: ");
  await page.click("#fmtBold");
  await page.keyboard.type("pogrubione");
  await page.fill(".cf-author", "Ola");
  await page.click(".compose-pop-form .lf-ok");
  await page.waitForFunction(() => !inlineLocksPending && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 10000 });
  await page.waitForTimeout(500);
  const cx = await page.evaluate(async () => (await (await JSZip.loadAsync(originalFileBytes)).file("word/comments.xml").async("string")));
  check("rozdz. 13: komentarz z pogrubieniem przyciskiem B z paska", /<w:b\/><\/w:rPr><w:t xml:space="preserve">pogrubione/.test(cx), cx.slice(-400));
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => /^Przycisk „A” na pasku/.test(x.textContent)); const r = formDomRange(p, 11, 11); p.closest(".docx-edit-root").focus({ preventScroll: true }); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await page.waitForTimeout(400);
  check("rozdz. 13: karta komentarza z ołówkiem (poprawianie treści)", await page.evaluate(() => !!document.querySelector(".comment-card .cc-edit")));
  await page.click("#pageLayoutBtn");
  await page.click('.compose-pop [data-preset="narrow"]');
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 10000 });
  await page.waitForTimeout(500);
  const sp = await page.evaluate(async () => (await (await JSZip.loadAsync(originalFileBytes)).file("word/document.xml").async("string")).match(/<w:pgMar [^>]*>/)?.[0]);
  check("rozdz. 13: „Układ” → „Wąskie” zmienia marginesy", /w:left="720"/.test(sp || ""), sp);

  // rozdział 15 (2026-10-29): karta z kropkami + pole tekstowe; ⌘/Ctrl przyciąga koniec kropek
  const card = await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    return ["Miejscowość:", "Teren nr:", "Data wydania:"].map((t) => ps.findIndex((p) => p.textContent.startsWith(t) && p.querySelector(".docx-tab")));
  });
  await page.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].scrollIntoView({ block: "center" }), card[0]);
  // przewijanie bywa płynne — mierzymy dopiero, gdy położenie stoi przez 3 klatki (przy pauzach ×0,5
  // stałe 500 ms wypadało w trakcie ruchu: klik w kropki trafiał obok i uchwyt się nie pokazywał)
  await page.evaluate(() => new Promise((res) => {
    const vp = document.getElementById("docViewport");
    let last = -1, still = 0, n = 0;
    const tick = () => { const y = vp ? vp.scrollTop : scrollY; still = y === last ? still + 1 : 0; last = y; if (still >= 3 || ++n > 180) res(); else requestAnimationFrame(tick); };
    requestAnimationFrame(tick);
  }));
  const tb = await page.evaluate((i) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    const el = [...document.querySelectorAll(".dwb-tb")].find((x) => x.__tbAnchor === p);
    if (!el) return null;
    const a = p.getBoundingClientRect(), r = el.getBoundingClientRect();
    return { text: el.textContent.trim(), dy: Math.round(r.top - a.top), right: r.left > a.left + a.width / 2 };
  }, card[0]);
  check("rozdz. 15: pole „ok. 34 m od drogi” obok karty (przy akapicie kotwicy, choć to dalsza strona)", card.every((i) => i > 0) && tb?.text === "ok. 34 m od drogi" && Math.abs(tb.dy) < 30 && tb.right, JSON.stringify({ card, tb }));
  const tabAt = (i) => page.evaluate((i) => { const r = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].querySelector(".docx-tab").getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2, r: r.right }; }, i);
  const goal = (await tabAt(card[0])).r; // koniec kropek „Miejscowość:”
  const dots = await tabAt(card[1]);
  await page.mouse.click(dots.x, dots.y);
  // uchwyt kropek pokazuje się po kliknięciu z opóźnieniem — czekamy na niego (w tłoku 150 ms nie starczało)
  await page.waitForFunction(() => { const e = document.querySelector(".dwb-tabstop-handle"); return e && e.style.visibility !== "hidden"; }, null, { timeout: 5000 }).catch(() => {});
  const hd = await page.evaluate(() => { const e = document.querySelector(".dwb-tabstop-handle"); if (!e || e.style.visibility === "hidden") return null; const r = e.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });
  const mod = process.platform === "darwin" ? "Meta" : "Control";
  let snapped = null;
  if (hd) {
    await page.keyboard.down(mod);
    await page.mouse.move(hd.x, hd.y);
    await page.mouse.down();
    await page.mouse.move(goal + 3, hd.y, { steps: 8 });
    const guideSeen = await page.evaluate(() => !!document.querySelector(".dwb-tabstop-guide"));
    await page.mouse.up();
    await page.keyboard.up(mod);
    await page.waitForTimeout(900);
    await page.waitForFunction(() => !inlineLocksPending && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 10000 });
    snapped = { guideSeen, end: (await tabAt(card[1])).r, goal };
  }
  check("rozdz. 15: kropki „Teren nr:” z ⌘/Ctrl kończą się równo z kropkami wiersza wyżej, linia pomocnicza", hd && snapped.guideSeen && Math.abs(snapped.end - goal) < 0.75, JSON.stringify({ hd, snapped }));
  await page.keyboard.press("Escape");

  // rozdział 16: pędzel z czerwonego fragmentu na „ten jest zwykły”
  const span = (text) => page.evaluate((text) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => x.textContent.includes("czeka na pędzel"));
    const at = p.textContent.indexOf(text);
    const r = formDomRange(p, at, at + text.length);
    const rs = r.getClientRects();
    return { a: at, p: true, x0: rs[0].left + 1, y0: rs[0].top + rs[0].height / 2, x1: rs[rs.length - 1].right - 1, y1: rs[rs.length - 1].top + rs[rs.length - 1].height / 2 };
  }, text);
  await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => x.textContent.includes("czeka na pędzel")).scrollIntoView({ block: "center" }));
  await page.waitForTimeout(400);
  const src = await span("fragment");
  await page.mouse.click((src.x0 + src.x1) / 2, src.y0);
  await page.waitForTimeout(200);
  await page.click("#fmtPainterBtn");
  await page.waitForTimeout(350);
  const dst = await span("ten jest zwykły");
  await page.mouse.move(dst.x0, dst.y0);
  await page.mouse.down();
  await page.mouse.move(dst.x1, dst.y1, { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(900);
  await page.waitForFunction(() => !inlineLocksPending && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 10000 });
  const painted = await page.evaluate(async () => {
    const xml = await (await JSZip.loadAsync(await buildDocumentForSave())).file("word/document.xml").async("string");
    const para = xml.split("</w:p>").find((x) => /czeka na pędzel/.test(x)) || "";
    const runs = para.split("</w:r>").filter((r) => /ten jest zwyk/.test(r));
    return runs.map((r) => ({ b: /<w:b\/>/.test(r), red: /C0392B/i.test(r) }));
  });
  check("rozdz. 16: malarz formatów — „ten jest zwykły” pogrubione i czerwone w pliku", painted.length > 0 && painted.every((r) => r.b && r.red), JSON.stringify(painted));

  // rozdział 17: „Zapytaj dokument” bez asystenta wskazuje zdanie z rozdziału 2
  const ask = await page.evaluate(async () => {
    document.getElementById("docAskQuery").value = "Do którego dnia najemca płaci czynsz?";
    await docAskRun();
    return document.getElementById("docAskAnswer").textContent;
  });
  check("rozdz. 17: „Zapytaj dokument” → zdanie o czynszu do dziesiątego dnia (rozdz. 2)", /dziesiątego dnia/.test(ask), ask.slice(0, 200));
  assertNoErrors(errors, "guide");
  await browser.close();
  let failed = 0;
  for (const x of results) { console.log(`${x.ok ? "✅" : "❌"} ${x.name}${!x.ok && x.detail ? `  (${x.detail})` : ""}`); if (!x.ok) failed++; }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ przewodnik: ${results.length}/${results.length}`);
}
run().catch((e) => { console.error(e); process.exit(1); });
