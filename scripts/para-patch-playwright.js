// para-patch-playwright.js — wyrównanie / wcięcie / odstępy BEZ przerysowania dokumentu
// (app/document.js: patchParagraphsInPlace). Niezmiennik: ŁATKA = PEŁNY RENDER — po każdej
// zmianie zapisany plik jest renderowany w tle tak jak przy pełnym przerysowaniu, a zmienione
// akapity porównane z żywymi (styl, klasy z renderu, wyrównanie, marginesy, wcięcie, interlinia,
// wysokość). Do tego: bez przerysowania, kursor zostaje, podgląd = plik, Cofnij przywraca.
//
//   node scripts/para-patch-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/para-patch-playwright.js
//   MOBILE=1 node scripts/para-patch-playwright.js        (telefon: Widok mobilny)
const fs = require("fs");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const SAMPLE = path.resolve(__dirname, "../docs/samples/przewodnik.docx");
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 }).then(() => page.waitForTimeout(500));

// żywe akapity vs pełny render zapisanego pliku (w tle, te same klasy płótna i szerokość)
const compareWithFullRender = (page, idx) => page.evaluate(async (indices) => {
  const APP = new Set(["docx-editable-p", "dwb-caret-p", "ff-para", "docx-locked-p", "dwb-tiny-active"]);
  const MEASURED = /^--dwb-(trail|last-fill|spw)$/; // pomiary na miejscu — porównujemy ich skutek (wymiary)
  const styleOf = (p) => [...p.style].filter((k) => !MEASURED.test(k)).map((k) => `${k}:${p.style.getPropertyValue(k)}`).sort().join(";");
  const classesOf = (p) => [...p.classList].filter((c) => !APP.has(c)).sort().join(" ");
  // + krój tekstu (pierwszy fragment) — styl akapitu z listy bez reguł CSS stylu wyszedłby tu inaczej
  const look = (p) => { const cs = getComputedStyle(p); const tw = document.createTreeWalker(p, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.data.trim() ? 1 : 3) }); const t = getComputedStyle(tw.nextNode()?.parentElement || p); return ["textAlign", "marginLeft", "marginRight", "marginTop", "marginBottom", "textIndent", "lineHeight", "paddingLeft"].map((k) => `${k}=${cs[k]}`).concat(["fontSize", "fontWeight", "color"].map((k) => `${k}=${t[k]}`)).join(" "); };
  const live = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
  const src = document.createElement("div");
  src.className = `${docCanvasEl.className} dwb-pp-src`;
  src.style.width = `${docCanvasEl.clientWidth}px`;
  document.body.appendChild(src);
  await renderDocxPreview(await buildDocumentForSave(), src);
  const full = collectPreviewParagraphElements(src.querySelector(".docx-preview-host"));
  const out = indices.map((i) => {
    const a = live[i], b = full[i];
    if (!a || !b) return { i, ok: false, why: "brak akapitu" };
    const diffs = [];
    if (styleOf(a) !== styleOf(b)) diffs.push(`styl: ${styleOf(a)} ≠ ${styleOf(b)}`);
    if (classesOf(a) !== classesOf(b)) diffs.push(`klasy: ${classesOf(a)} ≠ ${classesOf(b)}`);
    if (look(a) !== look(b)) diffs.push(`wygląd: ${look(a)} ≠ ${look(b)}`);
    if (a.textContent !== b.textContent) diffs.push("tekst");
    return { i, ok: !diffs.length, why: diffs.join(" | ").slice(0, 400) };
  });
  src.remove();
  return out;
}, idx);

async function main() {
  const browser = await pw[ENGINE].launch({ headless: true });
  // MOBILE=1 — telefon (Widok mobilny: tekst zawinięty do ekranu, wcięcia skalowane)
  const context = await browser.newContext(process.env.MOBILE
    ? { serviceWorkers: "block", viewport: { width: 390, height: 844 }, hasTouch: true, ...(ENGINE === "chromium" ? { isMobile: true } : {}) }
    : { serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "latka.docx", mimeType: DOCX_MIME, buffer: fs.readFileSync(SAMPLE) });
  await idle(page);
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await idle(page);
  await page.evaluate(() => {
    window.__renders = 0;
    const orig = window.renderCurrentDocument;
    window.renderCurrentDocument = function counted(...a) { window.__renders++; return orig.apply(this, a); };
  });

  // otwarty, nietknięty dokument (z polem tekstowym) — zero „zmian”; pisanie + Cofnij bez przerysowania
  check("bez edycji: brak fantomowych zmian akapitów (pole tekstowe w akapicie)", (await page.evaluate(() => collectInlineParagraphEdits().length)) === 0, JSON.stringify(await page.evaluate(() => collectInlineParagraphEdits().map((e) => e.index))));
  await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => !x.dataset.lock && x.textContent.length > 40);
    p.closest(".docx-edit-root")?.focus();
    placeCaret(p, 5);
  });
  await page.keyboard.type("xyz");
  await page.waitForTimeout(400);
  await page.evaluate(() => { window.__renders = 0; });
  await page.click("#undoBtn");
  await idle(page);
  check("pisanie → Cofnij: bez przerysowania dokumentu", (await page.evaluate(() => window.__renders)) === 0 && (await page.evaluate(() => collectInlineParagraphEdits().length)) === 0, String(await page.evaluate(() => window.__renders)));

  // akapity do zmian: zwykłe akapity treści (różne: z komentarzem/przypisem/tabulatorem, jeśli są)
  const targets = await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const ok = ps.map((p, i) => ({ p, i })).filter(({ p }) => !p.dataset.lock && !p.closest("td, th, header, footer, .dwb-tb") && !isListParagraph(p) && p.textContent.length > 30);
    const pick = (fn) => ok.find(fn)?.i;
    const out = new Set([ok[2]?.i, ok[Math.floor(ok.length / 2)]?.i, pick(({ p }) => p.querySelector("[data-cm]")), pick(({ p }) => p.querySelector(".docx-tab")), pick(({ p }) => /justify/.test(getComputedStyle(p).textAlign))]);
    return [...out].filter((i) => i != null);
  });
  check("akapity do sprawdzenia", targets.length >= 3, JSON.stringify(targets));

  const OPS = [
    { name: "do środka", edit: (i) => ({ op: "paraFormat", indices: [i], align: "center" }) },
    { name: "do prawej", edit: (i) => ({ op: "paraFormat", indices: [i], align: "right" }) },
    { name: "wyjustowany", edit: (i) => ({ op: "paraFormat", indices: [i], align: "both" }) },
    { name: "do lewej", edit: (i) => ({ op: "paraFormat", indices: [i], align: "left" }) },
    { name: "wcięcie +", edit: (i) => ({ op: "paraFormat", indices: [i], indentDelta: 1 }) },
    { name: "wcięcie −", edit: (i) => ({ op: "paraFormat", indices: [i], indentDelta: -1 }) },
    { name: "odstęp przed/po 12 pt", edit: (i) => ({ op: "paraFormat", indices: [i], spacing: { before: 240, after: 240 } }) },
    { name: "interlinia 1,5", edit: (i) => ({ op: "paraFormat", indices: [i], spacing: { line: 360, lineRule: "auto" } }) },
    { name: "pierwszy wiersz 1 cm", edit: (i) => ({ op: "paraFormat", indices: [i], ind: { firstLine: 567 } }) },
  ];
  let fast = 0, total = 0;
  for (const i of targets) {
    for (const op of OPS) {
      total++;
      const before = await page.evaluate((idx) => {
        const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[idx];
        p.scrollIntoView({ block: "center" });
        p.closest(".docx-edit-root")?.focus();
        placeCaret(p, 7);
        window.__renders = 0;
        return { off: getCaretOffset(p) };
      }, i);
      await page.evaluate((e) => applyDocumentEdit(e), op.edit(i));
      await idle(page);
      const st = await page.evaluate((idx) => {
        const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
        const sel = getSelection();
        const at = sel?.rangeCount ? ps.findIndex((p) => p.contains(sel.anchorNode)) : -1;
        return { renders: window.__renders, at, off: at >= 0 ? getCaretOffset(ps[at]) : -1 };
      }, i);
      if (st.renders === 0) fast++;
      const cmp = await compareWithFullRender(page, [i]);
      check(`akapit ${i} · ${op.name}: łatka = pełny render`, cmp[0].ok, cmp[0].why);
      if (st.renders === 0) check(`akapit ${i} · ${op.name}: kursor zostaje`, st.at === i && st.off === before.off, JSON.stringify({ st, before }));
    }
  }
  check(`szybka ścieżka użyta (bez przerysowania) w większości operacji`, fast >= total * 0.8, `${fast}/${total}`);

  // kilka akapitów naraz (zaznaczenie przez akapity)
  const span = await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    for (let i = 3; i < ps.length - 3; i++) {
      const trio = [ps[i], ps[i + 1], ps[i + 2]];
      if (trio.every((p) => !p.dataset.lock && !p.closest("td, th, header, footer, .dwb-tb") && !isListParagraph(p) && p.textContent.length > 10)) return [i, i + 1, i + 2];
    }
    return null;
  });
  if (span) {
    await page.evaluate(() => { window.__renders = 0; });
    await page.evaluate((ix) => applyDocumentEdit({ op: "paraFormat", indices: ix, align: "right" }), span);
    await idle(page);
    const cmp = await compareWithFullRender(page, span);
    check("trzy akapity naraz: łatka = pełny render", cmp.every((c) => c.ok), JSON.stringify(cmp.filter((c) => !c.ok)));
    check("…bez przerysowania", (await page.evaluate(() => window.__renders)) === 0);
  }

  // podgląd = plik i Cofnij (przerysowanie z migawki) — po całej serii
  const same = await page.evaluate(async () => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const dom = ps.map((p) => (p.dataset.lock ? null : previewRunsToPlainText(extractRunsFromPreviewParagraph(p)).replace(/[\t\n]/g, "")));
    const xml = (await extractParagraphTextsFromDocx(await buildDocumentForSave())).map((t) => String(t).replace(/[\t\n]/g, ""));
    return dom.length === xml.length && dom.every((t, k) => t === null || t === xml[k]);
  });
  check("podgląd = plik po serii zmian", same);
  const alignBefore = await page.evaluate((ix) => getComputedStyle(collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[ix[0]]).textAlign, span || [targets[0]]);
  await page.click("#undoBtn");
  await idle(page);
  const alignAfter = await page.evaluate((ix) => getComputedStyle(collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[ix[0]]).textAlign, span || [targets[0]]);
  check("Cofnij cofa ostatnią zmianę wyglądu", alignBefore !== alignAfter, `${alignBefore} → ${alignAfter}`);

  // ── łatka z różnicy plików (patchFromDiff): operacje, które tylko przepisują akapity ──────────
  const live = (i) => page.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].textContent, i);
  const viewIsFile = () => page.evaluate(async () => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const dom = ps.map((p) => (p.dataset.lock ? null : previewRunsToPlainText(extractRunsFromPreviewParagraph(p)).replace(/[\t\n]/g, "")));
    const xml = (await extractParagraphTextsFromDocx(await buildDocumentForSave())).map((t) => String(t).replace(/[\t\n]/g, ""));
    return dom.length === xml.length && dom.every((t, k) => t === null || t === xml[k]);
  });
  const t0 = targets[0];
  const renders = () => page.evaluate(() => window.__renders);
  const zero = () => page.evaluate(() => { window.__renders = 0; });
  // Cofnij / Ponów operacji na pliku (dawniej zawsze pełne przerysowanie)
  await zero();
  await page.evaluate(() => document.getElementById("redoBtn").click());
  await idle(page);
  check("Ponów zmiany wyglądu: bez przerysowania", (await renders()) === 0, String(await renders()));
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  check("Cofnij zmiany wyglądu: bez przerysowania", (await renders()) === 0, String(await renders()));
  const cu = await compareWithFullRender(page, span || [t0]);
  check("…Cofnij: łatka = pełny render", cu.every((c) => c.ok), JSON.stringify(cu.filter((c) => !c.ok)));
  // link (nowa relacja w pliku) — tylko ten akapit
  await zero();
  await page.evaluate((i) => applyDocumentEdit({ op: "link", index: i, start: 0, end: 5, href: "https://example.com/" }), t0);
  await idle(page);
  const cl = await compareWithFullRender(page, [t0]);
  check("link: bez przerysowania, łatka = pełny render", (await renders()) === 0 && cl[0].ok, `${await renders()} ${cl[0].why}`);
  check("…link w podglądzie", await page.evaluate((i) => !!collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].querySelector('a[href="https://example.com/"]'), t0));
  await zero();
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  check("…Cofnij link: bez przerysowania, linku nie ma", (await renders()) === 0 && await page.evaluate((i) => !collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].querySelector('a[href="https://example.com/"]'), t0), String(await renders()));
  // koniec Sugerowania: poprawki w dokumencie — tylko zmienione akapity
  await page.evaluate(() => localStorage.setItem("dwb.authorName", "Tester"));
  await page.evaluate(() => dwbTrack.start({ silent: true }));
  await idle(page);
  await page.evaluate((i) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; p.closest(".docx-edit-root")?.focus(); placeCaret(p, 3); }, t0);
  await page.keyboard.type("QQ ");
  await page.waitForTimeout(500);
  await zero();
  await page.evaluate(() => dwbTrack.stop());
  await idle(page);
  const ct = await compareWithFullRender(page, [t0]);
  check("koniec Sugerowania: bez przerysowania, łatka = pełny render", (await renders()) === 0 && ct[0].ok, `${await renders()} ${ct[0].why}`);
  check("…poprawka w podglądzie (wstawka)", await page.evaluate((i) => /QQ/.test(collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].querySelector("[data-dwb-rev]")?.textContent || ""), t0));
  check("…podgląd = plik", await viewIsFile());
  // styl akapitu z listy (styl jest w pliku — plik stylów bez zmian): tylko ten akapit
  for (const [key, name] of [["h1", "Nagłówek 1"], ["h2", "Nagłówek 2"], ["normal", "Normalny"]]) {
    await zero();
    await page.evaluate(([i, key]) => applyDocumentEdit({ op: "paraFormat", indices: [i], style: key, resetDirect: true }), [t0, key]);
    await idle(page);
    const cs = await compareWithFullRender(page, [t0]);
    const r = await renders();
    check(`styl „${name}”: ${r === 0 ? "bez przerysowania, " : ""}łatka = pełny render`, cs[0].ok, `${r} ${cs[0].why}`);
    if (key === "h1") {
      check("…„Nagłówek 1” bez przerysowania (styl w pliku)", r === 0, String(r));
      check("…nagłówek w panelu Struktura", await page.evaluate((i) => { const txt = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].textContent.trim().slice(0, 20); return (documentStructure?.headings || []).some((h) => String(h.label || "").includes(txt)); }, t0));
    }
  }
  // bezpiecznik: nowy akapit (podział strony) — dalej pełne przerysowanie i podgląd = plik
  await zero();
  await page.evaluate((i) => applyDocumentEdit({ op: "pageBreak", index: i }), t0);
  await idle(page);
  check("podział strony (nowy akapit): pełne przerysowanie, podgląd = plik", (await renders()) === 1 && await viewIsFile(), String(await renders()));

  check("bez błędów w konsoli", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail || ""}`}`);
  console.log(`\n[${ENGINE}${process.env.MOBILE ? ", telefon" : ""}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}

main().catch((e) => {
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail || ""}`}`);
  console.error(e);
  process.exit(1);
});
