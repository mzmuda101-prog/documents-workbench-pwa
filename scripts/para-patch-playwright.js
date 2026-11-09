// para-patch-playwright.js — wyrównanie / wcięcie / odstępy BEZ przerysowania dokumentu
// (app/document.js: patchParagraphsInPlace). Niezmiennik: ŁATKA = PEŁNY RENDER — po każdej
// zmianie zapisany plik jest renderowany w tle tak jak przy pełnym przerysowaniu, a zmienione
// akapity porównane z żywymi (styl, klasy z renderu, wyrównanie, marginesy, wcięcie, interlinia,
// wysokość). Do tego: bez przerysowania, kursor zostaje, podgląd = plik, Cofnij przywraca.
//
//   node scripts/para-patch-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/para-patch-playwright.js
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
  const look = (p) => { const cs = getComputedStyle(p); return ["textAlign", "marginLeft", "marginRight", "marginTop", "marginBottom", "textIndent", "lineHeight", "paddingLeft"].map((k) => `${k}=${cs[k]}`).join(" "); };
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
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
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

  check("bez błędów w konsoli", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail || ""}`}`);
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}

main().catch((e) => {
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail || ""}`}`);
  console.error(e);
  process.exit(1);
});
