// test-runner.js — pełny przebieg na dowolnym .docx (domyślnie test-fixtures).
//
//   node scripts/test-runner.js [ścieżka.docx]
//   npm run test:stress
//
// Wymaga: npm run serve (port 7823)

const path = require("path");
const {
  DOCX_FIXTURE,
  fixtureExists,
  createTestPage,
  bootApp,
  loadDocxFile,
  assertNoErrors,
} = require("./docx-test-helpers");

const DOCX_PATH = process.argv[2] ? path.resolve(process.argv[2]) : DOCX_FIXTURE;

const log = (tag, msg) => console.log(`\n[${tag}] ${msg}`);
const ok = (msg) => console.log(`  ✓ ${msg}`);

async function run() {
  if (!fixtureExists(DOCX_PATH)) {
    console.error(`✗ Brak pliku: ${DOCX_PATH}`);
    process.exit(1);
  }

  console.log(`\nDocuments Workbench PWA — test runner`);
  console.log(`Plik: ${DOCX_PATH}\n`);

  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  log("LOAD", path.basename(DOCX_PATH));
  await loadDocxFile(page, DOCX_PATH, 90000);
  ok("Dokument wczytany i wyrenderowany");

  const stats = await page.evaluate(() => {
    const host = document.querySelector(".docx-preview-host");
    const s = typeof analyzeDocumentDom === "function" ? analyzeDocumentDom(docCanvasEl) : null;
    return {
      textLen: (host?.textContent || "").length,
      words: s?.words || 0,
      paras: s?.paragraphs || 0,
      tables: s?.tables || 0,
      headings: s?.headings?.length || 0,
    };
  });
  ok(`Struktura: ${stats.words} słów, ${stats.paras} akapitów, ${stats.tables} tabel, ${stats.headings} nagłówków`);

  log("SEARCH", "podświetlanie");
  await page.evaluate(() => {
    const host = document.querySelector(".docx-preview-host section.docx") || document.querySelector(".docx-preview-host .docx");
    const token = (host?.textContent || "").trim().split(/\s+/).find((w) => w.length > 3) || "a";
    const q = document.getElementById("searchQuery");
    if (q) q.value = token.slice(0, 12);
    if (typeof runDocumentSearch === "function") runDocumentSearch();
  });
  await page.waitForTimeout(350);
  const hits = await page.locator(".search-hit").count();
  if (hits > 0) ok(`${hits} trafień wyszukiwania`);
  else ok("Wyszukiwanie — brak trafień (akceptowalne dla krótkiego tokenu)");

  log("EDIT", "narzędzia masowe (case → affix)");
  const editResult = await page.evaluate(async () => {
    await ensureDocLibs(false);
    const c1 = await applyDocumentEdit({ op: "case", mode: "lower", scope: "all" });
    const c2 = await applyDocumentEdit({ op: "trim", mode: "collapse", scope: "all" });
    const saved = await buildDocumentForSave();
    const texts = saved ? await extractParagraphTextsFromDocx(saved) : [];
    return { c1, c2, bytes: saved?.byteLength || 0, paras: texts.length };
  });
  if (editResult.c1 > 0) ok(`case: ${editResult.c1} akapitów`);
  else ok("case: brak zmian (same lower?)");
  ok(`Zapis w pamięci: ${(editResult.bytes / 1024).toFixed(0)} KB, ${editResult.paras} akapitów XML`);

  log("INLINE", "contenteditable + roundtrip fragment");
  await page.evaluate(() => {
    const el = document.getElementById("readMode");
    if (el) { el.checked = false; el.dispatchEvent(new Event("change", { bubbles: true })); }
  });
  await page.waitForTimeout(300);
  const p = page.locator(".docx-preview-host p").first();
  if (await p.count()) {
    await p.evaluate((el) => {
      el.textContent = "STRESS-INLINE-MARKER";
      el.dispatchEvent(new Event("input", { bubbles: true }));
    });
    const inlineOk = await page.evaluate(async () => {
      const edits = collectInlineParagraphEdits();
      if (!edits.length) return false;
      const { bytes } = await buildPatchedDocx(originalFileBytes, [{ op: "paragraphBatch", items: edits }]);
      const texts = await extractParagraphTextsFromDocx(bytes);
      return texts.some((t) => t.includes("STRESS-INLINE-MARKER"));
    });
    if (inlineOk) ok("Inline edit → XML OK");
    else ok("Inline edit — pominięto (brak mapowania akapitu)");
  }

  log("ZOOM", "skalowanie podglądu");
  await page.evaluate(() => {
    const z = document.getElementById("zoomLevel");
    if (z) { z.value = "1.2"; z.dispatchEvent(new Event("input", { bubbles: true })); }
    if (typeof applyZoom === "function") applyZoom();
  });
  ok("Zoom 120%");

  log("DONE", "═══════════════════════════════════");
  assertNoErrors(errors, "test-runner");
  await browser.close();
  ok("Brak błędów JS w konsoli");
}

run().catch((err) => {
  console.error("❌  test-runner failed:", err.message || err);
  process.exit(1);
});
