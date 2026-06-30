// grammar-playwright.js — skan typografii, apply all, weryfikacja XML.

const path = require("path");
const {
  BUILTIN_SAMPLE,
  fixtureExists,
  createTestPage,
  bootApp,
  loadDocxFile,
  loadBuiltinSample,
  assertNoErrors,
} = require("./docx-test-helpers");

const DOCX_ARG = process.argv[2];
const BAD_TEXT = "Test słowo  słowo , tekst i.";

async function run() {
  const filePath = DOCX_ARG ? path.resolve(DOCX_ARG) : BUILTIN_SAMPLE;
  if (!fixtureExists(filePath)) {
    console.warn(`⚠️  Pominięto — brak pliku: ${filePath}`);
    process.exit(0);
  }

  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  if (DOCX_ARG) await loadDocxFile(page, filePath, 90000);
  else await loadBuiltinSample(page);

  const result = await page.evaluate(async (badText) => {
    await ensureDocLibs(false);

    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    const targetIndex = texts.findIndex((t) => t && t.trim().length > 0);
    if (targetIndex < 0) return { ok: false, step: "fixture", msg: "Brak akapitu w dokumencie" };

    await applyDocumentEdit({
      op: "paragraphBatch",
      items: [{ index: targetIndex, text: badText }],
    });

    const scan = await scanDocument(originalFileBytes, { lang: "pl", nbspPl: false });
    if (scan.hits.length < 2) {
      return { ok: false, step: "scan", msg: `Oczekiwano ≥2 trafień, jest ${scan.hits.length}` };
    }

    const ruleIds = Object.keys(scan.byRule);
    if (!ruleIds.includes("double-space") || !ruleIds.includes("space-before-punct")) {
      return { ok: false, step: "rules", msg: `Brak oczekiwanych reguł: ${ruleIds.join(", ")}` };
    }

    const inline = scanParagraph(badText, "pl", { nbspPl: false });
    if (inline.length < 2) {
      return { ok: false, step: "scanParagraph", msg: "scanParagraph — za mało trafień" };
    }

    const items = buildGrammarBatchItems(
      await extractParagraphTextsFromDocx(originalFileBytes),
      scan.hits,
      { lang: "pl", nbspPl: false }
    );
    if (!items.length) return { ok: false, step: "batch", msg: "Brak pozycji batch" };

    const count = await applyDocumentEdit({ op: "paragraphBatch", items });
    if (!count) return { ok: false, step: "apply", msg: "apply all — 0 zmian" };

    const after = await extractParagraphTextsFromDocx(originalFileBytes);
    const fixed = after[targetIndex] || "";
    if (/ {2,}/.test(fixed)) {
      return { ok: false, step: "verify-space", msg: `Podwójne spacje pozostały: ${fixed}` };
    }
    if (/\s+[,;:.!?]/.test(fixed)) {
      return { ok: false, step: "verify-punct", msg: `Spacja przed interpunkcją: ${fixed}` };
    }

    const rescan = await scanDocument(originalFileBytes, { lang: "pl", nbspPl: false });
    const orphanLeft = rescan.byRule["orphan-i"]?.length || 0;

    return {
      ok: true,
      hits: scan.hits.length,
      rules: ruleIds.length,
      fixed,
      orphanLeft,
    };
  }, BAD_TEXT);

  if (!result.ok) throw new Error(`grammar failed at ${result.step}: ${result.msg}`);
  assertNoErrors(errors, "grammar");
  await browser.close();
  console.log(
    `✅  grammar-playwright passed (hits=${result.hits}, rules=${result.rules}, orphanLeft=${result.orphanLeft})`
  );
  console.log(`    fixed: ${result.fixed}`);
}

run().catch((err) => {
  console.error("❌  grammar-playwright failed:", err.message || err);
  process.exit(1);
});
