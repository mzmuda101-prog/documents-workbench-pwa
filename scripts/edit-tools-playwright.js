// edit-tools-playwright.js — narzędzia edycji na realnym .docx (case + trim + affix + replace pipeline).

const path = require("path");
const {
  DOCX_FIXTURE,
  fixtureExists,
  createTestPage,
  bootApp,
  loadDocxFile,
  assertNoErrors,
} = require("./docx-test-helpers");

const DOCX_ARG = process.argv[2];

async function run() {
  const filePath = DOCX_ARG ? path.resolve(DOCX_ARG) : DOCX_FIXTURE;
  if (!fixtureExists(filePath)) {
    console.warn(`⚠️  Pominięto — brak pliku: ${filePath}`);
    process.exit(0);
  }

  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await loadDocxFile(page, filePath, 90000);

  const result = await page.evaluate(async () => {
    await ensureDocLibs(false);
    const before = await extractParagraphTextsFromDocx(originalFileBytes);
    const hasLower = before.some((t) => /[a-ząćęłńóśźż]/.test(t));
    if (!hasLower) return { ok: false, step: "precheck", msg: "Brak małych liter do testu case" };

    const caseCount = await applyDocumentEdit({ op: "case", mode: "upper", scope: "all" });
    if (!caseCount) return { ok: false, step: "case", msg: "case upper — 0 zmian" };

    const afterCase = await extractParagraphTextsFromDocx(originalFileBytes);
    const stillLower = afterCase.some((t) => /[a-ząćęłńóśźż]/.test(t));
    if (stillLower) return { ok: false, step: "case-verify", msg: "Pozostały małe litery" };

    const trimCount = await applyDocumentEdit({ op: "trim", mode: "ends", scope: "all" });
    // trim może dać 0 — to OK

    const affixCount = await applyDocumentEdit({
      op: "affix",
      prefix: "[[T]]",
      suffix: "[[/T]]",
      scope: "body",
    });
    if (!affixCount) return { ok: false, step: "affix", msg: "affix — 0 zmian" };

    const saved = await buildDocumentForSave();
    if (!saved || saved.byteLength < 1000) return { ok: false, step: "save", msg: "buildDocumentForSave" };

    const texts = await extractParagraphTextsFromDocx(saved);
    const joined = texts.join(" ");
    if (!joined.includes("[[T]]") || !joined.includes("[[/T]]")) {
      return { ok: false, step: "affix-xml", msg: joined.slice(0, 120) };
    }

    return { ok: true, caseCount, trimCount, affixCount, paras: texts.length };
  });

  if (!result.ok) throw new Error(`edit-tools failed at ${result.step}: ${result.msg}`);
  assertNoErrors(errors, "edit-tools");
  await browser.close();
  console.log(
    `✅  edit-tools-playwright passed (${result.paras} akapitów, case=${result.caseCount}, affix=${result.affixCount})`
  );
}

run().catch((err) => {
  console.error("❌  edit-tools-playwright failed:", err.message || err);
  process.exit(1);
});
