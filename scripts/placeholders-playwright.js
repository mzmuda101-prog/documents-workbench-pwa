// placeholders-playwright.js — scan {{field}}, fill form, verify XML.

const {
  createTestPage,
  bootApp,
  loadBuiltinSample,
  assertNoErrors,
} = require("./docx-test-helpers");

const TEMPLATE = "Umowa: {{imie}} {{nazwisko}}, data {{data}}.";

async function run() {
  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await loadBuiltinSample(page);

  const result = await page.evaluate(async (template) => {
    await ensureDocLibs(false);

    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    const idx = texts.findIndex((t) => t && t.trim().length > 0);
    if (idx < 0) return { ok: false, step: "fixture", msg: "Brak akapitu" };

    await applyDocumentEdit({
      op: "paragraphBatch",
      items: [{ index: idx, text: template }],
    });

    const scan = await scanPlaceholders(originalFileBytes);
    if (scan.fields.length !== 3) {
      return { ok: false, step: "scan", msg: JSON.stringify(scan) };
    }
    if (scan.total !== 3) return { ok: false, step: "count", msg: String(scan.total) };

    const count = await applyDocumentEdit({
      op: "placeholderFill",
      values: { imie: "Jan", nazwisko: "Kowalski", data: "2026-06-30" },
      scope: "all",
    });
    if (count !== 3) return { ok: false, step: "fill", msg: `count=${count}` };

    const after = await extractParagraphTextsFromDocx(originalFileBytes);
    const joined = after.join("\n");
    if (!joined.includes("Jan Kowalski")) return { ok: false, step: "text", msg: joined };
    if (joined.includes("{{")) return { ok: false, step: "leftover", msg: joined };

    return { ok: true, count };
  }, TEMPLATE);

  if (!result.ok) throw new Error(`placeholders failed at ${result.step}: ${result.msg}`);
  assertNoErrors(errors, "placeholders");
  await browser.close();
  console.log(`✅  placeholders-playwright passed (${result.count} replacements)`);
}

run().catch((err) => {
  console.error("❌  placeholders-playwright failed:", err.message || err);
  process.exit(1);
});
