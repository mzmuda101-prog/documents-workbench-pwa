// metadata-playwright.js — read/write docProps/core.xml (title, author, keywords).

const {
  createTestPage,
  bootApp,
  loadBuiltinSample,
  assertNoErrors,
} = require("./docx-test-helpers");

async function run() {
  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await loadBuiltinSample(page);

  const result = await page.evaluate(async () => {
    await ensureDocLibs(false);

    const marker = `META-${Date.now()}`;
    const fields = {
      title: `Tytuł ${marker}`,
      creator: `Autor ${marker}`,
      keywords: `klucz, ${marker}`,
    };

    const count = await applyDocumentEdit({ op: "coreMetadata", fields });
    if (!count) return { ok: false, step: "apply", msg: "count=0" };

    const after = await extractCoreMetadataFromDocx(originalFileBytes);
    if (after.title !== fields.title) return { ok: false, step: "title", msg: JSON.stringify(after) };
    if (after.creator !== fields.creator) return { ok: false, step: "creator", msg: JSON.stringify(after) };
    if (after.keywords !== fields.keywords) return { ok: false, step: "keywords", msg: JSON.stringify(after) };

    const zip = await window.JSZip.loadAsync(originalFileBytes);
    if (!zip.file("docProps/core.xml")) return { ok: false, step: "core-file", msg: "missing core.xml" };
    const rels = await zip.file("_rels/.rels")?.async("string");
    if (!rels?.includes("core-properties")) return { ok: false, step: "rels", msg: rels };
    const ct = await zip.file("[Content_Types].xml")?.async("string");
    if (!ct?.includes("/docProps/core.xml")) return { ok: false, step: "ct", msg: ct };

    return { ok: true, marker };
  });

  if (!result.ok) throw new Error(`metadata failed at ${result.step}: ${result.msg}`);
  assertNoErrors(errors, "metadata");
  await browser.close();
  console.log(`✅  metadata-playwright passed (${result.marker})`);
}

run().catch((err) => {
  console.error("❌  metadata-playwright failed:", err.message || err);
  process.exit(1);
});
