// save-roundtrip-playwright.js — inline edit + bulk edit + save roundtrip (ZIP-patch).

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

  // Tryb edycji — inline WYSIWYG (panel może być zwinięty na wąskim viewport)
  await page.evaluate(() => {
    const el = document.getElementById("readMode");
    if (!el) return;
    el.checked = false;
    el.dispatchEvent(new Event("change", { bubbles: true }));
  });
  await page.waitForTimeout(400);

  const firstP = page.locator(".docx-preview-host p").first();
  await firstP.evaluate((el) => {
    el.textContent = "ROUNDTRIP INLINE EDIT";
    el.dispatchEvent(new Event("input", { bubbles: true }));
  });

  const result = await page.evaluate(async () => {
    await ensureDocLibs(false);
    const inline = collectInlineParagraphEdits();
    if (!inline.length) return { ok: false, step: "inline", msg: "No inline edits detected" };

    const saved = await buildDocumentForSave();
    if (!saved) return { ok: false, step: "save", msg: "buildDocumentForSave returned null" };

    const texts = await extractParagraphTextsFromDocx(saved);
    if (!texts[0]?.includes("ROUNDTRIP INLINE EDIT")) {
      return { ok: false, step: "inline-xml", msg: `Paragraph 0: ${texts[0]}` };
    }

    // Bulk: znajdź i zamień
    originalFileBytes = saved;
    pendingDocEdits = [];
    await refreshInlineEditBaseline(saved);
    const replaceCount = await applyDocumentEdit({
      op: "replace",
      find: "{{placeholder}}",
      replace: "WYPELNIONE",
      regex: false,
      scope: "all",
    });
    if (!replaceCount) return { ok: false, step: "replace", msg: "Replace made no changes" };

    const afterReplace = await buildDocumentForSave();
    const texts2 = await extractParagraphTextsFromDocx(afterReplace);
    const joined = texts2.join(" ");
    if (!joined.includes("WYPELNIONE")) {
      return { ok: false, step: "replace-xml", msg: joined };
    }

    // Roundtrip: wczytaj ponownie z bajtów
    originalFileBytes = afterReplace;
    pendingDocEdits = [];
    await reloadFromBytes(afterReplace);
    await new Promise((r) => setTimeout(r, 400));
    const hostText = document.querySelector(".docx-preview-host")?.textContent || "";
    if (!hostText.includes("ROUNDTRIP INLINE EDIT") || !hostText.includes("WYPELNIONE")) {
      return { ok: false, step: "rerender", msg: hostText.slice(0, 200) };
    }

    return { ok: true, replaceCount, paraCount: texts2.length };
  });

  if (!result.ok) throw new Error(`Roundtrip failed at ${result.step}: ${result.msg}`);
  assertNoErrors(errors, "save-roundtrip");
  await browser.close();
  console.log(`✅  save-roundtrip-playwright passed (${result.paraCount} paragraphs, ${result.replaceCount} replace hits)`);
}

run().catch((err) => {
  console.error("❌  save-roundtrip-playwright failed:", err.message || err);
  process.exit(1);
});
