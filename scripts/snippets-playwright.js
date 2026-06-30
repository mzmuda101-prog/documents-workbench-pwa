// snippets-playwright.js — !trigger expand + {{placeholder}} cooperation.

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

    upsertSnippet("klauzula", "Umowa z {{imie}} {{nazwisko}}.");
    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    const idx = texts.findIndex((t) => t && t.trim().length > 0);
    if (idx < 0) return { ok: false, step: "fixture", msg: "Brak akapitu" };

    await applyDocumentEdit({
      op: "paragraphBatch",
      items: [{ index: idx, text: "Wstęp !klauzula koniec." }],
    });

    const scanSn = await scanSnippetTriggers(originalFileBytes);
    if (!scanSn.triggers.some((t) => t.name === "klauzula")) {
      return { ok: false, step: "scan-sn", msg: JSON.stringify(scanSn) };
    }

    const expanded = await applyDocumentEdit({
      op: "snippetExpand",
      snippets: { klauzula: "Umowa z {{imie}} {{nazwisko}}." },
      scope: "all",
    });
    if (!expanded) return { ok: false, step: "expand", msg: "0" };

    const afterExpand = await extractParagraphTextsFromDocx(originalFileBytes);
    if (!afterExpand.join(" ").includes("{{imie}}")) {
      return { ok: false, step: "ph-visible", msg: afterExpand.join(" | ") };
    }
    if (afterExpand.join(" ").includes("!klauzula")) {
      return { ok: false, step: "trigger-left", msg: afterExpand.join(" | ") };
    }

    const filled = await applyDocumentEdit({
      op: "placeholderFill",
      values: { imie: "Anna", nazwisko: "Nowak" },
      scope: "all",
    });
    if (filled < 2) return { ok: false, step: "fill", msg: String(filled) };

    const finalText = (await extractParagraphTextsFromDocx(originalFileBytes)).join(" ");
    if (!finalText.includes("Anna Nowak")) return { ok: false, step: "final", msg: finalText };

    const dynamic = resolveSnippetBody("Dnia {date} o {time}.");
    if (!dynamic.includes("Dnia ") || dynamic.includes("{date}")) {
      return { ok: false, step: "dynamic", msg: dynamic };
    }

    setSnippetExpandMode("auto");
    if (getSnippetExpandMode() !== "auto") return { ok: false, step: "settings", msg: getSnippetExpandMode() };
    setSnippetExpandMode("manual");

    return { ok: true, expanded, filled };
  });

  if (!result.ok) throw new Error(`snippets failed at ${result.step}: ${result.msg}`);
  assertNoErrors(errors, "snippets");
  await browser.close();
  console.log(`✅  snippets-playwright passed (expand=${result.expanded}, fill=${result.filled})`);
}

run().catch((err) => {
  console.error("❌  snippets-playwright failed:", err.message || err);
  process.exit(1);
});
