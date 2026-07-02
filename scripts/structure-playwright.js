// structure-playwright.js — outline jump, highlight, quick paragraph edit.

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

  await page.evaluate(() => {
    if (typeof setSidebarOpen === "function") setSidebarOpen(true);
  });
  await page.locator("#panel-structure summary").click();
  await page.waitForSelector(".structure-outline-item", { timeout: 10000 });

  const listMetrics = await page.evaluate(() => {
    const item = document.querySelector(".structure-outline-item");
    const list = document.querySelector(".structure-outline-list");
    const rect = item?.getBoundingClientRect();
    return {
      itemCount: document.querySelectorAll(".structure-outline-item").length,
      itemHeight: rect?.height || 0,
      pageSize: Number(structurePageSize) || 0,
    };
  });
  if (listMetrics.itemCount > 40) throw new Error(`Za dużo elementów na stronie: ${listMetrics.itemCount}`);
  if (listMetrics.itemHeight > 0 && listMetrics.itemHeight < 36) {
    throw new Error(`Elementy zbyt spłaszczone: ${listMetrics.itemHeight}px`);
  }

  const searchResult = await page.evaluate(() => {
    const probe = documentStructure?.outline?.find((i) => (i.label || "").length > 4)?.label?.split(/\s+/)[0] || "Documents";
    structureSearchQuery = probe;
    structurePage = 0;
    renderStructureOutline(documentStructure);
    const items = document.querySelectorAll(".structure-outline-item").length;
    const info = document.getElementById("structurePageInfo")?.textContent || "";
    return { items, info, probe };
  });
  if (!searchResult.items) throw new Error(`search failed for "${searchResult.probe}": ${JSON.stringify(searchResult)}`);

  const jumpResult = await page.evaluate(() => {
    const items = documentStructure?.outline || [];
    const target = items.find((i) => i.type === "paragraph" || i.type === "heading");
    if (!target) return { ok: false, step: "outline", msg: "Brak elementów outline" };
    jumpToStructureItem(target);
    const active = document.querySelector(".search-hit-active");
    if (!active) return { ok: false, step: "highlight", msg: "Brak podświetlenia" };
    if (structureSelectionId !== target.id) return { ok: false, step: "selection", msg: structureSelectionId };
    return { ok: true, id: target.id, type: target.type };
  });
  if (!jumpResult.ok) throw new Error(`structure jump failed at ${jumpResult.step}: ${jumpResult.msg}`);

  const editResult = await page.evaluate(async () => {
    const items = documentStructure?.outline || [];
    const target = items.find((i) => Number.isFinite(i.paraIndex) && i.type !== "table");
    if (!target) return { ok: false, step: "target", msg: "Brak akapitu" };

    structureSelectionId = target.id;
    syncStructureQuickEdit(target);
    const marker = `STRUCT-EDIT-${Date.now()}`;
    if (structureQuickEditTextEl) structureQuickEditTextEl.value = marker;

    const count = await applyDocumentEdit({
      op: "paragraphBatch",
      items: [{ index: target.paraIndex, text: marker }],
    });
    if (!count) return { ok: false, step: "apply", msg: "count=0" };

    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    if (!texts[target.paraIndex]?.includes(marker)) {
      return { ok: false, step: "xml", msg: texts[target.paraIndex] };
    }
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    const paras = host?.querySelectorAll("p") || [];
    const domText = paras[target.paraIndex]?.textContent || "";
    if (!domText.includes(marker)) return { ok: false, step: "dom", msg: domText };
    return { ok: true, marker };
  });
  if (!editResult.ok) throw new Error(`structure edit failed at ${editResult.step}: ${editResult.msg}`);

  assertNoErrors(errors, "structure");
  await browser.close();
  console.log(`✅  structure-playwright passed (${jumpResult.type} → ${editResult.marker})`);
}

run().catch((err) => {
  console.error("❌  structure-playwright failed:", err.message || err);
  process.exit(1);
});
