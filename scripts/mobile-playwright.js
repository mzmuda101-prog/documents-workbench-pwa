// mobile-playwright.js — mobile viewport: fit zoom, panel sheet, no horizontal overflow.

const {
  createTestPage,
  bootApp,
  loadBuiltinSample,
  assertNoErrors,
} = require("./docx-test-helpers");

const MOBILE_VIEWPORT = { width: 390, height: 844 };

async function run() {
  console.log("\nDocuments Workbench PWA — mobile UX test (390×844)\n");

  const { browser, page, errors } = await createTestPage(MOBILE_VIEWPORT);
  await bootApp(page);

  const isMobile = await page.evaluate(() => document.documentElement.classList.contains("is-mobile"));
  if (!isMobile) throw new Error("Brak klasy is-mobile na viewporcie mobilnym");

  await loadBuiltinSample(page);

  const layout = await page.evaluate(() => {
    const vp = document.getElementById("docViewport");
    const host = document.querySelector(".docx-preview-host");
    const zoom = parseFloat(getComputedStyle(document.getElementById("docCanvas")).getPropertyValue("--doc-zoom")) || 1;
    return {
      viewportScrollW: vp?.scrollWidth || 0,
      viewportClientW: vp?.clientWidth || 0,
      hostW: host?.getBoundingClientRect().width || 0,
      zoom,
      hasHorizontalOverflow: (vp?.scrollWidth || 0) > (vp?.clientWidth || 0) + 2,
    };
  });

  if (layout.hasHorizontalOverflow) {
    throw new Error(`Poziomy overflow: scroll=${layout.viewportScrollW} client=${layout.viewportClientW}`);
  }
  console.log(`  ✓ Brak poziomego overflow (zoom ${Math.round(layout.zoom * 100)}%)`);

  await page.locator("#panelToggle").click();
  await page.waitForTimeout(250);
  const panelOpen = await page.evaluate(() => document.documentElement.classList.contains("sidebar-open"));
  if (!panelOpen) throw new Error("Panel nie otworzył się na mobile");
  console.log("  ✓ Bottom sheet panel otwarty");

  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
  await page.waitForTimeout(200);
  const panelClosed = await page.evaluate(() => !document.documentElement.classList.contains("sidebar-open"));
  if (!panelClosed) throw new Error("Panel nie zamknął się");
  console.log("  ✓ Panel zamknięty");

  await page.evaluate(() => {
    if (typeof applyFitToWidth === "function") applyFitToWidth();
  });
  const fitActive = await page.evaluate(() => document.getElementById("zoomFitBtn")?.classList.contains("primary"));
  if (!fitActive) throw new Error("Przycisk „Dopasuj szerokość” nie jest aktywny");
  console.log("  ✓ Fit-to-width aktywny");

  assertNoErrors(errors, "mobile-playwright");
  await browser.close();
  console.log("  ✓ Brak błędów JS w konsoli\n");
}

run().catch((err) => {
  console.error("❌  mobile-playwright failed:", err.message || err);
  process.exit(1);
});
