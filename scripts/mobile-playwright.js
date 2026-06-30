// mobile-playwright.js — mobile viewport: reflow layout, panel sheet, no horizontal overflow.

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
    const canvas = document.getElementById("docCanvas");
    const section = document.querySelector(".docx-preview-host section.docx") || document.querySelector(".docx-preview-host .docx");
    const host = document.querySelector(".docx-preview-host");
    return {
      hasDocument: document.body.classList.contains("has-document"),
      reflow: canvas?.classList.contains("doc-reflow-mode"),
      viewportScrollW: vp?.scrollWidth || 0,
      viewportClientW: vp?.clientWidth || 0,
      hostW: host?.getBoundingClientRect().width || 0,
      sectionW: section?.getBoundingClientRect().width || 0,
      hasHorizontalOverflow: (vp?.scrollWidth || 0) > (vp?.clientWidth || 0) + 2,
    };
  });

  if (!layout.hasDocument) throw new Error("Brak klasy has-document po wczytaniu");
  console.log("  ✓ Tryb pełnej szerokości (has-document)");

  const langVisible = await page.evaluate(() => {
    const el = document.getElementById("langSwitch");
    return !!el && getComputedStyle(el).display !== "none";
  });
  if (!langVisible) throw new Error("Przełącznik PL/EN ukryty w trybie czytania");
  console.log("  ✓ PL/EN dostępne podczas czytania");

  if (layout.hasHorizontalOverflow) {
    throw new Error(`Poziomy overflow: scroll=${layout.viewportScrollW} client=${layout.viewportClientW}`);
  }
  const widthUse = layout.hostW / layout.viewportClientW;
  if (widthUse < 0.96) {
    throw new Error(`Za wąski podgląd: host ${Math.round(layout.hostW)}px / viewport ${layout.viewportClientW}px (${Math.round(widthUse * 100)}%)`);
  }
  console.log(`  ✓ Pełna szerokość (${Math.round(layout.hostW)}px / ${layout.viewportClientW}px, reflow=${layout.reflow})`);

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

  const sliderHidden = await page.evaluate(() => document.getElementById("zoomSliderField")?.classList.contains("hidden"));
  if (!sliderHidden) throw new Error("Suwak zoom powinien być ukryty w trybie reflow");
  console.log("  ✓ Suwak zoom ukryty w reflow");

  assertNoErrors(errors, "mobile-playwright");
  await browser.close();
  console.log("  ✓ Brak błędów JS w konsoli\n");
}

run().catch((err) => {
  console.error("❌  mobile-playwright failed:", err.message || err);
  process.exit(1);
});
