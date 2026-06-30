// load-fixture-playwright.js — wczytanie realnego .docx z test-fixtures (render + struktura + search).

const {
  DOCX_FIXTURE,
  fixtureExists,
  createTestPage,
  bootApp,
  loadDocxFile,
  assertNoErrors,
} = require("./docx-test-helpers");

async function run() {
  if (!fixtureExists(DOCX_FIXTURE)) {
    console.warn(`⚠️  Pominięto — brak fixture: ${DOCX_FIXTURE}`);
    process.exit(0);
  }

  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await loadDocxFile(page, DOCX_FIXTURE, 90000);

  const stats = await page.evaluate(() => {
    const section = document.querySelector(".docx-preview-host section.docx") || document.querySelector(".docx-preview-host .docx");
    const text = section?.textContent || "";
    const words = document.querySelector(".structure-stat strong")?.textContent;
    const paras = section?.querySelectorAll("p")?.length || 0;
    return { textLen: text.length, words: Number(words) || 0, paras };
  });

  if (stats.textLen < 50) throw new Error(`Podgląd za krótki (${stats.textLen} znaków)`);
  if (stats.words < 1) throw new Error("Panel struktury — brak słów");
  if (stats.paras < 1) throw new Error("Brak akapitów w podglądzie");

  await page.evaluate(() => {
    const host = document.querySelector(".docx-preview-host section.docx") || document.querySelector(".docx-preview-host .docx");
    const words = (host?.textContent || "").match(/[\p{L}\p{N}]{4,}/gu) || [];
    for (const w of words.slice(0, 30)) {
      const q = document.getElementById("searchQuery");
      if (q) q.value = w;
      if (typeof runDocumentSearch === "function") runDocumentSearch();
      if (document.querySelectorAll(".search-hit").length) return;
    }
  });
  await page.waitForTimeout(300);

  const hitCount = await page.locator(".search-hit").count();
  if (hitCount < 1) throw new Error("Wyszukiwanie — brak trafień (sprawdź tokeny z treści)");

  assertNoErrors(errors, "load-fixture");
  await browser.close();
  console.log(`✅  load-fixture-playwright passed (${stats.words} słów, ${stats.paras} akapitów, ${hitCount} trafień)`);
}

run().catch((err) => {
  console.error("❌  load-fixture-playwright failed:", err.message || err);
  process.exit(1);
});
