// list-reflow-playwright.js — świeżo zmieniony poziom listy musi od razu przesunąć się
// w automatycznym widoku mobilnym; wcześniej styl inline reflow zatrzymywał stare wcięcie.

const { chromium } = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    viewport: { width: 390, height: 844 },
    isMobile: true,
    hasTouch: true,
    serviceWorkers: "block",
  });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await page.waitForTimeout(500);

  await page.keyboard.type("Nadrzędny punkt");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Podpunkt");
  await page.click("#fmtListBtn");
  await page.click('.compose-pop-list .compose-item-label:text-is("Lista punktowana")');
  await page.waitForTimeout(500);

  const itemLeft = () => page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1];
    const range = document.createRange();
    range.selectNodeContents(p);
    return range.getClientRects()[0]?.left || 0;
  });
  const before = await itemLeft();
  await page.keyboard.press("Tab");
  const after = await itemLeft();
  const state = await page.evaluate(() => ({
    layout: getViewLayout(),
    level: getListLevelFromDom(collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1]),
  }));

  await browser.close();
  if (errors.length) throw new Error(errors.join("\n"));
  if (state.layout !== "mobile" || state.level !== 1 || after <= before + 20) {
    throw new Error(`Brak natychmiastowego wcięcia podpunktu: ${JSON.stringify({ before, after, state })}`);
  }
  console.log("✅ list-reflow-playwright passed");
}

run().catch((error) => {
  console.error("❌ list-reflow-playwright failed:", error.message || error);
  process.exit(1);
});
