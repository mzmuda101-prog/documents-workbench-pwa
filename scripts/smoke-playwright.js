const { chromium } = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:7823/";

async function run() {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  const title = await page.title();
  if (!title.includes("Documents Workbench")) throw new Error(`Unexpected title: ${title}`);

  await page.click("#loadSampleBtn");
  await page.waitForTimeout(2500);

  const hasDocText = await page.locator(".docx-preview-host").evaluate((el) => (el?.textContent || "").includes("Documents Workbench"));
  if (!hasDocText) throw new Error("Sample document did not render");

  const words = await page.locator(".structure-stat strong").first().textContent();
  if (!words || Number(words) < 1) throw new Error("Structure panel empty");

  if (errors.length) throw new Error(`Console errors:\n${errors.join("\n")}`);
  await browser.close();
  console.log("✅  smoke-playwright passed");
}

run().catch((err) => {
  console.error("❌  smoke-playwright failed:", err.message || err);
  process.exit(1);
});
