// Shared Playwright helpers for documents-workbench test scripts.

const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:7823/";
const FIXTURES_ROOT = path.resolve(__dirname, "../../test-fixtures");
const DOCX_FIXTURE = path.join(FIXTURES_ROOT, "sample-docx-file-for-testing.docx");
const BUILTIN_SAMPLE = path.resolve(__dirname, "../docs/samples/sample.docx");

function fixtureExists(filePath) {
  return fs.existsSync(filePath);
}

async function createTestPage(viewport = { width: 1280, height: 900 }) {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
  return { browser, page, errors };
}

async function bootApp(page) {
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(async () => {
    if (typeof setSidebarOpen === "function") setSidebarOpen(false);
    if (typeof ensureLazyFeature === "function") await ensureLazyFeature("find-replace");
  });
}

async function waitDocReady(page, timeout = 60000) {
  await page.waitForFunction(
    () => document.getElementById("loadingOverlay")?.classList.contains("hidden")
      && !!document.querySelector(".docx-preview-host"),
    { timeout }
  );
  await page.waitForTimeout(400);
}

async function loadDocxFile(page, filePath, timeout = 60000) {
  if (!fixtureExists(filePath)) throw new Error(`Brak pliku: ${filePath}`);
  await page.locator("#fileInput").setInputFiles(filePath);
  await waitDocReady(page, timeout);
}

async function loadBuiltinSample(page) {
  await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
  await page.waitForSelector(".docx-preview-host p", { timeout: 15000 });
  await page.waitForTimeout(300);
}

function assertNoErrors(errors, label) {
  if (errors.length) throw new Error(`${label} — błędy konsoli:\n${errors.join("\n")}`);
}

module.exports = {
  APP_URL,
  FIXTURES_ROOT,
  DOCX_FIXTURE,
  BUILTIN_SAMPLE,
  fixtureExists,
  createTestPage,
  bootApp,
  waitDocReady,
  loadDocxFile,
  loadBuiltinSample,
  assertNoErrors,
};
