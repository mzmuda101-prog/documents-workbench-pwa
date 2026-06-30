// visual-audit-playwright.js — wykrywa problemy wizualne podglądu (np. tofu przy bulletach).
// Screenshot opcjonalnie: VISUAL_AUDIT_SHOT=1 npm run test:visual

const fs = require("fs");
const path = require("path");
const {
  DOCX_FIXTURE,
  fixtureExists,
  createTestPage,
  bootApp,
  loadDocxFile,
  assertNoErrors,
} = require("./docx-test-helpers");

const OUT_DIR = path.resolve(__dirname, "../output/visual-audit");

async function run() {
  if (!fixtureExists(DOCX_FIXTURE)) {
    console.warn(`⚠️  Pominięto — brak fixture: ${DOCX_FIXTURE}`);
    process.exit(0);
  }

  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await loadDocxFile(page, DOCX_FIXTURE, 90000);

  const audit = await page.evaluate(() => {
    const host = document.querySelector(".docx-preview-host");
    if (typeof fixDocxBulletRendering === "function") fixDocxBulletRendering(host);
    const report = typeof auditDocxVisualIssues === "function"
      ? auditDocxVisualIssues(host)
      : { issues: [], bulletsFixed: 0, brokenBullets: 0 };
    const listItems = document.querySelectorAll('.docx-preview-host p[class*="docx-num-"]').length;
    return { ...report, listItems };
  });

  if (process.env.VISUAL_AUDIT_SHOT === "1") {
    fs.mkdirSync(OUT_DIR, { recursive: true });
    const shot = path.join(OUT_DIR, "fixture-preview.png");
    await page.locator(".docx-preview-host").screenshot({ path: shot, fullPage: true });
    console.log(`  → screenshot: ${shot}`);
  }

  if (audit.issues.length > 0) {
    const sample = audit.issues.slice(0, 3).map((i) => `${i.type}:${i.className}`).join(", ");
    throw new Error(`Problemy wizualne (${audit.issues.length}): ${sample}`);
  }
  if (audit.brokenBullets > 0) {
    const sample = audit.issues.slice(0, 3).map((i) => `${i.type}:${i.className}`).join(", ");
    throw new Error(`${audit.brokenBullets} nienaprawionych bulletów (np. ${sample})`);
  }

  assertNoErrors(errors, "visual-audit");
  await browser.close();
  console.log(
    `✅  visual-audit-playwright passed (${audit.listItems} list-item, ${audit.bulletsFixed} bullet + ${audit.numberedFixed || 0} numbered)`
  );
}

run().catch((err) => {
  console.error("❌  visual-audit-playwright failed:", err.message || err);
  process.exit(1);
});
