// phone-bar-playwright.js — górny pasek telefonu (app/phone-bar.js, etap 7 wstążki, makieta v4).
// ← zamyka do ekranu startowego, Cofnij w Edycji, „N niezapisane · tryb”, menu trybu, Zapisz w ⋯,
// w Czytaniu bez pływającego „Edytuj”, szeroko nic z tego nie widać. ENGINE=webkit.

const pw = require("playwright");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const errors = [];

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 800 }, hasTouch: true, ...(ENGINE === "chromium" ? { isMobile: true } : {}) });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
  const page = await context.newPage();
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", path.join(__dirname, "..", "docs", "samples", "sample.docx"));
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  await sleep(600);
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  const vis = (sel) => page.evaluate((s) => { const e = document.querySelector(s); return !!e && getComputedStyle(e).display !== "none" && e.getBoundingClientRect().width > 0; }, sel);
  const st = () => page.evaluate(() => document.getElementById("phStatus").textContent);

  check("Czytanie: ←, tryb i ⋯ u góry, bez Cofnij i bez Zapisz", (await vis("#phBackBtn")) && (await vis("#phModeBtn")) && !(await vis("#phUndoBtn")) && !(await vis("#heroSaveBtn")) && (await st()) === "Czytanie");
  // Czytanie: stuknięcie w tekst nie pokazuje pływającego „Edytuj”
  await page.locator(".docx-preview-host p").first().tap();
  await sleep(500);
  check("Czytanie: bez pływającego „Edytuj”", !(await page.evaluate(() => !!document.querySelector(".read-hint"))));

  await page.tap("#phModeBtn");
  await sleep(250);
  const items = await page.evaluate(() => Array.from(document.querySelectorAll(".compose-pop [role=menuitemradio]")).map((b) => `${b.textContent.trim().slice(0, 12)}:${b.getAttribute("aria-checked")}`));
  check("menu trybu: Edycja / Sugerowanie / Czytanie (zaznaczone)", items.length === 3 && /^Edycja/.test(items[0]) && /^Czytanie.*:true$/.test(items[2]), JSON.stringify(items));
  await page.locator(".compose-pop [role=menuitemradio]").first().tap();
  await sleep(500);
  check("Edycja: Cofnij u góry (nieaktywne), status „Edycja”", (await vis("#phUndoBtn")) && (await page.evaluate(() => document.getElementById("phUndoBtn").disabled)) && (await st()) === "Edycja");

  await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.isContentEditable && p.textContent.trim().length > 3);
    placeCaret(ps[0], ps[0].textContent.length);
  });
  await page.keyboard.type(" XYZ", { delay: 10 });
  await sleep(600);
  const dirty = await page.evaluate(() => ({ st: document.getElementById("phStatus").textContent, undo: document.getElementById("phUndoBtn").disabled }));
  check("po pisaniu: „1 niezapisane · Edycja”, Cofnij aktywne", dirty.st === "1 niezapisane · Edycja" && !dirty.undo, JSON.stringify(dirty));
  await page.tap("#phUndoBtn");
  await sleep(600);
  const undone = await page.evaluate(() => ({ text: document.querySelector(".docx-preview-host").textContent.includes("XYZ"), st: document.getElementById("phStatus").textContent }));
  check("Cofnij u góry cofa pisanie", !undone.text && undone.st === "Edycja", JSON.stringify(undone));

  await page.tap("#appMenuBtn");
  await sleep(300);
  check("⋯: jest „Zapisz”", await vis("#phSaveMenuItem"));
  await page.evaluate(() => appFrame.setMenuOpen(false));
  await sleep(200);

  await page.tap("#phBackBtn");
  await sleep(800);
  check("←: dokument zamknięty, ekran startowy", await page.evaluate(() => !originalFileBytes && !document.querySelector(".hero").classList.contains("has-file")));

  await page.setViewportSize({ width: 1280, height: 800 });
  await sleep(400);
  check("szeroko: bez paska telefonu", !(await vis("#phBackBtn")) && !(await vis("#phModeBtn")));

  await browser.close();
  if (errors.length) check("bez błędów w konsoli", false, errors.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) { if (!r.ok) failed++; console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`); }
  console.log(`\n${failed ? "❌" : "✅"} Pasek telefonu [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}
run().catch((e) => { console.error(e); process.exit(1); });
