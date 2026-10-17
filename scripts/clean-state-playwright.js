// clean-state-playwright.js — licznik „niezapisane: N zmian” gaśnie, gdy treść wróciła do zapisanej
// bez Cofnij (zgłoszenie 2026-10-17: wpisane i skasowane → dalej „13 zmian”). undo.js
// checkBackToSaved: po pauzie te same bajty + brak zmian w podglądzie, a po operacjach na pliku
// porównanie zawartości części .docx. Prawdziwa zmiana dalej liczy się normalnie.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(300));
const dirty = (page) => page.evaluate(() => ({ dirty: hasUnsavedChanges, count: dwbUndo.changesSinceSave() }));
// czeka, aż sprawdzanie po pauzie (600 ms + porównanie) zgasi stan albo minie czas
const settle = (page, wantClean) => page.waitForFunction((w) => hasUnsavedChanges === !w, wantClean, { timeout: 4000 }).catch(() => {}).then(() => page.waitForTimeout(200));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);
  await page.keyboard.type("Akapit pierwszy.");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Akapit drugi.");
  await idle(page);
  // „zapisany” stan = ten (jak po Zapisz)
  await page.evaluate(async () => { await mergeInlineEditsIntoBytes(); await refreshInlineEditBaseline(originalFileBytes); setDirtyState(false); });
  await idle(page);
  check("start: stan zapisany", !(await dirty(page)).dirty, JSON.stringify(await dirty(page)));

  // 1) wpisane i skasowane
  await page.keyboard.type(" Dopisek", { delay: 20 });
  await page.waitForTimeout(1700); // osobny krok pisania
  let d = await dirty(page);
  check("po dopisaniu: niezapisane", d.dirty && d.count >= 1, JSON.stringify(d));
  for (let i = 0; i < 8; i++) await page.keyboard.press("Backspace");
  await settle(page, true);
  d = await dirty(page);
  check("wpisane i skasowane → licznik gaśnie", !d.dirty && d.count === 0, JSON.stringify(d));

  // 2) Enter + Backspace (operacja na pliku → nowe bajty, ta sama treść)
  await page.keyboard.press("Enter");
  await idle(page);
  await page.keyboard.press("Backspace");
  await idle(page);
  await settle(page, true);
  d = await dirty(page);
  check("Enter + Backspace → licznik gaśnie", !d.dirty, JSON.stringify(d));

  // 3) pogrubienie zaznaczenia i zdjęcie pogrubienia
  await page.keyboard.down("Shift");
  for (let i = 0; i < 6; i++) await page.keyboard.press("ArrowLeft");
  await page.keyboard.up("Shift");
  await page.keyboard.press(`${MOD}+b`);
  await page.waitForTimeout(300);
  check("po pogrubieniu: niezapisane", (await dirty(page)).dirty, JSON.stringify(await dirty(page)));
  await page.keyboard.press(`${MOD}+b`);
  await settle(page, true);
  d = await dirty(page);
  check("pogrubienie wł./wył. → licznik gaśnie", !d.dirty, JSON.stringify(d));

  // 4) prawdziwa zmiana zostaje
  await page.keyboard.press("End");
  await page.keyboard.type(" Zostaje", { delay: 10 });
  await page.waitForTimeout(1500);
  d = await dirty(page);
  check("prawdziwa zmiana: dalej niezapisane", d.dirty && d.count >= 1, JSON.stringify(d));

  // 5) Cofnij po „czystym” stanie dalej działa (historia nie zginęła)
  await page.keyboard.press(`${MOD}+z`);
  await idle(page);
  const txt = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[1].textContent);
  check("Cofnij nadal działa", !txt.includes("Zostaje"), txt);

  check("brak błędów w konsoli", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  if (failed.length) process.exit(1);
}

run().catch((err) => { console.error(err); process.exit(1); });
