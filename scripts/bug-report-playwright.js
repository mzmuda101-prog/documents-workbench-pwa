// bug-report-playwright.js — „Zgłoś błąd” (app/bug-report.js) + odtworzenie (scripts/replay-report.js):
// prawdziwa sesja (pisanie, Enter, Backspace, pogrubienie skrótem, wyrównanie z paska, Cofnij
// skrótem, wklejenie) → skrót Ctrl/⌘+Alt+Shift+B → pobrany plik → odtworzenie daje TEN SAM wynik.
//
//   node scripts/bug-report-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/bug-report-playwright.js
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const SAMPLE = path.resolve(__dirname, "../docs/samples/headings-sample.docx");
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

async function main() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 }, acceptDownloads: true });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "zgloszenie.docx", mimeType: DOCX_MIME, buffer: fs.readFileSync(SAMPLE) });
  await idle(page);
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await idle(page);

  // sesja jak u użytkownika: klik w tekst, pisanie, Enter, Backspace, skróty, pasek, Cofnij, wklejenie
  const para = (i) => page.locator(".docx-preview-host p.docx-editable-p").nth(i);
  await para(1).click({ position: { x: 30, y: 8 } });
  await page.keyboard.type(" Zażółć");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Nowy akapit");
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.type(" gruby");
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.press("Backspace");
  await page.keyboard.press("Backspace");
  await idle(page);
  await page.evaluate(() => composeUi.applyAlign("center"));
  await idle(page);
  await para(3).click({ position: { x: 20, y: 8 } });
  await page.keyboard.type("XYZ ");
  await page.waitForTimeout(1700);
  await page.keyboard.press(`${MOD}+z`); // cofa „XYZ ”
  await idle(page);
  await page.evaluate(() => {
    const dt = new DataTransfer();
    dt.setData("text/plain", "Wklejony tekst");
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[4];
    p.closest(".docx-edit-root")?.focus();
    placeCaret(p, 0);
    p.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  });
  await idle(page);
  // Cofnij i Ponów przyciskami (droga przez dwbUndo — w dzienniku jako osobne działania)
  await page.waitForTimeout(1700);
  await page.click("#undoBtn");
  await idle(page);
  await page.click("#redoBtn");
  await idle(page);
  await page.waitForTimeout(300);
  await page.click("#undoBtn"); // sam Cofnij — bez odtworzenia tego kroku wynik byłby inny
  await idle(page);
  const before = await page.evaluate(async () => (await extractParagraphTextsFromDocx(await buildDocumentForSave())).map(String));

  // Zgłoś błąd — skrót
  const dl = page.waitForEvent("download", { timeout: 15000 });
  await page.keyboard.press(`${MOD}+Alt+Shift+KeyB`);
  const download = await dl;
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "dwb-bug-")), download.suggestedFilename());
  await download.saveAs(file);
  const rep = JSON.parse(fs.readFileSync(file, "utf8"));
  check("skrót zapisuje plik zgłoszenia (.dwbbug.json)", /\.dwbbug\.json$/.test(file) && rep.kind === "dwb-bug-report", path.basename(file));
  const kinds = rep.actions.map((a) => a.a + (a.fromInput ? "*" : ""));
  check("dziennik: otwarcie, klawisze, klik, operacja z paska, wklejenie, Cofnij/Ponów przyciskami", ["open", "key", "click", "edit", "paste", "undo", "redo"].every((k) => kinds.includes(k)), kinds.join(" "));
  check("…Cofnij skrótem: klawisz + krok oznaczony jako jego skutek (odtworzenie bez podwójnego cofania)", rep.actions.some((a) => a.a === "key" && /Meta|Control/.test(a.mods) && a.key.toLowerCase() === "z") && rep.actions.some((a) => a.a === "undo" && a.fromInput), JSON.stringify(rep.actions.filter((a) => a.a === "undo" || (a.a === "key" && a.key.toLowerCase() === "z"))));
  check("…Cofnij / Ponów przyciskami: kroki do odtworzenia (nie skutek klawisza)", rep.actions.some((a) => a.a === "undo" && !a.fromInput) && rep.actions.some((a) => a.a === "redo" && !a.fromInput));
  check("…dokument z początku i stan końcowy w pliku", rep.start.length > 1000 && rep.end.length > 1000);
  check("bez błędów strony w sesji", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();

  // odtworzenie
  let out = "";
  try { out = execFileSync("node", [path.join(__dirname, "replay-report.js")], { env: { ...process.env, REPORT: file, ENGINE }, encoding: "utf8", timeout: 240000 }); } catch (e) { out = String(e.stdout || "") + String(e.stderr || e.message); }
  check("odtworzenie: wynik taki jak u zgłaszającego", /odtworzone: wynik taki jak u zgłaszającego/.test(out), out.split("\n").slice(-6).join(" / "));
  check("…podgląd = plik, bez błędów strony", /✅ podgląd = plik/.test(out) && /✅ błędy strony: brak/.test(out), out.split("\n").slice(-4).join(" / "));
  check("…zgłoszenie miało zmiany (sesja nie była pusta, wklejenie cofnięte przyciskiem)", before.join("\n").includes("Zażółć") && !before.join("\n").includes("Wklejony tekst"), before.slice(0, 6).join(" | "));

  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail || ""}`}`);
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}

main().catch((e) => {
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail || ""}`}`);
  console.error(e);
  process.exit(1);
});
