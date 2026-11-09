// quiet-edits-playwright.js — operacje, po których strona wygląda tak samo, NIE przerysowują
// dokumentu (etap „przeładowania”, 2026-11-10; pomiar scripts/reload-bench.js):
// edycja / odpowiedź / zamknięcie komentarza i metadane. Musi być prawdą: bez pełnego renderu,
// kursor zostaje, nowe dane widać (karty, licznik), podgląd = plik, Cofnij przywraca.
//
//   node scripts/quiet-edits-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/quiet-edits-playwright.js
const fs = require("fs");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const SAMPLE = path.resolve(__dirname, "../docs/samples/review-sample.docx");
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

async function main() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "cicho.docx", mimeType: DOCX_MIME, buffer: fs.readFileSync(SAMPLE) });
  await idle(page);
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await idle(page);
  await page.evaluate(() => {
    window.__renders = 0;
    const orig = window.renderCurrentDocument;
    window.renderCurrentDocument = function counted(...a) { window.__renders++; return orig.apply(this, a); };
  });

  // kursor w akapicie z tekstem (nie w komentarzu) — ma zostać po operacji
  const caretSpot = () => page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const sel = getSelection();
    const i = sel?.rangeCount ? ps.findIndex((p) => p.contains(sel.anchorNode)) : -1;
    return { i, off: i >= 0 ? getCaretOffset(ps[i]) : -1, focus: !!document.activeElement?.closest?.(".docx-preview-host, .docx-edit-root") };
  });
  const placeAt = () => page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const p = ps.find((x) => !x.dataset.lock && x.textContent.length > 20);
    p.closest(".docx-edit-root")?.focus();
    placeCaret(p, 5);
  });
  const previewIsFile = () => page.evaluate(async () => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const dom = ps.map((p) => (p.dataset.lock ? null : previewRunsToPlainText(extractRunsFromPreviewParagraph(p)).replace(/[\t\n]/g, "")));
    const xml = (await extractParagraphTextsFromDocx(await buildDocumentForSave())).map((t) => String(t).replace(/[\t\n]/g, ""));
    return dom.length === xml.length && dom.every((t, i) => t === null || t === xml[i]);
  });
  const comments = () => page.evaluate(async () => (await scanDocxRevisions(await buildDocumentForSave())).comments.map((c) => ({ id: c.id, text: c.text, done: !!c.done, parentId: c.parentId || null, replies: (c.replies || []).map((r) => r.text) })));

  const coreXml = () => page.evaluate(async () => { const f = (await JSZip.loadAsync(await buildDocumentForSave())).file("docProps/core.xml"); return f ? f.async("string") : "(brak docProps/core.xml)"; });
  const first = (await comments()).find((c) => !c.parentId);
  check("plik ma komentarz do edycji", !!first, JSON.stringify(await comments()));
  const ops = [
    { name: "edycja komentarza", edit: { op: "commentEdit", id: first.id, rich: [[{ text: "Nowa treść komentarza" }]] }, ok: (cs) => cs.some((c) => c.id === first.id && c.text === "Nowa treść komentarza") },
    { name: "odpowiedź na komentarz", edit: { op: "commentReply", id: first.id, rich: [[{ text: "Odpowiedź testowa" }]], author: "Test", initials: "T" }, ok: (cs) => cs.some((c) => c.id === first.id && c.replies.includes("Odpowiedź testowa")) },
    { name: "zamknięcie komentarza", edit: { op: "commentDone", id: first.id, done: true }, ok: (cs) => cs.some((c) => c.id === first.id && c.done) },
    { name: "metadane: tytuł", edit: { op: "coreMetadata", fields: { title: "Tytuł bez przerysowania" } }, ok: null },
  ];
  const coreBefore = await coreXml();
  for (const op of ops) {
    await placeAt();
    const before = await caretSpot();
    const textsBefore = await comments();
    await page.evaluate(() => { window.__renders = 0; });
    const n = await page.evaluate((e) => applyDocumentEdit(e), op.edit);
    await idle(page);
    const after = await caretSpot();
    const renders = await page.evaluate(() => window.__renders);
    check(`${op.name}: zmiana w pliku, bez przerysowania dokumentu`, n > 0 && renders === 0, `zmian ${n}, renderów ${renders}`);
    check(`…kursor zostaje (akapit i miejsce)`, after.i === before.i && after.off === before.off, JSON.stringify({ before, after }));
    if (op.ok) check(`…nowe dane komentarza w pliku`, op.ok(await comments()), JSON.stringify((await comments()).slice(0, 4)));
    else check(`…tytuł w pliku`, /Tytuł bez przerysowania/.test(await coreXml()), (await coreXml()).slice(0, 200));
    check(`…podgląd = plik`, await previewIsFile());
    // Cofnij — przywraca stan sprzed operacji
    await page.waitForTimeout(300);
    await page.click("#undoBtn");
    await idle(page);
    const back = await comments();
    if (op.ok) check(`…Cofnij przywraca`, JSON.stringify(back) === JSON.stringify(textsBefore), JSON.stringify(back.slice(0, 4)));
    else { const c = await coreXml(); check(`…Cofnij przywraca metadane sprzed zmiany (ten plik nie miał core.xml)`, c === coreBefore, c.slice(0, 200)); }
  }
  // podświetlenia komentarzy dalej na miejscu (dane z nowej wersji pliku)
  await page.evaluate((e) => applyDocumentEdit(e), { op: "commentEdit", id: first.id, rich: [[{ text: "Karta po cichu" }]] });
  await idle(page);
  const painted = await page.evaluate(async () => { await composeUi.loadComments(); return CSS.highlights ? (CSS.highlights.get("dwb-comment")?.size || 0) + (CSS.highlights.get("dwb-comment-done")?.size || 0) : -1; });
  check("podświetlenia komentarzy po cichej zmianie", painted !== 0, String(painted));

  check("bez błędów w konsoli", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();
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
