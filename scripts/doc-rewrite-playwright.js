// doc-rewrite-playwright.js — „✦ Przepisz” (atrapa AI): zaznaczenie w akapicie → podgląd różnicy
// → „Wstaw jako poprawkę” = w:ins/w:del autora „Asystent AI”, jeden krok Cofnij, Odrzuć wraca do
// oryginału; ostrzeżenie o zgubionej liczbie; „Zamień bez śledzenia”. ENGINE=webkit — Safari.
const { createTestPage, bootApp, loadDocxFile, assertNoErrors } = require("./docx-test-helpers");

const results = [];
const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail });
const ORIGINAL = "No to wyślijcie nam te papiery do piątku, bo inaczej nie zdążymy z terminem 14 dni.";

async function run() {
  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await loadDocxFile(page, "docs/samples/headings-sample.docx");
  const idx = await page.evaluate(async (original) => {
    appFrame.setReadOnly(false);
    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    const index = texts.map((t, i) => (t && t.trim() ? i : -1)).filter((i) => i >= 0)[3];
    await applyDocumentEdit({ op: "paragraphBatch", items: [{ index, text: original }] });
    return index;
  }, ORIGINAL);

  // zaznaczenie „wyślijcie nam te papiery” w podglądzie i klik w ✦ Przepisz
  const opened = await page.evaluate(async (index) => {
    const p = docBodyParagraphs(docCanvasEl)[index];
    p.scrollIntoView({ block: "center" });
    focusDocParagraph(p);
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let node; let found = null;
    while ((node = walker.nextNode())) { const i = node.data.indexOf("wyślijcie nam te papiery"); if (i >= 0) { found = { node, i }; break; } }
    if (!found) return { error: "no text node" };
    const r = document.createRange();
    r.setStart(found.node, found.i);
    r.setEnd(found.node, found.i + "wyślijcie nam te papiery".length);
    const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
    document.dispatchEvent(new Event("selectionchange"));
    await new Promise((res) => setTimeout(res, 50));
    window.__localAiMock = (messages) => {
      window.__rwPrompt = messages.map((m) => m.content).join("\n");
      return "Oto propozycja: „prześlijcie nam te dokumenty”";
    };
    document.getElementById("rewriteBtn").click();
    for (let i = 0; i < 40 && !document.querySelector(".rw-dialog[open]"); i++) await new Promise((res) => setTimeout(res, 50));
    return { open: !!document.querySelector(".rw-dialog[open]"), target: dwbRewrite._target() };
  }, idx);
  check("okno Przepisz otwarte dla zaznaczenia (nie całego akapitu)", opened.open && opened.target?.original === "wyślijcie nam te papiery", JSON.stringify(opened));

  const preview = await page.evaluate(async () => {
    document.querySelector('.rw-dialog button[data-mode="formal"]').click();
    for (let i = 0; i < 40 && !document.querySelector(".rw-dialog .rw-apply"); i++) await new Promise((r) => setTimeout(r, 50));
    return {
      ins: [...document.querySelectorAll(".rw-diff ins")].map((n) => n.textContent.trim()).join(" "),
      del: [...document.querySelectorAll(".rw-diff del")].map((n) => n.textContent.trim()).join(" "),
      prompt: window.__rwPrompt,
      badge: !!document.querySelector(".rw-dialog .ai-badge"),
    };
  });
  check("polecenie „Formalnie” + sam zaznaczony fragment idą do modelu", /stylu formalnym/.test(preview.prompt) && /Fragment:\nwyślijcie nam te papiery$/.test(preview.prompt), preview.prompt);
  check("odpowiedź oczyszczona ze wstępu i cudzysłowów; podgląd różnicy", /prześlijcie/.test(preview.ins) && /papiery/.test(preview.del) && !/Oto/.test(preview.ins), JSON.stringify(preview));
  check("podgląd podpisany jako AI", preview.badge);

  const applied = await page.evaluate(async (index) => {
    document.querySelector(".rw-dialog .rw-apply").click();
    for (let i = 0; i < 60 && !document.querySelector(".toast"); i++) await new Promise((r) => setTimeout(r, 50));
    await new Promise((r) => setTimeout(r, 200));
    const scan = await scanDocxRevisions(originalFileBytes);
    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    return {
      kinds: scan.changes.map((c) => c.kind || c.type),
      authors: [...new Set(scan.changes.map((c) => c.author))],
      count: scan.changes.length,
      text: texts[index],
      undo: dwbUndo._debug().undo.slice(-1)[0],
      dirty: hasUnsavedChanges,
    };
  }, idx);
  check("poprawki w pliku, autor „Asystent AI”", applied.count >= 1 && applied.authors.length === 1 && applied.authors[0] === "Asystent AI", JSON.stringify(applied));
  check("jeden krok Cofnij z nazwą „przepisanie (AI)”", applied.undo === "undoOpRewrite", JSON.stringify(applied.undo));

  // Odrzuć wszystkie → oryginał; Akceptuj wszystkie → nowy tekst
  const review = await page.evaluate(async (index) => {
    const before = originalFileBytes;
    const scan = await scanDocxRevisions(before);
    await applyDocumentEdit({ op: "revisions", action: "reject", ids: scan.changes.map((c) => c.id) });
    const rejected = (await extractParagraphTextsFromDocx(originalFileBytes))[index];
    await dwbUndo.undo();
    const scan2 = await scanDocxRevisions(originalFileBytes);
    await applyDocumentEdit({ op: "revisions", action: "accept", ids: scan2.changes.map((c) => c.id) });
    const accepted = (await extractParagraphTextsFromDocx(originalFileBytes))[index];
    return { rejected, accepted };
  }, idx);
  check("Odrzuć → dokładnie oryginalny tekst", review.rejected === ORIGINAL, JSON.stringify(review));
  check("Akceptuj → tekst z propozycją, reszta akapitu bez zmian",
    review.accepted === ORIGINAL.replace("wyślijcie nam te papiery", "prześlijcie nam te dokumenty"), JSON.stringify(review));

  // zgubiona liczba → ostrzeżenie; „Zamień bez śledzenia” → bez poprawek
  const plain = await page.evaluate(async (index) => {
    const p = docBodyParagraphs(docCanvasEl)[index];
    focusDocParagraph(p);
    const sel = window.getSelection(); const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); sel.removeAllRanges(); sel.addRange(r);
    document.dispatchEvent(new Event("selectionchange"));
    window.__localAiMock = () => "Prosimy o przesłanie dokumentów do piątku.";
    document.getElementById("rewriteBtn").click();
    for (let i = 0; i < 40 && !document.querySelector(".rw-dialog[open] .rw-modes"); i++) await new Promise((res) => setTimeout(res, 50));
    const whole = dwbRewrite._target()?.start === 0;
    document.querySelector('.rw-dialog button[data-mode="short"]').click();
    for (let i = 0; i < 40 && !document.querySelector(".rw-dialog .rw-apply-plain"); i++) await new Promise((res) => setTimeout(res, 50));
    const warn = document.querySelector(".rw-dialog .lai-warn")?.textContent || "";
    document.querySelector(".rw-dialog .rw-apply-plain").click();
    await new Promise((res) => setTimeout(res, 400));
    const scan = await scanDocxRevisions(originalFileBytes);
    delete window.__localAiMock;
    return { whole, warn, revisions: scan.changes.length, text: (await extractParagraphTextsFromDocx(originalFileBytes))[index] };
  }, idx);
  check("kursor bez zaznaczenia → cały akapit", plain.whole, JSON.stringify(plain));
  check("brak liczby z oryginału → ostrzeżenie „14”", /14/.test(plain.warn), plain.warn);
  check("„Zamień bez śledzenia” → nowy tekst bez poprawek", plain.revisions === 0 && plain.text === "Prosimy o przesłanie dokumentów do piątku.", JSON.stringify(plain));

  assertNoErrors(errors, "doc-rewrite");
  await browser.close();
  let failed = 0;
  results.forEach((r) => { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` (${String(r.detail).slice(0, 600)})`}`); if (!r.ok) failed++; });
  if (failed) process.exit(1);
  console.log(`\n✅ doc-rewrite: ${results.length}/${results.length} (${process.env.ENGINE || "chromium"})`);
}
run().catch((e) => { console.error(e); process.exit(1); });
