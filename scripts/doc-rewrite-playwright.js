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

  // „Popraw błędy” = słownik (bez modelu, nawet gdy asystent działa): polskie znaki i jedna litera
  const dict = await page.evaluate(async () => ({
    a: await dwbRewrite.offlineFix("Prosze przesłać dokumęty do piątku, spotkanie odbedzie sie o 12."),
    b: await dwbRewrite.offlineFix("Nie wiem czy zdąże, ale postaram sie."),
  }));
  check("słownik: „Prosze… dokumęty… odbedzie sie” → poprawione", dict.a === "Proszę przesłać dokumenty do piątku, spotkanie odbędzie się o 12.", dict.a);
  check("słownik: „zdąże” → „zdążę”, „sie” → „się”", dict.b === "Nie wiem czy zdążę, ale postaram się.", dict.b);
  // fałszywe „poprawki” złapane na całym przewodniku (2026-10-24): kod, nazwy, obce słowa zostają
  const keep = await page.evaluate(async () => ({
    a: await dwbRewrite.offlineFix("Zapis jako plik (.docx) albo TXT, Markdown, w przeglądarce Edge, papier A4 i Letter."),
    b: await dwbRewrite.offlineFix("Wpisz np. !dzis albo !pozdrawiam, a pole {{wynajmujacy}} zostanie wypełnione."),
    c: (await dwbRewrite.offlineFixMany(["pasek nad dokumentem", "pasek nad dokumentem"], { lists: [true, false] })),
  }));
  check("słownik nie rusza: .docx, Markdown, Edge, Letter", keep.a === "Zapis jako plik (.docx) albo TXT, Markdown, w przeglądarce Edge, papier A4 i Letter.", keep.a);
  check("słownik i reguły nie ruszają skrótów snippetów (!dzis) ani pól {{…}}", keep.b === "Wpisz np. !dzis albo !pozdrawiam, a pole {{wynajmujacy}} zostanie wypełnione.", keep.b);
  check("punkt listy / komórka tabeli: mała litera na początku zostaje (zwykły akapit — wielka)", keep.c[0] === "pasek nad dokumentem" && keep.c[1] === "Pasek nad dokumentem", JSON.stringify(keep.c));

  // ── zaznaczenie przez kilka akapitów: każdy akapit osobno do modelu, lista zmian ──
  const MULTI = [
    "Pierwszy akapit zostaje. Stary tekst do zmiany w pierwszym.",
    "Drugi akapit: stary tekst w całości, termin 7 dni.",
    "Trzeci akapit ma stary początek, a koniec zostaje bez zmian.",
  ];
  const multiIdx = await page.evaluate(async (texts) => {
    const all = await extractParagraphTextsFromDocx(originalFileBytes);
    const paras = docBodyParagraphs(docCanvasEl);
    let i = all.findIndex((x, k) => k > 0 && [0, 1, 2].every((d) => all[k + d]?.trim() && paras[k + d] && !paras[k + d].dataset.lock));
    await applyDocumentEdit({ op: "paragraphBatch", items: texts.map((text, d) => ({ index: i + d, text })) });
    return i;
  }, MULTI);
  const selectAcross = async (from, to) => page.evaluate(async ({ i, from, to }) => {
    const paras = docBodyParagraphs(docCanvasEl);
    // miejsce w akapicie → węzeł tekstu (tekst bywa pocięty na kilka węzłów)
    const pos = (p, at) => {
      const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      let n; let seen = 0;
      while ((n = w.nextNode())) { if (at <= seen + n.data.length) return { n, k: at - seen }; seen += n.data.length; }
      return null;
    };
    const a = pos(paras[i], paras[i].textContent.indexOf(from));
    const b = pos(paras[i + 2], paras[i + 2].textContent.indexOf(to) + to.length);
    paras[i].scrollIntoView({ block: "center" });
    focusDocParagraph(paras[i]);
    const r = document.createRange();
    r.setStart(a.n, a.k);
    r.setEnd(b.n, b.k);
    const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
    document.dispatchEvent(new Event("selectionchange"));
    await new Promise((res) => setTimeout(res, 50));
  }, { i: multiIdx, from, to });
  await selectAcross("Stary tekst", "stary początek");
  const multi = await page.evaluate(async () => {
    window.__rwCalls = [];
    // atrapa: „stary” → „nowy”; w drugim akapicie gubi liczbę 7 (ma wyjść odznaczony)
    window.__localAiMock = async (messages) => {
      const frag = messages[messages.length - 1].content.split("Fragment:\n")[1];
      window.__rwCalls.push(frag);
      await new Promise((r) => setTimeout(r, 30));
      return frag.replace(/stary/gi, (m) => (m[0] === "S" ? "Nowy" : "nowy")).replace(/termin 7 dni/, "krótki termin");
    };
    document.getElementById("rewriteBtn").click();
    for (let i = 0; i < 40 && !document.querySelector(".rw-dialog[open] .rw-modes"); i++) await new Promise((r) => setTimeout(r, 50));
    const scopeLabel = document.querySelector('.rw-scope-opt[data-scope="sel"]')?.textContent;
    const items = dwbRewrite._target().items.map((it) => it.original);
    document.querySelector('.rw-dialog button[data-mode="simple"]').click();
    for (let i = 0; i < 80 && !document.querySelector(".rw-dialog .rw-apply"); i++) await new Promise((r) => setTimeout(r, 50));
    const list = [...document.querySelectorAll(".rw-dialog .rw-item")].map((n) => ({ on: n.querySelector("input")?.checked, warn: n.querySelector(".lai-warn")?.textContent || "" }));
    return { scopeLabel, items, calls: window.__rwCalls, list, apply: document.querySelector(".rw-apply")?.textContent, summary: document.querySelector(".rw-summary")?.textContent };
  });
  check("zaznaczenie przez 3 akapity: „Zaznaczenie · 3 akapity”, pierwszy i ostatni częściowo, środkowy w całości",
    /3\s+akapity/.test(multi.scopeLabel) && multi.items.length === 3 && multi.items[0] === "Stary tekst do zmiany w pierwszym." && multi.items[1] === MULTI[1] && multi.items[2] === "Trzeci akapit ma stary początek", JSON.stringify(multi));
  check("każdy akapit osobno do modelu (3 wywołania, bez łączenia akapitów)", multi.calls.length === 3 && multi.calls.every((c) => !c.includes("\n")), JSON.stringify(multi.calls));
  check("lista 3 zmian; ta bez liczby „7” odznaczona z ostrzeżeniem; przycisk „(2)”",
    multi.list.length === 3 && multi.list.filter((x) => x.on).length === 2 && /7/.test(multi.list[1].warn) && !multi.list[1].on && /\(2\)/.test(multi.apply), JSON.stringify(multi));
  const multiApplied = await page.evaluate(async (i) => {
    const undoBefore = dwbUndo._debug().undo.length;
    document.querySelector(".rw-dialog .rw-apply").click();
    for (let k = 0; k < 60 && document.querySelector(".rw-dialog[open]"); k++) await new Promise((r) => setTimeout(r, 50));
    await new Promise((r) => setTimeout(r, 400));
    const texts = (await extractParagraphTextsFromDocx(originalFileBytes)).slice(i, i + 3);
    const scan = await scanDocxRevisions(originalFileBytes);
    const undoSteps = dwbUndo._debug().undo.length - undoBefore;
    await dwbUndo.undo();
    const undone = (await extractParagraphTextsFromDocx(originalFileBytes)).slice(i, i + 3);
    return { texts, authors: [...new Set(scan.changes.map((c) => c.author))], undoSteps, undone };
  }, multiIdx);
  check("wstawione 2 zaznaczone: początek 1. i koniec 3. akapitu nietknięte, 2. bez zmian (odznaczony)",
    multiApplied.texts[0] === "Pierwszy akapit zostaje. Nowy tekst do zmiany w pierwszym." && multiApplied.texts[1] === MULTI[1]
      && multiApplied.texts[2] === "Trzeci akapit ma nowy początek, a koniec zostaje bez zmian.", JSON.stringify(multiApplied));
  check("poprawki „Asystent AI”, jeden krok Cofnij przywraca wszystkie akapity",
    multiApplied.authors.join() === "Asystent AI" && multiApplied.undoSteps === 1 && multiApplied.undone.join("|") === MULTI.join("|"), JSON.stringify(multiApplied));

  // trzykrotne kliknięcie: zaznaczenie kończy się na samym początku następnego akapitu — ten się nie liczy
  const triple = await page.evaluate(async (i) => {
    const paras = docBodyParagraphs(docCanvasEl);
    focusDocParagraph(paras[i]);
    const first = (p) => document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode();
    const r = document.createRange();
    r.setStart(first(paras[i]), 0);
    r.setEnd(first(paras[i + 1]), 0);
    const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
    document.dispatchEvent(new Event("selectionchange"));
    await new Promise((res) => setTimeout(res, 50));
    document.getElementById("rewriteBtn").click();
    for (let k = 0; k < 40 && !document.querySelector(".rw-dialog[open] .rw-scope"); k++) await new Promise((res) => setTimeout(res, 50));
    const out = { label: document.querySelector('.rw-scope-opt[data-scope="sel"]')?.textContent, n: dwbRewrite._target().items ? dwbRewrite._target().items.length : 1, idx: dwbRewrite._target().paraIndex };
    document.querySelector(".rw-dialog").close();
    return out;
  }, multiIdx);
  check("trzykrotne kliknięcie (koniec na początku następnego akapitu) = jeden akapit", triple.n === 1 && triple.idx === multiIdx && triple.label === "Zaznaczenie", JSON.stringify(triple));

  // Zatrzymaj w połowie: gotowe zmiany da się wstawić
  await selectAcross("Pierwszy", "koniec zostaje bez zmian.");
  const stopped = await page.evaluate(async () => {
    window.__localAiMock = async (messages) => {
      const frag = messages[messages.length - 1].content.split("Fragment:\n")[1];
      await new Promise((r) => setTimeout(r, 400));
      return frag.replace(/stary/gi, "nowy");
    };
    document.getElementById("rewriteBtn").click();
    for (let i = 0; i < 40 && !document.querySelector(".rw-dialog[open] .rw-modes"); i++) await new Promise((r) => setTimeout(r, 50));
    document.querySelector('.rw-dialog button[data-mode="simple"]').click();
    for (let i = 0; i < 40 && !document.querySelector(".rw-dialog .rw-stop"); i++) await new Promise((r) => setTimeout(r, 25));
    const progress = document.querySelector(".rw-summary")?.textContent;
    await new Promise((r) => setTimeout(r, 450));
    document.querySelector(".rw-dialog .rw-stop").click();
    for (let i = 0; i < 40 && !document.querySelector(".rw-dialog .rw-apply"); i++) await new Promise((r) => setTimeout(r, 50));
    const out = {
      progress,
      done: dwbRewrite._results().filter((r) => r.done).length,
      stoppedNote: [...document.querySelectorAll(".rw-dialog .lai-warn")].some((n) => /Zatrzymano/.test(n.textContent)),
      apply: document.querySelector(".rw-apply")?.textContent,
    };
    document.querySelector(".rw-dialog").close();
    return out;
  });
  check("postęp „Akapit 1 z 3…”, Zatrzymaj → gotowe zostają do wstawienia",
    /Akapit 1 z 3/.test(stopped.progress) && stopped.done >= 1 && stopped.done < 3 && stopped.stoppedNote && /\(\d\)/.test(stopped.apply), JSON.stringify(stopped));

  // ── „Popraw błędy” dla całego dokumentu (słownik, bez AI) ──
  const docFix = await page.evaluate(async (i) => {
    delete window.__localAiMock;
    await applyDocumentEdit({ op: "paragraphBatch", items: [
      { index: i, text: "Prosze przesłać dokumęty do piątku." },
      { index: i + 2, text: "Spotkanie odbedzie sie o 12." },
    ] });
    const before = await extractParagraphTextsFromDocx(originalFileBytes);
    // kursor gdzie indziej (pierwszy akapit) — zakres przełączamy na „Cały dokument”
    const p0 = docBodyParagraphs(docCanvasEl).find((p) => !p.dataset.lock && p.textContent.trim());
    focusDocParagraph(p0);
    const sel = window.getSelection(); const r = document.createRange(); r.selectNodeContents(p0); r.collapse(true); sel.removeAllRanges(); sel.addRange(r);
    document.getElementById("rewriteBtn").click();
    for (let k = 0; k < 40 && !document.querySelector(".rw-dialog[open] .rw-scope"); k++) await new Promise((res) => setTimeout(res, 50));
    document.querySelector('.rw-scope-opt[data-scope="doc"]').click();
    const disabled = [...document.querySelectorAll(".rw-dialog .rw-modes button")].filter((b) => b.disabled).map((b) => b.dataset.mode);
    const ownOff = document.querySelector(".rw-dialog .rw-own input").disabled;
    document.querySelector('.rw-dialog button[data-mode="fix"]').click();
    for (let k = 0; k < 200 && !document.querySelector(".rw-dialog .rw-apply"); k++) await new Promise((res) => setTimeout(res, 50));
    const items = dwbRewrite._results().filter((x) => x.done && x.suggestion !== x.item.original).map((x) => x.item.paraIndex);
    const badge = document.querySelector(".rw-dialog .ai-badge")?.textContent;
    const undoBefore = dwbUndo._debug().undo.length;
    document.querySelector(".rw-dialog .rw-apply-plain").click();
    for (let k = 0; k < 60 && document.querySelector(".rw-dialog[open]"); k++) await new Promise((res) => setTimeout(res, 50));
    await new Promise((res) => setTimeout(res, 400));
    const after = await extractParagraphTextsFromDocx(originalFileBytes);
    const others = before.every((x, k) => k === i || k === i + 2 || x === after[k] || items.includes(k));
    return { disabled, ownOff, items, badge, a: after[i], c: after[i + 2], others, undoSteps: dwbUndo._debug().undo.length - undoBefore, revisions: (await scanDocxRevisions(originalFileBytes)).changes.length };
  }, multiIdx);
  check("„Cały dokument”: działa tylko „Popraw błędy” (opcje AI i własne polecenie wyłączone)",
    docFix.disabled.join() === "simple,formal,short,polite" && docFix.ownOff, JSON.stringify(docFix));
  check("cały dokument słownikiem: oba akapity z błędami poprawione, podpis „Słownik”",
    docFix.items.includes(multiIdx) && docFix.items.includes(multiIdx + 2) && /Słownik/.test(docFix.badge)
      && docFix.a === "Proszę przesłać dokumenty do piątku." && docFix.c === "Spotkanie odbędzie się o 12.", JSON.stringify(docFix));
  check("cały dokument: jeden krok Cofnij, bez śledzenia — bez poprawek; reszta tekstu nietknięta",
    docFix.undoSteps === 1 && docFix.revisions === 0 && docFix.others, JSON.stringify(docFix));

  // bez kursora w dokumencie: okno od razu na „Cały dokument” (zamiast komunikatu „postaw kursor”)
  const noCaret = await page.evaluate(async () => {
    document.activeElement?.blur?.();
    window.getSelection().removeAllRanges();
    lastDocCaret = null;
    document.getElementById("rewriteBtn").click();
    for (let k = 0; k < 40 && !document.querySelector(".rw-dialog[open] .rw-scope"); k++) await new Promise((res) => setTimeout(res, 50));
    const out = {
      doc: document.querySelector('.rw-scope-opt[data-scope="doc"]')?.getAttribute("aria-pressed"),
      selDisabled: document.querySelector('.rw-scope-opt[data-scope="sel"]')?.disabled,
      fix: !document.querySelector('.rw-dialog button[data-mode="fix"]')?.disabled,
    };
    document.querySelector(".rw-dialog")?.close();
    return out;
  });
  check("bez kursora: okno na „Cały dokument”, „Popraw błędy” dostępne", noCaret.doc === "true" && noCaret.selDisabled && noCaret.fix, JSON.stringify(noCaret));

  assertNoErrors(errors, "doc-rewrite");
  await browser.close();
  let failed = 0;
  results.forEach((r) => { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` (${String(r.detail).slice(0, 600)})`}`); if (!r.ok) failed++; });
  if (failed) process.exit(1);
  console.log(`\n✅ doc-rewrite: ${results.length}/${results.length} (${process.env.ENGINE || "chromium"})`);
}
run().catch((e) => { console.error(e); process.exit(1); });
