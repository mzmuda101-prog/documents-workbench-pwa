// doc-ask-playwright.js — „Zapytaj dokument”: dobór fragmentów bez modelu (odmiany polskie),
// odpowiedź atrapy AI z odnośnikami [n] → skok do akapitu, ostrzeżenie przy braku odnośnika,
// odmowa pobrania modelu → same fragmenty. ENGINE=webkit — Safari.
const { createTestPage, bootApp, loadDocxFile, assertNoErrors } = require("./docx-test-helpers");

const results = [];
const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail });

async function run() {
  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await loadDocxFile(page, "docs/samples/headings-sample.docx");
  const setup = await page.evaluate(async () => {
    setSidebarOpen(true);
    document.getElementById("panel-semantic-search").open = true;
    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    const idx = texts.map((t, i) => (t && t.trim() ? i : -1)).filter((i) => i >= 0).slice(3, 6);
    await applyDocumentEdit({ op: "paragraphBatch", items: [
      { index: idx[0], text: "Czynsz płatny jest do dziesiątego dnia każdego miesiąca na rachunek Wynajmującego." },
      { index: idx[1], text: "Najemca może wypowiedzieć umowę z zachowaniem trzydziestodniowego okresu wypowiedzenia, czyli 30 dni." },
      { index: idx[2], text: "Kaucja zwrotna wynosi dwukrotność miesięcznego czynszu." },
    ] });
    return { idx };
  });

  // 1) dobór fragmentów bez modelu: odmiana „wypowiedzenie” ↔ „wypowiedzenia”
  const lexical = await page.evaluate(() => docAskLexical(buildSemanticPassages(), "Ile dni ma najemca na wypowiedzenie umowy?").map((r) => r.text));
  check("dopasowanie słów: akapit o wypowiedzeniu na pierwszym miejscu", /wypowiedzieć umowę/.test(lexical[0] || ""), JSON.stringify(lexical.slice(0, 2)));

  // 2) atrapa AI z odnośnikiem
  const answered = await page.evaluate(async () => {
    let prompt = "";
    window.__localAiMock = (messages) => { prompt = messages.map((m) => m.content).join("\n"); return "Najemca ma 30 dni na wypowiedzenie umowy [1]."; };
    document.getElementById("docAskQuery").value = "Ile dni ma najemca na wypowiedzenie umowy?";
    await docAskRun();
    const cite = document.querySelector("#docAskAnswer .doc-ask-cite");
    const out = {
      text: document.querySelector("#docAskAnswer .doc-ask-text")?.textContent,
      cite: cite?.textContent,
      warn: !!document.querySelector("#docAskAnswer .doc-ask-warn"),
      sources: document.querySelectorAll("#docAskAnswer .doc-ask-source").length,
      promptHasFragment: /\[1\] Najemca może wypowiedzieć/.test(prompt),
      promptHasRule: /wyłącznie na podstawie/.test(prompt),
    };
    cite?.click();
    await new Promise((r) => setTimeout(r, 200));
    out.highlighted = document.querySelector(".semantic-hit-active")?.textContent || "";
    return out;
  });
  check("pytanie idzie do AI razem z fragmentem [1] i zasadą „tylko z fragmentów”", answered.promptHasFragment && answered.promptHasRule, JSON.stringify(answered));
  check("odnośnik [1] jako przycisk", answered.cite === "1" && !answered.warn, JSON.stringify(answered));
  check("klik w odnośnik zaznacza akapit w dokumencie", /wypowiedzieć umowę/.test(answered.highlighted), answered.highlighted);
  check("pod odpowiedzią lista fragmentów", answered.sources >= 1, String(answered.sources));

  // 3) odpowiedź bez odnośnika → ostrzeżenie; „Brak w dokumencie” → bez ostrzeżenia
  const warn = await page.evaluate(async () => {
    window.__localAiMock = () => "Najemca ma 30 dni.";
    await docAskRun();
    const a = !!document.querySelector("#docAskAnswer .doc-ask-warn");
    window.__localAiMock = () => "Brak w dokumencie.";
    await docAskRun();
    const b = !!document.querySelector("#docAskAnswer .doc-ask-warn");
    delete window.__localAiMock;
    return { a, b };
  });
  check("odpowiedź bez źródła oznaczona ostrzeżeniem", warn.a && !warn.b, JSON.stringify(warn));

  // 4) bez AI (okno asystenta zamknięte bez pobierania) → same fragmenty
  const noAi = await page.evaluate(async () => {
    const run = docAskRun();
    for (let i = 0; i < 40 && !document.querySelector(".lai-dialog[open]"); i++) await new Promise((r) => setTimeout(r, 50));
    const opened = !!document.querySelector(".lai-dialog[open]");
    document.querySelector(".lai-dialog")?.close();
    await run;
    return { opened, hint: document.querySelector("#docAskAnswer .hint")?.textContent, sources: document.querySelectorAll("#docAskAnswer .doc-ask-source").length };
  });
  check("bez modelu: okno zgody, potem same fragmenty", noAi.opened && noAi.sources >= 1 && /Bez asystenta/.test(noAi.hint || ""), JSON.stringify(noAi));

  assertNoErrors(errors, "doc-ask");
  await browser.close();
  let failed = 0;
  results.forEach((r) => { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` (${r.detail})`}`); if (!r.ok) failed++; });
  if (failed) process.exit(1);
  console.log(`\n✅ doc-ask: ${results.length}/${results.length} (${process.env.ENGINE || "chromium"})`);
}
run().catch((e) => { console.error(e); process.exit(1); });
