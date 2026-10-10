// grammar-playwright.js — skan typografii, apply all, weryfikacja XML.

const path = require("path");
const {
  BUILTIN_SAMPLE,
  fixtureExists,
  createTestPage,
  bootApp,
  loadDocxFile,
  loadBuiltinSample,
  assertNoErrors,
} = require("./docx-test-helpers");

const DOCX_ARG = process.argv[2];
const BAD_TEXT = "Test słowo  słowo , tekst i.";

async function run() {
  const filePath = DOCX_ARG ? path.resolve(DOCX_ARG) : BUILTIN_SAMPLE;
  if (!fixtureExists(filePath)) {
    console.warn(`⚠️  Pominięto — brak pliku: ${filePath}`);
    process.exit(0);
  }

  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  if (DOCX_ARG) await loadDocxFile(page, filePath, 90000);
  else await loadBuiltinSample(page);

  const result = await page.evaluate(async (badText) => {
    await ensureDocLibs(false);

    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    const targetIndex = texts.findIndex((t) => t && t.trim().length > 0);
    if (targetIndex < 0) return { ok: false, step: "fixture", msg: "Brak akapitu w dokumencie" };

    await applyDocumentEdit({
      op: "paragraphBatch",
      items: [{ index: targetIndex, text: badText }],
    });

    const scan = await scanDocument(originalFileBytes, { lang: "pl", nbspPl: false });
    if (scan.hits.length < 2) {
      return { ok: false, step: "scan", msg: `Oczekiwano ≥2 trafień, jest ${scan.hits.length}` };
    }

    const ruleIds = Object.keys(scan.byRule);
    if (!ruleIds.includes("double-space") || !ruleIds.includes("space-before-punct")) {
      return { ok: false, step: "rules", msg: `Brak oczekiwanych reguł: ${ruleIds.join(", ")}` };
    }

    const inline = scanParagraph(badText, "pl", { nbspPl: false });
    if (inline.length < 2) {
      return { ok: false, step: "scanParagraph", msg: "scanParagraph — za mało trafień" };
    }

    const items = buildGrammarBatchItems(
      await extractParagraphTextsFromDocx(originalFileBytes),
      scan.hits,
      { lang: "pl", nbspPl: false }
    );
    if (!items.length) return { ok: false, step: "batch", msg: "Brak pozycji batch" };

    const count = await applyDocumentEdit({ op: "paragraphBatch", items });
    if (!count) return { ok: false, step: "apply", msg: "apply all — 0 zmian" };

    const after = await extractParagraphTextsFromDocx(originalFileBytes);
    const fixed = after[targetIndex] || "";
    if (/ {2,}/.test(fixed)) {
      return { ok: false, step: "verify-space", msg: `Podwójne spacje pozostały: ${fixed}` };
    }
    if (/\s+[,;:.!?]/.test(fixed)) {
      return { ok: false, step: "verify-punct", msg: `Spacja przed interpunkcją: ${fixed}` };
    }

    // reguły na przypadkach, które dawniej psuły tekst
    const fix = (t, o = {}) => fixParagraphWithRules(t, getEnabledGrammarRules({ lang: "pl", ...o }), o);
    const cases = [
      ["ACME sp. z o.o. i Beta sp. j. oraz np. w lipcu, m.in. w domu, ul. zielona 5, art. 5 ust. 2 pkt. a.", "ACME sp. z o.o. i Beta sp. j. oraz np. w lipcu, m.in. w domu, ul. zielona 5, art. 5 ust. 2 pkt. a."],
      ["Koniec zdania. nowe zdanie", "Koniec zdania. Nowe zdanie"],
      ["J. kowalski podpisał", "J. kowalski podpisał"],
      ["a) pierwszy punkt listy", "a) pierwszy punkt listy"],
      ["To że że zostało wpisane, jest błędem.", "To że zostało wpisane, jest błędem."],
      ['Powiedział "tak" i "nie".', "Powiedział „tak” i „nie”."],
      ["Kupiłem jabłka i.", "Kupiłem jabłka i."], // dawna „sierota i” przestawiała słowa
      // brak spacji po interpunkcji — adresy, pliki, liczby i skróty zostają
      ["To koniec.nowe zdanie", "To koniec. Nowe zdanie"],
      ["Zrób to tak,jak trzeba", "Zrób to tak, jak trzeba"],
      ["Wejdź na www.firma.pl albo napisz: jan.kowalski@firma.pl", "Wejdź na www.firma.pl albo napisz: jan.kowalski@firma.pl"],
      ["Plik raport.docx i strona example.com/abc", "Plik raport.docx i strona example.com/abc"],
      ["Liczby 3,5 i 2.7 bez zmian", "Liczby 3,5 i 2.7 bez zmian"],
    ];
    for (const [input, want] of cases) {
      const got = fix(input);
      if (got !== want) return { ok: false, step: "rules-cases", msg: `${JSON.stringify(input)} → ${JSON.stringify(got)} (chciano ${JSON.stringify(want)})` };
    }
    const nb = fix("Cała sprawa w domu i u nas", { nbspPl: true });
    if (nb !== "Cała sprawa w\u00A0domu i\u00A0u\u00A0nas") return { ok: false, step: "nbsp", msg: JSON.stringify(nb) };
    const enRules = getEnabledGrammarRules({ lang: "en" });
    const repeatedEn = fixParagraphWithRules("The the draft is ready.", enRules, { lang: "en" });
    if (repeatedEn !== "The draft is ready.") return { ok: false, step: "repeated-en", msg: repeatedEn };
    const orphanLeft = 0;

    return {
      ok: true,
      hits: scan.hits.length,
      rules: ruleIds.length,
      fixed,
      orphanLeft,
    };
  }, BAD_TEXT);

  if (!result.ok) throw new Error(`grammar failed at ${result.step}: ${result.msg}`);

  // Panel: podgląd poprawki = kontekst + skreślone → wstawione, spacje widoczne („·”)
  const preview = await page.evaluate(async (badText) => {
    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    const idx = texts.findIndex((t) => t && t.trim().length > 0);
    await applyDocumentEdit({ op: "paragraphBatch", items: [{ index: idx, text: badText }] });
    setSidebarOpen(true);
    document.getElementById("panel-grammar").open = true;
    await ensureLazyFeature("grammar");
    document.getElementById("grammarScanBtn").click();
    for (let i = 0; i < 40 && !document.querySelector(".grammar-hit-item"); i++) await new Promise((r) => setTimeout(r, 100));
    const dbl = [...document.querySelectorAll(".grammar-hit-item")].find((b) => b.querySelector(".gr-del .gr-space"));
    return {
      items: document.querySelectorAll(".grammar-hit-item").length,
      del: document.querySelectorAll(".grammar-hit-item .gr-del").length,
      ins: document.querySelectorAll(".grammar-hit-item .gr-ins").length,
      doubleSpaceShown: dbl ? dbl.querySelector(".gr-del").textContent : null,
    };
  }, BAD_TEXT);
  if (!preview.items || !preview.del || !preview.ins || preview.doubleSpaceShown !== "··") {
    throw new Error(`grammar preview: ${JSON.stringify(preview)}`);
  }
  assertNoErrors(errors, "grammar");
  await browser.close();
  console.log(
    `✅  grammar-playwright passed (hits=${result.hits}, rules=${result.rules}, orphanLeft=${result.orphanLeft})`
  );
  console.log(`    fixed: ${result.fixed}`);
}

run().catch((err) => {
  console.error("❌  grammar-playwright failed:", err.message || err);
  process.exit(1);
});
