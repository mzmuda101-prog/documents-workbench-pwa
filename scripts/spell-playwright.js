// spell-playwright.js — ortografia w „Korekcie” (Hunspell offline): skan, podpowiedzi, zamiana
// z zachowaniem wielkiej litery, „Pomiń”, „Do słownika”, pomijanie skrótów / adresów / iPhone.
// ENGINE=webkit — to samo w silniku Safari (WebAssembly + module worker na iPhonie).

const { createTestPage, bootApp, loadBuiltinSample, assertNoErrors } = require("./docx-test-helpers");

const BAD = "Prosze przesłać dokumętów do piątku. Dokumętów brak, RODO i PESEL są ok, iPhone też, www.przyklad.pl i jan@firma.pl też, COVID-19 też, e-mailem wysłane.";

async function run() {
  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await loadBuiltinSample(page);

  const result = await page.evaluate(async (bad) => {
    const fail = (step, msg) => ({ ok: false, step, msg });
    await ensureDocLibs(false);
    const texts0 = await extractParagraphTextsFromDocx(originalFileBytes);
    const target = texts0.findIndex((t) => t && t.trim().length > 0);
    await applyDocumentEdit({ op: "paragraphBatch", items: [{ index: target, text: bad }] });

    // panel Korekta: leniwe wczytanie modułów jak po kliknięciu
    const panel = document.getElementById("panel-grammar");
    panel.open = true;
    panel.dispatchEvent(new Event("toggle"));
    const until = async (fn, ms = 20000) => {
      const t0 = Date.now();
      while (Date.now() - t0 < ms) { const v = fn(); if (v) return v; await new Promise((r) => setTimeout(r, 50)); }
      return null;
    };
    if (!(await until(() => typeof dwbSpell !== "undefined" && typeof runGrammarScan === "function"))) return fail("lazy", "brak modułów");
    if (!document.getElementById("grammarSpell")?.checked) return fail("default", "ortografia domyślnie wyłączona");

    const t0 = performance.now();
    await runGrammarScan();
    const scanMs = Math.round(performance.now() - t0);
    const words = (grammarScan.spelling || []).map((e) => e.word);
    // „Prosze” i „dokumętów”/„Dokumętów” (osobno: inna pisownia = osobny wpis)
    for (const w of ["Prosze", "dokumętów", "Dokumętów"]) if (!words.includes(w)) return fail("found", `brak „${w}” w ${JSON.stringify(words)}`);
    for (const w of ["RODO", "PESEL", "iPhone", "www", "przyklad", "jan", "firma", "COVID", "e-mailem", "przesłać"]) {
      if (words.includes(w)) return fail("skip", `„${w}” nie powinien być błędem (${JSON.stringify(words)})`);
    }
    if (!(await until(() => grammarScan.spelling.every((e) => e.suggestions != null)))) return fail("suggest", "podpowiedzi nie przyszły");
    const entry = grammarScan.spelling.find((e) => e.word === "dokumętów");
    if (!entry.suggestions.includes("dokumentów")) return fail("suggest", `podpowiedzi: ${entry.suggestions}`);

    // UI: grupa „Ortografia” z przyciskiem podpowiedzi; klik zamienia
    const row = [...document.querySelectorAll(".spell-row")].find((r) => r.querySelector(".spell-word")?.textContent.startsWith("Dokumętów"));
    const btn = row && [...row.querySelectorAll(".spell-suggestion")].find((b) => b.textContent === "Dokumentów");
    if (!btn) return fail("ui", "brak przycisku „Dokumentów” (wielka litera jak w oryginale)");
    btn.click();
    if (!(await until(() => !grammarScan?.spelling?.some((e) => e.word === "Dokumętów")))) return fail("apply", "wyraz dalej na liście");
    const after = (await extractParagraphTextsFromDocx(originalFileBytes))[target];
    if (!after.includes("Dokumentów brak")) return fail("apply", `tekst po zamianie: ${after}`);
    if (!after.includes("dokumętów do piątku")) return fail("apply", "zamieniono też małą pisownię (inny wpis)");

    // Cofnij przywraca
    {
      await dwbUndo.undo();
      const back = (await extractParagraphTextsFromDocx(originalFileBytes))[target];
      if (!back.includes("Dokumętów brak")) return fail("undo", back);
      await runGrammarScan();
    }

    // Pomiń / Do słownika
    const ignoreRow = [...document.querySelectorAll(".spell-row")].find((r) => r.querySelector(".spell-word")?.textContent.startsWith("Prosze"));
    ignoreRow.querySelector(".spell-add").click();
    if (!dwbSpell.userWords().has("prosze")) return fail("dict", "brak w słowniku użytkownika");
    await runGrammarScan();
    if (grammarScan.spelling.some((e) => e.word === "Prosze")) return fail("dict", "wyraz ze słownika dalej zgłaszany");
    localStorage.removeItem("dwb.userDictionary.v1");

    // licznik w nagłówku panelu uwzględnia ortografię
    const status = document.getElementById("grammarStatus").textContent;
    if (!/Ortografia/.test(status)) return fail("status", status);
    return { ok: true, scanMs, words };
  }, BAD);

  assertNoErrors(errors);
  await browser.close();
  if (!result.ok) {
    console.error(`❌ spell [${result.step}]: ${result.msg}`);
    process.exit(1);
  }
  console.log(`✅ spell: ${result.words.length} wyrazów do sprawdzenia, skan ${result.scanMs} ms (${process.env.ENGINE || "chromium"})`);
}

run().catch((err) => {
  console.error(err);
  process.exit(1);
});
