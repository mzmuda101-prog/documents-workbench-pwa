// compare-describe-playwright.js — „Opis zmian” w Porównaj wersje: liczby/kwoty/terminy na
// początku, zmiany słów, nowe akapity, kopiowanie, opis AI z kontrolą liczb (atrapa modelu).
const { createTestPage, bootApp, loadDocxFile, assertNoErrors } = require("./docx-test-helpers");

const results = [];
const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail });

async function run() {
  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await loadDocxFile(page, "docs/samples/headings-sample.docx");
  await page.evaluate(async () => {
    setSidebarOpen(true);
    document.getElementById("panel-compare").open = true;
    await ensureLazyFeature("compare");
    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    const idx = texts.map((t, i) => (t && t.trim() ? i : -1)).filter((i) => i >= 0).slice(1, 4);
    const older = [
      "§ 3. Termin płatności wynosi 14 dni od doręczenia faktury.",
      "§ 5. Wynagrodzenie wynosi 4 500 zł brutto miesięcznie.",
      "Strony ustalają, że Wykonawca wykona dzieło osobiście.",
    ];
    const newer = [
      "§ 3. Termin płatności wynosi 30 dni od doręczenia faktury.",
      "§ 5. Wynagrodzenie wynosi 5 000 zł brutto miesięcznie.",
      "Strony ustalają, że Zleceniobiorca wykona dzieło osobiście.",
    ];
    await applyDocumentEdit({ op: "paragraphBatch", items: idx.map((index, k) => ({ index, text: older[k] })) });
    const olderBytes = originalFileBytes;
    await applyDocumentEdit({ op: "paragraphBatch", items: idx.map((index, k) => ({ index, text: newer[k] })) });
    const dt = new DataTransfer();
    dt.items.add(new File([olderBytes], "starsza.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
    document.getElementById("compareFile").files = dt.files;
  });
  await page.click("#compareRunBtn");
  await page.waitForFunction(() => /Porównano/.test(document.getElementById("compareStatus").textContent), null, { timeout: 10000 });

  const view = await page.evaluate(() => ({
    hidden: document.getElementById("compareDescribe").hidden,
    numbers: [...document.querySelectorAll(".compare-describe-list.is-numbers li")].map((li) => li.textContent),
    all: document.getElementById("compareDescribeBody").textContent,
    text: compareDescribeText,
  }));
  check("sekcja „Opis zmian” widoczna", !view.hidden, JSON.stringify(view));
  check("termin: § 3 „14 dni” → „30 dni”", view.numbers.includes("§ 3: „14 dni” → „30 dni”"), JSON.stringify(view.numbers));
  check("kwota: § 5 „4 500 zł” → „5 000 zł”", view.numbers.includes("§ 5: „4 500 zł” → „5 000 zł”"), JSON.stringify(view.numbers));
  check("zmiana słowa: „Wykonawca” → „Zleceniobiorca”", /„Wykonawca” → „Zleceniobiorca”/.test(view.all), view.all);
  check("liczby przed innymi zmianami", view.text.indexOf("30 dni") < view.text.indexOf("Zleceniobiorca"), view.text);
  check("statystyka w pierwszym wierszu", /^Zmieniono akapitów: 3/.test(view.text), view.text.split("\n")[0]);

  // kierunek: nowszy = drugi plik → para odwrotna
  await page.click('#compareNewer button[data-newer="reference"]');
  const flipped = await page.evaluate(() => [...document.querySelectorAll(".compare-describe-list.is-numbers li")].map((li) => li.textContent));
  check("odwrócony kierunek odwraca pary", flipped.includes("§ 3: „30 dni” → „14 dni”"), JSON.stringify(flipped));
  await page.click('#compareNewer button[data-newer="current"]');

  // AI (atrapa): opis z liczbami z listy przechodzi, z wymyśloną liczbą — odrzucony
  const ai = await page.evaluate(async () => {
    const out = {};
    window.__localAiMock = () => "Termin płatności wydłużono z 14 do 30 dni, a wynagrodzenie wzrosło z 4 500 zł do 5 000 zł.";
    document.getElementById("compareDescribeAiBtn").click();
    await new Promise((r) => setTimeout(r, 300));
    out.ok = document.querySelector(".compare-describe-ai-text")?.textContent;
    out.okRejected = document.querySelector(".compare-describe-ai-text")?.classList.contains("is-rejected");
    window.__localAiMock = () => "Wynagrodzenie wzrosło do 6 000 zł.";
    document.getElementById("compareDescribeAiBtn").click();
    await new Promise((r) => setTimeout(r, 300));
    out.bad = document.querySelector(".compare-describe-ai-text")?.textContent;
    out.badRejected = document.querySelector(".compare-describe-ai-text")?.classList.contains("is-rejected");
    out.badge = document.querySelector("#compareDescribeAi .ai-badge")?.textContent;
    delete window.__localAiMock;
    return out;
  });
  check("opis AI z liczbami z listy — pokazany", !ai.okRejected && /30 dni/.test(ai.ok || ""), JSON.stringify(ai));
  check("opis AI z wymyśloną kwotą — odrzucony", ai.badRejected && /6000/.test(ai.bad || ""), JSON.stringify(ai));
  check("wynik AI podpisany plakietką", /AI/.test(ai.badge || ""), ai.badge);

  // kontrola liczb (czysta funkcja)
  const fn = await page.evaluate(() => [
    dwbDescribeCheckAi("Kwota 4 500,00 zł i 30 dni.", "„4 500 zł” → „5 000 zł”; „14 dni” → „30 dni”").ok,
    dwbDescribeCheckAi("Dodano 7 dni.", "„14 dni” → „30 dni”").ok,
  ]);
  check("dwbDescribeCheckAi: format liczb niezależny od zapisu, obce liczby wykryte", fn[0] === true && fn[1] === false, JSON.stringify(fn));

  assertNoErrors(errors, "compare-describe");
  await browser.close();
  let failed = 0;
  results.forEach((r) => { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` (${r.detail})`}`); if (!r.ok) failed++; });
  if (failed) process.exit(1);
  console.log(`\n✅ compare-describe: ${results.length}/${results.length}`);
}
run().catch((e) => { console.error(e); process.exit(1); });
