// Local Word-style comparison: content changes, no mutation of either source package.
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
    const data = await fetch("docs/samples/headings-sample.docx").then((r) => r.arrayBuffer());
    const zip = await JSZip.loadAsync(data);
    let xml = await zip.file("word/document.xml").async("string");
    xml = xml.replace("Dane Wynajmującego", "Dane Wynajmującego po redakcji");
    xml = xml.replace("</w:body>", '<w:p><w:r><w:t>Dodany akapit porównawczy.</w:t></w:r></w:p></w:body>');
    zip.file("word/document.xml", xml);
    const ref = await zip.generateAsync({ type: "blob" });
    const dt = new DataTransfer();
    dt.items.add(new File([ref], "druga-wersja.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
    document.getElementById("compareFile").files = dt.files;
  });
  await page.click("#compareRunBtn");
  await page.waitForFunction(() => /Porównano/.test(document.getElementById("compareStatus").textContent), null, { timeout: 10000 });
  const view = await page.evaluate(() => ({
    status: document.getElementById("compareStatus").textContent,
    changed: document.querySelector(".compare-cell.changed strong")?.textContent,
    added: document.querySelector(".compare-cell.added strong")?.textContent,
    removed: document.querySelector(".compare-cell.removed strong")?.textContent,
    rows: document.querySelectorAll(".compare-item").length,
    before: originalFileBytes.byteLength,
  }));
  check("panel shows changed paragraph", Number(view.changed) >= 1, JSON.stringify(view));
  check("panel shows paragraph missing from open document", Number(view.removed || 0) >= 1 && view.rows >= 2, JSON.stringify(view));
  check("comparison leaves open document unchanged", await page.evaluate(() => !hasUnsavedChanges && pendingDocEdits.length === 0));
  check("word-level diff marks edited words", await page.evaluate(() => document.querySelectorAll(".compare-word.compare-added, .compare-word.compare-removed").length >= 2));
  // Kierunek (jak w Wordzie: starsza → nowsza). Domyślnie nowszy = otwarty: akapit, którego nie ma
  // w otwartym, jest „usunięty”; po wyborze „Drugi plik” — „dodany”.
  const kindOf = (text) => page.evaluate((text) => [...document.querySelectorAll(".compare-item")].find((el) => el.textContent.includes(text))?.className.match(/compare-item-(\w+)/)?.[1], text);
  check("domyślnie nowszy = otwarty: akapit tylko w drugim pliku jest usunięty", await kindOf("Dodany akapit porównawczy") === "removed", await kindOf("Dodany akapit porównawczy"));
  await page.click('#compareNewer button[data-newer="reference"]');
  const flipped = await page.evaluate(() => ({
    added: document.querySelector(".compare-cell.added strong")?.textContent,
    removed: document.querySelector(".compare-cell.removed strong")?.textContent,
    checked: document.querySelector('#compareNewer button[data-newer="reference"]').getAttribute("aria-checked"),
    newWordAdded: [...document.querySelectorAll(".compare-item-changed .compare-word.compare-added")].some((m) => /redakcji/.test(m.textContent)),
  }));
  check("„Nowsza: Drugi plik” — ten akapit staje się dodany (bez ponownego porównania)", await kindOf("Dodany akapit porównawczy") === "added" && flipped.checked === "true", JSON.stringify(flipped));
  check("zmieniony akapit: słowa z nowszej wersji jako dodane", flipped.newWordAdded, JSON.stringify(flipped));
  await page.click('#compareNewer button[data-newer="current"]');
  check("powrót do „Otwarty dokument” przywraca kierunek", await kindOf("Dodany akapit porównawczy") === "removed");
  // plakietka rodzaju zmiany w jednym wierszu (dawniej pionowa łamała się na litery)
  const badges = await page.evaluate(() => [...document.querySelectorAll(".compare-kind")].slice(0, 6).map((b) => {
    const r = b.getBoundingClientRect(); return { h: Math.round(r.height), w: Math.round(r.width), wm: getComputedStyle(b).writingMode };
  }));
  check("plakietki rodzaju zmiany poziome, w jednym wierszu", badges.length && badges.every((b) => b.wm === "horizontal-tb" && b.h <= 18 && b.w > b.h), JSON.stringify(badges));
  assertNoErrors(errors, "compare");
  await browser.close();
  let failed = 0;
  results.forEach((r) => { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` (${r.detail})`}`); if (!r.ok) failed++; });
  if (failed) process.exit(1);
  console.log(`\n✅ comparison: ${results.length}/${results.length}`);
}
run().catch((e) => { console.error(e); process.exit(1); });
