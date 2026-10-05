// para-style-playwright.js — style akapitu: wyjście z ramki/cytatu i styl „Ramka” (2026-10-05).
//
// Zgłoszenie: Enter na końcu ramki „Spróbuj:” w przewodniku → nowy akapit też w ramce i nie dało
// się z niej wyjść; lista stylów pokazywała „Normalny”, więc wybranie „Normalny” nic nie robiło.
// Teraz: lista pokazuje nazwę stylu pliku („Wskazówka”), Enter/Backspace w PUSTYM akapicie w stylu
// → Normalny (jak wyjście z listy), a ramkę można wstawić samemu (styl „Ramka”).
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(500));
const keyOfCaret = (page) => page.evaluate(() => composeStyleKeyOf(docCaretParagraph(document.activeElement) || lastDocCaret?.p));
const styleSelVal = (page) => page.evaluate(() => ({ v: document.getElementById("fmtParaStyle").value, label: document.getElementById("fmtParaStyle").selectedOptions[0]?.textContent }));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => loadSampleDocument("przewodnik"));
  await page.waitForSelector(".docx-preview-host p", { timeout: 30000 });
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await idle(page);

  // kursor na końcu ostatniej ramki „Spróbuj:”
  const tipEnd = () => page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => /^Spróbuj:/.test(p.textContent) && !p.dataset.lock);
    const p = ps[ps.length - 1];
    p.scrollIntoView({ block: "center" });
    placeCaret(p, p.textContent.length);
    return resolveParaIndex(p);
  });
  let idx = await tipEnd();
  await page.waitForTimeout(300);
  let sel = await styleSelVal(page);
  check("kursor w ramce: lista stylów pokazuje „Wskazówka” (nie „Normalny”)", sel.label === "Wskazówka" && sel.v === "custom:Wskazówka", JSON.stringify(sel));
  await page.keyboard.press("Enter");
  await idle(page);
  check("Enter na końcu ramki: nowy akapit też w ramce (jak w Wordzie)", (await keyOfCaret(page)) === "custom:Wskazówka", await keyOfCaret(page));
  await page.keyboard.press("Enter");
  await idle(page);
  check("drugi Enter w PUSTEJ ramce: akapit wraca do Normalnego (wyjście)", (await keyOfCaret(page)) === "normal", await keyOfCaret(page));
  await page.keyboard.type("Zwykły tekst po ramce");
  await idle(page);
  let zip = await page.evaluate(async () => { const z = await JSZip.loadAsync(await buildDocumentForSave()); return z.file("word/document.xml").async("string"); });
  const after = zip.match(/<w:p[ >](?:(?!<\/w:p>).)*Zwykły tekst po ramce(?:(?!<\/w:p>).)*<\/w:p>/s)?.[0] || "";
  check("plik: tekst po ramce w zwykłym akapicie (bez stylu Tip)", after && !/w:pStyle w:val="Tip"/.test(after), after.slice(0, 200));

  // Backspace w pustej ramce → Normalny
  idx = await tipEnd();
  await page.keyboard.press("Enter");
  await idle(page);
  await page.keyboard.press("Backspace");
  await idle(page);
  check("Backspace w PUSTEJ ramce: akapit wraca do Normalnego", (await keyOfCaret(page)) === "normal", await keyOfCaret(page));

  // wybranie „Normalny” z listy w ramce zdejmuje ramkę
  idx = await tipEnd();
  await page.waitForTimeout(250);
  await page.selectOption("#fmtParaStyle", "normal");
  await idle(page);
  check("w ramce: wybranie „Normalny” z listy stylów zdejmuje ramkę", (await keyOfCaret(page)) === "normal", await keyOfCaret(page));

  // ── nowy dokument: styl „Ramka” ──────────────────────────────────────────────
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);
  await page.keyboard.type("Ważne: termin mija w piątek.");
  await page.selectOption("#fmtParaStyle", "callout");
  await idle(page);
  const box = await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0]; const cs = getComputedStyle(p); return { key: composeStyleKeyOf(p), border: cs.borderLeftStyle, bg: cs.backgroundColor }; });
  check("styl „Ramka”: pasek z lewej i tło w podglądzie", box.key === "callout" && box.border === "solid" && box.bg !== "rgba(0, 0, 0, 0)", JSON.stringify(box));
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0]; placeCaret(p, p.textContent.length); });
  await page.keyboard.press("Enter");
  await idle(page);
  check("Enter na końcu ramki z tekstem (styl „Ramka”): dalej zwykły tekst", (await keyOfCaret(page)) === "normal", await keyOfCaret(page));
  const files = await page.evaluate(async () => { const z = await JSZip.loadAsync(await buildDocumentForSave()); return { doc: await z.file("word/document.xml").async("string"), styles: await z.file("word/styles.xml").async("string") }; });
  check("plik: akapit w stylu DWBRamka, styl „Ramka” z obramowaniem i tłem w styles.xml", /w:pStyle w:val="DWBRamka"/.test(files.doc) && /w:styleId="DWBRamka"><w:name w:val="Ramka"\/>[\s\S]*?<w:pBdr><w:left w:val="single"[\s\S]*?w:fill="EEF3FB"/.test(files.styles), files.styles.match(/DWBRamka[\s\S]{0,200}/)?.[0]);

  await browser.close();
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Style akapitu [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
