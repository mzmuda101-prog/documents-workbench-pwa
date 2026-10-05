// replace-word-playwright.js — zamiana słowa przez przeglądarkę (insertReplacementText):
// podpowiedź z paska klawiatury telefonu, poprawka pisowni z menu pod prawym przyciskiem,
// autokorekta.
//
//   node scripts/replace-word-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/replace-word-playwright.js (Safari / iPhone / iPad)
//
// Zgłoszenie Mateusza 2026-10-05: kliknięcie podpowiedzi (telefon) albo poprawki pisowni
// (komputer, prawy przycisk) USUWAŁO słowo i zostawiała się dziura — nowe słowo przychodzi
// w e.dataTransfer, a nie w e.data, a zakres słowa w e.getTargetRanges().
// Test wysyła takie zdarzenie jak przeglądarka (bez działania domyślnego — sprawdza naszą obsługę):
// (a) kursor gdzie indziej, zakres = słowo (iPhone), (b) słowo zaznaczone (komputer),
// (c) słowo w formacie (pogrubione, inny rozmiar) — nowe słowo w tym samym formacie,
// (d) zapisany plik = podgląd, (e) Cofnij przywraca stare słowo, (f) bez zakresu i bez
// zaznaczenia nic nie znika (zamianę robi wtedy przeglądarka).
// WebKit nie przenosi dataTransfer ani targetRanges w sztucznym InputEvent (ograniczenie
// konstruktora) — tam nowe słowo idzie w data, a zakres to zaznaczenie (wariant „komputer”).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function buildDocx() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const r = (t, rpr = "") => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ""}<w:t xml:space="preserve">${t}</w:t></w:r>`;
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>
<w:p>${r("Pierwszy akapit ma bład w środku zdania.")}</w:p>
<w:p>${r("Drugi akapit: ")}${r("pogrubine", "<w:b/><w:sz w:val=\"28\"/>")}${r(" słowo z literówką.", "<w:sz w:val=\"28\"/>")}</w:p>
<w:p>${r("Trzeci akapit z kolejnym błedem do poprawki.")}</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: "nodebuffer" });
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "literowki.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await buildDocx() });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForFunction(() => !readOnlyMode && document.querySelector(".docx-edit-root"), null, { timeout: 10000 });
  await sleep(500);

  // zdarzenie jak z przeglądarki: zakres słowa + nowe słowo w dataTransfer (e.data puste)
  const replace = (wrong, right, mode, noTarget = false) => page.evaluate(({ wrong, right, mode, noTarget, webkit }) => {
    const root = docEditRoot();
    root.focus({ preventScroll: true });
    const p = [...document.querySelectorAll(".docx-preview-host p")].find((x) => x.textContent.includes(wrong));
    if (!p) return { missing: true, all: [...document.querySelectorAll(".docx-preview-host p")].map((x) => x.textContent).join(" | ") };
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let node = null;
    for (let t = w.nextNode(); t; t = w.nextNode()) if (t.data.includes(wrong)) { node = t; break; }
    const at = node.data.indexOf(wrong);
    const sel = getSelection();
    const r = document.createRange();
    if (mode === "selected") { r.setStart(node, at); r.setEnd(node, at + wrong.length); } // komputer: słowo zaznaczone
    else { r.selectNodeContents(p); r.collapse(false); } // telefon: kursor na końcu akapitu
    sel.removeAllRanges(); sel.addRange(r);
    const dt = new DataTransfer();
    dt.setData("text/plain", right);
    const target = new StaticRange({ startContainer: node, startOffset: at, endContainer: node, endOffset: at + wrong.length });
    const ev = new InputEvent("beforeinput", webkit
      ? { inputType: "insertReplacementText", data: right, bubbles: true, cancelable: true }
      : { inputType: "insertReplacementText", data: null, dataTransfer: dt, targetRanges: noTarget ? [] : [target], bubbles: true, cancelable: true });
    root.dispatchEvent(ev);
    return { prevented: ev.defaultPrevented, text: p.textContent };
  }, { wrong, right, mode, noTarget, webkit: ENGINE === "webkit" });

  // (f) bez zakresu i bez zaznaczenia — nic nie znika, zamianę zostawiamy przeglądarce
  const r0 = await replace("bład", "błąd", "caret-elsewhere", true);
  check("bez zakresu i bez zaznaczenia: słowo zostaje (zamiana dla przeglądarki), nic nie znika", !r0.prevented && r0.text === "Pierwszy akapit ma bład w środku zdania.", JSON.stringify(r0));
  const r1 = await replace("bład", "błąd", ENGINE === "webkit" ? "selected" : "caret-elsewhere");
  check(`${ENGINE === "webkit" ? "podpowiedź (słowo zaznaczone)" : "telefon: podpowiedź z paska klawiatury (kursor gdzie indziej)"} — słowo podmienione, nie usunięte`, r1.prevented && r1.text === "Pierwszy akapit ma błąd w środku zdania.", JSON.stringify(r1));
  const r2 = await replace("błedem", "błędem", "selected");
  check("komputer: poprawka pisowni z menu (słowo zaznaczone) — słowo podmienione, bez dziury", r2.prevented && r2.text === "Trzeci akapit z kolejnym błędem do poprawki.", JSON.stringify(r2));
  const r3 = await replace("pogrubine", "pogrubione", "selected");
  const fmt = await page.evaluate(() => {
    const p = [...document.querySelectorAll(".docx-preview-host p")].find((x) => x.textContent.includes("pogrubione"));
    if (!p) return null;
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = w.nextNode(); t; t = w.nextNode()) if (t.data.includes("pogrubione")) { const cs = getComputedStyle(t.parentElement); return { weight: cs.fontWeight, size: cs.fontSize }; }
    return null;
  });
  check("słowo w formacie (pogrubione, 14 pt): nowe słowo w tym samym formacie", r3.text === "Drugi akapit: pogrubione słowo z literówką." && fmt && +fmt.weight >= 600 && Math.abs(parseFloat(fmt.size) - 18.67) < 0.2, JSON.stringify({ r3, fmt }));
  // kursor za podmienionym słowem (dalsze pisanie trafia dalej, nie do środka)
  await page.keyboard.type("!");
  const afterType = await page.evaluate(() => [...document.querySelectorAll(".docx-preview-host p")].find((x) => x.textContent.includes("pogrubione"))?.textContent);
  check("kursor za podmienionym słowem", afterType === "Drugi akapit: pogrubione! słowo z literówką.", afterType);
  await page.keyboard.press("Backspace");
  await sleep(300);

  const saved = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const xml = await z.file("word/document.xml").async("string");
    const paras = (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).map((p) => ({ text: (p.match(/<w:t[^>]*>[^<]*/g) || []).map((t) => t.replace(/<w:t[^>]*>/, "")).join(""), xml: p }));
    return paras;
  });
  const p2 = saved.find((x) => x.text.includes("pogrubione"));
  check("zapis: wszystkie trzy poprawki w pliku", saved.map((x) => x.text).join(" | ") === "Pierwszy akapit ma błąd w środku zdania. | Drugi akapit: pogrubione słowo z literówką. | Trzeci akapit z kolejnym błędem do poprawki.", saved.map((x) => x.text).join(" | "));
  check("zapis: „pogrubione” pogrubione i 14 pt w pliku", /<w:b\b[^>]*\/>[\s\S]*?<w:sz w:val="28"\/>[\s\S]*?pogrubione<\/w:t>/.test(p2?.xml || ""), p2?.xml.slice(0, 300));

  // Cofnij — stare słowo wraca
  await sleep(1700);
  await page.evaluate(() => dwbUndo.undo());
  await sleep(500);
  const undone = await page.evaluate(() => [...document.querySelectorAll(".docx-preview-host p")].map((p) => p.textContent).join(" | "));
  check("Cofnij przywraca słowo sprzed poprawki", /pogrubine /.test(undone) || /błedem/.test(undone), undone);

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  let fail = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail})`}`);
    if (!r.ok) fail++;
  }
  console.log(`\n[${ENGINE}] ${fail ? `${fail} z ${results.length} nie przeszło` : `${results.length}/${results.length} OK`}`);
  process.exit(fail ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
