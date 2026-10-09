// format-state-playwright.js — B / I / U bez zaznaczenia i „wciśnięte” przyciski (2026-10-05).
//
// Zgłoszenie Mateusza: za pogrubionym słowem nie da się dopisać tekstu bez pogrubienia (nawet bez
// spacji) — Ctrl/⌘+B „nic nie robi”, wszystko dalej jest pogrubione. Przyczyna: przy samym kursorze
// Ctrl/⌘+B szło przez execCommand, a nasze wstawianie liter brało format z sąsiedniego znaku.
// Teraz B/I/U bez zaznaczenia przełączają format DALSZEGO pisania (jak w Wordzie), a wyłączenie
// wychodzi poza fragment (też podkreślenie). Przyciski B/I/U pokazują stan (is-on, aria-pressed).
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

// fragmenty pierwszego akapitu z tekstem w zapisanym pliku: [tekst, b, u]
const savedRuns = (page, needle) => page.evaluate(async (needle) => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  const doc = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const p = [...doc.getElementsByTagNameNS(W, "p")].find((x) => x.textContent.includes(needle));
  if (!p) return [];
  const on = (rPr, tag) => { const e = rPr?.getElementsByTagNameNS(W, tag)[0]; return !!e && !["0", "false", "none"].includes(e.getAttributeNS(W, "val") || e.getAttribute("w:val") || ""); };
  const out = [];
  for (const r of p.getElementsByTagNameNS(W, "r")) {
    const t = [...r.getElementsByTagNameNS(W, "t")].map((x) => x.textContent).join("");
    if (!t) continue;
    const rPr = r.getElementsByTagNameNS(W, "rPr")[0];
    const b = on(rPr, "b"), u = on(rPr, "u");
    const last = out[out.length - 1];
    if (last && last[1] === b && last[2] === u) last[0] += t; else out.push([t, b, u]);
  }
  return out;
}, needle);
const btn = (page) => page.evaluate(() => ({ b: document.getElementById("fmtBold").classList.contains("is-on"), u: document.getElementById("fmtUnderline").classList.contains("is-on"), aria: document.getElementById("fmtBold").getAttribute("aria-pressed") }));

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
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);

  // ── 1. pisanie z przełączaniem: zwykły → Ctrl/⌘+B → gruby → Ctrl/⌘+B → dalej ─────────
  await page.keyboard.type("Zwykły ");
  check("przycisk B zgaszony w zwykłym tekście", !(await btn(page)).b, JSON.stringify(await btn(page)));
  await page.keyboard.press(`${MOD}+b`);
  await page.waitForTimeout(150);
  const afterOn = await btn(page);
  check("Ctrl/⌘+B bez zaznaczenia: przycisk B od razu „wciśnięty” (aria-pressed)", afterOn.b && afterOn.aria === "true", JSON.stringify(afterOn));
  await page.keyboard.type("gruby");
  await page.keyboard.press(`${MOD}+b`);
  await page.waitForTimeout(150);
  check("drugi Ctrl/⌘+B: przycisk B zgaszony", !(await btn(page)).b, JSON.stringify(await btn(page)));
  await page.keyboard.type("dalej");
  await page.waitForTimeout(300);
  let runs = await savedRuns(page, "Zwykły");
  const noise = await page.evaluate(async () => { const z = await JSZip.loadAsync(await buildDocumentForSave()); const d = await z.file("word/document.xml").async("string"); return (d.match(/<w:p[ >](?:(?!<\/w:p>).)*Zwykły(?:(?!<\/w:p>).)*<\/w:p>/s) || [""])[0].includes('w:b w:val="0"'); });
  check("zwykły akapit po Ctrl/⌘+B: bez zbędnego w:b w:val=\"0\" w pliku", !noise, String(noise));
  check("plik: „Zwykły ” zwykły, „gruby” pogrubiony, „dalej” zwykły (bez spacji)", JSON.stringify(runs) === JSON.stringify([["Zwykły ", false, false], ["gruby", true, false], ["dalej", false, false]]), JSON.stringify(runs));

  // ── 2. zgłoszenie: kursor na końcu ISTNIEJĄCEGO pogrubionego słowa → Ctrl/⌘+B → pisanie ──────
  await page.keyboard.press("Enter");
  await page.keyboard.type("Przed ");
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.type("POGRUBIONE");
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.type(" po.");
  await idle(page);
  // kliknięcie w środek pogrubionego słowa → B świeci; kursor na jego końcu
  const pos = await page.evaluate(() => {
    const p = [...document.querySelectorAll(".docx-preview-host p.docx-editable-p")].find((x) => x.textContent.includes("POGRUBIONE"));
    const r = document.createRange();
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const i = n.data.indexOf("POGRUB");
      if (i >= 0) { r.setStart(n, i + 3); r.setEnd(n, i + 4); const b = r.getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + b.height / 2 }; }
    }
    return null;
  });
  await page.mouse.click(pos.x, pos.y);
  await page.waitForTimeout(250);
  check("kursor w pogrubionym słowie: przycisk B „wciśnięty”", (await btn(page)).b, JSON.stringify(await btn(page)));
  await page.evaluate(() => {
    const p = [...document.querySelectorAll(".docx-preview-host p.docx-editable-p")].find((x) => x.textContent.includes("POGRUBIONE"));
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const i = n.data.indexOf("POGRUBIONE");
      if (i >= 0) { const r = document.createRange(); r.setStart(n, i + 10); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); return; }
    }
  });
  await page.waitForTimeout(200);
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.type("xyz");
  await page.waitForTimeout(300);
  runs = await savedRuns(page, "POGRUBIONE");
  check("za pogrubionym słowem, Ctrl/⌘+B, pisanie bez spacji: „xyz” NIE pogrubione", JSON.stringify(runs) === JSON.stringify([["Przed ", false, false], ["POGRUBIONE", true, false], ["xyz po.", false, false]]), JSON.stringify(runs));

  // ── 3. podkreślenie: Ctrl/⌘+U wyłącza je dla dalszego pisania (tekst poza fragmentem) ──────
  await page.evaluate(() => { const p = [...document.querySelectorAll(".docx-preview-host p.docx-editable-p")].find((x) => x.textContent.includes("POGRUBIONE")); placeCaret(p, p.textContent.length); });
  await page.keyboard.press("Enter");
  await page.keyboard.type("Tekst ");
  await page.keyboard.press(`${MOD}+u`);
  await page.keyboard.type("podkreślony");
  check("Ctrl/⌘+U: przycisk U „wciśnięty”", (await btn(page)).u, JSON.stringify(await btn(page)));
  await page.keyboard.press(`${MOD}+u`);
  await page.keyboard.type("koniec");
  await page.waitForTimeout(300);
  runs = await savedRuns(page, "podkreślony");
  check("podkreślenie wyłączone w trakcie pisania: „koniec” bez podkreślenia", JSON.stringify(runs) === JSON.stringify([["Tekst ", false, false], ["podkreślony", false, true], ["koniec", false, false]]), JSON.stringify(runs));

  // ── 3a. przekreślenie: przycisk z paska używa modelu runów, nie tylko CSS podglądu ───────
  await page.keyboard.press("Enter");
  await page.keyboard.type("Korekta: ");
  await page.click("#fmtStrike");
  await page.keyboard.type("skreślone");
  const strikeOn = await page.evaluate(() => document.getElementById("fmtStrike").getAttribute("aria-pressed") === "true");
  await page.click("#fmtStrike");
  await page.keyboard.type(" zostaje");
  await page.waitForTimeout(300);
  const strikeRuns = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const xml = await z.file("word/document.xml").async("string");
    const p = (xml.match(/<w:p[ >](?:(?!<\/w:p>).)*skreślone(?:(?!<\/w:p>).)*<\/w:p>/s) || [""])[0];
    return [...p.matchAll(/<w:r>(.*?)<\/w:r>/gs)].map((m) => [(m[1].match(/<w:t[^>]*>([^<]*)/) || [])[1], /<w:strike(?:\s|\/)/.test(m[1]) && !/w:val="0"/.test(m[1])]);
  });
  check("przekreślenie z paska: przycisk świeci, a .docx ma w:strike tylko na zaznaczonym fragmencie", strikeOn && JSON.stringify(strikeRuns) === JSON.stringify([["Korekta: ", false], ["skreślone", true], [" zostaje", false]]), JSON.stringify({ strikeOn, strikeRuns }));

  // ¶ jest wyłącznie widokiem: nie zapisuje danych, a po przeładowaniu nadal odpowiada przyciskowi.
  await page.click("#formatMarksBtn");
  check("¶ na pasku włącza znaki niedrukowalne bez dodania operacji do kolejki pliku", await page.evaluate(() => document.getElementById("formatMarksBtn").getAttribute("aria-pressed") === "true" && docCanvasEl.classList.contains("show-formatting-marks") && pendingDocEdits.length === 0));

  // ── 3b. kursywa: Ctrl/⌘+I za pochylonym słowem; w Cytacie (pochylony ze stylu) → w:i w:val="0" ──
  await page.evaluate(() => { const ps = document.querySelectorAll(".docx-preview-host p.docx-editable-p"); const p = ps[ps.length - 1]; placeCaret(p, p.textContent.length); });
  await page.keyboard.press("Enter");
  await page.keyboard.type("Zdanie ");
  await page.keyboard.press(`${MOD}+i`);
  await page.keyboard.type("pochyłe");
  const iOn = await page.evaluate(() => document.getElementById("fmtItalic").classList.contains("is-on"));
  await page.keyboard.press(`${MOD}+i`);
  const iOff = await page.evaluate(() => document.getElementById("fmtItalic").classList.contains("is-on"));
  await page.keyboard.type("proste");
  await page.waitForTimeout(300);
  const iRuns = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const doc = await z.file("word/document.xml").async("string");
    const p = (doc.match(/<w:p[ >](?:(?!<\/w:p>).)*pochyłe(?:(?!<\/w:p>).)*<\/w:p>/s) || [""])[0];
    return [...p.matchAll(/<w:r>(.*?)<\/w:r>/gs)].map((m) => [(m[1].match(/<w:t[^>]*>([^<]*)/) || [])[1], /<w:i\/>|<w:i w:val="1"\/>/.test(m[1]), /<w:i w:val="0"\/>/.test(m[1])]);
  });
  check("kursywa: przycisk I świeci po Ctrl/⌘+I i gaśnie po drugim", iOn && !iOff, JSON.stringify({ iOn, iOff }));
  check("kursywa: „pochyłe” pochylone, „proste” zaraz za nim (bez spacji) NIE, bez zbędnego w:i w:val=\"0\"", JSON.stringify(iRuns) === JSON.stringify([["Zdanie ", false, false], ["pochyłe", true, false], ["proste", false, false]]), JSON.stringify(iRuns));
  await page.selectOption("#fmtParaStyle", "quote"); // Cytat = pochylony ze stylu
  await idle(page);
  await page.evaluate(() => { const p = [...document.querySelectorAll(".docx-preview-host p.docx-editable-p")].find((x) => x.textContent.includes("proste")); placeCaret(p, p.textContent.length); });
  await page.keyboard.press(`${MOD}+i`);
  await page.keyboard.type(" X");
  await page.waitForTimeout(300);
  const qx = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const doc = await z.file("word/document.xml").async("string");
    const p = (doc.match(/<w:p[ >](?:(?!<\/w:p>).)*proste(?:(?!<\/w:p>).)*<\/w:p>/s) || [""])[0];
    return /<w:i w:val="0"\/>(?:(?!<\/w:r>).)*<w:t[^>]*> X/s.test(p);
  });
  check("Cytat (pochylony ze stylu) + Ctrl/⌘+I: dopisek zapisany z w:i w:val=\"0\"", qx, String(qx));

  // ── 4. zaznaczenie całego pogrubionego słowa → B „wciśnięty”; przycisk B zdejmuje pogrubienie ──
  await page.evaluate(() => {
    const p = [...document.querySelectorAll(".docx-preview-host p.docx-editable-p")].find((x) => x.textContent.includes("gruby"));
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const i = n.data.indexOf("gruby");
      if (i >= 0) { const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 5); const s = getSelection(); s.removeAllRanges(); s.addRange(r); return; }
    }
  });
  await page.waitForTimeout(250);
  check("zaznaczone pogrubione słowo: przycisk B „wciśnięty”", (await btn(page)).b, JSON.stringify(await btn(page)));
  await page.click("#fmtBold");
  await page.waitForTimeout(300);
  runs = await savedRuns(page, "Zwykły");
  check("przycisk B na zaznaczeniu zdejmuje pogrubienie, przycisk gaśnie", !(await btn(page)).b && runs.every((r) => !r[1]), JSON.stringify({ runs, b: await btn(page) }));

  // ── 5. rozmiar przy kursorze (zgłoszenie: klik w słowo 8 pt pokazywał 9 pt) ─────────────
  await page.evaluate(() => { const ps = document.querySelectorAll(".docx-preview-host p.docx-editable-p"); const p = ps[ps.length - 1]; placeCaret(p, p.textContent.length); });
  await page.keyboard.press("Enter");
  await page.selectOption("#fmtFontSize", "12");
  await page.keyboard.type("Duże litery ");
  await page.selectOption("#fmtFontSize", "8");
  await page.keyboard.type("malutkie");
  await page.selectOption("#fmtFontSize", "12");
  await page.keyboard.type(" i dalej duże");
  await page.waitForTimeout(300);
  const edge = await page.evaluate(() => {
    const p = [...document.querySelectorAll(".docx-preview-host p.docx-editable-p")].find((x) => x.textContent.includes("malutkie"));
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      const i = n.data.indexOf("malutkie");
      if (i >= 0) { const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 1); const b = r.getBoundingClientRect(); return { x: b.left + 1, y: b.top + b.height / 2 }; }
    }
  });
  await page.mouse.click(edge.x, edge.y); // lewa połowa pierwszej litery → kursor PRZED słowem
  await page.waitForTimeout(200);
  const shownEdge = await page.evaluate(() => document.getElementById("fmtFontSize").value);
  check("klik w lewą krawędź słowa 8 pt (przed nim spacja 12 pt): lista rozmiarów pokazuje 8", shownEdge === "8", shownEdge);
  await page.keyboard.type("bardzo ");
  await page.waitForTimeout(300);
  const typedSize = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const doc = await z.file("word/document.xml").async("string");
    const m = doc.match(/<w:r>(?:(?!<\/w:r>).)*?<w:t[^>]*>[^<]*bardzo[^<]*<\/w:t>/);
    return m ? (m[0].match(/<w:sz w:val="(\d+)"/) || [])[1] : null;
  });
  check("dopisane na początku słowa 8 pt dostaje 8 pt (jak pokazuje lista)", typedSize === "16", String(typedSize));

  // ── 6. A+ na zaznaczeniu z różnymi rozmiarami: każdy o stopień osobno (8→9, 12→14) ──────────
  await page.evaluate(() => {
    const p = [...document.querySelectorAll(".docx-preview-host p.docx-editable-p")].find((x) => x.textContent.includes("malutkie"));
    const r = document.createRange(); r.selectNodeContents(p); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  });
  await page.waitForTimeout(200);
  await page.click("#fmtGrow");
  await page.waitForTimeout(400);
  const sizes = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const doc = await z.file("word/document.xml").async("string");
    const p = (doc.match(/<w:p[ >](?:(?!<\/w:p>).)*malutkie(?:(?!<\/w:p>).)*<\/w:p>/s) || [""])[0];
    const runsOnly = p.replace(/<w:pPr>.*?<\/w:pPr>/s, ""); // bez znaku końca akapitu (w:pPr/w:rPr)
    return [...new Set((runsOnly.match(/<w:sz w:val="(\d+)"/g) || []).map((m) => m.match(/\d+/)[0]))].sort();
  });
  check("A+ na zaznaczeniu 12 i 8 pt: każdy o stopień (14 i 9 pt) — różnice zostają", JSON.stringify(sizes) === JSON.stringify(["18", "28"]), JSON.stringify(sizes));

  await browser.close();
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Format B/I/U [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
