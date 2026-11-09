// corpus-invariants-playwright.js — NIEZMIENNIKI na prawdziwych plikach (korpus), nie na przykładach.
//
//   node scripts/corpus-invariants-playwright.js
//   CORPUS="/ścieżka/a:/ścieżka/b" ENGINE=webkit STEPS=40 SEED=7 node scripts/corpus-invariants-playwright.js
//   ONLY=349 node scripts/corpus-invariants-playwright.js     (tylko pliki z „349” w nazwie)
//   OCR=1 …                                                   (skany: z rozpoznawaniem tekstu)
//
// Testy przykładowe sprawdzają to, co ktoś wymyślił; błędy siedzą w przypadkach, o których nikt
// nie pomyślał (2026-11-08: znaczniki cięcia z PDF znikały z wydruku — żaden test nie miał
// obrazu zakotwiczonego do strony). Tu reguły ogólne, sprawdzane na KAŻDYM pliku korpusu:
//   I1  bez błędów strony (pageerror) przez całą sesję
//   I2  podgląd = plik: akapity na ekranie to dokładnie akapity zapisanego .docx
//   I3  wydruk = Edycja: każdy tekst, obraz i pole tekstowe z kartek Edycji jest na kartkach wydruku
//   I4  zapis → ponowne otwarcie: te same akapity
//   I5  po losowych edycjach (pisanie, Enter, Backspace, Delete; ziarno SEED) dalej I1, I2, I3
//   I6  Cofnij do skutku = dokument z otwarcia (tekst i wyspy każdego akapitu)
//
// Korpus domyślnie: docs/samples/*.docx + „../pdfy do polepszania pdf to docx” (*.pdf, *.docx) —
// pliki użytkownika poza repo. Nowy dziwny plik = wrzucić do folderu, od tej nocy jest testowany.
// Raport: output/corpus/<data>-<silnik>.md (+ .json). Kod wyjścia 1, gdy coś się nie zgadza.
const fs = require("fs");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const STEPS = Number(process.env.STEPS || 25);
const SEED = Number(process.env.SEED || 1);
const ONLY = process.env.ONLY || "";
const ROOT = path.resolve(__dirname, "..");
const DIRS = (process.env.CORPUS || [path.join(ROOT, "docs/samples"), path.resolve(ROOT, "../pdfy do polepszania pdf to docx")].join(":"))
  .split(":").map((d) => path.resolve(ROOT, d)).filter((d) => d && fs.existsSync(d));
const MIME = { ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".pdf": "application/pdf" };
const FILES = DIRS.flatMap((d) => fs.readdirSync(d).filter((f) => MIME[path.extname(f).toLowerCase()] && !f.startsWith("~$")).map((f) => path.join(d, f)))
  .filter((f) => !ONLY || path.basename(f).includes(ONLY));

// mulberry32 — powtarzalne „losowe” edycje dla danego ziarna i pliku
function rng(seed) {
  let a = seed >>> 0;
  return () => { a = (a + 0x6D2B79F5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}
const hash = (s) => [...s].reduce((h, c) => (Math.imul(h, 31) + c.charCodeAt(0)) >>> 0, 7);
const idle = (page, timeout = 30000) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout }).then(() => page.waitForTimeout(300));

// I2: akapity treści na ekranie (te, które aplikacja mapuje na akapity pliku) vs akapity zapisanego
// pliku — liczba i tekst w tej samej kolejności. Tekst z ekranu tak, jak go czyta zapis
// (extractRunsFromPreviewParagraph): odnośniki przypisów, pola i tabulatory jako w pliku.
const previewVsFile = (page) => page.evaluate(async () => {
  const bytes = await buildDocumentForSave();
  const els = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
  // bez tabulatorów, łamań i symboli (☐ pola wyboru) — odczyt tekstu z pliku ich nie oddaje;
  // akapit tylko do odczytu (data-lock) zapis przepisuje z pliku bez zmian — liczy się tylko jego obecność
  const norm = (t) => String(t).replace(/[\t\n\uFEFF]|\p{So}/gu, "");
  const dom = els.map((p) => (p.dataset.lock ? null : norm(previewRunsToPlainText(extractRunsFromPreviewParagraph(p)))));
  const xml = (await extractParagraphTextsFromDocx(bytes)).map(norm);
  let at = -1;
  for (let i = 0; i < Math.max(dom.length, xml.length); i++) if (i >= dom.length || i >= xml.length || (dom[i] !== null && dom[i] !== xml[i])) { at = i; break; }
  // wyspy (komentarze, przypisy, pola, rysunki — fragmenty z pliku przenoszone w całości): te same
  // w tej samej kolejności. Porównanie z PLIKIEM W PAMIĘCI (to, co ekran ma pokazywać), dla
  // akapitów nieedytowanych — z zapisem nie: zapis bierze akapit z ekranu, więc brak na ekranie
  // przechodzi do zapisu i oba się „zgadzają”. Tak przeszedł brak znaczników odpowiedzi na
  // komentarz (2026-11-10). Pola tekstowe (w:txbxContent) rysuje osobna warstwa — pomijamy.
  let islandAt = -1, islandDom = "", islandXml = "";
  if (at < 0) {
    const fileRuns = await extractParagraphRunsFromDocx(originalFileBytes);
    const edited = new Set(collectInlineParagraphEdits().map((e) => e.index));
    const isl = (runs) => (runs || []).filter((r) => r.island && !r.island.includes("txbxContent")).map((r) => r.island);
    for (let i = 0; i < els.length; i++) {
      if (els[i].dataset.lock || edited.has(i)) continue;
      const a = isl(extractRunsFromPreviewParagraph(els[i])), b = isl(fileRuns[i]);
      if (a.length !== b.length || a.some((x, k) => x !== b[k])) { islandAt = i; islandDom = `${a.length}: ${a.join(" ").slice(0, 120)}`; islandXml = `${b.length}: ${b.join(" ").slice(0, 120)}`; break; }
    }
  }
  const html = at < 0 ? undefined : els[at]?.innerHTML.slice(0, 300);
  if (islandAt >= 0) return { ok: false, wyspy: true, at: islandAt, dom: islandDom, xml: islandXml };
  return { ok: at < 0, nDom: dom.length, nXml: xml.length, at, dom: dom[at]?.slice(0, 60), xml: xml[at]?.slice(0, 60), html };
});

// I3: zawartość kartek Edycji vs kartek wydruku. Nagłówki/stopki (na wydruku na każdej stronie)
// i kopie bloków przeciętych granicą strony (pp-dup) — poza porównaniem.
const editVsPrint = (page) => page.evaluate(async () => {
  const norm = (s) => s.replace(/[\uFEFF\u200B]/g, "").replace(/\s+/g, " ").trim();
  const skip = (el) => !!el.closest("header, footer, .pp-hf, .pp-dup, .pp-page-label, .pp-unprintable");
  const take = (root) => {
    const texts = [...root.querySelectorAll("p")].filter((p) => !skip(p) && !p.parentElement.closest("p")).map((p) => norm(p.textContent)).filter(Boolean);
    const imgs = [...root.querySelectorAll("img")].filter((i) => !skip(i) && i.naturalWidth > 0 && getComputedStyle(i).display !== "none").length;
    const boxes = [...root.querySelectorAll(".dwb-tb")].filter((b) => !skip(b)).length;
    return { texts, imgs, boxes };
  };
  const edit = take(document.querySelector(".docx-preview-host"));
  dwbPrint.open();
  const t0 = Date.now();
  while (!(dwbPrint.isOpen() && document.querySelector(".pp-sheet")) && Date.now() - t0 < 30000) await new Promise((r) => setTimeout(r, 100));
  await new Promise((r) => setTimeout(r, 500));
  const sheets = [...document.querySelectorAll(".pp-sheet")];
  const pr = { texts: [], imgs: 0, boxes: 0 };
  sheets.forEach((s) => { const x = take(s); pr.texts.push(...x.texts); pr.imgs += x.imgs; pr.boxes += x.boxes; });
  dwbPrint.close();
  const have = new Set(pr.texts);
  const missing = [...new Set(edit.texts)].filter((t) => !have.has(t));
  return {
    ok: sheets.length > 0 && !missing.length && pr.imgs >= edit.imgs && pr.boxes >= edit.boxes,
    sheets: sheets.length, missing: missing.slice(0, 3).map((t) => t.slice(0, 60)), nMissing: missing.length,
    imgs: `${edit.imgs} → ${pr.imgs}`, boxes: `${edit.boxes} → ${pr.boxes}`,
  };
});

// I4: zapis → ponowne otwarcie z tych bajtów → te same akapity
const saveReload = (page) => page.evaluate(async () => {
  const bytes = await buildDocumentForSave();
  const before = (await extractParagraphTextsFromDocx(bytes)).map(String);
  await reloadFromBytes(bytes);
  await new Promise((r) => setTimeout(r, 300));
  const after = (await extractParagraphTextsFromDocx(await buildDocumentForSave())).map(String);
  const at = before.findIndex((t, i) => t !== after[i]);
  return { ok: at < 0 && before.length === after.length, n: `${before.length} → ${after.length}`, at, was: before[at]?.slice(0, 60), now: after[at]?.slice(0, 60) };
});

// I6: stan dokumentu do porównania — tekst i wyspy (komentarze, przypisy, pola) każdego akapitu
// zapisanego pliku
const docSnapshot = (page) => page.evaluate(async () => {
  const bytes = await buildDocumentForSave();
  const runs = await extractParagraphRunsFromDocx(bytes);
  return runs.map((rs) => (rs || []).map((r) => (r.island ? `[${r.island.length}]` : r.text || "")).join(""));
});
// Cofnij do skutku; jeśli stos był obcięty limitem (undo.js: limit()) — porównanie nie ma sensu
async function undoAll(page, start) {
  const depth = await page.evaluate(() => dwbUndo._debug().undo.length);
  const cap = await page.evaluate(() => { const n = originalFileBytes?.byteLength || 0; return n > 20 * 1024 * 1024 ? 10 : n > 5 * 1024 * 1024 ? 25 : 60; });
  for (let k = 0; k < 200 && (await page.evaluate(() => dwbUndo.canUndo())); k++) {
    await page.evaluate(() => dwbUndo.undo());
    await idle(page).catch(() => {});
  }
  const now = await docSnapshot(page);
  if (depth >= cap) return { ok: true, pominięte: `stos Cofnij obcięty (${depth}/${cap})` };
  const at = Math.max(start.length, now.length) === 0 ? -1 : [...Array(Math.max(start.length, now.length)).keys()].find((i) => start[i] !== now[i]);
  return { ok: at === undefined || at < 0, kroków: depth, akapitów: `${start.length} → ${now.length}`, at, było: start[at]?.slice(0, 60), jest: now[at]?.slice(0, 60) };
}

async function openFile(page, file) {
  const ext = path.extname(file).toLowerCase();
  await page.evaluate(() => {
    window.__conv = null;
    if (!window.__convHooked) { window.addEventListener("dwb:pdf-converted", (e) => (window.__conv = e.detail)); window.__convHooked = true; }
  });
  await page.setInputFiles("#fileInput", { name: path.basename(file), mimeType: MIME[ext], buffer: fs.readFileSync(file) });
  if (ext === ".pdf") {
    // skan / tekst jako krzywe — pytanie o rozpoznanie tekstu: domyślnie „Zostaw jako obraz” (szybko),
    // OCR=1 — „Rozpoznaj tekst” (wolno, ale sprawdza też ścieżkę OCR)
    const ask = page.waitForSelector("#pdfConvAsk:not(.hidden)", { timeout: 60000 })
      .then(() => page.click(process.env.OCR === "1" ? "#pdfConvOcrYes" : "#pdfConvOcrNo")).catch(() => {});
    await page.waitForFunction(() => window.__conv, null, { timeout: 180000 });
    void ask;
  }
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 120000 });
  await page.waitForTimeout(500);
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await page.waitForTimeout(300);
}

// I5: losowe edycje w akapitach treści (bez zablokowanych — pola, linki, przypisy)
async function randomEdits(page, rand, log) {
  const WORDS = ["test", "Zażółć", "ąę", "123", "a", "Koniec."];
  for (let s = 0; s < STEPS; s++) {
    const n = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => !p.dataset.lock).length);
    if (!n) return;
    const i = Math.floor(rand() * n);
    const atEnd = rand() < 0.5;
    const op = ["type", "type", "enter", "backspace", "delete"][Math.floor(rand() * 5)];
    await page.evaluate(([idx, end]) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((x) => !x.dataset.lock)[idx];
      p.scrollIntoView({ block: "center" });
      placeCaret(p, end ? p.textContent.length : 0);
    }, [i, atEnd]);
    log.push(`${op}@${i}${atEnd ? "$" : "^"}`);
    if (op === "type") await page.keyboard.type(` ${WORDS[Math.floor(rand() * WORDS.length)]}`);
    else if (op === "enter") await page.keyboard.press("Enter");
    else if (op === "backspace") await page.keyboard.press("Backspace");
    else await page.keyboard.press("Delete");
    await idle(page).catch(() => {});
    // przerwa w pisaniu (jak użytkownik) — osobny krok Cofnij; I6 przechodzi wtedy przez wiele
    // kroków (też szybką ścieżkę Cofnij), a nie jeden wielki (dawniej 25 edycji = 1 krok)
    if (rand() < 0.3) await page.evaluate(() => dwbUndo.breakTyping());
  }
}

async function runFile(browser, file) {
  const name = path.basename(file);
  const out = { file: name, checks: [], errors: [] };
  const check = (id, label, r) => out.checks.push({ id, label, ok: !!r?.ok, detail: r });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  page.on("pageerror", (e) => out.errors.push(e.message));
  page.on("dialog", (d) => d.dismiss().catch(() => {}));
  const t0 = Date.now();
  try {
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await openFile(page, file);
    check("I2", "podgląd = plik (po otwarciu)", await previewVsFile(page));
    check("I3", "wydruk = Edycja (po otwarciu)", await editVsPrint(page));
    await idle(page);
    check("I4", "zapis → ponowne otwarcie bez zmian", await saveReload(page));
    await idle(page); // przerysowanie kończy oznaczanie (przypisy, pola) po powrocie reloadFromBytes
    const start = await docSnapshot(page);
    const log = [];
    await randomEdits(page, rng(SEED ^ hash(name)), log);
    out.edits = log.join(" ");
    // odroczone porządki aplikacji po pisaniu (doc-notes.js watchRefs: usunięty odnośnik przypisu →
    // przerysowanie 450 ms po ostatnim klawiszu) — sprawdzamy stan jak po przerwie w pisaniu,
    // nie w jej trakcie (przy szybszym tempie testu przypis-widmo był jeszcze na ekranie)
    await page.evaluate(() => new Promise((r) => setTimeout(r, 700))); // zegar strony — runner go nie skraca (SLEEP_SCALE)
    await idle(page);
    check("I2", `podgląd = plik (po ${STEPS} losowych edycjach)`, await previewVsFile(page));
    check("I3", "wydruk = Edycja (po edycjach)", await editVsPrint(page));
    await idle(page);
    check("I6", "Cofnij wszystko = dokument z otwarcia", await undoAll(page, start));
  } catch (e) {
    out.checks.push({ id: "–", label: "sesja przerwana", ok: false, detail: { error: String(e.message || e).split("\n")[0] } });
  }
  out.checks.push({ id: "I1", label: "bez błędów strony", ok: !out.errors.length, detail: { errors: out.errors.slice(0, 3) } });
  out.seconds = Math.round((Date.now() - t0) / 1000);
  await context.close();
  return out;
}

async function main() {
  if (!FILES.length) { console.log("Korpus pusty:", DIRS.join(", ") || "(brak katalogów)"); process.exit(0); }
  const browser = await pw[ENGINE].launch({ headless: true });
  const all = [];
  for (const f of FILES) {
    const r = await runFile(browser, f);
    all.push(r);
    const bad = r.checks.filter((c) => !c.ok);
    console.log(`${bad.length ? "❌" : "✅"} ${r.file} (${r.seconds} s)${bad.length ? "" : ""}`);
    bad.forEach((c) => console.log(`   ✗ ${c.id} ${c.label} — ${JSON.stringify(c.detail).slice(0, 300)}`));
    if (bad.length && r.edits) console.log(`   edycje (SEED=${SEED}): ${r.edits.slice(0, 300)}`);
  }
  await browser.close();
  const failed = all.filter((r) => r.checks.some((c) => !c.ok));
  const day = new Date().toISOString().slice(0, 10);
  const dir = path.join(ROOT, "output/corpus");
  fs.mkdirSync(dir, { recursive: true });
  const md = [`# Niezmienniki na korpusie — ${day} [${ENGINE}], SEED=${SEED}, STEPS=${STEPS}`, "",
    `${all.length - failed.length}/${all.length} plików bez zastrzeżeń.`, "",
    ...all.map((r) => {
      const bad = r.checks.filter((c) => !c.ok);
      if (!bad.length) return `- ✅ ${r.file} (${r.seconds} s)`;
      return [`- ❌ **${r.file}** (${r.seconds} s)`, ...bad.map((c) => `  - ${c.id} ${c.label}: \`${JSON.stringify(c.detail).slice(0, 400)}\``),
        r.edits ? `  - powtórzenie: \`ONLY="${r.file}" SEED=${SEED} STEPS=${STEPS} ENGINE=${ENGINE} node scripts/corpus-invariants-playwright.js\`` : ""].filter(Boolean).join("\n");
    })].join("\n");
  fs.writeFileSync(path.join(dir, `${day}-${ENGINE}.md`), md + "\n");
  fs.writeFileSync(path.join(dir, `${day}-${ENGINE}.json`), JSON.stringify(all, null, 1));
  console.log(`\n[${ENGINE}] ${all.length - failed.length}/${all.length} plików OK — raport: output/corpus/${day}-${ENGINE}.md`);
  process.exit(failed.length ? 1 : 0);
}

main().catch((e) => { console.error(e); process.exit(1); });
