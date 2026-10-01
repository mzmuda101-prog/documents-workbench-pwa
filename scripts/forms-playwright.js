// forms-playwright.js — formularz Worda (app/docx-forms.js + app/forms-panel.js).
//
// Fixture: docs/samples/forms-sample.docx (node scripts/gen-forms-docx.mjs) — kontrolki zawartości
// (tekst z tekstem zastępczym, lista, kombi, data, pola wyboru, lista na cały akapit, pole
// zablokowane, dwa pola powiązane z customXml), stare pola FORMTEXT/FORMCHECKBOX/FORMDROPDOWN.
// Sprawdza TREŚĆ PLIKU po każdej zmianie, nie tylko UI. Uruchom też: ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden")
  && typeof formScan !== "undefined" && formScan?.bytes === originalFileBytes, null, { timeout: 20000 }).then(() => page.waitForTimeout(250));

// Stan pliku: kontrolki po w:id (tekst, placeholder, styl, lastValue, checked, fullDate), stare pola, customXml.
const fileState = (page, fromSave = false) => page.evaluate(async (fromSave) => {
  const bytes = fromSave ? await buildDocumentForSave() : originalFileBytes;
  const z = await JSZip.loadAsync(bytes);
  const x = await z.file("word/document.xml").async("string");
  const d = new DOMParser().parseFromString(x, "application/xml");
  const W = W_NS;
  const sdt = {};
  [...d.getElementsByTagNameNS(W, "sdt")].forEach((s) => {
    const pr = ffKid(s, "sdtPr");
    const id = ffAttr(ffKid(pr, "id"), "val");
    const c = ffKid(s, "sdtContent");
    const list = ffKid(pr, "dropDownList") || ffKid(pr, "comboBox");
    sdt[id] = {
      text: c ? ffText(c) : null,
      ph: !!ffKid(pr, "showingPlcHdr"),
      phStyle: c ? /PlaceholderText/.test(new XMLSerializer().serializeToString(c)) : false,
      last: list ? ffAttr(list, "lastValue") : null,
      checked: ffKid(pr, "checkbox", W14_NS) ? ffAttr(ffKid(ffKid(pr, "checkbox", W14_NS), "checked", W14_NS), "val", W14_NS) : null,
      font: c ? ffAttr(c.getElementsByTagNameNS(W, "rFonts")[0], "ascii") : null,
      fullDate: ffKid(pr, "date") ? ffAttr(ffKid(pr, "date"), "fullDate") : null,
      binding: !!ffKid(pr, "dataBinding"),
    };
  });
  const legacy = collectLegacyFormFields(d).map((lf) => ({
    result: legacyResultRuns(lf).map(ffText).join(""),
    checked: ffAttr(lf.ffData.getElementsByTagNameNS(W, "checked")[0], "val"),
    ddResult: ffAttr(lf.ffData.getElementsByTagNameNS(W, "result")[0], "val"),
  }));
  const item = await z.file("customXml/item1.xml").async("string");
  return { sdt, legacy, klient: (item.match(/<klient>([^<]*)<\/klient>/) || [])[1], paras: (await extractParagraphTextsFromDocx(bytes)).map(String) };
}, fromSave);

// Klik w pole w podglądzie (środek pierwszego prostokąta jego tekstu).
async function clickField(page, key) {
  const box = await page.evaluate((k) => {
    const r = formUi.ranges.get(k);
    const el = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
    el.scrollIntoView({ block: "center" });
    const rect = [...r.getClientRects()].find((q) => q.width > 0) || r.getBoundingClientRect();
    return { x: rect.left + Math.min(rect.width / 2, 20), y: rect.top + rect.height / 2 };
  }, key);
  await page.mouse.click(box.x, box.y);
  await page.waitForTimeout(150);
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
  await page.goto(`${APP_URL}?sample=forms-sample`, { waitUntil: "load" });
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await idle(page);

  // ── odczyt ─────────────────────────────────────────────────────────────────
  const scan = await page.evaluate(() => ({
    kinds: formScan.fields.map((f) => `${f.key}:${f.kind}`).join(" "),
    main: docFormCounts.fields, empty: docFormCounts.empty, prot: formScan.protection,
    badge: document.querySelector("#panel-forms > summary .panel-count")?.textContent,
    hl: typeof CSS !== "undefined" && CSS.highlights ? CSS.highlights.has("dwb-form") : "brak API",
    locks: collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => p.dataset.lock || "-"),
    date: formatWordDate("2026-10-01", "d MMMM yyyy", "pl-PL") + " | " + formatWordDate("2026-10-01", "dddd, dd.MM.yy", "pl-PL") + " | " + formatWordDate("2026-10-01", "MMMM yyyy", "pl-PL"),
  }));
  check("13 pól (12 + kopia powiązanego), spis treści i blok „Uwagi” to nie pola",
    scan.kinds === "s1:text s2:dropdown s3:combo s4:date s5:checkbox s6:checkbox s7:dropdown s8:text s9:text s10:text f0:text f1:checkbox f2:dropdown", scan.kinds);
  check("plakietka panelu = 12, do uzupełnienia 2, ochrona „forms”", scan.badge === "12" && scan.main === 12 && scan.empty === 2 && scan.prot === "forms", JSON.stringify(scan));
  check("wyróżnienie pól (CSS Highlight)", scan.hl === true || scan.hl === "brak API", String(scan.hl));
  check("akapity z polem tylko do odczytu, „Uwagi” (bogaty tekst) i spis treści edytowalne",
    scan.locks[2] === "lockForm" && scan.locks[8] === "lockForm" && scan.locks[14] === "lockForm" && scan.locks[1] === "-" && scan.locks[12] === "-", scan.locks.join(" "));
  check("daty w formacie Worda (dopełniacz przy dniu, mianownik bez dnia)", scan.date === "1 października 2026 | czwartek, 01.10.26 | październik 2026", scan.date);
  const glyph = await page.evaluate(() => document.querySelector('.ff-glyph[data-ff="f1"]')?.textContent);
  check("stare pole wyboru ma w podglądzie kratkę ☐", glyph === "☐", glyph);

  // ── BŁĄD: edycja akapitu obok pola nie może rozbić kontrolki ───────────────
  await page.click('.mode-btn[data-mode="edit"]');
  await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[12];
    placeCaret(p, p.textContent.length);
    document.execCommand("insertText", false, " Dopisane.");
  });
  await page.waitForTimeout(200);
  let s = await fileState(page, true);
  check("pisanie w innym akapicie: lista „Dział” cała (IT w środku), pola nietknięte",
    s.sdt["12"].text === "IT" && s.sdt["15"].text === "☐" && s.paras[12] === "Uwagi można pisać tu zwyczajnie. Dopisane.", JSON.stringify(s.sdt["12"]));
  // pełne porównanie (bez śledzenia zmian) też nie może ruszyć akapitów z polami
  s = await page.evaluate(async () => { inlineDirtyValid = false; const r = collectInlineParagraphEdits().map((e) => e.index); inlineDirtyValid = true; return r; });
  check("pełne porównanie podglądu z plikiem pomija akapity z polami", JSON.stringify(s) === "[12]", JSON.stringify(s));
  // Backspace na początku akapitu po zablokowanym — bez sklejania (zgubiłoby pole)
  const before = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).length);
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[12]; placeCaret(p, 0); });
  await page.keyboard.press("Backspace");
  await page.waitForTimeout(300);
  const after = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).length);
  s = await fileState(page, true);
  const mergeToast = await page.$$eval(".toast", (els) => els.map((e) => e.textContent).join(" | "));
  check("Backspace przed akapitem z polem nie skleja (pola całe) i mówi dlaczego",
    before === after && s.sdt["20"]?.text === "ACME sp. z o.o." && s.paras[12].startsWith("Uwagi") && /pole formularza/.test(mergeToast), `${before}→${after} | ${mergeToast}`);

  // ── pole wyboru: klik w dokumencie ─────────────────────────────────────────
  await clickField(page, "s5");
  await idle(page);
  s = await fileState(page);
  check("klik w ☐ zaznacza: checked=1, ☒, czcionka MS Gothic", s.sdt["15"].checked === "1" && s.sdt["15"].text === "☒" && s.sdt["15"].font === "MS Gothic", JSON.stringify(s.sdt["15"]));
  check("dopisany wcześniej tekst przetrwał", s.paras[12] === "Uwagi można pisać tu zwyczajnie. Dopisane.", s.paras[12]);
  const undoLabel = await page.getAttribute("#undoBtn", "aria-label");
  check("Cofnij: „pole formularza”", /pole formularza/.test(undoLabel || ""), undoLabel);
  await page.click("#undoBtn");
  await idle(page);
  s = await fileState(page);
  check("Cofnij odznacza", s.sdt["15"].checked === "0" && s.sdt["15"].text === "☐", JSON.stringify(s.sdt["15"]));

  // szybkie klikanie: kolejka, każde kliknięcie liczy stan od nowa (3× = zaznaczone)
  await page.evaluate(() => Promise.all([toggleFormCheckbox("s5"), toggleFormCheckbox("s5"), toggleFormCheckbox("s5")]));
  await idle(page);
  s = await fileState(page);
  const overlayHidden = await page.evaluate(() => document.getElementById("loadingOverlay").classList.contains("hidden"));
  check("3 szybkie kliknięcia w ☐ = zaznaczone (nic nie przepada)", s.sdt["15"].checked === "1", JSON.stringify(s.sdt["15"]));
  check("przy małym dokumencie bez nakładki „Renderowanie…”", overlayHidden);
  await page.evaluate(() => toggleFormCheckbox("s5"));
  await idle(page);

  // ── lista: okienko z pozycjami ─────────────────────────────────────────────
  await clickField(page, "s2");
  const opts = await page.$$eval(".ff-pop .ff-pop-opt", (els) => els.map((e) => e.textContent));
  check("klik w listę otwiera okienko z pozycjami", opts.join("|") === "Wybierz element.|Sprzedaż|Marketing|IT", opts.join("|"));
  await page.click(".ff-pop .ff-pop-opt:has-text('Marketing')");
  await idle(page);
  s = await fileState(page);
  check("wybór „Marketing”: tekst + lastValue=MKT", s.sdt["12"].text === "Marketing" && s.sdt["12"].last === "MKT", JSON.stringify(s.sdt["12"]));

  // ── lista na cały akapit ───────────────────────────────────────────────────
  await clickField(page, "s7");
  await page.click(".ff-pop .ff-pop-opt:has-text('Pilny')");
  await idle(page);
  s = await fileState(page);
  check("lista-akapit: „Pilny”", s.sdt["17"].text === "Pilny" && s.sdt["17"].last === "P", JSON.stringify(s.sdt["17"]));

  // ── tekst z tekstem zastępczym ─────────────────────────────────────────────
  await clickField(page, "s1");
  await page.fill(".ff-pop .ff-pop-input", "Jan Kowalski");
  await page.keyboard.press("Enter");
  await idle(page);
  s = await fileState(page);
  check("tekst: wpis zastępuje podpowiadacz (bez showingPlcHdr i szarego stylu)", s.sdt["11"].text === "Jan Kowalski" && !s.sdt["11"].ph && !s.sdt["11"].phStyle, JSON.stringify(s.sdt["11"]));
  check("akapit „Imię i nazwisko: Jan Kowalski”", s.paras[2] === "Imię i nazwisko: Jan Kowalski", s.paras[2]);

  // ── zablokowane ────────────────────────────────────────────────────────────
  await clickField(page, "s8");
  const lockedMsg = await page.textContent(".ff-pop").catch(() => "");
  check("pole zablokowane: okienko mówi „zablokowane”, bez pola do wpisania", /zablokowane/.test(lockedMsg) && !(await page.$(".ff-pop input")), lockedMsg);
  await page.keyboard.press("Escape");
  check("Esc zamyka okienko", !(await page.$(".ff-pop")));

  // ── panel ──────────────────────────────────────────────────────────────────
  await page.evaluate(async () => { setSidebarOpen(true); document.getElementById("panel-forms").open = true; await ensureLazyFeature("forms"); renderFormsPanel(); });
  await page.waitForTimeout(400);
  const panel = await page.evaluate(() => ({
    rows: document.querySelectorAll("#ffList .ff-row").length,
    summary: document.getElementById("ffSummary").textContent,
    prot: !document.getElementById("ffProtection").classList.contains("hidden"),
    lockedDisabled: document.getElementById("ff-s8").disabled,
    copies: [...document.querySelectorAll("#ffList .rv-meta")].some((m) => /w 2 miejscach/.test(m.textContent)),
  }));
  check("panel: 12 wierszy (kopia powiązanego raz), ochrona pokazana, zablokowane wyłączone, „w 2 miejscach”",
    panel.rows === 12 && panel.prot && panel.lockedDisabled && panel.copies, JSON.stringify(panel));
  check("panel: podsumowanie po wypełnieniu imienia — do uzupełnienia 1", /do uzupełnienia: 1/.test(panel.summary), panel.summary);

  await page.fill("#ff-s4", "2026-10-01");
  await page.dispatchEvent("#ff-s4", "change");
  await idle(page);
  s = await fileState(page);
  check("data z panelu: „1 października 2026”, fullDate", s.sdt["14"].text === "1 października 2026" && /^2026-10-01T/.test(s.sdt["14"].fullDate), JSON.stringify(s.sdt["14"]));

  await page.fill("#ff-s3", "Gdańsk");
  await page.press("#ff-s3", "Enter");
  await idle(page);
  s = await fileState(page);
  check("kombi z panelu: własny tekst „Gdańsk”", s.sdt["13"].text === "Gdańsk" && !s.sdt["13"].ph && s.sdt["13"].last === "Gdańsk", JSON.stringify(s.sdt["13"]));

  await page.fill("#ff-s9", "Beta SA");
  await page.press("#ff-s9", "Enter");
  await idle(page);
  s = await fileState(page);
  check("powiązane: obie kopie „Beta SA” + źródło customXml", s.sdt["19"].text === "Beta SA" && s.sdt["20"].text === "Beta SA" && s.klient === "Beta SA" && s.sdt["19"].binding, JSON.stringify({ a: s.sdt["19"].text, b: s.sdt["20"].text, k: s.klient }));

  // stare pola
  await page.fill("#ff-f0", "123 456 789 000 111 222");
  await page.press("#ff-f0", "Enter");
  await idle(page);
  await page.check("#ff-f1");
  await idle(page);
  await page.selectOption("#ff-f2", "L");
  await idle(page);
  s = await fileState(page);
  check("FORMTEXT: wynik przycięty do 15 znaków", s.legacy[0].result === "123 456 789 000", s.legacy[0].result);
  check("FORMCHECKBOX: checked=1", s.legacy[1].checked === "1", JSON.stringify(s.legacy[1]));
  check("FORMDROPDOWN: result=2 i tekst „L”", s.legacy[2].ddResult === "2" && s.legacy[2].result === "L", JSON.stringify(s.legacy[2]));
  const glyph2 = await page.evaluate(() => document.querySelector('.ff-glyph[data-ff="f1"]')?.textContent);
  check("kratka starego pola w podglądzie: ☒", glyph2 === "☒", glyph2);

  // ── Narzędzia edycji: WIELKIE LITERY na całym dokumencie omija akapity z polami ──
  await page.evaluate(() => applyDocumentEdit({ op: "case", mode: "upper", scope: "all" }));
  await idle(page);
  s = await fileState(page);
  check("WIELKIE LITERY: zwykłe akapity tak, pola nietknięte", s.paras[16] === "KONIEC FORMULARZA." && s.sdt["12"].text === "Marketing" && s.paras[2] === "Imię i nazwisko: Jan Kowalski", `${s.paras[16]} | ${s.paras[2]}`);

  // ── zapis i ponowne otwarcie ───────────────────────────────────────────────
  const reread = await page.evaluate(async () => {
    const bytes = await buildDocumentForSave();
    const sc = await scanFormFields(bytes);
    return sc.fields.map((f) => `${f.key}=${JSON.stringify(f.value)}`).join(" ");
  });
  check("plik po zapisie czyta się z nowymi wartościami",
    /s1="Jan Kowalski"/.test(reread) && /s2="MKT"/.test(reread) && /s4="2026-10-01"/.test(reread) && /f1=true/.test(reread) && /f2="L"/.test(reread), reread);

  await browser.close();
  if (errors.length) check("bez błędów w konsoli", false, errors.join(" | "));
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ formularz [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
