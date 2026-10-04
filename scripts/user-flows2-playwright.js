// user-flows2-playwright.js — druga tura funkcji klikanych jak użytkownik: cofanie wstawień
// z panelu, wklejanie, szybka edycja w Strukturze, auto-rozwijanie snippetów, zamykanie
// z niezapisanymi zmianami, język EN (bez surowych kluczy), motyw, zoom, skupienie, skróty sekcji.
//
//   node scripts/user-flows2-playwright.js                (Chromium)
//   ENGINE=webkit node scripts/user-flows2-playwright.js  (WebKit)

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1366, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.clear(); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  const dialogs = [];
  page.on("dialog", (d) => { dialogs.push(`${d.type()}: ${d.message()}`); d.accept(d.type() === "prompt" ? d.defaultValue() : undefined); });

  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.waitForTimeout(500);

  const fileParas = () => page.evaluate(async () => {
    await waitInlineStructuralIdle?.();
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
    return collectParagraphElements(d.documentElement, "all").map((p) => ({ text: paragraphSearchText(p), xml: new XMLSerializer().serializeToString(p) }));
  });
  const clickEndOf = async (i) => {
    const box = await page.evaluate((i) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
      p.scrollIntoView({ block: "center" });
      const r = document.createRange(); r.selectNodeContents(p);
      const rects = r.getClientRects(); const last = rects[rects.length - 1];
      return { x: Math.min(last.right + 4, p.getBoundingClientRect().right - 1), y: last.top + last.height / 2 };
    }, i);
    await page.mouse.click(box.x, box.y);
  };
  const openPanel = async (id) => {
    await page.evaluate((id) => {
      if (!document.documentElement.classList.contains("sidebar-open")) document.getElementById("panelToggle").click();
      document.getElementById(id).scrollIntoView({ block: "start" });
    }, id);
    await page.waitForTimeout(250);
    if (!(await page.evaluate((id) => document.getElementById(id).open, id))) await page.click(`#${id} > summary`);
    await page.waitForTimeout(400);
  };
  const editMode = async () => { await page.click('.mode-btn[data-mode="edit"]'); await page.waitForTimeout(400); };

  // ── 1. cofanie wstawienia z panelu (placeholder) — pisanie przed nim zostaje ──
  await editMode();
  await clickEndOf(1);
  await page.keyboard.type(" PRZED");
  await page.waitForTimeout(1700); // koniec „serii pisania” (1,5 s)
  await openPanel("panel-placeholders");
  await page.fill("#phInsertName", "pole");
  await page.click("#phInsertBtn");
  await page.waitForTimeout(300);
  let f = await fileParas();
  const withPh = f[1].text;
  await page.click("#undoBtn");
  await page.waitForTimeout(700);
  f = await fileParas();
  check("↶ po „Wstaw placeholder” usuwa sam placeholder, wpisane wcześniej „PRZED” zostaje",
    /\{\{pole\}\}/.test(withPh) && !/\{\{pole\}\}/.test(f[1].text) && / PRZED/.test(f[1].text), `${withPh} → ${f[1].text}`);

  // ── 2. wklejanie (tekst z Worda / strony jako HTML) ─────────────────────────
  const parasBefore = (await fileParas()).length;
  await clickEndOf(3);
  await page.evaluate(() => {
    const dt = new DataTransfer();
    dt.setData("text/html", '<p style="color:red"><b>Wklejony</b> <span style="font-family:Comic Sans MS">tekst</span></p><p>drugi akapit</p><script>alert(1)</script>');
    dt.setData("text/plain", "Wklejony tekst\ndrugi akapit");
    const ev = new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true });
    document.activeElement.dispatchEvent(ev);
  });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 15000 });
  await page.waitForTimeout(700);
  f = await fileParas();
  const pasted = f.map((p) => p.text).join("¶");
  // od 2026-10-04 (paste-rich.js): pogrubienie zostaje, akapity ze schowka = akapity; kroje,
  // kolory i skrypty ze źródła dalej NIE przechodzą
  check("wklejenie HTML: pogrubienie zostaje, akapity jako akapity, bez obcych krojów/kolorów/skryptów",
    /Wklejony tekst$/.test(f[3].text) && /<w:b w:val="1"\/><\/w:rPr><w:t>Wklejony/.test(f[3].xml) && f[4]?.text === "drugi akapit" && !/Comic|alert|color/.test(f[3].xml + f[4]?.xml) && f.length === parasBefore + 1,
    `${parasBefore} → ${f.length} akapitów; ${f[3].xml.slice(-300)}`);
  await page.click("#undoBtn");
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 15000 });
  await page.waitForTimeout(700);
  f = await fileParas();
  check("↶ cofa samo wklejenie", !/Wklejony/.test(f[3].text) && /elektroniczną\.$/.test(f[3].text), f[3].text.slice(-40));
  check("wklejenie: brak błędów strony po wklejeniu", !errors.length, errors.join(" | "));

  // Shift+Enter = łamanie wiersza w Wordzie (<w:br/>), nie „\n” w tekście (Word pokazał spację)
  await clickEndOf(4);
  await page.keyboard.type("LINIA1");
  await page.keyboard.press("Shift+Enter");
  await page.keyboard.type("LINIA2");
  await page.waitForTimeout(300);
  f = await fileParas();
  check("Shift+Enter zapisuje prawdziwe łamanie wiersza (<w:br/>)", /LINIA1<\/w:t><\/w:r><w:r><w:br\/><\/w:r><w:r><w:t[^>]*>LINIA2/.test(f[4].xml) && !/LINIA1\n/.test(f[4].xml), f[4].xml.slice(-160));

  // ── 3. Struktura: szybka edycja akapitu zachowuje formatowanie ──────────────
  await page.evaluate(async () => { // pogrub „Najemca” w akapicie 6, żeby było co chronić
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[6];
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); const n = w.nextNode();
    const r = document.createRange(); r.setStart(n, 0); r.setEnd(n, 7);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r); p.focus();
  });
  await page.click("#fmtBold");
  await page.waitForTimeout(300);
  await openPanel("panel-structure");
  const picked = await page.evaluate(() => {
    const item = documentStructure.outline.find((o) => o.paraIndex === 6);
    if (!item) return false;
    jumpToStructureItem(item);
    return true;
  });
  await page.waitForTimeout(300);
  const qe = await page.$("#structureQuickEditText");
  if (picked && qe && await qe.isVisible()) {
    const cur = await qe.inputValue();
    await qe.fill(cur.replace("zobowiązuje się", "zobowiązał się"));
    await page.click("#structureQuickApplyBtn");
    await page.waitForTimeout(700);
    f = await fileParas();
    check("Struktura → szybka edycja: tekst zmieniony w pliku", /Najemca zobowiązał się/.test(f[6].text), f[6].text.slice(0, 60));
    check("…i pogrubienie „Najemca” zostało", /<w:b( w:val="(1|true)")?\/><\/w:rPr><w:t[^>]*>Najemca</.test(f[6].xml), f[6].xml.slice(0, 220));
  } else check("Struktura → szybka edycja: pole dostępne po wybraniu akapitu", false, `picked=${picked}`);

  // ── 4. snippety: tryb auto (!nazwa + spacja) ────────────────────────────────
  await openPanel("panel-snippets");
  await page.fill("#snName", "adr");
  await page.fill("#snBody", "ul. Testowa 1");
  await page.click("#snSaveBtn");
  await page.selectOption("#snExpandMode", "auto");
  await clickEndOf(8);
  await page.keyboard.type(" !adr ");
  await page.waitForTimeout(600);
  f = await fileParas();
  check("snippet auto: „!adr” + spacja rozwija się w trakcie pisania", /ul\. Testowa 1/.test(f[8].text) && !/!adr/.test(f[8].text), f[8].text.slice(-40));

  // ── 5. zamknięcie z niezapisanymi zmianami pyta ─────────────────────────────
  const dirty = await page.evaluate(() => hasUnsavedChanges);
  dialogs.length = 0;
  await page.evaluate(() => { window.__origConfirm = window.confirm; window.confirm = () => false; }); // „Anuluj”
  await page.evaluate(() => document.getElementById("closeDocBtn").click());
  await page.waitForTimeout(300);
  const stillOpen = await page.evaluate(() => !!originalFileBytes);
  await page.evaluate(() => { window.confirm = window.__origConfirm; });
  check("zamknięcie z niezapisanymi zmianami: pyta, „Anuluj” zostawia dokument", dirty && stillOpen, JSON.stringify({ dirty, stillOpen }));
  // Od kart dokumentów (open-docs.js) otwarcie NIE podmienia dokumentu: nowy wchodzi w nową
  // kartę bez pytania, praca zostaje w karcie obok (niezapisana).
  const name0 = await page.evaluate(() => currentFileName);
  dialogs.length = 0;
  await page.evaluate(() => document.getElementById("loadSampleBtn").click());
  await page.waitForFunction(() => currentFileName === "przewodnik.docx", null, { timeout: 15000 }).catch(() => {});
  await page.evaluate(() => dwbOpenDocs._idle());
  const tabsNow = await page.evaluate(() => ({ name: currentFileName, list: dwbOpenDocs.list() }));
  const prevTab = tabsNow.list.find((d) => d.name === name0);
  check("„Przykładowy dokument” przy niezapisanych zmianach: nowa karta bez pytania, praca zostaje w karcie (niezapisana)", tabsNow.name === "przewodnik.docx" && prevTab && prevTab.dirty && !prevTab.active && !dialogs.length, JSON.stringify({ name0, ...tabsNow, dialogs }));
  if (prevTab) {
    await page.evaluate((id) => dwbOpenDocs.switchTo(id), prevTab.id);
    await page.waitForTimeout(400);
  }
  const name1 = await page.evaluate(() => ({ name: currentFileName, dirty: hasUnsavedChanges }));
  check("powrót na kartę z pracą: ten sam dokument, dalej niezapisany", name1.name === name0 && name1.dirty, JSON.stringify(name1));

  // ── 6. język EN: nigdzie surowych kluczy tłumaczeń ──────────────────────────
  await page.evaluate(() => { document.querySelectorAll("details.panel").forEach((d) => { d.open = true; }); });
  await page.evaluate(async () => { for (const k of Object.keys(LAZY_FEATURE_SCRIPTS)) await ensureLazyFeature(k); });
  await page.evaluate(() => document.querySelector('#langSwitch [data-lang="en"], #langSwitch button:last-child')?.click());
  await page.waitForTimeout(500);
  const raw = await page.evaluate(() => {
    const keys = new Set([...Object.keys(I18N.pl)]);
    const bad = [];
    const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = w.nextNode())) { const s = n.textContent.trim(); if (keys.has(s)) bad.push(s); }
    document.querySelectorAll("[placeholder],[aria-label]").forEach((el) => {
      ["placeholder", "aria-label"].forEach((a) => { const v = el.getAttribute(a); if (v && keys.has(v)) bad.push(`${a}=${v}`); });
    });
    return { lang: currentLang, bad: [...new Set(bad)] };
  });
  check("EN: brak surowych kluczy tłumaczeń w interfejsie", raw.lang === "en" && !raw.bad.length, JSON.stringify(raw));
  const plLeft = await page.evaluate(() => ["frPanel", "reviewPanel", "statsPanel", "exportPanel"].map((k) => document.querySelector(`[data-i18n="${k}"]`)?.textContent));
  check("EN: nowe panele przetłumaczone", plLeft.join("|") === "Find & replace|Review|Statistics|Export", plLeft.join("|"));

  // ── 7. motyw, zoom, skupienie, skróty sekcji ────────────────────────────────
  await page.evaluate(() => document.getElementById("themeToggle").click());
  await page.waitForTimeout(200);
  const theme = await page.evaluate(() => document.documentElement.getAttribute("data-theme"));
  check("motyw: przełącznik zmienia data-theme", !!theme, theme);
  const z0 = await page.textContent("#zoomNow");
  await page.click("#zoomInBtn");
  await page.waitForTimeout(150);
  const z1 = await page.textContent("#zoomNow");
  check("zoom +: procent rośnie", parseInt(z1, 10) > parseInt(z0, 10), `${z0} → ${z1}`);
  await page.click("#focusModeBtn");
  await page.waitForTimeout(300);
  const focus = await page.evaluate(() => document.documentElement.classList.contains("focus-mode") || document.body.classList.contains("focus-mode"));
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  check("tryb skupienia: włącza się przyciskiem", focus);
  const chip = await page.$("#sectionChips button:nth-child(3), #sectionChips .section-chip:nth-child(3)");
  if (chip) {
    // od góry dokumentu — po powrocie na kartę (open-docs.js) widok wraca w miejsce pracy,
    // które bywa dokładnie celem tego skrótu
    await page.evaluate(() => { docViewportEl.scrollTop = 0; });
    await page.waitForTimeout(200);
    const before = await page.evaluate(() => docViewportEl.scrollTop);
    await chip.click();
    await page.waitForTimeout(700);
    const after = await page.evaluate(() => docViewportEl.scrollTop);
    check("skrót sekcji przewija dokument", after !== before, `${before} → ${after}`);
  } else check("skróty sekcji są", false);

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ przepływy użytkownika 2 [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
