// a11y-playwright.js — „Sprawdź ułatwienia dostępu” w panelu Korekta (2026-10-10).
//
// Dokument z każdym rodzajem problemu (jak w Wordzie): obraz bez opisu, tabela bez wiersza
// nagłówka, Nagłówek 1 → Nagłówek 3, link „kliknij tutaj”, jasnoszary tekst na białym, puste
// akapity z rzędu, brak tytułu. Sprawdzamy wykrycie, skok do miejsca i szybkie poprawki
// (nagłówek tabeli, „zostaw jeden”, opis obrazu, tytuł w Metadanych). ENGINE=webkit.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));
const paraIdx = (page, needle) => page.evaluate((needle) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).findIndex((p) => p.textContent.includes(needle)), needle);

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

  // ── dokument z problemami ───────────────────────────────────────────────────
  await page.keyboard.type("Rozdział pierwszy");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Podrozdział za głęboko");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Więcej informacji: kliknij tutaj. Jasny tekst tutaj.");
  await page.keyboard.press("Enter");
  await page.keyboard.press("Enter");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Po pustych akapitach");
  await idle(page);
  await page.evaluate(async () => {
    const ps = () => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const idx = (n) => ps().findIndex((p) => p.textContent.includes(n));
    await applyDocumentEdit({ op: "paraFormat", indices: [idx("Rozdział")], style: "h1" });
    await applyDocumentEdit({ op: "paraFormat", indices: [idx("Podrozdział")], style: "h3" });
    const i = idx("kliknij tutaj");
    const txt = ps()[i].textContent;
    await applyDocumentEdit({ op: "link", index: i, start: txt.indexOf("kliknij tutaj"), end: txt.indexOf("kliknij tutaj") + "kliknij tutaj".length, href: "https://example.com" });
    const cv = document.createElement("canvas"); cv.width = 40; cv.height = 30;
    const cx = cv.getContext("2d"); cx.fillStyle = "#3a7"; cx.fillRect(0, 0, 40, 30);
    const blob = await new Promise((r) => cv.toBlob(r, "image/png"));
    await applyDocumentEdit({ op: "imageInsert", index: idx("Po pustych"), bytes: new Uint8Array(await blob.arrayBuffer()), ext: "png", mime: "image/png", width: 40, height: 30, descr: "" });
  });
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);
  // jasnoszary tekst
  await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => x.textContent.includes("Jasny tekst"));
    const n = [...(function* w(e) { const tw = document.createTreeWalker(e, NodeFilter.SHOW_TEXT); for (let x = tw.nextNode(); x; x = tw.nextNode()) yield x; })(p)].find((x) => x.data.includes("Jasny"));
    const r = document.createRange(); const at = n.data.indexOf("Jasny tekst");
    r.setStart(n, at); r.setEnd(n, at + "Jasny tekst".length);
    p.closest(".docx-edit-root").focus({ preventScroll: true });
    getSelection().removeAllRanges(); getSelection().addRange(r);
  });
  await page.evaluate(() => composeUi.applyColor("color", "E8E8E8"));
  await idle(page);
  // tabela 3×3
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => x.textContent.includes("Po pustych")); focusParagraphAtOffset(collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).indexOf(p), 5); });
  await page.evaluate(() => composeUi.insertTable(3, 3));
  await idle(page);
  for (const w of ["Kol A", "Kol B", "Kol C", "1", "2", "3"]) { await page.keyboard.type(w); await page.keyboard.press("Tab"); await page.waitForTimeout(80); }
  await idle(page);

  // ── sprawdzenie ─────────────────────────────────────────────────────────────
  await page.evaluate(async () => { setSidebarOpen(true); document.getElementById("panel-grammar").open = true; await ensureLazyFeature("grammar"); });
  const scan = async () => {
    await page.click("#a11yScanBtn");
    await page.waitForFunction(() => /\d|✓/.test(document.getElementById("a11yStatus").textContent) && !/…$/.test(document.getElementById("a11yStatus").textContent), null, { timeout: 15000 });
    return page.evaluate(() => Object.fromEntries([...document.querySelectorAll("#a11yResults .a11y-group")].map((d) => [d.dataset.rule, d.querySelectorAll(".grammar-hit-item").length])));
  };
  let found = await scan();
  check("wykryto: obraz bez opisu", found.imageAlt === 1, JSON.stringify(found));
  check("wykryto: tabela bez wiersza nagłówka", found.tableHeader === 1, JSON.stringify(found));
  check("wykryto: pominięty poziom nagłówka (1 → 3)", found.headingSkip === 1, JSON.stringify(found));
  check("wykryto: niejasny link „kliknij tutaj”", found.linkText === 1, JSON.stringify(found));
  check("wykryto: słaby kontrast (jasnoszary)", found.contrast >= 1, JSON.stringify(found));
  check("wykryto: puste akapity z rzędu", found.emptyParas === 1, JSON.stringify(found));
  check("wykryto: brak tytułu dokumentu", found.docTitle === 1, JSON.stringify(found));
  check("poziomy: Błąd / Ostrzeżenie / Wskazówka", await page.evaluate(() => !!document.querySelector('.a11y-error[data-rule="imageAlt"]') && !!document.querySelector('.a11y-warning[data-rule="tableHeader"]') && !!document.querySelector('.a11y-tip[data-rule="emptyParas"]')));

  // skok do miejsca
  await page.evaluate(() => document.querySelector('[data-rule="headingSkip"] .grammar-hit-item').click());
  await page.waitForTimeout(400);
  check("klik w uwagę podświetla miejsce w dokumencie", await page.evaluate(() => document.querySelector("#docCanvas .search-hit-active")?.textContent.includes("Podrozdział")));

  // ── szybkie poprawki ────────────────────────────────────────────────────────
  await page.evaluate(() => document.querySelector('[data-rule="tableHeader"] [data-fix="header"]').click());
  await idle(page);
  await page.waitForFunction(() => !/…$/.test(document.getElementById("a11yStatus").textContent), null, { timeout: 15000 });
  found = await page.evaluate(() => Object.fromEntries([...document.querySelectorAll("#a11yResults .a11y-group")].map((d) => [d.dataset.rule, d.querySelectorAll(".grammar-hit-item").length])));
  check("„Ustaw nagłówek”: tabela znika z listy", !found.tableHeader, JSON.stringify(found));
  const before = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).length);
  await page.evaluate(() => document.querySelector('[data-rule="emptyParas"] [data-fix="empty"]').click());
  await idle(page);
  await page.waitForFunction(() => !/…$/.test(document.getElementById("a11yStatus").textContent), null, { timeout: 15000 });
  const after = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).length);
  found = await page.evaluate(() => Object.fromEntries([...document.querySelectorAll("#a11yResults .a11y-group")].map((d) => [d.dataset.rule, d.querySelectorAll(".grammar-hit-item").length])));
  check("„Zostaw jeden”: z 2 pustych akapitów zostaje 1, uwaga znika", after === before - 1 && !found.emptyParas, JSON.stringify({ before, after, found }));
  // opis obrazu — okienko „Tekst alternatywny”
  await page.evaluate(() => document.querySelector('[data-rule="imageAlt"] [data-fix="alt"]').click());
  await page.waitForSelector(".compose-pop-form textarea", { timeout: 8000 });
  await page.fill(".compose-pop-form textarea", "Zielony prostokąt");
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
  found = await scan();
  check("„Dodaj opis” → okienko, po zapisie obraz bez uwagi", !found.imageAlt, JSON.stringify(found));
  // tytuł — Metadane
  await page.evaluate(() => document.querySelector('[data-rule="docTitle"] [data-fix="title"]').click());
  await page.waitForFunction(() => document.activeElement?.id === "metaTitle", null, { timeout: 5000 }).catch(() => {});
  check("„Uzupełnij” tytuł: otwiera Metadane z kursorem w polu Tytuł", await page.evaluate(() => document.getElementById("panel-metadata").open && document.activeElement?.id === "metaTitle"));
  await page.evaluate(async () => { await ensureLazyFeature("metadata"); });
  await page.fill("#metaTitle", "Dokument testowy");
  await page.click("#metaApplyBtn");
  await idle(page);
  found = await scan();
  check("po nadaniu tytułu — bez uwagi o tytule", !found.docTitle, JSON.stringify(found));

  // czysty dokument
  await page.evaluate(() => loadSampleDocument("headings-sample"));
  await idle(page);
  await page.evaluate(async () => { setSidebarOpen(true); document.getElementById("panel-grammar").open = true; });
  found = await scan();
  check("przykład z nagłówkami: bez błędów (tylko ewentualne wskazówki)", !found.imageAlt && !found.headingSkip && !found.contrast && !found.linkText, JSON.stringify(found));

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ ułatwienia dostępu [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
