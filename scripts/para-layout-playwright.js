// para-layout-playwright.js — odstępy i wcięcia akapitu (2026-10-10).
//
// Jak w Wordzie: przycisk „Odstępy i wcięcia” na pasku Edycji → interlinia 1,0…3,0, Dodaj/Usuń
// odstęp przed/po, Zwiększ/Zmniejsz wcięcie (w liście — poziom punktu), „Odstępy i wcięcia…” =
// formularz jak okno Akapit. Wartości w okienku są efektywne (ustawienia domyślne → styl → lista
// → akapit). Sprawdzamy zapisany XML (w:spacing, w:ind), podgląd i Cofnij. ENGINE=webkit.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

// w:spacing / w:ind / w:numPr akapitu z needle w zapisanym pliku
const savedPPr = (page, needle) => page.evaluate(async (needle) => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  const doc = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const p = [...doc.getElementsByTagNameNS(W, "p")].find((x) => x.textContent.includes(needle));
  if (!p) return null;
  const pPr = [...p.childNodes].find((n) => n.localName === "pPr");
  const attrs = (tag) => { const el = pPr && [...pPr.childNodes].find((n) => n.localName === tag); return el ? Object.fromEntries([...el.attributes].map((a) => [a.localName, a.value])) : null; };
  const ilvl = pPr?.getElementsByTagNameNS(W, "ilvl")[0];
  return { spacing: attrs("spacing"), ind: attrs("ind"), ilvl: ilvl ? ilvl.getAttributeNS(W, "val") || ilvl.getAttribute("w:val") : null };
}, needle);

const caretIn = (page, needle, at = 2) => page.evaluate(({ needle, at }) => {
  const p = [...document.querySelectorAll(".docx-editable-p")].find((x) => x.textContent.includes(needle));
  const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  const n = w.nextNode();
  const r = document.createRange();
  r.setStart(n, Math.min(at, n.length));
  r.collapse(true);
  p.closest(".docx-edit-root").focus({ preventScroll: true });
  getSelection().removeAllRanges();
  getSelection().addRange(r);
}, { needle, at });

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
  await page.keyboard.type("Pierwszy akapit testowy");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Drugi akapit testowy");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Trzeci akapit testowy");
  await idle(page);

  const settle = () => page.evaluate(() => new Promise((res) => {
    const vp = document.getElementById("docViewport");
    let last = vp.scrollTop, still = 0;
    const tick = () => { if (vp.scrollTop === last) still++; else { still = 0; last = vp.scrollTop; } if (still >= 6) res(); else requestAnimationFrame(tick); };
    tick();
  }));
  const openMenu = async () => {
    await settle();
    await page.click("#fmtSpacingBtn");
    await page.waitForSelector(".compose-pop-spacing .compose-item", { timeout: 8000 });
  };
  const pick = async (sel) => {
    const box = await page.evaluate((sel) => { const b = document.querySelector(`.compose-pop-spacing ${sel}`); const r = b?.getBoundingClientRect(); return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; }, sel);
    if (!box) throw new Error(`brak pozycji ${sel}`);
    await page.mouse.click(box.x, box.y);
    await idle(page);
  };

  // ── 1. interlinia z szybkiej listy ──────────────────────────────────────────
  check("przycisk Odstępy i wcięcia na pasku Edycji", await page.evaluate(() => !!document.getElementById("fmtSpacingBtn")?.offsetParent));
  await caretIn(page, "Drugi akapit");
  await openMenu();
  const items = await page.evaluate(() => [...document.querySelectorAll(".compose-pop-spacing .compose-item")].map((b) => b.textContent.trim()));
  check("menu: 6 interlinii, odstępy, wcięcia, „Odstępy i wcięcia…”", items.length === 11 && items[0].startsWith("1,0") && items[1].startsWith("1,15"), items.join(" | "));
  const lhBefore = await page.evaluate(() => getComputedStyle([...document.querySelectorAll(".docx-editable-p")].find((x) => x.textContent.includes("Drugi"))).lineHeight);
  await pick('[data-line="360"]');
  let pp = await savedPPr(page, "Drugi akapit");
  check("1,5 wiersza → w:spacing line=360 lineRule=auto", pp?.spacing?.line === "360" && pp?.spacing?.lineRule === "auto", JSON.stringify(pp));
  const lhAfter = await page.evaluate(() => getComputedStyle([...document.querySelectorAll(".docx-editable-p")].find((x) => x.textContent.includes("Drugi"))).lineHeight);
  check("podgląd: większa interlinia", parseFloat(lhAfter) > parseFloat(lhBefore) || (lhBefore === "normal" && lhAfter !== "normal"), `${lhBefore} → ${lhAfter}`);
  check("inne akapity bez zmian", (await savedPPr(page, "Pierwszy akapit"))?.spacing === null, JSON.stringify(await savedPPr(page, "Pierwszy akapit")));
  await caretIn(page, "Drugi akapit");
  await openMenu();
  check("menu pokazuje bieżącą interlinię (1,5 zaznaczone)", await page.evaluate(() => document.querySelector('.compose-pop-spacing [data-line="360"]')?.getAttribute("aria-checked") === "true"));

  // ── 2. odstęp przed / po ────────────────────────────────────────────────────
  await pick('[data-space="before"]');
  pp = await savedPPr(page, "Drugi akapit");
  check("Dodaj odstęp przed → before=240 (12 pt), interlinia zostaje", pp?.spacing?.before === "240" && pp?.spacing?.line === "360", JSON.stringify(pp));
  await caretIn(page, "Drugi akapit");
  await openMenu();
  const beforeLabel = await page.evaluate(() => document.querySelector('.compose-pop-spacing [data-space="before"]').textContent.trim());
  check("teraz w menu „Usuń odstęp przed akapitem”", beforeLabel === await page.evaluate(() => t("spaceBeforeRemove")), beforeLabel);
  await pick('[data-space="after"]'); // dokument domyślny ma odstęp po (8 pt) → „Usuń”
  pp = await savedPPr(page, "Drugi akapit");
  check("Usuń odstęp po → after=0", pp?.spacing?.after === "0", JSON.stringify(pp));

  // ── 3. wcięcie ──────────────────────────────────────────────────────────────
  await caretIn(page, "Drugi akapit");
  await openMenu();
  await pick('[data-indent="more"]');
  await caretIn(page, "Drugi akapit");
  await openMenu();
  await pick('[data-indent="more"]');
  pp = await savedPPr(page, "Drugi akapit");
  check("2× Zwiększ wcięcie → left=1418 (2,5 cm)", pp?.ind?.left === "1418", JSON.stringify(pp));
  await caretIn(page, "Drugi akapit");
  await openMenu();
  await pick('[data-indent="less"]');
  pp = await savedPPr(page, "Drugi akapit");
  check("Zmniejsz wcięcie → left=709", pp?.ind?.left === "709", JSON.stringify(pp));
  check("podgląd: akapit wcięty", await page.evaluate(() => {
    const ps = [...document.querySelectorAll(".docx-editable-p")];
    const a = ps.find((x) => x.textContent.includes("Pierwszy")).getBoundingClientRect().left;
    const b = ps.find((x) => x.textContent.includes("Drugi")).getBoundingClientRect().left;
    return b - a > 20;
  }));

  // ── 4. formularz „Odstępy i wcięcia…” ──────────────────────────────────────
  await caretIn(page, "Trzeci akapit");
  await openMenu();
  await pick("[data-more]");
  await page.waitForSelector(".compose-pop-layout .lf-left");
  const formNow = await page.evaluate(() => ({ before: document.querySelector(".lf-before").value, after: document.querySelector(".lf-after").value, rule: document.querySelector(".lf-rule").value, at: document.querySelector(".lf-at").value, left: document.querySelector(".lf-left").value }));
  check("formularz: wartości efektywne (domyślne dokumentu: po 8 pt, interlinia 1,08)", formNow.before === "0" && formNow.after === "8" && formNow.rule === "multiple" && formNow.at === "1,08" && formNow.left === "0", JSON.stringify(formNow));
  // zła wartość — ostrzeżenie, nic nie idzie do pliku
  await page.fill(".lf-before", "abc");
  check("zła liczba: ostrzeżenie", await page.evaluate(() => !document.querySelector(".compose-pop-layout .mf-warn").hidden));
  await page.fill(".lf-left", "2");
  await page.fill(".lf-right", "1");
  await page.selectOption(".lf-special", "first");
  await page.fill(".lf-by", "1,25");
  await page.fill(".lf-before", "6");
  await page.fill(".lf-after", "3");
  await page.selectOption(".lf-rule", "exact");
  await page.fill(".lf-at", "18");
  await page.click(".compose-pop-layout .lf-ok");
  await idle(page);
  pp = await savedPPr(page, "Trzeci akapit");
  check("formularz → ind left=1134 right=567 firstLine=709", pp?.ind?.left === "1134" && pp?.ind?.right === "567" && pp?.ind?.firstLine === "709" && !pp?.ind?.hanging, JSON.stringify(pp));
  check("formularz → spacing before=120 after=60 line=360 exact", pp?.spacing?.before === "120" && pp?.spacing?.after === "60" && pp?.spacing?.line === "360" && pp?.spacing?.lineRule === "exact", JSON.stringify(pp));
  // wysunięcie zastępuje pierwszy wiersz
  await caretIn(page, "Trzeci akapit");
  await openMenu();
  await pick("[data-more]");
  await page.waitForSelector(".compose-pop-layout .lf-left");
  check("formularz pokazuje zapisane: Pierwszy wiersz 1,25 / Dokładnie 18", await page.evaluate(() => document.querySelector(".lf-special").value === "first" && document.querySelector(".lf-by").value === "1,25" && document.querySelector(".lf-rule").value === "exact" && document.querySelector(".lf-at").value === "18"));
  await page.selectOption(".lf-special", "hanging");
  await page.fill(".lf-by", "0,5");
  await page.click(".compose-pop-layout .lf-ok");
  await idle(page);
  pp = await savedPPr(page, "Trzeci akapit");
  check("Wysunięcie → hanging=283, bez firstLine", pp?.ind?.hanging === "283" && pp?.ind?.firstLine === undefined, JSON.stringify(pp));

  // ── 5. kilka akapitów naraz + Cofnij ────────────────────────────────────────
  await page.evaluate(() => {
    const ps = [...document.querySelectorAll(".docx-editable-p")];
    const a = ps.find((x) => x.textContent.includes("Pierwszy")), b = ps.find((x) => x.textContent.includes("Drugi"));
    const r = document.createRange();
    r.setStart(a.firstChild?.firstChild || a.firstChild, 1);
    const w = document.createTreeWalker(b, NodeFilter.SHOW_TEXT); const n = w.nextNode();
    r.setEnd(n, 3);
    a.closest(".docx-edit-root").focus({ preventScroll: true });
    getSelection().removeAllRanges(); getSelection().addRange(r);
  });
  await openMenu();
  await pick('[data-line="480"]');
  const p1 = await savedPPr(page, "Pierwszy akapit"), p2 = await savedPPr(page, "Drugi akapit");
  check("zaznaczenie 2 akapitów → oba podwójna interlinia (480)", p1?.spacing?.line === "480" && p2?.spacing?.line === "480", JSON.stringify({ p1, p2 }));
  await settle();
  await page.click("#undoBtn");
  await idle(page);
  check("Cofnij → interlinia 1,5 wraca w drugim akapicie", (await savedPPr(page, "Drugi akapit"))?.spacing?.line === "360", JSON.stringify(await savedPPr(page, "Drugi akapit")));

  // ── 6. lista: Zwiększ wcięcie = poziom punktu ───────────────────────────────
  await caretIn(page, "Pierwszy akapit");
  await page.evaluate(() => composeUi.applyList ? composeUi.applyList("bullet") : null);
  if (!(await page.evaluate(() => typeof composeUi.applyList === "function"))) {
    await page.click("#fmtListBtn");
    await page.waitForSelector(".compose-pop .compose-item");
    await page.evaluate(() => [...document.querySelectorAll(".compose-pop .compose-item")][0].click());
  }
  await idle(page);
  await caretIn(page, "Pierwszy akapit");
  await openMenu();
  await pick('[data-indent="more"]');
  pp = await savedPPr(page, "Pierwszy akapit");
  check("lista: Zwiększ wcięcie → poziom 1 (ilvl=1)", pp?.ilvl === "1", JSON.stringify(pp));

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ odstępy i wcięcia [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
