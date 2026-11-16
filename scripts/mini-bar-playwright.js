// mini-bar-playwright.js — mini pasek przy zaznaczeniu (app/mini-bar.js, etap 5 wstążki).
//
// Zaznaczenie myszą w Edycji → pasek nad zaznaczeniem (pod nim, gdy u góry brak miejsca);
// B działa na zaznaczeniu i pasek pokazuje „wciśnięte”; krój z listy paska trafia do tekstu;
// kolor i styl otwierają menu przy pasku; przewijanie gasi pasek i stawia go na nowym miejscu;
// klawisz, klik w dokument i Czytanie go chowają; wysuwany panel go nie przykrywa.
// ENGINE=webkit (Safari/iPad z myszą).

const pw = require("playwright");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const errors = [];

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 800 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", path.join(__dirname, "..", "docs", "samples", "przewodnik.docx"));
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  await sleep(500);
  await page.evaluate(() => document.querySelector('.mode-btn[data-mode="edit"]').click());
  await page.waitForFunction(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).some((p) => p.isContentEditable));
  await sleep(300);

  const state = () => page.evaluate(() => {
    const r = dwbMiniBar.el.getBoundingClientRect();
    const sel = getSelection();
    const sr = sel.rangeCount && !sel.isCollapsed ? sel.getRangeAt(0).getClientRects()[0] : null;
    return { shown: dwbMiniBar.isShown(), text: sel.toString(), top: r.top, bottom: r.bottom, left: r.left, right: r.right, selTop: sr?.top ?? null, selBottom: sr?.bottom ?? null };
  });
  // zwykły akapit z kilkoma wyrazami (nie link), w środku ekranu
  const mark = () => page.evaluate(() => {
    if (document.querySelector("[data-mb-test]")) return;
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => {
      if (!p.isContentEditable || p.closest("table") || p.textContent.length < 90) return false;
      const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      const first = w.nextNode();
      const el = first?.parentElement;
      return el && !el.closest("a") && (parseInt(getComputedStyle(el).fontWeight, 10) || 400) < 600 && first.textContent.trim().length > 25;
    });
    const any = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.textContent.length > 60 && !p.closest("table"));
    (ps[1] || any[1]).dataset.mbTest = "1";
  });
  const para = page.locator("[data-mb-test]");
  async function dragSelect() {
    await mark(); // dokument przerysowany (np. inna szerokość) — akapit od nowa
    await para.evaluate((el) => el.scrollIntoView({ block: "center" }));
    await sleep(250);
    const bb = await para.boundingBox();
    const y = bb.y + 10;
    // najpierw zwykły klik (zwija poprzednie zaznaczenie — przeciąganie wewnątrz niego przenosiłoby tekst)
    await page.mouse.click(bb.x + bb.width - 30, bb.y + bb.height - 6);
    await sleep(150);
    await page.mouse.move(bb.x + 4, y);
    await page.mouse.down();
    await page.mouse.move(bb.x + 140, y, { steps: 8 });
    await page.mouse.up();
    await sleep(300);
  }

  // ── 1. zaznaczenie myszą → pasek nad zaznaczeniem ─────────────────────────
  await dragSelect();
  let st = await state();
  check("zaznaczenie myszą w Edycji: pasek nad zaznaczeniem", st.shown && st.text.length > 3 && st.bottom <= st.selTop && st.selTop - st.bottom < 20, JSON.stringify(st));
  const items = await page.evaluate(() => Array.from(dwbMiniBar.el.querySelectorAll("select, button")).map((b) => b.getAttribute("aria-label") || ""));
  check("pasek: krój, rozmiar, B, I, U, wyróżnienie, kolor, punktory, styl, komentarz, Przepisz (każdy z nazwą)", items.length === 11 && items.every(Boolean), JSON.stringify(items));

  // ── 2. B na zaznaczeniu ───────────────────────────────────────────────────
  const before = st.text;
  await page.click(".mini-bar .mini-btn >> nth=0");
  await sleep(400);
  const bold = await page.evaluate(() => ({
    pressed: document.querySelector(".mini-bar .mini-btn").getAttribute("aria-pressed"),
    src: document.getElementById("fmtBold").classList.contains("is-on"),
    text: getSelection().toString(),
    weight: getComputedStyle(getSelection().getRangeAt(0).startContainer.parentElement).fontWeight,
  }));
  check("B z paska: zaznaczenie pogrubione, zostaje zaznaczone, przycisk wciśnięty jak na wstążce", bold.pressed === "true" && bold.src && bold.text === before && Number(bold.weight) >= 600, JSON.stringify(bold));

  // ── 3. krój z listy paska → ta sama obsługa co lista na wstążce ──────────
  const fam = await page.evaluate(() => {
    const s = dwbMiniBar.el.querySelector(".mini-font");
    const opt = Array.from(s.options).find((o) => o.value && o.value !== "__custom" && o.value !== s.value);
    return opt?.value;
  });
  await page.selectOption(".mini-bar .mini-font", fam);
  await sleep(500);
  const famNow = await page.evaluate(() => ({ ribbon: document.getElementById("fmtFontFamily").value, text: getSelection().toString(), css: getComputedStyle(getSelection().getRangeAt(0).startContainer.parentElement).fontFamily }));
  check("krój z paska: trafia do zaznaczenia i na wstążkę", famNow.ribbon === fam && famNow.css.includes(fam) && famNow.text === before, JSON.stringify({ fam, famNow }));

  // ── 4. menu koloru i stylu przy pasku ─────────────────────────────────────
  await dragSelect();
  await page.click(".mini-bar .color-btn");
  await sleep(300);
  const pop = await page.evaluate(() => {
    const p = document.querySelector(".compose-pop-color");
    const b = document.querySelector(".mini-bar .color-btn").getBoundingClientRect();
    const r = p?.getBoundingClientRect();
    return { open: !!p && !!r.width, near: !!r && Math.abs(r.left - b.left) < 260 && Math.abs(r.top - b.bottom) < 80 };
  });
  check("kolor: menu kolorów otwiera się przy przycisku paska", pop.open && pop.near, JSON.stringify(pop));
  await page.keyboard.press("Escape");
  await sleep(200);
  await dragSelect();
  await page.click('.mini-bar .mini-btn[aria-haspopup="menu"]:not(.color-btn)');
  await sleep(300);
  const styles = await page.evaluate(() => Array.from(document.querySelectorAll(".compose-pop-insert [role=menuitemradio]")).map((b) => b.textContent.trim()));
  check("styl: menu stylów akapitu (te same co na wstążce)", styles.length >= 3 && styles.some((s) => /Normal|Nagłówek|Heading/i.test(s)), JSON.stringify(styles.slice(0, 6)));
  await page.keyboard.press("Escape");
  await sleep(200);

  // ── 5. przewijanie: gaśnie, wraca na nowym miejscu ───────────────────────
  await dragSelect();
  st = await state();
  await page.mouse.move(640, 300);
  await page.mouse.wheel(0, 120);
  await sleep(60);
  const during = await state();
  await sleep(500);
  const after = await state();
  check("przewijanie: pasek gaśnie, po zatrzymaniu wraca nad zaznaczenie", !during.shown && after.shown && Math.abs(after.top - st.top) > 40 && after.bottom <= after.selTop, JSON.stringify({ st, during, after }));

  // ── 6. brak miejsca nad zaznaczeniem → pod nim ────────────────────────────
  await page.evaluate(() => {
    const vp = document.getElementById("docViewport");
    const sel = getSelection().getRangeAt(0).getBoundingClientRect();
    vp.scrollTop += sel.top - vp.getBoundingClientRect().top - 6;
  });
  await sleep(500);
  st = await state();
  check("zaznaczenie przy górnej krawędzi: pasek pod zaznaczeniem", st.shown && st.top >= st.selBottom, JSON.stringify(st));

  // ── 7. chowanie ───────────────────────────────────────────────────────────
  await page.keyboard.press("ArrowRight");
  await sleep(200);
  check("klawisz: pasek znika", !(await state()).shown);
  await dragSelect();
  const bb = await para.boundingBox();
  await page.mouse.click(bb.x + 20, bb.y + bb.height - 4);
  await sleep(300);
  check("klik w dokument (bez zaznaczenia): pasek znika", !(await state()).shown);
  // zaznaczenie z klawiatury — bez paska (jak w Wordzie)
  await page.keyboard.press("Shift+End");
  await sleep(300);
  check("zaznaczenie klawiaturą: bez paska", !(await state()).shown);

  // ── 8. wysuwany panel: pasek nie wchodzi na niego ────────────────────────
  await dragSelect();
  await page.evaluate(() => dwbShortcuts.openToolPanel("panel-structure"));
  await sleep(500);
  await page.evaluate(() => dwbMiniBar.place());
  const cover = await page.evaluate(() => {
    const p = document.getElementById("toolsPanel").getBoundingClientRect();
    const r = dwbMiniBar.el.getBoundingClientRect();
    return { shown: dwbMiniBar.isShown(), barLeft: r.left, panelRight: p.right };
  });
  check("wysuwany panel: pasek stoi obok niego, nie na nim", !cover.shown || cover.barLeft >= cover.panelRight, JSON.stringify(cover));
  await page.keyboard.press("Escape");

  // ── 9. Czytanie: bez paska ────────────────────────────────────────────────
  await page.evaluate(() => document.querySelector('.mode-btn[data-mode="read"]').click());
  await sleep(400);
  await dragSelect();
  check("Czytanie: zaznaczenie bez paska", !(await state()).shown);

  // ── 10. telefon (bez wstążki): bez paska ─────────────────────────────────
  await page.evaluate(() => document.querySelector('.mode-btn[data-mode="edit"]').click());
  await page.setViewportSize({ width: 600, height: 800 });
  await sleep(500);
  await dragSelect();
  check("bez wstążki (wąsko): bez paska", !(await state()).shown);

  await browser.close();
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Mini pasek [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
