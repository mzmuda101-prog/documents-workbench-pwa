// @ts-check
// toc-pages-playwright.js — spis treści: numery stron jak na stronach Edycji (i wydruku), także po
// „Aktualizuj spis”. Dawniej przy każdym wpisie było „1”: strona liczona z min-height kartki,
// który przy odstępach stron jest sumą wszystkich stron (testerzy AI 2026-11-27, AUDYT-MOBILE #6).
//   node scripts/toc-pages-playwright.js        ENGINE=webkit node scripts/toc-pages-playwright.js
const pw = require("playwright");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending && !inlineStructuralPending, null, { timeout: 30000 }).then(() => page.waitForTimeout(1200));

(async () => {
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles(path.join(__dirname, "..", "docs", "samples", "headings-sample.docx"));
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);
  // nagłówek 3. poziomu (Ctrl+Alt+3, jak tester): jego wpis w spisie ma numer w tej samej linii
  await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((q) => q.textContent.trim() === "Dane Najemcy");
    p.closest(".docx-edit-root").focus();
    placeCaret(p, 2);
    document.dispatchEvent(new Event("selectionchange"));
  });
  await page.keyboard.press("Control+Alt+3");
  await idle(page);
  // kursor na początku dokumentu → Wstaw → Spis treści, potem Aktualizuj (drugie wywołanie)
  const toc = async () => {
    await page.evaluate(() => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0];
      p.closest(".docx-edit-root").focus();
      placeCaret(p, 0);
    });
    await page.evaluate(() => composeUi.insertToc());
    await idle(page);
    return page.evaluate(() => {
      const host = document.querySelector(".docx-preview-host");
      const sec = host.querySelector("section.docx");
      const scale = sec.getBoundingClientRect().height / sec.offsetHeight || 1;
      const bands = [...sec.querySelectorAll(":scope > .dwb-page-gap-band")].map((b) => parseFloat(b.style.top));
      const pageAt = (el) => { const y = (el.getBoundingClientRect().top - sec.getBoundingClientRect().top) / scale; return 1 + bands.filter((t) => t < y).length; };
      const ps = collectPreviewParagraphElements(host);
      const entries = ps.filter((p) => /^\s*\S.*\d+\s*$/.test(p.textContent) && [...p.classList].some((c) => COMPOSE_TOC_KEYS.includes(docComposeStyleClasses.get(c))))
        .map((p) => ({ title: p.textContent.replace(/[.\s ]*\d+\s*$/, "").trim(), n: Number(p.textContent.match(/(\d+)\s*$/)[1]) }));
      const heads = (documentStructure?.headings || []).map((h) => ({ title: h.el.textContent.trim(), page: pageAt(h.el) }));
      const tocPs = ps.filter((p) => /docx_toc[1-9]/.test(p.className));
      const lines = tocPs.map((p) => ({ cls: p.className.split(" ")[0], h: Math.round(p.getBoundingClientRect().height) }));
      return { pages: bands.length + 1, entries, heads, lines };
    });
  };
  const first = await toc();
  const second = await toc(); // „Aktualizuj spis”
  for (const [label, r] of [["wstawienie", first], ["aktualizacja", second]]) {
    const wrong = r.entries.filter((e) => { const h = r.heads.find((x) => x.title === e.title); return !h || h.page !== e.n; });
    check(`${label}: dokument ma ≥ 2 strony, spis ma wpisy`, r.pages >= 2 && r.entries.length >= 4, JSON.stringify({ pages: r.pages, n: r.entries.length }));
    check(`${label}: numer przy każdym wpisie = strona nagłówka w Edycji`, r.entries.length && !wrong.length, JSON.stringify({ wrong, heads: r.heads }));
    check(`${label}: nie same „1”`, r.entries.some((e) => e.n > 1), JSON.stringify(r.entries.map((e) => e.n)));
    check(`${label}: każdy wpis (też 3. poziomu) w jednej linii, numer bez zawijania`, r.lines.some((l) => l.cls === "docx_toc3") && r.lines.every((l) => l.h <= r.lines[0].h + 2), JSON.stringify(r.lines));
  }
  check("bez błędów na stronie", !errors.length, errors.join(" | "));
  await browser.close();
  const bad = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  ${String(r.detail).slice(0, 500)}`}`));
  console.log(`${bad.length ? "❌" : "✅"} Spis treści — numery stron [${ENGINE}]: ${results.length - bad.length}/${results.length}`);
  process.exit(bad.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
