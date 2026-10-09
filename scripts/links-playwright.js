// links-playwright.js — linki w dokumencie (app/doc-links.js) + obrys akapitu z wysuniętym
// pierwszym wierszem (fixHangingBox w docx-inline-edit.js).
//
// Fixture: docs/samples/links-sample.docx (node scripts/gen-links-docx.mjs). Klikamy jak
// użytkownik i sprawdzamy: skok w dokumencie bez zmiany adresu, „↩ Wróć”, adres WWW w nowej
// karcie (aplikacja zostaje), javascript: zablokowany, odsyłacze-pola klikalne, plik bez zmian.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(300));

// Czy element (akapit z danym tekstem) jest widoczny w obszarze dokumentu.
const inView = (page, needle) => page.evaluate((n) => {
  const vp = docViewportEl.getBoundingClientRect();
  const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => x.textContent === n);
  const r = p.getBoundingClientRect();
  return r.top >= vp.top - 2 && r.bottom <= vp.bottom + 2;
}, needle);
const clickLink = async (page, selector) => {
  await page.evaluate((s) => document.querySelector(s).scrollIntoView({ block: "center" }), selector);
  await page.waitForTimeout(150);
  await page.click(selector);
  await page.waitForTimeout(900); // płynne przewijanie
};

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 800 } });
  await context.addInitScript(() => {
    sessionStorage.setItem("introPlayed", "true");
    window.__opened = [];
    window.open = (...args) => { window.__opened.push(args); return null; };
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => { errors.push(`dialog: ${d.message()}`); d.dismiss(); });
  await page.goto(`${APP_URL}?sample=links-sample`, { waitUntil: "load" });
  await page.waitForSelector(".docx-preview-host a[href]", { timeout: 20000 });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await idle(page);
  const url0 = page.url();
  // stan PRZED poprawką ramki (zanim mysz/fokus dotknie akapitu)
  const hang0 = await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => x.textContent.startsWith("Nagłe zdarzenia 15-17"));
    const r = document.createRange(); const tn = document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode(); r.setStart(tn, 0); r.setEnd(tn, 1);
    return { box: p.getBoundingClientRect().left, char: r.getBoundingClientRect().left, hang: p.dataset.hang || "" };
  });

  // ── spis treści → rozdział, „Wróć” ─────────────────────────────────────────
  // miejsce sprzed skoku = po przewinięciu do linku (clickLink przewija go na środek)
  await page.evaluate(() => document.querySelector('.docx-preview-host a[href="#_Toc3"]').scrollIntoView({ block: "center" }));
  await page.waitForTimeout(150);
  const top0 = await page.evaluate(() => docViewportEl.scrollTop);
  await clickLink(page, '.docx-preview-host a[href="#_Toc3"]');
  check("spis treści: klik przewija do rozdziału „Zagrożenia bombowe”", await inView(page, "Zagrożenia bombowe"));
  check("adres strony bez zmian (bez #_Toc…)", page.url() === url0, page.url());
  const back = await page.evaluate(() => { const b = document.querySelector(".link-back"); return b && !b.hidden ? b.textContent : null; });
  check("pojawia się „↩ Wróć”", /Wróć/.test(back || ""), back);
  await page.keyboard.press("Alt+ArrowLeft");
  // płynne przewijanie z powrotem — dłuższe, gdy strony mają przerwy z marginesami
  await page.waitForFunction((t0) => Math.abs(docViewportEl.scrollTop - t0) < 30, top0, { timeout: 4000 }).catch(() => {});
  await page.waitForTimeout(150);
  const top1 = await page.evaluate(() => docViewportEl.scrollTop);
  check("Alt+← wraca do miejsca sprzed skoku i chowa „Wróć”", Math.abs(top1 - top0) < 30 && await page.evaluate(() => document.querySelector(".link-back").hidden), `${top0} → ${top1}`);

  // ── „Wróć” znika samo, gdy przewiniesz z powrotem w okolice miejsca sprzed skoku ──
  await page.evaluate(() => document.querySelector('.docx-preview-host a[href="#_Toc3"]').scrollIntoView({ block: "center" }));
  await page.waitForTimeout(150);
  const top2 = await page.evaluate(() => docViewportEl.scrollTop);
  await clickLink(page, '.docx-preview-host a[href="#_Toc3"]');
  const shown = await page.evaluate(() => !document.querySelector(".link-back").hidden);
  await page.evaluate((t) => { docViewportEl.scrollTop = t + 30; }, top2); // ręcznie, „mniej więcej” tam
  await page.waitForFunction(() => document.querySelector(".link-back").hidden, null, { timeout: 2000 }).catch(() => {});
  check("ręczny powrót w okolice miejsca sprzed skoku chowa „Wróć”", shown && await page.evaluate(() => document.querySelector(".link-back").hidden));
  // po samym skoku (bez powrotu) „Wróć” zostaje
  await clickLink(page, '.docx-preview-host a[href="#_Toc3"]');
  await page.waitForTimeout(300);
  check("bez powrotu „Wróć” zostaje na ekranie", await page.evaluate(() => !document.querySelector(".link-back").hidden && !document.querySelector(".link-back").classList.contains("is-leaving")));
  await page.click(".link-back");
  await page.waitForTimeout(900);

  // ── odsyłacze-pola ─────────────────────────────────────────────────────────
  await clickLink(page, '.docx-preview-host a.doc-xref[href="#_Ref9"]');
  check("odsyłacz REF \\h: klik przewija do celu", await inView(page, "Zagrożenia bombowe"));
  await page.click(".link-back");
  await page.waitForTimeout(900);
  await clickLink(page, '.docx-preview-host a.doc-xref[href="#_Toc2"]');
  check("pole HYPERLINK \\l: klik przewija do „Nagłe zdarzenia”", await inView(page, "Nagłe zdarzenia"));

  // ── zewnętrzne / niebezpieczne / zepsute ───────────────────────────────────
  await clickLink(page, '.docx-preview-host a[href^="https://"]');
  let opened = await page.evaluate(() => window.__opened.slice());
  check("adres WWW: nowa karta (noopener), aplikacja zostaje", opened.length === 1 && opened[0][0] === "https://www.example.com/" && /noopener/.test(opened[0][2]) && page.url() === url0, JSON.stringify(opened));
  await clickLink(page, '.docx-preview-host a[href^="javascript:"]');
  opened = await page.evaluate(() => window.__opened.length);
  const toastJs = await page.$$eval(".toast", (els) => els.map((e) => e.textContent).join(" | "));
  check("javascript: zablokowany z komunikatem", opened === 1 && /nie da się otworzyć/.test(toastJs) && page.url() === url0, toastJs);
  await clickLink(page, '.docx-preview-host a[href="#_TocBrak"]');
  const toastMissing = await page.$$eval(".toast", (els) => els.map((e) => e.textContent).join(" | "));
  check("link do brakującej zakładki: komunikat zamiast niczego", /nie ma w dokumencie/.test(toastMissing), toastMissing);

  // ── w trybie Edycja też działa ─────────────────────────────────────────────
  await page.click('.mode-btn[data-mode="edit"]');
  await page.evaluate(() => docViewportEl.scrollTo({ top: 0 }));
  await clickLink(page, '.docx-preview-host a[href="#_Toc1"]');
  check("Edycja: klik w spis treści też skacze", await inView(page, "Bezpieczeństwo"));
  // Ctrl/⌘+klik w link w edytowalnym akapicie = od razu otwarcie, bez karty linku pod kursorem
  {
    const MOD = process.platform === "darwin" ? "Meta" : "Control";
    const sel = '.docx-preview-host .docx-editable-p a[href^="https://"]';
    const n0 = await page.evaluate(() => window.__opened.length);
    await page.evaluate((s) => document.querySelector(s).scrollIntoView({ block: "center" }), sel);
    await page.waitForTimeout(150);
    await page.click(sel, { modifiers: [MOD] });
    await page.waitForTimeout(400);
    const r = await page.evaluate((n0) => ({ opened: window.__opened.length - n0, card: !!document.querySelector(".link-card") }), n0);
    check("Edycja: Ctrl/⌘+klik w link otwiera od razu, bez karty linku", r.opened === 1 && !r.card, JSON.stringify(r));
    await page.click(sel);
    await page.waitForTimeout(400);
    const r2 = await page.evaluate((n0) => ({ opened: window.__opened.length - n0, card: !!document.querySelector(".link-card") }), n0);
    check("Edycja: zwykły klik w ten link dalej pokazuje kartę (bez otwierania)", r2.opened === 1 && r2.card, JSON.stringify(r2));
  }

  // ── akapit z wysuniętym pierwszym wierszem (zgłoszenie: „N|agłe zdarzenia”) ─
  const hang = await page.evaluate(async () => {
    docViewportEl.scrollTo({ top: 0 });
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => x.textContent.startsWith("Nagłe zdarzenia 15-17"));
    const firstChar = () => { const r = document.createRange(); const tn = document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode(); r.setStart(tn, 0); r.setEnd(tn, 1); return r.getBoundingClientRect().left; };
    const charBefore = firstChar();
    const boxBefore = p.getBoundingClientRect().left;
    placeCaret(p, 3);
    await new Promise((r) => setTimeout(r, 200));
    return { editable: p.isContentEditable, charBefore, charAfter: firstChar(), boxBefore, boxAfter: p.getBoundingClientRect().left, hang: p.dataset.hang };
  });
  check("przed poprawką ramka zaczynała się za pierwszą literą (warunek testu)", !hang0.hang && hang0.box > hang0.char + 5, JSON.stringify(hang0));
  check("po fokusie podświetlenie obejmuje cały 1. wiersz (ramka ≤ pierwsza litera)", hang.editable && hang.hang === "1" && hang.boxAfter <= hang.charAfter + 1, JSON.stringify(hang));
  check("tekst nie przesunął się ani o piksel", Math.abs(hang.charAfter - hang0.char) < 0.5, JSON.stringify({ hang0, hang }));

  // ── nic z tego nie trafia do pliku ─────────────────────────────────────────
  const same = await page.evaluate(async () => { inlineDirtyValid = false; const n = collectInlineParagraphEdits().length; inlineDirtyValid = true; return { n, same: (await buildDocumentForSave()) === originalFileBytes }; });
  check("linki, odsyłacze i poprawka ramki nie są zmianą pliku", same.n === 0 && same.same, JSON.stringify(same));

  await browser.close();
  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ linki [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
