// review-margin-playwright.js — znaczniki recenzji jak w Wordzie na Macu (app/review-margin.js)
// i uchwyty rozmiaru obrazu z przyciąganiem (compose-ui.js).
//
// Dymki obok kartki: komentarze i zmiany, bez nachodzenia na siebie; wstawki podkreślone w tekście
// (CSS Highlight na DOKŁADNIE wstawionym tekście); kursor w komentarzu = aktywny dymek przy swoim
// wierszu zamiast karty nad tekstem; ✓ / Rozwiąż z dymka zmieniają PLIK (i Cofnij wraca); tryb
// „Bez” i wąski ekran — bez dymków, wraca karta. Śledź zmiany: wpisany tekst podkreślony od razu,
// usunięte znaki w dymku. Obraz: przeciągnięcie rogu „staje” na 50% szerokości tekstu.
// ENGINE=webkit (Safari).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(700));
const fileXml = (page, part) => page.evaluate(async (part) => { const z = await JSZip.loadAsync(originalFileBytes); return (await z.file(part)?.async("string")) || ""; }, part);
const marks = (x) => (x.match(/<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b/g) || []).length;
const items = (page) => page.evaluate(() => dwbMargin._items());
const visibleCards = (page) => page.evaluate(() => [...document.querySelectorAll(".mk-card:not([hidden])")].map((c) => { const r = c.getBoundingClientRect(); return { key: c.dataset.key, top: r.top, bottom: r.bottom, cls: c.className }; }));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1440, height: 950 } });
  await context.addInitScript(() => {
    sessionStorage.setItem("introPlayed", "true");
    localStorage.setItem("dwb.authorName", "Tester Testowy");
    window.__dwbAutoDialog = "accept";
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); setSidebarOpen(false); });
  await page.locator("#fileInput").setInputFiles("docs/samples/review-sample.docx");
  await idle(page);

  // ── Czytanie: dymki obok kartki ──────────────────────────────────────────
  const its = await items(page);
  const cards = await visibleCards(page);
  const pageRight = await page.evaluate(() => Math.max(...[...document.querySelectorAll(".docx-preview-host section.docx")].map((s) => s.getBoundingClientRect().right)));
  check("dymki: komentarze i zmiany na marginesie", its.some((i) => i.kind === "comment") && its.some((i) => i.kind === "change"), JSON.stringify(its.map((i) => i.key)));
  check("dymki stoją na prawo od kartki", cards.length > 0 && (await page.evaluate(() => [...document.querySelectorAll(".mk-card:not([hidden])")].every((c) => c.getBoundingClientRect().left > Math.max(...[...document.querySelectorAll(".docx-preview-host section.docx")].map((s) => s.getBoundingClientRect().right))))), String(pageRight));
  const sorted = [...cards].sort((a, b) => a.top - b.top);
  const overlap = sorted.some((c, i) => i && c.top < sorted[i - 1].bottom - 0.5);
  check("dymki nie nachodzą na siebie", !overlap, JSON.stringify(sorted.map((c) => [c.key, Math.round(c.top), Math.round(c.bottom)])));
  check("same wstawki bez dymka, dopóki kursor nie stanie w akapicie (jak Word)", its.some((i) => i.kind === "change") && (await page.evaluate(() => [...document.querySelectorAll(".mk-change[hidden]")].length)) >= 1);
  const hl = await page.evaluate(() => [...CSS.highlights.keys()].filter((k) => k.startsWith("dwb-rev-")).flatMap((k) => [...CSS.highlights.get(k)].map((r) => r.toString())));
  check("wstawki podkreślone w tekście: dokładnie wstawiony tekst", hl.includes("30") && hl.some((s) => s.trim() === "niezwłocznie"), JSON.stringify(hl));
  check("pasek przy akapitach ze zmianą", (await page.evaluate(() => document.querySelectorAll(".dwb-markup .mk-bar").length)) >= 5);
  // zgłoszenie 2026-11-13: trzy pary ✓ / ✗ przy jednej zamianie — usunięcie + wstawka w tym samym
  // miejscu = jedna decyzja „Zamieniono: 14 → 30”, przyciski tylko w nagłówku
  const rep = await page.evaluate(() => { const c = [...document.querySelectorAll(".mk-change")].find((x) => x.querySelector(".mk-row-replace")); return c ? { text: c.querySelector(".mk-row-replace").textContent.replace(/\s+/g, " "), pairs: c.querySelectorAll(".mk-ok").length, head: !!c.querySelector(".mk-tools-main .mk-ok") } : null; });
  check("usunięcie + wstawka = jedna zamiana „14 → 30” (Czytanie: bez przycisków)", rep && /14 → 30/.test(rep.text) && rep.pairs === 0, JSON.stringify(rep));
  // zgłoszenie 2026-11-13: pasek przy zmianie w komórce tabeli stał między kolumnami
  // tabela z dwiema kolumnami, zmiana w DRUGIEJ (jak w zgłoszeniu) — ten sam plik, komórka dopisana z lewej
  await page.evaluate(async () => {
    const z = await JSZip.loadAsync(originalFileBytes);
    let x = await z.file("word/document.xml").async("string");
    x = x.replace('<w:gridCol w:w="4000"/>', '<w:gridCol w:w="3000"/><w:gridCol w:w="4000"/>')
      .replace("<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t xml:space=\"preserve\">Komórka ", "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Parametr</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p><w:r><w:t xml:space=\"preserve\">Komórka ");
    z.file("word/document.xml", x);
    await ingestFile(new File([await z.generateAsync({ type: "blob" })], "review-sample.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
  });
  await idle(page);
  check("tabela: zmiana w drugiej kolumnie", await page.evaluate(() => [...document.querySelectorAll(".docx-preview-host td")].length >= 2));
  const barX = await page.evaluate(() => [...new Set([...document.querySelectorAll(".dwb-markup .mk-bar")].map((b) => Math.round(b.getBoundingClientRect().left)))]);
  check("paski zmian w jednej linii na lewym marginesie (też przy tabeli)", barX.length === 1, JSON.stringify(barX));
  // zgłoszenie 2026-11-13: dymki zepchnięte daleko od wiersza dawały długie pionowe linie
  const longWires = await page.evaluate(() => [...document.querySelectorAll(".mk-wire polyline:not(.is-focus)")].filter((pl) => { const pts = pl.getAttribute("points").split(" ").map((x) => x.split(",").map(Number)); return Math.abs(pts[2][1] - pts[0][1]) > 60; }).length);
  check("bez długich linii do dymków zepchniętych daleko", longWires === 0, String(longWires));

  // ── tekst przesuwa się PO narysowaniu (przerwa strony, czcionka, obraz) ────
  // zgłoszenie 2026-11-13: dymek, pasek i linia zostawały przy starym miejscu tekstu
  const shift = await page.evaluate(async () => {
    const art = document.querySelector(".docx-preview-host section.docx article");
    const spacer = document.createElement("div");
    spacer.style.height = "180px";
    spacer.className = "test-spacer";
    art.prepend(spacer);
    await new Promise((r) => setTimeout(r, 700));
    const p = [...document.querySelectorAll(".docx-preview-host p")].find((x) => x.textContent.includes("Termin płatności"));
    const bar = [...document.querySelectorAll(".mk-bar")].map((b) => b.getBoundingClientRect().top).sort((a, b) => Math.abs(a - p.getBoundingClientRect().top) - Math.abs(b - p.getBoundingClientRect().top))[0];
    const card = document.querySelector('.mk-card[data-key^="g:"]:not([hidden])');
    const it = dwbMargin._items().find((i) => i.key === card?.dataset.key);
    const out = { pTop: p.getBoundingClientRect().top, bar, anchorOk: it ? Math.abs(it.anchorY - (p.getBoundingClientRect().top - docViewportEl.getBoundingClientRect().top + docViewportEl.scrollTop)) < 30 : false };
    spacer.remove();
    await new Promise((r) => setTimeout(r, 700));
    return out;
  });
  check("tekst przesunięty po narysowaniu: pasek idzie za akapitem", Math.abs(shift.bar - shift.pTop) < 2, JSON.stringify(shift));
  check("tekst przesunięty po narysowaniu: dymek przeliczony do nowego miejsca", shift.anchorOk, JSON.stringify(shift));

  // ── Edycja: kursor w komentarzu = aktywny dymek przy wierszu, bez karty ───
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);
  const kara = await page.evaluate(() => { const r = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.includes("Kara umowna")).getBoundingClientRect(); return { x: r.left + 20, y: r.top + r.height / 2, top: r.top }; });
  await page.mouse.click(kara.x, kara.y);
  await page.waitForTimeout(600);
  const act = await page.evaluate(() => { const c = document.querySelector(".mk-card.is-active"); return c ? { key: c.dataset.key, top: c.getBoundingClientRect().top } : null; });
  check("kursor w komentarzu: jego dymek aktywny", act?.key === "c:10", JSON.stringify(act));
  check("aktywny dymek przy swoim wierszu (±12 px)", act && Math.abs(act.top - kara.top) <= 12, `${act?.top} vs ${kara.top}`);
  check("bez karty komentarza nad tekstem, gdy jest dymek", !(await page.$(".comment-card")));
  check("linia od tekstu do aktywnego dymka", await page.evaluate(() => !!document.querySelector(".mk-wire polyline")));
  const after = await visibleCards(page);
  const s2 = [...after].sort((a, b) => a.top - b.top);
  check("po aktywacji dalej bez nachodzenia", !s2.some((c, i) => i && c.top < s2[i - 1].bottom - 0.5), JSON.stringify(s2.map((c) => [c.key, Math.round(c.top), Math.round(c.bottom)])));

  // ── ✓ z dymka zmienia plik; Cofnij wraca ─────────────────────────────────
  const m0 = marks(await fileXml(page, "word/document.xml"));
  const key = await page.evaluate(() => [...document.querySelectorAll(".mk-change")].find((c) => c.querySelector(".mk-row-format"))?.dataset.key);
  await page.evaluate((k) => document.querySelector(`[data-key="${k}"] .mk-tools-main .mk-ok`).click(), key); // jedna decyzja — ✓ w nagłówku
  await idle(page);
  const m1 = marks(await fileXml(page, "word/document.xml"));
  check("✓ w dymku: zmiana przyjęta w pliku", m1 === m0 - 1, `${m0} → ${m1}`);
  check("po ✓: ten dymek zniknął", !(await items(page)).some((i) => i.key === key));
  await page.evaluate(() => dwbUndo.undo());
  await idle(page);
  check("Cofnij: zmiana wróciła", marks(await fileXml(page, "word/document.xml")) === m0);
  const repEdit = await page.evaluate(() => { const c = [...document.querySelectorAll(".mk-change")].find((x) => x.querySelector(".mk-row-replace")); return c ? { pairs: c.querySelectorAll(".mk-ok").length, head: !!c.querySelector(".mk-tools-main .mk-ok"), rowBtns: c.querySelectorAll(".mk-row-act").length } : null; });
  check("Edycja: przy zamianie jedna para ✓ / ✗ — w nagłówku, bez przycisków w wierszu", repEdit && repEdit.pairs === 1 && repEdit.head && !repEdit.rowBtns, JSON.stringify(repEdit));
  for (const [act, want] of [["mk-ok", "30"], ["mk-no", "14"]]) {
    await page.evaluate((act) => { const c = [...document.querySelectorAll(".mk-change")].find((x) => x.querySelector(".mk-row-replace")); c.querySelector(`.mk-tools-main .${act}`).click(); }, act);
    await idle(page);
    const x = await fileXml(page, "word/document.xml");
    const para = (x.match(/<w:p\b(?:(?!<w:p\b).)*?Termin(?:(?!<\/w:p>).)*<\/w:p>/s) || [""])[0];
    check(`zamiana ${act === "mk-ok" ? "✓" : "✗"}: obie poprawki naraz, zostaje „${want}”`, !/<w:(ins|del)\b/.test(para) && para.replace(/<[^>]+>/g, "").includes(`wynosi ${want} dni`), para.replace(/<[^>]+>/g, ""));
    await page.evaluate(() => dwbUndo.undo());
    await idle(page);
  }

  // ── Odpowiedz z dymka ────────────────────────────────────────────────────
  await page.click('[data-key="c:10"]');
  await page.waitForTimeout(400);
  await page.click('[data-key="c:10"] .mk-reply-btn');
  await page.waitForSelector(".compose-pop-comment .cf-rich");
  await page.locator(".compose-pop-comment .cf-rich").click();
  await page.keyboard.type("Odpowiedź z dymka");
  await page.click(".compose-pop-comment .lf-ok");
  await idle(page);
  check("Odpowiedz w dymku: odpowiedź w pliku", (await fileXml(page, "word/comments.xml")).includes("Odpowiedź z dymka"));
  check("odpowiedź widać w dymku", await page.evaluate(() => !!document.querySelector('[data-key="c:10"]')?.textContent.includes("Odpowiedź z dymka")));

  // ── Rozwiąż z dymka ─────────────────────────────────────────────────────
  await page.evaluate(() => document.querySelector('[data-key="c:10"] .mk-tools .mk-btn').click());
  await idle(page);
  const ex = await fileXml(page, "word/commentsExtended.xml");
  check("Rozwiąż w dymku: komentarz rozwiązany w pliku", /w15:done="1"/.test(ex));
  check("rozwiązany dymek zwinięty", await page.evaluate(() => document.querySelector('[data-key="c:10"]')?.classList.contains("is-done")));

  // ── kolor tekstu komentarza z pliku nieczytelny na ciemnym dymku (szary 595959 z Worda) ──
  await page.evaluate(() => { document.documentElement.dataset.theme = "dark"; dwbMargin.schedule(); });
  await page.waitForTimeout(500);
  const greyOk = await page.evaluate(() => {
    const body = document.querySelector('[data-key="c:20"] .mk-body');
    if (!body) return "brak dymka";
    return dwbMargin._readable({ text: "x", color: "595959" }).color === undefined && dwbMargin._readable({ text: "x", color: "FF0000" }).color === "FF0000" && dwbMargin._readable({ text: "x", hl: "yellow" }).color === "111111";
  });
  check("ciemny motyw: szary kolor z pliku pominięty, czytelne kolory zostają", greyOk === true, String(greyOk));
  await page.evaluate(() => { document.documentElement.dataset.theme = "light"; dwbMargin.schedule(); });
  await page.waitForTimeout(400);

  // ── tryb „Bez” = bez dymków i podkreśleń; karta komentarza wraca ───────────
  await page.evaluate(() => dwbMargin.setMode("none"));
  await page.waitForTimeout(500);
  check("„Bez”: bez dymków i podkreśleń", await page.evaluate(() => !document.querySelector(".mk-card") && ![...CSS.highlights.keys()].some((k) => k.startsWith("dwb-rev-"))));
  const zgl = await page.evaluate(() => { const r = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.includes("Ostatni akapit")).getBoundingClientRect(); return { x: r.left + 20, y: r.top + r.height / 2 }; });
  await page.mouse.click(zgl.x, zgl.y);
  await page.waitForTimeout(600);
  check("„Bez”: komentarz pokazuje karta nad tekstem", !!(await page.$(".comment-card")));
  await page.evaluate(() => dwbMargin.setMode("all"));
  await page.waitForTimeout(500);

  // ── wąski ekran: bez dymków ──────────────────────────────────────────────
  await page.setViewportSize({ width: 900, height: 900 });
  await page.waitForTimeout(800);
  check("wąski ekran: bez kolumny dymków (brak miejsca obok kartki)", await page.evaluate(() => !document.querySelector(".mk-card:not([hidden])")));
  check("wąski ekran: ikonki komentarzy przy krawędzi strony", await page.evaluate(() => document.querySelectorAll(".mk-cmark").length >= 1));
  await page.locator(".mk-bar-hit").first().click();
  await page.waitForTimeout(500);
  const pop = await page.evaluate(() => { const c = document.querySelector(".mk-card.is-compact:not([hidden])"); if (!c) return null; const r = c.getBoundingClientRect(); return { key: c.dataset.key, ask: !!c.querySelector(".mk-ask .mk-ok"), left: r.left, right: r.right, vw: docViewportEl.getBoundingClientRect().right }; });
  check("wąski ekran: stuknięcie w pasek = zmiana pod tekstem z ✓ / ✗", pop && pop.ask, JSON.stringify(pop));
  check("wąski ekran: dymek zmiany mieści się w widoku", pop && pop.left >= 0 && pop.right <= pop.vw + 1, JSON.stringify(pop));
  await page.mouse.click(5, 300);
  await page.waitForTimeout(300);
  check("wąski ekran: stuknięcie obok zamyka dymek", await page.evaluate(() => !document.querySelector(".mk-card.is-compact:not([hidden])")));
  await page.locator(".mk-cmark").first().click();
  await page.waitForTimeout(500);
  check("wąski ekran: ikonka komentarza otwiera kartę komentarza", !!(await page.$(".comment-card")));
  await page.keyboard.press("Escape");
  await page.setViewportSize({ width: 1440, height: 950 });
  await page.waitForTimeout(800);

  // ── Proste (jak „Proste znaczniki” w Wordzie): bez kolumny zmian, zmiana pod tekstem po kliknięciu ──
  await page.evaluate(() => dwbMargin.setMode("simple"));
  await page.waitForTimeout(600);
  check("Proste: bez podkreśleń wstawek i bez dymków zmian w kolumnie", await page.evaluate(() => ![...CSS.highlights.keys()].some((k) => k.startsWith("dwb-rev-")) && !document.querySelector(".mk-change:not(.is-compact):not([hidden])")));
  check("Proste: komentarze dalej obok strony", await page.evaluate(() => !!document.querySelector(".mk-comment:not([hidden])")));
  const tp = await page.evaluate(() => { const r = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.startsWith("Termin")).getBoundingClientRect(); return { x: r.left + 40, y: r.top + r.height / 2, bottom: r.bottom }; });
  await page.mouse.click(tp.x, tp.y);
  await page.waitForTimeout(600);
  const sp = await page.evaluate(() => { const c = document.querySelector(".mk-card.is-compact:not([hidden])"); return c ? { top: c.getBoundingClientRect().top, txt: c.textContent } : null; });
  check("Proste: klik w zmieniony akapit = dymek pod wierszem z pytaniem", sp && sp.top >= tp.bottom - 30 && /Zaakceptować|Accept/.test(sp.txt), JSON.stringify(sp));
  await page.evaluate(() => dwbMargin.setMode("all"));
  await page.waitForTimeout(500);

  // ── Śledź zmiany na żywo ─────────────────────────────────────────────────
  await page.evaluate(() => loadSampleDocument("sample"));
  await page.waitForSelector(".docx-preview-host p");
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);
  await page.evaluate(() => dwbTrack.start());
  await idle(page);
  const p2 = await page.evaluate(() => { const p = [...document.querySelectorAll(".docx-preview-host p")].find((x) => x.textContent.includes("przykładowy")); const r = p.getBoundingClientRect(); return { x: r.left + 2, y: r.top + 8 }; });
  await page.mouse.click(p2.x, p2.y);
  await page.keyboard.press("Home");
  await page.keyboard.type("Nowe ");
  await page.keyboard.press("End");
  for (let i = 0; i < 4; i++) await page.keyboard.press("Backspace");
  await page.waitForTimeout(1300);
  const live = await page.evaluate(() => ({
    hl: [...CSS.highlights.keys()].filter((k) => k.startsWith("dwb-rev-")).flatMap((k) => [...CSS.highlights.get(k)].map((r) => r.toString())),
    card: document.querySelector(".mk-live")?.textContent || "",
  }));
  check("śledzenie: wpisany tekst podkreślony od razu", live.hl.some((s) => s === "Nowe "), JSON.stringify(live.hl));
  check("śledzenie: usunięte znaki w dymku (tylko one)", /Usunięto:\s*PWA\./.test(live.card) || /Deleted:\s*PWA\./.test(live.card), live.card);
  await page.evaluate(() => dwbTrack.stop());
  await idle(page);

  // ── obraz: róg „staje” na 50% szerokości tekstu ──────────────────────────
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p");
  await idle(page);
  await page.keyboard.type("Tekst.");
  const file = await page.evaluateHandle(async () => { const c = document.createElement("canvas"); c.width = 900; c.height = 450; c.getContext("2d").fillRect(0, 0, 900, 450); const b = await new Promise((r) => c.toBlob(r, "image/png")); return new File([b], "a.png", { type: "image/png" }); });
  await page.evaluate((f) => composeUi.insertImageFile(f), file);
  await idle(page);
  await page.click(".docx-preview-host img");
  await page.waitForSelector('.img-h[data-h="se"]');
  await page.waitForTimeout(300);
  const g = await page.evaluate(() => {
    const img = document.querySelector(".docx-preview-host img");
    const h = document.querySelector('.img-h[data-h="se"]').getBoundingClientRect();
    return { x: h.left + h.width / 2, y: h.top + h.height / 2, w: img.getBoundingClientRect().width, text: img.closest("p").getBoundingClientRect().width, k: getComputedStyle(img.closest("p")).textAlign === "center" ? 2 : 1 };
  });
  await page.mouse.move(g.x, g.y);
  await page.mouse.down();
  const dx = (g.text * 0.5 + 4 - g.w) / g.k; // 4 px obok połowy — ma przyciągnąć
  for (let i = 1; i <= 10; i++) await page.mouse.move(g.x + dx * i / 10, g.y, { steps: 2 });
  const during = await page.evaluate(() => ({ badge: document.querySelector(".img-snap-badge")?.textContent, snap: document.querySelector(".img-snap-badge")?.classList.contains("is-snap"), guide: !document.querySelector(".img-snap-guide")?.hidden }));
  await page.mouse.up();
  await idle(page);
  check("róg obrazu: przyciąga do 50% z etykietą i linią", during.snap && during.badge === "50%" && during.guide, JSON.stringify(during));
  const ratio = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const x = await z.file("word/document.xml").async("string");
    const d = new DOMParser().parseFromString(x, "application/xml");
    const p = collectParagraphElements(d.documentElement, "all").find((q) => q.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "extent").length);
    const cx = +p.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "extent")[0].getAttribute("cx");
    return cx / (composeTextWidthTwips(d, p) * 635);
  });
  check("w pliku: obraz = połowa szerokości tekstu", Math.abs(ratio - 0.5) < 0.002, String(ratio));
  check("uchwyty idą za obrazem po zapisie", await page.evaluate(() => { const i = document.querySelector(".docx-preview-host img").getBoundingClientRect(); const h = document.querySelector(".img-handles").getBoundingClientRect(); return Math.abs(i.width - h.width) < 1.5 && Math.abs(i.left - h.left) < 1.5; }));

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`);
    if (!r.ok) failed++;
  }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  if (failed) process.exit(1);
}

run().catch((e) => { console.error(e); process.exit(1); });
