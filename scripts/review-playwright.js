// review-playwright.js — panel Recenzja (app/review-panel.js + app/docx-revisions.js).
//
// Fixture: docs/samples/review-sample.docx (node scripts/gen-review-docx.mjs) — dwóch autorów,
// wstawienia/usunięcia/przeniesienia/formatowanie/znaki akapitu/tabela/nagłówek, komentarze
// z odpowiedzią, przypisy. Sprawdza TREŚĆ PLIKU po operacji (nie tylko UI).

const { createTestPage, bootApp, loadDocxFile, assertNoErrors } = require("./docx-test-helpers");

const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

const fileState = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(originalFileBytes);
  const x = await z.file("word/document.xml").async("string");
  const d = new DOMParser().parseFromString(x, "application/xml");
  const hdr = await z.file("word/header1.xml").async("string");
  const comments = await z.file("word/comments.xml").async("string");
  return {
    body: collectParagraphElements(d.documentElement, "body").map(revParaText),
    marks: (x.match(/<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b/g) || []).length,
    hdr: revText(new DOMParser().parseFromString(hdr, "application/xml").documentElement),
    comments: (comments.match(/<w:comment\b/g) || []).length,
    refs: (x.match(/commentReference/g) || []).length,
  };
});

async function openReview(page) {
  await page.evaluate(async () => {
    setSidebarOpen(true);
    document.getElementById("panel-review").open = true;
    await ensureLazyFeature("review");
    await runReviewScan();
  });
  await page.waitForFunction(() => document.getElementById("toolsPanel").getBoundingClientRect().left >= 0, null, { timeout: 5000 });
  await page.waitForTimeout(300);
}

async function run() {
  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await page.evaluate(() => { window.confirm = () => true; });
  await loadDocxFile(page, "docs/samples/review-sample.docx");
  await page.waitForTimeout(400);

  // ── po otwarciu: komunikat i plakietka ─────────────────────────────────────
  const toasts = await page.$$eval(".toast", (els) => els.map((e) => e.textContent));
  check("komunikat po otwarciu: 12 zmian i 3 komentarze", toasts.some((s) => /12/.test(s) && /\(3\)/.test(s) && /Recenzja/.test(s)), toasts.join(" | "));
  check("plakietka przy panelu Recenzja = 15", (await page.textContent("#panel-review > summary .panel-count").catch(() => "")) === "15");

  // ── lista ──────────────────────────────────────────────────────────────────
  await openReview(page);
  const counts = await page.evaluate(() => ({
    changes: document.querySelectorAll("#rvChanges .rv-row").length,
    comments: document.querySelectorAll("#rvComments .rv-row").length,
    replies: document.querySelectorAll("#rvComments .rv-reply").length,
    notes: document.querySelectorAll("#rvNotes .rv-row").length,
    summary: document.getElementById("rvSummary").textContent,
    authors: [...document.querySelectorAll("#rvAuthor option")].map((o) => o.textContent),
    header: [...document.querySelectorAll("#rvChanges .rv-meta")].some((m) => /w nagłówku/.test(m.textContent)),
    del: document.querySelector("#rvChanges .rv-del .rv-text")?.textContent,
    anchor: document.querySelector("#rvComments .rv-anchor")?.textContent,
  }));
  check("12 zmian, 2 komentarze (+1 odpowiedź), 2 przypisy", counts.changes === 12 && counts.comments === 2 && counts.replies === 1 && counts.notes === 2, JSON.stringify(counts));
  check("filtr autorów: wszyscy + Anna + Jan", counts.authors.length === 3 && counts.authors.some((a) => /Anna Nowak \(6\)/.test(a)), counts.authors.join(", "));
  check("zmiana w nagłówku jest oznaczona", counts.header);
  check("usunięty tekst widać na liście („14”)", counts.del === "14", counts.del);
  check("komentarz pokazuje zakomentowany tekst", counts.anchor === "„Kara umowna”", counts.anchor);

  await page.click("#rvComments .rv-row.is-jump");
  await page.waitForTimeout(150);
  check("klik w komentarz podświetla akapit w podglądzie", await page.evaluate(() => /Kara umowna/.test(document.querySelector(".search-hit-active")?.textContent || "")));

  // ── Edycja: akapity ze złożoną treścią są zablokowane, zwykłe edytowalne ─────
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change", { bubbles: true })); });
  await page.waitForTimeout(400);
  const locks = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => `${p.dataset.lock || "-"}:${p.contentEditable}`));
  // akapit z odnośnikiem do przypisu (9) — od 2026-10-04 edytowalny, odnośnik to „wyspa” (doc-notes.js)
  check("blokady: śledzone zmiany (akapity 2, 4, 5 — też sama zmiana formatu); akapit z przypisem końcowym (9) edytowalny",
    locks[1] === "lockTracked:false" && locks[3] === "lockTracked:false" && locks[4] === "lockTracked:false" && locks[8] === "-:true", locks.join(" "));
  check("akapit ze zmianą tylko formatu AKAPITU (wyrównanie) edytowalny — zapis tekstu jej nie rusza", locks[5] === "-:true", locks[5]);
  check("akapit z samym komentarzem da się edytować", locks[2] === "-:true", locks[2]);
  check("mapowanie: w podglądzie tyle akapitów treści, ile w pliku (bez nagłówka i przypisów)",
    await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).length === baselineParagraphRuns.length));

  // wpisz w akapit 3 („Kara umowna…”) — potem operacja z panelu scala to z plikiem
  await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[2];
    p.focus();
    const r = document.createRange(); r.selectNodeContents(p); r.collapse(false);
    const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
  });
  await page.keyboard.type(" DOPISANE");
  await page.waitForTimeout(300);

  // ── pojedyncza zmiana: ✗ na „30” (wstawienie) → zostaje „14”? nie: ✗ odrzuca tylko wstawienie
  await page.click("#rvChanges .rv-ins .rv-reject");
  await page.waitForTimeout(500);
  let s = await fileState(page);
  // 12 zmian = 11 w treści + 1 w nagłówku
  check("✗ na jednym wstawieniu: „30” znika, usunięcie „14” dalej czeka", s.body[1] === "Termin płatności wynosi  dni." && s.marks === 10, `${s.body[1]} / ${s.marks}`);
  check("wpisany tekst trafił do SWOJEGO akapitu, nagłówek nie wszedł w treść",
    /Kara umowna wynosi 5% wartości zlecenia\. DOPISANE$/.test(s.body[2]) && !s.body.some((b) => /Poufne/.test(b)) && s.body.length === 13, JSON.stringify(s.body.slice(0, 4)));
  check("lista odświeżona po zmianie (11)", (await page.$$eval("#rvChanges .rv-row", (e) => e.length)) === 11);

  // Cofnij przywraca
  await page.click("#undoBtn");
  await page.waitForTimeout(700);
  s = await fileState(page);
  check("Cofnij przywraca zmianę (11 znaczników w treści)", s.marks === 11 && s.body[1] === "Termin płatności wynosi 30 dni.", `${s.body[1]} / ${s.marks}`);
  check("po Cofnij lista w panelu znów ma 12", (await page.$$eval("#rvChanges .rv-row", (e) => e.length)) === 12);

  // ── autor: Akceptuj tylko zmiany Jana ──────────────────────────────────────
  await page.selectOption("#rvAuthor", "Jan Kowalski");
  check("przycisk mówi, czyje zmiany", /Jan Kowalski/.test(await page.textContent("#rvAcceptAllBtn")));
  await page.click("#rvAcceptAllBtn");
  await page.waitForTimeout(600);
  s = await fileState(page);
  check("akceptacja Jana: jego wstawki są w tekście, zmiany Anny czekają",
    s.body.includes("Zleceniodawca niezwłocznie zgłasza usterki.") && s.body.includes("Cały nowy akapit od Jana.") && /projekt/.test(s.hdr) && s.marks === 6, `${s.marks} ${s.hdr}`);

  // ── komentarze: usuń wszystkie ─────────────────────────────────────────────
  await page.click("#rvRemoveCommentsBtn");
  await page.waitForTimeout(600);
  s = await fileState(page);
  check("usuń komentarze: brak komentarzy i odwołań w pliku", s.comments === 0 && s.refs === 0, JSON.stringify({ c: s.comments, r: s.refs }));

  // ── reszta (Anna): Odrzuć wszystkie ────────────────────────────────────────
  // pole autora bywa ukryte (jeden autor) — selectOption czekał wtedy 30 s na widoczność i błąd
  // był połykany; ustawiamy wartość wprost (2026-10-02: test 36 s → kilka sekund)
  await page.evaluate(() => { const s = document.getElementById("rvAuthor"); if (s) { s.value = ""; s.dispatchEvent(new Event("change", { bubbles: true })); } });
  await page.click("#rvRejectAllBtn");
  await page.waitForTimeout(600);
  s = await fileState(page);
  check("odrzuć resztę: „14 dni”, akapity osobno, przeniesienie wraca, brak znaczników",
    s.marks === 0 && s.body[1] === "Termin płatności wynosi 14 dni." && s.body.includes("Początek zdania") && s.body.includes("Przeniesiony fragment. Stały tekst."), JSON.stringify(s.body));
  check("po wszystkim: plakietka zniknęła", !(await page.$("#panel-review > summary .panel-count")));
  check("pusta lista mówi „Brak śledzonych zmian”", /Brak śledzonych zmian/.test(await page.textContent("#rvChanges")));

  // plik po zmianach dalej się otwiera (podgląd bez błędów)
  const unlocked = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => p.dataset.lock || "-"));
  check("po rozstrzygnięciu zmian akapit 2 znów edytowalny (i nic nie jest zablokowane)", unlocked[1] === "-" && unlocked.every((l) => l === "-"), unlocked.join(" "));
  check("podgląd dokumentu działa po zmianach", await page.evaluate(() => document.querySelectorAll(".docx-preview-host p").length > 5));

  // ── dokument z podziałami strony = kilka <section> w podglądzie ────────────
  // Dawniej brana była tylko PIERWSZA sekcja (6 z 34 akapitów) → edycja dalszych stron trafiała źle.
  const multi = await page.evaluate(async () => {
    const res = await fetch("docs/samples/headings-sample.docx");
    const z = await JSZip.loadAsync(await res.arrayBuffer());
    let x = await z.file("word/document.xml").async("string");
    let n = 0;
    x = x.replace(/<\/w:p>/g, (m) => (++n === 5 || n === 12 ? `</w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p>` : m));
    z.file("word/document.xml", x);
    const blob = await z.generateAsync({ type: "blob" });
    await ingestFile(new File([blob], "strony.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
    await new Promise((r) => setTimeout(r, 600));
    const host = document.querySelector(".docx-preview-host");
    return { sections: host.querySelectorAll("section.docx").length, preview: collectPreviewParagraphElements(host).length, base: baselineParagraphRuns.length };
  });
  check("kilka stron: akapity ze WSZYSTKICH sekcji podglądu = akapity pliku", multi.sections >= 3 && multi.preview === multi.base, JSON.stringify(multi));

  assertNoErrors(errors, "review");
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ recenzja: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
