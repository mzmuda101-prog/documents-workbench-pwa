// integrity-playwright.js — zwykłe operacje nie mogą niszczyć niczego obok (audyt 2026-10-04).
//
// Przewodnik (zakładki _GuideN przy rozdziałach = cele spisu treści i linków, komentarz, przypis,
// kontrolki formularza). Po każdej operacji ZAPISANY plik: zakładki obejmują swój tekst (dawniej
// przepisanie akapitu zostawiało je przed tekstem — pusty punkt; odsyłacz REF Worda pokazałby
// pusty tekst), komentarze mają odwołanie i zakres, nic nie jest osierocone. „Wielkie litery /
// Przytnij / Prefiks” nie gubią pogrubienia, kursywy ani komentarza (dawniej jeden goły fragment),
// „zaznacz wszystko + Delete” nie zostawia komentarza bez kotwicy. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

const audit = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  const x = await z.file("word/document.xml").async("string");
  const d = new DOMParser().parseFromString(x, "application/xml");
  const ids = (tag) => [...d.getElementsByTagNameNS(W_NS, tag)].map((e) => e.getAttributeNS(W_NS, "id"));
  const bs = ids("bookmarkStart"); const be = ids("bookmarkEnd");
  const cs = ids("commentRangeStart"); const ce = ids("commentRangeEnd"); const cr = ids("commentReference");
  const cx = z.file("word/comments.xml") ? await z.file("word/comments.xml").async("string") : "";
  const cIds = [...cx.matchAll(/<w:comment [^>]*w:id="(\d+)"/g)].map((m) => m[1]);
  const bm = {};
  [...d.getElementsByTagNameNS(W_NS, "bookmarkStart")].forEach((s) => {
    const id = s.getAttributeNS(W_NS, "id"); let t = ""; let n = s.nextSibling;
    while (n && !(n.localName === "bookmarkEnd" && n.getAttributeNS(W_NS, "id") === id)) { if (n.nodeType === 1) t += [...n.getElementsByTagNameNS(W_NS, "t")].map((q) => q.textContent).join(""); n = n.nextSibling; }
    bm[s.getAttributeNS(W_NS, "name")] = n ? t : null;
  });
  return {
    bm, bmCount: bs.length,
    orphans: bs.filter((i) => !be.includes(i)).length + be.filter((i) => !bs.includes(i)).length
      + cIds.filter((i) => !cr.includes(i)).length + cr.filter((i) => !cIds.includes(i)).length + cs.filter((i) => !ce.includes(i)).length + ce.filter((i) => !cs.includes(i)).length,
    comments: cIds.length, refs: cr.length,
    b: (x.match(/<w:b\/>/g) || []).length, i: (x.match(/<w:i\/>/g) || []).length,
    fn: (x.match(/footnoteReference/g) || []).length, sdt: (x.match(/<w:sdt>/g) || []).length,
  };
});

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); try { localStorage.setItem("dwb.authorName", "Jan"); } catch (_) {} });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(`${APP_URL}?sample=przewodnik`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.evaluate(() => appFrame.setReadOnly(false));
  const idle = () => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending && !inlineStructuralPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(450));
  await idle();
  const base = await audit(page);
  check("przewodnik: 14 zakładek rozdziałów, komentarz, przypis, kontrolki (warunek testu)", base.bmCount === 14 && base.comments === 1 && base.fn === 1 && base.sdt === 4 && !base.orphans, JSON.stringify(base));
  const head = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).findIndex((p) => !p.dataset.lock && p.querySelector('span[data-cm-kind="bm-start"]')));
  const caret = (off) => page.evaluate(([i, o]) => placeCaret(collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i], o), [head, off]);

  await caret(0);
  await page.keyboard.type("Nowy ");
  await page.waitForTimeout(250);
  let a = await audit(page);
  check("pisanie w nagłówku z zakładką: zakładka dalej obejmuje tekst rozdziału", a.bm._Guide1 === "Nowy 1. Czytanie i edycja" && !a.orphans, JSON.stringify(a.bm._Guide1));
  await caret(4);
  await page.keyboard.press("Enter"); await idle();
  await page.keyboard.press("Backspace"); await idle();
  a = await audit(page);
  check("Enter i Backspace w nagłówku: zakładka wraca w całości, nic osieroconego", a.bm._Guide1 === "Nowy 1. Czytanie i edycja" && !a.orphans, JSON.stringify(a.bm._Guide1));
  await page.evaluate((i) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; const r = formDomRange(p, 5, 7); p.closest(".docx-edit-root").focus({ preventScroll: true }); getSelection().removeAllRanges(); getSelection().addRange(r); }, head);
  await page.evaluate(() => composeUi.applyColor("color", "FF0000")); await idle();
  await page.evaluate(() => composeUi.applyStyle("h2")); await idle();
  a = await audit(page);
  check("kolor fragmentu i zmiana stylu: zakładka i tekst bez zmian", a.bm._Guide1 === "Nowy 1. Czytanie i edycja" && !a.orphans, JSON.stringify(a.bm._Guide1));

  const before = await audit(page);
  await page.evaluate(() => applyDocumentEdit({ op: "case", mode: "upper", scope: "all" })); await idle();
  a = await audit(page);
  check("WIELKIE LITERY w całym dokumencie: pogrubienia, kursywa, komentarz, przypis i zakładki zostają", a.b === before.b && a.i === before.i && a.comments === 1 && a.refs === before.refs && a.fn === 1 && a.bm._Guide2 === "2. SZUKANIE ORAZ ZNAJDŹ I ZAMIEŃ" && !a.orphans, JSON.stringify({ b: [before.b, a.b], i: [before.i, a.i], c: a.comments, bm: a.bm._Guide2, o: a.orphans }));
  await page.evaluate(() => applyDocumentEdit({ op: "affix", prefix: "» ", suffix: "", scope: "all" })); await idle();
  a = await audit(page);
  check("prefiks we wszystkich akapitach: formatowanie i komentarz zostają", a.b === before.b && a.comments === 1 && !a.orphans, JSON.stringify({ b: a.b, c: a.comments, o: a.orphans }));
  await page.keyboard.press(`${MOD}+KeyZ`); await idle();
  await page.keyboard.press(`${MOD}+KeyZ`); await idle();

  await page.evaluate(() => focusParagraphAtOffset(0, 0));
  await page.keyboard.press(`${MOD}+KeyA`);
  await page.keyboard.press("Delete");
  await idle();
  a = await audit(page);
  check("zaznacz wszystko + Delete: komentarz i przypis znikają razem z tekstem (bez sierot w comments.xml)", a.comments === 0 && a.fn === 0 && !a.orphans, JSON.stringify(a));
  await page.keyboard.press(`${MOD}+KeyZ`); await idle();
  a = await audit(page);
  check("Cofnij przywraca komentarz, przypis i zakładki", a.comments === 1 && a.fn === 1 && a.bmCount === 14 && !a.orphans, JSON.stringify({ c: a.comments, fn: a.fn, bm: a.bmCount, o: a.orphans }));

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : `\n   ${String(r.detail).slice(0, 700)}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
