// comments2-playwright.js — komentarze, runda 2 (2026-10-04):
//   - dodanie komentarza / odpowiedzi / rozwiązanie / usunięcie / edycja NIE przewijają dokumentu
//     (dawniej skok o sumę odstępów między stronami — page-breaks.js po przerysowaniu),
//   - formatowanie w komentarzu jak w Wordzie (B / I / U / przekreślenie, Ctrl/⌘+B…) → w:rPr w comments.xml,
//   - edycja komentarza i odpowiedzi w trybie Edycja (paraId zostaje — odpowiedzi i „rozwiązany” dalej
//     wiszą przy komentarzu), komentarz z czymś, czego model nie zapisze (link), jest zablokowany,
//   - długi komentarz: wygaszenie + „Więcej ↓”, znika po przewinięciu do końca,
//   - karta komentarza rośnie razem z przybliżeniem dokumentu (150% → tekst ×1,5).
// Jak użytkownik: klawiatura, kliknięcia w kartę. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(700));
const IDX = 40;
const select = (page, a, z) => page.evaluate(([i, a, z]) => {
  const el = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
  const r = formDomRange(el, a, z); el.closest(".docx-edit-root").focus({ preventScroll: true }); getSelection().removeAllRanges(); getSelection().addRange(r);
}, [IDX, a, z]);
const paraY = (page) => page.evaluate((i) => Math.round(collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].getBoundingClientRect().top), IDX);
const zipOf = async (page) => {
  const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
  return JSZip.loadAsync(Buffer.from(b64, "base64"));
};
const commentsXml = async (page) => (await zipOf(page)).file("word/comments.xml").async("string");

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); try { localStorage.setItem("dwb.authorName", "Jan Test"); } catch (_) {} });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(`${APP_URL}?sample=przewodnik`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);
  // akapit w połowie dokumentu, kilka kartek niżej (nad nim są odstępy między stronami)
  await page.evaluate((i) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; docViewportEl.scrollTop += p.getBoundingClientRect().top - docViewportEl.getBoundingClientRect().top - 300; }, IDX);
  await page.waitForTimeout(400);
  check("przewodnik: nad akapitem są odstępy między stronami (warunek testu)", await page.evaluate(() => document.querySelectorAll(".dwb-page-gap").length > 0));
  const y0 = await paraY(page);

  // ── nowy komentarz z pogrubieniem (Ctrl/⌘+B) i drugim akapitem ───────────────
  await select(page, 0, 4);
  await page.keyboard.press(`${MOD}+Alt+KeyM`);
  check("okienko: edytor treści z fokusem, bez osobnych przycisków (B / I / U z paska)", await page.evaluate(() => !document.querySelector(".compose-pop-comment .cf-fmt") && document.activeElement?.classList.contains("cf-rich") && /B I U/.test(document.querySelector(".cf-fmt-hint").textContent)));
  await page.keyboard.type("Zwykły ");
  await page.keyboard.press(`${MOD}+KeyB`);
  await page.keyboard.type("gruby");
  await page.keyboard.press(`${MOD}+KeyB`);
  await page.keyboard.type(" tekst.");
  await page.keyboard.press("Enter");
  await page.click("#fmtItalic"); // przycisk I na pasku Edycji — działa na komentarz
  await page.keyboard.type("Drugi akapit");
  check("klik „I” na pasku: okienko zostaje, fokus w komentarzu", await page.evaluate(() => !!document.querySelector(".compose-pop-comment") && document.activeElement?.classList.contains("cf-rich")));
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
  check("dodanie komentarza: akapit na tym samym miejscu ekranu", Math.abs((await paraY(page)) - y0) <= 1, `${y0} → ${await paraY(page)}`);
  let cx = await commentsXml(page);
  const last = cx.slice(cx.lastIndexOf("<w:comment "));
  check("plik: pogrubiony fragment <w:b/> w komentarzu, reszta bez", /<w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">gruby<\/w:t><\/w:r>/.test(last) && /<w:t xml:space="preserve">Zwykły <\/w:t>/.test(last), last.slice(0, 600));
  check("plik: Enter = drugi akapit komentarza, kursywa <w:i/>", (last.match(/<w:p\b/g) || []).length === 2 && /<w:rPr><w:i\/><\/w:rPr><w:t xml:space="preserve">Drugi akapit/.test(last), last);

  // ── karta: formatowanie widoczne, „Edytuj” ──────────────────────────────────
  await select(page, 1, 1);
  await page.waitForSelector(".comment-card", { timeout: 4000 });
  check("karta: pogrubienie i kursywa widoczne", await page.evaluate(() => document.querySelector(".comment-card .cc-text b")?.textContent === "gruby" && document.querySelector(".comment-card .cc-text i")?.textContent === "Drugi akapit"));
  const yEdit = await paraY(page);
  await page.click(".comment-card .cc-edit");
  check("Edytuj: okienko z obecną treścią (z formatowaniem)", await page.evaluate(() => document.querySelector(".compose-pop-comment .cf-rich b")?.textContent === "gruby" && !!document.querySelector(".compose-pop-comment .cf-author-field[hidden]")));
  await page.keyboard.type(" Poprawione");
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
  check("edycja: akapit na tym samym miejscu ekranu", Math.abs((await paraY(page)) - yEdit) <= 1);
  cx = await commentsXml(page);
  const edited = cx.slice(cx.lastIndexOf("<w:comment "));
  check("plik: nowa treść, pogrubienie zostało, autor bez zmian, znacznik annotationRef", /Poprawione/.test(edited) && /<w:b\/><\/w:rPr><w:t xml:space="preserve">gruby/.test(edited) && /w:author="Jan Test"/.test(edited) && /<w:annotationRef\/>/.test(edited), edited);

  // ── odpowiedź, edycja odpowiedzi, rozwiąż — powiązania po paraId zostają ────
  await select(page, 1, 1);
  await page.waitForSelector(".comment-card");
  let y = await paraY(page);
  await page.click(".comment-card .btn:has-text('Odpowiedz')");
  await page.keyboard.type("Odpowiedź");
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
  check("odpowiedź: akapit na tym samym miejscu ekranu", Math.abs((await paraY(page)) - y) <= 1);
  await select(page, 1, 1);
  await page.waitForSelector(".comment-card .cc-entry + .cc-entry");
  await page.click(".comment-card .cc-entry + .cc-entry .cc-edit");
  // zaznacz całą treść (natywne „zaznacz wszystko” na Macu to ⌘A, w Chromium testów Ctrl+A to początek wiersza)
  await page.evaluate(() => { const r = document.createRange(); r.selectNodeContents(document.querySelector(".cf-rich")); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await page.keyboard.type("Odpowiedź po edycji");
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
  await select(page, 1, 1);
  await page.waitForSelector(".comment-card .cc-entry + .cc-entry");
  const replyTxt = await page.evaluate(() => document.querySelector(".comment-card .cc-entry + .cc-entry .cc-text")?.textContent);
  check("karta: odpowiedź po edycji dalej pod komentarzem", replyTxt === "Odpowiedź po edycji", replyTxt);
  y = await paraY(page);
  await page.click(".comment-card .btn:has-text('Rozwiąż')");
  await idle(page);
  check("rozwiąż: akapit na tym samym miejscu ekranu", Math.abs((await paraY(page)) - y) <= 1);
  const z = await zipOf(page);
  const ex = await z.file("word/commentsExtended.xml").async("string");
  check("plik: odpowiedź powiązana z komentarzem (paraIdParent) i rozwiązany (done=1) po edycjach", /w15:paraIdParent/.test(ex) && /w15:done="1"/.test(ex), ex);

  // ── telefon: w trakcie pisania komentarza dokument się przesuwa (klawiatura, nagłówek) — po dodaniu
  // akapit wraca w to samo miejsce ekranu, w którym był przy otwieraniu okienka ────────────────
  await select(page, 20, 24);
  const ySpot = await paraY(page);
  await page.keyboard.press(`${MOD}+Alt+KeyM`);
  await page.waitForSelector(".compose-pop-comment .cf-rich");
  await page.keyboard.type("Przesunięcie w trakcie");
  // jak na telefonie: rama nad dokumentem zmienia wysokość (klawiatura chowa / pokazuje skróty sekcji, nagłówek)
  const stripWas = await page.evaluate(() => { const st = document.getElementById("sectionStrip"); const was = st.hidden; st.hidden = !was; return was; });
  await page.waitForTimeout(100);
  await page.evaluate(() => document.querySelector(".compose-pop-form .lf-ok").click());
  await idle(page);
  check("po dodaniu komentarza akapit wraca tam, gdzie był przy otwieraniu okienka (nawet gdy dokument się przesunął)", Math.abs((await paraY(page)) - ySpot) <= 2, `${ySpot} → ${await paraY(page)}`);
  await page.evaluate((was) => { document.getElementById("sectionStrip").hidden = was; }, stripWas);
  await page.waitForTimeout(200);

  // ── wklejanie z formatowaniem do komentarza ─────────────────────────────────
  await select(page, 3, 6);
  await page.keyboard.press(`${MOD}+Alt+KeyM`);
  await page.waitForSelector(".compose-pop-comment .cf-rich");
  await page.evaluate(() => {
    const dt = new DataTransfer();
    dt.setData("text/html", '<h2>Tytuł</h2><p>Zwykły <b>gruby</b> <i>pochyły</i> <span style="font-size:30px;color:red;font-family:Comic Sans MS">duży czerwony</span> <mark>zakreślony</mark> <s>skreślony</s> <a href="https://example.com">link</a><img src="x.png"></p><ul><li>punkt <u>pierwszy</u></li><li>drugi</li></ul>');
    dt.setData("text/plain", "Tytuł\nZwykły gruby pochyły duży czerwony zakreślony skreślony link\npunkt pierwszy\ndrugi");
    document.querySelector(".compose-pop-comment .cf-rich").dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  });
  await page.waitForTimeout(200);
  const pasted = await page.evaluate(() => ({ html: document.querySelector(".compose-pop-comment .cf-rich").innerHTML, img: !!document.querySelector(".compose-pop-comment .cf-rich img"), a: !!document.querySelector(".compose-pop-comment .cf-rich a") }));
  check("wklejenie: bez obrazka, linku, rozmiaru, kroju i koloru ze źródła", !pasted.img && !pasted.a && !/font-size|Comic|color:\s*red|rgb\(255, 0, 0\)/i.test(pasted.html), pasted.html);
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
  cx = await commentsXml(page);
  const pc = cx.slice(cx.lastIndexOf("<w:comment "));
  check("plik: wklejone formatowanie jak w Wordzie — nagłówek pogrubiony, B/I/U/przekreślenie/wyróżnienie, punkty jako akapity",
    /<w:b\/><\/w:rPr><w:t xml:space="preserve">Tytuł/.test(pc) && /<w:b\/><\/w:rPr><w:t xml:space="preserve">gruby/.test(pc) && /<w:i\/><\/w:rPr><w:t xml:space="preserve">pochyły/.test(pc)
    && /<w:highlight w:val="yellow"\/><\/w:rPr><w:t xml:space="preserve">zakreślony/.test(pc) && /<w:strike\/><\/w:rPr><w:t xml:space="preserve">skreślony/.test(pc)
    && /<w:u w:val="single"\/><\/w:rPr><w:t xml:space="preserve">pierwszy/.test(pc) && /• punkt/.test(pc) && (pc.match(/<w:p\b/g) || []).length === 4 && !/w:sz|w:rFonts|w:color/.test(pc), pc);

  // ── długi komentarz: „Więcej ↓” ─────────────────────────────────────────────
  await select(page, 6, 10);
  await page.keyboard.press(`${MOD}+Alt+KeyM`);
  await page.evaluate(() => { const ed = document.querySelector(".compose-pop-comment .cf-rich"); ed.innerHTML = Array.from({ length: 30 }, (_, i) => `<div>Wiersz ${i + 1} długiego komentarza, który nie mieści się w karcie.</div>`).join(""); });
  await page.waitForTimeout(150);
  check("edytor: długa treść — wygaszenie i „Więcej” także przy pisaniu", await page.evaluate(() => document.querySelector(".compose-pop-comment .cf-rich").classList.contains("more-b") && !document.querySelector(".compose-pop-comment .cc-more").hidden));
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
  await select(page, 7, 7);
  await page.waitForSelector(".comment-card");
  await page.waitForTimeout(150);
  const more = await page.evaluate(() => ({ cls: document.querySelector(".comment-card .cc-list").classList.contains("more-b"), btn: !document.querySelector(".comment-card .cc-more").hidden }));
  check("karta: długi komentarz ma wygaszenie i „Więcej ↓”", more.cls && more.btn, JSON.stringify(more));
  await page.click(".comment-card .cc-more");
  await page.waitForTimeout(500);
  check("„Więcej” przewija treść karty", await page.evaluate(() => document.querySelector(".comment-card .cc-list").scrollTop > 20));
  await page.evaluate(() => { const l = document.querySelector(".comment-card .cc-list"); l.scrollTop = l.scrollHeight; });
  await page.waitForTimeout(200);
  check("na końcu treści „Więcej” znika", await page.evaluate(() => document.querySelector(".comment-card .cc-more").hidden && !document.querySelector(".comment-card .cc-list").classList.contains("more-b")));
  await select(page, 1, 1);
  await page.waitForSelector(".comment-card");
  check("krótki komentarz: bez „Więcej”", await page.evaluate(() => document.querySelector(".comment-card .cc-more").hidden));

  // ── zoom: karta rośnie razem z dokumentem ───────────────────────────────────
  const fs100 = await page.evaluate(() => parseFloat(getComputedStyle(document.querySelector(".comment-card .cc-text")).fontSize));
  await page.evaluate(() => commitDocZoom(1.5));
  await page.waitForTimeout(300);
  await page.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].scrollIntoView({ block: "center" }), IDX);
  await page.waitForTimeout(200);
  await select(page, 1, 1);
  await page.waitForSelector(".comment-card");
  const fs150 = await page.evaluate(() => parseFloat(getComputedStyle(document.querySelector(".comment-card .cc-text")).fontSize));
  check("zoom 150%: tekst komentarza ×1,5", Math.abs(fs150 - fs100 * 1.5) < 0.6, `${fs100} → ${fs150}`);
  await page.evaluate(() => commitDocZoom(0.6));
  await page.waitForTimeout(300);
  const fs60 = await page.evaluate(() => parseFloat(getComputedStyle(document.querySelector(".comment-card .cc-text")).fontSize));
  check("zoom 60%: komentarz nie mniejszy niż zwykle (czytelność)", Math.abs(fs60 - fs100) < 0.6, `${fs60}`);
  await page.evaluate(() => commitDocZoom(1));
  await page.waitForTimeout(300);
  await page.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].scrollIntoView({ block: "center" }), IDX);
  await page.waitForTimeout(300);

  // ── usuń: bez skoku ──────────────────────────────────────────────────────────
  await select(page, 1, 1);
  await page.waitForSelector(".comment-card");
  y = await paraY(page);
  await page.click(".comment-card .btn.cc-del");
  await idle(page);
  check("usuń: akapit na tym samym miejscu ekranu", Math.abs((await paraY(page)) - y) <= 1);

  // ── komentarz z linkiem (z Worda) — edycja zablokowana, plik nietknięty ─────
  const blocked = await page.evaluate(async () => {
    const zip = await JSZip.loadAsync(originalFileBytes);
    let cx = await zip.file("word/comments.xml").async("string");
    cx = cx.replace(/(<w:comment [^>]*>)/, (m) => m); // bez zmian, tylko sprawdzenie, że jest
    const id = (cx.match(/<w:comment [^>]*w:id="(\d+)"/) || [])[1];
    cx = cx.replace(new RegExp(`(<w:comment [^>]*w:id="${id}"[^>]*>[\\s\\S]*?)(</w:p>)`), `$1<w:hyperlink w:anchor="x"><w:r><w:t>link</w:t></w:r></w:hyperlink>$2`);
    zip.file("word/comments.xml", cx);
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const r = await applyCommentEditInZip(await JSZip.loadAsync(bytes), "<x/>", { op: "commentEdit", id, text: "nadpisane" });
    const scan = await scanDocxRevisions(bytes);
    return { count: r.count, editable: scan.comments.concat(scan.comments.flatMap((c) => c.replies)).find((c) => c.id === id)?.editable };
  });
  check("komentarz z linkiem: model oznacza jako nieedytowalny, zapis go nie rusza", blocked.count === 0 && blocked.editable === false, JSON.stringify(blocked));

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : `\n   ${String(r.detail).slice(0, 700)}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
