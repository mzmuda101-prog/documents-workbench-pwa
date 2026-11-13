// track-edit-playwright.js — pisanie w akapicie ze śledzonymi zmianami (jak w Wordzie, 2026-11-13).
//
// Dawniej akapit ze wstawką / usunięciem był tylko do odczytu. Teraz (docx-inline-edit.js
// paragraphRevisionLock + stampRevisionWraps, docx-run-styles.js revKey): usunięcie = wyspa (zostaje
// w pliku w tym samym miejscu), tekst wpisany WE wstawce zostaje we wstawce (ten sam autor i data),
// obok niej — zwykły tekst. Sprawdza PLIK po każdej operacji. Przy włączonym „Śledź zmiany” akapit
// z cudzymi zmianami jest zablokowany (poprawki liczy różnica plików — docx-redline.js).
// ENGINE=webkit (Safari).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(600));

// akapit pliku (po scaleniu edycji) zaczynający się od tekstu — XML i zawartość wstawek/usunięć
const para = (page, start) => page.evaluate(async (start) => {
  await mergeInlineEditsIntoBytes();
  const d = await getDocumentXmlDom(originalFileBytes);
  const p = collectParagraphElements(d.documentElement, "all").find((x) => getParagraphText(x).startsWith(start));
  if (!p) return null;
  const pick = (tag) => Array.from(p.getElementsByTagNameNS(W_NS, tag)).filter((e) => e.parentNode === p).map((e) => ({ author: e.getAttributeNS(W_NS, "author"), date: e.getAttributeNS(W_NS, "date"), text: revText(e) }));
  return { text: getParagraphText(p), ins: pick("ins"), del: pick("del") };
}, start);
const caretIn = (page, start, rev, at) => page.evaluate(([start, rev, at]) => {
  const p = [...document.querySelectorAll(".docx-preview-host p")].find((x) => x.textContent.startsWith(start));
  const r = document.createRange();
  if (rev) { const w = p.querySelector("[data-dwb-rev]"); r.setStart(w.firstChild, at); }
  else r.setStart(...(() => { const tw = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); let n, pos = 0; while ((n = tw.nextNode())) { if (n.parentElement.closest("[data-cm]")) continue; if (pos + n.data.length >= at) return [n, at - pos]; pos += n.data.length; } return [p, 0]; })());
  r.collapse(true);
  p.closest(".docx-edit-root").focus({ preventScroll: true });
  getSelection().removeAllRanges();
  getSelection().addRange(r);
}, [start, rev, at]);

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.setItem("dwb.authorName", "Tester"); window.__dwbAutoDialog = "accept"; });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles("docs/samples/review-sample.docx");
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);

  const p0 = await para(page, "Termin");
  const locks = await page.evaluate(() => Object.fromEntries([...document.querySelectorAll(".docx-preview-host p")].map((p) => [p.textContent.slice(0, 12), p.dataset.lock || "-"])));
  check("akapit ze wstawką i usunięciem edytowalny", locks["Termin płatn"] === "-", JSON.stringify(locks));
  check("akapit ze wstawką i przypisem edytowalny", locks["Zleceniodawc"] === "-", JSON.stringify(locks));
  check("zmiana formatu fragmentu dalej zablokowana", locks["Pogrubione p"] === "lockTracked", JSON.stringify(locks));

  // ── pisanie WE wstawce: zostaje we wstawce, ten sam autor i data ──────────
  await caretIn(page, "Termin", true, 1);
  await page.keyboard.type("5");
  await idle(page);
  const p1 = await para(page, "Termin");
  check("wpisane we wstawce zostaje we wstawce", p1?.ins.length === 1 && p1.ins[0].text === "350", JSON.stringify(p1));
  check("…z tym samym autorem i datą", p1?.ins[0]?.author === p0.ins[0].author && p1?.ins[0]?.date === p0.ins[0].date, JSON.stringify([p0.ins, p1?.ins]));
  check("usunięcie w akapicie nietknięte", JSON.stringify(p1?.del) === JSON.stringify(p0.del), JSON.stringify([p0.del, p1?.del]));

  // ── pisanie OBOK wstawki: zwykły tekst ─────────────────────────────────────
  await page.keyboard.press("End");
  await page.keyboard.type(" roboczych");
  await idle(page);
  const p2 = await para(page, "Termin");
  check("dopisane za wstawką = zwykły tekst", p2?.text.endsWith("dni. roboczych") && p2.ins.length === 1 && p2.ins[0].text === "350", JSON.stringify(p2));

  // ── Backspace tuż za miejscem usunięcia: usunięcie zostaje ──────────────────
  await caretIn(page, "Termin", false, "Termin płatności wynosi ".length);
  await page.keyboard.press("Backspace");
  await idle(page);
  const p3 = await para(page, "Termin");
  check("Backspace przy usunięciu: usunięcie zostaje w pliku", p3?.del.length === 1 && p3.del[0].text === "14" && p3.text.startsWith("Termin płatności wynosi350"), JSON.stringify(p3));

  // ── Enter w akapicie ze zmianami: obie części, nic nie ginie ───────────────
  await caretIn(page, "Termin", false, 7);
  await page.keyboard.press("Enter");
  await idle(page);
  const split = await page.evaluate(async () => {
    const d = await getDocumentXmlDom(originalFileBytes);
    const x = new XMLSerializer().serializeToString(d);
    return { ins350: (x.match(/<w:ins\b[^>]*>(?:(?!<\/w:ins>).)*350/gs) || []).length, del14: (x.match(/<w:del\b[^>]*>(?:(?!<\/w:del>).)*>14</gs) || []).length, texts: collectParagraphElements(d.documentElement, "all").map(getParagraphText).filter((t) => /Termin|płatności/.test(t)) };
  });
  check("Enter: wstawka i usunięcie przechodzą do drugiej części", split.ins350 === 1 && split.del14 === 1 && split.texts.length === 2, JSON.stringify(split));

  // ── Cofnij wszystko = plik jak po otwarciu ─────────────────────────────────
  for (let i = 0; i < 6; i++) { await page.evaluate(() => dwbUndo.undo()); await idle(page); }
  const back = await para(page, "Termin");
  check("Cofnij: akapit jak w oryginale (tekst, wstawka, usunięcie)", JSON.stringify(back) === JSON.stringify(p0), JSON.stringify([p0, back]));

  // ── po edycji Akceptuj / Odrzuć dalej działa ───────────────────────────────
  await caretIn(page, "Termin", true, 1);
  await page.keyboard.type("5");
  await idle(page);
  await page.evaluate(async () => { await mergeInlineEditsIntoBytes(); await applyDocumentEdit({ op: "revisions", action: "reject", ids: null }); });
  await idle(page);
  const rej = await para(page, "Termin");
  check("Odrzuć wszystkie po edycji: wraca „14”, wstawka (z dopisanym) znika", rej?.text === "Termin płatności wynosi 14 dni." && !rej.ins.length && !rej.del.length, JSON.stringify(rej));
  await page.evaluate(() => dwbUndo.undo());
  await idle(page);

  // ── zapis → ponowne otwarcie: zmiany i autorzy jak w pliku ─────────────────
  const reopen = await page.evaluate(async () => {
    const bytes = await buildDocumentForSave();
    const a = await scanDocxRevisions(bytes);
    await ingestFile(new File([bytes], "po-edycji.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
    await new Promise((r) => setTimeout(r, 800));
    const b = await scanDocxRevisions(originalFileBytes);
    return { a: a.changes.length, b: b.changes.length, authors: [...b.authors.keys()].sort().join(",") };
  });
  check("zapis i ponowne otwarcie: ta sama liczba zmian, ci sami autorzy", reopen.a === reopen.b && reopen.authors === "Anna Nowak,Jan Kowalski", JSON.stringify(reopen));
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);

  // ── Śledź zmiany włączone: akapit z cudzymi zmianami zablokowany ───────────
  await page.evaluate(() => dwbTrack.start());
  await idle(page);
  await page.waitForTimeout(500);
  const on = await page.evaluate(() => [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.startsWith("Termin"))?.dataset.lock || "-");
  check("Śledź zmiany włączone: akapit z cudzymi zmianami tylko do odczytu", on === "lockTrackedOn", on);
  await page.evaluate(() => dwbTrack.stop());
  await idle(page);

  // ── przełącznik trybu: Sugerowanie = Edycja + śledzenie; Edycja kończy śledzenie ──
  const ctxMenu = async (start) => {
    const pt = await page.evaluate((start) => { const r = [...document.querySelectorAll(".docx-preview-host p")].find((p) => p.textContent.startsWith(start)).getBoundingClientRect(); return { x: r.left + 20, y: r.top + r.height / 2 }; }, start);
    await page.mouse.click(pt.x, pt.y);
    await page.waitForTimeout(250);
    await page.mouse.click(pt.x, pt.y, { button: "right" });
    await page.waitForSelector(".doc-menu");
    return page.evaluate(() => [...document.querySelectorAll(".doc-menu button")].filter((b) => !b.closest(".doc-menu-row")).map((b) => b.textContent.trim()).filter(Boolean));
  };
  const clickMenu = (re) => page.evaluate((src) => { const re = new RegExp(src); [...document.querySelectorAll(".doc-menu button")].find((b) => re.test(b.textContent))?.click(); }, re.source);
  await page.click('.mode-btn[data-mode="suggest"]');
  await idle(page);
  check("przełącznik „Sugerowanie”: Edycja ze śledzeniem zmian", await page.evaluate(() => dwbTrack.isOn() && !readOnlyMode && document.querySelector(".mode-btn.is-active")?.dataset.mode === "suggest"));
  await page.click('.mode-btn[data-mode="edit"]');
  await idle(page);
  check("…„Edycja” kończy śledzenie", await page.evaluate(() => !dwbTrack.isOn() && document.querySelector(".mode-btn.is-active")?.dataset.mode === "edit"));

  // ── menu kontekstowe: zmiany i komentarze ─────────────────────────────────
  let menu = await ctxMenu("Termin");
  check("menu kontekstowe na akapicie ze zmianą: Akceptuj / Odrzuć na górze", /Akceptuj zmian/.test(menu[0]) && /Odrzuć zmian/.test(menu[1]), menu.slice(0, 4).join(" | "));
  const before = await para(page, "Termin");
  await clickMenu(/Akceptuj zmian/);
  await idle(page);
  const acc = await para(page, "Termin");
  check("…Akceptuj z menu: zmiany akapitu przyjęte w pliku", before.ins.length + before.del.length > 0 && !acc.ins.length && !acc.del.length && acc.text === before.text, JSON.stringify([before, acc]));
  menu = await ctxMenu("Kara");
  check("menu kontekstowe w komentowanym tekście: Odpowiedz / Rozwiąż / Usuń", ["Odpowiedz na komentarz", "Rozwiąż komentarz", "Usuń komentarz"].every((x) => menu.includes(x)), menu.join(" | "));
  await clickMenu(/^Rozwiąż komentarz/);
  await idle(page);
  check("…Rozwiąż z menu: komentarz rozwiązany w pliku", await page.evaluate(async () => /w15:done="1"/.test(await (await JSZip.loadAsync(originalFileBytes)).file("word/commentsExtended.xml").async("string"))));
  menu = await ctxMenu("Ostatni");
  check("menu kontekstowe: „Sugeruj zmiany”", menu.some((x) => /^Sugeruj zmiany/.test(x)), menu.join(" | "));
  await clickMenu(/^Sugeruj zmiany/);
  await idle(page);
  check("…włącza śledzenie i przełącznik pokazuje Sugerowanie", await page.evaluate(() => dwbTrack.isOn() && document.querySelector(".mode-btn.is-active")?.dataset.mode === "suggest"));
  await page.evaluate(() => dwbTrack.stop());
  await idle(page);

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
