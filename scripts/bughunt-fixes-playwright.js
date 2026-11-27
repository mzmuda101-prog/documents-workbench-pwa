// Strażnik poprawek z polowania na błędy (2026-10-25) — funkcje zapisu wołane wprost na stronie.
//   node scripts/bughunt-fixes-playwright.js        ENGINE=webkit node scripts/bughunt-fixes-playwright.js
const playwright = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE || "chromium";

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${detail}`}`);
    if (!ok) fails.push(name);
  };

  // 1) funkcje zapisu
  const page = await browser.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.waitForFunction(() => typeof mergeParagraphInXml === "function" && typeof applyReplaceInXml === "function");
  const r = await page.evaluate(() => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const wrap = (body) => `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;
    const out = {};
    // Backspace na początku akapitu kończącego sekcję poziomą — podział sekcji przechodzi wyżej
    const sect = '<w:pPr><w:sectPr><w:pgSz w:orient="landscape"/></w:sectPr></w:pPr>';
    const m = mergeParagraphInXml(wrap(`<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p>${sect}<w:r><w:t>b</w:t></w:r></w:p><w:p><w:r><w:t>c</w:t></w:r></w:p>`), 1);
    out.mergeSect = m.count === 1 && /orient="landscape"/.test(m.xml) && (m.xml.match(/<w:p>/g) || []).length === 2 ? "" : m.xml;
    // usunięcie zaznaczenia przez akapity (zostaje pierwszy) — sectPr ostatniego zostaje
    const d = deleteParagraphRangeInXml(wrap(`<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>x</w:t></w:r></w:p><w:p>${sect}<w:r><w:t>b</w:t></w:r></w:p>`), { from: 0, to: 2, keep: "first", mergedRuns: [{ text: "ab" }] });
    out.rangeSect = d.count === 1 && /orient="landscape"/.test(d.xml) ? "" : d.xml;
    // Znajdź nie przechodzi przez łamanie wiersza / tabulator (jak Word)
    const br = wrap("<w:p><w:r><w:t>ab</w:t><w:br/><w:t>cd</w:t><w:tab/><w:t>ef</w:t></w:r></w:p>");
    const counts = ["bc", "de", "cd", "ab"].map((find) => applyReplaceInXml(br, { find, replace: "X" }, "all").count);
    out.findBreak = counts.join() === "0,0,1,1" ? "" : counts.join();
    // jawnie wyłączone pogrubienie (Ctrl+B w pogrubionym nagłówku) zostaje w pliku
    const doc = new DOMParser().parseFromString(wrap('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:lang w:val="pl-PL"/></w:rPr><w:t>Tytuł</w:t></w:r></w:p>'), "application/xml");
    const p = doc.getElementsByTagNameNS(W, "p")[0];
    applyRunsToParagraphXml(p, [{ text: "Tytuł" }, { text: " zwykły", bold: false }]);
    const ps = new XMLSerializer().serializeToString(p);
    out.boldOff = /<w:b w:val="(0|false)"\/>[\s\S]*zwykły/.test(ps) && /w:lang/.test(ps) ? "" : ps;
    // $10 przy jednej grupie = grupa 1 + „0” (jak String.replace)
    const re = applyReplaceInXml(wrap("<w:p><w:r><w:t>ab</w:t></w:r></w:p>"), { find: "(a)b", replace: "$10", regex: true }, "all");
    out.dollar10 = /<w:t[^>]*>a0<\/w:t>/.test(re.xml) ? "" : re.xml;
    // „Całe słowa” z wyrażeniem poprawnym tylko bez flagi u
    const ww = buildFindRegex({ find: "a\\-b", regex: true, wholeWord: true });
    out.wholeWordRegex = ww && "x a-b y".match(ww) && !"za-bc".match(ww) ? "" : String(ww);
    // znaki sterujące ze schowka Office nie trafiają do XML
    out.composeSanitize = composeXmlText("a\u000Bb<") === "ab&lt;" ? "" : JSON.stringify(composeXmlText("a\u000Bb<"));
    return out;
  });
  for (const [k, v] of Object.entries(r)) check(k, !v, v);
  check("bez wyjątków na stronie", !errors.length, errors.join(" | "));
  await page.close();

  // 2) zablokowana pamięć przeglądarki (Safari „Blokuj wszystkie cookie”) — aplikacja wstaje
  const ctx = await browser.newContext();
  await ctx.addInitScript(() => {
    Object.defineProperty(window, "localStorage", { get() { throw new DOMException("The operation is insecure.", "SecurityError"); }, configurable: true });
  });
  const p2 = await ctx.newPage();
  const errs2 = [];
  p2.on("pageerror", (e) => errs2.push(e.message));
  await p2.goto(APP_URL, { waitUntil: "load" });
  await p2.waitForTimeout(800);
  const booted = await p2.evaluate(() => typeof t === "function" && !!document.getElementById("fileInput") && typeof currentLang === "string");
  check("start przy zablokowanym localStorage", booted && !errs2.length, `booted=${booted} ${errs2.join(" | ")}`);
  await ctx.close();

  // 3) (chaos 2026-10-28) Esc w tekście wychodzi z pisania — WebKit dalej zmieniał zaznaczony tekst
  // (Enter dzielił akapit tylko w podglądzie, nie w pliku); teraz klawisze nic nie robią, jak w Chromium
  const p3 = await browser.newPage();
  await p3.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  await p3.goto(APP_URL, { waitUntil: "load" });
  await p3.evaluate(() => { document.getElementById("heroSplash")?.remove(); return composeUi.createNew("note"); });
  await p3.waitForSelector(".docx-preview-host p", { timeout: 30000 });
  await p3.waitForFunction(() => !inlineLocksPending && document.getElementById("loadingOverlay")?.classList.contains("hidden"));
  await p3.waitForTimeout(400);
  const escCase = async (key) => {
    const pt = await p3.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((q) => q.textContent.trim().length > 6); const r = p.getBoundingClientRect(); return { x: r.left + 3, y: r.top + r.height / 2 }; });
    await p3.mouse.move(pt.x, pt.y); await p3.mouse.down(); await p3.mouse.move(pt.x + 50, pt.y, { steps: 3 }); await p3.mouse.up();
    await p3.keyboard.press("Escape");
    await p3.waitForTimeout(150);
    const before = await p3.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => p.textContent).join("¶"));
    if (key === "type") await p3.keyboard.type("XY"); else await p3.keyboard.press(key);
    await p3.waitForTimeout(500);
    const after = await p3.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => p.textContent).join("¶"));
    const ck = (await p3.evaluate(() => dwbSelfCheck())).problems.filter((x) => x.level === "error").map((x) => x.msg);
    check(`Esc, potem ${key}: dokument bez zmian, podgląd = plik`, before === after && !ck.length, `${ck.join(" | ")} ${before === after ? "" : "zmiana tekstu"}`);
  };
  for (const key of ["Enter", "Shift+Enter", "type"]) await escCase(key);

  // 4) (testerzy AI 2026-11-27, AUDYT-MOBILE #3) zaznaczony cały tekst akapitu → Backspace zostawia
  // <br> przeglądarki w pustym akapicie; Delete w nim sklejał akapity tylko w podglądzie — zapis
  // przesuwał tekst wszystkich dalszych akapitów o jeden (tabela pod spodem: treść w złych komórkach)
  const target = await p3.evaluate(() => {
    const list = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const i = list.findIndex((q, k) => k > 0 && q.textContent.trim().length > 6 && list[k + 1]?.textContent.trim().length > 3 && (q.dataset.box || "") === (list[k + 1].dataset.box || "") && !q.closest("td") && !q.dataset.lock);
    const r = list[i].getBoundingClientRect();
    return { i, n: list.length, x: r.left + 2, y: r.top + Math.min(8, r.height / 2) };
  });
  await p3.mouse.click(target.x, target.y);
  await p3.keyboard.press("Home");
  await p3.keyboard.press("Shift+End");
  await p3.keyboard.press("Backspace");
  await p3.waitForTimeout(400);
  const emptied = await p3.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].textContent.replace(/﻿/g, ""), target.i);
  await p3.keyboard.press("Delete");
  await p3.waitForTimeout(800);
  await p3.waitForFunction(() => !inlineStructuralPending);
  const del = await p3.evaluate(async () => {
    const preview = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => p.textContent.replace(/﻿/g, "").trim());
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const xml = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const file = [...xml.getElementsByTagNameNS(W, "body")[0].getElementsByTagNameNS(W, "p")].map((p) => [...p.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join("").trim());
    const ck = (await dwbSelfCheck()).problems.filter((x) => x.level === "error").map((x) => x.msg);
    return { preview, file, ck };
  });
  check("pusty akapit po zaznaczeniu i Backspace: Delete skleja w pliku tak jak w podglądzie",
    emptied === "" && del.preview.length === target.n - 1 && del.file.length === del.preview.length && del.file.every((t, k) => t === del.preview[k]) && !del.ck.length,
    `pusty=${JSON.stringify(emptied)} podgląd ${del.preview.length} (było ${target.n}), plik ${del.file.length}; pierwsza różnica ${del.file.findIndex((t, k) => t !== del.preview[k])}; ${del.ck.join(" | ")}`);
  await p3.close();

  // 5) (testerzy AI 2026-11-27, AUDYT-MOBILE #11) lista punktowana: zaznaczenie całego tekstu
  // punktu i Backspace na WebKit usuwało cały element <p> (zostawało gołe <br> w <article>),
  // kursor wypadał poza akapity i pisanie / kolejny Backspace przestawały działać.
  const ctx5 = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 }, hasTouch: true });
  await ctx5.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const p5 = await ctx5.newPage();
  const errs5 = [];
  p5.on("pageerror", (e) => errs5.push(e.message));
  await p5.goto(APP_URL, { waitUntil: "load" });
  await p5.evaluate(() => { document.getElementById("heroSplash")?.remove(); return composeUi.createNew("blank"); });
  await p5.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 30000 });
  await p5.waitForFunction(() => !inlineLocksPending && document.getElementById("loadingOverlay")?.classList.contains("hidden"));
  await p5.waitForTimeout(300);
  await p5.keyboard.type("Lista");
  await p5.keyboard.press("Enter");
  await p5.evaluate(() => document.getElementById("fmtListBtn")?.click());
  await p5.waitForTimeout(150);
  await p5.evaluate(() => document.querySelector(".compose-pop button")?.click());
  await p5.waitForTimeout(300);
  await p5.keyboard.type("Jedzenie");
  await p5.keyboard.press("Enter");
  await p5.keyboard.press("Tab");
  await p5.waitForTimeout(150);
  await p5.keyboard.type("kanapki");
  await p5.keyboard.press("Enter");
  await p5.keyboard.type("woda");
  await p5.waitForTimeout(300);
  const kp5 = await p5.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((q) => q.textContent.includes("kanapki"));
    const r = p.getBoundingClientRect();
    return { x: r.left, y: r.top + r.height / 2, w: r.width };
  });
  await p5.mouse.move(kp5.x - 8, kp5.y);
  await p5.mouse.down();
  await p5.mouse.move(kp5.x + kp5.w + 8, kp5.y, { steps: 10 });
  await p5.mouse.up();
  await p5.waitForTimeout(150);
  await p5.keyboard.press("Backspace");
  await p5.waitForTimeout(300);
  const listAfterBs = await p5.evaluate(() => {
    const host = document.querySelector(".docx-preview-host");
    const ps = collectPreviewParagraphElements(host);
    const s = window.getSelection();
    const n = s?.rangeCount ? s.getRangeAt(0).startContainer : null;
    const inP = (n?.nodeType === 1 ? n : n?.parentElement)?.closest?.(".docx-editable-p") || null;
    return { count: ps.length, inP: !!inP, isList: !!inP && isListParagraph(inP), bareBr: !!host.querySelector("article > br") };
  });
  await p5.keyboard.type("owoce");
  await p5.waitForTimeout(300);
  const listAfterType = await p5.evaluate(async () => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => p.textContent.replace(/\uFEFF/g, "").trim());
    const ck = (await dwbSelfCheck()).problems.filter((x) => x.level === "error").map((x) => x.msg);
    return { ps, ck };
  });
  check("lista: zaznaczenie całego punktu + Backspace zostawia punkt z kursorem i pozwala pisać (#11)",
    listAfterBs.count === 4 && listAfterBs.inP && listAfterBs.isList && !listAfterBs.bareBr
    && listAfterType.ps[2] === "owoce" && !listAfterType.ck.length && !errs5.length,
    `${JSON.stringify(listAfterBs)} ${JSON.stringify(listAfterType)} ${errs5.join(" | ")}`);
  await ctx5.close();

  // 6) (testerzy AI 2026-11-27, AUDYT-MOBILE #9) komentarz przy kursorze w przypisie nie rzuca
  // „p.classList”, a zaznaczenie w zablokowanym akapicie (lockPageBreak) nie przeskakuje do
  // poprzedniego edytowalnego akapitu i nie otwiera okna komentarza przy złym słowie.
  const ctx6 = await browser.newContext({ serviceWorkers: "block", viewport: { width: 820, height: 1180 }, hasTouch: true });
  await ctx6.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const p6 = await ctx6.newPage();
  const errs6 = [];
  p6.on("pageerror", (e) => errs6.push(e.message));
  await p6.goto(APP_URL, { waitUntil: "load" });
  await p6.evaluate(() => document.getElementById("heroSplash")?.remove());
  await p6.locator("#fileInput").setInputFiles("docs/samples/notes-sample.docx");
  await p6.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await p6.waitForTimeout(400);
  await p6.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await p6.waitForFunction(() => !inlineLocksPending && !!document.querySelector(".docx-preview-host .docx-editable-p"), null, { timeout: 15000 });
  await p6.waitForTimeout(300);
  const fnComment = await p6.evaluate(() => {
    const fnP = document.querySelector(".docx-preview-host ol li p.docx-editable-p");
    fnP.focus();
    const r = document.createRange(); r.setStart(fnP, 0); r.collapse(true);
    const s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
    try { composeUi.openCommentForm(); return { ok: true, pop: !!document.querySelector(".compose-pop-comment") }; }
    catch (e) { return { ok: false, err: e.message }; }
  });
  check("komentarz przy kursorze w przypisie bez błędu p.classList (#9)",
    fnComment.ok && !fnComment.pop && !errs6.length, `${JSON.stringify(fnComment)} ${errs6.join(" | ")}`);

  const pts6 = await p6.evaluate(() => {
    const ps = [...document.querySelectorAll(".docx-preview-host section.docx > article > p")];
    const p1 = ps.find((p) => p.textContent.includes("pozycja"));
    const pLock = ps.find((p) => p.classList.contains("docx-locked-p"));
    const r1 = p1.getBoundingClientRect(), rL = pLock.getBoundingClientRect();
    return { p1: { x: r1.x + 120, y: r1.y + 8 }, pLock: { x: rL.x + 4, y: rL.y + 8 } };
  });
  await p6.touchscreen.tap(pts6.p1.x, pts6.p1.y);
  await p6.waitForTimeout(150);
  await p6.mouse.move(pts6.pLock.x, pts6.pLock.y);
  await p6.mouse.down();
  await p6.mouse.move(pts6.pLock.x + 90, pts6.pLock.y, { steps: 10 });
  await p6.mouse.up();
  await p6.waitForTimeout(200);
  const lockedSel = await p6.evaluate(() => {
    const miniShown = !document.querySelector(".mini-bar")?.hidden;
    composeUi.openCommentForm();
    const pop = !!document.querySelector(".compose-pop-comment");
    const toastTxt = [...document.querySelectorAll(".toast")].map((x) => x.textContent).join(" | ");
    return { miniShown, pop, toastTxt };
  });
  check("zaznaczenie w zablokowanym akapicie nie otwiera komentarza na poprzednim słowie ani mini-paska (#9)",
    !lockedSel.miniShown && !lockedSel.pop && !errs6.length && /podział strony|page break/i.test(lockedSel.toastTxt),
    `${JSON.stringify(lockedSel)} ${errs6.join(" | ")}`);
  await ctx6.close();

  await browser.close();
  if (fails.length) { console.error(`❌ bughunt-fixes (${ENGINE}): ${fails.length} błąd(y)`); process.exit(1); }
  console.log(`✅ bughunt-fixes (${ENGINE}) OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
