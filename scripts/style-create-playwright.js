// style-create-playwright.js — własne style jak w Wordzie (docx-compose.js applyStyleInZip, lista stylów
// w compose-ui.js): „Nowy styl z zaznaczenia…” (XML stylu jak z Worda — tools/word-oracle/
// probe-style-create.js), styl pliku na liście i nadany innemu akapitowi, „Zaktualizuj styl do
// zaznaczenia” zmienia wszystkie akapity w stylu, Cofnij, nazwa zajęta.
//   node scripts/style-create-playwright.js        ENGINE=webkit node scripts/style-create-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(500));

(async () => {
  const body = [
    { pPr: '<w:jc w:val="center"/><w:spacing w:before="240" w:after="120"/>', raw: '<w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:color w:val="C00000"/><w:sz w:val="32"/></w:rPr><w:t>Wzorcowy akapit z wyglądem</w:t></w:r>' },
    { text: "Zwykły akapit do nadania stylu" },
    { text: "Trzeci akapit bez zmian" },
  ];
  const bytes = await buildDocx({ body });
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "style.docx", mimeType: DOCX_MIME, buffer: bytes });
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);

  const caret = (i) => page.evaluate((i) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; p.closest(".docx-edit-root").focus(); placeCaret(p, 3); document.dispatchEvent(new Event("selectionchange")); }, i);
  const look = (i) => page.evaluate((i) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    const tw = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); const t = tw.nextNode()?.parentElement || p;
    const a = getComputedStyle(p), b = getComputedStyle(t);
    return { align: a.textAlign, before: a.marginTop, font: b.fontFamily.split(",")[0].replace(/["']/g, ""), weight: b.fontWeight, color: b.color, size: b.fontSize };
  }, i);
  const file = () => page.evaluate(async () => {
    await mergeInlineEditsIntoBytes();
    const zip = await JSZip.loadAsync(originalFileBytes);
    const styles = await zip.file("word/styles.xml").async("string");
    const st = (styles.match(/<w:style [^>]*>(?:(?!<\/w:style>)[\s\S])*?<w:name w:val="Nagłówek firmy"\/>[\s\S]*?<\/w:style>/) || [""])[0];
    const doc = await zip.file("word/document.xml").async("string");
    const paras = doc.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
    return { st, p: paras.slice(0, 3) };
  });
  const pick = (value) => page.evaluate((v) => { const s = document.getElementById("fmtParaStyle"); s.value = v; s.dispatchEvent(new Event("change", { bubbles: true })); }, value);

  // lista: polecenia stylu
  await caret(0);
  await page.waitForTimeout(300);
  const acts = await page.evaluate(() => [...document.querySelectorAll('#fmtParaStyle optgroup[data-act] option')].map((o) => o.textContent));
  check("lista stylów: „Nowy styl z zaznaczenia…” i „Zaktualizuj „Normalny” do zaznaczenia”", acts[0] === "Nowy styl z zaznaczenia…" && /Zaktualizuj „Normalny”/.test(acts[1] || ""), JSON.stringify(acts));

  // Nowy styl z zaznaczenia — akapit 1 (Georgia, pogrubiony, czerwony, 16 pt, do środka, odstęp przed)
  const before = await look(0);
  await page.evaluate(() => { window.__dwbAutoDialog = ({ input }) => (input ? "Nagłówek firmy" : true); });
  await pick("act:new");
  await idle(page);
  let f = await file();
  check("styl w pliku jak z Worda: bazowy Normal, qFormat, pPr (odstępy, środek), rPr (Georgia, pogrubienie, kolor, 16 pt)",
    /w:customStyle="1"/.test(f.st) && /w:styleId="Nagwekfirmy"/.test(f.st) && /<w:basedOn w:val="Normal"\/>/.test(f.st) && /<w:qFormat\/>/.test(f.st) && !/<w:next /.test(f.st)
      && /<w:spacing w:before="240" w:after="120"\/>/.test(f.st) && /<w:jc w:val="center"\/>/.test(f.st) && /Georgia/.test(f.st) && /<w:b\/>/.test(f.st) && /C00000/.test(f.st) && /<w:sz w:val="32"\/>/.test(f.st), f.st);
  check("akapit dostał styl, formatowanie wprost przeszło do stylu", /<w:pStyle w:val="Nagwekfirmy"\/>/.test(f.p[0]) && !/<w:jc |<w:spacing |Georgia|<w:b\/>|C00000/.test(f.p[0]), f.p[0]);
  const after = await look(0);
  check("…akapit wygląda tak samo jak przed", JSON.stringify(before) === JSON.stringify(after), JSON.stringify([before, after]));
  // styl na liście i nadany drugiemu akapitowi
  await caret(1);
  await page.waitForTimeout(300);
  const listed = await page.evaluate(() => [...document.querySelectorAll('#fmtParaStyle optgroup[data-file] option')].map((o) => o.value));
  check("„Nagłówek firmy” w grupie „Style tego dokumentu”", listed.includes("custom:Nagłówek firmy"), JSON.stringify(listed));
  await pick("custom:Nagłówek firmy");
  await idle(page);
  const p2 = await look(1);
  check("styl nadany innemu akapitowi: ten sam wygląd", p2.font === "Georgia" && p2.align === "center" && Number(p2.weight) >= 600 && p2.color === after.color && p2.size === after.size, JSON.stringify(p2));

  // Zaktualizuj styl do zaznaczenia: akapit 1 do prawej → styl do prawej, akapit 2 też
  await page.evaluate(() => applyDocumentEdit({ op: "paraFormat", indices: [0], align: "right" }));
  await idle(page);
  await caret(0);
  await page.waitForTimeout(300);
  await pick("act:update");
  await idle(page);
  f = await file();
  const a2 = await look(1);
  check("Zaktualizuj: styl do prawej, drugi akapit w tym stylu też, w akapicie 1 bez jc wprost", /<w:jc w:val="right"\/>/.test(f.st) && a2.align === "right" && !/<w:jc /.test(f.p[0]), JSON.stringify({ st: f.st.slice(0, 300), a2 }));
  // Cofnij — styl wraca do środka
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  check("Cofnij: styl z powrotem do środka (akapit 2)", (await look(1)).align === "center", JSON.stringify(await look(1)));
  // nazwa zajęta — drugi „Nagłówek firmy” nie powstaje
  await caret(2);
  let asked = 0;
  await page.exposeFunction("__nameAsked", () => { asked++; });
  await page.evaluate(() => { let n = 0; window.__dwbAutoDialog = ({ input }) => { if (!input) return true; window.__nameAsked(); return n++ ? null : "nagłówek FIRMY"; }; });
  await pick("act:new");
  await idle(page);
  const count = await page.evaluate(async () => ((await (await JSZip.loadAsync(originalFileBytes)).file("word/styles.xml").async("string")).match(/w:val="Nagłówek firmy"/g) || []).length);
  check("nazwa zajęta (bez względu na wielkość liter): pytanie jeszcze raz, styl nie dubluje się", asked === 2 && count === 1, JSON.stringify({ asked, count }));

  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
