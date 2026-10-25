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

  await browser.close();
  if (fails.length) { console.error(`❌ bughunt-fixes (${ENGINE}): ${fails.length} błąd(y)`); process.exit(1); }
  console.log(`✅ bughunt-fixes (${ENGINE}) OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
