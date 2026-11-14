// list-format-playwright.js — biblioteka list jak w Wordzie (compose-ui.js buildListMenu, docx-compose.js):
// format numeracji (a), I., 1.1.1…) i punktory z menu listy, zmiana formatu całej listy przy kursorze,
// „Kontynuuj numerację” i „Numeruj od nowa od 1”, Cofnij. Numery z podglądu (list-numbering.js) i plik.
//   node scripts/list-format-playwright.js        ENGINE=webkit node scripts/list-format-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(500));

(async () => {
  const bytes = await buildDocx({ body: ["Alfa", "Beta", "Gamma", "Przerwa w liście", "Delta", "Epsilon"] });
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "listy.docx", mimeType: DOCX_MIME, buffer: bytes });
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);

  const select = (a, b) => page.evaluate(([a, b]) => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    ps[a].closest(".docx-edit-root").focus();
    const r = document.createRange();
    r.setStart(ps[a].firstChild || ps[a], 0);
    if (b === a) r.collapse(true); else r.setEnd(ps[b], ps[b].childNodes.length);
    getSelection().removeAllRanges(); getSelection().addRange(r);
    document.dispatchEvent(new Event("selectionchange"));
  }, [a, b]);
  const chip = async (label) => {
    await page.click("#fmtListBtn");
    await page.waitForSelector(".list-chip");
    await page.locator(".list-chip", { hasText: new RegExp(`^${label.replace(/[.)]/g, "\\$&")}$`) }).click();
    await idle(page);
  };
  const menuItem = async (re) => {
    await page.click("#fmtListBtn");
    await page.waitForSelector(".compose-pop");
    await page.locator(".compose-pop .compose-item", { hasText: re }).click();
    await idle(page);
  };
  const nums = () => page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => (p.getAttribute("data-dwb-num") || (isListParagraph(p) ? "•" : "")).trim()));

  // format na zaznaczeniu: a)
  await select(0, 2);
  await chip("a)");
  check("zaznaczenie + „a)”: a) b) c)", JSON.stringify((await nums()).slice(0, 4)) === JSON.stringify(["a)", "b)", "c)", ""]), JSON.stringify(await nums()));
  // kursor w liście (bez zaznaczenia) + „I.” — cała lista
  await select(1, 1);
  await chip("I.");
  check("kursor w liście + „I.”: cała lista I. II. III. (jak w Wordzie)", JSON.stringify((await nums()).slice(0, 3)) === JSON.stringify(["I.", "II.", "III."]), JSON.stringify(await nums()));
  // druga lista po przerwie w tym samym formacie: od 1
  await select(4, 5);
  await chip("I.");
  let n = await nums();
  check("lista po przerwie w tym samym formacie: zaczyna od I.", n[4] === "I." && n[5] === "II.", JSON.stringify(n));
  // Kontynuuj numerację (Delta) → IV., V.
  await select(4, 4);
  await menuItem(/Kontynuuj numerację/);
  n = await nums();
  check("„Kontynuuj numerację”: IV. V.", n[4] === "IV." && n[5] === "V.", JSON.stringify(n));
  // Numeruj od nowa (Epsilon) → I.
  await select(5, 5);
  await menuItem(/Numeruj od nowa/);
  n = await nums();
  check("„Numeruj od nowa od 1” na Epsilon: IV. i I.", n[4] === "IV." && n[5] === "I.", JSON.stringify(n));
  // Cofnij
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  n = await nums();
  check("Cofnij: z powrotem IV. V.", n[4] === "IV." && n[5] === "V.", JSON.stringify(n));
  // 1.1.1 — wielopoziomowa: druga pozycja poziom niżej
  await select(0, 2);
  await chip("1.1.1");
  await select(1, 1);
  await page.keyboard.press("Tab");
  await idle(page);
  n = await nums();
  check("1.1.1: 1. / 1.1. / 2.", JSON.stringify(n.slice(0, 3)) === JSON.stringify(["1.", "1.1.", "2."]), JSON.stringify(n));
  // punktor ✓
  await select(3, 3);
  await chip("✓");
  const bullet = await page.evaluate(async () => {
    const d = await getDocumentXmlString(await buildDocumentForSave());
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const num = await z.file("word/numbering.xml").async("string");
    return { list: /<w:numPr>/.test(d.split("Przerwa")[0].slice(-400)), tick: /<w:lvlText w:val="✓"\/>/.test(num) };
  });
  check("punktor ✓ w pliku", bullet.list && bullet.tick, JSON.stringify(bullet));
  // plik: format a) / I. / 1.1.1 jak definicje Worda
  const file = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const num = await z.file("word/numbering.xml").async("string");
    return { roman: /<w:numFmt w:val="upperRoman"\/><w:lvlText w:val="%1\."\/>/.test(num), legal: /<w:lvlText w:val="%1\.%2\."\/>/.test(num), restart: /<w:startOverride w:val="1"\/>/.test(num) };
  });
  check("plik: I. (upperRoman „%1.”), 1.1. („%1.%2.”), numerowanie od nowa (startOverride)", file.roman && file.legal && file.restart, JSON.stringify(file));
  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
