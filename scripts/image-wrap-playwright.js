// image-wrap-playwright.js — zawijanie tekstu wokół obrazu jak w Wordzie (karta obrazu → „Zawijanie tekstu”,
// docx-compose.js composeWrapDrawing; tools/word-oracle/probe-image-wrap.js): W linii / Przyległe /
// Góra i dół / Za tekstem / Przed tekstem — plik (wp:anchor, wrap*, behindDoc, położenie) i podgląd,
// bieżący tryb zaznaczony w menu, powrót do „W linii”, Cofnij.
//   node scripts/image-wrap-playwright.js        ENGINE=webkit node scripts/image-wrap-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(500));

(async () => {
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);
  await page.keyboard.type("Akapit przed obrazem.");
  const img = await page.evaluateHandle(async () => {
    const c = document.createElement("canvas"); c.width = 400; c.height = 300;
    const g = c.getContext("2d"); g.fillStyle = "#3a6ea5"; g.fillRect(0, 0, 400, 300);
    const blob = await new Promise((r) => c.toBlob(r, "image/png"));
    return new File([blob], "obraz.png", { type: "image/png" });
  });
  await page.evaluate((f) => composeUi.insertImageFile(f), img);
  await idle(page);
  for (let i = 0; i < 4; i++) { await page.keyboard.type(`Dalszy tekst akapitu numer ${i + 1}, dość długi, żeby zawinął się w kilku wierszach obok obrazu.`); await page.keyboard.press("Enter"); }
  await idle(page);
  // obraz mniejszy (40%), żeby było obok niego miejsce na tekst
  // obraz „za tekstem” jest pod tekstem — wtedy jak użytkownik: prawy przycisk w akapicie → „Zaznacz obraz”
  const openCard = async () => {
    const box = await page.evaluate(() => { const i = document.querySelector(".docx-preview-host img"); i.scrollIntoView({ block: "center" }); const r = i.getBoundingClientRect(); const top = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2); return { x: r.left + r.width / 2, y: r.top + r.height / 2, onTop: top === i }; });
    if (box.onTop) await page.mouse.click(box.x, box.y);
    else {
      await page.mouse.click(box.x, box.y, { button: "right" });
      await page.waitForSelector(".doc-menu");
      await page.locator(".doc-menu button", { hasText: /Zaznacz obraz/ }).click();
    }
    await page.waitForSelector(".image-card");
    await page.waitForTimeout(250);
  };
  await openCard();
  await page.evaluate(() => { const r = document.querySelector(".image-card input[type=range]"); r.value = "40"; r.dispatchEvent(new Event("input", { bubbles: true })); r.dispatchEvent(new Event("change", { bubbles: true })); r.blur(); });
  await page.waitForTimeout(700);
  await idle(page);

  const anchor = () => page.evaluate(async () => {
    const xml = await getDocumentXmlString(await buildDocumentForSave());
    const a = (xml.match(/<wp:(anchor|inline)\b[^>]*>[\s\S]*?<\/wp:\1>/) || [""])[0];
    return { kind: (a.match(/^<wp:(anchor|inline)/) || [])[1], behind: (a.match(/behindDoc="(\d)"/) || [])[1], wrap: (a.match(/<wp:(wrap\w+)/) || [])[1] || "", h: (a.match(/<wp:positionH relativeFrom="(\w+)">/) || [])[1], v: (a.match(/<wp:positionV relativeFrom="(\w+)">/) || [])[1], order: (a.match(/<wp:(\w+)/g) || []).map((x) => x.slice(4)).join(",") };
  });
  const preview = () => page.evaluate(() => { const i = document.querySelector(".docx-preview-host img"); const a = i.closest("[data-dwb-anchor]"); return a ? JSON.parse(a.dataset.dwbAnchor) : null; });
  const pick = async (mode) => {
    await openCard();
    await page.click(".image-card .ic-wrap");
    await page.waitForSelector(".compose-pop-wrap");
    await page.click(`.compose-pop-wrap [data-wrap="${mode}"]`);
    await idle(page);
  };

  await openCard();
  await page.click(".image-card .ic-wrap");
  await page.waitForSelector(".compose-pop-wrap");
  const items = await page.evaluate(() => [...document.querySelectorAll(".compose-pop-wrap [data-wrap]")].map((b) => `${b.dataset.wrap}:${b.getAttribute("aria-checked")}`));
  check("menu: 5 trybów, „W linii z tekstem” zaznaczone", items.join(" ") === "inline:true square:false topBottom:false behind:false front:false", items.join(" "));
  await page.keyboard.press("Escape");

  await pick("square");
  let a = await anchor(), pv = await preview();
  check("Przyległe: wp:anchor, wrapSquare, kolumna / akapit, kolejność elementów jak w schemacie", a.kind === "anchor" && a.wrap === "wrapSquare" && a.behind === "0" && a.h === "column" && a.v === "paragraph" && /^anchor,simplePos,positionH,(posOffset|align),positionV,posOffset,extent,effectExtent,wrapSquare,docPr/.test(a.order), JSON.stringify(a));
  check("…podgląd: obraz pływający z oblewaniem", pv?.w === "wrapSquare" && !!(await page.evaluate(() => document.querySelector(".dwb-anchor-wrap"))), JSON.stringify(pv));
  await pick("behind");
  a = await anchor(); pv = await preview();
  check("Za tekstem: wrapNone + behindDoc 1, w podglądzie pod tekstem", a.wrap === "wrapNone" && a.behind === "1" && pv?.b === 1, JSON.stringify({ a, pv }));
  await pick("front");
  a = await anchor();
  check("Przed tekstem: wrapNone + behindDoc 0", a.wrap === "wrapNone" && a.behind === "0", JSON.stringify(a));
  await pick("topBottom");
  a = await anchor();
  check("Góra i dół: wrapTopAndBottom", a.wrap === "wrapTopAndBottom", JSON.stringify(a));
  await openCard();
  await page.click(".image-card .ic-wrap");
  await page.waitForSelector(".compose-pop-wrap");
  check("menu pokazuje bieżący tryb (Góra i dół)", await page.evaluate(() => document.querySelector('.compose-pop-wrap [data-wrap="topBottom"]').getAttribute("aria-checked") === "true"));
  await page.click('.compose-pop-wrap [data-wrap="inline"]');
  await idle(page);
  a = await anchor();
  check("W linii z tekstem: z powrotem wp:inline", a.kind === "inline" && !(await preview()), JSON.stringify(a));
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  a = await anchor();
  check("Cofnij: znów Góra i dół", a.kind === "anchor" && a.wrap === "wrapTopAndBottom", JSON.stringify(a));
  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
