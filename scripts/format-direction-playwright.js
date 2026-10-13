// format-direction-playwright.js — formatowanie zaznaczenia niezależnie od KIERUNKU zaznaczania.
// Zgłoszenie 2026-10-13: kolor przy zaznaczaniu od prawej do lewej działał, od lewej do prawej
// zwykle nie — przeciągnięcie do końca wiersza kończy zaznaczenie na początku NASTĘPNEGO akapitu,
// a kolor/wyróżnienie/komentarz brały akapit z końca zaznaczenia. Każdy przycisk × 4 rodzaje
// zaznaczenia; w pliku sformatowane mają być dokładnie zaznaczone znaki.
//   node scripts/format-direction-playwright.js / ENGINE=webkit node scripts/format-direction-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ACTIONS = {
  bold: { click: "#fmtBold", re: /<w:b\b(?![^>]*w:val="(0|false)")/ },
  italic: { click: "#fmtItalic", re: /<w:i\b(?![^>]*w:val="(0|false)")/ },
  underline: { click: "#fmtUnderline", re: /<w:u w:val="(?!none)/ },
  strike: { click: "#fmtStrike", re: /<w:strike\b(?![^>]*w:val="(0|false)")/ },
  super: { click: "#fmtSuper", re: /<w:vertAlign w:val="superscript"/ },
  color: { js: "composeUi.applyColor('color', 'C00000')", re: /<w:color w:val="C00000"/i },
  colorMenu: { menu: true, re: /<w:color w:val="(?!000000|auto)[0-9A-F]{6}"/i },
  highlight: { js: "composeUi.applyColor('highlight', 'yellow')", re: /<w:highlight w:val="yellow"/ },
  size: { select: ["#fmtFontSize", "20"], re: /<w:sz w:val="40"/ },
  grow: { click: "#fmtGrow", re: /<w:sz w:val="(2[4-9]|3\d)"/ },
};
const MODES = {
  "→ do początku nast. akapitu": [0, 9, 1, 0, [[0, 9, 32]]],
  "← (od prawej)": [0, 32, 0, 9, [[0, 9, 32]]],
  "→ w akapicie": [0, 9, 0, 15, [[0, 9, 15]]],
  "→ przez 2 akapity": [0, 9, 1, 5, [[0, 9, 32], [1, 0, 5]]],
};
(async () => {
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = []; page.on("pageerror", (e) => errors.push(e.message));
  const rows = [];
  const fails = [];
  for (const [aname, act] of Object.entries(ACTIONS).filter(([k]) => !process.env.ONLY || process.env.ONLY.split(',').includes(k))) {
    const row = [aname];
    for (const [mname, m] of Object.entries(MODES)) {
      await page.goto(APP_URL, { waitUntil: "load" });
      await page.evaluate(() => document.getElementById("heroSplash")?.remove());
      await page.evaluate(() => composeUi.createNew("blank"));
      await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
      await page.waitForTimeout(300);
      await page.keyboard.type("Pierwszy akapit tekstu do testu.");
      await page.keyboard.press("Enter");
      await page.keyboard.type("Drugi akapit.");
      await page.keyboard.press("Enter");
      await page.keyboard.type("Trzeci.");
      await page.waitForTimeout(400);
      await page.evaluate((m) => {
        const list = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
        const pt = (pi, off) => { const r = formDomRange(list[pi], off, off); return [r.startContainer, r.startOffset]; };
        const [a, ao, b, bo] = m;
        getSelection().setBaseAndExtent(...pt(a, ao), ...pt(b, bo));
      }, m);
      if (act.menu) { await page.click("#fmtColorBtn"); await page.waitForSelector(".compose-pop-color .color-swatch"); await page.click(".compose-pop-color .color-swatch >> nth=3"); }
      else if (act.click) await page.click(act.click);
      else if (act.select) await page.selectOption(act.select[0], act.select[1]);
      else await page.evaluate(act.js);
      await page.waitForTimeout(900);
      const ok = await page.evaluate(async ({ re, want }) => {
        await waitInlineStructuralIdle();
        const z = await JSZip.loadAsync(await buildDocumentForSave());
        const xml = await z.file("word/document.xml").async("string");
        const ps = [...xml.matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g)].map((x) => x[1]);
        const rx = new RegExp(re);
        const flags = ps.map((p) => { const out = []; for (const r of p.matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)) { const has = rx.test((r[1].match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [""])[0]); for (const t of r[1].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)) for (const ch of t[1]) out.push(has); } return out; });
        let bad = 0;
        flags.forEach((f, pi) => f.forEach((has, ci) => { const inSel = want.some(([wp, a, b]) => wp === pi && ci >= a && ci < b); if (has !== inSel) bad++; }));
        return bad === 0 ? "OK" : `ŹLE(${bad})`;
      }, { re: act.re.source, want: m[4] });
      row.push(`${mname}: ${ok}`);
      if (ok !== "OK") fails.push(`${aname} / ${mname}: ${ok}`);
    }
    rows.push(row);
    console.log(row.join(" | "));
  }

  // komentarz do zaznaczenia od lewej do prawej (koniec na początku następnego akapitu)
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); localStorage.setItem("dwb.authorName", "Test"); });
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await page.waitForTimeout(300);
  await page.keyboard.type("Pierwszy akapit tekstu do testu.");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Drugi akapit.");
  await page.evaluate(() => {
    const list = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const pt = (pi, off) => { const r = formDomRange(list[pi], off, off); return [r.startContainer, r.startOffset]; };
    getSelection().setBaseAndExtent(...pt(0, 9), ...pt(1, 0));
  });
  await page.keyboard.press(`${ENGINE === "webkit" ? "Meta" : "Control"}+Alt+KeyM`);
  await page.waitForSelector(".compose-pop-comment .cf-rich", { timeout: 5000 }).catch(() => {});
  const formOpen = await page.evaluate(() => !!document.querySelector(".compose-pop-comment .cf-rich"));
  if (formOpen) {
    await page.keyboard.type("Uwaga");
    await page.click(".compose-pop-form .lf-ok");
    await page.waitForTimeout(800);
  }
  const commented = await page.evaluate(() => {
    const s = document.querySelector("span[data-cm-kind='start']");
    const e = document.querySelector("span[data-cm-kind='end']");
    if (!s || !e) return null;
    const r = document.createRange(); r.setStartAfter(s); r.setEndBefore(e);
    return r.toString();
  });
  console.log(`komentarz → od lewej do prawej: ${commented === "akapit tekstu do testu." ? "OK" : `ŹLE (${commented})`}`);
  if (commented !== "akapit tekstu do testu.") fails.push(`komentarz: ${commented}`);

  if (errors.length) fails.push(...errors);
  await browser.close();
  if (fails.length) { console.error("❌ format-direction-playwright failed:\n" + fails.join("\n")); process.exit(1); }
  console.log(`✅ format-direction-playwright passed [${ENGINE}]`);
})();
