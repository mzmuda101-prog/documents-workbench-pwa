// tabs-dialog-playwright.js — okno „Tabulatory…” jak w Wordzie (compose-ui.js buildTabsForm, docx-compose.js
// applyTabsSetInZip): nowy tabulator prawy z kropkami (plik + położenie tekstu), tabulator ze stylu
// wyłączony przez „clear” + nowy środkowy, Cofnij, Anuluj, zła pozycja.
//   node scripts/tabs-dialog-playwright.js        ENGINE=webkit node scripts/tabs-dialog-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(500));
const TAB = (t) => `<w:r><w:t>${t[0]}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${t[1]}</w:t></w:r>`;

(async () => {
  const styles = `<w:style w:type="paragraph" w:customStyle="1" w:styleId="Formularz"><w:name w:val="Formularz"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="left" w:leader="dot" w:pos="5000"/></w:tabs></w:pPr></w:style>`;
  const bytes = await buildDocx({ styles, body: [{ raw: TAB(["Imię", "Nazwisko"]) }, { style: "Formularz", raw: TAB(["Pole", "wartość"]) }, "Trzeci"] });
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "tabs.docx", mimeType: DOCX_MIME, buffer: bytes });
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);

  const caret = (i) => page.evaluate((i) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; p.closest(".docx-edit-root").focus(); placeCaret(p, 1); document.dispatchEvent(new Event("selectionchange")); }, i);
  const openTabs = async (i) => {
    await caret(i);
    await page.click("#fmtSpacingBtn");
    await page.waitForSelector('.compose-pop [data-tabs="1"]');
    const desc = await page.locator('.compose-pop [data-tabs="1"]').textContent();
    await page.click('.compose-pop [data-tabs="1"]');
    await page.waitForSelector(".compose-pop-tabs");
    return desc;
  };
  const para = (i) => page.evaluate(async (i) => {
    await mergeInlineEditsIntoBytes();
    const d = await getDocumentXmlString(originalFileBytes);
    return ((d.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [])[i] || "").match(/<w:tabs>[\s\S]*?<\/w:tabs>/)?.[0] || "";
  }, i);

  // 1. nowy tabulator: 12,5 cm, prawy, kropki
  const d1 = await openTabs(0);
  check("menu odstępów: „Tabulatory…” — brak w akapicie", /Tabulatory/.test(d1) && /brak/.test(d1), d1);
  await page.fill(".tf-pos", "12,5");
  await page.selectOption(".tf-align", "right");
  await page.selectOption(".tf-leader", "dot");
  await page.click(".tf-set");
  check("„Ustaw”: tabulator na liście", /12,5 cm · Prawy/.test(await page.locator(".tabs-list").textContent()), await page.locator(".tabs-list").textContent());
  await page.click(".tf-ok");
  await idle(page);
  const x1 = await para(0);
  check("plik: w:tab right, kropki, 12,5 cm (7087 tw)", /<w:tab w:val="right" w:leader="dot" w:pos="7087"\/>/.test(x1), x1);
  const geo = await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0];
    const r = document.createRange(); const t = [...p.querySelectorAll("span")].map((s) => s.firstChild).find((n) => n?.nodeType === 3 && n.data.includes("Nazwisko"));
    r.selectNodeContents(t); const scale = p.getBoundingClientRect().width / p.offsetWidth;
    return { end: (r.getBoundingClientRect().right - p.getBoundingClientRect().left) / scale, dots: !!p.querySelector(".docx-tab.tab-dot") };
  });
  check("podgląd: „Nazwisko” kończy się na 12,5 cm, kropki przed nim", Math.abs(geo.end - 12.5 * 96 / 2.54) < 8 && geo.dots, JSON.stringify(geo));

  // 2. akapit w stylu z tabulatorem 5000 (8,82 cm): usuń go, dodaj środkowy 3 cm
  await openTabs(1);
  const listed = await page.locator(".tabs-list").textContent();
  check("tabulator ze stylu na liście (8,82 cm · Lewy · kropki)", /8,82 cm · Lewy · …/.test(listed), listed);
  await page.click(".tabs-del");
  await page.fill(".tf-pos", "3");
  await page.selectOption(".tf-align", "center");
  await page.click(".tf-ok"); // OK ustawia wpisaną pozycję (jak w Wordzie)
  await idle(page);
  const x2 = await para(1);
  check("plik: „clear” dla tabulatora stylu + nowy środkowy 3 cm", /<w:tab w:val="center" w:pos="1701"\/>/.test(x2) && /<w:tab w:val="clear" w:pos="5000"\/>/.test(x2), x2);
  // 3. Cofnij
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  check("Cofnij: akapit w stylu bez własnych tabulatorów", (await para(1)) === "", await para(1));
  // 4. Anuluj i zła pozycja
  await openTabs(2);
  await page.fill(".tf-pos", "abc");
  await page.click(".tf-set");
  const warn = await page.locator(".compose-pop-tabs .mf-warn").isVisible();
  await page.click(".tf-cancel");
  await idle(page);
  check("zła pozycja: komunikat; Anuluj: bez zmian w pliku", warn && (await para(2)) === "", String(warn));
  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
