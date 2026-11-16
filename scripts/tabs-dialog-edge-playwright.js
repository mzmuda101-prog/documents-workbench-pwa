// tabs-dialog-edge-playwright.js — okno „Tabulatory…” w nietypowych sytuacjach (przegląd błędów 2026-11-14):
// kilka akapitów naraz, pozycja z kropką / złe wartości, akapit listy numerowanej (numer zostaje),
// tabulator stylu usunięty i dodany z powrotem (plik bez zbędnych wpisów), komórka tabeli, tabulator
// dziesiętny (przecinek na pozycji), zapis → otwarcie, telefon (okno w ekranie).
//   node scripts/tabs-dialog-edge-playwright.js        ENGINE=webkit node scripts/tabs-dialog-edge-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 }).then(() => page.waitForTimeout(500));
const TAB = (a, b) => `<w:r><w:t xml:space="preserve">${a}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">${b}</w:t></w:r>`;
const TCELL = (inner) => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;

(async () => {
  const styles = `<w:style w:type="paragraph" w:customStyle="1" w:styleId="Formularz"><w:name w:val="Formularz"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="left" w:leader="dot" w:pos="5000"/></w:tabs></w:pPr></w:style>`;
  const body = [
    { raw: TAB("Jeden", "a") }, { raw: TAB("Dwa", "b") },
    { style: "Formularz", raw: TAB("Pole", "x") },
    { numId: 1, raw: TAB("Punkt", "opis") },
    { raw: TAB("Kwota", "12,50") }, { raw: TAB("Suma", "1.234,50") }, { raw: TAB("Ilość", "125") },
    `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr>${TCELL(`<w:p>${TAB("K", "v")}</w:p>`)}${TCELL("<w:p><w:r><w:t>B</w:t></w:r></w:p>")}</w:tr></w:tbl>`,
    "Koniec",
  ];
  const bytes = await buildDocx({ styles, body, abstractNums: [{ levels: [{ ilvl: 0, fmt: "decimal" }] }], nums: [{ id: 1, abs: 0 }] });
  const browser = await pw[ENGINE].launch({ headless: true });

  const open = async (ctxOpts = {}) => {
    const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 }, ...ctxOpts });
    await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
    const page = await ctx.newPage();
    page.errors = [];
    page.on("pageerror", (e) => page.errors.push(e.message));
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.locator("#fileInput").setInputFiles({ name: "tab-edge.docx", mimeType: DOCX_MIME, buffer: bytes });
    await idle(page);
    await page.evaluate(() => appFrame.setReadOnly(false));
    await idle(page);
    return page;
  };
  const P = (re) => `collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).findIndex((p) => ${re}.test(p.textContent))`;
  const select = (page, a, b = a) => page.evaluate(([a, b]) => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    ps[a].closest(".docx-edit-root").focus();
    const r = document.createRange();
    r.setStart(ps[a].querySelector("span")?.firstChild || ps[a], 1);
    if (a === b) r.collapse(true); else r.setEnd(ps[b].querySelector("span")?.firstChild || ps[b], 1);
    getSelection().removeAllRanges(); getSelection().addRange(r);
    document.dispatchEvent(new Event("selectionchange"));
  }, [a, b]);
  const idx = (page, re) => page.evaluate((src) => { const re = new RegExp(src); return collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).findIndex((p) => re.test(p.textContent)); }, re.source);
  const openTabs = async (page) => {
    await page.evaluate((s) => { if (typeof dwbRibbon !== "undefined" && dwbRibbon?.isOn()) { if (document.body.classList.contains("rb-reading")) document.querySelector('.mode-btn[data-mode="edit"]').click(); dwbRibbon.reveal(s); } }, "#fmtSpacingBtn").then(async () => page.click("#fmtSpacingBtn"));
    if (await page.evaluate(() => document.body.classList.contains("rb-on"))) {
      await page.evaluate(() => { dwbRibbon.select("layout"); document.querySelector('[data-rb-act="tabsForm"]').click(); });
    } else {
      await page.waitForSelector('.compose-pop [data-tabs="1"]');
      await page.click('.compose-pop [data-tabs="1"]');
    }
    await page.waitForSelector(".compose-pop-tabs");
  };
  const tabsXml = (page) => page.evaluate(async () => {
    const x = await getDocumentXmlString(await buildDocumentForSave());
    const doc = new DOMParser().parseFromString(x, "application/xml");
    return collectParagraphElements(doc.documentElement, "all").map((p) => ({ t: p.textContent, tabs: (new XMLSerializer().serializeToString(p).match(/<w:tabs>[\s\S]*?<\/w:tabs>/) || [""])[0].replace(/ xmlns:w="[^"]*"/g, ""), num: /<w:numPr>/.test(new XMLSerializer().serializeToString(p)) }));
  });

  const page = await open();
  // 1. dwa akapity naraz
  const a = await idx(page, /^Jeden/), b = await idx(page, /^Dwa/);
  await select(page, a, b);
  await openTabs(page);
  await page.fill(".tf-pos", "12.5"); // kropka zamiast przecinka
  await page.selectOption(".tf-align", "right");
  await page.click(".tf-ok");
  await idle(page);
  let t = await tabsXml(page);
  check("1: dwa zaznaczone akapity — oba dostają tabulator (pozycja „12.5” z kropką)", t.filter((x) => /^(Jeden|Dwa)/.test(x.t)).every((x) => /<w:tab w:val="right" w:pos="7087"\/>/.test(x.tabs)), JSON.stringify(t.slice(0, 2)));
  // 2. złe wartości
  await select(page, await idx(page, /^Koniec/));
  await openTabs(page);
  const bad = [];
  for (const v of ["0", "-1", "60", "abc", ""]) {
    await page.fill(".tf-pos", v);
    await page.click(".tf-set");
    bad.push(await page.evaluate(() => ({ warn: !document.querySelector(".compose-pop-tabs .mf-warn").hidden, n: document.querySelectorAll(".tabs-row").length })));
  }
  check("2: 0 / -1 / 60 / abc / puste — komunikat, nic nie ustawione", bad.every((x) => x.warn && x.n === 0), JSON.stringify(bad));
  await page.click(".tf-cancel");
  await idle(page);
  // 3. akapit listy numerowanej
  const li = await idx(page, /^Punkt/);
  const numBefore = await page.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].getAttribute("data-dwb-num"), li);
  await select(page, li);
  await openTabs(page);
  await page.fill(".tf-pos", "6");
  await page.click(".tf-ok");
  await idle(page);
  t = await tabsXml(page);
  const numAfter = await page.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].getAttribute("data-dwb-num"), li);
  check("3: akapit listy — tabulator dodany, numer i numPr zostają", /w:pos="3402"/.test(t.find((x) => /^Punkt/.test(x.t)).tabs) && t.find((x) => /^Punkt/.test(x.t)).num && numBefore && numAfter === numBefore, JSON.stringify({ numBefore, numAfter, row: t.find((x) => /^Punkt/.test(x.t)) }));
  // 4. tabulator stylu: usuń, potem dodaj z powrotem ten sam — plik bez własnych wpisów
  const fi = await idx(page, /^Pole/);
  await select(page, fi);
  await openTabs(page);
  await page.click(".tabs-del");
  await page.click(".tf-ok");
  await idle(page);
  const cleared = (await tabsXml(page)).find((x) => /^Pole/.test(x.t)).tabs;
  await select(page, fi);
  await openTabs(page);
  await page.fill(".tf-pos", "8,82");
  await page.selectOption(".tf-leader", "dot");
  await page.click(".tf-ok");
  await idle(page);
  const back = (await tabsXml(page)).find((x) => /^Pole/.test(x.t)).tabs;
  check("4: tabulator stylu usunięty („clear”), dodany z powrotem — akapit bez własnych tabulatorów", /w:val="clear"/.test(cleared) && back === "", JSON.stringify({ cleared, back }));
  // 5. komórka tabeli
  const ci = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).findIndex((p) => p.closest("td") && /^K/.test(p.textContent)));
  await select(page, ci);
  await openTabs(page);
  await page.fill(".tf-pos", "4");
  await page.selectOption(".tf-align", "center");
  await page.click(".tf-ok");
  await idle(page);
  t = await tabsXml(page);
  check("5: akapit w komórce tabeli — tabulator środkowy w pliku", /<w:tab w:val="center" w:pos="2268"\/>/.test(t.find((x) => /^K\s*v$/.test(x.t))?.tabs || ""), JSON.stringify(t.filter((x) => /^K/.test(x.t))));
  // 6. tabulator dziesiętny: przecinek „12,50” na pozycji 10 cm
  const ki = await idx(page, /^Kwota/);
  await select(page, ki);
  await openTabs(page);
  await page.fill(".tf-pos", "10");
  await page.selectOption(".tf-align", "decimal");
  await page.click(".tf-ok");
  await idle(page);
  const dec = await page.evaluate((i) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    const tn = [...p.querySelectorAll("span")].map((s) => s.firstChild).find((n) => n?.nodeType === 3 && n.data.includes("12,50"));
    const r = document.createRange(); const k = tn.data.indexOf(","); r.setStart(tn, k); r.setEnd(tn, k + 1);
    const scale = p.getBoundingClientRect().width / p.offsetWidth;
    return (r.getBoundingClientRect().left - p.getBoundingClientRect().left) / scale;
  }, ki);
  check("6: tabulator dziesiętny — przecinek „12,50” na 10 cm (±6 px)", Math.abs(dec - 10 * 96 / 2.54) < 6, `${Math.round(dec)} px, cel ${Math.round(10 * 96 / 2.54)}`);
  // jak Word: „1.234,50” — do pozycji kropka (pierwszy separator); „125” — koniec liczby
  const si = await idx(page, /^Suma/), ii = await idx(page, /^Ilo/);
  await select(page, si, ii);
  await openTabs(page);
  await page.fill(".tf-pos", "10");
  await page.selectOption(".tf-align", "decimal");
  await page.click(".tf-ok");
  await idle(page);
  const geo = await page.evaluate(([si, ii]) => {
    const pos = (i, needle, end) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
      const tn = [...p.querySelectorAll("span")].map((s) => s.firstChild).find((n) => n?.nodeType === 3 && n.data.includes(needle));
      const r = document.createRange(); const k = end ? tn.data.length : tn.data.indexOf(needle); r.setStart(tn, k); r.setEnd(tn, k);
      return (r.getBoundingClientRect().left - p.getBoundingClientRect().left) / (p.getBoundingClientRect().width / p.offsetWidth);
    };
    return { dot: pos(si, ".", false), end: pos(ii, "125", true) };
  }, [si, ii]);
  const goal = 10 * 96 / 2.54;
  check("6b: jak Word — „1.234,50” kropką na 10 cm, „125” końcem na 10 cm", Math.abs(geo.dot - goal) < 6 && Math.abs(geo.end - goal) < 6, JSON.stringify(geo));
  // 7. zapis → otwarcie: tabulatory zostają
  const before = await tabsXml(page);
  const b64 = await page.evaluate(async () => { const x = await buildDocumentForSave(); let s = ""; x.forEach((v) => { s += String.fromCharCode(v); }); return btoa(s); });
  await page.locator("#fileInput").setInputFiles({ name: "tab-edge-2.docx", mimeType: DOCX_MIME, buffer: Buffer.from(b64, "base64") });
  await idle(page);
  check("7: zapis → otwarcie: te same tabulatory", JSON.stringify(await tabsXml(page)) === JSON.stringify(before));
  check("bez błędów strony", !page.errors.length, page.errors.join(" | "));
  await page.context().close();

  // 8. telefon: okno w ekranie
  const ph = await open({ viewport: { width: 390, height: 844 }, hasTouch: true, ...(ENGINE === "chromium" ? { isMobile: true } : {}) });
  await select(ph, await idx(ph, /^Pole/));
  await ph.evaluate(() => document.getElementById("fmtSpacingBtn").click());
  await ph.waitForSelector('.compose-pop [data-tabs="1"]');
  await ph.click('.compose-pop [data-tabs="1"]');
  await ph.waitForSelector(".compose-pop-tabs");
  const fit = await ph.evaluate(() => { const r = document.querySelector(".compose-pop-tabs").getBoundingClientRect(); return { l: r.left, r: r.right, t: r.top, b: r.bottom, w: innerWidth, h: innerHeight }; });
  check("8: telefon — okno „Tabulatory” w ekranie", fit.l >= 0 && fit.r <= fit.w + 0.5 && fit.t >= 0 && fit.b <= fit.h + 0.5, JSON.stringify(fit));
  check("bez błędów strony (telefon)", !ph.errors.length, ph.errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
