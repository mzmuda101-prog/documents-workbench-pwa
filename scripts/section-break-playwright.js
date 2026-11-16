// section-break-playwright.js — Układ → Podziały → Podział sekcji jak w Wordzie (docx-compose.js
// applySectionBreakInXml; tools/word-oracle/probe-section-break.js): następna strona w środku akapitu
// (tekst za kursorem w nowej sekcji, nagłówek przeniesiony do pierwszej, druga dziedziczy), orientacja
// pozioma tylko drugiej sekcji, podział ciągły, w tabeli — odmowa, Cofnij, zapis → otwarcie.
//   node scripts/section-break-playwright.js        ENGINE=webkit node scripts/section-break-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 }).then(() => page.waitForTimeout(600));

(async () => {
  const cell = (t) => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
  const bytes = await buildDocx({
    body: ["Pierwszy akapit dokumentu.", "Drugi akapit do podziału w środku zdania.", "Trzeci akapit.", `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr>${cell("Komórka")}${cell("B")}</w:tr></w:tbl>`, "Ostatni."],
    sectPr: '<w:headerReference w:type="default" r:id="rIdX0"/>',
    extraParts: { "header1.xml": { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", rel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p><w:r><w:t>NAGŁÓWEK FIRMY</w:t></w:r></w:p></w:hdr>` } },
  });
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const load = async (b, name) => { await page.locator("#fileInput").setInputFiles({ name, mimeType: DOCX_MIME, buffer: b }); await idle(page); await page.evaluate(() => appFrame.setReadOnly(false)); await idle(page); };
  await load(bytes, "sekcje.docx");

  const caret = (re, at) => page.evaluate(([src, at]) => {
    const re = new RegExp(src);
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((q) => re.test(q.textContent));
    p.closest(".docx-edit-root").focus(); placeCaret(p, at === "end" ? p.textContent.length : at); document.dispatchEvent(new Event("selectionchange"));
  }, [re.source, at]);
  const pick = async (key) => {
    await page.evaluate(() => typeof dwbRibbon !== "undefined" && dwbRibbon?.isOn() && dwbRibbon.select("layout")).then(() => page.evaluate(() => document.getElementById("pageLayoutBtn").click()));
    await page.waitForSelector(`.compose-pop [data-brk="${key}"]`);
    await page.click(`.compose-pop [data-brk="${key}"]`);
    await idle(page);
  };
  const sects = () => page.evaluate(async () => {
    const x = await getDocumentXmlString(await buildDocumentForSave());
    const d = new DOMParser().parseFromString(x, "application/xml");
    return [...d.getElementsByTagNameNS(W_NS, "sectPr")].map((s) => {
      const p = s.parentNode.localName === "pPr" ? s.parentNode.parentNode.textContent : "(koniec)";
      const kids = [...s.childNodes].filter((n) => n.nodeType === 1);
      return { after: p.slice(0, 30), hdr: kids.some((n) => n.localName === "headerReference"), type: kids.find((n) => n.localName === "type")?.getAttributeNS(W_NS, "val") || "", orient: kids.find((n) => n.localName === "pgSz")?.getAttributeNS(W_NS, "orient") || "" };
    });
  });
  const view = () => page.evaluate(() => [...document.querySelectorAll(".docx-preview-host section.docx")].map((s) => ({ w: Math.round(s.offsetWidth), h: Math.round(s.offsetHeight), first: s.querySelector("article p")?.textContent.slice(0, 20) || "", hdr: (s.querySelector("header")?.textContent || "").trim() })));

  // 1. następna strona w środku akapitu
  await caret(/^Drugi/, 6);
  await pick("nextPage");
  let s = await sects(), v = await view();
  check("1: plik — „Drugi ” kończy sekcję 1 (z nagłówkiem), sekcja 2 bez odwołań (dziedziczy), bez w:type", s.length === 2 && s[0].after === "Drugi " && s[0].hdr && !s[1].hdr && s[1].type === "", JSON.stringify(s));
  check("1b: podgląd — dwie sekcje, druga zaczyna się od „akapit do podziału”", v.length >= 2 && /^akapit do podzia/.test(v[v.length - 1].first || v[1].first), JSON.stringify(v));
  check("1c: nagłówek na obu stronach (druga dziedziczy)", v.every((x) => x.hdr === "NAGŁÓWEK FIRMY"), JSON.stringify(v.map((x) => x.hdr)));
  // 2. orientacja pozioma tylko drugiej sekcji
  await caret(/^akapit do podzia/, 2);
  await page.evaluate(() => typeof dwbRibbon !== "undefined" && dwbRibbon?.isOn() && dwbRibbon.select("layout")).then(() => page.evaluate(() => document.getElementById("pageLayoutBtn").click()));
  await page.waitForSelector('.compose-pop [data-orient="landscape"]');
  await page.click('.compose-pop [data-orient="landscape"]');
  await idle(page);
  s = await sects(); v = await view();
  check("2: sekcja 2 pozioma, sekcja 1 pionowa (plik i podgląd)", s[0].orient === "" && s[1].orient === "landscape" && v[0].w < v[0].h && v[v.length - 1].w > v[v.length - 1].h, JSON.stringify({ s, v }));
  // 3. ciągły na końcu pierwszego akapitu
  await caret(/^Pierwszy/, "end");
  await pick("continuous");
  s = await sects();
  check("3: podział ciągły — nowa sekcja po „Pierwszy…” z w:type continuous", s.length === 3 && /^Pierwszy/.test(s[0].after) && s[1].type === "continuous" && s[0].hdr && !s[1].hdr, JSON.stringify(s));
  // 4. Cofnij
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  check("4: Cofnij — znów dwie sekcje", (await sects()).length === 2, JSON.stringify(await sects()));
  // 5. w tabeli — odmowa
  await page.evaluate(() => { const p = document.querySelector(".docx-preview-host td p"); p.closest(".docx-edit-root").focus(); placeCaret(p, 2); document.dispatchEvent(new Event("selectionchange")); });
  await pick("nextPage");
  check("5: w tabeli — bez podziału (komunikat)", (await sects()).length === 2, JSON.stringify(await sects()));
  // 6. zapis → otwarcie
  const before = await sects();
  const b64 = await page.evaluate(async () => { const x = await buildDocumentForSave(); let s = ""; x.forEach((v) => { s += String.fromCharCode(v); }); return btoa(s); });
  await load(Buffer.from(b64, "base64"), "sekcje-2.docx");
  check("6: zapis → otwarcie — te same sekcje", JSON.stringify(await sects()) === JSON.stringify(before), JSON.stringify(await sects()));
  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
