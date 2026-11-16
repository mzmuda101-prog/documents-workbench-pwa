// header-sections-playwright.js — nagłówki i stopki w sekcjach jak w Wordzie (docx-compose.js
// applyHeaderFooterInZip): formularz zna sekcję kursora, „Połącz z poprzednią” (rozłącz → własny
// nagłówek sekcji, zmiana w sekcji 1 nie rusza 2; połącz → sekcja dziedziczy), „Inne strony
// parzyste i nieparzyste” (settings.xml + części „even”, wyłączenie je usuwa), Cofnij, zapis →
// otwarcie, schemat Office.
//   node scripts/header-sections-playwright.js        ENGINE=webkit node scripts/header-sections-playwright.js
const pw = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");
const ooxml = require("./ooxml-validate");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 }).then(() => page.waitForTimeout(600));

(async () => {
  const bytes = await buildDocx({
    body: ["Pierwsza sekcja, akapit jeden.", "Druga sekcja zaczyna się tutaj.", "Ostatni akapit."],
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
  await load(bytes, "naglowki.docx");

  const caret = (re, at) => page.evaluate(([src, at]) => {
    const re = new RegExp(src);
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((q) => re.test(q.textContent));
    p.closest(".docx-edit-root").focus(); placeCaret(p, at); document.dispatchEvent(new Event("selectionchange"));
  }, [re.source, at]);
  // nagłówek każdej sekcji (część, do której prowadzi własne albo odziedziczone odwołanie) + ustawienia
  const state = () => page.evaluate(async () => {
    const zip = await JSZip.loadAsync(await buildDocumentForSave());
    const d = new DOMParser().parseFromString(await zip.file("word/document.xml").async("string"), "application/xml");
    const rels = await zip.file("word/_rels/document.xml.rels").async("string");
    const target = (rid) => (rels.match(new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${rid}"`)) || []).slice(1).find(Boolean);
    const text = async (rid) => { const f = rid && zip.file("word/" + target(rid)); return f ? (await f.async("string")).replace(/<[^>]+>/g, "").trim() : ""; };
    const sects = [...d.getElementsByTagNameNS(W_NS, "sectPr")];
    const out = [];
    for (let i = 0; i < sects.length; i++) {
      const refs = (kind, type) => [...sects[i].childNodes].find((n) => n.localName === `${kind}Reference` && (n.getAttributeNS(W_NS, "type") || "default") === type);
      let seen = null;
      for (let k = i; k >= 0 && !seen; k--) seen = [...sects[k].childNodes].find((n) => n.localName === "headerReference" && (n.getAttributeNS(W_NS, "type") || "default") === "default");
      const evenRef = refs("header", "even");
      out.push({ own: !!refs("header", "default"), header: await text(seen && (seen.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"))), even: evenRef ? await text(evenRef.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")) : null, evenFooter: !!refs("footer", "even") });
    }
    const settings = await zip.file("word/settings.xml")?.async("string") || "";
    return { sects: out, evenOdd: /<w:evenAndOddHeaders\b/.test(settings), parts: Object.keys(zip.files).filter((n) => /^word\/(header|footer)\d+\.xml$/.test(n)).length };
  });
  const view = () => page.evaluate(() => [...document.querySelectorAll(".docx-preview-host section.docx")].map((s) => (s.querySelector("header:not(.dwb-hf-alt)")?.textContent || "").trim()));
  // formularz nagłówka: odczyt i wypełnienie
  const form = async (fill) => {
    await page.evaluate(() => composeUi.openHeaderFooterForm());
    await page.waitForSelector(".compose-pop-hf .lf-ok");
    const seen = await page.evaluate(() => { const el = document.querySelector(".compose-pop-hf"); return { cap: el.querySelector(".compose-cap").textContent, linkShown: !el.querySelector(".hf-linkrow").hidden, linked: el.querySelector(".hf-link").checked, header: el.querySelector(".hf-h").value, eo: el.querySelector(".hf-eo").checked }; });
    if (fill) {
      await page.evaluate((f) => {
        const el = document.querySelector(".compose-pop-hf");
        const set = (sel, v) => { const x = el.querySelector(sel); if (x.type === "checkbox") { x.checked = v; x.dispatchEvent(new Event("change")); } else x.value = v; };
        for (const [sel, v] of Object.entries(f)) set(sel, v);
        el.querySelector(".lf-ok").click();
      }, fill);
      await idle(page);
    } else await page.evaluate(() => document.querySelector(".compose-pop-hf .lf-cancel").click());
    return seen;
  };

  // podział sekcji po „Pierwsza sekcja…” (Układ → Podziały → następna strona)
  await caret(/^Pierwsza/, 29);
  await page.evaluate(() => typeof dwbRibbon !== "undefined" && dwbRibbon?.isOn() && dwbRibbon.select("layout")).then(() => page.evaluate(() => document.getElementById("pageLayoutBtn").click()));
  await page.waitForSelector('.compose-pop [data-brk="nextPage"]');
  await page.click('.compose-pop [data-brk="nextPage"]');
  await idle(page);
  let s = await state();
  check("0: dwie sekcje, druga dziedziczy nagłówek", s.sects.length === 2 && s.sects[0].own && !s.sects[1].own && s.sects[1].header === "NAGŁÓWEK FIRMY", JSON.stringify(s));

  // 1. formularz w sekcji 2: „sekcja 2 z 2”, połączona; rozłącz i wpisz własny nagłówek
  await caret(/^Druga/, 3);
  const f1 = await form({ ".hf-link": false, ".hf-h": "ROZDZIAŁ 2" });
  s = await state();
  check("1: formularz zna sekcję kursora („sekcja 2 z 2”) i połączenie", /sekcja 2 z 2/.test(f1.cap) && f1.linkShown && f1.linked && f1.header === "NAGŁÓWEK FIRMY", JSON.stringify(f1));
  check("1b: rozłączona sekcja 2 ma własny nagłówek, sekcja 1 bez zmian", s.sects[0].header === "NAGŁÓWEK FIRMY" && s.sects[1].own && s.sects[1].header === "ROZDZIAŁ 2", JSON.stringify(s));
  check("1c: podgląd — każda sekcja swój nagłówek", JSON.stringify(await view()) === JSON.stringify(["NAGŁÓWEK FIRMY", "ROZDZIAŁ 2"]), JSON.stringify(await view()));

  // 2. zmiana w sekcji 1 nie rusza sekcji 2; w sekcji 1 nie ma „Połącz z poprzednią”
  await caret(/^Pierwsza/, 3);
  const f2 = await form({ ".hf-h": "FIRMA NOWA" });
  s = await state();
  check("2: sekcja 1 — bez „Połącz z poprzednią”", !f2.linkShown && /sekcja 1 z 2/.test(f2.cap), JSON.stringify(f2));
  check("2b: zmiana w sekcji 1 nie rusza rozłączonej sekcji 2", s.sects[0].header === "FIRMA NOWA" && s.sects[1].header === "ROZDZIAŁ 2", JSON.stringify(s));

  // 3. połącz sekcję 2 z powrotem — dziedziczy, jej część znika z paczki
  const partsBefore = s.parts;
  await caret(/^Druga/, 3);
  await form({ ".hf-link": true });
  s = await state();
  check("3: połączona sekcja 2 — bez własnych odwołań, nagłówek z sekcji 1, część usunięta", !s.sects[1].own && s.sects[1].header === "FIRMA NOWA" && s.parts === partsBefore - 1, JSON.stringify({ s, partsBefore }));

  // 4. inne strony parzyste: ustawienie dokumentu + części „even” (z numerem w stopce)
  await caret(/^Pierwsza/, 3);
  await form({ ".hf-eo": true, ".hf-eh": "PARZYSTE", ".hf-n": "n" });
  s = await state();
  check("4: parzyste — w:evenAndOddHeaders w settings.xml, nagłówek i stopka „even” w sekcji 1", s.evenOdd && s.sects[0].even === "PARZYSTE" && s.sects[0].evenFooter, JSON.stringify(s));
  const f4 = await form(null);
  check("4b: formularz pokazuje włączone parzyste", f4.eo, JSON.stringify(f4));

  // 5. Cofnij — bez stron parzystych
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  s = await state();
  check("5: Cofnij — bez stron parzystych", !s.evenOdd && s.sects.every((x) => x.even == null), JSON.stringify(s));

  // 6. ponów, potem wyłącz w formularzu — części „even” znikają ze wszystkich sekcji
  await page.evaluate(() => document.getElementById("redoBtn").click());
  await idle(page);
  await form({ ".hf-eo": false });
  s = await state();
  check("6: wyłączone parzyste — bez ustawienia i bez części „even”", !s.evenOdd && s.sects.every((x) => x.even == null && !x.evenFooter), JSON.stringify(s));

  // 7. zapis → otwarcie — ten sam stan; schemat Office
  const before = await state();
  const b64 = await page.evaluate(async () => { const x = await buildDocumentForSave(); let s = ""; x.forEach((v) => { s += String.fromCharCode(v); }); return btoa(s); });
  const saved = Buffer.from(b64, "base64");
  await load(saved, "naglowki-2.docx");
  check("7: zapis → otwarcie — te same nagłówki sekcji", JSON.stringify(await state()) === JSON.stringify(before), JSON.stringify(await state()));
  if (ooxml.available()) {
    const tmp = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "dwb-hf-")), "naglowki.docx");
    fs.writeFileSync(tmp, saved);
    const [v] = ooxml.validate([tmp]);
    check("schemat Office", v.ok, v.error || v.errors?.slice(0, 3).map(ooxml.short).join(" | "));
  }
  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
