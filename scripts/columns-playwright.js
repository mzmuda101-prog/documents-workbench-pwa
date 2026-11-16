// columns-playwright.js — Układ → Kolumny jak w Wordzie (docx-compose.js applyColumnsInXml) i kolumny
// strona po stronie (page-breaks.js columnPages): cały dokument w 2 kolumnach (kolumna 1, kolumna 2,
// potem następna strona; ostatnia strona — najpierw kolumna 1), linia między kolumnami, zaznaczony
// fragment (podziały ciągłe przed i za nim), „Jedna”, Cofnij, zapis → otwarcie, schemat Office.
//   node scripts/columns-playwright.js        ENGINE=webkit node scripts/columns-playwright.js
const pw = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");
const ooxml = require("./ooxml-validate");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 }).then(() => page.waitForTimeout(900));

(async () => {
  const body = Array.from({ length: 70 }, (_, i) => `Akapit ${i + 1}: tekst biuletynu w kolumnach, dość długi, żeby zawinął się na kilka wierszy w wąskiej kolumnie.`);
  const bytes = await buildDocx({ body });
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const load = async (b, name) => { await page.locator("#fileInput").setInputFiles({ name, mimeType: DOCX_MIME, buffer: b }); await idle(page); await page.evaluate(() => appFrame.setReadOnly(false)); await idle(page); };
  await load(bytes, "kolumny.docx");

  const paraEl = (n) => `collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((q) => q.textContent.startsWith("Akapit ${n}:"))`;
  const caret = (n) => page.evaluate((src) => { const p = eval(src); p.closest(".docx-edit-root").focus(); placeCaret(p, 3); document.dispatchEvent(new Event("selectionchange")); }, paraEl(n));
  const select = (a, b) => page.evaluate(([sa, sb]) => {
    const p = eval(sa), q = eval(sb);
    p.closest(".docx-edit-root").focus();
    const r = document.createRange();
    r.setStart(p.firstChild?.firstChild || p, 0);
    r.setEnd(q.lastChild?.lastChild || q, (q.lastChild?.lastChild || q).textContent.length);
    const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
    document.dispatchEvent(new Event("selectionchange"));
  }, [paraEl(a), paraEl(b)]);
  const pick = async (val) => {
    await page.evaluate(() => typeof dwbRibbon !== "undefined" && dwbRibbon?.isOn() && dwbRibbon.select("layout")).then(() => page.evaluate(() => document.getElementById("pageLayoutBtn").click()));
    await page.waitForSelector(`.compose-pop [data-cols="${val}"]`);
    await page.click(`.compose-pop [data-cols="${val}"]`);
    await idle(page);
  };
  const sects = () => page.evaluate(async () => {
    const d = new DOMParser().parseFromString(await getDocumentXmlString(await buildDocumentForSave()), "application/xml");
    return [...d.getElementsByTagNameNS(W_NS, "sectPr")].map((s) => {
      const c = [...s.childNodes].find((n) => n.localName === "cols");
      return { num: c?.getAttributeNS(W_NS, "num") || "", sep: c?.getAttributeNS(W_NS, "sep") || "", type: [...s.childNodes].find((n) => n.localName === "type")?.getAttributeNS(W_NS, "val") || "", last: (s.parentNode.localName === "pPr" ? s.parentNode.parentNode.textContent : "(koniec)").slice(0, 10) };
    });
  });
  // strony kolumn w podglądzie: pierwszy akapit każdej strony (za znacznikiem .dwb-col-break)
  const view = () => page.evaluate(() => {
    const host = document.querySelector(".docx-preview-host");
    const firstAfter = (b) => { let x = b.nextElementSibling; while (x && !x.matches("p")) x = x.nextElementSibling; return x?.textContent.split(":")[0]; };
    const art = host.querySelector("section.docx > article");
    const rect = (n) => { const p = [...host.querySelectorAll("p")].find((q) => q.textContent.startsWith(`Akapit ${n}:`)); const r = p?.getClientRects()[0]; return r && { left: r.left, top: r.top }; };
    // akapit na górze drugiej kolumny strony 1 (ta sama wysokość co „Akapit 1”, dalej w prawo)
    const a1 = rect(1);
    const beside = a1 && [...host.querySelectorAll("p")].find((q) => { const r = q.getClientRects()[0]; return r && r.left > a1.left + 100 && Math.abs(r.top - a1.top) < 3; });
    return { starts: [...host.querySelectorAll(".dwb-col-break")].map(firstAfter), rule: getComputedStyle(art).columnRuleStyle, a1, a2: rect(2), beside: beside ? parseInt(beside.textContent.split(":")[0].split(" ")[1], 10) : null };
  });

  // 1. cały dokument w 2 kolumnach — strony kolumn po kolei
  await caret(5);
  await pick("2");
  let s = await sects(), v = await view();
  check("1: plik — sekcja w 2 kolumnach (w:cols w:num=2), bez podziałów", s.length === 1 && s[0].num === "2", JSON.stringify(s));
  check("1b: podgląd — kolumny strona po stronie (nowe strony od dalszych akapitów, rosnąco)", v.starts.length >= 2 && v.starts.every((x, i) => i === 0 || parseInt(x.split(" ")[1], 10) > parseInt(v.starts[i - 1].split(" ")[1], 10)) && parseInt(v.starts[0].split(" ")[1], 10) > 15, JSON.stringify(v.starts));
  check("1c: strona 1 — Akapit 2 pod Akapitem 1, a obok niego (góra kolumny 2) Akapit ok. 14, nie 36 (jak Word)", v.a1 && v.a2 && Math.abs(v.a1.left - v.a2.left) < 2 && v.a2.top > v.a1.top && v.beside >= 12 && v.beside <= 16, JSON.stringify({ a1: v.a1, a2: v.a2, beside: v.beside }));

  // 2. linia między kolumnami
  await caret(5);
  await pick("sep");
  s = await sects(); v = await view();
  check("2: linia między kolumnami — w:sep, w podglądzie column-rule", s[0].sep === "1" && v.rule !== "none", JSON.stringify({ s, rule: v.rule }));

  // 3. „Jedna” — bez kolumn
  await caret(5);
  await pick("1");
  s = await sects(); v = await view();
  check("3: „Jedna” — bez w:num i w:sep, znaczniki kolumn zniknęły", s[0].num === "" && s[0].sep === "" && !v.starts.length, JSON.stringify({ s, starts: v.starts }));

  // 4. zaznaczony fragment (Akapity 10–20) — podziały ciągłe przed i za nim, kolumny tylko w nim
  await select(10, 20);
  await pick("2");
  s = await sects();
  check("4: fragment — 3 sekcje: do Akapitu 9 (1 kol.), Akapity 10–20 (2 kol., ciągła), reszta (1 kol., ciągła)", s.length === 3 && s[0].last.startsWith("Akapit 9:") && s[0].num === "" && s[1].last.startsWith("Akapit 20") && s[1].num === "2" && s[1].type === "continuous" && s[2].num === "" && s[2].type === "continuous", JSON.stringify(s));

  // 5. Cofnij — znów jedna sekcja bez kolumn
  await page.evaluate(() => document.getElementById("undoBtn").click());
  await idle(page);
  s = await sects();
  check("5: Cofnij — jedna sekcja, bez kolumn", s.length === 1 && s[0].num === "", JSON.stringify(s));

  // 6. zapis → otwarcie (dokument w 2 kolumnach) — te same sekcje; schemat Office
  await caret(5);
  await pick("2");
  const before = await sects();
  const b64 = await page.evaluate(async () => { const x = await buildDocumentForSave(); let s = ""; x.forEach((v) => { s += String.fromCharCode(v); }); return btoa(s); });
  const saved = Buffer.from(b64, "base64");
  await load(saved, "kolumny-2.docx");
  check("6: zapis → otwarcie — te same kolumny", JSON.stringify(await sects()) === JSON.stringify(before), JSON.stringify(await sects()));
  if (ooxml.available()) {
    const tmp = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "dwb-cols-")), "kolumny.docx");
    fs.writeFileSync(tmp, saved);
    const [r] = ooxml.validate([tmp]);
    check("schemat Office", r.ok, r.error || r.errors?.slice(0, 3).map(ooxml.short).join(" | "));
  }
  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
