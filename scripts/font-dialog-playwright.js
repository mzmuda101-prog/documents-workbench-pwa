// font-dialog-playwright.js — okno „Czcionka…” jak w Wordzie (compose-ui.js buildFontForm) i cechy tekstu
// w modelu fragmentów (docx-run-styles.js): wersaliki, kapitaliki, odstęp znaków, położenie, styl
// podkreślenia. Niezmienniki: plik z tymi cechami po otwarciu = bez zmian (zero „zmienionych”
// akapitów); pisanie w akapicie — każda cecha zostaje na SWOIM fragmencie; okno: zastosowanie na
// zaznaczeniu (plik + podgląd), Ctrl/⌘+D, odczyt bieżących cech, Cofnij, dalsze pisanie bez zaznaczenia.
//   node scripts/font-dialog-playwright.js        ENGINE=webkit node scripts/font-dialog-playwright.js
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 }).then(() => page.waitForTimeout(500));
const R = (rpr, t) => `<w:r><w:rPr>${rpr}</w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
const PLAIN = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;

(async () => {
  const body = [
    { raw: PLAIN("Zwykły ") + R("<w:caps/>", "wersaliki") + PLAIN(" i ") + R("<w:smallCaps/>", "kapitaliki") + PLAIN(" koniec.") },
    { raw: PLAIN("Odstęp ") + R('<w:spacing w:val="40"/>', "rozstrzelony") + PLAIN(" i ") + R('<w:spacing w:val="-10"/>', "zagęszczony") + PLAIN(".") },
    { raw: PLAIN("Położenie ") + R('<w:position w:val="6"/>', "wyżej") + PLAIN(" i ") + R('<w:position w:val="-4"/>', "niżej") + PLAIN(".") },
    { raw: PLAIN("Podkreślenie ") + R('<w:u w:val="double"/>', "podwójne") + PLAIN(" ") + R('<w:u w:val="dotted"/>', "kropki") + PLAIN(" ") + R('<w:u w:val="wave"/>', "fala") + PLAIN(" ") + R('<w:u w:val="single"/>', "zwykłe") + PLAIN(".") },
    "Akapit do okna czcionki z kilkoma słowami.",
  ];
  const bytes = await buildDocx({ body });
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "czcionka.docx", mimeType: DOCX_MIME, buffer: bytes });
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);

  // rPr fragmentów akapitu w pliku (po scaleniu pisania) — tekst → cechy
  const fileRuns = (page, start) => page.evaluate(async (start) => {
    const x = await getDocumentXmlString(await buildDocumentForSave());
    const doc = new DOMParser().parseFromString(x, "application/xml");
    const p = collectParagraphElements(doc.documentElement, "all").find((q) => q.textContent.replace("X", "").startsWith(start));
    if (!p) return ["BRAK: " + collectParagraphElements(doc.documentElement, "all").map((q) => q.textContent.slice(0, 15)).join(" / ")];
    return [...p.getElementsByTagNameNS(W_NS, "r")].map((r) => {
      const rPr = r.getElementsByTagNameNS(W_NS, "rPr")[0];
      const kids = rPr ? [...rPr.childNodes].filter((n) => n.nodeType === 1).map((n) => `${n.localName}${n.getAttributeNS(W_NS, "val") ? `=${n.getAttributeNS(W_NS, "val")}` : ""}`).join(",") : "";
      return `${r.textContent}|${kids}`;
    }).filter((s) => !s.startsWith("|"));
  }, start);
  const caretAt = (start, at) => page.evaluate(([start, at]) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((q) => q.textContent.startsWith(start));
    p.closest(".docx-edit-root").focus(); placeCaret(p, at);
  }, [start, at]);

  check("1: otwarty plik z wersalikami, kapitalikami, odstępem, położeniem, podkreśleniami — bez „zmienionych” akapitów", (await page.evaluate(() => collectInlineParagraphEdits().length)) === 0, JSON.stringify(await page.evaluate(() => collectInlineParagraphEdits().map((e) => e.index))));
  const look = await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const span = (t) => [...document.querySelectorAll(".docx-preview-host span")].find((s) => s.childNodes.length === 1 && s.textContent === t);
    const cs = (t) => getComputedStyle(span(t));
    return { caps: cs("wersaliki").textTransform, small: cs("kapitaliki").fontVariantCaps || cs("kapitaliki").fontVariant, sp: cs("rozstrzelony").letterSpacing, dbl: cs("podwójne").textDecorationStyle };
  });
  check("1b: podgląd rysuje cechy (wersaliki, kapitaliki, odstęp 2 pt, podwójne podkreślenie)", look.caps === "uppercase" && /small-caps/.test(look.small) && Math.abs(parseFloat(look.sp) - 2.667) < 0.1 && look.dbl === "double", JSON.stringify(look));
  // pisanie w każdym akapicie — cechy zostają na swoich fragmentach
  const before = {};
  for (const s of ["Zwykły", "Odstęp", "Położenie", "Podkreślenie"]) before[s] = await fileRuns(page, s);
  for (const s of ["Zwykły", "Odstęp", "Położenie", "Podkreślenie"]) { await caretAt(s, 2); await page.keyboard.type("X"); }
  await idle(page);
  for (const s of ["Zwykły", "Odstęp", "Położenie", "Podkreślenie"]) {
    const after = await fileRuns(page, s);
    const strip = (arr) => arr.map((x) => x.replace("X", ""));
    check(`2: pisanie w akapicie „${s}…” — cechy fragmentów bez zmian`, JSON.stringify(strip(after).slice(1)) === JSON.stringify(strip(before[s]).slice(1)) && after.join("").includes("X"), JSON.stringify({ before: before[s], after }));
  }

  // okno „Czcionka…”: zaznaczenie „kilkoma” → kapitaliki, odstęp rozstrzelony 1,5 pt, podniesione 3 pt, podkreślenie podwójne
  const sel = await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((q) => q.textContent.startsWith("Akapit do okna"));
    const tn = [...p.querySelectorAll("span")].map((s) => s.firstChild).find((n) => n?.nodeType === 3 && n.data.includes("kilkoma"));
    const k = tn.data.indexOf("kilkoma");
    p.closest(".docx-edit-root").focus();
    const r = document.createRange(); r.setStart(tn, k); r.setEnd(tn, k + 7);
    getSelection().removeAllRanges(); getSelection().addRange(r);
    document.dispatchEvent(new Event("selectionchange"));
    return r.toString();
  });
  await page.keyboard.press(`${MOD}+d`);
  const opened = await page.waitForSelector(".compose-pop-font", { timeout: 4000 }).then(() => true).catch(() => false);
  check("3: Ctrl/⌘+D otwiera „Czcionka…”", opened && sel === "kilkoma");
  if (opened) {
    await page.check(".ff-smallcaps");
    await page.selectOption(".ff-underline", "double");
    await page.selectOption(".ff-spacing", "expanded");
    await page.fill(".ff-spacing-by", "1,5");
    await page.selectOption(".ff-position", "raised");
    await page.fill(".ff-position-by", "3");
    const sample = await page.evaluate(() => { const s = document.querySelector(".ff-sample span"); const c = getComputedStyle(s); return { v: c.fontVariantCaps || c.fontVariant, d: c.textDecorationStyle }; });
    check("3b: podgląd próbki w oknie pokazuje kapitaliki i podwójne podkreślenie", /small-caps/.test(sample.v) && sample.d === "double", JSON.stringify(sample));
    await page.click(".ff-ok");
    await idle(page);
    const runs = await fileRuns(page, "Akapit do okna");
    const k = runs.find((x) => x.startsWith("kilkoma|")) || "";
    check("4: plik — „kilkoma”: smallCaps, spacing 30, position 6, u double; reszta akapitu bez nich", /smallCaps/.test(k) && /spacing=30/.test(k) && /position=6/.test(k) && /u=double/.test(k) && runs.filter((x) => !x.startsWith("kilkoma|")).every((x) => !/smallCaps|spacing|position|u=/.test(x)), JSON.stringify(runs));
    check("4b: po zastosowaniu — bez „zmienionych” akapitów (podgląd = plik)", (await page.evaluate(() => collectInlineParagraphEdits().length)) === 0);
    // okno czyta bieżące cechy
    await page.keyboard.press(`${MOD}+d`);
    await page.waitForSelector(".compose-pop-font");
    const cur = await page.evaluate(() => ({ sc: document.querySelector(".ff-smallcaps").checked, u: document.querySelector(".ff-underline").value, sp: document.querySelector(".ff-spacing").value, by: document.querySelector(".ff-spacing-by").value, pos: document.querySelector(".ff-position").value }));
    check("5: okno pokazuje bieżące cechy zaznaczenia", cur.sc && cur.u === "double" && cur.sp === "expanded" && /^1,5/.test(cur.by) && cur.pos === "raised", JSON.stringify(cur));
    // wersaliki + zdjęcie kapitalików
    await page.check(".ff-caps");
    await page.uncheck(".ff-smallcaps");
    await page.click(".ff-ok");
    await idle(page);
    const k2 = (await fileRuns(page, "Akapit do okna")).find((x) => x.startsWith("kilkoma|")) || "";
    check("6: wersaliki zamiast kapitalików", /caps/.test(k2) && !/smallCaps/.test(k2), k2);
    // Cofnij
    await page.evaluate(() => document.getElementById("undoBtn").click());
    await idle(page);
    const k3 = (await fileRuns(page, "Akapit do okna")).find((x) => x.startsWith("kilkoma|")) || "";
    check("7: Cofnij — znów kapitaliki", /smallCaps/.test(k3) && !/(^|,)caps/.test(k3.split("|")[1]), k3);
  }
  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`); if (!r.ok) failed++; }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
