// forms-light-playwright.js — pole wyboru / lista / data BEZ przerysowania dokumentu (2026-10-04).
//
// Dawniej każdy klik w ☐ albo wybór z listy = render całego dokumentu (przy długim pliku sekundy,
// mignięcie). Teraz zmiana idzie do pliku kolejką zmian (jak Enter/Backspace), a w podglądzie
// zmienia się tylko tekst pola. Sprawdza: zero przerysowań, plik = podgląd, pola powiązane
// (to samo źródło danych) zmieniają się razem, Cofnij, pisanie w akapicie z polem przed i po
// zmianie (nic nie ginie, kontrolka nie wraca do starej wartości), pole z tekstem zastępczym —
// dalej pełna (bezpieczna) ścieżka. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden")
  && !inlineLocksPending && !inlineStructuralPending && formScan?.bytes === originalFileBytes, null, { timeout: 20000 }).then(() => page.waitForTimeout(300));

const sdtState = (page, fromSave = false) => page.evaluate(async (fromSave) => {
  const bytes = fromSave ? await buildDocumentForSave() : originalFileBytes;
  const z = await JSZip.loadAsync(bytes);
  const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
  const out = {};
  [...d.getElementsByTagNameNS(W_NS, "sdt")].forEach((s, i) => { out[`s${i}`] = ffText(ffKid(s, "sdtContent")); });
  out.paras = (await extractParagraphTextsFromDocx(bytes)).map(String);
  out.legacy = collectLegacyFormFields(d).map((lf) => ffAttr(lf.ffData.getElementsByTagNameNS(W_NS, "checked")[0], "val") || ffAttr(lf.ffData.getElementsByTagNameNS(W_NS, "default")[0], "val"));
  return out;
}, fromSave);

async function clickField(page, key) {
  const box = await page.evaluate((k) => {
    const r = formUi.ranges.get(k);
    const el = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
    el.scrollIntoView({ block: "center" });
    const rect = [...r.getClientRects()].find((q) => q.width > 0) || r.getBoundingClientRect();
    return { x: rect.left + Math.min(rect.width / 2, 20), y: rect.top + rect.height / 2 };
  }, key);
  await page.mouse.click(box.x, box.y);
  await page.waitForTimeout(150);
}
const wrapText = (page, key) => page.evaluate((k) => document.querySelector(`.docx-preview-host .ff-field[data-ff="${k}"]`)?.textContent, key);
const renders = (page) => page.evaluate(() => window.__renders);

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=forms-sample`, { waitUntil: "load" });
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle(page);
  await page.evaluate(() => { window.__renders = 0; const orig = window.renderDocxPreview; window.renderDocxPreview = function (...a) { window.__renders++; return orig.apply(this, a); }; });

  // ── pole wyboru (kontrolka) ────────────────────────────────────────────────
  await clickField(page, "s5");
  await idle(page);
  let s = await sdtState(page);
  check("klik w ☐: bez przerysowania dokumentu", (await renders(page)) === 0, await renders(page));
  check("klik w ☐: w pliku ☒ i w podglądzie ☒", s.s5 === "☒" && (await wrapText(page, "s5")) === "☒", `${s.s5} / ${await wrapText(page, "s5")}`);
  check("„Zapisz” świeci (zmiana jest)", await page.evaluate(() => hasUnsavedChanges));
  await clickField(page, "s5");
  await idle(page);
  s = await sdtState(page);
  check("drugi klik odznacza (☐), dalej bez przerysowania", s.s5 === "☐" && (await wrapText(page, "s5")) === "☐" && (await renders(page)) === 0);

  // ── lista (okienko) ────────────────────────────────────────────────────────
  await clickField(page, "s2");
  await page.waitForSelector(".ff-pop .ff-pop-opt");
  const opt = await page.evaluate(() => [...document.querySelectorAll(".ff-pop .ff-pop-opt")].map((b) => b.textContent).find((x) => x !== "IT"));
  await page.click(`.ff-pop .ff-pop-opt:text-is("${opt}")`);
  await idle(page);
  s = await sdtState(page);
  check("wybór z listy: plik i podgląd, bez przerysowania", s.s2 === opt && (await wrapText(page, "s2")) === opt && (await renders(page)) === 0, `${opt}: ${s.s2} / ${await wrapText(page, "s2")} / r=${await renders(page)}`);

  // ── pola powiązane (to samo źródło danych): zmiana jednego zmienia oba ─────
  await clickField(page, "s9");
  await page.waitForSelector(".ff-pop .ff-pop-input");
  await page.fill(".ff-pop .ff-pop-input", "Nowa Firma S.A.");
  await page.keyboard.press("Enter");
  await idle(page);
  s = await sdtState(page);
  check("pola powiązane: oba w pliku i w podglądzie, bez przerysowania", s.s9 === "Nowa Firma S.A." && s.s10 === "Nowa Firma S.A." && (await wrapText(page, "s10")) === "Nowa Firma S.A." && (await renders(page)) === 0, JSON.stringify({ s9: s.s9, s10: s.s10, w: await wrapText(page, "s10"), r: await renders(page) }));

  // ── stare pole wyboru (FORMCHECKBOX) ───────────────────────────────────────
  const legacyBefore = (await sdtState(page)).legacy.join(",");
  await clickField(page, "f1");
  await idle(page);
  const legacyAfter = (await sdtState(page)).legacy.join(",");
  check("stare pole wyboru: zmiana w pliku i znaczek, bez przerysowania", legacyBefore !== legacyAfter && (await renders(page)) === 0, `${legacyBefore} → ${legacyAfter}`);

  // ── pisanie w akapicie z polem: przed i po zmianie pola ────────────────────
  const islandPara = await page.evaluate(() => { const w = document.querySelector('.docx-preview-host .ff-field[data-ff="s5"]'); return resolveParaIndex(w.closest("p")); });
  await page.evaluate((i) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; placeCaret(p, Number.MAX_SAFE_INTEGER); }, islandPara);
  await page.keyboard.type(" Przed");
  await clickField(page, "s5");
  await idle(page);
  await page.evaluate((i) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; placeCaret(p, Number.MAX_SAFE_INTEGER); }, islandPara);
  await page.keyboard.type(" Po");
  await page.waitForTimeout(300);
  s = await sdtState(page, true);
  check("pisanie przed i po kliknięciu pola w tym samym akapicie: tekst cały, pole ☒ (nie wraca do starej wartości)", s.s5 === "☒" && / Przed Po$/.test(s.paras[islandPara]), `${s.s5} | ${s.paras[islandPara]}`);

  // ── Cofnij ─────────────────────────────────────────────────────────────────
  await page.keyboard.press(`${MOD}+KeyZ`); // pisanie „ Po”
  await idle(page);
  await page.keyboard.press(`${MOD}+KeyZ`); // klik w ☐
  await idle(page);
  s = await sdtState(page, true);
  check("Cofnij: pole wraca do ☐, tekst sprzed kliknięcia zostaje", s.s5 === "☐" && / Przed$/.test(s.paras[islandPara]) && (await wrapText(page, "s5")) === "☐", `${s.s5} | ${s.paras[islandPara]}`);

  // ── pole z tekstem zastępczym: pełna (bezpieczna) ścieżka ──────────────────
  await page.evaluate(() => { window.__renders = 0; });
  await clickField(page, "s3");
  await page.waitForSelector(".ff-pop .ff-pop-opt");
  await page.click(".ff-pop .ff-pop-opt >> nth=0");
  await idle(page);
  const ph = await page.evaluate(() => formScan.fields.find((f) => f.key === "s3"));
  check("pole z tekstem zastępczym: wartość wpisana (przez przerysowanie — styl zastępczy znika)", !ph.placeholder && (await renders(page)) === 1, JSON.stringify({ ph: ph.placeholder, r: await renders(page) }));

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : `\n   ${String(r.detail).slice(0, 700)}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
