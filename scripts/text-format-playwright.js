// text-format-playwright.js — indeks górny / dolny, „Wyczyść formatowanie” i „Aa” (2026-10-10).
//
// Jak Narzędzia główne w Wordzie: x² / x₂ (też Ctrl/⌘+Shift+= / Ctrl/⌘+=, bez zaznaczenia —
// dalsze pisanie), gumka (format znaków zaznaczenia; akapit w całości → styl Normalny),
// Aa (jak w zdaniu / małe / WIELKIE / Jak Nazwa Własna / zAMIANA). Sprawdzamy zapisany XML:
// w:vertAlign tylko na właściwych znakach — także po ponownym wczytaniu pliku (podgląd rysuje
// indeks jako <sup>/<sub>; dawniej model fragmentu go nie znał i edycja akapitu mogła go zgubić).
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

// fragmenty akapitu z needle w zapisanym pliku: [tekst, vertAlign, b, kolor]
const savedRuns = (page, needle) => page.evaluate(async (needle) => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  const doc = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const p = [...doc.getElementsByTagNameNS(W, "p")].find((x) => x.textContent.includes(needle));
  if (!p) return null;
  const val = (rPr, tag) => { const e = rPr?.getElementsByTagNameNS(W, tag)[0]; return e ? (e.getAttributeNS(W, "val") || e.getAttribute("w:val") || "1") : ""; };
  const out = [];
  for (const r of p.getElementsByTagNameNS(W, "r")) {
    const t = [...r.getElementsByTagNameNS(W, "t")].map((x) => x.textContent).join("");
    if (!t) continue;
    const rPr = r.getElementsByTagNameNS(W, "rPr")[0];
    const key = [val(rPr, "vertAlign"), !["", "0", "false"].includes(val(rPr, "b")), val(rPr, "color")];
    const last = out[out.length - 1];
    if (last && last[1] === key[0] && last[2] === key[1] && last[3] === key[2]) last[0] += t; else out.push([t, ...key]);
  }
  const style = p.getElementsByTagNameNS(W, "pStyle")[0];
  return { runs: out, style: style ? style.getAttributeNS(W, "val") || style.getAttribute("w:val") : "" };
}, needle);

// zaznacz tekst `word` (n-te wystąpienie) w akapicie z needle
const selectText = (page, needle, word, nth = 0) => page.evaluate(({ needle, word, nth }) => {
  const p = [...document.querySelectorAll(".docx-editable-p")].find((x) => x.textContent.includes(needle));
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  const nodes = [];
  for (let n = walker.nextNode(); n; n = walker.nextNode()) nodes.push(n);
  const full = nodes.map((n) => n.data).join("");
  let at = -1;
  for (let i = 0; i <= nth; i++) at = full.indexOf(word, at + 1);
  if (at < 0) return false;
  const pos = (off) => { for (const n of nodes) { if (off <= n.length) return [n, off]; off -= n.length; } return [nodes.at(-1), nodes.at(-1).length]; };
  const r = document.createRange();
  r.setStart(...pos(at));
  r.setEnd(...pos(at + word.length));
  p.closest(".docx-edit-root")?.focus?.({ preventScroll: true });
  const sel = getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  return true;
}, { needle, word, nth });

// kursor na końcu akapitu z needle (klawisz End na Macu nie przesuwa kursora — przewija)
const caretEnd = (page, needle) => page.evaluate((needle) => {
  const p = [...document.querySelectorAll(".docx-editable-p")].find((x) => x.textContent.includes(needle));
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  let last = null;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) if (n.data.replace(/\uFEFF/g, "")) last = n;
  const r = document.createRange();
  if (last) r.setStart(last, last.length); else r.setStart(p, p.childNodes.length);
  r.collapse(true);
  p.closest(".docx-edit-root")?.focus?.({ preventScroll: true });
  getSelection().removeAllRanges();
  getSelection().addRange(r);
}, needle);

const pressed = (page) => page.evaluate(() => ({ sup: document.getElementById("fmtSuper").getAttribute("aria-pressed"), sub: document.getElementById("fmtSub").getAttribute("aria-pressed") }));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);

  check("przyciski x² / x₂ / Aa / gumka na pasku Edycji", await page.evaluate(() => ["fmtSuper", "fmtSub", "fmtCaseBtn", "fmtClearBtn"].every((id) => document.getElementById(id)?.offsetParent)));

  // ── 1. indeks na zaznaczeniu ────────────────────────────────────────────────
  await page.keyboard.type("Pole 20 m2 i woda H2O koniec");
  await selectText(page, "Pole 20", "2", 1); // „2” w „m2”
  await page.click("#fmtSuper");
  await page.waitForTimeout(150);
  check("x² wciśnięty na zaznaczonym indeksie", (await pressed(page)).sup === "true", JSON.stringify(await pressed(page)));
  await selectText(page, "Pole 20", "2", 2); // „2” w „H2O”
  await page.click("#fmtSub");
  await page.waitForTimeout(150);
  let saved = await savedRuns(page, "Pole 20");
  check("zapis: „2” w m² = superscript, w H₂O = subscript, reszta bez indeksu",
    JSON.stringify(saved.runs.map((r) => [r[0], r[1]])) === JSON.stringify([["Pole 20 m", ""], ["2", "superscript"], [" i woda H", ""], ["2", "subscript"], ["O koniec", ""]]), JSON.stringify(saved));
  check("podgląd: indeks mniejszy od tekstu", await page.evaluate(() => {
    const p = [...document.querySelectorAll(".docx-editable-p")].find((x) => x.textContent.includes("Pole 20"));
    const sup = [...p.querySelectorAll("span")].find((s) => s.style.verticalAlign === "super");
    return sup && parseFloat(getComputedStyle(sup).fontSize) < parseFloat(getComputedStyle(p).fontSize);
  }));

  await page.waitForTimeout(200);
  const sizeAtSub = await page.evaluate(() => document.getElementById("fmtFontSize").value);
  await selectText(page, "Pole 20", "woda");
  await page.waitForTimeout(200);
  const sizeAtText = await page.evaluate(() => document.getElementById("fmtFontSize").value);
  check("pole rozmiaru przy indeksie = rozmiar tekstu obok (jak Word), nie pomniejszony", sizeAtSub !== "" && sizeAtSub === sizeAtText, JSON.stringify({ sizeAtSub, sizeAtText }));
  // drugi klik zdejmuje
  await selectText(page, "Pole 20", "2", 1);
  await page.click("#fmtSuper");
  await page.waitForTimeout(150);
  saved = await savedRuns(page, "Pole 20");
  check("drugi klik x² zdejmuje indeks (H₂O zostaje)", saved.runs.filter((r) => r[1]).map((r) => r[1]).join() === "subscript", JSON.stringify(saved));

  // ── 2. skrót bez zaznaczenia = dalsze pisanie ────────────────────────────────
  await caretEnd(page, "Pole 20");
  await page.keyboard.press(`${MOD}+Shift+Equal`);
  await page.keyboard.type("x");
  await page.keyboard.press(`${MOD}+Shift+Equal`);
  await page.keyboard.type("y");
  saved = await savedRuns(page, "Pole 20");
  const tail = saved.runs.slice(-2).map((r) => [r[0], r[1]]);
  check("Ctrl/⌘+Shift+= bez zaznaczenia: „x” w indeksie, po drugim — „y” zwykłe", JSON.stringify(tail) === JSON.stringify([["x", "superscript"], ["y", ""]]), JSON.stringify(saved.runs));

  // ── 3. plik z indeksem z Worda (podgląd: <sup>) — edycja akapitu go nie gubi ─────
  await page.evaluate(async () => { await reloadFromBytes(await buildDocumentForSave()); });
  await idle(page);
  await page.evaluate(() => appFrame.setReadOnly(false));
  await page.waitForTimeout(300);
  check("po wczytaniu podgląd rysuje indeks jako <sub>", await page.evaluate(() => !![...document.querySelectorAll(".docx-editable-p")].find((x) => x.textContent.includes("Pole 20"))?.querySelector("sub")));
  await selectText(page, "Pole 20", "koniec");
  await page.keyboard.type("KONIEC");
  saved = await savedRuns(page, "Pole 20");
  check("edycja akapitu z <sub>: indeksy zostają tylko na „2” w H₂O i „x”", JSON.stringify(saved.runs.filter((r) => r[1]).map((r) => [r[0], r[1]])) === JSON.stringify([["2", "subscript"], ["x", "superscript"]]) && saved.runs.some((r) => r[0].includes("KONIEC") && !r[1]), JSON.stringify(saved.runs));
  // x² na tekście w <sub> z pliku: dolny → górny (bez zagnieżdżenia)
  await selectText(page, "Pole 20", "2", 2);
  await page.waitForTimeout(200); // stan paska nadąża za selectionchange
  check("x₂ wciśnięty w <sub> z pliku", (await pressed(page)).sub === "true", JSON.stringify(await pressed(page)));
  await page.click("#fmtSuper");
  await page.waitForTimeout(150);
  saved = await savedRuns(page, "Pole 20");
  check("x² na <sub> z pliku: dolny → górny", saved.runs.some((r) => r[0] === "2" && r[1] === "superscript") && !saved.runs.some((r) => r[1] === "subscript"), JSON.stringify(saved.runs));

  // ── 4. wyczyść formatowanie ─────────────────────────────────────────────────
  await caretEnd(page, "Pole 20");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Słowo ");
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.type("gruby");
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.type(" i dalej");
  await selectText(page, "Słowo gruby", "gruby");
  await page.click("#fmtClearBtn");
  await page.waitForTimeout(300);
  saved = await savedRuns(page, "Słowo gruby");
  check("gumka: zaznaczone słowo bez pogrubienia", saved && !saved.runs.some((r) => r[2]), JSON.stringify(saved));
  check("gumka: zaznaczenie zostaje na tekście", await page.evaluate(() => getSelection().toString()) === "gruby");

  // nagłówek, kursor w środku → akapit wraca do Normalnego
  await caretEnd(page, "Słowo gruby");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Tytuł rozdziału");
  await page.selectOption("#fmtParaStyle", "h1");
  await idle(page);
  saved = await savedRuns(page, "Tytuł rozdziału");
  const hadStyle = saved?.style;
  await selectText(page, "Tytuł rozdziału", "rozdz");
  await page.evaluate(() => getSelection().collapseToStart());
  await page.click("#fmtClearBtn");
  await idle(page);
  saved = await savedRuns(page, "Tytuł rozdziału");
  check("gumka bez zaznaczenia w nagłówku → styl Normalny", !!hadStyle && !saved.style, JSON.stringify({ hadStyle, after: saved }));

  // ── 5. Aa — wielkość liter ──────────────────────────────────────────────────
  await caretEnd(page, "Tytuł rozdziału");
  await page.keyboard.press("Enter");
  await page.keyboard.type("ala MA kota. ala ma psa");
  // okienko zamyka przewinięcie dokumentu (celowo) — najpierw czekamy, aż płynne przewijanie stanie
  const scrollSettled = () => page.evaluate(() => new Promise((res) => {
    const vp = document.getElementById("docViewport");
    let last = vp.scrollTop, still = 0;
    const tick = () => { if (vp.scrollTop === last) still++; else { still = 0; last = vp.scrollTop; } if (still >= 6) res(); else requestAnimationFrame(tick); };
    tick();
  }));
  const pickCase = async (key) => {
    await scrollSettled();
    await page.click("#fmtCaseBtn");
    await page.waitForSelector(".compose-pop-case .compose-item");
    const info = await page.evaluate((k) => {
      const b = [...document.querySelectorAll(".compose-pop-case .compose-item")].find((x) => x.textContent.trim() === t(k));
      const r = b?.getBoundingClientRect();
      return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null;
    }, key);
    if (!info) throw new Error(`brak pozycji ${key}`);
    await page.mouse.click(info.x, info.y);
    await page.waitForTimeout(150);
  };
  const paraText = () => page.evaluate(() => [...document.querySelectorAll(".docx-editable-p")].find((x) => /ma (kota|psa)/i.test(x.textContent))?.textContent);
  await selectText(page, "ala MA kota", "ala MA kota. ala ma psa");
  await pickCase("caseSentence");
  check("Aa: Jak w zdaniu", await paraText() === "Ala ma kota. Ala ma psa", await paraText());
  await pickCase("caseUpper");
  check("Aa: WIELKIE (zaznaczenie zostaje)", await paraText() === "ALA MA KOTA. ALA MA PSA", await paraText());
  await pickCase("caseTitle");
  check("Aa: Jak Nazwa Własna", await paraText() === "Ala Ma Kota. Ala Ma Psa", await paraText());
  await pickCase("caseToggle");
  check("Aa: zAMIANA", await paraText() === "aLA mA kOTA. aLA mA pSA", await paraText());
  await page.click("#undoBtn");
  await idle(page);
  check("Cofnij przywraca poprzednią wielkość liter", await paraText() === "Ala Ma Kota. Ala Ma Psa", await paraText());
  // bez zaznaczenia: słowo przy kursorze
  await selectText(page, "Ala Ma Kota", "Kota");
  await page.evaluate(() => getSelection().collapseToStart());
  await pickCase("caseLower");
  check("Aa bez zaznaczenia: tylko słowo przy kursorze", (await paraText()).startsWith("Ala Ma kota. Ala Ma Psa"), await paraText());
  saved = await savedRuns(page, "Ala Ma kota");
  check("zapis po Aa: nowy tekst w pliku", !!saved && saved.runs.map((r) => r[0]).join("").startsWith("Ala Ma kota. Ala Ma Psa"), JSON.stringify(saved));

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ format tekstu [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
