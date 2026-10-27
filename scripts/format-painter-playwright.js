// format-painter-playwright.js — malarz formatów (pędzel jak w Wordzie), 2026-10-27.
//
// Klik w pędzel = format znaków z kursora czeka na następne zaznaczenie (raz), dwuklik = pędzel
// zostaje (Esc kończy), klik w słowo bez zaznaczenia maluje całe słowo, Ctrl/⌘+Shift+C / V —
// kopiuj / wklej format. Sprawdzamy zapisany XML (w:b, w:i, w:sz, w:color) i jeden krok Cofnij.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

// format słowa w zapisanym pliku: "b i sz color" (pusta wartość = brak)
const savedLook = (page, word) => page.evaluate(async (word) => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  const doc = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const val = (rPr, tag) => { const e = rPr?.getElementsByTagNameNS(W, tag)[0]; return e ? (e.getAttributeNS(W, "val") || e.getAttribute("w:val") || "1") : ""; };
  const on = (v) => (["", "0", "false"].includes(v) ? "" : "1");
  for (const r of doc.getElementsByTagNameNS(W, "r")) {
    const tx = [...r.getElementsByTagNameNS(W, "t")].map((x) => x.textContent).join("");
    if (!tx.includes(word)) continue;
    const rPr = r.getElementsByTagNameNS(W, "rPr")[0];
    return { text: tx, look: [on(val(rPr, "b")), on(val(rPr, "i")), val(rPr, "sz"), val(rPr, "color").toUpperCase()].join(" ") };
  }
  return null;
}, word);

const selectText = (page, word, collapseAt = -1) => page.evaluate(({ word, collapseAt }) => {
  const p = [...document.querySelectorAll(".docx-editable-p")].find((x) => x.textContent.includes(word));
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  const nodes = [];
  for (let n = walker.nextNode(); n; n = walker.nextNode()) nodes.push(n);
  const at = nodes.map((n) => n.data).join("").indexOf(word);
  const pos = (off) => { for (const n of nodes) { if (off <= n.length) return [n, off]; off -= n.length; } return [nodes.at(-1), nodes.at(-1).length]; };
  const r = document.createRange();
  if (collapseAt >= 0) { r.setStart(...pos(at + collapseAt)); r.collapse(true); } else { r.setStart(...pos(at)); r.setEnd(...pos(at + word.length)); }
  p.closest(".docx-edit-root")?.focus?.({ preventScroll: true });
  getSelection().removeAllRanges();
  getSelection().addRange(r);
  return true;
}, { word, collapseAt });

// puszczenie myszy na akapicie (koniec zaznaczania celu)
const release = (page, word) => page.evaluate((word) => {
  const p = [...document.querySelectorAll(".docx-editable-p")].find((x) => x.textContent.includes(word));
  p.dispatchEvent(new PointerEvent("pointerup", { bubbles: true, button: 0, pointerType: "mouse" }));
}, word).then(() => page.waitForTimeout(200));

const pressed = (page) => page.evaluate(() => document.getElementById("fmtPainterBtn").getAttribute("aria-pressed"));

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

  check("pędzel na pasku Edycji", await page.evaluate(() => !!document.getElementById("fmtPainterBtn")?.offsetParent));

  await page.keyboard.type("Uwaga alfa beta gamma delta epsilon koniec");
  // źródło: „Uwaga” pogrubione, kursywa, 16 pt, czerwone
  await selectText(page, "Uwaga");
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.press(`${MOD}+i`);
  await page.evaluate(() => { applyFontSizePt(16); applyRunFormat({ color: "#C00000" }); });
  await page.waitForTimeout(200);
  const srcLook = (await savedLook(page, "Uwaga"))?.look;
  check("źródło ma format B I 16 pt czerwony", srcLook === "1 1 32 C00000", srcLook);

  // ── 1. pędzel raz: zaznaczenie celu ─────────────────────────────────────────
  await selectText(page, "Uwaga", 2);
  await page.click("#fmtPainterBtn");
  check("klik w pędzel — aktywny", (await pressed(page)) === "true");
  await selectText(page, "beta");
  await release(page, "beta");
  let look = await savedLook(page, "beta");
  check("„beta” dostała B I 16 pt czerwony", look?.look === "1 1 32 C00000" && look.text === "beta", JSON.stringify(look));
  check("po jednym malowaniu pędzel gaśnie", (await pressed(page)) === "false");
  look = await savedLook(page, "gamma");
  check("sąsiad „gamma” bez zmian", look && look.look === "   " && !look.text.includes("beta"), JSON.stringify(look));

  // ── 2. Cofnij = jeden krok ──────────────────────────────────────────────────
  await page.keyboard.press(`${MOD}+z`);
  await page.waitForTimeout(400);
  look = await savedLook(page, "beta");
  check("Cofnij zdejmuje malowanie jednym krokiem", look && look.look === "   ", JSON.stringify(look));

  // ── 3. dwuklik: pędzel przyklejony, klik w słowo maluje całe słowo ─────────
  await selectText(page, "Uwaga", 2);
  await page.dblclick("#fmtPainterBtn");
  await page.waitForTimeout(150);
  check("dwuklik — pędzel zostaje", (await pressed(page)) === "true");
  await selectText(page, "delta", 2); // kursor w środku słowa, bez zaznaczenia
  await release(page, "delta");
  await selectText(page, "epsilon");
  await release(page, "epsilon");
  const d = await savedLook(page, "delta");
  const e = await savedLook(page, "epsilon");
  check("klik w słowo maluje całe „delta”", d?.look === "1 1 32 C00000" && d.text === "delta", JSON.stringify(d));
  check("drugie malowanie tym samym pędzlem — „epsilon”", e?.look === "1 1 32 C00000", JSON.stringify(e));
  check("pędzel nadal aktywny po dwóch malowaniach", (await pressed(page)) === "true");
  await page.keyboard.press("Escape");
  check("Esc wyłącza pędzel", (await pressed(page)) === "false");

  // ── 4. zwykły format przenosi się też na sformatowany tekst (cel traci swój) ─
  await selectText(page, "koniec", 2);
  await page.click("#fmtPainterBtn");
  await selectText(page, "Uwaga");
  await release(page, "Uwaga");
  look = await savedLook(page, "Uwaga");
  check("format zwykłego tekstu zdejmuje B I rozmiar i kolor z „Uwaga” (bez zbędnych wyłączeń)", look?.look === "   ", JSON.stringify(look));

  // ── 5. skróty Ctrl/⌘+Shift+C / V ────────────────────────────────────────────
  await selectText(page, "delta", 1);
  await page.keyboard.press(`${MOD}+Shift+KeyC`);
  await selectText(page, "alfa");
  await page.keyboard.press(`${MOD}+Shift+KeyV`);
  await page.waitForTimeout(200);
  look = await savedLook(page, "alfa");
  check("Ctrl/⌘+Shift+C / V przenosi format na „alfa”", look?.look === "1 1 32 C00000", JSON.stringify(look));
  check("tekst dokumentu bez zmian", await page.evaluate(() => [...document.querySelectorAll(".docx-editable-p")].some((p) => p.textContent.replace(/﻿/g, "") === "Uwaga alfa beta gamma delta epsilon koniec")));

  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : `  (${r.detail})`}`));
  const bad = results.filter((r) => !r.ok).length;
  console.log(`\n[${ENGINE}] ${bad ? `${bad} z ${results.length} nie przeszło` : `wszystkie ${results.length} przeszły`}`);
  process.exit(bad ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
