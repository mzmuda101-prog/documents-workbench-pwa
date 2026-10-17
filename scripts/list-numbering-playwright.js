// list-numbering-playwright.js — numery list jak w Wordzie (app/list-numbering.js).
//
// Wzorzec: prawdziwy Word 16.114 (tools/word-oracle/run-lists.js — 26 przypadków, 2026-10-17).
// Tu bez Worda: wyniki Worda wpisane na sztywno, plus edycja na żywo:
//   • „%1.%2” — każdy poziom w SWOIM formacie („A.i”, nie „i.i”), isLgl = same cyfry,
//   • litery po „z”: aa, bb (nie ab), słownie / porządkowo (Jeden, Pierwszy, 1st),
//   • Enter w środku listy → nowy punkt ma kolejny numer, dalsze przesuwają się,
//   • Shift+Tab (poziom wyżej) → przeliczenie całej listy, Backspace (zdjęcie numeru) też.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(350));
const labels = (page) => page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => `${dwbListNum.labelOf(p) ?? "-"} ${p.textContent.replace(/﻿/g, "")}`));

async function open(page, name, spec) {
  const buffer = await buildDocx(spec);
  await page.locator("#fileInput").setInputFiles({ name, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.evaluate(() => { if (readOnlyMode) appFrame.setReadOnly(false); });
  await idle(page);
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());

  // ── 1) wielopoziomowa: litery + rzymskie, isLgl ─────────────────────────────
  await open(page, "multi.docx", {
    abstractNums: [
      { multi: "multilevel", levels: [{ ilvl: 0, fmt: "upperLetter", text: "%1." }, { ilvl: 1, fmt: "lowerRoman", text: "%1.%2" }] },
      { multi: "multilevel", levels: [{ ilvl: 0, fmt: "upperRoman", text: "%1." }, { ilvl: 1, fmt: "lowerLetter", text: "%1.%2.", isLgl: true }] },
    ],
    nums: [{ id: 1, abs: 0 }, { id: 2, abs: 1 }],
    body: [
      { text: "Alfa", numId: 1 }, { text: "pierwszy", numId: 1, ilvl: 1 }, { text: "drugi", numId: 1, ilvl: 1 }, { text: "Beta", numId: 1 },
      { text: "Przerwa" },
      { text: "Rozdział", numId: 2 }, { text: "ustęp", numId: 2, ilvl: 1 }, { text: "ustęp", numId: 2, ilvl: 1 },
    ],
  });
  let l = await labels(page);
  check("„%1.%2” — każdy poziom w swoim formacie (Word: A.i, A.ii)", l.slice(0, 4).join("|") === "A. Alfa|A.i pierwszy|A.ii drugi|B. Beta", JSON.stringify(l));
  check("numeracja prawnicza (isLgl): Word „1.1.”, nie „I.a.”", l.slice(5).join("|") === "I. Rozdział|1.1. ustęp|1.2. ustęp", JSON.stringify(l));

  // Enter na końcu „drugi” → nowy punkt A.iii, Beta dalej B.
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[2]; placeCaret(p, p.textContent.length); });
  await page.keyboard.press("Enter");
  await page.keyboard.type("trzeci");
  await idle(page);
  l = await labels(page);
  check("Enter w liście: nowy punkt dostaje kolejny numer (A.iii)", l.slice(0, 5).join("|") === "A. Alfa|A.i pierwszy|A.ii drugi|A.iii trzeci|B. Beta", JSON.stringify(l));
  // Shift+Tab na „trzeci” → poziom wyżej: staje się B., a Beta → C.
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[3]; placeCaret(p, 0); });
  await page.keyboard.press("Shift+Tab");
  await idle(page);
  l = await labels(page);
  check("Shift+Tab: punkt poziom wyżej, dalsze numery przeliczone (B., C.)", l.slice(0, 5).join("|") === "A. Alfa|A.i pierwszy|A.ii drugi|B. trzeci|C. Beta", JSON.stringify(l));
  // Backspace na początku „trzeci” → numer znika, Beta wraca do B.
  await page.keyboard.press("Backspace");
  await idle(page);
  l = await labels(page);
  check("Backspace zdejmuje numer: dalsze punkty wracają (Beta = B.)", l[3] === "- trzeci" && l[4] === "B. Beta", JSON.stringify(l));

  // ── 2) litery po „z”, słownie, porządkowo ──────────────────────────────────
  await open(page, "letters.docx", {
    abstractNums: [{ levels: [{ fmt: "lowerLetter", text: "%1)" }] }, { levels: [{ fmt: "cardinalText", text: "%1" }] }, { levels: [{ fmt: "ordinalText", text: "%1:" }] }, { levels: [{ fmt: "ordinal", text: "%1" }] }, { levels: [{ fmt: "decimal", text: "" }] }],
    nums: [{ id: 1, abs: 0 }, { id: 2, abs: 1, overrides: [{ ilvl: 0, start: 21 }] }, { id: 3, abs: 2, overrides: [{ ilvl: 0, start: 100 }] }, { id: 4, abs: 3 }, { id: 5, abs: 4 }],
    body: [
      ...Array.from({ length: 28 }, (_, i) => ({ text: `p${i + 1}`, numId: 1 })),
      { text: "słownie", numId: 2 }, { text: "porządkowo", numId: 3 }, { text: "porządkowo2", numId: 3 },
      { text: "en", numId: 4, lang: "en-US" }, { text: "pl", numId: 4 }, { text: "pusty numer", numId: 5 },
    ],
  });
  l = await labels(page);
  check("litery po „z”: aa), bb) — jak Word (CSS dawał ab)", l[26] === "aa) p27" && l[27] === "bb) p28", JSON.stringify(l.slice(25, 28)));
  check("słownie (cardinalText): „Dwadzieścia jeden”", l[28] === "Dwadzieścia jeden słownie", l[28]);
  check("porządkowo (ordinalText): „Setny:”, „Sto pierwszy:”", l[29] === "Setny: porządkowo" && l[30] === "Sto pierwszy: porządkowo2", JSON.stringify(l.slice(29, 31)));
  check("ordinal: angielski akapit „1st”, polski „2.”", l[31] === "1st en" && l[32] === "2. pl", JSON.stringify(l.slice(31, 33)));
  check("pusty tekst numeru: bez „1.” (Word nic nie pokazuje)", l[33] === " pusty numer" || l[33] === "- pusty numer" || l[33].startsWith(" "), JSON.stringify(l[33]));
  const shown = await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[27];
    return getComputedStyle(p, "::before").content;
  });
  check("widoczny numer = policzony (::before z data-dwb-num)", /bb\)/.test(shown), shown);

  check("brak błędów strony", errors.length === 0, errors.join(" | "));
  await browser.close();
  const bad = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`));
  console.log(`\n[${ENGINE}] ${results.length - bad.length}/${results.length} OK`);
  process.exit(bad.length ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
