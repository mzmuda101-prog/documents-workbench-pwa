// list-autoformat-playwright.js — listy z pisania jak w Wordzie (2026-10-31).
//
// Sprawdzone pisaniem w Wordzie (PL, nowy dokument): znacznik na początku akapitu + spacja albo Tab
// → lista w TYM formacie: „1.” 1. 2., „1)” 1) 2), „a.” a. b., „a)” a) b), „A.” A. B., „I.” I. II.,
// „i.” i. ii., „*” •, „>” ➢. NIE zaczynają listy: „-” (zostaje „- tekst”) i „2.”.
// Dawniej aplikacja znała tylko „1.”, „1)” (i numerowała „1.”), „*”, „-”, a „a. ” zamieniała
// autokorektą na „A.” (wielka litera na początku zdania). ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

const CASES = [
  // [klawisze, oczekiwane akapity „[znacznik] tekst”]
  ["1. jeden\ndwa\n", ["[1.] Jeden", "[2.] Dwa", "[3.] "]],
  ["1) jeden\ndwa\n", ["[1)] Jeden", "[2)] Dwa", "[3)] "]],
  ["a. jeden\ndwa\n", ["[a.] Jeden", "[b.] Dwa", "[c.] "]],
  ["a) jeden\ndwa\n", ["[a)] Jeden", "[b)] Dwa", "[c)] "]],
  ["A. jeden\ndwa\n", ["[A.] Jeden", "[B.] Dwa", "[C.] "]],
  ["I. jeden\ndwa\n", ["[I.] Jeden", "[II.] Dwa", "[III.] "]],
  ["i. jeden\ndwa\n", ["[i.] Jeden", "[ii.] Dwa", "[iii.] "]],
  ["1.\tjeden\n", ["[1.] Jeden", "[2.] "]],
  ["a)\tjeden\n", ["[a)] Jeden", "[b)] "]],
  ["* punkt\n", ["[•] Punkt", "[•] "]],
  ["> punkt\n", ["[➢] Punkt", "[➢] "]],
  ["- punkt\n", ["[tekst] - punkt", "[tekst] "]],
  ["2. dwa\n", ["[tekst] 2. dwa", "[tekst] "]],
];

(async () => {
  const browser = await pw[ENGINE].launch({ headless: true });
  const errors = [];
  // każdy przypadek w świeżym kontekście: nowy dokument otwiera się w karcie, a kart jest limit
  for (const [keys, want] of CASES) {
    const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
    // autokorekta jak u ludzi (pod automatem testowym jest domyślnie wyłączona)
    await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.setItem("dwb-autocorrect-v1", JSON.stringify({ on: true })); });
    const page = await context.newPage();
    page.on("pageerror", (e) => errors.push(e.message));
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.evaluate(() => composeUi.createNew("blank"));
    await page.waitForFunction(() => { const h = document.querySelector(".docx-preview-host"); const ps = h ? collectPreviewParagraphElements(h) : []; return document.getElementById("loadingOverlay")?.classList.contains("hidden") && ps.length === 1 && ps[0].classList.contains("docx-editable-p"); }, null, { timeout: 20000 });
    await page.waitForTimeout(500);
    for (const ch of keys) {
      await page.keyboard.press(ch === " " ? "Space" : ch === "\n" ? "Enter" : ch === "\t" ? "Tab" : ch);
      await page.waitForTimeout(50);
    }
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 });
    await page.waitForTimeout(500);
    const got = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => {
      const mark = p.classList.contains("docx-editable-list") ? (dwbListNum.labelOf(p) || getComputedStyle(p, "::before").content.replace(/^"|"$/g, "").replace(/\\9 ?|\s+$/g, "")) : "tekst";
      return `[${mark}] ${p.textContent}`;
    }));
    check(`${JSON.stringify(keys)} → ${want.join(" | ")}`, JSON.stringify(got) === JSON.stringify(want), JSON.stringify(got));
    await context.close();
  }
  check("brak błędów strony", errors.length === 0, errors.join("\n"));
  await browser.close();
  const bad = results.filter((r) => !r.ok);
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `\n   ${r.detail}`}`);
  console.log(`\n[${ENGINE}] ${results.length - bad.length}/${results.length}`);
  process.exit(bad.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
