// run-stats.js — statystyka wyrazów: Word kontra aplikacja (pasek stanu + panel Statystyki)
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { buildDocx } = require("./build-docx");
const word = require("./word");
const app = require("./app");
const CASES = require("./cases-stats");

const CACHE = path.join(word.DIR, "cache");
fs.mkdirSync(CACHE, { recursive: true });
const cached = (kind, bytes, fn) => {
  const f = path.join(CACHE, `${kind}-${crypto.createHash("sha1").update(bytes).digest("hex").slice(0, 16)}.json`);
  if (fs.existsSync(f) && !process.env.FRESH) return JSON.parse(fs.readFileSync(f, "utf8"));
  const r = fn();
  fs.writeFileSync(f, JSON.stringify(r));
  return r;
};

(async () => {
  const filter = process.argv[2];
  let bad = 0;
  const expected = {};
  for (const c of CASES.filter((x) => !filter || x.name.includes(filter))) {
    const bytes = await buildDocx(c);
    const w = cached("stats", bytes, () => word.statistics(bytes));
    expected[c.name] = { words: w.words, chars: w.chars, charsWithSpaces: w.charsWithSpaces, paragraphs: w.paragraphs };
    const { browser, page } = await app.openApp();
    await app.loadBytes(page, bytes, `stats-${c.name}.docx`);
    await page.evaluate(() => ensureLazyFeature("stats"));
    const a = await page.evaluate(() => {
      const s = computeDocumentStats();
      const st = analyzeDocumentDom(docCanvasEl);
      return { words: s.totals.words, chars: s.totals.charsNoSpaces, charsWithSpaces: s.totals.chars, paragraphs: s.a.paragraphs, barWords: st.words };
    });
    await browser.close();
    const diffs = ["words", "chars", "charsWithSpaces", "paragraphs"].filter((k) => w[k] !== a[k]).map((k) => `${k}: Word ${w[k]} / apka ${a[k]}`);
    if (w.words !== a.barWords) diffs.push(`pasek „słowa”: Word ${w.words} / apka ${a.barWords}`);
    console.log(`${diffs.length ? "❌" : "✅"} ${c.name}${diffs.length ? `  — ${diffs.join(", ")}` : ""}`);
    if (diffs.length) bad++;
  }
  // WRITE=1 — wyniki Worda jako oczekiwane dla testu scripts/word-stats-playwright.js
  if (process.env.WRITE && !filter) fs.writeFileSync(path.join(__dirname, "expected-stats.json"), `${JSON.stringify(expected, null, 1)}\n`);
  console.log(bad ? `\n${bad} przypadków różni się od Worda` : "\nwszystko jak w Wordzie");
})();
