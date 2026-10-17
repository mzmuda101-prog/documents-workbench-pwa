// run-find.js — „Znajdź i zamień” (Zamień wszystko): Word kontra aplikacja (tekst i format po zapisie)
// node tools/word-oracle/run-find.js [filtr]; WRITE=1 → expected-find.json (scripts/find-word-playwright.js)
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { buildDocx } = require("./build-docx");
const word = require("./word");
const app = require("./app");
const { CASES } = require("./cases-find");
const { normalizeParas } = require("./cases-quote");

const as = (s) => `"${String(s).replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
function wordReplace(bytes, c) {
  const key = crypto.createHash("sha1").update(bytes).update(JSON.stringify(c)).digest("hex").slice(0, 16);
  const f = path.join(word.DIR, "cache", `find-${key}.json`);
  if (fs.existsSync(f) && !process.env.FRESH) return JSON.parse(fs.readFileSync(f, "utf8"));
  const r = word.saveDocx(bytes, `
    set fo to find object of text object of it
    clear formatting fo
    clear formatting (replacement of fo)
    execute find fo find text ${as(c.find)} replace with ${as(c.replace)} replace replace all match case ${!!c.mc} match whole word ${!!c.ww} match wildcards false wrap find find continue`, ["word/document.xml", "word/styles.xml"]);
  const out = { doc: r.parts["word/document.xml"], styles: r.parts["word/styles.xml"] };
  fs.mkdirSync(path.dirname(f), { recursive: true });
  fs.writeFileSync(f, JSON.stringify(out));
  return out;
}

async function appReplace(page, c) {
  await page.evaluate(() => { if (readOnlyMode) appFrame.setReadOnly(false); });
  await page.evaluate((c) => applyDocumentEdit({ op: "replace", find: c.find, replace: c.replace, regex: false, matchCase: !!c.mc, wholeWord: !!c.ww, scope: "all", otherParts: false, wordLike: true }), c);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 });
  const x = await app.savedParts(page);
  return page.evaluate(normalizeParas, x);
}

(async () => {
  const filter = process.argv[2];
  let bad = 0;
  const expected = {};
  for (const c of CASES.filter((x) => !filter || x.name.includes(filter))) {
    const bytes = await buildDocx({ body: c.body });
    const wx = wordReplace(bytes, c);
    const { browser, page } = await app.openApp();
    await app.loadBytes(page, bytes, `find-${c.name}.docx`);
    const a = await appReplace(page, c);
    const w = await page.evaluate(normalizeParas, wx);
    await browser.close();
    expected[c.name] = w.map((p) => p.runs);
    const diff = w.map((p, i) => (p.runs === a[i]?.runs ? null : `   Word: ${p.runs}\n   apka: ${a[i]?.runs}`)).filter(Boolean);
    console.log(`${diff.length ? "❌" : "✅"} ${c.name}  [${c.find} → ${c.replace}${c.mc ? ", wielkość" : ""}${c.ww ? ", całe" : ""}]${diff.length ? `\n${diff.join("\n")}` : `  ${w.map((p) => p.runs).join(" ¶ ")}`}`);
    if (diff.length) bad++;
  }
  if (process.env.WRITE && !filter) fs.writeFileSync(path.join(__dirname, "expected-find.json"), `${JSON.stringify(expected, null, 1)}\n`);
  console.log(bad ? `\n${bad} przypadków różni się od Worda` : "\nwszystko jak w Wordzie");
})();
