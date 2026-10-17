// Statystyka wyrazów na prawdziwych plikach: node tools/word-oracle/run-stats-files.js a.docx b.docx
const fs = require("fs");
const path = require("path");
const word = require("./word");
const app = require("./app");
(async () => {
  const files = process.argv.slice(2);
  for (const f of files) {
    const bytes = fs.readFileSync(f);
    let w;
    try { w = word.statistics(bytes); } catch (e) { console.log(`⚠️  ${path.basename(f)}: Word — ${e.message}`); continue; }
    const { browser, page } = await app.openApp();
    await app.loadBytes(page, bytes, "stats-file.docx");
    const a = await page.evaluate(() => documentWordStats(docCanvasEl));
    await browser.close();
    const d = ["words", "chars", "charsWithSpaces", "paragraphs"].map((k) => `${k} ${w[k]}${w[k] === a[k] ? "" : `≠${a[k]}`}`);
    console.log(`${["words", "chars", "charsWithSpaces", "paragraphs"].every((k) => w[k] === a[k]) ? "✅" : "❌"} ${path.basename(f)}: ${d.join(", ")}`);
  }
})();
