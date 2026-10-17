// run-lists.js — numery list: prawdziwy Word kontra aplikacja (node tools/word-oracle/run-lists.js [filtr])
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { buildDocx } = require("./build-docx");
const word = require("./word");
const app = require("./app");
const CASES = require("./cases-lists");

const CACHE = path.join(word.DIR, "cache");
fs.mkdirSync(CACHE, { recursive: true });
function cachedWord(bytes) {
  const key = crypto.createHash("sha1").update(bytes).digest("hex").slice(0, 16);
  const f = path.join(CACHE, `lists-${key}.json`);
  if (fs.existsSync(f) && !process.env.FRESH) return JSON.parse(fs.readFileSync(f, "utf8"));
  const res = word.listStrings(bytes);
  fs.writeFileSync(f, JSON.stringify(res));
  return res;
}

(async () => {
  const filter = process.argv[2];
  let bad = 0;
  for (const c of CASES.filter((c) => !filter || c.name.includes(filter))) {
    const bytes = await buildDocx(c);
    const w = cachedWord(bytes);
    const { browser, page } = await app.openApp();
    await app.loadBytes(page, bytes, `lists-${c.name}.docx`);
    const a = await app.listLabels(page);
    await browser.close();
    const diffs = [];
    w.forEach((wp, i) => {
      const ap = a[i] || { label: "∅", text: "∅" };
      if (wp.label !== ap.label) diffs.push(`  #${i + 1} „${wp.text}”: Word „${wp.label}” / apka „${ap.label}”`);
    });
    if (w.length !== a.length) diffs.push(`  akapitów: Word ${w.length}, apka ${a.length}`);
    console.log(`${diffs.length ? "❌" : "✅"} ${c.name}${diffs.length ? `\n${diffs.slice(0, 14).join("\n")}${diffs.length > 14 ? `\n  … i ${diffs.length - 14} więcej` : ""}` : ""}`);
    if (diffs.length) bad++;
  }
  console.log(bad ? `\n${bad} przypadków różni się od Worda` : "\nwszystko jak w Wordzie");
})();
