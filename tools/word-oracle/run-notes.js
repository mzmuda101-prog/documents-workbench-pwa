// run-notes.js — numeracja przypisów (sekcje, numRestart, numFmt, numStart): Word kontra aplikacja
// node tools/word-oracle/run-notes.js [filtr]; WRITE=1 → expected-notes.json (scripts/notes-numbering-playwright.js)
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { buildDocx } = require("./build-docx");
const word = require("./word");
const app = require("./app");
const { CASES, toDocx } = require("./cases-notes");

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
    const bytes = await buildDocx(toDocx(c));
    const w = cached("notes", bytes, () => word.noteLabels(bytes));
    expected[c.name] = { refs: w.refs, notes: w.notes };
    const { browser, page } = await app.openApp();
    await app.loadBytes(page, bytes, `notes-${c.name}.docx`);
    const a = await app.noteLabels(page);
    await browser.close();
    const diffs = [];
    for (const k of Object.keys(w.refs)) if (w.refs[k] !== a.refs[k]) diffs.push(`  odnośnik ${k} (str. ${w.page[k]}): Word „${w.refs[k]}” / apka „${a.refs[k] ?? "∅"}”`);
    for (const k of Object.keys(w.notes)) if (w.notes[k] !== a.notes[k]) diffs.push(`  przypis ${k}: Word „${w.notes[k]}” / apka „${a.notes[k] ?? "∅"}”`);
    console.log(`${diffs.length ? "❌" : "✅"} ${c.name}${diffs.length ? `\n${diffs.slice(0, 12).join("\n")}` : ""}`);
    if (diffs.length) bad++;
  }
  if (process.env.WRITE && !filter) fs.writeFileSync(path.join(__dirname, "expected-notes.json"), `${JSON.stringify(expected, null, 1)}\n`);
  console.log(bad ? `\n${bad} przypadków różni się od Worda` : "\nwszystko jak w Wordzie");
})();
