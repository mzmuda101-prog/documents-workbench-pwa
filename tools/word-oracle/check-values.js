// Formaty numerów (słownie, porządkowo, rzymskie, litery) kontra wartości z Worda (values.json)
const fs = require("fs");
const vm = require("vm");
const path = require("path");
const ctx = { W_NS: "", console };
vm.createContext(ctx);
vm.runInContext(`${fs.readFileSync(path.join(__dirname, "../../app/list-numbering.js"), "utf8")};this.L=dwbListNum;`, ctx);
const ref = JSON.parse(fs.readFileSync(path.join(__dirname, "../../output/word-oracle/values.json"), "utf8"));
const extraFile = path.join(__dirname, "../../output/word-oracle/values-extra.json");
if (fs.existsSync(extraFile)) for (const [k, v] of Object.entries(JSON.parse(fs.readFileSync(extraFile, "utf8")))) ref[`${k}`] = { ...ref[k], ...v };
// błędy samego Worda (ogonki w złożeniach) — u nas poprawna polszczyzna
const WORD_BUGS = { "ordinalText/pl-PL": { 5000: "Pięciotysięczny", 6000: "Sześciotysięczny", 8000: "Ośmiotysięczny", 9000: "Dziewięciotysięczny", 10000: "Dziesięciotysięczny", 25000: "Dwudziestopięciotysięczny" } };
let bad = 0;
let all = 0;
for (const [key, vals] of Object.entries(ref)) {
  const [fmt, lang] = key.split("/");
  const wrong = [];
  for (const [v, w] of Object.entries(vals)) {
    if (+v > 32767) continue; // Word zawija większe wartości (nie odtwarzamy)
    all++;
    const a = ctx.L.formatValue(+v, fmt, lang);
    const fixedW = WORD_BUGS[key]?.[v];
    if (fixedW && a === fixedW) continue;
    if (a !== w) wrong.push(`${v}: Word „${w}” / apka „${a}”`);
  }
  bad += wrong.length;
  console.log(`${wrong.length ? "❌" : "✅"} ${key}${wrong.length ? `\n  ${wrong.slice(0, 12).join("\n  ")}` : ""}`);
}
console.log(`${all - bad}/${all} zgodnych`);
process.exitCode = bad ? 1 : 0;
