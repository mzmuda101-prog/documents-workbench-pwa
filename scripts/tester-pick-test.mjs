// @ts-check
// Test losowania składu testerów AI (bez przeglądarki): node scripts/tester-pick-test.mjs
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { parseOsoby, pick } from "./tester/pick.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const md = readFileSync(join(root, "docs/testerzy/osoby.md"), "utf8");
const all = parseOsoby(md);
const isPhone = (/** @type {{device:string}} */ o) => o.device === "iphone" || o.device === "iphone-landscape";

// Każda sekcja prawdziwego pliku ma pięć pól.
const sections = md.split(/^## /m).length - 1;
assert.equal(all.length, sections, "każda sekcja osoby.md się parsuje");
assert.ok(all.length >= 16 && all.length <= 20, `16-20 osób, jest ${all.length}`);
assert.equal(new Set(all.map((o) => o.id)).size, all.length, "id bez powtórzeń");
for (const o of all) {
  assert.ok(o.task && o.who && o.areas.length && [1, 2, 3].includes(o.weight), o.id);
}

// Parser rzuca z nazwą sekcji, gdy brak pola.
assert.throws(
  () => parseOsoby("# x\n\n## bez-zadania\n- urządzenie: ipad\n- waga: częste\n- obszary: pdf\n- kto: ktoś\n"),
  /bez-zadania/,
);
assert.throws(
  () => parseOsoby("## zle-urzadzenie\n- urządzenie: android\n- waga: częste\n- obszary: pdf\n- kto: k\n- zadanie: z\n"),
  /zle-urzadzenie/,
);

// Skład dla n=6, ziarna 1..20.
for (let seed = 1; seed <= 20; seed++) {
  const got = pick(all, { n: 6, seed });
  assert.equal(new Set(got.map((o) => o.id)).size, 6, `seed ${seed}: 6 różnych`);
  assert.ok(got.filter(isPhone).length >= 2, `seed ${seed}: >= 2 telefon`);
  assert.ok(got.some((o) => o.device === "ipad"), `seed ${seed}: >= 1 ipad`);
  assert.ok(got.some((o) => o.device === "desktop"), `seed ${seed}: >= 1 desktop`);
  assert.deepEqual(pick(all, { n: 6, seed }), got, `seed ${seed}: powtarzalne`);
}

// --area=pdf: tylko osoby z pdf; mniej niż n = wszystkie.
const pdfAll = all.filter((o) => o.areas.includes("pdf"));
const pdf = pick(all, { n: 6, area: "pdf", seed: 3 });
assert.ok(pdf.length > 0 && pdf.every((o) => o.areas.includes("pdf")), "area=pdf");
assert.equal(pick(all, { n: 99, area: "pdf", seed: 3 }).length, pdfAll.length, "area: wszystkie, gdy mało");

// Wagi: na osobę „częste” wypada > 2x częściej niż „rzadkie”.
/** @type {Record<number, number>} */
const hits = { 1: 0, 2: 0, 3: 0 };
for (let seed = 1; seed <= 2000; seed++) hits[pick(all, { n: 1, seed })[0].weight]++;
const per = (/** @type {number} */ w) => hits[w] / all.filter((o) => o.weight === w).length;
const ratio = per(3) / per(1);
assert.ok(ratio > 2, `stosunek częste/rzadkie ${ratio.toFixed(2)}`);

// CLI: ten sam seed, ten sam JSON.
const cli = (/** @type {string[]} */ args) =>
  execFileSync(process.execPath, [join(root, "scripts/tester/pick.mjs"), ...args], { encoding: "utf8" });
const a = cli(["--n=6", "--seed=7"]);
assert.equal(a, cli(["--n=6", "--seed=7"]), "CLI powtarzalne");
assert.equal(JSON.parse(a).length, 6);

console.log(`PASS tester-pick: ${all.length} osób, 20 składów, pdf=${pdfAll.length}, częste/rzadkie=${ratio.toFixed(2)}`);
