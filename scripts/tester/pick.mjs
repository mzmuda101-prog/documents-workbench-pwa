// @ts-check
// Losuje skład testerów AI z docs/testerzy/osoby.md.
// node scripts/tester/pick.mjs [--n=6] [--area=<słowo>] [--seed=<liczba>]  -> JSON na stdout

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

/** @typedef {{ id: string, name: string, device: string, weight: number, areas: string[], task: string, who: string }} Osoba */

const DEVICES = ["iphone", "iphone-landscape", "ipad", "desktop"];
/** @type {Record<string, number>} */
const WEIGHTS = { częste: 3, średnie: 2, rzadkie: 1 };
/** @type {Record<string, string>} */
const KEYS = { urządzenie: "device", waga: "weight", obszary: "areas", kto: "who", zadanie: "task" };

/** @param {string} md @returns {Osoba[]} */
export function parseOsoby(md) {
  return md
    .split(/^## /m)
    .slice(1)
    .map((section) => {
      const [head, ...lines] = section.split("\n");
      const id = head.trim();
      /** @type {Record<string, string>} */
      const f = {};
      let key = "";
      for (const line of lines) {
        const m = line.match(/^- (urządzenie|waga|obszary|kto|zadanie):\s*(.*)$/);
        if (m) {
          key = KEYS[m[1]];
          f[key] = m[2].trim();
        } else if (key && line.trim()) f[key] += ` ${line.trim()}`; // zadanie może ciągnąć się w dół
      }
      for (const [pl, k] of Object.entries(KEYS)) {
        if (!f[k]) throw new Error(`osoby.md, sekcja "${id}": brak pola "${pl}"`);
      }
      if (!/^[a-z0-9]+(-[a-z0-9]+)*$/.test(id)) throw new Error(`osoby.md, sekcja "${id}": id nie jest kebab-case ascii`);
      if (!DEVICES.includes(f.device)) throw new Error(`osoby.md, sekcja "${id}": złe urządzenie "${f.device}"`);
      if (!WEIGHTS[f.weight]) throw new Error(`osoby.md, sekcja "${id}": zła waga "${f.weight}"`);
      return {
        id,
        name: id,
        device: f.device,
        weight: WEIGHTS[f.weight],
        areas: f.areas.split(",").map((a) => a.trim()).filter(Boolean),
        task: f.task,
        who: f.who,
      };
    });
}

/** @param {number} a */
function mulberry32(a) {
  return () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Losowanie ważone bez powtórzeń: zabiera wylosowane z `pool`. @param {Osoba[]} pool @param {number} n @param {() => number} rand @param {(o: Osoba) => boolean} [ok] */
function draw(pool, n, rand, ok = () => true) {
  /** @type {Osoba[]} */
  const out = [];
  while (out.length < n) {
    const cand = pool.filter(ok);
    if (!cand.length) break;
    let r = rand() * cand.reduce((s, o) => s + o.weight, 0);
    const hit = cand.find((o) => (r -= o.weight) < 0) ?? cand[cand.length - 1];
    pool.splice(pool.indexOf(hit), 1);
    out.push(hit);
  }
  return out;
}

/** @param {Osoba[]} osoby @param {{ n?: number, area?: string, seed?: number }} [opts] @returns {Osoba[]} */
export function pick(osoby, { n = 6, area, seed = Date.now() } = {}) {
  const rand = mulberry32(seed);
  if (area) return draw(osoby.filter((o) => o.areas.includes(area)), n, rand);
  const pool = [...osoby];
  // Skład (2 telefon, 1 iPad, 1 komputer) tylko gdy się mieści; przy n < 4 zwykłe losowanie ważone.
  if (n < 4) return draw(pool, n, rand);
  return [
    ...draw(pool, 2, rand, (o) => o.device === "iphone" || o.device === "iphone-landscape"),
    ...draw(pool, 1, rand, (o) => o.device === "ipad"),
    ...draw(pool, 1, rand, (o) => o.device === "desktop"),
    ...draw(pool, n - 4, rand),
  ];
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  /** @type {Record<string, string>} */
  const args = Object.fromEntries(
    process.argv.slice(2).map((a) => {
      const m = a.match(/^--([a-z]+)=(.*)$/);
      if (!m) throw new Error(`nieznany argument: ${a}`);
      return [m[1], m[2]];
    }),
  );
  const root = join(dirname(fileURLToPath(import.meta.url)), "../..");
  const osoby = parseOsoby(readFileSync(join(root, "docs/testerzy/osoby.md"), "utf8"));
  const n = args.n ? Number.parseInt(args.n, 10) : 6;
  const seed = args.seed ? Number.parseInt(args.seed, 10) : Date.now();
  if (!(n > 0) || Number.isNaN(seed)) throw new Error("--n i --seed muszą być liczbami");
  console.log(JSON.stringify(pick(osoby, { n, area: args.area, seed }), null, 2));
}
