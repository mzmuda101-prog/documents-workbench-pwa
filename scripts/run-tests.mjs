// run-tests.mjs — równoległy runner testów (zamiast łańcucha kroków `&&` jeden po drugim).
// Przeniesiony 1:1 z Sheet Workbench (2026-09-28), port i serwer pod tę apkę.
//
// Lista testów jest JEDNA: skrypt "test:serial" w package.json (stary łańcuch, nadal
// działa sam). Runner go rozbiera na kroki i puszcza kilka naraz.
//
// Co robi:
//   - sam startuje serwer testów (scripts/test-server.py :7823), jeśli nie działa;
//   - puszcza N kroków naraz (domyślnie 2; JOBS=… nadpisuje) z obniżonym priorytetem;
//   - najdłuższe zaczyna pierwsze (czasy z poprzedniego przebiegu w .test-timings.json),
//     żeby na końcu nie czekać na jeden maruder;
//   - NIE zatrzymuje się na pierwszym błędzie — na końcu widać wszystkie porażki;
//   - to, co padło, powtarza RAZ, już pojedynczo (bez obciążenia innymi testami).
//     Przejście za drugim razem = „niestabilny" (wypisany osobno, NIE ukrywany),
//     wynik ogólny zielony; padnięcie drugi raz = prawdziwy błąd.
//
// Użycie:  npm test                  (wszystko, 2 naraz, niski priorytet)
//          npm run test:fast         (5 naraz — szybciej, ale komputer mocniej pracuje)
//          npm test -- freeze hero   (tylko kroki, których komenda zawiera któreś słowo)
//          JOBS=1 npm test           (po kolei, jak dawniej)

import { spawn } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import http from "node:http";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const TIMINGS = path.join(ROOT, ".test-timings.json");
const PORT = 7823; // ten sam co APP_URL w scripts/docx-test-helpers.js

const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8"));
const chain = pkg.scripts["test:serial"];
if (!chain) { console.error('Brak "test:serial" w package.json'); process.exit(2); }

// "ENGINE=webkit TOUCH=1 node scripts/x.js" → { env, cmd, args, label }
let steps = chain.split("&&").map((raw) => raw.trim()).filter(Boolean).map((raw) => {
  const parts = raw.split(/\s+/);
  const env = {};
  while (parts.length && /^[A-Z_][A-Z0-9_]*=/.test(parts[0])) {
    const [k, ...v] = parts.shift().split("=");
    env[k] = v.join("=");
  }
  const label = raw.replace(/node scripts\//, "").replace(/-playwright\.js|\.js/, "");
  return { raw, env, cmd: parts[0], args: parts.slice(1), label };
});

const filters = process.argv.slice(2).map((s) => s.toLowerCase());
if (filters.length) steps = steps.filter((s) => filters.some((f) => s.raw.toLowerCase().includes(f)));
if (!steps.length) { console.error("Żaden krok nie pasuje do filtra."); process.exit(2); }

let timings = {};
try { timings = JSON.parse(fs.readFileSync(TIMINGS, "utf8")); } catch { timings = {}; }
steps.sort((a, b) => (timings[b.raw] || 30000) - (timings[a.raw] || 30000));

// Domyślnie SPOKOJNIE: 2 naraz (~2,5 min zamiast ~5 po kolei), żeby komputer nie
// wchodził na pełne obroty (prośba Mateusza 2026-09-25 w Sheet Workbench — wentylator). `npm run test:fast`
// = 5 naraz (~1 min). Dodatkowo każdy test dostaje niższy priorytet (nice), więc
// reszta systemu ma pierwszeństwo.
// 2026-10-02 (prośba Mateusza: szybciej): domyślnie 4 naraz — Mac ma 10 rdzeni, a testy i tak
// idą z obniżonym priorytetem. `npm run test:calm` = 2 naraz (ciszej), JOBS=… nadpisuje.
const JOBS = Math.max(1, parseInt(process.env.JOBS || 4, 10));
// Sztywne pauzy w testach (page.waitForTimeout) w pierwszym przebiegu ×0,5 (sleep-scale.cjs);
// powtórka porażki — pojedynczo i z pełnymi pauzami. SLEEP_SCALE=1 wyłącza skracanie.
const SLEEP_SCALE = process.env.SLEEP_SCALE || "0.5";
const PRELOAD = path.join(ROOT, "scripts", "sleep-scale.cjs");
const NICE = 10;

function serverUp() {
  return new Promise((resolve) => {
    const req = http.get({ host: "127.0.0.1", port: PORT, path: "/", timeout: 1000 }, (res) => { res.resume(); resolve(true); });
    req.on("error", () => resolve(false));
    req.on("timeout", () => { req.destroy(); resolve(false); });
  });
}

let server = null;
async function ensureServer() {
  if (await serverUp()) return;
  server = spawn("python3", [path.join(ROOT, "scripts/test-server.py"), String(PORT)], { cwd: ROOT, stdio: "ignore" });
  for (let i = 0; i < 50; i++) {
    if (await serverUp()) return;
    await new Promise((r) => setTimeout(r, 100));
  }
  throw new Error("Serwer testów nie wstał na porcie " + PORT);
}

function runStep(step, sleepScale = SLEEP_SCALE) {
  return new Promise((resolve) => {
    const t0 = Date.now();
    // SLEEP_SCALE=1 w kroku (package.json) = test na timerach, zawsze pełne pauzy
    const env = { ...process.env, SLEEP_SCALE: sleepScale, ...step.env, NODE_OPTIONS: `${process.env.NODE_OPTIONS || ""} --require ${JSON.stringify(PRELOAD)}`.trim() };
    const child = spawn(step.cmd, step.args, { cwd: ROOT, env });
    try { os.setPriority(child.pid, NICE); } catch {} // przeglądarki odpalone przez test dziedziczą priorytet
    let out = "";
    child.stdout.on("data", (d) => { out += d; });
    child.stderr.on("data", (d) => { out += d; });
    child.on("close", (code) => resolve({ step, ok: code === 0, code, ms: Date.now() - t0, out }));
  });
}

async function runPool(list, jobs, onDone) {
  const results = [];
  let next = 0;
  async function worker() {
    while (next < list.length) {
      const step = list[next++];
      const r = await runStep(step);
      results.push(r);
      onDone(r);
    }
  }
  await Promise.all(Array.from({ length: Math.min(jobs, list.length) }, worker));
  return results;
}

const sec = (ms) => (ms / 1000).toFixed(1).replace(".", ",") + " s";

async function main() {
  await ensureServer();
  const t0 = Date.now();
  console.log(`Testy: ${steps.length} kroków, ${JOBS} naraz, pauzy ×${SLEEP_SCALE}\n`);
  let done = 0;
  const results = await runPool(steps, JOBS, (r) => {
    done += 1;
    console.log(`${r.ok ? "✅" : "❌"} [${done}/${steps.length}] ${r.step.label}  (${sec(r.ms)})`);
  });

  // Powtórka porażek — pojedynczo, bez tłoku.
  const failed = results.filter((r) => !r.ok);
  const flaky = [];
  const broken = [];
  if (failed.length) {
    console.log(`\nPowtarzam ${failed.length} pojedynczo…`);
    for (const r of failed) {
      const again = await runStep(r.step, "1");
      if (again.ok) {
        flaky.push(r);
        console.log(`⚠️  ${r.step.label} — przeszedł za drugim razem (niestabilny w tłoku). Pierwsza porażka:`);
        // najpierw linie z ❌ (które sprawdzenie padło), potem końcówka wyjścia
        const lines = r.out.trim().split("\n");
        const fails = lines.filter((l) => /❌|Error|Timeout/.test(l)).slice(0, 8);
        console.log([...fails, ...lines.slice(-3)].map((l) => "     │ " + l).join("\n"));
      }
      else { broken.push(again); console.log(`❌ ${r.step.label} — padł ponownie`); }
    }
  }

  for (const r of results) if (r.ok) timings[r.step.raw] = r.ms;
  try { fs.writeFileSync(TIMINGS, JSON.stringify(timings, null, 2)); } catch {}

  const wall = Date.now() - t0;
  const serial = results.reduce((a, r) => a + r.ms, 0);
  console.log(`\nCzas: ${sec(wall)} (po kolei byłoby ok. ${sec(serial)})`);
  const slow = [...results].sort((a, b) => b.ms - a.ms).slice(0, 5);
  console.log("Najwolniejsze: " + slow.map((r) => `${r.step.label} ${sec(r.ms)}`).join(" · "));

  if (flaky.length) console.log(`\n⚠️  Niestabilne w tłoku (${flaky.length}): ${flaky.map((r) => r.step.label).join(", ")}`);
  if (broken.length) {
    console.log(`\n❌ NIE PRZESZŁO (${broken.length}):`);
    for (const r of broken) {
      console.log(`\n──── ${r.step.raw} (kod ${r.code}) ────`);
      console.log(r.out.split("\n").slice(-40).join("\n"));
    }
  } else {
    console.log("\n✅ wszystko przeszło");
  }
  return broken.length ? 1 : 0;
}

main()
  .then((code) => { if (server) server.kill(); process.exit(code); })
  .catch((e) => { console.error(e); if (server) server.kill(); process.exit(1); });
