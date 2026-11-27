// @ts-check
// tester-drive-playwright.js — sterownik przeglądarki testerów AI (scripts/tester/drive.mjs):
// dwóch testerów naraz, każdy z własną przeglądarką; look/tap/open-file/errors; zapis → pobrany
// plik → ponowne otwarcie; start na zajętym imieniu = błąd; stop-all nie zostawia procesów.
// Użycie: node scripts/tester-drive-playwright.js   (serwer testów :7823 musi działać)

const { spawnSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const DRIVE = path.join(__dirname, "tester", "drive.mjs");
const SESS = path.join(ROOT, "output", "testerzy", ".sessions");
let failed = 0;
function check(name, ok, info = "") {
  console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : `  ${info}`}`);
  if (!ok) failed++;
}
function drive(...args) {
  const r = spawnSync(process.execPath, [DRIVE, ...args], { cwd: ROOT, encoding: "utf8", timeout: 60000 });
  let out = null;
  try { out = JSON.parse(r.stdout.trim().split("\n").pop() || "null"); } catch (_) { /* zostaje null */ }
  return { code: r.status, out, raw: r.stdout + r.stderr };
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const alive = (pid) => { try { process.kill(pid, 0); return true; } catch (_) { return false; } };

async function run() {
  drive("stop-all");
  const a = drive("start", "t-iphone", "--device=iphone", "--run=selftest");
  const b = drive("start", "t-desk", "--device=desktop", "--run=selftest");
  check("start: dwóch testerów naraz", a.code === 0 && b.code === 0, a.raw + b.raw);
  check("start: różne porty", a.out?.port && b.out?.port && a.out.port !== b.out.port, JSON.stringify([a.out, b.out]));
  const pids = fs.existsSync(SESS) ? fs.readdirSync(SESS).map((f) => JSON.parse(fs.readFileSync(path.join(SESS, f), "utf8")).pid) : [];
  check("sesje zapisane z pid", pids.length === 2 && pids.every(alive), JSON.stringify(pids));

  const again = drive("start", "t-desk", "--device=desktop", "--run=selftest");
  check("start na zajętym imieniu = błąd", again.code !== 0 && /już działa/.test(again.out?.error || ""), again.raw);

  for (const n of ["t-iphone", "t-desk"]) {
    const l = drive("look", n);
    const size = l.out?.shot && fs.existsSync(l.out.shot) ? fs.statSync(l.out.shot).size : 0;
    check(`look ${n}: zrzut > 10 kB i elementy`, l.code === 0 && size > 10000 && l.out.elements.length > 3, `${size} B, ${l.out?.elements?.length} el. ${l.raw.slice(0, 300)}`);
  }
  const lp = drive("look", "t-iphone");
  check("look iphone: zrzut w pikselach CSS (390 szer.)", lp.out?.viewport?.width === 390, JSON.stringify(lp.out?.viewport));

  const bad = drive("tap", "t-desk", "Tego przycisku nie ma 12345");
  check("tap nieistniejący tekst = błąd", bad.code !== 0 && !!bad.out?.error, bad.raw);

  const o = drive("open-file", "t-desk", "docs/samples/przewodnik.docx");
  await sleep(2500);
  const l2 = drive("look", "t-desk");
  check("open-file: dokument widać", o.code === 0 && /przewodnik/i.test(JSON.stringify(l2.out)), o.raw + JSON.stringify(l2.out?.elements?.slice(0, 5)));

  // dokument otwiera się w Czytaniu — jak człowiek: „Edytuj”, potem stuknięcie w tekst
  check("tap po tekście przycisku (Edytuj)", drive("tap", "t-desk", "Edytuj").code === 0, "");
  check("w Edycji dokument jest polem edycji", (drive("look", "t-desk").out?.elements || []).some((x) => x.role === "editable"), "");
  drive("tap", "t-desk", "--xy=700,300");
  const ps = drive("paste", "t-desk", "--html=<table><tr><td>WKLEJKA</td><td>123</td></tr></table>", "--text=WKLEJKA\t123");
  await sleep(1200);
  const tx = drive("text", "t-desk");
  check("paste: tabela z HTML w dokumencie (text)", ps.code === 0 && /WKLEJKA/.test(tx.out?.text || ""), ps.raw + (tx.out?.text || "").slice(0, 200));
  drive("type", "t-desk", "x");
  drive("key", "t-desk", "Control+s");
  let dl = [];
  for (let i = 0; i < 10 && !dl.length; i++) { await sleep(500); dl = drive("look", "t-desk").out?.downloads || []; }
  check("zapis: pobrany plik .docx", dl.length >= 1 && /\.docx$/.test(dl[0]) && fs.existsSync(dl[0]), JSON.stringify(dl));
  if (dl.length) {
    drive("errors", "t-desk");
    const re = drive("open-file", "t-desk", dl[0]);
    await sleep(2000);
    const er = drive("errors", "t-desk");
    check("ponowne otwarcie zapisanego pliku bez błędów", re.code === 0 && Array.isArray(er.out?.errors) && er.out.errors.length === 0, re.raw + JSON.stringify(er.out));
    check("po ponownym otwarciu wklejka jest w tekście", /WKLEJKA/.test(drive("text", "t-desk").out?.text || ""), "");
  }
  const tap = drive("tap", "t-iphone", "--xy=195,400");
  check("tap --xy na telefonie", tap.code === 0, tap.raw);
  const e = drive("errors", "t-iphone");
  check("errors zwraca tablicę", Array.isArray(e.out?.errors), e.raw);

  drive("stop-all");
  await sleep(2000);
  check("stop-all: procesy zamknięte", pids.every((p) => !alive(p)), JSON.stringify(pids.filter(alive)));
  check("stop-all: brak sesji", !fs.existsSync(SESS) || fs.readdirSync(SESS).length === 0, "");

  console.log(failed ? `\n${failed} nie przeszło` : "\nwszystko przeszło");
  process.exit(failed ? 1 : 0);
}
run().catch((err) => { console.error(err); drive("stop-all"); process.exit(1); });
