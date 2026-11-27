// @ts-check
// tester-drive-playwright.js — sterownik przeglądarki testerów AI (scripts/tester/drive.mjs):
// dwóch testerów naraz, każdy z własną przeglądarką; look/tap/open-file/errors; zapis → pobrany
// plik → ponowne otwarcie; start na zajętym imieniu = błąd; stop nie zostawia procesów (także
// zawieszonych); odczucie (martwe, reakcja). Zatrzymuje TYLKO swoich testerów — można go puścić
// w trakcie przebiegu /testerzy.
// Użycie: node scripts/tester-drive-playwright.js   (serwer testów :7823 musi działać)

const { spawn, spawnSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const DRIVE = path.join(__dirname, "tester", "drive.mjs");
const SESS = path.join(ROOT, "output", "testerzy", ".sessions");
const MINE = ["t-iphone", "t-desk", "t-race", "t-hang", "t-bad"];
let failed = 0;
function check(name, ok, info = "") {
  console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : `  ${info}`}`);
  if (!ok) failed++;
}
const parse = (stdout) => { try { return JSON.parse(stdout.trim().split("\n").pop() || "null"); } catch (_) { return null; } };
function drive(...args) {
  const r = spawnSync(process.execPath, [DRIVE, ...args], { cwd: ROOT, encoding: "utf8", timeout: 90000 });
  return { code: r.status, out: parse(r.stdout || ""), raw: (r.stdout || "") + (r.stderr || "") };
}
function driveAsync(args, env = {}) {
  return new Promise((resolve) => {
    const c = spawn(process.execPath, [DRIVE, ...args], { cwd: ROOT, env: { ...process.env, ...env } });
    let s = "";
    c.stdout.on("data", (d) => (s += d));
    c.stderr.on("data", (d) => (s += d));
    c.on("close", (code) => resolve({ code, out: parse(s), raw: s }));
  });
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const alive = (pid) => { try { process.kill(pid, 0); return true; } catch (_) { return false; } };
const sess = (n) => { try { return JSON.parse(fs.readFileSync(path.join(SESS, `${n}.json`), "utf8")); } catch (_) { return null; } };
const stopMine = () => MINE.forEach((n) => drive("stop", n));

async function run() {
  stopMine();
  fs.rmSync(path.join(ROOT, "output", "testerzy", "selftest"), { recursive: true, force: true });

  // apka nie odpowiada: jasny komunikat, szybko (nie wisi)
  const t0 = Date.now();
  const badApp = await driveAsync(["start", "t-bad", "--device=desktop", "--run=selftest"], { APP_URL: "http://127.0.0.1:9/" });
  check("start bez apki: komunikat, < 10 s", badApp.code !== 0 && /apka nie odpowiada/.test(badApp.out?.error || "") && Date.now() - t0 < 10000, badApp.raw);
  const badRun = drive("start", "t-bad", "--device=desktop", "--run=../../poza");
  check("--run spoza output/ odrzucone", badRun.code !== 0, badRun.raw);

  const [a, b] = await Promise.all([driveAsync(["start", "t-iphone", "--device=iphone", "--run=selftest"]), driveAsync(["start", "t-desk", "--device=desktop", "--run=selftest"])]);
  check("start: dwóch testerów naraz", a.code === 0 && b.code === 0, a.raw + b.raw);
  check("start: różne porty", a.out?.port && b.out?.port && a.out.port !== b.out.port, JSON.stringify([a.out, b.out]));
  const pids = ["t-iphone", "t-desk"].map((n) => sess(n)?.pid);
  check("sesje zapisane z pid", pids.every((p) => p && alive(p)), JSON.stringify(pids));

  const again = drive("start", "t-desk", "--device=desktop", "--run=selftest");
  check("start na zajętym imieniu = błąd", again.code !== 0 && /już działa/.test(again.out?.error || ""), again.raw);
  const race = await Promise.all([driveAsync(["start", "t-race", "--device=desktop", "--run=selftest"]), driveAsync(["start", "t-race", "--device=desktop", "--run=selftest"])]);
  check("dwa starty tego samego imienia naraz: działa jeden", race.filter((r) => r.code === 0).length === 1, race.map((r) => r.raw).join(" | "));
  drive("stop", "t-race");

  // śmieci w katalogu sesji (Finder, plik w trakcie zapisu) nie psują poleceń
  fs.writeFileSync(path.join(SESS, ".DS_Store"), "x");
  fs.writeFileSync(path.join(SESS, "zepsuty.json"), "{");
  const lj = drive("look", "t-iphone");
  check("look mimo śmieci w .sessions", lj.code === 0 && !!lj.out?.shot, lj.raw.slice(0, 300));
  fs.rmSync(path.join(SESS, ".DS_Store"), { force: true });
  fs.rmSync(path.join(SESS, "zepsuty.json"), { force: true });

  for (const n of ["t-iphone", "t-desk"]) {
    const l = drive("look", n);
    const size = l.out?.shot && fs.existsSync(l.out.shot) ? fs.statSync(l.out.shot).size : 0;
    check(`look ${n}: zrzut > 10 kB i elementy`, l.code === 0 && size > 10000 && l.out.elements.length > 3, `${size} B, ${l.out?.elements?.length} el. ${l.raw.slice(0, 300)}`);
  }
  check("look iphone: zrzut w pikselach CSS (390 szer.)", drive("look", "t-iphone").out?.viewport?.width === 390, "");
  const two = await Promise.all([driveAsync(["look", "t-desk"]), driveAsync(["look", "t-desk"])]);
  check("dwa polecenia naraz: oba działają, osobne zrzuty", two.every((r) => r.code === 0) && two[0].out?.shot !== two[1].out?.shot, two.map((r) => r.raw.slice(0, 120)).join(" | "));

  const bad = drive("tap", "t-desk", "Tego przycisku nie ma 12345");
  check("tap nieistniejący tekst = błąd", bad.code !== 0 && !!bad.out?.error, bad.raw);

  const o = drive("open-file", "t-desk", "docs/samples/przewodnik.docx");
  await sleep(2500);
  check("open-file: dokument widać", o.code === 0 && /przewodnik/i.test(JSON.stringify(drive("look", "t-desk").out)), o.raw);

  // dokument otwiera się w Czytaniu — jak człowiek: „Edytuj”, potem stuknięcie w tekst
  const ed = drive("tap", "t-desk", "Edytuj");
  check("tap po tekście przycisku (Edytuj)", ed.code === 0, ed.raw);
  check("tap po tekście: czas reakcji bez narzutu wyszukiwania (< 120 ms)", typeof ed.out?.feel?.reactMs === "number" && ed.out.feel.reactMs < 120, JSON.stringify(ed.out?.feel));
  check("w Edycji dokument jest polem edycji", (drive("look", "t-desk").out?.elements || []).some((x) => x.role === "editable"), "");
  drive("tap", "t-desk", "--xy=700,300");
  const caret = drive("look", "t-desk").out?.caret;
  check("look: kursor (akapit z tekstem) po stuknięciu w tekst", !!caret && typeof caret.text === "string" && caret.text.length > 0, JSON.stringify(caret));
  const endKey = drive("key", "t-desk", "End");
  check("feel: klawisz przesuwający tylko kursor nie jest martwy", endKey.out?.feel?.dead === false, JSON.stringify(endKey.out?.feel));

  // natywna lista wyboru (krój, styl akapitu)
  const sels = (drive("look", "t-desk").out?.elements || []).filter((x) => x.role === "select");
  const s0 = sels[0];
  const wrong = s0 ? drive("select", "t-desk", `--el=${s0.i}`, "Takiej opcji nie ma 999") : null;
  check("select: zła opcja = błąd z listą opcji", !!wrong && wrong.code !== 0 && Array.isArray(wrong.out?.options) && wrong.out.options.length > 0, wrong?.raw || "brak <select> na ekranie");
  const good = s0 && wrong?.out?.options ? drive("select", "t-desk", `--el=${s0.i}`, wrong.out.options[Math.min(1, wrong.out.options.length - 1)]) : null;
  check("select: wybór opcji po etykiecie", good?.code === 0, good?.raw || "");
  drive("key", "t-desk", "Escape");
  drive("tap", "t-desk", "--xy=700,300");

  const ps = drive("paste", "t-desk", "--html=<table><tr><td>WKLEJKA</td><td>123</td></tr></table>", "--text=WKLEJKA\t123");
  await sleep(1200);
  check("paste: tabela z HTML w dokumencie (text)", ps.code === 0 && /WKLEJKA/.test(drive("text", "t-desk").out?.text || ""), ps.raw);
  drive("type", "t-desk", "x");
  drive("key", "t-desk", "Control+s");
  let dl = [];
  for (let i = 0; i < 10 && !dl.length; i++) { await sleep(500); dl = drive("look", "t-desk").out?.downloads || []; }
  check("zapis: pobrany plik .docx", dl.length >= 1 && /\.docx$/.test(dl[0]) && fs.existsSync(dl[0]), JSON.stringify(dl));
  // drugi zapis tym samym „uchwytem”: plik nie może zawierać pierwszego zapisu
  drive("type", "t-desk", "y");
  drive("key", "t-desk", "Control+s");
  let dl2 = dl;
  for (let i = 0; i < 10 && dl2.length < 2; i++) { await sleep(500); dl2 = drive("look", "t-desk").out?.downloads || []; }
  if (dl2.length >= 2) {
    const s1 = fs.statSync(dl2[0]).size, s2 = fs.statSync(dl2[1]).size;
    const eocd = fs.readFileSync(dl2[1]).toString("latin1").split("PK\x05\x06").length - 1;
    check("drugi zapis: jeden plik .docx, nie sklejony z pierwszym", s2 < s1 * 1.5 && eocd === 1, `${s1} → ${s2} B, EOCD ×${eocd}`);
  } else check("drugi zapis: jeden plik .docx, nie sklejony z pierwszym", false, JSON.stringify(dl2));
  if (dl.length) {
    drive("errors", "t-desk");
    const re = drive("open-file", "t-desk", dl[0]);
    await sleep(2000);
    const er = drive("errors", "t-desk");
    check("ponowne otwarcie zapisanego pliku bez błędów", re.code === 0 && Array.isArray(er.out?.errors) && er.out.errors.length === 0, re.raw + JSON.stringify(er.out));
    check("po ponownym otwarciu wklejka jest w tekście", /WKLEJKA/.test(drive("text", "t-desk").out?.text || ""), "");
  }

  // telefon: przewijanie palcem (mouse.wheel nie działa w mobilnym WebKit)
  drive("open-file", "t-iphone", "docs/samples/przewodnik.docx");
  await sleep(2500);
  // przewijanie w bok (rząd skrótów sekcji nad dokumentem, rząd kart w arkuszu telefonu)
  const chip = (drive("look", "t-iphone").out?.elements || []).find((x) => x.y < 140 && x.y > 90 && x.role === "button");
  const sx = chip ? drive("scroll", "t-iphone", "0", "--dx=250", `--xy=${chip.x},${chip.y}`) : null;
  check("scroll w bok (--dx) przewija rząd", sx?.code === 0 && sx.out?.movedX > 0, sx?.raw || "brak rzędu skrótów na ekranie");
  const sc = drive("scroll", "t-iphone", "600", "--xy=195,500");
  check("scroll na telefonie przewija (moved > 0)", sc.code === 0 && sc.out?.moved > 0, sc.raw);

  // odczucie: martwe stuknięcie (nic się nie zmienia), reakcja przycisku z czasem, podsumowanie
  drive("tap", "t-desk", "--xy=1420,600"); // pierwsze zdejmuje kursor z dokumentu — to widoczna reakcja
  const dead = drive("tap", "t-desk", "--xy=1420,600");
  check("feel: stuknięcie w puste miejsce = martwe", dead.out?.feel?.dead === true, dead.raw);
  const live = drive("tap", "t-desk", "Plik");
  check("feel: przycisk reaguje, czas reakcji w ms", live.out?.feel?.dead === false && typeof live.out.feel.reactMs === "number", live.raw);
  drive("key", "t-desk", "Escape");
  const fs1 = drive("feel", "t-desk");
  check("feel: podsumowanie (czynności, martwe, mediana reakcji)", fs1.out?.actions >= 3 && fs1.out.dead >= 1 && typeof fs1.out.reactMedianMs === "number", fs1.raw);

  // zamknięcie karty i ponowne otwarcie apki w tej samej przeglądarce (pamięć przeglądarki zostaje)
  const ro = drive("reopen", "t-iphone");
  check("reopen: nowa karta z apką", ro.code === 0 && drive("look", "t-iphone").code === 0, ro.raw);
  check("tap --xy na telefonie", drive("tap", "t-iphone", "--xy=195,400").code === 0, "");
  check("errors zwraca tablicę", Array.isArray(drive("errors", "t-iphone").out?.errors), "");

  // restart pod tym samym imieniem nie nadpisuje starych zrzutów
  const lastShot = drive("look", "t-iphone").out?.shot;
  drive("stop", "t-iphone");
  drive("start", "t-iphone", "--device=iphone", "--run=selftest");
  const nextShot = drive("look", "t-iphone").out?.shot;
  check("po restarcie zrzuty numerowane dalej", !!lastShot && !!nextShot && nextShot > lastShot, `${lastShot} → ${nextShot}`);

  // zawieszony serwer testera (np. przeglądarka nie odpowiada): stop i tak go zamyka
  drive("start", "t-hang", "--device=desktop", "--run=selftest");
  const hp = sess("t-hang")?.pid;
  if (hp) process.kill(hp, "SIGSTOP");
  const hs = drive("stop", "t-hang");
  await sleep(500);
  check("stop zawieszonego testera: proces zabity, sesja usunięta", !!hp && hs.code === 0 && !alive(hp) && !sess("t-hang"), `${hs.raw} pid ${hp} żyje: ${hp && alive(hp)}`);

  const myPids = ["t-iphone", "t-desk"].map((n) => sess(n)?.pid).filter(Boolean);
  stopMine();
  await sleep(1500);
  check("stop: procesy zamknięte", myPids.every((p) => !alive(p)), JSON.stringify(myPids.filter(alive)));
  check("stop: brak moich sesji", MINE.every((n) => !sess(n)), "");

  console.log(failed ? `\n${failed} nie przeszło` : "\nwszystko przeszło");
  process.exit(failed ? 1 : 0);
}
run().catch((err) => { console.error(err); stopMine(); process.exit(1); });
