// @ts-check
// drive.mjs — sterownik przeglądarki dla testerów AI (agent .claude/agents/tester-apki.md).
// Każdy tester ma WŁASNĄ przeglądarkę: proces w tle (tryb „serve”) trzyma stronę Playwrighta,
// a polecenia z Basha idą do niego przez lokalny HTTP. Wynik każdego polecenia: jeden JSON na
// stdout; błąd = kod ≠ 0 i { "error": "…" }.
//
//   node scripts/tester/drive.mjs start <imię> --device=iphone|iphone-landscape|ipad|desktop [--run=<przebieg>]
//   look <imię>                     zrzut (piksele CSS) + widoczne elementy do stuknięcia + pobrane pliki
//   tap <imię> --el=<i> | --xy=x,y | "<tekst>"
//   type <imię> "<tekst>"   key <imię> <Klawisz>   drag <imię> x1,y1 x2,y2   scroll <imię> <dy> [--xy=x,y]
//   open-file <imię> <ścieżka>      wait <imię> <ms≤5000>   errors <imię>   stop <imię>   stop-all
//
// Obciążenie: najwyżej MAX_TESTERS (domyślnie 6) przeglądarek naraz, bez okna; serwer testera
// kończy się sam po 30 min bez poleceń.

import { spawn } from "node:child_process";
import fs from "node:fs";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
const BASE = path.join(ROOT, "output", "testerzy");
const SESS = path.join(BASE, ".sessions");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:7823/";
const MAX = Number(process.env.MAX_TESTERS || 6);
const IDLE_MS = 30 * 60 * 1000;

const DEVICES = {
  iphone: { engine: "webkit", viewport: { width: 390, height: 844 }, deviceScaleFactor: 3, isMobile: true, hasTouch: true },
  "iphone-landscape": { engine: "webkit", viewport: { width: 844, height: 390 }, deviceScaleFactor: 3, isMobile: true, hasTouch: true },
  ipad: { engine: "webkit", viewport: { width: 820, height: 1180 }, deviceScaleFactor: 2, isMobile: true, hasTouch: true },
  desktop: { engine: "chromium", viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1, isMobile: false, hasTouch: false },
};

const out = (obj, code = 0) => { process.stdout.write(`${JSON.stringify(obj)}\n`); process.exit(code); };
const fail = (msg) => out({ error: msg }, 1);
const sessFile = (name) => path.join(SESS, `${name}.json`);
const alive = (pid) => { try { process.kill(pid, 0); return true; } catch (_) { return false; } };
const flag = (args, k) => args.find((a) => a.startsWith(`--${k}=`))?.slice(k.length + 3);

function liveSessions() {
  if (!fs.existsSync(SESS)) return [];
  return fs.readdirSync(SESS).flatMap((f) => {
    const s = JSON.parse(fs.readFileSync(path.join(SESS, f), "utf8"));
    if (alive(s.pid)) return [s];
    fs.rmSync(path.join(SESS, f), { force: true }); // po padniętym serwerze
    return [];
  });
}

function post(port, body, timeout = 45000) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(body);
    const req = http.request({ host: "127.0.0.1", port, path: "/cmd", method: "POST", headers: { "content-type": "application/json", "content-length": Buffer.byteLength(data) }, timeout }, (res) => {
      let s = "";
      res.on("data", (c) => (s += c));
      res.on("end", () => { try { resolve(JSON.parse(s)); } catch (e) { reject(e); } });
    });
    req.on("timeout", () => req.destroy(new Error("brak odpowiedzi przeglądarki testera")));
    req.on("error", reject);
    req.end(data);
  });
}

// ---------------------------------------------------------------- klient
async function client(argv) {
  const [cmd, name, ...rest] = argv;
  if (cmd === "stop-all") {
    const all = liveSessions();
    for (const s of all) { try { await post(s.port, { cmd: "stop" }, 5000); } catch (_) { /* niżej kill */ } }
    for (const s of all) { if (alive(s.pid)) try { process.kill(s.pid, "SIGKILL"); } catch (_) { /* już nie żyje */ } fs.rmSync(sessFile(s.name), { force: true }); }
    return out({ stopped: all.map((s) => s.name) });
  }
  if (!cmd || !name) return fail("użycie: drive.mjs <polecenie> <imię> … (opis na górze pliku)");
  if (!/^[\w-]{1,40}$/.test(name)) return fail("imię: litery, cyfry, - i _");

  if (cmd === "start") {
    const device = flag(rest, "device") || "iphone";
    if (!DEVICES[device]) return fail(`nieznane urządzenie: ${device} (${Object.keys(DEVICES).join(", ")})`);
    const live = liveSessions();
    if (live.some((s) => s.name === name)) return fail(`tester ${name} już działa — najpierw stop ${name}`);
    if (live.length >= MAX) return fail(`działa już ${live.length} przeglądarek (limit ${MAX}) — zatrzymaj któregoś testera`);
    try { await fetch(APP_URL); } catch (_) { return fail(`apka nie odpowiada pod ${APP_URL} — uruchom: python3 scripts/test-server.py 7823`); }
    const d = new Date();
    const run = flag(rest, "run") || `${d.toISOString().slice(0, 10)}-${String(d.getHours()).padStart(2, "0")}${String(d.getMinutes()).padStart(2, "0")}`;
    const dir = path.join(BASE, run, name);
    fs.mkdirSync(dir, { recursive: true });
    fs.mkdirSync(SESS, { recursive: true });
    const log = fs.openSync(path.join(dir, "serve.log"), "a");
    const child = spawn(process.execPath, [fileURLToPath(import.meta.url), "serve", name, device, dir], { cwd: ROOT, detached: true, stdio: ["ignore", log, log] });
    child.unref();
    for (let i = 0; i < 80; i++) {
      await new Promise((r) => setTimeout(r, 250));
      if (fs.existsSync(sessFile(name))) {
        const s = JSON.parse(fs.readFileSync(sessFile(name), "utf8"));
        if (s.ready) return out({ name, device, engine: DEVICES[device].engine, dir, port: s.port });
      }
      if (!alive(/** @type {number} */ (child.pid))) break;
    }
    try { process.kill(/** @type {number} */ (child.pid), "SIGKILL"); } catch (_) { /* nie żyje */ }
    fs.rmSync(sessFile(name), { force: true });
    return fail(`przeglądarka nie wystartowała — zob. ${path.join(dir, "serve.log")}`);
  }

  const s = liveSessions().find((x) => x.name === name);
  if (!s) return fail(`tester ${name} nie działa — najpierw start ${name}`);
  let res;
  try { res = await post(s.port, { cmd, args: rest }); } catch (e) { return fail(String(e.message || e)); }
  if (cmd === "stop") fs.rmSync(sessFile(name), { force: true });
  return out(res, res && res.error ? 1 : 0);
}

// ---------------------------------------------------------------- serwer testera
async function serve(name, device, dir) {
  const pw = await import("playwright");
  const spec = DEVICES[device];
  const { engine, ...ctxOpts } = spec;
  const browser = await pw[engine].launch({ headless: true });
  const context = await browser.newContext({ ...ctxOpts, serviceWorkers: "block", acceptDownloads: true, locale: "pl-PL" });
  await context.addInitScript(() => {
    sessionStorage.setItem("introPlayed", "true");
    // Systemowe okna plików są poza stroną — agent ich nie kliknie. Otwieranie: zwykłe pole pliku
    // (drive.mjs open-file); zapis: okno udawane, plik trafia do pobranych testera.
    try { delete window.showOpenFilePicker; } catch (_) { /* brak */ }
    if ("showSaveFilePicker" in window) {
      window.showSaveFilePicker = async (opts) => {
        const name = opts?.suggestedName || "dokument.docx";
        const parts = [];
        return {
          kind: "file", name,
          queryPermission: async () => "granted", requestPermission: async () => "granted",
          getFile: async () => new File(parts, name),
          createWritable: async () => ({
            write: async (c) => { parts.push(c); },
            close: async () => {
              const a = document.createElement("a");
              a.href = URL.createObjectURL(new Blob(parts));
              a.download = name;
              a.click();
            },
          }),
        };
      };
    }
  });
  const page = await context.newPage();
  const errors = [];
  const downloads = [];
  let chooser = null;
  let lastEls = [];
  let shotN = 0;
  page.on("pageerror", (e) => errors.push(`pageerror: ${e.message}`));
  page.on("console", (m) => { if (m.type() === "error") errors.push(`console: ${m.text()}`); });
  page.on("filechooser", (fc) => { chooser = fc; });
  page.on("download", async (d) => {
    const dl = path.join(dir, "pobrane");
    fs.mkdirSync(dl, { recursive: true });
    let p = path.join(dl, d.suggestedFilename());
    for (let i = 2; fs.existsSync(p); i++) p = path.join(dl, d.suggestedFilename().replace(/(\.[^.]+)?$/, `-${i}$1`));
    await d.saveAs(p);
    downloads.push(p);
  });
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());

  const touch = spec.hasTouch;
  const tapAt = async (x, y) => (touch ? page.touchscreen.tap(x, y) : page.mouse.click(x, y));
  const xy = (s) => { const m = /^(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)$/.exec(s || ""); return m ? [Number(m[1]), Number(m[2])] : null; };

  const commands = {
    async look() {
      await page.waitForTimeout(150);
      const shot = path.join(dir, `shot-${String(++shotN).padStart(3, "0")}.png`);
      await page.screenshot({ path: shot, scale: "css" });
      const info = await page.evaluate(() => {
        const sel = "button, a[href], input:not([type=hidden]), select, textarea, summary, [role=button], [role=menuitem], [role=tab], [role=checkbox], [role=option], [role=switch], [contenteditable=true]";
        const vw = innerWidth, vh = innerHeight;
        const els = [];
        for (const el of document.querySelectorAll(sel)) {
          const r = el.getBoundingClientRect();
          if (r.width < 2 || r.height < 2 || r.bottom < 0 || r.right < 0 || r.top > vh || r.left > vw) continue;
          const st = getComputedStyle(el);
          if (st.visibility === "hidden" || Number(st.opacity) === 0) continue;
          const cx = Math.min(vw - 1, Math.max(0, r.left + r.width / 2)), cy = Math.min(vh - 1, Math.max(0, r.top + r.height / 2));
          const hit = document.elementFromPoint(cx, cy);
          if (!hit || !(hit === el || el.contains(hit) || hit.contains(el))) continue; // zasłonięty
          if (el.closest("[contenteditable=true]") && el.getAttribute("contenteditable") !== "true") continue;
          const raw = el.getAttribute("aria-label") || el.getAttribute("title") || (el.isContentEditable ? el.textContent : el.innerText) || el.getAttribute("placeholder") || /** @type {HTMLInputElement} */ (el).value || "";
          const name = raw.replace(/\s+/g, " ").trim().slice(0, el.isContentEditable ? 80 : 60);
          els.push({ role: el.getAttribute("role") || (el.isContentEditable ? "editable" : el.tagName.toLowerCase()), name, x: Math.round(cx), y: Math.round(cy), w: Math.round(r.width), h: Math.round(r.height), el });
        }
        const dlg = [...document.querySelectorAll("dialog[open], [role=dialog]:not([hidden]), [role=alertdialog]:not([hidden])")].find((d) => d.getBoundingClientRect().height > 0);
        // otwarte okno: tylko jego elementy (reszta i tak pod spodem)
        const list = (dlg ? els.filter((e) => dlg.contains(e.el)) : els).slice(0, 120).map(({ el, ...e }) => e);
        return { url: location.href, title: document.title, dialog: dlg ? (dlg.getAttribute("aria-label") || dlg.querySelector("h1,h2,h3")?.textContent || "okno").trim().slice(0, 80) : null, elements: list, viewport: { width: vw, height: vh } };
      });
      lastEls = info.elements;
      return { shot, ...info, elements: info.elements.map((e, i) => ({ i, ...e })), downloads: [...downloads], fileChooserOpen: !!chooser };
    },
    async tap(args) {
      const el = flag(args, "el"), p = xy(flag(args, "xy"));
      if (el != null) {
        const e = lastEls[Number(el)];
        if (!e) return { error: `brak elementu ${el} w ostatnim look (jest ${lastEls.length})` };
        await tapAt(e.x, e.y);
      } else if (p) await tapAt(p[0], p[1]);
      else {
        const text = args.filter((a) => !a.startsWith("--")).join(" ").trim();
        if (!text) return { error: "podaj --el=, --xy= albo tekst" };
        const cur = (await commands.look()).elements;
        const low = text.toLowerCase();
        let hits = cur.filter((e) => e.name.toLowerCase() === low);
        if (!hits.length) hits = cur.filter((e) => e.name.toLowerCase().includes(low));
        if (!hits.length) return { error: `nie widzę „${text}” na ekranie`, visible: cur.slice(0, 40).map((e) => `${e.i}: ${e.name}`) };
        if (hits.length > 1) return { error: `„${text}” pasuje do ${hits.length} elementów — użyj --el=`, candidates: hits.map((e) => `${e.i}: ${e.role} „${e.name}” @${e.x},${e.y}`) };
        await tapAt(hits[0].x, hits[0].y);
      }
      await page.waitForTimeout(250);
      return { ok: true };
    },
    async type(args) { await page.keyboard.type(args.join(" "), { delay: 15 }); return { ok: true }; },
    async key(args) { await page.keyboard.press(args[0]); await page.waitForTimeout(200); return { ok: true }; },
    async drag(args) {
      const a = xy(args[0]), b = xy(args[1]);
      if (!a || !b) return { error: "drag x1,y1 x2,y2" };
      if (touch) {
        // dotyk: przeciągnięcie palcem przez zdarzenia CDP nie działa w WebKit — przytrzymanie myszą jak w testach
        await page.mouse.move(a[0], a[1]); await page.mouse.down(); await page.mouse.move(b[0], b[1], { steps: 12 }); await page.mouse.up();
      } else { await page.mouse.move(a[0], a[1]); await page.mouse.down(); await page.mouse.move(b[0], b[1], { steps: 12 }); await page.mouse.up(); }
      return { ok: true };
    },
    async scroll(args) {
      const dy = Number(args[0]) || 400, p = xy(flag(args, "xy")) || [page.viewportSize().width / 2, page.viewportSize().height / 2];
      await page.mouse.move(p[0], p[1]); await page.mouse.wheel(0, dy); await page.waitForTimeout(300);
      return { ok: true };
    },
    async "open-file"(args) {
      const f = path.resolve(ROOT, args.join(" "));
      if (!fs.existsSync(f)) return { error: `brak pliku ${f}` };
      if (chooser) { const c = chooser; chooser = null; await c.setFiles(f); } else await page.setInputFiles("#fileInput", f);
      await page.waitForTimeout(500);
      return { ok: true };
    },
    async wait(args) { await page.waitForTimeout(Math.min(5000, Math.max(0, Number(args[0]) || 1000))); return { ok: true }; },
    async errors() { const e = errors.splice(0); return { errors: e }; },
    async stop() { setTimeout(async () => { await browser.close().catch(() => {}); fs.rmSync(sessFile(name), { force: true }); process.exit(0); }, 50); return { ok: true }; },
  };

  let idle = setTimeout(() => commands.stop(), IDLE_MS);
  const server = http.createServer((req, res) => {
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", async () => {
      clearTimeout(idle);
      idle = setTimeout(() => commands.stop(), IDLE_MS);
      let result;
      try {
        const { cmd, args = [] } = JSON.parse(body || "{}");
        const fn = commands[cmd];
        result = fn ? await fn(args) : { error: `nieznane polecenie: ${cmd}` };
      } catch (e) { result = { error: String(e?.message || e) }; }
      res.writeHead(200, { "content-type": "application/json" });
      res.end(JSON.stringify(result));
    });
  });
  server.listen(0, "127.0.0.1", () => {
    const addr = server.address();
    const port = typeof addr === "object" && addr ? addr.port : 0;
    fs.writeFileSync(sessFile(name), JSON.stringify({ name, device, dir, port, pid: process.pid, ready: true }));
  });
  const bye = async () => { await browser.close().catch(() => {}); fs.rmSync(sessFile(name), { force: true }); process.exit(0); };
  process.on("SIGTERM", bye);
  process.on("SIGINT", bye);
}

const argv = process.argv.slice(2);
if (argv[0] === "serve") serve(argv[1], argv[2], argv[3]).catch((e) => { console.error(e); fs.rmSync(sessFile(argv[1]), { force: true }); process.exit(1); });
else client(argv);
