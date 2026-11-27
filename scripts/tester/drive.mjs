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
//   open-file <imię> <ścieżka>   select <imię> --el=<i> "<opcja>"   paste <imię> --text=… [--html=…]
//   text <imię>   reopen <imię>   feel <imię>   wait <imię> <ms≤5000>   errors <imię>   stop <imię>   stop-all
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

const out = (obj, code = 0) => { process.stdout.write(`${JSON.stringify(obj)}\n`, () => process.exit(code)); };
const fail = (msg) => out({ error: msg }, 1);
const sessFile = (name) => path.join(SESS, `${name}.json`);
const alive = (pid) => { try { process.kill(pid, 0); return true; } catch (_) { return false; } };
const flag = (args, k) => args.find((a) => a.startsWith(`--${k}=`))?.slice(k.length + 3);

// Śmieci w katalogu (.DS_Store z Findera, plik w trakcie zapisu, sesja usunięta między odczytami)
// nie mogą wywrócić poleceń wszystkich testerów naraz.
function liveSessions() {
  if (!fs.existsSync(SESS)) return [];
  return fs.readdirSync(SESS).filter((f) => f.endsWith(".json")).flatMap((f) => {
    let s;
    try { s = JSON.parse(fs.readFileSync(path.join(SESS, f), "utf8")); } catch (_) { return []; }
    if (s && alive(s.pid)) return [s];
    fs.rmSync(path.join(SESS, f), { force: true }); // po padniętym serwerze
    return [];
  });
}
function writeSession(name, data) {
  const tmp = `${sessFile(name)}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(data));
  fs.renameSync(tmp, sessFile(name));
}
async function killHard(s) {
  for (let i = 0; i < 25 && alive(s.pid); i++) await new Promise((r) => setTimeout(r, 200));
  if (alive(s.pid)) { try { process.kill(s.pid, "SIGKILL"); } catch (_) { /* już nie żyje */ } }
  fs.rmSync(sessFile(s.name), { force: true });
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
    await Promise.all(all.map((s) => post(s.port, { cmd: "stop" }, 3000).catch(() => {})));
    await Promise.all(all.map(killHard));
    return out({ stopped: all.map((s) => s.name) });
  }
  if (!cmd || !name) return fail("użycie: drive.mjs <polecenie> <imię> … (opis na górze pliku)");
  if (!/^[\w-]{1,40}$/.test(name)) return fail("imię: litery, cyfry, - i _");

  if (cmd === "start") {
    const device = flag(rest, "device") || "iphone";
    if (!DEVICES[device]) return fail(`nieznane urządzenie: ${device} (${Object.keys(DEVICES).join(", ")})`);
    const d = new Date();
    const run = flag(rest, "run") || `${d.toISOString().slice(0, 10)}-${String(d.getHours()).padStart(2, "0")}${String(d.getMinutes()).padStart(2, "0")}`;
    if (!/^[\w-]{1,60}$/.test(run)) return fail("--run: litery, cyfry, - i _");
    const live = liveSessions();
    if (live.some((s) => s.name === name)) return fail(`tester ${name} już działa — najpierw stop ${name}`);
    if (live.length >= MAX) return fail(`działa już ${live.length} przeglądarek (limit ${MAX}) — zatrzymaj któregoś testera`);
    try {
      const r = await fetch(APP_URL, { signal: AbortSignal.timeout(5000) });
      if (!r.ok) throw new Error(String(r.status));
    } catch (_) { return fail(`apka nie odpowiada pod ${APP_URL} — uruchom: python3 scripts/test-server.py 7823`); }
    const dir = path.join(BASE, run, name);
    fs.mkdirSync(dir, { recursive: true });
    fs.mkdirSync(SESS, { recursive: true });
    // rezerwacja imienia, zanim przeglądarka wstanie (drugi start w tym czasie = „już działa”)
    try { fs.writeFileSync(sessFile(name), JSON.stringify({ name, pid: process.pid, ready: false }), { flag: "wx" }); } catch (_) { return fail(`tester ${name} już działa — najpierw stop ${name}`); }
    const log = fs.openSync(path.join(dir, "serve.log"), "a");
    const child = spawn(process.execPath, [fileURLToPath(import.meta.url), "serve", name, device, dir], { cwd: ROOT, detached: true, stdio: ["ignore", log, log] });
    child.unref();
    writeSession(name, { name, pid: child.pid, ready: false });
    for (let i = 0; i < 240; i++) {
      await new Promise((r) => setTimeout(r, 250));
      let s = null;
      try { s = JSON.parse(fs.readFileSync(sessFile(name), "utf8")); } catch (_) { /* w trakcie zapisu */ }
      if (s?.ready) return out({ name, device, engine: DEVICES[device].engine, dir, port: s.port });
      if (!alive(/** @type {number} */ (child.pid))) break;
    }
    try { process.kill(/** @type {number} */ (child.pid), "SIGKILL"); } catch (_) { /* nie żyje */ }
    fs.rmSync(sessFile(name), { force: true });
    return fail(`przeglądarka nie wystartowała — zob. ${path.join(dir, "serve.log")}`);
  }

  const s = liveSessions().find((x) => x.name === name);
  if (!s) return fail(`tester ${name} nie działa — najpierw start ${name}`);
  if (cmd === "stop") {
    // zawieszona przeglądarka nie odpowie — wtedy i tak zamykamy proces
    await post(s.port, { cmd: "stop" }, 3000).catch(() => {});
    await killHard(s);
    return out({ ok: true });
  }
  let res;
  try { res = await post(s.port, { cmd, args: rest }); } catch (e) { return fail(String(e.message || e)); }
  return out(res, res && res.error ? 1 : 0);
}

// ---------------------------------------------------------------- serwer testera
async function serve(name, device, dir) {
  const pw = await import("playwright");
  const spec = DEVICES[device];
  const { engine, ...ctxOpts } = spec;
  const browser = await pw[engine].launch({ headless: true });
  // koniec: zamknięcie przeglądarki, ale najwyżej 5 s (zawieszona nie trzyma procesu)
  const bye = async () => {
    setTimeout(() => process.exit(0), 5000).unref();
    await browser.close().catch(() => {});
    process.exit(0);
  };
  browser.on("disconnected", () => process.exit(1)); // padła przeglądarka — sesja zniknie przy następnym liveSessions
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
          createWritable: async () => {
            parts.length = 0; // nowy zapis tym samym uchwytem = nowy plik (dawniej doklejał się do poprzedniego)
            return {
            write: async (c) => { parts.push(c && c.type === "write" ? c.data : c); },
            close: async () => {
              const a = document.createElement("a");
              a.href = URL.createObjectURL(new Blob(parts));
              a.download = name;
              a.click();
            },
            };
          },
        };
      };
    }
  });
  let page = await context.newPage();
  const errors = [];
  const downloads = [];
  let chooser = null;
  let lastEls = [];
  // restart pod tym samym imieniem: numeracja zrzutów dalej (stare zostają w raporcie)
  let shotN = Math.max(0, ...fs.readdirSync(dir).map((f) => Number(/^shot-(\d+)\.png$/.exec(f)?.[1] || 0)));
  const attach = (pg) => {
  pg.on("pageerror", (e) => errors.push(`pageerror: ${e.message}`));
  pg.on("console", (m) => { if (m.type() === "error") errors.push(`console: ${m.text()}`); });
  pg.on("filechooser", (fc) => { chooser = fc; });
  pg.on("download", async (d) => {
    const dl = path.join(dir, "pobrane");
    fs.mkdirSync(dl, { recursive: true });
    let p = path.join(dl, d.suggestedFilename());
    for (let i = 2; fs.existsSync(p); i++) p = path.join(dl, d.suggestedFilename().replace(/(\.[^.]+)?$/, `-${i}$1`));
    try { await d.saveAs(p); downloads.push(p); } catch (e) { errors.push(`zapis pobranego pliku: ${e?.message || e}`); }
  });
  };
  attach(page);
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
        // najwyższe widoczne okno: to, które jest pod swoim środkiem (potwierdzenie nad panelem Plik)
        const dlg = [...document.querySelectorAll("dialog[open], [role=dialog]:not([hidden]), [role=alertdialog]:not([hidden])")].reverse().find((d) => {
          const r = d.getBoundingClientRect();
          if (r.height <= 0 || r.width <= 0) return false;
          const h = document.elementFromPoint(Math.min(vw - 1, r.left + r.width / 2), Math.min(vh - 1, Math.max(0, r.top + r.height / 2)));
          return !!h && d.contains(h);
        });
        // otwarte okno: tylko jego elementy (reszta i tak pod spodem)
        const list = (dlg ? els.filter((e) => dlg.contains(e.el)) : els).slice(0, 120).map(({ el, ...e }) => e);
        // kursor: akapit, w którym stoi, i komórka tabeli (wiersz/kolumna od 1)
        let caret = null;
        const cur = getSelection();
        const an = cur && cur.rangeCount ? cur.anchorNode : null;
        const at = an && (an.nodeType === 1 ? /** @type {Element} */ (an) : an.parentElement);
        if (at && at.closest("section.docx")) {
          const para = at.closest("p, li, h1, h2, h3, h4, h5, h6") || at;
          const td = at.closest("td, th");
          const tr = td?.parentElement;
          caret = {
            text: (para.textContent || "").replace(/\s+/g, " ").trim().slice(0, 100),
            offset: cur.anchorOffset,
            collapsed: cur.isCollapsed,
            cell: td && tr ? { row: [...tr.parentElement.children].indexOf(tr) + 1, col: [...tr.children].indexOf(td) + 1 } : null,
          };
        }
        return { url: location.href, title: document.title, caret, dialog: dlg ? (dlg.getAttribute("aria-label") || dlg.querySelector("h1,h2,h3")?.textContent || "okno").trim().slice(0, 80) : null, elements: list, viewport: { width: vw, height: vh } };
      });
      lastEls = info.elements;
      return { shot, ...info, elements: info.elements.map((e, i) => ({ i, ...e })), downloads: [...downloads], fileChooserOpen: !!chooser };
    },
    // gdzie stuknąć — liczone PRZED pomiarem odczucia (szukanie po tekście robi zrzut)
    async resolveTap(args) {
      const el = flag(args, "el"), p = xy(flag(args, "xy"));
      if (el != null) {
        const e = lastEls[Number(el)];
        return e ? { at: [e.x, e.y] } : { error: `brak elementu ${el} w ostatnim look (jest ${lastEls.length})` };
      }
      if (p) return { at: p };
      const text = args.filter((a) => !a.startsWith("--")).join(" ").trim();
      if (!text) return { error: "podaj --el=, --xy= albo tekst" };
      const cur = (await commands.look()).elements;
      const low = text.toLowerCase();
      let hits = cur.filter((e) => e.name.toLowerCase() === low);
      if (!hits.length) hits = cur.filter((e) => e.name.toLowerCase().includes(low));
      if (!hits.length) return { error: `nie widzę „${text}” na ekranie`, visible: cur.slice(0, 40).map((e) => `${e.i}: ${e.name}`) };
      if (hits.length > 1) return { error: `„${text}” pasuje do ${hits.length} elementów — użyj --el=`, candidates: hits.map((e) => `${e.i}: ${e.role} „${e.name}” @${e.x},${e.y}`) };
      return { at: [hits[0].x, hits[0].y] };
    },
    async tap(args, at) {
      const r = at ? { at } : await commands.resolveTap(args);
      if (r.error) return r;
      await tapAt(r.at[0], r.at[1]);
      await page.waitForTimeout(250);
      return { ok: true };
    },
    async type(args) { await page.keyboard.type(args.join(" "), { delay: 15 }); return { ok: true }; },
    async key(args) { await page.keyboard.press(args[0]); await page.waitForTimeout(200); return { ok: true }; },
    async drag(args) {
      const a = xy(args[0]), b = xy(args[1]);
      if (!a || !b) return { error: "drag x1,y1 x2,y2" };
      // także na dotyku myszą: przeciągnięcie palcem przez zdarzenia nie działa w WebKit (jak w testach)
      await page.mouse.move(a[0], a[1]); await page.mouse.down(); await page.mouse.move(b[0], b[1], { steps: 12 }); await page.mouse.up();
      return { ok: true };
    },
    // Przewija to, co jest pod punktem (najbliższy przewijany przodek, jak pod palcem/kółkiem).
    // Mobilny WebKit w Playwrighcie nie ma kółka ani gestu przesunięcia — tam przewijamy wprost.
    async scroll(args) {
      const dy = Number(args[0]) || 400, vp = page.viewportSize() || { width: 400, height: 800 };
      const p = xy(flag(args, "xy")) || [vp.width / 2, vp.height / 2];
      const where = (pt) => page.evaluate(([x, y]) => {
        let el = document.elementFromPoint(x, y);
        while (el && el !== document.documentElement) {
          const st = getComputedStyle(el);
          if (/(auto|scroll)/.test(st.overflowY) && el.scrollHeight > el.clientHeight + 1) break;
          el = el.parentElement;
        }
        const s = el && el !== document.documentElement ? el : document.scrollingElement || document.documentElement;
        return s.scrollTop;
      }, pt);
      const before = await where(p);
      if (touch) {
        await page.evaluate(([x, y, d]) => {
          let el = document.elementFromPoint(x, y);
          while (el && el !== document.documentElement) {
            const st = getComputedStyle(el);
            if (/(auto|scroll)/.test(st.overflowY) && el.scrollHeight > el.clientHeight + 1) break;
            el = el.parentElement;
          }
          (el && el !== document.documentElement ? el : document.scrollingElement || document.documentElement).scrollBy({ top: d });
        }, [p[0], p[1], dy]);
      } else {
        await page.mouse.move(p[0], p[1]);
        await page.mouse.wheel(0, dy);
      }
      await page.waitForTimeout(350);
      return { ok: true, moved: Math.round((await where(p)) - before) };
    },
    async "open-file"(args) {
      const f = path.resolve(ROOT, args.join(" "));
      if (!fs.existsSync(f)) return { error: `brak pliku ${f}` };
      if (chooser) { const c = chooser; chooser = null; await c.setFiles(f); } else await page.setInputFiles("#fileInput", f);
      await page.waitForTimeout(500);
      return { ok: true };
    },
    // wklejenie jak z innej apki (Excel, strona WWW, mail): zdarzenie paste w miejscu kursora,
    // tak samo jak w testach wklejania — prawdziwy schowek systemu jest poza zasięgiem
    async paste(args) {
      const text = flag(args, "text"), html = flag(args, "html");
      if (text == null && html == null) return { error: "paste --text=… [--html=…]" };
      const ok = await page.evaluate(({ text, html }) => {
        const sel = getSelection();
        const n = sel && sel.rangeCount ? sel.anchorNode : null;
        const target = (n && (n.nodeType === 1 ? n : n.parentElement)) || document.activeElement;
        if (!target || target === document.body) return false;
        const dt = new DataTransfer();
        if (text != null) dt.setData("text/plain", text);
        if (html != null) dt.setData("text/html", html);
        target.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
        return true;
      }, { text, html });
      await page.waitForTimeout(400);
      return ok ? { ok: true } : { error: "brak kursora — najpierw stuknij w miejsce wklejenia" };
    },
    // natywna lista wyboru (krój, rozmiar, styl akapitu): okno systemu jest poza stroną
    async select(args) {
      const el = flag(args, "el"), label = args.filter((a) => !a.startsWith("--")).join(" ").trim();
      const e = lastEls[Number(el)];
      if (!e || e.role !== "select") return { error: "select --el=<i> wskazuje element „select” z ostatniego look" };
      const handle = await page.evaluateHandle(([x, y]) => document.elementFromPoint(x, y)?.closest("select"), [e.x, e.y]);
      const sel = handle.asElement();
      if (!sel) return { error: "nie widzę tej listy w tym miejscu — zrób look" };
      const options = await sel.evaluate((s) => [.../** @type {HTMLSelectElement} */ (s).options].map((o) => o.label || o.text));
      if (!options.includes(label)) return { error: `brak opcji „${label}”`, options: options.slice(0, 60) };
      await sel.selectOption({ label });
      return { ok: true };
    },
    // zamknięcie karty i otwarcie apki od nowa w tej samej przeglądarce (pamięć przeglądarki,
    // np. szkice, zostaje — jak po zamknięciu i ponownym otwarciu karty w Safari)
    async reopen() {
      await page.close().catch(() => {});
      page = await context.newPage();
      attach(page);
      await page.goto(APP_URL, { waitUntil: "load" });
      await page.evaluate(() => document.getElementById("heroSplash")?.remove());
      chooser = null;
      lastEls = [];
      return { ok: true };
    },
    async text() {
      const t = await page.evaluate(() => [...document.querySelectorAll("section.docx")].map((s) => /** @type {HTMLElement} */ (s).innerText).join("\n\n"));
      return { text: t.replace(/\n{3,}/g, "\n\n").slice(0, 6000), length: t.length };
    },
    async wait(args) { await page.waitForTimeout(Math.min(5000, Math.max(0, Number(args[0]) || 1000))); return { ok: true }; },
    async errors() { const e = errors.splice(0); return { errors: e }; },
    async stop() { setTimeout(bye, 50); return { ok: true }; },
  };

  // ---- odczucie: każda czynność mierzona. Martwe = ekran po 700 ms taki sam jak przed (zrzuty).
  // Reakcja = czas do pierwszej zmiany DOM, zaznaczenia albo przewinięcia. Skok = treść pod
  // środkiem ekranu przesunęła się o > 80 px po stuknięciu lub pisaniu (nie po przewijaniu).
  // zrzut do porównania: bez kursora i dymków (znikają same z czasu — to nie reakcja na czynność)
  const CMP_SHOT = { scale: /** @type {"css"} */ ("css"), caret: /** @type {"hide"} */ ("hide"), style: ".toast{visibility:hidden!important}" };
  const MEASURED = new Set(["tap", "type", "key", "paste", "drag", "scroll"]);
  const feelLog = path.join(dir, "feel.jsonl");
  const feelBefore = () => page.evaluate(() => {
    const st = { t0: performance.now(), first: /** @type {number|null} */ (null), sel: false };
    const mark = () => { if (st.first == null) st.first = performance.now(); };
    const selMark = () => { st.sel = true; mark(); };
    const mo = new MutationObserver(mark);
    mo.observe(document, { subtree: true, childList: true, attributes: true, characterData: true });
    document.addEventListener("selectionchange", selMark, true);
    addEventListener("scroll", mark, true);
    const el = document.elementFromPoint(innerWidth / 2, innerHeight / 2);
    /** @type {any} */ (window).__dwbFeel = { st, mo, mark, selMark, el, top: el ? el.getBoundingClientRect().top : null };
  }).catch(() => {});
  const feelAfter = async (cmd) => {
    await page.waitForTimeout(700);
    return page.evaluate((cmd) => {
      const f = /** @type {any} */ (window).__dwbFeel;
      if (!f) return null; // strona przeładowana w trakcie — to reakcja
      f.mo.disconnect();
      document.removeEventListener("selectionchange", f.selMark, true);
      removeEventListener("scroll", f.mark, true);
      delete /** @type {any} */ (window).__dwbFeel;
      const reactMs = f.st.first == null ? null : Math.max(0, Math.round(f.st.first - f.st.t0));
      let jumpPx = 0;
      if ((cmd === "tap" || cmd === "type") && f.el && f.el.isConnected && f.top != null) {
        const d = Math.round(f.el.getBoundingClientRect().top - f.top);
        if (Math.abs(d) > 80) jumpPx = d;
      }
      return { reactMs, dead: reactMs == null, jumpPx, selMoved: f.st.sel };
    }, cmd).catch(() => ({ reactMs: null, dead: false, jumpPx: 0, reloaded: true }));
  };
  commands.feel = async () => {
    const rows = fs.existsSync(feelLog) ? fs.readFileSync(feelLog, "utf8").trim().split("\n").filter(Boolean).map((l) => JSON.parse(l)) : [];
    const react = rows.map((r) => r.reactMs).filter((v) => typeof v === "number").sort((a, b) => a - b);
    const q = (p) => (react.length ? react[Math.min(react.length - 1, Math.floor(p * react.length))] : null);
    return {
      actions: rows.length,
      dead: rows.filter((r) => r.dead).length,
      deadExamples: rows.filter((r) => r.dead).slice(-5).map((r) => `${r.cmd} ${r.args.join(" ")}`.slice(0, 80)),
      reactMedianMs: q(0.5), reactP90Ms: q(0.9),
      slow: rows.filter((r) => r.reactMs > 300).length,
      jumps: rows.filter((r) => r.jumpPx).length,
    };
  };

  let idle = setTimeout(() => commands.stop(), IDLE_MS);
  const handle = async (body) => {
    clearTimeout(idle);
    idle = setTimeout(() => commands.stop(), IDLE_MS);
    let result;
    try {
      const { cmd, args = [] } = JSON.parse(body || "{}");
      const fn = commands[cmd];
      if (!fn) result = { error: `nieznane polecenie: ${cmd}` };
      else if (MEASURED.has(cmd)) {
        let at = null;
        if (cmd === "tap") {
          const r = await commands.resolveTap(args);
          if (r.error) return r;
          at = r.at;
        }
        const chooserBefore = chooser;
        const shot0 = await page.screenshot(CMP_SHOT).catch(() => null);
        await feelBefore();
        result = await fn(args, at);
        const feel = await feelAfter(cmd);
        // martwe = ekran po czynności identyczny jak przed, kursor/zaznaczenie stoi i nie otwarło
        // się okno wyboru pliku (to widzi człowiek; zmiany DOM bez widocznego skutku się nie liczą)
        if (feel && shot0) {
          const shot1 = await page.screenshot(CMP_SHOT).catch(() => null);
          feel.dead = !!shot1 && shot0.equals(shot1) && !feel.selMoved && chooser === chooserBefore;
          if (feel.dead) feel.reactMs = null;
        }
        if (feel) delete feel.selMoved;
        if (!result.error && feel) {
          result.feel = feel;
          fs.appendFileSync(feelLog, `${JSON.stringify({ t: Date.now(), cmd, args, ...feel })}\n`);
        }
      } else result = await fn(args);
    } catch (e) { result = { error: String(e?.message || e) }; }
    return result;
  };
  // jedno polecenie naraz (agent potrafi wysłać kilka wywołań równolegle)
  let chain = Promise.resolve();
  const server = http.createServer((req, res) => {
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", () => {
      chain = chain.then(() => handle(body)).then((result) => {
        res.writeHead(200, { "content-type": "application/json" });
        res.end(JSON.stringify(result));
      });
    });
  });
  server.listen(0, "127.0.0.1", () => {
    const addr = server.address();
    const port = typeof addr === "object" && addr ? addr.port : 0;
    writeSession(name, { name, device, dir, port, pid: process.pid, ready: true });
  });
  process.on("SIGTERM", bye);
  process.on("SIGINT", bye);
}

const argv = process.argv.slice(2);
if (argv[0] === "serve") serve(argv[1], argv[2], argv[3]).catch((e) => { console.error(e); fs.rmSync(sessFile(argv[1]), { force: true }); process.exit(1); });
else client(argv);
