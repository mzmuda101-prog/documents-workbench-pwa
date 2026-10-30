// coverage-preload.cjs — nagrywa, które funkcje aplikacji wykonał test (wpinany przez run-tests.mjs
// w NODE_OPTIONS --require, gdy ustawione TEST_COV_OUT; testów nie trzeba zmieniać).
//
// Po co: `npm test` pomija testy, których nie dotyczy zmiana (scripts/test-impact.mjs). Do tego
// trzeba wiedzieć, co każdy test faktycznie uruchomił. Chromium daje to tanio: pokrycie V8 na
// poziomie FUNKCJI (callCount, bez bloków — prawie bez narzutu). WebKit tego nie ma — jego kroki
// korzystają z mapy tego samego skryptu z Chromium (albo, bez niej, lecą przy każdej zmianie).
//
// Zapisuje JSON: { files: { "app/x.js": { hash, exec: [[s,e]…], all: [[s,e]…] } },
//                 css: { "styles/x.css": { hash, used: [[s,e]…] } }, sleepMs, timers }
//   exec — zakresy (przesunięcia w tekście pliku) funkcji wywołanych choć raz,
//   all  — zakresy wszystkich funkcji znanych V8 (do odróżnienia kodu w funkcji od kodu na wierzchu),
//   used — reguły CSS, które choć raz do czegoś pasowały (CSS.startRuleUsageTracking),
//   sleepMs/timers — suma sztywnych pauz testu (page.waitForTimeout + setTimeout ≥ 100 ms w Node).
// Kopia pliku z chwili nagrania trafia do .test-cache/src/<hash> (do porównania przy kolejnym teście).
"use strict";

const OUT = process.env.TEST_COV_OUT;
if (OUT) {
  const fs = require("fs");
  const path = require("path");
  const crypto = require("crypto");
  const ROOT = process.env.TEST_COV_ROOT || process.cwd();
  const SRC_DIR = path.join(ROOT, ".test-cache", "src");
  const merged = new Map(); // "app/x.js" → { exec: Map(key→[s,e]), all: Map }
  const cssUsed = new Map(); // "styles/x.css" → Map(key→[s,e])
  let sleepMs = 0;
  let timers = 0;

  // sztywne pauzy testu (helper sleep = setTimeout) — tylko do statystyki; liczymy wyłącznie
  // wywołania z plików testów (scripts/), nie limity czasu Playwrighta z node_modules
  const SCRIPTS = path.join(ROOT, "scripts") + path.sep;
  const fromTest = () => {
    const frames = String(new Error().stack).split("\n").slice(2, 6);
    const f = frames.find((l) => !l.includes("coverage-preload.cjs") && !l.includes("sleep-scale.cjs"));
    return !!f && f.includes(SCRIPTS) && !f.includes("node_modules");
  };
  const origSetTimeout = global.setTimeout;
  global.setTimeout = function patchedSetTimeout(fn, ms, ...rest) {
    if (typeof ms === "number" && ms >= 50 && fromTest()) { sleepMs += ms; timers++; }
    return origSetTimeout.call(this, fn, ms, ...rest);
  };

  const toRepoPath = (url) => {
    try {
      const u = new URL(url);
      if (!/^(127\.0\.0\.1|localhost)$/.test(u.hostname)) return null;
      const p = decodeURIComponent(u.pathname).replace(/^\/+/, "");
      return /^app\/[^/]+\.js$/.test(p) || /^styles\/[^/]+\.css$/.test(p) ? p : null;
    } catch { return null; }
  };

  function mergeCoverage(result) {
    for (const script of result || []) {
      const rel = toRepoPath(script.url);
      if (!rel) continue;
      let m = merged.get(rel);
      if (!m) { m = { exec: new Map(), all: new Map() }; merged.set(rel, m); }
      for (const fn of script.functions || []) {
        const r = fn.ranges && fn.ranges[0];
        if (!r) continue;
        if (r.startOffset === 0 && !fn.functionName) continue; // sam skrypt (kod na wierzchu)
        const key = `${r.startOffset}:${r.endOffset}`;
        m.all.set(key, [r.startOffset, r.endOffset]);
        if (r.count > 0) m.exec.set(key, [r.startOffset, r.endOffset]);
      }
    }
  }

  const pages = new Set(); // { page, cdp, done }
  async function flushPage(rec) {
    if (rec.done) return;
    try {
      const { result } = await rec.cdp.send("Profiler.takePreciseCoverage");
      mergeCoverage(result);
    } catch { /* strona już zamknięta — trudno, część pokrycia mogła przepaść */ }
    if (process.env.TEST_COV_CSS !== "0") try {
      const { ruleUsage } = await rec.cdp.send("CSS.stopRuleUsageTracking");
      for (const r of ruleUsage || []) {
        if (!r.used) continue;
        const rel = toRepoPath(rec.sheets.get(r.styleSheetId) || "");
        if (!rel || !rel.endsWith(".css")) continue;
        if (!cssUsed.has(rel)) cssUsed.set(rel, new Map());
        cssUsed.get(rel).set(`${r.startOffset}:${r.endOffset}`, [r.startOffset, r.endOffset]);
      }
      await rec.cdp.send("CSS.startRuleUsageTracking"); // strona żyje dalej (nawigacja w teście)
    } catch { /* bez mapy CSS krok poleci przy każdej zmianie stylów */ }
  }
  async function flushAll() { await Promise.all([...pages].map(flushPage)); }

  async function hookPage(page) {
    if (!page || page.__covHooked) return;
    page.__covHooked = true;
    let cdp;
    try {
      cdp = await page.context().newCDPSession(page);
      await cdp.send("Profiler.enable");
      await cdp.send("Profiler.startPreciseCoverage", { callCount: true, detailed: false });
    } catch { return; }
    const rec = { page, cdp, done: false, sheets: new Map() };
    cdp.on("CSS.styleSheetAdded", (e) => rec.sheets.set(e.header.styleSheetId, e.header.sourceURL));
    // Śledzenie reguł CSS spowalnia przeliczanie stylów (~4× w dokumencie Docs) — testy, które
    // mierzą wydajność stylów, mają je wyłączone (TEST_COV_CSS=0, znacznik w pliku testu); dla nich
    // każda zmiana CSS uruchamia krok, co przy strażniku wydajności i tak jest właściwe.
    if (process.env.TEST_COV_CSS !== "0") {
      try {
        await cdp.send("DOM.enable");
        await cdp.send("CSS.enable");
        await cdp.send("CSS.startRuleUsageTracking");
      } catch { /* brak śledzenia CSS — wtedy zmiana stylów zawsze uruchamia krok */ }
    }
    pages.add(rec);
    // przed nawigacją / zamknięciem zbieramy (skrypty starego dokumentu mogą zniknąć)
    for (const name of ["goto", "reload", "setContent"]) {
      const orig = page[name];
      if (typeof orig !== "function") continue;
      page[name] = async function (...args) { await flushPage(rec); return orig.apply(this, args); };
    }
    const origClose = page.close;
    page.close = async function (...args) { await flushPage(rec); rec.done = true; return origClose.apply(this, args); };
    const origWait = page.waitForTimeout;
    if (typeof origWait === "function") page.waitForTimeout = function (ms, ...rest) { if (typeof ms === "number") { sleepMs += ms * (Number(process.env.SLEEP_SCALE) || 1); timers++; } return origWait.call(this, ms, ...rest); };
  }

  function hookContext(ctx) {
    if (!ctx || ctx.__covHooked) return ctx;
    ctx.__covHooked = true;
    ctx.on("page", (p) => { hookPage(p); }); // wyskakujące okna — w miarę możliwości
    const origNewPage = ctx.newPage;
    ctx.newPage = async function (...args) { const p = await origNewPage.apply(this, args); await hookPage(p); return p; };
    const origClose = ctx.close;
    ctx.close = async function (...args) { await flushAll(); return origClose.apply(this, args); };
    return ctx;
  }

  function hookBrowser(b) {
    if (!b || b.__covHooked) return b;
    b.__covHooked = true;
    const origNewContext = b.newContext;
    b.newContext = async function (...args) { return hookContext(await origNewContext.apply(this, args)); };
    const origNewPage = b.newPage;
    b.newPage = async function (...args) { const p = await origNewPage.apply(this, args); hookContext(p.context()); await hookPage(p); return p; };
    const origClose = b.close;
    b.close = async function (...args) { await flushAll(); return origClose.apply(this, args); };
    return b;
  }

  let pw = null;
  for (const name of ["playwright", "playwright-core", "@playwright/test"]) {
    try { pw = require(require.resolve(name, { paths: [ROOT] })); break; } catch { /* brak */ }
  }
  if (pw && pw.chromium) {
    const bt = pw.chromium;
    const origLaunch = bt.launch;
    bt.launch = async function (...args) { return hookBrowser(await origLaunch.apply(this, args)); };
    const origPersistent = bt.launchPersistentContext;
    if (typeof origPersistent === "function") bt.launchPersistentContext = async function (...args) { return hookContext(await origPersistent.apply(this, args)); };
  }

  process.on("exit", () => {
    const files = {};
    const snapshot = (rel) => {
      let text;
      try { text = fs.readFileSync(path.join(ROOT, rel), "utf8"); } catch { return null; }
      const hash = crypto.createHash("sha1").update(text).digest("hex");
      try {
        fs.mkdirSync(SRC_DIR, { recursive: true });
        const snap = path.join(SRC_DIR, hash);
        if (!fs.existsSync(snap)) fs.writeFileSync(snap, text);
      } catch { /* bez kopii krok po prostu poleci przy następnej zmianie pliku */ }
      return hash;
    };
    // CSS.ruleUsage podaje położenie w BAJTACH UTF-8 (polskie litery w komentarzach = 2 bajty),
    // a reszta (skaner reguł, porównanie) liczy w znakach — przeliczamy przy zapisie.
    const byteToChar = (text) => {
      const map = new Int32Array(Buffer.byteLength(text, "utf8") + 1);
      let b = 0;
      for (let i = 0; i < text.length; i++) {
        const c = text.codePointAt(i);
        const len = c < 0x80 ? 1 : c < 0x800 ? 2 : c < 0x10000 ? 3 : 4;
        for (let k = 0; k < len; k++) map[b + k] = i;
        b += len;
        if (c >= 0x10000) i++; // para zastępcza = 2 znaki JS
      }
      map[b] = text.length;
      return (x) => map[Math.min(Math.max(0, x), b)];
    };
    const css = {};
    for (const [rel, m] of cssUsed) {
      const hash = snapshot(rel);
      if (!hash) continue;
      let text;
      try { text = fs.readFileSync(path.join(ROOT, rel), "utf8"); } catch { continue; }
      const toChar = byteToChar(text);
      css[rel] = { hash, used: [...m.values()].map(([s, e]) => [toChar(s), toChar(e)]) };
    }
    for (const [rel, m] of merged) {
      if (rel.endsWith(".css")) continue;
      let text;
      try { text = fs.readFileSync(path.join(ROOT, rel), "utf8"); } catch { continue; }
      const hash = crypto.createHash("sha1").update(text).digest("hex");
      try {
        fs.mkdirSync(SRC_DIR, { recursive: true });
        const snap = path.join(SRC_DIR, hash);
        if (!fs.existsSync(snap)) fs.writeFileSync(snap, text);
      } catch { /* bez kopii test po prostu poleci przy następnej zmianie pliku */ }
      files[rel] = { hash, exec: [...m.exec.values()], all: [...m.all.values()] };
    }
    try {
      fs.mkdirSync(path.dirname(OUT), { recursive: true });
      fs.writeFileSync(OUT, JSON.stringify({ files, css, sleepMs, timers }));
    } catch { /* brak zapisu = test poleci następnym razem */ }
  });
}
