// test-impact.mjs — czy test trzeba puścić, skoro od jego ostatniego zielonego przebiegu coś się zmieniło?
//
// Używane przez run-tests.mjs: `npm test` pomija kroki, których zmiana na pewno nie dotyczy;
// `npm run test:all` (albo `--all`) puszcza wszystko. Zasada: przy każdej NIEPEWNOŚCI test leci.
//
// Pamięć (.test-cache/, poza gitem):
//   impact.json — dla kroku: odcisk „wszystkiego wspólnego” (CSS, index.html, lib/, sw.js…), odciski
//                 samego testu i jego pomocników/plików z danymi, oraz mapa funkcji apki, które test
//                 wykonał (z coverage-preload.cjs, Chromium) — z hashem pliku z chwili nagrania,
//   src/<hash>  — kopia pliku apki z tej chwili (do porównania linii z obecnym).
//
// Krok leci, gdy: brak zapisu / ostatnio nie przeszedł; zmienił się odcisk wspólny (index.html, lib/,
// sw.js, próbki); zmienił się test,
// jego pomocnik albo plik z danymi; zmieniony fragment pliku apki leży w funkcji, którą test wykonał,
// albo POZA funkcjami (kod na wierzchu skryptu — działa przy wczytaniu, więc dotyczy każdego testu);
// zmieniona reguła CSS była przez test użyta, zmienił się selektor, doszła nowa reguła albo at-reguła;
// plik apki zmienił się, a mapy funkcji dla niego nie ma (WebKit bez odpowiednika, test bez
// przeglądarki). Numery wersji wydań (20261029-03) nie liczą się jako zmiana.
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";

const VERSION_RE = /\d{8}-\d{2}/g;
// ta sama długość (przesunięcia w pliku z nagrania zostają ważne)
export const normalizeVersions = (text) => text.replace(VERSION_RE, "00000000-00");
const sha1 = (s) => crypto.createHash("sha1").update(s).digest("hex");
const readText = (p) => { try { return fs.readFileSync(p, "utf8"); } catch { return null; } };
const readBuf = (p) => { try { return fs.readFileSync(p); } catch { return null; } };

function listFiles(dir, acc = []) {
  let entries;
  try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return acc; }
  for (const e of entries) {
    if (e.name.startsWith(".") || e.name === "node_modules") continue;
    const p = path.join(dir, e.name);
    if (e.isDirectory()) listFiles(p, acc); else acc.push(p);
  }
  return acc;
}

// Odcisk rzeczy wspólnych dla wszystkich testów przeglądarkowych: zmiana = wszystko leci.
// (styles/ osobno — reguły CSS mają własną mapę użycia, patrz cssChangeHits)
export function globalFingerprint(root) {
  const files = [
    ...["index.html", "sw.js", "manifest.json"].map((f) => path.join(root, f)),
    ...listFiles(path.join(root, "lib")),
    ...listFiles(path.join(root, "docs", "samples")),
    path.join(root, "scripts", "sleep-scale.cjs"),
    path.join(root, "scripts", "test-server.py"),
  ].filter((f) => fs.existsSync(f)).sort();
  const h = crypto.createHash("sha1");
  for (const f of files) {
    const rel = path.relative(root, f);
    const buf = readBuf(f);
    if (!buf) continue;
    const isText = /\.(html|js|css|json|py|cjs|mjs)$/.test(f);
    h.update(rel + "\0" + (isText ? normalizeVersions(buf.toString("utf8")) : buf) + "\0");
  }
  return h.digest("hex");
}

const FIXTURE_RE = /["'`]([^"'`\n]{1,200}\.(?:xlsx|xlsm|xls|docx|pdf|csv|png|jpe?g|txt|json|xml|html|odt|rtf))["'`]/gi;
const REQUIRE_RE = /(?:require\(\s*|from\s+|import\(\s*)["'](\.{1,2}\/[^"']+)["']/g;

// Plik testu + lokalne moduły (rekurencyjnie) + pliki z danymi, których nazwy w nich padają.
export function testDependencies(root, step) {
  const deps = new Set();
  const scriptArg = step.args.find((a) => /\.(c|m)?js$/.test(a));
  if (!scriptArg) return deps;
  const fixtureDirs = [path.join(root, "scripts"), path.join(root, "..", "test-fixtures"), path.join(root, "docs", "samples")];
  const fixtureIndex = new Map(); // nazwa pliku → ścieżki
  for (const d of fixtureDirs) for (const f of listFiles(d)) {
    const k = path.basename(f).toLowerCase();
    if (!fixtureIndex.has(k)) fixtureIndex.set(k, []);
    fixtureIndex.get(k).push(f);
  }
  const visit = (file) => {
    if (deps.has(file) || !fs.existsSync(file)) return;
    deps.add(file);
    const src = readText(file) || "";
    for (const m of src.matchAll(REQUIRE_RE)) {
      const base = path.resolve(path.dirname(file), m[1]);
      for (const cand of [base, base + ".js", base + ".cjs", base + ".mjs", path.join(base, "index.js")]) {
        if (fs.existsSync(cand) && fs.statSync(cand).isFile()) { visit(cand); break; }
      }
    }
    for (const m of src.matchAll(FIXTURE_RE)) {
      const name = path.basename(m[1]).toLowerCase();
      for (const f of fixtureIndex.get(name) || []) deps.add(f);
    }
  };
  visit(path.resolve(root, scriptArg));
  // pliki podane w zmiennych kroku (DOC=…, FILE=…)
  for (const v of Object.values(step.env)) {
    const p = path.resolve(root, v);
    if (/\.\w{2,5}$/.test(v) && fs.existsSync(p)) deps.add(p);
  }
  return deps;
}

export function depsFingerprint(root, deps) {
  const out = {};
  for (const f of [...deps].sort()) {
    const buf = readBuf(f);
    if (!buf) continue;
    out[path.relative(root, f)] = sha1(/\.(c|m)?js$/.test(f) ? normalizeVersions(buf.toString("utf8")) : buf);
  }
  return out;
}

export function cssFileHashes(root) {
  const out = {};
  for (const f of listFiles(path.join(root, "styles"))) {
    if (!f.endsWith(".css")) continue;
    out[path.relative(root, f)] = sha1(readText(f));
  }
  return out;
}

// Komentarze CSS → spacje (ta sama długość, nowe linie zostają): zmiana samego komentarza
// po porównaniu z ignorowaniem białych znaków znika.
export function blankCssComments(t) {
  return t.replace(/\/\*[\s\S]*?\*\//g, (m) => m.replace(/[^\n]/g, " "));
}

// Reguły arkusza: { start (selektor), open ({), end (za }), kind: style | opaque }. Grupy (@media,
// @supports, @container, @layer, @scope) są przezroczyste — liczą się reguły w środku. Inne
// at-reguły (@keyframes, @font-face…) to „opaque”: zmiana w nich uruchamia każdy krok.
export function scanCss(t) {
  const rules = [];
  const n = t.length;
  const skipString = (i) => { const q = t[i]; i++; while (i < n && t[i] !== q) { if (t[i] === "\\") i++; i++; } return i + 1; };
  const matchBrace = (i) => { // i na „{” → indeks pasującego „}”
    let depth = 0;
    for (; i < n; i++) {
      const c = t[i];
      if (c === '"' || c === "'") { i = skipString(i) - 1; continue; }
      if (c === "/" && t[i + 1] === "*") { const e = t.indexOf("*/", i + 2); i = e < 0 ? n : e + 1; continue; }
      if (c === "{") depth++;
      else if (c === "}" && --depth === 0) return i;
    }
    return n - 1;
  };
  const list = (i) => {
    while (i < n) {
      while (i < n && /\s/.test(t[i])) i++;
      if (t[i] === "/" && t[i + 1] === "*") { const e = t.indexOf("*/", i + 2); i = e < 0 ? n : e + 2; continue; }
      if (i >= n || t[i] === "}") return i;
      const start = i;
      let j = i;
      while (j < n && t[j] !== "{" && t[j] !== ";" && t[j] !== "}") {
        if (t[j] === '"' || t[j] === "'") { j = skipString(j); continue; }
        if (t[j] === "/" && t[j + 1] === "*") { const e = t.indexOf("*/", j + 2); j = e < 0 ? n : e + 2; continue; }
        j++;
      }
      if (j >= n) return n;
      if (t[j] === ";") { rules.push({ start, open: j, end: j + 1, kind: "opaque" }); i = j + 1; continue; }
      if (t[j] === "}") return j;
      const prelude = t.slice(start, j).trim();
      if (prelude.startsWith("@")) {
        const name = prelude.slice(1).split(/[\s({]/)[0].toLowerCase();
        if (["media", "supports", "container", "layer", "scope", "document"].includes(name)) {
          rules.push({ start, open: j, end: j + 1, kind: "opaque" }); // sam nagłówek grupy
          const close = list(j + 1);
          rules.push({ start: close, open: close, end: close + 1, kind: "opaque" });
          i = close + 1;
          continue;
        }
        const close = matchBrace(j);
        rules.push({ start, open: j, end: close + 1, kind: "opaque" });
        i = close + 1;
        continue;
      }
      const close = matchBrace(j);
      rules.push({ start, open: j, end: close + 1, kind: "style" });
      i = close + 1;
    }
    return i;
  };
  list(0);
  return rules;
}

// Czy zmiana arkusza dotyczy testu, który użył reguł `used` (Chromium)? union — reguły użyte przez
// JAKIKOLWIEK test w Chromium (dla WebKitu: reguła nieużyta nigdzie może być regułą pod Safari).
function cssChangeHits(rec, root, rel, union) {
  const cur = readText(path.join(root, rel));
  if (cur == null) return `arkusz ${rel} zniknął`;
  if (sha1(cur) === rec.hash) return null;
  const old = readText(path.join(root, ".test-cache", "src", rec.hash));
  if (old == null) return `zmiana w ${rel} (brak kopii z nagrania)`;
  const a = blankCssComments(normalizeVersions(old)), b = blankCssComments(normalizeVersions(cur));
  const regions = changedRegions(a, b, ["-b", "-B"]);
  if (!regions) return `zmiana w ${rel} (porównanie się nie udało)`;
  if (!regions.length) return null; // same komentarze / białe znaki / numery wersji
  const rules = scanCss(a);
  const isUsed = (r, list) => list.some(([s, e]) => s <= r.open && r.open < e);
  for (const [lo, hi] of regions) {
    const r = rules.find((x) => x.start <= lo && (lo === hi ? lo < x.end : hi <= x.end));
    if (!r) return `nowa reguła albo zmiana między regułami (${rel})`;
    if (r.kind !== "style") return `zmiana w at-regule (${rel})`;
    if (lo <= r.open) return `zmiana selektora (${rel})`;
    if (isUsed(r, rec.used)) return `zmiana reguły CSS użytej przez test (${rel})`;
    if (union && !isUsed(r, union)) return `zmiana reguły CSS, której Chromium nie używa — może być pod Safari (${rel})`;
  }
  return null;
}

export function appFileHashes(root, normalized = false) {
  const out = {};
  for (const f of listFiles(path.join(root, "app"))) {
    if (!f.endsWith(".js")) continue;
    const t = readText(f);
    out[path.relative(root, f)] = sha1(normalized ? normalizeVersions(t) : t);
  }
  return out;
}

function lineStarts(text) {
  const starts = [0];
  for (let i = 0; i < text.length; i++) if (text.charCodeAt(i) === 10) starts.push(i + 1);
  return starts;
}

// Zmienione miejsca STAREGO tekstu: [[od, do]] (zmiana/usunięcie) albo [p, p] (wstawka w punkcie p).
function changedRegions(oldText, newText, flags = []) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "test-impact-"));
  const a = path.join(dir, "a"), b = path.join(dir, "b");
  fs.writeFileSync(a, oldText); fs.writeFileSync(b, newText);
  const r = spawnSync("diff", [...flags, a, b], { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
  fs.rmSync(dir, { recursive: true, force: true });
  if (r.status === 0) return [];
  if (r.status !== 1) return null; // diff się nie udał — traktuj jak „nie wiadomo”
  const starts = lineStarts(oldText);
  const nStarts = lineStarts(newText);
  const lineStart = (n) => (n <= 0 ? 0 : n > starts.length ? oldText.length : starts[n - 1]); // n: 1..
  const lineEnd = (n) => (n >= starts.length ? oldText.length : starts[n] - 1);
  const nLineStart = (n) => (n <= 0 ? 0 : n > nStarts.length ? newText.length : nStarts[n - 1]);
  const nLineEnd = (n) => (n >= nStarts.length ? newText.length : nStarts[n] - 1);
  const regions = [];
  for (const line of r.stdout.split("\n")) {
    const m = /^(\d+)(?:,(\d+))?([acd])(\d+)(?:,(\d+))?$/.exec(line);
    if (!m) continue;
    const s = +m[1], e = m[2] ? +m[2] : s;
    if (m[3] === "a") { const p = lineStart(s + 1); regions.push([p, p]); continue; }
    let lo = lineStart(s), hi = lineEnd(e);
    if (m[3] === "c") {
      // zmienione linie → zawężenie do znaków: wspólny początek i koniec starego i nowego fragmentu
      // nie są zmianą (deklaracja dopisana w regule zapisanej w jednej linii z selektorem itp.)
      const ns = +m[4], ne = m[5] ? +m[5] : ns;
      const a = oldText.slice(lo, hi), b = newText.slice(nLineStart(ns), nLineEnd(ne));
      let p = 0;
      while (p < a.length && p < b.length && a[p] === b[p]) p++;
      let q = 0;
      while (q < a.length - p && q < b.length - p && a[a.length - 1 - q] === b[b.length - 1 - q]) q++;
      lo += p; hi -= q;
      if (hi < lo) hi = lo;
    }
    regions.push([lo, hi]);
  }
  return regions;
}

// Czy zmiana pliku dotyczy testu, który wykonał funkcje `exec` (znane funkcje: `all`)?
function fileChangeHits(rec, root, rel, unionExec) {
  const cur = readText(path.join(root, rel));
  if (cur == null) return "plik apki zniknął";
  if (sha1(cur) === rec.hash) return null;
  if (rec.nhash && sha1(normalizeVersions(cur)) === rec.nhash) return null; // sam numer wersji wydania
  if (rec.unknown) return `zmiana w ${rel} (brak mapy funkcji)`;
  const old = readText(path.join(root, ".test-cache", "src", rec.hash));
  if (old == null) return `zmiana w ${rel} (brak kopii z nagrania)`;
  const a = normalizeVersions(old), b = normalizeVersions(cur);
  if (a === b) return null;
  const regions = changedRegions(a, b);
  if (!regions) return `zmiana w ${rel} (porównanie się nie udało)`;
  const inside = (lo, hi, [s, e]) => (lo === hi ? s < lo && lo < e : s <= lo && hi <= e);
  const overlaps = (lo, hi, [s, e]) => (lo === hi ? s < lo && lo < e : s < hi && e > lo);
  for (const [lo, hi] of regions) {
    if (rec.exec.some((r) => overlaps(lo, hi, r))) return `zmiana w funkcji wykonanej przez test (${rel})`;
    if (!rec.all.some((r) => inside(lo, hi, r))) return `zmiana poza funkcjami (${rel})`;
    // WebKit z mapą od Chromium: funkcja, której nie wykonał ŻADEN test w Chromium, może być
    // ścieżką tylko pod Safari (wykrywanie silnika, inne API) — wtedy krok WebKitu leci
    if (unionExec && !unionExec.some((r) => overlaps(lo, hi, r))) return `zmiana w funkcji, której Chromium nie wykonuje — może być ścieżką Safari (${rel})`;
  }
  return null; // zmienione tylko funkcje, których ten test nie uruchomił
}

export function loadImpact(root) {
  try { return JSON.parse(fs.readFileSync(path.join(root, ".test-cache", "impact.json"), "utf8")); } catch { return {}; }
}
export function saveImpact(root, data) {
  fs.mkdirSync(path.join(root, ".test-cache"), { recursive: true });
  fs.writeFileSync(path.join(root, ".test-cache", "impact.json"), JSON.stringify(data));
}

// → null (można pominąć) albo powód, dla którego krok leci
// Suma funkcji wykonanych (JS) / reguł użytych (CSS) przez wszystkie testy w Chromium dla danej
// wersji pliku.
export function unionExecFor(impact) {
  const cache = new Map();
  return (rel, hash) => {
    const k = rel + "\0" + hash;
    if (!cache.has(k)) {
      const out = [];
      for (const r of Object.values(impact)) {
        if (!r || r.cov !== "self") continue;
        const f = rel.endsWith(".css") ? r.css && r.css[rel] : r.files && r.files[rel];
        const list = f && f.hash === hash && (f.exec || f.used);
        if (list) out.push(...list);
      }
      cache.set(k, out);
    }
    return cache.get(k);
  };
}

export function whyRun(root, step, rec, globalNow, unionExec = null) {
  if (!rec) return "brak zapisu ostatniego zielonego przebiegu";
  if (rec.global !== globalNow) return "zmiana wspólna (index.html / lib / sw.js / próbki)";
  const depsNow = depsFingerprint(root, testDependencies(root, step));
  const names = new Set([...Object.keys(rec.deps || {}), ...Object.keys(depsNow)]);
  for (const n of names) if ((rec.deps || {})[n] !== depsNow[n]) return `zmiana w teście lub jego danych (${n})`;
  for (const [rel, f] of Object.entries(rec.files || {})) {
    const hit = fileChangeHits(f, root, rel, rec.cov === "sibling" && unionExec ? unionExec(rel, f.hash) : null);
    if (hit) return hit;
  }
  // arkusze: z mapą reguł — ocena reguła po regule; bez mapy — każda zmiana uruchamia krok
  const cssNow = cssFileHashes(root);
  for (const rel of new Set([...Object.keys(cssNow), ...Object.keys(rec.cssHashes || {})])) {
    const m = rec.css && rec.css[rel];
    if (m && m.used) {
      const hit = cssChangeHits(m, root, rel, rec.cov === "sibling" && unionExec ? unionExec(rel, m.hash) : null);
      if (hit) return hit;
    } else if ((rec.cssHashes || {})[rel] !== cssNow[rel]) return `zmiana w ${rel} (brak mapy reguł CSS)`;
  }
  return null;
}

// Zapis po zielonym przebiegu. cov — wynik coverage-preload (Chromium) albo null; sibling — zapis
// tego samego skryptu z Chromium (dla WebKitu), z którego bierzemy mapę funkcji, gdy pliki się zgadzają.
export function makeRecord(root, step, globalNow, cov, sibling) {
  const rec = { at: Date.now(), global: globalNow, deps: depsFingerprint(root, testDependencies(root, step)), files: {}, cssHashes: cssFileHashes(root), css: {} };
  const hashes = appFileHashes(root);
  const nhashes = appFileHashes(root, true);
  const mapped = cov && cov.files && Object.keys(cov.files).length ? cov.files : null;
  if (mapped) {
    for (const [rel, f] of Object.entries(mapped)) rec.files[rel] = { hash: f.hash, exec: f.exec, all: f.all };
    // mapa reguł tylko dla arkuszy, które strona faktycznie wczytała (i tylko gdy plik się zgadza)
    for (const [rel, m] of Object.entries((cov && cov.css) || {})) if (rec.cssHashes[rel] === m.hash) rec.css[rel] = m;
    rec.cov = "self";
    return rec;
  }
  if (sibling && sibling.cov === "self") {
    for (const [rel, f] of Object.entries(sibling.files)) {
      rec.files[rel] = hashes[rel] === f.hash ? f : { hash: hashes[rel], nhash: nhashes[rel], unknown: true };
    }
    for (const [rel, m] of Object.entries(sibling.css || {})) if (rec.cssHashes[rel] === m.hash) rec.css[rel] = m;
    rec.cov = "sibling";
    return rec;
  }
  for (const [rel, h] of Object.entries(hashes)) rec.files[rel] = { hash: h, nhash: nhashes[rel], unknown: true };
  rec.cov = "none";
  return rec;
}
