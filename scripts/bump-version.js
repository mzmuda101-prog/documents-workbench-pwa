#!/usr/bin/env node
/**
 * bump-version.js
 * Aktualizuje wersję cache/buildu we wszystkich plikach projektu jedną komendą.
 *
 * Użycie:
 *   cd documents-workbench-pwa
 *   npm run release            → auto: jeśli dziś już wydano, podbija -NN (np. 20260530-02),
 *                                w innym wypadku dzisiejsza data jako "YYYYMMDD-01"
 *   npm run release 20260601-02  → użyje podanej wersji
 *
 * ODPORNOŚĆ NA ROZJECHANIE WERSJI:
 *   Skrypt NIE zakłada, że wszystkie pliki mają tę samą starą wersję.
 *   Zamiast podmieniać jeden znany stary string, celuje wzorcami w:
 *     1) ?v=YYYYMMDD-NN przy plikach .css / .js / .mjs   (buildowe assety)
 *     2) const CACHE_VERSION = "YYYYMMDD-NN"             (sw.js)
 *     3) const APP_BUILD_VERSION = "YYYYMMDD-NN"         (core.js)
 *   Dzięki temu nawet rozjechane pliki zostają zrównane. Wersje ikon
 *   (.png?v=...) NIE są ruszane — mają własny, niezależny cykl.
 */

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const VERSION_RE = /^\d{8}-\d{2}$/;

// Pliki, w których pilnujemy wersji buildu. Wzorce same wybiorą właściwe miejsca.
const TARGET_FILES = ["sw.js", "app/core.js", "index.html"];

// ─── Pomocnicze ───────────────────────────────────────────────────────────────

function read(relPath) {
  return fs.readFileSync(path.join(ROOT, relPath), "utf8");
}
function write(relPath, content) {
  fs.writeFileSync(path.join(ROOT, relPath), content, "utf8");
}
function exists(relPath) {
  return fs.existsSync(path.join(ROOT, relPath));
}

// Podmienia wszystkie buildowe wystąpienia wersji na newVersion.
// Zwraca { content, count }. Pomija .png?v=... (ikony) bo wzorzec łapie tylko css/js/mjs.
function bumpBuildVersions(content, newVersion) {
  let count = 0;
  const bump = (re) => {
    content = content.replace(re, (match) => {
      count += 1;
      return match.replace(/\d{8}-\d{2}/, newVersion);
    });
  };
  bump(/\.(?:css|mjs|js)\?v=\d{8}-\d{2}/g);                // 1) assety css/js/mjs
  bump(/CACHE_VERSION\s*=\s*"\d{8}-\d{2}"/g);              // 2) sw.js
  bump(/APP_BUILD_VERSION\s*=\s*"\d{8}-\d{2}"/g);          // 3) core.js
  return { content, count };
}

// Wszystkie buildowe wersje w treści (bez ikon .png) — do wykrycia rozjazdu.
function collectBuildVersions(content) {
  const found = new Set();
  const add = (re) => {
    const m = content.match(re) || [];
    m.forEach((s) => {
      const v = (s.match(/\d{8}-\d{2}/) || [])[0];
      if (v) found.add(v);
    });
  };
  add(/\.(?:css|mjs|js)\?v=\d{8}-\d{2}/g);
  add(/CACHE_VERSION\s*=\s*"\d{8}-\d{2}"/g);
  add(/APP_BUILD_VERSION\s*=\s*"\d{8}-\d{2}"/g);
  return found;
}

// ─── Ustal aktualną wersję (z CACHE_VERSION w sw.js) i policz następną ─────────

const swRaw = read("sw.js");
const currentMatch = swRaw.match(/CACHE_VERSION\s*=\s*"(\d{8}-\d{2})"/);
const currentVersion = currentMatch ? currentMatch[1] : null;
if (!currentVersion) {
  console.error("❌  Nie znaleziono CACHE_VERSION w sw.js");
  process.exit(1);
}

function nextAutoVersion(current) {
  const now = new Date();
  const today = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}`;
  if (current && current.startsWith(`${today}-`)) {
    const nn = parseInt(current.slice(-2), 10) || 0;
    return `${today}-${String(nn + 1).padStart(2, "0")}`; // dziś już wydano → podbij -NN
  }
  return `${today}-01`;
}

const newVersion = process.argv[2] || nextAutoVersion(currentVersion);
if (!VERSION_RE.test(newVersion)) {
  console.error(`❌  Nieprawidłowy format wersji: "${newVersion}". Oczekiwany format: YYYYMMDD-NN (np. 20260601-01)`);
  process.exit(1);
}

console.log(`\n🔄  Wersja buildu: ${currentVersion} → ${newVersion}`);
if (process.argv[2] && process.argv[2] === currentVersion) {
  console.log("ℹ️  Podana wersja = bieżąca. Skrypt i tak przejedzie pliki, żeby zrównać ewentualny rozjazd.");
}
console.log("");

// ─── Przejedź pliki ───────────────────────────────────────────────────────────

const changes = [];
const driftWarnings = [];

for (const file of TARGET_FILES) {
  if (!exists(file)) {
    changes.push(`⚠️   ${file} — brak pliku (pominięto)`);
    continue;
  }
  const before = read(file);
  const { content: after, count } = bumpBuildVersions(before, newVersion);
  if (count > 0) write(file, after);

  // Wykryj resztki innych wersji buildu (rozjazd), które nie wpadły we wzorce
  const leftover = [...collectBuildVersions(after)].filter((v) => v !== newVersion);
  if (leftover.length) {
    driftWarnings.push(`   • ${file}: pozostały inne wersje buildu → ${leftover.join(", ")}`);
  }
  changes.push(`✅  ${file} — zaktualizowano ${count} wystąpień`);
}

// ─── Podsumowanie ─────────────────────────────────────────────────────────────

console.log(changes.join("\n"));
if (driftWarnings.length) {
  console.log(`\n⚠️  Wykryto pozostały rozjazd wersji (sprawdź ręcznie):`);
  console.log(driftWarnings.join("\n"));
}
console.log(`\n🎉  Gotowe! Nowa wersja: ${newVersion}`);
console.log(`\nNastępne kroki:`);
console.log(`  git add -A`);
console.log(`  git commit -m "release: ${newVersion}"`);
console.log(`  git push`);
