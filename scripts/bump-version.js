#!/usr/bin/env node
/**
 * bump-version.js
 * Podbija wersję aplikacji we wszystkich plikach projektu jedną komendą.
 *
 * SCHEMAT (od 2026-11-07, jak Smart Kalkulator): vMAJOR.MM.PP, np. v1.00.03
 *   MAJOR — przełom w aplikacji (rzadko),
 *   MM    — duża nowa funkcja / przebudowa,
 *   PP    — kolejne wydanie; drobna poprawka +1, większy etap skok do okrągłej …0 / …5.
 * Poprzedni schemat datowy (YYYYMMDD-NN) skończył się na wydaniu, które zastąpiło v1.00.00;
 * skrypt rozpoznaje go nadal, żeby zrównać ewentualne resztki.
 *
 * Użycie:
 *   cd documents-workbench-pwa
 *   npm run release                → +1 (v1.00.03 → v1.00.04; z wersji datowej → v1.00.00)
 *   npm run release -- skok        → do najbliższej okrągłej …0/…5 (v1.00.03 → v1.00.05, v1.00.05 → v1.00.10)
 *   npm run release -- minor       → v1.01.00
 *   npm run release -- major       → v2.00.00
 *   npm run release -- v1.02.00    → podana wersja
 *
 * ODPORNOŚĆ NA ROZJECHANIE WERSJI:
 *   Skrypt NIE zakłada, że wszystkie pliki mają tę samą starą wersję.
 *   Celuje wzorcami w:
 *     1) ?v=WERSJA przy plikach .css / .js / .mjs   (buildowe assety)
 *     2) const CACHE_VERSION = "WERSJA"              (sw.js)
 *     3) const APP_BUILD_VERSION = "WERSJA"          (core.js)
 *   Wersje ikon (.png?v=...) NIE są ruszane — mają własny, niezależny cykl.
 */

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const VERSION_RE = /^v\d+\.\d{2}\.\d{2}$/;
// Dowolna wersja buildu w plikach: nowa (v1.00.03) albo stara datowa (20261106-06).
const ANY_VER = "(?:v\\d+\\.\\d{2}\\.\\d{2}|\\d{8}-\\d{2})";

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

const PATTERNS = [
  new RegExp(`\\.(?:css|mjs|js)\\?v=${ANY_VER}`, "g"),          // 1) assety css/js/mjs
  new RegExp(`CACHE_VERSION\\s*=\\s*"${ANY_VER}"`, "g"),        // 2) sw.js
  new RegExp(`APP_BUILD_VERSION\\s*=\\s*"${ANY_VER}"`, "g"),    // 3) core.js
];
const ONE_VER = new RegExp(ANY_VER);

// Podmienia wszystkie buildowe wystąpienia wersji na newVersion. Zwraca { content, count }.
function bumpBuildVersions(content, newVersion) {
  let count = 0;
  for (const re of PATTERNS) {
    content = content.replace(re, (match) => {
      count += 1;
      return match.replace(ONE_VER, newVersion);
    });
  }
  return { content, count };
}

// Wszystkie buildowe wersje w treści (bez ikon .png) — do wykrycia rozjazdu.
function collectBuildVersions(content) {
  const found = new Set();
  for (const re of PATTERNS) {
    (content.match(re) || []).forEach((s) => {
      const v = (s.match(ONE_VER) || [])[0];
      if (v) found.add(v);
    });
  }
  return found;
}

// ─── Ustal aktualną wersję (z CACHE_VERSION w sw.js) i policz następną ─────────

const swRaw = read("sw.js");
const currentMatch = swRaw.match(new RegExp(`CACHE_VERSION\\s*=\\s*"(${ANY_VER})"`));
const currentVersion = currentMatch ? currentMatch[1] : null;
if (!currentVersion) {
  console.error("❌  Nie znaleziono CACHE_VERSION w sw.js");
  process.exit(1);
}

function fmt(major, minor, patch) {
  return `v${major}.${String(minor).padStart(2, "0")}.${String(patch).padStart(2, "0")}`;
}

function nextVersion(current, arg) {
  if (arg && VERSION_RE.test(arg)) return arg;
  const m = /^v(\d+)\.(\d{2})\.(\d{2})$/.exec(current);
  if (!m) return fmt(1, 0, 0); // pierwsze wydanie po schemacie datowym
  let [major, minor, patch] = [Number(m[1]), Number(m[2]), Number(m[3])];
  if (arg === "major") return fmt(major + 1, 0, 0);
  if (arg === "minor") return fmt(major, minor + 1, 0);
  if (arg === "skok" || arg === "jump") patch = (Math.floor(patch / 5) + 1) * 5;
  else if (!arg) patch += 1;
  else return null;
  if (patch > 99) return fmt(major, minor + 1, 0);
  return fmt(major, minor, patch);
}

const arg = process.argv[2];
const newVersion = nextVersion(currentVersion, arg);
if (!newVersion || !VERSION_RE.test(newVersion)) {
  console.error(`❌  Nieznany argument „${arg}”. Użyj: (nic) | skok | minor | major | vX.YY.ZZ (np. v1.02.00)`);
  process.exit(1);
}

console.log(`\n🔄  Wersja: ${currentVersion} → ${newVersion}`);
if (arg && arg === currentVersion) {
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
console.log(`\nCommit (styl: typ(zakres): krótki tytuł (wersja) + punkty w treści), np.:`);
console.log(`  fix(docs): … (${newVersion})`);
