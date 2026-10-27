// mail-merge-playwright.js — korespondencja seryjna (2026-10-27).
//
// Dokument z polami {{imie}} {{nazwisko}} {{kwota}} {{termin}} {{brak}} + dane CSV i XLSX →
// ZIP z osobnym .docx na wiersz. Sprawdzamy: dopasowanie pól bez wielkości liter i polskich znaków
// („Imię” = {{imie}}), pusta komórka = puste pole, pole bez kolumny zostaje, nazwy plików z kolumny,
// daty i kwoty z XLSX jak w polskim Excelu (15.10.2026, 1 234,50), oryginał dokumentu nietknięty.
// ENGINE=webkit (Safari/iPad).

const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const JSZip = require("../lib/jszip.min.js");
const XLSX = require("../lib/xlsx.full.min.js");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dwb-mm-"));
const csvPath = path.join(tmp, "dane.csv");
fs.writeFileSync(csvPath, "﻿Imię;Nazwisko;Kwota;Termin\r\nJan;Kowalski;\"1 200,00 zł\";31.10.2026\r\nAnna;\"Nowak; Wiśniewska\";;15.11.2026\r\n;;;\r\nOla;Kowalski;50;\r\n");
// XLSX: data (format krótkiej daty 14) i kwota w formacie #,##0.00
const ws = XLSX.utils.aoa_to_sheet([["imie", "NAZWISKO", "kwota", "termin"], ["Zenon", "Ćwikła", 1234.5, 46310]]);
ws.C2.z = "#,##0.00";
ws.D2.z = "m/d/yy";
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Dane");
const xlsxPath = path.join(tmp, "dane.xlsx");
fs.writeFileSync(xlsxPath, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));

async function zipTexts(download) {
  const file = path.join(tmp, `out-${Date.now()}.zip`);
  await download.saveAs(file);
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const out = {};
  for (const name of Object.keys(zip.files).sort()) {
    const inner = await JSZip.loadAsync(await zip.file(name).async("uint8array"));
    const xml = await inner.file("word/document.xml").async("string");
    const hdrs = Object.keys(inner.files).filter((n) => /^word\/header\d*\.xml$/.test(n));
    out[name + "#hdr"] = (await Promise.all(hdrs.map((h) => inner.file(h).async("string")))).map((x) => [...x.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("")).join("|");
    out[name] = [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");
  }
  return out;
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 }, acceptDownloads: true });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);
  await page.keyboard.type("Szanowny {{imie}} {{nazwisko}}, kwota {{kwota}} do {{termin}}. Pole {{brak}} koniec.");
  await page.waitForTimeout(300);
  // nagłówek z polem — korespondencja wypełnia też nagłówki i stopki
  await page.evaluate(() => applyDocumentEdit({ op: "headerFooter", lang: "pl", total: 1, header: { text: "Dla: {{nazwisko}}", align: "left" }, footer: { text: "", align: "center" }, number: { fmt: null }, firstDifferent: false }));
  await idle(page);

  // panel pól (leniwe skrypty)
  await page.evaluate(async () => {
    const d = document.getElementById("panel-placeholders");
    d.open = true;
    d.dispatchEvent(new Event("toggle"));
    if (typeof ensureLazyFeature === "function") await ensureLazyFeature("placeholders");
  });
  await page.waitForFunction(() => typeof mmRun === "function", null, { timeout: 10000 });

  // ── CSV ──────────────────────────────────────────────────────────────────────
  await page.setInputFiles("#mmFile", csvPath);
  await page.waitForFunction(() => /wierszy|rows/.test(document.getElementById("mmStatus").textContent), null, { timeout: 10000 });
  const status = await page.evaluate(() => document.getElementById("mmStatus").textContent);
  check("status: 3 wiersze (pusty pominięty), 4 z 5 pól, ostrzeżenie o {{brak}}", /3 wierszy/.test(status) && /4 z 5/.test(status) && /\{\{brak\}\}/.test(status), status);
  check("nazwy plików domyślnie z kolumny „Nazwisko”", await page.evaluate(() => document.getElementById("mmNameCol").selectedOptions[0]?.textContent === "Nazwisko"));

  let [dl] = await Promise.all([page.waitForEvent("download", { timeout: 30000 }), page.click("#mmRunBtn")]);
  check("ZIP nazwany od dokumentu", /seryjna \(3\)\.zip$/.test(dl.suggestedFilename()), dl.suggestedFilename());
  let docs = await zipTexts(dl);
  check("nagłówek „Dla: {{nazwisko}}” wypełniony w każdym pliku", docs["001 Kowalski.docx#hdr"]?.includes("Dla: Kowalski") && docs["002 Nowak; Wiśniewska.docx#hdr"]?.includes("Dla: Nowak; Wiśniewska"), JSON.stringify(Object.entries(docs).filter(([k]) => k.endsWith("#hdr"))));
  const names = Object.keys(docs).filter((k) => !k.endsWith("#hdr"));
  check("3 pliki z numerem i nazwiskiem, dubel nazwiska bez nadpisania", JSON.stringify(names) === JSON.stringify(["001 Kowalski.docx", "002 Nowak; Wiśniewska.docx", "003 Kowalski.docx"]), JSON.stringify(names));
  check("wiersz 1 wypełniony", docs[names[0]] === "Szanowny Jan Kowalski, kwota 1 200,00 zł do 31.10.2026. Pole {{brak}} koniec.", docs[names[0]]);
  check("wiersz 2: średnik w cudzysłowie, pusta kwota = puste pole", docs[names[1]] === "Szanowny Anna Nowak; Wiśniewska, kwota  do 15.11.2026. Pole {{brak}} koniec.", docs[names[1]]);
  check("wiersz 3: pusty termin = puste pole", docs[names[2]] === "Szanowny Ola Kowalski, kwota 50 do . Pole {{brak}} koniec.", docs[names[2]]);
  check("dokument w aplikacji nietknięty (pola zostają)", await page.evaluate(() => [...document.querySelectorAll(".docx-editable-p")].some((p) => p.textContent.includes("{{imie}} {{nazwisko}}"))));

  // ── podgląd wiersza w formularzu ────────────────────────────────────────────
  await page.fill("#mmRow", "2");
  await page.click("#mmPreviewBtn");
  await page.waitForTimeout(800);
  const form = await page.evaluate(() => Object.fromEntries([...document.querySelectorAll("#phForm input[data-ph-name]")].map((i) => [i.dataset.phName, i.value])));
  check("podgląd: wiersz 2 w formularzu pól", form.imie === "Anna" && form.termin === "15.11.2026" && form.kwota === "", JSON.stringify(form));

  // ── XLSX ─────────────────────────────────────────────────────────────────────
  await page.setInputFiles("#mmFile", xlsxPath);
  await page.waitForFunction(() => /dane\.xlsx/.test(document.getElementById("mmStatus").textContent), null, { timeout: 20000 });
  [dl] = await Promise.all([page.waitForEvent("download", { timeout: 30000 }), page.click("#mmRunBtn")]);
  docs = await zipTexts(dl);
  const xName = Object.keys(docs).find((k) => !k.endsWith("#hdr"));
  const x = docs[xName] || "";
  check("XLSX: data jak w polskim Excelu, kwota 1 234,50", x === "Szanowny Zenon Ćwikła, kwota 1 234,50 do 15.10.2026. Pole {{brak}} koniec.", JSON.stringify(x));
  check("XLSX: nazwa pliku z kolumny NAZWISKO", xName === "001 Ćwikła.docx", xName);

  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : `  (${r.detail})`}`));
  const bad = results.filter((r) => !r.ok).length;
  console.log(`\n[${ENGINE}] ${bad ? `${bad} z ${results.length} nie przeszło` : `wszystkie ${results.length} przeszły`}`);
  fs.rmSync(tmp, { recursive: true, force: true });
  process.exit(bad ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
