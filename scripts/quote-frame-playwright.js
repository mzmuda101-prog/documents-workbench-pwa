// quote-frame-playwright.js — cytaty i ramki jak w Wordzie (zgłoszenie 2026-10-17: cytat/ramka od AI,
// kursor w nim → lista stylów „Normalny” i nie dało się tego zdjąć).
// Oczekiwany XML po operacjach = to, co zapisuje Word 16.114 (tools/word-oracle/run-quote.js →
// expected-quote.json): nadanie stylu z listy zdejmuje bezpośrednie formatowanie akapitu (ramka, tło,
// wcięcia, wyrównanie, numeracja) i cechy znaków na >50% tekstu; gumka przy kursorze — cały akapit.
// Lista stylów: polskie nazwy stylów wbudowanych („Cytat intensywny”, „Tekst blokowy”) i — jak
// okienko Style w Wordzie („Normalny + Ramka: …”) — pozycja „Normalny + ramka”, więc wybranie
// „Normalny” jest zmianą i działa. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { buildDocx } = require("../tools/word-oracle/build-docx");
const { quoteOps } = require("../tools/word-oracle/app");
const { QUOTE_DOC, QUOTE_OPS } = require("../tools/word-oracle/cases-quote");
const EXPECTED = require("../tools/word-oracle/expected-quote.json");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
// lista stylów z kursorem w akapicie, zanim coś zrobimy (kolejność jak QUOTE_OPS)
const LABELS = ["Cytat", "Cytat intensywny", "Normalny + ramka", "Normalny + ramka", "Normalny + ramka", "Tekst blokowy", "Normalny", "Normalny + tło", "Normalny"];

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: "cytaty.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await buildDocx(QUOTE_DOC) });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p.docx-editable-p, .docx-preview-host p"), null, { timeout: 30000 });
  const a = await quoteOps(page, QUOTE_OPS);
  QUOTE_OPS.forEach((o, i) => check(`lista stylów w akapicie #${o.para}: „${LABELS[i]}”`, a.labels[i] === LABELS[i], a.labels[i]));
  EXPECTED.forEach((w, i) => check(`jak Word #${i}: ${w.style} „${w.runs}”${w.ppr.length ? ` [${w.ppr}]` : ""}`, JSON.stringify(a.paras[i]) === JSON.stringify(w), JSON.stringify(a.paras[i])));
  check("brak błędów strony", errors.length === 0, errors.join(" | "));
  await browser.close();
  const bad = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`));
  console.log(`\n[${ENGINE}] ${results.length - bad.length}/${results.length} OK`);
  process.exit(bad.length ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
