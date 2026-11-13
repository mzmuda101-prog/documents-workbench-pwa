// tools/break-ui/recent-files.js — najgorsze realne dane dla listy „Ostatnio otwierane” (ekran startowy).
// Narzędzie deweloperskie, nie trafia do dist (build kopiuje tylko assets/lib/docs + pliki aplikacji).
// Dane wchodzą tą samą drogą co prawdziwe: rekordy w IndexedDB „dwb-recent” (doc-comfort.js), bez ruszania kodu.
//
//   node tools/break-ui/recent-files.js                  → zrzuty wszystkich stanów do output/break-ui/
//   STATE=worst HEADED=1 node tools/break-ui/recent-files.js   → otwarte okno do oglądania (demo | worst | one)
//   ENGINE=webkit …                                      → to samo w Safari
// Wymaga serwera testów: npm run serve:test (port 7823).

const path = require("path");
const fs = require("fs");
const pw = require("playwright");

const APP = "http://localhost:7823/";
const OUT = path.resolve(__dirname, "../../output/break-ui");
const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const DAY = 86400000;

// Pola rekordu jak w rememberRecent(): id, name, size, type, openedAt, bytes, handle.
const now = Date.now();
const rec = (name, size, ago) => ({ id: `${name}|${size}`, name, size, type: "", openedAt: now - ago, bytes: new Uint8Array(0), handle: null });

const FIXTURES = {
  demo: [
    rec("Umowa-najmu.docx", 48213, 60e3),
    rec("Oferta 2026.docx", 120400, DAY),
    rec("Skan faktury.pdf", 812000, 3 * DAY),
  ],
  worst: [
    // dwie wersje różniące się dopiero przy końcu — ucięcie na końcu robi z nich dwa identyczne wiersze
    rec("Umowa najmu lokalu mieszkalnego — ul. Grzybowska 12 m. 34, Warszawa (aneks nr 3, wersja ostateczna po uwagach).docx", 15_728_000, 5 * 60e3),
    rec("Umowa najmu lokalu mieszkalnego — ul. Grzybowska 12 m. 34, Warszawa (aneks nr 4, wersja ostateczna po uwagach).docx", 1536, 2 * 3600e3),
    rec("IMG_20250914_183022_HDR_skan_umowy_edytowany_edytowany.pdf", 9_437_184, DAY),
    rec("a.docx", 0, 40 * DAY),
    rec("Đặng Thị Ngọc Hân — CV (Ước tính chi phí).docx", 1023, 200 * DAY),
    rec("نور الهدى عبد الرحمن — عقد العمل.docx", 2_621_440, 400 * DAY), // ponad rok temu
  ],
  one: [rec("Oferta.docx", 1, 60e3)],
};

async function seed(page, list) {
  await page.evaluate(async (items) => {
    const db = await new Promise((res, rej) => {
      const r = indexedDB.open("dwb-recent", 1);
      r.onupgradeneeded = () => r.result.createObjectStore("files", { keyPath: "id" });
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    });
    await new Promise((res, rej) => {
      const tr = db.transaction("files", "readwrite"); const s = tr.objectStore("files");
      s.clear(); items.forEach((x) => s.put({ ...x, bytes: new Uint8Array(0) }));
      tr.oncomplete = res; tr.onerror = () => rej(tr.error);
    });
    db.close();
  }, list);
}

async function open(browser, { width, height = 860, dark = false, lang = "pl", state }) {
  const ctx = await browser.newContext({ viewport: { width, height }, colorScheme: dark ? "dark" : "light", serviceWorkers: "block" });
  await ctx.addInitScript((l) => { try { sessionStorage.setItem("introPlayed", "true"); localStorage.setItem("documents-workbench-lang", l); } catch (_) {} }, lang);
  const page = await ctx.newPage();
  await page.goto(APP);
  await seed(page, FIXTURES[state]);
  await page.reload();
  await page.waitForSelector("#recentFiles:not([hidden]) .recent-item", { timeout: 8000 });
  return { ctx, page };
}

// Pomiar zamiast oka: przelewanie, wysokości, czy pełna nazwa jest gdziekolwiek do przeczytania.
const measure = (page) => page.evaluate(() => {
  const items = [...document.querySelectorAll(".recent-item")];
  const ul = document.querySelector(".recent-list");
  return {
    pageOverflowX: document.documentElement.scrollWidth > innerWidth,
    rows: items.map((li) => {
      const n = li.querySelector(".recent-name"), m = li.querySelector(".recent-meta"), d = li.querySelector(".recent-del");
      return { shown: n.textContent, truncated: n.scrollWidth > n.clientWidth, title: li.querySelector(".recent-open").title || null,
        meta: m.textContent, delW: Math.round(d.getBoundingClientRect().width),
        // wiersz szerszy od listy = ✕ za krawędzią (tak wyglądał błąd z kolumną `auto`); li.scrollWidth tego nie widzi
        delInside: d.getBoundingClientRect().right <= ul.getBoundingClientRect().right + 1,
        nameH: Math.round(n.getBoundingClientRect().height) };
    }),
  };
});


async function main() {
  const browser = await pw[ENGINE].launch({ headless: !process.env.HEADED });
  if (process.env.HEADED) {
    const { page } = await open(browser, { width: 390, state: process.env.STATE || "worst" });
    console.log(`Otwarte (${process.env.STATE || "worst"}). Zamknij okno, żeby skończyć.`);
    await page.waitForEvent("close", { timeout: 0 });
    return browser.close();
  }
  fs.mkdirSync(OUT, { recursive: true });
  const report = {};
  for (const state of ["demo", "worst", "one"]) {
    for (const v of [{ width: 320 }, { width: 390 }, { width: 640, tag: "zoom200" }, { width: 1280 }, { width: 390, dark: true }, { width: 390, lang: "en" }]) {
      const tag = `${state}-${v.width}${v.dark ? "-dark" : ""}${v.lang ? "-" + v.lang : ""}${v.tag ? "-" + v.tag : ""}-${ENGINE}`;
      const { ctx, page } = await open(browser, { ...v, state });
      await page.locator("#recentFiles").screenshot({ path: path.join(OUT, `${tag}.png`) });
      report[tag] = await measure(page);
      await ctx.close();
    }
  }
  fs.writeFileSync(path.join(OUT, `report-${ENGINE}.json`), JSON.stringify(report, null, 2));
  console.log(`Zrzuty i pomiary: ${OUT}`);
  await browser.close();
}

main().catch((e) => { console.error(e); process.exit(1); });
