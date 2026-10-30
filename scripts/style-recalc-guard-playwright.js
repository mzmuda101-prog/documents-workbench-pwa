// style-recalc-guard-playwright.js — strażnik pułapki wydajności „odczyt stylu = układ całego panelu”.
//
// 2026-10-29: `.doc-panel { container: docpanel / inline-size }` (pasek w jednym ciągu przy
// szerokim panelu) sprawiał, że każdy getComputedStyle akapitu po zmianie jego stylu wymuszał
// układ CAŁEGO panelu z dokumentem — łatki składu (applyWordLineMetrics) robią to dla każdego
// akapitu, więc otwarcie 100 stron trwało 2,5 s zamiast 0,8 s (Cofnij tak samo). Teraz klasa
// dwb-panel-wide z ResizeObservera (app-frame.js), a .doc-panel ma samo `contain`.
//
// Sprawdza:
//  1) koszt „zmień font-size akapitu + odczytaj styl” w dokumencie ≈ ten sam na kopiach akapitów
//     poza panelem (z zapytaniem kontenerowym nad dokumentem był wielokrotnie droższy),
//  2) żaden przodek dokumentu nie ma container-type (żadne nowe zapytanie kontenerowe),
//  3) pasek: klasa szerokiego panelu przy ≥ 860 px treści panelu, bez niej poniżej.
//
//   node scripts/style-recalc-guard-playwright.js      ENGINE=webkit node scripts/…
// test-impact: bez śledzenia CSS (mierzy koszt stylów — śledzenie reguł by go zawyżało)
const playwright = require("playwright");
const { APP_URL, loadBuiltinSample } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE || "chromium";

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${detail}`}`);
    if (!ok) fails.push(name);
  };

  for (const width of [1600, 700]) {
    const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width, height: 900 } });
    await ctx.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.setItem("dwb-panel-docked-open-v1", "0"); });
    const page = await ctx.newPage();
    const errors = [];
    page.on("pageerror", (e) => errors.push(e.message));
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await loadBuiltinSample(page);
    const r = await page.evaluate(() => {
      const panel = document.querySelector(".doc-panel");
      const article = document.querySelector(".docx-preview-host section.docx article");
      const src = article.querySelector("p");
      for (let i = 0; i < 600; i++) article.appendChild(src.cloneNode(true)); // dokument ~40 stron
      const ps = [...article.querySelectorAll(":scope > p")].slice(-300);
      getComputedStyle(ps[0]).lineHeight;
      const bench = (list, prop, val) => {
        const t = performance.now();
        list.forEach((p) => { const v = p.style[prop]; p.style[prop] = val; getComputedStyle(p).lineHeight; p.style[prop] = v; });
        return performance.now() - t;
      };
      const best = (list) => { const b = () => bench(list, "fontSize", "13px"); return Math.min(b(), b(), b()); };
      const font = best(ps);
      // te same akapity poza panelem (bez przodków dokumentu) — punkt odniesienia tego komputera
      const out = document.createElement("div");
      out.style.cssText = "position:absolute;left:-9999px;top:0;width:600px";
      ps.forEach((p) => out.appendChild(p.cloneNode(true)));
      document.body.appendChild(out);
      const ref = best([...out.children]);
      out.remove();
      const containers = [];
      for (let el = src; el; el = el.parentElement) {
        const cs = getComputedStyle(el);
        if (cs.containerType && cs.containerType !== "normal") containers.push(`${el.tagName}.${el.className}: ${cs.containerType}`);
      }
      const cs = getComputedStyle(panel);
      const content = panel.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
      return { font, ref, containers, wide: panel.classList.contains("dwb-panel-wide"), content };
    });
    if (width === 1600) {
      check(`font-size + odczyt stylu w dokumencie tani jak poza panelem (${Math.round(r.font)} ms vs ${Math.round(r.ref)} ms, 300 akapitów)`, r.font < Math.max(15, r.ref * 3), "odczyt stylu wymusza układ całego panelu — zapytanie kontenerowe albo podobna pułapka nad dokumentem");
      check("brak container-type nad dokumentem", !r.containers.length, r.containers.join("; "));
    }
    check(`${width} px: szeroki pasek = ${r.content >= 860} (treść panelu ${Math.round(r.content)} px)`, r.wide === (r.content >= 860), `klasa dwb-panel-wide = ${r.wide}`);
    check(`${width} px: bez błędów strony`, !errors.length, errors.join("\n"));
    await ctx.close();
  }

  await browser.close();
  if (fails.length) { console.log(`\n${fails.length} błędów`); process.exit(1); }
  console.log("\nOK");
}

run().catch((e) => { console.error(e); process.exit(1); });
