// pdf-score.js — „jak daleko od ideału”: każdy PDF z listy próbek → konwersja w aplikacji →
// porównanie wyglądu każdej strony z oryginałem (obraz strony PDF vs kartka w podglądzie) +
// odsetek słów z PDF obecnych w dokumencie. Wyniki z historią (poprzedni przebieg → zmiana).
//
// Próbki NIE leżą w repo. Lista ścieżek: plik ~/.dwb-pdf-samples.txt (jedna ścieżka na wiersz,
// # = komentarz) albo PDF_SAMPLES=katalog, albo ścieżki w argumentach. Historia wyników:
// ~/.dwb-pdf-scores.json (poza repo — są tam nazwy plików użytkownika).
//
// Zgodność wyglądu: piksele „tuszu” (ciemniejsze niż tło) obu obrazów z tolerancją ~1,5 pt;
// precyzja = ile tuszu wyniku leży tam, gdzie w oryginale; pełność = ile tuszu oryginału
// odtworzono; wynik = średnia harmoniczna (F1). 100% = obraz identyczny w granicach tolerancji.
//
// Użycie: npm run pdf:score   |   npm run pdf:score -- ścieżka.pdf …   |   IMAGES=1 (nakładki)

const pw = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const LIST = path.join(os.homedir(), ".dwb-pdf-samples.txt");
const HISTORY = path.join(os.homedir(), ".dwb-pdf-scores.json");
const IMG_DIR = path.join(os.tmpdir(), "dwb-pdf-score");

function samples() {
  const args = process.argv.slice(2).filter((a) => /\.pdf$/i.test(a));
  if (args.length) return args;
  if (process.env.PDF_SAMPLES) return fs.readdirSync(process.env.PDF_SAMPLES).filter((n) => /\.pdf$/i.test(n)).map((n) => path.join(process.env.PDF_SAMPLES, n));
  if (fs.existsSync(LIST)) return fs.readFileSync(LIST, "utf8").split("\n").map((l) => l.trim()).filter((l) => l && !l.startsWith("#"));
  return [];
}

async function run() {
  const files = samples().filter((f) => {
    if (fs.existsSync(f)) return true;
    console.log(`⚠️  brak pliku (np. odłączony pendrive): ${f}`);
    return false;
  });
  if (!files.length) {
    console.log(`Brak próbek. Wpisz ścieżki PDF do ${LIST} (jedna na wiersz) albo podaj je w argumentach.`);
    process.exit(2);
  }
  if (process.env.IMAGES) fs.mkdirSync(IMG_DIR, { recursive: true });
  const history = fs.existsSync(HISTORY) ? JSON.parse(fs.readFileSync(HISTORY, "utf8")) : { runs: [] };
  const prev = history.runs[history.runs.length - 1]?.results || {};

  const browser = await pw.chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1600, height: 2600 }, deviceScaleFactor: 1 }); // cała kartka A4 mieści się w polu dokumentu
  await context.addInitScript(() => {
    sessionStorage.setItem("introPlayed", "true");
    try { localStorage.setItem("dwb-page-breaks-v1", "0"); } catch (_) { /* nic */ }
  });
  const page = await context.newPage();
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.addStyleTag({ content: ".dwb-page-break{display:none!important}.toast,#toastHost{display:none!important}" });
  await page.evaluate(() => window.addEventListener("dwb:pdf-converted", (e) => (window.__conv = e.detail)));

  const results = {};
  for (const file of files) {
    const name = path.basename(file);
    const bytes = fs.readFileSync(file);
    // 1) oryginał: strony jako obrazy (1 pt = 4/3 px, jak kartka w podglądzie przy 100%)
    const orig = await page.evaluate(async (b64) => {
      await loadLazyScript("app/pdf-import.js");
      const pdfjs = await dwbPdfImport.loadEngine();
      const data = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
      const doc = await pdfjs.getDocument({ data, isEvalSupported: false }).promise;
      const pages = [];
      const words = new Set();
      for (let i = 1; i <= doc.numPages; i++) {
        const p = await doc.getPage(i);
        const vp = p.getViewport({ scale: 4 / 3 });
        const c = document.createElement("canvas");
        c.width = Math.round(vp.width);
        c.height = Math.round(vp.height);
        const ctx = c.getContext("2d");
        ctx.fillStyle = "#fff";
        ctx.fillRect(0, 0, c.width, c.height);
        await p.render({ canvasContext: ctx, viewport: vp, annotationMode: pdfjs.AnnotationMode.ENABLE }).promise; // z wyglądem pól formularza — jak w przeglądarce PDF
        pages.push(c.toDataURL("image/png"));
        const tc = await p.getTextContent();
        tc.items.forEach((it) => String(it.str).split(/\s+/).forEach((w) => { if (/\p{L}{3,}/u.test(w)) words.add(w.replace(/[.,:;()„”"!?]/g, "")); }));
      }
      try { await (doc.loadingTask || doc).destroy(); } catch (_) { /* nic */ }
      return { pages, words: [...words] };
    }, bytes.toString("base64"));

    // 2) konwersja w aplikacji (skan: „Rozpoznaj tekst”)
    await page.evaluate(() => { window.__conv = null; if (typeof setDirtyState === "function") setDirtyState(false); });
    const t0 = Date.now();
    await page.setInputFiles("#fileInput", file);
    const auto = setInterval(() => page.click("#pdfConvOcrYes", { timeout: 200 }).catch(() => {}), 400);
    try {
      await page.waitForFunction(() => window.__conv, null, { timeout: 180000 });
    } finally {
      clearInterval(auto);
    }
    const ms = Date.now() - t0;
    await page.waitForTimeout(1200);
    await page.evaluate(async () => {
      if (document.fonts?.ready) await document.fonts.ready;
      // Widok desktopowy, 100% — kartka w pikselach CSS = rozmiar strony PDF
      const zl = document.getElementById("zoomLevel");
      if (zl && typeof applyZoom === "function") {
        zl.value = "1";
        applyZoom();
      }
    });
    await page.waitForTimeout(300);
    // pasek narzędzi, nagłówek itp. (przyklejone/stałe) nachodziłyby na zrzut kartki
    await page.evaluate(() => {
      document.querySelectorAll("body *").forEach((el) => {
        if (el.closest(".docx-wrapper")) return;
        const pos = getComputedStyle(el).position;
        if (pos === "fixed" || pos === "sticky") {
          el.dataset.scoreHidden = "1";
          el.style.visibility = "hidden";
        }
      });
    });
    const secs = await page.$$("section.docx");
    const conv = [];
    for (const s of secs.slice(0, orig.pages.length)) {
      await s.evaluate((el) => el.scrollIntoView({ block: "start" }));
      await page.waitForTimeout(80);
      const box = await s.boundingBox();
      conv.push("data:image/png;base64," + (await page.screenshot({ clip: box, animations: "disabled" })).toString("base64"));
    }
    const text = await page.evaluate(() => [...document.querySelectorAll("section.docx")].map((s) => s.innerText).join(" "));
    await page.evaluate(() => document.querySelectorAll("[data-score-hidden]").forEach((el) => { el.style.visibility = ""; delete el.dataset.scoreHidden; }));

    // 3) porównanie obrazów (w przeglądarce, kanwa)
    const scores = await page.evaluate(async ({ a, b, wantImg }) => {
      const load = (src) => new Promise((r) => { const im = new Image(); im.onload = () => r(im); im.src = src; });
      const ink = (im, W, H) => {
        const c = document.createElement("canvas");
        c.width = W; c.height = H;
        const x = c.getContext("2d");
        x.fillStyle = "#fff"; x.fillRect(0, 0, W, H);
        x.drawImage(im, 0, 0, W, H);
        const d = x.getImageData(0, 0, W, H).data;
        const m = new Uint8Array(W * H);
        for (let i = 0; i < W * H; i++) m[i] = d[i * 4] * 0.3 + d[i * 4 + 1] * 0.59 + d[i * 4 + 2] * 0.11 < 200 ? 1 : 0;
        return m;
      };
      const dilate = (m, W, H, r) => {
        let cur = m;
        for (let k = 0; k < r; k++) {
          const o = new Uint8Array(W * H);
          for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
            const i = y * W + x;
            if (cur[i] || (x > 0 && cur[i - 1]) || (x < W - 1 && cur[i + 1]) || (y > 0 && cur[i - W]) || (y < H - 1 && cur[i + W])) o[i] = 1;
          }
          cur = o;
        }
        return cur;
      };
      const out = [];
      for (let i = 0; i < a.length; i++) {
        const A = await load(a[i]);
        if (!b[i]) { out.push({ f1: 0, missing: true }); continue; }
        const B = await load(b[i]);
        const W = A.width, H = A.height;
        const ma = ink(A, W, H), mb = ink(B, W, H);
        // Wyrównanie: przesunięcie całej kartki o kilka pikseli (marginesy zrzutu, zaokrąglenia)
        // nie jest błędem układu — szukamy najlepszego dx, dy (najpierw zgrubnie, potem dokładnie).
        const shiftMask = (m, dx, dy) => {
          const o = new Uint8Array(W * H);
          for (let y = 0; y < H; y++) {
            const sy = y - dy;
            if (sy < 0 || sy >= H) continue;
            for (let x = 0; x < W; x++) {
              const sx = x - dx;
              if (sx >= 0 && sx < W && m[sy * W + sx]) o[y * W + x] = 1;
            }
          }
          return o;
        };
        const overlap = (dx, dy, step) => {
          let n = 0;
          for (let y = 0; y < H; y += step) {
            const sy = y - dy;
            if (sy < 0 || sy >= H) continue;
            for (let x = 0; x < W; x += step) {
              const sx = x - dx;
              if (sx >= 0 && sx < W && ma[y * W + x] && mb[sy * W + sx]) n++;
            }
          }
          return n;
        };
        let best = [0, 0], bv = -1;
        for (let dy = -12; dy <= 12; dy += 2) for (let dx = -12; dx <= 12; dx += 2) {
          const v = overlap(dx, dy, 3);
          if (v > bv) { bv = v; best = [dx, dy]; }
        }
        const [cx, cy] = best;
        bv = -1;
        for (let dy = cy - 2; dy <= cy + 2; dy++) for (let dx = cx - 2; dx <= cx + 2; dx++) {
          const v = overlap(dx, dy, 1);
          if (v > bv) { bv = v; best = [dx, dy]; }
        }
        const mbs = shiftMask(mb, best[0], best[1]);
        mb.set(mbs);
        const da = dilate(ma, W, H, 2), db = dilate(mb, W, H, 2);
        let na = 0, nb = 0, hitB = 0, hitA = 0;
        for (let k = 0; k < W * H; k++) {
          if (ma[k]) { na++; if (db[k]) hitA++; }
          if (mb[k]) { nb++; if (da[k]) hitB++; }
        }
        const prec = nb ? hitB / nb : 1, rec = na ? hitA / na : 1;
        const f1 = prec + rec ? (2 * prec * rec) / (prec + rec) : 0;
        let img = null;
        if (wantImg) {
          // nakładka: czarny = zgodne, czerwony = tylko oryginał, niebieski = tylko wynik
          const c = document.createElement("canvas");
          c.width = W; c.height = H;
          const x = c.getContext("2d");
          const id = x.createImageData(W, H);
          for (let k = 0; k < W * H; k++) {
            const o = k * 4;
            const col = ma[k] && mb[k] ? [0, 0, 0] : ma[k] ? (db[k] ? [120, 120, 120] : [230, 30, 30]) : mb[k] ? (da[k] ? [120, 120, 120] : [30, 90, 230]) : [255, 255, 255];
            id.data[o] = col[0]; id.data[o + 1] = col[1]; id.data[o + 2] = col[2]; id.data[o + 3] = 255;
          }
          x.putImageData(id, 0, 0);
          img = c.toDataURL("image/png");
        }
        out.push({ f1, prec, rec, img, shift: best });
      }
      return out;
    }, { a: orig.pages, b: conv, wantImg: !!process.env.IMAGES });

    const norm = (w) => w.toLowerCase().normalize("NFC");
    const have = new Set(text.split(/\s+/).map((w) => norm(w.replace(/[.,:;()„”"!?]/g, ""))));
    const found = orig.words.filter((w) => have.has(norm(w))).length;
    const pagesF1 = scores.map((s) => s.f1);
    const r = {
      pages: orig.pages.length,
      look: Math.round((pagesF1.reduce((a, b) => a + b, 0) / Math.max(1, pagesF1.length)) * 1000) / 10,
      worst: Math.round(Math.min(...pagesF1) * 1000) / 10,
      worstPage: pagesF1.indexOf(Math.min(...pagesF1)) + 1,
      words: orig.words.length ? Math.round((found / orig.words.length) * 1000) / 10 : null,
      ms,
    };
    results[name] = r;
    if (process.env.IMAGES) scores.forEach((s, i) => s.img && fs.writeFileSync(path.join(IMG_DIR, `${name}-s${i + 1}.png`), Buffer.from(s.img.split(",")[1], "base64")));
    const p = prev[name];
    const delta = (k) => (p && p[k] != null && r[k] != null ? ` (${r[k] - p[k] >= 0 ? "+" : ""}${(r[k] - p[k]).toFixed(1)})` : "");
    console.log(`${name.padEnd(48)} wygląd ${String(r.look).padStart(5)}%${delta("look")}  najsłabsza s.${r.worstPage} ${r.worst}%  słowa z PDF ${r.words == null ? "—" : r.words + "%"}${delta("words")}  ${(ms / 1000).toFixed(1)} s`);
  }
  await browser.close();
  const looks = Object.values(results).map((r) => r.look);
  const avg = Math.round((looks.reduce((a, b) => a + b, 0) / looks.length) * 10) / 10;
  const prevAvg = history.runs[history.runs.length - 1]?.avg;
  console.log(`\nŚrednia zgodność wyglądu: ${avg}%${prevAvg != null ? ` (poprzednio ${prevAvg}%)` : ""}`);
  if (process.env.IMAGES) console.log(`Nakładki (czerwony = tylko oryginał, niebieski = tylko wynik): ${IMG_DIR}`);
  // do historii tylko pełny przebieg listy próbek — pomiar wybranych plików nie zmienia punktu odniesienia
  if (process.argv.slice(2).some((x) => /\.pdf$/i.test(x))) return;
  history.runs.push({ at: new Date().toISOString(), avg, results });
  history.runs = history.runs.slice(-50);
  fs.writeFileSync(HISTORY, JSON.stringify(history, null, 1));
}

run().catch((e) => {
  console.error(e);
  process.exit(1);
});
