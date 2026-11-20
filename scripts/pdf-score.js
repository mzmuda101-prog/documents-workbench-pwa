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
    // wyróżnienie pól formularza (ramki, tło) to interfejs jak znaczniki granic stron — na wydruku znika
    try { localStorage.setItem("dwb.forms.shade", "0"); } catch (_) { /* nic */ }
  });
  let page = await context.newPage();
  // Każdy plik na świeżej stronie: pomiar chowa elementy do zrzutu i zmienia powiększenie — po dużym
  // pliku następny potrafił się nie skończyć (2026-11-20: 3 min czekania i koniec całego przebiegu).
  const fresh = async () => {
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.addStyleTag({ content: ".dwb-page-break{display:none!important}.toast,#toastHost{display:none!important}" });
    await page.evaluate(() => window.addEventListener("dwb:pdf-converted", (e) => (window.__conv = e.detail)));
  };
  const failed = [];

  const results = {};
  // ta sama nazwa w dwóch folderach (Laboratorium_3.pdf) — z folderem, inaczej wyniki nadpisują się w historii
  const dup = (f) => files.filter((o) => path.basename(o) === path.basename(f)).length > 1;
  for (const file of files) {
    const name = dup(file) ? `${path.basename(path.dirname(file))}/${path.basename(file)}` : path.basename(file);
    const bytes = fs.readFileSync(file);
    try {
    if (page.isClosed()) page = await context.newPage(); // po awarii karty (brak pamięci) — dalej z nową
    await fresh();
    if (process.env.SCORE_DEBUG) console.error(`  [oryginał] ${name}`);
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
        // 1 pt = 4/3 px, ale najwyżej 2000 px na dłuższym boku: skan 2261 × 3329 pt dawał 3000 × 4400 px
        // na stronę i karta padała z braku pamięci (A4 = 1123 px — zwykłe strony bez zmian)
        const vp1 = p.getViewport({ scale: 1 });
        const vp = p.getViewport({ scale: Math.min(4 / 3, 2000 / Math.max(vp1.width, vp1.height)) });
        const c = document.createElement("canvas");
        c.width = Math.round(vp.width);
        c.height = Math.round(vp.height);
        const ctx = c.getContext("2d");
        ctx.fillStyle = "#fff";
        ctx.fillRect(0, 0, c.width, c.height);
        await p.render({ canvasContext: ctx, viewport: vp, annotationMode: pdfjs.AnnotationMode.ENABLE }).promise; // z wyglądem pól formularza — jak w przeglądarce PDF
        pages.push(c.toDataURL("image/png"));
        const tc = await p.getTextContent();
        // wzorzec słów też po kodowaniu T1 z LaTeX-a („Mo»emy” → „Możemy”), jak w konwersji (pdf-extract fixT1Polish)
        const T1 = { "¡": "ą", "¢": "ć", "¦": "ę", "ª": "ł", "«": "ń", "±": "ś", "¹": "ź", "»": "ż" };
        const t1 = (w) => w.replace(/(?<=\p{L})[¡¢¦ª«±¹»]|[¡¢¦ª«±¹»](?=\p{L})/gu, (c) => T1[c]);
        tc.items.forEach((it) => String(it.str).split(/\s+/).forEach((w) => { if (/\p{L}{3,}/u.test(w)) words.add(t1(w).replace(/[.,:;()„”"!?]/g, "")); }));
      }
      try { await (doc.loadingTask || doc).destroy(); } catch (_) { /* nic */ }
      return { pages, words: [...words] };
    }, bytes.toString("base64"));

    if (process.env.SCORE_DEBUG) console.error(`  [konwersja] ${name}`);
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
    // wszystkie kartki: wysoka strona (długa notatka z Notatek, paragon) wychodzi jako kilka kartek A4
    for (const s of secs.slice(0, Math.max(orig.pages.length, 60))) {
      await s.evaluate((el) => el.scrollIntoView({ block: "start" }));
      await page.waitForTimeout(80);
      const box = await s.boundingBox();
      conv.push("data:image/png;base64," + (await page.screenshot({ clip: box, animations: "disabled" })).toString("base64"));
    }
    const text = await page.evaluate(() => [...document.querySelectorAll("section.docx")].map((s) => s.innerText).join(" "));
    await page.evaluate(() => document.querySelectorAll("[data-score-hidden]").forEach((el) => { el.style.visibility = ""; delete el.dataset.scoreHidden; }));

    if (process.env.SCORE_DEBUG) console.error(`  [porównanie] ${name}: ${conv.length} kartek`);
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
      // Konwersja tnie wysoką stronę na kartki A4 (pdf-ocr / pdf-convert). Wtedy kartki tej strony
      // sklejamy w jeden obraz — dawniej strona 1 (cała notatka) szła do porównania z samą kartką 1
      // i długie notatki miały 4–14%, choć wynik był dobry.
      const As = await Promise.all(a.map(load));
      const Bs = await Promise.all(b.map(load));
      let groups = As.map((_, i) => (Bs[i] ? [Bs[i]] : []));
      // Notatka wyeksportowana jako wąski pasek na kartce A4 (Notatki na iPhonie): konwersja odcina puste
      // pole i powiększa pasek na kilka kartek. Oryginał przycinamy wtedy do treści — porównujemy treść.
      // prostokąt z „tuszem” (ciemniejsze niż tło) na obrazie
      // na zmniejszonej kopii (najwyżej 1500 px) — skany bywają 3000 × 4400 px i pełny odczyt pikseli
      // kilku stron naraz wywracał kartę („Target page… closed”)
      const inkBox = (im) => {
        const k = Math.min(1, 1500 / Math.max(im.width, im.height));
        const c = document.createElement("canvas");
        c.width = Math.max(1, Math.round(im.width * k)); c.height = Math.max(1, Math.round(im.height * k));
        const x = c.getContext("2d");
        x.drawImage(im, 0, 0, c.width, c.height);
        const d = x.getImageData(0, 0, c.width, c.height).data;
        let x0 = c.width, y0 = c.height, x1 = -1, y1 = -1;
        for (let y = 0; y < c.height; y++) for (let xx = 0; xx < c.width; xx++) {
          const i = (y * c.width + xx) * 4;
          if (d[i] * 0.3 + d[i + 1] * 0.59 + d[i + 2] * 0.11 < 200) { if (xx < x0) x0 = xx; if (xx > x1) x1 = xx; if (y < y0) y0 = y; if (y > y1) y1 = y; }
        }
        if (x1 < 0) return null;
        const up = (v, max) => Math.min(max, Math.max(0, Math.round(v / k)));
        return { x0: up(x0, im.width - 1), y0: up(y0, im.height - 1), x1: up(x1 + 1, im.width - 1), y1: up(y1 + 1, im.height - 1) };
      };
      const cropToInk = async (im, vertOnly) => {
        const r = inkBox(im);
        if (!r) return im;
        const x0 = vertOnly ? 0 : r.x0, x1 = vertOnly ? im.width - 1 : r.x1;
        const o = document.createElement("canvas");
        o.width = x1 - x0 + 1; o.height = r.y1 - r.y0 + 1;
        o.getContext("2d").drawImage(im, x0, r.y0, o.width, o.height, 0, 0, o.width, o.height);
        return load(o.toDataURL("image/png"));
      };
      if (Bs.length > As.length && Bs[0]) {
        const sheet = Bs[0].height / Bs[0].width;
        const tryGroups = (list) => {
          const n = list.map((im) => Math.max(1, Math.ceil(im.height / im.width / sheet - 0.05)));
          if (n.reduce((x, y) => x + y, 0) !== Bs.length) return null;
          let k = 0;
          return n.map((c) => Bs.slice(k, (k += c)));
        };
        let g = tryGroups(As);
        if (!g) {
          const cropped = [];
          for (const im of As) cropped.push(await cropToInk(im)); // po kolei — pamięć
          // z ilu kartek składa się wynik danej strony: wg proporcji przyciętej treści, resztę dostaje najwyższa
          const want = cropped.map((im) => Math.max(1, Math.round(im.height / im.width / sheet)));
          const diff = Bs.length - want.reduce((x, y) => x + y, 0);
          if (Math.abs(diff) <= Math.max(1, As.length)) {
            const tall = want.indexOf(Math.max(...want));
            want[tall] = Math.max(1, want[tall] + diff);
            if (want.reduce((x, y) => x + y, 0) === Bs.length) {
              let k = 0;
              g = want.map((c) => Bs.slice(k, (k += c)));
              cropped.forEach((im, i) => { if (g[i].length > 1) As[i] = im; });
            }
          }
        }
        if (g) groups = g;
      }
      const stitch = async (list) => {
        if (list.length === 1) return list[0];
        const w = list[0].width, h = list.reduce((s2, im) => s2 + im.height * (w / im.width), 0);
        const c = document.createElement("canvas");
        c.width = w; c.height = Math.round(h);
        const x = c.getContext("2d");
        let y = 0;
        for (const im of list) { const ih = im.height * (w / im.width); x.drawImage(im, 0, y, w, ih); y += ih; }
        return load(c.toDataURL("image/png"));
      };
      for (let i = 0; i < a.length; i++) {
        const A = As[i];
        if (!groups[i]?.length) { out.push({ f1: 0, missing: true }); continue; }
        const split = groups[i].length > 1;
        const W = split ? Math.min(1000, Math.max(A.width, 400)) : A.width;
        const H = split ? Math.round(A.height * (W / A.width)) : A.height;
        const ma = ink(A, W, H);
        let mb;
        if (split) {
          // Pocięta strona: każda kartka (bez pustych pasów u góry/dołu, w wspólnych granicach treści)
          // szuka swojego miejsca w oryginale blisko końca poprzedniej — cięcie zmienia odstęp w tym
          // miejscu, a jedno przesunięcie dla całej strony sumowało te różnice w dryf w dół.
          const sl = [];
          for (const im of groups[i]) sl.push(await cropToInk(im, true));
          let gx0 = Infinity, gx1 = -1;
          for (const im of sl) { const r = inkBox(im); if (r) { gx0 = Math.min(gx0, r.x0); gx1 = Math.max(gx1, r.x1); } }
          mb = new Uint8Array(W * H);
          let yExp = 0;
          for (const im of sl) {
            if (gx1 < 0) break;
            const sw = gx1 - gx0 + 1, sh = Math.max(1, Math.round(im.height * (W / sw)));
            const c = document.createElement("canvas");
            c.width = W; c.height = sh;
            const cx2 = c.getContext("2d");
            cx2.fillStyle = "#fff"; cx2.fillRect(0, 0, W, sh);
            cx2.drawImage(im, gx0, 0, sw, im.height, 0, 0, W, sh);
            const d = cx2.getImageData(0, 0, W, sh).data;
            const m = new Uint8Array(W * sh);
            for (let k = 0; k < W * sh; k++) m[k] = d[k * 4] * 0.3 + d[k * 4 + 1] * 0.59 + d[k * 4 + 2] * 0.11 < 200 ? 1 : 0;
            const fit = (dx, dy, step) => {
              let n = 0;
              for (let y = 0; y < sh; y += step) {
                const ty = y + dy;
                if (ty < 0 || ty >= H) continue;
                for (let x = 0; x < W; x += step) { const tx = x + dx; if (tx >= 0 && tx < W && m[y * W + x] && ma[ty * W + tx]) n++; }
              }
              return n;
            };
            const R = Math.round(sh * 0.25) + 20;
            let bp = [0, yExp], bv = -1;
            for (let dy = yExp - R; dy <= yExp + R; dy += 4) for (let dx = -8; dx <= 8; dx += 4) { const v = fit(dx, dy, 3); if (v > bv) { bv = v; bp = [dx, dy]; } }
            const [px, py] = bp;
            bv = -1;
            for (let dy = py - 3; dy <= py + 3; dy++) for (let dx = px - 3; dx <= px + 3; dx++) { const v = fit(dx, dy, 1); if (v > bv) { bv = v; bp = [dx, dy]; } }
            for (let y = 0; y < sh; y++) {
              const ty = y + bp[1];
              if (ty < 0 || ty >= H) continue;
              for (let x = 0; x < W; x++) { const tx = x + bp[0]; if (tx >= 0 && tx < W && m[y * W + x]) mb[ty * W + tx] = 1; }
            }
            yExp = bp[1] + sh;
          }
        } else mb = ink(await stitch(groups[i]), W, H);
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
        if (!split) {
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
    if (process.env.IMAGES) scores.forEach((s, i) => s.img && fs.writeFileSync(path.join(IMG_DIR, `${name.replace(/\//g, " - ")}-s${i + 1}.png`), Buffer.from(s.img.split(",")[1], "base64")));
    const p = prev[name];
    const delta = (k) => (p && p[k] != null && r[k] != null ? ` (${r[k] - p[k] >= 0 ? "+" : ""}${(r[k] - p[k]).toFixed(1)})` : "");
    console.log(`${name.padEnd(48)} wygląd ${String(r.look).padStart(5)}%${delta("look")}  najsłabsza s.${r.worstPage} ${r.worst}%  słowa z PDF ${r.words == null ? "—" : r.words + "%"}${delta("words")}  ${(ms / 1000).toFixed(1)} s`);
    } catch (e) {
      // jeden zły plik nie przerywa pomiaru całej puli — zostaje na liście „nie udało się”
      failed.push(name);
      console.log(`${name.padEnd(48)} NIE UDAŁO SIĘ: ${String(e?.message || e).split("\n")[0].slice(0, 120)}`);
    }
  }
  await browser.close();
  if (failed.length) console.log(`\nNie udało się (${failed.length}, poza średnią): ${failed.join(", ")}`);
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
