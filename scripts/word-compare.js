// word-compare.js — wygląd stron jak w Wordzie: Podgląd wydruku aplikacji vs PDF z prawdziwego Worda.
//
//   npm run word:compare                       (pliki z docs/samples)
//   node scripts/word-compare.js a.docx b.docx (własne pliki — zostają na komputerze)
//   FRESH=1 — PDF-y z Worda od nowa (domyślnie z pamięci, dopóki plik się nie zmienił)
//
// Wzorcem jest Word, nie poprzedni render aplikacji: porównanie z samą sobą pilnowałoby tylko,
// żeby nic się nie zmieniło — także gdy wcześniej było źle. Word (Microsoft Word dla Maca,
// AppleScript) zapisuje plik jako PDF; pdf.js rysuje jego strony, aplikacja rysuje Podgląd
// wydruku (te same kartki co druk), oba w tej samej skali (szerokość kartki = 794 px).
// Na każdej stronie: „tusz” (ciemne piksele) Worda i nasz, z tolerancją ±2 px:
//   zgodność = ile tuszu Worda ma nasz tusz w pobliżu i odwrotnie (średnia harmoniczna, %),
//   wiersze  = położenie kolejnych wierszy tekstu (z rzutu poziomego), mediana różnicy w px.
// Różne kroje (Calibri w Wordzie, Carlito u nas — te same szerokości, inny rysunek liter)
// i wygładzanie dają kilka-kilkanaście % różnicy nawet przy idealnym układzie; liczy się
// przesunięcie wierszy, liczba stron i obrazy nakładek.
// Wynik: output/word-compare/raport.html + nakładki PNG (czerwone = tylko Word, niebieskie =
// tylko aplikacja, szare = wspólne). Narzędzie lokalne: pliki nie opuszczają komputera.

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { execFileSync } = require("child_process");
const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

// Plik bez styles.xml / bez rozmiaru strony: Word bierze WŁASNE wartości domyślne (Aptos 12 pt,
// odstęp po akapicie, strona Letter, marginesy 1″) — różnice to nie błąd układu aplikacji.
async function fileNotes(file) {
  try {
    const zip = await JSZip.loadAsync(fs.readFileSync(file));
    const doc = (await zip.file("word/document.xml")?.async("string")) || "";
    const out = [];
    if (!zip.file("word/styles.xml")) out.push("bez styles.xml (domyślne Worda)");
    if (!/<w:pgSz\b/.test(doc)) out.push("bez rozmiaru strony (Word: swój papier)");
    if (!/<w:pgMar\b/.test(doc)) out.push("bez marginesów (Word: 1″)");
    return out;
  } catch (_) { return []; }
}

const ROOT = path.resolve(__dirname, "..");
const OUT = path.join(ROOT, "output/word-compare");
const CACHE = path.join(OUT, "word-pdf");
fs.mkdirSync(CACHE, { recursive: true });
const args = process.argv.slice(2);
const files = args.length ? args : fs.readdirSync(path.join(ROOT, "docs/samples")).filter((f) => f.endsWith(".docx")).map((f) => path.join(ROOT, "docs/samples", f));
const PAGE_W = 794;

// PDF z Worda (AppleScript); pamięć po skrócie pliku
function wordPdf(file) {
  const hash = crypto.createHash("sha1").update(fs.readFileSync(file)).digest("hex").slice(0, 12);
  const base = `${path.basename(file, ".docx")}-${hash}`;
  const pdf = path.join(CACHE, `${base}.pdf`);
  if (fs.existsSync(pdf) && !process.env.FRESH) return pdf;
  // starsze PDF-y tego pliku (inna treść = inny skrót) — precz, żeby pamięć nie rosła
  for (const f of fs.readdirSync(CACHE)) if (f !== path.basename(pdf) && f.startsWith(`${path.basename(file, ".docx")}-`) && /-[0-9a-f]{12}\.(pdf|docx)$/.test(f)) fs.rmSync(path.join(CACHE, f), { force: true });
  // Word (piaskownica macOS) pyta o zgodę na KAŻDĄ nową ścieżkę i pamięta ją dla tej ścieżki —
  // dawniej kopia i PDF miały nazwę ze skrótem, więc pytał przy każdym pliku. Teraz zawsze te same
  // dwa pliki: wejście nadpisywane w miejscu (ten sam plik na dysku), wynik kopiowany do pamięci.
  const docx = path.join(CACHE, "_word-in.docx");
  const out = path.join(CACHE, "_word-out.pdf");
  fs.writeFileSync(docx, fs.readFileSync(file));
  if (fs.existsSync(out)) fs.truncateSync(out, 0);
  const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  try {
    execFileSync("osascript", ["-e", `
      tell application "Microsoft Word"
        open file name "${esc(docx)}"
        tell active document
          save as file name "${esc(out)}" file format format PDF
        end tell
        try
          close active document saving no
        end try
      end tell`], { timeout: 180000, stdio: ["ignore", "pipe", "pipe"] });
  } catch (e) {
    if (!(fs.existsSync(out) && fs.statSync(out).size)) {
      // Word potrafi zawisnąć na pytaniu (naprawa pliku, zgoda na dostęp) — zamykamy wszystko,
      // żeby następne pliki nie padały kolejno (dawniej 7 porażek po jednej)
      try { execFileSync("osascript", ["-e", 'tell application "Microsoft Word" to close every document saving no'], { timeout: 30000, stdio: "ignore" }); } catch (_) { /* Word nie odpowiada */ }
      throw new Error(String(e.stderr || e.message).trim().split("\n").pop());
    }
  }
  if (!(fs.existsSync(out) && fs.statSync(out).size)) throw new Error("Word nie zapisał PDF-a");
  fs.copyFileSync(out, pdf);
  return pdf;
}

async function run() {
  const browser = await pw.chromium.launch({ headless: true });
  const report = [];
  let context = null;
  for (const file of files) {
    const name = path.basename(file, ".docx");
    let pdf;
    try { pdf = wordPdf(file); } catch (e) { report.push({ name, error: `Word: ${String(e.message).split("\n")[0]}` }); console.log(`❌ ${name}: Word — ${e.message}`); continue; }
    // każdy plik w osobnym, czystym kontekście: przy kartach dokumentów (open-docs) i szkicach
    // podgląd wydruku potrafił pokazać POPRZEDNI plik (np. „WYKŁAD” z treścią umowy najmu)
    await context?.close();
    context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1300, height: 1000 } });
    await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
    const page = await context.newPage();
    page.on("dialog", (d) => d.accept().catch(() => {}));
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.setInputFiles("#fileInput", file);
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 60000 });
    // po otwarciu potrafi przyjść drugie (ciche) przerysowanie — jego nakładka „Renderowanie…”
    // trafiała na zrzut kartki (fałszywe „przesunięcie wierszy” o setki px); czekamy na ciszę
    await page.waitForFunction(() => {
      const busy = !document.getElementById("loadingOverlay")?.classList.contains("hidden");
      const now = performance.now();
      if (busy || !window.__wcQuietSince) window.__wcQuietSince = busy ? 0 : now;
      return !busy && now - window.__wcQuietSince > 1200;
    }, null, { timeout: 60000, polling: 100 });
    await page.evaluate(() => { window.__wcQuietSince = 0; });
    // nasze kartki: Podgląd wydruku w 100 %
    await page.evaluate(() => dwbPrint.open());
    await page.waitForSelector(".pp-sheet", { timeout: 30000 });
    await page.evaluate(async () => { const st = dwbPrint._state(); if (st) { st.fit = false; st.zoom = 1; } document.querySelector(".pp-pages").style.setProperty("--pp-zoom", "1"); await document.fonts.ready; });
    // na zrzucie tylko kartka: pasek podglądu, komunikaty i podpowiedzi schowane
    await page.addStyleTag({ content: "#loadingOverlay, .pp-bar, .pp-warn, #toastContainer, .toast-container, .cursor-hint, [class*='hint-bubble'] { visibility: hidden !important; } .pp-zone { display: none !important; }" });
    await page.waitForTimeout(500);
    const sheets = await page.$$(".pp-sheet");
    const ours = [];
    for (const s of sheets) ours.push(`data:image/png;base64,${(await s.screenshot({ animations: "disabled" })).toString("base64")}`);
    await page.evaluate(() => dwbPrint.close());
    // strony Worda (pdf.js aplikacji) i porównanie — w przeglądarce, na kanwach
    const pdfB64 = fs.readFileSync(pdf).toString("base64");
    const res = await page.evaluate(async ({ pdfB64, ours, PAGE_W }) => {
      const pdfjs = await import("/lib/pdfjs/pdf.min.mjs");
      pdfjs.GlobalWorkerOptions.workerSrc = "/lib/pdfjs/pdf.worker.min.mjs";
      const data = Uint8Array.from(atob(pdfB64), (c) => c.charCodeAt(0));
      const doc = await pdfjs.getDocument({ data, cMapUrl: "/lib/pdfjs/cmaps/", cMapPacked: true, wasmUrl: "/lib/pdfjs/wasm/", isEvalSupported: false, verbosity: 0 }).promise;
      const load = (src) => new Promise((ok, bad) => { const im = new Image(); im.onload = () => ok(im); im.onerror = bad; im.src = src; });
      const canvas = (w, h) => Object.assign(document.createElement("canvas"), { width: w, height: h });
      const ink = (cv) => {
        const { width: w, height: h } = cv;
        const d = cv.getContext("2d").getImageData(0, 0, w, h).data;
        const m = new Uint8Array(w * h);
        for (let i = 0; i < w * h; i++) { const y = 0.299 * d[i * 4] + 0.587 * d[i * 4 + 1] + 0.114 * d[i * 4 + 2]; m[i] = y < 170 ? 1 : 0; }
        return m;
      };
      const dilate = (m, w, h, r) => { // kwadrat (2r+1)², rozdzielnie w poziomie i pionie
        const a = new Uint8Array(w * h), b = new Uint8Array(w * h);
        for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { let v = 0; for (let k = -r; k <= r && !v; k++) { const xx = x + k; if (xx >= 0 && xx < w && m[y * w + xx]) v = 1; } a[y * w + x] = v; }
        for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { let v = 0; for (let k = -r; k <= r && !v; k++) { const yy = y + k; if (yy >= 0 && yy < h && a[yy * w + x]) v = 1; } b[y * w + x] = v; }
        return b;
      };
      const rows = (m, w, h) => { // środki wierszy tekstu z rzutu poziomego
        const prof = []; for (let y = 0; y < h; y++) { let n = 0; for (let x = 0; x < w; x++) n += m[y * w + x]; prof.push(n); }
        const out = []; let start = -1;
        for (let y = 0; y <= h; y++) { const on = y < h && prof[y] > 2; if (on && start < 0) start = y; if (!on && start >= 0) { if (y - start >= 4) out.push((start + y) / 2); start = -1; } }
        return out;
      };
      const pages = [];
      const n = Math.max(doc.numPages, ours.length);
      for (let i = 0; i < n; i++) {
        const pg = { page: i + 1 };
        let wc = null, oc = null, W = PAGE_W, H = 0;
        if (i < doc.numPages) {
          const p = await doc.getPage(i + 1);
          const vp1 = p.getViewport({ scale: 1 });
          const vp = p.getViewport({ scale: PAGE_W / vp1.width });
          H = Math.round(vp.height);
          wc = canvas(W, H);
          const g = wc.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, W, H);
          await p.render({ canvasContext: g, viewport: vp }).promise;
        }
        if (i < ours.length) {
          const im = await load(ours[i]);
          if (!H) H = Math.round(im.height * (PAGE_W / im.width));
          oc = canvas(W, H);
          const g = oc.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, W, H);
          g.drawImage(im, 0, 0, W, Math.round(im.height * (PAGE_W / im.width)));
        }
        if (!wc || !oc) { pg.missing = wc ? "brak strony w aplikacji" : "brak strony w Wordzie"; pages.push(pg); continue; }
        const wm = ink(wc), om = ink(oc);
        const wd = dilate(wm, W, H, 2), od = dilate(om, W, H, 2);
        let wIn = 0, wHit = 0, oIn = 0, oHit = 0;
        const ov = canvas(W, H), og = ov.getContext("2d"), img = og.createImageData(W, H);
        for (let k = 0; k < W * H; k++) {
          if (wm[k]) { wIn++; if (od[k]) wHit++; }
          if (om[k]) { oIn++; if (wd[k]) oHit++; }
          let c = [255, 255, 255];
          if (wm[k] && om[k]) c = [120, 120, 120];
          else if (wm[k]) c = od[k] ? [190, 190, 190] : [220, 30, 30];
          else if (om[k]) c = wd[k] ? [190, 190, 190] : [30, 90, 230];
          img.data.set([...c, 255], k * 4);
        }
        og.putImageData(img, 0, 0);
        const rec = wIn ? wHit / wIn : 1, prec = oIn ? oHit / oIn : 1;
        pg.match = Math.round((rec + prec ? (2 * rec * prec) / (rec + prec) : 0) * 1000) / 10;
        const wr = rows(wm, W, H), or = rows(om, W, H);
        const k = Math.min(wr.length, or.length);
        const diffs = []; for (let j = 0; j < k; j++) diffs.push(Math.abs(wr[j] - or[j]));
        diffs.sort((a, b) => a - b);
        pg.rowsWord = wr.length; pg.rowsOurs = or.length;
        pg.rowShift = k ? Math.round(diffs[Math.floor(k / 2)] * 10) / 10 : null;
        pg.rowShiftMax = k ? Math.round(diffs[k - 1] * 10) / 10 : null;
        pg.overlay = ov.toDataURL("image/png");
        pages.push(pg);
      }
      return { wordPages: doc.numPages, ourPages: ours.length, pages };
    }, { pdfB64, ours, PAGE_W });
    res.pages.forEach((pg) => {
      if (!pg.overlay) return;
      const f = `${name}-s${pg.page}.png`;
      fs.writeFileSync(path.join(OUT, f), Buffer.from(pg.overlay.split(",")[1], "base64"));
      pg.overlay = f;
    });
    res.notes = await fileNotes(file);
    report.push({ name, ...res });
    const worst = res.pages.filter((p) => p.match != null).reduce((m, p) => Math.min(m, p.match), 100);
    console.log(`${res.wordPages === res.ourPages ? "✅" : "❌"} ${name}: strony Word ${res.wordPages} / aplikacja ${res.ourPages}; zgodność ${res.pages.map((p) => p.match ?? "—").join(" · ")} %; wiersze (mediana przesunięcia) ${res.pages.map((p) => p.rowShift ?? "—").join(" · ")} px${worst < 80 ? "  ⚠️" : ""}${res.notes.length ? `  (${res.notes.join(", ")})` : ""}`);
  }
  await browser.close();
  // raport HTML (lokalny plik)
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
  const html = `<!doctype html><html lang="pl"><meta charset="utf-8"><title>Aplikacja vs Word</title>
<style>body{font:14px system-ui;margin:24px;color:#1f2937}h2{margin:28px 0 6px}table{border-collapse:collapse}td,th{padding:4px 10px;border-bottom:1px solid #e5e7eb;text-align:left}
.bad{color:#b91c1c;font-weight:600}.pages{display:flex;gap:12px;flex-wrap:wrap}.pages figure{margin:0}.pages img{width:320px;border:1px solid #d1d5db}
.legend span{display:inline-block;width:12px;height:12px;margin:0 4px 0 12px;vertical-align:-1px}</style>
<h1>Podgląd wydruku aplikacji vs PDF z Worda</h1>
<p class="legend">Nakładka:<span style="background:#dc1e1e"></span>tylko Word<span style="background:#1e5ae6"></span>tylko aplikacja<span style="background:#787878"></span>wspólne<span style="background:#bebebe"></span>w tolerancji ±2 px</p>
<table><tr><th>Plik</th><th>Strony Word / aplikacja</th><th>Zgodność stron (%)</th><th>Przesunięcie wierszy — mediana / maks. (px)</th></tr>
${report.map((r) => r.error ? `<tr><td>${esc(r.name)}</td><td class="bad" colspan="3">${esc(r.error)}</td></tr>` : `<tr><td><a href="#${esc(r.name)}">${esc(r.name)}</a>${r.notes?.length ? `<br><small>${esc(r.notes.join(", "))}</small>` : ""}</td><td class="${r.wordPages !== r.ourPages ? "bad" : ""}">${r.wordPages} / ${r.ourPages}</td><td>${r.pages.map((p) => p.match ?? esc(p.missing)).join(" · ")}</td><td>${r.pages.map((p) => p.rowShift == null ? "—" : `${p.rowShift} / ${p.rowShiftMax}`).join(" · ")}</td></tr>`).join("")}
</table>
${report.filter((r) => !r.error).map((r) => `<h2 id="${esc(r.name)}">${esc(r.name)}</h2><div class="pages">${r.pages.map((p) => p.overlay ? `<figure><a href="${esc(p.overlay)}"><img src="${esc(p.overlay)}" alt=""></a><figcaption>str. ${p.page}: ${p.match} % · wiersze ${p.rowsWord}/${p.rowsOurs} · przes. ${p.rowShift} px</figcaption></figure>` : `<figure><figcaption class="bad">str. ${p.page}: ${esc(p.missing)}</figcaption></figure>`).join("")}</div>`).join("")}
</html>`;
  fs.writeFileSync(path.join(OUT, "raport.html"), html);
  // liczby bez obrazów — do porównań przed/po zmianie (SUMMARY=ścieżka, domyślnie summary.json)
  const summary = report.map((r) => ({ name: r.name, error: r.error, wordPages: r.wordPages, ourPages: r.ourPages, pages: (r.pages || []).map((p) => ({ match: p.match ?? null, rowShift: p.rowShift ?? null, missing: p.missing })) }));
  fs.writeFileSync(process.env.SUMMARY || path.join(OUT, "summary.json"), JSON.stringify(summary, null, 1));
  console.log(`\nRaport: ${path.relative(ROOT, path.join(OUT, "raport.html"))}`);
  // inna liczba stron (albo Word nie zrobił PDF-a) = kod błędu — do skryptów i porównań w czasie
  const bad = report.filter((r) => r.error || r.wordPages !== r.ourPages).length;
  if (bad) { console.log(`❌ ${bad} z ${report.length}: inna liczba stron albo błąd Worda`); process.exitCode = 1; }
}

run().catch((e) => { console.error(e); process.exit(1); });
