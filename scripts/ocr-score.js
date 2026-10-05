// ocr-score.js — „jak dobry jest OCR”: strony ze ZNANYM tekstem (wzorzec) → obraz → typowe
// wady skanu/zdjęcia (rozdzielczość, krzywo, obrót, szare tło, wyblakłe, rozmyte, JPEG,
// zdjęcie telefonem z cieniem) → PDF z samym obrazem → konwersja w aplikacji z OCR →
// porównanie tekstu dokumentu ze wzorcem.
//
// Miary (po słowach, z kolejnością — najdłuższy wspólny podciąg słów):
//   trafność  = ile wstawionych słów jest poprawnych (błędne słowo w dokumencie = najgorsze),
//   pełność   = ile słów wzorca trafiło do dokumentu jako tekst (reszta została na obrazie),
//   F1        = średnia harmoniczna; plus CER (błędy znaków na tekście wzorca) i czas.
// Historia: ~/.dwb-ocr-scores.json (porównanie z poprzednim pełnym przebiegiem).
//
// Użycie: npm run ocr:score   |   ONLY=krzywo,zdjecie npm run ocr:score   |   KEEP=1 (PDF-y w $TMPDIR/dwb-ocr-score)

const pw = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const HISTORY = path.join(os.homedir(), ".dwb-ocr-scores.json");
const TMP = path.join(os.tmpdir(), "dwb-ocr-score");

// ── strony wzorcowe (polski tekst „z życia”: pismo, formularz, dwie kolumny, drobny druk) ──
const CSS = `body{margin:0;padding:64px 72px;color:#111;background:#fff} p{margin:0 0 12px} h1,h2{margin:0 0 14px}`;
const PAGES = {
  pismo: `<style>${CSS} body{font-family:"Times New Roman",Times,serif;font-size:16px;line-height:1.35} .r{text-align:right}</style>
<p class="r">Lublin, 14 października 2026 r.</p>
<p><b>Spółdzielnia Mieszkaniowa „Zgoda”</b><br>ul. Żółkiewskiego 12/4<br>20-345 Lublin</p>
<h1 style="font-size:22px">Wezwanie do zapłaty nr 118/2026</h1>
<p>Szanowny Panie, zgodnie z umową najmu z dnia 3 marca 2025 r. uprzejmie przypominamy o zaległości w opłatach za lokal przy ulicy Grunwaldzkiej 7. Łączna kwota należności wynosi 1 284,56 zł i obejmuje czynsz za sierpień oraz wrzesień, a także opłatę za wywóz nieczystości.</p>
<p>Prosimy o uregulowanie płatności w terminie czternastu dni od otrzymania niniejszego pisma na rachunek bankowy: 12 1020 3147 0000 8902 0123 4567. W tytule przelewu należy wpisać numer lokalu oraz numer wezwania.</p>
<p>W przypadku braku wpłaty sprawa zostanie skierowana na drogę postępowania sądowego, co wiąże się z dodatkowymi kosztami. Jeżeli należność została już uregulowana, prosimy uznać niniejsze wezwanie za bezprzedmiotowe.</p>
<p>Z poważaniem<br>Grzegorz Brzęczyszczykiewicz<br>Kierownik Działu Księgowości</p>`,
  formularz: `<style>${CSS} body{font-family:Arial,Helvetica,sans-serif;font-size:14px} table{border-collapse:collapse;width:100%;margin:10px 0 18px} td,th{border:1.5px solid #222;padding:7px 9px;text-align:left;vertical-align:top} th{background:#ddd}</style>
<h2 style="font-size:20px">KARTA ANALIZY ZAGROŻEŃ</h2>
<table><tr><th>Opis zadania</th><td>Wymiana filtrów w prasie numer 3</td></tr><tr><th>Lokalizacja</th><td>Hala produkcyjna B, poziom 0</td></tr><tr><th>Data rozpoczęcia</th><td>21.10.2026</td></tr><tr><th>Numery alarmowe</th><td>112, wewnętrzny 4455</td></tr></table>
<table><tr><th>Kolejność prac</th><th>Potencjalne zagrożenia</th><th>Metody eliminowania</th></tr>
<tr><td>Odłączenie zasilania</td><td>Porażenie prądem</td><td>Blokada wyłącznika, kłódka</td></tr>
<tr><td>Spuszczenie ciśnienia</td><td>Wyciek szlamu pod ciśnieniem</td><td>Okulary ochronne, rękawice</td></tr>
<tr><td>Demontaż płyt filtracyjnych</td><td>Przygniecenie dłoni</td><td>Praca w dwie osoby, podnośnik</td></tr></table>
<p>Opracował: Mateusz Wiśniewski &nbsp; Przejrzał: Józef Łęcki</p>`,
  kolumny: `<style>${CSS} body{font-family:Georgia,serif;font-size:13px;line-height:1.4} .c{columns:2;column-gap:36px;text-align:justify}</style>
<h1 style="font-size:24px">Biuletyn informacyjny gminy</h1>
<div class="c"><p>Rada gminy na październikowej sesji przyjęła uchwałę w sprawie budowy nowej świetlicy wiejskiej. Inwestycja ma kosztować około dwóch milionów złotych, z czego połowę pokryje dotacja z funduszy europejskich. Prace ruszą wiosną przyszłego roku.</p>
<p>Mieszkańcy zgłaszali również potrzebę naprawy chodników przy szkole podstawowej. Wójt zapowiedział, że remont zostanie wykonany jeszcze przed zimą, a w budżecie zabezpieczono na ten cel 180 tysięcy złotych.</p>
<p>Przypominamy o zbiórce elektroodpadów, która odbędzie się w sobotę 25 października na placu przed urzędem. Można oddać stare pralki, lodówki, telewizory i drobny sprzęt kuchenny. Zbiórka jest bezpłatna.</p>
<p>Gminny ośrodek kultury zaprasza dzieci i młodzież na zajęcia teatralne, plastyczne oraz taneczne. Zapisy przyjmuje sekretariat od poniedziałku do piątku w godzinach od dziesiątej do szesnastej.</p></div>`,
  drobny: `<style>${CSS} body{font-family:Arial,Helvetica,sans-serif;font-size:11px;line-height:1.3} h3{font-size:13px;margin:0 0 8px}</style>
<h3>Ogólne warunki ubezpieczenia — fragment</h3>
<p>§ 4. Ubezpieczyciel ponosi odpowiedzialność za szkody powstałe w okresie ubezpieczenia, o ile zostały zgłoszone nie później niż w ciągu siedmiu dni od dnia, w którym ubezpieczony powziął wiadomość o ich wystąpieniu.</p>
<p>§ 5. Z zakresu ubezpieczenia wyłączone są szkody wyrządzone umyślnie albo wskutek rażącego niedbalstwa, a także szkody powstałe w następstwie działań wojennych, zamieszek, strajków oraz aktów terroryzmu.</p>
<p>§ 6. Suma ubezpieczenia stanowi górną granicę odpowiedzialności ubezpieczyciela i ulega zmniejszeniu o kwotę każdego wypłaconego odszkodowania. Składka jest płatna jednorazowo albo w czterech ratach kwartalnych.</p>
<p>§ 7. Spory wynikające z umowy ubezpieczenia można rozstrzygać przed sądem właściwym według przepisów o właściwości ogólnej albo przed sądem właściwym dla miejsca zamieszkania ubezpieczonego.</p>`,
  lista: `<style>${CSS} body{font-family:Arial;font-size:22px;padding:60px} h1{font-size:34px} .box{background:#f39a5b;padding:10px 16px}</style>
<div class="box"><b>KATEDRA AUTOMATYKI</b></div><h1>Lista obecności — lipiec</h1><p>Spóźnienia usprawiedliwione oraz nieusprawiedliwione.</p><p>Zażółć gęślą jaźń.</p>`,
};

// ── wady obrazu (wykonywane na kanwie w przeglądarce) ─────────────────────────────────
// dpi: rozdzielczość „skanera”; rot: obrót kartki (°); paper: szare/żółtawe tło z szumem;
// fade: wyblakły druk; blur: rozmycie (px przy danym dpi); jpeg: jakość; photo: cień i winieta.
const CASES = [
  { id: "czysty-300", page: "pismo", dpi: 300 },
  { id: "czysty-200", page: "formularz", dpi: 200 },
  { id: "niski-150", page: "pismo", dpi: 150, jpeg: 0.7 },
  { id: "krzywo-2", page: "pismo", dpi: 200, rot: 2.2, jpeg: 0.8 },
  { id: "krzywo-5", page: "formularz", dpi: 200, rot: -4.5, jpeg: 0.8 },
  { id: "bokiem-90", page: "kolumny", dpi: 200, rot: 90, jpeg: 0.8 },
  { id: "dogory-180", page: "pismo", dpi: 200, rot: 180, jpeg: 0.8 },
  { id: "szary-szum", page: "kolumny", dpi: 200, paper: true, jpeg: 0.6 },
  { id: "wyblakly", page: "drobny", dpi: 300, fade: 0.38, jpeg: 0.75 },
  { id: "rozmyty", page: "formularz", dpi: 200, blur: 1.1, jpeg: 0.7 },
  { id: "drobny-200", page: "drobny", dpi: 200, jpeg: 0.7 },
  { id: "tabela-183", page: "formularz", dpi: 200, rot: 183, jpeg: 0.8 },
  { id: "ekran-96", page: "lista", dpi: 96, jpeg: 0.85 },
  { id: "zdjecie", page: "pismo", dpi: 170, rot: 1.4, photo: true, paper: true, jpeg: 0.55 },
];

const norm = (s) => s.normalize("NFC").replace(/[„”"«»]/g, '"').replace(/[–—]/g, "-").replace(/ /g, " ");
const words = (s) => norm(s).split(/\s+/).map((w) => w.replace(/^[^\p{L}\p{N}§]+|[^\p{L}\p{N}§%]+$/gu, "")).filter(Boolean);

function lcs(a, b) {
  const m = a.length, n = b.length;
  let prev = new Uint16Array(n + 1), cur = new Uint16Array(n + 1);
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) cur[j] = a[i - 1] === b[j - 1] ? prev[j - 1] + 1 : Math.max(prev[j], cur[j - 1]);
    [prev, cur] = [cur, prev];
  }
  return prev[n];
}
function lev(a, b) {
  const n = b.length;
  let prev = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = cur;
  }
  return prev[n];
}

async function makeScans(browser) {
  fs.mkdirSync(TMP, { recursive: true });
  const ctx = await browser.newContext({ viewport: { width: 794, height: 1123 }, deviceScaleFactor: 300 / 96 });
  const page = await ctx.newPage();
  const truth = {};
  const shots = {};
  for (const [name, html] of Object.entries(PAGES)) {
    await page.setContent(`<!doctype html><html lang="pl"><head><meta charset="utf-8"></head><body>${html}</body></html>`, { waitUntil: "load" });
    await page.evaluate(() => document.fonts.ready);
    truth[name] = await page.evaluate(() => document.body.innerText);
    shots[name] = (await page.screenshot({ type: "png" })).toString("base64");
  }
  const out = {};
  for (const c of CASES) {
    const file = path.join(TMP, `${c.id}.pdf`);
    const img = await page.evaluate(async ({ src, c }) => {
      const im = new Image();
      im.src = "data:image/png;base64," + src;
      await im.decode();
      const W = Math.round(8.27 * c.dpi), H = Math.round(11.69 * c.dpi);
      const side = c.rot === 90 || c.rot === -90;
      const cv = document.createElement("canvas");
      cv.width = side ? H : W;
      cv.height = side ? W : H;
      const x = cv.getContext("2d");
      x.fillStyle = c.paper ? "#ebe6da" : "#fff";
      x.fillRect(0, 0, cv.width, cv.height);
      x.save();
      x.translate(cv.width / 2, cv.height / 2);
      x.rotate(((c.rot || 0) * Math.PI) / 180);
      if (c.blur) x.filter = `blur(${c.blur}px)`;
      if (c.paper) x.globalCompositeOperation = "multiply";
      x.drawImage(im, -W / 2, -H / 2, W, H);
      x.restore();
      x.filter = "none";
      x.globalCompositeOperation = "source-over";
      const d = x.getImageData(0, 0, cv.width, cv.height);
      const p = d.data;
      let seed = 7;
      const rnd = () => ((seed = (seed * 16807) % 2147483647) / 2147483647);
      for (let i = 0; i < p.length; i += 4) {
        const px = (i / 4) % cv.width, py = Math.floor(i / 4 / cv.width);
        for (let k = 0; k < 3; k++) {
          let v = p[i + k];
          if (c.fade) v = 255 - (255 - v) * c.fade;
          if (c.paper) v += (rnd() - 0.5) * 26;
          if (c.photo) {
            // światło z lewej góry, cień od telefonu w prawym dole, winieta
            const fx = px / cv.width, fy = py / cv.height;
            const light = 1 - 0.28 * fx * fy - 0.12 * ((fx - 0.5) ** 2 + (fy - 0.5) ** 2) * 4;
            v *= light;
          }
          p[i + k] = Math.max(0, Math.min(255, v));
        }
      }
      x.putImageData(d, 0, 0);
      return { url: cv.toDataURL(c.jpeg ? "image/jpeg" : "image/png", c.jpeg || undefined), w: cv.width, h: cv.height };
    }, { src: shots[c.page], c });
    const land = img.w > img.h;
    const [pw_, ph] = land ? ["297mm", "210mm"] : ["210mm", "297mm"];
    await page.setContent(`<html><body style="margin:0"><img src="${img.url}" style="width:${pw_};height:${ph};display:block"></body></html>`, { waitUntil: "load" });
    await page.pdf({ path: file, width: pw_, height: ph, margin: { top: 0, bottom: 0, left: 0, right: 0 }, printBackground: true });
    out[c.id] = { file, truth: truth[c.page] };
  }
  // Tekst zamieniony na krzywe (eksport „zamień czcionki na kontury”, Ghostscript -dNoOutputFonts):
  // w PDF nie ma ani jednej litery, same ścieżki — jak zgłoszenie „10. ZAGROŻENIE ATAKIEM…”.
  const { execFileSync } = require("child_process");
  for (const name of ["pismo", "formularz"]) {
    const id = `krzywe-${name}`;
    const src = path.join(TMP, `${id}-src.pdf`), file = path.join(TMP, `${id}.pdf`);
    await page.setContent(`<!doctype html><html lang="pl"><head><meta charset="utf-8"></head><body>${PAGES[name]}</body></html>`, { waitUntil: "load" });
    await page.pdf({ path: src, width: "210mm", height: "297mm", margin: { top: 0, bottom: 0, left: 0, right: 0 }, printBackground: true });
    try {
      execFileSync("gs", ["-q", "-o", file, "-sDEVICE=pdfwrite", "-dNoOutputFonts", src]);
      out[id] = { file, truth: truth[name] };
    } catch (_) { console.log(`⚠️  brak Ghostscripta (gs) — pomijam ${id}`); }
  }
  // Zdjęcia telefonem (JPG 4032×3024 jak z aparatu): kartka pod kątem na ciemnym blacie, cień,
  // kartka na całe zdjęcie — idą przez app/photo-import.js (wykrycie kartki + perspektywa).
  const PHOTOS = [
    { id: "foto-stol", page: "pismo", quad: [[520, 640], [2560, 470], [2730, 3560], [360, 3720]], light: 0.25 },
    { id: "foto-cien", page: "formularz", quad: [[700, 520], [2620, 760], [2480, 3480], [430, 3300]], light: 0.45, shadow: true },
    { id: "foto-cala", page: "kolumny", quad: [[-60, -40], [3080, 40], [3060, 4080], [-40, 4010]], light: 0.2 },
  ];
  for (const ph of PHOTOS) {
    const url = await page.evaluate(async ({ src, ph }) => {
      const im = new Image();
      im.src = "data:image/png;base64," + src;
      await im.decode();
      const pw = im.width, phh = im.height;
      const pc = Object.assign(document.createElement("canvas"), { width: pw, height: phh });
      pc.getContext("2d").drawImage(im, 0, 0);
      const pd = pc.getContext("2d").getImageData(0, 0, pw, phh).data;
      const W = 3024, H = 4032;
      // homografia: piksel zdjęcia → piksel strony (odwrotnie do czworokąta kartki)
      const solve = (src4, dst4) => {
        const A = [], B = [];
        for (let k = 0; k < 4; k++) {
          const [u, v] = src4[k], [x, y] = dst4[k];
          A.push([u, v, 1, 0, 0, 0, -u * x, -v * x]); B.push(x);
          A.push([0, 0, 0, u, v, 1, -u * y, -v * y]); B.push(y);
        }
        for (let i = 0; i < 8; i++) {
          let p = i;
          for (let r = i + 1; r < 8; r++) if (Math.abs(A[r][i]) > Math.abs(A[p][i])) p = r;
          [A[i], A[p]] = [A[p], A[i]]; [B[i], B[p]] = [B[p], B[i]];
          for (let r = 0; r < 8; r++) { if (r === i) continue; const m = A[r][i] / A[i][i]; for (let c = i; c < 8; c++) A[r][c] -= m * A[i][c]; B[r] -= m * B[i]; }
        }
        return B.map((v, i) => v / A[i][i]);
      };
      const h = solve(ph.quad, [[0, 0], [pw, 0], [pw, phh], [0, phh]]);
      const c = Object.assign(document.createElement("canvas"), { width: W, height: H });
      const x = c.getContext("2d");
      const id = x.createImageData(W, H);
      const d = id.data;
      let seed = 11;
      const rnd = () => ((seed = (seed * 16807) % 2147483647) / 2147483647);
      for (let y = 0; y < H; y++) for (let xx = 0; xx < W; xx++) {
        const o = (y * W + xx) * 4;
        const den = h[6] * xx + h[7] * y + 1;
        const u = (h[0] * xx + h[1] * y + h[2]) / den, v = (h[3] * xx + h[4] * y + h[5]) / den;
        let r, g, b;
        if (u >= 0 && v >= 0 && u < pw - 1 && v < phh - 1) {
          const i = ((v | 0) * pw + (u | 0)) * 4;
          r = pd[i] * 0.97; g = pd[i + 1] * 0.95; b = pd[i + 2] * 0.9; // ciepłe światło
        } else {
          const n = 60 + 25 * Math.sin(xx / 37 + y / 211) + rnd() * 18; // blat
          r = n * 1.3; g = n; b = n * 0.7;
        }
        const fx = xx / W, fy = y / H;
        let l = 1 - ph.light * (fx * 0.6 + fy * 0.8);
        if (ph.shadow && fx + fy * 0.4 > 0.85 && fx + fy * 0.4 < 1.15) l *= 0.62; // smuga cienia telefonu
        d[o] = Math.min(255, r * l + (rnd() - 0.5) * 10); d[o + 1] = Math.min(255, g * l + (rnd() - 0.5) * 10); d[o + 2] = Math.min(255, b * l + (rnd() - 0.5) * 10); d[o + 3] = 255;
      }
      x.putImageData(id, 0, 0);
      return c.toDataURL("image/jpeg", 0.85);
    }, { src: shots[ph.page], ph });
    const file = path.join(TMP, `${ph.id}.jpg`);
    fs.writeFileSync(file, Buffer.from(url.split(",")[1], "base64"));
    out[ph.id] = { file, truth: truth[ph.page] };
  }

  // MULTI=n: jeden PDF z n stron-skanów (różne strony wzorcowe) — pomiar szybkości wielu stron
  if (process.env.MULTI) {
    const n = Number(process.env.MULTI) || 6;
    const names = Object.keys(PAGES);
    const imgs = [];
    for (let k = 0; k < n; k++) {
      const url = await page.evaluate(async (src) => {
        const im = new Image();
        im.src = "data:image/png;base64," + src;
        await im.decode();
        const c = Object.assign(document.createElement("canvas"), { width: 1654, height: 2339 });
        c.getContext("2d").drawImage(im, 0, 0, 1654, 2339);
        return c.toDataURL("image/jpeg", 0.8);
      }, shots[names[k % names.length]]);
      imgs.push(url);
    }
    await page.setContent(`<html><body style="margin:0">${imgs.map((u) => `<img src="${u}" style="width:210mm;height:297mm;display:block;break-after:page">`).join("")}</body></html>`, { waitUntil: "load" });
    const file = path.join(TMP, "wiele.pdf");
    await page.pdf({ path: file, width: "210mm", height: "297mm", margin: { top: 0, bottom: 0, left: 0, right: 0 }, printBackground: true });
    out.wiele = { file, truth: Array.from({ length: n }, (_, k) => truth[names[k % names.length]]).join("\n") };
  }
  await ctx.close();
  return out;
}

async function run() {
  const only = process.env.ONLY ? process.env.ONLY.split(",") : null;
  const browser = await pw.chromium.launch({ headless: true });
  const scans = await makeScans(browser);
  const app = process.env.ENGINE === "webkit" ? await pw.webkit.launch({ headless: true }) : browser;
  const context = await app.newContext({ serviceWorkers: "block", viewport: { width: 1400, height: 1000 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  page.on("dialog", (d) => d.accept().catch(() => {})); // „zamknąć niezapisany?” przy sprzątaniu kart
  page.on("pageerror", (e) => console.log("BŁĄD STRONY:", e.message));
  page.on("console", (m) => { if (m.type() === "error") console.log("KONSOLA:", m.text().slice(0, 300)); });
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => window.addEventListener("dwb:pdf-converted", (e) => (window.__conv = e.detail)));
  if (process.env.MASK) await page.evaluate(() => (window.__maskWanted = true));
  if (process.env.LINES) await page.evaluate(() => (window.__linesWanted = true));
  if (process.env.PARALLEL) await page.evaluate((v) => (window.__dwbOcrParallel = v), Number(process.env.PARALLEL));
  if (process.env.INPUT) await page.evaluate(() => (window.__inputWanted = true));
  if (process.env.NOLINES) await page.evaluate(() => (window.__dwbOcrNoLines = true));
  if (process.env.SHARP) await page.evaluate((v) => (window.__dwbOcrSharpen = v), Number(process.env.SHARP));

  const history = fs.existsSync(HISTORY) ? JSON.parse(fs.readFileSync(HISTORY, "utf8")) : { runs: [] };
  const prev = history.runs[history.runs.length - 1]?.results || {};
  const results = {};
  const vector = Object.keys(scans).filter((k) => k.startsWith("krzywe-") || k.startsWith("foto-") || k === "wiele").map((id) => ({ id, ...scans[id] }));
  const extra = process.env.PDF ? [{ id: "plik", file: process.env.PDF, truth: process.env.TRUTH || "" }] : [];
  for (const c of [...CASES, ...vector, ...extra]) {
    if (only && !only.includes(c.id)) continue;
    const { file, truth } = c.file ? c : scans[c.id];
    await page.evaluate(() => { for (const d of dwbOpenDocs.list()) if (!d.active) dwbOpenDocs.closeTab(d.id); }); // limit 12 kart
    await page.evaluate(() => { window.__conv = null; window.__dwbOcrTrace = []; window.__dwbOcrLines = window.__linesWanted ? [] : null; window.__dwbOcrMask = !!window.__maskWanted; window.__dwbOcrInput = !!window.__inputWanted; if (typeof setDirtyState === "function") setDirtyState(false); });
    const t0 = Date.now();
    await page.setInputFiles("#fileInput", file);
    const auto = setInterval(() => page.click("#pdfConvOcrYes", { timeout: 200 }).catch(() => {}), 300);
    try {
      await page.waitForFunction(() => window.__conv, null, { timeout: 240000 });
    } finally {
      clearInterval(auto);
    }
    const ms = Date.now() - t0;
    const text = await page.evaluate(async () => {
      const z = await JSZip.loadAsync(originalFileBytes);
      const doc = await z.file("word/document.xml").async("string");
      const d = new DOMParser().parseFromString(doc, "application/xml");
      const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
      // tabulator / złamanie wiersza = odstęp (inaczej słowa z sąsiednich komórek „sklejały się”)
      // punktor listy Worda bez numeru (np. „§” wzięty za punktor) — z numbering.xml
      const numXml = z.file("word/numbering.xml") ? new DOMParser().parseFromString(await z.file("word/numbering.xml").async("string"), "application/xml") : null;
      const val = (el, tag) => el?.getElementsByTagNameNS(W, tag)[0]?.getAttributeNS(W, "val");
      const bullet = (p) => {
        const numId = val(p, "numId"), ilvl = val(p, "ilvl") || "0";
        if (!numXml || !numId) return "";
        const num = [...numXml.getElementsByTagNameNS(W, "num")].find((n) => n.getAttributeNS(W, "numId") === numId);
        const abs = [...numXml.getElementsByTagNameNS(W, "abstractNum")].find((a) => a.getAttributeNS(W, "abstractNumId") === val(num, "abstractNumId"));
        const lvl = abs && [...abs.getElementsByTagNameNS(W, "lvl")].find((l) => l.getAttributeNS(W, "ilvl") === ilvl);
        const t = val(lvl, "lvlText") || "";
        return t && !t.includes("%") ? t + " " : "";
      };
      // słowa w żółtych fragmentach („do sprawdzenia”)
      const hl = [...d.getElementsByTagNameNS(W, "r")].filter((r) => r.getElementsByTagNameNS(W, "highlight").length).map((r) => [...r.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join("")).join(" ");
      window.__ocrHl = hl;
      return [...d.getElementsByTagNameNS(W, "p")].map((p) => bullet(p) + [...p.getElementsByTagNameNS(W, "*")].map((n) => (n.localName === "t" ? n.textContent : n.localName === "tab" || n.localName === "br" ? " " : "")).join("")).join("\n");
    });
    const trace = await page.evaluate(() => window.__dwbOcrTrace);
    if (process.env.MASK) { const m = await page.evaluate(() => window.__dwbOcrMask); if (typeof m === "string") fs.writeFileSync(path.join(TMP, c.id + "-mask.png"), Buffer.from(m.split(",")[1], "base64")); }
    const hlWords = new Set(words(await page.evaluate(() => window.__ocrHl || "")));
    const g = words(truth), o = words(text);
    const hit = lcs(g, o);
    const prec = o.length ? hit / o.length : 0, rec = g.length ? hit / g.length : 0;
    const f1 = prec + rec ? (2 * prec * rec) / (prec + rec) : 0;
    const gt = g.join(" "), ot = o.join(" ");
    const cer = gt.length ? lev(gt, ot) / gt.length : 1;
    // błędne słowa BEZ żółtego zaznaczenia — najgroźniejsze (użytkownik ich nie sprawdzi)
    const gset = new Set(g);
    const badPlain = o.filter((w) => !gset.has(w) && !hlWords.has(w));
    const r = { badPlain: badPlain.length, hl: hlWords.size, f1: +(f1 * 100).toFixed(1), prec: +(prec * 100).toFixed(1), rec: +(rec * 100).toFixed(1), cer: +(cer * 100).toFixed(1), ms };
    results[c.id] = r;
    const d = prev[c.id] ? ` (${r.f1 - prev[c.id].f1 >= 0 ? "+" : ""}${(r.f1 - prev[c.id].f1).toFixed(1)})` : "";
    console.log(`${c.id.padEnd(12)} F1 ${String(r.f1).padStart(5)}%${d.padEnd(8)} trafność ${String(r.prec).padStart(5)}%  pełność ${String(r.rec).padStart(5)}%  CER ${String(r.cer).padStart(5)}%  żółte ${r.hl}  złe-bez-żółtego ${r.badPlain}${badPlain.length ? ` (${badPlain.join(" ")})` : ""}  ${(ms / 1000).toFixed(1)} s  ${trace.filter((t) => t.skip == null).map((t) => (t.conf != null ? `${t.deg}° ${t.conf}%` : t.textPx ? `linia ${t.textPx}px` : "")).join(" ")}${process.env.DEBUG ? "\n   pominięte: " + trace.filter((t) => t.skip != null).map((t) => `${t.skip}(${t.conf})`).join(" ") : ""}`);
    if (process.env.INPUT) { const m = await page.evaluate(() => window.__dwbOcrInput); if (typeof m === "string") { fs.mkdirSync(TMP, { recursive: true }); fs.writeFileSync(path.join(TMP, `${c.id}-${process.env.ENGINE || "chromium"}-input.png`), Buffer.from(m.split(",")[1], "base64")); } }
    if (process.env.TIMING) console.log("   czas:", JSON.stringify(trace.filter((t) => t.timing).map((t) => t.timing)));
    if (process.env.LEFT) console.log(JSON.stringify(trace.filter((t) => t.left)));
    if (process.env.SEGS) console.log(JSON.stringify(trace.filter((t) => t.segsDbg).map((t) => t.segsDbg)), JSON.stringify(trace.filter((t) => t.segsKept)));
    if (process.env.LINES) console.log((await page.evaluate(() => window.__dwbOcrLines)).join("\n"));
    if (process.env.TEXT) console.log(text.replace(/\n+/g, " ¶ "));
    if (process.env.DEBUG) {
      const miss = g.filter((w) => !o.includes(w)).slice(0, 25), bad = o.filter((w) => !g.includes(w)).slice(0, 25);
      console.log(`   brak: ${miss.join(" ")}\n   złe:  ${bad.join(" ")}`);
    }
  }
  await browser.close();
  if (app !== browser) await app.close();
  const vals = Object.values(results);
  if (!vals.length) return;
  const avg = (k) => (vals.reduce((n, r) => n + r[k], 0) / vals.length).toFixed(1);
  const prevAll = history.runs[history.runs.length - 1]?.avg;
  console.log(`\nŚrednio: F1 ${avg("f1")}%${prevAll ? ` (poprzednio ${prevAll}%)` : ""} · trafność ${avg("prec")}% · pełność ${avg("rec")}% · CER ${avg("cer")}% · czas ${(vals.reduce((n, r) => n + r.ms, 0) / vals.length / 1000).toFixed(1)} s/stronę`);
  if (!only) {
    history.runs.push({ at: new Date().toISOString(), avg: avg("f1"), results });
    fs.writeFileSync(HISTORY, JSON.stringify(history, null, 1));
  }
  if (!process.env.KEEP) fs.rmSync(TMP, { recursive: true, force: true });
}

run().catch((e) => { console.error(e); process.exit(1); });
