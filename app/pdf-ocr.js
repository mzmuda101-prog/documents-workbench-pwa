// PDF → DOCX: rozpoznawanie tekstu na skanach (tesseract.js, w workerze, na urządzeniu).
// Strona-skan: renderujemy ją (pdf.js) do kanwy ~250 dpi, tesseract zwraca słowa z położeniem
// i pewnością. Słowa pewne (≥ 70%) zamieniamy na znaki w modelu strony (dalej ten sam układ co
// dla zwykłego PDF), a ich miejsca na obrazie skanu wypełniamy kolorem tła (eraseBoxes) —
// obraz zostaje pod tekstem jako tło: kolory, linie tabel, pieczątki, podpisy odręczne.
(function () {
  "use strict";

  const DIR = "lib/tesseract/";
  let progressCb = null;

  const baseUrl = (rel) => new URL(rel, document.baseURI).href;

  // Ten sam test co wasm-feature-detect (simd): moduł z instrukcją v128.
  function hasSimd() {
    try {
      return WebAssembly.validate(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0, 1, 5, 1, 96, 0, 1, 123, 3, 2, 1, 0, 10, 10, 1, 8, 0, 65, 0, 253, 15, 253, 98, 11]));
    } catch (_) {
      return false;
    }
  }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      if (window.Tesseract) return resolve();
      const s = document.createElement("script");
      s.src = src;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error("Nie udało się wczytać silnika OCR"));
      document.head.appendChild(s);
    });
  }

  // Pula wątków tesseracta. Telefon/tablet: 1 (pamięć kanw — iOS ma ostry limit), komputer: 2
  // (setParallel z pdf-convert). Wątek jest zajmowany TYLKO na czas samego rozpoznawania, więc
  // przygotowanie następnej strony (wątek główny) idzie równolegle z czytaniem poprzedniej.
  const pool = []; // { w, busy, ready }
  const waiters = [];
  let poolSize = 1;

  async function createOne() {
    const v = typeof lazyAssetVersion === "function" ? lazyAssetVersion() : "1";
    await loadScript(baseUrl(DIR + "tesseract.min.js?v=" + encodeURIComponent(v)));
    const core = hasSimd() ? "tesseract-core-simd-lstm.js" : "tesseract-core-lstm.js";
    const w = await window.Tesseract.createWorker("pol", 1 /* OEM.LSTM_ONLY */, {
      workerPath: baseUrl(DIR + "worker.min.js?v=" + encodeURIComponent(v)),
      corePath: baseUrl(DIR + core),
      langPath: baseUrl(DIR + "lang"),
      gzip: true,
      workerBlobURL: false, // CSP: worker-src 'self' (bez blob:)
      cacheMethod: "write", // dane języka w IndexedDB po pierwszym użyciu — lokalnie
      logger: (m) => {
        if (progressCb && m && typeof m.progress === "number") progressCb(m.status, m.progress);
      },
    });
    await w.setParameters({ preserve_interword_spaces: "1" });
    return w;
  }

  function setParallel(n) {
    poolSize = Math.max(1, Math.min(4, n | 0));
  }

  async function acquire(onProgress) {
    progressCb = onProgress || progressCb;
    for (;;) {
      const free = pool.find((p) => p.ready && !p.busy);
      if (free) {
        free.busy = true;
        return free;
      }
      if (pool.length < poolSize) {
        const slot = { w: null, busy: true, ready: false };
        pool.push(slot);
        try {
          slot.w = await createOne();
        } catch (e) {
          pool.splice(pool.indexOf(slot), 1);
          waiters.splice(0).forEach((r) => r());
          throw e;
        }
        slot.ready = true;
        return slot;
      }
      await new Promise((r) => waiters.push(r));
    }
  }

  function release(slot) {
    slot.busy = false;
    const r = waiters.shift();
    if (r) r();
  }

  // Rozpoznanie obrazu na wolnym wątku z puli.
  async function recognizeOn(image, onProgress) {
    const slot = await acquire(onProgress);
    try {
      return (await slot.w.recognize(image, {}, { blocks: true, text: false })).data;
    } finally {
      release(slot);
    }
  }

  // Wczytanie silnika z wyprzedzeniem (pierwszy wątek).
  async function ensure(onProgress) {
    const slot = await acquire(onProgress);
    release(slot);
    return slot.w;
  }

  async function terminate() {
    const all = pool.splice(0);
    waiters.splice(0).forEach((r) => r());
    progressCb = null;
    for (const p of all) {
      try {
        await p.w?.terminate();
      } catch (_) { /* już zamknięty */ }
    }
  }

  function median(arr) {
    if (!arr.length) return 0;
    const s = arr.slice().sort((a, b) => a - b);
    return s[s.length >> 1];
  }

  const OCR_FONT = { id: "ocr", family: "Arial", bold: false, italic: false, mono: false, serif: false, symbolic: false, asc: 0.72, desc: 0.21 };

  // ── przygotowanie obrazu (przed tesseractem) ───────────────────────────────────
  // Tesseract dobrze czyta prosty, równo oświetlony druk. Skany bywają krzywe, obrócone o 90/180°,
  // ze zdjęcia (cień, winieta), na szarym papierze i z liniami tabel, które psują podział na
  // wiersze. Dlatego: (1) kąt strony z profilu wierszy, (2) góra/dół z wydłużeń liter,
  // (3) obraz wyprostowany — ten trafia też jako tło do dokumentu, (4) kopia dla OCR: tło
  // wyrównane do bieli, długie kreski (linie tabel, ramki) usunięte.

  const luma = (r, g, b) => (r * 299 + g * 587 + b * 114) / 1000;

  // Jasność papieru w okolicy każdego piksela: percentyl 90 w blokach B×B, między środkami
  // bloków liniowo. Zwraca funkcję bg(x, y).
  function backgroundMap(gray, w, h, B) {
    const bw = Math.ceil(w / B), bh = Math.ceil(h / B);
    const vals = new Float32Array(bw * bh);
    const hist = new Uint32Array(256);
    for (let by = 0; by < bh; by++) {
      for (let bx = 0; bx < bw; bx++) {
        hist.fill(0);
        let n = 0;
        const x1 = Math.min(w, (bx + 1) * B), y1 = Math.min(h, (by + 1) * B);
        for (let y = by * B; y < y1; y += 2) for (let x = bx * B; x < x1; x += 2) { hist[gray[y * w + x]]++; n++; }
        let k = 255, acc = 0;
        const want = n * 0.1;
        while (k > 0 && acc + hist[k] < want) acc += hist[k--];
        vals[by * bw + bx] = Math.max(40, k);
      }
    }
    // blok bez papieru (cały w tuszu / zdjęciu) bierze jaśniejszego sąsiada — inaczej druk w nim „znika”
    const sm = vals.slice();
    for (let by = 0; by < bh; by++) for (let bx = 0; bx < bw; bx++) {
      let m = vals[by * bw + bx];
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
        const yy = by + dy, xx = bx + dx;
        if (yy >= 0 && xx >= 0 && yy < bh && xx < bw) m = Math.max(m, vals[yy * bw + xx] * 0.97);
      }
      sm[by * bw + bx] = m;
    }
    return (x, y) => {
      const fx = Math.max(0, Math.min(bw - 1.001, x / B - 0.5)), fy = Math.max(0, Math.min(bh - 1.001, y / B - 0.5));
      const ix = fx | 0, iy = fy | 0, ax = fx - ix, ay = fy - iy;
      const i = iy * bw + ix, j = Math.min(i + 1, iy * bw + bw - 1), k = Math.min(i + bw, bw * bh - 1), l = Math.min(k + 1, bw * bh - 1);
      return (sm[i] * (1 - ax) + sm[j] * ax) * (1 - ay) + (sm[k] * (1 - ax) + sm[l] * ax) * ay;
    };
  }

  function grayOf(canvas, maxSide) {
    const f = Math.min(1, maxSide / Math.max(canvas.width, canvas.height));
    const w = Math.max(1, Math.round(canvas.width * f)), h = Math.max(1, Math.round(canvas.height * f));
    let src = canvas;
    if (f < 1) {
      src = document.createElement("canvas");
      src.width = w;
      src.height = h;
      const x = src.getContext("2d");
      x.imageSmoothingQuality = "high";
      x.drawImage(canvas, 0, 0, w, h);
    }
    const d = src.getContext("2d", { willReadFrequently: true }).getImageData(0, 0, w, h).data;
    const g = new Uint8Array(w * h);
    const rgb = [0, 0, 0];
    for (let i = 0, o = 0; i < g.length; i++, o += 4) g[i] = luma(d[o], d[o + 1], d[o + 2]);
    // kolor papieru: średnia z najjaśniejszych 30% pikseli (do wypełnienia rogów po obrocie)
    const order = Array.from({ length: Math.min(g.length, 40000) }, (_, k) => Math.floor((k * g.length) / Math.min(g.length, 40000)));
    order.sort((a, b) => g[b] - g[a]);
    const top = order.slice(0, Math.max(1, Math.floor(order.length * 0.3)));
    for (const i of top) for (let k = 0; k < 3; k++) rgb[k] += d[i * 4 + k];
    return { g, w, h, f, paper: rgb.map((v) => Math.round(v / top.length)) };
  }

  // Punkty „tuszu” (ciemniejsze od papieru w okolicy) — do szukania kąta. Tylko plamy wielkości
  // liter: krawędź kartki na zdjęciu (blat, cień), zdjęcia, pieczątki i długie linie to duże
  // plamy — jedna pionowa krawędź udawała „wiersz” i strona wychodziła obrócona o 90°.
  function inkPoints(a) {
    const { w, h } = a;
    const bg = backgroundMap(a.g, w, h, 24);
    const mask = new Uint8Array(w * h);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      const v = a.g[y * w + x], b = bg(x, y);
      if (v < b * 0.72 && v < b - 35) mask[y * w + x] = 1;
    }
    const maxH = Math.max(12, Math.max(w, h) * 0.045), maxW = Math.max(24, w * 0.2);
    const xs = [], ys = [], heights = [];
    const stack = new Int32Array(w * h);
    const comp = [];
    for (let i = 0; i < mask.length; i++) {
      if (mask[i] !== 1) continue;
      let sp = 0, x0 = w, x1 = 0, y0 = h, y1 = 0;
      comp.length = 0;
      stack[sp++] = i;
      mask[i] = 2;
      while (sp) {
        const j = stack[--sp], x = j % w, y = (j / w) | 0;
        comp.push(j);
        if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
        if (x > 0 && mask[j - 1] === 1) { mask[j - 1] = 2; stack[sp++] = j - 1; }
        if (x < w - 1 && mask[j + 1] === 1) { mask[j + 1] = 2; stack[sp++] = j + 1; }
        if (y > 0 && mask[j - w] === 1) { mask[j - w] = 2; stack[sp++] = j - w; }
        if (y < h - 1 && mask[j + w] === 1) { mask[j + w] = 2; stack[sp++] = j + w; }
      }
      if (y1 - y0 + 1 > maxH || x1 - x0 + 1 > maxW || comp.length < 2) continue;
      if (y1 - y0 >= 3) heights.push(y1 - y0 + 1);
      const step = comp.length > 400 ? 2 : 1;
      for (let k = 0; k < comp.length; k += step) { xs.push(comp[k] % w); ys.push((comp[k] / w) | 0); }
    }
    heights.sort((p, q) => p - q);
    // wysokość wysokich liter (wielkie, b d h k l): percentyl 90 wysokości plam-liter
    return { xs: Float32Array.from(xs), ys: Float32Array.from(ys), letterH: heights.length >= 20 ? heights[Math.floor(heights.length * 0.9)] : 0 };
  }

  // Profil wierszy po obrocie o kąt (stopnie, jak ctx.rotate): y' = x·sin + y·cos.
  function rowProfile(pts, deg, span) {
    const t = (deg * Math.PI) / 180, s = Math.sin(t), c = Math.cos(t);
    const prof = new Float32Array(span * 2 + 2);
    const { xs, ys } = pts;
    for (let i = 0; i < xs.length; i++) prof[(xs[i] * s + ys[i] * c + span) | 0]++;
    return prof;
  }
  // Ostrość profilu: wyraźne wiersze tekstu = duże skoki między sąsiednimi rzędami.
  function sharpness(prof) {
    let sc = 0;
    for (let k = 1; k < prof.length; k++) { const d = prof[k] - prof[k - 1]; sc += d * d; }
    return sc;
  }

  // Kąt strony: najostrzejszy profil wierszy wśród −12…12° oraz 78…102° (strona bokiem).
  function findAngle(pts, a) {
    const span = Math.ceil(Math.hypot(a.w, a.h));
    let best = { deg: 0, sc: -1 };
    const tryDeg = (deg) => {
      const sc = sharpness(rowProfile(pts, deg, span));
      if (sc > best.sc) best = { deg, sc };
      return sc;
    };
    const base0 = tryDeg(0);
    for (const center of [0, 90]) for (let d = -12; d <= 12; d += 0.5) tryDeg(center + d);
    const coarse = best.deg;
    for (let d = -0.5; d <= 0.5; d += 0.05) tryDeg(coarse + d);
    // drobne pochylenie bez wyraźnego zysku zostaw (zdjęcie/obraz bez tekstu, tabela z samymi liniami)
    if (window.__dwbOcrMask) { const c = document.createElement("canvas"); c.width = a.w; c.height = a.h; const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, a.w, a.h); x.fillStyle = "#000"; for (let i = 0; i < pts.xs.length; i++) x.fillRect(pts.xs[i], pts.ys[i], 1, 1); window.__dwbOcrMask = c.toDataURL(); }
    if (Array.isArray(window.__dwbOcrTrace)) window.__dwbOcrTrace.push({ pts: pts.xs.length, base0: base0.toExponential(2), best: best.deg, sc: best.sc.toExponential(2), s90: sharpness(rowProfile(pts, 90, span)).toExponential(2) });
    if (Math.abs(best.deg) < 0.15 || best.sc < base0 * 1.02) return 0;
    return Math.round(best.deg * 100) / 100;
  }

  // Góra/dół: nad pasem małych liter (wysokość „x”) jest więcej tuszu (wielkie litery, b d h k l,
  // kreski ó ś ż) niż pod nim (g j p y, ogonki). Tekst do góry nogami ma odwrotnie.
  function upsideDown(pts, a, deg) {
    const span = Math.ceil(Math.hypot(a.w, a.h));
    const prof = rowProfile(pts, deg, span);
    let max = 0;
    for (const v of prof) max = Math.max(max, v);
    const lo = max * 0.03;
    let above = 0, below = 0, lines = 0;
    for (let k = 0; k < prof.length; ) {
      if (prof[k] <= lo) { k++; continue; }
      let e = k;
      while (e < prof.length && prof[e] > lo) e++;
      const n = e - k;
      if (n >= 4 && n < span / 8) {
        let m = 0;
        for (let i = k; i < e; i++) m = Math.max(m, prof[i]);
        let cs = k, ce = e - 1;
        while (cs < e && prof[cs] < m * 0.45) cs++;
        while (ce > cs && prof[ce] < m * 0.45) ce--;
        let up = 0, dn = 0;
        for (let i = k; i < cs; i++) up += prof[i];
        for (let i = ce + 1; i < e; i++) dn += prof[i];
        if (up + dn > 0) { above += up; below += dn; lines++; }
      }
      k = e;
    }
    return lines >= 3 && below > above * 1.25;
  }

  // Obraz obrócony o kąt (stopnie) wokół środka; 90/270 → zamiana boków. Rogi kolorem papieru.
  // Wielokrotność 90° — kanwą (bez strat), reszta (pochylenie) — interpolacja dwuliniowa w JS:
  // Safari obraca kanwę o „krzywy” kąt bez wygładzania i postrzępione litery psuły OCR
  // (pewność 93% → 75% na tym samym skanie).
  function rotateCanvas(src, deg, paper) {
    const q = ((Math.round(deg / 90) % 4) + 4) % 4;
    const rest = deg - Math.round(deg / 90) * 90;
    let c = src;
    if (q) {
      c = document.createElement("canvas");
      c.width = q % 2 ? src.height : src.width;
      c.height = q % 2 ? src.width : src.height;
      const x = c.getContext("2d", { willReadFrequently: true });
      x.translate(c.width / 2, c.height / 2);
      x.rotate((q * Math.PI) / 2);
      x.drawImage(src, -src.width / 2, -src.height / 2);
    }
    if (Math.abs(rest) < 0.01) return c === src ? copyCanvas(src) : c;
    const w = c.width, h = c.height;
    const sd = c.getContext("2d", { willReadFrequently: true }).getImageData(0, 0, w, h).data;
    const out = document.createElement("canvas");
    out.width = w;
    out.height = h;
    const ox = out.getContext("2d", { willReadFrequently: true });
    const od = ox.createImageData(w, h);
    const d = od.data;
    const t = (rest * Math.PI) / 180, cs = Math.cos(t), sn = Math.sin(t);
    const cx = w / 2, cy = h / 2;
    // dla każdego piksela wyniku: skąd w źródle (obrót odwrotny)
    for (let y = 0; y < h; y++) {
      const dy = y + 0.5 - cy;
      let sx = cs * (0.5 - cx) + sn * dy + cx - 0.5, sy = -sn * (0.5 - cx) + cs * dy + cy - 0.5;
      for (let x = 0, o = y * w * 4; x < w; x++, o += 4, sx += cs, sy -= sn) {
        const x0 = Math.floor(sx), y0 = Math.floor(sy);
        if (x0 < 0 || y0 < 0 || x0 >= w - 1 || y0 >= h - 1) {
          d[o] = paper[0]; d[o + 1] = paper[1]; d[o + 2] = paper[2]; d[o + 3] = 255;
          continue;
        }
        const fx = sx - x0, fy = sy - y0;
        const i = (y0 * w + x0) * 4, j = i + w * 4;
        for (let k = 0; k < 3; k++) {
          const top = sd[i + k] + (sd[i + 4 + k] - sd[i + k]) * fx;
          const bot = sd[j + k] + (sd[j + 4 + k] - sd[j + k]) * fx;
          d[o + k] = top + (bot - top) * fy;
        }
        d[o + 3] = 255;
      }
    }
    ox.putImageData(od, 0, 0);
    if (c !== src) c.width = c.height = 0;
    return out;
  }

  function copyCanvas(src) {
    const c = document.createElement("canvas");
    c.width = src.width;
    c.height = src.height;
    c.getContext("2d", { willReadFrequently: true }).drawImage(src, 0, 0);
    return c;
  }

  // Kopia dla tesseracta: szarość z tłem wyrównanym do bieli (cień, szary papier), bez długich
  // kresek (linie tabel, ramki, podkreślenia pól) — tesseract brał je za tekst albo dzielił
  // przez nie wiersze. Litery są krótsze niż `minRun`, więc zostają.
  function ocrInput(canvas, textPx) {
    const T0i = performance.now();
    const w = canvas.width, h = canvas.height;
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    const id = ctx.getImageData(0, 0, w, h);
    const d = id.data;
    const g = new Uint8Array(w * h);
    for (let i = 0, o = 0; i < g.length; i++, o += 4) g[i] = luma(d[o], d[o + 1], d[o + 2]);
    const bg = backgroundMap(g, w, h, Math.max(32, Math.round(Math.max(w, h) / 60)));
    // rząd po rzędzie: tło liczone raz na kilka pikseli (interpolacja i tak jest gładka)
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x += 8) {
        const b = bg(x, y);
        const k = 255 / Math.max(40, b);
        for (let i = x, e = Math.min(w, x + 8); i < e; i++) g[y * w + i] = Math.min(255, g[y * w + i] * k);
      }
    }
    if (Array.isArray(window.__dwbOcrTrace)) window.__dwbOcrTrace.push({ textPx: Math.round(textPx), w, h });
    const Ti = performance.now();
    const sharp = window.__dwbOcrSharpen ?? 1.5; // zmierzone ocr:score: 0,8 → 93,4%, 1,5 → 94,6%, 2,5 → 93,7%
    if (sharp) unsharp(g, w, h, sharp, Math.max(1, textPx / 30));
    const Ts = performance.now();
    // Długie kreski (linie tabel, ramki, podkreślenia pól): poziome dłuższe niż ~2,5 wysokości
    // wiersza, pionowe niż ~1,5 (litery są niższe). Tylko cienkie — gruby pas (czarny nagłówek
    // z białym napisem) zostaje. Z kopii dla OCR znikają; jako odcinki idą do układu strony
    // (tabela Worda zamiast tekstu rozrzuconego po obrazie).
    // textPx = wysokość wysokich liter (L, b, Ł). Odcinek jest linią, gdy jest dużo dłuższy od
    // liter (poziomy ≥ 4×, pionowy ≥ 3×) albo krótszy, ale styka się z linią w drugim kierunku
    // (krawędź komórki jednowierszowej tabeli). Kreska „L” czy „—” w dużym nagłówku zostaje.
    // jasnoszara kreska (cienka linia po skanowaniu/JPEG) też jest linią — próg łagodniejszy niż dla
    // liter, bo liczą się tylko DŁUGIE ciągi; pas linii jest wybielany cały (z jasnymi brzegami)
    const ink = (v) => v < 170; // 165–175 OK na obu skanach testu; 190 brał krawędź kolorowej ramki za linię
    const maxThick = Math.max(8, Math.round(textPx * 0.45)); // po obrocie i rozmyciu kreska grubieje (WebKit: inne wygładzanie)
    let segs = [];
    if (!window.__dwbOcrNoLines) {
      const hs = findSegs(g, w, h, true, Math.max(30, Math.round(textPx * 1.5)), maxThick, ink);
      const vs = findSegs(g, w, h, false, Math.max(24, Math.round(textPx * 1.3)), maxThick, ink);
      const tol = Math.max(4, Math.round(textPx * 0.25));
      const touch = (a, b) => a.x0 - tol <= b.x1 && b.x0 - tol <= a.x1 && a.y0 - tol <= b.y1 && b.y0 - tol <= a.y1;
      if (Array.isArray(window.__dwbOcrTrace)) window.__dwbOcrTrace.push({ segsDbg: { letter: Math.round(textPx), hs: hs.map((q) => [q.x0, q.y0, q.x1 - q.x0 + 1, q.y1 - q.y0 + 1]), vs: vs.map((q) => [q.x0, q.y0, q.x1 - q.x0 + 1, q.y1 - q.y0 + 1]) } });
      segs = [
        ...hs.filter((sg) => sg.x1 - sg.x0 + 1 >= textPx * 4 || vs.some((v) => touch(sg, v))),
        ...vs.filter((sg) => sg.y1 - sg.y0 + 1 >= textPx * 3 || hs.some((hh) => hh.x1 - hh.x0 + 1 >= textPx * 4 && touch(sg, hh))),
      ];
    }
    if (Array.isArray(window.__dwbOcrTrace)) window.__dwbOcrTrace.push({ segsKept: segs.map((q) => [q.x0, q.y0, q.x1 - q.x0 + 1, q.y1 - q.y0 + 1]) });
    if (Array.isArray(window.__dwbOcrTrace)) window.__dwbOcrTrace.push({ timing: { normalize: Ti - T0i, sharpen: Ts - Ti, segs: performance.now() - Ts } });
    for (const sg of segs) {
      const y0 = Math.max(0, sg.y0 - 1), y1 = Math.min(h - 1, sg.y1 + 1), x0 = Math.max(0, sg.x0 - 1), x1 = Math.min(w - 1, sg.x1 + 1);
      for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) if (g[y * w + x] < 235) g[y * w + x] = 255;
    }
    for (let i = 0, o = 0; i < g.length; i++, o += 4) {
      d[o] = d[o + 1] = d[o + 2] = g[i];
      d[o + 3] = 255;
    }
    const out = document.createElement("canvas");
    out.width = w;
    out.height = h;
    out.getContext("2d").putImageData(id, 0, 0);
    return { canvas: out, segs };
  }

  // Odcinki z ciągów tuszu: ciągi ≥ minLen w kolejnych rzędach (poziome) / kolumnach (pionowe),
  // które się pokrywają, łączą się w jeden odcinek o grubości = liczba rzędów.
  function findSegs(g, w, h, horiz, minLen, maxThick, ink) {
    const L = horiz ? h : w, N = horiz ? w : h;
    const at = horiz ? (r, i) => g[r * w + i] : (r, i) => g[i * w + r];
    let open = [];
    const done = [];
    for (let r = 0; r <= L; r++) {
      const runs = [];
      if (r < L) {
        let run = 0;
        for (let i = 0; i <= N; i++) {
          if (i < N && ink(at(r, i))) run++;
          else {
            if (run >= minLen) runs.push([i - run, i - 1]);
            run = 0;
          }
        }
      }
      const next = [];
      for (const [a, b] of runs) {
        const sg = open.find((o) => !o.taken && Math.min(b, o.b) - Math.max(a, o.a) >= 0.7 * Math.min(b - a, o.b - o.a));
        if (sg) {
          sg.taken = true;
          next.push({ a: Math.min(a, sg.a), b: Math.max(b, sg.b), r0: sg.r0, r1: r });
        } else next.push({ a, b, r0: r, r1: r });
      }
      for (const o of open) if (!o.taken) done.push(o);
      open = next;
    }
    // Wyprostowana (obrócona) kreska rozpada się na główny odcinek i cienkie odpryski tuż obok
    // (wygładzanie krawędzi) — scal równoległe, stykające się kawałki w jeden.
    const merged = [];
    for (const o of done.sort((p, q) => p.r0 - q.r0 || p.a - q.a)) {
      const m = merged.find((q) => o.r0 <= q.r1 + 2 && o.r1 >= q.r0 - 2 && o.a <= q.b + 3 && o.b >= q.a - 3);
      if (m) {
        m.a = Math.min(m.a, o.a); m.b = Math.max(m.b, o.b);
        m.r0 = Math.min(m.r0, o.r0); m.r1 = Math.max(m.r1, o.r1);
      } else merged.push({ ...o });
    }
    return merged.filter((o) => o.r1 - o.r0 + 1 <= maxThick).map((o) => (horiz
      ? { horiz: true, x0: o.a, x1: o.b, y0: o.r0, y1: o.r1 }
      : { horiz: false, x0: o.r0, x1: o.r1, y0: o.a, y1: o.b }));
  }

  // Wyostrzenie (unsharp mask): g + k·(g − rozmyte g). Rozmycie: dwa przebiegi rozmycia
  // pudełkowego (≈ gaussowskie), w JS — ctx.filter nie działa w Safari (WebKit), a wynik ma
  // być ten sam na każdym urządzeniu.
  function unsharp(g, w, h, k, radius) {
    const r = Math.max(1, Math.round(radius)), n = 2 * r + 1;
    let src = new Float32Array(g), dst = new Float32Array(g.length);
    // przebieg poziomy i pionowy, każdy z sumą bieżącą (koszt nie zależy od promienia); 2× = ≈ gauss
    for (let rep = 0; rep < 2; rep++) {
      for (let y = 0; y < h; y++) {
        const row = y * w;
        let acc = src[row] * (r + 1);
        for (let i = 1; i <= r; i++) acc += src[row + Math.min(w - 1, i)];
        for (let x = 0; x < w; x++) {
          dst[row + x] = acc / n;
          acc += src[row + Math.min(w - 1, x + r + 1)] - src[row + Math.max(0, x - r)];
        }
      }
      [src, dst] = [dst, src];
      for (let x = 0; x < w; x++) {
        let acc = src[x] * (r + 1);
        for (let i = 1; i <= r; i++) acc += src[Math.min(h - 1, i) * w + x];
        for (let y = 0; y < h; y++) {
          dst[y * w + x] = acc / n;
          acc += src[Math.min(h - 1, y + r + 1) * w + x] - src[Math.max(0, y - r) * w + x];
        }
      }
      [src, dst] = [dst, src];
    }
    for (let i = 0; i < g.length; i++) {
      const v = g[i] + k * (g[i] - src[i]);
      g[i] = v < 0 ? 0 : v > 255 ? 255 : v;
    }
  }


  // Wysokość typowej linijki tekstu (px) z profilu — do progu „długiej kreski”.
  function typicalLineHeight(pts, a, deg) {
    const span = Math.ceil(Math.hypot(a.w, a.h));
    const prof = rowProfile(pts, deg, span);
    let max = 0;
    for (const v of prof) max = Math.max(max, v);
    const hs = [];
    for (let k = 0; k < prof.length; ) {
      if (prof[k] <= max * 0.03) { k++; continue; }
      let e = k;
      while (e < prof.length && prof[e] > max * 0.03) e++;
      if (e - k >= 3) hs.push(e - k);
      k = e;
    }
    hs.sort((x, y) => x - y);
    return (hs[hs.length >> 1] || 12) / a.f;
  }

  // Cała obróbka: kąt + góra/dół → obraz wyprostowany (tło dokumentu) + kopia dla OCR.
  function prepare(canvas, opts = {}) {
    const T = [performance.now()];
    const a = grayOf(canvas, 1400);
    T.push(performance.now());
    const pts = inkPoints(a);
    T.push(performance.now());
    let deg = 0;
    if (opts.straighten !== false && pts.xs.length > 500) {
      deg = findAngle(pts, a);
      if (upsideDown(pts, a, deg)) deg += 180;
    }
    deg = ((deg % 360) + 360) % 360;
    if (deg > 180) deg -= 360;
    T.push(performance.now());
    const upright = deg ? rotateCanvas(canvas, deg, a.paper) : canvas;
    T.push(performance.now());
    // wysokość wysokich liter w pikselach pełnej rozdzielczości (próg „to już linia, nie litera”)
    const letterH = pts.letterH ? pts.letterH / a.f : typicalLineHeight(pts, a, deg) * 0.75;
    const inp = ocrInput(upright, letterH);
    T.push(performance.now());
    if (Array.isArray(window.__dwbOcrTrace)) window.__dwbOcrTrace.push({ timing: { gray: T[1] - T[0], ink: T[2] - T[1], angle: T[3] - T[2], rotate: T[4] - T[3], input: T[5] - T[4] } });
    return { deg, upright, ocr: inp.canvas, segs: inp.segs, paper: a.paper, letterH };
  }

  // Kolor tekstu słowa: średnia najciemniejszych pikseli w jego prostokącie.
  function inkColor(data, W, box) {
    const px = [];
    for (let y = Math.max(0, box.y0 | 0); y < Math.min(box.y1, data.length / 4 / W); y += 2) {
      for (let x = Math.max(0, box.x0 | 0); x < Math.min(box.x1, W); x += 2) {
        const o = (y * W + x) * 4;
        px.push([data[o] + data[o + 1] + data[o + 2], data[o], data[o + 1], data[o + 2]]);
      }
    }
    if (!px.length) return "000000";
    px.sort((a, b) => a[0] - b[0]);
    const take = px.slice(0, Math.max(1, Math.floor(px.length * 0.15)));
    const avg = [1, 2, 3].map((k) => Math.round(take.reduce((n, p) => n + p[k], 0) / take.length));
    if (avg[0] + avg[1] + avg[2] < 200 && Math.max(...avg) - Math.min(...avg) < 40) return "000000";
    return avg.map((v) => v.toString(16).padStart(2, "0")).join("");
  }

  // „§” nie ma w polskim modelu tesseracta — wychodzi jako „8”, „$”, „S” albo „&” z niską
  // pewnością, a numer obok („4.”, „12”, „3a”) jest czytany dobrze. Pojedynczy taki znak tuż
  // przed numerem = paragraf (pisma urzędowe, umowy, regulaminy).
  function fixParagraphSign(words) {
    for (let i = 0; i < words.length - 1; i++) {
      const w = words[i], next = words[i + 1];
      const t = String(w.text || "").trim();
      if (!/^[8$S&§]$/.test(t) || !/^\d{1,3}[a-z]?[.,:]?$/.test(String(next.text || "").trim())) continue;
      if (w.confidence >= (i === 0 ? 90 : 85) && t !== "§") continue;
      w.text = "§";
      w.confidence = Math.max(w.confidence, next.confidence);
      if (w.symbols?.[0]) w.symbols[0].text = "§";
    }
  }

  /**
   * Rozpoznanie jednej strony.
   * @returns { glyphs, eraseBoxes, words, chars } — współrzędne w pt (jak extractPage)
   */
  async function recognizePage(page, raw, opts = {}) {
    if (opts.signal?.cancelled) throw new (window.DWPdf?.CancelledError || Error)("cancelled");
    const vp1 = page.getViewport({ scale: 1 });
    const scale = Math.min(250 / 72, 3600 / Math.max(vp1.width, vp1.height));
    const vp = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(vp.width);
    canvas.height = Math.ceil(vp.height);
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport: vp, annotationMode: 0 }).promise;
    if (opts.signal?.cancelled) throw new (window.DWPdf?.CancelledError || Error)("cancelled");
    // prosto, do góry, tło wyrównane, bez linii tabel (patrz prepare)
    const tm0 = performance.now();
    let prep = prepare(canvas, { straighten: opts.straighten });
    const tm1 = performance.now();
    let data = await recognizeOn(prep.ocr, opts.onProgress);
    if (Array.isArray(window.__dwbOcrTrace)) window.__dwbOcrTrace.push({ timing: { prep: Math.round(tm1 - tm0), ocr: Math.round(performance.now() - tm1) } });
    // Słaba pewność = może jednak do góry nogami (mało wierszy, sam nagłówek) — druga próba
    // obrócona o 180°; zostaje lepsza.
    if (opts.straighten !== false && (data.confidence || 0) < 55) {
      if (opts.signal?.cancelled) throw new (window.DWPdf?.CancelledError || Error)("cancelled");
      const flipped = rotateCanvas(prep.upright, 180, prep.paper);
      const inp = ocrInput(flipped, (prep.letterH || 12 * scale));
      const alt = { deg: prep.deg > 0 ? prep.deg - 180 : prep.deg + 180, upright: flipped, ocr: inp.canvas, segs: inp.segs, paper: prep.paper };
      const data2 = await recognizeOn(alt.ocr, opts.onProgress);
      if ((data2.confidence || 0) > (data.confidence || 0) + 10) { prep = alt; data = data2; }
    }
    if (opts.signal?.cancelled) throw new (window.DWPdf?.CancelledError || Error)("cancelled");
    if (window.__dwbOcrInput === true) window.__dwbOcrInput = prep.ocr.toDataURL("image/png"); // miernik: co widzi tesseract
    if (Array.isArray(window.__dwbOcrTrace)) window.__dwbOcrTrace.push({ deg: prep.deg, conf: Math.round(data.confidence || 0) }); // miernik ocr:score
    const up = prep.upright;
    const pixels = up.getContext("2d", { willReadFrequently: true }).getImageData(0, 0, up.width, up.height).data;
    const glyphs = [];
    const eraseBoxes = [];
    const wordBoxes = [];
    let words = 0, kept = 0, order = 0, paraNo = 0;
    // Progi pewności (decyzja Mateusza 2026-10-05, zastępuje 92% z 2026-10-02; pomiar ocr:score:
    // wśród słów 85–91% nie było ani jednego błędnego):
    //   ≥ 85%   → zwykły tekst,
    //   70–84%  → tekst z żółtym zaznaczeniem „do sprawdzenia” (z obrazu znika — bez nakładek),
    //   < 70%   → zostaje na obrazie (pismo odręczne, plamy, nieczytelne) i liczy się „do poprawy”.
    const minConf = opts.minConfidence ?? 85;
    const doubtConf = opts.doubtConfidence ?? 70;
    let skipped = 0, doubtful = 0;
    // Wysokość słowa w em (Arial): góra wg najwyższego znaku, dół wg wydłużeń dolnych.
    const wordEm = (text) => {
      let top = 0, bottom = 0;
      for (const ch of text) {
        if (/[ĆŃÓŚŹŻĄĘŁÁÉÍÚÜÖÄ]/.test(ch) && /[ĆŃÓŚŹŻÁÉÍÚ]/.test(ch)) top = Math.max(top, 0.92);
        else if (/[A-ZĄĘŁ0-9]/.test(ch)) top = Math.max(top, 0.716);
        else if (/[bdfhklłtćńóśźżi j!?]/.test(ch)) top = Math.max(top, 0.73);
        else if (/[\p{Ll}]/u.test(ch)) top = Math.max(top, 0.519);
        if (/[gjpqyąęç,;]/.test(ch)) bottom = 0.21;
      }
      return top ? top + bottom : 0;
    };
    // 1. przebieg: słowa pewne (druk); pismo odręczne ma zwykle < 80% — zostaje na obrazie
    const accepted = [];
    for (const block of data.blocks || []) {
      for (const para of block.paragraphs || []) {
        paraNo++;
        for (const line of para.lines || []) {
          fixParagraphSign(line.words || []);
          if (window.__dwbOcrLines) window.__dwbOcrLines.push((line.words || []).map((wd) => `${wd.text}(${Math.round(wd.confidence)})`).join(" "));
          for (const word of line.words || []) {
            words++;
            const text = String(word.text || "").trim();
            if (!text || !/[\p{L}\p{N}§]/u.test(text)) continue;
            const letters = [...text].filter((c) => /[\p{L}\p{N}]/u.test(c)).length;
            const ok = word.confidence >= minConf;
            // wątpliwe: tylko coś, co może być słowem — jednoliterowe tylko polskie spójniki/cyfry
            // (pojedyncze „U”, „R” z krawędzi kartki czy kreski to zwykle śmieci)
            const doubt = !ok && word.confidence >= doubtConf && (letters >= 2 || /^[aiouwzAIOUWZ0-9§]$/u.test(text));
            if (!ok && !doubt) {
              if (Array.isArray(window.__dwbOcrTrace)) window.__dwbOcrTrace.push({ skip: text, conf: Math.round(word.confidence) });
              if (letters >= 2) skipped++;
              continue;
            }
            accepted.push({ word, line, text, paraNo, doubt });
          }
        }
      }
    }
    // wysokość słów druku na stronie — słowo dużo wyższe od mediany to raczej odręczne
    const hs = accepted.map((a) => a.word.bbox.y1 - a.word.bbox.y0).sort((x, y) => x - y);
    const medH = hs.length ? hs[hs.length >> 1] : 0;
    const byLine = new Map();
    for (const a of accepted) {
      const h = a.word.bbox.y1 - a.word.bbox.y0;
      if (medH && h > medH * 2.6 && a.word.confidence < 95) {
        skipped++;
        continue;
      }
      if (!byLine.has(a.line)) byLine.set(a.line, []);
      byLine.get(a.line).push(a);
    }
    for (const [line, list] of byLine) {
      // rozmiar: z wysokości pojedynczych znaków (ramka słowa łapie czasem kreskę podpisu)
      const ests = [];
      for (const a of list) {
        const syms = (a.word.symbols || []).filter((sy) => sy.bbox && /[\p{L}\p{N}]/u.test(sy.text || ""));
        const per = syms.map((sy) => {
          const em = wordEm(sy.text);
          return em ? (sy.bbox.y1 - sy.bbox.y0) / em / scale : 0;
        }).filter((v) => v > 2);
        // dolna mediana: zawyżają tylko znaki sklejone z kreską podpisu, zaniżać nie ma czym
        if (per.length) ests.push(per.sort((x, y) => x - y)[(per.length - 1) >> 1]);
        else {
          const em = wordEm(a.text);
          if (em) ests.push((a.word.bbox.y1 - a.word.bbox.y0) / em / scale);
        }
      }
      const size = Math.max(4, Math.min(72, (ests.length ? ests.sort((x, y) => x - y)[(ests.length - 1) >> 1] : 0) || 10));
      const lb = line.baseline;
      let prevEnd = null;
      for (const a of list) {
        const word = a.word, wb = word.bbox;
        kept++;
        // linia bazowa: dół liter bez wydłużeń dolnych (najpewniejsze); inaczej odcinek bazowy
        // wiersza w miejscu słowa (pochylone skany) — ten bywa skrzywiony przez podpis obok
        let y;
        const baseSyms = (word.symbols || []).filter((sy) => sy.bbox && /[\p{L}\p{N}]/u.test(sy.text || "") && !/[gjpqyąęç,;Q]/.test(sy.text));
        if (baseSyms.length) {
          const bots = baseSyms.map((sy) => sy.bbox.y1).sort((p, q) => p - q);
          y = bots[(bots.length - 1) >> 1] / scale;
        } else if (lb && lb.has_baseline !== false && lb.x1 > lb.x0) {
          const cx = (wb.x0 + wb.x1) / 2;
          y = (lb.y0 + (lb.y1 - lb.y0) * (cx - lb.x0) / (lb.x1 - lb.x0)) / scale;
        } else y = (wb.y1 - (/[gjpqyąę,]/.test(a.text) ? 0.21 * size * scale : 0)) / scale;
        const color = inkColor(pixels, up.width, wb);
        const chars = [...a.text];
        if (a.doubt) doubtful++;
        const mk = (u, x, x1) => ({ u, x, x1, adv: x1 - x, y, size, asc: size * 0.72, desc: size * 0.21, angle: 0, font: OCR_FONT, isSpace: false, color, bold: false, italic: false, order: order++, ocr: true, ocrPara: a.paraNo, doubt: a.doubt || undefined });
        // Litery słowa jedna przy drugiej (szerokości jak w Arialu, dopasowane do ramki słowa) —
        // ramki pojedynczych znaków z OCR to ramki „tuszu”: przerwy między nimi udawałyby spacje.
        const x0 = wb.x0 / scale, x1 = wb.x1 / scale;
        // Tesseract wie, gdzie kończy się słowo — jawna spacja między słowami wiersza (z samej
        // przerwy przy niskiej rozdzielczości wychodziło „zostałajuż”). Duża przerwa = tabulator
        // (decyduje układ strony), więc tam spacji nie wstawiamy.
        if (prevEnd != null && x0 - prevEnd < size * 1.2) {
          glyphs.push({ ...mk(" ", prevEnd, Math.max(prevEnd, x0)), isSpace: true, doubt: undefined });
          // ramka poprzedniego słowa aż do tego — ramki jednoliterowych słów bywają za wąskie
          // (LSTM: „W” miało pół szerokości), a kształty liter (tekst-krzywe) mają znikać w całości
          if (wordBoxes.length) wordBoxes[wordBoxes.length - 1].x1 = Math.max(wordBoxes[wordBoxes.length - 1].x1, x0);
        }
        prevEnd = x1;
        const ws = chars.map((ch) => (window.DWPdf?.helvWidth ? window.DWPdf.helvWidth(ch) : 0.556));
        const total = ws.reduce((n, v) => n + v, 0) || 1;
        let x = x0;
        chars.forEach((ch, k) => {
          const w = ((x1 - x0) * ws[k]) / total;
          glyphs.push(mk(ch, x, x + w));
          x += w;
        });
        const pad = 2 / scale;
        eraseBoxes.push({ x0: wb.x0 / scale - pad, y0: wb.y0 / scale - pad, x1: wb.x1 / scale + pad, y1: wb.y1 / scale + pad });
        wordBoxes.push({ x0: wb.x0 / scale, y0: wb.y0 / scale, x1: wb.x1 / scale, y1: wb.y1 / scale });
      }
    }
    // Strona była krzywa/obrócona: współrzędne są w układzie WYPROSTOWANEJ strony (pt), a obraz
    // skanu w dokumencie trzeba obrócić tak samo (pdf-convert → pdf-import encodeImage).
    // Linie tabel/ramek ze skanu: odcinki dla układu strony (pt, środek kreski + grubość),
    // a na obrazie tła wymazane — rysuje je dokument (tabela albo warstwa grafiki).
    const lines = opts.lines === false ? [] : (prep.segs || []).map((sg) => {
      const color = inkColor(pixels, up.width, { x0: sg.x0, y0: sg.y0, x1: sg.x1 + 1, y1: sg.y1 + 1 });
      const pad = 1.5 / scale;
      eraseBoxes.push({ x0: sg.x0 / scale - pad, y0: sg.y0 / scale - pad, x1: (sg.x1 + 1) / scale + pad, y1: (sg.y1 + 1) / scale + pad });
      return sg.horiz
        ? { x0: sg.x0 / scale, x1: (sg.x1 + 1) / scale, y0: (sg.y0 + sg.y1 + 1) / 2 / scale, y1: (sg.y0 + sg.y1 + 1) / 2 / scale, width: Math.max(0.5, (sg.y1 - sg.y0 + 1) / scale), color, order: 0, ocr: true }
        : { x0: (sg.x0 + sg.x1 + 1) / 2 / scale, x1: (sg.x0 + sg.x1 + 1) / 2 / scale, y0: sg.y0 / scale, y1: (sg.y1 + 1) / scale, width: Math.max(0.5, (sg.x1 - sg.x0 + 1) / scale), color, order: 0, ocr: true };
    });
    const rotate = prep.deg ? { deg: prep.deg, w: up.width / scale, h: up.height / scale, ow: canvas.width / scale, oh: canvas.height / scale, paper: prep.paper } : null;
    for (const c of new Set([canvas, up, prep.ocr])) c.width = c.height = 0; // zwolnij pamięć (telefon)
    return { glyphs, eraseBoxes, wordBoxes, lines, words, kept, skipped, doubtful, confidence: data.confidence, rotate };
  }

  window.DWPdfOcr = { ensure, recognizePage, terminate, hasSimd, setParallel };
})();
