// PDF → DOCX: rozpoznawanie tekstu na skanach (tesseract.js, w workerze, na urządzeniu).
// Strona-skan: renderujemy ją (pdf.js) do kanwy ~250 dpi, tesseract zwraca słowa z położeniem
// i pewnością. Słowa pewne (≥ 70%) zamieniamy na znaki w modelu strony (dalej ten sam układ co
// dla zwykłego PDF), a ich miejsca na obrazie skanu wypełniamy kolorem tła (eraseBoxes) —
// obraz zostaje pod tekstem jako tło: kolory, linie tabel, pieczątki, podpisy odręczne.
(function () {
  "use strict";

  const DIR = "lib/tesseract/";
  let workerPromise = null;
  let worker = null;
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

  async function ensure(onProgress) {
    progressCb = onProgress || null;
    if (!workerPromise) {
      workerPromise = (async () => {
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
        worker = w;
        return w;
      })().catch((e) => {
        workerPromise = null;
        throw e;
      });
    }
    return workerPromise;
  }

  async function terminate() {
    const w = worker;
    worker = null;
    workerPromise = null;
    try {
      await w?.terminate();
    } catch (_) { /* już zamknięty */ }
  }

  function median(arr) {
    if (!arr.length) return 0;
    const s = arr.slice().sort((a, b) => a - b);
    return s[s.length >> 1];
  }

  const OCR_FONT = { id: "ocr", family: "Arial", bold: false, italic: false, mono: false, serif: false, symbolic: false, asc: 0.72, desc: 0.21 };

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

  /**
   * Rozpoznanie jednej strony.
   * @returns { glyphs, eraseBoxes, words, chars } — współrzędne w pt (jak extractPage)
   */
  async function recognizePage(page, raw, opts = {}) {
    const w = await ensure(opts.onProgress);
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
    const { data } = await w.recognize(canvas, {}, { blocks: true, text: false });
    const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
    const glyphs = [];
    const eraseBoxes = [];
    let words = 0, kept = 0, order = 0, paraNo = 0;
    // Próg pewności (decyzja Mateusza 2026-10-02): na tekst zamieniamy tylko pewny druk —
    // ≥ 92%, dłuższe słowa z samych liter ≥ 85%. Reszta (pismo odręczne, nieczytelne) zostaje
    // nietknięta na obrazie i jest liczona jako „do poprawy”.
    const minConf = opts.minConfidence ?? 92;
    let skipped = 0;
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
          for (const word of line.words || []) {
            words++;
            const text = String(word.text || "").trim();
            if (!text || !/[\p{L}\p{N}]/u.test(text)) continue;
            const letters = [...text].filter((c) => /[\p{L}\p{N}]/u.test(c)).length;
            const ok = word.confidence >= minConf || (word.confidence >= 85 && letters >= 5 && /^[\p{L}.,:;()\-–]+$/u.test(text));
            if (!ok) {
              if (letters >= 2) skipped++;
              continue;
            }
            accepted.push({ word, line, text, paraNo });
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
        const color = inkColor(pixels, canvas.width, wb);
        const chars = [...a.text];
        const mk = (u, x, x1) => ({ u, x, x1, adv: x1 - x, y, size, asc: size * 0.72, desc: size * 0.21, angle: 0, font: OCR_FONT, isSpace: false, color, bold: false, italic: false, order: order++, ocr: true, ocrPara: a.paraNo });
        // Litery słowa jedna przy drugiej (szerokości jak w Arialu, dopasowane do ramki słowa) —
        // ramki pojedynczych znaków z OCR to ramki „tuszu”: przerwy między nimi udawałyby spacje.
        const x0 = wb.x0 / scale, x1 = wb.x1 / scale;
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
      }
    }
    canvas.width = canvas.height = 0; // zwolnij pamięć (telefon)
    return { glyphs, eraseBoxes, words, kept, skipped, confidence: data.confidence };
  }

  window.DWPdfOcr = { ensure, recognizePage, terminate, hasSimd };
})();
