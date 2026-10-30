// PDF → DOCX, sterowanie: strony po kolei (z oddawaniem wątku — apka nie zamarza), nagłówki
// i stopki powtarzające się na stronach, poziomy nagłówków, marginesy, obrazy → plik .docx.
// Środowisko (przeglądarka/Node) podaje: pdf.js, JSZip, kodowanie obrazów, poprawki znaków.
(function (root) {
  "use strict";
  const NS = (root.DWPdf = root.DWPdf || {});

  const yieldNow = () => new Promise((r) => {
    if (typeof MessageChannel !== "undefined") {
      const ch = new MessageChannel();
      ch.port1.onmessage = () => {
        ch.port1.close();
        r();
      };
      ch.port2.postMessage(0);
    } else setTimeout(r, 0);
  });

  class CancelledError extends Error {
    constructor() {
      super("cancelled");
      this.name = "CancelledError";
    }
  }
  NS.CancelledError = CancelledError;

  function modeOf(values) {
    const m = new Map();
    for (const [v, w] of values) m.set(v, (m.get(v) || 0) + w);
    let best = null, bw = -1;
    for (const [v, w] of m) if (w > bw) {
      best = v;
      bw = w;
    }
    return best;
  }

  // Wszystkie bloki (rekurencyjnie: komórki, kolumny) — do statystyk i nagłówków.
  function walkBlocks(blocks, fn, inTable = false) {
    for (const b of blocks) {
      fn(b, inTable);
      if (b.type === "table") for (const c of b.cells) walkBlocks(c.blocks || [], fn, true);
      if (b.type === "columns") for (const c of b.cells) walkBlocks(c.blocks || [], fn, inTable);
    }
  }

  function paraText(b) {
    return b.tokens.map((t) => (t.type === "tab" ? "\t" : t.text)).join("");
  }

  // Nagłówki: krótkie akapity większe od tekstu podstawowego, poziomy wg rozmiaru — jak Word przy
  // otwieraniu PDF: okładka (największe kroje tylko na 1. stronie) to Tytuł / Podtytuł, a
  // rozdziały zaczynają się od Nagłówka 1. Dawniej tytuł był Nagłówkiem 1, a rozdziały dopiero
  // trzecim poziomem; nagłówki 14 pt przy tekście 12 pt (×1,17) przepadały (próg ×1,18).
  function headingCandidate(b, inTable) {
    if (b.type !== "para" || inTable || b.nLines > 3 || b.list) return false;
    const text = paraText(b).trim();
    return text.length >= 2 && text.length <= 160 && !/\t/.test(text);
  }
  function assignHeadings(pages, bodySize) {
    const where = new Map(); // rozmiar → strony, na których jest
    for (const p of pages) walkBlocks(p.blocks, (b, inTable) => {
      if (!headingCandidate(b, inTable) || b.size < bodySize * 1.12) return;
      const k = Math.round(b.size);
      if (!where.has(k)) where.set(k, new Set());
      where.get(k).add(p.index);
    });
    const sizes = [...where.keys()].sort((a, b) => b - a);
    // okładka: kroje WIĘKSZE od każdego nagłówka z dalszych stron i tylko na 1. stronie
    const later = sizes.filter((k) => [...where.get(k)].some((i) => i > 0));
    const cover = pages.length > 1 && later.length ? sizes.filter((k) => k > Math.max(...later) && [...where.get(k)].every((i) => i === 0)) : [];
    const levels = sizes.filter((k) => !cover.includes(k)).slice(0, 4);
    for (const p of pages) walkBlocks(p.blocks, (b, inTable) => {
      if (!headingCandidate(b, inTable)) return;
      const k = Math.round(b.size);
      if (cover.includes(k)) b.titleRole = k === cover[0] ? "Title" : "Subtitle";
      else if (levels.includes(k)) b.headingLevel = levels.indexOf(k) + 1;
    });
    return levels;
  }

  // Spis treści: wiersze „tekst …… numer strony” (tabulator z kropkami do prawej) → link do
  // zakładki przy nagłówku o tym samym tekście (klik przenosi w Wordzie i w aplikacji), styl
  // „Spis treści N” wg poziomu nagłówka; tytuł spisu → „Nagłówek spisu treści” (nie rozdział).
  function assignToc(pages) {
    const norm = (t) => t.replace(/\s+/g, " ").replace(/[.…\s]+$/, "").trim().toLowerCase();
    const heads = [];
    const tocs = [];
    let prev = null;
    let tocTitle = null;
    for (const p of pages) walkBlocks(p.blocks, (b, inTable) => {
      if (b.type !== "para" || inTable) { prev = null; return; }
      const i = b.tokens.findIndex((t) => t.type === "tab" && t.leader && t.align === "right");
      const tail = i >= 0 ? b.tokens.slice(i + 1).filter((t) => t.type === "text").map((t) => t.text).join("").trim() : "";
      if (i > 0 && /^\d{1,4}$/.test(tail)) {
        const text = norm(b.tokens.slice(0, i).filter((t) => t.type === "text").map((t) => t.text).join(""));
        if (text) {
          if (!tocs.length && prev?.headingLevel && /^(spis treści|spis rzeczy|zawartość|table of contents|contents)$/i.test(paraText(prev).trim())) tocTitle = prev;
          tocs.push({ b, text });
        }
      } else if (b.headingLevel) heads.push({ b, text: norm(paraText(b)) });
      prev = b;
    });
    let n = 0;
    const used = new Set();
    for (const t of tocs) {
      const h = heads.find((x) => !used.has(x) && x.text === t.text);
      if (!h) continue;
      used.add(h);
      if (!h.b.bookmark) h.b.bookmark = `_Toc${++n}`;
      t.b.tocLink = h.b.bookmark;
      t.b.tocLevel = Math.min(9, h.b.headingLevel);
    }
    if (n >= 2 && tocTitle) { // „Spis treści” to nie rozdział (Word też go tak nie liczy)
      delete tocTitle.headingLevel;
      tocTitle.tocHeading = true;
    }
    return n;
  }

  // Listy: kolejne akapity ze znacznikiem (•, o, ▪, 1., a), (i)…) → prawdziwa lista Worda
  // (numeracja i wcięcia jako lista — edytuje się jak w Wordzie). Poziom = położenie znacznika.
  // Gdy czegoś nie da się odtworzyć wiernie (numery z przerwami, mieszane rodzaje na jednym
  // poziomie) — ciąg zostaje jak dawniej: znacznik i tabulator jako tekst.
  const ROMAN = { i: 1, v: 5, x: 10, l: 50, c: 100 };
  function romanValue(r) {
    let v = 0;
    const s = r.toLowerCase();
    for (let i = 0; i < s.length; i++) {
      const a = ROMAN[s[i]], b = ROMAN[s[i + 1]] || 0;
      if (!a) return 0;
      v += a < b ? -a : a;
    }
    return v;
  }
  function parseMarker(m) {
    let x = /^(\(?)(\d{1,3})([.)])$/.exec(m);
    if (x) return { fmt: "decimal", value: +x[2], pre: x[1], post: x[3] };
    x = /^(\(?)([ivxlc]{1,5})([.)])$/.exec(m);
    if (x && romanValue(x[2]) && x[2] !== "c" && x[2] !== "l") return { fmt: "lowerRoman", value: romanValue(x[2]), pre: x[1], post: x[3], letter: x[2].length === 1 ? x[2] : null };
    x = /^(\(?)([IVXLC]{1,5})([.)])$/.exec(m);
    if (x && romanValue(x[2]) && x[2] !== "C" && x[2] !== "L") return { fmt: "upperRoman", value: romanValue(x[2]), pre: x[1], post: x[3], letter: x[2].length === 1 ? x[2] : null };
    x = /^(\(?)([a-z])([.)])$/.exec(m);
    if (x) return { fmt: "lowerLetter", value: x[2].charCodeAt(0) - 96, pre: x[1], post: x[3] };
    x = /^(\(?)([A-Z])([.)])$/.exec(m);
    if (x) return { fmt: "upperLetter", value: x[2].charCodeAt(0) - 64, pre: x[1], post: x[3] };
    return null;
  }
  function assignLists(pages) {
    const defs = []; // definicje list (numbering.xml) — pdf-docx.js
    // pojemniki: treść stron (lista może przejść na następną stronę) + każda komórka / kolumna
    const containers = [[]];
    const collect = (blocks, into) => {
      for (const b of blocks) {
        into.push(b);
        if (b.type === "table" || b.type === "columns") for (const c of b.cells) { const sub = []; containers.push(sub); collect(c.blocks || [], sub); }
      }
    };
    for (const p of pages) collect(p.blocks, containers[0]);
    for (const blocks of containers) {
      let run = [];
      const flushRun = () => { if (run.length) makeList(run, defs); run = []; };
      for (const b of blocks) {
        if (b.type === "para" && b.list && b.listMarker) run.push(b);
        else flushRun();
      }
      flushRun();
    }
    return defs;
  }
  function makeList(run, defs) {
    // poziomy: położenie znacznika od lewej krawędzi obszaru (tolerancja 4 pt)
    const pos = (b) => b.indLeft + b.firstLine;
    const cols = [];
    for (const b of run) if (!cols.some((c) => Math.abs(c - pos(b)) < 4)) cols.push(pos(b));
    cols.sort((a, b) => a - b);
    if (cols.length > 9) return;
    const lvlOf = (b) => cols.findIndex((c) => Math.abs(c - pos(b)) < 4);
    const levels = [];
    let last = []; // ostatnia wartość na poziomie (numeracja od nowa po wyższym poziomie)
    for (const b of run) {
      const L = lvlOf(b);
      const num = parseMarker(b.listMarker);
      const kind = b.list === "bullet" || !num ? "bullet" : "number";
      if (b.list === "number" && !num) return; // „1.1.”, „§ 3” itp. — zostaje tekstem
      const lv = levels[L] || (levels[L] = { kind, markers: new Map(), fmt: num?.fmt, pre: num?.pre, post: num?.post, start: num?.value, tok: b.tokens.find((t) => t.type === "text") });
      if (lv.kind !== kind) return;
      if (kind === "bullet") lv.markers.set(b.listMarker, (lv.markers.get(b.listMarker) || 0) + 1);
      else {
        // „i.” może być literą (a, b, … i) albo rzymską jedynką — rozstrzyga poziom
        let fmt = num.fmt, value = num.value;
        if (lv.fmt === "lowerLetter" && fmt === "lowerRoman" && num.letter) { fmt = "lowerLetter"; value = num.letter.charCodeAt(0) - 96; }
        if (lv.fmt === "upperLetter" && fmt === "upperRoman" && num.letter) { fmt = "upperLetter"; value = num.letter.charCodeAt(0) - 64; }
        if (fmt !== lv.fmt || num.pre !== lv.pre || num.post !== lv.post) return;
        const expect = last[L] == null ? (L === 0 ? lv.start : 1) : last[L] + 1;
        if (value !== expect) return;
        last[L] = value;
      }
      for (let k = L + 1; k < last.length; k++) last[k] = null; // głębsze poziomy liczą od nowa
    }
    if (levels.some((lv) => !lv)) return; // dziura w poziomach (np. tylko poziom 2) — zostaje tekstem
    const def = {
      id: defs.length + 1,
      levels: levels.map((lv, i) => {
        const sample = run.find((b) => lvlOf(b) === i);
        return {
          kind: lv.kind,
          fmt: lv.kind === "bullet" ? "bullet" : lv.fmt,
          text: lv.kind === "bullet" ? [...lv.markers].sort((a, b) => b[1] - a[1])[0][0] : `${lv.pre}%${i + 1}${lv.post}`,
          start: lv.kind === "bullet" ? 1 : lv.start,
          left: sample.indLeft,
          hanging: Math.max(0, -sample.firstLine),
          font: lv.tok?.font,
          size: lv.tok?.size,
        };
      }),
    };
    defs.push(def);
    for (const b of run) b.numbering = { numId: def.id, ilvl: lvlOf(b) };
  }

  // Treść: atomy drzewa układu (do marginesów).
  function treeBounds(tree, floats) {
    let x0 = Infinity, x1 = -Infinity, top = Infinity, bottom = -Infinity;
    const walk = (n) => {
      if (n.type === "flow") {
        for (const a of n.atoms) {
          x0 = Math.min(x0, a.x0);
          x1 = Math.max(x1, a.x1);
          top = Math.min(top, a.top);
          bottom = Math.max(bottom, a.bottom);
        }
      } else n.children.forEach(walk);
    };
    walk(tree);
    return { x0, x1, top, bottom, empty: !isFinite(x0) };
  }

  /**
   * Konwersja całego dokumentu.
   * @param {object} pdfDoc   PDFDocumentProxy
   * @param {object} env      { pdfjs, JSZip, encodeImage(page, img) → {bytes, ext}|null,
   *                            glyphFixes(page, raw) → Map|null, onProgress(i, n, phase),
   *                            signal: { cancelled }, maxPages }
   */
  // OCR wyprostował stronę (krzywy skan, kartka bokiem / do góry nogami): rozpoznane słowa są
  // w układzie wyprostowanej strony, więc strona dostaje jej wymiary, a obraz skanu — obrót wokół
  // środka (rotateM: punkt strony → punkt wyprostowanej strony, stosuje go encodeImage).
  // Reszta starej strony (niewidoczne znaki, drobne kreski) jest w starym układzie — pomijamy ją.
  function straightenScanPage(raw, scanImg, R) {
    const t = (R.deg * Math.PI) / 180, c = Math.cos(t), s = Math.sin(t);
    scanImg.rotateM = [c, s, -s, c, R.w / 2 - (c * R.ow / 2 - s * R.oh / 2), R.h / 2 - (s * R.ow / 2 + c * R.oh / 2)];
    scanImg.clipped = { x0: 0, y0: 0, x1: R.w, y1: R.h };
    scanImg.paper = R.paper; // rogi po obrocie w kolorze papieru
    raw.width = R.w;
    raw.height = R.h;
    raw.images = [scanImg];
    raw.glyphs = [];
    for (const k of ["rects", "lines", "graphics", "annotations"]) if (Array.isArray(raw[k])) raw[k] = [];
  }

  // Strona z tekstem-krzywymi: (prawie) bez znaków, za to ≥ 40 małych kształtów (wielkości liter).
  function looksLikeOutlinedText(raw, visible) {
    const small = raw.graphics.filter((g) => g.kind === "fill" && g.y1 - g.y0 > 1 && g.y1 - g.y0 <= 40 && g.x1 - g.x0 <= 60).length;
    return small >= 40 && visible < small * 0.05;
  }
  // Kształt (litera-rysunek) leży w ramce rozpoznanego słowa — znika, bo zastępuje go tekst.
  // Kształty nierozpoznane (logo, pismo, słowo poniżej progu pewności) zostają na stronie.
  function coveredByWords(g, boxes) {
    const area = Math.max(0.01, (g.x1 - g.x0) * (g.y1 - g.y0));
    const cx = (g.x0 + g.x1) / 2, gh = Math.max(0.01, g.y1 - g.y0);
    for (const b of boxes) {
      const ix = Math.min(g.x1, b.x1 + 0.8) - Math.max(g.x0, b.x0 - 0.8), iy = Math.min(g.y1, b.y1 + 0.8) - Math.max(g.y0, b.y0 - 0.8);
      if (ix > 0 && iy > 0 && ix * iy >= area * 0.7) return true;
      // środek kształtu w ramce słowa i większość wysokości w jej pasie (ramki OCR bywają ciasne)
      if (iy >= gh * 0.6 && cx >= b.x0 - 0.5 && cx <= b.x1 + 0.5) return true;
    }
    return false;
  }

  async function convertPdf(pdfDoc, env) {
    const n = Math.min(pdfDoc.numPages, env.maxPages || Infinity);
    const fontCache = new Map();
    const raws = [];
    const report = { pages: n, tables: 0, images: 0, scannedPages: 0, ocrPages: 0, fieldsText: 0, timings: {} };
    const t0 = Date.now();
    const pageObjs = [];
    let ocrChoice = null; // null = jeszcze nie pytano
    // OCR w potoku: strona startuje od razu, a czekamy dopiero, gdy w toku jest `ocrLimit` stron —
    // przygotowanie kolejnej (wątek główny) idzie równolegle z rozpoznawaniem poprzedniej (wątki
    // tesseracta). Wynik zmienia `raw` danej strony w miejscu, więc kolejność stron zostaje.
    const ocrLimit = Math.max(1, env.ocrParallel || 1);
    const inflight = [];
    const queueOcr = async (job) => {
      const p = job();
      p.catch(() => {}); // błąd i tak wyjdzie przy await poniżej
      inflight.push(p);
      if (inflight.length >= ocrLimit) await inflight.shift();
    };
    for (let i = 1; i <= n; i++) {
      if (env.signal?.cancelled) throw new CancelledError();
      env.onProgress?.(i, n, "read");
      const page = await pdfDoc.getPage(i);
      const raw = await NS.extractPage(page, env.pdfjs, { fontCache });
      if (env.glyphFixes) raw.glyphFixes = await env.glyphFixes(page, raw);
      // Skan: strona bez tekstu, za to z dużym obrazem. Gdy PDF ma już niewidoczną warstwę
      // tekstu (po OCR w skanerze) — bierzemy ją; inaczej (za zgodą) rozpoznajemy tekst sami.
      // W obu przypadkach słowa „wymazujemy” z obrazu, który zostaje tłem strony.
      const area = (b) => Math.max(0, b.x1 - b.x0) * Math.max(0, b.y1 - b.y0);
      const pageArea = raw.width * raw.height;
      const scanImg = raw.images.filter((im) => im.kind === "image" && area(im.clipped || im) >= pageArea * 0.4).sort((a, b) => area(b.clipped || b) - area(a.clipped || a))[0];
      const visible = raw.glyphs.filter((g) => !g.invisible && g.u.trim()).length;
      if (scanImg && visible < 3) {
        const hidden = raw.glyphs.filter((g) => g.invisible && g.u.trim());
        if (hidden.length >= 3) {
          scanImg.eraseBoxes = hidden.map((g) => ({ x0: g.x - 0.5, x1: g.x1 + 0.5, y0: g.y - g.asc, y1: g.y + g.desc }));
          report.ocrLayerPages = (report.ocrLayerPages || 0) + 1;
        } else if (env.ocrPage) {
          if (ocrChoice === null) ocrChoice = env.askOcr ? !!(await env.askOcr()) : true;
          if (env.signal?.cancelled) throw new CancelledError();
          if (ocrChoice) {
            env.onProgress?.(i, n, "ocr");
            await queueOcr(async () => {
            const res = await env.ocrPage(page, raw, (status, p) => env.onProgress?.(i, n, "ocr", p));
            if (res && res.glyphs.length) {
              if (res.rotate) straightenScanPage(raw, scanImg, res.rotate);
              raw.glyphs.push(...res.glyphs);
              if (res.lines?.length) raw.lines.push(...res.lines);
              scanImg.eraseBoxes = res.eraseBoxes;
              report.ocrPages++;
              if (res.rotate) report.straightened = (report.straightened || 0) + 1;
            }
            if (res) {
              report.ocrSkipped = (report.ocrSkipped || 0) + (res.skipped || 0);
              report.ocrDoubt = (report.ocrDoubt || 0) + (res.doubtful || 0);
            }
            });
          }
        }
      }
      // Tekst zapisany jako rysunek („zamień czcionki na krzywe”, eksport z programów graficznych,
      // Ghostscript): zero znaków, setki małych kształtów liter. Bez OCR dokument ma „0 słów”.
      else if (env.ocrPage && looksLikeOutlinedText(raw, visible)) {
        if (ocrChoice === null) ocrChoice = env.askOcr ? !!(await env.askOcr("curves")) : true;
        if (env.signal?.cancelled) throw new CancelledError();
        if (ocrChoice) {
          env.onProgress?.(i, n, "ocr");
          // strona wektorowa jest prosta, a linie tabel już są w PDF — bez prostowania i bez linii z obrazu
          await queueOcr(async () => {
          const res = await env.ocrPage(page, raw, (status, p) => env.onProgress?.(i, n, "ocr", p), { straighten: false, lines: false });
          if (res && res.glyphs.length) {
            if (Array.isArray(window.__dwbOcrTrace)) for (const g of raw.graphics) if (!coveredByWords(g, res.wordBoxes) && res.wordBoxes.some((b) => g.x0 < b.x1 && g.x1 > b.x0 && g.y0 < b.y1 && g.y1 > b.y0)) window.__dwbOcrTrace.push({ left: [g.x0, g.y0, g.x1, g.y1].map((v) => +v.toFixed(1)), box: res.wordBoxes.filter((b) => g.x0 < b.x1 && g.x1 > b.x0 && g.y0 < b.y1 && g.y1 > b.y0).map((b) => [b.x0, b.y0, b.x1, b.y1].map((v) => +v.toFixed(1))) });
            raw.graphics = raw.graphics.filter((g) => !coveredByWords(g, res.wordBoxes));
            // kreski liter (i, l, kropki) bywają zapisane jako prostokąty — też znikają (krawędzie tabel są długie, zostają)
            raw.rects = raw.rects.filter((r) => r.y1 - r.y0 > 40 || r.x1 - r.x0 > 60 || !coveredByWords(r, res.wordBoxes));
            raw.glyphs.push(...res.glyphs);
            report.ocrPages++;
            report.curvePages = (report.curvePages || 0) + 1;
          }
          if (res) {
            report.ocrSkipped = (report.ocrSkipped || 0) + (res.skipped || 0);
            report.ocrDoubt = (report.ocrDoubt || 0) + (res.doubtful || 0);
          }
          });
        }
      }
      raws.push(raw);
      pageObjs.push(page);
      await yieldNow();
    }
    while (inflight.length) await inflight.shift();
    if (env.signal?.cancelled) throw new CancelledError();
    report.timings.read = Date.now() - t0;

    // Powtarzające się nagłówki/stopki — wycinamy z treści, wstawiamy do nagłówka/stopki Worda.
    const running = NS.detectRunningText(raws);
    const runningRows = raws.map(() => ({ top: [], bottom: [], topImages: [], bottomImages: [] }));
    raws.forEach((raw, pi) => {
      if (!running.size) return;
      const H = raw.height;
      const rows = new Map();
      for (const g of raw.glyphs) {
        if (g.invisible) continue;
        const zone = NS.runningZone(g.y, H);
        if (!zone) continue;
        const k = zone.replace("+", "") + ":" + Math.round(g.y);
        if (!rows.has(k)) rows.set(k, { gl: [], wide: zone.endsWith("+"), y: Math.round(g.y) });
        rows.get(k).gl.push(g);
      }
      // od krawędzi strony do środka: w strefie poszerzonej wiersz musi mieć same powtarzające się
      // wiersze między sobą a krawędzią (tytuł dokumentu na tej samej wysokości co na innej stronie
      // nie jest nagłówkiem)
      for (const zone of ["top", "bottom"]) {
        const list = [...rows.entries()].filter(([k]) => k.startsWith(zone + ":")).sort((a, b) => (zone === "top" ? a[1].y - b[1].y : b[1].y - a[1].y));
        let contiguous = true;
        for (const [, row] of list) {
          const exact = row.gl.slice().sort((a, b) => a.x - b.x).map((g) => g.u).join("").replace(/\s+/g, " ").trim();
          if (!exact) continue;
          const isRunning = row.wide ? contiguous && running.has("=" + zone + "|" + exact) : running.has(zone + "|" + exact.replace(/\d+/g, "#"));
          if (!isRunning) {
            contiguous = false;
            continue;
          }
          for (const g of row.gl) g.running = true;
          runningRows[pi][zone].push(row.gl);
        }
      }
      raw.glyphs = raw.glyphs.filter((g) => !g.running);
    });
    // Obrazy w nagłówku/stopce (logo firmy): to samo miejsce i rozmiar na co najmniej połowie stron,
    // obok powtarzającego się tekstu (sięga ponad jego dolny brzeg) albo przy samej krawędzi strony.
    if (n >= 2) {
      const same = (a, b) => Math.abs(a.x0 - b.x0) < 1 && Math.abs(a.x1 - b.x1) < 1 && Math.abs(a.y0 - b.y0) < 1 && Math.abs(a.y1 - b.y1) < 1;
      const need = Math.max(2, Math.ceil(n * 0.5));
      raws.forEach((raw, pi) => {
        const H = raw.height;
        const rr = runningRows[pi];
        const textTop = rr.top.length ? Math.max(...rr.top.flat().map((g) => g.y + (g.desc || g.size * 0.2))) : -Infinity;
        const textBottom = rr.bottom.length ? Math.min(...rr.bottom.flat().map((g) => g.y - (g.asc || g.size * 0.8))) : Infinity;
        raw.images = raw.images.filter((img) => {
          if (img.kind !== "image") return true;
          const zone = img.y1 <= H * 0.2 || (img.y0 < H * 0.12 && img.y1 <= H * 0.3) ? "top" : img.y0 >= H * 0.85 ? "bottom" : null;
          if (!zone) return true;
          const nearText = zone === "top" ? img.y0 < textTop - 1 || img.y1 <= H * 0.1 : img.y1 > textBottom + 1 || img.y0 >= H * 0.9;
          if (!nearText) return true;
          const pagesWith = raws.filter((r) => r.images.some((o) => o.kind === "image" && same(o, img)) || runningRows[raws.indexOf(r)][zone + "Images"].some((o) => same(o, img))).length;
          if (pagesWith < need) return true;
          rr[zone + "Images"].push(img);
          return false;
        });
      });
    }

    // Układ stron
    const t1 = Date.now();
    const layouts = [];
    for (let i = 0; i < raws.length; i++) {
      if (env.signal?.cancelled) throw new CancelledError();
      env.onProgress?.(i + 1, n, "layout");
      const L = NS.layoutPage(raws[i], { glyphFixes: raws[i].glyphFixes, formFields: env.formFields !== false });
      layouts.push(L);
      report.tables += L.tables.length;
      report.images += raws[i].images.length;
      if (!L.hasText && raws[i].images.length) report.scannedPages++;
      await yieldNow();
    }
    report.timings.layout = Date.now() - t1;

    // Marginesy: wspólne dla stron o tym samym rozmiarze (minimum z treści).
    const groups = new Map();
    layouts.forEach((L, i) => {
      const k = Math.round(L.width) + "x" + Math.round(L.height);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(i);
    });
    const margins = new Array(layouts.length);
    for (const idx of groups.values()) {
      let x0 = Infinity, x1 = -Infinity, top = Infinity, bottom = -Infinity;
      for (const i of idx) {
        const b = treeBounds(layouts[i].tree);
        if (b.empty) continue;
        x0 = Math.min(x0, b.x0);
        x1 = Math.max(x1, b.x1);
        top = Math.min(top, b.top);
        bottom = Math.max(bottom, b.bottom);
      }
      const W = layouts[idx[0]].width, H = layouts[idx[0]].height;
      if (!isFinite(x0)) {
        x0 = 56; x1 = W - 56; top = 56; bottom = H - 56;
      }
      const ml = Math.max(4, Math.floor(x0 - 1));
      const mr = Math.max(4, Math.floor(W - x1 - 1));
      const mt = Math.max(4, Math.floor(top - 2));
      // dół: mały margines = zapas, gdy Word złoży tekst minimalnie wyżej niż PDF
      const mb = Math.max(4, Math.min(Math.floor(H - bottom - 2), 20));
      for (const i of idx) margins[i] = { ml, mr, mt, mb };
    }

    // Bloki stron + tekst podstawowy (najczęstszy rozmiar/krój)
    const pages = [];
    const sizeVotes = [], fontVotes = [];
    for (let i = 0; i < layouts.length; i++) {
      const L = layouts[i], m = margins[i];
      const blocks = NS.treeToBlocks(L.tree, { x0: m.ml, x1: L.width - m.mr, top: m.mt });
      pages.push({ width: L.width, height: L.height, blocks, floats: L.floats, margins: m, index: i });
    }
    for (const p of pages) walkBlocks(p.blocks, (b) => {
      if (b.type !== "para") return;
      for (const t of b.tokens) if (t.type === "text") {
        sizeVotes.push([t.size, t.text.length]);
        fontVotes.push([t.font, t.text.length]);
      }
    });
    const bodySize = modeOf(sizeVotes) || 11;
    const bodyFont = modeOf(fontVotes) || "Calibri";
    assignHeadings(pages, bodySize);
    const tocLinks = assignToc(pages);
    const listDefs = assignLists(pages);

    // Stopki/nagłówki z numerem strony → pole PAGE / NUMPAGES
    if (running.size || runningRows.some((r) => r.topImages.length || r.bottomImages.length)) {
      for (const zone of ["bottom", "top"]) {
        const first = runningRows.findIndex((r) => r[zone].length || r[zone + "Images"].length);
        if (first < 0) continue;
        const xml = runningXml(runningRows[first][zone], first + 1, n, layouts[first], margins[first], zone, raws[first].glyphFixes);
        const pics = runningRows[first][zone + "Images"];
        for (const p of pages) {
          if (zone === "bottom") {
            p.footer = xml.xml;
            p.footerKey = "f";
            p.footerImages = pics;
            p.margins = { ...p.margins, footerDist: xml.dist };
          } else {
            p.header = xml.xml;
            p.headerKey = "h";
            p.headerImages = pics;
            p.margins = { ...p.margins, headerDist: xml.dist };
          }
        }
      }
    }

    // Obrazy (kodowanie: przeglądarka — kanwa; Node — pomijamy albo test podaje własne)
    const t2 = Date.now();
    const images = new Map();
    if (env.encodeImage) {
      const seen = new Map();
      for (let i = 0; i < pages.length; i++) {
        const raw = raws[i];
        const own = [...(pages[i].headerImages || []), ...(pages[i].footerImages || [])].filter((im) => raw.images.includes(im) || runningRows[i].topImages.includes(im) || runningRows[i].bottomImages.includes(im));
        for (const img of [...raw.images, ...own]) {
          if (env.signal?.cancelled) throw new CancelledError();
          const key = (img.objId || "") + "|" + [img.x0, img.y0, img.x1, img.y1].map((v) => Math.round(v)).join(",") + "|" + i;
          if (seen.has(key)) {
            images.set(img, seen.get(key));
            continue;
          }
          env.onProgress?.(i + 1, n, "images");
          try {
            const enc = await env.encodeImage(pageObjs[i], img);
            if (enc) {
              images.set(img, enc);
              seen.set(key, enc);
            }
          } catch (_) { /* obraz pomijamy, reszta dokumentu dalej */ }
          await yieldNow();
        }
      }
    }
    // Warstwa grafiki wektorowej (linie, ramki, znaczniki cięcia…) — obraz pod tekstem.
    if (env.renderVectors) {
      for (let i = 0; i < pages.length; i++) {
        const L = layouts[i];
        if (!L.vectors || !L.vectors.length) continue;
        if (env.signal?.cancelled) throw new CancelledError();
        try {
          const enc = await env.renderVectors(L.vectors, L.width, L.height);
          if (enc) {
            const key = { vectorLayer: i };
            images.set(key, enc);
            pages[i].floats = [{ img: key, x0: 0, x1: L.width, top: 0, bottom: L.height, behind: true, vector: true }, ...(pages[i].floats || [])];
            report.vectorPages = (report.vectorPages || 0) + 1;
          }
        } catch (_) { /* bez warstwy grafiki — tekst i tak jest */ }
        await yieldNow();
      }
    }
    report.timings.images = Date.now() - t2;

    // Czcionki z PDF do osadzenia (krój, którego może nie być na innym urządzeniu)
    let fonts = [];
    if (NS.buildEmbeddedFonts && env.embedFonts !== false) {
      const usage = new Map();
      layouts.forEach((L, i) => {
        for (const g of L.glyphs || []) {
          if (!g.fontRef || g.composed || g.symbol || g.ocrLayer || !g.fontChar) continue;
          const cps = [...g.u];
          const fcs = [...String(g.fontChar)];
          if (cps.length !== 1 || fcs.length !== 1) continue;
          let u = usage.get(g.fontRef);
          if (!u) {
            let fontObj = null;
            try {
              fontObj = pageObjs[i].commonObjs.get(g.fontRef);
            } catch (_) { /* brak */ }
            u = { fontObj, family: g.font.family, bold: g.font.bold, italic: g.font.italic, symbolic: g.font.symbolic, chars: new Map() };
            usage.set(g.fontRef, u);
          }
          u.chars.set(cps[0].codePointAt(0), fcs[0].codePointAt(0));
        }
      });
      try {
        fonts = NS.buildEmbeddedFonts(usage);
      } catch (_) { fonts = []; }
      report.embeddedFonts = fonts.map((f) => `${f.family} ${f.style}`);
    }

    env.onProgress?.(n, n, "write");
    const meta = { bodyFont, bodySize, title: env.title || "", lang: env.lang || "pl-PL", fonts, listDefs };
    report.lists = listDefs.length;
    report.tocLinks = tocLinks;
    const t3 = Date.now();
    const bytes = await NS.buildDocx(pages, images, meta, env.JSZip);
    report.timings.write = Date.now() - t3;
    report.suspectChars = meta.suspectChars || 0;
    report.bodyFont = bodyFont;
    report.bodySize = bodySize;
    report.textChars = 0;
    for (const p of pages) walkBlocks(p.blocks, (b) => {
      if (b.type === "para") report.textChars += paraText(b).length;
    });
    for (const raw of raws) report.fieldsText += raw.annotations.filter((a) => a.type === "text").length;
    return { bytes, report, pages };
  }

  function runningXml(rowsGl, pageNo, total, L, m, zone, fixes) {
    // Każdy wiersz → akapit stopki; cyfry równe numerowi strony → pole PAGE, liczbie stron → NUMPAGES.
    const W = L.width, H = L.height;
    let xml = "";
    let dist = 18;
    // jak esc(clean()) w pdf-docx.js: nazwa kroju z PDF-a i tekst stopki mogą mieć „"”, „&”,
    // znaki sterujące albo samotne surogaty — bez tego header1/footer1.xml był niepoprawny
    const xmlSafe = (v) => String(v ?? "").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, "")
      .replace(/[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    // Znaki jak w treści (pdf-layout prepareGlyphs): poprawki kształtu ą/ę, punktory z krojów symboli
    // (Wingdings U+F09F → „•” — w kroju tekstu, prywatny kod bez Wingdings był kwadracikiem).
    const prep = (g0) => {
      const g = { ...g0 };
      const fix = fixes && fixes.get(g.fontRef + "|" + g.code);
      if (fix) g.u = fix;
      const sym = NS.mapSymbolChar ? NS.mapSymbolChar(g) : null;
      if (sym) {
        g.u = sym;
        g.symbol = true;
      }
      return g;
    };
    // wiersze o różnych liniach bazowych (punktor z innego kroju 1–2 pt niżej) to jeden wiersz
    const merged = [];
    for (const gl of rowsGl.slice().sort((a, b) => a[0].y - b[0].y)) {
      const y = gl[0].y, size = Math.max(...gl.map((g) => g.size));
      const prev = merged[merged.length - 1];
      if (prev && Math.abs(y - prev.y) < Math.min(size, prev.size) * 0.4) {
        prev.gl.push(...gl);
        if (gl.length > prev.main) { prev.y = y; prev.main = gl.length; }
        prev.size = Math.max(prev.size, size);
      } else merged.push({ gl: gl.slice(), y, size, main: gl.length });
    }
    let prevBase = null;
    merged.forEach((row, ri) => {
      // zdublowane znaki (pogrubienie przez kilkukrotny druk w tym samym miejscu) — jak w treści
      const g0 = [];
      for (const g of row.gl.map(prep).sort((a, b) => a.x - b.x)) {
        const p = g0.findLast((o) => o.u === g.u);
        if (p && g.x - p.x < g.size * 0.08 && Math.abs(p.y - g.y) < g.size * 0.08) continue;
        g0.push(g);
      }
      const x0 = g0[0].x, x1 = Math.max(...g0.map((g) => g.x1));
      const size = row.size;
      const center = (x0 + x1) / 2;
      const align = Math.abs(center - W / 2) < 12 ? "center" : x1 > W - m.mr - 8 && x0 > W / 2 ? "right" : "left";
      // przebiegi według stylu znaków (krój, pogrubienie, kursywa, kolor, rozmiar)
      const style = (g) => {
        const sym = g.symbol || /symbol|wingding|webding|dingbat/i.test(g.font.family);
        return [sym ? "" : g.font.family, !!g.font.bold, !!g.font.italic, g.color || "000000", Math.round(g.size * 2)].join("|");
      };
      const spans = [];
      let prev = null;
      for (const g of g0) {
        const k = g.isSpace && spans.length ? spans[spans.length - 1].k : style(g);
        let sp = spans[spans.length - 1];
        if (!sp || sp.k !== k) spans.push((sp = { k, g, gl: [] }));
        if (prev && !g.isSpace && !prev.isSpace && g.x - prev.x1 > g.size * 0.17) sp.gl.push({ isSpace: true });
        sp.gl.push(g);
        prev = g;
      }
      const baseFont = (g0.find((g) => !g.symbol) || g0[0]).font.family;
      let runs = "";
      spans.forEach((sp, si) => {
        let text = sp.gl.map((g) => (g.isSpace ? " " : g.u)).join("").replace(/ +/g, " ");
        if (si === 0) text = text.replace(/^ +/, "");
        if (si === spans.length - 1) text = text.replace(/ +$/, "");
        if (!text) return;
        const [fam, bold, italic, color, hp] = sp.k.split("|");
        const font = xmlSafe(fam || baseFont);
        const rPr = `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>${bold === "true" ? "<w:b/><w:bCs/>" : ""}${italic === "true" ? "<w:i/><w:iCs/>" : ""}${color && color !== "000000" ? `<w:color w:val="${color.toUpperCase()}"/>` : ""}<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/></w:rPr>`;
        for (const part of text.split(/(\d+)/)) {
          if (!part) continue;
          if (/^\d+$/.test(part) && (+part === pageNo || +part === total)) {
            const field = +part === pageNo ? "PAGE" : "NUMPAGES";
            runs += `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rPr}<w:instrText xml:space="preserve"> ${field} </w:instrText></w:r><w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rPr}<w:t>${part}</w:t></w:r><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`;
          } else {
            runs += `<w:r>${rPr}<w:t xml:space="preserve">${xmlSafe(part)}</w:t></w:r>`;
          }
        }
      });
      const L1 = size * 1.17;
      // odstęp przed wierszem = przerwa między liniami bazowymi ponad wysokość wiersza
      const before = prevBase != null ? Math.max(0, row.y - prevBase - L1) : 0;
      prevBase = row.y;
      const ind = align === "left" ? `<w:ind w:left="${Math.round(Math.max(0, x0 - m.ml) * 20)}"/>` : "";
      xml += `<w:p><w:pPr><w:spacing w:before="${before > 0.3 ? Math.round(before * 20) : 0}" w:after="0" w:line="${Math.round(L1 * 20)}" w:lineRule="exact"/>${ind}<w:jc w:val="${align}"/></w:pPr>${runs}</w:p>`;
      if (zone === "bottom") dist = Math.max(4, H - (row.y + (1 - NS.DOCX_BASE) * L1));
      else if (ri === 0) dist = Math.max(4, row.y - NS.DOCX_BASE * L1);
    });
    // sam obraz (logo bez tekstu) — nagłówek Worda musi mieć akapit, w nim kotwica obrazu
    if (!xml) xml = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>`;
    return { xml, dist };
  }

  NS.convertPdf = convertPdf;
  NS.yieldNow = yieldNow;
})(typeof globalThis !== "undefined" ? globalThis : window);
