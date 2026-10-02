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

  // Nagłówki: akapity wyraźnie większe od tekstu podstawowego (poziomy wg rozmiaru).
  function assignHeadings(pages, bodySize) {
    const sizes = new Set();
    for (const p of pages) walkBlocks(p.blocks, (b, inTable) => {
      if (b.type !== "para" || inTable || b.nLines > 3) return;
      const text = paraText(b).trim();
      if (text.length < 2 || text.length > 160 || /\t/.test(text)) return;
      if (b.size >= bodySize * 1.18) sizes.add(Math.round(b.size));
    });
    const levels = [...sizes].sort((a, b) => b - a).slice(0, 4);
    for (const p of pages) walkBlocks(p.blocks, (b, inTable) => {
      if (b.type !== "para" || inTable || b.nLines > 3) return;
      const text = paraText(b).trim();
      if (text.length < 2 || text.length > 160 || /\t/.test(text)) return;
      const lvl = levels.indexOf(Math.round(b.size));
      if (lvl >= 0) b.headingLevel = lvl + 1;
    });
    return levels;
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
  async function convertPdf(pdfDoc, env) {
    const n = Math.min(pdfDoc.numPages, env.maxPages || Infinity);
    const fontCache = new Map();
    const raws = [];
    const report = { pages: n, tables: 0, images: 0, scannedPages: 0, ocrPages: 0, fieldsText: 0, timings: {} };
    const t0 = Date.now();
    const pageObjs = [];
    let ocrChoice = null; // null = jeszcze nie pytano
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
            const res = await env.ocrPage(page, raw, (status, p) => env.onProgress?.(i, n, "ocr", p));
            if (res && res.glyphs.length) {
              raw.glyphs.push(...res.glyphs);
              scanImg.eraseBoxes = res.eraseBoxes;
              report.ocrPages++;
            }
            if (res) report.ocrSkipped = (report.ocrSkipped || 0) + (res.skipped || 0);
          }
        }
      }
      raws.push(raw);
      pageObjs.push(page);
      await yieldNow();
    }
    report.timings.read = Date.now() - t0;

    // Powtarzające się nagłówki/stopki — wycinamy z treści, wstawiamy do nagłówka/stopki Worda.
    const running = NS.detectRunningText(raws);
    const runningRows = raws.map(() => ({ top: [], bottom: [] }));
    raws.forEach((raw, pi) => {
      if (!running.size) return;
      const H = raw.height;
      const rows = new Map();
      for (const g of raw.glyphs) {
        if (g.invisible) continue;
        const zone = g.y < H * 0.1 ? "top" : g.y > H * 0.9 ? "bottom" : null;
        if (!zone) continue;
        const k = zone + ":" + Math.round(g.y);
        if (!rows.has(k)) rows.set(k, []);
        rows.get(k).push(g);
      }
      for (const [k, gl] of rows) {
        const text = gl.slice().sort((a, b) => a.x - b.x).map((g) => g.u).join("").replace(/\d+/g, "#").replace(/\s+/g, " ").trim();
        const zone = k.split(":")[0];
        if (running.has(zone + "|" + text)) {
          for (const g of gl) g.running = true;
          runningRows[pi][zone].push(gl);
        }
      }
      raw.glyphs = raw.glyphs.filter((g) => !g.running);
    });

    // Układ stron
    const t1 = Date.now();
    const layouts = [];
    for (let i = 0; i < raws.length; i++) {
      if (env.signal?.cancelled) throw new CancelledError();
      env.onProgress?.(i + 1, n, "layout");
      const L = NS.layoutPage(raws[i], { glyphFixes: raws[i].glyphFixes });
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

    // Stopki/nagłówki z numerem strony → pole PAGE / NUMPAGES
    if (running.size) {
      for (const zone of ["bottom", "top"]) {
        const first = runningRows.findIndex((r) => r[zone].length);
        if (first < 0) continue;
        const xml = runningXml(runningRows[first][zone], first + 1, n, layouts[first], margins[first], zone);
        for (const p of pages) {
          if (zone === "bottom") {
            p.footer = xml.xml;
            p.footerKey = "f";
            p.margins = { ...p.margins, footerDist: xml.dist };
          } else {
            p.header = xml.xml;
            p.headerKey = "h";
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
        for (const img of raw.images) {
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
    const meta = { bodyFont, bodySize, title: env.title || "", lang: env.lang || "pl-PL", fonts };
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

  function runningXml(rowsGl, pageNo, total, L, m, zone) {
    // Każdy wiersz → akapit stopki; cyfry równe numerowi strony → pole PAGE, liczbie stron → NUMPAGES.
    const W = L.width, H = L.height;
    let xml = "";
    let dist = 18;
    const rows = rowsGl.slice().sort((a, b) => a[0].y - b[0].y);
    for (const gl of rows) {
      const g0 = gl.slice().sort((a, b) => a.x - b.x);
      const x0 = g0[0].x, x1 = Math.max(...g0.map((g) => g.x1));
      const size = Math.max(...g0.map((g) => g.size));
      const text = joinGlyphs(g0);
      const center = (x0 + x1) / 2;
      const align = Math.abs(center - W / 2) < 12 ? "center" : x1 > W - m.mr - 8 && x0 > W / 2 ? "right" : "left";
      const font = g0[0].font.family;
      const color = g0[0].color && g0[0].color !== "000000" ? g0[0].color : null;
      const rPr = `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>${color ? `<w:color w:val="${color.toUpperCase()}"/>` : ""}<w:sz w:val="${Math.round(size * 2)}"/><w:szCs w:val="${Math.round(size * 2)}"/></w:rPr>`;
      let runs = "";
      const parts = text.split(/(\d+)/);
      for (const part of parts) {
        if (!part) continue;
        if (/^\d+$/.test(part) && (+part === pageNo || +part === total)) {
          const field = +part === pageNo ? "PAGE" : "NUMPAGES";
          runs += `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rPr}<w:instrText xml:space="preserve"> ${field} </w:instrText></w:r><w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rPr}<w:t>${part}</w:t></w:r><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`;
        } else {
          runs += `<w:r>${rPr}<w:t xml:space="preserve">${part.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</w:t></w:r>`;
        }
      }
      const L1 = size * 1.17;
      const ind = align === "left" ? `<w:ind w:left="${Math.round(Math.max(0, x0 - m.ml) * 20)}"/>` : "";
      xml += `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${Math.round(L1 * 20)}" w:lineRule="exact"/>${ind}<w:jc w:val="${align}"/></w:pPr>${runs}</w:p>`;
      if (zone === "bottom") dist = Math.max(4, H - (g0[0].y + (1 - NS.DOCX_BASE) * L1));
      else if (rows[0] === gl) dist = Math.max(4, g0[0].y - NS.DOCX_BASE * L1);
    }
    return { xml, dist };
  }

  function joinGlyphs(gl) {
    let s = "";
    let prev = null;
    for (const g of gl) {
      if (prev && !g.isSpace && !prev.isSpace && g.x - prev.x1 > g.size * 0.17) s += " ";
      s += g.isSpace ? " " : g.u;
      prev = g;
    }
    return s.replace(/ +/g, " ").trim();
  }

  NS.convertPdf = convertPdf;
  NS.yieldNow = yieldNow;
})(typeof globalThis !== "undefined" ? globalThis : window);
