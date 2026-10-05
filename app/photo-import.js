// Zdjęcie dokumentu (JPG/PNG/WEBP/HEIC) → PDF w pamięci → zwykła konwersja PDF → DOCX z OCR.
// Zdjęcie z telefonu to kartka na stole: pod kątem, z cieniem, z kawałkiem blatu. Dlatego:
//   1) kartka = jasny obszar wokół środka zdjęcia → 4 narożniki,
//   2) perspektywa wyprostowana (homografia, interpolacja dwuliniowa) — kartka staje się prostokątem,
//   3) tło wybielone jak w skanerze (jasność papieru w okolicy → biel; kolory pieczątek zostają),
//   4) każde zdjęcie = jedna strona PDF (JPEG), kilka zdjęć naraz = jeden dokument.
// Gdy kartki nie da się pewnie znaleźć (zrzut ekranu, skan, kartka na całe zdjęcie) — bierzemy
// całe zdjęcie; resztę (pochylenie, obrót) i tak prostuje OCR.
(function () {
  "use strict";

  const MAX_SIDE = 3200; // px — więcej nie poprawia OCR, a zjada pamięć (telefon)
  const A4 = 841.89 / 595.28;

  async function decode(file) {
    // createImageBitmap obraca wg EXIF (zdjęcie „bokiem” z telefonu); Image jako zapas
    if (typeof createImageBitmap === "function") {
      try {
        return await createImageBitmap(file, { imageOrientation: "from-image" });
      } catch (_) { /* np. HEIC w Chrome — spróbuj przez <img> */ }
    }
    const url = URL.createObjectURL(file);
    try {
      const im = new Image();
      im.src = url;
      await im.decode();
      return im;
    } finally {
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
  }

  function canvasOf(w, h) {
    const c = document.createElement("canvas");
    c.width = w;
    c.height = h;
    return c;
  }

  const luma = (r, g, b) => (r * 299 + g * 587 + b * 114) / 1000;

  // ── 1. kartka na zdjęciu ────────────────────────────────────────────────────────
  function otsu(g) {
    const hist = new Float64Array(256);
    for (let i = 0; i < g.length; i++) hist[g[i]]++;
    let sum = 0;
    for (let k = 0; k < 256; k++) sum += k * hist[k];
    let wB = 0, sumB = 0, best = 0, t = 128;
    for (let k = 0; k < 256; k++) {
      wB += hist[k];
      if (!wB) continue;
      const wF = g.length - wB;
      if (!wF) break;
      sumB += k * hist[k];
      const mB = sumB / wB, mF = (sum - sumB) / wF;
      const v = wB * wF * (mB - mF) * (mB - mF);
      if (v > best) { best = v; t = k; }
    }
    return t;
  }

  // Narożniki kartki (w pikselach zdjęcia) albo null.
  function findPage(src, W, H) {
    const f = Math.min(1, 640 / Math.max(W, H));
    const w = Math.max(8, Math.round(W * f)), h = Math.max(8, Math.round(H * f));
    const c = canvasOf(w, h);
    const x = c.getContext("2d", { willReadFrequently: true });
    x.imageSmoothingQuality = "high";
    x.drawImage(src, 0, 0, w, h);
    const d = x.getImageData(0, 0, w, h).data;
    const g = new Uint8Array(w * h);
    const sat = new Uint8Array(w * h); // nasycenie koloru 0–255 (papier ≈ szary, blat zwykle kolorowy)
    for (let i = 0, o = 0; i < g.length; i++, o += 4) {
      g[i] = luma(d[o], d[o + 1], d[o + 2]);
      const mx = Math.max(d[o], d[o + 1], d[o + 2]), mn = Math.min(d[o], d[o + 1], d[o + 2]);
      sat[i] = mx ? ((mx - mn) * 255) / mx : 0;
    }
    // rozmycie — druk na kartce nie może jej „dziurawić” na tyle, by rozpadła się na kawałki
    const b = new Uint8Array(g.length);
    const R = 3;
    for (let y = 0; y < h; y++) for (let xx = 0; xx < w; xx++) {
      let s = 0, n = 0;
      for (let dy = -R; dy <= R; dy += 1) {
        const yy = y + dy;
        if (yy < 0 || yy >= h) continue;
        for (let dx = -R; dx <= R; dx += 2) {
          const x2 = xx + dx;
          if (x2 < 0 || x2 >= w) continue;
          s += g[yy * w + x2]; n++;
        }
      }
      b[y * w + xx] = s / n;
    }
    const t = otsu(b);
    // Kartka = jasne piksele, a także jej część w CIENIU (ręka, telefon): ciemniejsza, ale dalej
    // szara, bez koloru — do połowy progu jasności. Kolorowy blat (drewno) się nie łapie.
    const bright = (i) => b[i] > t || (b[i] > t * 0.5 && sat[i] < 40);
    let seed = -1;
    const cx = w >> 1, cy = h >> 1;
    for (let r = 0; r < Math.min(w, h) / 3 && seed < 0; r += 2) {
      for (const [dx, dy] of [[0, 0], [r, 0], [-r, 0], [0, r], [0, -r]]) {
        const i = (cy + dy) * w + (cx + dx);
        if (bright(i)) { seed = i; break; }
      }
    }
    if (seed < 0) return null;
    const mark = new Uint8Array(w * h);
    const stack = [seed];
    mark[seed] = 1;
    let area = 0, edgeTouch = 0;
    let tl = [1e9, 0, 0], br = [-1e9, 0, 0], tr = [-1e9, 0, 0], bl = [1e9, 0, 0];
    let top = [0, 1e9], bot = [0, -1e9], lef = [1e9, 0], rig = [-1e9, 0];
    while (stack.length) {
      const i = stack.pop(), px = i % w, py = (i / w) | 0;
      area++;
      if (px === 0 || py === 0 || px === w - 1 || py === h - 1) edgeTouch++;
      if (px + py < tl[0]) tl = [px + py, px, py];
      if (px + py > br[0]) br = [px + py, px, py];
      if (px - py > tr[0]) tr = [px - py, px, py];
      if (px - py < bl[0]) bl = [px - py, px, py];
      if (py < top[1]) top = [px, py];
      if (py > bot[1]) bot = [px, py];
      if (px < lef[0]) lef = [px, py];
      if (px > rig[0]) rig = [px, py];
      for (const j of [i - 1, i + 1, i - w, i + w]) {
        if (j < 0 || j >= w * h || mark[j]) continue;
        if ((j === i - 1 && px === 0) || (j === i + 1 && px === w - 1)) continue;
        if (bright(j)) { mark[j] = 1; stack.push(j); }
      }
    }
    const frac = area / (w * h);
    // kartka zajmuje (prawie) całe zdjęcie albo jasne jest wszystko (biały blat) — bez kadrowania
    if (frac < 0.18 || frac > 0.97) return null;
    if (edgeTouch > (w + h) * 1.2) return null;
    const quadArea = (q) => {
      let s = 0;
      for (let k = 0; k < 4; k++) { const [x1, y1] = q[k], [x2, y2] = q[(k + 1) % 4]; s += x1 * y2 - x2 * y1; }
      return Math.abs(s) / 2;
    };
    // narożniki z przekątnych (kartka prosto lub lekko obrócona) albo ze skrajnych punktów (≈ 45°)
    const qa = [[tl[1], tl[2]], [tr[1], tr[2]], [br[1], br[2]], [bl[1], bl[2]]];
    const qb = [top, rig, bot, lef];
    const q = quadArea(qa) >= quadArea(qb) ? qa : qb;
    if (quadArea(q) < area * 0.85 || !convexOk(q)) return null;
    return q.map(([px, py]) => [(px + 0.5) / f, (py + 0.5) / f]);
  }

  // Czworokąt wypukły, kąty 50–130° — inaczej to nie kartka (cień, ręka, dwie kartki).
  function convexOk(q) {
    let sign = 0;
    for (let k = 0; k < 4; k++) {
      const [ax, ay] = q[k], [bx, by] = q[(k + 1) % 4], [cx, cy] = q[(k + 2) % 4];
      const cr = (bx - ax) * (cy - by) - (by - ay) * (cx - bx);
      if (!cr) return false;
      if (sign && Math.sign(cr) !== sign) return false;
      sign = Math.sign(cr);
      const v1 = [ax - bx, ay - by], v2 = [cx - bx, cy - by];
      const ang = (Math.acos((v1[0] * v2[0] + v1[1] * v2[1]) / (Math.hypot(...v1) * Math.hypot(...v2))) * 180) / Math.PI;
      if (ang < 50 || ang > 130) return false;
    }
    return true;
  }

  // ── 2. perspektywa ──────────────────────────────────────────────────────────────
  // Homografia: punkt prostokąta wyniku (u, v) → punkt zdjęcia (x, y). Układ 8×8 (eliminacja Gaussa).
  function homography(dst, src) {
    const A = [], B = [];
    for (let k = 0; k < 4; k++) {
      const [u, v] = dst[k], [x, y] = src[k];
      A.push([u, v, 1, 0, 0, 0, -u * x, -v * x]); B.push(x);
      A.push([0, 0, 0, u, v, 1, -u * y, -v * y]); B.push(y);
    }
    const n = 8;
    for (let i = 0; i < n; i++) {
      let p = i;
      for (let r = i + 1; r < n; r++) if (Math.abs(A[r][i]) > Math.abs(A[p][i])) p = r;
      [A[i], A[p]] = [A[p], A[i]];
      [B[i], B[p]] = [B[p], B[i]];
      for (let r = 0; r < n; r++) {
        if (r === i) continue;
        const m = A[r][i] / A[i][i];
        for (let c = i; c < n; c++) A[r][c] -= m * A[i][c];
        B[r] -= m * B[i];
      }
    }
    return B.map((v, i) => v / A[i][i]);
  }

  function warp(src, W, H, quad) {
    const len = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
    let ow = Math.max(len(quad[0], quad[1]), len(quad[3], quad[2]));
    let oh = Math.max(len(quad[0], quad[3]), len(quad[1], quad[2]));
    const s = Math.min(1, MAX_SIDE / Math.max(ow, oh));
    ow = Math.round(ow * s);
    oh = Math.round(oh * s);
    // proporcje prawie jak A4 → dokładnie A4 (perspektywa i tak je trochę przekłamuje)
    if (Math.abs(oh / ow / A4 - 1) < 0.08) oh = Math.round(ow * A4);
    else if (Math.abs(ow / oh / A4 - 1) < 0.08) ow = Math.round(oh * A4);
    const h8 = homography([[0, 0], [ow, 0], [ow, oh], [0, oh]], quad);
    const sc = canvasOf(W, H);
    const sx = sc.getContext("2d", { willReadFrequently: true });
    sx.drawImage(src, 0, 0, W, H);
    const sd = sx.getImageData(0, 0, W, H).data;
    sc.width = sc.height = 0;
    const out = canvasOf(ow, oh);
    const ox = out.getContext("2d", { willReadFrequently: true });
    const od = ox.createImageData(ow, oh);
    const d = od.data;
    const [a, b, c, dd, e, f, g, hh] = h8;
    for (let v = 0; v < oh; v++) {
      for (let u = 0, o = v * ow * 4; u < ow; u++, o += 4) {
        const uu = u + 0.5, vv = v + 0.5;
        const den = g * uu + hh * vv + 1;
        const x = (a * uu + b * vv + c) / den - 0.5, y = (dd * uu + e * vv + f) / den - 0.5;
        const x0 = Math.max(0, Math.min(W - 2, Math.floor(x))), y0 = Math.max(0, Math.min(H - 2, Math.floor(y)));
        const fx = Math.max(0, Math.min(1, x - x0)), fy = Math.max(0, Math.min(1, y - y0));
        const i = (y0 * W + x0) * 4, j = i + W * 4;
        for (let k = 0; k < 3; k++) {
          const t1 = sd[i + k] + (sd[i + 4 + k] - sd[i + k]) * fx;
          const t2 = sd[j + k] + (sd[j + 4 + k] - sd[j + k]) * fx;
          d[o + k] = t1 + (t2 - t1) * fy;
        }
        d[o + 3] = 255;
      }
    }
    ox.putImageData(od, 0, 0);
    return out;
  }

  function scaled(src, W, H) {
    const s = Math.min(1, MAX_SIDE / Math.max(W, H));
    const c = canvasOf(Math.round(W * s), Math.round(H * s));
    const x = c.getContext("2d", { willReadFrequently: true });
    x.imageSmoothingQuality = "high";
    x.drawImage(src, 0, 0, c.width, c.height);
    return c;
  }

  // ── 3. „jak ze skanera”: papier wybielony ───────────────────────────────────────
  // Jasność papieru w okolicy (percentyl 90 w blokach, między blokami liniowo) → każdy kanał
  // dzielony przez nią. Cień, żółte światło i szary papier znikają, a kolor druku/pieczątki zostaje.
  function whiten(canvas) {
    const w = canvas.width, h = canvas.height;
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    const id = ctx.getImageData(0, 0, w, h);
    const d = id.data;
    const B = Math.max(20, Math.round(Math.max(w, h) / 96));
    const bw = Math.ceil(w / B), bh = Math.ceil(h / B);
    // osobno dla każdego kanału — kartka w ciepłym świetle wychodzi biała, a nie kremowa
    const maps = [0, 1, 2].map(() => new Float32Array(bw * bh));
    const hist = [new Uint32Array(256), new Uint32Array(256), new Uint32Array(256)];
    for (let by = 0; by < bh; by++) for (let bx = 0; bx < bw; bx++) {
      hist.forEach((hh) => hh.fill(0));
      let n = 0;
      for (let y = by * B; y < Math.min(h, (by + 1) * B); y += 3) for (let x = bx * B; x < Math.min(w, (bx + 1) * B); x += 3) {
        const o = (y * w + x) * 4;
        hist[0][d[o]]++; hist[1][d[o + 1]]++; hist[2][d[o + 2]]++;
        n++;
      }
      for (let c = 0; c < 3; c++) {
        let k = 255, acc = 0;
        while (k > 0 && acc + hist[c][k] < n * 0.1) acc += hist[c][k--];
        maps[c][by * bw + bx] = Math.max(50, k);
      }
    }
    // Blok prawie cały w tuszu (zdjęcie, czarny pasek) bierze jasność sąsiada; blok w CIENIU
    // (ok. 60% jasności sąsiada) zostaje sobą — inaczej przy krawędzi cienia zostawał ząbkowany
    // szary pas, a słowa w nim nie dawały się odczytać.
    const smooth = (vals) => {
      const sm = vals.slice();
      for (let by = 0; by < bh; by++) for (let bx = 0; bx < bw; bx++) {
        const v = vals[by * bw + bx];
        let m = v;
        for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
          const yy = by + dy, xx = bx + dx;
          if (yy >= 0 && xx >= 0 && yy < bh && xx < bw) m = Math.max(m, vals[yy * bw + xx]);
        }
        sm[by * bw + bx] = v < m * 0.45 ? m * 0.97 : v;
      }
      return sm;
    };
    const sm = maps.map(smooth);
    for (let y = 0; y < h; y++) {
      const fy = Math.max(0, Math.min(bh - 1.001, y / B - 0.5)), iy = fy | 0, ay = fy - iy;
      for (let x = 0; x < w; x++) {
        const fx = Math.max(0, Math.min(bw - 1.001, x / B - 0.5)), ix = fx | 0, ax = fx - ix;
        const i = iy * bw + ix, j = Math.min(i + 1, iy * bw + bw - 1), k = Math.min(i + bw, bw * bh - 1), l = Math.min(k + 1, bw * bh - 1);
        const o = (y * w + x) * 4;
        for (let c = 0; c < 3; c++) {
          const m = sm[c];
          const bg = (m[i] * (1 - ax) + m[j] * ax) * (1 - ay) + (m[k] * (1 - ax) + m[l] * ax) * ay;
          d[o + c] = Math.min(255, (d[o + c] * 255) / bg);
        }
      }
    }
    ctx.putImageData(id, 0, 0);
  }

  async function jpegBytes(canvas) {
    const blob = await new Promise((r) => canvas.toBlob(r, "image/jpeg", 0.88));
    return new Uint8Array(await blob.arrayBuffer());
  }

  // ── 4. PDF: strona = obraz JPEG na całą stronę (szerokość A4) ──────────────────────
  function buildPdf(pages) {
    const enc = new TextEncoder();
    const parts = [];
    const offsets = [];
    let pos = 0;
    const push = (chunk) => {
      const b = typeof chunk === "string" ? enc.encode(chunk) : chunk;
      parts.push(b);
      pos += b.length;
    };
    push("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");
    const obj = (n, body, stream) => {
      offsets[n] = pos;
      if (stream) {
        push(`${n} 0 obj\n${body}\nstream\n`);
        push(stream);
        push("\nendstream\nendobj\n");
      } else push(`${n} 0 obj\n${body}\nendobj\n`);
    };
    const n = pages.length;
    const kids = pages.map((_, i) => `${3 + i * 3} 0 R`).join(" ");
    obj(1, "<< /Type /Catalog /Pages 2 0 R >>");
    obj(2, `<< /Type /Pages /Kids [${kids}] /Count ${n} >>`);
    pages.forEach((p, i) => {
      const pg = 3 + i * 3, img = pg + 1, cs = pg + 2;
      const W = 595.28, H = Math.round((W * p.h) / p.w * 100) / 100;
      const content = `q ${W} 0 0 ${H} 0 0 cm /Im0 Do Q`;
      obj(pg, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${W} ${H}] /Resources << /XObject << /Im0 ${img} 0 R >> >> /Contents ${cs} 0 R >>`);
      obj(img, `<< /Type /XObject /Subtype /Image /Width ${p.w} /Height ${p.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.jpeg.length} >>`, p.jpeg);
      obj(cs, `<< /Length ${content.length} >>`, enc.encode(content));
    });
    const total = 3 + n * 3;
    const xref = pos;
    let x = `xref\n0 ${total}\n0000000000 65535 f \n`;
    for (let k = 1; k < total; k++) x += `${String(offsets[k]).padStart(10, "0")} 00000 n \n`;
    push(x);
    push(`trailer\n<< /Size ${total} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
    const out = new Uint8Array(pos);
    let o = 0;
    for (const b of parts) { out.set(b, o); o += b.length; }
    return out;
  }

  // Jedno zdjęcie → strona (canvas po obróbce). info: { cropped }
  async function processPhoto(file) {
    const bmp = await decode(file);
    const W = bmp.width || bmp.naturalWidth, H = bmp.height || bmp.naturalHeight;
    if (!W || !H) throw new Error("decode");
    const quad = findPage(bmp, W, H);
    const canvas = quad ? warp(bmp, W, H, quad) : scaled(bmp, W, H);
    if (bmp.close) bmp.close();
    whiten(canvas);
    return { canvas, cropped: !!quad };
  }

  /**
   * Zdjęcia → PDF (File). Rzuca Error("decode") dla formatu, którego przeglądarka nie odczyta
   * (np. HEIC poza Safari).
   * @returns {Promise<{ file: File, cropped: number, pages: number }>}
   */
  async function toPdf(files, name) {
    const pages = [];
    let cropped = 0;
    for (const f of files) {
      const { canvas, cropped: c } = await processPhoto(f);
      if (c) cropped++;
      pages.push({ jpeg: await jpegBytes(canvas), w: canvas.width, h: canvas.height });
      canvas.width = canvas.height = 0;
      await new Promise((r) => setTimeout(r, 0));
    }
    const bytes = buildPdf(pages);
    return { file: new File([bytes], name, { type: "application/pdf" }), cropped, pages: pages.length };
  }

  window.dwbPhotoImport = { toPdf, _findPage: findPage, _buildPdf: buildPdf };
})();
