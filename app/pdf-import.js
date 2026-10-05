// PDF → DOCX w aplikacji: okno z postępem (Anuluj, hasło), pdf.js w workerze, obrazy przez
// kanwę, rozpoznawanie kształtu liter o złych kodach, warstwa grafiki wektorowej. Gotowy .docx
// trafia do ingestFile jako nowy, niezapisany dokument. Nic nie wychodzi z urządzenia.
(function () {
  "use strict";

  const SCRIPTS = ["app/pdf-extract.js", "app/pdf-layout.js", "app/pdf-docx.js", "app/pdf-fonts.js", "app/pdf-convert.js"];
  const PDFJS_DIR = "lib/pdfjs/";
  let enginePromise = null;
  let busy = false;

  const T = (pl, en) => (typeof currentLang === "string" && currentLang === "en" ? en : pl);
  const baseUrl = (rel) => new URL(rel, document.baseURI).href;
  const assetV = () => (typeof lazyAssetVersion === "function" ? lazyAssetVersion() : "1");

  async function loadEngine() {
    if (!enginePromise) {
      enginePromise = (async () => {
        if (typeof ensureDocLibs === "function") await ensureDocLibs(true);
        for (const s of SCRIPTS) {
          if (window.DWPdf && window.DWPdf[{ "app/pdf-extract.js": "extractPage", "app/pdf-layout.js": "layoutPage", "app/pdf-docx.js": "buildDocx", "app/pdf-fonts.js": "buildEmbeddedFonts", "app/pdf-convert.js": "convertPdf" }[s]]) continue;
          await loadLazyScript(s);
        }
        const pdfjs = await import(baseUrl(PDFJS_DIR + "pdf.min.mjs?v=" + encodeURIComponent(assetV())));
        pdfjs.GlobalWorkerOptions.workerSrc = baseUrl(PDFJS_DIR + "pdf.worker.min.mjs?v=" + encodeURIComponent(assetV()));
        return pdfjs;
      })().catch((e) => {
        enginePromise = null;
        throw e;
      });
    }
    return enginePromise;
  }

  // ------------------------------------------------------------------ okno postępu
  let dlg = null;
  function ensureDialog() {
    if (dlg) return dlg;
    const d = document.createElement("dialog");
    d.className = "pdfconv-dialog";
    d.id = "pdfConvDialog";
    d.setAttribute("aria-labelledby", "pdfConvTitle");
    d.innerHTML = `
      <div class="pdfconv-inner">
        <h2 id="pdfConvTitle">${T("Konwersja PDF → Word", "PDF → Word conversion")} <span class="beta-tag">beta</span></h2>
        <p class="pdfconv-file" id="pdfConvFile"></p>
        <div class="pdfconv-bar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0" id="pdfConvBar"><i></i></div>
        <p class="pdfconv-phase" id="pdfConvPhase" aria-live="polite"></p>
        <form class="pdfconv-pass hidden" id="pdfConvPass">
          <label for="pdfConvPassInput" id="pdfConvPassLabel">${T("Ten PDF jest chroniony hasłem:", "This PDF is password-protected:")}</label>
          <div class="pdfconv-pass-row">
            <input type="password" id="pdfConvPassInput" autocomplete="off" enterkeyhint="go" />
            <button class="btn primary" type="submit">${T("Otwórz", "Open")}</button>
          </div>
        </form>
        <div class="pdfconv-ask hidden" id="pdfConvAsk">
          <p id="pdfConvAskText">${T("Ten PDF to skan (zdjęcie strony). Rozpoznać tekst, żeby dało się go edytować?", "This PDF is a scan (an image of the page). Recognize the text so it can be edited?")}</p>
          <p class="pdfconv-ask-sub">${T("Rozpoznawanie (OCR) działa na tym urządzeniu, kilka sekund na stronę. Pismo odręczne zostaje jako obraz.", "Recognition (OCR) runs on this device, a few seconds per page. Handwriting stays as an image.")}</p>
          <div class="pdfconv-ask-row">
            <button class="btn primary" type="button" id="pdfConvOcrYes">${T("Rozpoznaj tekst", "Recognize text")}</button>
            <button class="btn" type="button" id="pdfConvOcrNo">${T("Zostaw jako obraz", "Keep as image")}</button>
          </div>
        </div>
        <p class="pdfconv-note">${T("Plik jest przetwarzany na tym urządzeniu — nigdzie nie jest wysyłany.", "The file is processed on this device — it is never uploaded.")}</p>
        <div class="pdfconv-actions"><button class="btn" type="button" id="pdfConvCancel">${T("Anuluj", "Cancel")}</button></div>
      </div>`;
    document.body.appendChild(d);
    d.addEventListener("cancel", (e) => {
      e.preventDefault();
      d.querySelector("#pdfConvCancel").click();
    });
    dlg = d;
    return d;
  }

  function setProgress(fraction, text) {
    const d = ensureDialog();
    const pct = Math.max(0, Math.min(100, Math.round(fraction * 100)));
    const bar = d.querySelector("#pdfConvBar");
    bar.style.setProperty("--p", pct + "%");
    bar.setAttribute("aria-valuenow", String(pct));
    if (text != null) d.querySelector("#pdfConvPhase").textContent = text;
  }

  function askPassword(wrong) {
    const d = ensureDialog();
    const form = d.querySelector("#pdfConvPass");
    const input = d.querySelector("#pdfConvPassInput");
    d.querySelector("#pdfConvPassLabel").textContent = wrong ? T("Hasło niepoprawne — spróbuj jeszcze raz:", "Wrong password — try again:") : T("Ten PDF jest chroniony hasłem:", "This PDF is password-protected:");
    form.classList.remove("hidden");
    input.value = "";
    setTimeout(() => input.focus(), 30);
    return new Promise((resolve) => {
      const onSubmit = (e) => {
        e.preventDefault();
        form.removeEventListener("submit", onSubmit);
        form.classList.add("hidden");
        resolve(input.value);
        input.value = "";
      };
      form.addEventListener("submit", onSubmit);
    });
  }

  // kind: "scan" (obraz strony) albo "curves" (tekst zapisany jako rysunek — same kształty liter)
  function askOcr(kind = "scan") {
    const d = ensureDialog();
    const box = d.querySelector("#pdfConvAsk");
    const txt = d.querySelector("#pdfConvAskText");
    if (txt) txt.textContent = kind === "photo"
      ? T("To zdjęcie dokumentu (kartka wyprostowana). Rozpoznać tekst, żeby dało się go edytować?", "This is a photo of a document (page straightened). Recognize the text so it can be edited?")
      : kind === "curves"
      ? T("Tekst w tym PDF jest zapisany jako rysunek (kształty liter, bez znaków). Rozpoznać go, żeby dało się go edytować i przeszukiwać?", "The text in this PDF is stored as drawings (letter shapes, no characters). Recognize it so it can be edited and searched?")
      : T("Ten PDF to skan (zdjęcie strony). Rozpoznać tekst, żeby dało się go edytować?", "This PDF is a scan (an image of the page). Recognize the text so it can be edited?");
    box.classList.remove("hidden");
    const yes = d.querySelector("#pdfConvOcrYes"), no = d.querySelector("#pdfConvOcrNo");
    setTimeout(() => yes.focus(), 30);
    return new Promise((resolve) => {
      const done = (v) => {
        yes.removeEventListener("click", onYes);
        no.removeEventListener("click", onNo);
        box.classList.add("hidden");
        resolve(v);
      };
      const onYes = () => done(true);
      const onNo = () => done(false);
      yes.addEventListener("click", onYes);
      no.addEventListener("click", onNo);
    });
  }

  // Ile stron OCR naraz: komputer (mysz, ≥ 4 rdzenie, ≥ 4 GB) — 2 wątki, ok. 2× szybciej przy
  // wielu stronach; telefon/tablet — 1 (iOS ma ostry limit pamięci kanw, a strona to kilka kanw).
  const OCR_PARALLEL = (() => {
    try {
      const coarse = window.matchMedia("(pointer: coarse)").matches;
      return !coarse && (navigator.hardwareConcurrency || 2) >= 4 && (navigator.deviceMemory || 8) >= 4 ? 2 : 1;
    } catch (_) {
      return 1;
    }
  })();

  async function ocrPage(page, raw, onProgress, opts = {}) {
    if (!window.DWPdfOcr) await loadLazyScript("app/pdf-ocr.js");
    window.DWPdfOcr.setParallel(window.__dwbOcrParallel || OCR_PARALLEL); // __dwbOcrParallel: miernik ocr:score
    return window.DWPdfOcr.recognizePage(page, raw, { ...opts, onProgress, signal: currentSignal });
  }
  let currentSignal = null;

  // Miejsca rozpoznanych słów na obrazie skanu: wypełnienie kolorem tła (mediana pikseli wokół).
  function eraseOnCanvas(ctx, boxes, s, box, cw, ch) {
    if (!boxes || !boxes.length) return;
    const img = ctx.getImageData(0, 0, cw, ch);
    const d = img.data;
    for (const b of boxes) {
      const x0 = Math.max(0, Math.floor((b.x0 - box.x0) * s)), x1 = Math.min(cw - 1, Math.ceil((b.x1 - box.x0) * s));
      const y0 = Math.max(0, Math.floor((b.y0 - box.y0) * s)), y1 = Math.min(ch - 1, Math.ceil((b.y1 - box.y0) * s));
      if (x1 <= x0 || y1 <= y0) continue;
      const ring = [];
      const take = (x, y) => {
        if (x < 0 || y < 0 || x >= cw || y >= ch) return;
        const o = (y * cw + x) * 4;
        ring.push([d[o], d[o + 1], d[o + 2]]);
      };
      for (let x = x0; x <= x1; x += 2) {
        take(x, y0 - 2);
        take(x, y1 + 2);
      }
      for (let y = y0; y <= y1; y += 2) {
        take(x0 - 2, y);
        take(x1 + 2, y);
      }
      if (!ring.length) continue;
      // jasność: bierzemy jaśniejszą połowę (tło), nie przypadkowe kreski przy słowie
      ring.sort((a, b) => b[0] + b[1] + b[2] - (a[0] + a[1] + a[2]));
      const top = ring.slice(0, Math.max(1, Math.floor(ring.length * 0.6)));
      const col = [0, 1, 2].map((k) => top.map((p) => p[k]).sort((a, b) => a - b)[top.length >> 1]);
      for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) {
        const o = (y * cw + x) * 4;
        d[o] = col[0]; d[o + 1] = col[1]; d[o + 2] = col[2]; d[o + 3] = 255;
      }
    }
    ctx.putImageData(img, 0, 0);
  }

  // ------------------------------------------------------------------ obrazy
  function getObj(page, id) {
    const objs = String(id).startsWith("g_") ? page.commonObjs : page.objs;
    return new Promise((resolve) => {
      try {
        objs.get(id, resolve);
      } catch (_) {
        resolve(null);
      }
      setTimeout(() => resolve(null), 15000);
    });
  }

  function imageSource(obj, maskColor) {
    if (!obj) return null;
    // ImageBitmap albo VideoFrame (pdf.js dekoduje JPEG przez ImageDecoder — wymiary w displayWidth)
    const bmp = obj.bitmap;
    if (bmp) {
      const bw = bmp.displayWidth || bmp.width || obj.width, bh = bmp.displayHeight || bmp.height || obj.height;
      if (bw > 0 && bh > 0) return { src: bmp, w: bw, h: bh, opaque: !maskColor && obj.kind !== 3 };
    }
    const { width: w, height: h, data } = obj;
    let kind = obj.kind;
    if (!kind && data) kind = data.length >= w * h * 4 ? 3 : data.length >= w * h * 3 ? 2 : 1;
    if (!w || !h || !data) return null;
    const c = new OffscreenCanvasOr(w, h);
    const ctx = c.getContext("2d");
    const id = ctx.createImageData(w, h);
    const out = id.data;
    if (maskColor) {
      // maska 1-bitowa: bit 0 = kolor wypełnienia, bit 1 = przezroczyste (jak w pdf.js)
      const r = parseInt(maskColor.slice(0, 2), 16), g = parseInt(maskColor.slice(2, 4), 16), b = parseInt(maskColor.slice(4, 6), 16);
      const rowBytes = (w + 7) >> 3;
      for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
        const bit = (data[y * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1;
        if (!bit) {
          const o = (y * w + x) * 4;
          out[o] = r; out[o + 1] = g; out[o + 2] = b; out[o + 3] = 255;
        }
      }
    } else if (kind === 3) out.set(data.subarray(0, w * h * 4));
    else if (kind === 2) {
      for (let i = 0, j = 0; i < w * h; i++, j += 3) {
        out[i * 4] = data[j]; out[i * 4 + 1] = data[j + 1]; out[i * 4 + 2] = data[j + 2]; out[i * 4 + 3] = 255;
      }
    } else if (kind === 1) {
      const rowBytes = (w + 7) >> 3;
      for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
        const v = (data[y * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1 ? 255 : 0;
        const o = (y * w + x) * 4;
        out[o] = out[o + 1] = out[o + 2] = v; out[o + 3] = 255;
      }
    } else return null;
    ctx.putImageData(id, 0, 0);
    return { src: c, w, h, opaque: !maskColor && kind !== 3 };
  }

  function OffscreenCanvasOr(w, h) {
    if (typeof OffscreenCanvas !== "undefined") return new OffscreenCanvas(w, h);
    const c = document.createElement("canvas");
    c.width = w;
    c.height = h;
    return c;
  }

  async function canvasToBytes(c, type, quality) {
    let blob;
    if (c.convertToBlob) blob = await c.convertToBlob({ type, quality });
    else blob = await new Promise((r) => c.toBlob(r, type, quality));
    return new Uint8Array(await blob.arrayBuffer());
  }

  const mul = (a, b) => [a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3], a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]];

  async function encodeImage(page, img) {
    let obj = null;
    if (img.objId) obj = await getObj(page, img.objId);
    else if (img.inline) obj = img.inline.data && typeof img.inline.data === "string" ? await getObj(page, img.inline.data) : img.inline;
    const srcInfo = imageSource(obj, img.kind === "mask" ? img.color || "000000" : null);
    if (!srcInfo) return null;
    const box = img.clipped || img;
    const bw = box.x1 - box.x0, bh = box.y1 - box.y0;
    // rozdzielczość: natywna obrazu (ale nie więcej niż 4 px/pt i 3000 px bok)
    const fullW = img.x1 - img.x0, fullH = img.y1 - img.y0;
    let s = Math.max(srcInfo.w / Math.max(1, fullW), srcInfo.h / Math.max(1, fullH));
    s = Math.max(1, Math.min(4, s));
    s = Math.min(s, 3000 / Math.max(bw, bh));
    const cw = Math.max(1, Math.round(bw * s)), ch = Math.max(1, Math.round(bh * s));
    const c = OffscreenCanvasOr(cw, ch);
    const ctx = c.getContext("2d");
    if (srcInfo.opaque && img.kind !== "mask") {
      ctx.fillStyle = img.rotateM && img.paper ? `rgb(${img.paper.join(",")})` : "#fff";
      ctx.fillRect(0, 0, cw, ch);
    }
    ctx.imageSmoothingQuality = "high";
    // piksel obrazu (px, py) → jednostka (px/w, 1 − py/h) → strona (macierz z PDF) → kanwa
    // rotateM: skan wyprostowany przez OCR (pdf-convert straightenScanPage) — strona → strona prosta
    const M = mul(img.rotateM || [1, 0, 0, 1, 0, 0], mul(img.matrix, [1 / srcInfo.w, 0, 0, -1 / srcInfo.h, 0, 1]));
    const A = mul([s, 0, 0, s, -box.x0 * s, -box.y0 * s], M);
    ctx.setTransform(A[0], A[1], A[2], A[3], A[4], A[5]);
    ctx.drawImage(srcInfo.src, 0, 0);
    if (img.eraseBoxes) {
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      eraseOnCanvas(ctx, img.eraseBoxes, s, box, cw, ch);
    }
    const photo = srcInfo.opaque && img.kind !== "mask" && cw * ch > 40000;
    const bytes = await canvasToBytes(c, photo ? "image/jpeg" : "image/png", photo ? 0.9 : undefined);
    return { bytes, ext: photo ? "jpeg" : "png" };
  }

  // ------------------------------------------------------------------ grafika wektorowa
  async function renderVectors(items, W, H) {
    if (!items.length) return null;
    const s = Math.min(3, 2400 / Math.max(W, H));
    // znaki-rysunki rysujemy czcionkami pdf.js z document.fonts — kanwa „poza ekranem” w Safari
    // potrafi ich nie widzieć, więc wtedy zwykła kanwa
    const hasGlyphs = items.some((v) => v.type === "glyph");
    const c = hasGlyphs ? Object.assign(document.createElement("canvas"), { width: Math.ceil(W * s), height: Math.ceil(H * s) }) : OffscreenCanvasOr(Math.ceil(W * s), Math.ceil(H * s));
    const ctx = c.getContext("2d");
    ctx.scale(s, s);
    let drawn = 0;
    for (const v of items.slice().sort((a, b) => (a.order || 0) - (b.order || 0))) {
      ctx.save();
      ctx.globalAlpha = v.alpha ?? 1;
      if (v.clip) {
        ctx.beginPath();
        ctx.rect(v.clip.x0, v.clip.y0, v.clip.x1 - v.clip.x0, v.clip.y1 - v.clip.y0);
        ctx.clip();
      }
      if (v.type === "rect") {
        ctx.fillStyle = "#" + v.color;
        ctx.fillRect(v.x0, v.y0, Math.max(0.3, v.x1 - v.x0), Math.max(0.3, v.y1 - v.y0));
        drawn++;
      } else if (v.type === "line") {
        ctx.strokeStyle = "#" + v.color;
        ctx.lineWidth = Math.max(0.3, v.w || 0.5);
        ctx.beginPath();
        ctx.moveTo(v.x0, v.y0);
        ctx.lineTo(v.x1, v.y1);
        ctx.stroke();
        drawn++;
      } else if (v.type === "glyph") {
        ctx.fillStyle = "#" + v.color;
        ctx.font = `${v.size}px "${v.fontRef}"`;
        ctx.textBaseline = "alphabetic";
        ctx.translate(v.x, v.y);
        if (v.angle) ctx.rotate(v.angle);
        if (v.xScale && Math.abs(v.xScale - 1) > 0.01) ctx.scale(v.xScale, 1);
        ctx.fillText(v.fontChar, 0, 0);
        drawn++;
      } else if (v.type === "path" && v.cmds) {
        ctx.beginPath();
        for (const cmd of v.cmds) {
          if (cmd[0] === "M") ctx.moveTo(cmd[1][0], cmd[1][1]);
          else if (cmd[0] === "L") ctx.lineTo(cmd[1][0], cmd[1][1]);
          else if (cmd[0] === "C") ctx.bezierCurveTo(cmd[1][0], cmd[1][1], cmd[2][0], cmd[2][1], cmd[3][0], cmd[3][1]);
          else if (cmd[0] === "Q") ctx.quadraticCurveTo(cmd[1][0], cmd[1][1], cmd[2][0], cmd[2][1]);
          else if (cmd[0] === "Z") ctx.closePath();
        }
        if (v.kind === "fill") {
          ctx.fillStyle = "#" + v.color;
          ctx.fill(v.evenOdd ? "evenodd" : "nonzero");
        } else {
          ctx.strokeStyle = "#" + v.color;
          ctx.lineWidth = Math.max(0.3, v.width || 0.5);
          ctx.stroke();
        }
        drawn++;
      }
      ctx.restore();
    }
    if (!drawn) return null;
    return { bytes: await canvasToBytes(c, "image/png"), ext: "png" };
  }

  // ------------------------------------------------------------------ rozpoznawanie liter
  // Litery bez wiarygodnego Unicode (nazwy glifów typu „167lc”, kody prywatne) rozpoznajemy po
  // kształcie: glif = znana litera tej samej czcionki + znak diakrytyczny (ą = a + ogonek).
  const STANDARD_NAME = /^(?:[A-Za-z]+(?:\.[A-Za-z0-9]+)?|uni[0-9A-Fa-f]{4,}|u[0-9A-Fa-f]{4,6}|[a-z]+[0-9]+(?:\.\w+)?)$/;
  function suspiciousName(n) {
    if (!n || typeof n !== "string") return false;
    if (/^\d/.test(n)) return true;
    if (/^(?:g|glyph|cid|index|c|G|a)\d+$/.test(n)) return true;
    return !STANDARD_NAME.test(n);
  }

  const glyphCache = new Map();

  // Zwykła kanwa (nie OffscreenCanvas): czcionki pdf.js są w document.fonts, a kanwa „poza
  // ekranem” w Safari potrafi ich nie widzieć — wtedy ą/ę z PDF zostawały jako „¨”/„ª”.
  let maskCanvas = null;
  function renderMask(family, ch, px) {
    const W = px * 3, H = px * 3;
    if (!maskCanvas) maskCanvas = document.createElement("canvas");
    const c = maskCanvas;
    c.width = W;
    c.height = H;
    const ctx = c.getContext("2d", { willReadFrequently: true });
    ctx.font = `${px}px "${family}"`;
    ctx.fillStyle = "#000";
    ctx.textBaseline = "alphabetic";
    const ox = px, oy = px * 2;
    ctx.fillText(ch, ox, oy);
    const d = ctx.getImageData(0, 0, W, H).data;
    const m = new Uint8Array(W * H);
    let n = 0, x0 = W, y0 = H, x1 = -1, y1 = -1;
    for (let i = 0; i < W * H; i++) if (d[i * 4 + 3] > 110) {
      m[i] = 1;
      n++;
      const x = i % W, y = (i / W) | 0;
      if (x < x0) x0 = x;
      if (x > x1) x1 = x;
      if (y < y0) y0 = y;
      if (y > y1) y1 = y;
    }
    return { m, W, H, n, x0, y0, x1, y1, base: oy, px };
  }

  function dilate(mask) {
    const { m, W, H } = mask;
    const out = new Uint8Array(m.length);
    for (let y = 1; y < H - 1; y++) for (let x = 1; x < W - 1; x++) {
      const i = y * W + x;
      if (m[i] || m[i - 1] || m[i + 1] || m[i - W] || m[i + W]) out[i] = 1;
    }
    return out;
  }

  // Residuum (U bez L) → znak łączący
  function classifyMark(U, L) {
    const dl = dilate(L);
    const { W } = U;
    const pts = [];
    for (let i = 0; i < U.m.length; i++) if (U.m[i] && !dl[i]) pts.push([i % W, (i / W) | 0]);
    if (pts.length < Math.max(4, U.n * 0.03)) return null;
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity, sx = 0, sy = 0;
    for (const [x, y] of pts) {
      x0 = Math.min(x0, x); x1 = Math.max(x1, x); y0 = Math.min(y0, y); y1 = Math.max(y1, y);
      sx += x; sy += y;
    }
    const cx = sx / pts.length, cy = sy / pts.length;
    const em = U.px;
    const lTop = L.y0, lBot = L.y1, base = U.base;
    if (cy > base - em * 0.02) {
      // pod linią bazową: ogonek z prawej, cedylla na środku
      const rel = (cx - L.x0) / Math.max(1, L.x1 - L.x0);
      return rel > 0.55 ? "̨" : "̧";
    }
    if (cy < lTop + (lBot - lTop) * 0.08 && y1 <= lTop + em * 0.08) {
      // nad literą: dwa skupiska = umlaut; ukośna kreska = ostry/słaby; zwarta kropka = kropka
      const comps = countComponents(pts, W);
      if (comps >= 2) {
        const w = x1 - x0, h = y1 - y0;
        if (w > h * 1.6) return "̈";
      }
      let sxy = 0, sxx = 0, syy = 0;
      for (const [x, y] of pts) {
        sxy += (x - cx) * (y - cy);
        sxx += (x - cx) ** 2;
        syy += (y - cy) ** 2;
      }
      const corr = sxy / Math.sqrt(sxx * syy || 1);
      const w = x1 - x0 + 1, h = y1 - y0 + 1;
      const fill = pts.length / (w * h);
      if (fill > 0.55 && w / h > 0.55 && w / h < 1.8) return "̇";
      if (corr < -0.45) return "́"; // od lewego dołu do prawej góry (oś y w dół) = ostry
      if (corr > 0.45) return "̀";
      // szeroki znak: daszek/haczyk wg tego, gdzie jest masa
      if (w > h * 1.2) return cy > (y0 + y1) / 2 ? "̌" : "̂";
      return "́";
    }
    // przez literę: kreska (ł, đ, ø)
    if (cy > lTop && cy < lBot) return "STROKE";
    return null;
  }

  function countComponents(pts, W) {
    const set = new Set(pts.map(([x, y]) => y * W + x));
    let comps = 0;
    const seen = new Set();
    for (const k of set) {
      if (seen.has(k)) continue;
      comps++;
      const stack = [k];
      seen.add(k);
      while (stack.length) {
        const q = stack.pop();
        for (const d of [1, -1, W, -W, W + 1, W - 1, -W + 1, -W - 1]) {
          const n = q + d;
          if (set.has(n) && !seen.has(n)) {
            seen.add(n);
            stack.push(n);
          }
        }
      }
      if (comps > 3) break;
    }
    return comps;
  }

  const STROKE_MAP = { l: "ł", L: "Ł", d: "đ", D: "Đ", o: "ø", O: "Ø", h: "ħ", b: "ƀ", z: "ƶ", Z: "Ƶ", t: "ŧ" };

  async function glyphFixes(page, raw) {
    const DW = window.DWPdf;
    const fixes = new Map();
    const fonts = new Map(); // fontRef → { font, suspects: Map(code → glyph), letters: Map(u → glyph) }
    for (const g of raw.glyphs) {
      // znak z kodem sterującym (paint) bywa literą z zepsutym kodem — „Ę” jako „\n” (DC-85)
      if (!g.fontRef || (g.isSpace && !g.paint)) continue;
      let f = fonts.get(g.fontRef);
      if (!f) {
        let fontObj = null;
        try {
          fontObj = page.commonObjs.get(g.fontRef);
        } catch (_) { /* brak */ }
        f = { fontObj, suspects: new Map(), letters: new Map() };
        fonts.set(g.fontRef, f);
      }
      const name = f.fontObj?.differences?.[g.code];
      const cp = g.u.codePointAt(0);
      const pua = cp >= 0xe000 && cp <= 0xf8ff && !g.font.symbolic;
      if (suspiciousName(name) || pua || g.u === "�" || g.paint) f.suspects.set(g.code, g);
      else if (/^\p{L}$/u.test(g.u) && !f.letters.has(g.u)) f.letters.set(g.u, g);
    }
    for (const [ref, f] of fonts) {
      if (!f.suspects.size || !f.fontObj) continue;
      const family = f.fontObj.loadedName;
      try {
        await document.fonts.load(`64px "${family}"`);
      } catch (_) { /* czcionka mogła nie wejść do document.fonts */ }
      for (const [code, g] of f.suspects) {
        const key = family + "|" + code;
        if (glyphCache.has(key)) {
          const v = glyphCache.get(key);
          if (v) fixes.set(ref + "|" + code, v);
          continue;
        }
        let result = null;
        try {
          const U = renderMask(family, g.fontChar, 64);
          // pełny prostokąt (szare tło malowane znakami) to nie litera — każda litera „mieści
          // się” w prostokącie i dostałby wtedy przypadkową literę z ogonkiem
          const solid = U.n > 8 && U.n / ((U.x1 - U.x0 + 1) * (U.y1 - U.y0 + 1)) > 0.8;
          if (U.n > 8 && !solid) {
            let best = null;
            for (const [u, lg] of f.letters) {
              const Lm = renderMask(family, lg.fontChar, 64);
              if (Lm.n < 8 || Lm.n > U.n * 1.02) continue;
              let inter = 0;
              for (let i = 0; i < Lm.m.length; i++) if (Lm.m[i] && U.m[i]) inter++;
              const cover = inter / Lm.n;
              if (cover >= 0.9 && (!best || Lm.n > best.L.n)) best = { u, L: Lm, cover };
            }
            if (best) {
              const extra = (U.n - best.L.n) / U.n;
              if (extra < 0.04) result = best.u;
              else {
                const mark = classifyMark(U, best.L);
                if (mark === "STROKE") result = STROKE_MAP[best.u] || null;
                else if (mark) {
                  // Znak pod „a/e” to w praktyce ogonek (ą, ę), pod „c/s/t” — cedylla; kształt
                  // ogonka pod „e” bywa brany za cedyllę i wychodziło „ȩ” zamiast „ę”.
                  let mk = mark;
                  if ((mk === "\u0327" || mk === "\u0328") && /[aeiuAEIU]/.test(best.u)) mk = "\u0328";
                  if ((mk === "\u0327" || mk === "\u0328") && /[cstCST]/.test(best.u)) mk = "\u0327";
                  const comp = (best.u + mk).normalize("NFC");
                  if ([...comp].length === 1) result = comp;
                }
              }
            }
          }
        } catch (_) { result = null; }
        glyphCache.set(key, result);
        if (result) fixes.set(ref + "|" + code, result);
        await DW.yieldNow();
      }
    }
    return fixes.size ? fixes : null;
  }

  // ------------------------------------------------------------------ główna ścieżka
  async function convertFile(file, options = {}) {
    if (busy) return false;
    busy = true;
    const d = ensureDialog();
    const signal = { cancelled: false };
    let loadingTask = null;
    const cancelBtn = d.querySelector("#pdfConvCancel");
    // „Anuluj” musi zadziałać od razu: po zamknięciu pdf.js w trakcie odczytu strony obietnica
    // potrafi nie wrócić nigdy (WebKit) — dlatego wyścig z osobną obietnicą anulowania.
    let rejectCancel = null;
    const cancelled = new Promise((_, rej) => (rejectCancel = rej));
    cancelled.catch(() => {});
    const onCancel = () => {
      if (signal.cancelled) return;
      signal.cancelled = true;
      rejectCancel(new (window.DWPdf?.CancelledError || Error)("cancelled"));
      setTimeout(() => {
        try {
          loadingTask?.destroy();
        } catch (_) { /* już zamknięty */ }
      }, 0);
    };
    cancelBtn.addEventListener("click", onCancel);
    d.querySelector("#pdfConvFile").textContent = options.photo?.name || file.name || "dokument.pdf";
    d.querySelector("#pdfConvPass").classList.add("hidden");
    setProgress(0.02, T("Wczytywanie silnika PDF…", "Loading the PDF engine…"));
    if (!d.open) {
      try {
        d.showModal();
      } catch (_) {
        d.setAttribute("open", "");
      }
    }
    const t0 = performance.now();
    const race = (p) => Promise.race([p, cancelled]);
    try {
      const pdfjs = await race(loadEngine());
      if (signal.cancelled) throw new window.DWPdf.CancelledError();
      const data = new Uint8Array(await race(file.arrayBuffer()));
      setProgress(0.05, T("Otwieranie pliku…", "Opening the file…"));
      loadingTask = pdfjs.getDocument({
        data,
        isEvalSupported: false,
        fontExtraProperties: true,
        cMapUrl: baseUrl(PDFJS_DIR + "cmaps/"),
        cMapPacked: true,
        wasmUrl: baseUrl(PDFJS_DIR + "wasm/"),
        enableXfa: false,
        stopAtErrors: false,
        verbosity: 0,
      });
      let wrong = false;
      loadingTask.onPassword = async (update) => {
        setProgress(0.05, "");
        const pw = await askPassword(wrong);
        wrong = true;
        update(pw);
      };
      const pdfDoc = await race(loadingTask.promise);
      let title = "";
      try {
        const meta = await pdfDoc.getMetadata();
        title = meta?.info?.Title || "";
      } catch (_) { /* bez metadanych */ }
      const phases = { read: T("odczyt", "reading"), ocr: T("rozpoznawanie tekstu (OCR)", "text recognition (OCR)"), layout: T("układ", "layout"), images: T("obrazy", "images"), write: T("zapis .docx", "writing .docx") };
      const weights = { read: [0.05, 0.6], ocr: [0.05, 0.6], layout: [0.6, 0.8], images: [0.8, 0.95], write: [0.95, 0.99] };
      currentSignal = signal;
      const res = await race(window.DWPdf.convertPdf(pdfDoc, {
        pdfjs,
        JSZip: window.JSZip,
        encodeImage,
        renderVectors,
        glyphFixes,
        askOcr: (kind) => askOcr(options.photo && kind !== "curves" ? "photo" : kind),
        ocrPage,
        ocrParallel: window.__dwbOcrParallel || OCR_PARALLEL,
        signal,
        title,
        lang: "pl-PL",
        onProgress(i, n, phase, sub) {
          const [a, b] = weights[phase] || [0, 1];
          if (signal.cancelled) return;
          const frac = phase === "ocr" && typeof sub === "number" ? (i - 1 + sub) / Math.max(1, n) : i / Math.max(1, n);
          const pct = phase === "ocr" && typeof sub === "number" ? ` ${Math.round(sub * 100)}%` : "";
          setProgress(a + (b - a) * frac, `${T("Strona", "Page")} ${i} ${T("z", "of")} ${n} — ${phases[phase] || phase}${pct}`);
        },
      }));
      try {
        await loadingTask.destroy();
      } catch (_) { /* nic */ }
      if (signal.cancelled) throw new window.DWPdf.CancelledError();
      setProgress(1, T("Otwieranie dokumentu…", "Opening the document…"));
      const name = (file.name || "dokument.pdf").replace(/\.pdf$/i, "") + ".docx";
      const docx = new File([res.bytes], name, { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      closeDialog();
      const ok = await ingestFile(docx, { silent: true });
      if (ok) {
        // nowy plik, którego nie ma jeszcze na dysku — „Zapisz” zapyta o miejsce
        setDirtyState(true);
        const r = res.report;
        const parts = [`${r.pages} ${plural(r.pages, "strona", "strony", "stron", "page", "pages")}`];
        if (r.tables) parts.push(`${r.tables} ${plural(r.tables, "tabela", "tabele", "tabel", "table", "tables")}`);
        if (r.images) parts.push(`${r.images} ${plural(r.images, "obraz", "obrazy", "obrazów", "image", "images")}`);
        const secs = ((performance.now() - t0) / 1000).toFixed(1);
        toast(`${options.photo ? T(options.photo.pages > 1 ? `Zdjęcia (${options.photo.pages}) zamienione na dokument` : "Zdjęcie zamienione na dokument", "Photo(s) converted to a document") : T("Przekonwertowano PDF", "PDF converted")}: ${parts.join(", ")} (${secs} s). ${T("Zapisz jako .docx.", "Save it as .docx.")}`, "success");
        if (r.suspectChars) toast(T(`${r.suspectChars} ${plural(r.suspectChars, "znak mógł", "znaki mogły", "znaków mogło")} się źle odczytać z PDF — ${r.suspectChars === 1 ? "jest zaznaczony" : "są zaznaczone"} na żółto. Sprawdź je.`, `${r.suspectChars} character(s) may have been read incorrectly — highlighted in yellow.`), "warning");
        if (r.ocrPages) toast(T(`Rozpoznano druk na ${r.ocrPages} ${plural(r.ocrPages, "stronie", "stronach", "stronach")} — sprawdź go, OCR bywa omylny.`, `Printed text recognized on ${r.ocrPages} page(s) — please check it, OCR can make mistakes.`), "info");
        if (r.ocrDoubt) toast(T(`${r.ocrDoubt} ${plural(r.ocrDoubt, "słowo rozpoznane", "słowa rozpoznane", "słów rozpoznanych")} z mniejszą pewnością ${r.ocrDoubt === 1 ? "jest zaznaczone" : "są zaznaczone"} na żółto — sprawdź je z oryginałem.`, `${r.ocrDoubt} word(s) recognized with lower confidence are highlighted in yellow — check them against the original.`), "warning");
        if (r.ocrSkipped) toast(T(`${r.ocrSkipped} ${plural(r.ocrSkipped, "fragment", "fragmenty", "fragmentów")} (np. pismo odręczne) ${r.ocrSkipped === 1 ? "został" : "zostało"} jako obraz, bez zmian — tego jeszcze nie rozpoznajemy.`, `${r.ocrSkipped} fragment(s) (e.g. handwriting) kept unchanged as image — not recognized yet.`), "info");
        if (r.scannedPages && !r.ocrPages && !r.ocrLayerPages) {
          toast(T(`${r.scannedPages === r.pages ? "To skan" : `${r.scannedPages} ${plural(r.scannedPages, "strona to skan", "strony to skany", "stron to skany")}`} — wstawiony jako obraz (bez rozpoznawania tekstu).`, `${r.scannedPages} scanned page(s) — inserted as images (no text recognition).`), "info");
        }
        setStatus(`${T("PDF → DOCX", "PDF → DOCX")}: ${parts.join(", ")}`);
        window.dispatchEvent(new CustomEvent("dwb:pdf-converted", { detail: { report: r, name } }));
      }
      return ok;
    } catch (e) {
      closeDialog();
      if (e && (e.name === "CancelledError" || signal.cancelled)) {
        toast(T("Konwersję anulowano.", "Conversion cancelled."), "info");
        return false;
      }
      const msg = String(e?.message || e);
      if (e?.name === "InvalidPDFException" || /Invalid PDF/i.test(msg)) toast(T("To nie jest poprawny plik PDF (albo jest uszkodzony).", "This is not a valid PDF file (or it is damaged)."), "error");
      else if (e?.name === "PasswordException") toast(T("Nie udało się otworzyć PDF-a chronionego hasłem.", "Could not open the password-protected PDF."), "error");
      else toast(T("Nie udało się przekonwertować PDF-a.", "Could not convert the PDF.") + " " + msg.slice(0, 120), "error");
      if (typeof log === "function") log("PDF → DOCX: " + msg, "error");
      return false;
    } finally {
      cancelBtn.removeEventListener("click", onCancel);
      busy = false;
      currentSignal = null;
      d.querySelector("#pdfConvAsk")?.classList.add("hidden");
      // silnik OCR zajmuje sporo pamięci (telefon) — zwalniamy po konwersji
      if (window.DWPdfOcr) window.DWPdfOcr.terminate();
    }
  }

  function closeDialog() {
    if (dlg?.open) {
      try {
        dlg.close();
      } catch (_) {
        dlg.removeAttribute("open");
      }
    }
  }

  function plural(n, one, few, many, en1, enN) {
    if (typeof currentLang === "string" && currentLang === "en") return n === 1 ? en1 : enN || en1;
    if (n === 1) return one;
    const m10 = n % 10, m100 = n % 100;
    return m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14) ? few : many;
  }

  window.dwbPdfImport = { convertFile, loadEngine, _glyphFixes: glyphFixes, _encodeImage: encodeImage, _renderVectors: renderVectors };
})();
