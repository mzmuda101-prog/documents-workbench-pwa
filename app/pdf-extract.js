// PDF → DOCX, krok 1: odczyt strony PDF (pdf.js) do prostego modelu.
// Interpretujemy listę operacji pdf.js tak, jak robi to jego kanwa (showText, macierze, kolory),
// ale zamiast rysować zbieramy: znaki z pozycją/czcionką/kolorem, prostokąty i linie (tabele,
// tła komórek, podkreślenia), obrazy i adnotacje z tekstem. Współrzędne: punkty PDF, oś Y w dół
// (początek w lewym górnym rogu strony po obrocie). Działa w przeglądarce i w Node (testy).
(function (root) {
  "use strict";
  const NS = (root.DWPdf = root.DWPdf || {});

  const IDENTITY = [1, 0, 0, 1, 0, 0];
  const mul = (m1, m2) => [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ];
  const apply = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

  function bboxOf(points) {
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const [x, y] of points) {
      if (x < x0) x0 = x;
      if (y < y0) y0 = y;
      if (x > x1) x1 = x;
      if (y > y1) y1 = y;
    }
    return { x0, y0, x1, y1 };
  }
  const intersect = (a, b) => {
    if (!a) return b;
    if (!b) return a;
    return { x0: Math.max(a.x0, b.x0), y0: Math.max(a.y0, b.y0), x1: Math.min(a.x1, b.x1), y1: Math.min(a.y1, b.y1) };
  };

  function hexToRgb(hex) {
    if (typeof hex !== "string") return null;
    const m = /^#?([0-9a-f]{6})$/i.exec(hex);
    return m ? m[1].toLowerCase() : null;
  }

  // Nazwa czcionki → rodzina dla Worda + pogrubienie/kursywa z nazwy.
  // „ABCDEF+TimesNewRomanPS-BoldItalicMT” → { family: "Times New Roman", bold, italic }.
  const FAMILY_ALIASES = {
    helvetica: "Arial", arialmt: "Arial", arial: "Arial", "helveticaneue": "Helvetica Neue",
    times: "Times New Roman", timesnewroman: "Times New Roman", timesnewromanps: "Times New Roman", timesroman: "Times New Roman",
    courier: "Courier New", couriernew: "Courier New", couriernewps: "Courier New",
    symbol: "Symbol", zapfdingbats: "Wingdings", wingdings: "Wingdings",
    calibri: "Calibri", cambria: "Cambria", verdana: "Verdana", tahoma: "Tahoma", georgia: "Georgia",
    garamond: "Garamond", bookantiqua: "Book Antiqua", centurygothic: "Century Gothic", segoeui: "Segoe UI",
    trebuchetms: "Trebuchet MS", consolas: "Consolas", candara: "Candara", corbel: "Corbel", constantia: "Constantia",
    bahnschrift: "Bahnschrift", aptos: "Aptos", liberationsans: "Liberation Sans", liberationserif: "Liberation Serif",
    dejavusans: "DejaVu Sans", dejavuserif: "DejaVu Serif", opensans: "Open Sans", roboto: "Roboto", lato: "Lato",
  };
  const STYLE_WORDS = /(bold|black|heavy|semibold|demibold|extrabold|ultrabold|medium|light|thin|book|regular|roman|italic|oblique|inclined|condensed|narrow|cond|it|bd|bi|ps|mt|psmt|std|pro|lt|ms)$/i;

  function parseFontName(raw) {
    let name = String(raw || "").replace(/^[A-Z]{6}\+/, "").replace(/^\//, "");
    const lower = name.toLowerCase();
    const bold = /bold|black|heavy|semibold|demibold|extrabold|ultrabold|,bd|-bd\b/.test(lower);
    const italic = /italic|oblique|inclined|,it\b|-it\b/.test(lower);
    // Odetnij część stylu: po myślniku/przecinku, a potem znane końcówki doklejone bez separatora.
    let base = name.split(/[-,]/)[0] || name;
    const keyOf = (s) => s.toLowerCase().replace(/[^a-z0-9]/g, "");
    let family = FAMILY_ALIASES[keyOf(base)];
    let prev;
    while (!family) {
      prev = base;
      base = base.replace(STYLE_WORDS, "");
      if (base === prev || base.length <= 3) break;
      family = FAMILY_ALIASES[keyOf(base)];
    }
    if (!base) base = name;
    if (!family) {
      // „TimesNewRoman” → „Times New Roman”; „WtTimesBold01jcnDxdw” zostawiamy do heurystyki.
      family = base.replace(/([a-z])([A-Z])/g, "$1 $2").replace(/_/g, " ").trim();
    }
    return { family, bold, italic, raw: name };
  }

  // --- Mini-czytnik tabel TrueType/OpenType: prawdziwa nazwa rodziny i waga z pliku czcionki.
  // „Microsoft Print to PDF” nazywa czcionki „CIDFont+F1”, ale w środku jest tabela name.
  function readSfntInfo(bytes) {
    try {
      if (!bytes || bytes.length < 12) return null;
      const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      const tag = (o) => String.fromCharCode(bytes[o], bytes[o + 1], bytes[o + 2], bytes[o + 3]);
      const numTables = dv.getUint16(4);
      if (numTables > 100) return null;
      const tables = {};
      for (let i = 0; i < numTables; i++) {
        const o = 12 + i * 16;
        tables[tag(o)] = { off: dv.getUint32(o + 8), len: dv.getUint32(o + 12) };
      }
      const info = {};
      const nameT = tables.name;
      if (nameT && nameT.off + 6 <= bytes.length) {
        const base = nameT.off;
        const count = dv.getUint16(base + 2);
        const strOff = base + dv.getUint16(base + 4);
        const found = {};
        for (let i = 0; i < count && base + 6 + i * 12 + 12 <= bytes.length; i++) {
          const r = base + 6 + i * 12;
          const platform = dv.getUint16(r), enc = dv.getUint16(r + 2), lang = dv.getUint16(r + 4);
          const nameId = dv.getUint16(r + 6), len = dv.getUint16(r + 8), off = dv.getUint16(r + 10);
          if (![1, 2, 4, 16, 17].includes(nameId)) continue;
          const s = strOff + off;
          if (s + len > bytes.length) continue;
          let str = "";
          if (platform === 3 || platform === 0) {
            for (let k = 0; k + 1 < len; k += 2) str += String.fromCharCode(dv.getUint16(s + k));
          } else if (platform === 1 && enc === 0) {
            for (let k = 0; k < len; k++) str += String.fromCharCode(bytes[s + k]);
          } else continue;
          const prio = platform === 3 && (lang === 0x409 || lang === 0x415) ? 2 : 1;
          if (!found[nameId] || found[nameId].prio < prio) found[nameId] = { str, prio };
        }
        info.family = (found[16] || found[1])?.str || null;
        info.subfamily = (found[17] || found[2])?.str || null;
      }
      const os2 = tables["OS/2"];
      if (os2 && os2.off + 64 <= bytes.length) {
        info.weight = dv.getUint16(os2.off + 4);
        info.fsSelection = dv.getUint16(os2.off + 62);
        // panose[0] = rodzina (2 = tekst łaciński), panose[1] = szeryfy (11–13 = bez szeryfów).
        info.panoseSerif = bytes[os2.off + 32 + 1];
      }
      const post = tables.post;
      if (post && post.off + 16 <= bytes.length) info.isFixedPitch = dv.getUint32(post.off + 12) !== 0;
      return info;
    } catch (_) {
      return null;
    }
  }

  const BAD_FAMILY = /^(cidfont|font|f\d+|t\d+|tt\d+|unnamed|untitled|ps)/i;

  // Opis czcionki pdf.js (commonObjs) → nasz opis (rodzina, pogrubienie, kursywa, metryki).
  function describeFont(fontObj, cache) {
    const id = fontObj?.loadedName || fontObj?.name || "?";
    if (cache.has(id)) return cache.get(id);
    const parsed = parseFontName(fontObj?.name || "");
    let family = parsed.family;
    let bold = parsed.bold || !!fontObj?.bold || !!fontObj?.black;
    let italic = parsed.italic || !!fontObj?.italic;
    let sfnt = null;
    const data = fontObj?.data;
    if (data && data.length) {
      const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
      const t = String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]);
      if (t === "\0\x01\0\0" || t === "OTTO" || t === "true") sfnt = readSfntInfo(bytes);
    }
    if (sfnt?.family && (!family || BAD_FAMILY.test(family) || /^cid/i.test(parsed.raw) || family.length < 3)) {
      family = parseFontName(sfnt.family).family || sfnt.family;
    }
    if (sfnt) {
      if (sfnt.weight >= 600 || /bold|black|heavy/i.test(sfnt.subfamily || "")) bold = true;
      if ((sfnt.fsSelection & 1) || /italic|oblique/i.test(sfnt.subfamily || "")) italic = true;
    }
    const mono = !!fontObj?.isMonospace || !!sfnt?.isFixedPitch;
    let serif = !!fontObj?.isSerifFont;
    if (sfnt?.panoseSerif >= 11 && sfnt?.panoseSerif <= 13) serif = false;
    if (!family || BAD_FAMILY.test(family)) family = mono ? "Courier New" : serif ? "Times New Roman" : "Arial";
    // Czcionki bez szeryfów/szeryfowe z niestandardowych zestawów (np. „WtTimesBold01jcnDxdw” z
    // Ghostscripta) — rozpoznajemy znane rdzenie w nazwie.
    if (!FAMILY_ALIASES[family.toLowerCase().replace(/[^a-z]/g, "")]) {
      const l = family.toLowerCase();
      if (/times/.test(l) && !/new/.test(l)) family = "Times New Roman";
      else if (/helvetica|arial/.test(l)) family = "Arial";
      else if (/courier/.test(l)) family = "Courier New";
    }
    const symbolic = /symbol|wingdings|dingbats|webdings/i.test(family);
    const asc = Number.isFinite(fontObj?.ascent) && fontObj.ascent > 0.3 && fontObj.ascent < 1.6 ? fontObj.ascent : 0.86;
    let desc = Number.isFinite(fontObj?.descent) ? Math.abs(fontObj.descent) : 0.22;
    if (desc > 0.6 || desc < 0.05) desc = 0.22;
    const out = { id, family, bold, italic, mono, serif, symbolic, asc, desc, raw: parsed.raw, type3: !!fontObj?.isType3Font };
    cache.set(id, out);
    return out;
  }

  // Znaki diakrytyczne rysowane osobno (Ghostscript, stare sterowniki drukarek): ´ nad „s” → „ś”.
  const SPACING_TO_COMBINING = {
    "´": "́", "ˊ": "́", "`": "̀", "ˋ": "̀",
    "˙": "̇", "¨": "̈", "ˇ": "̌", "˘": "̆",
    "˛": "̨", "¸": "̧", "˚": "̊", "°": "̊",
    "˜": "̃", "^": "̂", "ˆ": "̂", "˝": "̋", "¯": "̄",
  };
  NS.SPACING_TO_COMBINING = SPACING_TO_COMBINING;

  /**
   * Odczyt jednej strony.
   * @param {object} page   PDFPageProxy
   * @param {object} pdfjs  moduł pdf.js (OPS)
   * @param {object} [opts] { fontCache: Map }
   */
  async function extractPage(page, pdfjs, opts = {}) {
    const OPS = pdfjs.OPS;
    const viewport = page.getViewport({ scale: 1 });
    const fontCache = opts.fontCache || new Map();
    // Adnotacje najpierw: wyglądy pieczątek/obrazków rysujemy z listy operacji (ENABLE_FORMS),
    // a tekst pól formularza i dymków bierzemy z ich wartości (pełne Unicode — wygląd pola bywa
    // narysowany czcionką bez „Ł”, „Ś”).
    let rawAnnotations = [];
    try {
      rawAnnotations = await page.getAnnotations({ intent: "display" });
    } catch (_) { rawAnnotations = []; }
    // Linki/dymki: nic do rysowania. Pola i pola tekstowe: wygląd czytamy tylko po to, by poznać
    // rozmiar i położenie wierszy tekstu (wartość bierzemy z pola).
    const skipAnnotIds = new Set(rawAnnotations.filter((a) => a.subtype === "Link" || a.subtype === "Popup").map((a) => a.id));
    const textAnnotIds = new Set(rawAnnotations.filter((a) => a.subtype === "FreeText" || a.subtype === "Widget").map((a) => a.id));
    const annotGlyphs = {};
    let annotTextId = null;
    const opList = await page.getOperatorList({ annotationMode: pdfjs.AnnotationMode?.ENABLE ?? 1 });
    const { fnArray, argsArray } = opList;

    const glyphs = [];
    const rects = []; // wypełnione prostokąty (tła, grube linie)
    const lines = []; // odcinki rysowane obrysem
    const graphics = []; // inne kształty (krzywe) — tylko obwiednia
    const images = [];

    let gs = {
      ctm: viewport.transform.slice(),
      fill: "000000", stroke: "000000", fillAlpha: 1, strokeAlpha: 1, lineWidth: 1,
      clip: { x0: 0, y0: 0, x1: viewport.width, y1: viewport.height },
      font: null, fontSize: 0, charSpacing: 0, wordSpacing: 0, hscale: 1, leading: 0, rise: 0, renderMode: 0,
      fillPattern: false,
    };
    const stack = [];
    let tm = IDENTITY.slice();
    let tx = 0, ty = 0, lineX = 0, lineY = 0;
    let pendingClip = false;
    const mcStack = []; // znaczniki treści: Artifact (nagłówki/stopki w PDF z Worda)
    let artifactDepth = 0;
    let skipDepth = 0; // >0 = wnętrze adnotacji, której treść bierzemy z wartości (nie z wyglądu)
    let order = 0;

    const fontOf = (ref) => {
      try {
        return page.commonObjs.get(ref);
      } catch (_) {
        return null;
      }
    };

    const pushMark = (isArtifact) => {
      mcStack.push(isArtifact);
      if (isArtifact) artifactDepth++;
    };

    for (let i = 0; i < fnArray.length; i++) {
      const fn = fnArray[i];
      const args = argsArray[i];
      if (fn === OPS.beginAnnotation) {
        stack.push({ ...gs, ctm: gs.ctm.slice() });
        const [id, rect, transform, matrix] = args;
        if (skipDepth || skipAnnotIds.has(id)) {
          skipDepth++;
          continue;
        }
        if (textAnnotIds.has(id)) {
          annotTextId = id;
          annotGlyphs[id] = annotGlyphs[id] || [];
        }
        gs = { ...gs, fill: "000000", stroke: "000000", fillAlpha: 1, strokeAlpha: 1, fillPattern: false, patternImage: null };
        gs.ctm = mul(mul(viewport.transform, Array.from(transform || IDENTITY)), Array.from(matrix || IDENTITY));
        if (rect) {
          const vt = viewport.transform;
          gs.clip = bboxOf([apply(vt, rect[0], rect[1]), apply(vt, rect[2], rect[3])]);
        }
        continue;
      }
      if (fn === OPS.endAnnotation) {
        if (skipDepth) skipDepth--;
        else annotTextId = null;
        if (stack.length) gs = stack.pop();
        continue;
      }
      if (skipDepth) continue;
      switch (fn) {
        case OPS.save:
          stack.push({ ...gs, ctm: gs.ctm.slice() });
          break;
        case OPS.restore:
          if (stack.length) gs = stack.pop();
          break;
        case OPS.transform:
          gs.ctm = mul(gs.ctm, args);
          break;
        case OPS.paintFormXObjectBegin: {
          stack.push({ ...gs, ctm: gs.ctm.slice() });
          const [matrix, bbox] = args;
          if (Array.isArray(matrix) || ArrayBuffer.isView(matrix)) gs.ctm = mul(gs.ctm, Array.from(matrix));
          if (bbox && bbox.length === 4) {
            const b = bboxOf([apply(gs.ctm, bbox[0], bbox[1]), apply(gs.ctm, bbox[2], bbox[3]), apply(gs.ctm, bbox[0], bbox[3]), apply(gs.ctm, bbox[2], bbox[1])]);
            gs.clip = intersect(gs.clip, b);
          }
          break;
        }
        case OPS.paintFormXObjectEnd:
          if (stack.length) gs = stack.pop();
          break;
        case OPS.setLineWidth:
          gs.lineWidth = args[0];
          break;
        case OPS.setFillRGBColor:
          gs.fill = hexToRgb(args[0]) || gs.fill;
          gs.fillPattern = false;
          break;
        case OPS.setStrokeRGBColor:
          gs.stroke = hexToRgb(args[0]) || gs.stroke;
          break;
        case OPS.setFillColorN:
          gs.patternImage = null;
          if (args && args[0] === "TilingPattern" || args?.[0] === "Shading" || typeof args?.[0] === "string" && !/^#/.test(args[0])) gs.fillPattern = true;
          // Wzór z jednym obrazem (skaner/ONLYOFFICE wkleja tak cały skan strony) = zwykły obraz.
          if (args?.[0] === "TilingPattern" && args[2]?.fnArray) {
            const pOps = args[2];
            let pm = mul(viewport.transform, Array.from(args[3] || IDENTITY));
            for (let k = 0; k < pOps.fnArray.length; k++) {
              const f2 = pOps.fnArray[k];
              if (f2 === OPS.transform) pm = mul(pm, Array.from(pOps.argsArray[k]));
              else if (f2 === OPS.paintImageXObject) {
                gs.patternImage = { objId: pOps.argsArray[k][0], matrix: pm };
                break;
              }
            }
          }
          break;
        case OPS.setFillTransparent:
          gs.fillAlpha = 0;
          break;
        case OPS.setStrokeTransparent:
          gs.strokeAlpha = 0;
          break;
        case OPS.setGState:
          for (const [k, v] of args[0] || []) {
            if (k === "ca") gs.fillAlpha = v;
            else if (k === "CA") gs.strokeAlpha = v;
            else if (k === "LW") gs.lineWidth = v;
            else if (k === "Font" && Array.isArray(v)) {
              gs.font = fontOf(v[0]);
              gs.fontSize = v[1];
            }
          }
          break;
        case OPS.beginMarkedContent:
          pushMark(args?.[0]?.name === "Artifact" || args?.[0] === "Artifact");
          break;
        case OPS.beginMarkedContentProps:
          pushMark(args?.[0] === "Artifact" || args?.[0]?.name === "Artifact");
          break;
        case OPS.endMarkedContent:
          if (mcStack.length && mcStack.pop()) artifactDepth--;
          break;
        // ---------- tekst
        case OPS.beginText:
          tm = IDENTITY.slice();
          tx = ty = lineX = lineY = 0;
          break;
        case OPS.setFont: {
          const f = fontOf(args[0]);
          gs.font = f;
          gs.fontSize = args[1];
          break;
        }
        case OPS.setCharSpacing:
          gs.charSpacing = args[0];
          break;
        case OPS.setWordSpacing:
          gs.wordSpacing = args[0];
          break;
        case OPS.setHScale:
          gs.hscale = args[0] / 100;
          break;
        case OPS.setLeading:
          gs.leading = -args[0];
          break;
        case OPS.setTextRise:
          gs.rise = args[0];
          break;
        case OPS.setTextRenderingMode:
          gs.renderMode = args[0];
          break;
        case OPS.setTextMatrix:
          tm = Array.from(args[0] && args[0].length === 6 ? args[0] : args);
          tx = ty = lineX = lineY = 0;
          break;
        case OPS.moveText:
          tx = lineX += args[0];
          ty = lineY += args[1];
          break;
        case OPS.setLeadingMoveText:
          gs.leading = args[1];
          tx = lineX += args[0];
          ty = lineY += args[1];
          break;
        case OPS.nextLine:
          tx = lineX += 0;
          ty = lineY += gs.leading;
          break;
        case OPS.showText:
        case OPS.showSpacedText: {
          const font = gs.font;
          if (!font) break;
          const fontSize = gs.fontSize;
          if (!fontSize) break;
          const desc = describeFont(font, fontCache);
          const fontMatrix = font.fontMatrix || [0.001, 0, 0, 0.001, 0, 0];
          const dir = fontSize < 0 ? -1 : 1;
          const size = Math.abs(fontSize);
          const widthAdvanceScale = size * fontMatrix[0];
          const hs = gs.hscale * dir;
          const vertical = !!font.vertical;
          const spacingDir = vertical ? 1 : -1;
          // Macierz tekst → strona: CTM × Tm × przesunięcie (tx, ty+rise).
          const base = mul(gs.ctm, tm);
          const emH = size * (fontMatrix[3] ? Math.abs(fontMatrix[3]) * 1000 : 1);
          // Rozmiar na stronie = długość wektora osi Y tekstu.
          const yScale = Math.hypot(base[2], base[3]);
          const xScale = Math.hypot(base[0], base[1]);
          const pageSize = emH * yScale;
          const angle = Math.atan2(base[1], base[0]);
          const invisible = gs.renderMode === 3 || gs.renderMode === 7 || gs.fillAlpha === 0 && (gs.renderMode === 0);
          const fakeBold = gs.renderMode === 2;
          let x = 0;
          const list = args[0];
          for (let k = 0; k < list.length; k++) {
            const g = list[k];
            if (typeof g === "number") {
              x += spacingDir * g * size / 1000;
              continue;
            }
            const spacing = (g.isSpace ? gs.wordSpacing : 0) + gs.charSpacing;
            const adv = g.width * widthAdvanceScale;
            const gx = tx + x * gs.hscale;
            const charW = vertical ? adv - spacing * dir : adv + spacing * dir;
            const p0 = apply(base, gx, ty + gs.rise);
            const p1 = apply(base, gx + adv * hs, ty + gs.rise);
            const u = g.unicode || "";
            if (u && !(invisible && opts.skipInvisible)) {
              (annotTextId ? annotGlyphs[annotTextId] : glyphs).push({
                u,
                x: Math.min(p0[0], p1[0]),
                x1: Math.max(p0[0], p1[0]),
                y: p0[1], // linia bazowa
                adv: Math.hypot(p1[0] - p0[0], p1[1] - p0[1]),
                size: pageSize,
                asc: desc.asc * pageSize,
                desc: desc.desc * pageSize,
                angle,
                font: desc,
                fontRef: font.loadedName,
                fontChar: g.fontChar,
                code: g.originalCharCode,
                // Kod 32 to w PDF „spacja” dla odstępu między słowami, ale czcionka z własnym
                // kodowaniem może mieć tam literę (karty: „f” pod kodem 32) — liczy się znak.
                isSpace: /^\s+$/.test(u),
                color: gs.fillPattern ? "000000" : gs.fill,
                bold: desc.bold || fakeBold,
                italic: desc.italic || Math.abs(base[2] / (base[3] || 1)) > 0.12 && Math.abs(angle) < 0.01,
                invisible,
                artifact: artifactDepth > 0,
                clip: gs.clip,
                order: order++,
                xScale: xScale / (yScale || 1),
              });
            }
            x += vertical ? -charW : charW;
          }
          if (vertical) ty -= x;
          else tx += x * gs.hscale;
          break;
        }
        // ---------- ścieżki
        case OPS.clip:
        case OPS.eoClip:
          pendingClip = true;
          break;
        case OPS.constructPath: {
          if (annotTextId) break;
          const [paintOp, dataArr, minMax] = args;
          const data = Array.isArray(dataArr) ? dataArr[0] : dataArr;
          const subpaths = parsePath(data, gs.ctm);
          if (pendingClip) {
            pendingClip = false;
            if (minMax && minMax.length === 4 && isFinite(minMax[0])) {
              const b = bboxOf([apply(gs.ctm, minMax[0], minMax[1]), apply(gs.ctm, minMax[2], minMax[3]), apply(gs.ctm, minMax[0], minMax[3]), apply(gs.ctm, minMax[2], minMax[1])]);
              gs.clip = intersect(gs.clip, b);
            }
          }
          const isFill = paintOp === OPS.fill || paintOp === OPS.eoFill || paintOp === OPS.fillStroke || paintOp === OPS.eoFillStroke || paintOp === OPS.closeFillStroke || paintOp === OPS.closeEOFillStroke;
          const isStroke = paintOp === OPS.stroke || paintOp === OPS.closeStroke || paintOp === OPS.fillStroke || paintOp === OPS.eoFillStroke || paintOp === OPS.closeFillStroke || paintOp === OPS.closeEOFillStroke;
          if (!isFill && !isStroke) break;
          if (isFill && gs.patternImage && minMax && minMax.length === 4) {
            const pathBox = bboxOf([apply(gs.ctm, minMax[0], minMax[1]), apply(gs.ctm, minMax[2], minMax[3]), apply(gs.ctm, minMax[0], minMax[3]), apply(gs.ctm, minMax[2], minMax[1])]);
            const m = gs.patternImage.matrix;
            const b = bboxOf([apply(m, 0, 0), apply(m, 1, 0), apply(m, 0, 1), apply(m, 1, 1)]);
            const clipped = clipBox(b, intersect(gs.clip, pathBox));
            if (clipped && b.x1 - b.x0 > 1 && b.y1 - b.y0 > 1) {
              images.push({ ...b, clipped, order: order++, matrix: m.slice(), kind: "image", objId: gs.patternImage.objId, artifact: artifactDepth > 0 });
            }
            break;
          }
          const lw = Math.max(0.1, gs.lineWidth * Math.hypot(gs.ctm[0], gs.ctm[1]));
          const evenOdd = paintOp === OPS.eoFill || paintOp === OPS.eoFillStroke || paintOp === OPS.closeEOFillStroke;
          // Ścieżka złożona z krzywymi (litery zamienione na krzywe: „o” = obrys + dziura, „i” =
          // kreska + kropka) — wypełniamy RAZEM. Każdy kontur osobno zalewał dziury w a/o/e,
          // a prostokącik kropki/przecinka trafiał do teł i linii (i znikał).
          let fillDone = false;
          if (isFill && gs.fillAlpha > 0.05 && !gs.fillPattern && subpaths.length > 1 && subpaths.some((sp) => !sp.rect)) {
            const bb = bboxOf(subpaths.flatMap((sp) => [[sp.bbox.x0, sp.bbox.y0], [sp.bbox.x1, sp.bbox.y1]]));
            const clipped = clipBox(bb, gs.clip);
            if (clipped) graphics.push({ ...clipped, color: gs.fill, alpha: gs.fillAlpha, kind: "fill", cmds: subpaths.flatMap((sp) => sp.cmds), evenOdd, compound: true, clip: gs.clip, order: order++ });
            fillDone = true;
          }
          for (const sp of subpaths) {
            const clipped = clipBox(sp.bbox, gs.clip);
            if (!clipped) continue;
            if (isFill && !fillDone && gs.fillAlpha > 0.05 && !gs.fillPattern) {
              if (sp.rect) rects.push({ ...clipped, color: gs.fill, alpha: gs.fillAlpha, order: order++, artifact: artifactDepth > 0 });
              else graphics.push({ ...clipped, color: gs.fill, alpha: gs.fillAlpha, kind: "fill", cmds: sp.cmds, evenOdd, clip: gs.clip, order: order++ });
            }
            if (isStroke && gs.strokeAlpha > 0.05) {
              if (sp.rect) {
                const { x0, y0, x1, y1 } = sp.bbox;
                for (const seg of [[x0, y0, x1, y0], [x0, y1, x1, y1], [x0, y0, x0, y1], [x1, y0, x1, y1]]) lines.push(segLine(seg, lw, gs.stroke, order++));
              } else if (sp.segments) {
                for (const seg of sp.segments) lines.push(segLine(seg, lw, gs.stroke, order++));
              } else {
                graphics.push({ ...clipped, color: gs.stroke, alpha: gs.strokeAlpha, kind: "stroke", width: lw, cmds: sp.cmds, clip: gs.clip, order: order++ });
              }
            }
          }
          break;
        }
        // ---------- obrazy
        case OPS.paintImageXObject:
        case OPS.paintImageXObjectRepeat:
        case OPS.paintInlineImageXObject:
        case OPS.paintImageMaskXObject:
        case OPS.paintSolidColorImageMask: {
          if (annotTextId) break;
          const m = gs.ctm;
          const corners = [apply(m, 0, 0), apply(m, 1, 0), apply(m, 0, 1), apply(m, 1, 1)];
          const b = bboxOf(corners);
          const clipped = clipBox(b, gs.clip);
          if (!clipped) break;
          const w = b.x1 - b.x0, h = b.y1 - b.y0;
          if (w < 1 || h < 1) break;
          const img = {
            ...b, clipped, order: order++,
            matrix: m.slice(),
            kind: fn === OPS.paintImageMaskXObject || fn === OPS.paintSolidColorImageMask ? "mask" : "image",
            color: gs.fill,
            artifact: artifactDepth > 0,
          };
          if (fn === OPS.paintInlineImageXObject) img.inline = args[0];
          else if (fn === OPS.paintImageMaskXObject) img.inline = args[0];
          else img.objId = args[0];
          // Maski 1-bitowe w kształcie cienkiej linii = linie tabel w niektórych PDF-ach.
          if (img.kind !== "image" && (h < 2.5 || w < 2.5)) {
            rects.push({ ...clipped, color: gs.fill, alpha: 1, order: img.order });
            break;
          }
          images.push(img);
          break;
        }
        default:
          break;
      }
    }

    // Adnotacje z tekstem (FreeText — „pole tekstowe” wpisane w czytniku PDF) i linki.
    const annotations = [];
    try {
      const anns = rawAnnotations;
      for (const a of anns) {
        const r = a.rect;
        if (!r) continue;
        const [x0, y0] = viewport.convertToViewportPoint(r[0], r[3]);
        const [x1, y1] = viewport.convertToViewportPoint(r[2], r[1]);
        const box = { x0: Math.min(x0, x1), y0: Math.min(y0, y1), x1: Math.max(x0, x1), y1: Math.max(y0, y1) };
        if (a.subtype === "Link" && (a.url || a.unsafeUrl || a.dest)) {
          annotations.push({ type: "link", ...box, url: a.url || a.unsafeUrl || null, dest: a.dest || null });
        } else if (a.subtype === "FreeText") {
          const text = a.contentsObj?.str || a.richText?.str || "";
          if (!text.trim()) continue;
          const da = a.defaultAppearanceData || {};
          const color = da.fontColor ? Array.from(da.fontColor).map((v) => Math.round(v).toString(16).padStart(2, "0")).join("") : "000000";
          annotations.push({ type: "text", id: a.id, ...box, text, size: da.fontSize || 0, fontName: da.fontName || "Helv", color, hidden: !!a.hidden || (a.annotationFlags & 2) !== 0, glyphs: annotGlyphs[a.id] || [], align: a.textAlignment || 0 });
        } else if (a.subtype === "Widget" && a.fieldType === "Tx" && a.fieldValue) {
          annotations.push({ type: "text", id: a.id, ...box, text: String(a.fieldValue), size: a.defaultAppearanceData?.fontSize || 0, fontName: "Helv", color: "000000", field: true, glyphs: annotGlyphs[a.id] || [], align: a.textAlignment || 0, multiLine: !!a.multiLine, comb: !!a.comb });
        } else if (a.subtype === "Widget" && a.fieldType === "Btn" && a.checkBox) {
          annotations.push({ type: "check", ...box, checked: a.fieldValue && a.fieldValue !== "Off" });
        }
      }
    } catch (_) { /* adnotacje są dodatkiem */ }

    return {
      width: viewport.width,
      height: viewport.height,
      glyphs, rects, lines, graphics, images, annotations,
      opCount: fnArray.length,
    };
  }

  function segLine([x0, y0, x1, y1], width, color, order) {
    return { x0: Math.min(x0, x1), y0: Math.min(y0, y1), x1: Math.max(x0, x1), y1: Math.max(y0, y1), width, color, order };
  }

  function clipBox(b, clip) {
    if (!clip) return b;
    const r = intersect(b, clip);
    // Linie mają zerową grubość — pozwalamy na x0 == x1 / y0 == y1.
    if (r.x1 < r.x0 - 0.5 || r.y1 < r.y0 - 0.5) return null;
    return { x0: r.x0, y0: r.y0, x1: Math.max(r.x0, r.x1), y1: Math.max(r.y0, r.y1) };
  }

  // Dane ścieżki pdf.js (DrawOPS: 0 moveTo, 1 lineTo, 2 curveTo, 3 quadraticCurveTo, 4 closePath)
  // → podścieżki z obwiednią; prostokąty wyrównane do osi i pojedyncze odcinki rozpoznajemy.
  function parsePath(data, ctm) {
    const out = [];
    if (!data || !data.length) return out;
    let cur = null;
    const flush = () => {
      if (!cur || cur.pts.length < 2) {
        cur = null;
        return;
      }
      const pts = cur.pts;
      const bbox = bboxOf(pts);
      const sp = { bbox, cmds: cur.cmds, closed: cur.closed };
      if (!cur.curved) {
        // Prostokąt: 4–5 punktów, boki poziome/pionowe.
        const uniq = pts.filter((p, idx) => idx === 0 || Math.abs(p[0] - pts[idx - 1][0]) > 0.01 || Math.abs(p[1] - pts[idx - 1][1]) > 0.01);
        if (uniq.length > 1 && Math.abs(uniq[0][0] - uniq[uniq.length - 1][0]) < 0.01 && Math.abs(uniq[0][1] - uniq[uniq.length - 1][1]) < 0.01) uniq.pop();
        const axis = uniq.every((p, idx) => {
          const q = uniq[(idx + 1) % uniq.length];
          return Math.abs(p[0] - q[0]) < 0.05 || Math.abs(p[1] - q[1]) < 0.05;
        });
        if (uniq.length === 4 && axis) sp.rect = true;
        else if (uniq.length === 2 || (axis && !cur.closed)) {
          sp.segments = [];
          for (let k = 0; k + 1 < uniq.length; k++) sp.segments.push([uniq[k][0], uniq[k][1], uniq[k + 1][0], uniq[k + 1][1]]);
          if (cur.closed && uniq.length > 2) sp.segments.push([uniq[uniq.length - 1][0], uniq[uniq.length - 1][1], uniq[0][0], uniq[0][1]]);
        } else if (axis && cur.closed) {
          sp.segments = [];
          for (let k = 0; k < uniq.length; k++) {
            const q = uniq[(k + 1) % uniq.length];
            sp.segments.push([uniq[k][0], uniq[k][1], q[0], q[1]]);
          }
          sp.polygon = true;
        }
      }
      out.push(sp);
      cur = null;
    };
    let i = 0;
    while (i < data.length) {
      const op = data[i++];
      const fresh = () => ({ pts: [], cmds: [], curved: false, closed: false });
      if (op === 0) {
        flush();
        cur = fresh();
        const p = apply(ctm, data[i], data[i + 1]);
        cur.pts.push(p);
        cur.cmds.push(["M", p]);
        i += 2;
      } else if (op === 1) {
        if (!cur) cur = fresh();
        const p = apply(ctm, data[i], data[i + 1]);
        cur.pts.push(p);
        cur.cmds.push(["L", p]);
        i += 2;
      } else if (op === 2) {
        if (!cur) cur = fresh();
        const a = apply(ctm, data[i], data[i + 1]), b = apply(ctm, data[i + 2], data[i + 3]), c = apply(ctm, data[i + 4], data[i + 5]);
        cur.pts.push(a, b, c);
        cur.cmds.push(["C", a, b, c]);
        cur.curved = true;
        i += 6;
      } else if (op === 3) {
        if (!cur) cur = fresh();
        const a = apply(ctm, data[i], data[i + 1]), b = apply(ctm, data[i + 2], data[i + 3]);
        cur.pts.push(a, b);
        cur.cmds.push(["Q", a, b]);
        cur.curved = true;
        i += 4;
      } else if (op === 4) {
        if (cur) {
          cur.closed = true;
          cur.cmds.push(["Z"]);
          if (cur.pts.length) cur.pts.push(cur.pts[0]);
        }
      } else {
        break; // nieznany kod — przerywamy bezpiecznie
      }
    }
    flush();
    return out;
  }

  NS.extractPage = extractPage;
  NS.parseFontName = parseFontName;
  NS.readSfntInfo = readSfntInfo;
  NS._mul = mul;
  NS._apply = apply;
})(typeof globalThis !== "undefined" ? globalThis : window);
