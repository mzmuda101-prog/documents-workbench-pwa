// PDF → DOCX, krok 2: z surowych znaków/prostokątów/obrazów strony robimy strukturę dokumentu:
// wiersze → akapity (wyrównanie, wcięcia, odstępy, tabulatory, kropki spisu treści), tabele
// z linii i teł komórek, układ kolumn (np. 4 karty na stronie) przez cięcie XY, obrazy w tekście
// albo „pływające” (gdy w PDF nachodzą na tekst). Wszystko w punktach, oś Y w dół.
(function (root) {
  "use strict";
  const NS = (root.DWPdf = root.DWPdf || {});

  const median = (arr) => {
    if (!arr.length) return 0;
    const s = arr.slice().sort((a, b) => a - b);
    return s[s.length >> 1];
  };
  const overlap1d = (a0, a1, b0, b1) => Math.max(0, Math.min(a1, b1) - Math.max(a0, b0));

  // ---------------------------------------------------------------- znaki
  // Znaki z czcionek symboli (Word: punktory Symbol/Wingdings) trafiają do Unicode prywatnego
  // (U+F0xx). Zamieniamy je na zwykłe odpowiedniki — działają w każdej czcionce.
  const SYMBOL_PUA = {
    0xf0b7: "•", 0xf0a7: "▪", 0xf06e: "■", 0xf0fc: "✓", 0xf0fb: "✗",
    0xf0d8: "➢", 0xf0e8: "➔", 0xf076: "❖", 0xf0a8: "◻", 0xf06f: "□",
    0xf071: "❑", 0xf075: "◆", 0xf0b2: "◊", 0xf0e0: "→", 0xf0de: "⇒",
    0xf02d: "−", 0xf0d7: "⋅", 0xf06c: "●", 0xf09f: "•", 0xf0a1: "○",
    0xf0b0: "°", 0xf0b4: "×", 0xf0b8: "÷", 0xf0a3: "≤", 0xf0b3: "≥",
    0xf0b9: "≠", 0xf0bb: "≈", 0xf0ae: "→", 0xf0ac: "←", 0xf0ad: "↑",
    0xf0af: "↓", 0xf0e5: "∑", 0xf0d6: "√", 0xf0a5: "∞", 0xf0b1: "±",
  };
  const SYMBOL_ASCII = { "·": "•" };

  function mapSymbolChar(g) {
    const cp = g.u.codePointAt(0);
    if (cp >= 0xf000 && cp <= 0xf0ff) {
      if (SYMBOL_PUA[cp]) return SYMBOL_PUA[cp];
      if (g.font.symbolic) {
        // Czcionka Symbol: litery greckie pod kodami ASCII.
        const c = cp - 0xf000;
        if (/symbol/i.test(g.font.family) && c >= 0x41 && c <= 0x7a) return symbolGreek(c);
        return "•";
      }
      return String.fromCharCode(cp - 0xf000);
    }
    if (g.font.symbolic && /symbol/i.test(g.font.family) && SYMBOL_ASCII[g.u]) return SYMBOL_ASCII[g.u];
    return null;
  }
  function symbolGreek(c) {
    const map = "ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ[∴]⊥_‾αβχδεφγηιϕκλμνοπθρστυϖωξψζ";
    const i = c - 0x41;
    return i >= 0 && i < map.length ? map[i] : "•";
  }

  const COMBINING = NS.SPACING_TO_COMBINING || {};

  /**
   * Czyszczenie znaków strony: niewidoczne/obcięte, symbole, doklejanie osobno rysowanych
   * akcentów (´ nad „s” → „ś”), poprawki z rozpoznania kształtów (opts.glyphFixes).
   */
  // Szerokości znaków (1/1000 em) Arial / Times New Roman, zwykły i pogrubiony — zmierzone
  // z prawdziwych czcionek (Arial = metryka Helvetiki). Do zagęszczenia tekstu, gdy krój
  // z PDF zastępujemy Arialem/Timesem (niezależnie od czcionek na urządzeniu konwersji).
  const SUB_CHARS = " !\"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~ąćęłńóśźżĄĆĘŁŃÓŚŹŻ„”“‚‘’–—…•°§€×ü";
  const SUB_WIDTHS = {
    arial: [278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,334,260,334,584,556,500,556,222,556,556,500,500,500,667,722,667,556,722,778,667,611,611,333,333,333,222,222,222,556,1000,1000,350,400,556,556,584,556],
    arialB: [278,333,474,556,556,889,722,238,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,333,333,584,584,584,611,975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,667,778,722,667,611,722,667,944,667,667,611,333,278,333,584,556,333,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,611,611,389,556,333,611,556,778,556,556,500,389,280,389,584,556,556,556,278,611,611,556,500,500,722,722,667,611,722,778,667,611,611,500,500,500,278,278,278,556,1000,1000,350,400,556,556,584,611],
    times: [250,333,408,500,500,833,778,180,333,333,500,564,250,333,250,278,500,500,500,500,500,500,500,500,500,500,278,278,564,564,564,444,921,722,667,667,722,611,556,722,722,333,389,722,611,889,722,722,556,722,667,556,611,722,722,944,722,722,611,333,278,333,469,500,333,444,500,444,500,444,333,500,500,278,278,500,278,778,500,500,500,500,333,389,278,500,500,722,500,500,444,480,200,480,541,444,444,444,278,500,500,389,444,444,722,667,611,611,722,722,556,611,611,444,444,444,333,333,333,500,1000,1000,350,400,500,500,564,500],
    timesB: [250,333,555,500,500,1000,833,278,333,333,500,570,250,333,250,278,500,500,500,500,500,500,500,500,500,500,333,333,570,570,570,500,930,722,667,722,722,667,611,778,778,389,500,778,667,944,722,778,611,778,722,556,667,722,722,1000,722,722,667,333,278,333,581,500,333,500,556,444,556,444,333,500,556,278,333,556,278,833,556,500,556,556,444,389,333,556,500,722,500,500,444,394,220,394,520,500,444,444,278,556,500,389,444,444,722,722,667,667,722,778,556,667,667,500,500,500,333,333,333,500,1000,1000,350,400,500,500,570,556],
  };
  const SUB_INDEX = new Map([...SUB_CHARS].map((ch, i) => [ch, i]));

  // Krój z PDF zastąpiony Arialem/Timesem (pdf-extract describeFont → substituted): litery
  // w zastępniku są szersze/węższe niż w PDF — zawijałyby się inaczej (podtytuł formularza
  // DC-85 schodził na 2 wiersze i wszystko poniżej zjeżdżało). Liczymy dla każdego kroju
  // z PDF, o ile em różni się średnio litera, i zapisujemy to jako odstęp między znakami
  // (Word: „zagęszczone/rozstrzelone”, podgląd: letter-spacing).
  function fitSubstitutedFonts(glyphs) {
    const acc = new Map();
    // Rzeczywisty krok = odległość do NASTĘPNEJ litery w tym samym wierszu (obejmuje odstęp Tc
    // i korekty w tablicy TJ — Ghostscript dosuwa nimi każdą literę). Pary ze spacją pomijamy
    // (odstęp między słowami to nie szerokość litery).
    for (let k = 0; k + 1 < glyphs.length; k++) {
      const g = glyphs[k], nx = glyphs[k + 1];
      if (!g.font?.substituted || g.isSpace || g.symbol || g.ocrLayer || !g.size) continue;
      if (nx.isSpace || nx.fontRef !== g.fontRef || Math.abs(nx.y - g.y) > 0.2 || Math.abs(g.angle || 0) > 0.01) continue;
      const step = nx.x - g.x;
      // odstęp między słowami bez znaku spacji (samo przesunięcie — Ghostscript) to nie litera
      if (step <= 0 || step > g.size * 1.6 || step > (g.adv || 0) + g.size * 0.15) continue;
      const i = SUB_INDEX.get(g.u);
      if (i === undefined) continue;
      // tabela wg kroju, którym zastępujemy (rodzina), nie flagi „szeryfowy” z PDF — Ghostscript
      // jej nie ustawia i „WtTimes” porównywaliśmy z Arialem (wychodziło zagęszczenie zamiast rozstrzelenia)
      const serifSub = g.font.serif || /times|georgia|cambria|garamond|book antiqua/i.test(g.font.family || "");
      const tab = SUB_WIDTHS[(serifSub ? "times" : "arial") + (g.bold ? "B" : "")];
      const key = g.fontRef + "|" + (g.bold ? 1 : 0);
      let a = acc.get(key);
      if (!a) acc.set(key, (a = { orig: 0, pred: 0, em: 0, n: 0 }));
      a.orig += step;
      // skala pozioma liczy się tylko, gdy trafi do pliku jako w:w (styleKey: odchyłka > 5%);
      // mniejszą (np. 95,7% w DC-85) nadrabia odstęp między znakami
      const xs = g.xScale && Math.abs(g.xScale - 1) > 0.05 ? g.xScale : 1;
      a.pred += (tab[i] / 1000) * g.size * xs;
      a.em += g.size;
      a.n++;
    }
    const perEm = new Map();
    for (const [k, a] of acc) {
      if (a.n < 3 || !a.em) continue;
      const d = (a.orig - a.pred) / a.em; // em na znak (ujemne = zagęścić)
      if (Math.abs(d) >= 0.004 && Math.abs(d) < 0.25) perEm.set(k, d);
    }
    if (!perEm.size) return;
    for (const g of glyphs) {
      const d = perEm.get(g.fontRef + "|" + (g.bold ? 1 : 0));
      if (d !== undefined) g.fitSpacing = Math.round(d * g.size * 20); // dwudzieste punktu (twips)
    }
  }

  function prepareGlyphs(raw, opts = {}) {
    const W = raw.width, H = raw.height;
    const fixes = opts.glyphFixes || null;
    const out = [];
    let invisibleCount = 0;
    for (const g0 of raw.glyphs) {
      if (g0.invisible) {
        invisibleCount++;
        continue;
      }
      const cx = (g0.x + g0.x1) / 2;
      const cy = g0.y - g0.size * 0.3;
      if (cx < -2 || cx > W + 2 || cy < -2 || cy > H + 2) continue;
      const c = g0.clip;
      if (c && (cx < c.x0 - 1 || cx > c.x1 + 1 || cy < c.y0 - g0.size || cy > c.y1 + g0.size)) continue;
      if (g0.size < 1) continue;
      const g = { ...g0 };
      const fix = fixes && fixes.get(g.fontRef + "|" + g.code);
      if (fix) g.u = fix;
      const sym = mapSymbolChar(g);
      if (sym) {
        g.u = sym;
        g.symbol = true;
      }
      // Ligatury rozbijamy (ﬁ → fi), żeby wyszukiwanie i sprawdzanie pisowni działały.
      if (/[ﬀ-ﬆ]/.test(g.u)) g.u = g.u.normalize("NFKC");
      // kod sterujący bez rozpoznanej litery = rysunek (pdf-extract: paint) → warstwa grafiki
      if (g.paint && !fix) {
        if (opts.paints) opts.paints.push({ ...g.paint, order: g.order });
        continue;
      }
      if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(g.u)) continue;
      out.push(g);
    }
    // Strona bez widocznego tekstu, ale z niewidoczną warstwą (PDF po OCR) — bierzemy tę warstwę.
    if (out.filter((g) => g.u.trim()).length < 3 && invisibleCount >= 3) {
      for (const g0 of raw.glyphs) if (g0.invisible) out.push({ ...g0, ocrLayer: true });
    }
    composeAccents(out);
    const live = out.filter((g) => !g.dead);
    fitSubstitutedFonts(live);
    return live;
  }

  const BELOW_ACCENTS = new Set(["\u02db", "\u00b8"]);

  function composeAccents(glyphs) {
    const accents = [];
    for (const g of glyphs) {
      const comb = COMBINING[g.u];
      if (!comb) continue;
      // Akcent jako osobny znak: wąski (przesunięcie < 0,35 em) albo wyraźnie nad/pod linią.
      if (g.adv <= g.size * 0.35) accents.push(g);
    }
    if (!accents.length) return;
    for (const a of accents) {
      const ax = (a.x + a.x1) / 2 + (a.adv < 1 ? a.size * 0.12 : 0);
      let best = null, bestD = Infinity;
      for (const g of glyphs) {
        if (g === a || g.dead || g.isSpace || COMBINING[g.u]) continue;
        if (Math.abs(g.size - a.size) > a.size * 0.3) continue;
        // akcent nad literą: jego linia bazowa wyżej (mniejsze y) albo równa; pod literą
        // (ogonek, cedylla) — równa albo niżej
        const dy = g.y - a.y;
        if (BELOW_ACCENTS.has(a.u) ? (dy > a.size * 0.15 || dy < -a.size * 0.6) : (dy < -a.size * 0.15 || dy > a.size * 0.75)) continue;
        if (!/\p{L}/u.test(g.u)) continue;
        // Środek akcentu w obrębie litery (z marginesem) — akcent w PDF bywa rysowany
        // przed literą z ujemnym przesunięciem albo po niej.
        const gx0 = g.x - a.size * 0.08, gx1 = g.x1 + a.size * 0.08;
        const inside = a.adv < 1 ? (a.x >= gx0 && a.x <= gx1) : (ax >= gx0 && ax <= gx1);
        if (!inside) continue;
        const d = Math.abs((g.x + g.x1) / 2 - (a.adv < 1 ? a.x + a.size * 0.15 : ax)) + Math.abs(g.y - a.y);
        if (d < bestD) {
          bestD = d;
          best = g;
        }
      }
      if (best) {
        const composed = (best.u + COMBINING[a.u]).normalize("NFC");
        if (composed.length === 1 || composed.length === best.u.length) {
          best.u = composed;
          best.composed = true; // znak złożony z dwóch glifów — nie ma go w czcionce jako jednego
          a.dead = true;
        }
      }
    }
  }

  // ---------------------------------------------------------------- wiersze
  // Znaki → fragmenty wierszy (ciągi na tej samej linii bazowej bez dużej przerwy).
  function buildFragments(glyphs) {
    const gs = glyphs.filter((g) => Math.abs(g.angle) < 0.02 || Math.abs(Math.abs(g.angle) - Math.PI) < 0.02);
    const rotated = glyphs.filter((g) => !gs.includes(g));
    gs.sort((a, b) => a.y - b.y || a.x - b.x);
    // 1) pasy linii bazowych
    const rows = [];
    for (const g of gs) {
      let row = null;
      for (let i = rows.length - 1; i >= 0 && i >= rows.length - 4; i--) {
        const r = rows[i];
        const tol = Math.max(1, Math.min(r.size, g.size) * 0.3);
        if (Math.abs(r.y - g.y) <= tol) {
          row = r;
          break;
        }
      }
      if (!row) {
        row = { y: g.y, size: g.size, glyphs: [] };
        rows.push(row);
      }
      row.glyphs.push(g);
      if (g.size > row.size && !g.isSpace) row.size = g.size;
    }
    // indeksy górne/dolne: mały znak przesunięty względem sąsiedniego wiersza dołączamy do niego
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const nonSpace = r.glyphs.filter((g) => !g.isSpace);
      if (!nonSpace.length) continue;
      const maxSize = Math.max(...nonSpace.map((g) => g.size));
      for (const j of [i - 1, i + 1]) {
        const o = rows[j];
        if (!o || o === r || !o.glyphs.length) continue;
        // Rozmiar porównujemy tylko ze znakami OBOK (w poziomie): pas wiersza biegnie przez całą
        // stronę i duży tekst w innej kolumnie na tej samej wysokości (etykieta formularza DC-85)
        // robił z drobnego, ciasno złożonego napisu w trójkącie „indeks dolny” sąsiedniego wiersza.
        const rx0 = Math.min(...r.glyphs.map((g) => g.x)), rx1 = Math.max(...r.glyphs.map((g) => g.x1));
        const reach = maxSize * 1.2;
        const nearG = o.glyphs.filter((g) => !g.isSpace && g.x1 >= rx0 - reach && g.x <= rx1 + reach);
        if (!nearG.length) continue;
        const oMax = Math.max(...nearG.map((g) => g.size));
        if (maxSize > oMax * 0.88) continue;
        const dy = r.y - o.y;
        if (Math.abs(dy) > oMax * 0.55) continue;
        for (const g of r.glyphs) g.vert = dy < 0 ? "superscript" : "subscript";
        o.glyphs.push(...r.glyphs);
        r.glyphs = [];
        break;
      }
    }
    // indeks górny/dolny w tym samym pasie: mniejszy znak z linią bazową wyraźnie wyżej/niżej
    for (const r of rows) {
      const big = r.glyphs.filter((g) => !g.isSpace && g.size >= r.size * 0.9);
      if (!big.length || big.length === r.glyphs.length) continue;
      for (const g of r.glyphs) {
        if (g.isSpace || g.vert || g.size > r.size * 0.88) continue;
        // linia bazowa z DUŻYCH znaków obok (w poziomie) — pas biegnie przez całą stronę, a duży
        // napis w innej kolumnie (trójkąt w DC-85) robił z etykiety po lewej „indeks dolny”
        const reach = r.size * 1.5;
        const near = big.filter((b) => b.x1 >= g.x - reach && b.x <= g.x1 + reach);
        if (!near.length) continue;
        const base = median(near.map((b) => b.y));
        if (g.y < base - r.size * 0.12) g.vert = "superscript";
        else if (g.y > base + r.size * 0.08) g.vert = "subscript";
      }
    }
    // 2) fragmenty: dzielimy wiersz na dużych przerwach
    const frags = [];
    for (const r of rows) {
      if (!r.glyphs.length) continue;
      r.glyphs.sort((a, b) => a.x - b.x || a.order - b.order);
      // zdublowane znaki (pogrubienie przez podwójny druk): ta sama litera prawie w tym samym miejscu
      const dedup = [];
      for (const g of r.glyphs) {
        const p = dedup[dedup.length - 1];
        if (p && p.u === g.u && Math.abs(p.x - g.x) < g.size * 0.08 && Math.abs(p.y - g.y) < g.size * 0.08) {
          p.bold = true;
          continue;
        }
        dedup.push(g);
      }
      let cur = null;
      for (const g of dedup) {
        const size = Math.max(g.size, cur?.size || 0);
        if (cur) {
          const gap = g.x - cur.x1;
          const big = Math.max(size * 2.6, 18);
          if (gap > big) {
            frags.push(finishFrag(cur));
            cur = null;
          }
        }
        if (!cur) cur = { glyphs: [], x0: g.x, x1: g.x1, size: g.size };
        cur.glyphs.push(g);
        if (g.x1 > cur.x1) cur.x1 = g.x1;
        if (g.size > cur.size && !g.isSpace) cur.size = g.size;
      }
      if (cur) frags.push(finishFrag(cur));
    }
    // obrócony tekst: osobne fragmenty (bez obrotu w Wordzie — ważniejsza jest treść)
    if (rotated.length) {
      const groups = new Map();
      for (const g of rotated) {
        const k = Math.round(g.angle * 100) + ":" + Math.round((g.angle > 0 ? g.x : g.x) / 4);
        if (!groups.has(k)) groups.set(k, []);
        groups.get(k).push(g);
      }
      for (const arr of groups.values()) {
        arr.sort((a, b) => (Math.sin(a.angle) < 0 ? b.y - a.y : a.y - b.y));
        const f = finishFrag({ glyphs: arr, x0: Math.min(...arr.map((g) => g.x)), x1: Math.max(...arr.map((g) => g.x1)), size: median(arr.map((g) => g.size)) });
        f.rotated = true;
        f.y = Math.max(...arr.map((g) => g.y));
        f.top = Math.min(...arr.map((g) => g.y - g.asc));
        f.bottom = f.y;
        frags.push(f);
      }
    }
    return frags;
  }

  function finishFrag(f) {
    // usuń spacje z brzegów (obwiednia ma odpowiadać widocznemu tekstowi)
    while (f.glyphs.length && f.glyphs[0].isSpace) f.glyphs.shift();
    while (f.glyphs.length && f.glyphs[f.glyphs.length - 1].isSpace) f.glyphs.pop();
    const gl = f.glyphs.length ? f.glyphs : [];
    const main = gl.filter((g) => !g.vert);
    const base = main.length ? main : gl;
    const sizes = base.filter((g) => !g.isSpace).map((g) => g.size);
    const size = sizes.length ? Math.max(...sizes) : f.size;
    const y = median(base.map((g) => g.y));
    const asc = Math.max(...base.map((g) => g.asc), size * 0.7);
    const desc = Math.max(...base.map((g) => g.desc), size * 0.2);
    return {
      type: "text",
      glyphs: gl,
      x0: gl.length ? Math.min(...gl.map((g) => g.x)) : f.x0,
      x1: gl.length ? Math.max(...gl.map((g) => g.x1)) : f.x1,
      y,
      top: y - asc,
      bottom: y + desc,
      size,
    };
  }

  // ---------------------------------------------------------------- tabele z linii/teł
  function detectTables(raw, frags, W, H) {
    const tol = 1.6;
    let rects = raw.rects.filter((r) => {
      const w = r.x1 - r.x0, h = r.y1 - r.y0;
      if (w > W * 0.97 && h > H * 0.97) return false; // tło całej strony
      return w > 0.05 || h > 0.05;
    });
    // Białe wypełnienia na białej stronie są niewidoczne — nie są krawędziami.
    rects = rects.filter((r) => r.color !== "ffffff" || (r.x1 - r.x0) < 2.5 || (r.y1 - r.y0) < 2.5);
    // Wewnętrzne prostokąty „marginesu komórki” (ten sam kolor wewnątrz większego) — pomijamy.
    const fills = rects.filter((r) => r.x1 - r.x0 >= 2.5 && r.y1 - r.y0 >= 2.5);
    const fillsKept = fills.filter((r) => !fills.some((o) => o !== r && o.color === r.color && o.x0 <= r.x0 + 0.2 && o.x1 >= r.x1 - 0.2 && o.y0 <= r.y0 + 0.2 && o.y1 >= r.y1 - 0.2 && (o.x1 - o.x0) * (o.y1 - o.y0) > (r.x1 - r.x0) * (r.y1 - r.y0) + 0.5));
    const hSegs = [], vSegs = [], specks = [];
    // Pasek w kolorze sąsiedniego tła (Word rysuje tak marginesy komórek) to część tła, nie linia.
    const touchesSameFill = (r) => fillsKept.some((f) => f.color === r.color && f.x0 <= r.x1 + 0.6 && f.x1 >= r.x0 - 0.6 && f.y0 <= r.y1 + 0.6 && f.y1 >= r.y0 - 0.6);
    for (const r of rects) {
      const w = r.x1 - r.x0, h = r.y1 - r.y0;
      const thin = Math.min(w, h);
      if (thin > 0.9 && touchesSameFill(r)) continue;
      if (h < 2.5 && w >= 2.5) hSegs.push({ y: (r.y0 + r.y1) / 2, x0: r.x0, x1: r.x1, w: Math.max(h, 0.25), color: r.color });
      else if (w < 2.5 && h >= 2.5) vSegs.push({ x: (r.x0 + r.x1) / 2, y0: r.y0, y1: r.y1, w: Math.max(w, 0.25), color: r.color });
      // Kwadracik: kropka, kropka nad „i”, kropka w „ż” (tekst zamieniony na krzywe) — nie jest
      // ani linią, ani tłem, ale widać go w PDF, więc trafia do warstwy grafiki.
      else if (w < 2.5 && h < 2.5 && w > 0.2 && h > 0.2 && r.color !== "ffffff") specks.push(r);
    }
    for (const l of raw.lines) {
      const w = l.x1 - l.x0, h = l.y1 - l.y0;
      if (h < 1.2 && w >= 2) hSegs.push({ y: (l.y0 + l.y1) / 2, x0: l.x0, x1: l.x1, w: l.width, color: l.color });
      else if (w < 1.2 && h >= 2) vSegs.push({ x: (l.x0 + l.x1) / 2, y0: l.y0, y1: l.y1, w: l.width, color: l.color });
    }
    // Word dzieli krawędzie komórek na kawałki (na granicach marginesów komórki) — sklejamy
    // współliniowe odcinki, inaczej końce kawałków udają granice wierszy/kolumn.
    const hM = mergeCollinear(hSegs, "y", "x0", "x1");
    const vM = mergeCollinear(vSegs, "x", "y0", "y1");
    hSegs.length = 0;
    hSegs.push(...hM);
    vSegs.length = 0;
    vSegs.push(...vM);
    // elementy siatki: wypełnienia + linie; grupujemy w spójne skupiska
    const items = [
      ...fillsKept.map((r) => ({ kind: "fill", x0: r.x0, y0: r.y0, x1: r.x1, y1: r.y1, ref: r })),
      ...hSegs.map((s) => ({ kind: "h", x0: s.x0, x1: s.x1, y0: s.y - s.w / 2, y1: s.y + s.w / 2, ref: s })),
      ...vSegs.map((s) => ({ kind: "v", x0: s.x - s.w / 2, x1: s.x + s.w / 2, y0: s.y0, y1: s.y1, ref: s })),
    ];
    const parent = items.map((_, i) => i);
    const find = (i) => (parent[i] === i ? i : (parent[i] = find(parent[i])));
    const order = items.map((_, i) => i).sort((a, b) => items[a].x0 - items[b].x0);
    for (let a = 0; a < order.length; a++) {
      const A = items[order[a]];
      for (let b = a + 1; b < order.length; b++) {
        const B = items[order[b]];
        if (B.x0 > A.x1 + tol) break;
        if (B.y0 <= A.y1 + tol && B.y1 >= A.y0 - tol) parent[find(order[a])] = find(order[b]);
      }
    }
    const groups = new Map();
    items.forEach((it, i) => {
      const r = find(i);
      if (!groups.has(r)) groups.set(r, []);
      groups.get(r).push(it);
    });
    const tables = [];
    const leftovers = { hSegs: [], vSegs: [], fills: specks };
    for (const g of groups.values()) {
      const t = buildGrid(g, frags, tol);
      if (t) tables.push(t);
      else {
        for (const it of g) {
          if (it.kind === "h") leftovers.hSegs.push(it.ref);
          if (it.kind === "v") leftovers.vSegs.push(it.ref);
          if (it.kind === "fill") leftovers.fills.push(it.ref);
        }
      }
    }
    // tło całej strony (nie białe) też rysujemy w warstwie grafiki
    for (const r of raw.rects) if (r.x1 - r.x0 > W * 0.97 && r.y1 - r.y0 > H * 0.97 && r.color !== "ffffff") leftovers.fills.push(r);
    return { tables: mergeStackedTables(tables), leftovers };
  }

  // Tabela z Worda bywa narysowana w kawałkach przedzielonych wierszami bez tła i krawędzi
  // (np. „Świadczenia z tytułu…”). Kawałki o tych samych brzegach i z małą przerwą składamy
  // z powrotem w jedną tabelę; przerwa staje się wierszem scalonym na całą szerokość.
  function mergeStackedTables(tables) {
    if (tables.length < 2) return tables;
    const sorted = tables.slice().sort((a, b) => a.y0 - b.y0);
    const out = [sorted[0]];
    for (let i = 1; i < sorted.length; i++) {
      const a = out[out.length - 1], b = sorted[i];
      const gap = b.y0 - a.y1;
      const sameEdges = Math.abs(a.x0 - b.x0) < 3 && Math.abs(a.x1 - b.x1) < 3;
      if (sameEdges && gap > -1 && gap < 16) out[out.length - 1] = joinTables(a, b);
      else out.push(b);
    }
    return out;
  }

  function joinTables(a, b) {
    const X = mergeClose(clusterValues([...a.X, ...b.X], 1.2), 2.5);
    const gapRow = b.y0 - a.y1 > 0.8;
    const Y = [...a.Y, ...(gapRow ? [] : []), ...b.Y.slice(gapRow ? 0 : 1)];
    if (!gapRow) Y[a.Y.length - 1] = (a.y1 + b.y0) / 2;
    const near = (arr, v) => {
      let bi = 0, bd = Infinity;
      arr.forEach((x, i) => {
        const d = Math.abs(x - v);
        if (d < bd) {
          bd = d;
          bi = i;
        }
      });
      return bi;
    };
    const cells = [];
    const remap = (c) => {
      const n = { ...c };
      n.c0 = near(X, c.x0);
      n.c1 = Math.max(n.c0, near(X, c.x1) - 1);
      n.r0 = near(Y, c.y0);
      n.r1 = Math.max(n.r0, near(Y, c.y1) - 1);
      n.x0 = X[n.c0]; n.x1 = X[n.c1 + 1]; n.y0 = Y[n.r0]; n.y1 = Y[n.r1 + 1];
      return n;
    };
    for (const c of a.cells) cells.push(remap(c));
    if (gapRow) {
      const r = a.Y.length - 1;
      cells.push({ r0: r, r1: r, c0: 0, c1: X.length - 2, x0: X[0], x1: X[X.length - 1], y0: Y[r], y1: Y[r + 1], fill: null, borders: { top: null, bottom: null, left: null, right: null }, frags: [], gap: true });
    }
    for (const c of b.cells) cells.push(remap(c));
    const nR = Y.length - 1, nC = X.length - 1;
    const cellOf = Array.from({ length: nR }, () => new Array(nC).fill(null));
    for (const c of cells) for (let r = c.r0; r <= c.r1; r++) for (let k = c.c0; k <= c.c1; k++) if (!cellOf[r][k]) cellOf[r][k] = c;
    // dziury (nie powinny wystąpić) — wypełniamy pustymi komórkami
    for (let r = 0; r < nR; r++) for (let k = 0; k < nC; k++) if (!cellOf[r][k]) {
      const c = { r0: r, r1: r, c0: k, c1: k, x0: X[k], x1: X[k + 1], y0: Y[r], y1: Y[r + 1], fill: null, borders: {}, frags: [] };
      cells.push(c);
      cellOf[r][k] = c;
    }
    return { type: "table", x0: X[0], x1: X[nC], y0: Y[0], y1: Y[nR], top: Y[0], bottom: Y[nR], X, Y, cells, cellOf };
  }

  function mergeCollinear(segs, posKey, a0, a1) {
    const sorted = segs.slice().sort((p, q) => p[posKey] - q[posKey] || p[a0] - q[a0]);
    const out = [];
    for (const sg of sorted) {
      const last = out[out.length - 1];
      if (last && Math.abs(last[posKey] - sg[posKey]) < 0.6 && sg[a0] <= last[a1] + 0.6 && last.color === sg.color) {
        last[a1] = Math.max(last[a1], sg[a1]);
        last.w = Math.max(last.w, sg.w);
      } else out.push({ ...sg });
    }
    return out;
  }

  function clusterValues(vals, tol) {
    const s = vals.slice().sort((a, b) => a - b);
    const out = [];
    for (const v of s) {
      const last = out[out.length - 1];
      if (last && v - last.max <= tol) {
        last.max = v;
        last.sum += v;
        last.n++;
      } else out.push({ min: v, max: v, sum: v, n: 1 });
    }
    return out.map((c) => c.sum / c.n);
  }

  function buildGrid(items, frags, tol) {
    const bx0 = Math.min(...items.map((i) => i.x0)), bx1 = Math.max(...items.map((i) => i.x1));
    const by0 = Math.min(...items.map((i) => i.y0)), by1 = Math.max(...items.map((i) => i.y1));
    if (bx1 - bx0 < 8 || by1 - by0 < 4) return null;
    const xs = [], ys = [];
    for (const it of items) {
      if (it.kind === "v") xs.push((it.x0 + it.x1) / 2);
      else if (it.kind === "h") ys.push((it.y0 + it.y1) / 2);
      else {
        xs.push(it.x0, it.x1);
        ys.push(it.y0, it.y1);
      }
    }
    // poziome linie wyznaczają też granice kolumn tabeli (końce), pionowe — wierszy
    for (const it of items) {
      if (it.kind === "h") xs.push(it.x0, it.x1);
      if (it.kind === "v") ys.push(it.y0, it.y1);
    }
    let X = clusterValues(xs, tol);
    let Y = clusterValues(ys, tol);
    // usuń bardzo wąskie pasy (grubość linii) — scalamy granice bliższe niż 2,5 pt
    X = mergeClose(X, 2.5);
    Y = mergeClose(Y, 2.5);
    if (X.length < 2 || Y.length < 2) return null;
    const nC = X.length - 1, nR = Y.length - 1;
    if (nC * nR > 4000) return null;

    // Dowody granic: pionowa granica na x=X[c] w wierszu r, pozioma na y=Y[r] w kolumnie c.
    const vEdge = (x, y0, y1) => {
      let cov = 0;
      for (const it of items) {
        if (it.kind === "v" && Math.abs((it.x0 + it.x1) / 2 - x) <= tol) cov += overlap1d(it.y0, it.y1, y0, y1);
        else if (it.kind === "fill" && (Math.abs(it.x0 - x) <= tol || Math.abs(it.x1 - x) <= tol)) cov += overlap1d(it.y0, it.y1, y0, y1);
      }
      return cov >= (y1 - y0) * 0.55;
    };
    const hEdge = (y, x0, x1) => {
      let cov = 0;
      for (const it of items) {
        if (it.kind === "h" && Math.abs((it.y0 + it.y1) / 2 - y) <= tol) cov += overlap1d(it.x0, it.x1, x0, x1);
        else if (it.kind === "fill" && (Math.abs(it.y0 - y) <= tol || Math.abs(it.y1 - y) <= tol)) cov += overlap1d(it.x0, it.x1, x0, x1);
      }
      return cov >= (x1 - x0) * 0.55;
    };
    // Obszar komórki musi być zamknięty z 4 stron, żeby to była tabela (nie znaczniki cięcia).
    let closed = 0;
    for (let r = 0; r < nR; r++) for (let c = 0; c < nC; c++) {
      if (vEdge(X[c], Y[r], Y[r + 1]) && vEdge(X[c + 1], Y[r], Y[r + 1]) && hEdge(Y[r], X[c], X[c + 1]) && hEdge(Y[r + 1], X[c], X[c + 1])) closed++;
    }
    if (!closed) return null;

    // Tekst, który przecina granicę, oznacza scalenie komórek.
    const inBox = frags.filter((f) => f.type === "text" && !f.rotated && (f.x0 + f.x1) / 2 >= bx0 - 1 && (f.x0 + f.x1) / 2 <= bx1 + 1 && f.y >= by0 - 1 && f.y <= by1 + 1);
    // Pojedyncza komórka z tekstem = ramka/cieniowany akapit; wymagamy treści albo >1 komórki.
    if (nC * nR === 1 && !inBox.length) return null;

    // siatka właścicieli: union-find po komórkach (scalanie, gdy brak granicy)
    const id = (r, c) => r * nC + c;
    const par = Array.from({ length: nR * nC }, (_, i) => i);
    const f2 = (i) => (par[i] === i ? i : (par[i] = f2(par[i])));
    const textCrossesV = (x, y0, y1) => inBox.some((f) => f.y > y0 && f.y < y1 && f.glyphs.some((g, k) => k > 0 && f.glyphs[k - 1].x1 <= x + 0.5 && g.x >= x - 0.5 ? false : g.x < x - 1 && g.x1 > x + 1));
    for (let r = 0; r < nR; r++) for (let c = 0; c + 1 < nC; c++) {
      const x = X[c + 1];
      if (!vEdge(x, Y[r], Y[r + 1]) || textCrossesV(x, Y[r], Y[r + 1])) par[f2(id(r, c + 1))] = f2(id(r, c));
    }
    for (let c = 0; c < nC; c++) for (let r = 0; r + 1 < nR; r++) {
      const y = Y[r + 1];
      if (!hEdge(y, X[c], X[c + 1])) par[f2(id(r + 1, c))] = f2(id(r, c));
    }
    // każda grupa musi być prostokątem; jeśli nie — rozbijamy do prostokąta otaczającego
    const groups = new Map();
    for (let r = 0; r < nR; r++) for (let c = 0; c < nC; c++) {
      const k = f2(id(r, c));
      if (!groups.has(k)) groups.set(k, { r0: r, r1: r, c0: c, c1: c, n: 0 });
      const g = groups.get(k);
      g.r0 = Math.min(g.r0, r); g.r1 = Math.max(g.r1, r); g.c0 = Math.min(g.c0, c); g.c1 = Math.max(g.c1, c); g.n++;
    }
    const cellOf = Array.from({ length: nR }, () => new Array(nC).fill(null));
    const cells = [];
    for (const g of groups.values()) {
      const area = (g.r1 - g.r0 + 1) * (g.c1 - g.c0 + 1);
      if (area !== g.n) {
        // nieprostokątne scalenie (rzadkie) — rozbij na pojedyncze komórki
        for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) if (!cellOf[r][c] && f2(id(r, c)) === f2(id(g.r0, g.c0))) {
          const cell = { r0: r, r1: r, c0: c, c1: c };
          cells.push(cell);
          cellOf[r][c] = cell;
        }
        continue;
      }
      const cell = { r0: g.r0, r1: g.r1, c0: g.c0, c1: g.c1 };
      cells.push(cell);
      for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) cellOf[r][c] = cell;
    }
    for (const cell of cells) {
      cell.x0 = X[cell.c0]; cell.x1 = X[cell.c1 + 1]; cell.y0 = Y[cell.r0]; cell.y1 = Y[cell.r1 + 1];
      // tło: największe wypełnienie pokrywające środek komórki
      const cx = (cell.x0 + cell.x1) / 2, cy = (cell.y0 + cell.y1) / 2;
      let best = null;
      for (const it of items) if (it.kind === "fill" && it.x0 <= cx && it.x1 >= cx && it.y0 <= cy && it.y1 >= cy) {
        const a = (it.x1 - it.x0) * (it.y1 - it.y0);
        if (!best || a < best.a) best = { a, color: it.ref.color };
      }
      cell.fill = best && best.color !== "ffffff" ? best.color : null;
      cell.borders = {
        top: edgeStyle(items, "h", cell.y0, cell.x0, cell.x1, tol),
        bottom: edgeStyle(items, "h", cell.y1, cell.x0, cell.x1, tol),
        left: edgeStyle(items, "v", cell.x0, cell.y0, cell.y1, tol),
        right: edgeStyle(items, "v", cell.x1, cell.y0, cell.y1, tol),
      };
      cell.frags = [];
    }
    return { type: "table", x0: X[0], x1: X[X.length - 1], y0: Y[0], y1: Y[Y.length - 1], top: Y[0], bottom: Y[Y.length - 1], X, Y, cells, cellOf };
  }

  function mergeClose(vals, d) {
    const out = [];
    for (const v of vals) {
      if (out.length && v - out[out.length - 1] < d) out[out.length - 1] = (out[out.length - 1] + v) / 2;
      else out.push(v);
    }
    return out;
  }

  function edgeStyle(items, kind, pos, a0, a1, tol) {
    let best = null;
    for (const it of items) {
      if (it.kind !== kind) continue;
      const p = kind === "h" ? (it.y0 + it.y1) / 2 : (it.x0 + it.x1) / 2;
      if (Math.abs(p - pos) > tol) continue;
      const cov = kind === "h" ? overlap1d(it.x0, it.x1, a0, a1) : overlap1d(it.y0, it.y1, a0, a1);
      if (cov < (a1 - a0) * 0.4) continue;
      const w = it.ref.w || 0.5;
      if (!best || w > best.w) best = { w, color: it.ref.color || "000000" };
    }
    return best;
  }

  // ---------------------------------------------------------------- cięcie XY (kolumny, karty)
  function rowsCount(atoms) {
    const ys = clusterValues(atoms.filter((a) => a.type === "text").map((a) => a.y), 2);
    return ys.length + atoms.filter((a) => a.type !== "text").length * 2;
  }

  // Brzeg obrazu może lekko zachodzić na tekst obok (karty: mapka przykrywa końcówki wierszy
  // listy) — w podziale na kolumny obraz liczymy węższy o IMG_INSET z obu stron.
  const IMG_INSET = 6;
  const xSpan = (a) => (a.type === "image" && a.x1 - a.x0 > IMG_INSET * 4 ? [a.x0 + IMG_INSET, a.x1 - IMG_INSET] : [a.x0, a.x1]);

  function projectionGaps(atoms, axis) {
    const iv = atoms.map((a) => (axis === "x" ? xSpan(a) : [a.top, a.bottom])).sort((p, q) => p[0] - q[0]);
    const gaps = [];
    let end = iv.length ? iv[0][1] : 0;
    for (let i = 1; i < iv.length; i++) {
      if (iv[i][0] > end) gaps.push({ at: (end + iv[i][0]) / 2, from: end, to: iv[i][0], size: iv[i][0] - end });
      end = Math.max(end, iv[i][1]);
    }
    return gaps;
  }

  function xyCut(atoms, depth = 0) {
    if (atoms.length <= 1 || depth > 400) return { type: "flow", atoms };
    const sizes = atoms.filter((a) => a.type === "text").map((a) => a.size);
    const ms = median(sizes) || 10;
    // pionowe cięcie: kolumny/karty obok siebie
    let bestV = null;
    for (const g of projectionGaps(atoms, "x")) {
      const left = atoms.filter((a) => xSpan(a)[1] <= g.at), right = atoms.filter((a) => xSpan(a)[0] >= g.at);
      // Obok obrazu/tabeli wystarczy styk (obraz i tak „wcięty”) i jeden wiersz tekstu.
      const lObj = left.some((a) => a.type !== "text"), rObj = right.some((a) => a.type !== "text");
      const minGap = lObj || rObj ? 0.5 : Math.max(ms * 1.4, 10);
      if (g.size < minGap) continue;
      const okL = rowsCount(left) >= 2 || (rObj && left.length > 0) || lObj;
      const okR = rowsCount(right) >= 2 || (lObj && right.length > 0) || rObj;
      if (!okL || !okR) continue;
      if (!bestV || g.size > bestV.size) bestV = { ...g, left, right };
    }
    const hGaps = projectionGaps(atoms, "y");
    let bestH = null;
    for (const g of hGaps) if (!bestH || g.size > bestH.size) bestH = g;
    if (bestV && (!bestH || bestV.size >= bestH.size * 0.7)) {
      return { type: "cols", gap: bestV, children: [xyCut(bestV.left, depth + 1), xyCut(bestV.right, depth + 1)] };
    }
    // Dzielimy w poziomie na największej przerwie i schodzimy niżej (kolumny mogą być w części
    // strony). Przy bardzo gęstych stronach — tylko gdy szybki test widzi możliwe kolumny.
    if (bestH && (atoms.length <= 600 || hasAnyVerticalCut(atoms, ms))) {
      const top = atoms.filter((a) => (a.top + a.bottom) / 2 < bestH.at), bottom = atoms.filter((a) => (a.top + a.bottom) / 2 >= bestH.at);
      if (top.length && bottom.length) return { type: "stack", children: [xyCut(top, depth + 1), xyCut(bottom, depth + 1)] };
    }
    return { type: "flow", atoms };
  }

  // Szybkie sprawdzenie, czy gdziekolwiek w tym zbiorze jest sens szukać kolumn (oszczędza
  // rekurencję na zwykłych stronach z jedną kolumną tekstu).
  function hasAnyVerticalCut(atoms, ms) {
    const bands = [];
    const sorted = atoms.slice().sort((a, b) => a.top - b.top);
    let cur = null;
    for (const a of sorted) {
      if (cur && a.top < cur.bottom + ms * 3) {
        cur.atoms.push(a);
        cur.bottom = Math.max(cur.bottom, a.bottom);
      } else {
        cur = { atoms: [a], bottom: a.bottom };
        bands.push(cur);
      }
    }
    for (const b of bands) {
      if (b.atoms.length < 4) continue;
      for (const g of projectionGaps(b.atoms, "x")) if (g.size >= Math.max(ms * 1.4, 10)) {
        const left = b.atoms.filter((a) => a.x1 <= g.at), right = b.atoms.filter((a) => a.x0 >= g.at);
        if (rowsCount(left) >= 2 && rowsCount(right) >= 2) return true;
      }
    }
    return false;
  }

  function flattenLayout(node) {
    // stack z samych „flow” → jeden flow (kolejność od góry)
    if (node.type === "flow") return node;
    if (node.type === "cols") return { ...node, children: node.children.map(flattenLayout) };
    const kids = node.children.map(flattenLayout);
    const out = [];
    for (const k of kids) {
      const items = k.type === "stack" ? k.children : [k];
      for (const it of items) {
        const last = out[out.length - 1];
        if (it.type === "flow" && last && last.type === "flow") last.atoms = last.atoms.concat(it.atoms);
        else out.push(it.type === "flow" ? { type: "flow", atoms: it.atoms.slice() } : it);
      }
    }
    return out.length === 1 ? out[0] : { type: "stack", children: out };
  }

  // ---------------------------------------------------------------- akapity
  const BULLET_RE = /^([•▪■●○◦‣⁃∙➢➔❖◻□✓✗→⇒◆◊−⋅\-–—*o§])$/;
  const NUMBER_RE = /^(\(?\d{1,3}[.)]|\(?[a-zA-Z][.)]|[IVXLC]{1,5}[.)])$/;

  // Fragmenty z tej samej linii bazowej (w obrębie obszaru) → jeden wiersz z „kawałkami”.
  function linesFromFrags(frags) {
    const sorted = frags.slice().sort((a, b) => a.y - b.y || a.x0 - b.x0);
    const lines = [];
    for (const f of sorted) {
      const last = lines[lines.length - 1];
      if (last && Math.abs(last.y - f.y) <= Math.max(1, Math.min(last.size, f.size) * 0.35) && !f.rotated && !last.rotated) {
        last.parts.push(f);
        last.x0 = Math.min(last.x0, f.x0);
        last.x1 = Math.max(last.x1, f.x1);
        last.top = Math.min(last.top, f.top);
        last.bottom = Math.max(last.bottom, f.bottom);
        last.size = Math.max(last.size, f.size);
      } else {
        lines.push({ parts: [f], x0: f.x0, x1: f.x1, y: f.y, top: f.top, bottom: f.bottom, size: f.size, rotated: !!f.rotated });
      }
    }
    for (const l of lines) l.parts.sort((a, b) => a.x0 - b.x0);
    return lines;
  }

  // Wiersz → ciąg elementów: tekst (ze stylem), tabulatory (z pozycją i wypełnieniem kropkami).
  function lineTokens(line, regionX0) {
    const toks = [];
    let prev = null;
    const pushText = (g, text) => {
      const style = styleKey(g);
      const last = toks[toks.length - 1];
      if (last && last.type === "text" && last.key === style.key && last.link === (g.link || null)) last.text += text;
      else toks.push({ type: "text", text, ...style, link: g.link || null });
    };
    for (let p = 0; p < line.parts.length; p++) {
      const part = line.parts[p];
      let glyphs = part.glyphs;
      if (p > 0) {
        // przerwa między kawałkami = tabulator do pozycji kawałka
        toks.push({ type: "tab", pos: part.x0 - regionX0, leader: null, align: "left", startX: part.x0, endX: part.x1 });
        prev = null;
      }
      // Kropki/kreski wiodące (spis treści): ≥4 kropki z rzędu → tabulator z wypełnieniem.
      let i = 0;
      while (i < glyphs.length) {
        const g = glyphs[i];
        if ((g.u === "." || g.u === "…" || g.u === "_" || g.u === "-") && !g.isSpace) {
          let j = i;
          while (j < glyphs.length && (glyphs[j].u === g.u || glyphs[j].isSpace)) j++;
          const run = glyphs.slice(i, j).filter((x) => !x.isSpace);
          const span = run.length ? run[run.length - 1].x1 - run[0].x : 0;
          if (run.length >= (g.u === "_" ? 6 : 5) && span > g.size * 2.2) {
            const next = glyphs[j];
            const leader = g.u === "_" ? "underscore" : g.u === "-" ? "hyphen" : "dot";
            // Kropki do końca wiersza (formularz „Miejscowość: ........”) — tabulator do końca kropek.
            // Spis treści: po kropkach tylko krótki numer na końcu wiersza → tabulator prawy.
            const endX = next ? next.x : run[run.length - 1].x1;
            const tail = glyphs.slice(j).filter((x) => !x.isSpace);
            const isPageNo = next && p === line.parts.length - 1 && tail.length <= 6 && tail.every((x) => !/[.…_]/.test(x.u));
            if (!next) {
              toks.push({ type: "tab", pos: endX - regionX0, leader, align: "left", startX: run[0].x, endX, fill: true, trailing: true });
            } else {
              toks.push({ type: "tab", pos: endX - regionX0, leader, align: isPageNo ? "right" : "left", leaderTo: endX, startX: run[0].x, fill: true });
            }
            prev = null;
            i = j;
            continue;
          }
        }
        if (prev) {
          const gap = g.x - prev.x1;
          const sp = Math.max(prev.size, g.size) * 0.17;
          if (!g.isSpace && !prev.isSpace && gap > sp) {
            if (gap > Math.max(prev.size, g.size) * 1.6) {
              toks.push({ type: "tab", pos: g.x - regionX0, leader: null, align: "left", startX: g.x });
            } else pushText(g, " ");
          }
        }
        if (g.isSpace) {
          // spacja tylko jedna (wyjustowany tekst ma szerokie spacje, ale to dalej jedna spacja)
          const last = toks[toks.length - 1];
          if (prev && !(last && last.type === "text" && last.text.endsWith(" "))) pushText(g, " ");
        } else pushText(g, g.u);
        prev = g;
        i++;
      }
    }
    // spacje na brzegach tekstu przy tabulatorze
    for (let k = 0; k < toks.length; k++) {
      if (toks[k].type !== "tab") continue;
      if (k > 0 && toks[k - 1].type === "text") toks[k - 1].text = toks[k - 1].text.replace(/ +$/, "");
      if (k + 1 < toks.length && toks[k + 1].type === "text") toks[k + 1].text = toks[k + 1].text.replace(/^ +/, "");
    }
    return toks.filter((t) => t.type !== "text" || t.text.length);
  }

  function styleKey(g) {
    const size = Math.round(g.size * 2) / 2;
    const s = {
      font: g.font.family,
      size,
      bold: !!g.bold,
      italic: !!g.italic,
      color: g.color && g.color !== "000000" ? g.color : null,
      vert: g.vert || null,
      underline: !!g.underline,
      strike: !!g.strike,
      hScale: g.xScale && Math.abs(g.xScale - 1) > 0.05 ? Math.round(g.xScale * 100) : null,
      spacing: g.fitSpacing && Math.abs(g.fitSpacing) >= 2 ? g.fitSpacing : null,
    };
    if (s.vert) s.size = Math.round(size * 2 / 1.0) / 2; // indeks: zachowaj prawdziwy rozmiar
    s.key = [s.font, s.size, s.bold, s.italic, s.color, s.vert, s.underline, s.strike, s.hScale, s.spacing].join("|");
    return s;
  }

  // Szerokość pierwszego słowa wiersza (do bezpiecznego zapasu przy zawijaniu).
  function firstWordWidth(line, size) {
    const gl = (line.parts || []).flatMap((p) => p.glyphs || []).sort((a, b) => a.x - b.x);
    let x1 = null;
    for (let k = 0; k < gl.length; k++) {
      const g = gl[k];
      if (g.isSpace && x1 !== null) break;
      if (x1 !== null && g.x - x1 > size * 0.25) break;
      if (!g.isSpace) x1 = Math.max(x1 ?? g.x1, g.x1);
    }
    return x1 === null ? size : Math.max(0, x1 - line.x0);
  }

  function lineStartsList(line) {
    const first = line.parts[0];
    const gl = first.glyphs.filter((g) => !g.isSpace);
    if (!gl.length) return null;
    // znacznik = pierwsze „słowo”, po nim przerwa ≥ 0,3 em albo tabulator
    let word = "";
    let k = 0;
    while (k < gl.length && (k === 0 || gl[k].x - gl[k - 1].x1 < gl[k].size * 0.25)) {
      word += gl[k].u;
      k++;
      if (word.length > 6) break;
    }
    const after = k < gl.length ? gl[k] : line.parts[1]?.glyphs?.find((g) => !g.isSpace);
    if (!after) return null;
    const gap = after.x - gl[k - 1].x1;
    // „o”, „-”, „*” jako punktor tylko z wyraźnym odstępem (inaczej to zwykłe słowo: „o tym”)
    const weak = /^[o\-*]$/.test(word);
    if (BULLET_RE.test(word) && gap > gl[0].size * (weak ? 0.45 : 0.2)) return { kind: "bullet", marker: word, textX: after.x, markerX: gl[0].x };
    if (NUMBER_RE.test(word) && gap > gl[0].size * 0.25) return { kind: "number", marker: word, textX: after.x, markerX: gl[0].x };
    return null;
  }

  /**
   * Obszar (komórka, kolumna, strona) → lista bloków: akapity i obrazy w kolejności od góry.
   * box = { x0, x1, top } granice obszaru.
   */
  function flowBlocks(atoms, box) {
    const textFrags = atoms.filter((a) => a.type === "text");
    const others = atoms.filter((a) => a.type !== "text");
    const lines = linesFromFrags(textFrags);
    const items = [
      ...lines.map((l) => ({ kind: "line", top: l.top, bottom: l.bottom, line: l })),
      ...others.map((o) => ({ kind: o.type, top: o.top, bottom: o.bottom, atom: o })),
    ].sort((a, b) => a.top - b.top || (a.kind === "line" ? 1 : -1));
    const blocks = [];
    let para = null;
    const flush = () => {
      if (para) blocks.push(finishParagraph(para, box));
      para = null;
    };
    for (const it of items) {
      if (it.kind !== "line") {
        flush();
        blocks.push(it.atom.type === "image" ? { type: "image", atom: it.atom, top: it.top, bottom: it.bottom } : it.atom);
        continue;
      }
      const l = it.line;
      l.list = lineStartsList(l);
      l.tokens = lineTokens(l, box.x0);
      if (para && continuesParagraph(para, l, box)) {
        para.lines.push(l);
      } else {
        flush();
        para = { lines: [l] };
      }
    }
    flush();
    // Pojedyncze wiersze, których lewa krawędź pokrywa się z innym akapitem obszaru, są
    // wyrównane do lewej (lista „14/131 - 140” pod „1/1 - 10”), nie do prawej/środka.
    const paras = blocks.filter((b) => b.type === "para");
    for (const b of paras) {
      if (b.nLines !== 1 || b.align === "left" || b.x0 == null) continue;
      const twin = paras.some((o) => o !== b && Math.abs(o.x0 - b.x0) < 1.5 && (o.align === "left" || o.nLines > 1 || Math.abs(o.x1 - b.x1) > 3));
      if (twin) {
        b.align = "left";
        b.indLeft = Math.max(0, b.x0 - box.x0);
        b.indRight = 0;
      }
    }
    return blocks;
  }

  function lineStyleSig(l) {
    const t = l.tokens.find((x) => x.type === "text");
    return t ? `${t.font}|${t.bold}|${t.italic}` : "";
  }

  function continuesParagraph(para, l, box) {
    const prev = para.lines[para.lines.length - 1];
    if (l.rotated || prev.rotated) return false;
    if (l.list) return false;
    if (prev.parts.length > 1 || l.parts.length > 1) return false; // wiersze z tabulatorami osobno
    if (prev.parts.some((p) => p.fromField) || l.parts.some((p) => p.fromField)) return false; // wiersze pola formularza
    // skan po OCR: łączymy tylko wiersze z tego samego akapitu wg rozpoznawania
    const op = prev.parts[0].glyphs[0]?.ocrPara, ol = l.parts[0].glyphs[0]?.ocrPara;
    if ((op || ol) && op !== ol) return false;
    if (prev.tokens.some((t) => t.type === "tab") || l.tokens.some((t) => t.type === "tab")) return false;
    const size = Math.max(prev.size, l.size);
    // inna wielkość liter = inny akapit (nagłówek 13 pt nad tekstem 12 pt)
    if (Math.abs(prev.size - l.size) > Math.max(0.5, size * 0.04)) return false;
    const gap = l.y - prev.y; // odstęp linii bazowych
    if (gap <= 0) return false;
    if (gap > size * 1.75) return false;
    if (para.lines.length >= 2) {
      const pg = prev.y - para.lines[para.lines.length - 2].y;
      if (Math.abs(gap - pg) > size * 0.25) return false;
    }
    // poprzedni wiersz musi sięgać prawie do prawej krawędzi (inaczej to koniec akapitu), a
    // pierwsze słowo nowego wiersza nie zmieściłoby się na końcu poprzedniego
    const width = box.x1 - box.x0;
    const right = Math.max(maxRight(para), l.x1);
    const prevFill = prev.x1 >= right - Math.max(size * 3.2, width * 0.06);
    if (!prevFill) return false;
    const fw = firstWordWidth(l);
    if (fw > 0 && right - prev.x1 > fw + size * 0.3) return false;
    // lewa krawędź: równa z poprzednim (lub z wcięciem pierwszego wiersza / wiszącym)
    const first = para.lines[0];
    const leftRef = para.lines.length >= 2 ? prev.x0 : first.list ? first.list.textX : first.x0;
    const dx = Math.abs(l.x0 - leftRef);
    const centered = isCenteredLine(prev, box) && isCenteredLine(l, box);
    if (dx > size * 0.6 && !centered) {
      // wcięcie pierwszego wiersza: drugi wiersz bardziej na lewo niż pierwszy
      if (!(para.lines.length === 1 && l.x0 < first.x0 && first.x0 - l.x0 < size * 4)) return false;
    }
    if (lineStyleSig(prev) !== lineStyleSig(l)) {
      // zmiana kroju na całym wierszu (np. pogrubiony nagłówek) = nowy akapit
      const pt = prev.tokens.filter((t) => t.type === "text");
      if (pt.length === 1 && l.tokens.filter((t) => t.type === "text").length === 1) return false;
    }
    return true;
  }

  function firstWordWidth(l) {
    const gl = l.parts[0].glyphs;
    let w = 0;
    for (let i = 0; i < gl.length; i++) {
      if (gl[i].isSpace || (i > 0 && gl[i].x - gl[i - 1].x1 > gl[i].size * 0.17)) break;
      w = gl[i].x1 - gl[0].x;
    }
    return w;
  }

  function maxRight(para) {
    return Math.max(...para.lines.map((l) => l.x1));
  }

  function isCenteredLine(l, box) {
    const c = (l.x0 + l.x1) / 2, bc = (box.x0 + box.x1) / 2;
    return Math.abs(c - bc) < Math.max(3, l.size * 0.4) && l.x0 > box.x0 + l.size;
  }

  function finishParagraph(para, box) {
    const lines = para.lines;
    const first = lines[0], last = lines[lines.length - 1];
    const size = median(lines.map((l) => l.size));
    const lefts = lines.map((l) => l.x0), rights = lines.map((l) => l.x1);
    const boxW = box.x1 - box.x0;
    let align = "left";
    const multi = lines.length > 1;
    const tol = Math.max(2, size * 0.35);
    if (multi) {
      const body = lines.slice(0, -1);
      const rightsAligned = body.every((l) => Math.abs(l.x1 - Math.max(...rights)) < tol);
      const leftsAligned = lines.slice(1).every((l) => Math.abs(l.x0 - lines[1].x0) < tol) || lines.length === 2;
      const centers = lines.map((l) => (l.x0 + l.x1) / 2);
      const centered = centers.every((c) => Math.abs(c - centers[0]) < tol) && !leftsAligned;
      if (centered) align = Math.abs(centers[0] - (box.x0 + box.x1) / 2) < tol * 2 ? "center" : "center";
      else if (rightsAligned && leftsAligned && body.length >= 1) {
        // wyjustowany, gdy odstępy międzywyrazowe różnią się w wierszach albo prawa krawędź równa
        align = lines.length >= 2 && body.length >= 1 ? "both" : "left";
        // wiersz bez spacji (np. jedno słowo) — zostaw do lewej
        if (body.every((l) => !l.tokens.some((t) => t.type === "text" && / /.test(t.text)))) align = "left";
      } else if (lines.every((l) => Math.abs(l.x1 - rights[0]) < tol) && !leftsAligned) align = "right";
    } else if (!first.tokens.some((t) => t.type === "tab")) {
      const c = (first.x0 + first.x1) / 2;
      const bc = (box.x0 + box.x1) / 2;
      const nearLeft = first.x0 - box.x0 < size * 1.5;
      const nearRight = box.x1 - first.x1 < size * 1.0;
      if (Math.abs(c - bc) < Math.max(3, size * 0.5) && !nearLeft) align = "center";
      else if (nearRight && !nearLeft && first.parts.length === 1) align = "right";
    }
    // wcięcia
    let indLeft, firstLine = 0;
    const list = first.list;
    if (list) {
      indLeft = list.textX - box.x0;
      firstLine = list.markerX - list.textX; // ujemne = wiszące
    } else if (multi) {
      indLeft = Math.min(...lines.slice(1).map((l) => l.x0)) - box.x0;
      firstLine = first.x0 - box.x0 - indLeft;
      if (Math.abs(firstLine) < 1) firstLine = 0;
    } else {
      indLeft = first.x0 - box.x0;
    }
    if (align === "center" || align === "right") {
      indLeft = 0;
      firstLine = 0;
    }
    if (indLeft < 0.5) indLeft = Math.max(0, indLeft);
    let indRight = 0;
    if (align === "right") indRight = Math.max(0, box.x1 - Math.max(...rights));
    else if (!multi && align === "left" && !first.tokens.some((t) => t.type === "tab" && t.align === "right")) {
      // jeden wiersz w PDF = jeden wiersz w Wordzie: zapas na prawo (wcięcie ujemne), żeby
      // minimalnie szersza czcionka zastępcza nie zawinęła ostatniego słowa
      indRight = -Math.min(28, size * 2);
    } else if (multi && align !== "center") {
      // szerokość zawijania jak w PDF: najdłuższy wiersz + zapas. Zapas możliwie duży (krój
      // zastępczy bywa minimalnie szerszy — „ADMINISTRACYJNE” w DC-85 łamało się na kawałki),
      // ale mniejszy niż to, co wciągnęłoby pierwsze słowo następnego wiersza do poprzedniego:
      // w PDF wiersz i skończył się, bo x1(i) + spacja + pierwsze słowo(i+1) > szerokość.
      const maxR = Math.max(...rights);
      let safe = maxR + size * 1.2;
      for (let i = 0; i + 1 < lines.length; i++) safe = Math.min(safe, lines[i].x1 + size * 0.25 + firstWordWidth(lines[i + 1], size) - 0.5);
      indRight = Math.max(0, box.x1 - Math.max(maxR + size * 0.15 + 1, safe));
      if (align === "both") indRight = Math.max(0, box.x1 - maxR - 0.3);
    }
    // interlinia: mediana odstępów linii bazowych
    let lineGap = null;
    if (multi) {
      const gaps = [];
      for (let i = 1; i < lines.length; i++) gaps.push(lines[i].y - lines[i - 1].y);
      lineGap = median(gaps);
    }
    // treść: tokeny wierszy, łączone spacją (miękkie zawinięcie) albo łamaniem wiersza
    const tokens = [];
    lines.forEach((l, i) => {
      let toks = l.tokens;
      if (i === 0 && list) {
        // znacznik listy + tabulator (wcięcie wiszące = tabulator domyślny w Wordzie)
        const firstText = toks.findIndex((t) => t.type === "text");
        if (firstText >= 0) {
          const t0 = toks[firstText];
          const idx = t0.text.indexOf(list.marker);
          if (idx === 0) {
            const rest = t0.text.slice(list.marker.length).replace(/^ +/, "");
            toks = toks.slice();
            const head = { ...t0, text: list.marker };
            const tail = rest ? [{ ...t0, text: rest }] : [];
            // tabulator po znaczniku już jest, jeśli był duży odstęp
            const hadTab = !rest && toks[firstText + 1]?.type === "tab";
            if (!rest && !hadTab && toks[firstText + 1]?.type === "text") toks[firstText + 1] = { ...toks[firstText + 1], text: toks[firstText + 1].text.replace(/^ +/, "") };
            toks.splice(firstText, 1, head, ...(hadTab ? [] : [{ type: "tab", pos: list.textX - box.x0, align: "left", listTab: true }]), ...tail);
          }
        }
      }
      if (i > 0) {
        const prevT = tokens[tokens.length - 1];
        // dzielenie wyrazu na końcu wiersza: „przy-” + „kład” → zostawiamy myślnik (bezpieczniej)
        if (prevT && prevT.type === "text" && !/[ ­]$/.test(prevT.text)) {
          tokens.push({ ...prevT, type: "text", text: " " });
        }
      }
      for (const t of toks) tokens.push(t);
    });
    const tabs = [];
    for (const t of tokens) if (t.type === "tab" && !t.listTab) {
      const stopPos = t.align === "right" && t.leaderTo ? t.leaderTo - box.x0 : t.pos;
      tabs.push({ pos: stopPos, leader: t.leader, align: t.fill && t.align === "right" ? "right" : t.align === "right" ? "right" : "left" });
    }
    // Jeśli to tabulator z kropkami do numeru strony: pozycja tabulatora prawego = prawy brzeg numeru.
    for (let k = 0; k < tokens.length; k++) {
      const t = tokens[k];
      if (t.type === "tab" && t.fill && t.align === "right") {
        // znajdź koniec tekstu za tabulatorem
        const restLine = lines.find((l) => l.tokens.includes(t));
        if (restLine) {
          const stop = tabs.find((s) => Math.abs(s.pos - (t.leaderTo - box.x0)) < 0.01);
          if (stop) stop.pos = restLine.x1 - box.x0;
        }
      }
    }
    const dedupTabs = [];
    for (const s of tabs.sort((a, b) => a.pos - b.pos)) {
      if (!dedupTabs.some((d) => Math.abs(d.pos - s.pos) < 1.5)) dedupTabs.push(s);
    }
    return {
      type: "para",
      tokens,
      align,
      indLeft,
      indRight,
      firstLine,
      lineGap,
      size,
      tabs: dedupTabs,
      top: first.top,
      bottom: last.bottom,
      firstBaseline: first.y,
      lastBaseline: last.y,
      x0: Math.min(...lefts),
      x1: Math.max(...rights),
      nLines: lines.length,
      list: list ? list.kind : null,
      listMarker: list ? list.marker : null, // pdf-convert.js robi z tego prawdziwą listę Worda
      rotated: first.rotated,
    };
  }

  // ---------------------------------------------------------------- strona
  // Szerokości znaków Helvetiki/Ariala (1/1000 em) — gdy wygląd pola nie podaje szerokości znaku.
  const HELV_W = (() => {
    const t = {};
    const set = (chars, w) => {
      for (const ch of chars) t[ch] = w;
    };
    set(" !,./:;I[\\]ijlt|", 278); set("\"", 355); set("#$0123456789?L_abcdeghnopqsuvxyz", 556);
    set("%", 889); set("&ABEKPSVXY", 667); set("'", 191); set("()-`r{}", 333); set("*", 389); set("+<=>~", 584);
    set("@", 1015); set("CDHNRUw", 722); set("FTZ", 611); set("GOQ", 778); set("J", 500); set("M", 833); set("W", 944);
    set("fk", 500); set("m", 833); set("^", 469);
    t.k = 500; t.f = 278; t.j = 222; t.l = 222; t.i = 222; t.c = 500; t.s = 500; t.v = 500; t.x = 500; t.y = 500; t.z = 500; t.t = 278; t.r = 333;
    return t;
  })();
  function helvWidth(ch) {
    const base = ch.normalize("NFD")[0];
    const map = { "ł": "l", "Ł": "L", "đ": "d", "ø": "o", "Ø": "O" };
    return (HELV_W[ch] ?? HELV_W[map[ch]] ?? HELV_W[base] ?? 556) / 1000;
  }
  NS.helvWidth = helvWidth;

  // Tekst pól formularza i dymków: wartość pola (pełne Unicode), ułożona tam, gdzie rysuje ją
  // wygląd pola (rozmiar, linie bazowe, początek wiersza). Bez wyglądu — z prostokąta pola.
  function annotationFrags(raw) {
    const frags = [];
    for (const a of raw.annotations) {
      if (a.type !== "text" || a.hidden) continue;
      const ag = (a.glyphs || []).filter((g) => !g.isSpace);
      const apLines = [];
      for (const g of ag.slice().sort((p, q) => p.y - q.y || p.x - q.x)) {
        const last = apLines[apLines.length - 1];
        if (last && Math.abs(last.y - g.y) < g.size * 0.4) last.glyphs.push(g);
        else apLines.push({ y: g.y, glyphs: [g] });
      }
      for (const l of apLines) l.glyphs.sort((p, q) => p.x - q.x);
      const textLines = String(a.text).replace(/\r\n?/g, "\n").replace(/\n+$/, "").split("\n");
      const visible = textLines.map((t, i) => ({ t, i })).filter((o) => o.t.trim());
      // Rozmiar „auto” (0 w DA): przeglądarki liczą go różnie — bierzemy wygląd wygenerowany przez
      // pdf.js (rozmiar dopasowany do pola, odstęp ~1,35 em; miernik pdf:score: karty +2…+5 pp).
      // Gdy liczba wierszy wyglądu się nie zgadza — zasada Acrobata (12 pt, mniej gdy się nie mieści).
      const auto = !a.size;
      const h = a.y1 - a.y0;
      const multi = textLines.length > 1;
      const apMatches = apLines.length === visible.length && apLines.length > 0;
      const fromAp = !auto || apMatches;
      let size;
      if (!fromAp) size = Math.max(4, Math.min(12, multi ? (h - 4) / (textLines.length * 1.15) : (h - 2) / 1.15));
      else size = (ag.length ? median(ag.map((g) => g.size)) : 0) || a.size || 10;
      const font = ag[0]?.font || { id: "annot", family: /cour/i.test(a.fontName) ? "Courier New" : /tim/i.test(a.fontName) ? "Times New Roman" : "Arial", bold: false, italic: false, mono: false, serif: false, symbolic: false, asc: 0.72, desc: 0.21 };
      const color = ag[0]?.color || a.color || "000000";
      // szerokości znaków z wyglądu (ten sam krój) — dokładniejsze niż tabela
      const wmap = new Map();
      for (const g of ag) if (g.u.length === 1 && g.adv > 0) wmap.set(g.u, g.adv / g.size);
      const charW = (ch, sz) => (wmap.has(ch) ? wmap.get(ch) : helvWidth(ch)) * sz;
      const leading = fromAp && apLines.length >= 2 ? median(apLines.slice(1).map((l, i) => l.y - apLines[i].y)) : size * 1.15;
      const useAp = fromAp && apMatches;
      const single = textLines.length === 1;
      let y0 = useAp ? apLines[0].y : single ? (a.y0 + a.y1) / 2 + size * 0.3 : a.y0 + 2 + size * 0.85;
      // gdy wygląd ma inną liczbę wierszy — linie bazowe od pierwszej w odstępach z wyglądu
      visible.forEach((o, k) => {
        const tl = o.t.replace(/\s+$/, "");
        const txt = tl.replace(/^\s+/, "");
        let y, x;
        if (useAp) {
          y = apLines[k].y;
          x = apLines[k].glyphs[0].x;
        } else {
          const firstIdx = visible[0].i;
          y = y0 + (o.i - firstIdx) * leading;
          const lead = tl.length - txt.length;
          x = a.x0 + 2 + lead * charW(" ", size);
          if (a.align === 1 || a.align === 2) {
            const w = [...txt].reduce((acc, ch) => acc + charW(ch, size), 0);
            x = a.align === 1 ? (a.x0 + a.x1 - w) / 2 : a.x1 - 2 - w;
          }
        }
        const glyphs = [];
        for (const ch of txt) {
          const w = charW(ch, size);
          glyphs.push({ u: ch, x, x1: x + w, adv: w, y, size, asc: size * 0.72, desc: size * 0.21, angle: 0, font, isSpace: ch === " ", color, bold: font.bold, italic: font.italic, order: 1e9, fromField: true });
          x += w;
        }
        if (glyphs.length) {
          const f = finishFrag({ glyphs, x0: glyphs[0].x, x1: x, size });
          f.fromField = true;
          frags.push(f);
        }
      });
    }
    return frags;
  }

  // Tekst pola leżący na wierszu z kropkami („Miejscowość: ....”) wpisujemy W ten wiersz:
  // kropki pod tekstem znikają, a reszta wiersza zostaje (jak w wydrukowanym formularzu).
  function mergeFieldFrags(frags, fieldFrags) {
    const rest = [];
    for (const ff of fieldFrags) {
      let host = null, bd = Infinity;
      for (const f of frags) {
        if (f.fromField || f.rotated) continue;
        const dy = Math.abs(f.y - ff.y);
        if (dy > Math.max(ff.size, f.size) * 0.75) continue;
        if (ff.x0 > f.x1 + f.size * 0.5 || ff.x1 < f.x0 - f.size * 0.5) continue;
        if (dy < bd) {
          bd = dy;
          host = f;
        }
      }
      if (!host) {
        rest.push(ff);
        continue;
      }
      const covered = (g) => (g.u === "." || g.u === "_" || g.u === "\u2026" || g.isSpace) && g.x1 > ff.x0 - ff.size * 0.3 && g.x < ff.x1 + ff.size * 0.3;
      const kept = host.glyphs.filter((g) => !covered(g));
      for (const g of ff.glyphs) g.y = host.y; // jedna linia bazowa
      const merged = kept.concat(ff.glyphs).sort((p, q) => p.x - q.x);
      const nf = finishFrag({ glyphs: merged, x0: merged[0].x, x1: Math.max(...merged.map((g) => g.x1)), size: Math.max(host.size, ff.size) });
      frags[frags.indexOf(host)] = nf;
    }
    return frags.concat(rest);
  }

  function applyUnderlines(frags, hSegs) {
    // cienka pozioma linia tuż pod linią bazową, pod tekstem = podkreślenie (nie tabela)
    const used = new Set();
    for (const s of hSegs) {
      if (s.w > 2.2) continue;
      for (const f of frags) {
        if (f.type !== "text") continue;
        const dy = s.y - f.y;
        if (dy < -f.size * 0.05 || dy > f.size * 0.35) continue;
        if (overlap1d(s.x0, s.x1, f.x0, f.x1) < 2) continue;
        let hit = false;
        for (const g of f.glyphs) {
          const cx = (g.x + g.x1) / 2;
          if (cx >= s.x0 - 0.5 && cx <= s.x1 + 0.5) {
            g.underline = true;
            hit = true;
          }
        }
        // przekreślenie: linia w okolicy połowy wysokości x
        if (hit) used.add(s);
      }
    }
    for (const s of hSegs) {
      if (used.has(s) || s.w > 2.2) continue;
      for (const f of frags) {
        if (f.type !== "text") continue;
        const dy = f.y - s.y;
        if (dy < f.size * 0.15 || dy > f.size * 0.4) continue;
        if (overlap1d(s.x0, s.x1, f.x0, f.x1) < 2) continue;
        for (const g of f.glyphs) {
          const cx = (g.x + g.x1) / 2;
          if (cx >= s.x0 - 0.5 && cx <= s.x1 + 0.5) g.strike = true;
        }
        used.add(s);
      }
    }
    return hSegs.filter((s) => !used.has(s));
  }

  function applyLinks(frags, annotations) {
    const links = annotations.filter((a) => a.type === "link" && a.url);
    if (!links.length) return;
    for (const f of frags) {
      if (f.type !== "text") continue;
      for (const g of f.glyphs) {
        const cx = (g.x + g.x1) / 2, cy = g.y - g.size * 0.3;
        const L = links.find((a) => cx >= a.x0 - 0.5 && cx <= a.x1 + 0.5 && cy >= a.y0 - 1 && cy <= a.y1 + 1);
        if (L) g.link = L.url;
      }
    }
  }

  // Punktory rysowane jako kształt (kółko/kwadracik — Chromium, InDesign, wiele generatorów):
  // mały wypełniony kształt tuż przed początkiem wiersza tekstu → znak „•”/„▪”/„◦” w tym wierszu.
  function shapeBullets(raw, frags) {
    const used = new Set();
    const cands = [];
    for (const gr of raw.graphics) if (gr.cmds) cands.push({ ref: gr, x0: gr.x0, x1: gr.x1, y0: gr.y0, y1: gr.y1, ch: gr.kind === "fill" ? "\u2022" : "\u25e6" });
    for (const r of raw.rects) cands.push({ ref: r, x0: r.x0, x1: r.x1, y0: r.y0, y1: r.y1, ch: "\u25aa" });
    for (const c of cands) {
      const w = c.x1 - c.x0, h = c.y1 - c.y0;
      if (w < 1 || h < 1 || w > 9 || h > 9 || w / h > 1.6 || h / w > 1.6) continue;
      const cy = (c.y0 + c.y1) / 2;
      let best = null;
      for (const f of frags) {
        if (f.rotated || f.fromField) continue;
        if (w > f.size * 0.7 || h > f.size * 0.7) continue;
        if (cy < f.y - f.size * 0.75 || cy > f.y + f.size * 0.1) continue; // środek kształtu na wysokości liter
        const gap = f.x0 - c.x1;
        if (gap < 0.5 || gap > f.size * 2.2) continue;
        if (!best || gap < best.gap) best = { f, gap };
      }
      if (!best) continue;
      // nic tekstu między kształtem a wierszem / tuż przed kształtem w tym samym wierszu
      const f = best.f;
      if (frags.some((o) => o !== f && Math.abs(o.y - f.y) < f.size * 0.3 && o.x1 > c.x0 - f.size * 1.5 && o.x1 <= c.x0 + 0.5)) continue;
      const g0 = f.glyphs[0];
      f.glyphs.unshift({ ...g0, u: c.ch, x: c.x0, x1: c.x1, adv: w, isSpace: false, vert: null, underline: false, strike: false, link: null, bold: false, italic: false, color: c.ref.color || g0.color, symbol: true, order: (g0.order || 0) - 0.5 });
      f.x0 = c.x0;
      used.add(c.ref);
    }
    return used;
  }

  /**
   * Układ jednej strony.
   * @param raw   wynik extractPage
   * @param opts  { glyphFixes, headerFooter: Set fragmentów pominiętych (nagłówki/stopki) }
   */
  function layoutPage(raw, opts = {}) {
    const W = raw.width, H = raw.height;
    const paints = [];
    const glyphs = prepareGlyphs(raw, { ...opts, paints });
    let frags = buildFragments(glyphs).filter((f) => f.glyphs.length);
    frags = mergeFieldFrags(frags, annotationFrags(raw));
    const bulletShapes = shapeBullets(raw, frags);
    applyLinks(frags, raw.annotations);
    const { tables, leftovers } = detectTables(raw, frags, W, H);
    const rules = applyUnderlines(frags, leftovers.hSegs);

    // przypisanie tekstu do komórek
    const free = [];
    for (const f of frags) {
      const cx = (f.x0 + f.x1) / 2, cy = (f.top + f.bottom) / 2;
      const t = tables.find((t) => cx >= t.x0 - 1 && cx <= t.x1 + 1 && cy >= t.y0 - 1 && cy <= t.y1 + 1);
      if (!t) {
        free.push(f);
        continue;
      }
      // fragment może przechodzić przez kilka komórek (tekst bez podziału) — dzielimy po znakach
      const byCell = new Map();
      for (const g of f.glyphs) {
        const gx = (g.x + g.x1) / 2, gy = g.y - g.size * 0.3;
        const c = t.cells.find((c) => gx >= c.x0 - 0.5 && gx <= c.x1 + 0.5 && gy >= c.y0 - 1 && gy <= c.y1 + 1) || nearestCell(t, gx, gy);
        if (!byCell.has(c)) byCell.set(c, []);
        byCell.get(c).push(g);
      }
      for (const [c, gl] of byCell) {
        if (byCell.size === 1) c.frags.push(f);
        else c.frags.push(finishFrag({ glyphs: gl, x0: gl[0].x, x1: gl[gl.length - 1].x1, size: f.size }));
      }
    }

    // obrazy: w komórce / w tekście / pływające nad lub pod tekstem
    const imageAtoms = [];
    const floats = [];
    const textBoxes = free.map((f) => ({ x0: f.x0, x1: f.x1, y0: f.top, y1: f.bottom }));
    for (const im of raw.images) {
      const b = im.clipped || im;
      const area = (b.x1 - b.x0) * (b.y1 - b.y0);
      if (area < 16) continue;
      const atom = { type: "image", img: im, x0: b.x0, x1: b.x1, top: b.y0, bottom: b.y1 };
      // Obraz na całą stronę (skan) — tło strony
      if (b.x1 - b.x0 > W * 0.85 && b.y1 - b.y0 > H * 0.85) {
        floats.push({ ...atom, behind: true, fullPage: true });
        continue;
      }
      const cell = tables.flatMap((t) => t.cells).find((c) => (b.x0 + b.x1) / 2 >= c.x0 && (b.x0 + b.x1) / 2 <= c.x1 && (b.y0 + b.y1) / 2 >= c.y0 && (b.y0 + b.y1) / 2 <= c.y1);
      if (cell) {
        cell.frags.push(atom);
        continue;
      }
      // nachodzenie tylko na brzeg obrazu (≤ IMG_INSET) to wciąż „obok” — obraz zostaje w tekście
      const ix0 = b.x1 - b.x0 > IMG_INSET * 4 ? b.x0 + IMG_INSET : b.x0, ix1 = b.x1 - b.x0 > IMG_INSET * 4 ? b.x1 - IMG_INSET : b.x1;
      // ten sam próg co w podziale na kolumny (xSpan) — inaczej przy zachodzeniu 6–8 pt obraz nie
      // był ani „obok” (brak przerwy), ani „pływający”, i lądował nad tekstem
      const overlapsText = textBoxes.some((t) => overlap1d(t.x0, t.x1, ix0, ix1) > 0.3 && overlap1d(t.y0, t.y1, b.y0, b.y1) > 2);
      if (overlapsText) {
        const firstText = Math.min(...free.filter((f) => overlap1d(f.x0, f.x1, b.x0, b.x1) > 2 && overlap1d(f.top, f.bottom, b.y0, b.y1) > 2).flatMap((f) => f.glyphs.map((g) => g.order)));
        floats.push({ ...atom, behind: im.order < firstText });
      } else imageAtoms.push(atom);
    }

    // Grafika, która nie stała się tabelą ani podkreśleniem (linie, ramki, tła, znaczniki cięcia,
    // rysunki z krzywych) — rysujemy ją jako przezroczysty obraz pod tekstem, 1:1 z PDF.
    const vectors = [];
    for (const r of leftovers.fills) if (r.color !== "ffffff" && !bulletShapes.has(r)) vectors.push({ type: "rect", x0: r.x0, y0: r.y0, x1: r.x1, y1: r.y1, color: r.color, alpha: r.alpha ?? 1, order: r.order });
    for (const sg of rules) vectors.push({ type: "line", x0: sg.x0, y0: sg.y, x1: sg.x1, y1: sg.y, w: sg.w, color: sg.color });
    for (const sg of leftovers.vSegs) vectors.push({ type: "line", x0: sg.x, y0: sg.y0, x1: sg.x, y1: sg.y1, w: sg.w, color: sg.color });
    for (const gr of raw.graphics) if (gr.cmds && !bulletShapes.has(gr) && !(gr.kind === "fill" && gr.color === "ffffff")) vectors.push({ type: "path", ...gr });
    for (const pg of paints) vectors.push({ type: "glyph", ...pg });

    const atoms = [...free, ...imageAtoms, ...tables];
    const tree = flattenLayout(xyCut(atoms));

    // komórki tabel: bloki w środku
    for (const t of tables) {
      for (const c of t.cells) {
        c.blocks = flowBlocks(c.frags, { x0: c.x0 + 2, x1: c.x1 - 2, top: c.y0 });
      }
    }
    return { width: W, height: H, tree, floats, tables, vectors, glyphs, glyphCount: glyphs.length, hasText: glyphs.length > 0 };
  }

  function nearestCell(t, x, y) {
    let best = t.cells[0], bd = Infinity;
    for (const c of t.cells) {
      const dx = x < c.x0 ? c.x0 - x : x > c.x1 ? x - c.x1 : 0;
      const dy = y < c.y0 ? c.y0 - y : y > c.y1 ? y - c.y1 : 0;
      const d = dx + dy;
      if (d < bd) {
        bd = d;
        best = c;
      }
    }
    return best;
  }

  // Drzewo układu → bloki do zapisu (z obszarami); kolumny → tabela układu bez krawędzi.
  function treeToBlocks(node, box) {
    if (node.type === "flow") return flowBlocks(node.atoms, box);
    if (node.type === "stack") {
      const out = [];
      for (const ch of node.children) out.push(...treeToBlocks(ch, box));
      return out;
    }
    // cols: podziel na kolumny na środku przerwy
    const cols = [];
    const collect = (n) => {
      if (n.type === "cols") {
        collect(n.children[0]);
        collect(n.children[1]);
      } else cols.push(n);
    };
    collect(node);
    const bounds = cols.map((c) => atomsBox(c));
    const textB = cols.map((c) => atomsBox(c, true));
    const edges = [box.x0];
    for (let i = 1; i < cols.length; i++) {
      const L = bounds[i - 1], R = bounds[i];
      if (L.x1 <= R.x0) edges.push((L.x1 + R.x0) / 2);
      // brzeg obrazu zachodzi na tekst obok: granica tuż za tekstem (obraz przesunie się o te
      // kilka punktów), żeby wiersz tekstu się nie zawinął
      else if (textB[i - 1] && textB[i - 1].x1 > R.x0) edges.push(textB[i - 1].x1 + 0.5);
      else if (textB[i] && textB[i].x0 < L.x1) edges.push(textB[i].x0 - 0.5);
      else edges.push((L.x1 + R.x0) / 2);
    }
    edges.push(box.x1);
    const cells = cols.map((c, i) => ({ blocks: treeToBlocks(c, { x0: edges[i], x1: edges[i + 1], top: bounds[i].top }), x0: edges[i], x1: edges[i + 1], top: bounds[i].top }));
    const top = Math.min(...bounds.map((b) => b.top)), bottom = Math.max(...bounds.map((b) => b.bottom));
    return [{ type: "columns", cells, x0: box.x0, x1: box.x1, top, bottom }];
  }

  function atomsBox(node, textOnly = false) {
    let atoms = [];
    const walk = (n) => {
      if (n.type === "flow") atoms.push(...n.atoms);
      else n.children.forEach(walk);
    };
    walk(node);
    if (textOnly) {
      atoms = atoms.filter((a) => a.type === "text");
      if (!atoms.length) return null;
    }
    return {
      x0: Math.min(...atoms.map((a) => a.x0)), x1: Math.max(...atoms.map((a) => a.x1)),
      top: Math.min(...atoms.map((a) => a.top)), bottom: Math.max(...atoms.map((a) => a.bottom)),
    };
  }

  // ---------------------------------------------------------------- dokument
  // Nagłówki/stopki: wiersze przy górnej/dolnej krawędzi powtarzające się na stronach (cyfry
  // mogą się różnić — numer strony) albo oznaczone w PDF jako „Artifact”.
  function detectRunningText(pagesRaw) {
    const sigs = new Map();
    const n = pagesRaw.length;
    if (n < 2) return new Set();
    pagesRaw.forEach((raw, pi) => {
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
        const text = gl.sort((a, b) => a.x - b.x).map((g) => g.u).join("").replace(/\d+/g, "#").replace(/\s+/g, " ").trim();
        if (!text) continue;
        const sig = k.split(":")[0] + "|" + text;
        if (!sigs.has(sig)) sigs.set(sig, new Set());
        sigs.get(sig).add(pi);
      }
    });
    const running = new Set();
    for (const [sig, pages] of sigs) if (pages.size >= Math.max(2, Math.ceil(n * 0.5))) running.add(sig);
    return running;
  }

  NS.layoutPage = layoutPage;
  NS.treeToBlocks = treeToBlocks;
  NS.detectRunningText = detectRunningText;
  NS._internals = { prepareGlyphs, buildFragments, detectTables, xyCut, flowBlocks, linesFromFrags, lineTokens, finishParagraph, annotationFrags, median };
})(typeof globalThis !== "undefined" ? globalThis : window);
