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
    fitWordSpaces(live);
    return live;
  }

  // Spacje węższe niż w kroju: Word drukujący przez sterownik (PrimoPDF, wydruk do PDF ze
  // starego .doc) stawia słowo po spacji ~0,04 em wcześniej, niż wynika z szerokości spacji.
  // Litery się zgadzają, ale wiersz z dziesięcioma spacjami wychodzi u nas ~4 pt szerszy
  // i ostatnie słowo spada do następnego wiersza (cały akapit, a za nim strona, się rozjeżdża).
  // Mierzymy typową przerwę między słowami dla każdego kroju (bez wyjustowanych — szersze od
  // spacji pomijamy) i zagęszczamy same spacje (w:spacing na przebiegu spacji). Tylko kroje
  // z prawdziwymi szerokościami (zastępniki ma fitSubstitutedFonts).
  function fitWordSpaces(glyphs) {
    const acc = new Map();
    for (let k = 1; k + 1 < glyphs.length; k++) {
      const g = glyphs[k], pv = glyphs[k - 1], nx = glyphs[k + 1];
      if (!g.isSpace || pv.isSpace || nx.isSpace || !g.adv || !g.size || g.font?.substituted || g.ocrLayer) continue;
      if (pv.fontRef !== g.fontRef || Math.abs(pv.y - g.y) > 0.2 || Math.abs(nx.y - g.y) > 0.2 || Math.abs(g.angle || 0) > 0.01) continue;
      const gap = nx.x - pv.x1;
      if (gap < g.adv * 0.3 || gap > g.adv * 1.05) continue;
      const key = g.fontRef + "|" + (g.bold ? 1 : 0);
      if (!acc.has(key)) acc.set(key, []);
      acc.get(key).push((gap - g.adv) / g.size);
    }
    const perEm = new Map();
    for (const [k, list] of acc) {
      if (list.length < 8) continue;
      const d = median(list);
      if (d <= -0.012 && d > -0.25) perEm.set(k, d);
    }
    if (!perEm.size) return;
    for (const g of glyphs) {
      if (!g.isSpace || g.fitSpacing !== undefined) continue;
      const d = perEm.get(g.fontRef + "|" + (g.bold ? 1 : 0));
      if (d !== undefined) g.fitSpacing = Math.round(d * g.size * 20);
    }
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
    // Pas wiersza biegnie przez całą stronę, więc rozpatrujemy osobno każdą grupę znaków pasa
    // (przerwa > ~1 em dzieli grupy): ramki „LEGENDA” i „OZNACZENIA” po bokach dużego tytułu na
    // środku infografiki stawały się w całości jego indeksem górnym, a wiersz pod nimi — dolnym.
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const nonSpace = r.glyphs.filter((g) => !g.isSpace);
      if (!nonSpace.length) continue;
      const reach = Math.max(...nonSpace.map((g) => g.size)) * 1.2;
      const groups = [];
      for (const g of r.glyphs.slice().sort((a, b) => a.x - b.x)) {
        const last = groups[groups.length - 1];
        if (last && g.x - last.x1 <= reach) {
          last.glyphs.push(g);
          last.x1 = Math.max(last.x1, g.x1);
        } else groups.push({ glyphs: [g], x0: g.x, x1: g.x1 });
      }
      const moved = new Set();
      for (const grp of groups) {
        const ns = grp.glyphs.filter((g) => !g.isSpace);
        if (!ns.length) continue;
        const maxSize = Math.max(...ns.map((g) => g.size));
        for (const j of [i - 1, i + 1]) {
          const o = rows[j];
          if (!o || o === r || !o.glyphs.length) continue;
          // Rozmiar porównujemy tylko ze znakami OBOK (w poziomie): duży tekst w innej kolumnie na
          // tej samej wysokości (etykieta formularza DC-85) robił z drobnego, ciasno złożonego
          // napisu w trójkącie „indeks dolny” sąsiedniego wiersza.
          const nearG = o.glyphs.filter((g) => !g.isSpace && g.x1 >= grp.x0 - reach && g.x <= grp.x1 + reach);
          if (!nearG.length) continue;
          const oMax = Math.max(...nearG.map((g) => g.size));
          if (maxSize > oMax * 0.88) continue;
          const dy = r.y - o.y;
          if (Math.abs(dy) > oMax * 0.55) continue;
          for (const g of grp.glyphs) {
            g.vert = dy < 0 ? "superscript" : "subscript";
            moved.add(g);
          }
          o.glyphs.push(...grp.glyphs);
          break;
        }
      }
      if (moved.size) r.glyphs = r.glyphs.filter((g) => !moved.has(g));
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
        // Ten sam x i kąt nie oznaczają jeszcze jednego napisu: plan zajęć z Excela ma
        // pionowe nazwy dni w jednej kolumnie. Duża luka na osi pionowej to nowy podpis,
        // nie dalszy ciąg poprzedniego (np. „czwartek” + „piątek”).
        const chunks = [];
        let chunk = [];
        for (const glyph of arr) {
          const prev = chunk[chunk.length - 1];
          const gap = prev ? Math.abs(glyph.y - prev.y) : 0;
          const limit = prev ? Math.max(prev.size, glyph.size) * 2.4 : Infinity;
          if (prev && gap > limit) {
            chunks.push(chunk);
            chunk = [];
          }
          chunk.push(glyph);
        }
        if (chunk.length) chunks.push(chunk);
        for (const part of chunks) {
          const f = finishFrag({ glyphs: part, x0: Math.min(...part.map((g) => g.x)), x1: Math.max(...part.map((g) => g.x1)), size: median(part.map((g) => g.size)) });
          f.rotated = true;
          f.y = Math.max(...part.map((g) => g.y));
          f.top = Math.min(...part.map((g) => g.y - g.asc));
          f.bottom = f.y;
          frags.push(f);
        }
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
  // Podświetlenie tekstu (zakreślacz z PDF-a): kolorowy prostokąt na wysokość jednej linii, który
  // ciasno obejmuje litery z tej jednej linii. Dawniej trafiał do detectTables jako tło komórki —
  // każdy zakreślony wiersz robił się osobną tabelą o stałej szerokości (przesunięty tekst, zła
  // kolejność przy zawiniętym zakreśleniu, ucięcie na telefonie). Teraz litery dostają tło
  // (w:shd w runie, jak zakreślenie w Wordzie), a prostokąt znika z rysunków i z tabel.
  // Zakreślenie z warstwy adnotacji (r.annot) jest nim na pewno. Tło narysowane w treści — tylko
  // samotne: komórki tabeli stykają się z sąsiednimi tłami albo liniami (tak wyglądała tabela
  // w „informacje_ubezpieczenie…”: tło w każdej komórce z jednym wierszem tekstu).
  function applyHighlights(raw, frags) {
    const taken = new Set();
    const glyphs = frags.filter((f) => f.type === "text").flatMap((f) => f.glyphs);
    const near = (a, b2, d) => a.x0 <= b2.x1 + d && a.x1 >= b2.x0 - d && a.y0 <= b2.y1 + d && a.y1 >= b2.y0 - d;
    const edges = [...raw.rects.filter((o) => o.x1 - o.x0 < raw.width * 0.97 || o.y1 - o.y0 < raw.height * 0.97), ...raw.lines.map((l) => ({ x0: Math.min(l.x0, l.x1), x1: Math.max(l.x0, l.x1), y0: Math.min(l.y0, l.y1), y1: Math.max(l.y0, l.y1), line: true }))];
    // zakreślenie zawiniętego tekstu: prostokąty linii pod sobą o innych brzegach — to nie komórki
    const stackedLine = (a, o) => !o.line && Math.abs(a.x0 - o.x0) > 2 && Math.abs(a.x1 - o.x1) > 2;
    const alone = (r) => !edges.some((o) => o !== r && near(r, o, 1) && !stackedLine(r, o));
    // prostokąt o tym samym kolorze i szerokości co inne pasy na stronie (np. nagłówek tabeli nad szarymi
    // pasami harmonogramu) to wiersz tabeli, nie zakreślenie tekstu
    const nonPageRects = raw.rects.filter((o) => o.color && o.color !== "ffffff" && o.x1 - o.x0 >= 4 && o.y1 - o.y0 >= 4 && (o.x1 - o.x0 < raw.width * 0.97 || o.y1 - o.y0 < raw.height * 0.97));
    const hasSiblingBand = (r) => nonPageRects.some((o) => o !== r && o.color === r.color && Math.abs(o.x0 - r.x0) <= 2 && Math.abs(o.x1 - r.x1) <= 2);
    for (const r of raw.rects) {
      const w = r.x1 - r.x0, h = r.y1 - r.y0;
      if (!r.color || r.color === "ffffff" || w < 4 || h < 4) continue;
      if (!r.annot && (!alone(r) || hasSiblingBand(r))) continue;
      const inside = glyphs.filter((g) => {
        const cx = (g.x + g.x1) / 2, cy = g.y - g.size * 0.3;
        return cx >= r.x0 - 0.5 && cx <= r.x1 + 0.5 && cy >= r.y0 - 0.5 && cy <= r.y1 + 0.5;
      });
      const ink = inside.filter((g) => !g.isSpace);
      if (!ink.length) continue;
      const size = Math.max(...ink.map((g) => g.size));
      const ys = ink.map((g) => g.y);
      if (Math.max(...ys) - Math.min(...ys) > size * 0.5) continue; // kilka linii = tło akapitu/komórki
      if (h > size * 1.9) continue;
      const x0 = Math.min(...ink.map((g) => g.x)), x1 = Math.max(...ink.map((g) => g.x1));
      if (x1 - x0 < w * 0.6) continue; // tekst zajmuje mały kawałek prostokąta = komórka, nie zakreślenie
      // wiersz nagłówków kolumn („DATA      MIEJSCE      PROWADZĄCY”) ma szerokie puste przerwy — to pas tabeli, nie zakreślenie
      if (!r.annot) {
        const sortedInk = ink.slice().sort((a, b) => a.x - b.x);
        let wideGapSum = 0;
        for (let i = 1; i < sortedInk.length; i++) {
          const gap = sortedInk[i].x - sortedInk[i - 1].x1;
          if (gap >= size * 3) wideGapSum += gap;
        }
        if (wideGapSum > w * 0.3) continue;
      }
      for (const g of inside) g.shade = r.color;
      taken.add(r);
    }
    return taken.size ? { ...raw, rects: raw.rects.filter((r) => !taken.has(r)) } : raw;
  }

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
    const fillsNonNested = fills.filter((r) => !fills.some((o) => o !== r && o.color === r.color && o.x0 <= r.x0 + 0.2 && o.x1 >= r.x1 - 0.2 && o.y0 <= r.y0 + 0.2 && o.y1 >= r.y1 - 0.2 && (o.x1 - o.x0) * (o.y1 - o.y0) > (r.x1 - r.x0) * (r.y1 - r.y0) + 0.5));
    // Ozdobne pasy tła przy krawędzi kartki (np. szary margines boczny na komentarze Worda na ponad
    // połowę wysokości kartki albo pusty pas przy górnej/dolnej krawędzi): gdyby weszły do union-find
    // tabel, sklejałyby sąsiednie ramki w jedną wielką tabelę na całą stronę. Zostają w warstwie
    // grafiki pod tekstem (leftovers.fills), a ich ewentualny tekst układa xyCut jako osobną kolumnę.
    const isDecorEdgeStrip = (r) => {
      const w = r.x1 - r.x0, h = r.y1 - r.y0;
      const vStrip = (r.x0 <= 3 || r.x1 >= W - 3) && h > H * 0.5 && w < W * 0.45;
      if (vStrip) {
        if (h > H * 0.6) return true;
        return !frags.some((f) => f.type === "text" && (f.x0 + f.x1) / 2 >= r.x0 - 1 && (f.x0 + f.x1) / 2 <= r.x1 + 1 && (f.top + f.bottom) / 2 >= r.y0 - 1 && (f.top + f.bottom) / 2 <= r.y1 + 1);
      }
      const hStrip = (r.y0 <= 3 || r.y1 >= H - 3) && w > W * 0.85 && h < H * 0.25;
      if (!hStrip) return false;
      return !frags.some((f) => f.type === "text" && (f.x0 + f.x1) / 2 >= r.x0 - 1 && (f.x0 + f.x1) / 2 <= r.x1 + 1 && (f.top + f.bottom) / 2 >= r.y0 - 1 && (f.top + f.bottom) / 2 <= r.y1 + 1);
    };
    const decorStrips = fillsNonNested.filter(isDecorEdgeStrip);
    const fillsKept = fillsNonNested.filter((r) => !isDecorEdgeStrip(r));
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
    const leftovers = { hSegs: [], vSegs: [], fills: [...specks, ...decorStrips] };
    const failed = [];
    for (const g of groups.values()) {
      const t = buildGrid(g, frags, tol);
      if (t) tables.push(t);
      else failed.push(g);
    }
    // Pusty pas w serii pasów-tabel (Excel: szare wiersze, ostatnie jeszcze bez wpisów) — ten sam
    // kolor i te same brzegi co wykryta tabela: też pusty wiersz tabeli, nie obraz pod tekstem
    // (dawniej nieedytowalny i znikał w widoku telefonu).
    const sibling = (it) => tables.some((t) => Math.abs(t.x0 - it.x0) < 1.5 && Math.abs(t.x1 - it.x1) < 1.5 && t.cells.some((c) => c.fill === it.ref.color));
    for (const g of failed) {
      const t = g.length === 1 && g[0].kind === "fill" && g[0].ref.color !== "ffffff" && sibling(g[0]) ? buildGrid(g, frags, tol, { allowEmpty: true }) : null;
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
    // Etap 1 (jak w Wordzie): tabele z samymi poziomymi liniami (nowoczesne tabele Worda / booktabs),
    // sklejanie zebry (pasy cieniowane przedzielone białymi wierszami) i podział kolumn z wyrównanego tekstu
    // wewnątrz tabel bez pionowych linii (zebra, sama ramka zewnętrzna).
    const hTables = detectHorizontalTables(leftovers, tables, frags, tol);
    const merged = mergeStackedTables([...tables, ...hTables], frags);
    return { tables: refineImplicitTables(merged, frags), leftovers };
  }

  // Podział tekstu w prostokącie na wiersze i niejawne kolumny (bez pionowych linii).
  // Używane dla tabel z poziomymi liniami, zebry, pasów harmonogramu, tabel z samą ramką oraz tabel bez linii.
  function lineSpansInBox(inBox, ms, opts = {}) {
    const tolY = Math.max(2, ms * 0.4);
    const byY = clusterValues(inBox.map((f) => f.y), tolY);
    const gapMul = opts.spanGapMul ?? 0.85;
    const minGap = opts.minGapPt ?? 7;
    const lines = [];
    for (const y of byY) {
      const fs = inBox.filter((f) => Math.abs(f.y - y) <= tolY);
      const gs = fs.flatMap((f) => f.glyphs.filter((g) => !g.isSpace)).sort((a, b) => a.x - b.x || a.order - b.order);
      if (!gs.length) continue;
      const lineMs = median(fs.map((f) => f.size)) || ms;
      const spanGap = Math.max(lineMs * gapMul, minGap);
      const spans = [];
      for (const g of gs) {
        const last = spans[spans.length - 1];
        if (last && g.x - last.x1 <= spanGap) {
          last.x1 = Math.max(last.x1, g.x1);
          last.glyphs.push(g);
        } else spans.push({ x0: g.x, x1: g.x1, glyphs: [g] });
      }
      const top = Math.min(...fs.map((f) => f.top));
      const bottom = Math.max(...fs.map((f) => f.bottom));
      lines.push({ y, top, bottom, size: lineMs, spans, frags: fs });
    }
    return lines.sort((a, b) => a.y - b.y);
  }

  function inferTextColumns(inBox, bx0, bx1, opts = {}) {
    if (inBox.length < 2) return null;
    // spis treści z kropkami wiodącymi — to akapity z tabulatorem, nie tabela
    if (inBox.some((f) => /\.{4,}|…{2,}/.test(f.text || ""))) return null;
    const ms = median(inBox.map((f) => f.size)) || 10;
    const lines = lineSpansInBox(inBox, ms, opts);
    const multi = lines.filter((l) => l.spans.length >= 2);
    const minMulti = opts.minMultiLines ?? 2;
    if (multi.length < minMulti || multi.length < lines.length * 0.3) return null;
    // rzutowanie na oś X: bierzemy wiersze wielokolumnowe ORAZ wąskie jednosłupkowe linie (np. wyśrodkowana
    // w pionie 3. kolumna w pasie harmonogramu albo zawinięty opis jednej kolumny — bez nich znikają granice kolumn)
    const boxW = Math.max(1, bx1 - bx0);
    const projLines = lines.filter((l) => l.spans.length >= 2 || (l.spans.length === 1 && l.spans[0].x1 - l.spans[0].x0 <= boxW * 0.65));
    const iv = projLines.flatMap((l) => l.spans.map((s) => [s.x0, s.x1])).sort((a, b) => a[0] - b[0]);
    if (!iv.length) return null;
    const minGap = Math.max(ms * (opts.spanGapMul ?? 0.85), opts.minGapPt ?? 7);
    const gutters = [];
    let end = iv[0][1];
    for (let i = 1; i < iv.length; i++) {
      if (iv[i][0] - end >= minGap) gutters.push({ from: end, to: iv[i][0] });
      end = Math.max(end, iv[i][1]);
    }
    if (!gutters.length) return null;
    // granica kolumny: przy lewostronnym wyrównaniu kolumny po prawej — tuż przed jej tekstem (indLeft ≈ 0),
    // inaczej (np. liczby do prawej) — w połowie przerwy
    const padCut = opts.padCut ?? 3;
    const cuts = gutters.map((g, gi) => {
      const nextFrom = gi + 1 < gutters.length ? gutters[gi + 1].from : bx1 + 2;
      const rightStarts = projLines.flatMap((l) => l.spans.filter((s) => s.x0 >= g.to - 1 && s.x1 <= nextFrom + 1).map((s) => s.x0));
      const leftAligned = rightStarts.length >= 2 && Math.max(...rightStarts) - Math.min(...rightStarts) <= 3.5;
      if (leftAligned) return Math.max(g.from + 2, Math.min(g.to - 1.5, Math.min(...rightStarts) - padCut));
      return (g.from + g.to) / 2;
    });
    const X = [bx0, ...cuts, bx1];
    // upewnij się, że co najmniej minMulti wierszy obsadza ≥ 2 różne kolumny
    const colOf = (s) => {
      const cx = (s.x0 + s.x1) / 2;
      for (let c = 0; c < X.length - 1; c++) if (cx >= X[c] - 1 && cx <= X[c + 1] + 1) return c;
      return 0;
    };
    const multiCols = lines.filter((l) => new Set(l.spans.map(colOf)).size >= 2);
    if (multiCols.length < minMulti) return null;
    return { X, lines, colOf, ms };
  }

  // Podział pasma [y0, y1] między poziomymi liniami (albo wewnątrz ramki 1×1) na pod-wiersze:
  // nowa linia tekstu otwiera nowy wiersz, jeśli wnosi treść w kolumnie 0 lub w ≥ 2 kolumnach naraz;
  // zawinięta kontynuacja pojedynczej kolumny (np. opis w kolumnie 1) zostaje w tym samym wierszu.
  function splitBandRows(lines, y0, y1, colOf) {
    const inside = lines.filter((l) => l.y > y0 - 1 && l.y < y1 + 1);
    if (inside.length <= 1) return [y0, y1];
    const rowGroups = [[inside[0]]];
    for (let i = 1; i < inside.length; i++) {
      const l = inside[i];
      const cols = new Set(l.spans.map(colOf));
      const prevCols = new Set(rowGroups[rowGroups.length - 1][0].spans.map(colOf));
      const startsNew = cols.size >= 2 || (cols.has(0) && prevCols.has(0));
      if (startsNew) rowGroups.push([l]);
      else rowGroups[rowGroups.length - 1].push(l);
    }
    if (rowGroups.length <= 1) return [y0, y1];
    const ys = [y0];
    for (let i = 0; i + 1 < rowGroups.length; i++) {
      const a = rowGroups[i][rowGroups[i].length - 1];
      const b = rowGroups[i + 1][0];
      ys.push((a.bottom + b.top) / 2);
    }
    ys.push(y1);
    return ys;
  }

  // Wykrywanie tabel z samych poziomych linii (bez pionowych): nowoczesne tabele Worda i styl akademicki (booktabs).
  function detectHorizontalTables(leftovers, existingTables, frags, tol) {
    const cand = leftovers.hSegs.filter((s) => s.x1 - s.x0 >= 60);
    if (cand.length < 2) return [];
    const stacks = [];
    for (const s of cand) {
      const st = stacks.find((g) => Math.abs(g.x0 - s.x0) <= 6 && Math.abs(g.x1 - s.x1) <= 6);
      if (st) {
        st.segs.push(s);
        st.x0 = Math.min(st.x0, s.x0);
        st.x1 = Math.max(st.x1, s.x1);
      } else stacks.push({ x0: s.x0, x1: s.x1, segs: [s] });
    }
    const used = new Set();
    const out = [];
    const overlapsExisting = (x0, x1, y0, y1) => existingTables.some((t) => overlap1d(t.x0, t.x1, x0, x1) > 4 && overlap1d(t.y0, t.y1, y0, y1) > 4);
    for (const st of stacks) {
      if (st.segs.length < 2) continue;
      const segs = st.segs.slice().sort((a, b) => a.y - b.y);
      // rozdziel stos, jeśli między dwiema liniami jest wielka przerwa albo proza bez kolumn
      const groups = [[segs[0]]];
      for (let i = 1; i < segs.length; i++) {
        const prev = segs[i - 1], cur = segs[i];
        const between = frags.filter((f) => f.type === "text" && !f.rotated && (f.x0 + f.x1) / 2 >= st.x0 - 2 && (f.x0 + f.x1) / 2 <= st.x1 + 2 && f.y > prev.y + 1 && f.y < cur.y - 1);
        const hugeEmpty = !between.length && cur.y - prev.y > 40;
        const proseGap = cur.y - prev.y > 120 && between.length > 0 && !inferTextColumns(between, st.x0, st.x1);
        if (hugeEmpty || proseGap) groups.push([cur]);
        else groups[groups.length - 1].push(cur);
      }
      for (const g of groups) {
        if (g.length < 2) continue;
        let yTop = g[0].y;
        const yBot = g[g.length - 1].y;
        if (yBot - yTop < 12) continue;
        const bx0 = Math.min(...g.map((s) => s.x0)), bx1 = Math.max(...g.map((s) => s.x1));
        if (overlapsExisting(bx0, bx1, yTop, yBot)) continue;
        let inBox = frags.filter((f) => f.type === "text" && !f.rotated && (f.x0 + f.x1) / 2 >= bx0 - 2 && (f.x0 + f.x1) / 2 <= bx1 + 2 && f.y >= yTop - 1 && f.y <= yBot + 1);
        // nagłówek nad pierwszą linią (tabela tylko z podkreśleniem nagłówka i linią na dole)
        const msGuess = median(inBox.map((f) => f.size)) || 10;
        const above = frags.filter((f) => f.type === "text" && !f.rotated && (f.x0 + f.x1) / 2 >= bx0 - 2 && (f.x0 + f.x1) / 2 <= bx1 + 2 && f.bottom <= yTop + 1 && f.top >= yTop - msGuess * 2.2);
        if (above.length >= 2 && !overlapsExisting(bx0, bx1, Math.min(...above.map((f) => f.top)) - 3, yTop)) {
          const withAbove = [...above, ...inBox];
          if (inferTextColumns(withAbove, bx0, bx1)) {
            const rawTop = Math.min(...above.map((f) => f.top)) - 3;
            const prevSame = existingTables.find((t) => Math.abs(t.x0 - bx0) <= 6 && Math.abs(t.x1 - bx1) <= 6 && t.y1 <= rawTop + 2 && rawTop - t.y1 <= msGuess * 1.5);
            yTop = prevSame ? prevSame.y1 : rawTop;
            inBox = withAbove;
          }
        }
        const inferred = inferTextColumns(inBox, bx0, bx1);
        if (!inferred) continue;
        const { X, lines, colOf } = inferred;
        const ruleYs = mergeClose(clusterValues(g.map((s) => s.y), tol), 2.5);
        const anchorYs = yTop < ruleYs[0] - 2 ? [yTop, ...ruleYs] : ruleYs;
        const Y = [anchorYs[0]];
        for (let k = 0; k + 1 < anchorYs.length; k++) {
          const sub = splitBandRows(lines, anchorYs[k], anchorYs[k + 1], colOf);
          Y.push(...sub.slice(1));
        }
        const nR = Y.length - 1, nC = X.length - 1;
        if (nR < 2 || nC < 2) continue;
        const hItems = g.map((s) => ({ kind: "h", x0: s.x0, x1: s.x1, y0: s.y - s.w / 2, y1: s.y + s.w / 2, ref: s }));
        const cellOf = Array.from({ length: nR }, () => new Array(nC).fill(null));
        const cells = [];
        for (let r = 0; r < nR; r++) {
          for (let c = 0; c < nC; c++) {
            const cell = {
              r0: r, r1: r, c0: c, c1: c,
              x0: X[c], x1: X[c + 1], y0: Y[r], y1: Y[r + 1],
              fill: null,
              borders: {
                top: edgeStyle(hItems, "h", Y[r], X[c], X[c + 1], tol),
                bottom: edgeStyle(hItems, "h", Y[r + 1], X[c], X[c + 1], tol),
                left: null,
                right: null,
              },
              frags: [],
            };
            cells.push(cell);
            cellOf[r][c] = cell;
          }
        }
        for (const s of g) used.add(s);
        out.push({ type: "table", x0: X[0], x1: X[nC], y0: Y[0], y1: Y[nR], top: Y[0], bottom: Y[nR], X, Y, cells, cellOf });
      }
    }
    if (used.size) leftovers.hSegs = leftovers.hSegs.filter((s) => !used.has(s));
    return out;
  }

  // Podział tabel jednokolumnowych (np. sama ramka zewnętrzna 1×1, sklejona zebra nR×1 albo pasy
  // harmonogramu z wyśrodkowaną w pionie kolumną), gdy tekst wewnątrz tworzy wyraźną siatkę
  // wielokolumnową, oraz podział pełnoszerokościowego wiersza (np. cieniowany nagłówek sklejony
  // z tabelą z poziomych linii), jeśli jego tekst mieści się czysto w kolumnach tabeli.
  function refineImplicitTables(tables, frags) {
    const isBorderlessBand = (t) => t.X.length === 2 && t.Y.length === 2 && !!t.cellOf?.[0]?.[0]?.fill
      && !t.cellOf[0][0].borders?.top && !t.cellOf[0][0].borders?.bottom
      && !t.cellOf[0][0].borders?.left && !t.cellOf[0][0].borders?.right;
    const fragsIn = (t) => frags.filter((f) => f.type === "text" && !f.rotated && (f.x0 + f.x1) / 2 >= t.x0 - 1 && (f.x0 + f.x1) / 2 <= t.x1 + 1 && f.y >= t.y0 - 1 && f.y <= t.y1 + 1);
    const hasStaggeredCol = (lines, colOf, nC) => {
      const ys = Array.from({ length: nC }, () => []);
      for (const l of lines) {
        for (const c of new Set(l.spans.map(colOf))) ys[c].push(l.y);
      }
      for (let c = 0; c < nC; c++) ys[c].sort((a, b) => a - b);
      for (let c1 = 0; c1 < nC; c1++) {
        if (ys[c1].length !== 1) continue;
        const ym = ys[c1][0];
        for (let c2 = 0; c2 < nC; c2++) {
          if (c2 === c1 || ys[c2].length < 2) continue;
          if (ys[c2][0] + 2 < ym && ym < ys[c2][ys[c2].length - 1] - 2) return true;
        }
      }
      return false;
    };
    // Siostrzane pasy cieniowane o tych samych brzegach (np. harmonogram: nagłówek + wiersze + pusty pas):
    // wspólny podział na kolumny działa też dla 1-wierszowego nagłówka, pasów z brakującym polem i pustych pasów.
    const bandGroups = [];
    for (const t of tables) {
      if (!isBorderlessBand(t)) continue;
      const g = bandGroups.find((gr) => Math.abs(gr.x0 - t.x0) <= 3 && Math.abs(gr.x1 - t.x1) <= 3 && gr.fill === t.cellOf[0][0].fill);
      if (g) g.bands.push(t);
      else bandGroups.push({ x0: t.x0, x1: t.x1, fill: t.cellOf[0][0].fill, bands: [t] });
    }
    const sharedBandX = new Map();
    for (const g of bandGroups) {
      if (g.bands.length < 2) continue;
      const allFrags = g.bands.flatMap((b) => fragsIn(b));
      const inferred = inferTextColumns(allFrags, g.x0, g.x1, { padCut: 2 });
      if (!inferred || inferred.X.length < 3) continue;
      const cuts = inferred.X.slice(1, -1);
      for (const b of g.bands) {
        const gs = fragsIn(b).flatMap((f) => f.glyphs.filter((gl) => !gl.isSpace));
        const crosses = gs.some((gl) => cuts.some((x) => gl.x < x - 1 && gl.x1 > x + 1));
        if (!crosses) sharedBandX.set(b, [b.x0, ...cuts, b.x1]);
      }
    }

    return tables.map((t) => {
      const nC = t.X.length - 1, nR = t.Y.length - 1;
      if (nC >= 2) {
        let changed = false;
        const cells = [];
        const cellOf = t.cellOf.map((row) => row.slice());
        for (const cell of t.cells) {
          if (cell.c0 === 0 && cell.c1 === nC - 1 && cell.r0 === cell.r1) {
            const r = cell.r0;
            const rowFrags = frags.filter((f) => f.type === "text" && !f.rotated && (f.x0 + f.x1) / 2 >= t.x0 - 1 && (f.x0 + f.x1) / 2 <= t.x1 + 1 && f.y >= cell.y0 - 1 && f.y <= cell.y1 + 1);
            const gs = rowFrags.flatMap((f) => f.glyphs.filter((g) => !g.isSpace));
            const crosses = gs.some((g) => t.X.slice(1, -1).some((x) => g.x < x - 1 && g.x1 > x + 1));
            const colsHit = new Set(gs.map((g) => {
              const gx = (g.x + g.x1) / 2;
              for (let c = 0; c < nC; c++) if (gx >= t.X[c] - 1 && gx <= t.X[c + 1] + 1) return c;
              return 0;
            }));
            if (gs.length && !crosses && colsHit.size >= 2) {
              changed = true;
              for (let c = 0; c < nC; c++) {
                const nc = {
                  r0: r, r1: r, c0: c, c1: c,
                  x0: t.X[c], x1: t.X[c + 1], y0: cell.y0, y1: cell.y1,
                  fill: cell.fill || null,
                  borders: {
                    top: cell.borders?.top || null,
                    bottom: cell.borders?.bottom || null,
                    left: c === 0 ? cell.borders?.left || null : null,
                    right: c === nC - 1 ? cell.borders?.right || null : null,
                  },
                  frags: [],
                };
                cells.push(nc);
                cellOf[r][c] = nc;
              }
              continue;
            }
          }
          cells.push(cell);
        }
        return changed ? { ...t, cells, cellOf } : t;
      }
      if (nC !== 1) return t;
      let X, Y;
      if (sharedBandX.has(t)) {
        // pojedynczy pas harmonogramu z serii pasów → 1 wiersz × N kolumn (nie rozcina wyśrodkowanej kolumny)
        X = sharedBandX.get(t);
        Y = t.Y;
      } else {
        const inBox = fragsIn(t);
        const inferred = inferTextColumns(inBox, t.x0, t.x1);
        if (!inferred) return t;
        X = inferred.X;
        Y = t.Y;
        const keepSingleRow = isBorderlessBand(t) && hasStaggeredCol(inferred.lines, inferred.colOf, X.length - 1);
        // 1×1 (np. sama ramka zewnętrzna): podziel także na wiersze, chyba że pas ma kolumnę wyśrodkowaną w pionie
        if (nR === 1 && !keepSingleRow) Y = splitBandRows(inferred.lines, t.y0, t.y1, inferred.colOf);
      }
      const newR = Y.length - 1, newC = X.length - 1;
      if (newC < 2 || (newR < 2 && !sharedBandX.has(t) && !isBorderlessBand(t))) return t;
      const srcRowOf = (r) => {
        if (nR === 1) return t.cellOf[0][0];
        const cy = (Y[r] + Y[r + 1]) / 2;
        for (let sr = 0; sr < nR; sr++) if (cy >= t.Y[sr] - 0.5 && cy <= t.Y[sr + 1] + 0.5) return t.cellOf[sr][0];
        return t.cellOf[Math.min(r, nR - 1)][0];
      };
      const cellOf = Array.from({ length: newR }, () => new Array(newC).fill(null));
      const cells = [];
      for (let r = 0; r < newR; r++) {
        const src = srcRowOf(r);
        const origTop = nR === 1 ? (r === 0 ? src?.borders?.top || null : null) : src?.borders?.top || null;
        const origBot = nR === 1 ? (r === newR - 1 ? src?.borders?.bottom || null : null) : src?.borders?.bottom || null;
        for (let c = 0; c < newC; c++) {
          const cell = {
            r0: r, r1: r, c0: c, c1: c,
            x0: X[c], x1: X[c + 1], y0: Y[r], y1: Y[r + 1],
            fill: src?.fill || null,
            borders: {
              top: origTop,
              bottom: origBot,
              left: c === 0 ? src?.borders?.left || null : null,
              right: c === newC - 1 ? src?.borders?.right || null : null,
            },
            frags: [],
          };
          cells.push(cell);
          cellOf[r][c] = cell;
        }
      }
      return { type: "table", x0: X[0], x1: X[newC], y0: Y[0], y1: Y[newR], top: Y[0], bottom: Y[newR], X, Y, cells, cellOf };
    });
  }

  // Etap 2 (jak w Wordzie): wykrywanie tabel całkowicie bez linii i bez cieniowania (np. 3-kolumnowe
  // cenniki, 2-kolumnowe zestawienia klucz–wartość, małe 2-wierszowe tabelki). Ostre strażniki
  // odrzucają tekst dwukolumnowy (artykuły / proba.pdf), spisy treści z kropkami i pola formularzy.
  function detectBorderlessTables(freeFrags, existingTables, rawImages, W, H) {
    const textFrags = freeFrags.filter((f) => f.type === "text" && !f.rotated && !f.fromField);
    if (textFrags.length < 4) return [];
    const ms = median(textFrags.map((f) => f.size)) || 10;
    const allLines = lineSpansInBox(textFrags, ms, { spanGapMul: 1.35, minGapPt: 11 });
    if (allLines.length < 2) return [];

    const imgBoxes = (rawImages || []).map((im) => im.clipped || im).filter((b) => {
      const w = b.x1 - b.x0, h = b.y1 - b.y0;
      return w * h >= 16 && !(w > W * 0.85 && h > H * 0.85);
    });
    const blockedBox = (x0, x1, y0, y1) => existingTables.some((t) => overlap1d(t.x0, t.x1, x0, x1) > 2 && overlap1d(t.y0, t.y1, y0, y1) > 2)
      || imgBoxes.some((b) => overlap1d(b.x0, b.x1, x0, x1) > 2 && overlap1d(b.y0, b.y1, y0, y1) > 2);
    const spanText = (s) => {
      let out = "";
      for (let i = 0; i < s.glyphs.length; i++) {
        const g = s.glyphs[i];
        if (i > 0 && g.x - s.glyphs[i - 1].x1 > (g.size || 10) * 0.17) out += " ";
        out += g.u;
      }
      return out.trim();
    };
    const isLineEligible = (l) => l.spans.length >= 2
      && !l.frags.some((f) => f.fromField || /\.{4,}|…{2,}|_{4,}/.test(f.text || ""));

    // ciągłe pasma kolejnych wierszy wielokolumnowych o zbliżonym kroju i bez dużych przerw pionowych
    const runs = [];
    let cur = [];
    const flushRun = () => {
      if (cur.length >= 2) runs.push(cur);
      cur = [];
    };
    for (const l of allLines) {
      if (!isLineEligible(l)) {
        flushRun();
        continue;
      }
      if (cur.length) {
        const prev = cur[cur.length - 1];
        const dy = l.y - prev.y;
        const maxSz = Math.max(prev.size, l.size);
        const minSz = Math.min(...cur.map((x) => x.size), l.size);
        const hiSz = Math.max(...cur.map((x) => x.size), l.size);
        const dys = cur.slice(1).map((x, k) => x.y - cur[k].y);
        const medDy = dys.length ? median(dys) : dy;
        const bigJump = cur.length >= 2 && dy > medDy * 1.75 && dy > maxSz * 1.8;
        const midX0 = Math.min(prev.spans[0].x0, l.spans[0].x0);
        const midX1 = Math.max(prev.spans[prev.spans.length - 1].x1, l.spans[l.spans.length - 1].x1);
        const objBetween = existingTables.some((t) => overlap1d(t.x0, t.x1, midX0, midX1) > 2 && t.y0 > prev.bottom - 1 && t.y1 < l.top + 1);
        if (dy <= 0 || dy > maxSz * 3.2 || bigJump || hiSz / Math.max(1, minSz) > 1.35 || objBetween) flushRun();
      }
      cur.push(l);
    }
    flushRun();

    const out = [];
    for (const run of runs) {
      const runFrags = run.flatMap((l) => l.frags);
      const bx0 = Math.min(...run.map((l) => l.spans[0].x0));
      const bx1 = Math.max(...run.map((l) => l.spans[l.spans.length - 1].x1));
      const by0 = run[0].top, by1 = run[run.length - 1].bottom;
      if (blockedBox(bx0, bx1, by0, by1)) continue;
      const inferred = inferTextColumns(runFrags, bx0 - 2, bx1 + 4, { spanGapMul: 1.35, minGapPt: 11, padCut: 2 });
      if (!inferred) continue;
      const { X, lines, colOf, ms: rMs } = inferred;
      const nC = X.length - 1, nR = lines.length;
      if (nC < 2 || nR < 2) continue;
      if (X.slice(1).some((x, c) => x - X[c] < 10)) continue;

      // macierz (r, c) → połączone odcinki w tej komórce
      const grid = Array.from({ length: nR }, () => Array.from({ length: nC }, () => null));
      let filled = 0;
      let allRowsMulti = true;
      for (let r = 0; r < nR; r++) {
        const colsHit = new Set();
        for (const s of lines[r].spans) {
          const c = colOf(s);
          colsHit.add(c);
          if (!grid[r][c]) {
            grid[r][c] = { x0: s.x0, x1: s.x1, glyphs: s.glyphs.slice() };
            filled++;
          } else {
            grid[r][c].x0 = Math.min(grid[r][c].x0, s.x0);
            grid[r][c].x1 = Math.max(grid[r][c].x1, s.x1);
            grid[r][c].glyphs.push(...s.glyphs);
          }
        }
        if (colsHit.size < 2) { allRowsMulti = false; break; }
      }
      if (!allRowsMulti || filled / (nR * nC) < 0.72) continue;

      // sprawdzenie wyrównania kolumn (do lewej, prawej lub środka)
      let alignedCols = 0;
      let col0Aligned = false, col1Aligned = false;
      for (let c = 0; c < nC; c++) {
        const items = [];
        for (let r = 0; r < nR; r++) if (grid[r][c]) items.push(grid[r][c]);
        if (items.length < 2) { alignedCols++; if (c === 0) col0Aligned = true; if (c === 1) col1Aligned = true; continue; }
        const leftOk = Math.max(...items.map((x) => x.x0)) - Math.min(...items.map((x) => x.x0)) <= Math.max(4, rMs * 0.45);
        const rightOk = Math.max(...items.map((x) => x.x1)) - Math.min(...items.map((x) => x.x1)) <= Math.max(4, rMs * 0.45);
        const cxs = items.map((x) => (x.x0 + x.x1) / 2);
        const centerOk = Math.max(...cxs) - Math.min(...cxs) <= Math.max(5, rMs * 0.55);
        if (leftOk || rightOk || centerOk) {
          alignedCols++;
          if (c === 0) col0Aligned = true;
          if (c === 1) col1Aligned = true;
        }
      }
      if (!col0Aligned || !col1Aligned || alignedCols < nC - 1) continue;

      // strażnik przed prozą wielokolumnową (artykuł 2-kolumnowy: zdania zawijane z wiersza na wiersz)
      const cellTexts = grid.map((row) => row.map((cell) => (cell ? spanText(cell) : "")));
      let wrappedProse = 0;
      for (let c = 0; c < nC; c++) {
        const ws = [];
        for (let r = 0; r < nR; r++) if (grid[r][c]) ws.push(grid[r][c].x1 - grid[r][c].x0);
        const maxW = Math.max(1, ...ws);
        for (let r = 0; r + 1 < nR; r++) {
          const a = grid[r][c], b = grid[r + 1][c];
          if (!a || !b) continue;
          const ta = cellTexts[r][c], tb = cellTexts[r + 1][c];
          const wA = a.x1 - a.x0;
          if (wA >= maxW * 0.78 && ta.split(/\s+/).length >= 3 && !/[.:?!;)]$/.test(ta) && /^\p{Ll}/u.test(tb)) wrappedProse++;
        }
      }
      if (wrappedProse >= (nC === 2 ? 1 : 2)) continue;

      // zaostrzony strażnik dla 2 kolumn (rozróżnia zestawienie klucz–wartość od 2-kolumnowego tekstu w proba.pdf)
      if (nC === 2) {
        const wordsCol = (c) => median(cellTexts.map((row) => row[c]).filter(Boolean).map((t) => t.split(/\s+/).length)) || 0;
        const w0 = wordsCol(0), w1 = wordsCol(1);
        if (w0 >= 5 && w1 >= 5) continue;
        const c0List = cellTexts.map((row) => row[0]).filter(Boolean);
        const hasColonKeys = c0List.filter((t) => /:$/.test(t)).length >= Math.ceil(c0List.length * 0.6) && w0 <= 4 && w1 <= 5;
        const rowBoldRatio = (r) => {
          const gs = grid[r].filter(Boolean).flatMap((c) => c.glyphs);
          return gs.length ? gs.filter((g) => g.bold).length / gs.length : 0;
        };
        const bodyBold = median(Array.from({ length: nR - 1 }, (_, k) => rowBoldRatio(k + 1)));
        const hasBoldHeader = rowBoldRatio(0) >= 0.7 && bodyBold <= 0.3 && w0 <= 4 && w1 <= 4;
        // bez nagłówka i dwukropków 2-kolumnowe „towar … cena” (paragon, ocr-layout) zostaje wierszami z tabulatorem
        if (!hasColonKeys && !hasBoldHeader) continue;
        if (nR === 2 && !hasBoldHeader && !hasColonKeys) continue;
      }

      const Y = [lines[0].top - 2];
      for (let r = 1; r < nR; r++) Y.push((lines[r - 1].bottom + lines[r].top) / 2);
      Y.push(lines[nR - 1].bottom + 2);
      const cellOf = Array.from({ length: nR }, () => new Array(nC).fill(null));
      const cells = [];
      for (let r = 0; r < nR; r++) {
        for (let c = 0; c < nC; c++) {
          const cell = {
            r0: r, r1: r, c0: c, c1: c,
            x0: X[c], x1: X[c + 1], y0: Y[r], y1: Y[r + 1],
            fill: null,
            borders: { top: null, bottom: null, left: null, right: null },
            frags: [],
          };
          cells.push(cell);
          cellOf[r][c] = cell;
        }
      }
      out.push({ type: "table", borderless: true, x0: X[0], x1: X[nC], y0: Y[0], y1: Y[nR], top: Y[0], bottom: Y[nR], X, Y, cells, cellOf, _frags: new Set(runFrags) });
    }
    return out;
  }

  // Tabela z Worda bywa narysowana w kawałkach przedzielonych wierszami bez tła i krawędzi
  // (np. „Świadczenia z tytułu…” albo tabela zebrowa z cieniowaniem co drugiego wiersza).
  // Kawałki o tych samych brzegach i z małą przerwą (albo z wierszem tekstu wypełniającym przerwę zebry)
  // składamy z powrotem w jedną tabelę; puste przerwy harmonogramów (pdf-bands) zostają osobno.
  function mergeStackedTables(tables, frags = []) {
    if (tables.length < 2) return tables;
    const sorted = tables.slice().sort((a, b) => a.y0 - b.y0);
    const out = [sorted[0]];
    const zebraMerged = new WeakSet();
    for (let i = 1; i < sorted.length; i++) {
      const a = out[out.length - 1], b = sorted[i];
      const gap = b.y0 - a.y1;
      const sameEdges = Math.abs(a.x0 - b.x0) < 3 && Math.abs(a.x1 - b.x1) < 3;
      const maxRowH = Math.max(a.Y[a.Y.length - 1] - a.Y[a.Y.length - 2], b.Y[1] - b.Y[0]);
      const zebraRow = sameEdges && gap >= 16 && gap <= Math.max(36, maxRowH * 1.8)
        && frags.some((f) => f.type === "text" && !f.rotated && (f.x0 + f.x1) / 2 >= a.x0 - 1 && (f.x0 + f.x1) / 2 <= a.x1 + 1 && f.y > a.y1 + 1 && f.y < b.y0 - 1);
      if (sameEdges && ((gap > -1 && gap < 16) || zebraRow)) {
        const joined = joinTables(a, b);
        if (zebraRow || zebraMerged.has(a)) zebraMerged.add(joined);
        out[out.length - 1] = joined;
      } else out.push(b);
    }
    // Parzysta liczba wierszy w zebrze (ostatni wiersz biały, pod ostatnim cieniowanym pasem):
    // dołączamy ten pojedynczy wiersz, jeśli leży tuż pod tabelą w tym samym rozstawie kolumn.
    for (let i = 0; i < out.length; i++) {
      const t = out[i];
      if (!zebraMerged.has(t)) continue;
      const rowH = median(t.Y.slice(1).map((y, k) => y - t.Y[k])) || 20;
      const nextY = i + 1 < out.length ? out[i + 1].y0 : Infinity;
      const limitY = Math.min(t.y1 + rowH * 1.5, nextY - 2);
      const below = frags.filter((f) => f.type === "text" && !f.rotated && (f.x0 + f.x1) / 2 >= t.x0 - 1 && (f.x0 + f.x1) / 2 <= t.x1 + 1 && f.y > t.y1 + 1 && f.y <= limitY);
      if (!below.length) continue;
      const ms = median(below.map((f) => f.size)) || 10;
      const lines = lineSpansInBox(below, ms);
      if (lines.length === 1 && lines[0].spans.length >= 2) {
        const newY1 = Math.max(lines[0].bottom + 4, t.y1 + rowH);
        const r = t.Y.length - 1;
        const nC = t.X.length - 1;
        const Y = [...t.Y, newY1];
        const cell = { r0: r, r1: r, c0: 0, c1: nC - 1, x0: t.X[0], x1: t.X[nC], y0: t.y1, y1: newY1, fill: null, borders: { top: null, bottom: null, left: null, right: null }, frags: [] };
        const cellOf = [...t.cellOf, new Array(nC).fill(cell)];
        out[i] = { ...t, y1: newY1, bottom: newY1, Y, cells: [...t.cells, cell], cellOf };
      }
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

  function buildGrid(items, frags, tol, opts = {}) {
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
    if (nC * nR === 1 && !inBox.length && !opts.allowEmpty) return null;

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
    // W planach z Excela granice łączonych pól potrafią tworzyć niejednoznaczną, nieprostokątną
    // grupę. Każda pozycja siatki musi jednak dostać komórkę — Word nie ma odpowiednika „dziury”.
    // Bez tego późniejszy zapis DOCX próbował odczytać c1 z null i przerywał całą konwersję.
    for (let r = 0; r < nR; r++) for (let c = 0; c < nC; c++) if (!cellOf[r][c]) {
      const cell = { r0: r, r1: r, c0: c, c1: c };
      cells.push(cell);
      cellOf[r][c] = cell;
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
      if (labelValueRows(left, right, ms)) continue;
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

  // Lista „etykieta … wartość” (paragon: towar … cena, infografika: model … rok): po obu stronach
  // przerwy te same linie bazowe, a jedna strona dużo węższa. To wiersze z tabulatorem, nie dwie
  // kolumny — cięcie dawało najpierw wszystkie nazwy, potem wszystkie ceny. Dwie kolumny
  // artykułu (nawet z liniami bazowymi na jednej siatce) mają podobne szerokości — zostają.
  function labelValueRows(left, right, ms) {
    if (left.some((a) => a.type !== "text") || right.some((a) => a.type !== "text")) return false;
    const tol = Math.max(1, ms * 0.35);
    const rowsOf = (side) => clusterValues(side.map((a) => a.y), tol).map((y) => {
      const on = side.filter((a) => Math.abs(a.y - y) <= tol);
      return { y, w: Math.max(...on.map((a) => a.x1)) - Math.min(...on.map((a) => a.x0)) };
    });
    const L = rowsOf(left), R = rowsOf(right);
    if (L.length < 2 || R.length < 2) return false;
    const paired = (a, b) => a.filter((r) => b.some((o) => Math.abs(o.y - r.y) <= tol)).length / a.length;
    if (paired(L, R) < 0.8 || paired(R, L) < 0.8) return false;
    const wl = median(L.map((r) => r.w)), wr = median(R.map((r) => r.w));
    return Math.min(wl, wr) <= 0.6 * Math.max(wl, wr);
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
      if (last && last.type === "text" && last.key === style.key && last.link === (g.link || null) && last.field === (g.field || null) && !g.fieldMarker && !last.fieldMarker) last.text += text;
      else toks.push({ type: "text", text, ...style, link: g.link || null, field: g.field || null, fieldMarker: !!g.fieldMarker });
    };
    for (let p = 0; p < line.parts.length; p++) {
      const part = line.parts[p];
      let glyphs = part.glyphs;
      if (p > 0) {
        // przerwa między kawałkami = tabulator do pozycji kawałka
        toks.push({ type: "tab", pos: part.x0 - regionX0, leader: null, align: "left", startX: part.x0, endX: part.x1 });
        prev = null;
      }
      // Odstępy między słowami w tym kawałku (od końca litery do początku następnego słowa).
      // Szeroki odstęp (kilka spacji ze źródła — Excel: „SOBOTA      godz.”) staje się tabulatorem
      // na pozycji z PDF (jak większe odstępy niżej) — dawniej zwijał się do jednej spacji i kolumna
      // „godz.” w wierszach się rozjeżdżała. Spacje liczone szerokością kroju chybiały o 2 pt
      // (Excel skaluje stronę, krój zastępczy) — tabulator stawia słowo dokładnie tam, gdzie w PDF.
      // Wyjustowany tekst ma wszystkie odstępy podobnie szerokie — tam nic się nie zmienia.
      const wordGaps = new Map(); // glif → odstęp przed nim (pt)
      {
        let ink = null, sawSpace = false;
        for (const g of glyphs) {
          if (g.isSpace) { sawSpace = true; continue; }
          if (ink) {
            const gap = g.x - ink.x1;
            if (sawSpace || gap > Math.max(ink.size, g.size) * 0.17) wordGaps.set(g, gap);
          }
          ink = g;
          sawSpace = false;
        }
      }
      // wzór (indeksy w mniejszym kroju, znaki działań) — szerokie odstępy to układ wzoru, nie spacje
      const inkSizes = glyphs.filter((x) => !x.isSpace).map((x) => x.size || 0);
      const formula = Math.min(...inkSizes) < Math.max(...inkSizes) * 0.85 || glyphs.some((x) => /[=+<>±×÷∑∏∫√∞→←↔≤≥≠≈∂∆∇∈∉⊂⊃∪∩^_|]/.test(x.u || ""));
      const wideGap = (g) => {
        const gap = wordGaps.get(g);
        const size = Math.max(ink?.size || 0, g.size || 10);
        if (gap == null || gap < size * 0.9) return false;
        // Bardzo szeroka przerwa (> 1,6 em) złożona ze znaków spacji (np. „KMgF3        fluorek…”):
        // wcześniej wpadała w lukę między progiem !prev.isSpace a górnym limitem wideGap i zwijała się do 1 spacji.
        if (gap > size * 1.6) {
          const others = [...wordGaps.values()].filter((v) => v !== gap).sort((a, b) => a - b);
          return !(others.length >= 2 && gap < others[Math.floor(others.length / 2)] * 2.2);
        }
        if (formula) return false;
        const others = [...wordGaps.values()].filter((v) => v !== gap && v <= size * 1.6).sort((a, b) => a - b);
        return !(others.length && gap < others[Math.floor(others.length / 2)] * 2.5);
      };
      let ink = null; // ostatnia litera (nie spacja) — szeroki odstęp liczony od niej
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
            // Linia do wpisywania („recover ________ energy”): spacje przed i za linią zostają, a linia
            // kończy się tam, gdzie podkreślniki — dawniej sklejała się ze słowami po obu stronach.
            const blank = leader === "underscore" && !isPageNo;
            const lastTok = toks[toks.length - 1];
            // spacja w PDF-ie bywa znakiem albo samą przerwą między literami — liczymy obie
            const gap = (a, b2) => a && b2 && b2.x - a.x1 > g.size * 0.15;
            const before = glyphs[i - 1], last = run[run.length - 1];
            if (blank && before && (before.isSpace || gap(before, run[0])) && lastTok?.type === "text" && !/\s$/.test(lastTok.text)) lastTok.text += " ";
            const spaceAfter = blank && next && (glyphs[j - 1].isSpace || gap(last, next));
            if (!next) {
              toks.push({ type: "tab", pos: endX - regionX0, leader, align: "left", startX: run[0].x, endX, fill: true, trailing: true });
            } else if (spaceAfter) {
              const lineEnd = run[run.length - 1].x1;
              toks.push({ type: "tab", pos: lineEnd - regionX0, leader, align: "left", leaderTo: lineEnd, startX: run[0].x, fill: true, keepSpaces: true });
              pushText(glyphs[j - 1].isSpace ? glyphs[j - 1] : next, " ");
            } else {
              toks.push({ type: "tab", pos: endX - regionX0, leader, align: isPageNo ? "right" : "left", leaderTo: endX, startX: run[0].x, fill: true });
            }
            prev = null;
            i = j;
            continue;
          }
        }
        if (g.fieldMarker) { // puste pole: kontrolka między kropkami (kropki przed i za zostają)
          pushText(g, "");
          prev = null;
          i++;
          continue;
        }
        if (prev) {
          const gap = g.x - prev.x1;
          const sp = Math.max(prev.size, g.size) * 0.17;
          if (!g.isSpace && !prev.isSpace && gap > sp) {
            if (gap > Math.max(prev.size, g.size) * 1.6) {
              toks.push({ type: "tab", pos: g.x - regionX0, leader: null, align: "left", startX: g.x });
            } else pushText(g.field && prev.field !== g.field ? prev : g, " "); // odstęp przed polem — poza polem
          }
        }
        if (g.isSpace) {
          // spacja tylko jedna (wyjustowany tekst ma szerokie spacje, ale to dalej jedna spacja)
          const last = toks[toks.length - 1];
          if (prev && !(last && last.type === "text" && last.text.endsWith(" "))) pushText(g, " ");
        } else {
          if (ink && wideGap(g) && toks[toks.length - 1]?.type !== "tab") toks.push({ type: "tab", pos: g.x - regionX0, leader: null, align: "left", startX: g.x });
          pushText(g, g.u);
          ink = g;
        }
        prev = g;
        i++;
      }
    }
    // spacje na brzegach tekstu przy tabulatorze
    for (let k = 0; k < toks.length; k++) {
      if (toks[k].type !== "tab" || toks[k].keepSpaces) continue; // linia do wpisywania: spacje to część zdania
      if (k > 0 && toks[k - 1].type === "text") toks[k - 1].text = toks[k - 1].text.replace(/ +$/, "");
      if (k + 1 < toks.length && toks[k + 1].type === "text") toks[k + 1].text = toks[k + 1].text.replace(/^ +/, "");
    }
    return toks.filter((t) => t.type !== "text" || t.text.length || t.fieldMarker); // puste pole formularza zostaje
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
      hl: !!g.doubt, // słowo z OCR o mniejszej pewności — żółte „do sprawdzenia”
      shade: g.shade || null, // zakreślenie z PDF-a (applyHighlights)
    };
    if (s.vert) s.size = Math.round(size * 2 / 1.0) / 2; // indeks: zachowaj prawdziwy rozmiar
    s.key = [s.font, s.size, s.bold, s.italic, s.color, s.vert, s.underline, s.strike, s.hScale, s.spacing, s.hl, s.shade].join("|");
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
    // „§” z liczbą („§ 1.”, „§ 12a”) to numer paragrafu, nie punktor (kwadracik Wingdings bywa „§”)
    if (word === "§" && /\d/.test(after.u)) return null;
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
    // odstęp po numerze/punktorze listy („2.   Przejść…”) to nie tabulator układu — zawinięty punkt
    // listy (w komórce tabeli z Worda) ma być jednym akapitem, nie dwoma
    const markerTab = (t) => prev.list && Math.abs(box.x0 + t.pos - prev.list.textX) < 1;
    if (prev.tokens.some((t) => t.type === "tab" && !markerTab(t)) || l.tokens.some((t) => t.type === "tab")) return false;
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
    // dwa bardzo krótkie wiersze w szerokiej komórce/obszarze („Jan Kowalski” / „ul. Długa 55” zajmujące
    // < 55% szerokości komórki, gdzie pierwsze słowo kolejnego wiersza spokojnie mieści się na poprzednim)
    // to osobne wiersze, nie zawinięty akapit z wcięciem z prawej
    const bc0 = (box.x0 + box.x1) / 2, ctol0 = Math.max(3, size * 0.4);
    const looksCentered = (isCenteredLine(prev, box) && isCenteredLine(l, box)) || (Math.abs((prev.x0 + prev.x1) / 2 - bc0) < ctol0 && Math.abs((l.x0 + l.x1) / 2 - bc0) < ctol0);
    if (!looksCentered && para.lines.length === 1 && prev.x1 - box.x0 < width * 0.55 && prev.x1 + size * 0.3 + fw < box.x1 - size * 1.5) return false;
    // lewa krawędź: równa z poprzednim (lub z wcięciem pierwszego wiersza / wiszącym)
    const first = para.lines[0];
    const leftRef = para.lines.length >= 2 ? prev.x0 : first.list ? first.list.textX : first.x0;
    const dx = Math.abs(l.x0 - leftRef);
    // wyśrodkowane: oba wiersze na środku obszaru. Obszar bywa wyznaczony przez sam akapit (jego
    // najszerszy wiersz = krawędzie obszaru), więc wiersz bez odsunięcia od lewej też się liczy,
    // gdy środki się zgadzają, a lewe krawędzie nie (lista 3.11: inaczej akapit rozpadał się na wiersze)
    const bc = (box.x0 + box.x1) / 2, ctol = Math.max(3, size * 0.4);
    const nearC = (x) => Math.abs((x.x0 + x.x1) / 2 - bc) < ctol;
    // obie krawędzie inne (symetrycznie): wyjustowane wiersze z wysunięciem pierwszego mają wspólną
    // PRAWĄ krawędź i środek też blisko środka obszaru — to nie wyśrodkowanie (pdf:score „CO-23”)
    const centered = (isCenteredLine(prev, box) && isCenteredLine(l, box)) || (nearC(prev) && nearC(l) && dx > size * 0.6 && Math.abs(l.x1 - prev.x1) > size * 0.6);
    // wyśrodkowane zawija się na szerokości obszaru: krótki wiersz, w którym zmieściłoby się
    // pierwsze słowo następnego, kończy się ręcznie (tytuł „Zarządzenie Nr …” / „Rektora …” / „z dnia …”)
    // Gdy akapit ma już ≥ 2 wiersze, zawija się na WŁASNEJ szerokości (wcięcia po bokach, węższy
    // blok) — nie mniejszej niż jego najszerszy wiersz. Porównanie z całym obszarem odcinało
    // kolejne, nieco krótsze wiersze wyśrodkowanego akapitu (lista 3.11; PDF z Worda, wcięcia 3 cm).
    const wrapW = para.lines.length >= 2 ? Math.min(width, Math.max(...para.lines.map((x) => x.x1 - x.x0))) : width;
    if (centered && prev.x1 - prev.x0 + size * 0.3 + fw < wrapW - size * 0.5) return false;
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

  // Szerokość wiersza ze spacjami w ich zwykłej szerokości (bez rozpychania/ściskania z PDF):
  // znaki swoją szerokością, spacja swoją, przerwa bez znaku spacji jak spacja kroju.
  function naturalLineWidth(l) {
    const gl = l.parts.flatMap((p) => p.glyphs || []).filter((g) => !g.invisible).sort((a, b) => a.x - b.x);
    while (gl.length && gl[0].isSpace) gl.shift();
    while (gl.length && gl[gl.length - 1].isSpace) gl.pop();
    if (!gl.length) return l.x1 - l.x0;
    let w = 0, prev = null, lastInk = null, spaced = 0;
    // rozmiar w pliku idzie co pół punktu (9,96 pt z PDF → 10 pt): litery tyle szersze. W dół nie
    // zwężamy — zapas zostaje (podgląd bywa minimalnie szerszy niż szerokości z PDF). Tylko krój
    // z PDF = krój w pliku (zastępnik i krój Ghostscripta mają inne szerokości niż Arial/Times).
    const k = (g) => (!g.size || g.font?.substituted || g.font?.ghostscript ? 1 : Math.max(1, Math.round(g.size * 2) / 2 / g.size));
    for (const g of gl) {
      if (g.isSpace) {
        // spacja tak, jak ją złożymy (pdf-docx: zagęszczenie z fitWordSpaces)
        const sw = g.adv * k(g) + (g.fitSpacing || 0) / 20;
        w += sw;
        spaced += sw;
      } else {
        if (prev && !prev.isSpace && g.x - prev.x1 > g.size * 0.15) {
          const sw = g.size * (g.font?.serif ? 0.25 : 0.278);
          w += sw;
          spaced += sw;
        }
        // przerwa szersza niż kilka spacji (punktor listy → tekst, tabulator) zostaje jak w PDF —
        // liczona jako jedna spacja zaniżała wiersz o ~20 pt i zapas na zawijanie wychodził za mały
        if (lastInk && g.x - lastInk.x1 > g.size * 1.2) w += g.x - lastInk.x1 - spaced;
        w += (g.x1 - g.x) * k(g);
        lastInk = g;
        spaced = 0;
      }
      prev = g;
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
      if (align === "both") {
        // Wyjustowany wiersz w PDF bywa WĘŻSZY niż jego tekst ze zwykłymi spacjami: Word 2013+
        // ściska spacje, zanim przeniesie słowo („CO-23”: „…niebezpiecznym zacho-” 484,9 pt tekstu
        // w 481,6 pt). Szerokość jak w PDF nie wystarcza wtedy ani podglądowi, ani Wordowi (próg
        // ściskania zależy od kroju i rozmiaru — wyrocznia 2026-10-28) — słowo zjeżdżało i cała strona
        // za nim. Szerokość zawijania ≥ najdłuższy wiersz w naturalnej szerokości, ale nie tyle,
        // żeby wciągnąć pierwsze słowo następnego wiersza.
        let right = maxR + 0.3;
        const nat = lines.map((l) => l.x0 + naturalLineWidth(l));
        const want = Math.max(...nat.slice(0, -1)) + 0.3;
        if (want > right) {
          let cap = Infinity;
          for (let i = 0; i + 1 < lines.length; i++) cap = Math.min(cap, nat[i] + size * 0.25 + firstWordWidth(lines[i + 1]) - 0.5);
          right = Math.max(right, Math.min(want, cap));
        }
        indRight = right <= box.x1 ? box.x1 - right : Math.max(box.x1 - right, -Math.min(28, size * 2));
      }
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
  // opts.formFields: znaki wartości pola dostają g.field = { id, name, multi } — pdf-docx.js robi z nich
  // kontrolkę tekstową Worda (w:sdt). Puste pole = znacznik zerowej szerokości tam, gdzie pole się zaczyna.
  function annotationFrags(raw, opts = {}) {
    const frags = [];
    for (const a of raw.annotations) {
      if (a.type !== "text" || a.hidden) continue;
      const field = opts.formFields && a.field && !a.readOnly ? { id: a.id, name: a.fieldName || "", multi: !!a.multiLine } : null;
      if (!String(a.text).trim()) {
        if (field && !a.multiLine) {
          const size = a.size || Math.max(4, Math.min(12, (a.y1 - a.y0 - 2) / 1.15));
          const y = (a.y0 + a.y1) / 2 + size * 0.3;
          const x = a.x0 + 2;
          const g = { u: "", x, x1: x, adv: 0, y, size, asc: size * 0.72, desc: size * 0.21, angle: 0, font: { id: "annot", family: "Arial", bold: false, italic: false, mono: false, serif: false, symbolic: false, asc: 0.72, desc: 0.21 }, isSpace: false, color: "000000", bold: false, italic: false, order: 1e9, fromField: true, field, fieldMarker: true };
          const f = finishFrag({ glyphs: [g], x0: x, x1: x, size });
          f.fromField = true;
          f.fieldMarker = true;
          frags.push(f);
        }
        continue;
      }
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
          glyphs.push({ u: ch, x, x1: x + w, adv: w, y, size, asc: size * 0.72, desc: size * 0.21, angle: 0, font, isSpace: ch === " ", color, bold: font.bold, italic: font.italic, order: 1e9, fromField: true, field });
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
        if (!ff.fieldMarker) rest.push(ff); // puste pole poza wierszem z kropkami — bez kontrolki
        continue;
      }
      const covered = (g) => !ff.fieldMarker && (g.u === "." || g.u === "_" || g.u === "\u2026" || g.isSpace) && g.x1 > ff.x0 - ff.size * 0.3 && g.x < ff.x1 + ff.size * 0.3;
      const kept = host.glyphs.filter((g) => !covered(g));
      for (const g of ff.glyphs) g.y = host.y; // jedna linia bazowa
      const merged = kept.concat(ff.glyphs).sort((p, q) => p.x - q.x);
      const nf = finishFrag({ glyphs: merged, x0: merged[0].x, x1: Math.max(...merged.map((g) => g.x1)), size: Math.max(host.size, ff.size) });
      frags[frags.indexOf(host)] = nf;
    }
    return frags.concat(rest);
  }

  // Linia „pasuje do tekstu”: zaczyna się i kończy przy literach nad nią (najwyżej ~1 em dalej)
  // i nie ma pod sobą dużej dziury bez liter. Kreska pod tytułem ramki, krawędź wiersza tabeli
  // czy separator „Model … 2001” są dłuższe od tekstu albo mają w środku pustkę — to nie
  // podkreślenie (infografiki i skany: dziesiątki słów podkreślonych bez powodu).
  function lineFitsText(s, hits) {
    const gl = hits.filter((g) => !g.isSpace && g.u.trim()).sort((p, q) => p.x - q.x);
    if (!gl.length) return false;
    const em = Math.max(...gl.map((g) => g.size || 10));
    if (gl[0].x - s.x0 > em || s.x1 - gl[gl.length - 1].x1 > em) return false;
    let end = Math.max(s.x0, gl[0].x);
    for (const g of gl) {
      if (g.x - end > em * 1.5) return false;
      end = Math.max(end, g.x1);
    }
    return true;
  }

  function applyUnderlines(frags, hSegs) {
    // cienka pozioma linia tuż pod linią bazową, pod tekstem = podkreślenie (nie tabela)
    const used = new Set();
    // litery wszystkich fragmentów nad linią (wiersz bywa podzielony na kilka fragmentów)
    const under = (s, lo, hi, below) => {
      const hits = [];
      for (const f of frags) {
        if (f.type !== "text") continue;
        const dy = below ? s.y - f.y : f.y - s.y;
        if (dy < f.size * lo || dy > f.size * hi) continue;
        if (overlap1d(s.x0, s.x1, f.x0, f.x1) < 2) continue;
        for (const g of f.glyphs) {
          const cx = (g.x + g.x1) / 2;
          if (cx >= s.x0 - 0.5 && cx <= s.x1 + 0.5) hits.push(g);
        }
      }
      return hits;
    };
    for (const s of hSegs) {
      if (s.w > 2.2) continue;
      const hits = under(s, -0.05, 0.35, true);
      if (!lineFitsText(s, hits)) continue;
      for (const g of hits) g.underline = true;
      used.add(s);
    }
    // Skan: podkreślenie przerwane pod ogonkiem („serię”: przeglądarka i drukarki omijają „ę”)
    // nie obejmuje ostatniej litery — słowo podkreślone w większości jest podkreślone całe.
    for (const f of frags) {
      if (f.type !== "text" || !f.glyphs.some((g) => g.ocr && g.underline)) continue;
      let word = [];
      const flush = () => {
        if (word.length && word.filter((g) => g.underline).length >= word.length * 0.6) for (const g of word) g.underline = true;
        word = [];
      };
      for (const g of f.glyphs) {
        if (g.isSpace) flush();
        else word.push(g);
      }
      flush();
    }
    // przekreślenie: linia w okolicy połowy wysokości x
    for (const s of hSegs) {
      if (used.has(s) || s.w > 2.2) continue;
      const hits = under(s, 0.15, 0.4, false);
      if (!lineFitsText(s, hits)) continue;
      for (const g of hits) g.strike = true;
      used.add(s);
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
    frags = mergeFieldFrags(frags, annotationFrags(raw, opts));
    const bulletShapes = shapeBullets(raw, frags);
    applyLinks(frags, raw.annotations);
    raw = applyHighlights(raw, frags);
    const { tables, leftovers } = detectTables(raw, frags, W, H);
    const rules = applyUnderlines(frags, leftovers.hSegs);

    // przypisanie tekstu do komórek
    const assignFragToTable = (f, t) => {
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
    };
    let free = [];
    for (const f of frags) {
      const cx = (f.x0 + f.x1) / 2, cy = (f.top + f.bottom) / 2;
      const t = tables.find((t) => cx >= t.x0 - 1 && cx <= t.x1 + 1 && cy >= t.y0 - 1 && cy <= t.y1 + 1);
      if (!t) {
        free.push(f);
        continue;
      }
      assignFragToTable(f, t);
    }

    // Etap 2: tabele bez linii i bez cieniowania w wolnym tekście strony
    const bTables = detectBorderlessTables(free, tables, raw.images, W, H);
    if (bTables.length) {
      const taken = new Set();
      for (const bt of bTables) {
        for (const f of bt._frags) {
          taken.add(f);
          assignFragToTable(f, bt);
        }
        delete bt._frags;
        tables.push(bt);
      }
      free = free.filter((f) => !taken.has(f));
    }

    // Plan zajęć z Excela ma bardzo wąską kolumnę z pionowymi nazwami dni. Czasem PDF nie
    // rysuje poziomej kreski przez jej całą szerokość, więc wykrywacz uznaje dwa dni za jedną
    // pionowo scalaną komórkę. Gdy w takiej komórce leżą osobne obrócone podpisy, rozdzielamy
    // ją na najbliższych istniejących wierszach siatki. Dzięki temu „czwartek” i „piątek” nie
    // trafiają do jednego pola Worda.
    for (const table of tables) {
      for (const cell of table.cells.slice()) {
        const labels = (cell.frags || []).filter((fragment) => fragment.type === "text" && fragment.rotated)
          .sort((a, b) => a.top - b.top);
        if (labels.length < 2 || cell.x1 - cell.x0 > 24 || cell.r1 <= cell.r0) continue;
        const cuts = [];
        for (let i = 0; i + 1 < labels.length; i++) {
          const middle = (labels[i].bottom + labels[i + 1].top) / 2;
          let best = null;
          for (let row = cell.r0 + 1; row <= cell.r1; row++) {
            const distance = Math.abs(table.Y[row] - middle);
            if (!best || distance < best.distance) best = { row, distance };
          }
          if (best && !cuts.includes(best.row)) cuts.push(best.row);
        }
        if (!cuts.length) continue;
        cuts.sort((a, b) => a - b);
        const bounds = [cell.r0, ...cuts, cell.r1 + 1];
        const pieces = [];
        for (let i = 0; i + 1 < bounds.length; i++) {
          const r0 = bounds[i], r1 = bounds[i + 1] - 1;
          const part = {
            ...cell,
            r0, r1,
            y0: table.Y[r0], y1: table.Y[r1 + 1],
            frags: [],
          };
          pieces.push(part);
          for (let row = r0; row <= r1; row++) for (let col = cell.c0; col <= cell.c1; col++) table.cellOf[row][col] = part;
        }
        for (const fragment of cell.frags) {
          const middle = (fragment.top + fragment.bottom) / 2;
          const target = pieces.find((part) => middle >= part.y0 - 1 && middle <= part.y1 + 1)
            || pieces.reduce((nearest, part) => Math.abs((part.y0 + part.y1) / 2 - middle) < Math.abs((nearest.y0 + nearest.y1) / 2 - middle) ? part : nearest);
          target.frags.push(fragment);
        }
        const at = table.cells.indexOf(cell);
        table.cells.splice(at, 1, ...pieces);
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
      // nachodzenie tylko na brzeg obrazu (≤ IMG_INSET) to wciąż „obok” — obraz zostaje w tekście
      const ix0 = b.x1 - b.x0 > IMG_INSET * 4 ? b.x0 + IMG_INSET : b.x0, ix1 = b.x1 - b.x0 > IMG_INSET * 4 ? b.x1 - IMG_INSET : b.x1;
      const cell = tables.flatMap((t) => t.cells).find((c) => (b.x0 + b.x1) / 2 >= c.x0 && (b.x0 + b.x1) / 2 <= c.x1 && (b.y0 + b.y1) / 2 >= c.y0 && (b.y0 + b.y1) / 2 <= c.y1);
      if (cell) {
        // Obraz pokrywający tekst wewnątrz komórki (np. tło ramki kodu) nie może być obrazem w tekście
        // — inaczej spycha całą treść komórki o swoją wysokość w dół; idzie do floats pod/nad tekstem.
        const cellText = cell.frags.filter((f) => f.type === "text" && overlap1d(f.x0, f.x1, ix0, ix1) > 0.3 && overlap1d(f.top, f.bottom, b.y0, b.y1) > 2);
        if (cellText.length) {
          const firstCellText = Math.min(...cellText.flatMap((f) => f.glyphs.map((g) => g.order)));
          floats.push({ ...atom, behind: im.order < firstCellText });
        } else cell.frags.push(atom);
        continue;
      }
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
  // Strefa nagłówka/stopki: 10% od krawędzi zawsze; dalej (do 20% od góry, 15% od dołu — nagłówek
  // firmowy z logo, adresem i hasłem) tylko wiersze, nad którymi (pod którymi) wszystko też się
  // powtarza — pdf-convert sprawdza tę ciągłość od krawędzi strony.
  function runningZone(y, H) {
    if (y < H * 0.2) return y < H * 0.1 ? "top" : "top+";
    if (y > H * 0.85) return y > H * 0.9 ? "bottom" : "bottom+";
    return null;
  }

  function detectRunningText(pagesRaw) {
    const sigs = new Map();
    const n = pagesRaw.length;
    if (n < 2) return new Set();
    pagesRaw.forEach((raw, pi) => {
      const H = raw.height;
      const rows = new Map();
      for (const g of raw.glyphs) {
        if (g.invisible) continue;
        const zone = runningZone(g.y, H)?.replace("+", "");
        if (!zone) continue;
        const k = zone + ":" + Math.round(g.y);
        if (!rows.has(k)) rows.set(k, []);
        rows.get(k).push(g);
      }
      for (const [k, gl] of rows) {
        const exact = gl.sort((a, b) => a.x - b.x).map((g) => g.u).join("").replace(/\s+/g, " ").trim();
        if (!exact) continue;
        // cyfry jako „#” — numer strony; dokładny tekst („=…”) dla strefy poszerzonej, gdzie
        // „Rozdział 1”, „Rozdział 2”… na górze kolejnych stron to tytuły, nie nagłówek
        for (const sig of [k.split(":")[0] + "|" + exact.replace(/\d+/g, "#"), "=" + k.split(":")[0] + "|" + exact]) {
          if (!sigs.has(sig)) sigs.set(sig, new Set());
          sigs.get(sig).add(pi);
        }
      }
    });
    const running = new Set();
    for (const [sig, pages] of sigs) if (pages.size >= Math.max(2, Math.ceil(n * 0.5))) running.add(sig);
    return running;
  }

  NS.layoutPage = layoutPage;
  NS.treeToBlocks = treeToBlocks;
  NS.detectRunningText = detectRunningText;
  NS.runningZone = runningZone;
  NS.mapSymbolChar = mapSymbolChar;
  NS._internals = { prepareGlyphs, buildFragments, detectTables, xyCut, flowBlocks, linesFromFrags, lineTokens, finishParagraph, annotationFrags, median };
})(typeof globalThis !== "undefined" ? globalThis : window);
