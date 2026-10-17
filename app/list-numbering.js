// list-numbering.js — numery list liczone jak w Wordzie.
//
// docx-preview rysuje numery licznikami CSS, a te różnią się od Worda (porównanie z prawdziwym
// Wordem 16.114, tools/word-oracle — 17 z 26 przypadków inaczej, 2026-10-17):
//   • „%1.%2” brało format BIEŻĄCEGO poziomu dla wszystkich poziomów („A.i” → „i.i”),
//   • litery po „z”: Word „aa, bb, cc”, CSS „aa, ab, ac”; zero literą/rzymską: Word pusto,
//   • słownie / porządkowo (cardinalText, ordinalText, ordinal) — CSS zwykłe cyfry,
//   • isLgl (numeracja prawnicza), lvlRestart (kiedy poziom liczy od nowa), pominięty poziom,
//     pusty tekst numeru, wspólny licznik list z jednej definicji po startOverride.
// Tu: te same reguły co w Wordzie, po kolei po akapitach podglądu (także po Enter/Tab w liście —
// obserwator klas), wynik w data-dwb-num, który CSS wstawia zamiast licznika. Punktory (bullet)
// i formaty, których nie znamy, zostają przy rysunku docx-preview.

const dwbListNum = (() => {
  let model = null; // { abs: Map, nums: Map, lang }
  let modelBytes = null;
  let loading = null;

  // ── definicje z numbering.xml ──────────────────────────────────────────────
  const kids = (el, name) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
  const val = (el) => (el ? el.getAttributeNS(W_NS, "val") ?? el.getAttribute("w:val") : null);
  function parseLvl(el) {
    const one = (n) => kids(el, n)[0];
    const start = parseInt(val(one("start")), 10);
    const restartEl = one("lvlRestart");
    const lgl = one("isLgl");
    return {
      fmt: val(one("numFmt")) || "decimal",
      text: one("lvlText") ? (val(one("lvlText")) ?? "") : null,
      start: Number.isFinite(start) ? start : 0, // Word: brak w:start = 0
      restart: restartEl ? parseInt(val(restartEl), 10) : null,
      isLgl: !!lgl && val(lgl) !== "0" && val(lgl) !== "false",
      suff: val(one("suff")) || "tab",
    };
  }

  async function loadModel(bytes) {
    const zip = await loadDocxZipCached(bytes);
    const read = async (path) => {
      const s = await zip.file(path)?.async("string");
      return s ? new DOMParser().parseFromString(s, "application/xml") : null;
    };
    const [numbering, styles, docXml] = await Promise.all([read("word/numbering.xml"), read("word/styles.xml"), read("word/document.xml")]);
    const abs = new Map();
    const nums = new Map();
    const root = numbering?.documentElement;
    kids(root, "abstractNum").forEach((a) => {
      const levels = [];
      kids(a, "lvl").forEach((l) => { levels[parseInt(l.getAttributeNS(W_NS, "ilvl"), 10) || 0] = parseLvl(l); });
      // w:numStyleLink — definicja w innej (stylu listy); nie śledzimy: zostaje rysunek CSS
      abs.set(a.getAttributeNS(W_NS, "abstractNumId"), { levels, link: !!kids(a, "numStyleLink")[0] });
    });
    kids(root, "num").forEach((n) => {
      const overrides = new Map();
      kids(n, "lvlOverride").forEach((o) => {
        const so = kids(o, "startOverride")[0];
        const lv = kids(o, "lvl")[0];
        overrides.set(parseInt(o.getAttributeNS(W_NS, "ilvl"), 10) || 0, {
          start: so ? parseInt(val(so), 10) : null,
          lvl: lv ? parseLvl(lv) : null,
        });
      });
      nums.set(n.getAttributeNS(W_NS, "numId"), { abs: val(kids(n, "abstractNumId")[0]), overrides });
    });
    const defLang = styles ? val(styles.getElementsByTagNameNS(W_NS, "docDefaults")[0]?.getElementsByTagNameNS(W_NS, "lang")[0]) : null;
    const langOf = (el) => val(kids(kids(el, "rPr")[0], "lang")[0]);
    const paraLangs = docXml ? collectParagraphElements(docXml, "body").map((p) => {
      const ppr = kids(p, "pPr")[0];
      if (!ppr || !kids(ppr, "numPr")[0] && !kids(ppr, "pStyle")[0]) return null;
      return langOf(ppr) || langOf(p.getElementsByTagNameNS(W_NS, "r")[0]) || null;
    }) : [];
    return { abs, nums, paraLangs, lang: defLang || (typeof currentLang !== "undefined" && currentLang === "en" ? "en-US" : "pl-PL") };
  }

  function levelDef(numId, ilvl) {
    const num = model.nums.get(numId);
    if (!num) return null;
    const ov = num.overrides.get(ilvl);
    if (ov?.lvl) return ov.lvl;
    const a = model.abs.get(num.abs);
    return a && !a.link ? a.levels[ilvl] || null : null;
  }

  // ── wartości w formatach Worda ─────────────────────────────────────────────
  const roman = (n) => {
    let s = "";
    [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]].forEach(([v, r]) => { while (n >= v) { s += r; n -= v; } });
    return s;
  };
  // Word: a…z, aa…zz, aaa… — ta sama litera powtórzona; po 30 powtórzeniach (780) od nowa
  const letters = (n) => {
    const k = (n - 1) % 780;
    return String.fromCharCode(97 + (k % 26)).repeat(Math.floor(k / 26) + 1);
  };
  const cap = (s) => (s ? s[0].toLocaleUpperCase() + s.slice(1) : s);

  // polski: słownie (cardinalText) i porządkowo (ordinalText) — formy z Worda
  const PL = {
    ones: ["zero", "jeden", "dwa", "trzy", "cztery", "pięć", "sześć", "siedem", "osiem", "dziewięć", "dziesięć", "jedenaście", "dwanaście", "trzynaście", "czternaście", "piętnaście", "szesnaście", "siedemnaście", "osiemnaście", "dziewiętnaście"],
    tens: ["", "", "dwadzieścia", "trzydzieści", "czterdzieści", "pięćdziesiąt", "sześćdziesiąt", "siedemdziesiąt", "osiemdziesiąt", "dziewięćdziesiąt"],
    hundreds: ["", "sto", "dwieście", "trzysta", "czterysta", "pięćset", "sześćset", "siedemset", "osiemset", "dziewięćset"],
    oOnes: ["zerowy", "pierwszy", "drugi", "trzeci", "czwarty", "piąty", "szósty", "siódmy", "ósmy", "dziewiąty", "dziesiąty", "jedenasty", "dwunasty", "trzynasty", "czternasty", "piętnasty", "szesnasty", "siedemnasty", "osiemnasty", "dziewiętnasty"],
    oTens: ["", "", "dwudziesty", "trzydziesty", "czterdziesty", "pięćdziesiąty", "sześćdziesiąty", "siedemdziesiąty", "osiemdziesiąty", "dziewięćdziesiąty"],
    oHundreds: ["", "setny", "dwusetny", "trzysetny", "czterysetny", "pięćsetny", "sześćsetny", "siedemsetny", "osiemsetny", "dziewięćsetny"],
    // człon złożenia („dwu|tysięczny”, „dwudziesto|pięcio…”). Word 16.114 psuje tu ogonki
    // („Piąciotysięczny”, „Szeąciotysięczny”, „Óąmiotysięczny”) — piszemy poprawnie.
    cOnes: ["", "jedno", "dwu", "trzy", "cztero", "pięcio", "sześcio", "siedmio", "ośmio", "dziewięcio", "dziesięcio", "jedenasto", "dwunasto", "trzynasto", "czternasto", "piętnasto", "szesnasto", "siedemnasto", "osiemnasto", "dziewiętnasto"],
    cTens: ["", "", "dwudziesto", "trzydziesto", "czterdziesto", "pięćdziesięcio", "sześćdziesięcio", "siedemdziesięcio", "osiemdziesięcio", "dziewięćdziesięcio"],
    cHundreds: ["", "stu", "dwustu", "trzystu", "czterystu", "pięćset", "sześćset", "siedemset", "osiemset", "dziewięćset"],
  };
  function plBelow1000(n) {
    const out = [];
    if (n >= 100) out.push(PL.hundreds[Math.floor(n / 100)]);
    const r = n % 100;
    if (r >= 20) { out.push(PL.tens[Math.floor(r / 10)]); if (r % 10) out.push(PL.ones[r % 10]); } else if (r || !out.length) out.push(PL.ones[r]);
    return out.join(" ");
  }
  function plThousands(t) {
    if (t === 1) return "tysiąc";
    const u = t % 10;
    const tt = t % 100;
    const form = u >= 2 && u <= 4 && !(tt >= 12 && tt <= 14) ? "tysiące" : "tysięcy";
    return `${plBelow1000(t)} ${form}`;
  }
  function plCardinal(n) {
    if (n < 1000) return plBelow1000(n);
    const t = Math.floor(n / 1000);
    const r = n % 1000;
    return r ? `${plThousands(t)} ${plBelow1000(r)}` : plThousands(t);
  }
  function plOrdinalBelow1000(n) {
    const h = Math.floor(n / 100);
    const r = n % 100;
    if (!r && h) return PL.oHundreds[h];
    const out = h ? [PL.hundreds[h]] : [];
    if (r >= 20) { out.push(PL.oTens[Math.floor(r / 10)]); if (r % 10) out.push(PL.oOnes[r % 10]); } else out.push(PL.oOnes[r]);
    return out.join(" ");
  }
  function plCompound(t) { // 1…999 jako człon złożenia
    const h = Math.floor(t / 100);
    const r = t % 100;
    let s = h ? PL.cHundreds[h] : "";
    if (r >= 20) s += PL.cTens[Math.floor(r / 10)] + (r % 10 ? PL.cOnes[r % 10] : "");
    else if (r) s += PL.cOnes[r];
    return s;
  }
  function plOrdinal(n) {
    if (n < 1000) return plOrdinalBelow1000(n);
    const t = Math.floor(n / 1000);
    const r = n % 1000;
    if (r) return `${plThousands(t)} ${plOrdinalBelow1000(r)}`;
    return t === 1 ? "tysiączny" : `${plCompound(t)}tysięczny`;
  }

  const EN = {
    ones: ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"],
    tens: ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"],
  };
  function enBelow1000(n) {
    const out = [];
    if (n >= 100) out.push(`${EN.ones[Math.floor(n / 100)]} hundred`);
    const r = n % 100;
    if (r >= 20) out.push(EN.tens[Math.floor(r / 10)] + (r % 10 ? `-${EN.ones[r % 10]}` : ""));
    else if (r || !out.length) out.push(EN.ones[r]);
    return out.join(" ");
  }
  function enCardinal(n) {
    if (n < 1000) return enBelow1000(n);
    const r = n % 1000;
    return `${enBelow1000(Math.floor(n / 1000))} thousand${r ? ` ${enBelow1000(r)}` : ""}`;
  }
  const EN_ORD = { one: "first", two: "second", three: "third", five: "fifth", eight: "eighth", nine: "ninth", twelve: "twelfth" };
  function enOrdinal(n) {
    const c = enCardinal(n);
    return c.replace(/([a-z]+)$/, (w) => EN_ORD[w] || (w.endsWith("y") ? `${w.slice(0, -1)}ieth` : `${w}th`));
  }
  const enSuffix = (n) => (n % 100 >= 11 && n % 100 <= 13 ? "th" : ["th", "st", "nd", "rd"][n % 10] || "th");

  // null = format nieobsługiwany (zostaje licznik CSS)
  function formatValue(n, fmt, lang) {
    const en = /^en/i.test(lang || "");
    switch (fmt) {
      case "decimal": return String(n);
      case "decimalZero": return n < 10 && n >= 0 ? `0${n}` : String(n);
      case "lowerLetter": return n > 0 ? letters(n) : " ";
      case "upperLetter": return n > 0 ? letters(n).toUpperCase() : " ";
      case "lowerRoman": return n > 0 ? roman(n).toLowerCase() : " ";
      case "upperRoman": return n > 0 ? roman(n) : " ";
      case "ordinal": return en ? `${n}${enSuffix(n)}` : /^(pl|de|cs|sk|fi|da|no|nb|nn|hu|sl|hr|sr|tr|et|lv|lt)/i.test(lang || "pl") ? `${n}.` : null;
      case "cardinalText": return n < 0 || n > 999999 ? null : en ? cap(enCardinal(n)) : /^pl/i.test(lang || "pl") ? cap(plCardinal(n)) : null;
      case "ordinalText": return n < 0 || n > 999999 ? null : en ? cap(enOrdinal(n)) : /^pl/i.test(lang || "pl") ? cap(plOrdinal(n)) : null;
      case "none": return "";
      // dalsze formaty — sprawdzone w Wordzie (tools/word-oracle/probe-fmts.js)
      case "chicago": return n > 0 ? "*†‡§"[(n - 1) % 4].repeat(Math.floor((n - 1) / 4) + 1) : " ";
      case "decimalEnclosedCircle": return n >= 1 && n <= 20 ? String.fromCodePoint(0x2460 + n - 1) : String(n);
      case "decimalEnclosedParen": return n >= 1 && n <= 20 ? String.fromCodePoint(0x2474 + n - 1) : String(n);
      case "decimalEnclosedFullstop": return n >= 1 && n <= 20 ? String.fromCodePoint(0x2488 + n - 1) : String(n);
      case "decimalFullWidth": return String(n).replace(/\d/g, (d) => String.fromCodePoint(0xff10 + +d));
      case "hex": return n.toString(16).toUpperCase();
      case "numberInDash": return `- ${n} -`;
      default: return null;
    }
  }

  // ── liczenie po akapitach ──────────────────────────────────────────────────
  const NUM_CLASS = /(?:^|\s)docx-num-([\w-]+?)-(\d+)(?:\s|$)/;
  // te same akapity i ta sama kolejność co collectPreviewParagraphElements / akapity body w pliku
  function bodyParagraphs(host) {
    return Array.from(host.querySelectorAll("section.docx > article p:not([data-dwb-cont])"));
  }
  // język akapitu (słownie / porządkowo): znak akapitu, potem pierwszy fragment, potem dokument
  function paraLang(i) {
    return model.paraLangs[i] || model.lang;
  }

  function compute(host) {
    if (!model || !host) return 0;
    const state = new Map(); // abstractNumId → wartości poziomów (undefined = jeszcze nie liczony)
    const seen = new Set();
    let n = 0;
    bodyParagraphs(host).forEach((p, i) => {
      if (p.closest(".dwb-notes")) return;
      const m = NUM_CLASS.exec(p.className || "");
      const num = m && model.nums.get(m[1]);
      const def = num && levelDef(m[1], +m[2]);
      if (!def || def.fmt === "bullet" || def.text == null) {
        if (p.hasAttribute("data-dwb-num")) p.removeAttribute("data-dwb-num");
        // punktor zostaje z rysunku docx-preview; znak do statystyki (Word liczy go jak tekst)
        const bullet = def?.fmt === "bullet" ? `${def.text || ""}${def.suff === "space" ? "\u00a0" : def.suff === "nothing" ? "" : "\t"}` : null;
        if (bullet == null) p.removeAttribute("data-dwb-bullet");
        else if (p.getAttribute("data-dwb-bullet") !== bullet) p.setAttribute("data-dwb-bullet", bullet);
        return;
      }
      if (p.hasAttribute("data-dwb-bullet")) p.removeAttribute("data-dwb-bullet");
      const numId = m[1];
      const ilvl = +m[2];
      const key = num.abs;
      if (!state.has(key)) state.set(key, []);
      const vals = state.get(key);
      // pierwszy akapit danej listy (w:num): startOverride ustawia licznik definicji
      if (!seen.has(numId)) {
        seen.add(numId);
        num.overrides.forEach((o, lv) => {
          if (o.start != null) vals[lv] = o.start - 1;
          else if (o.lvl) vals[lv] = o.lvl.start - 1;
          for (let k = lv + 1; k < 9; k++) vals[k] = undefined;
        });
      }
      // pominięte wyższe poziomy liczą się jako swój start (Word: „2.1.1.” zaraz po „2.”)
      for (let k = 0; k < ilvl; k++) if (vals[k] === undefined) vals[k] = levelDef(numId, k)?.start ?? 1;
      vals[ilvl] = vals[ilvl] === undefined ? def.start : vals[ilvl] + 1;
      // niższe poziomy: od nowa po każdym wyższym (domyślnie), po wskazanym (lvlRestart) albo nigdy (0)
      for (let k = ilvl + 1; k < 9; k++) {
        const r = levelDef(numId, k)?.restart;
        if (r == null || (r > 0 && ilvl < r)) vals[k] = undefined;
      }
      let ok = true;
      let label = def.text.replace(/%(\d)/g, (_, d) => {
        const k = +d - 1;
        const lv = levelDef(numId, k);
        if (!lv) { ok = false; return ""; }
        const v = vals[k] === undefined ? lv.start : vals[k];
        const fmt = def.isLgl && lv.fmt !== "none" ? "decimal" : lv.fmt; // numeracja prawnicza: same cyfry (Word: „1.1.”, nie „1.a.”)
        if (fmt === "bullet") { ok = false; return ""; }
        const s = formatValue(v, fmt, fmt === "ordinal" || /Text$/.test(fmt) ? paraLang(i) : model.lang);
        if (s == null) ok = false;
        return s ?? "";
      });
      if (!ok) { if (p.hasAttribute("data-dwb-num")) p.removeAttribute("data-dwb-num"); return; }
      label += def.suff === "space" ? " " : def.suff === "nothing" ? "" : "\t";
      if (p.getAttribute("data-dwb-num") !== label) p.setAttribute("data-dwb-num", label);
      n++;
    });
    return n;
  }

  // ── wpięcie w podgląd ──────────────────────────────────────────────────────
  const watched = new WeakSet();
  function watch(host) {
    if (watched.has(host) || typeof MutationObserver !== "function") return;
    watched.add(host);
    let pending = false;
    const mo = new MutationObserver((records) => {
      // tylko zmiany akapitów (nowy/usunięty punkt, zmiana poziomu); nasze data-dwb-num nie
      if (pending || !records.some((r) => r.type === "attributes" || Array.from(r.addedNodes).concat(Array.from(r.removedNodes)).some((x) => x.nodeType === 1 && (x.localName === "p" || x.querySelector?.("p"))))) return;
      pending = true;
      requestAnimationFrame(() => {
        pending = false;
        if (!host.isConnected) { mo.disconnect(); return; }
        refresh(host);
      });
    });
    mo.observe(host, { childList: true, subtree: true, attributes: true, attributeFilter: ["class"] });
  }

  async function ensureModel(bytes) {
    if (!bytes) return false;
    if (modelBytes === bytes && model) return true;
    if (!loading || loading.bytes !== bytes) {
      const job = loadModel(bytes).then((m) => { if (loading === job) { model = m; modelBytes = bytes; } return m; });
      job.bytes = bytes;
      loading = job;
    }
    try { await loading; return modelBytes === bytes; } catch (_) { return false; }
  }

  // numer w podglądzie, którego definicji jeszcze nie znamy (nowa lista z panelu) → model z bieżącego pliku
  function refresh(host) {
    if (!model) return;
    const unknown = bodyParagraphs(host).some((p) => { const m = NUM_CLASS.exec(p.className || ""); return m && !model.nums.has(m[1]); });
    if (unknown && typeof originalFileBytes !== "undefined" && originalFileBytes && originalFileBytes !== modelBytes) {
      ensureModel(originalFileBytes).then((ok) => { if (ok) compute(host); });
      return;
    }
    compute(host);
  }

  async function apply(host, bytes) {
    try {
      if (!(await ensureModel(bytes))) return 0;
      const n = compute(host);
      watch(host);
      return n;
    } catch (err) {
      if (typeof log === "function") log(`Numeracja list: ${err.message || err}`, "warn");
      return 0;
    }
  }

  // tekst numeru akapitu (eksport, kopiowanie) — bez znaku po numerze
  function labelOf(p) {
    const s = p?.getAttribute?.("data-dwb-num");
    return s == null ? null : s.replace(/[\t ]$/, "");
  }

  return { apply, refresh, labelOf, formatValue, _compute: compute };
})();
