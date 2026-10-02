// PDF → DOCX: osadzanie czcionek z PDF w pliku .docx (word/fonts/*.odttf).
// Gdy na komputerze/telefonie nie ma kroju z PDF (np. Bahnschrift na Macu), Word i podgląd
// aplikacji użyliby zastępczego — inne szerokości liter, inne zawijanie wierszy. Osadzona
// czcionka (wycinek z PDF: tylko użyte znaki) trzyma wygląd 1:1.
//
// pdf.js przepisuje czcionkę tak, że znaki siedzą pod kodami prywatnymi (fontChar), więc
// budujemy nową tabelę cmap: Unicode z dokumentu → ten sam glif. Nazwa w tabeli name = nazwa
// kroju w dokumencie (Word dopasowuje po nazwie wewnętrznej). Tylko kontury TrueType (Word nie
// wczytuje osadzonych CFF) i tylko licencje pozwalające edytować (fsType 0 lub 8) — „podgląd
// i druk” otwierałby dokument w Wordzie tylko do odczytu.
(function (root) {
  "use strict";
  const NS = (root.DWPdf = root.DWPdf || {});

  // Kroje obecne praktycznie wszędzie — osadzony wycinek zasłoniłby pełną czcionkę z systemu.
  const CORE = new Set(["arial", "times new roman", "courier new", "calibri", "cambria", "symbol", "wingdings", "verdana", "georgia", "tahoma", "trebuchet ms", "segoe ui", "helvetica", "consolas", "calibri light", "aptos"]);

  const u16 = (b, o) => (b[o] << 8) | b[o + 1];
  const u32 = (b, o) => ((b[o] << 24) >>> 0) + (b[o + 1] << 16) + (b[o + 2] << 8) + b[o + 3];
  const s16 = (b, o) => {
    const v = u16(b, o);
    return v & 0x8000 ? v - 0x10000 : v;
  };

  function readTables(bytes) {
    const tag = (o) => String.fromCharCode(bytes[o], bytes[o + 1], bytes[o + 2], bytes[o + 3]);
    const n = u16(bytes, 4);
    const tables = {};
    for (let i = 0; i < n; i++) {
      const o = 12 + i * 16;
      const off = u32(bytes, o + 8), len = u32(bytes, o + 12);
      if (off + len > bytes.length) return null;
      tables[tag(o)] = bytes.subarray(off, off + len);
    }
    return tables;
  }

  // cmap (format 4 / 12 / 0 / 6) → Map(kod → glif)
  function readCmap(cmap) {
    const map = new Map();
    if (!cmap) return map;
    const n = u16(cmap, 2);
    const subs = [];
    for (let i = 0; i < n; i++) subs.push({ pid: u16(cmap, 4 + i * 8), eid: u16(cmap, 6 + i * 8), off: u32(cmap, 8 + i * 8) });
    const rank = (s) => (s.pid === 3 && s.eid === 10 ? 0 : s.pid === 3 && s.eid === 1 ? 1 : s.pid === 3 && s.eid === 0 ? 2 : s.pid === 0 ? 3 : 4);
    subs.sort((a, b) => rank(a) - rank(b));
    for (const s of subs) {
      const o = s.off;
      if (o + 4 > cmap.length) continue;
      const fmt = u16(cmap, o);
      if (fmt === 4) {
        const segX2 = u16(cmap, o + 6);
        const endO = o + 14, startO = endO + segX2 + 2, deltaO = startO + segX2, rangeO = deltaO + segX2;
        for (let k = 0; k < segX2 / 2; k++) {
          const end = u16(cmap, endO + k * 2), start = u16(cmap, startO + k * 2);
          const delta = s16(cmap, deltaO + k * 2), ro = u16(cmap, rangeO + k * 2);
          for (let c = start; c <= end && c !== 0xffff; c++) {
            let g;
            if (!ro) g = (c + delta) & 0xffff;
            else {
              const gi = rangeO + k * 2 + ro + (c - start) * 2;
              if (gi + 1 >= cmap.length) continue;
              g = u16(cmap, gi);
              if (g) g = (g + delta) & 0xffff;
            }
            if (g && !map.has(c)) map.set(c, g);
          }
        }
      } else if (fmt === 12) {
        const nGroups = u32(cmap, o + 12);
        for (let k = 0; k < nGroups; k++) {
          const b = o + 16 + k * 12;
          const sc = u32(cmap, b), ec = u32(cmap, b + 4), sg = u32(cmap, b + 8);
          for (let c = sc; c <= ec && c - sc < 70000; c++) if (!map.has(c)) map.set(c, sg + (c - sc));
        }
      } else if (fmt === 0) {
        for (let c = 0; c < 256; c++) if (cmap[o + 6 + c] && !map.has(c)) map.set(c, cmap[o + 6 + c]);
      } else if (fmt === 6) {
        const first = u16(cmap, o + 6), cnt = u16(cmap, o + 8);
        for (let k = 0; k < cnt; k++) {
          const g = u16(cmap, o + 10 + k * 2);
          if (g && !map.has(first + k)) map.set(first + k, g);
        }
      }
      if (map.size) break;
    }
    return map;
  }

  // Map(Unicode → glif) → tabela cmap: (3,1) format 4 + (3,10) format 12 gdy są znaki spoza BMP.
  function buildCmap(uniToGid) {
    const entries = [...uniToGid].filter(([u, g]) => u > 0 && g > 0).sort((a, b) => a[0] - b[0]);
    const bmp = entries.filter(([u]) => u < 0xffff);
    // segmenty format 4: jeden kod = jeden segment z deltą (proste i zawsze poprawne)
    const segs = [];
    for (const [u, g] of bmp) {
      const last = segs[segs.length - 1];
      if (last && u === last.end + 1 && ((g - u) & 0xffff) === last.delta) last.end = u;
      else segs.push({ start: u, end: u, delta: (g - u) & 0xffff });
    }
    segs.push({ start: 0xffff, end: 0xffff, delta: 1 });
    const segX2 = segs.length * 2;
    let p2 = 1, e = 0;
    while (p2 * 2 <= segs.length) {
      p2 *= 2;
      e++;
    }
    const f4len = 16 + segs.length * 8;
    const f4 = new Uint8Array(f4len);
    const dv4 = new DataView(f4.buffer);
    dv4.setUint16(0, 4); dv4.setUint16(2, f4len); dv4.setUint16(4, 0);
    dv4.setUint16(6, segX2); dv4.setUint16(8, p2 * 2); dv4.setUint16(10, e); dv4.setUint16(12, segX2 - p2 * 2);
    segs.forEach((sg, i) => {
      dv4.setUint16(14 + i * 2, sg.end);
      dv4.setUint16(16 + segX2 + i * 2, sg.start);
      dv4.setUint16(16 + segX2 * 2 + i * 2, sg.delta);
      dv4.setUint16(16 + segX2 * 3 + i * 2, 0);
    });
    const astral = entries.filter(([u]) => u > 0xffff);
    let f12 = null;
    if (astral.length) {
      const groups = entries.map(([u, g]) => [u, u, g]);
      f12 = new Uint8Array(16 + groups.length * 12);
      const dv = new DataView(f12.buffer);
      dv.setUint16(0, 12); dv.setUint32(4, f12.length); dv.setUint32(12, groups.length);
      groups.forEach(([s, en, g], i) => {
        dv.setUint32(16 + i * 12, s); dv.setUint32(20 + i * 12, en); dv.setUint32(24 + i * 12, g);
      });
    }
    const nSub = f12 ? 2 : 1;
    const head = 4 + nSub * 8;
    const out = new Uint8Array(head + f4.length + (f12 ? f12.length : 0));
    const dv = new DataView(out.buffer);
    dv.setUint16(0, 0); dv.setUint16(2, nSub);
    dv.setUint16(4, 3); dv.setUint16(6, 1); dv.setUint32(8, head);
    if (f12) {
      dv.setUint16(12, 3); dv.setUint16(14, 10); dv.setUint32(16, head + f4.length);
    }
    out.set(f4, head);
    if (f12) out.set(f12, head + f4.length);
    return out;
  }

  function buildName(family, sub, psName) {
    const recs = [[1, family], [2, sub], [3, `${family} ${sub}; Documents Workbench`], [4, sub === "Regular" ? family : `${family} ${sub}`], [6, psName]];
    const strs = recs.map(([, s]) => {
      const a = new Uint8Array(s.length * 2);
      for (let i = 0; i < s.length; i++) {
        a[i * 2] = s.charCodeAt(i) >> 8;
        a[i * 2 + 1] = s.charCodeAt(i) & 0xff;
      }
      return a;
    });
    const head = 6 + recs.length * 12;
    const total = head + strs.reduce((n, a) => n + a.length, 0);
    const out = new Uint8Array(total);
    const dv = new DataView(out.buffer);
    dv.setUint16(0, 0); dv.setUint16(2, recs.length); dv.setUint16(4, head);
    let off = 0;
    recs.forEach(([id], i) => {
      const r = 6 + i * 12;
      dv.setUint16(r, 3); dv.setUint16(r + 2, 1); dv.setUint16(r + 4, 0x409); dv.setUint16(r + 6, id);
      dv.setUint16(r + 8, strs[i].length); dv.setUint16(r + 10, off);
      out.set(strs[i], head + off);
      off += strs[i].length;
    });
    return out;
  }

  function checksum(t) {
    let sum = 0;
    const n = Math.ceil(t.length / 4);
    for (let i = 0; i < n; i++) {
      const o = i * 4;
      sum = (sum + (((t[o] || 0) << 24) >>> 0) + ((t[o + 1] || 0) << 16) + ((t[o + 2] || 0) << 8) + (t[o + 3] || 0)) >>> 0;
    }
    return sum >>> 0;
  }

  function writeFont(tables, sfntVersion) {
    const tags = Object.keys(tables).sort();
    const n = tags.length;
    let p2 = 1, e = 0;
    while (p2 * 2 <= n) {
      p2 *= 2;
      e++;
    }
    const headerLen = 12 + n * 16;
    let total = headerLen;
    for (const t of tags) total += (tables[t].length + 3) & ~3;
    const out = new Uint8Array(total);
    const dv = new DataView(out.buffer);
    dv.setUint32(0, sfntVersion);
    dv.setUint16(4, n); dv.setUint16(6, p2 * 16); dv.setUint16(8, e); dv.setUint16(10, n * 16 - p2 * 16);
    let off = headerLen;
    let headOff = -1;
    tags.forEach((t, i) => {
      const data = tables[t];
      const r = 12 + i * 16;
      for (let k = 0; k < 4; k++) out[r + k] = t.charCodeAt(k);
      if (t === "head") {
        headOff = off;
        const copy = data.slice();
        copy[8] = copy[9] = copy[10] = copy[11] = 0; // checkSumAdjustment liczony niżej
        out.set(copy, off);
        dv.setUint32(r + 4, checksum(copy));
      } else {
        out.set(data, off);
        dv.setUint32(r + 4, checksum(data));
      }
      dv.setUint32(r + 8, off);
      dv.setUint32(r + 12, data.length);
      off += (data.length + 3) & ~3;
    });
    if (headOff >= 0) dv.setUint32(headOff + 8, (0xb1b0afba - checksum(out)) >>> 0);
    return out;
  }

  function newGuid() {
    const b = new Uint8Array(16);
    if (root.crypto?.getRandomValues) root.crypto.getRandomValues(b);
    else for (let i = 0; i < 16; i++) b[i] = (Math.random() * 256) | 0;
    const h = [...b].map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase();
    return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
  }

  // Zaciemnienie czcionki wg ECMA-376 (17.8.1): XOR pierwszych 32 bajtów z kluczem GUID.
  function obfuscate(bytes, guid) {
    const hex = guid.replace(/[{}-]/g, "");
    const key = new Array(16);
    for (let i = 0; i < 16; i++) key[16 - i - 1] = parseInt(hex.substr(i * 2, 2), 16);
    const out = bytes.slice();
    for (let i = 0; i < 32 && i < out.length; i++) out[i] ^= key[i % 16];
    return out;
  }

  /**
   * @param usage Map(fontRef → { fontObj, family, bold, italic, chars: Map(unicode → fontChar) })
   * @returns [{ family, style: "Regular"|"Bold"|"Italic"|"BoldItalic", bytes (zaciemnione), key }]
   */
  function buildEmbeddedFonts(usage, opts = {}) {
    const chosen = new Map(); // rodzina|styl → kandydat z największą liczbą znaków
    for (const [, u] of usage) {
      const fam = String(u.family || "").trim();
      if (!fam || CORE.has(fam.toLowerCase()) || u.symbolic) continue;
      const data = u.fontObj?.data;
      if (!data || data.length < 100 || data.length > (opts.maxBytes || 4e6)) continue;
      const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
      const ver = u32(bytes, 0);
      if (ver !== 0x00010000 && ver !== 0x74727565) continue; // tylko TrueType (glyf)
      const tables = readTables(bytes);
      if (!tables || !tables.glyf || !tables.loca || !tables.head || !tables.cmap) continue;
      const os2 = tables["OS/2"];
      if (os2 && os2.length > 9) {
        const fsType = u16(os2, 8) & 0x000f;
        if (fsType !== 0 && !(fsType & 0x0008)) continue; // licencja: tylko instalowalne/edytowalne
      }
      const style = u.bold && u.italic ? "BoldItalic" : u.bold ? "Bold" : u.italic ? "Italic" : "Regular";
      const key = fam + "|" + style;
      const prev = chosen.get(key);
      if (!prev || u.chars.size > prev.u.chars.size) chosen.set(key, { u, tables, ver, style, fam });
    }
    const out = [];
    for (const { u, tables, ver, style, fam } of chosen.values()) {
      try {
        const oldMap = readCmap(tables.cmap);
        const uni = new Map();
        for (const [ucp, fc] of u.chars) {
          const g = oldMap.get(fc);
          if (g) uni.set(ucp, g);
        }
        if (!uni.size) continue;
        const nt = { ...tables };
        nt.cmap = buildCmap(uni);
        const sub = style === "BoldItalic" ? "Bold Italic" : style;
        nt.name = buildName(fam, sub, (fam + "-" + style).replace(/[^A-Za-z0-9-]/g, ""));
        if (nt["OS/2"] && nt["OS/2"].length >= 64) {
          const o = nt["OS/2"].slice();
          let sel = u16(o, 62) & ~0x61; // bez ITALIC/BOLD/REGULAR
          if (style === "Bold" || style === "BoldItalic") sel |= 0x20;
          if (style === "Italic" || style === "BoldItalic") sel |= 0x01;
          if (style === "Regular") sel |= 0x40;
          o[62] = sel >> 8;
          o[63] = sel & 0xff;
          // zakres znaków w OS/2: pierwszy/ostatni Unicode
          const keys = [...uni.keys()].sort((a, b) => a - b);
          const first = Math.min(keys[0], 0xffff), last = Math.min(keys[keys.length - 1], 0xffff);
          o[64] = first >> 8; o[65] = first & 0xff; o[66] = last >> 8; o[67] = last & 0xff;
          nt["OS/2"] = o;
        }
        delete nt.post; // nazwy glifów z pdf.js bywają niespójne z nową cmap — post v3 niżej
        nt.post = (() => {
          const p = new Uint8Array(32);
          const dv = new DataView(p.buffer);
          dv.setUint32(0, 0x00030000);
          return p;
        })();
        const font = writeFont(nt, 0x00010000);
        const guid = newGuid();
        out.push({ family: fam, style, bytes: obfuscate(font, guid), plain: opts.keepPlain ? font : null, key: guid, chars: uni.size });
      } catch (_) { /* tej czcionki nie osadzamy — dokument i tak powstanie */ }
    }
    return out;
  }

  NS.buildEmbeddedFonts = buildEmbeddedFonts;
  NS._fonts = { readCmap, buildCmap, readTables, obfuscate };
})(typeof globalThis !== "undefined" ? globalThis : window);
