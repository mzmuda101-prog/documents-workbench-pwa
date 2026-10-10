// Run-level styles — DOM spans ↔ Word w:r / w:rPr (bold, color, font, underline).

function parseCssColorToWordHex(color) {
  if (!color) return null;
  const rgb = color.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (rgb) {
    return [rgb[1], rgb[2], rgb[3]]
      .map((n) => parseInt(n, 10).toString(16).padStart(2, "0"))
      .join("")
      .toUpperCase();
  }
  if (color.startsWith("#")) return color.slice(1).toUpperCase();
  return null;
}

// Wyróżnienie tekstu (Word w:highlight = 16 nazw) albo cieniowanie (w:shd fill = dowolny kolor).
// Podgląd rysuje oba jako background-color: nazwę zostawia nazwą, cieniowanie daje „#hex”.
const WORD_HIGHLIGHTS = ["yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray", "black", "white"];
const WORD_HIGHLIGHT_BY_CSS = new Map(WORD_HIGHLIGHTS.map((n) => [n.toLowerCase(), n]));
function normHighlight(v) {
  if (!v) return "";
  const s = String(v).trim().toLowerCase();
  if (!s || s === "transparent" || s === "none" || s === "auto" || s === "rgba(0, 0, 0, 0)") return "";
  if (WORD_HIGHLIGHT_BY_CSS.has(s)) return s;
  const hex = parseCssColorToWordHex(s);
  return hex ? `#${hex.toLowerCase()}` : s;
}

function cssFontSizeToPt(val) {
  if (!val) return null;
  const px = String(val).match(/^([\d.]+)px$/i);
  if (px) return Math.round(parseFloat(px[1]) * 0.75 * 2) / 2;
  const pt = String(val).match(/^([\d.]+)pt$/i);
  if (pt) return parseFloat(pt[1]);
  const num = parseFloat(val);
  return Number.isFinite(num) ? num : null;
}

// Indeks górny / dolny: w:vertAlign (superscript / subscript / baseline) ↔ CSS vertical-align.
// docx-preview rysuje go jako <sup>/<sub> w fragmencie — czytamy oba zapisy.
const VERT_ALIGN_BY_CSS = { super: "superscript", sub: "subscript", baseline: "baseline" };
const VERT_ALIGN_CSS = { superscript: "super", subscript: "sub", baseline: "baseline" };
const normVertAlign = (v) => (v === "superscript" || v === "subscript" ? v : "");

function parseSpanStyle(cssText) {
  const style = {};
  if (!cssText) return style;
  cssText.split(";").forEach((chunk) => {
    const idx = chunk.indexOf(":");
    if (idx < 0) return;
    const key = chunk.slice(0, idx).trim().toLowerCase();
    const val = chunk.slice(idx + 1).trim();
    if (key === "font-weight") { if (val === "bold" || parseInt(val, 10) >= 600) style.bold = true; else if (val === "normal" || parseInt(val, 10) < 600) style.bold = false; }
    if (key === "font-style") { if (val === "italic") style.italic = true; else if (val === "normal") style.italic = false; }
    if ((key === "text-decoration" || key === "text-decoration-line") && val.includes("underline")) style.underline = true;
    if ((key === "text-decoration" || key === "text-decoration-line") && val.includes("line-through")) style.strike = true;
    if (key === "color") style.color = val;
    // indeks górny / dolny (w:vertAlign); „baseline” = jawnie wyłączony (pisanie za indeksem)
    if (key === "vertical-align") { const va = VERT_ALIGN_BY_CSS[val.toLowerCase()]; if (va) style.vertAlign = va; }
    if (key === "background-color" || key === "background") { const h = normHighlight(val); if (h) style.highlight = h; }
    // pierwsza rodzina z listy, bez cudzysłowów z OBU stron: podgląd ma „"DM Sans", sans-serif” —
    // dawniej obcinany był tylko początkowy cudzysłów i do pliku szła nazwa kroju „DM Sans"”
    if (key === "font-family") style.fontFamily = val.split(",")[0].trim().replace(/^["']+|["']+$/g, "").trim();
    if (key === "font-size") {
      const pt = cssFontSizeToPt(val);
      if (pt) style.fontSize = `${pt}pt`;
    }
  });
  return style;
}

function runStyleToCss(run) {
  const parts = [];
  if (run.bold) parts.push("font-weight:bold");
  else if (run.bold === false) parts.push("font-weight:normal"); // wyłączone w środku pogrubionego fragmentu
  if (run.italic) parts.push("font-style:italic");
  else if (run.italic === false) parts.push("font-style:normal");
  if (run.underline || run.strike) parts.push(`text-decoration:${[run.underline && "underline", run.strike && "line-through"].filter(Boolean).join(" ")}`);
  if (run.color) parts.push(`color:${run.color}`);
  if (run.fontFamily) {
    // z rodziną ogólną (bezszeryfowy/szeryfowy) — krój spoza urządzenia nie spada na Times
    const ff = `"${run.fontFamily}"`;
    parts.push(`font-family:${(typeof withGenericFontFallback === "function" && withGenericFontFallback(ff)) || ff}`);
  }
  if (run.fontSize) parts.push(`font-size:${run.fontSize}`);
  if (run.highlight) parts.push(`background-color:${run.highlight}`);
  if (VERT_ALIGN_CSS[run.vertAlign]) parts.push(`vertical-align:${VERT_ALIGN_CSS[run.vertAlign]}`);
  return parts.join(";");
}

// Trzy stany: włączone / jawnie WYŁĄCZONE (w:b w:val="0" — zwykły tekst w pogrubionym nagłówku)
// / nieustawione (jak styl akapitu). Dawniej wyłączone = nieustawione: sklejanie fragmentów gubiło
// wyłączenie i tekst w Wordzie wracał pogrubiony.
const tri = (v) => (v === true ? 1 : v === false ? 0 : -1);
function runsStyleEqual(a, b) {
  // podkreślenie/przekreślenie: dwa stany — podgląd nie zostawia śladu jawnego „wyłączone”
  // (w:u w:val="none" rysuje jak brak podkreślenia), więc porównanie i tak by się rozjechało
  return tri(a.bold) === tri(b.bold)
    && tri(a.italic) === tri(b.italic)
    && !!a.underline === !!b.underline
    && !!a.strike === !!b.strike
    && (parseCssColorToWordHex(a.color) || "") === (parseCssColorToWordHex(b.color) || "") // „rgb(5, 99, 193)” z podglądu = „#0563C1” z pliku
    && (a.fontFamily || "") === (b.fontFamily || "")
    && (a.fontSize || "") === (b.fontSize || "")
    && (a.link || "") === (b.link || "")
    && normHighlight(a.highlight) === normHighlight(b.highlight)
    && normVertAlign(a.vertAlign) === normVertAlign(b.vertAlign);
}

function mergeAdjacentRuns(runs) {
  const out = [];
  (runs || []).forEach((run) => {
    if (run.break) {
      out.push({ break: true });
      return;
    }
    if (run.island) { out.push({ ...run }); return; } // pole formularza — nienaruszalna „wyspa”
    if (run.tab) { out.push({ ...run }); return; } // tabulator — osobny element (<w:tab/>)
    const prev = out[out.length - 1];
    if (prev && !prev.break && !prev.island && !prev.tab && runsStyleEqual(prev, run)) {
      prev.text += run.text;
      return;
    }
    out.push({ ...run });
  });
  return out;
}

function extractRunsFromPreviewParagraph(pEl) {
  if (!pEl) return [];
  const runs = [];

  function walk(node, inherited = {}) {
    if (node.nodeType === Node.TEXT_NODE) {
      // U+FEFF = miejsce na kursor za odnośnikiem przypisu (doc-notes.js) — nie jest treścią
      const text = (node.textContent || "").replace(/\uFEFF/g, "");
      if (!text) return;
      // Shift+Enter / wklejone wiersze: przeglądarka przy white-space: pre-wrap wstawia znak
      // „\n” zamiast <br>. W pliku musi to być <w:br/> — „\n” w <w:t> Word pokazuje jako spację.
      text.split("\n").forEach((part, i) => {
        if (i) runs.push({ break: true });
        // tabulator wpisany klawiszem Tab = znak „\t” — w pliku <w:tab/>, jak tabulator z pliku
        part.split("\t").forEach((seg, j) => {
          if (j) runs.push({ tab: true, text: "\t", ...inherited });
          if (seg) runs.push({ text: seg, ...inherited });
        });
      });
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    // tabulator z pliku: docx-preview rysuje <span class="docx-tab">&emsp;</span>. Dawniej szedł do
    // pliku jako spacja szerokości „em” — po edycji akapitu wyrównanie do tabulatora znikało.
    if (node.classList?.contains("docx-tab")) {
      runs.push({ tab: true, text: "\t", ...inherited });
      return;
    }
    // pole formularza w zdaniu (docx-forms.js): cała kontrolka z pliku, bez zmian
    if (node.dataset?.ff && docIslandXml.has(node.dataset.ff)) {
      runs.push({ island: docIslandXml.get(node.dataset.ff), text: node.textContent || "" });
      return;
    }
    if (node.dataset?.cm && docIslandXml.has(node.dataset.cm)) { // znacznik komentarza (docx-inline-edit.js)
      runs.push({ island: docIslandXml.get(node.dataset.cm), text: node.dataset.cmKind === "sym" ? node.textContent : "" });
      return;
    }
    const tag = node.localName.toLowerCase();
    if (tag === "br") {
      if (!node.dataset?.dwbPh) runs.push({ break: true }); // <br data-dwb-ph> = widoczny pusty wiersz po Shift+Enter, nie treść
      return;
    }
    const style = { ...inherited };
    if (tag === "span") Object.assign(style, parseSpanStyle(node.getAttribute("style") || ""));
    if (tag === "b" || tag === "strong") style.bold = true;
    if (tag === "i" || tag === "em") style.italic = true;
    if (tag === "u") style.underline = true;
    if (tag === "s" || tag === "strike") style.strike = true;
    if (tag === "sup") style.vertAlign = "superscript";
    if (tag === "sub") style.vertAlign = "subscript";
    // link (w:hyperlink): podgląd rysuje <a>. data-dwb-link = odwołanie z pliku („#zakładka”
    // albo „rel:rIdN”, nadane przy oznaczaniu akapitów), nowy link ma sam adres.
    if (tag === "a" && !node.classList.contains("doc-xref")) style.link = node.dataset.dwbLink || node.getAttribute("href") || "";
    node.childNodes.forEach((child) => walk(child, style));
  }

  pEl.childNodes.forEach((child) => walk(child, {}));
  return mergeAdjacentRuns(runs);
}

function previewRunsToPlainText(runs) {
  return (runs || []).map((r) => (r.break ? "\n" : r.text || "")).join("");
}

function runsEqual(a, b) {
  const aa = mergeAdjacentRuns(a || []);
  const bb = mergeAdjacentRuns(b || []);
  if (aa.length !== bb.length) return false;
  return aa.every((run, i) => {
    const other = bb[i];
    if (!!run.break !== !!other.break) return false;
    if (run.break) return true;
    if (run.island || other.island) return run.island === other.island;
    if (!!run.tab !== !!other.tab) return false;
    if (run.tab) return true; // styl samego tabulatora nie ma znaczenia
    return run.text === other.text && runsStyleEqual(run, other);
  });
}

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// Odwołanie linku z <w:hyperlink>: „#zakładka” (w dokumencie) albo „rel:rIdN” (adres w powiązaniach).
function hyperlinkRef(h) {
  const anchor = h.getAttributeNS(W_NS, "anchor") || h.getAttribute("w:anchor");
  if (anchor) return `#${anchor}`;
  const rid = h.getAttributeNS(R_NS, "id") || h.getAttribute("r:id");
  return rid ? `rel:${rid}` : "";
}

function getParagraphRunElements(pEl) {
  const runs = [];
  for (let i = 0; i < pEl.childNodes.length; i++) {
    const child = pEl.childNodes[i];
    if (child.nodeType !== 1) continue;
    if (child.localName === "r" && child.namespaceURI === W_NS) runs.push(child);
    else if (child.localName === "hyperlink" && child.namespaceURI === W_NS) {
      for (let j = 0; j < child.childNodes.length; j++) {
        const sub = child.childNodes[j];
        if (sub.nodeType === 1 && sub.localName === "r" && sub.namespaceURI === W_NS) runs.push(sub);
      }
    }
  }
  return runs;
}

// Proste pole formularza w zdaniu (w:sdt bezpośrednio w akapicie) = „wyspa”: zapis akapitu
// wstawia je z powrotem dokładnie takie, jak w pliku (docx-forms.js formIslandSdt decyduje, które).
const docIslandXml = new Map(); // klucz pola („s<N>”) → XML kontrolki — dla podglądu
function paragraphXmlParts(pEl) {
  const parts = [];
  for (let i = 0; i < pEl.childNodes.length; i++) {
    const child = pEl.childNodes[i];
    if (child.nodeType !== 1 || child.namespaceURI !== W_NS) continue;
    if (child.localName === "r") parts.push(child);
    else if (child.localName === "commentRangeStart" || child.localName === "commentRangeEnd") parts.push(child); // znaczniki komentarza
    // zakładki (cel linku „W dokumencie”, odsyłacza „Rysunek 1”, spisu treści) — w swoim miejscu
    // tekstu; dawniej zostawały przed przepisanym tekstem i zakładka kurczyła się do pustego punktu
    else if (child.localName === "bookmarkStart" || child.localName === "bookmarkEnd") parts.push(child);
    else if (child.localName === "hyperlink") {
      for (let j = 0; j < child.childNodes.length; j++) {
        const sub = child.childNodes[j];
        if (sub.nodeType === 1 && sub.localName === "r" && sub.namespaceURI === W_NS) parts.push(sub);
      }
    } else if (child.localName === "sdt" && typeof formIslandSdt === "function" && formIslandSdt(child)) parts.push(child);
  }
  return parts;
}

function isCommentReferenceRun(r) {
  return r.localName === "r" && Array.from(r.childNodes).some((n) => n.localName === "commentReference");
}

// Przypisy (doc-notes.js): „ref” = odnośnik w treści (podgląd rysuje numer <sup>), „mark” = numer
// na początku tekstu przypisu (podgląd go nie rysuje — numer daje lista). Oba to „wyspy”: zapis
// akapitu oddaje cały fragment z pliku bez zmian, więc pisanie obok nie gubi przypisu.
function noteRunKind(r) {
  if (r.localName !== "r") return null;
  for (const n of r.childNodes) {
    if (n.localName === "footnoteReference" || n.localName === "endnoteReference") return "ref";
    if (n.localName === "footnoteRef" || n.localName === "endnoteRef") return "mark";
  }
  return null;
}

// Włączone/wyłączone (w:b, w:i, w:strike…): sam znacznik = włączone, w:val="0|false|off" = WYŁĄCZONE
// (np. zwykły tekst w pogrubionym nagłówku). Dawniej każdy znacznik był „włączony” — po Cofnij
// taki tekst wracał pogrubiony, a porównanie z podglądem widziało zmianę, której nie było.
function wOnOff(el) {
  if (!el) return undefined;
  return !/^(0|false|off|none)$/i.test(getWVal(el) || "");
}
// Krój jak w docx-preview (parseFont): ascii, potem krój motywu (var(--docx-…-font)), potem
// eastAsia — tak samo czyta go podgląd, więc porównanie i zapis mówią tym samym językiem.
function xmlRunFont(fonts) {
  const clean = (v) => String(v || "").replace(/["']/g, "").trim();
  const ascii = clean(fonts.getAttributeNS(W_NS, "ascii") || fonts.getAttributeNS(W_NS, "hAnsi"));
  if (ascii) return ascii;
  const theme = fonts.getAttributeNS(W_NS, "asciiTheme") || fonts.getAttributeNS(W_NS, "hAnsiTheme");
  if (theme) return `var(--docx-${theme}-font)`;
  return clean(fonts.getAttributeNS(W_NS, "eastAsia")) || null;
}
// Format fragmentu w:r w modelu akapitu (pogrubienie, kolor, krój, rozmiar, link…).
function xmlRunStyle(r) {
  const style = {};
  if (r.parentNode?.localName === "hyperlink") {
    const ref = hyperlinkRef(r.parentNode);
    if (ref) style.link = ref;
  }
  const rPr = Array.from(r.childNodes).find((n) => n.localName === "rPr" && n.namespaceURI === W_NS);
  if (!rPr) return style;
  const kid = (name) => Array.from(rPr.childNodes).find((n) => n.localName === name);
  for (const [k, tag] of [["bold", "b"], ["italic", "i"], ["strike", "strike"]]) {
    const v = wOnOff(kid(tag));
    if (v !== undefined) style[k] = v;
  }
  const u = kid("u");
  if (u) style.underline = !/^none$/i.test(getWVal(u) || "single");
  const colorEl = kid("color");
  const hex = colorEl ? getWVal(colorEl) : null;
  if (hex && hex !== "auto") style.color = `#${hex.replace(/^#/, "")}`;
  const fonts = kid("rFonts");
  if (fonts) style.fontFamily = xmlRunFont(fonts);
  const hl = kid("highlight");
  const shd = kid("shd");
  const hlVal = hl ? getWVal(hl) : null;
  const shdFill = shd ? (shd.getAttributeNS(W_NS, "fill") || shd.getAttribute("w:fill")) : null;
  if (hlVal && hlVal !== "none") style.highlight = normHighlight(hlVal);
  else if (shdFill && shdFill !== "auto") style.highlight = normHighlight(`#${shdFill}`);
  const sz = kid("sz");
  if (sz) {
    const half = parseInt(getWVal(sz) || "0", 10);
    if (half) style.fontSize = `${half / 2}pt`;
  }
  const va = kid("vertAlign");
  if (va && normVertAlign(getWVal(va))) style.vertAlign = getWVal(va);
  return style;
}

// Symbol Worda (w:sym — znak z kroju Symbol/Wingdings, np. „§”, kratka formularza): podgląd rysuje
// go jako <span style="font-family: Symbol">znak</span>. Fragment z SAMYM symbolem jest wyspą
// z tekstem = ten znak (zapis oddaje oryginalne w:sym). Dawniej odczyt z pliku go pomijał,
// a zapis pisanego akapitu po cichu gubił symbol.
function isSymRun(r) {
  return r?.localName === "r" && Array.from(r.childNodes).some((n) => n.localName === "sym" && n.namespaceURI === W_NS);
}
function symRunText(r) {
  return Array.from(r.childNodes).filter((n) => n.localName === "sym").map((n) => {
    const code = parseInt(n.getAttributeNS(W_NS, "char") || n.getAttribute("w:char") || "", 16);
    return Number.isFinite(code) ? String.fromCharCode(code) : "";
  }).join("");
}
// null = symbole da się zachować jako wyspy; inaczej powód blokady akapitu.
function paragraphSymbolLock(xp) {
  for (const s of Array.from(xp.getElementsByTagNameNS(W_NS, "sym"))) {
    const r = s.parentNode;
    if (r?.localName !== "r" || r.parentNode !== xp) return "lockSymbol";
    if (Array.from(r.childNodes).some((n) => n.nodeType === 1 && !["rPr", "sym"].includes(n.localName))) return "lockSymbol";
  }
  return null;
}

function extractRunsFromParagraphXml(pEl) {
  const runs = [];
  const symIslands = !paragraphSymbolLock(pEl);
  // rysunek jako wyspa — ta sama reguła co w podglądzie (docx-inline-edit.js paragraphObjectLock),
  // inaczej każdy akapit z rysunkiem wychodził „zmieniony” i był przepisywany przy zapisie
  const objIslands = typeof isObjectRun === "function" && typeof paragraphObjectLock === "function" && !paragraphObjectLock(pEl);
  paragraphXmlParts(pEl).forEach((r) => {
    if (r.localName === "sdt") {
      runs.push({ island: new XMLSerializer().serializeToString(r), text: ffText(ffKid(r, "sdtContent")) });
      return;
    }
    // komentarz: początek/koniec zakresu i fragment z odwołaniem — „wyspy” o zerowej długości
    if (symIslands && isSymRun(r)) {
      runs.push({ island: new XMLSerializer().serializeToString(r), text: symRunText(r) });
      return;
    }
    if (r.localName !== "r" || isCommentReferenceRun(r) || noteRunKind(r) || (objIslands && isObjectRun(r))) {
      runs.push({ island: new XMLSerializer().serializeToString(r), text: "" });
      return;
    }
    const style = xmlRunStyle(r);
    // Tekst i łamania w kolejności: jeden fragment może mieć <w:t>a</w:t><w:br/><w:t>b</w:t>
    // (tak zapisujemy wielowierszowe wstawienia). Dawniej samo <w:br/> kasowało tekst fragmentu.
    Array.from(r.childNodes).forEach((n) => {
      if (n.namespaceURI !== W_NS) return;
      if (n.localName === "br") runs.push({ break: true });
      else if (n.localName === "tab") runs.push({ tab: true, text: "\t", ...style });
      else if (n.localName === "t" && n.textContent) runs.push({ text: n.textContent, ...style });
    });
  });
  return mergeAdjacentRuns(runs);
}

function createRunElement(doc, run) {
  const r = doc.createElementNS(W_NS, "r");
  const rPr = doc.createElementNS(W_NS, "rPr");
  let hasPr = false;
  if (run.link) { // styl znakowy „Hiperłącze” — prawdziwe id podstawia finalizeComposeParts (docx-compose.js)
    const rs = doc.createElementNS(W_NS, "rStyle");
    setWVal(rs, "__DWB_HL__");
    rPr.appendChild(rs);
    hasPr = true;
  }
  // false = jawnie WYŁĄCZONE (Ctrl/⌘+B w pogrubionym nagłówku) — bez w:val="0" Word wziąłby
  // pogrubienie ze stylu akapitu
  if (run.bold != null) {
    const b = doc.createElementNS(W_NS, "b");
    setWVal(b, run.bold ? "1" : "0");
    rPr.appendChild(b);
    hasPr = true;
  }
  if (run.italic != null) {
    const i = doc.createElementNS(W_NS, "i");
    setWVal(i, run.italic ? "1" : "0");
    rPr.appendChild(i);
    hasPr = true;
  }
  if (run.strike != null) {
    const st = doc.createElementNS(W_NS, "strike");
    setWVal(st, run.strike ? "1" : "0");
    rPr.appendChild(st);
    hasPr = true;
  }
  if (run.underline != null) {
    const u = doc.createElementNS(W_NS, "u");
    setWVal(u, run.underline ? "single" : "none");
    rPr.appendChild(u);
    hasPr = true;
  }
  if (run.color) {
    const hex = parseCssColorToWordHex(run.color);
    if (hex) {
      const c = doc.createElementNS(W_NS, "color");
      setWVal(c, hex);
      rPr.appendChild(c);
      hasPr = true;
    }
  }
  if (run.fontFamily) {
    const rf = doc.createElementNS(W_NS, "rFonts");
    const theme = String(run.fontFamily).match(/^var\(--docx-(\w+)-font\)$/);
    if (theme) { // krój motywu (np. „Treść” = minorHAnsi) — odwołanie, nie nazwa
      rf.setAttributeNS(W_NS, "w:asciiTheme", theme[1]);
      rf.setAttributeNS(W_NS, "w:hAnsiTheme", theme[1]);
    } else {
      const name = String(run.fontFamily).replace(/["']/g, "").trim();
      rf.setAttributeNS(W_NS, "ascii", name);
      rf.setAttributeNS(W_NS, "hAnsi", name);
      rf.setAttributeNS(W_NS, "cs", name);
    }
    rPr.appendChild(rf);
    hasPr = true;
  }
  if (run.fontSize) {
    const pt = parseFloat(String(run.fontSize));
    if (pt) {
      const half = String(Math.round(pt * 2));
      const sz = doc.createElementNS(W_NS, "sz");
      setWVal(sz, half);
      rPr.appendChild(sz);
      const szCs = doc.createElementNS(W_NS, "szCs");
      setWVal(szCs, half);
      rPr.appendChild(szCs);
      hasPr = true;
    }
  }
  const hl = normHighlight(run.highlight);
  if (hl) { // kolejność w rPr: … sz, szCs, highlight, u (u dalej — Word toleruje), shd
    if (WORD_HIGHLIGHT_BY_CSS.has(hl)) {
      const h = doc.createElementNS(W_NS, "highlight");
      setWVal(h, WORD_HIGHLIGHT_BY_CSS.get(hl));
      rPr.appendChild(h);
    } else {
      const shd = doc.createElementNS(W_NS, "shd");
      setWVal(shd, "clear");
      shd.setAttributeNS(W_NS, "w:color", "auto");
      shd.setAttributeNS(W_NS, "w:fill", hl.replace(/^#/, "").toUpperCase());
      rPr.appendChild(shd);
    }
    hasPr = true;
  }
  if (VERT_ALIGN_CSS[run.vertAlign] && run.vertAlign !== "baseline") {
    const va = doc.createElementNS(W_NS, "vertAlign");
    setWVal(va, run.vertAlign);
    rPr.appendChild(va);
    hasPr = true;
  }
  if (hasPr) {
    // kolejność dzieci w:rPr wg schematu Worda (dawniej u przed color, rFonts po color)
    const order = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "strike", "color", "sz", "szCs", "highlight", "u", "shd", "vertAlign"];
    Array.from(rPr.childNodes).sort((x, y) => order.indexOf(x.localName) - order.indexOf(y.localName)).forEach((n) => rPr.appendChild(n));
    r.appendChild(rPr);
  }
  // Tekst z tabulatorami: „\t” → <w:tab/> (w <w:t> Word pokazałby go jako spację). Bez względu
  // na znacznik „tab” — fragment mógł odziedziczyć go razem ze stylem sąsiada.
  String(run.text || "").split("\t").forEach((seg, i) => {
    if (i) r.appendChild(doc.createElementNS(W_NS, "tab"));
    const text = sanitizeXmlText(seg);
    if (!text) return;
    const t = doc.createElementNS(W_NS, "t");
    if (/^\s|\s$/.test(text)) t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    t.textContent = text;
    r.appendChild(t);
  });
  return r;
}

// Kolejne fragmenty z tym samym linkiem trafiają do jednego <w:hyperlink>.
// keep = atrybuty dawnego linku o tym samym odwołaniu (etykietka ekranowa w:tooltip, ramka
// w:tgtFrame, w:docLocation) — dawniej każde pisanie w akapicie z linkiem je gubiło.
const HYPERLINK_KEEP_ATTRS = ["tooltip", "tgtFrame", "docLocation", "history"];
function createHyperlinkElement(doc, link, keep) {
  const h = doc.createElementNS(W_NS, "w:hyperlink");
  if (link.startsWith("#")) h.setAttributeNS(W_NS, "w:anchor", link.slice(1));
  else if (link.startsWith("rel:")) h.setAttributeNS(R_NS, "r:id", link.slice(4));
  else h.setAttribute("dwb-href", link); // nowy adres — powiązanie dopisze finalizeComposeParts
  h.setAttributeNS(W_NS, "w:history", "1");
  if (keep) HYPERLINK_KEEP_ATTRS.forEach((k) => { if (keep[k] != null) h.setAttributeNS(W_NS, `w:${k}`, keep[k]); });
  return h;
}
function originalHyperlinkAttrs(pEl) {
  const out = new Map();
  Array.from(pEl.childNodes).forEach((n) => {
    if (n.localName !== "hyperlink" || n.namespaceURI !== W_NS) return;
    const ref = hyperlinkRef(n);
    if (!ref || out.has(ref)) return;
    const attrs = {};
    HYPERLINK_KEEP_ATTRS.forEach((k) => { const v = n.getAttributeNS(W_NS, k); if (v != null && n.hasAttributeNS(W_NS, k)) attrs[k] = v; });
    out.set(ref, attrs);
  });
  return out;
}

// Oryginalne właściwości fragmentów akapitu (przed przepisaniem): model akapitu zna tylko część
// formatu Worda (pogrubienie, kursywa, kolor, krój, rozmiar…), a w:rPr ma też język, krój motywu,
// odstępy, kapitaliki, cień… Dawniej każdy fragment pisanego akapitu był budowany od zera i to
// wszystko przepadało. Teraz fragment o tym samym formacie dostaje KOPIĘ oryginału; inny format
// pogrubienia/kursywy/podkreślenia/koloru — kopię fragmentu o tym samym kroju i rozmiarze
// z poprawionymi tylko tymi cechami; zupełnie nowy format — budowany od zera, jak dawniej.
function originalRunProps(pEl) {
  const out = [];
  const walk = (parent) => Array.from(parent.childNodes).forEach((n) => {
    if (n.namespaceURI !== W_NS) return;
    if (n.localName === "hyperlink") return walk(n);
    if (n.localName !== "r" || !Array.from(n.childNodes).some((c) => c.localName === "t" && c.textContent)) return;
    const rPr = Array.from(n.childNodes).find((c) => c.localName === "rPr" && c.namespaceURI === W_NS);
    out.push({ style: xmlRunStyle(n), rPr: rPr || null });
  });
  walk(pEl);
  return out;
}
const TOGGLE_TAGS = { bold: ["b", "bCs"], italic: ["i", "iCs"], strike: ["strike"], underline: ["u"], color: ["color"], highlight: ["highlight", "shd"], vertAlign: ["vertAlign"] };
const RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"];
function runFromOriginal(doc, run, originals) {
  if (run.link) return null; // link: styl Hiperłącze zakłada createRunElement
  const same = originals.find((o) => !o.style.link && runsStyleEqual(o.style, run));
  const near = same || originals.find((o) => !o.style.link && (o.style.fontFamily || "") === (run.fontFamily || "") && (o.style.fontSize || "") === (run.fontSize || ""));
  if (!near) return null;
  const fresh = createRunElement(doc, run);
  if (!near.rPr) return same ? fresh : null; // oryginał bez w:rPr — nowy i tak ma tylko to, co trzeba
  const rPr = near.rPr.cloneNode(true);
  if (!same) {
    // podmień tylko cechy przełączane, resztę oryginału zostaw
    const freshPr = Array.from(fresh.childNodes).find((c) => c.localName === "rPr");
    for (const [k, tags] of Object.entries(TOGGLE_TAGS)) {
      const want = k === "color" ? parseCssColorToWordHex(run.color) || null : k === "highlight" ? normHighlight(run.highlight) : k === "vertAlign" ? normVertAlign(run.vertAlign) : run[k];
      const had = k === "color" ? parseCssColorToWordHex(near.style.color) || null : k === "highlight" ? normHighlight(near.style.highlight) : k === "vertAlign" ? normVertAlign(near.style.vertAlign) : near.style[k];
      if ((k === "color" || k === "highlight" || k === "vertAlign") ? (want || "") === (had || "") : !!want === !!had) continue;
      Array.from(rPr.childNodes).filter((c) => tags.includes(c.localName)).forEach((c) => rPr.removeChild(c));
      if (freshPr) Array.from(freshPr.childNodes).filter((c) => tags.includes(c.localName)).forEach((c) => rPr.appendChild(c.cloneNode(true)));
    }
    Array.from(rPr.childNodes).sort((x, y) => (RPR_ORDER.indexOf(x.localName) + 1 || 99) - (RPR_ORDER.indexOf(y.localName) + 1 || 99)).forEach((c) => rPr.appendChild(c));
  }
  const old = Array.from(fresh.childNodes).find((c) => c.localName === "rPr");
  if (old) fresh.replaceChild(doc.importNode(rPr, true), old);
  else fresh.insertBefore(doc.importNode(rPr, true), fresh.firstChild);
  return fresh;
}

function applyRunsToParagraphXml(pEl, runs) {
  const originals = originalRunProps(pEl);
  const linkAttrs = originalHyperlinkAttrs(pEl);
  // pola-wyspy wracają z listy fragmentów (w swoich miejscach) — stare kontrolki precz
  Array.from(pEl.childNodes).forEach((n) => {
    if (n.namespaceURI !== W_NS) return;
    if (n.localName === "sdt" && typeof formIslandSdt === "function" && formIslandSdt(n)) pEl.removeChild(n);
    else if (n.localName === "commentRangeStart" || n.localName === "commentRangeEnd") pEl.removeChild(n);
    else if (n.localName === "bookmarkStart" || n.localName === "bookmarkEnd") pEl.removeChild(n);
  });
  clearParagraphRuns(pEl);
  const doc = pEl.ownerDocument;
  let hl = null;
  (runs || []).forEach((run) => {
    if (run.island) {
      hl = null;
      const frag = new DOMParser().parseFromString(run.island, "application/xml").documentElement;
      if (frag && frag.namespaceURI === W_NS && ["sdt", "r", "commentRangeStart", "commentRangeEnd", "bookmarkStart", "bookmarkEnd"].includes(frag.localName)) {
        const node = doc.importNode(frag, true);
        // deklaracja xmlns:w z zapisu wyspy — dokument ma ją w korzeniu (bez tego każdy znacznik
        // dostawał w pliku własną kopię przestrzeni nazw)
        if (node.getAttributeNS("http://www.w3.org/2000/xmlns/", "w") === W_NS) node.removeAttributeNS("http://www.w3.org/2000/xmlns/", "w");
        pEl.appendChild(node);
      }
      return;
    }
    if (run.break) {
      hl = null;
      const brRun = doc.createElementNS(W_NS, "r");
      brRun.appendChild(doc.createElementNS(W_NS, "br"));
      pEl.appendChild(brRun);
      return;
    }
    if (!run.text) return;
    if (!run.link) { hl = null; pEl.appendChild(runFromOriginal(doc, run, originals) || createRunElement(doc, run)); return; }
    if (!hl || hl._dwbLink !== run.link) {
      hl = createHyperlinkElement(doc, run.link, linkAttrs.get(run.link));
      hl._dwbLink = run.link;
      pEl.appendChild(hl);
    }
    hl.appendChild(createRunElement(doc, run));
  });
}

async function extractParagraphRunsFromDocx(bytes) {
  if (!window.JSZip || !bytes) return [];
  const doc = await getDocumentXmlDom(bytes); // wspólny parse z extractParagraphTextsFromDocx
  if (!doc) return [];
  return collectParagraphElements(doc.documentElement, "all").map(extractRunsFromParagraphXml);
}

function applyRunsToPreviewParagraph(pEl, runs) {
  if (!pEl) return;
  pEl.replaceChildren();
  (runs || []).forEach((run) => {
    if (run.break) {
      pEl.appendChild(document.createElement("br"));
      return;
    }
    if (!run.text) return;
    const css = runStyleToCss(run);
    let node;
    if (css) {
      node = document.createElement("span");
      node.setAttribute("style", css);
      node.textContent = run.text;
    } else {
      node = document.createTextNode(run.text);
    }
    if (run.link) {
      const a = document.createElement("a");
      a.dataset.dwbLink = run.link;
      a.setAttribute("href", run.link.startsWith("rel:") ? (docLinkHrefs.get(run.link) || "#") : run.link);
      a.className = "dwb-link";
      a.appendChild(node);
      node = a;
    }
    pEl.appendChild(node);
  });
}

// „rel:rIdN” → adres (z oznaczania akapitów) — do odbudowy <a> w podglądzie.
const docLinkHrefs = new Map();

function runStyleHasProps(style) {
  return !!(style?.bold || style?.italic || style?.underline || style?.strike || style?.color || style?.fontFamily || style?.fontSize || style?.highlight || normVertAlign(style?.vertAlign))
    || styleTurnsOff(style); // np. pogrubienie WYŁĄCZONE (Ctrl/⌘+B bez zaznaczenia) — też format do wstawienia
}

// Format „wyłączony” dla dalszego pisania (pogrubienie/kursywa/podkreślenie = false).
function styleTurnsOff(style) {
  return style?.bold === false || style?.italic === false || style?.underline === false || style?.strike === false || style?.vertAlign === "baseline";
}

// Wyjście z fragmentów tekstu (<span>, <b>…) w miejscu kursora aż do akapitu: fragmenty dzielone
// są na „przed” i „za”, a kursor ląduje między nimi, bezpośrednio w akapicie. Potrzebne, gdy
// dalsze pisanie ma WYŁĄCZYĆ format otaczającego fragmentu — podkreślenia rodzica nie da się
// zdjąć stylem dziecka (text-decoration się „rysuje przez”), a pogrubienie z <b> też by zostało.
// Linków i wysp nie rozcina (pisanie w linku zostaje w linku). Zwraca znacznik miejsca.
function breakOutOfRunsAtCaret(range, rootEl) {
  const marker = document.createTextNode("");
  range.insertNode(marker);
  const SPLITTABLE = new Set(["span", "b", "strong", "i", "em", "u", "s", "strike", "sup", "sub"]);
  for (let parent = marker.parentElement; parent && parent !== rootEl && SPLITTABLE.has(parent.localName) && parent.getAttribute("contenteditable") !== "false"; parent = marker.parentElement) {
    const tail = parent.cloneNode(false);
    while (marker.nextSibling) tail.appendChild(marker.nextSibling);
    parent.after(marker);
    marker.after(tail);
    if (!tail.textContent) tail.remove();
    if (!parent.textContent) parent.remove();
  }
  return marker;
}

function accumulateElementStyle(el, style) {
  if (!el || el.nodeType !== 1) return;
  const tag = el.localName.toLowerCase();
  if (tag === "span") Object.assign(style, parseSpanStyle(el.getAttribute("style") || ""));
  if (tag === "b" || tag === "strong") style.bold = true;
  if (tag === "i" || tag === "em") style.italic = true;
  if (tag === "u") style.underline = true;
  if (tag === "s" || tag === "strike") style.strike = true;
  if (tag === "sup") style.vertAlign = "superscript";
  if (tag === "sub") style.vertAlign = "subscript";
}

// Formatowanie, które dostanie tekst wpisany w miejscu kursora: jak w Wordzie — znaku PRZED
// kursorem (na początku akapitu: pierwszego znaku za nim). Najgłębszy element wygrywa, tak samo
// jak w CSS i przy zapisie (extractRunsFromPreviewParagraph). Dawniej przodkowie nadpisywali
// potomków: po zmianie rozmiaru fragmentu wewnątrz tekstu z innym rozmiarem (zagnieżdżone
// <span>) kursor „widział” rozmiar zewnętrzny i dalsze pisanie dziedziczyło zły rozmiar/krój.
function getInheritedRunStyleAtCaret(rootEl) {
  const sel = window.getSelection();
  if (!sel?.rangeCount || !rootEl) return {};
  const range = sel.getRangeAt(0);
  if (!rootEl.contains(range.startContainer)) return {};
  const src = caretStyleSourceNode(rootEl, range.startContainer, range.startOffset);
  const style = {};
  const chain = [];
  for (let el = src?.nodeType === 1 ? src : src?.parentElement; el && el !== rootEl && rootEl.contains(el); el = el.parentElement) chain.push(el);
  for (let i = chain.length - 1; i >= 0; i--) accumulateElementStyle(chain[i], style); // od zewnątrz do środka
  return style;
}

// Węzeł tekstu, którego format „dziedziczy” kursor (container, offset) w akapicie rootEl.
// Pomija tekst wysp (odnośnik przypisu, pole formularza) i znaki-pomocnicze kursora (U+FEFF).
// Wyjątek (zgłoszenie: klik w słowo 8 pt pokazywał 9 pt): kursor na POCZĄTKU słowa, a przed nim
// spacja w innym formacie (klik w lewą połowę pierwszej litery) — liczy się słowo, w które kliknięto,
// i tak samo dostaje format to, co się dopisze przed nim. W środku i na końcu słowa — znak przed.
function caretStyleSourceNode(rootEl, container, offset) {
  const real = (n) => {
    if (!/[^\uFEFF]/.test(n.data)) return false;
    const island = n.parentElement?.closest('[contenteditable="false"]');
    return !(island && island !== rootEl && rootEl.contains(island));
  };
  const lastChar = (n, end) => n.data.slice(0, end).replace(/\uFEFF/g, "").slice(-1);
  const firstChar = (n, start) => n.data.slice(start).replace(/\uFEFF/g, "").charAt(0);
  // znak przed kursorem i znak za nim (z węzłami)
  const at = document.createRange();
  at.setStart(container, offset);
  const walker = document.createTreeWalker(rootEl, NodeFilter.SHOW_TEXT);
  let before = null, beforeCh = "", after = null, afterCh = "";
  if (container.nodeType === 3 && real(container)) {
    if (offset > 0 && lastChar(container, offset)) { before = container; beforeCh = lastChar(container, offset); }
    if (firstChar(container, offset)) { after = container; afterCh = firstChar(container, offset); }
  }
  if (!before || !after) {
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (!real(n) || n === container) continue;
      if (at.comparePoint(n, n.length) <= 0) { if (!before || before !== container) { before = n; beforeCh = lastChar(n, n.length); } continue; }
      if (!after) { after = n; afterCh = firstChar(n, 0); }
      break;
    }
  }
  if (before && after && after !== before && /\s/.test(beforeCh) && /\S/.test(afterCh)) return after;
  return before || after || container;
}

function mergeRunStyles(base, extra) {
  return { ...base, ...Object.fromEntries(Object.entries(extra || {}).filter(([, v]) => v != null && v !== "")) };
}

// opts.turnOff — format WYŁĄCZONY przez użytkownika (Ctrl/⌘+B bez zaznaczenia): tekst wychodzi
// z fragmentu. Dziedziczone „nie” od sąsiada (fragment z w:b w:val="0") dopisuje się normalnie.
function insertStyledTextAtCaret(text, style, rootEl, opts = {}) {
  const sel = window.getSelection();
  if (!sel?.rangeCount) return false;
  const range = sel.getRangeAt(0);
  if (rootEl && !rootEl.contains(range.startContainer)) return false;
  range.deleteContents();
  let css = runStyleToCss(style || {});
  // wyłączony format (bez zaznaczenia: Ctrl/⌘+B w pogrubionym słowie) — tekst poza fragmentem,
  // z pełnym formatem miejsca (krój, rozmiar, kolor) poza wyłączoną cechą
  if (opts.turnOff && styleTurnsOff(opts.turnOff)) {
    const sc0 = range.startContainer;
    const prev0 = sc0.nodeType === 1 ? sc0.childNodes[range.startOffset - 1] : null;
    if (!(prev0?.localName === "span" && prev0.getAttribute("style") === css && prev0.lastChild?.nodeType === 3)) {
      const para = rootEl || range.startContainer.parentElement?.closest("p");
      const marker = breakOutOfRunsAtCaret(range, para);
      range.setStartBefore(marker);
      range.collapse(true);
      marker.remove();
      // poza fragmentami decyduje styl akapitu: jawne „nie” tylko, gdy akapit SAM ma tę cechę
      // (pogrubiony nagłówek) — inaczej każdy dalszy tekst niósłby w pliku zbędne w:b w:val="0"
      if (para) {
        // format „gołego” fragmentu w tym akapicie: docx-preview daje styl akapitu fragmentom
        // (.styl span), nie samemu <p> — próbny pusty <span> pokazuje, co dostałby zwykły tekst
        const probe = document.createElement("span");
        probe.textContent = "x";
        para.appendChild(probe);
        const cs = getComputedStyle(probe);
        const deco = (() => { let d = ""; for (let a = probe; a && a !== para.parentElement; a = a.parentElement) d += " " + (getComputedStyle(a).textDecorationLine || ""); return d; })();
        const flags = { bold: cs.fontWeight === "bold" || parseInt(cs.fontWeight, 10) >= 600, italic: /italic|oblique/.test(cs.fontStyle), underline: /underline/.test(deco), strike: /line-through/.test(deco) };
        probe.remove();
        const style2 = { ...style };
        for (const k of ["bold", "italic", "underline", "strike"]) if (style2[k] === false && !flags[k]) delete style2[k];
        if (style2.vertAlign === "baseline") delete style2.vertAlign; // poza indeksem — zwykły tekst
        style = style2;
        css = runStyleToCss(style);
      }
    }
  }
  // kolejna litera tuż za fragmentem w tym samym stylu — dopisz do niego (dawniej każda litera
  // dostawała własny <span>; zapis i tak je sklejał, ale podgląd puchł przy dłuższym pisaniu)
  const sc = range.startContainer;
  const prevEl = sc.nodeType === 1 ? sc.childNodes[range.startOffset - 1]
    : sc.nodeType === 3 && range.startOffset === sc.length && sc.parentElement?.localName === "span" && !sc.nextSibling ? sc.parentElement : null;
  if (css && prevEl?.localName === "span" && prevEl.getAttribute("style") === css && prevEl.lastChild?.nodeType === 3) {
    const tn = prevEl.lastChild;
    tn.appendData(text);
    range.setStart(tn, tn.length);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    return true;
  }
  let node;
  if (css) {
    const span = document.createElement("span");
    span.setAttribute("style", css);
    span.textContent = text;
    node = span;
  } else {
    node = document.createTextNode(text);
  }
  range.insertNode(node);
  range.setStartAfter(node);
  range.collapse(true);
  sel.removeAllRanges();
  sel.addRange(range);
  return true;
}

function applyRunStyleToSelection(style, rootEl) {
  const sel = window.getSelection();
  if (!sel?.rangeCount || !rootEl?.contains(sel.anchorNode)) return false;
  const range = sel.getRangeAt(0);
  if (range.collapsed) return false;
  const fragment = range.extractContents();
  const css = runStyleToCss(style || {});
  if (!css) {
    range.insertNode(fragment);
    return true;
  }
  const span = document.createElement("span");
  span.setAttribute("style", css);
  span.appendChild(fragment);
  range.insertNode(span);
  range.selectNodeContents(span);
  range.collapse(false);
  sel.removeAllRanges();
  sel.addRange(range);
  return true;
}

// Formatowanie znakowe zaznaczenia (rozmiar, krój…) — także przez kilka akapitów. props: cechy
// albo funkcja (węzeł tekstu → cechy), np. „o stopień większa” dla każdego rozmiaru osobno. Każdy fragment
// tekstu w zaznaczeniu dostaje tę JEDNĄ cechę najgłębiej (we własnym <span> albo w <span>, który
// zawiera tylko ten fragment), reszta jego formatu zostaje. Dawniej całe zaznaczenie szło do
// jednego <span> z formatem z POCZĄTKU zaznaczenia: wewnętrzne fragmenty z własnym rozmiarem
// wygrywały (rozmiar „nie zmieniał się”), a pogrubienie czy krój pierwszego słowa rozlewały się
// na resztę. Zwraca zakres obejmujący sformatowany tekst (zaznaczenie zostaje, jak w Wordzie).
function applyRunPropsToRange(range, props, accept = () => true) {
  if (!range || range.collapsed) return null;
  const root = range.commonAncestorContainer.nodeType === 1 ? range.commonAncestorContainer : range.commonAncestorContainer.parentElement;
  const nodes = [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!range.intersectsNode(n) || !n.length) continue;
    const el = n.parentElement;
    if (!el || el.closest('[contenteditable="false"], .docx-tab') || !accept(n)) continue;
    let from = n === range.startContainer ? range.startOffset : 0;
    let to = n === range.endContainer ? range.endOffset : n.length;
    if (to <= from) continue;
    nodes.push({ n, from, to, css: typeof props === "function" ? runStyleToCss(props(n)) : null });
  }
  if (!nodes.length) return null;
  const fixedCss = typeof props === "function" ? "" : runStyleToCss(props);
  const wrapped = nodes.map(({ n, from, to, css: own }) => {
    const css = own ?? fixedCss;
    let node = n;
    if (to < node.length) node.splitText(to);
    if (from > 0) node = node.splitText(from);
    if (!/[^\uFEFF]/.test(node.data)) return node;
    const parent = node.parentElement;
    if (parent.localName === "span" && parent.childNodes.length === 1 && !parent.className && !parent.dataset.cm && !parent.dataset.ff) {
      // fragment ma już własny <span> (fragment z pliku) — cecha trafia do niego
      const probe = document.createElement("span");
      probe.setAttribute("style", css);
      for (let i = 0; i < probe.style.length; i++) {
        const prop = probe.style[i];
        parent.style.setProperty(prop, probe.style.getPropertyValue(prop));
      }
      return node;
    }
    const span = document.createElement("span");
    span.setAttribute("style", css);
    parent.insertBefore(span, node);
    span.appendChild(node);
    return node;
  });
  const out = document.createRange();
  out.setStart(wrapped[0], 0);
  const last = wrapped[wrapped.length - 1];
  out.setEnd(last, last.length);
  return out;
}

// Krój i rozmiar tekstu tak, jak go widać (styl akapitu, styl znakowy, format fragmentu) —
// to, co Word pokazuje na wstążce. Powiększenie widoku (CSS zoom / transform) nie zmienia
// wartości wyliczonej — oba silniki oddają rozmiar sprzed powiększenia.
function textFormatOfElement(el) {
  if (!el) return null;
  const cs = getComputedStyle(el);
  const family = String(cs.fontFamily || "").split(",")[0].trim().replace(/^["']|["']$/g, "");
  // indeks górny / dolny rysujemy mniejszy, ale w pliku (i w Wordzie) rozmiar tekstu jest ten sam —
  // pokazujemy rozmiar fragmentu (jego własny w:sz albo rozmiar wokół indeksu)
  for (let a = el; a && a.localName !== "p"; a = a.parentElement) {
    if (a.localName !== "sup" && a.localName !== "sub" && !/^(super|sub)$/.test(a.style?.verticalAlign || "")) continue;
    const own = a.localName === "span" ? parseSpanStyle(a.getAttribute("style") || "").fontSize : null;
    if (own) return { family, sizePt: parseFloat(own) };
    const outer = parseFloat(getComputedStyle(a.parentElement).fontSize) || 0;
    return { family, sizePt: Math.round(outer * 0.75 * 2) / 2 };
  }
  const px = parseFloat(cs.fontSize) || 0;
  return { family, sizePt: Math.round(px * 0.75 * 2) / 2 };
}

// Format w zakresie: { family, sizePt } — pusty („”/0) gdy w zaznaczeniu jest kilka różnych
// (jak puste pole kroju/rozmiaru w Wordzie); families / sizes = wszystkie spotkane (do podpowiedzi).
function textFormatOfRange(range, accept = () => true) {
  if (!range) return null;
  const sc = range.startContainer;
  if (range.collapsed) {
    const p = (sc.nodeType === 1 ? sc : sc.parentElement)?.closest?.("p");
    const src = p ? caretStyleSourceNode(p, sc, range.startOffset) : sc;
    const f = textFormatOfElement(src?.nodeType === 1 ? src : src?.parentElement);
    return f && { ...f, families: [f.family], sizes: [f.sizePt] };
  }
  const root = range.commonAncestorContainer.nodeType === 1 ? range.commonAncestorContainer : range.commonAncestorContainer.parentElement;
  const families = new Set();
  const sizes = new Set();
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let seen = 0;
  for (let n = walker.nextNode(); n && seen < 4000; n = walker.nextNode()) {
    if (!range.intersectsNode(n) || !/[^\s\uFEFF]/.test(n.data) || !accept(n)) continue;
    if (n === range.endContainer && range.endOffset === 0) continue;
    if (n === range.startContainer && range.startOffset >= n.length) continue;
    const f = textFormatOfElement(n.parentElement);
    if (!f) continue;
    families.add(f.family);
    sizes.add(f.sizePt);
    seen++;
  }
  if (!seen) { const r = range.cloneRange(); r.collapse(true); return textFormatOfRange(r, accept); } // same spacje
  return {
    family: families.size === 1 ? [...families][0] : "",
    sizePt: sizes.size === 1 ? [...sizes][0] : 0,
    families: [...families],
    sizes: [...sizes].sort((a, b) => a - b),
  };
}

function fontSizePtFromStyle(style) {
  if (!style?.fontSize) return "";
  const pt = cssFontSizeToPt(style.fontSize);
  return pt ? String(pt) : "";
}

function getTextBeforeCaret(rootEl) {
  const sel = window.getSelection();
  if (!sel?.rangeCount || !rootEl?.contains(sel.anchorNode)) return "";
  const range = sel.getRangeAt(0);
  if (!range.collapsed) return "";
  const pre = range.cloneRange();
  pre.selectNodeContents(rootEl);
  pre.setEnd(range.startContainer, range.startOffset);
  return pre.toString();
}

function replaceTextEndingBeforeCaret(rootEl, deleteLen, insertText, style) {
  const sel = window.getSelection();
  if (!sel?.rangeCount || deleteLen < 1) return false;
  const endRange = sel.getRangeAt(0);
  if (!rootEl.contains(endRange.startContainer)) return false;
  const endOffset = getTextBeforeCaret(rootEl).length;
  const startOffset = Math.max(0, endOffset - deleteLen);
  const startPos = resolveTextPosition(rootEl, startOffset);
  const endPos = resolveTextPosition(rootEl, endOffset);
  if (!startPos || !endPos) return false;
  const delRange = document.createRange();
  delRange.setStart(startPos.node, startPos.offset);
  delRange.setEnd(endPos.node, endPos.offset);
  delRange.deleteContents();
  const css = runStyleToCss(style || {});
  const node = css
    ? Object.assign(document.createElement("span"), { textContent: insertText })
    : document.createTextNode(insertText);
  if (css) node.setAttribute("style", css);
  delRange.insertNode(node);
  delRange.setStartAfter(node);
  delRange.collapse(true);
  sel.removeAllRanges();
  sel.addRange(delRange);
  return true;
}

function resolveTextPosition(rootEl, charOffset) {
  let remaining = charOffset;
  const walker = document.createTreeWalker(rootEl, NodeFilter.SHOW_TEXT);
  let node = walker.nextNode();
  while (node) {
    const len = (node.textContent || "").length;
    if (remaining <= len) return { node, offset: remaining };
    remaining -= len;
    node = walker.nextNode();
  }
  return null;
}
