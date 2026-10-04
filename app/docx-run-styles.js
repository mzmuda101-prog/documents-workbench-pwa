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
    if (key === "background-color" || key === "background") { const h = normHighlight(val); if (h) style.highlight = h; }
    if (key === "font-family") style.fontFamily = val.replace(/^["']|["']$/g, "").split(",")[0].trim();
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
  return parts.join(";");
}

function runsStyleEqual(a, b) {
  return !!a.bold === !!b.bold
    && !!a.italic === !!b.italic
    && !!a.underline === !!b.underline
    && !!a.strike === !!b.strike
    && (parseCssColorToWordHex(a.color) || "") === (parseCssColorToWordHex(b.color) || "") // „rgb(5, 99, 193)” z podglądu = „#0563C1” z pliku
    && (a.fontFamily || "") === (b.fontFamily || "")
    && (a.fontSize || "") === (b.fontSize || "")
    && (a.link || "") === (b.link || "")
    && normHighlight(a.highlight) === normHighlight(b.highlight);
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
      runs.push({ island: docIslandXml.get(node.dataset.cm), text: "" });
      return;
    }
    const tag = node.localName.toLowerCase();
    if (tag === "br") {
      runs.push({ break: true });
      return;
    }
    const style = { ...inherited };
    if (tag === "span") Object.assign(style, parseSpanStyle(node.getAttribute("style") || ""));
    if (tag === "b" || tag === "strong") style.bold = true;
    if (tag === "i" || tag === "em") style.italic = true;
    if (tag === "u") style.underline = true;
    if (tag === "s" || tag === "strike") style.strike = true;
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

function extractRunsFromParagraphXml(pEl) {
  const runs = [];
  paragraphXmlParts(pEl).forEach((r) => {
    if (r.localName === "sdt") {
      runs.push({ island: new XMLSerializer().serializeToString(r), text: ffText(ffKid(r, "sdtContent")) });
      return;
    }
    // komentarz: początek/koniec zakresu i fragment z odwołaniem — „wyspy” o zerowej długości
    if (r.localName !== "r" || isCommentReferenceRun(r) || noteRunKind(r)) {
      runs.push({ island: new XMLSerializer().serializeToString(r), text: "" });
      return;
    }
    const style = {};
    if (r.parentNode?.localName === "hyperlink") {
      const ref = hyperlinkRef(r.parentNode);
      if (ref) style.link = ref;
    }
    const rPr = Array.from(r.childNodes).find((n) => n.localName === "rPr" && n.namespaceURI === W_NS);
    if (rPr) {
      if (Array.from(rPr.childNodes).some((n) => n.localName === "b")) style.bold = true;
      if (Array.from(rPr.childNodes).some((n) => n.localName === "i")) style.italic = true;
      if (Array.from(rPr.childNodes).some((n) => n.localName === "u")) style.underline = true;
      const strikeEl = Array.from(rPr.childNodes).find((n) => n.localName === "strike");
      if (strikeEl && !/^(0|false|off)$/i.test(getWVal(strikeEl) || "")) style.strike = true;
      const colorEl = Array.from(rPr.childNodes).find((n) => n.localName === "color");
      const hex = colorEl ? getWVal(colorEl) : null;
      if (hex) style.color = `#${hex.replace(/^#/, "")}`;
      const fonts = Array.from(rPr.childNodes).find((n) => n.localName === "rFonts");
      if (fonts) {
        style.fontFamily = fonts.getAttributeNS(W_NS, "ascii") || fonts.getAttributeNS(W_NS, "hAnsi") || getWVal(fonts);
      }
      const hl = Array.from(rPr.childNodes).find((n) => n.localName === "highlight");
      const shd = Array.from(rPr.childNodes).find((n) => n.localName === "shd");
      const hlVal = hl ? getWVal(hl) : null;
      const shdFill = shd ? (shd.getAttributeNS(W_NS, "fill") || shd.getAttribute("w:fill")) : null;
      if (hlVal && hlVal !== "none") style.highlight = normHighlight(hlVal);
      else if (shdFill && shdFill !== "auto") style.highlight = normHighlight(`#${shdFill}`);
      const sz = Array.from(rPr.childNodes).find((n) => n.localName === "sz");
      if (sz) {
        const half = parseInt(getWVal(sz) || "0", 10);
        if (half) style.fontSize = `${half / 2}pt`;
      }
    }
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
  if (run.bold) {
    const b = doc.createElementNS(W_NS, "b");
    setWVal(b, "1");
    rPr.appendChild(b);
    hasPr = true;
  }
  if (run.italic) {
    const i = doc.createElementNS(W_NS, "i");
    setWVal(i, "1");
    rPr.appendChild(i);
    hasPr = true;
  }
  if (run.strike) {
    const st = doc.createElementNS(W_NS, "strike");
    setWVal(st, "1");
    rPr.appendChild(st);
    hasPr = true;
  }
  if (run.underline) {
    const u = doc.createElementNS(W_NS, "u");
    setWVal(u, "single");
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
    rf.setAttributeNS(W_NS, "ascii", run.fontFamily);
    rf.setAttributeNS(W_NS, "hAnsi", run.fontFamily);
    rf.setAttributeNS(W_NS, "cs", run.fontFamily);
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
  if (hasPr) {
    // kolejność dzieci w:rPr wg schematu Worda (dawniej u przed color, rFonts po color)
    const order = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "strike", "color", "sz", "szCs", "highlight", "u", "shd"];
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
function createHyperlinkElement(doc, link) {
  const h = doc.createElementNS(W_NS, "w:hyperlink");
  if (link.startsWith("#")) h.setAttributeNS(W_NS, "w:anchor", link.slice(1));
  else if (link.startsWith("rel:")) h.setAttributeNS(R_NS, "r:id", link.slice(4));
  else h.setAttribute("dwb-href", link); // nowy adres — powiązanie dopisze finalizeComposeParts
  h.setAttributeNS(W_NS, "w:history", "1");
  return h;
}

function applyRunsToParagraphXml(pEl, runs) {
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
      if (frag && frag.namespaceURI === W_NS && ["sdt", "r", "commentRangeStart", "commentRangeEnd", "bookmarkStart", "bookmarkEnd"].includes(frag.localName)) pEl.appendChild(doc.importNode(frag, true));
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
    if (!run.link) { hl = null; pEl.appendChild(createRunElement(doc, run)); return; }
    if (!hl || hl._dwbLink !== run.link) {
      hl = createHyperlinkElement(doc, run.link);
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
  return !!(style?.bold || style?.italic || style?.underline || style?.strike || style?.color || style?.fontFamily || style?.fontSize || style?.highlight);
}

function accumulateElementStyle(el, style) {
  if (!el || el.nodeType !== 1) return;
  const tag = el.localName.toLowerCase();
  if (tag === "span") Object.assign(style, parseSpanStyle(el.getAttribute("style") || ""));
  if (tag === "b" || tag === "strong") style.bold = true;
  if (tag === "i" || tag === "em") style.italic = true;
  if (tag === "u") style.underline = true;
  if (tag === "s" || tag === "strike") style.strike = true;
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
function caretStyleSourceNode(rootEl, container, offset) {
  const real = (n) => {
    if (!/[^\uFEFF]/.test(n.data)) return false;
    const island = n.parentElement?.closest('[contenteditable="false"]');
    return !(island && island !== rootEl && rootEl.contains(island));
  };
  if (container.nodeType === 3 && offset > 0 && real(container)) return container;
  // pozycja między węzłami albo na początku węzła tekstu — ostatni tekst PRZED nią,
  // a gdy go nie ma (początek akapitu) — pierwszy za nią
  const at = document.createRange();
  at.setStart(container, offset);
  const walker = document.createTreeWalker(rootEl, NodeFilter.SHOW_TEXT);
  let before = null;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!real(n)) continue;
    if (n !== container && at.comparePoint(n, n.length) <= 0) { before = n; continue; }
    return before || n;
  }
  return before || container;
}

function mergeRunStyles(base, extra) {
  return { ...base, ...Object.fromEntries(Object.entries(extra || {}).filter(([, v]) => v != null && v !== "")) };
}

function insertStyledTextAtCaret(text, style, rootEl) {
  const sel = window.getSelection();
  if (!sel?.rangeCount) return false;
  const range = sel.getRangeAt(0);
  if (rootEl && !rootEl.contains(range.startContainer)) return false;
  range.deleteContents();
  const css = runStyleToCss(style || {});
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
