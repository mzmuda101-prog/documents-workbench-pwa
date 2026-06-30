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
    if (key === "font-weight" && (val === "bold" || parseInt(val, 10) >= 600)) style.bold = true;
    if (key === "font-style" && val === "italic") style.italic = true;
    if (key === "text-decoration" && val.includes("underline")) style.underline = true;
    if (key === "color") style.color = val;
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
  if (run.italic) parts.push("font-style:italic");
  if (run.underline) parts.push("text-decoration:underline");
  if (run.color) parts.push(`color:${run.color}`);
  if (run.fontFamily) parts.push(`font-family:"${run.fontFamily}"`);
  if (run.fontSize) parts.push(`font-size:${run.fontSize}`);
  return parts.join(";");
}

function runsStyleEqual(a, b) {
  return !!a.bold === !!b.bold
    && !!a.italic === !!b.italic
    && !!a.underline === !!b.underline
    && (a.color || "") === (b.color || "")
    && (a.fontFamily || "") === (b.fontFamily || "")
    && (a.fontSize || "") === (b.fontSize || "");
}

function mergeAdjacentRuns(runs) {
  const out = [];
  (runs || []).forEach((run) => {
    if (run.break) {
      out.push({ break: true });
      return;
    }
    const prev = out[out.length - 1];
    if (prev && !prev.break && runsStyleEqual(prev, run)) {
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
      const text = node.textContent || "";
      if (text) runs.push({ text, ...inherited });
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
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
    return run.text === other.text && runsStyleEqual(run, other);
  });
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

function extractRunsFromParagraphXml(pEl) {
  const runs = [];
  getParagraphRunElements(pEl).forEach((r) => {
    const br = Array.from(r.childNodes).find((n) => n.localName === "br" && n.namespaceURI === W_NS);
    if (br) {
      runs.push({ break: true });
      return;
    }
    let text = "";
    Array.from(r.childNodes).forEach((n) => {
      if (n.localName === "t" && n.namespaceURI === W_NS) text += n.textContent || "";
    });
    if (!text) return;
    const run = { text };
    const rPr = Array.from(r.childNodes).find((n) => n.localName === "rPr" && n.namespaceURI === W_NS);
    if (rPr) {
      if (Array.from(rPr.childNodes).some((n) => n.localName === "b")) run.bold = true;
      if (Array.from(rPr.childNodes).some((n) => n.localName === "i")) run.italic = true;
      if (Array.from(rPr.childNodes).some((n) => n.localName === "u")) run.underline = true;
      const colorEl = Array.from(rPr.childNodes).find((n) => n.localName === "color");
      const hex = colorEl ? getWVal(colorEl) : null;
      if (hex) run.color = `#${hex.replace(/^#/, "")}`;
      const fonts = Array.from(rPr.childNodes).find((n) => n.localName === "rFonts");
      if (fonts) {
        run.fontFamily = fonts.getAttributeNS(W_NS, "ascii") || fonts.getAttributeNS(W_NS, "hAnsi") || getWVal(fonts);
      }
      const sz = Array.from(rPr.childNodes).find((n) => n.localName === "sz");
      if (sz) {
        const half = parseInt(getWVal(sz) || "0", 10);
        if (half) run.fontSize = `${half / 2}pt`;
      }
    }
    runs.push(run);
  });
  return mergeAdjacentRuns(runs);
}

function createRunElement(doc, run) {
  const r = doc.createElementNS(W_NS, "r");
  const rPr = doc.createElementNS(W_NS, "rPr");
  let hasPr = false;
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
  if (hasPr) r.appendChild(rPr);
  const t = doc.createElementNS(W_NS, "t");
  const text = sanitizeXmlText(run.text || "");
  if (/^\s|\s$/.test(text)) t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
  t.textContent = text;
  r.appendChild(t);
  return r;
}

function applyRunsToParagraphXml(pEl, runs) {
  clearParagraphRuns(pEl);
  const doc = pEl.ownerDocument;
  (runs || []).forEach((run) => {
    if (run.break) {
      const brRun = doc.createElementNS(W_NS, "r");
      brRun.appendChild(doc.createElementNS(W_NS, "br"));
      pEl.appendChild(brRun);
      return;
    }
    if (!run.text) return;
    pEl.appendChild(createRunElement(doc, run));
  });
}

async function extractParagraphRunsFromDocx(bytes) {
  if (!window.JSZip) return [];
  const zip = await window.JSZip.loadAsync(bytes);
  const docFile = zip.file("word/document.xml");
  if (!docFile) return [];
  const xml = await docFile.async("string");
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
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
    if (css) {
      const span = document.createElement("span");
      span.setAttribute("style", css);
      span.textContent = run.text;
      pEl.appendChild(span);
    } else {
      pEl.appendChild(document.createTextNode(run.text));
    }
  });
}

function runStyleHasProps(style) {
  return !!(style?.bold || style?.italic || style?.underline || style?.color || style?.fontFamily || style?.fontSize);
}

function accumulateElementStyle(el, style) {
  if (!el || el.nodeType !== 1) return;
  const tag = el.localName.toLowerCase();
  if (tag === "span") Object.assign(style, parseSpanStyle(el.getAttribute("style") || ""));
  if (tag === "b" || tag === "strong") style.bold = true;
  if (tag === "i" || tag === "em") style.italic = true;
  if (tag === "u") style.underline = true;
}

function getInheritedRunStyleAtCaret(rootEl) {
  const sel = window.getSelection();
  if (!sel?.rangeCount || !rootEl) return {};
  const range = sel.getRangeAt(0);
  if (!rootEl.contains(range.startContainer)) return {};
  const style = {};
  let node = range.startContainer;
  const offset = range.startOffset;
  if (node.nodeType === Node.TEXT_NODE && offset === 0) {
    let sib = node.previousSibling;
    while (sib) {
      if (sib.nodeType === 1) {
        accumulateElementStyle(sib, style);
        break;
      }
      if (sib.nodeType === Node.TEXT_NODE && sib.textContent) {
        let p = sib.parentElement;
        while (p && p !== rootEl) {
          accumulateElementStyle(p, style);
          p = p.parentElement;
        }
        break;
      }
      sib = sib.previousSibling;
    }
  }
  if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
  while (node && node !== rootEl) {
    accumulateElementStyle(node, style);
    node = node.parentElement;
  }
  return style;
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
