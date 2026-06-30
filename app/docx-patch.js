// ZIP-patch save for DOCX — preserves original package, patches word/document.xml.

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function sanitizeXmlText(s) {
  return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "");
}

function collectParagraphElements(rootEl, scope) {
  const paragraphs = [];
  const body = rootEl.getElementsByTagNameNS(W_NS, "body")[0];
  if (!body) return paragraphs;
  const start = scope === "body" ? body : rootEl;

  function walkTable(tbl) {
    const rows = tbl.getElementsByTagNameNS(W_NS, "tr");
    for (let r = 0; r < rows.length; r++) {
      const cells = rows[r].getElementsByTagNameNS(W_NS, "tc");
      for (let c = 0; c < cells.length; c++) walk(cells[c]);
    }
  }

  function walk(node) {
    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i];
      if (child.nodeType !== 1) continue;
      if (child.localName === "p" && child.namespaceURI === W_NS) paragraphs.push(child);
      else if (child.localName === "tbl" && child.namespaceURI === W_NS) walkTable(child);
      else if (child.localName !== "sectPr") walk(child);
    }
  }

  walk(start === body ? body : start);
  return paragraphs;
}

function collectTextNodes(doc, scope) {
  const nodes = [];
  const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
  if (!body) return nodes;
  const root = scope === "body" ? body : doc.documentElement;
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_ELEMENT, null);
  let el = walker.currentNode;
  while (el) {
    if (el.localName === "t" && el.namespaceURI === W_NS) nodes.push(el);
    el = walker.nextNode();
  }
  return nodes;
}

function getParagraphText(pEl) {
  let out = "";
  const runs = pEl.getElementsByTagNameNS(W_NS, "r");
  for (let r = 0; r < runs.length; r++) {
    const children = runs[r].childNodes;
    for (let c = 0; c < children.length; c++) {
      const child = children[c];
      if (child.nodeType !== 1) continue;
      if (child.localName === "t" && child.namespaceURI === W_NS) out += child.textContent || "";
      else if (child.localName === "br" && child.namespaceURI === W_NS) out += "\n";
    }
  }
  return out;
}

function clearParagraphRuns(pEl) {
  Array.from(pEl.getElementsByTagNameNS(W_NS, "r")).forEach((r) => r.parentNode.removeChild(r));
}

function setParagraphText(pEl, text) {
  const sanitized = sanitizeXmlText(text);
  clearParagraphRuns(pEl);
  if (!sanitized) return;
  const parts = sanitized.split("\n");
  const doc = pEl.ownerDocument;
  parts.forEach((part, i) => {
    const r = doc.createElementNS(W_NS, "r");
    if (part) {
      const t = doc.createElementNS(W_NS, "t");
      if (/^\s|\s$/.test(part)) t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
      t.textContent = part;
      r.appendChild(t);
    }
    pEl.appendChild(r);
    if (i < parts.length - 1) {
      const brRun = doc.createElementNS(W_NS, "r");
      brRun.appendChild(doc.createElementNS(W_NS, "br"));
      pEl.appendChild(brRun);
    }
  });
}

function splitParagraphInXml(xml, index, beforeText, afterText, beforeRuns, afterRuns) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[index];
  if (!p) return { xml, count: 0 };
  if (beforeRuns?.length) applyRunsToParagraphXml(p, beforeRuns);
  else setParagraphText(p, beforeText);
  const newP = p.cloneNode(true);
  clearParagraphRuns(newP);
  if (afterRuns?.length) applyRunsToParagraphXml(newP, afterRuns);
  else setParagraphText(newP, afterText);
  if (p.nextSibling) p.parentNode.insertBefore(newP, p.nextSibling);
  else p.parentNode.appendChild(newP);
  return { xml: new XMLSerializer().serializeToString(doc), count: 1 };
}

function getWVal(el) {
  if (!el) return null;
  return el.getAttributeNS(W_NS, "val") ?? el.getAttribute("w:val") ?? el.getAttribute("val");
}

function setWVal(el, value) {
  el.setAttributeNS(W_NS, "val", String(value));
}

function getParagraphListLevel(pEl) {
  const pPr = pEl.getElementsByTagNameNS(W_NS, "pPr")[0];
  const numPr = pPr?.getElementsByTagNameNS(W_NS, "numPr")[0];
  if (!numPr) return -1;
  const ilvl = numPr.getElementsByTagNameNS(W_NS, "ilvl")[0];
  const raw = ilvl ? getWVal(ilvl) : "0";
  const level = parseInt(raw, 10);
  return Number.isFinite(level) ? level : 0;
}

function mergeParagraphInXml(xml, index, mergedRuns) {
  if (index <= 0) return { xml, count: 0 };
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const prev = paragraphs[index - 1];
  const curr = paragraphs[index];
  if (!prev || !curr) return { xml, count: 0 };
  const joinAt = getParagraphText(prev).length;
  if (mergedRuns?.length) applyRunsToParagraphXml(prev, mergedRuns);
  else setParagraphText(prev, getParagraphText(prev) + getParagraphText(curr));
  curr.parentNode.removeChild(curr);
  return { xml: new XMLSerializer().serializeToString(doc), count: 1, joinAt };
}

function changeListLevelInXml(xml, index, delta) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[index];
  if (!p) return { xml, count: 0 };
  let pPr = p.getElementsByTagNameNS(W_NS, "pPr")[0];
  if (!pPr) {
    pPr = doc.createElementNS(W_NS, "pPr");
    p.insertBefore(pPr, p.firstChild);
  }
  const numPr = pPr.getElementsByTagNameNS(W_NS, "numPr")[0];
  if (!numPr) return { xml, count: 0 };
  let ilvl = numPr.getElementsByTagNameNS(W_NS, "ilvl")[0];
  if (!ilvl) {
    ilvl = doc.createElementNS(W_NS, "ilvl");
    numPr.insertBefore(ilvl, numPr.firstChild);
    setWVal(ilvl, "0");
  }
  const current = parseInt(getWVal(ilvl) || "0", 10);
  const next = Math.max(0, Math.min(8, current + delta));
  if (next === current) return { xml, count: 0 };
  setWVal(ilvl, String(next));
  return { xml: new XMLSerializer().serializeToString(doc), count: 1 };
}

async function extractParagraphTextsFromDocx(bytes) {
  if (!window.JSZip) return [];
  const zip = await window.JSZip.loadAsync(bytes);
  const docFile = zip.file("word/document.xml");
  if (!docFile) return [];
  const xml = await docFile.async("string");
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  return collectParagraphElements(doc.documentElement, "all").map(getParagraphText);
}

function countFindInText(text, find, useRegex) {
  if (!text || !find) return 0;
  if (useRegex) {
    try {
      const re = new RegExp(find, "g");
      return (text.match(re) || []).length;
    } catch (_) {
      return 0;
    }
  }
  let n = 0;
  let idx = 0;
  while ((idx = text.indexOf(find, idx)) !== -1) {
    n++;
    idx += find.length || 1;
  }
  return n;
}

function countReplacePreview(texts, edit) {
  const find = edit.find;
  const useRegex = !!edit.regex;
  let hits = 0;
  let paras = 0;
  const samples = [];
  (texts || []).forEach((text, index) => {
    const n = countFindInText(text, find, useRegex);
    if (!n) return;
    hits += n;
    paras++;
    if (samples.length < 8) {
      const snippet = text.length > 72 ? `${text.slice(0, 69)}…` : text;
      samples.push({ index, snippet, count: n });
    }
  });
  return { hits, paras, samples };
}

function applyReplaceInXml(xml, edit, scope, opts = {}) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const textNodes = collectTextNodes(doc, scope);
  let count = 0;
  const find = edit.find;
  const repl = edit.replace ?? "";
  const useRegex = !!edit.regex;
  const maxRepl = opts.maxReplacements ?? Infinity;
  let skip = opts.skipReplacements ?? 0;
  let done = 0;

  textNodes.forEach((tEl) => {
    if (done >= maxRepl) return;
    let raw = tEl.textContent || "";
    if (!raw) return;
    let next = raw;
    if (useRegex) {
      try {
        const re = new RegExp(find, "g");
        next = raw.replace(re, (match) => {
          if (done >= maxRepl) return match;
          if (skip > 0) { skip--; return match; }
          done++;
          count++;
          return repl;
        });
      } catch (_) { /* invalid regex */ }
    } else {
      let idx = 0;
      while (done < maxRepl) {
        const pos = next.indexOf(find, idx);
        if (pos === -1) break;
        if (skip > 0) {
          skip--;
          idx = pos + find.length;
          continue;
        }
        next = next.slice(0, pos) + repl + next.slice(pos + find.length);
        done++;
        count++;
        idx = pos + repl.length;
      }
    }
    if (next !== raw) tEl.textContent = sanitizeXmlText(next);
  });

  if (!count) return { xml, count: 0 };
  return { xml: new XMLSerializer().serializeToString(doc), count };
}

function applyParagraphBatchInXml(xml, items) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  let count = 0;
  (items || []).forEach(({ index, text, runs }) => {
    const p = paragraphs[index];
    if (!p) return;
    if (runs?.length) {
      const current = extractRunsFromParagraphXml(p);
      if (runsEqual(current, runs)) return;
      applyRunsToParagraphXml(p, runs);
      count++;
      return;
    }
    const raw = getParagraphText(p);
    const next = sanitizeXmlText(text);
    if (next === raw) return;
    setParagraphText(p, next);
    count++;
  });
  if (!count) return { xml, count: 0 };
  return { xml: new XMLSerializer().serializeToString(doc), count };
}

function buildParagraphTransformFn(edit) {
  const locale = (typeof I18N !== "undefined" && I18N[currentLang] && I18N[currentLang].locale) || "pl-PL";
  if (edit.op === "case") {
    const m = edit.mode;
    if (m === "upper") return (s) => s.toLocaleUpperCase(locale);
    if (m === "lower") return (s) => s.toLocaleLowerCase(locale);
    return (s) => s.replace(/\p{L}[\p{L}\p{M}]*/gu, (w) => w[0].toLocaleUpperCase(locale) + w.slice(1).toLocaleLowerCase(locale));
  }
  if (edit.op === "trim") {
    const m = edit.mode;
    if (m === "collapse") return (s) => s.replace(/\s+/gu, " ").trim();
    if (m === "hard") return (s) => s.replace(/[\u00A0\u2007\u202F]/g, " ").trim();
    return (s) => s.trim();
  }
  if (edit.op === "affix") {
    const pre = edit.prefix || "";
    const suf = edit.suffix || "";
    return (s) => pre + s + suf;
  }
  return (s) => s;
}

function applyParagraphTransformInXml(xml, edit, scope) {
  const fn = buildParagraphTransformFn(edit);
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const paragraphs = collectParagraphElements(doc.documentElement, scope);
  let count = 0;
  paragraphs.forEach((p) => {
    const raw = getParagraphText(p);
    if (!raw) return;
    let next;
    try { next = fn(raw); } catch { return; }
    if (typeof next !== "string" || next === raw) return;
    setParagraphText(p, next);
    count++;
  });
  if (!count) return { xml, count: 0 };
  return { xml: new XMLSerializer().serializeToString(doc), count };
}

function applyPlaceholderFillInXml(xml, values, scope) {
  let current = xml;
  let count = 0;
  Object.entries(values || {}).forEach(([name, val]) => {
    if (val == null || val === "") return;
    const safeName = String(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const res = applyReplaceInXml(current, {
      op: "replace",
      find: `\\{\\{\\s*${safeName}\\s*\\}\\}`,
      replace: sanitizeXmlText(String(val)),
      regex: true,
      scope: scope || "all",
    });
    current = res.xml;
    count += res.count;
  });
  if (!count) return { xml, count: 0 };
  return { xml: current, count };
}

function applySnippetExpandInXml(xml, snippetMap, scope) {
  let current = xml;
  let count = 0;
  const resolved = resolveSnippetMapBodies(snippetMap);
  const names = Object.keys(resolved).sort((a, b) => b.length - a.length);
  names.forEach((name) => {
    const safeName = String(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const res = applyReplaceInXml(current, {
      op: "replace",
      find: `!${safeName}\\b`,
      replace: sanitizeXmlText(String(resolved[name])),
      regex: true,
      scope: scope || "all",
    });
    current = res.xml;
    count += res.count;
  });
  if (!count) return { xml, count: 0 };
  return { xml: current, count };
}

function applyEditToXml(xml, edit, opts = {}) {
  const scope = edit.scope || "all";
  if (edit.op === "paragraphBatch") return applyParagraphBatchInXml(xml, edit.items);
  if (edit.op === "placeholderFill") return applyPlaceholderFillInXml(xml, edit.values, scope);
  if (edit.op === "snippetExpand") return applySnippetExpandInXml(xml, edit.snippets, scope);
  if (edit.op === "splitParagraph") return splitParagraphInXml(xml, edit.index, edit.before, edit.after, edit.beforeRuns, edit.afterRuns);
  if (edit.op === "mergeParagraph") return mergeParagraphInXml(xml, edit.index, edit.mergedRuns);
  if (edit.op === "listLevel") return changeListLevelInXml(xml, edit.index, edit.delta);
  if (edit.op === "case" || edit.op === "trim" || edit.op === "affix") return applyParagraphTransformInXml(xml, edit, scope);
  return applyReplaceInXml(xml, edit, scope, opts);
}

function recordPendingEdit(edit) {
  pendingDocEdits.push({ ...edit, ts: Date.now() });
}

async function buildPatchedDocx(bytes, edits, lastEditOpts = {}) {
  if (!window.JSZip) throw new Error("JSZip missing");
  const zip = await window.JSZip.loadAsync(bytes);
  const docFile = zip.file("word/document.xml");
  if (!docFile) throw new Error("word/document.xml missing");
  let xml = await docFile.async("string");
  let total = 0;
  const list = edits || [];
  for (let i = 0; i < list.length; i++) {
    const normalized = list[i].op ? list[i] : { ...list[i], op: "replace" };
    const opts = i === list.length - 1 ? lastEditOpts : {};
    const res = applyEditToXml(xml, normalized, opts);
    xml = res.xml;
    total += res.count;
  }
  zip.file("word/document.xml", xml);
  const out = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  return { bytes: out, changeCount: total };
}
