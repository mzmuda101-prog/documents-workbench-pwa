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
  const texts = pEl.getElementsByTagNameNS(W_NS, "t");
  let out = "";
  for (let i = 0; i < texts.length; i++) out += texts[i].textContent || "";
  return out;
}

function setParagraphText(pEl, text) {
  const runs = Array.from(pEl.getElementsByTagNameNS(W_NS, "r"));
  const sanitized = sanitizeXmlText(text);
  if (!runs.length) {
    const r = pEl.ownerDocument.createElementNS(W_NS, "r");
    const t = pEl.ownerDocument.createElementNS(W_NS, "t");
    t.textContent = sanitized;
    r.appendChild(t);
    pEl.appendChild(r);
    return;
  }
  let firstT = runs[0].getElementsByTagNameNS(W_NS, "t")[0];
  if (!firstT) {
    firstT = pEl.ownerDocument.createElementNS(W_NS, "t");
    runs[0].appendChild(firstT);
  }
  firstT.textContent = sanitized;
  for (let i = 1; i < runs.length; i++) {
    const ts = runs[i].getElementsByTagNameNS(W_NS, "t");
    for (let j = 0; j < ts.length; j++) ts[j].textContent = "";
  }
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

function applyReplaceInXml(xml, edit, scope) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const textNodes = collectTextNodes(doc, scope);
  let count = 0;
  const find = edit.find;
  const repl = edit.replace ?? "";
  const useRegex = !!edit.regex;

  textNodes.forEach((tEl) => {
    const raw = tEl.textContent || "";
    if (!raw) return;
    let next = raw;
    if (useRegex) {
      try {
        const re = new RegExp(find, "g");
        const replaced = raw.replace(re, repl);
        if (replaced !== raw) {
          next = replaced;
          count += (raw.match(re) || []).length;
        }
      } catch (_) { /* invalid regex */ }
    } else if (raw.includes(find)) {
      next = raw.split(find).join(repl);
      count += raw.split(find).length - 1;
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
  (items || []).forEach(({ index, text }) => {
    const p = paragraphs[index];
    if (!p) return;
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

function applyEditToXml(xml, edit) {
  const scope = edit.scope || "all";
  if (edit.op === "paragraphBatch") return applyParagraphBatchInXml(xml, edit.items);
  if (edit.op === "case" || edit.op === "trim" || edit.op === "affix") return applyParagraphTransformInXml(xml, edit, scope);
  return applyReplaceInXml(xml, edit, scope);
}

function recordPendingEdit(edit) {
  pendingDocEdits.push({ ...edit, ts: Date.now() });
}

async function buildPatchedDocx(bytes, edits) {
  if (!window.JSZip) throw new Error("JSZip missing");
  const zip = await window.JSZip.loadAsync(bytes);
  const docFile = zip.file("word/document.xml");
  if (!docFile) throw new Error("word/document.xml missing");
  let xml = await docFile.async("string");
  let total = 0;
  for (const edit of edits) {
    const normalized = edit.op ? edit : { ...edit, op: "replace" };
    const res = applyEditToXml(xml, normalized);
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
