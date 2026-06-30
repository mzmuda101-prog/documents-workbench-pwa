// ZIP-patch save for DOCX — preserves original package, patches word/document.xml.

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function sanitizeXmlText(s) {
  return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "");
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
    const res = applyReplaceInXml(xml, edit, edit.scope || "all");
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
