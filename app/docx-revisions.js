// docx-revisions.js — recenzja: śledzone zmiany, komentarze, przypisy (odczyt + Akceptuj/Odrzuć).
//
// Podgląd (docx-preview bez renderChanges) pokazuje dokument „po zaakceptowaniu”: wstawienia
// widać, usunięć nie. Ten moduł czyta XML bezpośrednio, więc widzi wszystko, i zapisuje
// Akceptuj / Odrzuć / Usuń komentarze we wszystkich częściach pliku (treść, nagłówki,
// stopki, przypisy). Operacja „revisions” jest wołana z buildPatchedDocx (docx-patch.js).
//
// Każda zmiana w Wordzie ma w:id — po nim działa akceptowanie/odrzucanie pojedynczych pozycji.

const REV_CHANGE_TAGS = ["rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "tcPrChange", "trPrChange", "tblGridChange", "tblPrExChange", "numberingChange"];
const REV_PARTS_RE = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

function revAttr(el, name) {
  return el.getAttributeNS(W_NS, name) ?? el.getAttribute(`w:${name}`) ?? "";
}

function revKids(el, name) {
  return Array.from(el.getElementsByTagNameNS(W_NS, name));
}

function revChildren(el) {
  return Array.from(el.childNodes).filter((n) => n.nodeType === 1);
}

// Tekst elementu, także usuniętego (w:delText) — dla opisu na liście.
function revText(el) {
  let out = "";
  const walk = (n) => {
    if (n.nodeType !== 1) return;
    if (n.namespaceURI === W_NS && (n.localName === "t" || n.localName === "delText")) out += n.textContent || "";
    else if (n.namespaceURI === W_NS && n.localName === "tab") out += " ";
    else for (let i = 0; i < n.childNodes.length; i++) walk(n.childNodes[i]);
  };
  walk(el);
  return out;
}

function revParaText(p) {
  let out = "";
  revKids(p, "t").forEach((t) => { out += t.textContent || ""; });
  return out;
}

function revClosest(el, localName) {
  let n = el;
  while (n && n.nodeType === 1) {
    if (n.namespaceURI === W_NS && n.localName === localName) return n;
    n = n.parentNode;
  }
  return null;
}

function revIsParaMark(el) {
  // <w:pPr><w:rPr><w:ins|w:del/></w:rPr></w:pPr> — wstawiony / usunięty znak akapitu
  const rPr = el.parentNode;
  return rPr?.localName === "rPr" && rPr.parentNode?.localName === "pPr";
}

function revIsRowMark(el) {
  return el.parentNode?.localName === "trPr";
}

function revPartLabel(name) {
  if (name === "word/document.xml") return "body";
  if (/header/.test(name)) return "header";
  if (/footer/.test(name)) return "footer";
  if (/footnotes/.test(name)) return "footnote";
  if (/endnotes/.test(name)) return "endnote";
  return "other";
}

// ── odczyt ───────────────────────────────────────────────────────────────────
async function scanDocxRevisions(bytes) {
  const zip = await loadDocxZipCached(bytes);
  const result = { changes: [], comments: [], notes: [], authors: new Map() };
  const parts = Object.keys(zip.files).filter((n) => REV_PARTS_RE.test(n)).sort((a, b) => (a === "word/document.xml" ? -1 : b === "word/document.xml" ? 1 : a.localeCompare(b)));
  let docParaIndex = new Map();
  let docXml = null;

  for (const name of parts) {
    const xml = await zip.file(name).async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const part = revPartLabel(name);
    const paraIndex = new Map();
    if (part === "body") {
      docXml = doc;
      collectParagraphElements(doc.documentElement, "body").forEach((p, i) => paraIndex.set(p, i));
      docParaIndex = paraIndex;
    }
    const where = (el) => {
      const p = revClosest(el, "p");
      return { part, partName: name, paraIndex: p && paraIndex.has(p) ? paraIndex.get(p) : null, context: p ? revParaText(p) : "" };
    };
    const add = (kind, el, text) => {
      const author = revAttr(el, "author") || "?";
      result.authors.set(author, (result.authors.get(author) || 0) + 1);
      result.changes.push({ id: revAttr(el, "id"), kind, author, date: revAttr(el, "date"), text: (text || "").trim(), ...where(el) });
    };
    ["ins", "del", "moveFrom", "moveTo"].forEach((tag) => {
      revKids(doc.documentElement, tag).forEach((el) => {
        if (revIsParaMark(el)) add(tag === "ins" ? "paraIns" : "paraDel", el, "");
        else if (revIsRowMark(el)) add(tag === "ins" ? "rowIns" : "rowDel", el, revText(el.parentNode.parentNode));
        else add(tag, el, revText(el));
      });
    });
    REV_CHANGE_TAGS.forEach((tag) => {
      revKids(doc.documentElement, tag).forEach((el) => {
        const owner = revClosest(el, "r") || revClosest(el, "p") || el.parentNode;
        add(tag === "rPrChange" ? "format" : tag === "pPrChange" ? "paraFormat" : "otherFormat", el, revText(owner));
      });
    });

    if (part === "footnote" || part === "endnote") {
      revChildren(doc.documentElement).forEach((note) => {
        const type = revAttr(note, "type");
        if (type === "separator" || type === "continuationSeparator" || type === "continuationNotice") return;
        result.notes.push({ type: part, id: revAttr(note, "id"), text: revText(note).trim(), paraIndex: null });
      });
    }
  }
  // kolejność w dokumencie: najpierw treść wg akapitu, potem nagłówki/stopki/przypisy
  const order = { body: 0, header: 1, footer: 2, footnote: 3, endnote: 4, other: 5 };
  result.changes.sort((a, b) => order[a.part] - order[b.part] || (a.paraIndex ?? 1e9) - (b.paraIndex ?? 1e9));

  // Przypisy: gdzie są odwołania w treści
  if (docXml) {
    ["footnote", "endnote"].forEach((type) => {
      revKids(docXml.documentElement, `${type}Reference`).forEach((ref) => {
        const note = result.notes.find((n) => n.type === type && n.id === revAttr(ref, "id"));
        const p = revClosest(ref, "p");
        if (note && p && docParaIndex.has(p)) { note.paraIndex = docParaIndex.get(p); note.context = revParaText(p); }
      });
    });
    result.notes.sort((a, b) => (a.type === b.type ? Number(a.id) - Number(b.id) : a.type === "footnote" ? -1 : 1));
  }

  // Komentarze + zakomentowany tekst + odpowiedzi (commentsExtended: paraIdParent)
  const cFile = zip.file("word/comments.xml");
  if (cFile) {
    const cDoc = new DOMParser().parseFromString(await cFile.async("string"), "application/xml");
    const anchors = docXml ? revCommentAnchors(docXml, docParaIndex) : new Map();
    const byParaId = new Map();
    revKids(cDoc.documentElement, "comment").forEach((c) => {
      const id = revAttr(c, "id");
      const paras = revKids(c, "p");
      const lastParaId = paras.length ? (paras[paras.length - 1].getAttributeNS(W14_NS, "paraId") || paras[paras.length - 1].getAttribute("w14:paraId")) : "";
      const a = anchors.get(id) || {};
      const item = { id, author: revAttr(c, "author") || "?", initials: revAttr(c, "initials"), date: revAttr(c, "date"), text: paras.map(revParaText).join("\n").trim(), anchor: (a.text || "").trim(), paraIndex: a.paraIndex ?? null, replies: [], done: false, parentId: null };
      if (lastParaId) byParaId.set(lastParaId.toUpperCase(), item);
      result.comments.push(item);
    });
    const exFile = zip.file("word/commentsExtended.xml");
    if (exFile) {
      const exDoc = new DOMParser().parseFromString(await exFile.async("string"), "application/xml");
      Array.from(exDoc.getElementsByTagNameNS(W15_NS, "commentEx")).forEach((ex) => {
        const item = byParaId.get((ex.getAttributeNS(W15_NS, "paraId") || "").toUpperCase());
        if (!item) return;
        item.done = ex.getAttributeNS(W15_NS, "done") === "1";
        const parent = byParaId.get((ex.getAttributeNS(W15_NS, "paraIdParent") || "").toUpperCase());
        if (parent && parent !== item) { item.parentId = parent.id; parent.replies.push(item); }
      });
    }
    result.comments = result.comments.filter((c) => !c.parentId);
    result.comments.sort((a, b) => (a.paraIndex ?? 1e9) - (b.paraIndex ?? 1e9));
  }
  return result;
}

// Tekst między commentRangeStart a commentRangeEnd (w kolejności dokumentu) + akapit początku.
function revCommentAnchors(doc, paraIndex) {
  const out = new Map();
  const open = new Set();
  const walker = doc.createTreeWalker(doc.documentElement, 1 /* SHOW_ELEMENT */);
  let n = walker.currentNode;
  while (n) {
    if (n.namespaceURI === W_NS) {
      const id = n.localName.startsWith("comment") ? revAttr(n, "id") : "";
      if (n.localName === "commentRangeStart") {
        open.add(id);
        const p = revClosest(n, "p") || revClosest(n.nextSibling || n, "p");
        out.set(id, { text: "", paraIndex: p && paraIndex.has(p) ? paraIndex.get(p) : null });
      } else if (n.localName === "commentRangeEnd") open.delete(id);
      else if (n.localName === "commentReference" && !out.has(id)) {
        const p = revClosest(n, "p");
        out.set(id, { text: "", paraIndex: p && paraIndex.has(p) ? paraIndex.get(p) : null });
      } else if (n.localName === "t" && open.size) {
        open.forEach((cid) => { const a = out.get(cid); if (a && a.text.length < 400) a.text += n.textContent || ""; });
      }
    }
    n = walker.nextNode();
  }
  return out;
}

// Szybkie liczenie przy otwarciu pliku (bez pełnego skanu) — do plakietki i komunikatu.
async function countDocxRevisions(bytes) {
  const zip = await loadDocxZipCached(bytes);
  let changes = 0;
  for (const name of Object.keys(zip.files).filter((n) => REV_PARTS_RE.test(n))) {
    const xml = await zip.file(name).async("string");
    changes += (xml.match(/<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|tcPrChange|trPrChange|tblGridChange)\b/g) || []).length;
  }
  const cXml = await zip.file("word/comments.xml")?.async("string");
  const comments = cXml ? (cXml.match(/<w:comment\b/g) || []).length : 0;
  return { changes, comments };
}

// ── zapis ────────────────────────────────────────────────────────────────────
function revUnwrap(el) {
  const parent = el.parentNode;
  while (el.firstChild) parent.insertBefore(el.firstChild, el);
  parent.removeChild(el);
}

function revRename(el, localName) {
  const doc = el.ownerDocument;
  const next = doc.createElementNS(W_NS, `w:${localName}`);
  Array.from(el.attributes).forEach((a) => next.setAttributeNS(a.namespaceURI, a.name, a.value));
  while (el.firstChild) next.appendChild(el.firstChild);
  el.parentNode.replaceChild(next, el);
}

// Usunięty/wstawiony znak akapitu = akapit łączy się z następnym (tekst przechodzi na początek
// następnego, który zachowuje swoje formatowanie — tak robi Word).
function revMergeWithNext(p) {
  let next = p.nextSibling;
  while (next && next.nodeType !== 1) next = next.nextSibling;
  if (!next || next.namespaceURI !== W_NS || next.localName !== "p") return false;
  const nextPPr = revChildren(next).find((c) => c.localName === "pPr");
  const anchor = nextPPr ? nextPPr.nextSibling : next.firstChild;
  revChildren(p).filter((c) => c.localName !== "pPr").forEach((c) => next.insertBefore(c, anchor));
  p.parentNode.removeChild(p);
  return true;
}

// Odrzucenie zmiany właściwości: w środku *PrChange jest poprzednia wersja właściwości.
const REV_KEEP_ON_RESTORE = new Set(["rPr", "sectPr", "ins", "del", "moveFrom", "moveTo", "headerReference", "footerReference"]);
function revRestoreProps(change) {
  const holder = change.parentNode;
  const old = revChildren(change)[0];
  const kept = [];
  revChildren(holder).forEach((c) => {
    if (c === change) return;
    if (REV_KEEP_ON_RESTORE.has(c.localName)) kept.push(c);
    else holder.removeChild(c);
  });
  // kolejność ze schematu: w pPr stare właściwości idą PRZED rPr/sectPr; w rPr znaczniki
  // ins/del są pierwsze, więc właściwości za nimi
  const anchor = holder.localName === "rPr" ? change : kept[0] || change;
  if (old) revChildren(old).forEach((c) => holder.insertBefore(c, anchor));
  holder.removeChild(change);
}

// mode: "accept" | "reject"; ids: Set<string> albo null (= wszystkie). Zwraca liczbę zmian.
function applyRevisionsToDoc(doc, mode, ids) {
  const pick = (el) => !ids || ids.has(revAttr(el, "id"));
  let count = 0;
  const merges = [];
  const rowsToRemove = [];

  ["ins", "del", "moveFrom", "moveTo"].forEach((tag) => {
    revKids(doc.documentElement, tag).forEach((el) => {
      if (!el.parentNode || !pick(el)) return;
      count++;
      const keepContent = mode === "accept" ? (tag === "ins" || tag === "moveTo") : (tag === "del" || tag === "moveFrom");
      if (revIsParaMark(el)) {
        const p = revClosest(el, "p");
        el.parentNode.removeChild(el);
        if (!keepContent && p) merges.push(p);
        return;
      }
      if (revIsRowMark(el)) {
        const tr = el.parentNode.parentNode;
        el.parentNode.removeChild(el);
        if (!keepContent) rowsToRemove.push(tr);
        return;
      }
      if (!keepContent) { el.parentNode.removeChild(el); return; }
      revKids(el, "delText").forEach((d) => revRename(d, "t"));
      revKids(el, "delInstrText").forEach((d) => revRename(d, "instrText"));
      revUnwrap(el);
    });
  });
  // zakresy przeniesień — same znaczniki
  ["moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"].forEach((tag) => {
    revKids(doc.documentElement, tag).forEach((el) => { if (!ids) el.parentNode?.removeChild(el); });
  });
  REV_CHANGE_TAGS.forEach((tag) => {
    revKids(doc.documentElement, tag).forEach((el) => {
      if (!el.parentNode || !pick(el)) return;
      count++;
      if (mode === "accept") el.parentNode.removeChild(el);
      else revRestoreProps(el);
    });
  });
  rowsToRemove.forEach((tr) => tr.parentNode?.removeChild(tr));
  // od końca, żeby łączenie kilku akapitów pod rząd szło jak w Wordzie
  merges.reverse().forEach((p) => { if (p.parentNode) revMergeWithNext(p); });
  return count;
}

function removeCommentMarkersFromDoc(doc, ids) {
  const pick = (el) => !ids || ids.has(revAttr(el, "id"));
  let count = 0;
  ["commentRangeStart", "commentRangeEnd"].forEach((tag) => {
    revKids(doc.documentElement, tag).forEach((el) => { if (pick(el)) el.parentNode.removeChild(el); });
  });
  revKids(doc.documentElement, "commentReference").forEach((el) => {
    if (!pick(el)) return;
    count++;
    const run = el.parentNode;
    el.parentNode.removeChild(el);
    if (run?.localName === "r" && !revChildren(run).some((c) => c.localName !== "rPr")) run.parentNode.removeChild(run);
  });
  return count;
}

// Wołane z buildPatchedDocx: edit = { op: "revisions", action: "accept"|"reject"|"removeComments", ids?: string[] }
// Zwraca { xml: nowy document.xml, count } i sam zapisuje pozostałe części w zip.
async function applyRevisionsInZip(zip, documentXml, edit) {
  const ids = Array.isArray(edit.ids) && edit.ids.length ? new Set(edit.ids.map(String)) : null;
  let total = 0;
  let outXml = documentXml;
  const names = Object.keys(zip.files).filter((n) => REV_PARTS_RE.test(n));
  for (const name of names) {
    const xml = name === "word/document.xml" ? documentXml : await zip.file(name).async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const n = edit.action === "removeComments" ? removeCommentMarkersFromDoc(doc, ids) : applyRevisionsToDoc(doc, edit.action, ids);
    if (!n) continue;
    total += n;
    const out = new XMLSerializer().serializeToString(doc);
    if (name === "word/document.xml") outXml = out;
    else zip.file(name, out);
  }
  if (edit.action === "removeComments") {
    const cFile = zip.file("word/comments.xml");
    if (cFile) {
      const cDoc = new DOMParser().parseFromString(await cFile.async("string"), "application/xml");
      const removedParaIds = new Set();
      let removed = 0;
      revKids(cDoc.documentElement, "comment").forEach((c) => {
        if (ids && !ids.has(revAttr(c, "id"))) return;
        revKids(c, "p").forEach((p) => { const pid = p.getAttributeNS(W14_NS, "paraId"); if (pid) removedParaIds.add(pid.toUpperCase()); });
        c.parentNode.removeChild(c);
        removed++;
      });
      zip.file("word/comments.xml", new XMLSerializer().serializeToString(cDoc));
      total = Math.max(total, removed);
      // dodatkowe części komentarzy (Word 2013+) — bez osieroconych wpisów
      for (const extra of ["word/commentsExtended.xml", "word/commentsIds.xml", "word/commentsExtensible.xml"]) {
        const f = zip.file(extra);
        if (!f) continue;
        const d = new DOMParser().parseFromString(await f.async("string"), "application/xml");
        revChildren(d.documentElement).forEach((el) => {
          const pid = (Array.from(el.attributes).find((a) => a.localName === "paraId")?.value || "").toUpperCase();
          if (!ids || removedParaIds.has(pid)) d.documentElement.removeChild(el);
        });
        zip.file(extra, new XMLSerializer().serializeToString(d));
      }
    }
  }
  return { xml: outXml, count: total };
}

// ── plakietka przy panelu + jednorazowy komunikat po otwarciu pliku ─────────
let docReviewCounts = { changes: 0, comments: 0 };
let reviewNoticePending = false;
let reviewCountJob = 0;

async function refreshReviewCounts() {
  const job = ++reviewCountJob;
  if (!originalFileBytes || !window.JSZip) {
    docReviewCounts = { changes: 0, comments: 0 };
  } else {
    try { docReviewCounts = await countDocxRevisions(originalFileBytes); } catch (_) { docReviewCounts = { changes: 0, comments: 0 }; }
  }
  if (job !== reviewCountJob) return;
  if (typeof appFrame !== "undefined") appFrame.syncPanelCounts();
  // otwarty panel Recenzja nadąża za edycjami (Cofnij, zmiany z innych paneli)
  if (document.getElementById("panel-review")?.open && typeof runReviewScan === "function") runReviewScan().catch(() => {});
  if (reviewNoticePending) {
    reviewNoticePending = false;
    const { changes, comments } = docReviewCounts;
    if (changes || comments) toast(t("reviewOnOpen", { changes, comments }), "info");
  }
}

document.addEventListener("DOMContentLoaded", () => {
  const orig = window.ingestFile;
  if (typeof orig !== "function") return;
  window.ingestFile = async function ingestFileReview(...args) {
    reviewNoticePending = true;
    return orig.apply(this, args);
  };
});
