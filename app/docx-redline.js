// docx-redline.js — plik z poprawkami Worda (w:ins / w:del) z dwóch wersji: „przed” i „po”.
//
// Jeden silnik dla dwóch funkcji:
//   - Śledź zmiany (review-panel.js): przed = stan z chwili włączenia, po = dokument teraz;
//   - Porównaj wersje → „Utwórz dokument z poprawkami” (compare-panel.js): starsza i nowsza.
// Wynik to NOWSZA wersja pliku (style, nagłówki, obrazy, relacje — z niej), w której:
//   - akapity zmienione — różnica na poziomie słów: tekst usunięty w w:del (w:delText),
//     wstawiony w w:ins, format fragmentów zachowany (ze swojej wersji),
//   - akapity dodane — treść w w:ins + znacznik akapitu wstawiony (w:pPr/w:rPr/w:ins),
//   - akapity usunięte — wracają na swoje miejsce jako w:del (+ znacznik akapitu usunięty);
//     bez zakładek, komentarzy, przypisów, obrazów i linków ze starej paczki (ich id i relacje
//     należą do tamtego pliku — w tym nie istnieją).
// Akapit „złożony” (pole, kontrolka, wzór, już istniejąca poprawka, fragment zagnieżdżony) nie
// jest rozbierany na słowa: stara wersja — usunięta, nowa — wstawiona (bezpieczne, jak Word przy
// akapitach, których nie umie dopasować). Zmiany samego formatowania nie są śledzone.

const RL_COMPLEX_TAGS = ["fldChar", "fldSimple", "instrText", "sdt", "smartTag", "customXml", "oMath", "oMathPara", "txbxContent", "ins", "del", "moveFrom", "moveTo", "rPrChange"];
const RL_DEL_DROP_TAGS = ["bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "permStart", "permEnd", "proofErr"];
const RL_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function rlKids(el) { return Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.namespaceURI === W_NS); }
function rlEl(doc, name, attrs) {
  const el = doc.createElementNS(W_NS, `w:${name}`);
  Object.entries(attrs || {}).forEach(([k, v]) => el.setAttributeNS(W_NS, `w:${k}`, String(v)));
  return el;
}

function rlComplex(p) {
  if (RL_COMPLEX_TAGS.some((tag) => p.getElementsByTagNameNS(W_NS, tag).length)) return true;
  // fragment tekstu głębiej niż akapit/link (np. w polu, kontrolce) — model akapitu by go zgubił
  return Array.from(p.getElementsByTagNameNS(W_NS, "r")).some((r) => r.parentNode !== p && r.parentNode.localName !== "hyperlink");
}

function rlRuns(p) { return extractRunsFromParagraphXml(p); }
function rlText(runs) { return (runs || []).map((r) => (r.break ? "\n" : r.island ? "" : r.text || "")).join(""); }

// Słowa, odstępy, znaki interpunkcji; łamanie wiersza, tabulator i wyspa — osobne tokeny.
function rlTokens(runs, side) {
  const out = [];
  runs.forEach((run, ri) => {
    if (run.break) { out.push({ key: "\u0001br", ri, side, kind: "break" }); return; }
    if (run.island) { out.push({ key: `\u0001is:${run.island}`, ri, side, kind: "island" }); return; }
    if (run.tab) { out.push({ key: "\t", ri, side, kind: "tab", text: "\t" }); return; }
    const parts = String(run.text || "").match(/\s+|[\p{L}\p{N}_'’]+|[^\s\p{L}\p{N}_'’]/gu) || [];
    parts.forEach((text) => out.push({ key: text, ri, side, kind: "text", text }));
  });
  return out;
}

// LCS na tokenach → [{ kind: same|ins|del, tok }]. Zbyt duży akapit — całość: usunięte + wstawione.
function rlDiffTokens(a, b) {
  if ((a.length + 1) * (b.length + 1) > 4_000_000) return [...a.map((tok) => ({ kind: "del", tok })), ...b.map((tok) => ({ kind: "ins", tok }))];
  const w = b.length + 1;
  const grid = new Uint32Array((a.length + 1) * w);
  for (let i = a.length - 1; i >= 0; i--) {
    for (let j = b.length - 1; j >= 0; j--) {
      grid[i * w + j] = a[i].key === b[j].key ? grid[(i + 1) * w + j + 1] + 1 : Math.max(grid[(i + 1) * w + j], grid[i * w + j + 1]);
    }
  }
  const out = [];
  let i = 0, j = 0;
  while (i < a.length || j < b.length) {
    if (i < a.length && j < b.length && a[i].key === b[j].key) { out.push({ kind: "same", tok: b[j], old: a[i] }); i++; j++; }
    else if (j < b.length && (i >= a.length || grid[i * w + j + 1] >= grid[(i + 1) * w + j])) out.push({ kind: "ins", tok: b[j++] });
    else out.push({ kind: "del", tok: a[i++] });
  }
  // usunięcia przed wstawieniami w każdym ciągu zmian (jak Word: „stare” → „nowe”)
  const sorted = [];
  for (let k = 0; k < out.length;) {
    if (out[k].kind === "same") { sorted.push(out[k++]); continue; }
    const block = [];
    while (k < out.length && out[k].kind !== "same") block.push(out[k++]);
    sorted.push(...block.filter((x) => x.kind === "del"), ...block.filter((x) => x.kind === "ins"));
  }
  return sorted;
}

// w:t → w:delText, w:instrText → w:delInstrText (w poddrzewie)
function rlToDeleted(node) {
  const doc = node.ownerDocument;
  [["t", "delText"], ["instrText", "delInstrText"]].forEach(([from, to]) => {
    Array.from(node.getElementsByTagNameNS(W_NS, from)).forEach((t) => {
      const d = rlEl(doc, to);
      if (t.getAttribute("xml:space") || /^\s|\s$/.test(t.textContent)) d.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
      d.textContent = t.textContent;
      t.parentNode.replaceChild(d, t);
    });
  });
}

function rlRevision(state, kind) {
  return rlEl(state.doc, kind, { id: state.nextId++, author: state.author, date: state.date });
}

// Znacznik akapitu wstawiony / usunięty: w:pPr/w:rPr/w:ins|del (na początku w:rPr)
function rlMarkParagraph(state, p, kind) {
  let pPr = rlKids(p).find((n) => n.localName === "pPr");
  if (!pPr) { pPr = rlEl(state.doc, "pPr"); p.insertBefore(pPr, p.firstChild); }
  let rPr = rlKids(pPr).find((n) => n.localName === "rPr");
  if (!rPr) {
    rPr = rlEl(state.doc, "rPr");
    const after = rlKids(pPr).find((n) => n.localName === "sectPr" || n.localName === "pPrChange");
    pPr.insertBefore(rPr, after || null);
  }
  rPr.insertBefore(rlRevision(state, kind), rPr.firstChild);
}

// Owinięcie całej treści akapitu w w:ins / w:del (fragmenty w linku, polu prostym — w środku nich).
function rlWrapRuns(state, container, kind) {
  let group = null;
  const close = () => { group = null; };
  Array.from(container.childNodes).forEach((n) => {
    if (n.nodeType !== 1 || n.namespaceURI !== W_NS) return;
    if (n.localName === "r") {
      if (kind === "del") rlToDeleted(n);
      if (!group) { group = rlRevision(state, kind); container.insertBefore(group, n); }
      group.appendChild(n);
      return;
    }
    close();
    if (n.localName === "hyperlink" || n.localName === "fldSimple" || n.localName === "smartTag") rlWrapRuns(state, n, kind);
    else if (n.localName === "sdt") { const c = rlKids(n).find((x) => x.localName === "sdtContent"); if (c) rlWrapRuns(state, c, kind); }
  });
}

// Usunięty akapit ze starej paczki — bez rzeczy, które w nowym pliku nie mają odniesienia.
function rlDeletedClone(state, oldP) {
  const p = state.doc.importNode(oldP, true);
  const pPr = rlKids(p).find((n) => n.localName === "pPr");
  if (pPr) rlKids(pPr).filter((n) => n.localName === "sectPr" || n.localName === "pPrChange").forEach((n) => pPr.removeChild(n));
  RL_DEL_DROP_TAGS.forEach((tag) => Array.from(p.getElementsByTagNameNS(W_NS, tag)).forEach((n) => n.parentNode.removeChild(n)));
  // fragmenty z odwołaniem do części / relacji starej paczki
  Array.from(p.getElementsByTagNameNS(W_NS, "r")).forEach((r) => {
    if (["drawing", "pict", "object", "footnoteReference", "endnoteReference", "commentReference"].some((tag) => r.getElementsByTagNameNS(W_NS, tag).length)) r.parentNode.removeChild(r);
  });
  Array.from(p.getElementsByTagNameNS(W_NS, "hyperlink")).forEach((h) => {
    if (!h.getAttributeNS(RL_REL_NS, "id")) return;
    while (h.firstChild) h.parentNode.insertBefore(h.firstChild, h);
    h.parentNode.removeChild(h);
  });
  // istniejące poprawki w starej wersji: wstawienia przyjmujemy, usunięte już nie istniały
  Array.from(p.getElementsByTagNameNS(W_NS, "del")).filter((d) => d.parentNode?.localName !== "rPr").forEach((d) => d.parentNode.removeChild(d));
  Array.from(p.getElementsByTagNameNS(W_NS, "ins")).filter((d) => d.parentNode?.localName !== "rPr").forEach((d) => { while (d.firstChild) d.parentNode.insertBefore(d.firstChild, d); d.parentNode.removeChild(d); });
  rlWrapRuns(state, p, "del");
  rlMarkParagraph(state, p, "del");
  return p;
}

// Akapit zmieniony: różnica słów. Zwraca false, gdy nic się nie zmieniło.
function rlRewriteChanged(state, newP, oldP) {
  const newRuns = rlRuns(newP), oldRuns = rlRuns(oldP);
  if (rlText(newRuns) === rlText(oldRuns) && newRuns.filter((r) => r.island).length === oldRuns.filter((r) => r.island).length) return false;
  const diff = rlDiffTokens(rlTokens(oldRuns, "old"), rlTokens(newRuns, "new"));
  const doc = state.doc;
  const originalsNew = originalRunProps(newP);
  const originalsOld = originalRunProps(oldP);
  const linkAttrs = originalHyperlinkAttrs(newP);
  // pozycje: [{ kind, link, node }] — kolejne tokeny tego samego fragmentu sklejane w jeden w:r
  const items = [];
  let cur = null;
  const flush = () => {
    if (!cur) return;
    const run = { ...cur.run, text: cur.text };
    delete run.island; delete run.break; delete run.tab;
    if (cur.kind === "del") delete run.link; // relacje linku starej paczki tu nie istnieją
    const node = (cur.kind === "del" ? null : runFromOriginal(doc, run, originalsNew)) || (cur.kind === "del" ? runFromOriginal(doc, run, originalsOld) : null) || createRunElement(doc, run);
    if (cur.kind === "del") rlToDeleted(node);
    items.push({ kind: cur.kind, link: run.link || "", node });
    cur = null;
  };
  diff.forEach(({ kind: k, tok }) => {
    const kind = k === "same" ? "same" : k;
    const runs = tok.side === "old" ? oldRuns : newRuns;
    const run = runs[tok.ri];
    if (tok.kind === "island") {
      flush();
      if (kind === "del") return; // wyspa starej wersji (komentarz, zakładka) — nie dublujemy
      const frag = new DOMParser().parseFromString(run.island, "application/xml").documentElement;
      const node = doc.importNode(frag, true);
      if (node.getAttributeNS("http://www.w3.org/2000/xmlns/", "w") === W_NS) node.removeAttributeNS("http://www.w3.org/2000/xmlns/", "w");
      items.push({ kind: kind === "ins" ? "ins" : "same", link: "", node, island: true });
      return;
    }
    if (tok.kind === "break") {
      flush();
      const r = rlEl(doc, "r");
      r.appendChild(rlEl(doc, "br"));
      items.push({ kind, link: kind === "del" ? "" : run.link || "", node: r });
      return;
    }
    if (cur && cur.kind === kind && cur.ri === tok.ri && cur.side === tok.side) { cur.text += tok.text; return; }
    flush();
    cur = { kind, ri: tok.ri, side: tok.side, run, text: tok.text };
  });
  flush();
  // przebudowa akapitu: w:pPr zostaje, reszta z listy (linki → w:hyperlink, zmiany → w:ins/w:del)
  rlKids(newP).filter((n) => n.localName !== "pPr").forEach((n) => newP.removeChild(n));
  Array.from(newP.childNodes).filter((n) => n.nodeType !== 1).forEach((n) => newP.removeChild(n));
  let host = newP, hostLink = "", wrap = null;
  items.forEach((it) => {
    // wyspy (zakładki, komentarze) poza w:ins — znaczniki zakresu nie mogą być w poprawce
    const link = it.island ? "" : it.link;
    if (link !== hostLink) {
      wrap = null;
      if (link) { host = createHyperlinkElement(doc, link, linkAttrs.get(link)); newP.appendChild(host); }
      else host = newP;
      hostLink = link;
    }
    if (it.kind === "same" || (it.island && it.node.localName !== "r")) { wrap = null; host.appendChild(it.node); return; }
    if (!wrap || wrap.localName !== it.kind) { wrap = rlRevision(state, it.kind); host.appendChild(wrap); }
    wrap.appendChild(it.node);
  });
  return true;
}

function rlParagraphKey(p) {
  return rlComplex(p) ? `\u0002${new XMLSerializer().serializeToString(p).replace(/\s(w:rsid\w*|w14:paraId|w14:textId)="[^"]*"/g, "")}` : rlText(rlRuns(p));
}

// Wyrównanie akapitów (LCS na dokładnym tekście; puste akapity też — zachowują pozycję).
function rlAlign(oldKeys, newKeys) {
  const n = oldKeys.length, m = newKeys.length;
  if ((n + 1) * (m + 1) > 12_000_000) { // bardzo długie dokumenty — po kolei, bez dopasowania
    const rows = [];
    for (let i = 0; i < Math.max(n, m); i++) rows.push(i < n && i < m ? { kind: oldKeys[i] === newKeys[i] ? "same" : "changed", o: i, n: i } : i < m ? { kind: "added", n: i } : { kind: "removed", o: i });
    return rows;
  }
  const w = m + 1;
  const grid = new Uint32Array((n + 1) * w);
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) grid[i * w + j] = oldKeys[i] === newKeys[j] ? grid[(i + 1) * w + j + 1] + 1 : Math.max(grid[(i + 1) * w + j], grid[i * w + j + 1]);
  const raw = [];
  let i = 0, j = 0;
  while (i < n || j < m) {
    if (i < n && j < m && oldKeys[i] === newKeys[j]) { raw.push({ kind: "same", o: i++, n: j++ }); }
    else if (j < m && (i >= n || grid[i * w + j + 1] >= grid[(i + 1) * w + j])) raw.push({ kind: "added", n: j++ });
    else raw.push({ kind: "removed", o: i++ });
  }
  // usunięty + dodany obok siebie (po równo) → „zmieniony” — różnica słów zamiast całych akapitów
  const rows = [];
  for (let k = 0; k < raw.length;) {
    if (raw[k].kind === "same") { rows.push(raw[k++]); continue; }
    const rem = [], add = [];
    while (k < raw.length && raw[k].kind !== "same") (raw[k].kind === "removed" ? rem : add).push(raw[k++]);
    const pairs = Math.min(rem.length, add.length);
    for (let x = 0; x < pairs; x++) rows.push({ kind: "changed", o: rem[x].o, n: add[x].n });
    rem.slice(pairs).forEach((r) => rows.push(r));
    add.slice(pairs).forEach((r) => rows.push(r));
  }
  return rows;
}

function rlMaxId(doc) {
  let max = 0;
  const all = doc.getElementsByTagName("*");
  for (let i = 0; i < all.length; i++) {
    const v = all[i].getAttributeNS(W_NS, "id");
    const n = v ? parseInt(v, 10) : NaN;
    if (Number.isFinite(n) && n > max) max = n;
  }
  return max;
}

// Główna funkcja: bajty starszej i nowszej wersji → bajty nowszej z poprawkami.
// opts: { author, date (ISO), trackOn (w:trackRevisions w settings.xml: true/false/undefined) }
async function dwbRedline(oldBytes, newBytes, opts = {}) {
  const [oldZip, newZip] = await Promise.all([JSZip.loadAsync(oldBytes), JSZip.loadAsync(newBytes)]);
  const parse = (x) => new DOMParser().parseFromString(x, "application/xml");
  const oldDoc = parse(await oldZip.file("word/document.xml").async("string"));
  const doc = parse(await newZip.file("word/document.xml").async("string"));
  const oldParas = collectParagraphElements(oldDoc.documentElement, "all");
  const newParas = collectParagraphElements(doc.documentElement, "all");
  const state = {
    doc,
    author: String(opts.author || "Autor").slice(0, 120),
    date: (opts.date || new Date().toISOString()).replace(/\.\d{3}Z$/, "Z"),
    nextId: Math.max(rlMaxId(doc), rlMaxId(oldDoc)) + 1,
  };
  const rows = rlAlign(oldParas.map(rlParagraphKey), newParas.map(rlParagraphKey));
  const stats = { changed: 0, added: 0, removed: 0 };
  // pozycje usuniętych akapitów liczymy na NIEZMIENIONYCH odniesieniach (przed wstawianiem)
  const placements = [];
  rows.forEach((row, k) => {
    if (row.kind !== "removed") return;
    const oldP = oldParas[row.o];
    const sameBox = (r) => r && r.n != null && r.o != null && oldParas[r.o].parentNode === oldP.parentNode;
    let anchor = null, where = "before";
    for (let x = k + 1; x < rows.length && !anchor; x++) if (sameBox(rows[x])) anchor = newParas[rows[x].n];
    if (!anchor) { where = "after"; for (let x = k - 1; x >= 0 && !anchor; x--) if (sameBox(rows[x])) anchor = newParas[rows[x].n]; }
    if (!anchor && oldP.parentNode.localName === "body") {
      // w treści głównej bez sąsiada z tego samego miejsca — przed następnym akapitem treści
      for (let x = k + 1; x < rows.length && !anchor; x++) if (rows[x].n != null && newParas[rows[x].n].parentNode.localName === "body") anchor = newParas[rows[x].n];
      where = "before";
    }
    if (anchor) placements.push({ oldP, anchor, where });
  });
  rows.forEach((row) => {
    if (row.kind === "changed" || (row.kind === "same" && row.o != null && row.n != null)) {
      const np = newParas[row.n], op = oldParas[row.o];
      if (row.kind === "same") return;
      if (rlComplex(np) || rlComplex(op)) {
        placements.push({ oldP: op, anchor: np, where: "before" });
        rlWrapRuns(state, np, "ins");
        rlMarkParagraph(state, np, "ins");
        stats.changed++;
        return;
      }
      if (rlRewriteChanged(state, np, op)) stats.changed++;
    } else if (row.kind === "added") {
      const np = newParas[row.n];
      rlWrapRuns(state, np, "ins");
      rlMarkParagraph(state, np, "ins");
      stats.added++;
    }
  });
  const lastAfter = new Map();
  placements.forEach(({ oldP, anchor, where }) => {
    const clone = rlDeletedClone(state, oldP);
    if (where === "before") anchor.parentNode.insertBefore(clone, anchor);
    else { const after = lastAfter.get(anchor) || anchor; after.parentNode.insertBefore(clone, after.nextSibling); lastAfter.set(anchor, clone); }
    stats.removed++;
  });
  newZip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
  if (opts.trackOn !== undefined) await rlSetTrackRevisions(newZip, !!opts.trackOn);
  const bytes = await newZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return { bytes, stats };
}

// Brakująca część ustawień: plik, typ zawartości i relacja z dokumentu.
async function rlCreateSettingsPart(zip) {
  zip.file("word/settings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${W_NS}"/>`);
  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    let ct = await ctFile.async("string");
    if (!/PartName="\/word\/settings\.xml"/.test(ct)) ct = ct.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>');
    zip.file("[Content_Types].xml", ct);
  }
  const relsPath = "word/_rels/document.xml.rels";
  const relsFile = zip.file(relsPath);
  if (relsFile) {
    const rels = new DOMParser().parseFromString(await relsFile.async("string"), "application/xml");
    const root = rels.documentElement;
    const all = Array.from(root.getElementsByTagName("Relationship"));
    if (!all.some((r) => /\/settings$/.test(r.getAttribute("Type") || ""))) {
      let n = all.length + 1;
      while (all.some((r) => r.getAttribute("Id") === `rId${n}`)) n++;
      const rel = rels.createElementNS(root.namespaceURI, "Relationship");
      rel.setAttribute("Id", `rId${n}`);
      rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings");
      rel.setAttribute("Target", "settings.xml");
      root.appendChild(rel);
      zip.file(relsPath, new XMLSerializer().serializeToString(rels));
    }
  }
}

// Ustawienia → „Śledź zmiany” włączone w Wordzie po otwarciu pliku (w:trackRevisions).
const RL_SETTINGS_AFTER_TRACK = ["doNotTrackMoves", "doNotTrackFormatting", "documentProtection", "autoFormatOverride", "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle", "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets", "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema", "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"];
async function rlSetTrackRevisions(zip, on) {
  let f = zip.file("word/settings.xml");
  if (!f && !on) return;
  if (!f) await rlCreateSettingsPart(zip); // plik bez ustawień (np. z innego programu) — jak Word: dopisz część
  f = zip.file("word/settings.xml");
  if (!f) return;
  const doc = new DOMParser().parseFromString(await f.async("string"), "application/xml");
  const root = doc.documentElement;
  const old = rlKids(root).find((n) => n.localName === "trackRevisions");
  if (old) root.removeChild(old);
  if (on) {
    const after = rlKids(root).find((n) => RL_SETTINGS_AFTER_TRACK.includes(n.localName));
    root.insertBefore(rlEl(doc, "trackRevisions"), after || null);
  }
  zip.file("word/settings.xml", new XMLSerializer().serializeToString(doc));
}
