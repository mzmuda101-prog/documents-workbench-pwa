// docx-compose.js — tworzenie dokumentu od zera (Etap 1, 2026-10-01): operacje na pliku.
//   - styl akapitu (Normalny, Tytuł, Podtytuł, Nagłówek 1–3, Cytat) — op "paraFormat",
//   - wyrównanie (do lewej / środek / do prawej / wyjustuj) — ta sama op,
//   - podział strony (Ctrl+Enter) — op "pageBreak", linia pozioma — op "hrule",
//   - nowy dokument z szablonu (pusty, pismo, notatka) — createComposeDocx.
// Interfejs (menu „＋ Wstaw”, lista stylów, wyrównanie, okno „Nowy dokument”) jest w compose-ui.js.
//
// Style szukamy w styles.xml po NAZWIE wbudowanej („heading 1”, „Title”…), nie po id —
// polski Word zapisuje Nagłówek 1 jako styleId „Nagwek1”, a nazwa zostaje „heading 1”.
// Brakujący styl dopisujemy (plik z generatora często ma tylko Normal albo nie ma styles.xml).
//
// Kolejność dzieci w:pPr jest narzucona przez schemat — Word przy złej kolejności potrafi
// zgłosić „nieczytelną zawartość”. Dlatego każdy element wstawiamy na swoje miejsce (PPR_ORDER).

const COMPOSE_PPR_ORDER = [
  "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
  "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
  "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
  "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
  "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
  "sectPr", "pPrChange",
];

const COMPOSE_ALIGNS = ["left", "center", "right", "both"];

// key → nazwa wbudowana (małe litery) i definicja dopisywana, gdy w pliku jej nie ma.
// {NORMAL} = id stylu domyślnego w tym pliku (basedOn/next).
const COMPOSE_STYLE_DEFS = {
  normal: { name: "normal", id: "Normal" },
  title: {
    name: "title", id: "Title",
    body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>`,
  },
  subtitle: {
    name: "subtitle", id: "Subtitle",
    body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="11"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:color w:val="595959"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>`,
  },
  h1: {
    name: "heading 1", id: "Heading1",
    body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>`,
  },
  h2: {
    name: "heading 2", id: "Heading2",
    body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="2F5496"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>`,
  },
  h3: {
    name: "heading 3", id: "Heading3",
    body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="1F3763"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>`,
  },
  // Ramka: pasek z lewej i jasne tło (jak „Spróbuj:” w przewodniku) — wyróżniony akapit; Enter na
  // końcu → dalej zwykły tekst (w:next), pusty akapit w ramce + Enter/Backspace → Normalny
  callout: {
    name: "ramka", id: "DWBRamka", label: "Ramka",
    body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="30"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="1F5FBF"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="EEF3FB"/><w:spacing w:before="120" w:after="160"/><w:ind w:left="240" w:right="120"/></w:pPr>`,
  },
  quote: {
    name: "quote", id: "Quote",
    body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr>`,
  },
};
// spis treści (nie ma ich na liście stylów — dopisywane przy wstawianiu spisu)
Object.assign(COMPOSE_STYLE_DEFS, {
  tocHeading: {
    name: "toc heading", id: "TOCHeading",
    body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="9"/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>`,
  },
  toc1: { name: "toc 1", id: "TOC1", body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/></w:pPr>` },
  toc2: { name: "toc 2", id: "TOC2", body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/><w:ind w:left="220"/></w:pPr>` },
  toc3: { name: "toc 3", id: "TOC3", body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/><w:ind w:left="440"/></w:pPr>` },
});
const COMPOSE_STYLE_KEYS = ["normal", "title", "subtitle", "h1", "h2", "h3", "quote", "callout"];
const COMPOSE_TOC_KEYS = ["toc1", "toc2", "toc3"];

const COMPOSE_DOC_DEFAULTS = `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="pl-PL" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`;
const COMPOSE_NORMAL_STYLE = `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`;

function composeEl(doc, localName, attrs) {
  const el = doc.createElementNS(W_NS, `w:${localName}`);
  Object.entries(attrs || {}).forEach(([k, v]) => el.setAttributeNS(W_NS, `w:${k}`, String(v)));
  return el;
}

function composeDirectChild(parent, localName) {
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && n.namespaceURI === W_NS && n.localName === localName) return n;
  }
  return null;
}

function composeEnsurePPr(p) {
  let pPr = composeDirectChild(p, "pPr");
  if (!pPr) {
    pPr = composeEl(p.ownerDocument, "pPr");
    p.insertBefore(pPr, p.firstChild);
  }
  return pPr;
}

// Ustaw (attrs = obiekt / gotowy element) albo usuń (null) dziecko w:pPr — na właściwym miejscu.
function composeSetPPrChild(pPr, localName, attrs) {
  const old = composeDirectChild(pPr, localName);
  if (old) pPr.removeChild(old);
  if (attrs == null) return;
  const el = attrs.nodeType === 1 ? attrs : composeEl(pPr.ownerDocument, localName, attrs);
  const rank = COMPOSE_PPR_ORDER.indexOf(localName);
  let before = null;
  for (let n = pPr.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const r = COMPOSE_PPR_ORDER.indexOf(n.localName);
    if (r > rank) { before = n; break; }
  }
  pPr.insertBefore(el, before);
}

function composeParse(xml) {
  return new DOMParser().parseFromString(xml, "application/xml");
}

function composeSerialize(doc) {
  return new XMLSerializer().serializeToString(doc);
}

// ── style w styles.xml ───────────────────────────────────────────────────────
function composeStylesIndex(stylesDoc) {
  const out = { byName: new Map(), ids: new Set(), defaultId: null };
  if (!stylesDoc) return out;
  Array.from(stylesDoc.getElementsByTagNameNS(W_NS, "style")).forEach((st) => {
    const id = st.getAttributeNS(W_NS, "styleId");
    if (id) out.ids.add(id);
    if (st.getAttributeNS(W_NS, "type") !== "paragraph" || !id) return;
    const name = (composeDirectChild(st, "name")?.getAttributeNS(W_NS, "val") || "").trim().toLowerCase();
    if (name && !out.byName.has(name)) out.byName.set(name, id);
    const def = st.getAttributeNS(W_NS, "default");
    if (!out.defaultId && (def === "1" || def === "true")) out.defaultId = id;
  });
  if (!out.defaultId) out.defaultId = out.byName.get("normal") || null;
  return out;
}

// Mapa klasa CSS akapitu (docx-preview) → klucz stylu — stan listy „Styl” na pasku.
async function readComposeStyleClasses(bytes) {
  const map = new Map();
  if (!bytes || !window.JSZip) return map;
  try {
    const zip = await loadDocxZipCached(bytes);
    const xml = await zip.file("word/styles.xml")?.async("string");
    if (!xml) return map;
    const idx = composeStylesIndex(composeParse(xml));
    // styl znakowy numeru strony (pole PAGE w stopce) — podgląd podstawia w nim numer
    const pn = Array.from(composeParse(xml).getElementsByTagNameNS(W_NS, "style")).find((st) => st.getAttributeNS(W_NS, "type") === "character"
      && (composeDirectChild(st, "name")?.getAttributeNS(W_NS, "val") || "").trim().toLowerCase() === "page number");
    if (pn) map.set(docxStyleClassName(pn.getAttributeNS(W_NS, "styleId")), "pagenum");
    [...COMPOSE_STYLE_KEYS, ...COMPOSE_TOC_KEYS].forEach((key) => {
      const id = key === "normal" ? idx.defaultId : idx.byName.get(COMPOSE_STYLE_DEFS[key].name);
      if (id && !map.has(docxStyleClassName(id))) map.set(docxStyleClassName(id), key);
    });
    // Pozostałe style akapitu z pliku („Wskazówka”, „Tekst podstawowy”…) — lista stylów pokazuje
    // ich nazwę. Dawniej pokazywała „Normalny”, więc wybranie „Normalny” nic nie robiło i z akapitu
    // w takim stylu (np. ramka z tłem) nie było jak wyjść. Akapit listy ma swój przycisk — pomijamy.
    Array.from(composeParse(xml).getElementsByTagNameNS(W_NS, "style")).forEach((st) => {
      const id = st.getAttributeNS(W_NS, "styleId");
      if (st.getAttributeNS(W_NS, "type") !== "paragraph" || !id) return;
      const cls = docxStyleClassName(id);
      const name = (composeDirectChild(st, "name")?.getAttributeNS(W_NS, "val") || id).trim();
      if (map.has(cls) || /^(list paragraph|akapit z listą|toc|spis treści)/i.test(name)) return;
      if (typeof docHeadingStyleClasses !== "undefined" && docHeadingStyleClasses.get?.(cls)) return;
      map.set(cls, `custom:${name}`);
    });
  } catch (_) { /* uszkodzony styles.xml — lista pokaże „Normalny” */ }
  return map;
}

async function composeEnsureStylesPart(zip) {
  const file = zip.file("word/styles.xml");
  if (file) return file.async("string");
  // Brak styles.xml (pliki z generatorów): dopisujemy część + typ + powiązanie.
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="${W_NS}">${COMPOSE_DOC_DEFAULTS}${COMPOSE_NORMAL_STYLE}</w:styles>`;
  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    let ct = await ctFile.async("string");
    if (!/PartName="\/word\/styles\.xml"/.test(ct)) {
      ct = ct.replace("</Types>", `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`);
      zip.file("[Content_Types].xml", ct);
    }
  }
  const relsPath = "word/_rels/document.xml.rels";
  let rels = zip.file(relsPath) ? await zip.file(relsPath).async("string")
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  if (!/relationships\/styles"/.test(rels)) {
    let n = 1;
    while (new RegExp(`Id="rId${n}"`).test(rels)) n++;
    rels = rels.replace("</Relationships>", `<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
    zip.file(relsPath, rels);
  }
  return xml;
}

// Zwraca styleId dla klucza (dopisuje definicję, gdy trzeba) albo null = styl domyślny (bez w:pStyle).
async function composeEnsureStyle(zip, key, cache) {
  if (!COMPOSE_STYLE_DEFS[key]) return null;
  if (cache.has(key)) return cache.get(key);
  if (!cache.stylesXml) cache.stylesXml = await composeEnsureStylesPart(zip);
  const doc = composeParse(cache.stylesXml);
  const idx = composeStylesIndex(doc);
  if (key === "normal") { cache.set(key, null); return null; }
  const def = COMPOSE_STYLE_DEFS[key];
  let id = idx.byName.get(def.name);
  if (!id) {
    id = def.id;
    while (idx.ids.has(id)) id += "1";
    const normal = idx.defaultId;
    let body = def.body;
    body = normal ? body.replaceAll("{NORMAL}", normal) : body.replace(/<w:basedOn[^>]*\/>|<w:next[^>]*\/>/g, "");
    const frag = composeParse(`<w:styles xmlns:w="${W_NS}"><w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${def.label || def.name}"/>${body}</w:style></w:styles>`);
    doc.documentElement.appendChild(doc.importNode(frag.documentElement.firstChild, true));
    cache.stylesXml = composeSerialize(doc);
    cache.stylesDirty = true;
  }
  cache.set(key, id);
  return id;
}

// ── op "paraFormat": styl i/lub wyrównanie akapitów ──────────────────────────
// ── odstępy i wcięcia akapitu (Word: Akapit → Wcięcia i odstępy) ──────────────
// Wartości jak w pliku: odstępy i wcięcia w twipach (1 pt = 20, 1 cm ≈ 567), interlinia
// „auto” w 240-tych częściach wiersza (1,15 = 276), „atLeast”/„exact” — w twipach.
// Efektywne wartości warstwami jak w Wordzie: ustawienia domyślne dokumentu → łańcuch stylów
// (basedOn) → poziom listy (numbering.xml) → sam akapit. Każda warstwa nadpisuje to, co podaje.
const COMPOSE_LAYOUT_ZERO = { before: 0, after: 0, line: 240, lineRule: "auto", left: 0, right: 0, firstLine: 0, hanging: 0 };
function composeWAttr(el, name) {
  if (!el) return null;
  const v = el.getAttributeNS(W_NS, name) ?? el.getAttribute(`w:${name}`);
  return v == null || v === "" ? null : v;
}
function composeWInt(el, name) {
  const v = composeWAttr(el, name);
  const n = v == null ? NaN : parseInt(v, 10);
  return Number.isFinite(n) ? n : null;
}
function composeLayoutLayer(pPr, out) {
  if (!pPr) return;
  const sp = composeDirectChild(pPr, "spacing");
  if (sp) {
    ["before", "after"].forEach((k) => { const v = composeWInt(sp, k); if (v != null) out[k] = v; });
    const line = composeWInt(sp, "line");
    if (line != null) { out.line = line; out.lineRule = composeWAttr(sp, "lineRule") || "auto"; }
  }
  const ind = composeDirectChild(pPr, "ind");
  if (ind) {
    const left = composeWInt(ind, "left") ?? composeWInt(ind, "start");
    const right = composeWInt(ind, "right") ?? composeWInt(ind, "end");
    if (left != null) out.left = left;
    if (right != null) out.right = right;
    const hanging = composeWInt(ind, "hanging");
    const first = composeWInt(ind, "firstLine");
    if (hanging != null) { out.hanging = hanging; out.firstLine = 0; } // oba podane — wygrywa wysunięcie
    else if (first != null) { out.firstLine = first; out.hanging = 0; }
  }
}
function composeStyleById(stylesDoc, id) {
  if (!stylesDoc || !id) return null;
  return Array.from(stylesDoc.getElementsByTagNameNS(W_NS, "style")).find((st) => composeWAttr(st, "styleId") === id) || null;
}
function composeParaLayout(p, stylesDoc, numberingDoc) {
  const out = { ...COMPOSE_LAYOUT_ZERO };
  const def = stylesDoc?.getElementsByTagNameNS(W_NS, "pPrDefault")[0];
  composeLayoutLayer(def && composeDirectChild(def, "pPr"), out);
  const pPr = composeDirectChild(p, "pPr");
  const pStyle = pPr && composeDirectChild(pPr, "pStyle");
  let id = composeWAttr(pStyle, "val");
  if (!id && stylesDoc) {
    const d = Array.from(stylesDoc.getElementsByTagNameNS(W_NS, "style")).find((st) => composeWAttr(st, "type") === "paragraph" && /^(1|true|on)$/.test(composeWAttr(st, "default") || ""));
    id = d ? composeWAttr(d, "styleId") : null;
  }
  const chain = [];
  for (let i = 0, st = composeStyleById(stylesDoc, id); st && i < 12; i++, st = composeStyleById(stylesDoc, composeWAttr(composeDirectChild(st, "basedOn"), "val"))) chain.unshift(st);
  let numPr = null;
  chain.forEach((st) => {
    const sp = composeDirectChild(st, "pPr");
    composeLayoutLayer(sp, out);
    if (sp && composeDirectChild(sp, "numPr")) numPr = composeDirectChild(sp, "numPr");
  });
  if (pPr && composeDirectChild(pPr, "numPr")) numPr = composeDirectChild(pPr, "numPr");
  const numId = composeWAttr(numPr && composeDirectChild(numPr, "numId"), "val");
  if (numberingDoc && numId && numId !== "0") {
    const ilvl = composeWAttr(composeDirectChild(numPr, "ilvl"), "val") || "0";
    const num = Array.from(numberingDoc.getElementsByTagNameNS(W_NS, "num")).find((n) => composeWAttr(n, "numId") === numId);
    const absId = composeWAttr(num && composeDirectChild(num, "abstractNumId"), "val");
    const abs = Array.from(numberingDoc.getElementsByTagNameNS(W_NS, "abstractNum")).find((a) => composeWAttr(a, "abstractNumId") === absId);
    const lvl = abs && Array.from(abs.getElementsByTagNameNS(W_NS, "lvl")).find((l) => composeWAttr(l, "ilvl") === ilvl);
    composeLayoutLayer(lvl && composeDirectChild(lvl, "pPr"), out);
  }
  composeLayoutLayer(pPr, out);
  return out;
}
async function composeLayoutDocs(zip) {
  const read = async (name) => { const f = zip.file(name); return f ? composeParse(await f.async("string")) : null; };
  return { stylesDoc: await read("word/styles.xml"), numberingDoc: await read("word/numbering.xml") };
}
// Efektywne odstępy/wcięcia akapitów (indeksy jak w podglądzie) — do okienka „Odstępy i wcięcia”.
async function composeParaLayoutsFromBytes(bytes, indices) {
  const zip = await JSZip.loadAsync(bytes);
  const doc = composeParse(await zip.file("word/document.xml").async("string"));
  const { stylesDoc, numberingDoc } = await composeLayoutDocs(zip);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  return indices.map((i) => (paragraphs[i] ? composeParaLayout(paragraphs[i], stylesDoc, numberingDoc) : null));
}
// Zapis: tylko podane cechy, reszta w:spacing / w:ind zostaje. Wcięcie pierwszego wiersza
// i wysunięcie wykluczają się (jak w Wordzie). edit.indentDelta — Zwiększ/Zmniejsz wcięcie:
// każdy akapit o krok od SWOJEGO wcięcia (efektywnego, z warstwami).
function composeApplyLayout(p, pPr, edit, layout) {
  const doc = p.ownerDocument;
  if (edit.spacing) {
    const sp = composeDirectChild(pPr, "spacing") || composeEl(doc, "spacing");
    Object.entries(edit.spacing).forEach(([k, v]) => {
      if (!["before", "after", "line", "lineRule"].includes(k)) return;
      if (v == null) { sp.removeAttributeNS(W_NS, k); sp.removeAttribute(`w:${k}`); return; }
      sp.setAttributeNS(W_NS, `w:${k}`, String(k === "lineRule" ? v : Math.max(0, Math.round(v))));
      if (k === "before" || k === "after") { // „auto” odstęp HTML wygrywałby z wartością
        sp.removeAttributeNS(W_NS, `${k}Autospacing`);
        sp.removeAttribute(`w:${k}Autospacing`);
      }
    });
    if (edit.spacing.line != null && edit.spacing.lineRule == null) sp.setAttributeNS(W_NS, "w:lineRule", "auto");
    composeSetPPrChild(pPr, "spacing", sp.attributes.length ? sp : null);
  }
  const indEdit = { ...(edit.ind || {}) };
  if (Number.isFinite(edit.indentDelta)) indEdit.left = Math.max(0, (layout?.left || 0) + edit.indentDelta);
  if (Object.keys(indEdit).length) {
    const ind = composeDirectChild(pPr, "ind") || composeEl(doc, "ind");
    const drop = (k) => { ind.removeAttributeNS(W_NS, k); ind.removeAttribute(`w:${k}`); };
    const set = (k, v) => ind.setAttributeNS(W_NS, `w:${k}`, String(Math.round(v)));
    if (indEdit.left != null) { drop("start"); set("left", indEdit.left); }
    if (indEdit.right != null) { drop("end"); set("right", Math.max(0, indEdit.right)); }
    if (indEdit.hanging != null && indEdit.hanging > 0) { drop("firstLine"); set("hanging", indEdit.hanging); }
    else if (indEdit.firstLine != null) { drop("hanging"); set("firstLine", Math.max(0, indEdit.firstLine)); }
    else if (indEdit.hanging === 0) { drop("hanging"); set("firstLine", 0); }
    composeSetPPrChild(pPr, "ind", ind);
  }
}

async function applyParaFormatInZip(zip, xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const targets = (edit.indices || []).map((i) => paragraphs[i]).filter(Boolean);
  if (!targets.length) return { xml, count: 0 };
  const cache = new Map();
  let styleId;
  if (edit.style) styleId = await composeEnsureStyle(zip, edit.style, cache);
  const layoutDocs = Number.isFinite(edit.indentDelta) ? await composeLayoutDocs(zip) : null;
  let count = 0;
  targets.forEach((p) => {
    const layout = layoutDocs ? composeParaLayout(p, layoutDocs.stylesDoc, layoutDocs.numberingDoc) : null;
    const pPr = composeEnsurePPr(p);
    if (edit.spacing || edit.ind || Number.isFinite(edit.indentDelta)) composeApplyLayout(p, pPr, edit, layout);
    if (edit.style) {
      composeSetPPrChild(pPr, "pStyle", styleId ? { val: styleId } : null);
      // Nagłówek z listy stylów jest poziomem konspektu ze stylu — bezpośredni w:outlineLvl
      // akapitu (z importu) by go przesłaniał w spisie treści i nawigacji Worda.
      composeSetPPrChild(pPr, "outlineLvl", null);
    }
    if (edit.align && COMPOSE_ALIGNS.includes(edit.align)) composeSetPPrChild(pPr, "jc", { val: edit.align });
    if (!pPr.firstChild) p.removeChild(pPr);
    count++;
  });
  if (cache.stylesDirty) zip.file("word/styles.xml", cache.stylesXml);
  return { xml: composeSerialize(doc), count };
}

// ── op "pageBreak" / "hrule": nowe akapity ───────────────────────────────────
// Nowy pusty akapit „po” danym: styl domyślny, bez listy, bez znaczników sekcji.
function composeNewParagraphAfter(p) {
  const doc = p.ownerDocument;
  const np = composeEl(doc, "p");
  p.parentNode.insertBefore(np, p.nextSibling);
  return np;
}

// Znacznik końca sekcji (w:sectPr w akapicie) należy do OSTATNIEGO akapitu sekcji —
// przy dzieleniu przechodzi na nowy, drugi akapit (inaczej sekcja kończyłaby się za wcześnie).
function composeMoveSectPr(fromP, toP) {
  const pPr = composeDirectChild(fromP, "pPr");
  const sect = pPr && composeDirectChild(pPr, "sectPr");
  if (!sect) return;
  pPr.removeChild(sect);
  composeSetPPrChild(composeEnsurePPr(toP), "sectPr", sect);
}

// mode: "before" (kursor na początku akapitu) — akapit zaczyna się od nowej strony;
//       "after" (na końcu albo akapit tylko do odczytu) — nowy pusty akapit na nowej stronie;
//       "split" — tekst za kursorem przechodzi na nową stronę (jak Ctrl+Enter w Wordzie).
function applyPageBreakInXml(xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[edit.index];
  if (!p) return { xml, count: 0 };
  if (edit.mode === "before") {
    // Pierwszy akapit dokumentu: Word pomija „od nowej strony” na samym początku, więc
    // przed nim stawiamy pusty akapit (strona 1 zostaje pusta — jak Ctrl+Enter w Wordzie).
    if (paragraphs[0] === p) p.parentNode.insertBefore(composeEl(doc, "p"), p);
    composeSetPPrChild(composeEnsurePPr(p), "pageBreakBefore", {});
    return { xml: composeSerialize(doc), count: 1 };
  }
  let np;
  if (edit.mode === "split") {
    np = p.cloneNode(true);
    applyRunsToParagraphXml(p, edit.beforeRuns || []);
    applyRunsToParagraphXml(np, edit.afterRuns || []);
    p.parentNode.insertBefore(np, p.nextSibling);
    const pPr = composeDirectChild(p, "pPr");
    const sect = pPr && composeDirectChild(pPr, "sectPr");
    if (sect) pPr.removeChild(sect); // klon już go ma — zostaje na drugim akapicie
  } else {
    np = composeNewParagraphAfter(p);
    composeMoveSectPr(p, np);
  }
  composeSetPPrChild(composeEnsurePPr(np), "pageBreakBefore", {});
  return { xml: composeSerialize(doc), count: 1 };
}

// Linia pozioma = pusty akapit z dolną krawędzią (tak robi to Word: „---” + Enter),
// a za nim pusty akapit do dalszego pisania.
function applyHruleInXml(xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[edit.index];
  if (!p) return { xml, count: 0 };
  const line = composeNewParagraphAfter(p);
  const pBdr = composeEl(doc, "pBdr");
  pBdr.appendChild(composeEl(doc, "bottom", { val: "single", sz: 8, space: 1, color: "auto" }));
  const pPr = composeEnsurePPr(line);
  composeSetPPrChild(pPr, "pBdr", pBdr);
  composeSetPPrChild(pPr, "spacing", { before: 0, after: 240 });
  const next = composeNewParagraphAfter(line);
  composeMoveSectPr(p, next);
  return { xml: composeSerialize(doc), count: 1 };
}

// ── nowy dokument ────────────────────────────────────────────────────────────
function composeXmlText(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Akapit z prostego opisu: { text, style, align, bold, size, pageBreak }.
function composeParagraphXml(spec, styleIds) {
  const pPr = [];
  if (spec.style && styleIds[spec.style]) pPr.push(`<w:pStyle w:val="${styleIds[spec.style]}"/>`);
  if (spec.spacingAfter != null) pPr.push(`<w:spacing w:after="${spec.spacingAfter}"/>`);
  if (spec.align) pPr.push(`<w:jc w:val="${spec.align}"/>`);
  const rPr = [];
  if (spec.bold) rPr.push("<w:b/><w:bCs/>");
  if (spec.size) rPr.push(`<w:sz w:val="${spec.size * 2}"/><w:szCs w:val="${spec.size * 2}"/>`);
  const runs = String(spec.text || "").split("\n").map((line, i) => {
    const br = i ? "<w:br/>" : "";
    const t = line ? `<w:t xml:space="preserve">${composeXmlText(line)}</w:t>` : "";
    return `<w:r>${rPr.length ? `<w:rPr>${rPr.join("")}</w:rPr>` : ""}${br}${t}</w:r>`;
  }).join("");
  return `<w:p>${pPr.length ? `<w:pPr>${pPr.join("")}</w:pPr>` : ""}${spec.text ? runs : ""}</w:p>`;
}

// Treść szablonów w języku interfejsu. Pola {{…}} wypełnia panel „Pola {{…}}”.
function composeTemplateBody(kind, lang) {
  const pl = lang !== "en";
  const today = new Date().toLocaleDateString(pl ? "pl-PL" : "en-GB", { day: "numeric", month: "long", year: "numeric" });
  if (kind === "letter") {
    return pl ? [
      { text: "{{Imię i nazwisko}}\n{{Adres}}\n{{Kod pocztowy}} {{Miejscowość}}", spacingAfter: 240 },
      { text: `{{Miejscowość}}, ${today}`, align: "right", spacingAfter: 240 },
      { text: "{{Odbiorca}}\n{{Adres odbiorcy}}", align: "right", bold: true, spacingAfter: 360 },
      { text: "Dotyczy: {{Temat}}", bold: true, spacingAfter: 240 },
      { text: "Szanowni Państwo," },
      { text: "Treść pisma." },
      { text: "Z poważaniem,", spacingAfter: 600 },
      { text: "{{Imię i nazwisko}}" },
    ] : [
      { text: "{{Full name}}\n{{Address}}\n{{City}} {{Postcode}}", spacingAfter: 240 },
      { text: `{{City}}, ${today}`, align: "right", spacingAfter: 240 },
      { text: "{{Recipient}}\n{{Recipient address}}", align: "right", bold: true, spacingAfter: 360 },
      { text: "Re: {{Subject}}", bold: true, spacingAfter: 240 },
      { text: "Dear Sir or Madam," },
      { text: "Body of the letter." },
      { text: "Yours faithfully,", spacingAfter: 600 },
      { text: "{{Full name}}" },
    ];
  }
  if (kind === "note") {
    return pl ? [
      { text: "Tytuł dokumentu", style: "title" },
      { text: today, style: "subtitle" },
      { text: "Wprowadzenie", style: "h1" },
      { text: "Krótko: czego dotyczy dokument i dla kogo jest." },
      { text: "Najważniejsze informacje", style: "h1" },
      { text: "Szczegóły", style: "h2" },
      { text: "Treść sekcji." },
      { text: "Podsumowanie", style: "h1" },
      { text: "Wnioski i dalsze kroki." },
    ] : [
      { text: "Document title", style: "title" },
      { text: today, style: "subtitle" },
      { text: "Introduction", style: "h1" },
      { text: "In short: what this document is about and who it is for." },
      { text: "Key information", style: "h1" },
      { text: "Details", style: "h2" },
      { text: "Section text." },
      { text: "Summary", style: "h1" },
      { text: "Conclusions and next steps." },
    ];
  }
  return [{ text: "" }];
}

async function createComposeDocx(kind = "blank", lang = "pl") {
  if (!window.JSZip) throw new Error("JSZip missing");
  const styleIds = { normal: "Normal" };
  const styleXml = [COMPOSE_NORMAL_STYLE];
  COMPOSE_STYLE_KEYS.filter((k) => k !== "normal").forEach((k) => {
    const def = COMPOSE_STYLE_DEFS[k];
    styleIds[k] = def.id;
    styleXml.push(`<w:style w:type="paragraph" w:styleId="${def.id}"><w:name w:val="${def.label || def.name}"/>${def.body.replaceAll("{NORMAL}", "Normal")}</w:style>`);
  });
  const langTag = lang === "en" ? "en-GB" : "pl-PL";
  const body = composeTemplateBody(kind, lang).map((s) => composeParagraphXml(s, styleIds)).join("");
  // A4, marginesy 2,5 cm (1418 twipów) — jak w polskim Wordzie.
  const sectPr = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/><w:cols w:space="708"/></w:sectPr>`;
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  const zip = new window.JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}${sectPr}</w:body></w:document>`);
  zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">${COMPOSE_DOC_DEFAULTS.replace('w:val="pl-PL"', `w:val="${langTag}"`)}${styleXml.join("")}</w:styles>`);
  zip.file("word/settings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W_NS}"><w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:themeFontLang w:val="${langTag}"/></w:settings>`);
  zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title></dc:title><dc:creator></dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Documents Workbench</Application></Properties>`);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
}

// ── stan podglądu ────────────────────────────────────────────────────────────
let docComposeStyleClasses = new Map(); // klasa CSS akapitu → klucz stylu (odświeżane przy renderze)

function composeStyleKeyOf(p) {
  if (!p) return "normal";
  for (const cls of p.classList) {
    const key = docComposeStyleClasses.get(cls);
    if (key && COMPOSE_STYLE_KEYS.includes(key)) return key;
  }
  for (const cls of p.classList) { // nagłówek w stylu własnym pliku (np. „Rozdział” z outlineLvl)
    const level = docHeadingStyleClasses?.get?.(cls);
    if (level) return level <= 3 ? `h${level}` : "h3";
  }
  for (const cls of p.classList) { // własny styl pliku („custom:Wskazówka”)
    const key = docComposeStyleClasses.get(cls);
    if (key?.startsWith("custom:")) return key;
  }
  return "normal";
}

// Nowy akapit po Enterze na końcu nagłówka/tytułu/cytatu: zdejmij klasę stylu w podglądzie.
// Zwraca true, gdy akapit miał taki styl (plik dostaje wtedy nextNormal).
function composeStripNextStyle(newP) {
  let had = false;
  Array.from(newP.classList).forEach((cls) => {
    const key = docComposeStyleClasses.get(cls);
    // własny styl pliku zostaje (jak w Wordzie bez „następnego stylu”) — wyjście: Enter w pustym akapicie
    if ((key && key !== "normal" && !key.startsWith("custom:")) || docHeadingStyleClasses?.get?.(cls)) {
      newP.classList.remove(cls);
      had = true;
    }
  });
  return had;
}

// ── domknięcie zapisu: nowe linki ────────────────────────────────────────────
// applyRunsToParagraphXml nie ma dostępu do paczki, więc nowy adres zostawia jako
// atrybut dwb-href, a styl znakowy linku jako „__DWB_HL__”. Tu: powiązanie (rels) + styl.
const COMPOSE_REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

async function composeReadRels(zip) {
  const path = "word/_rels/document.xml.rels";
  const xml = zip.file(path) ? await zip.file(path).async("string")
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  return { path, doc: composeParse(xml) };
}

function composeAddRel(relsDoc, type, target, external) {
  const root = relsDoc.documentElement;
  const rels = Array.from(root.getElementsByTagName("Relationship"));
  const hit = rels.find((r) => r.getAttribute("Type") === type && r.getAttribute("Target") === target);
  if (hit) return hit.getAttribute("Id");
  const ids = new Set(rels.map((r) => r.getAttribute("Id")));
  let n = rels.length + 1;
  while (ids.has(`rId${n}`)) n++;
  const el = relsDoc.createElementNS(root.namespaceURI, "Relationship");
  el.setAttribute("Id", `rId${n}`);
  el.setAttribute("Type", type);
  el.setAttribute("Target", target);
  if (external) el.setAttribute("TargetMode", "External");
  root.appendChild(el);
  return `rId${n}`;
}

async function composeEnsureCharStyle(zip, name, id, body) {
  const xml = await composeEnsureStylesPart(zip);
  const doc = composeParse(xml);
  const found = Array.from(doc.getElementsByTagNameNS(W_NS, "style")).find((st) => st.getAttributeNS(W_NS, "type") === "character"
    && (composeDirectChild(st, "name")?.getAttributeNS(W_NS, "val") || "").trim().toLowerCase() === name.toLowerCase());
  if (found) { zip.file("word/styles.xml", xml); return found.getAttributeNS(W_NS, "styleId"); }
  const frag = composeParse(`<w:styles xmlns:w="${W_NS}"><w:style w:type="character" w:styleId="${id}"><w:name w:val="${name}"/>${body}</w:style></w:styles>`);
  doc.documentElement.appendChild(doc.importNode(frag.documentElement.firstChild, true));
  zip.file("word/styles.xml", composeSerialize(doc));
  return id;
}

async function finalizeComposeParts(zip, xml) {
  if (!xml.includes("dwb-href") && !xml.includes("__DWB_HL__")) return xml;
  if (xml.includes("__DWB_HL__")) {
    const id = await composeEnsureCharStyle(zip, "Hyperlink", "Hyperlink",
      `<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>`);
    xml = xml.replaceAll('"__DWB_HL__"', `"${id}"`);
  }
  if (xml.includes("dwb-href")) {
    const doc = composeParse(xml);
    const rels = await composeReadRels(zip);
    Array.from(doc.getElementsByTagNameNS(W_NS, "hyperlink")).forEach((h) => {
      const href = h.getAttribute("dwb-href");
      if (href == null) return;
      h.removeAttribute("dwb-href");
      h.setAttributeNS(R_NS, "r:id", composeAddRel(rels.doc, COMPOSE_REL_HYPERLINK, href, true));
    });
    zip.file(rels.path, composeSerialize(rels.doc));
    xml = composeSerialize(doc);
  }
  return xml;
}

// ── op "link": wstaw / zmień / usuń link w akapicie ──────────────────────────
// { index, start, end, href? | targetIndex?, text?, remove? } — przesunięcia w tekście akapitu
// (łamanie wiersza = 1 znak, jak previewRunsToPlainText). targetIndex = akapit-cel w dokumencie
// (nagłówek): dostaje zakładkę (istniejąca zostaje użyta), link prowadzi do „#zakładka”.
// targetImage = który obraz w akapicie-celu (od 0, jak docImageTargets w podglądzie): zakładka
// obejmuje fragment z TYM obrazem (Word: zaznacz obraz → Zakładka) — przy kilku obrazach w akapicie
// albo obrazie za podziałem strony link prowadzi do właściwego, nie do początku akapitu.
const composeIsBookmarkStart = (n) => n?.localName === "bookmarkStart" && n.namespaceURI === W_NS && !/^_GoBack$/.test(n.getAttributeNS(W_NS, "name"));
// Fragmenty (w:r) z obrazem w akapicie — w kolejności, bez akapitów zagnieżdżonych (pole tekstowe).
function composeImageRuns(p) {
  return Array.from(p.getElementsByTagNameNS(W_NS, "r")).filter((r) => {
    for (let a = r.parentNode; a && a !== p; a = a.parentNode) if (a.localName === "p" && a.namespaceURI === W_NS) return false;
    return Array.from(r.getElementsByTagName("*")).some((n) => n.localName === "blip" || n.localName === "imagedata");
  });
}
// Zakładka tuż przed fragmentem (pomijając inne znaczniki zakładek między nimi).
function composeBookmarkBefore(run) {
  for (let n = run.previousSibling; n; n = n.previousSibling) {
    if (n.nodeType !== 1) continue;
    if (composeIsBookmarkStart(n)) return n;
    if (!(n.namespaceURI === W_NS && /^bookmark(Start|End)$/.test(n.localName))) return null;
  }
  return null;
}
function composeNewBookmark(doc) {
  const all = Array.from(doc.getElementsByTagNameNS(W_NS, "bookmarkStart"));
  const names = new Set(all.map((b) => b.getAttributeNS(W_NS, "name")));
  const maxId = all.reduce((m, b) => Math.max(m, parseInt(b.getAttributeNS(W_NS, "id"), 10) || 0), 0);
  let name;
  do { name = `_Ref${String(Math.floor(1e8 + Math.random() * 9e8))}`; } while (names.has(name));
  const id = String(maxId + 1);
  return { name, start: composeEl(doc, "bookmarkStart", { id, name }), end: composeEl(doc, "bookmarkEnd", { id }) };
}
function composeEnsureBookmark(doc, p) {
  // zakładka akapitu — nie ta, która stoi przy obrazie (cel linku do obrazu)
  const imageRuns = new Set(composeImageRuns(p));
  const own = Array.from(p.childNodes).find((n) => composeIsBookmarkStart(n) && ![...imageRuns].some((r) => composeBookmarkBefore(r) === n));
  if (own) return own.getAttributeNS(W_NS, "name");
  const { name, start, end } = composeNewBookmark(doc);
  const pPr = composeDirectChild(p, "pPr");
  p.insertBefore(start, pPr ? pPr.nextSibling : p.firstChild);
  p.appendChild(end);
  return name;
}
function composeEnsureImageBookmark(doc, p, nth) {
  const run = composeImageRuns(p)[nth];
  if (!run) return composeEnsureBookmark(doc, p);
  const own = composeBookmarkBefore(run);
  if (own) return own.getAttributeNS(W_NS, "name");
  const { name, start, end } = composeNewBookmark(doc);
  run.parentNode.insertBefore(start, run);
  run.parentNode.insertBefore(end, run.nextSibling);
  return name;
}

// Rozcina fragmenty tak, żeby [start, end) było osobno; zwraca { before, mid, after }.
function composeSliceRuns(runs, start, end) {
  const before = []; const mid = []; const after = [];
  let pos = 0;
  (runs || []).forEach((run) => {
    const len = run.break ? 1 : (run.text || "").length;
    const a = pos; const b = pos + len;
    pos = b;
    if (run.break) { (b <= start ? before : a >= end ? after : mid).push(run); return; }
    if (run.island) { // pole formularza — niepodzielne
      if (b <= start || (start === end && a < start)) before.push(run);
      else if (a >= end || start === end) after.push(run);
      else mid.push(run);
      return;
    }
    const cut = (from, to) => ({ ...run, text: run.text.slice(from - a, to - a) });
    if (b <= start) before.push(run);
    else if (a >= end) after.push(run);
    else {
      if (a < start) before.push(cut(a, start));
      mid.push(cut(Math.max(a, start), Math.min(b, end)));
      if (b > end) after.push(cut(end, b));
    }
  });
  return { before, mid, after };
}

function applyLinkInXml(xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[edit.index];
  if (!p) return { xml, count: 0 };
  let link = null;
  if (!edit.remove) {
    // gotowa zakładka (Word: Miejsce w tym dokumencie → Zakładki) albo „_top” = Początek dokumentu
    if (edit.ref && /^rel:/.test(edit.ref)) {
      link = edit.ref; // niezmieniony adres — to samo powiązanie w pliku
    } else if (edit.anchor) {
      if (edit.anchor !== "_top" && !composeFindBookmark(doc, edit.anchor)) return { xml, count: 0 };
      link = `#${edit.anchor}`;
    } else if (Number.isInteger(edit.targetIndex)) {
      const target = paragraphs[edit.targetIndex];
      if (!target) return { xml, count: 0 };
      link = `#${Number.isInteger(edit.targetImage) ? composeEnsureImageBookmark(doc, target, edit.targetImage) : composeEnsureBookmark(doc, target)}`;
    } else {
      link = String(edit.href || "").trim();
    }
    if (!link) return { xml, count: 0 };
  }
  const runs = extractRunsFromParagraphXml(p);
  const total = previewRunsToPlainText(runs).length;
  const start = Math.max(0, Math.min(total, edit.start | 0));
  const end = Math.max(start, Math.min(total, edit.end | 0));
  const { before, mid, after } = composeSliceRuns(runs, start, end);
  let middle = mid;
  const curText = previewRunsToPlainText(mid);
  if (edit.text != null && edit.text !== curText) {
    // nowy tekst linku w formatowaniu miejsca (bez dawnego linku)
    const base = { ...(mid.find((r) => !r.break) || before.filter((r) => !r.break).pop() || {}) };
    delete base.break;
    middle = edit.text ? [{ ...base, text: String(edit.text) }] : [];
  }
  middle = middle.map((r) => {
    if (r.break) return r;
    const out = { ...r };
    if (link) out.link = link; else delete out.link;
    return out;
  });
  applyRunsToParagraphXml(p, [...before, ...middle, ...after]);
  // etykietka ekranowa (Word: „Etykietka ekranowa…” → w:tooltip) — na linkach tego miejsca
  if (link && edit.tooltip != null) {
    const tip = String(edit.tooltip).trim().slice(0, 255);
    Array.from(p.getElementsByTagNameNS(W_NS, "hyperlink")).forEach((h) => {
      if (h._dwbLink !== link) return;
      if (tip) h.setAttributeNS(W_NS, "w:tooltip", tip); else h.removeAttributeNS(W_NS, "tooltip");
    });
  }
  return { xml: composeSerialize(doc), count: 1 };
}

// ── op "bookmark": zakładki (Word: Wstaw → Zakładka) ─────────────────────────
// { action: "add", name, from: { index, offset }, to: { index, offset } } — zakładka obejmuje
//   tekst od–do (kursor bez zaznaczenia = pusty punkt). Ta sama nazwa już jest → zakładka
//   przenosi się w nowe miejsce (jak w Wordzie: „Dodaj” z istniejącą nazwą).
// { action: "delete", name } — usuwa zakładkę (tekst zostaje). Linki do niej przestają działać —
//   okienko ostrzega wcześniej, ile ich jest.
// Nazwa jak w Wordzie: zaczyna się literą, dalej litery, cyfry i „_”, do 40 znaków.
const BOOKMARK_NAME_RE = /^[\p{L}][\p{L}\p{N}_]{0,39}$/u;
function composeFindBookmark(doc, name) {
  return Array.from(doc.getElementsByTagNameNS(W_NS, "bookmarkStart")).find((b) => b.getAttributeNS(W_NS, "name") === name) || null;
}
function composeRemoveBookmark(doc, name) {
  const start = composeFindBookmark(doc, name);
  if (!start) return false;
  const id = start.getAttributeNS(W_NS, "id");
  Array.from(doc.getElementsByTagNameNS(W_NS, "bookmarkEnd")).filter((e) => e.getAttributeNS(W_NS, "id") === id).forEach((e) => e.parentNode.removeChild(e));
  start.parentNode.removeChild(start);
  return true;
}
// Znacznik o zerowej długości w miejscu tekstu akapitu (jak zakres komentarza: model akapitu).
function composeInsertMarker(p, offset, markerXml) {
  const runs = extractRunsFromParagraphXml(p);
  const total = previewRunsToPlainText(runs).length;
  const at = Math.max(0, Math.min(total, offset | 0));
  const { before, after } = composeSliceRuns(runs, at, at);
  applyRunsToParagraphXml(p, [...before, { island: markerXml, text: "" }, ...after]);
}
function applyBookmarkInXml(xml, edit) {
  const doc = composeParse(xml);
  const name = String(edit.name || "");
  if (edit.action === "delete") return composeRemoveBookmark(doc, name) ? { xml: composeSerialize(doc), count: 1 } : { xml, count: 0 };
  if (edit.action !== "add" || !BOOKMARK_NAME_RE.test(name)) return { xml, count: 0 };
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  let from = edit.from || {};
  let to = edit.to || from;
  if (to.index < from.index || (to.index === from.index && to.offset < from.offset)) [from, to] = [to, from];
  const pFrom = paragraphs[from.index];
  const pTo = paragraphs[to.index];
  if (!pFrom || !pTo) return { xml, count: 0 };
  composeRemoveBookmark(doc, name); // ta sama nazwa = przeniesienie
  const all = Array.from(doc.getElementsByTagNameNS(W_NS, "bookmarkStart"));
  const id = String(all.reduce((m, b) => Math.max(m, parseInt(b.getAttributeNS(W_NS, "id"), 10) || 0), 0) + 1);
  const esc = name.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
  if (edit.whole) {
    // akapit tylko do odczytu (pole, zmiany, …) — bez przepisywania: zakładka na całe akapity
    const start = composeEl(doc, "bookmarkStart", { id, name });
    const end = composeEl(doc, "bookmarkEnd", { id });
    const pPr = composeDirectChild(pFrom, "pPr");
    pFrom.insertBefore(start, pPr ? pPr.nextSibling : pFrom.firstChild);
    pTo.appendChild(end);
  } else if (pFrom === pTo) {
    const runs = extractRunsFromParagraphXml(pFrom);
    const total = previewRunsToPlainText(runs).length;
    const a = Math.max(0, Math.min(total, from.offset | 0));
    const b = Math.max(a, Math.min(total, to.offset | 0));
    const { before, mid, after } = composeSliceRuns(runs, a, b);
    applyRunsToParagraphXml(pFrom, [...before,
      { island: `<w:bookmarkStart xmlns:w="${W_NS}" w:id="${id}" w:name="${esc}"/>`, text: "" }, ...mid,
      { island: `<w:bookmarkEnd xmlns:w="${W_NS}" w:id="${id}"/>`, text: "" }, ...after]);
  } else {
    composeInsertMarker(pTo, to.offset, `<w:bookmarkEnd xmlns:w="${W_NS}" w:id="${id}"/>`);
    composeInsertMarker(pFrom, from.offset, `<w:bookmarkStart xmlns:w="${W_NS}" w:id="${id}" w:name="${esc}"/>`);
  }
  return { xml: composeSerialize(doc), count: 1 };
}

// ── op "list": lista punktowana / numerowana / bez listy ─────────────────────
// Definicje list (numbering.xml) dopisujemy własne, rozpoznawane po w:name „DWB …” —
// kolejne listy używają ich ponownie. Numerowana lista zaczyna od 1 (startOverride), chyba że
// akapit tuż wyżej jest w tej samej numeracji — wtedy ją kontynuuje (jak w Wordzie).
const COMPOSE_LIST_NAMES = { bullet: "DWB Bullets", number: "DWB Numbering" };
const COMPOSE_BULLETS = ["•", "◦", "▪"];
const COMPOSE_NUMFMT = [["decimal", "%L."], ["lowerLetter", "%L."], ["lowerRoman", "%L."]];

function composeAbstractNumXml(kind, id) {
  const nsid = Math.floor(Math.random() * 0xffffffff).toString(16).toUpperCase().padStart(8, "0");
  let lvls = "";
  for (let i = 0; i < 9; i++) {
    const ind = `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr>`;
    if (kind === "bullet") {
      lvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${COMPOSE_BULLETS[i % 3]}"/><w:lvlJc w:val="left"/>${ind}<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/></w:rPr></w:lvl>`;
    } else {
      const [fmt, text] = COMPOSE_NUMFMT[i % 3];
      lvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text.replace("L", i + 1)}"/><w:lvlJc w:val="left"/>${ind}</w:lvl>`;
    }
  }
  return `<w:abstractNum w:abstractNumId="${id}"><w:nsid w:val="${nsid}"/><w:multiLevelType w:val="hybridMultilevel"/><w:name w:val="${COMPOSE_LIST_NAMES[kind]}"/>${lvls}</w:abstractNum>`;
}

async function composeEnsurePart(zip, path, rootXml, contentType, relType) {
  const file = zip.file(path);
  if (file) return file.async("string");
  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    let ct = await ctFile.async("string");
    if (!ct.includes(`PartName="/${path}"`)) {
      ct = ct.replace("</Types>", `<Override PartName="/${path}" ContentType="${contentType}"/></Types>`);
      zip.file("[Content_Types].xml", ct);
    }
  }
  const rels = await composeReadRels(zip);
  composeAddRel(rels.doc, relType, path.replace(/^word\//, ""), false);
  zip.file(rels.path, composeSerialize(rels.doc));
  return rootXml;
}

// Miejsce w numbering.xml zgodne ze schematem: numPicBullet* → abstractNum* → num* → numIdMacAtCleanup.
// Dawniej w:num szło na sam koniec — w pliku z Worda na Macu (w:numIdMacAtCleanup na końcu) za
// tym elementem, czyli błąd schematu (walidator Open XML SDK, audyt prawdziwych plików 2026-10-05).
function composeNumberingInsert(root, el) {
  const before = el.localName === "abstractNum" ? (composeDirectChild(root, "num") || composeDirectChild(root, "numIdMacAtCleanup")) : composeDirectChild(root, "numIdMacAtCleanup");
  root.insertBefore(el, before || null);
}

function composeNumberingIndex(numDoc) {
  const abstracts = new Map(); // abstractNumId → { el, name }
  Array.from(numDoc.getElementsByTagNameNS(W_NS, "abstractNum")).forEach((a) => {
    abstracts.set(a.getAttributeNS(W_NS, "abstractNumId"), { el: a, name: composeDirectChild(a, "name")?.getAttributeNS(W_NS, "val") || "" });
  });
  const nums = new Map(); // numId → abstractNumId
  Array.from(numDoc.getElementsByTagNameNS(W_NS, "num")).forEach((n) => {
    nums.set(n.getAttributeNS(W_NS, "numId"), composeDirectChild(n, "abstractNumId")?.getAttributeNS(W_NS, "val"));
  });
  return { abstracts, nums };
}

// Rodzaj listy akapitu: "bullet" / "number" / null.
function composeParagraphListKind(p, idx) {
  const numPr = composeDirectChild(composeDirectChild(p, "pPr") || p, "numPr");
  if (!numPr) return null;
  const numId = composeDirectChild(numPr, "numId")?.getAttributeNS(W_NS, "val");
  if (!numId || numId === "0") return null;
  const ilvl = composeDirectChild(numPr, "ilvl")?.getAttributeNS(W_NS, "val") || "0";
  const abs = idx.abstracts.get(idx.nums.get(numId));
  const lvl = abs && Array.from(abs.el.getElementsByTagNameNS(W_NS, "lvl")).find((l) => l.getAttributeNS(W_NS, "ilvl") === ilvl);
  const fmt = lvl && composeDirectChild(lvl, "numFmt")?.getAttributeNS(W_NS, "val");
  return fmt === "bullet" ? "bullet" : fmt && fmt !== "none" ? "number" : "number";
}

async function applyListInZip(zip, xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const targets = (edit.indices || []).map((i) => paragraphs[i]).filter(Boolean);
  if (!targets.length) return { xml, count: 0 };
  const numPath = "word/numbering.xml";
  let numXml = zip.file(numPath) ? await zip.file(numPath).async("string") : null;
  const numDoc = composeParse(numXml || `<w:numbering xmlns:w="${W_NS}"/>`);
  const idx = composeNumberingIndex(numDoc);
  let kind = edit.kind;
  // przełącznik: wszystkie zaznaczone już są taką listą → zdejmij listę
  if (edit.toggle && kind !== "none" && targets.every((p) => composeParagraphListKind(p, idx) === kind)) kind = "none";
  if (kind === "none") {
    let n = 0;
    targets.forEach((p) => {
      const pPr = composeDirectChild(p, "pPr");
      if (pPr && composeDirectChild(pPr, "numPr")) { composeSetPPrChild(pPr, "numPr", null); n++; }
      // styl „Akapit z listą” bez listy to tylko wcięcie — zostaje (jak w Wordzie)
    });
    return { xml: n ? composeSerialize(doc) : xml, count: n, kind };
  }
  // definicja (abstractNum) tego rodzaju
  let absId = Array.from(idx.abstracts.entries()).find(([, a]) => a.name === COMPOSE_LIST_NAMES[kind])?.[0];
  const root = numDoc.documentElement;
  if (absId == null) {
    absId = String(Array.from(idx.abstracts.keys()).reduce((m, k) => Math.max(m, parseInt(k, 10) || 0), -1) + 1);
    const frag = composeParse(`<w:numbering xmlns:w="${W_NS}">${composeAbstractNumXml(kind, absId)}</w:numbering>`);
    composeNumberingInsert(root, numDoc.importNode(frag.documentElement.firstChild, true)); // abstractNum przed num (schemat)
  }
  const newNum = (restart) => {
    const id = String(Array.from(idx.nums.keys()).reduce((m, k) => Math.max(m, parseInt(k, 10) || 0), 0) + 1);
    const num = composeEl(numDoc, "num", { numId: id });
    num.appendChild(composeEl(numDoc, "abstractNumId", { val: absId }));
    if (restart) {
      const ov = composeEl(numDoc, "lvlOverride", { ilvl: 0 });
      ov.appendChild(composeEl(numDoc, "startOverride", { val: 1 }));
      num.appendChild(ov);
    }
    composeNumberingInsert(root, num);
    idx.nums.set(id, absId);
    return id;
  };
  // numer listy: kontynuacja akapitu wyżej (ta sama definicja) albo nowy
  const first = paragraphs.indexOf(targets[0]);
  const prev = first > 0 ? paragraphs[first - 1] : null;
  const prevNumId = prev && composeDirectChild(composeDirectChild(composeDirectChild(prev, "pPr") || prev, "numPr") || prev, "numId")?.getAttributeNS(W_NS, "val");
  let numId = prevNumId && idx.nums.get(prevNumId) === absId ? prevNumId : null;
  if (!numId && kind === "bullet") numId = Array.from(idx.nums.entries()).find(([, a]) => a === absId)?.[0] || null;
  if (!numId) numId = newNum(kind === "number");
  targets.forEach((p) => {
    const pPr = composeEnsurePPr(p);
    const old = composeDirectChild(pPr, "numPr");
    const ilvl = old && composeDirectChild(old, "ilvl")?.getAttributeNS(W_NS, "val");
    const numPr = composeEl(doc, "numPr");
    numPr.appendChild(composeEl(doc, "ilvl", { val: ilvl || 0 }));
    numPr.appendChild(composeEl(doc, "numId", { val: numId }));
    composeSetPPrChild(pPr, "numPr", numPr);
  });
  numXml = await composeEnsurePart(zip, numPath, composeSerialize(numDoc),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
  zip.file(numPath, composeSerialize(numDoc));
  return { xml: composeSerialize(doc), count: targets.length, kind };
}

// Czy w podglądzie jest spis treści (akapity w stylach „toc 1–3”).
function composeHasToc() {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host) return false;
  return collectPreviewParagraphElements(host).some((p) => Array.from(p.classList).some((c) => COMPOSE_TOC_KEYS.includes(docComposeStyleClasses.get(c))));
}

// ── op "toc": spis treści — wstaw albo zaktualizuj ───────────────────────────
// Prawdziwe pole Worda TOC \o "1-3" \h \z \u: wpisy z linkami do zakładek przy nagłówkach i
// numerami stron z podglądu (pages: { indeksAkapitu: strona }). Word i tak przelicza je przy
// „Aktualizuj pole”. Istniejący spis (pole TOC) jest podmieniany w miejscu — także ten z Worda.
function composeHeadingLevels(stylesDoc) {
  const styles = new Map();
  if (stylesDoc) Array.from(stylesDoc.getElementsByTagNameNS(W_NS, "style")).forEach((st) => {
    if (st.getAttributeNS(W_NS, "type") !== "paragraph") return;
    const id = st.getAttributeNS(W_NS, "styleId");
    const pPr = composeDirectChild(st, "pPr");
    const ol = pPr && composeDirectChild(pPr, "outlineLvl")?.getAttributeNS(W_NS, "val");
    styles.set(id, {
      name: (composeDirectChild(st, "name")?.getAttributeNS(W_NS, "val") || "").toLowerCase(),
      basedOn: composeDirectChild(st, "basedOn")?.getAttributeNS(W_NS, "val"),
      outline: ol == null ? null : Number(ol),
    });
  });
  const levelOf = (id, depth = 0) => {
    const st = styles.get(id);
    if (!st || depth > 12) return 0;
    if (Number.isFinite(st.outline)) return st.outline < 9 ? st.outline + 1 : 0;
    const m = st.name.match(/^heading\s*(\d)$/);
    if (m) return Number(m[1]);
    return st.basedOn ? levelOf(st.basedOn, depth + 1) : 0;
  };
  return (p) => {
    const pPr = composeDirectChild(p, "pPr");
    const direct = pPr && composeDirectChild(pPr, "outlineLvl")?.getAttributeNS(W_NS, "val");
    if (direct != null) { const n = Number(direct); return n < 9 ? n + 1 : 0; }
    const sid = pPr && composeDirectChild(pPr, "pStyle")?.getAttributeNS(W_NS, "val");
    return sid ? levelOf(sid) : 0;
  };
}

function composeFieldInstr(beginRun) {
  let instr = "";
  for (let n = beginRun.nextSibling; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    if (composeDirectChild(n, "fldChar")) break;
    Array.from(n.getElementsByTagNameNS(W_NS, "instrText")).forEach((it) => { instr += it.textContent || ""; });
  }
  return instr.trim();
}

// [akapit z początkiem pola TOC, akapit z jego końcem] albo null.
function composeFindToc(paragraphs) {
  for (let i = 0; i < paragraphs.length; i++) {
    const begin = Array.from(paragraphs[i].getElementsByTagNameNS(W_NS, "fldChar")).find((fc) => fc.getAttributeNS(W_NS, "fldCharType") === "begin");
    if (!begin || !/^TOC\b/i.test(composeFieldInstr(begin.parentNode))) continue;
    let depth = 0;
    let started = false;
    for (let j = i; j < paragraphs.length; j++) {
      for (const fc of Array.from(paragraphs[j].getElementsByTagNameNS(W_NS, "fldChar"))) {
        if (!started) { if (fc === begin) started = true; else continue; }
        const type = fc.getAttributeNS(W_NS, "fldCharType");
        if (type === "begin") depth++;
        else if (type === "end" && --depth === 0) return [i, j];
      }
    }
    return null;
  }
  return null;
}

async function applyTocInZip(zip, xml, edit) {
  const doc = composeParse(xml);
  const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const cache = new Map();
  const ids = {};
  for (const k of ["tocHeading", ...COMPOSE_TOC_KEYS]) ids[k] = await composeEnsureStyle(zip, k, cache);
  const stylesDoc = composeParse(cache.stylesXml);
  const levelOf = composeHeadingLevels(stylesDoc);
  const existing = composeFindToc(paragraphs);
  const inToc = (i) => existing && i >= existing[0] && i <= existing[1];
  const pages = edit.pages || {};
  const heads = [];
  paragraphs.forEach((p, i) => {
    if (inToc(i) || p.parentNode !== body) return; // nagłówki w tabelach pomijamy (jak domyślnie Word)
    const level = levelOf(p);
    const text = getParagraphText(p).replace(/\s+/g, " ").trim();
    if (level >= 1 && level <= 3 && text) heads.push({ p, i, level, text });
  });
  // tabulator do prawego marginesu z kropkami — marginesy SEKCJI, w której stoi spis (dawniej
  // zawsze ostatniej sekcji dokumentu: spis na pionowej stronie przed sekcją poziomą wyjeżdżał
  // numerami stron za margines)
  const sect = composeDirectChild(body, "sectPr");
  const tabPos = Math.max(2000, composeTextWidthTwips(doc, paragraphs[existing ? existing[0] : edit.index] || null));
  const run = (inner) => { const r = composeEl(doc, "r"); inner.forEach((c) => r.appendChild(c)); return r; };
  const text = (s) => { const t = composeEl(doc, "t"); t.setAttribute("xml:space", "preserve"); t.textContent = s; return t; };
  const fld = (type) => run([composeEl(doc, "fldChar", { fldCharType: type })]);
  const entries = heads.map((h) => {
    const bm = composeEnsureBookmark(doc, h.p);
    const p = composeEl(doc, "p");
    const pPr = composeEnsurePPr(p);
    composeSetPPrChild(pPr, "pStyle", { val: ids[`toc${h.level}`] });
    const tabs = composeEl(doc, "tabs");
    tabs.appendChild(composeEl(doc, "tab", { val: "right", leader: "dot", pos: tabPos }));
    composeSetPPrChild(pPr, "tabs", tabs);
    const hl = composeEl(doc, "hyperlink", { anchor: bm, history: 1 });
    hl.appendChild(run([text(h.text)]));
    hl.appendChild(run([composeEl(doc, "tab")]));
    hl.appendChild(run([text(pages[h.i] ? String(pages[h.i]) : "")]));
    p.appendChild(hl);
    return p;
  });
  if (!entries.length) {
    const p = composeEl(doc, "p");
    composeSetPPrChild(composeEnsurePPr(p), "pStyle", { val: ids.toc1 });
    p.appendChild(run([text(edit.emptyText || "—")]));
    entries.push(p);
  }
  // pole: początek w pierwszym wpisie, koniec w ostatnim
  const instr = composeEl(doc, "instrText");
  instr.setAttribute("xml:space", "preserve");
  instr.textContent = ' TOC \\o "1-3" \\h \\z \\u ';
  const first = entries[0];
  const afterPPr = composeDirectChild(first, "pPr")?.nextSibling || null;
  [fld("begin"), run([instr]), fld("separate")].forEach((r) => first.insertBefore(r, afterPPr));
  entries[entries.length - 1].appendChild(fld("end"));
  if (existing) {
    const [a, b] = existing;
    const anchor = paragraphs[a];
    entries.forEach((p) => anchor.parentNode.insertBefore(p, anchor));
    for (let i = a; i <= b; i++) {
      const old = paragraphs[i];
      composeMoveSectPr(old, entries[entries.length - 1]);
      old.parentNode.removeChild(old);
    }
  } else {
    const at = paragraphs[edit.index] || null;
    const parent = at ? at.parentNode : body;
    const before = at || sect;
    const title = composeEl(doc, "p");
    composeSetPPrChild(composeEnsurePPr(title), "pStyle", { val: ids.tocHeading });
    title.appendChild(run([text(edit.title || "Spis treści")]));
    [title, ...entries].forEach((p) => parent.insertBefore(p, before));
  }
  if (cache.stylesDirty) zip.file("word/styles.xml", cache.stylesXml);
  return { xml: composeSerialize(doc), count: 1, updated: !!existing };
}

// ── op "formInsert": nowe pole formularza w miejscu kursora ──────────────────
// Kontrolka zawartości jak z Worda (Deweloper → Formanty): tekst / lista / data / pole wyboru,
// z tekstem zastępczym w stylu „Tekst zastępczy”. Do akapitu trafia jako pole-wyspa.
const COMPOSE_W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const COMPOSE_PH = {
  pl: { text: "Kliknij lub naciśnij tutaj, aby wprowadzić tekst.", date: "Kliknij lub naciśnij, aby wprowadzić datę.", dropdown: "Wybierz element." },
  en: { text: "Click or tap here to enter text.", date: "Click or tap to enter a date.", dropdown: "Choose an item." },
};

function composeSdtXml(edit, phStyleId) {
  const ph = COMPOSE_PH[edit.lang === "en" ? "en" : "pl"];
  const id = String(Math.floor(1e8 + Math.random() * 9e8));
  const esc = (v) => composeXmlText(v).replace(/"/g, "&quot;");
  const label = edit.label ? `<w:alias w:val="${esc(edit.label)}"/><w:tag w:val="${esc(edit.label)}"/>` : "";
  const phRun = (text) => `<w:r><w:rPr><w:rStyle w:val="${phStyleId}"/></w:rPr><w:t xml:space="preserve">${composeXmlText(text)}</w:t></w:r>`;
  let pr = "";
  let content = "";
  if (edit.kind === "checkbox") {
    pr = `<w14:checkbox><w14:checked w14:val="${edit.checked ? 1 : 0}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>`;
    content = `<w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/></w:rPr><w:t>${edit.checked ? "☒" : "☐"}</w:t></w:r>`;
    return `<w:sdt xmlns:w="${W_NS}" xmlns:w14="${COMPOSE_W14}"><w:sdtPr>${label}<w:id w:val="${id}"/>${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  }
  if (edit.kind === "date") {
    pr = `<w:date><w:dateFormat w:val="dd.MM.yyyy"/><w:lid w:val="${edit.lang === "en" ? "en-GB" : "pl-PL"}"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>`;
    content = phRun(ph.date);
  } else if (edit.kind === "dropdown") {
    const items = (edit.options || []).map((o) => String(o).trim()).filter(Boolean);
    pr = `<w:dropDownList><w:listItem w:displayText="${esc(ph.dropdown)}" w:value=""/>${items.map((o) => `<w:listItem w:displayText="${esc(o)}" w:value="${esc(o)}"/>`).join("")}</w:dropDownList>`;
    content = phRun(ph.dropdown);
  } else {
    pr = "<w:text/>";
    content = phRun(ph.text);
  }
  return `<w:sdt xmlns:w="${W_NS}"><w:sdtPr>${label}<w:id w:val="${id}"/><w:showingPlcHdr/>${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
}

async function applyFormInsertInZip(zip, xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[edit.index];
  if (!p || !["text", "date", "dropdown", "checkbox"].includes(edit.kind)) return { xml, count: 0 };
  const phId = await composeEnsureCharStyle(zip, "Placeholder Text", "PlaceholderText",
    `<w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:color w:val="666666"/></w:rPr>`);
  const sdtXml = composeSdtXml(edit, phId);
  const sdtEl = composeParse(sdtXml).documentElement;
  const runs = extractRunsFromParagraphXml(p);
  const total = previewRunsToPlainText(runs).length;
  const at = Math.max(0, Math.min(total, edit.offset | 0));
  const { before, after } = composeSliceRuns(runs, at, at);
  const island = { island: new XMLSerializer().serializeToString(sdtEl), text: ffText(ffKid(sdtEl, "sdtContent")) };
  applyRunsToParagraphXml(p, [...before, island, ...after]);
  return { xml: composeSerialize(doc), count: 1 };
}

// ── op "snippetInsert": snippet z polami formularza Worda (snippet-suggest.js) ─────────
// edit: { index, offset, deleteLen (wpisany „!nazwa” od offset), parts: [{ text } | { field:
// { kind, options, label } }], style (wygląd tekstu w miejscu kursora), lang }. Tekst → fragmenty
// (łamania wierszy z „\n”), pole → kontrolka jak z „Wstaw → pole formularza” (wyspa).
async function applySnippetInsertInZip(zip, xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[edit.index];
  if (!p || !Array.isArray(edit.parts)) return { xml, count: 0 };
  const needPh = edit.parts.some((x) => x.field && x.field.kind !== "checkbox");
  const phId = needPh ? await composeEnsureCharStyle(zip, "Placeholder Text", "PlaceholderText",
    `<w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:color w:val="666666"/></w:rPr>`) : "PlaceholderText";
  const runs = extractRunsFromParagraphXml(p);
  const total = previewRunsToPlainText(runs).length;
  const at = Math.max(0, Math.min(total, edit.offset | 0));
  const { before } = composeSliceRuns(runs, at, at);
  const { after } = composeSliceRuns(runs, Math.min(total, at + Math.max(0, edit.deleteLen | 0)), Math.min(total, at + Math.max(0, edit.deleteLen | 0)));
  const style = edit.style || {};
  const mid = [];
  edit.parts.forEach((part) => {
    if (part.field) {
      const kind = ["text", "date", "dropdown", "checkbox"].includes(part.field.kind) ? part.field.kind : "text";
      const sdtEl = composeParse(composeSdtXml({ kind, options: part.field.options, label: part.field.label, lang: edit.lang }, phId)).documentElement;
      mid.push({ island: new XMLSerializer().serializeToString(sdtEl), text: ffText(ffKid(sdtEl, "sdtContent")) });
      return;
    }
    const text = String(part.text || "").split(edit.cursorMark || "\uE000").join("");
    text.split("\n").forEach((line, i) => {
      if (i) mid.push({ break: true });
      if (line) mid.push({ text: line, ...style });
    });
  });
  applyRunsToParagraphXml(p, [...before, ...mid, ...after]);
  return { xml: composeSerialize(doc), count: 1 };
}

// ── op "pageVAlign": wyrównanie strony w pionie (Word: Ustawienia strony → Układ) ─────────
// Sekcja akapitu z kursorem: w:vAlign top (domyślne — bez wpisu) / center / both / bottom.
function applyPageVAlignInXml(xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  if (!["top", "center", "both", "bottom"].includes(edit.val)) return { xml, count: 0 };
  let sect = null;
  for (let i = Math.max(0, edit.index | 0); i < paragraphs.length && !sect; i++) {
    const pPr = composeDirectChild(paragraphs[i], "pPr");
    sect = pPr && composeDirectChild(pPr, "sectPr");
  }
  if (!sect) sect = composeDirectChild(doc.getElementsByTagNameNS(W_NS, "body")[0], "sectPr");
  if (!sect) return { xml, count: 0 };
  const old = composeDirectChild(sect, "vAlign");
  if (old) sect.removeChild(old);
  if (edit.val !== "top") {
    const after = ["noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"].map((n) => composeDirectChild(sect, n)).find(Boolean);
    sect.insertBefore(composeEl(doc, "vAlign", { val: edit.val }), after || null);
  }
  return { xml: composeSerialize(doc), count: 1 };
}

// ── op "pageSetup": marginesy, orientacja, rozmiar papieru (Word: Układ → Marginesy / Orientacja / Rozmiar)
// Sekcja akapitu z kursorem (scope „section”, jak Word dla bieżącej sekcji) albo wszystkie
// (scope „all” — „Zastosuj do: cały dokument”). Wartości w twipach (1 cm = 567).
// margins: { top, bottom, left, right }; orient: "portrait" | "landscape" (zmiana obraca kartkę
// i marginesy: górny↔lewy, dolny↔prawy — jak Word); size: { w, h } w pionie.
const COMPOSE_SECTPR_AFTER_PGMAR = ["paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"];
function composeAllSections(doc) {
  const out = [];
  collectParagraphElements(doc.documentElement, "all").forEach((p) => {
    const pPr = composeDirectChild(p, "pPr");
    const s = pPr && composeDirectChild(pPr, "sectPr");
    if (s) out.push(s);
  });
  const body = composeDirectChild(doc.getElementsByTagNameNS(W_NS, "body")[0], "sectPr");
  if (body) out.push(body);
  return out;
}
function composeSectionAt(doc, index) {
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  for (let i = Math.max(0, index | 0); i < paragraphs.length; i++) {
    const pPr = composeDirectChild(paragraphs[i], "pPr");
    const s = pPr && composeDirectChild(pPr, "sectPr");
    if (s) return s;
  }
  const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
  let s = composeDirectChild(body, "sectPr");
  if (!s) { s = doc.createElementNS(W_NS, "w:sectPr"); body.appendChild(s); }
  return s;
}
// Ustawienia strony sekcji (do zaznaczenia w menu); brak wpisów = A4 pionowo, marginesy 2,5 cm.
function composeSectionPageSetup(doc, index) {
  const sect = composeSectionAt(doc, index);
  const pgSz = composeDirectChild(sect, "pgSz");
  const mar = composeDirectChild(sect, "pgMar");
  const num = (el, a, d) => { const v = parseInt(el?.getAttributeNS(W_NS, a) || "", 10); return Number.isFinite(v) ? v : d; };
  const w = num(pgSz, "w", 11906), h = num(pgSz, "h", 16838);
  return {
    w, h, orient: pgSz?.getAttributeNS(W_NS, "orient") === "landscape" || w > h ? "landscape" : "portrait",
    top: num(mar, "top", 1418), bottom: num(mar, "bottom", 1418), left: num(mar, "left", 1418), right: num(mar, "right", 1418),
    sections: composeAllSections(doc).length || 1,
  };
}
function composeEnsurePageEls(doc, sect) {
  let pgSz = composeDirectChild(sect, "pgSz");
  let mar = composeDirectChild(sect, "pgMar");
  const after = () => COMPOSE_SECTPR_AFTER_PGMAR.map((n) => composeDirectChild(sect, n)).find(Boolean) || null;
  if (!mar) {
    mar = composeEl(doc, "pgMar", { top: "1418", right: "1418", bottom: "1418", left: "1418", header: "709", footer: "709", gutter: "0" });
    sect.insertBefore(mar, after());
  }
  if (!pgSz) {
    pgSz = composeEl(doc, "pgSz", { w: "11906", h: "16838" });
    sect.insertBefore(pgSz, mar);
  }
  return { pgSz, mar };
}
function applyPageSetupInXml(xml, edit) {
  const doc = composeParse(xml);
  const targets = edit.scope === "all" ? composeAllSections(doc) : [composeSectionAt(doc, edit.index)];
  if (!targets.length) targets.push(composeSectionAt(doc, 0));
  const get = (el, a) => parseInt(el.getAttributeNS(W_NS, a) || "0", 10) || 0;
  const set = (el, a, v) => el.setAttributeNS(W_NS, `w:${a}`, String(Math.max(0, Math.round(v))));
  let count = 0;
  for (const sect of targets) {
    const { pgSz, mar } = composeEnsurePageEls(doc, sect);
    let w = get(pgSz, "w") || 11906;
    let h = get(pgSz, "h") || 16838;
    const landscape = pgSz.getAttributeNS(W_NS, "orient") === "landscape" || w > h;
    if (edit.size) {
      const a = Math.min(edit.size.w, edit.size.h), b = Math.max(edit.size.w, edit.size.h);
      [w, h] = landscape ? [b, a] : [a, b];
    }
    if (edit.orient && (edit.orient === "landscape") !== landscape) {
      [w, h] = [h, w];
      const m = { top: get(mar, "top"), bottom: get(mar, "bottom"), left: get(mar, "left"), right: get(mar, "right") };
      set(mar, "top", m.left); set(mar, "left", m.top); set(mar, "bottom", m.right); set(mar, "right", m.bottom);
    }
    set(pgSz, "w", w);
    set(pgSz, "h", h);
    if (w > h) pgSz.setAttributeNS(W_NS, "w:orient", "landscape"); else pgSz.removeAttributeNS(W_NS, "orient");
    if (edit.margins) ["top", "bottom", "left", "right"].forEach((k) => { if (Number.isFinite(edit.margins[k])) set(mar, k, edit.margins[k]); });
    count++;
  }
  return { xml: composeSerialize(doc), count };
}

// Wyrównanie strony w pionie w sekcji akapitu (do zaznaczenia w menu).
function composeSectionVAlign(doc, index) {
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  let sect = null;
  for (let i = Math.max(0, index | 0); i < paragraphs.length && !sect; i++) {
    const pPr = composeDirectChild(paragraphs[i], "pPr");
    sect = pPr && composeDirectChild(pPr, "sectPr");
  }
  if (!sect) sect = composeDirectChild(doc.getElementsByTagNameNS(W_NS, "body")[0], "sectPr");
  return composeDirectChild(sect, "vAlign")?.getAttributeNS(W_NS, "val") || "top";
}

// ── op "pasteBlocks": wklejka ze strukturą (paste-rich.js) ─────────────────────
// Plan wklejki — ten sam w zapisie i przy ustawianiu kursora (paste-rich.js): pozycje kolejnych
// elementów względem akapitu z kursorem. Zwykły akapit wklejki dokleja się do tekstu przed
// kursorem; nagłówek/punkt listy/cytat na początku pustego miejsca zamienia ten akapit; tekst za
// kursorem dołącza do ostatniego akapitu wklejki (po tabeli — nowy akapit).
const COMPOSE_PASTE_TEXT = ["p", "h", "li", "quote", "code"];
function composePastePlan(blocks, beforeEmpty) {
  const items = []; // { block | null (sam tekst przed kursorem), merge: true = z tekstem przed kursorem }
  blocks.forEach((b, i) => {
    if (i === 0) {
      if (b.type === "p") { items.push({ block: b, merge: true }); return; }
      if (!beforeEmpty) items.push({ block: null, merge: true });
    }
    items.push({ block: b, merge: false });
  });
  const lastText = [...items].reverse().find((it) => !it.block || COMPOSE_PASTE_TEXT.includes(it.block.type) || it.block.type === "hr");
  const last = items[items.length - 1];
  const tailItem = last.block && (last.block.type === "table" || last.block.type === "hr");
  if (tailItem) items.push({ block: null, tail: true });
  const paras = (it) => (it.block?.type === "table" ? it.block.rows.reduce((n, r) => n + r.length, 0) : 1);
  let off = 0;
  let lastParaOffset = 0;
  let lastTextLen = 0;
  items.forEach((it) => {
    if (it.block?.type === "table") { off += paras(it); return; }
    lastParaOffset = off;
    lastTextLen = it.block ? previewRunsToPlainText(it.block.runs || []).length + (it.block.type === "li" && it.block.checked !== undefined ? 2 : 0) : 0;
    off += 1;
  });
  return { items, lastParaOffset, lastTextLen, lastIsFirst: items.length === 1 && items[0].merge, lastText };
}

function composePasteListLevels(group) {
  // rodzaj każdego poziomu z pierwszego punktu na tym poziomie (numerowana / punktowana)
  const kinds = [];
  group.forEach((b) => { if (kinds[b.level] == null) kinds[b.level] = b.ordered ? "n" : "b"; });
  for (let i = 0; i < 9; i++) if (kinds[i] == null) kinds[i] = kinds[i - 1] || "b";
  return kinds;
}

async function applyPasteBlocksInZip(zip, xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[edit.index];
  const blocks = (edit.blocks || []).filter((b) => b && (b.type === "table" ? b.rows?.length : true));
  if (!p || !blocks.length) return { xml, count: 0 };
  const cache = new Map();
  const runs = extractRunsFromParagraphXml(p);
  const total = previewRunsToPlainText(runs).length;
  const at = Math.max(0, Math.min(total, edit.offset | 0));
  const { before } = composeSliceRuns(runs, at, at);
  const { after } = composeSliceRuns(runs, at, at);
  const plan = composePastePlan(blocks, !previewRunsToPlainText(before).length);

  // wzór zwykłego akapitu: akapit z kursorem bez stylu nagłówka, listy, sekcji, podziału strony
  const origPPr = composeDirectChild(p, "pPr");
  // zwykły tekst w kilku wierszach (keepPara): format akapitu z kursorem zostaje (lista, styl), jak w Wordzie
  const plainPPr = () => {
    const np = origPPr ? origPPr.cloneNode(true) : null;
    if (np) (edit.keepPara ? ["sectPr", "pageBreakBefore"] : ["pStyle", "numPr", "sectPr", "pageBreakBefore", "outlineLvl", "pBdr", "keepNext"]).forEach((k) => composeSetPPrChild(np, k, null));
    return np;
  };
  // listy: kolejne punkty (bez list kontrolnych) = jedna lista z numeracją od 1
  let numDoc = null;
  let numIdx = null;
  const numPath = "word/numbering.xml";
  const listNum = new Map(); // blok → numId
  const groups = [];
  let g = null;
  blocks.forEach((b) => {
    if (b.type !== "li" || b.checked !== undefined) { g = null; return; }
    // punktowana zaraz po numerowanej (albo odwrotnie) na pierwszym poziomie = nowa lista
    const top = g && g.find((x) => !x.level);
    if (g && !b.level && top && !!top.ordered !== !!b.ordered) g = null;
    if (!g) groups.push((g = []));
    g.push(b);
  });
  if (groups.length) {
    const numXml = zip.file(numPath) ? await zip.file(numPath).async("string") : null;
    numDoc = composeParse(numXml || `<w:numbering xmlns:w="${W_NS}"/>`);
    numIdx = composeNumberingIndex(numDoc);
    const root = numDoc.documentElement;
    for (const group of groups) {
      const kinds = composePasteListLevels(group);
      const name = `DWB Paste ${kinds.join("")}`;
      let absId = Array.from(numIdx.abstracts.entries()).find(([, a]) => a.name === name)?.[0];
      if (absId == null) {
        absId = String(Array.from(numIdx.abstracts.keys()).reduce((m, k) => Math.max(m, parseInt(k, 10) || 0), -1) + 1);
        const nsid = Math.floor(Math.random() * 0xffffffff).toString(16).toUpperCase().padStart(8, "0");
        let lvls = "";
        kinds.forEach((k, i) => {
          const ind = `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr>`;
          if (k === "b") lvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${COMPOSE_BULLETS[i % 3]}"/><w:lvlJc w:val="left"/>${ind}<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/></w:rPr></w:lvl>`;
          else { const [fmt, text] = COMPOSE_NUMFMT[i % 3]; lvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text.replace("L", i + 1)}"/><w:lvlJc w:val="left"/>${ind}</w:lvl>`; }
        });
        const frag = composeParse(`<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="${absId}"><w:nsid w:val="${nsid}"/><w:multiLevelType w:val="hybridMultilevel"/><w:name w:val="${name}"/>${lvls}</w:abstractNum></w:numbering>`);
        composeNumberingInsert(root, numDoc.importNode(frag.documentElement.firstChild, true));
        numIdx.abstracts.set(absId, { el: null, name });
      }
      const numId = String(Array.from(numIdx.nums.keys()).reduce((m, k) => Math.max(m, parseInt(k, 10) || 0), 0) + 1);
      const num = composeEl(numDoc, "num", { numId });
      num.appendChild(composeEl(numDoc, "abstractNumId", { val: absId }));
      const ov = composeEl(numDoc, "lvlOverride", { ilvl: 0 });
      ov.appendChild(composeEl(numDoc, "startOverride", { val: 1 }));
      num.appendChild(ov);
      composeNumberingInsert(root, num);
      numIdx.nums.set(numId, absId);
      group.forEach((b) => listNum.set(b, numId));
    }
  }

  const styleOf = async (b) => {
    if (b.type === "h") return composeEnsureStyle(zip, `h${Math.max(1, Math.min(3, b.level || 1))}`, cache);
    if (b.type === "quote") return composeEnsureStyle(zip, "quote", cache);
    return null;
  };
  const tableStyle = blocks.some((b) => b.type === "table") ? await composeEnsureTableStyle(zip) : null;

  const nodes = []; // elementy w kolejności
  let reused = false;
  for (const it of plan.items) {
    const b = it.block;
    if (b?.type === "table") {
      const cols = Math.max(...b.rows.map((r) => r.length));
      const totalW = composeTextWidthTwips(doc, p);
      const colW = Math.floor(totalW / cols);
      const tbl = composeEl(doc, "tbl");
      tbl.appendChild(composeParse(`<w:tblPr xmlns:w="${W_NS}"><w:tblStyle w:val="${tableStyle}"/><w:tblW w:w="${colW * cols}" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>`).documentElement);
      const grid = composeEl(doc, "tblGrid");
      for (let c = 0; c < cols; c++) grid.appendChild(composeEl(doc, "gridCol", { w: colW }));
      tbl.appendChild(grid);
      b.rows.forEach((row) => {
        const tr = composeEl(doc, "tr");
        for (let c = 0; c < cols; c++) {
          const tc = composeEl(doc, "tc");
          const tcPr = composeEl(doc, "tcPr");
          tcPr.appendChild(composeEl(doc, "tcW", { w: colW, type: "dxa" }));
          tc.appendChild(tcPr);
          const cp = composeEl(doc, "p");
          applyRunsToParagraphXml(cp, row[c] || []);
          tc.appendChild(cp);
          tr.appendChild(tc);
        }
        tbl.appendChild(tr);
      });
      nodes.push(tbl);
      continue;
    }
    // akapit: pierwszy z tekstem przed kursorem (albo zamieniony) = ten sam element w:p
    const reuse = !reused && (it.merge || plan.items.indexOf(it) === 0 || (plan.items[0].block?.type === "table" && !reused));
    const el = reuse ? p : composeEl(doc, "p");
    if (reuse) reused = true;
    if (!reuse || (b && !it.merge)) {
      // nowy akapit albo akapit zamieniony w nagłówek / punkt listy / cytat
      const old = composeDirectChild(el, "pPr");
      const sect = old && composeDirectChild(old, "sectPr");
      if (old) el.removeChild(old);
      const pPr = plainPPr();
      if (pPr) el.insertBefore(pPr, el.firstChild);
      if (sect) composeSetPPrChild(composeEnsurePPr(el), "sectPr", sect);
    }
    let content = [];
    if (it.merge) content = [...before];
    if (b) {
      const pPr = composeEnsurePPr(el);
      const sid = await styleOf(b);
      if (sid) composeSetPPrChild(pPr, "pStyle", { val: sid });
      if (b.type === "li" && b.checked === undefined) {
        const numPr = composeEl(doc, "numPr");
        numPr.appendChild(composeEl(doc, "ilvl", { val: Math.max(0, Math.min(8, b.level | 0)) }));
        numPr.appendChild(composeEl(doc, "numId", { val: listNum.get(b) }));
        composeSetPPrChild(pPr, "numPr", numPr);
      }
      if (b.type === "li" && b.checked !== undefined) {
        // lista kontrolna: prawdziwe pole wyboru Worda (klikane w aplikacji i w Wordzie)
        if (b.level) composeSetPPrChild(pPr, "ind", { left: 360 * b.level });
        const sdtEl = composeParse(composeSdtXml({ kind: "checkbox", checked: !!b.checked, lang: edit.lang }, "PlaceholderText")).documentElement;
        content.push({ island: new XMLSerializer().serializeToString(sdtEl), text: b.checked ? "☒" : "☐" }, { text: " " });
      }
      if (b.type === "hr") {
        const pBdr = composeEl(doc, "pBdr");
        pBdr.appendChild(composeEl(doc, "bottom", { val: "single", sz: 8, space: 1, color: "auto" }));
        composeSetPPrChild(pPr, "pBdr", pBdr);
      }
      if (!pPr.firstChild) el.removeChild(pPr);
      content.push(...(b.runs || []));
    }
    nodes.push({ el, content });
  }
  // tekst za kursorem → ostatni akapit
  const lastPara = [...nodes].reverse().find((n) => n.el);
  lastPara.content.push(...after);
  // na miejsce akapitu z kursorem, w kolejności
  const parent = p.parentNode;
  const anchor = p.nextSibling;
  parent.removeChild(p);
  nodes.forEach((n) => {
    const node = n.el || n;
    parent.insertBefore(node, anchor);
    if (n.el) applyRunsToParagraphXml(n.el, n.content);
    // kopia pPr akapitu z kursorem w nowym akapicie — śledzone zmiany z nowymi numerami (już w dokumencie)
    if (n.el && n.el !== p) { const pr = composeDirectChild(n.el, "pPr"); if (pr) freshRevisionIds(pr, doc); }
  });
  // znacznik końca sekcji należy do ostatniego akapitu
  if (lastPara.el !== p) composeMoveSectPr(p, lastPara.el);
  if (numDoc) {
    await composeEnsurePart(zip, numPath, composeSerialize(numDoc),
      "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
    zip.file(numPath, composeSerialize(numDoc));
  }
  if (cache.stylesDirty) zip.file("word/styles.xml", cache.stylesXml);
  return { xml: composeSerialize(doc), count: blocks.length };
}

// ── tabele ───────────────────────────────────────────────────────────────────
// Szerokość tekstu (twipy) z sekcji akapitu (albo ostatniej w dokumencie): strona − marginesy.
function composeTextWidthTwips(doc, near) {
  let sect = null;
  if (near) {
    // sekcja akapitu = pierwszy w:sectPr w akapicie od tego miejsca w dół, inaczej sekcja dokumentu
    const paras = collectParagraphElements(doc.documentElement, "all");
    for (let i = Math.max(0, paras.indexOf(near)); i < paras.length && !sect; i++) {
      const pPr = composeDirectChild(paras[i], "pPr");
      sect = pPr && composeDirectChild(pPr, "sectPr");
    }
  }
  if (!sect) sect = composeDirectChild(doc.getElementsByTagNameNS(W_NS, "body")[0], "sectPr");
  const num = (el, a, d) => parseInt(el?.getAttributeNS(W_NS, a), 10) || d;
  const pgSz = sect && composeDirectChild(sect, "pgSz");
  const mar = sect && composeDirectChild(sect, "pgMar");
  const w = num(pgSz, "w", 11906);
  return Math.max(1440, w - num(mar, "left", 1418) - num(mar, "right", 1418));
}

function composeCellXml(widthTw) {
  return `<w:tc><w:tcPr><w:tcW w:w="${widthTw}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`;
}

// op "tableInsert": { index, rows, cols } — tabela pod akapitem z kursorem (pusty akapit zostaje
// pod tabelą, żeby dało się pisać dalej; tabela nie może być ostatnim elementem treści).
async function applyTableInsertInZip(zip, xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[edit.index];
  if (!p) return { xml, count: 0 };
  const rows = Math.max(1, Math.min(50, edit.rows | 0));
  const cols = Math.max(1, Math.min(20, edit.cols | 0));
  const styleId = await composeEnsureTableStyle(zip);
  const total = composeTextWidthTwips(doc, p);
  const colW = Math.floor(total / cols);
  const grid = Array.from({ length: cols }, () => `<w:gridCol w:w="${colW}"/>`).join("");
  const row = `<w:tr>${Array.from({ length: cols }, () => composeCellXml(colW)).join("")}</w:tr>`;
  const tblXml = `<w:tbl xmlns:w="${W_NS}"><w:tblPr><w:tblStyle w:val="${styleId}"/><w:tblW w:w="${colW * cols}" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${row.repeat(rows)}</w:tbl>`;
  const tbl = doc.importNode(composeParse(tblXml).documentElement, true);
  const empty = !getParagraphText(p).trim() && !p.getElementsByTagNameNS(W_NS, "drawing").length;
  if (empty) {
    p.parentNode.insertBefore(tbl, p); // pusty akapit zostaje pod tabelą
  } else {
    p.parentNode.insertBefore(tbl, p.nextSibling);
    const next = composeEl(doc, "p");
    tbl.parentNode.insertBefore(next, tbl.nextSibling);
    composeMoveSectPr(p, next);
  }
  return { xml: composeSerialize(doc), count: 1 };
}

async function composeEnsureTableStyle(zip) {
  const xml = await composeEnsureStylesPart(zip);
  const doc = composeParse(xml);
  const found = Array.from(doc.getElementsByTagNameNS(W_NS, "style")).find((st) => st.getAttributeNS(W_NS, "type") === "table"
    && (composeDirectChild(st, "name")?.getAttributeNS(W_NS, "val") || "").trim().toLowerCase() === "table grid");
  if (found) { zip.file("word/styles.xml", xml); return found.getAttributeNS(W_NS, "styleId"); }
  const b = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
  const frag = composeParse(`<w:styles xmlns:w="${W_NS}"><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map(b).join("")}</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`);
  doc.documentElement.appendChild(doc.importNode(frag.documentElement.firstChild, true));
  zip.file("word/styles.xml", composeSerialize(doc));
  return "TableGrid";
}

// Komórka z akapitem: { tbl, tr, tc, rowIndex, gridCol } (kolumna siatki z uwzględnieniem gridSpan).
function composeCellOf(p) {
  let tc = p.parentNode;
  while (tc && !(tc.localName === "tc" && tc.namespaceURI === W_NS)) tc = tc.parentNode;
  if (!tc) return null;
  const tr = tc.parentNode;
  const tbl = tr?.parentNode;
  if (!tr || tr.localName !== "tr" || !tbl || tbl.localName !== "tbl") return null;
  return { tbl, tr, tc, rowIndex: composeRows(tbl).indexOf(tr), gridCol: composeGridStart(tc) };
}
const composeRows = (tbl) => Array.from(tbl.childNodes).filter((n) => n.localName === "tr" && n.namespaceURI === W_NS);
const composeCells = (tr) => Array.from(tr.childNodes).filter((n) => n.localName === "tc" && n.namespaceURI === W_NS);
function composeSpan(tc) {
  const tcPr = composeDirectChild(tc, "tcPr");
  return parseInt(tcPr && composeDirectChild(tcPr, "gridSpan")?.getAttributeNS(W_NS, "val"), 10) || 1;
}
function composeGridStart(tc) {
  let col = 0;
  for (const c of composeCells(tc.parentNode)) { if (c === tc) return col; col += composeSpan(c); }
  return col;
}
// Komórka wiersza zajmująca kolumnę siatki col (albo null).
function composeCellAtCol(tr, col) {
  let at = 0;
  for (const c of composeCells(tr)) { const s = composeSpan(c); if (col >= at && col < at + s) return c; at += s; }
  return null;
}
// Pusta kopia komórki: te same właściwości (szerokość, tło), bez scalenia w pionie, jeden pusty akapit.
function composeBlankCell(doc, tc) {
  const nc = composeEl(doc, "tc");
  const tcPr = composeDirectChild(tc, "tcPr");
  if (tcPr) {
    const pr = tcPr.cloneNode(true);
    Array.from(pr.childNodes).filter((n) => n.localName === "vMerge").forEach((n) => pr.removeChild(n));
    nc.appendChild(pr);
  }
  const firstP = Array.from(tc.childNodes).find((n) => n.localName === "p");
  const np = composeEl(doc, "p");
  const pPr = firstP && composeDirectChild(firstP, "pPr");
  if (pPr) {
    const c = pPr.cloneNode(true);
    Array.from(c.childNodes).filter((n) => ["sectPr", "numPr"].includes(n.localName)).forEach((n) => c.removeChild(n));
    np.appendChild(c);
    freshRevisionIds(c, doc);
  }
  nc.appendChild(np);
  return nc;
}
// Scalenie w pionie: "restart" (początek), "continue" (ciąg dalszy) albo null.
function composeVMerge(tc) {
  const tcPr = tc && composeDirectChild(tc, "tcPr");
  const v = tcPr && composeDirectChild(tcPr, "vMerge");
  if (!v) return null;
  return v.getAttributeNS(W_NS, "val") === "restart" ? "restart" : "continue";
}
function composeSetVMerge(doc, tc, kind) {
  let tcPr = composeDirectChild(tc, "tcPr");
  if (!tcPr) { tcPr = composeEl(doc, "tcPr"); tc.insertBefore(tcPr, tc.firstChild); }
  const old = composeDirectChild(tcPr, "vMerge");
  if (old) tcPr.removeChild(old);
  if (!kind) return;
  const v = composeEl(doc, "vMerge", kind === "restart" ? { val: "restart" } : {});
  // kolejność w tcPr: tcW, gridSpan, (hMerge), vMerge, tcBorders… — za gridSpan/tcW
  const after = composeDirectChild(tcPr, "gridSpan") || composeDirectChild(tcPr, "tcW");
  tcPr.insertBefore(v, after ? after.nextSibling : tcPr.firstChild);
}

function composeSetCellWidth(tc, w) {
  const tcPr = composeDirectChild(tc, "tcPr");
  const tcW = tcPr && composeDirectChild(tcPr, "tcW");
  if (tcW && tcW.getAttributeNS(W_NS, "type") !== "pct") { tcW.setAttributeNS(W_NS, "w:w", String(Math.round(w))); tcW.setAttributeNS(W_NS, "w:type", "dxa"); }
}

// op "table": { index, action } — rowAbove / rowBelow / colLeft / colRight / delRow / delCol / delTable
function applyTableInXml(xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[edit.index];
  const cell = p && composeCellOf(p);
  if (!cell) return { xml, count: 0 };
  const { tbl, tr, tc, gridCol } = cell;
  const rows = composeRows(tbl);
  const gridEl = composeDirectChild(tbl, "tblGrid");
  const gridCols = gridEl ? Array.from(gridEl.childNodes).filter((n) => n.localName === "gridCol") : [];
  const a = edit.action;
  if (a === "rowAbove" || a === "rowBelow") {
    const nr = composeEl(doc, "tr");
    const trPr = composeDirectChild(tr, "trPr");
    if (trPr) nr.appendChild(trPr.cloneNode(true));
    // wiersz wstawiony w środek scalenia pionowego je przedłuża (jak w Wordzie)
    const below = a === "rowAbove" ? tr : rows[rows.indexOf(tr) + 1];
    composeCells(tr).forEach((c) => {
      const nc = composeBlankCell(doc, c);
      const under = below && composeCellAtCol(below, composeGridStart(c));
      if (composeVMerge(under) === "continue") composeSetVMerge(doc, nc, "continue");
      nr.appendChild(nc);
    });
    tbl.insertBefore(nr, a === "rowAbove" ? tr : tr.nextSibling);
  } else if (a === "colLeft" || a === "colRight") {
    const at = a === "colLeft" ? gridCol : gridCol + composeSpan(tc); // nowa kolumna siatki na tej pozycji
    const ref = gridCols[Math.min(gridCols.length - 1, a === "colLeft" ? gridCol : gridCol + composeSpan(tc) - 1)];
    const total = gridCols.reduce((s, g) => s + (parseInt(g.getAttributeNS(W_NS, "w"), 10) || 0), 0);
    const newW = parseInt(ref?.getAttributeNS(W_NS, "w"), 10) || Math.round(total / Math.max(1, gridCols.length));
    if (gridEl) {
      const g = composeEl(doc, "gridCol", { w: newW });
      gridEl.insertBefore(g, gridCols[at] || null);
      // ta sama szerokość tabeli: wszystkie kolumny proporcjonalnie węższe
      const all = Array.from(gridEl.childNodes).filter((n) => n.localName === "gridCol");
      const sum = all.reduce((s, x) => s + (parseInt(x.getAttributeNS(W_NS, "w"), 10) || 0), 0);
      const k = total > 0 && sum > 0 ? total / sum : 1;
      all.forEach((x) => x.setAttributeNS(W_NS, "w:w", String(Math.max(200, Math.round((parseInt(x.getAttributeNS(W_NS, "w"), 10) || 0) * k)))));
    }
    rows.forEach((r) => {
      // komórka scalona w poziomie PRZEZ miejsce nowej kolumny — tylko szersza (jak w Wordzie)
      const across = composeCells(r).find((c) => { const st = composeGridStart(c); return st < at && at < st + composeSpan(c); });
      if (across) {
        let tcPr = composeDirectChild(across, "tcPr");
        let span = tcPr && composeDirectChild(tcPr, "gridSpan");
        span.setAttributeNS(W_NS, "w:val", String(composeSpan(across) + 1));
        return;
      }
      const near = composeCellAtCol(r, a === "colLeft" ? at : at - 1) || composeCells(r).slice(-1)[0];
      if (!near) return;
      const nc = composeBlankCell(doc, near);
      const tcPr = composeDirectChild(nc, "tcPr");
      const span = tcPr && composeDirectChild(tcPr, "gridSpan");
      if (span) tcPr.removeChild(span);
      if (gridEl) composeSetCellWidth(nc, parseInt(Array.from(gridEl.childNodes).filter((n) => n.localName === "gridCol")[at]?.getAttributeNS(W_NS, "w"), 10) || newW);
      const start = composeGridStart(near);
      // wstaw przed komórką zaczynającą się na „at”, inaczej za komórką obejmującą at−1
      const before = composeCells(r).find((c) => composeGridStart(c) >= at);
      if (a === "colLeft" && start === at) r.insertBefore(nc, near);
      else if (before) r.insertBefore(nc, before);
      else r.appendChild(nc);
    });
    // szerokości komórek z nowej siatki (komórka scalona = suma swoich kolumn)
    if (gridEl) {
      const all = Array.from(gridEl.childNodes).filter((n) => n.localName === "gridCol").map((g) => parseInt(g.getAttributeNS(W_NS, "w"), 10) || 0);
      rows.forEach((r) => composeCells(r).forEach((c) => {
        const st = composeGridStart(c);
        const w = all.slice(st, st + composeSpan(c)).reduce((x, y) => x + y, 0);
        if (w) composeSetCellWidth(c, w);
      }));
    }
  } else if (a === "delRow") {
    if (rows.length <= 1) return applyTableInXml(xml, { ...edit, action: "delTable" });
    // usuwany wiersz zaczynał scalenie pionowe — jego ciąg dalszy staje się początkiem
    const next = rows[rows.indexOf(tr) + 1];
    if (next) composeCells(tr).forEach((c) => {
      if (composeVMerge(c) !== "restart") return;
      const under = composeCellAtCol(next, composeGridStart(c));
      if (composeVMerge(under) === "continue") composeSetVMerge(doc, under, "restart");
    });
    tbl.removeChild(tr);
  } else if (a === "delCol") {
    if (gridCols.length <= 1 || rows.every((r) => composeCells(r).length <= 1)) return applyTableInXml(xml, { ...edit, action: "delTable" });
    rows.forEach((r) => {
      const c = composeCellAtCol(r, gridCol);
      if (!c) return;
      const s = composeSpan(c);
      if (s > 1) { // komórka scalona w poziomie — tylko węższa
        const tcPr = composeDirectChild(c, "tcPr");
        composeDirectChild(tcPr, "gridSpan").setAttributeNS(W_NS, "w:val", String(s - 1));
      } else if (composeCells(r).length > 1) r.removeChild(c);
    });
    if (gridCols[gridCol]) gridEl.removeChild(gridCols[gridCol]);
  } else if (a === "align") {
    // Wyrównanie w komórce jak w Wordzie (Układ tabeli → Wyrównanie): poziomo = akapity komórki
    // (w:jc), pionowo = komórka (w:vAlign). Zakres: komórka / wiersz / kolumna / tabela.
    const col = composeGridStart(tc);
    const targets = edit.scope === "row" ? composeCells(tr)
      : edit.scope === "col" ? rows.map((r) => composeCellAtCol(r, col)).filter((c) => c && composeGridStart(c) <= col && col < composeGridStart(c) + composeSpan(c))
      : edit.scope === "table" ? rows.flatMap((r) => composeCells(r))
      : [tc];
    const TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"];
    targets.forEach((cell) => {
      if (["top", "center", "bottom"].includes(edit.v)) {
        let tcPr = composeDirectChild(cell, "tcPr");
        if (!tcPr) { tcPr = composeEl(doc, "tcPr"); cell.insertBefore(tcPr, cell.firstChild); }
        const old = composeDirectChild(tcPr, "vAlign");
        if (old) tcPr.removeChild(old);
        if (edit.v !== "top") { // „góra” = domyślne w Wordzie — bez wpisu
          const el = composeEl(doc, "vAlign", { val: edit.v });
          const after = TCPR_ORDER.slice(TCPR_ORDER.indexOf("vAlign") + 1).map((n) => composeDirectChild(tcPr, n)).find(Boolean);
          tcPr.insertBefore(el, after || null);
        }
      }
      if (["left", "center", "right", "both"].includes(edit.h)) {
        Array.from(cell.childNodes).filter((n) => n.localName === "p" && n.namespaceURI === W_NS).forEach((p) => {
          composeSetPPrChild(composeEnsurePPr(p), "jc", { val: edit.h });
        });
      }
    });
  } else if (a === "delTable") {
    const parent = tbl.parentNode;
    // tabela w komórce — komórka musi kończyć się akapitem
    if (parent.localName === "tc" && !Array.from(parent.childNodes).some((n) => n !== tbl && n.localName === "p")) parent.insertBefore(composeEl(doc, "p"), tbl.nextSibling);
    parent.removeChild(tbl);
  } else {
    return { xml, count: 0 };
  }
  return { xml: composeSerialize(doc), count: 1 };
}

// ── obrazy ───────────────────────────────────────────────────────────────────
const COMPOSE_NS = {
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
};
const COMPOSE_EMU_PER_TWIP = 635;
const COMPOSE_EMU_PER_PX = 9525; // 96 dpi

async function composeAddMedia(zip, bytes, ext, mime) {
  let n = 1;
  while (zip.file(`word/media/image${n}.${ext}`)) n++;
  const name = `image${n}.${ext}`;
  zip.file(`word/media/${name}`, bytes);
  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    let ct = await ctFile.async("string");
    if (!new RegExp(`Extension="${ext}"`, "i").test(ct)) {
      ct = ct.replace("<Default ", `<Default Extension="${ext}" ContentType="${mime}"/><Default `);
      zip.file("[Content_Types].xml", ct);
    }
  }
  const rels = await composeReadRels(zip);
  const rid = composeAddRel(rels.doc, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", `media/${name}`, false);
  zip.file(rels.path, composeSerialize(rels.doc));
  return { rid, name };
}

function composeDrawingXml({ rid, cx, cy, id, name, descr }) {
  const esc = (v) => composeXmlText(v || "").replace(/"/g, "&quot;");
  return `<w:r xmlns:w="${W_NS}"><w:drawing><wp:inline xmlns:wp="${COMPOSE_NS.wp}" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="${esc(name)}" descr="${esc(descr)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${COMPOSE_NS.a}" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="${COMPOSE_NS.a}"><a:graphicData uri="${COMPOSE_NS.pic}"><pic:pic xmlns:pic="${COMPOSE_NS.pic}"><pic:nvPicPr><pic:cNvPr id="0" name="${esc(name)}" descr="${esc(descr)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="${COMPOSE_NS.r}" r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

// op "imageInsert": { index, bytes (Uint8Array), ext, mime, width, height (px), descr }
// Obraz w osobnym akapicie (wyśrodkowany) pod akapitem z kursorem — albo w nim, gdy jest pusty.
// Szerszy niż tekst — zmniejszony do szerokości tekstu (proporcje zostają).
async function applyImageInsertInZip(zip, xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[edit.index];
  if (!p || !edit.bytes?.length) return { xml, count: 0 };
  const { rid, name } = await composeAddMedia(zip, edit.bytes, edit.ext, edit.mime);
  const maxCx = composeTextWidthTwips(doc, p) * COMPOSE_EMU_PER_TWIP;
  let cx = Math.round((edit.width || 600) * COMPOSE_EMU_PER_PX);
  let cy = Math.round((edit.height || 400) * COMPOSE_EMU_PER_PX);
  if (cx > maxCx) { cy = Math.round(cy * maxCx / cx); cx = maxCx; }
  const ids = Array.from(doc.getElementsByTagNameNS(COMPOSE_NS.wp, "docPr")).map((d) => parseInt(d.getAttribute("id"), 10) || 0);
  const id = Math.max(0, ...ids) + 1;
  const runEl = doc.importNode(composeParse(composeDrawingXml({ rid, cx, cy, id, name: edit.name || name, descr: edit.descr || "" })).documentElement, true);
  const empty = !getParagraphText(p).trim() && !p.getElementsByTagNameNS(W_NS, "drawing").length && !p.getElementsByTagNameNS(W_NS, "sdt").length;
  let target = p;
  if (!empty) {
    target = composeNewParagraphAfter(p);
    composeMoveSectPr(p, target);
  }
  const pPr = composeEnsurePPr(target);
  composeSetPPrChild(pPr, "jc", { val: "center" });
  target.appendChild(runEl);
  // pod obrazem musi być gdzie pisać dalej
  const after = target.nextSibling && target.nextSibling.nodeType === 1 ? target.nextSibling : null;
  if (!after || after.localName !== "p") {
    const np = composeNewParagraphAfter(target);
    composeMoveSectPr(target, np);
  }
  return { xml: composeSerialize(doc), count: 1 };
}

// op "image": { index, action: "size" (widthPct) | "align" (align) | "alt" (descr) | "delete" }
// Działa na pierwszym obrazie akapitu — także z plików z Worda (wp:inline albo wp:anchor).
function applyImageInXml(xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[edit.index];
  const drawing = p?.getElementsByTagNameNS(W_NS, "drawing")[0];
  const box = drawing && Array.from(drawing.childNodes).find((n) => n.nodeType === 1 && (n.localName === "inline" || n.localName === "anchor"));
  if (!box) return { xml, count: 0 };
  const extent = Array.from(box.childNodes).find((n) => n.localName === "extent");
  if (edit.action === "size") {
    const cx0 = parseInt(extent?.getAttribute("cx"), 10);
    const cy0 = parseInt(extent?.getAttribute("cy"), 10);
    if (!cx0 || !cy0) return { xml, count: 0 };
    const pct = Math.max(5, Math.min(100, Number(edit.widthPct) || 100));
    const cx = Math.round(composeTextWidthTwips(doc, p) * COMPOSE_EMU_PER_TWIP * pct / 100);
    const cy = Math.round(cy0 * cx / cx0);
    extent.setAttribute("cx", String(cx));
    extent.setAttribute("cy", String(cy));
    Array.from(box.getElementsByTagNameNS(COMPOSE_NS.a, "ext")).forEach((e) => {
      if (e.parentNode?.localName === "xfrm" && e.getAttribute("cx")) { e.setAttribute("cx", String(cx)); e.setAttribute("cy", String(cy)); }
    });
  } else if (edit.action === "align") {
    if (!COMPOSE_ALIGNS.includes(edit.align)) return { xml, count: 0 };
    composeSetPPrChild(composeEnsurePPr(p), "jc", { val: edit.align });
  } else if (edit.action === "alt") {
    Array.from(box.getElementsByTagNameNS(COMPOSE_NS.wp, "docPr")).forEach((d) => d.setAttribute("descr", String(edit.descr || "")));
    Array.from(box.getElementsByTagNameNS(COMPOSE_NS.pic, "cNvPr")).forEach((d) => d.setAttribute("descr", String(edit.descr || "")));
  } else if (edit.action === "delete") {
    let run = drawing.parentNode;
    while (run && run.localName !== "r") run = run.parentNode;
    if (!run) return { xml, count: 0 };
    run.parentNode.removeChild(run);
    // akapit był tylko obrazem — znika, jeśli coś zostaje w tym miejscu
    const parent = p.parentNode;
    const siblingsP = Array.from(parent.childNodes).filter((n) => n.localName === "p");
    if (!getParagraphText(p).trim() && !p.getElementsByTagNameNS(W_NS, "drawing").length && siblingsP.length > 1 && !composeDirectChild(composeDirectChild(p, "pPr") || p, "sectPr")) parent.removeChild(p);
  } else {
    return { xml, count: 0 };
  }
  return { xml: composeSerialize(doc), count: 1 };
}

// ── op "runStyle": kolor czcionki / wyróżnienie fragmentu akapitu ────────────
// { index, start, end, color?: "#RRGGBB" | "auto", highlight?: nazwa Worda | "#RRGGBB" | "" }
// Przez model akapitu (te same zasady co pisanie): linki i pola-wyspy zostają nietknięte.
function applyRunStyleInXml(xml, edit) {
  const doc = composeParse(xml);
  const p = collectParagraphElements(doc.documentElement, "all")[edit.index];
  if (!p) return { xml, count: 0 };
  const runs = extractRunsFromParagraphXml(p);
  const total = previewRunsToPlainText(runs).length;
  const start = Math.max(0, Math.min(total, edit.start | 0));
  const end = Math.max(start, Math.min(total, edit.end | 0));
  if (end <= start) return { xml, count: 0 };
  const { before, mid, after } = composeSliceRuns(runs, start, end);
  const styled = mid.map((r) => {
    if (r.break || r.island) return r;
    const out = { ...r };
    if (edit.color !== undefined) { if (!edit.color || edit.color === "auto") delete out.color; else out.color = edit.color; }
    if (edit.highlight !== undefined) { if (!edit.highlight) delete out.highlight; else out.highlight = edit.highlight; }
    return out;
  });
  applyRunsToParagraphXml(p, [...before, ...styled, ...after]);
  return { xml: composeSerialize(doc), count: 1 };
}

// ── komentarze: dodaj / odpowiedz / rozwiąż ──────────────────────────────────
// comments.xml (treść) + znaczniki w treści (commentRangeStart / End + fragment z
// commentReference) + commentsExtended.xml (Word 2013+: odpowiedzi i „rozwiązany”, po w14:paraId).
const COMPOSE_W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

function composeParaId(existing) {
  let id;
  do { id = Math.floor(Math.random() * 0x7fffffff).toString(16).toUpperCase().padStart(8, "0"); } while (existing?.has(id));
  existing?.add(id);
  return id;
}

async function composeCommentsParts(zip) {
  const cXml = await composeEnsurePart(zip, "word/comments.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="${W_NS}" xmlns:w14="${COMPOSE_W14}"/>`,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments");
  return composeParse(cXml);
}

async function composeCommentsExDoc(zip) {
  const xml = await composeEnsurePart(zip, "word/commentsExtended.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w15:commentsEx xmlns:w15="${COMPOSE_W15}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w15"/>`,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended");
  return composeParse(xml);
}

// Treść komentarza: rich = akapity → fragmenty { text, b, i, u, s, color, hl } (jak revCommentRich),
// inaczej zwykły tekst (wiersze = akapity). Formatowanie jak w Wordzie: pogrubienie, kursywa,
// przekreślenie, kolor, wyróżnienie, podkreślenie — w:rPr w kolejności schematu.
function composeCommentRichOf(rich, text) {
  if (Array.isArray(rich) && rich.length) return rich.map((p) => (Array.isArray(p) ? p : []));
  return String(text || "").replace(/\r\n?/g, "\n").split("\n").map((line) => (line ? [{ text: line }] : []));
}
function composeCommentRun(cDoc, run) {
  const r = cDoc.createElementNS(W_NS, "w:r");
  const props = [];
  if (run.b) props.push(["b"]);
  if (run.i) props.push(["i"]);
  if (run.s) props.push(["strike"]);
  if (/^[0-9A-F]{6}$/i.test(run.color || "")) props.push(["color", String(run.color).toUpperCase()]);
  if (run.hl && /^[A-Za-z]+$/.test(run.hl)) props.push(["highlight", run.hl]);
  if (run.u) props.push(["u", "single"]);
  if (props.length) {
    const rPr = cDoc.createElementNS(W_NS, "w:rPr");
    props.forEach(([name, val]) => {
      const el = cDoc.createElementNS(W_NS, `w:${name}`);
      if (val) el.setAttributeNS(W_NS, "w:val", val);
      rPr.appendChild(el);
    });
    r.appendChild(rPr);
  }
  String(run.text || "").split("\t").forEach((part, k) => {
    if (k) r.appendChild(cDoc.createElementNS(W_NS, "w:tab"));
    if (!part) return;
    const t = cDoc.createElementNS(W_NS, "w:t");
    t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    t.textContent = sanitizeXmlText(part);
    r.appendChild(t);
  });
  return r;
}
// Akapity komentarza; pierwszy z fragmentem-znacznikiem (annotationRef — „dymek” w Wordzie),
// ostatni z paraId (commentsExtended: odpowiedzi i „rozwiązany” wiążą się po nim).
// opts.pPr / opts.refRun — kopiowane z edytowanego komentarza (styl akapitu, styl znacznika).
function composeCommentParas(cDoc, paras, lastParaId, opts = {}) {
  return paras.map((runs, i) => {
    const p = cDoc.createElementNS(W_NS, "w:p");
    if (i === paras.length - 1) { p.setAttributeNS(COMPOSE_W14, "w14:paraId", lastParaId); p.setAttributeNS(COMPOSE_W14, "w14:textId", opts.textId || "77777777"); }
    if (opts.pPr) p.appendChild(opts.pPr.cloneNode(true));
    if (i === 0) {
      if (opts.refRun) p.appendChild(opts.refRun.cloneNode(true));
      else {
        const r = cDoc.createElementNS(W_NS, "w:r");
        r.appendChild(cDoc.createElementNS(W_NS, "w:annotationRef"));
        p.appendChild(r);
      }
    }
    runs.forEach((run) => { if (run?.text) p.appendChild(composeCommentRun(cDoc, run)); });
    return p;
  });
}

function composeCommentEl(cDoc, { id, author, initials, text, rich, paraIds }) {
  const c = cDoc.createElementNS(W_NS, "w:comment");
  c.setAttributeNS(W_NS, "w:id", String(id));
  c.setAttributeNS(W_NS, "w:author", author || "Autor");
  c.setAttributeNS(W_NS, "w:date", new Date().toISOString().replace(/\.\d+Z$/, "Z"));
  if (initials) c.setAttributeNS(W_NS, "w:initials", initials);
  const lastParaId = composeParaId(paraIds);
  composeCommentParas(cDoc, composeCommentRichOf(rich, text), lastParaId).forEach((p) => c.appendChild(p));
  cDoc.documentElement.appendChild(c);
  return lastParaId;
}

const composeRichText = (paras) => paras.map((runs) => runs.map((r) => r?.text || "").join("")).join("\n");

// op "commentEdit": { id, rich | text } — nowa treść komentarza (albo odpowiedzi). Autor, data,
// znaczniki w treści i powiązania (paraId ostatniego akapitu) zostają. Komentarza z czymś, czego
// model nie zapisze (link, pole, wzmianka, inny krój…), nie ruszamy — revCommentRich.editable.
async function applyCommentEditInZip(zip, xml, edit) {
  if (!zip.file("word/comments.xml")) return { xml, count: 0 };
  const paras = composeCommentRichOf(edit.rich, edit.text);
  if (!composeRichText(paras).trim()) return { xml, count: 0 };
  const cDoc = await composeCommentsParts(zip);
  const info = composeCommentInfo(cDoc);
  const c = info.comments.find((x) => x.getAttributeNS(W_NS, "id") === String(edit.id));
  if (!c || (typeof revCommentRich === "function" && !revCommentRich(c).editable)) return { xml, count: 0 };
  const old = Array.from(c.getElementsByTagNameNS(W_NS, "p"));
  const last = old[old.length - 1];
  const lastParaId = (last && last.getAttributeNS(COMPOSE_W14, "paraId")) || composeParaId(info.paraIds);
  const textId = last?.getAttributeNS(COMPOSE_W14, "textId") || undefined;
  const pPr = old[0] ? Array.from(old[0].childNodes).find((n) => n.nodeType === 1 && n.localName === "pPr") : null;
  const refRun = c.getElementsByTagNameNS(W_NS, "annotationRef")[0]?.parentNode || null;
  const fresh = composeCommentParas(cDoc, paras, lastParaId, { pPr, refRun: refRun?.localName === "r" ? refRun : null, textId });
  old.forEach((p) => p.parentNode.removeChild(p));
  fresh.forEach((p) => c.appendChild(p));
  zip.file("word/comments.xml", composeSerialize(cDoc));
  return { xml, count: 1 };
}

function composeCommentInfo(cDoc) {
  const comments = Array.from(cDoc.getElementsByTagNameNS(W_NS, "comment"));
  const ids = comments.map((c) => parseInt(c.getAttributeNS(W_NS, "id"), 10)).filter(Number.isFinite);
  const paraIds = new Set(Array.from(cDoc.getElementsByTagNameNS(W_NS, "p")).map((p) => (p.getAttributeNS(COMPOSE_W14, "paraId") || "").toUpperCase()).filter(Boolean));
  return { comments, nextId: (ids.length ? Math.max(...ids) : -1) + 1, paraIds };
}

// paraId ostatniego akapitu komentarza (dopisany, gdy go brak — potrzebny do odpowiedzi i „rozwiązany”)
function composeCommentParaId(cDoc, comment, paraIds) {
  const ps = Array.from(comment.getElementsByTagNameNS(W_NS, "p"));
  const last = ps[ps.length - 1];
  if (!last) return "";
  let pid = last.getAttributeNS(COMPOSE_W14, "paraId");
  if (!pid) { pid = composeParaId(paraIds); last.setAttributeNS(COMPOSE_W14, "w14:paraId", pid); }
  return pid.toUpperCase();
}

function composeCommentEx(exDoc, paraId) {
  let ex = Array.from(exDoc.getElementsByTagNameNS(COMPOSE_W15, "commentEx")).find((e) => (e.getAttributeNS(COMPOSE_W15, "paraId") || "").toUpperCase() === paraId);
  if (!ex) {
    ex = exDoc.createElementNS(COMPOSE_W15, "w15:commentEx");
    ex.setAttributeNS(COMPOSE_W15, "w15:paraId", paraId);
    ex.setAttributeNS(COMPOSE_W15, "w15:done", "0");
    exDoc.documentElement.appendChild(ex);
  }
  return ex;
}

// op "commentAdd": { index, start, end, text | rich, author, initials }
async function applyCommentAddInZip(zip, xml, edit) {
  const doc = composeParse(xml);
  const p = collectParagraphElements(doc.documentElement, "all")[edit.index];
  if (!p || !composeRichText(composeCommentRichOf(edit.rich, edit.text)).trim()) return { xml, count: 0 };
  const cDoc = await composeCommentsParts(zip);
  const info = composeCommentInfo(cDoc);
  const id = info.nextId;
  const paraId = composeCommentEl(cDoc, { id, author: edit.author, initials: edit.initials, text: edit.text, rich: edit.rich, paraIds: info.paraIds });
  const runs = extractRunsFromParagraphXml(p);
  const total = previewRunsToPlainText(runs).length;
  const start = Math.max(0, Math.min(total, edit.start | 0));
  const end = Math.max(start, Math.min(total, edit.end | 0));
  const { before, mid, after } = composeSliceRuns(runs, start, end);
  const isl = (x) => ({ island: x, text: "" });
  applyRunsToParagraphXml(p, [
    ...before,
    isl(`<w:commentRangeStart xmlns:w="${W_NS}" w:id="${id}"/>`),
    ...mid,
    isl(`<w:commentRangeEnd xmlns:w="${W_NS}" w:id="${id}"/>`),
    isl(`<w:r xmlns:w="${W_NS}"><w:commentReference w:id="${id}"/></w:r>`),
    ...after,
  ]);
  zip.file("word/comments.xml", composeSerialize(cDoc));
  const exDoc = await composeCommentsExDoc(zip);
  composeCommentEx(exDoc, paraId);
  zip.file("word/commentsExtended.xml", composeSerialize(exDoc));
  return { xml: composeSerialize(doc), count: 1 };
}

// op "commentReply": { id (komentarz nadrzędny), text, author, initials }
// op "commentDone":  { id, done: true|false }
async function applyCommentThreadInZip(zip, xml, edit) {
  if (!zip.file("word/comments.xml")) return { xml, count: 0 };
  const cDoc = await composeCommentsParts(zip);
  const info = composeCommentInfo(cDoc);
  const parent = info.comments.find((c) => c.getAttributeNS(W_NS, "id") === String(edit.id));
  if (!parent) return { xml, count: 0 };
  const parentPid = composeCommentParaId(cDoc, parent, info.paraIds);
  const exDoc = await composeCommentsExDoc(zip);
  const parentEx = composeCommentEx(exDoc, parentPid);
  if (edit.op === "commentDone") {
    parentEx.setAttributeNS(COMPOSE_W15, "w15:done", edit.done ? "1" : "0");
    zip.file("word/comments.xml", composeSerialize(cDoc));
    zip.file("word/commentsExtended.xml", composeSerialize(exDoc));
    return { xml, count: 1 };
  }
  if (!composeRichText(composeCommentRichOf(edit.rich, edit.text)).trim()) return { xml, count: 0 };
  const id = info.nextId;
  const pid = composeCommentEl(cDoc, { id, author: edit.author, initials: edit.initials, text: edit.text, rich: edit.rich, paraIds: info.paraIds });
  const ex = composeCommentEx(exDoc, pid);
  ex.setAttributeNS(COMPOSE_W15, "w15:paraIdParent", parentPid);
  // znaczniki odpowiedzi tuż przy znacznikach komentarza nadrzędnego (tak robi Word)
  const doc = composeParse(xml);
  const pid0 = String(edit.id);
  const find = (tag) => Array.from(doc.getElementsByTagNameNS(W_NS, tag)).find((m) => m.getAttributeNS(W_NS, "id") === pid0);
  const s = find("commentRangeStart");
  const e = find("commentRangeEnd");
  const ref = find("commentReference");
  if (s) s.parentNode.insertBefore(composeEl(doc, "commentRangeStart", { id }), s.nextSibling);
  if (e) e.parentNode.insertBefore(composeEl(doc, "commentRangeEnd", { id }), e.nextSibling);
  const refRun = ref?.parentNode;
  if (refRun?.localName === "r") {
    const r = composeEl(doc, "r");
    r.appendChild(composeEl(doc, "commentReference", { id }));
    refRun.parentNode.insertBefore(r, refRun.nextSibling);
  }
  zip.file("word/comments.xml", composeSerialize(cDoc));
  zip.file("word/commentsExtended.xml", composeSerialize(exDoc));
  return { xml: composeSerialize(doc), count: 1 };
}

// ── nagłówek i stopka (z numerem strony) ─────────────────────────────────────
// Prosty model, jak „Wstaw → Nagłówek/Stopka/Numer strony” w Wordzie: tekst nagłówka, tekst
// stopki, numer strony (format + wyrównanie, w stopce), „inna pierwsza strona” (bez nagłówka i
// stopki). Dotyczy wszystkich sekcji dokumentu. Numer = pole PAGE (i NUMPAGES dla „z N”) —
// Word liczy je sam; w podglądzie podstawia je fixPreviewPageNumbers (compose-ui.js).
const COMPOSE_REL_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const COMPOSE_REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
const COMPOSE_PN_FORMATS = ["n", "page", "pageOf", "dash"];

function composeSectPrs(doc) {
  return Array.from(doc.getElementsByTagNameNS(W_NS, "sectPr")).filter((s) => s.parentNode?.localName === "body" || s.parentNode?.localName === "pPr");
}

function composeHfRef(sect, kind, type) {
  return Array.from(sect.childNodes).find((n) => n.localName === `${kind}Reference` && (n.getAttributeNS(W_NS, "type") || "default") === type) || null;
}

function composeFieldRuns(instr, cached, rStyle) {
  const rpr = rStyle ? `<w:rPr><w:rStyle w:val="${rStyle}"/></w:rPr>` : "";
  return `<w:r>${rpr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rpr}<w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r>${rpr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rpr}<w:t>${cached}</w:t></w:r><w:r>${rpr}<w:fldChar w:fldCharType="end"/></w:r>`;
}

function composeHfParagraph(inner, align) {
  const jc = COMPOSE_ALIGNS.includes(align) && align !== "left" ? `<w:pPr><w:jc w:val="${align}"/></w:pPr>` : "";
  return `<w:p>${jc}${inner}</w:p>`;
}

function composeTextRun(text) {
  return text ? `<w:r><w:t xml:space="preserve">${composeXmlText(sanitizeXmlText(text))}</w:t></w:r>` : "";
}

function composePageNumberXml(fmt, lang, styleId, total) {
  const pl = lang !== "en";
  const page = composeFieldRuns("PAGE", "1", styleId);
  const pages = composeFieldRuns("NUMPAGES", String(total || 1), styleId);
  if (fmt === "page") return composeTextRun(pl ? "Strona " : "Page ") + page;
  if (fmt === "pageOf") return composeTextRun(pl ? "Strona " : "Page ") + page + composeTextRun(pl ? " z " : " of ") + pages;
  if (fmt === "dash") return composeTextRun("– ") + page + composeTextRun(" –");
  return page;
}

// Stan do okienka: { header: { text, align, complex }, footer: {…, number: fmt|null, numAlign}, firstDifferent, sections }
async function readHeaderFooterState(bytes) {
  const out = { header: { text: "", align: "left", complex: false }, footer: { text: "", align: "left", complex: false, number: null, numAlign: "center" }, firstDifferent: false, sections: 1 };
  if (!bytes || !window.JSZip) return out;
  const zip = await loadDocxZipCached(bytes);
  const doc = await getDocumentXmlDom(bytes);
  if (!doc) return out;
  const sects = composeSectPrs(doc);
  out.sections = Math.max(1, sects.length);
  const body = composeDirectChild(doc.getElementsByTagNameNS(W_NS, "body")[0], "sectPr") || sects[sects.length - 1];
  if (!body) return out;
  out.firstDifferent = !!composeDirectChild(body, "titlePg");
  const rels = (await composeReadRels(zip)).doc;
  const target = (rid) => Array.from(rels.documentElement.getElementsByTagName("Relationship")).find((r) => r.getAttribute("Id") === rid)?.getAttribute("Target");
  for (const kind of ["header", "footer"]) {
    const ref = composeHfRef(body, kind, "default");
    const path = ref && target(ref.getAttributeNS(R_NS, "id") || ref.getAttribute("r:id"));
    const file = path && zip.file(`word/${path.replace(/^\/?word\//, "")}`);
    if (!file) continue;
    const part = composeParse(await file.async("string"));
    const root = part.documentElement;
    const st = out[kind];
    st.complex = ["tbl", "drawing", "pict", "sdt", "txbxContent"].some((tag) => root.getElementsByTagNameNS(W_NS, tag).length)
      || Array.from(root.childNodes).filter((n) => n.localName === "p").length > 2;
    Array.from(root.childNodes).filter((n) => n.localName === "p").forEach((p) => {
      const instr = Array.from(p.getElementsByTagNameNS(W_NS, "instrText")).map((x) => x.textContent).join(" ")
        + " " + Array.from(p.getElementsByTagNameNS(W_NS, "fldSimple")).map((x) => x.getAttributeNS(W_NS, "instr")).join(" ");
      const jc = composeDirectChild(composeDirectChild(p, "pPr") || p, "jc")?.getAttributeNS(W_NS, "val") || "left";
      const align = jc === "both" || jc === "start" ? "left" : jc === "end" ? "right" : jc;
      if (/\bPAGE\b/.test(instr)) {
        const plain = Array.from(p.childNodes).filter((r) => r.localName === "r" && !r.getElementsByTagNameNS(W_NS, "fldChar").length && !r.getElementsByTagNameNS(W_NS, "instrText").length).map((r) => Array.from(r.getElementsByTagNameNS(W_NS, "t")).map((x) => x.textContent).join("")).join("");
        st.number = /\bNUMPAGES\b/.test(instr) ? "pageOf" : /[–-]/.test(plain) ? "dash" : /\S/.test(plain) ? "page" : "n";
        st.numAlign = align;
        if (kind === "header") { out.footer.number = st.number; out.footer.numAlign = align; st.number = null; }
      } else {
        const text = getParagraphText(p).replace(/\n/g, " ");
        if (text.trim() && !st.text) { st.text = text; st.align = align; }
      }
    });
  }
  return out;
}

// op "headerFooter": { header: { text, align }, footer: { text, align }, number: { fmt|null, align }, firstDifferent, lang, total }
async function applyHeaderFooterInZip(zip, xml, edit) {
  const doc = composeParse(xml);
  const sects = composeSectPrs(doc);
  if (!sects.length) return { xml, count: 0 };
  const relsInfo = await composeReadRels(zip);
  const rels = relsInfo.doc;
  const relEls = () => Array.from(rels.documentElement.getElementsByTagName("Relationship"));
  const ct = zip.file("[Content_Types].xml") ? await zip.file("[Content_Types].xml").async("string") : "";
  let ctXml = ct;
  const body = composeDirectChild(doc.getElementsByTagNameNS(W_NS, "body")[0], "sectPr") || sects[sects.length - 1];
  const numStyle = edit.number?.fmt ? await composeEnsureCharStyle(zip, "page number", "PageNumber", "<w:uiPriority w:val=\"99\"/><w:unhideWhenUsed/>") : null;
  const ns = `xmlns:w="${W_NS}" xmlns:r="${R_NS}"`;
  const newPart = (kind, inner) => {
    let n = 1;
    while (zip.file(`word/${kind}${n}.xml`)) n++;
    const name = `${kind}${n}.xml`;
    zip.file(`word/${name}`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${kind === "header" ? "hdr" : "ftr"} ${ns}>${inner}</w:${kind === "header" ? "hdr" : "ftr"}>`);
    if (!ctXml.includes(`PartName="/word/${name}"`)) ctXml = ctXml.replace("</Types>", `<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/></Types>`);
    return composeAddRel(rels, kind === "header" ? COMPOSE_REL_HEADER : COMPOSE_REL_FOOTER, name, false);
  };
  const setRef = (sect, kind, type, rid) => {
    const old = composeHfRef(sect, kind, type);
    if (old) sect.removeChild(old);
    if (!rid) return;
    const ref = composeEl(doc, `${kind}Reference`, { type });
    ref.setAttributeNS(R_NS, "r:id", rid);
    // odwołania nagłówków/stopek są pierwsze w w:sectPr
    let at = sect.firstChild;
    while (at && (at.localName === "headerReference" || at.localName === "footerReference")) at = at.nextSibling;
    sect.insertBefore(ref, at);
  };
  for (const kind of ["header", "footer"]) {
    const conf = edit[kind] || {};
    let inner = conf.text ? composeHfParagraph(composeTextRun(conf.text), conf.align) : "";
    if (kind === "footer" && edit.number?.fmt) inner += composeHfParagraph(composePageNumberXml(edit.number.fmt, edit.lang, numStyle, edit.total), edit.number.align || "center");
    if (!inner) { sects.forEach((s) => setRef(s, kind, "default", null)); continue; }
    // ta sama część dla wszystkich sekcji: istniejąca (z odwołania sekcji końcowej) albo nowa
    const oldRef = composeHfRef(body, kind, "default");
    const oldRid = oldRef && (oldRef.getAttributeNS(R_NS, "id") || oldRef.getAttribute("r:id"));
    const oldTarget = oldRid && relEls().find((r) => r.getAttribute("Id") === oldRid)?.getAttribute("Target");
    let rid;
    if (oldTarget && zip.file(`word/${oldTarget.replace(/^\/?word\//, "")}`)) {
      const path = `word/${oldTarget.replace(/^\/?word\//, "")}`;
      const part = composeParse(await zip.file(path).async("string"));
      const root = part.documentElement;
      while (root.firstChild) root.removeChild(root.firstChild);
      const frag = composeParse(`<x ${ns}>${inner}</x>`).documentElement;
      Array.from(frag.childNodes).forEach((n) => root.appendChild(part.importNode(n, true)));
      zip.file(path, composeSerialize(part));
      rid = oldRid;
    } else {
      rid = newPart(kind, inner);
    }
    sects.forEach((s) => setRef(s, kind, "default", rid));
  }
  // inna pierwsza strona: titlePg + puste części „first”
  let emptyH = null; let emptyF = null;
  sects.forEach((s) => {
    const tp = composeDirectChild(s, "titlePg");
    if (edit.firstDifferent) {
      if (!emptyH) { emptyH = newPart("header", "<w:p/>"); emptyF = newPart("footer", "<w:p/>"); }
      setRef(s, "header", "first", emptyH);
      setRef(s, "footer", "first", emptyF);
      if (!tp) {
        const el = composeEl(doc, "titlePg");
        const after = ["textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"].map((n) => composeDirectChild(s, n)).find(Boolean);
        s.insertBefore(el, after || null);
      }
    } else {
      if (tp) s.removeChild(tp);
      setRef(s, "header", "first", null);
      setRef(s, "footer", "first", null);
    }
  });
  // osierocone części nagłówków/stopek (nikt się do nich nie odwołuje) — precz z paczki
  const used = new Set(Array.from(doc.getElementsByTagNameNS(W_NS, "headerReference")).concat(Array.from(doc.getElementsByTagNameNS(W_NS, "footerReference"))).map((r) => r.getAttributeNS(R_NS, "id") || r.getAttribute("r:id")));
  relEls().forEach((r) => {
    const type = r.getAttribute("Type");
    if ((type !== COMPOSE_REL_HEADER && type !== COMPOSE_REL_FOOTER) || used.has(r.getAttribute("Id"))) return;
    const name = (r.getAttribute("Target") || "").replace(/^\/?word\//, "");
    zip.remove(`word/${name}`);
    ctXml = ctXml.replace(new RegExp(`<Override PartName="/word/${name.replace(/\./g, "\\.")}"[^>]*/>`), "");
    r.parentNode.removeChild(r);
  });
  zip.file(relsInfo.path, composeSerialize(rels));
  if (ctXml) zip.file("[Content_Types].xml", ctXml);
  return { xml: composeSerialize(doc), count: 1 };
}
