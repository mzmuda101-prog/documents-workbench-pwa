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
  quote: {
    name: "quote", id: "Quote",
    body: `<w:basedOn w:val="{NORMAL}"/><w:next w:val="{NORMAL}"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr>`,
  },
};
const COMPOSE_STYLE_KEYS = ["normal", "title", "subtitle", "h1", "h2", "h3", "quote"];

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
    COMPOSE_STYLE_KEYS.forEach((key) => {
      const id = key === "normal" ? idx.defaultId : idx.byName.get(COMPOSE_STYLE_DEFS[key].name);
      if (id) map.set(docxStyleClassName(id), key);
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
    const frag = composeParse(`<w:styles xmlns:w="${W_NS}"><w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${def.name}"/>${body}</w:style></w:styles>`);
    doc.documentElement.appendChild(doc.importNode(frag.documentElement.firstChild, true));
    cache.stylesXml = composeSerialize(doc);
    cache.stylesDirty = true;
  }
  cache.set(key, id);
  return id;
}

// ── op "paraFormat": styl i/lub wyrównanie akapitów ──────────────────────────
async function applyParaFormatInZip(zip, xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const targets = (edit.indices || []).map((i) => paragraphs[i]).filter(Boolean);
  if (!targets.length) return { xml, count: 0 };
  const cache = new Map();
  let styleId;
  if (edit.style) styleId = await composeEnsureStyle(zip, edit.style, cache);
  let count = 0;
  targets.forEach((p) => {
    const pPr = composeEnsurePPr(p);
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
    styleXml.push(`<w:style w:type="paragraph" w:styleId="${def.id}"><w:name w:val="${def.name}"/>${def.body.replaceAll("{NORMAL}", "Normal")}</w:style>`);
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
<w:settings xmlns:w="${W_NS}"><w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/><w:themeFontLang w:val="${langTag}"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
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
    if (key) return key;
  }
  for (const cls of p.classList) { // nagłówek w stylu własnym pliku (np. „Rozdział” z outlineLvl)
    const level = docHeadingStyleClasses?.get?.(cls);
    if (level) return level <= 3 ? `h${level}` : "h3";
  }
  return "normal";
}

// Nowy akapit po Enterze na końcu nagłówka/tytułu/cytatu: zdejmij klasę stylu w podglądzie.
// Zwraca true, gdy akapit miał taki styl (plik dostaje wtedy nextNormal).
function composeStripNextStyle(newP) {
  let had = false;
  Array.from(newP.classList).forEach((cls) => {
    const key = docComposeStyleClasses.get(cls);
    if ((key && key !== "normal") || docHeadingStyleClasses?.get?.(cls)) {
      newP.classList.remove(cls);
      had = true;
    }
  });
  return had;
}
