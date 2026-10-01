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
const COMPOSE_STYLE_KEYS = ["normal", "title", "subtitle", "h1", "h2", "h3", "quote"];
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
    [...COMPOSE_STYLE_KEYS, ...COMPOSE_TOC_KEYS].forEach((key) => {
      const id = key === "normal" ? idx.defaultId : idx.byName.get(COMPOSE_STYLE_DEFS[key].name);
      if (id && !map.has(docxStyleClassName(id))) map.set(docxStyleClassName(id), key);
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
    if (key && COMPOSE_STYLE_KEYS.includes(key)) return key;
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
function composeEnsureBookmark(doc, p) {
  const own = Array.from(p.childNodes).find((n) => n.localName === "bookmarkStart" && n.namespaceURI === W_NS
    && !/^_GoBack$/.test(n.getAttributeNS(W_NS, "name")));
  if (own) return own.getAttributeNS(W_NS, "name");
  const all = Array.from(doc.getElementsByTagNameNS(W_NS, "bookmarkStart"));
  const names = new Set(all.map((b) => b.getAttributeNS(W_NS, "name")));
  const maxId = all.reduce((m, b) => Math.max(m, parseInt(b.getAttributeNS(W_NS, "id"), 10) || 0), 0);
  let name;
  do { name = `_Ref${String(Math.floor(1e8 + Math.random() * 9e8))}`; } while (names.has(name));
  const id = String(maxId + 1);
  const start = composeEl(doc, "bookmarkStart", { id, name });
  const end = composeEl(doc, "bookmarkEnd", { id });
  const pPr = composeDirectChild(p, "pPr");
  p.insertBefore(start, pPr ? pPr.nextSibling : p.firstChild);
  p.appendChild(end);
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
    if (Number.isInteger(edit.targetIndex)) {
      const target = paragraphs[edit.targetIndex];
      if (!target) return { xml, count: 0 };
      link = `#${composeEnsureBookmark(doc, target)}`;
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
  const firstNum = composeDirectChild(root, "num");
  if (absId == null) {
    absId = String(Array.from(idx.abstracts.keys()).reduce((m, k) => Math.max(m, parseInt(k, 10) || 0), -1) + 1);
    const frag = composeParse(`<w:numbering xmlns:w="${W_NS}">${composeAbstractNumXml(kind, absId)}</w:numbering>`);
    root.insertBefore(numDoc.importNode(frag.documentElement.firstChild, true), firstNum); // abstractNum przed num (schemat)
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
    root.appendChild(num);
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
  // tabulator do prawego marginesu z kropkami
  const sect = composeDirectChild(body, "sectPr");
  const pgW = parseInt(sect && composeDirectChild(sect, "pgSz")?.getAttributeNS(W_NS, "w"), 10) || 11906;
  const mar = sect && composeDirectChild(sect, "pgMar");
  const tabPos = Math.max(2000, pgW - (parseInt(mar?.getAttributeNS(W_NS, "left"), 10) || 1418) - (parseInt(mar?.getAttributeNS(W_NS, "right"), 10) || 1418));
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
    pr = `<w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>`;
    content = `<w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/></w:rPr><w:t>☐</w:t></w:r>`;
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
