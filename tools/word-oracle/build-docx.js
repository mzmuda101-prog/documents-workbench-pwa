// build-docx.js — minimalne pliki .docx do porównań z prawdziwym Wordem (word-oracle).
// Tylko testy: nie trafia do dist/.
const JSZip = require("jszip");

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// p({ text, numId, ilvl, style, lang, raw }) → <w:p>
function p(o) {
  if (typeof o === "string") o = { text: o };
  const ppr = [];
  if (o.style) ppr.push(`<w:pStyle w:val="${o.style}"/>`);
  if (o.numId != null) ppr.push(`<w:numPr><w:ilvl w:val="${o.ilvl || 0}"/><w:numId w:val="${o.numId}"/></w:numPr>`);
  if (o.pPr) ppr.push(o.pPr);
  const rpr = o.lang ? `<w:rPr><w:lang w:val="${o.lang}"/></w:rPr>` : "";
  const run = o.raw != null ? o.raw : `<w:r>${rpr}<w:t xml:space="preserve">${esc(o.text || "")}</w:t></w:r>`;
  const mark = o.lang ? `<w:rPr><w:lang w:val="${o.lang}"/></w:rPr>` : "";
  return `<w:p>${ppr.length || mark ? `<w:pPr>${ppr.join("")}${mark}</w:pPr>` : ""}${run}</w:p>`;
}

// lvl({ ilvl, fmt, text, start, isLgl, restart, suff, jc }) → <w:lvl>
function lvl(o) {
  return `<w:lvl w:ilvl="${o.ilvl || 0}">${o.start != null ? `<w:start w:val="${o.start}"/>` : '<w:start w:val="1"/>'}<w:numFmt w:val="${o.fmt || "decimal"}"/>${o.restart != null ? `<w:lvlRestart w:val="${o.restart}"/>` : ""}${o.pStyle ? `<w:pStyle w:val="${o.pStyle}"/>` : ""}${o.isLgl ? "<w:isLgl/>" : ""}${o.suff ? `<w:suff w:val="${o.suff}"/>` : ""}<w:lvlText w:val="${esc(o.text != null ? o.text : `%${(o.ilvl || 0) + 1}.`)}"/><w:lvlJc w:val="${o.jc || "left"}"/><w:pPr><w:ind w:left="${720 * ((o.ilvl || 0) + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
}

async function buildDocx({ body, abstractNums = [], nums = [], styles = "", settings = "", sectPr = "", lang = "pl-PL", extraParts = {} }) {
  const real = new JSZip();
  // stała data w ZIP-ie: te same dane = te same bajty (pamięć wyników Worda po skrócie pliku)
  const zip = { file: (name, data) => real.file(name, data, { date: new Date("2026-01-01T00:00:00Z") }) };
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>${Object.keys(extraParts).map((k) => `<Override PartName="/word/${k}" ContentType="${extraParts[k].type}"/>`).join("")}</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>${Object.keys(extraParts).map((k, i) => `<Relationship Id="rIdX${i}" Type="${extraParts[k].rel}" Target="${k}"/>`).join("")}</Relationships>`);
  zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="${lang}" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>${styles}</w:styles>`);
  zip.file("word/settings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}>${settings}<w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
  zip.file("word/numbering.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}>${abstractNums.map((a, i) => `<w:abstractNum w:abstractNumId="${a.id != null ? a.id : i}"><w:multiLevelType w:val="${a.multi || "hybridMultilevel"}"/>${a.levels.map(lvl).join("")}</w:abstractNum>`).join("")}${nums.map((n) => `<w:num w:numId="${n.id}"><w:abstractNumId w:val="${n.abs}"/>${(n.overrides || []).map((o) => `<w:lvlOverride w:ilvl="${o.ilvl || 0}">${o.start != null ? `<w:startOverride w:val="${o.start}"/>` : ""}${o.lvl ? lvl({ ilvl: o.ilvl || 0, ...o.lvl }) : ""}</w:lvlOverride>`).join("")}</w:num>`).join("")}</w:numbering>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body.map((b) => (typeof b === "string" && b.startsWith("<") ? b : p(b))).join("")}<w:sectPr>${sectPr}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  for (const [k, v] of Object.entries(extraParts)) zip.file(`word/${k}`, v.xml);
  return real.generateAsync({ type: "nodebuffer" });
}

module.exports = { buildDocx, p, lvl, esc, W };
