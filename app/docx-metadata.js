// DOCX package metadata — read/write docProps/core.xml (title, author, keywords).

const CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS = "http://purl.org/dc/elements/1.1/";
const DCTERMS_NS = "http://purl.org/dc/terms/";
const XSI_NS = "http://www.w3.org/2001/XMLSchema-instance";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const CORE_REL_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const CORE_CONTENT_TYPE = "application/vnd.openxmlformats-package.core-properties+xml";

function isoDocxTimestamp(date = new Date()) {
  return date.toISOString().replace(/\.\d{3}Z$/, "Z");
}

function createDefaultCoreXml() {
  const now = isoDocxTimestamp();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="${CP_NS}" xmlns:dc="${DC_NS}" xmlns:dcterms="${DCTERMS_NS}" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="${XSI_NS}">
  <dc:title></dc:title>
  <dc:creator></dc:creator>
  <cp:keywords></cp:keywords>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}

function getMetaElement(doc, localName, ns) {
  const els = doc.getElementsByTagNameNS(ns, localName);
  return els.length ? els[0] : null;
}

function readMetaText(doc, localName, ns) {
  return (getMetaElement(doc, localName, ns)?.textContent || "").trim();
}

function parseCoreMetadataXml(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  return {
    title: readMetaText(doc, "title", DC_NS),
    creator: readMetaText(doc, "creator", DC_NS),
    keywords: readMetaText(doc, "keywords", CP_NS),
  };
}

async function extractCoreMetadataFromDocx(bytes) {
  if (!window.JSZip) return { title: "", creator: "", keywords: "" };
  const zip = await window.JSZip.loadAsync(bytes);
  const coreFile = zip.file("docProps/core.xml");
  if (!coreFile) return { title: "", creator: "", keywords: "" };
  return parseCoreMetadataXml(await coreFile.async("string"));
}

function setMetaText(doc, root, localName, ns, value) {
  const text = sanitizeXmlText(String(value ?? ""));
  let el = getMetaElement(doc, localName, ns);
  if (!el) {
    el = doc.createElementNS(ns, localName);
    root.appendChild(el);
  }
  if ((el.textContent || "") === text) return false;
  el.textContent = text;
  return true;
}

function touchCoreModified(doc, root) {
  const now = isoDocxTimestamp();
  let mod = getMetaElement(doc, "modified", DCTERMS_NS);
  if (!mod) {
    mod = doc.createElementNS(DCTERMS_NS, "modified");
    mod.setAttributeNS(XSI_NS, "xsi:type", "dcterms:W3CDTF");
    root.appendChild(mod);
  }
  if (mod.textContent !== now) mod.textContent = now;
}

function applyCoreMetadataInXml(xml, fields) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const root = doc.documentElement;
  if (!root || root.localName !== "coreProperties") {
    return applyCoreMetadataInXml(createDefaultCoreXml(), fields);
  }

  let count = 0;
  if (fields.title !== undefined && setMetaText(doc, root, "title", DC_NS, fields.title)) count++;
  if (fields.creator !== undefined && setMetaText(doc, root, "creator", DC_NS, fields.creator)) count++;
  if (fields.keywords !== undefined && setMetaText(doc, root, "keywords", CP_NS, fields.keywords)) count++;

  if (count > 0) touchCoreModified(doc, root);
  return { xml: new XMLSerializer().serializeToString(doc), count };
}

function nextRelationshipId(relsDoc) {
  const rels = relsDoc.getElementsByTagNameNS(REL_NS, "Relationship");
  let max = 0;
  for (let i = 0; i < rels.length; i++) {
    const m = (rels[i].getAttribute("Id") || "").match(/^rId(\d+)$/i);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return `rId${max + 1}`;
}

function ensureCoreRelationship(relsXml) {
  const doc = new DOMParser().parseFromString(relsXml, "application/xml");
  const root = doc.documentElement;
  const rels = root.getElementsByTagNameNS(REL_NS, "Relationship");
  for (let i = 0; i < rels.length; i++) {
    if (rels[i].getAttribute("Type") === CORE_REL_TYPE) return relsXml;
  }
  const rel = doc.createElementNS(REL_NS, "Relationship");
  rel.setAttribute("Id", nextRelationshipId(doc));
  rel.setAttribute("Type", CORE_REL_TYPE);
  rel.setAttribute("Target", "docProps/core.xml");
  root.appendChild(rel);
  return new XMLSerializer().serializeToString(doc);
}

function ensureCoreContentType(ctXml) {
  const doc = new DOMParser().parseFromString(ctXml, "application/xml");
  const root = doc.documentElement;
  const overrides = root.getElementsByTagNameNS(CT_NS, "Override");
  for (let i = 0; i < overrides.length; i++) {
    if (overrides[i].getAttribute("PartName") === "/docProps/core.xml") return ctXml;
  }
  const override = doc.createElementNS(CT_NS, "Override");
  override.setAttribute("PartName", "/docProps/core.xml");
  override.setAttribute("ContentType", CORE_CONTENT_TYPE);
  root.appendChild(override);
  return new XMLSerializer().serializeToString(doc);
}

async function prepareCoreMetadataInZip(zip) {
  let coreXml;
  const coreFile = zip.file("docProps/core.xml");
  if (coreFile) coreXml = await coreFile.async("string");
  else coreXml = createDefaultCoreXml();

  const relsFile = zip.file("_rels/.rels");
  if (relsFile) {
    const relsXml = await relsFile.async("string");
    const nextRels = ensureCoreRelationship(relsXml);
    if (nextRels !== relsXml) zip.file("_rels/.rels", nextRels);
  }

  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    const ctXml = await ctFile.async("string");
    const nextCt = ensureCoreContentType(ctXml);
    if (nextCt !== ctXml) zip.file("[Content_Types].xml", nextCt);
  }

  return coreXml;
}
