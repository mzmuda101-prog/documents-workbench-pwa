// docx-forms.js — formularze Worda: kontrolki zawartości (w:sdt) i stare pola formularza
// (FORMTEXT / FORMCHECKBOX / FORMDROPDOWN). Odczyt, zapis i klikanie w polu w podglądzie.
//
// Podgląd (docx-preview) rozpakowuje kontrolki do zwykłego tekstu — nie wie, że to pola.
// Ten moduł czyta je z XML, a w podglądzie znajduje ich tekst po pozycji w akapicie
// (zakres DOM, bez zmiany treści), więc kliknięcie w pole otwiera listę / kalendarz /
// pole tekstowe, a pole wyboru przełącza się od razu. Wyróżnienie pól = CSS Highlight.
//
// Zapis (operacja „formFill”, wołana z buildPatchedDocx) zmienia TYLKO zawartość kontrolki —
// opakowanie, właściwości i lista zostają, więc w Wordzie formularz dalej działa. Pola
// powiązane z danymi (w:dataBinding) aktualizują też źródło (customXml / docProps): Word przy
// otwarciu odświeża kontrolki z tego źródła, a bez tego wróciłaby stara wartość.
//
// Klucze pól: „s<N>” = N-te w:sdt w document.xml, „f<N>” = N-te stare pole formularza.
// Akapity z polem są w Edycji tylko do odczytu (lockForm, docx-inline-edit.js) — zapis akapitu
// przepisuje jego fragmenty od nowa, więc kontrolka by się rozpadła.

const FORM_STORE_CORE = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}";
const FORM_STORE_APP = "{6668398D-A668-4E3E-A5EB-62B293D839F1}";
const FORM_DS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
const FORM_FILLABLE = new Set(["checkbox", "dropdown", "combo", "date", "text"]);
const FORM_LOCK_KINDS = new Set(["checkbox", "dropdown", "combo", "date", "text", "picture"]);
const FORM_SHADE_KEY = "dwb.forms.shade";

function ffKid(el, name, ns = W_NS) {
  if (!el) return null;
  for (let i = 0; i < el.childNodes.length; i++) {
    const n = el.childNodes[i];
    if (n.nodeType === 1 && n.localName === name && n.namespaceURI === ns) return n;
  }
  return null;
}
function ffKids(el, name, ns = W_NS) {
  return el ? Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.localName === name && n.namespaceURI === ns) : [];
}
function ffAttr(el, name, ns = W_NS) {
  if (!el) return "";
  return el.getAttributeNS(ns, name) ?? el.getAttribute(`${ns === W14_NS ? "w14" : "w"}:${name}`) ?? "";
}
function ffOn(el, ns = W_NS) { // <w:x/> albo <w:x w:val="1|true|on"/>
  if (!el) return false;
  const v = ffAttr(el, "val", ns);
  return !v || v === "1" || v === "true" || v === "on";
}
function ffClosest(el, name) {
  let n = el?.parentNode;
  while (n && n.nodeType === 1) {
    if (n.localName === name && n.namespaceURI === W_NS) return n;
    n = n.parentNode;
  }
  return null;
}
// Tekst w kolejności dokumentu (w:t, w:br → \n, w:tab → \t) — bez usuniętego (delText) i instrukcji pól.
function ffText(el) {
  let out = "";
  const walk = (n) => {
    if (n.nodeType !== 1) return;
    if (n.namespaceURI === W_NS && n.localName === "t") out += n.textContent || "";
    else if (n.namespaceURI === W_NS && (n.localName === "br" || n.localName === "cr")) out += "\n";
    else if (n.namespaceURI === W_NS && n.localName === "tab") out += "\t";
    else for (let i = 0; i < n.childNodes.length; i++) walk(n.childNodes[i]);
  };
  walk(el);
  return out;
}

function formSdtKind(sdtPr) {
  for (let i = 0; i < (sdtPr?.childNodes.length || 0); i++) {
    const n = sdtPr.childNodes[i];
    if (n.nodeType !== 1) continue;
    if (n.localName === "checkbox" && n.namespaceURI === W14_NS) return "checkbox";
    if (n.namespaceURI === W15_NS && /^repeatingSection/.test(n.localName)) return "repeat";
    if (n.namespaceURI !== W_NS) continue;
    const k = { dropDownList: "dropdown", comboBox: "combo", date: "date", text: "text", richText: "rich", picture: "picture",
      docPartObj: "docpart", docPartList: "docpart", group: "group", citation: "other", bibliography: "other", equation: "other" }[n.localName];
    if (k) return k;
  }
  return "rich"; // bez typu = bogaty tekst
}

// Poziom kontrolki: w akapicie (inline) czy wokół akapitów / wierszy / komórek.
function formSdtLevel(sdt) {
  const content = ffKid(sdt, "sdtContent");
  const kids = content ? Array.from(content.childNodes).filter((n) => n.nodeType === 1 && n.namespaceURI === W_NS) : [];
  if (kids.some((n) => n.localName === "tr")) return "row";
  if (kids.some((n) => n.localName === "tc")) return "cell";
  if (kids.some((n) => n.localName === "p" || n.localName === "tbl" || n.localName === "sdt" && ffKid(ffKid(n, "sdtContent"), "p"))) return "block";
  return "inline";
}

// Akapit w Edycji tylko do odczytu, bo zawiera pole formularza (albo cały jest polem).
function formParagraphLock(p) {
  if (p.getElementsByTagNameNS(W_NS, "sdt").length || p.getElementsByTagNameNS(W_NS, "ffData").length) return "lockForm";
  for (let s = ffClosest(p, "sdt"); s; s = ffClosest(s, "sdt")) {
    if (FORM_LOCK_KINDS.has(formSdtKind(ffKid(s, "sdtPr")))) return "lockForm";
  }
  return null;
}

// Edycja akapitu wewnątrz kontrolki „bogaty tekst” = kontrolka nie pokazuje już tekstu
// zastępczego (inaczej Word dalej traktowałby wpis jak szary podpowiadacz).
function clearSdtPlaceholderAround(p) {
  for (let s = ffClosest(p, "sdt"); s; s = ffClosest(s, "sdt")) {
    const ph = ffKid(ffKid(s, "sdtPr"), "showingPlcHdr");
    if (ph) ph.parentNode.removeChild(ph);
  }
}

function parsePrefixMappings(s) {
  const map = {};
  String(s || "").replace(/xmlns:(\w+)\s*=\s*['"]([^'"]*)['"]/g, (_, k, v) => { map[k] = v; return ""; });
  return map;
}

// Tekst akapitu przed / po elemencie (do etykiety pola i do znalezienia go w podglądzie).
function formTextAround(p, el) {
  let before = "";
  let after = "";
  paragraphTextNodes(p).forEach((t) => {
    if (el.contains(t)) return;
    if (el.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_PRECEDING) before += t.textContent || "";
    else after += t.textContent || "";
  });
  return { before, after };
}

function formShortLabel(before, after, kind) {
  const clean = (s) => s.replace(/\s+/g, " ").trim();
  const b = clean(before).replace(/[:\s(–—-]+$/, "");
  const a = clean(after).replace(/^[\s:)–—-]+/, "");
  const pick = kind === "checkbox" ? a || b : b || a;
  if (!pick) return "";
  if (kind === "checkbox") return pick.length > 60 ? `${pick.slice(0, 59)}…` : pick;
  return pick.length > 60 ? `…${pick.slice(-59)}` : pick;
}

function readSdtField(sdt, key, paraIndexOf) {
  const sdtPr = ffKid(sdt, "sdtPr");
  const kind = formSdtKind(sdtPr);
  const level = formSdtLevel(sdt);
  const content = ffKid(sdt, "sdtContent");
  if (!content) return null;
  // bogaty tekst wewnątrz akapitu = zwykłe pole tekstowe; blok bogatego tekstu edytuje się w dokumencie
  const fillKind = kind === "rich" && level === "inline" ? "text" : kind;
  if (!FORM_FILLABLE.has(fillKind) || level === "row" || level === "cell") return null;
  if (fillKind !== "text" && content.getElementsByTagNameNS(W_NS, "tbl").length) return null;
  if (level === "block" && fillKind === "text" && ffKids(content, "p").length > 1) return null;
  const p = level === "inline" ? ffClosest(sdt, "p") : content.getElementsByTagNameNS(W_NS, "p")[0];
  const around = p && level === "inline" ? formTextAround(p, sdt) : { before: "", after: "" };
  const lock = ffAttr(ffKid(sdtPr, "lock"), "val");
  const display = ffText(content);
  const f = {
    key, source: "sdt", kind: fillKind, level,
    id: ffAttr(ffKid(sdtPr, "id"), "val"),
    alias: ffAttr(ffKid(sdtPr, "alias"), "val"),
    tag: ffAttr(ffKid(sdtPr, "tag"), "val"),
    placeholder: !!ffKid(sdtPr, "showingPlcHdr"),
    locked: lock === "contentLocked" || lock === "sdtContentLocked",
    display, value: display, options: [],
    before: around.before, after: around.after,
    paraIndex: p && paraIndexOf.has(p) ? paraIndexOf.get(p) : null,
  };
  if (f.kind === "checkbox") {
    const cb = ffKid(sdtPr, "checkbox", W14_NS);
    const state = (name, def) => {
      const el = ffKid(cb, name, W14_NS);
      const hex = ffAttr(el, "val", W14_NS);
      return { char: hex ? String.fromCodePoint(parseInt(hex, 16)) : def, font: ffAttr(el, "font", W14_NS) };
    };
    f.value = ffOn(ffKid(cb, "checked", W14_NS), W14_NS);
    f.on = state("checkedState", "☒");
    f.off = state("uncheckedState", "☐");
  } else if (f.kind === "dropdown" || f.kind === "combo") {
    const list = ffKid(sdtPr, f.kind === "dropdown" ? "dropDownList" : "comboBox");
    f.options = ffKids(list, "listItem").map((li) => {
      const value = ffAttr(li, "value");
      return { text: ffAttr(li, "displayText") || value, value };
    });
    const hit = f.options.find((o) => o.text === display);
    f.value = f.placeholder ? "" : hit ? hit.value : f.kind === "combo" ? display : ffAttr(list, "lastValue");
  } else if (f.kind === "date") {
    const date = ffKid(sdtPr, "date");
    f.value = f.placeholder ? "" : (ffAttr(date, "fullDate").match(/^\d{4}-\d{2}-\d{2}/) || [""])[0];
    f.dateFormat = ffAttr(ffKid(date, "dateFormat"), "val");
    f.lid = ffAttr(ffKid(date, "lid"), "val");
    f.storeAs = ffAttr(ffKid(date, "storeMappedDataAs"), "val");
  } else if (f.kind === "text") {
    f.multiLine = ["1", "true", "on"].includes(ffAttr(ffKid(sdtPr, "text"), "multiLine"));
    if (f.placeholder) f.value = "";
  }
  const bind = ffKid(sdtPr, "dataBinding");
  if (bind) {
    f.binding = { store: ffAttr(bind, "storeItemID"), xpath: ffAttr(bind, "xpath"), prefix: ffAttr(bind, "prefixMappings") };
    f.bindKey = `${f.binding.store.toUpperCase()}|${f.binding.xpath}`;
  }
  f.label = f.alias || formShortLabel(f.before, f.after, f.kind) || f.tag || "";
  return f;
}

// Stare pola formularza: <w:fldChar begin><w:ffData>… + instrukcja + [separate + wynik] + end.
function collectLegacyFormFields(doc) {
  const out = [];
  Array.from(doc.getElementsByTagNameNS(W_NS, "fldChar")).forEach((fc) => {
    if (ffAttr(fc, "fldCharType") !== "begin") return;
    const ffData = ffKid(fc, "ffData");
    if (!ffData) return;
    const beginRun = fc.parentNode;
    let sepRun = null;
    let endRun = null;
    let depth = 0;
    for (let n = beginRun.nextSibling; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const c = n.localName === "r" ? ffKid(n, "fldChar") : null;
      if (!c) continue;
      const type = ffAttr(c, "fldCharType");
      if (type === "begin") depth++;
      else if (type === "separate" && !depth) sepRun = n;
      else if (type === "end") { if (!depth) { endRun = n; break; } depth--; }
    }
    if (!endRun) return;
    out.push({ ffData, beginRun, sepRun, endRun });
  });
  return out;
}

function legacyResultRuns(lf) {
  const runs = [];
  if (!lf.sepRun) return runs;
  for (let n = lf.sepRun.nextSibling; n && n !== lf.endRun; n = n.nextSibling) if (n.nodeType === 1) runs.push(n);
  return runs;
}

function readLegacyField(lf, key, paraIndexOf) {
  const { ffData } = lf;
  const p = ffClosest(lf.beginRun, "p");
  const textInput = ffKid(ffData, "textInput");
  const checkBox = ffKid(ffData, "checkBox");
  const ddList = ffKid(ffData, "ddList");
  const kind = checkBox ? "checkbox" : ddList ? "dropdown" : textInput ? "text" : null;
  if (!kind) return null;
  // tekst przed/po: od początku pola do końca (instrukcja nie jest tekstem, wynik tak)
  let before = "";
  let after = "";
  if (p) {
    const nodes = paragraphTextNodes(p);
    nodes.forEach((t) => {
      if (lf.beginRun.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_PRECEDING) before += t.textContent || "";
      else if (lf.endRun.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_FOLLOWING) after += t.textContent || "";
    });
  }
  const display = legacyResultRuns(lf).map(ffText).join("");
  const enabled = ffKid(ffData, "enabled");
  const f = {
    key, source: "legacy", kind, level: "inline",
    name: ffAttr(ffKid(ffData, "name"), "val"),
    locked: !!enabled && !ffOn(enabled),
    placeholder: false, display, value: display, options: [], before, after,
    paraIndex: p && paraIndexOf.has(p) ? paraIndexOf.get(p) : null,
  };
  if (kind === "checkbox") {
    const checked = ffKid(checkBox, "checked"); // brak „checked” = stan domyślny
    f.value = ffOn(checked || ffKid(checkBox, "default"));
    f.on = { char: "☒" };
    f.off = { char: "☐" };
    f.display = "";
  } else if (kind === "dropdown") {
    f.options = ffKids(ddList, "listEntry").map((e) => ({ text: ffAttr(e, "val"), value: ffAttr(e, "val") }));
    const idxRaw = ffAttr(ffKid(ddList, "result"), "val") || ffAttr(ffKid(ddList, "default"), "val") || "0";
    f.value = f.options[parseInt(idxRaw, 10)]?.value ?? f.options[0]?.value ?? "";
  } else {
    f.maxLength = parseInt(ffAttr(ffKid(textInput, "maxLength"), "val") || "0", 10) || 0;
  }
  f.label = formShortLabel(before, after, kind) || f.name || "";
  return f;
}

// ── odczyt całego dokumentu ──────────────────────────────────────────────────
function scanFormFieldsInDoc(doc) {
  const paraIndexOf = new Map();
  collectParagraphElements(doc.documentElement, "all").forEach((p, i) => paraIndexOf.set(p, i));
  const fields = [];
  const sdts = Array.from(doc.getElementsByTagNameNS(W_NS, "sdt"));
  const taken = new Set();
  sdts.forEach((sdt, i) => {
    // pole wewnątrz już wziętego pola (np. lista w polu tekstowym) — nie osobno
    for (let s = ffClosest(sdt, "sdt"); s; s = ffClosest(s, "sdt")) if (taken.has(s)) return;
    const f = readSdtField(sdt, `s${i}`, paraIndexOf);
    if (!f) return;
    taken.add(sdt);
    fields.push(f);
  });
  collectLegacyFormFields(doc).forEach((lf, i) => {
    const f = readLegacyField(lf, `f${i}`, paraIndexOf);
    if (f) fields.push(f);
  });
  // kolejność jak w dokumencie
  fields.sort((a, b) => (a.paraIndex ?? 1e9) - (b.paraIndex ?? 1e9) || (a.before.length - b.before.length));
  // powiązane pola: jedno źródło danych → pierwsze jest „główne”, reszta to kopie
  const firstByBind = new Map();
  fields.forEach((f) => {
    if (!f.bindKey) return;
    const first = firstByBind.get(f.bindKey);
    if (first) { f.copyOf = first.key; first.copies = (first.copies || 1) + 1; } else firstByBind.set(f.bindKey, f);
  });
  return fields;
}

async function scanFormFields(bytes) {
  const doc = await getDocumentXmlDom(bytes);
  if (!doc) return { fields: [], protection: null };
  const fields = scanFormFieldsInDoc(doc);
  let protection = null;
  try {
    const zip = await loadDocxZipCached(bytes);
    const s = zip.file("word/settings.xml");
    if (s) {
      const sd = new DOMParser().parseFromString(await s.async("string"), "application/xml");
      const dp = sd.getElementsByTagNameNS(W_NS, "documentProtection")[0];
      const enf = ffAttr(dp, "enforcement");
      if (dp && (enf === "1" || enf === "true" || enf === "on")) protection = ffAttr(dp, "edit") || "readOnly";
    }
  } catch (_) { /* brak ustawień — bez ochrony */ }
  return { fields, protection };
}

// ── daty w formacie Worda (d MMMM yyyy, dd.MM.yyyy, dddd…) ───────────────────
function formatWordDate(ymd, fmt, lid) {
  const m = String(ymd || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return "";
  const date = new Date(+m[1], +m[2] - 1, +m[3]);
  const loc = lid || (typeof I18N !== "undefined" && I18N[currentLang]?.locale) || "pl-PL";
  const f = fmt || "dd.MM.yyyy";
  const hasDay = /d/.test(f.replace(/'[^']*'|"[^"]*"/g, ""));
  const part = (opts, type) => {
    try { return new Intl.DateTimeFormat(loc, opts).formatToParts(date).find((x) => x.type === type)?.value || ""; } catch (_) { return ""; }
  };
  const pad = (n) => String(n).padStart(2, "0");
  return f.replace(/'[^']*'|"[^"]*"|dddd|ddd|dd|d|MMMM|MMM|MM|M|yyyy|yy|HH|H|hh|h|mm|ss|AM\/PM|am\/pm/g, (tok) => {
    switch (tok) {
      case "dddd": return part({ weekday: "long" }, "weekday");
      case "ddd": return part({ weekday: "short" }, "weekday");
      case "dd": return pad(date.getDate());
      case "d": return String(date.getDate());
      // „15 września” (dopełniacz) przy dniu, „wrzesień” bez dnia — jak w Wordzie
      case "MMMM": return hasDay ? part({ day: "numeric", month: "long" }, "month") : part({ month: "long" }, "month");
      case "MMM": return hasDay ? part({ day: "numeric", month: "short" }, "month") : part({ month: "short" }, "month");
      case "MM": return pad(date.getMonth() + 1);
      case "M": return String(date.getMonth() + 1);
      case "yyyy": return String(date.getFullYear());
      case "yy": return pad(date.getFullYear() % 100);
      case "HH": case "hh": case "mm": case "ss": return "00";
      case "H": case "h": return "0";
      case "AM/PM": case "am/pm": return "";
      default: return tok.slice(1, -1); // tekst w cudzysłowie
    }
  });
}

// ── zapis ────────────────────────────────────────────────────────────────────
function ffSetVal(el, ns, value) {
  el.setAttributeNS(ns, ns === W14_NS ? "w14:val" : "w:val", String(value));
}

function ffRun(doc, rPr, text) {
  const r = doc.createElementNS(W_NS, "w:r");
  if (rPr) r.appendChild(rPr);
  String(text).split("\n").forEach((line, i) => {
    if (i) r.appendChild(doc.createElementNS(W_NS, "w:br"));
    if (!line) return;
    const t = doc.createElementNS(W_NS, "w:t");
    t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    t.textContent = line;
    r.appendChild(t);
  });
  return r;
}

function ffRunProps(rPr, { dropStyle, font }) {
  if (!rPr) return null;
  if (dropStyle) ffKids(rPr, "rStyle").forEach((s) => rPr.removeChild(s)); // styl „Tekst zastępczy” (szary)
  if (font) {
    let rf = ffKid(rPr, "rFonts");
    if (!rf) {
      rf = rPr.ownerDocument.createElementNS(W_NS, "w:rFonts");
      const style = ffKid(rPr, "rStyle");
      rPr.insertBefore(rf, style ? style.nextSibling : rPr.firstChild);
    }
    ["ascii", "hAnsi", "eastAsia", "cs"].forEach((a) => rf.setAttributeNS(W_NS, `w:${a}`, font));
  }
  return rPr.childNodes.length ? rPr : null;
}

// Nowa treść kontrolki: jeden fragment z formatem pierwszego fragmentu (bez stylu tekstu zastępczego).
function setSdtContentText(sdt, text, opts = {}) {
  const sdtPr = ffKid(sdt, "sdtPr");
  const content = ffKid(sdt, "sdtContent");
  if (!content) return;
  const wasPh = !!ffKid(sdtPr, "showingPlcHdr");
  const paras = ffKids(content, "p");
  const target = paras[0] || content;
  paras.slice(1).forEach((p) => content.removeChild(p));
  const runs = Array.from(target.getElementsByTagNameNS(W_NS, "r"));
  const rPr = ffRunProps(runs[0] && ffKid(runs[0], "rPr")?.cloneNode(true), { dropStyle: wasPh, font: opts.font });
  runs.forEach((r) => r.parentNode.removeChild(r));
  if (text !== "") target.appendChild(ffRun(sdt.ownerDocument, rPr, text));
  if (wasPh) sdtPr.removeChild(ffKid(sdtPr, "showingPlcHdr"));
}

function fillSdtField(sdt, f, val) {
  const sdtPr = ffKid(sdt, "sdtPr");
  if (f.kind === "checkbox") {
    const on = val === true || val === "1" || val === "true";
    const cb = ffKid(sdtPr, "checkbox", W14_NS);
    let checked = ffKid(cb, "checked", W14_NS);
    if (!checked) { checked = sdt.ownerDocument.createElementNS(W14_NS, "w14:checked"); cb.insertBefore(checked, cb.firstChild); }
    ffSetVal(checked, W14_NS, on ? 1 : 0);
    const state = on ? f.on : f.off;
    setSdtContentText(sdt, state.char, { font: state.font });
    return on ? "true" : "false";
  }
  if (f.kind === "dropdown" || f.kind === "combo") {
    const s = String(val ?? "");
    const opt = f.options.find((o) => o.value === s) || f.options.find((o) => o.text === s);
    if (f.kind === "dropdown" && !opt) return null;
    const list = ffKid(sdtPr, f.kind === "dropdown" ? "dropDownList" : "comboBox");
    list?.setAttributeNS(W_NS, "w:lastValue", opt ? opt.value : s);
    setSdtContentText(sdt, opt ? opt.text : s);
    return opt ? opt.value : s;
  }
  if (f.kind === "date") {
    const ymd = (String(val || "").match(/^\d{4}-\d{2}-\d{2}/) || [""])[0];
    if (!ymd) return null;
    ffKid(sdtPr, "date")?.setAttributeNS(W_NS, "w:fullDate", `${ymd}T00:00:00Z`);
    const shown = formatWordDate(ymd, f.dateFormat, f.lid);
    setSdtContentText(sdt, shown);
    return f.storeAs === "text" ? shown : f.storeAs === "date" ? ymd : `${ymd}T00:00:00`;
  }
  let s = sanitizeXmlText(String(val ?? ""));
  if (!f.multiLine) s = s.replace(/\s*\n\s*/g, " ");
  setSdtContentText(sdt, s);
  return s;
}

function fillLegacyField(lf, f, val) {
  const doc = lf.ffData.ownerDocument;
  if (f.kind === "checkbox") {
    const cb = ffKid(lf.ffData, "checkBox");
    let checked = ffKid(cb, "checked");
    if (!checked) { checked = doc.createElementNS(W_NS, "w:checked"); cb.appendChild(checked); }
    ffSetVal(checked, W_NS, val === true || val === "1" || val === "true" ? 1 : 0);
    return;
  }
  let text = String(val ?? "");
  if (f.kind === "dropdown") {
    const dd = ffKid(lf.ffData, "ddList");
    const idx = f.options.findIndex((o) => o.value === text);
    if (idx < 0) return;
    let res = ffKid(dd, "result");
    if (!res) { res = doc.createElementNS(W_NS, "w:result"); dd.insertBefore(res, dd.firstChild); }
    ffSetVal(res, W_NS, idx);
    if (!lf.sepRun) return; // Word rysuje wybraną pozycję sam
  } else {
    text = sanitizeXmlText(text).replace(/\s*\n\s*/g, " ");
    if (f.maxLength) text = text.slice(0, f.maxLength);
  }
  const old = legacyResultRuns(lf);
  const firstRun = old.find((n) => n.localName === "r");
  const rPr = (firstRun && ffKid(firstRun, "rPr")) || ffKid(lf.beginRun, "rPr");
  old.forEach((n) => n.parentNode.removeChild(n));
  if (!lf.sepRun) {
    lf.sepRun = doc.createElementNS(W_NS, "w:r");
    if (rPr) lf.sepRun.appendChild(rPr.cloneNode(true));
    const sep = doc.createElementNS(W_NS, "w:fldChar");
    sep.setAttributeNS(W_NS, "w:fldCharType", "separate");
    lf.sepRun.appendChild(sep);
    lf.endRun.parentNode.insertBefore(lf.sepRun, lf.endRun);
  }
  if (text) lf.endRun.parentNode.insertBefore(ffRun(doc, rPr ? rPr.cloneNode(true) : null, text), lf.endRun);
}

// Źródło danych pola powiązanego (customXml/itemN.xml albo docProps/*.xml).
async function writeFormBinding(zip, binding, raw) {
  const store = String(binding.store || "").toUpperCase();
  let part = store === FORM_STORE_CORE ? "docProps/core.xml" : store === FORM_STORE_APP ? "docProps/app.xml" : null;
  if (!part) {
    for (const name of Object.keys(zip.files).filter((n) => /^customXml\/itemProps\d+\.xml$/i.test(n))) {
      const d = new DOMParser().parseFromString(await zip.file(name).async("string"), "application/xml");
      const id = d.documentElement.getAttributeNS(FORM_DS_NS, "itemID") || d.documentElement.getAttribute("ds:itemID") || "";
      if (id.toUpperCase() === store) { part = name.replace(/itemProps(\d+)\.xml$/i, "item$1.xml"); break; }
    }
  }
  const file = part && zip.file(part);
  if (!file) return false;
  const xdoc = new DOMParser().parseFromString(await file.async("string"), "application/xml");
  const ns = parsePrefixMappings(binding.prefix);
  let node = null;
  try {
    node = xdoc.evaluate(binding.xpath, xdoc, (pfx) => ns[pfx] || null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
  } catch (_) { return false; }
  if (!node) return false;
  if (node.nodeType === 2) node.value = raw; else node.textContent = raw;
  zip.file(part, new XMLSerializer().serializeToString(xdoc));
  return true;
}

// Operacja „formFill” (buildPatchedDocx): edit.values = { klucz: wartość }.
// Wartości: pole wyboru true/false, lista — wartość pozycji, data — RRRR-MM-DD, tekst — tekst.
async function applyFormFillInZip(zip, xml, edit) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const paraIndexOf = new Map();
  const sdts = Array.from(doc.getElementsByTagNameNS(W_NS, "sdt"));
  const legacy = collectLegacyFormFields(doc);
  const before = new XMLSerializer().serializeToString(doc);
  const bound = new Map(); // bindKey → { f, raw, done:Set<sdt> }
  Object.entries(edit.values || {}).forEach(([key, val]) => {
    const n = parseInt(key.slice(1), 10);
    if (key[0] === "s" && sdts[n]) {
      const f = readSdtField(sdts[n], key, paraIndexOf);
      if (!f || f.locked) return;
      const raw = fillSdtField(sdts[n], f, val);
      if (raw == null || !f.bindKey) return;
      const entry = bound.get(f.bindKey) || { f, val, raw, done: new Set() };
      entry.done.add(sdts[n]);
      bound.set(f.bindKey, entry);
    } else if (key[0] === "f" && legacy[n]) {
      const f = readLegacyField(legacy[n], key, paraIndexOf);
      if (f && !f.locked) fillLegacyField(legacy[n], f, val);
    }
  });
  // powiązane: ta sama wartość we wszystkich kopiach + w źródle danych
  for (const [bindKey, entry] of bound) {
    const copies = sdts.filter((s, i) => {
      if (entry.done.has(s)) return false;
      const g = readSdtField(s, `s${i}`, paraIndexOf);
      return g?.bindKey === bindKey;
    });
    copies.forEach((s) => {
      const g = readSdtField(s, "", paraIndexOf);
      if (g && g.kind === entry.f.kind) fillSdtField(s, g, entry.val);
    });
    let ok = false;
    try { ok = await writeFormBinding(zip, entry.f.binding, entry.raw); } catch (_) { ok = false; }
    if (!ok) { // źródła nie ma / nie da się go zmienić — odpinamy, żeby Word nie przywrócił starej wartości
      [...entry.done, ...copies].forEach((s) => {
        const b = ffKid(ffKid(s, "sdtPr"), "dataBinding");
        if (b) b.parentNode.removeChild(b);
      });
    }
  }
  const after = new XMLSerializer().serializeToString(doc);
  if (after === before) return { xml, count: 0 };
  return { xml: after, count: Object.keys(edit.values || {}).length };
}

// ── stan dla UI ──────────────────────────────────────────────────────────────
let formScan = null; // { bytes, fields, protection }
let docFormCounts = { fields: 0, empty: 0 };
let formNoticePending = false;
const formUi = { byPara: new Map(), ranges: new Map(), pop: null, panelBytes: null };

function formMainFields() {
  return (formScan?.fields || []).filter((f) => !f.copyOf);
}
function formIsEmpty(f) {
  return f.kind !== "checkbox" && (f.placeholder || !String(f.value ?? "").trim());
}

async function refreshFormScan() {
  const bytes = originalFileBytes;
  if (!bytes || !window.JSZip || currentFileType !== "docx") {
    formScan = null;
  } else if (formScan?.bytes !== bytes) {
    try { formScan = { bytes, ...(await scanFormFields(bytes)) }; } catch (_) { formScan = { bytes, fields: [], protection: null }; }
    if (bytes !== originalFileBytes) return; // w międzyczasie nowa wersja — zrobi to kolejne wywołanie
  }
  const main = formMainFields();
  docFormCounts = { fields: main.length, empty: main.filter(formIsEmpty).length };
  paintFormFields();
  if (typeof appFrame !== "undefined") appFrame.syncPanelCounts();
  // panel tylko przy NOWEJ wersji pliku — przebudowa podmienia pola w trakcie wpisywania/zmiany
  if (document.getElementById("panel-forms")?.open && typeof renderFormsPanel === "function" && formUi.panelBytes !== formScan?.bytes) renderFormsPanel();
  if (formNoticePending) {
    formNoticePending = false;
    if (docFormCounts.fields) toast(t("formsOnOpen", { count: docFormCounts.fields }), "info");
  }
}

// ── pola w podglądzie: „kafelki” do kliknięcia ───────────────────────────────
// Tekst pola dostaje w podglądzie opakowanie <span class="ff-field ff-<rodzaj>"> (ramka, znaczek
// rodzaju — ▾ lista, kalendarz, przerywana ramka tekstu, kłódka). Treść się NIE zmienia: akapit
// z polem jest tylko do odczytu, a zapis z podglądu pomija takie akapity — więc nic z tego nie
// trafia do pliku. Znaczki są z CSS (::after), więc nie wchodzą też do tekstu (szukanie, eksport).
const FF_HINT = {
  checkbox: ["Pole wyboru — kliknij, żeby zaznaczyć lub odznaczyć", "Check box — click to tick or untick"],
  dropdown: ["Lista — kliknij, żeby wybrać", "List — click to choose"],
  combo: ["Lista lub własny tekst — kliknij", "List or your own text — click"],
  date: ["Data — kliknij, żeby wybrać z kalendarza", "Date — click to pick from a calendar"],
  text: ["Pole tekstowe — kliknij, żeby wpisać", "Text field — click to type"],
  locked: ["Pole zablokowane w Wordzie", "Field locked in Word"],
};

function formShadeOn() {
  try { return localStorage.getItem(FORM_SHADE_KEY) !== "0"; } catch (_) { return true; }
}
function setFormShade(on) {
  try { localStorage.setItem(FORM_SHADE_KEY, on ? "1" : "0"); } catch (_) { /* tryb prywatny */ }
  paintFormFields();
}

// Zakres DOM tekstu [start, end) w elemencie (węzły tekstowe po kolei).
function formDomRange(el, start, end) {
  const range = document.createRange();
  let pos = 0;
  let startSet = false;
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const len = n.textContent.length;
    if (!startSet && start <= pos + len) { range.setStart(n, start - pos); startSet = true; }
    if (startSet && end <= pos + len) { range.setEnd(n, end - pos); return range; }
    pos += len;
  }
  return null;
}

// Gdzie w podglądzie jest tekst pola (null = pole bez widocznego tekstu).
function formPreviewRange(p, f) {
  if (f.level === "block") {
    const range = document.createRange();
    range.selectNodeContents(p);
    return range;
  }
  if (!f.display || (f.source === "legacy" && f.kind === "checkbox")) return null;
  const text = p.textContent || "";
  let at = text.indexOf(f.display, Math.max(0, f.before.length - 4));
  if (at < 0 || Math.abs(at - f.before.length) > 12) at = text.indexOf(f.display);
  return at >= 0 ? formDomRange(p, at, at + f.display.length) : null;
}

function formWrap(p, f) {
  const w = document.createElement("span");
  w.dataset.ff = f.key;
  const range = formPreviewRange(p, f);
  if (range) {
    try {
      w.appendChild(range.extractContents()); // fragmenty z różnym formatem zostają w środku
      range.insertNode(w);
      return w;
    } catch (_) { /* zakres przez granice elementów, których nie da się rozciąć — znaczek obok */ }
  }
  // stare pole wyboru / lista bez wyniku w pliku / pusta kontrolka — znaczek w miejscu pola
  w.classList.add("ff-glyph");
  const at = formDomRange(p, f.before.length, f.before.length);
  if (at) at.insertNode(w); else p.prepend(w);
  return w;
}

function formDecorate(w, f) {
  const glyph = w.classList.contains("ff-glyph");
  w.className = `ff-field ff-${f.kind}`;
  w.classList.toggle("ff-glyph", glyph);
  w.classList.toggle("is-block", f.level === "block");
  w.classList.toggle("is-placeholder", !!f.placeholder);
  w.classList.toggle("is-locked", !!f.locked);
  if (glyph) {
    // Word rysuje stare pole wyboru i wybraną pozycję starej listy sam — pokazujemy je tu
    const shown = f.kind === "checkbox" ? (f.value ? f.on.char : f.off.char) : f.kind === "dropdown" && f.value ? f.value : "";
    w.classList.toggle("is-empty", !shown);
    w.textContent = shown || "  ";
  }
  w.tabIndex = 0;
  w.setAttribute("role", f.kind === "checkbox" ? "checkbox" : "button");
  if (f.kind === "checkbox") w.setAttribute("aria-checked", String(!!f.value));
  else w.setAttribute("aria-haspopup", "dialog");
  const shownText = f.kind === "checkbox" ? "" : f.placeholder ? "" : `: ${f.kind === "date" || f.kind === "text" || f.kind === "combo" ? f.display : f.options.find((o) => o.value === f.value)?.text || f.display}`;
  w.setAttribute("aria-label", `${formFieldName(f)}${shownText}`);
  const [pl, en] = FF_HINT[f.locked ? "locked" : f.kind];
  w.dataset.hint = "";
  w.dataset.hintPl = pl;
  w.dataset.hintEn = en;
}

function paintFormFields() {
  formUi.byPara = new Map();
  formUi.ranges = new Map();
  docCanvasEl?.classList.toggle("ff-plain", !formShadeOn());
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host || !formScan?.fields.length || formScan.bytes !== originalFileBytes) return;
  const previews = collectPreviewParagraphElements(host);
  formScan.fields.forEach((f) => {
    const p = Number.isFinite(f.paraIndex) ? previews[f.paraIndex] : null;
    if (!p) return;
    // ten sam podgląd (np. kolejne odświeżenie panelu) — opakowanie już jest
    const w = p.querySelector(`.ff-field[data-ff="${f.key}"]`) || formWrap(p, f);
    formDecorate(w, f);
    const range = document.createRange();
    range.selectNodeContents(w);
    formUi.ranges.set(f.key, range);
    if (!formUi.byPara.has(p)) formUi.byPara.set(p, []);
    formUi.byPara.get(p).push(f);
    p.classList.add("ff-para");
  });
}

// Dotyk: palec trafia obok małego pola (☐) — liczymy też kilka px wokół.
function formFieldAt(p, x, y) {
  const list = formUi.byPara.get(p);
  if (!list) return null;
  const pad = window.matchMedia?.("(pointer: coarse)").matches ? 10 : 3;
  for (const f of list) {
    const range = formUi.ranges.get(f.key);
    if (!range) continue;
    for (const r of range.getClientRects()) {
      if (x >= r.left - pad && x <= r.right + pad && y >= r.top - pad && y <= r.bottom + pad) return f;
    }
  }
  return null;
}

// ── wypełnianie ──────────────────────────────────────────────────────────────
// Zmiany idą po kolei (szybkie klikanie kilku pól). `values` może być funkcją — liczy wartości
// dopiero w swojej kolejce, ze świeżego odczytu (dwa szybkie kliknięcia w ☐ = zaznacz, odznacz).
let formFillQueue = Promise.resolve();
function fillFormFields(values) {
  const job = formFillQueue.then(async () => {
    if (!originalFileBytes) return 0;
    await refreshFormScan();
    const vals = typeof values === "function" ? values() : values;
    if (!vals) return 0;
    const top = docViewportEl?.scrollTop || 0;
    quietRenderOnce = true;
    const n = await applyDocumentEdit({ op: "formFill", values: vals });
    quietRenderOnce = false;
    if (docViewportEl) docViewportEl.scrollTop = top; // klik w pole w połowie dokumentu — nie skacz na górę
    await refreshFormScan();
    return n;
  });
  formFillQueue = job.catch(() => 0);
  return job;
}

function toggleFormCheckbox(key) {
  return fillFormFields(() => {
    const f = formScan?.fields.find((x) => x.key === key);
    return f && !f.locked ? { [key]: !f.value } : null;
  });
}

function formFieldName(f) {
  return f.label || t("formsUnnamed");
}

function closeFormPop() {
  formUi.pop?.remove();
  formUi.pop = null;
}

function openFormPop(f, anchorRect) {
  closeFormPop();
  const pop = document.createElement("div");
  pop.className = "ff-pop";
  pop.setAttribute("role", "dialog");
  pop.setAttribute("aria-label", formFieldName(f));
  const title = document.createElement("div");
  title.className = "ff-pop-title";
  title.textContent = formFieldName(f);
  pop.append(title);
  const commit = async (val) => {
    closeFormPop();
    await fillFormFields({ [f.key]: val });
  };
  const btn = (text, cls, onClick) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = cls;
    b.textContent = text;
    b.addEventListener("click", onClick);
    return b;
  };
  if (f.locked) {
    const p = document.createElement("p");
    p.className = "hint";
    p.textContent = t("formsLocked");
    pop.append(p);
  } else if (f.kind === "dropdown" || (f.kind === "combo" && f.options.length)) {
    const list = document.createElement("div");
    list.className = "ff-pop-list";
    list.setAttribute("role", "listbox");
    f.options.forEach((o) => {
      const b = btn(o.text, `ff-pop-opt${o.value === f.value && !f.placeholder ? " is-current" : ""}`, () => commit(o.value));
      b.setAttribute("role", "option");
      b.setAttribute("aria-selected", String(o.value === f.value));
      list.append(b);
    });
    pop.append(list);
  }
  if (!f.locked && (f.kind === "combo" || f.kind === "text" || f.kind === "date")) {
    const row = document.createElement("form");
    row.className = "ff-pop-row";
    const input = document.createElement(f.multiLine ? "textarea" : "input");
    if (f.kind === "date") input.type = "date";
    else if (!f.multiLine) input.type = "text";
    input.className = "ff-pop-input";
    input.value = f.kind === "date" ? f.value || "" : f.placeholder ? "" : f.kind === "combo" ? f.display : f.value;
    if (f.maxLength) input.maxLength = f.maxLength;
    if (f.kind !== "date") input.placeholder = f.placeholder ? f.display : "";
    input.setAttribute("aria-label", formFieldName(f));
    row.append(input);
    if (f.kind === "date") {
      row.append(btn(t("formsToday"), "btn", () => {
        const d = new Date();
        commit(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`);
      }));
    }
    const ok = btn("OK", "btn primary", () => {});
    ok.type = "submit";
    row.append(ok);
    row.addEventListener("submit", (e) => { e.preventDefault(); commit(input.value); });
    if (f.multiLine) input.addEventListener("keydown", (e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); commit(input.value); } });
    pop.append(row);
  }
  document.body.append(pop);
  formUi.pop = pop;
  // pod polem, a gdy brak miejsca — nad nim
  const w = Math.min(320, window.innerWidth - 16);
  pop.style.width = `${w}px`;
  const left = Math.max(8, Math.min(anchorRect.left, window.innerWidth - w - 8));
  const h = pop.offsetHeight;
  const below = anchorRect.bottom + 6;
  const top = below + h > window.innerHeight - 8 && anchorRect.top - h - 6 > 8 ? anchorRect.top - h - 6 : Math.min(below, window.innerHeight - h - 8);
  pop.style.left = `${left}px`;
  pop.style.top = `${Math.max(8, top)}px`;
  const first = pop.querySelector(".ff-pop-opt.is-current, .ff-pop-input, .ff-pop-opt, button");
  // dotyk: bez automatycznej klawiatury przy liście — tylko przy polu tekstowym
  first?.focus({ preventScroll: true });
}

function activateFormField(f, anchorRect) {
  if (f.kind === "checkbox" && !f.locked) { toggleFormCheckbox(f.key); return; }
  openFormPop(f, anchorRect);
}

function onFormDocClick(e) {
  if (!formUi.byPara.size || e.button > 0) return;
  const sel = window.getSelection();
  if (sel && !sel.isCollapsed && docCanvasEl.contains(sel.anchorNode)) return; // zaznaczanie tekstu
  const w = e.target.closest?.(".ff-field");
  const p = e.target.closest?.(".ff-para");
  const f = w ? formScan?.fields.find((x) => x.key === w.dataset.ff) : p ? formFieldAt(p, e.clientX, e.clientY) : null;
  if (!f) return;
  e.preventDefault();
  const range = formUi.ranges.get(f.key);
  const rects = range ? [...range.getClientRects()] : [];
  const rect = rects.find((r) => e.clientY >= r.top - 12 && e.clientY <= r.bottom + 12) || rects[0] || p.getBoundingClientRect();
  activateFormField(f, rect);
}

// Klawiatura: Tab dochodzi do pola, Enter / spacja = jak kliknięcie.
function onFormDocKeydown(e) {
  const w = e.target.closest?.(".ff-field");
  if (!w || (e.key !== "Enter" && e.key !== " ") || e.ctrlKey || e.metaKey || e.altKey) return;
  const f = formScan?.fields.find((x) => x.key === w.dataset.ff);
  if (!f) return;
  e.preventDefault();
  activateFormField(f, w.getBoundingClientRect());
}

// Skok do pola z panelu: przewiń, mrugnij polem, otwórz okienko.
function jumpToFormField(key, opts = {}) {
  const f = formScan?.fields.find((x) => x.key === key);
  const range = formUi.ranges.get(key);
  if (!f || !range) return;
  const w = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
  const p = w?.closest?.("p");
  if (p) jumpToStructureItem({ el: p, id: "form" }, { silentSelect: true });
  if (w?.classList.contains("ff-field")) {
    w.classList.remove("is-flash");
    void w.offsetWidth; // od nowa, gdy klikane drugi raz
    w.classList.add("is-flash");
    setTimeout(() => w.classList.remove("is-flash"), 1600);
  }
  if (opts.open) setTimeout(() => openFormPop(f, range.getBoundingClientRect()), 350);
}

document.addEventListener("DOMContentLoaded", () => {
  docCanvasEl?.addEventListener("click", onFormDocClick);
  docCanvasEl?.addEventListener("keydown", onFormDocKeydown);
  document.addEventListener("pointerdown", (e) => { if (formUi.pop && !formUi.pop.contains(e.target)) closeFormPop(); }, true);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && formUi.pop) { e.preventDefault(); e.stopPropagation(); closeFormPop(); }
  }, true);
  docViewportEl?.addEventListener("scroll", () => { if (formUi.pop && !formUi.pop.contains(document.activeElement)) closeFormPop(); }, { passive: true });
  // Po każdym narysowaniu dokumentu (otwarcie, edycja z panelu, Cofnij): pola od nowa.
  const origRender = window.renderStructurePanel;
  if (typeof origRender === "function") {
    window.renderStructurePanel = function renderStructureAndForms(...args) {
      const r = origRender.apply(this, args);
      refreshFormScan().catch(() => {});
      return r;
    };
  }
  const origIngest = window.ingestFile;
  if (typeof origIngest === "function") {
    window.ingestFile = async function ingestFileForms(...args) {
      formNoticePending = true;
      closeFormPop();
      return origIngest.apply(this, args);
    };
  }
  const origClear = window.clearDocumentState;
  if (typeof origClear === "function") {
    window.clearDocumentState = function clearDocumentStateForms(...args) {
      const r = origClear.apply(this, args);
      formScan = null;
      docFormCounts = { fields: 0, empty: 0 };
      closeFormPop();
      paintFormFields();
      return r;
    };
  }
});
