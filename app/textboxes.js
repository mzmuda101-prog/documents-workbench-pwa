// textboxes.js — pola tekstowe jak w Wordzie (Wstaw → Pole tekstowe, 2026-10-28).
//
// Zgłoszenie: w karcie terenu z PDF nie dało się dopisać napisu z boku mapki („ok 34 m.”) —
// tekst można było pisać tylko w istniejących akapitach. Pole tekstowe = kształt Worda z tekstem
// (wps:wsp + wps:txbx) przypięty do akapitu (wp:anchor), położenie względem strony. W pliku
// zawsze w mc:AlternateContent z kopią VML (w:pict) — tak zapisuje Word, starsze programy czytają
// kopię. Pola z plików Worda (dawniej akapit z polem był tylko do odczytu) obsługujemy tak samo.
//
// Podgląd: docx-preview nie rysuje kształtów wps (bierze kopię VML i gubi tekst). Przed
// rysowaniem wycinamy pola z document.xml (prepare), a po ułożeniu strony rysujemy je sami
// (render) — warstwa nad kartką, w punktach strony. Fragment z polem zostaje w pliku „wyspą”
// akapitu (docx-inline-edit.js), więc pisanie w akapicie kotwicy go nie gubi.
//
// Edycja: klik = zaznaczenie (uchwyty, karta), przeciąganie = przesuwanie (z przyciąganiem do
// marginesów, środka strony, krawędzi innych obiektów i linijek tekstu — jak linie wyrównania
// w Wordzie; Alt wyłącza), uchwyty = rozmiar, drugi klik / dwuklik / Enter = pisanie w polu
// (pole tekstowe nad polem, ten sam krój i rozmiar). Każda zmiana to jedna operacja na pliku
// (op "textbox") — jeden krok Cofnij, zapis zgodny z Wordem.

const TB_NS = {
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
};
const TB_EMU_PX = 9525; // 96 dpi
const TB_XMLNS = "http://www.w3.org/2000/xmlns/";

function tbKids(el, ns, name) {
  if (!el) return [];
  return Array.from(el.childNodes).filter((n) => n.nodeType === 1 && (!name || n.localName === name) && (!ns || n.namespaceURI === ns));
}
function tbKid(el, ns, name) { return tbKids(el, ns, name)[0] || null; }
function tbW(el, name) {
  if (!el) return null;
  const v = el.getAttributeNS(W_NS, name);
  return v != null && v !== "" ? v : el.getAttribute(`w:${name}`);
}
function tbOn(el) { // w:b, w:i… (w:val="0"/"false" = wyłączone)
  if (!el) return false;
  const v = tbW(el, "val");
  return v == null || !/^(0|false|off|none)$/i.test(v);
}

// ── model pliku ──────────────────────────────────────────────────────────────
// Pole tekstowe w fragmencie akapitu → { run, anchor, wsp, txbx, vTxbx, vShape } albo null.
// Tylko pojedynczy kształt z tekstem przypięty (wp:anchor) — grupy, pola w linii tekstu i stare
// pola tylko w VML zostają jak dawniej (akapit tylko do odczytu).
function tbBoxOfRun(r) {
  if (r?.nodeType !== 1 || r.localName !== "r" || r.namespaceURI !== W_NS) return null;
  let drawing = null, ac = null, count = 0;
  for (const c of tbKids(r)) {
    if (c.localName === "AlternateContent" && c.namespaceURI === TB_NS.mc) {
      const choice = tbKids(c, TB_NS.mc, "Choice").find((ch) => /(^|\s)wps(\s|$)/.test(ch.getAttribute("Requires") || ""));
      const d = choice && tbKid(choice, W_NS, "drawing");
      if (d) { drawing = d; ac = c; count++; } else if (tbKids(c).length) return null;
    } else if (c.localName === "drawing" && c.namespaceURI === W_NS) { drawing = c; count++; }
    else if (c.localName !== "rPr" && c.localName !== "lastRenderedPageBreak") return null; // tekst we wspólnym fragmencie
  }
  if (count !== 1) return null;
  const anchor = tbKid(drawing, TB_NS.wp, "anchor");
  const gd = anchor && anchor.getElementsByTagNameNS(TB_NS.a, "graphicData")[0];
  const wsp = gd && tbKid(gd, TB_NS.wps, "wsp");
  const txbx = wsp && tbKid(tbKid(wsp, TB_NS.wps, "txbx"), W_NS, "txbxContent");
  if (!txbx) return null;
  const fb = ac && tbKid(ac, TB_NS.mc, "Fallback");
  const vTxbx = fb ? fb.getElementsByTagNameNS(W_NS, "txbxContent")[0] || null : null;
  let vShape = vTxbx?.parentNode || null;
  while (vShape && !(vShape.namespaceURI === TB_NS.v && vShape.localName !== "textbox")) vShape = vShape.parentNode;
  return { run: r, anchor, wsp, txbx, vTxbx, vShape, ac };
}
function tbBoxesOfParagraph(p) {
  return p ? tbKids(p, W_NS, "r").map(tbBoxOfRun).filter(Boolean) : [];
}
// docx-inline-edit.js: fragment z polem tekstowym może być wyspą akapitu
function textboxRunSupported(r) { return !!tbBoxOfRun(r); }

function tbPos(anchor, axis) {
  if (anchor.getAttribute("simplePos") === "1") {
    const sp = tbKid(anchor, TB_NS.wp, "simplePos");
    return { rel: "page", off: (parseInt(sp?.getAttribute(axis === "H" ? "x" : "y"), 10) || 0) / TB_EMU_PX, align: null };
  }
  const el = tbKid(anchor, TB_NS.wp, `position${axis}`);
  const off = tbKid(el, TB_NS.wp, "posOffset");
  const al = tbKid(el, TB_NS.wp, "align");
  return {
    rel: el?.getAttribute("relativeFrom") || (axis === "H" ? "column" : "paragraph"),
    off: off ? (parseInt(off.textContent, 10) || 0) / TB_EMU_PX : al ? null : 0,
    align: al ? al.textContent.trim() : null,
  };
}

function tbColor(fillEl) { // a:solidFill → "#RRGGBB"
  const c = tbKids(fillEl)[0];
  if (!c) return null;
  if (c.localName === "srgbClr") return `#${c.getAttribute("val")}`;
  if (c.localName === "sysClr") return `#${c.getAttribute("lastClr") || "000000"}`;
  if (c.localName === "schemeClr") return /^(bg1|lt1)$/.test(c.getAttribute("val")) ? "#ffffff" : /^(tx1|dk1)$/.test(c.getAttribute("val")) ? "#000000" : null;
  return null;
}

function tbParaModel(p) {
  const pPr = tbKid(p, W_NS, "pPr");
  const sp = tbKid(pPr, W_NS, "spacing");
  const num = (v) => (v == null ? null : parseInt(v, 10));
  const runs = [];
  const walk = (el) => tbKids(el).forEach((n) => {
    if (n.namespaceURI !== W_NS) return;
    if (n.localName === "r") {
      const rPr = tbKid(n, W_NS, "rPr");
      const f = tbKid(rPr, W_NS, "rFonts");
      const props = {
        b: tbOn(tbKid(rPr, W_NS, "b")), i: tbOn(tbKid(rPr, W_NS, "i")),
        u: !!tbKid(rPr, W_NS, "u") && tbW(tbKid(rPr, W_NS, "u"), "val") !== "none",
        strike: tbOn(tbKid(rPr, W_NS, "strike")),
        sz: num(tbW(tbKid(rPr, W_NS, "sz"), "val")),
        color: tbW(tbKid(rPr, W_NS, "color"), "val"),
        font: f ? tbW(f, "ascii") || tbW(f, "hAnsi") : null,
        highlight: tbW(tbKid(rPr, W_NS, "highlight"), "val"),
      };
      let text = "";
      tbKids(n).forEach((c) => {
        if (c.localName === "t") text += c.textContent;
        else if (c.localName === "tab") text += "\t";
        else if (c.localName === "br" || c.localName === "cr") text += "\n";
        else if (c.localName === "noBreakHyphen") text += "‑";
      });
      if (text) runs.push({ ...props, text });
    } else if (/^(hyperlink|ins|smartTag|fldSimple|sdt|sdtContent|customXml)$/.test(n.localName)) walk(n);
  });
  walk(p);
  const markSz = num(tbW(tbKid(tbKid(pPr, W_NS, "rPr"), W_NS, "sz"), "val"));
  return {
    style: tbW(tbKid(pPr, W_NS, "pStyle"), "val"),
    jc: tbW(tbKid(pPr, W_NS, "jc"), "val"),
    spacing: sp ? { before: num(tbW(sp, "before")), after: num(tbW(sp, "after")), line: num(tbW(sp, "line")), rule: tbW(sp, "lineRule") } : null,
    runs, markSz,
  };
}
function tbParaPlain(p) { return tbParaModel(p).runs.map((r) => r.text).join(""); }

// Opis pola do rysowania (zwykły obiekt — przechodzi przez przerysowanie)
function tbDescribe(box) {
  const { anchor, wsp } = box;
  const ext = tbKid(anchor, TB_NS.wp, "extent");
  const spPr = tbKid(wsp, TB_NS.wps, "spPr");
  const bodyPr = tbKid(wsp, TB_NS.wps, "bodyPr");
  const style = tbKid(wsp, TB_NS.wps, "style");
  const ln = tbKid(spPr, TB_NS.a, "ln");
  let fill = null;
  if (tbKid(spPr, TB_NS.a, "solidFill")) fill = tbColor(tbKid(spPr, TB_NS.a, "solidFill")) || "#ffffff";
  else if (!tbKid(spPr, TB_NS.a, "noFill") && style && (parseInt(tbKid(style, TB_NS.a, "fillRef")?.getAttribute("idx"), 10) || 0) > 0) fill = "#ffffff";
  let line = null;
  if (ln && !tbKid(ln, TB_NS.a, "noFill")) line = { w: Math.max(1, (parseInt(ln.getAttribute("w"), 10) || 9525) / TB_EMU_PX), color: tbColor(tbKid(ln, TB_NS.a, "solidFill")) || "#000000" };
  else if (!ln && style && (parseInt(tbKid(style, TB_NS.a, "lnRef")?.getAttribute("idx"), 10) || 0) > 0) line = { w: 1, color: "#000000" };
  const ins = (name, def) => { const v = bodyPr?.getAttribute(name); return (v != null ? parseInt(v, 10) || 0 : def) / TB_EMU_PX; };
  const paras = tbKids(box.txbx, W_NS, "p");
  return {
    h: tbPos(anchor, "H"), v: tbPos(anchor, "V"),
    w: (parseInt(ext?.getAttribute("cx"), 10) || 0) / TB_EMU_PX,
    hgt: (parseInt(ext?.getAttribute("cy"), 10) || 0) / TB_EMU_PX,
    behind: anchor.getAttribute("behindDoc") === "1",
    z: parseInt(anchor.getAttribute("relativeHeight"), 10) || 0,
    fill, line,
    inset: { l: ins("lIns", 91440), t: ins("tIns", 45720), r: ins("rIns", 91440), b: ins("bIns", 45720) },
    vAnchor: bodyPr?.getAttribute("anchor") || "t",
    autoFit: !!tbKid(bodyPr, TB_NS.a, "spAutoFit"),
    // tabela / kontrolka w polu — tekst tylko do podglądu (zapis tekstu by ją zgubił)
    complex: tbKids(box.txbx).some((n) => n.localName !== "p"),
    paras: paras.map(tbParaModel),
    text: paras.map(tbParaPlain).join("\n"),
  };
}

// ── budowa XML ───────────────────────────────────────────────────────────────
function tbEsc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
function tbPt(px) { return Math.round(px * 0.75 * 100) / 100; }

// Akapit pola: bez odstępów (styl Normalny bywa z odstępem „po” 8 pt — w polu Word i tak go
// pokazuje, a karta potrzebuje zwartego napisu)
function tbParaXml(text, { jc, sz, bold } = {}) {
  const rPr = (sz || bold) ? `<w:rPr>${bold ? "<w:b/><w:bCs/>" : ""}${sz ? `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` : ""}</w:rPr>` : "";
  const runs = String(text || "").split("\t").map((part, i) => `${i ? `<w:r>${rPr}<w:tab/></w:r>` : ""}${part ? `<w:r>${rPr}<w:t xml:space="preserve">${tbEsc(part)}</w:t></w:r>` : ""}`).join("");
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>${jc ? `<w:jc w:val="${jc}"/>` : ""}${rPr}</w:pPr>${runs}</w:p>`;
}

function tbVmlStyle({ x, y, w, h, z }) {
  return `position:absolute;margin-left:${tbPt(x)}pt;margin-top:${tbPt(y)}pt;width:${tbPt(w)}pt;height:${tbPt(h)}pt;z-index:${z};mso-wrap-style:square;mso-position-horizontal-relative:page;mso-position-vertical-relative:page`;
}

// Fragment z polem (x, y, w, h w px strony; wartości jak w „Wstaw → Pole tekstowe” Worda, bez
// ramki i tła — napis na karcie, mapie, formularzu)
function tbRunXml({ id, name, z, x, y, w, h, paraXml, border = false, fill = false, autoFit = true }) {
  const E = (px) => Math.round(px * TB_EMU_PX);
  const fillXml = fill ? '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' : "<a:noFill/>";
  const lnXml = border ? '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:miter lim="800000"/></a:ln>' : '<a:ln w="9525"><a:noFill/></a:ln>';
  return `<w:r xmlns:w="${W_NS}" xmlns:mc="${TB_NS.mc}" xmlns:wp="${TB_NS.wp}" xmlns:a="${TB_NS.a}" xmlns:wps="${TB_NS.wps}" xmlns:v="${TB_NS.v}" xmlns:o="${TB_NS.o}"><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="${z}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${E(x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${E(y)}</wp:posOffset></wp:positionV><wp:extent cx="${E(w)}" cy="${E(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="${name}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${TB_NS.wps}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${E(w)}" cy="${E(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fillXml}${lnXml}</wps:spPr><wps:txbx><w:txbxContent>${paraXml}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0">${autoFit ? "<a:spAutoFit/>" : "<a:noAutofit/>"}</wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect id="${tbEsc(name)}" style="${tbVmlStyle({ x, y, w, h, z })}" filled="${fill ? "t" : "f"}" stroked="${border ? "t" : "f"}"${fill ? ' fillcolor="white"' : ""}${border ? ' strokecolor="black" strokeweight=".75pt"' : ""}><v:textbox${autoFit ? ' style="mso-fit-shape-to-text:t"' : ""}><w:txbxContent>${paraXml}</w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
}

// Prefiksy na korzeniu dokumentu (Word rozwiązuje „Requires” i pliki z innych programów
// mogą ich nie mieć)
function tbEnsureNamespaces(doc) {
  const root = doc.documentElement;
  [["mc", TB_NS.mc], ["wp", TB_NS.wp], ["a", TB_NS.a], ["wps", TB_NS.wps], ["v", TB_NS.v], ["o", TB_NS.o]].forEach(([p, ns]) => {
    if (!root.getAttributeNS(TB_XMLNS, p)) root.setAttributeNS(TB_XMLNS, `xmlns:${p}`, ns);
  });
}

function tbVmlStyleMap(shape) {
  const map = new Map();
  String(shape.getAttribute("style") || "").split(";").forEach((kv) => {
    const i = kv.indexOf(":");
    if (i > 0) map.set(kv.slice(0, i).trim(), kv.slice(i + 1).trim());
  });
  return map;
}
function tbSetVmlStyle(shape, map) {
  shape.setAttribute("style", [...map].map(([k, v]) => `${k}:${v}`).join(";"));
}

function tbSetGeometry(box, { x, y, w, h, fixedH }) {
  const { anchor, wsp } = box;
  const doc = anchor.ownerDocument;
  const E = (px) => String(Math.round(px * TB_EMU_PX));
  anchor.setAttribute("simplePos", "0");
  ["H", "V"].forEach((axis) => {
    let el = tbKid(anchor, TB_NS.wp, `position${axis}`);
    if (!el) {
      el = doc.createElementNS(TB_NS.wp, `wp:position${axis}`);
      const before = axis === "H" ? tbKid(anchor, TB_NS.wp, "positionV") || tbKid(anchor, TB_NS.wp, "extent") : tbKid(anchor, TB_NS.wp, "extent");
      anchor.insertBefore(el, before);
    }
    el.setAttribute("relativeFrom", "page");
    while (el.firstChild) el.removeChild(el.firstChild); // też wp14:pctPos… (ma pierwszeństwo w Wordzie)
    const off = doc.createElementNS(TB_NS.wp, "wp:posOffset");
    off.textContent = E(axis === "H" ? x : y);
    el.appendChild(off);
  });
  const ext = tbKid(anchor, TB_NS.wp, "extent");
  if (ext) { ext.setAttribute("cx", E(w)); ext.setAttribute("cy", E(h)); }
  // rozmiar w procentach strony (Word 2010+) wygrywa z extent — usuwamy
  tbKids(anchor).filter((n) => /^sizeRel[HV]$/.test(n.localName)).forEach((n) => anchor.removeChild(n));
  tbKids(anchor, TB_NS.mc, "AlternateContent").filter((n) => n.getElementsByTagNameNS("*", "sizeRelH").length || n.getElementsByTagNameNS("*", "sizeRelV").length).forEach((n) => anchor.removeChild(n));
  const xfrmExt = tbKid(tbKid(tbKid(wsp, TB_NS.wps, "spPr"), TB_NS.a, "xfrm"), TB_NS.a, "ext");
  if (xfrmExt) { xfrmExt.setAttribute("cx", E(w)); xfrmExt.setAttribute("cy", E(h)); }
  if (fixedH) {
    const bodyPr = tbKid(wsp, TB_NS.wps, "bodyPr");
    const fit = tbKid(bodyPr, TB_NS.a, "spAutoFit");
    if (fit) bodyPr.replaceChild(doc.createElementNS(TB_NS.a, "a:noAutofit"), fit);
  }
  if (box.vShape) {
    const map = tbVmlStyleMap(box.vShape);
    ["left", "top", "margin-left", "margin-top", "mso-position-horizontal", "mso-position-vertical", "mso-left-percent", "mso-top-percent", "mso-width-percent", "mso-height-percent"].forEach((k) => map.delete(k));
    map.set("position", "absolute");
    map.set("margin-left", `${tbPt(x)}pt`);
    map.set("margin-top", `${tbPt(y)}pt`);
    map.set("width", `${tbPt(w)}pt`);
    map.set("height", `${tbPt(h)}pt`);
    map.set("mso-position-horizontal-relative", "page");
    map.set("mso-position-vertical-relative", "page");
    tbSetVmlStyle(box.vShape, map);
    if (fixedH) {
      const tb = box.vTxbx?.parentNode;
      if (tb?.localName === "textbox") tb.setAttribute("style", String(tb.getAttribute("style") || "").replace(/mso-fit-shape-to-text:\s*t;?/, ""));
    }
  }
}

function tbSetHeight(box, h) {
  const cy = String(Math.round(h * TB_EMU_PX));
  tbKid(box.anchor, TB_NS.wp, "extent")?.setAttribute("cy", cy);
  tbKid(tbKid(tbKid(box.wsp, TB_NS.wps, "spPr"), TB_NS.a, "xfrm"), TB_NS.a, "ext")?.setAttribute("cy", cy);
  if (box.vShape) {
    const map = tbVmlStyleMap(box.vShape);
    map.set("height", `${tbPt(h)}pt`);
    tbSetVmlStyle(box.vShape, map);
  }
}

function tbFirstRPr(p) {
  for (const r of p ? p.getElementsByTagNameNS(W_NS, "r") : []) {
    const rPr = tbKid(r, W_NS, "rPr");
    if (rPr && r.getElementsByTagNameNS(W_NS, "t").length) return rPr;
  }
  return tbKid(tbKid(p, W_NS, "pPr"), W_NS, "rPr");
}

// Nowy tekst pola (wiersze = akapity). Akapit bez zmian zostaje dokładnie taki, jak był
// (z mieszanym formatem); zmieniony bierze akapit i format pierwszego fragmentu z tego samego
// miejsca — jak przepisanie akapitu w Wordzie.
function tbSetText(box, text) {
  const doc = box.txbx.ownerDocument;
  const oldPs = tbKids(box.txbx, W_NS, "p");
  const lines = String(text).replace(/\r\n?/g, "\n").split("\n");
  const newPs = lines.map((line, i) => {
    const old = oldPs[i];
    if (old && tbParaPlain(old) === line) return old.cloneNode(true);
    const tmpl = old || oldPs[oldPs.length - 1] || null;
    const np = doc.createElementNS(W_NS, "w:p");
    const pPr = tmpl && tbKid(tmpl, W_NS, "pPr");
    np.appendChild(pPr ? pPr.cloneNode(true) : doc.importNode(composeParse(`<w:pPr xmlns:w="${W_NS}"><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`).documentElement, true));
    const rPr = tbFirstRPr(tmpl);
    line.split("\t").forEach((part, k) => {
      const mk = (child) => {
        const r = doc.createElementNS(W_NS, "w:r");
        if (rPr) r.appendChild(doc.importNode(rPr, true));
        r.appendChild(child);
        np.appendChild(r);
      };
      if (k) mk(doc.createElementNS(W_NS, "w:tab"));
      if (part) {
        const t = doc.createElementNS(W_NS, "w:t");
        t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
        t.textContent = part;
        mk(t);
      }
    });
    return np;
  });
  [box.txbx, box.vTxbx].filter(Boolean).forEach((tx) => {
    while (tx.firstChild) tx.removeChild(tx.firstChild);
    newPs.forEach((p) => tx.appendChild(tx === box.txbx ? p : p.cloneNode(true)));
  });
}

function tbSetRPrChild(doc, holder, name, attrs) { // holder = w:r albo w:pPr (znak akapitu)
  let rPr = tbKid(holder, W_NS, "rPr");
  if (!rPr) {
    if (attrs == null) return;
    rPr = doc.createElementNS(W_NS, "w:rPr");
    if (holder.localName === "pPr") composeOrderedChild(holder, rPr, COMPOSE_PPR_ORDER);
    else holder.insertBefore(rPr, holder.firstChild);
  }
  const old = tbKid(rPr, W_NS, name);
  if (old) rPr.removeChild(old);
  if (attrs == null) return;
  composeOrderedChild(rPr, composeEl(doc, name, attrs), RPR_ORDER);
}

function tbSetSpFill(spPr, kind, el) { // kind: "fill" | "ln"
  const doc = spPr.ownerDocument;
  const FILLS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];
  if (kind === "fill") tbKids(spPr, TB_NS.a).filter((n) => FILLS.includes(n.localName)).forEach((n) => spPr.removeChild(n));
  else tbKids(spPr, TB_NS.a, "ln").forEach((n) => spPr.removeChild(n));
  const node = doc.importNode(composeParse(`<x xmlns:a="${TB_NS.a}">${el}</x>`).documentElement.firstChild, true);
  // kolejność w spPr: xfrm, geometria, wypełnienie, linia, efekty…
  const ORDER = ["xfrm", "custGeom", "prstGeom", ...FILLS, "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"];
  composeOrderedChild(spPr, node, ORDER);
}

// Format całego pola (pasek karty): rozmiar, pogrubienie, wyrównanie, ramka, tło
function tbSetFormat(box, f) {
  const doc = box.txbx.ownerDocument;
  [box.txbx, box.vTxbx].filter(Boolean).forEach((tx) => {
    tbKids(tx, W_NS, "p").forEach((p) => {
      const pPr = composeEnsurePPr(p);
      if (f.jc) composeSetPPrChild(pPr, "jc", { val: f.jc });
      const holders = [pPr, ...Array.from(p.getElementsByTagNameNS(W_NS, "r"))];
      holders.forEach((h) => {
        if (f.sizePt) {
          const v = String(Math.round(f.sizePt * 2));
          tbSetRPrChild(doc, h, "sz", { val: v });
          tbSetRPrChild(doc, h, "szCs", { val: v });
        }
        if (f.bold != null) {
          tbSetRPrChild(doc, h, "b", f.bold ? {} : null);
          tbSetRPrChild(doc, h, "bCs", f.bold ? {} : null);
        }
      });
    });
  });
  const spPr = tbKid(box.wsp, TB_NS.wps, "spPr");
  if (spPr && f.border != null) {
    tbSetSpFill(spPr, "ln", f.border ? '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:miter lim="800000"/></a:ln>' : '<a:ln w="9525"><a:noFill/></a:ln>');
    if (box.vShape) {
      box.vShape.setAttribute("stroked", f.border ? "t" : "f");
      if (f.border) { box.vShape.setAttribute("strokecolor", "black"); box.vShape.setAttribute("strokeweight", ".75pt"); }
    }
  }
  if (spPr && f.fill != null) {
    tbSetSpFill(spPr, "fill", f.fill ? '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' : "<a:noFill/>");
    if (box.vShape) {
      box.vShape.setAttribute("filled", f.fill ? "t" : "f");
      if (f.fill) box.vShape.setAttribute("fillcolor", "white");
    }
  }
}

// Wolny numer kształtu (wp:docPr id), nazwa niepowtarzalna także wśród id kształtów VML
// (walidator Office: id w kopii VML musi być jedyne — pole z Worda bywa „Pole tekstowe 2” przy id 1)
function tbNextIds(doc) {
  const docPrs = Array.from(doc.getElementsByTagNameNS(TB_NS.wp, "docPr"));
  const ids = docPrs.map((d) => parseInt(d.getAttribute("id"), 10) || 0);
  const zs = Array.from(doc.getElementsByTagNameNS(TB_NS.wp, "anchor")).map((a) => parseInt(a.getAttribute("relativeHeight"), 10) || 0);
  const taken = new Set([...docPrs.map((d) => d.getAttribute("name")), ...Array.from(doc.getElementsByTagNameNS(TB_NS.v, "*")).map((v) => v.getAttribute("id")).filter(Boolean)]);
  const id = Math.max(0, ...ids) + 1;
  let n = id;
  while (taken.has(`Pole tekstowe ${n}`)) n++;
  return { id, name: `Pole tekstowe ${n}`, z: Math.max(251659264, ...zs) + 1024 };
}

// op "textbox": { action: "insert" | "geometry" | "text" | "format" | "delete", index, nth, … }
function applyTextboxInXml(xml, edit) {
  const doc = composeParse(xml);
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[edit.index];
  if (!p) return { xml, count: 0 };
  if (edit.action === "insert") {
    tbEnsureNamespaces(doc);
    const { id, name, z } = tbNextIds(doc);
    const paraXml = String(edit.text || "").split("\n").map((line) => tbParaXml(line, { jc: edit.jc, sz: edit.sz })).join("");
    const run = doc.importNode(composeParse(tbRunXml({ id, name, z, x: edit.x, y: edit.y, w: edit.w, h: edit.h, paraXml: paraXml.replace(/<w:p>/g, `<w:p xmlns:w="${W_NS}">`), border: !!edit.border, fill: !!edit.fill })).documentElement, true);
    p.appendChild(run);
    return { xml: composeSerialize(doc), count: 1 };
  }
  const box = tbBoxesOfParagraph(p)[edit.nth | 0];
  if (!box) return { xml, count: 0 };
  if (edit.action === "geometry") {
    tbSetGeometry(box, edit);
    const to = edit.toIndex != null && edit.toIndex !== edit.index ? paragraphs[edit.toIndex] : null;
    if (to) to.appendChild(box.run); // inna strona — kotwica w akapicie tamtej strony (jak w Wordzie)
  } else if (edit.action === "text") {
    if (tbKids(box.txbx).some((n) => n.localName !== "p")) return { xml, count: 0 };
    tbSetText(box, edit.text);
    // „dopasuj kształt do tekstu”: wysokość jak po pisaniu (Word przelicza ją sam, starsze programy nie)
    if (edit.h && tbKid(tbKid(box.wsp, TB_NS.wps, "bodyPr"), TB_NS.a, "spAutoFit")) tbSetHeight(box, edit.h);
  } else if (edit.action === "format") {
    tbSetFormat(box, edit);
  } else if (edit.action === "delete") {
    box.run.parentNode.removeChild(box.run);
  } else {
    return { xml, count: 0 };
  }
  return { xml: composeSerialize(doc), count: 1 };
}

// ── podgląd ──────────────────────────────────────────────────────────────────
const dwbTextbox = (() => {
  const cache = new WeakMap(); // bytes → wynik prepare (to samo przerysowanie dwa razy)

  // Pola wycięte z document.xml do rysowania przez docx-preview; opis każdego do render()
  async function prepare(bytes) {
    if (!bytes || typeof JSZip === "undefined") return null;
    if (cache.has(bytes)) return cache.get(bytes);
    let out = null;
    try {
      const zip = await JSZip.loadAsync(bytes);
      const f = zip.file("word/document.xml");
      const xml = f ? await f.async("string") : "";
      if (xml.includes("txbxContent")) {
        const doc = composeParse(xml);
        const boxes = [];
        collectParagraphElements(doc.documentElement, "all").forEach((p, index) => {
          tbBoxesOfParagraph(p).forEach((box, nth) => {
            boxes.push({ index, nth, ...tbDescribe(box) });
            box.run.parentNode.removeChild(box.run);
          });
        });
        if (boxes.length) {
          zip.file("word/document.xml", composeSerialize(doc));
          out = { bytes: await zip.generateAsync({ type: "uint8array", compression: "STORE" }), boxes };
        }
      }
    } catch (err) {
      if (typeof log === "function") log(`Pola tekstowe: ${err.message || err}`, "warning");
      out = null;
    }
    cache.set(bytes, out);
    return out;
  }

  // Krój/rozmiar/odstępy akapitu danego stylu i jego tekstu — z CSS docx-preview (próbny akapit)
  function probeStyle(wrapper, style, probes) {
    const key = style || "";
    if (probes.has(key)) return probes.get(key);
    const art = wrapper.querySelector("section.docx > article");
    let out = null;
    if (art) {
      const p = document.createElement("p");
      if (style) p.className = `docx_${String(style).replace(/[ .]+/g, "-").replace(/[&]+/g, "and").toLowerCase()}`;
      p.style.cssText = "position:absolute;visibility:hidden;left:0;top:0;margin:0";
      const sp = document.createElement("span");
      sp.textContent = "x";
      p.appendChild(sp);
      art.appendChild(p);
      const cp = getComputedStyle(p), cs = getComputedStyle(sp);
      const pfs = parseFloat(cp.fontSize) || 16;
      const lh = parseFloat(cp.lineHeight);
      p.style.margin = "";
      const cm = getComputedStyle(p);
      out = {
        family: cs.fontFamily, size: parseFloat(cs.fontSize) || pfs, weight: cs.fontWeight, italic: cs.fontStyle === "italic", color: cs.color,
        ratio: Number.isFinite(lh) ? lh / pfs : 1,
        mt: parseFloat(cm.marginTop) || 0, mb: parseFloat(cm.marginBottom) || 0, align: cm.textAlign,
      };
      p.remove();
    }
    probes.set(key, out);
    return out;
  }

  const JC = { left: "left", start: "left", center: "center", right: "right", end: "right", both: "justify", distribute: "justify" };
  function buildParagraph(pm, wrapper, probes) {
    const pr = probeStyle(wrapper, pm.style, probes) || { family: "", size: 14.67, weight: "400", italic: false, color: "", ratio: 1, mt: 0, mb: 0, align: "left" };
    const div = document.createElement("div");
    div.className = "dwb-tb-p";
    div.style.textAlign = JC[pm.jc] || pr.align || "left";
    const sp = pm.spacing || {};
    div.style.marginTop = `${sp.before != null ? sp.before / 15 : pr.mt}px`; // twipy → px (20 tw = 1 pt)
    div.style.marginBottom = `${sp.after != null ? sp.after / 15 : pr.mb}px`;
    const firstFont = pm.runs.find((r) => r.font)?.font;
    const family = firstFont ? (typeof withGenericFontFallback === "function" && withGenericFontFallback(`"${firstFont}"`)) || `"${firstFont}"` : pr.family;
    const factor = typeof wordLineFactor === "function" ? wordLineFactor(family) : 1.15;
    const maxSz = Math.max(pm.markSz ? pm.markSz / 1.5 : 0, ...pm.runs.map((r) => (r.sz ? r.sz / 1.5 : pr.size)), pm.runs.length ? 0 : pr.size);
    if (sp.line != null && sp.rule && sp.rule !== "auto") div.style.lineHeight = `${sp.line / 15}px`; // dokładnie / co najmniej
    else div.style.lineHeight = String(Math.round((sp.line != null ? sp.line / 240 : pr.ratio) * factor * 1000) / 1000);
    div.style.fontFamily = family;
    div.style.fontSize = `${maxSz || pr.size}px`;
    div.style.color = pr.color;
    pm.runs.forEach((r) => {
      const s = document.createElement("span");
      s.style.fontFamily = r.font ? (typeof withGenericFontFallback === "function" && withGenericFontFallback(`"${r.font}"`)) || `"${r.font}"` : pr.family;
      s.style.fontSize = `${r.sz ? r.sz / 1.5 : pr.size}px`; // pół-punkty → px
      s.style.fontWeight = r.b ? "700" : pr.weight;
      s.style.fontStyle = r.i || pr.italic ? "italic" : "normal";
      const deco = [r.u && "underline", r.strike && "line-through"].filter(Boolean).join(" ");
      if (deco) s.style.textDecoration = deco;
      if (r.color && r.color !== "auto") s.style.color = `#${r.color}`;
      else if (pr.color) s.style.color = pr.color;
      if (r.highlight && r.highlight !== "none") s.style.backgroundColor = r.highlight;
      r.text.split("\n").forEach((part, i) => {
        if (i) s.appendChild(document.createElement("br"));
        if (part) s.appendChild(document.createTextNode(part.replace(/\t/g, "    ")));
      });
      div.appendChild(s);
    });
    if (!pm.runs.length) div.appendChild(document.createElement("br"));
    return div;
  }

  function sectionMetrics(section) {
    const cs = getComputedStyle(section);
    return {
      W: section.offsetWidth, H: section.offsetHeight,
      pl: parseFloat(cs.paddingLeft) || 0, pr: parseFloat(cs.paddingRight) || 0,
      pt: parseFloat(cs.paddingTop) || 0, pb: parseFloat(cs.paddingBottom) || 0,
    };
  }
  function scaleOf(el) {
    return el.offsetWidth ? el.getBoundingClientRect().width / el.offsetWidth || 1 : 1;
  }
  // Kartka, na której leży y (px sekcji bez skali): page-breaks.js stawia w sekcji prowadnice
  // marginesów, po jednej na stronę (góra i wysokość kartki). Bez nich (jedna strona, granice
  // wyłączone, Widok mobilny) cała sekcja to jedna kartka → null.
  function sheetAt(section, y) {
    let sheet = null;
    section.querySelectorAll(":scope > .dwb-page-margin-guide").forEach((g) => {
      const top = parseFloat(g.style.top) || 0;
      if (top <= y + 1) sheet = { top, H: parseFloat(g.style.height) || 0 };
    });
    return sheet;
  }
  function sectionY(section, el) {
    return (el.getBoundingClientRect().top - section.getBoundingClientRect().top) / scaleOf(section);
  }
  // Położenie pola na kartce (px bez skali widoku) — jak Word: od strony/marginesu/akapitu
  function boxPosition(b, p, section, m, w, h) {
    const frameH = { page: [0, m.W], margin: [m.pl, m.W - m.pr], column: [m.pl, m.W - m.pr], character: [m.pl, m.W - m.pr], leftMargin: [0, m.pl], rightMargin: [m.W - m.pr, m.W], insideMargin: [0, m.pl], outsideMargin: [m.W - m.pr, m.W] }[b.h.rel] || [m.pl, m.W - m.pr];
    let pTop = m.pt;
    if (/^(paragraph|line)$/.test(b.v.rel) || b.v.align == null && b.v.off == null) {
      const s = scaleOf(section);
      pTop = (p.getBoundingClientRect().top - section.getBoundingClientRect().top) / s;
    }
    const frameV = { page: [0, m.H], margin: [m.pt, m.H - m.pb], topMargin: [0, m.pt], bottomMargin: [m.H - m.pb, m.H], insideMargin: [0, m.pt], outsideMargin: [m.H - m.pb, m.H], paragraph: [pTop, pTop], line: [pTop, pTop] }[b.v.rel] || [pTop, pTop];
    const along = (frame, size, align, off) => {
      if (align === "center") return frame[0] + (frame[1] - frame[0] - size) / 2;
      if (align === "right" || align === "bottom" || align === "outside") return frame[1] - size;
      if (align === "left" || align === "top" || align === "inside") return frame[0];
      return frame[0] + (off || 0);
    };
    return { x: along(frameH, w, b.h.align, b.h.off), y: along(frameV, h, b.v.align, b.v.off) };
  }

  function buildBox(b, wrapper, probes) {
    const el = document.createElement("div");
    el.className = "dwb-tb";
    el.contentEditable = "false";
    el.setAttribute("role", "group");
    el.setAttribute("aria-label", typeof t === "function" ? t("textboxName") : "Pole tekstowe");
    const body = document.createElement("div");
    body.className = "dwb-tb-body";
    body.style.padding = `${b.inset.t}px ${b.inset.r}px ${b.inset.b}px ${b.inset.l}px`;
    if (b.fill) body.style.background = b.fill;
    if (b.line) body.style.border = `${b.line.w}px solid ${b.line.color}`;
    else el.classList.add("dwb-tb-noline");
    body.style.justifyContent = b.vAnchor === "ctr" ? "center" : b.vAnchor === "b" ? "flex-end" : "flex-start";
    b.paras.forEach((pm) => body.appendChild(buildParagraph(pm, wrapper, probes)));
    el.appendChild(body);
    el.style.width = `${b.w}px`;
    if (b.autoFit) el.style.minHeight = "0";
    else el.style.height = `${b.hgt}px`;
    if (b.autoFit) el.classList.add("dwb-tb-fit");
    el.__tb = b;
    return el;
  }

  // Miejsce pola na kartce akapitu kotwicy: względem akapitu — od akapitu (tam, gdzie stoi TERAZ,
  // także po odstępach między stronami), względem strony/marginesu — od góry jego kartki.
  function place(el, section) {
    const b = el.__tb;
    const p = el.__tbAnchor;
    const sheet = sheetAt(section, sectionY(section, p));
    const m = sectionMetrics(section);
    const h = b.autoFit ? el.offsetHeight || b.hgt : b.hgt;
    const pos = boxPosition(b, p, section, sheet ? { ...m, H: sheet.H } : m, b.w, h);
    el.style.left = `${pos.x}px`;
    el.style.top = `${pos.y + (sheet && !/^(paragraph|line)$/.test(b.v.rel) ? sheet.top : 0)}px`;
  }

  // Po ułożeniu stron (page-breaks.js): odstępy między kartkami przesuwają akapity w dół, a pola
  // ustawione przy rysowaniu zostawały na starym miejscu (pole przy akapicie na 7. stronie
  // wisiało ~1000 px wyżej) — ustawiamy je jeszcze raz.
  function relayout(wrapper) {
    let n = 0;
    wrapper?.querySelectorAll(".dwb-tb:not(.dwb-tb-flow):not(.is-dragging)").forEach((el) => {
      const section = el.parentElement;
      if (!el.__tb || !el.__tbAnchor?.isConnected || !section?.matches("section.docx")) return;
      place(el, section);
      n++;
    });
    if (n) ui.refreshCard();
    return n;
  }

  // Rysowanie pól po ułożeniu stron (renderDocxPreview). Widok mobilny (bez kartek): pole jako
  // zwykły blok pod akapitem kotwicy.
  function render(wrapper, prep, { reflow = false } = {}) {
    if (!prep?.boxes?.length || !wrapper) return 0;
    const paras = typeof collectPreviewParagraphElements === "function" ? collectPreviewParagraphElements(wrapper) : [];
    const probes = new Map();
    let n = 0;
    prep.boxes.forEach((b) => {
      const p = paras[b.index];
      if (!p) return;
      const el = buildBox(b, wrapper, probes);
      el.__tbAnchor = p;
      el.dataset.tbIndex = String(b.index);
      el.dataset.tbNth = String(b.nth);
      if (reflow) {
        el.classList.add("dwb-tb-flow");
        el.style.maxWidth = "100%";
        let after = p;
        while (after.nextElementSibling?.classList.contains("dwb-tb")) after = after.nextElementSibling;
        after.after(el);
        n++;
        return;
      }
      const section = p.closest("section.docx");
      if (!section) return;
      if (getComputedStyle(section).position === "static") section.style.position = "relative";
      section.style.isolation = "isolate";
      section.insertBefore(el, section.firstChild);
      place(el, section);
      el.style.zIndex = b.behind ? "-1" : "3";
      if (b.behind) el.classList.add("dwb-tb-behind");
      n++;
    });
    if (n) ui.afterRender(wrapper);
    return n;
  }

  // ── edycja ─────────────────────────────────────────────────────────────────
  const ui = (() => {
    let sel = null; // { el }
    let card = null;
    let editor = null; // { ta, el, text0, finish }
    let pendingSelect = null; // { index, nth, edit, char }
    let committing = null; // zapis tekstu pola w toku (Cofnij czeka — undo.js)
    let guides = [];

    const inDoc = (el) => !!el && !!docCanvasEl?.contains(el) && !!el.closest(".docx-edit-root");
    const editable = () => typeof readOnlyMode !== "undefined" && !readOnlyMode;
    const boxEls = () => [...(docCanvasEl?.querySelectorAll(".docx-preview-host .dwb-tb") || [])];
    const identity = (el) => {
      const index = typeof resolveParaIndex === "function" ? resolveParaIndex(el.__tbAnchor) : -1;
      return { index, nth: +el.dataset.tbNth || 0 };
    };

    async function edit(op, reselect) {
      const id = op.index != null ? op : null;
      if (!id || id.index < 0) return;
      pendingSelect = reselect || null;
      try {
        const n = await applyDocumentEdit({ op: "textbox", ...op });
        if (!n) pendingSelect = null;
      } catch (err) {
        pendingSelect = null;
        if (typeof log === "function") log(`Pole tekstowe: ${err.message || err}`, "error");
        if (typeof toast === "function") toast(t("saveFailed"), "error");
      }
    }

    function afterRender(wrapper) {
      if (!pendingSelect || !docCanvasEl?.contains(wrapper)) return;
      const want = pendingSelect;
      pendingSelect = null;
      const t0 = performance.now();
      const tick = () => {
        if (typeof inlineLocksPending !== "undefined" && inlineLocksPending && performance.now() - t0 < 4000) { requestAnimationFrame(tick); return; }
        const el = boxEls().find((b) => +b.dataset.tbIndex === want.index && +b.dataset.tbNth === want.nth);
        if (!el || !editable()) return;
        select(el);
        if (want.edit) startEdit(el, want.char || "");
      };
      requestAnimationFrame(tick);
    }

    // ── zaznaczenie + uchwyty + karta ──
    function select(el) {
      if (sel?.el === el) return;
      deselect();
      if (typeof composeUi !== "undefined") { composeUi.hideImageCard?.(); composeUi.hideLinkCard?.(); }
      sel = { el };
      el.classList.add("is-selected");
      el.tabIndex = -1;
      if (!el.classList.contains("dwb-tb-flow")) {
        // „dopasuj kształt do tekstu”: wysokość wynika z tekstu — bez uchwytów góra/dół (niskie pole
        // jednowierszowe na dotyku byłoby w całości uchwytem i nie dałoby się go przesunąć)
        const hs = el.__tb.autoFit ? ["nw", "ne", "e", "se", "sw", "w"] : ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
        hs.forEach((h) => {
          const d = document.createElement("div");
          d.className = "dwb-tb-h";
          d.dataset.h = h;
          el.appendChild(d);
        });
      }
      window.getSelection?.()?.removeAllRanges();
      try { el.focus({ preventScroll: true }); } catch (_) { /* stary Safari */ }
      showCard();
    }
    function deselect() {
      if (editor) editor.finish();
      if (!sel) return;
      sel.el.classList.remove("is-selected");
      sel.el.querySelectorAll(".dwb-tb-h").forEach((h) => h.remove());
      sel = null;
      hideCard();
      clearGuides();
    }

    const ICON = (path) => `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${path}</svg>`;
    const ALIGN_ICONS = {
      left: '<line x1="4" y1="6" x2="20" y2="6"/><line x1="4" y1="10" x2="14" y2="10"/><line x1="4" y1="14" x2="20" y2="14"/><line x1="4" y1="18" x2="14" y2="18"/>',
      center: '<line x1="4" y1="6" x2="20" y2="6"/><line x1="7" y1="10" x2="17" y2="10"/><line x1="4" y1="14" x2="20" y2="14"/><line x1="7" y1="18" x2="17" y2="18"/>',
      right: '<line x1="4" y1="6" x2="20" y2="6"/><line x1="10" y1="10" x2="20" y2="10"/><line x1="4" y1="14" x2="20" y2="14"/><line x1="10" y1="18" x2="20" y2="18"/>',
    };
    function boxFormat(b) {
      const runs = b.paras.flatMap((p) => p.runs);
      const sz = runs.find((r) => r.sz)?.sz || b.paras.find((p) => p.markSz)?.markSz || 0;
      const first = sel?.el.querySelector(".dwb-tb-p span, .dwb-tb-p");
      const sizePt = sz ? sz / 2 : first ? Math.round(parseFloat(getComputedStyle(first).fontSize) * 0.75 * 2) / 2 : 11;
      return { sizePt, bold: runs.length > 0 && runs.every((r) => r.b), jc: JC[b.paras[0]?.jc] || "left" };
    }
    function showCard() {
      hideCard();
      if (!sel) return;
      const b = sel.el.__tb;
      const fmt = boxFormat(b);
      const el = document.createElement("div");
      el.className = "image-card dwb-tb-card";
      el.setAttribute("role", "toolbar");
      el.setAttribute("aria-label", t("textboxTools"));
      const btn = (cls, key, html, on) => `<button type="button" class="tb-btn ${cls}" data-key="${key}"${on != null ? ` aria-pressed="${on}"` : ""}>${html}</button>`;
      el.innerHTML = [
        btn("tbc-edit", "textboxEdit", ICON('<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>')),
        '<span class="tb-sep" aria-hidden="true"></span>',
        btn("tbc-smaller", "textboxSmaller", '<span class="tbc-a">A</span><span class="tbc-pm">−</span>'),
        `<span class="tbc-size" aria-live="polite">${String(fmt.sizePt).replace(".", ",")}</span>`,
        btn("tbc-bigger", "textboxBigger", '<span class="tbc-a tbc-a-big">A</span><span class="tbc-pm">+</span>'),
        btn(`tbc-bold${fmt.bold ? " is-on" : ""}`, "textboxBold", "<b>B</b>", fmt.bold),
        '<span class="tb-sep" aria-hidden="true"></span>',
        ...["left", "center", "right"].map((a) => btn(`tbc-align${fmt.jc === a ? " is-on" : ""}`, `textboxAlign_${a}`, ICON(ALIGN_ICONS[a]), fmt.jc === a).replace("<button", `<button data-align="${a}"`)),
        '<span class="tb-sep" aria-hidden="true"></span>',
        btn(`tbc-border${b.line ? " is-on" : ""}`, "textboxBorder", ICON('<rect x="4" y="4" width="16" height="16" rx="1"/>'), !!b.line),
        btn(`tbc-fill${b.fill ? " is-on" : ""}`, "textboxFill", ICON('<rect x="4" y="4" width="16" height="16" rx="1" fill="currentColor" fill-opacity=".25"/>'), !!b.fill),
        btn("tbc-del", "textboxDelete", ICON('<polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>')),
      ].join("");
      el.querySelectorAll("button[data-key]").forEach((bt) => {
        const key = bt.dataset.key;
        bt.setAttribute("aria-label", t(key));
        bt.dataset.hint = "";
        bt.dataset.hintPl = I18N.pl[key];
        bt.dataset.hintEn = I18N.en[key];
        bt.dataset.hintDelay = "0.4";
      });
      if (b.complex) el.querySelector(".tbc-edit").disabled = true;
      el.addEventListener("mousedown", (e) => { if (e.target.closest("button")) e.preventDefault(); });
      el.addEventListener("click", (e) => {
        const bt = e.target.closest("button");
        if (!bt || !sel) return;
        const target = sel.el;
        if (bt.classList.contains("tbc-edit")) { startEdit(target); return; }
        if (bt.classList.contains("tbc-del")) { removeBox(target); return; }
        const f = {};
        if (bt.classList.contains("tbc-smaller") || bt.classList.contains("tbc-bigger")) {
          const sizes = typeof WORD_FONT_SIZES !== "undefined" ? WORD_FONT_SIZES : [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
          const cur = boxFormat(target.__tb).sizePt;
          f.sizePt = bt.classList.contains("tbc-bigger") ? sizes.find((s) => s > cur + 0.01) || cur + 2 : [...sizes].reverse().find((s) => s < cur - 0.01) || Math.max(1, cur - 1);
        } else if (bt.classList.contains("tbc-bold")) f.bold = !boxFormat(target.__tb).bold;
        else if (bt.dataset.align) f.jc = bt.dataset.align === "left" ? "left" : bt.dataset.align;
        else if (bt.classList.contains("tbc-border")) f.border = !target.__tb.line;
        else if (bt.classList.contains("tbc-fill")) f.fill = !target.__tb.fill;
        formatBox(target, f);
      });
      document.body.appendChild(el);
      card = el;
      placeCard();
    }
    function hideCard() { card?.remove(); card = null; }
    function placeCard() {
      if (!card || !sel) return;
      const vv = window.visualViewport;
      const r = sel.el.getBoundingClientRect();
      const vp = docViewportEl?.getBoundingClientRect();
      if (vp && (r.bottom < vp.top || r.top > vp.bottom)) { card.style.visibility = "hidden"; return; }
      card.style.visibility = "";
      const viewW = vv ? vv.width : window.innerWidth;
      const viewH = vv ? vv.height : window.innerHeight;
      const w = card.offsetWidth, h = card.offsetHeight;
      const left = Math.max(8, Math.min(r.left + r.width / 2 - w / 2, viewW - w - 8));
      let top = r.bottom + 12;
      if (top + h > viewH - 8) top = r.top - h - 12 >= 8 ? r.top - h - 12 : viewH - h - 8;
      card.style.left = `${left + (vv ? vv.offsetLeft : 0)}px`;
      card.style.top = `${top + (vv ? vv.offsetTop : 0)}px`;
    }

    async function formatBox(el, f) {
      const commitFirst = editor?.el === el ? editor.finish() : null;
      if (commitFirst) await commitFirst;
      const id = identity(el);
      await edit({ action: "format", ...id, ...f }, { ...id });
    }
    async function removeBox(el) {
      const id = identity(el);
      deselect();
      await edit({ action: "delete", ...id }, null);
    }

    // ── przesuwanie / rozmiar z przyciąganiem ──
    function clearGuides() { guides.forEach((g) => g.remove()); guides = []; }
    function showGuides(section, xs, ys, m) {
      clearGuides();
      xs.forEach((x) => {
        const g = document.createElement("div");
        g.className = "dwb-tb-guide dwb-tb-guide-v";
        g.style.left = `${x}px`;
        g.style.height = `${m.H}px`;
        section.appendChild(g);
        guides.push(g);
      });
      ys.forEach((y) => {
        const g = document.createElement("div");
        g.className = "dwb-tb-guide dwb-tb-guide-h";
        g.style.top = `${y}px`;
        g.style.width = `${m.W}px`;
        section.appendChild(g);
        guides.push(g);
      });
    }
    // Linie do przyciągania na kartce: marginesy, środek, krawędzie innych pól i obrazów,
    // linijki tekstu (góra i dół wiersza — napis w polu staje „na linii” tekstu obok)
    // Spód liter (linia bazowa) dla kroju i rozmiaru: zstępniki z metryk kroju (kanwa), jak liczy
    // przeglądarka prostokąt tekstu (góra = wysokość wydłużeń, dół = linia bazowa + zstępniki)
    const descents = new Map();
    function descentOf(cs) {
      const font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
      if (!descents.has(font)) {
        let d = parseFloat(cs.fontSize) * 0.21;
        try {
          const ctx = document.createElement("canvas").getContext("2d");
          ctx.font = font;
          const m = ctx.measureText("Hg");
          if (Number.isFinite(m.fontBoundingBoxDescent)) d = m.fontBoundingBoxDescent;
        } catch (_) { /* stary silnik — przybliżenie */ }
        descents.set(font, d);
      }
      return descents.get(font);
    }
    // Wiersze tekstu w elemencie: { T, B, base, L, R } (px kartki bez skali)
    function textLines(root, sr, s, limit = 800, own = false) {
      const out = new Map();
      const range = document.createRange();
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      let n = 0;
      for (let t = walker.nextNode(); t && n < limit; t = walker.nextNode()) {
        if (!t.data.trim() || !t.parentElement || t.parentElement.closest(own ? ".ff-glyph, .docx-tab" : ".dwb-tb, .ff-glyph, .docx-tab")) continue;
        n++;
        const cs = getComputedStyle(t.parentElement);
        const desc = descentOf(cs);
        range.selectNodeContents(t);
        for (const r of range.getClientRects()) {
          if (r.height < 4 || r.width < 1) continue;
          const base = Math.round(((r.bottom - sr.top) / s - desc / 1) * 2) / 2;
          const T = (r.top - sr.top) / s, B = (r.bottom - sr.top) / s, L = (r.left - sr.left) / s, R = (r.right - sr.left) / s;
          // jeden wiersz = ta sama linia bazowa w tym samym pasie (kolumna karty) — łączymy fragmenty
          const key = `${base}|${Math.round(L / 400)}`;
          const cur = out.get(key);
          if (cur) { cur.L = Math.min(cur.L, L); cur.R = Math.max(cur.R, R); cur.T = Math.min(cur.T, T); cur.B = Math.max(cur.B, B); }
          else out.set(key, { T, B, base, L, R });
        }
      }
      return [...out.values()];
    }
    // Linie do przyciągania na kartce (jak Affinity / Word): marginesy, środek, krawędzie i środki
    // innych pól i obrazów, linijki tekstu — linia bazowa (napis w polu staje „w tym samym wierszu”
    // także przy innym rozmiarze czcionki), lewy i prawy brzeg tekstu. Cel pamięta prostokąt źródła
    // (src): dalekie obiekty nie przyciągają, bliższe wygrywają przy remisie.
    function snapTargets(section, self) {
      const m = sectionMetrics(section);
      const s = scaleOf(section);
      const sr = section.getBoundingClientRect();
      const xs = [m.pl, m.W - m.pr, m.W / 2].map((v) => ({ v }));
      const ys = [m.pt, m.H - m.pb, m.H / 2].map((v) => ({ v }));
      const sizes = []; // szerokości innych obiektów (uchwyt: „ta sama szerokość”)
      section.querySelectorAll(":scope > .dwb-tb, :scope > .dwb-page-anchor, img").forEach((o) => {
        if (o === self || self.contains(o) || o.closest(".dwb-tb") === self) return;
        const r = o.getBoundingClientRect();
        if (!r.width || !r.height || r.width >= sr.width * 0.98) return; // warstwa całej strony nie jest obiektem
        const src = { L: (r.left - sr.left) / s, T: (r.top - sr.top) / s, R: (r.right - sr.left) / s, B: (r.bottom - sr.top) / s };
        xs.push({ v: src.L, src }, { v: src.R, src }, { v: (src.L + src.R) / 2, src });
        ys.push({ v: src.T, src }, { v: src.B, src }, { v: (src.T + src.B) / 2, src });
        sizes.push({ v: src.R - src.L, src });
      });
      const lines = [];
      section.querySelectorAll(":scope > article").forEach((art) => lines.push(...textLines(art, sr, s)));
      lines.forEach((l) => {
        const src = { L: l.L, T: l.T, R: l.R, B: l.B };
        xs.push({ v: l.L, src, kind: "text" }, { v: l.R, src, kind: "text" });
      });
      const bases = lines.map((l) => ({ v: l.base, src: { L: l.L, T: l.T, R: l.R, B: l.B }, kind: "base" }));
      return { m, xs, ys, bases, lines, sizes };
    }
    // Linia bazowa 1. wiersza tekstu pola — od górnej krawędzi pola (px kartki)
    function boxBaseline(el, sr, s) {
      const top = (el.getBoundingClientRect().top - sr.top) / s;
      const own = textLines(el, sr, s, 4, true).sort((a, b) => a.base - b.base)[0];
      if (own) return own.base - top;
      const pd = el.querySelector(".dwb-tb-p");
      const b = el.__tb;
      const cs = pd ? getComputedStyle(pd) : null;
      const lh = cs ? parseFloat(cs.lineHeight) || parseFloat(cs.fontSize) * 1.15 : 14;
      return b.inset.t + lh * 0.8;
    }
    // Odległość prostokąta źródła od pola (0 = nachodzą) — w pikselach kartki
    function gapTo(src, box) {
      if (!src) return 0;
      const gx = Math.max(0, src.L - box.R, box.L - src.R);
      const gy = Math.max(0, src.T - box.B, box.T - src.B);
      return Math.hypot(gx, gy);
    }
    // Najlepsze przyciągnięcie jednej osi. features: [{ f, id }] (położenia cech pola), targets: [{ v, src }].
    // Lepkość: przyciągnięta para trzyma, dopóki odchyłka < 1,8 × tolerancja (bez skakania między
    // bliskimi liniami). Dalekie źródła (> ⅓ szerokości kartki) pomijane; przy remisie bliższe.
    function snapAxis(features, targets, tol, box, near, lock) {
      let best = null;
      for (const f of features) {
        for (const tg of targets) {
          const d = tg.v - f.f;
          if (Math.abs(d) > tol) continue;
          const gap = gapTo(tg.src, box);
          if (gap > near) continue;
          const score = Math.abs(d) + gap * 0.01;
          if (!best || score < best.score) best = { d, at: tg.v, id: f.id, v: tg.v, kind: tg.kind, src: tg.src, score };
        }
      }
      if (best) return best; // najbliższe w tolerancji zawsze wygrywa
      // nic w pobliżu — przy odjeżdżaniu trzymana para jeszcze chwilę trzyma (lepkość)
      if (lock) {
        const f = features.find((x) => x.id === lock.id);
        if (f && Math.abs(lock.v - f.f) <= tol * 1.8) return { d: lock.v - f.f, at: lock.v, id: lock.id, v: lock.v, kind: lock.kind, src: lock.src, held: true };
      }
      return null;
    }
    function sectionAt(clientY, fallback) {
      const secs = [...(docCanvasEl?.querySelectorAll(".docx-preview-host section.docx") || [])];
      let best = fallback, dist = Infinity;
      secs.forEach((sc) => {
        const r = sc.getBoundingClientRect();
        const d = clientY < r.top ? r.top - clientY : clientY > r.bottom ? clientY - r.bottom : 0;
        if (d < dist) { dist = d; best = sc; }
      });
      return best;
    }

    let drag = null;
    function onPointerDown(e) {
      const el = e.target.closest?.(".dwb-tb");
      if (!el || !inDoc(el)) {
        if (sel && !e.target.closest?.(".dwb-tb-card, .dwb-tb-editor, .compose-pop")) deselect();
        return;
      }
      if (!editable() || e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      const wasSelected = sel?.el === el;
      if (editor?.el === el && e.target.closest(".dwb-tb-editor")) return;
      if (editor) editor.finish();
      select(el);
      if (el.classList.contains("dwb-tb-flow")) { if (wasSelected) startEdit(el); return; }
      const section = el.parentElement;
      const s = scaleOf(section);
      const handle = e.target.closest(".dwb-tb-h")?.dataset.h || null;
      drag = {
        el, handle, wasSelected, pointerId: e.pointerId, sx: e.clientX, sy: e.clientY, moved: false, s,
        x0: parseFloat(el.style.left) || 0, y0: parseFloat(el.style.top) || 0, w0: el.offsetWidth, h0: el.offsetHeight,
        grabX: (e.clientX - el.getBoundingClientRect().left) / s, grabY: (e.clientY - el.getBoundingClientRect().top) / s,
        section, targets: null, fixedH: false, touch: e.pointerType === "touch", lock: {},
      };
      try { el.setPointerCapture(e.pointerId); } catch (_) { /* już puszczony */ }
    }
    function onPointerMove(e) {
      if (!drag || e.pointerId !== drag.pointerId) return;
      const dx = e.clientX - drag.sx, dy = e.clientY - drag.sy;
      if (!drag.moved && Math.hypot(dx, dy) < 4) return;
      e.preventDefault();
      if (!drag.moved) {
        drag.moved = true;
        drag.el.classList.add("is-dragging");
        hideCard();
      }
      const el = drag.el;
      const b = el.__tb;
      // tolerancja w pikselach ekranu: palec mniej precyzyjny niż mysz
      const tol = (drag.touch ? 12 : 6) / drag.s;
      if (!drag.handle) {
        const sec = sectionAt(e.clientY, drag.section);
        if (sec !== el.parentElement) { sec.insertBefore(el, sec.firstChild); drag.targets = null; drag.lock = {}; }
        const s = scaleOf(sec);
        const sr = sec.getBoundingClientRect();
        if (!drag.targets) { drag.targets = snapTargets(sec, el); drag.base = boxBaseline(el, sr, s); }
        const { m, xs, ys, bases } = drag.targets;
        let x = (e.clientX - sr.left) / s - drag.grabX;
        let y = (e.clientY - sr.top) / s - drag.grabY;
        const w = drag.w0, h = drag.h0;
        const gx = [], gy = [];
        if (!e.altKey) {
          const box = { L: x, R: x + w, T: y, B: y + h };
          const near = m.W / 3;
          // poziomo: krawędzie, środek, początek i koniec tekstu pola (po wewnętrznych marginesach)
          const sx = snapAxis([{ f: x, id: "l" }, { f: x + w, id: "r" }, { f: x + w / 2, id: "c" }, { f: x + b.inset.l, id: "tl" }, { f: x + w - b.inset.r, id: "tr" }], xs, tol, box, near, drag.lock.x);
          drag.lock.x = sx || null;
          if (sx) { x += sx.d; gx.push(sx.at); }
          // pionowo: linia bazowa tekstu pola ↔ linia bazowa wiersza obok (pierwszeństwo przy remisie),
          // krawędzie i środek pola ↔ marginesy, środek strony, inne obiekty
          const sb = snapAxis([{ f: y + drag.base, id: "base" }], bases, tol, box, near, drag.lock.y?.id === "base" ? drag.lock.y : null);
          const se = snapAxis([{ f: y, id: "t" }, { f: y + h, id: "b" }, { f: y + h / 2, id: "m" }], ys, tol, box, near, drag.lock.y && drag.lock.y.id !== "base" ? drag.lock.y : null);
          // świeże przyciągnięcie wygrywa z trzymanym; między świeżymi — linia bazowa przy remisie
          const sy = sb && se ? (sb.held && !se.held ? se : se.held && !sb.held ? sb : Math.abs(sb.d) <= Math.abs(se.d) + 0.5 ? sb : se) : sb || se;
          drag.lock.y = sy || null;
          if (sy) { y += sy.d; gy.push(sy.at); }
        } else drag.lock = {};
        x = Math.max(-w + 8, Math.min(m.W - 8, x));
        y = Math.max(-h + 8, Math.min(m.H - 8, y));
        el.style.left = `${x}px`;
        el.style.top = `${y}px`;
        showGuides(sec, gx, gy, m);
      } else {
        const sec = el.parentElement;
        if (!drag.targets) drag.targets = snapTargets(sec, el);
        const { m, xs, ys, lines, sizes } = drag.targets;
        let L = drag.x0, T = drag.y0, R = drag.x0 + drag.w0, B = drag.y0 + drag.h0;
        const hd = drag.handle;
        const gx = [], gy = [];
        const box = { L, R, T, B };
        const near = m.W / 3;
        const yTargets = [...ys, ...lines.flatMap((l) => [{ v: l.T, src: l }, { v: l.B, src: l }])];
        const snapY = (v, id) => { const sn = e.altKey ? null : snapAxis([{ f: v, id }], yTargets, tol, box, near, drag.lock[id]); drag.lock[id] = sn || null; if (sn) gy.push(sn.at); return sn ? sn.at : v; };
        // krawędź w poziomie: linia (krawędź obiektu, brzeg tekstu, margines) albo ta sama szerokość
        // co inne pole / obraz — wygrywa bliższe (przy remisie szerokość)
        drag.sameW = null;
        const sideX = (raw, other, moving) => {
          if (e.altKey) return raw;
          const id = `h${moving}`;
          const sn = snapAxis([{ f: raw, id }], xs, tol, box, near, drag.lock[id]);
          const width = moving === "e" ? raw - other : other - raw;
          const sw = snapAxis([{ f: width, id: "w" }], sizes, tol, box, Infinity, drag.lock.w);
          const useW = sw && (!sn || Math.abs(sw.d) <= Math.abs(sn.d) + 0.25);
          drag.lock[id] = useW ? null : sn || null;
          drag.lock.w = useW ? sw : null;
          if (useW) { drag.sameW = sw.src || null; return moving === "e" ? other + sw.v : other - sw.v; }
          if (sn) { gx.push(sn.at); return sn.at; }
          return raw;
        };
        if (hd.includes("w")) L = Math.min(R - 12, sideX(drag.x0 + dx / drag.s, R, "w"));
        if (hd.includes("e")) R = Math.max(L + 12, sideX(drag.x0 + drag.w0 + dx / drag.s, L, "e"));
        if (hd.includes("n")) { T = Math.min(B - 8, snapY(drag.y0 + dy / drag.s, "hn")); drag.fixedH = true; }
        if (hd.includes("s")) { B = Math.max(T + 8, snapY(drag.y0 + drag.h0 + dy / drag.s, "hs")); drag.fixedH = true; }
        if (drag.sameW) gx.push(hd.includes("e") ? R : L, drag.sameW.L, drag.sameW.R); // „tej samej szerokości co…”
        el.style.left = `${L}px`;
        el.style.top = `${T}px`;
        el.style.width = `${R - L}px`;
        if (drag.fixedH || !b.autoFit) { el.style.height = `${B - T}px`; el.classList.remove("dwb-tb-fit"); }
        showGuides(sec, gx, gy, m);
      }
    }
    function onPointerUp(e) {
      if (!drag || e.pointerId !== drag.pointerId) return;
      const d = drag;
      drag = null;
      clearGuides();
      d.el.classList.remove("is-dragging");
      if (!d.moved) {
        if (d.wasSelected && !d.handle) startEdit(d.el);
        else placeCard();
        return;
      }
      commitGeometry(d.el, d.fixedH);
    }
    async function commitGeometry(el, fixedH = false) {
      const b = el.__tb;
      const sec = el.parentElement;
      const id = identity(el);
      let toIndex = id.index;
      let nth = id.nth;
      const top = parseFloat(el.style.top) || 0;
      const sheet = sheetAt(sec, top);
      const anchorSheet = el.__tbAnchor?.closest("section.docx") === sec ? sheetAt(sec, sectionY(sec, el.__tbAnchor)) : undefined;
      // inna sekcja albo inna kartka tej samej sekcji — kotwica w akapicie tamtej strony (jak w Wordzie)
      if (anchorSheet === undefined || (anchorSheet?.top ?? 0) !== (sheet?.top ?? 0)) {
        const target = anchorParagraphIn(sec, top);
        const ti = target ? resolveParaIndex(target) : -1;
        if (ti >= 0) {
          toIndex = ti;
          nth = boxEls().filter((o) => o !== el && o.__tbAnchor === target).length;
        }
      }
      const geo = { x: parseFloat(el.style.left) || 0, y: top - (sheet?.top || 0), w: el.offsetWidth, h: el.offsetHeight, fixedH: fixedH || !b.autoFit };
      await edit({ action: "geometry", ...id, toIndex, ...geo }, { index: toIndex, nth });
    }
    // Akapit kotwicy na danej kartce: akapit poza tabelą najbliżej nad polem (Word kotwiczy do
    // akapitu w pobliżu); w tabeli Word liczyłby położenie od komórki
    function anchorParagraphIn(section, y) {
      const ps = [...section.querySelectorAll(":scope > article p")].filter((p) => !p.dataset.dwbCont);
      const s = scaleOf(section);
      const sr = section.getBoundingClientRect();
      const top = (p) => (p.getBoundingClientRect().top - sr.top) / s;
      const body = ps.filter((p) => !p.closest("td, th"));
      const pool = body.length ? body : ps;
      let best = pool[0] || null;
      pool.forEach((p) => { if (top(p) <= y + 1) best = p; });
      return best;
    }

    // ── pisanie w polu ──
    function startEdit(el, char = "") {
      if (!el || !editable()) return;
      const b = el.__tb;
      if (b.complex) { toast?.(t("textboxComplex"), "info"); return; }
      if (editor?.el === el) return;
      if (editor) editor.finish();
      if (sel?.el !== el) select(el);
      const body = el.querySelector(".dwb-tb-body");
      const firstP = body.querySelector(".dwb-tb-p");
      const firstSpan = firstP?.querySelector("span") || firstP;
      const ta = document.createElement("textarea");
      ta.className = "dwb-tb-editor";
      ta.setAttribute("aria-label", t("textboxName"));
      ta.spellcheck = true;
      ta.value = b.text + char;
      const csP = firstP ? getComputedStyle(firstP) : null;
      const csS = firstSpan ? getComputedStyle(firstSpan) : null;
      const place = () => {
        const s = scaleOf(el.parentElement || el);
        const r = el.getBoundingClientRect();
        ta.style.left = `${r.left + window.scrollX}px`;
        ta.style.top = `${r.top + window.scrollY}px`;
        ta.style.width = `${r.width}px`;
        ta.style.padding = `${b.inset.t * s}px ${b.inset.r * s}px ${b.inset.b * s}px ${b.inset.l * s}px`;
        if (csS) {
          ta.style.fontFamily = csS.fontFamily;
          ta.style.fontSize = `${parseFloat(csS.fontSize) * s}px`;
          ta.style.fontWeight = csS.fontWeight;
          ta.style.fontStyle = csS.fontStyle;
          ta.style.color = csS.color;
        }
        if (csP) {
          ta.style.textAlign = csP.textAlign;
          const lh = parseFloat(csP.lineHeight);
          ta.style.lineHeight = Number.isFinite(lh) ? `${lh * s}px` : csP.lineHeight;
        }
        ta.style.height = "0px";
        const need = ta.scrollHeight;
        ta.style.height = `${Math.max(r.height, need)}px`;
        if (b.autoFit) el.style.height = `${Math.max(r.height, need) / s}px`;
      };
      el.classList.add("is-editing");
      document.body.appendChild(ta);
      place();
      ta.focus({ preventScroll: true });
      ta.setSelectionRange(ta.value.length, ta.value.length);
      const onScroll = () => place();
      docViewportEl?.addEventListener("scroll", onScroll, { passive: true });
      window.visualViewport?.addEventListener("resize", onScroll);
      ta.addEventListener("input", () => place());
      ta.addEventListener("keydown", (e) => {
        e.stopPropagation(); // skróty dokumentu nie działają na tekst pola
        if (e.key === "Escape") { e.preventDefault(); editor?.finish(); try { el.focus({ preventScroll: true }); } catch (_) { /* — */ } }
        if ((e.ctrlKey || e.metaKey) && /^s$/i.test(e.key)) { // zapis: najpierw tekst pola do pliku
          e.preventDefault();
          editor?.finish().then(() => document.dispatchEvent(new KeyboardEvent("keydown", { key: e.key, code: e.code, ctrlKey: e.ctrlKey, metaKey: e.metaKey, shiftKey: e.shiftKey, bubbles: true })));
        }
      });
      ["keyup", "keypress", "beforeinput", "input", "paste", "cut", "copy"].forEach((ev) => ta.addEventListener(ev, (e) => e.stopPropagation()));
      ta.addEventListener("blur", () => setTimeout(() => { if (editor?.ta === ta && document.activeElement !== ta) editor.finish(); }, 0));
      let done = null;
      const finish = () => {
        if (done) return done;
        docViewportEl?.removeEventListener("scroll", onScroll);
        window.visualViewport?.removeEventListener("resize", onScroll);
        const text = ta.value.replace(/\r\n?/g, "\n");
        const s = scaleOf(el.parentElement || el);
        const hNeed = ta.scrollHeight / s;
        ta.remove();
        el.classList.remove("is-editing");
        editor = null;
        if (text === b.text) { done = Promise.resolve(); el.style.height = b.autoFit ? "" : el.style.height; return done; }
        const id = identity(el);
        const op = { action: "text", ...id, text };
        if (b.autoFit) op.h = Math.max(8, hNeed);
        done = edit(op, { ...id });
        committing = done;
        done.finally(() => { if (committing === done) committing = null; });
        return done;
      };
      editor = { ta, el, finish };
    }

    // ── klawiatura przy zaznaczonym polu (przed obsługą dokumentu) ──
    let nudge = null;
    function onKeyDown(e) {
      if (!sel || editor || drag) return;
      if (document.activeElement !== sel.el && !sel.el.contains(document.activeElement)) return;
      const el = sel.el;
      const k = e.key;
      const mod = e.ctrlKey || e.metaKey;
      if (k === "Escape") { e.preventDefault(); e.stopPropagation(); deselect(); return; }
      if (k === "Delete" || k === "Backspace") { e.preventDefault(); e.stopPropagation(); removeBox(el); return; }
      if (k === "Enter" || k === "F2") { e.preventDefault(); e.stopPropagation(); startEdit(el); return; }
      if (/^Arrow/.test(k) && !mod && !el.classList.contains("dwb-tb-flow")) {
        e.preventDefault();
        e.stopPropagation();
        const step = (e.shiftKey ? 10 : 1) / 0.75; // 1 pt / 10 pt
        const dx = k === "ArrowLeft" ? -step : k === "ArrowRight" ? step : 0;
        const dy = k === "ArrowUp" ? -step : k === "ArrowDown" ? step : 0;
        el.style.left = `${(parseFloat(el.style.left) || 0) + dx}px`;
        el.style.top = `${(parseFloat(el.style.top) || 0) + dy}px`;
        placeCard();
        clearTimeout(nudge);
        nudge = setTimeout(() => { if (sel?.el === el) commitGeometry(el); }, 500);
        return;
      }
      // zwykły znak przy zaznaczonym polu = pisanie w nim (jak w Wordzie)
      if (k.length === 1 && !mod && !e.altKey) { e.preventDefault(); e.stopPropagation(); startEdit(el, k); }
    }

    // ── Wstaw → Pole tekstowe ──
    async function insertAtCaret() {
      if (!editable()) return;
      const host = docCanvasEl?.querySelector(".docx-preview-host");
      const p = typeof restoreDocCaret === "function" ? restoreDocCaret() : null;
      const section = p?.closest("section.docx") || host?.querySelector("section.docx");
      if (!section) { toast?.(t("insertNoCaret"), "info"); return; }
      const reflow = !p?.closest("section.docx > article") && !section;
      if (reflow) return;
      const s = scaleOf(section);
      const sr = section.getBoundingClientRect();
      const m = sectionMetrics(section);
      let cx = m.pl, cy = m.pt;
      const range = window.getSelection?.()?.rangeCount ? window.getSelection().getRangeAt(0) : null;
      const rr = range && section.contains(range.startContainer) ? range.getBoundingClientRect() : null;
      if (rr && (rr.width || rr.height)) { cx = (rr.left - sr.left) / s; cy = (rr.top - sr.top) / s; }
      else if (p) { const pr = p.getBoundingClientRect(); cx = (pr.left - sr.left) / s; cy = (pr.top - sr.top) / s; }
      const w = Math.min(189, m.W - m.pr - cx > 60 ? m.W - m.pr - cx : 189); // 5 cm
      const lineH = 20;
      const x = Math.max(0, Math.min(m.W - w, cx - 9.6)); // tekst pola zaczyna się w miejscu kursora
      const y = Math.max(0, Math.min(m.H - lineH, cy - 4.8));
      const anchor = anchorParagraphIn(section, y) || p;
      const index = anchor ? resolveParaIndex(anchor) : -1;
      if (index < 0) return;
      const nth = boxEls().filter((o) => o.__tbAnchor === anchor).length;
      const sheetTop = sheetAt(section, y)?.top || 0; // plik: położenie względem strony
      await edit({ action: "insert", index, x, y: y - sheetTop, w, h: lineH + 9.6, text: "" }, { index, nth, edit: true });
    }

    function init() {
      if (!docCanvasEl) return;
      window.addEventListener("pointerdown", onPointerDown, true);
      window.addEventListener("pointermove", onPointerMove, true);
      window.addEventListener("pointerup", onPointerUp, true);
      window.addEventListener("pointercancel", onPointerUp, true);
      window.addEventListener("keydown", onKeyDown, true);
      // klik/dotyk w polu nie stawia kursora w tekście pod spodem
      ["mousedown", "click", "dblclick", "touchstart", "contextmenu"].forEach((ev) => window.addEventListener(ev, (e) => {
        const el = e.target.closest?.(".dwb-tb");
        if (!el || !inDoc(el) || !editable()) return;
        if (ev === "dblclick") startEdit(el);
        if (ev !== "touchstart") e.preventDefault();
        e.stopPropagation();
      }, { capture: true, passive: false }));
      docViewportEl?.addEventListener("scroll", () => placeCard(), { passive: true });
      window.addEventListener("resize", () => placeCard());
      if (typeof readModeEl !== "undefined") readModeEl?.addEventListener("change", () => deselect());
    }
    return { init, afterRender, insertAtCaret, deselect, refreshCard: () => { if (sel?.el) placeCard(); }, isEditing: () => !!editor, finishEdit: () => editor?.finish(), settle: () => (editor ? editor.finish() : committing), pending: () => committing, snapAxis,
      dropAnchoredTo: (p) => boxEls().forEach((el) => { if (el.__tbAnchor === p) { if (sel?.el === el) deselect(); el.remove(); } }) };
  })();

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", () => ui.init());
  else ui.init();

  return { prepare, render, relayout, insert: ui.insertAtCaret, deselect: ui.deselect, finishEdit: ui.finishEdit, settle: ui.settle, pending: ui.pending, dropAnchoredTo: ui.dropAnchoredTo, isEditing: ui.isEditing, _snapAxis: ui.snapAxis }; // _snapAxis: testy
})();
