// keys-state.js — stan dokumentu w zapisie tabeli klawiszy (scripts/keys-table.js) z word/document.xml.
// Ten sam odczyt dla aplikacji (scripts/keys-table-playwright.js) i prawdziwego Worda
// (tools/word-oracle/run-keys.js) — obie strony porównywane jednakowo.
//
// Akapity po kolei (także w komórkach tabeli), „ ¶ ” między nimi, `rodzaj:tekst`, pusty = ∅.
// Rodzaje: cell (w komórce tabeli), pb (sam podział strony), pbb („od nowej strony”), list, rule,
// line, h1, quote, ind — pierwszy pasujący. Podział strony wewnątrz tekstu = „⤓”.
// Liczone w PRZEGLĄDARCE (DOMParser): keysStateIn(page, xml) — tak samo dla obu stron.

// Samodzielna funkcja (bez odwołań na zewnątrz) — przekazywana do przeglądarki jako tekst.
function stateFromDocumentXml(xml) {
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const kids = (el, name) => Array.from(el?.childNodes || []).filter((n) => n.nodeType === 1 && n.namespaceURI === W_NS && (!name || n.localName === name));
  const kid = (el, name) => kids(el, name)[0] || null;
  const val = (el, name) => (el ? el.getAttributeNS(W_NS, name) || el.getAttribute(`w:${name}`) || null : null);

  // tekst akapitu: w:t (bez usuniętych w śledzonych zmianach), podział strony jako „⤓”
  const paraText = (p) => {
    let out = "";
    const walk = (n) => {
      for (const c of Array.from(n.childNodes || [])) {
        if (c.nodeType !== 1 || c.namespaceURI !== W_NS) continue;
        if (c.localName === "del" || c.localName === "pPr" || c.localName === "rPr") continue;
        if (c.localName === "t") out += c.textContent;
        else if (c.localName === "br" && val(c, "type") === "page") out += "⤓";
        else walk(c);
      }
    };
    walk(p);
    return out;
  };

  const describe = (p, inCell) => {
    let text = paraText(p);
    const pPr = kid(p, "pPr");
    const style = val(kid(pPr, "pStyle"), "val") || "";
    const ind = kid(pPr, "ind");
    const pbb = kid(pPr, "pageBreakBefore");
    let kind = "";
    if (inCell) kind = "cell";
    else if (text === "⤓") { kind = "pb"; text = ""; }
    else if (pbb && !/^(0|false|off)$/.test(val(pbb, "val") || "")) kind = "pbb";
    else if (kid(pPr, "numPr")) kind = "list";
    else if (kid(kid(pPr, "pBdr"), "bottom")) kind = text ? "line" : "rule";
    else if (/heading ?1|nagwek1/i.test(style)) kind = "h1";
    else if (/quote|cytat/i.test(style)) kind = "quote";
    else if (parseInt(val(ind, "left") || val(ind, "start") || "0", 10) > 0) kind = "ind";
    const t = text || "∅";
    return kind ? `${kind}:${t}` : t;
  };

  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const out = [];
  const visit = (el, inCell) => {
    for (const c of kids(el)) {
      if (c.localName === "p") out.push(describe(c, inCell));
      else if (c.localName === "tbl") kids(c, "tr").forEach((tr) => kids(tr, "tc").forEach((tc) => visit(tc, true)));
      else if (c.localName === "sdt") visit(kid(c, "sdtContent"), inCell);
    }
  };
  visit(doc.getElementsByTagNameNS(W_NS, "body")[0], false);
  return out.join(" ¶ ");
}

// w przeglądarce Playwrighta (strona aplikacji albo pusta)
const keysStateIn = (page, xml) => page.evaluate(([src, x]) => (0, eval)(`(${src})`)(x), [stateFromDocumentXml.toString(), xml]);

module.exports = { stateFromDocumentXml, keysStateIn };
