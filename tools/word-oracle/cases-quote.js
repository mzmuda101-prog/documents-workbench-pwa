// Cytaty i ramki: style wbudowane Worda + obramowanie/cieniowanie wprost (jak z generatorów AI).
// QUOTE_OPS: co zrobić z którym akapitem — styl z listy stylów (op "style") albo gumka przy kursorze
// ("clear"); porównujemy akapity po zapisie (normalizeParas — działa w przeglądarce).
const QUOTE_STYLES = `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="IntenseQuote"><w:name w:val="Intense Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="30"/><w:qFormat/><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="10" w:color="4472C4"/><w:bottom w:val="single" w:sz="4" w:space="10" w:color="4472C4"/></w:pBdr><w:spacing w:before="360" w:after="360"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="4472C4"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="BlockText"><w:name w:val="Block Text"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="1152" w:right="1152"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>`;
const BOX = '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="4" w:color="auto"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>';
const LEFT = '<w:pBdr><w:left w:val="single" w:sz="24" w:space="8" w:color="999999"/></w:pBdr><w:ind w:left="720"/>';
const SHADE = '<w:shd w:val="clear" w:color="auto" w:fill="DEEAF6"/><w:ind w:left="360" w:right="360"/>';
const R = (rpr, text) => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const QUOTE_BODY = [
  "Zwykły akapit.",
  { text: "Cytat stylem Quote.", style: "Quote" },
  { text: "Cytat intensywny.", style: "IntenseQuote" },
  { text: "Ramka wprost z tłem, styl Normalny z listy.", pPr: BOX },
  { text: "Pasek z lewej wprost, gumka przy kursorze.", pPr: LEFT },
  { pPr: BOX, raw: R("<w:i/><w:color w:val=\"1F4E79\"/>", "Cały pochylony kolorowy cytat od AI, ") + R("<w:i/><w:color w:val=\"1F4E79\"/><w:b/>", "pogrubione") + R("<w:i/><w:color w:val=\"1F4E79\"/>", " na końcu.") },
  { text: "Tekst blokowy na Normalny.", style: "BlockText" },
  { text: "Wyśrodkowany i wcięty wprost.", pPr: '<w:ind w:left="1440"/><w:jc w:val="center"/>' },
  { text: "Tło wprost bez ramki.", pPr: SHADE },
  { text: "Punkt listy na Normalny.", numId: 1 },
  "Koniec.",
];
const QUOTE_OPS = [
  { para: 1, op: "style", style: "normal" },
  { para: 2, op: "style", style: "normal" },
  { para: 3, op: "style", style: "normal" },
  { para: 4, op: "clear" },
  { para: 5, op: "style", style: "quote" },
  { para: 6, op: "style", style: "normal" },
  { para: 7, op: "style", style: "normal" },
  { para: 8, op: "style", style: "normal" },
  { para: 9, op: "style", style: "normal" },
];
const QUOTE_DOC = { styles: QUOTE_STYLES, body: QUOTE_BODY, abstractNums: [{ levels: [{ fmt: "decimal" }] }], nums: [{ id: 1, abs: 0 }] };

// akapity treści po zapisie → [{ style: nazwa stylu, ppr: [cechy akapitu], runs: "tekst{cechy znaków}…" }]
// (styl po NAZWIE: Word zapisuje polskie identyfikatory „Cytat”, aplikacja zostawia „Quote”).
// Wołać przez page.evaluate(normalizeParas, { doc, styles }) — strona ma CSP bez eval, a funkcję
// przekazaną wprost Playwright uruchamia przez protokół przeglądarki.
function normalizeParas({ doc: docXml, styles: stylesXml }) {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const parse = (x) => new DOMParser().parseFromString(x, "application/xml");
  const kids = (el, name) => Array.from(el?.childNodes || []).filter((n) => n.nodeType === 1 && n.namespaceURI === W && (!name || n.localName === name));
  const names = new Map();
  let def = "Normal";
  Array.from(parse(stylesXml).getElementsByTagNameNS(W, "style")).forEach((st) => {
    const name = kids(st, "name")[0]?.getAttributeNS(W, "val") || "";
    names.set(st.getAttributeNS(W, "styleId"), name);
    if (st.getAttributeNS(W, "type") === "paragraph" && st.getAttributeNS(W, "default") === "1") def = name;
  });
  const body = parse(docXml).getElementsByTagNameNS(W, "body")[0];
  return kids(body, "p").map((p) => {
    const pPr = kids(p, "pPr")[0];
    const sid = kids(pPr, "pStyle")[0]?.getAttributeNS(W, "val");
    const ppr = (pPr ? kids(pPr) : []).map((c) => c.localName).filter((n) => !["pStyle", "rPr", "sectPr"].includes(n)).sort();
    // sąsiednie fragmenty o tych samych cechach razem (Word i aplikacja dzielą fragmenty inaczej)
    const merged = [];
    Array.from(p.getElementsByTagNameNS(W, "r")).forEach((r) => {
      const t = kids(r, "t").map((x) => x.textContent).join("");
      if (!t) return;
      const rpr = kids(kids(r, "rPr")[0]).map((c) => c.localName).filter((n) => !["lang", "noProof"].includes(n)).sort().join(",");
      const last = merged[merged.length - 1];
      if (last && last.rpr === rpr) last.t += t; else merged.push({ t, rpr });
    });
    const runs = merged.map((x) => `${x.t}${x.rpr ? `{${x.rpr}}` : ""}`).join("");
    return { style: (sid ? names.get(sid) || sid : def).toLowerCase(), ppr, runs };
  });
}

module.exports = { QUOTE_STYLES, QUOTE_BODY, QUOTE_OPS, QUOTE_DOC, BOX, LEFT, normalizeParas };
