// Numeracja przypisów dolnych/końcowych: przypadki (każdy = osobny dokument).
// sections: [{ fn, en (wnętrze <w:footnotePr>/<w:endnotePr> w sectPr), type, items }],
// items: "F" przypis dolny, "E" końcowy, "|" podział strony; settings: { fn, en } (settings.xml).
// Tekst odnośnika „R01:” / „E01:”, treść przypisu „N01” / „M01” — po nich czytamy numer z PDF Worda.
const { W, esc } = require("./build-docx");

const FT = "application/vnd.openxmlformats-officedocument.wordprocessingml.";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const fmt = (v) => `<w:numFmt w:val="${v}"/>`;
const start = (v) => `<w:numStart w:val="${v}"/>`;
const restart = (v) => `<w:numRestart w:val="${v}"/>`;
const PG = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>';

function sectInner(s) {
  return `${s.fn ? `<w:footnotePr>${s.fn}</w:footnotePr>` : ""}${s.en ? `<w:endnotePr>${s.en}</w:endnotePr>` : ""}${s.type ? `<w:type w:val="${s.type}"/>` : ""}`;
}

function noteXml(kind, id, text) {
  const st = kind === "footnote" ? "Footnote" : "Endnote";
  return `<w:${kind} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="${st}Text"/></w:pPr><w:r><w:rPr><w:rStyle w:val="${st}Reference"/></w:rPr><w:${kind}Ref/></w:r><w:r><w:t xml:space="preserve"> ${esc(text)}</w:t></w:r></w:p></w:${kind}>`;
}
function partXml(kind, notes) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind}s ${W}><w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>${notes.join("")}</w:${kind}s>`;
}
const STYLES = ["Footnote", "Endnote"].map((k) => `<w:style w:type="paragraph" w:styleId="${k}Text"><w:name w:val="${k.toLowerCase()} text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style><w:style w:type="character" w:styleId="${k}Reference"><w:name w:val="${k.toLowerCase()} reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>`).join("");

// przypadek → argumenty buildDocx
function toDocx(c) {
  const body = [];
  const fns = [];
  const ens = [];
  c.sections.forEach((s, si) => {
    for (const it of s.items.split(/\s+/).filter(Boolean)) {
      if (it === "|") { body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>'); continue; }
      const kind = it === "F" ? "footnote" : "endnote";
      const list = kind === "footnote" ? fns : ens;
      const n = String(list.length + 1).padStart(2, "0");
      const id = list.length + 1;
      list.push(noteXml(kind, id, `${kind === "footnote" ? "N" : "M"}${n}`));
      body.push(`<w:p><w:r><w:t xml:space="preserve">${kind === "footnote" ? "R" : "E"}${n}:</w:t></w:r><w:r><w:rPr><w:rStyle w:val="${kind === "footnote" ? "Footnote" : "Endnote"}Reference"/></w:rPr><w:${kind}Reference w:id="${id}"/></w:r></w:p>`);
    }
    if (si < c.sections.length - 1) body.push(`<w:p><w:pPr><w:sectPr>${sectInner(s)}${PG}</w:sectPr></w:pPr></w:p>`);
  });
  const last = c.sections[c.sections.length - 1];
  const settingsTail = `${c.settings?.fn ? `<w:footnotePr>${c.settings.fn}<w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>` : ""}${c.settings?.en ? `<w:endnotePr>${c.settings.en}<w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>` : ""}`;
  return {
    body,
    styles: STYLES,
    settingsTail,
    sectPr: sectInner(last),
    extraParts: {
      "footnotes.xml": { type: `${FT}footnotes+xml`, rel: `${REL}footnotes`, xml: partXml("footnote", fns) },
      "endnotes.xml": { type: `${FT}endnotes+xml`, rel: `${REL}endnotes`, xml: partXml("endnote", ens) },
    },
  };
}

const S = (items, o = {}) => ({ items, ...o });
const CASES = [
  { name: "ciagle-domyslne", sections: [S("F F E F E")] },
  { name: "settings-fmt-start", settings: { fn: fmt("upperRoman") + start(5), en: fmt("decimal") }, sections: [S("F F E F E")] },
  { name: "settings-litery-ponad-26", settings: { fn: fmt("lowerLetter") }, sections: [S(Array(29).fill("F").join(" "))] },
  { name: "chicago", settings: { fn: fmt("chicago") }, sections: [S("F F F F F F F F F")] },
  { name: "sekcje-ciagle", sections: [S("F F E"), S("F E F")] },
  { name: "sekcje-eachSect-obie", sections: [S("F F", { fn: restart("eachSect") }), S("F F F", { fn: restart("eachSect") })] },
  { name: "sekcje-eachSect-tylko-druga", sections: [S("F F"), S("F F", { fn: restart("eachSect") }), S("F F")] },
  { name: "sekcje-eachSect-tylko-pierwsza", sections: [S("F F", { fn: restart("eachSect") }), S("F F"), S("F")] },
  { name: "sekcje-eachSect-w-settings", settings: { fn: restart("eachSect") }, sections: [S("F F"), S("F F")] },
  { name: "sekcje-rozne-fmt", sections: [S("F F", { fn: fmt("lowerLetter") }), S("F F", { fn: fmt("upperRoman") }), S("F")] },
  { name: "sekcje-fmt-dziedziczenie", settings: { fn: fmt("lowerRoman") }, sections: [S("F F", { fn: fmt("upperLetter") }), S("F F")] },
  { name: "sekcja-numStart-ciagle", sections: [S("F F"), S("F F", { fn: start(10) })] },
  { name: "sekcja-numStart-eachSect", sections: [S("F F", { fn: start(3) + restart("eachSect") }), S("F F", { fn: start(7) + restart("eachSect") })] },
  { name: "sekcja-ciagla-eachSect", sections: [S("F F", { fn: restart("eachSect") }), S("F F", { fn: restart("eachSect"), type: "continuous" })] },
  { name: "eachPage-podzialy", sections: [S("F F | F F F | F", { fn: restart("eachPage") })] },
  { name: "eachPage-w-settings", settings: { fn: restart("eachPage") }, sections: [S("F F | F | F F")] },
  { name: "eachPage-fmt-start", sections: [S("F F | F F", { fn: restart("eachPage") + fmt("upperLetter") + start(2) })] },
  { name: "eachPage-sekcje", sections: [S("F | F F", { fn: restart("eachPage") }), S("F F | F", { fn: restart("eachPage") })] },
  { name: "koncowe-eachSect", sections: [S("E E F", { en: restart("eachSect") }), S("E F E", { en: restart("eachSect") })] },
  { name: "koncowe-fmt-sekcji", sections: [S("E E", { en: fmt("upperLetter") }), S("E E", { en: fmt("decimal") })] },
  { name: "koncowe-numStart", settings: { en: start(4) }, sections: [S("E E E")] },
  { name: "fmt-rozne", sections: [S("F F F", { fn: fmt("decimalEnclosedCircle") }), S("F F", { fn: fmt("numberInDash") }), S("F F", { fn: fmt("ordinal") })] },
  { name: "fmt-slownie-pl", sections: [S("F F F", { fn: fmt("cardinalText") }), S("F F F", { fn: fmt("ordinalText") }), S("E E", { en: fmt("ordinal") })] },
  { name: "eachPage-numStart-zero", sections: [S("F F | F", { fn: restart("eachPage") + start(0) })] },
  { name: "numStart-zero-ciagle", sections: [S("F F F", { fn: start(0) + fmt("lowerRoman") })] },
  { name: "fmt-zero-wiodace-sekcja", sections: [S("F F F F F F F F F F F", { fn: fmt("decimalZero") })] },
  { name: "fmt-zero-wiodace", settings: { fn: fmt("decimalZero") }, sections: [S("F F F F F F F F F F F")] },
];

module.exports = { CASES, toDocx };
