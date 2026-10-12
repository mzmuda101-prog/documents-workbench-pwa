// paste-edge-summary.js — streszczenie dokumentu do porównania z wynikiem Worda (paste-edge-playwright.js):
// akapity (treść + [PRZYPIS]/[OBRAZ], nazwa stylu, lista tak/nie, formatowanie fragmentów B/I/kolor),
// liczba przypisów i czy numery obrazów są jedyne. Bez parsera XML — działa w Node (wzorce) i w teście.
function summarize(docXml, stylesXml = "", footnotesXml = "") {
  const names = {};
  for (const m of String(stylesXml).matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>[\s\S]*?<w:name w:val="([^"]+)"/g)) names[m[1]] = m[2].toLowerCase();
  const body = String(docXml).slice(String(docXml).indexOf("<w:body>"));
  const lines = [];
  const paraRe = /<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  const tblRanges = [...body.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((m) => [m.index, m.index + m[0].length]);
  for (const m of body.matchAll(paraRe)) {
    const p = m[0];
    const inCell = tblRanges.some(([a, b]) => m.index >= a && m.index < b);
    const pPr = (p.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [""])[0].replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "");
    const sid = (pPr.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || "";
    const style = sid ? names[sid] || sid.toLowerCase() : "-";
    const list = /<w:numPr>/.test(pPr) ? "tak" : "-";
    const segs = [];
    for (const r of p.matchAll(/<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g)) {
      const rPr = (r[1].match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [""])[0];
      const on = (tag) => new RegExp(`<w:${tag}(?:\\s+w:val="(?!0|false)[^"]*")?\\s*/>`).test(rPr);
      const col = (rPr.match(/<w:color w:val="([^"]+)"/) || [])[1] || "";
      const fmt = (on("b") ? "B" : "") + (on("i") ? "I" : "") + (col && col.toLowerCase() !== "auto" ? col.toUpperCase() : "");
      for (const k of r[1].matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>|<w:t\b[^>]*\/>|<w:footnoteReference\b|<w:drawing\b|<w:tab\/>/g)) {
        const text = k[1] != null ? k[1] : k[0].startsWith("<w:footnote") ? "[PRZYPIS]" : k[0].startsWith("<w:drawing") ? "[OBRAZ]" : k[0] === "<w:tab/>" ? "\t" : "";
        const f = text.startsWith("[") ? "" : fmt;
        const last = segs[segs.length - 1];
        if (last && last[1] === f) last[0] += text; else segs.push([text, f]);
      }
    }
    const text = segs.map((s) => s[0]).join("").replace(/&amp;/g, "&").replace(/&quot;/g, '"');
    const fmts = segs.filter((s) => s[1] && s[0]).map((s) => `${s[0]}<${s[1]}>`).join("|");
    lines.push(`${inCell ? "[KOM] " : ""}${JSON.stringify(text)} styl=${style} lista=${list}${fmts ? ` fmt: ${fmts}` : ""}`);
  }
  const notes = [...String(footnotesXml).matchAll(/<w:footnote\b[^>]*w:id="(\d+)"/g)].filter((m) => +m[1] > 0).length;
  const docPr = [...String(docXml).matchAll(/<wp:docPr\b[^>]*?\sid="(\d+)"/g)].map((m) => m[1]);
  lines.push(`przypisów: ${notes}, odwołań: ${(String(docXml).match(/<w:footnoteReference\b/g) || []).length}, obrazów: ${docPr.length}, numery obrazów jedyne: ${docPr.length === new Set(docPr).size}`);
  return lines;
}
if (typeof module !== "undefined") module.exports = { summarize };
