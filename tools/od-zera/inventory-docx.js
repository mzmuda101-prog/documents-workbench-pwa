// inventory-docx.js — co z Worda używa każdy plik .docx (do pytania „czy da się to zrobić od zera
// w aplikacji?”). Tylko odczyt; wynik: tabela cech na plik + suma. Pliki prywatne zostają na dysku.
//   node tools/od-zera/inventory-docx.js plik.docx folder/ …   (JSON=1 — wynik jako JSON)
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");

const files = process.argv.slice(2).flatMap((a) => (fs.statSync(a).isDirectory() ? fs.readdirSync(a).filter((f) => /\.docx$/i.test(f) && !f.startsWith("~$")).map((f) => path.join(a, f)) : [a]));
const count = (xml, re) => (xml.match(re) || []).length;

async function inspect(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const read = async (n) => (zip.file(n) ? zip.file(n).async("string") : "");
  const doc = await read("word/document.xml");
  const styles = await read("word/styles.xml");
  const numbering = await read("word/numbering.xml");
  const settings = await read("word/settings.xml");
  const names = Object.keys(zip.files);
  const hf = (await Promise.all(names.filter((n) => /^word\/(header|footer)\d*\.xml$/.test(n)).map(read))).join("");
  const all = doc + hf;
  const f = {};
  const set = (k, v) => { if (v) f[k] = v; };
  // tekst
  set("kolor tekstu", count(doc, /<w:color w:val="(?!auto|000000)/g));
  set("wyróżnienie / cieniowanie tekstu", count(doc, /<w:highlight |<w:rPr>[^]*?<w:shd /g) ? count(doc, /<w:highlight /g) || 1 : 0);
  set("kapitaliki / wersaliki", count(doc, /<w:(smallCaps|caps)\/>/g));
  set("odstęp między znakami", count(doc, /<w:rPr>(?:(?!<\/w:rPr>).)*<w:spacing w:val="-?[1-9]/g));
  set("indeks górny / dolny", count(doc, /<w:vertAlign w:val="(superscript|subscript)"/g));
  set("symbole (w:sym)", count(all, /<w:sym /g));
  // akapit
  set("obramowanie akapitu", count(doc, /<w:pBdr>/g));
  set("cieniowanie akapitu", count(doc, /<w:pPr>(?:(?!<\/w:pPr>).)*<w:shd [^>]*w:fill="(?!auto|FFFFFF)/g));
  set("tabulatory", count(doc, /<w:tab w:val="(left|right|center|decimal)"/g));
  set("tabulatory z wypełnieniem (kropki)", count(doc, /<w:tab [^>]*w:leader="(dot|hyphen|underscore)"/g));
  set("tabulator dziesiętny / prawy / środkowy", count(doc, /<w:tab w:val="(right|center|decimal)"/g));
  set("inicjał (drop cap)", count(doc, /w:dropCap=/g));
  set("ramka akapitu (framePr)", count(doc, /<w:framePr /g) - count(doc, /w:dropCap=/g));
  set("podział strony", count(doc, /<w:br w:type="page"/g) + count(doc, /<w:pageBreakBefore\/>/g));
  set("podział kolumny", count(doc, /<w:br w:type="column"/g));
  // listy
  set("listy", count(doc, /<w:numPr>/g));
  const fmts = new Set((numbering.match(/<w:numFmt w:val="([a-zA-Z]+)"/g) || []).map((s) => s.replace(/.*="|"/g, "")));
  set("formaty numeracji", [...fmts].filter((x) => x !== "none").join(", "));
  set("listy wielopoziomowe", count(doc, /<w:ilvl w:val="[1-8]"/g));
  // style
  const custom = count(styles, /<w:style [^>]*w:customStyle="1"/g);
  set("własne style", custom);
  set("style tabel", count(styles, /<w:style w:type="table"/g) > 1 ? count(styles, /<w:style w:type="table"/g) : 0);
  // tabele
  set("tabele", count(doc, /<w:tbl>/g));
  set("tabele zagnieżdżone", count(doc, /<w:tc>(?:(?!<\/w:tc>).)*<w:tbl>/gs));
  set("scalanie w poziomie", count(doc, /<w:gridSpan /g));
  set("scalanie w pionie", count(doc, /<w:vMerge/g));
  set("cieniowanie komórek", count(doc, /<w:tcPr>(?:(?!<\/w:tcPr>).)*<w:shd [^>]*w:fill="(?!auto|FFFFFF)/g));
  set("obramowania komórek", count(doc, /<w:tcBorders>/g));
  set("wyrównanie w pionie w komórce", count(doc, /<w:vAlign w:val="(center|bottom)"/g));
  set("kierunek tekstu w komórce", count(doc, /<w:textDirection /g));
  set("powtarzany wiersz nagłówka", count(doc, /<w:tblHeader\/>/g));
  set("tabela pływająca (tblpPr)", count(doc, /<w:tblpPr /g));
  // obiekty
  set("obraz w linii", count(doc, /<wp:inline/g));
  set("obraz / kształt pływający", count(doc, /<wp:anchor/g));
  const wraps = new Set((doc.match(/<wp:wrap(Square|Tight|Through|TopAndBottom|None)/g) || []).map((s) => s.slice(8)));
  set("zawijanie tekstu", [...wraps].join(", "));
  set("za tekstem", count(doc, /<wp:anchor [^>]*behindDoc="1"/g));
  set("pole tekstowe", count(all, /<w:txbxContent>/g));
  set("kształty (bez tekstu)", Math.max(0, count(all, /<wps:wsp>/g) - count(all, /<wps:txbx>/g)));
  set("grupa kształtów", count(all, /<wpg:wgp>/g));
  set("wykres", count(all, /<c:chart /g));
  set("SmartArt", count(all, /<dgm:relIds /g));
  set("obiekt OLE / osadzony", count(all, /<w:object /g));
  set("równanie", count(all, /<m:oMath>/g));
  set("VML (stary format)", count(all, /<w:pict>/g));
  set("znak wodny", count(hf, /PowerPlusWaterMarkObject|<v:textpath /g));
  // strona
  const sects = count(doc, /<w:sectPr/g);
  set("kilka sekcji", sects > 1 ? sects : 0);
  set("kolumny tekstu", count(doc, /<w:cols [^>]*w:num="[2-9]"/g));
  set("strona pozioma", count(doc, /<w:pgSz [^>]*w:orient="landscape"/g));
  set("obramowanie strony", count(doc, /<w:pgBorders/g));
  set("numeracja wierszy", count(doc, /<w:lnNumType /g));
  set("tło strony", count(doc, /<w:background /g));
  set("wyrównanie strony w pionie", count(doc, /<w:vAlign w:val="(center|both|bottom)"\/><\/w:sectPr>|<w:sectPr>(?:(?!<\/w:sectPr>).)*<w:vAlign /gs));
  // nagłówki
  set("nagłówki / stopki", names.filter((n) => /^word\/(header|footer)\d*\.xml$/.test(n)).length);
  set("inna pierwsza strona", count(doc, /<w:titlePg\/>/g));
  set("parzyste / nieparzyste", count(settings, /<w:evenAndOddHeaders\/>/g));
  // pola
  const instr = (all.match(/<w:instrText[^>]*>([^<]*)</g) || []).map((s) => s.replace(/<[^>]*>/g, "").trim().split(/\s+/)[0]).concat((all.match(/w:instr="\s*([A-Z]+)/g) || []).map((s) => s.replace(/.*"\s*/, "")));
  const fieldSet = [...new Set(instr.filter(Boolean).map((s) => s.toUpperCase()))];
  set("pola", fieldSet.join(", "));
  // odnośniki
  set("hiperłącza", count(doc, /<w:hyperlink /g));
  set("zakładki", count(doc, /<w:bookmarkStart [^>]*w:name="(?!_GoBack)/g));
  set("przypisy dolne", count(doc, /<w:footnoteReference /g));
  set("przypisy końcowe", count(doc, /<w:endnoteReference /g));
  set("komentarze", count(doc, /<w:commentReference /g));
  set("śledzone zmiany", count(doc, /<w:(ins|del) w:id/g));
  // formularze
  const sdt = new Set((doc.match(/<w:sdtPr>(?:(?!<\/w:sdtPr>).)*?<w(?:14)?:(date|dropDownList|comboBox|checkbox|text|richText|picture|docPartObj)\b/gs) || []).map((s) => s.match(/<w(?:14)?:(\w+)\b[^<]*$/)[1]));
  set("kontrolki treści", [...sdt].join(", ") || (count(doc, /<w:sdt>/g) ? "inne" : 0));
  set("pola formularza (stare)", count(doc, /<w:ffData>/g));
  return f;
}

(async () => {
  const out = {};
  for (const file of files) {
    try { out[path.basename(file)] = await inspect(file); } catch (e) { out[path.basename(file)] = { błąd: String(e.message || e) }; }
  }
  if (process.env.JSON) { console.log(JSON.stringify(out, null, 1)); return; }
  const total = {};
  for (const [name, f] of Object.entries(out)) {
    console.log(`\n## ${name}`);
    for (const [k, v] of Object.entries(f)) { console.log(`  ${k}: ${v}`); total[k] = (total[k] || 0) + 1; }
  }
  console.log(`\n## Ile plików używa (z ${files.length})`);
  Object.entries(total).sort((a, b) => b[1] - a[1]).forEach(([k, n]) => console.log(`  ${n}  ${k}`));
})();
