// Document structure analysis and search highlighting.

const STRUCTURE_OUTLINE_LIMIT = 150;

// Cały obszar dokumentu (wszystkie sekcje/strony). Akapity treści bierz przez
// docBodyParagraphs — bez nagłówków, stopek i listy przypisów.
function getDocContentRoot(root) {
  const host = root?.querySelector(".docx-preview-host") || root;
  return host?.querySelector(".docx-wrapper") || host?.querySelector("section.docx") || host;
}

function docBodyParagraphs(root) {
  const host = root?.querySelector?.(".docx-preview-host") || root;
  return typeof collectPreviewParagraphElements === "function" ? collectPreviewParagraphElements(host) : Array.from(host?.querySelectorAll("p") || []);
}

function isOutsideDocBody(node) {
  const tag = node.tagName;
  if (tag === "HEADER" || tag === "FOOTER") return true;
  return tag === "OL" && node.parentElement?.matches?.("section.docx");
}

function truncateOutlineText(text, max = 72) {
  const raw = String(text || "").replace(/\s+/g, " ").trim();
  if (!raw) return "";
  return raw.length > max ? `${raw.slice(0, max - 1)}…` : raw;
}

function detectHeadingLevel(el) {
  const tag = el.tagName?.toLowerCase() || "";
  const hm = tag.match(/^h([1-6])$/);
  if (hm) return Number(hm[1]);
  const cls = el.className || "";
  if (/Title/i.test(cls)) return 1;
  if (/Subtitle/i.test(cls)) return 2;
  const cm = cls.match(/heading-?(\d)/i) || cls.match(/docx-heading-?(\d)/i);
  if (cm) return Number(cm[1]);
  if (/heading|docx-heading/i.test(cls)) return 2;
  return 2;
}

// Style nagłówków z word/styles.xml → klasa CSS, którą docx-preview nadaje akapitowi.
// docx-preview renderuje nagłówki jako zwykłe <p class="docx_<styleId>"> (poziom konspektu
// czyta, ale nie używa), a styleId jest zlokalizowany: polski „Nagłówek 1” to „Nagwek1”.
// Dlatego poziom bierzemy z definicji stylu: w:outlineLvl albo nazwa „heading N” / „Title”
// (Word zapisuje nazwę po angielsku niezależnie od języka), z dziedziczeniem przez basedOn.
let docHeadingStyleClasses = new Map(); // "docx_nagwek1" → 1

function docxStyleClassName(styleId) {
  // ta sama zamiana co w docx-preview (escapeClassName) + prefiks „docx”
  return `docx_${String(styleId || "").replace(/[ .]+/g, "-").replace(/[&]+/g, "and").toLowerCase()}`;
}

async function readHeadingStyleClasses(bytes) {
  const map = new Map();
  if (!bytes || !window.JSZip) return map;
  try {
    const zip = await loadDocxZipCached(bytes);
    const xml = await zip.file("word/styles.xml")?.async("string");
    if (!xml) return map;
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const styles = new Map();
    Array.from(doc.getElementsByTagNameNS(W, "style")).forEach((st) => {
      if (st.getAttributeNS(W, "type") !== "paragraph") return;
      const id = st.getAttributeNS(W, "styleId");
      if (!id) return;
      const val = (tag) => st.getElementsByTagNameNS(W, tag)[0]?.getAttributeNS(W, "val") ?? null;
      const outline = val("outlineLvl");
      styles.set(id, { name: (val("name") || "").trim(), basedOn: val("basedOn"), outline: outline == null ? null : Number(outline) });
    });
    const levelOf = (id, depth = 0) => {
      const st = styles.get(id);
      if (!st || depth > 12) return 0;
      if (Number.isFinite(st.outline) && st.outline >= 0 && st.outline < 9) return st.outline + 1;
      const m = st.name.match(/^heading\s*(\d)$/i);
      if (m) return Number(m[1]);
      if (/^title$/i.test(st.name)) return 1;
      return st.basedOn ? levelOf(st.basedOn, depth + 1) : 0;
    };
    styles.forEach((_, id) => {
      const level = levelOf(id);
      if (level > 0) map.set(docxStyleClassName(id), Math.min(level, 6));
    });
  } catch (_) { /* uszkodzony styles.xml — zostaje wykrywanie z DOM */ }
  return map;
}

function styleHeadingLevel(el) {
  if (!docHeadingStyleClasses.size) return 0;
  for (const cls of el.classList) {
    const level = docHeadingStyleClasses.get(cls);
    if (level) return level;
  }
  return 0;
}

function collectDomHeadings(content, allParas) {
  const headings = [];
  // 1) style nagłówków z pliku — najpewniejsze źródło (działa też dla polskiego Worda)
  allParas.forEach((el, i) => {
    const level = styleHeadingLevel(el);
    const label = (el.textContent || "").trim();
    if (level && label) headings.push({ index: headings.length, paraIndex: i, label, el, level, source: "style" });
  });
  if (headings.length) return headings;
  const headingEls = content?.querySelectorAll("h1, h2, h3, h4, h5, h6, p[style*='heading']") || [];
  headingEls.forEach((el) => {
    const label = (el.textContent || "").trim();
    if (!label) return;
    const paraIndex = allParas.indexOf(el);
    headings.push({ index: headings.length, paraIndex, label, el, level: detectHeadingLevel(el), source: "tag" });
  });
  if (!headings.length && content) {
    allParas.forEach((el, i) => {
      const label = (el.textContent || "").trim();
      const cls = el.className || "";
      const isStyleHeading = /heading|docx-heading|Title|Subtitle/i.test(cls);
      if (label && (isStyleHeading || (label.length < 120 && /^[A-Z0-9ĄĆĘŁŃÓŚŹŻ]/.test(label)))) {
        // zgadywanie (krótki akapit od wielkiej litery) — dobre dla Struktury, za słabe na skróty sekcji
        headings.push({ index: headings.length, paraIndex: i, label, el, level: detectHeadingLevel(el), source: isStyleHeading ? "tag" : "guess" });
      }
    });
  }
  return headings;
}

function buildDocumentOutline(content, headings) {
  if (!content) return [];
  const headingByPara = new Map();
  (headings || []).forEach((h) => {
    if (Number.isFinite(h.paraIndex) && h.paraIndex >= 0) headingByPara.set(h.paraIndex, h);
  });

  const outline = [];
  let paraIndex = 0;
  let tableIndex = 0;
  const walker = document.createTreeWalker(content, NodeFilter.SHOW_ELEMENT, {
    acceptNode(node) {
      if (isOutsideDocBody(node)) return NodeFilter.FILTER_REJECT; // nagłówki, stopki, przypisy
      const tag = node.tagName;
      if (tag === "P" || tag === "TABLE") return NodeFilter.FILTER_ACCEPT;
      return NodeFilter.FILTER_SKIP;
    },
  });

  let node = walker.nextNode();
  while (node) {
    if (node.tagName === "P") {
      const headingMeta = headingByPara.get(paraIndex);
      const isHeading = !!headingMeta;
      const raw = (node.textContent || "").replace(/\s+/g, " ").trim();
      const label = raw || (typeof t === "function" ? t("structureEmptyPara") : "(empty)");
      outline.push({
        id: `p-${paraIndex}`,
        type: isHeading ? "heading" : "paragraph",
        paraIndex,
        tableIndex: null,
        level: isHeading ? (headingMeta.level || detectHeadingLevel(node)) : 0,
        label,
        preview: truncateOutlineText(label),
        el: node,
      });
      paraIndex++;
    } else if (node.tagName === "TABLE") {
      const rows = node.querySelectorAll("tr").length;
      const cols = node.querySelectorAll("tr:first-child th, tr:first-child td").length;
      const firstCell = (node.querySelector("td, th")?.textContent || "").replace(/\s+/g, " ").trim();
      const preview = firstCell ? truncateOutlineText(firstCell, 48) : "";
      const label = typeof t === "function"
        ? (preview ? t("structureTablePreview", { rows, cols, preview }) : t("structureTableEmpty", { rows, cols }))
        : `Table ${rows}×${cols}`;
      outline.push({
        id: `t-${tableIndex}`,
        type: "table",
        paraIndex: null,
        tableIndex,
        level: 0,
        label,
        preview: truncateOutlineText(label, 80),
        el: node,
        rows,
        cols,
      });
      tableIndex++;
    }
    node = walker.nextNode();
  }
  return outline;
}

// ── Statystyka wyrazów jak w Wordzie ──────────────────────────────────────────
// Sprawdzone z Word 16.114 (tools/word-oracle/run-stats.js, 2026-10-17). Dawniej pasek stanu
// liczył słowa z textContent całego artykułu (akapity i komórki tabel sklejone: „kota.Kot” =
// jedno słowo), a panel Statystyki tylko „wyrazy z liter” (bez „3,5”, „1/2”, „—” osobno…),
// znaki z UTF-16 (emoji = 2) i z dopisanym „\n” między akapitami. Word:
//   • słowo = ciąg między odstępami (spacja, tabulator, łamanie wiersza, twarda spacja, spacja
//     ideograficzna); „—”, „–”, „•” (też punktor Symbol U+F0B7) dzielą, ale same słowem nie są; reszta interpunkcji
//     stojąca osobno (np. „,” albo „...”) — jest słowem; cienka spacja (U+2009) nie dzieli,
//   • numer listy („1.”, „a)”) to tekst akapitu: słowo i znaki (+ tabulator po nim),
//   • ukryty tekst (w:vanish) się nie liczy, znaki zerowej szerokości też nie,
//   • znak = punkt kodowy (emoji = 1), „ze spacjami” bez znaków końca akapitu,
//   • akapity = tylko z jakimś znakiem innym niż odstęp (puste i same spacje — nie).
const WORD_SPLIT_RE = /[ \t\n\r\f\u0001\u000b\u00a0\u3000\u2013\u2014\u2022\uf0a8\uf0b7]+/; // F0B7/F0A8 = punktory z kroju Symbol
const WORD_SPACE_RE = /[ \t\n\r\u000b\u00a0\u3000]/g;
const WORD_ZERO_WIDTH_RE = /[\u200b-\u200d\u2060\ufeff]/g;

// tekst akapitu podglądu tak, jak widzi go Word: numer listy, łamania, tabulatory; bez ukrytego
function wordParagraphText(p) {
  let out = "";
  // numer listy z suffiksem (tabulator / spacja) albo punktor — list-numbering.js
  const marker = p.getAttribute("data-dwb-num") ?? p.getAttribute("data-dwb-bullet");
  if (marker != null) out += marker.replace(/\u00a0$/, " ");
  const walk = (node) => {
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3) out += n.data;
      else if (n.nodeType === 1) {
        if (n.tagName === "BR") { if (!n.dataset.dwbPh) out += "\n"; continue; }
        if (n.dataset?.dwbNote) { out += "\u0001"; continue; } // znak przypisu: 1 znak, dzieli słowa
        if (n.style?.display === "none" || n.hidden) continue; // ukryty tekst (w:vanish)
        if (n.classList.contains("docx-tab")) { out += "\t"; continue; }
        walk(n);
      }
    }
  };
  walk(p);
  return out.replace(WORD_ZERO_WIDTH_RE, "");
}

function wordTextStats(texts) {
  let words = 0;
  let chars = 0;
  let charsWithSpaces = 0;
  let paragraphs = 0;
  texts.forEach((raw) => {
    // \f = podział strony/kolumny (dzieli słowa, nie jest znakiem), \u0001 = znak przypisu
    const text = String(raw || "").replace(WORD_ZERO_WIDTH_RE, "");
    const breaks = (text.match(/\f/g) || []).length;
    const all = [...text].length - breaks;
    const spaces = (text.match(WORD_SPACE_RE) || []).length;
    charsWithSpaces += all;
    chars += all - spaces;
    if (all > spaces) paragraphs++;
    words += text.split(WORD_SPLIT_RE).filter(Boolean).length;
  });
  return { words, chars, charsWithSpaces, paragraphs };
}

// cały dokument: akapity treści (z komórkami tabel), bez nagłówków/stopek i przypisów
// (akapit przecięty podziałem strony = dwa <p> w podglądzie, druga połówka z data-dwb-cont — jeden akapit)
function documentWordStats(root) {
  const host = root?.querySelector?.(".docx-preview-host") || root;
  const texts = [];
  Array.from(host?.querySelectorAll("section.docx > article p") || []).forEach((p) => {
    if (p.closest(".dwb-notes, header, footer")) return;
    if (p.dataset.dwbCont && texts.length) texts[texts.length - 1] += `\f${wordParagraphText(p)}`;
    else texts.push(wordParagraphText(p));
  });
  return wordTextStats(texts);
}

function analyzeDocumentDom(root) {
  const content = getDocContentRoot(root);
  const allParas = docBodyParagraphs(root);
  const articles = content?.querySelectorAll?.("section.docx > article");
  const text = articles?.length ? Array.from(articles).map((a) => a.textContent || "").join("\n") : content?.textContent || "";
  const counts = documentWordStats(root); // jak „Statystyka wyrazów” Worda
  const words = counts.words;
  const headings = collectDomHeadings(content, allParas);
  const outline = buildDocumentOutline(content, headings);
  const paragraphs = counts.paragraphs; // jak Word: tylko akapity z treścią (puste się nie liczą)
  const tables = (articles?.length ? Array.from(articles).reduce((n, a) => n + a.querySelectorAll("table").length, 0) : content?.querySelectorAll("table")?.length) || 0;
  return { words, chars: counts.charsWithSpaces, counts, headings, paragraphs, tables, text, outline };
}
