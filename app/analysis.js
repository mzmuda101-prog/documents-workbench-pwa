// Document structure analysis and search highlighting.

const STRUCTURE_OUTLINE_LIMIT = 150;

function getDocContentRoot(root) {
  const host = root?.querySelector(".docx-preview-host") || root;
  return host?.querySelector("section.docx") || host?.querySelector(".docx") || host;
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

function collectDomHeadings(content, allParas) {
  const headings = [];
  const headingEls = content?.querySelectorAll("h1, h2, h3, h4, h5, h6, p[style*='heading']") || [];
  headingEls.forEach((el) => {
    const label = (el.textContent || "").trim();
    if (!label) return;
    const paraIndex = allParas.indexOf(el);
    headings.push({ index: headings.length, paraIndex, label, el, level: detectHeadingLevel(el) });
  });
  if (!headings.length && content) {
    allParas.forEach((el, i) => {
      const label = (el.textContent || "").trim();
      const cls = el.className || "";
      const isStyleHeading = /heading|docx-heading|Title|Subtitle/i.test(cls);
      if (label && (isStyleHeading || (label.length < 120 && /^[A-Z0-9ĄĆĘŁŃÓŚŹŻ]/.test(label)))) {
        headings.push({ index: headings.length, paraIndex: i, label, el, level: detectHeadingLevel(el) });
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

function analyzeDocumentDom(root) {
  const content = getDocContentRoot(root);
  const text = content?.textContent || "";
  const words = text.trim() ? text.trim().split(/\s+/).filter(Boolean).length : 0;
  const allParas = Array.from(content?.querySelectorAll("p") || []);
  const headings = collectDomHeadings(content, allParas);
  const outline = buildDocumentOutline(content, headings);
  const paragraphs = allParas.length;
  const tables = content?.querySelectorAll("table")?.length || 0;
  return { words, chars: text.length, headings, paragraphs, tables, text, outline };
}
