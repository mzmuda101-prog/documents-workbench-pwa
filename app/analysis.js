// Document structure analysis and search highlighting.

function getDocContentRoot(root) {
  const host = root?.querySelector(".docx-preview-host") || root;
  return host?.querySelector("section.docx") || host?.querySelector(".docx") || host;
}

function analyzeDocumentDom(root) {
  const content = getDocContentRoot(root);
  const text = content?.textContent || "";
  const words = text.trim() ? text.trim().split(/\s+/).filter(Boolean).length : 0;
  const allParas = Array.from(content?.querySelectorAll("p") || []);
  const headings = [];
  const headingEls = content?.querySelectorAll("h1, h2, h3, h4, h5, h6, p[style*='heading']") || [];
  headingEls.forEach((el) => {
    const label = (el.textContent || "").trim();
    if (!label) return;
    const paraIndex = allParas.indexOf(el);
    headings.push({ index: headings.length, paraIndex, label, el });
  });
  if (!headings.length && content) {
    allParas.forEach((el, i) => {
      const label = (el.textContent || "").trim();
      const cls = el.className || "";
      const isStyleHeading = /heading|docx-heading|Title|Subtitle/i.test(cls);
      if (label && (isStyleHeading || (label.length < 120 && /^[A-Z0-9ĄĆĘŁŃÓŚŹŻ]/.test(label)))) {
        headings.push({ index: headings.length, paraIndex: i, label, el });
      }
    });
  }
  const paragraphs = allParas.length;
  const tables = content?.querySelectorAll("table")?.length || 0;
  return { words, chars: text.length, headings, paragraphs, tables, text };
}

function renderStructurePanel(structure) {
  if (!structureSummaryEl || !headingNavEl) return;
  structureSummaryEl.replaceChildren();
  if (!structure) {
    headingNavEl.replaceChildren();
    return;
  }

  const stats = [
    [t("words"), structure.words],
    [t("chars"), structure.chars],
    [t("headings"), structure.headings.length],
    [t("paragraphs"), structure.paragraphs],
    [t("tables"), structure.tables],
  ];
  const grid = document.createElement("div");
  grid.className = "structure-stats";
  stats.forEach(([label, value]) => {
    const item = document.createElement("div");
    item.className = "structure-stat";
    item.innerHTML = `<span class="structure-stat-label">${label}</span><strong>${value}</strong>`;
    grid.appendChild(item);
  });
  structureSummaryEl.appendChild(grid);

  headingNavEl.replaceChildren();
  if (!structure.headings.length) return;
  const list = document.createElement("div");
  list.className = "heading-nav-list";
  structure.headings.slice(0, 40).forEach((h, idx) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "heading-nav-item";
    btn.textContent = h.label.length > 80 ? `${h.label.slice(0, 77)}…` : h.label;
    btn.addEventListener("click", () => {
      h.el?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    list.appendChild(btn);
  });
  headingNavEl.appendChild(list);
}
