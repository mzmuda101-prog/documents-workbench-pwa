// Document structure analysis and search highlighting.

function getDocContentRoot(root) {
  const host = root?.querySelector(".docx-preview-host") || root;
  return host?.querySelector("section.docx") || host?.querySelector(".docx") || host;
}

function analyzeDocumentDom(root) {
  const content = getDocContentRoot(root);
  const text = content?.textContent || "";
  const words = text.trim() ? text.trim().split(/\s+/).filter(Boolean).length : 0;
  const headings = [];
  const headingEls = content?.querySelectorAll("h1, h2, h3, h4, h5, h6, p[style*='heading']") || [];
  headingEls.forEach((el, i) => {
    const label = (el.textContent || "").trim();
    if (!label) return;
    headings.push({ index: i, label, el });
  });
  if (!headings.length && content) {
    content.querySelectorAll("p").forEach((el, i) => {
      const label = (el.textContent || "").trim();
      if (label && label.length < 120 && /^[A-Z0-9]/.test(label)) {
        headings.push({ index: i, label, el });
      }
    });
  }
  const paragraphs = content?.querySelectorAll("p")?.length || 0;
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

function clearSearchHighlights() {
  docCanvasEl?.querySelectorAll(".search-hit").forEach((el) => {
    el.classList.remove("search-hit", "search-hit-active");
  });
}

function runDocumentSearch() {
  clearSearchHighlights();
  searchMatches = [];
  activeSearchIndex = -1;
  const q = (searchQueryEl?.value || "").trim();
  if (!q || !docCanvasEl) {
    if (searchCountEl) searchCountEl.textContent = "";
    return;
  }
  const scope = searchScopeEl?.value || "all";
  const root = getDocContentRoot(docCanvasEl);
  const nodes = scope === "headings"
    ? root.querySelectorAll("h1, h2, h3, h4, h5, h6, p")
    : root.querySelectorAll("p, span, td, th, li");
  const lower = q.toLowerCase();
  nodes.forEach((el) => {
    const text = el.textContent || "";
    if (text.toLowerCase().includes(lower)) {
      el.classList.add("search-hit");
      searchMatches.push(el);
    }
  });
  if (searchCountEl) {
    searchCountEl.textContent = searchMatches.length
      ? t("searchMatches", { count: searchMatches.length })
      : t("searchNoMatches");
  }
  if (searchMatches.length) {
    activeSearchIndex = 0;
    focusSearchMatch(0);
  }
}

function focusSearchMatch(index) {
  if (!searchMatches.length) return;
  const i = ((index % searchMatches.length) + searchMatches.length) % searchMatches.length;
  searchMatches.forEach((el) => el.classList.remove("search-hit-active"));
  activeSearchIndex = i;
  const el = searchMatches[i];
  el.classList.add("search-hit-active");
  el.scrollIntoView({ behavior: "smooth", block: "center" });
}
