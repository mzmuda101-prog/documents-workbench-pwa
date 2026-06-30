// Post-render fixes for docx-preview output (Symbol bullets, visual tofu, etc.).

function getPreviewSection(host) {
  const root = host?.querySelector?.(".docx-preview-host") || host;
  return root?.querySelector("section.docx") || root?.querySelector(".docx") || root;
}

function isNumberedListMarker(beforeStyle) {
  const content = beforeStyle?.content || "";
  return content.includes("counter(");
}

function isImageBullet(beforeStyle) {
  const bg = beforeStyle?.backgroundImage || "";
  return bg && bg !== "none";
}

function usesLegacyBulletFont(beforeStyle) {
  const font = beforeStyle?.fontFamily || "";
  return /Symbol|Wingdings|Webdings|MT Extra|Marlett/i.test(font);
}

function fixDocxBulletRendering(host) {
  const section = getPreviewSection(host);
  if (!section) return 0;
  let fixed = 0;
  section.querySelectorAll('p[class*="docx-num-"]').forEach((p) => {
    const paraStyle = getComputedStyle(p);
    if (paraStyle.display !== "list-item") return;
    const before = getComputedStyle(p, "::before");
    if (isImageBullet(before)) return;

    if (isNumberedListMarker(before)) {
      p.classList.add("docx-list-numbered-fixed");
      fixed++;
      return;
    }
    if (!usesLegacyBulletFont(before) && before.content === "none") return;

    const levelMatch = p.className.match(/docx-num-\d+-(\d+)/);
    const level = levelMatch ? parseInt(levelMatch[1], 10) : 0;
    p.classList.add("docx-bullet-fixed", `docx-bullet-l${level % 3}`);
    fixed++;
  });
  return fixed;
}

function auditDocxVisualIssues(host) {
  const section = getPreviewSection(host);
  const issues = [];
  if (!section) return { issues, bulletsFixed: 0, brokenBullets: 0, numberedFixed: 0 };

  let brokenBullets = 0;
  let bulletsFixed = 0;
  let numberedFixed = 0;
  section.querySelectorAll('p[class*="docx-num-"]').forEach((p) => {
    const before = getComputedStyle(p, "::before");
    if (isImageBullet(before)) return;

    if (isNumberedListMarker(before)) {
      if (p.classList.contains("docx-list-numbered-fixed")) {
        numberedFixed++;
        const ti = parseFloat(getComputedStyle(p).textIndent) || 0;
        if (ti < 0) issues.push({ type: "numbered-indent", className: p.className, textIndent: ti });
      }
      return;
    }

    if (!usesLegacyBulletFont(before) && before.content === "none") return;
    if (p.classList.contains("docx-bullet-fixed")) {
      bulletsFixed++;
      const ti = parseFloat(getComputedStyle(p).textIndent) || 0;
      if (ti < 0) issues.push({ type: "bullet-indent", className: p.className, textIndent: ti });
      return;
    }
    brokenBullets++;
    issues.push({
      type: "bullet-tofu",
      className: p.className,
      font: before.fontFamily,
      content: before.content,
    });
  });

  return { issues, bulletsFixed, brokenBullets, numberedFixed };
}
