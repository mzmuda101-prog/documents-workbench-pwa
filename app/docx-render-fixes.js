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

// Odstęp między wierszami jak w Wordzie. W Wordzie „1,15” (w:spacing line=276 auto) znaczy
// 1,15 × NATURALNA wysokość wiersza czcionki (dla Arial ok. 1,15 em, Calibri 1,22 em), a
// docx-preview pisze `line-height: 1.15` = 1,15 em — wiersze wychodziły o kilka procent
// ciaśniej, a do tego akapit miał 16 px dziedziczone z aplikacji, więc jego „strut” narzucał
// 18,4 px niezależnie od liter (Arial 11 pt w Wordzie ≈ 19,3 px). Skutek: strona mieściła
// więcej tekstu niż w Wordzie, granice stron (page-breaks.js) wypadały za późno.
// Poprawka: akapit dostaje wielkość liter swojego pierwszego fragmentu tekstu (jak znak
// akapitu w Wordzie) i line-height = mnożnik × współczynnik czcionki. Dokładny odstęp
// (exact/atLeast, w pt) zostaje bez zmian — rozpoznajemy go po tym, że NIE skaluje się
// z wielkością liter.
const WORD_LINE_FACTORS = {
  arial: 1.149, helvetica: 1.149, "liberation sans": 1.149, "arial narrow": 1.149,
  calibri: 1.2207, "calibri light": 1.2207, cambria: 1.1724, aptos: 1.2, "aptos display": 1.2,
  "times new roman": 1.149, times: 1.149, georgia: 1.1362, garamond: 1.12, "book antiqua": 1.17,
  verdana: 1.2153, tahoma: 1.2075, "segoe ui": 1.33, "trebuchet ms": 1.1641,
  "courier new": 1.1328, consolas: 1.1709, "century gothic": 1.2251,
};

function wordLineFactor(fontFamily) {
  const first = String(fontFamily || "").split(",")[0].trim().replace(/^["']|["']$/g, "").toLowerCase();
  return WORD_LINE_FACTORS[first] || 1.17;
}

function applyWordLineMetrics(host) {
  const root = host?.querySelector?.(".docx-wrapper") || host;
  if (!root) return 0;
  const multiplierCache = new Map(); // klasa + inline line-height → mnożnik albo null (dokładny)
  let fixed = 0;
  root.querySelectorAll("section.docx p").forEach((p) => {
    const key = `${p.className}|${p.style.lineHeight}`;
    let m = multiplierCache.get(key);
    if (m === undefined) {
      const prevFs = p.style.fontSize;
      p.style.fontSize = "100px";
      const a = parseFloat(getComputedStyle(p).lineHeight);
      p.style.fontSize = "200px";
      const b = parseFloat(getComputedStyle(p).lineHeight);
      p.style.fontSize = prevFs;
      m = Number.isFinite(a) && Number.isFinite(b) && Math.abs(b - 2 * a) < 1 ? a / 100 : null;
      multiplierCache.set(key, m);
    }
    if (!m) return; // dokładny odstęp w pt albo „normal” — zostaje
    const run = [...p.querySelectorAll("span")].find((s) => s.textContent.trim()) || p.querySelector("span") || p;
    const cs = getComputedStyle(run);
    const fs = parseFloat(cs.fontSize);
    if (!fs) return;
    p.style.fontSize = `${fs}px`;
    p.style.lineHeight = String(Math.round(m * wordLineFactor(cs.fontFamily) * 1000) / 1000);
    fixed++;
  });
  return fixed;
}
