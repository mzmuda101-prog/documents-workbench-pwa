// a11y-check.js — „Sprawdź ułatwienia dostępu” jak w Wordzie (Recenzja → Sprawdź ułatwienia
// dostępu), w panelu Korekta. Ładowany razem z Korektą (lazy-features.js).
//
// Sprawdzamy na podglądzie (to, co widać, z dokładnym miejscem do skoku) i w pliku (tytuł):
//   - obraz bez tekstu alternatywnego (błąd) — poprawka: okienko „Tekst alternatywny”,
//   - tabela bez wiersza nagłówka (ostrzeżenie) — poprawka: pierwszy wiersz jako nagłówek,
//   - pominięty poziom nagłówka, np. Nagłówek 1 → Nagłówek 3 (ostrzeżenie),
//   - kilka pustych akapitów z rzędu zamiast odstępu (wskazówka) — poprawka: zostaw jeden,
//   - niejasny tekst linku („kliknij tutaj”, sam długi adres) (ostrzeżenie),
//   - słaby kontrast tekstu z tłem (< 4,5:1, duży tekst < 3:1) (ostrzeżenie),
//   - brak tytułu dokumentu we właściwościach pliku (wskazówka) — poprawka: Metadane.
// Klik w uwagę = skok i podświetlenie miejsca (zdejmuje je Esc / „Odznacz” na dotyku).

const a11yScanBtn = document.getElementById("a11yScanBtn");
const a11yStatusEl = document.getElementById("a11yStatus");
const a11yResultsEl = document.getElementById("a11yResults");
const A11Y_MAX_PER_RULE = 40;
const A11Y_UNCLEAR_LINK = new Set(["tutaj", "kliknij", "kliknij tutaj", "kliknij tu", "tu", "link", "ten link", "więcej", "czytaj więcej", "here", "click here", "click", "link here", "more", "read more", "this link"]);
const A11Y_RULES = [
  ["imageAlt", "error"], ["tableHeader", "warning"], ["headingSkip", "warning"], ["linkText", "warning"],
  ["contrast", "warning"], ["emptyParas", "tip"], ["docTitle", "tip"],
];

function a11yHost() { return docCanvasEl?.querySelector(".docx-preview-host") || null; }
const a11yText = (el, n = 60) => { const s = (el?.textContent || "").replace(/\s+/g, " ").trim(); return s.length > n ? `${s.slice(0, n - 1)}…` : s; };

// ── kontrast (WCAG 2): jasność względna i stosunek ────────────────────────────
function a11yRgb(css) {
  const m = String(css || "").match(/rgba?\(\s*([\d.]+)[ ,]+([\d.]+)[ ,]+([\d.]+)(?:[ ,/]+([\d.]+))?/);
  return m ? { r: +m[1], g: +m[2], b: +m[3], a: m[4] === undefined ? 1 : +m[4] } : null;
}
function a11yLum({ r, g, b }) {
  const ch = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; };
  return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b);
}
function a11yRatio(fg, bg) {
  const a = a11yLum(fg), b = a11yLum(bg);
  return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
}
// Tło pod elementem: pierwsze nieprzezroczyste tło w górę (wyróżnienie, komórka, kartka); kartka = biała.
function a11yBackground(el) {
  for (let a = el; a && a !== document.body; a = a.parentElement) {
    const bg = a11yRgb(getComputedStyle(a).backgroundColor);
    if (bg && bg.a > 0.5) return bg;
    if (a.matches?.("section.docx")) break;
  }
  return { r: 255, g: 255, b: 255, a: 1 };
}

function a11yParaIndex(el) {
  const p = el?.closest?.("p");
  return p && typeof resolveParaIndex === "function" ? resolveParaIndex(p) : -1;
}

async function a11yDocTitle() {
  try {
    const zip = await JSZip.loadAsync(originalFileBytes);
    const core = zip.file("docProps/core.xml");
    if (!core) return "";
    const doc = new DOMParser().parseFromString(await core.async("string"), "application/xml");
    return (doc.getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "title")[0]?.textContent || "").trim();
  } catch (_) { return ""; }
}

async function runA11yCheck() {
  const host = a11yHost();
  if (!host || !originalFileBytes) { a11yStatusEl.textContent = t("noFileToSave"); return null; }
  const found = Object.fromEntries(A11Y_RULES.map(([id]) => [id, []]));
  const add = (rule, item) => { if (found[rule].length < A11Y_MAX_PER_RULE) found[rule].push(item); };
  const paras = typeof collectPreviewParagraphElements === "function" ? collectPreviewParagraphElements(host) : [...host.querySelectorAll("p")];

  // obrazy bez opisu: opis jest w pliku (wp:docPr descr / title) — podgląd go nie przenosi.
  // Obraz w podglądzie ↔ rysunek w pliku: ten sam akapit, ta sama kolejność (docImageTargets).
  // Oznaczony w Wordzie jako dekoracyjny (adec:decorative) — bez uwagi, jak w Wordzie.
  const docXml = typeof getDocumentXmlDom === "function" ? await getDocumentXmlDom(originalFileBytes).catch(() => null) : null;
  const xParas = docXml ? collectParagraphElements(docXml.documentElement, "all") : [];
  const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const imgInfo = (paraIndex, nth) => {
    const d = xParas[paraIndex]?.getElementsByTagNameNS(WP, "docPr")[nth];
    if (!d) return null;
    const decorative = [...d.getElementsByTagName("*")].some((n) => n.localName === "decorative" && /^(1|true)$/.test(n.getAttribute("val") || ""));
    return { described: !!(d.getAttribute("descr") || d.getAttribute("title") || "").trim(), decorative };
  };
  const targets = typeof docImageTargets === "function" ? docImageTargets(host) : [...host.querySelectorAll("img")].map((img) => ({ img, paraIndex: a11yParaIndex(img), nth: 0 }));
  targets.forEach(({ img, paraIndex, nth }) => {
    const info = imgInfo(paraIndex, nth);
    if (info ? info.described || info.decorative : (img.getAttribute("alt") || "").trim()) return;
    add("imageAlt", { el: img, label: t("a11yImageN", { n: found.imageAlt.length + 1 }), fix: "alt" });
  });

  // tabele bez wiersza nagłówka (pierwszy wiersz bez w:tblHeader — docx-preview nie zostawia
  // śladu w DOM, więc czytamy plik); tabele-układy (1 wiersz albo 1 kolumna) pomijamy jak Word
  if (docXml) {
    const tables = [...host.querySelectorAll("table")].filter((tb) => !tb.parentElement?.closest("table"));
    tables.forEach((tb) => {
      if (tb.rows.length < 2 || [...tb.rows].every((r) => r.cells.length < 2)) return;
      const firstP = tb.querySelector("td p, th p");
      const idx = a11yParaIndex(firstP);
      let tr = xParas[idx]?.parentNode;
      while (tr && tr.localName !== "tr") tr = tr.parentNode;
      const trPr = tr && [...tr.childNodes].find((n) => n.localName === "trPr");
      const header = trPr && [...trPr.childNodes].some((n) => n.localName === "tblHeader");
      const head = [...tb.rows[0].cells].map((c) => (c.textContent || "").replace(/\s+/g, " ").trim()).filter(Boolean).join(" · ");
      if (!header) add("tableHeader", { el: tb, label: (head.length > 50 ? `${head.slice(0, 49)}…` : head) || t("a11yTable"), fix: "header", paraIndex: idx });
    });
  }

  // kolejność nagłówków: skok o więcej niż jeden poziom w dół
  const structure = typeof analyzeDocumentDom === "function" ? analyzeDocumentDom(docCanvasEl) : null;
  let prev = 0;
  (structure?.headings || []).filter((h) => h.source !== "guess").forEach((h) => {
    if (prev && h.level > prev + 1) add("headingSkip", { el: h.el, label: `${t("a11yHeadingLevels", { from: prev, to: h.level })}: ${a11yText(h.el, 40)}` });
    prev = h.level;
  });

  // linki z niejasnym tekstem
  host.querySelectorAll("a[href]").forEach((a) => {
    if (a.classList.contains("doc-xref") || a.closest("header, footer")) return;
    const txt = (a.textContent || "").replace(/\s+/g, " ").trim();
    const low = txt.toLocaleLowerCase("pl-PL").replace(/[.:!…]+$/, "");
    if (A11Y_UNCLEAR_LINK.has(low) || (/^(https?:\/\/|www\.)\S{30,}$/i.test(txt))) add("linkText", { el: a, label: txt.length > 50 ? `${txt.slice(0, 49)}…` : txt });
  });

  // kontrast: tekst z własnym kolorem albo na kolorowym tle (jeden wpis na akapit)
  const contrastParas = new Set();
  paras.forEach((p) => {
    if (contrastParas.has(p)) return;
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (!/\S/.test(n.data)) continue;
      const el = n.parentElement;
      // tekst zastępczy pola formularza („Kliknij tutaj, aby wpisać.”) — podpowiedź, nie treść
      if (el.closest('[data-ff], [contenteditable="false"]')) continue;
      const cs = getComputedStyle(el);
      const fg = a11yRgb(cs.color);
      if (!fg) continue;
      const bg = a11yBackground(el);
      const px = parseFloat(cs.fontSize) || 14;
      const bold = (parseInt(cs.fontWeight, 10) || 400) >= 600;
      const large = px >= 24 || (bold && px >= 18.66);
      const ratio = a11yRatio(fg, bg);
      if (ratio < (large ? 3 : 4.5)) {
        contrastParas.add(p);
        add("contrast", { el: p, label: `${ratio.toFixed(1).replace(".", currentLang === "en" ? "." : ",")}:1 — ${a11yText(el, 40)}` });
        break;
      }
    }
  });

  // puste akapity z rzędu (poza tabelami): 2 i więcej → jedna uwaga na ciąg
  let run = [];
  const flush = () => { if (run.length >= 2) add("emptyParas", { el: run[0].p, label: t("a11yEmptyRun", { n: run.length }), fix: "empty", from: run[0].i, to: run[run.length - 1].i }); run = []; };
  paras.forEach((p, i) => {
    // akapit z zakładką (cel linku) albo zablokowany — nie jest „pustym odstępem”
    const empty = !p.closest("td, th, header, footer") && !/\S/.test((p.textContent || "").replace(/\uFEFF/g, "")) && !p.querySelector("img, svg, canvas, [contenteditable=false], [id]") && !p.dataset.lock;
    if (!empty) { flush(); return; }
    // ciąg tylko w obrębie jednej sekcji (granica sekcji to osobny element podglądu)
    if (run.length && (run[run.length - 1].i !== i - 1 || run[run.length - 1].p.closest("section") !== p.closest("section"))) flush();
    run.push({ p, i });
  });
  flush();

  if (!(await a11yDocTitle())) add("docTitle", { el: null, label: t("a11yNoTitle"), fix: "title" });
  return found;
}

function a11yClearMarks() {
  docCanvasEl?.querySelectorAll(".search-hit, .search-hit-active").forEach((n) => n.classList.remove("search-hit", "search-hit-active"));
}
function a11yJump(item) {
  if (!item.el?.isConnected) return;
  a11yClearMarks();
  const target = item.el.matches("img") ? item.el.closest("p") || item.el : item.el;
  target.classList.add("search-hit", "search-hit-active");
  target.scrollIntoView({ behavior: "smooth", block: "center" });
}

async function a11yFix(item) {
  if (item.fix === "title") {
    if (typeof setSidebarOpen === "function") setSidebarOpen(true);
    const panel = document.getElementById("panel-metadata");
    if (panel) { panel.open = true; panel.scrollIntoView({ block: "start", behavior: "smooth" }); }
    setTimeout(() => document.getElementById("metaTitle")?.focus(), 350);
    return;
  }
  if (readOnlyMode && typeof appFrame !== "undefined") appFrame.setReadOnly(false);
  if (item.fix === "alt" && typeof composeUi !== "undefined") {
    // po przełączeniu w Edycję podgląd jest rysowany od nowa — obraz wskazujemy kolejnością
    composeUi.openImageAlt([...a11yHost().querySelectorAll("img")].indexOf(item.el));
    return;
  }
  if (item.fix === "header" && item.paraIndex >= 0) {
    await applyDocumentEdit({ op: "table", index: item.paraIndex, action: "headerRow" });
  } else if (item.fix === "empty" && item.from >= 0 && item.to > item.from) {
    // zostaje jeden pusty akapit (jak Delete w zaznaczeniu od początku pierwszego do ostatniego)
    // zostaje OSTATNI z ciągu — może nieść znacznik końca sekcji (w:sectPr)
    await applyDocumentEdit({ op: "deleteRange", from: item.from, to: item.to, keep: "last", mergedRuns: [] });
  }
  await renderA11y(); // po poprawce lista od nowa
}

async function renderA11y() {
  if (!a11yResultsEl) return;
  a11yResultsEl.replaceChildren();
  a11yStatusEl.textContent = t("a11yWorking");
  const found = await runA11yCheck();
  if (!found) return;
  const total = Object.values(found).reduce((s, list) => s + list.length, 0);
  a11yStatusEl.textContent = total ? t("a11yFound", { n: total }) : t("a11yNone");
  A11Y_RULES.forEach(([rule, level]) => {
    const list = found[rule];
    if (!list.length) return;
    const details = document.createElement("details");
    details.className = `grammar-rule-group a11y-group a11y-${level}`;
    details.open = true;
    details.dataset.rule = rule;
    const summary = document.createElement("summary");
    summary.className = "grammar-rule-summary";
    summary.innerHTML = `<span class="a11y-level"></span><span class="grammar-rule-name"></span><span class="grammar-rule-count"></span>`;
    summary.querySelector(".a11y-level").textContent = t(`a11yLevel_${level}`);
    summary.querySelector(".grammar-rule-name").textContent = t(`a11yRule_${rule}`);
    summary.querySelector(".grammar-rule-count").textContent = `(${list.length})`;
    details.appendChild(summary);
    const why = document.createElement("p");
    why.className = "hint a11y-why";
    why.textContent = t(`a11yWhy_${rule}`);
    details.appendChild(why);
    const box = document.createElement("div");
    box.className = "fr-preview-list grammar-hit-list";
    list.forEach((item) => {
      const row = document.createElement("div");
      row.className = "grammar-hit-row";
      const b = document.createElement("button");
      b.type = "button";
      b.className = "fr-preview-item grammar-hit-item";
      b.textContent = item.label;
      b.disabled = !item.el;
      b.addEventListener("click", () => a11yJump(item));
      row.appendChild(b);
      if (item.fix) {
        const fix = document.createElement("button");
        fix.type = "button";
        fix.className = "btn grammar-apply-one";
        fix.dataset.fix = item.fix;
        fix.textContent = t(`a11yFix_${item.fix}`);
        fix.addEventListener("click", () => a11yFix(item));
        row.appendChild(fix);
      }
      box.appendChild(row);
    });
    details.appendChild(box);
    a11yResultsEl.appendChild(details);
  });
}

a11yScanBtn?.addEventListener("click", () => renderA11y());
