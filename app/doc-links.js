// doc-links.js — linki w dokumencie: spis treści, zakładki, odsyłacze, adresy WWW.
//
// Podgląd (docx-preview) rysuje <w:hyperlink> jako <a href> — wewnętrzny jako „#zakładka”
// (zakładka = <span id>), zewnętrzny jako adres. Bez obsługi kliknięcie robiło to, co
// przeglądarka: adres WWW ZASTĘPOWAŁ aplikację (niezapisane zmiany przepadały), a skok
// dopisywał #… do adresu. Tu:
//   - wewnętrzny link → płynny skok w dokumencie + „↩ Wróć” (też Alt+←), adres bez zmian,
//   - http(s) / mailto / tel → nowa karta (aplikacja zostaje),
//   - reszta (javascript:, plik na dysku, względny adres) → zablokowane z komunikatem,
//   - odsyłacze Worda jako POLA (REF / PAGEREF / NOTEREF z \h, HYPERLINK) — podgląd rysuje
//     tylko ich tekst; tu dostają opakowanie <a class="doc-xref">. Akapit z polem jest tylko do
//     odczytu, a zapis z podglądu pomija takie akapity — opakowanie nie trafia do pliku.

const LINK_SAFE_RE = /^(https?:|mailto:|tel:)/i;
const linkUi = { back: null, backTop: 0, backTimer: 0 };

// ── odsyłacze-pola z XML ─────────────────────────────────────────────────────
function linkFieldTarget(instr) {
  const s = String(instr || "").trim();
  let m = s.match(/^(?:PAGEREF|REF|NOTEREF)\s+"?([^\s"\\]+)"?(.*)$/i);
  if (m) return /\\h\b/i.test(m[2]) ? { href: `#${m[1]}` } : null; // bez \h Word też nie skacze
  m = s.match(/^HYPERLINK\s+(.*)$/i);
  if (!m) return null;
  const anchor = m[1].match(/\\l\s+"([^"]+)"/i);
  const url = m[1].match(/^"([^"]+)"/);
  if (url) return { href: url[1] + (anchor ? `#${anchor[1]}` : "") };
  return anchor ? { href: `#${anchor[1]}` } : null;
}

// Pola-odsyłacze w obrębie jednego akapitu: { paraIndex, before, text, href }.
function scanLinkFields(doc) {
  const out = [];
  const paraIndexOf = new Map();
  collectParagraphElements(doc.documentElement, "all").forEach((p, i) => paraIndexOf.set(p, i));
  const textBefore = (p, el) => paragraphTextNodes(p).filter((t) => el.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_PRECEDING).map((t) => t.textContent || "").join("");
  // proste: <w:fldSimple w:instr="…">wynik</w:fldSimple>
  Array.from(doc.getElementsByTagNameNS(W_NS, "fldSimple")).forEach((fs) => {
    const target = linkFieldTarget(fs.getAttributeNS(W_NS, "instr") || fs.getAttribute("w:instr"));
    const p = ffClosest(fs, "p");
    const text = ffText(fs);
    if (target && p && paraIndexOf.has(p) && text) out.push({ paraIndex: paraIndexOf.get(p), before: textBefore(p, fs), text, ...target });
  });
  // złożone: begin … instrText … separate … wynik … end (tylko w jednym akapicie)
  Array.from(doc.getElementsByTagNameNS(W_NS, "fldChar")).forEach((fc) => {
    if (ffAttr(fc, "fldCharType") !== "begin" || ffKid(fc, "ffData")) return;
    const beginRun = fc.parentNode;
    const p = ffClosest(beginRun, "p");
    if (!p || !paraIndexOf.has(p)) return;
    let instr = "";
    let sep = null;
    let end = null;
    let depth = 0;
    for (let n = beginRun.nextSibling; n; n = n.nextSibling) {
      if (n.nodeType !== 1 || n.localName !== "r") continue;
      const c = ffKid(n, "fldChar");
      if (c) {
        const type = ffAttr(c, "fldCharType");
        if (type === "begin") depth++;
        else if (type === "separate" && !depth) sep = n;
        else if (type === "end") { if (!depth) { end = n; break; } depth--; }
        continue;
      }
      if (!sep && !depth) ffKids(n, "instrText").forEach((it) => { instr += it.textContent || ""; });
    }
    const target = end && sep ? linkFieldTarget(instr) : null;
    if (!target) return;
    let text = "";
    for (let n = sep.nextSibling; n && n !== end; n = n.nextSibling) if (n.nodeType === 1) text += ffText(n);
    if (text) out.push({ paraIndex: paraIndexOf.get(p), before: textBefore(p, beginRun), text, ...target });
  });
  return out;
}

let linkScanBytes = null;
let linkFields = [];

async function paintDocLinks() {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host || !originalFileBytes) return;
  const bytes = originalFileBytes;
  if (linkScanBytes !== bytes) {
    try {
      const doc = await getDocumentXmlDom(bytes);
      linkFields = doc ? scanLinkFields(doc) : [];
    } catch (_) { linkFields = []; }
    if (bytes !== originalFileBytes) return;
    linkScanBytes = bytes;
  }
  const previews = collectPreviewParagraphElements(host);
  linkFields.forEach((lf, i) => {
    const p = previews[lf.paraIndex];
    if (!p || p.querySelector(`a.doc-xref[data-xref="${i}"]`)) return;
    const text = p.textContent || "";
    let at = text.indexOf(lf.text, Math.max(0, lf.before.length - 4));
    if (at < 0 || Math.abs(at - lf.before.length) > 12) at = text.indexOf(lf.text);
    const range = at >= 0 ? formDomRange(p, at, at + lf.text.length) : null;
    if (!range) return;
    const anc = range.commonAncestorContainer;
    if ((anc.nodeType === 1 ? anc : anc.parentElement)?.closest("a")) return; // już w linku (np. spis treści)
    const a = document.createElement("a");
    a.className = "doc-xref";
    a.dataset.xref = String(i);
    a.href = lf.href;
    try {
      a.appendChild(range.extractContents());
      range.insertNode(a);
    } catch (_) { /* zakres przez granice elementów — zostaje zwykły tekst */ }
  });
  // podpowiedzi: dokąd prowadzi link
  host.querySelectorAll("a[href]").forEach((a) => {
    if (a.dataset.hintPl) return;
    const href = a.getAttribute("href") || "";
    a.dataset.hint = "";
    if (href.startsWith("#")) {
      const target = linkTargetEl(href.slice(1));
      const label = (target?.closest("p")?.textContent || "").replace(/\s+/g, " ").trim().slice(0, 70);
      a.dataset.hintPl = label ? `Przejdź do: ${label}` : "Przejdź do miejsca w dokumencie";
      a.dataset.hintEn = label ? `Go to: ${label}` : "Go to a place in the document";
    } else if (LINK_SAFE_RE.test(href)) {
      a.dataset.hintPl = `Otwórz w nowej karcie: ${href}`;
      a.dataset.hintEn = `Open in a new tab: ${href}`;
    } else {
      a.dataset.hintPl = "Ten link nie otworzy się tutaj (plik na dysku lub nieobsługiwany adres)";
      a.dataset.hintEn = "This link can't be opened here (local file or unsupported address)";
    }
  });
}

function linkTargetEl(id) {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host || !id) return null;
  let decoded = id;
  try { decoded = decodeURIComponent(id); } catch (_) { /* zostaje surowe */ }
  return host.querySelector(`[id="${CSS.escape(decoded)}"]`);
}

// ── skok + „Wróć” ────────────────────────────────────────────────────────────
function hideLinkBack() {
  clearTimeout(linkUi.backTimer);
  if (linkUi.back) linkUi.back.hidden = true;
}
function showLinkBack(top) {
  if (!linkUi.back) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "link-back";
    b.addEventListener("click", () => linkGoBack());
    document.body.append(b);
    linkUi.back = b;
  }
  linkUi.backTop = top;
  linkUi.back.textContent = t("linkBack");
  linkUi.back.hidden = false;
  clearTimeout(linkUi.backTimer);
  linkUi.backTimer = setTimeout(hideLinkBack, 12000);
}
function linkGoBack() {
  if (!linkUi.back || linkUi.back.hidden) return false;
  docViewportEl?.scrollTo({ top: linkUi.backTop, behavior: "smooth" });
  hideLinkBack();
  return true;
}

function jumpToLinkTarget(id) {
  const target = linkTargetEl(id);
  if (!target) { toast(t("linkTargetMissing"), "info"); return; }
  const el = target.closest("p, td, li") || target;
  const before = docViewportEl?.scrollTop || 0;
  jumpToStructureItem({ el, id: "link" }, { silentSelect: true });
  setTimeout(() => el.classList.remove("search-hit", "search-hit-active"), 2200);
  showLinkBack(before);
  if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
}

function onDocLinkClick(e) {
  const a = e.target.closest?.("a[href]");
  if (!a || !docCanvasEl.contains(a) || e.button > 0) return;
  const sel = window.getSelection();
  if (sel && !sel.isCollapsed && a.contains(sel.anchorNode)) return; // zaznaczanie tekstu linku
  e.preventDefault(); // nigdy nie zastępuj aplikacji stroną z linku
  // Edycja: klik w link w edytowalnym akapicie stawia kursor (jak w Wordzie), Ctrl/⌘+klik
  // otwiera; karta linku (compose-ui.js) ma też „Otwórz”.
  if (a.closest(".docx-editable-p") && !(e.ctrlKey || e.metaKey)) return;
  const href = a.getAttribute("href") || "";
  if (href.startsWith("#")) { jumpToLinkTarget(href.slice(1)); return; }
  if (LINK_SAFE_RE.test(href)) { window.open(href, "_blank", "noopener,noreferrer"); return; }
  toast(t("linkUnsupported"), "info");
}

document.addEventListener("DOMContentLoaded", () => {
  docCanvasEl?.addEventListener("click", onDocLinkClick);
  // Alt+← jak w Wordzie — tylko gdy jest dokąd wracać (inaczej przeglądarka cofnęłaby stronę)
  document.addEventListener("keydown", (e) => {
    if (e.altKey && !e.ctrlKey && !e.metaKey && e.key === "ArrowLeft" && linkGoBack()) e.preventDefault();
  });
  const origRender = window.renderStructurePanel;
  if (typeof origRender === "function") {
    window.renderStructurePanel = function renderStructureAndLinks(...args) {
      const r = origRender.apply(this, args);
      hideLinkBack();
      paintDocLinks().catch(() => {});
      return r;
    };
  }
});
