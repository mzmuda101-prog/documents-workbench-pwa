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
const linkUi = { back: null, backTop: 0, backTimer: 0, backToken: 0, backArmed: false };
// użytkownik sam przewija — dociągnięcie „Wróć” odpada
["wheel", "touchstart", "pointerdown"].forEach((type) => document.addEventListener(type, (e) => { if (docViewportEl?.contains(e.target)) linkUi.backToken++; }, { passive: true, capture: true }));

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
  // podpowiedzi: dokąd prowadzi link (link do obrazu — zamiast napisu podgląd obrazu, niżej)
  host.querySelectorAll("a[href]").forEach((a) => {
    if (a.dataset.hintPl || a.dataset.dwbImgLink) return;
    const href = a.getAttribute("href") || "";
    if (href.startsWith("#") && linkTargetImage(href.slice(1))) { a.dataset.dwbImgLink = "1"; return; }
    a.dataset.hint = "";
    // etykietka ekranowa z pliku (w:tooltip) zamiast adresu — jak w Wordzie
    if (a.dataset.dwbTip) {
      a.dataset.hintPl = a.dataset.dwbTip;
      a.dataset.hintEn = a.dataset.dwbTip;
    } else if (href === "#_top") {
      a.dataset.hintPl = "Przejdź na początek dokumentu";
      a.dataset.hintEn = "Go to the top of the document";
    } else if (href.startsWith("#")) {
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

// Zakładki dokumentu (Word: Wstaw → Zakładka) z pliku, w kolejności dokumentu:
// [{ name, paraIndex, hidden (nazwa od „_” — ukryta, jak _Toc/_Ref Worda), text (objęty tekst), links }]
// links = ile linków i odsyłaczy (REF/PAGEREF/HYPERLINK \l) prowadzi do zakładki.
async function docBookmarks(bytes = originalFileBytes) {
  const doc = bytes ? await getDocumentXmlDom(bytes) : null;
  if (!doc) return [];
  const paras = collectParagraphElements(doc.documentElement, "all");
  const paraIndexOf = new Map(paras.map((p, i) => [p, i]));
  const refs = new Map();
  const addRef = (name) => refs.set(name, (refs.get(name) || 0) + 1);
  Array.from(doc.getElementsByTagNameNS(W_NS, "hyperlink")).forEach((h) => { const a = h.getAttributeNS(W_NS, "anchor"); if (a) addRef(a); });
  scanLinkFields(doc).forEach((f) => { if (f.href?.startsWith("#")) addRef(f.href.slice(1)); });
  const ends = new Map(Array.from(doc.getElementsByTagNameNS(W_NS, "bookmarkEnd")).map((e) => [e.getAttributeNS(W_NS, "id"), e]));
  const out = [];
  Array.from(doc.getElementsByTagNameNS(W_NS, "bookmarkStart")).forEach((b) => {
    const name = b.getAttributeNS(W_NS, "name") || "";
    if (!name) return;
    // akapit zakładki: jej akapit albo pierwszy akapit za nią (zakładka między akapitami / nad tabelą)
    let p = ffClosest(b, "p");
    if (!p) p = paras.find((x) => b.compareDocumentPosition(x) & Node.DOCUMENT_POSITION_FOLLOWING) || null;
    let text = "";
    const end = ends.get(b.getAttributeNS(W_NS, "id"));
    if (end) {
      const walker = doc.createTreeWalker(doc.documentElement, NodeFilter.SHOW_ELEMENT);
      walker.currentNode = b;
      for (let n = walker.nextNode(), k = 0; n && n !== end && text.length < 80 && k < 4000; n = walker.nextNode(), k++) {
        if (n.localName === "t" && n.namespaceURI === W_NS) text += n.textContent || "";
        else if (n.localName === "p" && text) text += " ";
      }
    }
    out.push({ name, paraIndex: p && paraIndexOf.has(p) ? paraIndexOf.get(p) : -1, hidden: name.startsWith("_"), text: text.replace(/\s+/g, " ").trim(), links: refs.get(name) || 0 });
  });
  return out;
}

// Ta sama lista z pamięci, gdy plik się nie zmienił — okienka „Link” i „Zakładka” budują się od
// razu (na iOS fokus pola musi paść w geście stuknięcia, inaczej nie wysuwa się klawiatura).
const docBookmarksCache = { bytes: null, list: null, pending: null };
function docBookmarksNow() {
  if (docBookmarksCache.bytes === originalFileBytes && docBookmarksCache.list) return docBookmarksCache.list;
  return null;
}
function docBookmarksFresh() {
  const bytes = originalFileBytes;
  if (docBookmarksCache.bytes === bytes && docBookmarksCache.list) return Promise.resolve(docBookmarksCache.list);
  if (docBookmarksCache.pending?.bytes === bytes) return docBookmarksCache.pending.p;
  const p = docBookmarks(bytes).catch(() => []).then((list) => {
    if (bytes === originalFileBytes) Object.assign(docBookmarksCache, { bytes, list });
    return list;
  });
  docBookmarksCache.pending = { bytes, p };
  return p;
}

function linkTargetEl(id) {
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host || !id) return null;
  let decoded = id;
  try { decoded = decodeURIComponent(id); } catch (_) { /* zostaje surowe */ }
  return host.querySelector(`[id="${CSS.escape(decoded)}"]`);
}

// ── link do obrazu: podgląd bez skoku ────────────────────────────────────────
// Word: link „Miejsce w tym dokumencie” do zakładki przy rysunku albo odsyłacz \h do podpisu
// („Rysunek 2”) — klik przenosi do obrazu. U nas dodatkowo: najechanie myszą (na dotyku
// przytrzymanie palca) pokazuje obraz w okienku przy linku — bez zmiany miejsca w dokumencie.
// Klik dalej skacze do obrazu z „↩ Wróć”.

// Podpis obrazu: akapit pod nim (styl Legenda / „Rysunek 1…”), inaczej tekst alternatywny, inaczej „Obraz N”.
const DOC_CAPTION_RE = /^(rys(unek|\.)?|ilustracja|zdj(ęcie|\.)|fot(\.|ografia)?|wykres|schemat|fig(ure|\.)?|image|picture|chart)\s*\d/i;
function docImageCaption(p) {
  for (const sib of [p.nextElementSibling, p.previousElementSibling]) {
    if (sib?.tagName !== "P" || sib.querySelector("img")) continue;
    const text = (sib.textContent || "").replace(/\s+/g, " ").trim();
    const styled = Array.from(sib.classList).some((c) => /caption|legenda/i.test(c));
    if (text && text.length <= 160 && (styled || DOC_CAPTION_RE.test(text))) return text;
  }
  return "";
}
// Obrazy dokumentu w kolejności pliku, KAŻDY osobno: także kilka obrazów w jednym akapicie i obraz
// w drugiej połówce akapitu przeciętego podziałem strony (data-dwb-cont — to dalej ten sam akapit
// pliku). Dawniej brany był tylko pierwszy obraz akapitu, więc np. obraz za łamaniem wiersza albo
// za podziałem strony w ogóle nie trafiał na listę celów linku (4 obrazy w pliku, 3 na liście).
// → [{ img, part (<p> z obrazem), p (akapit pliku), paraIndex, nth (który obraz w akapicie, od 0), count }]
// nth liczy tak samo jak zapis (docx-compose.js: composeImageRuns — fragmenty z a:blip / v:imagedata).
function docParagraphParts(host) {
  const out = [];
  let main = null;
  let index = -1;
  (host?.querySelectorAll("section.docx > article p") || []).forEach((part) => {
    if (part.hasAttribute("data-dwb-cont")) { if (!main) return; } else { main = part; index++; }
    out.push({ part, p: main, paraIndex: index });
  });
  return out;
}
function docImageTargets(host) {
  const out = [];
  const perPara = new Map();
  docParagraphParts(host).forEach(({ part, p, paraIndex }) => {
    part.querySelectorAll("img").forEach((img) => {
      if (img.closest("p") !== part) return; // obraz w akapicie zagnieżdżonym (pole tekstowe) liczy jego akapit
      const nth = perPara.get(p) || 0;
      perPara.set(p, nth + 1);
      out.push({ img, part, p, paraIndex, nth });
    });
  });
  out.forEach((x) => { x.count = perPara.get(x.p); });
  return out;
}
// Podpis obrazu: podpis pod/nad akapitem albo tekst akapitu — tylko gdy obraz jest w akapicie sam
// (przy kilku obrazach w akapicie nie wiadomo, do którego należy); dalej tekst alternatywny, „Obraz N”.
function docImageLabel(p, n, img = p?.querySelector?.("img"), count = 1) {
  const own = count > 1 ? "" : (p?.textContent || "").replace(/\s+/g, " ").trim();
  return (count > 1 ? "" : docImageCaption(p)) || own || (img?.getAttribute("alt") || "").trim() || (n ? t("linkImageN", { n }) : t("linkGroupImages"));
}
// Obraz, do którego prowadzi zakładka: pierwszy obraz ZA zakładką w tym samym akapicie pliku
// (zakładka stoi tuż przed obrazem albo na początku akapitu z obrazem) albo obraz, którego
// podpisem jest cel (odsyłacz Worda do „Rysunek 2” wskazuje podpis pod obrazem).
// Wołane dla każdego linku przy rysowaniu — tylko części jednego akapitu, bez liczenia całego dokumentu.
function linkTargetImage(id) {
  const el = linkTargetEl(id);
  const p = el?.closest?.("p");
  if (!p) return null;
  const all = Array.from(docCanvasEl?.querySelectorAll(".docx-preview-host section.docx > article p") || []);
  let i = all.indexOf(p);
  while (i > 0 && all[i].hasAttribute("data-dwb-cont")) i--;
  const parts = [];
  for (let k = Math.max(0, i); k < all.length && (k === i || all[k].hasAttribute("data-dwb-cont")); k++) parts.push(all[k]);
  const imgs = parts.flatMap((part) => Array.from(part.querySelectorAll("img")).filter((img) => img.closest("p") === part));
  const own = imgs.find((img) => el.compareDocumentPosition(img) & Node.DOCUMENT_POSITION_FOLLOWING);
  if (own) return { img: own, p: own.closest("p"), count: imgs.length };
  const text = (p.textContent || "").replace(/\s+/g, " ").trim();
  for (const sib of [p.previousElementSibling, p.nextElementSibling]) {
    const img = sib?.tagName === "P" ? sib.querySelector("img") : null;
    if (img && text && docImageCaption(sib) === text) return { img, p: sib, count: 1 };
  }
  return null;
}
// Podpis obrazu-celu do okienka podglądu — ten sam co na liście celów linku („Obraz N” = N-ty w dokumencie).
function linkTargetImageLabel(hit) {
  const targets = docImageTargets(docCanvasEl?.querySelector(".docx-preview-host"));
  const n = targets.findIndex((x) => x.img === hit.img) + 1;
  return docImageLabel(hit.p, n || undefined, hit.img, hit.count);
}

const linkPeek = { el: null, a: null, img: null, timer: 0, hideTimer: 0, touchTimer: 0, touchShown: false, pending: null };
function hideLinkPeek() {
  clearTimeout(linkPeek.timer);
  clearTimeout(linkPeek.hideTimer);
  linkPeek.pending = null; // przerwane czekanie (przewinięcie, klik) — następny ruch zaczyna od nowa
  linkPeek.el?.remove();
  linkPeek.el = null;
  linkPeek.a = null;
  linkPeek.img = null;
}
// Mysz zjeżdża z linku na okienko (przycisk „Powiększ”) — chwila zwłoki, żeby nie znikło po drodze.
function hideLinkPeekSoon() {
  clearTimeout(linkPeek.hideTimer);
  linkPeek.hideTimer = setTimeout(hideLinkPeek, 260);
}
// Podgląd obrazu na cały / pół ekranu prosto z linku — bez skoku do obrazu w dokumencie.
function zoomLinkTarget(a) {
  const hit = a ? linkTargetImage((a.getAttribute("href") || "").slice(1)) : null;
  if (!hit || typeof imageViewer === "undefined") return false;
  hideLinkPeek();
  linkPeek.touchShown = false;
  imageViewer.open(hit.img, { keepDocPlace: true });
  return true;
}
function showLinkPeek(a, touch = false) {
  const hit = linkTargetImage((a.getAttribute("href") || "").slice(1));
  if (!hit) return false;
  if (linkPeek.a === a && linkPeek.el) return true;
  hideLinkPeek();
  const el = document.createElement("div");
  el.className = "link-peek";
  el.setAttribute("role", "dialog");
  el.setAttribute("aria-label", t("linkPeekZoom"));
  const img = document.createElement("img");
  img.src = hit.img.currentSrc || hit.img.src;
  img.alt = hit.img.getAttribute("alt") || "";
  const cap = document.createElement("div");
  cap.className = "link-peek-cap";
  cap.textContent = linkTargetImageLabel(hit);
  const tip = document.createElement("div");
  tip.className = "link-peek-tip";
  tip.textContent = t(touch ? "linkPeekGoTouch" : "linkPeekGo");
  // okienko jest klikalne: obraz albo „Powiększ” = podgląd obrazu (image-viewer.js)
  const zoom = document.createElement("button");
  zoom.type = "button";
  zoom.className = "btn link-peek-zoom";
  zoom.innerHTML = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16" y2="16"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/></svg><span></span>';
  zoom.querySelector("span").textContent = t("linkPeekZoom");
  if (!touch) { zoom.dataset.hint = ""; zoom.dataset.hintPl = I18N.pl.linkPeekZoomHint; zoom.dataset.hintEn = I18N.en.linkPeekZoomHint; zoom.dataset.hintDelay = "0.4"; }
  const row = document.createElement("div");
  row.className = "link-peek-row";
  row.append(tip, zoom);
  el.append(img, cap, row);
  el.addEventListener("click", (e) => { if (e.target === img || e.target.closest(".link-peek-zoom")) zoomLinkTarget(a); });
  el.addEventListener("pointerenter", () => clearTimeout(linkPeek.hideTimer));
  el.addEventListener("pointerleave", (e) => { if (e.pointerType !== "touch" && !a.contains(e.relatedTarget)) hideLinkPeekSoon(); });
  document.body.appendChild(el);
  linkPeek.el = el;
  linkPeek.a = a;
  linkPeek.img = hit.img;
  const place = () => {
    if (linkPeek.el !== el) return;
    const vv = window.visualViewport;
    const r = a.getBoundingClientRect();
    const viewW = vv ? vv.width : window.innerWidth;
    const viewH = vv ? vv.height : window.innerHeight;
    const w = el.offsetWidth; const h = el.offsetHeight;
    const left = Math.max(8, Math.min(r.left, viewW - w - 8));
    const top = r.bottom + 8 + h <= viewH - 8 ? r.bottom + 8 : Math.max(8, r.top - h - 8);
    el.style.left = `${left + (vv ? vv.offsetLeft : 0)}px`;
    el.style.top = `${top + (vv ? vv.offsetTop : 0)}px`;
  };
  place();
  if (!img.complete) img.addEventListener("load", place, { once: true });
  return true;
}

// ── skok + „Wróć” ────────────────────────────────────────────────────────────
// Wejście i wyjście „Wróć” tylko przez opacity/transform (CSS .link-back, .is-leaving); przy
// „Ogranicz ruch” — od razu.
function hideLinkBack() {
  clearTimeout(linkUi.backTimer);
  const b = linkUi.back;
  if (!b || b.hidden || b.classList.contains("is-leaving")) return;
  linkUi.backArmed = false;
  if (matchMedia("(prefers-reduced-motion: reduce)").matches) { b.hidden = true; return; }
  b.classList.add("is-leaving");
  const done = () => { if (b.classList.contains("is-leaving")) { b.classList.remove("is-leaving"); b.hidden = true; } };
  b.addEventListener("animationend", done, { once: true });
  setTimeout(done, 260); // zapas, gdyby animationend nie przyszło (karta w tle)
}
function showLinkBack(top) {
  if (!linkUi.back) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "link-back";
    b.hidden = true;
    b.addEventListener("click", () => linkGoBack());
    document.body.append(b);
    linkUi.back = b;
  }
  const b = linkUi.back;
  linkUi.backTop = top;
  linkUi.backArmed = false; // samo-chowanie czeka, aż widok odjedzie od miejsca sprzed skoku
  b.textContent = t("linkBack");
  if (b.hidden || b.classList.contains("is-leaving")) {
    b.classList.remove("is-leaving");
    b.hidden = false;
  }
  clearTimeout(linkUi.backTimer);
  linkUi.backTimer = setTimeout(hideLinkBack, 12000);
}
// Wróciłeś sam (przewijaniem) mniej więcej tam, skąd był skok — „Wróć” nie jest już potrzebne.
// Najpierw widok musi od tego miejsca odjechać (płynny skok startuje właśnie stamtąd).
function onDocScrollForLinkBack() {
  const b = linkUi.back;
  if (!b || b.hidden || b.classList.contains("is-leaving") || !docViewportEl) return;
  const near = Math.max(80, docViewportEl.clientHeight * 0.3);
  const dist = Math.abs(docViewportEl.scrollTop - linkUi.backTop);
  if (dist > near) linkUi.backArmed = true;
  else if (linkUi.backArmed) hideLinkBack();
}
// Płynny powrót przerywa każda natychmiastowa zmiana przewijania — np. przeliczenie granic stron
// (page-breaks.js trzyma akapit u góry na miejscu, gdy doszły czcionki/odstępy stron): widok
// zostawał przy celu skoku (test links, niestabilny w tłoku). Po animacji dociągamy bez animacji,
// chyba że użytkownik sam zaczął przewijać.
function linkGoBack() {
  if (!linkUi.back || linkUi.back.hidden || linkUi.back.classList.contains("is-leaving") || !docViewportEl) return false;
  const top = linkUi.backTop;
  docViewportEl.scrollTo({ top, behavior: "smooth" });
  hideLinkBack();
  const token = ++linkUi.backToken;
  const settle = () => {
    if (token !== linkUi.backToken || Math.abs(docViewportEl.scrollTop - top) <= 24) return;
    docViewportEl.scrollTo({ top, behavior: "auto" });
  };
  setTimeout(settle, 700);
  setTimeout(settle, 1500);
  return true;
}

function jumpToLinkTarget(id) {
  // „Początek dokumentu” (Word: w:anchor="_top" — bez zakładki w pliku)
  if (id === "_top") {
    const before = docViewportEl?.scrollTop || 0;
    docViewportEl?.scrollTo({ top: 0, behavior: "smooth" });
    showLinkBack(before);
    return;
  }
  const target = linkTargetEl(id);
  if (!target) { toast(t("linkTargetMissing"), "info"); return; }
  // link do obrazu: przewijamy do samego obrazu (drugi obraz akapitu, obraz za podziałem strony)
  const el = linkTargetImage(id)?.img || target.closest("p, td, li") || target;
  const before = docViewportEl?.scrollTop || 0;
  jumpToStructureItem({ el, id: "link" }, { silentSelect: true });
  setTimeout(() => el.classList.remove("search-hit", "search-hit-active"), 2200);
  showLinkBack(before);
  if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
}

function onDocLinkClick(e) {
  const a = e.target.closest?.("a[href]");
  if (!a || !docCanvasEl.contains(a) || e.button > 0) return;
  // przytrzymanie = sam podgląd; okienko zostaje po puszczeniu palca (można stuknąć „Powiększ”),
  // znika po stuknięciu obok albo przewinięciu
  if (linkPeek.touchShown) { linkPeek.touchShown = false; e.preventDefault(); return; }
  hideLinkPeek();
  // Shift+klik w link do obrazu = od razu podgląd obrazu (bez skoku)
  if (e.shiftKey && a.hasAttribute("data-dwb-img-link") && zoomLinkTarget(a)) { e.preventDefault(); return; }
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
  // podgląd obrazu przy linku: mysz — po chwili bezruchu na linku; dotyk — przytrzymanie palca
  const imgLinkAt = (e) => { const a = e.target.closest?.("a[data-dwb-img-link]"); return a && docCanvasEl.contains(a) ? a : null; };
  // też „move”: link, który podjechał pod stojący kursor przy przewijaniu kółkiem, nie dostaje
  // „over” — pierwszy ruch myszy nad nim ma pokazać podgląd
  const peekSoon = (e) => {
    if (e.pointerType === "touch") return;
    const a = imgLinkAt(e);
    if (a && linkPeek.a === a) clearTimeout(linkPeek.hideTimer); // wrócił z okienka na link
    if (!a || linkPeek.a === a || linkPeek.pending === a) return;
    linkPeek.pending = a;
    clearTimeout(linkPeek.timer);
    linkPeek.timer = setTimeout(() => { linkPeek.pending = null; showLinkPeek(a); }, 220);
  };
  docCanvasEl?.addEventListener("pointerover", peekSoon);
  docCanvasEl?.addEventListener("pointermove", peekSoon, { passive: true });
  docCanvasEl?.addEventListener("pointerout", (e) => {
    if (e.pointerType === "touch") return;
    const a = imgLinkAt(e);
    if (!a || a.contains(e.relatedTarget)) return;
    if (linkPeek.a === a && linkPeek.el) hideLinkPeekSoon(); // może jedzie na okienko
    else hideLinkPeek();
  });
  docCanvasEl?.addEventListener("pointerdown", (e) => {
    clearTimeout(linkPeek.touchTimer);
    if (e.pointerType !== "touch") { hideLinkPeek(); return; }
    const a = imgLinkAt(e);
    if (linkPeek.el && linkPeek.a !== a) hideLinkPeek();
    if (!a) return;
    linkPeek.touchShown = false;
    linkPeek.touchTimer = setTimeout(() => { if (showLinkPeek(a, true)) linkPeek.touchShown = true; }, 450);
  });
  ["pointerup", "pointercancel", "pointermove"].forEach((type) => docCanvasEl?.addEventListener(type, (e) => {
    if (e.pointerType !== "touch" || (type === "pointermove" && Math.hypot(e.movementX || 0, e.movementY || 0) < 6)) return;
    clearTimeout(linkPeek.touchTimer);
  }));
  // przytrzymanie na linku: bez systemowego menu (Kopiuj / Otwórz link) zamiast podglądu
  docCanvasEl?.addEventListener("contextmenu", (e) => { if (linkPeek.touchShown || (e.pointerType === "touch" && imgLinkAt(e))) e.preventDefault(); });
  docViewportEl?.addEventListener("scroll", hideLinkPeek, { passive: true });
  docViewportEl?.addEventListener("scroll", onDocScrollForLinkBack, { passive: true });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && linkPeek.el) hideLinkPeek();
    // Spacja przy okienku obrazu (jak „Szybki podgląd” na Macu) = powiększ; nie przy pisaniu
    if (e.key === " " && linkPeek.el && linkPeek.a && !e.ctrlKey && !e.metaKey && !e.altKey
      && !document.activeElement?.closest?.("input, textarea, select, [contenteditable]:not([contenteditable=false]), button")) {
      if (zoomLinkTarget(linkPeek.a)) e.preventDefault();
    }
  });
  document.addEventListener("pointerdown", (e) => { if (linkPeek.el && !docCanvasEl?.contains(e.target) && !linkPeek.el.contains(e.target)) hideLinkPeek(); }, true);
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
      docBookmarksFresh();
      return r;
    };
  }
});
