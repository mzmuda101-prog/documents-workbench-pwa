// Post-render fixes for docx-preview output (Symbol bullets, visual tofu, etc.).

// Cały podgląd (wszystkie strony/sekcje). Dawniej poprawki list brały tylko PIERWSZĄ sekcję —
// listy od drugiej strony (podział strony / sekcji, PDF → DOCX) zostawały nienaprawione.
function getPreviewRoot(host) {
  const root = host?.querySelector?.(".docx-preview-host") || host;
  return root?.querySelector(".docx-wrapper") || root;
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
  const section = getPreviewRoot(host);
  if (!section) return 0;
  let fixed = 0;
  section.querySelectorAll('p[class*="docx-num-"]').forEach((p) => {
    const paraStyle = getComputedStyle(p);
    if (paraStyle.display !== "list-item") return;
    const before = getComputedStyle(p, "::before");
    if (isImageBullet(before)) return;

    // Jak w Wordzie: numer / punktor na początku wcięcia wiszącego, tekst (po tabulatorze) ZAWSZE
    // na lewym wcięciu. docx-preview pisał „1.\t” w tekście, a tabulator zwijał się do spacji —
    // tekst punktu zaczynał się kilka-kilkanaście px za wcześnie (listy z Worda i z PDF → DOCX).
    // Znacznik wychodzi z biegu tekstu (absolutnie) — kursor przy edycji stoi na początku tekstu.
    const hang = -(parseFloat(paraStyle.textIndent) || 0);
    if (hang > 0.5 && before.content !== "none" && !usesLegacyBulletFont(before)) {
      p.classList.add("dwb-list-hang");
      p.style.setProperty("--dwb-hang", `${hang}px`);
      fixed++;
      return;
    }

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
  const section = getPreviewRoot(host);
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

// Obrazy zakotwiczone do STRONY (łatka w vendor-libs: position:absolute, „za tekstem” = z-index −1)
// przenosimy wprost do strony: w komórce tabeli czy akapicie z position:relative liczyłyby się
// od tamtego elementu, nie od rogu kartki.
function fixPageAnchoredDrawings(host) {
  const root = host?.querySelector?.(".docx-wrapper") || host;
  if (!root) return 0;
  let moved = 0;
  root.querySelectorAll("section.docx").forEach((section) => {
    section.querySelectorAll("div").forEach((d) => {
      if (d.style.position !== "absolute" || d.parentElement === section) return;
      if (!d.querySelector("img")) return;
      d.classList.add("dwb-page-anchor");
      section.insertBefore(d, section.firstChild);
      moved++;
    });
  });
  return moved;
}

// Tabulatory jak w Wordzie: docx-preview rysuje każdy tabulator jako stałą spację, a łatka w
// vendor-libs zostawia na <span class="docx-tab"> pozycje tabulatorów akapitu. Tu liczymy
// szerokość każdego: lewy (do pozycji), prawy/środkowy (tekst za nim kończy się/centruje na
// pozycji), domyślna siatka co data-dt pt, wcięcie wiszące = ukryty tabulator (punktory).
// Wypełnienie kropkami/kreską (spis treści, „Miejscowość: ……”). Współrzędne bez skali widoku
// (powiększenie robi transform). Przebiegi: najpierw wszystkie 1. tabulatory akapitów, potem 2.…
// — jedno przeliczenie układu na przebieg, nie na każdy tabulator.
function layoutTabStops(host) {
  const root = host?.querySelector?.(".docx-wrapper") || host;
  if (!root) return 0;
  const spans = [...root.querySelectorAll("span.docx-tab")];
  if (!spans.length) return 0;
  const byPara = new Map();
  for (const sp of spans) {
    const p = sp.closest("p");
    if (!p) continue;
    if (!byPara.has(p)) byPara.set(p, []);
    byPara.get(p).push(sp);
  }
  const PX2PT = 0.75;
  // ponowne liczenie (np. po dojściu czcionki): najpierw zerujemy szerokości — tabulator z
  // poprzedniego liczenia mógł zepchnąć resztę wiersza do nowej linijki i stamtąd byśmy mierzyli
  for (const sp of spans) if (sp.classList.contains("docx-tab-laid")) sp.style.width = "0pt";
  let done = 0;
  for (let pass = 0; ; pass++) {
    const jobs = [];
    for (const [p, list] of byPara) if (list[pass]) jobs.push([p, list[pass], list[pass + 1] || null]);
    if (!jobs.length || pass > 40) break;
    // odczyt
    const plans = jobs.map(([p, sp, next]) => {
      const pr = p.getBoundingClientRect();
      const scale = p.offsetWidth ? pr.width / p.offsetWidth : 1;
      if (!scale) return null;
      const cs = getComputedStyle(p);
      const ml = parseFloat(cs.marginLeft) || 0, pl = parseFloat(cs.paddingLeft) || 0;
      const originPx = pr.left - ml * scale;
      const r = sp.getBoundingClientRect();
      const x = ((r.left - originPx) / scale) * PX2PT;
      const indent = (ml + pl) * PX2PT;
      const hanging = (parseFloat(cs.textIndent) || 0) < -0.5;
      let stops = [];
      try {
        stops = JSON.parse(sp.dataset.stops || "[]").map(([pos, leader, style]) => ({ pos, leader, style }));
      } catch (_) { stops = []; }
      stops = stops.filter((t) => t.style !== "clear").sort((a, b) => a.pos - b.pos);
      if (hanging && indent > x + 0.5 && !stops.some((t) => t.pos > x + 0.5 && t.pos < indent)) stops.push({ pos: indent, leader: "none", style: "left" });
      stops.sort((a, b) => a.pos - b.pos);
      let stop = stops.find((t) => t.pos > x + 0.5);
      if (!stop) {
        const dt = parseFloat(sp.dataset.dt) || 36;
        const last = stops.length ? stops[stops.length - 1].pos : 0;
        let pos = Math.max(last, 0);
        while (pos <= x + 0.5) pos += dt;
        stop = { pos, leader: "none", style: "left" };
      }
      let after = 0;
      if (stop.style === "right" || stop.style === "center" || stop.style === "decimal") {
        const range = document.createRange();
        range.setStartAfter(sp);
        if (next) range.setEndBefore(next);
        else range.setEndAfter(p.lastChild || p);
        after = (range.getBoundingClientRect().width / scale) * PX2PT;
        if (stop.style === "center") after /= 2;
      }
      return { sp, width: Math.max(0, stop.pos - x - after), leader: stop.leader };
    });
    // zapis
    for (const pl of plans) {
      if (!pl) continue;
      const st = pl.sp.style;
      st.display = "inline-block";
      st.width = `${pl.width.toFixed(2)}pt`;
      st.whiteSpace = "pre";
      st.overflow = "hidden"; // wyrównanie w pionie: CSS (.docx-tab-laid, p.dwb-exact)
      pl.sp.innerHTML = "&nbsp;";
      pl.sp.classList.add("docx-tab-laid");
      pl.sp.classList.toggle("tab-dot", pl.leader === "dot" || pl.leader === "middleDot");
      pl.sp.classList.toggle("tab-line", pl.leader === "underscore" || pl.leader === "heavy" || pl.leader === "hyphen");
      done++;
    }
  }
  return done;
}

// docx-preview daje domyślny krój i rozmiar dokumentu (docDefaults, style akapitów) tylko
// fragmentom tekstu: „.docx span { font-family: Calibri; font-size: 11pt }”,
// „p.docx_heading1 span { … }”. Sam akapit dziedziczył krój APLIKACJI — pusty akapit, nowy po
// Enterze i tekst wpisany w nim bez fragmentu wyglądały jak „Space Grotesk 12 pt” (pasek też tak
// pokazywał), choć w pliku i w Wordzie to Calibri 11 pt; puste akapity miały też inną wysokość
// niż w Wordzie (granice stron). Te same cechy kopiujemy na akapit — tuż za oryginałem, żeby
// kolejność (styl domyślny → style akapitów) się nie zmieniła.
const RUN_DEFAULT_PROPS = ["font-family", "font-size", "font-weight", "font-style", "color"];
function applyRunDefaultsToParagraphs(host) {
  if (!host) return 0;
  let n = 0;
  host.querySelectorAll("style").forEach((el) => {
    let sheet;
    try { sheet = el.sheet; } catch (_) { sheet = null; }
    if (!sheet?.cssRules) return;
    for (let i = sheet.cssRules.length - 1; i >= 0; i--) {
      const rule = sheet.cssRules[i];
      if (rule.type !== 1 || !rule.selectorText) continue;
      const sels = rule.selectorText.split(",").map((s) => s.trim());
      // tylko „X span” (nie „span.docx_hyperlink” — styl znakowy) — X = .docx albo p.<styl akapitu>
      if (!sels.every((s) => / span$/.test(s))) continue;
      const targets = sels.map((s) => {
        const base = s.slice(0, -" span".length).trim();
        return /(^|\s)p\.[\w-]+$/.test(base) ? base : /^\.[\w-]+$/.test(base) ? `${base} p` : null;
      });
      if (targets.some((s) => !s)) continue;
      const decl = RUN_DEFAULT_PROPS.map((prop) => {
        const v = rule.style.getPropertyValue(prop);
        return v ? `${prop}: ${v};` : "";
      }).join(" ").trim();
      if (!decl) continue;
      try { sheet.insertRule(`${targets.join(", ")} { ${decl} }`, i + 1); n++; } catch (_) { /* nieznany selektor */ }
    }
  });
  return n;
}

// „Nie dodawaj odstępu między akapitami tego samego stylu” (w:contextualSpacing; łatka 13 w
// vendor-libs daje takiemu akapitowi data-dwb-ctx). Jak w Wordzie: akapit z tą cechą traci własny
// odstęp przed, gdy poprzedni akapit ma ten sam styl, i odstęp po, gdy następny ma ten sam styl.
// Styl = klasa docx_<id> (akapit bez stylu: brak klasy = styl domyślny). Sąsiad tylko w tym samym
// pojemniku (komórka tabeli, strona) — jak w Wordzie.
const PAGE_LAYOUT_OWN = ".dwb-page-gap, .dwb-page-gap-spacer, .dwb-page-break, .dwb-page-gap-band, .dwb-page-margin-guide";
function docxParagraphStyleKey(el) {
  return [...el.classList].find((c) => c.startsWith("docx_")) || "";
}

function applyContextualSpacing(host) {
  if (!host) return 0;
  let n = 0;
  const set = (p, prop, zero, orig) => {
    const want = zero ? "0px" : orig;
    if (p.style[prop] !== want) { p.style[prop] = want; n++; }
  };
  host.querySelectorAll("section.docx p[data-dwb-ctx]").forEach((p) => {
    // własne marginesy akapitu z pliku (docx-preview: w:spacing wprost) — do przywrócenia, gdy
    // sąsiad przestanie mieć ten sam styl (Enter, koniec listy, sklejenie akapitów w Edycji)
    if (p.dataset.dwbCtxOrig == null) p.dataset.dwbCtxOrig = JSON.stringify([p.style.marginTop, p.style.marginBottom]);
    let orig = ["", ""];
    try { orig = JSON.parse(p.dataset.dwbCtxOrig); } catch (_) { /* uszkodzony atrybut — bez własnych marginesów */ }
    const key = docxParagraphStyleKey(p);
    // odstępy między stronami w Edycji (page-breaks.js) to nie sąsiedzi — granica strony w środku
    // listy nie może przywracać odstępów (przesunięcie układu → kolejne przeliczenie stron)
    let prev = p.previousElementSibling;
    while (prev?.matches(PAGE_LAYOUT_OWN)) prev = prev.previousElementSibling;
    let next = p.nextElementSibling;
    while (next?.matches(PAGE_LAYOUT_OWN)) next = next.nextElementSibling;
    // sąsiadem akapitu przy tabeli jest akapit w tabeli (Word: nagłówek z „bez odstępu…” nad tabelą
    // w tym samym stylu nie ma odstępu — „Zapotrzebowanie”: tabele o 8 pt niżej)
    if (next?.tagName === "TABLE") next = next.querySelector("td p");
    if (prev?.tagName === "TABLE") { const ps = prev.querySelectorAll("td p"); prev = ps[ps.length - 1] || null; }
    set(p, "marginTop", prev?.tagName === "P" && docxParagraphStyleKey(prev) === key, orig[0]);
    set(p, "marginBottom", next?.tagName === "P" && docxParagraphStyleKey(next) === key, orig[1]);
  });
  return n;
}

// Pusty akapit dodany w Edycji (Enter, wklejka) nie przeszedł przez applyWordLineMetrics — bez
// wysokości linijki po przerysowaniu podskakiwałby o kilka px. Wysokość = jego interlinia.
function fixEmptyParagraphHeights(host) {
  host?.querySelectorAll("section.docx p").forEach((p) => {
    // „0px” to wyzerowane minimum akapitu z treścią (applyWordLineMetrics) — klon po Enterze je dziedziczy
    if ((p.style.minHeight && p.style.minHeight !== "0px") || p.textContent.length || p.querySelector(EMPTY_P_CONTENT)) return;
    const cs = getComputedStyle(p);
    const fs = parseFloat(cs.fontSize);
    const lh = p.style.lineHeight;
    if (!fs) return;
    if (/^[\d.]+$/.test(lh)) p.style.minHeight = `${Math.round(fs * parseFloat(lh) * 100) / 100}px`;
    else if (/(pt|px)$/.test(lh)) p.style.minHeight = lh;
    else p.style.minHeight = `${Math.round(fs * wordLineFactor(cs.fontFamily) * 100) / 100}px`;
  });
}

// Po zmianach akapitów w podglądzie (Edycja) — odstępy list i puste akapity jak po narysowaniu.
function watchParagraphLayout(wrapper) {
  if (!wrapper || typeof MutationObserver !== "function") return;
  // w następnej klatce, nie w trakcie operacji edycji (natychmiast psuło kotwiczenie widoku przy
  // Enter/Backspace w pustym punkcie listy — test page-gaps-edit)
  let pending = false;
  const mo = new MutationObserver(() => {
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => {
      pending = false;
      if (!wrapper.isConnected) { mo.disconnect(); return; }
      applyContextualSpacing(wrapper);
      fixEmptyParagraphHeights(wrapper);
    });
  });
  mo.observe(wrapper, { childList: true, subtree: true });
}

// Linie poziome tabeli dokładane do wysokości wiersza — jak w Wordzie. W Wordzie wiersz
// „co najmniej 17 pt” z linią 0,5 pt ma 17,5 pt (PDF z Worda, „S-99_P”: 23,3 px na wiersz),
// a przeglądarka mieści linię w 17 pt (22,7 px) — przy 30 wierszach strona kończyła się 18 px
// wcześniej; wiersz z treścią (bez minimum) też był o linię niższy. Grubość z DEKLARACJI
// (styl komórki albo reguła tabeli): obliczona jest zaokrąglona do 1 px. Tabele bez linii — bez zmian.
const PT_PX = 96 / 72;

// Indeks górny/dolny: rozmiar jak w Wordzie = 2/3 tekstu zaokrąglone w dół do pół punktu
// (pomiar w Wordzie, 13 rozmiarów — app.css). Odnośniki przypisów też (styl z vertAlign).
function applyWordSuperscripts(host) {
  if (!host) return 0;
  let n = 0;
  host.querySelectorAll("section.docx sup, section.docx sub").forEach((el) => {
    const base = parseFloat(getComputedStyle(el.parentElement).fontSize) * 0.75; // pt
    if (!base) return;
    el.style.fontSize = `${Math.floor(base * 4 / 3) / 2}pt`;
    n++;
  });
  return n;
}
function cssLengthPx(v) {
  const m = /^(-?[\d.]+)(pt|px)$/.exec(String(v || "").trim());
  return m ? parseFloat(m[1]) * (m[2] === "pt" ? PT_PX : 1) : null;
}
function addTableBorderHeights(host) {
  if (!host) return 0;
  const rules = [];
  host.querySelectorAll("style").forEach((el) => {
    let list;
    try { list = el.sheet?.cssRules; } catch (_) { list = null; }
    for (const r of list || []) if (r.type === 1 && /\btd\b/.test(r.selectorText) && (r.style.borderTopWidth || r.style.borderBottomWidth)) rules.push(r);
  });
  // grubość linii z deklaracji: styl komórki, inaczej ostatnia pasująca reguła
  const declared = (td, side) => {
    const prop = side === "top" ? "borderTopWidth" : "borderBottomWidth";
    const styleProp = side === "top" ? "borderTopStyle" : "borderBottomStyle";
    const cs = getComputedStyle(td);
    if (cs[styleProp] === "none" || cs[styleProp] === "hidden" || !parseFloat(cs[prop])) return 0;
    let v = td.style[prop];
    if (!v) for (let i = rules.length - 1; i >= 0; i--) if (rules[i].style[prop] && td.matches(rules[i].selectorText)) { v = rules[i].style[prop]; break; }
    return cssLengthPx(v) ?? parseFloat(cs[prop]);
  };
  let n = 0;
  host.querySelectorAll("section.docx table").forEach((table) => {
    if (getComputedStyle(table).borderCollapse !== "collapse") return;
    const rows = [...table.rows];
    rows.forEach((tr, i) => {
      const cells = [...tr.cells];
      if (!cells.length) return;
      let extra = Math.max(0, ...cells.map((td) => declared(td, "bottom")), ...(rows[i + 1] ? [...rows[i + 1].cells].map((td) => declared(td, "top")) : []));
      if (i === 0) extra += Math.max(0, ...cells.map((td) => declared(td, "top")));
      if (!(extra > 0.05)) return;
      // treść: dopełnienie komórek po połowie u góry i u dołu; minimum wiersza: + linia
      cells.forEach((td) => {
        const cs = getComputedStyle(td);
        td.style.paddingTop = `${(parseFloat(cs.paddingTop) || 0) + extra / 2}px`;
        td.style.paddingBottom = `${(parseFloat(cs.paddingBottom) || 0) + extra / 2}px`;
      });
      const h = cssLengthPx(tr.style.height);
      if (h) tr.style.height = `${Math.round((h + extra) * 100) / 100}px`;
      n++;
    });
  });
  return n;
}

// Komórka scalona w pionie (w:vMerge → rowspan): w Wordzie PIERWSZY wiersz scalenia ma co
// najmniej wysokość pierwszej linijki jej tekstu (bez odstępu przed), dalsze wiersze swoje minima,
// a całość rośnie, gdy treść jest wyższa. Przeglądarka rozkłada komórkę tylko na minima wierszy.
// Zmierzone w Wordzie na plikach testowych (2026-10-09): linijka 9,4 pt, minima 3,5 + 9,95 pt →
// Word 19,4 pt, przeglądarka 13,45 („ubezp” z konwersji PDF: tabele o 30 % za niskie, 4 strony
// zamiast 7). Wysokość linijki = interlinia pierwszego akapitu komórki.
function applyMergedCellFirstLine(host) {
  if (!host) return 0;
  let n = 0;
  host.querySelectorAll("section.docx td[rowspan]").forEach((td) => {
    if (td.rowSpan < 2) return;
    const p = td.querySelector("p");
    const tr = td.parentElement;
    if (!p || !p.textContent.trim() || tr?.tagName !== "TR") return;
    const cs = getComputedStyle(p);
    const lh = parseFloat(cs.lineHeight);
    if (!lh) return;
    const tcs = getComputedStyle(td);
    const need = lh + (parseFloat(tcs.paddingTop) || 0) + (parseFloat(tcs.paddingBottom) || 0);
    const cur = cssLengthPx(tr.style.height) || 0;
    if (need > cur + 0.1) { tr.style.height = `${Math.round(need * 100) / 100}px`; n++; }
  });
  return n;
}

// Obrazy zakotwiczone (łatka 15: data-dwb-anchor) — miejsce jak w Wordzie: poziomo względem
// strony / marginesu / kolumny, pionowo względem strony / marginesu / akapitu, przesunięciem albo
// wyrównaniem. Obraz zostaje w swoim akapicie (edycja i zapis go widzą) i stoi bezwzględnie
// WZGLĘDEM AKAPITU (position: relative na <p> — ten sam punkt odniesienia w podglądzie i w podglądzie
// wydruku, gdzie akapity są przenoszone na kartki). Oblewanie tekstem (wrapSquare/Tight/Through):
// niewidoczne „rezerwacje” (span.dwb-anchor-wrap, float) na początku akapitu, ułożone od skrajnego
// — kilka obrazów po jednej stronie też stoi na swoich miejscach; ujemny margin-top = obraz
// wystający nad akapit zabiera miejsce tylko tam, gdzie zachodzi na tekst. Liczenie stron pomija
// obraz (.dwb-anchor-abs) i rezerwacje. Strona akapitu w długiej sekcji: z jego położenia.
function positionAnchoredDrawings(host) {
  if (!host) return 0;
  const byPara = new Map();
  host.querySelectorAll("section.docx [data-dwb-anchor]").forEach((el) => {
    let d;
    try { d = JSON.parse(el.dataset.dwbAnchor); } catch (_) { return; }
    const wrapText = /^wrap(Square|Tight|Through)$/.test(d.w);
    if (!wrapText && !(d.w === "wrapNone" && !(d.hr === "page" && d.vr === "page"))) return; // reszta jak dotąd
    if (el.closest("td, header, footer")) return; // w komórce / nagłówku: układ docx-preview
    const p = el.closest("p");
    if (!p || !el.closest("section.docx")) return;
    if (!byPara.has(p)) byPara.set(p, []);
    byPara.get(p).push({ el, d, wrapText });
  });
  let n = 0;
  byPara.forEach((items, p) => {
    const sec = p.closest("section.docx");
    const scale = sec.getBoundingClientRect().height / sec.offsetHeight || 1;
    const cs = getComputedStyle(sec);
    const padL = parseFloat(cs.paddingLeft) || 0, padR = parseFloat(cs.paddingRight) || 0;
    const padT = parseFloat(cs.paddingTop) || 0, padB = parseFloat(cs.paddingBottom) || 0;
    const pageW = sec.offsetWidth;
    const pageH = parseFloat(cs.minHeight) || pageW * Math.SQRT2;
    const sr = sec.getBoundingClientRect();
    const pr = p.getBoundingClientRect();
    const pTop = (pr.top - sr.top) / scale, pLeft = (pr.left - sr.left) / scale;
    const pW = pr.width / scale, pH = pr.height / scale;
    const bodyH = Math.max(1, pageH - padT - padB);
    const pageTop = Math.max(0, Math.floor((pTop - padT) / bodyH)) * pageH; // strona akapitu
    const off = (v) => cssLengthPx(v) ?? 0;
    const placed = items.map(({ el, d, wrapText }) => {
      // wrapNone: docx-preview daje kotwicy 0×0 (obraz wystaje) — rozmiar z samego obrazu
      const img = el.querySelector("img, svg");
      const w = cssLengthPx(el.style.width) || cssLengthPx(img?.style.width) || (img ? img.getBoundingClientRect().width / scale : 0);
      const h = cssLengthPx(el.style.height) || cssLengthPx(img?.style.height) || (img ? img.getBoundingClientRect().height / scale : 0);
      const hF = d.hr === "page" ? [0, pageW] : d.hr === "leftMargin" ? [0, padL] : d.hr === "rightMargin" ? [pageW - padR, pageW] : [padL, pageW - padR];
      let x = hF[0] + off(d.ho);
      if (d.ha === "right" || d.ha === "outside") x = hF[1] - w;
      else if (d.ha === "center") x = (hF[0] + hF[1] - w) / 2;
      else if (d.ha === "left" || d.ha === "inside") x = hF[0];
      const vF = d.vr === "page" ? [pageTop, pageTop + pageH]
        : d.vr === "margin" ? [pageTop + padT, pageTop + pageH - padB]
        : d.vr === "topMargin" ? [pageTop, pageTop + padT]
        : d.vr === "bottomMargin" ? [pageTop + pageH - padB, pageTop + pageH]
        : [pTop, pTop + pH]; // paragraph / line
      let y = vF[0] + off(d.vo);
      if (d.va === "bottom" || d.va === "outside") y = vF[1] - h;
      else if (d.va === "center") y = (vF[0] + vF[1] - h) / 2;
      else if (d.va === "top" || d.va === "inside") y = vF[0];
      return { el, d, wrapText, x, y, w, h };
    });
    if (getComputedStyle(p).position === "static") p.style.position = "relative";
    placed.forEach(({ el, d, wrapText, x, y, w, h }) => {
      el.style.position = "absolute";
      el.style.float = "none";
      el.style.margin = "0";
      el.style.left = `${Math.round((x - pLeft) * 100) / 100}px`;
      el.style.top = `${Math.round((y - pTop) * 100) / 100}px`;
      el.style.zIndex = d.b ? "-1" : "1";
      el.classList.add("dwb-anchor-abs");
      if (wrapText) { el.style.width = `${w}px`; el.style.height = `${h}px`; }
      n++;
    });
    // rezerwacje miejsca dla oblewanych: prawa strona od prawej krawędzi, lewa od lewej
    const wraps = placed.filter((x) => x.wrapText && x.y + x.h > pTop);
    const side = (x) => x.x + x.w / 2 > pLeft + pW / 2;
    const make = (o, right, gap) => {
      const ph = document.createElement("span");
      ph.className = "dwb-anchor-wrap";
      ph.setAttribute("aria-hidden", "true");
      ph.style.cssText = `float:${right ? "right" : "left"};display:block;width:${o.w}px;height:${o.h}px;margin-top:${Math.round((o.y - pTop) * 100) / 100}px;`
        + (right ? `margin-right:${Math.max(0, gap)}px;margin-left:${cssLengthPx(o.d.dl) ?? 12}px;` : `margin-left:${Math.max(0, gap)}px;margin-right:${cssLengthPx(o.d.dr) ?? 12}px;`);
      return ph;
    };
    const out = [];
    let edge = pLeft + pW;
    wraps.filter(side).sort((a, b) => (b.x + b.w) - (a.x + a.w)).forEach((o) => { out.push(make(o, true, edge - (o.x + o.w))); edge = o.x - (cssLengthPx(o.d.dl) ?? 12); });
    edge = pLeft;
    wraps.filter((o) => !side(o)).sort((a, b) => a.x - b.x).forEach((o) => { out.push(make(o, false, o.x - edge)); edge = o.x + o.w + (cssLengthPx(o.d.dr) ?? 12); });
    if (out.length) p.prepend(...out);
  });
  return n;
}

// Akapit kończący się ręcznym końcem wiersza (w:br, Shift+Enter): w Wordzie znak akapitu stoi w
// NASTĘPNEJ linijce, więc „tekst⏎(Shift+Enter)” to dwie linijki, a dwa same w:br — trzy. Przeglądarka
// nie robi linijki za <br> na końcu bloku („Podsumowanie dla porządkowych”: akapit z dwoma
// łamaniami pod tytułem o linijkę niższy niż w Wordzie, cała strona przesunięta). Dokładamy
// <br data-dwb-ph> — ten sam znacznik „pusty wiersz tylko dla oka”, którego używa Edycja (zapis go
// pomija, docx-run-styles.js).
function addTrailingBreakLines(host) {
  if (!host) return 0;
  let n = 0;
  host.querySelectorAll("section.docx p").forEach((p) => {
    let last = p.lastChild;
    while (last) {
      if (last.nodeType === 3 && !last.data.length) { last = last.previousSibling; continue; }
      if (last.nodeType === 1 && last.tagName !== "BR" && last.lastChild && !last.matches?.("[contenteditable=false], img, svg")) { last = last.lastChild; continue; }
      break;
    }
    if (last?.nodeType !== 1 || last.tagName !== "BR" || last.dataset.dwbPh || last.closest(".dwb-page-break")) return;
    const ph = document.createElement("br");
    ph.dataset.dwbPh = "1";
    last.after(ph);
    n++;
  });
  return n;
}

// Coś, co daje akapitowi linijkę mimo braku tekstu (obraz, wyspa, przerwa wiersza, pole).
const EMPTY_P_CONTENT = "img, svg, canvas, video, object, iframe, br, input, select, textarea, [contenteditable=false]";

// Krój „szkieletu” linijki = krój tekstu. Przeglądarka buduje linijkę także z niewidocznego
// szkieletu w kroju AKAPITU (styl domyślny, np. Calibri), gdy tekst jest w innym (Arial): różne
// proporcje góra/dół podnosiły linijkę o ~1 px — i tylko wtedy, gdy krój akapitu był już wczytany
// („CO-23”: raz 98 %, raz 58 % zgodności podglądu wydruku). Word liczy linijkę z samego tekstu.
// Punktor/numer listy bez własnego kroju zostaje przy dotychczasowym (--dwb-p-font, app.css).
// Fragment wzorcowy linijki: największy tekst, przy równym rozmiarze — najdłuższy; kroje symboli
// (punktor „” w Symbol/Wingdings jako pierwszy fragment akapitu z PDF) tylko, gdy nie ma innych.
const SYMBOL_FONT = /^["']?(symbol|wingdings( \d)?|webdings|zapf ?dingbats|marlett)["']?$/i;
function dominantRun(p) {
  let best = null, bestFs = 0, bestLen = 0, bestSym = true;
  p.querySelectorAll("span").forEach((sp) => {
    // tekst fragmentu: jego własne węzły tekstu i indeks górny/dolny w nim (podpis cały w <sup>)
    const own = [...sp.childNodes].map((n) => (n.nodeType === 3 ? n.data : n.nodeType === 1 && /^(SUP|SUB)$/.test(n.tagName) ? n.textContent : "")).join("").trim();
    if (!own || sp.closest("sup, sub")) return;
    const cs = getComputedStyle(sp);
    const fs = parseFloat(cs.fontSize);
    const sym = SYMBOL_FONT.test(cs.fontFamily.split(",")[0].trim());
    const better = (bestSym && !sym) || (sym === bestSym && (fs > bestFs + 0.01 || (Math.abs(fs - bestFs) <= 0.01 && own.length > bestLen)));
    if (!best || better) { best = sp; bestFs = fs; bestLen = own.length; bestSym = sym; }
  });
  return best;
}

function matchStrutFont(p, runFamily) {
  const own = getComputedStyle(p).fontFamily;
  if (!runFamily || own === runFamily) return;
  if (!p.style.getPropertyValue("--dwb-p-font")) p.style.setProperty("--dwb-p-font", own);
  p.style.fontFamily = runFamily;
}

// Punktor/numer listy w Wordzie ma cechy ZNAKU AKAPITU (gdy definicja listy ich nie podaje), nie
// tekstu: „WYKŁAD” — tekst 5 pt, znak akapitu (styl) 11 pt → punktor 11 pt i to on wyznacza
// wysokość linijki. Akapit dostaje rozmiar tekstu (applyWordLineMetrics), więc punktor zachowuje
// swój przez --dwb-p-size / --dwb-p-font (app.css, najniższa ważność — rozmiar z listy wygrywa).
function keepMarkerFont(p) {
  if (!/(^|\s)docx-num-/.test(p.className) || p.style.getPropertyValue("--dwb-p-size")) return;
  p.style.setProperty("--dwb-p-size", p.dataset.dwbMarkSize || getComputedStyle(p).fontSize);
  if (p.dataset.dwbMarkFont) p.style.setProperty("--dwb-p-font", withGenericFontFallback(`"${p.dataset.dwbMarkFont}"`) || `"${p.dataset.dwbMarkFont}"`);
}

// Kolejność stylów jak w Wordzie (OOXML): ustawienia domyślne → STYL TABELI → styl akapitu → styl
// znakowy → wprost. docx-preview pisze styl tabeli jako „table.docx_X span” — ważniejszy w CSS niż
// styl akapitu Normalny („.docx :where(p) span”, łatka 10), więc w tabeli wygrywał krój tabeli
// („Zapotrzebowanie”: Table Grid = Calibri, Normalny = Arial — Word pokazuje Arial, my Calibri).
// Reguły tekstu stylów tabel i ustawień domyślnych dostają najniższą ważność (:where) — kolejność
// w arkuszu (domyślne → style) zostaje, więc styl tabeli dalej wygrywa z domyślnymi.
// Wołać PO applyRunDefaultsToParagraphs (szuka „.docx span”).
function lowerTableStylePrecedence(host) {
  if (!host) return 0;
  let n = 0;
  host.querySelectorAll("style").forEach((el) => {
    let rules;
    try { rules = el.sheet?.cssRules; } catch (_) { rules = null; }
    for (const rule of rules || []) {
      if (rule.type !== 1 || !rule.selectorText) continue;
      const parts = rule.selectorText.split(",").map((x) => x.trim());
      let changed = false;
      const out = parts.map((sel) => {
        let m = /^(\.[\w-]+) span$/.exec(sel); // ustawienia domyślne: „.docx span”
        if (m) { changed = true; return `:where(${m[1]}) span`; }
        m = /^(table\.[\w-]+(?:[ >][^,]*)?) span$/.exec(sel); // styl tabeli (też warunkowy: pierwszy wiersz…)
        if (m && !/:where\(/.test(m[1])) { changed = true; return `:where(${m[1]}) span`; }
        return sel;
      });
      if (changed) { try { rule.selectorText = out.join(", "); n++; } catch (_) { /* nieobsługiwany selektor */ } }
    }
  });
  return n;
}

// Punktor większy niż tekst (11 pt przy tekście 5 pt) w Wordzie podnosi pierwszą linijkę akapitu
// do swojej wysokości. U nas punktor stoi poza tekstem (absolute, .dwb-list-hang) — dokładamy
// różnicę jako dopełnienie (≈ 80 % nad tekstem, 20 % pod, jak wysokość nad/pod linią bazową),
// a punktor podnosimy o górne dopełnienie (--dwb-mk-top), żeby trzymał linię bazową tekstu.
// Interlinia „dokładnie” — bez zmian (Word też trzyma stałą wysokość).
let markerCanvas = null;
function markerGlyphHeight(b, mfs) {
  try {
    markerCanvas = markerCanvas || document.createElement("canvas").getContext("2d");
    markerCanvas.font = `${b.fontStyle} ${b.fontWeight} ${mfs}px ${b.fontFamily}`;
    const m = markerCanvas.measureText("•");
    const h = (m.fontBoundingBoxAscent || 0) + (m.fontBoundingBoxDescent || 0);
    if (h > 0) return h;
  } catch (_) { /* brak canvas — współczynnik kroju */ }
  return mfs * wordLineFactor(b.fontFamily);
}

function addMarkerLineSpace(p) {
  if (!p.style.getPropertyValue("--dwb-p-size") || p.classList.contains("dwb-exact")) return;
  const b = getComputedStyle(p, "::before");
  if (!b.content || b.content === "none") return;
  const cs = getComputedStyle(p);
  const mfs = parseFloat(b.fontSize), pfs = parseFloat(cs.fontSize), plh = parseFloat(cs.lineHeight);
  if (!(mfs > pfs * 1.05) || !plh || !pfs) return;
  // wysokość punktora = zasięg liter JEGO kroju nad i pod linią bazową (bez odstępu kroju i bez
  // mnożnika interlinii) − linijka tekstu („WYKŁAD”: punktor Arial 11 pt przy tekście 5 pt i
  // interlinii 1,08 → Word 12,1 pt na punkt; Arial: 0,905 + 0,212 = 1,117 × 11 = 12,3)
  const extra = markerGlyphHeight(b, mfs) - plh;
  if (extra < 0.5) return;
  const top = Math.round(extra * 0.8 * 100) / 100;
  p.style.paddingTop = `${top}px`;
  p.style.paddingBottom = `${Math.round(extra * 0.2 * 100) / 100}px`;
  p.style.setProperty("--dwb-mk-top", `${-top}px`);
}

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
    // Pusty akapit (bez tekstu i obiektów) nie ma w przeglądarce żadnej linijki — jego wysokość
    // brała się z min-height = sam rozmiar czcionki stylu. W Wordzie to jedna linijka: rozmiar ×
    // interlinia × współczynnik kroju, gdzie rozmiar = fragmentu, jeśli akapit jakiś ma (np. z
    // kształtem, którego podgląd nie rysuje — linie do wypełnienia w „umowie najmu”), a bez
    // fragmentów = ZNAKU KOŃCA akapitu (w:pPr/w:rPr/w:sz, łatka 12).
    const empty = !p.textContent.length && !p.querySelector(EMPTY_P_CONTENT);
    const emptyRun = empty ? p.querySelector("span[style*='font-size']") : null;
    if (empty && !emptyRun && p.dataset.dwbMarkSize) p.style.fontSize = p.dataset.dwbMarkSize;
    if (empty && !emptyRun && p.dataset.dwbMarkFont) p.style.fontFamily = withGenericFontFallback(`"${p.dataset.dwbMarkFont}"`) || `"${p.dataset.dwbMarkFont}"`;
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
    if (!m && getComputedStyle(p).lineHeight === "normal") {
      // Bez interlinii w pliku (brak w:spacing w akapicie, stylu i docDefaults) = pojedyncza, jak w
      // Wordzie: rozmiar × współczynnik kroju. Przeglądarka brała „normal” z metryk swojego kroju
      // i z rozmiaru AKAPITU (stylu), nie tekstu — podpis 8 pt pod linią miał linijkę 12 pt, a
      // zwykły wiersz Times 13,5 pt zamiast 13,8 („A-RAVI-PROJEKT”: 15 pt dryfu na stronie).
      if (empty) {
        const cs = getComputedStyle(emptyRun || p);
        const fs = parseFloat(cs.fontSize);
        if (fs) p.style.minHeight = `${Math.round(fs * wordLineFactor(cs.fontFamily) * 100) / 100}px`;
        return;
      }
      // rozmiar linijki = największy tekst w akapicie (Word liczy linijkę z fragmentów)
      const big = dominantRun(p);
      const bigFs = big ? parseFloat(getComputedStyle(big).fontSize) : 0;
      if (big) keepMarkerFont(p);
      if (!big) {
        // bez tekstu, same łamania wiersza: linijki w kroju i rozmiarze znaku akapitu (Word)
        if (!p.querySelector("br")) return;
        if (p.dataset.dwbMarkSize) p.style.fontSize = p.dataset.dwbMarkSize;
        if (p.dataset.dwbMarkFont) p.style.fontFamily = withGenericFontFallback(`"${p.dataset.dwbMarkFont}"`) || `"${p.dataset.dwbMarkFont}"`;
        p.style.lineHeight = String(wordLineFactor(getComputedStyle(p).fontFamily));
        fixed++;
        return;
      }
      p.style.fontSize = `${bigFs}px`;
      p.style.lineHeight = String(wordLineFactor(getComputedStyle(big).fontFamily));
      matchStrutFont(p, getComputedStyle(big).fontFamily);
      addMarkerLineSpace(p);
      fixed++;
      return;
    }
    if (!m) {
      // Interlinia „dokładnie” (wartość w pt): wysokość linijki w przeglądarce bierze się też
      // z czcionki SAMEGO akapitu (domyślnej z dokumentu), nie tylko tekstu. Gdy tekst jest
      // mniejszy, każda linijka rosła o 1–2 pt (tabela z PDF „puchła” o 90 pt na stronie).
      // Word liczy tylko podaną wysokość — ustawiamy akapitowi rozmiar jego tekstu.
      if (/(pt|px)$/.test(p.style.lineHeight)) {
        p.classList.add("dwb-exact"); // fragmenty tekstu nie podnoszą linijki (CSS niżej)
        // docx-preview daje akapitowi min-height z domyślnej czcionki dokumentu — mały tekst
        // (podpis 6 pt przy domyślnych 10 pt) zajmował więcej miejsca niż w Wordzie. Wysokość
        // = podana wysokość linijki (pusty akapit-odstęp też ją ma — bez min-height miałby 0).
        if (p.style.minHeight || empty) p.style.minHeight = p.style.lineHeight;
        const run = [...p.querySelectorAll("span")].find((sp) => sp.textContent.trim());
        if (run) {
          const fs = parseFloat(getComputedStyle(run).fontSize);
          if (fs && Math.abs(fs - parseFloat(getComputedStyle(p).fontSize)) > 0.1) {
            p.style.fontSize = `${fs}px`;
            fixed++;
          }
        }
      }
      return;
    }
    const run = empty ? emptyRun || p : dominantRun(p) || p.querySelector("span") || p;
    const cs = getComputedStyle(run);
    const fs = parseFloat(cs.fontSize);
    if (!fs) return;
    const ratio = Math.round(m * wordLineFactor(cs.fontFamily) * 1000) / 1000;
    if (!empty) keepMarkerFont(p);
    p.style.fontSize = `${fs}px`;
    p.style.lineHeight = String(ratio);
    if (!empty) { matchStrutFont(p, cs.fontFamily); addMarkerLineSpace(p); }
    if (empty) p.style.minHeight = `${Math.round(fs * ratio * 100) / 100}px`;
    fixed++;
  });
  return fixed;
}

// Brakujący krój → właściwy RODZAJ zastępczy. docx-preview pisze „font-family: Calibri” bez
// niczego dalej, więc gdy kroju (albo jego pogrubionej odmiany) nie ma na urządzeniu, przeglądarka
// brała czcionkę domyślną — na Androidzie szeryfową — a odmiana pogrubiona zastępnika Calibri
// potrafiła trafić w ZWYKŁY plik Roboto (pogrubienie znikało; zgłoszenie Mateusza z Androida,
// 2026-10-03). Dopisujemy „sans-serif” / „serif” / „monospace”: system wybiera wtedy swoją
// czcionkę danego rodzaju z prawdziwym pogrubieniem. Tam, gdzie krój jest, nic się nie zmienia.
const DOC_FONT_SANS = /^(calibri( light)?|aptos( display| narrow)?|arial( narrow| black)?|arial mt|helvetica( neue)?|verdana|tahoma|segoe ui( light| semibold)?|trebuchet ms|century gothic|franklin gothic( book| medium| demi)?|candara|corbel|gill sans( mt)?|open sans|roboto|lato|bahnschrift|carlito|liberation sans|dejavu sans|noto sans|source sans pro|montserrat|lucida sans( unicode)?|microsoft sans serif|ms sans serif)$/i;
const DOC_FONT_SERIF = /^(cambria|times new roman|times|georgia|garamond|book antiqua|palatino( linotype)?|constantia|caladea|century schoolbook|century|liberation serif|dejavu serif|noto serif|baskerville( old face)?|bookman old style|cambria math)$/i;
const DOC_FONT_MONO = /^(consolas|courier new|courier|lucida console|menlo|monaco|liberation mono|dejavu sans mono|cascadia (code|mono))$/i;

function withGenericFontFallback(value) {
  const v = String(value || "").trim();
  if (!v || v.includes(",") || v.includes("var(")) return null;
  const name = v.replace(/^["']|["']$/g, "").trim();
  const generic = DOC_FONT_SANS.test(name) ? "sans-serif" : DOC_FONT_SERIF.test(name) ? "serif" : DOC_FONT_MONO.test(name) ? "monospace" : "";
  return generic ? `${v}, ${generic}` : null;
}

function addGenericFontFallbacks(host) {
  if (!host) return 0;
  let n = 0;
  // reguły z <style> docx-preview (style akapitów/znaków, motyw); @font-face pomijamy
  host.querySelectorAll("style").forEach((el) => {
    let rules;
    try { rules = el.sheet?.cssRules; } catch (_) { rules = null; }
    if (!rules) return;
    for (const rule of rules) {
      if (!rule.style || rule.type !== 1) continue; // tylko CSSStyleRule
      const ff = withGenericFontFallback(rule.style.fontFamily);
      if (ff) { rule.style.fontFamily = ff; n++; }
      for (let i = 0; i < rule.style.length; i++) {
        const prop = rule.style[i];
        if (!/^--docx-.*-font$/.test(prop)) continue;
        const fv = withGenericFontFallback(rule.style.getPropertyValue(prop));
        if (fv) { rule.style.setProperty(prop, fv); n++; }
      }
    }
  });
  // style w treści (fragmenty z własnym krojem)
  host.querySelectorAll('[style*="font-family"], [style*="--docx-"]').forEach((el) => {
    const ff = withGenericFontFallback(el.style.fontFamily);
    if (ff) { el.style.fontFamily = ff; n++; }
    for (let i = 0; i < el.style.length; i++) {
      const prop = el.style[i];
      if (!/^--docx-.*-font$/.test(prop)) continue;
      const fv = withGenericFontFallback(el.style.getPropertyValue(prop));
      if (fv) { el.style.setProperty(prop, fv); n++; }
    }
  });
  return n;
}
