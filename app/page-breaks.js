// page-breaks.js — granice stron w Widoku desktopowym (opcja w panelu Widok, domyślnie wł.).
//
// docx-preview zaczyna nową kartkę tylko przy jawnym podziale strony/sekcji — dłuższy tekst
// bez podziałów to jedna wysoka „kartka”. Tu liczymy, gdzie skończyłaby się strona: wysokość
// strony i marginesy z pliku (min-height i padding kartki), a granica idzie PRZED pierwszą
// linijką, która się nie mieści (nigdy przez tekst; wiersz tabeli = całość), z zasadami Worda:
// „razem z następnym” i „wiersze razem” ze stylów pliku, wdowy i sieroty. Przerywana linia
// + „str. N” na marginesie. Odstępy wierszy jak w Wordzie daje applyWordLineMetrics
// (docx-render-fixes.js). Dalej przybliżenie (np. Word dzieli wiersze tabel, inne czcionki
// na Macu) — granica może czasem wypaść o linijkę inaczej.
//
// Znacznik to tylko nakładka (position:absolute w kartce, etykieta z ::after — bez tekstu
// w DOM), więc nie przesuwa tekstu, nie wchodzi do liczenia słów, szukania ani zapisu.
// W Widoku desktopowym zoom to transform — położenia w układzie nie zależą od zoomu, więc
// przeliczamy tylko, gdy zmienia się treść (pisanie, przebudowa, obrazy, czcionki).

// Odstępy między stronami (domyślnie, decyzja 2026-10-04): jak „Układ wydruku” w Wordzie —
// dolny margines strony, przerwa między kartkami i górny margines następnej. Robi to niewidoczny
// odstęp (div.dwb-page-gap, nieedytowalny) PRZED pierwszym blokiem nowej strony, a pas przerwy
// (.dwb-page-gap-band) rysuje się nad nim. Gdy strona kończy się w środku akapitu, akapit dzieli
// się jak w Wordzie (decyzja Mateusza 2026-10-05; dawniej cały akapit szedł na następną stronę
// i zostawiał dziurę): przed akapitem stoją dwa niewidoczne elementy pływające (float) — pusty
// „dystans” do linijki granicy i pas pełnej szerokości, który spycha wiersze od niej w dół.
// W edytowanym akapicie nic nie przybywa (kursor, zaznaczanie, zapis bez zmian). Wyjątki Worda
// zostają: wiersze razem, wdowy, sieroty, „razem z następnym”. Tabela przechodzi w całości, gdy
// się mieści. Druk i Podgląd wydruku dzielą osobno (print-preview.js), też jak Word. Akapit
// niedzielony (tabela, akapit-pojemnik) dłuższy niż strona: zostaje sama kreska. Opcja „Ukryj
// biały obszar” (jak w Wordzie, też dwuklik w przerwę) = dawna kreska styku kartek.

const dwbPageBreaks = (() => {
  const KEY = "dwb-page-breaks-v1";
  const GAP_KEY = "dwb-page-gaps-v1";
  const GAP = 22; // przerwa między kartkami — jak między kartkami z jawnym podziałem
  const box = document.getElementById("showPageBreaks");
  const gapBox = document.getElementById("hideWhiteSpace");
  let timer = 0;

  function enabled() {
    try { return localStorage.getItem(KEY) !== "0"; } catch (_) { return true; }
  }
  function gapsOn() {
    try { return localStorage.getItem(GAP_KEY) !== "0"; } catch (_) { return true; }
  }
  function setGaps(on) {
    try { localStorage.setItem(GAP_KEY, on ? "1" : "0"); } catch (_) { /* prywatne */ }
    if (gapBox) gapBox.checked = !on;
    compute();
  }

  function host() {
    return docCanvasEl?.querySelector(".docx-preview-host");
  }

  function clear(h = host()) {
    h?.querySelectorAll(".dwb-page-break, .dwb-page-gap, .dwb-page-gap-spacer, .dwb-page-gap-band").forEach((el) => el.remove());
    h?.querySelectorAll("section.dwb-gapped").forEach((sec) => { sec.classList.remove("dwb-gapped"); sec.style.removeProperty("--dwb-sec-min"); });
  }

  // „Razem z następnym” / „Wiersze razem” ze stylów pliku (z dziedziczeniem po stylu bazowym)
  // — Word nie zostawia nagłówka na dole strony. Klasa akapitu w podglądzie = docx_<styleId>.
  let keep = { bytes: null, next: new Set(), lines: new Set() };
  async function ensureKeepInfo() {
    if (keep.bytes === originalFileBytes) return;
    keep = { bytes: originalFileBytes, next: new Set(), lines: new Set() };
    try {
      const zip = await loadDocxZipCached(originalFileBytes);
      const xml = await zip.file("word/styles.xml")?.async("string");
      if (!xml) return;
      const styles = [...composeParse(xml).getElementsByTagNameNS(W_NS, "style")].filter((st) => st.getAttributeNS(W_NS, "type") === "paragraph");
      const byId = new Map(styles.map((st) => [st.getAttributeNS(W_NS, "styleId"), st]));
      const flag = (st, name, depth = 0) => {
        if (!st || depth > 12) return false;
        const pPr = composeDirectChild(st, "pPr");
        const el = pPr && composeDirectChild(pPr, name);
        if (el) return !/^(0|false|off)$/i.test(el.getAttributeNS(W_NS, "val") || "");
        const based = composeDirectChild(st, "basedOn")?.getAttributeNS(W_NS, "val");
        return based ? flag(byId.get(based), name, depth + 1) : false;
      };
      byId.forEach((st, id) => {
        if (flag(st, "keepNext")) keep.next.add(docxStyleClassName(id));
        if (flag(st, "keepLines")) keep.lines.add(docxStyleClassName(id));
      });
    } catch (_) { /* uszkodzony styles.xml — bez tych zasad */ }
  }
  const hasClass = (el, set) => !!el && [...el.classList].some((c) => set.has(c));

  // Wiersze tekstu elementu: [[góra, dół], …] w układzie okna (prostokąty znaków złączone w linijki)
  function textLines(el, range = document.createRange()) {
    range.selectNodeContents(el);
    const rects = [...range.getClientRects()].filter((r) => r.height > 0).sort((a, b) => a.top - b.top);
    if (!rects.length) return [];
    const lines = [];
    let top = rects[0].top;
    let bottom = rects[0].bottom;
    for (const r of rects.slice(1)) {
      if (r.top < bottom - 2) { bottom = Math.max(bottom, r.bottom); continue; } // ta sama linijka
      lines.push([top, bottom]);
      top = r.top;
      bottom = r.bottom;
    }
    lines.push([top, bottom]);
    return lines;
  }

  // Czy wiersze akapitu opłyną odstęp pływający przed nim: akapit i jego pojemniki aż do <article>
  // to zwykłe bloki (flex, grid, tabela, overflow ≠ visible to osobny kontekst — taki akapit
  // przeszedłby pod pas w całości, więc granica idzie przed nim jak dawniej).
  function canSplit(el, article) {
    for (let x = el; x && x !== article; x = x.parentElement) {
      const cs = getComputedStyle(x);
      if (!/^(block|list-item)$/.test(cs.display) || cs.overflow !== "visible" || cs.cssFloat !== "none" || /^(absolute|fixed)$/.test(cs.position)) return false;
    }
    return true;
  }

  // Linijki treści kartki (y w układzie kartki, bez zoomu): akapity rozbite na wiersze tekstu,
  // wiersz tabeli jako całość. Każda linijka wie, z którego akapitu jest i którym jest wierszem.
  function lineUnits(article, secTop, scale) {
    const out = [];
    const range = document.createRange();
    const add = (el, lines, row = false) => {
      const next = !row && hasClass(el, keep.next);
      const together = !row && hasClass(el, keep.lines);
      lines.forEach(([top, bottom], i) => out.push({ top: (top - secTop) / scale, bottom: (bottom - secTop) / scale,
        el, i, n: lines.length, row, next, together }));
    };
    article.querySelectorAll("p, tr").forEach((el) => {
      if (el.tagName === "TR") {
        if (el.parentElement?.closest("tr")) return; // tabela w tabeli — liczy się zewnętrzny wiersz
        const r = el.getBoundingClientRect();
        if (r.height) add(el, [[r.top, r.bottom]], true);
        return;
      }
      if (el.closest("table")) return;
      const lines = textLines(el, range);
      if (!lines.length) {
        const r = el.getBoundingClientRect();
        if (r.height) add(el, [[r.top, r.bottom]]);
        return;
      }
      add(el, lines);
    });
    return out.sort((a, b) => a.top - b.top);
  }

  // Gdzie naprawdę zaczyna się następna strona, gdy linijka u[k] się nie mieści — zasady Worda:
  // wiersze razem, sierota (sam pierwszy wiersz na dole), wdowa (sam ostatni na górze) i
  // „razem z następnym” (nagłówek idzie z akapitem). Nigdy przed początkiem bieżącej strony.
  function pageStartIndex(u, k, first) {
    let b = k;
    const cur = u[b];
    if (!cur.row && cur.i > 0) {
      const head = b - cur.i; // pierwsza linijka tego akapitu
      if (cur.together || cur.i === 1) b = head; // wiersze razem / sierota → cały akapit dalej
      else if (cur.n - cur.i === 1 && cur.n >= 3) b -= 1; // wdowa → jeszcze jedna linijka dalej
      if (b <= first) b = k;
    }
    // poprzedzające akapity „razem z następnym” idą razem z tym, co zaczyna nową stronę
    for (let guard = 0; guard < 20; guard++) {
      const prev = u[b - 1];
      if (!prev || prev.row || !prev.next || prev.el === u[b].el) break;
      const head = b - 1 - prev.i;
      if (head <= first) break;
      b = head;
    }
    return b;
  }

  // Strony jednej kartki podglądu (section.docx): gdzie zaczyna się każda kolejna strona.
  // Wynik w układzie kartki bez zoomu: contentTop (początek treści 1. strony), cuts — dla każdej
  // następnej strony { y: miejsce znacznika między linijkami, top: pierwsza linijka strony },
  // end — dół ostatniej linijki. Wspólne dla znaczników w Edycji i podglądu wydruku
  // (print-preview.js — tam też z przypisami pod treścią: includeNotes).
  // Najwyższy element w <article> zawierający el (akapit albo cała tabela).
  function topBlock(el, article) {
    let b = el;
    while (b && b.parentElement !== article) b = b.parentElement;
    return b;
  }

  function paginateSection(sec, { includeNotes = false, blockBreaks = false } = {}) {
    const article = sec.querySelector(":scope > article");
    if (!article || !sec.offsetHeight) return null;
    const secRect = sec.getBoundingClientRect();
    const scale = secRect.height / sec.offsetHeight || 1; // zoom Widoku desktopowego (transform)
    const cs = getComputedStyle(sec);
    const padT = parseFloat(cs.paddingTop) || 0;
    const padB = parseFloat(cs.paddingBottom) || 0;
    const pageW = sec.offsetWidth;
    // wysokość strony = min-height kartki z pliku; przy odstępach min-height kartki to suma stron
    // (klasa dwb-gapped) — wtedy z zapamiętanej wartości
    if (!sec.classList.contains("dwb-gapped")) sec.dataset.dwbPageH = String(parseFloat(cs.minHeight) || pageW * Math.SQRT2); // bez rozmiaru strony: A4
    const pageH = parseFloat(sec.dataset.dwbPageH) || pageW * Math.SQRT2;
    const bodyH = pageH - padT - padB;
    const contentTop = (article.getBoundingClientRect().top - secRect.top) / scale;
    const out = { sec, pageW, pageH, padT, padB, bodyH, contentTop, cuts: [], end: contentTop };
    if (!(bodyH > 40)) return out;
    const u = lineUnits(article, secRect.top, scale);
    // Położenia BEZ obecnych odstępów między stronami (każdy przesuwa dalszą treść o data-shift) —
    // przeliczenie nie musi ich zdejmować (dawniej zdejmowanie i wstawianie co chwilę przestawiało
    // przewijanie: skok do sekcji, „Wróć”, zoom, zmiana widoku).
    const gaps = [...article.querySelectorAll(":scope > .dwb-page-gap")].map((g) => ({ top: (g.getBoundingClientRect().top - secRect.top) / scale, shift: parseFloat(g.dataset.shift) || 0 }));
    if (gaps.length) u.forEach((x) => { let d = 0; for (const g of gaps) if (g.top <= x.top) d += g.shift; x.top -= d; x.bottom -= d; });
    if (includeNotes) sec.querySelectorAll(":scope > ol").forEach((ol) => u.push(...lineUnits(ol, secRect.top, scale)));
    u.sort((a, b) => a.top - b.top);
    let limit = pageH - padB; // dół treści pierwszej strony
    let first = 0; // indeks pierwszej linijki bieżącej strony
    for (let k = 0; k < u.length; k++) {
      if (u[k].bottom <= limit + 1) continue;
      if (k > first) {
        let b = pageStartIndex(u, k, first);
        let block = null;
        let line = 0; // > 0: granica w środku akapitu, przed jego wierszem nr line
        if (blockBreaks) {
          if (u[b].row) {
            // tabela przechodzi w całości (gdy się mieści na stronie)
            const head = u.findIndex((x) => x.row && topBlock(x.el, article) === topBlock(u[b].el, article));
            if (head > first) b = head;
            if (u[b] === u.find((x) => x.row && topBlock(x.el, article) === topBlock(u[b].el, article))) block = topBlock(u[b].el, article);
          } else if (u[b].i === 0) {
            block = topBlock(u[b].el, article);
          } else if (canSplit(u[b].el, article)) {
            block = topBlock(u[b].el, article); // akapit dzieli się jak w Wordzie
            line = u[b].i;
          } else {
            const head = b - u[b].i; // akapit, którego wiersze nie opłyną odstępu — w całości
            if (head > first) { b = head; block = topBlock(u[b].el, article); }
          }
        }
        const prevBottom = u[b - 1]?.bottom;
        const y = prevBottom != null && prevBottom <= u[b].top ? (prevBottom + u[b].top) / 2 : u[b].top;
        out.cuts.push({ y, top: u[b].top, h: u[b].bottom - u[b].top, block, line, el: u[b].el });
        first = b;
        limit = u[b].top + bodyH;
        k = b - 1; // od nowej strony sprawdzamy jeszcze raz
      } else {
        // jedna rzecz wyższa niż strona (duży obraz, wysoki wiersz) — granica w środku niej
        while (u[k].bottom > limit + 1) {
          out.cuts.push({ y: limit, top: limit });
          limit += bodyH;
        }
        first = k + 1;
      }
    }
    out.end = u.length ? Math.max(...u.map((x) => x.bottom)) : contentTop;
    return out;
  }

  // Wszystkie kartki podglądu po kolei (podgląd wydruku liczy na osobnym, niewidocznym renderze).
  async function paginate(h, opts) {
    await ensureKeepInfo();
    return [...h.querySelectorAll(".docx-wrapper > section.docx")].map((sec) => paginateSection(sec, opts)).filter(Boolean);
  }

  let run = 0;
  async function compute() {
    timer = 0;
    const myRun = ++run;
    const h = host();
    if (!h) return;
    const on = enabled() && !(typeof shouldUseMobileReflow === "function" && shouldUseMobileReflow());
    rootEl.classList.toggle("page-breaks-on", on);
    if (!on || docCanvasEl.classList.contains("hidden")) { clear(h); return; }
    await ensureKeepInfo();
    if (myRun !== run) return; // w międzyczasie zaczęło się nowsze liczenie
    let page = 1;
    const editable = !!h.querySelector(".docx-edit-root");
    const gaps = gapsOn();
    const plan = [];
    h.querySelectorAll(".docx-wrapper > section.docx").forEach((sec) => {
      const res = paginateSection(sec, { blockBreaks: gaps });
      if (res) plan.push(res);
    });
    // co powinno stać w kartkach: odstępy (przed którym blokiem, jaka wysokość), pasy, kreski
    const want = [];
    plan.forEach((res, si) => {
      if (si > 0) page += 1; // jawny podział = nowa kartka = nowa strona
      let shift = 0; // o ile przesunęły się dalsze linijki przez odstępy
      let sheetTop = 0; // góra bieżącej strony (układ kartki, po przesunięciach)
      const items = [];
      for (const cut of res.cuts) {
        page += 1;
        const label = t("pageBreakLabel", { n: page });
        if (gaps && cut.block) {
          const block = cut.block;
          const move = sheetTop + res.pageH + GAP + res.padT - (cut.top + shift); // pierwsza linijka nowej strony
          if (cut.line) {
            // w środku akapitu: pas pływający od linijki granicy (at), wysokość dociągana po wstawieniu
            items.push({ kind: "gap", block, line: cut.line, el: cut.el, at: Math.round((cut.top + shift) * 10) / 10, inset: Math.round(Math.max(1, cut.h / 2) * 10) / 10, height: Math.round(move * 10) / 10, shift: move, bandTop: Math.round(sheetTop + res.pageH), label });
            shift += move;
            sheetTop += res.pageH + GAP;
            continue;
          }
          let prev = block.previousElementSibling;
          while (prev?.matches(".dwb-page-gap, .dwb-page-gap-spacer")) prev = prev.previousElementSibling;
          const mb = prev ? Math.max(0, parseFloat(getComputedStyle(prev).marginBottom) || 0) : 0;
          const mt = Math.max(0, parseFloat(getComputedStyle(block).marginTop) || 0);
          // z odstępem między blokami ich marginesy już się nie łączą (mb + mt zamiast max)
          items.push({ kind: "gap", block, line: 0, height: Math.max(0, Math.round((move - Math.min(mb, mt)) * 10) / 10), shift: move, bandTop: Math.round(sheetTop + res.pageH), label });
          shift += move;
          sheetTop += res.pageH + GAP;
        } else {
          items.push({ kind: "seam", top: Math.round(cut.y + shift), label });
          sheetTop = cut.top + shift - res.padT;
        }
      }
      want.push({ sec: res.sec, items, minH: gaps && sheetTop > 0 ? Math.ceil(sheetTop + res.pageH) : 0 });
    });
    // bez zmian (np. przeliczenie po przewinięciu, zoomie, skoku) — nic nie ruszamy
    const same = want.every(({ sec, items, minH }) => {
      const gapEls = [...sec.querySelectorAll(":scope > article > .dwb-page-gap")];
      const bands = [...sec.querySelectorAll(":scope > .dwb-page-gap-band")];
      const seams = [...sec.querySelectorAll(":scope > .dwb-page-break")];
      const wg = items.filter((x) => x.kind === "gap");
      const ws = items.filter((x) => x.kind === "seam");
      return gapEls.length === wg.length && bands.length === wg.length && seams.length === ws.length
        && wg.every((x, i) => gapEls[i].nextElementSibling === x.block && (gapEls[i].dataset.line || "0") === String(x.line)
          && (x.line ? Math.abs((parseFloat(gapEls[i].dataset.at) || 0) - x.at) < 0.6 && Math.abs((parseFloat(gapEls[i].dataset.shift) || 0) - x.shift) < 0.6 : Math.abs(parseFloat(gapEls[i].style.height) - x.height) < 0.6)
          && bands[i].dataset.label === x.label && parseFloat(bands[i].style.top) === x.bandTop)
        && ws.every((x, i) => parseFloat(seams[i].style.top) === x.top && seams[i].dataset.label === x.label)
        && (sec.style.getPropertyValue("--dwb-sec-min") || "") === (minH ? `${minH}px` : "");
    }) && h.querySelectorAll(".dwb-page-gap, .dwb-page-gap-spacer, .dwb-page-gap-band, .dwb-page-break").length === want.reduce((n, w) => n + w.items.reduce((m, x) => m + (x.kind === "gap" ? (x.line ? 3 : 2) : 1), 0), 0);
    if (same) return;
    // zmiana: akapit widoczny u góry zostaje na swoim miejscu na ekranie (jak zakotwiczenie przewijania);
    // tuż po przerysowaniu — akapit sprzed przerysowania (document.js, reloadFromBytes)
    const pending = typeof takeRenderScrollAnchor === "function" ? takeRenderScrollAnchor() : null;
    const anchor = pending ? null : viewportAnchor(h);
    clear(h);
    const bandBg = canvasBackground();
    const hint = (el) => {
      el.dataset.hint = "";
      el.dataset.hintPl = I18N.pl[gaps ? "pageGapHint" : "pageSeamHint"];
      el.dataset.hintEn = I18N.en[gaps ? "pageGapHint" : "pageSeamHint"];
      el.dataset.hintDelay = "0.8";
    };
    const splits = []; // odstępy w środku akapitów — do dociągnięcia po wstawieniu wszystkiego
    want.forEach(({ sec, items, minH }) => {
      items.forEach((x) => {
        if (x.kind === "gap") {
          const gap = document.createElement("div");
          gap.className = "dwb-page-gap";
          gap.setAttribute("aria-hidden", "true");
          if (editable) gap.contentEditable = "false";
          gap.style.height = `${x.height}px`;
          gap.dataset.shift = String(x.shift);
          if (x.line) {
            const spacer = document.createElement("div");
            spacer.className = "dwb-page-gap-spacer";
            spacer.setAttribute("aria-hidden", "true");
            if (editable) spacer.contentEditable = "false";
            gap.classList.add("dwb-gap-split");
            gap.dataset.line = String(x.line);
            gap.dataset.at = String(x.at);
            // kolor kartki: zakrywa podświetlenie akapitu z kursorem na marginesach stron
            gap.style.background = getComputedStyle(sec).backgroundColor;
            x.block.before(spacer, gap);
            splits.push({ sec, spacer, gap, x });
          } else x.block.before(gap);
          const band = document.createElement("div");
          band.className = "dwb-page-gap-band";
          band.setAttribute("aria-hidden", "true");
          if (editable) band.contentEditable = "false";
          band.dataset.label = x.label;
          band.style.top = `${x.bandTop}px`;
          band.style.height = `${GAP}px`;
          if (bandBg) band.style.background = bandBg;
          hint(band);
          sec.appendChild(band);
          return;
        }
        const el = document.createElement("div");
        el.className = "dwb-page-break";
        el.setAttribute("aria-hidden", "true");
        if (editable) el.contentEditable = "false"; // w polu edycji dokumentu — nie da się w nim pisać
        el.dataset.label = x.label;
        el.style.top = `${x.top}px`;
        hint(el);
        sec.appendChild(el);
      });
      if (minH) { // ostatnia strona pełnej wysokości (kartka nie kończy się w pół)
        sec.classList.add("dwb-gapped");
        sec.style.setProperty("--dwb-sec-min", `${minH}px`);
      }
    });
    splits.forEach(fitSplit);
    if (typeof syncPageScaleHeightNow === "function") syncPageScaleHeightNow(); // inaczej przewinięcie niżej obcina stary, krótszy obszar
    if (pending) restoreDocScrollAnchor(pending);
    else if (anchor && docViewportEl) {
      const d = anchor.el.getBoundingClientRect().top - anchor.top;
      if (Math.abs(d) > 0.5) docViewportEl.scrollTop += d;
    }
  }

  // Odstęp w środku akapitu: „dystans” tak wysoki, żeby pas zaczął się w połowie linijki granicy
  // (at + inset; na samej jej górze zaokrąglenie potrafiło zepchnąć też linijkę wyżej — Chromium),
  // a pas tak wysoki, żeby ta linijka wypadła na początku treści następnej strony (at + shift).
  // Mierzone po wstawieniu (pas spycha całe pole wiersza, a liczymy po prostokątach znaków —
  // różnica to interlinia nad tekstem). Po kolei: każdy odstęp przesuwa tylko to, co za nim.
  function fitSplit({ sec, spacer, gap, x }) {
    // położenie kartki mierzone przy każdym odczycie: zmiana wysokości potrafi przesunąć przewinięcie
    // (zakotwiczenie przewijania przeglądarki) między pomiarami
    const local = (y) => { const r = sec.getBoundingClientRect(); return (y - r.top) / (r.height / sec.offsetHeight || 1); };
    const top0 = local(gap.getBoundingClientRect().top);
    spacer.style.height = `${Math.max(0, Math.round((x.at + x.inset - top0) * 10) / 10)}px`;
    const ln = textLines(x.el)[x.line];
    if (!ln) return;
    const err = local(ln[0]) - (x.at + x.shift);
    if (Math.abs(err) > 0.3) gap.style.height = `${Math.max(0, Math.round((x.height - err) * 10) / 10)}px`;
  }

  // Pierwszy akapit widoczny w obszarze dokumentu i jego położenie na ekranie.
  function viewportAnchor(h) {
    if (!docViewportEl) return null;
    const vpTop = docViewportEl.getBoundingClientRect().top;
    for (const p of h.querySelectorAll("section.docx > article p")) {
      const r = p.getBoundingClientRect();
      if (r.bottom > vpTop + 1 && r.height) return { el: p, top: r.top };
    }
    return null;
  }

  // Tło obszaru wokół kartek (pas przerwy ma je „przeciąć”)
  function canvasBackground() {
    for (let el = docCanvasEl; el && el !== document.documentElement; el = el.parentElement) {
      const bg = getComputedStyle(el).backgroundColor;
      if (bg && !/rgba\(0, 0, 0, 0\)|transparent/.test(bg)) return bg;
    }
    return "";
  }

  function schedule(delay = 150) {
    clearTimeout(timer);
    timer = setTimeout(() => {
      const idle = window.requestIdleCallback || ((fn) => setTimeout(fn, 0));
      idle(compute, { timeout: 600 });
    }, delay);
  }

  if (gapBox) {
    gapBox.checked = !gapsOn();
    gapBox.addEventListener("change", () => setGaps(!gapBox.checked));
  }
  // dwuklik w przerwę między stronami = ukryj / pokaż biały obszar (jak w Wordzie)
  docCanvasEl?.addEventListener("dblclick", (e) => {
    if (!e.target.closest?.(".dwb-page-gap-band, .dwb-page-break, .dwb-page-gap")) return;
    e.preventDefault();
    setGaps(!gapsOn());
  });
  if (box) {
    box.checked = enabled();
    box.addEventListener("change", () => {
      try { localStorage.setItem(KEY, box.checked ? "1" : "0"); } catch (_) {}
      compute();
    });
  }
  if (docCanvasEl) {
    // wysokość treści się zmienia (pisanie, przebudowa, inne okno w Widoku desktopowym)
    let lastH = 0;
    new ResizeObserver(() => {
      const hgt = docCanvasEl.offsetHeight;
      if (Math.abs(hgt - lastH) < 1) return;
      lastH = hgt;
      schedule(200);
    }).observe(docCanvasEl);
    docCanvasEl.addEventListener("input", () => schedule(400));
    docCanvasEl.addEventListener("load", (e) => { if (e.target.tagName === "IMG") schedule(200); }, true);
    // Enter / Backspace / sklejanie akapitów nie wysyłają „input”, a wysokość płótna często stoi
    // (ostatnia kartka ma wysokość pełnych stron) — obserwator rozmiaru milczał. Odstęp strony
    // zostawał przy starym akapicie: tekst wjeżdżał pod pas „str. N” albo zostawała pusta dziura
    // (zgłoszenie 2026-10-05). Każda zmiana akapitów → przeliczenie; własne odstępy pomijamy.
    const ours = (n) => n.nodeType === 1 && n.matches?.(".dwb-page-gap, .dwb-page-gap-spacer, .dwb-page-gap-band, .dwb-page-break");
    new MutationObserver((records) => {
      if (records.some((r) => [...r.addedNodes, ...r.removedNodes].some((n) => !ours(n)))) schedule(150);
    }).observe(docCanvasEl, { childList: true, subtree: true });
  }
  document.fonts?.ready?.then(() => schedule(0));

  return { schedule, compute, enabled, paginate, gapsOn, setGaps };
})();
