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

const dwbPageBreaks = (() => {
  const KEY = "dwb-page-breaks-v1";
  const box = document.getElementById("showPageBreaks");
  let timer = 0;

  function enabled() {
    try { return localStorage.getItem(KEY) !== "0"; } catch (_) { return true; }
  }

  function host() {
    return docCanvasEl?.querySelector(".docx-preview-host");
  }

  function clear(h = host()) {
    h?.querySelectorAll(".dwb-page-break").forEach((el) => el.remove());
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
      range.selectNodeContents(el);
      const rects = [...range.getClientRects()].filter((r) => r.height > 0).sort((a, b) => a.top - b.top);
      if (!rects.length) {
        const r = el.getBoundingClientRect();
        if (r.height) add(el, [[r.top, r.bottom]]);
        return;
      }
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
  function paginateSection(sec, { includeNotes = false } = {}) {
    const article = sec.querySelector(":scope > article");
    if (!article || !sec.offsetHeight) return null;
    const secRect = sec.getBoundingClientRect();
    const scale = secRect.height / sec.offsetHeight || 1; // zoom Widoku desktopowego (transform)
    const cs = getComputedStyle(sec);
    const padT = parseFloat(cs.paddingTop) || 0;
    const padB = parseFloat(cs.paddingBottom) || 0;
    const pageW = sec.offsetWidth;
    const pageH = parseFloat(cs.minHeight) || pageW * Math.SQRT2; // bez rozmiaru strony: A4
    const bodyH = pageH - padT - padB;
    const contentTop = (article.getBoundingClientRect().top - secRect.top) / scale;
    const out = { sec, pageW, pageH, padT, padB, bodyH, contentTop, cuts: [], end: contentTop };
    if (!(bodyH > 40)) return out;
    const u = lineUnits(article, secRect.top, scale);
    if (includeNotes) sec.querySelectorAll(":scope > ol").forEach((ol) => u.push(...lineUnits(ol, secRect.top, scale)));
    u.sort((a, b) => a.top - b.top);
    let limit = pageH - padB; // dół treści pierwszej strony
    let first = 0; // indeks pierwszej linijki bieżącej strony
    for (let k = 0; k < u.length; k++) {
      if (u[k].bottom <= limit + 1) continue;
      if (k > first) {
        const b = pageStartIndex(u, k, first);
        const prevBottom = u[b - 1]?.bottom;
        const y = prevBottom != null && prevBottom <= u[b].top ? (prevBottom + u[b].top) / 2 : u[b].top;
        out.cuts.push({ y, top: u[b].top });
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
    clear(h);
    let page = 1;
    const editable = !!h.querySelector(".docx-edit-root");
    h.querySelectorAll(".docx-wrapper > section.docx").forEach((sec, si) => {
      if (si > 0) page += 1; // jawny podział = nowa kartka = nowa strona
      const res = paginateSection(sec);
      if (!res) return;
      for (const cut of res.cuts) {
        page += 1;
        const el = document.createElement("div");
        el.className = "dwb-page-break";
        el.setAttribute("aria-hidden", "true");
        if (editable) el.contentEditable = "false"; // w polu edycji dokumentu — nie da się w nim pisać
        el.dataset.label = t("pageBreakLabel", { n: page });
        el.style.top = `${Math.round(cut.y)}px`;
        sec.appendChild(el);
      }
    });
  }

  function schedule(delay = 150) {
    clearTimeout(timer);
    timer = setTimeout(() => {
      const idle = window.requestIdleCallback || ((fn) => setTimeout(fn, 0));
      idle(compute, { timeout: 600 });
    }, delay);
  }

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
  }
  document.fonts?.ready?.then(() => schedule(0));

  return { schedule, compute, enabled, paginate };
})();
