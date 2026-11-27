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
  // odstępy stron w kartce: między blokami, między wierszami tabeli (dwb-gap-row) i przed listą końcowych
  const GAP_SEL = ":scope > article > .dwb-page-gap, :scope > article > table .dwb-page-gap, :scope > .dwb-page-gap";
  const box = document.getElementById("showPageBreaks");
  const gapBox = document.getElementById("hideWhiteSpace");
  let timer = 0;
  let contentMo = null;

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

  // Żywe <li> przypisów dolnych wracają do macierzystego <ol> zanim usuniemy pojemniki stron
  // (.dwb-page-notes) — inaczej clear() usunęłoby edytowalne przypisy z DOM.
  function restoreFootnotes(h = host()) {
    if (!h) return;
    h.querySelectorAll("section.docx").forEach((sec) => {
      const home = sec.querySelector(":scope > ol.dwb-notes-footnote:not(.dwb-page-notes)");
      const extras = [...sec.querySelectorAll(":scope > ol.dwb-page-notes")];
      if (home && extras.length) {
        const items = [...home.querySelectorAll(":scope > li"), ...extras.flatMap((ol) => [...ol.querySelectorAll(":scope > li")])];
        items.sort((a, b) => (Number(a.dataset.dwbNoteOrder) || 0) - (Number(b.dataset.dwbNoteOrder) || 0));
        items.forEach((li) => home.appendChild(li));
      }
      extras.forEach((ol) => ol.remove());
      home?.classList.remove("dwb-notes-paged");
    });
  }
  function captureNoteCaret(h) {
    const sel = window.getSelection?.();
    if (!sel?.rangeCount) return null;
    const r = sel.getRangeAt(0);
    const el = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
    const p = el?.closest?.("ol.dwb-notes-footnote p[data-note-key]");
    if (!p || !h?.contains(p)) return null;
    return { p, startContainer: r.startContainer, startOffset: r.startOffset, endContainer: r.endContainer, endOffset: r.endOffset };
  }
  function restoreNoteCaret(saved) {
    if (!saved?.p?.isConnected) return;
    try {
      saved.p.focus({ preventScroll: true });
      const sel = window.getSelection?.();
      const r = document.createRange();
      r.setStart(saved.startContainer, saved.startOffset);
      r.setEnd(saved.endContainer, saved.endOffset);
      sel?.removeAllRanges();
      sel?.addRange(r);
    } catch (_) { /* zakres nieaktualny — zostaje fokus na przypisie */ }
  }

  function clear(h = host()) {
    restoreFootnotes(h);
    h?.querySelectorAll(".dwb-page-break, .dwb-page-gap, .dwb-page-gap-spacer, .dwb-page-gap-band, .dwb-page-margin-guide, .dwb-page-repeat, .dwb-page-hf").forEach((el) => el.remove());
    h?.querySelectorAll("section[data-dwb-hf-sig]").forEach((sec) => delete sec.dataset.dwbHfSig);
    h?.querySelectorAll("section.dwb-gapped").forEach((sec) => { sec.classList.remove("dwb-gapped"); sec.style.removeProperty("--dwb-sec-min"); });
  }

  // „Razem z następnym” / „Wiersze razem” ze stylów pliku (z dziedziczeniem po stylu bazowym)
  // — Word nie zostawia nagłówka na dole strony. Klasa akapitu w podglądzie = docx_<styleId>.
  // Kontrola bękartów (w:widowControl): Word ma ją włączoną, dopóki plik jej nie wyłączy — w
  // domyślnych akapitu (docDefaults), stylu albo samym akapicie (data-dwb-widow, łatka 22).
  let keep = { bytes: null, next: new Set(), lines: new Set(), noWidow: new Set(), widow: true };
  async function ensureKeepInfo() {
    if (keep.bytes === originalFileBytes) return;
    keep = { bytes: originalFileBytes, next: new Set(), lines: new Set(), noWidow: new Set(), widow: true };
    try {
      const zip = await loadDocxZipCached(originalFileBytes);
      const xml = await zip.file("word/styles.xml")?.async("string");
      if (!xml) return;
      const styles = [...composeParse(xml).getElementsByTagNameNS(W_NS, "style")].filter((st) => st.getAttributeNS(W_NS, "type") === "paragraph");
      const byId = new Map(styles.map((st) => [st.getAttributeNS(W_NS, "styleId"), st]));
      const on = (el) => !/^(0|false|off)$/i.test(el.getAttributeNS(W_NS, "val") || "");
      // wartość ze stylu albo jego bazowego; null — nigdzie w łańcuchu
      const flag = (st, name, depth = 0) => {
        if (!st || depth > 12) return null;
        const pPr = composeDirectChild(st, "pPr");
        const el = pPr && composeDirectChild(pPr, name);
        if (el) return on(el);
        const based = composeDirectChild(st, "basedOn")?.getAttributeNS(W_NS, "val");
        return based ? flag(byId.get(based), name, depth + 1) : null;
      };
      const dflt = composeParse(xml).getElementsByTagNameNS(W_NS, "pPrDefault")[0];
      const dWidow = dflt && [...dflt.getElementsByTagNameNS(W_NS, "widowControl")][0];
      if (dWidow) keep.widow = on(dWidow);
      byId.forEach((st, id) => {
        if (flag(st, "keepNext")) keep.next.add(docxStyleClassName(id));
        if (flag(st, "keepLines")) keep.lines.add(docxStyleClassName(id));
        if ((flag(st, "widowControl") ?? keep.widow) === false) keep.noWidow.add(docxStyleClassName(id));
      });
    } catch (_) { /* uszkodzony styles.xml — bez tych zasad */ }
  }
  const hasClass = (el, set) => !!el && [...el.classList].some((c) => set.has(c));

  // Wiersze tekstu elementu: [[góra, dół], …] w układzie okna (prostokąty znaków złączone w linijki).
  // Obrazy i inne „wyspy” w linii (rysunek, wykres, pole formularza) dokładamy osobno: prostokąty
  // zakresu ich nie dają, a wtedy akapit z obrazem był jedną niską linijką na dole obrazu — odstęp
  // strony liczony od niej zostawiał pas „str. N” w środku obrazu (zgłoszenie 2026-10-06).
  // (opakowanie wyspy [contenteditable=false] to span — jego prostokąt ma wysokość linijki, nie
  // obrazu; liczymy też to, co w środku: obraz, blok inline-block)
  const ATOMS = "img, svg, canvas, video, object, iframe, [contenteditable=false], [contenteditable=false] > *";
  function textLines(el, range = document.createRange()) {
    range.selectNodeContents(el);
    const rects = [...range.getClientRects()];
    el.querySelectorAll(ATOMS).forEach((a) => {
      if (a.closest(".dwb-page-gap, .dwb-page-gap-spacer, .dwb-anchor-abs, .dwb-anchor-wrap")) return; // obraz postawiony na kartce nie jest linijką akapitu
      const r = a.getBoundingClientRect();
      if (r.height > 0 && r.width > 0) rects.push(r);
    });
    rects.splice(0, rects.length, ...rects.filter((r) => r.height > 0).sort((a, b) => a.top - b.top));
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
  function canSplit(el) {
    for (let x = el; x && x.tagName !== "ARTICLE"; x = x.parentElement) {
      const cs = getComputedStyle(x);
      if (!/^(block|list-item)$/.test(cs.display) || cs.overflow !== "visible" || cs.cssFloat !== "none" || /^(absolute|fixed)$/.test(cs.position)) return false;
    }
    return !multiColumn(el.closest("article"));
  }

  // Sekcja w kilku kolumnach (w:cols) — <article> z column-count. Linijki kolumn leżą na tych samych
  // wysokościach, więc każdy zestaw kolumn (strona, patrz columnPages) liczymy jako całość, jak
  // wiersz tabeli: odstęp strony staje przed nim, nigdy w środku kolumn.
  function multiColumn(article) {
    if (!article) return false;
    const n = parseInt(getComputedStyle(article).columnCount, 10);
    return n > 1;
  }

  // Kolumny strona po stronie, jak w Wordzie: kolumna 1, kolumna 2, potem następna strona. Sam CSS
  // (column-count) rozkładał cały tekst sekcji na dwie długie kolumny przez wszystkie strony —
  // obok „Akapitu 1” stał „Akapit 36”, a granica strony cięła wiersze. Przed pierwszym blokiem
  // każdej strony stoi znacznik .dwb-col-break (column-span: all) — przerywa zestaw kolumn, więc
  // każda strona ma własny. Strony po kolei: bloki z sumy wysokości w szerokości kolumny (offsetHeight
  // — prostokąty fragmentów w kolumnach WebKit podaje błędnie), potem pomiar zestawu między
  // znacznikami; wystaje poza stronę — znacznik o blok wcześniej (wyrównywanie kolumn i podział
  // akapitu między nie bywa o linijkę inne niż suma).
  // Ostatnia strona kolumn: Word wypełnia najpierw pierwszą kolumnę (wyrównuje tylko przed sekcją
  // ciągłą — u nas następny <article> tej kartki). Wypełniacz .dwb-col-fill na końcu dopełnia
  // zestaw do pełnej strony, więc wyrównywanie CSS daje to samo: pełna kolumna 1, reszta w 2.
  // shortcut: granica strony między akapitami (Word dzieli też akapit) — poprawić, gdy plik z bazy
  // tego wymaga.
  const COL_SPAN = ".dwb-page-gap, .dwb-col-break";
  function columnPages(article, firstAvail, bodyH, scale) {
    const n = parseInt(getComputedStyle(article).columnCount, 10) || 1;
    article.querySelectorAll(":scope > .dwb-col-break, :scope > .dwb-col-fill").forEach((x) => x.remove());
    // odstępy stron z poprzedniego liczenia też przerywają zestaw kolumn (column-span) i zawyżały
    // pomiar — na czas układania schowane (compute i tak ustawia je od nowa przed zestawami)
    const hidden = [...article.querySelectorAll(":scope > .dwb-page-gap, :scope > .dwb-page-gap-spacer")];
    hidden.forEach((g) => { g.style.display = "none"; });
    try { layColumnPages(article, n, firstAvail, bodyH, scale); } finally { hidden.forEach((g) => { g.style.display = ""; }); }
  }
  function layColumnPages(article, n, firstAvail, bodyH, scale) {
    const blocks = [...article.children].filter((el) => !el.matches(".dwb-page-gap, .dwb-page-gap-spacer"));
    // kontrola wdów i sierot Worda: podział akapitu zostawia ≥ 2 wiersze z każdej strony, więc akapit
    // do 3 wierszy nie dzieli się między kolumny (Chrome przy wyrównywaniu łamał swoje widows: 2)
    for (const el of blocks) {
      if (el.tagName !== "P") continue;
      const lh = parseFloat(getComputedStyle(el).lineHeight) || (parseFloat(getComputedStyle(el).fontSize) || 12) * 1.2;
      el.classList.toggle("dwb-col-keep", el.offsetHeight / lh < 3.5);
    }
    const hs = blocks.map((el) => { const cs = getComputedStyle(el); return el.offsetHeight + (parseFloat(cs.marginTop) || 0) + (parseFloat(cs.marginBottom) || 0); });
    const edit = !!article.closest(".docx-edit-root");
    const mk = (cls) => { const x = document.createElement("div"); x.className = cls; x.setAttribute("aria-hidden", "true"); if (edit) x.contentEditable = "false"; return x; };
    // przed odstępem strony, jeśli już stoi przed blokiem (kolejność: znacznik, odstęp, blok)
    const at = (el) => { let x = el; while (x.previousElementSibling?.matches(".dwb-page-gap, .dwb-page-gap-spacer")) x = x.previousElementSibling; return x; };
    // góra zestawu: dół znacznika poprzedniej strony, inaczej góra sekcji (odstępy schowane)
    const setTop = (first, prevBr) => () => (prevBr ? prevBr.getBoundingClientRect().bottom : article.getBoundingClientRect().top);
    let i = 0, page = 0, prev = null, used = 0;
    for (;;) {
      const allowed = page === 0 ? Math.max(0, firstAvail) : bodyH;
      let j = i;
      used = 0;
      while (j < blocks.length && (j === i || used + hs[j] <= n * allowed)) used += hs[j++];
      if (j >= blocks.length) break; // reszta mieści się na tej stronie — ostatnia
      const br = mk("dwb-col-break");
      at(blocks[j]).before(br);
      const from = setTop(blocks[i], prev);
      const height = () => (br.getBoundingClientRect().top - from()) / scale;
      while (j > i + 1 && height() > allowed + 0.5) { j--; at(blocks[j]).before(br); }
      i = j; page++; prev = br;
      used = 0;
    }
    if (article.nextElementSibling?.matches("article")) {
      // przed sekcją ciągłą Word wyrównuje kolumny i zostawia odstęp „po” ostatniego akapitu (CSS
      // obcina marginesy na końcu kolumn — tekst pod kolumnami stał ~1 odstęp za wysoko)
      const lastBlock = blocks[blocks.length - 1];
      article.style.paddingBottom = lastBlock ? `${parseFloat(getComputedStyle(lastBlock).marginBottom) || 0}px` : "";
      return;
    }
    const allowed = page === 0 ? Math.max(0, firstAvail) : bodyH;
    let fillH = Math.floor(n * allowed - used);
    if (fillH <= 0) return;
    const fill = mk("dwb-col-fill");
    article.appendChild(fill);
    const lastTop = setTop(blocks[i], prev);
    const lastH = () => (article.getBoundingClientRect().bottom - lastTop()) / scale;
    // wypełniacz o tyle mniejszy, o ile zestaw wystaje (podział akapitów między kolumny)
    for (let guard = 0; guard < 6 && fillH > 0; guard++) {
      fill.style.height = `${fillH}px`;
      const over = lastH() - allowed;
      if (over <= 0.5) break;
      fillH = Math.floor(fillH - Math.max(n * over, 4));
    }
    if (fillH <= 0) fill.remove();
  }

  // Linijki treści kartki (y w układzie kartki, bez zoomu): akapity rozbite na wiersze tekstu,
  // wiersz tabeli jako całość. Każda linijka wie, z którego akapitu jest i którym jest wierszem.
  function lineUnits(article, secTop, scale) {
    const out = [];
    const range = document.createRange();
    const add = (el, lines, row = false) => {
      // wpisane w akapit (data-dwb-*, łatka 22) wygrywa ze stylem
      const own = (k, set) => (el.dataset[k] ? el.dataset[k] === "1" : hasClass(el, set));
      const next = !row && own("dwbKeepNext", keep.next);
      const together = !row && own("dwbKeepLines", keep.lines);
      // styl bez klasy (Normalny) → domyślne pliku
      const widow = !row && (el.dataset.dwbWidow ? el.dataset.dwbWidow === "1" : !hasClass(el, keep.noWidow) && (keep.widow || [...el.classList].some((c) => c.startsWith("docx_"))));
      // Dół linijki do sprawdzania, czy mieści się na stronie — jak w Wordzie (porównanie z
      // PDF-ami): środek liter + POŁOWA wysokości pojedynczej interlinii kroju (rozmiar ×
      // współczynnik Worda); dodatkowy odstęp interlinii wielokrotnej może wystawać w dolny
      // margines. „Podsumowanie…” (pojedyncza): nagłówek, którego litery kończą się 0,2 px nad
      // marginesem, w Wordzie przechodzi na następną stronę; „umowa najmu” (1,17): ostatnia linijka
      // akapitu z pełną interlinią 2,8 px za marginesem zostaje. Góra bez zmian (góra liter) — od
      // niej liczą się odstępy stron w Edycji (fitSplit mierzy litery; inna góra = odstęp „pływał”).
      if (!row && lines.length) {
        const cs = getComputedStyle(el);
        const single = (parseFloat(cs.fontSize) || 0) * wordLineFactor(cs.fontFamily) * scale;
        if (single > 0) lines = lines.map(([t, b]) => [t, Math.max(b, (t + b) / 2 + single / 2)]);
      }
      lines.forEach(([top, bottom], i) => out.push({ top: (top - secTop) / scale, bottom: (bottom - secTop) / scale,
        el, i, n: lines.length, row, next, together, widow }));
    };
    if (multiColumn(article)) {
      // każdy zestaw kolumn (strona) = całość: od elementu „na całą szerokość” tuż przed pierwszym
      // blokiem (odstęp strony / znacznik) do następnego znacznika — prostokąty akapitów w kolumnach
      // WebKit podaje błędnie, a te elementy leżą zwyczajnie w przepływie
      const ar = article.getBoundingClientRect();
      const breaks = [...article.querySelectorAll(":scope > .dwb-col-break")];
      [null, ...breaks].forEach((br, k) => {
        let el = br ? br.nextElementSibling : article.firstElementChild;
        while (el?.matches(`${COL_SPAN}, .dwb-page-gap-spacer, .dwb-col-fill`)) el = el.nextElementSibling;
        if (!el) return;
        const before = el.previousElementSibling;
        const top = before?.matches(COL_SPAN) ? before.getBoundingClientRect().bottom : ar.top;
        const bottom = breaks[k] ? breaks[k].getBoundingClientRect().top : ar.bottom;
        if (bottom > top) { add(el, [[top, bottom]], true); out[out.length - 1].multi = true; }
      });
      return out;
    }
    article.querySelectorAll("p, tr").forEach((el) => {
      if (el.tagName === "TR") {
        if (el.parentElement?.closest("tr")) return; // tabela w tabeli — liczy się zewnętrzny wiersz
        const r = el.getBoundingClientRect();
        // Komórka scalona w pionie (rowspan) sięga w następne wiersze, a jej tekst może wystawać
        // poza ten wiersz. Linijki tekstu granica strony nie przetnie — wiersz kończy się pod
        // ostatnią linijką, która zaczyna się w nim (Word przenosi go w całości; „ubezp”: odstęp
        // przed akapitem + linijka wyższe niż wiersz, tekst wjeżdżał na następną kartkę).
        let bottom = r.bottom;
        for (const td of el.cells) {
          if (td.rowSpan > 1) td.querySelectorAll("p").forEach((p) => textLines(p, range).forEach(([t, b]) => { if (t < r.bottom - 0.5) bottom = Math.max(bottom, b); }));
        }
        // wiersz podzielony między strony (podgląd wydruku: splitRowsAtPages) — części osobno
        if (el.dataset.ppSplit && r.height) {
          let a = r.top;
          for (const y of el.dataset.ppSplit.split(",").map(Number).sort((x, z) => x - z)) {
            const at = secTop + y * scale;
            if (at > a + 0.5 && at < bottom - 0.5) { add(el, [[a, at]], true); a = at; }
          }
          add(el, [[a, bottom]], true);
          return;
        }
        if (r.height) {
          add(el, [[r.top, bottom]], true);
          if (bottom > r.bottom + 0.5) out[out.length - 1].spill = true; // tekst wystaje pod wiersz
        }
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
      if (cur.together || cur.widow && cur.i === 1) b = head; // wiersze razem / sierota → cały akapit dalej
      else if (cur.widow && cur.n - cur.i === 1 && cur.n >= 3) b -= 1; // wdowa → jeszcze jedna linijka dalej
      if (b <= first) b = k;
    }
    // wiersz tabeli, którego tekst (komórka scalona w pionie) wchodzi w wiersz zaczynający stronę:
    // linijka nie zmieściła się w części komórki na tej stronie — Word przenosi i ten wiersz
    for (let guard = 0; guard < 20 && u[b].row; guard++) {
      const prev = u[b - 1];
      if (!prev?.row || !prev.spill || b - 1 <= first || prev.bottom <= u[b].top + 0.5) break;
      b -= 1;
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
  // (kartka może mieć kilka <article> — sekcje ciągłe na jednej stronie; liczy się ten, w którym jest el)
  function topBlock(el) {
    const article = el.closest("article");
    let b = el;
    while (b && b.parentElement !== article) b = b.parentElement;
    return b;
  }

  // Wiersze nagłówka do powtórzenia, gdy strona zaczyna się wierszem `tr` tej samej tabeli co
  // `prevTr`: kolejne od początku tabeli z data-dwb-header (jak w Wordzie). { table, rows, top, h }
  // w układzie kartki; null, gdy tabela ich nie ma albo `tr` sam jest nagłówkiem.
  function repeatedHeader(tr, prevTr, secTop, scale) {
    const table = tr.closest("table");
    if (!table || prevTr.closest("table") !== table) return null;
    const rows = [];
    for (const r of table.rows) { if (r.dataset.dwbHeader !== "1") break; rows.push(r); }
    if (!rows.length || rows.includes(tr)) return null;
    const a = rows[0].getBoundingClientRect();
    const z = rows[rows.length - 1].getBoundingClientRect();
    const h = (z.bottom - a.top) / scale;
    return h > 0 ? { table, rows, top: (a.top - secTop) / scale, h } : null;
  }

  // Histereza granic stron (Edycja). Linijka, której dół leży niemal na dolnej granicy strony, raz
  // „mieściła się”, raz nie — przerwy stron mają ułamkowe wysokości i każde przeliczenie zmieniało
  // pozycje o ~0,1 px. Sierota/wdowa przełączały wtedy „cały akapit dalej” ↔ „podział po linijce”
  // w kółko (przewodnik: przewinięcie skakało i zamykało okienka — testy guide, chaos 101). W pasie
  // ±STICKY px od granicy zostaje decyzja z poprzedniego przeliczenia (el → linijka → mieści się?);
  // pisanie przesuwa treść o całe linijki, więc wychodzi poza pas. Podgląd wydruku — bez histerezy.
  const STICKY = 1.5;
  let lastFit = new WeakMap();
  // O ile odstęp naprawdę przesunął treść za nim: w środku akapitu fitSplit dopasowuje go do linijki,
  // a WebKit stawia linijki w pikselach ekranu — rzeczywiste przesunięcie różni się od planu (data-shift)
  // o ułamek piksela. Pomiar odejmujący plan widział przy następnym liczeniu „zmianę”, odstęp był
  // budowany od nowa o 1 px inaczej — i tak na zmianę (przewodnik w WebKit, ¶ ruszało stronę).
  function gapDisplacement(g) {
    const d = parseFloat(g.dataset.disp);
    return Number.isFinite(d) ? d : parseFloat(g.dataset.shift) || 0;
  }

  function paginateSection(sec, { includeNotes = false, blockBreaks = false, decided = null } = {}) {
    // sekcje ciągłe (Word: nowa sekcja na tej samej stronie) = kolejne <article> tej samej kartki
    const articles = [...sec.querySelectorAll(":scope > article")];
    const article = articles[0];
    if (!article || !sec.offsetHeight) return null;
    const secRect = sec.getBoundingClientRect();
    const scale = secRect.height / sec.offsetHeight || 1; // zoom Widoku desktopowego (transform)
    const cs = getComputedStyle(sec);
    const padT = parseFloat(cs.paddingTop) || 0;
    const padB = parseFloat(cs.paddingBottom) || 0;
    const padL = parseFloat(cs.paddingLeft) || 0;
    const padR = parseFloat(cs.paddingRight) || 0;
    const pageW = sec.offsetWidth;
    // wysokość strony = min-height kartki z pliku; przy odstępach min-height kartki to suma stron
    // (klasa dwb-gapped) — wtedy z zapamiętanej wartości
    if (!sec.classList.contains("dwb-gapped")) sec.dataset.dwbPageH = String(parseFloat(cs.minHeight) || pageW * Math.SQRT2); // bez rozmiaru strony: A4
    const pageH = parseFloat(sec.dataset.dwbPageH) || pageW * Math.SQRT2;
    // Stopka wyższa niż miejsce między jej odległością od krawędzi a dolnym marginesem wypycha
    // treść w górę — jak w Wordzie (S-99: stopka „S-99-P 2/26” na odległości = marginesowi, w
    // Wordzie wiersz mniej na stronie). docx-preview: margin-bottom stopki = odległość − margines.
    const footer = sec.querySelector(":scope > footer");
    let bottomZone = padB;
    if (footer && footer.offsetHeight) {
      const dist = padB + (parseFloat(getComputedStyle(footer).marginBottom) || 0);
      let fh = 0;
      for (const c of footer.children) fh += c.offsetHeight + (parseFloat(getComputedStyle(c).marginTop) || 0) + (parseFloat(getComputedStyle(c).marginBottom) || 0);
      if (dist >= 0 && fh > 0) bottomZone = Math.max(padB, dist + fh);
    }
    const bodyH = pageH - padT - bottomZone;
    // kolumny strona po stronie: znaczniki zestawów kolumn przed liczeniem linijek. Pierwsza strona
    // kolumn zaczyna się tam, gdzie sekcja w kolumnach (położenie bez odstępów stron nad nią).
    for (const a of articles) {
      if (!multiColumn(a)) continue;
      let top = (a.getBoundingClientRect().top - secRect.top) / scale;
      for (const g of sec.querySelectorAll(GAP_SEL)) if (g.compareDocumentPosition(a) & Node.DOCUMENT_POSITION_FOLLOWING) top -= gapDisplacement(g);
      columnPages(a, top < pageH - bottomZone ? pageH - bottomZone - top : bodyH, bodyH, scale);
    }
    const contentTop = (article.getBoundingClientRect().top - secRect.top) / scale;
    // Nagłówek wyższy niż górny margines spycha tekst KAŻDEJ strony (jak w Wordzie i w podglądzie
    // wydruku, który zaczyna treść stron od contentTop) — dawniej od strony 2 tekst zaczynał się na
    // marginesie, pod kopią nagłówka (testerzy AI 2026-11-27). Lustro bottomZone dla stopki.
    const topZone = Math.max(padT, contentTop);
    const nextBodyH = pageH - topZone - bottomZone;
    const out = { sec, pageW, pageH, padT, padB, padL, padR, bodyH, bottomZone, topZone, contentTop, cuts: [], end: contentTop };
    if (!(bodyH > 40)) return out;
    const u = articles.flatMap((a) => lineUnits(a, secRect.top, scale));
    // Położenia BEZ obecnych odstępów między stronami (każdy przesuwa dalszą treść o data-shift) —
    // przeliczenie nie musi ich zdejmować (dawniej zdejmowanie i wstawianie co chwilę przestawiało
    // przewijanie: skok do sekcji, „Wróć”, zoom, zmiana widoku).
    const gaps = [...sec.querySelectorAll(GAP_SEL)].map((g) => ({ top: (g.getBoundingClientRect().top - secRect.top) / scale, shift: gapDisplacement(g) }));
    const unshiftY = (y) => { if (!gaps.length) return y; let d = 0; for (const g of gaps) if (g.top <= y) d += g.shift; return y - d; };
    if (gaps.length) u.forEach((x) => { const d = x.top - unshiftY(x.top); x.top -= d; x.bottom -= d; });
    // Przypisy DOLNE na dole strony z odnośnikiem (jak w Wordzie; Edycja i podgląd wydruku: includeNotes).
    // Każda linijka zna przypisy, do których ma odnośnik; mieści się na stronie tylko razem z nimi
    // (i kreską nad przypisami przy pierwszym na stronie) — inaczej idzie na następną stronę.
    // Wynik: out.pageNotes[i] = przypisy strony i. Przypisy końcowe — dalej na końcu sekcji.
    let footOl = null, noteH = null, sepH = 0;
    if (includeNotes) {
      footOl = sec.querySelector(":scope > ol.dwb-notes-footnote:not(.dwb-page-notes)");
      const footItems = [...sec.querySelectorAll(":scope > ol.dwb-notes-footnote > li[data-dwb-note]")];
      if (footOl && footItems.length) {
        noteH = new Map();
        footItems.forEach((li, idx) => {
          if (li.dataset.dwbNoteOrder == null) li.dataset.dwbNoteOrder = String(idx);
          const lcs = getComputedStyle(li);
          noteH.set(li.dataset.dwbNote, li.getBoundingClientRect().height / scale + (parseFloat(lcs.marginTop) || 0) + (parseFloat(lcs.marginBottom) || 0));
        });
        const ocs = getComputedStyle(footOl);
        sepH = parseFloat(footOl.dataset.dwbSepH) || ((parseFloat(ocs.marginTop) || 0) + (parseFloat(ocs.paddingTop) || 0));
        if (sepH > 0) footOl.dataset.dwbSepH = String(sepH);
        const refs = [...sec.querySelectorAll(":scope > article sup[data-dwb-note^='footnote:']")].map((el) => {
          const r = el.getBoundingClientRect();
          return { id: el.dataset.dwbNote, el, y: unshiftY(((r.top + r.bottom) / 2 - secRect.top) / scale) };
        }).filter((r) => noteH.has(r.id));
        const used = new Set();
        u.forEach((x) => {
          x.notes = refs.filter((r) => !used.has(r) && x.el.contains(r.el) && (x.row || (r.y >= x.top - 1 && r.y <= x.bottom + 1))).map((r) => { used.add(r); return r.id; });
        });
      }
      sec.querySelectorAll(":scope > ol").forEach((ol) => {
        if (ol.classList.contains("dwb-notes-footnote") && noteH) return;
        const olUnits = lineUnits(ol, secRect.top, scale);
        if (gaps.length) olUnits.forEach((x) => { const d = x.top - unshiftY(x.top); x.top -= d; x.bottom -= d; });
        u.push(...olUnits);
      });
    }
    if (footOl && noteH) out.footOl = footOl;
    u.sort((a, b) => a.top - b.top);
    let pageNotes = new Set(), pageNoteH = 0;
    const noteExtra = (x) => {
      if (!noteH || !x.notes?.length) return 0;
      let h = 0;
      for (const id of x.notes) if (!pageNotes.has(id)) h += noteH.get(id);
      return h ? h + (pageNotes.size ? 0 : sepH) : 0;
    };
    const notesOf = (a, b) => { const ids = []; for (let i = a; i < b; i++) for (const id of u[i].notes || []) if (!ids.includes(id)) ids.push(id); return ids; };
    if (noteH) out.pageNotes = [];
    let limit = pageH - bottomZone; // dół treści pierwszej strony
    let first = 0; // indeks pierwszej linijki bieżącej strony
    for (let k = 0; k < u.length; k++) {
      const extra = noteExtra(u[k]);
      const lim = limit - pageNoteH - extra + 0.5;
      let fits = u[k].bottom <= lim;
      // tylko decyzje w pasie granicznym: po wstawieniu granicy te same linijki są sprawdzane jeszcze
      // raz wobec następnej strony — taki zapis nadpisałby decyzję graniczną
      if (decided && Math.abs(u[k].bottom - lim) < STICKY) {
        const prev = lastFit.get(u[k].el)?.get(u[k].i);
        if (prev !== undefined) fits = prev;
        if (!decided.has(u[k].el)) decided.set(u[k].el, new Map());
        decided.get(u[k].el).set(u[k].i, fits);
      }
      if (fits) {
        if (extra || u[k].notes?.length) { for (const id of u[k].notes) pageNotes.add(id); pageNoteH += extra; }
        continue;
      }
      if (k > first) {
        let b = pageStartIndex(u, k, first);
        let block = null;
        let line = 0; // > 0: granica w środku akapitu, przed jego wierszem nr line
        if (blockBreaks) {
          if (!u[b].el.closest("article")) {
            // przypis końcowy (<ol> poza <article>): odstęp strony przed całą listą, gdy zaczyna nową stronę
            const ol = u[b].el.closest("ol");
            if (ol && u[b].i === 0 && !u.slice(first, b).some((x) => x.el.closest("ol") === ol)) block = ol;
          } else if (u[b].multi) {
            block = u[b].el; // zestaw kolumn (strona) w całości: odstęp przed nim (column-span: all)
          } else if (u[b].row) {
            // Tabela dzieli się między wierszami, jak w Wordzie i w podglądzie wydruku (dawniej szła
            // w całości na następną stronę — Kolokwium 18.11: od strony 3 każda strona w Edycji inna niż
            // w Wordzie). W całości idzie, gdy na tej stronie zostałyby same wiersze nagłówka
            // powtarzanego albo nic; tabela w tabeli — przez wiersz tabeli zewnętrznej.
            const tbl = topBlock(u[b].el, article);
            const head = u.findIndex((x) => x.row && topBlock(x.el, article) === tbl);
            const tr = u[b].el.tagName === "TR" ? u[b].el : null;
            const own = tr && tr.closest("table") === tbl && tbl.tagName === "TABLE";
            const onlyHeaders = own && [...tbl.rows].slice(0, tr.rowIndex).every((r) => r.dataset.dwbHeader === "1");
            if (b === head || !own || onlyHeaders) {
              if (head > first) b = head;
              if (u[b] === u[head]) block = tbl;
            } else block = tr;
          } else if (u[b].i === 0) {
            block = topBlock(u[b].el, article);
          } else if (canSplit(u[b].el)) {
            block = topBlock(u[b].el, article); // akapit dzieli się jak w Wordzie
            line = u[b].i;
          } else {
            const head = b - u[b].i; // akapit, którego wiersze nie opłyną odstępu — w całości
            if (head > first) { b = head; block = topBlock(u[b].el, article); }
          }
        }
        const prevBottom = u[b - 1]?.bottom;
        const y = prevBottom != null && prevBottom <= u[b].top ? (prevBottom + u[b].top) / 2 : u[b].top;
        // tabela idzie dalej na nowej stronie, a ma wiersze „Powtórz jako wiersz nagłówka”
        // (łatka 14): Word rysuje je u góry strony — miejsce na nie odejmujemy od strony
        const header = u[b].row && u[b - 1]?.row ? repeatedHeader(u[b].el, u[b - 1].el, secRect.top, scale) : null;
        out.cuts.push({ y, top: u[b].top, h: u[b].bottom - u[b].top, block, line, el: u[b].el, header });
        if (noteH) { out.pageNotes.push(notesOf(first, b)); pageNotes = new Set(); pageNoteH = 0; }
        first = b;
        limit = u[b].top + nextBodyH - (header?.h || 0);
        k = b - 1; // od nowej strony sprawdzamy jeszcze raz
      } else {
        // jedna rzecz wyższa niż strona (duży obraz, wysoki wiersz) — granica w środku niej
        let firstCut = true;
        while (u[k].bottom > limit + 1) {
          out.cuts.push({ y: limit, top: limit });
          if (noteH) out.pageNotes.push(firstCut ? notesOf(first, k + 1) : []);
          firstCut = false;
          limit += nextBodyH;
        }
        if (noteH) { pageNotes = new Set(); pageNoteH = 0; }
        first = k + 1;
      }
    }
    if (noteH) out.pageNotes.push(notesOf(first, u.length));
    out.end = u.length ? Math.max(...u.map((x) => x.bottom)) : contentTop;
    return out;
  }

  // ── Wiersz tabeli na granicy strony (podgląd wydruku) ─────────────────────────────────────
  // Word dzieli wiersz bez „nie dziel” (w:cantSplit) między strony KAŻDĄ KOMÓRKĘ OSOBNO, między
  // linijkami: w komórce zostaje to, co się mieści, a reszta zaczyna się u góry następnej strony —
  // też gdy w sąsiedniej komórce granica wypada gdzie indziej. Dawniej wiersz szedł w całości na
  // następną stronę, a wyższy niż strona był cięty jedną kreską (przez linijki; „karty” z PDF:
  // druga kolumna 70 px wyżej niż w Wordzie). Na niewidocznym renderze wydruku w każdej takiej
  // komórce wstawiamy przed pierwszą niemieszczącą się linijkę odstęp do granicy strony, a wiersz
  // dostaje data-pp-split (miejsce podziału) — lineUnits liczy jego części osobno.
  function cellLines(td, Y, range) {
    const table = td.closest("table");
    const out = [];
    td.querySelectorAll("p, table").forEach((el) => {
      if (el.tagName === "TABLE") {
        if (el.parentElement?.closest("td, th") !== td) return; // tabela w tabeli — w całości
        const r = el.getBoundingClientRect();
        if (r.height) out.push({ top: Y(r.top), bottom: Y(r.bottom), el, i: 0, n: 1 });
        return;
      }
      if (el.closest("table") !== table) return;
      let lines = textLines(el, range);
      if (!lines.length) { const r = el.getBoundingClientRect(); if (r.height) lines = [[r.top, r.bottom]]; }
      const widow = el.dataset.dwbWidow ? el.dataset.dwbWidow === "1" : !hasClass(el, keep.noWidow) && (keep.widow || [...el.classList].some((c) => c.startsWith("docx_")));
      lines.forEach(([t, b], i) => out.push({ top: Y(t), bottom: Y(b), el, i, n: lines.length, widow }));
    });
    return out.sort((a, b) => a.top - b.top);
  }

  // Odstęp przed linijką `ln` komórki: jej górny brzeg (z półinterlinią) staje na `y`.
  function pushLineTo(ln, y, Y, range) {
    const el = ln.el;
    let spacer;
    if (el.tagName === "TABLE" || ln.i === 0) {
      spacer = document.createElement("div");
      spacer.className = "pp-row-split";
      el.before(spacer);
    } else {
      // w środku akapitu: pełnej szerokości blok w linii przed pierwszym znakiem linijki — tekst
      // przed nim zostaje w swoich wierszach (wyjustowanie też), dalszy idzie pod niego
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
      let at = null;
      for (let n = walker.nextNode(); n && !at; n = walker.nextNode()) {
        if (n.nodeType === 1) {
          if (n.matches(ATOMS) && Y(n.getBoundingClientRect().top) >= ln.top - 1) at = { node: n };
          continue;
        }
        for (let i = 0; i < n.data.length; i++) {
          range.setStart(n, i);
          range.setEnd(n, i + 1);
          const r = range.getBoundingClientRect();
          if (r.height && Y(r.top) >= ln.top - 1) { at = { node: n, i }; break; }
        }
      }
      if (!at) return false;
      spacer = document.createElement("span");
      spacer.className = "pp-row-split";
      Object.assign(spacer.style, { display: "inline-block", width: "100%", verticalAlign: "top", lineHeight: "0" });
      if (at.i != null) at.node.splitText(at.i).before(spacer);
      else at.node.before(spacer);
    }
    spacer.setAttribute("aria-hidden", "true");
    spacer.style.margin = "0";
    const lh = parseFloat(getComputedStyle(el).lineHeight);
    const half = el.tagName === "TABLE" || !(lh > 0) ? 0 : Math.max(0, (lh - (ln.bottom - ln.top)) / 2);
    const mt = el.tagName === "TABLE" || ln.i === 0 ? parseFloat(getComputedStyle(el).marginTop) || 0 : 0;
    // wysokość dociągana pomiarem (marginesy, wiersz z blokiem) — dwa podejścia wystarczają
    let hgt = Math.max(0, y - (ln.top - half - mt));
    for (let pass = 0; pass < 3; pass++) {
      spacer.style.height = `${hgt}px`;
      const lines = el.tagName === "TABLE" ? [[el.getBoundingClientRect().top]] : textLines(el, range);
      const now = lines.map(([t]) => Y(t)).find((t) => t >= y - half - mt - 2);
      if (now == null) break;
      const err = y - (now - half - (ln.i === 0 ? mt : 0));
      if (Math.abs(err) < 0.3) break;
      hgt = Math.max(0, hgt + err);
    }
    return true;
  }

  // Jeden podział na raz (potem przeliczenie stron — dalsza treść się przesunęła). true = zmiana.
  function splitRowsAtPages(sections) {
    const range = document.createRange();
    for (const S of sections) {
      const sec = S.sec;
      const secRect = sec.getBoundingClientRect();
      const scale = secRect.height / sec.offsetHeight || 1;
      const Y = (v) => (v - secRect.top) / scale;
      const starts = [S.contentTop, ...S.cuts.map((c) => c.top)];
      for (let j = 0; j < S.cuts.length; j++) {
        const limit = j === 0 ? S.pageH - S.bottomZone : starts[j] + S.bodyH - (S.cuts[j - 1].header?.h || 0);
        for (const tr of sec.querySelectorAll(":scope > article tr")) {
          if (tr.parentElement.closest("tr") || [...tr.cells].some((td) => td.rowSpan > 1)) continue;
          const r = tr.getBoundingClientRect();
          const top = Y(r.top), bottom = Y(r.bottom);
          if (!(top >= starts[j] - 0.5 && top < limit - 0.5 && bottom > limit + 0.5)) continue;
          // „nie dziel”: chyba że wiersz jest wyższy niż cała strona — wtedy Word i tak go dzieli
          // (zaczyna go na nowej stronie)
          if (tr.dataset.dwbCantSplit === "1" && !(top <= starts[j] + 1 && bottom - top > limit - starts[j] + 0.5)) continue;
          if ((tr.dataset.ppSplit || "").split(",").some((v) => v && Math.abs(+v - limit) < 1)) continue;
          // Word dzieli wiersz tylko wtedy, gdy w KAŻDEJ komórce coś zostaje na tej stronie
          // (S-99: komórka z dwulinijkowym akapitem — sierota — przenosi cały wiersz)
          const plan = [];
          let fits = true;
          for (const td of tr.cells) {
            const lines = cellLines(td, Y, range);
            let k = lines.findIndex((l) => l.bottom > limit + 0.5);
            if (k < 0) continue;
            const ln = lines[k];
            if (ln.widow && ln.i === 1) k -= 1; // sierota: cały akapit dalej
            else if (ln.widow && ln.n - ln.i === 1 && ln.n >= 3) k -= 1; // wdowa: jeszcze jedna linijka
            if (k <= 0) { fits = false; break; }
            plan.push(lines[k]);
          }
          // bez linijki do przeniesienia (za stronę wystaje tylko dół komórki) — jak dotąd w całości
          if (!fits || !plan.length) continue; // wiersz idzie w całości na następną stronę (jak dotąd)
          const ys = (tr.dataset.ppSplit || "").split(",").filter(Boolean);
          plan.forEach((ln) => pushLineTo(ln, limit, Y, range));
          tr.dataset.ppSplit = [...ys, String(Math.round(limit * 100) / 100)].join(",");
          return true;
        }
      }
    }
    return false;
  }

  // Wszystkie kartki podglądu po kolei (podgląd wydruku liczy na osobnym, niewidocznym renderze).
  async function paginate(h, opts) {
    await ensureKeepInfo();
    const run = () => [...h.querySelectorAll(".docx-wrapper > section.docx")].map((sec) => paginateSection(sec, opts)).filter(Boolean);
    let res = run();
    // wydruk: wiersze tabel dzielone między strony jak w Wordzie (zmienia render — tylko na kopii)
    if (opts?.splitRows) for (let guard = 0; guard < 200 && splitRowsAtPages(res); guard++) res = run();
    return res;
  }

  let run = 0;
  async function compute() {
    timer = 0;
    const myRun = ++run;
    const h = host();
    if (!h) return;
    const on = enabled() && !(typeof shouldUseMobileReflow === "function" && shouldUseMobileReflow());
    rootEl.classList.toggle("page-breaks-on", on);
    if (!on || docCanvasEl.classList.contains("hidden")) { clear(h); h.querySelectorAll(".dwb-col-break, .dwb-col-fill").forEach((b) => b.remove()); relayoutTextboxes(h); return; }
    await ensureKeepInfo();
    if (myRun !== run) return; // w międzyczasie zaczęło się nowsze liczenie
    let page = 1;
    const editable = !!h.querySelector(".docx-edit-root");
    const gaps = gapsOn();
    const plan = [];
    const decided = new WeakMap();
    h.querySelectorAll(".docx-wrapper > section.docx").forEach((sec) => {
      const res = paginateSection(sec, { includeNotes: gaps, blockBreaks: gaps, decided });
      if (res) plan.push(res);
    });
    lastFit = decided;
    // co powinno stać w kartkach: odstępy (przed którym blokiem, jaka wysokość), pasy, kreski, przypisy stron
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
          // wiersz tabeli na nowej stronie stoi pod powtórzonym nagłówkiem (cut.header, jak w Wordzie)
          const head = block.tagName === "TR" && cut.header ? cut.header : null;
          const move = sheetTop + res.pageH + GAP + res.topZone + (head?.h || 0) - (cut.top + shift); // pierwsza linijka nowej strony
          if (cut.line) {
            // w środku akapitu: pas pływający od linijki granicy (at), wysokość dociągana po wstawieniu
            items.push({ kind: "gap", block, line: cut.line, el: cut.el, at: Math.round((cut.top + shift) * 10) / 10, inset: Math.round(Math.max(1, cut.h / 2) * 10) / 10, height: Math.round(move * 10) / 10, shift: move, bandTop: Math.round(sheetTop + res.pageH), label });
            shift += move;
            sheetTop += res.pageH + GAP;
            continue;
          }
          let prev = block.previousElementSibling;
          while (prev?.matches(".dwb-page-gap, .dwb-page-gap-spacer, .dwb-col-break")) prev = prev.previousElementSibling;
          const mb = prev ? Math.max(0, parseFloat(getComputedStyle(prev).marginBottom) || 0) : 0;
          const mt = Math.max(0, parseFloat(getComputedStyle(block).marginTop) || 0);
          // z odstępem między blokami ich marginesy już się nie łączą (mb + mt zamiast max)
          items.push({ kind: "gap", block, line: 0, height: Math.max(0, Math.round((move - Math.min(mb, mt)) * 10) / 10), shift: move, bandTop: Math.round(sheetTop + res.pageH), label,
            head: head && { rows: head.rows, table: head.table, top: Math.round((sheetTop + res.pageH + GAP + res.topZone) * 10) / 10 } });
          shift += move;
          sheetTop += res.pageH + GAP;
        } else {
          items.push({ kind: "seam", top: Math.round(cut.y + shift), label });
          sheetTop = cut.top + shift - res.topZone;
        }
      }
      const guideTops = [0];
      items.forEach((x) => guideTops.push(x.kind === "gap" ? x.bandTop + GAP : x.top));
      // Nie obrysowujemy całego pola tekstu: na czterech rogach kartki dajemy tylko
      // prowadnice biegnące OD marginesu do krawędzi arkusza (jak znaki cięcia w DTP).
      const guides = guideTops.map((top) => ({ top: Math.round(top * 10) / 10, height: res.pageH, insetTop: res.padT, insetBottom: res.padB, insetLeft: res.padL, insetRight: res.padR }));
      const minH = gaps && sheetTop > 0 ? Math.ceil(sheetTop + res.pageH) : 0;
      // Przypisy dolne wielostronicowej sekcji: na dole swoich stron (nad stopką i dolnym marginesem),
      // jak w Wordzie i w podglądzie wydruku (AUDYT-MOBILE #7). Żywe <li> pozostają edytowalne.
      const gapBands = items.filter((x) => x.kind === "gap");
      const pageNotePlan = [];
      if (gaps && res.footOl && res.pageNotes && gapBands.length) {
        const lastBottom = minH || Math.ceil(sheetTop + res.pageH);
        res.pageNotes.forEach((ids, j) => {
          if (!ids?.length) return;
          const pageBottom = j < gapBands.length ? gapBands[j].bandTop : lastBottom;
          pageNotePlan.push({ page: j, ids, left: res.padL, right: res.padR, bottomY: pageBottom - res.bottomZone });
        });
      }
      want.push({ sec: res.sec, items, guides, minH, footOl: res.footOl, pageNotePlan });
    });
    // bez zmian (np. przeliczenie po przewinięciu, zoomie, skoku) — nic nie ruszamy
    const same = want.every(({ sec, items, guides, minH, pageNotePlan }) => {
      const gapEls = [...sec.querySelectorAll(GAP_SEL)];
      const bands = [...sec.querySelectorAll(":scope > .dwb-page-gap-band")];
      const seams = [...sec.querySelectorAll(":scope > .dwb-page-break")];
      const guideEls = [...sec.querySelectorAll(":scope > .dwb-page-margin-guide")];
      const noteEls = [...sec.querySelectorAll(":scope > ol.dwb-page-notes")];
      const wg = items.filter((x) => x.kind === "gap");
      const ws = items.filter((x) => x.kind === "seam");
      return gapEls.length === wg.length && bands.length === wg.length && seams.length === ws.length
        && wg.every((x, i) => gapEls[i].nextElementSibling === x.block && (gapEls[i].dataset.line || "0") === String(x.line)
          && (x.line ? Math.abs((parseFloat(gapEls[i].dataset.at) || 0) - x.at) < 0.6 && Math.abs((parseFloat(gapEls[i].dataset.shift) || 0) - x.shift) < 0.6 : Math.abs(parseFloat(gapEls[i].style.height) - x.height) < 0.6)
          && bands[i].dataset.label === x.label && parseFloat(bands[i].style.top) === x.bandTop)
        && ws.every((x, i) => parseFloat(seams[i].style.top) === x.top && seams[i].dataset.label === x.label)
        && guideEls.length === guides.length
        && guides.every((x, i) => parseFloat(guideEls[i].style.top) === x.top
          && parseFloat(guideEls[i].style.height) === x.height
          && parseFloat(guideEls[i].style.getPropertyValue("--dwb-guide-top")) === x.insetTop
          && parseFloat(guideEls[i].style.getPropertyValue("--dwb-guide-bottom")) === x.insetBottom
          && parseFloat(guideEls[i].style.getPropertyValue("--dwb-guide-left")) === x.insetLeft
          && parseFloat(guideEls[i].style.getPropertyValue("--dwb-guide-right")) === x.insetRight)
        && noteEls.length === pageNotePlan.length
        && pageNotePlan.every((pnp, i) => noteEls[i].dataset.page === String(pnp.page)
          && [...noteEls[i].querySelectorAll(":scope > li")].map((li) => li.dataset.dwbNote).join(",") === pnp.ids.join(",")
          && Math.abs(parseFloat(noteEls[i].style.top) - Math.round((pnp.bottomY - noteEls[i].offsetHeight) * 10) / 10) < 1)
        && (sec.style.getPropertyValue("--dwb-sec-min") || "") === (minH ? `${minH}px` : "");
    }) && h.querySelectorAll(".dwb-page-gap, .dwb-page-gap-spacer, .dwb-page-gap-band, .dwb-page-break, .dwb-page-margin-guide, .dwb-page-repeat:not(.dwb-page-hf)").length === want.reduce((n, w) => n + w.items.reduce((m, x) => m + (x.kind === "gap" ? (x.line ? 3 : 2) + (x.head ? 1 : 0) : 1), 0) + w.guides.length, 0);
    if (same) { pageHeadersFooters(h); contentMo?.takeRecords(); return; }
    // zmiana: akapit widoczny u góry zostaje na swoim miejscu na ekranie (jak zakotwiczenie przewijania);
    // tuż po przerysowaniu — akapit sprzed przerysowania (document.js, reloadFromBytes)
    const pending = typeof takeRenderScrollAnchor === "function" ? takeRenderScrollAnchor() : null;
    const anchor = pending ? null : viewportAnchor(h);
    const noteCaret = captureNoteCaret(h);
    clear(h);
    const bandBg = canvasBackground();
    const hint = (el) => {
      el.dataset.hint = "";
      el.dataset.hintPl = I18N.pl[gaps ? "pageGapHint" : "pageSeamHint"];
      el.dataset.hintEn = I18N.en[gaps ? "pageGapHint" : "pageSeamHint"];
      el.dataset.hintDelay = "0.8";
    };
    const splits = []; // odstępy w środku akapitów — do dociągnięcia po wstawieniu wszystkiego
    want.forEach(({ sec, items, guides, minH, footOl, pageNotePlan }) => {
      guides.forEach((g) => {
        const guide = document.createElement("div");
        guide.className = "dwb-page-margin-guide";
        guide.setAttribute("aria-hidden", "true");
        if (editable) guide.contentEditable = "false";
        guide.style.top = `${g.top}px`;
        guide.style.left = "0px";
        guide.style.right = "0px";
        guide.style.height = `${g.height}px`;
        guide.style.setProperty("--dwb-guide-top", `${g.insetTop}px`);
        guide.style.setProperty("--dwb-guide-bottom", `${g.insetBottom}px`);
        guide.style.setProperty("--dwb-guide-left", `${g.insetLeft}px`);
        guide.style.setProperty("--dwb-guide-right", `${g.insetRight}px`);
        sec.appendChild(guide);
      });
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
          } else if (x.block.tagName === "TR") {
            // przerwa między wierszami tabeli: element w roli wiersza (display: table-row), nie <tr> —
            // nie wchodzi do table.rows ani numeracji wierszy (narzędzia tabeli, plik)
            gap.classList.add("dwb-gap-row");
            const cell = document.createElement("div");
            cell.className = "dwb-gap-cell";
            gap.appendChild(cell);
            x.block.before(gap);
            if (x.head) sec.appendChild(repeatedHeadOverlay(sec, x.head));
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
      if (footOl && pageNotePlan.length) {
        const byId = new Map([...footOl.querySelectorAll(":scope > li[data-dwb-note]")].map((li) => [li.dataset.dwbNote, li]));
        footOl.classList.add("dwb-notes-paged");
        pageNotePlan.forEach((pnp) => {
          const ol = document.createElement("ol");
          ol.className = `${footOl.className.replace(/\bdwb-notes-paged\b/g, "").trim()} dwb-page-notes`;
          ol.dataset.page = String(pnp.page);
          if (editable) ol.contentEditable = "false";
          pnp.ids.forEach((id) => { const li = byId.get(id); if (li) ol.appendChild(li); });
          if (!ol.children.length) return;
          ol.style.left = `${pnp.left}px`;
          ol.style.right = `${pnp.right}px`;
          // przed macierzystym <ol>: zachowuje kolejność DOM (przypisy dolne przed końcowymi)
          footOl.before(ol);
          ol.style.top = `${Math.round((pnp.bottomY - ol.offsetHeight) * 10) / 10}px`;
        });
      }
    });
    splits.forEach(fitSplit);
    relayoutTextboxes(h); // pola tekstowe za swoimi akapitami / na swoich kartkach
    pageHeadersFooters(h);
    restoreNoteCaret(noteCaret);
    contentMo?.takeRecords();
    if (typeof syncPageScaleHeightNow === "function") syncPageScaleHeightNow(); // inaczej przewinięcie niżej obcina stary, krótszy obszar
    if (pending) restoreDocScrollAnchor(pending);
    else if (anchor && docViewportEl) {
      const d = anchor.el.getBoundingClientRect().top - anchor.top;
      if (Math.abs(d) > 0.5) docViewportEl.scrollTop += d;
    }
  }

  // Powtórzony wiersz nagłówka tabeli u góry strony w Edycji (Word: „Powtórz jako wiersz nagłówka”).
  // Kopia tylko do oglądania w zamkniętym shadow DOM: wyszukiwanie, liczniki wyrazów, pisownia,
  // edycja i mapowanie akapitów na plik jej nie widzą. Wygląd przepisany z obliczonych stylów
  // oryginału (style dokumentu nie sięgają do shadow DOM).
  function styledClone(el, deep = true) {
    const c = el.cloneNode(false);
    if (c.nodeType !== 1) return c;
    [...c.attributes].forEach((a) => { if (!/^(colspan|rowspan)$/i.test(a.name)) c.removeAttribute(a.name); });
    const cs = getComputedStyle(el);
    let css = "";
    for (let i = 0; i < cs.length; i++) { const k = cs[i]; if (!k.startsWith("--")) css += `${k}:${cs.getPropertyValue(k)};`; }
    c.style.cssText = css;
    if (deep) el.childNodes.forEach((n) => { if (n.nodeType === 1 || n.nodeType === 3) c.appendChild(styledClone(n)); });
    return c;
  }
  // Nagłówek i stopka na KAŻDEJ stronie Edycji, jak na wydruku i w Wordzie (dawniej nagłówek tylko
  // nad stroną 1, stopka tylko pod ostatnią — testerzy AI 2026-11-27, AUDYT-MOBILE #2). Oryginały
  // zostają, gdzie są (nagłówek strony 1, stopka ostatniej strony); pozostałe strony dostają kopie
  // tylko do oglądania (shadow DOM, jak powtarzany wiersz nagłówka tabeli) z numerem swojej strony.
  // Wersja na stronie jak w print-preview.js (hfFor): pierwsza strona sekcji — narysowana, dalsze —
  // parzysta (w:evenAndOddHeaders) albo domyślna; bez domyślnej po „first” — nic.
  let evenOdd = { bytes: null, on: false };
  function evenOddOn() {
    if (evenOdd.bytes !== originalFileBytes) {
      const bytes = originalFileBytes;
      evenOdd = { bytes, on: false };
      if (bytes && typeof loadDocxZipCached === "function") {
        loadDocxZipCached(bytes).then((z) => z.file("word/settings.xml")?.async("string")).then((x) => {
          if (evenOdd.bytes === bytes && /<w:evenAndOddHeaders(?![^>]*w:val="(0|false|off)")/.test(x || "")) { evenOdd.on = true; schedule(0); }
        }).catch(() => { /* bez ustawień — bez stron parzystych */ });
      }
    }
    return evenOdd.on;
  }
  function hfVersion(sec, tag, j, pageNo) {
    const shown = sec.querySelector(`:scope > ${tag}:not(.dwb-hf-alt)`);
    if (!shown) return null;
    const alt = (k) => sec.querySelector(`:scope > ${tag}.dwb-hf-alt[data-dwb-hf="${k}"]`);
    if (j === 0) return shown;
    if (pageNo % 2 === 0 && evenOddOn()) { const e = shown.dataset.dwbHf === "even" ? shown : alt("even"); if (e) return e; }
    if ((shown.dataset.dwbHf || "default") === "default") return shown;
    return alt("default") || (shown.dataset.dwbHf === "first" ? null : shown);
  }
  // Kopia części z numerami tej strony: numery wpisane na chwilę w oryginał (te same reguły co
  // podgląd wydruku: composeUi.stampPageNumbers), wygląd przepisany, tekst oryginału przywrócony.
  function hfCopy(part, kind, page, top, box, pageNo, total, secPages) {
    const hidden = part.classList.contains("dwb-hf-alt");
    const shown = hidden ? part.parentElement.querySelector(`:scope > ${part.localName}:not(.dwb-hf-alt)`) : null;
    if (hidden) { shown.style.display = "none"; part.style.display = "block"; }
    const leaves = [...part.querySelectorAll("span")].filter((s) => !s.firstElementChild).map((s) => [s, s.textContent]);
    try {
      if (typeof composeUi !== "undefined" && composeUi.stampPageNumbers) composeUi.stampPageNumbers(part, pageNo, total, secPages);
      const wrap = document.createElement("div");
      wrap.className = "dwb-page-hf dwb-page-repeat";
      wrap.dataset.kind = kind;
      wrap.dataset.page = String(pageNo);
      wrap.setAttribute("aria-hidden", "true");
      wrap.contentEditable = "false";
      const b = box || part.getBoundingClientRect();
      const sr = part.parentElement.getBoundingClientRect();
      const scale = sr.width / part.parentElement.offsetWidth || 1;
      wrap.style.top = `${Math.round(top * 10) / 10}px`;
      wrap.style.left = `${Math.round((b.left - sr.left) / scale * 10) / 10}px`;
      wrap.style.width = `${Math.round(b.width / scale * 10) / 10}px`;
      const root = wrap.attachShadow({ mode: window.__dwbOpenShadow ? "open" : "closed" });
      const c = styledClone(part);
      c.style.margin = "0";
      c.style.position = "static";
      root.appendChild(c);
      return wrap;
    } finally {
      leaves.forEach(([s, txt]) => { if (s.textContent !== txt) s.textContent = txt; });
      if (hidden) { part.style.display = ""; shown.style.display = ""; }
    }
  }
  function pageHeadersFooters(h) {
    const secs = [...h.querySelectorAll(".docx-wrapper > section.docx")];
    const counts = secs.map((sec) => sec.querySelectorAll(":scope > .dwb-page-gap-band").length + 1);
    const total = counts.reduce((a, b) => a + b, 0);
    let first = 1;
    secs.forEach((sec, si) => {
      const n = counts[si], start = first;
      first += n;
      const bands = [...sec.querySelectorAll(":scope > .dwb-page-gap-band")].map((b) => parseFloat(b.style.top) || 0);
      const sig = `${total}|${start}|${bands.join(",")}|${evenOdd.on}|${sec.offsetHeight}`;
      if (sec.dataset.dwbHfSig === sig) return;
      sec.querySelectorAll(":scope > .dwb-page-hf").forEach((el) => el.remove());
      sec.dataset.dwbHfSig = sig;
      if (!bands.length) return;
      const sr = sec.getBoundingClientRect();
      const scale = sr.height / sec.offsetHeight || 1;
      const add = [];
      // nagłówki stron 2…n: tyle samo pod górą strony, co oryginał pod górą strony 1
      for (let j = 1; j < n; j++) {
        const v = hfVersion(sec, "header", j, start + j);
        if (!v) continue;
        const show = v.classList.contains("dwb-hf-alt");
        if (show) { v.style.display = "block"; v.parentElement.querySelector(":scope > header:not(.dwb-hf-alt)").style.display = "none"; }
        const b = v.getBoundingClientRect();
        if (show) { v.style.display = ""; v.parentElement.querySelector(":scope > header:not(.dwb-hf-alt)").style.display = ""; }
        add.push(hfCopy(v, "header", j, bands[j - 1] + GAP + (b.top - sr.top) / scale, b, start + j, total, n));
      }
      // stopki stron 1…n−1: tyle samo nad dołem strony, co oryginał nad dołem ostatniej strony
      // (stopka rośnie od dołu — jej dół zostaje na miejscu dołu narysowanej, jak w podglądzie wydruku)
      const shownF = sec.querySelector(":scope > footer:not(.dwb-hf-alt)");
      const fBottom = shownF ? (sr.bottom - shownF.getBoundingClientRect().bottom) / scale : 0;
      for (let j = 0; j < n - 1; j++) {
        const v = hfVersion(sec, "footer", j, start + j);
        if (!v) continue;
        const show = v.classList.contains("dwb-hf-alt");
        if (show) { v.style.display = "block"; shownF.style.display = "none"; }
        const b = v.getBoundingClientRect();
        if (show) { v.style.display = ""; shownF.style.display = ""; }
        add.push(hfCopy(v, "footer", j, bands[j] - fBottom - b.height / scale, b, start + j, total, n));
      }
      add.forEach((el) => sec.appendChild(el));
    });
    // numery w oryginałach (nagłówek strony 1, stopka ostatniej) — liczba stron z odstępów
    if (typeof composeUi !== "undefined" && composeUi.fixPreviewPageNumbers) composeUi.fixPreviewPageNumbers();
  }
  function repeatedHeadOverlay(sec, head) {
    const r = sec.getBoundingClientRect();
    const scale = r.width / sec.offsetWidth || 1;
    const tb = head.table.getBoundingClientRect();
    const wrap = document.createElement("div");
    wrap.className = "dwb-page-repeat";
    wrap.setAttribute("aria-hidden", "true");
    wrap.contentEditable = "false";
    wrap.style.top = `${head.top}px`;
    wrap.style.left = `${Math.round((tb.left - r.left) / scale * 10) / 10}px`;
    wrap.style.width = `${Math.round(tb.width / scale * 10) / 10}px`;
    const root = wrap.attachShadow({ mode: window.__dwbOpenShadow ? "open" : "closed" }); // testy: __dwbOpenShadow
    const t = styledClone(head.table, false);
    t.style.margin = "0";
    t.style.width = "100%";
    t.style.height = "auto"; // obliczona wysokość oryginału = cała tabela, kopia to tylko nagłówek
    t.style.minHeight = "0";
    const cg = head.table.querySelector(":scope > colgroup");
    if (cg) t.appendChild(cg.cloneNode(true));
    const body = document.createElement("tbody");
    head.rows.forEach((row) => body.appendChild(styledClone(row)));
    t.appendChild(body);
    root.appendChild(t);
    return wrap;
  }

  function relayoutTextboxes(h) {
    if (typeof dwbTextbox !== "undefined") dwbTextbox.relayout?.(h);
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
    // rzeczywiste przesunięcie linijki granicy (po dopasowaniu) — do pomiaru przy następnym liczeniu
    const now = textLines(x.el)[x.line];
    if (now) gap.dataset.disp = String(Math.round((local(now[0]) - x.at) * 1000) / 1000);
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
    const ours = (n) => n.nodeType === 1 && n.matches?.(".dwb-page-gap, .dwb-page-gap-spacer, .dwb-page-gap-band, .dwb-page-break, .dwb-page-margin-guide, .dwb-page-repeat, .dwb-page-notes");
    // Raz na zawsze (zgłoszenie 2026-10-06: „przy różnych manewrach coś wjeżdża, jakby nie było
    // układu strony”): obserwujemy rozmiar KAŻDEGO bloku treści kartki (akapit, tabela, obraz).
    // Cokolwiek zmieni wysokość bloku — suwak obrazu, formatowanie, pole formularza, krój, który
    // doszedł później, zmiana wyrównania — przelicza granice, bez wyliczania, które funkcje to robią.
    // Własne odstępy nie są obserwowane; „nic się nie zmieniło” kończy liczenie bez dotykania DOM.
    const sizes = new WeakMap(); // blok → ostatnia wysokość (pierwszy pomiar to nie zmiana)
    const watched = new Set(); // obserwowane bloki — odpięte od DOM zwalniamy (unobserve) w watchBlocks
    const blockRO = new ResizeObserver((entries) => {
      let changed = false;
      for (const e of entries) {
        const h = Math.round(e.contentRect.height * 2) / 2;
        if (sizes.has(e.target) && sizes.get(e.target) !== h) changed = true;
        sizes.set(e.target, h);
      }
      if (changed) schedule(150);
    });
    const watchBlocks = () => {
      // ResizeObserver trzyma obserwowane elementy mocno: bez unobserve każde przerysowanie
      // (Cofnij, wklejenie obiektu, usunięcie przez sekcje) zostawiało w pamięci całe stare drzewo
      watched.forEach((b) => { if (!b.isConnected) { blockRO.unobserve(b); watched.delete(b); } });
      docCanvasEl.querySelectorAll(".docx-preview-host section.docx > article > *").forEach((b) => {
        if (!watched.has(b) && !ours(b)) { watched.add(b); blockRO.observe(b); }
      });
    };
    contentMo = new MutationObserver((records) => {
      let content = false;
      for (const r of records) {
        if (r.type === "childList") { if ([...r.addedNodes, ...r.removedNodes].some((n) => !ours(n))) content = true; }
        else if (!ours(r.target) && !r.target.closest?.(".dwb-page-gap-band, .dwb-page-break, .dwb-page-margin-guide")) content = true; // styl obrazu, tekst
        if (content) break;
      }
      if (!content) return;
      watchBlocks();
      schedule(150);
    });
    contentMo.observe(docCanvasEl, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ["style", "width", "height", "src"] });
    watchBlocks();
  }
  document.fonts?.ready?.then(() => schedule(0));

  // prepare: ustawienia stylów pliku (razem z następnym, wdowy) wczytane z góry — document.js woła je
  // w trakcie renderu w tle, żeby compute() po podmianie nie czekało na plik (bez klatki pośredniej)
  return { schedule, compute, enabled, paginate, gapsOn, setGaps, prepare: ensureKeepInfo };
})();
