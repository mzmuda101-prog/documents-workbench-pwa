// print-preview.js — Podgląd wydruku (Ctrl/⌘+P, Eksport → Drukuj, menu ⋯), jak w Wordzie.
//
// Prawdziwe kartki: rozmiar papieru i marginesy z pliku, nagłówek i stopka z numerem strony na
// KAŻDEJ stronie, odstęp między kartkami. Granice stron liczone tymi samymi zasadami Worda co
// znaczniki „str. N” w Edycji (page-breaks.js: wdowy/sieroty, razem z następnym, wiersz tabeli
// w całości). „Drukuj” drukuje DOKŁADNIE te kartki: @page bez marginesów przeglądarki (bez jej
// nagłówków z datą i adresem), każda kartka = jedna strona — wydruk = podgląd.
//
// Dawniej druk szedł z jednej wysokiej „kartki” podglądu z @page { margin: 0 }: przeglądarka
// cięła ją, gdzie popadnie — na 2. i dalszych stronach tekst stał przy samej krawędzi papieru
// (bez górnego/dolnego marginesu), nagłówek i stopka były tylko raz.
//
// Fizyczne ograniczenia drukarki: większość nie drukuje przy samej krawędzi (tańsze ~5–6 mm).
// Treść bliżej niż 6,35 mm od krawędzi = ostrzeżenie (jak Word: „marginesy poza obszarem
// wydruku”) i widoczna strefa na kartkach.
//
// Kartki budowane z osobnego, niewidocznego renderu bieżącego stanu (buildDocumentForSave) —
// zawsze w układzie stron, także w Widoku mobilnym; edytowany podgląd zostaje nietknięty.
// Bloki (akapity, tabele) są PRZENOSZONE na kolejne kartki w kolejności dokumentu — liczniki CSS
// numeracji list idą dalej przez strony; ciąg dalszy bloku na następnej stronie to kopia bez
// liczników (pp-dup), przesunięta i przycięta do okna strony.

const dwbPrint = (() => {
  const UNPRINTABLE_MM = 6.35;
  const PX_PER_MM = 96 / 25.4;
  let overlay = null;
  let building = false;
  let state = null; // { zoom, fit, sheets, risky, pageStyle }

  const frame = () => new Promise((r) => requestAnimationFrame(() => r()));

  function rel(el, base) {
    const r = el.getBoundingClientRect();
    return { top: r.top - base.top, bottom: r.bottom - base.top, left: r.left - base.left, right: r.right - base.left, width: r.width, height: r.height };
  }

  function pageNumberSelector() {
    try { return typeof composeUi !== "undefined" && composeUi.pageNumberSelector ? composeUi.pageNumberSelector() : ""; } catch (_) { return ""; }
  }

  function stampPageNumbers(part, page, total, sel) {
    if (!part || !sel) return;
    part.querySelectorAll("p").forEach((p) => {
      const spans = Array.from(p.querySelectorAll(sel));
      if (spans[0]) spans[0].textContent = String(page);
      if (spans[1]) spans[1].textContent = String(total);
    });
  }

  function placeAbs(el, box) {
    el.style.position = "absolute";
    el.style.top = `${box.top}px`;
    el.style.left = `${box.left}px`;
    el.style.width = `${box.width}px`;
    el.style.margin = "0";
  }

  // Kartki ze zrenderowanego, niewidocznego dokumentu (host) i podziału na strony (sections).
  function buildSheets(host, sections) {
    const zone = UNPRINTABLE_MM * PX_PER_MM;
    const plan = [];
    // 1) pomiary — przed przenoszeniem czegokolwiek
    sections.forEach((S) => {
      const sec = S.sec;
      const base = sec.getBoundingClientRect();
      const secH = sec.offsetHeight;
      const header = sec.querySelector(":scope > header");
      const footer = sec.querySelector(":scope > footer");
      const blocks = [];
      // wszystkie <article> kartki (sekcje ciągłe, np. fragment w 2 kolumnach) — bloki i tak stoją
      // bezwzględnie w zmierzonym miejscu, więc kolumny zostają, gdzie były
      sec.querySelectorAll(":scope > article").forEach((article) => [...article.children].forEach((el) => blocks.push({ el, holder: "article", box: rel(el, base) })));
      // przypisy dolne na dole swoich stron (page-breaks.js: S.pageNotes) — nie płyną z treścią
      const footOl = S.pageNotes ? sec.querySelector(":scope > ol.dwb-notes-footnote") : null;
      sec.querySelectorAll(":scope > ol").forEach((ol, oi) => { if (ol !== footOl) [...ol.children].forEach((el) => blocks.push({ el, holder: oi, box: rel(el, base) })); });
      const cs = getComputedStyle(sec);
      const starts = [S.contentTop, ...S.cuts.map((c) => c.top)];
      // powtarzane wiersze nagłówka tabeli (page-breaks.js: cut.header) — miejsce tabeli zmierzone teraz
      const heads = [null, ...S.cuts.map((c) => c.header && { ...c.header, box: rel(c.header.table, base) })];
      plan.push({
        S, sec, secH, blocks, heads, footOl,
        header: header && { el: header, box: rel(header, base) },
        footer: footer && { el: footer, box: rel(footer, base), fromBottom: secH - rel(footer, base).top },
        ols: [...sec.querySelectorAll(":scope > ol")],
        starts,
        pad: { l: parseFloat(cs.paddingLeft) || 0, r: parseFloat(cs.paddingRight) || 0 },
      });
    });
    const total = plan.reduce((n, P) => n + P.starts.length, 0);
    const numSel = pageNumberSelector();
    // 2) kartki
    const sheets = [];
    const risky = [];
    let pageNo = 0;
    plan.forEach((P) => {
      const { S, sec } = P;
      const moved = new Set();
      P.starts.forEach((ws, j) => {
        pageNo += 1;
        const we = j + 1 < P.starts.length ? P.starts[j + 1] : Infinity;
        const sheet = sec.cloneNode(false);
        sheet.classList.add("pp-sheet");
        sheet.removeAttribute("contenteditable");
        sheet.dataset.page = String(pageNo);
        sheet.style.height = `${S.pageH}px`;
        sheet.style.minHeight = `${S.pageH}px`;
        sheet.style.width = `${S.pageW}px`;
        // okno strony: treść od ws do we (pierwsza linijka następnej strony), na miejscu treści kartki
        const top = S.contentTop;
        const head = P.heads[j]; // strona zaczyna się w środku tabeli z nagłówkiem: najpierw jego kopia
        const headH = head ? head.h : 0;
        const clip = document.createElement("div");
        clip.className = "pp-clip";
        // z nagłówkiem okno strony zaczyna się pod nim (tabela to jeden blok — wiersze sprzed
        // granicy prześwitywałyby spod kopii nagłówka)
        clip.style.top = `${top + headH}px`;
        clip.style.height = `${Math.max(0, (we === Infinity ? S.pageH - top - headH : we - ws))}px`;
        const shift = document.createElement("div");
        shift.className = "pp-shift";
        shift.style.top = `${-ws}px`;
        shift.style.width = `${S.pageW}px`;
        const art = sec.querySelector(":scope > article")?.cloneNode(false) || document.createElement("article");
        art.removeAttribute("style");
        shift.appendChild(art);
        const olClones = P.ols.map((ol) => { const c = ol.cloneNode(false); shift.appendChild(c); return c; });
        let near = false;
        P.blocks.forEach((b) => {
          if (b.box.bottom <= ws || b.box.top >= we || !b.box.height) return;
          let el;
          if (!moved.has(b.el)) { el = b.el; moved.add(b.el); } else { el = b.el.cloneNode(true); el.classList.add("pp-dup"); }
          placeAbs(el, b.box);
          (b.holder === "article" ? art : olClones[b.holder]).appendChild(el);
          if (b.box.left < zone || b.box.right > S.pageW - zone) near = true;
        });
        clip.appendChild(shift);
        if (head) {
          // kopia tabeli z samymi wierszami nagłówka (kolumny i style z oryginału)
          const t = head.table.cloneNode(false);
          t.removeAttribute("id");
          t.classList.add("pp-dup", "pp-repeat-head");
          placeAbs(t, { top, left: head.box.left, width: head.box.width });
          const cols = head.table.querySelector(":scope > colgroup");
          if (cols) t.appendChild(cols.cloneNode(true));
          const body = document.createElement("tbody");
          head.rows.forEach((r) => body.appendChild(r.cloneNode(true)));
          t.appendChild(body);
          sheet.appendChild(t);
        }
        sheet.appendChild(clip);
        const notes = P.footOl && S.pageNotes?.[j];
        if (notes?.length) {
          // przypisy tej strony przy dolnym marginesie (nad stopką, jak w Wordzie); numer z oryginału
          const ol = P.footOl.cloneNode(false);
          ol.classList.add("pp-page-notes");
          const all = [...P.footOl.children];
          notes.forEach((id) => {
            const li = all.find((x) => x.dataset.dwbNote === id);
            if (!li) return;
            const c = li.cloneNode(true);
            if (li.classList.contains("dwb-note-native")) c.value = all.indexOf(li) + 1;
            ol.appendChild(c);
          });
          ol.style.position = "absolute";
          ol.style.left = `${P.pad.l}px`;
          ol.style.right = `${P.pad.r}px`;
          ol.style.top = "auto";
          ol.style.bottom = `${S.bottomZone ?? S.padB}px`;
          ol.style.margin = "0";
          sheet.appendChild(ol);
        }
        if (P.header) {
          const h = P.header.el.cloneNode(true);
          h.classList.add("pp-hf");
          placeAbs(h, P.header.box);
          stampPageNumbers(h, pageNo, total, numSel);
          sheet.appendChild(h);
          if (P.header.box.top < zone && h.textContent.trim()) near = true;
        }
        if (P.footer) {
          const f = P.footer.el.cloneNode(true);
          f.classList.add("pp-hf");
          placeAbs(f, { ...P.footer.box, top: S.pageH - P.footer.fromBottom });
          stampPageNumbers(f, pageNo, total, numSel);
          sheet.appendChild(f);
          if (S.pageH - P.footer.fromBottom + P.footer.box.height > S.pageH - zone && f.textContent.trim()) near = true;
        }
        if (S.padT < zone || S.padB < zone || P.pad.l < zone || P.pad.r < zone) near = true;
        const z = document.createElement("div");
        z.className = "pp-unprintable";
        z.style.borderWidth = `${zone}px`;
        z.style.setProperty("--pp-zone", `${zone}px`);
        z.setAttribute("aria-hidden", "true");
        sheet.appendChild(z);
        const lab = document.createElement("div");
        lab.className = "pp-page-label";
        lab.setAttribute("aria-hidden", "true");
        lab.textContent = t("ppPageOf", { n: pageNo, total });
        sheet.appendChild(lab);
        if (near) risky.push(pageNo);
        sheets.push({ sheet, w: S.pageW, h: S.pageH });
      });
    });
    return { sheets, risky, total };
  }

  function rangesText(nums) {
    const out = [];
    for (let i = 0; i < nums.length; i++) {
      let j = i;
      while (j + 1 < nums.length && nums[j + 1] === nums[j] + 1) j++;
      out.push(i === j ? String(nums[i]) : `${nums[i]}–${nums[j]}`);
      i = j;
    }
    return out.join(", ");
  }

  // @page: rozmiar kartek, bez marginesów przeglądarki (marginesy są w kartkach); kartki innego
  // rozmiaru (sekcja pozioma) — strony nazwane.
  function pageCss(sheets) {
    const sizes = [];
    const key = (s) => `${Math.round(s.w)}x${Math.round(s.h)}`;
    sheets.forEach((s) => { if (!sizes.includes(key(s))) sizes.push(key(s)); });
    const mm = (px) => `${(px / PX_PER_MM).toFixed(1)}mm`;
    let css = "";
    sizes.forEach((k, i) => {
      const [w, h] = k.split("x").map(Number);
      css += `${i === 0 ? "@page" : `@page pp${i}`} { size: ${mm(w)} ${mm(h)}; margin: 0; }\n`;
    });
    sheets.forEach((s) => { const i = sizes.indexOf(key(s)); if (i > 0) s.sheet.style.page = `pp${i}`; });
    return css;
  }

  function fitZoom() {
    const scroller = overlay?.querySelector(".pp-scroll");
    if (!scroller || !state) return 1;
    const maxW = Math.max(...state.sheets.map((s) => s.w));
    return Math.max(0.2, Math.min(1, (scroller.clientWidth - 32) / maxW));
  }

  function applyZoom() {
    if (!overlay || !state) return;
    if (state.fit) state.zoom = fitZoom();
    overlay.querySelector(".pp-pages").style.setProperty("--pp-zoom", String(state.zoom));
    overlay.querySelector(".pp-zoom-val").textContent = `${Math.round(state.zoom * 100)}%`;
  }

  function close() {
    if (!overlay) return;
    overlay.remove();
    overlay = null;
    state = null;
    document.getElementById("ppPageStyle")?.remove();
    document.body.classList.remove("pp-open");
    window.removeEventListener("resize", onResize);
    if (lastFocus?.isConnected) lastFocus.focus({ preventScroll: true });
  }
  function onResize() { if (state?.fit) applyZoom(); }

  let lastFocus = null;
  function showOverlay(styles, result) {
    close();
    lastFocus = document.activeElement;
    overlay = document.createElement("div");
    overlay.className = "pp-overlay";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.setAttribute("aria-label", t("ppTitle"));
    overlay.innerHTML = `
      <div class="pp-bar">
        <button type="button" class="tb-btn pp-close" data-act="close"></button>
        <div class="pp-title"><strong></strong><span class="pp-count"></span></div>
        <div class="pp-zoom" role="group">
          <button type="button" class="tb-btn" data-act="out">−</button>
          <button type="button" class="tb-btn pp-zoom-val" data-act="fit"></button>
          <button type="button" class="tb-btn" data-act="in">+</button>
        </div>
        <button type="button" class="tb-btn pp-margins" data-act="margins"></button>
        <label class="pp-zone-toggle"><input type="checkbox" data-act="zone"><span></span></label>
        <button type="button" class="btn primary pp-print" data-act="print"></button>
      </div>
      <div class="pp-warn" hidden></div>
      <div class="pp-scroll" tabindex="0"><div class="doc-canvas pp-canvas"><div class="docx-preview-host pp-host"><div class="docx-wrapper pp-pages"></div></div></div></div>`;
    overlay.querySelector(".pp-close").textContent = `← ${t("ppClose")}`;
    overlay.querySelector(".pp-close").setAttribute("aria-label", t("ppClose"));
    overlay.querySelector(".pp-title strong").textContent = t("ppTitle");
    overlay.querySelector(".pp-count").textContent = t("ppPages", { n: result.total });
    overlay.querySelector('[data-act="out"]').setAttribute("aria-label", t("zoomOut"));
    overlay.querySelector('[data-act="in"]').setAttribute("aria-label", t("zoomIn"));
    overlay.querySelector('[data-act="fit"]').setAttribute("aria-label", t("zoomFit"));
    overlay.querySelector(".pp-zone-toggle span").textContent = t("ppZone");
    overlay.querySelector(".pp-zone-toggle").dataset.hint = "";
    overlay.querySelector(".pp-zone-toggle").dataset.hintPl = I18N.pl.ppZoneHint.replace("{mm}", "6,35");
    overlay.querySelector(".pp-zone-toggle").dataset.hintEn = I18N.en.ppZoneHint.replace("{mm}", "6.35");
    overlay.querySelector(".pp-print").textContent = t("ppPrint");
    overlay.querySelector(".pp-margins").textContent = t("ppMargins");
    const host = overlay.querySelector(".pp-host");
    styles.forEach((s) => host.insertBefore(s, host.firstChild));
    const pages = overlay.querySelector(".pp-pages");
    result.sheets.forEach((s) => pages.appendChild(s.sheet));
    if (result.risky.length) {
      const w = overlay.querySelector(".pp-warn");
      w.hidden = false;
      w.textContent = t("ppWarn", { pages: rangesText(result.risky), mm: String(UNPRINTABLE_MM).replace(".", currentLang === "pl" ? "," : ".") });
      overlay.querySelector('[data-act="zone"]').checked = true;
      overlay.classList.add("pp-show-zone");
    }
    const style = document.createElement("style");
    style.id = "ppPageStyle";
    style.textContent = `@media print {\n${pageCss(result.sheets)}}`;
    document.head.appendChild(style);
    state = { zoom: 1, fit: true, sheets: result.sheets, risky: result.risky };
    overlay.addEventListener("click", (e) => {
      const act = e.target.closest("[data-act]")?.dataset.act;
      if (act === "close") close();
      else if (act === "print") print();
      else if (act === "margins" && typeof composeUi !== "undefined") {
        // marginesy jak w Wordzie przy drukowaniu — po zmianie kartki układają się od nowa
        composeUi.openPageSetup(e.target.closest("[data-act]"), () => { if (overlay) open(); });
      }
      else if (act === "fit") { state.fit = true; applyZoom(); }
      else if (act === "in" || act === "out") {
        state.fit = false;
        state.zoom = Math.max(0.25, Math.min(3, Math.round((state.zoom + (act === "in" ? 0.1 : -0.1)) * 10) / 10));
        applyZoom();
      }
    });
    overlay.querySelector('[data-act="zone"]').addEventListener("change", (e) => overlay.classList.toggle("pp-show-zone", e.target.checked));
    document.body.appendChild(overlay);
    document.body.classList.add("pp-open");
    window.addEventListener("resize", onResize);
    applyZoom();
    overlay.querySelector(".pp-print").focus({ preventScroll: true });
  }

  function print() {
    if (!overlay) return;
    try { window.print(); } catch (_) { toast(t("exportPrintFail"), "error"); }
  }

  async function open() {
    if (!originalFileBytes) { toast(t("noFileToSave"), "error"); return; }
    if (building) return;
    building = true;
    setLoading(true, t("ppBuilding"));
    const src = document.createElement("div");
    src.className = "doc-canvas dwb-pp-src";
    src.setAttribute("aria-hidden", "true");
    try {
      const bytes = await buildDocumentForSave();
      document.body.appendChild(src);
      await renderDocxPreview(bytes, src, { pages: true });
      await waitDocFontsSettled(src); // pomiary dopiero na właściwych krojach
      await frame();
      await frame();
      const host = src.querySelector(".docx-preview-host");
      if (!host) throw new Error("render");
      // Przypisy dolne stoją na dole swoich stron (page-breaks.js: pageNotes) — ich lista na końcu
      // sekcji nie może spychać przypisów końcowych pod sobą (Word: końcowe zaraz za treścią).
      // Wyjęta z przepływu, z tą samą szerokością (wysokości przypisów dalej mierzalne).
      host.querySelectorAll("section.docx > ol.dwb-notes-footnote").forEach((ol) => {
        ol.style.width = `${ol.offsetWidth}px`;
        ol.style.position = "absolute";
      });
      // kartka docx-preview to kolumna flex z „margin-bottom: auto” na treści — przypisy końcowe
      // lądowały przy dole kartki; Word stawia je zaraz za tekstem (stopka zostaje na dole)
      host.querySelectorAll("section.docx > ol.dwb-notes-endnote").forEach((ol) => {
        const arts = ol.parentElement.querySelectorAll(":scope > article");
        if (arts.length) arts[arts.length - 1].style.marginBottom = "0";
        ol.style.marginBottom = "auto";
      });
      const sections = await dwbPageBreaks.paginate(host, { includeNotes: true });
      const result = buildSheets(host, sections);
      // Przeniesienie <style> buduje arkusz od nowa z TEKSTU — zmiany w CSSOM (zamienniki krojów z
      // rodzajem, krój pustych akapitów, kolejność stylów tabeli: docx-render-fixes.js) przepadały i
      // wydruk różnił się od podglądu („Zapotrzebowanie”: tabela w Calibri zamiast Arial).
      // Najpierw zapisujemy bieżące reguły z powrotem do tekstu.
      const styles = [...host.querySelectorAll(":scope > style")];
      styles.forEach((st) => {
        try { const rules = st.sheet?.cssRules; if (rules) st.textContent = [...rules].map((r) => r.cssText).join("\n"); } catch (_) { /* arkusz niedostępny — zostaje tekst */ }
      });
      showOverlay(styles, result);
    } catch (err) {
      log(`Podgląd wydruku: ${err.message || err}`, "error");
      toast(t("ppFail"), "error");
    } finally {
      src.remove();
      setLoading(false);
      building = false;
    }
  }

  // Esc zamyka podgląd — też gdy fokus wyszedł poza okno (np. po zamknięciu okna druku)
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape" || !overlay || document.querySelector(".compose-pop")) return; // najpierw okienko marginesów
    e.preventDefault();
    e.stopPropagation();
    close();
  }, true);

  // Ctrl/⌘+P = podgląd wydruku (jak Word: Plik → Drukuj); w podglądzie — od razu druk.
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || e.code !== "KeyP") return;
    if (!originalFileBytes) return;
    e.preventDefault();
    if (overlay) print(); else open();
  }, true);

  document.getElementById("printPreviewMenuItem")?.addEventListener("click", () => open());

  return { open, close, print, isOpen: () => !!overlay, _state: () => state };
})();
