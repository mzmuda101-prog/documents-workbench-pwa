// doc-notes.js — przypisy dolne i końcowe w podglądzie.
//
//   - numer jak w Wordzie: ciągły przez cały dokument (docx-preview liczył od 1 na każdej stronie
//     podglądu — łatka nr 5 w scripts/vendor-libs.mjs), format z ustawień pliku (1, i, a, *…),
//   - dymek z treścią przypisu po najechaniu na odnośnik (na dotyku: przytrzymanie palca),
//   - skok do przypisu: klik w odnośnik (Czytanie) / dwuklik albo Ctrl/⌘+klik (Edycja) + „↩ Wróć”;
//     klik w numer przypisu na dole strony wraca do odnośnika,
//   - akapit treści z odnośnikiem jest edytowalny — odnośnik to „wyspa” (cały fragment z pliku
//     wraca przy zapisie; dawniej taki akapit był tylko do odczytu),
//   - tekst przypisu edytowalny na dole strony (Enter = nowy akapit przypisu).
//
// Zmiany w przypisach idą TĄ SAMĄ listą co zmiany w treści (collectInlineParagraphEdits →
// op „paragraphBatch” → applyNoteEditsInZip), więc Zapisz, Cofnij, szkic odzyskiwania, karty
// i przerysowanie obejmują je bez osobnych ścieżek.

const NOTE_PART = { footnote: "word/footnotes.xml", endnote: "word/endnotes.xml" };

const dwbNotes = (() => {
  // klucz „footnote:2” → { li, editable, lock, base: [runs akapitu], mark: wzór numeru na początku }
  let notes = new Map();
  let labels = new Map(); // klucz → numer jak w Wordzie („3”, „iv”, „*”)
  let nums = new Map(); // klucz → { n: numer kolejny, fmt } (lista przypisów bez numeru w treści)
  let texts = new Map(); // klucz → treść przypisu (dymek)
  let parts = null; // { footnote, endnote, settings } — DOM części z pliku

  const wKids = (el, name) => Array.from(el?.childNodes || []).filter((n) => n.nodeType === 1 && n.namespaceURI === W_NS && (!name || n.localName === name));
  const noteParas = (li) => Array.from(li?.children || []).filter((c) => c.tagName === "P");
  const hostEl = () => docCanvasEl?.querySelector(".docx-preview-host");

  // ── numeracja ──────────────────────────────────────────────────────────────
  function roman(n) {
    const table = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
    let s = "";
    for (const [v, r] of table) while (n >= v) { s += r; n -= v; }
    return s;
  }
  function formatNum(n, fmt) {
    switch (fmt) {
      case "lowerRoman": return roman(n);
      case "upperRoman": return roman(n).toUpperCase();
      case "lowerLetter": return String.fromCharCode(97 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1); // a…z, aa…zz
      case "upperLetter": return String.fromCharCode(65 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
      case "chicago": return ["*", "†", "‡", "§"][(n - 1) % 4].repeat(Math.floor((n - 1) / 4) + 1);
      default: return String(n);
    }
  }
  // Word: przypisy dolne domyślnie 1, 2, 3…, końcowe i, ii, iii… (chyba że plik mówi inaczej)
  function numProps(kind) {
    const pr = parts?.settings?.getElementsByTagNameNS(W_NS, `${kind}Pr`)[0];
    const fmtEl = pr ? wKids(pr, "numFmt")[0] : null;
    const startEl = pr ? wKids(pr, "numStart")[0] : null;
    const fmt = fmtEl ? getWVal(fmtEl) : null;
    const start = startEl ? parseInt(getWVal(startEl), 10) : 1;
    return { fmt: fmt || (kind === "endnote" ? "lowerRoman" : "decimal"), start: start > 0 ? start : 1 };
  }

  async function readParts(bytes) {
    const zip = await loadDocxZipCached(bytes);
    const read = async (path) => {
      const f = zip.file(path);
      return f ? new DOMParser().parseFromString(await f.async("string"), "application/xml") : null;
    };
    return { footnote: await read(NOTE_PART.footnote), endnote: await read(NOTE_PART.endnote), settings: await read("word/settings.xml") };
  }

  function noteXml(kind, id) {
    const doc = parts?.[kind];
    if (!doc) return null;
    return Array.from(doc.getElementsByTagNameNS(W_NS, kind)).find((n) => n.getAttributeNS(W_NS, "id") === id) || null;
  }

  function plainOf(runsList) {
    return runsList.map((runs) => previewRunsToPlainText(runs)).join("\n").replace(/[ \t]+/g, " ").trim();
  }

  // ── 1) przed oznaczaniem akapitów: pliki, numery, treść dymków ─────────────
  async function prepare(bytes) {
    notes = new Map();
    labels = new Map();
    nums = new Map();
    texts = new Map();
    parts = null;
    const host = hostEl();
    const sups = host ? Array.from(host.querySelectorAll("sup[data-dwb-note]")) : [];
    if (!sups.length && !host?.querySelector("ol.dwb-notes")) return;
    parts = await readParts(bytes);
    const fmt = { footnote: numProps("footnote"), endnote: numProps("endnote") };
    sups.forEach((sup) => {
      const key = sup.dataset.dwbNote;
      const f = fmt[key.split(":")[0]] || fmt.footnote;
      const label = formatNum(f.start - 1 + (Number(sup.dataset.dwbNoteNum) || 1), f.fmt);
      sup.textContent = label;
      if (!labels.has(key)) { labels.set(key, label); nums.set(key, { n: f.start - 1 + (Number(sup.dataset.dwbNoteNum) || 1), fmt: f.fmt }); }
    });
    labels.forEach((_, key) => {
      const [kind, id] = key.split(":");
      const el = noteXml(kind, id);
      if (el) texts.set(key, plainOf(wKids(el, "p").map((p) => extractRunsFromParagraphXml(p))));
    });
  }

  // ── 2) akapit TREŚCI: odnośniki jako wyspy ─────────────────────────────────
  // null = gotowe (akapit edytowalny), inaczej powód blokady.
  function stampRefs(xp, el, nextKey) {
    const refs = Array.from(xp.getElementsByTagNameNS(W_NS, "footnoteReference")).concat(Array.from(xp.getElementsByTagNameNS(W_NS, "endnoteReference")));
    if (!refs.length) return null;
    const xmlParts = paragraphXmlParts(xp);
    const runs = xmlParts.filter((r) => noteRunKind(r) === "ref");
    // odnośnik w zwykłym fragmencie akapitu (nie w linku, polu, śledzonej zmianie), sam — bez tekstu obok
    const plain = runs.length === refs.length && runs.every((r) => r.parentNode === xp
      && wKids(r).every((n) => ["rPr", "footnoteReference", "endnoteReference"].includes(n.localName))
      && wKids(r).filter((n) => /noteReference$/.test(n.localName)).length === 1);
    const sups = Array.from(el.querySelectorAll("sup[data-dwb-note]"));
    if (!plain || sups.length !== runs.length) return "lockNote";
    for (let i = 0; i < runs.length; i++) {
      const ref = wKids(runs[i]).find((n) => /noteReference$/.test(n.localName));
      const kind = ref.localName === "footnoteReference" ? "footnote" : "endnote";
      if (sups[i].dataset.dwbNote !== `${kind}:${ref.getAttributeNS(W_NS, "id")}`) return "lockNote";
    }
    runs.forEach((r, i) => {
      const key = nextKey();
      docIslandXml.set(key, new XMLSerializer().serializeToString(r));
      // docx-preview: <span (fragment)>[<sup> indeks górny ze stylu]<sup>1</sup>…</span> — wyspą jest
      // całe opakowanie fragmentu (inaczej kursor stawał w zewnętrznym <sup> i pisanie szło w indeks)
      let island = sups[i];
      while (island.parentElement && island.parentElement !== el && ["SPAN", "SUP", "SUB"].includes(island.parentElement.tagName) && island.parentElement.childNodes.length === 1) island = island.parentElement;
      island.dataset.cm = key;
      island.dataset.cmKind = "note";
      island.contentEditable = "false";
    });
    // Odnośnik na KOŃCU akapitu („…z planu⁴”): Chrome nie postawi kursora kliknięciem za
    // nieedytowalnym elementem na końcu bloku (stawiał go na początku akapitu). Niewidoczny znak
    // za odnośnikiem daje miejsce na kursor; odczyt tekstu z podglądu go pomija (docx-run-styles.js).
    const last = sups[sups.length - 1].closest('[data-cm-kind="note"]');
    const after = document.createRange();
    after.setStartAfter(last);
    after.setEnd(el, el.childNodes.length);
    if (!after.toString().replace(/\uFEFF/g, "").length && !after.toString().includes("\uFEFF")) el.appendChild(document.createTextNode("\uFEFF"));
    return null;
  }

  // długości numerów odnośników w akapicie (stampCommentMarks liczy po nich położenia w tekście)
  function refLabelLengths(el) {
    return Array.from(el.querySelectorAll("sup[data-dwb-note]")).map((s) => s.textContent.length);
  }

  // ── 3) lista przypisów: numer, blokady, edycja ─────────────────────────────
  function noteParaLock(xp) {
    const hit = INLINE_LOCK_TAGS.find(([tag]) => xp.getElementsByTagNameNS(W_NS, tag).length);
    if (hit) return hit[1];
    if (xp.getElementsByTagNameNS(W_NS, "sdt").length) return "lockField";
    if (paragraphPageBreakLock(xp)) return "lockNoteText";
    if (Array.from(xp.getElementsByTagNameNS(W_NS, "footnoteReference")).length || Array.from(xp.getElementsByTagNameNS(W_NS, "endnoteReference")).length) return "lockNoteText";
    // numer przypisu z tekstem w tym samym fragmencie — wyspa zabrałaby tekst
    if (paragraphXmlParts(xp).some((r) => noteRunKind(r) === "mark" && wKids(r).some((n) => n.localName !== "rPr" && n.localName !== "footnoteRef" && n.localName !== "endnoteRef"))) return "lockNoteText";
    return paragraphNestedRunLock(xp) || paragraphCommentLock(xp);
  }

  // Numer na początku przypisu: w pliku fragment <w:footnoteRef/>, podgląd go nie rysuje (numer
  // dawała lista). Wstawiamy go jako wyspę z numerem — wygląda jak w Wordzie i zapis go oddaje.
  function stampMark(xp, el, nextKey, label) {
    let offset = 0;
    for (const r of paragraphXmlParts(xp)) {
      if (noteRunKind(r) === "mark") {
        const key = nextKey();
        docIslandXml.set(key, new XMLSerializer().serializeToString(r));
        const sup = document.createElement("sup");
        sup.className = "note-mark";
        sup.textContent = label;
        sup.dataset.cm = key;
        sup.dataset.cmKind = "note";
        sup.contentEditable = "false";
        const range = offset > 0 ? formDomRange(el, offset, offset) : null;
        if (range) range.insertNode(sup);
        else el.insertBefore(sup, el.firstChild);
        return sup;
      }
      if (r.localName === "r") wKids(r).forEach((n) => { if (n.localName === "t") offset += n.textContent.length; else if (n.localName === "br" || n.localName === "tab") offset += 1; });
    }
    return null;
  }

  function setupList(nextKey) {
    const host = hostEl();
    refCount = bodyRefCount();
    if (!host || !parts) return;
    host.querySelectorAll("ol.dwb-notes > li[data-dwb-note]").forEach((li) => {
      const key = li.dataset.dwbNote;
      if (notes.has(key)) { notes.get(key).extra = true; return; } // ten sam przypis drugi raz — tylko do odczytu
      const [kind, id] = key.split(":");
      const label = labels.get(key) || "";
      const el = noteXml(kind, id);
      const xps = el ? wKids(el, "p") : [];
      const paras = noteParas(li);
      let lock = null;
      if (!el || wKids(el).length !== xps.length || xps.length !== paras.length || li.children.length !== paras.length) lock = "lockNoteText";
      for (let i = 0; !lock && i < xps.length; i++) {
        lock = noteParaLock(xps[i]) || (stampParagraphLinks(xps[i], paras[i]) ? null : "lockLink");
      }
      let hasMark = false;
      if (!lock) {
        paras.forEach((p, i) => {
          const mark = stampMark(xps[i], p, nextKey, label);
          if (mark) hasMark = true;
          stampCommentMarks(xps[i], p, nextKey, mark ? [label.length] : []);
        });
        // zgodność podglądu z plikiem: ten sam tekst (bez tego zapis mógłby zgubić np. tabulator)
        const fromXml = xps.map((xp) => extractRunsFromParagraphXml(xp));
        const fromDom = paras.map((p) => extractRunsFromPreviewParagraph(p));
        const same = fromXml.every((runs, i) => previewRunsToPlainText(runs) === previewRunsToPlainText(fromDom[i])
          && runs.filter((r) => r.island).map((r) => r.island).join() === fromDom[i].filter((r) => r.island).map((r) => r.island).join());
        if (!same) lock = "lockNoteText";
        notes.set(key, { li, lock, base: fromDom, markHtml: hasMark ? noteParas(li)[0]?.querySelector("sup.note-mark")?.outerHTML : null });
      } else {
        notes.set(key, { li, lock, base: [], markHtml: null });
        // tylko do odczytu — numer i tak jak w Wordzie (sam wygląd, zapis tego przypisu nie dotyka)
        if (label && paras[0] && li.firstElementChild === paras[0]) {
          const sup = document.createElement("sup");
          sup.className = "note-mark";
          sup.textContent = label;
          sup.contentEditable = "false";
          paras[0].insertBefore(sup, paras[0].firstChild);
          hasMark = true;
        }
      }
      if (!hasMark && label) { // przypis bez numeru w treści (inny program) — numer z listy
        li.classList.add("dwb-note-native");
        // numer kolejny + rodzaj numeracji listy (końcowe „i, ii…”) — dawniej parseInt(„i”) = 0 → „0.”
        const nm = nums.get(key);
        li.value = nm?.n || Number.parseInt(label, 10) || 1;
        const css = { lowerRoman: "lower-roman", upperRoman: "upper-roman", lowerLetter: "lower-alpha", upperLetter: "upper-alpha" }[nm?.fmt];
        if (css) li.style.listStyleType = css;
      }
      paras.forEach((p, i) => { p.dataset.noteKey = key; p.dataset.noteSrc = String(i); });
    });
  }

  // ── 4) tryb edycji / czytania ──────────────────────────────────────────────
  function hintText(key, editable) {
    const text = texts.get(key);
    if (text == null) return "";
    const short = text.length > 700 ? `${text.slice(0, 700)}…` : text;
    return `${short.replace(/\n/g, "/|")}/|/|${t(editable ? "noteHintJumpEdit" : "noteHintJumpRead")}`;
  }
  function paintHints(editable) {
    hostEl()?.querySelectorAll("sup[data-dwb-note]").forEach((sup) => {
      const text = hintText(sup.dataset.dwbNote, editable);
      if (!text) return;
      sup.dataset.hint = "";
      sup.dataset.hintPl = text;
      sup.dataset.hintEn = text;
      sup.dataset.hintClass = "note-hint";
      sup.dataset.hintTouch = "on";
      sup.dataset.hintDelay = "0.15";
    });
  }

  function sync(editable) {
    notes.forEach((n) => {
      const lock = n.lock || (n.extra ? "lockNoteText" : null);
      noteParas(n.li).forEach((p) => {
        if (lock) {
          p.contentEditable = "false";
          p.classList.remove("docx-editable-p", "docx-editable-list");
          p.classList.toggle("docx-locked-p", editable);
          if (editable) {
            p.dataset.hint = "";
            p.dataset.hintPl = I18N.pl[lock];
            p.dataset.hintEn = I18N.en[lock];
            p.dataset.hintTouch = "on";
          } else ["hint", "hintPl", "hintEn", "hintTouch"].forEach((k) => delete p.dataset[k]);
          return;
        }
        if (editable) prepareEditableParagraph(p);
        else {
          p.contentEditable = "false";
          p.classList.remove("docx-editable-p", "docx-editable-list");
        }
      });
    });
    paintHints(editable);
  }

  // ── 5) zmiany do zapisu ────────────────────────────────────────────────────
  function collectEdits() {
    const out = [];
    notes.forEach((n, key) => {
      if (n.lock || n.extra || !n.li.isConnected) return;
      const paras = noteParas(n.li);
      const runs = paras.map((p) => extractRunsFromPreviewParagraph(p));
      const same = runs.length === n.base.length
        && runs.every((r, i) => runsEqual(r, n.base[i]))
        && paras.every((p, i) => p.dataset.noteSrc === String(i));
      if (!same) out.push({ note: key, paras: paras.map((p, i) => ({ runs: runs[i], src: Number(p.dataset.noteSrc) || 0 })) });
    });
    return out;
  }

  // po wpisaniu zmian do pliku (bez przerysowania): obecny podgląd = nowy stan pliku
  function rebase() {
    notes.forEach((n) => {
      if (n.lock || n.extra || !n.li.isConnected) return;
      const paras = noteParas(n.li);
      n.base = paras.map((p) => extractRunsFromPreviewParagraph(p));
      paras.forEach((p, i) => { p.dataset.noteSrc = String(i); });
    });
  }

  // ── 6) pisanie w przypisie ─────────────────────────────────────────────────
  // tekst przed kursorem bez numeru przypisu (wyspy) — „początek akapitu” jak w Wordzie
  function atParaStart(p) {
    const sel = window.getSelection();
    if (!sel?.rangeCount || !sel.isCollapsed) return false;
    const r = document.createRange();
    r.selectNodeContents(p);
    r.setEnd(sel.getRangeAt(0).startContainer, sel.getRangeAt(0).startOffset);
    const frag = r.cloneContents();
    frag.querySelectorAll?.('[data-cm-kind="note"]').forEach((x) => x.remove());
    return !frag.textContent.length;
  }

  // Zaznaczenie z numerem przypisu skasowane (np. zaznacz wszystko + pisanie) — numer wraca
  function ensureMark(key) {
    const n = notes.get(key);
    const first = noteParas(n?.li)[0];
    if (!n?.markHtml || !first || n.li.querySelector("sup.note-mark")) return;
    const tpl = document.createElement("template");
    tpl.innerHTML = n.markHtml;
    first.insertBefore(tpl.content.firstChild, first.firstChild);
  }

  function refreshText(key) {
    const n = notes.get(key);
    if (!n) return;
    texts.set(key, plainOf(noteParas(n.li).map((p) => extractRunsFromPreviewParagraph(p))));
    const editable = !readOnlyMode;
    hostEl()?.querySelectorAll(`sup[data-dwb-note="${CSS.escape(key)}"]`).forEach((sup) => {
      const text = hintText(key, editable);
      sup.dataset.hintPl = text;
      sup.dataset.hintEn = text;
    });
  }

  // Odnośnik skasowany w treści (Backspace, usunięcie zaznaczenia): w Wordzie przypis znika od
  // razu, a dalsze numery się przesuwają — rysujemy dokument od nowa (zapis i tak usuwa przypis
  // bez odnośnika: pruneOrphanNotesInZip). Kursor i przewinięcie wracają.
  let refTimer = 0;
  function bodyRefCount() {
    return hostEl()?.querySelectorAll("section.docx > article sup[data-dwb-note]").length || 0;
  }
  function watchRefs() {
    clearTimeout(refTimer);
    if (readOnlyMode || !refCount || bodyRefCount() >= refCount) return;
    refTimer = setTimeout(async () => {
      if (bodyRefCount() >= refCount) return;
      const p = typeof restoreDocCaret === "function" ? restoreDocCaret() : null;
      await mergeInlineEditsIntoBytes();
      const caret = p?.isConnected ? { paraIndex: resolveParaIndex(p), offset: getCaretOffset(p) } : null;
      if (caret && caret.paraIndex >= 0) pendingInlineCursor = caret;
      await reloadFromBytes(originalFileBytes);
    }, 450);
  }
  let refCount = 0;

  function onInput(e) {
    watchRefs();
    const p = e.target?.closest?.("p[data-note-key]");
    if (!p) return;
    ensureMark(p.dataset.noteKey);
    refreshText(p.dataset.noteKey);
  }

  function keydown(p, e) {
    if (e.key === "Enter") {
      e.preventDefault();
      if (e.shiftKey) document.execCommand("insertLineBreak");
      else {
        const np = splitParagraphDomAtCaret(p); // kopia akapitu: ten sam klucz i wzór (noteSrc)
        if (!np) return;
        prepareEditableParagraph(np);
        placeCaret(np, 0);
      }
      onInlineParagraphInput();
      refreshText(p.dataset.noteKey);
      return;
    }
    if (e.key === "Backspace" && atParaStart(p)) {
      e.preventDefault(); // nigdy nie kasuj numeru przypisu
      const prev = p.previousElementSibling;
      if (!prev || prev.tagName !== "P" || prev.dataset.noteKey !== p.dataset.noteKey || !prev.isContentEditable) return;
      const at = (prev.textContent || "").length;
      mergeParagraphDom(prev, p);
      placeCaret(prev, at);
      onInlineParagraphInput();
      refreshText(prev.dataset.noteKey);
      return;
    }
    if (e.key === "Tab") {
      e.preventDefault();
      insertTextAtCaret("\t");
      onInlineParagraphInput();
    }
  }

  // ── 7) skok odnośnik ⇄ przypis ─────────────────────────────────────────────
  function flash(el) {
    el.classList.add("note-flash");
    setTimeout(() => el.classList.remove("note-flash"), 1600);
  }
  function scrollToEl(el) {
    const vp = docViewportEl;
    const before = vp?.scrollTop || 0;
    el.scrollIntoView({ behavior: "smooth", block: "center" });
    if (typeof showLinkBack === "function") showLinkBack(before);
    flash(el);
  }
  function jumpToNote(key) {
    const li = hostEl()?.querySelector(`ol.dwb-notes > li[data-dwb-note="${CSS.escape(key)}"]`);
    if (!li) return;
    scrollToEl(li);
    const last = noteParas(li).filter((p) => p.isContentEditable).pop();
    if (last && !readOnlyMode) setTimeout(() => placeCaret(last, Number.MAX_SAFE_INTEGER), 30);
  }
  function jumpToRef(key) {
    const sup = hostEl()?.querySelector(`sup[data-dwb-note="${CSS.escape(key)}"]`);
    if (sup) scrollToEl(sup.closest("p") || sup);
  }

  function onClick(e) {
    const sup = e.target?.closest?.("sup[data-dwb-note]");
    if (sup && docCanvasEl.contains(sup)) {
      // Edycja: zwykły klik stawia kursor (jak w Wordzie) — skok dwuklikiem albo Ctrl/⌘+klik
      if (!readOnlyMode && !(e.ctrlKey || e.metaKey)) return;
      e.preventDefault();
      jumpToNote(sup.dataset.dwbNote);
      return;
    }
    const mark = e.target?.closest?.("sup.note-mark");
    if (mark) {
      e.preventDefault();
      jumpToRef(mark.closest("p[data-note-key]")?.dataset.noteKey || "");
    }
  }
  function onDblClick(e) {
    const sup = e.target?.closest?.("sup[data-dwb-note]");
    if (!sup || readOnlyMode || !docCanvasEl.contains(sup)) return;
    e.preventDefault();
    jumpToNote(sup.dataset.dwbNote);
  }

  // Kursor W ŚRODKU numeru przypisu (Chromium stawia go tam po kliknięciu za odnośnikiem albo
  // strzałką) — tam nie da się pisać, tekst lądował na początku akapitu. Przenosimy go tuż za
  // numer (albo przed, gdy stał na jego początku).
  function onSelectionChange() {
    if (readOnlyMode) return;
    const sel = window.getSelection();
    if (!sel?.rangeCount || !sel.isCollapsed) return;
    const node = sel.anchorNode;
    const island = (node?.nodeType === 1 ? node : node?.parentElement)?.closest?.('[data-cm-kind="note"]');
    const p = island?.closest?.(".docx-editable-p");
    if (!island || !p || !docCanvasEl?.contains(p)) return;
    const range = document.createRange();
    const atStart = sel.anchorOffset === 0 && (node === island || node === island.firstChild || node.parentElement === island && !node.previousSibling);
    if (atStart) range.setStartBefore(island); else range.setStartAfter(island);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
  }

  document.addEventListener("DOMContentLoaded", () => {
    document.addEventListener("selectionchange", onSelectionChange);
    docCanvasEl?.addEventListener("input", onInput);
    docCanvasEl?.addEventListener("click", onClick);
    docCanvasEl?.addEventListener("dblclick", onDblClick);
  });

  return { watchRefs, prepare, stampRefs, refLabelLengths, setupList, sync, collectEdits, rebase, keydown, hasEditable: () => [...notes.values()].some((n) => !n.lock && !n.extra) };
})();

// ── zapis: zmiany przypisów do footnotes.xml / endnotes.xml (buildPatchedDocx) ──
// item: { note: "footnote:2", paras: [{ runs, src }] } — src = akapit przypisu w pliku, z którego
// akapit podglądu powstał (Enter dzieli akapit → oba mają ten sam wzór wyglądu).
async function applyNoteEditsInZip(zip, items) {
  let count = 0;
  for (const kind of ["footnote", "endnote"]) {
    const mine = (items || []).filter((it) => String(it.note).startsWith(`${kind}:`));
    const file = mine.length ? zip.file(NOTE_PART[kind]) : null;
    if (!file) continue;
    const doc = new DOMParser().parseFromString(await file.async("string"), "application/xml");
    let changed = 0;
    mine.forEach((it) => {
      const id = it.note.slice(kind.length + 1);
      const noteEl = Array.from(doc.getElementsByTagNameNS(W_NS, kind)).find((n) => n.getAttributeNS(W_NS, "id") === id);
      if (!noteEl || !it.paras?.length) return;
      const kids = Array.from(noteEl.childNodes).filter((n) => n.nodeType === 1);
      const old = kids.filter((n) => n.namespaceURI === W_NS && n.localName === "p");
      if (!old.length || old.length !== kids.length) return; // tabela itp. w przypisie — nie ruszamy
      const used = new Set();
      const next = it.paras.map(({ runs, src }) => {
        const base = old[Math.min(Math.max(0, src | 0), old.length - 1)];
        let p = base;
        if (used.has(base)) { // kolejny akapit z tego samego wzoru — kopia bez identyfikatorów Worda
          p = base.cloneNode(true);
          Array.from(p.attributes).filter((a) => /^(paraId|textId|rsidR|rsidRDefault|rsidP)$/.test(a.localName)).forEach((a) => p.removeAttributeNode(a));
        }
        used.add(base);
        return { p, runs, fresh: p !== base };
      });
      let touched = old.some((p) => !used.has(p)) || next.length !== old.length;
      next.forEach(({ p, runs, fresh }, i) => {
        if (fresh || p !== old[i] || !runsEqual(extractRunsFromParagraphXml(p), runs)) {
          applyRunsToParagraphXml(p, runs);
          touched = true;
        }
      });
      if (!touched) return;
      old.forEach((p) => { if (!used.has(p)) noteEl.removeChild(p); });
      next.forEach(({ p }) => noteEl.appendChild(p)); // kolejność jak w podglądzie
      changed++;
    });
    if (!changed) continue;
    let xml = new XMLSerializer().serializeToString(doc);
    xml = await finalizeNoteLinks(zip, xml, kind);
    zip.file(NOTE_PART[kind], xml);
    count += changed;
  }
  return count;
}

// Link w przypisie: styl „Hiperłącze” i powiązanie adresu w powiązaniach TEJ części
// (finalizeComposeParts w docx-compose.js robi to samo dla document.xml).
async function finalizeNoteLinks(zip, xml, kind) {
  if (xml.includes('"__DWB_HL__"') && typeof composeEnsureCharStyle === "function") {
    const id = await composeEnsureCharStyle(zip, "Hyperlink", "Hyperlink",
      `<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>`);
    xml = xml.replaceAll('"__DWB_HL__"', `"${id}"`);
  }
  if (xml.includes("dwb-href") && typeof composeAddRel === "function") {
    const relPath = `word/_rels/${kind}s.xml.rels`;
    const relXml = zip.file(relPath) ? await zip.file(relPath).async("string")
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    const rels = composeParse(relXml);
    const doc = composeParse(xml);
    Array.from(doc.getElementsByTagNameNS(W_NS, "hyperlink")).forEach((h) => {
      const href = h.getAttribute("dwb-href");
      if (href == null) return;
      h.removeAttribute("dwb-href");
      h.setAttributeNS(R_NS, "r:id", composeAddRel(rels, COMPOSE_REL_HYPERLINK, href, true));
    });
    zip.file(relPath, composeSerialize(rels));
    xml = composeSerialize(doc);
  }
  return xml;
}

// ── nowy przypis (Wstaw → Przypis dolny / końcowy, Ctrl/⌘+Alt+F / D — jak w Wordzie) ──────────
// edit: { kind, index (akapit treści), offset (miejsce w tekście akapitu jak w pliku), id }.
// Brakująca część footnotes.xml / endnotes.xml powstaje z separatorami (jak w nowym pliku Worda);
// style „footnote text” / „footnote reference” (i końcowe) — dopisywane, gdy ich nie ma.
const NOTE_CT = { footnote: "footnotes", endnote: "endnotes" };
function noteStyleDefs(kind, normalId) {
  const based = normalId ? `<w:basedOn w:val="${normalId}"/>` : "";
  return {
    text: { name: `${kind} text`, id: kind === "footnote" ? "FootnoteText" : "EndnoteText",
      body: `${based}<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>` },
    ref: { name: `${kind} reference`, id: kind === "footnote" ? "FootnoteReference" : "EndnoteReference",
      body: `<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>` },
  };
}
async function noteEnsureStyles(zip, kind) {
  const xml = await composeEnsureStylesPart(zip);
  const doc = composeParse(xml);
  const styles = Array.from(doc.getElementsByTagNameNS(W_NS, "style"));
  const normalId = styles.find((st) => st.getAttributeNS(W_NS, "type") === "paragraph" && /^(1|true)$/.test(st.getAttributeNS(W_NS, "default") || ""))?.getAttributeNS(W_NS, "styleId") || null;
  const ids = new Set(styles.map((st) => st.getAttributeNS(W_NS, "styleId")));
  const defs = noteStyleDefs(kind, normalId);
  const out = {};
  for (const [key, type] of [["text", "paragraph"], ["ref", "character"]]) {
    const def = defs[key];
    const found = styles.find((st) => st.getAttributeNS(W_NS, "type") === type && (composeDirectChild(st, "name")?.getAttributeNS(W_NS, "val") || "").trim().toLowerCase() === def.name);
    if (found) { out[key] = found.getAttributeNS(W_NS, "styleId"); continue; }
    let id = def.id;
    while (ids.has(id)) id += "1";
    ids.add(id);
    const frag = composeParse(`<w:styles xmlns:w="${W_NS}"><w:style w:type="${type}" w:styleId="${id}"><w:name w:val="${def.name}"/>${def.body}</w:style></w:styles>`);
    doc.documentElement.appendChild(doc.importNode(frag.documentElement.firstChild, true));
    out[key] = id;
  }
  zip.file("word/styles.xml", composeSerialize(doc));
  return out;
}
function noteNextId(partXml, kind) {
  let max = 0;
  for (const m of String(partXml || "").matchAll(new RegExp(`<w:${kind}\\b[^>]*\\bw:id="(-?\\d+)"`, "g"))) max = Math.max(max, parseInt(m[1], 10));
  return max + 1;
}
async function applyNoteInsertInZip(zip, xml, edit) {
  const kind = edit.kind === "endnote" ? "endnote" : "footnote";
  const doc = composeParse(xml);
  const p = collectParagraphElements(doc.documentElement, "all")[edit.index];
  if (!p) return { xml, count: 0 };
  const sep = (type, id, tag) => `<w:${kind} w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${tag}/></w:r></w:p></w:${kind}>`;
  const partXml = await composeEnsurePart(zip, NOTE_PART[kind],
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${kind}s xmlns:w="${W_NS}" xmlns:r="${R_NS}">${sep("separator", -1, "separator")}${sep("continuationSeparator", 0, "continuationSeparator")}</w:${kind}s>`,
    `application/vnd.openxmlformats-officedocument.wordprocessingml.${NOTE_CT[kind]}+xml`,
    `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${NOTE_CT[kind]}`);
  const st = await noteEnsureStyles(zip, kind);
  const nextId = noteNextId(partXml, kind);
  const id = Number.isInteger(edit.id) && edit.id >= nextId ? edit.id : nextId;
  const partDoc = composeParse(partXml);
  const noteFrag = composeParse(`<w:${kind}s xmlns:w="${W_NS}"><w:${kind} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="${st.text}"/></w:pPr><w:r><w:rPr><w:rStyle w:val="${st.ref}"/></w:rPr><w:${kind}Ref/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r></w:p></w:${kind}></w:${kind}s>`);
  partDoc.documentElement.appendChild(partDoc.importNode(noteFrag.documentElement.firstChild, true));
  zip.file(NOTE_PART[kind], composeSerialize(partDoc));
  // odnośnik w treści — „wyspa” w modelu akapitu (cały fragment z pliku wraca przy zapisie)
  const runs = extractRunsFromParagraphXml(p);
  const total = previewRunsToPlainText(runs).length;
  const at = Math.max(0, Math.min(total, edit.offset | 0));
  const { before, after } = composeSliceRuns(runs, at, at);
  const ref = `<w:r xmlns:w="${W_NS}"><w:rPr><w:rStyle w:val="${st.ref}"/></w:rPr><w:${kind}Reference w:id="${id}"/></w:r>`;
  applyRunsToParagraphXml(p, [...before, { island: ref, text: "" }, ...after]);
  return { xml: composeSerialize(doc), count: 1, id };
}

// Przypisy bez odnośnika w treści (odnośnik usunięty razem z tekstem) — usuwane z pliku jak
// w Wordzie; separatory (w:type) zostają.
async function pruneOrphanNotesInZip(zip, docXml) {
  for (const kind of ["footnote", "endnote"]) {
    const file = zip.file(NOTE_PART[kind]);
    if (!file) continue;
    const used = new Set([...String(docXml).matchAll(new RegExp(`<w:${kind}Reference\\b[^>]*\\bw:id="(-?\\d+)"`, "g"))].map((m) => m[1]));
    const xml = await file.async("string");
    const doc = composeParse(xml);
    const orphans = Array.from(doc.getElementsByTagNameNS(W_NS, kind)).filter((n) => !n.getAttributeNS(W_NS, "type") && !used.has(n.getAttributeNS(W_NS, "id")));
    if (!orphans.length) continue;
    orphans.forEach((n) => n.parentNode.removeChild(n));
    zip.file(NOTE_PART[kind], composeSerialize(doc));
  }
}
