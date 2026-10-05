// compose-ui.js — tworzenie dokumentu od zera (Etap 1, 2026-10-01): interfejs.
// Operacje na pliku są w docx-compose.js; tu tylko przyciski i okienka.
//
// Zasada ekranu (prośba Mateusza: „żeby nie zawalić całego ekranu”):
//   - na pasku przybywa JEDEN przycisk „＋ Wstaw”, lista stylu i jeden przycisk wyrównania,
//     tylko w trybie Edycja; cały rząd formatowania przewija się w bok z wygaszeniem krawędzi,
//   - menu „Wstaw” i wybór wyrównania to okienka nad dokumentem, otwierane na żądanie,
//     zamykane klikiem obok / Esc / przewinięciem. Okienko jest poza rzędem (rząd przewija się,
//     overflow by je przycinał), kotwiczone POD przyciskiem — na telefonie z klawiaturą
//     arkusz od dołu chowałby się pod klawiaturą (iOS: fixed liczy się od layout viewportu).
//   - przyciski w okienkach nie zabierają fokusu z tekstu (mousedown → preventDefault; NIE
//     pointerdown — w Safari/iOS kasuje to tap), więc na iPhonie klawiatura i kursor zostają.
//   - rzeczy kontekstowe pojawiają się tylko, gdy są potrzebne: przycisk „Tabela” na pasku, gdy
//     kursor stoi w tabeli (Etap 3), karta obrazu po kliknięciu obrazu, karta linku w linku.

const composeUi = (() => {
  const bar = document.getElementById("formatToolbar");
  const insertBtn = document.getElementById("insertMenuBtn");
  const styleSel = document.getElementById("fmtParaStyle");
  const alignBtn = document.getElementById("fmtAlignBtn");
  const listBtn = document.getElementById("fmtListBtn");
  const tableBtn = document.getElementById("tableToolsBtn");
  const layoutBtn = document.getElementById("pageLayoutBtn");
  const colorBtn = document.getElementById("fmtColorBtn");
  const colorBar = document.getElementById("fmtColorBar");
  const dialog = document.getElementById("newDocDialog");

  const ALIGN_ICONS = {
    left: '<line x1="4" y1="6" x2="20" y2="6"/><line x1="4" y1="10" x2="14" y2="10"/><line x1="4" y1="14" x2="20" y2="14"/><line x1="4" y1="18" x2="14" y2="18"/>',
    center: '<line x1="4" y1="6" x2="20" y2="6"/><line x1="7" y1="10" x2="17" y2="10"/><line x1="4" y1="14" x2="20" y2="14"/><line x1="7" y1="18" x2="17" y2="18"/>',
    right: '<line x1="4" y1="6" x2="20" y2="6"/><line x1="10" y1="10" x2="20" y2="10"/><line x1="4" y1="14" x2="20" y2="14"/><line x1="10" y1="18" x2="20" y2="18"/>',
    both: '<line x1="4" y1="6" x2="20" y2="6"/><line x1="4" y1="10" x2="20" y2="10"/><line x1="4" y1="14" x2="20" y2="14"/><line x1="4" y1="18" x2="20" y2="18"/>',
  };
  const ALIGN_KEYS = { left: "alignLeft", center: "alignCenter", right: "alignRight", both: "alignBoth" };
  const SPECIAL_CHARS = ["„", "”", "«", "»", "–", "—", "…", "§", "©", "®", "™", "°", "±", "×", "÷", "€", "½", "✓", "•", "→"];

  const alignSvg = (key) => `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" aria-hidden="true">${ALIGN_ICONS[key]}</svg>`;

  // ── kursor / akapity ──────────────────────────────────────────────────────
  function caretParagraph() {
    const p = typeof restoreDocCaret === "function" ? restoreDocCaret() : null;
    if (!p) toast(t("insertNoCaret"), "info");
    return p;
  }

  // Akapity objęte zaznaczeniem (styl/wyrównanie dla kilku naraz), inaczej akapit z kursorem.
  function selectedParagraphIndices(p) {
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    const paras = collectPreviewParagraphElements(host);
    const sel = window.getSelection();
    const out = [];
    if (sel?.rangeCount && !sel.getRangeAt(0).collapsed) {
      const range = sel.getRangeAt(0);
      paras.forEach((el, i) => { if (range.intersectsNode(el)) out.push(i); });
    }
    if (!out.length) {
      const i = resolveParaIndex(p);
      if (i >= 0) out.push(i);
    }
    return out;
  }

  function caretState(p) {
    return { paraIndex: resolveParaIndex(p), offset: getCaretOffset(p) };
  }

  // Zmiana w pliku + przerysowanie; kursor wraca, miejsce na ekranie trzyma reloadFromBytes.
  async function runFileEdit(edit, caret) {
    if (caret) pendingInlineCursor = caret;
    try {
      await applyDocumentEdit(edit);
    } catch (err) {
      log(`Wstaw/format: ${err.message || err}`, "error");
      toast(t("saveFailed"), "error");
    }
  }

  function textAround(p) {
    const sel = window.getSelection();
    if (!sel?.rangeCount || !p.contains(sel.getRangeAt(0).startContainer)) return { before: "", after: "x" };
    const r = sel.getRangeAt(0);
    const before = document.createRange();
    before.selectNodeContents(p);
    before.setEnd(r.startContainer, r.startOffset);
    const after = document.createRange();
    after.selectNodeContents(p);
    after.setStart(r.endContainer, r.endOffset);
    const hasBr = (range) => { const f = range.cloneContents(); return !!f.querySelector?.("br"); };
    return {
      before: before.toString() + (hasBr(before) ? "\n" : ""),
      after: after.toString() + (hasBr(after) ? "\n" : ""),
    };
  }

  // ── operacje ──────────────────────────────────────────────────────────────
  async function applyStyle(key) {
    const p = caretParagraph();
    if (!p) { syncState(); return; }
    await runFileEdit({ op: "paraFormat", indices: selectedParagraphIndices(p), style: key }, caretState(p));
  }

  async function applyAlign(key) {
    const p = caretParagraph();
    if (!p) return;
    await runFileEdit({ op: "paraFormat", indices: selectedParagraphIndices(p), align: key }, caretState(p));
  }

  // Lista: toggle = ta sama lista drugi raz zdejmuje ją (jak przycisk w Wordzie)
  async function applyList(kind, toggle = true) {
    const p = caretParagraph();
    if (!p) return;
    await runFileEdit({ op: "list", indices: selectedParagraphIndices(p), kind, toggle }, caretState(p));
  }

  async function changeListLevel(delta) {
    const p = caretParagraph();
    if (!p) return;
    if (!isListParagraph(p)) { toast(t("listNotInList"), "info"); return; }
    const index = resolveParaIndex(p);
    asUndoStep("undoOpList", () => applyDomListLevel({ index, delta }));
    await applyInlineStructuralEdit({ op: "listLevel", index, delta });
  }

  // Enter w pustym punkcie / Backspace na początku punktu (docx-inline-edit.js) — koniec listy.
  function endListAt(p) {
    const index = resolveParaIndex(p);
    if (index < 0) return false;
    runFileEdit({ op: "list", indices: [index], kind: "none" }, { paraIndex: index, offset: 0 });
    return true;
  }

  // ── spis treści ───────────────────────────────────────────────────────────
  // Numery stron z podglądu (strony = sekcje docx-preview; w trybie dopasowania do telefonu ich nie ma).
  function previewPages() {
    const pages = {};
    if (typeof shouldUseMobileReflow === "function" && shouldUseMobileReflow()) return pages;
    const host = docCanvasEl?.querySelector(".docx-preview-host");
    if (!host) return pages;
    // docx-preview łamie strony tylko przy jawnych podziałach — dłuższy tekst bez nich to jedna
    // wysoka „kartka”. Numer strony szacujemy z położenia nagłówka i wysokości strony
    // (min-height kartki = wysokość strony z pliku). Word i tak przelicza je przy „Aktualizuj pole”.
    const sections = Array.from(host.querySelectorAll("section.docx"));
    const startPage = new Map();
    let page = 1;
    sections.forEach((sec) => {
      startPage.set(sec, page);
      const ph = parseFloat(getComputedStyle(sec).minHeight) || sec.offsetHeight || 1;
      page += Math.max(1, Math.ceil((sec.offsetHeight - 2) / ph));
    });
    (documentStructure?.headings || []).forEach((h) => {
      const sec = h.el?.closest?.("section.docx");
      if (!sec || !Number.isFinite(h.paraIndex)) return;
      const ph = parseFloat(getComputedStyle(sec).minHeight) || sec.offsetHeight || 1;
      const y = h.el.getBoundingClientRect().top - sec.getBoundingClientRect().top;
      pages[h.paraIndex] = startPage.get(sec) + Math.max(0, Math.floor(y / ph));
    });
    return pages;
  }

  function paragraphCount() {
    return collectPreviewParagraphElements(docCanvasEl?.querySelector(".docx-preview-host")).length;
  }

  async function insertToc() {
    const has = typeof composeHasToc === "function" && composeHasToc();
    const p = has ? (activeParagraph() || lastP) : caretParagraph();
    if (!has && !p) return;
    const index = p ? resolveParaIndex(p) : 0;
    const offset = p ? getCaretOffset(p) : 0;
    const countBefore = paragraphCount();
    // pierwszy akapit spisu (przy aktualizacji) — kursor za spisem przesuwa się o zmianę liczby akapitów
    const tocFirst = has ? collectPreviewParagraphElements(docCanvasEl?.querySelector(".docx-preview-host"))
      .findIndex((q) => Array.from(q.classList).some((c) => COMPOSE_TOC_KEYS.includes(docComposeStyleClasses.get(c)))) : index;
    const base = { op: "toc", index, title: t("tocTitle"), emptyText: t("tocEmpty") };
    await runFileEdit({ ...base, pages: previewPages() }, null);
    // drugi przebieg: numery stron PO wstawieniu spisu (spis sam zajmuje miejsce) — bez nowego kroku cofania
    try {
      const pages = previewPages();
      if (Object.keys(pages).length) {
        const { bytes } = await buildPatchedDocx(originalFileBytes, [{ ...base, pages }]);
        originalFileBytes = bytes;
        quietRenderOnce = true;
        await reloadFromBytes(bytes);
      }
    } catch (err) { log(`Spis treści: ${err.message || err}`, "error"); }
    const shift = paragraphCount() - countBefore;
    if (p && !readOnlyMode) {
      const target = { paraIndex: index >= tocFirst ? index + shift : index, offset };
      if (inlineLocksPending) pendingInlineCursor = target; // podgląd jeszcze się przygotowuje
      else focusParagraphAtOffset(target.paraIndex, target.offset);
    }
    toast(t(has ? "tocUpdated" : "tocInserted"), "success");
  }

  // ── pola formularza ───────────────────────────────────────────────────────
  async function insertFormField(kind, options) {
    const p = caretParagraph();
    if (!p) return;
    const index = resolveParaIndex(p);
    if (index < 0) return;
    const sel = window.getSelection();
    const range = sel.getRangeAt(0);
    const offset = textOffset(p, range.startContainer, range.startOffset);
    await runFileEdit({ op: "formInsert", index, offset, kind, options, lang: currentLang }, { paraIndex: index, offset: offset + 1 });
  }

  function openListFieldForm() {
    const p = caretParagraph();
    if (!p) return;
    openPop(insertBtn, (el) => {
      el.classList.add("compose-pop-form");
      el.setAttribute("role", "dialog");
      el.setAttribute("aria-label", t("formInsertList"));
      el.innerHTML = `<div class="compose-cap"></div>
        <label class="compose-field"><span></span><textarea class="lf-options" rows="5" data-autofocus="1"></textarea></label>
        <div class="compose-actions"><span class="compose-actions-gap"></span><button type="button" class="btn lf-cancel"></button><button type="button" class="btn primary lf-ok"></button></div>`;
      el.querySelector(".compose-cap").textContent = t("formInsertList");
      el.querySelector(".compose-field span").textContent = t("formListOptions");
      const ta = el.querySelector(".lf-options");
      ta.placeholder = t("formListPlaceholder");
      el.querySelector(".lf-cancel").textContent = t("linkCancel");
      el.querySelector(".lf-cancel").addEventListener("click", () => closePop());
      const ok = el.querySelector(".lf-ok");
      ok.textContent = t("formInsertBtn");
      ok.addEventListener("click", () => {
        const options = ta.value.split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
        if (!options.length) { toast(t("formListEmpty"), "info"); ta.focus(); return; }
        closePop();
        restoreDocCaret();
        insertFormField("dropdown", options);
      });
    });
  }

  async function insertPageBreak() {
    const p = caretParagraph();
    if (!p) return;
    const index = resolveParaIndex(p);
    if (index < 0) return;
    const { before, after } = textAround(p);
    if (!after) {
      await runFileEdit({ op: "pageBreak", index, mode: "after" }, { paraIndex: index + 1, offset: 0 });
    } else if (!before) {
      await runFileEdit({ op: "pageBreak", index, mode: "before" }, { paraIndex: index ? index : 1, offset: 0 });
    } else {
      // Tekst za kursorem idzie na nową stronę: formatowanie obu części z podglądu.
      // Podział w DOM tylko na chwilę (odczyt fragmentów) — zaraz scalony z powrotem,
      // bo applyDocumentEdit najpierw zapisuje do pliku akapity z podglądu w ich kolejności.
      const tail = splitParagraphDomAtCaret(p);
      if (!tail) return;
      const beforeRuns = extractRunsFromPreviewParagraph(p);
      const afterRuns = extractRunsFromPreviewParagraph(tail);
      mergeParagraphDom(p, tail);
      await runFileEdit({ op: "pageBreak", index, mode: "split", beforeRuns, afterRuns }, { paraIndex: index + 1, offset: 0 });
    }
  }

  async function insertHrule() {
    const p = caretParagraph();
    if (!p) return;
    const index = resolveParaIndex(p);
    if (index < 0) return;
    await runFileEdit({ op: "hrule", index }, { paraIndex: index + 2, offset: 0 });
  }

  function insertText(text) {
    const p = caretParagraph();
    if (!p) return;
    asUndoStep("undoOpInsert", () => {
      if (!document.execCommand("insertText", false, text)) insertTextAtCaret(text);
    });
    onInlineParagraphInput();
  }

  function todayText() {
    return new Date().toLocaleDateString(currentLang === "en" ? "en-GB" : "pl-PL", { day: "numeric", month: "long", year: "numeric" });
  }

  // ── okienko (menu Wstaw / wyrównanie) ─────────────────────────────────────
  let pop = null; // { el, anchor }

  function closePop(focusAnchor = false) {
    if (!pop) return;
    const { el, anchor } = pop;
    pop = null;
    el.remove();
    anchor.setAttribute("aria-expanded", "false");
    anchor.classList.remove("is-on");
    if (focusAnchor) anchor.focus();
  }

  function placePop(el, anchor) {
    const vv = window.visualViewport;
    const r = anchor.getBoundingClientRect(); // układ widocznego obszaru (iOS z klawiaturą ≠ fixed)
    const viewW = vv ? vv.width : window.innerWidth;
    const viewH = vv ? vv.height : window.innerHeight;
    const w = el.offsetWidth;
    const left = Math.max(8, Math.min(r.left, viewW - w - 8));
    const top = r.bottom + 6;
    el.style.maxHeight = `${Math.max(160, viewH - top - 10)}px`;
    el.style.left = `${left + (vv ? vv.offsetLeft : 0)}px`;
    el.style.top = `${top + (vv ? vv.offsetTop : 0)}px`;
  }

  function openPop(anchor, build) {
    const same = pop?.anchor === anchor;
    closePop();
    hideLinkCard(); // karty pod kursorem nie wiszą pod okienkiem
    if (typeof hideCommentCard === "function") hideCommentCard();
    if (same) return; // drugi klik w ten sam przycisk = zamknij
    const el = document.createElement("div");
    el.className = "compose-pop";
    el.setAttribute("role", "menu");
    const built = build(el);
    // okienko budowane asynchronicznie (czyta plik, np. Układ strony) — położenie po zbudowaniu
    if (built?.then) built.then(() => { if (pop?.el === el) placePop(el, anchor); }).catch(() => {});
    // przyciski nie zabierają fokusu z dokumentu (kursor i klawiatura ekranowa zostają)
    el.addEventListener("mousedown", (e) => { if (e.target.closest("button")) e.preventDefault(); });
    el.addEventListener("keydown", (e) => {
      const items = Array.from(el.querySelectorAll("button"));
      const i = items.indexOf(document.activeElement);
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closePop(true); }
      else if (e.target.closest("input, select, textarea")) { /* pole formularza — strzałki dla pola */ }
      else if (e.key === "ArrowDown" || e.key === "ArrowRight") { e.preventDefault(); items[(i + 1) % items.length]?.focus(); }
      else if (e.key === "ArrowUp" || e.key === "ArrowLeft") { e.preventDefault(); items[(i - 1 + items.length) % items.length]?.focus(); }
      else if (e.key === "Tab") closePop();
    });
    document.body.appendChild(el);
    pop = { el, anchor };
    anchor.setAttribute("aria-expanded", "true");
    anchor.classList.add("is-on");
    placePop(el, anchor);
    // z klawiatury (Enter/spacja na przycisku) — fokus do pierwszej pozycji
    const autofocus = el.querySelector("[data-autofocus]");
    if (autofocus) autofocus.focus({ preventScroll: true });
    else if (document.activeElement === anchor) el.querySelector("button")?.focus();
  }

  function popItem(el, { label, desc, kbd, icon, onPick }) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "compose-item";
    b.setAttribute("role", "menuitem");
    b.innerHTML = `<span class="compose-item-icon" aria-hidden="true">${icon || ""}</span><span class="compose-item-text"><span class="compose-item-label"></span>${desc ? '<span class="compose-item-desc"></span>' : ""}</span>${kbd ? `<span class="app-menu-kbd">${kbd}</span>` : ""}`;
    b.querySelector(".compose-item-label").textContent = label;
    if (desc) b.querySelector(".compose-item-desc").textContent = desc;
    b.addEventListener("click", () => { closePop(); onPick(); });
    el.appendChild(b);
    return b;
  }

  function popCap(el, text) {
    const c = document.createElement("div");
    c.className = "compose-cap";
    c.textContent = text;
    el.appendChild(c);
  }

  const ICON = (path) => `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${path}</svg>`;

  function buildInsertMenu(el) {
    el.classList.add("compose-pop-insert");
    popCap(el, t("insertGroupPage"));
    popItem(el, {
      label: t("insertPageBreak"), desc: t("insertPageBreakDesc"), kbd: "Ctrl/⌘+Enter",
      icon: ICON('<path d="M6 3v5h12V3"/><path d="M6 21v-5h12v5"/><line x1="3" y1="12" x2="6" y2="12"/><line x1="9" y1="12" x2="11" y2="12"/><line x1="13" y1="12" x2="15" y2="12"/><line x1="18" y1="12" x2="21" y2="12"/>'),
      onPick: insertPageBreak,
    });
    popItem(el, {
      label: t("hfInsert"), desc: t("hfInsertDesc"),
      icon: ICON('<rect x="4" y="3" width="16" height="18" rx="2"/><line x1="7" y1="6.5" x2="17" y2="6.5"/><line x1="7" y1="17.5" x2="12" y2="17.5"/><line x1="15" y1="17.5" x2="17" y2="17.5"/>'),
      onPick: () => openHeaderFooterForm(),
    });
    popItem(el, {
      label: t("insertHrule"), desc: t("insertHruleDesc"),
      icon: ICON('<line x1="3" y1="12" x2="21" y2="12"/>'),
      onPick: insertHrule,
    });
    popCap(el, t("insertGroupText"));
    popItem(el, {
      label: t("insertLink"), desc: t("insertLinkDesc"), kbd: "Ctrl/⌘+K",
      icon: ICON('<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>'),
      onPick: () => openLinkForm(insertBtn),
    });
    popItem(el, {
      label: t("insertComment"), desc: t("insertCommentDesc"), kbd: "Ctrl/⌘+Alt+M",
      icon: ICON('<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><line x1="8" y1="9" x2="16" y2="9"/><line x1="8" y1="13" x2="13" y2="13"/>'),
      onPick: () => openCommentForm(insertBtn),
    });
    popItem(el, {
      label: t("insertFootnote"), desc: t("insertFootnoteDesc"), kbd: "Ctrl/⌘+Alt+F",
      icon: ICON('<line x1="4" y1="7" x2="15" y2="7"/><path d="M17.5 4.5v4"/><line x1="4" y1="12" x2="20" y2="12"/><line x1="4" y1="18.5" x2="11" y2="18.5"/><path d="M13 18.5h7" stroke-dasharray="1.5 1.5"/>'),
      onPick: () => insertNote("footnote"),
    });
    popItem(el, {
      label: t("insertEndnote"), desc: t("insertEndnoteDesc"), kbd: "Ctrl/⌘+Alt+D",
      icon: ICON('<line x1="4" y1="6" x2="15" y2="6"/><path d="M17.5 3.5v4"/><line x1="4" y1="11" x2="20" y2="11"/><line x1="4" y1="16" x2="20" y2="16"/><path d="M4 20.5h16"/>'),
      onPick: () => insertNote("endnote"),
    });
    popItem(el, {
      label: t("insertDate"), desc: todayText(),
      icon: ICON('<rect x="3" y="5" width="18" height="16" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="8" y1="3" x2="8" y2="7"/><line x1="16" y1="3" x2="16" y2="7"/>'),
      onPick: () => insertText(todayText()),
    });
    popCap(el, t("insertGroupDoc"));
    popItem(el, {
      label: t("insertTable"), desc: t("insertTableDesc"),
      icon: ICON('<rect x="3" y="4" width="18" height="16" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="4" x2="9" y2="20"/><line x1="15" y1="4" x2="15" y2="20"/>'),
      onPick: () => openPop(insertBtn, buildTablePicker),
    });
    popItem(el, {
      label: t("insertImage"), desc: t("insertImageDesc"),
      icon: ICON('<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="m21 17-5-5-8 8"/>'),
      onPick: pickImage,
    });
    const hasToc = typeof composeHasToc === "function" && composeHasToc();
    popItem(el, {
      label: t(hasToc ? "tocUpdate" : "tocInsert"), desc: t(hasToc ? "tocUpdateDesc" : "tocInsertDesc"),
      icon: ICON('<line x1="4" y1="6" x2="14" y2="6"/><line x1="4" y1="11" x2="17" y2="11"/><line x1="7" y1="16" x2="17" y2="16"/><line x1="18" y1="6" x2="20" y2="6"/><line x1="19.5" y1="11" x2="20" y2="11"/><line x1="19.5" y1="16" x2="20" y2="16"/>'),
      onPick: insertToc,
    });
    buildInsertMoreGroups(el);
    popCap(el, t("insertGroupChars"));
    const grid = document.createElement("div");
    grid.className = "compose-chars";
    SPECIAL_CHARS.forEach((ch) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "compose-char";
      b.textContent = ch;
      b.setAttribute("aria-label", ch);
      b.addEventListener("click", () => { closePop(); insertText(ch); });
      grid.appendChild(b);
    });
    el.appendChild(grid);
  }

  function buildListMenu(el) {
    el.classList.add("compose-pop-list");
    popItem(el, {
      label: t("listBullet"), desc: "• • •",
      icon: ICON('<line x1="9" y1="6" x2="20" y2="6"/><line x1="9" y1="12" x2="20" y2="12"/><line x1="9" y1="18" x2="20" y2="18"/><circle cx="4.5" cy="6" r="1" fill="currentColor"/><circle cx="4.5" cy="12" r="1" fill="currentColor"/><circle cx="4.5" cy="18" r="1" fill="currentColor"/>'),
      onPick: () => applyList("bullet"),
    });
    popItem(el, {
      label: t("listNumber"), desc: "1. 2. 3.",
      icon: ICON('<line x1="10" y1="6" x2="21" y2="6"/><line x1="10" y1="12" x2="21" y2="12"/><line x1="10" y1="18" x2="21" y2="18"/><path d="M4 6h1v4"/><path d="M4 10h2"/><path d="M6 18H4c0-1 2-2 2-3s-1-1.5-2-1"/>'),
      onPick: () => applyList("number"),
    });
    const p = activeParagraph() || lastP;
    const inList = !!p && isListParagraph(p);
    if (inList) {
      popItem(el, { label: t("listNone"), icon: ICON('<line x1="5" y1="12" x2="19" y2="12"/>'), onPick: () => applyList("none", false) });
      popCap(el, t("listLevel"));
      const row = document.createElement("div");
      row.className = "compose-row";
      [["listOutdent", -1, '<polyline points="11 17 6 12 11 7"/><line x1="18" y1="12" x2="6" y2="12"/>'], ["listIndent", 1, '<polyline points="13 7 18 12 13 17"/><line x1="6" y1="12" x2="18" y2="12"/>']].forEach(([key, delta, svg]) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "btn compose-row-btn";
        b.innerHTML = `${ICON(svg)}<span></span>`;
        b.querySelector("span").textContent = t(key);
        b.addEventListener("click", () => { closePop(); changeListLevel(delta); });
        row.appendChild(b);
      });
      el.appendChild(row);
    }
  }

  function buildInsertMoreGroups(el) {
    popCap(el, t("insertGroupForm"));
    const grid = document.createElement("div");
    grid.className = "compose-grid2";
    [
      ["formInsertText", '<rect x="3" y="7" width="18" height="10" rx="2"/><line x1="7" y1="12" x2="12" y2="12"/>', () => insertFormField("text")],
      ["formInsertList", '<rect x="3" y="7" width="18" height="10" rx="2"/><polyline points="14 11 16 13 18 11"/>', openListFieldForm],
      ["formInsertDate", '<rect x="3" y="5" width="18" height="16" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="8" y1="3" x2="8" y2="7"/><line x1="16" y1="3" x2="16" y2="7"/>', () => insertFormField("date")],
      ["formInsertCheck", '<rect x="5" y="5" width="14" height="14" rx="2"/><polyline points="8.5 12 11 14.5 15.5 9.5"/>', () => insertFormField("checkbox")],
    ].forEach(([key, svg, fn]) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "compose-item compose-item-sm";
      b.setAttribute("role", "menuitem");
      b.innerHTML = `<span class="compose-item-icon" aria-hidden="true">${ICON(svg)}</span><span class="compose-item-label"></span>`;
      b.querySelector(".compose-item-label").textContent = t(key);
      b.addEventListener("click", () => { closePop(); fn(); });
      grid.appendChild(b);
    });
    el.appendChild(grid);
  }

  function buildAlignMenu(el) {
    el.classList.add("compose-pop-align");
    const current = currentAlign();
    Object.keys(ALIGN_ICONS).forEach((key) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "tb-btn compose-align";
      b.setAttribute("role", "menuitemradio");
      b.setAttribute("aria-checked", String(key === current));
      b.classList.toggle("is-on", key === current);
      b.setAttribute("aria-label", t(ALIGN_KEYS[key]));
      b.dataset.hint = "";
      b.dataset.hintPl = I18N.pl[ALIGN_KEYS[key]];
      b.dataset.hintEn = I18N.en[ALIGN_KEYS[key]];
      b.dataset.hintDelay = "0.4";
      b.innerHTML = alignSvg(key);
      b.addEventListener("click", () => { closePop(); applyAlign(key); });
      el.appendChild(b);
    });
  }

  // ── link: okienko „Link” i karta linku pod kursorem ───────────────────────
  const LINK_OK_RE = /^(https?:|mailto:|tel:)/i;

  // Przesunięcie w tekście akapitu (jak w pliku: łamanie wiersza = 1 znak).
  function textOffset(p, node, offset) {
    const r = document.createRange();
    r.selectNodeContents(p);
    try { r.setEnd(node, offset); } catch (_) { return 0; }
    const div = document.createElement("div");
    div.appendChild(r.cloneContents());
    return previewRunsToPlainText(extractRunsFromPreviewParagraph(div)).length;
  }

  function linkAtCaret() {
    const sel = window.getSelection();
    if (!sel?.rangeCount) return null;
    const node = sel.getRangeAt(0).startContainer;
    const a = (node.nodeType === 1 ? node : node.parentElement)?.closest?.("a[href]");
    return a && !a.classList.contains("doc-xref") && a.closest(".docx-editable-p") ? a : null;
  }

  function linkContext() {
    const p = caretParagraph();
    if (!p) return null;
    const index = resolveParaIndex(p);
    if (index < 0) return null;
    const a = linkAtCaret();
    if (a && p.contains(a)) {
      const r = document.createRange();
      r.selectNodeContents(a);
      return { p, index, a, start: textOffset(p, r.startContainer, r.startOffset), end: textOffset(p, r.endContainer, r.endOffset), text: a.textContent || "", link: a.dataset.dwbLink || a.getAttribute("href") || "" };
    }
    const range = window.getSelection().getRangeAt(0);
    const start = textOffset(p, range.startContainer, range.startOffset);
    const end = p.contains(range.endContainer) ? textOffset(p, range.endContainer, range.endOffset) : previewRunsToPlainText(extractRunsFromPreviewParagraph(p)).length;
    return { p, index, a: null, start, end, text: p.contains(range.endContainer) ? range.toString() : "", link: "" };
  }

  function linkHeadings() {
    const s = documentStructure || (docCanvasEl ? analyzeDocumentDom(docCanvasEl) : null);
    return (s?.headings || []).filter((h) => h.source !== "guess" && Number.isFinite(h.paraIndex));
  }

  // Obrazy jako cele linku „W dokumencie” (Word: Link → Miejsce w tym dokumencie → zakładka przy
  // rysunku; odsyłacz do rysunku). Zakładka trafia do akapitu z obrazem — klik w link skacze do
  // obrazu, a najechanie pokazuje jego podgląd bez skoku (doc-links.js).
  function linkImages() {
    const paras = collectPreviewParagraphElements(host());
    const out = [];
    paras.forEach((p, i) => {
      const img = p.querySelector("img");
      if (!img) return;
      out.push({ paraIndex: i, label: typeof docImageLabel === "function" ? docImageLabel(p, out.length + 1) : t("linkImageN", { n: out.length + 1 }), image: true });
    });
    return out;
  }

  // „#zakładka” → akapit-cel (indeks), żeby okienko pokazało wybrany nagłówek
  function linkTargetParaIndex(link) {
    if (!link.startsWith("#") || typeof linkTargetEl !== "function") return -1;
    const p = linkTargetEl(link.slice(1))?.closest?.("p");
    return p ? resolveParaIndex(p) : -1;
  }

  function normalizeUrl(v) {
    const s = String(v || "").trim();
    if (!s) return "";
    if (LINK_OK_RE.test(s)) return s;
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) return `mailto:${s}`;
    if (/^[\w-]+(\.[\w-]+)+(:\d+)?([/?#].*)?$/.test(s)) return `https://${s}`;
    return "";
  }

  function linkShownTarget(link) {
    if (link.startsWith("#")) {
      const p = typeof linkTargetEl === "function" ? linkTargetEl(link.slice(1))?.closest?.("p") : null;
      const label = (p?.querySelector("img") && typeof docImageLabel === "function" ? docImageLabel(p) : (p?.textContent || "")).replace(/\s+/g, " ").trim().slice(0, 50);
      return label ? t("linkToPlace", { label }) : t("linkTabDoc");
    }
    return link.startsWith("rel:") ? (docLinkHrefs.get(link) || "") : link;
  }

  function openLinkForm(anchor) {
    if (readOnlyMode) return;
    const ctx = linkContext();
    if (!ctx) return;
    hideLinkCard();
    closePop();
    openPop(anchor || insertBtn, (el) => {
      el.classList.add("compose-pop-form");
      el.setAttribute("role", "dialog");
      el.setAttribute("aria-label", t("linkTitle"));
      const heads = [...linkHeadings(), ...linkImages()];
      const targetIdx = linkTargetParaIndex(ctx.link);
      const docMode = ctx.link.startsWith("#");
      const url = ctx.link && !docMode ? linkShownTarget(ctx.link) : "";
      el.innerHTML = `
        <div class="compose-cap"></div>
        <label class="compose-field"><span></span><input type="text" class="lf-text" autocomplete="off"></label>
        <div class="seg lf-mode" role="group">
          <button type="button" data-mode="web"></button><button type="button" data-mode="doc"></button>
        </div>
        <label class="compose-field lf-web"><span></span><input type="url" class="lf-url" inputmode="url" autocomplete="url" autocapitalize="off" spellcheck="false" enterkeyhint="done"></label>
        <label class="compose-field lf-doc"><span></span><select class="lf-target"></select></label>
        <p class="compose-note lf-none" hidden></p>
        <div class="compose-actions">
          <button type="button" class="btn lf-remove" hidden></button>
          <span class="compose-actions-gap"></span>
          <button type="button" class="btn lf-cancel"></button>
          <button type="button" class="btn primary lf-ok"></button>
        </div>`;
      el.querySelector(".compose-cap").textContent = t("linkTitle");
      const [textLbl, webLbl, docLbl] = el.querySelectorAll(".compose-field > span");
      textLbl.textContent = t("linkText");
      webLbl.textContent = t("linkTabWeb");
      docLbl.textContent = t("linkTabDoc");
      const textIn = el.querySelector(".lf-text");
      const urlIn = el.querySelector(".lf-url");
      const sel = el.querySelector(".lf-target");
      textIn.value = ctx.text;
      urlIn.value = url;
      urlIn.placeholder = t("linkUrlPlaceholder");
      const groups = {};
      heads.forEach((h) => {
        const key = h.image ? "linkGroupImages" : "linkGroupHeadings";
        if (!groups[key]) { groups[key] = document.createElement("optgroup"); groups[key].label = t(key); sel.appendChild(groups[key]); }
        const o = document.createElement("option");
        o.value = String(h.paraIndex);
        o.textContent = `${h.image ? "" : "  ".repeat(Math.max(0, (h.level || 1) - 1))}${h.label}`.slice(0, 80);
        if (h.paraIndex === targetIdx) o.selected = true;
        groups[key].appendChild(o);
      });
      const none = el.querySelector(".lf-none");
      none.textContent = t("linkNoHeadings");
      const [webBtn, docBtn] = el.querySelectorAll(".lf-mode button");
      webBtn.textContent = t("linkTabWeb");
      docBtn.textContent = t("linkTabDoc");
      let mode = docMode ? "doc" : "web";
      const setMode = (m) => {
        mode = m;
        webBtn.classList.toggle("is-on", m === "web");
        docBtn.classList.toggle("is-on", m === "doc");
        webBtn.setAttribute("aria-pressed", String(m === "web"));
        docBtn.setAttribute("aria-pressed", String(m === "doc"));
        el.querySelector(".lf-web").hidden = m !== "web";
        el.querySelector(".lf-doc").hidden = m !== "doc" || !heads.length;
        none.hidden = m !== "doc" || !!heads.length;
      };
      webBtn.addEventListener("click", () => { setMode("web"); urlIn.focus(); });
      docBtn.addEventListener("click", () => { setMode("doc"); sel.focus(); });
      setMode(mode);
      (mode === "web" ? urlIn : sel).dataset.autofocus = "1";
      const removeBtn = el.querySelector(".lf-remove");
      removeBtn.textContent = t("linkRemove");
      removeBtn.hidden = !ctx.a;
      removeBtn.addEventListener("click", () => { closePop(); removeLink(ctx); });
      el.querySelector(".lf-cancel").textContent = t("linkCancel");
      el.querySelector(".lf-cancel").addEventListener("click", () => closePop());
      const ok = el.querySelector(".lf-ok");
      ok.textContent = ctx.a ? t("linkSave") : t("linkInsert");
      const submit = () => {
        const edit = { op: "link", index: ctx.index, start: ctx.start, end: ctx.end };
        let label = "";
        if (mode === "doc") {
          const idx = parseInt(sel.value, 10);
          if (!heads.length || !Number.isInteger(idx)) { toast(t("linkNoHeadings"), "info"); return; }
          edit.targetIndex = idx;
          label = heads.find((h) => h.paraIndex === idx)?.label || "";
        } else {
          const href = normalizeUrl(urlIn.value);
          if (!href) { toast(t("linkBadUrl"), "warning"); urlIn.focus(); return; }
          edit.href = href;
          label = urlIn.value.trim();
        }
        let text = textIn.value;
        if (!text.trim()) text = ctx.text || label;
        if (text !== ctx.text) edit.text = text;
        closePop();
        runFileEdit(edit, { paraIndex: ctx.index, offset: ctx.start + text.length });
      };
      ok.addEventListener("click", submit);
      el.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && e.target.closest("input")) { e.preventDefault(); submit(); }
      });
    });
  }

  function removeLink(ctx) {
    if (!ctx?.a) return;
    runFileEdit({ op: "link", index: ctx.index, start: ctx.start, end: ctx.end, remove: true }, { paraIndex: ctx.index, offset: ctx.end });
  }

  function openLink(link, href) {
    if (link.startsWith("#")) { if (typeof jumpToLinkTarget === "function") jumpToLinkTarget(link.slice(1)); return; }
    if (LINK_OK_RE.test(href)) { window.open(href, "_blank", "noopener,noreferrer"); return; }
    toast(t("linkUnsupported"), "info");
  }

  // Karta linku: tylko gdy kursor stoi w linku (Edycja). Nie zabiera fokusu z tekstu.
  let card = null; // { el, a }
  function hideLinkCard() {
    if (!card) return;
    card.el.remove();
    card = null;
  }
  function showLinkCard(a) {
    if (card?.a === a) { placeLinkCard(); return; }
    hideLinkCard();
    const link = a.dataset.dwbLink || a.getAttribute("href") || "";
    const href = link.startsWith("#") ? link : (a.getAttribute("href") || "");
    const el = document.createElement("div");
    el.className = "link-card";
    el.setAttribute("role", "toolbar");
    el.setAttribute("aria-label", t("linkTitle"));
    el.innerHTML = `<span class="link-card-target"></span><button type="button" class="tb-btn lc-open"></button><button type="button" class="tb-btn lc-edit"></button><button type="button" class="tb-btn lc-remove"></button>`;
    el.querySelector(".link-card-target").textContent = linkShownTarget(link) || href;
    const btn = (cls, key, svg, fn) => {
      const b = el.querySelector(cls);
      b.setAttribute("aria-label", t(key));
      b.dataset.hint = "";
      b.dataset.hintPl = I18N.pl[key];
      b.dataset.hintEn = I18N.en[key];
      b.dataset.hintDelay = "0.4";
      b.innerHTML = ICON(svg);
      b.addEventListener("click", fn);
    };
    btn(".lc-open", "linkOpen", '<path d="M14 3h7v7"/><path d="M10 14 21 3"/><path d="M21 14v5a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5"/>', () => openLink(link, href));
    btn(".lc-edit", "linkEdit", '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>', () => openLinkForm(insertBtn));
    btn(".lc-remove", "linkRemove", '<path d="M18.84 12.25l1.72-1.71a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M5.17 11.75l-1.71 1.71a5 5 0 0 0 7.07 7.07l1.71-1.71"/><line x1="8" y1="2" x2="8" y2="5"/><line x1="2" y1="8" x2="5" y2="8"/><line x1="16" y1="19" x2="16" y2="22"/><line x1="19" y1="16" x2="22" y2="16"/>', () => {
      const ctx = linkContext();
      hideLinkCard();
      removeLink(ctx);
    });
    el.addEventListener("mousedown", (e) => { if (e.target.closest("button")) e.preventDefault(); });
    document.body.appendChild(el);
    card = { el, a };
    placeLinkCard();
  }
  function placeLinkCard() {
    if (!card) return;
    const vv = window.visualViewport;
    const rects = card.a.getClientRects();
    const r = rects[rects.length - 1] || card.a.getBoundingClientRect();
    const vp = docViewportEl?.getBoundingClientRect();
    if (vp && (r.bottom < vp.top || r.top > vp.bottom)) { card.el.style.visibility = "hidden"; return; }
    card.el.style.visibility = "";
    const viewW = vv ? vv.width : window.innerWidth;
    const w = card.el.offsetWidth;
    const left = Math.max(8, Math.min(r.left, viewW - w - 8));
    card.el.style.left = `${left + (vv ? vv.offsetLeft : 0)}px`;
    card.el.style.top = `${r.bottom + 6 + (vv ? vv.offsetTop : 0)}px`;
  }

  // ── tabele ────────────────────────────────────────────────────────────────
  const host = () => docCanvasEl?.querySelector(".docx-preview-host");

  // Czeka, aż po przerysowaniu akapity będą edytowalne (oznaczenia „tylko do odczytu” gotowe).
  function whenEditable() {
    return new Promise((resolve) => {
      const t0 = performance.now();
      const tick = () => (!inlineLocksPending || performance.now() - t0 > 4000 ? resolve() : requestAnimationFrame(tick));
      tick();
    });
  }
  function firstParaIndexIn(el) {
    const p = el?.querySelector?.("p");
    return p ? resolveParaIndex(p) : -1;
  }

  // Siatka rozmiaru jak w Wordzie: najechanie podświetla, klik/tap wstawia. Klawiatura: strzałki + Enter.
  const PICK_MAX = 8;
  function buildTablePicker(el) {
    el.classList.add("compose-pop-table");
    popCap(el, t("insertTable"));
    const grid = document.createElement("div");
    grid.className = "table-pick";
    grid.setAttribute("role", "grid");
    const label = document.createElement("div");
    label.className = "table-pick-size";
    const mark = (r, c) => {
      grid.querySelectorAll("button").forEach((b) => b.classList.toggle("on", +b.dataset.r <= r && +b.dataset.c <= c));
      label.textContent = r && c ? `${c} × ${r}` : t("insertTablePick");
    };
    for (let r = 1; r <= PICK_MAX; r++) {
      for (let c = 1; c <= PICK_MAX; c++) {
        const b = document.createElement("button");
        b.type = "button";
        b.dataset.r = r;
        b.dataset.c = c;
        b.setAttribute("aria-label", t("insertTableCells", { cols: c, rows: r }));
        b.addEventListener("mouseenter", () => mark(r, c));
        b.addEventListener("focus", () => mark(r, c));
        b.addEventListener("click", () => { closePop(); insertTable(r, c); });
        grid.appendChild(b);
      }
    }
    grid.addEventListener("keydown", (e) => {
      const b = e.target.closest("button");
      if (!b) return;
      let r = +b.dataset.r; let c = +b.dataset.c;
      if (e.key === "ArrowRight") c = Math.min(PICK_MAX, c + 1);
      else if (e.key === "ArrowLeft") c = Math.max(1, c - 1);
      else if (e.key === "ArrowDown") r = Math.min(PICK_MAX, r + 1);
      else if (e.key === "ArrowUp") r = Math.max(1, r - 1);
      else return;
      e.preventDefault();
      e.stopPropagation();
      grid.querySelector(`button[data-r="${r}"][data-c="${c}"]`)?.focus();
    }, true);
    el.appendChild(grid);
    el.appendChild(label);
    mark(0, 0);
    // klawiatura: fokus od razu w siatce; dotyk: bez podświetlonego „1 × 1” przed wyborem
    const first = grid.querySelector("button");
    if (first && !window.matchMedia?.("(pointer: coarse)").matches) first.dataset.autofocus = "1";
  }

  async function insertTable(rows, cols) {
    const p = caretParagraph();
    if (!p) return;
    const index = resolveParaIndex(p);
    if (index < 0) return;
    const empty = !previewRunsToPlainText(extractRunsFromPreviewParagraph(p)).trim();
    // kursor w pierwszej komórce: tabela przed pustym akapitem albo za akapitem z tekstem
    await runFileEdit({ op: "tableInsert", index, rows, cols }, { paraIndex: empty ? index : index + 1, offset: 0 });
  }

  function cellOf(p) {
    const td = p?.closest?.("td, th");
    const tr = td?.parentElement;
    const table = tr?.closest("table");
    if (!td || !table) return null;
    const tables = Array.from(host()?.querySelectorAll("table") || []);
    return { td, tr, table, ti: tables.indexOf(table), r: Array.from(table.rows).indexOf(tr), c: Array.from(tr.cells).indexOf(td) };
  }

  // Po przerysowaniu: kursor do komórki (tabela nr ti, wiersz r, kolumna c) — granice przycięte.
  function focusCell(ti, r, c) {
    const table = host()?.querySelectorAll("table")[ti];
    if (!table) return false;
    const row = table.rows[Math.max(0, Math.min(table.rows.length - 1, r))];
    const cell = row?.cells[Math.max(0, Math.min(row.cells.length - 1, c))];
    const idx = firstParaIndexIn(cell);
    if (idx < 0) return false;
    focusParagraphAtOffset(idx, 0);
    return true;
  }

  async function tableAction(action) {
    const p = caretParagraph();
    const cell = cellOf(p);
    if (!cell) { toast(t("tableNoCaret"), "info"); return; }
    const index = resolveParaIndex(p);
    const firstIdx = firstParaIndexIn(cell.table);
    await applyDocumentEdit({ op: "table", index, action }).catch((err) => log(`Tabela: ${err.message || err}`, "error"));
    await whenEditable();
    if (readOnlyMode) return;
    const { ti, r, c } = cell;
    const target = { rowAbove: [r, c], rowBelow: [r + 1, c], colLeft: [r, c], colRight: [r, c + 1], delRow: [r, c], delCol: [r, c - 1] }[action];
    if (target) focusCell(ti, target[0], Math.max(0, target[1]));
    else if (action === "delTable" && firstIdx >= 0) focusParagraphAtOffset(firstIdx, 0); // akapit, który był pod tabelą
  }

  function buildTableMenu(el) {
    el.classList.add("compose-pop-tabletools");
    popCap(el, t("tableInsertGroup"));
    const grid = document.createElement("div");
    grid.className = "compose-grid2";
    [
      ["tableRowAbove", "rowAbove", '<rect x="3" y="12" width="18" height="8" rx="1"/><line x1="12" y1="3" x2="12" y2="9"/><line x1="9" y1="6" x2="15" y2="6"/>'],
      ["tableRowBelow", "rowBelow", '<rect x="3" y="4" width="18" height="8" rx="1"/><line x1="12" y1="15" x2="12" y2="21"/><line x1="9" y1="18" x2="15" y2="18"/>'],
      ["tableColLeft", "colLeft", '<rect x="12" y="3" width="8" height="18" rx="1"/><line x1="3" y1="12" x2="9" y2="12"/><line x1="6" y1="9" x2="6" y2="15"/>'],
      ["tableColRight", "colRight", '<rect x="4" y="3" width="8" height="18" rx="1"/><line x1="15" y1="12" x2="21" y2="12"/><line x1="18" y1="9" x2="18" y2="15"/>'],
    ].forEach(([key, action, svg]) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "compose-item compose-item-sm";
      b.setAttribute("role", "menuitem");
      b.innerHTML = `<span class="compose-item-icon" aria-hidden="true">${ICON(svg)}</span><span class="compose-item-label"></span>`;
      b.querySelector(".compose-item-label").textContent = t(key);
      b.addEventListener("click", () => { closePop(); tableAction(action); });
      grid.appendChild(b);
    });
    el.appendChild(grid);
    popCap(el, t("tableDeleteGroup"));
    [["tableDelRow", "delRow"], ["tableDelCol", "delCol"], ["tableDelTable", "delTable"]].forEach(([key, action]) => {
      popItem(el, { label: t(key), icon: ICON('<polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>'), onPick: () => tableAction(action) });
    });
    buildCellAlign(el);
    const hint = document.createElement("p");
    hint.className = "compose-note compose-note-pad";
    hint.textContent = t("tableTabHint");
    el.appendChild(hint);
  }

  // ── wyrównanie w komórce (Word: Układ tabeli → Wyrównanie): siatka 3×3 + zakres ──────────
  const CELL_V = ["top", "center", "bottom"];
  const CELL_H = ["left", "center", "right"];
  function cellAlignIcon(h, v) {
    const y = { top: [6, 10], center: [10, 14], bottom: [14, 18] }[v];
    const line = (len, yy) => { const x0 = h === "left" ? 5 : h === "right" ? 19 - len : 12 - len / 2; return `<line x1="${x0}" y1="${yy}" x2="${x0 + len}" y2="${yy}"/>`; };
    return ICON(`<rect x="2.5" y="2.5" width="19" height="19" rx="2" stroke-width="1.4" opacity=".55"/>${line(10, y[0])}${line(6, y[1])}`);
  }
  function currentCellAlign() {
    const p = docCaretParagraph(document.activeElement) || (typeof lastDocCaret !== "undefined" ? lastDocCaret?.p : null);
    const td = p?.closest?.("td, th");
    if (!td) return {};
    const va = getComputedStyle(td).verticalAlign;
    const ta = getComputedStyle(p).textAlign;
    return { v: va === "middle" ? "center" : va === "bottom" ? "bottom" : "top", h: ta === "center" ? "center" : ta === "right" || ta === "end" ? "right" : "left" };
  }
  function buildCellAlign(el) {
    popCap(el, t("tableAlignGroup"));
    let scope = "cell";
    const seg = document.createElement("div");
    seg.className = "seg compose-cellalign-scope";
    seg.setAttribute("role", "group");
    seg.setAttribute("aria-label", t("tableAlignScope"));
    [["cell", "tableScopeCell"], ["row", "tableScopeRow"], ["col", "tableScopeCol"], ["table", "tableScopeTable"]].forEach(([v, key]) => {
      const b = document.createElement("button");
      b.type = "button";
      b.dataset.v = v;
      b.textContent = t(key);
      b.classList.toggle("is-on", v === scope);
      b.setAttribute("aria-pressed", String(v === scope));
      b.addEventListener("click", () => {
        scope = v;
        seg.querySelectorAll("button").forEach((x) => { x.classList.toggle("is-on", x === b); x.setAttribute("aria-pressed", String(x === b)); });
      });
      seg.appendChild(b);
    });
    el.appendChild(seg);
    const cur = currentCellAlign();
    const grid = document.createElement("div");
    grid.className = "compose-cellalign";
    grid.setAttribute("role", "group");
    CELL_V.forEach((v) => CELL_H.forEach((h) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "tb-btn compose-align";
      const on = cur.v === v && cur.h === h;
      b.classList.toggle("is-on", on);
      b.setAttribute("aria-pressed", String(on));
      const label = `${t(`cellAlignV_${v}`)}, ${t(`cellAlignH_${h}`)}`;
      b.setAttribute("aria-label", label);
      b.dataset.hint = label;
      b.dataset.hintDelay = "0.3";
      b.innerHTML = cellAlignIcon(h, v);
      b.addEventListener("click", () => { closePop(); tableAlign(h, v, scope); });
      grid.appendChild(b);
    }));
    el.appendChild(grid);
  }
  async function tableAlign(h, v, scope) {
    const p = caretParagraph();
    const cell = cellOf(p);
    if (!cell) { toast(t("tableNoCaret"), "info"); return; }
    const index = resolveParaIndex(p);
    await applyDocumentEdit({ op: "table", index, action: "align", h, v, scope }).catch((err) => log(`Tabela: ${err.message || err}`, "error"));
    await whenEditable();
    if (!readOnlyMode) focusCell(cell.ti, cell.r, cell.c);
  }

  // ── wyrównanie strony w pionie (Word: Ustawienia strony → Układ → Wyrównanie w pionie) ───
  const PAGE_VALIGN = [
    ["top", "pageVAlignTop", '<rect x="5" y="2.5" width="14" height="19" rx="1.5" opacity=".55"/><line x1="8" y1="6" x2="16" y2="6"/><line x1="8" y1="9" x2="14" y2="9"/>'],
    ["center", "pageVAlignCenter", '<rect x="5" y="2.5" width="14" height="19" rx="1.5" opacity=".55"/><line x1="8" y1="10.5" x2="16" y2="10.5"/><line x1="8" y1="13.5" x2="14" y2="13.5"/>'],
    ["both", "pageVAlignBoth", '<rect x="5" y="2.5" width="14" height="19" rx="1.5" opacity=".55"/><line x1="8" y1="6" x2="16" y2="6"/><line x1="8" y1="12" x2="16" y2="12"/><line x1="8" y1="18" x2="14" y2="18"/>'],
    ["bottom", "pageVAlignBottom", '<rect x="5" y="2.5" width="14" height="19" rx="1.5" opacity=".55"/><line x1="8" y1="15" x2="16" y2="15"/><line x1="8" y1="18" x2="14" y2="18"/>'],
  ];
  function buildPageVAlignMenu(el) {
    popCap(el, t("pageVAlign"));
    const p = typeof restoreDocCaret === "function" ? restoreDocCaret() : null;
    const index = p ? resolveParaIndex(p) : 0;
    el.classList.add("compose-pop-valign");
    const items = PAGE_VALIGN.map(([val, key, svg]) => {
      const b = popItem(el, { label: t(key), icon: ICON(svg), onPick: () => runFileEdit({ op: "pageVAlign", index: Math.max(0, index), val }, p ? { paraIndex: index, offset: 0 } : null) });
      b.setAttribute("role", "menuitemradio");
      b.dataset.v = val;
      return b;
    });
    const note = document.createElement("p");
    note.className = "compose-note compose-note-pad";
    note.textContent = t("pageVAlignNote");
    el.appendChild(note);
    // zaznaczenie obecnego ustawienia sekcji (z pliku)
    getDocumentXmlDom(originalFileBytes).then((doc) => {
      const cur = doc ? composeSectionVAlign(doc, index) : "top";
      items.forEach((b) => { b.setAttribute("aria-checked", String(b.dataset.v === cur)); b.classList.toggle("is-current", b.dataset.v === cur); });
    }).catch(() => {});
  }

  // ── przypisy: Wstaw → Przypis dolny / końcowy (jak Word: Odwołania → Wstaw przypis) ──────
  // Odnośnik w miejscu kursora (przy zaznaczeniu — za nim), numer jak w Wordzie, kursor od razu
  // w tekście nowego przypisu na dole strony / na końcu dokumentu.
  function modelOffsetAt(p, node, offset) {
    const pre = document.createRange();
    pre.selectNodeContents(p);
    pre.setEnd(node, offset);
    const tmp = document.createElement("p");
    tmp.appendChild(pre.cloneContents());
    return previewRunsToPlainText(extractRunsFromPreviewParagraph(tmp)).length;
  }
  async function insertNote(kind) {
    const p = caretParagraph();
    if (!p) return;
    if (p.dataset.noteKey) { toast(t("noteInNote"), "info"); return; }
    if (p.dataset.lock) { toast(t(p.dataset.lock), "info"); return; }
    const index = resolveParaIndex(p);
    if (index < 0) return;
    const r = window.getSelection().getRangeAt(0);
    const offset = p.contains(r.endContainer) ? modelOffsetAt(p, r.endContainer, r.endOffset) : 0;
    let id = 1;
    try {
      await mergeInlineEditsIntoBytes();
      const zip = await loadDocxZipCached(originalFileBytes);
      id = noteNextId(await zip.file(NOTE_PART[kind])?.async("string"), kind);
    } catch (_) { /* id policzy zapis */ }
    await runFileEdit({ op: "noteInsert", kind, index, offset, id }, { paraIndex: index, offset });
    await whenEditable();
    const li = host()?.querySelector(`ol.dwb-notes > li[data-dwb-note="${kind}:${id}"]`);
    const np = li && [...li.children].filter((c) => c.tagName === "P" && c.isContentEditable).pop();
    if (np) {
      np.scrollIntoView({ block: "center" });
      placeCaret(np, Number.MAX_SAFE_INTEGER);
    }
  }

  // ── układ strony: marginesy, orientacja, rozmiar (Word: Układ → Marginesy / Orientacja / Rozmiar) ──
  // Gotowe marginesy jak w polskim Wordzie (twipy; 1 cm = 567). Zmiana dotyczy sekcji z kursorem
  // (w dokumencie z jedną sekcją = całego), „Marginesy niestandardowe…” pozwala wybrać zakres.
  const MARGIN_PRESETS = [
    ["normal", "marginNormal", { top: 1418, bottom: 1418, left: 1418, right: 1418 }],
    ["narrow", "marginNarrow", { top: 720, bottom: 720, left: 720, right: 720 }],
    ["moderate", "marginModerate", { top: 1440, bottom: 1440, left: 1080, right: 1080 }],
    ["wide", "marginWide", { top: 1440, bottom: 1440, left: 2880, right: 2880 }],
  ];
  const PAPER_SIZES = [["A4", 11906, 16838], ["A5", 8391, 11906], ["Letter", 12240, 15840], ["Legal", 12240, 20160], ["A3", 16838, 23811]];
  const TW_PER_CM = 1440 / 2.54;
  // martwa strefa typowej drukarki (¼ cala) — margines mniejszy = ostrzeżenie (jak Word)
  const PRINTER_MIN_TW = 360;
  const cmText = (tw) => (Math.round(tw / TW_PER_CM * 100) / 100).toLocaleString(currentLang === "en" ? "en-GB" : "pl-PL", { maximumFractionDigits: 2 });
  const sameTw = (a, b) => Math.abs(a - b) <= 6;

  async function pageSetupAtCaret() {
    const p = typeof restoreDocCaret === "function" ? restoreDocCaret() : null;
    const index = p ? Math.max(0, resolveParaIndex(p)) : 0;
    const doc = await getDocumentXmlDom(originalFileBytes);
    return { p, index, cur: doc ? composeSectionPageSetup(doc, index) : null };
  }

  // Wejścia: „＋ Wstaw → Układ strony…”, menu ⋯ i przycisk „Marginesy” w Podglądzie wydruku
  // (openPageSetup — w trybie Czytanie przełącza na Edycję; after = np. przerysuj podgląd).
  let pageSetupAnchor = null;
  let pageSetupAfter = null;
  async function openPageSetup(anchor, after = null) {
    if (!originalFileBytes) return;
    if (readOnlyMode && typeof appFrame !== "undefined") {
      appFrame.setReadOnly(false);
      await whenEditable();
      // przełączenie trybu poprawia przewinięcie (tekst zostaje w miejscu) — zdarzenie „scroll”
      // przychodzi klatkę później i zamknęłoby świeżo otwarte okienko
      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    }
    pageSetupAnchor = anchor || (layoutBtn && !layoutBtn.closest(".hidden") ? layoutBtn : insertBtn);
    pageSetupAfter = after;
    openPop(pageSetupAnchor, buildPageSetupMenu);
  }
  async function runPageSetup(edit, ctx) {
    pageSetupReflowHint();
    await runFileEdit({ op: "pageSetup", index: ctx.index, scope: "section", ...edit }, ctx.p ? caretState(ctx.p) : null);
    const after = pageSetupAfter;
    pageSetupAfter = null;
    after?.();
  }
  // Widok mobilny przekłada tekst na szerokość ekranu — kartek i marginesów tam nie widać
  function pageSetupReflowHint() {
    if (typeof dwbPrint !== "undefined" && dwbPrint.isOpen()) return; // w podglądzie wydruku widać od razu
    if (typeof shouldUseMobileReflow === "function" && shouldUseMobileReflow()) toast(t("pageSetupMobileHint"), "info");
  }

  // Czy marginesy zostawiają miejsce na tekst (Word odmawia, gdy kolumna byłaby za wąska).
  function marginsFit(m, cur) {
    return cur.w - m.left - m.right >= 1440 && cur.h - m.top - m.bottom >= 1440;
  }

  async function buildPageSetupMenu(el) {
    el.classList.add("compose-pop-insert");
    await mergeInlineEditsIntoBytes().catch(() => {});
    const ctx = await pageSetupAtCaret();
    const cur = ctx.cur || { w: 11906, h: 16838, orient: "portrait", top: 1418, bottom: 1418, left: 1418, right: 1418, sections: 1 };
    const mark = (b, on) => { b.setAttribute("role", "menuitemradio"); b.setAttribute("aria-checked", String(on)); b.classList.toggle("is-current", on); };
    popCap(el, t("pageMargins"));
    MARGIN_PRESETS.forEach(([key, label, m]) => {
      const desc = m.left === m.top ? t("marginAll", { v: cmText(m.top) }) : t("marginTBLR", { tb: cmText(m.top), lr: cmText(m.left) });
      const b = popItem(el, {
        label: t(label), desc,
        icon: ICON(`<rect x="4" y="2.5" width="16" height="19" rx="1.5"/><rect x="${4 + m.left / 720}" y="${2.5 + m.top / 720}" width="${16 - (m.left + m.right) / 720}" height="${19 - (m.top + m.bottom) / 720}" rx=".4" stroke-dasharray="1.6 1.4"/>`),
        onPick: () => {
          if (!marginsFit(m, cur)) { toast(t("marginTooBig"), "warning"); return; }
          runPageSetup({ margins: m }, ctx);
        },
      });
      b.dataset.preset = key;
      mark(b, ["top", "bottom", "left", "right"].every((k) => sameTw(m[k], cur[k])));
    });
    const custom = popItem(el, {
      label: t("marginCustom"),
      desc: t("marginCurrent", { t: cmText(cur.top), b: cmText(cur.bottom), l: cmText(cur.left), r: cmText(cur.right) }),
      icon: ICON('<rect x="4" y="2.5" width="16" height="19" rx="1.5"/><path d="M8 8h8M8 12h5M8 16h6"/>'),
      onPick: () => openPop(pageSetupAnchor || insertBtn, (form) => buildMarginForm(form, ctx, cur)),
    });
    custom.dataset.preset = "custom";
    popCap(el, t("pageOrient"));
    [["portrait", "orientPortrait", '<rect x="6" y="2.5" width="12" height="19" rx="1.5"/>'], ["landscape", "orientLandscape", '<rect x="2.5" y="6" width="19" height="12" rx="1.5"/>']].forEach(([val, key, svg]) => {
      const b = popItem(el, { label: t(key), icon: ICON(svg), onPick: () => runPageSetup({ orient: val }, ctx) });
      b.dataset.orient = val;
      mark(b, cur.orient === val);
    });
    popCap(el, t("pageSize"));
    PAPER_SIZES.forEach(([name, w, h]) => {
      const b = popItem(el, {
        label: name, desc: `${cmText(w)} × ${cmText(h)} cm`,
        icon: ICON('<path d="M7 2.5h7l4 4v15H7z"/><path d="M14 2.5v4h4"/>'),
        onPick: () => runPageSetup({ size: { w, h } }, ctx),
      });
      b.dataset.size = name;
      mark(b, sameTw(Math.min(cur.w, cur.h), w) && sameTw(Math.max(cur.w, cur.h), h));
    });
    // Word: Ustawienia strony → Układ → Wyrównanie w pionie — w tym samym okienku (dawniej w „Wstaw”)
    buildPageVAlignMenu(el);
  }

  function buildMarginForm(el, ctx, cur) {
    el.classList.add("compose-pop-form", "compose-pop-margins");
    el.innerHTML = `<div class="compose-cap"></div>
      <div class="mf-grid">
        ${["top", "bottom", "left", "right"].map((k) => `<label class="compose-field"><span></span><input type="text" class="mf-${k}" inputmode="decimal" autocomplete="off" enterkeyhint="done"></label>`).join("")}
      </div>
      <label class="compose-field mf-scope-field"><span></span><select class="mf-scope"></select></label>
      <p class="compose-note mf-warn" hidden></p>
      <div class="compose-actions"><span class="compose-actions-gap"></span><button type="button" class="btn lf-cancel"></button><button type="button" class="btn primary lf-ok"></button></div>`;
    el.querySelector(".compose-cap").textContent = t("marginCustomTitle");
    const keys = ["top", "bottom", "left", "right"];
    const labels = el.querySelectorAll(".mf-grid .compose-field > span");
    keys.forEach((k, i) => {
      labels[i].textContent = t(`margin_${k}`);
      const inp = el.querySelector(`.mf-${k}`);
      inp.value = cmText(cur[k]); // przecinek po polsku; przyjmujemy też kropkę
      if (i === 0) inp.dataset.autofocus = "1";
    });
    el.querySelector(".mf-scope-field > span").textContent = t("marginScope");
    const scope = el.querySelector(".mf-scope");
    scope.appendChild(new Option(t("marginScopeAll"), "all"));
    if (cur.sections > 1) scope.appendChild(new Option(t("marginScopeSection"), "section"));
    el.querySelector(".mf-scope-field").hidden = cur.sections <= 1;
    const warn = el.querySelector(".mf-warn");
    const read = () => Object.fromEntries(keys.map((k) => [k, Math.round(parseFloat(String(el.querySelector(`.mf-${k}`).value).replace(",", ".")) * TW_PER_CM)]));
    const check = () => {
      const m = read();
      const bad = keys.some((k) => !Number.isFinite(m[k]) || m[k] < 0);
      const tight = !bad && keys.some((k) => m[k] < PRINTER_MIN_TW);
      warn.hidden = !(bad || tight || !marginsFit(m, cur));
      warn.textContent = bad ? t("marginBad") : !marginsFit(m, cur) ? t("marginTooBig") : t("marginPrinterWarn");
      warn.classList.toggle("is-error", bad || !marginsFit(m, cur));
      return !bad && marginsFit(m, cur);
    };
    el.addEventListener("input", check);
    check();
    el.querySelector(".lf-cancel").textContent = t("linkCancel");
    el.querySelector(".lf-cancel").addEventListener("click", () => closePop());
    const ok = el.querySelector(".lf-ok");
    ok.textContent = t("marginApply");
    const submit = () => {
      if (!check()) return;
      closePop();
      pageSetupReflowHint();
      const after = pageSetupAfter;
      pageSetupAfter = null;
      runFileEdit({ op: "pageSetup", index: ctx.index, scope: scope.value || "all", margins: read() }, ctx.p ? caretState(ctx.p) : null).then(() => after?.());
    };
    ok.addEventListener("click", submit);
    el.addEventListener("keydown", (e) => { if (e.key === "Enter" && e.target.closest("input")) { e.preventDefault(); submit(); } });
  }

  // Tab / Shift+Tab w komórce = następna / poprzednia komórka; Tab w ostatniej = nowy wiersz (jak Word).
  function tableTab(p, back) {
    const cell = cellOf(p);
    if (!cell) return false;
    const cells = Array.from(cell.table.querySelectorAll("td, th")).filter((td) => td.closest("table") === cell.table);
    const i = cells.indexOf(cell.td);
    const next = cells[i + (back ? -1 : 1)];
    if (next) {
      const idx = firstParaIndexIn(next);
      if (idx >= 0) {
        const el = collectPreviewParagraphElements(host())[idx];
        focusParagraphAtOffset(idx, back ? previewRunsToPlainText(extractRunsFromPreviewParagraph(el)).length : 0);
      }
      return true;
    }
    if (back) return true;
    tableAction("rowBelow").then(() => focusCell(cell.ti, cell.r + 1, 0));
    return true;
  }

  // ── obrazy ────────────────────────────────────────────────────────────────
  const IMG_MAX_PX = 2400;
  const IMG_KEEP_BYTES = 2.5 * 1024 * 1024;

  // Obraz → bajty do .docx: PNG/JPEG/GIF zostają (gdy nie za duże), resztę (WebP, HEIC…) i zbyt
  // duże zdjęcia przerabiamy w przeglądarce (canvas) — Word nie zna WebP, a 12 MP z aparatu
  // niepotrzebnie puchnie plik.
  async function prepareImage(file) {
    const type = (file.type || "").toLowerCase();
    let bitmap;
    try { bitmap = await createImageBitmap(file); } catch (_) {
      bitmap = await new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = reject;
        img.src = URL.createObjectURL(file);
      });
    }
    const w0 = bitmap.width; const h0 = bitmap.height;
    const keep = { "image/png": "png", "image/jpeg": "jpeg", "image/jpg": "jpeg", "image/gif": "gif" }[type];
    if (keep && file.size <= IMG_KEEP_BYTES && Math.max(w0, h0) <= IMG_MAX_PX) {
      return { bytes: new Uint8Array(await file.arrayBuffer()), ext: keep === "jpeg" ? "jpeg" : keep, mime: keep === "gif" ? "image/gif" : `image/${keep}`, width: w0, height: h0 };
    }
    const k = Math.min(1, IMG_MAX_PX / Math.max(w0, h0));
    const w = Math.max(1, Math.round(w0 * k)); const h = Math.max(1, Math.round(h0 * k));
    const canvas = document.createElement("canvas");
    canvas.width = w; canvas.height = h;
    canvas.getContext("2d").drawImage(bitmap, 0, 0, w, h);
    const png = type === "image/png" || type === "image/gif" || type === "image/webp" || type === "image/svg+xml";
    const mime = png ? "image/png" : "image/jpeg";
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, mime, 0.86));
    return { bytes: new Uint8Array(await blob.arrayBuffer()), ext: png ? "png" : "jpeg", mime, width: w, height: h };
  }

  async function insertImageFile(file, p) {
    if (!file || !/^image\//.test(file.type || "")) { toast(t("imageBadType"), "warning"); return; }
    p = p || caretParagraph();
    if (!p) return;
    const index = resolveParaIndex(p);
    if (index < 0) return;
    let img;
    try { img = await prepareImage(file); } catch (err) {
      log(`Obraz: ${err.message || err}`, "error");
      toast(t("imageBadType"), "warning");
      return;
    }
    const empty = !previewRunsToPlainText(extractRunsFromPreviewParagraph(p)).trim();
    const imgIdx = empty ? index : index + 1;
    await runFileEdit({ op: "imageInsert", index, ...img, name: (file.name || "").replace(/\.[^.]+$/, "") }, { paraIndex: imgIdx + 1, offset: 0 });
  }

  function pickImage() {
    const p = caretParagraph(); // zapamiętane, zanim okno wyboru zabierze fokus
    if (!p) return;
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.addEventListener("change", () => { if (input.files?.[0]) insertImageFile(input.files[0], p); }, { once: true });
    input.click();
  }

  // Karta obrazu: klik w obraz w Edycji — rozmiar (suwak, % szerokości tekstu), wyrównanie, opis, usuń.
  let imgCard = null; // { el, img, p }
  function hideImageCard() {
    if (!imgCard) return;
    imgCard.img.classList.remove("img-selected");
    imgCard.el.remove();
    imgCard = null;
  }
  function imagePct(img, p) {
    const w = p.clientWidth - parseFloat(getComputedStyle(p).paddingLeft || 0) - parseFloat(getComputedStyle(p).paddingRight || 0);
    return Math.max(5, Math.min(100, Math.round((img.getBoundingClientRect().width / (w || 1)) * 100 / (docZoomScale(img) || 1))));
  }
  // Skala dokumentu na ekranie (zoom, transform Widoku desktopowego, gest) — mierzona wprost:
  // szerokość akapitu na ekranie / jego szerokość w układzie.
  function docZoomScale(el) {
    const p = el?.closest?.("p") || docCanvasEl?.querySelector(".docx-preview-host section.docx article p");
    const w = p?.offsetWidth;
    return w ? p.getBoundingClientRect().width / w || 1 : 1;
  }
  // Szerokość tekstu akapitu (układ, bez zoomu) — zapas, zanim przyjdzie szerokość z pliku.
  function imageTextWidth(p) {
    const cs = getComputedStyle(p);
    return p.clientWidth - parseFloat(cs.paddingLeft || 0) - parseFloat(cs.paddingRight || 0);
  }
  // 100% suwaka = szerokość tekstu STRONY z pliku (papier − marginesy sekcji), tak jak liczy zapis
  // (composeTextWidthTwips) i Word. W Widoku mobilnym akapit ma szerokość ekranu, więc pomiar
  // z ekranu dawał inny procent niż plik: „25%” pokazywało obraz na ~44% ekranu, a po puszczeniu
  // suwaka obraz „dociągał” (sprawdzone na symulatorze iPhone'a).
  async function pageTextPx(index) {
    if (!originalFileBytes || typeof composeTextWidthTwips !== "function") return 0;
    const doc = await getDocumentXmlDom(originalFileBytes).catch(() => null);
    const p = doc ? collectParagraphElements(doc.documentElement, "all")[index] : null;
    return p ? composeTextWidthTwips(doc, p) / 15 : 0; // 1440 tw = 96 px
  }
  // Szerokość obrazu w dokumencie (z pliku: docx-preview daje ją w pt), nie przycięta do ekranu.
  function imageDocWidthPx(img) {
    const w = img.style.width || "";
    const n = parseFloat(w);
    if (!n) return img.offsetWidth;
    return /pt$/.test(w) ? n * 96 / 72 : /cm$/.test(w) ? n * 96 / 2.54 : n;
  }
  function showImageCard(img) {
    const p = img.closest("p");
    if (!p) return;
    const index = resolveParaIndex(p);
    if (index < 0) return;
    hideImageCard();
    hideLinkCard();
    img.classList.add("img-selected");
    const el = document.createElement("div");
    el.className = "image-card";
    el.setAttribute("role", "toolbar");
    el.setAttribute("aria-label", t("imageTools"));
    const pct0 = imagePct(img, p);
    el.innerHTML = `<label class="image-size"><span class="image-size-val"></span><input type="range" min="10" max="100" step="5"></label>
      <span class="tb-sep" aria-hidden="true"></span>
      <button type="button" class="tb-btn" data-align="left"></button><button type="button" class="tb-btn" data-align="center"></button><button type="button" class="tb-btn" data-align="right"></button>
      <span class="tb-sep" aria-hidden="true"></span>
      <button type="button" class="tb-btn ic-alt"></button><button type="button" class="tb-btn ic-del"></button>`;
    const range = el.querySelector("input");
    const val = el.querySelector(".image-size-val");
    range.value = String(Math.round(pct0 / 5) * 5);
    range.setAttribute("aria-label", t("imageSize"));
    val.textContent = `${range.value}%`;
    // Suwak (zgłoszenie: „miga, skacze”): karta stała w miejscu, dopóki trzymasz suwak — dawniej
    // jechała za zmieniającym się obrazem, suwak uciekał spod palca/myszy i sam zmieniał wartość
    // (sprzężenie zwrotne). Podgląd na żywo liczony z szerokości tekstu (tak jak plik), więc po
    // zapisie obraz nie „dociąga” o kilka pikseli. Zapis: po puszczeniu suwaka, a ze strzałek —
    // gdy przestaniesz naciskać (jedno przerysowanie zamiast jednego na krok).
    let dragging = false;
    let commitTimer = 0;
    const commit = () => {
      clearTimeout(commitTimer);
      if (!imgCard || imgCard.el !== el || +range.value === imgCard.savedPct) return;
      imgCard.savedPct = +range.value;
      imageEdit(imgCard.index, { action: "size", widthPct: +range.value });
    };
    const release = () => {
      if (!dragging) return;
      dragging = false;
      window.removeEventListener("pointerup", release, true);
      window.removeEventListener("pointercancel", release, true);
      commit();
      placeImageCard();
    };
    range.addEventListener("pointerdown", () => {
      dragging = true;
      // puszczenie także poza suwakiem (przeciągnięcie za kartę)
      window.addEventListener("pointerup", release, true);
      window.addEventListener("pointercancel", release, true);
    });
    range.addEventListener("lostpointercapture", release);
    range.addEventListener("input", () => {
      if (!imgCard || imgCard.el !== el) return;
      val.textContent = `${range.value}%`;
      const im = imgCard.img;
      im.style.width = `${(imgCard.textPx || imageTextWidth(imgCard.p)) * range.value / 100}px`; // podgląd na żywo (jak w pliku)
      im.style.height = "auto";
    });
    range.addEventListener("change", () => {
      if (dragging) return; // zapis przy puszczeniu
      clearTimeout(commitTimer);
      commitTimer = setTimeout(commit, 450); // strzałki / Page Up/Down
    });
    range.addEventListener("blur", () => { if (commitTimer) commit(); });
    const hint = (b, key) => { b.setAttribute("aria-label", t(key)); b.dataset.hint = ""; b.dataset.hintPl = I18N.pl[key]; b.dataset.hintEn = I18N.en[key]; b.dataset.hintDelay = "0.4"; };
    el.querySelectorAll("[data-align]").forEach((b) => {
      hint(b, ALIGN_KEYS[b.dataset.align]);
      b.innerHTML = alignSvg(b.dataset.align);
      b.addEventListener("click", () => imageEdit(imgCard.index, { action: "align", align: b.dataset.align }));
    });
    const alt = el.querySelector(".ic-alt");
    hint(alt, "imageAlt");
    alt.innerHTML = '<span class="ic-alt-text">ALT</span>';
    alt.addEventListener("click", () => openAltForm(imgCard.index, imgCard.img));
    const del = el.querySelector(".ic-del");
    hint(del, "imageDelete");
    del.innerHTML = ICON('<polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>');
    del.addEventListener("click", () => { const i = imgCard.index; hideImageCard(); imageEdit(i, { action: "delete" }, false); });
    el.addEventListener("mousedown", (e) => { if (e.target.closest("button")) e.preventDefault(); });
    document.body.appendChild(el);
    imgCard = { el, img, p, index, savedPct: +range.value, isDragging: () => dragging, textPx: 0 };
    placeImageCard();
    // procent względem strony z pliku (Widok mobilny ≠ szerokość ekranu)
    pageTextPx(index).then((tp) => {
      if (!tp || imgCard?.el !== el || dragging) return;
      imgCard.textPx = tp;
      const pct = Math.max(10, Math.min(100, Math.round(imageDocWidthPx(img) / tp * 100 / 5) * 5));
      range.value = String(pct);
      val.textContent = `${pct}%`;
      imgCard.savedPct = pct;
    });
  }
  function placeImageCard() {
    if (!imgCard || imgCard.busy || imgCard.isDragging?.()) return; // przy przesuwaniu suwaka karta stoi
    const vv = window.visualViewport;
    const r = imgCard.img.getBoundingClientRect();
    const vp = docViewportEl?.getBoundingClientRect();
    if (vp && (r.bottom < vp.top || r.top > vp.bottom)) { imgCard.el.style.visibility = "hidden"; return; }
    imgCard.el.style.visibility = "";
    const viewW = vv ? vv.width : window.innerWidth;
    const viewH = vv ? vv.height : window.innerHeight;
    const w = imgCard.el.offsetWidth; const h = imgCard.el.offsetHeight;
    const left = Math.max(8, Math.min(r.left + r.width / 2 - w / 2, viewW - w - 8));
    // pod obrazem, a gdy nie ma miejsca — nad nim (albo przy dolnej krawędzi widoku)
    let top = r.bottom + 8;
    if (top + h > viewH - 8) top = r.top - h - 8 >= 8 ? r.top - h - 8 : viewH - h - 8;
    imgCard.el.style.left = `${left + (vv ? vv.offsetLeft : 0)}px`;
    imgCard.el.style.top = `${top + (vv ? vv.offsetTop : 0)}px`;
  }
  // Zmiana obrazu w pliku + przerysowanie. Karta zostaje TA SAMA (bez znikania i ponownego
  // „wjazdu”) i po przerysowaniu przypina się do nowego elementu obrazu.
  async function imageEdit(index, edit, reselect = true) {
    if (imgCard && reselect) imgCard.busy = true;
    await applyDocumentEdit({ op: "image", index, ...edit }).catch((err) => log(`Obraz: ${err.message || err}`, "error"));
    if (!reselect) return;
    await whenEditable();
    const img = collectPreviewParagraphElements(host())[index]?.querySelector("img");
    if (!img || readOnlyMode) { hideImageCard(); return; }
    if (imgCard?.el.isConnected && imgCard.index === index) {
      imgCard.img.classList.remove("img-selected");
      img.classList.add("img-selected");
      imgCard.img = img;
      imgCard.p = img.closest("p");
      imgCard.busy = false;
      placeImageCard();
      return;
    }
    showImageCard(img);
  }
  function openAltForm(index, img) {
    openPop(imgCard?.el.querySelector(".ic-alt") || insertBtn, (el) => {
      el.classList.add("compose-pop-form");
      el.setAttribute("role", "dialog");
      el.setAttribute("aria-label", t("imageAlt"));
      el.innerHTML = `<div class="compose-cap"></div><label class="compose-field"><span></span><textarea rows="3" data-autofocus="1"></textarea></label>
        <div class="compose-actions"><span class="compose-actions-gap"></span><button type="button" class="btn lf-cancel"></button><button type="button" class="btn primary lf-ok"></button></div>`;
      el.querySelector(".compose-cap").textContent = t("imageAlt");
      el.querySelector(".compose-field span").textContent = t("imageAltHelp");
      const ta = el.querySelector("textarea");
      ta.value = img.getAttribute("alt") || "";
      el.querySelector(".lf-cancel").textContent = t("linkCancel");
      el.querySelector(".lf-cancel").addEventListener("click", () => closePop());
      el.querySelector(".lf-ok").textContent = t("linkSave");
      el.querySelector(".lf-ok").addEventListener("click", () => { closePop(); imageEdit(index, { action: "alt", descr: ta.value.trim() }); });
    });
  }

  // ── kolor czcionki i wyróżnienie (jak w Wordzie) ───────────────────────────
  const COLOR_KEY = "dwb.lastFontColor";
  const THEME_BASE = [
    ["FFFFFF", "colWhite"], ["000000", "colBlack"], ["E7E6E6", "colLightGray"], ["44546A", "colBlueGray"], ["4472C4", "colBlue"],
    ["ED7D31", "colOrange"], ["A5A5A5", "colGray"], ["FFC000", "colGold"], ["5B9BD5", "colLightBlue"], ["70AD47", "colGreen"],
  ];
  const STANDARD = [
    ["C00000", "colDarkRed"], ["FF0000", "colRed"], ["FFC000", "colOrange"], ["FFFF00", "colYellow"], ["92D050", "colLightGreen"],
    ["00B050", "colGreen"], ["00B0F0", "colLightBlue"], ["0070C0", "colBlue"], ["002060", "colDarkBlue"], ["7030A0", "colPurple"],
  ];
  // nazwa wyróżnienia Worda → kolor, jaki pokazuje Word (nazwy CSS bywają inne: „green” Worda to jasna zieleń)
  const HIGHLIGHTS = [
    ["yellow", "FFFF00", "colYellow"], ["green", "00FF00", "colBrightGreen"], ["cyan", "00FFFF", "colTurquoise"], ["magenta", "FF00FF", "colPink"],
    ["blue", "0000FF", "colBlue"], ["red", "FF0000", "colRed"], ["darkBlue", "000080", "colDarkBlue"], ["darkCyan", "008080", "colTeal"],
    ["darkGreen", "008000", "colDarkGreen"], ["darkMagenta", "800080", "colViolet"], ["darkRed", "800000", "colDarkRed"],
    ["darkGray", "808080", "colGray"], ["lightGray", "C0C0C0", "colLightGray"], ["black", "000000", "colBlack"],
  ];
  const mix = (hex, k, toWhite) => hex.match(/../g).map((h) => {
    const c = parseInt(h, 16);
    return Math.round(toWhite ? c + (255 - c) * k : c * (1 - k)).toString(16).padStart(2, "0");
  }).join("").toUpperCase();
  // kolumna motywu: kolor + 5 odcieni (jak paleta Worda)
  function themeColumn([hex, key]) {
    const variants = hex === "FFFFFF" ? [[0.05, 0], [0.15, 0], [0.25, 0], [0.35, 0], [0.5, 0]].map(([k]) => [mix(hex, k, false), `-${k * 100}`])
      : hex === "000000" ? [0.5, 0.35, 0.25, 0.15, 0.05].map((k) => [mix(hex, k, true), `+${k * 100}`])
      : [[0.8, true], [0.6, true], [0.4, true], [0.25, false], [0.5, false]].map(([k, w]) => [mix(hex, k, w), `${w ? "+" : "-"}${k * 100}`]);
    return [[hex, key, ""], ...variants.map(([h, d]) => [h, key, d])];
  }

  function lastColor() {
    try { return localStorage.getItem(COLOR_KEY) || "C00000"; } catch (_) { return "C00000"; }
  }
  function syncColorBar() {
    if (colorBar) colorBar.style.background = `#${lastColor()}`;
  }

  // Zaznaczenie w akapicie z kursorem (przycięte do akapitu) jako przesunięcia w jego tekście.
  function selectionInParagraph() {
    const p = caretParagraph();
    if (!p) return null;
    const index = resolveParaIndex(p);
    const range = window.getSelection().getRangeAt(0);
    const start = p.contains(range.startContainer) ? textOffset(p, range.startContainer, range.startOffset) : 0;
    const end = p.contains(range.endContainer) ? textOffset(p, range.endContainer, range.endOffset) : previewRunsToPlainText(extractRunsFromPreviewParagraph(p)).length;
    return { p, index, start, end, collapsed: range.collapsed || end <= start };
  }

  function reselect(index, start, end) {
    const el = collectPreviewParagraphElements(host())[index];
    const r = el && formDomRange(el, start, end);
    if (!r) return;
    focusDocParagraph(el);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  }

  // kind: "color" | "highlight"; value: "RRGGBB" / nazwa wyróżnienia / "" (= automatyczny / brak)
  async function applyColor(kind, value) {
    const ctx = selectionInParagraph();
    if (!ctx) return;
    if (kind === "color" && value) { try { localStorage.setItem(COLOR_KEY, value); } catch (_) { /* prywatne */ } syncColorBar(); }
    const css = kind === "color" ? (value ? `#${value}` : "#000000") : (value ? (value.length === 6 && /^[0-9A-F]+$/i.test(value) ? `#${value}` : value.toLowerCase()) : "");
    if (ctx.collapsed) {
      // bez zaznaczenia — kolor dla dalszego pisania (jak w Wordzie)
      restoreDocCaret();
      setTypingStyle({ [kind]: css || undefined }); // obowiązuje w tym miejscu kursora (docx-inline-edit.js)
      if (!css && activeTypingStyle) delete activeTypingStyle[kind];
      return;
    }
    const edit = { op: "runStyle", index: ctx.index, start: ctx.start, end: ctx.end };
    edit[kind] = kind === "color" ? (value ? `#${value}` : "auto") : css;
    await applyDocumentEdit(edit).catch((err) => log(`Kolor: ${err.message || err}`, "error"));
    await whenEditable();
    if (!readOnlyMode) reselect(ctx.index, ctx.start, ctx.end); // zaznaczenie zostaje — można dodać wyróżnienie
  }

  function swatch(hex, label, onPick, extraClass = "") {
    const b = document.createElement("button");
    b.type = "button";
    b.className = `color-swatch ${extraClass}`.trim();
    b.style.setProperty("--sw", `#${hex}`);
    b.setAttribute("aria-label", label);
    b.dataset.hint = label;
    b.dataset.hintDelay = "0.3";
    b.addEventListener("click", () => { closePop(); onPick(); });
    return b;
  }

  function buildColorMenu(el) {
    el.classList.add("compose-pop-color");
    popCap(el, t("colorFont"));
    const auto = document.createElement("button");
    auto.type = "button";
    auto.className = "compose-item compose-item-sm color-auto";
    auto.innerHTML = '<span class="color-auto-chip" aria-hidden="true"></span><span class="compose-item-label"></span>';
    auto.querySelector(".compose-item-label").textContent = t("colorAuto");
    auto.addEventListener("click", () => { closePop(); applyColor("color", ""); });
    el.appendChild(auto);
    const theme = document.createElement("div");
    theme.className = "color-grid color-grid-theme";
    const cols = THEME_BASE.map(themeColumn);
    for (let row = 0; row < 6; row++) {
      cols.forEach((col) => {
        const [hex, key, d] = col[row];
        const name = d ? t(d.startsWith("+") ? "colLighter" : "colDarker", { name: t(key), pct: d.slice(1) }) : t(key);
        theme.appendChild(swatch(hex, name, () => applyColor("color", hex), row === 0 ? "is-base" : ""));
      });
    }
    el.appendChild(theme);
    const std = document.createElement("div");
    std.className = "color-grid";
    STANDARD.forEach(([hex, key]) => std.appendChild(swatch(hex, t(key), () => applyColor("color", hex))));
    el.appendChild(std);
    const more = document.createElement("label");
    more.className = "compose-item compose-item-sm color-more";
    more.innerHTML = '<span class="color-more-chip" aria-hidden="true"></span><span class="compose-item-label"></span><input type="color">';
    more.querySelector(".compose-item-label").textContent = t("colorMore");
    const picker = more.querySelector("input");
    picker.value = `#${lastColor()}`;
    picker.addEventListener("change", () => { closePop(); applyColor("color", picker.value.replace("#", "").toUpperCase()); });
    el.appendChild(more);
    popCap(el, t("colorHighlight"));
    const hl = document.createElement("div");
    hl.className = "color-grid color-grid-hl";
    HIGHLIGHTS.forEach(([name, hex, key]) => hl.appendChild(swatch(hex, t(key), () => applyColor("highlight", name))));
    const none = swatch("FFFFFF", t("colorNoHighlight"), () => applyColor("highlight", ""), "is-none");
    hl.appendChild(none);
    el.appendChild(hl);
  }

  // ── komentarze ────────────────────────────────────────────────────────────
  const AUTHOR_KEY = "dwb.authorName";
  function authorName() { try { return localStorage.getItem(AUTHOR_KEY) || ""; } catch (_) { return ""; } }
  function initialsOf(name) { return String(name || "").trim().split(/\s+/).map((w) => w[0] || "").join("").slice(0, 3).toUpperCase(); }

  // dane komentarzy (autor, treść, odpowiedzi, „rozwiązany”) — z panelu Recenzja, raz na wersję pliku
  let commentData = { bytes: null, byId: new Map(), job: null };
  function loadComments() {
    const bytes = originalFileBytes;
    if (!bytes || typeof scanDocxRevisions !== "function") return Promise.resolve(commentData.byId);
    if (commentData.bytes === bytes) return commentData.job || Promise.resolve(commentData.byId);
    commentData = { bytes, byId: new Map(), job: null };
    commentData.job = scanDocxRevisions(bytes).then((r) => {
      if (commentData.bytes !== bytes) return commentData.byId;
      r.comments.forEach((c) => commentData.byId.set(c.id, c));
      commentData.job = null;
      paintCommentHighlights();
      queueSync();
      return commentData.byId;
    }).catch(() => commentData.byId);
    return commentData.job;
  }

  // Zakresy komentarzy w podglądzie (między znacznikami <span data-cm>).
  function commentRanges() {
    const h = host();
    if (!h) return [];
    const out = [];
    h.querySelectorAll('span[data-cm-kind="start"]').forEach((sp) => {
      const end = h.querySelector(`span[data-cm-kind="end"][data-cm-id="${CSS.escape(sp.dataset.cmId)}"]`);
      if (!end) return;
      const r = document.createRange();
      r.setStartAfter(sp);
      r.setEndBefore(end);
      out.push({ id: sp.dataset.cmId, range: r });
    });
    return out;
  }

  function paintCommentHighlights() {
    if (typeof CSS === "undefined" || !CSS.highlights || typeof Highlight === "undefined") return;
    const open = new Highlight();
    const done = new Highlight();
    commentRanges().forEach(({ id, range }) => {
      const c = commentData.byId.get(id);
      if (c?.parentId) return; // odpowiedź — zakres ma komentarz nadrzędny
      (c?.done ? done : open).add(range);
    });
    CSS.highlights.set("dwb-comment", open);
    CSS.highlights.set("dwb-comment-done", done);
    if (originalFileBytes && commentData.bytes !== originalFileBytes) loadComments();
  }

  function commentAtCaret() {
    const sel = window.getSelection();
    if (!sel?.rangeCount || !sel.isCollapsed) return null;
    const node = sel.anchorNode; const off = sel.anchorOffset;
    let hit = null;
    commentRanges().forEach(({ id, range }) => {
      try { if (range.isPointInRange(node, off) && commentData.byId.get(id) && !commentData.byId.get(id).parentId) hit = { id, range }; } catch (_) { /* inny dokument */ }
    });
    return hit;
  }

  // ── treść komentarza z formatowaniem (jak w Wordzie: B, I, U, przekreślenie; kolor i
  // wyróżnienie z pliku zostają) ↔ model { text, b, i, u, s, color, hl } (docx-revisions.js) ──
  const HL_CSS = Object.fromEntries(HIGHLIGHTS.map(([name, hex]) => [name, `#${hex}`]));
  function richRunEl(run) {
    let node = document.createTextNode(run.text);
    const wrap = (tag) => { const el = document.createElement(tag); el.appendChild(node); node = el; };
    if (run.b) wrap("b");
    if (run.i) wrap("i");
    if (run.u) wrap("u");
    if (run.s) wrap("s");
    if (run.color || run.hl) {
      wrap("span");
      if (run.color) { node.dataset.color = run.color; node.style.color = `#${run.color}`; }
      if (run.hl) { node.dataset.hl = run.hl; node.style.backgroundColor = HL_CSS[run.hl] || run.hl; }
    }
    return node;
  }
  function richToDom(paras, box) {
    box.textContent = "";
    (paras?.length ? paras : [[]]).forEach((runs) => {
      const p = document.createElement("div");
      runs.forEach((r) => { if (r.text) p.appendChild(richRunEl(r)); });
      if (!p.firstChild) p.appendChild(document.createElement("br"));
      box.appendChild(p);
    });
  }
  // DOM edytora → akapity. Blok (div/p) i <br> = nowy akapit (Enter w komentarzu Worda), <br>
  // zamykający blok to tylko wypełniacz pustego wiersza. Formatowanie z najbliższego przodka,
  // który je określa (Chrome/Safari: <b>/<i>/<u>/<strike>, „odpogrubienie” = font-weight: normal).
  const RICH_BLOCK = /^(DIV|P|LI|H[1-6]|BLOCKQUOTE)$/;
  function richFmtOf(node, root) {
    const f = {};
    const seen = new Set();
    const set = (k, v) => { if (!seen.has(k)) { seen.add(k); if (v) f[k] = v; } };
    for (let el = node.parentElement; el && el !== root; el = el.parentElement) {
      const tag = el.tagName;
      const st = el.style;
      if (st.fontWeight) set("b", st.fontWeight === "bold" || parseInt(st.fontWeight, 10) >= 600);
      if (tag === "B" || tag === "STRONG") set("b", true);
      if (st.fontStyle) set("i", st.fontStyle === "italic");
      if (tag === "I" || tag === "EM") set("i", true);
      const deco = st.textDecorationLine || st.textDecoration || "";
      if (/underline/.test(deco) || tag === "U") set("u", true);
      if (/line-through/.test(deco) || tag === "S" || tag === "STRIKE" || tag === "DEL") set("s", true);
      if (el.dataset.color) set("color", el.dataset.color);
      if (el.dataset.hl) set("hl", el.dataset.hl);
    }
    return f;
  }
  function domToRich(root) {
    const paras = [[]];
    const cur = () => paras[paras.length - 1];
    const walk = (node) => {
      if (node.nodeType === 3) {
        const text = node.nodeValue.replace(/\u00a0/g, " ").replace(/[\r\n]+/g, " ");
        if (!text) return;
        const f = richFmtOf(node, root);
        const last = cur()[cur().length - 1];
        if (last && ["b", "i", "u", "s", "color", "hl"].every((k) => (last[k] || false) === (f[k] || false))) last.text += text;
        else cur().push({ ...f, text });
        return;
      }
      if (node.nodeType !== 1) return;
      if (node.tagName === "BR") {
        const block = node.parentElement;
        if (block !== root && RICH_BLOCK.test(block.tagName) && node === block.lastChild) return; // wypełniacz
        paras.push([]);
        return;
      }
      const isBlock = node !== root && RICH_BLOCK.test(node.tagName);
      if (isBlock && cur().length) paras.push([]); // blok zaczyna wiersz
      node.childNodes.forEach(walk);
      if (isBlock) paras.push([]); // i go kończy (pusty blok = pusty wiersz)
    };
    walk(root);
    while (paras.length > 1 && !paras[paras.length - 1].length) paras.pop();
    while (paras.length > 1 && !paras[0].length) paras.shift();
    return paras;
  }
  // Schowek → akapity komentarza. Bloki z paste-rich.js (ten sam odczyt co wklejanie do
  // dokumentu): nagłówek = pogrubiony akapit, punkt listy = „• ” / „1. ”, tabela = wiersze z
  // tabulatorami, kod i link = sam tekst. null = nic do wklejenia.
  function clipboardToRich(cd) {
    const text = cd?.getData("text/plain") || "";
    const blocks = typeof dwbPaste !== "undefined" ? dwbPaste.parse(cd) : null;
    if (!blocks) return text ? text.replace(/\r\n?/g, "\n").split("\n").map((line) => (line ? [{ text: line }] : [])) : null;
    const paras = [];
    const counters = [];
    const conv = (r, extra = {}) => ({ text: r.text, b: !!(r.bold || extra.b), i: !!r.italic, u: !!r.underline, s: !!r.strike, ...(r.highlight ? { hl: r.highlight } : {}) });
    const pushRuns = (runs, prefix = "", extra = {}) => {
      let cur = prefix ? [{ text: prefix }] : [];
      runs.forEach((r) => {
        if (r.break) { paras.push(cur); cur = []; return; }
        if (r.text) cur.push(conv(r, extra));
      });
      paras.push(cur);
    };
    blocks.forEach((b) => {
      if (b.type !== "li") counters.length = 0;
      if (b.type === "hr") return;
      if (b.type === "table") { b.rows.forEach((row) => paras.push([{ text: row.map((cell) => cell.map((r) => (r.break ? " " : r.text)).join("")).join("\t") }])); return; }
      if (b.type === "li") {
        const lvl = b.level || 0;
        counters.length = lvl + 1;
        counters[lvl] = (counters[lvl] || 0) + 1;
        const mark = b.checked !== undefined ? (b.checked ? "☒ " : "☐ ") : b.ordered ? `${counters[lvl]}. ` : "• ";
        pushRuns(b.runs, "\t".repeat(lvl) + mark);
        return;
      }
      pushRuns(b.runs, "", b.type === "h" ? { b: true } : {});
    });
    return paras.map((runs) => runs.filter((r) => r.text)).filter((runs, i, all) => runs.length || (i && i < all.length - 1));
  }

  const richPlain = (paras) => paras.map((runs) => runs.map((r) => r.text).join("")).join("\n");

  // Przewijany obszar (lista w karcie, edytor): wygaszenie krawędzi + „Więcej ↓”, gdy treść
  // wychodzi poza widoczną część — długi komentarz nie może udawać krótkiego.
  function attachMoreBelow(scroller, host) {
    const more = document.createElement("button");
    more.type = "button";
    more.className = "cc-more";
    more.tabIndex = -1;
    more.innerHTML = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="6 9 12 15 18 9"/></svg><span></span>';
    more.querySelector("span").textContent = t("commentMore");
    more.setAttribute("aria-hidden", "true");
    more.addEventListener("mousedown", (e) => e.preventDefault());
    more.addEventListener("click", () => scroller.scrollBy({ top: Math.max(40, scroller.clientHeight * 0.8), behavior: "smooth" }));
    host.appendChild(more);
    const sync = () => {
      const max = scroller.scrollHeight - scroller.clientHeight;
      scroller.classList.toggle("more-t", max > 1 && scroller.scrollTop > 2);
      const below = max > 1 && scroller.scrollTop < max - 2;
      scroller.classList.toggle("more-b", below);
      more.hidden = !below;
    };
    scroller.addEventListener("scroll", sync, { passive: true });
    if (typeof ResizeObserver === "function") new ResizeObserver(sync).observe(scroller);
    if (typeof MutationObserver === "function") new MutationObserver(sync).observe(scroller, { childList: true, subtree: true, characterData: true });
    sync();
    return sync;
  }

  // Skala karty i edytora komentarza = przybliżenie dokumentu (dostępność: przy 150% komentarz też
  // 150%), nigdy mniej niż zwykły rozmiar (dopasowana do ekranu strona bywa pomniejszona).
  function commentZoom() {
    const z = typeof getDocZoom === "function" ? getDocZoom() : 1;
    return Math.max(1, Math.min(2.5, Number.isFinite(z) ? z : 1));
  }

  // Formatowanie treści komentarza — te same przyciski B / I / U z paska Edycji i te same skróty
  // (prośba Mateusza: bez osobnych przycisków w okienku). Gdy kursor stoi w treści komentarza,
  // przyciski paska działają na komentarz, nie na dokument.
  // Akapit, przy którym otwarto okienko komentarza, wraca po zmianie w to samo miejsce EKRANU
  // (nie obszaru dokumentu — po schowaniu klawiatury wraca rząd skrótów i obszar zaczyna się niżej). Na telefonie w międzyczasie wysuwa się i chowa klawiatura,
  // zwija i rozwija nagłówek — sama kotwica przerysowania tego nie widzi (zgłoszenie: „po dodaniu
  // komentarza muszę szukać, gdzie byłem”).
  function paragraphSpot(index) {
    const p = index >= 0 ? collectPreviewParagraphElements(host())[index] : null;
    return p && docViewportEl ? { index, off: p.getBoundingClientRect().top } : null;
  }
  function restoreParagraphSpot(spot) {
    if (!spot || !docViewportEl) return;
    const apply = () => {
      const p = collectPreviewParagraphElements(host())[spot.index];
      if (!p) return;
      const vp = docViewportEl.getBoundingClientRect();
      // na ekranie tam, gdzie był — o ile mieści się w obszarze dokumentu
      const want = Math.max(vp.top, Math.min(spot.off, vp.bottom - 40));
      const d = p.getBoundingClientRect().top - want;
      if (Math.abs(d) > 1) docViewportEl.scrollTop += d;
      placeCommentCard();
    };
    apply();
    // klawiatura chowa się z animacją (obszar dokumentu rośnie) — jeszcze raz, gdy się ustatkuje
    const vv = window.visualViewport;
    if (!vv) return;
    let t = 0;
    const onResize = () => { clearTimeout(t); t = setTimeout(apply, 120); };
    vv.addEventListener("resize", onResize);
    setTimeout(() => { vv.removeEventListener("resize", onResize); clearTimeout(t); }, 1500);
  }

  const RICH_KEYS = { KeyB: "bold", KeyI: "italic", KeyU: "underline" };
  const RICH_BTNS = { fmtBold: "bold", fmtItalic: "italic", fmtUnderline: "underline" };
  function runRich(cmd) {
    try { document.execCommand("styleWithCSS", false, false); } catch (_) { /* Safari bez tej opcji */ }
    document.execCommand(cmd, false, null);
  }
  const activeCommentEditor = () => {
    const ed = pop?.el.querySelector(".cf-rich");
    return ed && (document.activeElement === ed || ed.contains(document.activeElement)) ? ed : null;
  };
  // capture na dokumencie — przed obsługą przycisków w docx-inline-edit.js (format dokumentu)
  document.addEventListener("click", (e) => {
    const b = e.target.closest?.("#fmtBold, #fmtItalic, #fmtUnderline");
    if (!b || !activeCommentEditor()) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    runRich(RICH_BTNS[b.id]);
  }, true);
  function openCommentForm(anchor, opts = {}) {
    if (readOnlyMode) return;
    const editing = !!opts.editId;
    const ctx = opts.replyTo || editing ? null : selectionInParagraph();
    if (!opts.replyTo && !editing && !ctx) return;
    hideCommentCard();
    closePop();
    const capKey = editing ? "commentEditTitle" : opts.replyTo ? "commentReply" : "insertComment";
    const spotIndex = ctx ? ctx.index : opts.caret?.paraIndex ?? -1;
    const spot = paragraphSpot(spotIndex);
    openPop(anchor || insertBtn, (el) => {
      el.classList.add("compose-pop-form", "compose-pop-comment");
      el.style.setProperty("--cc-z", String(commentZoom()));
      el.setAttribute("role", "dialog");
      el.setAttribute("aria-label", t(capKey));
      el.innerHTML = `<div class="compose-cap"></div>
        <div class="compose-field cf-field"><span class="cf-label"></span>
          <div class="cf-wrap"><div class="cf-rich" contenteditable="true" role="textbox" aria-multiline="true" spellcheck="true" data-autofocus="1"></div></div>
          <span class="compose-note cf-fmt-hint"></span>
        </div>
        <label class="compose-field cf-author-field"><span></span><input type="text" class="cf-author" autocomplete="name" enterkeyhint="done"></label>
        <div class="compose-actions"><span class="compose-actions-gap"></span><button type="button" class="btn lf-cancel"></button><button type="button" class="btn primary lf-ok"></button></div>`;
      el.querySelector(".compose-cap").textContent = t(capKey);
      const label = el.querySelector(".cf-label");
      label.id = `cfLabel${Date.now()}`;
      label.textContent = t(opts.replyTo ? "commentReplyText" : "commentText");
      const ed = el.querySelector(".cf-rich");
      ed.setAttribute("aria-labelledby", label.id);
      richToDom(editing ? opts.rich : null, ed);
      const au = el.querySelector(".cf-author");
      el.querySelector(".cf-author-field > span").textContent = t("commentAuthor");
      el.querySelector(".cf-author-field").hidden = editing; // edycja nie zmienia podpisu
      au.value = authorName();
      au.placeholder = t("commentAuthorPh");
      el.querySelector(".cf-fmt-hint").textContent = t("commentFmtHint");
      ed.addEventListener("keydown", (e) => {
        const mod = e.ctrlKey || e.metaKey;
        if (mod && !e.altKey && !e.shiftKey && RICH_KEYS[e.code]) { e.preventDefault(); e.stopPropagation(); runRich(RICH_KEYS[e.code]); return; }
        if (mod && e.key === "Enter") { e.preventDefault(); e.stopPropagation(); submit(); }
      });
      // wklejanie z formatowaniem, które komentarz Worda zapisuje (B / I / U / przekreślenie /
      // wyróżnienie) — z Worda, Notatek, stron WWW, Markdownu; kroje, rozmiary, obrazy i linki — nie
      ed.addEventListener("paste", (e) => {
        e.preventDefault();
        e.stopPropagation();
        const paras = clipboardToRich(e.clipboardData);
        if (!paras) return;
        if (paras.length === 1 && paras[0].every((r) => !["b", "i", "u", "s", "hl"].some((k) => r[k]))) {
          document.execCommand("insertText", false, paras[0].map((r) => r.text).join(""));
          return;
        }
        const box = document.createElement("div");
        paras.forEach((runs, i) => {
          if (i) box.appendChild(document.createElement("br"));
          runs.forEach((r) => box.appendChild(richRunEl(r)));
        });
        document.execCommand("insertHTML", false, box.innerHTML); // HTML zbudowany z modelu (tekst w węzłach tekstowych)
      });
      ed.addEventListener("drop", (e) => e.preventDefault());
      attachMoreBelow(ed, el.querySelector(".cf-wrap"));
      el.querySelector(".lf-cancel").textContent = t("linkCancel");
      el.querySelector(".lf-cancel").addEventListener("click", () => closePop());
      el.querySelector(".lf-ok").textContent = t(editing ? "linkSave" : opts.replyTo ? "commentReplyBtn" : "commentAddBtn");
      const submit = async () => {
        const rich = domToRich(ed);
        const text = richPlain(rich).trim();
        const author = au.value.trim();
        if (!text) { ed.focus(); return; }
        if (editing) {
          closePop();
          await runFileEdit({ op: "commentEdit", id: opts.editId, rich }, opts.caret || null);
          await whenEditable();
          restoreParagraphSpot(spot);
          return;
        }
        if (!author) { toast(t("commentNeedAuthor"), "info"); au.focus(); return; }
        try { localStorage.setItem(AUTHOR_KEY, author); } catch (_) { /* prywatne */ }
        closePop();
        if (opts.replyTo) {
          await runFileEdit({ op: "commentReply", id: opts.replyTo, rich, author, initials: initialsOf(author) }, opts.caret || null);
          await whenEditable();
          restoreParagraphSpot(spot);
          return;
        }
        let { start, end } = ctx;
        if (end <= start) { // bez zaznaczenia — słowo pod kursorem (jak w Wordzie)
          const full = previewRunsToPlainText(extractRunsFromPreviewParagraph(ctx.p));
          while (start > 0 && /\S/.test(full[start - 1])) start--;
          while (end < full.length && /\S/.test(full[end])) end++;
        }
        await runFileEdit({ op: "commentAdd", index: ctx.index, start, end, rich, author, initials: initialsOf(author) }, { paraIndex: ctx.index, offset: end });
        await whenEditable();
        restoreParagraphSpot(spot);
      };
      el.querySelector(".lf-ok").addEventListener("click", submit);
      au.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); submit(); } });
      if (editing) { // kursor na końcu treści
        const r = document.createRange();
        r.selectNodeContents(ed);
        r.collapse(false);
        setTimeout(() => { if (document.activeElement === ed) { const s = window.getSelection(); s.removeAllRanges(); s.addRange(r); } }, 0);
      }
    });
    if (!authorName() && !editing) setTimeout(() => pop?.el.querySelector(".cf-rich")?.focus(), 0);
  }

  // Karta komentarza: NAD wierszem z kursorem (pod nim bywa karta linku), tylko gdy kursor stoi
  // w komentowanym tekście; pisanie ją chowa do następnego ruchu kursora.
  let cCard = null; // { el, id }
  let cCardMuted = false;
  function hideCommentCard() { if (cCard) { cCard.el.remove(); cCard = null; } }
  function fmtDate(iso) {
    const d = iso ? new Date(iso) : null;
    return d && !isNaN(d) ? d.toLocaleDateString(currentLang === "en" ? "en-GB" : "pl-PL", { day: "numeric", month: "short", year: "numeric" }) : "";
  }
  // ro = tryb Czytanie: sama treść (autor, data, odpowiedzi), bez przycisków zmian
  function showCommentCard(hit, ro = false) {
    const c = commentData.byId.get(hit.id);
    if (!c) return;
    const sig = `${c.done}|${ro}|${commentData.bytes?.byteLength}`;
    if (cCard?.id === hit.id && cCard.sig === sig) { placeCommentCard(); return; }
    hideCommentCard();
    const el = document.createElement("div");
    el.className = `comment-card${c.done ? " is-done" : ""}`;
    el.setAttribute("role", "region");
    el.setAttribute("aria-label", t("commentCard"));
    el.style.setProperty("--cc-z", String(commentZoom()));
    const caretNow = () => { const p = activeParagraph() || lastP; return p?.isConnected ? caretState(p) : null; };
    const entry = (x) => {
      const box = document.createElement("div");
      box.className = "cc-entry";
      const head = document.createElement("div");
      head.className = "cc-head";
      const who = document.createElement("span");
      who.className = "cc-who";
      who.textContent = [x.author, fmtDate(x.date)].filter(Boolean).join(" · ");
      head.appendChild(who);
      if (!ro) {
        const ed = document.createElement("button");
        ed.type = "button";
        ed.className = "tb-btn cc-edit";
        ed.innerHTML = ICON('<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>');
        const key = x.editable === false ? "commentEditBlocked" : "commentEdit";
        ed.setAttribute("aria-label", t(key));
        ed.dataset.hint = "";
        ed.dataset.hintPl = I18N.pl[key];
        ed.dataset.hintEn = I18N.en[key];
        ed.dataset.hintDelay = "0.3";
        if (x.editable === false) ed.setAttribute("aria-disabled", "true");
        ed.addEventListener("click", () => {
          if (x.editable === false) { toast(t("commentEditBlocked"), "info"); return; }
          openCommentForm(ed, { editId: x.id, rich: x.rich?.length ? x.rich : [[{ text: x.text }]], caret: caretNow() });
        });
        head.appendChild(ed);
      }
      const body = document.createElement("div");
      body.className = "cc-text";
      if (x.rich?.length) {
        x.rich.forEach((runs) => {
          const p = document.createElement("div");
          p.className = "cc-p";
          runs.forEach((r) => p.appendChild(richRunEl(r)));
          if (!runs.length) p.appendChild(document.createElement("br"));
          body.appendChild(p);
        });
      } else body.textContent = x.text;
      box.append(head, body);
      return box;
    };
    const list = document.createElement("div");
    list.className = "cc-list";
    list.tabIndex = 0; // przewijanie strzałkami
    list.append(entry(c), ...c.replies.map(entry));
    const listWrap = document.createElement("div");
    listWrap.className = "cc-list-wrap";
    listWrap.appendChild(list);
    const actions = document.createElement("div");
    actions.className = "cc-actions";
    const mk = (key, fn, cls = "") => { const b = document.createElement("button"); b.type = "button"; b.className = `btn ${cls}`.trim(); b.textContent = t(key); b.addEventListener("click", fn); return b; };
    actions.append(
      mk("commentReply", () => openCommentForm(actions.firstChild, { replyTo: c.id, caret: caretNow() })),
      mk(c.done ? "commentReopen" : "commentResolve", () => runFileEdit({ op: "commentDone", id: c.id, done: !c.done }, caretNow())),
      mk("commentDelete", () => { hideCommentCard(); runFileEdit({ op: "revisions", action: "removeComments", ids: [c.id, ...c.replies.map((r) => r.id)] }, caretNow()); }, "cc-del"),
    );
    if (c.done) { const badge = document.createElement("div"); badge.className = "cc-done"; badge.textContent = t("commentResolved"); el.appendChild(badge); }
    if (ro) {
      const note = document.createElement("div");
      note.className = "cc-ro-note";
      note.textContent = t("commentReadOnlyNote");
      el.append(listWrap, note);
    } else el.append(listWrap, actions);
    el.addEventListener("mousedown", (e) => { if (e.target.closest("button")) e.preventDefault(); });
    document.body.appendChild(el);
    attachMoreBelow(list, listWrap);
    cCard = { el, id: hit.id, sig, ro, range: hit.range };
    placeCommentCard();
  }
  const coarsePointer = () => !!window.matchMedia?.("(pointer: coarse)").matches;
  // Wysokość karty komentarza (Mateusz 2026-10-05). Karta jest tak wysoka jak treść, ale:
  //   max    = najwyżej taka część widocznej wysokości ekranu (dłuższa treść przewija się z „Więcej ↓”),
  //   lines / screen = najmniej (gdy brakuje miejsca): linijki treści w skali tekstu albo część ekranu,
  //            co większe (jedna linijka to za mało), chrome = autor, odstępy i przyciski (px).
  const COMMENT_CARD_MIN = { touch: { lines: 5, screen: 0.30, chrome: 82, max: 0.5 }, desktop: { lines: 8, screen: 0.40, chrome: 66, max: 0.61 } };
  function placeCommentCard() {
    if (!cCard) return;
    cCard.el.style.setProperty("--cc-z", String(commentZoom()));
    const sel = window.getSelection();
    if (!sel?.rangeCount) return;
    const rects = sel.getRangeAt(0).getClientRects();
    let r = rects[0] || sel.getRangeAt(0).getBoundingClientRect();
    // kursor tuż przy znaczniku (wyspa bez szerokości) — przeglądarka daje pusty prostokąt (0,0);
    // wtedy kotwicą jest komentowany tekst (dawniej karta chowała się jak „poza ekranem”)
    if ((!r || (!r.width && !r.height && !r.top && !r.left)) && cCard.range?.startContainer?.isConnected) {
      r = cCard.range.getClientRects()[0] || cCard.range.getBoundingClientRect();
    }
    const vv = window.visualViewport;
    const vp = docViewportEl?.getBoundingClientRect();
    if (!r || (vp && (r.bottom < vp.top || r.top > vp.bottom))) { cCard.el.style.visibility = "hidden"; return; }
    cCard.el.style.visibility = "";
    const viewW = vv ? vv.width : window.innerWidth;
    const viewH = vv ? vv.height : window.innerHeight;
    // nad wierszem; brak miejsca — pod nim (z zapasem na kartę linku); długi komentarz, który nie
    // mieści się nigdzie — tam, gdzie więcej miejsca, niższy (treść przewija się z „Więcej ↓”)
    const minTop = Math.max(8, vp?.top ?? 8);
    const above = r.top - 8 - minTop;
    const below = viewH - 8 - (r.bottom + 40);
    const m = COMMENT_CARD_MIN[coarsePointer() ? "touch" : "desktop"];
    const maxH = Math.round(viewH * m.max);
    cCard.el.style.maxHeight = `${maxH}px`;
    const w = cCard.el.offsetWidth; let h = cCard.el.offsetHeight;
    const left = Math.max(8, Math.min(r.left - 16, viewW - w - 8));
    let top;
    if (h <= above) top = r.top - h - 8;
    else if (h <= below) top = r.bottom + 40;
    else {
      const up = above >= below;
      // minimum: linijki treści w skali tekstu komentarza (przy przybliżeniu 150% wyższe) albo część
      // widocznego ekranu (większy ekran = wyższa karta) — co większe; progi w COMMENT_CARD_MIN
      const z = commentZoom();
      const minH = Math.min(maxH, Math.max(Math.round((22 + m.lines * 19) * z) + m.chrome, Math.round(viewH * m.screen)));
      cCard.el.style.maxHeight = `${Math.min(maxH, Math.max(minH, up ? above : below))}px`;
      h = cCard.el.offsetHeight;
      top = up ? Math.max(minTop, r.top - h - 8) : r.bottom + 40;
      top = Math.max(minTop, Math.min(top, viewH - h - 8)); // cała karta na ekranie (najwyżej zachodzi na wiersz)
    }
    cCard.el.style.left = `${left + (vv ? vv.offsetLeft : 0)}px`;
    cCard.el.style.top = `${top + (vv ? vv.offsetTop : 0)}px`;
  }

  // ── nagłówek i stopka ─────────────────────────────────────────────────────
  // Strony w podglądzie: docx-preview łamie tylko przy jawnych podziałach, więc wysoka kartka =
  // kilka stron (szacunek z wysokości strony). Nagłówek kartki = jej pierwsza strona, stopka = ostatnia.
  function sectionPages() {
    const h = host();
    const out = [];
    let page = 1;
    (h ? Array.from(h.querySelectorAll("section.docx")) : []).forEach((sec) => {
      const ph = parseFloat(getComputedStyle(sec).minHeight) || sec.offsetHeight || 1;
      const n = Math.max(1, Math.ceil((sec.offsetHeight - 2) / ph));
      out.push({ sec, first: page, last: page + n - 1 });
      page += n;
    });
    return { list: out, total: Math.max(1, page - 1) };
  }

  // Fragmenty z numerem strony (styl znakowy „page number” pól PAGE / NUMPAGES) — też dla
  // podglądu wydruku (print-preview.js), który numeruje każdą kartkę.
  function pageNumberSelector() {
    const cls = [...docComposeStyleClasses].filter(([, k]) => k === "pagenum").map(([c]) => c);
    return cls.map((c) => `span.${CSS.escape(c)}`).join(",");
  }

  function fixPreviewPageNumbers() {
    const sel = pageNumberSelector();
    if (!sel) return;
    const { list, total } = sectionPages();
    list.forEach(({ sec, first, last }) => {
      sec.querySelectorAll(":scope > header, :scope > footer").forEach((part) => {
        const page = part.localName === "header" ? first : last;
        part.querySelectorAll("p").forEach((p) => {
          const spans = Array.from(p.querySelectorAll(sel));
          if (spans[0]) spans[0].textContent = String(page);
          if (spans[1]) spans[1].textContent = String(total);
        });
      });
    });
  }

  const ALIGN3 = ["left", "center", "right"];
  function segAlign(name, value) {
    return `<div class="seg seg-icons" role="group" data-seg="${name}">${ALIGN3.map((a) => `<button type="button" data-v="${a}" class="${a === value ? "is-on" : ""}" aria-pressed="${a === value}" aria-label="${t(ALIGN_KEYS[a])}">${alignSvg(a)}</button>`).join("")}</div>`;
  }

  async function openHeaderFooterForm(opts = {}) {
    if (readOnlyMode || !originalFileBytes) return;
    let st;
    try { await mergeInlineEditsIntoBytes(); st = await readHeaderFooterState(originalFileBytes); } catch (err) { log(`Nagłówek: ${err.message || err}`, "error"); return; }
    if (opts.pageNumber && !st.footer.number) { st.footer.number = "n"; st.footer.numAlign = "center"; }
    hideImageCard();
    closePop();
    openPop(insertBtn, (el) => {
      el.classList.add("compose-pop-form", "compose-pop-hf");
      el.setAttribute("role", "dialog");
      el.setAttribute("aria-label", t("hfTitle"));
      const fmts = [["", "hfNumNone"], ["n", "hfNumN"], ["page", "hfNumPage"], ["pageOf", "hfNumPageOf"], ["dash", "hfNumDash"]];
      el.innerHTML = `<div class="compose-cap"></div>
        <label class="compose-field"><span></span><input type="text" class="hf-h" autocomplete="off" enterkeyhint="next"></label>
        ${segAlign("h", st.header.align)}
        <label class="compose-field"><span></span><input type="text" class="hf-f" autocomplete="off" enterkeyhint="done"></label>
        ${segAlign("f", st.footer.align)}
        <label class="compose-field"><span></span><select class="hf-n">${fmts.map(([v, k]) => `<option value="${v}">${t(k)}</option>`).join("")}</select></label>
        ${segAlign("n", st.footer.numAlign || "center")}
        <label class="compose-check"><input type="checkbox" class="hf-first"><span></span></label>
        <p class="compose-note hf-note"></p>
        <div class="compose-actions"><span class="compose-actions-gap"></span><button type="button" class="btn lf-cancel"></button><button type="button" class="btn primary lf-ok"></button></div>`;
      el.querySelector(".compose-cap").textContent = t("hfTitle");
      const labels = el.querySelectorAll(".compose-field > span");
      labels[0].textContent = t("hfHeader");
      labels[1].textContent = t("hfFooter");
      labels[2].textContent = t("hfNumber");
      el.querySelector(".compose-check span").textContent = t("hfFirst");
      const hIn = el.querySelector(".hf-h"); const fIn = el.querySelector(".hf-f"); const nSel = el.querySelector(".hf-n"); const first = el.querySelector(".hf-first");
      hIn.value = st.header.text; fIn.value = st.footer.text; nSel.value = st.footer.number || ""; first.checked = st.firstDifferent;
      hIn.placeholder = t("hfHeaderPh"); fIn.placeholder = t("hfFooterPh");
      (opts.pageNumber ? nSel : hIn).dataset.autofocus = "1";
      const notes = [];
      if (st.header.complex || st.footer.complex) notes.push(t("hfComplex"));
      if (st.sections > 1) notes.push(t("hfSections", { n: st.sections }));
      notes.push(t("hfWordUpdates"));
      el.querySelector(".hf-note").textContent = notes.join(" ");
      const segs = {};
      el.querySelectorAll("[data-seg]").forEach((g) => {
        segs[g.dataset.seg] = g.querySelector(".is-on")?.dataset.v || "left";
        g.addEventListener("click", (e) => {
          const b = e.target.closest("button[data-v]");
          if (!b) return;
          segs[g.dataset.seg] = b.dataset.v;
          g.querySelectorAll("button").forEach((x) => { x.classList.toggle("is-on", x === b); x.setAttribute("aria-pressed", String(x === b)); });
        });
      });
      const syncNum = () => { el.querySelector('[data-seg="n"]').hidden = !nSel.value; };
      nSel.addEventListener("change", syncNum);
      syncNum();
      el.querySelector(".lf-cancel").textContent = t("linkCancel");
      el.querySelector(".lf-cancel").addEventListener("click", () => closePop());
      el.querySelector(".lf-ok").textContent = t("linkSave");
      el.querySelector(".lf-ok").addEventListener("click", async () => {
        const edit = {
          op: "headerFooter", lang: currentLang, total: sectionPages().total,
          header: { text: hIn.value.trim(), align: segs.h },
          footer: { text: fIn.value.trim(), align: segs.f },
          number: { fmt: nSel.value || null, align: segs.n },
          firstDifferent: first.checked,
        };
        closePop();
        const p = activeParagraph() || lastP;
        await runFileEdit(edit, p?.isConnected ? caretState(p) : null);
      });
      el.addEventListener("keydown", (e) => { if (e.key === "Enter" && e.target.matches("input[type=text]")) { e.preventDefault(); el.querySelector(".lf-ok").click(); } });
    });
  }

  // ── stan paska (styl i wyrównanie akapitu z kursorem) ──────────────────────
  function activeParagraph() {
    const sel = window.getSelection();
    const node = sel?.rangeCount ? sel.getRangeAt(0).startContainer : null;
    const el = node && (node.nodeType === 1 ? node : node.parentElement);
    const p = el?.closest?.(".docx-editable-p");
    return p && docCanvasEl?.contains(p) ? p : null;
  }

  let lastP = null;
  function currentAlign() {
    const p = activeParagraph() || lastP;
    if (!p) return "left";
    const a = getComputedStyle(p).textAlign;
    if (a === "center") return "center";
    if (a === "right" || a === "end") return "right";
    if (a === "justify") return "both";
    return "left";
  }

  function syncState() {
    const p = activeParagraph();
    if (p) lastP = p;
    const target = p || lastP;
    if (styleSel && document.activeElement !== styleSel) {
      styleSel.value = target && typeof composeStyleKeyOf === "function" ? composeStyleKeyOf(target) : "normal";
    }
    if (alignBtn) alignBtn.innerHTML = alignSvg(currentAlign());
    listBtn?.classList.toggle("is-on", !!target && isListParagraph(target) && !pop);
    // tylko gdy kursor stoi (albo ostatnio stał — fokus mógł przejść na pasek) w tabeli
    if (tableBtn) tableBtn.hidden = !(target?.isConnected && target.closest("td, th") && !readOnlyMode);
    const sel = window.getSelection();
    const a = !readOnlyMode && !pop && sel?.isCollapsed ? linkAtCaret() : null;
    if (a) showLinkCard(a); else hideLinkCard();
    const hit = !readOnlyMode && !pop && !cCardMuted ? commentAtCaret() : null;
    if (hit) showCommentCard(hit); else hideCommentCard();
  }

  let syncQueued = false;
  function queueSync() {
    if (syncQueued || readOnlyMode) return;
    syncQueued = true;
    requestAnimationFrame(() => { syncQueued = false; syncState(); });
  }

  // ── nowy dokument ─────────────────────────────────────────────────────────
  function openNewDialog() {
    if (!dialog) return;
    if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
    if (typeof appFrame !== "undefined") appFrame.setMenuOpen?.(false);
    if (dialog.open) return;
    if (typeof dialog.showModal === "function") dialog.showModal();
    else dialog.setAttribute("open", "");
    dialog.querySelector(".newdoc-card")?.focus();
  }

  async function createNew(kind) {
    if (dialog?.open) dialog.close();
    if (typeof confirmDiscardChanges === "function" && !confirmDiscardChanges()) return false;
    try {
      const ok = await ensureDocLibs(true);
      if (!ok) return false;
      const bytes = await createComposeDocx(kind, currentLang);
      const nameKey = { letter: "newDocFileLetter", note: "newDocFileNote" }[kind] || "newDocFileBlank";
      const file = new File([bytes], `${t(nameKey)}.docx`, {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const loaded = await window.ingestFile(file, { silent: true });
      if (!loaded) return false;
      // od razu pisanie: Edycja + kursor na początku (zużyje go syncInlineEditMode)
      pendingInlineCursor = { paraIndex: 0, offset: 0 };
      if (typeof appFrame !== "undefined") appFrame.setReadOnly(false);
      toast(t("newDocCreated"), "success");
      return true;
    } catch (err) {
      log(`Nowy dokument: ${err.message || err}`, "error");
      toast(t("newDocFailed"), "error");
      return false;
    }
  }

  // ── podpięcie ─────────────────────────────────────────────────────────────
  if (bar && typeof attachOverflowFade === "function") attachOverflowFade(bar);
  if (alignBtn) alignBtn.innerHTML = alignSvg("left");

  // Przyciski nie zabierają fokusu z tekstu — przez mousedown, NIE pointerdown: w Safari/iOS
  // preventDefault na pointerdown przy dotyku kasuje kliknięcie (tap nic nie robił).
  insertBtn?.addEventListener("mousedown", (e) => e.preventDefault()); // kursor zostaje w tekście
  insertBtn?.addEventListener("click", () => openPop(insertBtn, buildInsertMenu));
  alignBtn?.addEventListener("mousedown", (e) => e.preventDefault());
  listBtn?.addEventListener("mousedown", (e) => e.preventDefault());
  listBtn?.addEventListener("click", () => openPop(listBtn, buildListMenu));
  alignBtn?.addEventListener("click", () => openPop(alignBtn, buildAlignMenu));
  styleSel?.addEventListener("change", () => applyStyle(styleSel.value));
  colorBtn?.addEventListener("mousedown", (e) => e.preventDefault());
  colorBtn?.addEventListener("click", () => openPop(colorBtn, buildColorMenu));
  syncColorBar();
  tableBtn?.addEventListener("mousedown", (e) => e.preventDefault());
  tableBtn?.addEventListener("click", () => openPop(tableBtn, buildTableMenu));

  // Obraz: klik w Edycji pokazuje kartę; klik obok / Esc / tryb Czytanie chowa.
  // Klik/stuknięcie w obraz NIE stawia kursora (obraz się zaznacza, jak w Wordzie) — na iPhonie
  // kursor wysuwał klawiaturę, która zasłaniała kartę z suwakiem. Przez mousedown, NIE pointerdown
  // (w Safari/iOS preventDefault na pointerdown kasuje kliknięcie).
  docCanvasEl?.addEventListener("mousedown", (e) => {
    const img = e.target.closest?.(".docx-preview-host img");
    if (!img || readOnlyMode || !collectPreviewParagraphElements(host()).includes(img.closest("p"))) return;
    e.preventDefault();
    if (document.activeElement?.closest?.(".docx-edit-root")) document.activeElement.blur(); // klawiatura ekranowa chowa się
  });
  docCanvasEl?.addEventListener("click", (e) => {
    const img = e.target.closest?.(".docx-preview-host img");
    if (!img || readOnlyMode || !collectPreviewParagraphElements(host()).includes(img.closest("p"))) return;
    e.preventDefault();
    showImageCard(img);
  });
  document.addEventListener("pointerdown", (e) => {
    if (!imgCard || imgCard.el.contains(e.target) || e.target === imgCard.img || pop?.el.contains(e.target)) return;
    hideImageCard();
  }, true);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && imgCard && !pop) { e.preventDefault(); hideImageCard(); }
    if ((e.key === "Delete" || e.key === "Backspace") && imgCard && !pop && !e.target.closest?.("input, textarea")) {
      e.preventDefault();
      const { index } = imgCard;
      hideImageCard();
      imageEdit(index, { action: "delete" }, false);
    }
  }, true);
  document.getElementById("readMode")?.addEventListener("change", () => { hideImageCard(); hideCommentCard(); if (tableBtn) tableBtn.hidden = true; });

  // Czytanie: klik / stuknięcie w komentowany tekst pokazuje komentarz (tylko do podglądu, jak
  // w Wordzie w trybie tylko do odczytu); klik obok albo zaznaczanie tekstu — karta znika.
  function commentAtPoint(x, y) {
    let node = null; let off = 0;
    if (document.caretRangeFromPoint) { const r = document.caretRangeFromPoint(x, y); if (r) { node = r.startContainer; off = r.startOffset; } }
    else if (document.caretPositionFromPoint) { const p = document.caretPositionFromPoint(x, y); if (p) { node = p.offsetNode; off = p.offset; } }
    if (!node) return null;
    let hit = null;
    commentRanges().forEach(({ id, range }) => {
      try { if (range.isPointInRange(node, off) && commentData.byId.get(id) && !commentData.byId.get(id).parentId) hit = { id, range }; } catch (_) { /* inny dokument */ }
    });
    return hit;
  }
  docCanvasEl?.addEventListener("click", (e) => {
    if (!readOnlyMode || e.button > 0) return;
    if (!host()?.contains(e.target) || e.target.closest("a[href], sup[data-dwb-note], sup.note-mark")) { hideCommentCard(); return; }
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed && sel.toString().trim()) { hideCommentCard(); return; } // zaznaczanie tekstu
    const hit = commentAtPoint(e.clientX, e.clientY);
    if (hit) showCommentCard(hit, true); else hideCommentCard();
  });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && cCard?.ro) hideCommentCard(); });
  // klik poza dokumentem i poza kartą (pasek, panel) — karta Czytania znika
  document.addEventListener("pointerdown", (e) => { if (cCard?.ro && !cCard.el.contains(e.target) && !docCanvasEl?.contains(e.target)) hideCommentCard(); }, true);

  // Wklejony obraz (zrzut ekranu, skopiowane zdjęcie) — w miejscu kursora, jak „Wstaw → Obraz”.
  docCanvasEl?.addEventListener("paste", (e) => {
    if (readOnlyMode) return;
    const file = Array.from(e.clipboardData?.files || []).find((f) => /^image\//.test(f.type));
    if (!file) return;
    if (typeof dwbSel !== "undefined") dwbSel.collapseForInsert(); // obraz w miejsce zaznaczenia kilku akapitów
    const p = docCaretParagraph(e.target);
    if (!p) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    insertImageFile(file, p);
  }, true);

  // klik obok / przewinięcie / zmiana rozmiaru — okienko znika
  document.addEventListener("pointerdown", (e) => {
    if (!pop || pop.el.contains(e.target) || pop.anchor.contains(e.target)) return;
    if (e.target.closest?.("#fmtBold, #fmtItalic, #fmtUnderline") && activeCommentEditor()) return; // B / I / U dla komentarza
    closePop();
  }, true);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && pop) { e.preventDefault(); e.stopPropagation(); closePop(true); }
  }, true);
  docViewportEl?.addEventListener("scroll", () => { closePop(); placeLinkCard(); placeImageCard(); placeCommentCard(); }, { passive: true });
  // zoom dokumentu (przyciski, suwak, dwa palce): karta i edytor komentarza w tej samej skali
  docCanvasEl?.addEventListener("dwb-zoom", () => {
    if (pop?.el.classList.contains("compose-pop-comment")) pop.el.style.setProperty("--cc-z", String(commentZoom()));
    placeCommentCard();
    placeImageCard();
  });
  // Edycja: klik w nagłówek / stopkę strony w podglądzie = okienko nagłówka i stopki (jak dwuklik w Wordzie)
  docCanvasEl?.addEventListener("click", (e) => {
    if (readOnlyMode || !e.target.closest?.(".docx-preview-host section.docx > header, .docx-preview-host section.docx > footer")) return;
    if (e.target.closest("a")) return;
    openHeaderFooterForm();
  });

  // Ctrl/⌘+Alt+M = komentarz (jak w Wordzie; po e.code — Alt na Macu zmienia znak)
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || !e.altKey || e.shiftKey || e.code !== "KeyM") return;
    if (readOnlyMode || !originalFileBytes || (!docCaretParagraph(e.target) && !lastDocCaret)) return;
    e.preventDefault();
    e.stopPropagation();
    openCommentForm(insertBtn);
  }, true);
  // przewinięcie rzędu paska: okienko jedzie za swoim przyciskiem; znika, gdy przycisk schowa się
  // pod krawędzią (okienko komentarza zostaje zawsze — B / I / U są w tym rzędzie)
  bar?.addEventListener("scroll", () => {
    if (!pop) return;
    const a = pop.anchor.getBoundingClientRect();
    const b = bar.getBoundingClientRect();
    const visible = bar.contains(pop.anchor) ? a.right > b.left + 4 && a.left < b.right - 4 : true;
    if (visible || pop.el.classList.contains("compose-pop-comment")) placePop(pop.el, pop.anchor); else closePop();
  }, { passive: true });
  window.addEventListener("resize", () => closePop());
  // klawiatura ekranowa wysuwa się / chowa: okienko i karty liczą miejsce od nowa (karta ułożona
  // przy klawiaturze była ściśnięta do jednej linijki i taka zostawała)
  window.visualViewport?.addEventListener("resize", () => { if (pop) placePop(pop.el, pop.anchor); placeCommentCard(); placeLinkCard(); placeImageCard(); });

  // Ctrl/⌘+Alt+F / D = przypis dolny / końcowy (jak w Wordzie) — w Edycji z kursorem w tekście;
  // poza tekstem Ctrl/⌘+Alt+F dalej włącza tryb skupienia (keyboard.js)
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || !e.altKey || e.shiftKey || (e.code !== "KeyF" && e.code !== "KeyD")) return;
    if (readOnlyMode || !originalFileBytes || !docCaretParagraph(e.target)) return;
    e.preventDefault();
    e.stopPropagation();
    insertNote(e.code === "KeyF" ? "footnote" : "endnote");
  }, true);

  // Ctrl/⌘+K = link (jak w Wordzie)
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || e.code !== "KeyK") return;
    if (readOnlyMode || !originalFileBytes) return;
    if (!docCaretParagraph(e.target) && !lastDocCaret) return;
    e.preventDefault();
    e.stopPropagation();
    openLinkForm(insertBtn);
  }, true);

  // Ctrl/⌘+Enter w tekście = podział strony (jak w Wordzie). Przed obsługą Entera w akapicie.
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" || !(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || e.isComposing) return;
    if (readOnlyMode || !docCaretParagraph(e.target)) return;
    e.preventDefault();
    e.stopPropagation();
    insertPageBreak();
  }, true);

  // pisanie chowa kartę komentarza; wraca przy ruchu kursora, który NIE jest skutkiem pisania
  // (zmiana zaznaczenia > 150 ms po ostatnim wpisanym znaku: klik, strzałki, kursor po operacji)
  let lastInputAt = 0;
  docCanvasEl?.addEventListener("beforeinput", () => { cCardMuted = true; lastInputAt = performance.now(); hideCommentCard(); }, true);
  docCanvasEl?.addEventListener("input", () => { lastInputAt = performance.now(); }, true);
  const unmute = () => { cCardMuted = false; lastInputAt = 0; };
  docCanvasEl?.addEventListener("pointerdown", unmute, true);
  docCanvasEl?.addEventListener("keydown", (e) => { if (/^(Arrow|Home|End|Page)/.test(e.key)) unmute(); }, true);
  document.addEventListener("selectionchange", () => {
    if (cCardMuted && performance.now() - lastInputAt > 150) cCardMuted = false;
    queueSync();
  });

  dialog?.addEventListener("click", (e) => {
    const card = e.target.closest(".newdoc-card");
    if (card) { createNew(card.dataset.template); return; }
    if (e.target === dialog) dialog.close(); // klik w tło
  });
  // menu ⋯ → Marginesy i układ strony (po zamknięciu menu, z kotwicą „＋ Wstaw” na pasku)
  document.getElementById("pageSetupMenuItem")?.addEventListener("click", () => setTimeout(() => openPageSetup(), 60));
  // „Układ” na pasku Edycji (jak karta Układ w Wordzie): marginesy, orientacja, rozmiar, wyrównanie w pionie
  layoutBtn?.addEventListener("mousedown", (e) => e.preventDefault()); // kursor zostaje w tekście
  layoutBtn?.addEventListener("click", () => { if (pop?.anchor === layoutBtn) closePop(); else openPageSetup(layoutBtn); });
  ["emptyNewBtn", "newDocBtn", "newDocMenuItem"].forEach((id) => {
    document.getElementById(id)?.addEventListener("click", openNewDialog);
  });

  return { openPageSetup, insertNote, openHeaderFooterForm, fixPreviewPageNumbers, pageNumberSelector, openCommentForm, paintCommentHighlights, loadComments, applyColor, insertTable, tableAction, tableTab, insertImageFile, imageEdit, showImageCard, hideImageCard, insertToc, insertFormField, applyList, changeListLevel, endListAt, openLinkForm, removeLink, hideLinkCard, openNewDialog, createNew, applyStyle, applyAlign, insertPageBreak, insertHrule, insertText, syncState };
})();
