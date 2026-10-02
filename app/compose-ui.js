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

  // Zmiana w pliku + przerysowanie; kursor i miejsce przewinięcia wracają.
  async function runFileEdit(edit, caret) {
    const top = docViewportEl?.scrollTop || 0;
    if (caret) pendingInlineCursor = caret;
    try {
      await applyDocumentEdit(edit);
    } catch (err) {
      log(`Wstaw/format: ${err.message || err}`, "error");
      toast(t("saveFailed"), "error");
    }
    if (docViewportEl) docViewportEl.scrollTop = top;
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
    build(el);
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
      const label = (p?.textContent || "").replace(/\s+/g, " ").trim().slice(0, 50);
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
      const heads = linkHeadings();
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
      heads.forEach((h) => {
        const o = document.createElement("option");
        o.value = String(h.paraIndex);
        o.textContent = `${"  ".repeat(Math.max(0, (h.level || 1) - 1))}${h.label}`.slice(0, 80);
        if (h.paraIndex === targetIdx) o.selected = true;
        sel.appendChild(o);
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
    const top = docViewportEl?.scrollTop || 0;
    await applyDocumentEdit({ op: "table", index, action }).catch((err) => log(`Tabela: ${err.message || err}`, "error"));
    if (docViewportEl) docViewportEl.scrollTop = top;
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
    const hint = document.createElement("p");
    hint.className = "compose-note compose-note-pad";
    hint.textContent = t("tableTabHint");
    el.appendChild(hint);
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
    const baseW = img.getBoundingClientRect().width / (pct0 / 100);
    range.addEventListener("input", () => {
      val.textContent = `${range.value}%`;
      img.style.width = `${(baseW / docZoomScale(img)) * range.value / 100}px`; // podgląd na żywo
      img.style.height = "auto";
      placeImageCard();
    });
    range.addEventListener("change", () => imageEdit(index, { action: "size", widthPct: +range.value }));
    const hint = (b, key) => { b.setAttribute("aria-label", t(key)); b.dataset.hint = ""; b.dataset.hintPl = I18N.pl[key]; b.dataset.hintEn = I18N.en[key]; b.dataset.hintDelay = "0.4"; };
    el.querySelectorAll("[data-align]").forEach((b) => {
      hint(b, ALIGN_KEYS[b.dataset.align]);
      b.innerHTML = alignSvg(b.dataset.align);
      b.addEventListener("click", () => imageEdit(index, { action: "align", align: b.dataset.align }));
    });
    const alt = el.querySelector(".ic-alt");
    hint(alt, "imageAlt");
    alt.innerHTML = '<span class="ic-alt-text">ALT</span>';
    alt.addEventListener("click", () => openAltForm(index, img));
    const del = el.querySelector(".ic-del");
    hint(del, "imageDelete");
    del.innerHTML = ICON('<polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>');
    del.addEventListener("click", () => { hideImageCard(); imageEdit(index, { action: "delete" }, false); });
    el.addEventListener("mousedown", (e) => { if (e.target.closest("button")) e.preventDefault(); });
    document.body.appendChild(el);
    imgCard = { el, img, p, index };
    placeImageCard();
  }
  function placeImageCard() {
    if (!imgCard) return;
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
  async function imageEdit(index, edit, reselect = true) {
    const top = docViewportEl?.scrollTop || 0;
    await applyDocumentEdit({ op: "image", index, ...edit }).catch((err) => log(`Obraz: ${err.message || err}`, "error"));
    if (docViewportEl) docViewportEl.scrollTop = top;
    if (!reselect) return;
    await whenEditable();
    const img = collectPreviewParagraphElements(host())[index]?.querySelector("img");
    if (img && !readOnlyMode) showImageCard(img);
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
    el.focus({ preventScroll: true });
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
      activeTypingStyle = mergeRunStyles(activeTypingStyle || {}, { [kind]: css || undefined });
      if (!css && activeTypingStyle) delete activeTypingStyle[kind];
      restoreDocCaret();
      return;
    }
    const edit = { op: "runStyle", index: ctx.index, start: ctx.start, end: ctx.end };
    edit[kind] = kind === "color" ? (value ? `#${value}` : "auto") : css;
    const top = docViewportEl?.scrollTop || 0;
    await applyDocumentEdit(edit).catch((err) => log(`Kolor: ${err.message || err}`, "error"));
    if (docViewportEl) docViewportEl.scrollTop = top;
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

  function openCommentForm(anchor, opts = {}) {
    if (readOnlyMode) return;
    const ctx = opts.replyTo ? null : selectionInParagraph();
    if (!opts.replyTo && !ctx) return;
    hideCommentCard();
    closePop();
    openPop(anchor || insertBtn, (el) => {
      el.classList.add("compose-pop-form");
      el.setAttribute("role", "dialog");
      el.setAttribute("aria-label", t(opts.replyTo ? "commentReply" : "insertComment"));
      el.innerHTML = `<div class="compose-cap"></div>
        <label class="compose-field"><span></span><textarea rows="3" class="cf-text" data-autofocus="1"></textarea></label>
        <label class="compose-field"><span></span><input type="text" class="cf-author" autocomplete="name" enterkeyhint="done"></label>
        <div class="compose-actions"><span class="compose-actions-gap"></span><button type="button" class="btn lf-cancel"></button><button type="button" class="btn primary lf-ok"></button></div>`;
      el.querySelector(".compose-cap").textContent = t(opts.replyTo ? "commentReply" : "insertComment");
      const [l1, l2] = el.querySelectorAll(".compose-field > span");
      l1.textContent = t(opts.replyTo ? "commentReplyText" : "commentText");
      l2.textContent = t("commentAuthor");
      const ta = el.querySelector(".cf-text");
      const au = el.querySelector(".cf-author");
      au.value = authorName();
      au.placeholder = t("commentAuthorPh");
      el.querySelector(".lf-cancel").textContent = t("linkCancel");
      el.querySelector(".lf-cancel").addEventListener("click", () => closePop());
      el.querySelector(".lf-ok").textContent = t(opts.replyTo ? "commentReplyBtn" : "commentAddBtn");
      el.querySelector(".lf-ok").addEventListener("click", async () => {
        const text = ta.value.trim();
        const author = au.value.trim();
        if (!text) { ta.focus(); return; }
        if (!author) { toast(t("commentNeedAuthor"), "info"); au.focus(); return; }
        try { localStorage.setItem(AUTHOR_KEY, author); } catch (_) { /* prywatne */ }
        closePop();
        if (opts.replyTo) {
          await runFileEdit({ op: "commentReply", id: opts.replyTo, text, author, initials: initialsOf(author) }, null);
          return;
        }
        let { start, end } = ctx;
        if (end <= start) { // bez zaznaczenia — słowo pod kursorem (jak w Wordzie)
          const full = previewRunsToPlainText(extractRunsFromPreviewParagraph(ctx.p));
          while (start > 0 && /\S/.test(full[start - 1])) start--;
          while (end < full.length && /\S/.test(full[end])) end++;
        }
        await runFileEdit({ op: "commentAdd", index: ctx.index, start, end, text, author, initials: initialsOf(author) }, { paraIndex: ctx.index, offset: end });
      });
    });
    if (!authorName()) setTimeout(() => pop?.el.querySelector(".cf-text")?.focus(), 0);
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
  function showCommentCard(hit) {
    const c = commentData.byId.get(hit.id);
    if (!c) return;
    if (cCard?.id === hit.id && cCard.done === c.done) { placeCommentCard(); return; }
    hideCommentCard();
    const el = document.createElement("div");
    el.className = `comment-card${c.done ? " is-done" : ""}`;
    el.setAttribute("role", "region");
    el.setAttribute("aria-label", t("commentCard"));
    const entry = (x) => {
      const box = document.createElement("div");
      box.className = "cc-entry";
      const head = document.createElement("div");
      head.className = "cc-head";
      head.textContent = [x.author, fmtDate(x.date)].filter(Boolean).join(" · ");
      const body = document.createElement("div");
      body.className = "cc-text";
      body.textContent = x.text;
      box.append(head, body);
      return box;
    };
    const list = document.createElement("div");
    list.className = "cc-list";
    list.append(entry(c), ...c.replies.map(entry));
    const actions = document.createElement("div");
    actions.className = "cc-actions";
    const mk = (key, fn, cls = "") => { const b = document.createElement("button"); b.type = "button"; b.className = `btn ${cls}`.trim(); b.textContent = t(key); b.addEventListener("click", fn); return b; };
    actions.append(
      mk("commentReply", () => openCommentForm(actions.firstChild, { replyTo: c.id })),
      mk(c.done ? "commentReopen" : "commentResolve", () => runFileEdit({ op: "commentDone", id: c.id, done: !c.done }, null)),
      mk("commentDelete", () => { hideCommentCard(); applyDocumentEdit({ op: "revisions", action: "removeComments", ids: [c.id, ...c.replies.map((r) => r.id)] }); }, "cc-del"),
    );
    if (c.done) { const badge = document.createElement("div"); badge.className = "cc-done"; badge.textContent = t("commentResolved"); el.appendChild(badge); }
    el.append(list, actions);
    el.addEventListener("mousedown", (e) => { if (e.target.closest("button")) e.preventDefault(); });
    document.body.appendChild(el);
    cCard = { el, id: hit.id, done: c.done };
    placeCommentCard();
  }
  function placeCommentCard() {
    if (!cCard) return;
    const sel = window.getSelection();
    if (!sel?.rangeCount) return;
    const rects = sel.getRangeAt(0).getClientRects();
    const r = rects[0] || sel.getRangeAt(0).getBoundingClientRect();
    const vv = window.visualViewport;
    const vp = docViewportEl?.getBoundingClientRect();
    if (!r || (vp && (r.bottom < vp.top || r.top > vp.bottom))) { cCard.el.style.visibility = "hidden"; return; }
    cCard.el.style.visibility = "";
    const viewW = vv ? vv.width : window.innerWidth;
    const w = cCard.el.offsetWidth; const h = cCard.el.offsetHeight;
    const left = Math.max(8, Math.min(r.left - 16, viewW - w - 8));
    let top = r.top - h - 8;
    if (top < (vp?.top ?? 8)) top = r.bottom + 40; // brak miejsca nad wierszem — pod nim (z zapasem na kartę linku)
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

  function fixPreviewPageNumbers() {
    const cls = [...docComposeStyleClasses].filter(([, k]) => k === "pagenum").map(([c]) => c);
    if (!cls.length) return;
    const sel = cls.map((c) => `span.${CSS.escape(c)}`).join(",");
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
  // kolor „do dalszego pisania” obowiązuje do kliknięcia w inne miejsce (jak w Wordzie)
  docCanvasEl?.addEventListener("pointerdown", () => { if (activeTypingStyle) { delete activeTypingStyle.color; delete activeTypingStyle.highlight; } }, true);
  tableBtn?.addEventListener("mousedown", (e) => e.preventDefault());
  tableBtn?.addEventListener("click", () => openPop(tableBtn, buildTableMenu));

  // Obraz: klik w Edycji pokazuje kartę; klik obok / Esc / tryb Czytanie chowa.
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
  document.getElementById("readMode")?.addEventListener("change", () => { hideImageCard(); if (tableBtn) tableBtn.hidden = true; });

  // Wklejony obraz (zrzut ekranu, skopiowane zdjęcie) — w miejscu kursora, jak „Wstaw → Obraz”.
  docCanvasEl?.addEventListener("paste", (e) => {
    const p = e.target?.closest?.(".docx-editable-p");
    if (!p || readOnlyMode) return;
    const file = Array.from(e.clipboardData?.files || []).find((f) => /^image\//.test(f.type));
    if (!file) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    insertImageFile(file, p);
  }, true);

  // klik obok / przewinięcie / zmiana rozmiaru — okienko znika
  document.addEventListener("pointerdown", (e) => {
    if (!pop || pop.el.contains(e.target) || pop.anchor.contains(e.target)) return;
    closePop();
  }, true);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && pop) { e.preventDefault(); e.stopPropagation(); closePop(true); }
  }, true);
  docViewportEl?.addEventListener("scroll", () => { closePop(); placeLinkCard(); placeImageCard(); placeCommentCard(); }, { passive: true });
  // Edycja: klik w nagłówek / stopkę strony w podglądzie = okienko nagłówka i stopki (jak dwuklik w Wordzie)
  docCanvasEl?.addEventListener("click", (e) => {
    if (readOnlyMode || !e.target.closest?.(".docx-preview-host section.docx > header, .docx-preview-host section.docx > footer")) return;
    if (e.target.closest("a")) return;
    openHeaderFooterForm();
  });

  // Ctrl/⌘+Alt+M = komentarz (jak w Wordzie; po e.code — Alt na Macu zmienia znak)
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || !e.altKey || e.shiftKey || e.code !== "KeyM") return;
    if (readOnlyMode || !originalFileBytes || (!e.target.closest?.(".docx-editable-p") && !lastDocCaret)) return;
    e.preventDefault();
    e.stopPropagation();
    openCommentForm(insertBtn);
  }, true);
  bar?.addEventListener("scroll", () => closePop(), { passive: true });
  window.addEventListener("resize", () => closePop());
  window.visualViewport?.addEventListener("resize", () => { if (pop) placePop(pop.el, pop.anchor); });

  // Ctrl/⌘+K = link (jak w Wordzie)
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || e.code !== "KeyK") return;
    if (readOnlyMode || !originalFileBytes) return;
    if (!e.target.closest?.(".docx-editable-p") && !lastDocCaret) return;
    e.preventDefault();
    e.stopPropagation();
    openLinkForm(insertBtn);
  }, true);

  // Ctrl/⌘+Enter w tekście = podział strony (jak w Wordzie). Przed obsługą Entera w akapicie.
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" || !(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || e.isComposing) return;
    if (readOnlyMode || !e.target.closest?.(".docx-editable-p")) return;
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
  ["emptyNewBtn", "newDocBtn", "newDocMenuItem"].forEach((id) => {
    document.getElementById(id)?.addEventListener("click", openNewDialog);
  });

  return { openHeaderFooterForm, fixPreviewPageNumbers, openCommentForm, paintCommentHighlights, loadComments, applyColor, insertTable, tableAction, tableTab, insertImageFile, imageEdit, showImageCard, hideImageCard, insertToc, insertFormField, applyList, changeListLevel, endListAt, openLinkForm, removeLink, hideLinkCard, openNewDialog, createNew, applyStyle, applyAlign, insertPageBreak, insertHrule, insertText, syncState };
})();
