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
//   - przyciski w okienkach nie zabierają fokusu z tekstu (pointerdown → preventDefault),
//     więc na iPhonie klawiatura i kursor zostają.

const composeUi = (() => {
  const bar = document.getElementById("formatToolbar");
  const insertBtn = document.getElementById("insertMenuBtn");
  const styleSel = document.getElementById("fmtParaStyle");
  const alignBtn = document.getElementById("fmtAlignBtn");
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
    if (same) return; // drugi klik w ten sam przycisk = zamknij
    const el = document.createElement("div");
    el.className = "compose-pop";
    el.setAttribute("role", "menu");
    build(el);
    // przyciski nie zabierają fokusu z dokumentu (kursor i klawiatura ekranowa zostają)
    el.addEventListener("pointerdown", (e) => { if (e.target.closest("button")) e.preventDefault(); });
    el.addEventListener("keydown", (e) => {
      const items = Array.from(el.querySelectorAll("button"));
      const i = items.indexOf(document.activeElement);
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closePop(true); }
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
    if (document.activeElement === anchor) el.querySelector("button")?.focus();
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
      label: t("insertHrule"), desc: t("insertHruleDesc"),
      icon: ICON('<line x1="3" y1="12" x2="21" y2="12"/>'),
      onPick: insertHrule,
    });
    popCap(el, t("insertGroupText"));
    popItem(el, {
      label: t("insertDate"), desc: todayText(),
      icon: ICON('<rect x="3" y="5" width="18" height="16" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="8" y1="3" x2="8" y2="7"/><line x1="16" y1="3" x2="16" y2="7"/>'),
      onPick: () => insertText(todayText()),
    });
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

  insertBtn?.addEventListener("pointerdown", (e) => e.preventDefault()); // kursor zostaje w tekście
  insertBtn?.addEventListener("click", () => openPop(insertBtn, buildInsertMenu));
  alignBtn?.addEventListener("pointerdown", (e) => e.preventDefault());
  alignBtn?.addEventListener("click", () => openPop(alignBtn, buildAlignMenu));
  styleSel?.addEventListener("change", () => applyStyle(styleSel.value));

  // klik obok / przewinięcie / zmiana rozmiaru — okienko znika
  document.addEventListener("pointerdown", (e) => {
    if (!pop || pop.el.contains(e.target) || pop.anchor.contains(e.target)) return;
    closePop();
  }, true);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && pop) { e.preventDefault(); e.stopPropagation(); closePop(true); }
  }, true);
  docViewportEl?.addEventListener("scroll", () => closePop(), { passive: true });
  bar?.addEventListener("scroll", () => closePop(), { passive: true });
  window.addEventListener("resize", () => closePop());
  window.visualViewport?.addEventListener("resize", () => { if (pop) placePop(pop.el, pop.anchor); });

  // Ctrl/⌘+Enter w tekście = podział strony (jak w Wordzie). Przed obsługą Entera w akapicie.
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" || !(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || e.isComposing) return;
    if (readOnlyMode || !e.target.closest?.(".docx-editable-p")) return;
    e.preventDefault();
    e.stopPropagation();
    insertPageBreak();
  }, true);

  document.addEventListener("selectionchange", queueSync);

  dialog?.addEventListener("click", (e) => {
    const card = e.target.closest(".newdoc-card");
    if (card) { createNew(card.dataset.template); return; }
    if (e.target === dialog) dialog.close(); // klik w tło
  });
  ["emptyNewBtn", "newDocBtn", "newDocMenuItem"].forEach((id) => {
    document.getElementById(id)?.addEventListener("click", openNewDialog);
  });

  return { openNewDialog, createNew, applyStyle, applyAlign, insertPageBreak, insertHrule, insertText, syncState };
})();
