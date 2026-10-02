// snippet-suggest.js — wstawianie snippetów w podglądzie: podpowiedzi po „!”, pytanie o pola
// {{…}}, kursor w miejscu {cursor}, jeden krok cofania.
//
// Wszystkie drogi (lista podpowiedzi, tryb auto, „Wstaw” w panelu) idą przez
// expandSnippetAtCaret — dzięki temu działają tak samo.
//
// Podpowiedzi: po wpisaniu „!” (na początku słowa) pod kursorem pojawia się lista pasujących
// snippetów z podglądem treści. ↑ ↓ wybór, Enter/Tab wstawia, Esc zamyka; na dotyku — stuknięcie.

(() => {
  const MAX_ITEMS = 6;
  const TRIGGER_BEFORE_CARET_RE = /(?<![\p{L}\p{N}_!])!([\p{L}\p{N}_-]*)$/u;

  // ── pytanie o pola {{…}} ────────────────────────────────────────────────────
  function askSnippetFields(sn, fields) {
    return new Promise((resolve) => {
      const remembered = loadSnippetFieldValues();
      const dlg = document.createElement("dialog");
      dlg.className = "sn-dialog";
      const form = document.createElement("form");
      form.method = "dialog";
      const title = document.createElement("h3");
      title.textContent = t("snippetsFieldsTitle", { name: formatSnippetTrigger(sn.name) });
      form.append(title);
      const inputs = fields.map((name, i) => {
        const label = document.createElement("label");
        label.className = "field";
        const span = document.createElement("span");
        span.textContent = typeof placeholderFieldLabel === "function" ? placeholderFieldLabel(name) : name;
        const input = document.createElement("input");
        input.type = "text";
        input.autocomplete = "off";
        input.value = remembered[name] || "";
        input.dataset.name = name;
        if (!i) input.autofocus = true;
        label.append(span, input);
        form.append(label);
        return input;
      });
      const hint = document.createElement("p");
      hint.className = "hint";
      hint.textContent = t("snippetsFieldsHint");
      const row = document.createElement("div");
      row.className = "btn-row";
      const cancel = Object.assign(document.createElement("button"), { type: "button", className: "btn", textContent: t("cancel") });
      const ok = Object.assign(document.createElement("button"), { type: "submit", className: "btn primary", textContent: t("snippetsFieldsInsert") });
      row.append(cancel, ok);
      form.append(hint, row);
      dlg.append(form);
      document.body.append(dlg);
      let done = false;
      const finish = (values) => { if (done) return; done = true; dlg.close(); dlg.remove(); resolve(values); };
      cancel.addEventListener("click", () => finish(null));
      dlg.addEventListener("cancel", (e) => { e.preventDefault(); finish(null); }); // Esc
      form.addEventListener("submit", (e) => {
        e.preventDefault();
        const values = {};
        inputs.forEach((inp) => { if (inp.value.trim()) values[inp.dataset.name] = inp.value.trim(); });
        saveSnippetFieldValues(values);
        finish(values);
      });
      dlg.showModal();
      inputs[0]?.focus();
    });
  }

  function fillFields(text, values) {
    return text.replace(new RegExp(PLACEHOLDER_TOKEN_RE.source, "g"), (tok, name) => (values[name] != null ? values[name] : tok));
  }

  // Kursor tam, gdzie był {cursor} — znacznik usuwamy i stawiamy w jego miejscu zaznaczenie.
  function placeCursorMark(p) {
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = w.nextNode())) {
      const at = n.textContent.indexOf(SNIPPET_CURSOR_MARK);
      if (at < 0) continue;
      n.textContent = n.textContent.slice(0, at) + n.textContent.slice(at + 1);
      const r = document.createRange();
      r.setStart(n, at);
      r.collapse(true);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
      return true;
    }
    return false;
  }

  // Wstaw snippet w akapicie p w miejscu kursora; deleteLen — ile znaków przed kursorem zastąpić
  // (wpisany „!nazwa”), trailing — znak, który wywołał rozwinięcie w trybie auto (spacja, kropka).
  async function expandSnippetAtCaret(p, sn, deleteLen = 0, trailing = "", lead = "") {
    if (!p || !sn) return false;
    let body = lead + resolveSnippetBody(sn.body, { cursorMark: true });
    const fields = [...new Set(scanPlaceholdersInText(body).map((h) => h.name))];
    if (fields.length) {
      const saved = window.getSelection()?.rangeCount ? window.getSelection().getRangeAt(0).cloneRange() : null;
      const values = await askSnippetFields(sn, fields);
      if (!values) { p.focus(); if (saved) { const s = getSelection(); s.removeAllRanges(); s.addRange(saved); } return false; }
      body = fillFields(body, values);
      p.focus({ preventScroll: true });
      if (saved) { const s = getSelection(); s.removeAllRanges(); s.addRange(saved); }
    }
    const style = mergeRunStyles(getInheritedRunStyleAtCaret(p), activeTypingStyle);
    const hasMark = body.includes(SNIPPET_CURSOR_MARK);
    asUndoStep("undoOpSnippet", () => {
      if (deleteLen > 0) replaceTextEndingBeforeCaret(p, deleteLen, body + trailing, style);
      else if (runStyleHasProps(style)) insertStyledTextAtCaret(body + trailing, style, p);
      else insertTextAtCaret(body + trailing);
      if (hasMark) placeCursorMark(p);
    });
    onInlineParagraphInput();
    return true;
  }
  window.expandSnippetAtCaret = expandSnippetAtCaret;

  // ── iOS: spacja zjedzona przed „!” ─────────────────────────────────────────
  // Klawiatura iPhone'a po słowie z paska podpowiedzi sama dokleja spację, a gdy wpiszesz
  // znak interpunkcyjny („!”), kasuje ją („Dobrze !” → „Dobrze!”) — w Notatkach tak samo.
  // Dla snippetów to psuło wyzwalacz (zgłoszenie Mateusza 2026-10-02: trzeba było dawać dwie
  // spacje). Zapamiętujemy, że iOS zjadł spację tuż przed „!”; jeśli zaraz potem pisana jest
  // nazwa (litera po „!” — wykrzyknik kończący zdanie nie ma litery bez spacji), to wyzwalacz,
  // a przy wstawieniu snippetu spacja wraca. Samo „Dobrze!” zostaje bez zmian.
  let eatenSpace = null; // { p, index: pozycja „!” w tekście akapitu }
  // iOS bywa, że robi to jedną zmianą („ ” → „!”), a bywa, że dwiema (skasuj spację, wstaw „!”) —
  // dlatego krótka historia stanów tekstu przed kursorem (ostatnie zmiany z ~0,6 s).
  const recent = []; // { p, text, at }
  docCanvasEl?.addEventListener("beforeinput", (e) => {
    const p = e.target?.closest?.(".docx-editable-p");
    if (!p) return;
    recent.push({ p, text: getTextBeforeCaret(p), at: performance.now() });
    if (recent.length > 4) recent.shift();
  }, true);
  docCanvasEl?.addEventListener("input", (e) => {
    const p = e.target?.closest?.(".docx-editable-p");
    if (!p) return;
    const now = getTextBeforeCaret(p);
    if (!/[^\s!]!$/u.test(now)) return; // „słowo!” — tylko wtedy coś mogło zjeść spację
    const want = `${now.slice(0, -1)} `;
    const t0 = performance.now() - 600;
    if (recent.some((r) => r.p === p && r.at >= t0 && (r.text === want || r.text === `${now.slice(0, -1)}\u00a0`))) eatenSpace = { p, index: now.length - 1 };
  }, true);
  // Nazwa snippetu pisana za „!”, przed którym iOS zjadł spację: zwraca nazwę albo null.
  function eatenSpaceQuery(p, before) {
    if (!eatenSpace || eatenSpace.p !== p || before.length <= eatenSpace.index || before[eatenSpace.index] !== "!") return null;
    const m = before.slice(eatenSpace.index).match(/^!([\p{L}\p{N}_-]*)$/u);
    return m ? m[1] : null;
  }
  window.snippetEatenSpaceQuery = eatenSpaceQuery;

  // ── podpowiedzi po „!” ──────────────────────────────────────────────────────
  const box = document.createElement("div");
  box.className = "sn-suggest";
  box.setAttribute("role", "listbox");
  box.hidden = true;
  document.body.append(box);
  let state = null; // { p, query, items, active }

  function swallowGhostClick(x, y) {
    const until = Date.now() + 700;
    const guard = (ev) => {
      if (Date.now() > until) { document.removeEventListener("click", guard, true); return; }
      if (Math.abs(ev.clientX - x) < 30 && Math.abs(ev.clientY - y) < 30) {
        ev.preventDefault();
        ev.stopImmediatePropagation();
        document.removeEventListener("click", guard, true);
      }
    };
    document.addEventListener("click", guard, true);
    setTimeout(() => document.removeEventListener("click", guard, true), 800);
  }

  function close() {
    state = null;
    box.hidden = true;
    box.replaceChildren();
  }

  function preview(body) {
    const one = resolveSnippetBody(body).replace(/\s*\n\s*/g, " ⏎ ");
    return one.length > 60 ? `${one.slice(0, 59)}…` : one;
  }

  function matches(query) {
    const q = query.toLocaleLowerCase("pl-PL");
    const list = loadSnippets();
    const starts = list.filter((s) => s.name.toLocaleLowerCase("pl-PL").startsWith(q));
    const contains = q ? list.filter((s) => !starts.includes(s) && s.name.toLocaleLowerCase("pl-PL").includes(q)) : [];
    return [...starts, ...contains].slice(0, MAX_ITEMS);
  }

  function render() {
    box.replaceChildren();
    state.items.forEach((sn, i) => {
      const item = document.createElement("div");
      item.className = `sn-suggest-item${i === state.active ? " is-active" : ""}`;
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", String(i === state.active));
      const name = document.createElement("strong");
      name.textContent = formatSnippetTrigger(sn.name);
      const pv = document.createElement("span");
      pv.textContent = preview(sn.body);
      item.append(name, pv);
      // pointerdown/mousedown + preventDefault: akapit nie traci fokusu (klawiatura zostaje).
      // Wybór dopiero na puszczeniu palca — i połykamy „kliknięcie”, które iOS wysyła chwilę
      // później w TO SAMO miejsce: lista już znika, więc trafiało w przycisk pod nią
      // (na telefonie ↶ na pasku — snippet wstawiał się i od razu cofał).
      item.addEventListener("pointerdown", (e) => e.preventDefault());
      item.addEventListener("mousedown", (e) => e.preventDefault());
      item.addEventListener("pointerup", (e) => { e.preventDefault(); swallowGhostClick(e.clientX, e.clientY); accept(i); });
      box.append(item);
    });
    const hint = document.createElement("div");
    hint.className = "sn-suggest-hint";
    hint.textContent = t("snippetsSuggestHint");
    box.append(hint);
  }

  function position() {
    const sel = window.getSelection();
    if (!sel?.rangeCount) return;
    const r = sel.getRangeAt(0).getClientRects()[0] || state.p.getBoundingClientRect();
    const vv = window.visualViewport;
    const vh = vv ? vv.height + vv.offsetTop : window.innerHeight;
    box.hidden = false;
    const h = box.offsetHeight;
    const w = box.offsetWidth;
    const top = r.bottom + 6 + h > vh - 8 ? Math.max(8, r.top - 6 - h) : r.bottom + 6;
    box.style.top = `${top}px`;
    box.style.left = `${Math.max(8, Math.min(r.left - 8, window.innerWidth - w - 8))}px`;
  }

  function update() {
    if (readOnlyMode) { close(); return; }
    const sel = window.getSelection();
    const node = sel?.anchorNode;
    const el = node?.nodeType === 1 ? node : node?.parentElement;
    const p = el?.closest?.(".docx-editable-p");
    if (!p || !sel.isCollapsed) { close(); return; }
    const before = getTextBeforeCaret(p);
    const m = before.match(TRIGGER_BEFORE_CARET_RE);
    // iOS zjadł spację przed „!”, a za nim jest nazwa (co najmniej jedna litera) — też wyzwalacz
    const eaten = !m ? eatenSpaceQuery(p, before) : null;
    if (!m && !eaten) { close(); return; }
    const query = m ? m[1] : eaten;
    const items = matches(query);
    if (!items.length) { close(); return; }
    const same = state && state.p === p && state.query === query;
    state = { p, query, items, active: same ? Math.min(state.active, items.length - 1) : 0, lead: m ? "" : " " };
    render();
    position();
  }

  async function accept(i) {
    if (!state) return;
    const { p, query, items, lead } = state;
    const sn = items[i];
    close();
    eatenSpace = null;
    await expandSnippetAtCaret(p, sn, query.length + 1, "", lead || "");
  }

  docCanvasEl?.addEventListener("input", () => requestAnimationFrame(update));
  document.addEventListener("selectionchange", () => { if (state) requestAnimationFrame(update); });
  docCanvasEl?.addEventListener("focusout", () => setTimeout(() => { if (!docCanvasEl.contains(document.activeElement)) close(); }, 0));
  docViewportEl?.addEventListener("scroll", () => { if (state) position(); }, { passive: true });

  // przed obsługą Enter (nowy akapit) i przed zapisem kroku cofania „Pisanie”
  docCanvasEl?.addEventListener("keydown", (e) => {
    if (!state || e.isComposing) return;
    const k = e.key;
    if (k === "ArrowDown" || k === "ArrowUp") {
      state.active = (state.active + (k === "ArrowDown" ? 1 : -1) + state.items.length) % state.items.length;
      render();
      position();
    } else if (k === "Enter" || k === "Tab") {
      accept(state.active);
    } else if (k === "Escape") {
      close();
    } else return;
    e.preventDefault();
    e.stopImmediatePropagation();
  }, true);
})();
