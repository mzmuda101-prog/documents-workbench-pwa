// doc-rewrite.js — „✦ Przepisz” na pasku Edycji: asystent AI proponuje nową wersję zaznaczenia
// (także przez kilka akapitów) albo akapitu z kursorem, a zmiana trafia do dokumentu jako
// POPRAWKA Worda (w:ins / w:del, autor „Asystent AI”) — do przejrzenia w Recenzji: Akceptuj /
// Odrzuć. Jeden krok Cofnij. „Popraw błędy” (słownik, bez AI) działa też na cały dokument.
//
// Zasady: podgląd różnicy przed wstawieniem; liczby z oryginału muszą zostać (inaczej ostrzeżenie);
// akapity zablokowane (przypis, pole, link, obraz) — bez zmian, jak przy zwykłym pisaniu.
// Kilka akapitów = każdy osobno do modelu: mały model (0.8B na telefonie) przy dłuższym tekście
// zlewał / dzielił akapity i gubił liczby; osobno układ dokumentu zostaje 1:1, a postęp widać.

const dwbRewrite = (() => {
  const MODES = ["fix", "simple", "formal", "short", "polite"];
  const PROMPTS = {
    fix: "Popraw błędy ortograficzne, interpunkcyjne i gramatyczne. Nie zmieniaj stylu ani treści.",
    simple: "Przepisz prostszym, zrozumiałym językiem, zachowując sens i wszystkie fakty.",
    formal: "Przepisz w stylu formalnym, urzędowym, zachowując sens i wszystkie fakty.",
    short: "Skróć, zachowując sens, wszystkie liczby, daty i nazwy.",
    polite: "Przepisz uprzejmiej, zachowując sens i wszystkie fakty.",
  };
  const SYSTEM =
    "Jesteś redaktorem tekstu. Dostajesz fragment dokumentu i polecenie. Odpowiadasz WYŁĄCZNIE nową wersją " +
    "fragmentu — bez wstępu, komentarza, cudzysłowów i formatowania Markdown. Zachowujesz język oryginału, " +
    "wszystkie liczby, kwoty, daty, nazwy własne i odwołania (np. § 3).";
  const AUTHOR = () => t("rewriteAuthor");
  const STRIP = (s) => String(s || "").replace(/﻿/g, ""); // znaki-pomocniki kursora z podglądu

  let dialog = null;
  // { scope: "sel" | "doc", selItems: [...], items: [{ paraIndex, start, end, original, paraText }], skipped, note }
  let target = null;
  let results = []; // [{ item, suggestion, warn, on, done }]
  let state = null; // { instruction, source: "ai" | "dict", busy, stopped, current }
  let abort = null;

  function plural(n, one, few, many) {
    if (n === 1) return one;
    const m10 = n % 10, m100 = n % 100;
    return m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14) ? few : many;
  }

  // ── co przepisujemy ───────────────────────────────────────────────────────
  // Kawałki zaznaczenia: akapit podglądu, zaznaczony w nim tekst i jego miejsce. Pierwszy
  // i ostatni akapit zwykle częściowo, środkowe w całości.
  function selectionPieces(range) {
    const list = collectPreviewParagraphElements(docCanvasEl?.querySelector(".docx-preview-host"));
    const pieces = list.filter((p) => { try { return range.intersectsNode(p); } catch (_) { return false; } }).map((p) => {
      const r = range.cloneRange();
      if (!p.contains(range.startContainer)) r.setStart(p, 0);
      if (!p.contains(range.endContainer)) r.setEnd(p, p.childNodes.length);
      const pre = document.createRange();
      pre.selectNodeContents(p);
      pre.setEnd(r.startContainer, r.startOffset);
      const post = document.createRange();
      post.selectNodeContents(p);
      post.setStart(r.endContainer, r.endOffset);
      const before = STRIP(pre.toString());
      return { p, text: STRIP(r.toString()), offset: before.length, whole: !before.trim() && !STRIP(post.toString()).trim() };
    });
    // trzykrotne kliknięcie kończy zaznaczenie na samym początku NASTĘPNEGO akapitu — ten się nie liczy
    while (pieces.length > 1 && !pieces[pieces.length - 1].text.trim()) pieces.pop();
    while (pieces.length > 1 && !pieces[0].text.trim()) pieces.shift();
    return pieces;
  }

  async function captureTarget() {
    if (readOnlyMode || !originalFileBytes) return { error: "rewriteNoDoc" };
    const p = restoreDocCaret?.();
    const sel = window.getSelection();
    const range = sel?.rangeCount ? sel.getRangeAt(0) : null;
    let pieces = p && range && !range.collapsed ? selectionPieces(range) : [];
    if (p && !pieces.length) pieces = [{ p, text: "", offset: 0, whole: true }];
    const locked = pieces.filter((x) => x.p.dataset.lock && STRIP(x.p.textContent).trim());
    if (pieces.length === 1 && locked.length === 1) return { error: "rewriteLocked" };
    // indeksy z podglądu PRZED scaleniem edycji (tak jak przy jednym akapicie)
    const usable = pieces.filter((x) => !x.p.dataset.lock).map((x) => ({ ...x, paraIndex: resolveParaIndex(x.p), list: lowerOk(x.p) })).filter((x) => x.paraIndex >= 0);
    await mergeInlineEditsIntoBytes();
    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    const items = [];
    let fallback = false;
    for (const x of usable) {
      const paraText = texts[x.paraIndex];
      if (paraText == null || !paraText.trim() || items.some((it) => it.paraIndex === x.paraIndex)) continue;
      const whole = { paraIndex: x.paraIndex, start: 0, end: paraText.length, original: paraText, paraText, list: x.list };
      if (x.whole || !x.text.trim()) { items.push(whole); continue; }
      // zaznaczenie w tekście z pliku: najbliższe wystąpienie wokół miejsca z podglądu
      let start = paraText.indexOf(x.text, Math.max(0, x.offset - 4));
      if (start < 0) start = paraText.indexOf(x.text);
      if (start < 0) { fallback = true; items.push(whole); continue; }
      items.push({ paraIndex: x.paraIndex, start, end: start + x.text.length, original: x.text, paraText, list: x.list && start === 0 });
    }
    let note = "";
    if (fallback) note = t("rewriteWholeFallback");
    else if (p && !items.length) note = t(locked.length ? "rewriteSkipped" : "rewriteEmpty", { n: locked.length });
    // bez kursora (albo nic do przepisania) okno otwiera się na „Cały dokument” — tam działa „Popraw błędy”
    const hadSel = pieces.some((x) => x.text.trim());
    return { scope: items.length ? "sel" : "doc", selItems: items, items, skipped: locked.length, note, hadSel };
  }

  // Punkt listy albo komórka tabeli: mała litera na początku jest tam poprawna („pasek nad dokumentem”).
  const lowerOk = (p) => isListParagraph(p) || !!p.closest("td, th");

  // Cały dokument: każdy akapit treści z tekstem (bez nagłówków, stopek i przypisów — jak Ctrl+A).
  async function documentItems() {
    await mergeInlineEditsIntoBytes();
    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    const list = collectPreviewParagraphElements(docCanvasEl?.querySelector(".docx-preview-host"));
    const items = [];
    let skipped = 0;
    list.forEach((p, i) => {
      const paraText = texts[i];
      if (paraText == null || !paraText.trim()) return;
      if (p.dataset.lock) { skipped++; return; }
      items.push({ paraIndex: i, start: 0, end: paraText.length, original: paraText, paraText, list: lowerOk(p) });
    });
    return { items, skipped };
  }

  // ── odpowiedź modelu → czysty tekst ───────────────────────────────────────
  function clean(answer, original) {
    let s = String(answer || "").trim();
    s = s.replace(/^```[a-z]*\n?|```$/gi, "").trim();
    s = s.replace(/^(oto|poniżej|propozycja|wersja)[^:\n]{0,60}:\s*/i, "");
    s = s.replace(/^["„“”«»']+|["„“”«»']+$/g, "").trim();
    // zaznaczenie jednym wierszem — model nie dokłada akapitów
    if (!/\n/.test(original)) s = s.replace(/\s*\n+\s*/g, " ");
    return s;
  }
  function numbersOf(text) {
    return (String(text).match(/\d[\d\s .,]*\d|\d/g) || []).map((n) => n.replace(/[\s ]/g, ""));
  }
  function missingNumbers(original, rewritten) {
    const have = new Set(numbersOf(rewritten));
    return [...new Set(numbersOf(original))].filter((n) => !have.has(n));
  }

  // ── „Popraw błędy” bez modelu: słownik Hunspell (spell-check.js) + reguły Korekty ─────────
  // Działa od razu, także na telefonie bez pobranego asystenta. Zamieniamy tylko pewne
  // przypadki: brak polskich znaków („prosze” → „proszę”) albo jedna litera różnicy.
  const fold = (w) => String(w).normalize("NFD").replace(/\p{M}/gu, "").replace(/ł/g, "l").replace(/Ł/g, "L").toLocaleLowerCase("pl");
  function oneEdit(a, b) {
    a = a.toLocaleLowerCase("pl"); b = b.toLocaleLowerCase("pl");
    if (a === b || Math.abs(a.length - b.length) > 1) return a === b;
    let i = 0; let j = 0; let e = 0;
    while (i < a.length && j < b.length) {
      if (a[i] === b[j]) { i++; j++; continue; }
      if (++e > 1) return false;
      if (a.length > b.length) i++; else if (b.length > a.length) j++; else { i++; j++; }
    }
    return e + (a.length - i) + (b.length - j) <= 1;
  }
  // Wyraz będący częścią „kodu”, nie zdania: rozszerzenie („.docx” → „.doc”), pole szablonu
  // („{{wynajmujacy}}”), skrót snippetu („!dzis”), adres, #tag, nazwa_z_podkreślnikiem.
  function isCode(text, o) {
    return /[.!{_@/#\\$]/.test(text[o.start - 1] || "") || /[}_@/]/.test(text[o.end] || "");
  }
  // Wiele tekstów naraz (cały dokument): jeden skan słownikiem, podpowiedzi partiami z postępem.
  // lists[k]: akapit k to punkt listy / komórka tabeli — mała litera na początku zostaje
  async function offlineFixMany(texts, { lists = [], onProgress } = {}) {
    if (typeof ensureLazyFeature === "function") await ensureLazyFeature("grammar");
    const out = texts.slice();
    if (typeof dwbSpell !== "undefined") {
      const found = await dwbSpell.scan(texts, "pl");
      // wyraz z wielkiej litery: Hunspell podsuwa najpierw nazwy własne („Rosze”) — pytamy też o małą
      const ask = [...new Set(found.flatMap((e) => [e.word, e.word.toLocaleLowerCase("pl")]))];
      const sugg = {};
      for (let i = 0; i < ask.length; i += 40) {
        Object.assign(sugg, await dwbSpell.suggest(ask.slice(i, i + 40), "pl"));
        onProgress?.(Math.min(ask.length, i + 40), ask.length);
      }
      const edits = texts.map(() => []);
      for (const e of found) {
        const low = e.word.toLocaleLowerCase("pl");
        const list = [...(sugg[low] || []), ...(sugg[e.word] || [])];
        // pewne: różnica tylko w polskich znakach, albo jedna litera po ich zdjęciu („dokumęty” → „dokumenty”)
        // Bez: samej zmiany wielkości liter („Edge” → „EDGE”), rozbijania wyrazu („Markdown” → „Mark down”),
        // a w wyrazie z wielkiej litery (nazwa, obce słowo) tylko polskie znaki („Letter” ≠ „Getter”).
        const sure = list.filter((x) => !/[\s-]/.test(x) && x.toLocaleLowerCase("pl") !== low);
        const pick = sure.find((x) => fold(x) === fold(e.word))
          || (e.word === low ? sure.find((x) => x === x.toLocaleLowerCase("pl") && oneEdit(fold(x), fold(e.word))) : undefined);
        if (!pick) continue;
        e.occurrences.forEach((o) => { if (!isCode(texts[o.paraIndex], o)) edits[o.paraIndex]?.push({ ...o, word: dwbSpell.matchCase(e.word, pick) }); });
      }
      edits.forEach((list, k) => {
        list.sort((a, b) => b.start - a.start).forEach((o) => { out[k] = out[k].slice(0, o.start) + o.word + out[k].slice(o.end); });
      });
    }
    if (typeof fixParagraphWithRules === "function" && typeof getEnabledGrammarRules === "function") {
      const rules = getEnabledGrammarRules({ lang: "pl" });
      for (let k = 0; k < out.length; k++) {
        const fixed = fixParagraphWithRules(out[k], rules, {});
        out[k] = lists[k] && /^\p{Ll}/u.test(out[k]) ? out[k][0] + fixed.slice(1) : fixed;
      }
    }
    return out;
  }
  async function offlineFix(text) {
    return (await offlineFixMany([text]))[0];
  }

  // ── okno ──────────────────────────────────────────────────────────────────
  function el(tag, cls, text) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }
  function button(cls, text, onClick) {
    const b = el("button", cls, text);
    b.type = "button";
    b.addEventListener("click", onClick);
    return b;
  }

  function ensureDialog() {
    if (dialog) return dialog;
    dialog = el("dialog", "lai-dialog rw-dialog");
    dialog.setAttribute("aria-labelledby", "rwTitle");
    dialog.addEventListener("close", () => { abort?.abort(); abort = null; });
    dialog.addEventListener("click", (e) => { if (e.target === dialog) dialog.close(); });
    document.body.appendChild(dialog);
    return dialog;
  }
  function titleBox() {
    const box = el("div", "lai-box");
    const h = el("h2", "lai-title", t("rewriteTitle"));
    h.id = "rwTitle";
    box.append(h);
    return box;
  }
  const clip = (s, n = 400) => (s.length > n ? `${s.slice(0, n)}…` : s);

  function renderChooser(note = "") {
    const d = ensureDialog();
    d.replaceChildren();
    const box = titleBox();
    const doc = target.scope === "doc";
    // zakres: zaznaczenie / akapit z kursorem albo cały dokument
    const scope = el("div", "rw-scope");
    scope.setAttribute("role", "group");
    scope.setAttribute("aria-label", t("rewriteScopeAria"));
    const n = target.selItems.length;
    const selLabel = n > 1
      ? t("rewriteScopeMulti", { n, word: (typeof currentLang === "string" && currentLang === "en") ? "paragraphs" : plural(n, "akapit", "akapity", "akapitów") })
      : t(target.hadSel ? "rewriteScopeSel" : "rewriteScopePara");
    [["sel", selLabel], ["doc", t("rewriteScopeDoc")]].forEach(([key, label]) => {
      const b = button("rw-scope-opt", label, () => {
        if (target.scope === key) return;
        target.scope = key;
        target.items = key === "sel" ? target.selItems : [];
        renderChooser();
      });
      b.dataset.scope = key;
      b.setAttribute("aria-pressed", String(target.scope === key));
      if (key === "sel" && !n) b.disabled = true;
      scope.append(b);
    });
    box.append(scope);
    box.append(el("p", "rw-original", doc ? t("rewriteDocDesc") : clip(target.selItems.map((it) => it.original).join("\n"))));
    if (note) box.append(el("p", "lai-note", note));
    if (!doc && target.skipped) box.append(el("p", "lai-note", t("rewriteSkipped", { n: target.skipped })));
    if (!doc && n > 1) box.append(el("p", "lai-note", t("rewriteMultiNote")));
    if (doc) box.append(el("p", "lai-note", t("rewriteDocOnlyFix")));
    else if (typeof LocalAI !== "undefined" && !LocalAI.isReady()) box.append(el("p", "lai-note", t("rewriteOfflineHint")));
    const modes = el("div", "rw-modes");
    const aiReady = typeof LocalAI !== "undefined" && LocalAI.isReady();
    MODES.forEach((m) => {
      const b = button("lai-btn", t(`rewriteMode_${m}`), () => run(PROMPTS[m], m));
      b.dataset.mode = m;
      // bez asystenta: „Popraw błędy” działa słownikiem, reszta potrzebuje AI (✦)
      if (!aiReady && m !== "fix") b.append(el("span", "rw-needs-ai", " ✦"));
      if (doc && m !== "fix") b.disabled = true;
      modes.append(b);
    });
    box.append(modes);
    const own = el("div", "rw-own");
    const input = el("input");
    input.type = "text";
    input.placeholder = t("rewriteOwnPlaceholder");
    input.setAttribute("aria-label", t("rewriteOwnPlaceholder"));
    const runOwn = () => { if (input.value.trim()) run(input.value.trim()); };
    const go = button("lai-btn is-primary", t("rewriteOwnGo"), runOwn);
    input.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); runOwn(); } });
    input.disabled = go.disabled = doc;
    own.append(input, go);
    box.append(own);
    const row = el("div", "lai-actions");
    row.append(button("lai-btn", t("rewriteCancel"), () => d.close()));
    box.append(row);
    d.append(box);
  }

  const changed = (r) => r.done && !!r.suggestion && r.suggestion !== r.item.original;
  const chosen = () => results.filter((r) => r.on && changed(r));

  function diffNode(r, pending) {
    const diff = el("p", "rw-diff");
    if (pending) diff.textContent = r.suggestion || "…";
    else {
      dwbCompareWords(r.item.original, r.suggestion).forEach((part) => {
        diff.append(el(part.kind === "same" ? "span" : part.kind === "added" ? "ins" : "del", "", part.text));
      });
    }
    return diff;
  }

  function renderResult() {
    const d = ensureDialog();
    d.replaceChildren();
    const box = titleBox();
    box.append(el("span", "ai-badge", t(state.source === "dict" ? "rewriteDictBadge" : "aiBadge")));
    const single = results.length === 1;
    if (single) {
      const r = results[0];
      box.append(diffNode(r, !r.done));
      if (r.warn) box.append(el("p", "lai-warn", r.warn));
    } else {
      const doneCount = results.filter((r) => r.done).length;
      const summary = state.busy
        ? (state.source === "dict" ? t("rewriteChecking") : t("rewriteProgress", { i: Math.min(results.length, doneCount + 1), n: results.length }))
        : t("rewriteSummary", { n: results.filter(changed).length, total: results.length });
      box.append(el("p", "rw-summary", summary));
      if (state.busy) {
        const bar = el("div", "lai-bar");
        const fill = el("i");
        fill.style.width = `${Math.round((doneCount / results.length) * 100)}%`;
        bar.append(fill);
        box.append(bar);
      }
      if (target.skipped) box.append(el("p", "lai-note", t("rewriteSkipped", { n: target.skipped })));
      const list = el("div", "rw-list");
      results.forEach((r, k) => {
        const current = state.busy && state.current === k;
        if (!changed(r) && !current && !r.warn) return;
        const item = el("div", "rw-item");
        item.dataset.k = String(k);
        if (changed(r)) {
          const pick = el("label", "rw-pick");
          const cb = el("input");
          cb.type = "checkbox";
          cb.checked = r.on;
          cb.addEventListener("change", () => { r.on = cb.checked; item.classList.toggle("is-off", !r.on); syncApply(); });
          pick.append(cb, el("span", "", t("rewritePick")));
          item.append(pick);
          item.classList.toggle("is-off", !r.on);
        }
        item.append(diffNode(r, !r.done));
        if (r.warn) item.append(el("p", "lai-warn", r.warn));
        list.append(item);
      });
      if (list.childElementCount) box.append(list);
      if (state.stopped) box.append(el("p", "lai-warn", t("rewriteStopped")));
      if (!state.busy && !results.some(changed) && state.source === "dict") box.append(el("p", "lai-warn", t(target.scope === "doc" ? "rewriteDocNothing" : "rewriteDictNothing")));
    }
    const row = el("div", "lai-actions");
    if (state.busy) {
      if (!single) row.append(button("lai-btn rw-stop", t("rewriteStop"), () => { state.stopped = true; abort?.abort(); }));
    } else {
      row.append(
        button("lai-btn is-primary rw-apply", "", () => applySuggestion(true)),
        button("lai-btn rw-apply-plain", "", () => applySuggestion(false)),
        button("lai-btn", t("rewriteAgain"), () => renderChooser()),
      );
    }
    row.append(button("lai-btn", t("rewriteCancel"), () => d.close()));
    box.append(row);
    d.append(box);
    syncApply();
  }

  function syncApply() {
    const n = chosen().length;
    const many = results.length > 1;
    const apply = dialog?.querySelector(".rw-apply");
    const plain = dialog?.querySelector(".rw-apply-plain");
    if (apply) { apply.textContent = many ? t("rewriteApplyTrackedN", { n }) : t("rewriteApplyTracked"); apply.disabled = !n; }
    if (plain) { plain.textContent = many ? t("rewriteApplyPlainN", { n }) : t("rewriteApplyPlain"); plain.disabled = !n; }
  }

  // Strumień tokenów: tylko tekst bieżącego akapitu, bez przebudowy całego okna.
  function showPartial(k) {
    if (!dialog?.open) return;
    const node = results.length === 1 ? dialog.querySelector(".rw-diff") : dialog.querySelector(`.rw-item[data-k="${k}"] .rw-diff`);
    if (node) node.textContent = results[k].suggestion || "…";
    else renderResult();
  }

  async function run(instruction, mode = "") {
    if (target.scope === "doc") {
      if (mode !== "fix") return;
      const got = await documentItems();
      target.items = got.items;
      target.skipped = got.skipped;
    }
    if (!target.items.length) { renderChooser(t("rewriteDocEmpty")); return; }
    results = target.items.map((item) => ({ item, suggestion: "", warn: "", on: true, done: false }));
    // „Popraw błędy” zawsze słownikiem: pewne i natychmiastowe. Mały model „poprawiał” też
    // dobre słowa (zdążę → zdąży — pomiar docs/llm-bench/prompts.html, 2026-10-17).
    if (mode === "fix") {
      state = { instruction, source: "dict", busy: true, current: -1 };
      if (!dialog.open) dialog.showModal();
      renderResult();
      const fixed = await offlineFixMany(target.items.map((it) => it.original), { lists: target.items.map((it) => !!it.list) });
      results.forEach((r, k) => { r.suggestion = fixed[k]; r.done = true; });
      state.busy = false;
      if (results.length === 1 && !changed(results[0])) results[0].warn = t("rewriteDictNothing");
      if (dialog.open) renderResult();
      return;
    }
    if (!(await LocalAI.ensure({ reason: t("rewriteAiReason") }))) {
      renderChooser();
      if (!dialog.open) dialog.showModal();
      return;
    }
    if (!dialog.open) dialog.showModal();
    state = { instruction, source: "ai", busy: true, current: 0, stopped: false };
    abort = new AbortController();
    const signal = abort.signal;
    renderResult();
    for (let k = 0; k < results.length && !signal.aborted; k++) {
      const r = results[k];
      state.current = k;
      if (k) renderResult();
      try {
        const answer = await LocalAI.generate({
          system: SYSTEM,
          // przykład w poleceniu — mały model trzyma format „tylko nowy tekst” (pomiar v1 → v2)
          shots: [[`Polecenie: ${PROMPTS.formal}\n\nFragment:\nDajcie znać, czy pasuje wam wtorek o 9.`, "Uprzejmie prosimy o informację, czy termin we wtorek o godzinie 9 jest dla Państwa odpowiedni."]],
          user: `Polecenie: ${instruction}\n\nFragment:\n${r.item.original}`,
          maxTokens: Math.min(900, Math.max(80, Math.round(r.item.original.length / 2.2) + 60)),
          signal,
          onPartial: (s) => { r.suggestion = s; showPartial(k); },
        });
        r.suggestion = clean(answer, r.item.original);
        if (!r.suggestion) r.warn = t("rewriteEmptyAnswer");
        else if (r.suggestion.length > r.item.original.length * 3 + 200) { r.warn = t("rewriteTooLong"); r.suggestion = ""; }
        else {
          const missing = missingNumbers(r.item.original, r.suggestion);
          // kilka akapitów: zmiana z brakującą liczbą zostaje na liście, ale odznaczona
          if (missing.length) {
            r.warn = t(results.length > 1 ? "rewriteMissingNumbersOff" : "rewriteMissingNumbers", { nums: missing.join(", ") });
            r.on = results.length === 1;
          }
        }
        r.done = true;
      } catch (err) {
        if (signal.aborted) { r.suggestion = ""; break; }
        r.warn = t("aiFailed", { msg: err.message || err });
        r.suggestion = "";
        r.done = true;
      }
    }
    state.busy = false;
    if (dialog.open) renderResult();
  }

  async function applySuggestion(tracked) {
    const picked = chosen();
    if (!target || !picked.length) return;
    dialog.close();
    // tekst akapitów nie mógł się zmienić w międzyczasie (okno modalne), ale plik — sprawdzamy
    await mergeInlineEditsIntoBytes();
    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    if (picked.some((r) => texts[r.item.paraIndex] !== r.item.paraText)) { toast(t("rewriteChanged"), "warning"); return; }
    const items = picked.map(({ item, suggestion }) => ({
      index: item.paraIndex,
      text: item.paraText.slice(0, item.start) + suggestion + item.paraText.slice(item.end),
    }));
    const n = await applyDocumentEdit(tracked
      ? { op: "rewriteTracked", items, author: AUTHOR() }
      : { op: "paragraphBatch", items });
    if (n <= 0) return;
    if (items.length > 1) toast(t(tracked ? "rewriteDoneTrackedMany" : "rewriteDoneMany", { n: items.length }), "success");
    else toast(t(tracked ? "rewriteDoneTracked" : "rewriteDone"), "success");
  }

  async function open() {
    const got = await captureTarget();
    if (got.error) { toast(t(got.error), "info"); return; }
    target = got;
    results = [];
    renderChooser(got.note);
    ensureDialog().showModal();
  }

  function init() {
    const btn = document.getElementById("rewriteBtn");
    btn?.addEventListener("mousedown", (e) => e.preventDefault()); // zaznaczenie w tekście zostaje
    btn?.addEventListener("click", open);
  }

  return {
    init, open, clean, missingNumbers, offlineFix, offlineFixMany,
    _target: () => (target?.items?.length === 1 ? target.items[0] : target),
    _results: () => results,
  };
})();

document.addEventListener("DOMContentLoaded", () => dwbRewrite.init());
