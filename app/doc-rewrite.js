// doc-rewrite.js — „✦ Przepisz” na pasku Edycji: asystent AI proponuje nową wersję zaznaczenia
// (albo akapitu z kursorem), a zmiana trafia do dokumentu jako POPRAWKA Worda (w:ins / w:del,
// autor „Asystent AI”) — do przejrzenia w Recenzji: Akceptuj / Odrzuć. Jeden krok Cofnij.
//
// Zasady: podgląd różnicy przed wstawieniem; liczby z oryginału muszą zostać (inaczej ostrzeżenie);
// akapity zablokowane (przypis, pole, link, obraz) — bez zmian, jak przy zwykłym pisaniu.

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

  let dialog = null;
  let target = null; // { paraIndex, start, end, original, paraText }
  let suggestion = "";
  let abort = null;

  // ── co przepisujemy ───────────────────────────────────────────────────────
  async function captureTarget() {
    if (readOnlyMode || !originalFileBytes) return { error: "rewriteNoDoc" };
    const p = restoreDocCaret?.();
    if (!p) return { error: "rewriteNoCaret" };
    if (p.dataset.lock) return { error: "rewriteLocked" };
    const sel = window.getSelection();
    const range = sel?.rangeCount ? sel.getRangeAt(0) : null;
    if (range && !range.collapsed) {
      const startP = docCaretParagraph(range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement);
      const endP = docCaretParagraph(range.endContainer.nodeType === 1 ? range.endContainer : range.endContainer.parentElement);
      if (startP && endP && startP !== endP) return { error: "rewriteOneParagraph" };
    }
    const selected = range && !range.collapsed ? range.toString() : "";
    const paraIndex = resolveParaIndex(p);
    let offset = 0;
    if (selected) {
      const pre = range.cloneRange();
      pre.selectNodeContents(p);
      pre.setEnd(range.startContainer, range.startOffset);
      offset = pre.toString().length;
    }
    await mergeInlineEditsIntoBytes();
    const paraText = (await extractParagraphTextsFromDocx(originalFileBytes))[paraIndex];
    if (paraText == null || !paraText.trim()) return { error: "rewriteEmpty" };
    if (!selected.trim()) return { paraIndex, start: 0, end: paraText.length, original: paraText, paraText };
    // zaznaczenie w tekście z pliku: najbliższe wystąpienie wokół miejsca z podglądu
    let start = paraText.indexOf(selected, Math.max(0, offset - 4));
    if (start < 0) start = paraText.indexOf(selected);
    if (start < 0) return { paraIndex, start: 0, end: paraText.length, original: paraText, paraText, wholeFallback: true };
    return { paraIndex, start, end: start + selected.length, original: selected, paraText };
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
    return (String(text).match(/\d[\d\s .,]*\d|\d/g) || []).map((n) => n.replace(/[\s ]/g, ""));
  }
  function missingNumbers(original, rewritten) {
    const have = new Set(numbersOf(rewritten));
    return [...new Set(numbersOf(original))].filter((n) => !have.has(n));
  }

  // ── okno ──────────────────────────────────────────────────────────────────
  function el(tag, cls, text) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
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

  function renderChooser(note = "") {
    const d = ensureDialog();
    d.replaceChildren();
    const box = el("div", "lai-box");
    const h = el("h2", "lai-title", t("rewriteTitle"));
    h.id = "rwTitle";
    box.append(h);
    box.append(el("p", "rw-original", target.original.length > 400 ? `${target.original.slice(0, 400)}…` : target.original));
    if (note) box.append(el("p", "lai-note", note));
    const modes = el("div", "rw-modes");
    MODES.forEach((m) => {
      const b = el("button", "lai-btn", t(`rewriteMode_${m}`));
      b.type = "button";
      b.dataset.mode = m;
      b.addEventListener("click", () => run(PROMPTS[m]));
      modes.append(b);
    });
    box.append(modes);
    const own = el("div", "rw-own");
    const input = el("input");
    input.type = "text";
    input.placeholder = t("rewriteOwnPlaceholder");
    input.setAttribute("aria-label", t("rewriteOwnPlaceholder"));
    const go = el("button", "lai-btn is-primary", t("rewriteOwnGo"));
    go.type = "button";
    const runOwn = () => { if (input.value.trim()) run(input.value.trim()); };
    go.addEventListener("click", runOwn);
    input.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); runOwn(); } });
    own.append(input, go);
    box.append(own);
    const row = el("div", "lai-actions");
    const close = el("button", "lai-btn", t("rewriteCancel"));
    close.type = "button";
    close.addEventListener("click", () => d.close());
    row.append(close);
    box.append(row);
    d.append(box);
  }

  function renderResult(instruction, { pending = false, warn = "" } = {}) {
    const d = ensureDialog();
    d.replaceChildren();
    const box = el("div", "lai-box");
    const h = el("h2", "lai-title", t("rewriteTitle"));
    h.id = "rwTitle";
    box.append(h, el("span", "ai-badge", t("aiBadge")));
    const diff = el("p", "rw-diff");
    if (pending) diff.textContent = suggestion || "…";
    else {
      dwbCompareWords(target.original, suggestion).forEach((part) => {
        const node = el(part.kind === "same" ? "span" : part.kind === "added" ? "ins" : "del", "", part.text);
        diff.append(node);
      });
    }
    box.append(diff);
    if (warn) box.append(el("p", "lai-warn", warn));
    const row = el("div", "lai-actions");
    if (!pending) {
      const apply = el("button", "lai-btn is-primary rw-apply", t("rewriteApplyTracked"));
      apply.type = "button";
      apply.addEventListener("click", () => applySuggestion(true));
      const plain = el("button", "lai-btn rw-apply-plain", t("rewriteApplyPlain"));
      plain.type = "button";
      plain.addEventListener("click", () => applySuggestion(false));
      const again = el("button", "lai-btn", t("rewriteAgain"));
      again.type = "button";
      again.addEventListener("click", () => renderChooser());
      apply.disabled = plain.disabled = !suggestion || suggestion === target.original;
      row.append(apply, plain, again);
    }
    const close = el("button", "lai-btn", t("rewriteCancel"));
    close.type = "button";
    close.addEventListener("click", () => d.close());
    row.append(close);
    box.append(row);
    d.append(box);
  }

  async function run(instruction) {
    if (!(await LocalAI.ensure({ reason: t("rewriteAiReason") }))) {
      renderChooser();
      if (!dialog.open) dialog.showModal();
      return;
    }
    if (!dialog.open) dialog.showModal();
    suggestion = "";
    renderResult(instruction, { pending: true });
    abort = new AbortController();
    try {
      const answer = await LocalAI.generate({
        system: SYSTEM,
        user: `Polecenie: ${instruction}\n\nFragment:\n${target.original}`,
        maxTokens: Math.min(900, Math.max(80, Math.round(target.original.length / 2.2) + 60)),
        signal: abort.signal,
        onPartial: (s) => { suggestion = s; if (dialog.open) renderResult(instruction, { pending: true }); },
      });
      suggestion = clean(answer, target.original);
      let warn = "";
      if (!suggestion) warn = t("rewriteEmptyAnswer");
      else if (suggestion.length > target.original.length * 3 + 200) { warn = t("rewriteTooLong"); suggestion = ""; }
      else {
        const missing = missingNumbers(target.original, suggestion);
        if (missing.length) warn = t("rewriteMissingNumbers", { nums: missing.join(", ") });
      }
      if (dialog.open) renderResult(instruction, { warn });
    } catch (err) {
      if (dialog.open) renderResult(instruction, { warn: t("aiFailed", { msg: err.message || err }) });
    }
  }

  async function applySuggestion(tracked) {
    if (!target || !suggestion) return;
    dialog.close();
    // tekst akapitu nie mógł się zmienić w międzyczasie (okno modalne), ale plik — sprawdzamy
    await mergeInlineEditsIntoBytes();
    const now = (await extractParagraphTextsFromDocx(originalFileBytes))[target.paraIndex];
    if (now !== target.paraText) { toast(t("rewriteChanged"), "warning"); return; }
    const text = target.paraText.slice(0, target.start) + suggestion + target.paraText.slice(target.end);
    const n = await applyDocumentEdit(tracked
      ? { op: "rewriteTracked", index: target.paraIndex, text, author: AUTHOR() }
      : { op: "paragraphBatch", items: [{ index: target.paraIndex, text }] });
    if (n > 0) toast(t(tracked ? "rewriteDoneTracked" : "rewriteDone"), "success");
  }

  async function open() {
    const got = await captureTarget();
    if (got.error) { toast(t(got.error), "info"); return; }
    target = got;
    suggestion = "";
    renderChooser(got.wholeFallback ? t("rewriteWholeFallback") : "");
    ensureDialog().showModal();
  }

  function init() {
    const btn = document.getElementById("rewriteBtn");
    btn?.addEventListener("mousedown", (e) => e.preventDefault()); // zaznaczenie w tekście zostaje
    btn?.addEventListener("click", open);
  }

  return { init, open, clean, missingNumbers, _target: () => target };
})();

document.addEventListener("DOMContentLoaded", () => dwbRewrite.init());
