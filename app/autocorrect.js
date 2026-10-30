// autocorrect.js — Autokorekta i Autoformatowanie podczas pisania, jak w Wordzie (2026-10-17).
//
// Reguły zbadane w Microsoft Word 16.114 (REA: lista opcji z WDPreferences.plist, format listy
// Default_ACL_[Polish] z mso99.framework; potem pisanie w Wordzie z językiem polskim):
//   - poprawka słowa po spacji / interpunkcji / Enter / Tab: „ktory” → „który”, „juz” → „już”,
//     wielkość liter zachowana („Juz” → „Już”),
//   - wielka litera na początku zdania i akapitu (też w punkcie listy i komórce tabeli); NIE po
//     skrótach („np. tak”, „tzn. to”, „r. będzie”) ani po pojedynczej literze z kropką,
//   - „PIerwszy” → „Pierwszy” (DWie WIelkie na początku),
//   - "…" → „…” i '…' → ‘…’ (w tekście angielskim “…”), apostrof w słowie → ’,
//   - „a--b” → „a—b”, „raz - dwa” / „raz -- dwa” → „raz – dwa”, „1/2” → „½”, „...” → „…”,
//     „(c)” → „©”, „(r)” → „®”, „(tm)” → „™”, „(e)” → „€”,
//   - „1. ” / „1) ” / „* ” / „- ” na początku akapitu → lista numerowana / punktowana,
//   - Ctrl/⌘+Z zaraz po poprawce cofa SAMĄ poprawkę (w Wordzie: „Cofnij Autokorekta”),
//     kolejne — pisanie. Tu: poprawka = osobny krok cofania (dwbUndo.record / group).
// Własna lista poprawek: brak polskich znaków w częstych słowach i pisownia łączna/rozdzielna —
// bez słów, które bez ogonków też istnieją („ze”, „tez”, „wiec”, „chce”, „odpowiedz”…).
// Działa tylko przy pisaniu na końcu słowa (za kursorem nie ma litery) — poprawianie środka
// istniejącego tekstu niczego nie zmienia samo. Wyłączniki: Narzędzia edycji → Autokorekta.

const dwbAutoCorrect = (() => {
  const KEY = "dwb-autocorrect-v1";
  // Pod automatem testowym (navigator.webdriver) domyślnie wyłączona: pozostałe testy wpisują
  // dosłowny tekst i porównują go z plikiem. test:autocorrect włącza ją jawnie (localStorage).
  const DEFAULTS = { on: !navigator.webdriver, replace: true, caps: true, quotes: true, lists: true };

  function load() {
    try { return { ...DEFAULTS, ...(JSON.parse(localStorage.getItem(KEY) || "{}") || {}) }; } catch (_) { return { ...DEFAULTS }; }
  }
  let opts = load();
  function save() { try { localStorage.setItem(KEY, JSON.stringify(opts)); } catch (_) { /* bez pamięci */ } }
  const isOn = (k) => opts.on && opts[k] !== false;

  // ── lista poprawek (polski) ───────────────────────────────────────────────
  const PL_FIX = {
    juz: "już", sie: "się", dziala: "działa", dzialaja: "działają", jesli: "jeśli", jezeli: "jeżeli", mozna: "można", moze: "może", mozemy: "możemy",
    mozesz: "możesz", mozecie: "możecie", mozliwe: "możliwe", mozliwy: "możliwy", mozliwa: "możliwa",
    mozliwosc: "możliwość", mozliwosci: "możliwości", musze: "muszę", prosze: "proszę", dziekuje: "dziękuję",
    dziekujemy: "dziękujemy", dzieki: "dzięki", bede: "będę", beda: "będą", bedzie: "będzie",
    bedziemy: "będziemy", bedziesz: "będziesz", bedziecie: "będziecie", byc: "być", byl: "był",
    byla: "była", bylo: "było", byly: "były", bylem: "byłem", bylam: "byłam", bylismy: "byliśmy",
    miec: "mieć", mial: "miał", miala: "miała", mialem: "miałem", wziac: "wziąć", wzial: "wziął",
    ktory: "który", ktora: "która", ktore: "które", ktorego: "którego", ktorej: "której",
    ktoremu: "któremu", ktorych: "których", ktorym: "którym", ktorymi: "którymi", ktorzy: "którzy",
    kazdy: "każdy", kazda: "każda", kazde: "każde", kazdego: "każdego", kazdej: "każdej",
    zaden: "żaden", zadna: "żadna", zadne: "żadne", zadnego: "żadnego", zeby: "żeby", niz: "niż",
    iz: "iż", az: "aż", wciaz: "wciąż", dzis: "dziś", zle: "źle", duzo: "dużo", malo: "mało",
    wiecej: "więcej", czesto: "często", rowniez: "również", takze: "także", poniewaz: "ponieważ",
    oczywiscie: "oczywiście", szczegolnie: "szczególnie", wlasnie: "właśnie", wlasciwie: "właściwie",
    nastepnie: "następnie", nastepny: "następny", nastepna: "następna", nastepne: "następne",
    pozniej: "później", wczesniej: "wcześniej", wszedzie: "wszędzie", wiekszosc: "większość",
    wazne: "ważne", wazny: "ważny", wazna: "ważna", goraco: "gorąco", prosba: "prośba",
    swoj: "swój", twoj: "twój", moj: "mój", dopoki: "dopóki", wokol: "wokół", okolo: "około",
    sposob: "sposób", wsrod: "wśród", miedzy: "między", pomiedzy: "pomiędzy", wzgledem: "względem",
    wzgledu: "względu", dzien: "dzień", tydzien: "tydzień", miesiac: "miesiąc", piatek: "piątek",
    sroda: "środa", styczen: "styczeń", kwiecien: "kwiecień", sierpien: "sierpień",
    wrzesien: "wrzesień", pazdziernik: "październik", grudzien: "grudzień",
    // pisownia łączna / rozdzielna
    narazie: "na razie", napewno: "na pewno", wogole: "w ogóle", "wogóle": "w ogóle",
    poprostu: "po prostu", odrazu: "od razu", conajmniej: "co najmniej", nietylko: "nie tylko",
    pokolei: "po kolei", naprzyklad: "na przykład", "naprzykład": "na przykład",
    "zwyjątkiem": "z wyjątkiem", zwyjatkiem: "z wyjątkiem", niewiem: "nie wiem",
    wkoncu: "w końcu", "wkońcu": "w końcu", "niebędzie": "nie będzie", "niemoże": "nie może",
  };
  // skróty bez kropki → z kropką (tylko przed spacją: „itd.,” piszemy sami)
  const PL_ABBR_DOT = { np: "np.", itd: "itd.", itp: "itp.", tzw: "tzw.", "m.in": "m.in.", prof: "prof." };
  // po tych skrótach z kropką zdanie trwa — bez wielkiej litery („itd.”, „zł.” często kończą
  // zdanie, więc ich tu nie ma — Word też po nich zaczyna wielką)
  const PL_NO_CAP_AFTER = new Set(("np tj tzn tzw m.in ok por zob ul al nr godz min tel ds ew lp nb nt pkt poz par pn pt " +
    "rozdz rys str tab wyd ww zał im św dr prof mgr inż hab płk gen ks ur zm wg jw ang łac niem przyp red tłum oprac " +
    "ust art lit ryc").split(/\s+/));

  const SYMBOLS = { "(c)": "©", "(r)": "®", "(tm)": "™", "(e)": "€" };
  const FRACTIONS = { "1/2": "½", "1/4": "¼", "3/4": "¾" };
  const DELIM_RE = /^[\s .,;:!?)\]}"'»”’…]$/u;
  const LETTER_RE = /[\p{L}\p{N}]/u;

  function docLang(p) {
    const tagged = p?.closest?.("[lang]")?.getAttribute("lang");
    if (tagged) return tagged.toLowerCase().startsWith("pl") ? "pl" : "other";
    const g = document.getElementById("grammarLang")?.value;
    return g && g !== "pl" ? "other" : "pl";
  }

  // Kursor w zwykłym tekście akapitu (nie w linku, polu formularza, odsyłaczu).
  function caretSpot() {
    const sel = window.getSelection();
    if (!sel?.isCollapsed || !sel.rangeCount) return null;
    let node = sel.focusNode;
    let offset = sel.focusOffset;
    // kursor za wstawionym fragmentem (<span> ze stylem) — ostatni tekst przed nim
    if (node?.nodeType === 1 && offset > 0) {
      let n = node.childNodes[offset - 1];
      while (n?.nodeType === 1 && n.getAttribute("contenteditable") !== "false") n = n.lastChild;
      if (n?.nodeType === 3) { node = n; offset = n.length; }
    }
    if (node?.nodeType !== 3) return null;
    const el = node.parentElement;
    const p = el?.closest(".docx-editable-p");
    if (!p || el.closest("a, [contenteditable=false], .doc-xref")) return null;
    return { sel, node, offset, p };
  }
  function textBefore(p, node, offset) {
    const r = document.createRange();
    r.selectNodeContents(p);
    r.setEnd(node, offset);
    return r.toString();
  }
  function charAfter(p, node, offset) {
    if (offset < node.length) return node.data[offset];
    const r = document.createRange();
    r.selectNodeContents(p);
    r.setStart(node, offset);
    return r.toString().charAt(0);
  }
  function setCaret(sel, node, offset) {
    const r = document.createRange();
    r.setStart(node, Math.max(0, Math.min(node.length, offset)));
    r.collapse(true);
    sel.removeAllRanges();
    sel.addRange(r);
  }
  function matchCase(src, repl) {
    if (src.length > 1 && src === src.toLocaleUpperCase("pl") && src !== src.toLocaleLowerCase("pl")) return repl.toLocaleUpperCase("pl");
    if (src[0] !== src[0].toLocaleLowerCase("pl")) return repl[0].toLocaleUpperCase("pl") + repl.slice(1);
    return repl;
  }

  // Zmiany słowa kończącego się w node na pozycji `end` (przed wpisanym ogranicznikiem).
  // Zwraca listę { start, len, text } w kolejności od końca (żeby przesunięcia się nie psuły).
  function planWord(spot, end, delim, lang) {
    const { node, p } = spot;
    const head = node.data.slice(0, end);
    const edits = [];
    const token = head.match(/[^\s ]+$/u)?.[0] || "";
    if (!token) return edits;
    const tokenStart = end - token.length;
    const before = textBefore(p, node, tokenStart);
    // słowo zaczyna się w poprzednim węźle (np. połowa pogrubiona) — nie ruszamy
    if (tokenStart === 0 && before && !/[\s ]$/u.test(before)) return edits;

    if (isOn("replace")) {
      // myślniki: „a--b” → pauza, „ - b” / „ -- b” → półpauza (słowo po myślniku już wpisane)
      const em = token.match(/^([^-]*[^\s-])--([^\s-][^-]*)$/u);
      if (em) { edits.push({ start: tokenStart + em[1].length, len: 2, text: "—" }); return edits; }
      const en = head.match(/(\S)([\s ])(--?)([\s ])[^\s ]+$/u);
      if (en && !/[-–—]/u.test(en[1])) {
        const dashAt = end - token.length - en[4].length - en[3].length;
        edits.push({ start: dashAt, len: en[3].length, text: "–" });
      }
      if (FRACTIONS[token] && (tokenStart > 0 || !before || /[\s ]$/u.test(before))) {
        edits.push({ start: tokenStart, len: token.length, text: FRACTIONS[token] });
        return edits;
      }
    }

    // samo słowo z liter (może z kropką w środku: „m.in”)
    const wordM = token.match(/[\p{L}][\p{L}.]*$/u);
    if (!wordM || !/^[\p{L}][\p{L}.]*$/u.test(token.replace(/^[(\[„"'“‘«]+/u, ""))) return edits;
    let word = wordM[0];
    const wordStart = end - word.length;
    let replacement = word;

    if (isOn("replace") && lang === "pl") {
      const low = word.toLocaleLowerCase("pl");
      if (PL_FIX[low]) replacement = matchCase(word, PL_FIX[low]);
      else if (PL_ABBR_DOT[low] && /[\s ]/u.test(delim)) replacement = matchCase(word, PL_ABBR_DOT[low]);
    }
    if (isOn("caps")) {
      // DWie WIelkie: dwie wielkie + co najmniej dwie małe
      if (/^\p{Lu}\p{Lu}\p{Ll}{2,}$/u.test(replacement)) replacement = replacement[0] + replacement.slice(1).toLocaleLowerCase("pl");
      // początek zdania / akapitu
      const ctx = before + node.data.slice(tokenStart, wordStart);
      if (/^\p{Ll}/u.test(replacement) && !/\./u.test(word) && sentenceStart(ctx)) {
        replacement = replacement[0].toLocaleUpperCase("pl") + replacement.slice(1);
      }
    }
    if (replacement !== word) edits.push({ start: wordStart, len: word.length, text: replacement });
    return edits;
  }

  function sentenceStart(ctx) {
    // otwierający cudzysłów / nawias tuż przed słowem nie przeszkadza
    const c = ctx.replace(/[(\[„"'“‘«]+$/u, "");
    if (!c.trim()) return true; // początek akapitu (też punkt listy, komórka tabeli)
    if (/[!?]["”’)»]*[\s ]+$/u.test(c)) return true;
    const m = c.match(/(\S*)\.["”’)»]*[\s ]+$/u);
    if (!m) return false;
    const prev = m[1].replace(/^[(\[„"'“‘«]+/u, "");
    if (!prev) return true; // sama kropka
    if (/\d$/u.test(prev)) return false; // „2026. tak”, „§ 3. ust” — liczby, daty
    if (/\.$/u.test(prev)) return false; // wielokropek z kropek
    if (/^\p{L}$/u.test(prev)) return false; // „A. Kowalski”, „s. 12”, „r. będzie”
    return !PL_NO_CAP_AFTER.has(prev.toLocaleLowerCase("pl"));
  }

  let toldOnce = false;
  function tell() {
    if (toldOnce) return;
    toldOnce = true;
    try {
      if (sessionStorage.getItem("dwbAutoCorrectTold")) return;
      sessionStorage.setItem("dwbAutoCorrectTold", "1");
    } catch (_) { /* bez pamięci sesji */ }
    if (typeof toast === "function") toast(t("autoCorrectUndoHint"), "info");
  }

  // Zastosuj edycje w węźle (od końca) i przesuń kursor; jeden krok Cofnij.
  function applyEdits(spot, edits) {
    if (!edits.length) return false;
    const { node, sel } = spot;
    let caret = spot.offset;
    const ordered = edits.slice().sort((a, b) => b.start - a.start);
    const run = () => {
      ordered.forEach((e) => {
        node.replaceData(e.start, e.len, e.text);
        if (e.start < caret) caret += e.text.length - e.len;
      });
      setCaret(sel, node, caret);
    };
    if (typeof asUndoStep === "function") asUndoStep("undoOpAutoCorrect", run); else run();
    if (typeof onInlineParagraphInput === "function") onInlineParagraphInput();
    tell();
    return true;
  }

  // ── wpisany znak ──────────────────────────────────────────────────────────
  function onInput(e) {
    if (!opts.on || readOnlyMode || e.isComposing || e.inputType !== "insertText") return;
    const ch = e.data || "";
    if (ch.length !== 1) return;
    const spot = caretSpot();
    if (!spot || spot.offset < 1 || spot.node.data[spot.offset - 1] !== ch && !(ch === " " && spot.node.data[spot.offset - 1] === " ")) return;
    const lang = docLang(spot.p);
    const after = charAfter(spot.p, spot.node, spot.offset);
    const edits = [];

    // cudzysłowy drukarskie — sam wpisany znak
    if (isOn("quotes") && (ch === '"' || ch === "'")) {
      const prev = textBefore(spot.p, spot.node, spot.offset - 1).slice(-1);
      const opening = !prev || /[\s (\[{„“‘«—–-]/u.test(prev);
      let q;
      if (ch === '"') q = lang === "pl" ? (opening ? "„" : "”") : (opening ? "“" : "”");
      else q = opening ? "‘" : "’";
      edits.push({ start: spot.offset - 1, len: 1, text: q });
    }
    // „...” → „…” (przy trzeciej kropce, nie czwartej)
    if (isOn("replace") && ch === "." && /(^|[^.])\.\.\.$/u.test(spot.node.data.slice(0, spot.offset))) {
      edits.push({ start: spot.offset - 3, len: 3, text: "…" });
      return void applyEdits(spot, edits);
    }
    // „(c)” → „©” przy nawiasie zamykającym
    if (isOn("replace") && ch === ")") {
      const m = spot.node.data.slice(0, spot.offset).match(/\((c|r|tm|e)\)$/iu);
      if (m && SYMBOLS[m[0].toLowerCase()]) {
        edits.push({ start: spot.offset - m[0].length, len: m[0].length, text: SYMBOLS[m[0].toLowerCase()] });
        return void applyEdits(spot, edits);
      }
    }
    if (!DELIM_RE.test(ch) || LETTER_RE.test(after)) return void applyEdits(spot, edits);

    // listy: „1. ” „1) ” „* ” „- ” na samym początku akapitu
    if (isOn("lists") && /^[\s ]$/u.test(ch) && !spot.p.dataset.noteKey && !isListParagraph(spot.p)) {
      const before = textBefore(spot.p, spot.node, spot.offset);
      const m = before.match(/^(1[.)]|\*|-)[\s ]$/u);
      if (m) { makeList(spot, m[1] === "*" || m[1] === "-" ? "bullet" : "number", before.length); return; }
    }
    edits.push(...planWord(spot, spot.offset - 1, ch, lang));
    applyEdits(spot, edits);
  }

  // Zamiana na listę idzie tą samą drogą co przycisk „Lista” (zmiana w pliku + przerysowanie —
  // numeracja to liczniki CSS z renderu). „1. ” pisze się w biegu, więc klawisze wciśnięte
  // w trakcie przerysowania trzyma holdTypingWhile (docx-inline-edit.js) i dopisuje potem.
  async function makeList(spot, kind, markerLen) {
    const p = spot.p;
    const index = resolveParaIndex(p);
    if (index < 0 || typeof applyDocumentEdit !== "function" || typingHeld()) return;
    const run = async () => {
      const r = document.createRange();
      r.selectNodeContents(p);
      r.setEnd(spot.node, spot.offset);
      if (r.toString().length !== markerLen) return;
      r.deleteContents();
      if (typeof onInlineParagraphInput === "function") onInlineParagraphInput();
      pendingInlineCursor = { paraIndex: index, offset: 0 };
      await applyDocumentEdit({ op: "list", indices: [index], kind, toggle: false });
    };
    try {
      await holdTypingWhile(() => (typeof dwbUndo !== "undefined" && dwbUndo.group ? dwbUndo.group("undoOpAutoList", run) : run()), index);
      tell();
    } catch (err) {
      if (typeof log === "function") log(`Autoformatowanie listy: ${err.message || err}`, "error");
    }
  }

  // Enter / Tab kończą słowo jak spacja — przed ich obsługą (podział akapitu czyta już poprawkę)
  function onKeyDown(e) {
    if (!opts.on || readOnlyMode || e.isComposing || (e.key !== "Enter" && e.key !== "Tab") || e.ctrlKey || e.metaKey || e.altKey) return;
    if (!docCaretParagraph(e.target)) return;
    const spot = caretSpot();
    if (!spot || LETTER_RE.test(charAfter(spot.p, spot.node, spot.offset))) return;
    // „1.” / „1)” / „*” / „-” + Tab na samym początku akapitu → lista, jak ze spacją (Word 16.114,
    // sprawdzone pisaniem: „1.⇥” = lista numerowana, tabulator nie zostaje w tekście). Dawniej
    // zostawał tekst „1.⇥”, a kolejne „1.⇥” w akapicie pod nim dawały dziwnie wciętą listę.
    if (e.key === "Tab" && !e.shiftKey && isOn("lists") && !spot.p.dataset.noteKey && !isListParagraph(spot.p)) {
      const before = textBefore(spot.p, spot.node, spot.offset);
      const m = before.match(/^(1[.)]|\*|-)$/u);
      if (m && !charAfter(spot.p, spot.node, spot.offset)) {
        e.preventDefault();
        e.stopPropagation();
        makeList(spot, m[1] === "*" || m[1] === "-" ? "bullet" : "number", before.length);
        return;
      }
    }
    applyEdits(spot, planWord(spot, spot.offset, e.key === "Enter" ? "\n" : "\t", docLang(spot.p)));
  }

  docCanvasEl?.addEventListener("input", onInput);
  document.addEventListener("keydown", onKeyDown, true);

  // ── wyłączniki (Narzędzia edycji → Autokorekta) ───────────────────────────
  function bindToggles() {
    document.querySelectorAll("[data-autocorrect]").forEach((box) => {
      const k = box.dataset.autocorrect;
      box.checked = opts[k] !== false;
      box.addEventListener("change", () => {
        opts[k] = box.checked;
        save();
        syncToggles();
      });
    });
    syncToggles();
  }
  function syncToggles() {
    document.querySelectorAll("[data-autocorrect]:not([data-autocorrect=on])").forEach((box) => {
      box.disabled = !opts.on;
      box.closest("label")?.classList.toggle("is-disabled", !opts.on);
    });
  }
  bindToggles();

  return {
    get options() { return { ...opts }; },
    set(k, v) { opts[k] = v; save(); document.querySelectorAll(`[data-autocorrect="${k}"]`).forEach((b) => { b.checked = v !== false; }); syncToggles(); },
  };
})();
