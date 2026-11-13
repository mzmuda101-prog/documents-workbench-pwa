// review-margin.js — znaczniki recenzji jak w Wordzie na Macu (Recenzja → Wszystkie znaczniki).
//
// Obok kartki, na prawym marginesie, dymki: komentarze (awatar w kolorze autora, kiedy, treść,
// odpowiedzi; Odpowiedz / Rozwiąż / Edytuj / Usuń) i śledzone zmiany („Usunięto: …”, „Wstawiono:
// …”, ✓ / ✗). Dymek stoi na wysokości swojego miejsca w tekście; aktywny (kursor w komentarzu,
// klik w dymek) stoi dokładnie przy nim, z linią łączącą, a pozostałe rozsuwają się w górę i w dół.
// W tekście: wstawki podkreślone w kolorze autora (CSS Highlight — bez zmiany DOM dokumentu, więc
// edycja i zapis nie widzą znaczników), pasek przy akapicie ze zmianą.
// Podczas „Śledź zmiany” (track-changes.js) to samo na żywo z różnicy słów względem pliku „przed”.
//
// Warstwa leży NAD obszarem dokumentu (position: fixed, przycięta do niego), nie w nim: kartki,
// powiększenie, Dopasuj i liczenie stron zostają bez zmian. Przewijanie przesuwa tylko warstwę
// (transform), więc nic się nie przelicza przy każdym ruchu. Za mało miejsca obok kartki (telefon,
// duże powiększenie, otwarty panel) — dymków nie ma, komentarz pokazuje dalej karta nad wierszem.

const dwbMargin = (() => {
  const MODE_KEY = "dwb.markupMode";
  const COL_MIN = 200, COL_MAX = 300, GAP = 8;
  const COLORS = 6;
  let mode = "all";
  try { mode = localStorage.getItem(MODE_KEY) || "all"; } catch (_) {}
  if (!/^(all|simple|none)$/.test(mode)) mode = "all";

  let cache = { bytes: null, job: null };
  let layer = null, inner = null, wire = null;
  let items = []; // { key, kind, el, anchor: {x, y, h}, top }
  let active = null, hover = null;
  let showing = false;
  let timer = 0, raf = 0;
  const authors = new Map(); // autor → numer koloru (stały w obrębie pliku)

  const host = () => docCanvasEl?.querySelector(".docx-preview-host");
  const vp = () => docViewportEl;
  const hlOk = () => typeof CSS !== "undefined" && CSS.highlights && typeof Highlight !== "undefined";
  const colorOf = (name) => {
    if (!authors.has(name)) authors.set(name, authors.size % COLORS);
    return authors.get(name);
  };
  const initials = (c) => (c.initials || String(c.author || "?").trim().split(/\s+/).map((w) => w[0] || "").join("").slice(0, 2)).toUpperCase();
  const ICONS = {
    ok: '<polyline points="20 6 9 17 4 12"/>',
    no: '<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>',
    reply: '<polyline points="9 17 4 12 9 7"/><path d="M20 18v-2a4 4 0 0 0-4-4H4"/>',
    resolve: '<circle cx="12" cy="12" r="9"/><polyline points="8 12.5 11 15.5 16.5 9.5"/>',
    reopen: '<path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><polyline points="3 3 3 8 8 8"/>',
    edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
    del: '<polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>',
  };
  const svg = (k) => `<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[k]}</svg>`;

  function when(iso) {
    const d = iso ? new Date(iso) : null;
    if (!d || Number.isNaN(d.getTime())) return "";
    const locale = currentLang === "en" ? "en-GB" : "pl-PL";
    const s = (Date.now() - d.getTime()) / 1000;
    if (s >= 0 && s < 7 * 86400 && typeof Intl.RelativeTimeFormat === "function") {
      const rtf = new Intl.RelativeTimeFormat(locale, { numeric: "auto" });
      if (s < 60) return rtf.format(0, "second");
      if (s < 3600) return rtf.format(-Math.round(s / 60), "minute");
      if (s < 86400) return rtf.format(-Math.round(s / 3600), "hour");
      return rtf.format(-Math.round(s / 86400), "day");
    }
    return d.toLocaleDateString(locale, { day: "numeric", month: "short", year: "numeric" });
  }

  // ── zakres w tekście akapitu podglądu ─────────────────────────────────────
  // Położenia liczone jak tekst akapitu „po przyjęciu” (łamanie wiersza = 1 znak). Gdy w tym
  // miejscu nie stoi oczekiwany tekst (numer przypisu, tabulator rysowany bez znaku) — najbliższe
  // wystąpienie w pobliżu, a gdy go nie ma, nic (lepiej bez podkreślenia niż w złym miejscu).
  function textNodes(p) {
    const out = [];
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (n.nodeType === 3) out.push(n);
      else if (n.tagName === "BR" && !n.dataset.dwbPh) out.push(n);
    }
    return out;
  }
  function rangeIn(p, start, expect, len = expect ? expect.length : 0) {
    const nodes = textNodes(p);
    const text = nodes.map((n) => (n.nodeType === 3 ? n.data : "\n")).join("");
    let at = Math.max(0, Math.min(start, text.length));
    if (expect && text.slice(at, at + expect.length) !== expect) {
      let best = -1;
      for (let i = text.indexOf(expect); i >= 0; i = text.indexOf(expect, i + 1)) if (best < 0 || Math.abs(i - start) < Math.abs(best - start)) best = i;
      if (best < 0 || Math.abs(best - start) > 24) return null;
      at = best;
    }
    const end = at + len;
    const r = document.createRange();
    let pos = 0, set = false;
    for (const n of nodes) {
      const l = n.nodeType === 3 ? n.data.length : 1;
      if (!set && at <= pos + l) {
        if (n.nodeType === 3) r.setStart(n, at - pos); else r.setStartBefore(n);
        set = true;
      }
      if (set && end <= pos + l) {
        if (n.nodeType === 3) r.setEnd(n, end - pos); else r.setEndAfter(n);
        return r;
      }
      pos += l;
    }
    if (!set) { r.selectNodeContents(p); r.collapse(false); }
    return r;
  }
  const firstRect = (range) => {
    const rs = range?.getClientRects?.();
    const r = (rs && rs.length ? rs[0] : null) || range?.getBoundingClientRect?.();
    return r && (r.width || r.height || r.top || r.left) ? r : null;
  };

  // ── dane ──────────────────────────────────────────────────────────────────
  function scan() {
    const bytes = originalFileBytes;
    if (!bytes || typeof scanDocxRevisions !== "function") return Promise.resolve(null);
    if (cache.bytes !== bytes) {
      const c = { bytes, job: scanDocxRevisions(bytes).catch(() => null), data: null };
      c.job.then((d) => { c.data = d; });
      cache = c;
    }
    return cache.job;
  }

  // ── warstwa ───────────────────────────────────────────────────────────────
  function ensureLayer() {
    if (layer) return;
    layer = document.createElement("div");
    layer.className = "dwb-markup";
    layer.setAttribute("role", "region");
    layer.setAttribute("aria-label", t("marginRegion"));
    inner = document.createElement("div");
    inner.className = "dwb-markup-inner";
    wire = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    wire.classList.add("mk-wire");
    wire.setAttribute("aria-hidden", "true");
    inner.appendChild(wire);
    layer.appendChild(inner);
    document.body.appendChild(layer);
  }
  function frame() {
    const v = vp();
    if (!layer || !v) return;
    const r = v.getBoundingClientRect();
    layer.style.left = `${r.left + v.clientLeft}px`;
    layer.style.top = `${r.top + v.clientTop}px`;
    layer.style.width = `${v.clientWidth}px`;
    layer.style.height = `${v.clientHeight}px`;
    inner.style.transform = `translate(${-v.scrollLeft}px, ${-v.scrollTop}px)`;
  }
  // prostokąt ekranu → współrzędne treści obszaru dokumentu (te, które przewija warstwa)
  function toContent(r) {
    const v = vp();
    const b = v.getBoundingClientRect();
    return { x: r.left - b.left - v.clientLeft + v.scrollLeft, y: r.top - b.top - v.clientTop + v.scrollTop, r: r.right - b.left - v.clientLeft + v.scrollLeft, h: r.height };
  }

  function clearHighlights() {
    if (!hlOk()) return;
    for (let k = 0; k < COLORS; k++) CSS.highlights.delete(`dwb-rev-${k}`);
    CSS.highlights.delete("dwb-mk-focus");
  }
  function clear() {
    items.forEach((it) => it.el.remove());
    items = [];
    inner?.querySelectorAll(".mk-bar").forEach((b) => b.remove());
    clearHighlights();
    showing = false;
    if (wire) wire.replaceChildren();
    layer?.classList.remove("has-cards");
  }

  function schedule() {
    clearTimeout(timer);
    timer = setTimeout(build, 140);
  }

  // ── budowa ────────────────────────────────────────────────────────────────
  let building = false, again = false;
  async function build() {
    if (building) { again = true; return; }
    building = true;
    try { await buildNow(); } finally {
      building = false;
      if (again) { again = false; schedule(); }
    }
  }
  async function buildNow() {
    const h = host();
    const v = vp();
    if (!h || !v || !originalFileBytes || mode === "none" || docCanvasEl?.classList.contains("hidden")) { clear(); return; }
    const data = await scan();
    if (host() !== h) return;
    const live = typeof dwbTrack !== "undefined" ? dwbTrack.liveChanges() : [];
    const comments = (data?.comments || []).filter((c) => !c.parentId);
    const changes = (data?.changes || []).filter((c) => c.part === "body" && Number.isFinite(c.paraIndex));
    if (!comments.length && !changes.length && !live.length) { clear(); return; }
    ensureLayer();
    frame();
    (data?.authors ? [...data.authors.keys()] : []).forEach(colorOf);
    const paras = collectPreviewParagraphElements(h);

    // geometria kolumny: od prawej krawędzi kartek do prawej krawędzi obszaru
    const secs = [...h.querySelectorAll("section.docx")];
    const pageR = Math.max(0, ...secs.map((s) => toContent(s.getBoundingClientRect()).r));
    const room = Math.max(v.clientWidth, v.scrollWidth) - pageR - 2 * 12;
    const colW = Math.min(COL_MAX, room);
    const cards = colW >= COL_MIN;
    const colX = pageR + 12;

    // podkreślenia wstawek (Wszystkie) i paski przy akapitach
    const hl = Array.from({ length: COLORS }, () => []);
    const bars = new Set();
    const groups = new Map(); // akapit → { p, author, rows: [] }
    changes.forEach((c) => {
      const p = paras[c.paraIndex];
      if (!p) return;
      bars.add(p);
      const k = colorOf(c.author);
      if (mode === "all" && (c.kind === "ins" || c.kind === "moveTo") && c.shown && c.shown.trim()) {
        const r = rangeIn(p, c.at, c.shown);
        if (r) hl[k].push(r);
      }
      const key = `g:${c.paraIndex}:${c.author}`;
      if (!groups.has(key)) groups.set(key, { key, p, author: c.author, date: c.date, rows: [] });
      groups.get(key).rows.push(c);
    });
    const me = live.length ? (dwbTrack.author() || t("trackAuthorDefault")) : "";
    const liveGroups = [];
    live.forEach((g) => {
      if (!g.p?.isConnected) return;
      const k = colorOf(me);
      const text = previewRunsToPlainText(extractRunsFromPreviewParagraph(g.p));
      const rows = [];
      g.ins.forEach((x) => {
        const s = text.slice(x.start, x.end);
        if (!s.trim()) return;
        if (mode === "all") { const r = rangeIn(g.p, x.start, s); if (r) hl[k].push(r); }
        rows.push({ kind: "ins", text: s, at: x.start });
      });
      g.dels.forEach((d) => rows.push({ kind: d.para ? "paraDel" : "del", text: d.text, at: d.at }));
      if (rows.length) liveGroups.push({ key: `l:${g.n}`, p: g.p, rows, live: true });
    });
    if (hlOk()) {
      clearHighlights();
      hl.forEach((rs, k) => { if (rs.length) CSS.highlights.set(`dwb-rev-${k}`, new Highlight(...rs)); });
    }
    inner.querySelectorAll(".mk-bar").forEach((b) => b.remove());
    bars.forEach((p) => {
      const c = toContent(p.getBoundingClientRect());
      const bar = document.createElement("div");
      bar.className = "mk-bar";
      bar.style.left = `${Math.max(2, toContent({ left: textLeft(p), top: 0, right: 0, height: 0 }).x - 14)}px`;
      bar.style.top = `${c.y}px`;
      bar.style.height = `${Math.max(4, c.h)}px`;
      inner.appendChild(bar);
    });

    // dymki
    const next = [];
    const marks = []; // telefon / wąsko: ikonki komentarzy na marginesie strony (jak w Wordzie)
    compact = !cards;
    // dymek zmiany POD tekstem (jak w Wordzie: „Zaakceptować czy odrzucić?”) zamiast w kolumnie:
    // gdy kolumny nie ma (telefon, wąsko) albo w trybie Proste (czysty tekst, bez bocznego pasa zmian)
    const popChanges = !cards || mode === "simple";
    {
      const ranges = new Map((typeof composeUi !== "undefined" && composeUi.commentRanges ? composeUi.commentRanges() : []).map((x) => [x.id, x.range]));
      comments.forEach((c) => {
        const range = ranges.get(c.id);
        let r = firstRect(range);
        if (!r && range?.startContainer?.isConnected) r = (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest("p")?.getBoundingClientRect();
        if (!r && Number.isFinite(c.paraIndex)) r = paras[c.paraIndex]?.getBoundingClientRect();
        if (!r) return;
        if (!cards) { marks.push({ id: c.id, author: c.author, done: c.done, anchor: toContent(r) }); return; }
        next.push({ key: `c:${c.id}`, kind: "comment", data: c, range, anchor: toContent(r), p: range?.startContainer?.isConnected ? (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest("p") : null });
      });
      [...groups.values()].forEach((g) => {
        const del = g.rows.find((x) => (x.kind === "del" || x.kind === "moveFrom") && Number.isFinite(x.at));
        const ins = g.rows.find((x) => (x.kind === "ins" || x.kind === "moveTo") && x.shown);
        const range = del ? rangeIn(g.p, del.at, "", 0) : ins ? rangeIn(g.p, ins.at, ins.shown) : null;
        const r = firstRect(range) || g.p.getBoundingClientRect();
        // same wstawki widać w tekście — ich dymek dopiero, gdy kursor stanie w tym akapicie (jak Word)
        if (g.rows.length > 1 && g.rows.some((x) => x.kind === "ins")) g.rows = g.rows.filter((x) => x.kind !== "paraIns");
        const quiet = g.rows.every((x) => x.kind === "ins" || x.kind === "moveTo" || x.kind === "paraIns");
        next.push({ key: g.key, kind: "change", data: g, range, anchor: toContent(r), quiet, p: g.p, pop: popChanges });
      });
      liveGroups.forEach((g) => {
        const first = g.rows[0];
        const range = rangeIn(g.p, first.at, first.kind === "ins" ? first.text : "", first.kind === "ins" ? first.text.length : 0);
        const r = firstRect(range) || g.p.getBoundingClientRect();
        next.push({ key: g.key, kind: "live", data: g, range, anchor: toContent(r), quiet: g.rows.every((x) => x.kind === "ins"), p: g.p, pop: popChanges });
      });
    }
    items.forEach((it) => it.el.remove());
    items = next;
    items.forEach((it) => { it.el = card(it); inner.appendChild(it.el); });
    const popW = Math.min(360, v.clientWidth - 24);
    items.forEach((it) => { it.el.style.width = `${it.pop ? popW : colW}px`; it.el.style.left = `${colX}px`; it.el.classList.toggle("is-compact", !!it.pop); });
    paintMarks(marks, pageR, bars);
    showing = items.length > 0;
    layer.classList.toggle("has-cards", showing);
    if (active && !items.some((it) => it.key === active)) active = null;
    if (!active) { const k = itemAtCaret(); if (k) active = k; }
    place();
  }

  // ── dymek ─────────────────────────────────────────────────────────────────
  function btn(icon, key, fn, cls = "") {
    const b = document.createElement("button");
    b.type = "button";
    b.className = `mk-btn ${cls}`.trim();
    b.innerHTML = svg(icon);
    b.setAttribute("aria-label", t(key));
    b.dataset.hint = "";
    b.dataset.hintPl = I18N.pl[key];
    b.dataset.hintEn = I18N.en[key];
    b.dataset.hintDelay = "0.4";
    b.addEventListener("click", (e) => { e.stopPropagation(); fn(b); });
    return b;
  }
  function head(author, date, extra) {
    const hd = document.createElement("div");
    hd.className = "mk-head";
    const av = document.createElement("span");
    av.className = "mk-avatar";
    av.textContent = initials({ author });
    av.setAttribute("aria-hidden", "true");
    const who = document.createElement("span");
    who.className = "mk-who";
    const name = document.createElement("b");
    name.textContent = author || "?";
    who.appendChild(name);
    const w = when(date);
    if (w) { const tm = document.createElement("time"); tm.textContent = w; if (date) tm.dateTime = date; who.appendChild(tm); }
    hd.append(av, who);
    if (extra) hd.appendChild(extra);
    return hd;
  }
  // Pasek zmiany na lewym marginesie STRONY (jak w Wordzie), nie przy akapicie — w tabeli stał
  // między kolumnami (zgłoszenie 2026-11-13). Akapit wysunięty w lewo poza tekst — przy nim.
  function textLeft(p) {
    const pr = p.getBoundingClientRect();
    const sec = p.closest("section.docx");
    if (!sec) return pr.left;
    const sr = sec.getBoundingClientRect();
    const k = sec.offsetWidth ? sr.width / sec.offsetWidth : 1;
    return Math.min(pr.left, sr.left + (parseFloat(getComputedStyle(sec).paddingLeft) || 0) * k);
  }
  // Kolor tekstu z pliku (Word: szary 595959, biały…) bywa nieczytelny na tle dymka (ciemny motyw)
  // — wtedy bez niego; tekst z wyróżnieniem zawsze ciemny (wyróżnienia są jasne).
  function readableRun(r) {
    if (r.hl) return { ...r, color: "111111" };
    if (!r.color) return r;
    // kontrast WCAG z tłem dymka (--surface: jasny #f2f5fb, ciemny #111827); poniżej 3:1 — bez koloru
    const L = (hex) => {
      const n = parseInt(hex, 16);
      const ch = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; };
      return 0.2126 * ch(n >> 16) + 0.7152 * ch((n >> 8) & 255) + 0.0722 * ch(n & 255);
    };
    const a = L(r.color), bg = L(document.documentElement.dataset.theme === "dark" ? "111827" : "F2F5FB");
    return (Math.max(a, bg) + 0.05) / (Math.min(a, bg) + 0.05) < 3 ? { ...r, color: undefined } : r;
  }
  function richBody(x) {
    const body = document.createElement("div");
    body.className = "mk-body";
    if (x.rich?.length && typeof composeUi !== "undefined" && composeUi.richRunEl) {
      x.rich.forEach((runs) => {
        const p = document.createElement("div");
        p.className = "mk-p";
        runs.forEach((r) => p.appendChild(composeUi.richRunEl(readableRun(r))));
        if (!runs.length) p.appendChild(document.createElement("br"));
        body.appendChild(p);
      });
    } else body.textContent = x.text;
    return body;
  }
  const ro = () => !!readOnlyMode;
  const kindLabel = (k) => t({ ins: "reviewKindIns", del: "reviewKindDel", moveFrom: "reviewKindMoveFrom", moveTo: "reviewKindMoveTo", format: "reviewKindFormat", paraFormat: "reviewKindParaFormat", paraIns: "reviewKindParaIns", paraDel: "reviewKindParaDel", rowIns: "reviewKindRowIns", rowDel: "reviewKindRowDel" }[k] || "reviewKindOtherFormat");

  async function revisions(action, ids) {
    if (!originalFileBytes || !ids.length) return;
    const n = await applyDocumentEdit({ op: "revisions", action, ids });
    if (n > 0) toastUndoable(t(action === "accept" ? "reviewAccepted" : "reviewRejected", { count: n }));
    else toast(t("editNothing"), "info");
    if (typeof runReviewScan === "function" && document.getElementById("panel-review")?.open) runReviewScan();
  }

  // Usunięcie i wstawka w tym samym miejscu (ten sam autor) = jedna decyzja „Zamieniono: 14 → 30”
  // (Word zapisuje to jako dwie poprawki; ✓ / ✗ działa na obie naraz). Dawniej dwa wiersze i trzy pary
  // przycisków, a po przyjęciu samego usunięcia dymek ze zostawioną wstawką się chował.
  function pairRows(rows) {
    const out = [];
    for (let i = 0; i < rows.length; i++) {
      const a = rows[i], b = rows[i + 1];
      const isDel = (x) => x && (x.kind === "del" || x.kind === "moveFrom");
      const isIns = (x) => x && (x.kind === "ins" || x.kind === "moveTo");
      const near = (x, y) => Number.isFinite(x.at) && Number.isFinite(y.at) && x.at === y.at && (x.author || "") === (y.author || "");
      if ((isDel(a) && isIns(b) && near(a, b)) || (isIns(a) && isDel(b) && near(a, b))) {
        const d = isDel(a) ? a : b, n = isIns(a) ? a : b;
        out.push({ kind: "replace", ids: [d.id, n.id].filter((x) => x != null), old: d.text || d.shown || "", text: n.text || n.shown || "", at: d.at, author: d.author });
        i++;
      } else out.push({ ...a, ids: a.id != null ? [a.id] : [] });
    }
    return out;
  }

  function card(it) {
    const el = document.createElement("div");
    el.className = `mk-card mk-${it.kind}`;
    el.dataset.key = it.key;
    el.tabIndex = 0;
    const c = it.data;
    if (it.kind === "comment") {
      el.style.setProperty("--mk-c", `var(--mk-a${colorOf(c.author)})`);
      el.setAttribute("aria-label", `${t("commentCard")}: ${c.author}`);
      if (c.done) el.classList.add("is-done");
      let tools = null;
      if (!ro()) {
        tools = document.createElement("span");
        tools.className = "mk-tools";
        tools.append(btn(c.done ? "reopen" : "resolve", c.done ? "commentReopen" : "commentResolve", () => composeUi.commentAction(c.id, "resolve")));
      }
      el.appendChild(head(c.author, c.date, tools));
      if (c.done) { const s = document.createElement("div"); s.className = "mk-done-note"; s.textContent = t("marginCollapsed"); el.appendChild(s); }
      el.appendChild(richBody(c));
      c.replies.forEach((r) => {
        const box = document.createElement("div");
        box.className = "mk-reply";
        box.style.setProperty("--mk-c", `var(--mk-a${colorOf(r.author)})`);
        box.append(head(r.author, r.date), richBody(r));
        if (!ro()) {
          const ed = btn("edit", "commentEdit", (b) => composeUi.commentAction(r.id, "edit", b), "mk-reply-edit");
          box.querySelector(".mk-head").appendChild(ed);
        }
        el.appendChild(box);
      });
      if (!ro()) {
        const act = document.createElement("div");
        act.className = "mk-actions";
        const reply = document.createElement("button");
        reply.type = "button";
        reply.className = "mk-reply-btn";
        reply.innerHTML = `${svg("reply")}<span></span>`;
        reply.querySelector("span").textContent = t("commentReply");
        reply.addEventListener("click", (e) => { e.stopPropagation(); composeUi.commentAction(c.id, "reply", reply); });
        act.append(reply,
          btn("edit", "commentEdit", (b) => composeUi.commentAction(c.id, "edit", b)),
          btn("del", "reviewDeleteComment", () => composeUi.commentAction(c.id, "delete"), "mk-danger"));
        el.appendChild(act);
      }
    } else {
      const g = c;
      const author = it.kind === "live" ? (dwbTrack.author() || t("trackAuthorDefault")) : g.author;
      el.style.setProperty("--mk-c", `var(--mk-a${colorOf(author)})`);
      el.setAttribute("aria-label", `${t("reviewChangesHead")}: ${author}`);
      let tools = null;
      const units = pairRows(g.rows);
      // jedna decyzja — ✓ / ✗ w nagłówku (jak karta w Wordzie); kilka — przy każdym wierszu
      const single = units.length === 1;
      if (it.kind === "change" && !ro() && single && !it.pop) {
        tools = document.createElement("span");
        tools.className = "mk-tools mk-tools-main";
        const ids = units[0].ids;
        tools.append(btn("ok", "reviewAcceptOne", () => revisions("accept", ids), "mk-ok"), btn("no", "reviewRejectOne", () => revisions("reject", ids), "mk-no"));
      }
      // wąsko: karta pod wierszem jak dymek w Wordzie — pytanie i ✓ / ✗ dla całej zmiany na górze
      if (it.pop && it.kind === "change" && !ro() && single) {
        const ask = document.createElement("div");
        ask.className = "mk-ask";
        const q = document.createElement("span");
        q.textContent = t("marginAsk");
        const ids = g.rows.map((x) => x.id);
        const a2 = document.createElement("span");
        a2.className = "mk-ask-act";
        a2.append(btn("ok", "reviewAcceptOne", () => revisions("accept", ids), "mk-ok"), btn("no", "reviewRejectOne", () => revisions("reject", ids), "mk-no"));
        ask.append(q, a2);
        el.appendChild(ask);
        tools = null;
      }
      el.appendChild(head(author, it.kind === "live" ? "" : g.date, tools));
      if (it.kind === "live") { const s = document.createElement("div"); s.className = "mk-live-tag"; s.textContent = t("marginLive"); el.querySelector(".mk-who").appendChild(s); }
      const clip = (v) => { const s = String(v || "").replace(/\s+/g, " ").trim(); return s.length > 140 ? `${s.slice(0, 139)}…` : s; };
      units.slice(0, 8).forEach((x) => {
        const row = document.createElement("div");
        row.className = `mk-row mk-row-${x.kind}`;
        const k = document.createElement("span");
        k.className = "mk-kind";
        const main = document.createElement("span");
        main.className = "mk-row-main";
        if (x.kind === "replace") {
          k.textContent = `${t("reviewKindReplace")}:`;
          const o = document.createElement("span"); o.className = "mk-txt mk-old"; o.textContent = clip(x.old);
          const n = document.createElement("span"); n.className = "mk-txt mk-new"; n.textContent = clip(x.text);
          main.append(k, " ", o, " → ", n);
        } else {
          const s = clip(x.text || x.shown);
          const tx = document.createElement("span");
          tx.className = "mk-txt";
          tx.textContent = s;
          k.textContent = `${x.kind === "paraDel" && it.kind === "live" ? t("marginParaDeleted") : kindLabel(x.kind)}${s ? ":" : ""}`;
          main.append(k);
          if (s) main.append(" ", tx);
        }
        row.append(main);
        if (it.kind === "change" && !ro() && !single && x.ids.length) {
          const a = document.createElement("span");
          a.className = "mk-row-act";
          a.append(btn("ok", "reviewAcceptOne", () => revisions("accept", x.ids), "mk-ok"), btn("no", "reviewRejectOne", () => revisions("reject", x.ids), "mk-no"));
          row.appendChild(a);
        }
        el.appendChild(row);
      });
      if (it.pop && ro() && it.kind === "change") { const n = document.createElement("div"); n.className = "mk-more"; n.textContent = t("marginRoNote"); el.appendChild(n); }
      if (units.length > 8) { const more = document.createElement("div"); more.className = "mk-more"; more.textContent = `+${units.length - 8}`; el.appendChild(more); }
      if (it.kind === "live" && !ro()) {
        const foot = document.createElement("div");
        foot.className = "mk-live-foot";
        const note = document.createElement("span");
        note.textContent = t("marginLiveNote");
        const show = document.createElement("button");
        show.type = "button";
        show.className = "btn mk-show";
        show.textContent = t("marginShowAll");
        show.addEventListener("click", (e) => { e.stopPropagation(); dwbTrack.bake({ keepOn: true }); });
        foot.append(note, show);
        el.appendChild(foot);
      }
    }
    el.addEventListener("mousedown", (e) => { if (e.target.closest("button")) e.preventDefault(); }); // fokus zostaje w dokumencie
    el.addEventListener("click", () => focusItem(it.key, { reveal: true }));
    el.addEventListener("keydown", (e) => { if ((e.key === "Enter" || e.key === " ") && e.target === el) { e.preventDefault(); focusItem(it.key, { reveal: true }); } });
    el.addEventListener("pointerenter", () => { hover = it.key; paintFocus(); drawWire(); });
    el.addEventListener("pointerleave", () => { if (hover === it.key) { hover = null; paintFocus(); drawWire(); } });
    return el;
  }

  // ── układ: aktywny przy swoim miejscu, reszta rozsuwa się w górę i w dół ──
  // Wąsko (telefon, Widok mobilny, duże powiększenie): obok kartki nie ma miejsca na kolumnę
  // dymków. Zmiana pokazuje się wtedy jako zwarta karta POD swoim wierszem — tylko aktywna
  // (kursor albo stuknięcie w akapit / znacznik przy pasku); komentarze mają ikonki przy
  // prawej krawędzi strony, stuknięcie otwiera kartę komentarza (compose-ui.js).
  let compact = false;
  function paintMarks(marks, pageR, bars) {
    inner.querySelectorAll(".mk-cmark, .mk-bar-hit").forEach((b) => b.remove());
    const v = vp();
    const right = Math.min(pageR, v.scrollLeft + v.clientWidth) - 30;
    const used = [];
    if (compact) marks.forEach((m) => {
      if (used.some((y) => Math.abs(y - m.anchor.y) < 6)) return; // kilka komentarzy w wierszu — jedna ikonka
      used.push(m.anchor.y);
      const b = document.createElement("button");
      b.type = "button";
      b.className = `mk-cmark${m.done ? " is-done" : ""}`;
      b.style.left = `${right}px`;
      b.style.top = `${m.anchor.y + Math.min(m.anchor.h, 22) / 2 - 13}px`;
      b.style.setProperty("--mk-c", `var(--mk-a${colorOf(m.author)})`);
      b.innerHTML = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>';
      b.setAttribute("aria-label", `${t("commentCard")}: ${m.author}`);
      b.addEventListener("mousedown", (e) => e.preventDefault());
      b.addEventListener("click", () => composeUi.showCommentById?.(m.id));
      inner.appendChild(b);
    });
    // pasek przy zmienionym akapicie: większy cel do stuknięcia (≥ 24 px) — pokazuje kartę zmiany
    bars.forEach((p) => {
      const it = items.find((x) => x.pop && x.p === p);
      if (!it) return;
      const c = toContent(p.getBoundingClientRect());
      const b = document.createElement("button");
      b.type = "button";
      b.className = "mk-bar-hit";
      b.style.left = `${Math.max(0, c.x - 26)}px`;
      b.style.top = `${c.y}px`;
      b.style.height = `${Math.max(24, c.h)}px`;
      b.setAttribute("aria-label", t("reviewChangesHead"));
      b.addEventListener("mousedown", (e) => e.preventDefault());
      b.addEventListener("click", () => focusItem(active === it.key ? null : it.key));
      inner.appendChild(b);
    });
  }
  function placePop() {
    const v = vp();
    const it = items.find((x) => x.pop && x.key === active);
    if (!it) return;
    const w = it.el.offsetWidth;
    const minX = v.scrollLeft + 12, maxX = v.scrollLeft + v.clientWidth - w - 12;
    it.el.style.left = `${Math.max(minX, Math.min(it.anchor.x - 16, maxX))}px`;
    const line = Math.min(it.anchor.h, 28);
    it.top = it.anchor.y + line + 8;
    it.h = it.el.offsetHeight;
    // pod wierszem; przy dolnej krawędzi widoku — nad nim
    if (it.top + it.h > v.scrollTop + v.clientHeight - 8 && it.anchor.y - it.h - 8 > v.scrollTop) it.top = it.anchor.y - it.h - 8;
    setCardY(it.el, it.top);
  }
  // Pionowo przez transform (GPU), nie top (układ). Pierwsze ustawienie bez przejścia: z „none”
  // dymek zjeżdżałby z góry kartki (top startował z auto, którego nie da się animować).
  function setCardY(el, y) {
    const v = `translateY(${Math.round(y)}px)`;
    if (el.style.transform) { el.style.transform = v; return; }
    el.style.transition = "none";
    el.style.transform = v;
    void el.offsetHeight; // zatwierdź pozycję, zanim wróci przejście
    el.style.transition = "";
  }
  function place() {
    if (!items.length) { if (wire) wire.replaceChildren(); return; }
    // najpierw stan (aktywny ma więcej przycisków = inną wysokość), potem pomiar; dymek pod tekstem — tylko aktywny
    items.forEach((it) => { it.el.hidden = (it.pop || !!it.quiet) && it.key !== active; it.el.classList.toggle("is-active", it.key === active); });
    placePop();
    const list = items.filter((it) => !it.pop && !it.el.hidden).sort((a, b) => a.anchor.y - b.anchor.y || a.anchor.x - b.anchor.x);
    list.forEach((it) => { it.h = it.el.offsetHeight; it.top = it.anchor.y - 4; });
    const ai = active ? list.findIndex((it) => it.key === active) : -1;
    if (!list.length) { paintFocus(); drawWire(); return; }
    if (ai < 0) {
      let floor = -Infinity;
      list.forEach((it) => { it.top = Math.max(it.top, floor); floor = it.top + it.h + GAP; });
    } else {
      let floor = list[ai].top + list[ai].h + GAP;
      for (let i = ai + 1; i < list.length; i++) { list[i].top = Math.max(list[i].top, floor); floor = list[i].top + list[i].h + GAP; }
      let ceil = list[ai].top - GAP;
      for (let i = ai - 1; i >= 0; i--) { list[i].top = Math.min(list[i].top, ceil - list[i].h); ceil = list[i].top - GAP; }
      // nad aktywnym brakuje miejsca — dymki nad nim wyjeżdżają w górę poza widok (jak w Wordzie);
      // wrócą na swoje miejsca, gdy aktywny przestanie nim być
    }
    // rozsuwanie dymków: transform liczy karta graficzna, top przeliczałby układ w każdej klatce
    list.forEach((it) => setCardY(it.el, it.top));
    paintFocus();
    drawWire();
  }
  // Linie jak w Wordzie: przy każdym usunięciu (dymek ze zmianą „usunięto”) przerywana linia
  // w kolorze autora od miejsca w tekście do dymka — stale; przy aktywnym / wskazanym dymku
  // (też komentarzu) linia wyraźniejsza.
  function drawWire() {
    if (!wire) return;
    wire.replaceChildren();
    const focus = hover || active;
    const hasDel = (it) => it.kind !== "comment" && it.data.rows.some((x) => /^(del|moveFrom|paraDel)$/.test(x.kind));
    // stała linia tylko gdy dymek stoi blisko swojego wiersza — dymek zepchnięty daleko dawał długie,
    // prawie pionowe linie przez cały margines; wtedy linia dopiero przy wskazaniu / aktywnym
    const near = (it) => Math.abs(it.top + 11 - it.anchor.y) < 48;
    const list = items.filter((it) => !it.pop && !it.el.hidden && (it.key === focus || (hasDel(it) && near(it))));
    if (!list.length) return;
    const h = host();
    const secs = h ? [...h.querySelectorAll("section.docx")] : [];
    const pageR = Math.max(0, ...secs.map((s) => toContent(s.getBoundingClientRect()).r));
    let w = 0, hh = 0;
    list.forEach((it) => {
      const y0 = it.anchor.y + 1; // górna krawędź wiersza (jak Word), nie po literach
      const x1 = parseFloat(it.el.style.left);
      const y1 = it.top + 11;
      const pl = document.createElementNS("http://www.w3.org/2000/svg", "polyline");
      pl.setAttribute("points", `${it.anchor.x},${y0} ${pageR},${y0} ${x1},${y1}`);
      pl.style.setProperty("--mk-c", it.el.style.getPropertyValue("--mk-c"));
      if (it.key === focus) pl.classList.add("is-focus");
      wire.appendChild(pl);
      w = Math.max(w, x1 + 4); hh = Math.max(hh, y0, y1);
    });
    wire.setAttribute("width", String(Math.max(w, inner.scrollWidth)));
    wire.setAttribute("height", String(hh + 4));
  }
  // podświetlenie miejsca w tekście aktywnego / wskazanego dymka
  function paintFocus() {
    if (!hlOk()) return;
    const it = items.find((x) => x.key === (hover || active));
    const r = it?.range;
    if (r && !r.collapsed && r.startContainer?.isConnected) CSS.highlights.set("dwb-mk-focus", new Highlight(r));
    else CSS.highlights.delete("dwb-mk-focus");
  }

  function focusItem(key, { reveal = false } = {}) {
    if (key && !items.some((it) => it.key === key)) key = null;
    if (active === key && !reveal) return !!key;
    active = key;
    place();
    if (reveal && key) {
      const it = items.find((x) => x.key === key);
      const v = vp();
      if (it && v) {
        const top = Math.min(it.anchor.y, it.top) - 40;
        const bottom = Math.max(it.anchor.y + it.anchor.h, it.top + it.h) + 40;
        if (top < v.scrollTop || bottom > v.scrollTop + v.clientHeight) v.scrollTo({ top: Math.max(0, top), behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth" });
      }
    }
    return !!key;
  }

  // Kursor (Edycja) albo klik (Czytanie) w tekście: dymek tego miejsca staje się aktywny —
  // komentarz pod kursorem, a poza komentarzem zmiana w tym akapicie.
  function itemAtCaret() {
    const sel = window.getSelection();
    if (!sel?.rangeCount || !items.length) return null;
    const node = sel.anchorNode, off = sel.anchorOffset;
    if (!host()?.contains(node)) return undefined; // fokus poza dokumentem (dymek, pasek) — bez zmian
    const inRange = (r) => { try { return !!r && r.startContainer.isConnected && r.isPointInRange(node, off); } catch (_) { return false; } };
    const c = items.find((it) => it.kind === "comment" && inRange(it.range));
    if (c) return c.key;
    const el = node.nodeType === 1 ? node : node.parentElement;
    const p = el?.closest("p");
    const g = p && items.find((it) => it.kind !== "comment" && it.p === p);
    return g ? g.key : null;
  }
  let selTimer = 0;
  function onSelection() {
    clearTimeout(selTimer);
    selTimer = setTimeout(() => {
      const k = itemAtCaret();
      if (k !== undefined && k !== active) focusItem(k);
    }, 60);
  }

  function setMode(m) {
    if (!/^(all|simple|none)$/.test(m)) return;
    mode = m;
    try { localStorage.setItem(MODE_KEY, m); } catch (_) {}
    syncSeg();
    schedule();
  }
  function syncSeg() {
    document.querySelectorAll("#rvMarkup [data-markup]").forEach((b) => {
      const on = b.dataset.markup === mode;
      b.classList.toggle("is-on", on);
      b.setAttribute("aria-checked", String(on));
    });
  }

  function init() {
    syncSeg();
    const seg = document.getElementById("rvMarkup");
    seg?.addEventListener("click", (e) => { const b = e.target.closest("[data-markup]"); if (b) setMode(b.dataset.markup); });
    seg?.addEventListener("keydown", (e) => {
      if (!["ArrowLeft", "ArrowRight"].includes(e.key)) return;
      const bs = [...seg.querySelectorAll("[data-markup]")];
      const i = bs.findIndex((b) => b.dataset.markup === mode);
      const n = bs[(i + (e.key === "ArrowRight" ? 1 : bs.length - 1)) % bs.length];
      e.preventDefault();
      setMode(n.dataset.markup);
      n.focus();
    });
    const v = vp();
    v?.addEventListener("scroll", () => { if (!layer) return; cancelAnimationFrame(raf); raf = requestAnimationFrame(frame); }, { passive: true });
    window.addEventListener("resize", () => { frame(); schedule(); });
    docCanvasEl?.addEventListener("dwb-zoom", schedule);
    docCanvasEl?.addEventListener("input", schedule);
    document.getElementById("readMode")?.addEventListener("change", schedule);
    if (v && typeof ResizeObserver === "function") new ResizeObserver(() => { frame(); schedule(); }).observe(v);
    // Tekst przesuwa się już PO policzeniu dymków (przerwy między stronami z page-breaks.js,
    // doczytane czcionki i obrazy, rozwinięty pasek): dymki, paski i linie stały przy starym
    // miejscu, a podkreślenie (CSS Highlight) szło za tekstem — dymek zmiany z pkt 7 przy
    // akapicie z pkt 6 (zgłoszenie 2026-11-13). Zmiana wysokości dokumentu = przeliczenie.
    if (docCanvasEl && typeof ResizeObserver === "function") new ResizeObserver(() => { if (layer) schedule(); }).observe(docCanvasEl);
    docCanvasEl?.addEventListener("load", () => { if (layer) schedule(); }, true); // obraz w dokumencie
    document.fonts?.addEventListener?.("loadingdone", () => { if (layer) schedule(); });
    document.addEventListener("dwb:undo-step", schedule);
    document.addEventListener("selectionchange", () => { if (showing) onSelection(); });
    // dymek pod tekstem: stuknięcie obok (poza dymkiem i znacznikiem przy pasku) go zamyka —
    // w Czytaniu na telefonie stuknięcie w tekst nie przesuwa kursora, więc samo zaznaczenie nie wystarcza
    document.addEventListener("pointerdown", (e) => {
      const it = items.find((x) => x.pop && x.key === active);
      if (!it || it.el.contains(e.target) || e.target.closest?.(".mk-bar-hit") || e.target.closest?.(".compose-pop, .app-dialog")) return;
      focusItem(null);
    }, true);
    // Esc w dymku — z powrotem do dokumentu
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape" || !layer?.contains(document.activeElement)) return;
      focusItem(null);
      docViewportEl?.focus({ preventScroll: true });
    });
  }

  return {
    init, schedule, setMode, focusItem,
    mode: () => mode,
    visible: () => showing,
    hasItem: (key) => showing && items.some((it) => it.key === key),
    _items: () => items.map((it) => ({ key: it.key, kind: it.kind, top: it.top, h: it.h, anchorY: it.anchor.y, left: parseFloat(it.el.style.left), width: it.el.offsetWidth })),
    // menu kontekstowe (doc-context-menu.js): śledzone zmiany w akapicie podglądu (z ostatniego skanu
    // TEJ wersji pliku) i Akceptuj / Odrzuć jak w dymku
    revisionsAt: (p) => {
      if (!p || cache.bytes !== originalFileBytes || !cache.data) return [];
      const paras = collectPreviewParagraphElements(host());
      return cache.data.changes.filter((c) => c.part === "body" && Number.isFinite(c.paraIndex) && paras[c.paraIndex] === p);
    },
    revisions,
    _rangeIn: rangeIn,
    _readable: readableRun,
  };
})();

document.addEventListener("DOMContentLoaded", () => dwbMargin.init());
