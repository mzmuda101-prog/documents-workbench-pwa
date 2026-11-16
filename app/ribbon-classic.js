// ribbon-classic.js — wstążka Klasyczna (etap 6b; docs/WSTAZKA.md, makieta v4 „Klasyczna”).
//
// Duże przyciski (ikona nad podpisem), kolumny małych przycisków z podpisami, nazwy grup pod
// spodem i ↘ „więcej ustawień” — układ każdej karty jak w makiecie. Buduje panele .rb-panel.rb-cl
// obok paneli Uproszczonej; ribbon.js pokazuje jedne albo drugie.
//
// Bez kopii logiki:
//   - przyciski z data-rb-act to te same akcje co w Uproszczonej (obsługa kliknięć w ribbon.js);
//     ikona, podpis i podpowiedź są brane z przycisku Uproszczonej (ta sama funkcja = ta sama ikona),
//   - kontrolki z id (B, I, U, listy kroju…) to TE SAME elementy: w Klasycznej przenoszą się do
//     gniazd [data-rbc] (ribbon.js applyLayout), więc ich stan i obsługa zostają,
//   - „click:#id” klika istniejący przycisk (np. Tryb skupienia) i pokazuje jego stan.

const dwbRibbonClassic = (() => {
  const root = document.getElementById("ribbon");
  const host = root?.querySelector('[data-strip="cmds"] .rb-sc');
  if (!host) return null;

  const svg = (d, size = 16) => `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
  // ikony makiety dla przycisków, których Uproszczona nie ma
  const IC = {
    search: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>',
    replace: '<path d="M14 4c0-1.1.9-2 2-2h4a2 2 0 0 1 2 2v4a2 2 0 0 1-2 2h-4a2 2 0 0 1-2-2M14 14a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v4a2 2 0 0 1-2 2h-4a2 2 0 0 1-2-2M4 6v4a2 2 0 0 0 2 2h4M7 9l3 3-3 3"/>',
    a11y: '<circle cx="12" cy="5" r="2"/><path d="M5 9h14M12 9v5M9 22l3-8 3 8"/>',
    globe: '<circle cx="12" cy="12" r="10"/><path d="M2 12h20M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>',
    track: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><path d="M12 6.5v4M10 8.5h4M10 13h4"/>',
    cmt: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
    spell: '<path d="m3 15 4-8 4 8M4 13h6M15 11h4.5a2 2 0 0 1 0 4H15V7h4a2 2 0 0 1 0 4"/><path d="m14 20 2 2 4-4"/>',
    navp: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18M13 8h4M13 12h4"/>',
    ask: '<circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3M12 17h.01"/>',
    sense: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3M8 11h.01M11 11h.01M14 11h.01"/>',
    info: '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/>',
    merge: '<rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/>',
    rowA: '<rect x="3" y="12" width="18" height="8" rx="1"/><line x1="12" y1="3" x2="12" y2="9"/><line x1="9" y1="6" x2="15" y2="6"/>',
    rowB: '<rect x="3" y="4" width="18" height="8" rx="1"/><line x1="12" y1="15" x2="12" y2="21"/><line x1="9" y1="18" x2="15" y2="18"/>',
    colL: '<rect x="12" y="3" width="8" height="18" rx="1"/><line x1="3" y1="12" x2="9" y2="12"/><line x1="6" y1="9" x2="6" y2="15"/>',
    colR: '<rect x="4" y="3" width="8" height="18" rx="1"/><line x1="15" y1="12" x2="21" y2="12"/><line x1="18" y1="9" x2="18" y2="15"/>',
    launch: '<path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/>',
    up: '<path d="m18 15-6-6-6 6"/>',
    down: '<path d="m6 9 6 6 6-6"/>',
    marks: '<path d="M13 4v16M17 4v16M19 4H9.5a4.5 4.5 0 0 0 0 9H13"/>',
    more: '<circle cx="5" cy="12" r="1.5"/><circle cx="12" cy="12" r="1.5"/><circle cx="19" cy="12" r="1.5"/>',
  };
  const DD = '<svg class="rb-dd" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" aria-hidden="true"><path d="m6 9 6 6 6-6"/></svg>';
  const alignIcon = (k) => (typeof composeUi !== "undefined" && composeUi.alignSvg ? composeUi.alignSvg(k) : svg(""));

  // ── opis kart (jak w makiecie) ────────────────────────────────────────────
  const B = (act, o = {}) => ({ k: "big", act, ...o });
  const S = (act, o = {}) => ({ k: "sm", act, ...o });
  const I = (act, o = {}) => ({ k: "ic", act, ...o });
  const L = (sel, o = {}) => ({ k: "slot", sel, ...o });
  const ST = (...items) => ({ k: "stack", items });
  const RW = (...items) => ({ k: "row", items });
  const G = (key, items, launch) => ({ key, items, launch });
  const TXT = (key, icon) => ({ k: "txt", key, icon });
  const ALIGN = (k, key) => I(`align:${k}`, { iconHtml: alignIcon(k), key });

  const SPEC = {
    home: [
      G("rbGClipboard", [B("paste", { dd: 1 }), ST(S("cut"), S("copy"), L("#fmtPainterBtn", { lbl: 1 }))]),
      G("rbGFont", [ST(
        RW(L("#fmtFontFamily"), L("#fmtFontSize"), L("#fmtGrow"), L("#fmtShrink"), L("#fmtCaseBtn")),
        RW(L("#fmtBold"), L("#fmtItalic"), L("#fmtUnderline"), L("#fmtStrike"), L("#fmtSub"), L("#fmtSuper"), I("highlight"), L("#fmtColorBtn"), L("#fmtClearBtn")),
      )], "fontForm"),
      G("rbGPara", [ST(
        RW(I("bullets"), I("numbering"), L("#fmtListBtn"), I("outdent"), I("indent"), L("#formatMarksBtn")),
        RW(ALIGN("left", "alignLeft"), ALIGN("center", "alignCenter"), ALIGN("right", "alignRight"), ALIGN("both", "alignBoth"), L("#fmtSpacingBtn")),
      )], "layoutForm"),
      G("rbGStyles", [{ k: "styles" }]),
      G("rbGEditing", [ST(S("find", { icon: IC.search, key: "rbFind" }), S("replace", { icon: IC.replace, key: "rbReplace" }), S("panel:panel-edit-tools"))]),
    ],
    insert: [
      G("rbGPages", [B("pageBreak"), B("hf")]),
      G("rbGTables", [B("table", { dd: 1 })]),
      G("rbGIllus", [B("image", { dd: 1 })]),
      G("rbGLinks", [ST(S("link"), S("footnote"), S("endnote")), ST(S("bookmark"), S("comment"))]),
      G("rbGText", [B("textbox"), B("toc"), ST(S("hrule"), S("date"), S("panel:panel-snippets"))]),
      G("rbGSymbols", [B("symbol", { dd: 1 })]),
      G("rbGFormFields", [B("formFields", { dd: 0 })]),
    ],
    layout: [
      G("rbGPage", [B("ps:margins", { dd: 1 }), B("ps:orient", { dd: 1 }), B("ps:size", { dd: 1 }), B("ps:cols", { dd: 1 }), ST(S("ps:breaks", { dd: 1 }), S("ps:valign", { dd: 1 }))], "pageSetup"),
      G("rbGParaLayout", [B("layoutForm")], "layoutForm"),
      G("rbGTabs", [B("tabsForm")]),
    ],
    review: [
      G("rbGProof", [B("panel:panel-grammar", { icon: IC.spell, key: "railProof" }), ST(S("panel:panel-stats"), S("a11y", { icon: IC.a11y, key: "rbA11y" }))]),
      G("rbGLang", [B("lang", { icon: IC.globe, key: "rbLang", dd: 1 })]),
      G("rbGComments", [B("comment"), ST(S("removeComments"))]),
      G("rbGTracking", [B("suggest", { icon: IC.track, key: "modeSuggest" }), ST(S("markupMenu"), S("panel:panel-review", { icon: IC.cmt, key: "rbReviewPane" }))]),
      G("rbGChanges", [B("acceptAll"), B("rejectAll"), ST(S("walk"))]),
      G("rbGCompare", [B("panel:panel-compare")]),
    ],
    mail: [
      G("rbGFields", [B("panel:panel-placeholders")]),
      G("rbGMerge", [B("click:#mmPickBtn", { icon: IC.merge, key: "rbMerge" })]),
      G("rbGForms", [B("panel:panel-forms")]),
    ],
    view: [
      G("rbGModes", [B("click:#focusModeBtn"), ST(S("fullscreen"))]),
      G("rbGShow", [ST(S("panel:panel-structure", { icon: IC.navp, key: "rbNavPane" }), S("sectionStrip"), S("click:#formatMarksBtn", { icon: IC.marks }))]),
    ],
    ai: [
      G("rbGWriting", [B("click:#rewriteBtn", { cls: "rb-ai" })]),
      G("rbGQuestions", [B("panel:panel-ask", { icon: IC.ask, key: "askTitle", cls: "rb-ai" }), B("panel:panel-semantic-search", { icon: IC.sense, key: "semanticPanel", cls: "rb-ai" })]),
      G("rbGModel", [ST(S("aiSettings"), TXT("rbAiLocal", IC.info))]),
    ],
    table: [
      G("rbGRowsCols", [B("tbl:rowAbove", { icon: IC.rowA, key: "tableRowAbove" }), B("tbl:rowBelow", { icon: IC.rowB, key: "tableRowBelow" }),
        ST(S("tbl:colLeft", { icon: IC.colL, key: "tableColLeft" }), S("tbl:colRight", { icon: IC.colR, key: "tableColRight" }), S("tblp:delete", { dd: 1 }))]),
      G("rbGMergeCells", [ST(S("tblp:merge", { dd: 1 }))]),
      G("rbGTableStyle", [ST(S("tbl:headerRow"), S("tblp:borders", { dd: 1 }), S("tblp:shade", { dd: 1 })), ST(S("tblp:align", { dd: 1 }), S("tblp:textdir", { dd: 1 }), S("tbl:evenCols"))]),
      G("rbGData", [B("tblp:sort")]),
    ],
  };

  // ── budowa ────────────────────────────────────────────────────────────────
  const simple = (act) => root.querySelector(`.rb-panel:not(.rb-cl) [data-rb-act="${CSS.escape(act)}"]`);
  const mirrors = []; // [kopia, oryginał] — stan „wciśnięty” przycisków click:#id
  function labelKey(src) { return src?.dataset.i18nAria || src?.getAttribute("data-i18n") || ""; }
  function button(it) {
    const src = it.act.startsWith("click:") ? document.querySelector(it.act.slice(6)) : simple(it.act);
    const key = it.key || labelKey(src);
    const b = document.createElement("button");
    b.type = "button";
    b.className = `tb-btn rb-btn rb-cb ${it.k === "big" ? "rb-big" : it.k === "ic" ? "rb-ic" : "rb-sm"} ${it.cls || ""}`.trim();
    b.dataset.rbAct = it.act;
    if (key) { b.dataset.i18nAria = key; b.setAttribute("aria-label", t(key)); }
    for (const k of ["hint", "hintPl", "hintEn", "hintDelay"]) if (src?.dataset[k] != null) b.dataset[k] = src.dataset[k];
    if (src?.hasAttribute("aria-haspopup")) b.setAttribute("aria-haspopup", src.getAttribute("aria-haspopup"));
    const icon = it.iconHtml || (it.icon ? svg(it.icon) : src?.querySelector("svg:not(.rb-dd)")?.outerHTML || "");
    b.innerHTML = icon;
    b.querySelector("svg")?.removeAttribute("id");
    if (it.k !== "ic" && key) {
      const s = document.createElement("span");
      s.className = "rb-cl-t";
      s.dataset.i18n = key;
      s.textContent = t(key);
      b.append(s);
    }
    if (it.dd ?? src?.querySelector(".rb-dd")) b.insertAdjacentHTML("beforeend", DD);
    src?.querySelectorAll(".rb-bar").forEach((x) => b.append(x.cloneNode(true)));
    if (it.act.startsWith("click:") && src) mirrors.push([b, src]);
    return b;
  }
  const classicSlots = [];
  function slot(it) {
    const s = document.createElement("span");
    s.className = "rb-slot";
    s.dataset.rbc = it.sel;
    classicSlots.push(s);
    if (!it.lbl) return s;
    // mały przycisk z podpisem (np. Malarz formatów): podpis obok prawdziwej kontrolki, klik też działa
    const wrap = document.createElement("span");
    wrap.className = "rb-cl-item";
    const el = document.querySelector(it.sel);
    const key = labelKey(el);
    const lbl = document.createElement("span");
    lbl.className = "rb-cl-lbl";
    lbl.dataset.i18n = key;
    lbl.textContent = t(key);
    lbl.addEventListener("click", () => document.querySelector(it.sel)?.click());
    wrap.append(s, lbl);
    return wrap;
  }
  function txt(it) {
    const s = document.createElement("span");
    s.className = "rb-cl-txt";
    s.innerHTML = svg(it.icon);
    const l = document.createElement("span");
    l.dataset.i18n = it.key;
    l.textContent = t(it.key);
    s.append(l);
    return s;
  }
  function node(it) {
    if (it.k === "stack" || it.k === "row") {
      const d = document.createElement("div");
      d.className = it.k === "stack" ? "rb-cst" : "rb-crw";
      d.append(...it.items.map(node));
      return d;
    }
    if (it.k === "slot") return slot(it);
    if (it.k === "txt") return txt(it);
    if (it.k === "styles") return styles();
    return button(it);
  }
  function group(g) {
    const d = document.createElement("div");
    d.className = "rb-cg";
    d.setAttribute("role", "group");
    d.dataset.i18nAria = g.key;
    d.setAttribute("aria-label", t(g.key));
    const items = document.createElement("div");
    items.className = "rb-gi";
    items.append(...g.items.map(node));
    const cap = document.createElement("div");
    cap.className = "rb-gcap";
    const name = document.createElement("span");
    name.dataset.i18n = g.key;
    name.textContent = t(g.key);
    cap.append(name);
    if (g.launch) {
      const l = document.createElement("button");
      l.type = "button";
      l.className = "rb-launch";
      l.dataset.rbAct = g.launch;
      l.dataset.i18nAria = "rbLaunch";
      l.setAttribute("aria-label", t("rbLaunch"));
      l.dataset.hint = "";
      l.dataset.hintPl = `${I18N.pl[g.key] || ""} — ${I18N.pl.rbLaunch}`;
      l.dataset.hintEn = `${I18N.en[g.key] || ""} — ${I18N.en.rbLaunch}`;
      l.dataset.hintDelay = "0.6";
      l.innerHTML = svg(IC.launch, 11);
      cap.append(l);
    }
    d.append(items, cap);
    return d;
  }

  // ── galeria stylów (te same style co lista „Styl akapitu”, klik = composeUi.applyStyle) ──
  let gallery = null;
  function styles() {
    const wrap = document.createElement("div");
    wrap.className = "rb-sty";
    const list = document.createElement("div");
    list.className = "rb-sty-list";
    const more = document.createElement("div");
    more.className = "rb-sty-more";
    const mk = (ic, key, fn) => {
      const b = document.createElement("button");
      b.type = "button";
      b.innerHTML = svg(IC[ic], 12);
      b.dataset.i18nAria = key;
      b.setAttribute("aria-label", t(key));
      b.addEventListener("click", fn);
      return b;
    };
    more.append(
      mk("up", "rbStylesPrev", () => list.scrollBy({ left: -list.clientWidth, behavior: "smooth" })),
      mk("down", "rbStylesNext", () => list.scrollBy({ left: list.clientWidth, behavior: "smooth" })),
      mk("more", "rbStylesAll", (e) => allStyles(e.currentTarget)),
    );
    wrap.append(list, more);
    gallery = list;
    return wrap;
  }
  const styleSel = () => document.getElementById("fmtParaStyle");
  const styleOpts = () => Array.from(styleSel()?.options || []).filter((o) => o.value && !o.value.startsWith("act:") && !o.value.startsWith("plus:") && !o.dataset.plus && !o.dataset.cur && !o.disabled && !o.hidden);
  function allStyles(anchor) {
    const sel = styleSel();
    if (typeof composeUi === "undefined" || !sel) return;
    composeUi.openMenu(anchor, styleOpts().map((o) => ({ label: o.textContent.trim(), checked: o.value === sel.value, onPick: () => composeUi.applyStyle(o.value) })));
  }
  function renderStyles() {
    if (!gallery) return;
    const sel = styleSel();
    gallery.replaceChildren(...styleOpts().map((o) => {
      const name = o.textContent.trim();
      const kind = /heading|nagłówek|^h\d/i.test(o.value + " " + name) ? "h" : /title|tytuł/i.test(o.value + " " + name) ? "t" : /quote|cytat/i.test(o.value + " " + name) ? "q" : "n";
      const b = document.createElement("button");
      b.type = "button";
      b.className = `rb-sty-i is-${kind}`;
      b.dataset.style = o.value;
      b.setAttribute("aria-pressed", String(o.value === sel.value));
      const pv = document.createElement("b");
      pv.textContent = kind === "n" ? "Aa" : name;
      const nm = document.createElement("small");
      nm.textContent = name;
      b.append(pv, nm);
      b.addEventListener("mousedown", (e) => e.preventDefault());
      b.addEventListener("click", () => composeUi?.applyStyle(o.value));
      return b;
    }));
  }
  function syncStyles() {
    const v = styleSel()?.value;
    gallery?.querySelectorAll(".rb-sty-i").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.style === v)));
  }

  for (const [tab, groups] of Object.entries(SPEC)) {
    const p = document.createElement("div");
    p.className = "rb-panel rb-cl";
    p.setAttribute("role", "tabpanel");
    p.dataset.rtab = tab;
    p.setAttribute("aria-labelledby", `rbTab-${tab}`);
    p.hidden = true;
    p.append(...groups.map(group));
    host.append(p);
  }
  renderStyles();
  const ss = styleSel();
  if (ss) new MutationObserver(renderStyles).observe(ss, { childList: true, subtree: true });
  let sq = 0;
  document.addEventListener("selectionchange", () => { cancelAnimationFrame(sq); sq = requestAnimationFrame(syncStyles); });
  ss?.addEventListener("change", syncStyles);
  // stan „wciśnięty” kopii (Tryb skupienia, Znaki formatowania) za oryginałem
  for (const [b, src] of mirrors) {
    const sync = () => {
      const onNow = src.classList.contains("is-on") || src.getAttribute("aria-pressed") === "true";
      b.classList.toggle("is-on", onNow);
      if (src.hasAttribute("aria-pressed") || src.classList.contains("is-on")) b.setAttribute("aria-pressed", String(onNow));
      b.hidden = src.hidden;
    };
    new MutationObserver(sync).observe(src, { attributes: true, attributeFilter: ["class", "aria-pressed", "hidden"] });
    sync();
  }

  return { slots: classicSlots, syncStyles };
})();
