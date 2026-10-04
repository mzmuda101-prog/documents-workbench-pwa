// paste-rich.js — wklejanie z zachowaniem bezpiecznych stylów (Apple Notes, Markdown, strony WWW,
// Google Docs, Word): akapity, nagłówki, listy punktowane / numerowane z poziomami, listy kontrolne
// (prawdziwe pola wyboru), cytaty, kod, tabele, linia pozioma; w tekście: pogrubienie, kursywa,
// podkreślenie, przekreślenie, wyróżnienie, link, kod (czcionka stała).
// NIE bierzemy krojów, rozmiarów ani kolorów ze źródła — tekst przyjmuje styl dokumentu (jak
// „Scal formatowanie” w Wordzie). Obrazy, skrypty, style, ramki — pomijane.
//
// Model: blok = { type: "p" | "h" | "li" | "quote" | "code" | "hr" | "table", level, ordered,
// checked, runs: [{ text, bold, italic, underline, strike, highlight, link, code } | { break }],
// rows: [[runs]], head }. Jeden zwykły akapit → wstawienie w miejscu kursora (bez przerysowania);
// więcej → op „pasteBlocks” (docx-compose.js) — jeden krok Cofnij.

const dwbPaste = (() => {
  const SAFE_LINK = /^(https?:|mailto:|tel:)/i;
  const MONO = /(mono|courier|menlo|monaco|consolas|sfmono|sf mono|source code)/i;

  // ── wspólne ────────────────────────────────────────────────────────────────
  function pushText(runs, text, style) {
    if (!text) return;
    const last = runs[runs.length - 1];
    if (last && !last.break && sameStyle(last, style)) last.text += text;
    else runs.push({ text, ...style });
  }
  function sameStyle(a, b) {
    return ["bold", "italic", "underline", "strike", "highlight", "code", "link"].every((k) => (a[k] || "") === (b[k] || ""));
  }
  // przycięcie spacji na brzegach bloku i podwójnych spacji (HTML zwija białe znaki)
  function tidy(runs) {
    const out = [];
    for (const r of runs) {
      if (r.break) { out.push(r); continue; }
      let text = r.text.replace(/[ \s]+/g, " ");
      const prev = out[out.length - 1];
      if ((!prev || prev.break || /\s$/.test(prev.text || "")) && !r.code) text = text.replace(/^ /, "");
      if (text) out.push({ ...r, text });
    }
    while (out.length && out[out.length - 1].break) out.pop();
    while (out.length && out[0].break) out.shift();
    const last = out[out.length - 1];
    if (last && !last.code) last.text = last.text.replace(/ $/, "");
    return out.filter((r) => r.break || r.text);
  }
  const plain = (runs) => runs.map((r) => (r.break ? "\n" : r.text)).join("");

  // ── HTML ───────────────────────────────────────────────────────────────────
  function fromHtml(html) {
    // DOMParser tworzy dokument bez wykonywania skryptów i bez ładowania obrazków
    const doc = new DOMParser().parseFromString(html, "text/html");
    doc.querySelectorAll("script, style, meta, link, title, template, svg, iframe, object, embed, img, video, audio, canvas, noscript, head").forEach((n) => n.remove());
    const blocks = [];
    let cur = null; // bieżący blok tekstu
    const flush = () => {
      if (cur) {
        cur.runs = tidy(cur.runs);
        if (cur.type === "li" || cur.runs.length) blocks.push(cur);
      }
      cur = null;
    };
    const open = (type, extra = {}) => { flush(); cur = { type, runs: [], ...extra }; return cur; };
    const ensure = (ctx) => cur || open(ctx.quote ? "quote" : "p");

    function inlineStyle(el, base) {
      const st = { ...base };
      const tag = el.localName;
      const css = (el.getAttribute("style") || "").toLowerCase();
      const fw = /font-weight\s*:\s*([a-z0-9]+)/.exec(css)?.[1];
      if (tag === "b" || tag === "strong") st.bold = !(fw && (fw === "normal" || +fw < 600)); // Google Docs: <b style="font-weight:normal">
      else if (fw) st.bold = fw === "bold" || fw === "bolder" || +fw >= 600;
      if (tag === "i" || tag === "em" || /font-style\s*:\s*italic/.test(css)) st.italic = true;
      if (tag === "u" || tag === "ins" || /text-decoration[^;]*underline/.test(css)) st.underline = true;
      if (tag === "s" || tag === "strike" || tag === "del" || /text-decoration[^;]*line-through/.test(css)) st.strike = true;
      if (tag === "mark") st.highlight = "yellow";
      if (tag === "code" || tag === "tt" || tag === "kbd" || tag === "samp" || MONO.test(/font-family\s*:\s*([^;]+)/.exec(css)?.[1] || "")) st.code = true;
      if (tag === "a") {
        const href = (el.getAttribute("href") || "").trim();
        if (SAFE_LINK.test(href)) st.link = href;
      }
      return st;
    }

    const HEAD = { h1: 1, h2: 2, h3: 3, h4: 3, h5: 3, h6: 3 };
    function walk(node, style, ctx) {
      if (node.nodeType === 3) {
        const text = ctx.pre ? node.textContent : node.textContent.replace(/[\r\n\t ]+/g, " ");
        if (!text.trim() && !cur && !ctx.pre) return;
        if (ctx.pre) text.split("\n").forEach((line, i) => { if (i) ensure(ctx).runs.push({ break: true }); pushText(ensure(ctx).runs, line, style); });
        else pushText(ensure(ctx).runs, text, style);
        return;
      }
      if (node.nodeType !== 1) return;
      const el = node;
      const tag = el.localName;
      const css = (el.getAttribute("style") || "").toLowerCase();
      if (/mso-list\s*:\s*ignore/.test(css)) return; // Word: znak punktora jako tekst
      if (/display\s*:\s*none/.test(css)) return;
      if (tag === "br") { if (cur) cur.runs.push({ break: true }); return; }
      if (tag === "hr") { flush(); blocks.push({ type: "hr", runs: [] }); return; }
      if (tag === "input" && el.getAttribute("type") === "checkbox") {
        if (cur && cur.type === "li") cur.checked = el.checked || el.hasAttribute("checked");
        return;
      }
      if (HEAD[tag]) {
        open("h", { level: HEAD[tag] });
        el.childNodes.forEach((c) => walk(c, { ...inlineStyle(el, style), bold: false }, ctx));
        flush();
        return;
      }
      if (tag === "ul" || tag === "ol") {
        flush();
        const depth = ctx.list ? ctx.list.depth + 1 : 0;
        // Apple Notes: lista kontrolna jako ul.checklist / li z atrybutem stanu
        const checklist = /check/i.test(el.className || "");
        el.childNodes.forEach((c) => walk(c, style, { ...ctx, list: { depth, ordered: tag === "ol", checklist } }));
        flush();
        return;
      }
      if (tag === "li") {
        const L = ctx.list || { depth: 0, ordered: false };
        const done = /checked|done/i.test(el.className || "") || el.getAttribute("aria-checked") === "true";
        open("li", { level: Math.min(8, L.depth), ordered: L.ordered, checked: L.checklist ? done : undefined });
        el.childNodes.forEach((c) => walk(c, style, { ...ctx, list: L, inLi: true }));
        flush();
        return;
      }
      // Word (HTML ze schowka): akapity listy jako p.MsoListParagraph… z „mso-list: l0 level2”
      if ((tag === "p" || tag === "div") && /mso-list\s*:\s*l\d+\s+level(\d)/.test(css)) {
        const lvl = +/level(\d)/.exec(css)[1] - 1;
        open("li", { level: Math.max(0, Math.min(8, lvl)), ordered: false });
        el.childNodes.forEach((c) => walk(c, style, ctx));
        flush();
        return;
      }
      if (tag === "table") {
        flush();
        const rows = [];
        let head = false;
        el.querySelectorAll(":scope > tr, :scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr").forEach((tr, ri) => {
          const cells = [];
          tr.querySelectorAll(":scope > td, :scope > th").forEach((td) => {
            if (ri === 0 && td.localName === "th") head = true;
            // treść komórki: bloki → jeden akapit (wiersze = łamania)
            const sub = fromHtml(td.innerHTML);
            const runs = [];
            sub.forEach((b, i) => { if (i) runs.push({ break: true }); (b.type === "table" ? [{ text: b.rows.map((r) => r.map(plain).join("\t")).join(" ") }] : b.runs).forEach((r) => runs.push(r)); });
            cells.push(tidy(runs));
          });
          if (cells.length) rows.push(cells);
        });
        if (rows.length) blocks.push({ type: "table", rows, head, runs: [] });
        return;
      }
      if (tag === "pre") {
        open("code");
        el.childNodes.forEach((c) => walk(c, { ...style, code: true }, { ...ctx, pre: true }));
        flush();
        return;
      }
      if (tag === "blockquote") {
        flush();
        el.childNodes.forEach((c) => walk(c, style, { ...ctx, quote: true }));
        flush();
        return;
      }
      const block = /^(p|div|section|article|header|footer|address|figure|figcaption|dd|dt|aside|main|nav|center)$/.test(tag);
      if (block && !ctx.inLi) flush();
      if (block && ctx.inLi && cur && plain(cur.runs).trim()) cur.runs.push({ break: true }); // kolejny akapit w punkcie listy
      el.childNodes.forEach((c) => walk(c, inlineStyle(el, style), ctx));
      if (block && !ctx.inLi) flush();
    }
    walk(doc.body, {}, {});
    flush();
    // pusty punkt listy zostaje tylko jako pole wyboru; puste akapity (odstępy w Notatkach) — precz
    return blocks.filter((b) => b.type !== "li" || b.runs.length || b.checked !== undefined);
  }

  // ── Markdown ───────────────────────────────────────────────────────────────
  const MD_BLOCK = /^(#{1,6}\s|\s*[-*+]\s|\s*\d{1,3}[.)]\s|>\s?|```|~~~|\s*\|.*\|\s*$|\s*(?:[-*_]\s*){3,}$)/m;
  const MD_INLINE = /\*\*[^*\n]+\*\*|__[^_\n]+__|~~[^~\n]+~~|==[^=\n]+==|\[[^\]\n]+\]\([^)\s]+\)|`[^`\n]+`|(?<![\p{L}\p{N}*])\*[^*\s][^*\n]*\*(?![\p{L}\p{N}*])/u;
  function looksLikeMarkdown(text) {
    return MD_BLOCK.test(text) || MD_INLINE.test(text);
  }

  function mdInline(src, base = {}) {
    const runs = [];
    const RULES = [
      ["code", /`([^`\n]+)`/],
      ["link", /\[([^\]\n]+)\]\(([^)\s]+)(?:\s+"[^"]*")?\)/],
      ["auto", /<((?:https?:|mailto:)[^>\s]+)>/],
      ["bold", /\*\*([^\n]+?)\*\*|__([^\n]+?)__/],
      ["strike", /~~([^\n]+?)~~/],
      ["highlight", /==([^\n]+?)==/],
      ["italic", /(?<![\p{L}\p{N}*])\*([^*\s](?:[^*\n]*[^*\s])?)\*(?![\p{L}\p{N}*])|(?<![\p{L}\p{N}_])_([^_\s](?:[^_\n]*[^_\s])?)_(?![\p{L}\p{N}_])/u],
    ];
    let rest = src;
    while (rest) {
      let best = null;
      for (const [kind, re] of RULES) {
        const m = re.exec(rest);
        if (m && (!best || m.index < best.m.index)) best = { kind, m };
      }
      if (!best) { pushText(runs, rest.replace(/\\([\\`*_{}[\]()#+\-.!~=|>])/g, "$1"), base); break; }
      const { kind, m } = best;
      if (m.index) pushText(runs, rest.slice(0, m.index).replace(/\\([\\`*_{}[\]()#+\-.!~=|>])/g, "$1"), base);
      if (kind === "code") pushText(runs, m[1], { ...base, code: true });
      else if (kind === "link" || kind === "auto") {
        const href = kind === "auto" ? m[1] : m[2];
        const text = kind === "auto" ? m[1].replace(/^mailto:/i, "") : m[1];
        mdInline(text, SAFE_LINK.test(href) ? { ...base, link: href } : base).forEach((r) => runs.push(r));
      } else {
        const inner = m[1] ?? m[2];
        mdInline(inner, { ...base, [kind === "highlight" ? "highlight" : kind]: kind === "highlight" ? "yellow" : true }).forEach((r) => runs.push(r));
      }
      rest = rest.slice(m.index + m[0].length);
    }
    return runs;
  }

  function fromMarkdown(text) {
    const lines = text.replace(/\r\n?/g, "\n").split("\n");
    const blocks = [];
    let para = null; // zbierane wiersze akapitu
    const listIndents = []; // wcięcia kolejnych poziomów listy
    const flushPara = () => {
      if (para) {
        const runs = [];
        para.lines.forEach((l, i) => { if (i) runs.push({ break: true }); mdInline(l).forEach((r) => runs.push(r)); });
        const t = tidy(runs);
        if (t.length) blocks.push({ type: para.type, runs: t });
      }
      para = null;
    };
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const fence = /^\s*(```|~~~)/.exec(line);
      if (fence) {
        flushPara();
        const runs = [];
        let j = i + 1;
        for (; j < lines.length && !lines[j].trim().startsWith(fence[1]); j++) {
          if (runs.length) runs.push({ break: true });
          if (lines[j]) runs.push({ text: lines[j], code: true });
        }
        if (runs.length) blocks.push({ type: "code", runs });
        i = j;
        continue;
      }
      if (!line.trim()) { flushPara(); listIndents.length = 0; continue; }
      let m = /^(#{1,6})\s+(.*?)\s*#*\s*$/.exec(line);
      if (m) { flushPara(); blocks.push({ type: "h", level: Math.min(3, m[1].length), runs: tidy(mdInline(m[2])) }); continue; }
      if (/^\s*(?:[-*_]\s*){3,}$/.test(line) && !para) { blocks.push({ type: "hr", runs: [] }); continue; }
      // tabela: | a | b | + wiersz |---|---|
      if (/^\s*\|.*\|\s*$/.test(line) && /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(lines[i + 1] || "")) {
        flushPara();
        const cells = (l) => l.trim().replace(/^\||\|$/g, "").split(/(?<!\\)\|/).map((c) => tidy(mdInline(c.trim())));
        const rows = [cells(line)];
        let j = i + 2;
        for (; j < lines.length && /^\s*\|.*\|\s*$/.test(lines[j]); j++) rows.push(cells(lines[j]));
        blocks.push({ type: "table", rows, head: true, runs: [] });
        i = j - 1;
        continue;
      }
      m = /^(\s*)([-*+]|\d{1,3}[.)])\s+(?:\[([ xX])\]\s+)?(.*)$/.exec(line);
      if (m) {
        flushPara();
        const indent = m[1].replace(/\t/g, "    ").length;
        while (listIndents.length && indent < listIndents[listIndents.length - 1]) listIndents.pop();
        if (!listIndents.length || indent > listIndents[listIndents.length - 1]) listIndents.push(indent);
        blocks.push({
          type: "li",
          level: Math.min(8, listIndents.length - 1),
          ordered: /\d/.test(m[2]),
          checked: m[3] == null ? undefined : m[3] !== " ",
          runs: tidy(mdInline(m[4])),
        });
        continue;
      }
      m = /^\s*>\s?(.*)$/.exec(line);
      if (m) {
        if (para && para.type !== "quote") flushPara();
        para = para || { type: "quote", lines: [] };
        para.lines.push(m[1]);
        continue;
      }
      // dalszy ciąg punktu listy (wcięty wiersz bez znacznika)
      const prev = blocks[blocks.length - 1];
      if (!para && prev?.type === "li" && /^\s{2,}\S/.test(line)) {
        prev.runs.push({ break: true }, ...tidy(mdInline(line.trim())));
        continue;
      }
      if (para && para.type !== "p") flushPara();
      para = para || { type: "p", lines: [] };
      para.lines.push(line.trim());
    }
    flushPara();
    return blocks;
  }

  // ── decyzja: co ze schowka ─────────────────────────────────────────────────
  // null = zwykły tekst (dotychczasowe wklejanie). Inaczej lista bloków.
  function parse(clipboardData) {
    const html = clipboardData?.getData("text/html") || "";
    const text = clipboardData?.getData("text/plain") || "";
    let blocks = null;
    if (html && /<[a-z]/i.test(html)) {
      blocks = fromHtml(html);
      // HTML bez żadnego formatowania (np. z edytora tekstu), a tekst wygląda na Markdown
      const rich = blocks.some((b) => b.type !== "p" || b.runs.some((r) => !r.break && (r.bold || r.italic || r.underline || r.strike || r.link || r.code || r.highlight)));
      if (!rich && text && looksLikeMarkdown(text)) blocks = fromMarkdown(text);
      else if (!rich) return null;
    } else if (text && looksLikeMarkdown(text)) {
      blocks = fromMarkdown(text);
    }
    if (!blocks || !blocks.length) return null;
    return blocks;
  }

  // ── wstawianie ─────────────────────────────────────────────────────────────
  // style fragmentu wklejki + wygląd tekstu w miejscu kursora (krój, rozmiar, kolor dokumentu)
  function runStyle(r, base) {
    const st = { ...base };
    ["bold", "italic", "underline", "strike"].forEach((k) => { if (r[k]) st[k] = true; });
    if (r.highlight) st.highlight = r.highlight;
    if (r.code) st.fontFamily = "Courier New";
    if (r.link) st.link = r.link;
    return st;
  }

  // Bloki → jeden ciąg fragmentów (przypis, komórka, wklejka bez struktury): bloki rozdziela łamanie.
  function flatten(blocks) {
    const runs = [];
    blocks.forEach((b, i) => {
      if (i) runs.push({ break: true });
      if (b.type === "table") runs.push({ text: b.rows.map((r) => r.map(plain).join("\t")).join(" / ") });
      else if (b.type === "li") runs.push({ text: b.checked === undefined ? (b.ordered ? "" : "• ") : b.checked ? "☒ " : "☐ " }, ...b.runs);
      else b.runs.forEach((r) => runs.push(r));
    });
    return runs.filter((r) => r.break || r.text);
  }

  // Wstawienie fragmentów w miejscu kursora (bez przerysowania) — jak pisanie.
  function insertInline(p, runs) {
    const sel = window.getSelection();
    if (!sel?.rangeCount) return false;
    const base = {};
    const at = typeof getInheritedRunStyleAtCaret === "function" ? getInheritedRunStyleAtCaret(p) : {};
    ["fontFamily", "fontSize", "color"].forEach((k) => { if (at[k]) base[k] = at[k]; });
    const frag = document.createDocumentFragment();
    let last = null;
    runs.forEach((r) => {
      if (r.break) { last = document.createElement("br"); frag.append(last); return; }
      const st = runStyle(r, base);
      const css = runStyleToCss(st);
      let node = css ? Object.assign(document.createElement("span"), { textContent: r.text }) : document.createTextNode(r.text);
      if (css) node.setAttribute("style", css);
      if (st.link) {
        const a = document.createElement("a");
        a.setAttribute("href", st.link);
        a.className = "dwb-link";
        a.append(node);
        node = a;
      }
      frag.append(node);
      last = node;
    });
    const range = sel.getRangeAt(0);
    range.deleteContents();
    range.insertNode(frag);
    if (last) {
      range.setStartAfter(last);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    }
    return true;
  }

  async function apply(p, blocks) {
    const inNote = !!p.dataset.noteKey;
    const single = blocks.length === 1 && blocks[0].type === "p";
    if (inNote || single) {
      asUndoStep("undoOpPaste", () => insertInline(p, single ? blocks[0].runs : flatten(blocks)));
      onInlineParagraphInput();
      return;
    }
    const index = resolveParaIndex(p);
    const sel = window.getSelection();
    if (index < 0 || !sel?.rangeCount) return;
    // komórka tabeli: tabelę z wklejki wpisujemy tekstem (tabela w tabeli — nie)
    if (p.closest("td, th")) blocks = blocks.map((b) => (b.type === "table" ? { type: "p", runs: [{ text: b.rows.map((r) => r.map(plain).join("\t")).join("\n") }] } : b));
    const range = sel.getRangeAt(0);
    if (!range.collapsed) { range.deleteContents(); onInlineParagraphInput(); }
    const pre = document.createRange();
    pre.selectNodeContents(p);
    pre.setEnd(range.startContainer, range.startOffset);
    const box = document.createElement("p");
    box.append(pre.cloneContents());
    const offset = previewRunsToPlainText(extractRunsFromPreviewParagraph(box)).length;
    const at = typeof getInheritedRunStyleAtCaret === "function" ? getInheritedRunStyleAtCaret(p) : {};
    const base = {};
    ["fontFamily", "fontSize", "color"].forEach((k) => { if (at[k]) base[k] = at[k]; });
    const edit = {
      op: "pasteBlocks", index, offset, lang: currentLang,
      blocks: blocks.map((b) => ({
        ...b,
        runs: (b.runs || []).map((r) => (r.break ? { break: true } : { text: r.text, ...runStyle(r, base) })),
        rows: b.rows ? b.rows.map((row) => row.map((cell) => cell.map((r) => (r.break ? { break: true } : { text: r.text, ...runStyle(r, base), ...(b.head && row === b.rows[0] ? { bold: true } : {}) })))) : undefined,
      })),
    };
    // kursor po wklejeniu: koniec ostatniego wklejonego bloku (composePastePlan — ten sam plan co zapis)
    const plan = composePastePlan(edit.blocks, offset === 0);
    pendingInlineCursor = { paraIndex: index + plan.lastParaOffset, offset: plan.lastTextLen };
    await applyDocumentEdit(edit);
  }

  return { parse, apply, fromHtml, fromMarkdown, looksLikeMarkdown, flatten };
})();
