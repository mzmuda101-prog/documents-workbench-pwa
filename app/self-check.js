// self-check.js — strażnik spójności dokumentu (2026-10-06).
//
// Błędy edytora rzadko widać od razu: zapis gubi poprawkę akapitu, link prowadzi do zakładki,
// której nie ma, obraz traci powiązanie, Word przy otwarciu „naprawia” plik. Strażnik sprawdza
// to, co MUSI się zgadzać — niezależnie od tego, jaką drogą użytkownik doszedł do stanu:
//
//   podgląd ↔ plik   — tyle samo akapitów; tekst każdego edytowalnego akapitu w podglądzie
//                      = tekst tego akapitu w zapisanym pliku (ten sam model co zapis);
//                      po zapisie nie zostają niezapisane poprawki
//   paczka .docx     — każda część XML się parsuje; każde r:id / r:embed ma powiązanie;
//                      każda część ma typ w [Content_Types].xml
//   zakładki         — id i nazwy bez powtórzeń, każdy początek ma koniec (i odwrotnie)
//   linki            — link „#zakładka” prowadzi do istniejącej zakładki (albo _top)
//   komentarze       — zakres i odwołanie mają komentarz w comments.xml
//   przypisy         — odnośnik ma przypis w footnotes/endnotes.xml
//   listy            — numId akapitu jest w numbering.xml
//   tekst            — <w:t> ze spacją na brzegu ma xml:space="preserve" (inaczej Word ją zjada)
//   obrazy           — wp:docPr id bez powtórzeń (Word naprawia plik przy duplikatach)
//
// Używają go: test „chaos” (scripts/chaos-playwright.js) po każdym losowym kroku użytkownika
// i aplikacja — po każdym zapisie, w tle: problem trafia do Logu (z opisem) i do dyskretnego
// komunikatu, zamiast czekać, aż ktoś przypadkiem zauważy zepsuty plik.
//
// dwbSelfCheck({ bytes?, afterSave? }) → { problems: [{ level: "error"|"warn", code, msg }], stats }

const SELF_CHECK_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const SELF_CHECK_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

async function dwbSelfCheck(opts = {}) {
  const problems = [];
  const add = (level, code, msg) => { if (problems.length < 60) problems.push({ level, code, msg }); };
  const bytes = opts.bytes || (await buildDocumentForSave());
  if (!bytes || !window.JSZip) return { problems, stats: {} };
  const zip = await JSZip.loadAsync(bytes);
  const parts = {};
  const parseErr = (d) => d.getElementsByTagName("parsererror").length > 0;
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir || !/\.(xml|rels)$/i.test(name)) continue;
    const text = await zip.file(name).async("string");
    const d = new DOMParser().parseFromString(text, "application/xml");
    if (parseErr(d)) { add("error", "XML_PARSE", `${name}: XML się nie parsuje`); continue; }
    parts[name] = d;
  }
  const doc = parts["word/document.xml"];
  if (!doc) { add("error", "NO_DOCUMENT", "brak word/document.xml"); return { problems, stats: {} }; }

  // ── [Content_Types] ──
  const ct = parts["[Content_Types].xml"];
  if (ct) {
    const defaults = new Set([...ct.getElementsByTagName("Default")].map((e) => (e.getAttribute("Extension") || "").toLowerCase()));
    const overrides = new Set([...ct.getElementsByTagName("Override")].map((e) => (e.getAttribute("PartName") || "").replace(/^\//, "")));
    Object.keys(zip.files).forEach((name) => {
      if (zip.files[name].dir || name === "[Content_Types].xml") return;
      const ext = (name.split(".").pop() || "").toLowerCase();
      if (!overrides.has(name) && !defaults.has(ext)) add("error", "CONTENT_TYPE", `${name}: brak typu w [Content_Types].xml`);
    });
  }

  // ── powiązania (r:id, r:embed …) w każdej części z własnym .rels ──
  Object.keys(parts).forEach((name) => {
    if (!/^word\/[^/]+\.xml$/.test(name)) return;
    const relsName = name.replace(/^word\//, "word/_rels/") + ".rels";
    const rels = parts[relsName];
    const ids = new Set(rels ? [...rels.getElementsByTagName("Relationship")].map((r) => r.getAttribute("Id")) : []);
    const missing = new Set();
    const walk = parts[name].getElementsByTagName("*");
    for (let i = 0; i < walk.length; i++) {
      const el = walk[i];
      for (const a of ["id", "embed", "link", "pict", "dm", "lo", "qs", "cs"]) {
        const v = el.getAttributeNS(SELF_CHECK_R, a);
        if (v && !ids.has(v)) missing.add(`${el.localName}@r:${a}=${v}`);
      }
    }
    if (missing.size) add("error", "REL_MISSING", `${name}: odwołania bez powiązania: ${[...missing].slice(0, 5).join(", ")}`);
  });

  // ── zakładki ──
  const W = SELF_CHECK_W;
  const bStarts = [...doc.getElementsByTagNameNS(W, "bookmarkStart")];
  const bEnds = [...doc.getElementsByTagNameNS(W, "bookmarkEnd")];
  const startIds = new Map();
  const names = new Map();
  bStarts.forEach((b) => {
    const id = b.getAttributeNS(W, "id");
    const name = b.getAttributeNS(W, "name");
    startIds.set(id, (startIds.get(id) || 0) + 1);
    names.set(name, (names.get(name) || 0) + 1);
  });
  const endIds = new Map();
  bEnds.forEach((e) => { const id = e.getAttributeNS(W, "id"); endIds.set(id, (endIds.get(id) || 0) + 1); });
  startIds.forEach((n, id) => {
    if (n > 1) add("error", "BOOKMARK_DUP_ID", `zakładka: id ${id} użyte ${n}×`);
    if (!endIds.has(id)) add("error", "BOOKMARK_NO_END", `zakładka id ${id}: brak końca`);
  });
  endIds.forEach((n, id) => { if (!startIds.has(id)) add("error", "BOOKMARK_NO_START", `koniec zakładki id ${id} bez początku`); });
  names.forEach((n, name) => { if (n > 1) add("error", "BOOKMARK_DUP_NAME", `zakładka „${name}” ${n}×`); });

  // ── linki do zakładek ──
  [...doc.getElementsByTagNameNS(W, "hyperlink")].forEach((h) => {
    const a = h.getAttributeNS(W, "anchor");
    if (a && a !== "_top" && !names.has(a)) add("warn", "LINK_BROKEN", `link „${(h.textContent || "").slice(0, 40)}” → #${a}: nie ma takiej zakładki`);
    if (!h.getElementsByTagNameNS(W, "r").length) add("warn", "LINK_EMPTY", "pusty link (bez tekstu)");
  });

  // ── komentarze ──
  const cDoc = parts["word/comments.xml"];
  const commentIds = new Set(cDoc ? [...cDoc.getElementsByTagNameNS(W, "comment")].map((c) => c.getAttributeNS(W, "id")) : []);
  ["commentRangeStart", "commentRangeEnd", "commentReference"].forEach((tag) => {
    [...doc.getElementsByTagNameNS(W, tag)].forEach((m) => {
      const id = m.getAttributeNS(W, "id");
      if (!commentIds.has(id)) add("error", "COMMENT_ORPHAN", `${tag} id ${id}: nie ma komentarza w comments.xml`);
    });
  });

  // ── przypisy ──
  [["footnoteReference", "word/footnotes.xml", "footnote"], ["endnoteReference", "word/endnotes.xml", "endnote"]].forEach(([tag, part, el]) => {
    const refs = [...doc.getElementsByTagNameNS(W, tag)];
    if (!refs.length) return;
    const have = new Set(parts[part] ? [...parts[part].getElementsByTagNameNS(W, el)].map((n) => n.getAttributeNS(W, "id")) : []);
    refs.forEach((r) => { const id = r.getAttributeNS(W, "id"); if (!have.has(id)) add("error", "NOTE_ORPHAN", `${tag} id ${id}: brak przypisu w ${part}`); });
  });

  // ── listy ──
  const numDoc = parts["word/numbering.xml"];
  const numIds = new Set(numDoc ? [...numDoc.getElementsByTagNameNS(W, "num")].map((n) => n.getAttributeNS(W, "numId")) : []);
  const badNum = new Set();
  [...doc.getElementsByTagNameNS(W, "numId")].forEach((n) => {
    const v = n.getAttributeNS(W, "val");
    if (v && v !== "0" && !numIds.has(v)) badNum.add(v);
  });
  if (badNum.size) add("error", "NUMID_MISSING", `listy: numId ${[...badNum].join(", ")} nie ma w numbering.xml`);

  // ── spacje na brzegach <w:t> ──
  let lostSpaces = 0;
  let sample = "";
  Object.keys(parts).forEach((name) => {
    if (!/^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/.test(name)) return;
    [...parts[name].getElementsByTagNameNS(W, "t")].forEach((t) => {
      const v = t.textContent || "";
      if (/^\s|\s$/.test(v) && t.getAttribute("xml:space") !== "preserve") { lostSpaces++; if (!sample) sample = `${name}: „${v.slice(0, 30)}”`; }
    });
  });
  if (lostSpaces) add("error", "SPACE_NOT_PRESERVED", `${lostSpaces}× tekst ze spacją na brzegu bez xml:space="preserve" (Word zje spację), np. ${sample}`);

  // ── obrazy: wp:docPr id ──
  const docPr = new Map();
  [...doc.getElementsByTagName("*")].forEach((el) => {
    if (el.localName !== "docPr") return;
    const id = el.getAttribute("id");
    docPr.set(id, (docPr.get(id) || 0) + 1);
  });
  const dupPr = [...docPr].filter(([, n]) => n > 1).map(([id]) => id);
  if (dupPr.length) add("warn", "DOCPR_DUP", `obrazy: powtórzone wp:docPr id ${dupPr.slice(0, 5).join(", ")}`);

  // ── podgląd ↔ plik ──
  const host = typeof docCanvasEl !== "undefined" ? docCanvasEl?.querySelector(".docx-preview-host") : null;
  const xmlParas = collectParagraphElements(doc.documentElement, "all");
  let compared = 0;
  if (host && !opts.bytes) {
    const previews = collectPreviewParagraphElements(host);
    if (previews.length && previews.length !== xmlParas.length) {
      add("error", "PARA_COUNT", `akapity: podgląd ${previews.length}, plik ${xmlParas.length}`);
    } else {
      const norm = (s) => String(s || "").replace(/ /g, " ");
      previews.forEach((p, i) => {
        if (!p.classList.contains("docx-editable-p") || !xmlParas[i]) return;
        compared++;
        const a = norm(previewRunsToPlainText(extractRunsFromPreviewParagraph(p)));
        const b = norm(previewRunsToPlainText(extractRunsFromParagraphXml(xmlParas[i])));
        if (a !== b) add("error", "TEXT_MISMATCH", `akapit ${i}: podgląd „${a.slice(0, 60)}” ≠ plik „${b.slice(0, 60)}”`);
      });
    }
  }
  if (opts.afterSave && typeof collectInlineParagraphEdits === "function") {
    const left = collectInlineParagraphEdits();
    if (left.length) add("error", "SAVE_LEFT_EDITS", `po zapisie zostały niezapisane poprawki akapitów: ${left.map((x) => x.index).slice(0, 8).join(", ")}`);
  }
  return { problems, stats: { paragraphs: xmlParas.length, compared, bookmarks: bStarts.length, parts: Object.keys(parts).length } };
}

// W aplikacji: po każdym zapisie (w tle, bez blokowania). Problem → Log + komunikat raz na plik.
// Bez zapisu na dysk nic się nie dzieje — strażnik nie spowalnia pisania.
// Liczą się tylko problemy NOWE względem pliku, który użytkownik otworzył (wada samego pliku,
// np. spacja bez xml:space z innego programu, to nie błąd aplikacji).
const selfCheckBase = { bytes: null, keys: null };
const selfCheckKey = (p) => `${p.code}|${p.msg}`;
async function dwbSelfCheckNew(opts = {}) {
  const res = await dwbSelfCheck(opts);
  if (selfCheckBase.bytes && !selfCheckBase.keys) {
    const base = await dwbSelfCheck({ bytes: selfCheckBase.bytes }).catch(() => ({ problems: [] }));
    selfCheckBase.keys = new Set(base.problems.map(selfCheckKey));
  }
  const known = selfCheckBase.keys || new Set();
  return { ...res, problems: res.problems.filter((p) => !known.has(selfCheckKey(p))) };
}

(() => {
  let told = null;
  const runAfterSave = () => {
    const bytes = originalFileBytes;
    const go = () => dwbSelfCheckNew({ afterSave: true }).then(({ problems }) => {
      if (bytes !== originalFileBytes) return;
      const errs = problems.filter((p) => p.level === "error");
      problems.forEach((p) => log(`Strażnik spójności [${p.code}]: ${p.msg}`, p.level === "error" ? "error" : "warning"));
      if (errs.length && told !== bytes) {
        told = bytes;
        toast(t("selfCheckFound", { n: errs.length }), "warning");
      }
    }).catch((e) => log(`Strażnik spójności: ${e.message || e}`, "warning"));
    if ("requestIdleCallback" in window) requestIdleCallback(go, { timeout: 4000 }); else setTimeout(go, 600);
  };
  document.addEventListener("dwb:saved", runAfterSave); // ui-controls.js — po udanym zapisie
  // otwarty plik = punkt odniesienia (jego własne wady się nie liczą)
  const origIngest = window.ingestFile;
  if (typeof origIngest === "function") {
    window.ingestFile = async function ingestFileSelfCheck(...args) {
      const ok = await origIngest.apply(this, args);
      if (ok) Object.assign(selfCheckBase, { bytes: originalFileBytes, keys: null });
      return ok;
    };
  }
})();
