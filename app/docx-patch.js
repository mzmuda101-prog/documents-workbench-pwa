// ZIP-patch save for DOCX — preserves original package, patches word/document.xml.

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function sanitizeXmlText(s) {
  return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "");
}

function collectParagraphElements(rootEl, scope) {
  const paragraphs = [];
  const body = rootEl.getElementsByTagNameNS(W_NS, "body")[0];
  if (!body) return paragraphs;
  const start = scope === "body" ? body : rootEl;

  // Tabela = zwykłe zejście w dół (wiersze → komórki → akapity, w kolejności dokumentu). Dawniej
  // wiersze brane były getElementsByTagNameNS („w głąb”) — tabela W tabeli (formularze, układ z
  // PDF) liczyła się dwa razy: numery akapitów pliku rozjeżdżały się z podglądem i zapis edycji
  // wpisywał tekst w cudze akapity (DC-85: „KOLEJNOŚĆ PRAC” zamiast „POTENCJALNE ZAGROŻENIA”).
  function walk(node) {
    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i];
      if (child.nodeType !== 1) continue;
      if (child.localName === "p" && child.namespaceURI === W_NS) paragraphs.push(child);
      else if (child.localName !== "sectPr" && child.localName !== "tblPr" && child.localName !== "tblGrid" && child.localName !== "trPr" && child.localName !== "tcPr") walk(child);
    }
  }

  walk(start === body ? body : start);
  return paragraphs;
}

function collectTextNodes(doc, scope) {
  const nodes = [];
  const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
  if (!body) return nodes;
  const root = scope === "body" ? body : doc.documentElement;
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_ELEMENT, null);
  let el = walker.currentNode;
  while (el) {
    if (el.localName === "t" && el.namespaceURI === W_NS) nodes.push(el);
    el = walker.nextNode();
  }
  return nodes;
}

function getParagraphText(pEl) {
  let out = "";
  const runs = pEl.getElementsByTagNameNS(W_NS, "r");
  for (let r = 0; r < runs.length; r++) {
    const children = runs[r].childNodes;
    for (let c = 0; c < children.length; c++) {
      const child = children[c];
      if (child.nodeType !== 1) continue;
      if (child.localName === "t" && child.namespaceURI === W_NS) out += child.textContent || "";
      else if (child.localName === "br" && child.namespaceURI === W_NS) out += "\n";
    }
  }
  return out;
}

function clearParagraphRuns(pEl) {
  Array.from(pEl.getElementsByTagNameNS(W_NS, "r")).forEach((r) => r.parentNode.removeChild(r));
  // puste opakowania linków (fragmenty już usunięte) — nowe linki tworzy applyRunsToParagraphXml
  Array.from(pEl.childNodes).forEach((n) => { if (n.localName === "hyperlink" && n.namespaceURI === W_NS) pEl.removeChild(n); });
}

function setParagraphText(pEl, text) {
  const sanitized = sanitizeXmlText(text);
  clearParagraphRuns(pEl);
  if (!sanitized) return;
  const parts = sanitized.split("\n");
  const doc = pEl.ownerDocument;
  parts.forEach((part, i) => {
    const r = doc.createElementNS(W_NS, "r");
    if (part) {
      const t = doc.createElementNS(W_NS, "t");
      if (/^\s|\s$/.test(part)) t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
      t.textContent = part;
      r.appendChild(t);
    }
    pEl.appendChild(r);
    if (i < parts.length - 1) {
      const brRun = doc.createElementNS(W_NS, "r");
      brRun.appendChild(doc.createElementNS(W_NS, "br"));
      pEl.appendChild(brRun);
    }
  });
}

// Kopia właściwości akapitu (w:pPr) do NOWEGO akapitu: śledzone zmiany w środku (w:pPrChange,
// wstawiony/usunięty znak akapitu w:rPr/w:ins|w:del…) dostają nowe numery. Ten sam w:id w dwóch
// akapitach to błąd schematu (walidator Open XML SDK: „should have unique value”, 2026-10-05:
// Enter w akapicie ze śledzoną zmianą formatu). Numery od największego w:id w części + 1.
function freshRevisionIds(el, doc) {
  const withId = (root) => Array.from(root.getElementsByTagNameNS(W_NS, "*")).filter((n) => n.hasAttributeNS(W_NS, "id"));
  const mine = withId(el);
  if (!mine.length) return;
  let next = 0;
  for (const n of withId(doc.documentElement)) next = Math.max(next, parseInt(n.getAttributeNS(W_NS, "id"), 10) || 0);
  mine.forEach((n) => n.setAttributeNS(W_NS, "w:id", String(++next)));
}

function splitParagraphInXml(xml, index, beforeText, afterText, beforeRuns, afterRuns, nextNormal) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[index];
  if (!p) return { xml, count: 0 };
  if (beforeRuns?.length) applyRunsToParagraphXml(p, beforeRuns);
  else setParagraphText(p, beforeText);
  // Nowy akapit = TYLKO właściwości akapitu (w:pPr) + treść „po kursorze”. Dawniej pełna kopia
  // starego bez fragmentów tekstu: zostawały w niej pusta kopia pola formularza (zdublowana
  // kontrolka, przesunięta numeracja pól → cudze pola w następnych akapitach), zakładki
  // o tej samej nazwie i zdublowane zakresy komentarzy (Word zgłasza błąd pliku).
  const newP = doc.createElementNS(W_NS, "w:p");
  const pPrSrc = Array.from(p.childNodes).find((n) => n.localName === "pPr" && n.namespaceURI === W_NS);
  if (pPrSrc) freshRevisionIds(newP.appendChild(pPrSrc.cloneNode(true)), doc);
  if (afterRuns?.length) applyRunsToParagraphXml(newP, afterRuns);
  else setParagraphText(newP, afterText);
  if (p.nextSibling) p.parentNode.insertBefore(newP, p.nextSibling);
  else p.parentNode.appendChild(newP);
  // Znacznik końca sekcji zostaje tylko na drugim (ostatnim) akapicie — dawniej klon
  // dostawał kopię i sekcja (marginesy, orientacja) kończyła się o akapit za wcześnie.
  const pPr = p.getElementsByTagNameNS(W_NS, "pPr")[0];
  const sect = pPr && Array.from(pPr.childNodes).find((n) => n.localName === "sectPr");
  if (sect) pPr.removeChild(sect);
  // Enter na końcu nagłówka/tytułu — dalej pisze się zwykłym tekstem (jak „styl następnego
  // akapitu” w Wordzie), bez stylu, podziału strony i poziomu konspektu poprzedniego.
  // Podział strony „przed” należy do pierwszej części — druga nie zaczyna kolejnej strony.
  const nPr = newP.getElementsByTagNameNS(W_NS, "pPr")[0];
  if (nPr) {
    const drop = nextNormal ? ["pStyle", "pageBreakBefore", "outlineLvl", "keepNext"] : ["pageBreakBefore"];
    Array.from(nPr.childNodes).filter((n) => drop.includes(n.localName)).forEach((n) => nPr.removeChild(n));
    if (!nPr.firstChild) newP.removeChild(nPr);
  }
  return { xml: new XMLSerializer().serializeToString(doc), count: 1 };
}

function getWVal(el) {
  if (!el) return null;
  return el.getAttributeNS(W_NS, "val") ?? el.getAttribute("w:val") ?? el.getAttribute("val");
}

function setWVal(el, value) {
  el.setAttributeNS(W_NS, "val", String(value));
}

function getParagraphListLevel(pEl) {
  const pPr = pEl.getElementsByTagNameNS(W_NS, "pPr")[0];
  const numPr = pPr?.getElementsByTagNameNS(W_NS, "numPr")[0];
  if (!numPr) return -1;
  const ilvl = numPr.getElementsByTagNameNS(W_NS, "ilvl")[0];
  const raw = ilvl ? getWVal(ilvl) : "0";
  const level = parseInt(raw, 10);
  return Number.isFinite(level) ? level : 0;
}

function mergeParagraphInXml(xml, index, mergedRuns) {
  if (index <= 0) return { xml, count: 0 };
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const prev = paragraphs[index - 1];
  const curr = paragraphs[index];
  if (!prev || !curr) return { xml, count: 0 };
  const joinAt = getParagraphText(prev).length;
  if (mergedRuns?.length) applyRunsToParagraphXml(prev, mergedRuns);
  else setParagraphText(prev, getParagraphText(prev) + getParagraphText(curr));
  curr.parentNode.removeChild(curr);
  return { xml: new XMLSerializer().serializeToString(doc), count: 1, joinAt };
}

// Usunięcie zaznaczenia przez kilka akapitów (doc-selection.js): akapity from..to (indeksy jak
// w podglądzie) i wszystko między nimi (tabele, obrazy) znikają, zostaje jeden akapit z treścią
// mergedRuns — z właściwościami pierwszego (keep „first”) albo ostatniego (keep „last”: cały
// pierwszy akapit był zaznaczony, jak usunięcie akapitu ze znacznikiem w Wordzie).
function deleteParagraphRangeInXml(xml, edit) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const first = paragraphs[edit.from];
  const last = paragraphs[edit.to];
  if (!first || !last || edit.to <= edit.from || first.parentNode !== last.parentNode) return { xml, count: 0 };
  const keep = edit.keep === "last" ? last : first;
  const drop = [];
  for (let n = edit.keep === "last" ? first : first.nextSibling; n && n !== last; n = n.nextSibling) drop.push(n);
  if (edit.keep !== "last") drop.push(last);
  drop.forEach((n) => n.parentNode.removeChild(n));
  applyRunsToParagraphXml(keep, edit.mergedRuns || []);
  return { xml: new XMLSerializer().serializeToString(doc), count: 1 };
}

function changeListLevelInXml(xml, index, delta) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  const p = paragraphs[index];
  if (!p) return { xml, count: 0 };
  let pPr = p.getElementsByTagNameNS(W_NS, "pPr")[0];
  if (!pPr) {
    pPr = doc.createElementNS(W_NS, "pPr");
    p.insertBefore(pPr, p.firstChild);
  }
  const numPr = pPr.getElementsByTagNameNS(W_NS, "numPr")[0];
  if (!numPr) return { xml, count: 0 };
  let ilvl = numPr.getElementsByTagNameNS(W_NS, "ilvl")[0];
  if (!ilvl) {
    ilvl = doc.createElementNS(W_NS, "ilvl");
    numPr.insertBefore(ilvl, numPr.firstChild);
    setWVal(ilvl, "0");
  }
  const current = parseInt(getWVal(ilvl) || "0", 10);
  const next = Math.max(0, Math.min(8, current + delta));
  if (next === current) return { xml, count: 0 };
  setWVal(ilvl, String(next));
  return { xml: new XMLSerializer().serializeToString(doc), count: 1 };
}

// ── wspólna pamięć rozpakowanego pliku (paczka F) ────────────────────────────
// Przy otwarciu ten sam .docx był rozpakowywany 3× (style nagłówków, teksty akapitów,
// formatowanie akapitów), a document.xml parsowany 2× — przy ~300 stronach to ~0,6 s
// (CPU ×4). Bajty są niezmienne (każda edycja daje NOWE), więc wynik trzymamy przy nich.
// Pamiętamy tylko OSTATNI plik (po jednym wpisie na rodzaj): historia cofania trzyma
// stare wersje bajtów, a WeakMap trzymałby przy każdej sparsowany XML (dziesiątki MB).
// Tylko do ODCZYTU: buildPatchedDocx zmienia zip, więc ładuje własny, a wynikowy
// document.xml podaje dalej (seedDocumentXml) — kolejny odczyt nie rozpakowuje.
const docxCache = { zip: [null, null], xml: [null, null], dom: [null, null] }; // [bytes, Promise]

function docxCached(kind, bytes, make) {
  const slot = docxCache[kind];
  if (slot[0] !== bytes) { slot[0] = bytes; slot[1] = make(); }
  return slot[1];
}
function loadDocxZipCached(bytes) {
  return docxCached("zip", bytes, () => window.JSZip.loadAsync(bytes));
}
function getDocumentXmlString(bytes) {
  return docxCached("xml", bytes, () => loadDocxZipCached(bytes).then((zip) => zip.file("word/document.xml")?.async("string") ?? null));
}
function getDocumentXmlDom(bytes) {
  return docxCached("dom", bytes, () => getDocumentXmlString(bytes).then((xml) => (xml ? new DOMParser().parseFromString(xml, "application/xml") : null)));
}
function seedDocumentXml(bytes, xml) {
  if (!bytes || typeof xml !== "string") return;
  docxCache.xml = [bytes, Promise.resolve(xml)];
}

async function extractParagraphTextsFromDocx(bytes) {
  if (!window.JSZip || !bytes) return [];
  const doc = await getDocumentXmlDom(bytes);
  if (!doc) return [];
  return collectParagraphElements(doc.documentElement, "all").map(getParagraphText);
}

// ── Znajdź i zamień: JEDEN silnik dla szukania, licznika i zamiany ─────────────
// Dawniej szukanie w podglądzie ignorowało wielkość liter, a zamiana i licznik ją
// rozróżniały → „Zamień bieżące” trafiało w INNE wystąpienie niż podświetlone. Zamiana szła
// po pojedynczych <w:t>, więc słowo rozcięte między fragmenty (poprawka pisowni, zmiana
// formatu w środku) było znajdowane, ale nie zamieniane. Teraz wszystko liczy po tekście
// całego akapitu (jego <w:t> po kolei) tym samym wyrażeniem.
//
// edit: { find, replace, regex, matchCase, wholeWord }. matchCase domyślnie TAK (jak dawniej —
// snippety, pola, Narzędzia edycji); panel Znajdź i zamień podaje go jawnie.
function buildFindRegex(edit) {
  if (!edit?.find) return null;
  const matchCase = edit.matchCase !== false;
  let src = edit.regex ? edit.find : edit.find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let flags = matchCase ? "g" : "gi";
  if (edit.unicode && !edit.wholeWord) flags += "u";
  if (edit.wholeWord) {
    src = `(?<![\\p{L}\\p{N}_])(?:${src})(?![\\p{L}\\p{N}_])`;
    flags += "u"; // \p{…} wymaga u; bez „całych słów” u nie włączamy (psuje część zwykłych wzorców)
  }
  try { return new RegExp(src, flags); } catch (_) { return null; }
}

// Wszystkie trafienia w tekście: [{ start, end, m }] (m = wynik exec — grupy dla $1).
function findAllMatches(text, re) {
  const out = [];
  if (!text || !re) return out;
  re.lastIndex = 0;
  let m;
  while ((m = re.exec(text)) !== null) {
    if (!m[0].length) { re.lastIndex++; continue; } // puste dopasowanie nic nie zamienia
    out.push({ start: m.index, end: m.index + m[0].length, m });
  }
  return out;
}

// Tekst zastępczy: w trybie wyrażeń $& $1…$99 $$ jak w JS; w zwykłym dosłownie.
function expandReplacement(repl, m, regex) {
  const s = String(repl ?? "");
  if (!regex) return s;
  return s.replace(/\$(\$|&|\d{1,2})/g, (all, g) => {
    if (g === "$") return "$";
    if (g === "&") return m[0];
    const v = m[Number(g)];
    return v === undefined ? all : v;
  });
}

// <w:t> NALEŻĄCE do akapitu (bez akapitów zagnieżdżonych, np. w polu tekstowym).
function paragraphTextNodes(p) {
  return Array.from(p.getElementsByTagNameNS(W_NS, "t")).filter((t) => {
    let n = t.parentNode;
    while (n && !(n.localName === "p" && n.namespaceURI === W_NS)) n = n.parentNode;
    return n === p;
  });
}

function paragraphSearchText(p) {
  return paragraphTextNodes(p).map((t) => t.textContent || "").join("");
}

// Zamienia wybrane trafienia w akapicie. pick(k) — czy zamienić k-te trafienie w akapicie.
// Tekst zastępczy trafia do fragmentu, w którym zaczyna się trafienie (bierze jego format —
// jak w Wordzie); z kolejnych fragmentów znika zamieniona część.
function replaceInParagraphXml(p, re, edit, pick) {
  const nodes = paragraphTextNodes(p);
  if (!nodes.length) return 0;
  const texts = nodes.map((t) => t.textContent || "");
  const starts = [];
  let acc = 0;
  texts.forEach((s) => { starts.push(acc); acc += s.length; });
  const matches = findAllMatches(texts.join(""), re);
  let count = 0;
  for (let k = matches.length - 1; k >= 0; k--) { // od końca: wcześniejsze przesunięcia się nie zmieniają
    if (!pick(k)) continue;
    const { start, end, m } = matches[k];
    spliceTextNodes(texts, starts, start, end, sanitizeXmlText(expandReplacement(edit.replace, m, !!edit.regex && !edit.literalReplace)));
    count++;
  }
  if (!count) return 0;
  writeTextNodes(nodes, texts);
  return count;
}

// Wstaw `repl` w miejsce [start, end) tekstu akapitu rozłożonego na fragmenty (texts/starts
// z CHWILI przed zmianami). Zmiany robić od końca akapitu. Tekst trafia do fragmentu, w którym
// zaczyna się zakres (bierze jego formatowanie).
function spliceTextNodes(texts, starts, start, end, repl) {
  let a = 0;
  while (a < texts.length - 1 && starts[a] + texts[a].length <= start && !(start === end && starts[a] + texts[a].length === start)) a++;
  if (start === end) { // samo wstawienie
    const from = Math.max(0, start - starts[a]);
    texts[a] = texts[a].slice(0, from) + repl + texts[a].slice(from);
    return;
  }
  for (let i = a; i < texts.length && starts[i] < end; i++) {
    const cur = texts[i];
    const from = Math.max(0, start - starts[i]);
    const to = Math.min(cur.length, end - starts[i]);
    texts[i] = cur.slice(0, from) + (i === a ? repl : "") + cur.slice(to);
  }
}

function writeTextNodes(nodes, texts) {
  nodes.forEach((t, i) => {
    if (t.textContent === texts[i]) return;
    t.textContent = texts[i];
    t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    if (texts[i].includes("\n")) splitNewlinesIntoBreaks(t);
  });
}

// „\n” we wstawianym tekście (wielowierszowy snippet, wartość pola z adresem) = łamanie
// wiersza. W <w:t> Word pokazałby spację — w obrębie tego samego fragmentu robimy
// <w:t>a</w:t><w:br/><w:t>b</w:t> (ten sam format dla wszystkich wierszy).
function splitNewlinesIntoBreaks(t) {
  const parts = t.textContent.split("\n");
  const doc = t.ownerDocument;
  const XML_NS = "http://www.w3.org/XML/1998/namespace";
  t.textContent = parts[0];
  let after = t;
  for (let i = 1; i < parts.length; i++) {
    const br = doc.createElementNS(W_NS, "w:br");
    after.parentNode.insertBefore(br, after.nextSibling);
    after = br;
    if (!parts[i]) continue;
    const nt = doc.createElementNS(W_NS, "w:t");
    nt.setAttributeNS(XML_NS, "xml:space", "preserve");
    nt.textContent = parts[i];
    after.parentNode.insertBefore(nt, after.nextSibling);
    after = nt;
  }
}

// Minimalne różnice a → b: [{ start, end, text }] we współrzędnych a (Myers, O((N+M)·D)).
// Duże albo bardzo różne teksty → jeden zakres (środek między wspólnym początkiem i końcem).
function diffTextSegments(a, b) {
  if (a === b) return [];
  let pre = 0;
  while (pre < a.length && pre < b.length && a[pre] === b[pre]) pre++;
  let suf = 0;
  while (suf < a.length - pre && suf < b.length - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
  const A = a.slice(pre, a.length - suf);
  const B = b.slice(pre, b.length - suf);
  const whole = [{ start: pre, end: pre + A.length, text: B }];
  const N = A.length;
  const M = B.length;
  if (!N || !M || N + M > 6000) return whole;
  const max = N + M;
  const off = max;
  let v = new Int32Array(2 * max + 2);
  const trace = [];
  let found = false;
  for (let d = 0; d <= max && d <= 300; d++) {
    trace.push(v.slice());
    for (let k = -d; k <= d; k += 2) {
      let x = (k === -d || (k !== d && v[off + k - 1] < v[off + k + 1])) ? v[off + k + 1] : v[off + k - 1] + 1;
      let y = x - k;
      while (x < N && y < M && A[x] === B[y]) { x++; y++; }
      v[off + k] = x;
      if (x >= N && y >= M) { found = true; break; }
    }
    if (found) break;
  }
  if (!found) return whole;
  const ops = []; // { at, del, ins } w A
  let x = N;
  let y = M;
  for (let d = trace.length - 1; d > 0; d--) {
    const vv = trace[d];
    const k = x - y;
    const prevK = (k === -d || (k !== d && vv[off + k - 1] < vv[off + k + 1])) ? k + 1 : k - 1;
    const prevX = vv[off + prevK];
    const prevY = prevX - prevK;
    while (x > prevX && y > prevY) { x--; y--; }
    if (x === prevX) ops.push({ at: prevX, del: 0, ins: B[prevY] });
    else ops.push({ at: prevX, del: 1, ins: "" });
    x = prevX;
    y = prevY;
  }
  ops.reverse();
  const segs = [];
  ops.forEach((o) => {
    const last = segs[segs.length - 1];
    if (last && last.end === o.at) { last.end += o.del; last.text += o.ins; }
    else segs.push({ start: o.at, end: o.at + o.del, text: o.ins });
  });
  return segs.map((sg) => ({ start: sg.start + pre, end: sg.end + pre, text: sg.text }));
}

// Nowy tekst akapitu BEZ gubienia formatowania: zmieniamy tylko różniące się znaki
// (Korekta, szybka edycja w Strukturze). Dawniej cały akapit stawał się jednym zwykłym
// fragmentem — znikały pogrubienia, kursywa, kolory. false = nie da się (np. łamanie wiersza).
function setParagraphTextPreservingRuns(p, next) {
  const nodes = paragraphTextNodes(p);
  if (!nodes.length || next.includes("\n")) return false;
  const texts = nodes.map((t) => t.textContent || "");
  const cur = texts.join("");
  if (cur !== getParagraphText(p)) return false; // łamania wiersza, tekst poza <w:t> itp.
  const starts = [];
  let acc = 0;
  texts.forEach((s) => { starts.push(acc); acc += s.length; });
  const segs = diffTextSegments(cur, next);
  for (let i = segs.length - 1; i >= 0; i--) spliceTextNodes(texts, starts, segs[i].start, segs[i].end, segs[i].text);
  writeTextNodes(nodes, texts);
  return true;
}

// Akapity w zakresie: "headings" = tylko nagłówki sekcji (style z docHeadingStyleClasses).
function paragraphsInFindScope(doc, scope) {
  const all = collectParagraphElements(doc.documentElement, scope === "body" ? "body" : "all");
  if (scope !== "headings") return all;
  const headingClasses = typeof docHeadingStyleClasses !== "undefined" ? docHeadingStyleClasses : new Map();
  return all.map((p, i) => [p, i]).filter(([p]) => {
    const pStyle = p.getElementsByTagNameNS(W_NS, "pStyle")[0];
    const id = pStyle ? getWVal(pStyle) : "";
    return id && typeof docxStyleClassName === "function" && headingClasses.has(docxStyleClassName(id));
  });
}

// Lista trafień w treści: [{ paraIndex, occurrence, start, end, text, context }] — kolejność
// dokumentu; ta sama numeracja co w zamianie (target).
// textOverride: Map(paraIndex → tekst) — akapity zmienione w podglądzie, jeszcze nie w pliku
// (tańsze niż przebudowa całego pliku przed każdym szukaniem).
function scanFindMatchesInDoc(doc, edit, scope, textOverride) {
  const re = buildFindRegex(edit);
  if (!re) return [];
  const paras = collectParagraphElements(doc.documentElement, "all");
  const inScope = scope === "headings"
    ? new Set(paragraphsInFindScope(doc, "headings").map(([, i]) => i))
    : null;
  const out = [];
  paras.forEach((p, paraIndex) => {
    if (inScope && !inScope.has(paraIndex)) return;
    const text = textOverride?.has(paraIndex) ? textOverride.get(paraIndex) : paragraphSearchText(p);
    findAllMatches(text, re).forEach(({ start, end }, occurrence) => {
      out.push({ paraIndex, occurrence, start, end, text: text.slice(start, end), context: text });
    });
  });
  return out;
}

// Dla testów/zgodności: licznik po gotowych tekstach akapitów.
function countReplacePreview(texts, edit) {
  const re = buildFindRegex(edit);
  let hits = 0;
  let paras = 0;
  const samples = [];
  (texts || []).forEach((text, index) => {
    const n = findAllMatches(text, re).length;
    if (!n) return;
    hits += n;
    paras++;
    if (samples.length < 8) samples.push({ index, snippet: text.length > 72 ? `${text.slice(0, 69)}…` : text, count: n });
  });
  return { hits, paras, samples };
}

// opts.target = { paraIndex, occurrence } — dokładnie to jedno trafienie (panel: „Zamień bieżące”).
// opts.maxReplacements / skipReplacements — dawne API (zostaje dla zgodności).
function applyReplaceInXml(xml, edit, scope, opts = {}) {
  const re = buildFindRegex(edit);
  if (!re) return { xml, count: 0 };
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const count = replaceInDocParagraphs(doc, re, edit, scope, opts);
  if (!count) return { xml, count: 0 };
  return { xml: new XMLSerializer().serializeToString(doc), count };
}

function replaceInDocParagraphs(doc, re, edit, scope, opts = {}) {
  const isBody = !!doc.getElementsByTagNameNS(W_NS, "body")[0];
  const list = isBody
    ? (scope === "headings" ? paragraphsInFindScope(doc, "headings") : collectParagraphElements(doc.documentElement, scope === "body" ? "body" : "all").map((p, i) => [p, i]))
    : Array.from(doc.getElementsByTagNameNS(W_NS, "p")).map((p, i) => [p, i]); // nagłówek/stopka/przypisy
  const target = opts.target;
  const max = opts.maxReplacements ?? Infinity;
  let skip = opts.skipReplacements ?? 0;
  let done = 0;
  let count = 0;
  for (const [p, paraIndex] of list) {
    if (done >= max) break;
    if (target && paraIndex !== target.paraIndex) continue;
    if (target) { count += replaceInParagraphXml(p, re, edit, (k) => k === target.occurrence); break; }
    if (max === Infinity && !skip) { count += replaceInParagraphXml(p, re, edit, () => true); continue; }
    // dawne API: pomiń `skip` pierwszych, zamień do `max` — liczone w kolejności dokumentu
    const n = findAllMatches(paragraphSearchText(p), re).length;
    const chosen = new Set();
    for (let k = 0; k < n && done < max; k++) {
      if (skip > 0) { skip--; continue; }
      chosen.add(k);
      done++;
    }
    if (chosen.size) count += replaceInParagraphXml(p, re, edit, (k) => chosen.has(k));
  }
  return count;
}

// Nagłówki, stopki, przypisy: części poza document.xml (lista części z docx-revisions.js).
async function replaceInOtherParts(zip, edit) {
  const re = buildFindRegex(edit);
  if (!re) return 0;
  let total = 0;
  const names = Object.keys(zip.files).filter((n) => /^word\/(header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(n));
  for (const name of names) {
    const doc = new DOMParser().parseFromString(await zip.file(name).async("string"), "application/xml");
    const n = replaceInDocParagraphs(doc, re, edit, "all");
    if (!n) continue;
    total += n;
    zip.file(name, new XMLSerializer().serializeToString(doc));
  }
  return total;
}

async function countInOtherParts(bytes, edit) {
  const re = buildFindRegex(edit);
  if (!re || !bytes) return 0;
  const zip = await loadDocxZipCached(bytes);
  let total = 0;
  for (const name of Object.keys(zip.files).filter((n) => /^word\/(header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(n))) {
    const doc = new DOMParser().parseFromString(await zip.file(name).async("string"), "application/xml");
    Array.from(doc.getElementsByTagNameNS(W_NS, "p")).forEach((p) => { total += findAllMatches(paragraphSearchText(p), re).length; });
  }
  return total;
}

function applyParagraphBatchInXml(xml, items) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const paragraphs = collectParagraphElements(doc.documentElement, "all");
  let count = 0;
  (items || []).forEach(({ index, text, runs }) => {
    const p = paragraphs[index];
    if (!p) return;
    // Pusta lista fragmentów = akapit wyczyszczony (dawniej szła ścieżka „text”, a text był
    // pusty — do pliku trafiało dosłowne „undefined”)
    if (Array.isArray(runs)) {
      const current = extractRunsFromParagraphXml(p);
      if (runsEqual(current, runs)) return;
      applyRunsToParagraphXml(p, runs);
      clearSdtPlaceholderAround(p);
      count++;
      return;
    }
    const raw = getParagraphText(p);
    const next = sanitizeXmlText(text);
    if (next === raw) return;
    if (!setParagraphTextPreservingRuns(p, next)) setParagraphText(p, next);
    clearSdtPlaceholderAround(p);
    count++;
  });
  if (!count) return { xml, count: 0 };
  return { xml: new XMLSerializer().serializeToString(doc), count };
}

function buildParagraphTransformFn(edit) {
  const locale = (typeof I18N !== "undefined" && I18N[currentLang] && I18N[currentLang].locale) || "pl-PL";
  if (edit.op === "case") {
    const m = edit.mode;
    if (m === "upper") return (s) => s.toLocaleUpperCase(locale);
    if (m === "lower") return (s) => s.toLocaleLowerCase(locale);
    return (s) => s.replace(/\p{L}[\p{L}\p{M}]*/gu, (w) => w[0].toLocaleUpperCase(locale) + w.slice(1).toLocaleLowerCase(locale));
  }
  if (edit.op === "trim") {
    const m = edit.mode;
    if (m === "collapse") return (s) => s.replace(/\s+/gu, " ").trim();
    if (m === "hard") return (s) => s.replace(/[\u00A0\u2007\u202F]/g, " ").trim();
    return (s) => s.trim();
  }
  if (edit.op === "affix") {
    const pre = edit.prefix || "";
    const suf = edit.suffix || "";
    return (s) => pre + s + suf;
  }
  return (s) => s;
}

// setParagraphText przepisuje cały akapit jednym fragmentem — pole formularza, pole Worda, link,
// przypis, obraz czy śledzona zmiana w środku by przepadły. Takie akapity przekształcenia pomijają
// (te same, które w podglądzie są tylko do odczytu).
const TRANSFORM_SKIP_TAGS = ["sdt", "fldChar", "fldSimple", "hyperlink", "footnoteReference", "endnoteReference", "drawing", "pict", "object", "ins", "del", "moveFrom", "moveTo"];
function paragraphHasProtectedContent(p) {
  return TRANSFORM_SKIP_TAGS.some((tag) => p.getElementsByTagNameNS(W_NS, tag).length);
}

// Ten sam tekst co do długości (wielkość liter): znak po znaku do istniejących <w:t> — formatowanie,
// znaczniki komentarzy i zakładki zostają dokładnie na miejscu (też przy łamaniu wiersza).
function setParagraphTextSameLength(p, raw, next) {
  if (next.length !== raw.length) return false;
  let pos = 0;
  const runs = p.getElementsByTagNameNS(W_NS, "r");
  for (let r = 0; r < runs.length; r++) {
    Array.from(runs[r].childNodes).forEach((c) => {
      if (c.nodeType !== 1 || c.namespaceURI !== W_NS) return;
      if (c.localName === "t") {
        const len = (c.textContent || "").length;
        c.textContent = next.slice(pos, pos + len);
        if (/^\s|\s$/.test(c.textContent)) c.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
        pos += len;
      } else if (c.localName === "br") pos += 1;
    });
  }
  return pos === next.length;
}
// Akapit bez niczego poza zwykłym tekstem — tylko taki wolno przepisać jednym fragmentem.
function paragraphIsPlainText(p) {
  return Array.from(p.childNodes).every((n) => n.nodeType !== 1 || (n.namespaceURI === W_NS && (n.localName === "pPr" || n.localName === "proofErr"
    || (n.localName === "r" && Array.from(n.childNodes).every((c) => c.nodeType !== 1 || ["t", "br"].includes(c.localName))))));
}

function applyParagraphTransformInXml(xml, edit, scope) {
  const fn = buildParagraphTransformFn(edit);
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const paragraphs = collectParagraphElements(doc.documentElement, scope);
  let count = 0;
  paragraphs.forEach((p) => {
    const raw = getParagraphText(p);
    if (!raw || paragraphHasProtectedContent(p)) return;
    let next;
    try { next = fn(raw); } catch { return; }
    if (typeof next !== "string" || next === raw) return;
    // Dawniej zawsze setParagraphText — jeden goły fragment: „WIELKIE LITERY” / „Przytnij” gubiły
    // pogrubienie, kursywę, kolor, zakładki i odwołanie komentarza (komentarz zostawał bez kotwicy).
    // Teraz tekst wraca do istniejących fragmentów; gdy się nie da — akapit bez formatowania albo pominięty.
    if (!setParagraphTextSameLength(p, raw, next) && !setParagraphTextPreservingRuns(p, next)) {
      if (!paragraphIsPlainText(p)) return;
      setParagraphText(p, next);
    }
    count++;
  });
  if (!count) return { xml, count: 0 };
  return { xml: new XMLSerializer().serializeToString(doc), count };
}

function applyPlaceholderFillInXml(xml, values, scope) {
  let current = xml;
  let count = 0;
  Object.entries(values || {}).forEach(([name, val]) => {
    if (val == null || val === "") return;
    const safeName = String(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const res = applyReplaceInXml(current, {
      op: "replace",
      find: `\\{\\{\\s*${safeName}\\s*\\}\\}`,
      replace: sanitizeXmlText(String(val)),
      regex: true,
      literalReplace: true, // wartość/treść wstawiana dosłownie — „$100” to nie odwołanie do grupy
      scope: scope || "all",
    });
    current = res.xml;
    count += res.count;
  });
  if (!count) return { xml, count: 0 };
  return { xml: current, count };
}

function applySnippetExpandInXml(xml, snippetMap, scope) {
  let current = xml;
  let count = 0;
  const resolved = resolveSnippetMapBodies(snippetMap);
  const names = Object.keys(resolved).sort((a, b) => b.length - a.length);
  names.forEach((name) => {
    const safeName = String(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const res = applyReplaceInXml(current, {
      op: "replace",
      // jak w podglądzie: tylko na początku słowa i całe słowo (polskie litery)
      find: `(?<![\\p{L}\\p{N}_!])!${safeName}(?![\\p{L}\\p{N}_-])`,
      unicode: true,
      replace: sanitizeXmlText(String(resolved[name])),
      regex: true,
      literalReplace: true, // wartość/treść wstawiana dosłownie — „$100” to nie odwołanie do grupy
      scope: scope || "all",
    });
    current = res.xml;
    count += res.count;
  });
  if (!count) return { xml, count: 0 };
  return { xml: current, count };
}

function applyEditToXml(xml, edit, opts = {}) {
  const scope = edit.scope || "all";
  if (edit.op === "paragraphBatch") return applyParagraphBatchInXml(xml, edit.items);
  if (edit.op === "placeholderFill") return applyPlaceholderFillInXml(xml, edit.values, scope);
  if (edit.op === "snippetExpand") return applySnippetExpandInXml(xml, edit.snippets, scope);
  if (edit.op === "splitParagraph") return splitParagraphInXml(xml, edit.index, edit.before, edit.after, edit.beforeRuns, edit.afterRuns, edit.nextNormal);
  if (edit.op === "pageBreak") return applyPageBreakInXml(xml, edit); // docx-compose.js
  if (edit.op === "hrule") return applyHruleInXml(xml, edit);
  if (edit.op === "pageVAlign") return applyPageVAlignInXml(xml, edit); // docx-compose.js
  if (edit.op === "pageSetup") return applyPageSetupInXml(xml, edit); // docx-compose.js
  if (edit.op === "link") return applyLinkInXml(xml, edit);
  if (edit.op === "runStyle") return applyRunStyleInXml(xml, edit);
  if (edit.op === "table") return applyTableInXml(xml, edit);
  if (edit.op === "image") return applyImageInXml(xml, edit);
  if (edit.op === "mergeParagraph") return mergeParagraphInXml(xml, edit.index, edit.mergedRuns);
  if (edit.op === "deleteRange") return deleteParagraphRangeInXml(xml, edit); // doc-selection.js
  if (edit.op === "listLevel") return changeListLevelInXml(xml, edit.index, edit.delta);
  if (edit.op === "case" || edit.op === "trim" || edit.op === "affix") return applyParagraphTransformInXml(xml, edit, scope);
  return applyReplaceInXml(xml, edit, scope, opts);
}

function recordPendingEdit(edit) {
  pendingDocEdits.push({ ...edit, ts: Date.now() });
}

function commentRefIds(xml) {
  return new Set([...String(xml).matchAll(/<w:commentReference\b[^>]*\bw:id="(-?\d+)"/g)].map((m) => m[1]));
}

// Naprawa nazw krojów zepsutych przez starsze wersje (do 2026-10-05): podgląd ma
// „"DM Sans", sans-serif”, odczyt obcinał tylko początkowy cudzysłów i do pliku szło
// w:ascii="DM Sans&quot;" — Word nie znajdował takiego kroju. Przy każdym zapisie: bez cudzysłowów.
function repairFontNames(xml) {
  if (!xml.includes("&quot;") && !xml.includes("&apos;")) return xml;
  return xml.replace(/(<w:rFonts\b[^>]*>)/g, (tag) => tag.replace(/(w:(?:ascii|hAnsi|cs|eastAsia)=")([^"]*)"/g, (m, a, v) => `${a}${v.replace(/&quot;|&apos;/g, "").trim()}"`));
}

// Naprawy przy OTWARCIU pliku (w pamięci, przed podglądem; do pliku trafiają przy zapisie):
//  • BOM (\uFEFF) na początku części XML — tak zapisują niektóre programy (.NET). Parser XML
//    w Safari odrzuca taki tekst: podgląd działał (ma własny parser), ale ZAPIS w Safari po cichu
//    gubił wpisany tekst. Word czyta XML bez BOM tak samo.
//  • zepsute nazwy krojów ze starszych wersji („DM Sans&quot;”) — podgląd pokazuje właściwy krój.
// Zwraca te same bajty, gdy nie ma czego naprawiać.
// Nowe dokumenty z aplikacji (do 2026-10-05) miały w settings.xml w:themeFontLang PRZED w:compat —
// schemat Office wymaga odwrotnie (walidator Open XML SDK: „unexpected child element compat”).
// Przenosimy go za compat (i za docVars / rsids / mathPr / attachedSchema, które też go poprzedzają).
function repairSettingsOrder(xml) {
  const lang = /<w:themeFontLang\b[^>]*\/>/.exec(xml);
  const compat = xml.indexOf("<w:compat");
  if (!lang || compat < 0 || lang.index > compat) return xml;
  const rest = xml.slice(0, lang.index) + xml.slice(lang.index + lang[0].length);
  const ends = [/<\/w:compat>/g, /<w:compat\s*\/>/g, /<\/w:docVars>/g, /<\/w:rsids>/g, /<\/m:mathPr>/g, /<w:attachedSchema\b[^>]*\/>/g].map((re) => {
    let at = -1;
    for (const m of rest.matchAll(re)) at = m.index + m[0].length;
    return at;
  });
  const at = Math.max(...ends);
  return at < 0 ? xml : rest.slice(0, at) + lang[0] + rest.slice(at);
}

async function repairDocxFontNames(bytes) {
  if (!window.JSZip || !bytes) return bytes;
  try {
    const zip = await window.JSZip.loadAsync(bytes);
    let changed = false;
    for (const f of Object.keys(zip.files)) {
      if (zip.files[f].dir || !/\.(xml|rels)$/i.test(f)) continue;
      const part = await zip.file(f).async("string");
      let fixed = part.charCodeAt(0) === 0xfeff ? part.slice(1) : part;
      if (/^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments|styles)\.xml$/.test(f)) fixed = repairFontNames(fixed);
      if (f === "word/settings.xml") fixed = repairSettingsOrder(fixed);
      if (fixed !== part) { zip.file(f, fixed); changed = true; }
    }
    return changed ? await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } }) : bytes;
  } catch (_) {
    return bytes;
  }
}

async function buildPatchedDocx(bytes, edits, lastEditOpts = {}) {
  if (!window.JSZip) throw new Error("JSZip missing");
  const zip = await window.JSZip.loadAsync(bytes);
  const docFile = zip.file("word/document.xml");
  if (!docFile) throw new Error("word/document.xml missing");
  let xml = await docFile.async("string");
  const commentRefsBefore = commentRefIds(xml);
  let total = 0;
  let coreXml = null;
  const list = edits || [];
  for (let i = 0; i < list.length; i++) {
    let normalized = list[i].op ? list[i] : { ...list[i], op: "replace" };
    // tekst przypisów (doc-notes.js) — do footnotes.xml / endnotes.xml; akapity treści dalej niżej
    if (normalized.op === "paragraphBatch" && normalized.items?.some((it) => it.note)) {
      if (typeof applyNoteEditsInZip === "function") total += await applyNoteEditsInZip(zip, normalized.items.filter((it) => it.note));
      normalized = { ...normalized, items: normalized.items.filter((it) => !it.note) };
    }
    if (normalized.op === "coreMetadata") {
      if (typeof prepareCoreMetadataInZip === "function" && coreXml === null) {
        coreXml = await prepareCoreMetadataInZip(zip);
      } else if (coreXml === null) {
        const coreFile = zip.file("docProps/core.xml");
        coreXml = coreFile ? await coreFile.async("string") : (typeof createDefaultCoreXml === "function" ? createDefaultCoreXml() : "");
      }
      if (typeof applyCoreMetadataInXml === "function" && coreXml) {
        const res = applyCoreMetadataInXml(coreXml, normalized.fields || {});
        coreXml = res.xml;
        total += res.count;
      }
      continue;
    }
    if (normalized.op === "formFill") { // docx-forms.js — pola powiązane zmieniają też customXml / docProps
      const res = await applyFormFillInZip(zip, xml, normalized);
      xml = res.xml;
      total += res.count;
      continue;
    }
    if (normalized.op === "headerFooter") { // docx-compose.js — części nagłówków/stopek
      const res = await applyHeaderFooterInZip(zip, xml, normalized);
      xml = res.xml;
      total += res.count;
      continue;
    }
    if (normalized.op === "commentAdd" || normalized.op === "commentReply" || normalized.op === "commentDone" || normalized.op === "commentEdit") { // docx-compose.js
      const res = normalized.op === "commentAdd" ? await applyCommentAddInZip(zip, xml, normalized)
        : normalized.op === "commentEdit" ? await applyCommentEditInZip(zip, xml, normalized)
        : await applyCommentThreadInZip(zip, xml, normalized);
      xml = res.xml;
      total += res.count;
      continue;
    }
    if (normalized.op === "tableInsert" || normalized.op === "imageInsert") { // docx-compose.js — style / pliki w paczce
      const res = normalized.op === "tableInsert" ? await applyTableInsertInZip(zip, xml, normalized) : await applyImageInsertInZip(zip, xml, normalized);
      xml = res.xml;
      total += res.count;
      continue;
    }
    if (normalized.op === "pasteBlocks") { // docx-compose.js — wklejka ze strukturą (paste-rich.js)
      const res = await applyPasteBlocksInZip(zip, xml, normalized);
      xml = res.xml;
      total += res.count;
      continue;
    }
    if (normalized.op === "formInsert" || normalized.op === "snippetInsert") { // docx-compose.js — styl „Tekst zastępczy”
      const res = normalized.op === "formInsert" ? await applyFormInsertInZip(zip, xml, normalized) : await applySnippetInsertInZip(zip, xml, normalized);
      xml = res.xml;
      total += res.count;
      continue;
    }
    if (normalized.op === "toc") { // docx-compose.js — style spisu treści + pole TOC
      const res = await applyTocInZip(zip, xml, normalized);
      xml = res.xml;
      total += res.count;
      continue;
    }
    if (normalized.op === "list") { // docx-compose.js — dopisuje definicje do numbering.xml
      const res = await applyListInZip(zip, xml, normalized);
      xml = res.xml;
      total += res.count;
      continue;
    }
    if (normalized.op === "paraFormat") { // docx-compose.js — styl może wymagać dopisania do styles.xml
      const res = await applyParaFormatInZip(zip, xml, normalized);
      xml = res.xml;
      total += res.count;
      continue;
    }
    if (normalized.op === "noteInsert") { // doc-notes.js — nowy przypis: część footnotes/endnotes, style, odnośnik
      const res = await applyNoteInsertInZip(zip, xml, normalized);
      xml = res.xml;
      total += res.count;
      continue;
    }
    if (normalized.op === "revisions") { // docx-revisions.js — dotyka też nagłówków, stopek, przypisów, komentarzy
      const res = await applyRevisionsInZip(zip, xml, normalized);
      xml = res.xml;
      total += res.count;
      continue;
    }
    const opts = i === list.length - 1 ? lastEditOpts : {};
    const res = applyEditToXml(xml, normalized, opts);
    xml = res.xml;
    total += res.count;
    // „Zamień wszystkie” z zaznaczonym „też w nagłówkach, stopkach i przypisach”
    if (normalized.op === "replace" && normalized.otherParts && !opts.target) total += await replaceInOtherParts(zip, normalized);
  }
  if (typeof finalizeComposeParts === "function") xml = await finalizeComposeParts(zip, xml); // nowe linki: powiązania + styl
  // przypis bez odnośnika w treści (odnośnik skasowany) znika z pliku — jak w Wordzie
  if (list.length && typeof pruneOrphanNotesInZip === "function") await pruneOrphanNotesInZip(zip, xml);
  // komentarz, któremu ta zmiana skasowała odwołanie w treści (np. zaznacz wszystko + Delete) — też
  // znika, razem z resztką zakresu i wpisami w commentsExtended (dawniej zostawał bez kotwicy)
  if (list.length && commentRefsBefore.size && typeof applyRevisionsInZip === "function") {
    const after = commentRefIds(xml);
    const lost = [...commentRefsBefore].filter((id) => !after.has(id));
    if (lost.length) xml = (await applyRevisionsInZip(zip, xml, { op: "revisions", action: "removeComments", ids: lost })).xml;
  }
  xml = repairFontNames(xml);
  zip.file("word/document.xml", xml);
  if (coreXml !== null) zip.file("docProps/core.xml", coreXml);
  for (const f of Object.keys(zip.files)) {
    if (!/^word\/(header\d*|footer\d*|footnotes|endnotes|comments|styles)\.xml$/.test(f)) continue;
    const part = await zip.file(f).async("string");
    const fixed = repairFontNames(part);
    if (fixed !== part) zip.file(f, fixed);
  }
  const out = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  seedDocumentXml(out, xml); // następny odczyt tych bajtów nie musi ich rozpakowywać
  return { bytes: out, changeCount: total };
}
