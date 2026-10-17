// app.js — to samo pytanie zadane aplikacji (Playwright): numery list i tekst akapitów.
const path = require("path");
const fs = require("fs");
const { createTestPage, bootApp, waitDocReady } = require("../../scripts/docx-test-helpers");

async function openApp(engine) {
  const t = await createTestPage();
  await bootApp(t.page);
  return t;
}

async function loadBytes(page, bytes, name = "oracle.docx") {
  const file = path.join(__dirname, "../../output/word-oracle", name);
  fs.writeFileSync(file, bytes);
  await page.locator("#fileInput").setInputFiles(file);
  await waitDocReady(page);
}

// Numery list rysuje CSS (::before z licznikami) — tekst z migawki układu Chrome (DOMSnapshot),
// gdzie liczniki są już rozwiązane. Akapity treści (bez nagłówka/stopki/przypisów), po kolei.
async function listLabels(page, selector = ".docx-preview-host article p") {
  await page.evaluate((sel) => document.querySelectorAll(sel).forEach((p, i) => {
    if (!p.closest(".dwb-notes, header, footer") && !p.parentElement.closest("p")) p.setAttribute("data-oracle", i);
  }), selector);
  const s = await page.context().newCDPSession(page);
  const snap = await s.send("DOMSnapshot.captureSnapshot", { computedStyles: [] });
  await s.detach();
  const d = snap.documents[0];
  const str = snap.strings;
  const N = d.nodes;
  const L = d.layout;
  const pseudo = new Set(N.pseudoType?.index || []);
  const oracleId = new Map();
  N.attributes.forEach((attrs, n) => { for (let a = 0; a < attrs.length; a += 2) if (str[attrs[a]] === "data-oracle") oracleId.set(n, Number(str[attrs[a + 1]])); });
  const out = new Map();
  L.nodeIndex.forEach((ni, k) => {
    const t = L.text[k];
    if (t < 0) return;
    let n = ni;
    let inPseudo = false;
    while (n >= 0 && !oracleId.has(n)) { if (pseudo.has(n)) inPseudo = true; n = N.parentIndex[n]; }
    if (n < 0) return;
    const id = oracleId.get(n);
    const e = out.get(id) || { label: "", text: "", before: true };
    if (inPseudo && e.before) e.label += str[t];
    else { e.before = false; if (inPseudo) e.after = (e.after || "") + str[t]; }
    out.set(id, e);
  });
  const texts = await page.evaluate(() => Array.from(document.querySelectorAll("[data-oracle]")).map((p) => [Number(p.dataset.oracle), p.textContent]));
  return texts.map(([id, text]) => ({ label: (out.get(id)?.label || "").replace(/[\s\u00a0\u2003]+$/, ""), text }));
}

// numery przypisów w podglądzie: odnośnik „R01:<numer>”, przypis „<numer> N01” (jak noteLabels w word.js)
async function noteLabels(page) {
  return page.evaluate(() => {
    const host = document.querySelector(".docx-preview-host");
    const out = { refs: {}, notes: {} };
    const clean = (s) => s.replace(/[\uFEFF\u200B]/g, "").replace(/[\s\u00a0]+/g, " ").trim();
    host.querySelectorAll("p").forEach((p) => {
      const t = clean(p.textContent);
      const ol = p.closest("ol.dwb-notes");
      let m;
      if (!ol && (m = /^([RE]\d\d):(.*)$/.exec(t))) out.refs[m[1]] = m[2].trim();
      if (ol && (m = /^(.+?)\s*([NM]\d\d)$/.exec(t))) out.notes[m[2]] = m[1];
      else if (ol && (m = /^([NM]\d\d)$/.exec(t))) out.notes[m[1]] = `(lista:${p.closest("li")?.value})`;
    });
    return out;
  });
}

// cytaty/ramki: kursor w akapicie → co pokazuje lista stylów; potem styl z listy albo gumka; po
// wszystkim akapity zapisanego pliku (normalizeParas z cases-quote.js)
async function savedParts(page) {
  return page.evaluate(async () => {
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    return { doc: await z.file("word/document.xml").async("string"), styles: await z.file("word/styles.xml").async("string") };
  });
}
async function quoteOps(page, ops) {
  const { normalizeParas } = require("./cases-quote");
  await page.evaluate(() => { if (readOnlyMode) appFrame.setReadOnly(false); });
  const idle = () => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(350));
  await idle();
  const labels = [];
  for (const o of ops) {
    await page.evaluate((i) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
      p.focus();
      placeCaret(p, 3);
      document.dispatchEvent(new Event("selectionchange"));
    }, o.para);
    await page.waitForTimeout(250);
    labels.push(await page.evaluate(() => document.getElementById("fmtParaStyle").selectedOptions[0]?.textContent));
    if (o.op === "style") await page.selectOption("#fmtParaStyle", o.style);
    else await page.evaluate(() => clearTextFormatting());
    await idle();
  }
  const x = await savedParts(page);
  return { labels, paras: await page.evaluate(normalizeParas, x) };
}

module.exports = { openApp, loadBytes, listLabels, noteLabels, quoteOps, savedParts };
