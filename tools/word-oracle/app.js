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

module.exports = { openApp, loadBytes, listLabels };
