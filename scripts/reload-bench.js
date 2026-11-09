// reload-bench.js — co kosztuje PRZERYSOWANIE dokumentu przy typowych operacjach (etap
// „przeładowania”, 2026-11-10). Dla każdej operacji: czas do gotowości, czy był pełny render
// (renderCurrentDocument), przeskok tekstu na ekranie (ten sam akapit przed i po), czy kursor
// został w tym samym miejscu.
//
//   node scripts/reload-bench.js                 (przewodnik.docx + ~100 stron, CPU ×4)
//   CPU=1 PARAS=4500 node scripts/reload-bench.js
//   MOBILE=1 node scripts/reload-bench.js        (telefon: Widok mobilny)
//
// Wymaga serwera testowego (python3 scripts/test-server.py 7823).
const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const CPU = Number(process.env.CPU || 4);
const PARAS = Number(process.env.PARAS || 1500);
const ROOT = path.resolve(__dirname, "..");

// duży dokument jak w perf-bench (nagłówki, listy, tabele)
async function bigDocx(n) {
  const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const LOREM = "Najemca zobowiązuje się używać lokalu zgodnie z jego przeznaczeniem i dbać o jego stan techniczny, a wszelkie zmiany wymagają pisemnej zgody Wynajmującego.";
  const para = (t, style) => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r></w:p>`;
  const body = [];
  for (let i = 0; i < n; i++) {
    if (i % 25 === 0) body.push(para(`§${i / 25 + 1} Rozdział`, "Heading1"));
    body.push(para(`${i + 1}. ${LOREM}`));
  }
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(ROOT, "docs/samples/headings-sample.docx")));
  const xml = await zip.file("word/document.xml").async("string");
  const head = xml.slice(0, xml.indexOf("<w:body>") + 8);
  const sect = xml.match(/<w:sectPr[\s\S]*<\/w:sectPr>/)[0];
  zip.file("word/document.xml", `${head}${body.join("")}${sect}</w:body></w:document>`);
  const file = path.join(os.tmpdir(), `dwb-reload-bench-${n}.docx`);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return file;
}

// operacje: przygotowanie (zaznaczenie / kursor) i wykonanie — tak jak robi to interfejs
const OPS = [
  { name: "wyrównanie do środka", run: () => composeUi.applyAlign("center") },
  { name: "styl Nagłówek 2", run: () => composeUi.applyStyle("h2") },
  { name: "lista punktowana", run: () => composeUi.applyList("bullet") },
  { name: "wcięcie +", run: (i) => { pendingInlineCursor = { paraIndex: i, offset: 10 }; return applyDocumentEdit({ op: "paraFormat", indices: [i], indentDelta: 1 }); } },
  { name: "linia pozioma", run: () => composeUi.insertHrule() },
  { name: "pogrubienie zaznaczenia", select: true, run: () => execInlineFormat("bold") },
  { name: "komentarz do słowa", run: (i) => { pendingInlineCursor = { paraIndex: i, offset: 4 }; return applyDocumentEdit({ op: "commentAdd", index: i, start: 0, end: 4, rich: [[{ text: "Uwaga" }]], author: "Test", initials: "T" }); } },
  // metadane: panel woła applyDocumentEdit bez miejsca kursora (metadata-panel.js) — jak w aplikacji
  { name: "metadane: tytuł", run: () => applyDocumentEdit({ op: "coreMetadata", fields: { title: `Bench ${Date.now()}` } }) },
  { name: "pisanie → Cofnij", typeUndo: true },
];

async function measureDoc(browser, file, label) {
  // MOBILE=1 — telefon (Widok mobilny)
  const context = await browser.newContext(process.env.MOBILE
    ? { serviceWorkers: "block", viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: true }
    : { serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  const cdp = await context.newCDPSession(page);
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles(file);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 300000 });
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await page.waitForTimeout(1500);
  await page.evaluate(() => {
    window.__renders = 0;
    const orig = window.renderCurrentDocument;
    window.renderCurrentDocument = function counted(...a) { window.__renders++; return orig.apply(this, a); };
  });
  await cdp.send("Emulation.setCPUThrottlingRate", { rate: CPU });

  const rows = [];
  for (const op of OPS) {
    // akapit treści w środku dokumentu, długi (bez blokady), na środku ekranu
    const target = await page.evaluate(() => {
      const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
      const ok = ps.filter((p) => !p.dataset.lock && p.textContent.length > 40 && !p.closest("td") && !isListParagraph(p) && !/heading|nagwek/i.test(p.className));
      const p = ok[Math.floor(ok.length / 2)];
      p.scrollIntoView({ block: "center" });
      return ps.indexOf(p);
    });
    await page.waitForTimeout(300);
    const before = await page.evaluate(([i, sel]) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
      p.closest(".docx-edit-root")?.focus();
      if (sel) { const r = document.createRange(); const tn = p.querySelector("span")?.firstChild || p.firstChild; r.setStart(tn, 3); r.setEnd(tn, 12); getSelection().removeAllRanges(); getSelection().addRange(r); }
      else placeCaret(p, 10);
      return { text: p.textContent.slice(0, 30) };
    }, [target, !!op.select]);
    // fokus na telefonie przewija widok chwilę później — pozycja „przed” po ustabilizowaniu
    await page.waitForTimeout(500);
    before.top = await page.evaluate((i) => { window.__renders = 0; return collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].getBoundingClientRect().top; }, target);
    let t0;
    if (op.typeUndo) {
      await page.keyboard.type("xyz");
      await page.locator(".docx-preview-host p").nth(0).evaluate(() => {}); // nic — krok zamyka pauza niżej
      await page.waitForTimeout(1700);
      await page.evaluate(() => { window.__renders = 0; });
      t0 = Date.now();
      await page.click("#undoBtn");
    } else {
      t0 = Date.now();
      await page.evaluate(([fn, i]) => (0, eval)(`(${fn})`)(i), [op.run.toString(), target]);
    }
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 120000 });
    await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
    const ms = Date.now() - t0;
    const where = (late) => page.evaluate(([i, txt]) => {
      const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
      const p = ps.find((x, k) => k >= i - 2 && x.textContent.slice(0, 30).includes(txt.slice(0, 20))) || ps[i];
      const sel = getSelection();
      const at = sel?.rangeCount ? ps.findIndex((x) => x.contains(sel.anchorNode)) : -1;
      return { top: p.getBoundingClientRect().top, renders: window.__renders, caretPara: at, caretOff: at >= 0 ? getCaretOffset(ps[at]) : null, selText: sel?.toString().length || 0, pi: ps.indexOf(p) };
    }, [target, before.text]);
    const after = await where();
    await page.waitForTimeout(1500);
    const late = await where();
    rows.push({ op: op.name, ms, renders: after.renders, jump: Math.round(after.top - before.top), lateJump: Math.round(late.top - before.top), caret: after.caretPara === after.pi ? (op.select ? (after.selText ? "zaznaczenie zostało" : "zaznaczenie zgubione") : `ok (${after.caretOff})`) : `inny akapit (${after.caretPara} zamiast ${after.pi})` });
    await cdp.send("Emulation.setCPUThrottlingRate", { rate: 1 });
    await page.evaluate(() => { if (typeof setDirtyState === "function") setDirtyState(false); });
    await cdp.send("Emulation.setCPUThrottlingRate", { rate: CPU });
  }
  await context.close();
  console.log(`\n■ ${label} (CPU ×${CPU})`);
  console.log("  operacja                   czas     render  przeskok (od razu / po 1,5 s)  kursor");
  rows.forEach((r) => console.log(`  ${r.op.padEnd(26)} ${String(r.ms).padStart(6)} ms  ${r.renders ? "PEŁNY " : "—     "}  ${String(r.jump).padStart(6)} / ${String(r.lateJump).padStart(6)} px      ${r.caret}`));
  if (errors.length) console.log(`  błędy strony: ${errors.slice(0, 2).join(" | ")}`);
  return rows;
}

(async () => {
  const browser = await pw.chromium.launch({ headless: true });
  const out = {};
  out.przewodnik = await measureDoc(browser, path.join(ROOT, "docs/samples/przewodnik.docx"), "przewodnik.docx (8 stron)");
  out.duzy = await measureDoc(browser, await bigDocx(PARAS), `wygenerowany, ${PARAS} akapitów (~${Math.round(PARAS / 15)} stron)`);
  await browser.close();
  fs.mkdirSync(path.join(ROOT, "output"), { recursive: true });
  fs.writeFileSync(path.join(ROOT, "output/reload-bench.json"), JSON.stringify(out, null, 1));
})();
