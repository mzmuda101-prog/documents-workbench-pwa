// perf-bench.js — pomiar wydajności na dużych dokumentach (paczka F).
//
//   npm run bench                    (≈100 i ≈300 stron, CPU ×4 — przybliżenie telefonu)
//   PARAS=6000 CPU=6 npm run bench
//
// Metodyka jak w Sheet Workbench: Chromium + CDP Emulation.setCPUThrottlingRate (jedyny
// sensowny proxy telefonu — WebKita z Playwrighta nie da się dławić), czasy mierzone
// PO namalowaniu (podwójny rAF), mediana z kilku powtórzeń tam, gdzie się da.
// Rozbicie otwarcia: funkcje aplikacji są owijane z zewnątrz (kod apki nietknięty).

const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const { spawn } = require("child_process");
const pw = require("playwright");
const JSZip = require("jszip");

const ROOT = path.resolve(__dirname, "..");
const PORT = 7823;
const APP_URL = `http://127.0.0.1:${PORT}/`;
const CPU = Number(process.env.CPU || 4);
const SIZES = (process.env.PARAS ? [Number(process.env.PARAS)] : [1500, 4500]);

// ── generator dużego .docx ───────────────────────────────────────────────────
const LOREM = "Najemca zobowiązuje się używać lokalu zgodnie z jego przeznaczeniem i dbać o jego stan techniczny, a wszelkie zmiany wymagają pisemnej zgody Wynajmującego. Strony ustalają, że korespondencja będzie prowadzona drogą elektroniczną na adresy wskazane w umowie.";
const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
function para(text, style, opts = {}) {
  const pPr = style || opts.num ? `<w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ""}${opts.num ? `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>` : ""}</w:pPr>` : "";
  const runs = opts.bold
    ? `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${esc(text.slice(0, 40))}</w:t></w:r><w:r><w:t xml:space="preserve">${esc(text.slice(40))}</w:t></w:r>`
    : `<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
  return `<w:p>${pPr}${runs}</w:p>`;
}
function table(i) {
  const row = (cells) => `<w:tr>${cells.map((c) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${para(c)}</w:tc>`).join("")}</w:tr>`;
  return `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>${row([`Poz. ${i}`, "Opis", "Kwota"])}${row(["1", "Czynsz", "2 400 zł"])}${row(["2", "Media", "350 zł"])}</w:tbl>`;
}
async function genDocx(n, file) {
  const body = [];
  let section = 0;
  for (let i = 0; i < n; i++) {
    if (i % 25 === 0) body.push(para(`§${++section} Rozdział ${section}`, "Heading1"));
    else if (i % 12 === 0) body.push(para(`Podrozdział ${section}.${i % 25}`, "Heading2"));
    if (i % 60 === 30) body.push(table(i));
    body.push(para(`${i + 1}. ${LOREM}`, null, { num: i % 9 === 3, bold: i % 7 === 0 }));
  }
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`);
  zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style></w:styles>`);
  zip.file("word/numbering.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.join("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

// ── serwer ────────────────────────────────────────────────────────────────────
function up() {
  return new Promise((r) => { const q = http.get({ host: "127.0.0.1", port: PORT, path: "/", timeout: 800 }, (s) => { s.resume(); r(true); }); q.on("error", () => r(false)); q.on("timeout", () => { q.destroy(); r(false); }); });
}
async function ensureServer() {
  if (await up()) return null;
  const srv = spawn("python3", [path.join(ROOT, "scripts/test-server.py"), String(PORT)], { cwd: ROOT, stdio: "ignore" });
  for (let i = 0; i < 50; i++) { if (await up()) return srv; await new Promise((r) => setTimeout(r, 100)); }
  throw new Error("serwer nie wstał");
}

const median = (a) => { const s = [...a].sort((x, y) => x - y); return s.length ? s[Math.floor(s.length / 2)] : 0; };
const ms = (x) => `${Math.round(x)} ms`;

async function benchOne(browser, file, n) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.setItem("dwb-panel-docked-open-v1", "0"); });
  const page = await context.newPage();
  const cdp = await context.newCDPSession(page);
  await cdp.send("Emulation.setCPUThrottlingRate", { rate: CPU });

  // eksperyment A/B: dodatkowy CSS przed startem (np. INJECT_CSS='.x{...}')
  if (process.env.INJECT_CSS) await page.addInitScript((css) => { document.addEventListener("DOMContentLoaded", () => { const st = document.createElement("style"); st.textContent = css; document.head.appendChild(st); }); }, process.env.INJECT_CSS);
  const tBoot = Date.now();
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
  const boot = Date.now() - tBoot;
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  // biblioteki DOCX wczytane przed pomiarem otwarcia? (idle-prefetch po starcie)
  await page.waitForTimeout(3000);
  const libsReady = await page.evaluate(() => !!(window.JSZip && window.docx));

  // rozbicie otwarcia: owijamy funkcje z zewnątrz
  await page.evaluate(() => {
    window.__perf = {};
    const wrap = (name) => {
      const orig = window[name];
      if (typeof orig !== "function") return;
      window[name] = function timed(...args) {
        const t0 = performance.now();
        const r = orig.apply(this, args);
        const done = () => { window.__perf[name] = (window.__perf[name] || 0) + performance.now() - t0; };
        if (r && typeof r.then === "function") return r.finally(done);
        done();
        return r;
      };
    };
    ["ensureDocLibs", "renderDocxPreview", "readHeadingStyleClasses", "analyzeDocumentDom", "renderStructurePanel", "refreshInlineEditBaseline", "extractParagraphTextsFromDocx", "extractParagraphRunsFromDocx", "collectInlineParagraphEdits", "buildPatchedDocx"].forEach(wrap);
  });
  const tOpen = Date.now();
  await page.locator("#fileInput").setInputFiles(file);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 300000 });
  // gotowe = baza edycji też wczytana (bez niej pierwsza edycja czeka)
  await page.waitForFunction(() => typeof baselineParagraphRuns !== "undefined" && baselineParagraphRuns.length > 0, null, { timeout: 300000 });
  await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
  const open = Date.now() - tOpen;
  const breakdown = await page.evaluate(() => ({ ...window.__perf }));
  const info = await page.evaluate(() => ({ paras: document.querySelectorAll(".docx-preview-host p").length, pages: document.querySelectorAll(".docx-preview-host section.docx").length, nodes: document.querySelectorAll(".docx-preview-host *").length }));

  // przewijanie: klatki podczas płynnego zjazdu przez dokument
  const scroll = await page.evaluate(async () => {
    const vp = document.getElementById("docViewport");
    const frames = [];
    let last = performance.now();
    const end = Math.min(vp.scrollHeight - vp.clientHeight, 40000);
    await new Promise((resolve) => {
      const step = () => {
        const now = performance.now();
        frames.push(now - last);
        last = now;
        vp.scrollTop += 120;
        if (vp.scrollTop >= end - 2 || frames.length > 240) resolve(); else requestAnimationFrame(step);
      };
      requestAnimationFrame(step);
    });
    frames.shift();
    frames.sort((a, b) => a - b);
    return { median: frames[Math.floor(frames.length / 2)], p95: frames[Math.floor(frames.length * 0.95)], long: frames.filter((f) => f > 50).length, n: frames.length };
  });

  // pisanie: czas od klawisza do namalowania (Event Timing), osobno 1. klawisz „serii”
  await page.evaluate(() => { document.getElementById("docViewport").scrollTop = 0; });
  await page.click('.mode-btn[data-mode="edit"]');
  await page.waitForTimeout(400);
  const typing = [];
  for (let burst = 0; burst < 3; burst++) {
    await page.evaluate((b) => {
      const el = [...document.querySelectorAll(".docx-editable-p")][5 + b * 3];
      el.focus();
      const r = document.createRange(); r.selectNodeContents(el); r.collapse(false);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
    }, burst);
    await page.waitForTimeout(1700); // przerwa > 1,5 s = nowy krok cofania (najdroższy klawisz)
    const res = await page.evaluate(async () => {
      const times = [];
      const el = document.activeElement;
      for (const ch of " abcdef") {
        const t0 = performance.now();
        el.dispatchEvent(new KeyboardEvent("keydown", { key: ch, bubbles: true }));
        el.dispatchEvent(new InputEvent("beforeinput", { inputType: "insertText", data: ch, bubbles: true, cancelable: true }));
        document.execCommand("insertText", false, ch);
        await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
        times.push(performance.now() - t0);
      }
      return times;
    });
    typing.push(res);
  }
  const firstKey = median(typing.map((t) => t[0]));
  const nextKeys = median(typing.flatMap((t) => t.slice(1)));

  // cofnięcie samego pisania (ostatnia seria, bez Entera)
  const undoTyping = await page.evaluate(async () => { const t0 = performance.now(); await dwbUndo.undo(); await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))); return performance.now() - t0; });

  // Enter w środku dokumentu (kolejka przebudowy) — czas do „plik gotowy”
  await page.evaluate(() => {
    const el = [...document.querySelectorAll(".docx-editable-p")][20];
    el.focus();
    const r = document.createRange(); r.selectNodeContents(el); r.collapse(false);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  });
  const enter = await page.evaluate(async () => {
    const t0 = performance.now();
    document.activeElement.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    const visible = performance.now() - t0;
    await waitInlineStructuralIdle();
    return { visible, file: performance.now() - t0 };
  });

  // szukanie, cofnięcie, zapis
  await page.click('.mode-btn[data-mode="read"]');
  const search = await page.evaluate(async () => {
    await ensureLazyFeature("find-replace");
    searchQueryEl.value = "Wynajmującego";
    const t0 = performance.now();
    await runFindReplaceScan();
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    return { ms: performance.now() - t0, hits: frMatches.length };
  });
  const undo = await page.evaluate(async () => { const t0 = performance.now(); await dwbUndo.undo(); return performance.now() - t0; });
  // zapis z niezapisanym pisaniem w kilku akapitach (tak wygląda zwykły zapis)
  const save = await page.evaluate(async () => {
    [...document.querySelectorAll(".docx-preview-host p")].slice(10, 13).forEach((p) => p.append(" zmiana"));
    await new Promise((r) => setTimeout(r, 50));
    const t0 = performance.now();
    const b = await buildDocumentForSave();
    return { ms: performance.now() - t0, kb: Math.round(b.byteLength / 1024) };
  });

  await context.close();
  return { n, info, boot, libsReady, open, breakdown, scroll, firstKey, nextKeys, undoTyping, enter, search, undo, save };
}

async function main() {
  const srv = await ensureServer();
  const browser = await pw.chromium.launch({ headless: true });
  try {
    console.log(`\nDocuments Workbench — pomiar wydajności (CPU ×${CPU})\n`);
    for (const n of SIZES) {
      const file = path.join(os.tmpdir(), `dwb-bench-${n}.docx`);
      if (!fs.existsSync(file)) await genDocx(n, file);
      const r = await benchOne(browser, file, n);
      const kb = Math.round(fs.statSync(file).size / 1024);
      console.log(`── ${n} akapitów tekstu (${r.info.paras} <p>, ${r.info.nodes} węzłów DOM, plik ${kb} KB) ──`);
      console.log(`  start apki:            ${ms(r.boot)}   biblioteki gotowe przed otwarciem: ${r.libsReady ? "tak" : "nie"}`);
      console.log(`  OTWARCIE (do edycji):  ${ms(r.open)}`);
      Object.entries(r.breakdown).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => console.log(`      ${k.padEnd(30)} ${ms(v)}`));
      console.log(`  przewijanie:           mediana klatki ${ms(r.scroll.median)}, p95 ${ms(r.scroll.p95)}, klatek >50 ms: ${r.scroll.long}/${r.scroll.n}`);
      console.log(`  pisanie:               1. klawisz serii ${ms(r.firstKey)}, kolejne ${ms(r.nextKeys)}`);
      console.log(`  Enter:                 widać ${ms(r.enter.visible)}, plik gotowy ${ms(r.enter.file)}`);
      console.log(`  szukanie:              ${ms(r.search.ms)} (${r.search.hits} trafień)`);
      console.log(`  cofnij pisanie:        ${ms(r.undoTyping)}`);
      console.log(`  cofnij (z Enterem):    ${ms(r.undo)}`);
      console.log(`  zapis (budowa pliku):  ${ms(r.save.ms)} (${r.save.kb} KB)\n`);
    }
  } finally {
    await browser.close();
    if (srv) srv.kill();
  }
}

main().catch((e) => { console.error(e); process.exit(1); });
