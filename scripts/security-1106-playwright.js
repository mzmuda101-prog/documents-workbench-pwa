// Strażnik poprawek z audytu bezpieczeństwa Docs (2026-11-06, run-1):
//   1. DOCX z w:altChunk (HTML z pliku): bez <iframe> w podglądzie i bez żądań na zewnątrz,
//   2. wklejenie application/x-dwb-comment-fragment z obcej strony: surowy HTML nie trafia
//      do dokumentu (w schowku tylko klucz do fragmentu w pamięci),
//   3. CSP strony bez https://esm.sh, z frame-src 'none',
//   4. worker modułowy dalej importuje z esm.sh (workery AI mają własne CSP, nie strony).
//   node scripts/security-1106-playwright.js        ENGINE=webkit node scripts/security-1106-playwright.js
const playwright = require("playwright");
const fs = require("fs");
const path = require("path");
const JSZip = require(path.join(__dirname, "..", "lib", "jszip.min.js"));
const { APP_URL, BUILTIN_SAMPLE, waitDocReady } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE || "chromium";

async function buildAltChunkDocx() {
  const z = await JSZip.loadAsync(fs.readFileSync(BUILTIN_SAMPLE));
  let doc = await z.file("word/document.xml").async("string");
  doc = doc.replace("<w:body>", '<w:body><w:p><w:r><w:t>ALTCHUNK-DOC</w:t></w:r></w:p><w:altChunk r:id="rIdAltTest"/>');
  if (!/xmlns:r=/.test(doc)) doc = doc.replace("<w:document ", '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
  z.file("word/document.xml", doc);
  const rels = await z.file("word/_rels/document.xml.rels").async("string");
  z.file("word/_rels/document.xml.rels", rels.replace("</Relationships>", '<Relationship Id="rIdAltTest" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="chunk.html"/></Relationships>'));
  const ct = await z.file("[Content_Types].xml").async("string");
  if (!/Extension="html"/.test(ct)) z.file("[Content_Types].xml", ct.replace("<Default ", '<Default Extension="html" ContentType="text/html"/><Default '));
  z.file("word/chunk.html", '<!doctype html><html><body><p>altchunk</p><script src="https://esm.sh/dwb-test-payload"></script></body></html>');
  return z.generateAsync({ type: "nodebuffer" });
}

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const esmHits = [];
  await context.route("https://esm.sh/**", (route) => {
    esmHits.push(route.request().url());
    const stub = /dwb-test-stub/.test(route.request().url());
    route.fulfill({ status: 200, contentType: "application/javascript", headers: { "Access-Control-Allow-Origin": "*" }, body: stub ? "export const ok = 'działa';" : "parent.window.__dwbPwned = 1;" });
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  const fails = [];
  const check = (name, ok, detail) => { console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`); if (!ok) fails.push(name); };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());

  // 3. CSP
  const csp = await page.evaluate(() => document.querySelector('meta[http-equiv="Content-Security-Policy"]').content);
  check("CSP strony bez esm.sh, z frame-src 'none'", !/esm\.sh|huggingface/.test(csp) && /frame-src 'none'/.test(csp), csp);

  // 1. altChunk
  await page.locator("#fileInput").setInputFiles({ name: "altchunk.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await buildAltChunkDocx() });
  await waitDocReady(page);
  await page.waitForTimeout(800);
  const alt = await page.evaluate(() => ({ text: (document.querySelector(".docx-preview-host")?.textContent || "").includes("ALTCHUNK-DOC"), iframes: document.querySelectorAll(".docx-preview-host iframe").length, pwned: !!window.__dwbPwned }));
  check("altChunk: dokument otwarty, bez <iframe>, skrypt z pliku nie ruszył", alt.text && alt.iframes === 0 && !alt.pwned && !esmHits.length, { ...alt, esmHits });

  // 2. cudza wklejka
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForFunction(() => document.querySelector(".docx-editable-p"), null, { timeout: 10000 });
  await page.locator(".docx-editable-p").first().click();
  await page.keyboard.press("End");
  const pasted = await page.evaluate(() => {
    const dt = new DataTransfer();
    dt.setData("text/plain", "zwykły tekst");
    dt.setData("application/x-dwb-comment-fragment", '<span data-cm="a" data-cm-kind="start" data-cm-xml="&lt;w:commentRangeStart w:id=&quot;7&quot;/&gt;">x</span><iframe srcdoc="&lt;b&gt;obce&lt;/b&gt;"></iframe><img src="x" onerror="window.__dwbPwned=2">');
    const target = window.getSelection().anchorNode?.parentElement || document.querySelector(".docx-editable-p");
    target.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
    const host = document.querySelector(".docx-preview-host");
    return { iframes: host.querySelectorAll("iframe").length, foreignMarks: host.querySelectorAll('[data-cm^="paste"]').length, imgs: host.querySelectorAll('img[src="x"]').length, pwned: !!window.__dwbPwned, plain: host.textContent.includes("zwykły tekst") };
  });
  check("cudza wklejka: bez ramki, znaczników i obrazów z obcego HTML", !pasted.iframes && !pasted.foreignMarks && !pasted.imgs && !pasted.pwned, pasted);

  // 4. worker z importem z esm.sh
  const workerMsg = await page.evaluate(() => new Promise((res) => {
    const w = new Worker("scripts/fixtures/esm-import-worker.js", { type: "module" });
    w.onmessage = (e) => { res(e.data); w.terminate(); };
    w.onerror = (e) => res(`onerror ${e.message || ""}`);
    setTimeout(() => res("timeout"), 8000);
  }));
  check("worker modułowy dalej importuje z esm.sh", workerMsg === "działa", workerMsg);

  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  if (fails.length) { console.log(`❌ security-1106 (${ENGINE}): ${fails.length} błędów`); process.exit(1); }
  console.log(`✅ security-1106 (${ENGINE}): wszystko OK`);
}
run().catch((e) => { console.error(e); process.exit(1); });
