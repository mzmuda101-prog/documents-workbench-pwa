// replay-report.js — odtworzenie zgłoszenia „Zgłoś błąd” (app/bug-report.js) krok po kroku na tym
// samym dokumencie: ten sam plik z początku sesji, te same klawisze w tych samych miejscach,
// te same operacje. Na końcu: czy wynik jest taki jak u zgłaszającego + niezmienniki (podgląd =
// plik, wyspy, bez błędów strony).
//
//   REPORT=~/Downloads/zgloszenie-….dwbbug.json node scripts/replay-report.js
//   REPORT=… ENGINE=webkit VERBOSE=1 node scripts/replay-report.js
//
// Zgłoszenie, które pokazuje błąd → poprawka → zgłoszenie (bez prywatnego dokumentu albo
// z dokumentem zastępczym) jako stały test regresji.
const fs = require("fs");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const FILE = process.env.REPORT ? path.resolve(process.env.REPORT.replace(/^~/, process.env.HOME)) : null;
const VERBOSE = !!process.env.VERBOSE;
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 }).catch(() => {}).then(() => page.waitForTimeout(250));

// kursor / zaznaczenie: [akapit, znak] — znaki bez niewidocznych ﻿ (tak liczy dziennik)
const place = (page, at, to) => page.evaluate(([a, b]) => {
  const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
  const point = ([i, k]) => {
    const p = ps[i];
    if (!p) return null;
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let left = k, last = null;
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      last = n;
      for (let j = 0; j < n.data.length; j++) {
        if (n.data[j] === "﻿") continue;
        if (left === 0) return [n, j];
        left--;
      }
      if (left === 0) return [n, n.data.length];
    }
    return last ? [last, last.data.length] : [p, 0];
  };
  const s = point(a);
  if (!s) return false;
  const e = b ? point(b) : s;
  ps[a[0]].closest(".docx-edit-root")?.focus();
  const sel = getSelection();
  sel.removeAllRanges();
  const r = document.createRange();
  r.setStart(s[0], s[1]);
  r.collapse(true);
  sel.addRange(r);
  if (b && e) sel.extend(e[0], e[1]);
  return true;
}, [at, to || null]);

async function main() {
  if (!FILE || !fs.existsSync(FILE)) { console.log("Podaj REPORT=ścieżka do pliku .dwbbug.json"); process.exit(2); }
  const rep = JSON.parse(fs.readFileSync(FILE, "utf8"));
  if (rep.kind !== "dwb-bug-report") { console.log("To nie jest zgłoszenie Documents Workbench"); process.exit(2); }
  console.log(`Zgłoszenie: ${rep.file}, ${rep.at}, aplikacja ${rep.app}, ${rep.actions.length} działań, ${rep.viewport?.mobile ? "telefon" : "komputer"}`);
  if (rep.errors?.length) console.log(`  błędy strony u zgłaszającego: ${rep.errors.slice(0, 3).join(" | ")}`);

  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext(rep.viewport?.mobile
    ? { serviceWorkers: "block", viewport: { width: rep.viewport.w, height: rep.viewport.h }, hasTouch: true, ...(ENGINE === "chromium" ? { isMobile: true } : {}) }
    : { serviceWorkers: "block", viewport: { width: Math.max(1024, rep.viewport?.w || 1280), height: rep.viewport?.h || 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept().catch(() => {}));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles({ name: rep.file || "zgloszenie.docx", mimeType: DOCX_MIME, buffer: Buffer.from(rep.start, "base64") });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host p"), null, { timeout: 60000 });
  await idle(page);

  const setMode = (mode) => page.evaluate((m) => { const ro = m === "read"; if (readOnlyMode !== ro) { readModeEl.checked = ro; readModeEl.dispatchEvent(new Event("change")); } }, mode);
  let skipped = 0;
  let prevMs = null;
  for (const [n, a] of rep.actions.entries()) {
    // przerwa ≥ 1,5 s u zgłaszającego zamyka krok pisania (undo.js) — bez niej kroki Cofnij
    // układałyby się inaczej i Cofnij cofałoby co innego; krótsze przerwy pomijamy (szybciej)
    if (prevMs != null && a.ms - prevMs >= 1500) await page.waitForTimeout(1600);
    prevMs = a.ms;
    if (VERBOSE) console.log(`  ${n + 1}. ${a.a}${a.key ? ` ${a.mods ? `${a.mods}+` : ""}${a.key}` : ""}${a.at ? ` @${a.at.join(":")}${a.to ? `→${a.to.join(":")}` : ""}` : ""}${a.edit ? ` ${a.edit.op}` : ""}${a.fromInput ? " (skutek klawisza — pomijam)" : ""}`);
    try {
      if (a.a === "open") { if (a.mode) await setMode(a.mode); await idle(page); continue; }
      if (a.a === "mode") { await setMode(a.mode); await idle(page); continue; }
      if (a.fromInput) continue; // operację wywołał klawisz / wklejenie — odtworzy się sama
      if (a.at) await place(page, a.at, a.to);
      if (a.a === "click") continue;
      if (a.a === "key") {
        const printable = a.key.length === 1 && (!a.mods || a.mods === "Shift");
        if (printable) await page.keyboard.insertText(a.key);
        else await page.keyboard.press(a.mods ? `${a.mods}+${a.key}` : a.key);
      } else if (a.a === "text") {
        await page.keyboard.insertText(a.data || "");
      } else if (a.a === "paste") {
        await page.evaluate(([text, html]) => {
          const dt = new DataTransfer();
          if (text) dt.setData("text/plain", text);
          if (html) dt.setData("text/html", html);
          const target = getSelection().anchorNode?.parentElement || document.activeElement;
          target.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
        }, [a.text, a.html]);
      } else if (a.a === "edit") {
        if (JSON.stringify(a.edit).includes("…[przycięte")) { skipped++; console.log(`  ⚠️ ${n + 1}. operacja ${a.edit.op} przycięta w zgłoszeniu — pomijam`); continue; }
        await page.evaluate((e) => applyDocumentEdit(e), a.edit);
      } else if (a.a === "undo") await page.evaluate(() => dwbUndo.undo());
      else if (a.a === "redo") await page.evaluate(() => dwbUndo.redo());
      await idle(page);
    } catch (err) {
      console.log(`  ❌ ${n + 1}. ${a.a}: ${String(err.message || err).split("\n")[0]}`);
    }
  }
  await idle(page);

  // wynik: tekst akapitów — u zgłaszającego i w odtworzeniu; niezmienniki
  const res = await page.evaluate(async (endB64) => {
    const bin = Uint8Array.from(atob(endB64), (c) => c.charCodeAt(0));
    const want = (await extractParagraphTextsFromDocx(bin)).map(String);
    const bytes = await buildDocumentForSave();
    const got = (await extractParagraphTextsFromDocx(bytes)).map(String);
    const at = [...Array(Math.max(want.length, got.length)).keys()].find((i) => want[i] !== got[i]);
    const els = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const norm = (t) => String(t).replace(/[\t\n﻿]|\p{So}/gu, "");
    const dom = els.map((p) => (p.dataset.lock ? null : norm(previewRunsToPlainText(extractRunsFromPreviewParagraph(p)))));
    const file = got.map(norm);
    const pv = dom.findIndex((t, i) => t !== null && t !== file[i]);
    return { same: at === undefined, at, want: want[at]?.slice(0, 80), got: got[at]?.slice(0, 80), nWant: want.length, nGot: got.length, previewIsFile: pv < 0 && dom.length === file.length, pv };
  }, rep.end);
  await browser.close();
  console.log(res.same ? "✅ odtworzone: wynik taki jak u zgłaszającego" : `⚠️ wynik inny niż u zgłaszającego (akapit ${res.at}): „${res.want}” vs odtworzenie „${res.got}” (${res.nWant} / ${res.nGot} akapitów)`);
  console.log(`${res.previewIsFile ? "✅" : "❌"} podgląd = plik${res.previewIsFile ? "" : ` (akapit ${res.pv})`}`);
  console.log(`${errors.length ? "❌" : "✅"} błędy strony${errors.length ? `: ${errors.slice(0, 3).join(" | ")}` : ": brak"}`);
  if (skipped) console.log(`⚠️ pominięte operacje (przycięte w zgłoszeniu): ${skipped}`);
  process.exit(res.previewIsFile && !errors.length ? 0 : 1);
}

main().catch((e) => { console.error(e); process.exit(1); });
