// keys-table-playwright.js — przejście przez TABELĘ ZACHOWAŃ KLAWISZY (scripts/keys-table.js).
//
//   node scripts/keys-table-playwright.js               (Chromium)
//   ENGINE=webkit node scripts/keys-table-playwright.js
//   ONLY=rule node scripts/keys-table-playwright.js     (tylko jeden rodzaj akapitu)
//
// Każda komórka od świeżego dokumentu: budowa stanu początkowego (sprawdzana z tabelą), kursor,
// klawisz, odczyt wyniku z ZAPISANEGO pliku (rodzaje akapitów) + kursor z ekranu. Komórki „?”
// nie psują wyniku — raport pokazuje, co aplikacja robi dziś (do decyzji).
// Raport: output/keys-table-<silnik>.md.
const fs = require("fs");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const { contexts, expect } = require("./keys-table");
const { keysStateIn } = require("./keys-state");
const { buildDocx } = require("../tools/word-oracle/build-docx");
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ONLY = process.env.ONLY || "";
const KEYS = { bs: ["Backspace", false], del: ["Delete", true], enterStart: ["Enter", false], enterEnd: ["Enter", true], enter: ["Enter", false] };
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(350));

// stan dokumentu w zapisie tabeli: rodzaje z pliku (scripts/keys-state.js — ten sam odczyt co
// wyrocznia Worda), kursor z ekranu
async function readState(page, withCaret = true) {
  const r = await page.evaluate(async (caret) => {
    const zip = await JSZip.loadAsync(await buildDocumentForSave());
    const xml = await zip.file("word/document.xml").async("string");
    if (!caret) return { xml };
    const sel = getSelection();
    const els = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const i = sel?.rangeCount ? els.findIndex((p) => p.contains(sel.anchorNode)) : -1;
    return { xml, at: `${i}:${i >= 0 ? getCaretOffset(els[i]) : "?"}` };
  }, withCaret);
  return (await keysStateIn(page, r.xml)) + (r.at ? ` @${r.at}` : "");
}

const caretAt = (page, i, end) => page.evaluate(([idx, e]) => {
  const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[idx];
  p.scrollIntoView({ block: "center" });
  placeCaret(p, e ? p.textContent.replace(/﻿/g, "").length : 0);
}, [i, end]);

async function build(page, name) {
  const ctx = contexts[name];
  if (ctx.body) { // gotowy plik — ten sam, który dostaje wyrocznia Worda
    await page.locator("#fileInput").setInputFiles({ name: `${name}.docx`, mimeType: DOCX_MIME, buffer: await buildDocx({ body: ctx.body }) });
    await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
    await idle(page);
    await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
    await idle(page);
    return;
  }
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);
  const edit = (e) => page.evaluate((x) => applyDocumentEdit(x), e).then(() => idle(page));
  if (["rule", "underRule", "line"].includes(name)) {
    await page.keyboard.type("Alfa");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Gamma");
    await idle(page);
    await edit({ op: "hrule", index: 0 });
    if (name === "line") { await caretAt(page, 1, false); await page.keyboard.type("Beta"); await idle(page); }
    return;
  }
  const beta = /Empty$|^empty$/.test(name) ? "" : "Beta";
  await page.keyboard.type("Alfa");
  await page.keyboard.press("Enter");
  if (beta) await page.keyboard.type(beta);
  await page.keyboard.press("Enter");
  await page.keyboard.type("Gamma");
  await idle(page);
  if (name.startsWith("list")) await edit({ op: "list", indices: [1], kind: "bullet" });
  else if (name === "ind") await edit({ op: "paraFormat", indices: [1], ind: { left: 720 } });
  else if (name === "h1") await edit({ op: "paraFormat", indices: [1], style: "h1" });
  else if (name.startsWith("quote")) await edit({ op: "paraFormat", indices: [1], style: "quote" });
}

async function main() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const errors = [];
  // każda komórka w świeżym oknie — stan jednej (niezapisany dokument, szkic) nie wpływa na następną
  async function freshPage() {
    const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
    await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
    const page = await context.newPage();
    page.on("pageerror", (e) => errors.push(e.message));
    page.on("dialog", (d) => d.accept());
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    return { page, context };
  }

  const rows = [];
  for (const [name, ctx] of Object.entries(contexts)) {
    if (ONLY && name !== ONLY) continue;
    for (const [key, want] of Object.entries(expect[name] || {})) {
      const cell = { name, key, want };
      const { page, context } = await freshPage();
      try {
        await build(page, name);
        cell.start = await readState(page, false);
        if (cell.start !== ctx.start) { cell.status = "build"; cell.got = cell.start; rows.push(cell); await context.close(); continue; }
        const [press, atEnd] = KEYS[key];
        await caretAt(page, ctx.at, atEnd);
        await page.keyboard.press(press);
        await idle(page);
        cell.got = await readState(page);
        cell.status = want === "?" ? "open" : cell.got === want ? "ok" : "fail";
      } catch (e) {
        cell.status = "fail";
        cell.got = `błąd: ${String(e.message || e).split("\n")[0]}`;
      }
      await context.close();
      rows.push(cell);
    }
  }
  await browser.close();

  const icon = { ok: "✅", fail: "❌", open: "❔", build: "⚠️" };
  const lines = rows.map((r) => `${icon[r.status]} ${r.name} · ${r.key}${r.status === "ok" ? "" : r.status === "open" ? ` — dziś: ${r.got}` : r.status === "build" ? ` — zły stan początkowy: ${r.got} (tabela: ${contexts[r.name].start})` : ` — oczekiwane: ${r.want} | jest: ${r.got}`}`);
  lines.forEach((l) => console.log(l));
  if (errors.length) console.log(`❌ błędy strony: ${errors.slice(0, 3).join(" | ")}`);
  const bad = rows.filter((r) => r.status === "fail" || r.status === "build").length + (errors.length ? 1 : 0);
  const open = rows.filter((r) => r.status === "open").length;
  const summary = `[${ENGINE}] ${rows.filter((r) => r.status === "ok").length}/${rows.length} OK, do decyzji: ${open}, niezgodne: ${bad}`;
  console.log(`\n${summary}`);
  fs.mkdirSync(path.join(__dirname, "../output"), { recursive: true });
  fs.writeFileSync(path.join(__dirname, `../output/keys-table-${ENGINE}.md`), `# Tabela klawiszy — ${new Date().toISOString().slice(0, 10)}\n\n${summary}\n\n${lines.map((l) => `- ${l}`).join("\n")}\n`);
  process.exit(bad ? 1 : 0);
}

main().catch((e) => { console.error(e); process.exit(1); });
