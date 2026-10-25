// Monkey-test: otwiera plik i wykonuje N losowych czynności (klik w widoczny
// przycisk/akapit, pisanie, klawisze, przewijanie), zbierając wyjątki JS i błędy
// konsoli. Ziarno (SEED) czyni przebieg powtarzalnym — błąd da się odtworzyć krok po kroku.
//
//   npm run serve:test   (osobno)
//   FILE=docs/samples/przewodnik.docx SEED=7 STEPS=400 node scripts/monkey-playwright.js
//   ENGINE=webkit MOBILE=1 VERBOSE=1 …
// UWAGA: WebKit blokuje część portów („restricted network port”, np. 4190) — używaj 7823.
const playwright = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:7823/";
const FILE = process.env.FILE;
const SEED = Number(process.env.SEED || 1);
const STEPS = Number(process.env.STEPS || 300);
const ENGINE = process.env.ENGINE || "chromium";
const MOBILE = process.env.MOBILE === "1";
const VERBOSE = process.env.VERBOSE === "1";

let s = SEED >>> 0 || 1;
const rnd = () => ((s = (s * 1664525 + 1013904223) >>> 0) / 2 ** 32);
const pick = (a) => a[Math.floor(rnd() * a.length)];

const TEXTS = ["abc", "Zażółć gęślą", "123", "-4,5", "=SUMA(A1:A3)", "2026-10-24", "<b>x</b>", "😀👍🏽", "  ", "\t", "x".repeat(300), "1e309", "#ADR!", "'quoted", "a\"b", "&amp;"];
const KEYS = ["Escape", "Enter", "Tab", "Shift+Tab", "ArrowDown", "ArrowUp", "ArrowLeft", "ArrowRight", "Backspace", "Delete", "Home", "End", "PageDown", "Control+z", "Control+y", "Meta+z", "Shift+ArrowRight", "Control+a", "F2"];

async function main() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const context = await browser.newContext({
    viewport: MOBILE ? { width: 390, height: 844 } : { width: 1366, height: 900 },
    isMobile: MOBILE && ENGINE !== "firefox",
    hasTouch: MOBILE,
    serviceWorkers: "block",
    acceptDownloads: true,
  });
  await context.addInitScript(() => {
    try { localStorage.setItem("introPlayed", "true"); } catch {}
    window.print = () => {};
    window.open = () => null;
    window.alert = () => {};
    window.confirm = () => true;
    window.prompt = () => null;
  });
  const page = await context.newPage();
  const errors = [];
  const history = [];
  const record = (kind, msg) => {
    if (errors.some((e) => e.msg === msg)) return;
    errors.push({ kind, msg, step: history.length, last: history.slice(-6) });
  };
  page.on("pageerror", (e) => record("pageerror", (e.stack || e.message).split("\n").slice(0, 4).join(" | ")));
  page.on("console", (m) => { if (m.type() === "error") record("console", m.text().slice(0, 400)); });
  page.on("dialog", (d) => d.dismiss().catch(() => {}));
  page.on("download", (d) => d.cancel().catch(() => {}));

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForTimeout(600);
  if (FILE) {
    await page.setInputFiles("#fileInput", path.resolve(FILE));
    await page.waitForTimeout(2500);
  }

  for (let i = 0; i < STEPS; i++) {
    if (!page.url().startsWith(APP_URL)) { history.push("nav-back " + page.url()); await page.goto(APP_URL).catch(() => {}); }
    const r = rnd();
    let desc = "";
    try {
      if (r < 0.62) {
        // klik w losowy widoczny element sterujący
        const cands = await page.evaluate(() => {
          const sel = "button, [role=button], [role=menuitem], [role=tab], summary, select, input[type=checkbox], input[type=radio], a[href^='#'], td, th, [contenteditable=true] p, label";
          const out = [];
          document.querySelectorAll(sel).forEach((el) => {
            if (el.closest("[hidden], [inert]") || el.disabled) return;
            if (el.matches("input[type=file]") || el.closest("a[download]")) return;
            const b = el.getBoundingClientRect();
            if (b.width < 2 || b.height < 2 || b.bottom < 0 || b.right < 0 || b.top > innerHeight || b.left > innerWidth) return;
            const cs = getComputedStyle(el);
            if (cs.visibility === "hidden" || cs.pointerEvents === "none" || +cs.opacity === 0) return;
            const label = (el.id ? "#" + el.id : el.tagName.toLowerCase()) + " " + (el.getAttribute("aria-label") || el.textContent || "").trim().slice(0, 30);
            out.push({ x: b.left + b.width / 2, y: b.top + b.height / 2, label, tag: el.tagName });
          });
          return out;
        });
        if (!cands.length) continue;
        const c = pick(cands);
        desc = "click " + c.label.replace(/\s+/g, " ");
        if (c.tag === "SELECT") {
          desc = "select " + c.label;
          await page.mouse.click(c.x, c.y);
          await page.keyboard.press("ArrowDown");
          await page.keyboard.press("Enter");
        } else if (MOBILE && rnd() < 0.7) await page.touchscreen.tap(c.x, c.y);
        else await page.mouse.click(c.x, c.y, { clickCount: rnd() < 0.1 ? 2 : 1 });
      } else if (r < 0.77) {
        const t = pick(TEXTS);
        desc = "type " + JSON.stringify(t.slice(0, 20));
        await page.keyboard.type(t, { delay: 0 });
      } else if (r < 0.93) {
        const k = pick(KEYS);
        desc = "key " + k;
        await page.keyboard.press(k);
      } else {
        const dy = (rnd() - 0.4) * 1600;
        desc = "scroll " + Math.round(dy);
        await page.mouse.move(400, 400);
        await page.mouse.wheel(0, dy);
      }
    } catch (e) {
      desc += " (!" + String(e.message).split("\n")[0].slice(0, 60) + ")";
    }
    history.push(desc);
    if (VERBOSE) console.log(i, desc);
    await page.waitForTimeout(rnd() < 0.2 ? 400 : 60);
  }
  await page.waitForTimeout(800);
  await browser.close();
  console.log(`monkey ${ENGINE}${MOBILE ? " mobile" : ""} seed=${SEED} steps=${STEPS} file=${FILE ? path.basename(FILE) : "-"} → ${errors.length} error(s)`);
  for (const e of errors) console.log(JSON.stringify(e, null, 1));
  process.exit(errors.length ? 1 : 0);
}
main().catch((e) => { console.error(e); process.exit(2); });
