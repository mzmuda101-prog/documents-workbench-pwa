// chaos-playwright.js — losowy użytkownik + strażnik spójności po każdym kroku (2026-10-06).
//
// Po co: błędy edytora wychodzą zwykle przy dziwnej kolejności zwykłych czynności (zaznacz przez
// dwa akapity → wklej → Cofnij → Enter w linku…), której nikt nie pisze w teście ręcznie. Tu
// „użytkownik” robi to losowo — PRAWDZIWĄ myszą i klawiaturą tam, gdzie się da: klika w tekst,
// zaznacza przeciąganiem, pisze (też polskie znaki, adresy, e-maile), Enter / Shift+Enter /
// Backspace / Delete / Tab, strzałki z Shiftem, B/I/U, styl akapitu, listy, wyrównanie, okienko
// linku (WWW / miejsce / e-mail), zakładki, wklejanie (tekst, wiersze, adres, HTML), podział
// strony, tabela, obraz, przypis, Cofnij / Ponów, Czytanie ↔ Edycja.
//
// Po KAŻDYM kroku: brak błędów strony, nic nie wisi (nakładka „wczytywanie”), strażnik spójności
// (app/self-check.js: podgląd = plik akapit po akapicie, paczka .docx, zakładki, linki, komentarze,
// przypisy, listy, spacje) — liczą się tylko problemy NOWE względem pliku na starcie. Na końcu:
// zapis → ponowne otwarcie = ten sam tekst; Cofnij wszystko = dokument z początku.
//
// Błąd → seria jest SKRACANA do najmniejszej, która go jeszcze wywołuje (ddmin), i zapisywana:
//   output/chaos/<silnik>-<ziarno>.json  (+ czytelna lista kroków w konsoli)
// Powtórzenie:  REPLAY=output/chaos/chromium-7.json node scripts/chaos-playwright.js
// Regresje:     CASES=scripts/chaos-cases node scripts/chaos-playwright.js  (naprawione znaleziska — w npm test)
//
// Zmienne: ENGINE=webkit · VIEW=mobile · SEEDS=1,2,3 (albo RUNS=50 → ziarna 1..50) · STEPS=40 · DOC=przewodnik|blank
//          SHRINK=0 (bez skracania) · STRICT_UNDO=0 (bez „Cofnij wszystko”) · VERBOSE=1

const fs = require("fs");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const STEPS = Number(process.env.STEPS || 40);
const SEEDS = process.env.RUNS ? Array.from({ length: Number(process.env.RUNS) }, (_, i) => i + 1) : (process.env.SEEDS || "1,2,3").split(",").map(Number);
const SHRINK = process.env.SHRINK !== "0";
const SHRINK_MAX_MS = Number(process.env.SHRINK_MAX_S || 240) * 1000;
const STRICT_UNDO = process.env.STRICT_UNDO !== "0";
const VERBOSE = !!process.env.VERBOSE;
// VIEW=mobile — telefon (390×844, dotyk): stuknięcia zamiast kliknięć, układ telefonu
const MOBILE = process.env.VIEW === "mobile";
const OUT_DIR = path.resolve(__dirname, "../output/chaos");

// ── losowość z ziarnem ────────────────────────────────────────────────────────
function rng(seed) {
  let s = (seed * 2654435761) % 2147483647 || 1;
  return () => ((s = (s * 16807) % 2147483647) / 2147483647);
}

// ── słownik działań: każde dostaje własne liczby losowe (r[0..7]) — powtórka = te same liczby ──
const TEXTS = [
  "zażółć gęślą jaźń", "Ala ma kota", " ", "  ", "test", "ĄĘÓŁŚŻŹĆŃ", "1234", "„cytat”", "a–b—c…", "x",
  "www.example.com ", "zobacz https://example.org/a?b=1. ", "jan@example.com ", "(nawias) ", "tab\there", "😀 emoji", "koniec.",
];
const PASTES = [
  { kind: "text", text: "wklejony tekst" },
  { kind: "text", text: "pierwszy wiersz\ndrugi wiersz\n\nczwarty" },
  { kind: "text", text: "https://example.com/wklejka" },
  { kind: "html", html: "<p><b>Pogrubione</b> i <a href=\"https://example.net\">link</a></p><ul><li>punkt 1</li><li>punkt 2</li></ul>", text: "Pogrubione i link\npunkt 1\npunkt 2" },
  { kind: "text", text: "   " },
];
const STYLES = ["normal", "h1", "h2", "h3", "title", "quote", "callout"];

const WEIGHTS = {
  click: 5, type: 8, enter: 3, shiftEnter: 1, backspace: 4, del: 2, tab: 1, shiftTab: 0.5,
  arrows: 3, drag: 2.5, selectAll: 0.3, bold: 1.5, italic: 0.7, underline: 0.5, style: 1, list: 1,
  align: 0.5, link: 1, autolink: 1, bookmark: 0.6, paste: 2, pageBreak: 0.5, table: 0.4, image: 0.3,
  footnote: 0.3, undo: 2.5, redo: 1, mode: 0.5, typeInLink: 0.5,
};
const NAMES = Object.keys(WEIGHTS);
const TOTAL_W = NAMES.reduce((a, n) => a + WEIGHTS[n], 0);
function pickAction(r) {
  let x = r * TOTAL_W;
  for (const n of NAMES) { x -= WEIGHTS[n]; if (x <= 0) return n; }
  return NAMES[NAMES.length - 1];
}
function planActions(seed, n) {
  const rnd = rng(seed);
  return Array.from({ length: n }, () => ({ name: pickAction(rnd()), r: Array.from({ length: 8 }, () => Math.round(rnd() * 1e6) / 1e6) }));
}

// ── czekanie na spokój ────────────────────────────────────────────────────────
async function idle(page, ms = 250) {
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 });
  await page.evaluate(() => (typeof waitInlineStructuralIdle === "function" ? Promise.race([waitInlineStructuralIdle(), new Promise((r) => setTimeout(r, 8000))]) : null));
  await page.waitForTimeout(ms);
}

// Punkt w tekście (współrzędne na ekranie) w akapicie fracP (0..1) na pozycji fracC (0..1).
const pointIn = (page, fracP, fracC, editableOnly = true) => page.evaluate(([fp, fc, eo]) => {
  const host = document.querySelector(".docx-preview-host");
  const ps = collectPreviewParagraphElements(host).filter((p) => (!eo || p.classList.contains("docx-editable-p")) && p.getClientRects().length);
  if (!ps.length) return null;
  const p = ps[Math.min(ps.length - 1, Math.floor(fp * ps.length))];
  p.scrollIntoView({ block: "center" });
  const len = (p.textContent || "").length;
  const off = Math.min(len, Math.floor(fc * (len + 1)));
  let rect = null;
  try {
    const r = formDomRange(p, Math.max(0, off - (off === len && len ? 1 : 0)), Math.max(0, off - (off === len && len ? 1 : 0)) + (len ? 1 : 0));
    const rs = r ? [...r.getClientRects()].filter((x) => x.width || x.height) : [];
    rect = rs[0] || null;
  } catch (_) { /* pusty akapit */ }
  const b = rect || p.getBoundingClientRect();
  const x = rect ? (off === len && len ? b.right - 0.5 : b.left + 0.5) : b.left + 4;
  return { x, y: b.top + Math.min(b.height / 2, 9), i: resolveParaIndex(p), off };
}, [fracP, fracC, editableOnly]);

// ── wykonanie działań ─────────────────────────────────────────────────────────
async function closePops(page) {
  for (let k = 0; k < 2; k++) {
    if (!(await page.$(".compose-pop, .image-viewer"))) return;
    await page.keyboard.press("Escape");
    await page.waitForTimeout(80);
  }
}
async function ensureEdit(page) {
  if (await page.evaluate(() => readOnlyMode)) { await page.evaluate(() => appFrame.setReadOnly(false)); await idle(page); }
}
async function clickAt(page, r0, r1) {
  const pt = await pointIn(page, r0, r1);
  if (!pt) return null;
  if (MOBILE) await page.touchscreen.tap(pt.x, pt.y); else await page.mouse.click(pt.x, pt.y);
  return pt;
}

const ACT = {
  async click(page, r) { const pt = await clickAt(page, r[0], r[1]); return pt && `klik w akapit ${pt.i} (znak ${pt.off})`; },
  async type(page, r) { const t = TEXTS[Math.floor(r[0] * TEXTS.length)]; await page.keyboard.type(t, { delay: 5 }); return `pisze „${t.replace(/\t/g, "⇥")}”`; },
  async enter(page) { await page.keyboard.press("Enter"); return "Enter"; },
  async shiftEnter(page) { await page.keyboard.press("Shift+Enter"); return "Shift+Enter"; },
  async backspace(page, r) { const n = 1 + Math.floor(r[0] * 4); for (let i = 0; i < n; i++) await page.keyboard.press("Backspace"); return `Backspace ×${n}`; },
  async del(page, r) { const n = 1 + Math.floor(r[0] * 3); for (let i = 0; i < n; i++) await page.keyboard.press("Delete"); return `Delete ×${n}`; },
  async tab(page) { await page.keyboard.press("Tab"); return "Tab"; },
  async shiftTab(page) { await page.keyboard.press("Shift+Tab"); return "Shift+Tab"; },
  async arrows(page, r) {
    const keys = ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Home", "End"];
    const k = keys[Math.floor(r[0] * keys.length)];
    const shift = r[1] < 0.4;
    const n = 1 + Math.floor(r[2] * 5);
    for (let i = 0; i < n; i++) await page.keyboard.press(`${shift ? "Shift+" : ""}${k}`);
    return `${shift ? "Shift+" : ""}${k} ×${n}`;
  },
  async drag(page, r) {
    const a = await pointIn(page, r[0], r[1]);
    if (!a) return null;
    // drugi koniec: ten sam albo jeden–dwa akapity dalej (na ekranie)
    const b = await pointIn(page, Math.min(0.999, r[0] + (r[2] < 0.5 ? 0 : r[3] * 0.08)), r[4]);
    if (!b) return null;
    const a2 = await pointIn(page, r[0], r[1]); // po przewinięciu do b — odczyt a jeszcze raz
    await page.mouse.move(a2.x, a2.y);
    await page.mouse.down();
    await page.mouse.move((a2.x + b.x) / 2, (a2.y + b.y) / 2, { steps: 3 });
    await page.mouse.move(b.x, b.y, { steps: 3 });
    await page.mouse.up();
    return `zaznacza myszą od akapitu ${a2.i}:${a2.off} do ${b.i}:${b.off}`;
  },
  async selectAll(page) { await page.keyboard.press(`${MOD}+a`); return "Ctrl/⌘+A"; },
  async bold(page) { await page.keyboard.press(`${MOD}+b`); return "Ctrl/⌘+B"; },
  async italic(page) { await page.keyboard.press(`${MOD}+i`); return "Ctrl/⌘+I"; },
  async underline(page) { await page.keyboard.press(`${MOD}+u`); return "Ctrl/⌘+U"; },
  async style(page, r) { const st = STYLES[Math.floor(r[0] * STYLES.length)]; await page.evaluate((st) => composeUi.applyStyle(st), st); return `styl akapitu „${st}”`; },
  async list(page, r) { const k = ["bullet", "number", "none"][Math.floor(r[0] * 3)]; await page.evaluate((k) => composeUi.applyList(k), k); return `lista: ${k}`; },
  async align(page, r) { const k = ["left", "center", "right", "both"][Math.floor(r[0] * 4)]; await page.evaluate((k) => composeUi.applyAlign(k), k); return `wyrównanie: ${k}`; },
  async link(page, r) {
    await page.keyboard.press(`${MOD}+k`);
    if (!(await page.waitForSelector(".compose-pop-link .lf-ok", { timeout: 1500 }).catch(() => null))) return "Ctrl/⌘+K (bez okienka)";
    const mode = ["web", "doc", "mail"][Math.floor(r[0] * 3)];
    await page.click(`.lf-mode button[data-mode=${mode}]`);
    if (mode === "web") await page.fill(".lf-url", ["example.com", "https://example.org/x", "zły adres", ""][Math.floor(r[1] * 4)]);
    if (mode === "mail") await page.fill(".lf-addr", ["a@example.com", "zły"][Math.floor(r[1] * 2)]);
    if (mode === "doc") await page.evaluate((x) => { const b = [...document.querySelectorAll(".lf-place")]; b[Math.floor(x * b.length)]?.click(); }, r[1]);
    if (r[2] < 0.3) await page.fill(".lf-text", ["nowy tekst", "", "ĄĘ"][Math.floor(r[3] * 3)]);
    if (r[4] < 0.2) { await page.click(".lf-tip-toggle").catch(() => {}); await page.fill(".lf-tip-in", "etykietka").catch(() => {}); }
    await page.click(".lf-ok");
    return `okienko linku: ${mode}`;
  },
  async autolink(page, r) { const t = ["www.chaos.pl ", "kontakt@chaos.pl ", "https://chaos.pl/a). "][Math.floor(r[0] * 3)]; await page.keyboard.type(t); return `pisze adres „${t}”`; },
  async typeInLink(page, r) {
    // kursor na końcu/początku istniejącego linku i pisanie
    const ok = await page.evaluate((x) => {
      const as = [...document.querySelectorAll(".docx-preview-host .docx-editable-p a[href]:not(.doc-xref)")];
      const a = as[Math.floor(x[0] * as.length)];
      if (!a) return false;
      a.scrollIntoView({ block: "center" });
      let n = a; while (n.lastChild) n = n.lastChild;
      let f = a; while (f.firstChild) f = f.firstChild;
      const r = document.createRange();
      if (x[1] < 0.5) r.setStart(n, (n.textContent || "").length); else r.setStart(f, 0);
      r.collapse(true);
      a.closest(".docx-edit-root")?.focus({ preventScroll: true });
      getSelection().removeAllRanges(); getSelection().addRange(r);
      return true;
    }, r);
    if (!ok) return null;
    await page.keyboard.type("Q");
    return `pisze przy krawędzi linku (${r[1] < 0.5 ? "koniec" : "początek"})`;
  },
  async bookmark(page, r) {
    await page.keyboard.press(`${MOD}+Shift+F5`);
    if (!(await page.waitForSelector(".compose-pop .bm-name", { timeout: 1500 }).catch(() => null))) return "Ctrl/⌘+Shift+F5 (bez okienka)";
    if (r[0] < 0.2) {
      await page.waitForTimeout(150);
      const has = await page.evaluate(() => { const b = document.querySelector(".bm-list .lf-place"); b?.click(); return !!b; });
      if (has) { await page.click(".bm-delete"); if (await page.$(".compose-pop .bm-delete")) await page.click(".bm-delete").catch(() => {}); await idle(page); await closePops(page); return "usuwa zakładkę"; }
    }
    const name = ["Zakl_1", "Zakl_2", "Cel", "zła nazwa"][Math.floor(r[1] * 4)];
    await page.fill(".bm-name", name);
    if (await page.evaluate(() => document.querySelector(".bm-add").disabled)) { await closePops(page); return `zakładka „${name}” (odrzucona nazwa)`; }
    await page.click(".bm-add");
    return `zakładka „${name}”`;
  },
  async paste(page, r) {
    const p = PASTES[Math.floor(r[0] * PASTES.length)];
    await page.evaluate((p) => {
      const dt = new DataTransfer();
      dt.setData("text/plain", p.text);
      if (p.html) dt.setData("text/html", p.html);
      const n = getSelection().anchorNode;
      const target = (n?.nodeType === 1 ? n : n?.parentElement) || document.querySelector(".docx-edit-root");
      target.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
    }, p);
    return `wkleja ${p.kind}: „${p.text.replace(/\n/g, "⏎").slice(0, 30)}”`;
  },
  async pageBreak(page) { await page.keyboard.press(`${MOD}+Enter`); return "Ctrl/⌘+Enter (podział strony)"; },
  async table(page, r) { const rows = 1 + Math.floor(r[0] * 3); const cols = 1 + Math.floor(r[1] * 3); await page.evaluate(([a, b]) => composeUi.insertTable(a, b), [rows, cols]); return `tabela ${rows}×${cols}`; },
  async image(page) {
    await page.evaluate(async () => {
      const c = document.createElement("canvas"); c.width = 120; c.height = 60;
      const g = c.getContext("2d"); g.fillStyle = "#3a6ea5"; g.fillRect(0, 0, 120, 60);
      const blob = await new Promise((r) => c.toBlob(r, "image/png"));
      await composeUi.insertImageFile(new File([blob], "chaos.png", { type: "image/png" }));
    });
    return "wstawia obraz";
  },
  async footnote(page) {
    await page.keyboard.press(`${MOD}+Alt+f`);
    await idle(page);
    await page.keyboard.type("przypis");
    await page.keyboard.press("Escape");
    return "przypis (Ctrl/⌘+Alt+F) + tekst";
  },
  async undo(page) { await page.keyboard.press(`${MOD}+z`); return "Cofnij"; },
  async redo(page) { await page.keyboard.press(`${MOD}+Shift+z`); return "Ponów"; },
  async mode(page) {
    await page.evaluate(() => appFrame.setReadOnly(!readOnlyMode));
    await idle(page);
    const ro = await page.evaluate(() => readOnlyMode);
    return ro ? "Czytanie" : "Edycja";
  },
};

// ── jedna sesja: dokument startowy + działania; zwraca pierwszy błąd (albo null) ──
async function openDoc(browser, doc) {
  const context = await browser.newContext(MOBILE
    ? { serviceWorkers: "block", viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, isMobile: true, hasTouch: true }
    : { serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); sessionStorage.setItem("dwbAutoLinkTold", "1"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(`pageerror: ${e.message}`));
  page.on("console", (m) => { if (m.type() === "error" && !/Failed to load resource|favicon/.test(m.text())) errors.push(`console.error: ${m.text()}`); });
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  if (doc === "blank") {
    await page.evaluate(() => composeUi.createNew("note"));
  } else {
    await page.evaluate((d) => loadSampleDocument(d), doc);
  }
  await page.waitForSelector(".docx-preview-host p", { timeout: 30000 });
  await ensureEdit(page);
  await idle(page, 400);
  return { context, page, errors };
}

const docXml = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  return (await z.file("word/document.xml").async("string")).replace(/ w14:paraId="[^"]*"| w14:textId="[^"]*"/g, "");
});
const allText = (page) => page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => previewRunsToPlainText(extractRunsFromPreviewParagraph(p)).replace(/ /g, " ")).join("\n"));

async function runSession(browser, doc, actions, { final = true } = {}) {
  const { context, page, errors } = await openDoc(browser, doc);
  const log = [];
  let failure = null;
  try {
    const base = await page.evaluate(() => dwbSelfCheck().then((r) => r.problems.map((p) => `${p.code}|${p.msg}`)));
    const known = new Set(base);
    const x0 = STRICT_UNDO ? await docXml(page) : null;
    for (let k = 0; k < actions.length && !failure; k++) {
      const a = actions[k];
      let label = a.name;
      try {
        if (a.name !== "mode" && a.name !== "undo" && a.name !== "redo") await ensureEdit(page);
        label = (await ACT[a.name](page, a.r)) || `${a.name} (pominięte)`;
        await idle(page);
        await closePops(page);
      } catch (e) {
        failure = { step: k, code: "ACTION_THROW", msg: `${label}: ${String(e.message).split("\n")[0]}` };
      }
      log.push(label);
      if (VERBOSE) console.log(`   ${k + 1}. ${label}`);
      if (failure) break;
      if (errors.length) { failure = { step: k, code: "PAGE_ERROR", msg: errors.join(" | ").slice(0, 400) }; break; }
      const res = await page.evaluate(() => dwbSelfCheck().then((r) => r.problems)).catch((e) => [{ level: "error", code: "SELF_CHECK_THROW", msg: String(e.message || e) }]);
      const fresh = res.filter((p) => p.level === "error" && !known.has(`${p.code}|${p.msg}`));
      if (fresh.length) failure = { step: k, code: fresh[0].code, msg: fresh.map((p) => `${p.code}: ${p.msg}`).join(" | ").slice(0, 600) };
      res.forEach((p) => known.add(`${p.code}|${p.msg}`)); // ostrzeżenia i znane — bez powtórzeń
    }
    if (!failure && final) {
      // zapis → ponowne otwarcie: ten sam tekst
      await ensureEdit(page);

      // Cofnij wszystko = dokument z początku (cofanie przywraca migawki bajtów — dokładnie), potem
      // Ponów wszystko = najdalszy stan historii. Seria mogła skończyć się na „Cofnij”, więc
      // najpierw Ponów do końca — to jest stan odniesienia (nie stan po ostatnim kroku).
      let xN = null;
      if (STRICT_UNDO) {
        let t = 0;
        while (t < 120 && (await page.evaluate(() => dwbUndo.canRedo()))) { await page.evaluate(() => dwbUndo.redo()); await idle(page, 120); t++; }
        xN = await docXml(page);
        let n = 0;
        while (n < 120 && (await page.evaluate(() => dwbUndo.canUndo()))) { await page.evaluate(() => dwbUndo.undo()); await idle(page, 120); n++; }
        const xU = await docXml(page);
        if (xU !== x0) failure = { step: actions.length, code: "UNDO_ALL", msg: `Cofnij wszystko (${n}) ≠ dokument z początku: ${firstDiff(xU, x0)}` };
        else {
          let m = 0;
          while (m < 120 && (await page.evaluate(() => dwbUndo.canRedo()))) { await page.evaluate(() => dwbUndo.redo()); await idle(page, 120); m++; }
          const xR = await docXml(page);
          if (xR !== xN) failure = { step: actions.length, code: "REDO_ALL", msg: `Ponów wszystko (${m}) ≠ koniec historii: ${firstDiff(xR, xN)}` };
        }
      }
      if (!failure) {
        await ensureEdit(page);
        await idle(page);
        const before = await allText(page);
        const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
        await page.locator("#fileInput").setInputFiles({ name: "chaos.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: Buffer.from(b64, "base64") });
        await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
        await ensureEdit(page);
        await idle(page, 400);
        const after = await allText(page);
        if (after !== before) failure = { step: actions.length, code: "ROUNDTRIP", msg: `zapis → otwarcie zmienia tekst: ${firstDiff(after, before)}` };
      }
      if (!failure && errors.length) failure = { step: actions.length, code: "PAGE_ERROR", msg: errors.join(" | ").slice(0, 400) };
    }
  } catch (e) {
    failure = failure || { step: log.length, code: "HARNESS", msg: String(e.message).split("\n")[0] };
  }
  await context.close();
  return { failure, log };
}

function firstDiff(a, b) {
  let i = 0;
  while (i < a.length && a[i] === b[i]) i++;
  return `@${i}: „${a.slice(Math.max(0, i - 60), i + 80)}” ≠ „${b.slice(Math.max(0, i - 60), i + 80)}”`;
}

// ── skracanie serii (ddmin): najmniejsza seria z tym samym kodem błędu ──
async function shrink(browser, doc, actions, code) {
  const t0 = Date.now();
  let cur = actions;
  let n = 2;
  const fails = async (cand) => {
    if (Date.now() - t0 > SHRINK_MAX_MS) return false;
    const final = code === "UNDO_ALL" || code === "REDO_ALL" || code === "ROUNDTRIP";
    const { failure } = await runSession(browser, doc, cand, { final });
    return failure?.code === code;
  };
  while (cur.length >= 2 && Date.now() - t0 < SHRINK_MAX_MS) {
    const size = Math.ceil(cur.length / n);
    let reduced = false;
    for (let i = 0; i < cur.length; i += size) {
      const cand = [...cur.slice(0, i), ...cur.slice(i + size)];
      if (cand.length && (await fails(cand))) { cur = cand; n = Math.max(n - 1, 2); reduced = true; break; }
    }
    if (!reduced) { if (n >= cur.length) break; n = Math.min(cur.length, n * 2); }
  }
  return cur;
}

async function main() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const failures = [];
  // CASES=katalog — znalezione wcześniej (i naprawione) serie jako stałe testy regresji
  const caseDir = process.env.CASES ? path.resolve(process.cwd(), process.env.CASES) : null;
  const cases = process.env.REPLAY
    ? [JSON.parse(fs.readFileSync(process.env.REPLAY, "utf8"))]
    : caseDir
      ? fs.readdirSync(caseDir).filter((f) => f.endsWith(".json")).sort().map((f) => ({ ...JSON.parse(fs.readFileSync(path.join(caseDir, f), "utf8")), seed: f.replace(/\.json$/, "") }))
      : SEEDS.map((seed) => ({ seed, doc: process.env.DOC || (seed % 2 ? "przewodnik" : "blank"), actions: planActions(seed, STEPS) }));
  for (const c of cases) {
    const t0 = Date.now();
    const { failure, log } = await runSession(browser, c.doc, c.actions);
    const secs = ((Date.now() - t0) / 1000).toFixed(0);
    if (!failure) { console.log(`✅ ziarno ${c.seed} (${c.doc}, ${c.actions.length} kroków, ${secs} s)`); continue; }
    console.log(`❌ ziarno ${c.seed} (${c.doc}) — krok ${failure.step + 1}: ${failure.code}\n   ${failure.msg}`);
    let minimal = c.actions.slice(0, Math.min(c.actions.length, failure.step + 1));
    if (SHRINK && !process.env.REPLAY && !caseDir) {
      console.log(`   skracam serię ${minimal.length} kroków…`);
      minimal = await shrink(browser, c.doc, minimal, failure.code);
    }
    const rerun = await runSession(browser, c.doc, minimal, { final: ["UNDO_ALL", "REDO_ALL", "ROUNDTRIP"].includes(failure.code) });
    const rec = { engine: ENGINE, seed: c.seed, doc: c.doc, failure: rerun.failure || failure, steps: rerun.log, actions: minimal };
    fs.mkdirSync(OUT_DIR, { recursive: true });
    const file = path.join(OUT_DIR, `${ENGINE}-${c.seed}.json`);
    fs.writeFileSync(file, JSON.stringify(rec, null, 1));
    console.log(`   najkrótsza seria (${minimal.length}): ${rerun.log.map((s, i) => `${i + 1}. ${s}`).join(" → ")}`);
    console.log(`   ${rerun.failure ? `${rerun.failure.code}: ${rerun.failure.msg.slice(0, 300)}` : "(po skróceniu nie powtarza się — błąd zależny od czasu)"}`);
    console.log(`   powtórka: REPLAY=${path.relative(process.cwd(), file)}${ENGINE === "webkit" ? " ENGINE=webkit" : ""} node scripts/chaos-playwright.js`);
    failures.push(rec);
  }
  await browser.close();
  console.log(`\n${ENGINE}: ${cases.length - failures.length}/${cases.length} serii bez błędu`);
  process.exit(failures.length ? 1 : 0);
}

main().catch((e) => { console.error(e); process.exit(1); });
