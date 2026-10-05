// undo-fuzz-playwright.js — Cofnij/Ponów na losowej serii operacji (2026-10-05, polowanie na błędy r2).
//
// Przewodnik → N losowych operacji (ziarno = powtarzalne): pisanie, Enter w środku akapitu,
// Backspace sklejający akapity, pogrubienie słowa, styl akapitu, lista, A+, wklejenie tekstu.
// Potem Cofnij do skutku → document.xml MA BYĆ taki jak po otwarciu; Ponów do skutku → taki jak
// po ostatniej operacji. Do tego: bez błędów strony, Cofnij nie „gubi” kursora poza dokument.
// Użycie: node scripts/undo-fuzz-playwright.js   |   ENGINE=webkit   |   SEEDS=1,2,3   |   OPS=14

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const SEEDS = (process.env.SEEDS || "1,2,3").split(",").map(Number);
const OPS = Number(process.env.OPS || 12);
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 30000 }).then(() => page.waitForTimeout(300));

const docXml = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  return (await z.file("word/document.xml").async("string")).replace(/ w14:paraId="[^"]*"| w14:textId="[^"]*"/g, "");
});
const firstDiff = (a, b) => { let i = 0; while (i < a.length && a[i] === b[i]) i++; return `@${i}: „${a.slice(Math.max(0, i - 120), i + 160)}” ≠ „${b.slice(Math.max(0, i - 120), i + 160)}”`; };
const plain = (xml) => (xml.match(/<w:t[^>]*>[^<]*/g) || []).map((x) => x.replace(/<w:t[^>]*>/, "")).join("").slice(0, 120);

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  for (const seed of SEEDS) {
    let s = seed * 9973;
    const rnd = () => ((s = (s * 16807) % 2147483647) / 2147483647);
    const pick = (a) => a[Math.floor(rnd() * a.length)];
    const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1300, height: 900 } });
    await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
    const page = await context.newPage();
    const errors = [];
    page.on("pageerror", (e) => errors.push(e.message));
    page.on("dialog", (d) => d.accept());
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.evaluate(() => loadSampleDocument("przewodnik"));
    await page.waitForSelector(".docx-preview-host p", { timeout: 30000 });
    await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
    await idle(page);
    const x0 = await docXml(page);
    if (seed === SEEDS[0]) {
      // stały scenariusz (zgłoszenie z polowania): Enter z kursorem W ŚRODKU pola formularza —
      // dawniej rozcinał kontrolkę, numeracja pól się przesuwała i następne akapity dostawały cudze pola
      const fields = () => page.evaluate(async () => { const z = await JSZip.loadAsync(await buildDocumentForSave()); const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml"); const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"; return { sdt: d.getElementsByTagNameNS(W, "sdt").length, texts: collectParagraphElements(d.documentElement, "all").map((x) => [...x.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join("")).filter((t) => /^(Dział|Data rozpoczęcia|☐ Akceptuję)/.test(t)) }; });
      const f0 = await fields();
      await page.evaluate(() => { const el = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((x) => x.textContent.startsWith("Imię i nazwisko")); el.scrollIntoView({ block: "center" }); placeCaret(el, 21); });
      await page.keyboard.press("Enter");
      await idle(page);
      await page.waitForTimeout(600);
      const f1 = await fields();
      check("Enter z kursorem w polu formularza: tyle samo pól, następne akapity z własnymi polami", f1.sdt === f0.sdt && JSON.stringify(f1.texts) === JSON.stringify(f0.texts), JSON.stringify({ f0, f1 }));
      await page.evaluate(() => dwbUndo.undo());
      await idle(page);
      check("…i Cofnij przywraca dokument dokładnie", (await docXml(page)) === x0, "");
    }

    // akapit edytowalny z tekstem (losowy)
    const target = () => page.evaluate((r) => {
      const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.isContentEditable && !p.dataset.lock && p.textContent.trim().length > 25 && !p.closest("td"));
      const p = ps[Math.floor(r * ps.length)];
      p.scrollIntoView({ block: "center" });
      return resolveParaIndex(p);
    }, rnd());
    const caretAt = (i, where) => page.evaluate(([i, where]) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
      const n = p.textContent.length;
      placeCaret(p, where === "end" ? n : where === "start" ? 0 : Math.floor(n / 2));
    }, [i, where]);
    const selectWord = (i) => page.evaluate((i) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
      const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      for (let n = w.nextNode(); n; n = w.nextNode()) {
        const m = n.data.match(/\p{L}{4,}/u);
        if (m && !n.parentElement.closest('[contenteditable="false"]')) { const r = document.createRange(); r.setStart(n, m.index); r.setEnd(n, m.index + m[0].length); const s = getSelection(); s.removeAllRanges(); s.addRange(r); return m[0]; }
      }
      return null;
    }, i);
    const done = [];
    const ops = {
      type: async () => { const i = await target(); await caretAt(i, "end"); await page.keyboard.type(" dopisek" + Math.floor(rnd() * 100)); return `pisanie w ${i}`; },
      enter: async () => { const i = await target(); await caretAt(i, "mid"); await page.keyboard.press("Enter"); return `Enter w ${i}`; },
      backspace: async () => { const i = await target(); await caretAt(i, "start"); await page.keyboard.press("Backspace"); return `Backspace na początku ${i}`; },
      bold: async () => { const i = await target(); const w = await selectWord(i); if (!w) return null; await page.keyboard.press(`${MOD}+b`); return `pogrubienie „${w}”`; },
      style: async () => { const i = await target(); await caretAt(i, "mid"); const st = pick(["h2", "normal", "callout", "quote"]); await page.selectOption("#fmtParaStyle", st); return `styl ${st} w ${i}`; },
      list: async () => { const i = await target(); await caretAt(i, "mid"); await page.evaluate(() => composeUi.applyList("bullet")); return `lista w ${i}`; },
      grow: async () => { const i = await target(); const w = await selectWord(i); if (!w) return null; await page.click("#fmtGrow"); return `A+ „${w}”`; },
      paste: async () => {
        const i = await target(); await caretAt(i, "mid");
        await page.evaluate(() => { const dt = new DataTransfer(); dt.setData("text/plain", "WKLEJONE zdanie"); docCanvasEl.querySelector(".docx-edit-root, .docx-preview-host")?.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true })); });
        return `wklejenie w ${i}`;
      },
    };
    const names = Object.keys(ops);
    for (let k = 0; k < OPS; k++) {
      const name = pick(names);
      try {
        const label = await ops[name]();
        if (!label) continue;
        await idle(page);
        await page.waitForTimeout(1700); // przerwa > 1,5 s zamyka krok pisania
        done.push(label);
        if (process.env.STEPCHECK) { // od razu Cofnij + Ponów — która operacja psuje
          const before = await docXml(page);
          await page.evaluate(() => dwbUndo.undo()); await idle(page);
          await page.evaluate(() => dwbUndo.redo()); await idle(page);
          const again = await docXml(page);
          if (again !== before) console.log(`   ⚠️  po „${label}”: Cofnij+Ponów zmienia dokument ${firstDiff(again, before).slice(0, 700)}`);
        }
      } catch (e) {
        check(`ziarno ${seed}: operacja ${name}`, false, String(e.message).split("\n")[0]);
      }
    }
    const xN = await docXml(page);
    // Cofnij do skutku
    let undos = 0;
    while (undos < 80 && (await page.evaluate(() => dwbUndo.canUndo()))) {
      await page.evaluate(() => dwbUndo.undo());
      await idle(page);
      undos++;
    }
    const xU = await docXml(page);
    check(`ziarno ${seed}: Cofnij wszystko (${undos} kroków po ${done.length} operacjach) = dokument jak po otwarciu`, xU === x0, xU === x0 ? "" : `${firstDiff(xU, x0)} | operacje: ${done.join("; ")}`);
    let redos = 0;
    while (redos < 80 && (await page.evaluate(() => dwbUndo.canRedo()))) {
      await page.evaluate(() => dwbUndo.redo());
      await idle(page);
      redos++;
    }
    const xR = await docXml(page);
    check(`ziarno ${seed}: Ponów wszystko (${redos}) = dokument jak po ostatniej operacji`, xR === xN, xR === xN ? "" : `${firstDiff(xR, xN)} | operacje: ${done.join("; ")}`);
    const real = errors.filter((e) => !/ResizeObserver/.test(e));
    if (real.length) check(`ziarno ${seed}: bez błędów strony`, false, real.join(" | ").slice(0, 300));
    await context.close();
  }
  await browser.close();
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Cofnij/Ponów na losowych operacjach [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
