// docx-roundtrip.js — prawdziwe pliki .docx: otwarcie → Edycja → dopisanie słowa → zapis →
// sprawdzenie, że NIC nie zginęło (2026-10-05, polowanie na błędy).
//
// Dla każdego pliku: (1) otwiera się bez błędów, (2) samo wejście w Edycję bez pisania nie
// zmienia pliku, (3) po dopisaniu słowa w jednym akapicie zapis jest poprawnym XML-em, tekst =
// stary tekst + słowo, a liczby rysunków, tabel, komentarzy, przypisów, pól, linków, zakładek
// i akapitów się zgadzają, (4) zapisany plik otwiera się ponownie bez błędów, (5) zapis nie dokłada
// błędów schematu Office (Open XML SDK, scripts/ooxml-validate.js — gdy jest .NET).
// Pliki NIE leżą w repo: lista ścieżek z pliku podanego w argumencie albo ~/.dwb-docx-samples.txt.
// Użycie: node scripts/docx-roundtrip.js [lista.txt]   |   ENGINE=webkit

const pw = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");
const ooxml = require("./ooxml-validate");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const LIST = process.argv[2] || path.join(os.homedir(), ".dwb-docx-samples.txt");
const MARK = " QX7";

async function run() {
  const files = fs.readFileSync(LIST, "utf8").split("\n").map((l) => l.trim()).filter((l) => l && !l.startsWith("#") && fs.existsSync(l));
  const browser = await pw[ENGINE].launch({ headless: true });
  const report = [];
  const schema = ooxml.available();
  if (!schema) console.log("ℹ️  bez sprawdzania schematu (brak .NET — brew install dotnet)");
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dwb-roundtrip-"));
  for (const file of files) {
    const name = path.basename(file);
    const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1300, height: 900 } });
    await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
    const page = await context.newPage();
    const errs = [];
    page.on("pageerror", (e) => errs.push("WYJĄTEK " + e.message));
    page.on("console", (m) => { if (m.type() === "error" && !/favicon|ResizeObserver|Failed to load resource/.test(m.text())) errs.push(m.text().slice(0, 200)); });
    page.on("dialog", (d) => d.accept().catch(() => {}));
    const r = { name, problems: [] };
    try {
      await page.goto(APP_URL, { waitUntil: "load" });
      await page.evaluate(() => document.getElementById("heroSplash")?.remove());
      await page.setInputFiles("#fileInput", file);
      await page.waitForFunction(() => originalFileBytes && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 60000 });
      await page.waitForTimeout(1500);
      const loadErr = await page.evaluate(() => [...document.querySelectorAll("#logList li, .log-entry")].map((e) => e.textContent).filter((t) => /error|błąd|Cannot|undefined/i.test(t)).slice(-3));
      if (!(await page.$(".docx-preview-host p"))) r.problems.push("NIE OTWIERA SIĘ (brak akapitów w podglądzie) " + loadErr.join(" | "));
      // stan pliku (liczby elementów, tekst)
      const stats = () => page.evaluate(async (bytesB64) => {
        const bytes = bytesB64 ? Uint8Array.from(atob(bytesB64), (c) => c.charCodeAt(0)) : await buildDocumentForSave();
        const z = await JSZip.loadAsync(bytes);
        const out = { bad: [] };
        for (const f of Object.keys(z.files)) {
          if (!/\.xml$|\.rels$/.test(f) || z.files[f].dir) continue;
          const x = (await z.file(f).async("string")).replace(/^\uFEFF/, ""); // BOM: Word go toleruje
          const d = new DOMParser().parseFromString(x, "application/xml");
          if (d.getElementsByTagName("parsererror").length) out.bad.push(f);
        }
        const doc = await z.file("word/document.xml").async("string");
        const d = new DOMParser().parseFromString(doc, "application/xml");
        const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const n = (tag, ns = W) => d.getElementsByTagNameNS(ns, tag).length;
        out.counts = { p: n("p"), tbl: n("tbl"), drawing: n("drawing"), pict: n("pict"), sdt: n("sdt"), hyperlink: n("hyperlink"), bookmarkStart: n("bookmarkStart"), commentRef: n("commentReference"), footnoteRef: n("footnoteReference"), fldChar: n("fldChar"), ins: n("ins"), del: n("del") };
        out.text = [...d.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join("");
        // nazwy krojów bez śmieci (dawniej „DM Sans"” z podglądu szło do pliku)
        out.badFonts = [...new Set([...d.getElementsByTagNameNS(W, "rFonts")].flatMap((f) => [...f.attributes].map((a) => a.value)).filter((v) => /["',;]/.test(v)))];
        out.parts = Object.keys(z.files).filter((f) => !z.files[f].dir).length;
        return out;
      }, null);
      const b64 = await page.evaluate(() => { let s = ""; originalFileBytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
      const before = await page.evaluate(async (b64) => {
        const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
        const z = await JSZip.loadAsync(bytes);
        const doc = await z.file("word/document.xml").async("string");
        const d = new DOMParser().parseFromString(doc, "application/xml");
        const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const n = (tag) => d.getElementsByTagNameNS(W, tag).length;
        return { counts: { p: n("p"), tbl: n("tbl"), drawing: n("drawing"), pict: n("pict"), sdt: n("sdt"), hyperlink: n("hyperlink"), bookmarkStart: n("bookmarkStart"), commentRef: n("commentReference"), footnoteRef: n("footnoteReference"), fldChar: n("fldChar"), ins: n("ins"), del: n("del") }, text: [...d.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join(""), parts: Object.keys(z.files).filter((f) => !z.files[f].dir).length };
      }, b64);
      // (2) Edycja bez pisania — plik bez zmian
      await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
      await page.waitForTimeout(1200);
      const same = await page.evaluate(async () => { inlineDirtyValid = false; const n = collectInlineParagraphEdits().length; inlineDirtyValid = true; return { n, same: (await buildDocumentForSave()) === originalFileBytes }; });
      if (same.n || !same.same) r.problems.push(`EDYCJA BEZ PISANIA ZMIENIA PLIK (akapitów „zmienionych”: ${same.n})`);
      // (3) dopisz słowo w środkowym edytowalnym akapicie z tekstem
      const target = await page.evaluate((mark) => {
        const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.isContentEditable && !p.dataset.lock && p.textContent.trim().length > 15);
        if (!ps.length) return null;
        const p = ps[Math.floor(ps.length / 2)];
        p.scrollIntoView({ block: "center" });
        placeCaret(p, p.textContent.length);
        return p.textContent.slice(-30);
      }, MARK);
      if (target != null) {
        await page.keyboard.type(MARK);
        await page.waitForTimeout(500);
        const after = await stats();
        if (after.bad.length) r.problems.push("ZEPSUTY XML po zapisie: " + after.bad.join(", "));
        if (after.badFonts.length) r.problems.push("ZEPSUTA NAZWA KROJU: " + after.badFonts.join(" | "));
        for (const [k, v] of Object.entries(before.counts)) if (after.counts[k] !== v) r.problems.push(`ZMIENIONA LICZBA ${k}: ${v} → ${after.counts[k]}`);
        if (after.parts !== before.parts) r.problems.push(`ZMIENIONA LICZBA części pliku: ${before.parts} → ${after.parts}`);
        const plain = after.text.replace(MARK, "");
        if (!after.text.includes(MARK.trim())) r.problems.push("DOPISANE SŁOWO NIE TRAFIŁO DO PLIKU");
        else if (plain !== before.text) {
          let i = 0;
          while (i < plain.length && plain[i] === before.text[i]) i++;
          r.problems.push(`TEKST ZMIENIONY POZA DOPISKIEM przy „${before.text.slice(Math.max(0, i - 20), i + 20)}” → „${plain.slice(Math.max(0, i - 20), i + 20)}”`);
        }
        // (5) schemat Office: tylko błędy, których nie było w oryginale
        if (schema) {
          const out = path.join(tmp, "zapis.docx");
          const b64 = await page.evaluate(async () => { const b = await buildDocumentForSave(); let s = ""; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000)); return btoa(s); });
          fs.writeFileSync(out, Buffer.from(b64, "base64"));
          const [vo, vs] = ooxml.validate([file, out]);
          const added = ooxml.newErrors(vo, vs);
          if (vs.error) r.problems.push("SCHEMAT: zapis nie otwiera się w Open XML SDK — " + vs.error);
          else if (added.length) r.problems.push(`SCHEMAT: ${added.length} nowych błędów — ${added.slice(0, 3).map(ooxml.short).join(" | ")}`);
        }
        // (4) zapisany plik otwiera się ponownie
        await page.evaluate(async () => {
          const bytes = await buildDocumentForSave();
          setDirtyState(false);
          window.__reopened = await ingestFile(new File([bytes], "ponownie.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
        });
        await page.waitForTimeout(1200);
        const ok = await page.evaluate(() => window.__reopened && !!document.querySelector(".docx-preview-host p") && document.querySelector(".docx-preview-host").textContent.includes("QX7"));
        if (!ok) r.problems.push("ZAPISANY PLIK NIE OTWIERA SIĘ PONOWNIE");
      } else r.note = "brak edytowalnego akapitu z tekstem";
    } catch (e) {
      r.problems.push("BŁĄD SKRYPTU: " + String(e.message).split("\n")[0]);
    }
    r.problems.push(...[...new Set(errs)].map((e) => "KONSOLA: " + e));
    report.push(r);
    console.log(`${r.problems.length ? "❌" : "✅"} ${name}${r.note ? ` (${r.note})` : ""}${r.problems.length ? "\n     " + r.problems.join("\n     ") : ""}`);
    await context.close();
  }
  await browser.close();
  fs.rmSync(tmp, { recursive: true, force: true });
  const bad = report.filter((r) => r.problems.length).length;
  console.log(`\n${bad ? "❌" : "✅"} Zapis prawdziwych plików [${ENGINE}]: ${report.length - bad}/${report.length} bez problemów`);
  process.exit(bad ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
