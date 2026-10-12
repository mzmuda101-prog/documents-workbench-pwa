// paste-edge-playwright.js — Kopiuj/Wklej w trudnych miejscach, porównane z WORDEM (wzorce w
// paste-edge-expected.json: te same kroki wykonane w Wordzie dla Maca i zapisane jako .docx):
// słowo w czerwone pogrubione zdanie, punkty listy w nagłówek / komórkę tabeli, akapit w punkt
// numerowany, kawałek z komórki, tekst z przypisem (Word: nowy przypis), z obrazem (nowy numer
// obrazu), ramka „Wskazówka” tuż nad tabelą, kawałek na początek punktu. Do tego: Wytnij tekst
// z przypisem → przerysowanie → Wklej ×2, wklejenie do przypisu, zwykły tekst z zewnątrz.
// Każdy wynik przez walidator Open XML (gdy jest .NET).
//   node scripts/paste-edge-playwright.js / ENGINE=webkit node scripts/paste-edge-playwright.js

const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");
const { summarize } = require("./paste-edge-summary");
const validator = require("./ooxml-validate");
const { cases } = require("./paste-edge-expected.json");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";

async function buildDocx() {
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(__dirname, "..", "docs", "samples", "przewodnik.docx")));
  const xml = await zip.file("word/document.xml").async("string");
  const head = xml.slice(0, xml.indexOf("<w:body>") + 8);
  const sect = xml.match(/<w:sectPr[\s\S]*<\/w:sectPr>/)[0];
  const drawing = xml.match(/<w:r>(?:(?!<w:r>)[\s\S])*?<w:drawing>[\s\S]*?<\/w:drawing><\/w:r>/)[0];
  const t = (s, rpr = "") => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ""}<w:t xml:space="preserve">${s}</w:t></w:r>`;
  const pp = (inner) => `<w:pPr>${inner}</w:pPr>`;
  const cell = (s) => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr><w:p>${t(s)}</w:p></w:tc>`;
  const body = [
    `<w:p>${pp('<w:pStyle w:val="Heading1"/>')}${t("Rozdział pierwszy")}</w:p>`,
    `<w:p>${t("Ważne zdanie na czerwono", '<w:b/><w:color w:val="C00000"/>')}</w:p>`,
    `<w:p>${pp('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')}${t("Punkt A ", "<w:b/>")}${t("zwykły")}</w:p>`,
    `<w:p>${pp('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')}${t("Punkt B")}</w:p>`,
    `<w:p>${pp('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>')}${t("Numer jeden")}</w:p>`,
    `<w:p>${t("Tekst z przypisem")}<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>${t(" i dalej")}</w:p>`,
    `<w:p>${t("Obrazek ")}${drawing}${t(" w tekście")}</w:p>`,
    `<w:p>${pp('<w:pStyle w:val="Tip"/>')}${t("Wskazówka w ramce")}</w:p>`,
    '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>',
    `<w:tr>${cell("Komórka A1 tekst")}${cell("Komórka B1")}</w:tr>`,
    `<w:tr>${cell("Komórka A2")}${cell("Komórka B2")}</w:tr>`,
    "</w:tbl>",
    `<w:p>${t("Koniec dokumentu")}</w:p>`,
    "<w:p/>",
  ].join("");
  zip.file("word/document.xml", `${head}${body}${sect}</w:body></w:document>`);
  const file = path.join(os.tmpdir(), `dwb-paste-edge-${process.pid}.docx`);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer" }));
  return file;
}

async function run() {
  const file = await buildDocx();
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  const fails = [];
  const check = (name, ok, detail = "") => { if (!ok) fails.push(`${name} ${detail}`); console.log(`${ok ? "✓" : "✗"} ${name}${ok ? "" : ` ${detail}`}`); };
  const canValidate = validator.available();
  const validate = (b64, label) => {
    if (!canValidate) return;
    const f = path.join(os.tmpdir(), `dwb-paste-edge-out-${process.pid}.docx`);
    fs.writeFileSync(f, Buffer.from(b64, "base64"));
    const r = validator.validate([f])[0];
    fs.rmSync(f, { force: true });
    check(`${label}: plik poprawny (Open XML)`, r.ok, JSON.stringify(r.errors).slice(0, 400));
  };

  const open = async () => {
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.locator("#fileInput").setInputFiles(file);
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && collectPreviewParagraphElements(document.querySelector(".docx-preview-host") || document.body).some((p) => p.textContent.startsWith("Koniec")), null, { timeout: 30000 });
    await page.evaluate(() => { if (readModeEl.checked) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change", { bubbles: true })); } });
    await page.waitForSelector(".docx-edit-root");
    await page.waitForTimeout(400);
    await page.evaluate(() => {
      window.__find = (prefix) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).find((p) => p.textContent.startsWith(prefix));
      window.__at = (p, off) => {
        if (off === "end") { const e = document.createRange(); e.selectNodeContents(p); e.collapse(false); return [e.endContainer, e.endOffset]; }
        const r = formDomRange(p, off, off);
        if (r) return [r.startContainer, r.startOffset];
        const e = document.createRange(); e.selectNodeContents(p); e.collapse(true); return [e.startContainer, e.startOffset];
      };
      window.__sel = (a, b) => { const r = document.createRange(); r.setStart(...a); r.setEnd(...(b || a)); getSelection().removeAllRanges(); getSelection().addRange(r); };
      window.__event = (type, clip, target = docEditRoot()) => target.dispatchEvent(new ClipboardEvent(type, { clipboardData: clip, bubbles: true, cancelable: true }));
      window.__save = async () => {
        await new Promise((res) => setTimeout(res, 900));
        await waitInlineStructuralIdle();
        const bytes = await buildDocumentForSave();
        const z = await JSZip.loadAsync(bytes);
        let s = ""; bytes.forEach((c) => (s += String.fromCharCode(c)));
        return { b64: btoa(s), doc: await z.file("word/document.xml").async("string"), styles: await z.file("word/styles.xml").async("string"), notes: await z.file("word/footnotes.xml")?.async("string") };
      };
    });
  };

  // ── te same kroki co w Wordzie ──
  for (const c of cases) {
    await open();
    const out = await page.evaluate(async (c) => {
      const ps = __find(c.src[0]);
      const pe = __find(c.src[2]);
      ps.closest(".docx-edit-root").focus();
      const end = c.src[3] === "mark" ? __at(collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).indexOf(pe) + 1], 0) : __at(pe, c.src[3]);
      __sel(__at(ps, c.src[1]), end);
      const clip = new DataTransfer();
      __event("copy", clip);
      __sel(__at(__find(c.dst[0]), c.dst[1]));
      __event("paste", clip);
      return __save();
    }, c);
    const got = summarize(out.doc, out.styles, out.notes);
    const same = JSON.stringify(got) === JSON.stringify(c.word);
    const diff = same ? "" : got.map((l, i) => (l === c.word[i] ? null : `\n   app:  ${l}\n   Word: ${c.word[i]}`)).filter(Boolean).join("");
    check(`${c.id} ${c.name}: jak w Wordzie`, same, diff);
    validate(out.b64, c.id);
  }

  // ── X1: Wytnij tekst z przypisem → przerysowanie dokumentu → Wklej dwa razy ──
  await open();
  let r = await page.evaluate(async () => {
    const p = __find("Tekst z przypisem");
    p.closest(".docx-edit-root").focus();
    __sel(__at(p, 0), __at(p, "end"));
    const clip = new DataTransfer();
    __event("cut", clip);
    await new Promise((res) => setTimeout(res, 300));
    await mergeInlineEditsIntoBytes();
    await reloadFromBytes(originalFileBytes);
    await new Promise((res) => setTimeout(res, 800));
    __sel(__at(__find("Koniec"), 7));
    __event("paste", clip);
    await new Promise((res) => setTimeout(res, 1500));
    await waitInlineStructuralIdle();
    __sel(__at(__find("Koniec"), 0));
    __event("paste", clip);
    return __save();
  });
  const notesX1 = [...(r.notes || "").matchAll(/<w:footnote\b[^>]*w:id="(\d+)"[^>]*>([\s\S]*?)<\/w:footnote>/g)].filter((m) => +m[1] > 0).map((m) => [...m[2].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]).join("").trim());
  const refsX1 = [...r.doc.matchAll(/footnoteReference w:id="(\d+)"/g)].map((m) => m[1]);
  check("X1 Wytnij z przypisem → przerysowanie → Wklej ×2: dwa przypisy z treścią, różne numery", notesX1.length === 2 && notesX1.every((t) => t && t === notesX1[0]) && new Set(refsX1).size === 2, JSON.stringify({ notesX1, refsX1 }));
  validate(r.b64, "X1");

  // ── X2: wklejenie do przypisu ──
  await open();
  r = await page.evaluate(async () => {
    const src = __find("Punkt A");
    src.closest(".docx-edit-root").focus();
    __sel(__at(src, 8), __at(src, 14));
    const clip = new DataTransfer();
    __event("copy", clip);
    const np = [...document.querySelectorAll("ol.dwb-notes li p")].find((p) => p.isContentEditable);
    const e = document.createRange(); e.selectNodeContents(np); e.collapse(false);
    getSelection().removeAllRanges(); getSelection().addRange(e);
    __event("paste", clip, np);
    return __save();
  });
  check("X2 wklejenie do przypisu: tekst w przypisie", /przypis dolny\.\s?zwykły/.test((r.notes || "").replace(/<[^>]+>/g, "")), (r.notes || "").replace(/<[^>]+>/g, "").slice(0, 200));
  validate(r.b64, "X2");

  // ── X3: zwykły tekst z zewnątrz → formatowanie miejsca (Word: zwykły tekst bierze format otoczenia) ──
  await open();
  r = await page.evaluate(async () => {
    const p = __find("Ważne");
    p.closest(".docx-edit-root").focus();
    __sel(__at(p, 6));
    const clip = new DataTransfer();
    clip.setData("text/plain", "WKLEJONE ");
    __event("paste", clip);
    return __save();
  });
  const line = summarize(r.doc, r.styles, r.notes).find((l) => l.includes("WKLEJONE")) || "";
  check("X3 zwykły tekst z zewnątrz w czerwonym pogrubionym zdaniu: czerwony pogrubiony", /Ważne WKLEJONE zdanie na czerwono<BC00000>/.test(line), line);
  validate(r.b64, "X3");

  await browser.close();
  fs.rmSync(file, { force: true });
  if (errors.length) throw new Error(errors.join("\n"));
  if (fails.length) throw new Error(fails.join("\n"));
  console.log(`✅ paste-edge-playwright passed [${ENGINE}]`);
}

run().catch((error) => {
  console.error("❌ paste-edge-playwright failed:", error.message || error);
  process.exit(1);
});
