// paste-format-playwright.js — Kopiuj/Wytnij → Wklej w tym samym dokumencie z formatowaniem źródła,
// jak w Wordzie (domyślne „Zachowaj formatowanie źródła”). Oczekiwane wyniki sprawdzone w Wordzie
// (AppleScript, 2026-10-12) na tym samym pliku: nagłówek, dwa punkty listy, akapit z wcięciem, pusty.
// Zgłoszenie: kilka punktów listy wklejało się jako jeden sklejony akapit bez punktorów i kolorów.
//   node scripts/paste-format-playwright.js / ENGINE=webkit node scripts/paste-format-playwright.js

const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";

async function buildDocx() {
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(__dirname, "..", "docs", "samples", "przewodnik.docx")));
  const xml = await zip.file("word/document.xml").async("string");
  const head = xml.slice(0, xml.indexOf("<w:body>") + 8);
  const sect = xml.match(/<w:sectPr[\s\S]*<\/w:sectPr>/)[0];
  const t = (s, rpr = "") => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ""}<w:t xml:space="preserve">${s}</w:t></w:r>`;
  const list = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
  const body = [
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${t("Nagłówek", '<w:b/><w:color w:val="C00000"/><w:sz w:val="32"/>')}</w:p>`,
    `<w:p>${list}${t("Punkt A ", "<w:b/>")}${t("zwykły")}</w:p>`,
    `<w:p>${list}${t("Punkt B")}</w:p>`,
    `<w:p><w:pPr><w:ind w:left="1440"/></w:pPr>${t("Cel wklejenia tutaj")}</w:p>`,
    "<w:p/>",
  ].join("");
  zip.file("word/document.xml", `${head}${body}${sect}</w:body></w:document>`);
  const file = path.join(os.tmpdir(), `dwb-paste-format-${process.pid}.docx`);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer" }));
  return file;
}

// [tekst akapitu, punkt listy?, wcięcie 1440 (akapit „Cel”)?] — wynik Worda dla tych samych kroków.
// Spacje na styku („tutaj Punkt”, „Punkt A wklejenia”) dokłada Word (inteligentne wklejanie: całe
// słowa przy słowie); kawałek słowa („Pun”) — bez spacji.
const L = (text) => [text, true, false];
const N = (text) => [text, false, false];
const C = (text) => [text, false, true];
const HEAD = [["Nagłówek", false, false], L("Punkt A zwykły"), L("Punkt B")];
const CASES = [
  { name: "całe punkty → początek akapitu", copy: [1, 0, 3, 0], at: [3, 0], want: [...HEAD, L("Punkt A zwykły"), L("Punkt B"), C("Cel wklejenia tutaj"), N("")] },
  { name: "całe punkty → środek akapitu", copy: [1, 0, 3, 0], at: [3, 4], want: [...HEAD, L("Cel Punkt A zwykły"), L("Punkt B"), C("wklejenia tutaj"), N("")] },
  { name: "całe punkty → koniec akapitu", copy: [1, 0, 3, 0], at: [3, 19], want: [...HEAD, L("Cel wklejenia tutaj Punkt A zwykły"), L("Punkt B"), C(""), N("")] },
  { name: "całe punkty → pusty akapit", copy: [1, 0, 3, 0], at: [4, 0], want: [...HEAD, C("Cel wklejenia tutaj"), L("Punkt A zwykły"), L("Punkt B"), N("")] },
  { name: "od środka punktu do środka następnego → środek akapitu", copy: [1, 6, 2, 3], at: [3, 4], want: [...HEAD, L("Cel A zwykły"), C("Punwklejenia tutaj"), N("")] },
  { name: "kawałek tekstu → środek akapitu", copy: [1, 0, 1, 7], at: [3, 4], want: [...HEAD, C("Cel Punkt A wklejenia tutaj"), N("")], bold: "Punkt A" },
  // Wytnij: punkty listy przeniesione pod „Cel…” (z punktorami), na górze ich nie ma
  { name: "Wytnij całe punkty → pusty akapit (przeniesienie)", cut: true, copy: [1, 0, 3, 0], at: [2, 0], want: [["Nagłówek", false, false], C("Cel wklejenia tutaj"), L("Punkt A zwykły"), L("Punkt B"), N("")] },
  // wklejenie w miejsce zaznaczenia („wklejenia” zastąpione punktem)
  { name: "kawałek tekstu w miejsce zaznaczonego słowa", copy: [2, 0, 2, 7], at: [3, 4, 3, 13], want: [...HEAD, C("Cel Punkt B tutaj"), N("")] },
];

async function run() {
  const file = await buildDocx();
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const fails = [];
  const check = (name, ok, detail = "") => { if (!ok) fails.push(`${name} ${detail}`); console.log(`${ok ? "✓" : "✗"} ${name}${ok ? "" : ` ${detail}`}`); };

  for (const [k, c] of CASES.entries()) {
    // czysta strona na przypadek (niezapisany poprzedni dokument otwierałby się obok w karcie)
    if (k) {
      await page.goto(APP_URL, { waitUntil: "load" });
      await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    }
    await page.locator("#fileInput").setInputFiles(file);
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && collectPreviewParagraphElements(document.querySelector(".docx-preview-host") || document.body).some((p) => p.textContent === "Punkt B"), null, { timeout: 30000 });
    await page.evaluate(() => { if (readModeEl.checked) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change", { bubbles: true })); } });
    await page.waitForSelector(".docx-edit-root");
    await page.waitForTimeout(300);
    const got = await page.evaluate(async ({ copy, at, bold, cut }) => {
      const list = () => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
      const point = (pi, off) => {
        const p = list()[pi];
        const r = formDomRange(p, off, off);
        if (r) return [r.startContainer, r.startOffset];
        const e = document.createRange(); e.selectNodeContents(p); e.collapse(true); // pusty akapit
        return [e.startContainer, e.startOffset];
      };
      const r = document.createRange();
      r.setStart(...point(copy[0], copy[1]));
      r.setEnd(...point(copy[2], copy[3]));
      list()[copy[0]].closest(".docx-edit-root").focus();
      getSelection().removeAllRanges();
      getSelection().addRange(r);
      const clip = new DataTransfer();
      docEditRoot().dispatchEvent(new ClipboardEvent(cut ? "cut" : "copy", { clipboardData: clip, bubbles: true, cancelable: true }));
      if (cut) { await new Promise((res) => setTimeout(res, 400)); await waitInlineStructuralIdle(); }
      const q = document.createRange();
      q.setStart(...point(at[0], at[1]));
      if (at.length > 2) q.setEnd(...point(at[2], at[3]));
      getSelection().removeAllRanges();
      getSelection().addRange(q);
      docEditRoot().dispatchEvent(new ClipboardEvent("paste", { clipboardData: clip, bubbles: true, cancelable: true }));
      await new Promise((res) => setTimeout(res, 900));
      await waitInlineStructuralIdle();
      const zip = await JSZip.loadAsync(await buildDocumentForSave());
      const xml = await zip.file("word/document.xml").async("string");
      const ps = [...xml.matchAll(/<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((m) => m[0]);
      const rows = ps.map((p) => [[...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join(""), /<w:numPr>/.test(p), /<w:ind w:left="1440"/.test(p)]);
      let boldOk = null;
      if (bold) {
        const p = ps.find((x) => x.includes("Cel ")) || ""; // akapit, w który wklejono
        const runs = [...p.matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)].map((m) => [/<w:b\b(?![^>]*w:val="(0|false)")/.test(m[1]), [...m[1].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((x) => x[1]).join("")]);
        boldOk = runs.filter(([b]) => b).map(([, t]) => t).join("") === bold;
        if (!boldOk) boldOk = JSON.stringify(runs);
      }
      return { rows, boldOk, caret: getSelection().anchorNode?.textContent?.slice(0, 20) };
    }, c);
    check(`${c.name}: akapity i formatowanie jak w Wordzie`, JSON.stringify(got.rows) === JSON.stringify(c.want), `\n   jest:  ${JSON.stringify(got.rows)}\n   Word:  ${JSON.stringify(c.want)}`);
    if (c.bold) check(`${c.name}: pogrubione tylko „${c.bold}” (formatowanie źródła, nie miejsca)`, got.boldOk === true, JSON.stringify(got));
    if (k === 1) {
      // Cofnij: jedno naciśnięcie zdejmuje całe wklejenie
      await page.keyboard.press(ENGINE === "webkit" ? "Meta+KeyZ" : "Control+KeyZ");
      await page.waitForTimeout(1200);
      const undone = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => p.textContent));
      check("Cofnij po wklejeniu akapitów: jeden krok, dokument jak przed", JSON.stringify(undone) === JSON.stringify(["Nagłówek", "Punkt A zwykły", "Punkt B", "Cel wklejenia tutaj", ""]), JSON.stringify(undone));
    }
  }

  await browser.close();
  fs.rmSync(file, { force: true });
  if (errors.length) throw new Error(errors.join("\n"));
  if (fails.length) throw new Error(fails.join("\n"));
  console.log(`✅ paste-format-playwright passed [${ENGINE}]`);
}

run().catch((error) => {
  console.error("❌ paste-format-playwright failed:", error.message || error);
  process.exit(1);
});
