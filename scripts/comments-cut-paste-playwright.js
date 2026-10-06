// comments-cut-paste-playwright.js — wycięty cały komentowany fragment ma przenieść kotwicę
// komentarza razem z tekstem, zamiast zostawić po wklejeniu sam tekst.

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1120, height: 800 } });
  await context.addInitScript(() => {
    sessionStorage.setItem("introPlayed", "true");
    localStorage.setItem("dwb.authorName", "Test komentarza");
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });

  await page.keyboard.type("Początek zaznaczony fragment koniec");
  await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0];
    const from = p.textContent.indexOf("zaznaczony");
    const r = formDomRange(p, from, from + "zaznaczony fragment".length);
    p.closest(".docx-edit-root").focus();
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  });
  await page.keyboard.press(`${MOD}+Alt+KeyM`);
  await page.waitForSelector(".compose-pop-comment .cf-rich");
  await page.keyboard.type("Przenoszony komentarz");
  await page.click(".compose-pop-form .lf-ok");
  await page.waitForTimeout(500);

  const result = await page.evaluate(async () => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0];
    const from = p.textContent.indexOf("zaznaczony");
    const r = formDomRange(p, from, from + "zaznaczony fragment".length);
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    const clip = new DataTransfer();
    docEditRoot().dispatchEvent(new ClipboardEvent("cut", { clipboardData: clip, bubbles: true, cancelable: true }));
    docEditRoot().dispatchEvent(new ClipboardEvent("paste", { clipboardData: clip, bubbles: true, cancelable: true }));
    const bytes = await buildDocumentForSave();
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("word/document.xml").async("string");
    const comments = await zip.file("word/comments.xml").async("string");
    const ranges = [...document.querySelectorAll("span[data-cm-kind='start']")].map((start) => {
      const end = document.querySelector(`span[data-cm-kind="end"][data-cm-id="${CSS.escape(start.dataset.cmId)}"]`);
      if (!end) return { id: start.dataset.cmId, text: "" };
      const range = document.createRange();
      range.setStartAfter(start);
      range.setEndBefore(end);
      return { id: start.dataset.cmId, text: range.toString() };
    });
    const highlight = CSS.highlights?.get("dwb-comment");
    return {
      text: p.textContent,
      starts: (xml.match(/<w:commentRangeStart\b/g) || []).length,
      ends: (xml.match(/<w:commentRangeEnd\b/g) || []).length,
      refs: (xml.match(/<w:commentReference\b/g) || []).length,
      comment: comments.includes("Przenoszony komentarz"),
      ranges,
      highlightRanges: highlight ? highlight.size : null,
    };
  });

  await browser.close();
  if (errors.length) throw new Error(errors.join("\n"));
  const moved = result.ranges.length === 1 && result.ranges[0].text === "zaznaczony fragment";
  if (result.starts !== 1 || result.ends !== 1 || result.refs !== 1 || !result.comment || !moved
    || (result.highlightRanges !== null && result.highlightRanges < 1)) {
    throw new Error(`Komentarz zgubiony po Wytnij/Wklej: ${JSON.stringify(result)}`);
  }
  console.log(`✅ comments-cut-paste-playwright passed [${ENGINE}]`);
}

run().catch((error) => {
  console.error("❌ comments-cut-paste-playwright failed:", error.message || error);
  process.exit(1);
});
