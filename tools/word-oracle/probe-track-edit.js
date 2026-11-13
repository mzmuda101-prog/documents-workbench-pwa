// probe-track-edit.js — Word czyta plik po pisaniu w akapicie ze śledzonymi zmianami tak, jak aplikacja?
// Zmiany (liczba, typ, autor, tekst) w oryginale i po edycji w aplikacji (wpisane „5” we wstawce „30”,
// „ roboczych” na końcu). Uruchom: node tools/word-oracle/probe-track-edit.js
const fs = require("fs");
const path = require("path");
const word = require("./word");
const { createTestPage, bootApp, loadDocxFile } = require("../../scripts/docx-test-helpers");

const REV = `
  set n to count of revisions
  repeat with i from 1 to n
    set rv to revision i
    set res to res & (revision type of rv as string) & "|" & (author of rv) & "|" & (content of text object of rv) & linefeed
  end repeat
  set res to res & "##" & (content of text object of paragraph 2)`;

(async () => {
  const orig = fs.readFileSync(path.join(__dirname, "../../docs/samples/review-sample.docx"));
  const { browser, page } = await createTestPage();
  await bootApp(page);
  await loadDocxFile(page, "docs/samples/review-sample.docx");
  await page.evaluate(() => appFrame.setReadOnly(false));
  await page.waitForTimeout(1500);
  await page.evaluate(() => { const p = [...document.querySelectorAll(".docx-preview-host p")].find((x) => x.textContent.startsWith("Termin")); const w = p.querySelector("[data-dwb-rev]"); const r = document.createRange(); r.setStart(w.firstChild, 1); r.collapse(true); p.closest(".docx-edit-root").focus(); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await page.keyboard.type("5");
  await page.keyboard.press("End");
  await page.keyboard.type(" roboczych");
  await page.waitForTimeout(800);
  const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
  await browser.close();
  const edited = Buffer.from(b64, "base64");
  console.log("── oryginał\n" + word.withDocument(orig, REV));
  console.log("── po edycji w aplikacji\n" + word.withDocument(edited, REV));
})();
