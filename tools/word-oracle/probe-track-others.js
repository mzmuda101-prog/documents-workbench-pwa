// probe-track-others.js — co robi PRAWDZIWY Word, gdy przy włączonym śledzeniu (autor „Tester”) pisze się
// lub kasuje w akapicie z cudzymi zmianami (review-sample.docx, akapit „Termin płatności wynosi [del 14][ins 30] dni.”
// i „Zleceniodawca [ins niezwłocznie ]zgłasza usterki.”). node tools/word-oracle/probe-track-others.js
const fs = require("fs");
const path = require("path");
const word = require("./word");

const orig = fs.readFileSync(path.join(__dirname, "../../docs/samples/review-sample.docx"));
const para = (xml, start) => (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).find((p) => p.replace(/<[^>]+>/g, "").startsWith(start)) || "";
const short = (p) => p.replace(/ w:(rsid\w*|date|id)="[^"]*"/g, "").replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "").replace(/ xml:space="preserve"/g, "");
// pozycje w treści Worda liczą też tekst usunięty: „Termin płatności wynosi ” 24 + „14” + „30” + „ dni.”
const CASES = [
  ["pisanie we wstawce Anny (3|0 → 350)", "Termin", 27, 27, "5"],
  ["kasowanie znaku we wstawce Anny (30 → 3)", "Termin", 27, 28, ""],
  ["kasowanie całej wstawki Anny", "Termin", 26, 28, ""],
  ["kasowanie zwykłego tekstu obok usunięcia („wynosi ”)", "Termin", 17, 24, ""],
  ["zamiana słowa we wstawce Jana (niezwłocznie → szybko)", "Zleceniodawca", 14, 26, "szybko"],
  ["pisanie tuż przed wstawką Jana", "Zleceniodawca", 14, 14, "X"],
];
if (!process.env.APP) for (const [name, start, a, b, text] of CASES) {
  const pi = start === "Termin" ? 2 : +(process.env.PI || 4);
  const { res, parts } = word.saveDocx(orig, `
    tell application "Microsoft Word" to set user name to "Tester"
    set track revisions to true
    set ps to start of content of text object of paragraph ${pi}
    set res to content of text object of paragraph ${pi}
    select (create range it start (ps + ${a}) end (ps + ${b}))
    ${text ? `type text (get selection of application "Microsoft Word") text "${text}"` : `type backspace (get selection of application "Microsoft Word")`}`);
  console.log(`── ${name}\n   akapit Worda: ${JSON.stringify(res)}\n   ${short(para(parts["word/document.xml"] || "", start))}\n`);
}

// ── to samo z aplikacji: plik zapisany przy włączonym śledzeniu — Word go czyta, przyjmuje i odrzuca?
if (process.env.APP) (async () => {
  const { createTestPage, bootApp } = require("../../scripts/docx-test-helpers");
  const { browser, page } = await createTestPage();
  await bootApp(page);
  await page.locator("#fileInput").setInputFiles("docs/samples/review-sample.docx");
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 });
  await page.evaluate(async () => { localStorage.setItem("dwb.authorName", "Tester"); appFrame.setReadOnly(false); });
  await page.waitForTimeout(800);
  await page.evaluate(() => dwbTrack.start());
  await page.waitForTimeout(800);
  const sel = (start, a, b) => page.evaluate(([start, a, b]) => {
    const p = [...document.querySelectorAll(".docx-preview-host p")].find((x) => x.textContent.startsWith(start));
    const at = (n) => { const tw = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); let t, pos = 0; while ((t = tw.nextNode())) { if (pos + t.data.length >= n) return [t, n - pos]; pos += t.data.length; } };
    const r = document.createRange(); r.setStart(...at(a)); r.setEnd(...at(b));
    p.closest(".docx-edit-root").focus(); getSelection().removeAllRanges(); getSelection().addRange(r);
  }, [start, a, b]);
  await sel("Termin", 17, 24); await page.keyboard.press("Backspace"); await page.waitForTimeout(600);
  await sel("Termin", 18, 18); await page.keyboard.type("5"); await page.waitForTimeout(600);
  await sel("Zleceniodawca", 14, 26); await page.keyboard.type("szybko"); await page.waitForTimeout(800);
  const b64 = await page.evaluate(async () => { const bytes = await dwbTrack.bytesForDisk(await buildDocumentForSave()); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
  await browser.close();
  const bytes = Buffer.from(b64, "base64");
  const LIST = `repeat with i from 1 to count of revisions
      set rv to revision i
      set res to res & (revision type of rv as string) & "|" & (author of rv) & "|" & (content of text object of rv) & linefeed
    end repeat`;
  console.log("── plik z aplikacji: zmiany w Wordzie\n" + word.withDocument(bytes, LIST));
  console.log("── przyjmij wszystkie\n" + word.withDocument(bytes, `accept all revisions
    set res to (content of text object of paragraph 2) & "#" & (content of text object of paragraph 4)`));
  console.log("── odrzuć wszystkie\n" + word.withDocument(bytes, `reject all revisions
    set res to (content of text object of paragraph 2) & "#" & (content of text object of paragraph 4)`));
})();
