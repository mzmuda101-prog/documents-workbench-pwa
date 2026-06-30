// run-styles-playwright.js — bold/color/font preserved on save roundtrip.

const fs = require("fs");
const path = require("path");
const {
  createTestPage,
  bootApp,
  assertNoErrors,
} = require("./docx-test-helpers");

const STYLED_DOCX = path.resolve(__dirname, "../docs/samples/styled-sample.docx");

async function ensureStyledFixture() {
  if (fs.existsSync(STYLED_DOCX)) return;
  const JSZip = require("jszip");
  const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr><w:t>Red bold</w:t></w:r>
      <w:r><w:rPr><w:i/><w:u w:val="single"/></w:rPr><w:t> italic underline</w:t></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t> mono</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`);
  zip.file("word/document.xml", DOCUMENT);
  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(STYLED_DOCX, buf);
}

async function run() {
  await ensureStyledFixture();
  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await page.locator("#fileInput").setInputFiles(STYLED_DOCX);
  await page.waitForSelector(".docx-preview-host p");

  const result = await page.evaluate(async () => {
    await ensureDocLibs(false);
    const p = document.querySelector(".docx-preview-host p");
    const previewRuns = extractRunsFromPreviewParagraph(p);
    if (!previewRuns.some((r) => r.bold && r.color)) {
      return { ok: false, step: "preview", msg: "Brak bold+color w podglądzie" };
    }

    const saved = await buildDocumentForSave();
    const xmlRuns = (await extractParagraphRunsFromDocx(saved))[0] || [];
    const hasBold = xmlRuns.some((r) => r.bold);
    const hasColor = xmlRuns.some((r) => r.color && r.color.includes("FF"));
    const hasItalic = xmlRuns.some((r) => r.italic);
    const hasUnderline = xmlRuns.some((r) => r.underline);
    const hasMono = xmlRuns.some((r) => (r.fontFamily || "").includes("Courier"));
    if (!hasBold || !hasColor || !hasItalic || !hasUnderline || !hasMono) {
      return { ok: false, step: "xml", msg: JSON.stringify(xmlRuns) };
    }

    document.getElementById("readMode").checked = false;
    document.getElementById("readMode").dispatchEvent(new Event("change", { bubbles: true }));
    await new Promise((r) => setTimeout(r, 200));
    p.focus();
    const sel = window.getSelection();
    const endRange = document.createRange();
    endRange.selectNodeContents(p);
    endRange.collapse(false);
    sel.removeAllRanges();
    sel.addRange(endRange);
    document.execCommand("insertText", false, " NOWY");
    const needle = " NOWY";
    let found = false;
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let textNode;
    while ((textNode = walker.nextNode())) {
      const idx = (textNode.textContent || "").indexOf(needle);
      if (idx < 0) continue;
      const fmtRange = document.createRange();
      fmtRange.setStart(textNode, idx);
      fmtRange.setEnd(textNode, idx + needle.length);
      sel.removeAllRanges();
      sel.addRange(fmtRange);
      found = true;
      break;
    }
    if (!found) return { ok: false, step: "select", msg: "Nie znaleziono NOWY" };
    document.execCommand("bold", false, null);
    document.execCommand("underline", false, null);
    p.dispatchEvent(new Event("input", { bubbles: true }));
    const domAfterEdit = extractRunsFromPreviewParagraph(p);
    const nowyRun = domAfterEdit.find((r) => (r.text || "").includes("NOWY"));
    if (!nowyRun?.bold || !nowyRun?.underline) {
      return { ok: false, step: "dom-edit", msg: JSON.stringify(domAfterEdit) };
    }
    const afterEdit = await buildDocumentForSave();
    const afterRuns = (await extractParagraphRunsFromDocx(afterEdit))[0] || [];
    const nowyXml = afterRuns.find((r) => (r.text || "").includes("NOWY"));
    if (!nowyXml?.bold) {
      return { ok: false, step: "edit-bold", msg: JSON.stringify(afterRuns) };
    }
    if (!nowyXml?.underline) {
      return { ok: false, step: "edit-underline", msg: JSON.stringify(afterRuns) };
    }

    return { ok: true, runs: afterRuns.length };
  });

  if (!result.ok) throw new Error(`run-styles failed at ${result.step}: ${result.msg}`);
  assertNoErrors(errors, "run-styles");
  await browser.close();
  console.log(`✅  run-styles-playwright passed (${result.runs} runs)`);
}

run().catch((err) => {
  console.error("❌  run-styles-playwright failed:", err.message || err);
  process.exit(1);
});
