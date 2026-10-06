// Tabulatory w edytowanym DOCX są szerokimi elementami widoku. Klik obok nich musi ustawiać
// kursor po właściwej stronie, a Backspace/Delete muszą usunąć sam tabulator z zapisu Worda.

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const check = (name, ok, detail) => ({ name, ok: !!ok, detail });

async function fixture() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>A</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>B</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  return zip.generateAsync({ type: "nodebuffer" });
}

async function savedXml(page) {
  const b64 = await page.evaluate(async () => {
    const bytes = await buildDocumentForSave();
    let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return btoa(s);
  });
  return (await JSZip.loadAsync(Buffer.from(b64, "base64"))).file("word/document.xml").async("string");
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", { name: "tab.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await fixture() });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  await page.evaluate(() => { if (readOnlyMode) appFrame.setReadOnly(false); });
  const tab = page.locator(".docx-editable-p .docx-tab");
  await tab.waitFor();
  const box = await tab.boundingBox();
  if (!box) throw new Error("Brak widocznego tabulatora");

  const results = [];
  // Klik przy lewej/prawej krawędzi rozciągniętego tabulatora powinien trafić odpowiednio
  // przed i za znak, a nie przeskoczyć za literę B albo do środka sztucznej spacji.
  await page.mouse.click(box.x + 1, box.y + box.height / 2);
  const left = await page.evaluate(() => {
    const p = document.querySelector(".docx-editable-p");
    const s = getSelection();
    return { offset: getCaretOffset(p), inTab: !!(s.anchorNode?.parentElement?.closest?.(".docx-tab")) };
  });
  results.push(check("klik przy lewej krawędzi tabulatora = przed nim", left.offset === 1 && !left.inTab, JSON.stringify(left)));

  await page.mouse.click(box.x + box.width - 1, box.y + box.height / 2);
  const right = await page.evaluate(() => {
    const p = document.querySelector(".docx-editable-p");
    const s = getSelection();
    return { offset: getCaretOffset(p), inTab: !!(s.anchorNode?.parentElement?.closest?.(".docx-tab")) };
  });
  results.push(check("klik przy prawej krawędzi tabulatora = za nim", right.offset === 2 && !right.inTab, JSON.stringify(right)));

  await page.keyboard.press("Backspace");
  const xml = await savedXml(page);
  results.push(check("Backspace za tabulatorem usuwa go z DOCX", !/<w:tab\/>/.test(xml) && /<w:t>AB<\/w:t>/.test(xml), xml));

  // Druga strona tego samego zachowania: Delete przed tabulatorem usuwa go, nie literę B.
  await page.setInputFiles("#fileInput", { name: "tab.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await fixture() });
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  await page.evaluate(() => { if (readOnlyMode) appFrame.setReadOnly(false); });
  const tab2 = page.locator(".docx-editable-p .docx-tab");
  const box2 = await tab2.boundingBox();
  if (!box2) throw new Error("Brak tabulatora po ponownym otwarciu");
  await page.mouse.click(box2.x + 1, box2.y + box2.height / 2);
  await page.keyboard.press("Delete");
  const xmlDelete = await savedXml(page);
  results.push(check("Delete przed tabulatorem usuwa go z DOCX", !/<w:tab\/>/.test(xmlDelete) && /<w:t>AB<\/w:t>/.test(xmlDelete), xmlDelete));
  results.push(check("brak błędów strony", errors.length === 0, errors.join(" | ")));
  await browser.close();
  const bad = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`));
  if (bad.length) process.exit(1);
  console.log(`✅ tab-caret [${ENGINE}]`);
}

run().catch((e) => { console.error(e); process.exit(1); });
