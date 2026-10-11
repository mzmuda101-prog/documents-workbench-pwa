// empty-after-table-playwright.js — pusty akapit za tabelą (zgłoszenie 2026-10-11: „pusta strona
// na końcu”, Backspace mówił „nie da się skleić”). Tabela wypełnia stronę, a za nią plik Worda
// musi mieć akapit — spadał sam na nową kartkę.
//   • ostatni akapit dokumentu → Backspace zmniejsza go do 1 pt (pusta strona znika), Cofnij wraca
//   • pusty akapit za tabelą, a dalej tekst → Backspace go usuwa, kursor na końcu tabeli

const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";

// review-sample.docx z podmienioną treścią: tytuł, tabela rows×2, pusty akapit (+ opcjonalnie tekst)
async function buildDocx(rows, withTail) {
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(__dirname, "..", "docs", "samples", "review-sample.docx")));
  const xml = await zip.file("word/document.xml").async("string");
  const head = xml.slice(0, xml.indexOf("<w:body>") + 8);
  const sect = xml.match(/<w:sectPr[\s\S]*<\/w:sectPr>/)[0];
  const para = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  const cell = (t) => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${para(t)}</w:tc>`;
  let body = para("Tytuł dokumentu z tabelą na końcu");
  body += '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>';
  for (let i = 1; i <= rows; i++) body += `<w:tr>${cell(`Wiersz ${i}`)}${cell(`Wartość ${i}`)}</w:tr>`;
  body += "</w:tbl><w:p/>";
  if (withTail) body += para("Tekst po pustym akapicie");
  zip.file("word/document.xml", `${head}${body}${sect}</w:body></w:document>`);
  const file = path.join(os.tmpdir(), `dwb-empty-after-table-${rows}-${withTail ? 1 : 0}-${process.pid}.docx`);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer" }));
  return file;
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 1000 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const fails = [];
  const check = (name, ok, detail = "") => { if (!ok) fails.push(`${name} ${detail}`); console.log(`${ok ? "✓" : "✗"} ${name}${ok ? "" : ` ${detail}`}`); };

  const open = async (file) => {
    await page.locator("#fileInput").setInputFiles(file);
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !!document.querySelector(".docx-preview-host"), null, { timeout: 30000 });
    await page.evaluate(() => {
      if (!readModeEl.checked) return;
      readModeEl.checked = false;
      readModeEl.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await page.waitForSelector(".docx-edit-root");
    await page.waitForTimeout(800);
  };
  const backspaceAt = async (fromEnd) => {
    await page.evaluate((i) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).at(i);
      p.closest(".docx-edit-root").focus();
      placeCaret(p, 0);
    }, fromEnd);
    await page.keyboard.press("Backspace");
    await page.waitForTimeout(1500);
  };
  const state = () => page.evaluate(async () => {
    const h = document.querySelector(".docx-preview-host");
    const list = collectPreviewParagraphElements(h);
    const xml = await (await JSZip.loadAsync(await buildDocumentForSave())).file("word/document.xml").async("string");
    const anchor = getSelection().anchorNode;
    return {
      gaps: h.querySelectorAll(".dwb-page-gap").length,
      paras: list.length,
      lastH: list.at(-1).getBoundingClientRect().height,
      lastText: list.at(-1).textContent,
      afterTable: xml.slice(xml.lastIndexOf("</w:tbl>"), xml.lastIndexOf("</w:tbl>") + 160),
      caretInTable: !!(anchor?.nodeType === 1 ? anchor : anchor?.parentElement)?.closest("td, th"),
    };
  });

  // ── pusta strona na końcu: tabela do dołu strony, za nią pusty ostatni akapit ──
  let file = null;
  let before = null;
  for (const rows of [47, 48, 49, 50, 51, 52, 46, 45]) { // granica strony zależy od silnika (czcionki)
    file = await buildDocx(rows, false);
    await open(file);
    before = await page.evaluate(() => {
      const list = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
      let prev = list.at(-1).previousElementSibling;
      return { onOwnPage: !!prev?.matches(".dwb-page-gap, .dwb-page-gap-spacer"), paras: list.length };
    });
    if (before.onOwnPage) break;
  }
  check("odtworzenie: pusty ostatni akapit sam na nowej stronie", before.onOwnPage);
  await backspaceAt(-1);
  const shrunk = await state();
  check("Backspace: akapit 1 pt bez odstępów w pliku", /<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"\/>/.test(shrunk.afterTable), shrunk.afterTable);
  check("…pusta strona znika (bez granicy przed nim), akapit zostaje", shrunk.gaps === 0 && shrunk.lastH < 4 && shrunk.paras === before.paras, JSON.stringify(shrunk));
  check("…bez komunikatu „nie da się skleić”", await page.evaluate(() => ![...document.querySelectorAll("[class*='toast']")].some((t) => /skleić|joined/.test(t.textContent))));
  // Zmniejszony akapit ma 1 pt — kliknięcie pod tabelą musi w niego trafić, a pisanie przywraca
  // mu zwykłą wysokość (zgłoszenie: „ciężko by było w to trafić palcem, kursorem”).
  const below = await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).at(-1);
    p.scrollIntoView({ block: "center" });
    const r = p.getBoundingClientRect();
    return { x: r.left + 40, y: r.bottom + 30 };
  });
  await page.waitForTimeout(300);
  await page.mouse.click(below.x, below.y);
  await page.waitForTimeout(300);
  const caret = await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).at(-1);
    const a = getSelection().anchorNode;
    return { inside: p.contains(a) || p === a, active: p.classList.contains("dwb-tiny-active"), bar: getComputedStyle(p, "::after").height };
  });
  check("klik 30 px pod tabelą: kursor w akapicie 1 pt, widoczna kreska kursora", caret.inside && caret.active && caret.bar === "15px", JSON.stringify(caret));
  await page.keyboard.type("Dalej");
  await page.waitForTimeout(1500);
  const typed = await state();
  check("…pisanie przywraca zwykły akapit (w pliku bez 1 pt, tekst jest)", typed.lastText === "Dalej" && typed.lastH > 10 && !/w:lineRule="exact"/.test(typed.afterTable) && /Dalej/.test(typed.afterTable), JSON.stringify(typed));
  for (let i = 0; i < 2; i++) {
    await page.keyboard.press(`${MOD}+KeyZ`);
    await page.waitForTimeout(1200);
  }
  const undone = await state();
  check("Cofnij (pisanie, potem zmniejszenie) przywraca pusty zwykły akapit", !/w:lineRule="exact"/.test(undone.afterTable) && undone.lastH > 10 && !undone.lastText, JSON.stringify(undone));

  // ── pusty akapit za tabelą, a dalej tekst ──
  file = await buildDocx(30, true);
  await open(file);
  const midBefore = await state();
  await backspaceAt(-2);
  const mid = await state();
  check("Backspace w pustym akapicie za tabelą (dalej tekst): akapit znika", mid.paras === midBefore.paras - 1 && mid.lastText === "Tekst po pustym akapicie", JSON.stringify(mid));
  check("…w pliku tabela i od razu tekst", /^<\/w:tbl><w:p>(<w:pPr>.*?<\/w:pPr>)?<w:r>.*?Tekst po pustym akapicie/.test(mid.afterTable), mid.afterTable);
  check("…kursor na końcu tabeli", mid.caretInTable, JSON.stringify(mid));

  await browser.close();
  if (errors.length) throw new Error(errors.join("\n"));
  if (fails.length) throw new Error(fails.join("\n"));
  console.log(`✅ empty-after-table-playwright passed [${ENGINE}]`);
}

run().catch((error) => {
  console.error("❌ empty-after-table-playwright failed:", error.message || error);
  process.exit(1);
});
