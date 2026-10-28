// image-replace-playwright.js — „Zamień obraz” (2026-10-28, etap 2 edycji kart z PDF).
//
// Zgłoszenie: nowa mapka ma wejść w miejsce starej bez ruszania układu karty. Sprawdzamy: karta
// obrazu → Zamień (ramka wp:extent bez zmian, nowy r:embed, stary plik usunięty z paczki), inne
// proporcje → „Wypełnij” (obraz przycięty do proporcji ramki) i przełącznik „Cała”, wklejenie obrazu
// na zaznaczony obraz, upuszczenie pliku na obraz (zamiana, nie otwarcie jako dokument), Cofnij.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(350));
const mod = process.platform === "darwin" ? "Meta" : "Control";

// 2×1 px PNG (proporcje 2:1 jak ramka 2:1)
const PNG21 = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR42mP8z8DwHwiZGBgAJAsDAfDYyqMAAAAASUVORK5CYII=", "base64");
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const inlinePic = (id, rid, cx, cy) => `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Mapka ${id}" descr="mapa terenu"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Mapka ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:srcRect l="5000" r="5000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

async function fixture() {
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/stara.png"/><Relationship Id="rIdB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/druga.png"/></Relationships>');
  z.file("word/media/stara.png", PNG21);
  z.file("word/media/druga.png", PNG21);
  const body = [
    `<w:p><w:r><w:t>Karta z mapą terenu</w:t></w:r></w:p>`,
    `<w:p><w:r><w:t xml:space="preserve">Mapa: </w:t></w:r>${inlinePic(1, "rIdA", 2286000, 1143000)}</w:p>`, // 6 × 3 cm (2:1)
    `<w:p>${inlinePic(2, "rIdB", 1143000, 1143000)}</w:p>`, // 3 × 3 cm (1:1)
    `<w:p><w:r><w:t>Koniec.</w:t></w:r></w:p>`,
  ].join("");
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return z.generateAsync({ type: "nodebuffer" });
}

// stan pliku: obrazy (rid, extent, srcRect, opis), pliki w word/media, rels obrazów
const fileState = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  const xml = await z.file("word/document.xml").async("string");
  const rels = await z.file("word/_rels/document.xml.rels").async("string");
  const pics = [...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"\/>[\s\S]*?descr="([^"]*)"[\s\S]*?r:embed="([^"]+)"([\s\S]*?)<\/pic:blipFill>/g)].map((m) => ({ cx: +m[1], cy: +m[2], descr: m[3], rid: m[4], srcRect: /srcRect/.test(m[5]) }));
  const media = Object.keys(z.files).filter((f) => f.startsWith("word/media/") && !z.files[f].dir);
  const relTargets = [...rels.matchAll(/Id="([^"]+)"[^>]*Target="(media\/[^"]+)"|Target="(media\/[^"]+)"[^>]*Id="([^"]+)"/g)].map((m) => ({ id: m[1] || m[4], target: m[2] || m[3] }));
  const dims = {};
  for (const f of media) {
    const blob = new Blob([await z.file(f).async("uint8array")]);
    try { const bm = await createImageBitmap(blob); dims[f] = [bm.width, bm.height]; bm.close?.(); } catch (_) { dims[f] = null; }
  }
  return { pics, media, relTargets, dims };
});
const mediaOf = (st, rid) => { const t = st.relTargets.find((r) => r.id === rid)?.target; return t ? `word/${t}` : null; };

// obraz testowy z przeglądarki (kanwa) jako plik w Node
const makeImage = (page, w, h, color) => page.evaluate(async ([w, h, color]) => {
  const c = document.createElement("canvas"); c.width = w; c.height = h;
  const x = c.getContext("2d"); x.fillStyle = color; x.fillRect(0, 0, w, h); x.fillStyle = "#000"; x.fillRect(0, 0, w / 4, h / 4);
  const b = await new Promise((r) => c.toBlob(r, "image/png"));
  return Array.from(new Uint8Array(await b.arrayBuffer()));
}, [w, h, color]).then((a) => Buffer.from(a));

const imgCenter = (page, i) => page.evaluate((i) => { const im = docImageTargets(document.querySelector(".docx-preview-host"))[i].img; im.scrollIntoView({ block: "center" }); const r = im.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; }, i);

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  const tmp = require("os").tmpdir();
  const file = require("path").join(tmp, `dwb-zamien-${process.pid}.docx`);
  require("fs").writeFileSync(file, await fixture());
  await page.setInputFiles("#fileInput", file);
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  require("fs").rmSync(file, { force: true });
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await idle(page);
  const st0 = await fileState(page);

  // 1. przycisk „Zamień” — te same proporcje (2:1): ramka bez zmian, nowy plik, stary usunięty
  const same = require("path").join(tmp, `dwb-21-${process.pid}.png`);
  require("fs").writeFileSync(same, await makeImage(page, 400, 200, "#3a7"));
  let c = await imgCenter(page, 0);
  await page.mouse.click(c.x, c.y);
  await page.waitForSelector(".image-card .ic-replace", { timeout: 5000 });
  let [fc] = await Promise.all([page.waitForEvent("filechooser"), page.click(".image-card .ic-replace")]);
  await fc.setFiles(same);
  await idle(page);
  let st = await fileState(page);
  const p0 = st.pics[0];
  check("Zamień: ramka (wp:extent) bez zmian", p0.cx === st0.pics[0].cx && p0.cy === st0.pics[0].cy, JSON.stringify(p0));
  check("Zamień: nowy plik obrazu (inny r:embed, 400×200)", p0.rid !== "rIdA" && JSON.stringify(st.dims[mediaOf(st, p0.rid)]) === "[400,200]", JSON.stringify({ p0, d: st.dims }));
  check("Zamień: przycięcie starego obrazu (srcRect) usunięte, opis zostaje", !p0.srcRect && p0.descr === "mapa terenu", JSON.stringify(p0));
  check("Zamień: stary plik i jego powiązanie usunięte z paczki", !st.media.includes("word/media/stara.png") && !st.relTargets.some((r) => r.id === "rIdA") && st.media.includes("word/media/druga.png"), JSON.stringify({ media: st.media, rels: st.relTargets }));
  check("Zamień: te same proporcje — bez przełącznika Wypełnij/Cała", await page.evaluate(() => !!document.querySelector(".image-card") && document.querySelector(".image-card .ic-fit")?.hidden), "");
  check("Zamień: drugi obraz nietknięty", st.pics[1].rid === "rIdB", JSON.stringify(st.pics[1]));

  // 2. Cofnij — wraca stary obraz (plik i powiązanie)
  await page.evaluate(() => document.querySelector(".docx-edit-root")?.focus());
  await page.keyboard.press(`${mod}+z`);
  await idle(page);
  st = await fileState(page);
  check("Cofnij przywraca stary obraz (rIdA, stara.png)", st.pics[0].rid === "rIdA" && st.media.includes("word/media/stara.png"), JSON.stringify({ p: st.pics[0], m: st.media }));

  // 3. inne proporcje (pionowy 200×400 w ramce 2:1) → Wypełnij: plik przycięty do 2:1, potem „Cała”
  const tall = require("path").join(tmp, `dwb-12-${process.pid}.png`);
  require("fs").writeFileSync(tall, await makeImage(page, 200, 400, "#c33"));
  c = await imgCenter(page, 0);
  await page.mouse.click(c.x, c.y);
  await page.waitForSelector(".image-card .ic-replace", { timeout: 5000 });
  [fc] = await Promise.all([page.waitForEvent("filechooser"), page.click(".image-card .ic-replace")]);
  await fc.setFiles(tall);
  await idle(page);
  st = await fileState(page);
  let d = st.dims[mediaOf(st, st.pics[0].rid)];
  check("inne proporcje → Wypełnij: obraz przycięty do proporcji ramki (2:1)", d && Math.abs(d[0] / d[1] - 2) < 0.02 && st.pics[0].cx === st0.pics[0].cx, JSON.stringify({ d, p: st.pics[0] }));
  const fitShown = await page.evaluate(() => { const b = document.querySelector(".image-card .ic-fit"); return !!b && !b.hidden; });
  check("po zamianie z innymi proporcjami: przełącznik „Cała” w karcie", fitShown);
  if (fitShown) {
    await page.click(".image-card .ic-fit");
    await idle(page);
    st = await fileState(page);
    d = st.dims[mediaOf(st, st.pics[0].rid)];
    check("„Cała”: cały obraz na płótnie 2:1 (białe pasy), ramka bez zmian", d && Math.abs(d[0] / d[1] - 2) < 0.02 && d[1] >= 400 && st.pics[0].cx === st0.pics[0].cx, JSON.stringify({ d, p: st.pics[0] }));
    const media = st.media.filter((f) => f !== "word/media/druga.png");
    check("przełączanie nie zostawia starych plików w paczce", media.length === 1, JSON.stringify(st.media));
  }

  // 4. wklejenie obrazu przy zaznaczonym obrazie = zamiana (drugi obraz, 1:1)
  c = await imgCenter(page, 1);
  await page.mouse.click(c.x, c.y);
  await page.waitForSelector(".image-card", { timeout: 5000 });
  const sq = await makeImage(page, 120, 120, "#36c");
  await page.evaluate(async (arr) => {
    const f = new File([new Uint8Array(arr)], "wklejony.png", { type: "image/png" });
    const dt = new DataTransfer(); dt.items.add(f);
    const ev = new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true });
    (document.querySelector(".docx-edit-root") || docCanvasEl).dispatchEvent(ev);
  }, Array.from(sq));
  await idle(page);
  st = await fileState(page);
  check("wklejenie na zaznaczony obraz: zamiana (nie nowy obraz)", st.pics.length === 2 && st.pics[1].rid !== "rIdB" && JSON.stringify(st.dims[mediaOf(st, st.pics[1].rid)]) === "[120,120]", JSON.stringify({ pics: st.pics, dims: st.dims }));

  // 5. upuszczenie pliku na obraz: zamiana, dokument dalej ten sam (nie otwarty nowy)
  const name0 = await page.evaluate(() => document.title);
  const drop = await makeImage(page, 300, 150, "#fa0");
  await page.evaluate(async (arr) => {
    const im = docImageTargets(document.querySelector(".docx-preview-host"))[0].img;
    const f = new File([new Uint8Array(arr)], "upuszczony.png", { type: "image/png" });
    const dt = new DataTransfer(); dt.items.add(f);
    const r = im.getBoundingClientRect();
    const opts = { dataTransfer: dt, bubbles: true, cancelable: true, clientX: r.left + 5, clientY: r.top + 5 };
    im.dispatchEvent(new DragEvent("dragenter", opts));
    im.dispatchEvent(new DragEvent("dragover", opts));
    window.__hl = im.classList.contains("img-drop-target");
    im.dispatchEvent(new DragEvent("drop", opts));
  }, Array.from(drop));
  await idle(page);
  st = await fileState(page);
  check("przeciąganie pliku nad obraz: obraz podświetlony", await page.evaluate(() => window.__hl));
  check("upuszczenie pliku na obraz: zamiana w tym samym dokumencie", st.pics.length === 2 && JSON.stringify(st.dims[mediaOf(st, st.pics[0].rid)]) === "[300,150]" && await page.evaluate(() => document.title) === name0, JSON.stringify({ pics: st.pics, dims: st.dims }));

  if (process.env.OUT) require("fs").writeFileSync(process.env.OUT, Buffer.from(await page.evaluate(async () => Array.from(await buildDocumentForSave()))));
  [same, tall].forEach((f) => require("fs").rmSync(f, { force: true }));
  await browser.close();
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Zamień obraz [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => {
  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  console.error(e);
  process.exit(1);
});
