// image-card-playwright.js — karta obrazu i linki do obrazów (2026-10-04, runda 2).
//
// Suwak rozmiaru (zgłoszenie: „miga, skacze, niestabilny”): przeciąganie PRAWDZIWĄ myszą krok po
// kroku — karta stoi w miejscu, wartość zmienia się tylko w jedną stronę (bez sprzężenia z ruchem
// karty), obraz na żywo = wartość suwaka, jeden zapis po puszczeniu, po zapisie ta sama karta
// (bez znikania) i ten sam rozmiar obrazu (bez „dociągania”). Strzałki: zapis raz, po przerwie.
// Linki do obrazów (Word: Link → Miejsce w tym dokumencie / odsyłacz do rysunku): obraz jako cel
// w okienku linku, najechanie pokazuje podgląd obrazu bez skoku, klik skacze z „↩ Wróć”.
// Z okienka linku: „Powiększ” / klik w obraz / Spacja / Shift+klik w link = podgląd obrazu
// (image-viewer.js) bez przewijania dokumentu; na dotyku — przytrzymanie, potem „Powiększ”.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(500));
const makeImage = (page, w, h) => page.evaluateHandle(async ([w, h]) => {
  const c = document.createElement("canvas"); c.width = w; c.height = h;
  const g = c.getContext("2d"); g.fillStyle = "#3a6ea5"; g.fillRect(0, 0, w, h); g.fillStyle = "#f4c430"; g.fillRect(w / 4, h / 4, w / 2, h / 2);
  const blob = await new Promise((r) => c.toBlob(r, "image/png"));
  return new File([blob], "prostokat.png", { type: "image/png" });
}, [w, h]);
const savedZip = async (page) => {
  const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
  return JSZip.loadAsync(Buffer.from(b64, "base64"));
};

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);
  await page.keyboard.type("Zobacz obraz poniżej.");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Akapit przed obrazem.");
  const img = await makeImage(page, 800, 400);
  await page.evaluate((f) => composeUi.insertImageFile(f), img);
  await idle(page);
  await page.keyboard.type("Rysunek 1. Niebieski prostokąt");
  await page.keyboard.press("Enter");
  for (let i = 0; i < 6; i++) { await page.keyboard.type(`Dalszy tekst ${i + 1}.`); await page.keyboard.press("Enter"); }
  await idle(page);

  // ── suwak: prawdziwe przeciąganie ──────────────────────────────────────────
  await page.click(".docx-preview-host img");
  await page.waitForSelector(".image-card");
  await page.waitForTimeout(300); // animacja pojawienia się karty
  const cardEl = await page.$(".image-card");
  const box = await page.$eval(".image-card input[type=range]", (r) => { const b = r.getBoundingClientRect(); return { x: b.left, y: b.top + b.height / 2, w: b.width }; });
  // kciuk przy 100% → w lewo do ~40%, krokami
  await page.mouse.move(box.x + box.w - 4, box.y);
  await page.mouse.down();
  const samples = [];
  for (let k = 1; k <= 12; k++) {
    await page.mouse.move(box.x + box.w - 4 - (box.w * 0.6) * (k / 12), box.y);
    await page.waitForTimeout(40);
    samples.push(await page.evaluate(() => {
      const card = document.querySelector(".image-card");
      const i = document.querySelector(".docx-preview-host img");
      const p = i.closest("p");
      const cs = getComputedStyle(p);
      const tw = p.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
      return { top: Math.round(card.getBoundingClientRect().top), val: +card.querySelector("input").value, pct: Math.round(i.offsetWidth / tw * 100) };
    }));
  }
  const tops = new Set(samples.map((s) => s.top));
  check("przeciąganie suwaka: karta stoi w miejscu (bez skakania pod kursorem)", tops.size === 1, JSON.stringify(samples.map((s) => s.top)));
  const vals = samples.map((s) => s.val);
  check("przeciąganie w lewo: wartość tylko maleje (bez drgań w przód i w tył)", vals.every((v, i) => !i || v <= vals[i - 1]) && vals[vals.length - 1] <= 50, JSON.stringify(vals));
  check("obraz na żywo = wartość suwaka", samples.every((s) => Math.abs(s.pct - s.val) <= 1), JSON.stringify(samples.map((s) => [s.val, s.pct])));
  const before = await page.evaluate(() => ({ undo: document.getElementById("undoBtn")?.getAttribute("aria-label") || "" }));
  await page.mouse.up();
  await idle(page);
  const after = await page.evaluate(() => {
    const card = document.querySelector(".image-card");
    const i = document.querySelector(".docx-preview-host img");
    const p = i.closest("p"); const cs = getComputedStyle(p);
    const tw = p.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
    return { cards: document.querySelectorAll(".image-card").length, val: +card.querySelector("input").value, pct: i.offsetWidth / tw * 100, sel: i.classList.contains("img-selected") };
  });
  check("po puszczeniu: ta sama karta (nie znika i nie wjeżdża od nowa), obraz zaznaczony", after.cards === 1 && await cardEl.evaluate((el) => el.isConnected) && after.sel, JSON.stringify(after));
  check("po zapisie obraz ma rozmiar z suwaka (bez dociągania)", Math.abs(after.pct - after.val) < 1.2, JSON.stringify(after));
  let zip = await savedZip(page);
  let xml = await zip.file("word/document.xml").async("string");
  const cx = +(xml.match(/<wp:extent cx="(\d+)"/) || [])[1];
  // A4, marginesy 2,5 cm → tekst 16 cm = 5 760 000 EMU (z pliku nowego dokumentu)
  const textEmu = await page.evaluate(async () => { const d = await getDocumentXmlDom(originalFileBytes); const s = d.getElementsByTagNameNS(W_NS, "pgSz")[0]; const m = d.getElementsByTagNameNS(W_NS, "pgMar")[0]; return (+s.getAttributeNS(W_NS, "w") - +m.getAttributeNS(W_NS, "left") - +m.getAttributeNS(W_NS, "right")) * 635; });
  check("plik: szerokość obrazu = wartość suwaka × szerokość tekstu", Math.abs(cx / textEmu * 100 - after.val) < 1, `${cx} / ${textEmu} vs ${after.val}%`);
  const undoAfter = await page.evaluate(() => document.getElementById("undoBtn")?.getAttribute("aria-label") || "");
  check("jedno przeciągnięcie = jeden krok Cofnij „obraz”", /obraz/i.test(undoAfter) && undoAfter !== before.undo, undoAfter);

  // strzałki: zapis raz, po przerwie
  await page.focus(".image-card input[type=range]");
  const renders = await page.evaluate(() => { window.__renders = 0; const orig = window.renderDocxPreview; window.renderDocxPreview = function (...a) { window.__renders++; return orig.apply(this, a); }; return 0; });
  for (let i = 0; i < 4; i++) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(60); }
  await page.waitForTimeout(900);
  await idle(page);
  const kb = await page.evaluate(() => ({ renders: window.__renders, val: +document.querySelector(".image-card input").value }));
  check("strzałki ×4 na suwaku: jedno przerysowanie, wartość +20", kb.renders === 1 && kb.val === after.val + 20, JSON.stringify({ ...kb, from: after.val, renders }));
  await page.keyboard.press("Escape");

  // ── link do obrazu ─────────────────────────────────────────────────────────
  await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0]; const r = formDomRange(p, 7, 12); p.closest(".docx-edit-root").focus({ preventScroll: true }); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await page.evaluate(() => composeUi.openLinkForm(document.getElementById("insertMenuBtn")));
  await page.click(".lf-mode button[data-mode=doc]");
  const opts = await page.evaluate(() => { const out = []; let cur = null; document.querySelectorAll(".lf-places > *").forEach((x) => { if (x.classList.contains("lf-group")) { cur = { label: x.textContent, items: [] }; out.push(cur); } else if (cur && x.classList.contains("lf-place")) cur.items.push(x.querySelector(".lf-place-label").textContent); }); return out.map((g) => `${g.label}: ${g.items.join(", ")}`); });
  check("okienko linku: grupa „Obrazy” z podpisem rysunku", opts.some((o) => /^Obrazy: Rysunek 1\. Niebieski prostokąt/.test(o)), JSON.stringify(opts));
  await page.evaluate((label) => [...document.querySelectorAll(".lf-place")].find((b) => b.querySelector(".lf-place-label").textContent === label)?.click(), "Rysunek 1. Niebieski prostokąt");
  await page.click(".compose-pop-form .lf-ok");
  await idle(page);
  zip = await savedZip(page);
  xml = await zip.file("word/document.xml").async("string");
  const bm = (xml.match(/<w:hyperlink w:anchor="([^"]+)"/) || [])[1];
  const bmPara = bm ? xml.split("<w:p>").concat(xml.split("<w:p ")).find((part) => part.includes(`w:name="${bm}"`)) || "" : "";
  check("plik: link do zakładki w akapicie z obrazem (jak Word: Miejsce w tym dokumencie)", !!bm && /<w:drawing>/.test(bmPara), bm);
  await page.evaluate(() => appFrame.setReadOnly(true));
  await idle(page);
  await page.evaluate(() => { docViewportEl.scrollTop = 0; });
  await page.waitForTimeout(200);
  const st0 = await page.evaluate(() => docViewportEl.scrollTop);
  check("link do obrazu oznaczony (bez dymka z napisem — podgląd zamiast niego)", await page.evaluate(() => !!document.querySelector(".docx-preview-host a[data-dwb-img-link]")));
  await page.hover(".docx-preview-host a[data-dwb-img-link]");
  await page.waitForTimeout(500);
  const peek = await page.evaluate(() => { const el = document.querySelector(".link-peek"); const i = el?.querySelector("img"); return el && { cap: el.querySelector(".link-peek-cap").textContent, loaded: i.complete && i.naturalWidth > 0, inView: el.getBoundingClientRect().bottom <= innerHeight && el.getBoundingClientRect().left >= 0 }; });
  check("najechanie: podgląd obrazu z podpisem, na ekranie", peek && peek.loaded && peek.inView && /Rysunek 1/.test(peek.cap), JSON.stringify(peek));
  check("podgląd nie przewija dokumentu", (await page.evaluate(() => docViewportEl.scrollTop)) === st0);
  const viewerOpen = () => page.evaluate(() => !!document.querySelector(".image-viewer:not([hidden])"));
  const zb = await page.$eval(".link-peek-zoom", (b) => { const r = b.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });
  await page.mouse.move(zb.x, zb.y, { steps: 8 }); // z linku na okienko — nie może zniknąć po drodze
  await page.waitForTimeout(400);
  check("zjazd myszą z linku na okienko: okienko zostaje", !!(await page.$(".link-peek")));
  await page.mouse.click(zb.x, zb.y);
  await page.waitForTimeout(300);
  check("„Powiększ” w okienku linku otwiera podgląd obrazu, dokument stoi w miejscu", await viewerOpen() && !(await page.$(".link-peek")) && (await page.evaluate(() => docViewportEl.scrollTop)) === st0);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(250);
  check("po zamknięciu podglądu dokument dalej w tym samym miejscu (bez skoku do obrazu)", !(await viewerOpen()) && (await page.evaluate(() => docViewportEl.scrollTop)) === st0 && !(await page.$(".link-back:not([hidden])")));
  await page.click(".docx-preview-host a[data-dwb-img-link]", { modifiers: ["Shift"] });
  await page.waitForTimeout(300);
  check("Shift+klik w link do obrazu = od razu podgląd, bez skoku", await viewerOpen() && (await page.evaluate(() => docViewportEl.scrollTop)) === st0 && !(await page.$(".link-back:not([hidden])")));
  await page.keyboard.press("Escape");
  await page.mouse.move(5, 5);
  await page.waitForTimeout(500);
  await page.hover(".docx-preview-host a[data-dwb-img-link]");
  await page.waitForSelector(".link-peek", { timeout: 3000 });
  await page.keyboard.press(" ");
  await page.waitForTimeout(300);
  check("Spacja przy okienku linku = podgląd (jak Szybki podgląd na Macu)", await viewerOpen() && (await page.evaluate(() => docViewportEl.scrollTop)) === st0);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(200);
  await page.hover(".docx-preview-host a[data-dwb-img-link]");
  await page.waitForSelector(".link-peek", { timeout: 3000 });
  await page.mouse.move(5, 5);
  await page.waitForTimeout(900);
  check("zjechanie z linku chowa podgląd", !(await page.$(".link-peek")));
  await page.click(".docx-preview-host a[data-dwb-img-link]");
  await page.waitForTimeout(900);
  const jumped = await page.evaluate(() => ({ back: !!document.querySelector(".link-back:not([hidden])"), imgTop: document.querySelector(".docx-preview-host img").getBoundingClientRect().top, vp: docViewportEl.getBoundingClientRect() }));
  check("klik: skok do obrazu + „↩ Wróć”", jumped.back && jumped.imgTop >= jumped.vp.top - 2 && jumped.imgTop < jumped.vp.bottom, JSON.stringify(jumped));

  // dotyk: przytrzymanie pokazuje podgląd, stuknięcie skacze
  const touch = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 800 }, isMobile: true, hasTouch: true });
  await touch.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const tp = await touch.newPage();
  tp.on("pageerror", (e) => errors.push(`touch: ${e.message}`));
  const bytes = await page.evaluate(async () => { const b = await buildDocumentForSave(); let s = ""; b.forEach((x) => { s += String.fromCharCode(x); }); return s; });
  await tp.goto(APP_URL, { waitUntil: "load" });
  await tp.evaluate(() => document.getElementById("heroSplash")?.remove());
  await tp.evaluate(async (s) => { const u = Uint8Array.from(s, (c) => c.charCodeAt(0)); await ingestFile(new File([u], "obraz-link.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }), { silent: true }); }, bytes);
  await tp.waitForSelector(".docx-preview-host a[data-dwb-img-link]", { timeout: 20000 });
  await tp.waitForTimeout(500);
  const a = await tp.$eval(".docx-preview-host a[data-dwb-img-link]", (el) => { const r = el.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });
  const cdpOk = ENGINE === "chromium";
  if (cdpOk) {
    const cdp = await touch.newCDPSession(tp);
    await cdp.send("Input.dispatchTouchEvent", { type: "touchStart", touchPoints: [{ x: a.x, y: a.y }] });
    await tp.waitForFunction(() => !!document.querySelector(".link-peek"), null, { timeout: 3000 }).catch(() => {});
    check("dotyk: przytrzymanie pokazuje podgląd obrazu", !!(await tp.$(".link-peek")));
    await cdp.send("Input.dispatchTouchEvent", { type: "touchEnd", touchPoints: [] });
    await tp.waitForTimeout(400);
    check("dotyk: po przytrzymaniu nie ma skoku (sam podgląd)", !(await tp.$(".link-back:not([hidden])")));
    const tst0 = await tp.evaluate(() => docViewportEl.scrollTop);
    await tp.tap(".link-peek-zoom");
    await tp.waitForTimeout(400);
    check("dotyk: „Powiększ” w okienku po przytrzymaniu otwiera podgląd, bez skoku", await tp.evaluate(() => !!document.querySelector(".image-viewer:not([hidden])")) && (await tp.evaluate(() => docViewportEl.scrollTop)) === tst0 && !(await tp.$(".link-back:not([hidden])")));
    await tp.tap('.image-viewer .iv-btn[data-act="close"]');
    await tp.waitForTimeout(300);
  }
  await tp.tap(".docx-preview-host a[data-dwb-img-link]");
  await tp.waitForTimeout(900);
  check("dotyk: stuknięcie skacze do obrazu z „Wróć”", !!(await tp.$(".link-back:not([hidden])")));

  // Widok mobilny (telefon): procent suwaka = procent szerokości tekstu STRONY w pliku, bez dociągania
  const filePct = (pg) => pg.evaluate(async () => {
    const d = await getDocumentXmlDom(await buildDocumentForSave());
    const cx = +d.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "extent")[0].getAttribute("cx");
    const s = d.getElementsByTagNameNS(W_NS, "pgSz")[0]; const m = d.getElementsByTagNameNS(W_NS, "pgMar")[0];
    return cx / ((+s.getAttributeNS(W_NS, "w") - +m.getAttributeNS(W_NS, "left") - +m.getAttributeNS(W_NS, "right")) * 635) * 100;
  });
  await tp.evaluate(() => appFrame.setReadOnly(false));
  await tp.waitForFunction(() => !inlineLocksPending, null, { timeout: 10000 });
  await tp.evaluate(() => document.querySelector(".docx-preview-host img").scrollIntoView({ block: "center" }));
  await tp.waitForTimeout(300);
  await tp.tap(".docx-preview-host img");
  await tp.waitForSelector(".image-card");
  await tp.waitForTimeout(500);
  const mob = await tp.evaluate(() => ({ reflow: shouldUseMobileReflow(), val: +document.querySelector(".image-card input").value, kb: document.activeElement?.closest?.(".docx-edit-root") ? "fokus w tekście" : "" }));
  const fp0 = await filePct(tp);
  check("telefon (Widok mobilny): suwak pokazuje procent z pliku, stuknięcie obrazu nie stawia kursora (bez klawiatury)", mob.reflow && Math.abs(mob.val - fp0) <= 5 && !mob.kb, JSON.stringify({ ...mob, fp0 }));
  const live = await tp.evaluate(() => { const r = document.querySelector(".image-card input"); r.value = "30"; r.dispatchEvent(new Event("input")); return document.querySelector(".docx-preview-host img").offsetWidth; });
  await tp.evaluate(() => document.querySelector(".image-card input").dispatchEvent(new Event("change")));
  await tp.waitForTimeout(700);
  await tp.waitForFunction(() => !inlineLocksPending && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 10000 });
  await tp.waitForTimeout(400);
  const after30 = await tp.evaluate(() => document.querySelector(".docx-preview-host img").offsetWidth);
  const fp1 = await filePct(tp);
  check("telefon: 30% na suwaku = 30% w pliku, obraz po zapisie tej samej szerokości co podczas przesuwania", Math.abs(fp1 - 30) < 1 && Math.abs(after30 - live) <= 2, JSON.stringify({ fp1, live, after30 }));

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : `\n   ${String(r.detail).slice(0, 700)}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
