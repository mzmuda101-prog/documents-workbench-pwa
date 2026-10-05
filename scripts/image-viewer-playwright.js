// image-viewer-playwright.js — podgląd obrazu (image-viewer.js, 2026-10-06).
//
// Tylko oglądanie (rozmiar w pliku bez zmian): otwarcie z karty obrazu, dwuklikiem w Edycji
// i kliknięciem w Czytaniu; dopasowanie do okna, przybliżanie w miejscu kursora (kółko, +),
// przesuwanie przeciąganiem bez wyjeżdżania poza krawędź, 1:1, dwuklik przybliż ↔ dopasuj,
// poprzedni / następny obraz, „Pół ekranu” (dokument obok dalej przewijalny, klawisze tylko
// z fokusem w podglądzie), Esc zamyka i oddaje fokus, Delete pod podglądem nie usuwa obrazu.
// Telefon: szczypanie dwoma palcami (CDP touch, tylko Chromium), pół ekranu = dolna połowa.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));
const makeImage = (page, w, h, color) => page.evaluateHandle(async ([w, h, color]) => {
  const c = document.createElement("canvas"); c.width = w; c.height = h;
  const g = c.getContext("2d"); g.fillStyle = color; g.fillRect(0, 0, w, h); g.fillStyle = "#f4c430"; g.fillRect(w / 4, h / 4, w / 2, h / 2);
  const blob = await new Promise((r) => c.toBlob(r, "image/png"));
  return new File([blob], "obraz.png", { type: "image/png" });
}, [w, h, color]);
const state = (page) => page.evaluate(() => {
  const v = document.querySelector(".image-viewer");
  const im = v?.querySelector(".iv-img");
  const st = v?.querySelector(".iv-stage")?.getBoundingClientRect();
  const r = im?.getBoundingClientRect();
  return v && !v.hidden ? {
    open: true, half: v.classList.contains("iv-half"), zoom: v.querySelector(".iv-zoom-val").textContent,
    count: v.querySelector(".iv-count").textContent, cap: v.querySelector(".iv-cap").textContent,
    nw: im.naturalWidth, w: Math.round(r.width), h: Math.round(r.height), l: Math.round(r.left - st.left), t: Math.round(r.top - st.top),
    sw: Math.round(st.width), sh: Math.round(st.height), sl: Math.round(st.left), stop: Math.round(st.top), focusIn: v.contains(document.activeElement),
  } : { open: false };
});
const docImgWidths = async (page) => {
  const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
  const zip = await JSZip.loadAsync(Buffer.from(b64, "base64"));
  const xml = await zip.file("word/document.xml").async("string");
  return [...xml.matchAll(/<wp:extent cx="(\d+)"/g)].map((m) => +m[1]);
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
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); localStorage.removeItem("dwb.ivMode"); });
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);
  await page.keyboard.type("Pierwszy obraz:");
  await page.evaluate((f) => composeUi.insertImageFile(f), await makeImage(page, 2400, 1600, "#3a6ea5"));
  await idle(page);
  for (let i = 0; i < 4; i++) { await page.keyboard.type(`Tekst między obrazami ${i + 1}.`); await page.keyboard.press("Enter"); }
  await page.evaluate((f) => composeUi.insertImageFile(f), await makeImage(page, 300, 500, "#a53a6e"));
  await idle(page);
  await page.keyboard.type("Koniec.");
  await idle(page);
  const widths0 = await docImgWidths(page);

  // ── z karty obrazu ─────────────────────────────────────────────────────────
  await page.evaluate(() => document.querySelector(".docx-preview-host img").scrollIntoView({ block: "center" }));
  await page.click(".docx-preview-host img");
  await page.waitForSelector(".image-card .ic-view");
  await page.click(".image-card .ic-view");
  await page.waitForSelector(".image-viewer:not([hidden])");
  await page.waitForTimeout(300);
  let s = await state(page);
  check("karta obrazu → podgląd: otwarty w pełnej rozdzielczości, „Obraz 1 z 2”, fokus w podglądzie", s.open && s.nw === 2400 && /1/.test(s.count) && /2/.test(s.count) && s.focusIn, JSON.stringify(s));
  check("dopasowany do okna: cały obraz widać, wyśrodkowany", s.w <= s.sw && s.h <= s.sh && Math.abs(s.l - (s.sw - s.w) / 2) <= 2 && Math.abs(s.t - (s.sh - s.h) / 2) <= 2, JSON.stringify(s));
  const fitW = s.w;

  // kółko w miejscu kursora: punkt pod kursorem zostaje pod kursorem
  const px = s.sl + s.l + s.w * 0.25; const py = s.stop + s.t + s.h * 0.25;
  const before = await page.evaluate(([x, y]) => { const r = document.querySelector(".iv-img").getBoundingClientRect(); return [(x - r.left) / r.width, (y - r.top) / r.height]; }, [px, py]);
  await page.mouse.move(px, py);
  for (let i = 0; i < 5; i++) { await page.mouse.wheel(0, -120); await page.waitForTimeout(30); }
  await page.waitForTimeout(100);
  s = await state(page);
  const after = await page.evaluate(([x, y]) => { const r = document.querySelector(".iv-img").getBoundingClientRect(); return [(x - r.left) / r.width, (y - r.top) / r.height]; }, [px, py]);
  check("kółko przybliża w miejscu kursora (ten sam punkt obrazu pod kursorem)", s.w > fitW * 1.5 && Math.abs(after[0] - before[0]) < 0.01 && Math.abs(after[1] - before[1]) < 0.01, JSON.stringify({ w: s.w, fitW, before, after }));

  // przeciąganie daleko w prawo-dół: lewa/górna krawędź obrazu nie odjeżdża od krawędzi okna
  await page.mouse.move(s.sl + s.sw / 2, s.stop + s.sh / 2);
  await page.mouse.down();
  await page.mouse.move(s.sl + s.sw / 2 + 3000, s.stop + s.sh / 2 + 3000, { steps: 8 });
  await page.mouse.up();
  s = await state(page);
  check("przeciąganie przesuwa, ale obraz nie wyjeżdża poza krawędź (bez pustego pasa)", s.l === 0 && s.t === 0, JSON.stringify(s));

  await page.click('.iv-btn[data-act="actual"]');
  s = await state(page);
  check("1:1 = rzeczywisty rozmiar (100%)", s.w === 2400 && s.zoom === "100%", JSON.stringify(s));
  await page.keyboard.press("0");
  s = await state(page);
  check("klawisz 0 = dopasuj", s.w === fitW, JSON.stringify(s));
  await page.keyboard.press("+");
  const plusW = (await state(page)).w;
  await page.keyboard.press("-");
  check("+ / − przybliża i oddala", plusW > fitW && (await state(page)).w === fitW, JSON.stringify({ plusW, fitW }));

  // dwuklik: przybliż ↔ dopasuj
  const c = { x: s.sl + s.sw / 2, y: s.stop + s.sh / 2 };
  await page.mouse.dblclick(c.x, c.y);
  await page.waitForTimeout(100);
  const dz = (await state(page)).w;
  await page.mouse.dblclick(c.x, c.y);
  await page.waitForTimeout(100);
  check("dwuklik: przybliża, drugi dwuklik wraca do dopasowania", dz > fitW * 2 && (await state(page)).w === fitW, JSON.stringify({ dz, fitW }));

  // Delete / Backspace w podglądzie nie usuwa obrazu z dokumentu
  await page.keyboard.press("Delete");
  await page.keyboard.press("Backspace");
  await page.waitForTimeout(400);
  check("Delete w podglądzie nie usuwa obrazu pod spodem", (await page.$$(".docx-preview-host img")).length === 2 && (await state(page)).open);

  // następny / poprzedni
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(200);
  s = await state(page);
  check("→ następny obraz (Obraz 2 z 2)", s.nw === 300 && /^\D*2\D+2/.test(s.count), JSON.stringify(s));
  await page.click(".iv-nav-next");
  await page.waitForTimeout(200);
  check("następny po ostatnim = pierwszy (w kółko)", (await state(page)).nw === 2400);

  // pół ekranu
  await page.click('.iv-btn[data-act="mode"]');
  await page.waitForTimeout(250);
  s = await state(page);
  check("Pół ekranu: panel z prawej, połowa szerokości, obraz dalej dopasowany", s.half && Math.abs(s.sl - 640) <= 2 && s.w <= s.sw && s.h <= s.sh, JSON.stringify(s));
  // granica panelu: przeciągnięcie w prawo zmniejsza panel, aplikacja obok się poszerza; zapamiętane
  const sepBox = await page.$eval(".iv-split", (el) => { const r = el.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });
  await page.mouse.move(sepBox.x, sepBox.y);
  await page.mouse.down();
  await page.mouse.move(1280 * 0.7, sepBox.y, { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(250);
  const split = await page.evaluate(() => ({
    panel: Math.round(document.querySelector(".image-viewer").getBoundingClientRect().width),
    app: Math.round(document.querySelector(".app").getBoundingClientRect().right),
    saved: localStorage.getItem("dwb.ivSplitSide"),
    img: Math.round(document.querySelector(".iv-img").getBoundingClientRect().width),
    stage: Math.round(document.querySelector(".iv-stage").getBoundingClientRect().width),
  }));
  check("przeciągnięcie granicy: panel węższy (~30%), aplikacja obok szersza, obraz dopasowany, rozmiar zapamiętany", Math.abs(split.panel - 384) <= 4 && split.app <= 1280 - 380 && split.img <= split.stage && Math.abs(+split.saved - 0.3) < 0.01, JSON.stringify(split));
  await page.focus(".iv-split");
  await page.keyboard.press("ArrowLeft");
  await page.waitForTimeout(150);
  const kb = await page.evaluate(() => Math.round(document.querySelector(".image-viewer").getBoundingClientRect().width));
  check("strzałka na uchwycie powiększa panel o 5%", Math.abs(kb - 448) <= 4, String(kb));
  await page.evaluate(() => { localStorage.removeItem("dwb.ivSplitSide"); });
  await page.click('.iv-btn[data-act="fit"]');
  const docClick = await page.evaluate(() => {
    const el = document.elementFromPoint(300, 500);
    return !!el && !el.closest(".image-viewer");
  });
  check("Pół ekranu: lewa połowa = dokument (do czytania / klikania)", docClick);
  // klik w drugi obraz w dokumencie (w Czytaniu) podmienia podgląd
  await page.keyboard.press("Escape");
  await page.evaluate(() => appFrame.setReadOnly(true));
  await page.waitForTimeout(400);
  await page.locator(".docx-preview-host img").nth(1).click();
  await page.waitForTimeout(300);
  s = await state(page);
  check("Czytanie: klik w obraz otwiera podgląd (pamięta Pół ekranu)", s.open && s.half && s.nw === 300, JSON.stringify(s));
  // klawisze z fokusem w dokumencie nie sterują podglądem; Esc w podglądzie zamyka
  await page.evaluate(() => document.activeElement?.blur());
  await page.mouse.click(200, 400); // klik w dokument
  const z0 = (await state(page)).zoom;
  await page.keyboard.press("+");
  const zKey = (await state(page)).zoom;
  await page.click('.iv-btn[data-act="in"]');
  const zBtn = (await state(page)).zoom;
  check("Pół ekranu: + z fokusem w dokumencie nie przybliża, przycisk przybliża", zKey === z0 && zBtn !== z0, JSON.stringify({ z0, zKey, zBtn }));
  await page.click('.iv-btn[data-act="mode"]');
  await page.waitForTimeout(200);
  check("przełącznik wraca na Całe okno", !(await state(page)).half);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(200);
  check("Esc zamyka podgląd", !(await state(page)).open);

  // Edycja: dwuklik w obraz otwiera podgląd
  await page.evaluate(() => appFrame.setReadOnly(false));
  await page.waitForFunction(() => !inlineLocksPending, null, { timeout: 10000 });
  await page.waitForTimeout(300);
  await page.evaluate(() => document.querySelector(".docx-preview-host img").scrollIntoView({ block: "center" }));
  await page.dblclick(".docx-preview-host img");
  await page.waitForTimeout(300);
  s = await state(page);
  check("Edycja: dwuklik w obraz otwiera podgląd", s.open && s.nw === 2400, JSON.stringify(s));
  await page.click('.iv-btn[data-act="close"]');
  await page.waitForTimeout(200);
  check("Zamknij: podgląd znika, karta obrazu zostaje", !(await state(page)).open && !!(await page.$(".image-card")));
  const widths1 = await docImgWidths(page);
  check("tylko oglądanie: rozmiary obrazów w pliku bez zmian", JSON.stringify(widths0) === JSON.stringify(widths1), JSON.stringify({ widths0, widths1 }));
  check("brak błędów strony", !errors.length, errors.join(" | "));

  // ── telefon: szczypanie, pół ekranu = dół ───────────────────────────────────
  const phone = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 }, isMobile: ENGINE === "chromium", hasTouch: true, deviceScaleFactor: 2 });
  await phone.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const tp = await phone.newPage();
  tp.on("pageerror", (e) => errors.push(`telefon: ${e.message}`));
  await tp.goto(APP_URL, { waitUntil: "load" });
  await tp.evaluate(() => { document.getElementById("heroSplash")?.remove(); localStorage.removeItem("dwb.ivMode"); });
  await tp.evaluate(() => composeUi.createNew("blank"));
  await tp.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(tp);
  await tp.evaluate((f) => composeUi.insertImageFile(f), await makeImage(tp, 1200, 900, "#3a6ea5"));
  await idle(tp);
  await tp.evaluate(() => appFrame.setReadOnly(true));
  await tp.waitForTimeout(400);
  await tp.evaluate(() => document.querySelector(".docx-preview-host img").scrollIntoView({ block: "center" }));
  await tp.waitForTimeout(200);
  await tp.tap(".docx-preview-host img");
  await tp.waitForSelector(".image-viewer:not([hidden])");
  await tp.waitForTimeout(300);
  s = await state(tp);
  check("telefon: stuknięcie w obraz (Czytanie) otwiera podgląd na całe okno", s.open && !s.half && s.w <= s.sw, JSON.stringify(s));
  const toolsFit = await tp.evaluate(() => { const b = document.querySelector(".iv-bar"); return b.scrollWidth <= b.clientWidth + 1; });
  check("telefon: pasek podglądu nie rozpycha ekranu (przyciski przewijane w bok)", toolsFit && await tp.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
  if (ENGINE === "chromium") {
    const fitW2 = s.w;
    const cdp = await phone.newCDPSession(tp);
    const cx = s.sl + s.sw / 2; const cy = s.stop + s.sh / 2;
    const tps = (d) => [{ x: cx - d, y: cy, id: 0 }, { x: cx + d, y: cy, id: 1 }];
    await cdp.send("Input.dispatchTouchEvent", { type: "touchStart", touchPoints: tps(30) });
    for (let d = 40; d <= 120; d += 10) { await cdp.send("Input.dispatchTouchEvent", { type: "touchMove", touchPoints: tps(d) }); await tp.waitForTimeout(16); }
    await cdp.send("Input.dispatchTouchEvent", { type: "touchEnd", touchPoints: [] });
    await tp.waitForTimeout(150);
    s = await state(tp);
    check("telefon: rozsunięcie palców przybliża obraz (~4×), a nie całą stronę", s.w > fitW2 * 3 && await tp.evaluate(() => (window.visualViewport?.scale || 1) === 1), JSON.stringify({ w: s.w, fitW2 }));
  }
  await tp.click('.iv-btn[data-act="mode"]');
  await tp.waitForTimeout(250);
  s = await state(tp);
  check("telefon: Pół ekranu = dolna połowa", s.half && s.stop > 844 * 0.45, JSON.stringify(s));
  check("brak błędów strony (telefon)", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : `\n   ${String(r.detail).slice(0, 700)}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
