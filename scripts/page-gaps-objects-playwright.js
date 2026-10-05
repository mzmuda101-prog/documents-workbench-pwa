// page-gaps-objects-playwright.js — nic nie wjeżdża w przerwę między stronami (2026-10-06).
//
// Zgłoszenie: obraz na końcu strony stał w poprzek pasa „str. N” (połowa na jednej stronie,
// połowa na drugiej) i „przy różnych manewrach coś wjeżdża, jakby nie było układu strony”.
// Przyczyna: liczenie linijek akapitu nie widziało obrazu w wyspie (span contenteditable=false
// ma wysokość linijki), a część zmian (styl obrazu z suwaka) nie przeliczała granic.
// Zasada sprawdzana po każdym manewrze: żadna linijka tekstu ani obraz nie leży w pasie przerwy
// ani w marginesach wokół niego; obraz, który mieści się na stronie, jest w całości na jednej.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(700));

// Treść w pasie przerwy (z marginesami stron wokół niego) — lista naruszeń.
const violations = (page) => page.evaluate(() => {
  const out = [];
  document.querySelectorAll(".docx-preview-host section.docx").forEach((sec) => {
    const cs = getComputedStyle(sec);
    const scale = sec.getBoundingClientRect().height / sec.offsetHeight || 1;
    const padT = (parseFloat(cs.paddingTop) || 0) * scale;
    const padB = (parseFloat(cs.paddingBottom) || 0) * scale;
    const bands = [...sec.querySelectorAll(":scope > .dwb-page-gap-band")].map((b) => b.getBoundingClientRect());
    if (!bands.length) return;
    const zones = bands.map((b) => [b.top - padB + 2, b.bottom + padT - 2]);
    const rects = [];
    const rg = document.createRange();
    sec.querySelectorAll(":scope > article p").forEach((p) => {
      rg.selectNodeContents(p);
      [...rg.getClientRects()].filter((r) => r.height > 0 && r.width > 0).forEach((r) => rects.push({ r, what: (p.textContent || "").trim().slice(0, 20) || "(pusty)" }));
    });
    sec.querySelectorAll(":scope > article img").forEach((img) => rects.push({ r: img.getBoundingClientRect(), what: "OBRAZ" }));
    rects.forEach(({ r, what }) => {
      zones.forEach(([t, b]) => { if (r.top < b && r.bottom > t) out.push(`${what} [${Math.round(r.top)}–${Math.round(r.bottom)}] w przerwie [${Math.round(t)}–${Math.round(b)}]`); });
    });
  });
  return out;
});

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1400, height: 900 } });
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
  for (let i = 0; i < 24; i++) { await page.keyboard.type(`Linijka ${i + 1}.`); await page.keyboard.press("Enter"); }
  await idle(page);
  const file = await page.evaluateHandle(async () => {
    const c = document.createElement("canvas"); c.width = 1200; c.height = 600;
    const g = c.getContext("2d"); g.fillStyle = "#3a4ea5"; g.fillRect(0, 0, 1200, 600);
    const b = await new Promise((r) => c.toBlob(r, "image/png"));
    return new File([b], "a.png", { type: "image/png" });
  });
  await page.evaluate((f) => composeUi.insertImageFile(f), file);
  await idle(page);
  await page.keyboard.type("Tekst za obrazem.");
  await page.keyboard.press("Enter");
  for (let i = 0; i < 6; i++) { await page.keyboard.type(`Dalej ${i + 1}.`); await page.keyboard.press("Enter"); }
  await idle(page);

  // stan końcowy (przeliczenie idzie po chwili bezczynności — pod obciążeniem bywa później)
  const nBands = () => page.$$eval(".dwb-page-gap-band", (b) => b.length);
  const settled = async () => {
    let v = await violations(page);
    let n = await nBands();
    for (let t = 0; (v.length || !n) && t < 20; t++) { await page.waitForTimeout(200); v = await violations(page); n = await nBands(); }
    return v;
  };
  const step = async (name) => {
    await idle(page);
    const v = await settled();
    const bands = await page.$$eval(".dwb-page-gap-band", (b) => b.length);
    const ok = bands > 0 && !v.length;
    check(`${name}: nic w przerwie między stronami${ok ? "" : ` — ${bands} pasów; ${v.slice(0, 3).join(" | ")}`}`, ok);
  };
  await step("obraz wstawiony na końcu strony (przechodzi w całości na następną)");

  // suwak rozmiaru obrazu: zmiana stylu (bez nowych węzłów) też przelicza granice
  await page.evaluate(() => document.querySelector(".docx-preview-host img").scrollIntoView({ block: "center" }));
  await page.click(".docx-preview-host img");
  await page.waitForSelector(".image-card");
  await page.evaluate(() => { const r = document.querySelector(".image-card input"); r.value = "40"; r.dispatchEvent(new Event("input")); });
  await page.waitForTimeout(900); // podgląd na żywo, przed zapisem
  const live = await settled();
  check("suwak obrazu w trakcie (sam styl, bez zapisu): granice przeliczone", !live.length, live.join(" | "));
  await page.evaluate(() => document.querySelector(".image-card input").dispatchEvent(new Event("change")));
  await page.waitForTimeout(700);
  await step("obraz zmniejszony do 40%");
  await page.evaluate(() => { const r = document.querySelector(".image-card input"); r.value = "100"; r.dispatchEvent(new Event("input")); r.dispatchEvent(new Event("change")); });
  await page.waitForTimeout(700);
  await step("obraz powiększony do 100%");
  await page.click('.image-card [data-align="right"]');
  await step("obraz do prawej");
  await page.keyboard.press("Escape");

  // tekst nad obrazem: Enter dopisuje linijki, Backspace je zabiera — obraz przeskakuje między stronami
  await page.evaluate(() => { const p = [...document.querySelectorAll(".docx-preview-host p.docx-editable-p")].find((x) => x.textContent === "Linijka 10."); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); p.focus?.(); });
  for (let i = 0; i < 3; i++) { await page.keyboard.press("Enter"); await page.keyboard.type(`Nowa ${i + 1}`); }
  await step("Enter nad obrazem (obraz w dół)");
  for (let i = 0; i < 18; i++) await page.keyboard.press("Backspace");
  await step("Backspace nad obrazem (obraz w górę)");
  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  await step("Cofnij");

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : `\n   ${String(r.detail).slice(0, 700)}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
