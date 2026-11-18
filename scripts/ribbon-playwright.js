// ribbon-playwright.js — wstążka (app/ribbon.js): opcje wyświetlania (etap 6) i przewijany rząd.
//
// Kliknięcie kontrolki z gniazda (¶) nie przewija rzędu poleceń (błąd z nagrania 2026-11-16);
// Zawsze / Chowaj przy przewijaniu / Tylko karty (menu przy rzędzie, Plik → Opcje, Ctrl/⌘+F1);
// gęstość Mysz / Dotyk; zapamiętanie po przeładowaniu. ENGINE=webkit.

const pw = require("playwright");
const path = require("path");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const errors = [];

async function open(context) {
  const page = await context.newPage();
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", path.join(__dirname, "..", "docs", "samples", "przewodnik.docx"));
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  await page.evaluate(() => document.querySelector('.mode-btn[data-mode="edit"]').click());
  await sleep(700);
  return page;
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1180, height: 760 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  let page = await open(context);
  const r2h = () => page.evaluate(() => document.querySelector(".rb-r2").getBoundingClientRect().height);
  const vis = () => page.evaluate(() => { const r = document.querySelector(".rb-r2"); return getComputedStyle(r).display !== "none" && r.getBoundingClientRect().height > 20; });

  // ── 1. ¶ w przewiniętym rzędzie: rząd stoi ────────────────────────────────
  await page.evaluate(() => { const s = document.querySelector('[data-strip="cmds"] .rb-sc'); s.scrollLeft = s.scrollWidth; });
  await sleep(300);
  const sl = [];
  const scl = () => page.evaluate(() => Math.round(document.querySelector('[data-strip="cmds"] .rb-sc').scrollLeft));
  sl.push(await scl());
  for (let i = 0; i < 3; i++) { await page.click("#formatMarksBtn"); await sleep(400); sl.push(await scl()); }
  check("klik ¶ w przewiniętym rzędzie: rząd nie skacze", sl[0] > 0 && sl.every((v) => Math.abs(v - sl[0]) <= 1), sl.join(" "));

  // przewijanie touchpadem w bok małymi krokami: rząd jedzie płynnie, bez dociągania do grup
  await page.evaluate(() => { document.querySelector('[data-strip="cmds"] .rb-sc').scrollLeft = 0; });
  await sleep(200);
  const box = await page.locator('[data-strip="cmds"] .rb-sc').boundingBox();
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  const steps = [];
  for (let i = 0; i < 3; i++) { await page.mouse.wheel(23, 0); await sleep(350); steps.push(await scl()); }
  check("touchpad w bok: rząd przesuwa się o tyle, ile ruch (bez przyciągania)", steps.every((v, i) => Math.abs(v - 23 * (i + 1)) <= 2), steps.join(" "));

  // ── 2. menu opcji przy rzędzie ────────────────────────────────────────────
  await page.click("#rbOptsBtn");
  await sleep(250);
  const items = await page.evaluate(() => Array.from(document.querySelectorAll(".compose-pop [role=menuitemradio]")).map((b) => `${b.textContent.trim().slice(0, 24)}:${b.getAttribute("aria-checked")}`));
  check("menu opcji: Uproszczona / Klasyczna, Zawsze (zaznaczone) / Chowaj przy przewijaniu / Tylko karty", items.length === 5 && /^Uproszczona.*:true$/.test(items[0]) && /^Zawsze.*:true$/.test(items[2]), JSON.stringify(items));

  // ── 3. Tylko karty ────────────────────────────────────────────────────────
  await page.click('.compose-pop [role=menuitemradio]:has-text("Tylko karty")');
  await sleep(300);
  check("Tylko karty: rząd poleceń schowany", !(await vis()));
  await page.click("#rbTab-insert");
  await sleep(300);
  const peek = await page.evaluate(() => ({ vis: getComputedStyle(document.querySelector(".rb-r2")).position, panel: !document.getElementById("rbPanel-insert").hidden }));
  check("Tylko karty: klik w kartę wysuwa jej polecenia nad dokumentem", (await vis()) && peek.vis === "absolute" && peek.panel, JSON.stringify(peek));
  await page.mouse.click(640, 600);
  await sleep(300);
  check("Tylko karty: klik w dokument chowa polecenia", !(await vis()));
  await page.keyboard.press("Control+F1");
  await sleep(300);
  check("Ctrl+F1: z powrotem Zawsze", (await vis()) && (await page.evaluate(() => dwbRibbon.prefs().show)) === "pin");
  await page.keyboard.press("Control+F1");
  await sleep(200);
  check("Ctrl+F1 drugi raz: Tylko karty", !(await vis()) && (await page.evaluate(() => dwbRibbon.prefs().show)) === "tabs");

  // ── 4. Chowaj przy przewijaniu (z Plik → Opcje) ───────────────────────────
  await page.evaluate(() => dwbBackstage.open("opts"));
  await sleep(400);
  await page.click('#rbShowSeg [data-rb-show="auto"]');
  await page.evaluate(() => dwbBackstage.close());
  await sleep(400);
  const h0 = await r2h();
  await page.mouse.move(640, 450);
  for (let i = 0; i < 4; i++) { await page.mouse.wheel(0, 200); await sleep(80); }
  await sleep(400);
  const h1 = await r2h();
  await page.mouse.wheel(0, -120);
  await sleep(500);
  const h2 = await r2h();
  check("Chowaj przy przewijaniu: w dół rząd się zwija, w górę wraca", h0 > 30 && h1 < 2 && h2 > 30, JSON.stringify({ h0, h1, h2 }));

  // ── 5. gęstość Dotyk ──────────────────────────────────────────────────────
  await page.evaluate(() => dwbBackstage.open("opts"));
  await sleep(300);
  await page.click('#rbDensSeg [data-rb-dens="touch"]');
  await page.evaluate(() => { dwbBackstage.close(); dwbRibbon.select("home"); });
  await sleep(400);
  const touch = await page.evaluate(() => ({ r2: document.querySelector(".rb-r2").getBoundingClientRect().height, btn: document.getElementById("fmtBold").getBoundingClientRect().height }));
  check("Dotyk: wyższy rząd i przyciski 40 px", touch.r2 >= 50 && touch.btn >= 39, JSON.stringify(touch));

  // ── 6a. Klasyczna ─────────────────────────────────────────────────────────
  await page.evaluate(() => { dwbRibbon.setDens("mouse"); dwbRibbon.setShow("pin"); });
  await page.click("#rbOptsBtn");
  await sleep(250);
  await page.click('.compose-pop [role=menuitemradio]:has-text("Klasyczna")');
  await sleep(400);
  const cl = await page.evaluate(() => ({
    h: document.querySelector(".rb-r2").getBoundingClientRect().height,
    panel: !!document.querySelector(".rb-panel.rb-cl[data-rtab=home]:not([hidden])"),
    simpleHidden: document.querySelector(".rb-panel:not(.rb-cl)[data-rtab=home]").hidden,
    boldIn: !!document.getElementById("fmtBold").closest(".rb-cl"),
    caps: Array.from(document.querySelectorAll(".rb-panel.rb-cl[data-rtab=home] .rb-gcap > span")).map((x) => x.textContent),
  }));
  check("Klasyczna: rząd ~98 px, grupy z nazwami, B to ten sam przycisk (przeniesiony)", cl.h >= 90 && cl.panel && cl.simpleHidden && cl.boldIn && cl.caps.join("|") === "Schowek|Czcionka|Akapit|Style|Edytowanie", JSON.stringify(cl));
  // kursor w zwykłym akapicie, zaznaczenie słowa → B z Klasycznej
  const para = page.locator(".docx-preview-host p").filter({ hasText: "Aplikacja startuje" }).first();
  await para.evaluate((el) => el.scrollIntoView({ block: "center" }));
  await para.dblclick({ position: { x: 30, y: 8 } });
  await sleep(300);
  await page.click("#fmtBold");
  await sleep(500);
  const bold = await page.evaluate(() => ({ on: document.getElementById("fmtBold").classList.contains("is-on"), w: getComputedStyle(getSelection().getRangeAt(0).startContainer.parentElement).fontWeight, fam: document.getElementById("fmtFontFamily").value }));
  check("Klasyczna: B działa na zaznaczeniu, lista kroju pokazuje krój", bold.on && Number(bold.w) >= 600 && !!bold.fam, JSON.stringify(bold));
  await page.click('.rb-cl [data-rb-act="align:center"]');
  await sleep(700);
  const al = await page.evaluate(() => ({ css: getComputedStyle(document.querySelector(".docx-preview-host p.dwb-caret-para, .docx-preview-host p:focus-within") || getSelection().anchorNode.parentElement.closest("p")).textAlign, pressed: document.querySelector('.rb-cl [data-rb-act="align:center"]').getAttribute("aria-pressed") }));
  check("Klasyczna: Wyśrodkuj — akapit wyśrodkowany, przycisk wciśnięty", al.css === "center" && al.pressed === "true", JSON.stringify(al));
  await page.click('.rb-cl [data-rb-act="align:left"]');
  await sleep(500);
  const sty = await page.evaluate(() => Array.from(document.querySelectorAll(".rb-sty-i")).map((b) => b.dataset.style));
  const hStyle = sty.find((v) => /heading1|nagłówek1|^1$/i.test(v.replace(/\s/g, ""))) || sty[3];
  await page.click(`.rb-sty-i[data-style="${hStyle}"]`);
  await sleep(700);
  const st = await page.evaluate(() => ({ sel: document.getElementById("fmtParaStyle").value, pressed: document.querySelector('.rb-sty-i[aria-pressed="true"]')?.dataset.style }));
  check("Klasyczna: galeria stylów nadaje styl i pokazuje bieżący", sty.length >= 4 && st.sel === hStyle && st.pressed === hStyle, JSON.stringify({ n: sty.length, hStyle, st }));
  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  await sleep(500);
  await page.click('.rb-cl .rb-cg:has([data-i18n="rbGFont"]) .rb-launch');
  await sleep(400);
  check("Klasyczna: ↘ przy Czcionce otwiera okno Czcionka", await page.evaluate(() => !!document.querySelector(".compose-pop .ff-family, .compose-pop-font")));
  await page.keyboard.press("Escape");
  // telefon i z powrotem: kontrolki wracają na miejsce i znów do Klasycznej
  await page.setViewportSize({ width: 600, height: 760 });
  await sleep(500);
  const ph = await page.evaluate(() => ({ ribbon: !document.getElementById("ribbon").hidden, boldInRibbon: !!document.getElementById("fmtBold").closest("#ribbon") }));
  await page.setViewportSize({ width: 1180, height: 760 });
  await sleep(500);
  const back = await page.evaluate(() => ({ boldIn: !!document.getElementById("fmtBold").closest(".rb-cl"), cls: document.body.classList.contains("rb-classic") }));
  check("Klasyczna: na telefonie B wraca na swoje miejsce, szeroko znów w Klasycznej", !ph.ribbon && !ph.boldInRibbon && back.boldIn && back.cls, JSON.stringify({ ph, back }));
  // Uproszczona z powrotem: B w swoim gnieździe Uproszczonej
  await page.evaluate(() => dwbRibbon.setLayout("simple"));
  await sleep(300);
  check("Uproszczona z powrotem: B w swoim gnieździe", await page.evaluate(() => !!document.getElementById("fmtBold").closest(".rb-panel:not(.rb-cl)")));
  await page.evaluate(() => { dwbRibbon.setLayout("classic"); dwbRibbon.setShow("auto"); });

  // ── 6. zapamiętane po przeładowaniu ───────────────────────────────────────
  await page.close();
  page = await open(context);
  const kept = await page.evaluate(() => ({ ...dwbRibbon.prefs(), cls: document.body.className }));
  check("po przeładowaniu: Chowaj przy przewijaniu + Klasyczna zostają", kept.show === "auto" && kept.layout === "classic" && /rb-classic/.test(kept.cls) && /rb-show-auto/.test(kept.cls), JSON.stringify(kept));
  await page.evaluate(() => { dwbRibbon.setShow("pin"); dwbRibbon.setDens("mouse"); dwbRibbon.setLayout("simple"); });

  // Esc w polu Szukaj (wstążka): wyjście z pola — pole wraca do zwykłej szerokości, podświetlenia znikają
  {
    const w0 = await page.evaluate(() => Math.round(document.querySelector(".rb-act .doc-search").getBoundingClientRect().width));
    await page.click("#searchQuery");
    await page.keyboard.type("Wstęp");
    await page.waitForTimeout(400);
    const w1 = await page.evaluate(() => Math.round(document.querySelector(".rb-act .doc-search").getBoundingClientRect().width));
    await page.keyboard.press("Escape");
    await page.waitForTimeout(350);
    const st = await page.evaluate(() => ({ inField: document.activeElement?.id === "searchQuery", w: Math.round(document.querySelector(".rb-act .doc-search").getBoundingClientRect().width), hl: (CSS.highlights?.get?.("dwb-search")?.size) || 0 }));
    check("Esc w Szukaj: wyjście z pola, pole wraca do zwykłej szerokości", w1 > w0 && !st.inField && st.w === w0, JSON.stringify({ w0, w1, ...st }));
    await page.fill("#searchQuery", "");
  }

  await browser.close();
  const real = errors.filter((e) => !/ResizeObserver/.test(e));
  if (real.length) check("bez błędów w konsoli", false, real.join(" | ").slice(0, 400));
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed++;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail || ""})`}`);
  }
  console.log(`\n${failed ? "❌" : "✅"} Wstążka [${ENGINE}]: ${results.length - failed}/${results.length}`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
