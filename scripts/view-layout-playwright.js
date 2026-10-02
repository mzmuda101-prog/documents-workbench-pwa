// view-layout-playwright.js — Widok mobilny / Widok desktopowy (app/mobile-doc-zoom.js).
//
//   node scripts/view-layout-playwright.js                (Chromium)
//   ENGINE=webkit node scripts/view-layout-playwright.js  (WebKit = Safari/iPad)
//
// Sprawdza to, co widać: czy tekst naprawdę zawija się do ekranu (brak poziomego przewijania,
// kolumna na szerokość), czy w desktopowym są strony o szerokości kartki, czy po zmianie
// widoku i obrocie widać ten sam akapit, i czy NIEZAPISANY tekst przeżywa przełączanie
// (na ekranie i w pliku do zapisu). Pamięć wyboru osobno: dotyk w pionie / w poziomie / komputer.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

const look = (page) => page.evaluate(() => {
  const vp = document.getElementById("docViewport");
  const host = document.querySelector(".docx-preview-host");
  const sec = document.querySelector(".docx-preview-host section.docx");
  const p = [...document.querySelectorAll(".docx-preview-host section.docx > article p")].find((e) => e.textContent.trim().length > 60);
  const lines = p ? (() => { const r = document.createRange(); r.selectNodeContents(p); return new Set([...r.getClientRects()].map((x) => Math.round(x.top))).size; })() : 0;
  return {
    layout: getViewLayout(),
    pref: getViewLayoutPref(),
    reflow: document.getElementById("docCanvas").classList.contains("doc-reflow-mode"),
    zoom: parseFloat(document.getElementById("docCanvas").style.getPropertyValue("--doc-zoom")) || 1,
    label: document.getElementById("zoomNow").textContent.trim(),
    icon: document.getElementById("zoomNow").dataset.layout,
    overflowX: vp.scrollWidth - vp.clientWidth,
    vpW: vp.clientWidth,
    hostW: Math.round(host?.getBoundingClientRect().width || 0),
    hostL: Math.round((host?.getBoundingClientRect().left || 0) - vp.getBoundingClientRect().left),
    secW: Math.round(sec?.getBoundingClientRect().width || 0),
    lines,
  };
});

// pierwszy akapit treści widoczny u góry obszaru dokumentu
const topPara = (page) => page.evaluate(() => {
  const vp = document.getElementById("docViewport");
  const top = vp.getBoundingClientRect().top;
  const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
  const i = ps.findIndex((p) => p.getBoundingClientRect().bottom > top + 1);
  return { i, text: ps[i]?.textContent.slice(0, 30) };
});
const scrollToPara = (page, i) => page.evaluate((i) => {
  const vp = document.getElementById("docViewport");
  const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
  vp.scrollTop += p.getBoundingClientRect().top - vp.getBoundingClientRect().top;
}, i);

async function pickInPopover(page, layout) {
  if (await page.evaluate(() => document.getElementById("viewPop")?.hidden !== false)) await page.click("#zoomNow");
  await page.waitForFunction(() => document.getElementById("viewPop")?.hidden === false, null, { timeout: 5000 }).catch(async () => {
    throw new Error("dymek się nie otworzył: " + JSON.stringify(await page.evaluate(() => ({ hidden: document.getElementById("viewPop")?.hidden, exp: document.getElementById("zoomNow").getAttribute("aria-expanded"), r: document.getElementById("zoomNow").getBoundingClientRect().toJSON() }))));
  });
  await page.click(`#viewPop .view-layout-opt[data-layout="${layout}"]`);
  await idle(page);
}

async function openDoc(page, sample) {
  await page.goto(`${APP_URL}?sample=${sample}`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await idle(page);
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const errors = [];

  // ── telefon (dotyk) ──────────────────────────────────────────────────────────
  const phone = await browser.newContext({
    serviceWorkers: "block", viewport: { width: 390, height: 844 }, hasTouch: true,
    ...(ENGINE === "chromium" ? { isMobile: true } : {}),
  });
  await phone.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await phone.newPage();
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await openDoc(page, "przewodnik");

  const m0 = await look(page);
  check("telefon w pionie, Auto → Widok mobilny (tekst zawija się, bez przewijania w bok)",
    m0.pref === "auto" && m0.layout === "mobile" && m0.reflow && m0.overflowX <= 1 && m0.hostW >= m0.vpW - 2, JSON.stringify(m0));
  check("pasek: ikonka telefonu + 100%", m0.icon === "mobile" && m0.label === "100%", JSON.stringify(m0));

  // dymek z paska
  await page.click("#zoomNow");
  const pop = await page.evaluate(() => {
    const el = document.getElementById("viewPop");
    const r = el.getBoundingClientRect();
    return {
      open: !el.hidden, inside: r.left >= 0 && r.right <= window.innerWidth,
      opts: [...el.querySelectorAll(".view-layout-opt")].map((b) => `${b.dataset.layout}:${b.getAttribute("aria-checked")}`).join(","),
      names: [...el.querySelectorAll(".vl-name")].map((n) => n.textContent).join("|"),
      auto: el.querySelector('[data-layout="auto"] .vl-desc').textContent,
      note: el.querySelector(".view-layout-note").textContent,
    };
  });
  check("stuknięcie w procent otwiera wybór widoku w granicach ekranu", pop.open && pop.inside, JSON.stringify(pop));
  check("wybór: Auto / Widok mobilny / Widok desktopowy, zaznaczone Auto", pop.opts === "auto:true,mobile:false,desktop:false" && pop.names === "Auto|Widok mobilny|Widok desktopowy", JSON.stringify(pop));
  check("Auto mówi, co teraz działa, i dla jakiej sytuacji jest pamiętane", pop.auto === "Teraz: Widok mobilny" && /dotykowy w pionie/.test(pop.note), JSON.stringify(pop));
  await page.keyboard.press("Escape");
  check("Esc zamyka dymek", await page.evaluate(() => document.getElementById("viewPop").hidden));

  // niezapisany tekst + miejsce w dokumencie
  await page.click('.mode-btn[data-mode="edit"]');
  await page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.isContentEditable && p.textContent.trim().length > 20);
    const el = ps[3];
    placeCaret(el, el.textContent.length);
  });
  await page.keyboard.type(" NIEZAPISANE1", { delay: 5 });
  await page.waitForTimeout(700); // touch.js pilnuje kursora nad klawiaturą jeszcze ~0,4 s po pisaniu
  const total = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).length);
  const target = Math.floor(total * 0.6);
  await scrollToPara(page, target);
  await page.waitForTimeout(150);
  const before = await topPara(page);

  await pickInPopover(page, "desktop");
  const d0 = await look(page);
  check("telefon → Widok desktopowy: strony jak na wydruku, cała kartka na szerokość (zoom < 100%)",
    d0.layout === "desktop" && !d0.reflow && d0.zoom < 0.6 && d0.secW > 0 && d0.secW <= d0.vpW && d0.overflowX <= 1, JSON.stringify(d0));
  check("pasek: ikonka monitora", d0.icon === "desktop", JSON.stringify(d0));
  const afterD = await topPara(page);
  check("po zmianie widoku ten sam akapit u góry (nie początek, nie inne miejsce)", Math.abs(afterD.i - before.i) <= 1, JSON.stringify({ before, afterD }));

  await pickInPopover(page, "mobile");
  const afterM = await topPara(page);
  check("i z powrotem — dalej ten sam akapit", Math.abs(afterM.i - before.i) <= 1, JSON.stringify({ before, afterM }));
  await pickInPopover(page, "desktop");

  const kept = await page.evaluate(async () => {
    const view = document.querySelector(".docx-preview-host").textContent;
    const file = (await extractParagraphTextsFromDocx(await buildDocumentForSave())).join("\n");
    return { view: view.includes("NIEZAPISANE1"), file: file.includes("NIEZAPISANE1"), dirty: hasUnsavedChanges };
  });
  check("niezapisany tekst po 3 zmianach widoku: na ekranie, w pliku do zapisu, „Zapisz” dalej świeci", kept.view && kept.file && kept.dirty, JSON.stringify(kept));

  // obrót: poziom ma swoje ustawienie, pion pamięta „desktopowy”
  await page.setViewportSize({ width: 844, height: 390 });
  await page.waitForTimeout(500);
  await idle(page);
  const l0 = await look(page);
  check("obrót do poziomu: Auto → Widok desktopowy, strona się mieści", l0.pref === "auto" && l0.layout === "desktop" && l0.overflowX <= 1, JSON.stringify(l0));
  await pickInPopover(page, "mobile");
  const l1 = await look(page);
  check("w poziomie można wybrać Widok mobilny", l1.pref === "mobile" && l1.reflow && l1.overflowX <= 1, JSON.stringify(l1));
  const beforeRot = await topPara(page);
  await page.setViewportSize({ width: 390, height: 844 });
  await page.waitForTimeout(500);
  await idle(page);
  const r0 = await look(page);
  check("powrót do pionu: pamięta „desktopowy” wybrany w pionie", r0.pref === "desktop" && r0.layout === "desktop" && !r0.reflow, JSON.stringify(r0));
  const afterRot = await topPara(page);
  check("obrót ze zmianą widoku: ten sam akapit u góry", Math.abs(afterRot.i - beforeRot.i) <= 1, JSON.stringify({ beforeRot, afterRot }));
  const kept2 = await page.evaluate(async () => (await extractParagraphTextsFromDocx(await buildDocumentForSave())).join("\n").includes("NIEZAPISANE1") && hasUnsavedChanges);
  check("niezapisany tekst przeżył też obroty", kept2);

  // pamięć po ponownym otwarciu
  await openDoc(page, "headings-sample");
  const re = await look(page);
  check("po ponownym otwarciu pion dalej w Widoku desktopowym", re.pref === "desktop" && !re.reflow, JSON.stringify(re));
  const stored = await page.evaluate(() => JSON.parse(localStorage.getItem("dwb-view-layout-v1")));
  check("zapamiętane osobno: pion = desktopowy, poziom = mobilny", stored["touch-portrait"] === "desktop" && stored["touch-landscape"] === "mobile", JSON.stringify(stored));

  // panel Widok + klawiatura (strzałki w grupie)
  await page.evaluate(() => { setSidebarOpen(true); document.getElementById("panel-view").open = true; });
  await page.focus('#viewLayoutPanel .view-layout-opt[aria-checked="true"]');
  await page.keyboard.press("ArrowUp"); // desktop → mobile
  await idle(page);
  const kb = await look(page);
  check("panel Widok: strzałka zmienia widok (desktopowy → mobilny)", kb.pref === "mobile" && kb.reflow, JSON.stringify(kb));
  await page.keyboard.press("ArrowUp"); // mobile → auto
  await idle(page);
  check("strzałka dalej: Auto (fokus zostaje na wyborze)", await page.evaluate(() => getViewLayoutPref() === "auto" && document.activeElement?.dataset.layout === "auto"));
  await page.evaluate(() => setSidebarOpen(false));

  // − + w Widoku mobilnym zmieniają wielkość liter, tekst dalej się zawija
  const z0 = await look(page);
  await page.evaluate(() => { zoomLevelEl.value = "1.5"; zoomLevelEl.dispatchEvent(new Event("input", { bubbles: true })); });
  const z1 = await look(page);
  check("suwak w Widoku mobilnym: większe litery, więcej linii, bez przewijania w bok", z1.reflow && z1.lines > z0.lines && z1.overflowX <= 1 && z1.label === "150%", JSON.stringify({ z0, z1 }));
  await page.click("#zoomFitBtn");
  check("„Dopasuj” w Widoku mobilnym wraca do 100%", (await look(page)).label === "100%");

  // Widok mobilny przy 200%: litery ×2, wcięcia wolniej (√zoom) — tekst wypełnia ekran jak w Wordzie
  const ind = async () => page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const vpL = document.getElementById("docViewport").getBoundingClientRect().left;
    const left = (re) => { const el = ps.find((x) => re.test(x.textContent)); const r = document.createRange(); r.selectNodeContents(el); const q = r.getClientRects()[0]; return { x: q.left - vpL, h: q.height }; };
    const body = left(/^Pole wyboru|^Ten plik|^Numer:/);
    const ref = ps.find((x) => parseFloat(x.style.getPropertyValue("--dwb-ml0")) > 20);
    const r = document.createRange(); r.selectNodeContents(ref);
    return { body: body.x, h: body.h, indent: r.getClientRects()[0].left - vpL - body.x, ox: document.getElementById("docViewport").scrollWidth - document.getElementById("docViewport").clientWidth };
  });
  await openDoc(page, "przewodnik");
  await page.evaluate(() => setViewLayoutPref("mobile"));
  await idle(page);
  const i1 = await ind();
  await page.evaluate(() => { zoomLevelEl.value = "2"; zoomLevelEl.dispatchEvent(new Event("input", { bubbles: true })); });
  await page.waitForTimeout(200);
  const i2 = await ind();
  check("Widok mobilny 200%: litery ×2, wcięcia i margines tylko ~×1,4, bez przewijania w bok",
    i2.h > i1.h * 1.8 && i2.indent < i1.indent * 1.6 && i2.indent > i1.indent * 1.2 && i2.body < i1.body * 1.6 && i2.ox <= 1, JSON.stringify({ i1, i2 }));
  await page.click("#zoomFitBtn");

  // Widok desktopowy, plik BEZ rozmiaru strony: stała kartka A4, zoom = zdjęcie (tekst się nie przekłada)
  await openDoc(page, "sample");
  await page.evaluate(() => setViewLayoutPref("desktop"));
  await idle(page);
  const page1 = async () => page.evaluate(() => {
    const sec = document.querySelector(".docx-preview-host section.docx");
    const p = [...sec.querySelectorAll("article p")].sort((a, b) => b.textContent.length - a.textContent.length)[0];
    const r = document.createRange(); r.selectNodeContents(p);
    return { z: zoomLevelEl.value, w: Math.round(sec.getBoundingClientRect().width), lines: new Set([...r.getClientRects()].map((x) => Math.round(x.top))).size };
  });
  const a0 = await page1();
  await page.evaluate(() => { zoomLevelEl.value = "2"; zoomLevelEl.dispatchEvent(new Event("input", { bubbles: true })); });
  await page.waitForTimeout(200);
  const a1 = await page1();
  check("Widok desktopowy (plik bez rozmiaru strony): kartka A4 rośnie z zoomem, tekst się nie przekłada",
    Math.abs(a0.w / parseFloat(a0.z) - 794) < 3 && Math.abs(a1.w - 1588) < 4 && a1.lines === a0.lines, JSON.stringify({ a0, a1 }));
  await page.evaluate(() => setViewLayoutPref("auto"));
  await idle(page);

  // EN
  await page.evaluate(() => setLanguage("en"));
  await page.click("#zoomNow");
  const en = await page.evaluate(() => [...document.querySelectorAll("#viewPop .vl-name")].map((n) => n.textContent).join("|") + "/" + document.querySelector("#viewPop .view-layout-note").textContent);
  check("EN: Mobile view / Desktop view", en.startsWith("Auto|Mobile view|Desktop view/Remembered for: touch screen, portrait"), en);
  await page.keyboard.press("Escape");
  await page.evaluate(() => setLanguage("pl"));
  await phone.close();

  // ── komputer (mysz) ──────────────────────────────────────────────────────────
  const desk = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 800 } });
  await desk.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const dp = await desk.newPage();
  dp.on("pageerror", (e) => errors.push(e.message));
  await openDoc(dp, "przewodnik");
  const c0 = await look(dp);
  check("komputer, Auto → Widok desktopowy, 100%", c0.layout === "desktop" && !c0.reflow && c0.zoom === 1 && c0.label === "100%", JSON.stringify(c0));
  // Ctrl + kółko (jak w Wordzie): przybliża dokument, nie aplikację; tekst pod kursorem zostaje pod kursorem
  await dp.evaluate(() => { document.getElementById("docViewport").scrollTop = 900; });
  await dp.waitForTimeout(250);
  const box = await dp.locator("#docViewport").boundingBox();
  const cx = Math.round(box.x + box.width / 2);
  const cy = Math.round(box.y + 260);
  // Znak „pod kursorem” bywa obok kursora (kursor w pustym miejscu za krótkim nagłówkiem —
  // przeglądarka zwraca najbliższą pozycję w wierszu). Po przybliżeniu o u jego odległość od
  // kursora ma urosnąć dokładnie u razy — mierzymy odchyłkę od tego położenia, nie od kursora.
  const caretAt = ([x, y]) => {
    const r = document.caretRangeFromPoint ? document.caretRangeFromPoint(x, y) : (() => { const p = document.caretPositionFromPoint(x, y); const q = document.createRange(); q.setStart(p.offsetNode, p.offset); return q; })();
    window.__n = r.startContainer; window.__o = r.startOffset;
    const q = document.createRange();
    q.setStart(window.__n, window.__o);
    q.setEnd(window.__n, Math.min(window.__o + 1, window.__n.length || 0));
    const b = q.getBoundingClientRect();
    window.__d0 = [b.left - x, b.top + b.height / 2 - y];
  };
  const drift = ([x, y, u]) => {
    const r = document.createRange();
    r.setStart(window.__n, window.__o);
    r.setEnd(window.__n, Math.min(window.__o + 1, window.__n.length || 0));
    const b = r.getBoundingClientRect();
    const ex = x + window.__d0[0] * u, ey = y + window.__d0[1] * u;
    return Math.round(Math.hypot(b.left - ex, b.top + b.height / 2 - ey));
  };
  await dp.evaluate(caretAt, [cx, cy]);
  await dp.mouse.move(cx, cy);
  await dp.keyboard.down("Control");
  await dp.mouse.wheel(0, -100);
  await dp.waitForTimeout(40);
  await dp.mouse.wheel(0, -100);
  await dp.keyboard.up("Control");
  await dp.waitForTimeout(600);
  const w1 = await look(dp);
  const d1 = await dp.evaluate(drift, [cx, cy, 1.2]);
  check("Ctrl + kółko: dwa ząbki = 120% (co 10%, jak Word)", w1.zoom === 1.2 && w1.label === "120%" && !w1.reflow, JSON.stringify(w1));
  check("Ctrl + kółko: tekst pod kursorem zostaje pod kursorem", d1 < 40, String(d1));
  // szczypanie na gładziku = kółko z Ctrl i małymi ułamkowymi krokami → płynnie
  await dp.evaluate(([x, y]) => {
    const vp = document.getElementById("docViewport");
    for (let i = 0; i < 8; i++) vp.dispatchEvent(new WheelEvent("wheel", { deltaY: 4.5, ctrlKey: true, clientX: x, clientY: y, bubbles: true, cancelable: true }));
  }, [cx, cy]);
  await dp.waitForTimeout(500);
  const w2 = await look(dp);
  check("szczypanie na gładziku: płynnie (nie co 10%) w dół", w2.zoom < 1.0 && w2.zoom > 0.75 && Math.round(w2.zoom * 100) % 10 !== 0, JSON.stringify(w2));
  check("zwykłe kółko bez Ctrl dalej przewija", await dp.evaluate(async () => {
    const vp = document.getElementById("docViewport"); const t0 = vp.scrollTop; const z0 = zoomLevelEl.value;
    vp.dispatchEvent(new WheelEvent("wheel", { deltaY: 100, bubbles: true, cancelable: true }));
    return zoomLevelEl.value === z0 && t0 >= 0;
  }));

  // − + i „Dopasuj” zostają w tym samym miejscu dokumentu (jak w Wordzie)
  await scrollToPara(dp, 45);
  await dp.waitForTimeout(250);
  const k0 = await topPara(dp);
  await dp.click("#zoomInBtn");
  await dp.click("#zoomInBtn");
  const k1 = await topPara(dp);
  await dp.click("#zoomFitBtn");
  const k2 = await topPara(dp);
  await dp.click("#zoomOutBtn");
  const k3 = await topPara(dp);
  check("Widok desktopowy: +, +, Dopasuj, − — ten sam akapit u góry", [k1, k2, k3].every((k) => Math.abs(k.i - k0.i) <= 1), JSON.stringify({ k0, k1, k2, k3 }));

  await pickInPopover(dp, "mobile");
  const c1 = await look(dp);
  check("komputer: Widok mobilny wypełnia całą szerokość obszaru, tekst się zawija",
    c1.reflow && c1.hostW >= c1.vpW * 0.95 && c1.overflowX <= 1, JSON.stringify(c1));
  await dp.evaluate(() => { zoomLevelEl.value = "0.5"; zoomLevelEl.dispatchEvent(new Event("input", { bubbles: true })); });
  await dp.waitForTimeout(200);
  const c1b = await look(dp);
  check("…także przy oddaleniu do 50% (kolumna się nie zwęża)", c1b.reflow && c1b.hostW >= c1b.vpW * 0.95 && c1b.overflowX <= 1, JSON.stringify(c1b));
  await dp.click("#zoomFitBtn");
  await scrollToPara(dp, 45);
  await dp.waitForTimeout(250);
  const m0b = await topPara(dp);
  await dp.click("#zoomInBtn");
  const m1b = await topPara(dp);
  await dp.click("#zoomOutBtn");
  await dp.click("#zoomOutBtn");
  const m2b = await topPara(dp);
  check("Widok mobilny: − + (tekst zawija się inaczej) — ten sam akapit u góry", [m1b, m2b].every((k) => Math.abs(k.i - m0b.i) <= 1), JSON.stringify({ m0b, m1b, m2b }));
  check("Widok mobilny nie schodzi poniżej 50%", await dp.evaluate(() => { for (let i = 0; i < 8; i++) document.getElementById("zoomOutBtn").click(); return zoomLevelEl.value === "0.5"; }));
  await dp.click("#zoomFitBtn");
  const note = await dp.evaluate(() => document.querySelector("#viewPop .view-layout-note")?.textContent || document.querySelector("#viewLayoutPanel .view-layout-note").textContent);
  check("komputer: pamiętane „dla: komputer”", /komputer/.test(note), note);
  await pickInPopover(dp, "auto");
  await dp.setViewportSize({ width: 700, height: 800 });
  await dp.waitForTimeout(500);
  await idle(dp);
  const c2 = await look(dp);
  check("wąskie okno komputera (Auto) → Widok mobilny", c2.layout === "mobile" && c2.reflow && c2.overflowX <= 1, JSON.stringify(c2));
  const tip = await dp.evaluate(() => [...document.querySelectorAll(".toast")].map((t) => t.textContent).join("|"));
  check("przy samoczynnej zmianie widoku jednorazowa podpowiedź", /Widok mobilny — zmienisz go/.test(tip), tip);
  await dp.setViewportSize({ width: 1280, height: 800 });
  await dp.waitForTimeout(500);
  await idle(dp);
  const c3 = await look(dp);
  check("szerokie okno z powrotem: Widok desktopowy, strona się mieści", c3.layout === "desktop" && c3.overflowX <= 1, JSON.stringify(c3));
  const tips = await dp.evaluate(() => [...document.querySelectorAll(".toast")].filter((t) => /zmienisz go/.test(t.textContent)).length);
  check("podpowiedź tylko raz", tips <= 1, String(tips));
  await desk.close();

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ widok mobilny/desktopowy [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
