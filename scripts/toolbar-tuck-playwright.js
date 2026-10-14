// toolbar-tuck-playwright.js — pasek nad dokumentem chowa się przy przewijaniu (telefon, app-frame.js).
//
//   node scripts/toolbar-tuck-playwright.js                (Chromium, telefon z dotykiem)
//   ENGINE=webkit node scripts/toolbar-tuck-playwright.js  (WebKit, telefon z dotykiem)
//
// Pilnuje: pasek leży NAD dokumentem i nie zasłania początku; chowa się i wraca tylko po ruchu
// użytkownika (kółko / palec), skoki programu go nie ruszają; chowanie nie przesuwa tekstu;
// pigułka, fokus w pasku, zaznaczenie w Edycji i dojechanie na górę przywracają pasek;
// skok do sekcji nie ląduje pod paskiem; szeroko nic się nie zmienia.
// Prawdziwego przeciągania palcem z rozpędem Playwright nie zrobi — to symulator/telefon.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({
    serviceWorkers: "block",
    viewport: { width: 390, height: 844 },
    hasTouch: true,
    ...(ENGINE === "chromium" ? { isMobile: true } : {}),
  });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));

  await page.goto(`${APP_URL}?sample=przewodnik`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.waitForTimeout(600);

  const state = () => page.evaluate(() => {
    const vp = document.getElementById("docViewport");
    const tb = document.getElementById("docToolbar");
    const panel = document.querySelector(".doc-panel").getBoundingClientRect();
    const peek = document.getElementById("tbPeek");
    return {
      float: document.body.classList.contains("tb-float"),
      tucked: document.body.classList.contains("tb-tucked"),
      tbBottom: tb.getBoundingClientRect().bottom,
      tbH: tb.offsetHeight,
      panelTop: panel.top,
      pad: parseFloat(getComputedStyle(vp).paddingTop),
      y: vp.scrollTop,
      peek: !peek.hidden && getComputedStyle(peek).display !== "none",
      peekText: peek.textContent.trim(),
      islandTop: document.getElementById("docIsland").getBoundingClientRect().top,
      islandShown: !document.getElementById("docIsland").hidden && getComputedStyle(document.getElementById("docIsland")).opacity > 0.5,
      panelBottom: panel.bottom,
    };
  });
  // pozycja akapitu w dokumencie niezależna od przewinięcia — zmienia się tylko, gdy układ skacze
  const docPos = () => page.evaluate(() => {
    const vp = document.getElementById("docViewport");
    const p = document.querySelectorAll(".docx-preview-host p")[12];
    return p.getBoundingClientRect().top - vp.getBoundingClientRect().top + vp.scrollTop;
  });
  const vpBox = await page.locator("#docViewport").boundingBox();
  const wheel = async (dy, steps = 4) => {
    await page.mouse.move(vpBox.x + vpBox.width / 2, vpBox.y + vpBox.height * 0.7);
    for (let i = 0; i < steps; i++) { await page.mouse.wheel(0, dy / steps); await page.waitForTimeout(40); }
    await page.waitForTimeout(350);
  };

  let s = await state();
  check("telefon: pasek leży nad dokumentem (tb-float)", s.float, JSON.stringify(s));
  check("odstęp u góry dokumentu = wysokość paska (nic nie zasłonięte na starcie)", Math.abs(s.pad - s.tbH) <= 1 && s.tbH > 40, `${s.pad} vs ${s.tbH}`);
  const firstTop = await page.evaluate(() => document.querySelector(".docx-preview-host p, .docx-preview-host h1").getBoundingClientRect().top);
  check("pierwszy akapit zaczyna się pod paskiem", firstTop >= s.tbBottom - 1, `${firstTop} vs ${s.tbBottom}`);

  // ── wyspa u dołu ──
  const isl = await page.evaluate(() => {
    const island = document.getElementById("docIsland");
    const ms = document.querySelector(".mode-switch").getBoundingClientRect();
    const zg = document.querySelector(".zoom-group").getBoundingClientRect();
    return {
      shown: !island.hidden,
      search: island.contains(document.getElementById("searchQuery")),
      tools: island.contains(document.getElementById("panelToggle")),
      statusHidden: getComputedStyle(document.querySelector(".status-bar")).display === "none",
      oneRow: Math.abs(ms.top + ms.height / 2 - (zg.top + zg.height / 2)) < 4,
      bottomGap: document.querySelector(".doc-panel").getBoundingClientRect().bottom - island.getBoundingClientRect().bottom,
    };
  });
  check("wyspa u dołu: szukanie i narzędzia przeniesione", isl.shown && isl.search && isl.tools, JSON.stringify(isl));
  check("pasek stanu na telefonie schowany", isl.statusHidden);
  check("Czytanie: reszta paska w jednym rzędzie (zoom obok Czytanie/Edycja)", isl.oneRow);
  check("wyspa unosi się nad dołem obszaru dokumentu", isl.bottomGap >= 6 && isl.bottomGap <= 16, String(isl.bottomGap));
  const end = await page.evaluate(async () => {
    const vp = document.getElementById("docViewport");
    vp.scrollTop = vp.scrollHeight;
    await new Promise((r) => setTimeout(r, 400));
    const ps = [...document.querySelectorAll(".docx-preview-host p")].filter((p) => p.textContent.trim());
    return { last: ps[ps.length - 1].getBoundingClientRect().bottom, island: document.getElementById("docIsland").getBoundingClientRect().top };
  });
  check("koniec dokumentu da się wysunąć nad wyspę", end.last <= end.island + 1, JSON.stringify(end));
  await page.evaluate(() => { document.getElementById("docViewport").scrollTop = 0; });
  await page.waitForTimeout(300);

  const before = await docPos();
  await wheel(600);
  s = await state();
  check("przewinięcie w dół (kółko) chowa pasek", s.tucked && s.tbBottom <= s.panelTop + 1, JSON.stringify(s));
  check("pigułka widoczna, z nazwą sekcji albo pliku", s.peek && s.peekText.length > 0, s.peekText);
  check("wyspa chowa się razem z paskiem", !s.islandShown && s.islandTop >= s.panelBottom - 1, JSON.stringify(s));
  check("schowanie paska nie przesuwa tekstu", Math.abs((await docPos()) - before) < 1);

  await page.evaluate(() => { document.getElementById("docViewport").scrollTop -= 300; });
  await page.waitForTimeout(300);
  check("przewinięcie przez program (bez kółka/palca) nie pokazuje paska", (await state()).tucked);

  await wheel(-160);
  s = await state();
  check("ruch w górę pokazuje pasek i wyspę", !s.tucked && s.islandShown);

  await wheel(500);
  check("znowu w dół — schowany", (await state()).tucked);
  await page.locator("#tbPeek").tap();
  await page.waitForTimeout(350);
  check("stuknięcie w pigułkę pokazuje pasek", !(await state()).tucked);

  await wheel(500);
  await page.evaluate(() => document.getElementById("searchQuery").focus());
  await page.waitForTimeout(300);
  check("fokus w szukaniu (Ctrl/⌘+F, Tab) pokazuje pasek", !(await state()).tucked);
  await page.evaluate(() => document.activeElement.blur());
  await wheel(500);
  check("bez fokusu w pasku — chowa się dalej", (await state()).tucked);

  await page.evaluate(() => { document.getElementById("docViewport").scrollTop = 0; });
  await page.waitForTimeout(300);
  check("na samej górze pasek wraca (bez pustego pasa nad tekstem)", !(await state()).tucked);

  // szukanie w wyspie: trafienia, „Odznacz” stoi NAD wyspą (nie na niej)
  await page.locator("#searchQuery").tap();
  await page.keyboard.type("najemca");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(900);
  const hit = await page.evaluate(() => {
    const d = document.getElementById("docDismissBtn");
    // dokładne trafienia to podświetlenia CSS (dwb-find), nie elementy — bierzemy te w oknie
    const vp = document.getElementById("docViewport").getBoundingClientRect();
    const ranges = [...(CSS.highlights?.get("dwb-find") || [])].map((r) => r.getBoundingClientRect()).filter((r) => r.height && r.bottom > vp.top && r.top < vp.bottom);
    return {
      pos: document.getElementById("searchPos").textContent.trim(),
      dismiss: !d.hidden ? d.getBoundingClientRect().bottom : null,
      island: document.getElementById("docIsland").getBoundingClientRect().top,
      hitBottom: ranges.length ? Math.max(...ranges.map((r) => r.bottom)) : null,
    };
  });
  check("szukanie z wyspy znajduje trafienia", /\d/.test(hit.pos), hit.pos);
  check("trafienie nie leży pod wyspą", hit.hitBottom != null && hit.hitBottom <= hit.island, JSON.stringify(hit));
  if (hit.dismiss != null) check("„Odznacz” nad wyspą", hit.dismiss <= hit.island - 4, JSON.stringify(hit));
  await page.evaluate(() => { const q = document.getElementById("searchQuery"); q.value = ""; q.dispatchEvent(new Event("input", { bubbles: true })); q.blur(); window.dwbKeyboard?.clearHighlights?.(); });
  await page.evaluate(() => { document.getElementById("docViewport").scrollTop = 0; });
  await page.waitForTimeout(300);

  // skok do sekcji: nagłówek nie ląduje pod paskiem
  const chip = page.locator("#sectionChips .section-chip").nth(3);
  if (await chip.count()) {
    await chip.tap();
    await page.waitForTimeout(2100);
    const r = await page.evaluate(() => {
      const label = document.querySelectorAll("#sectionChips .section-chip")[3].textContent.replace("…", "").trim();
      // ostatnie trafienie = nagłówek w treści (wcześniejsze to pozycja spisu treści)
      const h = [...document.querySelectorAll(".docx-preview-host h1, .docx-preview-host h2, .docx-preview-host p")].filter((el) => el.textContent.trim().startsWith(label)).pop();
      return { top: h?.getBoundingClientRect().top, tb: document.getElementById("docToolbar").getBoundingClientRect().bottom };
    });
    check("skok do sekcji: nagłówek pod paskiem, nie za nim", r.top != null && r.top >= r.tb - 1, JSON.stringify(r));
  } else check("skok do sekcji: są skróty sekcji", false);

  // Edycja: zaznaczenie tekstu przywraca pasek (formatowanie)
  await page.locator('.mode-btn[data-mode="edit"]').tap();
  await page.waitForTimeout(500);
  s = await state();
  check("Edycja: odstęp nadąża za wyższym paskiem", Math.abs(s.pad - s.tbH) <= 1, `${s.pad} vs ${s.tbH}`);
  await wheel(600);
  check("Edycja: przewinięcie w dół chowa pasek", (await state()).tucked);
  await page.evaluate(() => {
    const vp = document.getElementById("docViewport").getBoundingClientRect();
    const p = [...document.querySelectorAll(".docx-editable-p")].find((el) => { const r = el.getBoundingClientRect(); return r.top > vp.top + 80 && r.bottom < vp.bottom && el.textContent.length > 20; });
    const tn = document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode();
    const range = document.createRange();
    range.setStart(tn, 0); range.setEnd(tn, Math.min(8, tn.length));
    const sel = getSelection(); sel.removeAllRanges(); sel.addRange(range);
  });
  await page.waitForTimeout(300);
  check("Edycja: zaznaczenie tekstu pokazuje pasek", !(await state()).tucked);

  // szeroko: bez zmian
  await page.setViewportSize({ width: 1280, height: 900 });
  await page.waitForTimeout(500);
  s = await state();
  check("szerokie okno: pasek w układzie jak dawniej (bez tb-float, bez odstępu)", !s.float && !s.tucked && s.pad === 0, JSON.stringify(s));
  const back = await page.evaluate(() => ({
    search: document.querySelector("#docToolbar > .doc-toolbar-row:first-child")?.contains(document.getElementById("searchQuery")),
    dirtyInStatus: document.querySelector(".status-bar").contains(document.getElementById("statusDirty")),
    trackInStatus: document.querySelector(".status-bar").contains(document.getElementById("statusTrack")),
    island: document.getElementById("docIsland").hidden,
    status: getComputedStyle(document.querySelector(".status-bar")).display !== "none",
  }));
  check("szeroko: szukanie i wskaźniki wracają na swoje miejsca, wyspa ukryta", back.search && back.dirtyInStatus && back.trackInStatus && back.island && back.status, JSON.stringify(back));

  check("bez błędów strony", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();
}

run().then(() => {
  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`${r.ok ? "✓" : "✗"} ${r.name}${!r.ok && r.detail ? ` — ${r.detail}` : ""}`);
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
