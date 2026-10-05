// explore-playwright.js — „przeklikanie” całej aplikacji w poszukiwaniu błędów (2026-10-05).
//
// Przewodnik otwarty, tryb Edycja, kursor w tekście. Każdy widoczny przycisk / pole wyboru na
// pasku, w panelach (wszystkie rozwinięte) i w menu ⋯ jest klikany po kolei; po każdym: Esc,
// sprawdzenie błędów strony (wyjątki, odrzucone obietnice, console.error) i tego, że dokument
// dalej ma treść. Pomijane: rzeczy wychodzące poza stronę (druk, odświeżenie, pełny ekran,
// inne aplikacje, zamknięcie dokumentu, zapis do pliku).
// Użycie: node scripts/explore-playwright.js   |   ENGINE=webkit   |   PHONE=1 (390×844, dotyk)

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const PHONE = !!process.env.PHONE;
const SKIP = /sidebarCloseBtn|panelToggle|focusMode|tbPanelBtn|brandRefresh|printPreview|fullscreen|closeDoc|saveBtn|saveAsBtn|heroSaveBtn|appUpdateBtn|loadBtn|photoDocBtn|fileInput|photoInput|phImportFile|snImportFile|exportDocx|download|print/i;
const SKIP_TEXT = /Drukuj|Odśwież|Pełny ekran|Zamknij dokument|Zapisz|Sheet Workbench|Otwórz plik|Zdjęcie dokumentu|Pobierz|Eksport/i;

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: PHONE ? { width: 390, height: 844 } : { width: 1360, height: 900 }, isMobile: PHONE && ENGINE === "chromium", hasTouch: PHONE });
  await context.addInitScript(() => {
    sessionStorage.setItem("introPlayed", "true");
    window.print = () => {};
    window.addEventListener("unhandledrejection", (e) => console.error("UNHANDLED " + (e.reason?.stack || e.reason)));
  });
  const page = await context.newPage();
  const problems = [];
  let current = "start";
  page.on("pageerror", (e) => problems.push({ at: current, msg: "WYJĄTEK " + e.message }));
  page.on("console", (m) => { if (m.type() === "error" && !/favicon|ResizeObserver|Failed to load resource/.test(m.text())) problems.push({ at: current, msg: m.text().slice(0, 300) }); });
  page.on("dialog", (d) => d.dismiss().catch(() => {}));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => loadSampleDocument("przewodnik"));
  await page.waitForSelector(".docx-preview-host p", { timeout: 30000 });
  await page.waitForTimeout(1500);
  await page.evaluate(() => { if (readOnlyMode) { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); } });
  await page.waitForTimeout(800);
  const caret = () => page.evaluate(() => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).filter((p) => p.isContentEditable && p.textContent.length > 20);
    const p = ps[Math.min(5, ps.length - 1)];
    if (p) placeCaret(p, 5);
  }).catch(() => {});
  await caret();

  // lista kontrolek: pasek nad dokumentem, panel (wszystkie sekcje), menu ⋯
  const controls = await page.evaluate((skipSrc) => {
    const skip = new RegExp(skipSrc, "i");
    document.querySelectorAll("details.panel").forEach((d) => (d.open = true));
    const sel = "button, select, [role=button], summary";
    const all = [...document.querySelectorAll(`.doc-toolbar ${sel}, #formatToolbar ${sel}, .sidebar ${sel}, .doc-head ${sel}, .toolbar ${sel}`)];
    const seen = new Set();
    return all.map((el, i) => {
      el.dataset.exploreId = String(i);
      const label = (el.id || "") + " " + (el.getAttribute("aria-label") || el.textContent || "").trim().replace(/\s+/g, " ").slice(0, 40);
      return { id: String(i), label, tag: el.tagName.toLowerCase(), skip: skip.test(el.id || "") || el.disabled };
    }).filter((c) => !c.skip && !seen.has(c.label) && seen.add(c.label));
  }, SKIP.source);

  let clicked = 0;
  for (const c of controls) {
    if (SKIP_TEXT.test(c.label)) continue;
    current = c.label;
    const vis = await page.evaluate((id) => {
      const el = document.querySelector(`[data-explore-id="${id}"]`);
      if (!el) return false;
      // telefon/tablet: panel to szuflada — otwórz ją przed klikaniem w panel, zamknij przed paskiem
      if (typeof setSidebarOpen === "function") setSidebarOpen(!!el.closest(".sidebar"));
      el.closest("details.panel")?.setAttribute("open", "");
      el.scrollIntoView({ block: "center" });
      const r = el.getBoundingClientRect();
      return r.width > 0 && r.height > 0 && getComputedStyle(el).visibility !== "hidden" && !el.closest("[hidden]");
    }, c.id).catch(() => false);
    await page.waitForTimeout(PHONE ? 350 : 0); // animacja szuflady
    if (!vis) continue;
    try {
      if (c.tag === "select") {
        const vals = await page.evaluate((id) => [...document.querySelector(`[data-explore-id="${id}"]`).options].filter((o) => !o.disabled && !o.hidden && o.value && o.value !== "__custom").map((o) => o.value).slice(0, 3), c.id);
        for (const v of vals) await page.selectOption(`[data-explore-id="${c.id}"]`, v, { timeout: 2000 }).catch(() => {});
      } else {
        await page.click(`[data-explore-id="${c.id}"]`, { timeout: 2000 });
      }
      clicked++;
      await page.waitForTimeout(250);
      // otwarte okienko/menu — kliknij jego pierwszy zwykły przycisk (bez „Usuń”), potem Esc
      await page.keyboard.press("Escape").catch(() => {});
      await page.waitForTimeout(120);
      await page.keyboard.press("Escape").catch(() => {});
    } catch (e) {
      const why = (String(e.message).match(/<[^>]+> (?:from <[^>]+> subtree )?intercepts pointer events/) || [""])[0];
      problems.push({ at: c.label, msg: "KLIK NIEUDANY: " + (why || String(e.message).split("\n")[0]) });
    }
    const alive = await page.evaluate(() => !!document.querySelector(".docx-preview-host p") && !!originalFileBytes).catch(() => false);
    if (!alive) { problems.push({ at: c.label, msg: "DOKUMENT ZNIKNĄŁ po kliknięciu" }); break; }
    await caret();
  }
  // menu ⋯
  await page.click("#appMenuBtn").catch(() => {});
  await page.waitForTimeout(300);
  const menuItems = await page.evaluate((skipSrc) => [...document.querySelectorAll("#appMenu .app-menu-item, #appMenu button")].filter((b) => !new RegExp(skipSrc, "i").test(b.id || "") && b.offsetParent).map((b, i) => { b.dataset.exploreMenu = String(i); return { i: String(i), label: (b.id || "") + " " + b.textContent.trim().slice(0, 30) }; }), SKIP.source);
  await page.keyboard.press("Escape");
  for (const m of menuItems) {
    if (SKIP_TEXT.test(m.label)) continue;
    current = "menu ⋯ " + m.label;
    await page.click("#appMenuBtn").catch(() => {});
    await page.waitForTimeout(200);
    await page.click(`[data-explore-menu="${m.i}"]`, { timeout: 2000 }).catch((e) => problems.push({ at: current, msg: "KLIK NIEUDANY " + String(e.message).split("\n")[0] }));
    clicked++;
    await page.waitForTimeout(400);
    await page.keyboard.press("Escape").catch(() => {});
    await page.keyboard.press("Escape").catch(() => {});
  }
  await browser.close();
  console.log(`[${ENGINE}${PHONE ? " telefon" : ""}] kliknięto ${clicked} kontrolek, problemów: ${problems.length}`);
  for (const p of problems) console.log(`❌ ${p.at.trim()} → ${p.msg}`);
  process.exit(problems.length ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
