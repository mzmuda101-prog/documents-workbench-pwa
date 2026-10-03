// mobile-playwright.js — mobile viewport: reflow layout, panel sheet, no horizontal overflow.

const {
  createTestPage,
  bootApp,
  loadBuiltinSample,
  assertNoErrors,
} = require("./docx-test-helpers");

const MOBILE_VIEWPORT = { width: 390, height: 844 };

async function run() {
  console.log("\nDocuments Workbench PWA — mobile UX test (390×844)\n");

  const { browser, page, errors } = await createTestPage(MOBILE_VIEWPORT);
  await bootApp(page);

  const isMobile = await page.evaluate(() => document.documentElement.classList.contains("is-mobile"));
  if (!isMobile) throw new Error("Brak klasy is-mobile na viewporcie mobilnym");

  await loadBuiltinSample(page);

  const layout = await page.evaluate(() => {
    const vp = document.getElementById("docViewport");
    const canvas = document.getElementById("docCanvas");
    const section = document.querySelector(".docx-preview-host section.docx") || document.querySelector(".docx-preview-host .docx");
    const host = document.querySelector(".docx-preview-host");
    return {
      hasDocument: document.body.classList.contains("has-document"),
      reflow: canvas?.classList.contains("doc-reflow-mode"),
      viewportScrollW: vp?.scrollWidth || 0,
      viewportClientW: vp?.clientWidth || 0,
      hostW: host?.getBoundingClientRect().width || 0,
      sectionW: section?.getBoundingClientRect().width || 0,
      hasHorizontalOverflow: (vp?.scrollWidth || 0) > (vp?.clientWidth || 0) + 2,
    };
  });

  if (!layout.hasDocument) throw new Error("Brak klasy has-document po wczytaniu");
  console.log("  ✓ Tryb pełnej szerokości (has-document)");

  // PL/EN mieszka w menu ⋯ (paczka B) — musi być osiągalne jednym tapnięciem
  await page.locator("#appMenuBtn").click();
  const langVisible = await page.evaluate(() => {
    const menu = document.getElementById("appMenu");
    const el = document.getElementById("langSwitch");
    const r = el?.getBoundingClientRect();
    return !!menu && !menu.hidden && menu.contains(el) && !!r && r.width > 0 && r.right <= window.innerWidth;
  });
  if (!langVisible) throw new Error("Przełącznik PL/EN niedostępny w menu ⋯");
  await page.keyboard.press("Escape");
  console.log("  ✓ PL/EN dostępne w menu ⋯");

  if (layout.hasHorizontalOverflow) {
    throw new Error(`Poziomy overflow: scroll=${layout.viewportScrollW} client=${layout.viewportClientW}`);
  }
  const widthUse = layout.hostW / layout.viewportClientW;
  if (widthUse < 0.96) {
    throw new Error(`Za wąski podgląd: host ${Math.round(layout.hostW)}px / viewport ${layout.viewportClientW}px (${Math.round(widthUse * 100)}%)`);
  }
  console.log(`  ✓ Pełna szerokość (${Math.round(layout.hostW)}px / ${layout.viewportClientW}px, reflow=${layout.reflow})`);

  await page.locator("#panelToggle").click();
  await page.waitForTimeout(250);
  const panelOpen = await page.evaluate(() => document.documentElement.classList.contains("sidebar-open"));
  if (!panelOpen) throw new Error("Panel nie otworzył się na mobile");
  console.log("  ✓ Bottom sheet panel otwarty");

  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
  await page.waitForTimeout(200);
  const panelClosed = await page.evaluate(() => !document.documentElement.classList.contains("sidebar-open"));
  if (!panelClosed) throw new Error("Panel nie zamknął się");
  console.log("  ✓ Panel zamknięty");

  // Widok mobilny: suwak zostaje (skaluje tekst), w panelu wybór Auto / mobilny / desktopowy
  const viewUi = await page.evaluate(() => ({
    slider: !document.getElementById("zoomSliderField")?.classList.contains("hidden"),
    opts: [...document.querySelectorAll("#viewLayoutPanel .view-layout-opt")].map((b) => `${b.dataset.layout}:${b.getAttribute("aria-checked")}`).join(","),
  }));
  if (!viewUi.slider) throw new Error("Suwak zoom powinien być widoczny także w Widoku mobilnym");
  if (viewUi.opts !== "auto:true,mobile:false,desktop:false") throw new Error(`Wybór widoku w panelu: ${viewUi.opts}`);
  console.log("  ✓ Suwak + wybór widoku (Auto zaznaczone) w panelu Widok");

  // Brakujący krój (Calibri/Segoe UI na telefonie): dopisany rodzaj zapasowy i PRAWDZIWE
  // pogrubienie (Android: odmiana bold zastępnika trafiała w zwykły plik → bold znikał).
  const fonts = await page.evaluate(async () => {
    const z = await JSZip.loadAsync(originalFileBytes);
    let doc = await z.file("word/document.xml").async("string");
    const run = (font, bold, txt) => `<w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>${bold ? "<w:b/>" : ""}<w:sz w:val="40"/></w:rPr><w:t>${txt}</w:t></w:r>`;
    const para = ["Calibri", "Segoe UI", "Cambria"].map((f) => `<w:p>${run(f, false, "Wwmmaaoo")}</w:p><w:p>${run(f, true, "Wwmmaaoo")}</w:p>`).join("");
    doc = doc.replace("<w:body>", `<w:body>${para}`);
    z.file("word/document.xml", doc);
    const bytes = await z.generateAsync({ type: "uint8array" });
    await ingestFile(new File([bytes], "fonty.docx"), { silent: true });
    await document.fonts.ready;
    const spans = [...document.querySelectorAll(".docx-preview-host span")].filter((sp) => sp.textContent === "Wwmmaaoo").slice(0, 6);
    const ink = (el) => {
      const cs = getComputedStyle(el);
      const c = document.createElement("canvas");
      c.width = 400; c.height = 60;
      const x = c.getContext("2d");
      x.font = `${cs.fontWeight} 40px ${cs.fontFamily}`;
      x.fillText(el.textContent, 4, 44);
      const d = x.getImageData(0, 0, 400, 60).data;
      let n = 0;
      for (let i = 3; i < d.length; i += 4) n += d[i];
      return n;
    };
    return spans.map((sp) => ({ family: getComputedStyle(sp).fontFamily, weight: getComputedStyle(sp).fontWeight, ink: ink(sp) }));
  });
  const expectGeneric = ["sans-serif", "sans-serif", "sans-serif", "sans-serif", "serif", "serif"];
  if (fonts.length !== 6 || fonts.some((f, i) => !f.family.endsWith(`, ${expectGeneric[i]}`))) throw new Error(`Rodzaj zapasowy czcionki: ${JSON.stringify(fonts)}`);
  for (let i = 0; i < 6; i += 2) {
    if (!(fonts[i + 1].ink > fonts[i].ink * 1.12)) throw new Error(`Pogrubienie niewidoczne (${fonts[i].family}): ${fonts[i].ink} → ${fonts[i + 1].ink}`);
  }
  console.log("  ✓ Brakujące kroje: rodzaj zapasowy (sans-serif/serif) i widoczne pogrubienie");

  assertNoErrors(errors, "mobile-playwright");
  await browser.close();
  console.log("  ✓ Brak błędów JS w konsoli\n");
}

run().catch((err) => {
  console.error("❌  mobile-playwright failed:", err.message || err);
  process.exit(1);
});
