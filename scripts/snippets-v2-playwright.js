// snippets-v2-playwright.js — Snippety klikane i pisane jak użytkownik (Chromium + WebKit):
// podpowiedzi po „!”, pytanie o pola {{…}}, {cursor}, daty z przesunięciem, snippet w snippecie,
// polskie nazwy, tryb auto, „Rozwiń w dokumencie” z wieloma wierszami (<w:br/>), cofanie.
//
//   node scripts/snippets-v2-playwright.js
//   ENGINE=webkit node scripts/snippets-v2-playwright.js

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const pad = (n) => String(n).padStart(2, "0");
const short = (d) => `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()}`;

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1366, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.clear(); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(`${APP_URL}?sample=headings-sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForSelector(".docx-preview-host p", { timeout: 20000 });
  await page.waitForTimeout(400);

  const fileParas = () => page.evaluate(async () => {
    await waitInlineStructuralIdle?.();
    const z = await JSZip.loadAsync(await buildDocumentForSave());
    const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
    return collectParagraphElements(d.documentElement, "all").map((p) => ({ text: paragraphSearchText(p), xml: new XMLSerializer().serializeToString(p) }));
  });
  const clickEndOf = async (i) => {
    const box = await page.evaluate((i) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
      p.scrollIntoView({ block: "center" });
      const r = document.createRange(); r.selectNodeContents(p);
      const rects = r.getClientRects(); const last = rects[rects.length - 1];
      return { x: Math.min(last.right + 4, p.getBoundingClientRect().right - 1), y: last.top + last.height / 2 };
    }, i);
    await page.mouse.click(box.x, box.y);
  };
  const suggest = () => page.evaluate(() => {
    const b = document.querySelector(".sn-suggest");
    return { open: !b.hidden, items: [...b.querySelectorAll(".sn-suggest-item strong")].map((e) => e.textContent) };
  });

  // ── zestaw startowy ─────────────────────────────────────────────────────────
  await page.evaluate(() => { if (!document.documentElement.classList.contains("sidebar-open")) document.getElementById("panelToggle").click(); });
  await page.waitForTimeout(250);
  await page.evaluate(() => { const d = document.getElementById("panel-snippets"); d.scrollIntoView(); });
  await page.click("#panel-snippets > summary");
  await page.waitForTimeout(400);
  await page.click("#snStarterBtn");
  await page.waitForTimeout(150);
  const rows = await page.$$eval("#snList .snippet-row .snippet-name", (e) => e.map((x) => x.textContent));
  check("pusty panel: „Dodaj przykładowe snippety” → 5 snippetów z podglądem", rows.length === 5 && rows.includes("!pozdrawiam"), rows.join(" "));

  // ── podpowiedzi po „!” + pytanie o pola ─────────────────────────────────────
  await page.click('.mode-btn[data-mode="edit"]');
  await page.waitForTimeout(400);
  await clickEndOf(1);
  await page.keyboard.type(" !po");
  await page.waitForTimeout(200);
  let sg = await suggest();
  check("po wpisaniu „!po” lista podpowiedzi z !pozdrawiam", sg.open && sg.items[0] === "!pozdrawiam", JSON.stringify(sg));
  await page.keyboard.press("Enter");
  await page.waitForSelector("dialog.sn-dialog[open]", { timeout: 3000 }).catch(() => {});
  const dlg = await page.$("dialog.sn-dialog[open] input");
  check("snippet z {{imie_nazwisko}} pyta o wartość w oknie", !!dlg);
  if (dlg) { await dlg.fill("Jan Nowak"); await page.keyboard.press("Enter"); }
  await page.waitForTimeout(400);
  let f = await fileParas();
  check("Enter wstawia snippet: „Z poważaniem” + łamanie wiersza + „Jan Nowak” (w pliku <w:br/>)",
    /Struktura\. Z poważaniem<\/w:t>(<\/w:r><w:r>)?<w:br\/>(<\/w:r><w:r>)?<w:t[^>]*>Jan Nowak/.test(f[1].xml) || (/Z poważaniem/.test(f[1].text) && /<w:br\/>/.test(f[1].xml) && /Jan Nowak/.test(f[1].text) && !/!po/.test(f[1].text)),
    f[1].xml.slice(-240));
  check("Enter nie tworzy przy tym nowego akapitu", f.length === 32, String(f.length));
  await page.click("#undoBtn");
  await page.waitForTimeout(700);
  f = await fileParas();
  check("↶ cofa wstawienie snippetu (zostaje wpisane „!po”)", / !po$/.test(f[1].text) && !/poważaniem/.test(f[1].text), f[1].text.slice(-30));

  // ── wybór z listy kliknięciem/stuknięciem; „kliknięcie-duch” iOS nie trafia w ↶ pod listą ──
  await clickEndOf(5);
  await page.keyboard.type(" !dz");
  await page.waitForTimeout(200);
  const itemBox = await page.evaluate(() => { const r = document.querySelector(".sn-suggest-item").getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });
  await page.mouse.click(itemBox.x, itemBox.y);
  // iOS: po zniknięciu listy w to samo miejsce przychodzi jeszcze „click” — symulujemy go
  await page.evaluate(({ x, y }) => { const el = document.elementFromPoint(x, y); el?.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, clientX: x, clientY: y })); }, itemBox);
  await page.waitForTimeout(400);
  f = await fileParas();
  check("stuknięcie w podpowiedź wstawia snippet i nic pod listą nie dostaje kliknięcia", f[5].text.endsWith(` ${short(new Date())}`), f[5].text.slice(-20));

  // ── Tab + data z przesunięciem ──────────────────────────────────────────────
  await clickEndOf(3);
  await page.keyboard.type(" !ter");
  await page.waitForTimeout(200);
  await page.keyboard.press("Tab");
  await page.waitForTimeout(300);
  f = await fileParas();
  const d14 = new Date(); d14.setDate(d14.getDate() + 14);
  check("Tab wstawia !termin14 = dziś + 14 dni", f[3].text.endsWith(` ${short(d14)}`), `${f[3].text.slice(-20)} vs ${short(d14)}`);

  // ── własny snippet: {cursor}, polska nazwa, snippet w snippecie ──────────────
  await page.fill("#snName", "zwrotŁ");
  await page.fill("#snBody", "Szanowny Panie {cursor}, dnia !dzis");
  await page.click("#snSaveBtn");
  await page.waitForTimeout(200);
  check("nazwa z polską literą (!zwrotŁ) zapisuje się", await page.evaluate(() => loadSnippets().some((s) => s.name === "zwrotŁ")));
  await clickEndOf(4);
  await page.keyboard.type(" !zwr");
  await page.waitForTimeout(200);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(300);
  await page.keyboard.type("Kowalski");
  await page.waitForTimeout(300);
  f = await fileParas();
  check("{cursor}: kursor staje w miejscu znacznika (dopisane „Kowalski” przed przecinkiem)", f[4].text.endsWith(`Szanowny Panie Kowalski, dnia ${short(new Date())}`), f[4].text.slice(-50));
  check("snippet w snippecie (!dzis w treści) rozwinięty", !/!dzis/.test(f[4].text));

  // ── Esc zamyka listę, zwykły Enter dalej robi nowy akapit ───────────────────
  await clickEndOf(6);
  await page.keyboard.type(" !d");
  await page.waitForTimeout(200);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(100);
  sg = await suggest();
  const stillEditing = await page.evaluate(() => !!docCaretParagraph(document.activeElement));
  check("Esc zamyka podpowiedzi, a kursor zostaje w tekście", !sg.open && stillEditing, JSON.stringify({ sg, stillEditing }));
  await page.keyboard.type("x");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(500);
  f = await fileParas();
  check("bez podpowiedzi Enter robi nowy akapit (jak zawsze)", f.length === 33, String(f.length));

  // ── anulowanie okna pól zostawia wpisany trigger ────────────────────────────
  await clickEndOf(8);
  await page.keyboard.type(" !adr");
  await page.waitForTimeout(200);
  await page.keyboard.press("Enter");
  await page.waitForSelector("dialog.sn-dialog[open]", { timeout: 3000 }).catch(() => {});
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  f = await fileParas();
  check("Anuluj w oknie pól: nic nie wstawione, „!adr” zostaje", /!adr$/.test(f[8].text), f[8].text.slice(-20));

  // ── tryb auto: „!dzis” + spacja ─────────────────────────────────────────────
  await page.selectOption("#snExpandMode", "auto");
  await clickEndOf(10);
  await page.keyboard.type(" !dzis ");
  await page.waitForTimeout(300);
  f = await fileParas();
  check("tryb auto: „!dzis ” → dzisiejsza data", f[10].text.endsWith(` ${short(new Date())} `), f[10].text.slice(-20));
  await page.selectOption("#snExpandMode", "manual");

  // ── „Rozwiń w dokumencie”: wiele wierszy = <w:br/>, „Uwaga!pozdrawiam” to nie trigger ──
  await page.fill("#snName", "blok");
  await page.fill("#snBody", "Linia A\nLinia B");
  await page.click("#snSaveBtn");
  await clickEndOf(12);
  await page.keyboard.type(" !blok i Uwaga!blok");
  await page.waitForTimeout(200);
  await page.keyboard.press("Escape");
  await page.click("#snExpandBtn");
  await page.waitForTimeout(800);
  f = await fileParas();
  check("„Rozwiń”: wielowierszowy snippet zapisany z <w:br/> (nie „\\n” = spacja w Wordzie)", /Linia A<\/w:t><w:br\/><w:t[^>]*>Linia B/.test(f[12].xml) && !/Linia A\n/.test(f[12].xml), f[12].xml.slice(-200));
  check("„Uwaga!blok” (w środku słowa) nie jest rozwijane", /Uwaga!blok/.test(f[12].text), f[12].text.slice(-30));

  // ── iOS: klawiatura zjada spację przed „!” (zgłoszenie 2026-10-02) ───────────
  // Po słowie z paska podpowiedzi iOS dokleja spację, a „!” ją kasuje („Dobrze !” → „Dobrze!”).
  // Symulacja tej zmiany (beforeinput „!” + podmiana spacji na „!” + input), potem nazwa snippetu.
  const iosBang = () => page.evaluate(() => {
    const sel = getSelection();
    const tn = sel.anchorNode;
    const el = tn.parentElement.closest(".docx-editable-p");
    el.dispatchEvent(new InputEvent("beforeinput", { inputType: "insertText", data: "!", bubbles: true }));
    const at = sel.anchorOffset;
    tn.textContent = `${tn.textContent.slice(0, at - 1)}!${tn.textContent.slice(at)}`;
    const r = document.createRange(); r.setStart(tn, at); r.collapse(true); sel.removeAllRanges(); sel.addRange(r);
    el.dispatchEvent(new InputEvent("input", { inputType: "insertText", data: "!", bubbles: true }));
  });
  await clickEndOf(14);
  await page.keyboard.type(" Dobrze ");
  await iosBang();
  await page.keyboard.type("bl");
  await page.waitForTimeout(250);
  const iosSug = await suggest();
  check("iOS zjadł spację przed „!”: „Dobrze!bl” i tak podpowiada !blok", iosSug.open && iosSug.items.includes("!blok"), JSON.stringify(iosSug));
  await page.keyboard.press("Enter");
  await page.waitForTimeout(300);
  f = await fileParas();
  check("…a po wstawieniu spacja wraca: „Dobrze Linia A”", /Dobrze Linia A/.test(f[14].text), f[14].text.slice(-30));
  await page.keyboard.type(" Super ");
  await iosBang();
  await page.keyboard.type(" Dalej");
  await page.waitForTimeout(250);
  f = await fileParas();
  check("zwykłe „Super!” na końcu zdania zostaje bez zmian (bez podpowiedzi)", /Super! Dalej$/.test(f[14].text) && !(await suggest()).open, f[14].text.slice(-20));

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ snippety v2 [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
