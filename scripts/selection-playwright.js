// selection-playwright.js — zaznaczanie przez kilka akapitów (doc-selection.js) i krój/rozmiar
// w miejscu kursora (docx-inline-edit.js), jak w Wordzie.
//
// Fixture: docs/samples/selection-sample.docx (node scripts/gen-selection-docx.mjs).
// Dawniej każdy akapit był osobnym polem edycji: przeciąganie kończyło się na granicy akapitu,
// Ctrl+A nie zaznaczał nic; rozmiar ustawiony bez zaznaczenia „przyklejał się” do pisania
// w innych akapitach, a lista rozmiarów zostawała przy starej wartości.
// Każda zmiana sprawdzana też w ZAPISANYM pliku (teksty akapitów pliku = podgląd).
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = ENGINE === "webkit" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1400, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); localStorage.setItem("dwb-panel-docked-open-v1", "0"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.dismiss());
  await page.goto(`${APP_URL}?sample=selection-sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForFunction(() => !readOnlyMode && document.querySelector(".docx-edit-root"), null, { timeout: 10000 });
  await sleep(500);

  const idle = async () => {
    await page.evaluate(() => waitInlineStructuralIdle());
    await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 });
    await sleep(350);
  };
  const texts = () => page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).map((p) => previewRunsToPlainText(extractRunsFromPreviewParagraph(p))));
  const fileXml = async () => {
    const b64 = await page.evaluate(async () => {
      const bytes = await buildDocumentForSave();
      let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s);
    });
    const zip = await JSZip.loadAsync(Buffer.from(b64, "base64"));
    return zip.file("word/document.xml").async("string");
  };
  const fileTexts = (xml) => (xml.match(/<w:p[ >][\s\S]*?<\/w:p>|<w:p\/>/g) || []).map((p) => (p.match(/<w:t[^>]*>[^<]*/g) || []).map((x) => x.replace(/<w:t[^>]*>/, "")).join("").replace(/&quot;/g, '"').replace(/&amp;/g, "&"));
  const idx = (needle) => page.evaluate((n) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).findIndex((p) => p.textContent.includes(n)), needle);
  // zaznaczenie od (akapit a, znak oa) do (akapit b, znak ob) — przesunięcia w tekście akapitu
  const select = (a, oa, b, ob) => page.evaluate(([a, oa, b, ob]) => {
    const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
    const s = resolveTextPosition(ps[a], oa) || { node: ps[a], offset: 0 };
    const e = resolveTextPosition(ps[b], ob) || { node: ps[b], offset: ps[b].childNodes.length };
    docEditRoot().focus({ preventScroll: true });
    const r = document.createRange(); r.setStart(s.node, s.offset); r.setEnd(e.node, e.offset);
    const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
    return sel.toString();
  }, [a, oa, b, ob]);
  const undo = async () => { await page.keyboard.press(`${MOD}+KeyZ`); await idle(); };
  const consistent = async (label) => {
    const prev = await texts();
    const file = fileTexts(await fileXml());
    check(`${label}: plik = podgląd (${prev.length} akapitów)`, JSON.stringify(prev) === JSON.stringify(file), `podgląd ${JSON.stringify(prev)}\nplik ${JSON.stringify(file)}`);
    return prev;
  };

  const orig = await texts();
  check("fixture: 14 akapitów (nagłówek, tabela 2×2, podział strony)", orig.length === 14 && orig[0] === "Rozdział pierwszy", JSON.stringify(orig));
  check("cały dokument to jedno pole edycji (akapity bez własnego contenteditable)", await page.evaluate(() => docEditRoot().isContentEditable && !document.querySelector(".docx-edit-root p.docx-editable-p[contenteditable]")));

  // ── przeciąganie myszą przez 3 akapity ──
  const iAlfa = await idx("Alfa"), iGamma = await idx("Gamma");
  const pos = (i, x) => page.evaluate(([i, x]) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; p.scrollIntoView({ block: "center" }); const r = p.getBoundingClientRect(); return { x: r.left + x, y: r.top + r.height / 2 }; }, [i, x]);
  let a = await pos(iAlfa, 40);
  await page.mouse.click(a.x, a.y);
  await sleep(150);
  a = await pos(iAlfa, 40);
  const g = await pos(iGamma, 60);
  await page.mouse.move(a.x, a.y);
  await page.mouse.down();
  await page.mouse.move(g.x, g.y, { steps: 12 });
  await page.mouse.up();
  const dragSel = await page.evaluate(() => getSelection().toString());
  check("przeciąganie myszą zaznacza przez kilka akapitów", /Beta drugi akapit zwykłego tekstu\./.test(dragSel) && /Gamma/.test(dragSel), JSON.stringify(dragSel));

  // ── Ctrl/⌘+A ──
  await page.keyboard.press(`${MOD}+KeyA`);
  const all = await page.evaluate(() => getSelection().toString());
  check("Ctrl/⌘+A zaznacza całą treść (od nagłówka do ostatniego akapitu)", all.startsWith("Rozdział pierwszy") && /Eta ostatni akapit\.\s*$/.test(all), JSON.stringify(all.slice(0, 40) + "…" + all.slice(-30)));

  // ── Shift+↓ wydłuża zaznaczenie do następnego akapitu ──
  const iBeta = await idx("Beta");
  await page.evaluate((i) => focusParagraphAtOffset(i, 0), iBeta);
  await page.keyboard.press("Shift+ArrowDown");
  await page.keyboard.press("Shift+ArrowDown");
  const shiftSel = await page.evaluate(() => getSelection().toString());
  check("Shift+↓ zaznacza przez granicę akapitu", /^Beta drugi akapit zwykłego tekstu\.\s+Gamma/.test(shiftSel), JSON.stringify(shiftSel));

  // ── pisanie w miejsce zaznaczenia przez 3 akapity ──
  await select(iAlfa, 7, iGamma, 6);
  await page.keyboard.type("X");
  await idle();
  let now = await texts();
  check("pisanie zastępuje zaznaczenie: 3 akapity → 1, tekst sklejony", now.length === 12 && now[iAlfa] === "Alfa piXtrzeci akapit zwykłego tekstu.", JSON.stringify(now.slice(0, 4)));
  await consistent("pisanie w miejsce zaznaczenia");
  await undo();
  check("Cofnij przywraca 3 akapity", JSON.stringify(await texts()) === JSON.stringify(orig), JSON.stringify((await texts()).slice(0, 5)));

  // ── cały nagłówek + część następnego: zostaje formatowanie ostatniego ──
  await select(0, 0, iAlfa, 5);
  await page.keyboard.press("Backspace");
  await idle();
  now = await texts();
  const headInfo = await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0]; return { cls: p.className, fs: getComputedStyle(p).fontSize }; });
  check("usunięty cały nagłówek: następny tekst NIE staje się nagłówkiem (jak w Wordzie)", now[0] === "pierwszy akapit zwykłego tekstu." && !/Heading1/i.test(headInfo.cls), JSON.stringify({ t: now[0], ...headInfo }));
  const xmlD = await fileXml();
  check("…w pliku też bez stylu Nagłówek 1", !/Heading1/.test(xmlD.split("</w:p>")[0]), xmlD.split("</w:p>")[0].slice(-200));
  await consistent("usunięcie nagłówka");
  await undo();

  // ── trzykrotne kliknięcie + pisanie: zastępuje tekst akapitu, akapit zostaje ──
  const al = await pos(iAlfa, 30);
  await page.mouse.click(al.x, al.y, { clickCount: 3 });
  await page.keyboard.type("Nowy");
  await idle();
  now = await texts();
  check("trzykrotne kliknięcie + pisanie: zastąpiony tekst akapitu, liczba akapitów ta sama", now.length === 14 && now[iAlfa] === "Nowy" && now[iBeta] === orig[iBeta], JSON.stringify(now.slice(0, 4)));
  await undo();

  // ── tabela w środku zaznaczenia znika w całości ──
  const iDelta = await idx("Delta"), iEps = await idx("Epsilon");
  await select(iDelta, 5, iEps, 9);
  await page.keyboard.press("Delete");
  await idle();
  now = await texts();
  const xmlT = await fileXml();
  check("zaznaczenie z całą tabelą w środku: tabela usunięta (podgląd i plik)", now[iDelta] === "Deltao tabeli." && !(await page.evaluate(() => !!document.querySelector(".docx-preview-host table"))) && !/<w:tbl>/.test(xmlT), JSON.stringify(now));
  await consistent("usunięcie z tabelą");
  await undo();

  // ── przez podział strony (dwie kartki podglądu) ──
  const iZeta = await idx("Zeta");
  await select(iEps, 8, iZeta, 5);
  await page.keyboard.press("Delete");
  await idle();
  now = await texts();
  const xmlP = await fileXml();
  const sections = await page.evaluate(() => document.querySelectorAll(".docx-preview-host section.docx").length);
  check("przez podział strony: akapity sklejone, podział usunięty, jedna kartka", now.includes("Epsilon na drugiej stronie.") && !/w:type="page"/.test(xmlP) && sections === 1, JSON.stringify({ now, sections }));
  await consistent("usunięcie przez podział strony");
  await undo();

  // ── komórki jednej tabeli: tekst czyszczony, tabela zostaje ──
  const iA1 = await idx("A1 komórka"), iB2 = await idx("B2 komórka");
  await select(iA1, 2, iB2, 2);
  await page.keyboard.press("Delete");
  await idle();
  now = await texts();
  check("zaznaczenie przez komórki: tekst usunięty, tabela 2×2 zostaje", now.length === 14 && now[iA1] === "A1" && now[iA1 + 1] === "" && now[iB2] === " komórka", JSON.stringify(now.slice(iA1, iB2 + 1)));
  await consistent("czyszczenie komórek");
  await undo();

  // ── z treści do środka tabeli: odmowa z wyjaśnieniem, nic nie znika ──
  await select(iDelta, 2, iA1, 2);
  await page.keyboard.press("Delete");
  await idle();
  check("zaznaczenie z tekstu do części tabeli: nic nie znika (komunikat)", JSON.stringify(await texts()) === JSON.stringify(orig) && await page.evaluate(() => [...document.querySelectorAll(".toast")].some((x) => /część tabeli/.test(x.textContent))));

  // ── Enter w miejsce zaznaczenia ──
  await select(iAlfa, 4, iBeta, 4);
  await page.keyboard.press("Enter");
  await idle();
  now = await texts();
  check("Enter w miejsce zaznaczenia: usuwa i dzieli w miejscu", now[iAlfa] === "Alfa" && now[iAlfa + 1] === " drugi akapit zwykłego tekstu." && now.length === 14, JSON.stringify(now.slice(0, 4)));
  await consistent("Enter w miejsce zaznaczenia");
  await undo();

  // ── wytnij ──
  await select(iBeta, 5, iGamma, 6);
  const cut = await page.evaluate(() => {
    const dt = new DataTransfer();
    const ev = new ClipboardEvent("cut", { clipboardData: dt, bubbles: true, cancelable: true });
    docEditRoot().dispatchEvent(ev);
    return { text: dt.getData("text/plain"), html: dt.getData("text/html").length };
  });
  await idle();
  now = await texts();
  check("Wytnij przez akapity: do schowka tekst, akapity sklejone", /drugi akapit zwykłego tekstu\.\s+Gamma/.test(cut.text) && cut.html > 0 && now[iBeta] === "Beta trzeci akapit zwykłego tekstu.", JSON.stringify({ cut, t: now[iBeta] }));
  await consistent("wytnij");
  await undo();

  // ── klawiatura ekranowa (Android): Enter / Backspace tylko jako beforeinput ──
  await page.evaluate((i) => focusParagraphAtOffset(i, 4), iAlfa);
  await page.evaluate(() => docEditRoot().dispatchEvent(new InputEvent("beforeinput", { inputType: "insertParagraph", bubbles: true, cancelable: true })));
  await idle();
  now = await texts();
  check("Enter z klawiatury ekranowej (sam beforeinput) dzieli akapit", now[iAlfa] === "Alfa" && now[iAlfa + 1] === " pierwszy akapit zwykłego tekstu.", JSON.stringify(now.slice(0, 4)));
  await page.evaluate((i) => focusParagraphAtOffset(i, 0), iAlfa + 1);
  await page.evaluate(() => docEditRoot().dispatchEvent(new InputEvent("beforeinput", { inputType: "deleteContentBackward", bubbles: true, cancelable: true })));
  await idle();
  now = await texts();
  check("Backspace z klawiatury ekranowej na początku akapitu skleja (bez rozbicia podglądu)", now[iAlfa] === orig[iAlfa] && now.length === 14, JSON.stringify(now.slice(0, 4)));
  await consistent("Enter/Backspace z klawiatury ekranowej");

  // ── formatowanie przez kilka akapitów ──
  await select(iAlfa, 0, iGamma, 5);
  await page.keyboard.press(`${MOD}+KeyB`);
  await sleep(200);
  await page.selectOption("#fmtFontSize", "16");
  await sleep(200);
  const xmlF = await fileXml();
  const parasF = xmlF.split("</w:p>");
  const B = /<w:b( w:val="(1|true)")?\/>/;
  const fmtOk = [iAlfa, iBeta].every((i) => B.test(parasF[i]) && /<w:sz w:val="32"\/>/.test(parasF[i])) && /<w:sz w:val="32"\/>/.test(parasF[iGamma]) && !B.test(parasF[iDelta]);
  check("B i rozmiar 16 na zaznaczeniu przez akapity — w pliku w każdym z nich", fmtOk, [iAlfa, iBeta, iGamma].map((i) => parasF[i].replace(/<[^>]*w:t[^>]*>/g, "").slice(-160)).join("\n"));
  check("…liczba akapitów bez zmian, zaznaczenie zostaje", (await texts()).length === 14 && (await page.evaluate(() => getSelection().toString())).startsWith("Alfa"));
  await undo();
  await undo();

  // ── krój i rozmiar w miejscu kursora (pasek i pasek stanu) ──
  const fmt = () => page.evaluate(() => ({ fam: document.getElementById("fmtFontFamily").value, size: document.getElementById("fmtFontSize").value, status: document.getElementById("statusFont").textContent }));
  await page.evaluate(() => focusParagraphAtOffset(0, 3));
  await sleep(150);
  const fHead = await fmt();
  await page.evaluate((i) => focusParagraphAtOffset(i, 3), iBeta);
  await sleep(150);
  const fBeta = await fmt();
  check("rozmiar z dziedziczenia: nagłówek 16, zwykły tekst 11 (pole i pasek stanu)", fHead.size === "16" && fBeta.size === "11" && fBeta.fam === "Arial" && fBeta.status === "Arial · 11 pt", JSON.stringify({ fHead, fBeta }));
  const iMaly = await idx("dziewiątka");
  await page.evaluate((i) => focusParagraphAtOffset(i, 8), iMaly);
  await sleep(150);
  const f9 = await fmt();
  await page.evaluate((i) => focusParagraphAtOffset(i, 37), iMaly);
  await sleep(150);
  const fGeo = await fmt();
  await select(iMaly, 0, iMaly, 43);
  await sleep(150);
  const fMix = await fmt();
  check("fragment 9 pt / Georgia / zaznaczenie mieszane = puste pole i zakres w pasku stanu", f9.size === "9" && fGeo.fam === "Georgia" && fMix.size === "" && fMix.fam === "" && /9–14 pt/.test(fMix.status) && /różne czcionki/.test(fMix.status), JSON.stringify({ f9, fGeo, fMix }));
  const famList = await page.evaluate(() => [...document.querySelectorAll("#fmtFontFamily optgroup")].map((g) => `${g.label}: ${[...g.children].map((o) => o.value).join(",")}`));
  check("lista krojów: najpierw kroje z dokumentu (bez symboli), potem popularne", /^W tym dokumencie: Arial,Georgia$/.test(famList[0]) && /Calibri/.test(famList[1] || ""), JSON.stringify(famList));

  // Ctrl/⌘+Shift+> — każdy rozmiar o stopień (9→10, 11→12, 14→16)
  await select(iMaly, 0, iMaly, 30);
  await page.keyboard.press(`${MOD}+Shift+Period`);
  await sleep(200);
  const grown = await page.evaluate((i) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; return extractRunsFromPreviewParagraph(p).map((r) => `${r.text.trim()}=${r.fontSize || ""}`).join("|"); }, iMaly);
  check("Ctrl/⌘+Shift+> powiększa każdy rozmiar osobno (Word)", /dziewiątka=10pt/.test(grown) && /czternastka=16pt/.test(grown) && /Mały=12pt/.test(grown), grown);
  await undo();

  // „przyklejony” rozmiar (zgłoszenie): 9 ustawione w jednym zdaniu nie idzie do innego akapitu
  await select(iAlfa, 0, iAlfa, 4);
  await page.selectOption("#fmtFontSize", "9");
  await sleep(150);
  await page.evaluate((i) => focusParagraphAtOffset(i, 4), iBeta);
  await sleep(150);
  const afterMove = await fmt();
  await page.keyboard.type("Q");
  await sleep(150);
  const qSize = await page.evaluate((i) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; return extractRunsFromPreviewParagraph(p).find((r) => r.text.includes("Q"))?.fontSize || "dziedziczony"; }, iBeta);
  check("po zmianie na 9 w jednym akapicie: w innym pole pokazuje 11, pisanie bez 9", afterMove.size === "11" && qSize === "dziedziczony", JSON.stringify({ afterMove, qSize }));
  // rozmiar bez zaznaczenia: tylko w tym miejscu kursora
  await page.selectOption("#fmtFontSize", "20");
  await sleep(100);
  check("rozmiar bez zaznaczenia: pole od razu pokazuje 20", (await fmt()).size === "20");
  await page.keyboard.type("W");
  await page.keyboard.press("ArrowLeft");
  await page.keyboard.press("ArrowLeft");
  await page.keyboard.type("v");
  await sleep(150);
  const runsB = await page.evaluate((i) => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]; return extractRunsFromPreviewParagraph(p).map((r) => `${r.text}=${r.fontSize || "-"}`).join("|"); }, iBeta);
  check("…wpisane tam „W” ma 20 pt, a po przestawieniu kursora „v” już nie", /W=20pt/.test(runsB) && /v[^=|]*=-/.test(runsB) && !/v=20pt/.test(runsB), runsB);
  check("zmiana bez zaznaczenia nie robi pustego kroku Cofnij", !(await page.evaluate(() => dwbUndo._debug().undo)).slice(-3).every((l) => l === "undoOpFormat"));

  // ── akapit z samym podziałem strony: podpis i usuwanie jak w Wordzie ──
  page.removeAllListeners("dialog");
  page.on("dialog", (d) => (d.type() === "beforeunload" ? d.accept() : d.dismiss())); // niezapisane zmiany z testów wyżej
  await page.reload();
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector(".docx-preview-host p"), null, { timeout: 30000 });
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForFunction(() => !readOnlyMode && document.querySelector('p[data-lock="lockPageBreakOnly"]'), null, { timeout: 10000 });
  await sleep(400);
  const pbLabel = await page.evaluate(() => getComputedStyle(document.querySelector('p[data-lock="lockPageBreakOnly"]:not([data-dwb-cont])'), "::before").content);
  check("sam podział strony: w Edycji podpis „Podział strony” (zamiast „edytuj w Wordzie”)", /Podział strony/.test(pbLabel), pbLabel);
  const pbState = () => page.evaluate(async () => ({
    secs: document.querySelectorAll("#docCanvas section.docx").length,
    n: collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).length,
    br: (await (await JSZip.loadAsync(await buildDocumentForSave())).file("word/document.xml").async("string")).includes('w:type="page"'),
  }));
  const iZeta2 = await idx("Zeta");
  await page.evaluate((i) => focusParagraphAtOffset(i, 0), iZeta2);
  await page.keyboard.press("Backspace");
  await idle();
  let pb = await pbState();
  check("Backspace na początku akapitu za podziałem usuwa podział (podgląd i plik)", pb.secs === 1 && pb.n === 13 && !pb.br, JSON.stringify(pb));
  await undo();
  const iEps2 = await idx("Epsilon");
  await page.evaluate((i) => focusParagraphAtOffset(i, 99), iEps2);
  await page.keyboard.press("Delete");
  await idle();
  pb = await pbState();
  check("Delete na końcu akapitu przed podziałem usuwa podział", pb.secs === 1 && pb.n === 13 && !pb.br, JSON.stringify(pb));

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  let fail = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : `  (${r.detail})`}`);
    if (!r.ok) fail++;
  }
  console.log(`\n[${ENGINE}] ${fail ? `${fail} z ${results.length} nie przeszło` : `${results.length}/${results.length} OK`}`);
  process.exit(fail ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
