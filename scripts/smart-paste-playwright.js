// smart-paste-playwright.js — inteligentne wycinanie i wklejanie jak w Wordzie (2026-10-17).
//
// Word: Preferencje → Edycja → „Use smart cut and paste” + „When selecting, automatically select
// entire word”. Oczekiwane wyniki to wyniki Worda 16 (AppleScript cut/copy/paste object, Backspace
// na zaznaczeniu, Cmd+X/Cmd+V i dwuklik prawdziwą myszą): wycięte/usunięte całe słowo nie zostawia
// podwójnej spacji ani spacji przed kropką, wklejone słowo dostaje brakującą spację po właściwej
// stronie, dwuklik zaznacza słowo ze spacją za nim. Wytnij i Wklej = jeden krok Cofnij.
// Wyłącznik w Narzędziach edycji. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 })
  .then(() => page.evaluate(() => waitInlineStructuralIdle())).then(() => page.waitForTimeout(150));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => {
    sessionStorage.setItem("introPlayed", "true");
    localStorage.setItem("dwb-smart-cut-v1", "1"); // pod automatem domyślnie wyłączone
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => composeUi.createNew("blank"));
  await page.waitForSelector(".docx-preview-host p.docx-editable-p", { timeout: 20000 });
  await idle(page);

  // nowy akapit na końcu z podanym tekstem; zwraca jego numer
  let first = true;
  const para = async (text) => {
    await page.evaluate(() => {
      const ps = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
      placeCaret(ps[ps.length - 1], Number.MAX_SAFE_INTEGER);
    });
    if (!first) { await page.keyboard.press("Enter"); await idle(page); }
    first = false;
    await page.keyboard.insertText(text);
    await idle(page);
    return page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host")).length - 1);
  };
  const textOf = (i) => page.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].textContent.replace(/﻿/g, ""), i);
  const select = (i, s, e) => page.evaluate(({ i, s, e }) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    const r = formDomRange(p, s, e);
    docEditRoot().focus();
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  }, { i, s, e });
  // zdarzenia schowka jak w paste-format-playwright.js (headless nie ma systemowego schowka)
  const cut = () => page.evaluate(() => {
    window.__clip = new DataTransfer();
    docEditRoot().dispatchEvent(new ClipboardEvent("cut", { clipboardData: window.__clip, bubbles: true, cancelable: true }));
  }).then(() => idle(page));
  const copy = () => page.evaluate(() => {
    window.__clip = new DataTransfer();
    docEditRoot().dispatchEvent(new ClipboardEvent("copy", { clipboardData: window.__clip, bubbles: true, cancelable: true }));
  });
  const pasteClip = () => page.evaluate(() => {
    docEditRoot().dispatchEvent(new ClipboardEvent("paste", { clipboardData: window.__clip, bubbles: true, cancelable: true }));
  }).then(() => idle(page));
  const pasteText = (text, html) => page.evaluate(({ text, html }) => {
    const dt = new DataTransfer();
    dt.setData("text/plain", text);
    if (html) dt.setData("text/html", html);
    docEditRoot().dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  }, { text, html }).then(() => idle(page));

  // ── 1) Wytnij całe słowo (wyniki Worda) ────────────────────────────────────
  const CUTS = [
    ["Ala ma kota.", 4, 6, "Ala kota.", "słowo w środku → jedna spacja"],
    ["Ala ma kota.", 7, 11, "Ala ma.", "słowo przed kropką → bez spacji przed kropką"],
    ["Ala ma kota, psa.", 7, 11, "Ala ma, psa.", "słowo przed przecinkiem"],
    ["Ala ma kota!", 7, 11, "Ala ma!", "słowo przed wykrzyknikiem"],
    ["Ala ma kota.", 0, 3, "ma kota.", "pierwsze słowo akapitu → bez spacji na początku"],
    ["Ala ma kota", 7, 11, "Ala ma ", "ostatnie słowo akapitu → spacja zostaje (jak w Wordzie)"],
    ["Ala ma  kota.", 4, 6, "Ala  kota.", "przy dwóch spacjach znika jedna"],
    ["Ala ma kota psa.", 4, 11, "Ala psa.", "dwa słowa"],
    ["Ala ma kota.", 4, 7, "Ala kota.", "zaznaczenie ze spacją — zwykłe wycięcie"],
    ["Ala ma kota.", 8, 10, "Ala ma ka.", "kawałek słowa — bez poprawki"],
    ["Mam 5 kotów.", 4, 5, "Mam  kotów.", "sama cyfra — Word nie poprawia"],
    ["Ala (ma) kota", 4, 8, "Ala  kota", "nawias na brzegu — Word nie poprawia"],
    ["Ala ma kota-psa.", 7, 11, "Ala ma -psa.", "łącznik za słowem — bez poprawki"],
  ];
  for (const [text, s, e, want, name] of CUTS) {
    const i = await para(text);
    await select(i, s, e);
    await cut();
    const got = await textOf(i);
    check(`Wytnij: ${name}`, got === want, `„${got}” ≠ „${want}”`);
  }
  // schowek dostaje samo zaznaczenie (bez spacji)
  {
    const i = await para("Ola ma psa.");
    await select(i, 4, 6);
    await cut();
    const plain = await page.evaluate(() => window.__clip.getData("text/plain"));
    check("Wytnij: w schowku samo słowo „ma”", plain === "ma", JSON.stringify(plain));
    // Cofnij: jeden krok przywraca słowo i spację
    await page.keyboard.press(`${MOD}+z`);
    await idle(page);
    const back = await textOf(i);
    check("Cofnij po Wytnij: jeden krok przywraca „Ola ma psa.”", back === "Ola ma psa.", back);
  }

  // ── 2) Backspace / Delete na zaznaczonym słowie ────────────────────────────
  {
    const i = await para("Ala ma kota.");
    await select(i, 4, 6);
    await page.keyboard.press("Backspace");
    await idle(page);
    let got = await textOf(i);
    check("Backspace na „ma” → „Ala kota.”", got === "Ala kota.", got);
    await select(i, 4, 8);
    await page.keyboard.press("Delete");
    await idle(page);
    got = await textOf(i);
    check("Delete na „kota” → „Ala.”", got === "Ala.", got);
    await page.keyboard.insertText("x");
    await idle(page);
    got = await textOf(i);
    check("kursor po Delete w miejscu słowa", got === "Alax.", got);
  }

  // ── 3) Wklej zwykły tekst z innego programu ────────────────────────────────
  const PASTES = [
    ["Ala kota.", 3, "ma", "Ala ma kota.", "za słowem, przed spacją → spacja przed"],
    ["Ala kota.", 4, "ma", "Ala ma kota.", "po spacji, przed słowem → spacja za"],
    ["Ala ma.", 6, "kota", "Ala ma kota.", "przed kropką → spacja przed"],
    ["Ala ma.", 7, "kota", "Ala ma. kota", "za kropką na końcu akapitu → spacja przed"],
    ["Ola psa, a Ewa.", 8, "pies", "Ola psa, pies a Ewa.", "za przecinkiem → spacja przed"],
    ["Ala kota.", 0, "pies", "pies Ala kota.", "początek akapitu → spacja za"],
    ["Ala kota.", 5, "pies", "Ala k pies ota.", "w środku słowa → spacje z obu stron (Word)"],
    ["Ala kota.", 4, " ma", "Ala ma kota.", "wklejka ze spacją przy spacji — bez podwójnej"],
    ["Ala kota.", 3, "ma ", "Ala ma kota.", "wklejka ze spacją na końcu przy spacji — bez podwójnej"],
    ["Ala kota.", 3, ",", "Ala, kota.", "sam przecinek — bez spacji"],
    ["Ala ().", 5, "ma", "Ala (ma).", "w nawiasie — bez spacji"],
  ];
  for (const [text, at, clip, want, name] of PASTES) {
    const i = await para(text);
    await select(i, at, at);
    await pasteText(clip);
    const got = await textOf(i);
    check(`Wklej: ${name}`, got === want, `„${got}” ≠ „${want}”`);
  }
  // wklejenie w miejsce zaznaczenia: liczą się znaki obok zaznaczenia
  {
    const i = await para("Ala kot psa.");
    await select(i, 4, 7);
    await pasteText("ma");
    const got = await textOf(i);
    check("Wklej w miejsce zaznaczonego słowa", got === "Ala ma psa.", got);
  }
  // HTML z innego programu (paste-rich.js) — spacja w formacie miejsca, pogrubione tylko słowo
  {
    const i = await para("Ala kota.");
    await select(i, 3, 3);
    await pasteText("ma", "<b>ma</b>");
    const r = await page.evaluate((i) => {
      const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
      const bold = [...p.querySelectorAll("b, strong, [style*='bold'], [style*='700']")].map((x) => x.textContent).join("|");
      return { text: p.textContent.replace(/﻿/g, ""), bold };
    }, i);
    check("Wklej HTML: „Ala ma kota.”, pogrubione samo „ma”", r.text === "Ala ma kota." && r.bold === "ma", JSON.stringify(r));
  }
  // jeden krok Cofnij po wklejeniu
  {
    const i = await para("Ewa nic.");
    await select(i, 3, 3);
    await pasteText("ma");
    let got = await textOf(i);
    check("Wklej: „Ewa ma nic.”", got === "Ewa ma nic.", got);
    await page.keyboard.press(`${MOD}+z`);
    await idle(page);
    got = await textOf(i);
    check("Cofnij po Wklej: jeden krok, bez dołożonej spacji", got === "Ewa nic.", got);
  }

  // ── 4) Wytnij → Wklej w tym dokumencie (schowek z formatowaniem źródła) ─────
  {
    const i = await para("Ala ma kota.");
    await select(i, 4, 6);
    await cut();
    await select(i, 8, 8); // „Ala kota|.”
    await pasteClip();
    const got = await textOf(i);
    check("Wytnij „ma” + Wklej przed kropką → „Ala kota ma.”", got === "Ala kota ma.", got);
  }
  {
    const i = await para("Alakota.");
    await select(i, 0, 3);
    await copy(); // kawałek słowa — bez spacji (Word)
    await select(i, 3, 3);
    await pasteClip();
    const got = await textOf(i);
    check("Kopiuj kawałek słowa + Wklej — bez spacji (Word)", got === "AlaAlakota.", got);
  }

  // ── 5) Dwuklik zaznacza słowo ze spacją za nim ─────────────────────────────
  const wordBox = (i, s, e) => page.evaluate(({ i, s, e }) => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
    const b = formDomRange(p, s, e).getBoundingClientRect();
    return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  }, { i, s, e });
  {
    const i = await para("Ola ma psa i kota.");
    await page.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i].scrollIntoView({ block: "center" }), i);
    let b = await wordBox(i, 4, 6);
    await page.mouse.dblclick(b.x, b.y);
    await page.waitForTimeout(100);
    let sel = await page.evaluate(() => getSelection().toString());
    check("dwuklik „ma” zaznacza „ma ”", sel === "ma ", JSON.stringify(sel));
    await page.keyboard.press("Backspace");
    await idle(page);
    let got = await textOf(i);
    check("dwuklik + Backspace → „Ola psa i kota.”", got === "Ola psa i kota.", got);
    b = await wordBox(i, 10, 14);
    await page.mouse.dblclick(b.x, b.y);
    await page.waitForTimeout(100);
    sel = await page.evaluate(() => getSelection().toString());
    check("dwuklik „kota” przed kropką zaznacza samo „kota”", sel === "kota", JSON.stringify(sel));
    await cut();
    got = await textOf(i);
    check("dwuklik + Wytnij „kota” → „Ola psa i.”", got === "Ola psa i.", got);
  }

  // ── 6) wyłącznik ───────────────────────────────────────────────────────────
  {
    const on = await page.evaluate(() => document.getElementById("smartCutPaste")?.checked);
    check("wyłącznik w Narzędziach edycji jest i jest włączony", on === true, String(on));
    await page.evaluate(() => { const b = document.getElementById("smartCutPaste"); b.checked = false; b.dispatchEvent(new Event("change", { bubbles: true })); });
    const stored = await page.evaluate(() => localStorage.getItem("dwb-smart-cut-v1"));
    check("wyłącznik zapamiętany", stored === "0", stored);
    let i = await para("Ala ma kota.");
    await select(i, 4, 6);
    await cut();
    let got = await textOf(i);
    check("wyłączone: Wytnij zostawia podwójną spację", got === "Ala  kota.", got);
    i = await para("Ala kota.");
    await select(i, 3, 3);
    await pasteText("ma");
    got = await textOf(i);
    check("wyłączone: Wklej bez dokładania spacji", got === "Alama kota.", got);
    await page.evaluate(() => { const b = document.getElementById("smartCutPaste"); b.checked = true; b.dispatchEvent(new Event("change", { bubbles: true })); });
  }

  // ── 7) zapis ───────────────────────────────────────────────────────────────
  await idle(page);
  const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
  const zip = await JSZip.loadAsync(Buffer.from(b64, "base64"));
  const xml = await zip.file("word/document.xml").async("string");
  const texts = [...xml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((m) => [...m[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((x) => x[1]).join(""));
  check("zapis: wyniki w pliku", texts.includes("Ala ma, psa.") && texts.includes("Ola psa, pies a Ewa.") && texts.includes("Ala kota ma.") && texts.includes("Ola psa i."), JSON.stringify(texts.slice(0, 6)));

  check("brak błędów w konsoli", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();

  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`));
  console.log(`\n[${ENGINE}] ${results.length - failed.length}/${results.length} OK`);
  if (failed.length) process.exit(1);
}

run().catch((err) => { console.error(err); process.exit(1); });
