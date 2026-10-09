// links-word-playwright.js — linki i zakładki jak w Wordzie (2026-10-06).
//
// Okienko „Link” (Ctrl/⌘+K) jak „Wstaw hiperłącze”: Adres WWW / W dokumencie / E-mail (adres +
// temat → mailto:?subject=), Początek dokumentu (w:anchor="_top"), Nagłówki, Zakładki, Obrazy,
// szukanie miejsca, etykietka ekranowa (w:tooltip — zostaje po pisaniu w akapicie), ten sam adres
// = to samo powiązanie w pliku. Okienko „Zakładka” (Ctrl/⌘+Shift+F5) jak w Wordzie: nazwa
// (walidacja), Dodaj / Przenieś / Usuń (ostrzeżenie, gdy prowadzą do niej linki) / Przejdź do.
// Pisanie: adres/e-mail + spacja/Enter = link (Ctrl/⌘+Z zdejmuje sam link), kropka na końcu
// zdania poza adresem, pisanie tuż za linkiem nie przedłuża linku. Wklejanie: sam adres = link,
// adres na zaznaczony tekst = ten tekst staje się linkiem. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const MOD = process.platform === "darwin" ? "Meta" : "Control";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const idle = (page) => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 20000 }).then(() => page.waitForTimeout(400));

const saved = async (page) => {
  const b64 = await page.evaluate(async () => { const bytes = await buildDocumentForSave(); let s = ""; bytes.forEach((x) => { s += String.fromCharCode(x); }); return btoa(s); });
  const zip = await JSZip.loadAsync(Buffer.from(b64, "base64"));
  return { xml: await zip.file("word/document.xml").async("string"), rels: await zip.file("word/_rels/document.xml.rels").async("string") };
};
// kursor / zaznaczenie w akapicie i (przesunięcia w tekście akapitu)
const select = (page, i, a, b = a) => page.evaluate(([i, a, b]) => {
  const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i];
  const r = formDomRange(p, a, b);
  p.closest(".docx-edit-root").focus({ preventScroll: true });
  getSelection().removeAllRanges();
  getSelection().addRange(r);
}, [i, a, b]);
const paraText = (page, i) => page.evaluate((i) => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[i]?.textContent || "", i);
const linkForm = async (page) => {
  await page.evaluate(() => composeUi.openLinkForm(document.getElementById("insertMenuBtn")));
  await page.waitForSelector(".compose-pop-link .lf-ok");
};
const placeLabels = (page) => page.evaluate(() => [...document.querySelectorAll(".lf-places > *")].map((x) => (x.classList.contains("lf-group") ? `#${x.textContent}` : x.querySelector(".lf-place-label")?.textContent || x.textContent)));
// zakładki dociągają się po zmianie pliku — czekamy, aż miejsce będzie na liście
const pickPlace = async (page, label) => {
  await page.waitForFunction((label) => [...document.querySelectorAll(".lf-place-label")].some((x) => x.textContent === label), label, { timeout: 5000 }).catch(() => {});
  return page.evaluate((label) => [...document.querySelectorAll(".lf-place")].find((b) => b.querySelector(".lf-place-label").textContent === label)?.click(), label);
};

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  if (ENGINE === "chromium") await context.grantPermissions(["clipboard-read", "clipboard-write"]);
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
  // 0: zdanie z linkami, 1: nagłówek, 2..: tekst
  await page.keyboard.type("Zobacz rozdział i napisz do nas.");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Rozdział drugi");
  await page.keyboard.press("Enter");
  for (let i = 0; i < 4; i++) { await page.keyboard.type(`Treść rozdziału ${i + 1}.`); await page.keyboard.press("Enter"); }
  await idle(page);
  await select(page, 1, 0);
  await page.evaluate(() => composeUi.applyStyle("h1"));
  await idle(page);

  // ── zakładka na zaznaczonym tekście (Wstaw → Zakładka) ──
  await select(page, 2, 0, 5); // „Treść”
  await page.keyboard.press(`${MOD}+Shift+F5`);
  await page.waitForSelector(".compose-pop .bm-name");
  const sug = await page.inputValue(".bm-name");
  check("zakładka: nazwa podpowiedziana z zaznaczenia", sug === "Treść", sug);
  await page.fill(".bm-name", "zła nazwa");
  const bad = await page.evaluate(() => ({ msg: document.querySelector(".bm-name-msg").textContent, dis: document.querySelector(".bm-add").disabled }));
  check("zakładka: nazwa ze spacją odrzucona z wyjaśnieniem (jak Word)", bad.dis && /spacji/.test(bad.msg), JSON.stringify(bad));
  await page.fill(".bm-name", "1abc");
  check("zakładka: nazwa od cyfry odrzucona", await page.evaluate(() => document.querySelector(".bm-add").disabled));
  await page.fill(".bm-name", "Cel_1");
  await page.click(".bm-add");
  await idle(page);
  let f = await saved(page);
  check("plik: zakładka obejmuje zaznaczony tekst", /<w:bookmarkStart w:id="\d+" w:name="Cel_1"\/><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>Treść<\/w:t><\/w:r><w:bookmarkEnd w:id="\d+"\/>/.test(f.xml), (f.xml.match(/.{0,80}Cel_1.{0,200}/) || [""])[0]);

  // ── okienko linku: miejsca w dokumencie ──
  await select(page, 0, 7, 15); // „rozdział”
  await page.keyboard.press(`${MOD}+k`);
  await page.waitForSelector(".compose-pop-link .lf-ok");
  const modes = await page.evaluate(() => [...document.querySelectorAll(".lf-mode button")].map((b) => b.textContent));
  check("okienko linku: trzy rodzaje (WWW / w dokumencie / e-mail)", modes.length === 3, modes.join(" | "));
  await page.click(".lf-mode button[data-mode=doc]");
  await page.waitForFunction(() => [...document.querySelectorAll(".lf-place-label")].some((x) => x.textContent === "Cel_1"), null, { timeout: 5000 }).catch(() => {});
  const labels = await placeLabels(page);
  check("miejsca: Początek dokumentu, Nagłówki, Zakładki (jak Word)", labels.includes("Początek dokumentu") && labels.includes("#Nagłówki") && labels.includes("Rozdział drugi") && labels.includes("#Zakładki") && labels.includes("Cel_1"), labels.join(", "));
  await page.fill(".lf-find", "cel");
  const found = await placeLabels(page);
  check("szukanie miejsca zawęża listę", found.filter((x) => !x.startsWith("#")).join() === "Cel_1", found.join(", "));
  await page.press(".lf-find", "Enter");
  await idle(page);
  f = await saved(page);
  check("plik: link do zakładki (w:anchor=nazwa)", /<w:hyperlink w:anchor="Cel_1"[^>]*>.*?rozdział/.test(f.xml));

  // zmiana tego linku: okienko pokazuje zakładkę, przestawiamy na Początek dokumentu + etykietka
  await select(page, 0, 9);
  await linkForm(page);
  await page.waitForFunction(() => !!document.querySelector(".lf-place.is-on"), null, { timeout: 5000 }).catch(() => {});
  const st = await page.evaluate(() => ({ mode: document.querySelector(".lf-mode .is-on")?.dataset.mode, sel: document.querySelector(".lf-place.is-on .lf-place-label")?.textContent, title: document.querySelector(".compose-pop-link .compose-cap").textContent }));
  check("„Zmień link”: tryb w dokumencie, wybrana zakładka", st.mode === "doc" && st.sel === "Cel_1" && st.title === "Zmień link", JSON.stringify(st));
  await pickPlace(page, "Początek dokumentu");
  await page.click(".lf-tip-toggle");
  await page.fill(".lf-tip-in", "Wróć na górę");
  await page.click(".lf-ok");
  await idle(page);
  f = await saved(page);
  check("plik: Początek dokumentu = w:anchor=\"_top\" + w:tooltip", /<w:hyperlink [^>]*w:anchor="_top"[^>]*>/.test(f.xml) && /w:tooltip="Wróć na górę"/.test(f.xml), (f.xml.match(/<w:hyperlink [^>]*>/) || [""])[0]);
  // pisanie w tym akapicie nie gubi etykietki (dawniej przepisanie akapitu ją zrzucało)
  await select(page, 0, (await paraText(page, 0)).length);
  await page.keyboard.type(" Dalej.");
  await idle(page);
  f = await saved(page);
  check("etykietka zostaje po pisaniu w akapicie z linkiem", /w:tooltip="Wróć na górę"/.test(f.xml));
  const hint = await page.evaluate(() => collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[0].querySelector("a")?.dataset.hintPl);
  check("podpowiedź linku = etykietka ekranowa (jak Word)", hint === "Wróć na górę", hint);

  // ── e-mail z tematem ──
  const t0 = await paraText(page, 0);
  const at = t0.indexOf("do nas");
  await select(page, 0, at, at + 6);
  await linkForm(page);
  await page.click(".lf-mode button[data-mode=mail]");
  await page.fill(".lf-addr", "biuro@example.com");
  await page.fill(".lf-subject", "Pytanie o ofertę");
  await page.click(".lf-ok");
  await idle(page);
  f = await saved(page);
  check("plik: e-mail z tematem = mailto:?subject=", f.rels.includes("mailto:biuro@example.com?subject=Pytanie%20o%20ofert%C4%99"), (f.rels.match(/Target="mailto[^"]*"/) || [""])[0]);
  await select(page, 0, at + 2);
  await linkForm(page);
  const mail = await page.evaluate(() => ({ mode: document.querySelector(".lf-mode .is-on")?.dataset.mode, a: document.querySelector(".lf-addr").value, s: document.querySelector(".lf-subject").value }));
  check("„Zmień link” e-mail: adres i temat wracają do pól", mail.mode === "mail" && mail.a === "biuro@example.com" && mail.s === "Pytanie o ofertę", JSON.stringify(mail));
  await page.click(".lf-cancel");

  // ── adres WWW: ten sam adres = to samo powiązanie ──
  await select(page, 2, 6, 15); // „rozdziału”
  await linkForm(page);
  await page.fill(".lf-url", "example.org/strona");
  await page.click(".lf-ok");
  await idle(page);
  f = await saved(page);
  const relCount = (f.rels.match(/TargetMode="External"/g) || []).length;
  check("plik: adres bez https:// dostaje https://", f.rels.includes('Target="https://example.org/strona"'));
  await select(page, 2, 8);
  await linkForm(page);
  await page.fill(".lf-text", "rozdziału (WWW)");
  await page.click(".lf-ok");
  await idle(page);
  f = await saved(page);
  check("zmiana samego tekstu linku: bez nowego powiązania w pliku", (f.rels.match(/TargetMode="External"/g) || []).length === relCount && (await paraText(page, 2)).includes("rozdziału (WWW)"), String((f.rels.match(/TargetMode="External"/g) || []).length));

  // ── usuwanie zakładki z linkami: ostrzeżenie, potem usunięcie ──
  await select(page, 3, 0);
  await linkForm(page);
  await page.click(".lf-mode button[data-mode=doc]");
  await pickPlace(page, "Cel_1");
  await page.click(".lf-ok");
  await idle(page);
  await select(page, 4, 0);
  await page.keyboard.press(`${MOD}+Shift+F5`);
  await page.waitForSelector(".compose-pop .bm-name");
  await page.waitForFunction(() => !!document.querySelector('.bm-list .lf-place[data-name="Cel_1"]'), null, { timeout: 5000 }).catch(() => {});
  await page.evaluate(() => [...document.querySelectorAll(".bm-list .lf-place")].find((b) => b.dataset.name === "Cel_1")?.click());
  const listInfo = await page.evaluate(() => document.querySelector(".bm-info").textContent);
  await page.click(".bm-delete");
  const warn = await page.evaluate(() => ({ info: document.querySelector(".bm-info").textContent, btn: document.querySelector(".bm-delete").textContent, open: !!document.querySelector(".compose-pop .bm-name") }));
  check("usuwanie zakładki z linkiem: najpierw ostrzeżenie z liczbą linków", warn.open && /linków: 1/.test(warn.info) && warn.btn === "Usuń mimo to", `${listInfo} | ${JSON.stringify(warn)}`);
  await page.click(".bm-delete");
  await idle(page);
  await page.waitForSelector(".compose-pop .bm-name", { timeout: 5000 }).catch(() => {});
  f = await saved(page);
  check("po „Usuń mimo to”: zakładki nie ma, okienko otwarte dalej (jak Word)", !/w:name="Cel_1"/.test(f.xml) && !!(await page.$(".compose-pop .bm-name")), `${/w:name="Cel_1"/.test(f.xml)} ${!!(await page.$(".compose-pop .bm-name"))}`);
  await page.keyboard.press("Escape");

  // ── przenoszenie zakładki (ta sama nazwa) ──
  await select(page, 5, 0, 5);
  await page.keyboard.press(`${MOD}+Shift+F5`);
  await page.waitForSelector(".compose-pop .bm-name");
  await page.fill(".bm-name", "Koniec");
  await page.click(".bm-add");
  await idle(page);
  await select(page, 4, 0);
  await page.keyboard.press(`${MOD}+Shift+F5`);
  await page.waitForSelector(".compose-pop .bm-name");
  await page.fill(".bm-name", "Koniec");
  const mv = await page.evaluate(() => document.querySelector(".bm-add").textContent);
  await page.click(".bm-add");
  await idle(page);
  f = await saved(page);
  check("ta sama nazwa = „Przenieś tutaj”: jedna zakładka, w nowym miejscu", mv === "Przenieś tutaj" && (f.xml.match(/w:name="Koniec"/g) || []).length === 1 && /w:name="Koniec"\/>(?:<w:bookmarkEnd[^>]*\/>)?<w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>Treść rozdziału 3/.test(f.xml), mv);

  // ── autoformatowanie: adres + spacja ──
  await select(page, 5, (await paraText(page, 5)).length);
  await page.keyboard.type(" Strona www.example.com. Pisz na jan@example.com teraz");
  await page.keyboard.press("Enter");
  await idle(page);
  f = await saved(page);
  check("wpisany www.… + kropka + spacja = link bez kropki", f.rels.includes('Target="https://www.example.com"') && /<w:hyperlink [^>]*>(?:(?!<\/w:hyperlink>).)*>www\.example\.com<\/w:t>/.test(f.xml), (f.xml.match(/.{0,40}www\.example.{0,40}/) || [""])[0]);
  check("wpisany e-mail + spacja = link mailto:", f.rels.includes('Target="mailto:jan@example.com"'));
  const p5 = await paraText(page, 5);
  check("tekst akapitu bez zmian (kropka, spacje)", p5.endsWith("Strona www.example.com. Pisz na jan@example.com teraz"), p5);

  // Ctrl/⌘+Z zaraz po zamianie zdejmuje sam link
  await page.keyboard.type("Adres https://example.net/a");
  await page.keyboard.press("Space");
  await page.waitForTimeout(200);
  const linked = await page.evaluate(() => !!document.querySelector('.docx-preview-host a[data-dwb-link="https://example.net/a"]'));
  await page.keyboard.press(`${MOD}+z`);
  await idle(page);
  const after = await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[6];
    return { a: !!p.querySelector("a"), text: p.textContent };
  });
  check("Ctrl/⌘+Z po automatycznym linku: link znika, tekst zostaje", linked && !after.a && /Adres https:\/\/example\.net\/a\s?$/.test(after.text), JSON.stringify({ linked, ...after }));

  // pisanie tuż za linkiem nie przedłuża linku
  await select(page, 2, 0);
  const t2 = await paraText(page, 2);
  const endOfLink = await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[2]; const a = p.querySelector("a"); const r = document.createRange(); r.selectNodeContents(a); const last = r.endContainer; const sel = getSelection(); const rr = document.createRange(); let n = a; while (n.lastChild) n = n.lastChild; rr.setStart(n, n.textContent.length); rr.collapse(true); p.closest(".docx-edit-root").focus({ preventScroll: true }); sel.removeAllRanges(); sel.addRange(rr); return a.textContent; });
  await page.keyboard.type("XYZ");
  const edge = await page.evaluate(() => { const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[2]; return { a: p.querySelector("a").textContent, p: p.textContent }; });
  check("pisanie tuż za linkiem: litery poza linkiem", edge.a === endOfLink && edge.p.includes(`${endOfLink}XYZ`), JSON.stringify({ endOfLink, t2, ...edge }));
  await idle(page);

  // ── wklejanie adresu ──
  const paste = async (text) => page.evaluate((text) => {
    const dt = new DataTransfer();
    dt.setData("text/plain", text);
    const target = getSelection().anchorNode?.parentElement || document.querySelector(".docx-edit-root");
    target.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  }, text);
  await select(page, 4, 0, 5); // „Treść”
  await paste("https://example.com/wklejony");
  await idle(page);
  f = await saved(page);
  check("adres wklejony na zaznaczony tekst = ten tekst jest linkiem", f.rels.includes('Target="https://example.com/wklejony"') && (await paraText(page, 4)).startsWith("Treść rozdziału 3"));
  await select(page, 3, (await paraText(page, 3)).length);
  await paste("www.wklejka.pl");
  await idle(page);
  f = await saved(page);
  check("sam wklejony adres = link z tym adresem jako tekstem", f.rels.includes('Target="https://www.wklejka.pl"') && (await paraText(page, 3)).endsWith("www.wklejka.pl"), await paraText(page, 3));
  // wyłącznik w „Narzędziach edycji”: adres wkleja się jako zwykły tekst (zaznaczenie zastąpione)
  await page.evaluate(() => { const box = document.getElementById("pasteAutoLink"); box.checked = false; box.dispatchEvent(new Event("change", { bubbles: true })); });
  await select(page, 5, 0, 5);
  await paste("https://example.com/bez-linku");
  await idle(page);
  f = await saved(page);
  const plain = await page.evaluate(() => {
    const p = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"))[5];
    return { a: !!p.querySelector('a[data-dwb-link*="bez-linku"], a[href*="bez-linku"]'), text: p.textContent, stored: localStorage.getItem("dwb-paste-autolink-v1") };
  });
  check("opcja wyłączona: wklejony adres = zwykły tekst, bez linku", !f.rels.includes("bez-linku") && !plain.a && plain.text.startsWith("https://example.com/bez-linku") && plain.stored === "0", JSON.stringify(plain));
  await page.evaluate(() => { const box = document.getElementById("pasteAutoLink"); box.checked = true; box.dispatchEvent(new Event("change", { bubbles: true })); });

  // ── czytanie: Początek dokumentu przewija na górę ──
  await page.evaluate(() => appFrame.setReadOnly(true));
  await idle(page);
  await page.evaluate(() => { docViewportEl.scrollTop = 400; });
  await page.evaluate(() => document.querySelector('.docx-preview-host a[href="#_top"]').click());
  await page.waitForTimeout(900);
  check("klik „Początek dokumentu”: dokument na górze, bez komunikatu o braku miejsca", (await page.evaluate(() => docViewportEl.scrollTop)) < 5 && !(await page.evaluate(() => [...document.querySelectorAll(".toast")].some((x) => /nie ma w dokumencie/.test(x.textContent)))));

  check("brak błędów strony", errors.length === 0, errors.join(" | "));
  await browser.close();
}

run().then(() => {
  let failed = 0;
  results.forEach((r) => { if (!r.ok) failed++; console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` — ${r.detail}`}`); });
  console.log(`\n${ENGINE}: ${results.length - failed}/${results.length} OK`);
  process.exit(failed ? 1 : 0);
}).catch((e) => { console.error(e); process.exit(1); });
