// paste-rich-playwright.js — wklejanie z zachowaniem stylów (app/paste-rich.js, op „pasteBlocks”).
//
// HTML (jak z Notatek Apple / strony / Google Docs): nagłówek, pogrubienie/kursywa/przekreślenie,
// lista punktowana z poziomem, numerowana, lista kontrolna → pola wyboru Worda, link, tabela;
// kroje/rozmiary/kolory ze źródła NIE przechodzą. Markdown z text/plain: #, **, ~~, `kod`, [link](),
// - [ ], tabela |a|b|. Jedno zdanie z pogrubieniem — bez przerysowania, w miejscu kursora. Zwykły
// tekst — jak dawniej. Tekst przed i za kursorem zostaje na swoich miejscach. Jedno Cofnij.
// ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const NOTES_HTML = `<meta charset="utf-8"><h1 style="font-family:'Helvetica Neue';font-size:24px">Zakupy na budowę</h1>
<div style="font-family:'Helvetica Neue';font-size:13px;color:#ff2600"><b>Pilne</b> rzeczy, <i>kursywa</i> i <span style="text-decoration:line-through">skreślone</span>.</div>
<div><br></div>
<ul><li>kabel YDY</li><li>puszki<ul><li>podtynkowe</li></ul></li></ul>
<ol><li>zamówić</li><li>odebrać</li></ol>
<ul class="checklist"><li class="checked">zmierzone</li><li>zamówione</li></ul>
<div>Sklep: <a href="https://example.com/sklep">example.com</a> <a href="javascript:alert(1)">zły</a></div>
<table><tr><th>Pozycja</th><th>Ilość</th></tr><tr><td>Kabel</td><td>120 m</td></tr></table>
<script>window.__xss = 1</script><img src="https://example.com/x.png" onerror="window.__xss=2">`;

const MARKDOWN = `## Notatka z odbioru
Tekst z **pogrubieniem**, *kursywą*, ~~skreśleniem~~, \`kodem\` i [linkiem](https://example.com).

- punkt pierwszy
  - pod-punkt
1. krok jeden
2. krok dwa
- [x] zrobione
- [ ] do zrobienia

| A | B |
|---|---|
| 1 | 2 |`;

async function paste(page, data) {
  await page.evaluate((data) => {
    const p = document.querySelector("p[data-paste]");
    const dt = new DataTransfer();
    Object.entries(data).forEach(([k, v]) => dt.setData(k, v));
    p.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  }, data);
}

async function freshPara(page, mark) {
  // akapit z kursorem w środku: „PRZED|ZA”
  await page.evaluate((mark) => {
    document.querySelectorAll("p[data-paste]").forEach((x) => delete x.dataset.paste);
    const p = [...document.querySelectorAll(".docx-editable-p")].filter((x) => x.textContent.trim().length > 5 && !x.closest("td"))[mark];
    p.dataset.paste = "1";
    p.textContent = "";
    p.append(document.createTextNode("PRZED ZA"));
    p.dispatchEvent(new Event("input", { bubbles: true }));
    placeCaret(p, 6);
  }, mark);
}

const docXml = (page) => page.evaluate(async () => {
  const z = await JSZip.loadAsync(await buildDocumentForSave());
  return { doc: await z.file("word/document.xml").async("string"), num: z.file("word/numbering.xml") ? await z.file("word/numbering.xml").async("string") : "" };
});
const paraTexts = (xml) => (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).map((p) => ({ xml: p, text: (p.match(/<w:t[^>]*>[^<]*/g) || []).map((t) => t.replace(/<w:t[^>]*>/, "")).join("") }));

(async () => {
  const browser = await pw[ENGINE].launch();
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=sample`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.waitForFunction(() => !!originalFileBytes && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForFunction(() => document.querySelector(".docx-editable-p"), null, { timeout: 10000 });
  await sleep(300);

  // ── modele (bez dokumentu) ──
  const md = await page.evaluate((t) => dwbPaste.fromMarkdown(t).map((b) => `${b.type}${b.level ?? ""}${b.ordered ? "#" : ""}${b.checked === true ? "x" : b.checked === false ? "o" : ""}`).join(" "), MARKDOWN);
  check("Markdown → bloki: nagłówek, akapit, listy z poziomem, numerowana, kontrolna, tabela", md === "h2 p li0 li1 li0# li0# li0x li0o table", md);
  const inl = await page.evaluate(() => JSON.stringify(dwbPaste.fromMarkdown("a **b** *c* ~~d~~ `e` [f](https://x.pl) snake_case_name 2*3*4")[0].runs));
  check("Markdown w tekście: pogrubienie, kursywa, przekreślenie, kod, link; snake_case i 2*3*4 bez zmian", inl === JSON.stringify([{ text: "a " }, { text: "b", bold: true }, { text: " " }, { text: "c", italic: true }, { text: " " }, { text: "d", strike: true }, { text: " " }, { text: "e", code: true }, { text: " " }, { text: "f", link: "https://x.pl" }, { text: " snake_case_name 2*3*4" }]), inl);
  const plain = await page.evaluate(() => {
    const dt = new DataTransfer();
    dt.setData("text/plain", "zwykły tekst\ndrugi wiersz");
    return dwbPaste.parse(dt);
  });
  check("zwykły tekst bez formatowania → dawne wklejanie", plain === null, JSON.stringify(plain));

  // ── HTML jak z Notatek ──
  await freshPara(page, 1);
  await paste(page, { "text/html": NOTES_HTML, "text/plain": "Zakupy na budowę" });
  await page.waitForFunction(() => document.body.innerText.includes("Zakupy na budowę") && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 });
  await sleep(600);
  let { doc, num } = await docXml(page);
  let ps = paraTexts(doc);
  const iH = ps.findIndex((x) => x.text === "Zakupy na budowę");
  check("tekst przed kursorem zostaje w swoim akapicie", ps[iH - 1]?.text === "PRZED ", ps[iH - 1]?.text);
  check("nagłówek → styl Nagłówek 1", /<w:pStyle w:val="Heading1"\/>/.test(ps[iH]?.xml || ""), ps[iH]?.xml.slice(0, 200));
  const sent = ps.find((x) => x.text.startsWith("Pilne"));
  check("pogrubienie, kursywa, przekreślenie zachowane; krój/rozmiar/kolor ze źródła NIE", sent && /<w:b w:val="1"\/>[^]*?Pilne/.test(sent.xml) && /<w:i w:val="1"\/>[^]*?kursywa/.test(sent.xml) && /<w:strike w:val="1"\/>[^]*?skreślone/.test(sent.xml) && !/Helvetica|FF2600/i.test(sent.xml), sent?.xml.slice(0, 400));
  const li = (t) => ps.find((x) => x.text === t)?.xml || "";
  check("lista punktowana z poziomem 2", /<w:ilvl w:val="0"\/>/.test(li("kabel YDY")) && /<w:ilvl w:val="1"\/>/.test(li("podtynkowe")) && /<w:numFmt w:val="bullet"\/>/.test(num));
  check("lista numerowana (osobna numeracja od 1)", /<w:numPr>/.test(li("zamówić")) && /<w:startOverride w:val="1"\/>/.test(num) && /<w:numFmt w:val="decimal"\/>/.test(num));
  check("lista kontrolna → pola wyboru Worda (zaznaczone / puste)", /<w14:checked w14:val="1"\/>[^]*?zmierzone/.test(li("☒ zmierzone")) && /<w14:checked w14:val="0"\/>[^]*?zamówione/.test(li("☐ zamówione")), [li("☒ zmierzone").slice(0, 200)]);
  const shop = ps.find((x) => x.text.startsWith("Sklep:"));
  check("bezpieczny link zostaje linkiem, javascript: — sam tekst", shop && /<w:hyperlink /.test(shop.xml) && (shop.xml.match(/<w:hyperlink /g) || []).length === 1 && shop.text.includes("zły"), shop?.xml.slice(0, 300));
  check("tabela z nagłówkiem (pogrubiony)", /<w:tbl>[\s\S]*Pozycja[\s\S]*120 m[\s\S]*<\/w:tbl>/.test(doc) && /<w:b w:val="1"\/>(?:(?!<\/w:p>).)*Pozycja/.test(doc));
  const tail = ps.findIndex((x) => x.text === "ZA");
  check("tekst za kursorem w nowym akapicie po tabeli", tail > iH, ps.slice(-3).map((x) => x.text));
  check("bez skryptów i obrazków ze schowka", await page.evaluate(() => !window.__xss) && !/<w:drawing/.test(doc.slice(doc.indexOf("Zakupy"), doc.indexOf("120 m"))));
  const lists = await page.evaluate(() => document.querySelectorAll("p.dwb-list-hang, p[class*='docx-num-']").length);
  check("podgląd: listy narysowane jako listy", lists >= 5, lists);
  await page.evaluate(() => dwbUndo.undo());
  await page.waitForFunction(() => !document.body.innerText.includes("Zakupy na budowę"), null, { timeout: 15000 }).catch(() => {});
  check("jedno Cofnij zdejmuje całą wklejkę", await page.evaluate(() => !document.body.innerText.includes("Zakupy na budowę") && document.body.innerText.includes("PRZED ZA")));

  // ── Markdown ──
  await freshPara(page, 2);
  await paste(page, { "text/plain": MARKDOWN });
  await page.waitForFunction(() => document.body.innerText.includes("Notatka z odbioru") && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 });
  await sleep(600);
  ({ doc, num } = await docXml(page));
  ps = paraTexts(doc);
  check("Markdown: ## → Nagłówek 2", /<w:pStyle w:val="Heading2"\/>/.test(ps.find((x) => x.text === "Notatka z odbioru")?.xml || ""));
  const mdp = ps.find((x) => x.text.startsWith("Tekst z pogrubieniem"));
  check("Markdown: pogrubienie, kursywa, przekreślenie, kod (Courier New), link", mdp && /<w:b w:val="1"\/>[^]*?pogrubieniem/.test(mdp.xml) && /<w:strike w:val="1"\/>/.test(mdp.xml) && /Courier New[^]*?kodem/.test(mdp.xml) && /<w:hyperlink /.test(mdp.xml), mdp?.xml.slice(0, 500));
  check("Markdown: pod-punkt na poziomie 2, lista kontrolna, tabela", /<w:ilvl w:val="1"\/>/.test(ps.find((x) => x.text === "pod-punkt")?.xml || "") && /<w14:checked w14:val="1"\/>/.test(ps.find((x) => x.text === "☒ zrobione")?.xml || "") && /<w:tbl>[\s\S]*?>A<[\s\S]*?>2</.test(doc));
  const caret = await page.evaluate(() => { const s = getSelection(); const n = s.anchorNode; return (n?.nodeType === 1 ? n : n?.parentElement)?.closest("td, p")?.textContent || ""; });
  check("kursor po wklejeniu na końcu wklejki (przed „ZA”)", /ZA$/.test(caret) || caret === "ZA", caret);

  // ── jedno zdanie ze stylem — w miejscu kursora, bez przerysowania ──
  await freshPara(page, 3);
  const before = await page.evaluate(() => document.querySelector("p[data-paste]"));
  await paste(page, { "text/html": "<span>Bardzo <b>ważne</b> zdanie</span>", "text/plain": "Bardzo ważne zdanie" });
  await sleep(300);
  const inline = await page.evaluate(() => { const p = document.querySelector("p[data-paste]"); return { text: p?.textContent, bold: [...(p?.querySelectorAll("span") || [])].some((s) => s.textContent === "ważne" && /bold/.test(s.getAttribute("style") || "")), same: !!p }; });
  check("jedno zdanie: w miejscu kursora, pogrubienie, bez przerysowania", inline.same && inline.text === "PRZED Bardzo ważne zdanieZA" && inline.bold, JSON.stringify(inline));

  check("brak błędów JS", !errors.length, errors.slice(0, 3).join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`  ${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`);
  console.log(`\n${ENGINE}: ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
