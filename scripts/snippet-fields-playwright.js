// snippet-fields-playwright.js — pola w snippetach z typem (snippets.js, snippet-suggest.js).
//
// {{termin:data}} → kalendarz (input type=date), {{status:lista=…}} → lista, {{uwagi:długi}} →
// pole wielowierszowe (wiersze = łamania), {{zgoda:zaznacz}} → ☒/☐; {{podpis:formularz-lista=…}} i
// {{kiedy:formularz-data}} → PRAWDZIWE pola formularza Worda (w:sdt) — jeden krok Cofnij; wartości
// zapamiętane na następny raz; „＋ Pole” w panelu buduje zapis; „Rozwiń w dokumencie” zostawia
// {{pole}}. ENGINE=webkit (Safari/iPad).

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  const browser = await pw[ENGINE].launch();
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(`${APP_URL}?sample=sample`, { waitUntil: "load" });
  await page.evaluate(() => { localStorage.removeItem("documents-workbench-snippets"); document.getElementById("heroSplash")?.remove(); });
  await page.waitForFunction(() => !!originalFileBytes && document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 30000 });
  await page.evaluate(() => { readModeEl.checked = false; readModeEl.dispatchEvent(new Event("change")); });
  await page.waitForFunction(() => document.querySelector(".docx-editable-p"), null, { timeout: 10000 });
  await sleep(300);

  // ── parsowanie ──
  const parsed = await page.evaluate(() => scanSnippetFields("A {{termin:data=długa}} {{status:lista=Nowy|W toku}} {{x}} {{x}} {{p:formularz-lista=Jan|Anna}} {{p:formularz-lista=Jan|Anna}} {{u:długi}}")
    .map((f) => `${f.name}:${f.kind}${f.form ? "+form" : ""}${f.options.length ? `[${f.options.join("|")}]` : ""}${f.format ? `(${f.format})` : ""}`).join(" "));
  check("składnia pól: rodzaj, opcje, format, formularz; zwykłe pole raz, formularz każde wystąpienie", parsed === "termin:date(długa) status:list[Nowy|W toku] x:text p:list+form[Jan|Anna] p:list+form[Jan|Anna] u:long", parsed);

  // ── snippet z polami do wypełnienia (tekst do dokumentu) ──
  await page.evaluate(() => {
    upsertSnippet("zgl", "Termin: {{termin:data}}, status: {{status:lista=Nowy|W toku|Zamknięty}}{cursor}, zgoda {{zgoda:zaznacz}}.\nUwagi: {{uwagi:długi}}");
    const p = [...document.querySelectorAll(".docx-editable-p")].find((x) => x.textContent.trim().length > 5);
    p.dataset.t = "1";
    placeCaret(p, p.textContent.length);
    window.__ins = expandSnippetAtCaret(p, getSnippetByName("zgl"));
  });
  await page.waitForSelector("dialog.sn-dialog[open]");
  const dlg = await page.evaluate(() => [...document.querySelectorAll("dialog.sn-dialog [data-name]")].map((x) => `${x.dataset.name}:${x.tagName.toLowerCase()}${x.type ? `/${x.type}` : ""}`).join(" "));
  check("okienko: kalendarz (date), lista (select), pole wyboru, pole wielowierszowe", dlg === "termin:input/date status:select/select-one zgoda:input/checkbox uwagi:textarea/textarea", dlg);
  await page.fill('dialog.sn-dialog [data-name="termin"]', "2026-11-03");
  await page.selectOption('dialog.sn-dialog [data-name="status"]', "W toku");
  await page.check('dialog.sn-dialog [data-name="zgoda"]');
  await page.fill('dialog.sn-dialog [data-name="uwagi"]', "pierwszy wiersz\ndrugi wiersz");
  await page.click("dialog.sn-dialog button[type=submit]");
  await page.evaluate(() => window.__ins);
  await sleep(300);
  await page.keyboard.type("!");
  const p1 = await page.evaluate(() => ({ text: document.querySelector("p[data-t]").innerText, remembered: JSON.parse(localStorage.getItem("documents-workbench-snippet-fields") || "{}") }));
  check("wstawiony tekst: data 03.11.2026, wybór z listy, ☒, łamanie wiersza, kursor w miejscu {cursor}", /Termin: 03\.11\.2026, status: W toku!, zgoda ☒\.\nUwagi: pierwszy wiersz\ndrugi wiersz$/.test(p1.text), p1.text.slice(-120));
  check("wartości zapamiętane na następny raz", p1.remembered.status === "W toku" && p1.remembered.termin === "2026-11-03" && p1.remembered.zgoda === "1", JSON.stringify(p1.remembered));

  // ── snippet z polami formularza Worda ──
  await page.evaluate(() => {
    upsertSnippet("podp", "Zatwierdza: {{podpis:formularz-lista=Jan Kowalski|Anna Nowak}} dnia {{kiedy:formularz-data}} {{ok:formularz-zaznacz}}");
    const p = [...document.querySelectorAll(".docx-editable-p")].filter((x) => x.textContent.trim().length > 5)[1];
    p.dataset.f = "1";
    placeCaret(p, 0);
    window.__ins2 = expandSnippetAtCaret(p, getSnippetByName("podp"));
  });
  const noDialog = await page.evaluate(() => !document.querySelector("dialog.sn-dialog[open]"));
  await page.evaluate(() => window.__ins2);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden"), null, { timeout: 20000 });
  await sleep(600);
  check("same pola formularza: bez okienka", noDialog);
  const xml = await page.evaluate(async () => (await JSZip.loadAsync(await buildDocumentForSave())).file("word/document.xml").async("string"));
  const seg = xml.slice(xml.indexOf("Zatwierdza:"), xml.indexOf("Zatwierdza:") + 4000);
  check("pole-lista Worda z pozycjami i nazwą", /<w:alias w:val="podpis"\/>[\s\S]*?<w:dropDownList>[\s\S]*?w:displayText="Jan Kowalski"[\s\S]*?w:displayText="Anna Nowak"/.test(seg), seg.slice(0, 300));
  check("pole-data Worda (kalendarz w Wordzie i w aplikacji)", /<w:alias w:val="kiedy"\/>[\s\S]*?<w:date>/.test(seg));
  check("pole wyboru Worda", /<w:alias w:val="ok"\/>[\s\S]*?<w14:checkbox>/.test(seg));
  check("tekst snippetu wokół pól", /Zatwierdza: <\/w:t>[\s\S]*?dnia <\/w:t>/.test(seg));
  const ff = await page.evaluate(() => document.querySelectorAll("p .ff-field").length);
  check("pola w podglądzie działają (formularz)", ff >= 3, ff);
  await page.evaluate(() => dwbUndo.undo());
  await page.waitForFunction(() => !document.body.innerText.includes("Zatwierdza:"), null, { timeout: 15000 }).catch(() => {});
  const undone = await page.evaluate(() => !document.body.innerText.includes("Zatwierdza:") && document.body.innerText.includes("W toku"));
  check("jedno Cofnij zdejmuje cały snippet z polami formularza (wcześniejszy zostaje)", undone);

  // ── „＋ Pole” w panelu ──
  await page.evaluate(() => ensureLazyFeature("snippets-panel"));
  await page.evaluate(() => { const b = document.getElementById("snBody"); b.value = "Data: "; b.focus(); b.setSelectionRange(6, 6); });
  await page.evaluate(() => document.getElementById("snAddFieldBtn").click());
  await page.waitForSelector("dialog.sn-dialog[open] .sf-name");
  await page.fill("dialog.sn-dialog .sf-name", "termin odbioru");
  await page.selectOption("dialog.sn-dialog .sf-type", "lista");
  await page.fill("dialog.sn-dialog .sf-opts", "Pilne\nZwykłe");
  await page.check("dialog.sn-dialog .sf-form");
  await page.click("dialog.sn-dialog .sf-ok");
  const body = await page.evaluate(() => document.getElementById("snBody").value);
  check("„＋ Pole” wstawia zapis w miejscu kursora", body === "Data: {{termin_odbioru:formularz-lista=Pilne|Zwykłe}}", body);

  // ── „Rozwiń w dokumencie” bez pytania: pole z typem → {{pole}} ──
  const plain = await page.evaluate(() => resolveSnippetMapBodies({ a: "X {{t:data}} {{s:lista=A|B}} {{n}}" }).a);
  check("rozwinięcie w całym pliku zostawia {{pole}} (do panelu Placeholdery)", plain === "X {{t}} {{s}} {{n}}", plain);

  check("brak błędów JS", !errors.length, errors.slice(0, 3).join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  for (const r of results) console.log(`  ${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : ` — ${r.detail}`}`);
  console.log(`\n${ENGINE}: ${results.length - failed.length}/${results.length} OK`);
  process.exit(failed.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
