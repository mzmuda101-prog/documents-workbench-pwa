// table-keys-playwright.js — klawisze w tabeli jak w Wordzie (2026-11-13).
//
// Tab / Shift+Tab: następna / poprzednia komórka z ZAZNACZONĄ treścią (pisanie ją zastępuje — skróty
// Worda: „Tab — przejdź do następnej komórki i zaznacz jej zawartość”); Tab w ostatniej komórce —
// nowy wiersz. Zaznaczenie przez kilka komórek = całe komórki (wyrocznia
// tools/word-oracle/probe-table-select.js): pisanie zastępuje treść PIERWSZEJ komórki, reszta
// zostaje; Delete / Backspace czyści wszystkie. Backspace / Delete na brzegu komórki nic nie skleja.
// Każdy przypadek na świeżym pliku; stan = akapity PLIKU + kursor/zaznaczenie. ENGINE=webkit.

const { buildDocx } = require("../tools/word-oracle/build-docx");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const cell = (t) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p></w:tc>`;
const TBL = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell("A1") + cell("B1")}</w:tr><w:tr>${cell("A2") + cell("B2")}</w:tr></w:tbl>`;

// przypadek: [nazwa, kroki, oczekiwany stan]; stan = akapity „ ¶ ”, potem @akapit:znak [sel="…"]
const CASES = [
  ["Tab: następna komórka z zaznaczoną treścią", async (k) => { await k.at("A1", -1); await k.press("Tab"); }, 'Przed ¶ A1 ¶ B1 ¶ A2 ¶ B2 ¶ Po @2:0 sel="B1"'],
  ["Shift+Tab: poprzednia komórka z zaznaczoną treścią", async (k) => { await k.at("B1", 0); await k.press("Shift+Tab"); }, 'Przed ¶ A1 ¶ B1 ¶ A2 ¶ B2 ¶ Po @1:0 sel="A1"'],
  ["Tab i pisanie: treść komórki zastąpiona", async (k) => { await k.at("A1", 1); await k.press("Tab"); await k.type("Z"); }, "Przed ¶ A1 ¶ Z ¶ A2 ¶ B2 ¶ Po @2:1"],
  ["Tab w ostatniej komórce: nowy wiersz", async (k) => { await k.at("B2", -1); await k.press("Tab"); }, "Przed ¶ A1 ¶ B1 ¶ A2 ¶ B2 ¶ ∅ ¶ ∅ ¶ Po @5:0"],
  ["Backspace na początku komórki: nic", async (k) => { await k.at("B1", 0); await k.press("Backspace"); }, "Przed ¶ A1 ¶ B1 ¶ A2 ¶ B2 ¶ Po @2:0"],
  ["Delete na końcu komórki: nic", async (k) => { await k.at("A1", -1); await k.press("Delete"); }, "Przed ¶ A1 ¶ B1 ¶ A2 ¶ B2 ¶ Po @1:2"],
  ["przez dwie komórki + pisanie: tylko pierwsza komórka (Word)", async (k) => { await k.sel("A1", 0, "B1", 1); await k.type("X"); }, "Przed ¶ X ¶ B1 ¶ A2 ¶ B2 ¶ Po @1:1"],
  ["przez dwie komórki + Delete: obie całe wyczyszczone", async (k) => { await k.sel("A1", 1, "B1", 1); await k.press("Delete"); }, "Przed ¶ ∅ ¶ ∅ ¶ A2 ¶ B2 ¶ Po @1:0"],
  ["przez dwie komórki + Backspace: obie całe wyczyszczone", async (k) => { await k.sel("A2", 1, "B2", 1); await k.press("Backspace"); }, "Przed ¶ A1 ¶ B1 ¶ ∅ ¶ ∅ ¶ Po @3:0"],
];

async function run() {
  const bytes = await buildDocx({ body: ["Przed", TBL, "Po"] });
  const b64 = Buffer.from(bytes).toString("base64");
  const browser = await pw[ENGINE].launch({ headless: true });
  const errors = [];
  let n = 0;
  for (const [name, steps, want] of CASES) {
    const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
    await context.addInitScript(() => { sessionStorage.setItem("introPlayed", "true"); window.__dwbAutoDialog = "accept"; });
    const page = await context.newPage();
    page.on("pageerror", (e) => errors.push(e.message));
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.evaluate(async ([b, name]) => { const bin = atob(b); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); await ingestFile(new File([u], name, { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" })); }, [b64, `tabela${n++}.docx`]);
    const idle = () => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending && !!document.querySelector(".docx-preview-host"), null, { timeout: 15000 }).then(() => page.waitForTimeout(500));
    await idle();
    await page.evaluate(() => appFrame.setReadOnly(false));
    await idle();
    const node = (txt, off) => `(() => { const p = [...document.querySelectorAll(".docx-preview-host p")].find((x) => x.textContent === ${JSON.stringify(txt)}); const tn = document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode(); return [tn, ${off} < 0 ? tn.data.length : ${off}, p]; })()`;
    const k = {
      at: (txt, off) => page.evaluate(`(() => { const [n, o, p] = ${node(txt, off)}; const r = document.createRange(); r.setStart(n, o); r.collapse(true); p.closest(".docx-edit-root").focus(); getSelection().removeAllRanges(); getSelection().addRange(r); })()`),
      sel: (t1, o1, t2, o2) => page.evaluate(`(() => { const [a, x, p] = ${node(t1, o1)}; const [b, y] = ${node(t2, o2)}; const r = document.createRange(); r.setStart(a, x); r.setEnd(b, y); p.closest(".docx-edit-root").focus(); getSelection().removeAllRanges(); getSelection().addRange(r); })()`),
      press: (key) => page.keyboard.press(key),
      type: (s) => page.keyboard.type(s),
    };
    try {
      await steps(k);
      await idle();
      const got = await page.evaluate(async () => {
        await mergeInlineEditsIntoBytes();
        const d = await getDocumentXmlDom(originalFileBytes);
        const ps = collectParagraphElements(d.documentElement, "all").map((p) => getParagraphText(p) || "∅");
        const s = getSelection();
        const nd = s.anchorNode;
        const pe = nd && (nd.nodeType === 1 ? nd : nd.parentElement).closest("p");
        const all = collectPreviewParagraphElements(document.querySelector(".docx-preview-host"));
        return `${ps.join(" ¶ ")} @${all.indexOf(pe)}:${s.anchorOffset}${s.isCollapsed ? "" : ` sel="${s.toString()}"`}`;
      });
      check(name, got === want, `chce: ${want}\n   jest: ${got}`);
    } catch (e) { check(name, false, e.message.split("\n")[0]); }
    await context.close();
  }
  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `\n   ${r.detail}` : ""}`);
    if (!r.ok) failed++;
  }
  console.log(`\n[${ENGINE}] ${results.length - failed}/${results.length} ${failed ? "FAIL" : "OK"}`);
  if (failed) process.exit(1);
}

run().catch((e) => { console.error(e); process.exit(1); });
