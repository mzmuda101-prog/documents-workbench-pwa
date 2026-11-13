// word-shortcuts-playwright.js — skróty formatowania jak w Wordzie (2026-11-13).
//
// Ctrl/⌘+B / I / U z kursorem W ŚRODKU słowa (bez zaznaczenia): format dostaje całe słowo, kursor
// zostaje w miejscu; na końcu słowa — tylko dalsze pisanie (słowo bez zmian). Ctrl/⌘+L / E / R / J —
// wyrównanie akapitu. Ctrl/⌘+Alt+1 / 2 / 3 z kursorem w tekście — Nagłówek 1 / 2 / 3 (poza tekstem
// dalej obszary ekranu). Sprawdza PLIK. ENGINE=webkit.

const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const M = process.platform === "darwin" ? "Meta" : "Control";

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); setSidebarOpen(false); loadSampleDocument("sample"); });
  await page.waitForSelector(".docx-preview-host p");
  const idle = () => page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && !inlineLocksPending, null, { timeout: 15000 }).then(() => page.waitForTimeout(600));
  await idle();
  await page.evaluate(() => appFrame.setReadOnly(false));
  await idle();
  const caret = (needle, off) => page.evaluate(([needle, off]) => {
    const p = [...document.querySelectorAll(".docx-preview-host p")].find((x) => x.textContent.includes(needle));
    const tw = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = tw.nextNode(); n; n = tw.nextNode()) {
      const i = n.data.indexOf(needle);
      if (i < 0) continue;
      const r = document.createRange(); r.setStart(n, i + off); r.collapse(true);
      p.closest(".docx-edit-root").focus(); getSelection().removeAllRanges(); getSelection().addRange(r);
      return;
    }
  }, [needle, off]);
  const para = (needle) => page.evaluate(async (needle) => {
    await mergeInlineEditsIntoBytes();
    const d = await getDocumentXmlDom(originalFileBytes);
    const p = collectParagraphElements(d.documentElement, "all").find((x) => getParagraphText(x).includes(needle));
    const bold = Array.from(p.getElementsByTagNameNS(W_NS, "r")).filter((r) => r.getElementsByTagNameNS(W_NS, "b").length).map((r) => r.textContent).join("|");
    const jc = p.getElementsByTagNameNS(W_NS, "jc")[0]?.getAttributeNS(W_NS, "val") || "";
    const style = p.getElementsByTagNameNS(W_NS, "pStyle")[0]?.getAttributeNS(W_NS, "val") || "";
    return { bold, jc, style };
  }, needle);

  // ── B w środku słowa: całe słowo, kursor w miejscu ─────────────────────────
  await caret("przykładowy", 4);
  await page.keyboard.press(`${M}+KeyB`);
  await idle();
  const b1 = await para("przykładowy");
  const c1 = await page.evaluate(() => { const s = getSelection(); return { collapsed: s.isCollapsed, text: s.anchorNode?.data, at: s.anchorOffset }; });
  check("Ctrl/⌘+B w środku słowa: pogrubione całe słowo", b1.bold === "przykładowy", JSON.stringify(b1));
  check("…kursor zostaje w tym samym miejscu słowa", c1.collapsed && c1.text === "przykładowy" && c1.at === 4, JSON.stringify(c1));
  await page.keyboard.type("Q");
  await idle();
  check("…pisanie dalej w słowie (pogrubione)", (await para("przyQkładowy")).bold === "przyQkładowy", JSON.stringify(await para("przyQkładowy")));

  // ── B na końcu słowa: słowo bez zmian, pogrubione tylko dalsze pisanie ─────
  await caret("dokument", 8);
  await page.keyboard.press(`${M}+KeyB`);
  await page.keyboard.type("X");
  await idle();
  const b2 = await para("dokumentX");
  check("Ctrl/⌘+B na końcu słowa: słowo bez zmian, pogrubione tylko dopisane", /(^|\|)X(\||$)/.test(b2.bold) && !/dokument/.test(b2.bold), JSON.stringify(b2));

  // ── wyrównanie ─────────────────────────────────────────────────────────────
  for (const [key, want] of [["KeyE", "center"], ["KeyR", "right"], ["KeyJ", "both"], ["KeyL", "left"]]) {
    await caret("testowy", 2);
    await page.keyboard.press(`${M}+${key}`);
    await idle();
    const got = (await para("testowy")).jc;
    check(`Ctrl/⌘+${key.slice(3)}: wyrównanie „${want}”`, want === "left" ? got === "left" || got === "" : got === want, got);
  }

  // ── Nagłówek 1 / 2 / 3 z kursorem w tekście ────────────────────────────────
  for (const [d, want] of [["1", "Heading1"], ["2", "Heading2"], ["3", "Heading3"]]) {
    await caret("testowy", 2);
    await page.keyboard.press(`${M}+Alt+Digit${d}`);
    await idle();
    check(`Ctrl/⌘+Alt+${d} w tekście: Nagłówek ${d}`, (await para("testowy")).style === want, (await para("testowy")).style);
  }
  // poza tekstem — dalej obszar ekranu (panel)
  await page.evaluate(() => { getSelection().removeAllRanges(); document.getElementById("fmtParaStyle")?.focus(); });
  await page.keyboard.press(`${M}+Alt+Digit1`);
  await page.waitForTimeout(400);
  check("Ctrl/⌘+Alt+1 poza tekstem: panel (jak dotąd)", await page.evaluate(() => !!document.activeElement?.closest("#toolsPanel, .tools-panel, aside")), await page.evaluate(() => document.activeElement?.id || document.activeElement?.className));

  // ── Wstaw → Symbol: znak w miejscu kursora, potem w „Ostatnio używanych” ───
  await caret("Szukaj", 6);
  await page.click("#insertMenuBtn");
  await page.locator(".compose-item", { hasText: /^Symbol/ }).click();
  await page.locator(".symbol-pick button", { hasText: "§" }).first().click();
  await idle();
  const symText = await page.evaluate(async () => { await mergeInlineEditsIntoBytes(); const d = await getDocumentXmlDom(originalFileBytes); return collectParagraphElements(d.documentElement, "all").map(getParagraphText).find((t) => t.startsWith("Szukaj")); });
  check("Wstaw → Symbol: § w miejscu kursora (w pliku)", /^Szukaj§:/.test(symText || ""), symText);
  await page.click("#insertMenuBtn");
  await page.locator(".compose-item", { hasText: /^Symbol/ }).click();
  const recent = await page.evaluate(() => { const g = document.querySelectorAll(".symbol-pick"); return g.length === 2 ? g[0].textContent : ""; });
  check("…§ w „Ostatnio używanych”", recent === "§", recent);
  // podzbiory jak w oknie Symbol Worda: Matematyka → ∑, Greckie → α
  await page.locator(".symbol-sets button", { hasText: "Matematyka" }).click();
  const mathSet = await page.evaluate(() => ({ pressed: document.querySelector('.symbol-sets [data-set="math"]').getAttribute("aria-pressed"), grids: document.querySelectorAll(".symbol-pick").length, label: [...document.querySelectorAll(".symbol-pick button")].find((b) => b.textContent === "∑")?.getAttribute("aria-label") }));
  check("Symbol → Matematyka: ∑ z nazwą, jedna siatka zamiast ogólnej", mathSet.pressed === "true" && mathSet.grids === 2 && mathSet.label === "suma", JSON.stringify(mathSet));
  await page.locator(".symbol-pick button", { hasText: "∑" }).click();
  await idle();
  await page.click("#insertMenuBtn");
  await page.locator(".compose-item", { hasText: /^Symbol/ }).click();
  await page.locator(".symbol-sets button", { hasText: "Greckie" }).click();
  await page.locator(".symbol-pick button", { hasText: "α" }).first().click();
  await idle();
  const mathText = await page.evaluate(async () => { await mergeInlineEditsIntoBytes(); const d = await getDocumentXmlDom(originalFileBytes); return collectParagraphElements(d.documentElement, "all").map(getParagraphText).find((t) => t.startsWith("Szukaj")); });
  check("Wstaw → Symbol: ∑ i α w pliku", /^Szukaj§∑α:/.test(mathText || ""), mathText);
  await page.click("#insertMenuBtn");
  await page.locator(".compose-item", { hasText: /^Symbol/ }).click();
  const recent2 = await page.evaluate(() => document.querySelectorAll(".symbol-pick")[0].textContent);
  check("…α, ∑, § w „Ostatnio używanych”", recent2 === "α∑§", recent2);
  await page.keyboard.press("Escape");

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
