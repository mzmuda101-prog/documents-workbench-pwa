// justify-last-line-playwright.js — ostatnia linijka wyjustowanego akapitu na EKRANIE ma zwykłe
// spacje, jak w Wordzie (lista 3.11). CSS ściska spacje o ¼ we wszystkich linijkach (łamanie jak
// Word), więc bez poprawki ostatnia, niewyjustowana linijka wyglądała na ściśniętą.
//   • odstęp między słowami w ostatniej linijce = odstęp w takim samym akapicie do lewej (±0,6 px)
//   • liczba linijek akapitu bez zmian (dopełnienie nigdy nie łamie linijki)
//   • po dopisaniu słowa w Edycji dopełnienie liczy się od nowa

const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const JSZip = require("jszip");
const { APP_URL } = require("./docx-test-helpers");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const TEXT = "Wykonawca zobowiązuje się wykonać przedmiot umowy z należytą starannością, zgodnie z obowiązującymi przepisami prawa oraz zasadami wiedzy technicznej, a także przekazać zamawiającemu wszystkie dokumenty niezbędne do odbioru prac w terminie siedmiu dni.";

async function buildDocx() {
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(__dirname, "..", "docs", "samples", "review-sample.docx")));
  const xml = await zip.file("word/document.xml").async("string");
  const head = xml.slice(0, xml.indexOf("<w:body>") + 8);
  const sect = xml.match(/<w:sectPr[\s\S]*<\/w:sectPr>/)[0];
  const para = (jc, t) => `<w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  zip.file("word/document.xml", `${head}${para("both", TEXT)}${para("left", TEXT)}${sect}</w:body></w:document>`);
  const file = path.join(os.tmpdir(), `dwb-justify-last-${process.pid}.docx`);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer" }));
  return file;
}

// średni odstęp między słowami w linijce nr `which` ("last" | 0) akapitu
const GAPS = ({ sel, which }) => {
  const ps = [...document.querySelectorAll("section.docx p")].filter((x) => x.textContent.startsWith("Wykonawca zobowiązuje"));
  const p = ps[sel];
  const range = document.createRange();
  const words = [];
  const tw = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  for (let t = tw.nextNode(); t; t = tw.nextNode()) {
    const re = /\S+/g;
    let m;
    while ((m = re.exec(t.data))) {
      range.setStart(t, m.index); range.setEnd(t, m.index + 1);
      const a = range.getBoundingClientRect();
      range.setStart(t, m.index + m[0].length - 1); range.setEnd(t, m.index + m[0].length);
      const b = range.getBoundingClientRect();
      words.push({ top: Math.round(a.top), x0: a.left, x1: b.right });
    }
  }
  const tops = [...new Set(words.map((w) => w.top))];
  const top = which === "last" ? tops[tops.length - 1] : tops[which];
  const line = words.filter((w) => w.top === top);
  const gaps = line.slice(1).map((w, i) => w.x0 - line[i].x1);
  const ws = -parseFloat(getComputedStyle(p.querySelector("span") || p).wordSpacing) || 0;
  return { lines: tops.length, end: line.length ? line[line.length - 1].x1 : 0, ws, gap: gaps.reduce((a, b) => a + b, 0) / Math.max(1, gaps.length), n: gaps.length, fill: p.style.getPropertyValue("--dwb-last-fill"), justify: p.classList.contains("dwb-justify") };
};

(async () => {
  const file = await buildDocx();
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 1000 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.locator("#fileInput").setInputFiles(file);
  await page.waitForFunction(() => document.getElementById("loadingOverlay")?.classList.contains("hidden") && document.querySelector("section.docx p.dwb-justify"), null, { timeout: 30000 });
  await page.waitForTimeout(600);
  const fails = [];
  const check = (name, ok, d) => { if (!ok) fails.push(`${name} — ${JSON.stringify(d)}`); };

  const j = await page.evaluate(GAPS, { sel: 0, which: "last" });
  const ref = await page.evaluate(GAPS, { sel: 1, which: 0 });
  check("akapit wyjustowany ma kilka linijek", j.justify && j.lines >= 2 && j.n >= 2, j);
  // WebKit: prostokąty pojedynczych znaków są niedokładne — porównanie odstępów tylko w Chromium
  if (ENGINE === "chromium") check("ostatnia linijka: zwykłe spacje jak w akapicie do lewej", Math.abs(j.gap - ref.gap) <= 0.6, { last: j.gap, normal: ref.gap, fill: j.fill });
  // bez dopełnienia (stan sprzed poprawki) — spacje ściśnięte: kontrola, że pomiar coś mierzy
  const off = await page.evaluate((G) => {
    const p = [...document.querySelectorAll("section.docx p")].find((x) => x.textContent.startsWith("Wykonawca zobowiązuje"));
    const v = p.style.getPropertyValue("--dwb-last-fill");
    p.style.removeProperty("--dwb-last-fill");
    const r = (0, eval)(`(${G})`)({ sel: 0, which: "last" });
    p.style.setProperty("--dwb-last-fill", v);
    return r;
  }, GAPS.toString());
  if (ENGINE === "chromium") check("kontrola: bez dopełnienia spacje węższe", off.gap < ref.gap - 0.6, { off: off.gap, normal: ref.gap });
  // każdy odstęp odzyskuje dokładnie swoje ściśnięcie: koniec linijki przesunięty o n × |word-spacing|
  check("koniec ostatniej linijki przesunięty o ściśnięcie spacji", j.ws > 0.3 && Math.abs((j.end - off.end) - j.n * j.ws) <= 1.5, { shift: j.end - off.end, want: j.n * j.ws });
  check("liczba linijek bez zmian", off.lines === j.lines, { with: j.lines, without: off.lines });

  // Edycja: dopisz słowo na końcu → dopełnienie od nowa
  const modeBtn = page.locator('[data-mode="edit"], #modeEdit, button:has-text("Edycja")').first();
  if (await modeBtn.count()) await modeBtn.click().catch(() => {});
  await page.waitForTimeout(400);
  const before = j.fill;
  await page.evaluate(() => {
    const p = document.querySelector("section.docx p.dwb-justify");
    const t = [...p.querySelectorAll("span")].reverse().find((s) => s.firstChild && s.firstChild.nodeType === 3).firstChild;
    t.data = t.data.replace(/dni\.$/, "dni, tj.");
  });
  await page.waitForTimeout(500);
  const j2 = await page.evaluate(GAPS, { sel: 0, which: "last" });
  check("po dopisaniu: dopełnienie przeliczone", j2.fill && j2.fill !== before && (ENGINE !== "chromium" || Math.abs(j2.gap - ref.gap) <= 0.6), { before, after: j2.fill, gap: j2.gap, normal: ref.gap });

  if (errors.length) fails.push("błędy strony: " + errors.join(" | "));
  await browser.close();
  fs.unlinkSync(file);
  if (fails.length) { console.log(`❌ justify-last-line (${ENGINE}): ${fails.length} błędów`); fails.forEach((f) => console.log("  ✗", f)); process.exit(1); }
  console.log(`✅ justify-last-line (${ENGINE}): wszystko przeszło (odstęp ${j.gap.toFixed(2)} px vs ${ref.gap.toFixed(2)} px)`);
})().catch((e) => { console.error(e); process.exit(1); });
