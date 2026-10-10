// redline-playwright.js — silnik poprawek Worda z dwóch wersji (docx-redline.js, 2026-10-10).
//
// Dla każdego przykładu z docs/samples: wersja „po” = zmienione słowa, dopisek, usunięty akapit,
// nowy akapit, zmiana w komórce tabeli. Plik z poprawkami musi:
//   - po „Akceptuj wszystkie” dać dokładnie tekst wersji „po”,
//   - po „Odrzuć wszystkie” dać dokładnie tekst wersji „przed” (akapity, puste też, w kolejności),
//   - przejść walidator Open XML SDK bez nowych błędów (jeśli jest .NET).
// Akceptacja/odrzucenie — ten sam kod co panel Recenzja (applyRevisionsToDoc). ENGINE=webkit.

const fs = require("fs");
const os = require("os");
const path = require("path");
const pw = require("playwright");
const { APP_URL } = require("./docx-test-helpers");
const ooxml = require("./ooxml-validate");

const ENGINE = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const SAMPLES = path.resolve(__dirname, "../docs/samples");
const files = fs.readdirSync(SAMPLES).filter((f) => f.endsWith(".docx"));
const OUT = fs.mkdtempSync(path.join(os.tmpdir(), "dwb-redline-"));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block" });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => ensureDocLibs(true));
  const toValidate = [];

  for (const f of files) {
    const b0 = fs.readFileSync(path.join(SAMPLES, f)).toString("base64");
    const res = await page.evaluate(async (b64) => {
      const bytes0 = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
      const W = W_NS;
      const zip = await JSZip.loadAsync(bytes0);
      const doc = new DOMParser().parseFromString(await zip.file("word/document.xml").async("string"), "application/xml");
      const paras = collectParagraphElements(doc.documentElement, "all");
      const simple = (p) => !rlComplex(p) && Array.from(p.getElementsByTagNameNS(W, "t")).map((t) => t.textContent).join("").split(/\s+/).filter(Boolean).length >= 3;
      const body = paras.filter((p) => p.parentNode.localName === "body" && simple(p));
      const cell = paras.find((p) => p.parentNode.localName === "tc" && simple(p));
      const did = [];
      // 1) zmiana słowa + dopisek
      if (body[0]) {
        const t = Array.from(body[0].getElementsByTagNameNS(W, "t")).find((x) => /\S+\s+\S+/.test(x.textContent));
        if (t) { t.textContent = t.textContent.replace(/(\S+\s+)(\S+)/, "$1ZMIANA"); did.push("change"); }
        const last = Array.from(body[0].getElementsByTagNameNS(W, "t")).pop();
        last.textContent += " dopisek końcowy";
      }
      // 2) usunięty akapit
      if (body[2]) { body[2].parentNode.removeChild(body[2]); did.push("delete"); }
      // 3) nowy akapit (za czwartym)
      if (body[3]) {
        const np = doc.createElementNS(W, "w:p");
        const r = doc.createElementNS(W, "w:r"); const t = doc.createElementNS(W, "w:t"); t.textContent = "Nowy akapit testowy";
        r.appendChild(t); np.appendChild(r);
        body[3].parentNode.insertBefore(np, body[3].nextSibling);
        did.push("insert");
      }
      // 4) komórka tabeli
      if (cell) { const t = cell.getElementsByTagNameNS(W, "t")[0]; t.textContent = `${t.textContent} w komórce`; did.push("cell"); }
      zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
      const bytes1 = await zip.generateAsync({ type: "uint8array" });
      const { bytes: red, stats } = await dwbRedline(bytes0, bytes1, { author: "Test Redline", date: "2026-10-10T10:00:00Z" });
      const texts = async (bytes, mode) => {
        const z = await JSZip.loadAsync(bytes);
        const d = new DOMParser().parseFromString(await z.file("word/document.xml").async("string"), "application/xml");
        if (mode) applyRevisionsToDoc(d, mode, null);
        return collectParagraphElements(d.documentElement, "all").map((p) => Array.from(p.getElementsByTagNameNS(W, "t")).map((t) => t.textContent).join(""));
      };
      // przykłady z WCZEŚNIEJSZYMI poprawkami (Recenzja): odrzucenie/akceptacja dotyczy też ich
      const [t0, t1, acc, rej] = await Promise.all([texts(bytes0, "reject"), texts(bytes1, "accept"), texts(red, "accept"), texts(red, "reject")]);
      const rz = await JSZip.loadAsync(red);
      const rx = await rz.file("word/document.xml").async("string");
      const firstDiff = (a, b) => { const i = a.findIndex((x, k) => x !== b[k]); return i < 0 && a.length === b.length ? null : { i, a: a[i], b: b[i], la: a.length, lb: b.length }; };
      let b64out = ""; const chunk = 0x8000;
      for (let i = 0; i < red.length; i += chunk) b64out += String.fromCharCode(...red.subarray(i, i + chunk));
      return {
        did, stats,
        accept: firstDiff(acc, t1), reject: firstDiff(rej, t0),
        hasIns: /<w:ins /.test(rx), hasDel: /<w:del /.test(rx), hasDelText: /<w:delText/.test(rx), hasParaDel: /<w:rPr><w:del /.test(rx), hasParaIns: /<w:rPr><w:ins /.test(rx),
        b64: btoa(b64out),
      };
    }, b0);
    check(`${f}: Akceptuj wszystkie = wersja „po”`, res.accept === null, JSON.stringify(res.accept));
    check(`${f}: Odrzuć wszystkie = wersja „przed”`, res.reject === null, JSON.stringify(res.reject));
    if (res.did.includes("change")) check(`${f}: zmiana słowa → w:del (w:delText) + w:ins`, res.hasIns && res.hasDel && res.hasDelText, JSON.stringify(res.stats));
    if (res.did.includes("delete")) check(`${f}: usunięty akapit — znacznik akapitu usunięty`, res.hasParaDel, JSON.stringify(res.stats));
    if (res.did.includes("insert")) check(`${f}: nowy akapit — znacznik akapitu wstawiony`, res.hasParaIns, JSON.stringify(res.stats));
    const out = path.join(OUT, f.replace(/\.docx$/, ".redline.docx"));
    fs.writeFileSync(out, Buffer.from(res.b64, "base64"));
    toValidate.push([path.join(SAMPLES, f), out, f]);
  }

  // poprawki z dwóch identycznych wersji — bez zmian
  const same = await page.evaluate(async () => {
    const b = new Uint8Array(await (await fetch("docs/samples/headings-sample.docx")).arrayBuffer());
    const { bytes, stats } = await dwbRedline(b, b, { author: "T" });
    const x = await (await JSZip.loadAsync(bytes)).file("word/document.xml").async("string");
    return { stats, ins: /<w:ins /.test(x), del: /<w:del /.test(x) };
  });
  check("dwie identyczne wersje — zero poprawek", !same.ins && !same.del && same.stats.changed + same.stats.added + same.stats.removed === 0, JSON.stringify(same));

  if (ooxml.available()) {
    const all = ooxml.validate(toValidate.flatMap(([a, b]) => [a, b]));
    const by = new Map(all.map((r) => [path.resolve(r.file), r]));
    toValidate.forEach(([orig, red, f]) => {
      const fresh = ooxml.newErrors(by.get(path.resolve(orig)), by.get(path.resolve(red)));
      check(`${f}: plik z poprawkami zgodny ze schematem (Open XML SDK)`, !fresh.length, fresh.slice(0, 3).map((e) => `${e.part} ${e.path} — ${e.description}`).join(" | "));
    });
  } else console.log("(walidator .NET niedostępny — pomijam sprawdzenie schematu)");

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ silnik poprawek [${ENGINE}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
