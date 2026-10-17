// run-quote.js — cytaty i ramki: styl z listy stylów / gumka w Wordzie kontra aplikacja (XML po zapisie)
// node tools/word-oracle/run-quote.js; WRITE=1 → expected-quote.json (scripts/quote-frame-playwright.js)
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { buildDocx } = require("./build-docx");
const word = require("./word");
const app = require("./app");
const { QUOTE_DOC, QUOTE_OPS, normalizeParas } = require("./cases-quote");

const WSTYLE = { normal: "style normal", quote: "style quote", h1: "style heading1", title: "style title" };

(async () => {
  const bytes = await buildDocx(QUOTE_DOC);
  const key = crypto.createHash("sha1").update(bytes).update(JSON.stringify(QUOTE_OPS)).digest("hex").slice(0, 16);
  const cacheF = path.join(word.DIR, "cache", `quote-${key}.json`);
  let wx;
  if (fs.existsSync(cacheF) && !process.env.FRESH) wx = JSON.parse(fs.readFileSync(cacheF, "utf8"));
  else {
    const body = QUOTE_OPS.map((o) => (o.op === "style" ? `set style of paragraph ${o.para + 1} to ${WSTYLE[o.style]}`
      : `set ps to start of content of text object of paragraph ${o.para + 1}
    select (create range it start (ps + 3) end (ps + 3))
    clear formatting (get selection of application "Microsoft Word")`)).join("\n");
    const r = word.saveDocx(bytes, body, ["word/document.xml", "word/styles.xml"]);
    wx = { doc: r.parts["word/document.xml"], styles: r.parts["word/styles.xml"] };
    fs.mkdirSync(path.dirname(cacheF), { recursive: true });
    fs.writeFileSync(cacheF, JSON.stringify(wx));
  }
  const { browser, page } = await app.openApp();
  await app.loadBytes(page, bytes, "quote.docx");
  const a = await app.quoteOps(page, QUOTE_OPS);
  const w = await page.evaluate(normalizeParas, wx);
  await browser.close();
  if (process.env.WRITE) fs.writeFileSync(path.join(__dirname, "expected-quote.json"), `${JSON.stringify(w, null, 1)}\n`);
  let bad = 0;
  w.forEach((wp, i) => {
    const ap = a.paras[i] || {};
    if (JSON.stringify(wp) !== JSON.stringify(ap)) { bad++; console.log(`❌ #${i} ${wp.runs}\n   Word: ${JSON.stringify(wp)}\n   apka: ${JSON.stringify(ap)}`); } else console.log(`✅ #${i} ${wp.style} ${wp.runs}`);
  });
  console.log(`lista stylów przed: ${JSON.stringify(a.labels)}`);
  console.log(bad ? `\n${bad} akapitów różni się od Worda` : "\nwszystko jak w Wordzie");
})();
