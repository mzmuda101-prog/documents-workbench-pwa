#!/usr/bin/env node
/** Bundle docx-preview for browser (global docx.renderAsync). */
import esbuild from "esbuild";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const OUT = path.join(ROOT, "lib", "docx-preview.bundle.js");
const ENTRY = path.join(ROOT, "node_modules", "docx-preview", "dist", "docx-preview.js");

async function main() {
  if (!fs.existsSync(ENTRY)) {
    console.warn("⚠️  docx-preview not installed — run npm install first");
    return;
  }
  await esbuild.build({
    entryPoints: [ENTRY],
    outfile: OUT,
    bundle: true,
    format: "iife",
    globalName: "docx",
    minify: true,
    legalComments: "none",
    target: "es2020",
  });
  patchBundle();
  const kb = (fs.statSync(OUT).size / 1024).toFixed(1);
  console.log(`✅  lib/docx-preview.bundle.js (${kb} KB)`);
}

// Łatki na zbudowanej bibliotece (nazwy po minifikacji się zmieniają — szukamy wzorcem).
// 1) „Od nowej strony” (w:pageBreakBefore) docx-preview bierze tylko ze STYLU akapitu;
//    bezpośrednie ustawienie akapitu (Ctrl+Enter w Documents Workbench, Word „Podział
//    wiersza i strony → Podział strony przed”) nie łamało strony w podglądzie.
function patchBundle() {
  let src = fs.readFileSync(OUT, "utf8");
  const re = /(\w+)\.type==(\w+)\.Paragraph&&this\.findStyle\(\1\.styleName\)\?\.paragraphProps\?\.pageBreakBefore/;
  if (!re.test(src)) throw new Error("łatka pageBreakBefore: wzorzec nie pasuje (nowa wersja docx-preview?)");
  src = src.replace(re, (_, p, T) => `${p}.type==${T}.Paragraph&&(${p}.pageBreakBefore||this.findStyle(${p}.styleName)?.paragraphProps?.pageBreakBefore)`);
  console.log("  ✅  łatka: pageBreakBefore akapitu");

  // 2) Obraz „pływający” (wp:anchor, wrapNone) położony względem STRONY: docx-preview stawiał go
  //    względem akapitu, w którym siedzi kotwica, i ignorował „za tekstem”. PDF → DOCX kotwiczy
  //    mapki/skany/warstwę grafiki do strony — bez tej łatki lądowały w złym miejscu.
  const reBehind = /(\w+)\.boolAttr\((\w+),"behindDoc"\);let (\w+)=\{relative:"page",align:"left",offset:"0"\},(\w+)=\{relative:"page",align:"top",offset:"0"\};/;
  const mB = src.match(reBehind);
  if (!mB) throw new Error("łatka anchor: wzorzec behindDoc nie pasuje (nowa wersja docx-preview?)");
  const [, P, NODE, PX, PY] = mB;
  src = src.replace(reBehind, `var __dwbBehind=${P}.boolAttr(${NODE},"behindDoc");let ${PX}={relative:"page",align:"left",offset:"0"},${PY}={relative:"page",align:"top",offset:"0"};`);
  const reWrap = /(\w+)=="wrapNone"\?\((\w+)\.cssStyle\.display="block",\2\.cssStyle\.position="relative",/;
  if (!reWrap.test(src)) throw new Error("łatka anchor: wzorzec wrapNone nie pasuje");
  src = src.replace(reWrap, (_, O, E) => `${O}=="wrapNone"?(${E}.cssStyle.display="block",${E}.cssStyle.position=(${PX}.relative=="page"&&${PY}.relative=="page"?"absolute":"relative"),${PX}.relative=="page"&&${PY}.relative=="page"&&(${E}.cssStyle["z-index"]=__dwbBehind?"-1":"1",${E}.cssStyle["pointer-events"]="none"),`);
  console.log("  ✅  łatka: obraz zakotwiczony do strony (absolute, za tekstem)");

  // 3) Tabulatory: docx-preview rysuje każdy jako stałą spację (&emsp;), a rozkład wg pozycji
  //    tabulatorów ma tylko w trybie „experimental” (po 500 ms, na współrzędnych ekranu — psuje
  //    się przy powiększeniu). Zostawiamy na znaczniku pozycje z akapitu; rozkład robi
  //    app/docx-render-fixes.js (layoutTabStops) od razu po renderze.
  const reTab = /renderTab\((\w+)\)\{var (\w+)=this\.createElement\("span"\);if\(\2\.innerHTML="&emsp;",this\.options\.experimental\)\{\2\.className=this\.tabStopClass\(\);var (\w+)=([\w$]+)\(\1,([\w$]+)\.Paragraph\)\?\.tabs;/;
  if (!reTab.test(src)) throw new Error("łatka tabulatorów: wzorzec renderTab nie pasuje (nowa wersja docx-preview?)");
  src = src.replace(reTab, (m, T, E, R, F, D) => `renderTab(${T}){var ${E}=this.createElement("span");${E}.className="docx-tab";try{var __dwbStops=${F}(${T},${D}.Paragraph)?.tabs;__dwbStops&&__dwbStops.length&&(${E}.dataset.stops=JSON.stringify(__dwbStops.map(s=>[parseFloat(s.position),s.leader||"none",s.style||"left"])));this.defaultTabSize&&(${E}.dataset.dt=parseFloat(this.defaultTabSize))}catch(__e){}if(${E}.innerHTML="&emsp;",this.options.experimental){${E}.className+=" "+this.tabStopClass();var ${R}=${F}(${T},${D}.Paragraph)?.tabs;`);
  fs.writeFileSync(OUT, src);
  console.log("  ✅  łatka: tabulatory z pozycjami (dane dla layoutTabStops)");
}

// pdf.js (konwersja PDF → DOCX): moduł główny + worker + wasm (obrazy JPEG2000/JBIG2, profile
// kolorów) + cmapy (PDF-y z tekstem CJK). Wszystko lokalnie — plik nie wychodzi z urządzenia.
function vendorPdfjs() {
  const SRC = path.join(ROOT, "node_modules", "pdfjs-dist");
  if (!fs.existsSync(SRC)) {
    console.warn("⚠️  pdfjs-dist not installed — run npm install first");
    return;
  }
  const OUT_DIR = path.join(ROOT, "lib", "pdfjs");
  fs.rmSync(OUT_DIR, { recursive: true, force: true });
  fs.mkdirSync(path.join(OUT_DIR, "wasm"), { recursive: true });
  fs.mkdirSync(path.join(OUT_DIR, "cmaps"), { recursive: true });
  for (const f of ["pdf.min.mjs", "pdf.worker.min.mjs"]) fs.copyFileSync(path.join(SRC, "build", f), path.join(OUT_DIR, f));
  for (const f of ["jbig2.wasm", "openjpeg.wasm", "qcms_bg.wasm", "jbig2_nowasm_fallback.js", "openjpeg_nowasm_fallback.js"]) {
    const from = path.join(SRC, "wasm", f);
    if (fs.existsSync(from)) fs.copyFileSync(from, path.join(OUT_DIR, "wasm", f));
  }
  for (const f of fs.readdirSync(path.join(SRC, "cmaps"))) fs.copyFileSync(path.join(SRC, "cmaps", f), path.join(OUT_DIR, "cmaps", f));
  fs.copyFileSync(path.join(SRC, "LICENSE"), path.join(OUT_DIR, "LICENSE"));
  const ver = JSON.parse(fs.readFileSync(path.join(SRC, "package.json"), "utf8")).version;
  fs.writeFileSync(path.join(OUT_DIR, "VERSION"), ver + "\n");
  console.log(`✅  lib/pdfjs (pdf.js ${ver})`);
}

// tesseract.js (rozpoznawanie tekstu na skanach, PDF → DOCX): skrypt główny, worker, rdzeń wasm
// (SIMD + zwykły, wariant LSTM) i dane języka polskiego. Ładowane dopiero przy pierwszym skanie.
function vendorTesseract() {
  const T = path.join(ROOT, "node_modules", "tesseract.js");
  const CORE = path.join(ROOT, "node_modules", "tesseract.js-core");
  const POL = path.join(ROOT, "node_modules", "@tesseract.js-data", "pol", "4.0.0_best_int", "pol.traineddata.gz");
  if (!fs.existsSync(T) || !fs.existsSync(CORE) || !fs.existsSync(POL)) {
    console.warn("⚠️  tesseract.js / dane pol nie zainstalowane — run npm install first");
    return;
  }
  const OUT_DIR = path.join(ROOT, "lib", "tesseract");
  fs.rmSync(OUT_DIR, { recursive: true, force: true });
  fs.mkdirSync(path.join(OUT_DIR, "lang"), { recursive: true });
  fs.copyFileSync(path.join(T, "dist", "tesseract.min.js"), path.join(OUT_DIR, "tesseract.min.js"));
  fs.copyFileSync(path.join(T, "dist", "worker.min.js"), path.join(OUT_DIR, "worker.min.js"));
  for (const f of ["tesseract-core-simd-lstm.js", "tesseract-core-simd-lstm.wasm", "tesseract-core-lstm.js", "tesseract-core-lstm.wasm"]) {
    // obok worker.min.js: rdzeń szuka swojego .wasm względem adresu workera
    fs.copyFileSync(path.join(CORE, f), path.join(OUT_DIR, f));
  }
  fs.copyFileSync(POL, path.join(OUT_DIR, "lang", "pol.traineddata.gz"));
  fs.copyFileSync(path.join(T, "LICENSE.md"), path.join(OUT_DIR, "LICENSE.md"));
  const ver = JSON.parse(fs.readFileSync(path.join(T, "package.json"), "utf8")).version;
  fs.writeFileSync(path.join(OUT_DIR, "VERSION"), `tesseract.js ${ver}, dane pol 4.0.0_best_int\n`);
  console.log(`✅  lib/tesseract (tesseract.js ${ver})`);
}

main().then(vendorPdfjs).then(vendorTesseract).catch((err) => {
  console.error("❌  vendor-libs failed:", err.message || err);
  process.exit(1);
});
