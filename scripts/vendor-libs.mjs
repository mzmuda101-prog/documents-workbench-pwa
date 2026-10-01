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
  fs.writeFileSync(OUT, src);
  console.log("  ✅  łatka: pageBreakBefore akapitu");
}

main().catch((err) => {
  console.error("❌  vendor-libs failed:", err.message || err);
  process.exit(1);
});
