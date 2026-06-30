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
  const kb = (fs.statSync(OUT).size / 1024).toFixed(1);
  console.log(`✅  lib/docx-preview.bundle.js (${kb} KB)`);
}

main().catch((err) => {
  console.error("❌  vendor-libs failed:", err.message || err);
  process.exit(1);
});
