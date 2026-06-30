#!/usr/bin/env node
import esbuild from "esbuild";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const DIST = path.join(ROOT, "dist");
const COPY_DIRS = ["assets", "lib", "docs"];
const COPY_FILES = ["manifest.json", "index.html", "sw.js", "README.md", "LICENSE"];

function readVersion() {
  const sw = fs.readFileSync(path.join(ROOT, "sw.js"), "utf8");
  const m = sw.match(/CACHE_VERSION\s*=\s*"(\d{8}-\d{2})"/);
  if (!m) throw new Error("Brak CACHE_VERSION w sw.js");
  return m[1];
}

function rmrf(dir) {
  if (!fs.existsSync(dir)) return;
  fs.rmSync(dir, { recursive: true, force: true });
}

function copyRecursive(src, dest) {
  const stat = fs.statSync(src);
  if (stat.isDirectory()) {
    fs.mkdirSync(dest, { recursive: true });
    for (const name of fs.readdirSync(src)) {
      if (name === "node_modules" || name === ".git" || name === "dist") continue;
      copyRecursive(path.join(src, name), path.join(dest, name));
    }
    return;
  }
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.copyFileSync(src, dest);
}

async function minifyFile(inFile, outFile) {
  const ext = path.extname(inFile);
  if (ext === ".css") {
    const css = await esbuild.transform(fs.readFileSync(inFile, "utf8"), { loader: "css", minify: true });
    fs.mkdirSync(path.dirname(outFile), { recursive: true });
    fs.writeFileSync(outFile, css.code);
    return;
  }
  if (ext === ".js") {
    await esbuild.build({
      entryPoints: [inFile],
      outfile: outFile,
      minify: true,
      bundle: false,
      legalComments: "none",
      target: "es2020",
    });
  }
}

async function main() {
  const version = readVersion();
  console.log(`\n📦  Build release → dist/ (wersja ${version})\n`);
  rmrf(DIST);
  fs.mkdirSync(DIST, { recursive: true });
  for (const dir of COPY_DIRS) {
    const src = path.join(ROOT, dir);
    if (fs.existsSync(src)) copyRecursive(src, path.join(DIST, dir));
  }
  for (const file of COPY_FILES) {
    const src = path.join(ROOT, file);
    if (fs.existsSync(src)) {
      fs.mkdirSync(path.dirname(path.join(DIST, file)), { recursive: true });
      fs.copyFileSync(src, path.join(DIST, file));
    }
  }
  const appDir = path.join(ROOT, "app");
  for (const file of fs.readdirSync(appDir).filter((f) => f.endsWith(".js"))) {
    await minifyFile(path.join(appDir, file), path.join(DIST, "app", file));
    console.log(`  ✅  app/${file}`);
  }
  await minifyFile(path.join(ROOT, "styles", "app.css"), path.join(DIST, "styles", "app.css"));
  console.log(`  ✅  styles/app.css\n🎉  Gotowe: dist/\n`);
}

main().catch((err) => {
  console.error("❌  Build failed:", err.message || err);
  process.exit(1);
});
