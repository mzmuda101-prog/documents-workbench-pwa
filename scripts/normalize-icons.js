#!/usr/bin/env node
// normalize-icons.js — skaluje wgrane ikony PWA do standardowych rozmiarów.
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const IMG = path.resolve(__dirname, "../assets/images");
const CANDIDATES = ["icon-source.png", "apple-touch-icon.png", "icon-512.png", "favicon.png"];
const TMP = path.join(IMG, ".icon-normalize-src.png");

function pickSource() {
  for (const name of CANDIDATES) {
    const p = path.join(IMG, name);
    if (fs.existsSync(p)) return p;
  }
  throw new Error("Brak pliku źródłowego ikony w assets/images/");
}

function sips(size, out) {
  execSync(`sips -z ${size} ${size} "${TMP}" --out "${out}"`, { stdio: "inherit" });
}

const src = pickSource();
fs.copyFileSync(src, TMP);
sips(512, path.join(IMG, "icon-512.png"));
sips(192, path.join(IMG, "icon-192.png"));
sips(180, path.join(IMG, "apple-touch-icon.png"));
sips(64, path.join(IMG, "favicon.png"));
fs.unlinkSync(TMP);
console.log("✅  Ikony PWA znormalizowane (64 / 180 / 192 / 512 px)");
