// mail-merge.js — korespondencja seryjna (2026-10-27): dokument z polami {{nazwa}} + tabela danych
// (CSV albo XLSX, pierwszy wiersz = nazwy kolumn) → osobny .docx dla każdego wiersza, w jednym ZIP-ie.
//
// Pola dopasowujemy do kolumn bez wielkości liter i polskich znaków („Imię” = {{imie}}), spacja
// i łącznik = podkreślnik. Pusta komórka = puste pole; kolumna, której brak w danych — pole
// zostaje {{tak}} (ostrzeżenie przed tworzeniem). Wypełniamy treść, nagłówki, stopki i przypisy.
// Daty i liczby z XLSX jak w polskim Excelu (15.10.2026, 1 234,50), nie w zapisie SheetJS (US).
// Ładowany leniwie z panelem pól (lazy-features.js); SheetJS (lib/xlsx.full.min.js, ~1 MB) —
// dopiero przy pierwszym pliku XLSX.

let mailMergeData = null; // { name, headers: [..], keys: [..], rows: [[..]] }

const mmPickBtn = document.getElementById("mmPickBtn");
const mmFileEl = document.getElementById("mmFile");
const mmStatusEl = document.getElementById("mmStatus");
const mmOptionsEl = document.getElementById("mmOptions");
const mmNameColEl = document.getElementById("mmNameCol");
const mmRowEl = document.getElementById("mmRow");
const mmPreviewBtn = document.getElementById("mmPreviewBtn");
const mmRunBtn = document.getElementById("mmRunBtn");

function mmFoldKey(s) {
  return String(s || "").trim().toLowerCase().replace(/ł/g, "l").normalize("NFD").replace(/\p{M}/gu, "")
    .replace(/[\s-]+/g, "_").replace(/[^a-z0-9_.]/g, "");
}

// CSV: separator ; , albo tabulator (ten, którego najwięcej w 1. wierszu poza cudzysłowami),
// pola w "…" z "" w środku, BOM, CRLF.
function mmParseCsv(text) {
  text = String(text || "").replace(/^﻿/, "");
  const firstLine = text.split(/\r?\n/, 1)[0].replace(/"[^"]*"/g, "");
  const sep = [";", ",", "\t"].map((c) => [c, firstLine.split(c).length]).sort((a, b) => b[1] - a[1])[0][0];
  const rows = [];
  let row = [], cell = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') q = false; else cell += ch;
    } else if (ch === '"' && cell === "") q = true;
    else if (ch === sep) { row.push(cell); cell = ""; }
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(cell); rows.push(row); row = []; cell = "";
    } else cell += ch;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

// Tekst komórki XLSX tak, jak pokazuje go polski Excel.
function mmXlsxCellText(cell) {
  if (!cell || cell.v == null) return "";
  if (cell.t === "n" && cell.z && XLSX.SSF.is_date(cell.z)) {
    // format 14 („data krótka”) zależy od ustawień regionalnych — u nas dd.mm.rrrr
    if (cell.z === "m/d/yy" || cell.z === 14) {
      const d = XLSX.SSF.parse_date_code(cell.v);
      if (d) return `${String(d.d).padStart(2, "0")}.${String(d.m).padStart(2, "0")}.${d.y}`;
    }
    return cell.w != null ? String(cell.w) : String(cell.v);
  }
  if (cell.t === "n") {
    const w = cell.w != null ? String(cell.w) : String(cell.v);
    // zapis SheetJS jest amerykański: 1,234.50 → 1 234,50 (twarda spacja nie łamie liczby)
    return currentLang === "pl" ? w.replace(/[,.]/g, (ch) => (ch === "," ? " " : ",")) : w;
  }
  if (cell.t === "b") return cell.v ? (currentLang === "pl" ? "PRAWDA" : "TRUE") : (currentLang === "pl" ? "FAŁSZ" : "FALSE");
  return cell.w != null ? String(cell.w) : String(cell.v);
}

async function mmReadXlsx(file) {
  if (typeof XLSX === "undefined") await loadLazyScript("lib/xlsx.full.min.js");
  const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: "array", cellNF: true }) // cellNF — kod formatu (data / kwota);
  // pierwszy arkusz z danymi
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws || !ws["!ref"]) continue;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const rows = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) row.push(mmXlsxCellText(ws[XLSX.utils.encode_cell({ r, c })]));
      rows.push(row);
    }
    if (rows.some((row) => row.some((v) => v.trim()))) return rows;
  }
  return [];
}

function mmTableFromRows(rows, name) {
  const filled = rows.filter((row) => row.some((v) => String(v).trim()));
  if (filled.length < 2) return null;
  const headers = filled[0].map((h) => String(h).trim());
  return {
    name,
    headers,
    keys: headers.map(mmFoldKey),
    rows: filled.slice(1).map((row) => headers.map((_, i) => String(row[i] ?? "").trim())),
  };
}

async function mmLoadFile(file) {
  if (!file) return;
  try {
    const rows = /\.xlsx?$|\.xlsm$/i.test(file.name) ? await mmReadXlsx(file) : mmParseCsv(await file.text());
    const data = mmTableFromRows(rows, file.name);
    if (!data) { toast(t("mmNoRows"), "error"); return; }
    mailMergeData = data;
  } catch (e) {
    console.warn("mail merge read", e);
    toast(t("mmReadFailed"), "error");
    return;
  }
  await mmRefresh();
}

// Pola dokumentu ↔ kolumny danych
async function mmMatch() {
  if (typeof mergeInlineEditsIntoBytes === "function") await mergeInlineEditsIntoBytes();
  const scan = await scanPlaceholders(originalFileBytes);
  const names = scan.fields.map((f) => f.name);
  const map = {}; // nazwa pola → indeks kolumny
  const missing = [];
  names.forEach((n) => {
    const i = mailMergeData.keys.indexOf(mmFoldKey(n));
    if (i >= 0) map[n] = i; else missing.push(n);
  });
  return { names, map, missing };
}

async function mmRefresh() {
  if (!mailMergeData || !mmStatusEl) return;
  const { names, map, missing } = await mmMatch();
  const matched = Object.keys(map);
  let msg = t("mmLoaded", { file: mailMergeData.name, rows: mailMergeData.rows.length, cols: mailMergeData.headers.length });
  if (!names.length) msg += " " + t("mmNoFields");
  else {
    msg += " " + t("mmMatched", { n: matched.length, total: names.length, list: matched.map((n) => `{{${n}}}`).join(", ") || "—" });
    if (missing.length) msg += " " + t("mmMissing", { list: missing.map((n) => `{{${n}}}`).join(", ") });
  }
  mmStatusEl.textContent = msg;
  if (mmOptionsEl) mmOptionsEl.hidden = false;
  if (mmNameColEl) {
    const prev = mmNameColEl.value;
    mmNameColEl.replaceChildren(new Option(t("mmNameNumber"), ""), ...mailMergeData.headers.map((h, i) => new Option(h || `#${i + 1}`, String(i))));
    // domyślnie kolumna z nazwiskiem / nazwą, jeśli jest
    const guess = mailMergeData.keys.findIndex((k) => /nazwisk|nazwa|name|firma/.test(k));
    mmNameColEl.value = prev && mmNameColEl.querySelector(`option[value="${prev}"]`) ? prev : (guess >= 0 ? String(guess) : "");
  }
  if (mmRowEl) { mmRowEl.max = String(mailMergeData.rows.length); if (!mmRowEl.value || +mmRowEl.value > mailMergeData.rows.length) mmRowEl.value = "1"; }
  if (mmRunBtn) mmRunBtn.disabled = !matched.length;
  if (mmPreviewBtn) mmPreviewBtn.disabled = !matched.length;
}

function mmValuesForRow(row, map) {
  const values = {};
  Object.entries(map).forEach(([name, i]) => { values[name] = row[i] ?? ""; });
  return values;
}

// Jeden wiersz → bajty .docx (treść przez zwykłą operację; nagłówki/stopki/przypisy osobno).
async function mmMergeOne(baseBytes, values) {
  const { bytes } = await buildPatchedDocx(baseBytes, [{ op: "placeholderFill", values, scope: "all", allowEmpty: true }]);
  const zip = await JSZip.loadAsync(bytes);
  const parts = Object.keys(zip.files).filter((n) => /^word\/(header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(n));
  for (const name of parts) {
    const xml = await zip.file(name).async("string");
    if (!xml.includes("{{")) continue;
    const res = applyPlaceholderFillInXml(xml, values, "all", { allowEmpty: true });
    if (res.count) zip.file(name, res.xml);
  }
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

function mmSafeName(s) {
  return String(s || "").replace(/[\\/:*?"<>|\u0000-\u001f]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 80);
}

async function mmRun() {
  if (!mailMergeData || !originalFileBytes) return;
  const { map, missing } = await mmMatch();
  if (!Object.keys(map).length) { toast(t("mmNoMatch"), "error"); return; }
  if (missing.length && !confirm(t("mmMissingConfirm", { list: missing.map((n) => `{{${n}}}`).join(", ") }))) return;
  const base = await buildDocumentForSave();
  const zip = new JSZip();
  const nameCol = mmNameColEl?.value === "" || mmNameColEl == null ? -1 : Number(mmNameColEl.value);
  const stem = mmSafeName(String(currentFileName || "dokument").replace(/\.docx?$/i, "")) || "dokument";
  const width = String(mailMergeData.rows.length).length;
  const used = new Set();
  if (mmRunBtn) mmRunBtn.disabled = true;
  try {
    for (let i = 0; i < mailMergeData.rows.length; i++) {
      const row = mailMergeData.rows[i];
      if (mmStatusEl) mmStatusEl.textContent = t("mmProgress", { i: i + 1, n: mailMergeData.rows.length });
      const bytes = await mmMergeOne(base, mmValuesForRow(row, map));
      const num = String(i + 1).padStart(Math.max(3, width), "0");
      let name = `${num} ${mmSafeName(nameCol >= 0 ? row[nameCol] : "") || stem}`.trim();
      while (used.has(name.toLowerCase())) name += "_";
      used.add(name.toLowerCase());
      zip.file(`${name}.docx`, bytes);
    }
    const out = await zip.generateAsync({ type: "uint8array", compression: "STORE" }); // .docx są już spakowane
    const blob = new Blob([out], { type: "application/zip" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${stem} — seryjna (${mailMergeData.rows.length}).zip`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    toast(t("mmDone", { n: mailMergeData.rows.length }), "success");
  } catch (e) {
    console.warn("mail merge", e);
    toast(t("mmFailed"), "error");
  } finally {
    if (mmRunBtn) mmRunBtn.disabled = false;
    await mmRefresh();
  }
}

// Podgląd: wartości wybranego wiersza do formularza pól (Wypełnij = zobaczysz w dokumencie, Cofnij wraca).
async function mmPreviewRow() {
  if (!mailMergeData) return;
  const idx = Math.min(Math.max(1, parseInt(mmRowEl?.value, 10) || 1), mailMergeData.rows.length) - 1;
  const { map } = await mmMatch();
  if (typeof runPlaceholderScan === "function") await runPlaceholderScan();
  const values = mmValuesForRow(mailMergeData.rows[idx], map);
  let filled = 0;
  document.querySelectorAll("#phForm input[data-ph-name]").forEach((input) => {
    if (input.dataset.phName in values) { input.value = values[input.dataset.phName]; filled++; }
  });
  toast(t("mmPreviewed", { row: idx + 1, n: filled }), "info");
}

function wireMailMerge() {
  mmPickBtn?.addEventListener("click", () => {
    if (!originalFileBytes) { toast(t("noFileToSave"), "error"); return; }
    mmFileEl?.click();
  });
  mmFileEl?.addEventListener("change", () => {
    const f = mmFileEl.files?.[0];
    mmFileEl.value = "";
    mmLoadFile(f);
  });
  mmRunBtn?.addEventListener("click", () => { mmRun(); });
  mmPreviewBtn?.addEventListener("click", () => { mmPreviewRow(); });
}

wireMailMerge();
