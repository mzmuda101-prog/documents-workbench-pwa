// Panel „Eksport” — TXT / HTML / Markdown z aktualnego dokumentu (razem z Twoimi edycjami) + druk / PDF.
// Wszystko lokalnie: budujemy z DOM podglądu, nic nie opuszcza urządzenia.

const exTxtBtn = document.getElementById("exTxtBtn");
const exMdBtn = document.getElementById("exMdBtn");
const exHtmlBtn = document.getElementById("exHtmlBtn");
const exPrintBtn = document.getElementById("exPrintBtn");
const exStatusEl = document.getElementById("exStatus");

function exportBaseName() {
  return (currentFileName || "dokument").replace(/\.docx$/i, "") || "dokument";
}

function exportListLevel(p) {
  const m = /docx-num-\d+-(\d+)/.exec(p.className || "");
  return m ? Number(m[1]) : -1;
}

function exportCellText(cell) {
  return (cell.innerText ?? cell.textContent ?? "").replace(/\s*\n+\s*/g, " ").replace(/\s+/g, " ").trim();
}

// Model neutralny: [{ kind: "heading"|"para"|"li"|"table", level, text, rows }]
function buildExportModel() {
  const { outline } = analyzeDocumentDom(docCanvasEl);
  const model = [];
  outline.forEach((item) => {
    if (item.type === "table") {
      const rows = [...item.el.querySelectorAll("tr")].map((tr) => [...tr.children].map(exportCellText));
      if (rows.length) model.push({ kind: "table", rows });
      return;
    }
    const text = (item.el.textContent || "").replace(/\s+/g, " ").trim();
    if (!text) return;
    if (item.type === "heading") { model.push({ kind: "heading", level: item.level || 2, text }); return; }
    const lvl = exportListLevel(item.el);
    model.push(lvl >= 0 ? { kind: "li", level: lvl, text } : { kind: "para", text });
  });
  return model;
}

function exportToTxt(model) {
  return model.map((b) => {
    if (b.kind === "table") return b.rows.map((r) => r.join("\t")).join("\n");
    if (b.kind === "li") return `${"  ".repeat(b.level)}• ${b.text}`;
    return b.text;
  }).join("\n\n").replace(/\n\n(?=(?: *• ))/g, "\n") + "\n";
}

function mdEscape(s) {
  return s.replace(/([\\`*_{}\[\]<>|])/g, "\\$1");
}

function exportToMarkdown(model) {
  const out = [];
  model.forEach((b, i) => {
    if (b.kind === "heading") out.push(`${"#".repeat(Math.min(6, b.level))} ${mdEscape(b.text)}`);
    else if (b.kind === "li") {
      const item = `${"  ".repeat(b.level)}- ${mdEscape(b.text)}`;
      const prev = model[i - 1];
      if (prev?.kind === "li" && out.length) out[out.length - 1] += `\n${item}`;
      else out.push(item);
    } else if (b.kind === "table") {
      const width = Math.max(...b.rows.map((r) => r.length));
      const line = (r) => `| ${Array.from({ length: width }, (_, c) => mdEscape(r[c] || "")).join(" | ")} |`;
      out.push([line(b.rows[0]), `|${" --- |".repeat(width)}`, ...b.rows.slice(1).map(line)].join("\n"));
    } else out.push(mdEscape(b.text));
  });
  return `${out.join("\n\n")}\n`;
}

function escHtmlEx(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function exportToHtmlBody(model) {
  const out = [];
  let inList = false;
  model.forEach((b) => {
    if (b.kind !== "li" && inList) { out.push("</ul>"); inList = false; }
    if (b.kind === "heading") out.push(`<h${Math.min(6, b.level)}>${escHtmlEx(b.text)}</h${Math.min(6, b.level)}>`);
    else if (b.kind === "li") {
      if (!inList) { out.push("<ul>"); inList = true; }
      out.push(`<li style="margin-left:${b.level * 1.2}em">${escHtmlEx(b.text)}</li>`);
    } else if (b.kind === "table") {
      out.push("<table>", ...b.rows.map((r, i) => `<tr>${r.map((c) => `<${i ? "td" : "th"}>${escHtmlEx(c)}</${i ? "td" : "th"}>`).join("")}</tr>`), "</table>");
    } else out.push(`<p>${escHtmlEx(b.text)}</p>`);
  });
  if (inList) out.push("</ul>");
  return out.join("\n");
}

function exportToHtml(model, title) {
  return `<!doctype html>
<html lang="${currentLang}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${escHtmlEx(title)}</title>
<style>body{font:16px/1.6 system-ui,sans-serif;max-width:760px;margin:2rem auto;padding:0 1rem;color:#111}
table{border-collapse:collapse;margin:1rem 0}td,th{border:1px solid #bbb;padding:.35rem .6rem;vertical-align:top}th{background:#f2f2f2}</style>
</head><body>
${exportToHtmlBody(model)}
</body></html>
`;
}

function downloadTextFile(text, name, mime) {
  const url = URL.createObjectURL(new Blob([text], { type: `${mime};charset=utf-8` }));
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function runExport(kind) {
  if (!originalFileBytes) { toast(t("noFileToSave"), "error"); return; }
  const model = buildExportModel();
  if (!model.length) { toast(t("exportEmpty"), "info"); return; }
  const base = exportBaseName();
  if (kind === "txt") downloadTextFile(exportToTxt(model), `${base}.txt`, "text/plain");
  else if (kind === "md") downloadTextFile(exportToMarkdown(model), `${base}.md`, "text/markdown");
  else downloadTextFile(exportToHtml(model, base), `${base}.html`, "text/html");
  const msg = t("exportDone", { name: `${base}.${kind}` });
  if (exStatusEl) exStatusEl.textContent = msg;
  toast(msg, "success");
}

// Druk / PDF: wierny układ stron z podglądu w ukrytej ramce (sama strona, bez interfejsu).
// Na telefonie w trybie „Dopasuj” podgląd jest przeformatowany — wtedy drukujemy tekst z nagłówkami.
function printDocument() {
  if (!originalFileBytes) { toast(t("noFileToSave"), "error"); return; }
  const host = docCanvasEl?.querySelector(".docx-preview-host");
  if (!host) return;
  const reflow = typeof shouldUseMobileReflow === "function" && shouldUseMobileReflow();
  const base = exportBaseName();
  const html = reflow
    ? exportToHtml(buildExportModel(), base)
    : `<!doctype html><html lang="${currentLang}"><head><meta charset="utf-8"><title>${escHtmlEx(base)}</title>
<style>@page{margin:0}html,body{margin:0;background:#fff}
.docx-wrapper{background:none!important;padding:0!important}
.docx-wrapper>section.docx{box-shadow:none!important;margin:0 auto!important;break-after:page}
.search-hit,.search-hit-active{background:none!important;outline:none!important}</style></head>
<body>${host.outerHTML}</body></html>`;
  const frame = document.createElement("iframe");
  frame.setAttribute("aria-hidden", "true");
  frame.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden";
  frame.onload = () => {
    setTimeout(() => {
      try { frame.contentWindow.focus(); frame.contentWindow.print(); } catch (_) { toast(t("exportPrintFail"), "error"); }
      setTimeout(() => frame.remove(), 60000);
    }, 250);
  };
  document.body.appendChild(frame);
  frame.srcdoc = html;
}

exTxtBtn?.addEventListener("click", () => runExport("txt"));
exMdBtn?.addEventListener("click", () => runExport("md"));
exHtmlBtn?.addEventListener("click", () => runExport("html"));
exPrintBtn?.addEventListener("click", printDocument);
