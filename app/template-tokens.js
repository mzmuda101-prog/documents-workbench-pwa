// Template tokens — {{placeholder}} vs !snippet (separate syntax, composable workflow).

const PLACEHOLDER_TOKEN_RE = /\{\{\s*([a-zA-Z0-9_.-]+)\s*\}\}/g;
// Nazwy snippetów z polskimi literami (!podpisŁK). Trigger tylko na początku słowa —
// „Uwaga!Ważne” to zwykły tekst, nie !Ważne. (\b w JS nie zna polskich liter.)
const SNIPPET_NAME_RE = /^[\p{L}][\p{L}\p{N}_-]*$/u;
const SNIPPET_TRIGGER_RE = /(?<![\p{L}\p{N}_!])!([\p{L}][\p{L}\p{N}_-]*)(?![\p{L}\p{N}_-])/gu;

function formatSnippetTrigger(name) {
  return `!${String(name || "").trim()}`;
}

function normalizeSnippetName(raw) {
  const s = String(raw || "").trim().replace(/^!+/, "");
  return SNIPPET_NAME_RE.test(s) ? s : "";
}

function normalizePlaceholderName(raw) {
  const s = String(raw || "").trim().replace(/^\{\{\s*|\s*\}\}$/g, "").trim();
  return /^[a-zA-Z0-9_.-]+$/.test(s) ? s : "";
}

function formatPlaceholderToken(name) {
  return `{{${String(name || "").trim()}}}`;
}

function scanPlaceholdersInText(text) {
  const hits = [];
  if (!text) return hits;
  const re = new RegExp(PLACEHOLDER_TOKEN_RE.source, "g");
  let m;
  while ((m = re.exec(text)) !== null) {
    hits.push({ token: m[0], name: m[1], start: m.index, end: m.index + m[0].length, kind: "placeholder" });
  }
  return hits;
}

function scanSnippetTriggersInText(text) {
  const hits = [];
  if (!text) return hits;
  const re = new RegExp(SNIPPET_TRIGGER_RE.source, "gu");
  let m;
  while ((m = re.exec(text)) !== null) {
    hits.push({ token: m[0], name: m[1], start: m.index, end: m.index + m[0].length, kind: "snippet" });
  }
  return hits;
}
