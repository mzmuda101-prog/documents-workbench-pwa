// Template tokens — {{placeholder}} vs !snippet (separate syntax, composable workflow).

const PLACEHOLDER_TOKEN_RE = /\{\{\s*([a-zA-Z0-9_.-]+)\s*\}\}/g;
const SNIPPET_NAME_RE = /^[a-zA-Z][a-zA-Z0-9_-]*$/;
const SNIPPET_TRIGGER_RE = /!([a-zA-Z][a-zA-Z0-9_-]*)\b/g;

function formatSnippetTrigger(name) {
  return `!${String(name || "").trim()}`;
}

function normalizeSnippetName(raw) {
  const s = String(raw || "").trim().replace(/^!+/, "");
  return SNIPPET_NAME_RE.test(s) ? s : "";
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
  const re = new RegExp(SNIPPET_TRIGGER_RE.source, "g");
  let m;
  while ((m = re.exec(text)) !== null) {
    hits.push({ token: m[0], name: m[1], start: m.index, end: m.index + m[0].length, kind: "snippet" });
  }
  return hits;
}
