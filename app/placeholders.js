// Placeholder fields — scan {{name}} tokens and fill via XML replace.

const PLACEHOLDER_TOKEN_RE = /\{\{\s*([a-zA-Z0-9_.-]+)\s*\}\}/g;

function scanPlaceholdersInText(text) {
  const hits = [];
  if (!text) return hits;
  const re = new RegExp(PLACEHOLDER_TOKEN_RE.source, "g");
  let m;
  while ((m = re.exec(text)) !== null) {
    hits.push({ token: m[0], name: m[1], start: m.index, end: m.index + m[0].length });
  }
  return hits;
}

async function scanPlaceholders(bytes) {
  const texts = await extractParagraphTextsFromDocx(bytes);
  const fields = new Map();
  let total = 0;
  texts.forEach((text, paraIndex) => {
    scanPlaceholdersInText(text).forEach((hit) => {
      total++;
      const entry = fields.get(hit.name) || { name: hit.name, count: 0, samples: [] };
      entry.count++;
      if (entry.samples.length < 4) entry.samples.push({ paraIndex, token: hit.token });
      fields.set(hit.name, entry);
    });
  });
  return {
    fields: [...fields.values()].sort((a, b) => a.name.localeCompare(b.name, "pl")),
    total,
  };
}

function placeholderFieldLabel(name) {
  return name.replace(/_/g, " ").replace(/-/g, " ");
}
