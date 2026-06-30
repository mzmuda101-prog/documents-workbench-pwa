// Placeholder fields — scan {{name}} tokens and fill via XML replace.

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
