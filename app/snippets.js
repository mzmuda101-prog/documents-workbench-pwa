// Snippets — localStorage blocks triggered by !name; composable with {{placeholders}}.

const SNIPPETS_STORAGE_KEY = "documents-workbench-snippets";

function loadSnippets() {
  try {
    const raw = localStorage.getItem(SNIPPETS_STORAGE_KEY);
    const list = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(list)) return [];
    return list
      .map((item) => ({
        name: normalizeSnippetName(item?.name),
        body: String(item?.body ?? ""),
        updatedAt: item?.updatedAt || 0,
      }))
      .filter((item) => item.name && item.body);
  } catch (_) {
    return [];
  }
}

function persistSnippets(list) {
  localStorage.setItem(SNIPPETS_STORAGE_KEY, JSON.stringify(list));
}

function upsertSnippet(name, body) {
  const cleanName = normalizeSnippetName(name);
  if (!cleanName) return null;
  const text = String(body ?? "");
  if (!text.trim()) return null;
  const list = loadSnippets().filter((s) => s.name !== cleanName);
  const entry = { name: cleanName, body: text, updatedAt: Date.now() };
  list.push(entry);
  list.sort((a, b) => a.name.localeCompare(b.name, "pl"));
  persistSnippets(list);
  return entry;
}

function deleteSnippet(name) {
  const cleanName = normalizeSnippetName(name);
  const list = loadSnippets().filter((s) => s.name !== cleanName);
  persistSnippets(list);
}

function snippetsToMap(list) {
  const map = {};
  (list || []).forEach((s) => { if (s.name) map[s.name] = s.body; });
  return map;
}

async function scanSnippetTriggers(bytes) {
  const texts = await extractParagraphTextsFromDocx(bytes);
  const stored = snippetsToMap(loadSnippets());
  const triggers = new Map();
  let total = 0;
  texts.forEach((text, paraIndex) => {
    scanSnippetTriggersInText(text).forEach((hit) => {
      total++;
      const entry = triggers.get(hit.name) || {
        name: hit.name,
        count: 0,
        samples: [],
        hasDefinition: !!stored[hit.name],
      };
      entry.count++;
      if (entry.samples.length < 4) entry.samples.push({ paraIndex, token: hit.token });
      entry.hasDefinition = !!stored[hit.name];
      triggers.set(hit.name, entry);
    });
  });
  return {
    triggers: [...triggers.values()].sort((a, b) => a.name.localeCompare(b.name, "pl")),
    total,
    stored,
  };
}

function buildSnippetExpandMap(triggers, storedMap) {
  const map = {};
  (triggers || []).forEach((t) => {
    if (storedMap?.[t.name]) map[t.name] = storedMap[t.name];
  });
  return map;
}
