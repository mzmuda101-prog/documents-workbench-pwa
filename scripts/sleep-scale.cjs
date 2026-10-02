// sleep-scale.cjs — skraca sztywne pauzy testów (page.waitForTimeout) o czynnik SLEEP_SCALE.
// Ładowany przez runner (NODE_OPTIONS=--require) przed każdym testem; same testy bez zmian.
// Pierwszy przebieg idzie z 0,5 (szybciej), porażka jest powtarzana POJEDYNCZO z 1 — więc test,
// który potrzebuje pełnej pauzy, nie zgłasza fałszywego błędu, tylko „niestabilny” (wypisany).
// Wzorzec z Sheet Workbench (2026-09-28).
const scale = parseFloat(process.env.SLEEP_SCALE || "1");
if (Number.isFinite(scale) && scale > 0 && scale !== 1) {
  let pw;
  try { pw = require(require.resolve("playwright", { paths: [process.cwd()] })); } catch { pw = null; }
  const patchPage = (p) => {
    if (!p || p.__dwbScaled) return p;
    p.__dwbScaled = true;
    const orig = p.waitForTimeout.bind(p);
    p.waitForTimeout = (ms) => orig(Math.max(0, Math.round(ms * scale)));
    return p;
  };
  const patchContext = (c) => {
    if (!c || c.__dwbScaled) return c;
    c.__dwbScaled = true;
    const np = c.newPage.bind(c);
    c.newPage = async (...a) => patchPage(await np(...a));
    return c;
  };
  const patchBrowser = (b) => {
    const nc = b.newContext.bind(b);
    b.newContext = async (...a) => patchContext(await nc(...a));
    const np = b.newPage.bind(b);
    b.newPage = async (...a) => patchPage(await np(...a));
    return b;
  };
  for (const type of pw ? [pw.chromium, pw.webkit, pw.firefox] : []) {
    if (!type) continue;
    const launch = type.launch.bind(type);
    type.launch = async (...a) => patchBrowser(await launch(...a));
    if (type.launchPersistentContext) {
      const lpc = type.launchPersistentContext.bind(type);
      type.launchPersistentContext = async (...a) => patchContext(await lpc(...a));
    }
  }
}
