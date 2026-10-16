// local-ai.js — wspólny asystent AI NA URZĄDZENIU dla Sheet i Documents Workbench.
// KOPIA 1:1 w sheet-workbench-pwa/app i documents-workbench-pwa/app — zmieniaj oba pliki.
//
// Zasady (pakiet AI, 2026-10-16):
//  • Plik nigdy nie wychodzi z urządzenia. Model działa w przeglądarce: albo wbudowany w Chrome
//    (Prompt API, bez pobierania), albo pobrany raz model WebGPU (local-ai-worker.js).
//  • Nic nie pobiera się samo — pierwsze użycie otwiera okno z rozmiarem i zgodą.
//  • Model tylko PROPONUJE. Liczy i zmienia zawsze kod aplikacji; wynik modelu jest sprawdzany
//    (JSON wg schematu, liczby z danych, poprawna formuła), a zmiana idzie przez Cofnij.
//  • Testy podstawiają atrapę: window.__localAiMock = (messages, opts) => "odpowiedź".
//
// API: LocalAI.ensure({ reason }) → Promise<boolean>, LocalAI.generate({ system, user, maxTokens,
//      onPartial, signal }), LocalAI.json({ system, user, schema, validate }), LocalAI.isReady(),
//      LocalAI.openSettings(), LocalAI.onChange(fn), LocalAI.release().

const LocalAI = (() => {
  const PREF_KEY = "localAi.prefs.v1";
  const CACHE_KEY = "local-ai-models-v1";
  // Worker z tą samą wersją co ten plik (?v= z wydania) — trafia w cache service workera.
  const OWN_VERSION = (document.currentScript?.src?.match(/[?&]v=([^&]+)/) || [])[1] || "1";
  const WORKER_URL = `app/local-ai-worker.js?v=${OWN_VERSION}`;

  // Wyniki Etapu 0 (Mac, 2026-10-16): Qwen3.5 2B — 4/4 testy po polsku, 0.8B — 2/4.
  const MODELS = {
    "qwen3.5-2b": {
      id: "onnx-community/Qwen3.5-2B-ONNX-OPT", kind: "vlm", mb: { q4f16: 1580, q4: 1750 },
      files: { q4f16: ["decoder_model_merged_q4f16.onnx", "decoder_model_merged_q4f16.onnx_data", "embed_tokens_q4f16.onnx_data"], q4: ["decoder_model_merged_q4.onnx", "decoder_model_merged_q4.onnx_data", "embed_tokens_q4.onnx_data"] },
      label: { pl: "Qwen3.5 2B — dokładny", en: "Qwen3.5 2B — accurate" },
      note: { pl: "Najlepszy po polsku. Polecany do „Przepisz” i pisania formuł. Duży — raczej na komputer.", en: "Best Polish quality. Recommended for Rewrite and writing formulas. Large — rather for a computer." },
    },
    "qwen3.5-0.8b": {
      id: "onnx-community/Qwen3.5-0.8B-Text-ONNX", kind: "causal", mb: { q4f16: 470, q4: 550 },
      files: { q4f16: ["model_q4f16.onnx", "model_q4f16.onnx_data"], q4: ["model_q4.onnx", "model_q4.onnx_data"] },
      label: { pl: "Qwen3.5 0.8B — lekki", en: "Qwen3.5 0.8B — light" },
      note: { pl: "Wystarczy do „Zapytaj dokument” i krótkich podsumowań. Dłuższe teksty pisze słabszą polszczyzną. Dobry na telefon.", en: "Good enough for Ask the document and short summaries. Weaker Polish in longer texts. Good for a phone." },
    },
  };

  const TXT = {
    pl: {
      title: "Asystent AI na urządzeniu",
      intro: "Działa w tej przeglądarce. Treść plików nie jest nigdzie wysyłana — z internetu pobiera się tylko sam model, raz. Wiele funkcji działa też bez niego (słownik, wzory, reguły); asystent jest potrzebny do opcji oznaczonych ✦.",
      chrome: "Wbudowany w Chrome — bez pobierania",
      chromeNote: "Model Google z przeglądarki Chrome (Prompt API).",
      size: "ok. {mb} MB do pobrania",
      cached: "pobrany na tym urządzeniu",
      noGpu: "Ta przeglądarka nie ma WebGPU z kartą graficzną — modelu nie da się tu uruchomić. Działa w aktualnym Chrome, Edge i Safari (iOS 26+).",
      start: "Pobierz i uruchom",
      startCached: "Uruchom",
      release: "Zwolnij pamięć",
      remove: "Usuń pobrany model",
      close: "Zamknij",
      off: "Wyłączony",
      loading: "Pobieranie… {done} / {total} MB",
      compile: "Przygotowanie modelu…",
      warmup: "Rozgrzewanie…",
      ready: "Gotowy: {name}",
      error: "Błąd: {msg}",
      removed: "Model usunięty z urządzenia.",
      reasonDefault: "Ta funkcja potrzebuje asystenta AI.",
      phoneWarn: "Na telefonie wybierz model, który zmierzysz na stronie „Pomiar modeli AI”; duży model może zamknąć kartę przy braku pamięci.",
    },
    en: {
      title: "On-device AI assistant",
      intro: "Runs in this browser. File contents are never sent anywhere — only the model itself is downloaded, once. Many features work without it (dictionary, templates, rules); the assistant is needed for options marked ✦.",
      chrome: "Built into Chrome — no download",
      chromeNote: "Google's model shipped with Chrome (Prompt API).",
      size: "about {mb} MB to download",
      cached: "downloaded on this device",
      noGpu: "This browser has no WebGPU adapter — the model cannot run here. Use current Chrome, Edge or Safari (iOS 26+).",
      start: "Download and start",
      startCached: "Start",
      release: "Free memory",
      remove: "Delete downloaded model",
      close: "Close",
      off: "Off",
      loading: "Downloading… {done} / {total} MB",
      compile: "Preparing model…",
      warmup: "Warming up…",
      ready: "Ready: {name}",
      error: "Error: {msg}",
      removed: "Model removed from this device.",
      reasonDefault: "This feature needs the AI assistant.",
      phoneWarn: "On a phone pick a model you measured on the “AI model benchmark” page; a large model may close the tab when memory runs out.",
    },
  };

  const lang = () => (typeof currentLang === "string" && currentLang === "en" ? "en" : "pl");
  const tx = (key, vars = {}) => String(TXT[lang()][key] ?? TXT.pl[key] ?? key).replace(/\{(\w+)\}/g, (_, k) => vars[k] ?? "");

  const state = {
    backend: null, // "chrome" | "webgpu" | "mock"
    status: "off", // off | loading | ready | error
    phase: "",
    progress: null,
    error: "",
    gpu: null, // { ok, f16 }
    chrome: null, // "available" | "downloadable" | "downloading" | "unavailable" | null
    cached: {},
  };
  let prefs = { choice: "" };
  try { prefs = { ...prefs, ...JSON.parse(localStorage.getItem(PREF_KEY) || "{}") }; } catch {}
  const savePrefs = () => { try { localStorage.setItem(PREF_KEY, JSON.stringify(prefs)); } catch {} };

  let worker = null;
  let chromeSession = null;
  let reqSeq = 0;
  const pending = new Map();
  let loadWaiters = [];
  const listeners = new Set();
  const emit = () => listeners.forEach((fn) => { try { fn(snapshot()); } catch {} });
  const snapshot = () => ({ ...state, choice: currentChoice() });

  const mock = () => (typeof window !== "undefined" && typeof window.__localAiMock === "function" ? window.__localAiMock : null);

  // ── możliwości przeglądarki ──────────────────────────────────────────────
  async function probe() {
    if (!state.gpu) {
      state.gpu = { ok: false, f16: false };
      try {
        const adapter = navigator.gpu ? await navigator.gpu.requestAdapter() : null;
        if (adapter) state.gpu = { ok: true, f16: adapter.features.has("shader-f16") };
      } catch {}
    }
    if (state.chrome == null) {
      state.chrome = "unavailable";
      try {
        if (typeof LanguageModel !== "undefined" && LanguageModel.availability) {
          state.chrome = await LanguageModel.availability({
            expectedInputs: [{ type: "text", languages: ["pl", "en"] }],
            expectedOutputs: [{ type: "text", languages: ["pl"] }],
          });
        }
      } catch {}
    }
    return state;
  }

  const dtype = () => (state.gpu?.f16 ? "q4f16" : "q4");
  function currentChoice() {
    if (prefs.choice === "chrome" && state.chrome && state.chrome !== "unavailable") return "chrome";
    if (MODELS[prefs.choice]) return prefs.choice;
    if (state.chrome === "available") return "chrome";
    // telefon / mały tablet: domyślnie lekki model (470 MB zamiast 1,6 GB) — pomiar poleceń 2026-10-17
    const phone = typeof matchMedia === "function" && matchMedia("(pointer: coarse)").matches && Math.min(screen.width, screen.height) < 820;
    return phone ? "qwen3.5-0.8b" : "qwen3.5-2b";
  }

  // ── WebGPU worker ────────────────────────────────────────────────────────
  function ensureWorker() {
    if (worker) return worker;
    worker = new Worker(WORKER_URL, { type: "module", name: "local-ai" });
    worker.addEventListener("message", ({ data }) => onWorker(data));
    worker.addEventListener("error", (e) => fail(e.message || "worker"));
    return worker;
  }

  function onWorker(data) {
    if (data.type === "progress") {
      state.progress = { loaded: data.loaded, total: data.total };
      state.phase = "download";
      emit();
    } else if (data.type === "phase") {
      state.phase = data.phase;
      emit();
    } else if (data.type === "ready") {
      state.status = "ready";
      state.backend = "webgpu";
      state.phase = "";
      state.progress = null;
      state.cached[currentChoice()] = true;
      loadWaiters.forEach((w) => w.resolve(true));
      loadWaiters = [];
      emit();
    } else if (data.type === "cache-status") {
      state.cached[data.model] = !!data.cached;
      emit();
    } else if (data.type === "partial") {
      pending.get(data.id)?.onPartial?.(data.text);
    } else if (data.type === "done") {
      pending.get(data.id)?.resolve(data.text);
      pending.delete(data.id);
    } else if (data.type === "released") {
      state.status = "off";
      state.backend = null;
      emit();
    } else if (data.type === "deleted") {
      state.status = "off";
      state.backend = null;
      state.cached = {};
      worker?.terminate();
      worker = null;
      emit();
    } else if (data.type === "error") {
      if (data.id != null && pending.has(data.id)) {
        pending.get(data.id).reject(new Error(data.message));
        pending.delete(data.id);
      } else fail(data.message);
    }
  }

  function fail(message) {
    state.status = "error";
    state.error = String(message || "");
    state.phase = "";
    loadWaiters.forEach((w) => w.resolve(false));
    loadWaiters = [];
    pending.forEach((p) => p.reject(new Error(state.error)));
    pending.clear();
    emit();
  }

  function refreshCache() {
    if (!state.gpu?.ok) return;
    for (const [key, m] of Object.entries(MODELS)) {
      ensureWorker().postMessage({ type: "cache-status", model: key, cacheKey: CACHE_KEY, id: m.id, files: m.files[dtype()] });
    }
  }

  async function start(choice = currentChoice()) {
    await probe();
    if (state.status === "ready" && state.backend && (state.backend !== "webgpu" || state.loadedChoice === choice)) return true;
    state.error = "";
    if (choice === "chrome") {
      state.status = "loading";
      state.phase = "compile";
      emit();
      try {
        chromeSession = await LanguageModel.create({
          expectedInputs: [{ type: "text", languages: ["pl", "en"] }],
          expectedOutputs: [{ type: "text", languages: ["pl"] }],
          monitor(m) {
            m.addEventListener("downloadprogress", (e) => {
              state.phase = "download";
              state.progress = { loaded: e.loaded * 1000e6, total: 1000e6 };
              emit();
            });
          },
        });
        state.status = "ready";
        state.backend = "chrome";
        state.phase = "";
        emit();
        return true;
      } catch (err) {
        fail(err?.message || err);
        return false;
      }
    }
    if (!state.gpu?.ok) {
      fail(tx("noGpu"));
      return false;
    }
    const m = MODELS[choice];
    state.status = "loading";
    state.phase = "download";
    state.loadedChoice = choice;
    emit();
    const done = new Promise((resolve) => loadWaiters.push({ resolve }));
    ensureWorker().postMessage({ type: "load", spec: { id: m.id, kind: m.kind }, dtype: dtype(), cacheKey: CACHE_KEY });
    return done;
  }

  // ── generowanie ──────────────────────────────────────────────────────────
  function isReady() {
    return !!mock() || state.status === "ready";
  }

  // shots: [[pytanie, odpowiedź], …] — przykłady w poleceniu (małe modele uczą się z nich formatu);
  // prefix: wymuszony początek odpowiedzi („=”, „{"szukaj": ["”) — model nie zaczyna od „Oto…”.
  // Pomiar: documents-workbench-pwa/docs/llm-bench/prompts.html.
  async function generate({ system = "", shots = [], user = "", prefix = "", maxTokens = 256, onPartial = null, signal = null } = {}) {
    const messages = [];
    if (system) messages.push({ role: "system", content: system });
    for (const [u, a] of shots) messages.push({ role: "user", content: u }, { role: "assistant", content: a });
    messages.push({ role: "user", content: user });
    const m = mock();
    if (m) {
      const text = String(await m(messages, { maxTokens }));
      onPartial?.(text);
      return text;
    }
    if (state.status !== "ready") throw new Error("not-ready");
    if (state.backend === "chrome") {
      // Każde pytanie w świeżej sesji — bez pamięci poprzednich rozmów; rola „system” wolno
      // podać tylko przy tworzeniu sesji (initialPrompts).
      const s = await LanguageModel.create({
        initialPrompts: messages.slice(0, -1),
        expectedInputs: [{ type: "text", languages: ["pl", "en"] }],
        expectedOutputs: [{ type: "text", languages: ["pl"] }],
      });
      try {
        let out = "";
        const stream = s.promptStreaming(user, { signal });
        for await (const chunk of stream) {
          out += chunk;
          onPartial?.(out);
        }
        return out.trim();
      } finally {
        s.destroy();
      }
    }
    const id = ++reqSeq;
    return new Promise((resolve, reject) => {
      pending.set(id, { resolve, reject, onPartial });
      signal?.addEventListener("abort", () => worker?.postMessage({ type: "abort", id }), { once: true });
      worker.postMessage({ type: "generate", id, messages, maxTokens, prefix });
    });
  }

  // Pierwszy obiekt JSON z odpowiedzi (model lubi dodać ```json … ```).
  function firstJson(text) {
    const s = String(text || "");
    const start = s.indexOf("{");
    if (start < 0) return null;
    let depth = 0;
    let inStr = false;
    for (let i = start; i < s.length; i++) {
      const ch = s[i];
      if (inStr) {
        if (ch === "\\") i++;
        else if (ch === '"') inStr = false;
        continue;
      }
      if (ch === '"') inStr = true;
      else if (ch === "{") depth++;
      else if (ch === "}" && --depth === 0) {
        try { return JSON.parse(s.slice(start, i + 1)); } catch { return null; }
      }
    }
    return null;
  }

  // JSON z walidacją; przy złej odpowiedzi jedna poprawka z informacją, co było źle.
  async function json({ system, shots = [], prefix = "", user, maxTokens = 200, validate = null, signal = null } = {}) {
    let answer = await generate({ system, shots, prefix, user, maxTokens, signal });
    let parsed = firstJson(answer);
    let problem = !parsed ? "invalid-json" : validate ? validate(parsed) : "";
    if (!problem) return { value: parsed, raw: answer };
    const retryUser = `${user}\n\nPoprzednia odpowiedź była niepoprawna (${problem}). Odpowiedz ponownie WYŁĄCZNIE poprawnym obiektem JSON.`;
    answer = await generate({ system, shots, prefix, user: retryUser, maxTokens, signal });
    parsed = firstJson(answer);
    problem = !parsed ? "invalid-json" : validate ? validate(parsed) : "";
    return problem ? { value: null, raw: answer, problem } : { value: parsed, raw: answer };
  }

  function release() {
    if (chromeSession) { try { chromeSession.destroy(); } catch {} chromeSession = null; }
    if (state.backend === "webgpu") worker?.postMessage({ type: "release" });
    state.status = "off";
    state.backend = null;
    emit();
  }

  function removeModel() {
    ensureWorker().postMessage({ type: "delete-cache", cacheKey: CACHE_KEY });
  }

  // ── okno ustawień / zgody ────────────────────────────────────────────────
  let dialog = null;
  let dialogReason = "";
  let dialogWaiter = null;

  function el(tag, cls, text) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }

  function buildDialog() {
    dialog = el("dialog", "lai-dialog");
    dialog.setAttribute("aria-labelledby", "laiTitle");
    dialog.addEventListener("close", () => {
      dialogWaiter?.(isReady());
      dialogWaiter = null;
    });
    dialog.addEventListener("click", (e) => { if (e.target === dialog) dialog.close(); });
    document.body.appendChild(dialog);
    listeners.add(() => { if (dialog.open) renderDialog(); });
  }

  function renderDialog() {
    const L = lang();
    dialog.replaceChildren();
    const box = el("div", "lai-box");
    const title = el("h2", "lai-title", tx("title"));
    title.id = "laiTitle";
    box.append(title);
    if (dialogReason) box.append(el("p", "lai-reason", dialogReason));
    box.append(el("p", "lai-intro", tx("intro")));
    const list = el("div", "lai-options");
    list.setAttribute("role", "radiogroup");
    const choice = currentChoice();
    const busy = state.status === "loading";
    const addOption = (key, label, meta, note, disabled) => {
      const opt = el("label", "lai-option");
      const input = document.createElement("input");
      input.type = "radio";
      input.name = "laiChoice";
      input.value = key;
      input.checked = key === choice;
      input.disabled = disabled || busy;
      input.addEventListener("change", () => {
        prefs.choice = key;
        savePrefs();
        if (state.status === "ready" && (state.backend === "chrome" ? key !== "chrome" : state.loadedChoice !== key)) release();
        renderDialog();
      });
      const body = el("span", "lai-option-body");
      body.append(el("strong", "", label), el("span", "lai-meta", meta));
      if (note) body.append(el("span", "lai-note", note));
      opt.append(input, body);
      list.append(opt);
    };
    if (state.chrome && state.chrome !== "unavailable") addOption("chrome", tx("chrome"), "", tx("chromeNote"), false);
    for (const [key, m] of Object.entries(MODELS)) {
      const meta = state.cached[key] ? tx("cached") : tx("size", { mb: m.mb[dtype()] });
      addOption(key, m.label[L], meta, m.note[L], !state.gpu?.ok);
    }
    box.append(list);
    if (!state.gpu?.ok && !(state.chrome && state.chrome !== "unavailable")) box.append(el("p", "lai-warn", tx("noGpu")));
    else if (window.matchMedia?.("(pointer: coarse)").matches) box.append(el("p", "lai-note", tx("phoneWarn")));

    const status = el("div", "lai-status");
    status.setAttribute("role", "status");
    if (state.status === "loading") {
      if (state.phase === "download" && state.progress?.total) {
        status.textContent = tx("loading", { done: Math.round(state.progress.loaded / 1e6), total: Math.round(state.progress.total / 1e6) });
        const bar = el("div", "lai-bar");
        const fill = el("i");
        fill.style.width = `${Math.min(100, (state.progress.loaded / state.progress.total) * 100).toFixed(1)}%`;
        bar.append(fill);
        status.append(bar);
      } else status.textContent = tx(state.phase === "warmup" ? "warmup" : "compile");
    } else if (state.status === "ready") {
      const name = state.backend === "chrome" ? tx("chrome") : MODELS[state.loadedChoice]?.label[L] || "";
      status.textContent = tx("ready", { name });
      status.classList.add("is-ready");
    } else if (state.status === "error") {
      status.textContent = tx("error", { msg: state.error });
      status.classList.add("is-error");
    } else status.textContent = tx("off");
    box.append(status);

    const row = el("div", "lai-actions");
    const startBtn = el("button", "lai-btn is-primary", choice !== "chrome" && !state.cached[choice] ? tx("start") : tx("startCached"));
    startBtn.type = "button";
    const canStart = choice === "chrome" || state.gpu?.ok;
    startBtn.disabled = busy || !canStart || (state.status === "ready" && (state.backend === "chrome" ? choice === "chrome" : state.loadedChoice === choice));
    startBtn.addEventListener("click", async () => {
      const ok = await start(choice);
      if (ok && dialogWaiter) dialog.close();
    });
    const releaseBtn = el("button", "lai-btn", tx("release"));
    releaseBtn.type = "button";
    releaseBtn.disabled = state.status !== "ready";
    releaseBtn.addEventListener("click", release);
    const removeBtn = el("button", "lai-btn", tx("remove"));
    removeBtn.type = "button";
    removeBtn.disabled = busy || !Object.values(state.cached).some(Boolean);
    removeBtn.addEventListener("click", () => {
      release();
      removeModel();
    });
    const closeBtn = el("button", "lai-btn", tx("close"));
    closeBtn.type = "button";
    closeBtn.addEventListener("click", () => dialog.close());
    row.append(startBtn, releaseBtn, removeBtn, closeBtn);
    box.append(row);
    dialog.append(box);
  }

  async function openSettings(reason = "") {
    await probe();
    if (!dialog) buildDialog();
    dialogReason = reason;
    refreshCache();
    renderDialog();
    if (!dialog.open) dialog.showModal();
  }

  // Funkcja potrzebuje modelu: gotowy → od razu; inaczej okno ze zgodą (nic nie pobiera się samo).
  async function ensure({ reason = "" } = {}) {
    if (isReady()) return true;
    await openSettings(reason || tx("reasonDefault"));
    return new Promise((resolve) => { dialogWaiter = resolve; });
  }

  // Nazwa działającego modelu do komunikatów („Odpowiedź pisze Asystent AI (Qwen3.5 0.8B — lekki)”)
  function readyName() {
    if (mock()) return "test";
    if (state.status !== "ready") return "";
    return state.backend === "chrome" ? tx("chrome") : MODELS[state.loadedChoice]?.label[lang()] || "";
  }

  function onChange(fn) {
    listeners.add(fn);
    return () => listeners.delete(fn);
  }

  // Każdy element z data-local-ai-settings (pozycja menu ⋯) otwiera okno asystenta.
  document.addEventListener("click", (e) => {
    const trigger = e.target.closest?.("[data-local-ai-settings]");
    if (!trigger) return;
    e.preventDefault();
    openSettings();
  });

  return { ensure, generate, json, firstJson, isReady, readyName, openSettings, onChange, release, probe, status: snapshot, MODELS };
})();

window.LocalAI = LocalAI;
