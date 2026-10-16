// local-ai-worker.js — model językowy na urządzeniu (WebGPU), osobny wątek dla local-ai.js.
// KOPIA 1:1 w sheet-workbench-pwa/app i documents-workbench-pwa/app — zmieniaj oba pliki.
//
// Tworzony dopiero, gdy użytkownik uruchomi asystenta. Tekst pliku nie opuszcza urządzenia;
// z sieci przychodzą wyłącznie wagi modelu (Hugging Face) i biblioteka (esm.sh).
// Pomiar modeli: documents-workbench-pwa/docs/llm-bench (Etap 0, 2026-10-16).

const TRANSFORMERS_URL = "https://esm.sh/@huggingface/transformers@4.3.1";

let lib = null;
let loaded = null; // { key, model, tokenizer }
let stopper = null;

const post = (type, payload = {}) => self.postMessage({ type, ...payload });

async function loadLib(cacheKey) {
  lib ||= await import(TRANSFORMERS_URL);
  lib.env.useBrowserCache = true;
  lib.env.cacheKey = cacheKey;
  lib.env.allowLocalModels = false;
  return lib;
}

async function cacheStatus({ cacheKey, id, files }) {
  try {
    if (!self.caches) return false;
    const urls = (await (await caches.open(cacheKey)).keys()).map((r) => r.url);
    return files.every((f) => urls.some((u) => u.includes(id) && u.endsWith("/" + f)));
  } catch {
    return false;
  }
}

async function load({ spec, dtype, cacheKey }) {
  const key = `${spec.id}|${dtype}`;
  if (loaded?.key === key) return post("ready", { key });
  if (loaded) await release();
  const L = await loadLib(cacheKey);
  const files = new Map();
  let lastReport = 0;
  const progress_callback = (p) => {
    if (!p.file) return;
    const f = files.get(p.file) || { loaded: 0, total: 0 };
    if (p.status === "progress") {
      f.loaded = p.loaded || f.loaded;
      f.total = p.total || f.total;
    } else if (p.status === "done" && f.total) f.loaded = f.total;
    files.set(p.file, f);
    const t = performance.now();
    if (t - lastReport < 250 && p.status !== "done") return;
    lastReport = t;
    let loadedBytes = 0;
    let total = 0;
    for (const v of files.values()) {
      loadedBytes += v.loaded;
      total += v.total;
    }
    post("progress", { loaded: loadedBytes, total });
  };
  let model;
  let tokenizer;
  if (spec.kind === "vlm") {
    const processor = await L.AutoProcessor.from_pretrained(spec.id, { progress_callback });
    tokenizer = processor.tokenizer;
    post("phase", { phase: "compile" });
    model = await L.AutoModelForImageTextToText.from_pretrained(spec.id, {
      device: "webgpu",
      dtype: { embed_tokens: dtype, decoder_model_merged: dtype, vision_encoder: dtype },
      progress_callback,
    });
  } else {
    tokenizer = await L.AutoTokenizer.from_pretrained(spec.id, { progress_callback });
    post("phase", { phase: "compile" });
    model = await L.AutoModelForCausalLM.from_pretrained(spec.id, { dtype, device: "webgpu", progress_callback });
  }
  loaded = { key, model, tokenizer };
  // rozgrzewka — pierwsze wywołanie kompiluje shadery; lepiej teraz niż przy pierwszym pytaniu
  post("phase", { phase: "warmup" });
  await run({ messages: [{ role: "user", content: "OK" }], maxTokens: 2 }, null);
  post("ready", { key });
}

async function run({ messages, maxTokens }, id) {
  const { model, tokenizer } = loaded;
  const text = tokenizer.apply_chat_template(messages, { tokenize: false, add_generation_prompt: true, enable_thinking: false });
  const inputs = tokenizer(text, { add_special_tokens: false });
  const promptTokens = inputs.input_ids.dims.at(-1);
  let partial = "";
  let lastSent = 0;
  const streamer = new lib.TextStreamer(tokenizer, {
    skip_prompt: true,
    skip_special_tokens: true,
    callback_function: (piece) => {
      partial += piece;
      const now = performance.now();
      if (id != null && now - lastSent > 120) {
        lastSent = now;
        post("partial", { id, text: partial });
      }
    },
  });
  stopper = new lib.InterruptableStoppingCriteria();
  const out = await model.generate({
    ...inputs,
    max_new_tokens: maxTokens,
    do_sample: false,
    repetition_penalty: 1.05,
    streamer,
    stopping_criteria: stopper,
  });
  stopper = null;
  const seq = out.slice(null, [promptTokens, null]);
  return (tokenizer.batch_decode(seq, { skip_special_tokens: true })[0] || "").replace(/<think>[\s\S]*?<\/think>/g, "").trim();
}

async function release() {
  try {
    await loaded?.model?.dispose();
  } catch {}
  loaded = null;
}

self.addEventListener("unhandledrejection", (e) => {
  post("error", { message: String(e.reason?.message || e.reason).slice(0, 400) });
});

self.onmessage = async ({ data }) => {
  try {
    if (data.type === "cache-status") {
      post("cache-status", { model: data.model, cached: await cacheStatus(data) });
    } else if (data.type === "load") {
      await load(data);
    } else if (data.type === "generate") {
      if (!loaded) throw new Error("model not loaded");
      const text = await run(data, data.id);
      post("done", { id: data.id, text });
    } else if (data.type === "abort") {
      stopper?.interrupt();
    } else if (data.type === "release") {
      await release();
      post("released");
    } else if (data.type === "delete-cache") {
      await release();
      await caches.delete(data.cacheKey);
      post("deleted");
    }
  } catch (err) {
    post("error", { id: data.id, message: String(err?.message || err).slice(0, 400) });
  }
};
