// Strojenie poleceń pod małe modele: jeden model wczytany raz, potem seria zadań z aplikacji.
// Zadanie: { messages, maxTokens, prefix } — prefix = wymuszony początek odpowiedzi („{”, „=”),
// dopisany po szablonie czatu; mały model nie zaczyna wtedy od „Oto…”.

const TRANSFORMERS_URL = "https://esm.sh/@huggingface/transformers@4.3.1";
const CACHE_KEY = "dwb-llm-bench-v1";
let L = null;
let model = null;
let tokenizer = null;
const post = (type, payload = {}) => self.postMessage({ type, ...payload });

async function load(spec, dtype) {
  L ||= await import(TRANSFORMERS_URL);
  L.env.useBrowserCache = true;
  L.env.cacheKey = CACHE_KEY;
  L.env.allowLocalModels = false;
  let last = 0;
  const progress_callback = (p) => {
    if (p.status === "progress" && performance.now() - last > 500) { last = performance.now(); post("progress", { file: p.file, loaded: p.loaded, total: p.total }); }
  };
  if (spec.kind === "vlm") {
    const processor = await L.AutoProcessor.from_pretrained(spec.id, { progress_callback });
    tokenizer = processor.tokenizer;
    model = await L.AutoModelForImageTextToText.from_pretrained(spec.id, { device: "webgpu", dtype: { embed_tokens: dtype, decoder_model_merged: dtype, vision_encoder: dtype }, progress_callback });
  } else {
    tokenizer = await L.AutoTokenizer.from_pretrained(spec.id, { progress_callback });
    model = await L.AutoModelForCausalLM.from_pretrained(spec.id, { dtype, device: "webgpu", progress_callback });
  }
}

async function run({ messages, maxTokens, prefix = "" }) {
  const text = tokenizer.apply_chat_template(messages, { tokenize: false, add_generation_prompt: true, enable_thinking: false }) + prefix;
  const inputs = tokenizer(text, { add_special_tokens: false });
  const n = inputs.input_ids.dims.at(-1);
  const t0 = performance.now();
  const out = await model.generate({ ...inputs, max_new_tokens: maxTokens, do_sample: false, repetition_penalty: 1.05 });
  const decoded = tokenizer.batch_decode(out.slice(null, [n, null]), { skip_special_tokens: true })[0] || "";
  return { text: prefix + decoded.replace(/<think>[\s\S]*?<\/think>/g, "").trim(), ms: performance.now() - t0 };
}

self.addEventListener("unhandledrejection", (e) => post("error", { message: String(e.reason?.message || e.reason) }));
self.onmessage = async ({ data }) => {
  try {
    if (data.type === "load") { await load(data.spec, data.dtype); post("loaded"); }
    if (data.type === "run") post("result", { id: data.id, ...(await run(data.task)) });
  } catch (err) {
    post("error", { id: data.id, message: String(err?.message || err) });
  }
};
