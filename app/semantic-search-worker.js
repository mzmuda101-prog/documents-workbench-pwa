// Optional EmbeddingGemma 2 runtime. This worker is never created during normal app startup.
// The model and all document text stay in the browser; only model files are fetched from HF.

const MODEL_ID = "onnx-community/embeddinggemma-2-ONNX";
// This names the persistent on-device download, independently of the PWA shell version.
// Bump it only when deliberately shipping a different model or quantization.
const MODEL_CACHE_KEY = "dwb-embeddinggemma-2-q4-v1";
const DIMENSIONS = 256; // Matryoshka: 3x smaller index, near-full text retrieval quality.
const DOCUMENT_PREFIX = "title: none | text: ";
const QUERY_PREFIX = "task: search result | query: ";

let embedder = null;
let runtime = null;
let textOnlyConfig = null;

function post(type, payload = {}, transfer = []) {
  self.postMessage({ type, ...payload }, transfer);
}

function truncateAndNormalize(values) {
  const out = new Float32Array(DIMENSIONS);
  let norm = 0;
  for (let i = 0; i < DIMENSIONS; i++) {
    const v = Number(values[i] || 0);
    out[i] = v;
    norm += v * v;
  }
  norm = Math.sqrt(norm) || 1;
  for (let i = 0; i < DIMENSIONS; i++) out[i] /= norm;
  return out;
}

async function getEmbedder() {
  if (embedder) return embedder;
  if (!navigator.gpu) throw new Error("WebGPU is unavailable");
  // Dynamic import means a load failure is reported back to the panel instead of killing the
  // module worker before it can explain what happened.
  // Keep the optional runtime off the release payload. esm.sh resolves the browser bundle's
  // `onnxruntime-web/webgpu` dependency to an ESM URL; the plain npm CDN file leaves that
  // bare specifier unresolved in a static PWA.
  runtime ||= import("https://esm.sh/@huggingface/transformers@4.3.1");
  const { AutoTokenizer, EmbeddingGemma2Model, env, mean_pooling } = await runtime;
  if (!textOnlyConfig) {
    const response = await fetch(`https://huggingface.co/${MODEL_ID}/raw/main/config.json`);
    if (!response.ok) throw new Error(`Could not load model configuration (${response.status})`);
    textOnlyConfig = await response.json();
    // Transformers.js uses these keys to decide which ONNX sessions to download. Omitting
    // both prevents the 170M vision and 300M audio encoders from ever entering this feature.
    textOnlyConfig.vision_config = null;
    textOnlyConfig.audio_config = null;
  }
  env.allowLocalModels = false;
  // Cache model weights and ONNX Runtime persistently in the browser. Releasing app UI updates
  // must not make a user download the same q4 model again.
  env.useBrowserCache = true;
  env.useWasmCache = true;
  env.cacheKey = MODEL_CACHE_KEY;
  // WebGPU avoids a very slow WASM fallback and keeps UI work off the main thread.
  const progress_callback = (event) => {
      if (event.status === "progress") post("progress", { loaded: event.loaded, total: event.total, file: event.file });
      else if (event.status === "initiate" || event.status === "download") post("status", { message: "Pobieranie modelu…" });
    };
  // Do not use the generic pipeline here: it detects the multimodal model and eagerly loads
  // its vision/audio sessions. The text model's forward path accepts token inputs directly.
  const [tokenizer, model] = await Promise.all([
    AutoTokenizer.from_pretrained(MODEL_ID, { progress_callback }),
    EmbeddingGemma2Model.from_pretrained(MODEL_ID, {
      config: textOnlyConfig,
      device: "webgpu",
      dtype: "q4",
      progress_callback,
    }),
  ]);
  embedder = async (inputs) => {
    const modelInputs = tokenizer(inputs, { padding: true, truncation: true });
    const outputs = await model(modelInputs);
    return mean_pooling(outputs.last_hidden_state, modelInputs.attention_mask).normalize(2, -1);
  };
  return embedder;
}

async function embed(texts, kind) {
  const model = await getEmbedder();
  const input = texts.map((text) => (kind === "query" ? QUERY_PREFIX : DOCUMENT_PREFIX) + text);
  const tensor = await model(input);
  const vectors = [];
  const stride = tensor.dims[tensor.dims.length - 1];
  for (let i = 0; i < input.length; i++) vectors.push(truncateAndNormalize(tensor.data.subarray(i * stride, (i + 1) * stride)));
  return vectors;
}

self.onmessage = async ({ data }) => {
  try {
    if (data.type === "warmup") {
      await getEmbedder();
      post("ready");
      return;
    }
    if (data.type === "embed-documents") {
      const vectors = await embed(data.texts || [], "document");
      const buffers = vectors.map((v) => v.buffer);
      post("document-vectors", { requestId: data.requestId, vectors: buffers }, buffers);
      return;
    }
    if (data.type === "embed-query") {
      const [vector] = await embed([data.query || ""], "query");
      post("query-vector", { requestId: data.requestId, vector: vector.buffer }, [vector.buffer]);
      return;
    }
    if (data.type === "release") {
      embedder = null;
      post("released");
    }
  } catch (error) {
    post("error", { requestId: data.requestId, message: error?.message || "EmbeddingGemma 2 failed to start." });
  }
};
