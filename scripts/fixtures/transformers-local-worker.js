// transformers-local-playwright.js: biblioteka AI z lib/transformers w workerze modułowym
// (jak local-ai-worker.js). FULL=1 → prawdziwy mały model (sieć: Hugging Face), backend wasm.
const TRANSFORMERS_URL = new URL("../../lib/transformers/transformers.min.mjs?v=4.3.1", import.meta.url).href;
const ORT_WASM_PATHS = {
  mjs: new URL("../../lib/transformers/ort-wasm-simd-threaded.asyncify.mjs?v=4.3.1", import.meta.url).href,
  wasm: new URL("../../lib/transformers/ort-wasm-simd-threaded.asyncify.wasm", import.meta.url).href,
};
self.onmessage = async ({ data }) => {
  try {
    const lib = await import(TRANSFORMERS_URL);
    lib.env.backends.onnx.wasm.wasmPaths = { ...ORT_WASM_PATHS };
    lib.env.allowLocalModels = false;
    lib.env.useBrowserCache = false;
    const out = { imported: typeof lib.pipeline === "function", paths: { ...lib.env.backends.onnx.wasm.wasmPaths } };
    if (data.full) {
      const extractor = await lib.pipeline("feature-extraction", "Xenova/all-MiniLM-L6-v2", { device: "wasm", dtype: "q8" });
      const t = await extractor("Ala ma kota", { pooling: "mean", normalize: true });
      out.dims = t.dims;
      out.norm = Math.sqrt([...t.data].reduce((s, v) => s + v * v, 0));
    }
    postMessage(out);
  } catch (e) {
    postMessage({ error: String(e && (e.stack || e.message) || e) });
  }
};
