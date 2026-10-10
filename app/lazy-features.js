// [EN] Lazy-load sidebar feature bundles on first panel open — keeps boot parse lighter.
const LAZY_FEATURE_SCRIPTS = {
  grammar: ["app/grammar-panel.js", "app/a11y-check.js"],
  metadata: ["app/metadata-panel.js"],
  placeholders: ["app/placeholders-panel.js"],
  "snippets-panel": ["app/snippets-panel.js"],
  "find-replace": ["app/find-replace-workbench.js"],
  export: ["app/export-panel.js"],
  stats: ["app/stats-panel.js"],
  review: ["app/review-panel.js"],
  forms: ["app/forms-panel.js"],
  compare: ["app/compare-panel.js"],
};

const PANEL_LAZY_FEATURE = {
  "panel-metadata": "metadata",
  "panel-grammar": "grammar",
  "panel-placeholders": "placeholders",
  "panel-snippets": "snippets-panel",
  "panel-search": "find-replace",
  "panel-export": "export",
  "panel-stats": "stats",
  "panel-review": "review",
  "panel-forms": "forms",
  "panel-compare": "compare",
};

const lazyFeatureLoaded = new Set();
const lazyFeatureLoading = new Map();

function lazyAssetVersion() {
  if (typeof APP_BUILD_VERSION === "string" && APP_BUILD_VERSION) return APP_BUILD_VERSION;
  const src = document.querySelector('script[src*="core.js"]')?.getAttribute("src") || "";
  const m = src.match(/[?&]v=([^&]+)/);
  return m ? m[1] : "1";
}

function loadLazyScript(relativePath) {
  const base = relativePath.split("?")[0];
  if (document.querySelector(`script[src*="${base}"]`)) return Promise.resolve();
  const v = lazyAssetVersion();
  return new Promise((resolve, reject) => {
    const el = document.createElement("script");
    el.src = `${relativePath}?v=${encodeURIComponent(v)}`;
    el.defer = true;
    el.onload = () => resolve();
    el.onerror = () => reject(new Error(`lazy script failed: ${relativePath}`));
    document.head.appendChild(el);
  });
}

function ensureLazyFeature(featureKey) {
  if (!featureKey || lazyFeatureLoaded.has(featureKey)) return Promise.resolve();
  if (lazyFeatureLoading.has(featureKey)) return lazyFeatureLoading.get(featureKey);
  const files = LAZY_FEATURE_SCRIPTS[featureKey];
  if (!files?.length) return Promise.resolve();
  const job = (async () => {
    for (const file of files) await loadLazyScript(file);
    lazyFeatureLoaded.add(featureKey);
  })().finally(() => {
    lazyFeatureLoading.delete(featureKey);
  });
  lazyFeatureLoading.set(featureKey, job);
  return job;
}

function initLazyFeaturePanels() {
  document.querySelectorAll("details.panel").forEach((panel) => {
    const featureKey = PANEL_LAZY_FEATURE[panel.id];
    if (!featureKey) return;
    panel.addEventListener("toggle", () => {
      if (!panel.open) return;
      ensureLazyFeature(featureKey).then(() => {
        if (panel.id === "panel-metadata" && typeof loadMetadataFromDocument === "function") {
          loadMetadataFromDocument().catch(() => {});
        }
        if (panel.id === "panel-stats" && typeof renderDocumentStats === "function") renderDocumentStats();
        if (panel.id === "panel-review" && typeof runReviewScan === "function") runReviewScan().catch(() => {});
        if (panel.id === "panel-forms" && typeof renderFormsPanel === "function") renderFormsPanel();
      }).catch(() => {});
    });
  });
}
