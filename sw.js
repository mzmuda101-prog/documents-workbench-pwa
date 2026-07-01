const CACHE_VERSION = "20260702-01";
const APP_CACHE = `docs-wb-shell-${CACHE_VERSION}`;
const HEAVY_CACHE = `docs-wb-heavy-${CACHE_VERSION}`;
const RUNTIME_CACHE = `docs-wb-runtime-${CACHE_VERSION}`;
const ASSET_V = "20260702-01";

const SHELL_ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  `./styles/app.css?v=${ASSET_V}`,
  "./assets/fonts/space-grotesk-latin.woff2",
  "./assets/fonts/space-grotesk-latin-ext.woff2",
  `./app/core.js?v=${ASSET_V}`,
  `./app/language.js?v=${ASSET_V}`,
  `./app/docx-patch.js?v=${ASSET_V}`,
  `./app/template-tokens.js?v=${ASSET_V}`,
  `./app/docx-run-styles.js?v=${ASSET_V}`,
  `./app/docx-inline-edit.js?v=${ASSET_V}`,
  `./app/document.js?v=${ASSET_V}`,
  `./app/docx-render-fixes.js?v=${ASSET_V}`,
  `./app/docx-viewer.js?v=${ASSET_V}`,
  `./app/analysis.js?v=${ASSET_V}`,
  `./app/grammar-style.js?v=${ASSET_V}`,
  `./app/placeholders.js?v=${ASSET_V}`,
  `./app/snippets.js?v=${ASSET_V}`,
  `./app/edit-tools.js?v=${ASSET_V}`,
  `./app/mobile-doc-zoom.js?v=${ASSET_V}`,
  `./app/ui-controls.js?v=${ASSET_V}`,
  `./app/lazy-features.js?v=${ASSET_V}`,
  `./app/bootstrap.js?v=${ASSET_V}`,
  "./assets/images/favicon.png",
  "./assets/images/apple-touch-icon.png",
  "./assets/images/icon-192.png",
  "./assets/images/icon-512.png",
  "./docs/samples/sample.docx",
  // [EN] Lazy panel UI — cached for offline after first open
  `./app/grammar-panel.js?v=${ASSET_V}`,
  `./app/placeholders-panel.js?v=${ASSET_V}`,
  `./app/snippets-panel.js?v=${ASSET_V}`,
  `./app/find-replace-workbench.js?v=${ASSET_V}`,
];

// [EN] Large libs + media — separate bucket; still precached for offline after install
const HEAVY_ASSETS = [
  "./lib/jszip.min.js",
  "./lib/docx-preview.bundle.js",
  "./assets/media/mateusz-intro.mp4",
];

function isStaticAsset(url) {
  return /\.(?:css|js|png|svg|jpg|jpeg|gif|webp|ico|woff2?|mp4|docx|pdf)$/i.test(url.pathname);
}

function isHeavyAsset(url) {
  return /\/lib\/(?:jszip\.min|docx-preview\.bundle)\.js$/i.test(url.pathname)
    || /\/assets\/media\/mateusz-intro\.mp4$/i.test(url.pathname);
}

function cacheNameForUrl(url) {
  if (isHeavyAsset(url)) return HEAVY_CACHE;
  return RUNTIME_CACHE;
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    Promise.all([
      caches.open(APP_CACHE).then((cache) => cache.addAll(SHELL_ASSETS)),
      caches.open(HEAVY_CACHE).then((cache) => cache.addAll(HEAVY_ASSETS)),
    ]).catch(() => {})
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== APP_CACHE && key !== HEAVY_CACHE && key !== RUNTIME_CACHE)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener("message", (event) => {
  if (event.data && event.data.type === "SKIP_WAITING") self.skipWaiting();
});

self.addEventListener("fetch", (event) => {
  const { request } = event;
  if (request.method !== "GET") return;
  const reqUrl = new URL(request.url);
  const sameOrigin = reqUrl.origin === self.location.origin;

  if (request.mode === "navigate") {
    event.respondWith(
      fetch(request)
        .then((response) => {
          const copy = response.clone();
          caches.open(RUNTIME_CACHE).then((c) => c.put(request, copy)).catch(() => {});
          return response;
        })
        .catch(async () => (await caches.match(request)) || caches.match("./index.html"))
    );
    return;
  }

  if (sameOrigin && isStaticAsset(reqUrl)) {
    event.respondWith(
      caches.match(request).then((cached) => {
        const network = fetch(request)
          .then((response) => {
            if (response && response.ok) {
              const copy = response.clone();
              const bucket = isHeavyAsset(reqUrl) ? HEAVY_CACHE : cacheNameForUrl(reqUrl);
              caches.open(bucket).then((c) => c.put(request, copy)).catch(() => {});
            }
            return response;
          })
          .catch(() => cached);
        return cached || network;
      })
    );
  }
});
