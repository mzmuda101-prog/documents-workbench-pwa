const CACHE_VERSION = "20260630-18";
const APP_CACHE = `docs-wb-shell-${CACHE_VERSION}`;
const RUNTIME_CACHE = `docs-wb-runtime-${CACHE_VERSION}`;

const APP_ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./styles/app.css?v=20260630-18",
  "./assets/fonts/space-grotesk-latin.woff2",
  "./assets/fonts/space-grotesk-latin-ext.woff2",
  "./assets/media/mateusz-intro.mp4",
  "./lib/jszip.min.js",
  "./lib/docx-preview.bundle.js",
  "./app/core.js?v=20260630-18",
  "./app/language.js?v=20260630-18",
  "./app/docx-patch.js?v=20260630-18",
  "./app/docx-inline-edit.js?v=20260630-18",
  "./app/document.js?v=20260630-18",
  "./app/docx-render-fixes.js?v=20260630-18",
  "./app/docx-viewer.js?v=20260630-18",
  "./app/analysis.js?v=20260630-18",
  "./app/grammar-style.js?v=20260630-18",
  "./app/grammar-panel.js?v=20260630-18",
  "./app/find-replace-workbench.js?v=20260630-18",
  "./app/edit-tools.js?v=20260630-18",
  "./app/mobile-doc-zoom.js?v=20260630-18",
  "./app/ui-controls.js?v=20260630-18",
  "./app/bootstrap.js?v=20260630-18",
  "./assets/images/favicon.png",
  "./assets/images/apple-touch-icon.png",
  "./assets/images/icon-192.png",
  "./assets/images/icon-512.png",
  "./docs/samples/sample.docx",
];

function isStaticAsset(url) {
  return /\.(?:css|js|png|svg|jpg|jpeg|gif|webp|ico|woff2?|mp4|docx|pdf)$/i.test(url.pathname);
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(APP_CACHE).then((cache) => cache.addAll(APP_ASSETS)).catch(() => {})
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== APP_CACHE && k !== RUNTIME_CACHE).map((k) => caches.delete(k)))
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
              caches.open(RUNTIME_CACHE).then((c) => c.put(request, copy)).catch(() => {});
            }
            return response;
          })
          .catch(() => cached);
        return cached || network;
      })
    );
  }
});
