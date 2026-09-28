const CACHE_VERSION = "20260928-24";
const APP_CACHE = `docs-wb-shell-${CACHE_VERSION}`;
const HEAVY_CACHE = `docs-wb-heavy-${CACHE_VERSION}`;
const RUNTIME_CACHE = `docs-wb-runtime-${CACHE_VERSION}`;
// Spięte z CACHE_VERSION (nie osobna stała) — inaczej `npm run release` podbija tylko
// CACHE_VERSION i ?v= w index.html, a precache celuje w adresy, o które strona już nie prosi.
const ASSET_V = CACHE_VERSION;
// Po tylu ms bez odpowiedzi sieci start idzie z cache (patrz handler nawigacji).
const NAVIGATION_TIMEOUT_MS = 3000;

const SHELL_ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  `./styles/app.css?v=${ASSET_V}`,
  "./assets/fonts/space-grotesk-latin.woff2",
  "./assets/fonts/space-grotesk-latin-ext.woff2",
  `./app/core.js?v=${ASSET_V}`,
  `./app/language.js?v=${ASSET_V}`,
  `./app/docx-metadata.js?v=${ASSET_V}`,
  `./app/docx-revisions.js?v=${ASSET_V}`,
  `./app/docx-patch.js?v=${ASSET_V}`,
  `./app/template-tokens.js?v=${ASSET_V}`,
  `./app/docx-run-styles.js?v=${ASSET_V}`,
  `./app/docx-inline-edit.js?v=${ASSET_V}`,
  `./app/document.js?v=${ASSET_V}`,
  `./app/docx-render-fixes.js?v=${ASSET_V}`,
  `./app/docx-viewer.js?v=${ASSET_V}`,
  `./app/analysis.js?v=${ASSET_V}`,
  `./app/structure-panel.js?v=${ASSET_V}`,
  `./app/grammar-style.js?v=${ASSET_V}`,
  `./app/placeholders.js?v=${ASSET_V}`,
  `./app/snippets.js?v=${ASSET_V}`,
  `./app/edit-tools.js?v=${ASSET_V}`,
  `./app/mobile-doc-zoom.js?v=${ASSET_V}`,
  `./app/ui-controls.js?v=${ASSET_V}`,
  `./app/lazy-features.js?v=${ASSET_V}`,
  `./app/app-frame.js?v=${ASSET_V}`,
  `./app/touch.js?v=${ASSET_V}`,
  `./app/pinch-zoom.js?v=${ASSET_V}`,
  `./app/undo.js?v=${ASSET_V}`,
  `./app/view-mode.js?v=${ASSET_V}`,
  `./app/keyboard.js?v=${ASSET_V}`,
  `./app/cursor-hint.js?v=${ASSET_V}`,
  `./app/bootstrap.js?v=${ASSET_V}`,
  "./assets/images/favicon.png",
  "./assets/images/apple-touch-icon.png",
  "./assets/images/icon-192.png",
  "./assets/images/icon-512.png",
  "./docs/samples/sample.docx",
  // [EN] Lazy panel UI — cached for offline after first open
  `./app/export-panel.js?v=${ASSET_V}`,
  `./app/stats-panel.js?v=${ASSET_V}`,
  `./app/review-panel.js?v=${ASSET_V}`,
  `./app/grammar-panel.js?v=${ASSET_V}`,
  `./app/placeholders-panel.js?v=${ASSET_V}`,
  `./app/snippets-panel.js?v=${ASSET_V}`,
  `./app/find-replace-workbench.js?v=${ASSET_V}`,
  `./app/metadata-panel.js?v=${ASSET_V}`,
];

// Ciężkie biblioteki + film intro — osobny kubełek, dogrywany PO aktywacji (niżej).
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

// Lokalnie (npm run dev) pliki zmieniają się bez podbicia ?v=, więc tam zostaje
// stale-while-revalidate — inaczej po edycji widać by było stary kod aż do release.
function isImmutableAsset(url) {
  if (/^(?:localhost|127\.0\.0\.1|\[::1\])$/.test(self.location.hostname)) return false;
  return url.searchParams.has("v");
}

function cacheNameForUrl(url) {
  if (isHeavyAsset(url)) return HEAVY_CACHE;
  return RUNTIME_CACHE;
}

// Instalacja pobiera TYLKO lekką powłokę. Ciężkie zasoby (docx-preview, jszip, film
// intro ~1,6 MB) szły wcześniej w tej samej paczce — czyli zaraz po każdej aktualizacji
// telefon ściągał kilka megabajtów dokładnie wtedy, gdy użytkownik otwiera dokument.
// Teraz dogrywamy je po aktywacji, z opóźnieniem; gdyby SW został w międzyczasie
// uśpiony — i tak trafią do cache przy pierwszym użyciu (handler fetch niżej).
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(APP_CACHE).then((cache) => cache.addAll(SHELL_ASSETS)).catch(() => {})
  );
});

function precacheHeavyAssetsLater(delayMs = 8000) {
  return new Promise((resolve) => {
    setTimeout(async () => {
      try {
        const cache = await caches.open(HEAVY_CACHE);
        for (const asset of HEAVY_ASSETS) {
          // Pojedynczo i sekwencyjnie — równoległe addAll potrafi zapchać łącze telefonu.
          if (await cache.match(asset)) continue;
          try { await cache.add(asset); } catch (_) { /* dogramy przy pierwszym użyciu */ }
        }
      } catch (_) {
        // brak miejsca / tryb prywatny — zostaje ścieżka „cache przy pierwszym użyciu"
      }
      resolve();
    }, delayMs);
  });
}

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
  // Ciężkie zasoby dogrywamy po chwili, już poza ścieżką krytyczną startu.
  event.waitUntil(precacheHeavyAssetsLater());
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
    // Sieć najpierw (świeży index.html), ale z LIMITEM czasu. Bez niego przy słabym
    // zasięgu („jest kreska, a nic nie przechodzi") start z ikony wisiał na białym
    // ekranie nawet kilkadziesiąt sekund, choć cała apka leży w cache. Po limicie
    // podajemy stronę z cache; żądanie leci dalej w tle i odświeża cache na następny raz.
    const network = fetch(request).then((response) => {
      if (response && response.ok) {
        const copy = response.clone();
        caches.open(RUNTIME_CACHE).then((cache) => cache.put(request, copy)).catch(() => {});
      }
      return response;
    });
    const fromCache = async () => (await caches.match(request)) || caches.match("./index.html");
    event.respondWith(new Promise((resolve) => {
      let settled = false;
      const finish = (res) => { if (!settled && res) { settled = true; resolve(res); } };
      const timer = setTimeout(async () => {
        const cached = await fromCache();
        if (cached) finish(cached);
      }, NAVIGATION_TIMEOUT_MS);
      network
        .then((res) => { clearTimeout(timer); finish(res); })
        .catch(async () => {
          clearTimeout(timer);
          const cached = await fromCache();
          finish(cached || Response.error());
        });
    }));
    event.waitUntil(network.catch(() => {}));
    return;
  }

  if (sameOrigin && isStaticAsset(reqUrl) && isImmutableAsset(reqUrl)) {
    // Plik z wersją w adresie (?v=…) nigdy się nie zmienia — nowa wersja apki to nowy
    // adres. Dawniej każdy start dopytywał sieć o ~25 takich plików; teraz z cache,
    // a sieć tylko gdy pliku w cache brak.
    event.respondWith(
      caches.match(request).then((cached) => cached || fetch(request).then((response) => {
        if (response && response.ok) {
          const copy = response.clone();
          caches.open(cacheNameForUrl(reqUrl)).then((cache) => cache.put(request, copy)).catch(() => {});
        }
        return response;
      }))
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
              caches.open(cacheNameForUrl(reqUrl)).then((c) => c.put(request, copy)).catch(() => {});
            }
            return response;
          })
          .catch(() => cached);
        return cached || network;
      })
    );
  }
});
