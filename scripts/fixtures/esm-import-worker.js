// security-1106-playwright.js: worker modułowy z importem z innej domeny (jak workery AI).
// Sprawdza, że zaostrzone CSP strony (bez https://esm.sh) nie blokuje importu w workerze.
import("https://esm.sh/dwb-test-stub").then((m) => postMessage(m.ok || "brak"), (e) => postMessage(`ERR ${e && e.message}`));
