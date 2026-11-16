// offline-status.js — czy aplikacja wstanie bez internetu (wiersz w menu ⋯).
// KOPIA 1:1 w sheet-workbench-pwa/app/offline-status.js — zmieniaj oba pliki.
//
// Zgłoszenie (2026-11-16): na niektórych laptopach zainstalowana aplikacja nie wstaje bez sieci.
// Kod offline działa (sprawdzone na produkcji), ale przeglądarka może nie zapisywać danych strony
// (tryb prywatny, zablokowane dane witryn) albo kasować je przy zamknięciu (komputer szkolny /
// firmowy, „usuwaj dane przy zamknięciu”, dodatki do prywatności). Użytkownik ma to zobaczyć
// od razu, zanim zostanie bez internetu.
//
// Stany: ready (gotowa), preparing (pierwsze wejście — pliki się zapisują), volatile (zainstalowana,
// ale przeglądarka nie dała trwałego miejsca — może skasować dane), blocked (nie da się zapisać).
// Przy aplikacji zainstalowanej prosimy od razu o trwałe miejsce (navigator.storage.persist) —
// w zwykłej karcie nie, bo Firefox pyta wtedy użytkownika okienkiem.

const dwbOffline = (() => {
  const el = document.getElementById("offlineState");
  if (!el) return null;
  const standalone = matchMedia("(display-mode: standalone)").matches || navigator.standalone === true;
  const ICONS = {
    ready: '<path d="M20 6 9 17l-5-5"/>',
    preparing: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    volatile: '<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
    blocked: '<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
  };
  let state = null;
  let registerFailed = false;

  async function detect() {
    if (!("serviceWorker" in navigator) || typeof caches === "undefined") return "blocked";
    try {
      const keys = await caches.keys(); // zablokowane dane witryny → SecurityError
      const reg = await navigator.serviceWorker.getRegistration();
      // bez rejestracji po kilku sekundach = przeglądarka jej nie dopuściła (nie każda zgłasza błąd)
      if (!reg) return registerFailed || performance.now() > 10000 ? "blocked" : "preparing";
      const shell = keys.filter((k) => /-shell-/.test(k));
      let ok = false;
      for (const k of shell) {
        if (await (await caches.open(k)).match("./index.html")) { ok = true; break; }
      }
      if (!ok || !navigator.serviceWorker.controller) return "preparing";
      if (standalone && navigator.storage?.persisted) {
        let p = await navigator.storage.persisted().catch(() => null);
        if (p === false && navigator.storage.persist) p = await navigator.storage.persist().catch(() => false);
        if (p === false) return "volatile";
      }
      return "ready";
    } catch (_) {
      return "blocked";
    }
  }
  function render() {
    if (!state) return;
    el.hidden = false;
    el.dataset.state = state;
    const key = { ready: "offlineReady", preparing: "offlinePreparing", volatile: "offlineVolatile", blocked: "offlineBlocked" }[state];
    el.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[state]}</svg><span></span>`;
    el.querySelector("span").innerHTML = t(key); // tekst z tłumaczeń (zawiera <b>)
  }
  async function check() {
    state = await detect();
    render();
  }

  // rejestracja SW nie przeszła (ui-controls.js) → przeglądarka nie zapisze aplikacji
  window.addEventListener("dwb:sw-failed", () => { registerFailed = true; check(); });
  navigator.serviceWorker?.addEventListener?.("controllerchange", check);
  window.addEventListener("load", () => setTimeout(check, 2500));
  // przy otwarciu menu ⋯ — stan aktualny (np. pliki dograły się w tle)
  document.getElementById("appMenuBtn")?.addEventListener("click", check);
  if (typeof window.applyLanguage === "function") {
    const orig = window.applyLanguage;
    window.applyLanguage = function applyLanguageOffline(...args) { const r = orig.apply(this, args); render(); return r; };
  }

  return { check, state: () => state };
})();
