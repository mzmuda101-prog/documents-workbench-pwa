// backstage.js — ekran „Plik” jak w Wordzie (etap 2 wstążki; docs/WSTAZKA.md, makieta v4).
//
// Wjeżdża pod rzędem kart tym samym ruchem co dawny panel boczny (280 ms, --ease-drawer).
// Po lewej lista: Wróć, Nowy, Otwórz, Zdjęcie, Informacje, Zapisz jako, Eksportuj, Drukuj,
// Porównaj, Zamknij · Opcje, Pomoc, Zgłoś błąd. Akcje klikają TE SAME przyciski co dawne menu ⋯
// (ich obsługa zostaje jedna); sekcje przejmują istniejące panele i przełączniki przez gniazda
// [data-rb] (ribbon.js je przenosi — na telefonie wracają na swoje miejsca).

const dwbBackstage = (() => {
  const el = document.getElementById("backstage");
  if (!el) return null;
  const nav = el.querySelector(".bs-nav");
  const secs = Array.from(el.querySelectorAll(".bs-sec"));
  const recentHost = document.getElementById("bsRecent");
  let current = "new";
  let opener = null;

  function place() {
    const r1 = document.querySelector(".rb-r1");
    if (r1) el.style.top = `${Math.round(r1.getBoundingClientRect().bottom)}px`;
  }
  function show(sec) {
    if (!secs.some((s) => s.dataset.bsSec === sec)) sec = "new";
    current = sec;
    secs.forEach((s) => { s.hidden = s.dataset.bsSec !== sec; });
    nav.querySelectorAll("[data-bs-go]").forEach((b) => {
      const on = b.dataset.bsGo === sec;
      b.classList.toggle("is-on", on);
      if (on) b.setAttribute("aria-current", "page"); else b.removeAttribute("aria-current");
    });
    // panele przeniesione do sekcji: zawsze rozwinięte (nagłówek <summary> schowany w CSS)
    secs.find((s) => s.dataset.bsSec === sec)?.querySelectorAll("details.panel").forEach((d) => { d.open = true; });
    if (sec === "new" && recentHost && typeof dwbComfort !== "undefined") dwbComfort.renderRecent(recentHost);
    // wskaźnik języka liczy się z wymiarów — w schowanym miejscu były zerowe
    if (sec === "opts" && typeof syncLangSwitchPill === "function") requestAnimationFrame(syncLangSwitchPill);
  }
  function isOpen() { return !el.hidden && el.classList.contains("is-open"); }
  function open(sec = "new") {
    if (typeof dwbRibbon !== "undefined" && dwbRibbon && !dwbRibbon.isOn()) return false;
    // fokus wraca tam, skąd przyszedł — przy kliknięciu „Plik” (bez fokusu: mousedown) na „Plik”
    opener = document.activeElement && document.activeElement !== document.body ? document.activeElement : document.querySelector(".rb-file");
    document.body.classList.toggle("bs-has-doc", !!originalFileBytes);
    place();
    el.hidden = false;
    show(sec);
    document.body.classList.add("bs-open");
    requestAnimationFrame(() => {
      el.classList.add("is-open");
      (nav.querySelector(".is-on") || nav.querySelector("button"))?.focus({ preventScroll: true });
    });
    return true;
  }
  let hideTimer = 0;
  function close({ refocus = true } = {}) {
    if (el.hidden) return;
    el.classList.remove("is-open");
    document.body.classList.remove("bs-open");
    clearTimeout(hideTimer);
    hideTimer = setTimeout(() => { if (!el.classList.contains("is-open")) el.hidden = true; }, 300);
    if (refocus) (opener?.isConnected ? opener : document.querySelector(".rb-file"))?.focus?.({ preventScroll: true });
  }

  el.addEventListener("click", (e) => {
    const go = e.target.closest("[data-bs-go]");
    if (go) { show(go.dataset.bsGo); return; }
    const d = e.target.closest("[data-bs-do]");
    if (!d) return;
    const id = d.dataset.bsDo;
    if (id === "back") { close(); return; }
    // akcja dawnego menu ⋯ / ekranu startowego (ta sama obsługa) — ekran Plik się zamyka
    close({ refocus: false });
    document.getElementById(id)?.click();
  });
  el.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !e.target.closest("input, select, textarea, [contenteditable='true']")) {
      e.preventDefault();
      e.stopPropagation();
      close();
      return;
    }
    // strzałki w liście po lewej (jak w menu)
    if ((e.key === "ArrowDown" || e.key === "ArrowUp") && e.target.closest(".bs-nav")) {
      const items = Array.from(nav.querySelectorAll("button:not([hidden])")).filter((b) => b.offsetParent);
      const i = items.indexOf(document.activeElement);
      e.preventDefault();
      items[(i + (e.key === "ArrowDown" ? 1 : -1) + items.length) % items.length]?.focus();
    }
  });
  window.addEventListener("resize", () => { if (!el.hidden) place(); }, { passive: true });
  // wczytanie / zamknięcie pliku albo przejście na telefon — ekran Plik się zamyka
  document.getElementById("fileInput")?.addEventListener("change", () => close({ refocus: false }));
  matchMedia("(min-width: 769px)").addEventListener("change", (e) => { if (!e.matches) close({ refocus: false }); });

  return { open, close, show, isOpen, current: () => current };
})();
