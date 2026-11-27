// rail.js — szyna paneli przy wstążce (etap 3; docs/WSTAZKA.md, makieta v4).
//
// Wąska szyna po lewej: Nawigacja, Znajdź i zamień, Zmiany i komentarze, Korekta, Asystent —
// z licznikami (te same, co dotąd w nagłówkach sekcji panelu). Klik wysuwa DAWNY panel boczny
// (ten sam ruch i to samo przypinanie co wcześniej), ale pokazuje w nim jeden panel naraz,
// z tytułem, pinezką i ×. Bez pinezki panel kładzie się na marginesie (dokument zostaje),
// z pinezką stoi obok dokumentu (dawne „panel obok od 1024 px”).
// Każda droga do panelu (skrót, przycisk karty, kod, który ustawia details.open) pokazuje go
// tutaj — nasłuch zdarzenia „toggle”. Na telefonie (bez wstążki) panel działa jak dawniej.

const dwbRail = (() => {
  const rail = document.getElementById("rbRail");
  const sidebar = document.getElementById("toolsPanel");
  if (!rail || !sidebar) return null;
  const titleEl = document.getElementById("railTitle");
  const pinBtn = document.getElementById("railPinBtn");
  const PIN_KEY = "dwb-rail-pin-v1";
  const ITEMS = {
    nav: ["panel-structure"],
    find: ["panel-search"],
    review: ["panel-review"],
    proof: ["panel-grammar"],
    ai: ["panel-ask", "panel-semantic-search"],
  };
  let active = [];
  const on = () => document.body.classList.contains("rb-on");
  const itemOf = (id) => Object.keys(ITEMS).find((k) => ITEMS[k].includes(id));

  function pinned() { try { return localStorage.getItem(PIN_KEY) === "1"; } catch (_) { return false; } }

  // ── który panel jest pokazany ─────────────────────────────────────────────
  function show(ids) {
    active = ids;
    sidebar.querySelectorAll("details.panel").forEach((d) => {
      const a = ids.includes(d.id);
      d.classList.toggle("rb-active", a);
      if (a && !d.open) d.open = true;
    });
    const sum = sidebar.querySelector(`#${ids[0]} > summary`);
    if (titleEl) {
      const item = itemOf(ids[0]);
      titleEl.textContent = item ? t({ nav: "railNav", find: "railFind", review: "railReview", proof: "railProof", ai: "railAi" }[item])
        : (sum?.querySelector("[data-i18n]")?.textContent || sum?.textContent || "").trim();
    }
    if (typeof appFrame !== "undefined") {
      const wasDocked = document.documentElement.classList.contains("sidebar-docked");
      appFrame.syncDock();
      if (wasDocked !== document.documentElement.classList.contains("sidebar-docked") && isSidebarOpen()) setSidebarOpen(true);
    }
    sync();
  }
  function open(id) {
    const item = itemOf(id);
    show(item ? ITEMS[item] : [id]);
    if (!isSidebarOpen()) setSidebarOpen(true);
    requestAnimationFrame(() => document.getElementById(id)?.scrollIntoView({ block: "start" }));
  }
  function toggle(item) {
    if (isSidebarOpen() && active[0] === ITEMS[item][0]) { setSidebarOpen(false); return; }
    open(ITEMS[item][0]);
  }
  rail.addEventListener("click", (e) => {
    const b = e.target.closest("[data-rail]");
    if (b) toggle(b.dataset.rail);
  });
  // panel otwarty skądkolwiek (details.open = true) staje się tym pokazanym; „toggle” nie bąbelkuje,
  // ale przechodzi fazę przechwytywania
  sidebar.addEventListener("toggle", (e) => {
    const d = e.target;
    if (!on() || !d.matches?.("details.panel") || !d.open || active.includes(d.id)) return;
    const item = itemOf(d.id);
    show(item ? ITEMS[item] : [d.id]);
  }, true);

  // ── stan przycisków szyny i pinezki ───────────────────────────────────────
  function sync() {
    // szyna jedzie razem z obszarem roboczym (przypięty panel poszerza ramę) — panel za nią
    place();
    setTimeout(place, 320);
    const openNow = isSidebarOpen();
    rail.querySelectorAll("[data-rail]").forEach((b) => {
      b.setAttribute("aria-pressed", String(openNow && ITEMS[b.dataset.rail][0] === active[0]));
    });
    const p = pinned();
    pinBtn?.setAttribute("aria-pressed", String(p));
    pinBtn?.setAttribute("aria-label", t(p ? "railUnpin" : "railPin"));
  }
  new MutationObserver(sync).observe(document.documentElement, { attributes: true, attributeFilter: ["class"] });
  pinBtn?.addEventListener("click", () => {
    try { localStorage.setItem(PIN_KEY, pinned() ? "0" : "1"); } catch (_) {}
    if (typeof appFrame !== "undefined") appFrame.syncDock();
    setSidebarOpen(isSidebarOpen()); // przelicza dokument (przypięty = węższy)
    sync();
  });

  // ── liczniki: te same co w nagłówkach sekcji (app-frame.js setSummaryCount) ──
  let cq = 0;
  function syncCounts() {
    cq = 0;
    rail.querySelectorAll("[data-rail]").forEach((b) => {
      const n = ITEMS[b.dataset.rail].reduce((sum, id) => sum + (parseInt(sidebar.querySelector(`#${id} > summary .panel-count`)?.textContent, 10) || 0), 0);
      const badge = b.querySelector(".rb-rail-n");
      badge.hidden = !n;
      badge.textContent = n > 99 ? "99+" : String(n);
    });
  }
  new MutationObserver(() => { if (!cq) cq = requestAnimationFrame(syncCounts); })
    .observe(sidebar, { subtree: true, childList: true, characterData: true });

  // ── położenie: panel stoi tuż za szyną, na wysokości obszaru roboczego ─────
  function place() {
    if (!on()) return;
    const r = rail.getBoundingClientRect();
    document.documentElement.style.setProperty("--rail-x", `${Math.round(r.right + 8)}px`);
    document.documentElement.style.setProperty("--rail-top", `${Math.round(r.top)}px`);
    const bar = document.querySelector(".status-bar");
    const bt = bar?.offsetParent ? bar.getBoundingClientRect().top : 0;
    document.documentElement.style.setProperty("--rail-bottom", bt ? `${Math.round(innerHeight - bt + 8)}px` : "16px");
  }
  if (typeof ResizeObserver === "function") new ResizeObserver(() => requestAnimationFrame(place)).observe(rail);
  window.addEventListener("resize", () => requestAnimationFrame(place), { passive: true });

  // start: bez panelu — dokument ma całe miejsce (panel otwiera szyna albo polecenie)
  place();
  syncCounts();
  if (on() && !pinned() && isSidebarOpen()) setSidebarOpen(false);
  sync();

  return { open, toggle, show, pinned, isOn: on };
})();
