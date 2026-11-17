// phone-bar.js — górny pasek telefonu jak w makiecie v4 (etap 7 wstążki; docs/WSTAZKA.md).
//
// Do 768 px, z otwartym plikiem: ← (zamknij do ekranu startowego) · Cofnij (w Edycji) · nazwa
// pliku i „N niezapisane · tryb” · przycisk trybu (menu: Edycja / Sugerowanie / Czytanie) · ⋯.
// Wszystko klika istniejące kontrolki (Zamknij z menu ⋯, #undoBtn, przyciski trybu, Zapisz), więc
// pytania o niezapisane zmiany, historia Cofnij i przełączanie trybu działają jak dotąd.

const dwbPhoneBar = (() => {
  const $ = (id) => document.getElementById(id);
  const back = $("phBackBtn"), undo = $("phUndoBtn"), status = $("phStatus"), modeBtn = $("phModeBtn");
  if (!back || !undo || !status || !modeBtn) return null;
  const readModeEl = $("readMode");
  const ICONS = {
    read: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>',
    edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
    suggest: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><path d="M12 6.5v4M10 8.5h4M10 13h4"/>',
  };
  const svg = (d, s = 18) => `<svg width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
  const NAME = { read: "modeRead", edit: "modeEdit", suggest: "modeSuggest" };
  const DESC = { edit: "phModeEditDesc", suggest: "phModeSuggestDesc", read: "phModeReadDesc" };

  function mode() {
    if (readModeEl?.checked) return "read";
    return typeof dwbTrack !== "undefined" && dwbTrack.isOn() ? "suggest" : "edit";
  }
  function sync() {
    const m = mode();
    document.body.classList.toggle("ph-reading", m === "read");
    modeBtn.dataset.mode = m;
    modeBtn.querySelector(".ph-mode-ico").innerHTML = svg(ICONS[m]);
    modeBtn.setAttribute("aria-label", t("phModeAria", { mode: t(NAME[m]) }));
    const saveBtn = $("heroSaveBtn");
    const dirty = !!saveBtn?.classList.contains("is-dirty");
    const n = dirty ? $("heroSaveCount")?.textContent || "1" : "";
    status.textContent = m === "read" ? t(NAME.read) : [n ? t("phUnsaved", { n }) : "", t(NAME[m])].filter(Boolean).join(" · ");
    const src = $("undoBtn");
    undo.disabled = !src || src.disabled;
  }
  let q = 0;
  const soon = () => { cancelAnimationFrame(q); q = requestAnimationFrame(sync); };

  back.addEventListener("click", () => $("closeDocBtn")?.click());
  undo.addEventListener("mousedown", (e) => e.preventDefault()); // kursor zostaje (pointerdown w WebKit zjadał klik)
  undo.addEventListener("click", () => $("undoBtn")?.click());
  $("phSaveMenuItem")?.addEventListener("click", () => {
    if (typeof appFrame !== "undefined") appFrame.setMenuOpen?.(false);
    $("heroSaveBtn")?.click();
  });
  modeBtn.addEventListener("click", () => {
    if (typeof composeUi === "undefined") return;
    const cur = mode();
    modeBtn.setAttribute("aria-expanded", "true");
    composeUi.openMenu(modeBtn, ["edit", "suggest", "read"].map((m) => ({
      label: t(NAME[m]), desc: t(DESC[m]), icon: svg(ICONS[m], 16), checked: m === cur,
      onPick: () => document.querySelector(`.mode-btn[data-mode="${m}"]`)?.click(),
    })));
    // menu zamyka się samo (compose-ui) — stan przycisku wraca przy najbliższej zmianie
    setTimeout(() => modeBtn.setAttribute("aria-expanded", "false"), 0);
  });

  readModeEl?.addEventListener("change", soon);
  document.addEventListener("dwb:track", soon);
  for (const id of ["heroSaveBtn", "heroSaveCount", "undoBtn"]) {
    const el = $(id);
    if (el) new MutationObserver(soon).observe(el, { attributes: true, childList: true, characterData: true, subtree: true });
  }
  if (typeof window.applyLanguage === "function") {
    const orig = window.applyLanguage;
    window.applyLanguage = function applyLanguagePhone(...args) { const r = orig.apply(this, args); sync(); return r; };
  }
  sync();
  return { sync };
})();
