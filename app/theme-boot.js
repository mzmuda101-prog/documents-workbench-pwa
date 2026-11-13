// theme-boot.js — motyw ustawiony ZANIM strona się narysuje (bez mignięcia jasnym przy ciemnym).
// Ładowany synchronicznie w <head> (CSP nie pozwala na skrypt wpisany w HTML). Brak klucza = motyw
// systemowy; reszta (nasłuch zmiany systemu, wybór w menu ⋯) — ui-controls.js.
(() => {
  let pref = null;
  try { pref = localStorage.getItem("documents-workbench-theme"); } catch (_) {}
  const dark = pref === "dark" || (pref !== "light" && !!(window.matchMedia && matchMedia("(prefers-color-scheme: dark)").matches));
  document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
  // motyw wybrany ręcznie wygrywa z systemem także w pasku stanu (kolory = --bg z app.css)
  document.querySelectorAll('meta[name="theme-color"]').forEach((m) => m.setAttribute("content", dark ? "#0b1018" : "#e8ecf4"));
})();
