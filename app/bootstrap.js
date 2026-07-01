// App bootstrap: event wiring and startup.

if (typeof initLazyFeaturePanels === "function") initLazyFeaturePanels();

if (panelToggle) panelToggle.addEventListener("click", toggleSidebar);
if (panelHandle) panelHandle.addEventListener("click", toggleSidebar);
if (sidebarScrim) sidebarScrim.addEventListener("click", () => setSidebarOpen(false));

applyLanguage();
syncLangSwitchPill();
setSidebarOpen(window.matchMedia("(min-width: 1100px)").matches);
if (typeof initMobileDocZoom === "function") initMobileDocZoom();
syncDocViewportHeight();
applyZoom();
syncDocumentShellClass();
setDirtyState(false);
setStatus(t("noFile"));

window.addEventListener("beforeunload", (e) => {
  if (!hasUnsavedChanges) return;
  e.preventDefault();
  e.returnValue = "";
});

if ("serviceWorker" in navigator) {
  let refreshing = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (refreshing) return;
    refreshing = true;
  });
}
