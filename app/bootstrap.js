// App bootstrap: event wiring and startup.

if (typeof initLazyFeaturePanels === "function") initLazyFeaturePanels();

if (panelToggle) panelToggle.addEventListener("click", toggleSidebar);
if (panelHandle) panelHandle.addEventListener("click", toggleSidebar);
if (sidebarScrim) sidebarScrim.addEventListener("click", () => setSidebarOpen(false));

applyLanguage();
syncLangSwitchPill();
// Od 1024 px panel stoi obok dokumentu i pamięta ostatni wybór; węziej startuje schowany.
setSidebarOpen(appFrame.dockedInitialOpen());
// animacje ramy dopiero po pierwszym malowaniu — start bez „odjeżdżania” dokumentu
requestAnimationFrame(() => requestAnimationFrame(() => rootEl.classList.add("frame-ready")));
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
