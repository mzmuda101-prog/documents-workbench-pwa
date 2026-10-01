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

// Biblioteki DOCX (docx-preview + JSZip) dociągamy w wolnej chwili po starcie — pierwsze
// otwarcie pliku nie czeka wtedy na ich pobranie i parsowanie (paczka F). Z oszczędzaniem
// danych (Save-Data) zostaje ładowanie przy pierwszym użyciu.
if (!navigator.connection?.saveData) {
  (window.requestIdleCallback ? (fn) => requestIdleCallback(fn, { timeout: 4000 }) : (fn) => setTimeout(fn, 2000))(() => {
    if (!originalFileBytes) ensureDocLibs(false).catch(() => {});
  });
}

// ?sample=nazwa — otwiera docs/samples/nazwa.docx od razu (symulator iOS, testy, pokaz).
// Na telefonie nie ma jak wskazać pliku z dysku Maca, a przykład z nagłówkami jest potrzebny.
{
  const sampleParam = new URLSearchParams(location.search).get("sample");
  if (sampleParam) loadSampleDocument(sampleParam);
}

window.addEventListener("beforeunload", (e) => {
  if (!hasUnsavedChanges || reloadConfirmed) return; // przy „Aktualizuj” już zapytaliśmy
  e.preventDefault();
  e.returnValue = "";
});
