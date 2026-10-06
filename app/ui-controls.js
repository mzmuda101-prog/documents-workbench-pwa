// UI controls: theme, sidebar, file pickers, save, network, SW updates.

function syncLangSwitchPill() {
  const switchEl = document.getElementById("langSwitch");
  if (!switchEl) return;
  const active = switchEl.querySelector(".lang-button.is-active");
  if (!active) return;
  const switchRect = switchEl.getBoundingClientRect();
  const rect = active.getBoundingClientRect();
  const pad = 4;
  const x = Math.max(0, rect.left - switchRect.left - pad);
  const maxX = switchRect.width - rect.width - pad * 2;
  switchEl.style.setProperty("--lang-pill-x", `${Math.min(x, Math.max(0, maxX))}px`);
  switchEl.style.setProperty("--lang-pill-width", `${rect.width}px`);
  switchEl.classList.add("is-ready");
}

function requestCloseDocument() {
  if (!originalFileBytes) return;
  if (hasUnsavedChanges && !window.confirm(t("closeDocWarn"))) return;
  clearDocumentState();
  setStatus(t("noFile"));
  toast(t("docClosed"), "info");
}

function initTheme() {
  const saved = localStorage.getItem(THEME_KEY);
  const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  const theme = saved || (prefersDark ? "dark" : "light");
  rootEl.setAttribute("data-theme", theme);
}

function initIntroSplash() {
  const splash = document.getElementById("heroSplash");
  const vid = document.getElementById("introVideo");
  if (!splash) return;

  if (sessionStorage.getItem(INTRO_PLAYED_KEY)) {
    splash.style.display = "none";
    document.body.classList.remove("splashing");
    return;
  }

  document.body.classList.add("splashing");

  // Intro znika dopiero, gdy OBA warunki są spełnione: film się skończył i aplikacja pod spodem
  // jest gotowa (load + czcionki + 2 klatki). Pomiar 2026-10-01: gotowość ≤ 0,05 s po starcie
  // filmu nawet przy CPU ×6 i wolnym 4G. Film trwa INTRO_MS (prośba Mateusza: 2 s, nie krócej).
  // Od 2026-10-01 plik (mateusz-intro-2s.mp4) MA już 2 s przy 60 kl./s — dawniej 5,6 s grane
  // 2,8× szybciej: przeglądarka dekodowała ~168 kl./s i wyrzucała klatki nierówno (2–3 naraz),
  // co wyglądało jak przycinanie (zgłoszenie Mateusza). Tempo niżej zostaje jako zabezpieczenie
  // dla dłuższego pliku. Gdyby ładowanie trwało dłużej, zostaje ostatnia klatka aż do gotowości.
  const INTRO_MS = 2000;
  const introRate = () => (vid?.duration > 0 && Number.isFinite(vid.duration) ? Math.min(4, Math.max(1, (vid.duration * 1000) / INTRO_MS)) : 2.8);
  const INTRO_MAX_MS = 15000; // bezpiecznik: nigdy nie wisi na intro
  let videoDone = !vid;
  let appReady = false;
  let maxTimer = 0;

  const hideSplash = () => {
    if (!splash || splash.classList.contains("hide")) return;
    clearTimeout(maxTimer);
    splash.classList.add("hide");
    sessionStorage.setItem(INTRO_PLAYED_KEY, "true");
    setTimeout(() => {
      splash.style.display = "none";
      document.body.classList.remove("splashing");
    }, 700);
  };
  const maybeHide = () => { if (videoDone && appReady) hideSplash(); };
  const finishVideo = () => { videoDone = true; maybeHide(); };

  const markReady = () => {
    const fonts = document.fonts?.ready || Promise.resolve();
    fonts.catch(() => {}).then(() => requestAnimationFrame(() => requestAnimationFrame(() => { appReady = true; maybeHide(); })));
  };
  if (document.readyState === "complete") markReady();
  else window.addEventListener("load", markReady, { once: true });
  maxTimer = setTimeout(hideSplash, INTRO_MAX_MS);

  if (vid) {
    vid.addEventListener("ended", finishVideo, { once: true });
    try {
      vid.currentTime = 0;
      vid.muted = true;
      vid.defaultPlaybackRate = introRate();
      vid.playbackRate = introRate();
      // długość filmu znana dopiero z metadanych — wtedy tempo dokładnie na INTRO_MS
      vid.addEventListener("loadedmetadata", () => { vid.playbackRate = introRate(); }, { once: true });
      const playPromise = vid.play();
      // autoodtwarzanie zablokowane (np. tryb oszczędzania energii) — bez filmu, ale dalej
      // zakrywamy ładowanie aż do gotowości aplikacji
      if (playPromise && typeof playPromise.catch === "function") playPromise.catch(finishVideo);
    } catch {
      finishVideo();
    }
  }
}

function toggleTheme() {
  const next = rootEl.getAttribute("data-theme") === "dark" ? "light" : "dark";
  rootEl.setAttribute("data-theme", next);
  localStorage.setItem(THEME_KEY, next);
}

function syncNetworkBadge() {
  if (!networkBadge) return;
  const online = navigator.onLine;
  networkBadge.textContent = online ? t("online") : t("offline");
  networkBadge.classList.toggle("offline", !online);
}

function syncSidebarHandle() {
  if (!panelHandle) return;
  const sidebar = document.querySelector(".sidebar");
  const open = isSidebarOpen();
  panelHandle.textContent = "";
  panelHandle.setAttribute("aria-expanded", open ? "true" : "false");
  panelHandle.setAttribute("aria-label", open ? t("sidebarCloseAria") : t("sidebarOpenAria"));
  panelHandle.setAttribute("data-hint-pl", open ? "Schowaj sidebar" : "Wysuń sidebar");
  panelHandle.setAttribute("data-hint-en", open ? "Hide the sidebar" : "Open the sidebar");
  panelHandle.style.setProperty("--handle-open-label", `"${t("sidebarHandleLabel")}"`);
  panelHandle.style.setProperty("--handle-close-label", `"${t("sidebarHandleCloseLabel")}"`);
  panelHandle.style.setProperty("--handle-hide-panel-label", `"${t("sidebarHandleHideLabel")}"`);
  if (open && sidebar) {
    const rect = sidebar.getBoundingClientRect();
    const overlap = 8; // [EN] sit on the sidebar’s right edge, same as Sheet
    panelHandle.style.left = `${Math.max(8, Math.round(rect.right - overlap))}px`;
  } else {
    panelHandle.style.removeProperty("left");
  }
}

async function downloadBytes(bytes, name) {
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  // Nie od razu: Safari (iPhone — jedyna droga zapisu bez File System Access) potrafi
  // przerwać pobieranie, gdy adres zniknie w tym samym zadaniu co kliknięcie.
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

// Prawo zapisu do otwartego pliku. Tylko gdy mamy jego uchwyt (Chrome/Edge: „Otwórz” albo
// przeciągnięcie pliku) — inaczej null i „Zapisz” idzie w „Zapisz jako” (okno ZAPISU).
// Dawniej bez uchwytu otwierało się okno OTWIERANIA („wskaż plik do nadpisania”) — mylące:
// wyglądało jak otwieranie od nowa (zgłoszenie Mateusza, Windows, 2026-10-01).
async function ensureWriteAccess() {
  if (!fileHandle || typeof fileHandle.createWritable !== "function") return null;
  try {
    let perm = await fileHandle.queryPermission({ mode: "readwrite" });
    if (perm !== "granted") perm = await fileHandle.requestPermission({ mode: "readwrite" });
    return perm === "granted" ? fileHandle : null;
  } catch (_) {
    return null; // np. wygasła aktywacja użytkownika — wtedy zapis kopii
  }
}

let overwriteConfirmedFor = null; // „Nadpisać oryginalny plik?” — raz na plik, nie przy każdym zapisie

async function saveDocument() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  // Uprawnienie PRZED budowaniem pliku: requestPermission wymaga świeżego kliknięcia,
  // a przy dużym dokumencie budowanie trwa — przeglądarka odrzuciłaby pytanie.
  const handle = await ensureWriteAccess();
  if (!handle) {
    await saveDocumentAs();
    return;
  }
  if (overwriteConfirmedFor !== handle) {
    if (!window.confirm(t("saveInPlaceWarn", { name: handle.name || currentFileName || "" }))) return;
    overwriteConfirmedFor = handle;
  }
  setLoading(true, t("savingFile"));
  try {
    const bytes = await buildDocumentForSave();
    const writable = await handle.createWritable();
    await writable.write(bytes);
    await writable.close();
    pendingDocEdits = [];
    originalFileBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    await refreshInlineEditBaseline(originalFileBytes);
    setDirtyState(false);
    toast(t("saveDone"), "success");
    document.dispatchEvent(new CustomEvent("dwb:saved")); // strażnik spójności (self-check.js)
    if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
  } catch (e) {
    log(String(e.message || e), "error");
    toast(t("saveFailed"), "error");
  } finally {
    setLoading(false);
  }
}

async function saveDocumentAs() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  const bytes = await buildDocumentForSave();
  const base = (currentFileName || "document.docx").replace(/\.docx$/i, "");
  const suggested = `${base}_edited.docx`;

  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: suggested,
        types: [{
          description: "Word Document",
          accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] },
        }],
      });
      const writable = await handle.createWritable();
      await writable.write(bytes);
      await writable.close();
      fileHandle = handle;
      overwriteConfirmedFor = handle; // sam wybrał ten plik w oknie zapisu — kolejne „Zapisz” bez pytania
      currentFileName = handle.name || suggested;
      setFileUi(currentFileName, bytes.byteLength);
      originalFileBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      pendingDocEdits = [];
      await refreshInlineEditBaseline(originalFileBytes);
      setDirtyState(false);
      toast(t("saveDone"), "success");
      document.dispatchEvent(new CustomEvent("dwb:saved"));
    document.dispatchEvent(new CustomEvent("dwb:saved")); // strażnik spójności (self-check.js)
      if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
      return;
    } catch (e) {
      if (e && e.name === "AbortError") return;
    }
  }

  const nameRaw = window.prompt(t("saveAsPrompt"), suggested);
  if (!nameRaw) return;
  const name = nameRaw.toLowerCase().endsWith(".docx") ? nameRaw : `${nameRaw}.docx`;
  await downloadBytes(bytes, name);
  originalFileBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  currentFileName = name;
  pendingDocEdits = [];
  await refreshInlineEditBaseline(originalFileBytes);
  setDirtyState(false);
  // pobranie: przeglądarka (Safari/iPhone) jeszcze pyta „Pobrać?” — nie mówimy „zapisano”
  toast(t("saveDownloaded", { name }), "success");
  document.dispatchEvent(new CustomEvent("dwb:saved"));
  if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
}

// Nowy plik w miejsce dokumentu z niezapisanymi zmianami — zapytaj (jak przy zamykaniu
// i upuszczaniu pliku). Dawniej „Otwórz” i „Przykładowy dokument” podmieniały bez słowa.
function confirmDiscardChanges() {
  return !originalFileBytes || !hasUnsavedChanges || window.confirm(t("closeDocWarn"));
}

async function openFilePicker() {
  if (!confirmDiscardChanges()) return;
  if (window.showOpenFilePicker) {
    try {
      const handles = await window.showOpenFilePicker({
        mode: "readwrite",
        types: [{
          description: "Word / PDF / zdjęcie",
          accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"], "application/pdf": [".pdf"], "image/*": [".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif"] },
        }],
        multiple: true, // kilka plików = kilka kart (app/open-docs.js)
      });
      const items = [];
      for (const handle of handles) {
        const file = await handle.getFile();
        // PDF/zdjęcie zamienia się w nowy .docx — jego uchwyt nie może służyć do „Zapisz”
        items.push(detectFileType(file.name, file.type) !== "docx" ? { file } : { file, handle });
      }
      if (typeof openDocumentFiles === "function") await openDocumentFiles(items);
      else await ingestFile(items[0].file, items[0].handle ? { handle: items[0].handle } : {});
      return;
    } catch (e) {
      if (e && e.name === "AbortError") return;
    }
  }
  fileInput?.click();
}

function syncActionButtons() {
  const hasDoc = !!originalFileBytes;
  if (saveBtn) {
    saveBtn.disabled = !hasDoc;
    saveBtn.classList.toggle("primary", hasDoc && hasUnsavedChanges);
  }
  if (saveAsBtn) saveAsBtn.disabled = !hasDoc;
}

function wireFileDrop() {
  if (!dropZone) return;
  dropZone.addEventListener("click", (e) => {
    if (e.target.closest("#fileInput")) return;
    openFilePicker();
  });
  ["dragenter", "dragover"].forEach((ev) => {
    dropZone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropZone.classList.add("drag-over");
    });
  });
  ["dragleave", "drop"].forEach((ev) => {
    dropZone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropZone.classList.remove("drag-over");
    });
  });
  dropZone.addEventListener("drop", (e) => {
    const files = Array.from(e.dataTransfer?.files || []);
    if (files.length > 1 && typeof openDocumentFiles === "function") openDocumentFiles(files.map((file) => ({ file })));
    else if (files[0]) ingestDroppedFile(files[0], droppedFileHandle(e));
  });
}

// Upuszczony plik: w Chrome/Edge da się z niego wziąć uchwyt (jak z „Otwórz”), więc późniejsze
// „Zapisz” nadpisze ten plik. Wołać SYNCHRONICZNIE w zdarzeniu drop (potem dane znikają).
function droppedFileHandle(e) {
  const item = Array.from(e.dataTransfer?.items || []).find((i) => i.kind === "file");
  return item && typeof item.getAsFileSystemHandle === "function" ? item.getAsFileSystemHandle().catch(() => null) : Promise.resolve(null);
}
async function ingestDroppedFile(file, handlePromise) {
  const handle = await handlePromise;
  return ingestFile(file, handle?.kind === "file" ? { handle } : {});
}

// Przeładowanie (Aktualizuj / Odśwież aplikację) przy niezapisanych zmianach: najpierw pytamy.
// Na iPhonie/iPadzie Safari NIE pokazuje „Opuścić stronę?” — praca przepadała bez słowa.
// Po „OK” przeglądarka nie pyta drugi raz (beforeunload w bootstrap.js patrzy na tę flagę).
let reloadConfirmed = false;
function confirmReloadWithUnsaved() {
  if (!hasUnsavedChanges || reloadConfirmed) return true;
  if (!window.confirm(t("reloadUnsavedWarn"))) return false;
  reloadConfirmed = true;
  return true;
}

async function hardRefreshApp() {
  if (!confirmReloadWithUnsaved()) return;
  // Sygnał NATYCHMIAST po kliknięciu — czyszczenie cache i update() potrafią na
  // telefonie trwać sekundę i dłużej, a bez toastu klik wyglądał na niezłapany.
  toast(t("cacheRefresh"), "info");
  try {
    if ("serviceWorker" in navigator) {
      const registrations = await navigator.serviceWorker.getRegistrations();
      await Promise.all(registrations.map((registration) => registration.update().catch(() => {})));
      // Czekająca nowa wersja NIE włącza się sama przy przeładowaniu (stara dalej trzyma
      // stronę) — bez tego po przeładowaniu znów było „Aktualizuj”. Każemy jej przejąć
      // i chwilę czekamy, aż przejmie.
      const waiting = registrations.map((r) => r.waiting).filter(Boolean);
      if (waiting.length) {
        const took = new Promise((resolve) => navigator.serviceWorker.addEventListener("controllerchange", resolve, { once: true }));
        waiting.forEach((w) => w.postMessage({ type: "SKIP_WAITING" }));
        await Promise.race([took, new Promise((r) => setTimeout(r, 3000))]);
      }
    }
    if ("caches" in window) {
      const keys = await caches.keys();
      await Promise.all(keys.filter((key) => key.startsWith("docs-wb-")).map((key) => caches.delete(key).catch(() => false)));
    }
  } catch {
    // cache/SW nie dały się posprzątać — i tak przeładuj, toast już poszedł
  }
  window.location.reload();
}

function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) return;
  let waitingServiceWorker = null;
  let refreshingForUpdate = false;

  const showAppUpdate = (worker) => {
    waitingServiceWorker = worker;
    if (!appUpdateBtn) return;
    appUpdateBtn.classList.remove("hidden");
    toast(t("updateAvailable"), "info");
  };

  if (appUpdateBtn) {
    appUpdateBtn.addEventListener("click", () => {
      if (appUpdateBtn.classList.contains("is-busy")) return;
      if (!confirmReloadWithUnsaved()) return;
      // Najpierw WIDOCZNA zmiana stanu, dopiero potem robota — aktywacja nowego workera
      // i przeładowanie trwają od kilkuset ms do paru sekund.
      appUpdateBtn.classList.add("is-busy");
      appUpdateBtn.setAttribute("aria-busy", "true");
      appUpdateBtn.disabled = true;
      appUpdateBtn.textContent = t("refreshingApp");
      if (!waitingServiceWorker) {
        hardRefreshApp();
        return;
      }
      // Przeładowanie robi controllerchange (niżej) — dopiero gdy nowy SW przejął stronę.
      // Dawniej reload szedł od razu i strona wstawała jeszcze pod starym workerem.
      waitingServiceWorker.postMessage({ type: "SKIP_WAITING" });
      // Bezpiecznik: gdyby nowy worker nie przejął kontroli, nie zostawiaj wiecznego
      // „Odświeżam…" — po 4 s twarda ścieżka (czyszczenie cache + reload).
      window.setTimeout(() => {
        if (!refreshingForUpdate) hardRefreshApp();
      }, 4000);
    });
  }

  // Strona była JUŻ kontrolowana przy starcie → późniejszy controllerchange = aktualizacja
  // → przeładuj. Przy pierwszym wejściu controllerchange pochodzi z clients.claim() — to
  // nie aktualizacja, więc bez zbędnego mignięcia.
  const hadControllerAtStart = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (refreshingForUpdate || !hadControllerAtStart) return;
    // Aktualizację kliknięto w INNYM oknie aplikacji (np. drugi plik z „Otwórz za pomocą”),
    // a tu są niezapisane zmiany — nie przeładowuj sam; „Aktualizuj” zostaje na później.
    if (hasUnsavedChanges && !reloadConfirmed) {
      waitingServiceWorker = null; // nowa wersja już działa — „Aktualizuj” = zwykłe przeładowanie
      if (appUpdateBtn) {
        appUpdateBtn.classList.remove("hidden", "is-busy");
        appUpdateBtn.disabled = false;
        appUpdateBtn.removeAttribute("aria-busy");
        appUpdateBtn.textContent = t("updateApp");
      }
      toast(t("updateWaitsForSave"), "info");
      return;
    }
    refreshingForUpdate = true;
    window.location.reload();
  });

  // ?v= w adresie SW: każde wydanie = nowy adres = pewne wykrycie aktualizacji.
  navigator.serviceWorker.register(`./sw.js?v=${APP_BUILD_VERSION}`).then((registration) => {
    if (registration.waiting && navigator.serviceWorker.controller) showAppUpdate(registration.waiting);
    registration.addEventListener("updatefound", () => {
      const worker = registration.installing;
      if (!worker) return;
      worker.addEventListener("statechange", () => {
        if (worker.state === "installed" && navigator.serviceWorker.controller) showAppUpdate(worker);
      });
    });
    // PWA wznowiona z tła / trzymana otwarta: sprawdzaj nową wersję przy powrocie na
    // pierwszy plan i co 30 min — inaczej z ikony siedzi się na starym buildzie.
    const checkForUpdate = () => registration.update().catch(() => {});
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") checkForUpdate();
    });
    window.setInterval(checkForUpdate, 30 * 60 * 1000);
  }).catch(() => {});
}

if (themeToggle) themeToggle.addEventListener("click", toggleTheme);
if (brandRefresh) brandRefresh.addEventListener("click", hardRefreshApp);
if (closeDocBtn) closeDocBtn.addEventListener("click", requestCloseDocument);
const closeDocPanelBtn = document.getElementById("closeDocPanelBtn");
if (closeDocPanelBtn) closeDocPanelBtn.addEventListener("click", requestCloseDocument);
if (loadBtn) loadBtn.addEventListener("click", openFilePicker);
if (loadSampleBtn) loadSampleBtn.addEventListener("click", () => loadSampleDocument("przewodnik"));
// „Zdjęcie dokumentu”: aparat na telefonie/tablecie (capture), na komputerze wybór zdjęć.
// Kilka zdjęć naraz = jeden dokument z kilkoma stronami (open-docs.js groupPhotos).
const photoDocBtn = document.getElementById("photoDocBtn");
const photoInput = document.getElementById("photoInput");
if (photoDocBtn && photoInput) {
  photoDocBtn.addEventListener("click", () => {
    if (typeof confirmDiscardChanges === "function" && !confirmDiscardChanges()) return;
    photoInput.click();
  });
  photoInput.addEventListener("change", () => {
    const files = Array.from(photoInput.files || []);
    photoInput.value = "";
    if (!files.length) return;
    if (typeof openDocumentFiles === "function") openDocumentFiles(files.map((file) => ({ file })));
    else ingestFile(files[0]);
  });
}
if (saveBtn) saveBtn.addEventListener("click", saveDocument);
if (saveAsBtn) saveAsBtn.addEventListener("click", saveDocumentAs);
if (fileInput) {
  fileInput.addEventListener("change", () => {
    const files = Array.from(fileInput.files || []);
    fileInput.value = "";
    if (files.length > 1 && typeof openDocumentFiles === "function") openDocumentFiles(files.map((file) => ({ file })));
    else if (files[0]) (typeof openDocumentFiles === "function" ? openDocumentFiles([{ file: files[0] }]) : ingestFile(files[0]));
  });
}
const frWorkbenchActive = !!document.getElementById("frScanBtn");
if (searchHighlightBtn && !frWorkbenchActive) {
  searchHighlightBtn.addEventListener("click", runDocumentSearch);
}
if (searchQueryEl && !frWorkbenchActive) {
  searchQueryEl.addEventListener("keydown", (e) => {
    if (e.key === "Enter") runDocumentSearch();
  });
}
if (zoomLevelEl) zoomLevelEl.addEventListener("input", typeof onZoomSliderInput === "function" ? onZoomSliderInput : applyZoom);
if (readModeEl) {
  readModeEl.addEventListener("change", () => {
    readOnlyMode = readModeEl.checked;
    keepDocTextInPlace(syncInlineEditMode); // pasek Edycji znika / wraca nad dokumentem
  });
  docCanvasEl?.classList.add("read-only");
}

document.querySelectorAll(".lang-button").forEach((btn) => {
  btn.addEventListener("click", () => setLanguage(btn.dataset.lang));
});

window.addEventListener("online", syncNetworkBadge);
window.addEventListener("offline", syncNetworkBadge);
window.addEventListener("resize", () => {
  syncDocViewportHeight();
  syncLangSwitchPill();
  if (typeof onViewportChange === "function") onViewportChange();
  if (typeof syncSidebarHandle === "function") syncSidebarHandle();
});

wireFileDrop();
initIntroSplash();
initTheme();
syncNetworkBadge();
syncActionButtons();
registerServiceWorker();
