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

  const hideSplash = () => {
    if (!splash || splash.classList.contains("hide")) return;
    splash.classList.add("hide");
    sessionStorage.setItem(INTRO_PLAYED_KEY, "true");
    setTimeout(() => {
      splash.style.display = "none";
      document.body.classList.remove("splashing");
    }, 700);
  };

  if (vid) {
    try {
      vid.currentTime = 0;
      vid.muted = true;
      vid.playbackRate = 1.5;
      const playPromise = vid.play();
      if (playPromise && typeof playPromise.catch === "function") playPromise.catch(() => hideSplash());
    } catch {
      hideSplash();
    }
    const fallback = setTimeout(hideSplash, 10000);
    vid.addEventListener("ended", () => { clearTimeout(fallback); hideSplash(); }, { once: true });
  } else {
    setTimeout(hideSplash, 6000);
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
  if (!sidebar) return;
  const open = isSidebarOpen();
  if (open) {
    const rect = sidebar.getBoundingClientRect();
    panelHandle.style.left = `${Math.max(8, rect.right + 6)}px`;
  } else {
    panelHandle.style.left = "";
  }
  if (panelToggle) {
    const handleVisible = getComputedStyle(panelHandle).display !== "none";
    panelToggle.hidden = handleVisible;
    panelToggle.setAttribute("aria-hidden", handleVisible ? "true" : "false");
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
  URL.revokeObjectURL(url);
}

async function ensureWriteAccess() {
  if (fileHandle && typeof fileHandle.createWritable === "function") {
    let perm = await fileHandle.queryPermission({ mode: "readwrite" });
    if (perm !== "granted") perm = await fileHandle.requestPermission({ mode: "readwrite" });
    if (perm === "granted") return fileHandle;
  }
  if (!window.showOpenFilePicker) return null;
  toast(t("savePickFileHint"), "info");
  try {
    const [handle] = await window.showOpenFilePicker({
      mode: "readwrite",
      types: [{
        description: "Word Document",
        accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] },
      }],
      multiple: false,
    });
    const file = await handle.getFile();
    if (currentFileName && file.name && file.name !== currentFileName) {
      if (!window.confirm(t("saveDifferentFileWarn", { name: file.name }))) return null;
    }
    fileHandle = handle;
    if (file.name) {
      currentFileName = file.name;
      setFileUi(currentFileName, originalFileBytes?.byteLength || file.size);
    }
    return handle;
  } catch (e) {
    if (e && e.name === "AbortError") return null;
    throw e;
  }
}

async function saveDocument() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  setLoading(true, t("savingFile"));
  try {
    const bytes = await buildDocumentForSave();
    const handle = await ensureWriteAccess();
    if (!handle) {
      toast(t("saveCancelled"), "info");
      return;
    }
    if (!window.confirm(t("saveInPlaceWarn"))) return;
    const writable = await handle.createWritable();
    await writable.write(bytes);
    await writable.close();
    pendingDocEdits = [];
    originalFileBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    await refreshInlineEditBaseline(originalFileBytes);
    setDirtyState(false);
    toast(t("saveDone"), "success");
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
      currentFileName = handle.name || suggested;
      setFileUi(currentFileName, bytes.byteLength);
      originalFileBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      pendingDocEdits = [];
      await refreshInlineEditBaseline(originalFileBytes);
      setDirtyState(false);
      toast(t("saveDone"), "success");
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
  toast(t("saveDone"), "success");
  if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
}

async function openFilePicker() {
  if (window.showOpenFilePicker) {
    try {
      const [handle] = await window.showOpenFilePicker({
        mode: "readwrite",
        types: [{
          description: "Word Document",
          accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] },
        }],
        multiple: false,
      });
      const file = await handle.getFile();
      await ingestFile(file, { handle });
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
    const file = e.dataTransfer?.files?.[0];
    if (file) ingestFile(file);
  });
}

function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) return;
  navigator.serviceWorker.register("./sw.js").then((reg) => {
    reg.addEventListener("updatefound", () => {
      const worker = reg.installing;
      if (!worker) return;
      worker.addEventListener("statechange", () => {
        if (worker.state === "installed" && navigator.serviceWorker.controller && appUpdateBtn) {
          appUpdateBtn.classList.remove("hidden");
        }
      });
    });
  }).catch(() => {});
}

if (themeToggle) themeToggle.addEventListener("click", toggleTheme);
if (brandRefresh) brandRefresh.addEventListener("click", () => location.reload());
if (closeDocBtn) closeDocBtn.addEventListener("click", requestCloseDocument);
const closeDocPanelBtn = document.getElementById("closeDocPanelBtn");
if (closeDocPanelBtn) closeDocPanelBtn.addEventListener("click", requestCloseDocument);
if (appUpdateBtn) {
  appUpdateBtn.addEventListener("click", () => {
    navigator.serviceWorker.getRegistration().then((reg) => {
      reg?.waiting?.postMessage({ type: "SKIP_WAITING" });
      location.reload();
    });
  });
}
if (loadBtn) loadBtn.addEventListener("click", openFilePicker);
if (loadSampleBtn) loadSampleBtn.addEventListener("click", loadSampleDocument);
if (saveBtn) saveBtn.addEventListener("click", saveDocument);
if (saveAsBtn) saveAsBtn.addEventListener("click", saveDocumentAs);
if (fileInput) {
  fileInput.addEventListener("change", () => {
    const file = fileInput.files?.[0];
    if (file) ingestFile(file);
    fileInput.value = "";
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
    const label = readModeEl.closest(".field")?.querySelector("span[data-i18n]");
    if (label) label.dataset.i18n = readOnlyMode ? "readModeOn" : "readModeOff";
    applyLanguage();
    syncInlineEditMode();
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
});

wireFileDrop();
initIntroSplash();
initTheme();
syncNetworkBadge();
syncActionButtons();
registerServiceWorker();
