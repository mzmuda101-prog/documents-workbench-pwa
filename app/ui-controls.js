// UI controls: theme, sidebar, file pickers, save, network, SW updates.

function syncLangSwitchPill() {
  const switchEl = document.getElementById("langSwitch");
  if (!switchEl) return;
  const active = switchEl.querySelector(".lang-button.is-active");
  if (!active) return;
  const switchRect = switchEl.getBoundingClientRect();
  const rect = active.getBoundingClientRect();
  switchEl.style.setProperty("--lang-pill-x", `${rect.left - switchRect.left}px`);
  switchEl.style.setProperty("--lang-pill-width", `${rect.width}px`);
  switchEl.classList.add("is-ready");
}

function initTheme() {
  const saved = localStorage.getItem(THEME_KEY);
  const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  const theme = saved || (prefersDark ? "dark" : "light");
  rootEl.setAttribute("data-theme", theme);
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

async function saveDocument() {
  if (!originalFileBytes) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  setLoading(true, t("savingFile"));
  try {
    const bytes = await buildDocumentForSave();
    if (fileHandle && typeof fileHandle.createWritable === "function") {
      if (!window.confirm(t("saveInPlaceWarn"))) return;
      try {
        const writable = await fileHandle.createWritable();
        await writable.write(bytes);
        await writable.close();
        pendingDocEdits = [];
        setDirtyState(false);
        toast(t("saveDone"), "success");
        return;
      } catch (_) {
        toast(t("savePermissionDenied"), "warning");
      }
    }
    await saveDocumentAs();
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
      setDirtyState(false);
      toast(t("saveDone"), "success");
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
  setDirtyState(false);
  toast(t("saveDone"), "success");
}

async function openFilePicker() {
  if (window.showOpenFilePicker) {
    try {
      const [handle] = await window.showOpenFilePicker({
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
if (searchHighlightBtn) searchHighlightBtn.addEventListener("click", runDocumentSearch);
if (searchQueryEl) {
  searchQueryEl.addEventListener("keydown", (e) => {
    if (e.key === "Enter") runDocumentSearch();
  });
}
if (zoomLevelEl) zoomLevelEl.addEventListener("input", applyZoom);
if (readModeEl) {
  readModeEl.addEventListener("change", () => {
    readOnlyMode = readModeEl.checked;
    docCanvasEl?.classList.toggle("read-only", readOnlyMode);
    const label = readModeEl.closest(".field")?.querySelector("span[data-i18n]");
    if (label) label.dataset.i18n = readOnlyMode ? "readModeOn" : "readModeOff";
    applyLanguage();
  });
  docCanvasEl?.classList.add("read-only");
}

document.querySelectorAll(".lang-button").forEach((btn) => {
  btn.addEventListener("click", () => setLanguage(btn.dataset.lang));
});

window.addEventListener("online", syncNetworkBadge);
window.addEventListener("offline", syncNetworkBadge);
window.addEventListener("resize", syncDocViewportHeight);

wireFileDrop();
initTheme();
syncNetworkBadge();
syncActionButtons();
registerServiceWorker();
