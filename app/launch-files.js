// launch-files.js — otwieranie pliku „z zewnątrz”: Otwórz za pomocą / Udostępnij.
//
// Windows / macOS / ChromeOS (zainstalowana aplikacja w Chrome / Edge): manifest.json
// `file_handlers` wpisuje aplikację do „Otwórz za pomocą” dla .docx. Plik przychodzi przez
// launchQueue jako UCHWYT — więc „Zapisz” może zapisać prosto do oryginału (po zgodzie
// przeglądarki). `launch_handler: navigate-new` = każdy plik we własnym oknie, więc
// niezapisane zmiany w innym oknie nie przepadają.
//
// Android (zainstalowana aplikacja w Chrome): manifest `share_target` = pozycja w menu
// „Udostępnij”. System wysyła plik POST-em na ./share-target; sw.js go odkłada do cache
// i przekierowuje na ./?open=shared — tu go odbieramy. To KOPIA (bez zapisu do oryginału).
//
// iPhone / iPad i Safari: nie obsługują żadnego z tych mechanizmów — kod nic tam nie robi.

const SHARE_CACHE = "docs-wb-share";
const SHARE_KEY = "./__shared-file";

function initFileLaunch() {
  if ("launchQueue" in window && window.LaunchParams && "files" in window.LaunchParams.prototype) {
    window.launchQueue.setConsumer(async (params) => {
      const handle = params?.files?.find((h) => h.kind === "file");
      if (!handle) return;
      if (typeof confirmDiscardChanges === "function" && !confirmDiscardChanges()) return;
      try {
        const file = await handle.getFile();
        await ingestFile(file, { handle });
      } catch (e) {
        log(String(e?.message || e), "error");
        toast(t("openExternalFailed"), "error");
      }
    });
  }
  const open = new URLSearchParams(location.search).get("open");
  if (open) history.replaceState(null, "", location.pathname); // odświeżenie strony nie otwiera pliku drugi raz
  if (open === "shared") openSharedFile();
}

async function openSharedFile() {
  if (!("caches" in window)) return;
  try {
    const cache = await caches.open(SHARE_CACHE);
    const res = await cache.match(SHARE_KEY);
    if (!res) return;
    const name = decodeURIComponent(res.headers.get("X-File-Name") || "") || "dokument.docx";
    const blob = await res.blob();
    await cache.delete(SHARE_KEY);
    await ingestFile(new File([blob], name, { type: blob.type || "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
  } catch (e) {
    log(String(e?.message || e), "error");
    toast(t("openExternalFailed"), "error");
  }
}

// Po DOMContentLoaded: moduły (Cofnij, Recenzja, Formularz) podpinają się wtedy pod ingestFile.
if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", initFileLaunch);
else initFileLaunch();
