// DOCX preview rendering via docx-preview (lazy-loaded global).

// opts.pages: zawsze układ stron (podgląd wydruku rysuje osobny render także w Widoku mobilnym)
async function renderDocxPreview(bytes, container, opts = {}) {
  if (!container) return;
  container.replaceChildren();
  const wrapper = document.createElement("div");
  wrapper.className = "docx-preview-host";
  container.appendChild(wrapper);

  const mobileReflow = !opts.pages && typeof shouldUseMobileReflow === "function" && shouldUseMobileReflow();
  const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  await window.docx.renderAsync(ab, wrapper, null, {
    className: "docx",
    inWrapper: true,
    ignoreWidth: mobileReflow,
    ignoreHeight: mobileReflow,
    ignoreFonts: false,
    breakPages: !mobileReflow,
    renderHeaders: true,
    renderFooters: !mobileReflow,
    renderFootnotes: true,
    renderEndnotes: true,
  });
  addGenericFontFallbacks(wrapper); // brak kroju na urządzeniu → systemowy bezszeryfowy/szeryfowy z prawdziwym pogrubieniem
  applyRunDefaultsToParagraphs(wrapper); // pusty / nowy akapit: krój i rozmiar dokumentu, nie aplikacji
  fixDocxBulletRendering(wrapper);
  fixPageAnchoredDrawings(wrapper);
  applyWordLineMetrics(wrapper); // odstępy między wierszami jak w Wordzie (też granice stron)
  if (!mobileReflow) {
    layoutTabStops(wrapper); // tabulatory na pozycjach z akapitu (spis treści, formularze)
    // Czcionki osadzone w pliku (np. z PDF) docx-preview wczytuje z opóźnieniem — po ich
    // dojściu szerokości tekstu się zmieniają: liczymy tabulatory jeszcze raz.
    if (document.fonts) {
      const again = () => { if (wrapper.isConnected) layoutTabStops(wrapper); };
      document.fonts.ready.then(again).catch(() => {});
      let t = 0;
      const onDone = () => { clearTimeout(t); t = setTimeout(again, 60); };
      document.fonts.addEventListener("loadingdone", onDone);
      setTimeout(() => document.fonts.removeEventListener("loadingdone", onDone), 8000);
    }
  }
  if (mobileReflow && typeof applyMobileReflowLayout === "function") applyMobileReflowLayout(wrapper);
}
