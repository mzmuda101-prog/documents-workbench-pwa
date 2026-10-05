// DOCX preview rendering via docx-preview (lazy-loaded global).

// Aptos (domyślny krój Word 365): na urządzeniu zwykle go nie ma — wtedy zamiennik z app.css
// (Arial/Arimo przeskalowany do szerokości Aptosa, blok doc-fonts). Gdy prawdziwy Aptos JEST
// (Office udostępnił go systemowi), dokładamy go jako ostatnią regułę rodziny „Aptos” — wygrywa
// z zamiennikiem, a size-adjust go nie dotyczy.
(function useLocalAptos() {
  if (typeof FontFace !== "function" || !document.fonts?.add) return;
  [["Aptos", "Aptos", "normal", "400"], ["Aptos Bold", "Aptos-Bold", "normal", "700"], ["Aptos Italic", "Aptos-Italic", "italic", "400"], ["Aptos Bold Italic", "Aptos-BoldItalic", "italic", "700"]]
    .forEach(([full, ps, style, weight]) => {
      new FontFace("Aptos", `local("${full}"), local("${ps}")`, { style, weight }).load().then((f) => document.fonts.add(f)).catch(() => {});
    });
})();

// Plik bez kroju (brak w:rFonts w docDefaults): tekst dostaje krój zastępczy kartki jak w Wordzie
// 365 — Aptos (app.css). Font dociągany dopiero przy pierwszym użyciu przestawiał tekst PO
// pomiarach — granice stron i powrót na to samo miejsce (zmiana widoku, gest szczypania)
// liczyły się na starych wymiarach. Wczytujemy go od razu (lokalnie, z limitem czasu).
async function loadFallbackDocFont(wrapper) {
  const span = [...wrapper.querySelectorAll("section.docx span")].find((s) => s.textContent.trim());
  if (!span || !document.fonts?.load || !/^"?Aptos/.test(getComputedStyle(span).fontFamily)) return;
  const faces = ["normal 400", "normal 700", "italic 400", "italic 700"].map((v) => document.fonts.load(`${v} 16px Aptos`, "AaĄąŻż").catch(() => null));
  await Promise.race([Promise.all(faces), new Promise((r) => setTimeout(r, 1500))]);
}

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
  await loadFallbackDocFont(wrapper); // plik bez kroju: zamiennik wczytany przed pomiarami
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
