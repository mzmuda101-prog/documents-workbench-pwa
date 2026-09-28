// =============================================================================
// *cursor-hint.js — podążający za kursorem dymek z podpowiedzią
// =============================================================================
//
// *WSZYSTKIE ATRYBUTY HTML (data-*) — ściągawka
// -----------------------------------------------------------------------------
//
// data-hint
//   Główny atrybut — jego obecność na elemencie włącza hint.
//   Trzy warianty użycia:
//
//   a) data-hint="Zapisz plik"
//      Tekst podany wprost. Używaj gdy hint nie musi się tłumaczyć.
//
//   b) data-hint (bez wartości) lub data-hint=""
//      Pusty atrybut = "chcę hint, ale tekst pochodzi skądinąd".
//      Wyświetli fallbackHint ("Kliknij" / "Click") chyba że JS
//      dynamicznie wpisze tekst przez setAttribute("data-hint", "..."),
//      np. jak robi to language.js przez setHint().
//
//   c) data-hint + data-hint-pl + data-hint-en  ← PREFEROWANE dla i18n
//      Samo data-hint jako "znacznik aktywacji", a teksty w osobnych
//      atrybutach językowych (patrz niżej). Wtedy data-hint może być puste.
//
// data-hint-pl="..."
//   Tekst hinta po polsku. Używany gdy lang aplikacji = "pl".
//   Jeśli istnieje, ma pierwszeństwo nad data-hint.
//   Przykład: data-hint-pl="Zmień motyw jasny / ciemny"
//
// data-hint-en="..."
//   Tekst hinta po angielsku. Używany gdy lang aplikacji = "en".
//   Jeśli istnieje, ma pierwszeństwo nad data-hint.
//   Przykład: data-hint-en="Switch light / dark theme"
//
//   UWAGA: data-hint-pl i data-hint-en działają niezależnie —
//   możesz podać tylko jeden z nich, drugi nie jest wymagany.
//   Jeśli brakuje wersji dla aktualnego języka, system spada
//   na data-hint, a potem na fallback.
//
// data-hint-delay="1.1"
//   Opóźnienie pojawienia się hinta przy kursorem myszy, w sekundach.
//   Timer startuje gdy mysz wejdzie na element i resetuje się gdy
//   mysz ruszy się o więcej niż ~4px (MOVE_THRESHOLD).
//   Hint pojawia się dopiero po zatrzymaniu myszy na zadany czas.
//   Wartość: liczba dziesiętna, akceptuje przecinek i kropkę.
//   Domyślnie: 0 (pojawia się od razu).
//   Przykład: data-hint-delay="1.1"  ← pojawi się po 1.1 sekundy bezruchu
//
// data-hint-touch[="sekundy"]
//   Włącza hint dla urządzeń dotykowych (przytrzymanie palca).
//   Domyślnie hint działa tylko dla myszy — dodaj ten atrybut żeby
//   włączyć obsługę dotyku. Czas przytrzymania i delay myszy są NIEZALEŻNE.
//   Przy aktywnym timerze dotykowym blokowane jest systemowe context menu
//   (iOS "Kopiuj / Szukaj", Android long-press menu) — brak przerywania hintu.
//
//   Warianty:
//     data-hint-touch            ← dotyk włączony, domyślny delay (650ms)
//     data-hint-touch="on"       ← to samo co powyżej
//     data-hint-touch="0.8"      ← dotyk włączony, delay 800ms
//     data-hint-touch="1.2"      ← dotyk włączony, delay 1200ms
//
// data-hint-tap  (opt-in, działa razem z data-hint-touch)
//   Zmienia model DOTYKU na „tap vs przytrzymanie" — dla elementów, które NIE mają
//   własnej akcji na tapnięcie (np. liczba/wynik, nie przycisk):
//     • TAP (krótkie tapnięcie, <500ms)  → hint pojawia się na chwilę i SAM znika
//       (czas: data-hint-duration jeśli ustawione, inaczej ~2.6s; ładnie z data-hint-fade).
//     • PRZYTRZYMANIE (≥~280ms trzymania) → hint pokazany TAK DŁUGO jak trzymasz palec,
//       znika dopiero po puszczeniu. Ruch palcem (scroll) anuluje.
//   Bez tego atrybutu dotyk działa po staremu („przytrzymaj, by pokazać; puść = schowaj"),
//   więc przyciski z akcją się nie zmieniają — to dlatego jest opt-in.
//   WAŻNE (iOS/Safari): na kontenerze daj `-webkit-user-select:none; user-select:none;
//   -webkit-touch-callout:none;` — inaczej długie przytrzymanie zaznacza tekst / pokazuje
//   lupę systemową RÓWNOLEGLE z hintem. (Patrz .monthly-summary w app.css.)
//   Przykład: <span data-hint data-hint-pl="..." data-hint-touch="on" data-hint-tap data-hint-fade>
//
// data-hint-class="nazwa-klasy"
//   Dodatkowa klasa CSS doklejana do dymka gdy jest widoczny.
//   Przydatne do stylowania konkretnych hintów inaczej niż reszta,
//   np. inny kolor, rozmiar, wariant. Klasa jest usuwana gdy hint znika.
//   Przykład: data-hint-class="smaller"  ← jest już taka klasa w app.css
//
// data-hint-duration="2.5"
//   Czas po którym hint automatycznie znika, w sekundach.
//   Odliczanie startuje w momencie gdy hint staje się widoczny (po delay).
//   Mysz opuszczająca element nadal chowa hint normalnie — duration to
//   tylko górny limit widoczności gdy kursor stoi w miejscu.
//   Wartość: liczba dziesiętna, akceptuje przecinek i kropkę.
//   Domyślnie: brak (hint znika tylko gdy mysz opuści element).
//   Przykład: data-hint-duration="3"  ← hint zniknie po 3 sekundach
//
// data-hint-fade[="sekundy"]
//   Włącza powolne zanikanie przez mgłę (blur + opacity).
//   Działa przy KAŻDYM ukryciu hinta — zarówno przez auto-hide (data-hint-duration)
//   jak i przy normalnym opuszczeniu elementu myszą.
//   Bez tego atrybutu hint znika szybko (140ms) — domyślne zachowanie.
//
//   Warianty:
//     data-hint-fade              ← domyślny czas fade (460ms)
//     data-hint-fade="0.5"        ← fade przez 0.5 sekundy
//     data-hint-fade="1.2"        ← fade przez 1.2 sekundy
//
//   Wartość: liczba dziesiętna w sekundach (>0). Brak wartości lub 0 = 460ms.
//
// -----------------------------------------------------------------------------
// *SEPARATOR LINII W TEKŚCIE HINTA
// -----------------------------------------------------------------------------
//
//   /|  (ukośnik + pipe) — ręczny podział na nową linię w treści hinta.
//   Możesz go użyć w każdym atrybucie tekstowym (data-hint, data-hint-pl itd.).
//   Przykład: data-hint="Filtruj dane /| i sortuj wyniki"
//             wyświetli się jako dwie linie:
//               Filtruj dane
//               i sortuj wyniki
//
// -----------------------------------------------------------------------------
// *KOMPLETNY PRZYKŁAD — element z PEŁNYM zestawem atrybutów
// -----------------------------------------------------------------------------
//
//   <button
// -    data-hint
// -    data-hint-pl="Pobierz kopię pliku /| jako XLSX"
// -    data-hint-en="Download a copy /| as XLSX"
// -    data-hint-delay="1.1"
// -    data-hint-touch="on"
// -    data-hint-class="smaller"
//   >
//     Zapisz
//   </button>
//
// -----------------------------------------------------------------------------
// *CSS — customizacja offsetu dymka przez zmienne
// -----------------------------------------------------------------------------
//
//   --hint-offset-x  (domyślnie 22px) — poziome przesunięcie od kursora
//   --hint-offset-y  (domyślnie 18px) — pionowe przesunięcie od kursora
//
//   Ustaw na elemencie #cursorHint lub przez klasę z data-hint-class:
//   .cursor-hint.smaller { --hint-offset-x: 18; --hint-offset-y: 22; }
//
// =============================================================================

window.MateuszCursorHint = (() => {
  function createCursorHintController({ cursorHint, prefersReducedMotion = false, getFallbackHint = () => "" }) {
    // Aktualna pozycja dymka (interpolowana, zmienia się co klatkę animacji)
    let cursorHintX = -999;
    let cursorHintY = -999;
    // Docelowa pozycja dymka (ustawiana od razu przy ruchu myszy)
    let cursorHintTargetX = -999;
    let cursorHintTargetY = -999;
    // ID klatki requestAnimationFrame — null gdy animacja nie jest aktywna
    let cursorHintFrame = null;
    // ID setTimeout dla opóźnionego pokazania hinta — null gdy brak aktywnego timera
    let cursorHintTimer = null;
    // ID setTimeout dla automatycznego ukrycia hinta — null gdy brak aktywnego timera
    let autoHideTimer = null;
    // ID setTimeout dla opóźnionego sprzątania po powolnym fade-out — null gdy brak
    let fadingOutTimer = null;
    // Element na którym aktualnie czeka lub pokazuje hint
    let activeHintEl = null;
    // Typ wskaźnika który aktywował hint ("mouse" | "touch" | "pen")
    let activePointerType = "";

    // Dotyk w trybie opt-in `data-hint-tap`: TAP = zerknięcie (pokaż na chwilę, sam znika),
    // PRZYTRZYMANIE = pokazuj póki trzymasz (znika po puszczeniu). Bez data-hint-tap dotyk
    // działa po staremu (przytrzymaj-by-pokazać) — przyciski z akcją się nie zmieniają.
    let touchHoldTimer = null;     // timer wykrycia przytrzymania
    let touchDownTime = 0;         // moment przyciśnięcia palca
    let touchHeldShown = false;    // czy hint pokazany w trybie przytrzymania
    let touchMoved = false;        // czy palec ruszył (scroll → anuluj)
    let touchStartX = 0;
    let touchStartY = 0;
    const TOUCH_HOLD_MS = 280;     // po tylu ms trzymania pokazujemy „na trzymanie"
    const TOUCH_TAP_MAX = 500;     // krótsze przyciśnięcie = tap (zerknięcie)
    const TOUCH_MOVE_CANCEL = 12;  // ruch > tylu px przed pokazaniem = scroll, anuluj
    const TOUCH_PEEK_MS = 2600;    // domyślny czas „zerknięcia" (override: data-hint-duration)
    const nowMs = () => (window.performance && performance.now ? performance.now() : Date.now());

    // Sprawdza czy element jawnie zezwala na hint dotykowy przez data-hint-touch.
    // Akceptuje: brak wartości / "on" / "true" / "1" / "yes" / liczba dziesiętna (delay w sek.)
    function allowsTouchHint(el) {
      if (el.dataset.hintTouch === undefined) return false;
      const value = String(el.dataset.hintTouch).toLowerCase();
      if (value === "" || value === "on" || value === "true" || value === "1" || value === "yes") return true;
      const num = parseFloat(value.replace(",", "."));
      return Number.isFinite(num) && num >= 0;
    }

    // Zwraca true gdy hint ma być zablokowany — np. brak elementu dymka w DOM,
    // użytkownik preferuje reduced motion, albo dotyk bez jawnego zezwolenia,
    // albo urządzenie z grubym wskaźnikiem (tablet) bez myszy
    function isDisabled(el, pointerType = "") {
      if (!cursorHint || prefersReducedMotion) return true;
      if (pointerType === "touch") return !allowsTouchHint(el);
      // (pointer: coarse) = ekran dotykowy bez myszy — tam hint myszowy nie ma sensu
      return pointerType !== "mouse" && window.matchMedia("(pointer: coarse)").matches;
    }

    // Czyta lang z <html lang="..."> i normalizuje do "pl" lub "en"
    function getCurrentLang() {
      return (document.documentElement.lang || "pl").toLowerCase().startsWith("en") ? "en" : "pl";
    }

    // Zwraca tekst do wyświetlenia w hincie dla danego elementu.
    // Kolejność priorytetów: data-hint-pl/en → data-hint → fallback globalny
    function getHintText(el) {
      const lang = getCurrentLang();
      const langKey = lang === "en" ? "hintEn" : "hintPl";

      // data-hint-pl / data-hint-en — zawsze mają pierwszeństwo jeśli istnieją
      if (el.dataset[langKey] !== undefined) return el.dataset[langKey];

      // data-hint="tekst" — użyj tekstu; data-hint="" lub samo data-hint — fallback
      if (el.dataset.hint !== undefined && el.dataset.hint !== "") return el.dataset.hint;

      // brak wartości lub pusty atrybut → domyślny hint (np. "Kliknij")
      return getFallbackHint() || "";
    }

    // Zwraca czas auto-hide w ms dla danego elementu.
    // 0 = brak auto-hide (domyślne zachowanie).
    function getHintDurationMs(el) {
      const raw = el.dataset.hintDuration || "";
      const parsed = parseFloat(raw.replace(",", "."));
      if (Number.isFinite(parsed) && parsed > 0) return parsed * 1000;
      return 0;
    }

    // Zwraca opóźnienie w ms dla danego elementu i typu wskaźnika.
    // Mysz i dotyk mają osobne atrybuty i osobne wartości domyślne.
    function getHintDelayMs(el, pointerType = "") {
      if (pointerType === "touch") {
        // Delay dotyku pochodzi z wartości data-hint-touch (np. "0.8" → 800ms).
        // data-hint-delay nie ma wpływu na dotyk — osobne atrybuty, osobne timery.
        const rawTouch = el.dataset.hintTouch || "";
        const parsedTouch = parseFloat(rawTouch.replace(",", "."));
        if (Number.isFinite(parsedTouch) && parsedTouch >= 0) return Math.round(parsedTouch * 1000);
        return 650; // domyślny delay gdy brak wartości liczbowej
      }
      // Mysz — czyta data-hint-delay, domyślnie 0 (od razu)
      const rawDelay = el.dataset.hintDelay || "";
      const parsedDelay = parseFloat(rawDelay.replace(",", "."));
      if (Number.isFinite(parsedDelay) && parsedDelay >= 0) return parsedDelay * 1000;
      return 0;
    }

    // Oblicza docelową pozycję dymka tak żeby nie wychodził poza viewport.
    // Odczytuje rozmiar dymka przez getBoundingClientRect (działa bo dymek
    // jest już w DOM, tylko niewidoczny — opacity/scale przez CSS).
    // Zwraca też originX/Y dla transformOrigin żeby animacja scale
    // zawsze wychodziła z narożnika bliższego kursorowi.
    function computeHintPosition(x, y) {
      if (!cursorHint) return { tx: x + 22, ty: y - 18, originX: "left", originY: "bottom" };

      // Offset dymka od kursora — można nadpisać przez CSS custom properties
      const style = getComputedStyle(cursorHint);
      const parsedOffsetX = parseInt(style.getPropertyValue("--hint-offset-x"), 10);
      const parsedOffsetY = parseInt(style.getPropertyValue("--hint-offset-y"), 10);
      const offsetX = Number.isNaN(parsedOffsetX) ? 22 : parsedOffsetX;
      const offsetY = Number.isNaN(parsedOffsetY) ? 18 : parsedOffsetY;

      const W = window.innerWidth;
      const H = window.innerHeight;
      const MARGIN = 8; // minimalna odległość od krawędzi ekranu

      // Rozmiar dymka znany dopiero po renderze — dlatego liczymy pozycję
      // tutaj, a nie przed pokazaniem
      const rect = cursorHint.getBoundingClientRect();
      const hintW = rect.width || 0;
      const hintH = rect.height || 0;

      // Domyślna pozycja: prawo-góra względem kursora
      let tx = x + offsetX;
      let ty = y - offsetY - hintH;
      let originX = "left";
      let originY = "bottom";

      // Wychodzi poza prawą krawędź → przesuń w lewo od kursora
      if (tx + hintW + MARGIN > W) {
        tx = x - offsetX - hintW;
        originX = "right";
      }

      // Wychodzi poza górną krawędź → pokaż pod kursorem
      if (ty < MARGIN) {
        ty = y + offsetY;
        originY = "top";
      }

      // Zabezpieczenia przed wyjściem poza lewą i dolną krawędź
      // (rzadkie, ale możliwe przy bardzo wąskich lub niskich viewportach)
      if (tx < MARGIN) tx = MARGIN;
      if (ty + hintH + MARGIN > H) ty = H - hintH - MARGIN;

      return { tx, ty, originX, originY };
    }

    // Przesuwa dymek w kierunku pozycji (x, y) z płynną interpolacją (lerp).
    // Każde wywołanie tylko aktualizuje target — animacja sama dobiegnie do celu.
    function moveCursorHint(x, y) {
      if (!cursorHint) return;

      const { tx, ty, originX, originY } = computeHintPosition(x, y);

      // transformOrigin musi być ustawiony przed każdym ruchem bo może się
      // zmienić gdy hint przeskakuje między narożnikami ekranu
      cursorHint.style.transformOrigin = `${originX} ${originY}`;
      cursorHintTargetX = tx;
      cursorHintTargetY = ty;

      // Nie startuj nowej pętli animacji jeśli już działa
      if (cursorHintFrame !== null) return;

      const animateHint = () => {
        // Lerp 24% per klatka — hint "goni" kursor z lekkim opóźnieniem
        cursorHintX += (cursorHintTargetX - cursorHintX) * 0.24;
        cursorHintY += (cursorHintTargetY - cursorHintY) * 0.24;

        // Snap do celu gdy różnica jest sub-pikselowa — zatrzymuje pętlę rAF
        if (Math.abs(cursorHintTargetX - cursorHintX) < 0.2) cursorHintX = cursorHintTargetX;
        if (Math.abs(cursorHintTargetY - cursorHintY) < 0.2) cursorHintY = cursorHintTargetY;

        cursorHint.style.transform = `translate3d(${cursorHintX}px, ${cursorHintY}px, 0)`;

        // Kontynuuj animację dopóki nie dotrzemy do celu
        if (cursorHintX !== cursorHintTargetX || cursorHintY !== cursorHintTargetY) {
          cursorHintFrame = window.requestAnimationFrame(animateHint);
        } else {
          cursorHintFrame = null; // animacja skończona, pętla się zatrzyma
        }
      };

      cursorHintFrame = window.requestAnimationFrame(animateHint);
    }

    // Anuluje oczekujący timer opóźnienia (jeśli istnieje)
    function clearCursorHintTimer() {
      if (cursorHintTimer === null) return;
      window.clearTimeout(cursorHintTimer);
      cursorHintTimer = null;
    }

    // Wpisuje tekst do dymka, obsługując separator /| jako podział linii.
    // Czyści span przed wpisaniem żeby nie dublować zawartości.
    function setHintText(hintContent) {
      const span = cursorHint && cursorHint.querySelector("span");
      if (!span) return;

      span.textContent = "";
      String(hintContent).split("/|").forEach((line, index) => {
        if (index > 0) span.appendChild(document.createElement("br"));
        span.appendChild(document.createTextNode(line.trim()));
      });
    }

    // Pokazuje hint natychmiast — ustawia tekst, klasę CSS i uruchamia animację pozycji.
    // opts.autoHide=false → nie chowaj po czasie (tryb „przytrzymanie"); opts.durationMs →
    // własny czas auto-hide (tryb „tap = zerknięcie"). Bez opts → zachowanie domyślne.
    function showCursorHint(el, x, y, opts) {
      if (isDisabled(el, activePointerType)) return;
      setHintText(getHintText(el));
      // Resetujemy className żeby wyczyścić poprzednią data-hint-class,
      // potem dokładamy is-visible i ewentualną klasę z atrybutu
      cursorHint.className = `cursor-hint is-visible ${el.dataset.hintClass || ""}`.trim();
      moveCursorHint(x, y);
      // Auto-hide: jeśli element ma data-hint-duration (lub opts.durationMs), ukryj po czasie
      if (autoHideTimer) { window.clearTimeout(autoHideTimer); autoHideTimer = null; }
      const autoHide = !opts || opts.autoHide !== false;
      const durationMs = (opts && Number.isFinite(opts.durationMs)) ? opts.durationMs : getHintDurationMs(el);
      if (autoHide && durationMs > 0) {
        autoHideTimer = window.setTimeout(() => {
          autoHideTimer = null;
          hideCursorHint();
        }, durationMs);
      }
    }

    // Ostatnia zarejestrowana pozycja myszy podczas oczekiwania na timer —
    // hint pojawi się dokładnie tam gdzie mysz stała gdy timer dobiegł
    let pendingHintX = 0;
    let pendingHintY = 0;
    // Pozycja myszy przy ostatnim resecie timera — do pomiaru dystansu ruchu
    let lastMoveX = 0;
    let lastMoveY = 0;
    // Minimalny ruch myszy (w px) który resetuje timer opóźnienia
    const MOVE_THRESHOLD = 4;

    // Planuje pokazanie hinta po zadanym opóźnieniu.
    // Jeśli delay = 0, pokazuje od razu bez timera.
    // Każde wywołanie anuluje poprzedni timer — timer liczy od nowa.
    function scheduleCursorHint(el, event) {
      clearCursorHintTimer();
      activeHintEl = el;
      activePointerType = event.pointerType || "mouse";
      // Zapamiętaj pozycję — hint pojawi się tu gdy timer dobiegnie
      pendingHintX = event.clientX;
      pendingHintY = event.clientY;

      const delayMs = getHintDelayMs(el, activePointerType);

      // Delay 0 → pokaż natychmiast, bez timera (pointermove nie będzie go resetować)
      if (delayMs <= 0) {
        showCursorHint(el, pendingHintX, pendingHintY);
        return;
      }

      cursorHintTimer = window.setTimeout(() => {
        cursorHintTimer = null;
        // Sprawdź czy użytkownik nadal jest na tym samym elemencie —
        // mógł zjechać myszą zanim timer dobiegł
        if (activeHintEl !== el) return;
        showCursorHint(el, pendingHintX, pendingHintY);
      }, delayMs);
    }

    // Finalizuje ukrycie — usuwa klasy, kasuje zmienną fade i wysyła dymek poza ekran
    function finishHide() {
      if (!cursorHint) return;
      cursorHint.classList.remove("is-fading", "is-visible");
      cursorHint.style.removeProperty("--hint-fade-ms");
      cursorHintTargetX = -999;
      cursorHintTargetY = -999;
      moveCursorHint(-999, -999);
    }

    // Chowa hint i anuluje wszystkie aktywne timery i stany.
    // Jeśli element miał data-hint-fade, używa powolnego zanikania przez mgłę.
    // Czas fade: data-hint-fade="0.5" → 500ms; samo data-hint-fade → 460ms (domyślnie).
    function hideCursorHint() {
      clearCursorHintTimer();
      if (autoHideTimer) { window.clearTimeout(autoHideTimer); autoHideTimer = null; }
      if (fadingOutTimer) { window.clearTimeout(fadingOutTimer); fadingOutTimer = null; }

      const fadingEl = activeHintEl; // zapamiętaj przed wyczyszczeniem stanu
      activeHintEl = null;
      activePointerType = "";
      if (!cursorHint) return;

      const slowFade = fadingEl && fadingEl.dataset.hintFade !== undefined
        && cursorHint.classList.contains("is-visible");

      if (slowFade) {
        const rawFade = fadingEl.dataset.hintFade;
        const parsedFade = parseFloat(String(rawFade).replace(",", "."));
        const fadeDurationMs = Number.isFinite(parsedFade) && parsedFade > 0
          ? Math.round(parsedFade * 1000)
          : 460;
        cursorHint.style.setProperty("--hint-fade-ms", `${fadeDurationMs}ms`);
        cursorHint.classList.add("is-fading");
        fadingOutTimer = window.setTimeout(() => {
          fadingOutTimer = null;
          finishHide();
        }, fadeDurationMs + 20);
      } else {
        finishHide();
      }
    }

    // Podpina eventy na liście elementów. Pomija elementy już podpięte
    // (data-cursor-hint-bound="1") — bezpieczne przy wielokrotnym wywołaniu
    // np. przez MutationObserver gdy DOM się zmienia dynamicznie.
    function setupCursorHint(elements, clickCallback = null) {
      elements.forEach((el) => {
        if (!el || el.dataset.cursorHintBound === "1") return;
        el.dataset.cursorHintBound = "1"; // znacznik "już podpięty"

        // Mysz wchodzi na element — start timera (dotyk obsługuje pointerdown)
        el.addEventListener("pointerenter", (event) => {
          event.stopPropagation(); // blokuj bąbelkowanie do rodzica który też może mieć hint
          if (event.pointerType === "touch" || isDisabled(el, event.pointerType || "mouse")) return;
          lastMoveX = event.clientX;
          lastMoveY = event.clientY;
          scheduleCursorHint(el, event);
        });

        el.addEventListener("pointermove", (event) => {
          event.stopPropagation();
          const pointerType = event.pointerType || activePointerType || "mouse";
          if (isDisabled(el, pointerType)) return;

          // Dotyk tap/hold: większy ruch zanim pokażemy = scroll → anuluj
          if (pointerType === "touch" && el.dataset.hintTap !== undefined
              && activeHintEl === el && !touchHeldShown && !touchMoved) {
            const mx = event.clientX - touchStartX;
            const my = event.clientY - touchStartY;
            if (Math.sqrt(mx * mx + my * my) >= TOUCH_MOVE_CANCEL) {
              touchMoved = true;
              if (touchHoldTimer) { window.clearTimeout(touchHoldTimer); touchHoldTimer = null; }
              hideCursorHint();
              return;
            }
          }

          if (cursorHint && cursorHint.classList.contains("is-visible")) {
            // Hint już widoczny — tylko przesuń za kursorem
            moveCursorHint(event.clientX, event.clientY);
          } else if (activeHintEl === el) {
            // Hint jeszcze niewidoczny — sprawdź czy mysz ruszyła wystarczająco
            // żeby zresetować timer (mikrodrgania ręki poniżej progu są ignorowane)
            const dx = event.clientX - lastMoveX;
            const dy = event.clientY - lastMoveY;
            if (Math.sqrt(dx * dx + dy * dy) >= MOVE_THRESHOLD) {
              lastMoveX = event.clientX;
              lastMoveY = event.clientY;
              scheduleCursorHint(el, event); // reset — odliczaj od nowa
            }
          }
        });

        // Dotyk: timer startuje przy przyciśnięciu palca (nie przy wejściu)
        el.addEventListener("pointerdown", (event) => {
          if (event.pointerType !== "touch" || isDisabled(el, "touch")) return;
          event.stopPropagation();
          if (el.dataset.hintTap === undefined) { scheduleCursorHint(el, event); return; }
          // Tryb tap/hold: ustaw stan i odpal timer przytrzymania
          clearCursorHintTimer();
          if (touchHoldTimer) { window.clearTimeout(touchHoldTimer); touchHoldTimer = null; }
          activeHintEl = el;
          activePointerType = "touch";
          touchDownTime = nowMs();
          touchHeldShown = false;
          touchMoved = false;
          touchStartX = event.clientX;
          touchStartY = event.clientY;
          const px = event.clientX;
          const py = event.clientY;
          touchHoldTimer = window.setTimeout(() => {
            touchHoldTimer = null;
            if (activeHintEl !== el || touchMoved) return;
            showCursorHint(el, px, py, { autoHide: false }); // trzymanie → pokazuj póki trzymasz
            touchHeldShown = true;
          }, TOUCH_HOLD_MS);
        });

        // Blokuj systemowe context menu (iOS "Kopiuj/Szukaj", Android long-press)
        // gdy hint dotykowy jest w trakcie oczekiwania lub już widoczny na tym elemencie
        el.addEventListener("contextmenu", (event) => {
          if (activeHintEl === el) event.preventDefault();
        });

        // Mysz opuszcza element — schowaj. Dotyk obsługują pointerup/cancel/auto-hide,
        // więc tu go pomijamy (inaczej pointerleave po tapie ubiłby „zerknięcie").
        el.addEventListener("pointerleave", (event) => {
          if (event && event.pointerType === "touch") return;
          hideCursorHint();
        });

        // Podniesienie palca: w trybie tap/hold rozróżnij tap (zerknięcie) od trzymania.
        el.addEventListener("pointerup", (event) => {
          if (event && event.pointerType === "touch" && el.dataset.hintTap !== undefined) {
            if (touchHoldTimer) { window.clearTimeout(touchHoldTimer); touchHoldTimer = null; }
            if (touchMoved) { hideCursorHint(); return; }
            if (nowMs() - touchDownTime < TOUCH_TAP_MAX) {
              const dur = getHintDurationMs(el);
              showCursorHint(el, event.clientX, event.clientY, { autoHide: true, durationMs: dur > 0 ? dur : TOUCH_PEEK_MS });
            } else {
              hideCursorHint(); // długie trzymanie → puszczenie chowa
            }
            return;
          }
          hideCursorHint();
        });
        el.addEventListener("pointercancel", () => {
          if (touchHoldTimer) { window.clearTimeout(touchHoldTimer); touchHoldTimer = null; }
          hideCursorHint();
        });

        if (clickCallback) {
          el.addEventListener("click", () => clickCallback(el));
        }
      });
    }

    return { setupCursorHint };
  }

  // Publiczne API — inicjalizuje system hintów dla całej strony.
  // Automatycznie podpina nowe elementy gdy pojawią się w DOM (MutationObserver).
  function initCursorHints({
    selector = "[data-hint], [data-hint-pl], [data-hint-en]",
    fallbackHint = "", // tekst gdy element ma pusty data-hint (np. "Kliknij")
  } = {}) {
    const cursorHint = document.getElementById("cursorHint") || document.getElementById("cursor-hint");
    const prefersReducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    const controller = createCursorHintController({
      cursorHint,
      prefersReducedMotion,
      getFallbackHint: () => fallbackHint,
    });
    const bindTargets = () => controller.setupCursorHint(document.querySelectorAll(selector));
    bindTargets(); // podepnij elementy istniejące przy starcie
    if ("MutationObserver" in window) {
      // Obserwuj zmiany w DOM — np. dynamicznie dodawane przyciski w panelach
      const observer = new MutationObserver(bindTargets);
      observer.observe(document.body, { childList: true, subtree: true });
    }
    return controller;
  }

  return { createCursorHintController, initCursorHints };
})();