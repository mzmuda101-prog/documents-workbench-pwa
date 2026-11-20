// @ts-check
// „Co nowego”: lista zmian w wydaniach. Po aktualizacji okno pokazuje się samo (raz na wersję),
// a całą listę da się zawsze otworzyć z menu ⋯ → Pomoc → Co nowego — treść nie znika razem z dymkiem.
// Przy wydaniu (npm run release) dopisz NA GÓRZE wpis z nową wersją; skrypt przypomina, gdy go brak.

/* global APP_BUILD_VERSION, appDialog, t, currentLang */

const WHATS_NEW = [
  {
    v: "v1.01.35",
    pl: ["PDF → Word: wykłady i notatki z LaTeX-a — polskie litery są poprawne („Możemy”, a nie „Mo»emy”), słowa mają spacje, a wiersze stoją w dobrej kolejności."],
    en: ["PDF → Word: LaTeX lectures and notes — Polish letters are correct, words keep their spaces and lines stay in the right order."],
  },
  {
    v: "v1.01.34",
    pl: ["Telefon: w arkuszu „Główne” jest teraz „Cofnij” obok „Ponów” — nie trzeba go szukać w górnym pasku."],
    en: ["Phone: the “Home” sheet now has “Undo” next to “Redo” — no need to look for it in the top bar."],
  },
  {
    v: "v1.01.33",
    pl: [
      "PDF → Word: zakreślenia z PDF-a są tłem tekstu, jak w Wordzie — tekst nie przesuwa się, nie zmienia kolejności i nie ucina się na telefonie.",
      "Linie do wpisywania („recover ______ energy”) widać na całą długość, ze spacjami po obu stronach — także w widoku mobilnym.",
      "Wydruk zachowuje zakreślenia, tło komórek i linie tabulatorów.",
    ],
    en: [
      "PDF → Word: highlights from a PDF become text shading, like in Word — text no longer shifts, changes order or gets cut off on phones.",
      "Fill-in lines (“recover ______ energy”) show at full length with spaces on both sides — in mobile view too.",
      "Printing keeps highlights, cell shading and tab lines.",
    ],
  },
  {
    v: "v1.01.32",
    pl: [
      "Nowy ekran startowy: krótszy tytuł, równe przyciski, a na telefonie ostatnie pliki od razu widoczne.",
      "Ostatnio otwierane: pliki o tej samej nazwie jako jedna pozycja — starsze wersje rozwiniesz pod „+3 starsze wersje”. Przy każdym pliku widać, czy to Word, czy PDF.",
      "Bez otwartego pliku nie ma narzędzi do dokumentu (lewy pasek, tryb, powiększenie), a szukanie jest wyszarzone.",
    ],
    en: [
      "New start screen: a shorter title, even buttons, and on phones your recent files are visible right away.",
      "Recently opened: files with the same name are one entry — expand older versions under “+3 older versions”. Each file shows whether it is Word or PDF.",
      "With no file open, document tools (left bar, mode, zoom) are hidden and search is greyed out.",
    ],
  },
  {
    v: "v1.01.31",
    pl: ["Lepszy kontrast: szare podpisy są ciemniejsze, a niebieskie przyciski w ciemnym motywie czytelniejsze."],
    en: ["Better contrast: grey captions are darker and blue buttons in dark mode are easier to read."],
  },
  {
    v: "v1.01.30",
    pl: [
      "Nowa wersja aplikacji włącza się dopiero po kliknięciu „Aktualizuj” — nie wchodzi już „po części” sama przy otwarciu.",
      "Okno „Co nowego” po każdej aktualizacji. Otworzysz je też w każdej chwili: menu ⋯ → Co nowego.",
    ],
    en: [
      "A new app version starts only after you tap “Update” — it no longer half-loads by itself on launch.",
      "A “What's new” window after every update. Open it any time from the ⋯ menu → What's new.",
    ],
  },
  {
    v: "v1.01.29",
    pl: ["Szukanie: po wpisaniu nowej frazy strzałki ↑ ↓ szukają nowej — bez klikania „Odznacz”."],
    en: ["Search: after typing a new phrase, the ↑ ↓ arrows search for it — no need to tap “Clear” first."],
  },
  {
    v: "v1.01.28",
    pl: ["Tryb skupienia: wyjście na tablecie i przycisk na dolnym pasku."],
    en: ["Focus mode: exit on tablets and a button on the bottom bar."],
  },
  {
    v: "v1.01.27",
    pl: ["Esc w polu Szukaj wraca do dokumentu."],
    en: ["Esc in the Search field returns to the document."],
  },
  {
    v: "v1.01.26",
    pl: ["Wstążka przyklejona do górnej krawędzi okna."],
    en: ["The ribbon sticks to the top edge of the window."],
  },
];

const WHATS_NEW_SEEN_KEY = "documents-workbench-seen-version";

/** @param {string} [since] tylko wydania nowsze niż ta wersja; bez — cała lista */
function showWhatsNew(since) {
  // vMAJOR.MM.PP z zerami — porównanie napisów = porównanie wersji
  const items = since ? WHATS_NEW.filter((e) => e.v > since) : WHATS_NEW;
  if (!items.length) return;
  const body = document.createElement("div");
  body.className = "whats-new";
  for (const e of items) {
    const h = document.createElement("h3");
    h.textContent = e.v;
    const ul = document.createElement("ul");
    for (const line of currentLang === "en" ? e.en : e.pl) {
      const li = document.createElement("li");
      li.textContent = line;
      ul.append(li);
    }
    body.append(h, ul);
  }
  const title = since && items.length === 1 ? t("whatsNewTitle", { v: items[0].v }) : t("whatsNewAllTitle");
  appDialog({ title, body, actions: [{ label: t("whatsNewOk"), value: true, kind: "primary" }] });
}

(function initWhatsNew() {
  document.getElementById("whatsNewMenuItem")?.addEventListener("click", () => showWhatsNew());
  let seen = null;
  try {
    seen = localStorage.getItem(WHATS_NEW_SEEN_KEY);
    if (seen === APP_BUILD_VERSION) return;
    localStorage.setItem(WHATS_NEW_SEEN_KEY, APP_BUILD_VERSION);
  } catch (_) {
    return; // bez pamięci przeglądarki okno wyskakiwałoby przy każdym starcie
  }
  // Pierwsze wejście w ogóle (brak workera) — nic. Pierwsza aktualizacja po dodaniu „Co nowego”
  // (worker jest, zapisanej wersji jeszcze nie) — tylko ta wersja.
  const returning = !!navigator.serviceWorker?.controller;
  if (!seen && !returning) return;
  const since = seen || WHATS_NEW[1]?.v || "";
  const show = () => setTimeout(() => showWhatsNew(since), 600);
  if (document.readyState === "complete") show();
  else window.addEventListener("load", show, { once: true });
})();
