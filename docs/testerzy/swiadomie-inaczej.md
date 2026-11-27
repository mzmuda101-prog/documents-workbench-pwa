# Świadomie inaczej niż Word

Do czego służy ten plik: lista rzeczy, w których aplikacja różni się od Worda za zgodą właściciela
(Mateusza), oraz znanych i zaakceptowanych ograniczeń. Testerzy AI NIE zgłaszają ich jako błędów
ani tarcia. Kto prowadzi przebieg, usuwa wpis, gdy decyzja się zmieni. Nie dotyczy to otwartych błędów.

Format wpisu: `- co się różni od Worda — dlaczego / decyzja (data, jeśli znana) (źródło: plik)`.
Źródła w `memory/` leżą w `~/.claude/projects/-Users-mateuszuser-Programistyka-documents-workbench-pwa/memory/`.

## Tabele

- Sortuj…: wiersz z `w:tblHeader` zostaje na górze po sortowaniu — świadomie inaczej niż Word (2026-11-15) (źródło: memory/plan-after-v1-00-56.md)
- Sortuj…: rodzaj danych (liczby / daty / tekst) wybierany automatycznie wg większości komórek (2026-11-15) (źródło: memory/plan-after-v1-00-56.md)
- Okno Sortuj… na start wybiera kolumnę z kursorem, a Word pokazał pierwszą kolumnę — zostawione, „na razie tak zostawmy” (2026-11-15) (źródło: memory/sort-dialog-word-diff.md)
- Okno Sortuj… nie pamięta ustawień poprzedniego sortowania (Word pamięta, nawet z innej tabeli); startuje od rozpoznania bieżącej tabeli (2026-11-15) (źródło: memory/sort-dialog-word-diff.md)
- Backspace w tabeli z zaznaczeniem wielu komórek tylko czyści zawartość; Word pokazuje okno „usuń komórki / wiersz / kolumnę” (źródło: memory/next-review-and-word-gaps.md, pkt 5 — na liście kroków; patrz też „Znane, odłożone”)
- Szerokość kolumn i wysokość wierszy zmienia się przeciąganiem krawędzi myszą lub rysikiem; dotyk bez przeciągania krawędzi (źródło: docs/OD-ZERA.md, tabela „Gdzie co ma być”; memory/next-review-and-word-gaps.md)
- Brak galerii stylów tabel (Projekt tabeli → Style); jest siatka, cieniowanie i krawędzie — niski zysk, pliki rzadko używają więcej niż „Tabela – siatka” (źródło: docs/OD-ZERA.md, brak nr 9)

## Klawiatura

- Litery nie są skrótami klawiszowymi (klawiatury tabletowe); F1–F12 tylko jako alias skrótu (źródło: AGENTS.md, „Zasady UI”)
- Backspace tuż za podziałem strony i Delete tuż przed nim: aplikacja skleja akapity i usuwa podział; Word zostawia podział i usuwa go dopiero drugim Backspace. ODSTĘPSTWO do czasu edycji akapitu z podziałem w środku (decyzje 2026-11-08 / 2026-11-09) (źródło: scripts/keys-table.js, bloki `afterBreak` i `beforeBreak`; docs/JAKOSC.md)
- Backspace w pustej linijce pod linią poziomą skleja akapity, a samą linię usuwa menu; nowa linijka nie dziedziczy linii (zgodne z Wordem wg wyroczni, decyzja 2026-11-08) (źródło: scripts/keys-table.js)
- Reguła decyzji: gdzie Word ma odpowiednik, zachowanie ma być jak w Wordzie (sprawdzane wyrocznią `tools/word-oracle`); tylko gdzie Word nie ma odpowiednika, obowiązuje rekomendacja (2026-11-08) (źródło: memory/behavior-decisions-like-word.md)

## Wstążka i układ

- Układ interfejsu wg makiety `wstazka-v4.html`, wdrażany 1:1; różnice od Worda są zatwierdzone, nie kopiujemy wstążki 1:1 (telefon) (źródło: memory/follow-mockup.md, memory/ribbon-mockup.md, memory/feature-placement.md)
- Jedna funkcja = jedno miejsce; wyjątek jak w Wordzie: Komentarz na kartach Wstaw i Recenzja. Zoom tylko na dolnym pasku, tryb (Edycja / Sugerowanie / Czytanie) tylko w przycisku u góry — karta Widok nie ma ani Czytania, ani grupy Powiększenie (źródło: memory/no-duplicate-controls.md; docs/WSTAZKA.md, etap 1b)
- Różne funkcje mają różne ikony; samotna ikona bez znanego znaczenia dostaje podpis (źródło: memory/no-duplicate-controls.md)
- Przy braku miejsca rzędy wstążki przewijają się w bok (wygaszenie, strzałka „+N”, jednorazowe drgnięcie); polecenia nie chowają się do „⋯” ani list (źródło: memory/ribbon-mockup.md; AGENTS.md: co się nie mieści, przewija się)
- To, co może się nie zmieścić, przewija się — nic nie jest ucinane po cichu (źródło: AGENTS.md, „Zasady UI”)
- Brak w karcie Układ pól Wcięcie/Odstęp (zamiast nich przycisk „Odstępy i wcięcia…”), brak Linijki, „Zaznacz”, Poprzedni/Następny komentarz i zmiana, Okienka „Snippety” w Korespondencji, Powiększenia i Wyglądu w Widoku — decyzje z przeglądu (źródło: docs/WSTAZKA.md, etap 6b)
- Menu Odstępy na Narzędziach głównych zawiera tylko interlinię i odstęp przed/po; wcięcia to osobne przyciski; Tabulatory… i „Odstępy i wcięcia…” leżą na karcie Układ (źródło: docs/WSTAZKA.md, etap 1b; memory/no-duplicate-controls.md)
- Recenzja: Akceptuj wszystkie / Odrzuć wszystkie / Przejrzyj po kolei; bez przycisku Sugerowanie (tryb jest u góry); Plik bez Porównaj i bez Asystent AI (są w Recenzji i na karcie Asystent) (źródło: docs/WSTAZKA.md, etap 1b)
- Czwarty kafelek „Nowy” to „Przewodnik po aplikacji” zamiast „Pismo z placeholderami” z makiety — takiego szablonu jeszcze nie ma (źródło: docs/WSTAZKA.md, etap 2)
- Karta kontekstowa „Tabela” pojawia się tylko, gdy kursor jest w tabeli (jak w Wordzie); funkcje aplikacji spoza makiety dołożone do najbliższych grup (źródło: docs/WSTAZKA.md, etap 6b)
- Dziennik techniczny („Log”) i pigułki „niezapisane” / „Dokument wczytany” są schowane (licznik zmian jest na przycisku Zapisz; Log widać z `?debug`) (źródło: docs/WSTAZKA.md, etap 1b)
- Przycisk × zamknięcia karty dokumentu pokazuje się pod kursorem (na dotyku stoi zawsze) — odstępstwo od makiety zatwierdzone przez Mateusza; „Dopasuj” zostaje przy powiększeniu; przy zoomie trzy ikonki widoku Auto / mobilny / desktopowy (v1.01.08) (źródło: docs/WSTAZKA.md, etap 4)
- Intro aplikacji gaśnie 0,6 s, ponad budżet ruchu 500 ms — użytkownik nie zdecydował (źródło: memory/emil-polish-pass.md)
- Poniżej 1000 px na dolnym pasku chowają się wyrazy i język (źródło: docs/WSTAZKA.md, etap 4)
- Telefon: górny pasek (←, Cofnij, nazwa i tryb, ⋯), Zapisz jest w ⋯ i na wyspie; rząd poleceń nad klawiaturą i arkusz kart zamiast wstążki; w Czytaniu na telefonie nie ma paska poleceń ani pływającego „Edytuj” (tryb zmienia się u góry) (źródło: docs/WSTAZKA.md, etap 7a / 7b)
- Telefon: „←” zamyka dokument do ekranu startowego (pyta o niezapisane); wyspa szukania u dołu jest ustawieniem, domyślnie włączona (decyzje Mateusza 2026-11-17) (źródło: docs/WSTAZKA.md, etap 7a)
- Na telefonie te same menu co na komputerze, przyciski jako ikony (do 480 px); polecenia mają jeden dom, a nie 1:1 kopię wstążki Worda (źródło: docs/OD-ZERA.md, „Gdzie co ma być”)
- Dymki powiadomień układają się w stos jak w Sonner; stos i pojedynczy dymek reagują na RUCH myszy, nie na samo wejście (v1.00.59–60) (źródło: memory/plan-after-v1-00-56.md)

## PDF

- PDF → Word: kolejność czytania ramek infografiki bywa inna niż w oryginale (ramki obok siebie z kreskami na tych samych wysokościach sklejają się w tabelę); naprawa dotknęłaby tabel we wszystkich PDF-ach — odłożone świadomie, ryzyko > zysk (źródło: docs/JAKOSC.md, „Dalej” pkt 5; memory/plan-after-v1-00-56.md)
- PDF → Word z OCR: w bardzo gęstych tabelach podgląd jest ok. 1 pt na wiersz wyższy niż Word (ostatni wiersz mapy B na stronie 2 w podglądzie, na 1 w Wordzie) (źródło: docs/JAKOSC.md, „Dalej” pkt 5)
- PDF → Word z OCR: mapa Canona A w Wordzie jest o kilka wierszy dłuższa (krótkie napisy w wąskich komórkach zawijają się); gęsta infografika po OCR bywa o 2–3 wiersze dłuższa niż strona (źródło: docs/JAKOSC.md, „Dalej” pkt 5)
- Strona pisana ręcznie w PDF / skanie zostaje w całości obrazem (bez OCR tekstu); strona większa niż 22″ jest zmniejszana do A4 (źródło: docs/JAKOSC.md, „Dalej” pkt 5; app/whats-new.js)

## Edycja i strony

- Wysoki wiersz tabeli w Edycji przenosi się na następną stronę w całości, a Podgląd wydruku dzieli go między strony jak Word — świadomie NIE zmieniamy (2026-11-18): w 23 prawdziwych plikach ani jeden wiersz nie jest dzielony przez Word, ryzyko duże; wrócić przy pierwszym prawdziwym pliku (źródło: docs/JAKOSC.md, „Dalej”)
- Akapit z podziałem strony w środku jest tylko do odczytu (lockPageBreak) — świadomie NIE edytujemy (2026-11-10): w 11 plikach Worda wystąpił raz, w pliku testowym; założenie „akapit na ekranie = akapit w pliku” (źródło: docs/JAKOSC.md, „Dalej” pkt 1; scripts/keys-table.js)
- Dodanie komentarza przerysowuje cały dokument (bez szybkiej ścieżki) — świadomie NIE (2026-11-10), rzadka operacja, ryzyko > zysk (źródło: docs/JAKOSC.md, pkt 0; memory/next-stage-reloads.md)
- Pełne przerysowanie zostaje przy: stylu, liście, linii poziomej, przypisach, ustawieniach strony, przyjmij / odrzuć wszystko, spisie treści, nagłówku i stopce — zaakceptowane; zmiana wyrównania, wcięć i odstępów jest łatką bez przerysowania (źródło: docs/JAKOSC.md, pkt 0 B i C; memory/next-stage-reloads.md)
- Edycja w akapitach z zapisanymi zmianami (rPrChange — zmiana formatowania; przeniesienia z markerami zakresu) jest zablokowana celowo; edycja akapitów z cudzymi zmianami przy włączonym śledzeniu — zablokowana (`lockTrackedOn`) (źródło: memory/next-review-and-word-gaps.md)
- Zawijanie tekstu wokół obrazu: tryby w linii / przyległe / góra i dół / za / przed tekstem są; przeciąganie obrazu myszą (zmiana pozycji) — nie ma (źródło: docs/OD-ZERA.md, brak nr 4)
- Okno Czcionka…: brak podwójnego przekreślenia i skali znaków (podgląd ich nie rysuje) (źródło: docs/OD-ZERA.md, brak nr 5)
- Tabulatory ustawia się oknem „Tabulatory…”; Linijki do ustawiania nie ma (źródło: docs/OD-ZERA.md, brak nr 3; docs/WSTAZKA.md, etap 6b)
- Kolumny tekstu: jedna / dwie / trzy o równej szerokości; kolumny nierównej szerokości nie ma. Wyrównanie kolumn jak w Wordzie: tylko przed sekcją ciągłą (źródło: docs/OD-ZERA.md; memory/plan-after-v1-00-56.md)
- Wydruk: treść co najmniej 12,7 mm od krawędzi kartki (źródło: AGENTS.md, „Zasady UI”)
- Każde :hover działa tylko w `@media (hover: hover)`; ramka fokusu przycisków przez `:focus-visible` / `.kbd-focus` — na dotyku nie ma stanu hover (źródło: AGENTS.md, „Zasady UI”)

## Inne

- Nie ma emoji w UI, kodzie ani commitach; ikony z zestawu lucide (źródło: AGENTS.md, „Zasady UI”)
- Wklej z paska: odczyt schowka przez `clipboard.read`; przy odmowie przeglądarki pojawia się komunikat „użyj ⌘V” (ograniczenie przeglądarki) (źródło: memory/ribbon-mockup.md, „Zrobione w 1b”)
- Test schema w WebKit („Target crashed”), mini-bar w runnerze i guide z pokryciem padają także na czystym main — nie błędy aplikacji (stan 2026-11-21) (źródło: memory/known-failing-tests.md)
- Mierzone porównanie z Wordem w widoku Edycji (80,8 %) i w Podglądzie wydruku (81,5 %) na 23 plikach — reszta różnicy to szum miernika wierszy na odręcznych obrazach (źródło: docs/JAKOSC.md, „Dalej”)
- Pasek postępu konwersji PDF animuje `width` (nie `scaleX`), uchwyt panelu `.sidebar-handle` animuje `left` — celowo (źródło: memory/emil-polish-pass.md)

## Znane, odłożone

Testerzy mogą to pominąć, nie jest to nieznany błąd.

- Edycja wysokiego wiersza tabeli dzielonego między strony jak w Wordzie — wrócić przy prawdziwym pliku z takim wierszem (źródło: docs/JAKOSC.md, „Dalej”)
- Edycja akapitu z podziałem strony w środku (i zgodność Backspace / Delete z Wordem wokół podziału) — wrócić przy prawdziwym pliku z problemem (źródło: docs/JAKOSC.md, „Dalej” pkt 1; scripts/keys-table.js)
- Znak wodny (Układ → Znak wodny; Word zapisuje go jako kształt w nagłówku) — następny na liście (źródło: memory/plan-after-v1-00-56.md; docs/OD-ZERA.md)
- Kolumny nierównej szerokości i granica strony w środku akapitu w kolumnach (źródło: memory/plan-after-v1-00-56.md)
- Braki z listy „Od zera”: inicjał, obramowanie strony, podpisy i odsyłacze („Rysunek 1”), kształty, wykresy, SmartArt, edytor równań, Alt+X, galeria stylów tabel (źródło: docs/OD-ZERA.md)
- „Zmień styl…” (okno: nazwa, krój, rozmiar, kolor, wyrównanie, odstępy, bazowy, następny) i przeciąganie obrazu z zawijaniem myszą, linijka dla tabulatorów, przekreślenie podwójne i skala znaków (źródło: docs/OD-ZERA.md)
- Backspace na wielu zaznaczonych komórkach tabeli ma pokazać okno „usuń komórki / wiersz / kolumnę” jak Word; dziś tylko czyści (źródło: memory/next-review-and-word-gaps.md, pkt 5)
- Zmiana szerokości kolumn tabeli dotykiem (dziś mysz i rysik) (źródło: memory/next-review-and-word-gaps.md, pkt 5)
- Edycja akapitów z cudzymi zmianami przy włączonym śledzeniu (potrzebny diff wewnątrz akapitu w docx-redline.js) (źródło: memory/next-review-and-word-gaps.md, pkt 2)
- Komentarz → propozycja zmiany przez lokalne AI (źródło: memory/next-review-and-word-gaps.md, pkt 6)
- Ikony na wąskiej wstążce telefonu (Wstaw / Układ / Przepisz, lista stylów) i przegląd zmian po kolei z licznikiem i swipe — wg planu z v1.00.32; część mogła już wejść (sprawdź wydanie przed zgłoszeniem) (źródło: memory/next-review-and-word-gaps.md, pkt 1 i 3)
- Zostały pełne rendery listy / linii / komentarza / Cofnij w akapicie z komentarzem — szybkie, ale nie łatki (źródło: memory/next-stage-reloads.md)
- Do sprawdzenia na prawdziwym iPhonie (niezweryfikowane): gest „dymek w dół = zamknij”, `:active` scale na iOS, kolor paska stanu (theme-color), stuknięcie rozkładające stos powiadomień (źródło: memory/emil-polish-pass.md; memory/plan-after-v1-00-56.md)
- Panel narzędzi (klucz) na wyspie telefonu — docelowo może nie być potrzebny w tej formie, do przemyślenia (źródło: docs/WSTAZKA.md, etap 7b)
- Język i Dostępność w Recenzji — z etapu 1b zostały do zrobienia (źródło: memory/ribbon-mockup.md)
- Okno Sortuj…: ewentualne dopasowanie 1:1 do Worda (kolumna startowa, pamięć ustawień) — jedna linijka `keys[0].col = "0"` dla kolumny (źródło: memory/sort-dialog-word-diff.md)
- OCR w Wordzie: odłożone zmiany z listy pkt 5 (kolejność czytania ramek, gęste tabele OCR) — patrz sekcja PDF (źródło: docs/JAKOSC.md)
