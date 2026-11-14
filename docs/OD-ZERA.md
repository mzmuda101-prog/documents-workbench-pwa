# Od zera — czy da się zrobić taki dokument w aplikacji?

Pomysł Mateusza (2026-11-14): nie tylko otwierać i poprawiać pliki z bazy testów, ale umieć
każdy z nich **napisać od zera** — docelowo na poziomie Worda. Ten plik to wynik przeglądu bazy
i lista braków w kolejności, w jakiej warto je robić.

**Jak powstał:** `node tools/od-zera/inventory-docx.js <pliki|foldery>` wypisuje, czego z Worda
używa każdy plik (style, listy, tabele, obrazy, sekcje, pola…). Potem dla każdej cechy:
czy aplikacja pozwala ją zrobić samemu (menu, pasek, skróty), a nie tylko wyświetlić z pliku.
Baza: 24 pliki — prawdziwe (zarządzenie z załącznikami, foldery FELU, „Folder v2 – materiały”,
opis pras, transkrypcja, zapotrzebowanie, podania, sprawozdanie, notatki z lekcji i mapy z PDF)
oraz próbki z `docs/samples`. Sondy (`vm-*`, `output/word-oracle`) pominięte.
Po dodaniu plików do bazy wystarczy uruchomić to polecenie jeszcze raz.

## Co już da się zrobić od zera

Tekst (krój, rozmiar, kolor, wyróżnienie, indeks górny/dolny, wielkość liter, malarz formatów),
akapit (wyrównanie, wcięcia, odstępy, interlinia), style z listy (Normalny, Tytuł, Nagłówki 1–3,
Cytat, Ramka), listy punktowane i numerowane z poziomami, tabele (wstawianie, wiersze/kolumny,
scalanie i dzielenie, wiersz nagłówka, cieniowanie, krawędzie, szerokości), obraz w linii,
pole tekstowe, nagłówek i stopka z numerem strony (inna pierwsza strona), spis treści, przypisy
dolne i końcowe, komentarze, linki, zakładki, pola formularza (tekst, pole wyboru, lista, data),
podział strony, linia pozioma, symbole (też matematyczne i greckie), układ strony (marginesy,
orientacja, rozmiar, wyrównanie w pionie), śledzenie zmian i recenzja.

## Braki — w kolejności (ile prawdziwych plików tego używa)

| # | Brak | Plików | Co w Wordzie | Uwagi |
|---|---|---|---|---|
| 1 | **Własne style: utwórz / zmień** | 17 z 24 | Style → Utwórz styl, Modyfikuj | Pisma urzędowe i foldery mają 18–34 własnych stylów. Dziś można tylko wybrać z 8 stałych. Najpierw: „Nowy styl z zaznaczenia” i „Zaktualizuj styl do zaznaczenia”. |
| 2 | **Format numeracji list** | 13 z 24 | Numerowanie → a) / i. / I. / 1.1. | Pliki używają liter i rzymskich. Dziś tylko domyślne „1.” (poziomy dają a., i. same). Do tego: zacznij od nowa / kontynuuj numerację, wielopoziomowa „1.1.1”. |
| 3 | **Tabulatory od zera** | 10 z 24 (3 prawe / środkowe / dziesiętne, 2 z kropkami) | Linijka, Akapit → Tabulatory | Dziś da się tylko przesunąć istniejące kropki. Potrzebne okienko „Tabulatory” (pozycja, wyrównanie, wypełnienie) i może linijka. |
| 4 | **Obraz pływający i zawijanie tekstu** | 5 (4 za tekstem) | Zawijanie: przyległe, za tekstem, przed tekstem | Już w planie (punkt 4). Pliki z PDF mają tło-skan „za tekstem”. |
| 5 | **Czcionka zaawansowana** | 5 | Czcionka → Zaawansowane | Odstęp między znakami, kapitaliki, wersaliki, przekreślenie podwójne. |
| 6 | **Komórka tabeli: wyrównanie w pionie, kierunek tekstu** | 2 + 1 | Układ tabeli → Wyrównanie, Kierunek | Proste w XML (w:vAlign, w:textDirection). |
| 7 | **Sekcje** | 2 | Układ → Podziały → Następna strona | Inne nagłówki w sekcji, parzyste/nieparzyste, strona pozioma w środku pionowego dokumentu. |
| 8 | **Tabela w tabeli** | 2 | Wstaw tabelę w komórce | Mapy z PDF tak wyglądają. |
| 9 | **Style tabel (galeria)** | 12 (zwykle sama „Tabela – siatka”) | Projekt tabeli → Style | Niski zysk: pliki rzadko mają coś ponad siatkę. |

Bez plików w bazie, ale Word to ma (w kolejce z planu): kolumny tekstu, znak wodny, sortowanie
tabeli, edytor równań, Alt+X. Jeszcze dalej: inicjał, obramowanie strony, podpisy i odsyłacze
(„Rysunek 1”), kształty, wykresy, SmartArt.

## Jak to sprawdzać

Każda pozycja kończy się testem „zrobione w aplikacji = jak w Wordzie”: dokument zbudowany od
zera w aplikacji, zapisany, otwarty prawdziwym Wordem (`tools/word-oracle`) — Word widzi te same
style / listy / tabulatory co plik wzorcowy z bazy. Dobra miara postępu: dla każdego pliku z
bazy lista „czego nie da się zrobić od zera” (inventory + mapa wyżej) ma maleć do zera.
