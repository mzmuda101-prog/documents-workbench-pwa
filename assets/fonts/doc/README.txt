Czcionki zastępcze dokumentów (Documents Workbench)
===================================================

Darmowe kroje o tych samych szerokościach liter co kroje Office. Podgląd używa ich tylko
wtedy, gdy na urządzeniu nie ma oryginału. Pliki generuje scripts/gen-doc-fonts.py.

  Calibri          → Carlito   SIL OFL 1.1     carlito-OFL.txt   (Reserved Font Name "Carlito")
  Cambria          → Caladea   Apache 2.0      caladea-Apache-2.0.txt
                                (Copyright (c) 2012 Huerta Tipografia; crosextrafonts-20130214)
  Arial, Helvetica → Arimo     SIL OFL 1.1     arimo-OFL.txt
  Times New Roman  → Tinos     SIL OFL 1.1     tinos-OFL.txt
  Courier New      → Cousine   SIL OFL 1.1     cousine-OFL.txt
  Georgia          → Gelasio   SIL OFL 1.1     gelasio-OFL.txt

Źródła: https://github.com/google/fonts (ofl/carlito, ofl/arimo, ofl/tinos, ofl/cousine,
ofl/gelasio) oraz https://commondatastorage.googleapis.com/chromeos-localmirror/distfiles/crosextrafonts-20130214.tar.gz
(Caladea — wersja zgodna z metryką Cambrii).

Zmiany względem oryginałów: przycięcie do wybranych zakresów znaków (pliki *-core-*
i *-ext-*), usunięcie hintingu, konwersja do WOFF2, wewnętrzne nazwy krojów zmienione na
„DWB …” (wymóg licencji przy zastrzeżonej nazwie), w Gelasio metryki pionowe wyrównane do
Georgii. Nazwy Calibri, Cambria, Arial, Times New Roman, Courier New, Georgia i Helvetica
są znakami towarowymi ich właścicieli i służą wyłącznie do wskazania zgodności metryk.
