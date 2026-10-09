#!/usr/bin/env python3
"""Darmowe zamienniki krojów Office o IDENTYCZNYCH szerokościach liter → assets/fonts/doc/.

Calibri → Carlito, Cambria → Caladea, Arial → Arimo, Times New Roman → Tinos,
Courier New → Cousine, Georgia → Gelasio. Na telefonie / tablecie / Macu bez Office
dokument łamie wiersze i strony tak jak w Wordzie (zamiast Roboto / Noto / Arial).
Gdzie oryginał jest zainstalowany, zostaje oryginał (local() przed url()).

Sprawdzone szerokości (hmtx, 215 znaków łaciny z polskimi): Arimo/Tinos/Cousine/Carlito
0 różnic, Caladea z crosextrafonts-20130214 0 różnic (Caladea z Google Fonts ma inne
cyfry — NIE jest zgodna z Cambrią), Gelasio ≈ Georgia (różne tylko ¸ ¯ ¼ ½ ¾).

Na odmianę DWA pliki (unicode-range): „core” — łacina z polskimi i innymi znakami
środkowoeuropejskimi + typowe symbole (~20 KB, apka zapisuje je na urządzeniu z góry),
„ext” — greka, cyrylica, reszta łaciny i symboli (pobierany dopiero, gdy dokument go użyje).
Bez hintingu (szerokości projektowe, jak liczy Word). Wewnętrzne nazwy krojów zmienione
(„DWB …”) — Carlito ma w licencji zastrzeżoną nazwę, a pliki są przycięte. Prawa autorskie
i licencja zostają w plikach; teksty licencji w assets/fonts/doc/.

Pliki są niezmienne (sw.js trzyma je w osobnym, niewersjonowanym cache) — przy zmianie
zawartości zmień SUFFIX, żeby adresy były nowe.

Wymaga: pip install fonttools brotli
Uruchom: python3 scripts/gen-doc-fonts.py   (źródła w ~/.cache/dwb-doc-fonts)
Skrypt nadpisuje też blok @font-face w styles/app.css (między znacznikami doc-fonts).
"""
import io
import re
import sys
import tarfile
import urllib.request
from pathlib import Path

from fontTools.subset import Options, Subsetter, parse_unicodes
from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "assets" / "fonts" / "doc"
CSS = ROOT / "styles" / "app.css"
CACHE = Path.home() / ".cache" / "dwb-doc-fonts"
GF = "https://raw.githubusercontent.com/google/fonts/main/"
CROS = "https://commondatastorage.googleapis.com/chromeos-localmirror/distfiles/crosextrafonts-20130214.tar.gz"
SUFFIX = "1"

STYLES = [("", 400, "normal"), ("b", 700, "normal"), ("i", 400, "italic"), ("bi", 700, "italic")]

# klon → (nazwa wewnętrzna, {styl: (źródło, wght dla krojów zmiennych)}, metryki pionowe do wyrównania)
CLONES = {
    "carlito": ("DWB Sans C", {
        "": ("gf:ofl/carlito/Carlito-Regular.ttf", None), "b": ("gf:ofl/carlito/Carlito-Bold.ttf", None),
        "i": ("gf:ofl/carlito/Carlito-Italic.ttf", None), "bi": ("gf:ofl/carlito/Carlito-BoldItalic.ttf", None)}, None),
    "caladea": ("DWB Serif C", {
        "": ("cros:Caladea-Regular.ttf", None), "b": ("cros:Caladea-Bold.ttf", None),
        "i": ("cros:Caladea-Italic.ttf", None), "bi": ("cros:Caladea-BoldItalic.ttf", None)}, None),
    "arimo": ("DWB Sans A", {
        "": ("gf:ofl/arimo/Arimo[wght].ttf", 400), "b": ("gf:ofl/arimo/Arimo[wght].ttf", 700),
        "i": ("gf:ofl/arimo/Arimo-Italic[wght].ttf", 400), "bi": ("gf:ofl/arimo/Arimo-Italic[wght].ttf", 700)}, None),
    "tinos": ("DWB Serif T", {
        "": ("gf:ofl/tinos/Tinos-Regular.ttf", None), "b": ("gf:ofl/tinos/Tinos-Bold.ttf", None),
        "i": ("gf:ofl/tinos/Tinos-Italic.ttf", None), "bi": ("gf:ofl/tinos/Tinos-BoldItalic.ttf", None)}, None),
    "cousine": ("DWB Mono C", {
        "": ("gf:ofl/cousine/Cousine-Regular.ttf", None), "b": ("gf:ofl/cousine/Cousine-Bold.ttf", None),
        "i": ("gf:ofl/cousine/Cousine-Italic.ttf", None), "bi": ("gf:ofl/cousine/Cousine-BoldItalic.ttf", None)}, None),
    # Gelasio ma wyższe/niższe ascent/descent niż Georgia — wyrównujemy (położenie linii bazowej w wierszu)
    "gelasio": ("DWB Serif G", {
        "": ("gf:ofl/gelasio/Gelasio[wght].ttf", 400), "b": ("gf:ofl/gelasio/Gelasio[wght].ttf", 700),
        "i": ("gf:ofl/gelasio/Gelasio-Italic[wght].ttf", 400), "bi": ("gf:ofl/gelasio/Gelasio-Italic[wght].ttf", 700)},
        (1878 / 2048, 449 / 2048)),
}

CORE = (
    "U+0000-017F,U+0218-021B,U+02C6-02C7,U+02D8-02DD,U+2000-206F,U+20AC,U+2116,U+2122,"
    "U+2190-2193,U+2212,U+2215,U+221E,U+2248,U+2260,U+2264-2265,U+25A0-25A1,U+25AA-25AB,"
    "U+25CB,U+25CF,U+2713,U+FB01-FB02,U+FEFF,U+FFFD"
)
FULL = (
    "U+0000-024F,U+0259,U+02B0-02FF,U+0300-036F,U+0370-03FF,U+0400-04FF,"
    "U+1E00-1EFF,U+2000-206F,U+2070-209F,U+20A0-20C0,U+2100-214F,U+2150-218F,"
    "U+2190-21FF,U+2200-22FF,U+2300-23FF,U+2500-257F,U+25A0-25FF,U+2600-26FF,"
    "U+2700-27BF,U+FB00-FB06,U+FEFF,U+FFFD"
)

# Nazwy lokalne (pełna nazwa + nazwa PostScript) — local() w Chrome/Safari/Firefox pasuje po nich.
LOCALS = {
    "Calibri": {"": ["Calibri"], "b": ["Calibri Bold", "Calibri-Bold"], "i": ["Calibri Italic", "Calibri-Italic"], "bi": ["Calibri Bold Italic", "Calibri-BoldItalic"]},
    "Carlito": {"": ["Carlito", "Carlito Regular", "Carlito-Regular"], "b": ["Carlito Bold", "Carlito-Bold"], "i": ["Carlito Italic", "Carlito-Italic"], "bi": ["Carlito Bold Italic", "Carlito-BoldItalic"]},
    "Cambria": {"": ["Cambria"], "b": ["Cambria Bold", "Cambria-Bold"], "i": ["Cambria Italic", "Cambria-Italic"], "bi": ["Cambria Bold Italic", "Cambria-BoldItalic"]},
    "Caladea": {"": ["Caladea", "Caladea Regular", "Caladea-Regular"], "b": ["Caladea Bold", "Caladea-Bold"], "i": ["Caladea Italic", "Caladea-Italic"], "bi": ["Caladea Bold Italic", "Caladea-BoldItalic"]},
    "Arial": {"": ["Arial", "ArialMT"], "b": ["Arial Bold", "Arial-BoldMT"], "i": ["Arial Italic", "Arial-ItalicMT"], "bi": ["Arial Bold Italic", "Arial-BoldItalicMT"]},
    "Helvetica": {"": ["Helvetica"], "b": ["Helvetica Bold", "Helvetica-Bold"], "i": ["Helvetica Oblique", "Helvetica-Oblique"], "bi": ["Helvetica Bold Oblique", "Helvetica-BoldOblique"]},
    "Liberation Sans": {"": ["Liberation Sans", "LiberationSans"], "b": ["Liberation Sans Bold", "LiberationSans-Bold"], "i": ["Liberation Sans Italic", "LiberationSans-Italic"], "bi": ["Liberation Sans Bold Italic", "LiberationSans-BoldItalic"]},
    "Arimo": {"": ["Arimo", "Arimo Regular", "Arimo-Regular"], "b": ["Arimo Bold", "Arimo-Bold"], "i": ["Arimo Italic", "Arimo-Italic"], "bi": ["Arimo Bold Italic", "Arimo-BoldItalic"]},
    "Times New Roman": {"": ["Times New Roman", "TimesNewRomanPSMT"], "b": ["Times New Roman Bold", "TimesNewRomanPS-BoldMT"], "i": ["Times New Roman Italic", "TimesNewRomanPS-ItalicMT"], "bi": ["Times New Roman Bold Italic", "TimesNewRomanPS-BoldItalicMT"]},
    "Times": {"": ["Times Roman", "Times-Roman"], "b": ["Times Bold", "Times-Bold"], "i": ["Times Italic", "Times-Italic"], "bi": ["Times Bold Italic", "Times-BoldItalic"]},
    "Liberation Serif": {"": ["Liberation Serif", "LiberationSerif"], "b": ["Liberation Serif Bold", "LiberationSerif-Bold"], "i": ["Liberation Serif Italic", "LiberationSerif-Italic"], "bi": ["Liberation Serif Bold Italic", "LiberationSerif-BoldItalic"]},
    "Tinos": {"": ["Tinos", "Tinos Regular", "Tinos-Regular"], "b": ["Tinos Bold", "Tinos-Bold"], "i": ["Tinos Italic", "Tinos-Italic"], "bi": ["Tinos Bold Italic", "Tinos-BoldItalic"]},
    "Courier New": {"": ["Courier New", "CourierNewPSMT"], "b": ["Courier New Bold", "CourierNewPS-BoldMT"], "i": ["Courier New Italic", "CourierNewPS-ItalicMT"], "bi": ["Courier New Bold Italic", "CourierNewPS-BoldItalicMT"]},
    "Courier": {"": ["Courier"], "b": ["Courier Bold", "Courier-Bold"], "i": ["Courier Oblique", "Courier-Oblique"], "bi": ["Courier Bold Oblique", "Courier-BoldOblique"]},
    "Liberation Mono": {"": ["Liberation Mono", "LiberationMono"], "b": ["Liberation Mono Bold", "LiberationMono-Bold"], "i": ["Liberation Mono Italic", "LiberationMono-Italic"], "bi": ["Liberation Mono Bold Italic", "LiberationMono-BoldItalic"]},
    "Cousine": {"": ["Cousine", "Cousine Regular", "Cousine-Regular"], "b": ["Cousine Bold", "Cousine-Bold"], "i": ["Cousine Italic", "Cousine-Italic"], "bi": ["Cousine Bold Italic", "Cousine-BoldItalic"]},
    "Georgia": {"": ["Georgia"], "b": ["Georgia Bold", "Georgia-Bold"], "i": ["Georgia Italic", "Georgia-Italic"], "bi": ["Georgia Bold Italic", "Georgia-BoldItalic"]},
    "Gelasio": {"": ["Gelasio", "Gelasio Regular", "Gelasio-Regular"], "b": ["Gelasio Bold", "Gelasio-Bold"], "i": ["Gelasio Italic", "Gelasio-Italic"], "bi": ["Gelasio Bold Italic", "Gelasio-BoldItalic"]},
    "Calibri Light": {"": ["Calibri Light", "Calibri-Light"], "i": ["Calibri Light Italic", "Calibri-LightItalic"]},
    "Aptos": {"": ["Aptos"], "b": ["Aptos Bold", "Aptos-Bold"], "i": ["Aptos Italic", "Aptos-Italic"], "bi": ["Aptos Bold Italic", "Aptos-BoldItalic"]},
}

# rodzina w CSS → (klon, kolejność nazw lokalnych, czy plik „ext”)
FAMILIES = [
    # Kroje Office: z urządzenia TYLKO oryginał. Lokalne kopie zamienników bywają w innej
    # wersji — Caladea z Google Fonts (2023+) ma inne cyfry niż Cambria, a stała u Mateusza
    # w ~/Library/Fonts i przejmowała Cambrię. Nasz plik jest sprawdzony.
    ("Calibri", "carlito", ["Calibri"], True),
    ("Cambria", "caladea", ["Cambria"], True),
    ("Arial", "arimo", ["Arial"], True),
    ("Times New Roman", "tinos", ["Times New Roman"], True),
    ("Courier New", "cousine", ["Courier New"], True),
    ("Georgia", "gelasio", ["Georgia"], True),
    # te same szerokości pod innymi nazwami (pliki z LibreOffice, Maca, PDF)
    ("Carlito", "carlito", ["Carlito"], False),
    # „Caladea” w pliku: Word pokazuje Cambrię (PDF z Worda, „Punkt Marzeny…” 2026-10-09: panose
    # = Cambria, osadzona Cambria) — lokalna Caladea z Google Fonts jest o ~6% szersza i
    # przestawiała wiersze; z urządzenia tylko Cambria, potem nasz plik (szerokości Cambrii)
    ("Caladea", "caladea", ["Cambria"], False),
    ("Helvetica", "arimo", ["Helvetica", "Arial"], False),
    ("Liberation Sans", "arimo", ["Liberation Sans", "Arimo", "Arial"], False),
    ("Arimo", "arimo", ["Arimo"], False),
    ("Times", "tinos", ["Times", "Times New Roman"], False),
    ("Liberation Serif", "tinos", ["Liberation Serif", "Tinos", "Times New Roman"], False),
    ("Tinos", "tinos", ["Tinos"], False),
    ("Courier", "cousine", ["Courier", "Courier New"], False),
    ("Liberation Mono", "cousine", ["Liberation Mono", "Cousine", "Courier New"], False),
    ("Cousine", "cousine", ["Cousine"], False),
    ("Gelasio", "gelasio", ["Gelasio"], False),
    # bez darmowego odpowiednika — najbliższy: Calibri Light ≈ Carlito (1,3% szerszy),
    # Aptos ≈ Arial/Arimo zmniejszony do 96% (sam Arial 3–5% szerszy od Aptosa w PDF-ie z Worda,
    # Carlito 5–6% węższy — pomiar npm run word:compare 2026-10-05). Prawdziwy Aptos z urządzenia
    # dokłada app/docx-viewer.js (FontFace z local) — tu bez local("Aptos"), bo size-adjust
    # zmniejszyłby i jego.
    ("Calibri Light", "carlito", ["Calibri Light"], True),
    ("Aptos", "arimo", ["Arial"], True),
]
SIZE_ADJUST = {"Aptos": "96%"}


def cp_set(spec):
    return set(parse_unicodes(spec))


def ranges(cps):
    out, cps = [], sorted(cps)
    i = 0
    while i < len(cps):
        j = i
        while j + 1 < len(cps) and cps[j + 1] == cps[j] + 1:
            j += 1
        out.append(f"U+{cps[i]:04X}" if i == j else f"U+{cps[i]:04X}-{cps[j]:04X}")
        i = j + 1
    return ",".join(out)


def source(src):
    CACHE.mkdir(parents=True, exist_ok=True)
    kind, rel = src.split(":", 1)
    dest = CACHE / f"{kind}_{rel.replace('/', '_')}"
    if dest.exists():
        return dest
    if kind == "gf":
        url = GF + urllib.request.quote(rel)
        print(f"  ↓ {url}")
        with urllib.request.urlopen(url) as r:
            dest.write_bytes(r.read())
    else:
        print(f"  ↓ {CROS}")
        with urllib.request.urlopen(CROS) as r:
            with tarfile.open(fileobj=io.BytesIO(r.read()), mode="r:gz") as tar:
                for m in tar.getmembers():
                    if m.name.endswith(".ttf"):
                        (CACHE / f"cros_{Path(m.name).name}").write_bytes(tar.extractfile(m).read())
    return dest


def rename(font, family, style_key):
    sub = {"": "Regular", "b": "Bold", "i": "Italic", "bi": "Bold Italic"}[style_key]
    full = f"{family} {sub}"
    ps = f"{family.replace(' ', '')}-{sub.replace(' ', '')}"
    names = font["name"]
    names.names = [n for n in names.names if n.nameID not in (1, 2, 3, 4, 6, 16, 17, 21, 22, 25)]
    for nid, val in ((1, family), (2, sub), (3, f"{ps};{SUFFIX}"), (4, full), (6, ps)):
        names.setName(val, nid, 3, 1, 0x409)
        names.setName(val, nid, 1, 0, 0)


def set_vertical(font, asc, desc):
    upm = font["head"].unitsPerEm
    a, d = round(asc * upm), round(desc * upm)
    font["hhea"].ascent, font["hhea"].descent, font["hhea"].lineGap = a, -d, 0
    os2 = font["OS/2"]
    os2.sTypoAscender, os2.sTypoDescender, os2.sTypoLineGap = a, -d, 0
    os2.usWinAscent, os2.usWinDescent = a, d


def subset_to(font_path, wght, cps, out, internal, style_key, vertical):
    font = TTFont(font_path)
    if wght is not None:
        font = instantiateVariableFont(font, {"wght": wght})
    opts = Options()
    opts.flavor = "woff2"
    opts.hinting = False
    opts.layout_features = ["*"]
    opts.name_IDs = ["*"]
    opts.name_languages = ["*"]
    opts.notdef_outline = True
    opts.drop_tables += ["DSIG", "hdmx", "LTSH", "VDMX", "STAT"]
    sub = Subsetter(options=opts)
    sub.populate(unicodes=cps)
    sub.subset(font)
    rename(font, internal, style_key)
    if vertical:
        set_vertical(font, *vertical)
    font.flavor = "woff2"
    font.save(out)
    return out.stat().st_size


def build_files():
    OUT.mkdir(parents=True, exist_ok=True)
    core, full = cp_set(CORE), cp_set(FULL)
    ext = full - core
    sizes = {"core": 0, "ext": 0}
    for key, (internal, styles, vertical) in CLONES.items():
        for style_key, _, _ in STYLES:
            src, wght = styles[style_key]
            path = source(src)
            tag = style_key or "r"
            sizes["core"] += subset_to(path, wght, core, OUT / f"{key}-{tag}-core-{SUFFIX}.woff2", internal, style_key, vertical)
            sizes["ext"] += subset_to(path, wght, ext, OUT / f"{key}-{tag}-ext-{SUFFIX}.woff2", internal, style_key, vertical)
    print(f"  core {sizes['core'] / 1024:.0f} KB · ext {sizes['ext'] / 1024:.0f} KB")
    return ranges(core), ranges(ext)


def css_block(core_range, ext_range):
    lines = [
        "/* >>> doc-fonts — WYGENEROWANE przez scripts/gen-doc-fonts.py, nie edytuj ręcznie.",
        "   Kroje Office spoza urządzenia → darmowe zamienniki o tych samych szerokościach liter",
        "   (assets/fonts/doc, licencje OFL/Apache obok plików). Najpierw oryginał z urządzenia",
        "   (local), dopiero gdy go nie ma — plik zamiennika. Dzięki temu na telefonie i Macu bez",
        "   Office wiersze i strony łamią się jak w Wordzie. Każda odmiana ma prawdziwy plik",
        "   (bez sztucznego pogrubiania/pochylania); „ext” (greka, cyrylica…) pobiera się tylko,",
        "   gdy dokument ma takie znaki. */",
    ]
    for family, clone, local_order, has_ext in FAMILIES:
        for style_key, weight, style in STYLES:
            locs = []
            for name in local_order:
                locs += LOCALS.get(name, {}).get(style_key, [])
            if family == "Calibri Light" and style_key not in LOCALS["Calibri Light"]:
                continue  # pogrubienie Calibri Light = przeglądarka pogrubi lżejszy krój, jak Word
            tag = style_key or "r"
            src_local = ", ".join(f'local("{n}")' for n in dict.fromkeys(locs))
            parts = [("core", core_range)] + ([("ext", ext_range)] if has_ext else [])
            for part, rng in parts:
                url = f'url("../assets/fonts/doc/{clone}-{tag}-{part}-{SUFFIX}.woff2") format("woff2")'
                src = f"{src_local}, {url}" if src_local else url
                adjust = f" size-adjust: {SIZE_ADJUST[family]};" if family in SIZE_ADJUST else ""
                lines.append(
                    f'@font-face {{ font-family: "{family}"; font-style: {style}; font-weight: {weight}; '
                    f"font-display: swap; src: {src};{adjust} unicode-range: {rng}; }}"
                )
    lines.append("/* <<< doc-fonts */")
    return "\n".join(lines)


def write_css(block):
    css = CSS.read_text(encoding="utf-8")
    pattern = re.compile(r"/\* >>> doc-fonts.*?/\* <<< doc-fonts \*/", re.S)
    if not pattern.search(css):
        raise SystemExit("styles/app.css: brak znaczników /* >>> doc-fonts … /* <<< doc-fonts */")
    CSS.write_text(pattern.sub(lambda _: block, css), encoding="utf-8")
    print(f"  ✅ styles/app.css — blok doc-fonts ({block.count('@font-face')} reguł)")


def main():
    for old in OUT.glob("*.woff2"):
        old.unlink()
    core_range, ext_range = build_files()
    write_css(css_block(core_range, ext_range))


if __name__ == "__main__":
    sys.exit(main())
