# ==================================================================
# PDF-BILDCHECK: Auflösung und Farbraum eingebetteter Bilder
# ==================================================================
# Prüft, mit welcher effektiven Auflösung und in welchem Farbraum Bilder
# in einer PDF liegen. Gedacht für den Vorher/Nachher-Vergleich rund um
# die OCR-Verarbeitung: "Wurde mein Foto heruntergerechnet?"
#
# NUTZUNG:
#   python pdf_bildcheck.py <datei-oder-ordner> [weitere ...]
#   python pdf_bildcheck.py D:\Skripte\OCR-Problem
#
# Ohne Argument wird nach einem Pfad gefragt.
#
# Benötigt: pip install pymupdf
# ==================================================================

import os
import re
import sys
import glob
import math
from typing import Optional

# ==================================================================
# Konsolen-Encoding (UTF-8) - muss VOR jedem print stehen
# ==================================================================
# Die Ausgabe verwendet Haken- und Warnzeichen; ohne diesen Block bricht
# sie unter der Windows-Standardcodepage mit UnicodeEncodeError ab.
try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import fitz  # PyMuPDF
except ImportError:
    print("❌ PyMuPDF fehlt.  Installation:  pip install pymupdf")
    sys.exit(1)


# Schwellen analog zu 5_OCR_PDF.py
PRINT_SAFE_MIN_DPI    = 300
PRINT_SAFE_MIN_PIXELS = 500_000

_CS_NAME_COMPONENTS = {
    "/DeviceGray": 1, "/CalGray": 1, "/G": 1,
    "/DeviceRGB": 3, "/CalRGB": 3, "/Lab": 3, "/RGB": 3,
    "/DeviceCMYK": 4, "/CMYK": 4,
}


def _cs_components_from_text(doc, txt: Optional[str], depth: int) -> int:
    if not txt or depth > 3:
        return 0
    if "/DeviceCMYK" in txt or "/CMYK" in txt:
        return 4
    if "/Separation" in txt or "/DeviceN" in txt:
        return 4
    if "/DeviceRGB" in txt or "/CalRGB" in txt:
        return 3
    if "/DeviceGray" in txt or "/CalGray" in txt:
        return 1
    m = re.search(r"/N\s+(\d+)", txt)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)\s+0\s+R", txt)
    if m:
        target = int(m.group(1))
        try:
            k, v = doc.xref_get_key(target, "N")
            if k == "int":
                return int(v)
            return _cs_components_from_text(
                doc, doc.xref_object(target, compressed=False), depth + 1)
        except Exception:
            return 0
    return 0


def image_color_components(doc, xref: int) -> int:
    """Farbkanäle eines Bildes, ohne den Bildstrom zu dekodieren."""
    try:
        kind, val = doc.xref_get_key(xref, "ColorSpace")
    except Exception:
        return 0
    if kind == "null" or not val:
        return 0
    if kind == "name":
        return _CS_NAME_COMPONENTS.get(val, 0)
    if kind == "xref":
        m = re.match(r"(\d+)", val)
        if not m:
            return 0
        target = int(m.group(1))
        try:
            k, v = doc.xref_get_key(target, "N")
            if k == "int":
                return int(v)
            return _cs_components_from_text(
                doc, doc.xref_object(target, compressed=False), 1)
        except Exception:
            return 0
    return _cs_components_from_text(doc, val, 1)


def _farbraum_name(n: int) -> str:
    return {1: "Graustufen", 3: "RGB", 4: "CMYK/Sonderfarbe"}.get(n, "unbekannt")


def check_pdf(path: str) -> None:
    size_mb = os.path.getsize(path) / 1024 / 1024
    print("\n" + "=" * 78)
    print(f"  {os.path.basename(path)}   ({size_mb:.2f} MB)")
    print("=" * 78)

    try:
        doc = fitz.open(path)
    except Exception as e:
        print(f"  ❌ Nicht lesbar: {type(e).__name__}: {e}")
        return

    max_dpi        = 0.0
    min_dpi_relevant = None
    has_cmyk       = False
    n_images       = 0
    seiten_bitmaps = 0

    with doc:
        meta = doc.metadata or {}
        if meta.get("subject"):
            print(f"  Subject: {meta['subject'][:70]}")
        if meta.get("producer"):
            print(f"  Producer: {meta['producer'][:70]}")
        print(f"  Seiten: {len(doc)}")
        print()
        print(f"  {'Seite':>5}  {'Pixel':>13}  {'platziert':>13}  {'dpi':>6}  Farbraum")
        print("  " + "-" * 72)

        for pno in range(len(doc)):
            page = doc.load_page(pno)
            prect = page.rect
            for img in page.get_images(full=False):
                try:
                    xref = int(img[0])
                    w, h = int(img[2] or 0), int(img[3] or 0)
                except (IndexError, TypeError, ValueError):
                    continue
                if w * h < PRINT_SAFE_MIN_PIXELS:
                    continue
                n_images += 1
                comps = image_color_components(doc, xref)
                if comps >= 4:
                    has_cmyk = True

                # Bildkanten ueber die Platzierungsmatrix vermessen, nicht
                # ueber das umschliessende Rechteck: bei einem um 90 Grad
                # gedrehten Bild liegt die Pixelbreite entlang der
                # Rechteckhoehe. Gemessen: 1000x500 px gedreht in 3x6 Zoll
                # platziert = 166,7 dpi; die alte Rechnung max(w/Breite,
                # h/Hoehe) meldete 333 dpi. Die Matrix bildet das
                # Einheitsquadrat des Bildes auf die Seite ab; die Laengen
                # ihrer Zeilen (a, b) und (c, d) sind die platzierten Kanten
                # in pt - unabhaengig von Drehung und Scherung.
                try:
                    platzierungen = page.get_image_rects(xref, transform=True)
                except Exception:
                    platzierungen = []
                if not platzierungen:
                    print(f"  {pno+1:>5}  {w:>6}x{h:<6}  {'—':>13}  {'—':>6}  {_farbraum_name(comps)}")
                    continue
                for r, mat in platzierungen:
                    wi = math.hypot(mat.a, mat.b) / 72.0
                    hi = math.hypot(mat.c, mat.d) / 72.0
                    if wi <= 0 or hi <= 0:
                        continue
                    dpi = max(w / wi, h / hi)
                    max_dpi = max(max_dpi, dpi)
                    if min_dpi_relevant is None or dpi < min_dpi_relevant:
                        min_dpi_relevant = dpi
                    # Seitenfüllend? -> Hinweis auf gerasterte Seite
                    fuellt = (r.width >= prect.width * 0.98
                              and r.height >= prect.height * 0.98)
                    if fuellt:
                        seiten_bitmaps += 1
                    flag = "  ⚠️ unter 300 dpi" if dpi < PRINT_SAFE_MIN_DPI else ""
                    ganz = "  [seitenfüllend]" if fuellt else ""
                    print(f"  {pno+1:>5}  {w:>6}x{h:<6}  "
                          f"{r.width/72*25.4:>5.0f}x{r.height/72*25.4:<5.0f} mm  "
                          f"{dpi:>6.0f}  {_farbraum_name(comps)}{flag}{ganz}")

    print()
    if n_images == 0:
        print("  Keine nennenswerten Bilder gefunden (nur Vektor/Text?).")
        return

    print(f"  Höchste Auflösung:  {max_dpi:>6.0f} dpi")
    if min_dpi_relevant is not None:
        print(f"  Niedrigste:         {min_dpi_relevant:>6.0f} dpi")
    print(f"  CMYK/Sonderfarben:  {'ja' if has_cmyk else 'nein'}")

    if seiten_bitmaps:
        print()
        print(f"  ⚠️  {seiten_bitmaps} seitenfüllende(s) Bitmap(s) gefunden.")
        print("     Typisch für eine Seite, die komplett neu gerastert wurde –")
        print("     Vektorgrafik und Text sind dann keine Vektoren mehr.")

    if max_dpi < PRINT_SAFE_MIN_DPI:
        print()
        print(f"  ⚠️  Kein Bild erreicht {PRINT_SAFE_MIN_DPI} dpi – für professionellen")
        print("     Druck in der Regel zu wenig.")
    else:
        print()
        print("  ✓ Mindestens ein Bild erfüllt die 300-dpi-Grenze.")

    if has_cmyk:
        print("  ✓ CMYK/Sonderfarben vorhanden – die OCR-Automatisierung")
        print("    verarbeitet solche Dateien bildschonend (keine RGB-Wandlung).")


def collect(arg: str):
    if os.path.isdir(arg):
        # glob.escape: eckige Klammern im Ordnernamen ('Akte [2024]') sind
        # sonst Zeichenklassen - gemessen: 0 statt 1 Treffer.
        return sorted(glob.glob(os.path.join(glob.escape(arg), "**", "*.pdf"), recursive=True))
    if os.path.isfile(arg) and arg.lower().endswith(".pdf"):
        return [arg]
    return []


def main() -> None:
    args = sys.argv[1:]
    if not args:
        print("\nPDF-BILDCHECK – Auflösung und Farbraum eingebetteter Bilder")
        try:
            raw = input("\nDatei oder Ordner: ").strip().strip('"').strip("'")
        except (EOFError, KeyboardInterrupt):
            return
        args = [raw] if raw else []

    dateien = []
    for a in args:
        gefunden = collect(a.strip().strip('"').strip("'"))
        if not gefunden:
            print(f"  ❌ Nichts gefunden: {a}")
        dateien.extend(gefunden)

    if not dateien:
        return

    for f in dateien:
        try:
            check_pdf(f)
        except Exception as e:
            print(f"\n  ❌ Fehler bei {f}: {type(e).__name__}: {e}")

    print("\n" + "=" * 78)
    print(f"  {len(dateien)} Datei(en) geprüft.")
    try:
        input("\nEnter zum Beenden ...")
    except Exception:
        pass


if __name__ == "__main__":
    main()
