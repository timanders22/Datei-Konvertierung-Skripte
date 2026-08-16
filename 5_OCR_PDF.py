# ==================================================================
# PDF-OCR-AUTOMATISIERUNG: MASCHINELLE TEXTERKENNUNG
# ==================================================================
# Stand: 12.06.2026
# ANLEITUNG:
# 1. Python installieren: https://www.python.org/
# 2. Erforderliche Programme installieren:
#    a. Tesseract OCR (inkl. DEU+ENG+FRA+SPA)
#       → https://github.com/UB-Mannheim/tesseract/wiki
#    b. Ghostscript (optional ab ocrmypdf >= 17)
#       → https://ghostscript.com/releases/gsdnld.html
#    c. Python-Bibliotheken:
#       pip install ocrmypdf pikepdf pymupdf psutil pypdfium2 pywin32 tqdm watchdog
#    d. Optimierungs-Tools (optional, in C:\OCR):
#       jbig2.exe  pngquant.exe  unpaper.exe  leptonica-*.dll
#    e. Java (JRE) und VeraPDF (am besten in C:\OCR)
#       → https://verapdf.org/software/
#    f. Windows neu starten
# 3. In der Konsole: python 5_OCR_PDF.py
#
# --watch / --daemon:  Hotfolder-Modus
# --verify-markers:    Vorhandene PDF/A-Marker gegen veraPDF prüfen
#                      und bei Abweichung korrigieren/neu verarbeiten
# --cleanup-backups:   Zu Beginn Backup-Leichen rekursiv löschen
#                      (Default: Nein – spart Zeit auf großen Ablagen)
# --clear-mru:         Recent-Documents-Liste leeren (Default: Aus)
# --no-initial-scan:   Watch-Modus: vorhandene PDFs beim Start NICHT
#                      einreihen (Default: Initial-Scan aktiv)
# --print-safe:        Bildschonendes Profil für ALLE Dateien erzwingen
# --no-print-safe:     Automatische Erkennung druckrelevanter Bilder aus
#
# BILDQUALITÄT: PDFs mit Bildern ab 300 dpi oder mit CMYK-/Separations-
# farbräumen werden automatisch bildschonend verarbeitet (kein Neurastern
# der Seite, keine Neukodierung der Bilder, keine RGB-Konvertierung).
# Siehe Kommentarblock bei PRINT_SAFE_OPTIONS.
#
# HINWEIS FÜR PyInstaller-Build (Umwandlung in .exe, Windows CMD, eine Zeile):
#   python -m PyInstaller --onefile --console --icon="python_icon.ico" --name 5_OCR_PDF --collect-all ocrmypdf --collect-all pikepdf --collect-binaries pypdfium2 --collect-submodules win32com --hidden-import win32timezone --hidden-import pywintypes --hidden-import win32api --hidden-import win32file --hidden-import win32process --hidden-import watchdog --hidden-import watchdog.observers.polling --hidden-import fitz --hidden-import psutil --hidden-import tqdm --hidden-import PIL --hidden-import PIL.ImageFile 5_OCR_PDF.py
#
# ==================================================================

import os
import sys
import time
import shutil
import io
import re
import csv
import uuid
import hashlib
import tempfile
import traceback
import contextlib
import platform
import subprocess
import logging
import logging.handlers
import winreg
import argparse
import queue
import signal
import threading
import multiprocessing
import concurrent.futures
import ctypes
from concurrent.futures import ProcessPoolExecutor
from concurrent.futures.process import BrokenProcessPool
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, List, Dict, Any, NamedTuple, Set

# ==================================================================
# Konsolen-Encoding (UTF-8) - muss VOR jedem print stehen
# ==================================================================
# Ohne diesen Block bricht die erste Ausgabe mit Rahmenzeichen oder Emoji
# unter der Windows-Standardcodepage (cp850/cp1252) mit UnicodeEncodeError
# ab - und zwar sowohl im Konsolenfenster als auch bei Umleitung in eine
# Datei. Betroffen waren hier 169 print-Zeilen, die erste davon in der
# Hinweismeldung zu pypdfium2, also noch vor der ersten PDF.
try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ==================================================================
# Gemeinsame Grundbibliothek (mit Rueckfall)
# ==================================================================
# Fehlt _gemeinsam.py - etwa weil nur dieses eine Skript weitergegeben
# wurde -, laeuft alles unveraendert weiter; die daraus bedienten
# Zusatzfunktionen schalten sich dann still ab.
try:
    _eigener_ordner = os.path.dirname(os.path.abspath(__file__))
    if _eigener_ordner not in sys.path:
        sys.path.insert(0, _eigener_ordner)
    import _gemeinsam as gem
except Exception:
    gem = None

os.environ["PATH"] = r"C:\OCR;" + r"C:\OCR\bin;" + os.environ.get("PATH", "")

# Tesseract nutzt intern OpenMP und startet sonst je Instanz mehrere
# Threads (typisch 4-8). Bei ProcessPoolExecutor mit N Workern multipli-
# ziert sich das zu N*Threads und ueberlastet die CPU (Thermal Throttling,
# gegenseitige Blockade, Timeout-Haeufung). OMP_THREAD_LIMIT=1 zwingt
# jede Tesseract-Instanz auf genau einen Thread; die Parallelitaet kommt
# allein aus dem Worker-Pool. Muss VOR dem Laden von ocrmypdf/Tesseract
# gesetzt werden und wird via spawn an die Worker (und deren Tesseract-
# Subprozesse) vererbt.
os.environ["OMP_THREAD_LIMIT"] = "1"

# ==================================================================
# Externe Bibliotheken
# ==================================================================
try:
    import ocrmypdf
    from ocrmypdf.exceptions import PriorOcrFoundError
    import fitz
    try:
        fitz.TOOLS.mupdf_display_errors(False)
    except Exception:
        pass
    import psutil
    from tqdm import tqdm
    from PIL import ImageFile, Image
    ImageFile.LOAD_TRUNCATED_IMAGES = True
    Image.MAX_IMAGE_PIXELS = None  # Decompression-Bomb-Schutz aus (Enterprise-Pläne)
except ImportError as e:
    print(f"❌ FEHLENDE MODULE: {e}")
    print("Bitte installieren: pip install ocrmypdf pikepdf pymupdf psutil pypdfium2 pywin32 tqdm watchdog Pillow")
    sys.exit(1)

_DIGITAL_SIG_FALLBACK = False
try:
    from ocrmypdf.exceptions import DigitalSignatureError
except ImportError:
    _DIGITAL_SIG_FALLBACK = True
    class DigitalSignatureError(Exception):
        pass

_INPUT_FILE_ERR_FALLBACK = False
try:
    from ocrmypdf.exceptions import InputFileError
except ImportError:
    _INPUT_FILE_ERR_FALLBACK = True
    class InputFileError(Exception):
        pass

try:
    import pypdfium2
    HAS_PYPDFIUM2 = True
except ImportError:
    HAS_PYPDFIUM2 = False
    print("⚠️  pypdfium2 nicht installiert (empfohlen): pip install pypdfium2")

try:
    from watchdog.observers import Observer
    from watchdog.observers.polling import PollingObserver
    from watchdog.events import FileSystemEventHandler
    HAS_WATCHDOG = True
except ImportError:
    HAS_WATCHDOG = False

    class FileSystemEventHandler:
        pass

    class Observer:
        def schedule(self, *a, **kw): pass
        def start(self): pass
        def stop(self): pass
        def join(self, timeout=None): pass
        def is_alive(self): return False

    class PollingObserver(Observer):
        pass


# ==================================================================
# Prozess-Kontext (Main vs. Worker)
# ==================================================================

_IS_WORKER = multiprocessing.current_process().name != "MainProcess"


# ==================================================================
# Plattform-Konstanten (Subprocess ohne Fensterblitz)
# ==================================================================

if sys.platform == "win32":
    _CREATE_NO_WINDOW = 0x08000000
    _SUBPROCESS_FLAGS: Dict[str, Any] = {"creationflags": _CREATE_NO_WINDOW}
else:
    _SUBPROCESS_FLAGS = {}


# ==================================================================
# Protokoll-Objekte fruehzeitig binden
# ==================================================================
# Zahlreiche except-Zweige weiter unten 'verwerfen' Ausnahmen mit
# detail_logger.debug(...). Diese Zweige sind aber schon waehrend des Imports
# erreichbar - lange bevor _setup_logging() am Ende der Datei die Namen
# belegt. Der Fehlerschlucker loeste dann selbst einen NameError aus, der
# NICHT abgefangen wurde: das Skript startete gar nicht, und die Meldung
# zeigte auf den falschen Fehler.
# Am eindeutigsten ist _ensure_bom(): es wird von _setup_logging() selbst
# aufgerufen, dort KANN der Name konstruktionsbedingt noch nicht existieren.
# Deshalb hier bereits gueltige Logger-Objekte binden; _setup_logging()
# ergaenzt spaeter nur noch die Handler und Stufen.
error_logger  = logging.getLogger("ErrorLogger")
detail_logger = logging.getLogger("DetailLogger")


def _fmt_exc(e: BaseException) -> str:
    return f"{type(e).__name__}: {e}"


def _run(cmd: List[str], timeout: Optional[int] = None) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd,
        capture_output=True, text=True,
        encoding="utf-8", errors="ignore",
        timeout=timeout,
        **_SUBPROCESS_FLAGS,
    )


# ==================================================================
# EXE-Verzeichnis und DLL-Suchpfade
# ==================================================================

def _get_exe_dir() -> str:
    if getattr(sys, "frozen", False):
        try:
            return os.path.dirname(sys.executable)
        except Exception as _e:
            detail_logger.debug(f"_get_exe_dir: Exception verworfen: {_e!r}")
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def _get_meipass_dir() -> Optional[str]:
    return getattr(sys, "_MEIPASS", None)


_DLL_DIR_HANDLES: List[Any] = []


def _register_dll_dirs(paths: List[str]) -> None:
    if not hasattr(os, "add_dll_directory"):
        return
    for p in paths:
        try:
            if p and os.path.isdir(p):
                handle = os.add_dll_directory(p)
                if handle is not None:
                    _DLL_DIR_HANDLES.append(handle)
        except Exception as _e:
            detail_logger.debug(f"_register_dll_dirs: Exception verworfen: {_e!r}")


# ==================================================================
# Absicherung der Werkzeugverzeichnisse
# ==================================================================
# C:\OCR und C:\OCR\bin stehen laut Zeile 102 in PATH VOR allen System-
# verzeichnissen, bilden die obersten Kandidaten der Werkzeugsuche
# (tesseract.exe, gswin64c.exe, jbig2.exe, verapdf.bat) und landen ueber
# EXTRA_DLL_DIRS im DLL-Suchpfad des Prozesses. Ein direkt unter C:\
# angelegter Ordner erbt aber die Standard-ACL von C:\ und ist damit fuer
# JEDEN authentifizierten Benutzer beschreibbar. Wird das Skript 'als
# Administrator' gestartet - wozu der Hinweis bei der Zielauswahl
# ausdruecklich einlaedt -, kann ein Standardbenutzer dort vorab eine
# eigene tesseract.exe oder leptonica-*.dll ablegen und damit Code im
# Administratorkontext ausfuehren: lokale Rechteausweitung.

# Rechte, mit denen sich in einem Verzeichnis eine Programmdatei
# unterschieben oder eine vorhandene ersetzen laesst.
_GEFAEHRLICHE_RECHTE = (
    0x00000002      # FILE_ADD_FILE / FILE_WRITE_DATA
    | 0x00000004    # FILE_ADD_SUBDIRECTORY / FILE_APPEND_DATA
    | 0x00000040    # FILE_DELETE_CHILD
    | 0x00010000    # DELETE
    | 0x00040000    # WRITE_DAC
    | 0x00080000    # WRITE_OWNER
    | 0x10000000    # GENERIC_ALL
    | 0x40000000    # GENERIC_WRITE
)

# Konten, denen Schreibrecht auf ein Programmverzeichnis zusteht. Als SID
# statt als Name, weil die Klartextnamen sprachabhaengig sind
# ('Authentifizierte Benutzer' / 'Authenticated Users').
_VERTRAUTE_SIDS = {
    "S-1-5-18",      # SYSTEM
    "S-1-5-32-544",  # Administratoren
    "S-1-3-0",       # ERSTELLER-BESITZER (greift nur auf selbst erzeugte Dateien)
    "S-1-3-4",       # Besitzerrechte
}


def _sid_klartext(sid: Any, rueckfall: str) -> str:
    try:
        import win32security
        name, domaene, _typ = win32security.LookupAccountSid(None, sid)
        return f"{domaene}\\{name}" if domaene else name
    except Exception:
        return rueckfall


def _unsichere_schreiber(pfad: str) -> List[str]:
    """Konten nennen, die in `pfad` eine Programmdatei unterschieben koennen.

    Leere Liste heisst unbedenklich. Ist die ACL nicht lesbar, ebenfalls
    leer - lieber kein Befund als ein Fehlalarm, der den Lauf blockiert.
    Inherit-only-ACEs werden mitgezaehlt: sie gelten zwar nicht fuer den
    Ordner selbst, vererben sich aber auf die Dateien darin und sind
    damit genau der Weg, eine vorhandene .exe zu ersetzen.
    """
    try:
        import win32security
    except Exception:
        return []
    try:
        sd = win32security.GetFileSecurity(
            pfad, win32security.DACL_SECURITY_INFORMATION)
        dacl = sd.GetSecurityDescriptorDacl()
    except Exception as _e:
        detail_logger.debug(f"_unsichere_schreiber({pfad}): Exception verworfen: {_e!r}")
        return []
    if dacl is None:
        # Fehlende DACL bedeutet Vollzugriff fuer jeden.
        return ["<ohne DACL – Vollzugriff für alle>"]

    ACCESS_ALLOWED_ACE_TYPE        = 0
    ACCESS_ALLOWED_OBJECT_ACE_TYPE = 5
    treffer: List[str] = []
    for i in range(dacl.GetAceCount()):
        try:
            ace = dacl.GetAce(i)
            ace_typ = ace[0][0]
            if ace_typ not in (ACCESS_ALLOWED_ACE_TYPE, ACCESS_ALLOWED_OBJECT_ACE_TYPE):
                continue
            maske = ace[1]
            sid   = ace[-1]
            if not (maske & _GEFAEHRLICHE_RECHTE):
                continue
            sid_text = win32security.ConvertSidToStringSid(sid)
            if sid_text in _VERTRAUTE_SIDS:
                continue
            # Dienstkonten (S-1-5-80-*, u.a. TrustedInstaller) sind unkritisch.
            if sid_text.startswith("S-1-5-80-"):
                continue
            treffer.append(_sid_klartext(sid, sid_text))
        except Exception as _e:
            detail_logger.debug(f"_unsichere_schreiber: ACE {i} verworfen: {_e!r}")
    return treffer


def _ist_erhoeht() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _pruefe_werkzeugverzeichnisse() -> None:
    """Werkzeugverzeichnisse vor der ersten Programmausfuehrung pruefen.

    Bei erhoehten Rechten ist ein beschreibbares Werkzeugverzeichnis eine
    Rechteausweitung - dann Abbruch. Ohne erhoehte Rechte bleibt der
    Angreifer auf dem Rechteniveau des Anwenders; dann genuegt eine
    deutliche Warnung, damit der Lauf auf den Ablagen nicht ausfaellt.
    """
    if os.name != "nt":
        return
    befunde: Dict[str, List[str]] = {}
    for pfad in EXTRA_DLL_DIRS:
        if not os.path.isdir(pfad):
            continue
        schreiber = _unsichere_schreiber(pfad)
        if schreiber:
            befunde[pfad] = sorted(set(schreiber))
    if not befunde:
        return

    for pfad, konten in befunde.items():
        detail_logger.error(
            f"Unsicheres Werkzeugverzeichnis {pfad}: Schreibrecht für {', '.join(konten)}")

    print()
    print("=" * 72)
    print("  ⚠  Werkzeugverzeichnis ist für normale Benutzer beschreibbar")
    print("=" * 72)
    for pfad, konten in befunde.items():
        print(f"  {pfad}")
        for k in konten:
            print(f"        Schreibrecht: {k}")
    print()
    print("  Aus diesem Verzeichnis werden tesseract.exe, gswin64c.exe,")
    print("  jbig2.exe und verapdf.bat gestartet und DLLs geladen – mit")
    print("  Vorrang vor C:\\Program Files. Wer dort schreiben darf, kann")
    print("  dem Skript ein eigenes Programm unterschieben.")
    print()
    print("  Reparatur (Eingabeaufforderung als Administrator):")
    # Unterordner eines ohnehin genannten Ordners weglassen: sie erben die
    # neuen Rechte. Ein eigener /inheritance:r-Aufruf wuerde die Vererbung
    # dort unnoetig wieder kappen.
    wurzeln = [
        p for p in befunde
        if not any(o != p and p.lower().startswith(o.lower().rstrip("\\") + "\\")
                   for o in befunde)
    ]
    for pfad in wurzeln:
        print(f'      icacls "{pfad}" /inheritance:r'
              f' /grant *S-1-5-32-544:(OI)(CI)F'
              f' /grant *S-1-5-18:(OI)(CI)F'
              f' /grant *S-1-5-32-545:(OI)(CI)RX')
    print()
    if _ist_erhoeht():
        print("  Dieser Lauf hat erhöhte Rechte. Ein untergeschobenes Programm")
        print("  liefe damit als Administrator – das ist eine lokale Rechte-")
        print("  ausweitung. ABBRUCH.")
        print("=" * 72)
        print()
        detail_logger.error(
            "Abbruch: erhöhte Rechte bei beschreibbarem Werkzeugverzeichnis.")
        sys.exit(2)
    print("  Dieser Lauf hat keine erhöhten Rechte; ein untergeschobenes")
    print("  Programm käme nicht über Ihre eigenen Rechte hinaus. Der Lauf")
    print("  wird fortgesetzt – die Rechte bitte trotzdem korrigieren.")
    print("=" * 72)
    print()


# ==================================================================
# Windows Known Folders (Desktop, Downloads, Documents)
# ==================================================================

class _GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]


def _str_to_guid(s: str) -> _GUID:
    s = s.strip("{}").replace("-", "")
    g = _GUID()
    g.Data1 = int(s[0:8], 16)
    g.Data2 = int(s[8:12], 16)
    g.Data3 = int(s[12:16], 16)
    for i in range(8):
        g.Data4[i] = int(s[16 + i * 2:18 + i * 2], 16)
    return g


FOLDERID_DESKTOP   = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
FOLDERID_DOWNLOADS = "{374DE290-123F-4565-9164-39C4925E467B}"
FOLDERID_DOCUMENTS = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"


def _get_known_folder(folder_id_guid: str) -> Optional[str]:
    if sys.platform != "win32":
        return None
    try:
        guid = _str_to_guid(folder_id_guid)
        SHGetKnownFolderPath = ctypes.windll.shell32.SHGetKnownFolderPath
        SHGetKnownFolderPath.argtypes = [
            ctypes.POINTER(_GUID),
            ctypes.c_ulong,
            ctypes.c_void_p,
            ctypes.POINTER(ctypes.c_wchar_p),
        ]
        SHGetKnownFolderPath.restype = ctypes.c_long

        path_ptr = ctypes.c_wchar_p()
        hr = SHGetKnownFolderPath(ctypes.byref(guid), 0, None, ctypes.byref(path_ptr))
        if hr == 0 and path_ptr.value:
            path = path_ptr.value
            try:
                ctypes.windll.ole32.CoTaskMemFree(path_ptr)
            except Exception as _e:
                detail_logger.debug(f"_get_known_folder: Exception verworfen: {_e!r}")
            return path
    except Exception as _e:
        detail_logger.debug(f"_get_known_folder: Exception verworfen: {_e!r}")
    return None


def _get_user_shell_folder_from_registry(name: str) -> Optional[str]:
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        ) as key:
            val, _ = winreg.QueryValueEx(key, name)
            expanded = os.path.expandvars(val)
            if expanded and os.path.isdir(expanded):
                return expanded
    except Exception as _e:
        detail_logger.debug(f"_get_user_shell_folder_from_registry: Exception verworfen: {_e!r}")
    return None


def _get_desktop_folder() -> str:
    p = _get_known_folder(FOLDERID_DESKTOP)
    if p and os.path.isdir(p):
        return p
    p = _get_user_shell_folder_from_registry("Desktop")
    if p:
        return p
    up = os.environ.get("USERPROFILE", "")
    if up:
        d = os.path.join(up, "Desktop")
        if os.path.isdir(d):
            return d
    d = os.path.join(os.path.expanduser("~"), "Desktop")
    return d if os.path.isdir(d) else os.path.expanduser("~")


def _get_downloads_folder() -> str:
    p = _get_known_folder(FOLDERID_DOWNLOADS)
    if p and os.path.isdir(p):
        return p
    p = _get_user_shell_folder_from_registry("{374DE290-123F-4565-9164-39C4925E467B}")
    if p:
        return p
    up = os.environ.get("USERPROFILE", "")
    if up:
        d = os.path.join(up, "Downloads")
        if os.path.isdir(d):
            return d
    d = os.path.join(os.path.expanduser("~"), "Downloads")
    return d if os.path.isdir(d) else os.path.expanduser("~")


# ==================================================================
# Anwendungsdaten-Verzeichnis (deterministisch für EXE-Betrieb)
# ==================================================================

def _get_app_data_dir() -> str:
    base = (
        os.environ.get("LOCALAPPDATA")
        or os.environ.get("APPDATA")
        or tempfile.gettempdir()
    )
    d = os.path.join(base, "PDF-OCR-Automatisierung")
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        d = tempfile.gettempdir()
    return d


def _get_documents_folder() -> str:
    p = _get_known_folder(FOLDERID_DOCUMENTS)
    if p and os.path.isdir(p):
        return p
    userprofile = os.environ.get("USERPROFILE", "")
    if userprofile:
        docs = os.path.join(userprofile, "Documents")
        if os.path.isdir(docs):
            return docs
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    if os.path.isdir(docs):
        return docs
    return os.environ.get("TEMP", os.environ.get("TMP", r"C:\Temp"))


# ==================================================================
# Konfiguration
# ==================================================================

APP_DATA_DIR = _get_app_data_dir()


def _get_log_dir() -> str:
    # Bevorzugt Skript-/EXE-Verzeichnis; Fallback APP_DATA_DIR bei fehlenden Schreibrechten
    candidate = _get_exe_dir()
    try:
        test_path = os.path.join(candidate, ".log_write_test")
        with open(test_path, "wb") as fh:
            fh.write(b"x")
        os.remove(test_path)
        return candidate
    except Exception:
        return APP_DATA_DIR


LOG_DIR = _get_log_dir()

TEMP_DIR = os.path.join(
    _get_documents_folder(),
    f"5_OCR_PDF_Processing_{os.getpid()}"
)
MAX_PATH_LEN      = 200
MAX_RETRIES       = 3
RETRY_DELAY       = 2

_RUN_TIMESTAMP     = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE           = os.path.join(LOG_DIR, f"5_OCR_PDF_errors_{_RUN_TIMESTAMP}.log")
DETAILED_LOG_FILE  = os.path.join(LOG_DIR, f"5_OCR_PDF_detailed_{_RUN_TIMESTAMP}.log")
CSV_LOG_FILE       = os.path.join(LOG_DIR, "5_OCR_PDF_results.csv")
RESUME_FILE_PREFIX = os.path.join(LOG_DIR, "5_OCR_PDF_resume")
RUN_SUMMARY_FILE   = os.path.join(LOG_DIR, "5_OCR_PDF_last_run.txt")
# Der Probelauf schreibt in eine EIGENE Datei. Sonst ueberschrieb er die
# Zusammenfassung des letzten ECHTEN Laufs - und meldete dort "Status: OK"
# fuer einen Lauf, der gar nichts geschrieben hat.
RUN_SUMMARY_DRYRUN = os.path.join(LOG_DIR, "5_OCR_PDF_last_dryrun.txt")

LOG_MAX_BYTES     = 10 * 1024 * 1024
LOG_BACKUP_COUNT  = 5

UTF8_BOM = b'\xef\xbb\xbf'

DEFAULT_LANGUAGE   = "deu"
FALLBACK_LANGUAGES = ["eng", "fra", "spa"]

PDF_RASTERIZER = "auto"

# Default auch fuer --auto: bewusst PDF/A-2u (Archiv-Default) -- vorher
# fiel --auto ohne --output-type still auf Standard-PDF zurueck, waehrend
# der interaktive Modus PDF/A-2u empfahl.
DEFAULT_OUTPUT_TYPE = "pdfa-2u"

OCR_OPTIONS: Dict[str, Any] = {
    "skip_text":               True,
    "deskew":                  True,
    "rotate_pages":            True,
    "rotate_pages_threshold":  0.5,
    "fast_web_view":           1_000_000,
    "optimize":                1,
    "clean":                   True,
    "jbig2_lossy":             False,
    "png_quality":             75,
    "unpaper_args":            "--no-blurfilter",
}

# ==================================================================
# Bildschonendes Profil (Druckvorlagen)
# ==================================================================
# HINTERGRUND – warum es dieses Profil gibt:
#
# ocrmypdf setzt intern lossless_reconstruction=False, sobald eine der
# Optionen deskew, clean_final, force_ocr oder remove_background aktiv ist
# (siehe ocrmypdf/_validation.py: set_lossless_reconstruction). Ist die
# verlustfreie Rekonstruktion aus, wird JEDE Seite, die tatsaechlich OCR
# durchlaeuft, komplett neu gerastert und als ein einziges Bitmap in die
# Ausgabe geschrieben (_pipelines/_common.py: create_pdf_page_from_image).
#
# Die Rasteraufloesung stammt aus calculate_image_dpi(). Liegt das
# Verhaeltnis von durchschnittlicher zu maximaler Bild-DPI unter 0.8 –
# also immer dann, wenn eine Seite ein hochaufgeloestes Foto UND ein
# niedrig aufgeloestes Element (Logo, Strichgrafik) enthaelt – rastert
# ocrmypdf mit dem GEWICHTETEN Mittelwert. Das hochaufgeloeste Foto wird
# dabei heruntergerechnet.
#
# Gemessen an einer Testseite (A4, Foto mit 600 dpi + Logo mit 72 dpi):
#   deskew=True  -> eine Seiten-Bitmap mit 374 dpi, Foto nur noch 2210 px
#                   statt 3543 px (38 % Aufloesungsverlust), Vektoren weg
#   deskew=False -> beide Bilder unveraendert, Vektorinhalt bleibt Vektor
#
# Zusaetzlich rekodiert Ghostscript im PDF/A-Schritt Bilder standardmaessig
# nach JPEG (-dAutoFilterColorImages=true, -dJPEGQ=95) – auch verlustfrei
# gespeicherte. pdfa_image_compression="lossless" unterbindet das.
#
# Und: ocrmypdf rekodiert ab Ghostscript 10.6.0 selbst bei optimize=1 jedes
# eingebettete JPEG mit Qualitaet 75 neu (optimize.py: _should_optimize_jpeg).
# Nur optimize=0 schaltet den Optimierer vollstaendig ab.

# Ab dieser Aufloesung gilt ein eingebettetes Bild als druckrelevant.
# 300 dpi ist die uebliche Untergrenze professioneller Druckereien.
PRINT_SAFE_MIN_DPI = 300

# Mindestgroesse in Pixeln, ab der ein Bild ueberhaupt betrachtet wird –
# verhindert, dass winzige Icons mit hoher rechnerischer DPI das Profil
# ausloesen.
PRINT_SAFE_MIN_PIXELS = 500_000

# Optionen, die im bildschonenden Profil ueberschrieben werden.
PRINT_SAFE_OPTIONS: Dict[str, Any] = {
    "deskew":       False,   # verhindert das Neurastern der Seite
    "clean":        False,   # unpaper wuerde ebenfalls in die Seite eingreifen
    "unpaper_args": None,
    "optimize":     0,       # kein JPEG-Requantisieren durch ocrmypdf
    "pdfa_image_compression": "lossless",   # kein JPEG-Requantisieren durch Ghostscript
    "color_conversion_strategy": "LeaveColorUnchanged",
}

# OCR-Marker in PDF-Metadaten
MARKER_KEY           = "subject"
MARKER_VALUE_PDF     = "OCR-verarbeitet (Standard-PDF)"
MARKER_VALUE_PDFA_2U = "OCR-verarbeitet (PDF/A-2u)"
MARKER_VALUE_PDFA_2B = "OCR-verarbeitet (PDF/A-2b)"
LEGACY_MARKER        = "OCR-verarbeitet"

# Verzeichnis-Presets fuer die Startauswahl. Hier die im eigenen Netz
# gebraeuchlichen Laufwerke/Shares eintragen. Format: (Pfad, Hinweistext)
DIRECTORY_PRESETS: Tuple[Tuple[str, str], ...] = (
    ("Q:\\",                 ""),
    ("R:\\",                 ""),
    ("G:\\Geteilte Ablagen", ""),
    ("G:\\Meine Ablage",     ""),
    ("\\\\server\\dfs",      "Platzhalter – eigenen Dateiserver eintragen"),
)

EXCLUDE_DIRS: Set[str] = {
    "$RECYCLE.BIN", "System Volume Information", "Windows",
    "Program Files", "Program Files (x86)", "AppData",
    "Temp", "tmp", ".git", ".svn", ".vs", "__pycache__",
}
EXCLUDE_DIRS_LOWER: Set[str] = {d.lower() for d in EXCLUDE_DIRS}
EXCLUDE_FILES: Set[str] = {"thumbs.db", "desktop.ini", ".ds_store"}

WARN_KEYWORDS = [
    "error", "exception", "lots of diacritics", "poor ocr",
    "corrupt", "warning", "failed", "aborted",
]

COLOR_SPACE_ERR_KEYWORDS = [
    "unusual color space",
    "color-conversion-strategy",
    "color conversion",
    "colorspace",
]

TEMP_BASENAME_MAX = 80
MAX_PAGE_DIMENSION = 10_000

# has_text ist rein informativ (die OCR-Entscheidung trifft skip_text);
# der teure get_text-Durchlauf wird daher auf die ersten N Seiten begrenzt.
PDF_INFO_TEXT_SCAN_PAGES = 50

# Maximale Pixelzahl eines einzelnen eingebetteten Bildes vor dem OCR-Versuch.
# Pillow wirft per Default DecompressionBombError ab ca. 178 956 970 Pixel.
MAX_EMBEDDED_IMAGE_PIXELS = 178_956_970

# Magic-Number-Such-Bereich (Header-Müll tolerieren)
PDF_MAGIC_SCAN_BYTES = 1024

# Memory-Schätzung (realistisch für seitenweise ocrmypdf-Verarbeitung)
MEMORY_MB_PER_PAGE_MIN = 25
MEMORY_BASE_MB_FLOOR   = 300
MEMORY_BASE_MB_CAP     = 2000
MEMORY_MB_PER_JOB      = 200

# Disk-Space: Mindest-Overhead (MB) zzgl. Faktor × PDF-Größe
DISK_MIN_FREE_MB   = 100
DISK_FACTOR_PER_MB = 3

# AV-Scanner-Toleranz beim Temp-Kopieren
COPY_RETRIES               = 5
COPY_RETRY_DELAY_BASE      = 1.0
COPY_AV_STABILIZE_DELAY    = 0.3
COPY_AV_LOCK_WAIT_SECONDS  = 15

# Watch-Modus: Heartbeat alle N Sekunden
HEARTBEAT_INTERVAL = 300

# Watch-Modus: maximale Retry-Anzahl bei Lock (je 5s → 30 Min)
MAX_LOCK_RETRIES = 360

# Watch-Modus: Deduplizierungsfenster für _enqueue (Sekunden)
ENQUEUE_DEDUP_WINDOW = 30.0

# veraPDF: Validierungs-Timeout (Basis + pro Seite)
VERAPDF_BASE_TIMEOUT     = 60
VERAPDF_PER_PAGE_SECONDS = 0.5
VERAPDF_MAX_TIMEOUT      = 600

# Worker-Timeout (Watch + parallel)
WORKER_TIMEOUT = 3600

# Pool-Reset-Schutz gegen Endlosschleifen bei systemischen Fehlern
MAX_POOL_GENERATIONS = 500
MAX_FILE_REQUEUES    = 2

# Backup-Leichen beim Start: Mindest-Alter in Stunden
BACKUP_CLEANUP_MIN_AGE_HOURS = 1

# Zusätzliche DLL-Suchpfade (Leptonica, Tesseract-Abhängigkeiten)
EXTRA_DLL_DIRS: List[str] = [
    r"C:\OCR",
    r"C:\OCR\bin",
    r"C:\Program Files\Tesseract-OCR",
    r"C:\Program Files (x86)\Tesseract-OCR",
    r"C:\Tesseract-OCR",
]


# ==================================================================
# Logging (Dual-Logger + CSV + Worker-Puffer)
# ==================================================================

class _LogEntry(NamedTuple):
    level:    str
    context:  str
    message:  str
    exc_info: bool = False


class _ProcessResult(NamedTuple):
    status:        str
    detail:        str
    messages:      List[str]
    log_entries:   List[_LogEntry]
    duration_sec:  float
    size_mb:       float
    pages:         int
    error_msg:     Optional[str]
    file_path:     str
    size_after_mb: float = 0.0


def _ensure_bom(path: str) -> None:
    try:
        if not os.path.exists(path):
            with open(path, "wb") as fh:
                fh.write(UTF8_BOM)
            return
        if os.path.getsize(path) == 0:
            with open(path, "wb") as fh:
                fh.write(UTF8_BOM)
            return
        with open(path, "rb") as fh:
            head = fh.read(3)
        if head.startswith(UTF8_BOM):
            return
        with open(path, "rb") as fh:
            data = fh.read()
        with open(path, "wb") as fh:
            fh.write(UTF8_BOM)
            fh.write(data)
    except Exception as _e:
        detail_logger.debug(f"_ensure_bom: Exception verworfen: {_e!r}")


def _setup_logging(
    log_file: str = LOG_FILE,
    detail_file: str = DETAILED_LOG_FILE,
) -> Tuple[logging.Logger, logging.Logger]:
    _ensure_bom(log_file)
    _ensure_bom(detail_file)

    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    el = logging.getLogger("ErrorLogger")
    el.handlers.clear()
    fh = logging.handlers.RotatingFileHandler(
        log_file, mode="a", encoding="utf-8",
        maxBytes=LOG_MAX_BYTES, backupCount=LOG_BACKUP_COUNT,
    )
    fh.setFormatter(fmt)
    el.setLevel(logging.ERROR)
    el.addHandler(fh)
    el.propagate = False

    dl = logging.getLogger("DetailLogger")
    dl.handlers.clear()
    dh = logging.handlers.RotatingFileHandler(
        detail_file, mode="a", encoding="utf-8",
        maxBytes=LOG_MAX_BYTES, backupCount=LOG_BACKUP_COUNT,
    )
    dh.setFormatter(fmt)
    dl.setLevel(logging.DEBUG)
    dl.addHandler(dh)
    dl.propagate = False

    return el, dl


if not _IS_WORKER:
    error_logger, detail_logger = _setup_logging()
else:
    error_logger  = logging.getLogger("ErrorLogger")
    detail_logger = logging.getLogger("DetailLogger")
    if not error_logger.handlers:
        error_logger.addHandler(logging.NullHandler())
    if not detail_logger.handlers:
        detail_logger.addHandler(logging.NullHandler())


# Worker-Puffer
_worker_entries_buffer: List[_LogEntry] = []
_worker_entries_lock = threading.Lock()


def _worker_buffer_drain() -> List[_LogEntry]:
    with _worker_entries_lock:
        out = list(_worker_entries_buffer)
        _worker_entries_buffer.clear()
    return out


def _worker_buffer_clear() -> None:
    with _worker_entries_lock:
        _worker_entries_buffer.clear()


def log_error(context: str, message: str, exc_info: bool = False) -> None:
    if _IS_WORKER:
        msg = message
        if exc_info:
            msg += "\n" + traceback.format_exc()
        with _worker_entries_lock:
            _worker_entries_buffer.append(_LogEntry("ERROR", context, msg, False))
        return
    error_logger.error(f"[{context}] {message}", exc_info=exc_info)
    detail_logger.error(f"[{context}] {message}", exc_info=exc_info)


def log_warning(context: str, message: str) -> None:
    if _IS_WORKER:
        with _worker_entries_lock:
            _worker_entries_buffer.append(_LogEntry("WARN", context, message))
        return
    detail_logger.warning(f"[{context}] {message}")


def log_info(context: str, message: str) -> None:
    if _IS_WORKER:
        with _worker_entries_lock:
            _worker_entries_buffer.append(_LogEntry("INFO", context, message))
        return
    detail_logger.info(f"[{context}] {message}")


def log_debug(context: str, message: str) -> None:
    if _IS_WORKER:
        return
    detail_logger.debug(f"[{context}] {message}")


def _flush_log_entries(entries: List[_LogEntry]) -> None:
    for e in entries:
        if e.level == "ERROR":
            if _IS_WORKER:
                continue
            error_logger.error(f"[{e.context}] {e.message}", exc_info=e.exc_info)
            detail_logger.error(f"[{e.context}] {e.message}", exc_info=e.exc_info)
        elif e.level == "WARN":
            if _IS_WORKER:
                continue
            detail_logger.warning(f"[{e.context}] {e.message}")
        else:
            if _IS_WORKER:
                continue
            detail_logger.info(f"[{e.context}] {e.message}")


# ==================================================================
# CSV-Log
# ==================================================================

CSV_COLUMNS = [
    "timestamp", "path", "status", "detail",
    "duration_sec", "size_mb", "pages", "error", "size_after_mb",
]
_csv_lock = threading.Lock()


def _setup_csv_log(path: str = CSV_LOG_FILE) -> None:
    try:
        # Altes CSV-Format (ohne size_after_mb) wegrotieren, damit Header
        # und Datenzeilen konsistent bleiben.
        if os.path.exists(path) and os.path.getsize(path) > 0:
            try:
                with open(path, "r", encoding="utf-8-sig", newline="") as fh:
                    first_line = fh.readline()
                if "size_after_mb" not in first_line:
                    os.replace(path, path + ".old")
            except Exception:
                pass
        needs_header = (not os.path.exists(path)) or os.path.getsize(path) == 0
        if needs_header:
            with open(path, "wb") as fh:
                fh.write(UTF8_BOM)
            with open(path, "a", encoding="utf-8", newline="") as fh:
                writer = csv.writer(fh, delimiter=";")
                writer.writerow(CSV_COLUMNS)
    except Exception:
        pass



# ==================================================================
# Gemeinsames Laufprotokoll (migration.jsonl)
# ==================================================================
# Ergaenzt das skripteigene Protokoll, ersetzt es nicht. Erst damit
# laesst sich der Fortschritt ueber alle elf Schritte auswerten.
_laufprotokoll = None


def _protokoll(pfad: str, aktion: str, status: str, detail: str = "") -> None:
    global _laufprotokoll
    if gem is None:
        return
    try:
        if _laufprotokoll is None:
            _laufprotokoll = gem.Laufprotokoll(
                os.path.splitext(os.path.basename(__file__))[0])
        _laufprotokoll.schreibe(pfad, aktion, status, detail)
    except Exception:
        pass

def _write_csv_row(result: _ProcessResult, path: str = CSV_LOG_FILE) -> None:
    _protokoll(result.file_path, "OCR", result.status, result.detail)
    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        result.file_path,
        result.status,
        result.detail,
        f"{result.duration_sec:.2f}",
        f"{result.size_mb:.2f}",
        str(result.pages),
        (result.error_msg or "").replace("\n", " ").replace("\r", " "),
        f"{result.size_after_mb:.2f}",
    ]
    with _csv_lock:
        try:
            with open(path, "a", encoding="utf-8", newline="") as fh:
                writer = csv.writer(fh, delimiter=";")
                writer.writerow(row)
        except Exception:
            pass


def _write_error_csv(file_path: str, detail: str, error_msg: str) -> None:
    # CSV-Zeile fuer Fehler, die AUSSERHALB von process_pdf_file entstehen
    # (Worker-Timeout/Hard-Kill, Requeue-/Generation-Limit) -- ohne diesen
    # Helper fehlten solche Dateien im Ergebnis-CSV komplett.
    _write_csv_row(_ProcessResult(
        status="ERROR", detail=detail, messages=[], log_entries=[],
        duration_sec=0.0, size_mb=0.0, pages=0,
        error_msg=error_msg, file_path=file_path,
    ))


# ==================================================================
# Resume-Datei
# ==================================================================

_resume_lock = threading.Lock()


def get_resume_file_path(target_dir: str, output_type: str) -> str:
    key = hashlib.md5(
        f"{os.path.normcase(os.path.abspath(target_dir))}|{output_type}".encode("utf-8")
    ).hexdigest()[:12]
    return f"{RESUME_FILE_PREFIX}_{key}.txt"


def load_resume_set(resume_path: str) -> Set[str]:
    if not os.path.exists(resume_path):
        return set()
    try:
        with open(resume_path, "r", encoding="utf-8") as fh:
            return {line.strip() for line in fh if line.strip()}
    except Exception:
        return set()


def append_resume(resume_path: str, file_path: str) -> None:
    with _resume_lock:
        try:
            with open(resume_path, "a", encoding="utf-8") as fh:
                fh.write(file_path + "\n")
        except Exception as _e:
            detail_logger.debug(f"append_resume: Exception verworfen: {_e!r}")


def delete_resume_file(resume_path: str) -> None:
    try:
        if os.path.exists(resume_path):
            os.remove(resume_path)
    except Exception as _e:
        detail_logger.debug(f"delete_resume_file: Exception verworfen: {_e!r}")


# ==================================================================
# Graceful Shutdown
# ==================================================================

shutdown_event = threading.Event()
_shutdown_count = [0]


def _install_signal_handlers() -> None:
    def handler(sig, frame):
        _shutdown_count[0] += 1
        if _shutdown_count[0] == 1:
            shutdown_event.set()
            print("\n⚠️  Beende nach aktueller Datei … (Ctrl+C erneut für Sofort-Abbruch)")
        else:
            print("\n❌ Sofort-Abbruch!")
            os._exit(130)
    try:
        signal.signal(signal.SIGINT, handler)
    except Exception as _e:
        detail_logger.debug(f"_install_signal_handlers: Exception verworfen: {_e!r}")
    try:
        signal.signal(signal.SIGTERM, handler)
    except Exception as _e:
        detail_logger.debug(f"_install_signal_handlers: Exception verworfen: {_e!r}")


# ==================================================================
# MRU-Bereinigung (optional)
# ==================================================================

def clear_mru() -> None:
    try:
        SHARD_PIDL = 0x00000001
        ctypes.windll.shell32.SHAddToRecentDocs(SHARD_PIDL, None)
    except Exception as _e:
        detail_logger.debug(f"clear_mru: Exception verworfen: {_e!r}")


# ==================================================================
# Pfad-Hilfsfunktionen
# ==================================================================

def _lp(path: str) -> str:
    if not path or len(path) <= MAX_PATH_LEN:
        return path
    return prepare_long_path(path)


def prepare_long_path(path: str) -> str:
    if path.startswith("\\\\?\\"):
        return path
    path = os.path.abspath(os.path.normpath(path))
    if path.startswith("\\\\"):
        return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path


def _strip_long_path(path: str) -> str:
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\"):
        return path[4:]
    return path


def sanitize_path(raw: str) -> str:
    path = raw.strip().strip('"').strip("'")
    if len(path) == 2 and path[1] == ":":
        path += "\\"
    return path


def ask_directory() -> str:
    desktop_dir   = _get_desktop_folder()
    downloads_dir = _get_downloads_folder()

    entries: List[Tuple[str, str]] = [
        (path, path) for path, _hint in DIRECTORY_PRESETS
    ]
    entries.append((f"Desktop   ({desktop_dir})",   desktop_dir))
    entries.append((f"Downloads ({downloads_dir})", downloads_dir))
    custom = len(entries) + 1

    print("\nZielverzeichnis auswählen:")
    for i, (label, _p) in enumerate(entries, start=1):
        print(f"  [{i}] {label}")
    print(f"  [{custom}] Eigenen Pfad eingeben")
    print()
    print("  Hinweis: Bei Start 'als Administrator' sind Netzlaufwerks-Buchstaben")
    print("  oft nicht verbunden – dann den UNC-Pfad verwenden.")
    print()
    while True:
        try:
            choice = input(f"Auswahl [1–{custom}]: ").strip()
        except EOFError:
            print("\nKeine Eingabe möglich (EOF) – Abbruch.")
            sys.exit(1)

        if choice.isdigit() and 1 <= int(choice) <= len(entries):
            path = entries[int(choice) - 1][1]
            if os.path.isdir(path):
                return path
            print(f"  ❌ Nicht erreichbar: {path}")
            continue

        if choice == str(custom):
            try:
                raw = input("Pfad eingeben: ")
            except EOFError:
                print("\nKeine Eingabe möglich (EOF) – Abbruch.")
                sys.exit(1)
            path = sanitize_path(raw)
            if os.path.isdir(path):
                return path
            print(f"  ❌ Verzeichnis nicht gefunden: '{path}'")
        else:
            print(f"  Bitte 1–{custom} eingeben.")


def ask_yes_no(prompt: str, default_yes: bool = False) -> bool:
    hint = "[J/n]" if default_yes else "[j/N]"
    while True:
        try:
            answer = input(f"{prompt} {hint}: ").strip().lower()
        except EOFError:
            return default_yes
        if answer == "":
            return default_yes
        if answer in ("j", "ja", "y", "yes"):
            return True
        if answer in ("n", "nein", "no"):
            return False
        print("  Bitte 'j' oder 'n' eingeben.")


def ask_output_type() -> str:
    print("\n" + "—" * 70)
    print("  HINWEIS ZUR ARCHIVIERUNG (Stand 2026):")
    print("  Das Bundesarchiv empfiehlt die Formate in dieser Reihenfolge:")
    print("  PDF/A-2a → PDF/A-2u → PDF/A-2b → PDF/A-1a → PDF/A-1b")
    print()
    print("  Da PDF/A-2a (semantische Struktur) nachträglich kaum")
    print("  ohne manuelle Nachbearbeitung erreichbar ist, wird PDF/A-2u")
    print("  empfohlen.")
    print()
    print("  TECHNIK: Die tatsächliche Konformitätsstufe (2u vs. 2b) hängt")
    print("  von den ToUnicode-CMaps der Schriftarten ab. Wenn veraPDF")
    print("  installiert ist, wird das Ergebnis nach OCR automatisch")
    print("  validiert und der Metadaten-Marker entsprechend gesetzt.")
    print("  Ohne veraPDF wird konservativ höchstens PDF/A-2b markiert.")
    print("—" * 70)

    print("\n  Ausgabeformat wählen:")
    print("  [1] PDF/A-2u  (Unicode – EMPFOHLEN; benötigt veraPDF)")
    print("  [2] PDF/A-2b  (Basic – immer erreichbar)")
    print("  [3] PDF       (Standard – kleiner, nicht archivoptimiert)")

    while True:
        try:
            choice = input("\n  Format [Standard: 1]: ").strip()
        except EOFError:
            return "pdfa-2u"
        if choice in ("", "1"):
            return "pdfa-2u"
        elif choice == "2":
            return "pdfa-2b"
        elif choice == "3":
            return "pdf"
        else:
            print("  Bitte 1, 2 oder 3 eingeben.")


# ==================================================================
# Long-Path- und Datei-Hilfsfunktionen
# ==================================================================

def safe_exists(path: str) -> bool:
    try:
        return os.path.exists(_lp(path))
    except Exception:
        return False


# Original-System-Temp VOR jeder Umlenkung sichern: _worker_init_tempdir
# verbiegt TEMP/TMP auf das eigene Arbeitsverzeichnis - der Cleanup am
# Skriptende muss aber den ECHTEN Windows-Temp adressieren.
_ORIG_WINDOWS_TEMP = (
    os.environ.get("TEMP")
    or os.environ.get("TMP")
    or tempfile.gettempdir()
)

# Whitelist fuer den Windows-Temp-Cleanup: NUR eigene Artefakte und
# ocrmypdf-Reste. NIEMALS pauschal leeren: Fremdprozesse legen aktive
# Daten ohne Lock in %TEMP% ab; blindes Loeschen zerstoert sie.
# Office-/COM-Praefixe (gen_py, ~$, vbe, excel8.0, cvr*.tmp) gehoeren
# bewusst NICHT hierher: das sind Artefakte fremder Programme, und
# gen_py wird von laufenden pywin32-COM-Anwendungen aktiv genutzt.
_WINDOWS_TEMP_WHITELIST_PREFIXES = (
    "ocrmypdf.", "test_ocr", "5_ocr_pdf",
)

# 'ocrmypdf.'-Ordner im ECHTEN System-Temp koennen seit der Temp-Umlenkung
# (_worker_init_tempdir lenkt tempfile.tempdir und TMP/TEMP in das eigene,
# PID-getrennte Arbeitsverzeichnis) gar nicht mehr von diesem Lauf stammen -
# sie gehoeren also einem FREMDEN ocrmypdf-Prozess, dessen unkomprimierte
# Seitenbilder womoeglich gerade in Benutzung sind. Sie am Skriptende blind zu
# loeschen zerstoert fremde Zwischendaten. Vollstaendig streichen waere aber
# auch falsch: Reste ABGEBROCHENER frueherer Laeufe dieses Skripts (aus der
# Zeit vor der Umlenkung oder nach einem harten Kill) sind GB-gross und
# gehoeren geraeumt. Deshalb: nur, was schon vor dem Start dieses Laufs da war
# und seither nicht mehr angefasst wurde.
_LAUF_BEGINN = time.time()
_TEMP_MINDESTALTER_S = 3600.0


def _is_whitelisted_temp_entry(name: str) -> bool:
    nl = name.lower()
    return any(nl.startswith(p) for p in _WINDOWS_TEMP_WHITELIST_PREFIXES)


def _temp_eintrag_ist_verwaist(pfad: str) -> bool:
    """Wurde der Eintrag zuletzt vor dem Laufbeginn veraendert?"""
    try:
        mtime = os.path.getmtime(pfad)
    except Exception as _e:
        detail_logger.debug(f"_temp_eintrag_ist_verwaist: Exception verworfen: {_e!r}")
        return False
    return mtime < min(_LAUF_BEGINN, time.time() - _TEMP_MINDESTALTER_S)


def _cleanup_windows_temp() -> None:
    # Best-effort Bereinigung des ECHTEN Windows-Temp am Skriptende -
    # ausschliesslich per Whitelist bekannter Praefixe (ocrmypdf.io.*-
    # Zwischenordner, eigene Artefakte). Gesperrte/in-Nutzung-Dateien
    # werden stillschweigend uebersprungen.
    win_temp = _ORIG_WINDOWS_TEMP
    if not win_temp or not os.path.isdir(win_temp):
        return
    try:
        for name in os.listdir(win_temp):
            if not _is_whitelisted_temp_entry(name):
                continue
            full = os.path.join(win_temp, name)
            # Nur eindeutig verwaiste Reste anfassen - siehe Begruendung an
            # _temp_eintrag_ist_verwaist. Ein laufender Fremdprozess haelt
            # seine Zwischendaten aktuell und bleibt damit verschont.
            if not _temp_eintrag_ist_verwaist(full):
                detail_logger.debug(
                    f"System-Temp: '{name}' ist aktuell in Benutzung – nicht angetastet.")
                continue
            try:
                if os.path.isdir(full):
                    shutil.rmtree(full, ignore_errors=True)
                else:
                    os.remove(full)
            except Exception as _e:
                detail_logger.debug(f"_cleanup_windows_temp: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_cleanup_windows_temp: Exception verworfen: {_e!r}")


def _worker_init_tempdir(worker_temp_dir: str) -> None:
    # Wird beim Start jedes Worker-Prozesses aufgerufen (ProcessPool-
    # initializer). Lenkt sowohl Python's tempfile (das ocrmypdf intern
    # fuer seine ocrmypdf.io.*-Zwischenordner nutzt) als auch die an
    # Tesseract/Ghostscript vererbten TMP/TEMP-Env-Variablen in unser
    # kontrolliertes, PID-getrenntes Arbeitsverzeichnis um. Sonst landen
    # die unkomprimierten Seitenbilder im globalen System-Temp und bleiben
    # bei einem harten Worker-Kill (Timeout/Absturz) als GB-grosser Muell
    # liegen - und zwar an einem Ort, den wir nicht gefahrlos raeumen
    # koennen (andere Programme/parallele Skript-Instanzen nutzen denselben
    # System-Temp). In unserem temp_dir werden sie dagegen von
    # _purge_orphan_temp_files (Pool-Reset) und rmtree (Skriptende) erfasst.
    try:
        os.makedirs(worker_temp_dir, exist_ok=True)
        tempfile.tempdir = worker_temp_dir
        os.environ["TMP"]  = worker_temp_dir
        os.environ["TEMP"] = worker_temp_dir
    except Exception as _e:
        detail_logger.debug(f"_worker_init_tempdir: Exception verworfen: {_e!r}")

    # Restore-Privilegien (Admin-Kontext) fuer ACL/Owner-Erhalt aktivieren –
    # Token-Privilegien gelten pro Prozess, daher hier je Worker.
    _enable_restore_privileges()


def _purge_orphan_temp_files(temp_dir: str) -> int:
    # Nach einem Pool-Bruch/Timeout-Reset sind alle Worker tot und alle
    # Tasks requeued (mit neuen UUIDs bei Wiederholung). Die UUID-benannten
    # Temp-Dateien (in/out) der abgestuerzten Tasks sind damit verwaist:
    # ihr finally-Block im Worker lief nie. Sie werden zwar am Skriptende
    # per rmtree(temp_dir) entfernt, koennen aber bei langen Naechten mit
    # vielen Crashes die Platte fuellen. Daher hier zwischendurch raeumen.
    # Locked-Dateien (falls ein Worker doch noch haengt) werden still
    # uebersprungen.
    removed = 0
    # Sicherheits-Riegel: niemals den System-Temp pauschal leeren -
    # dort liegen aktive Daten fremder Prozesse ohne Lock.
    try:
        if (os.path.normcase(os.path.abspath(temp_dir))
                == os.path.normcase(os.path.abspath(_ORIG_WINDOWS_TEMP))):
            return 0
    except Exception as _e:
        detail_logger.debug(f"_purge_orphan_temp_files: Exception verworfen: {_e!r}")
    try:
        for name in os.listdir(_lp(temp_dir)):
            full = os.path.join(temp_dir, name)
            try:
                if os.path.isdir(_lp(full)):
                    # ocrmypdf.io.*-Zwischenordner mit unkomprimierten
                    # Seitenbildern - die GB-Fresser bei einem Crash.
                    shutil.rmtree(_lp(full), ignore_errors=True)
                    removed += 1
                elif os.path.isfile(_lp(full)) and safe_remove(full):
                    removed += 1
            except Exception as _e:
                detail_logger.debug(f"_purge_orphan_temp_files: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_purge_orphan_temp_files: Exception verworfen: {_e!r}")
    return removed


# ==================================================================
# NTFS-Sicherheitsuebernahme (ACL/Owner) bei Datei-Ersetzung
# ==================================================================
# Das Skript ersetzt Dateien, statt sie in-place zu aendern. Neu
# erzeugte Dateien erben nur die vererbbaren ACEs des Zielordners und
# gehoeren dem AUSFUEHRENDEN Konto. Laeuft das Skript unter einem
# Admin-Konto ueber Ablagen, deren Zugriff per explizitem Benutzer-ACE
# oder Owner-Mapping (NAS, Unix-Security-Style) geregelt ist, verliert
# der Nutzer dadurch den Zugriff. Daher: Owner/Group/DACL des Originals
# VOR der Ersetzung sichern und auf der neuen Datei wiederherstellen.
#
# Beide Ausfuehrungskontexte werden beruecksichtigt:
#  - Admin-Konto: SeRestore-/SeTakeOwnership-Privileg wird beim Start
#    aktiviert; Owner-Wiederherstellung funktioniert vollstaendig.
#  - Nutzer-Konto: Privileg-Aktivierung schlaegt still fehl; die neue
#    Datei gehoert ohnehin dem Nutzer. DACL-Restore funktioniert (der
#    Owner hat implizit WRITE_DAC), Owner-Restore auf sich selbst
#    ebenfalls; ein fremder Original-Owner wird toleriert und nur im
#    Detail-Log vermerkt.
# Alles best effort: ein fehlgeschlagener ACL-Restore bricht die
# Dateiverarbeitung nicht ab.

_restore_privileges_enabled = False


def _sec_path(path: str) -> str:
    return _lp(path)


def _enable_restore_privileges() -> None:
    global _restore_privileges_enabled
    try:
        import win32api
        import win32security
        htoken = win32security.OpenProcessToken(
            win32api.GetCurrentProcess(),
            win32security.TOKEN_ADJUST_PRIVILEGES | win32security.TOKEN_QUERY,
        )
        privs = []
        for name in ("SeRestorePrivilege", "SeTakeOwnershipPrivilege"):
            try:
                luid = win32security.LookupPrivilegeValue(None, name)
                privs.append((luid, win32security.SE_PRIVILEGE_ENABLED))
            except Exception as _e:
                detail_logger.debug(f"_enable_restore_privileges: Exception verworfen: {_e!r}")
        if not privs:
            return
        win32security.AdjustTokenPrivileges(htoken, 0, privs)
        if win32api.GetLastError() == 0:
            _restore_privileges_enabled = True
            log_debug("ACL", 
                "Restore-Privilegien aktiviert (Admin-Kontext) – "
                "Owner-Wiederherstellung vollstaendig verfuegbar.")
        else:
            log_debug("ACL", 
                "Restore-Privilegien nicht zugewiesen (Nutzer-Kontext) – "
                "Owner-Restore nur auf eigene Dateien moeglich.")
    except Exception as e:
        log_debug("ACL", f"Privileg-Aktivierung fehlgeschlagen: {e}")


def _get_security_descriptor(path: str):
    try:
        import win32security
        return win32security.GetNamedSecurityInfo(
            _sec_path(path),
            win32security.SE_FILE_OBJECT,
            win32security.OWNER_SECURITY_INFORMATION
            | win32security.GROUP_SECURITY_INFORMATION
            | win32security.DACL_SECURITY_INFORMATION,
        )
    except Exception as e:
        log_debug("ACL", f"Sicherheitsinfo nicht lesbar ({path}): {e}")
        return None


def _apply_security_descriptor(path: str, sd) -> None:
    """Eigentuemer, Gruppe und DACL einer ersetzten Datei wiederherstellen.

    Alles wird in EINEM SetNamedSecurityInfo-Aufruf gesetzt. Frueher liefen
    zwei getrennte Aufrufe (erst DACL, dann Owner) - und das Setzen des
    Eigentuemers ordnet die Vererbung neu. Nachgestellt: eine Datei verlor
    dabei die Kennzeichnung ihrer geerbten ACEs, und eine geerbte
    EIGENTUEMERRECHTE-ACE (S-1-3-4) bekam zusaetzlich INHERIT_ONLY - damit galt
    sie fuer die Datei selbst nicht mehr. Wer seinen Zugriff allein daraus
    bezog, konnte die eigene Datei anschliessend nicht mehr oeffnen
    (PermissionError). Das ist das Gegenteil dessen, was diese Funktion
    bezweckt. Ein gemeinsamer Aufruf laesst Windows die Rechte in einem Zug
    berechnen; der schaedliche Zwischenzustand entsteht gar nicht erst.

    Der Eigentuemer wird ausserdem nur gesetzt, wenn er tatsaechlich abweicht -
    ein privilegierter Schreibvorgang ohne Wirkung entfaellt damit.

    Schlaegt der gemeinsame Aufruf fehl (typisch: kein SeRestorePrivilege im
    Nutzer-Kontext, dann verweigert bereits das Owner-Feld), wird die DACL
    einzeln nachgezogen. Damit bleibt das bisherige Verhalten erhalten, dass
    wenigstens die Rechte ankommen.
    """
    if sd is None:
        return
    try:
        import win32security
    except Exception:
        return
    p = _sec_path(path)

    try:
        dacl = sd.GetSecurityDescriptorDacl()
    except Exception:
        dacl = None
    try:
        owner = sd.GetSecurityDescriptorOwner()
    except Exception:
        owner = None
    try:
        group = sd.GetSecurityDescriptorGroup()
    except Exception:
        group = None

    info = 0
    if dacl is not None:
        info |= win32security.DACL_SECURITY_INFORMATION
        try:
            ctrl, _rev = sd.GetSecurityDescriptorControl()
            if ctrl & win32security.SE_DACL_PROTECTED:
                info |= win32security.PROTECTED_DACL_SECURITY_INFORMATION
            else:
                info |= win32security.UNPROTECTED_DACL_SECURITY_INFORMATION
        except Exception as _e:
            log_debug("ACL", f"_apply_security_descriptor: Exception verworfen: {_e!r}")

    # Eigentuemer nur setzen, wenn er wirklich abweicht.
    if owner is not None:
        try:
            akt = win32security.GetNamedSecurityInfo(
                p, win32security.SE_FILE_OBJECT,
                win32security.OWNER_SECURITY_INFORMATION
            ).GetSecurityDescriptorOwner()
            if (win32security.ConvertSidToStringSid(akt)
                    == win32security.ConvertSidToStringSid(owner)):
                owner = None
        except Exception as _e:
            log_debug("ACL", f"_apply_security_descriptor: Owner-Vergleich verworfen: {_e!r}")

    if owner is not None:
        info |= win32security.OWNER_SECURITY_INFORMATION
    if group is not None:
        info |= win32security.GROUP_SECURITY_INFORMATION
    if not info:
        return

    nur_dacl = info & ~(win32security.OWNER_SECURITY_INFORMATION
                        | win32security.GROUP_SECURITY_INFORMATION)

    try:
        win32security.SetNamedSecurityInfo(
            p, win32security.SE_FILE_OBJECT, info, owner, group, dacl, None)
        log_debug("ACL", f"Sicherheitsinfo wiederhergestellt: {path}")
        return
    except Exception as e:
        if owner is None and group is None:
            log_warning("ACL", f"DACL-Wiederherstellung fehlgeschlagen ({path}): {e}")
            return
        if _restore_privileges_enabled:
            log_warning("ACL", f"Owner-Wiederherstellung fehlgeschlagen ({path}): {e}")
        else:
            log_debug("ACL", f"Owner nicht gesetzt (kein Admin-Privileg - im Nutzer-Kontext unkritisch): {path} - {e}")

    # Rueckfall: wenigstens die DACL setzen.
    if dacl is not None and nur_dacl:
        try:
            win32security.SetNamedSecurityInfo(
                p, win32security.SE_FILE_OBJECT, nur_dacl, None, None, dacl, None)
            log_debug("ACL", f"DACL wiederhergestellt (ohne Owner): {path}")
        except Exception as e2:
            log_warning("ACL", f"DACL-Wiederherstellung fehlgeschlagen ({path}): {e2}")


def safe_remove(path: str) -> bool:
    try:
        p = _lp(path)
        if os.path.exists(p):
            try:
                import win32api, win32con
                win32api.SetFileAttributes(p, win32con.FILE_ATTRIBUTE_NORMAL)
            except Exception as _e:
                detail_logger.debug(f"safe_remove: Exception verworfen: {_e!r}")
            os.remove(p)
            return True
        return False
    except Exception as e:
        print(f"  ⚠️  Löschen fehlgeschlagen {path}: {_fmt_exc(e)}")
        return False


def safe_move(src: str, dst: str) -> bool:
    try:
        shutil.move(_lp(src), _lp(dst))
        return True
    except Exception as e:
        print(f"  ⚠️  Verschieben fehlgeschlagen {src} → {dst}: {_fmt_exc(e)}")
        return False


def safe_move_with_retry(
    src: str,
    dst: str,
    retries: int = 5,
    delay: float = 1.0,
) -> bool:
    # Read-Only-Attribut auf dem Ziel zurueckspielen, bevor shutil.move
    # versucht zu ueberschreiben. Ohne das schlaegt der Move auf NTFS
    # ueber read-only-Zieldateien zwingend mit PermissionError fehl
    # (analog zur Logik in safe_remove).
    dst_lp = _lp(dst)
    try:
        if os.path.exists(dst_lp):
            import win32api, win32con
            try:
                win32api.SetFileAttributes(dst_lp, win32con.FILE_ATTRIBUTE_NORMAL)
            except Exception as _e:
                detail_logger.debug(f"safe_move_with_retry: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"safe_move_with_retry: Exception verworfen: {_e!r}")

    for attempt in range(retries):
        try:
            shutil.move(_lp(src), dst_lp)
            return True
        except PermissionError:
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                print(f"  ⚠️  Verschieben fehlgeschlagen (nach {retries} Versuchen): {src} → {dst}")
                return False
        except Exception as e:
            print(f"  ⚠️  Verschiebefehler: {_fmt_exc(e)}")
            return False
    return False


def _same_volume(a: str, b: str) -> bool:
    try:
        da = os.path.splitdrive(os.path.abspath(_strip_long_path(a)))[0]
        db = os.path.splitdrive(os.path.abspath(_strip_long_path(b)))[0]
        return bool(da) and da.lower() == db.lower()
    except Exception:
        return False


def safe_replace_with_retry(src: str, dst: str, retries: int = 5, delay: float = 1.0) -> bool:
    # Ersetzt dst durch src OHNE Truncate-Fenster: shutil.move() faellt
    # cross-volume auf copy2 zurueck und schreibt das Ziel direkt neu --
    # ein harter Prozess-Kill in diesem Fenster (Worker-Timeout!)
    # hinterliesse eine halbe Datei. Stattdessen wird src zunaechst als
    # dst + '.tmp_new' AUF DAS ZIELVOLUME kopiert und dann per
    # os.replace() atomar uebergeschoben; auf demselben Volume genuegt
    # os.replace() direkt. Das Ziel ist zu jedem Zeitpunkt entweder die
    # alte oder die neue Datei, nie ein Zwischenzustand.
    dst_lp = _lp(dst)
    try:
        if os.path.exists(dst_lp):
            import win32api, win32con
            try:
                win32api.SetFileAttributes(dst_lp, win32con.FILE_ATTRIBUTE_NORMAL)
            except Exception as _e:
                detail_logger.debug(f"safe_replace_with_retry: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"safe_replace_with_retry: Exception verworfen: {_e!r}")

    if _same_volume(src, dst):
        stage = src
    else:
        stage = _strip_long_path(dst) + ".tmp_new"
        copied = False
        last_err: Optional[BaseException] = None
        for attempt in range(retries):
            try:
                shutil.copy2(_lp(src), _lp(stage))
                copied = True
                break
            except (PermissionError, OSError) as e:
                last_err = e
                if attempt < retries - 1:
                    time.sleep(delay)
        if not copied:
            print(f"  ⚠️  Ersetzen fehlgeschlagen (Staging-Kopie): {src} → {dst}: "
                  f"{_fmt_exc(last_err) if last_err else '?'}")
            safe_remove(stage)
            return False

    for attempt in range(retries):
        try:
            os.replace(_lp(stage), dst_lp)
            if stage != src:
                safe_remove(src)
            return True
        except PermissionError as e:
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                print(f"  ⚠️  Ersetzen fehlgeschlagen (nach {retries} Versuchen): {src} → {dst}: {_fmt_exc(e)}")
        except Exception as e:
            print(f"  ⚠️  Ersetzen fehlgeschlagen: {src} → {dst}: {_fmt_exc(e)}")
            break
    if stage != src:
        safe_remove(stage)
    return False


def safe_copy(src: str, dst: str) -> bool:
    try:
        shutil.copy2(_lp(src), _lp(dst))
        return True
    except Exception as e:
        print(f"  ⚠️  Kopieren fehlgeschlagen {src} → {dst}: {_fmt_exc(e)}")
        return False


def get_file_size_mb(file_path: str) -> float:
    try:
        if safe_exists(file_path):
            return os.path.getsize(_lp(file_path)) / (1024 * 1024)
    except Exception as _e:
        detail_logger.debug(f"get_file_size_mb: Exception verworfen: {_e!r}")
    return 0.0


def is_excluded_dir(dir_name: str) -> bool:
    return dir_name.lower() in EXCLUDE_DIRS_LOWER


def is_excluded_file(file_name: str) -> bool:
    return file_name.lower() in EXCLUDE_FILES


def check_disk_space(path: str, required_mb: float = DISK_MIN_FREE_MB) -> bool:
    try:
        _, _, free = shutil.disk_usage(path)
        return (free / (1024 * 1024)) >= required_mb
    except Exception:
        return True


def is_file_locked(filepath: str) -> bool:
    if not safe_exists(filepath):
        return False
    # Read-Only-Attribut darf NICHT als "in Bearbeitung gesperrt" gewertet
    # werden: open('ab') wirft bei read-only-Dateien einen PermissionError,
    # der von OSError abgedeckt ist.
    try:
        # Echter Lock-Test: oeffnen wie OCR-Tool es spaeter macht.
        with open(_lp(filepath), 'ab'):
            pass
        return False
    except PermissionError:
        # Pruefen, ob es das Read-Only-Attribut ist (nicht ein echter Lock):
        try:
            import win32api, win32con
            attrs = win32api.GetFileAttributes(_lp(filepath))
            if attrs != -1 and (attrs & win32con.FILE_ATTRIBUTE_READONLY):
                # Read-Only - aber NICHT durch fremden Prozess gelockt.
                try:
                    with open(_lp(filepath), 'rb'):
                        pass
                    return False  # nur read-only, kein Lock
                except OSError:
                    return True   # zusaetzlich gelockt
        except Exception as _e:
            detail_logger.debug(f"is_file_locked: Exception verworfen: {_e!r}")
        return True  # PermissionError ohne erkennbares Read-Only -> Lock
    except OSError:
        return True
    except Exception:
        return False


def wait_for_file_unlock(filepath: str, max_wait: int = 10) -> bool:
    for _ in range(max_wait):
        if not is_file_locked(filepath):
            return True
        time.sleep(1)
    return False


def is_real_pdf(file_path: str) -> bool:
    # Sucht das PDF-Magic in den ersten 1024 Bytes (toleriert Header-Müll)
    try:
        with open(_lp(file_path), "rb") as fh:
            head = fh.read(PDF_MAGIC_SCAN_BYTES)
        return b"%PDF-" in head
    except Exception:
        return False


def create_temp_copy(file_path: str, temp_dir: str) -> Optional[str]:
    # Temp-Kopie nach Documents (lokal) mit AV-Scanner-Toleranz
    try:
        base = os.path.basename(file_path)
        max_name = TEMP_BASENAME_MAX - 33
        if len(base) > max_name:
            stem = os.path.splitext(base)[0]
            base = stem[:max_name - 4] + ".pdf"
        temp_name = f"{uuid.uuid4().hex}_{base}"
        temp_path = os.path.join(temp_dir, temp_name)
    except Exception as e:
        print(f"  ⚠️  Temp-Pfad nicht erstellbar: {_fmt_exc(e)}")
        return None

    # Quelldatei darf nicht gesperrt sein
    if is_file_locked(file_path):
        if not wait_for_file_unlock(file_path, max_wait=10):
            print(f"  ⚠️  Quelldatei dauerhaft gesperrt: {file_path}")
            return None

    # Kopieren mit Retry (PermissionError → AV-Scanner-Hold)
    last_err: Optional[BaseException] = None
    copied = False
    for attempt in range(COPY_RETRIES):
        try:
            shutil.copy2(_lp(file_path), _lp(temp_path))
            copied = True
            break
        except PermissionError as e:
            last_err = e
            if attempt < COPY_RETRIES - 1:
                time.sleep(COPY_RETRY_DELAY_BASE + attempt * 0.5)
        except OSError as e:
            last_err = e
            if attempt < COPY_RETRIES - 1:
                time.sleep(COPY_RETRY_DELAY_BASE + attempt * 0.5)
        except Exception as e:
            print(f"  ⚠️  Temp-Kopie nicht möglich: {_fmt_exc(e)}")
            return None

    if not copied:
        if last_err:
            print(f"  ⚠️  Temp-Kopie nach {COPY_RETRIES} Versuchen: {_fmt_exc(last_err)}")
        return None

    # Stabilisierungspause für AV-Scanner-Erstzugriff
    time.sleep(COPY_AV_STABILIZE_DELAY)

    # Lock-Check der Zielkopie (AV-Scanner hält evtl. Handle)
    if is_file_locked(temp_path):
        if not wait_for_file_unlock(temp_path, max_wait=COPY_AV_LOCK_WAIT_SECONDS):
            print(f"  ⚠️  Temp-Kopie bleibt gesperrt (AV-Scanner?): {temp_path}")
            try:
                os.remove(_lp(temp_path))
            except Exception as _e:
                detail_logger.debug(f"create_temp_copy: Exception verworfen: {_e!r}")
            return None

    return temp_path


# ==================================================================
# Backup-Leichen aufräumen
# ==================================================================

def _pdf_opens_ok(path: str) -> bool:
    # Leichtgewichtiger Validitaetstest: Magic-Number + fitz kann oeffnen.
    if not is_real_pdf(path):
        return False
    try:
        with fitz.open(_lp(path)) as doc:
            _ = doc.page_count
        return True
    except Exception:
        return False


def cleanup_orphaned_backups(
    directory: str,
    min_age_hours: int = BACKUP_CLEANUP_MIN_AGE_HOURS,
) -> int:
    count  = 0
    kept   = 0
    cutoff = time.time() - (min_age_hours * 3600)
    try:
        # Walk-Root immer praefixieren: ohne LongPathsEnabled scheitert
        # os.walk sonst still an Pfaden > 260 Zeichen.
        safe_dir = prepare_long_path(directory)
        for root, dirs, files in os.walk(
            safe_dir, topdown=True, onerror=lambda e: None, followlinks=False,
        ):
            dirs[:] = [d for d in dirs if not is_excluded_dir(d)]
            for f in files:
                low = f.lower()

                # Verwaiste Staging-Kopien aus safe_replace_with_retry und
                # set_ocr_marker.
                #
                # Beim volumeuebergreifenden Ersetzen wird die neue Datei
                # zuerst als '<name>.pdf.tmp_new' NEBEN dem Original
                # abgelegt und dann per os.replace() darueber geschoben.
                # Stirbt der Prozess in diesem Fenster - Worker-Timeout-Kill,
                # Stromausfall, harter Abbruch -, bleibt die .tmp_new liegen.
                # Aufgeraeumt wurde sie bisher NIRGENDS: weder vom
                # Temp-Verzeichnis-Aufraeumen (das betrifft nur den eigenen
                # Arbeitsordner) noch von diesem Backup-Lauf. Damit wanderten
                # solche Reste am Ende mit in die Cloud.
                #
                # Anders als beim Backup ist hier kein Schutz noetig: die
                # .tmp_new ist immer nur eine Zwischenkopie, das Original
                # liegt zu jedem Zeitpunkt unveraendert daneben.
                # '.tmp_marker' gehoert zur selben Klasse: set_ocr_marker legt
                # die neue Fassung ebenfalls NEBEN dem Original an und schiebt
                # sie per os.replace darueber. Auch diese Zwischendatei war von
                # keinem Aufraeumpfad erfasst.
                if low.endswith(".pdf.tmp_new") or low.endswith(".pdf.tmp_marker"):
                    full = os.path.join(root, f)
                    try:
                        if os.path.getmtime(_lp(full)) < cutoff and safe_remove(full):
                            count += 1
                            detail_logger.info(
                                f"Verwaiste Staging-Kopie entfernt: {_strip_long_path(full)}")
                    except Exception as _e:
                        detail_logger.debug(f"cleanup_orphaned_backups: Exception verworfen: {_e!r}")
                    continue

                # NUR eigene Artefakte (<name>.pdf.backup): ein generischer
                # *.backup-Filter wuerde fremde Sicherungsdateien loeschen.
                if low.endswith(".pdf.backup"):
                    full = os.path.join(root, f)
                    try:
                        if os.path.getmtime(_lp(full)) < cutoff:
                            # Schutz vor Datenverlust: Stammt das Backup aus
                            # einem hart abgebrochenen Lauf, kann es die
                            # einzige intakte Kopie sein. Nur loeschen, wenn
                            # die zugehoerige PDF existiert und lesbar ist.
                            sibling = _strip_long_path(full)[: -len(".backup")]
                            if not (safe_exists(sibling) and _pdf_opens_ok(sibling)):
                                kept += 1
                                msg = (f"Backup NICHT geloescht – Original fehlt oder ist "
                                       f"defekt (Backup ist evtl. die einzige intakte Kopie): "
                                       f"{_strip_long_path(full)}")
                                print(f"  ⚠️  {msg}")
                                log_warning("cleanup_backups", msg)
                                continue
                            if safe_remove(full):
                                count += 1
                    except Exception as _e:
                        detail_logger.debug(f"cleanup_orphaned_backups: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"cleanup_orphaned_backups: Exception verworfen: {_e!r}")
    if kept:
        print(f"  ⚠️  {kept} Backup(s) wegen fehlendem/defektem Original behalten – bitte manuell pruefen.")
    return count


# ==================================================================
# Memory- und Disk-Check vor OCR
# ==================================================================

def estimate_ocr_memory_mb(pages: int, jobs: int = 1) -> float:
    # Realistische Schätzung für ocrmypdf (seitenweise Verarbeitung)
    base = max(MEMORY_BASE_MB_FLOOR,
               min(pages * MEMORY_MB_PER_PAGE_MIN, MEMORY_BASE_MB_CAP))
    return base + max(1, jobs) * MEMORY_MB_PER_JOB


def check_memory_before_ocr(pages: int, jobs: int = 1) -> Tuple[bool, str]:
    # Kleine PDFs (< 50 Seiten) immer versuchen – RAM-Bedarf ist bei wenigen Seiten harmlos
    if pages < 50:
        return True, ""
    try:
        available_mb = psutil.virtual_memory().available / (1024 * 1024)
        estimated_mb = estimate_ocr_memory_mb(pages, jobs)
        # Skip nur, wenn der Schätzbedarf deutlich (> 1.2×) über dem verfügbaren RAM liegt
        if estimated_mb > available_mb * 1.2:
            return False, (
                f"Schätzbedarf {estimated_mb:.0f} MB > 1.2 × verfügbar {available_mb:.0f} MB"
            )
        return True, ""
    except Exception:
        return True, ""


def required_disk_mb_for_pdf(size_mb: float) -> float:
    return max(DISK_MIN_FREE_MB, size_mb * DISK_FACTOR_PER_MB)


# ==================================================================
# PDF-Prüfungs- und Markierungsfunktionen
# ==================================================================

def _read_xmp_pdfa_code(pdf_path: str) -> Optional[str]:
    try:
        import pikepdf
        with pikepdf.Pdf.open(_lp(pdf_path)) as pdf:
            try:
                with pdf.open_metadata() as meta:
                    part = str(meta.get("pdfaid:part", "") or "").strip()
                    conf = str(meta.get("pdfaid:conformance", "") or "").strip().upper()
            except Exception:
                return None
            if part == "2" and conf == "U":
                return "PDFA_2U"
            if part == "2" and conf == "B":
                return "PDFA_2B"
            if part == "1" and conf in ("A", "B"):
                return "PDFA_1B"
            if part == "3" and conf in ("U", "A"):
                return "PDFA_3U"
            if part == "3" and conf == "B":
                return "PDFA_3B"
    except Exception:
        return None
    return None


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
    # Separation/DeviceN sind Sonderfarben (Pantone o. ae.) und fuer den
    # Druck genauso schuetzenswert wie CMYK.
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
    """Anzahl Farbkanaele eines Bildes, ohne den Bildstrom zu dekodieren.

    Der Farbraumname aus get_images() reicht nicht: CMYK-Bilder aus
    Layoutprogrammen sind fast immer /ICCBased und melden sich damit als
    "ICCBased", nicht als "DeviceCMYK". Die Kanalzahl steht im /N-Eintrag
    des ICC-Streams, auf den ueber ein bis zwei Referenzen verwiesen wird.
    """
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


def _analyze_image_for_print(doc, page, img_tuple, w: int, h: int, px: int,
                             info: Dict[str, Any]) -> None:
    """Ermittelt effektive Aufloesung und Farbraum eines eingebetteten Bildes.

    Die effektive DPI ergibt sich aus der Pixelbreite geteilt durch die
    Breite, mit der das Bild auf der Seite platziert ist (in Zoll). Genau
    diese Groesse interessiert eine Druckerei – die reine Pixelzahl sagt
    nichts darueber aus, wie fein das Bild gedruckt erscheint.
    """
    if px < PRINT_SAFE_MIN_PIXELS:
        return
    try:
        xref = int(img_tuple[0])
    except (IndexError, TypeError, ValueError):
        return

    if image_color_components(doc, xref) >= 4:
        info["has_cmyk_images"] = True

    # Platzierung auf der Seite suchen; ohne Treffer keine DPI-Aussage
    try:
        rects = page.get_image_rects(xref)
    except Exception:
        rects = []
    for r in rects or []:
        try:
            width_in  = float(r.width) / 72.0
            height_in = float(r.height) / 72.0
        except Exception:
            continue
        if width_in <= 0 or height_in <= 0:
            continue
        dpi = max(w / width_in, h / height_in)
        if dpi > info["max_image_dpi"]:
            info["max_image_dpi"] = dpi
        if dpi >= PRINT_SAFE_MIN_DPI:
            info["has_highres_images"] = True


def needs_print_safe_profile(pdf_info: Dict[str, Any]) -> Tuple[bool, str]:
    """Entscheidet, ob eine Datei bildschonend verarbeitet werden muss."""
    reasons = []
    if pdf_info.get("has_highres_images"):
        reasons.append(f"Bild mit {pdf_info.get('max_image_dpi', 0):.0f} dpi "
                       f"(Schwelle {PRINT_SAFE_MIN_DPI} dpi)")
    if pdf_info.get("has_cmyk_images"):
        reasons.append("CMYK-/Separations-Bild")
    return bool(reasons), ", ".join(reasons)


def get_pdf_info(file_path: str) -> Dict[str, Any]:
    info: Dict[str, Any] = {
        "path":           file_path,
        "size_mb":        0.0,
        "pages":          0,
        "has_text":       False,
        "has_ocr_marker": False,
        "existing_marker": None,
        "is_valid":       False,
        "is_encrypted":   False,
        "needs_repair":   False,
        "has_forms":      False,
        "is_oversized":   False,
        "max_image_pixels": 0,
        "max_image_dpi":  0.0,
        "has_highres_images": False,
        "has_cmyk_images":    False,
        "subject":        "",
        "error":          None,
    }

    if not safe_exists(file_path):
        info["error"] = "Datei existiert nicht"
        return info

    info["size_mb"] = get_file_size_mb(file_path)

    try:
        p = _lp(file_path)
        try:
            doc = fitz.open(p)
        except fitz.FileDataError:
            info["needs_repair"] = True
            info["is_valid"]     = True
            return info
        except Exception as e:
            info["error"] = _fmt_exc(e)
            return info

        with doc:
            if doc.is_encrypted:
                info["is_encrypted"] = True
                info["is_valid"]     = True
                return info

            info["pages"]    = len(doc)
            info["is_valid"] = True

            try:
                if getattr(doc, "is_repaired", False):
                    info["needs_repair"] = True
            except Exception as _e:
                detail_logger.debug(f"get_pdf_info: Exception verworfen: {_e!r}")

            metadata = doc.metadata or {}
            subject = str(metadata.get(MARKER_KEY, ""))
            info["subject"] = subject

            if MARKER_VALUE_PDFA_2U in subject:
                info["existing_marker"] = "PDFA_2U"
            elif MARKER_VALUE_PDFA_2B in subject:
                info["existing_marker"] = "PDFA_2B"
            elif MARKER_VALUE_PDF in subject:
                info["existing_marker"] = "PDF"
            elif LEGACY_MARKER in subject:
                info["existing_marker"] = "LEGACY"

            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                # get_text ist der teuerste Teil dieser Schleife und liefe bei
                # reinen Scan-PDFs (kein Text) ueber ALLE Seiten -- has_text
                # ist aber rein informativ, daher auf die ersten N Seiten
                # begrenzt. Forms/Oversize/Bildgroessen bleiben Volldurchlauf
                # (billig, keine Inhalts-Dekodierung).
                if (not info["has_text"]
                        and page_num < PDF_INFO_TEXT_SCAN_PAGES
                        and page.get_text("text").strip()):
                    info["has_text"] = True
                if not info["has_forms"] and list(page.widgets()):
                    info["has_forms"] = True
                if not info["is_oversized"]:
                    rect = page.rect
                    if rect.width > MAX_PAGE_DIMENSION or rect.height > MAX_PAGE_DIMENSION:
                        info["is_oversized"] = True
                # Groesstes eingebettetes Bild verfolgen. Breite/Hoehe stehen
                # direkt im get_images()-Tupel (xref, smask, width, height, ...);
                # extract_image() wuerde dafuer unnoetig den kompletten
                # komprimierten Bild-Stream in den Speicher laden.
                try:
                    for img in page.get_images(full=False):
                        try:
                            w = int(img[2] or 0)
                            h = int(img[3] or 0)
                        except (IndexError, TypeError, ValueError):
                            continue
                        px = w * h
                        if px > info["max_image_pixels"]:
                            info["max_image_pixels"] = px
                        _analyze_image_for_print(doc, page, img, w, h, px, info)
                except Exception as _e:
                    detail_logger.debug(f"get_pdf_info: Exception verworfen: {_e!r}")

        # XMP-Rueckfall NUR bei tatsaechlich vorhandener Textebene.
        # Die Marker aus dem Subject setzt ausschliesslich dieses Skript, sie
        # belegen also wirklich eine OCR-Verarbeitung. Die XMP-Kennung sagt
        # dagegen nur, dass die Datei als PDF/A vorliegt - das kann jeder
        # Scanner oder jedes Archivwerkzeug erzeugt haben, ganz ohne OCR.
        # Ein reiner Bild-Scan im Format PDF/A-2u galt damit als 'bereits
        # verarbeitet' und wurde uebersprungen, obwohl er keine Textebene hat:
        # genau die Dateien, deretwegen das Skript laeuft.
        if info["existing_marker"] is None and info["has_text"]:
            xmp_code = _read_xmp_pdfa_code(file_path)
            if xmp_code in ("PDFA_2U", "PDFA_2B"):
                info["existing_marker"] = xmp_code

        info["has_ocr_marker"] = info["existing_marker"] is not None

    except Exception as e:
        info["error"] = _fmt_exc(e)

    return info


def _sichere_datei_metadaten(file_path: str):
    """Zeitstempel und NTFS-Sicherheitsinfo einer Datei festhalten.

    Wird gebraucht, wo eine Datei per os.replace ERSETZT wird: das Ergebnis
    ist dann eine neue Datei mit aktuellem Zeitstempel, dem ausfuehrenden
    Konto als Eigentuemer und nur den vom Ordner geerbten ACEs. Auf einer
    Ablage mit Owner-Mapping kostet das den Fachnutzer den Zugriff auf seine
    eigene Datei - genau der Fall, den der Kommentarblock zur ACL-Uebernahme
    weiter oben beschreibt.
    """
    times = None
    sd    = None
    try:
        times = _read_file_times(file_path)
    except Exception as _e:
        detail_logger.debug(f"_sichere_datei_metadaten (Zeiten): {_e!r}")
    try:
        sd = _get_security_descriptor(file_path)
    except Exception as _e:
        detail_logger.debug(f"_sichere_datei_metadaten (ACL): {_e!r}")
    return (times, sd)


def _stelle_datei_metadaten_her(file_path: str, gesichert) -> None:
    """Gegenstueck zu _sichere_datei_metadaten. Fehler sind nicht fatal."""
    times, sd = gesichert if gesichert else (None, None)
    if sd is not None:
        try:
            _apply_security_descriptor(file_path, sd)
        except Exception as _e:
            detail_logger.debug(f"_stelle_datei_metadaten_her (ACL): {_e!r}")
    # Zeitstempel zuletzt: das Setzen der ACL wuerde ihn sonst wieder
    # ueberschreiben koennen.
    if times is not None:
        try:
            _write_file_times(file_path, times)
        except Exception as _e:
            detail_logger.debug(f"_stelle_datei_metadaten_her (Zeiten): {_e!r}")


def _add_ocr_marker_legacy(file_path: str, marker_value: str = MARKER_VALUE_PDF) -> bool:
    if not safe_exists(file_path):
        return False
    p = _lp(file_path)
    tmp = p + ".tmp_marker"
    # Vor jeder Aenderung sichern - der Nicht-Inkrementell-Zweig unten ersetzt
    # die Datei und wuerde Zeitstempel, Eigentuemer und explizite ACEs sonst
    # verlieren.
    gesichert = _sichere_datei_metadaten(file_path)
    try:
        doc = fitz.open(p)
        try:
            metadata = doc.metadata or {}
            metadata[MARKER_KEY] = marker_value
            metadata["producer"] = f"OCR-Automation {datetime.now().strftime('%Y-%m-%d')}"
            metadata["creator"]  = "PDF OCR Automation"
            doc.set_metadata(metadata)
            try:
                doc.save(p, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
                _stelle_datei_metadaten_her(file_path, gesichert)
                return True
            except Exception:
                doc.save(tmp, garbage=2, deflate=True)
        finally:
            doc.close()
        if safe_replace_with_retry(tmp, p):
            _stelle_datei_metadaten_her(file_path, gesichert)
            return True
        safe_remove(tmp)
        return False
    except Exception as e:
        log_warning("_add_ocr_marker_legacy",
                    f"Marker-Setzen fehlgeschlagen ({file_path}): {_fmt_exc(e)}")
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception as _e:
            detail_logger.debug(f"_add_ocr_marker_legacy: Exception verworfen: {_e!r}")
        return False


def set_ocr_marker(file_path: str, marker_value: str) -> bool:
    # Zeitstempel und Sicherheitsinfo VOR der Ersetzung festhalten. Die
    # Funktion schreibt eine komplette Neufassung nach '<name>.pdf.tmp_marker'
    # und schiebt sie ueber das Original; danach ist es eine neue Datei.
    # Beim Lauf mit --verify-markers endet der aufrufende Zweig direkt danach
    # mit SKIPPED, die Wiederherstellung am Ende der Verarbeitung wird also
    # nie erreicht - und die spaeter gelesenen orig_times/orig_sd enthielten
    # ohnehin schon die zerstoerten Werte. Deshalb hier, in der Funktion, die
    # die Ersetzung tatsaechlich vornimmt.
    gesichert = _sichere_datei_metadaten(file_path)
    try:
        import pikepdf
        p   = _lp(file_path)
        tmp = p + ".tmp_marker"
        with pikepdf.Pdf.open(p, allow_overwriting_input=False) as pdf:
            try:
                pdf.docinfo["/Subject"] = pikepdf.String(marker_value)
                pdf.docinfo["/Producer"] = pikepdf.String(
                    f"OCR-Automation {datetime.now().strftime('%Y-%m-%d')}"
                )
                pdf.docinfo["/Creator"] = pikepdf.String("PDF OCR Automation")
            except Exception as _e:
                detail_logger.debug(f"set_ocr_marker: Exception verworfen: {_e!r}")
            try:
                with pdf.open_metadata() as meta:
                    meta["dc:description"] = marker_value
            except Exception as _e:
                detail_logger.debug(f"set_ocr_marker: Exception verworfen: {_e!r}")
            pdf.save(tmp)
        if safe_replace_with_retry(tmp, p):
            _stelle_datei_metadaten_her(file_path, gesichert)
            return True
        safe_remove(tmp)
    except ImportError:
        pass
    except Exception as e:
        log_warning("set_ocr_marker",
                    f"pikepdf-Marker setzen fehlgeschlagen ({file_path}): {_fmt_exc(e)}")
        try:
            tmp_path = _lp(file_path) + ".tmp_marker"
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception as _e:
            detail_logger.debug(f"set_ocr_marker: Exception verworfen: {_e!r}")
    return _add_ocr_marker_legacy(file_path, marker_value)


def repair_pdf(file_path: str, temp_dir: str) -> bool:
    tmp_path = os.path.join(temp_dir, f"{uuid.uuid4().hex}_repaired.pdf")
    # Zeitstempel und Sicherheitsinfo VOR dem Ersetzen festhalten. Die
    # reparierte Datei wird per safe_replace_with_retry ueber das Original
    # geschoben und ist danach eine neue Datei. Die Sicherung im Aufrufer
    # greift zu spaet: sie liest erst nach dem Reparaturblock und wuerde
    # damit den Reparaturzeitpunkt und die bereits verlorenen Rechte
    # wiederherstellen. Gleiche Ursache und gleiche Loesung wie bei
    # set_ocr_marker.
    gesichert = _sichere_datei_metadaten(file_path)
    try:
        p = _lp(file_path)
        with fitz.open(p) as src, fitz.open() as dst:
            dst.insert_pdf(src)
            dst.save(tmp_path, garbage=4, deflate=True)
        if safe_replace_with_retry(tmp_path, file_path):
            _stelle_datei_metadaten_her(file_path, gesichert)
            return True
        return False
    except Exception as e:
        log_warning("repair_pdf", f"Reparatur fehlgeschlagen ({file_path}): {_fmt_exc(e)}")
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception as _e:
            detail_logger.debug(f"repair_pdf: Exception verworfen: {_e!r}")
        return False


def try_remove_empty_password(file_path: str, temp_dir: str) -> bool:
    # Scanner/Multifunktionsgeraete erzeugen haeufig PDFs mit reinem
    # Owner-Passwort: is_encrypted ist wahr, aber das Benutzerpasswort
    # ist leer. Solche Dateien lassen sich mit pikepdf oeffnen und
    # unverschluesselt neu speichern; Zeitstempel und NTFS-Sicherheits-
    # info des Originals bleiben dabei erhalten. PDFs mit echtem
    # Benutzerpasswort scheitern hier und werden wie bisher uebersprungen.
    try:
        import pikepdf
    except ImportError:
        return False
    tmp_path   = os.path.join(temp_dir, f"{uuid.uuid4().hex}_decrypted.pdf")
    orig_times = _read_file_times(file_path)
    orig_sd    = _get_security_descriptor(file_path)
    try:
        with pikepdf.Pdf.open(_lp(file_path), password="") as pdf:
            pdf.save(tmp_path)
        if not safe_replace_with_retry(tmp_path, file_path):
            safe_remove(tmp_path)
            return False
        _apply_security_descriptor(file_path, orig_sd)
        if orig_times:
            _write_file_times(file_path, orig_times)
        return True
    except Exception as e:
        log_debug("decrypt",
                  f"Leeres-Passwort-Entschluesselung nicht moeglich ({file_path}): {_fmt_exc(e)}")
        safe_remove(tmp_path)
        return False


# ==================================================================
# veraPDF-Validierung
# ==================================================================

def _verapdf_timeout_for(pages: int) -> int:
    t = int(VERAPDF_BASE_TIMEOUT + pages * VERAPDF_PER_PAGE_SECONDS)
    return max(VERAPDF_BASE_TIMEOUT, min(VERAPDF_MAX_TIMEOUT, t))


def verify_pdfa_with_verapdf(
    pdf_path: str,
    verapdf_path: str,
    desired_output: str,
    pages: int = 0,
) -> Optional[str]:
    if not verapdf_path or not os.path.exists(verapdf_path):
        return None
    p = _lp(pdf_path)
    timeout = _verapdf_timeout_for(pages)

    flavour_order: List[Tuple[str, str]] = []
    if desired_output == "pdfa-2u":
        flavour_order = [("2u", "PDFA_2U"), ("2b", "PDFA_2B")]
    elif desired_output == "pdfa-2b":
        flavour_order = [("2b", "PDFA_2B")]
    elif desired_output == "pdfa-2":
        flavour_order = [("2u", "PDFA_2U"), ("2b", "PDFA_2B")]
    elif desired_output == "pdfa-1":
        flavour_order = [("1b", "PDFA_1B")]
    elif desired_output == "pdfa-3":
        flavour_order = [("3u", "PDFA_3U"), ("3b", "PDFA_3B")]
    else:
        return None

    # verapdf.bat laeuft durch cmd.exe, und cmd expandiert %VAR% auch
    # INNERHALB von Anfuehrungszeichen. Originalpfade mit cmd-Sonder-
    # zeichen daher ueber eine neutral benannte Temp-Kopie validieren.
    cleanup_tmp: Optional[str] = None
    if (verapdf_path.lower().endswith((".bat", ".cmd"))
            and re.search(r"[%^&!]", pdf_path)):
        tmp_v = os.path.join(tempfile.gettempdir(), f"verapdf_{uuid.uuid4().hex}.pdf")
        if safe_copy(pdf_path, tmp_v):
            p = _lp(tmp_v)
            cleanup_tmp = tmp_v

    try:
        for flavour_arg, marker_code in flavour_order:
            try:
                r = _run(
                    [verapdf_path, "--format", "text", "--flavour", flavour_arg, p],
                    timeout=timeout,
                )
                if r.returncode == 0:
                    return marker_code
            except subprocess.TimeoutExpired:
                log_warning("verapdf", f"Timeout bei Validierung: {pdf_path}")
            except Exception as e:
                log_warning("verapdf", f"Validierung fehlgeschlagen: {_fmt_exc(e)}")
        return None
    finally:
        if cleanup_tmp:
            safe_remove(cleanup_tmp)


# ==================================================================
# Marker-Re-Verifikation (für --verify-markers)
# ==================================================================

def verify_existing_marker(
    file_path: str,
    existing_marker_code: str,
    config: Dict,
    pages: int = 0,
) -> Tuple[bool, Optional[str]]:
    if not config.get("verify_markers"):
        return True, existing_marker_code
    if not config.get("has_verapdf"):
        return True, existing_marker_code
    if existing_marker_code not in ("PDFA_2U", "PDFA_2B"):
        return True, existing_marker_code

    verapdf_path = config.get("verapdf_path", "")
    desired = "pdfa-2u" if existing_marker_code == "PDFA_2U" else "pdfa-2b"
    actual  = verify_pdfa_with_verapdf(file_path, verapdf_path, desired, pages=pages)

    if actual is None:
        return False, None

    if actual == existing_marker_code:
        return True, existing_marker_code

    return False, actual


# ==================================================================
# CPU- und RAM-Info, Worker-Empfehlung
# ==================================================================

def get_cpu_info() -> Tuple[int, str]:
    cpu_count = os.cpu_count() or 1
    cpu_name = "Unbekannt"
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"HARDWARE\DESCRIPTION\System\CentralProcessor\0",
        ) as key:
            cpu_name = winreg.QueryValueEx(key, "ProcessorNameString")[0].strip()
    except Exception:
        try:
            cpu_name = platform.processor() or "Unbekannt"
        except Exception as _e:
            detail_logger.debug(f"get_cpu_info: Exception verworfen: {_e!r}")
    return cpu_count, cpu_name


def get_total_ram_gb() -> float:
    try:
        return psutil.virtual_memory().total / (1024 ** 3)
    except Exception:
        return 16.0


def print_cpu_recommendation(cpu_count: int, cpu_name: str) -> None:
    ram_gb = get_total_ram_gb()
    print(f"  CPU:   {cpu_name[:70]}")
    print(f"  Kerne: {cpu_count} logische Kerne")
    print(f"  RAM:   {ram_gb:.1f} GB gesamt")
    if ram_gb < 8:
        print("  ⚠️  Wenig RAM (<8 GB) – konservative Worker-Empfehlung (1)")
    elif ram_gb < 12:
        print("  ℹ️  Knapper RAM (<12 GB) – moderate Worker-Empfehlung (max 2)")
    elif cpu_count >= 8:
        print("  ✓ Mehrkern-CPU – Parallelisierung sehr empfohlen (3–4 Worker)")
    elif cpu_count >= 4:
        print("  ✓ Quad-Core – Parallelisierung empfohlen (2–3 Worker)")
    elif cpu_count >= 2:
        print("  ℹ️  Dual-Core – Parallelisierung möglich (1–2 Worker)")
    else:
        print("  ⚠️  Single-Core – 1 Worker empfohlen")


def recommended_worker_count(cpu_count: int) -> int:
    # Berücksichtigt Kerne und RAM für 8/16 GB-Laptops
    ram_gb = get_total_ram_gb()
    if ram_gb < 8:
        return 1
    if ram_gb < 12:
        return min(2, max(1, cpu_count // 2))
    if ram_gb < 20:
        return min(3, max(1, cpu_count // 2))
    return min(4, max(1, cpu_count // 2))


def recommended_jobs_per_worker(cpu_count: int, num_workers: int) -> int:
    # jobs nicht zu hoch für RAM-knappe Systeme
    ram_gb = get_total_ram_gb()
    raw = max(1, (cpu_count or 1) // max(1, num_workers))
    if ram_gb < 12:
        return 1
    if ram_gb < 20:
        return min(2, raw)
    return min(3, raw)


def ask_worker_count(cpu_count: int) -> int:
    default = recommended_worker_count(cpu_count)
    max_rec = max(1, min(6, cpu_count - 1))
    print("\n  Wie viele PDFs sollen gleichzeitig verarbeitet werden?")
    print(f"  1 = sequenziell | {default} = empfohlen | max. sinnvoll: {max_rec}")
    while True:
        try:
            answer = input(f"  Worker-Anzahl [Standard: {default}]: ").strip()
        except EOFError:
            return default
        if answer == "":
            return default
        try:
            n = int(answer)
            max_workers = min(61, cpu_count)
            if 1 <= n <= max_workers:
                return n
            print(f"  Bitte eine Zahl zwischen 1 und {max_workers} eingeben.")
        except ValueError:
            print("  Bitte eine Zahl eingeben.")


# ==================================================================
# Zeitstempel-Hilfsfunktionen
# ==================================================================

def _read_file_times(file_path: str) -> Optional[Any]:
    p = _lp(file_path)
    try:
        import win32file
        FILE_READ_ATTRIBUTES = 0x0080
        h = win32file.CreateFile(
            p, FILE_READ_ATTRIBUTES,
            win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
            None, win32file.OPEN_EXISTING, 0, None
        )
        try:
            times = win32file.GetFileTime(h)
            return ("win32", times)
        finally:
            h.Close()
    except Exception as _e:
        detail_logger.debug(f"_read_file_times: Exception verworfen: {_e!r}")

    try:
        st = os.stat(p)
        return ("stat", (st.st_atime, st.st_mtime))
    except Exception:
        return None


def _write_file_times(file_path: str, times_data: Any) -> bool:
    p = _lp(file_path)
    kind, times = times_data

    if kind == "win32":
        import win32file
        FILE_WRITE_ATTRIBUTES = 0x0100
        for retry in range(5):
            try:
                h = win32file.CreateFile(
                    p, FILE_WRITE_ATTRIBUTES,
                    win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | 4,
                    None, win32file.OPEN_EXISTING, 0, None
                )
                try:
                    win32file.SetFileTime(h, times[0], times[1], times[2])
                finally:
                    h.Close()
                return True
            except Exception as e:
                if retry == 4:
                    log_warning("write_times", f"Zeitstempel nicht wiederherstellbar: {_fmt_exc(e)}")
                else:
                    time.sleep(0.5)
        return False

    elif kind == "stat":
        atime, mtime = times
        try:
            os.utime(p, (atime, mtime))
            return True
        except Exception as e:
            log_warning("write_times", f"os.utime fehlgeschlagen: {_fmt_exc(e)}")
            return False

    return False


# ==================================================================
# Systemprüfung: Externe Abhängigkeiten
# ==================================================================

def check_external_dependencies() -> Dict:
    results: Dict = {}

    print("\n" + "=" * 80)
    print("  SYSTEMPRÜFUNG: EXTERNE ABHÄNGIGKEITEN")
    print("=" * 80)

    exe_dir = _get_exe_dir()
    meipass = _get_meipass_dir()

    # 1. Tesseract
    print("\n[1] Prüfe Tesseract OCR ...")
    tesseract_found   = False
    tesseract_version = "Nicht gefunden"
    tesseract_path    = ""
    tesseract_langs: List[str] = []

    try:
        candidate_paths = [
            os.path.join(exe_dir, "tesseract.exe"),
            os.path.join(exe_dir, "Tesseract-OCR", "tesseract.exe"),
            r"C:\OCR\Tesseract-OCR\tesseract.exe",
            r"C:\OCR\tesseract.exe",
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            r"C:\Tesseract-OCR\tesseract.exe",
        ]
        if meipass:
            candidate_paths.insert(0, os.path.join(meipass, "tesseract.exe"))

        for p in candidate_paths:
            if os.path.exists(p):
                tesseract_path = p
                break

        if not tesseract_path:
            tesseract_path = shutil.which("tesseract") or ""

        if tesseract_path and os.path.exists(tesseract_path):
            tess_dir = os.path.dirname(tesseract_path)
            tessdata_dir = os.path.join(tess_dir, "tessdata")
            if os.path.exists(tessdata_dir):
                os.environ["TESSDATA_PREFIX"] = tessdata_dir
            _register_dll_dirs([tess_dir])

            r = _run([tesseract_path, "--version"], timeout=5)
            if r.returncode == 0:
                tesseract_found = True
                for line in r.stdout.split("\n"):
                    if "tesseract" in line.lower():
                        tesseract_version = line.strip()
                        break

                lr = _run([tesseract_path, "--list-langs"], timeout=5)
                if lr.returncode == 0:
                    tesseract_langs = [
                        lang.strip() for lang in lr.stdout.split("\n")
                        if lang.strip() and not lang.startswith("List of available")
                    ]

                print(f"  ✓ {tesseract_path}")
                print(f"  Version: {tesseract_version}")
                print(f"  Sprachen: {len(tesseract_langs)} verfügbar")

                missing_langs = [
                    lang for lang in ([DEFAULT_LANGUAGE] + FALLBACK_LANGUAGES)
                    if lang not in tesseract_langs
                ]
                if missing_langs:
                    print(f"  ⚠️  Fehlende Sprachpakete: {', '.join(missing_langs)}")
                else:
                    print("  ✓ Alle benötigten Sprachpakete vorhanden")
            else:
                print("  ✗ Tesseract gefunden, aber Version nicht ermittelbar")
        else:
            print("  ✗ Tesseract nicht gefunden!")
    except Exception as e:
        print(f"  ✗ Tesseract-Prüfung fehlgeschlagen: {_fmt_exc(e)}")

    results["tesseract"]       = (tesseract_found, tesseract_version, tesseract_path)
    results["tesseract_langs"] = tesseract_langs

    # 2. Leptonica
    print("\n[2] Prüfe Leptonica DLL ...")
    leptonica_found   = False
    leptonica_version = "Nicht gefunden"
    leptonica_path    = ""

    try:
        search_paths: List[str] = []
        if meipass:
            search_paths.append(meipass)
        search_paths.append(exe_dir)
        if tesseract_path:
            tdir = os.path.dirname(tesseract_path)
            search_paths += [tdir, os.path.join(tdir, "..")]
        search_paths += EXTRA_DLL_DIRS
        search_paths += os.environ.get("PATH", "").split(os.pathsep)
        search_paths += [r"C:\Windows\System32", r"C:\Windows\SysWOW64"]

        seen: Set[str] = set()
        unique_paths: List[str] = []
        for sp in search_paths:
            sp_norm = os.path.normcase(os.path.abspath(sp)) if sp else ""
            if sp_norm and sp_norm not in seen and os.path.isdir(sp):
                seen.add(sp_norm)
                unique_paths.append(sp)

        patterns = [
            "leptonica-1.*.dll",
            "liblept-*.dll",
            "liblept*.dll",
            "leptonica.dll",
            "libleptonica-*.dll",
            "libleptonica.dll",
        ]

        found_dll = None
        for sdir in unique_paths:
            for pattern in patterns:
                try:
                    matches = sorted(Path(sdir).glob(pattern))
                except Exception:
                    matches = []
                if matches:
                    found_dll = str(matches[0])
                    break
            if found_dll:
                break

        if found_dll:
            leptonica_path  = found_dll
            leptonica_found = True
            m = re.search(r"(\d+\.\d+\.\d+)", os.path.basename(found_dll))
            leptonica_version = (
                f"Leptonica {m.group(1)}" if m
                else f"Leptonica ({os.path.basename(found_dll)})"
            )
            _register_dll_dirs([os.path.dirname(found_dll)])
            print(f"  ✓ {leptonica_path}")
            print(f"  Version: {leptonica_version}")
        else:
            print("  ℹ️  Keine separate Leptonica-DLL gefunden.")
            print("     (Der UB-Mannheim-Tesseract linkt Leptonica statisch – dann ist")
            print("      keine DLL noetig. Massgeblich ist der ocrmypdf-Systemtest unten.)")
            print(f"     → Suchpfade: {', '.join(unique_paths[:5])} ...")
    except Exception as e:
        print(f"  ✗ Leptonica-Prüfung fehlgeschlagen: {_fmt_exc(e)}")

    results["leptonica"] = (leptonica_found, leptonica_version, leptonica_path)

    # 3. Ghostscript
    print("\n[3] Prüfe Ghostscript ...")
    gs_found   = False
    gs_version = "Nicht gefunden"
    gs_path    = ""

    try:
        candidates = []
        for name in ("gswin64c.exe", "gswin32c.exe", "gs.exe"):
            candidates.append(os.path.join(exe_dir, name))
            candidates.append(os.path.join(r"C:\OCR", name))
            candidates.append(os.path.join(r"C:\OCR\bin", name))
        for p in candidates:
            if os.path.exists(p):
                gs_path = p
                break
        if not gs_path:
            for exe in ["gswin64c.exe", "gswin32c.exe", "gs.exe"]:
                found = shutil.which(exe)
                if found:
                    gs_path = found
                    break

        if gs_path and os.path.exists(gs_path):
            r = _run([gs_path, "--version"], timeout=5)
            if r.returncode == 0:
                gs_found   = True
                gs_version = r.stdout.strip()
                print(f"  ✓ {gs_path}")
                print(f"  Version: {gs_version}")
            else:
                print("  ✗ Ghostscript gefunden, aber Version nicht ermittelbar")
        else:
            print("  ✗ Ghostscript nicht gefunden!")
    except Exception as e:
        print(f"  ✗ Ghostscript-Prüfung fehlgeschlagen: {_fmt_exc(e)}")

    results["ghostscript"] = (gs_found, gs_version, gs_path)

    # 4. JBIG2
    print("\n[4] Prüfe JBIG2 Encoder ...")
    jbig2_found   = False
    jbig2_version = "Nicht gefunden"
    jbig2_path    = ""

    try:
        import importlib.util
        if importlib.util.find_spec("jbig2enc") is not None:
            jbig2_found   = True
            jbig2_version = "Python-Modul (jbig2enc)"
            print("  ✓ Python-Modul jbig2enc verfügbar")
        else:
            candidates = [
                shutil.which("jbig2"),
                os.path.join(exe_dir, "jbig2.exe"),
                r"C:\OCR\jbig2.exe",
                r"C:\OCR\bin\jbig2.exe",
                r"C:\Program Files\jbig2\jbig2.exe",
                r"C:\jbig2\jbig2.exe",
            ]
            for p in candidates:
                if p and os.path.exists(p):
                    jbig2_path = p
                    break
            if jbig2_path:
                r = _run([jbig2_path, "--version"], timeout=5)
                if r.returncode == 0:
                    jbig2_found   = True
                    jbig2_version = r.stdout.strip()
                    print(f"  ✓ {jbig2_path}  Version: {jbig2_version}")
                else:
                    print("  ⚠️  JBIG2 gefunden, aber Version nicht ermittelbar")
            else:
                print("  ⚠️  JBIG2 nicht gefunden (optional)")
    except Exception as e:
        print(f"  ✗ JBIG2-Prüfung fehlgeschlagen: {_fmt_exc(e)}")

    results["jbig2"] = (jbig2_found, jbig2_version, jbig2_path)

    # 5. pngquant
    print("\n[5] Prüfe pngquant ...")
    pngquant_found   = False
    pngquant_version = "Nicht gefunden"
    pngquant_path    = ""

    try:
        candidates = [
            shutil.which("pngquant"),
            os.path.join(exe_dir, "pngquant.exe"),
            r"C:\OCR\pngquant.exe",
            r"C:\OCR\bin\pngquant.exe",
            r"C:\Program Files\pngquant\pngquant.exe",
            r"C:\pngquant\pngquant.exe",
        ]
        for p in candidates:
            if p and os.path.exists(p):
                pngquant_path = p
                break
        if pngquant_path:
            r = _run([pngquant_path, "--version"], timeout=5)
            if r.returncode == 0:
                pngquant_found   = True
                pngquant_version = r.stdout.strip()
                print(f"  ✓ {pngquant_path}  Version: {pngquant_version}")
            else:
                print("  ⚠️  pngquant gefunden, aber Version nicht ermittelbar")
        else:
            print("  ⚠️  pngquant nicht gefunden (optional)")
    except Exception as e:
        print(f"  ✗ pngquant-Prüfung fehlgeschlagen: {_fmt_exc(e)}")

    results["pngquant"] = (pngquant_found, pngquant_version, pngquant_path)

    # 6. unpaper
    print("\n[6] Prüfe unpaper ...")
    unpaper_found   = False
    unpaper_version = "Nicht gefunden"
    unpaper_path    = ""

    try:
        candidates = [
            shutil.which("unpaper"),
            os.path.join(exe_dir, "unpaper.exe"),
            r"C:\OCR\unpaper.exe",
            r"C:\OCR\bin\unpaper.exe",
            r"C:\Program Files\unpaper\unpaper.exe",
            r"C:\unpaper\unpaper.exe",
        ]
        for p in candidates:
            if p and os.path.exists(p):
                unpaper_path = p
                break
        if unpaper_path:
            r = _run([unpaper_path, "--version"], timeout=5)
            if r.returncode == 0:
                unpaper_found   = True
                unpaper_version = r.stdout.strip()
                print(f"  ✓ {unpaper_path}  Version: {unpaper_version}")
            else:
                print("  ⚠️  unpaper gefunden, aber Version nicht ermittelbar")
        else:
            print("  ⚠️  unpaper nicht gefunden (optional)")
    except Exception as e:
        print(f"  ✗ unpaper-Prüfung fehlgeschlagen: {_fmt_exc(e)}")

    results["unpaper"] = (unpaper_found, unpaper_version, unpaper_path)

    # 7. PDF-Rasterizer
    print("\n[7] Prüfe PDF-Rasterizer ...")
    if HAS_PYPDFIUM2:
        v = getattr(pypdfium2, "__version__", "unbekannt")
        print(f"  ✓ pypdfium2 verfügbar: {v}")
    else:
        print("  ⚠️  pypdfium2 nicht verfügbar (pip install pypdfium2)")
    results["pypdfium2"] = (HAS_PYPDFIUM2, "", "")

    pdfium_path = shutil.which("pdfium") or ""
    results["pdfium"] = (bool(pdfium_path), "", pdfium_path)
    if pdfium_path:
        print(f"  ✓ PDFium (extern): {pdfium_path}")
    else:
        print("  ⚠️  PDFium nicht verfügbar (nicht erforderlich)")

    # 8. Visual C++ Redistributable
    print("\n[8] Prüfe Microsoft Visual C++ Redistributable ...")
    vc_found    = False
    vc_versions = []

    vc_paths = [
        r"SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64",
        r"SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64",
        r"SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x86",
    ]
    for key_path in vc_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                installed = winreg.QueryValueEx(key, "Installed")[0]
                if installed == 1:
                    version = winreg.QueryValueEx(key, "Version")[0]
                    arch = "x64" if "x64" in key_path else "x86"
                    vc_versions.append(f"VC++ Redistributable {arch}: {version}")
                    vc_found = True
        except Exception:
            continue

    if vc_found:
        for v in vc_versions:
            print(f"  ✓ {v}")
    else:
        print("  ⚠️  Visual C++ Redistributable nicht eindeutig gefunden")
    results["vc_redist"] = (vc_found, "; ".join(vc_versions), "")

    # 9. ocrmypdf
    print("\n[9] Prüfe ocrmypdf ...")
    ocrmypdf_major = 0
    try:
        ocrmypdf_ver = ocrmypdf.__version__
        print(f"  ✓ ocrmypdf Version: {ocrmypdf_ver}")
        try:
            ocrmypdf_major = int(ocrmypdf_ver.split(".")[0])
        except (ValueError, IndexError):
            ocrmypdf_major = 0
        if ocrmypdf_major >= 17:
            print("  ℹ️  Version >= 17: pypdfium2 nativ integriert, Ghostscript optional")
    except Exception as e:
        print(f"  ✗ ocrmypdf-Prüfung fehlgeschlagen: {_fmt_exc(e)}")
    results["ocrmypdf_major"] = ocrmypdf_major

    # 10. pikepdf
    print("\n[10] Prüfe pikepdf ...")
    pikepdf_found = False
    try:
        import pikepdf
        pikepdf_ver = getattr(pikepdf, "__version__", "unbekannt")
        pikepdf_found = True
        print(f"  ✓ pikepdf Version: {pikepdf_ver}")
    except ImportError:
        print("  ⚠️  pikepdf nicht gefunden")
        print("     → Wird von ocrmypdf mitinstalliert (pip install ocrmypdf)")
    results["pikepdf"] = (pikepdf_found, "", "")

    # 11. veraPDF
    print("\n[11] Prüfe veraPDF ...")
    verapdf_found = False
    verapdf_path  = ""

    try:
        # .exe-Varianten bevorzugt: .bat-Aufrufe laufen durch cmd.exe,
        # das Sonderzeichen (%, ^, &) in Dateipfaden expandieren kann.
        verapdf_candidates = [
            shutil.which("verapdf.exe"),
            shutil.which("verapdf"),
            shutil.which("verapdf.bat"),
            os.path.join(exe_dir, "verapdf.exe"),
            os.path.join(exe_dir, "verapdf.bat"),
            os.path.join(exe_dir, "veraPDF", "verapdf.exe"),
            os.path.join(exe_dir, "veraPDF", "verapdf.bat"),
            r"C:\OCR\veraPDF\verapdf.exe",
            r"C:\OCR\veraPDF\verapdf.bat",
            r"C:\OCR\verapdf.exe",
            r"C:\OCR\verapdf.bat",
            r"C:\Program Files\veraPDF\verapdf.exe",
            r"C:\Program Files\veraPDF\verapdf.bat",
            r"C:\Program Files (x86)\veraPDF\verapdf.bat",
            r"C:\veraPDF\verapdf.exe",
            r"C:\veraPDF\verapdf.bat",
        ]
        for p in verapdf_candidates:
            if p and os.path.exists(p):
                verapdf_path = p
                break
        if verapdf_path:
            r = _run([verapdf_path, "--version"], timeout=10)
            if r.returncode == 0:
                verapdf_found = True
                print(f"  ✓ {verapdf_path}")
                print(f"  Version: {r.stdout.strip()[:80]}")
            else:
                print("  ⚠️  veraPDF gefunden, aber Version nicht ermittelbar")
        else:
            if ocrmypdf_major >= 17:
                print("  ⚠️  veraPDF nicht gefunden")
                print("     → Für PDF/A-Validierung empfohlen: https://verapdf.org/")
            else:
                print("  ℹ️  veraPDF nicht gefunden (optional)")
    except Exception as e:
        print(f"  ✗ veraPDF-Prüfung fehlgeschlagen: {_fmt_exc(e)}")

    results["verapdf"] = (verapdf_found, "", verapdf_path)

    print("=" * 80)
    return results


# ==================================================================
# Rasterizer-Auswahl
# ==================================================================

def _detect_pypdfium2_plugin() -> Optional[str]:
    import importlib
    candidates = [
        "ocrmypdf._plugins.pypdfium2",
        "ocrmypdf.builtin_plugins.pypdfium2",
    ]
    for candidate in candidates:
        try:
            importlib.import_module(candidate)
            return candidate
        except (ImportError, ModuleNotFoundError):
            continue
    return None


def determine_best_rasterizer(deps: Dict) -> Tuple[str, Optional[str]]:
    if PDF_RASTERIZER != "auto":
        return PDF_RASTERIZER, None

    if deps["pypdfium2"][0]:
        plugin = _detect_pypdfium2_plugin()
        if plugin:
            print(f"  → Rasterizer: pypdfium2 (Plugin: {plugin})")
            return "pypdfium2", plugin
        else:
            print("  → Rasterizer: pypdfium2 (nativ integriert)")
            return "auto", None

    if deps["pdfium"][0]:
        print("  → Rasterizer: PDFium")
        return "pdfium", None

    if deps["ghostscript"][0]:
        print("  → Rasterizer: Ghostscript (Fallback)")
        return "ghostscript", None

    print("  ⚠️  Kein PDF-Rasterizer verfügbar!")
    return "none", None


# ==================================================================
# Validierung und Konfiguration
# ==================================================================

def validate_dependencies(deps: Dict) -> Tuple[bool, Dict]:
    print("\n" + "=" * 80)
    print("  ABHÄNGIGKEITSVALIDIERUNG & KONFIGURATION")
    print("=" * 80)

    critical_missing = []
    warnings         = []
    config: Dict[str, Any] = {}

    ocrmypdf_major     = deps.get("ocrmypdf_major", 0)
    has_pypdfium2      = deps["pypdfium2"][0]
    has_ghostscript    = deps["ghostscript"][0]
    has_verapdf        = deps.get("verapdf", (False,))[0]
    has_pikepdf        = deps.get("pikepdf", (False,))[0]

    if not deps["tesseract"][0]:
        critical_missing.append("Tesseract OCR")
    if not deps["leptonica"][0]:
        warnings.append(
            "Keine separate Leptonica-DLL gefunden – beim UB-Mannheim-Tesseract "
            "unkritisch (statisch gelinkt); nur relevant, falls der ocrmypdf-"
            "Systemtest fehlschlägt."
        )
    if not has_pikepdf:
        warnings.append(
            "pikepdf fehlt – PDF/A-Marker können bei bereits verarbeiteten "
            "Dateien nicht gesetzt werden!"
        )

    if not has_ghostscript:
        if ocrmypdf_major >= 17 and has_pypdfium2:
            warnings.append(
                "Ghostscript nicht gefunden – nicht kritisch (ocrmypdf >= 17 + pypdfium2)"
            )
            if not has_verapdf:
                warnings.append(
                    "veraPDF nicht gefunden – PDF/A-Erzeugung ohne Ghostscript "
                    "kann stillschweigend auf Standard-PDF zurückfallen!"
                )
        else:
            critical_missing.append("Ghostscript")

    config["ocrmypdf_major"] = ocrmypdf_major
    config["has_verapdf"]    = has_verapdf
    config["verapdf_path"]   = deps.get("verapdf", (False, "", ""))[2]
    config["has_pikepdf"]    = has_pikepdf

    config["rasterizer"], config["rasterizer_plugin"] = determine_best_rasterizer(deps)

    available_langs = deps.get("tesseract_langs", [])
    if available_langs:
        if DEFAULT_LANGUAGE not in available_langs:
            print(f"  ⚠️  Standardsprache '{DEFAULT_LANGUAGE}' nicht verfügbar")
            for variant in ["deu", "deu_frak", "deu_latf", "ger", "de"]:
                if variant in available_langs:
                    print(f"  → Alternative: {variant}")
                    config["ocr_language"] = variant
                    break
            else:
                fallback = "eng" if "eng" in available_langs else (
                    available_langs[0] if available_langs else "eng"
                )
                print(f"  → Sprach-Fallback: {fallback}")
                config["ocr_language"] = fallback
        else:
            config["ocr_language"] = DEFAULT_LANGUAGE

        primary_lang = config.get("ocr_language")
        available_fallbacks = [
            lang for lang in FALLBACK_LANGUAGES
            if lang in available_langs and lang != primary_lang
        ]
        if available_fallbacks:
            config["ocr_language"] += "+" + "+".join(available_fallbacks)
            print(f"  → Mehrsprachige OCR: {config['ocr_language']}")
    else:
        config["ocr_language"] = DEFAULT_LANGUAGE

    config["optimizations"] = {
        "jbig2":    deps["jbig2"][0],
        "pngquant": deps["pngquant"][0],
        "unpaper":  deps["unpaper"][0],
    }
    if deps["jbig2"][0]:    print("  ✓ JBIG2-Optimierung verfügbar")
    if deps["pngquant"][0]: print("  ✓ PNG-Optimierung verfügbar")
    if deps["unpaper"][0]:  print("  ✓ Unpaper-Scanaufbereitung verfügbar")

    if critical_missing:
        print("\n❌ KRITISCHE FEHLER – folgende Programme fehlen:")
        for m in critical_missing:
            print(f"  • {m}")
        print("\nInstallationsanleitung:")
        step = 1
        if "Tesseract OCR" in critical_missing:
            print(f"  {step}. Tesseract: https://github.com/UB-Mannheim/tesseract/wiki")
            print("     → 'Additional language data' (DEU+ENG+FRA+SPA) aktivieren")
            print("     → 'Add to PATH' aktivieren")
            step += 1
        if "Ghostscript" in critical_missing:
            print(f"  {step}. Ghostscript: https://ghostscript.com/releases/gsdnld.html")
            print("     → AGPL-Version, Windows 64-bit")
            if ocrmypdf_major >= 17:
                print("     → Alternativ: pypdfium2 (pip install pypdfium2)")
                print("       + veraPDF für PDF/A: https://verapdf.org/")
            step += 1
        print(f"  {step}. System neu starten")
        return False, config

    print("\n✅ ALLE KRITISCHEN ABHÄNGIGKEITEN VORHANDEN")
    if warnings:
        print("\n⚠️  WARNUNGEN:")
        for w in warnings:
            print(f"  • {w}")

    # Systemtest
    print("\n🔧 ocrmypdf-Systemtest ...")
    test_pdf    = os.path.join(tempfile.gettempdir(), "test_ocr.pdf")
    test_output = os.path.join(tempfile.gettempdir(), "test_ocr_out.pdf")
    try:
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((50, 50), "Systemtest", fontsize=12)
            doc.save(test_pdf)

        test_opts = {
            "skip_text": False,
            "force_ocr": True,
            "output_type": "pdf",
            "optimize": 1,
            "jobs": 1,
        }

        with io.StringIO() as buf, contextlib.redirect_stderr(buf):
            try:
                ocrmypdf.ocr(
                    test_pdf, test_output,
                    language=config.get("ocr_language", DEFAULT_LANGUAGE),
                    **test_opts,
                )
                print("  ✅ ocrmypdf-Systemtest erfolgreich")
            except Exception as e:
                captured = buf.getvalue().strip()
                msg = _fmt_exc(e)
                if captured:
                    msg += f" | {captured[:200]}"
                if "leptonica" in msg.lower():
                    print("  ❌ KRITISCH: Leptonica-Fehler – DLL fehlt!")
                    return False, config
                print(f"  ⚠️  Systemtest mit Warnung: {msg[:400]}")
    except Exception as e:
        print(f"  ⚠️  Systemtest fehlgeschlagen: {_fmt_exc(e)}")
    finally:
        safe_remove(test_pdf)
        safe_remove(test_output)

    print("=" * 80)
    return True, config


# ==================================================================
# Output-Type-Mapping und Marker-Strategie
# ==================================================================

def map_output_type_for_ocrmypdf(output_type: str) -> str:
    if output_type in ("pdfa-2u", "pdfa-2b", "pdfa"):
        return "pdfa-2"
    if output_type in ("pdfa-1", "pdfa-2", "pdfa-3", "pdf"):
        return output_type
    return "pdf"


def marker_for_output_type(output_type: str) -> str:
    # Vorläufiger Marker (für ocrmypdf-Hint); finaler Marker via resolve_pdfa_marker
    if output_type == "pdfa-2u":
        return MARKER_VALUE_PDFA_2U
    if output_type == "pdfa-2b":
        return MARKER_VALUE_PDFA_2B
    if output_type.startswith("pdfa"):
        return MARKER_VALUE_PDFA_2B
    return MARKER_VALUE_PDF


def marker_code_for_output_type(output_type: str) -> str:
    if output_type == "pdfa-2u":
        return "PDFA_2U"
    if output_type == "pdfa-2b":
        return "PDFA_2B"
    if output_type.startswith("pdfa"):
        return "PDFA_2B"
    return "PDF"


def _compose_subject(marker_value: str, old_subject: str) -> str:
    # Vorhandene INHALTLICHE Subject-Metadaten nicht zerstoeren: alte
    # Marker-Varianten werden entfernt, der verbleibende Beschreibungstext
    # hinter dem neuen Marker angehaengt. Die Marker-Erkennung arbeitet
    # mit Substring-Matching und funktioniert daher unveraendert.
    old = str(old_subject or "")
    for m in (MARKER_VALUE_PDFA_2U, MARKER_VALUE_PDFA_2B,
              MARKER_VALUE_PDF, LEGACY_MARKER):
        old = old.replace(m, "")
    old = old.strip().strip("|–-").strip()
    if old:
        return f"{marker_value} | {old}"
    return marker_value


def resolve_pdfa_marker(
    output_type: str,
    has_verapdf: bool,
    verified_code: Optional[str],
) -> Tuple[str, str, Optional[str]]:
    # Zentrale Marker-Strategie: PDFA_2U nur wenn veraPDF bestätigt hat.
    # Returns: (marker_value, marker_code, warning_message_or_None)
    if not output_type.startswith("pdfa"):
        return MARKER_VALUE_PDF, "PDF", None

    if output_type == "pdfa-2u":
        if has_verapdf:
            if verified_code == "PDFA_2U":
                return MARKER_VALUE_PDFA_2U, "PDFA_2U", None
            if verified_code == "PDFA_2B":
                return (MARKER_VALUE_PDFA_2B, "PDFA_2B",
                        "veraPDF: nur PDF/A-2b erreicht (ToUnicode-CMaps unvollständig)")
            return (MARKER_VALUE_PDF, "PDF",
                    "veraPDF: Keine PDF/A-Konformität validiert – Standard-PDF-Marker")
        return (MARKER_VALUE_PDFA_2B, "PDFA_2B",
                "veraPDF nicht verfügbar – konservativ PDF/A-2b statt 2u markiert")

    if output_type == "pdfa-2b":
        if has_verapdf and verified_code is None:
            return (MARKER_VALUE_PDF, "PDF",
                    "veraPDF: PDF/A-2b nicht validiert – Standard-PDF-Marker")
        return MARKER_VALUE_PDFA_2B, "PDFA_2B", None

    return MARKER_VALUE_PDFA_2B, "PDFA_2B", None


# ==================================================================
# ocrmypdf-Logger-Handler (singleton, pro Worker)
# ==================================================================

class _OcrMypdfLogCapture:
    def __init__(self) -> None:
        self._buf = io.StringIO()
        self._handler = logging.StreamHandler(self._buf)
        self._handler.setLevel(logging.WARNING)
        self._logger = logging.getLogger("ocrmypdf")
        self._attached = False

    def __enter__(self):
        self._buf.seek(0)
        self._buf.truncate(0)
        if not self._attached:
            self._logger.addHandler(self._handler)
            self._attached = True
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            self._logger.removeHandler(self._handler)
        finally:
            self._attached = False
        return False

    def get_output(self) -> str:
        return self._buf.getvalue().strip()


# ==================================================================
# PDF-Dateiverarbeitung (Worker-kompatibel)
# ==================================================================

def process_pdf_file(
    file_path: str,
    config: Dict,
    temp_dir: str,
) -> _ProcessResult:
    start_time = time.time()
    messages:    List[str]       = []
    log_entries: List[_LogEntry] = []

    _worker_buffer_clear()

    def pwrite(msg: str) -> None:
        messages.append(msg)

    def linfo(ctx: str, msg: str) -> None:
        log_entries.append(_LogEntry("INFO", ctx, msg))

    def lwarn(ctx: str, msg: str) -> None:
        log_entries.append(_LogEntry("WARN", ctx, msg))

    def lerr(ctx: str, msg: str, exc_info: bool = False) -> None:
        if exc_info:
            msg += "\n" + traceback.format_exc()
        log_entries.append(_LogEntry("ERROR", ctx, msg, False))

    def build_result(status: str, detail: str,
                     size_mb: float = 0.0, pages: int = 0,
                     error_msg: Optional[str] = None,
                     size_after_mb: float = 0.0) -> _ProcessResult:
        log_entries.extend(_worker_buffer_drain())
        return _ProcessResult(
            status=status, detail=detail,
            messages=messages, log_entries=log_entries,
            duration_sec=time.time() - start_time,
            size_mb=size_mb, pages=pages,
            error_msg=error_msg, file_path=file_path,
            size_after_mb=size_after_mb,
        )

    file_name   = os.path.basename(file_path)
    log_context = f"process_pdf:{file_name}"

    pwrite(f"\n📄 Verarbeite: {file_name}")
    linfo(log_context, f"Starte: {file_path}")

    # --- Grundlegende Prüfungen ---
    if not safe_exists(file_path):
        pwrite("  ✗ Datei existiert nicht")
        lerr(log_context, f"Existiert nicht: {file_path}")
        return build_result("ERROR", "ERROR_MISSING", error_msg="Datei existiert nicht")

    if not is_real_pdf(file_path):
        pwrite("  ⚠️  Keine gültige PDF-Signatur (%PDF-) – überspringe")
        lwarn(log_context, f"Keine PDF-Signatur: {file_path}")
        return build_result("SKIPPED", "SKIP_NOT_PDF",
                            size_mb=get_file_size_mb(file_path))

    file_size_mb = get_file_size_mb(file_path)
    required_mb = required_disk_mb_for_pdf(file_size_mb)
    if not check_disk_space(temp_dir, required_mb=required_mb):
        pwrite(f"  ✗ Zu wenig Speicherplatz im Temp-Laufwerk (benötigt: ~{required_mb:.0f} MB)")
        lerr(log_context, f"Kritischer Speichermangel auf {temp_dir} (benötigt ~{required_mb:.0f} MB)")
        return build_result("ERROR", "ERROR_DISK", error_msg="Speichermangel")

    if is_file_locked(file_path):
        pwrite("  ⚠️  Datei gesperrt – warte ...")
        if not wait_for_file_unlock(file_path):
            pwrite("  ✗ Bleibt gesperrt – überspringe")
            lwarn(log_context, f"Dauerhaft gesperrt: {file_path}")
            return build_result("SKIPPED", "SKIP_LOCKED",
                                size_mb=file_size_mb)

    # --- PDF-Infos ---
    pdf_info = get_pdf_info(file_path)
    if not pdf_info["is_valid"]:
        pwrite(f"  ✗ Ungültige PDF: {pdf_info.get('error', '?')}")
        lerr(log_context, f"Ungültige PDF: {pdf_info.get('error')}")
        return build_result("ERROR", "ERROR_INVALID",
                            size_mb=pdf_info["size_mb"],
                            error_msg=pdf_info.get("error"))

    pwrite(f"  ℹ️  Seiten: {pdf_info['pages']}, "
           f"Größe: {pdf_info['size_mb']:.2f} MB")

    if pdf_info["is_encrypted"]:
        # Owner-Only-Verschluesselung (leeres Benutzerpasswort) ist bei
        # Scanner-PDFs haeufig und laesst sich verlustfrei entfernen.
        pwrite("  🔐 Verschlüsselt – prüfe auf leeres Benutzerpasswort ...")
        if try_remove_empty_password(file_path, temp_dir):
            pwrite("  🔓 Owner-Only-Verschlüsselung entfernt – Verarbeitung möglich")
            linfo(log_context, f"Leeres Benutzerpasswort – Verschluesselung entfernt: {file_path}")
            pdf_info = get_pdf_info(file_path)
            if not pdf_info["is_valid"] or pdf_info["is_encrypted"]:
                pwrite("  ⚠️  Nach Entschlüsselung nicht lesbar – überspringe")
                lwarn(log_context, f"Nach Entschluesselung nicht lesbar: {file_path}")
                return build_result("SKIPPED", "SKIP_ENCRYPTED",
                                    size_mb=pdf_info["size_mb"],
                                    pages=pdf_info["pages"])
        else:
            pwrite("  ⚠️  Verschlüsselt (echtes Benutzerpasswort) – überspringe")
            lwarn(log_context, f"Verschluesselt: {file_path}")
            return build_result("SKIPPED", "SKIP_ENCRYPTED",
                                size_mb=pdf_info["size_mb"],
                                pages=pdf_info["pages"])

    # Schutzpruefungen als Block, damit sie nach einer Reparatur ERNEUT laufen
    # koennen. get_pdf_info() kehrt bei fitz.FileDataError sofort zurueck und
    # liefert dann needs_repair=True, is_valid=True, aber pages=0,
    # has_forms=False, is_oversized=False und max_image_pixels=0. Auf diesen
    # leeren Werten gingen alle Pruefungen unten folgenlos durch - der
    # Formular-, Uebergroessen-, Bildpixel- und RAM-Schutz griff bei genau den
    # Dateien nicht, die am ehesten Probleme machen.
    def _schutzpruefungen(info):
        if info["has_forms"]:
            pwrite("  ⚠️  Enthält Formularfelder – wird zum Schutz übersprungen")
            lwarn(log_context, f"Formularfelder erkannt – uebersprungen: {file_path}")
            return build_result("SKIPPED", "SKIP_HAS_FORMS",
                                size_mb=info["size_mb"], pages=info["pages"])

        if info["is_oversized"]:
            pwrite(f"  ⚠️  Physische Dimensionen zu groß (>{MAX_PAGE_DIMENSION} pt ≈ "
                   f"{MAX_PAGE_DIMENSION / 72:.0f} Zoll) – überspringe")
            lwarn(log_context, f"Übergröße: {file_path}")
            return build_result("SKIPPED", "SKIP_OVERSIZED",
                                size_mb=info["size_mb"], pages=info["pages"])

        max_img_px = info.get("max_image_pixels", 0)
        if max_img_px > MAX_EMBEDDED_IMAGE_PIXELS:
            pwrite(f"  ⚠️  Eingebettetes Bild zu groß ({max_img_px / 1_000_000:.0f} MP "
                   f"> {MAX_EMBEDDED_IMAGE_PIXELS / 1_000_000:.0f} MP Pillow-Limit) – überspringe")
            lwarn(log_context,
                  f"Übergroßes Bild: {file_path} – max_image_pixels={max_img_px}")
            return build_result("SKIPPED", "SKIP_OVERSIZED_IMAGE",
                                size_mb=info["size_mb"], pages=info["pages"])

        # --- Memory-Check (gelockert; Skip nur bei physischer Unmöglichkeit) ---
        mem_ok, mem_msg = check_memory_before_ocr(
            info["pages"], jobs=config.get("jobs", 1))
        if not mem_ok:
            pwrite(f"  ⚠️  RAM-Schutz aktiv – überspringe ({mem_msg})")
            lwarn(log_context, f"Memory-Schutz: {file_path} – {mem_msg}")
            return build_result("SKIPPED", "SKIP_MEMORY",
                                size_mb=info["size_mb"], pages=info["pages"],
                                error_msg=mem_msg)
        return None

    _abbruch = _schutzpruefungen(pdf_info)
    if _abbruch is not None:
        return _abbruch

    # --- Beschädigte Struktur reparieren ---
    if pdf_info["needs_repair"]:
        pwrite("  🔧 Beschädigte PDF-Struktur – versuche Reparatur ...")
        if repair_pdf(file_path, temp_dir):
            pwrite("  ✓ Reparatur erfolgreich")
            linfo(log_context, f"Repariert: {file_path}")
            # Erst jetzt sind Seitenzahl, Formularfelder, Bildgroessen und
            # Marker ueberhaupt lesbar - Schutzpruefungen deshalb wiederholen.
            pdf_info = get_pdf_info(file_path)
            if not pdf_info["is_valid"]:
                pwrite("  ✗ Nach Reparatur nicht lesbar – überspringe")
                lwarn(log_context, f"Nach Reparatur nicht lesbar: {file_path}")
                return build_result("ERROR", "ERROR_INVALID",
                                    size_mb=pdf_info["size_mb"],
                                    pages=pdf_info["pages"])
            _abbruch = _schutzpruefungen(pdf_info)
            if _abbruch is not None:
                return _abbruch
        else:
            pwrite("  ⚠️  Reparatur fehlgeschlagen – versuche OCR trotzdem")
            lwarn(log_context, f"Reparatur fehlgeschlagen: {file_path}")

    # --- Marker-Prüfung (mit optionaler Re-Verifikation) ---
    output_type    = config.get("output_type", DEFAULT_OUTPUT_TYPE)
    target_is_pdfa = output_type.startswith("pdfa")
    has_verapdf    = bool(config.get("has_verapdf"))
    found_marker   = pdf_info.get("existing_marker")

    if found_marker in ("PDFA_2U", "PDFA_2B") and config.get("verify_markers"):
        pwrite("  🔎 Verifiziere vorhandenen PDF/A-Marker gegen veraPDF ...")
        ok, actual_code = verify_existing_marker(
            file_path, found_marker, config, pages=pdf_info["pages"]
        )
        if not ok:
            if actual_code is None:
                pwrite(f"  ⚠️  Marker '{found_marker}' konnte nicht validiert werden – "
                       f"behandle als unmarkiert")
                lwarn(log_context, f"Marker-Re-Verifikation fehlgeschlagen: {file_path}")
                found_marker = None
            else:
                pwrite(f"  ⚠️  Marker-Korrektur: {found_marker} → {actual_code}")
                lwarn(log_context,
                      f"Marker-Korrektur: {found_marker} → {actual_code} ({file_path})")
                if actual_code == "PDFA_2B":
                    set_ocr_marker(file_path, _compose_subject(
                        MARKER_VALUE_PDFA_2B, str(pdf_info.get("subject") or "")))
                found_marker = actual_code

    if found_marker == "PDFA_2U":
        pwrite("  ✓ Bereits im Zielformat PDF/A-2u (Metadaten-Marker)")
        linfo(log_context, "Übersprungen – PDF/A-2u-Marker vorhanden")
        return build_result("SKIPPED", "SKIP_MARKER_PDFA_2U",
                            size_mb=pdf_info["size_mb"],
                            pages=pdf_info["pages"])

    if found_marker == "PDFA_2B":
        if output_type == "pdfa-2u":
            pass
        else:
            pwrite("  ✓ Bereits im Format PDF/A-2b (Metadaten-Marker)")
            linfo(log_context, "Übersprungen – PDF/A-2b-Marker vorhanden")
            return build_result("SKIPPED", "SKIP_MARKER_PDFA_2B",
                                size_mb=pdf_info["size_mb"],
                                pages=pdf_info["pages"])

    if found_marker in ("PDF", "LEGACY") and not target_is_pdfa:
        pwrite("  ✓ Bereits als Standard-PDF verarbeitet (Metadaten-Marker)")
        linfo(log_context, "Übersprungen – PDF-Marker vorhanden")
        return build_result("SKIPPED", "SKIP_MARKER_PDF",
                            size_mb=pdf_info["size_mb"],
                            pages=pdf_info["pages"])

    is_upgrade = False
    if found_marker == "PDFA_2B" and output_type == "pdfa-2u":
        pwrite("  ℹ️  Upgrade von PDF/A-2b auf PDF/A-2u ...")
        linfo(log_context, "Upgrade PDF/A-2b → PDF/A-2u")
        is_upgrade = True
    elif found_marker in ("PDF", "LEGACY") and target_is_pdfa:
        fmt_name = "PDF/A-2u" if output_type == "pdfa-2u" else "PDF/A-2b"
        pwrite(f"  ℹ️  Konvertierung nach {fmt_name} ...")
        linfo(log_context, f"Rekonvertierung nach {fmt_name}")
        is_upgrade = True

    if not found_marker and pdf_info["has_text"]:
        pwrite("  ℹ️  Enthält bereits Text (ocrmypdf entscheidet mit skip_text=True)")

    # --- Temporäre Kopie als Input (mit AV-Scanner-Toleranz) ---
    temp_input = create_temp_copy(file_path, temp_dir)
    if not temp_input:
        pwrite("  ✗ Temp-Kopie nicht erstellbar")
        lerr(log_context, "Temp-Kopie fehlgeschlagen")
        return build_result("ERROR", "ERROR_TEMP_COPY",
                            size_mb=pdf_info["size_mb"],
                            pages=pdf_info["pages"],
                            error_msg="Temp-Kopie fehlgeschlagen")

    temp_output = os.path.join(temp_dir, f"{uuid.uuid4().hex}_out.pdf")

    # --- OCR-Optionen ---
    current_options = OCR_OPTIONS.copy()
    ocr_language    = config.get("ocr_language", DEFAULT_LANGUAGE)
    optimizations   = config.get("optimizations", {})

    actual_output_type = map_output_type_for_ocrmypdf(output_type)
    is_pdfa = actual_output_type.startswith("pdfa")

    current_options["output_type"] = actual_output_type
    current_options["jobs"] = config.get("jobs", 1)

    rasterizer_plugin = config.get("rasterizer_plugin")
    if rasterizer_plugin:
        current_options["plugins"] = [rasterizer_plugin]

    if is_pdfa:
        current_options.pop("fast_web_view", None)
        # KEINE proaktive RGB-Konvertierung mehr: sie blaehte Graustufen-/
        # Bitonal-Scans unnoetig auf. Farbraum-Probleme werden reaktiv im
        # Retry behandelt (COLOR_SPACE_ERR_KEYWORDS + Sicherheitsnetz im
        # letzten Versuch).

    if not optimizations.get("jbig2", False):
        current_options["jbig2_lossy"] = False
    if not optimizations.get("pngquant", False):
        current_options.pop("png_quality", None)
    if not optimizations.get("unpaper", False):
        current_options["unpaper_args"] = None
        current_options["clean"] = False

    # --- Bildschonendes Profil fuer Druckvorlagen -------------------------
    # Muss NACH den Optimierungs-Schaltern stehen, damit es sie ueberstimmt.
    print_safe, print_safe_reason = needs_print_safe_profile(pdf_info)
    if config.get("force_print_safe"):
        print_safe = True
        print_safe_reason = print_safe_reason or "per Schalter erzwungen"
    if config.get("no_print_safe"):
        print_safe = False

    if print_safe:
        current_options.update(PRINT_SAFE_OPTIONS)
        current_options.pop("png_quality", None)
        # ocrmypdf lehnt pdfa_image_compression != "auto" ab, wenn die Ausgabe
        # kein PDF/A ist (builtin_plugins/ghostscript.py) -> BadArgsError.
        if not is_pdfa:
            current_options.pop("pdfa_image_compression", None)
        pwrite(f"  🖼️  Bildschonendes Profil aktiv ({print_safe_reason})")
        pwrite("      deskew/clean aus, keine Bild-Neukodierung – "
               "Auflösung und Farbraum bleiben erhalten")
        linfo(log_context,
              f"Bildschonendes Profil: {print_safe_reason}; "
              f"max. Bild-DPI {pdf_info.get('max_image_dpi', 0):.0f}, "
              f"CMYK={pdf_info.get('has_cmyk_images')}")

    ocr_successful   = False
    ocr_error_msg    = None
    last_error_kind  = None
    backup_file      = file_path + ".backup"

    # Harte Abbrüche (Signatur, korrupter Stream): Status, Detail, Fehlermeldung
    hard_status: Optional[Tuple[str, str, Optional[str]]] = None

    # Vorlaeufiger Marker als ocrmypdf-Hint (subject in PDF-Metadaten).
    # Ein vorhandenes inhaltliches Subject wird NICHT zerstoert, sondern
    # hinter dem Marker erhalten (_compose_subject).
    original_subject    = str(pdf_info.get("subject") or "")
    preliminary_marker  = marker_for_output_type(output_type)
    preliminary_subject = _compose_subject(preliminary_marker, original_subject)
    ocr_metadata = {"subject": preliminary_subject}
    if original_subject.strip() and preliminary_subject != preliminary_marker:
        linfo(log_context,
              f"Bisheriges Subject bleibt hinter dem Marker erhalten: '{original_subject[:150]}'")

    # Final ermittelter Marker nach veraPDF (oder Fallback-Strategie)
    final_marker_value: Optional[str] = None
    final_marker_code:  Optional[str] = None

    orig_times = _read_file_times(file_path)

    # NTFS-Sicherheitsinfo (Owner/Group/DACL) des Originals sichern –
    # wird nach der Datei-Ersetzung wieder angewendet (Admin-Kontext:
    # vollständig inkl. Owner; Nutzer-Kontext: DACL).
    orig_sd = _get_security_descriptor(file_path)

    color_fallback_applied  = False
    bitmap_fallback_applied = False

    for attempt in range(MAX_RETRIES):
        if safe_exists(temp_output):
            safe_remove(temp_output)

        pwrite(f"  🔧 OCR Versuch {attempt + 1}/{MAX_RETRIES} "
               f"(Sprache: {ocr_language}, Format: {output_type}) ...")

        # Adaptive Retry-Strategie
        if attempt > 0:
            pwrite("  ⚠️  Deaktiviere 'clean', 'deskew', 'unpaper'")
            current_options["clean"]        = False
            current_options["deskew"]       = False
            current_options["unpaper_args"] = None

            # RGB-Konvertierung bei CMYK-/Separations-Bildern NIE erzwingen:
            # Ghostscript schreibt die Bilder dann als DeviceRGB zurueck, die
            # Farbauszuege sind unwiederbringlich weg und die Datei fuer den
            # Druck unbrauchbar. Lieber als Fehler melden.
            rgb_forbidden = bool(pdf_info.get("has_cmyk_images"))

            if ocr_error_msg and any(
                k in ocr_error_msg.lower() for k in COLOR_SPACE_ERR_KEYWORDS
            ):
                if rgb_forbidden:
                    pwrite("  🎨 Farbraum-Problem, aber CMYK-Bilder vorhanden – "
                           "KEINE RGB-Konvertierung (Druckqualität)")
                    lwarn(log_context,
                          "Farbraum-Fallback unterdrückt: Datei enthält CMYK-/Separations-Bilder")
                elif not color_fallback_applied:
                    pwrite("  🎨 Farbraum-Konvertierung nach RGB aktiviert")
                    current_options["color_conversion_strategy"] = "RGB"
                    color_fallback_applied = True
            elif (attempt == MAX_RETRIES - 1 and is_pdfa
                  and not color_fallback_applied and not rgb_forbidden):
                pwrite("  🎨 Letzter Versuch: Farbraum-Konvertierung nach RGB (Sicherheitsnetz)")
                current_options["color_conversion_strategy"] = "RGB"
                color_fallback_applied = True

        backup_success = False
        try:
            if not safe_copy(file_path, backup_file):
                if safe_exists(backup_file):
                    safe_remove(backup_file)
                last_error_kind = "BACKUP"
                raise Exception("Backup-Erstellung fehlgeschlagen")

            backup_success = True

            with _OcrMypdfLogCapture() as ocr_cap:
                with io.StringIO() as err_buf, contextlib.redirect_stderr(err_buf):
                    ocrmypdf.ocr(
                        temp_input,
                        temp_output,
                        language=ocr_language,
                        **ocr_metadata,
                        **current_options,
                    )
                    err_output = err_buf.getvalue()
                ocr_log_output = ocr_cap.get_output()

            combined_warnings = "\n".join(
                part for part in (err_output, ocr_log_output) if part
            )
            if combined_warnings:
                for kw in WARN_KEYWORDS:
                    if kw in combined_warnings.lower():
                        lwarn(log_context, f"OCR-Warnung: {combined_warnings[:300]}")
                        break

            if os.path.exists(temp_output):
                # --- veraPDF-Validierung VOR dem Verschieben ---
                verified_code: Optional[str] = None
                if is_pdfa and has_verapdf:
                    verified_code = verify_pdfa_with_verapdf(
                        temp_output,
                        config.get("verapdf_path", ""),
                        output_type if output_type != "pdfa" else "pdfa-2",
                        pages=pdf_info["pages"],
                    )

                # --- Zentrale Marker-Strategie anwenden ---
                final_marker_value, final_marker_code, marker_warning = resolve_pdfa_marker(
                    output_type, has_verapdf, verified_code
                )

                if marker_warning:
                    pwrite(f"  ⚠️  {marker_warning}")
                    lwarn(log_context, marker_warning)
                elif is_pdfa and verified_code:
                    linfo(log_context, f"veraPDF: {verified_code} bestätigt")

                # --- Marker-Korrektur LOKAL auf temp_output ---
                # ocrmypdf hat den vorlaeufigen Marker bereits in die Ausgabe
                # geschrieben. Nur bei Abweichung ist ein Rewrite noetig --
                # und der passiert auf der lokalen Temp-Datei statt auf der
                # Netzfreigabe; im Normalfall (veraPDF bestaetigt das Ziel-
                # format) bleiben die validierten Bytes unveraendert.
                if final_marker_value:
                    final_subject = _compose_subject(final_marker_value, original_subject)
                    if final_subject != preliminary_subject:
                        if not set_ocr_marker(temp_output, final_subject):
                            lwarn(log_context, "OCR-Marker konnte nicht gesetzt werden")

                # --- Atomare Ersetzung temp_output → file_path ---
                # os.replace laesst das Original bei JEDEM Fehlschlag intakt;
                # ein Backup-Rueckspielen ist deshalb nicht mehr noetig.
                if not safe_replace_with_retry(temp_output, file_path):
                    last_error_kind = "MOVE"
                    raise Exception(f"Ersetzen von {file_path} fehlgeschlagen (SMB-Lock?)")

                safe_remove(backup_file)
                ocr_successful = True
                pwrite(f"  ✅ OCR erfolgreich (Marker: {final_marker_code or 'PDF'})")
                linfo(log_context, f"OCR erfolgreich – Marker: {final_marker_code}")
                break
            else:
                # OCR hat keine Ausgabedatei produziert. file_path wurde NIE
                # angetastet (ocrmypdf liest aus temp_input, schreibt nach
                # temp_output - beide sind temporaere Pfade). Backup deshalb
                # nur loeschen, NICHT zurueckspielen: Wuerden wir das Backup
                # ueber file_path schreiben, koennte das im Hotfolder-Modus
                # zwischenzeitlich vom Nutzer ersetzte Originale ueberschreiben
                # (z.B. User legt neue Version waehrend OCR laeuft).
                if backup_success and safe_exists(backup_file):
                    safe_remove(backup_file)
                last_error_kind = "NO_OUTPUT"
                raise Exception("Ausgabedatei wurde nicht erstellt")

        except PriorOcrFoundError:
            pwrite("  ✓ Bereits OCR-verarbeitet (ocrmypdf-Erkennung)")
            linfo(log_context, "PriorOcrFoundError – Datei bereits verarbeitet")
            # file_path wurde nicht modifiziert - Backup einfach loeschen.
            if backup_success and safe_exists(backup_file):
                safe_remove(backup_file)

            # --- Auch hier zentrale Marker-Strategie auf bestehender Datei ---
            verified_code = None
            if is_pdfa and has_verapdf:
                verified_code = verify_pdfa_with_verapdf(
                    file_path,
                    config.get("verapdf_path", ""),
                    output_type if output_type != "pdfa" else "pdfa-2",
                    pages=pdf_info["pages"],
                )

            final_marker_value, final_marker_code, marker_warning = resolve_pdfa_marker(
                output_type, has_verapdf, verified_code
            )
            if marker_warning:
                pwrite(f"  ⚠️  {marker_warning}")
                lwarn(log_context, marker_warning)

            target_code_now = found_marker
            if final_marker_value and target_code_now != final_marker_code:
                if set_ocr_marker(file_path,
                                  _compose_subject(final_marker_value, original_subject)):
                    pwrite(f"  ✓ OCR-Marker gesetzt ({final_marker_code})")
                    linfo(log_context, f"OCR-Marker gesetzt: {final_marker_code}")
                else:
                    lwarn(log_context, "OCR-Marker konnte nicht gesetzt werden")
                    pwrite("  ⚠️  OCR-Marker nicht setzbar")
            ocr_successful = True
            break

        except DigitalSignatureError:
            pwrite("  ⚠️  Digital signiertes PDF – OCR würde Signatur ungültig machen")
            lwarn(log_context, f"Digital signiert – übersprungen: {file_path}")
            # file_path wurde nicht modifiziert - Backup einfach loeschen.
            if backup_success and safe_exists(backup_file):
                safe_remove(backup_file)
            hard_status = ("SKIPPED", "SKIP_SIGNED", None)
            break

        except InputFileError as e:
            ocr_error_msg = _fmt_exc(e)
            pwrite(f"  ✗ PDF-Inhaltsstrom defekt – Retry chancenlos: {ocr_error_msg}")
            lerr(log_context, f"Korrupter PDF-Inhaltsstrom: {ocr_error_msg}")
            # file_path wurde nicht modifiziert - Backup einfach loeschen.
            if backup_success and safe_exists(backup_file):
                safe_remove(backup_file)
            hard_status = ("ERROR", "ERROR_INVALID_STREAM", ocr_error_msg)
            break

        except Exception as e:
            ocr_error_msg = _fmt_exc(e)
            err_lower = ocr_error_msg.lower()
            pwrite(f"  ✗ OCR-Fehler: {ocr_error_msg}")
            lerr(log_context, f"OCR fehlgeschlagen: {ocr_error_msg}", exc_info=True)

            # ocrmypdf hat file_path nicht angefasst (liest aus temp_input,
            # schreibt nach temp_output), und die finale Ersetzung laeuft
            # atomar ueber os.replace - auch im MOVE-Fehlerfall bleibt das
            # Original intakt. Backup daher in ALLEN Fehlerfaellen nur
            # loeschen, nie zurueckspielen: Rueckspielen koennte im
            # Hotfolder-Modus eine zwischenzeitlich vom Nutzer ersetzte
            # Datei ueberschreiben.
            if backup_success and safe_exists(backup_file):
                safe_remove(backup_file)

            # String-Match-Fallback (Klassen-Match versagt sporadisch in spawn-Workern)
            if (
                "digitalsignatureerror" in err_lower
                or "has a digital signature" in err_lower
            ):
                pwrite("  ⚠️  Digital signiertes PDF (String-Fallback) – übersprungen")
                lwarn(log_context, f"Digital signiert (Fallback): {file_path}")
                hard_status = ("SKIPPED", "SKIP_SIGNED", None)
                break

            if (
                "dynamic xfa forms" in err_lower
                or "livecycle designer" in err_lower
            ):
                pwrite("  ✗ Dynamische XFA-Form (Adobe LiveCycle) – nicht OCR-fähig")
                hard_status = ("ERROR", "ERROR_INVALID_STREAM", ocr_error_msg)
                break

            # Decompression-Bomb-Schutz wurde global deaktiviert; Restfälle defensiv abfangen.
            # Diese Dateien sind nicht defekt, sondern enthalten gigantische eingebettete
            # Bilder. Sauberer Skip statt Stream-Fehler – Pre-Check sollte sie eigentlich
            # vorher abfangen, das hier ist die Sicherheitsnetz-Klassifikation.
            if "decompressionbombe" in err_lower or "decompression bomb" in err_lower:
                pwrite("  ⚠️  Pillow-Decompression-Bomb (sehr großes Bild) – überspringe")
                hard_status = ("SKIPPED", "SKIP_OVERSIZED_IMAGE", ocr_error_msg)
                break

            # Strukturell defekte Streams → kein Retry sinnvoll
            if (
                "graphics stack overflowed" in err_lower
                or "content stream is corrupt" in err_lower
                or ("inputfileerror" in err_lower and "corrupt" in err_lower)
            ):
                pwrite("  ❌ Strukturell defekt – Abbruch ohne Retry")
                hard_status = ("ERROR", "ERROR_INVALID_STREAM", ocr_error_msg)
                break

            # Rasterizer am Speicherlimit → Plugin-Reset (Ghostscript-Fallback in ocrmypdf 17+)
            bitmap_problem = (
                isinstance(e, MemoryError)
                or "failed to fill bitmap" in err_lower
                or "memoryerror" in err_lower
            )
            if (
                bitmap_problem
                and not bitmap_fallback_applied
                and attempt < MAX_RETRIES - 1
            ):
                # KEIN oversample setzen. Der Parameter ist eine DPI-UNTER-
                # grenze, keine Reduktion: die CLI-Hilfe von ocrmypdf lautet
                # 'Oversample images to AT LEAST the specified DPI', und
                # _pipeline.py verrechnet den Wert in einem max(...) mit der
                # tatsaechlichen Bildaufloesung. 'oversample=200' konnte die
                # Aufloesung also nur ANHEBEN - und damit den Speicherbedarf
                # erhoehen, den dieser Rueckfall gerade senken soll. Bei
                # Scans mit 300 dpi blieb er wirkungslos, bei 150 dpi machte
                # er es schlimmer.
                # Wirksam ist stattdessen, die Parallelitaet auf einen Job zu
                # senken (jeder Job haelt seine eigene Rasterbitmap) und die
                # Obergrenze fuer Bildpixel zu druecken.
                pwrite("  🖼️  pypdfium2 am Speicherlimit – wechsle auf Ghostscript-Pfad, "
                       "1 Job, kleinere Bildobergrenze")
                current_options["plugins"] = []
                current_options["jobs"] = 1
                current_options["max_image_mpixels"] = min(
                    int(current_options.get("max_image_mpixels") or 250), 128)
                bitmap_fallback_applied = True

            # Farbraum-Problem (falls is_pdfa nicht greift, z.B. bei nicht-pdfa-Output)
            if (
                not color_fallback_applied
                and any(k in err_lower for k in COLOR_SPACE_ERR_KEYWORDS)
                and attempt < MAX_RETRIES - 1
            ):
                if pdf_info.get("has_cmyk_images"):
                    pwrite("  🎨 Farbraum-Problem, aber CMYK-Bilder vorhanden – "
                           "KEINE RGB-Konvertierung (Druckqualität)")
                    lwarn(log_context,
                          "Farbraum-Fallback unterdrückt (CMYK/Separation vorhanden)")
                else:
                    pwrite("  🎨 Farbraum-Problem erkannt – nächster Versuch mit RGB-Konvertierung")
                    current_options["color_conversion_strategy"] = "RGB"
                    color_fallback_applied = True

            if attempt < MAX_RETRIES - 1:
                pwrite(f"  ⏳ Warte {RETRY_DELAY}s vor Wiederholung ...")
                time.sleep(RETRY_DELAY)
            else:
                pwrite("  ❌ Maximale Wiederholungen erreicht")

    # --- Aufräumen ---
    safe_remove(temp_input)
    safe_remove(temp_output)

    # Harte Abbruch-Fälle (Signatur, korrupter Stream)
    if hard_status:
        status, detail, err = hard_status
        return build_result(status, detail,
                            size_mb=pdf_info["size_mb"],
                            pages=pdf_info["pages"],
                            error_msg=err)

    if ocr_successful:
        _apply_security_descriptor(file_path, orig_sd)
    if ocr_successful and orig_times:
        _write_file_times(file_path, orig_times)

    if ocr_successful:
        status = "UPGRADED" if is_upgrade else "PROCESSED"
        detail = "OK_UPGRADE" if is_upgrade else "OK"
        return build_result(status, detail,
                            size_mb=pdf_info["size_mb"],
                            pages=pdf_info["pages"],
                            size_after_mb=get_file_size_mb(file_path))
    else:
        if last_error_kind == "BACKUP":
            detail = "ERROR_BACKUP"
        elif last_error_kind == "MOVE":
            detail = "ERROR_MOVE"
        else:
            detail = "ERROR_OCR"
        return build_result("ERROR", detail,
                            size_mb=pdf_info["size_mb"],
                            pages=pdf_info["pages"],
                            error_msg=ocr_error_msg)


# ==================================================================
# Generator + Verzeichnis-Verarbeitung
# ==================================================================

def pdf_generator(directory: str, skip_set: Optional[Set[str]] = None):
    # Walk-Root IMMER mit \\?\ praefixieren: ohne LongPathsEnabled wuerde
    # os.walk sonst tiefe Baeume (> 260 Zeichen) still ueberspringen.
    # Nach aussen werden die Pfade wieder ohne Praefix geliefert (saubere
    # Anzeige/CSV/Resume); _lp() praefixiert bei Bedarf erneut.
    safe_dir = prepare_long_path(directory)
    skip_set = skip_set or set()

    def _walk_error(err: OSError) -> None:
        name = _strip_long_path(err.filename or "")
        log_warning("pdf_generator",
                    f"Verzeichnis nicht lesbar – übersprungen: {name}")

    for root, dirs, files in os.walk(safe_dir, topdown=True,
                                     onerror=_walk_error, followlinks=False):
        dirs[:] = [d for d in dirs if not is_excluded_dir(d)]
        for f in files:
            if f.lower().endswith(".pdf") and not is_excluded_file(f):
                full = _strip_long_path(os.path.join(root, f))
                if full in skip_set:
                    continue
                yield full


def _new_stats() -> Dict[str, int]:
    return {
        "PROCESSED": 0,
        "UPGRADED":  0,
        "SKIPPED":   0,
        "ERROR":     0,
        "TOTAL":     0,
        "SKIP_NOT_PDF":        0,
        "SKIP_ENCRYPTED":      0,
        "SKIP_OVERSIZED":      0,
        "SKIP_OVERSIZED_IMAGE": 0,
        "SKIP_HAS_FORMS":      0,
        "SKIP_LOCKED":         0,
        "SKIP_MEMORY":         0,
        "SKIP_SIGNED":         0,
        "SKIP_MARKER_PDF":     0,
        "SKIP_MARKER_PDFA_2B": 0,
        "SKIP_MARKER_PDFA_2U": 0,
        "ERROR_OCR":            0,
        "ERROR_INVALID":        0,
        "ERROR_INVALID_STREAM": 0,
        "ERROR_MISSING":        0,
        "ERROR_DISK":           0,
        "ERROR_TEMP_COPY":      0,
        "ERROR_BACKUP":         0,
        "ERROR_MOVE":           0,
        "ERROR_TIMEOUT":        0,
    }


def _apply_result_to_stats(stats: Dict[str, int], result: _ProcessResult) -> None:
    stats[result.status] = stats.get(result.status, 0) + 1
    if result.detail and result.detail in stats:
        stats[result.detail] += 1
    if result.status in ("PROCESSED", "UPGRADED"):
        stats["MB_BEFORE"] = stats.get("MB_BEFORE", 0) + result.size_mb
        stats["MB_AFTER"]  = stats.get("MB_AFTER", 0)  + result.size_after_mb


# ==================================================================
# Pool-Hard-Kill, Log-Flush, Run-Summary
# ==================================================================

def _terminate_pool_workers(executor) -> int:
    # Hard-Kill aller Worker-Prozesse eines ProcessPoolExecutor
    procs = getattr(executor, "_processes", None) or {}
    killed = 0
    for pid, proc in list(procs.items()):
        try:
            if not proc.is_alive():
                continue
            proc.terminate()
            proc.join(timeout=5)
            if proc.is_alive():
                if hasattr(proc, "kill"):
                    proc.kill()
                else:
                    try:
                        os.kill(int(pid), signal.SIGTERM)
                    except Exception as _e:
                        detail_logger.debug(f"_terminate_pool_workers: Exception verworfen: {_e!r}")
                proc.join(timeout=2)
            killed += 1
        except Exception as _e:
            detail_logger.debug(f"_terminate_pool_workers: Exception verworfen: {_e!r}")
    return killed


def _prepend_iter(items, tail_iter):
    # Liefert zuerst alle Pfade aus 'items', danach den Rest aus 'tail_iter'.
    # Wird benutzt, um nach einem Pool-Reset unschuldige Dateien wieder einzureihen.
    for it in items:
        yield it
    for it in tail_iter:
        yield it


def _flush_log_handlers() -> None:
    for logger in (error_logger, detail_logger):
        for h in logger.handlers:
            try:
                h.flush()
            except Exception:
                pass


def write_run_summary(
    stats: Dict[str, int],
    target_dir: str,
    output_type: str,
    start_time: datetime,
    end_time: datetime,
    aborted: bool = False,
    abort_reason: Optional[str] = None,
    summary_path: str = RUN_SUMMARY_FILE,
    ist_probelauf: bool = False,
) -> None:
    duration = end_time - start_time
    total = stats.get("TOTAL", 0)
    try:
        with open(summary_path, "w", encoding="utf-8") as fh:
            fh.write("PDF-OCR-AUTOMATISIERUNG – LETZTER LAUF\n")
            fh.write("=" * 60 + "\n")
            fh.write(f"Verzeichnis:    {target_dir}\n")
            fh.write(f"Ausgabeformat:  {output_type}\n")
            fh.write(f"Start:          {start_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            fh.write(f"Ende:           {end_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            fh.write(f"Dauer:          {str(duration).split('.')[0]}\n")
            if aborted:
                fh.write(f"Status:         ABBRUCH ({abort_reason or 'unbekannt'})\n")
            else:
                fh.write("Status:         OK\n")
            fh.write("\n")
            fh.write(f"Verarbeitet:    {stats.get('PROCESSED', 0)}\n")
            fh.write(f"Upgrades:       {stats.get('UPGRADED', 0)}\n")
            fh.write(f"Übersprungen:   {stats.get('SKIPPED', 0)}\n")
            fh.write(f"Fehler:         {stats.get('ERROR', 0)}\n")
            fh.write(f"Gesamt:         {total}\n")
            mb_before = stats.get("MB_BEFORE", 0)
            mb_after  = stats.get("MB_AFTER", 0)
            if mb_before or mb_after:
                fh.write(f"Größe vorher:   {mb_before:.1f} MB\n")
                fh.write(f"Größe nachher:  {mb_after:.1f} MB\n")
                fh.write(f"Differenz:      {mb_before - mb_after:+.1f} MB\n")
            fh.write("\n")
            for k in sorted(stats.keys()):
                if k in ("PROCESSED", "UPGRADED", "SKIPPED", "ERROR", "TOTAL",
                         "MB_BEFORE", "MB_AFTER"):
                    continue
                v = stats.get(k, 0)
                if v > 0:
                    fh.write(f"  {k}: {v}\n")
    except Exception as e:
        log_warning("write_run_summary", f"Konnte Summary nicht schreiben: {_fmt_exc(e)}")


def dry_run_directory(
    directory:   str,
    config:      Dict,
    resume_path: Optional[str] = None,
) -> Dict[str, int]:
    """Probelauf: zeigt je Datei die Entscheidung, ohne etwas zu schreiben.

    Dieses Skript war das einzige schreibende der Sammlung ohne
    Probelauf - und es ersetzt PDFs an Ort und Stelle
    (safe_replace_with_retry). Ein Erstlauf auf fremdem Bestand war damit
    nicht verantwortbar zu machen. Aufbau nach dem Vorbild von
    dry_run_directory in 4a bis 4c.

    Geprueft wird alles, was ohne Schreibzugriff geht: Lesbarkeit,
    Sperre, vorhandene Textebene, OCR-Marker, Druckprofil, geschaetzter
    Speicher- und Plattenbedarf.
    """
    stats = _new_stats()

    print(f"\n📁 PROBELAUF (es wird nichts geschrieben): {directory}")
    if not safe_exists(directory):
        print(f"❌ Verzeichnis existiert nicht: {directory}")
        return stats

    skip_set: Set[str] = load_resume_set(resume_path) if resume_path else set()
    if skip_set:
        print(f"♻️  Resume-Datei: {len(skip_set)} Datei(en) gelten bereits als erledigt")

    print()
    print(f"  {'Entscheidung':<14} {'Seiten':>6} {'MB':>8}  Datei / Grund")
    print("  " + "-" * 78)

    wuerde_ocr = 0
    gesamt     = 0
    ram_max    = 0.0
    platte_mb  = 0.0

    for pdf_path in pdf_generator(directory, skip_set=skip_set):
        gesamt += 1
        name = os.path.basename(pdf_path)

        if is_file_locked(pdf_path):
            stats["SKIPPED"] += 1
            print(f"  {'ÜBERSPRINGEN':<14} {'-':>6} {'-':>8}  {name}  (gesperrt)")
            continue

        if not is_real_pdf(pdf_path):
            stats["ERROR"] += 1
            print(f"  {'FEHLER':<14} {'-':>6} {'-':>8}  {name}  (keine gültige PDF)")
            continue

        info = get_pdf_info(pdf_path)

        if info.get("error"):
            stats["ERROR"] += 1
            print(f"  {'FEHLER':<14} {'-':>6} {info['size_mb']:>8.1f}  {name}  ({info['error']})")
            continue

        seiten = info.get("pages", 0)
        groesse = info.get("size_mb", 0.0)

        if info.get("is_encrypted"):
            # Der echte Lauf ueberspringt NICHT pauschal: bei Owner-Only-
            # Verschluesselung (leeres Benutzerpasswort, bei Scanner-PDFs
            # haeufig) entfernt er den Schutz und ersetzt die Datei. Ob das
            # gelingt, laesst sich ohne Schreibzugriff nicht sicher sagen -
            # deshalb hier als 'moeglicherweise' ausweisen statt zu
            # behaupten, die Datei bleibe unangetastet.
            stats["SKIPPED"] += 1
            print(f"  {'ÜBERSPRINGEN':<14} {seiten:>6} {groesse:>8.1f}  {name}  "
                  f"(verschlüsselt – bei leerem Benutzerpasswort würde der echte "
                  f"Lauf entschlüsseln und ersetzen)")
            continue

        # Dieselbe Entscheidung wie im echten Lauf treffen. Frueher galt hier
        # 'has_text ODER irgendein Marker -> uebersprungen'. Der echte Lauf
        # sieht das anders:
        #   - has_text ohne Marker ist KEIN Skip, sondern nur ein Hinweis;
        #     ocrmypdf ueberspringt mit skip_text lediglich die OCR der
        #     Textseiten, schreibt aber sehr wohl eine Ausgabedatei, die
        #     anschliessend ueber das Original geschoben wird.
        #   - die Marker 'PDF', 'LEGACY' und 'PDFA_2B' fuehren bei der
        #     Standard-Zielvorgabe pdfa-2u zu einem Upgrade, also ebenfalls
        #     zu einer Ersetzung.
        # Ausgerechnet die Funktion, die den Erstlauf auf fremdem Bestand
        # verantwortbar machen soll, unterschaetzte die Zahl der ersetzten
        # Dateien damit erheblich - bei Bestaenden mit vielen digital
        # erzeugten PDFs um Groessenordnungen.
        output_type    = config.get("output_type", DEFAULT_OUTPUT_TYPE)
        target_is_pdfa = output_type.startswith("pdfa")
        marker         = info.get("existing_marker")

        skip_grund = None
        if marker == "PDFA_2U":
            skip_grund = "bereits PDF/A-2u"
        elif marker == "PDFA_2B" and output_type != "pdfa-2u":
            skip_grund = "bereits PDF/A-2b"
        elif marker in ("PDF", "LEGACY") and not target_is_pdfa:
            skip_grund = "bereits als Standard-PDF verarbeitet"

        if skip_grund:
            stats["SKIPPED"] += 1
            print(f"  {'ÜBERSPRINGEN':<14} {seiten:>6} {groesse:>8.1f}  {name}  ({skip_grund})")
            continue

        hinweise = []
        # Warum diese Datei trotz Marker/Textebene angefasst wird - genau die
        # Information, die der Bediener vor der Freigabe braucht.
        if marker == "PDFA_2B" and output_type == "pdfa-2u":
            hinweise.append("Upgrade PDF/A-2b → PDF/A-2u, wird ersetzt")
        elif marker in ("PDF", "LEGACY") and target_is_pdfa:
            fmt = "PDF/A-2u" if output_type == "pdfa-2u" else "PDF/A-2b"
            hinweise.append(f"Rekonvertierung nach {fmt}, wird ersetzt")
        elif info.get("has_text"):
            hinweise.append("Textebene vorhanden (OCR übersprungen), Datei wird ersetzt")
        try:
            druckprofil, warum = needs_print_safe_profile(info)
            if druckprofil:
                hinweise.append(f"bildschonend: {warum}")
        except Exception as _e:
            detail_logger.debug(f"dry_run_directory: Exception verworfen: {_e!r}")
        if info.get("needs_repair"):
            hinweise.append("Reparatur nötig")
        if info.get("has_forms"):
            hinweise.append("Formularfelder")

        try:
            ram = estimate_ocr_memory_mb(seiten or 1, 1)
            ram_max = max(ram_max, ram)
        except Exception as _e:
            detail_logger.debug(f"dry_run_directory: Exception verworfen: {_e!r}")
        try:
            platte_mb += required_disk_mb_for_pdf(groesse)
        except Exception as _e:
            detail_logger.debug(f"dry_run_directory: Exception verworfen: {_e!r}")

        wuerde_ocr += 1
        zusatz = ("  [" + ", ".join(hinweise) + "]") if hinweise else ""
        print(f"  {'WÜRDE OCR':<14} {seiten:>6} {groesse:>8.1f}  {name}{zusatz}")

    print("  " + "-" * 78)
    print(f"\n  Gesamt geprüft        : {gesamt}")
    print(f"  Würde verarbeitet     : {wuerde_ocr}")
    print(f"  Würde übersprungen    : {stats['SKIPPED']}")
    print(f"  Nicht lesbar / Fehler : {stats['ERROR']}")
    if ram_max:
        print(f"  Größter Speicherbedarf: {ram_max:,.0f} MB (eine Datei, ein Job)")
    if platte_mb:
        print(f"  Temporärer Plattenplatz: {platte_mb:,.0f} MB in Summe")
    print("\n  Es wurde nichts geschrieben.")

    stats["TOTAL"] = gesamt
    return stats


def process_directory(
    directory:    str,
    config:       Dict,
    temp_dir:     str,
    count_first:  bool = False,
    workers:      int  = 1,
    resume_path:  Optional[str] = None,
) -> Dict[str, int]:
    stats = _new_stats()

    print(f"\n📁 Verarbeite: {directory}")
    if workers > 1:
        print(f"⚙️  Parallelisierung: {workers} Worker")

    if not safe_exists(directory):
        print(f"❌ Verzeichnis existiert nicht: {directory}")
        return stats

    skip_set: Set[str] = load_resume_set(resume_path) if resume_path else set()
    if skip_set:
        print(f"♻️  Resume-Datei gefunden: {len(skip_set)} bereits verarbeitete Dateien werden übersprungen")

    gen = pdf_generator(directory, skip_set=skip_set)

    if count_first:
        print("🔍 Zähle PDF-Dateien ...")
        files_list = list(gen)
        total      = len(files_list)
        iterable   = iter(files_list)
        bar_fmt    = "{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]"
        print(f"📊 Gefunden: {total} PDF-Dateien (neu zu verarbeiten)\n")
    else:
        iterable = gen
        total    = None
        bar_fmt  = "{desc}: {n_fmt} Dateien [{elapsed}, {rate_fmt}]"
        print("Starte direkte Verarbeitung ...\n")

    print("-" * 66)

    disable_tqdm = not sys.stderr.isatty()

    def _handle_result(result: _ProcessResult, pbar) -> None:
        for line in result.messages:
            if disable_tqdm:
                print(line)
            else:
                pbar.write(line)
        _flush_log_entries(result.log_entries)
        _apply_result_to_stats(stats, result)
        _write_csv_row(result)
        # Transiente Skips (Datei gerade gesperrt / RAM gerade knapp) NICHT
        # ins Resume schreiben - sonst wuerden sie nach einem Neustart
        # dauerhaft uebersprungen, obwohl der Hinderungsgrund weg ist.
        if (resume_path
                and result.status in ("PROCESSED", "UPGRADED", "SKIPPED")
                and result.detail not in ("SKIP_LOCKED", "SKIP_MEMORY")):
            append_resume(resume_path, result.file_path)
        _flush_log_handlers()

    with tqdm(total=total, desc="OCR", unit="PDF",
              bar_format=bar_fmt, disable=disable_tqdm) as pbar:

        if workers <= 1:
            # Sequenzieller Modus
            for file_path in iterable:
                if shutdown_event.is_set():
                    pbar.write("  ⚠️  Graceful Shutdown – Abbruch")
                    break
                try:
                    result = process_pdf_file(file_path, config, temp_dir)
                    _handle_result(result, pbar)
                except Exception as e:
                    log_error("process_directory", f"Kritisch: {file_path}: {_fmt_exc(e)}",
                              exc_info=True)
                    stats["ERROR"] += 1
                    stats["ERROR_OCR"] = stats.get("ERROR_OCR", 0) + 1
                    pbar.write(f"  ✗ KRITISCH: {os.path.basename(file_path)}")
                finally:
                    pbar.update(1)

        else:
            # Paralleler Modus mit Hard-Kill, Pool-Reset und Heartbeat
            pool_kwargs = {
                "max_workers": workers,
                "initializer": _worker_init_tempdir,
                "initargs": (temp_dir,),
            }
            if sys.version_info >= (3, 11):
                pool_kwargs["max_tasks_per_child"] = 50

            heartbeat_stop = threading.Event()
            heartbeat_lock = threading.Lock()
            active_workers: Dict[str, float] = {}

            def _heartbeat_loop() -> None:
                last = time.time()
                while not heartbeat_stop.is_set():
                    time.sleep(5)
                    if time.time() - last < HEARTBEAT_INTERVAL:
                        continue
                    last = time.time()
                    with heartbeat_lock:
                        snapshot = list(active_workers.items())
                    now = time.time()
                    if snapshot:
                        msg_lines = [f"💓 Heartbeat: {len(snapshot)} Worker aktiv"]
                        for fp, started in sorted(snapshot, key=lambda x: x[1])[:5]:
                            mins = int((now - started) / 60)
                            msg_lines.append(f"    • {os.path.basename(fp)} (seit {mins} Min)")
                        for line in msg_lines:
                            if disable_tqdm:
                                print(line)
                            else:
                                pbar.write(line)
                        log_info("heartbeat", " | ".join(msg_lines))
                    else:
                        line = "💓 Heartbeat: keine aktiven Worker"
                        if disable_tqdm:
                            print(line)
                        else:
                            pbar.write(line)
                    _flush_log_handlers()

            heartbeat_thread = threading.Thread(
                target=_heartbeat_loop, daemon=True, name="ocr-heartbeat"
            )
            heartbeat_thread.start()

            iterable_active = iterable
            generation = 0
            requeue_counts: Dict[str, int] = {}
            try:
                while True:
                    generation += 1
                    if generation > MAX_POOL_GENERATIONS:
                        log_error(
                            "process_directory",
                            f"Generation-Limit ({MAX_POOL_GENERATIONS}) erreicht – "
                            f"verbleibende Dateien werden als ERROR markiert")
                        pbar.write(
                            f"  ⚠ Generation-Limit erreicht ({MAX_POOL_GENERATIONS}) – "
                            f"weitere Verarbeitung abgebrochen")
                        for fp_drain in iterable_active:
                            stats["ERROR"] += 1
                            stats["ERROR_GENERATION_LIMIT"] = stats.get(
                                "ERROR_GENERATION_LIMIT", 0) + 1
                            log_error(
                                "process_directory",
                                f"Abgebrochen (Generation-Limit): {fp_drain}")
                            _write_error_csv(fp_drain, "ERROR_GENERATION_LIMIT",
                                             "Generation-Limit erreicht – nicht verarbeitet")
                            pbar.update(1)
                        _flush_log_handlers()
                        break
                    pool_reset_needed = False
                    executor = ProcessPoolExecutor(**pool_kwargs)
                    future_to_path: Dict = {}

                    def _fill_queue() -> None:
                        # Queue-Tiefe = Worker-Anzahl: jeder submittete Task
                        # startet sofort, damit der Timeout (gemessen ab
                        # Submit) der tatsaechlichen Laufzeit entspricht.
                        # Bei groesserer Tiefe liefen wartende Tasks hinter
                        # einer langsamen Datei faelschlich in den Timeout
                        # und loesten unnoetige Pool-Resets aus.
                        while len(future_to_path) < workers:
                            if shutdown_event.is_set() or pool_reset_needed:
                                return
                            try:
                                fp = next(iterable_active)
                            except StopIteration:
                                break
                            future = executor.submit(
                                process_pdf_file, fp, config, temp_dir
                            )
                            future_to_path[future] = (fp, time.time())
                            with heartbeat_lock:
                                active_workers[fp] = time.time()

                    try:
                        _fill_queue()

                        while future_to_path:
                            done, _pending = concurrent.futures.wait(
                                future_to_path.keys(),
                                timeout=60,
                                return_when=concurrent.futures.FIRST_COMPLETED,
                            )

                            # Timeout-Erkennung: Hard-Kill aller Worker, Pool-Reset
                            now = time.time()
                            timed_out: List[Tuple[Any, str]] = []
                            for f in list(future_to_path.keys()):
                                fp, started = future_to_path[f]
                                if f not in done and (now - started) > WORKER_TIMEOUT:
                                    timed_out.append((f, fp))

                            if timed_out:
                                timed_out_futures = {f for f, _ in timed_out}

                                # ZUERST die bereits fertigen Futures auswerten.
                                # concurrent.futures.wait liefert in 'done' die
                                # abgeschlossenen Ergebnisse; die Auswertung
                                # dahinter wurde aber vom 'break' am Ende dieses
                                # Blocks uebersprungen. Diese Dateien waren auf
                                # der Platte bereits ersetzt - nur ohne CSV-Zeile,
                                # ohne Resume-Eintrag und ohne Statistik. Beim
                                # naechsten Lauf gaelten sie als unbearbeitet und
                                # wuerden erneut durch die OCR geschickt.
                                for f_done in list(done):
                                    if f_done not in future_to_path:
                                        continue
                                    fp_done, _st = future_to_path.pop(f_done)
                                    with heartbeat_lock:
                                        active_workers.pop(fp_done, None)
                                    try:
                                        _handle_result(f_done.result(), pbar)
                                    except Exception as e_done:
                                        log_error(
                                            "process_directory",
                                            f"Ergebnis nach Timeout-Reset nicht auswertbar "
                                            f"({fp_done}): {_fmt_exc(e_done)}")

                                for _, fp in timed_out:
                                    log_error("process_directory",
                                              f"Worker-Timeout ({WORKER_TIMEOUT//60} Min): {fp}")
                                    stats["ERROR"] += 1
                                    stats["ERROR_TIMEOUT"] = stats.get("ERROR_TIMEOUT", 0) + 1
                                    _write_error_csv(fp, "ERROR_TIMEOUT",
                                                     f"Worker-Timeout ({WORKER_TIMEOUT // 60} Min) – Hard-Kill")
                                    pbar.write(f"  ✗ TIMEOUT (Hard-Kill): {os.path.basename(fp)}")
                                    pbar.update(1)
                                pbar.write("  🔪 Terminiere Worker-Pool und starte ihn neu ...")
                                killed = _terminate_pool_workers(executor)
                                log_warning("process_directory",
                                            f"Pool-Reset nach Timeout – {killed} Worker hart beendet")
                                requeue_paths: List[str] = []
                                exhausted_paths: List[str] = []
                                for f, info in list(future_to_path.items()):
                                    fp_inflight, _ = info
                                    with heartbeat_lock:
                                        active_workers.pop(fp_inflight, None)
                                    if f in timed_out_futures:
                                        continue
                                    n = requeue_counts.get(fp_inflight, 0)
                                    if n < MAX_FILE_REQUEUES:
                                        requeue_counts[fp_inflight] = n + 1
                                        requeue_paths.append(fp_inflight)
                                    else:
                                        exhausted_paths.append(fp_inflight)
                                for efp in exhausted_paths:
                                    stats["ERROR"] += 1
                                    stats["ERROR_REQUEUE_LIMIT"] = stats.get(
                                        "ERROR_REQUEUE_LIMIT", 0) + 1
                                    log_error(
                                        "process_directory",
                                        f"Requeue-Limit ({MAX_FILE_REQUEUES}) erreicht: {efp}")
                                    _write_error_csv(efp, "ERROR_REQUEUE_LIMIT",
                                                     "Requeue-Limit nach Pool-Reset erreicht")
                                    pbar.write(
                                        f"  ✗ Requeue-Limit: {os.path.basename(efp)}")
                                    pbar.update(1)
                                if requeue_paths:
                                    log_warning("process_directory",
                                                f"Pool-Reset: {len(requeue_paths)} unschuldige Datei(en) "
                                                f"werden requeued")
                                    iterable_active = _prepend_iter(requeue_paths, iterable_active)
                                future_to_path.clear()
                                pool_reset_needed = True
                                _flush_log_handlers()
                                break

                            for future in done:
                                fp, _started = future_to_path.pop(future)
                                with heartbeat_lock:
                                    active_workers.pop(fp, None)
                                _suppress_pbar_update = False
                                try:
                                    result = future.result()
                                    _handle_result(result, pbar)
                                except BrokenProcessPool as e:
                                    log_error("process_directory",
                                              f"BrokenProcessPool: {fp}: {_fmt_exc(e)}", exc_info=True)
                                    # Ueberlebende Worker hart beenden - siehe
                                    # Begruendung im Watch-Modus. Ohne das
                                    # bleiben bei einem Teil-Crash Kindprozesse
                                    # stehen, die das Arbeitsverzeichnis halten.
                                    _terminate_pool_workers(executor)

                                    reporter_requeue: List[str] = []
                                    n_rep = requeue_counts.get(fp, 0)
                                    if n_rep < MAX_FILE_REQUEUES:
                                        requeue_counts[fp] = n_rep + 1
                                        reporter_requeue.append(fp)
                                        _suppress_pbar_update = True
                                        pbar.write(
                                            f"  ⚠ POOL-BRUCH (Reporter, Versuch "
                                            f"{n_rep + 2}/{MAX_FILE_REQUEUES + 1}): "
                                            f"{os.path.basename(fp)} – Pool wird neu aufgebaut")
                                    else:
                                        stats["ERROR"] += 1
                                        stats["ERROR_OCR"] = stats.get("ERROR_OCR", 0) + 1
                                        _write_error_csv(fp, "ERROR_OCR",
                                                         f"Pool-Bruch, Requeue-Limit erreicht: {_fmt_exc(e)}")
                                        pbar.write(
                                            f"  ✗ POOL-BRUCH (Reporter, Limit erreicht): "
                                            f"{os.path.basename(fp)}")

                                    requeue_paths_bp: List[str] = []
                                    exhausted_paths_bp: List[str] = []
                                    for f, info in list(future_to_path.items()):
                                        afp, _ = info
                                        with heartbeat_lock:
                                            active_workers.pop(afp, None)
                                        n_inn = requeue_counts.get(afp, 0)
                                        if n_inn < MAX_FILE_REQUEUES:
                                            requeue_counts[afp] = n_inn + 1
                                            requeue_paths_bp.append(afp)
                                        else:
                                            exhausted_paths_bp.append(afp)
                                    for efp in exhausted_paths_bp:
                                        stats["ERROR"] += 1
                                        stats["ERROR_REQUEUE_LIMIT"] = stats.get(
                                            "ERROR_REQUEUE_LIMIT", 0) + 1
                                        log_error(
                                            "process_directory",
                                            f"Requeue-Limit ({MAX_FILE_REQUEUES}) erreicht: {efp}")
                                        _write_error_csv(efp, "ERROR_REQUEUE_LIMIT",
                                                         "Requeue-Limit nach Pool-Bruch erreicht")
                                        pbar.write(
                                            f"  ✗ Requeue-Limit: {os.path.basename(efp)}")
                                        pbar.update(1)

                                    all_requeue = reporter_requeue + requeue_paths_bp
                                    if all_requeue:
                                        log_warning(
                                            "process_directory",
                                            f"Pool-Bruch: {len(all_requeue)} Datei(en) werden requeued "
                                            f"({len(reporter_requeue)} Reporter, "
                                            f"{len(requeue_paths_bp)} parallel)")
                                        iterable_active = _prepend_iter(all_requeue, iterable_active)
                                    future_to_path.clear()
                                    pool_reset_needed = True
                                    _flush_log_handlers()
                                    break
                                except Exception as e:
                                    log_error("process_directory",
                                              f"Worker-Fehler: {fp}: {_fmt_exc(e)}", exc_info=True)
                                    stats["ERROR"] += 1
                                    stats["ERROR_OCR"] = stats.get("ERROR_OCR", 0) + 1
                                    pbar.write(f"  ✗ KRITISCH: {os.path.basename(fp)}")
                                finally:
                                    if not _suppress_pbar_update:
                                        pbar.update(1)

                            if pool_reset_needed:
                                break

                            if shutdown_event.is_set():
                                for f in list(future_to_path.keys()):
                                    f.cancel()
                                break

                            _fill_queue()
                    finally:
                        if pool_reset_needed:
                            executor.shutdown(wait=False)
                            # Pool ist tot, alle Tasks requeued - jetzt sind
                            # die UUID-Temp-Dateien der abgestuerzten Worker
                            # verwaist und koennen gefahrlos geraeumt werden.
                            n_purged = _purge_orphan_temp_files(temp_dir)
                            if n_purged:
                                log_warning("process_directory",
                                            f"Pool-Reset: {n_purged} verwaiste Temp-Datei(en) entfernt")
                        else:
                            executor.shutdown(wait=True)

                    if shutdown_event.is_set():
                        break
                    if not pool_reset_needed:
                        # Iterable erschöpft, Pool sauber durchgelaufen
                        break
                    # Sonst: nächste Generation – frischer Pool, gleicher Generator
            finally:
                heartbeat_stop.set()
                heartbeat_thread.join(timeout=2)

    stats["TOTAL"] = (
        total if (count_first and total is not None)
        else sum(stats[k] for k in ("PROCESSED", "UPGRADED", "SKIPPED", "ERROR"))
    )

    _flush_log_handlers()
    return stats


# ==================================================================
# Hotfolder / Watch-Modus (watchdog)
# ==================================================================

class _PdfEventHandler(FileSystemEventHandler):

    def __init__(self, work_queue: queue.Queue, watch_dir: str, temp_dir: str,
                 stable_delay: float = 2.0) -> None:
        super().__init__()
        self._queue        = work_queue
        self._watch_dir    = os.path.abspath(watch_dir)
        self._temp_dir     = os.path.abspath(temp_dir)
        self._stable_delay = stable_delay
        self._dedup: Dict[str, float] = {}
        self._dedup_lock = threading.Lock()

    def _should_enqueue(self, abs_path: str) -> bool:
        now = time.time()
        with self._dedup_lock:
            for k in list(self._dedup.keys()):
                if now - self._dedup[k] > ENQUEUE_DEDUP_WINDOW:
                    del self._dedup[k]
            last = self._dedup.get(abs_path)
            if last is not None and (now - last) < ENQUEUE_DEDUP_WINDOW:
                self._dedup[abs_path] = now
                return False
            self._dedup[abs_path] = now
        return True

    def _enqueue(self, path: str) -> None:
        path = _strip_long_path(path)
        abs_path = os.path.abspath(path)

        if not abs_path.lower().endswith(".pdf"):
            return

        abs_path_nc = os.path.normcase(abs_path)

        watch_dir_strict = os.path.normcase(os.path.join(self._watch_dir, ""))
        if not abs_path_nc.startswith(watch_dir_strict):
            return

        temp_dir_strict = os.path.normcase(os.path.join(self._temp_dir, ""))
        if abs_path_nc.startswith(temp_dir_strict):
            return

        if is_excluded_file(os.path.basename(abs_path)):
            return

        try:
            rel_path = os.path.relpath(abs_path, self._watch_dir)
            parts = Path(rel_path).parts[:-1]
            if any(is_excluded_dir(p) for p in parts):
                return
        except ValueError:
            pass

        if not self._should_enqueue(abs_path):
            return

        self._queue.put((abs_path, time.time() + self._stable_delay, 0))

    def on_created(self, event) -> None:
        if event.is_directory:
            safe_dir = prepare_long_path(event.src_path)
            for root, dirs, files in os.walk(safe_dir, followlinks=False):
                dirs[:] = [d for d in dirs if not is_excluded_dir(d)]
                for f in files:
                    self._enqueue(os.path.join(root, f))
        else:
            self._enqueue(event.src_path)

    def on_moved(self, event) -> None:
        if event.is_directory:
            safe_dir = prepare_long_path(event.dest_path)
            for root, dirs, files in os.walk(safe_dir, followlinks=False):
                dirs[:] = [d for d in dirs if not is_excluded_dir(d)]
                for f in files:
                    self._enqueue(os.path.join(root, f))
        else:
            self._enqueue(event.dest_path)

    def on_modified(self, event) -> None:
        if event.is_directory:
            return
        self._enqueue(event.src_path)


def _is_network_path(directory: str) -> bool:
    if directory.startswith("\\\\") or directory.startswith("//"):
        return True
    if len(directory) >= 2 and directory[1] == ":":
        try:
            drive_root = directory[:2] + "\\"
            drive_type = ctypes.windll.kernel32.GetDriveTypeW(drive_root)
            if drive_type == 4:
                return True
        except Exception as _e:
            detail_logger.debug(f"_is_network_path: Exception verworfen: {_e!r}")
    return False


def run_watch_mode(
    directory:    str,
    config:       Dict,
    temp_dir:     str,
    workers:      int   = 1,
    stable_delay: float = 3.0,
    initial_scan: bool  = True,
) -> None:
    if not HAS_WATCHDOG:
        print("\n❌ Watchdog-Bibliothek nicht installiert.")
        print("   pip install watchdog")
        return

    watch_start     = datetime.now()
    is_network_path = _is_network_path(directory)

    print("\n👁️  HOTFOLDER-MODUS")
    print(f"  Überwachtes Verzeichnis: {directory}")
    print(f"  Stabilisierungs-Delay:   {stable_delay}s")
    print(f"  Worker:                  {workers}")
    print(f"  Heartbeat:               alle {HEARTBEAT_INTERVAL}s")
    if is_network_path:
        print("  Netzwerkpfad erkannt    → PollingObserver (zuverlässiger auf SMB/DFS)")
    print("  Beenden mit Ctrl+C\n")
    print("-" * 66)

    work_queue:  queue.Queue     = queue.Queue()
    stop_event:  threading.Event = threading.Event()

    event_handler = _PdfEventHandler(work_queue, watch_dir=directory,
                                     temp_dir=temp_dir, stable_delay=stable_delay)
    if is_network_path:
        observer = PollingObserver(timeout=5)
    else:
        observer = Observer()
    observer.schedule(event_handler, directory, recursive=True)
    observer.start()
    print("  ✓ Ordner-Überwachung gestartet\n")

    stats = _new_stats()

    echo_cache_seconds = 120.0 if is_network_path else 10.0
    if is_network_path:
        print(f"  ℹ️  Netzlaufwerk – Echo-Cache: {echo_cache_seconds:.0f}s")

    thread_lock        = threading.Lock()
    active_files: Set[str]           = set()
    active_started: Dict[str, float] = {}
    recently_processed: Dict[str, float] = {}

    def _worker_loop(thread_temp_dir: str) -> None:
        # Eigener Pool pro Thread – Hard-Kill betrifft nur diesen Thread.
        # EIGENES Temp-Unterverzeichnis pro Thread: _purge_orphan_temp_files
        # nach Timeout/Pool-Bruch darf nur die eigenen verwaisten Dateien
        # treffen – im gemeinsamen temp_dir wuerde es die AKTIVEN Temp-
        # Dateien der anderen Worker-Threads loeschen.
        try:
            os.makedirs(thread_temp_dir, exist_ok=True)
        except Exception:
            thread_temp_dir = temp_dir
        pool_kwargs_single = {
            "max_workers": 1,
            "initializer": _worker_init_tempdir,
            "initargs": (thread_temp_dir,),
        }
        if sys.version_info >= (3, 11):
            pool_kwargs_single["max_tasks_per_child"] = 50
        executor = ProcessPoolExecutor(**pool_kwargs_single)
        try:
            while not stop_event.is_set():
                try:
                    file_path, ready_at, retries = work_queue.get(timeout=0.5)
                except queue.Empty:
                    continue

                wait_secs = ready_at - time.time()
                if wait_secs > 0:
                    work_queue.put((file_path, ready_at, retries))
                    work_queue.task_done()
                    time.sleep(min(0.5, wait_secs))
                    continue

                if not safe_exists(file_path):
                    work_queue.task_done()
                    continue

                with thread_lock:
                    current_time = time.time()
                    for k in list(recently_processed.keys()):
                        if current_time - recently_processed[k] > echo_cache_seconds:
                            del recently_processed[k]

                    if file_path in active_files or file_path in recently_processed:
                        work_queue.task_done()
                        continue

                    if is_file_locked(file_path):
                        if retries < MAX_LOCK_RETRIES:
                            work_queue.put((file_path, time.time() + 5.0, retries + 1))
                        else:
                            stats["ERROR"] += 1
                            stats["ERROR_OCR"] = stats.get("ERROR_OCR", 0) + 1
                            print(f"  [ERROR] {os.path.basename(file_path)}: Dauerhaft gesperrt")
                            log_error("watch_mode", f"Dauerhaft gesperrt: {file_path}")
                        work_queue.task_done()
                        continue

                    active_files.add(file_path)
                    active_started[file_path] = time.time()

                print(f"\n🆕 Neue PDF erkannt: {os.path.basename(file_path)}")

                # Fehlergrund mitfuehren. Der aeussere except-Zweig unten
                # umfasst nicht nur den Timeout-Pfad, sondern auch den
                # Pool-Bruch (der bewusst ein generisches Exception-Objekt
                # wirft) und das Einsammeln/Auswerten des Ergebnisses. Alles
                # landete pauschal als ERROR_TIMEOUT in Statistik und CSV -
                # eine Fehlersuche lief damit in die falsche Richtung.
                fehlergrund = "ERROR_OCR"
                try:
                    future = executor.submit(process_pdf_file, file_path, config, thread_temp_dir)
                    try:
                        result = future.result(timeout=WORKER_TIMEOUT)
                    except concurrent.futures.TimeoutError:
                        fehlergrund = "ERROR_TIMEOUT"
                        future.cancel()
                        _terminate_pool_workers(executor)
                        executor.shutdown(wait=False)
                        executor = ProcessPoolExecutor(**pool_kwargs_single)
                        if thread_temp_dir != temp_dir:
                            _purge_orphan_temp_files(thread_temp_dir)
                        raise Exception(
                            f"Worker-Timeout ({WORKER_TIMEOUT // 60} Min) – "
                            "Thread-Pool hart terminiert und neu gestartet"
                        )
                    except BrokenProcessPool as e:
                        # Ueberlebende Worker hart beenden, BEVOR der Pool
                        # verworfen wird. shutdown(wait=False) allein laesst
                        # bei einem Teil-Crash die noch laufenden Kindprozesse
                        # stehen - sie halten dann das alte Arbeitsverzeichnis
                        # samt unkomprimierter Seitenbilder offen, und
                        # _purge_orphan_temp_files darunter kann es nicht
                        # raeumen. Der Timeout-Zweig macht es bereits so.
                        _terminate_pool_workers(executor)
                        executor.shutdown(wait=False)
                        executor = ProcessPoolExecutor(**pool_kwargs_single)
                        if thread_temp_dir != temp_dir:
                            _purge_orphan_temp_files(thread_temp_dir)
                        raise Exception(f"Worker-Prozess gestorben – Pool neu gestartet: {_fmt_exc(e)}")

                    with thread_lock:
                        for line in result.messages:
                            print(line)
                        print(f"  [{result.status}] {os.path.basename(file_path)}")
                        _apply_result_to_stats(stats, result)

                    _flush_log_entries(result.log_entries)
                    _write_csv_row(result)
                    _flush_log_handlers()

                except Exception as e:
                    with thread_lock:
                        stats["ERROR"] += 1
                        stats[fehlergrund] = stats.get(fehlergrund, 0) + 1
                        print(f"  [ERROR] {os.path.basename(file_path)}: {_fmt_exc(e)}")
                    log_error("watch_mode", f"Fehler bei {file_path}: {_fmt_exc(e)}", exc_info=True)
                    _write_error_csv(file_path, fehlergrund, _fmt_exc(e))

                finally:
                    with thread_lock:
                        active_files.discard(file_path)
                        active_started.pop(file_path, None)
                        recently_processed[file_path] = time.time()
                    work_queue.task_done()
        finally:
            try:
                _terminate_pool_workers(executor)
            except Exception as _e:
                detail_logger.debug(f"_worker_loop: Exception verworfen: {_e!r}")
            executor.shutdown(wait=False)

    def _heartbeat_loop() -> None:
        last_beat = time.time()
        while not stop_event.is_set():
            time.sleep(5)
            if time.time() - last_beat < HEARTBEAT_INTERVAL:
                continue
            last_beat = time.time()
            with thread_lock:
                active_snapshot  = list(active_files)
                started_snapshot = dict(active_started)
            if active_snapshot:
                print(f"\n💓 Heartbeat: {len(active_snapshot)} Datei(en) aktiv")
                now = time.time()
                for fp in active_snapshot[:5]:
                    mins = int((now - started_snapshot.get(fp, now)) / 60)
                    print(f"    • {os.path.basename(fp)} (seit {mins} Min)")
            else:
                print("\n💓 Heartbeat: keine aktiven Verarbeitungen – warte auf neue PDFs")
            _flush_log_handlers()

    worker_threads = [
        threading.Thread(target=_worker_loop, daemon=True, name=f"ocr-worker-{i}",
                         args=(os.path.join(temp_dir, f"watch_thread_{i}"),))
        for i in range(workers)
    ]
    for t in worker_threads:
        t.start()

    heartbeat_thread = threading.Thread(
        target=_heartbeat_loop, daemon=True, name="ocr-heartbeat"
    )
    heartbeat_thread.start()

    if workers > 1:
        print(f"  ✓ {workers} Worker-Threads (je eigener Sandbox-Pool und Temp-Unterordner) gestartet\n")
    else:
        print("  ✓ 1 Worker-Thread mit Sandbox-Pool gestartet\n")

    if initial_scan:
        # Watchdog liefert nur NEUE Ereignisse – PDFs, die beim Start
        # bereits im Ordner liegen (z.B. ueber Nacht aufgelaufen, waehrend
        # der Daemon nicht lief), wuerden sonst nie verarbeitet. Die
        # Worker laufen bereits und arbeiten parallel zum Scan.
        print("  🔍 Initial-Scan: reihe bereits vorhandene PDFs ein ...")
        n_init = 0
        for fp in pdf_generator(directory):
            if shutdown_event.is_set():
                break
            event_handler._enqueue(fp)
            n_init += 1
        print(f"  ✓ Initial-Scan abgeschlossen: {n_init} vorhandene PDF(s) eingereiht\n")

    try:
        while True:
            time.sleep(1)
            if shutdown_event.is_set():
                print("\n⚠️  Graceful Shutdown – beende Watch-Modus")
                break
            if not observer.is_alive():
                print("\n⚠️  Watchdog-Observer unerwartet beendet (Netzwerkabbruch?)")
                break
    except KeyboardInterrupt:
        print("\n\n⚠️  Hotfolder-Modus beendet (Ctrl+C)")
    finally:
        stop_event.set()
        observer.stop()
        observer.join(timeout=5)
        for t in worker_threads:
            t.join(timeout=5)
        heartbeat_thread.join(timeout=2)

    stats["TOTAL"] = sum(stats[k] for k in ("PROCESSED", "UPGRADED", "SKIPPED", "ERROR"))

    print("\n  Statistik:")
    print(f"    [✓] Verarbeitet: {stats['PROCESSED']}")
    print(f"    [↑] Upgrades:    {stats['UPGRADED']}")
    print(f"    [=] Übersprungen: {stats['SKIPPED']}")
    print(f"    [✗] Fehler:       {stats['ERROR']}")

    write_run_summary(
        stats, directory, str(config.get("output_type", "?")),
        watch_start, datetime.now(),
        aborted=False, abort_reason=None,
    )


# ==================================================================
# Statistik-Ausgabe
# ==================================================================

def print_detailed_stats(stats: Dict[str, int]) -> None:
    total = stats.get("TOTAL", 0)
    print("  STATISTIK:")
    print(f"    [✓] Neu verarbeitet (OCR):  {stats.get('PROCESSED', 0)}")
    print(f"    [↑] Upgrades nach PDF/A:    {stats.get('UPGRADED', 0)}")
    print(f"    [=] Übersprungen:            {stats.get('SKIPPED', 0)}")

    skip_rows = [
        ("Bereits als PDF/A-2u markiert", stats.get("SKIP_MARKER_PDFA_2U", 0)),
        ("Bereits als PDF/A-2b markiert", stats.get("SKIP_MARKER_PDFA_2B", 0)),
        ("Bereits als PDF markiert",      stats.get("SKIP_MARKER_PDF", 0)),
        ("Enthält Formularfelder",        stats.get("SKIP_HAS_FORMS", 0)),
        ("Verschlüsselt",                  stats.get("SKIP_ENCRYPTED", 0)),
        ("Digital signiert",               stats.get("SKIP_SIGNED", 0)),
        ("Physisch übergroß",              stats.get("SKIP_OVERSIZED", 0)),
        ("Eingebettetes Bild übergroß",    stats.get("SKIP_OVERSIZED_IMAGE", 0)),
        ("Dauerhaft gesperrt",             stats.get("SKIP_LOCKED", 0)),
        ("Kein gültiges PDF",              stats.get("SKIP_NOT_PDF", 0)),
        ("RAM-Schutz",                     stats.get("SKIP_MEMORY", 0)),
    ]
    for label, count in skip_rows:
        if count > 0:
            print(f"         • {label}: {count}")

    print(f"    [✗] Fehler:                  {stats.get('ERROR', 0)}")
    err_rows = [
        ("OCR-Fehler",          stats.get("ERROR_OCR", 0)),
        ("Ungültige PDF",       stats.get("ERROR_INVALID", 0)),
        ("Korrupter Stream",    stats.get("ERROR_INVALID_STREAM", 0)),
        ("Datei fehlt",         stats.get("ERROR_MISSING", 0)),
        ("Speichermangel",      stats.get("ERROR_DISK", 0)),
        ("Temp-Kopie-Fehler",   stats.get("ERROR_TEMP_COPY", 0)),
        ("Backup-Fehler",       stats.get("ERROR_BACKUP", 0)),
        ("Verschiebe-Fehler",   stats.get("ERROR_MOVE", 0)),
        ("Worker-Timeout",      stats.get("ERROR_TIMEOUT", 0)),
        ("Requeue-Limit",       stats.get("ERROR_REQUEUE_LIMIT", 0)),
        ("Generation-Limit",    stats.get("ERROR_GENERATION_LIMIT", 0)),
    ]
    for label, count in err_rows:
        if count > 0:
            print(f"         • {label}: {count}")

    mb_before = stats.get("MB_BEFORE", 0)
    mb_after  = stats.get("MB_AFTER", 0)
    if mb_before or mb_after:
        diff = mb_before - mb_after
        trend = "eingespart" if diff >= 0 else "größer geworden"
        print(f"    [≡] Verarbeitete Dateien:    {mb_before:.1f} MB → {mb_after:.1f} MB "
              f"({abs(diff):.1f} MB {trend})")

    print(f"        GESAMT:                  {total}")


# ==================================================================
# Hauptprogramm
# ==================================================================

def main() -> None:

    _install_signal_handlers()
    _setup_csv_log()

    # Muss VOR _register_dll_dirs und vor der Werkzeugsuche laufen: danach
    # ist C:\OCR bereits im DLL-Suchpfad bzw. eine dort abgelegte .exe
    # schon gestartet.
    _pruefe_werkzeugverzeichnisse()

    # DLL-Suchpfade früh registrieren (relevant für EXE ohne Adminrechte)
    exe_dir = _get_exe_dir()
    meipass = _get_meipass_dir()
    early_dll_dirs = [exe_dir] + ([meipass] if meipass else []) + EXTRA_DLL_DIRS
    _register_dll_dirs(early_dll_dirs)

    parser = argparse.ArgumentParser(
        description="PDF-OCR-Automatisierung",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Beispiele:\n"
            "  python 5_OCR_PDF.py Q:\\\n"
            "  python 5_OCR_PDF.py Q:\\ --workers 3\n"
            "  python 5_OCR_PDF.py Q:\\ --watch\n"
            "  python 5_OCR_PDF.py Q:\\ --daemon\n"
            "  python 5_OCR_PDF.py Q:\\ --auto --workers 2\n"
            "  python 5_OCR_PDF.py Q:\\ --pdfa\n"
            "  python 5_OCR_PDF.py Q:\\ --no-resume\n"
            "  python 5_OCR_PDF.py Q:\\ --verify-markers\n"
            "  python 5_OCR_PDF.py Q:\\ --cleanup-backups\n"
        ),
    )
    parser.add_argument("directory", nargs="?", default=None,
                        help="Verzeichnis (optional; sonst interaktive Auswahl)")
    parser.add_argument("--lang", "-l", default=None,
                        help="OCR-Sprache, z.B. 'deu+eng'")
    parser.add_argument("--no-count", action="store_true",
                        help="Nicht vorher zählen")
    parser.add_argument("--workers", "-w", type=int, default=None,
                        help="Anzahl Worker-Prozesse")
    parser.add_argument("--auto", action="store_true",
                        help="Nicht-interaktiv (keine Rückfragen)")
    parser.add_argument("--watch", "--daemon", action="store_true",
                        dest="watch",
                        help="Hotfolder-Modus (watchdog)")
    parser.add_argument("--pdfa", action="store_true",
                        help="Ausgabe als PDF/A-2u")
    parser.add_argument("--output-type", choices=[
        "pdf", "pdfa", "pdfa-1", "pdfa-2", "pdfa-2b", "pdfa-2u", "pdfa-3",
    ], default=None, help="Ausgabeformat")
    parser.add_argument("--no-resume", action="store_true",
                        help="Resume-Datei ignorieren (von vorne beginnen)")
    parser.add_argument("--verify-markers", action="store_true",
                        help="Vorhandene PDF/A-Marker gegen veraPDF prüfen und ggf. korrigieren")
    parser.add_argument("--cleanup-backups", action="store_true",
                        help="Backup-Leichen (*.backup) beim Start rekursiv aufräumen")
    parser.add_argument("--clear-mru", action="store_true",
                        help="Recent-Documents-Liste leeren (Default: aus)")
    parser.add_argument("--no-initial-scan", action="store_true",
                        help="Watch-Modus: vorhandene PDFs beim Start nicht einreihen")
    parser.add_argument("--print-safe", action="store_true",
                        help="Bildschonendes Profil für ALLE Dateien erzwingen "
                             "(kein Neurastern, keine Bild-Neukodierung)")
    parser.add_argument("--no-print-safe", action="store_true",
                        help="Automatische Erkennung druckrelevanter Bilder abschalten "
                             "(altes Verhalten – nicht empfohlen)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Probelauf: zeigt je Datei die Entscheidung, "
                             "schreibt nichts")
    args = parser.parse_args()

    if args.clear_mru:
        clear_mru()

    # --- Einzelinstanz-Schutz ---
    # Zwei parallele Laeufe ueber denselben Bestand wuerden sich die
    # Temp-Kopien und Ersetzungen gegenseitig wegziehen. Im Probelauf
    # nicht noetig - der schreibt nichts.
    _sperre = None
    if gem is not None and not args.dry_run:
        _sperre = gem.Einzelinstanz("5_OCR_PDF")
        if not _sperre.belegen():
            print(_sperre.hinweis())
            sys.exit(1)
        import atexit
        atexit.register(_sperre.freigeben)

    # Restore-Privilegien (Admin-Kontext) für ACL/Owner-Erhalt aktivieren.
    _enable_restore_privileges()

    print()
    print("╔══════════════════════════════════════════════════════════════════════╗")
    print("║          PDF-OCR-AUTOMATISIERUNG: MASCHINELLE TEXTERKENNUNG          ║")
    print("╚══════════════════════════════════════════════════════════════════════╝")
    print(f"  Start:   {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Python:  {sys.version.split()[0]}")
    print(f"  System:  {platform.system()} {platform.version()[:40]}")
    print(f"  Log-Dir: {LOG_DIR}")
    print(f"  EXE-Dir: {exe_dir}")

    cpu_count, cpu_name = get_cpu_info()
    print()
    print_cpu_recommendation(cpu_count, cpu_name)

    # --- Zielverzeichnis ---
    if args.directory:
        target_dir = sanitize_path(args.directory)
        print(f"\n📂 Verzeichnis per Argument: {target_dir}")
    elif args.auto:
        print("\n⚠️  --auto ohne Verzeichnis – bitte Verzeichnis angeben.")
        sys.exit(1)
    else:
        target_dir = ask_directory()

    if not safe_exists(target_dir):
        print(f"\n❌ VERZEICHNIS NICHT GEFUNDEN: {target_dir}")
        if re.match(r"^[A-Za-z]:", target_dir):
            print("   Hinweis: Bei erhöhten Rechten ('als Administrator') sind Netz-")
            print("   laufwerksbuchstaben oft nicht gemappt – bitte UNC-Pfad versuchen")
            print("   (z.B. \\\\server\\freigabe).")
        if not args.auto:
            try:
                input("Drücke Enter zum Beenden ...")
            except EOFError:
                pass
        return

    # --- Backup-Leichen aufräumen (nur auf ausdrücklichen Wunsch) ---
    do_cleanup = False
    if args.cleanup_backups:
        do_cleanup = True
    elif not args.auto and not args.watch:
        print()
        print("—" * 70)
        print("  Backup-Leichen aufräumen?")
        print("  Sucht rekursiv nach *.backup-Dateien aus abgebrochenen Läufen.")
        print("  Vorteil: Saubere Ablage, keine Altlasten nach Abstürzen.")
        print("  Nachteil: Bei großen Netzlaufwerken mit tiefen Strukturen")
        print("            kann die Suche vor Beginn der Arbeit mehrere Minuten")
        print("            bis Stunden dauern. Im Normalbetrieb werden Backups")
        print("            ohnehin nach jeder erfolgreichen Datei einzeln gelöscht.")
        print("—" * 70)
        do_cleanup = ask_yes_no("  Jetzt nach Backup-Leichen suchen?", default_yes=False)

    # Im Probelauf wird auch hier NICHTS geloescht. Das Aufraeumen entfernt
    # zwar nur eigene Artefakte, aber ein Probelauf, der Dateien loescht,
    # ist keiner - und genau darauf verlaesst sich der Erstlauf auf fremdem
    # Bestand.
    if do_cleanup and args.dry_run:
        print("\n🧹 Probelauf: Suche nach Resten wird übersprungen "
              "(im Probelauf wird nichts gelöscht).")
    elif do_cleanup:
        print(f"\n🧹 Suche nach Resten (*.backup, *.tmp_new, *.tmp_marker, älter als "
              f"{BACKUP_CLEANUP_MIN_AGE_HOURS}h) ...")
        removed = cleanup_orphaned_backups(target_dir)
        if removed > 0:
            print(f"  ✓ {removed} Rest-Datei(en) entfernt")
        else:
            print("  ✓ Keine Reste gefunden")

    # --- Temporäres Verzeichnis ---
    temp_dir = TEMP_DIR
    print(f"\n🔧 Temp-Verzeichnis: {temp_dir}")
    try:
        os.makedirs(temp_dir, exist_ok=True)
        print("  ✓ Bereit")
    except Exception as e:
        print(f"  ⚠️  Konnte nicht erstellt werden: {_fmt_exc(e)}")
        # Fallback als EIGENER Unterordner im System-Temp - niemals der
        # System-Temp selbst, sonst wuerden rmtree/_purge_orphan_temp_files
        # fremde Daten treffen.
        temp_dir = os.path.join(
            _ORIG_WINDOWS_TEMP, f"5_OCR_PDF_Processing_{os.getpid()}")
        try:
            os.makedirs(temp_dir, exist_ok=True)
        except Exception:
            temp_dir = _ORIG_WINDOWS_TEMP
        print(f"  → Fallback: {temp_dir}")

    # Temp-Umlenkung auch fuer den MAIN-Prozess: im sequenziellen Modus
    # (workers=1) laeuft ocrmypdf direkt im Hauptprozess - ohne Umlenkung
    # landeten dessen ocrmypdf.io.*-Zwischenordner im System-Temp und
    # blieben bei Abstuerzen als GB-grosser Muell liegen. Der echte
    # System-Temp-Pfad ist vorab in _ORIG_WINDOWS_TEMP gesichert; auch
    # der ocrmypdf-Systemtest nutzt ab hier das kontrollierte Verzeichnis.
    if temp_dir != _ORIG_WINDOWS_TEMP:
        _worker_init_tempdir(temp_dir)

    # --- Systemprüfung ---
    deps          = check_external_dependencies()
    valid, config = validate_dependencies(deps)

    if not valid:
        print("\n❌ Kritische Abhängigkeiten fehlen. Skript wird beendet.")
        if not args.auto:
            try:
                input("Drücke Enter ...")
            except EOFError:
                pass
        sys.exit(1)

    # --- Ausgabeformat ---
    if args.output_type:
        output_type = args.output_type
    elif args.pdfa:
        output_type = "pdfa-2u"
    elif args.auto:
        output_type = DEFAULT_OUTPUT_TYPE
    else:
        output_type = None

    # --- Marker-Verifikation ---
    config["verify_markers"] = bool(args.verify_markers)
    if args.verify_markers and not config.get("has_verapdf"):
        print("\n⚠️  --verify-markers angefordert, aber veraPDF ist nicht verfügbar.")
        print("    Marker-Re-Verifikation wird übersprungen.")

    # --- Interaktive Konfiguration ---
    if args.auto:
        count_first = not args.no_count
        if args.workers is not None:
            num_workers = min(61, max(1, args.workers))
        else:
            num_workers = recommended_worker_count(cpu_count)
        if args.lang:
            config["ocr_language"] = args.lang
        if output_type is None:
            output_type = DEFAULT_OUTPUT_TYPE
    else:
        print("\n" + "=" * 70)
        print("  EINSTELLUNGEN")
        print("=" * 70)

        if args.lang:
            config["ocr_language"] = args.lang
            print(f"  Sprache (Argument): {args.lang}")
        else:
            default_lang = config.get("ocr_language", DEFAULT_LANGUAGE)
            try:
                lang_input = input(f"  OCR-Sprache [Standard: {default_lang}]: ").strip()
            except EOFError:
                lang_input = ""
            if lang_input:
                config["ocr_language"] = lang_input

        if output_type is None:
            output_type = ask_output_type()

        if args.no_count:
            count_first = False
        else:
            count_first = ask_yes_no(
                "\n  Gesamtzahl der Dateien vorab ermitteln?\n"
                "    (Ja = länger bis zum Start, dafür exakte Fortschritts-Anzeige)\n"
                "    (Hinweis: Bei UNC-Pfaden oder sehr großen Freigaben kann die\n"
                "     Indizierung 10–30 Minuten dauern, bevor die erste Datei\n"
                "     verarbeitet wird. Bei > 100.000 Dateien besser mit 'Nein'\n"
                "     antworten und die Generator-Variante verwenden.)",
            )

        if args.workers is not None:
            num_workers = min(61, max(1, args.workers))
            print(f"  Worker (Argument): {num_workers}")
        else:
            num_workers = ask_worker_count(cpu_count)

        if not args.verify_markers and config.get("has_verapdf") \
                and output_type.startswith("pdfa"):
            if ask_yes_no(
                "\n  Vorhandene PDF/A-Marker via veraPDF re-verifizieren?\n"
                "  (Ja = langsamer, erkennt aber fehlerhafte Alt-Marker)",
                default_yes=False,
            ):
                config["verify_markers"] = True

    config["output_type"]      = output_type
    config["force_print_safe"] = bool(args.print_safe)
    config["no_print_safe"]    = bool(args.no_print_safe)
    if args.print_safe and args.no_print_safe:
        print("  ⚠️  --print-safe und --no-print-safe gleichzeitig gesetzt – "
              "--print-safe hat Vorrang.")
        config["no_print_safe"] = False

    # --- Hinweis bei pdfa-2u ohne veraPDF ---
    if output_type == "pdfa-2u" and not config.get("has_verapdf"):
        print()
        print("⚠️  HINWEIS: PDF/A-2u angefordert, aber veraPDF nicht verfügbar.")
        print("   Es wird konservativ höchstens PDF/A-2b markiert (kein 2u-Marker")
        print("   ohne validierte Konformität).")

    # --- Resume-Datei ---
    resume_path: Optional[str] = None
    if not args.watch and not args.no_resume:
        resume_path = get_resume_file_path(target_dir, output_type)

    # --- Zusammenfassung ---
    is_pdfa = output_type.startswith("pdfa")
    format_labels = {
        "pdfa-2u": "PDF/A-2u (Unicode – Bundesarchiv-Empfehlung)",
        "pdfa-2b": "PDF/A-2b (Basic – Visuelle Integrität)",
        "pdfa-2":  "PDF/A-2b (Langzeitarchivierung)",
        "pdfa-1":  "PDF/A-1b",
        "pdfa-3":  "PDF/A-3b",
        "pdfa":    "PDF/A-2b",
        "pdf":     "PDF (Standard)",
    }
    format_display = format_labels.get(output_type, output_type)
    print()
    print("=" * 70)
    print("  KONFIGURATION")
    print("=" * 70)
    print(f"  Verzeichnis:       {target_dir}")
    print(f"  OCR-Sprache:       {config.get('ocr_language', DEFAULT_LANGUAGE)}")
    print(f"  PDF-Rasterizer:    {config.get('rasterizer', 'ghostscript')}")
    print(f"  Ausgabeformat:     {format_display}")
    print(f"  Worker:            {num_workers}")
    print(f"  Modus:             {'Hotfolder (watch)' if args.watch else ('Vorab zählen (ETA)' if count_first else 'Generator (kein ETA)')}")
    print(f"  Temp-Verzeichnis:  {temp_dir}")
    print(f"  CSV-Log:           {os.path.abspath(CSV_LOG_FILE)}")
    if resume_path:
        print(f"  Resume-Datei:      {resume_path}")
    print(f"  veraPDF:           {'aktiv' if config.get('has_verapdf') else 'nicht verfügbar'}")
    print(f"  Marker-Verify:     {'aktiv' if config.get('verify_markers') else 'aus'}")
    print()
    print("  Verarbeitungslogik:")
    print("    • Magic-Number-Check (%PDF- in den ersten 1024 Bytes)")
    print("    • RAM-Schutz nur bei klar unzureichendem RAM (>1.2× Bedarf, ab 50 Seiten)")
    print("    • Disk-Check dynamisch (max(100 MB, 3×Dateigröße))")
    print("    • skip_text=True   → Seiten mit Text werden übersprungen")
    print("    • OCR-Marker       → Info-Dict /Subject + XMP dc:description")
    print("    • XMP-Fallback     → pdfaid:part/conformance wird mitgelesen")
    print("    • Backup           → Original gesichert, erst nach Erfolg gelöscht")
    print("    • Retry × 3        → ab Versuch 2 ohne clean/deskew/unpaper")
    print("    • Farbraum-Fix     → RGB-Konvertierung (proaktiv bei PDF/A)")
    print("    • Pillow-Toleranz  → LOAD_TRUNCATED_IMAGES = True, MAX_IMAGE_PIXELS = None")
    print("    • Signatur-Skip    → digital signierte PDFs ohne Retry übersprungen")
    print("    • Stream-Schutz    → korrupte Inhaltsströme ohne Retry abgebrochen")
    print("    • Bitmap-Fallback  → bei pypdfium2-Speicherlimit Wechsel auf Ghostscript")
    print("    • Temp-Kopie       → AV-Scanner-Toleranz (Retry + Stabilisierung)")
    print(f"    • Worker-Timeout   → {WORKER_TIMEOUT // 60} Min pro Datei (parallel, Hard-Kill + Pool-Reset)")
    print(f"    • Heartbeat        → alle {HEARTBEAT_INTERVAL // 60} Min Lebenszeichen mit aktiven Workern")
    print("    • Log-Flush        → nach jeder Datei explizit auf Platte geschrieben")
    print("    • Run-Summary      → letzter Lauf wird zusätzlich als Datei abgelegt")
    if is_pdfa:
        print("    • PDF/A-Modus     → keine Linearisierung")
        if config.get("has_verapdf"):
            print("    • veraPDF-Check   → Marker spiegelt validierte Konformität")
        else:
            print("    • Ohne veraPDF    → maximal PDF/A-2b-Marker (konservativ)")
        if config.get("verify_markers"):
            print("    • Marker-Verify   → alte Marker werden geprüft und korrigiert")
    if resume_path:
        print("    • Resume-Log      → verarbeitete Dateien beim Neustart übersprungen")
    print("=" * 70)

    if not args.auto:
        if not ask_yes_no("\nJetzt starten?"):
            print("Abgebrochen.")
            if os.path.exists(temp_dir) and temp_dir != _ORIG_WINDOWS_TEMP:
                shutil.rmtree(temp_dir, ignore_errors=True)
            return

    # --- Verarbeitung ---
    start_time   = datetime.now()
    stats_result = None
    run_aborted   = False
    abort_reason: Optional[str] = None

    config["jobs"] = recommended_jobs_per_worker(os.cpu_count() or 1, num_workers)

    if args.watch and args.dry_run:
        # Der Watch-Modus kannte den Probelauf nicht: '--watch --dry-run'
        # lief als ECHTER Dauerbetrieb und ersetzte PDFs, obwohl der
        # Anwender einen Probelauf angefordert hat. Zusaetzlich haette
        # er ohne Einzelinstanz-Sperre gearbeitet, weil die Sperre oben
        # bei --dry-run bewusst uebersprungen wird.
        #
        # Ein dauerhaft ueberwachender Probelauf ergibt keinen Sinn -
        # er wuerde dieselben Dateien endlos erneut melden. Deshalb wird
        # der Bestand einmal geprueft und danach beendet.
        print("\n⚠️  --dry-run und --watch zusammen: der Hotfolder-Betrieb wird")
        print("    NICHT gestartet. Stattdessen einmalige Prüfung des Bestands,")
        print("    danach Ende. Für den echten Dauerbetrieb --dry-run weglassen.")
        stats_result = dry_run_directory(target_dir, config, None)
    elif args.watch:
        run_watch_mode(target_dir, config, temp_dir, workers=num_workers,
                       initial_scan=not args.no_initial_scan)
    else:
        print("\n" + "=" * 70)
        print("  STARTE PDF-VERARBEITUNG")
        print("=" * 70)

        try:
            if args.dry_run:
                # Probelauf: keine Worker, kein Temp-Verzeichnis, kein
                # Schreibzugriff - und die Resume-Datei bleibt unangetastet.
                stats_result = dry_run_directory(target_dir, config, resume_path)
            else:
                stats_result = process_directory(
                    target_dir, config, temp_dir,
                    count_first, num_workers, resume_path,
                )
                if stats_result and not shutdown_event.is_set() and resume_path:
                    delete_resume_file(resume_path)
            if shutdown_event.is_set():
                run_aborted = True
                abort_reason = "Graceful Shutdown"
        except KeyboardInterrupt:
            print("\n\n⚠️  Abgebrochen durch Benutzer (Ctrl+C)")
            run_aborted = True
            abort_reason = "KeyboardInterrupt"
        except Exception as e:
            print(f"\n❌ KRITISCHER FEHLER: {_fmt_exc(e)}")
            traceback.print_exc()
            run_aborted = True
            abort_reason = _fmt_exc(e)

    # --- Aufräumen ---
    print("\n🧹 Temp-Verzeichnis aufräumen ...")
    if os.path.exists(temp_dir) and temp_dir != _ORIG_WINDOWS_TEMP:
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
            print(f"  ✓ Gelöscht: {temp_dir}")
        except Exception as e:
            print(f"  ⚠️  {_fmt_exc(e)}")
    _cleanup_windows_temp()

    # --- Abschluss ---
    end_time = datetime.now()
    duration = end_time - start_time

    print()
    print("=" * 70)
    print("  VERARBEITUNG ABGESCHLOSSEN")
    print("=" * 70)
    print(f"  Ende:  {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Dauer: {str(duration).split('.')[0]}")
    print()

    if stats_result:
        print_detailed_stats(stats_result)
        write_run_summary(
            stats_result, target_dir, output_type,
            start_time, end_time,
            aborted=run_aborted, abort_reason=abort_reason,
            summary_path=(RUN_SUMMARY_DRYRUN if args.dry_run else RUN_SUMMARY_FILE),
            ist_probelauf=bool(args.dry_run),
        )

    _flush_log_handlers()

    print()
    print(f"  Fehler-Log:  {os.path.abspath(LOG_FILE)}")
    print(f"  Detail-Log:  {os.path.abspath(DETAILED_LOG_FILE)}")
    print(f"  CSV-Log:     {os.path.abspath(CSV_LOG_FILE)}")
    print(f"  Run-Summary: "
          f"{os.path.abspath(RUN_SUMMARY_DRYRUN if args.dry_run else RUN_SUMMARY_FILE)}")
    if resume_path and os.path.exists(resume_path):
        print(f"  Resume-Log:  {os.path.abspath(resume_path)} (erhalten – Neustart möglich)")
    print("=" * 70)


# ==================================================================
# Einstiegspunkt
# ==================================================================
if __name__ == "__main__":
    multiprocessing.freeze_support()
    try:
        multiprocessing.set_start_method("spawn", force=True)
    except RuntimeError:
        pass
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n❌ Abgebrochen (Ctrl+C)")
    except Exception as e:
        print(f"\n❌ UNBEKANNTER FEHLER: {_fmt_exc(e)}")
        traceback.print_exc()
    finally:
        if "--auto" not in sys.argv:
            try:
                input("\nDrücke Enter zum Beenden ...")
            except EOFError:
                pass
