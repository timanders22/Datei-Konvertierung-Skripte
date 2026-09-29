# ==================================================================
# EXCEL FONT-ERSETZUNG
# ==================================================================
# ANLEITUNG:
# 1. Python installieren: https://www.python.org/
# 2. In der Konsole (CMD): pip install pywin32 tqdm psutil
# 3. Falls pywin32 erstmals installiert: python Scripts/pywin32_postinstall.py -install
# 4. In der Konsole (CMD): python 4b_ersetze_font_in_excel.py
#
# AUFRUF-OPTIONEN (argparse):
#   --auto              Trust-Center-Check und alle Rückfragen überspringen
#   --dir <PFAD>        Zielverzeichnis
#   --font <NAME>       Ziel-Schriftart (Standard: Arial)
#   --count-first       Dateien vorab zählen (ETA-Anzeige)
#   --no-meta           Metadaten nicht entfernen
#
# WICHTIG: EXCEL SICHERHEITSEINSTELLUNGEN (TRUST CENTER)
# ------------------------------------------------------------------
# Excel öffnen → Datei → Optionen → Trust Center →
# "Einstellungen für das Trust Center..."
#
# 1. GESCHÜTZTE ANSICHT:
#    ☐ Dateien aus dem Internet in geschützter Ansicht öffnen
#    ☐ Dateien an potenziell unsicheren Orten in geschützter Ansicht öffnen
#    ☐ Anlagen aus Outlook in geschützter Ansicht öffnen
#    → ALLE 3 DEAKTIVIEREN
#
# 2. MAKROEINSTELLUNGEN:
#    ⦿ Alle Makros aktivieren
#    ☑ Zugriff auf das VBA-Projektobjektmodell vertrauen
#    (Erforderlich für korrekte .xlsm/.xlsx-Unterscheidung)
#
# 3. VERTRAUENSWÜRDIGE SPEICHERORTE:
#    ☑ Vertrauenswürdige Speicherorte im Netzwerk zulassen
#    Folgende Speicherorte hinzufügen:
#    • %LOCALAPPDATA%\Dateimigration-Arbeitskopien (temporärer Arbeitsordner des Skripts)
#    • Q:\ oder \\server\dfs (Quelldateien)
#    Jeweils: ☑ Unterordner ... sind ebenfalls vertrauenswürdig
#
# HINWEIS ZU METADATEN-ENTFERNUNG:
# ------------------------------------------------------------------
# Die Option "Persönliche Informationen" setzt das Excel-interne Flag
# RemovePersonalInformation=TRUE. Dieses Flag bleibt dauerhaft in der
# Arbeitsmappe aktiv. Excel entfernt dann bei JEDEM späteren manuellen
# Speichern erneut persönliche Informationen (Autor, letzter Bearbeiter).
# Wer das nicht wünscht, muss die Option in Excel manuell deaktivieren:
# Datei → Informationen → "Auf Probleme überprüfen" → Dokumentprüfung
# (oder Trust Center → Datenschutzoptionen).
#
# HINWEIS ZUM BLATTSCHUTZ:
# ------------------------------------------------------------------
# Nicht-passwortgeschützte Blätter werden vor der Schriftartänderung
# ENTSPERRT und bleiben anschließend ENTSPERRT (kein Re-Protect).
# Passwortgeschützte Blätter werden übersprungen (das Skript versucht
# kein Passwort-Brute-Force).
#
# HINWEIS ZU AV-SCANNERN:
# ------------------------------------------------------------------
# Nach Excel-SaveAs beginnen Antiviren-Programme (Defender, Sophos, ...)
# die Datei zu scannen und sperren sie kurz. Das Skript wartet daher nach
# dem Anlegen der Temp-Kopie auf Freigabe und wiederholt alle Datei-
# Operationen (move/copy/replace/remove) mit exponentiellem Backoff bis
# zu ca. 60 Sekunden pro Operation.
#
# HINWEIS FÜR PyInstaller-Build (Umwandlung in .exe):
#   # Einmalig: PyInstaller installieren
#   pip install pyinstaller
#
#   # Build als ein einzeiliger Befehl (am sichersten in jeder Shell):
#   python -m PyInstaller --onefile --noupx --noconfirm --clean --console --icon="python_icon.ico" --name "4b_ersetze_font_in_excel" --collect-submodules win32com --hidden-import pywintypes --hidden-import pythoncom --hidden-import win32api --hidden-import win32com.client --hidden-import win32process --hidden-import win32file --hidden-import win32con --hidden-import win32timezone --hidden-import winreg --hidden-import psutil --hidden-import tqdm 4b_ersetze_font_in_excel.py
#
#   # Mehrzeilig (Fortsetzungszeichen je nach Shell):
#   #   cmd.exe        : ^ am Zeilenende
#   #   PowerShell     : ` (Backtick) am Zeilenende
#   #   bash/WSL/Git-Bash: \ am Zeilenende
#
#   Optionen:
#     --noupx       vermeidet AV-False-Positives bei UPX-gepackten Binaries.
#     --noconfirm   überschreibt vorhandenes dist/-Verzeichnis ohne Rückfrage.
#     --clean       leert den PyInstaller-Cache vor dem Build (sauberer Neubau).
#     --hidden-import  pythoncom/pywintypes/win32com.client werden von win32com
#                      dynamisch geladen und sind ohne explizite Angabe
#                      manchmal nicht im Build enthalten.
#     --collect-submodules win32com  erfasst die dynamisch geladenen
#                      Office-Constants-Submodule.
#     Code-Signing nach Build empfohlen für Enterprise-Rollout.
#
#   Bitness (32 vs. 64 Bit) bei gemischten Office-Umgebungen:
#     PyInstaller erzeugt eine EXE in der Architektur des verwendeten Python-
#     Interpreters (64-Bit-Python -> x64-EXE; 32-Bit-Python -> x86-EXE). Excel
#     registriert seinen COM-Server als LocalServer32 (out-of-process,
#     EXCEL.EXE); das Windows-COM-Subsystem marshallt Aufrufe zwischen
#     x64-Aufrufer und x86-Server (und umgekehrt) automatisch über DCOM/LRPC.
#     Eine x64-EXE arbeitet daher auch mit 32-Bit-Excel zusammen (und umgekehrt).
#     Bitness-Match liefert minimal bessere Performance, ist aber nicht
#     erforderlich. 64-Bit-Python ist eine sichere Default-Wahl, da modernes
#     Office (2019+, M365) standardmäßig x64 ist.
#
# Stand: 13.06.2026
# ==================================================================

import argparse
import os
import re
import sys
import csv
import stat
import shutil
import signal
import hashlib
import logging
import time
import threading
import uuid
import zipfile
import tempfile
from datetime import datetime
from typing import Optional, Callable, Any

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
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

# Zuletzt verwendet: Ausgangszustand merken, damit spaeter nur die
# eigenen Eintraege verschwinden (siehe _cleanup_user_recent).
if gem is not None and hasattr(gem, "recent_momentaufnahme"):
    gem.recent_momentaufnahme()

# Symbolschriften (Wingdings, Symbol ...) werden nie ersetzt: dort steht
# hinter jedem Zeichencode ein Bild, und aus einem Haekchen wuerde in Arial
# ein Buchstabe. Erkennung in _gemeinsam.py (feste Liste plus alle
# installierten Schriften mit Symbolzeichensatz); fehlt es, gilt diese Liste.
_SYMBOLSCHRIFTEN_RUECKFALL = (
    "Symbol", "Wingdings", "Wingdings 2", "Wingdings 3", "Webdings",
    "Marlett", "MT Extra", "Bookshelf Symbol 7", "MS Reference Specialty",
    "MS Outlook", "ZapfDingbats", "Zapf Dingbats", "Segoe MDL2 Assets",
    "Segoe Fluent Icons",
)


def _symbolschrift_namen() -> list:
    if gem is not None and hasattr(gem, "symbolschrift_namen"):
        return gem.symbolschrift_namen()
    return list(_SYMBOLSCHRIFTEN_RUECKFALL)


def _ist_symbolschrift(name) -> bool:
    if gem is not None and hasattr(gem, "ist_symbolschrift"):
        return gem.ist_symbolschrift(name)
    return bool(name) and str(name).strip().lower() in {
        n.lower() for n in _SYMBOLSCHRIFTEN_RUECKFALL}

try:
    import pythoncom
    import win32com.client
    import win32process
    import win32file
    import pywintypes
    import psutil
    from tqdm import tqdm
except ImportError as _e:
    _mod = str(_e).split("'")[1].split(".")[0] if "'" in str(_e) else str(_e)
    _pkg = {"pythoncom": "pywin32", "win32com": "pywin32",
            "win32process": "pywin32", "win32file": "pywin32",
            "pywintypes": "pywin32",
            "psutil": "psutil", "tqdm": "tqdm"}.get(_mod, _mod)
    print("=" * 66)
    print("FEHLENDES MODUL – bitte in der Konsole (CMD) installieren:")
    print(f"   pip install {_pkg}")
    print("=" * 66)
    sys.exit(1)

try:
    import winreg
except ImportError:
    winreg = None

# msoffcrypto ist OPTIONAL (nicht in der Standard-Installationsanleitung).
# Wenn vorhanden, erkennt es klassische BIFF-Verschluesselung in .xls
# zuverlaessig (FilePass-Record) – sonst greift der manuelle Marker-Scan.
try:
    import msoffcrypto as _msoffcrypto
except ImportError:
    _msoffcrypto = None

# ==================================================================
# Shell-Folder-Lookup (OneDrive-Redirect-tauglich)
# ==================================================================
def _resolve_user_shell_folder(value_name: str, fallback_subdir: str) -> str:
    user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    fallback = os.path.join(user_profile, fallback_subdir)
    if winreg is None:
        return fallback
    try:
        with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders") as key:
            val, _ = winreg.QueryValueEx(key, value_name)
            expanded = os.path.expandvars(val)
            if expanded and os.path.isdir(expanded):
                return expanded
    except (OSError, FileNotFoundError):
        pass
    return fallback

# ==================================================================
# COM-Konstanten
# ==================================================================
COM_TRUE  = -1
COM_FALSE =  0

XL_CALC_MANUAL = -4135
XL_CALC_AUTO   = -4105

XL_FORMAT_XLSB = 50
XL_FORMAT_XLSX = 51
XL_FORMAT_XLSM = 52
XL_FORMAT_XLTX = 54
XL_FORMAT_XLTM = 53

_EXT_TO_FORMAT = {
    ".xlsx": XL_FORMAT_XLSX,
    ".xlsm": XL_FORMAT_XLSM,
    ".xlsb": XL_FORMAT_XLSB,
    ".xltx": XL_FORMAT_XLTX,
    ".xltm": XL_FORMAT_XLTM,
}

MSO_GROUP = 6
MSO_CHART = 3

FILE_READ_ATTRIBUTES  = 0x0080
FILE_WRITE_ATTRIBUTES = 0x0100
FILE_SHARE_DELETE     = 0x00000004

VBEXT_CT_DOCUMENT = 100

XL_EXCEL4_MACRO_SHEET      = 3
XL_EXCEL4_INTL_MACRO_SHEET = 4

# ==================================================================
# Metadaten-Optionen
# ==================================================================
BUILTIN_PROPS_TO_CLEAR = {
    "Author":         "Autor (Ersteller)",
    "Last author":    "Letzter Bearbeiter",
    "Title":          "Titel",
    "Subject":        "Betreff",
    "Keywords":       "Stichwörter",
    "Comments":       "Notizen/Kommentare (Eigenschaft)",
    "Company":        "Firma",
    "Manager":        "Vorgesetzter",
    "Category":       "Kategorie",
    "Hyperlink base": "Hyperlink-Basis",
}

METADATA_OPTIONS = {
    "props":    "Integrierte Dokumenteigenschaften (Autor, Titel, Firma …)",
    "custom":   "Benutzerdefinierte Dokumenteigenschaften (alle löschen)",
    "personal": "Persönliche Informationen (RemovePersonalInformation-Flag)",
    "comments": "Zell-Kommentare (alter Stil – INHALT wird gelöscht!)",
    "threads":  "Thread-Kommentare / Notizen (Excel 365 – INHALT wird gelöscht!)",
}

# ==================================================================
# Konfiguration
# ==================================================================
DEFAULT_FONT_NAME = "Arial"
NEW_FONT_NAME     = DEFAULT_FONT_NAME
auto_mode         = False

# Gemeinsamer Arbeitsordner der Office-Skripte. Bis 29.09.2026 lag er in
# "Dokumente"; bei OneDrive-Ordnersicherung wanderte so jede Arbeitskopie
# in die Cloud. %LOCALAPPDATA% wird nie umgeleitet.
ARBEITS_BASIS     = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP")
    or _resolve_user_shell_folder("Personal", "Documents"),
    "Dateimigration-Arbeitskopien")
TEMP_BASE_PATH    = os.path.join(ARBEITS_BASIS, "4b_ersetze_font_in_excel")
TEMP_PROCESS_PATH = TEMP_BASE_PATH
MAX_PATH_LEN      = 240
WINDOWS_MAX_PATH  = 260

_preserved_temp_files: set = set()

OPEN_TIMEOUT_SECONDS        = 60
SAVE_TIMEOUT_SECONDS        = 180
MAX_SHAPE_RECURSION_DEPTH   = 20
MAX_POINTS_PER_SERIES       = 500
FILE_PROCESSING_ATTEMPTS    = 2
ERROR_RATE_THRESHOLD        = 0.20
MIN_FILES_FOR_RATE_CHECK    = 30
EXCEL_MEMORY_LIMIT_MB       = 2048
EXCEL_RESTART_INTERVAL      = 200

AV_MAX_RETRIES              = 40
AV_INITIAL_DELAY            = 0.2
AV_MAX_DELAY                = 3.0
AV_RETRYABLE_WIN_ERRORS     = {5, 32, 33}
AV_POST_COPY_WAIT_SECONDS   = 15.0

SKIP_FILE_PREFIXES = ("~$", "._", ".~tmp")
SKIP_FILE_SUFFIXES = (".tmp", ".bak")

# Verzeichnisnamen, die nie betreten werden: geloeschte Arbeitsmappen im
# $RECYCLE.BIN wuerden sonst mitverarbeitet (und .xls-Originale ersetzt);
# ~snapshot/.snapshot sind read-only NAS-Schattenkopien.
EXCLUDE_DIR_NAMES = {"$recycle.bin", "system volume information",
                     "~snapshot", ".snapshot"}

FILE_ATTRIBUTE_REPARSE_POINT = 0x0400

# Nur DAUERHAFT erledigte Status ins Resume schreiben. SKIPPED (Passwort/
# OLE-verschluesselt) wird beim naechsten Lauf erneut versucht – die
# Erkennung dafuer ist billig (vor bzw. beim Open).
RESUME_STATUSES = ("SUCCESS", "PARTIAL")

def _get_base_dir() -> str:
    try:
        if getattr(sys, "frozen", False):
            return os.path.dirname(os.path.abspath(sys.executable))
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()

BASE_DIR = _get_base_dir()
_LOG_TIMESTAMP    = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE          = os.path.join(
    BASE_DIR, f"4b_ersetze_font_in_excel_{_LOG_TIMESTAMP}.log")
DETAILED_LOG_FILE = os.path.join(
    BASE_DIR, f"4b_ersetze_font_in_excel_detailed_{_LOG_TIMESTAMP}.log")
RUN_SUMMARY_FILE  = os.path.join(
    BASE_DIR, "4b_ersetze_font_in_excel_last_run.txt")
CONVERSIONS_CSV   = os.path.join(
    BASE_DIR, f"4b_ersetze_font_in_excel_konvertierungen_{_LOG_TIMESTAMP}.csv")
# Tatsaechliches Log-Verzeichnis; wird in _setup_logging ggf. auf den
# Fallback-Ordner umgebogen (wenn BASE_DIR nicht beschreibbar ist).
_LOG_DIR = BASE_DIR

# ==================================================================
# Logging
# ==================================================================
def _setup_logging() -> tuple:
    # LOG_FILE und DETAILED_LOG_FILE gehoeren MIT in die global-Zeile:
    # im Fallback-Zweig unten werden die Handler auf einen anderen
    # Ordner umgelenkt, die Abschlussmeldungen zeigten aber weiterhin
    # die urspruenglichen, NICHT beschreibbaren Pfade an - der Anwender
    # suchte die Protokolle dort vergeblich.
    global RUN_SUMMARY_FILE, CONVERSIONS_CSV, _LOG_DIR
    global LOG_FILE, DETAILED_LOG_FILE
    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    try:
        fh = logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8-sig")
        dh = logging.FileHandler(DETAILED_LOG_FILE, mode="w", encoding="utf-8-sig")
    except (PermissionError, OSError) as e:
        fallback = os.path.join(
            _resolve_user_shell_folder("Personal", "Documents"),
            "4b_ersetze_font_in_excel_logs")
        os.makedirs(fallback, exist_ok=True)
        fh_path = os.path.join(fallback, os.path.basename(LOG_FILE))
        dh_path = os.path.join(fallback, os.path.basename(DETAILED_LOG_FILE))
        fh = logging.FileHandler(fh_path, mode="w", encoding="utf-8-sig")
        dh = logging.FileHandler(dh_path, mode="w", encoding="utf-8-sig")
        # Run-Summary, Konvertierungs-CSV und Resume-Dateien in denselben
        # beschreibbaren Fallback-Ordner umlenken.
        _LOG_DIR        = fallback
        LOG_FILE          = fh_path
        DETAILED_LOG_FILE = dh_path
        RUN_SUMMARY_FILE = os.path.join(fallback, os.path.basename(RUN_SUMMARY_FILE))
        CONVERSIONS_CSV  = os.path.join(fallback, os.path.basename(CONVERSIONS_CSV))
        print(f"[WARN] Log-Ziel nicht beschreibbar ({e}) → {fallback}")

    fh.setFormatter(fmt)
    fl = logging.getLogger("FileLogger")
    fl.setLevel(logging.ERROR)
    fl.handlers.clear()
    fl.addHandler(fh)

    dh.setFormatter(fmt)
    dl = logging.getLogger("DetailLogger")
    dl.setLevel(logging.DEBUG)
    dl.handlers.clear()
    dl.addHandler(dh)

    return fl, dl

file_logger, detail_logger = _setup_logging()



# ==================================================================
# Gemeinsames Laufprotokoll (migration.jsonl)
# ==================================================================
# Jedes Skript schreibt sein eigenes Format: CSV mit Semikolon, CSV mit
# Komma, .log, XLSX, teils mit BOM, teils ohne. Der Gesamtfortschritt
# ueber die elf Schritte liess sich damit nicht auswerten - etwa die
# Frage, welche Dateien in Schritt 2 liegen blieben und in Schritt 7
# wieder auftauchen. Diese Zeile ERGAENZT die bestehenden Protokolle.
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

def log_conversion(old_path: str, new_path: str) -> None:
    _protokoll(new_path or old_path, "konvertiert", "OK",
               f"aus {old_path}")
    # Nachweisliste der Format-Konvertierungen (.xls/.xlt -> neu, und
    # Ausweichnamen bei Kollision) mit altem und neuem Pfad fuers Archiv.
    if not CONVERSIONS_CSV:
        return
    try:
        needs_header = not os.path.exists(CONVERSIONS_CSV)
        with open(CONVERSIONS_CSV, "a", encoding="utf-8-sig", newline="") as fh:
            writer = csv.writer(fh, delimiter=";")
            if needs_header:
                writer.writerow(["timestamp", "alter_pfad", "neuer_pfad"])
            writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                old_path, new_path,
            ])
    except Exception as e:
        detail_logger.warning(f"Konvertierungs-CSV nicht schreibbar: {e}")


def write_run_summary(stats: dict, directory: str, font_name: str,
                      metadata_selected: dict, duration) -> None:
    if not RUN_SUMMARY_FILE:
        return
    try:
        total = sum(stats.values())
        n_meta = sum(1 for v in (metadata_selected or {}).values() if v)
        with open(RUN_SUMMARY_FILE, "w", encoding="utf-8-sig") as fh:
            fh.write("EXCEL FONT-ERSETZUNG – LETZTER LAUF\n")
            fh.write("=" * 60 + "\n")
            fh.write(f"Verzeichnis:    {directory}\n")
            fh.write(f"Schriftart:     {font_name}\n")
            fh.write(f"Metadaten:      "
                     f"{('Ja (' + str(n_meta) + ' Optionen)') if n_meta else 'Nein'}\n")
            fh.write(f"Dauer:          {str(duration).split('.')[0]}\n")
            fh.write("\n")
            fh.write(f"Erfolgreich:    {stats.get('SUCCESS', 0)}\n")
            fh.write(f"Teilweise:      {stats.get('PARTIAL', 0)}\n")
            fh.write(f"Übersprungen:   {stats.get('SKIPPED', 0)}\n")
            fh.write(f"Fehler:         {stats.get('ERROR', 0)}\n")
            fh.write(f"Gesamt:         {total}\n")
    except Exception as e:
        detail_logger.warning(f"Run-Summary nicht schreibbar: {e}")


# ==================================================================
# Auto-Resume (Neustart-Faehigkeit grosser Laeufe)
# ==================================================================
def get_auto_resume_path(start_dir: str) -> str:
    key = hashlib.md5(
        os.path.normcase(os.path.abspath(start_dir)).encode("utf-8")
    ).hexdigest()[:12]
    return os.path.join(_LOG_DIR, f"4b_ersetze_font_in_excel_resume_{key}.txt")


def load_resume_set(resume_path: str) -> set:
    if not resume_path or not os.path.exists(resume_path):
        return set()
    try:
        with open(resume_path, encoding="utf-8-sig") as fh:
            return {l.strip().lower() for l in fh if l.strip()}
    except Exception as e:
        detail_logger.warning(f"Resume-Datei nicht lesbar: {e}")
        return set()


def append_resume(resume_path: str, file_path: str) -> None:
    if not resume_path:
        return
    try:
        with open(resume_path, "a", encoding="utf-8") as fh:
            fh.write(file_path + "\n")
    except Exception as _e:
        detail_logger.debug(f"append_resume: Exception verworfen: {_e!r}")


def delete_resume_file(resume_path: str) -> None:
    try:
        if resume_path and os.path.exists(resume_path):
            os.remove(resume_path)
    except Exception as _e:
        detail_logger.debug(f"delete_resume_file: Exception verworfen: {_e!r}")


# ==================================================================
# Globale Excel-Referenz und PID
# ==================================================================
excel_app_global: Optional[win32com.client.CDispatch] = None
excel_pid_global: Optional[int] = None
# Prozess-Erstellungszeit der Skript-Excel-Instanz. Schuetzt zusammen mit
# dem Prozessnamen-Check vor PID-Recycling: Stirbt Excel waehrend einer
# langen Watchdog-Wartezeit und vergibt Windows die PID an eine neue,
# vom Benutzer geoeffnete Excel-Sitzung, unterscheidet nur die
# Erstellungszeit die beiden Prozesse.
excel_create_time_global: Optional[float] = None

# Urspruengliche Connect-Zustaende der COM-Add-Ins (ProgId -> bool).
# COMAddIn.Connect=False schreibt den LoadBehavior-Wert DAUERHAFT in die
# Benutzer-Registry - ohne Wiederherstellung blieben alle Add-Ins
# (Acrobat, SAP, ELO, ...) auch fuer die regulaere Excel-Nutzung des
# Benutzers deaktiviert.
_orig_addin_states: dict = {}

# ==================================================================
# Signal-Handler
# ==================================================================
def _signal_handler(sig, frame) -> None:
    print("\n\n*** ABBRUCH durch Benutzer – raeume auf ...")

    global excel_app_global
    if excel_app_global is not None:
        try:
            _restore_com_addins(excel_app_global)
        except Exception as _e:
            detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")
        try:
            excel_app_global.Quit()
            time.sleep(1)
        except Exception as _e:
            detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")
        excel_app_global = None

    _kill_specific_excel()
    _safe_cleanup_temp()
    # Best effort Cache-Cleanup beim Strg+C-Abbruch. Schnell genug fuer
    # einen Signal-Handler-Kontext (nur lokale Datei-Loeschungen).
    try:
        _cleanup_excel_inetcache()
    except Exception as _e:
        detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")
    try:
        _cleanup_user_recent()
    except Exception as _e:
        detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")

    try:
        pythoncom.CoUninitialize()
    except Exception as _e:
        detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")

    sys.exit(130)

# ==================================================================
# AV-Scanner-sichere Datei-Operationen
# ==================================================================
def _av_safe_op(op: Callable[[], Any],
                max_retries: int = AV_MAX_RETRIES,
                initial_delay: float = AV_INITIAL_DELAY,
                max_delay: float = AV_MAX_DELAY) -> Any:
    delay    = initial_delay
    last_exc = None
    for i in range(max_retries):
        try:
            return op()
        except PermissionError as e:
            last_exc = e
        except OSError as e:
            winerr = getattr(e, "winerror", None)
            if winerr is not None and winerr not in AV_RETRYABLE_WIN_ERRORS:
                raise
            last_exc = e
        except pywintypes.error as e:
            if e.winerror not in AV_RETRYABLE_WIN_ERRORS:
                raise
            last_exc = e
        time.sleep(delay)
        delay = min(delay * 1.5, max_delay)
    if last_exc is not None:
        raise last_exc


def _av_safe_remove(path: str) -> None:
    _av_safe_op(lambda: os.remove(path))


def _av_safe_replace(src: str, dst: str) -> None:
    _av_safe_op(lambda: os.replace(src, dst))


def _av_safe_move(src: str, dst: str) -> None:
    _av_safe_op(lambda: shutil.move(src, dst))


def _av_safe_copy2(src: str, dst: str) -> None:
    _av_safe_op(lambda: shutil.copy2(src, dst))


def _wait_file_released(path: str, timeout: float = 30.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with open(path, "ab"):
                return True
        except (PermissionError, OSError):
            time.sleep(0.3)
    return False

# ==================================================================
# Hilfsfunktionen
# ==================================================================
def log_error(file_path: str, exc: Exception) -> None:
    file_logger.error(f"Datei: {file_path}\n  → {exc}\n")
    detail_logger.error(f"Datei: {file_path}\n  → {exc}\n")


def _rmtree_onerror(func, path, exc_info):
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception as _e:
        detail_logger.debug(f"_rmtree_onerror: Exception verworfen: {_e!r}")


# ==================================================================
# Startwarnung: laufende Excel-Sitzungen
# ==================================================================
def find_running_excel_pids() -> list:
    """PIDs aller Excel-Prozesse des angemeldeten Benutzers."""
    pids = []
    try:
        me = psutil.Process().username().lower().split("\\")[-1].split("@")[0]
    except Exception:
        return pids
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            nm = proc.info.get("name") or ""
            if "EXCEL.EXE" not in nm.upper():
                continue
            try:
                u = proc.username().lower().split("\\")[-1].split("@")[0]
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
            if u == me:
                pids.append(proc.info["pid"])
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
    return pids


def warn_running_excel(auto_mode: bool = False) -> None:
    """Warnt vor bereits laufenden Excel-Sitzungen.

    Excel wird per COM automatisiert. Laeuft bereits eine Sitzung des
    Anwenders, hat das zwei Auswirkungen, die man vorher kennen sollte:

      1. Die Automatisierung kann sich an die vorhandene Sitzung haengen.
         Warnhinweise werden dann abgeschaltet und das Fenster ausgeblendet
         - fuer den Anwender sieht das aus, als sei Excel abgestuerzt.
      2. Die Aufraeumroutinen des Skripts koennen die eigene Instanz nicht
         mehr sicher von der fremden unterscheiden.

    Die Sitzung wird NICHT beendet (dafuer sorgt die Momentaufnahme der
    fremden Prozess-IDs beim Start), aber ein sauberer Lauf setzt ein
    geschlossenes Excel voraus.
    """
    pids = find_running_excel_pids()
    if not pids:
        return
    pids_str = ", ".join(str(p) for p in pids)

    if auto_mode:
        detail_logger.warning(
            f"Automatikmodus: laufende Excel-Sitzungen (PID: {pids_str}) - "
            f"sie werden geschuetzt, aber nicht geschlossen."
        )
        return

    print()
    print("=" * 66)
    print("  WARNUNG: Excel laeuft bereits")
    print("=" * 66)
    print(f"  Gefundene Excel-Prozesse (PID): {pids_str}")
    print()
    print("  Ihre Sitzung wird vom Skript NICHT beendet. Waehrend des Laufs")
    print("  kann sie aber ausgeblendet werden und Warnhinweise sind")
    print("  abgeschaltet - das wirkt wie ein Absturz.")
    print("  Ausserdem laesst sich die eigene Automatisierungs-Instanz dann")
    print("  nicht mehr zuverlaessig von Ihrer Sitzung unterscheiden.")
    print()
    print("  EMPFEHLUNG: Excel jetzt schliessen und das Skript neu starten.")
    print("=" * 66)
    print()
    if not ask_yes_no("Trotzdem fortfahren?"):
        print("Abgebrochen. Bitte Excel schliessen und neu starten.")
        sys.exit(0)


def _kill_specific_excel() -> None:
    global excel_pid_global, excel_create_time_global
    if excel_pid_global is None:
        return
    try:
        proc = psutil.Process(excel_pid_global)
        # Process-Name-Check vor kill() - schuetzt vor PID-Recycling:
        # Nach einem regulaeren excel.Quit() gibt Windows die PID sofort
        # wieder frei. Ein anderer Prozess (Browser-Worker, AV-Scan-Thread,
        # svchost-Child) kann sie in Millisekunden erben. Ohne diese
        # Pruefung wuerde der Kill blind einen unbeteiligten Systemprozess
        # erwischen. Analog zur Implementierung in _kill_specific_word
        # im 4a_ersetze_font_in_word-Skript.
        try:
            proc_name = proc.name().lower()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            proc_name = ""
        if proc_name != "excel.exe":
            detail_logger.debug(
                f"PID {excel_pid_global} ist nicht EXCEL.EXE "
                f"(sondern '{proc_name}') – kein Kill (PID-Recycling)")
            return
        if excel_create_time_global is not None:
            try:
                if proc.create_time() != excel_create_time_global:
                    detail_logger.debug(
                        f"PID {excel_pid_global} hat abweichende "
                        f"Erstellungszeit – kein Kill (PID-Recycling)")
                    return
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                return
        proc.kill()
        proc.wait(timeout=3)
        detail_logger.debug(f"Skript-Excel beendet (PID {excel_pid_global})")
    except psutil.NoSuchProcess:
        pass
    except Exception as e:
        detail_logger.warning(
            f"Fehler beim Beenden von PID {excel_pid_global}: {e}")
    finally:
        excel_pid_global = None
        excel_create_time_global = None


def _disable_com_addins(app) -> None:
    """Deaktiviert COM-Add-Ins fuer den Massenlauf (Performance, keine
    Popups). ACHTUNG: Connect=False schreibt den LoadBehavior-Wert
    dauerhaft in die Benutzer-Registry. Die Original-Zustaende werden
    daher in _orig_addin_states gesichert - nur beim ERSTEN Aufruf je
    ProgId, damit geplante Excel-Neustarts den bereits deaktivierten
    Zustand nicht als "Original" ueberschreiben - und am Skriptende via
    _restore_com_addins wiederhergestellt."""
    try:
        for i in range(1, app.COMAddIns.Count + 1):
            try:
                ad = app.COMAddIns.Item(i)
                prog_id = str(ad.ProgId)
                if prog_id not in _orig_addin_states:
                    _orig_addin_states[prog_id] = bool(ad.Connect)
                if ad.Connect:
                    ad.Connect = False
                    detail_logger.debug(f"COM-Add-In deaktiviert: {ad.Description}")
            except Exception as _e:
                detail_logger.debug(f"_disable_com_addins: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_disable_com_addins: Exception verworfen: {_e!r}")


def _restore_com_addins(app) -> bool:
    """Stellt die urspruenglichen Connect-Zustaende der COM-Add-Ins
    wieder her. Rueckgabe False, wenn die Add-In-Collection nicht
    erreichbar war (Instanz tot) - dann muss der Fallback greifen."""
    if not _orig_addin_states:
        return True
    try:
        count = app.COMAddIns.Count
    except Exception:
        return False
    for i in range(1, count + 1):
        try:
            ad = app.COMAddIns.Item(i)
            prog_id = str(ad.ProgId)
            if prog_id in _orig_addin_states:
                orig = _orig_addin_states[prog_id]
                if bool(ad.Connect) != orig:
                    ad.Connect = orig
                    detail_logger.debug(
                        f"COM-Add-In wiederhergestellt: {ad.Description}")
        except Exception as _e:
            detail_logger.debug(f"_restore_com_addins: Exception verworfen: {_e!r}")
    return True


def _restore_com_addins_fallback() -> None:
    """Wiederherstellung der COM-Add-In-Zustaende, wenn keine lebende
    Excel-Instanz mehr existiert (Crash/Watchdog-Kill): startet eine
    kurzlebige Instanz nur fuer die Registry-Reparatur. Best effort."""
    if not _orig_addin_states:
        return
    if not any(_orig_addin_states.values()):
        return
    app = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        try:
            app.Visible       = COM_FALSE
            app.DisplayAlerts = COM_FALSE
        except Exception as _e:
            detail_logger.debug(f"_restore_com_addins_fallback: Exception verworfen: {_e!r}")
        _restore_com_addins(app)
    except Exception as e:
        detail_logger.warning(
            f"COM-Add-In-Wiederherstellung fehlgeschlagen: {e}\n"
            f"  Betroffene Add-Ins ggf. manuell reaktivieren "
            f"(Datei → Optionen → Add-Ins → COM-Add-Ins).")
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception as _e:
                detail_logger.debug(f"_restore_com_addins_fallback: Exception verworfen: {_e!r}")


def _safe_cleanup_temp() -> None:
    if TEMP_PROCESS_PATH == TEMP_BASE_PATH:
        return
    if not os.path.exists(TEMP_PROCESS_PATH):
        return
    try:
        rescue_files = [
            f for f in os.listdir(TEMP_PROCESS_PATH)
            if f.startswith("RESCUE_")
        ]
        preserved_alive = [
            p for p in _preserved_temp_files
            if os.path.normcase(os.path.abspath(p)).startswith(
                    os.path.normcase(os.path.abspath(TEMP_PROCESS_PATH)))
            and os.path.exists(p)
        ]
        if rescue_files or preserved_alive:
            all_protected = rescue_files + [os.path.basename(p) for p in preserved_alive]
            detail_logger.warning(
                f"GESCHÜTZTE DATEIEN IM TEMP-ORDNER – Ordner wird NICHT gelöscht!\n"
                f"  Pfad: {TEMP_PROCESS_PATH}\n"
                f"  Dateien ({len(all_protected)}):\n"
                + "\n".join(f"    * {f}" for f in all_protected)
                + "\n  → Bitte manuell sichern, dann Ordner löschen.")
            print("\n  GESCHÜTZTE DATEIEN vorhanden – Temp-Ordner NICHT gelöscht:")
            print(f"   {TEMP_PROCESS_PATH}")
            for f in all_protected:
                print(f"   * {f}")
            print("   → Bitte manuell sichern, dann Ordner löschen.")
        else:
            shutil.rmtree(TEMP_PROCESS_PATH, onerror=_rmtree_onerror)
    except Exception as e:
        detail_logger.warning(f"Aufräumen fehlgeschlagen: {e}")


def _cleanup_excel_inetcache() -> None:
    """Raeumt den Excel/Office-INetCache-Ordner auf
    (%LOCALAPPDATA%\\Microsoft\\Windows\\INetCache\\Content.MSO).

    Hinweis zum Pfadnamen: Excel hat KEINEN eigenen Content.Excel-Ordner.
    Excel, PowerPoint und andere Office-Anwendungen teilen sich den
    gemeinsamen MSO-Cache. Nur Word (Content.Word) und Outlook
    (Content.Outlook) haben eigene Cache-Ordner.

    Best effort: gelockte Dateien werden still uebersprungen (DEBUG).

    DARF NUR AUFGERUFEN WERDEN, WENN DAS SKRIPT-EXCEL NICHT LAEUFT
    (also nach Quit + _kill_specific_excel, oder am Skriptende).
    """
    local_appdata = os.environ.get("LOCALAPPDATA")
    if not local_appdata:
        return

    cache_dir = os.path.join(
        local_appdata, "Microsoft", "Windows", "INetCache", "Content.MSO")

    # SICHERHEITS-CHECK: Pfad muss auf den erwarteten Suffix enden -
    # schuetzt vor Pfad-Manipulation durch seltsame Env-Variablen.
    expected_suffix = os.path.join(
        "Microsoft", "Windows", "INetCache", "Content.MSO")
    if not cache_dir.lower().endswith(expected_suffix.lower()):
        return
    if not os.path.isdir(cache_dir):
        return

    files_deleted = 0
    files_skipped = 0
    bytes_freed   = 0

    try:
        for root, dirs, files in os.walk(cache_dir, topdown=False):
            for fname in files:
                fpath = os.path.join(root, fname)
                try:
                    fsize = os.path.getsize(fpath)
                except OSError:
                    fsize = 0
                try:
                    try:
                        os.chmod(fpath, stat.S_IWRITE)
                    except Exception as _e:
                        detail_logger.debug(f"_cleanup_excel_inetcache: Exception verworfen: {_e!r}")
                    os.remove(fpath)
                    files_deleted += 1
                    bytes_freed   += fsize
                except (OSError, PermissionError) as e:
                    files_skipped += 1
                    detail_logger.debug(
                        f"INetCache: Datei gesperrt, uebersprungen: "
                        f"{fpath} ({e})")
            for dname in dirs:
                dpath = os.path.join(root, dname)
                try:
                    os.rmdir(dpath)
                except OSError:
                    pass
    except Exception as e:
        detail_logger.debug(
            f"INetCache-Cleanup unerwarteter Fehler bei {cache_dir}: {e}")
        return

    if files_deleted > 0 or files_skipped > 0:
        mb = bytes_freed / (1024 * 1024)
        detail_logger.info(
            f"Office-INetCache aufgeraeumt (Content.MSO): "
            f"{files_deleted} Dateien geloescht ({mb:.1f} MB), "
            f"{files_skipped} gesperrt/uebersprungen")


def _cleanup_user_recent() -> None:
    # Bis 29.09.2026 loeschte diese Funktion JEDE Verknuepfung im
    # Windows-Ordner "Zuletzt verwendet" - auch die des Anwenders. Jetzt
    # entfernt sie nur Verknuepfungen, die seit dem Skriptstart entstanden
    # sind und auf das bearbeitete Verzeichnis oder den Arbeitsordner
    # zeigen (_gemeinsam.recent_eigene_entfernen). Ohne _gemeinsam.py
    # bleibt der Ordner unangetastet.
    entfernen = getattr(gem, "recent_eigene_entfernen", None)
    if entfernen is None:
        return
    try:
        geloescht, gesperrt = entfernen()
    except Exception as e:
        detail_logger.debug(f"Recent-Cleanup unerwarteter Fehler: {e}")
        return
    if geloescht or gesperrt:
        detail_logger.info(
            f"Windows-Recent: {geloescht} eigene Verknuepfung(en) entfernt, "
            f"{gesperrt} gesperrt/uebersprungen")


# Whitelist fuer den Windows-Temp-Cleanup: NUR Excel-/Office-eigene Reste
# und eigene Skript-Artefakte. NIEMALS pauschal leeren: Fremdprozesse
# legen aktive Daten ohne Lock in %TEMP% ab; blindes Loeschen zerstoert
# sie. gen_py (pywin32-COM-Cache) gehoert bewusst NICHT hierher: er wird
# von parallel laufenden COM-Skripten (3a/3b/3c/4a) aktiv genutzt.
# excel8.0/vbe/cvr sind Excel-/Office-typisch und bleiben.
_WINDOWS_TEMP_WHITELIST_PREFIXES = (
    "~$", "~df", "vbe", "excel8.0",
    "4b_ersetze_font_in_excel",
)


def _is_whitelisted_temp_entry(name: str) -> bool:
    nl = name.lower()
    if any(nl.startswith(p) for p in _WINDOWS_TEMP_WHITELIST_PREFIXES):
        return True
    if nl.startswith("cvr") and nl.endswith(".tmp"):
        return True
    return False


def _cleanup_windows_temp() -> None:
    # Best-effort Bereinigung von %LOCALAPPDATA%\Temp am Skriptende –
    # ausschliesslich per Whitelist bekannter Praefixe.
    # Gesperrte/in-Nutzung-Dateien werden stillschweigend übersprungen.
    win_temp = os.environ.get("TEMP") or os.environ.get("TMP") or tempfile.gettempdir()
    if not win_temp or not os.path.isdir(win_temp):
        return
    removed_files = 0
    removed_dirs  = 0
    skipped       = 0
    for entry in os.scandir(win_temp):
        try:
            if not _is_whitelisted_temp_entry(entry.name):
                skipped += 1
                continue
            if entry.is_file() or entry.is_symlink():
                try:
                    os.remove(entry.path)
                    removed_files += 1
                except Exception as _e:
                    detail_logger.debug(f"_cleanup_windows_temp: Exception verworfen: {_e!r}")
            elif entry.is_dir():
                try:
                    shutil.rmtree(entry.path, ignore_errors=True)
                    if not os.path.exists(entry.path):
                        removed_dirs += 1
                except Exception as _e:
                    detail_logger.debug(f"_cleanup_windows_temp: Exception verworfen: {_e!r}")
        except Exception as _e:
            detail_logger.debug(f"_cleanup_windows_temp: Exception verworfen: {_e!r}")
    try:
        detail_logger.info(
            f"Windows-Temp-Cleanup (Whitelist): {removed_files} Dateien, "
            f"{removed_dirs} Ordner aus {win_temp} entfernt, "
            f"{skipped} fremde Einträge unangetastet.")
    except Exception as _e:
        detail_logger.debug(f"_cleanup_windows_temp: Exception verworfen: {_e!r}")

# ==================================================================
# Trust-Center-Registry-Check
# ==================================================================
def check_trust_center_registry() -> list:
    issues: list = []
    if winreg is None:
        return issues

    ver = "16.0"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            rf"Software\Microsoft\Office\{ver}\Excel"):
            pass
    except OSError:
        return issues

    sec_base = rf"Software\Microsoft\Office\{ver}\Excel\Security"

    pv_flags = {
        "DisableInternetFilesInPV":    "Geschützte Ansicht für Internet-Dateien",
        "DisableUnsafeLocationsInPV":  "Geschützte Ansicht für unsichere Orte",
        "DisableAttachmentsInPV":      "Geschützte Ansicht für Outlook-Anlagen",
    }
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            sec_base + r"\ProtectedView") as key:
            for name, label in pv_flags.items():
                try:
                    val, _ = winreg.QueryValueEx(key, name)
                    if val != 1:
                        issues.append(f"Geschützte Ansicht aktiv: {label}")
                except FileNotFoundError:
                    issues.append(f"Geschützte Ansicht vermutlich aktiv: {label}")
    except OSError:
        issues.append("Registry-Pfad ProtectedView nicht gefunden (Einstellungen unklar)")

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sec_base) as key:
            try:
                val, _ = winreg.QueryValueEx(key, "AccessVBOM")
                if val != 1:
                    issues.append("VBA-Projektobjektmodell-Zugriff NICHT aktiviert")
            except FileNotFoundError:
                issues.append("VBA-Projektobjektmodell-Zugriff nicht gesetzt (default = aus)")
    except OSError:
        pass

    try:
        with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                sec_base + r"\Trusted Locations") as key:
            try:
                val, _ = winreg.QueryValueEx(key, "AllowNetworkLocations")
                if val != 1:
                    issues.append(
                        "'Vertrauenswürdige Speicherorte im Netzwerk zulassen' NICHT aktiviert")
            except FileNotFoundError:
                issues.append(
                    "'Vertrauenswürdige Speicherorte im Netzwerk' nicht gesetzt")
    except OSError:
        pass

    return issues

# ==================================================================
# COM-Watchdog  (Open / Save)
# ==================================================================
def _run_com_with_watchdog(op: Callable[[], Any],
                           excel_pid: Optional[int],
                           timeout: float,
                           op_name: str = "COM-Operation",
                           excel_create_time: Optional[float] = None) -> Any:
    done         = threading.Event()
    timeout_flag = [False]
    state_lock   = threading.Lock()

    def watchdog():
        if not done.wait(timeout):
            with state_lock:
                if not done.is_set():
                    timeout_flag[0] = True
                    if excel_pid:
                        # Process-Name- UND Erstellungszeit-Check vor
                        # kill() schuetzen vor PID-Recycling. Bei Save-
                        # Timeouts bis 180s kann Excel in der Wartezeit
                        # anderweitig sterben (AV-Kill, OOM, User-Kill),
                        # Windows die PID freigeben und an einen
                        # unbeteiligten Prozess vergeben - auch an eine
                        # NEUE Excel-Sitzung des Benutzers, die der
                        # Name-Check allein nicht erkennen wuerde.
                        try:
                            proc = psutil.Process(excel_pid)
                            if ("excel" in proc.name().lower()
                                    and (excel_create_time is None
                                         or proc.create_time()
                                            == excel_create_time)):
                                proc.kill()
                        except (psutil.NoSuchProcess, psutil.AccessDenied):
                            pass
                        except Exception as _e:
                            detail_logger.debug(f"watchdog: Exception verworfen: {_e!r}")

    wd_thread = threading.Thread(target=watchdog, daemon=True)
    wd_thread.start()

    result = [None]
    error  = [None]
    try:
        result[0] = op()
    except Exception as e:
        error[0] = e
    finally:
        with state_lock:
            done.set()

    if timeout_flag[0]:
        raise TimeoutError(
            f"{op_name}-Timeout nach {timeout}s – Excel-Prozess beendet")
    if error[0] is not None:
        raise error[0]
    return result[0]


def safe_excel_open(excel_app, file_path: str,
                    timeout: float = OPEN_TIMEOUT_SECONDS,
                    **open_kwargs):
    excel_pid = None
    try:
        _, excel_pid = win32process.GetWindowThreadProcessId(excel_app.Hwnd)
    except Exception as _e:
        detail_logger.debug(f"safe_excel_open: Exception verworfen: {_e!r}")
    if not excel_pid and excel_pid_global is not None:
        excel_pid = excel_pid_global
    excel_create_time = None
    if excel_pid and excel_pid == excel_pid_global:
        excel_create_time = excel_create_time_global
    elif excel_pid:
        try:
            excel_create_time = psutil.Process(excel_pid).create_time()
        except Exception as _e:
            detail_logger.debug(f"safe_excel_open: Exception verworfen: {_e!r}")

    return _run_com_with_watchdog(
        lambda: excel_app.Workbooks.Open(file_path, **open_kwargs),
        excel_pid,
        timeout,
        op_name="Excel-Open",
        excel_create_time=excel_create_time,
    )


def safe_excel_save(excel_app, save_func: Callable[[], Any],
                    timeout: float = SAVE_TIMEOUT_SECONDS) -> Any:
    excel_pid = None
    try:
        _, excel_pid = win32process.GetWindowThreadProcessId(excel_app.Hwnd)
    except Exception as _e:
        detail_logger.debug(f"safe_excel_save: Exception verworfen: {_e!r}")
    if not excel_pid and excel_pid_global is not None:
        excel_pid = excel_pid_global
    excel_create_time = None
    if excel_pid and excel_pid == excel_pid_global:
        excel_create_time = excel_create_time_global
    elif excel_pid:
        try:
            excel_create_time = psutil.Process(excel_pid).create_time()
        except Exception as _e:
            detail_logger.debug(f"safe_excel_save: Exception verworfen: {_e!r}")

    return _run_com_with_watchdog(
        save_func, excel_pid, timeout, op_name="Excel-Save",
        excel_create_time=excel_create_time)


# ==================================================================
# Trust-Center Smoke-Test
# ==================================================================
# Hintergrund:
#   Excel kann beim Open einer .xlsx in einen unsichtbaren Geschuetzte-
#   Ansicht-Dialog laufen, wenn der Pfad nicht als vertrauenswuerdiger
#   Speicherort eingetragen ist. In COM ist dieser Dialog nicht sichtbar -
#   der Aufruf haengt einfach. Pro Datei waeren das minutenlange Hänger.
#   Der Smoke-Test prueft daher VOR dem Hauptlauf, ob Open im Temp-Ordner
#   ohne Hänger funktioniert. Der Watchdog beendet Excel nach 25 s hart;
#   dann liefert der Test einen klaren Fehlertext statt eines Freezes.
#
#   Ergaenzt die Registry-basierte check_trust_center_registry() um den
#   tatsaechlichen Praxistest: Group Policies oder beschaedigte
#   Office-Profile koennen die Registry-Werte ignorieren.

def _build_minimal_xlsx(target_path: str) -> None:
    """Baut eine minimale gueltige .xlsx per zipfile (kein COM noetig)."""
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '</Types>'
    )
    rels_root = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    )
    rels_workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '</Relationships>'
    )
    sheet1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData/>'
        '</worksheet>'
    )

    if os.path.exists(target_path):
        try:
            os.remove(target_path)
        except Exception as _e:
            detail_logger.debug(f"_build_minimal_xlsx: Exception verworfen: {_e!r}")

    with zipfile.ZipFile(target_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",          content_types)
        zf.writestr("_rels/.rels",                  rels_root)
        zf.writestr("xl/workbook.xml",              workbook)
        zf.writestr("xl/_rels/workbook.xml.rels",   rels_workbook)
        zf.writestr("xl/worksheets/sheet1.xml",     sheet1)


def test_trust_center_smoke(timeout: float = 25.0) -> tuple:
    """
    Excel-Smoke-Test fuer 4b:
      Erzeugt eine kurzlebige Excel-Instanz, baut eine minimale Test-.xlsx
      im TEMP_BASE_PATH, oeffnet sie und schliesst sie wieder. Im
      Erfolgsfall wird Excel sauber beendet; bei Timeout per Watchdog hart
      gekillt.
      Rueckgabe: (ok, fehlertext, kategorie)
        kategorie: "ok" - Test erfolgreich
                   "com" - COM-Subsystem-Problem (z.B. CO_E_NOTINITIALIZED)
                   "xl"  - Excel/Trust-Center-Problem
    """
    try:
        os.makedirs(TEMP_BASE_PATH, exist_ok=True)
    except Exception as e:
        return False, f"Temp-Ordner nicht erstellbar ({TEMP_BASE_PATH}): {e}", "xl"

    test_path = os.path.join(
        TEMP_BASE_PATH, f"trustcheck_{uuid.uuid4().hex}.xlsx"
    )

    try:
        _build_minimal_xlsx(test_path)
    except Exception as e:
        return False, f"Test-XLSX konnte nicht erstellt werden: {e}", "xl"

    excel            = None
    test_pid         = None
    test_create_time = None
    wb               = None
    pythoncom.CoInitialize()  # COM-Init fuer Main-Thread (sonst CO_E_NOTINITIALIZED bei DispatchEx)
    try:
        # Eigene kurzlebige Excel-Instanz fuer den Test.
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except pythoncom.com_error as e:
            # HRESULT 0x800401F0 (-2147221008) = CO_E_NOTINITIALIZED -> COM-Subsystem, kein Trust-Center
            hresult = e.args[0] if e.args else None
            if hresult == -2147221008:
                return False, (f"COM-Subsystem nicht initialisiert "
                               f"(HRESULT 0x800401F0 CO_E_NOTINITIALIZED): {e}"), "com"
            return False, f"Excel-Instanz konnte nicht gestartet werden: {e}", "xl"
        except Exception as e:
            return False, f"Excel-Instanz konnte nicht gestartet werden: {e}", "xl"

        # PID via Hwnd (fuer Watchdog-Kill bei Timeout).
        try:
            _, test_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        except Exception:
            test_pid = None
        if test_pid:
            try:
                test_create_time = psutil.Process(test_pid).create_time()
            except Exception:
                test_create_time = None

        # Excel lautlos konfigurieren.
        try: excel.Visible            = COM_FALSE
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try: excel.DisplayAlerts      = COM_FALSE
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try: excel.ScreenUpdating     = COM_FALSE
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try: excel.EnableEvents       = COM_FALSE
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try: excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try: excel.AskToUpdateLinks   = False
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")

        # --- Open mit Watchdog ---
        try:
            wb = _run_com_with_watchdog(
                lambda: excel.Workbooks.Open(
                    test_path,
                    UpdateLinks=0,
                    ReadOnly=True,
                    IgnoreReadOnlyRecommended=True,
                    Notify=False,
                ),
                test_pid,
                timeout,
                op_name="Smoke-Test (Workbooks.Open)",
                excel_create_time=test_create_time,
            )
        except TimeoutError:
            # Watchdog hat Excel beendet; excel-Referenz ist tot.
            excel = None
            return False, (f"TIMEOUT nach {timeout:.0f}s bei Workbooks.Open "
                           "- Trust Center vermutlich nicht konfiguriert "
                           "oder Geschuetzte Ansicht aktiv"), "xl"
        except Exception as e:
            return False, f"Workbooks.Open fehlgeschlagen: {e}", "xl"

        if wb is None:
            return False, "Workbooks.Open lieferte kein Workbook-Objekt", "xl"

        return True, "", "ok"
    finally:
        if wb is not None:
            try: wb.Close(SaveChanges=False)
            except Exception as _e:
                detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        if excel is not None:
            try: excel.Quit()
            except Exception as _e:
                detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
            # Falls Quit nicht greift: harter Kill ueber gemerkte PID.
            if test_pid:
                try:
                    p = psutil.Process(test_pid)
                    if (p.is_running()
                            and p.name().lower() == "excel.exe"
                            and (test_create_time is None
                                 or p.create_time() == test_create_time)):
                        p.kill()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
                except Exception as _e:
                    detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try:
            if os.path.exists(test_path):
                os.remove(test_path)
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try:
            pythoncom.CoUninitialize()
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")


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
    return prepare_long_path(path) if len(path) > MAX_PATH_LEN else path


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
            detail_logger.debug(
                "Restore-Privilegien aktiviert (Admin-Kontext) – "
                "Owner-Wiederherstellung vollstaendig verfuegbar.")
        else:
            detail_logger.debug(
                "Restore-Privilegien nicht zugewiesen (Nutzer-Kontext) – "
                "Owner-Restore nur auf eigene Dateien moeglich.")
    except Exception as e:
        detail_logger.debug(f"Privileg-Aktivierung fehlgeschlagen: {e}")


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
        detail_logger.debug(f"Sicherheitsinfo nicht lesbar ({path}): {e}")
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
            detail_logger.debug(f"_apply_security_descriptor: Exception verworfen: {_e!r}")

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
            detail_logger.debug(f"_apply_security_descriptor: Owner-Vergleich verworfen: {_e!r}")

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
        detail_logger.debug(f"Sicherheitsinfo wiederhergestellt: {path}")
        return
    except Exception as e:
        if owner is None and group is None:
            detail_logger.warning(f"DACL-Wiederherstellung fehlgeschlagen ({path}): {e}")
            return
        if _restore_privileges_enabled:
            detail_logger.warning(f"Owner-Wiederherstellung fehlgeschlagen ({path}): {e}")
        else:
            detail_logger.debug(f"Owner nicht gesetzt (kein Admin-Privileg - im Nutzer-Kontext unkritisch): {path} - {e}")

    # Rueckfall: wenigstens die DACL setzen.
    if dacl is not None and nur_dacl:
        try:
            win32security.SetNamedSecurityInfo(
                p, win32security.SE_FILE_OBJECT, nur_dacl, None, None, dacl, None)
            detail_logger.debug(f"DACL wiederhergestellt (ohne Owner): {path}")
        except Exception as e2:
            detail_logger.warning(f"DACL-Wiederherstellung fehlgeschlagen ({path}): {e2}")


# ==================================================================
# Pfad-Hilfsfunktionen
# ==================================================================
def sanitize_path(raw: str) -> str:
    path = raw.strip().strip('"').strip("'")
    path = os.path.expandvars(path)
    path = os.path.expanduser(path)
    if len(path) == 2 and path[1] == ":":
        path = path + "\\"
    return path



def _utime_rueckfall(pfad: str, zeiten) -> bool:
    """Rueckfall, wenn win32file.SetFileTime scheitert.

    Erhaelt Zugriffs- und Aenderungszeit - nicht die Erstellungszeit,
    aber das ist deutlich besser als der vollstaendige Verlust des
    Datums. Diesen Rueckfall hatte bisher nur 5_OCR_PDF.py; ohne ihn
    verloren die uebrigen Skripte die Zeitstempel stillschweigend,
    sobald pywin32 fehlte oder der Handle nicht zu oeffnen war.
    """
    try:
        zugriff, geaendert = zeiten[1], zeiten[2]
        os.utime(prepare_long_path(pfad),
                 (zugriff.timestamp(), geaendert.timestamp()))
        return True
    except Exception:
        return False

def prepare_long_path(path: str) -> str:
    if path.startswith("\\\\?\\"):
        return path
    path = os.path.abspath(os.path.normpath(path))
    if path.startswith("\\\\"):
        return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path

# ==================================================================
# OLE-Verschlüsselungserkennung
# ==================================================================
def _is_xls_encrypted(file_path: str) -> bool:
    # Klassische BIFF-Verschluesselung (.xls, FilePass-Record) traegt keine
    # OOXML-Marker. Falls msoffcrypto verfuegbar ist, zuerst damit pruefen –
    # es erkennt das Legacy-Format zuverlaessig; sonst manueller Marker-Scan.
    if _msoffcrypto is not None:
        try:
            with open(prepare_long_path(file_path), "rb") as f:
                if _msoffcrypto.OfficeFile(f).is_encrypted():
                    return True
        except Exception as e_crypto:
            detail_logger.debug(f"msoffcrypto-Check (.xls) fehlgeschlagen: {e_crypto}")
    try:
        with open(prepare_long_path(file_path), "rb") as f:
            data = f.read(65536)
        if len(data) < 8:
            return False
        if not (data[0] == 0xD0 and data[1] == 0xCF and
                data[2] == 0x11 and data[3] == 0xE0):
            return False

        content_utf16 = data.decode("utf-16-le", errors="ignore")
        content_latin = data.decode("latin-1",   errors="ignore")
        markers = ("EncryptedPackage", "EncryptionInfo",
                   "EncryptedSummary", "DataSpaces")
        for m in markers:
            if m in content_utf16 or m in content_latin:
                return True

        if b"\x00\x00\x2F\x00" in data and b"\x00E\x00n\x00c\x00r\x00y\x00p\x00t" in data:
            return True
        return False
    except Exception:
        return False

def _xls_hat_schreibkennwort(lp: str) -> bool:
    """BIFF8: FILESHARING-Satz (0x005B) mit Kennwort-Hash im Globals-Teil."""
    try:
        import olefile   # kommt mit msoffcrypto; fehlt es, keine Erkennung
    except ImportError:
        return False
    import struct
    try:
        with olefile.OleFileIO(lp) as ole:
            name = "Workbook" if ole.exists("Workbook") else "Book"
            if not ole.exists(name):
                return False
            daten = ole.openstream(name).read()
    except Exception as _e:
        detail_logger.debug(f"_xls_hat_schreibkennwort: {_e!r}")
        return False
    pos = 0
    while pos + 4 <= len(daten):
        typ, laenge = struct.unpack_from("<HH", daten, pos)
        if typ == 0x005B and laenge >= 4:
            return struct.unpack_from("<H", daten, pos + 6)[0] != 0
        if typ == 0x000A:                    # EOF des Globals-Teils
            return False
        pos += 4 + laenge
    return False


def _xlsb_hat_schreibkennwort(wb_bin: bytes) -> bool:
    """BIFF12 (xl/workbook.bin): Schreibreservierung mit Kennwort?

    Gemessen am 29.09.2026 mit Excel 2024: mit Kennwort schreibt Excel den
    Satz BrtFileSharingIso (0x2A4, neuer Hash) und daneben BrtFileSharing
    (0x224) ohne Hash; bei "schreibgeschuetzt empfohlen" ohne Kennwort steht
    nur 0x224 mit leerem Hash. Aeltere Dateien tragen den Hash (wResPass) in
    0x224 selbst.
    """
    pos = 0
    n = len(wb_bin)
    while pos < n:
        b0 = wb_bin[pos]; pos += 1
        typ = b0 & 0x7F
        if b0 & 0x80 and pos < n:
            typ |= (wb_bin[pos] & 0x7F) << 7; pos += 1
        laenge = 0
        for i in range(4):
            if pos >= n:
                return False
            b = wb_bin[pos]; pos += 1
            laenge |= (b & 0x7F) << (7 * i)
            if not b & 0x80:
                break
        if typ == 0x2A4:
            return True
        if typ == 0x224 and laenge >= 4:
            return int.from_bytes(wb_bin[pos + 2:pos + 4], "little") != 0
        if typ == 0x08F:               # BrtBeginBundleShs: Kopfteil vorbei
            return False
        pos += laenge
    return False


def _excel_kennwort_grund(path: str) -> str:
    """'' wenn Excel die Mappe ohne Kennwortabfrage oeffnet, sonst den Grund.

    Gemessen am 29.09.2026 mit Excel 2024: eine verschluesselte .xlsx
    (Oeffnungskennwort) und eine mit Schreibreservierungs-Kennwort hingen
    bis zum Waechter (je 60 s plus Excel-Neustart). 4a erkennt beides seit
    demselben Tag vorab; hier fehlte es, nur verschluesselte .xls wurden
    erkannt.
    - Verschluesselt: OOXML-Datei im CFB-Container statt ZIP; .xls ueber
      _is_xls_encrypted.
    - Schreibkennwort: <fileSharing> mit Kennwort-Hash in xl/workbook.xml
      bzw. FILESHARING-Satz in der .xls (die reine Empfehlung
      "schreibgeschuetzt oeffnen" ohne Kennwort bleibt erlaubt).
    """
    ext = os.path.splitext(path)[1].lower()
    try:
        lp = prepare_long_path(path)
        if ext in (".xls", ".xlt"):
            if _is_xls_encrypted(path):
                return "verschluesselt"
            return "schreibkennwort" if _xls_hat_schreibkennwort(lp) else ""
        with open(lp, "rb") as f:
            kopf = f.read(8)
        if kopf.startswith(b"\xD0\xCF\x11\xE0"):
            return "verschluesselt"
        if not kopf.startswith(b"PK"):
            return ""
        with zipfile.ZipFile(lp) as z:
            try:
                wb = z.read("xl/workbook.xml").decode("utf-8", errors="ignore")
            except KeyError:
                try:
                    wb_bin = z.read("xl/workbook.bin")          # .xlsb
                except KeyError:
                    return ""
                return "schreibkennwort" if _xlsb_hat_schreibkennwort(wb_bin) else ""
        m = re.search(r"<(?:\w+:)?fileSharing\b[^>]*>", wb)
        if m and re.search(r"\b(reservationPassword|hashValue|algorithmName)=", m.group(0)):
            return "schreibkennwort"
    except Exception as _e:
        detail_logger.debug(f"_excel_kennwort_grund: {_e!r}")
    return ""


# ==================================================================
# Dateiattribute
# ==================================================================
# Bis 29.09.2026 gingen Schreibschutz/Versteckt/System verloren: der
# Schreibschutz wurde zum Speichern entfernt und nie zurueckgesetzt, und eine
# umgewandelte .xls-Mappe entsteht per SaveAs ohnehin mit Normalattributen.
# 3b erhaelt sie (gleiche Maske).
_SETZBARE_ATTRIBUTE = (
    0x0001      # READONLY
    | 0x0002    # HIDDEN
    | 0x0004    # SYSTEM
    | 0x0020    # ARCHIVE
    | 0x2000    # NOT_CONTENT_INDEXED
)


def _dateiattribute_lesen(pfad: str):
    try:
        return win32file.GetFileAttributesW(prepare_long_path(pfad))
    except Exception as _e:
        detail_logger.debug(f"_dateiattribute_lesen: {_e!r}")
        return None


def _dateiattribute_zurueck(pfad: str, attribute) -> None:
    """Schreibt gemerkte Dateiattribute zurueck (best effort, protokolliert)."""
    if attribute is None or not pfad or not os.path.exists(prepare_long_path(pfad)):
        return
    wert = attribute & _SETZBARE_ATTRIBUTE
    if not wert:
        wert = 0x0080           # NORMAL
    try:
        win32file.SetFileAttributesW(prepare_long_path(pfad), wert)
        detail_logger.debug(f"Dateiattribute wiederhergestellt (0x{wert:X}): {pfad}")
    except Exception as _e:
        detail_logger.warning(
            f"Dateiattribute (0x{wert:X}) nicht wiederherstellbar: {pfad} – {_e!r}")


# ==================================================================
# Integritätsprüfung
# ==================================================================
def verify_saved_file(path: str) -> bool:
    try:
        lp = prepare_long_path(path)
        if not os.path.exists(lp):
            return False
        if os.path.getsize(lp) < 100:
            return False
        ext = os.path.splitext(path)[1].lower()
        if ext in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb"):
            try:
                return zipfile.is_zipfile(lp)
            except Exception:
                return False
        return True
    except Exception:
        return False

# ==================================================================
# Interaktiver Start
# ==================================================================
def _is_tty() -> bool:
    try:
        return sys.stdin.isatty()
    except Exception:
        return False


def ask_directory() -> str:
    desktop_path   = _resolve_user_shell_folder("Desktop", "Desktop")
    downloads_path = _resolve_user_shell_folder(
        "{374DE290-123F-4565-9164-39C4925E467B}", "Downloads")

    print("\nZielverzeichnis auswählen:")
    print("  [1] Q:\\")
    print("  [2] R:\\")
    print("  [3] G:\\Geteilte Ablagen")
    print("  [4] G:\\Meine Ablage")
    print("  [5] \\\\server\\dfs\\dm")
    print(f"  [6] Desktop      ({desktop_path})")
    print(f"  [7] Downloads    ({downloads_path})")
    print("  [8] Eigenen Pfad eingeben")
    print()
    print("  Hinweis: Bei Start 'als Administrator' sind Netzlaufwerks-")
    print("  Buchstaben (Q:, R:) oft nicht verbunden – dann Option [5]")
    print("  (UNC-Pfad) verwenden.")
    print()
    while True:
        try:
            choice = input("Auswahl [1-8]: ").strip()
        except EOFError:
            print("\nKeine Eingabe möglich (EOF) – Abbruch.")
            sys.exit(1)
        if choice == "1":
            path = "Q:\\"
        elif choice == "2":
            path = "R:\\"
        elif choice == "3":
            path = "G:\\Geteilte Ablagen"
        elif choice == "4":
            path = "G:\\Meine Ablage"
        elif choice == "5":
            path = "\\\\server\\dfs\\dm"
        elif choice == "6":
            path = desktop_path
        elif choice == "7":
            path = downloads_path
        elif choice == "8":
            try:
                raw = input("Pfad eingeben: ")
            except EOFError:
                print("\nKeine Eingabe möglich (EOF) – Abbruch.")
                sys.exit(1)
            path = sanitize_path(raw)
        else:
            print("  Bitte 1-8 eingeben.")
            continue

        if os.path.isdir(path):
            return path

        if choice in ("1", "2"):
            print(f"  Laufwerk {path} nicht erreichbar – ist es gemappt?")
        elif choice in ("3", "4"):
            print(f"  Pfad nicht erreichbar: {path}")
            print("     (Google Drive gestartet? Laufwerk G:\\ verfügbar?)")
        elif choice == "5":
            print(f"  Netzwerkpfad nicht erreichbar: {path}")
            print("     (VPN aktiv? Netzwerkverbindung prüfen)")
        elif choice in ("6", "7"):
            print(f"  Pfad nicht gefunden: {path}")
        else:
            print(f"  Verzeichnis nicht gefunden: '{path}'")


def ask_font() -> str:
    try:
        val = input(f"\nZiel-Schriftart [Standard: {NEW_FONT_NAME}]: ").strip()
    except EOFError:
        return NEW_FONT_NAME
    return val if val else NEW_FONT_NAME


def ask_yes_no(prompt: str, default_yes: bool = True) -> bool:
    hint = "J/n" if default_yes else "j/N"
    while True:
        try:
            answer = input(f"{prompt} [{hint}]: ").strip().lower()
        except EOFError:
            return default_yes
        if not answer:
            return default_yes
        if answer in ("j", "ja", "y", "yes"):
            return True
        if answer in ("n", "nein", "no"):
            return False
        print("  Bitte 'j' oder 'n' eingeben.")


def ask_progress_mode() -> tuple:
    """
    Returns (show_progress: bool, count_first: bool).
      [1] Standard – kein tqdm-Balken, jede Datei wird inline geloggt.
      [2] tqdm mit ETA – benoetigt Vorab-Scan aller Dateien.
      [3] tqdm ohne ETA – sofortiger Start, kein Total.
    """
    print()
    print("Fortschrittsanzeige:")
    print("  [1] Keine Fortschrittsleiste (laufende Zählung inline) [Standard]")
    print("  [2] Fortschrittsleiste mit ETA (Vorab-Scan nötig, Start verzögert!)")
    print("  [3] Fortschrittsleiste ohne ETA (sofortiger Start, kein Total)")
    print()
    print("  Hinweis zu [2]: Bei UNC-Pfaden oder sehr großen Freigaben kann")
    print("  die Indizierung 10–30 Minuten dauern, bevor die erste Datei")
    print("  verarbeitet wird. Bei > 100.000 Dateien besser [1] oder [3]")
    print("  wählen (Generator-Variante ohne Vorab-Zählung).")
    print()
    while True:
        try:
            choice = input("Auswahl [1-3]: ").strip()
        except EOFError:
            return (False, False)
        if choice in ("", "1"):
            return (False, False)
        if choice == "2":
            return (True, True)
        if choice == "3":
            return (True, False)
        print("  Bitte 1, 2 oder 3 eingeben.")


def ask_metadata_detail() -> dict:
    print()
    print("  Welche Metadaten sollen entfernt werden?")
    print("  (Jede Option wird einzeln abgefragt)")
    print()

    selected = {}

    print("  ── Metadaten ohne Inhaltsverlust ──────────────────────────")
    for key in ["props", "custom", "personal"]:
        selected[key] = ask_yes_no(f"  {METADATA_OPTIONS[key]} entfernen?")

    print()
    print("  ──  Folgende Optionen entfernen ZELLINHALTE  ──────────────")
    for key in ["comments", "threads"]:
        selected[key] = ask_yes_no(f"  {METADATA_OPTIONS[key]} entfernen?")

    return selected

# ==================================================================
# Datei-Generator
# ==================================================================
def _is_reparse_point(entry_or_path) -> bool:
    # Junctions/Mount-Points erkennen: entry.is_symlink() ist unter Windows
    # nur fuer echte Symlinks True; Junctions gelten als normale Ordner ->
    # Zyklus-Gefahr (Endlosschleife).
    try:
        if hasattr(entry_or_path, "stat"):
            attrs = entry_or_path.stat(follow_symlinks=False).st_file_attributes
        else:
            attrs = os.stat(entry_or_path, follow_symlinks=False).st_file_attributes
        return bool(attrs & FILE_ATTRIBUTE_REPARSE_POINT)
    except (OSError, AttributeError):
        return False



def is_locked_by_other(path: str) -> bool:
    """Erkennt die Office-Owner-Datei (~$...) neben der Zieldatei.

    Excel haengt "~$" immer vor den vollen Namen (anders als Word, das je
    nach Namenslaenge die ersten Zeichen ersetzt) - eine Variante genuegt.
    """
    try:
        d = os.path.dirname(path)
        b = os.path.basename(path)
        if not d or not b:
            return False
        try:
            return os.path.exists(prepare_long_path(os.path.join(d, "~$" + b)))
        except Exception:
            return False
    except Exception:
        return False


def _verwaiste_besitzerdatei_entfernen(path: str) -> None:
    """Entfernt die Office-Besitzerdatei '~$<name>' neben path, die eine
    getoetete Excel-Instanz hinterlassen hat. Ohne das meldete
    is_locked_by_other die Mappe in Versuch 2 und in jedem Folgelauf als
    'in Excel geöffnet'. Eine von einer lebenden Excel-Sitzung offen-
    gehaltene Besitzerdatei laesst sich nicht loeschen (Freigabeverletzung)
    und bleibt unangetastet. Kein _av_safe_remove: dessen Wiederholungen
    wuerden eine echte fremde Sperre nur abwarten."""
    try:
        d = os.path.dirname(path)
        b = os.path.basename(path)
        if not d or not b:
            return
        besitzer = prepare_long_path(os.path.join(d, "~$" + b))
        if not os.path.exists(besitzer):
            return
        try:
            os.chmod(besitzer, stat.S_IWRITE)
        except Exception as _e:
            detail_logger.debug(f"_verwaiste_besitzerdatei_entfernen: {_e!r}")
        os.remove(besitzer)
        detail_logger.warning(f"Verwaiste Besitzerdatei entfernt: {os.path.join(d, '~$' + b)}")
    except Exception as e:
        detail_logger.warning(
            f"Besitzerdatei nicht entfernbar (von Excel belegt?): "
            f"{os.path.join(os.path.dirname(path), '~$' + os.path.basename(path))} – {e}")


def _excel_antwortet(excel_app) -> bool:
    try:
        _ = excel_app.Version
        return True
    except Exception:
        return False


def dry_run_directory(directory: str) -> dict:
    """Listet auf, was der Echtlauf tun WUERDE - insbesondere welche
    .xls/.xlt konvertiert (und deren Originale ersetzt) wuerden.
    Es wird nichts geoeffnet, gespeichert oder geloescht; Excel wird
    nicht gestartet. Identisch aufgebaut zu 4a/4c."""
    stats = {"VERARBEITEN": 0, "KONVERTIEREN": 0,
             "GESPERRT": 0, "VERSCHLUESSELT": 0}
    print("\nPROBELAUF – es wird nichts geändert.\n")
    print("-" * 66)
    for file_path in file_generator(directory):
        ext = os.path.splitext(file_path)[1].lower()
        if is_locked_by_other(file_path):
            stats["GESPERRT"] += 1
            tag = "GESPERRT     "
        elif _excel_kennwort_grund(file_path):
            stats["VERSCHLUESSELT"] += 1
            tag = "KENNWORT     "
        elif ext in (".xls", ".xlt"):
            stats["KONVERTIEREN"] += 1
            tag = "KONVERTIEREN "
        else:
            stats["VERARBEITEN"] += 1
            tag = "VERARBEITEN  "
        print(f"  [{tag}] {file_path}")
        detail_logger.info(f"Probelauf [{tag.strip()}]: {file_path}")
    print("-" * 66)
    print("\n  PROBELAUF-ERGEBNIS (es wurde nichts geändert):")
    print(f"    Würden verarbeitet (Font/Metadaten):     {stats['VERARBEITEN']}")
    print( "    Würden konvertiert (.xls/.xlt → neu,")
    print(f"                        Original ersetzt!):  {stats['KONVERTIEREN']}")
    print(f"    Gesperrt (würden übersprungen):          {stats['GESPERRT']}")
    print(f"    Kennwortgeschützt (würden übersprungen): {stats['VERSCHLUESSELT']}")
    print(f"    GESAMT:                                  {sum(stats.values())}")
    return stats


def file_generator(directory: str):
    extensions = {".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xls", ".xlt"}
    # Eigenen Arbeitsordner ausnehmen: liegt das Ziel z.B. auf Documents,
    # wuerde der Generator sonst die eigenen Stage-/RESCUE-Dateien
    # einsammeln, waehrend Excel sie gerade schreibt.
    temp_nc = os.path.normcase(os.path.abspath(TEMP_BASE_PATH))
    queue = [os.path.abspath(directory)]
    while queue:
        current_dir = queue.pop()
        try:
            with os.scandir(current_dir) as it:
                entries = list(it)
        except (PermissionError, OSError):
            continue
        for entry in entries:
            try:
                if entry.is_symlink():
                    continue
                if entry.is_dir(follow_symlinks=False):
                    # Papierkorb/Systemordner/NAS-Snapshots, eigenen Temp-
                    # Ordner und Junctions nicht betreten.
                    if entry.name.lower() in EXCLUDE_DIR_NAMES:
                        continue
                    d_nc = os.path.normcase(os.path.abspath(entry.path))
                    if d_nc == temp_nc or d_nc.startswith(temp_nc + os.sep):
                        continue
                    if _is_reparse_point(entry):
                        continue
                    queue.append(entry.path)
                elif entry.is_file(follow_symlinks=False):
                    nl = entry.name.lower()
                    if any(nl.startswith(p) for p in SKIP_FILE_PREFIXES):
                        continue
                    if any(nl.endswith(s) for s in SKIP_FILE_SUFFIXES):
                        continue
                    if any(nl.endswith(ext) for ext in extensions):
                        yield entry.path
            except (PermissionError, OSError):
                continue

# ==================================================================
# Schrift setzen, Symbolschriften schonen
# ==================================================================
# Bis 29.09.2026 setzte 4b die Zielschrift pauschal (UsedRange, Formatvor-
# lagen, Textfelder, Kopfzeilen) - auch ueber Wingdings/Symbol. Aus einem
# Haekchen wurde dabei ein "ü". Jetzt bleibt Text in Symbolschrift stehen.

# Obergrenze fuer die Halbierung je Blatt (COM-Aufrufe). Wird sie gerissen,
# setzt _zellschrift_setzen den Rest pauschal und holt reine Symbolzellen
# danach zurueck (siehe dort).
ZELLSCHRIFT_BUDGET   = 20000
# Laengere Mischzellen werden nicht zeichenweise gelesen (ein Aufruf je
# Zeichen); sie behalten ihre Schriften.
ZEICHENWEISE_MAX_LEN = 2000


def _com_eigenschaft(obj, name: str, *args):
    """Eigenschaft MIT Argumenten (Characters, Runs) in spaeter Bindung.

    Gemessen am 29.09.2026: Range.Characters(1, 3) wirft "Mitglied nicht
    gefunden", TextRange2.Runs(1) "Auflistung nicht unterstuetzt" - pywin32
    holt erst die Eigenschaft ohne Argumente und ruft dann deren
    Standardmitglied auf. Die Get-Form (GetCharacters/GetRuns) reicht die
    Argumente richtig durch.
    """
    try:
        methode = getattr(obj, "Get" + name)
    except AttributeError:
        methode = getattr(obj, name)
    return methode(*args)


def _font_name_setzen(font, font_name: str) -> None:
    """Setzt font.Name, ausser die bisherige Schrift ist eine Symbolschrift."""
    try:
        if _ist_symbolschrift(font.Name):
            return
    except Exception as _e:
        detail_logger.debug(f"_font_name_setzen: Exception verworfen: {_e!r}")
    font.Name = font_name


def _textrange2_setzen(tr, font_name: str) -> None:
    """TextRange2 (Textfeld, SmartArt): Symbol-Abschnitte bleiben stehen.

    Gemessen am 29.09.2026: bei gemischten Schriften liefert
    TextRange2.Font.Name einen leeren Namen - dann wird abschnittsweise
    (Runs) gesetzt.
    """
    name = tr.Font.Name
    if name:
        if not _ist_symbolschrift(name):
            tr.Font.Name = font_name
        return
    for i in range(1, tr.Runs.Count + 1):
        try:
            run = _com_eigenschaft(tr, "Runs", i, 1)
            if not _ist_symbolschrift(run.Font.Name):
                run.Font.Name = font_name
        except Exception as _e:
            detail_logger.debug(f"_textrange2_setzen: Exception verworfen: {_e!r}")


def _zelle_zeichenweise(zelle, font_name: str) -> None:
    """Mischzelle (Rich Text): nur Abschnitte ohne Symbolschrift umstellen."""
    try:
        wert = zelle.Value2
        if not isinstance(wert, str) or not wert or len(wert) > ZEICHENWEISE_MAX_LEN:
            return
        namen = [_com_eigenschaft(zelle, "Characters", i, 1).Font.Name
                 for i in range(1, len(wert) + 1)]
    except Exception as _e:
        detail_logger.debug(f"_zelle_zeichenweise: Exception verworfen: {_e!r}")
        return
    start = 0
    while start < len(namen):
        ende = start
        while ende + 1 < len(namen) and namen[ende + 1] == namen[start]:
            ende += 1
        if namen[start] != font_name and not _ist_symbolschrift(namen[start]):
            try:
                _com_eigenschaft(zelle, "Characters", start + 1,
                                 ende - start + 1).Font.Name = font_name
            except Exception as _e:
                detail_logger.debug(f"_zelle_zeichenweise: Exception verworfen: {_e!r}")
        start = ende + 1


def _symbolzellen_finden(rng) -> list:
    """(Adresse, Schrift) der Zellen, die GANZ in einer Symbolschrift stehen.

    Gemessen am 29.09.2026: Range.Find mit SearchFormat findet so nur Zellen
    mit einheitlicher Zellschrift, keine Mischzellen. FindNext ignoriert
    SearchFormat, deshalb wird Find mit After= wiederholt.
    """
    treffer = []
    try:
        xl = rng.Application
    except Exception:
        return treffer
    try:
        for name in _symbolschrift_namen():
            try:
                xl.FindFormat.Clear()
                xl.FindFormat.Font.Name = name
                letzte = rng.Cells(rng.Cells.Count)
                erste = rng.Find("", letzte, -4123, 2, 1, 1, False, False, True)
                zelle = erste
                while zelle is not None and len(treffer) < ZELLSCHRIFT_BUDGET:
                    treffer.append((zelle.Address, zelle.Font.Name))
                    zelle = rng.Find("", zelle, -4123, 2, 1, 1, False, False, True)
                    if zelle is not None and zelle.Address == erste.Address:
                        break
            except Exception as _e:
                detail_logger.debug(f"_symbolzellen_finden: Exception verworfen: {_e!r}")
    finally:
        try:
            xl.FindFormat.Clear()
        except Exception as _e:
            detail_logger.debug(f"_symbolzellen_finden: Exception verworfen: {_e!r}")
    return treffer


def _zellschrift_setzen(sheet, rng, font_name: str) -> bool:
    """Setzt die Zellschrift in rng und schont Symbolschriften.

    Ein einheitlicher Bereich kostet einen Aufruf; ist er gemischt, wird er
    halbiert, bis jeder Teil einheitlich ist. Einzelne Mischzellen werden
    zeichenweise bearbeitet. Reisst die Halbierung ZELLSCHRIFT_BUDGET, wird
    der Rest pauschal gesetzt und reine Symbolzellen danach zurueckgeholt -
    dann koennen Symbolzeichen INNERHALB von Mischzellen umgestellt sein.
    Rueckgabe: False in diesem Rueckfall.
    """
    stapel = [rng]
    budget = ZELLSCHRIFT_BUDGET
    while stapel:
        r = stapel.pop()
        budget -= 1
        if budget < 0:
            stapel.append(r)
            break
        name = r.Font.Name
        if name:
            if name != font_name and not _ist_symbolschrift(name):
                r.Font.Name = font_name
            continue
        zeilen = r.Rows.Count
        spalten = r.Columns.Count
        if zeilen == 1 and spalten == 1:
            _zelle_zeichenweise(r, font_name)
            continue
        if zeilen >= spalten:
            h = zeilen // 2
            stapel.append(sheet.Range(r.Cells(1, 1), r.Cells(h, spalten)))
            stapel.append(sheet.Range(r.Cells(h + 1, 1), r.Cells(zeilen, spalten)))
        else:
            h = spalten // 2
            stapel.append(sheet.Range(r.Cells(1, 1), r.Cells(zeilen, h)))
            stapel.append(sheet.Range(r.Cells(1, h + 1), r.Cells(zeilen, spalten)))
    if not stapel:
        return True
    for r in stapel:
        gerettet = _symbolzellen_finden(r)
        r.Font.Name = font_name
        for adresse, alt in gerettet:
            try:
                sheet.Range(adresse).Font.Name = alt
            except Exception as _e:
                detail_logger.debug(f"_zellschrift_setzen: Exception verworfen: {_e!r}")
    return False


# ==================================================================
# Diagramm-Schriftarten
# ==================================================================
def _set_chart_fonts(chart, font_name: str, depth: int = 0) -> None:
    try:
        _textrange2_setzen(chart.ChartArea.Format.TextFrame2.TextRange, font_name)
    except Exception as _e:
        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")

    try:
        if chart.HasTitle:
            _font_name_setzen(chart.ChartTitle.Font, font_name)
    except Exception as _e:
        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")

    try:
        for axis in chart.Axes():
            try:
                _font_name_setzen(axis.TickLabels.Font, font_name)
            except Exception as _e:
                detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
            try:
                if axis.HasTitle:
                    _font_name_setzen(axis.AxisTitle.Font, font_name)
            except Exception as _e:
                detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")

    try:
        if chart.HasLegend:
            _font_name_setzen(chart.Legend.Font, font_name)
    except Exception as _e:
        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")

    try:
        if chart.HasDataTable:
            try:
                _font_name_setzen(chart.DataTable.Font, font_name)
            except Exception as _e:
                detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")

    try:
        for series in chart.SeriesCollection():
            try:
                if series.HasDataLabels:
                    try:
                        _font_name_setzen(series.DataLabels().Font, font_name)
                    except Exception as _e:
                        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
                    try:
                        pts_coll = series.Points()
                        pt_count = pts_coll.Count
                        if 0 < pt_count <= MAX_POINTS_PER_SERIES:
                            for j in range(1, pt_count + 1):
                                try:
                                    pt = pts_coll.Item(j)
                                    if pt.HasDataLabel:
                                        _font_name_setzen(pt.DataLabel.Font, font_name)
                                except Exception as _e:
                                    detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
                    except Exception as _e:
                        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
            except Exception as _e:
                detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
            try:
                for tl in series.Trendlines():
                    try:
                        if tl.DisplayEquation or tl.DisplayRSquared:
                            _font_name_setzen(tl.DataLabel.Font, font_name)
                    except Exception as _e:
                        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
            except Exception as _e:
                detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")

    try:
        _textrange2_setzen(chart.PlotArea.Format.TextFrame2.TextRange, font_name)
    except Exception as _e:
        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")

    try:
        for shp in chart.Shapes:
            try:
                _process_sheet_shape(shp, font_name, depth + 1)
            except Exception as _e:
                detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")


def _process_sheet_shape(shape, font_name: str, depth: int = 0) -> None:
    if depth > MAX_SHAPE_RECURSION_DEPTH:
        detail_logger.warning(
            f"Shape-Rekursionstiefe {depth} überschritten – Abbruch")
        return

    try:
        shape_type = shape.Type
    except Exception:
        return

    if shape_type == MSO_GROUP:
        try:
            for item in shape.GroupItems:
                try:
                    _process_sheet_shape(item, font_name, depth + 1)
                except Exception as _e:
                    detail_logger.debug(f"_process_sheet_shape: Exception verworfen: {_e!r}")
        except Exception as _e:
            detail_logger.debug(f"_process_sheet_shape: Exception verworfen: {_e!r}")
        return

    if shape_type == MSO_CHART:
        try:
            _set_chart_fonts(shape.Chart, font_name, depth)
        except Exception as _e:
            detail_logger.debug(f"_process_sheet_shape: Exception verworfen: {_e!r}")
        return

    # Excel-Formen kennen kein HasTextFrame (das ist PowerPoint): die
    # Abfrage warf bis 29.09.2026 immer AttributeError, und 4b liess jedes
    # Textfeld unveraendert. TextFrame2.HasText ist die Excel-Form.
    try:
        if shape.TextFrame2.HasText:
            try:
                _textrange2_setzen(shape.TextFrame2.TextRange, font_name)
            except Exception:
                try:
                    _font_name_setzen(shape.TextFrame.Characters().Font, font_name)
                except Exception as _e:
                    detail_logger.debug(f"_process_sheet_shape: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_process_sheet_shape: Exception verworfen: {_e!r}")

    try:
        if getattr(shape, "HasSmartArt", False):
            for node in shape.SmartArt.AllNodes:
                try:
                    _textrange2_setzen(node.TextFrame2.TextRange, font_name)
                except Exception as _e:
                    detail_logger.debug(f"_process_sheet_shape: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_process_sheet_shape: Exception verworfen: {_e!r}")


_HEADER_FONT_RE = re.compile(r'&"([^",]+)(?:,([^"]+))?"')


def _replace_header_footer_fonts(sheet, font_name: str) -> None:
    try:
        ps = sheet.PageSetup
    except Exception:
        return

    hf_attrs = (
        "LeftHeader", "CenterHeader", "RightHeader",
        "LeftFooter", "CenterFooter", "RightFooter",
        "FirstPageLeftHeader", "FirstPageCenterHeader", "FirstPageRightHeader",
        "FirstPageLeftFooter", "FirstPageCenterFooter", "FirstPageRightFooter",
        "EvenPageLeftHeader", "EvenPageCenterHeader", "EvenPageRightHeader",
        "EvenPageLeftFooter", "EvenPageCenterFooter", "EvenPageRightFooter",
    )
    for attr in hf_attrs:
        try:
            value = getattr(ps, attr, None)
            if not value:
                continue

            if '&"' in value:
                new_value = _HEADER_FONT_RE.sub(
                    lambda m: m.group(0) if _ist_symbolschrift(m.group(1))
                              else (f'&"{font_name},{m.group(2)}"' if m.group(2)
                                    else f'&"{font_name}"'),
                    value
                )
            else:
                new_value = f'&"{font_name}"' + value

            if new_value != value:
                try:
                    setattr(ps, attr, new_value)
                except Exception as _e:
                    detail_logger.debug(f"_replace_header_footer_fonts: Exception verworfen: {_e!r}")
        except Exception as _e:
            detail_logger.debug(f"_replace_header_footer_fonts: Exception verworfen: {_e!r}")

# ==================================================================
# Conditional Formatting / Tabellen / Pivot
# ==================================================================
def _hat_bedingte_formatierung(sheet) -> bool:
    """Traegt das Blatt bedingte Formatierung?

    Die Schriftart laesst sich dort NICHT umstellen: das Font-Objekt einer
    FormatCondition unterstuetzt nur Schriftschnitt, Unterstreichung, Farbe
    und Durchstreichung. Mit echtem Excel geprueft - 'Name' und 'Size' werfen
    'Die Name-Eigenschaft des Font-Objektes kann nicht festgelegt werden',
    waehrend Bold/Italic/Underline/Strikethrough anstandslos durchgehen. Es
    ist dieselbe Einschraenkung, die im Dialog 'Zellen formatieren' fuer
    bedingte Formatierung sichtbar ist (Feld 'Schriftart' ausgegraut); das
    zugrunde liegende dxf-Format kennt schlicht keinen Schriftnamen.

    Frueher versuchte _set_conditional_formatting_fonts genau das und
    verschluckte den Fehler still - das Blatt galt als vollstaendig
    umgestellt, obwohl bedingt formatierte Zellen die alte Schrift behielten.
    Statt es weiter zu versuchen, wird der Umstand jetzt gemeldet.
    """
    try:
        return sheet.UsedRange.FormatConditions.Count > 0
    except Exception:
        return False


def _set_list_object_fonts(sheet, font_name: str) -> None:
    try:
        los = sheet.ListObjects
        count = los.Count
    except Exception:
        return

    for i in range(1, count + 1):
        try:
            lo = los.Item(i)
            try:
                _zellschrift_setzen(sheet, lo.Range, font_name)
            except Exception as _e:
                detail_logger.debug(f"_set_list_object_fonts: Exception verworfen: {_e!r}")
        except Exception as _e:
            detail_logger.debug(f"_set_list_object_fonts: Exception verworfen: {_e!r}")


def _set_pivot_table_fonts(sheet, font_name: str) -> None:
    try:
        pts = sheet.PivotTables()
        count = pts.Count
    except Exception:
        return

    for i in range(1, count + 1):
        try:
            pt = pts.Item(i)
            try:
                _zellschrift_setzen(sheet, pt.TableRange2, font_name)
            except Exception:
                try:
                    _zellschrift_setzen(sheet, pt.TableRange1, font_name)
                except Exception as _e:
                    detail_logger.debug(f"_set_pivot_table_fonts: Exception verworfen: {_e!r}")
        except Exception as _e:
            detail_logger.debug(f"_set_pivot_table_fonts: Exception verworfen: {_e!r}")

# ==================================================================
# Blattschutz temporär aufheben
# ==================================================================
def _unprotect_sheet_if_possible(sheet) -> bool:
    """
    Versucht, ein Tabellenblatt zu entsperren. Blaetter, die OHNE Passwort
    entsperrt werden konnten, bleiben dauerhaft entsperrt (kein Re-Protect).

    Unprotect("") statt Unprotect(): ohne Argument oeffnet Excel bei einem
    passwortgeschuetzten Blatt einen Kennwortdialog - auch mit
    Visible=False und Interactive=False - und der Aufruf blockiert, bis
    jemand den (u. U. gar nicht sichtbaren) Dialog schliesst. Gemessen mit
    Excel 2024: Unprotect() haengt ueber 60 s, Unprotect("") wirft bei
    Kennwortschutz nach 0,03 s und entsperrt ein Blatt ohne Kennwort.
    Danach wird ProtectContents nachgelesen: vorher meldete die Funktion
    "entsperrt", sobald Unprotect() ohne Ausnahme zurueckkam - im Testlauf
    auch fuer ein Blatt, das hinterher weiterhin geschuetzt war.
    """
    try:
        if not sheet.ProtectContents:
            return False
    except Exception:
        return False

    try:
        sheet.Unprotect("")
    except Exception:
        return False
    try:
        return not sheet.ProtectContents
    except Exception:
        return False


# ==================================================================
# Metadaten-Entfernung
# ==================================================================
def _remove_excel_metadata(workbook, selected: dict) -> bool:
    any_removed = False

    if selected.get("props"):
        for prop_name, prop_label in BUILTIN_PROPS_TO_CLEAR.items():
            try:
                workbook.BuiltinDocumentProperties(prop_name).Value = ""
                detail_logger.debug(f"Eigenschaft geleert: {prop_label}")
                any_removed = True
            except Exception as e:
                detail_logger.debug(f"Eigenschaft '{prop_label}' nicht änderbar: {e}")

    if selected.get("custom"):
        try:
            props = workbook.CustomDocumentProperties
            for i in range(props.Count, 0, -1):
                try:
                    props(i).Delete()
                    any_removed = True
                except Exception as e:
                    detail_logger.debug(f"Custom-Prop {i} nicht löschbar: {e}")
            detail_logger.debug("Benutzerdefinierte Eigenschaften geleert")
        except Exception as e:
            detail_logger.debug(f"CustomDocumentProperties nicht verfügbar: {e}")

        try:
            ct_props = workbook.ContentTypeProperties
            for i in range(ct_props.Count, 0, -1):
                try:
                    ct_props(i).Delete()
                    any_removed = True
                except Exception as _e:
                    detail_logger.debug(f"_remove_excel_metadata: Exception verworfen: {_e!r}")
            detail_logger.debug("ContentTypeProperties geleert")
        except Exception as e:
            detail_logger.debug(f"ContentTypeProperties nicht verfügbar: {e}")

    if selected.get("personal"):
        try:
            workbook.RemovePersonalInformation = COM_TRUE
            detail_logger.debug("RemovePersonalInformation = COM_TRUE gesetzt")
            any_removed = True
        except Exception as e:
            detail_logger.debug(f"RemovePersonalInformation nicht setzbar: {e}")

    if selected.get("comments"):
        for sheet in workbook.Worksheets:
            try:
                sheet.Cells.ClearComments()
                detail_logger.debug(f"Kommentare geleert: {sheet.Name}")
                any_removed = True
            except Exception as e:
                detail_logger.debug(f"ClearComments auf {sheet.Name} fehlgeschlagen: {e}")

    if selected.get("threads"):
        for sheet in workbook.Worksheets:
            try:
                tc = sheet.CommentsThreaded
                for i in range(tc.Count, 0, -1):
                    try:
                        tc(i).Delete()
                        any_removed = True
                    except Exception as _e:
                        detail_logger.debug(f"_remove_excel_metadata: Exception verworfen: {_e!r}")
                detail_logger.debug(f"Thread-Kommentare geleert: {sheet.Name}")
            except Exception as e:
                detail_logger.debug(f"CommentsThreaded auf {sheet.Name} nicht verfügbar: {e}")

    return any_removed

# ==================================================================
# VBA-Erkennung
# ==================================================================
def _workbook_has_real_macros(workbook) -> tuple:
    try:
        if not workbook.HasVBProject:
            return (False, False)
    except Exception as e_has:
        try:
            comps = workbook.VBProject.VBComponents
            for i in range(1, comps.Count + 1):
                try:
                    comp = comps.Item(i)
                    if comp.Type != VBEXT_CT_DOCUMENT:
                        return (True, False)
                    if comp.CodeModule.CountOfLines > 0:
                        return (True, False)
                except Exception:
                    continue
            return (False, False)
        except Exception:
            detail_logger.warning(
                f"VBA-Prüfung nicht möglich (Trust-Center-Konfiguration prüfen): {e_has}")
            return (True, True)

    try:
        comps = workbook.VBProject.VBComponents
        for i in range(1, comps.Count + 1):
            try:
                comp = comps.Item(i)
                if comp.Type != VBEXT_CT_DOCUMENT:
                    return (True, False)
                if comp.CodeModule.CountOfLines > 0:
                    return (True, False)
            except Exception:
                continue
        return (False, False)
    except Exception:
        return (True, False)


# ==================================================================
# XLM-Funktionen in definierten Namen (gleiche Regel wie 3b)
# ==================================================================
# Excel-4.0-Makrofunktionen (XLM) stehen nicht nur auf Makroblaettern,
# sondern auch in definierten Namen - verbreitet fuer Blattlisten
# (GET.WORKBOOK), Zellformat-Abfragen (GET.CELL) und Textformeln
# (EVALUATE). Beim Speichern in ein makrofreies Format (.xlsx/.xltx)
# nennt Excel sie als nicht speicherbar ("Excel 4.0-Funktionen, die in
# definierten Namen gespeichert sind"); mit DisplayAlerts=False wird ohne
# Rueckfrage makrofrei gespeichert, jede Zelle mit Bezug auf einen solchen
# Namen zeigt danach #NAME?. Die VBA-Pruefung sieht davon nichts - Stage 1
# muss solche .xls/.xlt deshalb wie VBA-Mappen als .xlsm/.xltm sichern.
#
# Auswahl der Muster (Gross/Klein egal):
#   - GET.<...>(       alle XLM-Abfragefunktionen heissen GET.* (GET.CELL,
#                      GET.WORKBOOK, GET.DOCUMENT, GET.WORKSPACE,
#                      GET.FORMULA, GET.NAME, GET.DEF, GET.OBJECT,
#                      GET.WINDOW, GET.NOTE, GET.LINK.INFO ...). Keine
#                      Tabellenfunktion beginnt mit "GET." (GETPIVOTDATA
#                      hat keinen Punkt).
#   - <...>.ZUORDNEN(  deutsche Form derselben Familie (ZELLE.ZUORDNEN,
#                      ARBEITSMAPPE.ZUORDNEN, DOKUMENT.ZUORDNEN ...); keine
#                      Tabellenfunktion endet so.
#   - einzeln          XLM-Funktionen ohne GET, die in Namen vorkommen und
#                      KEINE gleichnamige Tabellenfunktion haben:
#                      EVALUATE/AUSWERTEN, FILES/DATEIEN, DOCUMENTS/
#                      DOKUMENTE, DIRECTORY/VERZEICHNIS, ACTIVE.CELL/
#                      AKTIVE.ZELLE, CALLER/AUFRUFER, SELECTION, NAMES/
#                      NAMEN, LINKS, WINDOWS, REFTEXT, TEXTREF, ABSREF,
#                      RELREF, DEREF, CALL, REGISTER.ID.
#                      Bewusst NICHT: CELL/ZELLE, INFO (Tabellenfunktionen).
# Name.RefersTo liefert die Formel laut Dokumentation in der Makrosprache
# (englisch); die deutschen Formen sind Absicherung. Sie stammen aus
# Literatur/Foren, nicht aus einer Messung an diesem Rechner - ein
# Fehlgriff kostet hoechstens .xlsm statt .xlsx, nie Daten. Vor dem Namen
# darf kein Buchstabe, keine Ziffer, kein "_" und kein "." stehen.
_XLM_IN_NAMEN_RE = re.compile(
    r"(?<![A-Za-z0-9_.À-ſ])"
    r"(?:GET\.[A-Za-z0-9_.]+"
    r"|[A-Za-zÀ-ſ][A-Za-z0-9_.À-ſ]*\.ZUORDNEN"
    r"|EVALUATE|AUSWERTEN|FILES|DATEIEN|DOCUMENTS|DOKUMENTE"
    r"|DIRECTORY|VERZEICHNIS|ACTIVE\.CELL|AKTIVE\.ZELLE|CALLER|AUFRUFER"
    r"|SELECTION|NAMES|NAMEN|LINKS|WINDOWS|REFTEXT|TEXTREF|ABSREF|RELREF"
    r"|DEREF|CALL|REGISTER\.ID)\(",
    re.IGNORECASE,
)

# Name.MacroType: xlFunction = 1, xlCommand = 2 (xlNotXLM = 3). Ein Name
# mit Typ 1/2 IST ein XLM-Makro (Makroblatt-Funktion bzw. -Befehl).
_XL_NAME_MACRO_TYPES = (1, 2)


def _name_text(nm, ersatz) -> str:
    # getattr(..., Vorgabe) faengt nur AttributeError, ein COM-Fehler
    # beim Lesen kaeme durch - daher eigenes try.
    try:
        return str(nm.Name)
    except Exception:
        return f"Nr. {ersatz}"


def _hat_xlm_namen(workbook) -> tuple:
    """(True, Grund), wenn ein definierter Name XLM-Funktionen nutzt oder
    selbst ein XLM-Makro ist. Nicht lesbare Namen/Bezuege zaehlen
    KONSERVATIV als Makro: ein unnoetiges .xlsm kostet nichts, ein
    faelschliches .xlsx kostet die Formeln."""
    try:
        names  = workbook.Names
        anzahl = int(names.Count)
    except Exception as e:
        return True, f"Namensliste nicht lesbar ({e!r})"
    for i in range(1, anzahl + 1):
        try:
            nm = names.Item(i)
        except Exception as e:
            return True, f"Name Nr. {i} nicht lesbar ({e!r})"
        try:
            bezug = str(nm.RefersTo)
        except Exception as e:
            return True, f"Bezug von Name Nr. {i} nicht lesbar ({e!r})"
        if _XLM_IN_NAMEN_RE.search(bezug):
            return True, f"XLM-Funktion in Name '{_name_text(nm, i)}'"
        try:
            if int(nm.MacroType) in _XL_NAME_MACRO_TYPES:
                return True, f"Name '{_name_text(nm, i)}' ist ein XLM-Makro"
        except Exception as _e:
            detail_logger.debug(f"_hat_xlm_namen: MacroType nicht lesbar: {_e!r}")
    return False, ""

# ==================================================================
# Schriftart-Ersetzung + Metadaten  (Kern-Routine)
# ==================================================================
def replace_fonts_in_workbook(
    file_path: str,
    excel_app: win32com.client.CDispatch,
    pbar: tqdm,
    font_name: str,
    metadata_selected: dict,
) -> str:
    pbar.write(f"Prüfe: {os.path.basename(file_path)}")
    detail_logger.info(f"=== Starte: {file_path} ===")

    # Sperrpruefung wie im Probelauf. is_locked_by_other wurde bisher
    # AUSSCHLIESSLICH von dry_run_directory aufgerufen: der Probelauf meldete
    # geoeffnete Mappen als 'wuerde uebersprungen', der Echtlauf pruefte die
    # Office-Owner-Datei '~$<name>' an keiner Stelle und griff die Datei
    # trotzdem an. Damit sagte der Probelauf etwas anderes voraus, als der
    # Echtlauf tat - und genau dafuer ist er da.
    if is_locked_by_other(file_path):
        pbar.write("  → ÜBERSPRUNGEN: Datei ist in Excel geöffnet (~$-Datei vorhanden).")
        detail_logger.warning(f"Uebersprungen (geoeffnet): {file_path}")
        return "SKIPPED"

    original_path       = file_path
    is_temp_copy        = False
    was_converted       = False
    workbook            = None
    new_path            = None
    temp_stage1_path    = None
    # temp_stage2_path haelt den Pfad zur Stage-2-SaveCopyAs-Datei, in die
    # die finalen Font-Ersetzungen geschrieben werden (statt direkt ins
    # Original). Erst nach Close + erfolgreichem _av_safe_move wird das
    # Original ueberschrieben. Schuetzt vor Datenverlust bei Crash
    # (Watchdog-Kill, Netzwerk-Drop, COM-Fehler) waehrend Save.
    temp_stage2_path    = None
    # True, sobald Stage 1 per SaveAs DIREKT auf new_path (Freigabe)
    # schreibt - auch wenn dieser SaveAs selbst abbricht, kann dort schon
    # eine (halbe) Datei liegen. Siehe Aufraeumen im except.
    stage1_gestartet    = False
    # True, wenn die Mappe schreibend direkt von der Ablage geoeffnet wurde
    # (keine Temp-Kopie): dann legt Excel dort die Besitzerdatei '~$<name>'
    # an, die nach einem Watchdog-Kill liegen bleibt.
    direkt_geoeffnet    = False
    _lp_cleanup_handled = False
    _orig_calc          = XL_CALC_AUTO
    had_protected_sheet = False
    ext                 = os.path.splitext(file_path)[1].lower()

    # ── Zeitstempel sichern ───────────────────────────────────────────
    orig_times = None
    try:
        h_src = win32file.CreateFile(
            prepare_long_path(original_path),
            FILE_READ_ATTRIBUTES,
            win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
            None, win32file.OPEN_EXISTING, 0, None
        )
        try:
            orig_times = win32file.GetFileTime(h_src)
        finally:
            h_src.Close()
    except Exception:
        orig_times = None

    # NTFS-Sicherheitsinfo (Owner/Group/DACL) des Originals sichern –
    # wird nach der Datei-Ersetzung wieder angewendet (Admin-Kontext:
    # vollständig inkl. Owner; Nutzer-Kontext: DACL).
    orig_sd = _get_security_descriptor(original_path)
    orig_attr = _dateiattribute_lesen(original_path)

    # ── Kennwort-Erkennung (Verschluesselung, Schreibkennwort) ─────────
    _grund = _excel_kennwort_grund(file_path)
    if _grund:
        _text = "verschlüsselt" if _grund == "verschluesselt" else "Schreibkennwort"
        pbar.write(f"  →  ÜBERSPRUNGEN ({_text}): "
                   f"{os.path.basename(original_path)}")
        detail_logger.info(f"Übersprungen ({_text}): {original_path}")
        return "SKIPPED"

    # ── Long-Path-Behandlung ──────────────────────────────────────────
    if len(file_path) > MAX_PATH_LEN:
        detail_logger.debug("Long-Path erkannt, erstelle Temp-Kopie")
        try:
            os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)
            uid         = uuid.uuid4().hex
            basename    = os.path.basename(file_path)
            unique_name = f"{uid}_{basename}"
            temp_path   = os.path.join(TEMP_PROCESS_PATH, unique_name)

            if len(temp_path) >= WINDOWS_MAX_PATH:
                unique_name = f"{uid}{os.path.splitext(basename)[1].lower()}"
                temp_path   = os.path.join(TEMP_PROCESS_PATH, unique_name)
                detail_logger.debug(
                    f"Dateiname zu lang für Temp-Pfad, kürze auf UUID: {temp_path}")

            _av_safe_copy2(prepare_long_path(file_path), temp_path)
            # WICHTIG: Schreibschutz JETZT entfernen, BEVOR _wait_file_released
            # versucht open(temp_path, "ab"). _av_safe_copy2 (shutil.copy2)
            # uebernimmt das Read-Only-Attribut der Quelldatei. _wait_file_released
            # wuerde sonst den PermissionError aus open('ab') als AV-Lock
            # fehldeuten und 15s sinnlos warten, bevor dann unten chmod
            # ausgefuehrt wird.
            try:
                os.chmod(temp_path, os.stat(temp_path).st_mode | stat.S_IWRITE)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
            if not _wait_file_released(temp_path, timeout=AV_POST_COPY_WAIT_SECONDS):
                detail_logger.warning(
                    f"Temp-Datei nach {AV_POST_COPY_WAIT_SECONDS}s noch gesperrt "
                    f"(AV-Scanner?): {temp_path}")
            file_path    = temp_path
            is_temp_copy = True
            detail_logger.debug(f"Temp-Kopie: {temp_path}")
        except Exception as e:
            log_error(original_path, Exception(f"Long-Path-Kopie fehlgeschlagen: {e}"))
            return "ERROR"

    # ── Schreibschutz aufheben ────────────────────────────────────────
    if not is_temp_copy:
        try:
            lp = prepare_long_path(file_path)
            current_mode = os.stat(lp).st_mode
            if not (current_mode & stat.S_IWRITE):
                os.chmod(lp, current_mode | stat.S_IWRITE)
                detail_logger.debug(f"Schreibschutz (Dateiattribut) entfernt: {file_path}")
        except Exception as e_chmod:
            detail_logger.debug(f"Schreibschutz nicht änderbar: {e_chmod}")

    try:
        direkt_geoeffnet = not is_temp_copy
        workbook = safe_excel_open(
            excel_app,
            file_path,
            timeout          = OPEN_TIMEOUT_SECONDS,
            UpdateLinks      = 0,
            ReadOnly         = COM_FALSE,
            IgnoreReadOnlyRecommended = COM_TRUE,
            Notify           = COM_FALSE,
            AddToMru         = COM_FALSE,
            Password         = "",
            WriteResPassword = "",
            CorruptLoad      = 0,
        )
        if workbook is None:
            raise RuntimeError(
                "Arbeitsmappe konnte nicht geöffnet werden (COM retournierte None).")
        detail_logger.debug("Arbeitsmappe geöffnet")

        try:
            _orig_calc = excel_app.Calculation
            excel_app.Calculation = XL_CALC_MANUAL
        except Exception as _e:
            detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")

        try:
            excel_app.DisplayAlerts = COM_FALSE
        except Exception as _e:
            detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")

        # ── STAGE 1: Alte Formate konvertieren ────────────────────────
        if ext in (".xls", ".xlt"):
            has_macros, vba_check_failed = _workbook_has_real_macros(workbook)

            if not has_macros and not vba_check_failed:
                try:
                    for sht in workbook.Sheets:
                        try:
                            if sht.Type in (XL_EXCEL4_MACRO_SHEET,
                                            XL_EXCEL4_INTL_MACRO_SHEET):
                                has_macros = True
                                detail_logger.info(
                                    f"Excel 4.0-Makroblatt erkannt (Typ {sht.Type}): "
                                    f"'{sht.Name}' → Speichere als Makro-Format")
                                break
                        except Exception as _e:
                            detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
                except Exception as _e:
                    detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")

            # XLM in definierten Namen (GET.CELL, EVALUATE ...) - siehe
            # _hat_xlm_namen. Frueher nur VBA und Makroblaetter geprueft:
            # solche .xls wurden .xlsx, die Namen fielen weg (#NAME?).
            if not has_macros:
                _xlm_namen, _xlm_grund = _hat_xlm_namen(workbook)
                if _xlm_namen:
                    has_macros = True
                    detail_logger.info(
                        f"XLM in definierten Namen ({_xlm_grund}) → "
                        f"Speichere als Makro-Format: {original_path}")
                    pbar.write(f"  ⚠  {_xlm_grund} → Makroformat")

            if ext == ".xlt":
                new_format = XL_FORMAT_XLTM if has_macros else XL_FORMAT_XLTX
                new_ext    = ".xltm"         if has_macros else ".xltx"
            else:
                new_format = XL_FORMAT_XLSM if has_macros else XL_FORMAT_XLSX
                new_ext    = ".xlsm"         if has_macros else ".xlsx"

            if vba_check_failed:
                pbar.write(
                    f"  ⚠  VBA-Prüfung nicht möglich → als {new_ext} gesichert: "
                    f"{os.path.basename(original_path)}")

            new_path = os.path.splitext(original_path)[0] + new_ext

            if os.path.exists(prepare_long_path(new_path)):
                base_no_ext = os.path.splitext(original_path)[0]
                counter = 1
                while os.path.exists(prepare_long_path(
                        f"{base_no_ext}_{counter}{new_ext}")) and counter < 100:
                    counter += 1
                new_path = f"{base_no_ext}_{counter}{new_ext}"
                if os.path.exists(prepare_long_path(new_path)):
                    raise RuntimeError(
                        "Mehr als 100 Namenskollisionen! Abbruch zum Schutz fremder Dateien.")
                pbar.write(
                    f"  ⚠  Ziel existiert bereits – speichere als "
                    f"'{os.path.basename(new_path)}'")
                detail_logger.info(
                    f"Namenskollision gelöst: {original_path} → {new_path}")

            excel_app.DisplayAlerts = COM_FALSE

            if is_temp_copy:
                temp_stage1_name = f"{uuid.uuid4().hex}{new_ext}"
                temp_stage1_path = os.path.join(TEMP_PROCESS_PATH, temp_stage1_name)
                safe_excel_save(
                    excel_app,
                    lambda: workbook.SaveAs(
                        temp_stage1_path, FileFormat=new_format, AddToMru=COM_FALSE),
                )
                detail_logger.debug(
                    f"Stage 1 (Long-Path): Als {new_ext} gespeichert → {temp_stage1_path}")

                if os.path.exists(file_path):
                    try:
                        _av_safe_remove(file_path)
                    except Exception as _e:
                        detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
                file_path    = temp_stage1_path
                is_temp_copy = False
            else:
                stage1_gestartet = True
                safe_excel_save(
                    excel_app,
                    lambda: workbook.SaveAs(
                        new_path, FileFormat=new_format, AddToMru=COM_FALSE),
                )
                detail_logger.debug(f"Stage 1: Als {new_ext} gespeichert → {new_path}")

            was_converted = True
            detail_logger.debug(
                "Stage 1 abgeschlossen – Font-Verarbeitung startet auf modernem Format")

        # ── 1. FORMATVORLAGEN ─────────────────────────────────────────
        try:
            for style in workbook.Styles:
                try:
                    _font_name_setzen(style.Font, font_name)
                except Exception as _e:
                    detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
        except Exception as e:
            detail_logger.debug(f"Styles nicht änderbar: {e}")

        # ── 2. TABELLENBLÄTTER ────────────────────────────────────────
        # Nicht-passwortgeschuetzte Blaetter werden hier ENTSPERRT und
        # bleiben anschliessend ENTSPERRT (kein Re-Protect).

        for sheet in workbook.Worksheets:
            try:
                _sheet_protected = False
                try:
                    _sheet_protected = sheet.ProtectContents
                except Exception as _e:
                    detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")

                if _sheet_protected:
                    _was_unprotected = _unprotect_sheet_if_possible(sheet)
                    if _was_unprotected:
                        detail_logger.info(
                            f"Blatt '{sheet.Name}' entsperrt "
                            f"(bleibt entsperrt)")
                    else:
                        had_protected_sheet = True
                        pbar.write(
                            f"  ⚠  Blatt '{sheet.Name}' ist passwortgeschützt "
                            f"– Schriftarten nicht vollständig änderbar")
                        detail_logger.warning(
                            f"Blatt '{sheet.Name}' passwortgeschützt – skip")

                try:
                    _ur = sheet.UsedRange
                    _ur_rows = _ur.Rows.Count
                    _ur_cols = _ur.Columns.Count
                    if _ur_rows > 500000 or _ur_cols > 16000:
                        detail_logger.warning(
                            f"UsedRange zu groß auf '{sheet.Name}': "
                            f"{_ur_rows} Zeilen × {_ur_cols} Spalten – übersprungen")
                        pbar.write(
                            f"  ⚠  Blatt '{sheet.Name}': UsedRange unrealistisch groß "
                            f"({_ur_rows}×{_ur_cols}) – Schriftart nur via Styles gesetzt")
                    elif not _zellschrift_setzen(sheet, _ur, font_name):
                        detail_logger.warning(
                            f"Blatt '{sheet.Name}': sehr kleinteilig formatiert - "
                            f"Rest pauschal gesetzt; Symbolzeichen in Mischzellen "
                            f"koennen umgestellt sein")
                except Exception as e:
                    detail_logger.debug(
                        f"UsedRange.Font.Name fehlgeschlagen auf '{sheet.Name}': {e}")

                # Bedingte Formatierung traegt technisch keinen Schriftnamen
                # (siehe _hat_bedingte_formatierung). Nur vermerken, nicht
                # vergeblich setzen.
                if _hat_bedingte_formatierung(sheet):
                    detail_logger.info(
                        f"Blatt '{sheet.Name}': bedingte Formatierung vorhanden – "
                        f"deren Schriftart bleibt unveraendert (von Excel nicht "
                        f"aenderbar).")
                _set_list_object_fonts(sheet, font_name)
                _set_pivot_table_fonts(sheet, font_name)

                if not metadata_selected.get("comments"):
                    try:
                        for comment in sheet.Comments:
                            try:
                                _font_name_setzen(comment.Shape.TextFrame.Characters().Font, font_name)
                            except Exception as _e:
                                detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
                    except Exception as _e:
                        detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")

                try:
                    for shape in sheet.Shapes:
                        try:
                            _process_sheet_shape(shape, font_name, 0)
                        except Exception as _e:
                            detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
                except Exception as _e:
                    detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")

                _replace_header_footer_fonts(sheet, font_name)

            except Exception as e:
                detail_logger.warning(f"Fehler bei Tabellenblatt '{sheet.Name}': {e}")

        # ── 3. EIGENSTÄNDIGE DIAGRAMMBLÄTTER ──────────────────────────
        try:
            for chart_sheet in workbook.Charts:
                try:
                    _set_chart_fonts(chart_sheet, font_name)
                except Exception as _e:
                    detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
        except Exception as e:
            detail_logger.debug(f"Charts-Collection nicht verfügbar: {e}")

        detail_logger.debug("Schriftarten ersetzt")

        # ── 3.5 METADATEN ENTFERNEN ───────────────────────────────────
        if any(metadata_selected.values()):
            try:
                _remove_excel_metadata(workbook, metadata_selected)
                detail_logger.debug("Metadaten bereinigt")
            except Exception as e_meta:
                detail_logger.warning(
                    f"Metadaten-Bereinigung fehlgeschlagen (nicht kritisch): {e_meta}")

        # ── 4. SPEICHERN ──────────────────────────────────────────────
        try:
            excel_app.Calculation = _orig_calc
        except Exception as _e:
            detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
        excel_app.DisplayAlerts = COM_FALSE

        # DATENVERLUST-SCHUTZ:
        #  - was_converted=True : Workbook wurde in Stage 1 ueber SaveAs auf
        #                         temp_stage1_path gewechselt. Save() schreibt
        #                         dorthin - Original (.xls) ist intakt.
        #  - is_temp_copy=True  : Workbook wurde aus Long-Path-Temp geoeffnet.
        #                         Save() schreibt in die Temp-Kopie - Original
        #                         wird erst spaeter via Block 6/7 zurueck-
        #                         geschrieben.
        #  - sonst              : Workbook wurde DIREKT vom Original geoeffnet.
        #                         Ein nacktes Save() wuerde In-Place auf das
        #                         Original schreiben. Bei Crash mittendrin
        #                         (Watchdog-Kill, Netzwerk-Drop, COM-Fehler)
        #                         waere das Original korrupt/leer. Daher:
        #                         SaveCopyAs in temp_stage2_path, dann nach
        #                         Close per _av_safe_move atomar zum Original.
        save_in_place_safe = was_converted or is_temp_copy

        if save_in_place_safe:
            # Anders als PowerPoint braucht Excel kein erzwungenes SaveAs zur
            # Schema-Re-Serialisierung - das uebernimmt der regulaere Save bei
            # modernen OOXML-Formaten ohnehin.
            safe_excel_save(excel_app, workbook.Save)
            detail_logger.debug("Gespeichert (Fonts + Metadaten, in-place sicher)")
        else:
            # Stage-2-SaveCopyAs in Temp-Pfad. SaveCopyAs schreibt die
            # aktuelle Workbook-Sicht in eine NEUE Datei, OHNE die offene
            # Workbook-Referenz zu wechseln (im Gegensatz zu SaveAs).
            # Das Original auf der Platte wird dabei NICHT angefasst.
            os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)
            temp_stage2_name = f"stage2_{uuid.uuid4().hex}{ext}"
            temp_stage2_path = os.path.join(TEMP_PROCESS_PATH, temp_stage2_name)

            safe_excel_save(
                excel_app,
                lambda: workbook.SaveCopyAs(temp_stage2_path)
            )
            detail_logger.debug(
                f"Stage 2: SaveCopyAs -> {temp_stage2_path} (Fonts + Metadaten)")

        detail_logger.debug("Gespeichert (Fonts + Metadaten)")

        # ── 5. SCHLIESSEN ─────────────────────────────────────────────
        if workbook is not None:
            workbook.Close(SaveChanges=COM_FALSE)
            workbook = None
            detail_logger.debug("Arbeitsmappe geschlossen")

        # ── 5.5. Stage-2-Tempdatei → Original verschieben ──────────────
        # Nur wenn der direct-on-original-Save-Pfad genommen wurde
        # (was_converted=False, is_temp_copy=False). In den anderen
        # Faellen erledigen Block 6 (Stage-1-Konvertierung) oder Block 7
        # (Long-Path-Restore) den Move.
        if temp_stage2_path is not None:
            bak_orig = None
            # Die Stage-2-Datei traegt die fertigen Aenderungen und ist bis zum
            # bestaetigten Move die einzige Quelle dafuer. Ein Strg+C in diesem
            # Fenster laesst den Signal-Handler _safe_cleanup_temp() aufrufen,
            # der TEMP_PROCESS_PATH per rmtree abraeumt - 'stage2_*' war weder
            # RESCUE_ noch geschuetzt und damit weg. Deshalb ab hier als
            # geschuetzt fuehren; nach erfolgreichem Move wieder freigeben.
            _preserved_temp_files.add(temp_stage2_path)
            try:
                safe_orig = prepare_long_path(original_path)
                if not (os.path.exists(temp_stage2_path)
                        and os.path.getsize(temp_stage2_path) > 0):
                    raise RuntimeError(
                        f"Stage-2-Tempdatei ist leer/fehlt: {temp_stage2_path}")
                # Wait-File-Released sicherstellen, dass Excel die
                # Datei nicht mehr lockt (sollte nach Close eigentlich
                # nicht mehr der Fall sein, aber AV-Scanner koennen
                # die Datei noch kurz halten).
                _wait_file_released(temp_stage2_path, timeout=10.0)
                # Original VOR dem Move per atomarem Rename auf demselben
                # Laufwerk sichern: shutil.move (_av_safe_move) faellt bei
                # existierendem Ziel bzw. Cross-Volume-Move (Temp lokal,
                # Ziel auf Netzlaufwerk) auf copy2 + unlink zurueck und
                # trunkiert das Ziel sofort auf 0 Bytes ('wb'). Bricht der
                # Kopiervorgang ab (Netzwerk-Drop, AV-Scanner), waere das
                # Original ohne Backup unwiderruflich zerstoert. Analog
                # zum .bak-Schutz in Block 8 (Long-Path-Restore).
                if os.path.exists(safe_orig):
                    try:
                        current_mode = os.stat(safe_orig).st_mode
                        os.chmod(safe_orig, current_mode | stat.S_IWRITE)
                    except Exception as _e:
                        detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
                    bak_orig = prepare_long_path(
                        original_path + f"_{uuid.uuid4().hex[:6]}.bak")
                    _av_safe_replace(safe_orig, bak_orig)
                _av_safe_move(temp_stage2_path, safe_orig)
                if not verify_saved_file(safe_orig):
                    raise RuntimeError(
                        f"Zurueckgeschobene Datei nicht valide: {original_path}")
                # Move bestaetigt - der Schutz wird nicht mehr gebraucht.
                _preserved_temp_files.discard(temp_stage2_path)
                temp_stage2_path = None
                if bak_orig and os.path.exists(bak_orig):
                    try:
                        _av_safe_remove(bak_orig)
                    except Exception as e_bak_del:
                        detail_logger.warning(
                            f"Backup-Datei nicht löschbar: {bak_orig} ({e_bak_del})")
                detail_logger.debug(
                    f"Stage-2-Tempdatei -> Original verschoben: {original_path}")
            # BaseException, nicht Exception: Der Signal-Handler endet mit
            # sys.exit(130), und das loest SystemExit aus - eine BaseException.
            # Bei Strg+C mitten im Move lief die Wiederherstellung aus dem
            # .bak deshalb NIE, obwohl das Original oben bereits auf den
            # .bak-Namen umbenannt war. Zurueck blieb nur
            # 'Datei.xlsx_9f3a1c.bak' - ein Name, der unter
            # SKIP_FILE_SUFFIXES faellt und damit auch von jedem Folgelauf
            # ignoriert wird. Das Fenster ist nicht schmal: _av_safe_move
            # wiederholt bei AV-Sperren bis zu AV_MAX_RETRIES-mal.
            except BaseException as e_s2_move:
                # Move fehlgeschlagen - Original aus dem .bak-Backup
                # wiederherstellen, falls der Move es bereits angetastet
                # hat. Stage-2-Tempdatei zusaetzlich als RESCUE_-Datei
                # sichern.
                if bak_orig and os.path.exists(bak_orig):
                    try:
                        if os.path.exists(safe_orig):
                            _av_safe_remove(safe_orig)
                        _av_safe_replace(bak_orig, safe_orig)
                        detail_logger.warning(
                            f"Stage-2-Move fehlgeschlagen – Original aus "
                            f"Backup wiederhergestellt: {original_path}")
                    except Exception as e_restore:
                        detail_logger.error(
                            f"Original konnte NICHT aus Backup wiederhergestellt "
                            f"werden: {e_restore}\n"
                            f"  Backup liegt unter: {bak_orig}")
                rescue_name = f"RESCUE_{uuid.uuid4().hex}{ext}"
                rescue_path = os.path.join(TEMP_PROCESS_PATH, rescue_name)
                detail_logger.warning(
                    f"Stage-2-Move fehlgeschlagen: {e_s2_move}\n"
                    f"  Versuche Rescue-Kopie nach: {rescue_path}\n"
                    f"  Original wiederhergestellt bzw. Backup vorhanden: "
                    f"{original_path}"
                )
                try:
                    if temp_stage2_path and os.path.exists(temp_stage2_path):
                        _av_safe_copy2(temp_stage2_path, rescue_path)
                        detail_logger.warning(
                            f"Stage-2-Rescue gesichert: {rescue_path}")
                        pbar.write(
                            f"  ⚠  RESCUE (Font-Replace): "
                            f"{os.path.basename(rescue_path)} "
                            f"(Original geschützt)")
                except Exception as e_s2_rescue:
                    detail_logger.error(f"Stage-2-Rescue fehlgeschlagen: {e_s2_rescue}")
                    if temp_stage2_path and os.path.exists(temp_stage2_path):
                        _preserved_temp_files.add(temp_stage2_path)
                        detail_logger.error(
                            f"DATENVERLUST-SCHUTZ: Stage-2-Temp wird NICHT geloescht: "
                            f"{temp_stage2_path}\n"
                            f"  → Datei manuell pruefen und sichern!")
                log_error(
                    original_path,
                    Exception(f"Stage-2-Move fehlgeschlagen: {e_s2_move}"))
                pbar.write(
                    f"  ✗  FEHLER (Stage-2-Move): "
                    f"{os.path.basename(original_path)}")
                # SystemExit/KeyboardInterrupt nach der Rettung UNVERAENDERT
                # weiterreichen. Sonst wuerde der Abbruch zu einem blossen
                # 'ERROR' fuer diese eine Datei und der Lauf liefe weiter -
                # das Gegenteil dessen, was Strg+C bedeutet.
                if not isinstance(e_s2_move, Exception):
                    raise
                return "ERROR"

        # ── 6. TEMP-STAGE1 → ZIELDATEI verschieben ────────────────────
        if was_converted and temp_stage1_path is not None:
            safe_new = prepare_long_path(new_path)
            _wait_file_released(temp_stage1_path, timeout=10.0)
            _av_safe_move(temp_stage1_path, safe_new)
            temp_stage1_path = None

            if not verify_saved_file(safe_new):
                raise RuntimeError(
                    f"Konvertierte Datei nicht valide oder leer: {new_path}")

            safe_orig = prepare_long_path(original_path)
            if os.path.exists(safe_orig):
                try:
                    os.chmod(safe_orig, stat.S_IWRITE)
                    _av_safe_remove(safe_orig)
                    detail_logger.debug(f"Alte Datei gelöscht: {original_path}")
                except Exception as e_del_orig:
                    detail_logger.warning(
                        f"Original nach Long-Path-Konvertierung nicht löschbar "
                        f"({e_del_orig}), versuche Umbenennung zu .bak")
                    bak_orig = prepare_long_path(
                        original_path + f"_{uuid.uuid4().hex[:6]}.bak")
                    try:
                        _av_safe_replace(safe_orig, bak_orig)
                        detail_logger.warning(
                            f"Original umbenannt: {os.path.basename(original_path)} → "
                            f"{os.path.basename(bak_orig)}")
                        pbar.write(
                            f"  ⚠  Alte Datei gesichert als: "
                            f"{os.path.basename(bak_orig)}")
                    except Exception as e_ren_orig:
                        detail_logger.warning(
                            f"Original weder lösch- noch umbenennbar: {e_ren_orig}\n"
                            f"  new_path ist gültig: {new_path}")

            _lp_cleanup_handled = True

        # ── 7. ORIGINAL LÖSCHEN (normale Konvertierung) ───────────────
        elif was_converted and not is_temp_copy and not _lp_cleanup_handled \
                and os.path.exists(prepare_long_path(original_path)):
            # Die neue Datei PRUEFEN, bevor das Original geloescht wird.
            # Dieser Zweig ist der haeufigste Produktionsfall (kurzer Pfad,
            # .xls -> .xlsx) und war als einziger Ersetzungspfad ohne
            # Integritaetspruefung - Block 5.5, 6 und 8 haben je einen
            # verify_saved_file-Aufruf, Block 7 hatte keinen. Geschrieben
            # wurde new_path direkt auf die Freigabe (SaveAs + Save); liefert
            # ein Netzwerk-Aussetzer oder ein AV-Eingriff dort eine
            # trunkierte Datei zurueck, ohne dass COM eine Ausnahme wirft,
            # wurde anschliessend das intakte Original geloescht und der Lauf
            # meldete Erfolg. verify_saved_file prueft Existenz, Mindestgroesse
            # und bei OOXML zusaetzlich die ZIP-Struktur.
            if not verify_saved_file(new_path):
                raise RuntimeError(
                    f"Konvertierte Datei nicht valide oder leer: {new_path} "
                    f"– Original bleibt erhalten.")
            safe_orig = prepare_long_path(original_path)
            try:
                os.chmod(safe_orig, stat.S_IWRITE)
                _av_safe_remove(safe_orig)
                detail_logger.debug(f"Alte Datei gelöscht: {original_path}")
            except Exception as e_del:
                detail_logger.warning(
                    f"Alte Datei nicht löschbar ({e_del}), versuche Umbenennung zu .bak")
                bak_path = prepare_long_path(
                    original_path + f"_{uuid.uuid4().hex[:6]}.bak")
                try:
                    _av_safe_replace(safe_orig, bak_path)
                    detail_logger.warning(
                        f"Alte Datei umbenannt: "
                        f"{os.path.basename(original_path)} → {os.path.basename(bak_path)}")
                    pbar.write(
                        f"  ⚠  Alte Datei gesichert als: "
                        f"{os.path.basename(bak_path)}")
                except Exception as e_ren:
                    detail_logger.error(
                        f"Alte Datei weder lösch- noch umbenennbar: {e_ren}")
                    log_error(original_path, Exception(
                        f"Alte Datei nach Konvertierung nicht entfernbar: "
                        f"Löschen: {e_del} | Umbenennen: {e_ren}"))
                    pbar.write(
                        f"  ✗  FEHLER (alte Datei nicht entfernbar): "
                        f"{os.path.basename(original_path)}")
                    return "ERROR"

        # ── 8. LONG-PATH TEMP → ORIGINAL (moderne Formate) ────────────
        if is_temp_copy:
            safe_orig = prepare_long_path(original_path)
            bak_path  = prepare_long_path(
                original_path + f"_{uuid.uuid4().hex[:6]}.bak")
            bak_created = False
            # Erst nach bestandener Pruefung gilt die zurueckgeschobene Datei
            # als fertig; ein Abbruch danach (etwa Strg+C beim Loeschen des
            # .bak) darf sie nicht mehr durch das alte Original ersetzen.
            rueck_bestaetigt = False
            try:
                if os.path.exists(safe_orig):
                    os.chmod(safe_orig, stat.S_IWRITE)
                    _av_safe_replace(safe_orig, bak_path)
                    bak_created = True
                _wait_file_released(file_path, timeout=10.0)
                _av_safe_move(file_path, safe_orig)

                if not verify_saved_file(safe_orig):
                    raise RuntimeError(
                        f"Zurückverschobene Datei nicht valide: {original_path}")
                rueck_bestaetigt = True

                detail_logger.debug("Temp-Datei zurückverschoben")
                is_temp_copy = False
                if bak_created:
                    try:
                        _av_safe_remove(bak_path)
                    except Exception as e_bak_del:
                        detail_logger.warning(
                            f"Backup-Datei nicht löschbar: {bak_path} ({e_bak_del})")
            # BaseException wie in Block 5.5: sonst ueberspringt ein Strg+C
            # mitten im Move die Wiederherstellung aus dem .bak.
            except BaseException as e_move:
                # Frueher: "and not os.path.exists(safe_orig)". Genau im
                # gefaehrlichen Fall existiert safe_orig aber: shutil.move
                # (_av_safe_move) faellt ueber Laufwerksgrenzen (Temp lokal,
                # Ziel auf der Freigabe) auf copy2 zurueck und legt das Ziel
                # sofort an - ein Abbruch mittendrin oder eine ungueltige
                # Kopie (verify_saved_file) hinterliess eine halbe Datei am
                # Originalpfad, das intakte Original blieb als '<name>_xxxxxx.bak'
                # liegen und fiel unter SKIP_FILE_SUFFIXES aus jedem Folgelauf.
                # Jetzt wie Block 5.5: halbe Datei entfernen, .bak zurueck.
                if bak_created and not rueck_bestaetigt and os.path.exists(bak_path):
                    try:
                        if os.path.exists(safe_orig):
                            _av_safe_remove(safe_orig)
                        _av_safe_replace(bak_path, safe_orig)
                        bak_created = False
                        detail_logger.warning(
                            f"Zurückschieben fehlgeschlagen – Original aus Backup "
                            f"wiederhergestellt: {original_path} (Backup war: {bak_path})")
                    except Exception as e_restore:
                        detail_logger.error(
                            f"Original konnte NICHT aus Backup wiederhergestellt werden: {e_restore}\n"
                            f"  Backup liegt unter: {bak_path}")
                        pbar.write(f"  ⚠  Original liegt als Backup: {bak_path}")
                elif bak_created and os.path.exists(bak_path):
                    detail_logger.warning(
                        f"Backup-Datei bleibt liegen (Rueckschieben war bestaetigt): {bak_path}")
                rescue_ext  = os.path.splitext(os.path.basename(original_path))[1]
                rescue_name = f"RESCUE_{uuid.uuid4().hex}{rescue_ext}"
                rescue_path = os.path.join(TEMP_PROCESS_PATH, rescue_name)
                detail_logger.warning(
                    f"Zurückschieben fehlgeschlagen: {e_move}\n"
                    f"  Versuche Rescue-Kopie nach: {rescue_path}"
                )
                try:
                    _av_safe_copy2(file_path, rescue_path)
                    detail_logger.warning(f"Rescue-Kopie gesichert: {rescue_path}")
                    pbar.write(f"  ⚠  RESCUE: {os.path.basename(rescue_path)}")
                except Exception as e_rescue:
                    detail_logger.error(f"Rescue fehlgeschlagen: {e_rescue}")
                    is_temp_copy = False
                    _preserved_temp_files.add(file_path)
                    detail_logger.error(
                        f"DATENVERLUST-SCHUTZ: Temp-Datei wird NICHT gelöscht: {file_path}\n"
                        f"  → Datei manuell prüfen und sichern!")
                log_error(original_path, Exception(f"Zurückschieben fehlgeschlagen: {e_move}"))
                pbar.write(f"  ✗  FEHLER (move): {os.path.basename(original_path)}")
                # Abbruch nach der Rettung unveraendert weiterreichen.
                if not isinstance(e_move, Exception):
                    raise
                return "ERROR"

        # ── 8.5. ACL/OWNER WIEDERHERSTELLEN ───────────────────────────
        acl_target = new_path if (was_converted and new_path) else original_path
        if acl_target and os.path.exists(prepare_long_path(acl_target)):
            _apply_security_descriptor(acl_target, orig_sd)

        # ── 9. ZEITSTEMPEL WIEDERHERSTELLEN ───────────────────────────
        if orig_times:
            ts_target = new_path if (was_converted and new_path) else original_path
            if ts_target and os.path.exists(prepare_long_path(ts_target)):
                for av_retry in range(5):
                    try:
                        h_dst = win32file.CreateFile(
                            prepare_long_path(ts_target),
                            FILE_WRITE_ATTRIBUTES,
                            win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                            None, win32file.OPEN_EXISTING, 0, None
                        )
                        try:
                            win32file.SetFileTime(
                                h_dst,
                                orig_times[0],
                                orig_times[1],
                                orig_times[2],
                            )
                        finally:
                            h_dst.Close()
                        break
                    except Exception as e_utime:
                        if av_retry == 4:
                            # Rueckfall auf os.utime: erhaelt Zugriffs- und
                            # Aenderungszeit (nicht die Erstellungszeit),
                            # aber das ist deutlich besser als der
                            # vollstaendige Verlust des Datums. Bisher
                            # hatte nur 5_OCR_PDF.py diesen Rueckfall.
                            if not _utime_rueckfall(ts_target, orig_times):
                                detail_logger.warning(
                                    f"Zeitstempel nicht wiederherstellbar: {e_utime}")
                            else:
                                detail_logger.info(
                                    "Zeitstempel über os.utime-Rückfall gesetzt "
                                    "(ohne Erstellungszeit).")
                        else:
                            time.sleep(0.5)

        # ── 9.5. DATEIATTRIBUTE WIEDERHERSTELLEN (zuletzt: Schreibschutz) ─
        _dateiattribute_zurueck(
            new_path if (was_converted and new_path) else original_path, orig_attr)

        # Nachweisliste: Pfadwechsel (.xls/.xlt -> neu, oder Ausweichname
        # bei Kollision) in die Konvertierungs-CSV.
        if was_converted and new_path and new_path.lower() != original_path.lower():
            log_conversion(original_path, new_path)

        marker = " [konvertiert]" if was_converted else ""
        if had_protected_sheet:
            marker += " [teilweise]"
        pbar.write(f"  ✓  {os.path.basename(original_path)}{marker}")
        detail_logger.info(f"Erfolgreich: {original_path}")
        return "PARTIAL" if had_protected_sheet else "SUCCESS"

    # Frueher stand hier ein eigenes "except TimeoutError: return 'ERROR'"
    # VOR diesem Zweig. Damit lief ausgerechnet beim Watchdog-Kill das
    # Aufraeumen unten nie: bei .xls blieb die unfertige Stage-1-Datei
    # '<name>.xlsx' liegen (Folgelauf wich auf '<name>_1.xlsx' aus), und
    # bei direkt geoeffneten Mappen blieb Excels Besitzerdatei '~$<name>'
    # liegen - Versuch 2 und JEDER Folgelauf meldeten die Mappe dann als
    # "in Excel geöffnet" (is_locked_by_other). TimeoutError ist eine
    # Exception und landet jetzt hier.
    except Exception as e:
        err_msg = str(e).lower()
        ist_timeout = isinstance(e, TimeoutError)

        # Mappe VOR dem Aufraeumen schliessen: nach Stage 1 ist new_path die
        # in Excel offene Datei, Excel haelt sie ohne Loeschfreigabe - das
        # Entfernen unten scheiterte bei jedem Nicht-Timeout-Fehler an der
        # eigenen Sperre (das finally schloss erst danach). Nach einem Kill
        # schlaegt Close still fehl; das ist gleichgueltig.
        if workbook is not None:
            try:
                excel_app.Calculation = _orig_calc
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
            try:
                workbook.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
            workbook = None

        # Das Original steht unveraendert da, nur der Schreibschutz war fuer
        # das Speichern entfernt - zuruecksetzen.
        _dateiattribute_zurueck(original_path, orig_attr)

        # Bereits geschriebene, aber nie gepruefte Zieldatei entfernen.
        # Bei .xls/.xlt schreibt Stage 1 new_path auf die Freigabe, BEVOR
        # irgendeine Schriftersetzung stattgefunden hat. Bricht der Lauf
        # danach ab, blieb diese Datei liegen - ungeprueft, mit alten
        # Schriftarten oder trunkiert. Der Wiederholungslauf
        # (FILE_PROCESSING_ATTEMPTS) fand sie vor, wich auf '<name>_1.xlsx'
        # aus und loeschte am Ende das Original: der Anwender hatte zwei
        # Zieldateien, von denen nur eine brauchbar war, und die verwaiste
        # stand in keinem Protokoll. Das Original wird hier NICHT angetastet.
        # stage1_gestartet: auch ein im SaveAs selbst abgebrochener Stage-1-
        # Schritt kann dort schon eine halbe Datei hinterlassen haben; der
        # Name war vorher frei (Kollisionspruefung), sie ist also unsere.
        # Dieses Aufraeumen steht jetzt VOR der Kennwort-Erkennung: eine
        # Meldung mit "geschützt"/"protected" (z. B. geschuetztes Blatt)
        # nach Stage 1 fuehrte frueher zu SKIPPED ohne Aufraeumen.
        if ((was_converted or stage1_gestartet) and new_path and not is_temp_copy
                and new_path.lower() != original_path.lower()
                and os.path.exists(prepare_long_path(new_path))):
            try:
                _av_safe_remove(new_path)
            except Exception as _e_del:
                detail_logger.debug(f"Zieldatei-Aufraeumen: {_e_del!r}")
            if os.path.exists(prepare_long_path(new_path)):
                detail_logger.error(
                    f"Unfertige Zieldatei NICHT entfernbar - bitte pruefen: {new_path}")
                pbar.write(f"  ⚠  Unfertige Zieldatei bitte prüfen: {os.path.basename(new_path)}")
            else:
                detail_logger.warning(
                    f"Unfertige Zieldatei nach Abbruch entfernt: {new_path}")
                pbar.write(f"  ↺  Unfertige Zieldatei entfernt: {os.path.basename(new_path)}")

        # Besitzerdatei '~$<name>' der getoeteten Instanz entfernen. Nur
        # wenn Excel nicht mehr antwortet (Kill/Absturz) und die Mappe
        # direkt von der Ablage geoeffnet war. Beim Start dieser Funktion
        # gab es fuer original_path keine (is_locked_by_other oben), sie
        # stammt also von uns; eine von einer LEBENDEN Excel-Sitzung
        # offengehaltene Besitzerdatei laesst sich nicht loeschen und
        # bleibt dann unangetastet.
        if direkt_geoeffnet and (ist_timeout or not _excel_antwortet(excel_app)):
            _verwaiste_besitzerdatei_entfernen(original_path)
            if (was_converted or stage1_gestartet) and new_path:
                _verwaiste_besitzerdatei_entfernen(new_path)

        if ist_timeout:
            pbar.write(f"  ✗  TIMEOUT: {os.path.basename(original_path)}")
            detail_logger.error(f"Timeout: {original_path} → {e}")
            log_error(original_path, e)
            return "ERROR"

        is_password_error = any(kw in err_msg for kw in (
            "password", "passwort", "kennwort",
            "protected", "geschützt",
        ))

        if is_password_error:
            pbar.write(f"  →  ÜBERSPRUNGEN (Passwort): {os.path.basename(original_path)}")
            detail_logger.info(f"Übersprungen (Passwort): {original_path}")
            return "SKIPPED"

        log_error(original_path, e)
        pbar.write(f"  ✗  FEHLER: {os.path.basename(original_path)}")
        return "ERROR"

    finally:
        if workbook is not None:
            try:
                excel_app.Calculation = _orig_calc
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
            try:
                workbook.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
        if is_temp_copy and os.path.exists(file_path):
            if file_path in _preserved_temp_files:
                detail_logger.debug(
                    f"Temp-Datei ist geschützt (preserved) – wird NICHT gelöscht: {file_path}")
            else:
                try:
                    _av_safe_remove(file_path)
                except Exception as _e:
                    detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
        if (temp_stage1_path is not None
                and os.path.exists(temp_stage1_path)):
            try:
                _av_safe_remove(temp_stage1_path)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")
        if (temp_stage2_path is not None
                and os.path.exists(temp_stage2_path)
                and temp_stage2_path not in _preserved_temp_files):
            try:
                _av_safe_remove(temp_stage2_path)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_workbook: Exception verworfen: {_e!r}")

# ==================================================================
# Verzeichnis-Verarbeitung
# ==================================================================
def process_directory(
    directory: str,
    font_name: str,
    metadata_selected: dict,
    count_files_first: bool,
    show_progress: bool = True,
    resume_path: Optional[str] = None,
    resume_set: Optional[set] = None,
) -> tuple:
    global TEMP_PROCESS_PATH
    TEMP_PROCESS_PATH = os.path.join(TEMP_BASE_PATH, uuid.uuid4().hex)
    os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)

    resume_set = resume_set or set()

    def _gen():
        for fp in file_generator(directory):
            if resume_set and fp.lower() in resume_set:
                continue
            yield fp

    if count_files_first:
        print("\nIndiziere Dateien (kann einige Minuten dauern) ...")
        files_list = list(_gen())
        total      = len(files_list)
        iterable   = files_list
        bar_fmt    = "{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]"
        print(f"{total} Dateien gefunden"
              + (f" ({len(resume_set)} via Resume übersprungen)" if resume_set else "")
              + ". Starte Verarbeitung ...\n")
    else:
        iterable = _gen()
        total    = None
        bar_fmt  = "{desc}: {n_fmt} Dateien [{elapsed}, {rate_fmt}]"
        print("\nStarte direkte Verarbeitung (ohne Vorab-Zählung) ...\n")

    print("-" * 66)

    stats = {"SUCCESS": 0, "PARTIAL": 0, "ERROR": 0, "SKIPPED": 0}
    run_completed = False

    global excel_app_global
    excel = None

    # ── Excel-Lifecycle ───────────────────────────────────────────────
    def _start_excel():
        nonlocal excel
        global excel_app_global, excel_pid_global, excel_create_time_global
        _kill_specific_excel()
        time.sleep(1)

        _pids_before = set()
        try:
            _pids_before = {
                p.pid for p in psutil.process_iter(["name"])
                if p.info["name"] and p.info["name"].lower() == "excel.exe"
            }
        except Exception as _e:
            detail_logger.debug(f"_start_excel: Exception verworfen: {_e!r}")

        excel            = win32com.client.DispatchEx("Excel.Application")
        excel_app_global = excel

        try:
            _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
            excel_pid_global = pid
            detail_logger.debug(f"Excel gestartet, PID {pid} (via Hwnd)")
        except Exception:
            try:
                _pids_now = {
                    p.pid for p in psutil.process_iter(["name"])
                    if p.info["name"] and p.info["name"].lower() == "excel.exe"
                }
                _new = _pids_now - _pids_before
                if len(_new) == 1:
                    excel_pid_global = _new.pop()
                else:
                    excel_pid_global = None
            except Exception:
                excel_pid_global = None

        # Erstellungszeit der Excel-Instanz sichern (PID-Recycling-Schutz
        # fuer Watchdog-Kill und _kill_specific_excel).
        excel_create_time_global = None
        if excel_pid_global is not None:
            try:
                excel_create_time_global = psutil.Process(
                    excel_pid_global).create_time()
            except Exception as e_ct:
                detail_logger.debug(
                    f"Excel-Erstellungszeit nicht ermittelbar: {e_ct}")

        excel.Visible            = COM_FALSE
        excel.DisplayAlerts      = COM_FALSE
        excel.ScreenUpdating     = COM_FALSE
        excel.EnableEvents       = COM_FALSE
        excel.AutomationSecurity = 3
        try:
            excel.Interactive = COM_FALSE
        except Exception as _e:
            detail_logger.debug(f"_start_excel: Exception verworfen: {_e!r}")

        _disable_com_addins(excel)

    def _check_excel_memory() -> bool:
        if excel_pid_global is None:
            return False
        try:
            mem_mb = psutil.Process(excel_pid_global).memory_info().rss / 1024 / 1024
            if mem_mb > EXCEL_MEMORY_LIMIT_MB:
                detail_logger.warning(
                    f"Excel-Speicher {mem_mb:.0f}MB > Limit {EXCEL_MEMORY_LIMIT_MB}MB – Neustart")
                return True
        except Exception as _e:
            detail_logger.debug(f"_check_excel_memory: Exception verworfen: {_e!r}")
        return False

    # ── Hauptschleife ─────────────────────────────────────────────────
    try:
        _start_excel()
        file_counter = 0

        with tqdm(total=total, desc="Verarbeite", unit="Datei",
                  bar_format=bar_fmt, disable=not show_progress) as pbar:
            for file_path in iterable:
                # Geplanter Restart (gegen 32-Bit-OOM bei Langlaeufen)
                if file_counter > 0 and file_counter % EXCEL_RESTART_INTERVAL == 0:
                    pbar.write(
                        f"  ↻  Geplanter Excel-Neustart nach {file_counter} Dateien ...")
                    detail_logger.info(
                        f"Geplanter Excel-Neustart nach {file_counter} Dateien")
                    try:
                        excel.Quit()
                    except Exception as _e:
                        detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                    # Excel garantiert toetscheckpoint - Cache-Cleanup
                    # analog 3a/3b/3c/4a (Excel tot vor Neustart).
                    # _kill_specific_excel ist idempotent; _start_excel
                    # ruft es selbst nochmal auf - harmlos.
                    _kill_specific_excel()
                    try:
                        _cleanup_excel_inetcache()
                    except Exception as _e:
                        detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                    try:
                        _cleanup_user_recent()
                    except Exception as _e:
                        detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                    try:
                        _start_excel()
                    except Exception as e_planned:
                        detail_logger.error(
                            f"Geplanter Neustart fehlgeschlagen: {e_planned}")

                # COM Health-Check
                try:
                    _ = excel.Version
                except Exception:
                    pbar.write("  ↻  Excel abgestürzt – starte neu ...")
                    detail_logger.warning("COM Health-Check fehlgeschlagen, starte Excel neu")
                    # Excel abgestuerzt - vor Cache-Cleanup explizit
                    # toten Zustand sicherstellen (idempotent).
                    _kill_specific_excel()
                    try:
                        _cleanup_excel_inetcache()
                    except Exception as _e:
                        detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                    try:
                        _cleanup_user_recent()
                    except Exception as _e:
                        detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                    try:
                        _start_excel()
                    except Exception as e_restart:
                        pbar.write(
                            f"  ✗  Excel-Neustart fehlgeschlagen – "
                            f"verbleibende Dateien werden übersprungen: {e_restart}")
                        log_error(file_path, e_restart)
                        stats["ERROR"] += 1
                        pbar.update(1)
                        processed = sum(stats.values())
                        if total is not None and processed < total:
                            remaining = total - processed
                            stats["SKIPPED"] += remaining
                            detail_logger.warning(
                                f"Excel-Neustart permanent fehlgeschlagen – "
                                f"{remaining} verbleibende Dateien als SKIPPED gezählt")
                            pbar.write(
                                f"  ⚠  {remaining} verbleibende Dateien übersprungen "
                                f"(Excel nicht wiederherstellbar)")
                        else:
                            detail_logger.warning(
                                "Excel-Neustart permanent fehlgeschlagen – "
                                "Anzahl übersprungener Dateien unbekannt (Generator-Modus)")
                        break

                # Memory-Watchdog
                if _check_excel_memory():
                    pbar.write("  ↻  Excel-Memory-Limit erreicht – starte neu ...")
                    try:
                        _start_excel()
                    except Exception as e_mem_restart:
                        detail_logger.error(f"Neustart nach Memory-Limit fehlgeschlagen: {e_mem_restart}")

                # Datei-Verarbeitung mit Retry
                result = "ERROR"
                for attempt in range(FILE_PROCESSING_ATTEMPTS):
                    # Vor der Wiederholung lebt Excel nach einem Timeout
                    # nicht mehr (Watchdog-Kill). Frueher lief Versuch 2
                    # gegen die tote Instanz und scheiterte sicher; der
                    # Lebendtest oben greift erst bei der NAECHSTEN Datei.
                    if attempt > 0 and not _excel_antwortet(excel):
                        pbar.write("  ↻  Excel antwortet nicht – starte neu vor Wiederholung ...")
                        detail_logger.warning(
                            "Excel vor Wiederholung tot (Timeout/Absturz) – Neustart")
                        try:
                            _kill_specific_excel()
                            _start_excel()
                        except Exception as e_retry_start:
                            log_error(file_path, e_retry_start)
                            result = "ERROR"
                            break
                    try:
                        result = replace_fonts_in_workbook(
                            file_path, excel, pbar, font_name,
                            metadata_selected)
                        if result != "ERROR":
                            break
                        if attempt < FILE_PROCESSING_ATTEMPTS - 1:
                            detail_logger.info(
                                f"Wiederholung {attempt+1}/{FILE_PROCESSING_ATTEMPTS-1} für: {file_path}")
                            time.sleep(2)
                    except Exception as e:
                        log_error(file_path, e)
                        result = "ERROR"
                        pbar.write(f"  ✗  KRITISCH: {os.path.basename(file_path)}")
                        break

                stats[result] += 1
                if resume_path and result in RESUME_STATUSES:
                    append_resume(resume_path, file_path)
                pbar.update(1)
                file_counter += 1

                # Fehlerrate-Abbruch
                processed = sum(stats.values())
                if processed >= MIN_FILES_FOR_RATE_CHECK:
                    error_rate = stats["ERROR"] / processed
                    if error_rate > ERROR_RATE_THRESHOLD:
                        pbar.write(
                            f"\n  ⛔  ABBRUCH: Fehlerrate {error_rate:.1%} > "
                            f"Schwellwert {ERROR_RATE_THRESHOLD:.0%}")
                        detail_logger.error(
                            f"Verarbeitung wegen hoher Fehlerrate abgebrochen "
                            f"({stats['ERROR']} von {processed})")
                        break
            else:
                # for-else: Schleife lief ohne break vollstaendig durch.
                run_completed = True

    except Exception as e:
        print(f"\nKRITISCHER FEHLER: Excel konnte nicht gestartet werden: {e}")
        log_error("GLOBAL", e)

    finally:
        addins_restored = False
        if excel is not None:
            try:
                addins_restored = _restore_com_addins(excel)
            except Exception:
                addins_restored = False
            try:
                excel.Quit()
            except Exception as _e:
                detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
        _kill_specific_excel()
        excel_app_global = None
        if not addins_restored:
            try:
                _restore_com_addins_fallback()
            except Exception as _e:
                detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
        time.sleep(2)
        _safe_cleanup_temp()
        # Office-INetCache (Content.MSO) und Windows-Recent (.lnk) ebenfalls
        # aufraeumen - analog zu 3a/3b/3c/4a. Beide best effort.
        try:
            _cleanup_excel_inetcache()
        except Exception as _e:
            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
        try:
            _cleanup_user_recent()
        except Exception as _e:
            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")

    return stats, run_completed

# ==================================================================
# Kommandozeilen-Argumente
# ==================================================================
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Excel Font-Ersetzung – ersetzt Schriftarten in Excel-Dateien",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--auto", action="store_true",
                        help="Trust-Center-Check und interaktive Rückfragen überspringen")
    parser.add_argument("--dir", dest="directory", default=None,
                        help="Zielverzeichnis")
    parser.add_argument("--font", default=None,
                        help=f"Ziel-Schriftart (Standard: {DEFAULT_FONT_NAME})")
    parser.add_argument("--count-first", action="store_true",
                        help="Dateien vorab zählen (ETA-Anzeige)")
    parser.add_argument("--no-meta", action="store_true",
                        help="Metadaten nicht entfernen")
    parser.add_argument("--dry-run", action="store_true",
                        help="Probelauf: listet auf, was passieren würde. "
                             "Excel wird nicht gestartet, nichts geöffnet, "
                             "gespeichert oder gelöscht.")
    parser.add_argument("--resume", dest="resume", default=None,
                        help="Resume-Datei: bereits fertige Dateien überspringen "
                             "(für geplante Tasks)")
    # Rueckwaertskompatibilitaet: Flag wird stillschweigend akzeptiert.
    # Blaetter werden jetzt immer entsperrt (und bleiben entsperrt).
    parser.add_argument("--unprotect-sheets", action="store_true",
                        help=argparse.SUPPRESS)
    return parser.parse_args()

# ==================================================================
# Einstiegspunkt
# ==================================================================
if __name__ == "__main__":
    args = parse_args()
    auto_mode = args.auto

    print()
    print("╔══════════════════════════════════════════════════════════════╗")
    print("║                     EXCEL FONT-ERSETZUNG                     ║")
    print("╚══════════════════════════════════════════════════════════════╝")
    print(f"  Start: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    # Vor dem ersten COM-Zugriff auf bereits laufende Sitzungen hinweisen.
    warn_running_excel(auto_mode)


    # ── Trust-Center-Check ────────────────────────────────────────────
    if not auto_mode:
        issues = check_trust_center_registry()
        if issues:
            print("\n" + "!" * 66)
            print("TRUST-CENTER: Potenzielle Probleme gefunden:")
            for iss in issues:
                print(f"  • {iss}")
            print("!" * 66)

        print("!" * 66)
        print("  WICHTIGER CHECK: TRUST-CENTER-EINSTELLUNGEN")
        print("!" * 66)
        print("  Excel öffnen → Datei → Optionen → Trust Center →")
        print("  'Einstellungen für das Trust Center...'")
        print()
        print("  1. GESCHÜTZTE ANSICHT:")
        print("     ☐ Geschützte Ansicht für Dateien aus dem Internet")
        print("     ☐ Geschützte Ansicht für potenziell unsichere Speicherorte")
        print("     ☐ Geschützte Ansicht für Outlook-Anlagen")
        print("     → ALLE 3 DEAKTIVIEREN")
        print()
        print("  2. MAKROEINSTELLUNGEN:")
        print("     ⦿ Aktivieren von VBA-Makros")
        print("     ☑ Zugriff auf das VBA-Projektobjektmodell vertrauen")
        print()
        print("  3. VERTRAUENSWÜRDIGE SPEICHERORTE:")
        print("     ☑ Vertrauenswürdige Speicherorte im Netzwerk zulassen")
        print("     Folgende Speicherorte hinzufügen:")
        print(f"     • {TEMP_BASE_PATH} (temporärer Arbeitsordner des Skripts)")
        print(r"     • Q:\ oder \\server\dfs (Quelldateien)")
        print("     Jeweils: ☑ Unterordner ... sind ebenfalls vertrauenswürdig")
        print("!" * 66)

        if not ask_yes_no("Wurden diese Einstellungen in Excel vorgenommen?", default_yes=False):
            print("\n❌ Abbruch. Bitte konfigurieren Sie erst das Trust-Center in Excel.")
            sys.exit(0)

    # ── Zielverzeichnis & Schriftart ──────────────────────────────────
    if args.directory:
        start_dir = sanitize_path(args.directory)
        if not os.path.isdir(start_dir):
            print(f"Verzeichnis nicht gefunden: {start_dir}")
            sys.exit(2)
    elif auto_mode:
        print("Im --auto Modus muss --dir angegeben werden.")
        sys.exit(2)
    else:
        start_dir = ask_directory()

    # ── 3-Optionen-Fortschrittsmenue ─────────────────────────────────
    if auto_mode:
        show_progress = True
        count_first   = args.count_first
    elif args.count_first:
        show_progress = True
        count_first   = True
    else:
        show_progress, count_first = ask_progress_mode()

    # ── Schriftart ────────────────────────────────────────────────────
    if args.font:
        NEW_FONT_NAME = args.font
    elif auto_mode:
        NEW_FONT_NAME = DEFAULT_FONT_NAME
    else:
        NEW_FONT_NAME = ask_font()

    # ── Metadaten ─────────────────────────────────────────────────────
    metadata_selected = {}
    if args.no_meta:
        remove_meta = False
    elif auto_mode:
        remove_meta = False
    else:
        remove_meta = ask_yes_no("\nMetadaten entfernen?")

    if remove_meta:
        metadata_selected = ask_metadata_detail()

    # ── Probelauf (Dry-Run) ───────────────────────────────────────────
    if args.dry_run:
        dry_run = True
    elif auto_mode:
        dry_run = False
    else:
        dry_run = ask_yes_no(
            "\nProbelauf (Dry-Run)? Zeigt nur, was passieren würde –\n"
            "  es wird nichts geöffnet, gespeichert oder gelöscht",
            default_yes=False)

    # ── Resume-Datei bestimmen ────────────────────────────────────────
    # CLI --resume hat Vorrang (bewusst persistent fuer geplante Tasks).
    # Ohne --resume wird im interaktiven Modus eine verzeichnisspezifische
    # Resume-Datei angeboten und nach vollstaendigem Lauf geloescht.
    resume_file       = args.resume
    auto_resume_owned = False
    resume_set: set   = set()
    if resume_file:
        resume_set = load_resume_set(resume_file)
    elif not auto_mode:
        candidate = get_auto_resume_path(start_dir)
        existing  = load_resume_set(candidate)
        if existing:
            print(f"\n♻  Resume-Datei eines früheren Laufs gefunden: "
                  f"{len(existing)} bereits verarbeitete Datei(en).")
            if ask_yes_no("  Lauf fortsetzen (J) oder von vorn beginnen (n)?",
                          default_yes=True):
                resume_file = candidate
                resume_set  = existing
            else:
                delete_resume_file(candidate)
                resume_file = candidate
        else:
            resume_file = candidate
        auto_resume_owned = True

    # ── Konfiguration ─────────────────────────────────────────────────
    print()
    print("=" * 66)
    print("  KONFIGURATION")
    print("=" * 66)
    print(f"  Verzeichnis:          {start_dir}")
    print(f"  Neue Schriftart:      {NEW_FONT_NAME}")
    print(f"  Modus:                {'Vorab zählen (ETA)' if count_first else 'Generator (kein ETA)'}")
    print(f"  Fortschrittsbalken:   {'an' if show_progress else 'aus (inline)'}")
    if resume_file and (resume_set or not auto_resume_owned):
        _rlabel = ("automatisch" if auto_resume_owned else resume_file)
        _rcount = f"{len(resume_set)} übersprungen" if resume_set else "neu"
        print(f"  Resume:               {_rcount} ({_rlabel})")
    print( "  Sheets entsperren:    Ja (ohne Passwort, bleiben entsperrt)")
    if metadata_selected and any(metadata_selected.values()):
        print("  Metadaten entfernen:")
        for key, active in metadata_selected.items():
            if active:
                print(f"    * {METADATA_OPTIONS[key]}")
    else:
        print("  Metadaten entfernen:  Nein")
    print("  Verarbeitungsumfang:")
    print("    * Formatvorlagen (Styles)")
    print("    * Tabellenblätter: UsedRange (Zellen), Kommentare")
    print("    * Tabellen (ListObjects), PivotTables")
    print("  NICHT umstellbar:")
    print("    * Bedingte Formatierung – trägt technisch keinen Schriftnamen")
    print("      (Excel lässt dort nur Schnitt, Unterstreichung, Farbe zu)")
    print("    * Kopf- und Fußzeilen (mit & ohne Font-Code)")
    print("    * Eingebettete Diagramme + Diagrammblätter (inkl. DataTable, Point-Labels)")
    print("    * Shapes: Textfelder, Formen, Gruppen (rekursiv), SmartArts")
    print("    * Konvertierung: .xls/.xlt → .xlsx/.xltx (xlsb bleibt erhalten)")
    print("  AV-Scanner-Schutz:")
    print(f"    * Datei-Operationen mit Retry bis {AV_MAX_RETRIES}x")
    print(f"    * Excel-Save mit Timeout {SAVE_TIMEOUT_SECONDS}s")
    print(f"    * Wartezeit nach Temp-Kopie bis {AV_POST_COPY_WAIT_SECONDS:.0f}s")
    print("=" * 66)

    if not auto_mode:
        if not ask_yes_no("\nJetzt starten?"):
            print("Abgebrochen.")
            sys.exit(0)

    # ── Probelauf: ohne Excel, ohne Aenderungen, dann Ende ────────────
    if dry_run:
        dr_start = datetime.now()
        try:
            dry_run_directory(start_dir)
        except Exception as e:
            print(f"\nKRITISCHER FEHLER im Probelauf: {e}")
            log_error("GLOBAL", e)
        print(f"\n  Dauer:       {str(datetime.now() - dr_start).split('.')[0]}")
        print(f"  Detail-Log:  {os.path.abspath(DETAILED_LOG_FILE)}")
        if not auto_mode:
            try:
                input("\nBeliebige Taste drücken, um das Fenster zu schließen ...")
            except EOFError:
                pass
        sys.exit(0)

    # ── Excel Smoke-Test (nach Bestaetigung, vor Hauptlauf) ───────────
    # Praxistest, ergaenzend zur Registry-Pruefung weiter oben:
    # Erzeugt eine kurzlebige Excel-Instanz, oeffnet eine minimale .xlsx
    # im TEMP_BASE_PATH und prueft, ob das Open ohne Geschuetzte-Ansicht-
    # Hänger durchläuft. Bei Hänger killt der Watchdog Excel nach 25 s.
    # Ohne diesen Test wuerde ein falsch konfiguriertes Trust-Center die
    # spaeteren Open-Aufrufe pro Datei minutenlang blockieren.
    # Kategorien: "com" = COM-Subsystem-Problem, "xl" = Excel/Trust-Center.
    if not auto_mode:
        print("\nPrüfe COM-Subsystem und Trust-Center ...")
        smoke_ok, smoke_msg, smoke_cat = test_trust_center_smoke(timeout=25.0)
        if not smoke_ok:
            print()
            print("=" * 66)
            if smoke_cat == "com":
                print("  ⚠  COM-SUBSYSTEM NICHT VERFÜGBAR")
                print("=" * 66)
                print(f"  Grund: {smoke_msg}")
                print()
                print("  Hinweise:")
                print("    • Das ist KEIN Trust-Center-Problem, sondern ein COM-/pywin32-Problem.")
                print("    • pywin32 nicht korrekt registriert:")
                print("        python Scripts/pywin32_postinstall.py -install")
                print("    • Bei EXE-Build: --hidden-import pythoncom + --hidden-import pywintypes prüfen")
                print("    • Antivirus / EDR blockiert COM-Aufrufe auf EXCEL.EXE")
                print("=" * 66)
                file_logger.error(f"COM-Subsystem nicht verfuegbar: {smoke_msg}")
                log_warn_text = "Abbruch durch Benutzer nach COM-Subsystem-Fehler."
                log_ignore_text = "COM-Subsystem-Fehler vom Benutzer ignoriert – Fortsetzung."
            else:
                print("  ⚠  EXCEL-SMOKE-TEST FEHLGESCHLAGEN")
                print("=" * 66)
                print(f"  Grund: {smoke_msg}")
                print()
                print("  Mögliche Ursachen:")
                print("    • Temp-Ordner nicht als vertrauenswürdiger Speicherort eingetragen")
                print(f"      ({TEMP_BASE_PATH})")
                print("    • Geschützte Ansicht für unsichere Speicherorte noch aktiv")
                print("    • Excel/Office-Profil beschädigt oder fehlende Desktop-Ordner")
                print("=" * 66)
                file_logger.error(f"Excel-Smoke-Test fehlgeschlagen: {smoke_msg}")
                log_warn_text = "Abbruch durch Benutzer nach Smoke-Test-Warnung."
                log_ignore_text = "Smoke-Test-Warnung vom Benutzer ignoriert – Fortsetzung."

            if not ask_yes_no("\nTrotzdem fortfahren? (Timeout-Risiko pro Datei!)",
                              default_yes=False):
                print("Abgebrochen.")
                file_logger.error(log_warn_text)
                sys.exit(0)
            file_logger.error(log_ignore_text)
        else:
            detail_logger.info("Excel-Smoke-Test erfolgreich.")
            print("  → Smoke-Test OK.")

    signal.signal(signal.SIGINT,  _signal_handler)
    try:
        signal.signal(signal.SIGTERM, _signal_handler)
    except Exception:
        pass
    # SIGBREAK wie in 3b: Strg+Pause kommt unter Windows als SIGBREAK, nicht
    # als SIGINT - ohne Handler endete der Prozess sofort, _restore_com_addins
    # lief nie und die per Connect=False dauerhaft (LoadBehavior in HKCU)
    # abgeschalteten COM-Add-Ins blieben aus. Die Laufzeitbibliothek meldet
    # auch das Schliessen des Konsolenfensters als SIGBREAK; dort gibt
    # Windows dem Prozess aber nur wenige Sekunden und beendet ihn, sobald
    # der Konsolen-Handler zurueckkehrt - der Python-Handler kommt dann
    # meist nicht mehr zum Zug. Fuer das Fensterschliessen ist das also nur
    # ein Versuch, kein Schutz.
    if hasattr(signal, "SIGBREAK"):
        try:
            signal.signal(signal.SIGBREAK, _signal_handler)
        except Exception as _e:
            detail_logger.debug(f"SIGBREAK-Handler nicht setzbar: {_e!r}")

    # --- Einzelinstanz-Schutz ---
    # Diese Sperre gab es bisher nur in 3a-3c. Ohne sie konnten zwei
    # Laeufe gleichzeitig ueber denselben Bestand gehen und sich
    # gegenseitig die Temp-Kopien und Zieldateien wegziehen.
    # Freigabe ueber atexit, damit sie auch bei sys.exit greift.
    _sperre = None
    if gem is not None:
        _sperre = gem.Einzelinstanz("4b_ersetze_font_in_excel")
        if not _sperre.belegen():
            print(_sperre.hinweis())
            sys.exit(1)
        import atexit
        atexit.register(_sperre.freigeben)

    pythoncom.CoInitialize()

    # Restore-Privilegien (Admin-Kontext) für ACL/Owner-Erhalt aktivieren.
    _enable_restore_privileges()

    start_time   = datetime.now()
    stats_result = None
    exit_code    = 0

    try:
        if gem is not None and hasattr(gem, "recent_wurzel_hinzufuegen"):
            gem.recent_wurzel_hinzufuegen(start_dir)
        stats_result, run_completed = process_directory(
            start_dir, NEW_FONT_NAME, metadata_selected,
            count_first, show_progress=show_progress,
            resume_path=resume_file, resume_set=resume_set)

        # Auto-Resume-Datei nach vollstaendigem Lauf entfernen (eine vom
        # Nutzer via --resume uebergebene Datei bleibt unangetastet).
        if auto_resume_owned and run_completed and resume_file:
            delete_resume_file(resume_file)
            detail_logger.info("Auto-Resume-Datei nach vollständigem Lauf gelöscht.")
    except Exception as e:
        print(f"\nKRITISCHER FEHLER: {e}")
        log_error("GLOBAL", e)
        exit_code = 2
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    duration = datetime.now() - start_time

    print()
    print("=" * 66)
    print("  VERARBEITUNG ABGESCHLOSSEN")
    print("=" * 66)
    print(f"  Ende:  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Dauer: {str(duration).split('.')[0]}")
    print()

    if stats_result:
        total_files = sum(stats_result.values())
        print("  STATISTIK:")
        print(f"    [OK] Erfolgreich:      {stats_result.get('SUCCESS', 0)}")
        print(f"    [~ ] Teilweise:        {stats_result.get('PARTIAL', 0)}")
        print(f"    [→ ] Übersprungen:     {stats_result.get('SKIPPED', 0)}")
        print(f"    [!!] Fehler:           {stats_result.get('ERROR', 0)}")
        print(f"         GESAMT:            {total_files}")

        if stats_result.get("ERROR", 0) > 0 and exit_code == 0:
            exit_code = 1

        write_run_summary(stats_result, start_dir, NEW_FONT_NAME,
                          metadata_selected, duration)

    if not auto_mode:
        print()
        print("!" * 66)
        print("HINWEIS: Bitte setzen Sie die Trust-Center-Einstellungen in Excel")
        print("wieder auf die ursprünglichen Werte zurück.")
        print("!" * 66)

    print()
    print(f"  Fehler-Log:  {os.path.abspath(LOG_FILE)}")
    print(f"  Detail-Log:  {os.path.abspath(DETAILED_LOG_FILE)}")
    if RUN_SUMMARY_FILE and os.path.exists(RUN_SUMMARY_FILE):
        print(f"  Run-Summary: {os.path.abspath(RUN_SUMMARY_FILE)}")
    if CONVERSIONS_CSV and os.path.exists(CONVERSIONS_CSV):
        print("  Konvertierungs-CSV (alte → neue Pfade):")
        print(f"               {os.path.abspath(CONVERSIONS_CSV)}")
    if (not auto_mode and auto_resume_owned and resume_file
            and os.path.exists(resume_file)):
        print(f"  Resume-Log:  {os.path.abspath(resume_file)}")
        print("               (erhalten – Neustart setzt dort fort)")
    print("=" * 66)
    print()

    if _is_tty() and not auto_mode:
        try:
            input("Beliebige Taste drücken, um das Fenster zu schließen ...")
        except EOFError:
            pass

    _cleanup_windows_temp()
    sys.exit(exit_code)
