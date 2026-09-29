# ==================================================================
# WORD FONT-ERSETZUNG
# ==================================================================
# ANLEITUNG:
# 1. Python installieren: https://www.python.org/
# 2. In der Konsole (CMD): pip install pywin32 tqdm psutil
# 3. Falls pywin32 erstmals installiert: python Scripts/pywin32_postinstall.py -install
# 4. In der Konsole (CMD): python 4a_ersetze_font_in_word.py
#
# FUNKTIONEN (interaktiv abgefragt):
#   - Probelauf (Dry-Run): zeigt nur, was passieren wuerde (insbesondere
#     welche .doc/.dot konvertiert und deren Originale ersetzt wuerden)
#   - Resume: bei Abbruch grosser Laeufe setzt der naechste Lauf dort
#     fort, wo der vorige aufgehoert hat (Resume-Datei im Skript-Ordner)
#   - Konvertierungs-CSV: jede .doc/.dot-Konvertierung wird mit altem
#     und neuem Pfad in einer CSV protokolliert (Nachweisliste)
#
# WICHTIG: WORD SICHERHEITSEINSTELLUNGEN (TRUST CENTER)
# ------------------------------------------------------------------
# Word öffnen → Datei → Optionen → Trust Center →
# "Einstellungen für das Trust Center..."
#
# 1. GESCHÜTZTE ANSICHT:
#    ☐ Geschützte Ansicht für Dateien aus dem Internet aktivieren
#    ☐ Geschützte Ansicht für Dateien an potenziell unsicheren Speicherorten aktivieren
#    ☐ Geschützte Ansicht für Outlook-Anlagen aktivieren
#    → ALLE 3 DEAKTIVIEREN
#
# 2. MAKROEINSTELLUNGEN:
#    ⦿ Alle Makros aktivieren
#    ☑ Zugriff auf das VBA-Projektobjektmodell vertrauen
#    (Erforderlich für korrekte .docm/.docx-Unterscheidung)
#
# 3. VERTRAUENSWÜRDIGE SPEICHERORTE:
#    ☑ Vertrauenswürdige Speicherorte im Netzwerk zulassen (nur bei Netzwerkquellen)
#    Folgende Speicherorte hinzufügen:
#    • %LOCALAPPDATA%\Dateimigration-Arbeitskopien\4a_ersetze_font_in_word
#      (zwingend – temporärer Arbeitsordner des Skripts für Long-Path-Dateien)
#    • Bei Quellen auf Q:\, R:\, G:\ oder \\server\dfs zusätzlich:
#      jeweiligen Pfad eintragen, ☑ Unterordner sind ebenfalls vertrauenswürdig
#    • Bei Quellen auf Desktop / Downloads / lokal: keine zusätzliche Eintragung
#
# NACH ABSCHLUSS: Trust-Center-Einstellungen wieder zurücksetzen!
#
# HINWEIS ZU doc.RemovePersonalInformation:
#   Das Skript setzt dieses dokument-persistente Property bewusst auf False,
#   damit ausgewählte Metadaten-Typen gezielt über RemoveDocumentInformation()
#   entfernt werden (und nicht pauschal beim Speichern). Der Flag-Wert wird in
#   der Datei gespeichert. Dies ist eine bewusste Design-Entscheidung. Eine
#   saubere Roundtrip-Lösung via SaveAs2 + Flag-Restore ist nicht implementiert.
#
# HINWEIS FÜR PyInstaller-Build (Umwandlung in .exe):
#   # Einmalig: PyInstaller installieren
#   pip install pyinstaller
#
#   # Build als ein einzeiliger Befehl (am sichersten in jeder Shell):
#   python -m PyInstaller --onefile --noupx --noconfirm --clean --console --icon="python_icon.ico" --name "4a_ersetze_font_in_word" --hidden-import pywintypes --hidden-import pythoncom --hidden-import win32com --hidden-import win32com.client --hidden-import win32api --hidden-import win32file --hidden-import win32process --hidden-import win32con --hidden-import winreg --hidden-import psutil --hidden-import tqdm 4a_ersetze_font_in_word.py
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
#     Code-Signing nach Build empfohlen für Enterprise-Rollout.
#
#   Bitness (32 vs. 64 Bit) bei gemischten Office-Umgebungen:
#     PyInstaller erzeugt eine EXE in der Architektur des verwendeten Python-
#     Interpreters (64-Bit-Python -> x64-EXE; 32-Bit-Python -> x86-EXE). Word
#     registriert seinen COM-Server als LocalServer32 (out-of-process,
#     WINWORD.EXE); das Windows-COM-Subsystem marshallt Aufrufe zwischen
#     x64-Aufrufer und x86-Server (und umgekehrt) automatisch über DCOM/LRPC.
#     Eine x64-EXE arbeitet daher auch mit 32-Bit-Word zusammen (und umgekehrt).
#     Bitness-Match liefert minimal bessere Performance, ist aber nicht
#     erforderlich. 64-Bit-Python ist eine sichere Default-Wahl, da modernes
#     Office (2019+, M365) standardmäßig x64 ist.
#
# Stand: 12.06.2026
# ==================================================================

import os
import re
import sys
import csv
import argparse
import stat
import shutil
import hashlib
import logging
import signal
import time
import uuid
import winreg
import threading
import zipfile
import tempfile
from datetime import datetime
from typing import Optional, Tuple

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

# ==================================================================
# Konsolen-Encoding auf UTF-8
# ==================================================================
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

# ==================================================================
# COM-Konstanten
# ==================================================================
COM_TRUE  = -1
COM_FALSE =  0

WD_FORMAT_DOCX = 12
WD_FORMAT_DOCM = 13
WD_FORMAT_DOTX = 14
WD_FORMAT_DOTM = 15

MSO_GROUP    = 6
MSO_SMARTART = 24

MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

FILE_READ_ATTRIBUTES  = 0x0080
FILE_WRITE_ATTRIBUTES = 0x0100

# WdRemoveDocInfoType - Nummern und Namen aus der Typbibliothek
# MSWORD.OLB (Office 2024) ausgelesen (pythoncom.LoadTypeLib, ohne Word).
# Die frueheren Beschriftungen waren ab Typ 4 verschoben (4 hiess
# "Dokumenteigenschaften", 8 "Kopf-/Fußzeilen-Metadaten", 10 "XML-Teile",
# 18 "Veröffentlichungsinfos", 19 "Persönliche Informationen" ...). Dadurch
# liefen 18 (@-Erwähnungen) und 19 (Dokumentaufgaben) als "sichere" Typen
# mit - obwohl beide Kommentarinhalte veraendern und die Kommentar-
# entfernung abgelehnt sein konnte.
METADATA_TYPES = {
    1:  "Kommentare (Inhalte)",                          # wdRDIComments
    2:  "Überarbeitungen",                               # wdRDIRevisions
    3:  "Versionen",                                     # wdRDIVersions
    4:  "Persönliche Informationen (Autorennamen …)",     # wdRDIRemovePersonalInformation
    5:  "E-Mail-Kopfzeilen",                             # wdRDIEmailHeader
    6:  "Routing-Informationen",                         # wdRDIRoutingSlip
    7:  "Zur Überprüfung senden",                        # wdRDISendForReview
    8:  "Dokumenteigenschaften (Autor, Titel …)",        # wdRDIDocumentProperties
    9:  "Dokumentvorlage",                               # wdRDITemplate
    10: "Dokumentarbeitsbereich",                        # wdRDIDocumentWorkspace
    14: "Dokumentserver-Eigenschaften",                  # wdRDIDocumentServerProperties
    15: "Dokumentverwaltungsrichtlinie",                 # wdRDIDocumentManagementPolicy
    16: "Inhaltstyp",                                    # wdRDIContentType
    18: "@-Erwähnungen (in Kommentaren)",                # wdRDIAtMentions
    19: "Dokumentaufgaben (in Kommentaren)",             # wdRDIDocumentTasks
    20: "Dokumentintelligenz",                           # wdRDIDocumentIntelligence
}

# Typen, die Kommentarinhalte veraendern: nur zusammen mit der
# ausdruecklichen Zustimmung zur Kommentarentfernung (Typ 1).
METADATA_TYPES_KOMMENTARE = (1, 18, 19)

DOWNLOADS_GUID = "{374DE290-123F-4565-9164-39C4925E467B}"

# ==================================================================
# Logging  (Dual-Logger mit Zeitstempel-Suffix pro Run)
# ==================================================================
def _resolve_script_directory() -> str:
    """
    Ermittelt das Verzeichnis, in dem das Skript bzw. die kompilierte EXE liegt.
    - PyInstaller --onefile / --onedir: sys.executable zeigt auf die EXE
    - Normaler Python-Aufruf: __file__ zeigt auf die .py-Datei
    Fallback auf Arbeitsverzeichnis, falls beides fehlschlaegt.
    """
    try:
        if getattr(sys, "frozen", False):
            return os.path.dirname(os.path.abspath(sys.executable))
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()

_SCRIPT_DIR       = _resolve_script_directory()
_RUN_TIMESTAMP    = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE          = os.path.join(_SCRIPT_DIR, f"4a_ersetze_font_in_word_{_RUN_TIMESTAMP}.log")
DETAILED_LOG_FILE = os.path.join(_SCRIPT_DIR, f"4a_ersetze_font_in_word_detailed_{_RUN_TIMESTAMP}.log")
RUN_SUMMARY_FILE  = os.path.join(_SCRIPT_DIR, "4a_ersetze_font_in_word_last_run.txt")
CONVERSIONS_CSV   = os.path.join(_SCRIPT_DIR, f"4a_ersetze_font_in_word_konvertierungen_{_RUN_TIMESTAMP}.csv")
RESUME_FILE_PREFIX = os.path.join(_SCRIPT_DIR, "4a_ersetze_font_in_word_resume")


def _setup_logging() -> tuple:
    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    fh = logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8-sig")
    fh.setFormatter(fmt)
    fl = logging.getLogger("FileLogger")
    fl.setLevel(logging.ERROR)
    fl.addHandler(fh)

    dh = logging.FileHandler(DETAILED_LOG_FILE, mode="w", encoding="utf-8-sig")
    dh.setFormatter(fmt)
    dl = logging.getLogger("DetailLogger")
    dl.setLevel(logging.DEBUG)
    dl.addHandler(dh)

    return fl, dl


file_logger, detail_logger = _setup_logging()


# ==================================================================
# User-Shell-Folder-Lookup (Registry mit Fallback)
# ==================================================================

def _get_user_shell_folder(value_name: str, fallback_subdir: str) -> Optional[str]:
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        ) as key:
            raw, _ = winreg.QueryValueEx(key, value_name)
        expanded = os.path.expandvars(raw)
        if os.path.isdir(expanded):
            return expanded
    except Exception as e:
        detail_logger.debug(
            f"User-Shell-Folder-Registry-Lookup '{value_name}' fehlgeschlagen: {e}")

    fallback = os.path.join(os.path.expanduser("~"), fallback_subdir)
    if os.path.isdir(fallback):
        return fallback
    return None


def _get_documents_base() -> str:
    p = _get_user_shell_folder("Personal", "Documents")
    if p:
        return p
    return os.path.join(os.path.expanduser("~"), "Documents")


# ==================================================================
# Konfiguration
# ==================================================================
NEW_FONT_NAME     = "Arial"
# Gemeinsamer Arbeitsordner der Office-Skripte. Bis 29.09.2026 lag er in
# "Dokumente"; bei OneDrive-Ordnersicherung wanderte so jede Arbeitskopie
# in die Cloud. %LOCALAPPDATA% wird nie umgeleitet.
ARBEITS_BASIS     = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or _get_documents_base(),
    "Dateimigration-Arbeitskopien")
TEMP_BASE_PATH    = os.path.join(ARBEITS_BASIS, "4a_ersetze_font_in_word")
TEMP_PROCESS_PATH = None
MAX_PATH_LEN      = 240
WINDOWS_MAX_PATH  = 260

WORD_RESTART_INTERVAL = 200

# Verzeichnisnamen, die nie betreten werden: geloeschte Dokumente im
# $RECYCLE.BIN wuerden sonst mitkonvertiert (und ihr Original geloescht);
# ~snapshot/.snapshot sind read-only NAS-Schattenkopien.
EXCLUDE_DIR_NAMES = {"$recycle.bin", "system volume information",
                     "~snapshot", ".snapshot"}

FILE_ATTRIBUTE_REPARSE_POINT = 0x0400

# Watchdog-Timeouts fuer COM-Aufrufe in der Hauptverarbeitung. Werte sind
# grosszuegig bemessen, um grosse Dateien (50+ MB, viele eingebettete
# Objekte) nicht voreilig abzuschiessen, fangen aber haengende
# Modal-Dialoge zuverlaessig ab (DRM-Popup, Verknuepfungs-Update,
# Makro-Sicherheitswarnung trotz DisplayAlerts=False, beschaedigte
# Tabellenstrukturen). Smoke-Test (25s) bleibt unveraendert, weil dort
# nur eine winzige Test-Datei verarbeitet wird.
WORD_OPEN_TIMEOUT  = 180.0   # Sekunden fuer Documents.Open
WORD_SAVE_TIMEOUT  = 240.0   # Sekunden fuer SaveAs2 / Save (kann bei
                             # grossen Dateien laenger dauern als Open)

# AV-Retry: Wartezeiten in Sekunden, exponentielles Backoff (gesamt ~17.5 s)
AV_RETRY_DELAYS = (0.5, 1.0, 2.0, 4.0, 4.0, 4.0, 4.0)
TS_RETRY_DELAYS = AV_RETRY_DELAYS

_preserved_temp_files: set = set()

# ==================================================================
# Globale Word-Referenz und PID (für Cleanup)
# ==================================================================
word_app_global: Optional[win32com.client.CDispatch] = None
word_pid_global: Optional[int] = None
# Prozess-Erstellungszeit der Skript-Word-Instanz. Schuetzt zusammen mit
# dem Prozessnamen-Check vor PID-Recycling: Stirbt Word waehrend einer
# langen Watchdog-Wartezeit und vergibt Windows die PID an eine neue,
# vom Benutzer geoeffnete Word-Sitzung, unterscheidet nur die
# Erstellungszeit die beiden Prozesse.
word_create_time_global: Optional[float] = None


# ==================================================================
# Hilfsfunktionen
# ==================================================================

def log_error(file_path: str, exc: Exception) -> None:
    file_logger.error(f"Datei: {file_path}\n  -> {exc}\n")
    detail_logger.error(f"Datei: {file_path}\n  -> {exc}\n")


def _rmtree_onerror(func, path, exc_info):
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception as _e:
        detail_logger.debug(f"_rmtree_onerror: Exception verworfen: {_e!r}")


# Whitelist fuer den Windows-Temp-Cleanup: NUR Word-eigene Reste
# (~$ Owner-Dateien, ~DF*.tmp, VBE-Reste) und eigene Skript-Artefakte.
# NIEMALS pauschal leeren: Fremdprozesse legen aktive Daten ohne Lock
# in %TEMP% ab; blindes Loeschen zerstoert sie. gen_py (pywin32-COM-
# Cache) und excel8.0 gehoeren bewusst NICHT hierher: gen_py wird von
# parallel laufenden COM-Skripten (3a/3b/3c) aktiv genutzt, excel8.0
# ist ein Excel-Artefakt.
_WINDOWS_TEMP_WHITELIST_PREFIXES = (
    "~$", "~df", "vbe",
    "4a_ersetze_font_in_word",
)


def _is_whitelisted_temp_entry(name: str) -> bool:
    nl = name.lower()
    return any(nl.startswith(p) for p in _WINDOWS_TEMP_WHITELIST_PREFIXES)


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


def _safe_cleanup_temp_process() -> None:
    # Bereinigt den Run-spezifischen TEMP_PROCESS_PATH (UUID-Subordner).
    # Wenn RESCUE_*-Dateien oder vom Skript "preserved" Dateien drinliegen,
    # bleibt der Ordner stehen (Hinweis fuer den Benutzer).
    if not TEMP_PROCESS_PATH or not os.path.exists(TEMP_PROCESS_PATH):
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
            try:
                detail_logger.warning(
                    f"GESCHÜTZTE DATEIEN IM TEMP-ORDNER – Ordner wird NICHT gelöscht!\n"
                    f"  Pfad: {TEMP_PROCESS_PATH}\n"
                    f"  Dateien ({len(all_protected)}):\n"
                    + "\n".join(f"    * {f}" for f in all_protected)
                    + "\n  → Bitte manuell sichern, dann Ordner löschen.")
            except Exception as _e:
                detail_logger.debug(f"_safe_cleanup_temp_process: Exception verworfen: {_e!r}")
            print("\n  GESCHÜTZTE DATEIEN vorhanden – Temp-Ordner NICHT gelöscht:")
            print(f"   {TEMP_PROCESS_PATH}")
            for f in all_protected:
                print(f"   * {f}")
            print("   → Bitte manuell sichern, dann Ordner löschen.")
        else:
            shutil.rmtree(TEMP_PROCESS_PATH, onerror=_rmtree_onerror)
    except Exception as e:
        try:
            detail_logger.warning(f"Aufräumen fehlgeschlagen: {e}")
        except Exception as _e:
            detail_logger.debug(f"_safe_cleanup_temp_process: Exception verworfen: {_e!r}")


def _cleanup_word_inetcache() -> None:
    """Raeumt den Word-INetCache-Ordner auf
    (%LOCALAPPDATA%\\Microsoft\\Windows\\INetCache\\Content.Word).

    Word legt dort Zwischenspeicher-Kopien ab, die ueber lange Laeufe
    massiv anwachsen koennen. Best effort: nur Dateien, die nicht gelockt
    sind, werden geloescht; gesperrte Dateien werden uebersprungen und
    auf DEBUG-Level geloggt.

    Zusaetzlich wird Content.MSO mitgenommen (Office-uebergreifender
    Cache, enthaelt auch Word-bezogene Daten z.B. fuer eingebettete
    Excel-Tabellen).

    DARF NUR AUFGERUFEN WERDEN, WENN DAS SKRIPT-WORD NICHT LAEUFT
    (also nach Quit + _kill_specific_word, oder am Skriptende).
    """
    local_appdata = os.environ.get("LOCALAPPDATA")
    if not local_appdata:
        return

    cache_subdirs = [
        os.path.join("Microsoft", "Windows", "INetCache", "Content.Word"),
        os.path.join("Microsoft", "Windows", "INetCache", "Content.MSO"),
    ]

    for sub in cache_subdirs:
        cache_dir = os.path.join(local_appdata, sub)

        # SICHERHEITS-CHECK: Pfad muss auf den erwarteten Suffix enden.
        if not cache_dir.lower().endswith(sub.lower()):
            continue
        if not os.path.isdir(cache_dir):
            continue

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
                            detail_logger.debug(f"_cleanup_word_inetcache: Exception verworfen: {_e!r}")
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
            continue

        if files_deleted > 0 or files_skipped > 0:
            mb = bytes_freed / (1024 * 1024)
            detail_logger.info(
                f"Word-INetCache aufgeraeumt ({os.path.basename(cache_dir)}): "
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


# ==================================================================
# Startwarnung: laufende Word-Sitzungen
# ==================================================================
def find_running_word_pids() -> list:
    """PIDs aller Word-Prozesse des angemeldeten Benutzers."""
    pids = []
    try:
        me = psutil.Process().username().lower().split("\\")[-1].split("@")[0]
    except Exception:
        return pids
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            nm = proc.info.get("name") or ""
            if "WINWORD.EXE" not in nm.upper():
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


def warn_running_word(auto_mode: bool = False) -> None:
    """Warnt vor bereits laufenden Word-Sitzungen.

    Word wird per COM automatisiert. Laeuft bereits eine Sitzung des
    Anwenders, hat das zwei Auswirkungen, die man vorher kennen sollte:

      1. Die Automatisierung kann sich an die vorhandene Sitzung haengen.
         Warnhinweise werden dann abgeschaltet und das Fenster ausgeblendet
         - fuer den Anwender sieht das aus, als sei Word abgestuerzt.
      2. Die Aufraeumroutinen des Skripts koennen die eigene Instanz nicht
         mehr sicher von der fremden unterscheiden.

    Die Sitzung wird NICHT beendet (dafuer sorgt die Momentaufnahme der
    fremden Prozess-IDs beim Start), aber ein sauberer Lauf setzt ein
    geschlossenes Word voraus.
    """
    pids = find_running_word_pids()
    if not pids:
        return
    pids_str = ", ".join(str(p) for p in pids)

    if auto_mode:
        detail_logger.warning(
            f"Automatikmodus: laufende Word-Sitzungen (PID: {pids_str}) - "
            f"sie werden geschuetzt, aber nicht geschlossen."
        )
        return

    print()
    print("=" * 66)
    print("  WARNUNG: Word laeuft bereits")
    print("=" * 66)
    print(f"  Gefundene Word-Prozesse (PID): {pids_str}")
    print()
    print("  Ihre Sitzung wird vom Skript NICHT beendet. Waehrend des Laufs")
    print("  kann sie aber ausgeblendet werden und Warnhinweise sind")
    print("  abgeschaltet - das wirkt wie ein Absturz.")
    print("  Ausserdem laesst sich die eigene Automatisierungs-Instanz dann")
    print("  nicht mehr zuverlaessig von Ihrer Sitzung unterscheiden.")
    print()
    print("  EMPFEHLUNG: Word jetzt schliessen und das Skript neu starten.")
    print("=" * 66)
    print()
    if not ask_yes_no("Trotzdem fortfahren?"):
        print("Abgebrochen. Bitte Word schliessen und neu starten.")
        sys.exit(0)


def _alle_word_pids() -> set:
    """PIDs aller WINWORD.EXE-Prozesse (jeder Benutzer) - Momentaufnahme."""
    found = set()
    try:
        for p in psutil.process_iter(["pid", "name"]):
            try:
                if (p.info["name"] or "").upper() == "WINWORD.EXE":
                    found.add(p.info["pid"])
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
    except Exception as _e:
        detail_logger.debug(f"_alle_word_pids: Exception verworfen: {_e!r}")
    return found


def _ermittle_word_pid(word, pids_vorher: set) -> Optional[int]:
    """PID der gerade per DispatchEx gestarteten Word-Instanz.

    Frueher: win32process.GetWindowThreadProcessId(word.Hwnd). Word.
    Application HAT KEINE Hwnd-Eigenschaft - nachgeprueft in MSWORD.OLB
    (pythoncom.LoadTypeLib, Office 2024): _Application kennt 212 Mitglieder,
    darunter kein Hwnd; nur Window.Hwnd existiert. Der Aufruf warf also
    immer AttributeError: im Smoke-Test blieb word_pid None und der
    Watchdog konnte einen Haenger nie beenden.

    Jetzt: 1. Momentaufnahme vor/nach DispatchEx (wie im Hauptlauf) -
    genau EINE neue WINWORD-PID. 2. Rueckfall bei mehrdeutiger Aufnahme:
    Documents.Add + ActiveWindow.Hwnd (Window.Hwnd laut Typbibliothek
    vorhanden; das Laufzeitverhalten einer unsichtbaren Instanz ist hier
    ohne Office ungemessen), nur gueltig, wenn die PID NEU ist.
    Rueckgabe None = unbekannt - der Aufrufer muss dann sicher abbrechen."""
    neu = _alle_word_pids() - set(pids_vorher)
    if len(neu) == 1:
        pid = next(iter(neu))
        detail_logger.debug(f"Word-PID via psutil-Snapshot ermittelt: {pid}")
        return pid
    detail_logger.warning(
        f"PID-Snapshot nicht eindeutig ({len(neu)} neue Prozesse) "
        f"– versuche Fenster-Handle eines Hilfsdokuments")
    doc = None
    try:
        doc = word.Documents.Add()
        hwnd = doc.ActiveWindow.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(int(hwnd))
        if (pid and pid not in pids_vorher
                and psutil.Process(pid).name().upper() == "WINWORD.EXE"):
            detail_logger.debug(f"Word-PID via Window.Hwnd ermittelt: {pid}")
            return pid
        detail_logger.warning(f"Window.Hwnd lieferte keine neue Word-PID ({pid})")
    except Exception as e_hwnd:
        detail_logger.warning(f"PID-Ermittlung via Window.Hwnd fehlgeschlagen: {e_hwnd}")
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"_ermittle_word_pid: Exception verworfen: {_e!r}")
    return None


def _kill_specific_word() -> None:
    global word_pid_global, word_create_time_global
    if word_pid_global is None:
        return
    try:
        proc = psutil.Process(word_pid_global)
        try:
            proc_name = proc.name().upper()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            proc_name = ""
        if proc_name != "WINWORD.EXE":
            detail_logger.debug(
                f"PID {word_pid_global} ist nicht WINWORD.EXE "
                f"(sondern '{proc_name}') – kein Kill (PID-Recycling)")
            return
        if word_create_time_global is not None:
            try:
                if proc.create_time() != word_create_time_global:
                    detail_logger.debug(
                        f"PID {word_pid_global} hat abweichende "
                        f"Erstellungszeit – kein Kill (PID-Recycling)")
                    return
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                return
        proc.kill()
        proc.wait(timeout=3)
        detail_logger.debug(f"Skript-Word beendet (PID {word_pid_global})")
    except psutil.NoSuchProcess:
        pass
    except Exception as e:
        detail_logger.warning(
            f"Fehler beim Beenden von PID {word_pid_global}: {e}")
    finally:
        word_pid_global = None
        word_create_time_global = None


def _word_process_alive() -> bool:
    """Prueft, ob die Skript-Word-Instanz noch lebt (PID + Name +
    Erstellungszeit). Liefert True, wenn keine PID bekannt ist
    (Liveness unbekannt -> kein unnoetiger Neustart)."""
    if word_pid_global is None:
        return True
    try:
        proc = psutil.Process(word_pid_global)
        if not proc.is_running() or proc.name().upper() != "WINWORD.EXE":
            return False
        if (word_create_time_global is not None
                and proc.create_time() != word_create_time_global):
            return False
        return True
    except psutil.NoSuchProcess:
        return False
    except Exception:
        return True


# HRESULTs, die einen toten COM-Server anzeigen:
#   0x800706BA RPC_S_SERVER_UNAVAILABLE ("Der RPC-Server ist nicht verfuegbar")
#   0x80010108 RPC_E_DISCONNECTED ("Das aufgerufene Objekt wurde von den
#   Clients getrennt")
_RPC_DEAD_HRESULTS = {-2147023174, -2147417848}


def _is_rpc_dead_error(exc: Exception) -> bool:
    try:
        if isinstance(exc, pythoncom.com_error):
            hresult = exc.args[0] if exc.args else None
            if hresult in _RPC_DEAD_HRESULTS:
                return True
    except Exception as _e:
        detail_logger.debug(f"_is_rpc_dead_error: Exception verworfen: {_e!r}")
    msg = str(exc).lower()
    if "0x800706ba" in msg or "0x80010108" in msg:
        return True
    return "rpc" in msg and ("nicht verfügbar" in msg or "unavailable" in msg)


# Word-Laufzeitfehler 5408 ("Das Kennwort ist falsch. Word kann das
# Dokument nicht oeffnen.") kommt per COM als DISP_E_EXCEPTION mit
# excepinfo-scode 0x800A1520 (= 0x800A0000 + 5408).
_WD_KENNWORT_SCODES = {-2146823904}


def _ist_kennwortfehler(exc: Exception, pfade) -> bool:
    """Passwortschutz erkennen - vorrangig am COM-Fehlercode.

    Frueher nur per Stichwort im gesamten Fehlertext. Word nennt darin
    aber den Pfad: "Kennwortliste.docx" oder ein Ordner "Geschützt" machten
    JEDEN Fehler zum SKIPPED_PASSWORD - und der steht in RESUME_STATUSES,
    die Datei wurde also dauerhaft uebersprungen. Jetzt: 1. excepinfo-scode;
    2. Rueckfall auf Stichworte erst, nachdem alle Pfade, Dateinamen und
    Ordnernamen der beteiligten Dateien aus dem Text entfernt sind."""
    try:
        if isinstance(exc, pywintypes.com_error):
            excepinfo = exc.args[2] if len(exc.args) > 2 else None
            if (excepinfo and len(excepinfo) > 5
                    and excepinfo[5] in _WD_KENNWORT_SCODES):
                return True
    except Exception as _e:
        detail_logger.debug(f"_ist_kennwortfehler: Exception verworfen: {_e!r}")
    text = str(exc).lower()
    teile = set()
    for p in pfade:
        if not p:
            continue
        sauber = _strip_long_path_prefix(str(p))
        # str(com_error) ist die repr des Tupels - Rueckstriche doppelt.
        teile.update({str(p), sauber, os.path.basename(sauber),
                      sauber.replace("\\", "\\\\")})
        teile.update(t for t in sauber.replace("/", "\\").split("\\") if len(t) >= 3)
    # Laengste zuerst, damit ein Ordnername nicht einen Teil des vollen
    # Pfads zerschneidet, bevor dieser als Ganzes entfernt ist.
    for t in sorted(teile, key=len, reverse=True):
        text = text.replace(t.lower(), " ")
    return any(kw in text for kw in (
        "password", "passwort", "kennwort",
        "protected", "geschützt",
    ))


# ==================================================================
# Watchdog (Timeout-Schutz fuer COM-Aufrufe)
# ==================================================================

def _word_call_with_watchdog(call_label: str, timeout: float, word_pid,
                              call_fn, word_create_time=None):
    """
    Watchdog-Wrapper fuer COM-Calls (analog zu 3a/3b/3c).
    Fuehrt call_fn() aus und beendet den Word-Prozess (per psutil.kill),
    falls der Aufruf nicht innerhalb von 'timeout' Sekunden zurueckkehrt.

    Hintergrund:
      Trust-Center-Probleme oder Geschuetzte-Ansicht-Dialoge in einer
      unsichtbaren Word-Instanz koennen den Aufrufer ohne diesen Schutz
      unbegrenzt blockieren. Mit Watchdog liefert der Aufruf nach
      'timeout' Sekunden eine TimeoutError zurueck, der Prozess wird hart
      beendet und der Smoke-Test/Hauptlauf kann sauber reagieren.

    call_label: Klartext fuer Fehlermeldungen.
    word_pid:   PID der Word-Instanz (von win32process.GetWindowThreadProcessId).
    call_fn:    Lambda mit dem eigentlichen COM-Aufruf.
    word_create_time: psutil-Erstellungszeit der Word-Instanz. Wenn gesetzt,
                wird vor dem Kill validiert, dass die PID noch zu genau
                diesem Prozess gehoert (Schutz vor PID-Recycling auf eine
                neue, vom Benutzer geoeffnete Word-Sitzung).
    """
    done_event   = threading.Event()
    timeout_flag = [False]
    # Schloss + Fertig-Kennzeichen gegen ein schmales, aber echtes
    # Zeitfenster: der Rueckgabewert von call_fn() steht fest, BEVOR das
    # finally done_event setzt. Laeuft der Timeout genau dazwischen ab,
    # toetet der Waechter Word, obwohl der Aufruf gelungen ist - der
    # Aufrufer bekaeme ein Ergebnis und arbeitete danach mit einer toten
    # COM-Instanz weiter. Schloss ALLEIN genuegt nicht (nachgemessen:
    # 43 -> 30 von 300 Faellen); erst die unteilbare Pruefung auf dem
    # Erfolgspfad unten macht Toeten und Erfolg eindeutig (0 von 300).
    state_lock   = threading.Lock()
    completed    = [False]

    def _watchdog():
        if not done_event.wait(timeout):
            with state_lock:
                if completed[0]:
                    return          # Aufruf war bereits fertig
                timeout_flag[0] = True
            if word_pid:
                # Process-Name- UND Erstellungszeit-Check vor kill()
                # schuetzen vor PID-Recycling. Besonders kritisch hier
                # wegen der grosszuegigen Timeouts (bis 240s bei Save):
                # in dieser langen Wartezeit kann Word anderweitig sterben
                # (AV-Kill, OOM, User-Kill), Windows die PID freigeben und
                # an einen unbeteiligten Prozess vergeben - auch an eine
                # NEUE Word-Sitzung des Benutzers, die der Name-Check
                # allein nicht erkennen wuerde.
                try:
                    proc = psutil.Process(word_pid)
                    if ("winword" in proc.name().lower()
                            and (word_create_time is None
                                 or proc.create_time() == word_create_time)):
                        proc.kill()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
                except Exception as _e:
                    detail_logger.debug(f"_watchdog: Exception verworfen: {_e!r}")

    wd_thread = threading.Thread(target=_watchdog, daemon=True)
    wd_thread.start()

    try:
        _ergebnis = call_fn()
        with state_lock:
            if timeout_flag[0]:
                # Der Waechter hat bereits zugeschlagen: die Instanz ist
                # tot, das Ergebnis damit unbrauchbar. Als Timeout melden,
                # statt dem Aufrufer einen Erfolg vorzuspiegeln.
                raise TimeoutError(f"Word Timeout bei {call_label}")
            completed[0] = True
        return _ergebnis
    except Exception as e:
        done_event.set()
        if timeout_flag[0]:
            raise TimeoutError(f"Word Timeout bei {call_label}")
        raise e
    finally:
        done_event.set()
        wd_thread.join(timeout=1.0)


# ==================================================================
# Trust-Center Smoke-Test
# ==================================================================
# Hintergrund:
#   Word kann beim Open einer .docx in einen unsichtbaren Geschuetzte-
#   Ansicht-Dialog laufen, wenn der Pfad nicht als vertrauenswuerdiger
#   Speicherort eingetragen ist. In COM ist dieser Dialog nicht sichtbar -
#   der Aufruf haengt einfach. Pro Datei waeren das minutenlange Hänger.
#   Der Smoke-Test prueft daher VOR dem Hauptlauf, ob Open im Temp-Ordner
#   ohne Hänger funktioniert. Der Watchdog beendet Word nach 25 s hart;
#   dann liefert der Test einen klaren Fehlertext statt eines Freezes.

def _build_minimal_docx(target_path: str) -> None:
    """Baut eine minimale gueltige .docx per zipfile (kein COM noetig)."""
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        '</Relationships>'
    )
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body><w:p/></w:body>'
        '</w:document>'
    )

    if os.path.exists(target_path):
        try:
            os.remove(target_path)
        except Exception as _e:
            detail_logger.debug(f"_build_minimal_docx: Exception verworfen: {_e!r}")

    with zipfile.ZipFile(target_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels",         rels)
        zf.writestr("word/document.xml",   document)


def test_trust_center_smoke(timeout: float = 25.0) -> Tuple[bool, str, str]:
    """
    Word-Smoke-Test fuer 4a:
      Erzeugt eine kurzlebige Word-Instanz, baut eine minimale Test-.docx
      im TEMP_BASE_PATH, oeffnet sie und schliesst sie wieder. Im
      Erfolgsfall wird Word sauber beendet; bei Timeout per Watchdog hart
      gekillt.
      Rueckgabe: (ok, fehlertext, kategorie)
        kategorie: "ok"  - Test erfolgreich
                   "com" - COM-Subsystem-Problem (z.B. CO_E_NOTINITIALIZED)
                   "wd"  - Word/Trust-Center-Problem
    """
    try:
        os.makedirs(TEMP_BASE_PATH, exist_ok=True)
    except Exception as e:
        return False, f"Temp-Ordner nicht erstellbar ({TEMP_BASE_PATH}): {e}", "wd"

    test_path = os.path.join(
        TEMP_BASE_PATH, f"trustcheck_{uuid.uuid4().hex}.docx"
    )

    try:
        _build_minimal_docx(test_path)
    except Exception as e:
        return False, f"Test-DOCX konnte nicht erstellt werden: {e}", "wd"

    word             = None
    word_pid         = None
    word_create_time = None
    doc              = None
    pythoncom.CoInitialize()  # COM-Init fuer Main-Thread (sonst CO_E_NOTINITIALIZED bei DispatchEx)
    try:
        # Eigene kurzlebige Word-Instanz fuer den Test.
        pids_vorher = _alle_word_pids()
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except pythoncom.com_error as e:
            # HRESULT 0x800401F0 (-2147221008) = CO_E_NOTINITIALIZED -> COM-Subsystem, kein Trust-Center
            hresult = e.args[0] if e.args else None
            if hresult == -2147221008:
                return False, (f"COM-Subsystem nicht initialisiert "
                               f"(HRESULT 0x800401F0 CO_E_NOTINITIALIZED): {e}"), "com"
            return False, f"Word-Instanz konnte nicht gestartet werden: {e}", "wd"
        except Exception as e:
            return False, f"Word-Instanz konnte nicht gestartet werden: {e}", "wd"

        # PID fuer den Watchdog-Kill bei Timeout. Frueher per word.Hwnd -
        # das gibt es an Word.Application nicht (AttributeError, PID immer
        # None, Watchdog wirkungslos). Siehe _ermittle_word_pid.
        word_pid = _ermittle_word_pid(word, pids_vorher)
        if word_pid:
            try:
                word_create_time = psutil.Process(word_pid).create_time()
            except Exception:
                word_create_time = None
        else:
            # Ohne PID koennte der Watchdog einen Haenger beim Open nicht
            # beenden - der Test haengt dann unbegrenzt. Deshalb das Open
            # gar nicht erst versuchen, sondern als Fehlschlag melden.
            return False, ("PID der Word-Instanz nicht ermittelbar - ohne sie "
                           "kann der Watchdog einen Haenger nicht beenden"), "wd"

        # Word lautlos konfigurieren.
        try: word.Visible = False
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try: word.DisplayAlerts = 0   # wdAlertsNone
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try: word.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try: word.Options.UpdateLinksAtOpen   = False
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try: word.Options.DoNotPromptForConvert = True
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")

        # --- Open mit Watchdog ---
        def _do_open():
            return word.Documents.Open(
                test_path,
                ConfirmConversions = COM_FALSE,
                ReadOnly           = COM_TRUE,
                AddToRecentFiles   = COM_FALSE,
            )

        try:
            doc = _word_call_with_watchdog(
                "Documents.Open (Smoke-Test)",
                timeout, word_pid, _do_open, word_create_time
            )
        except TimeoutError:
            # Watchdog hat Word beendet; word ist tot.
            word = None
            return False, (f"TIMEOUT nach {timeout:.0f}s bei Documents.Open "
                           "- Trust Center vermutlich nicht konfiguriert "
                           "oder Geschuetzte Ansicht aktiv"), "wd"
        except Exception as e:
            return False, f"Documents.Open fehlgeschlagen: {e}", "wd"

        if doc is None:
            return False, "Documents.Open lieferte kein Document-Objekt", "wd"

        return True, "", "ok"
    finally:
        if doc is not None:
            try: doc.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        if word is not None:
            try: word.Quit(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
            # Falls Quit nicht greift: harter Kill ueber gemerkte PID.
            if word_pid:
                try:
                    p = psutil.Process(word_pid)
                    if (p.is_running()
                            and p.name().upper() == "WINWORD.EXE"
                            and (word_create_time is None
                                 or p.create_time() == word_create_time)):
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


def _eigene_word_mru_entfernen(word) -> None:
    # Word traegt die Arbeitskopien trotz AddToRecentFiles=False in seine
    # Liste "Zuletzt verwendet" ein (am 29.09.2026 in der Registry
    # gesehen: sechs Eintraege auf 4a-Arbeitsordner). Entfernt werden nur
    # Eintraege im bearbeiteten Verzeichnis oder im Arbeitsordner; ohne
    # _gemeinsam.py bleibt die Liste unangetastet.
    ist_eigen = getattr(gem, "recent_ist_eigen", None)
    if ist_eigen is None or word is None:
        return
    try:
        for i in range(word.RecentFiles.Count, 0, -1):
            try:
                eintrag = word.RecentFiles(i)
                if ist_eigen(os.path.join(eintrag.Path, eintrag.Name)):
                    eintrag.Delete()
            except Exception as _e:
                detail_logger.debug(f"_eigene_word_mru_entfernen: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_eigene_word_mru_entfernen: Exception verworfen: {_e!r}")


def _restore_word_settings(
    word,
    original_screen_updating: bool,
    original_pagination: bool,
    original_spell_check: bool,
    original_grammar_check: bool,
    original_background_save: bool,
    original_update_links: bool = False,
    original_no_prompt_convert: bool = False,
) -> None:
    _eigene_word_mru_entfernen(word)
    try:
        word.ScreenUpdating = original_screen_updating
    except Exception as _e:
        detail_logger.debug(f"_restore_word_settings: Exception verworfen: {_e!r}")
    try:
        word.Options.Pagination = original_pagination
    except Exception as _e:
        detail_logger.debug(f"_restore_word_settings: Exception verworfen: {_e!r}")
    try:
        word.Options.CheckSpellingAsYouType = original_spell_check
    except Exception as _e:
        detail_logger.debug(f"_restore_word_settings: Exception verworfen: {_e!r}")
    try:
        word.Options.CheckGrammarAsYouType = original_grammar_check
    except Exception as _e:
        detail_logger.debug(f"_restore_word_settings: Exception verworfen: {_e!r}")
    try:
        word.Options.BackgroundSave = original_background_save
    except Exception as _e:
        detail_logger.debug(f"_restore_word_settings: Exception verworfen: {_e!r}")
    # Word-Options sind profil-persistent - die beiden Dialog-
    # Unterdrueckungen daher ebenfalls auf den Ausgangswert zuruecksetzen.
    try:
        word.Options.UpdateLinksAtOpen = original_update_links
    except Exception as _e:
        detail_logger.debug(f"_restore_word_settings: Exception verworfen: {_e!r}")
    try:
        word.Options.DoNotPromptForConvert = original_no_prompt_convert
    except Exception as _e:
        detail_logger.debug(f"_restore_word_settings: Exception verworfen: {_e!r}")


# ==================================================================
# Signal-Handler (Ctrl+C / SIGTERM – SIGTERM auf Windows kosmetisch)
# ==================================================================
def _signal_handler(sig, frame) -> None:
    print("\n\n*** ABBRUCH durch Benutzer – raeume auf ...")

    _kill_specific_word()

    global word_app_global
    if word_app_global is not None:
        word_app_global = None

    _safe_cleanup_temp_process()

    # Best effort Cache-Cleanup beim Strg+C-Abbruch. Schnell genug fuer
    # einen Signal-Handler-Kontext (nur lokale Datei-Loeschungen).
    try:
        _cleanup_word_inetcache()
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

    sys.exit(1)


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
# AV-tolerante Datei-Operationen (Retry mit exponentiellem Backoff)
# ==================================================================

def _is_av_lock_error(exc: Exception) -> bool:
    if isinstance(exc, PermissionError):
        return True
    err_no  = getattr(exc, "errno",   None)
    win_err = getattr(exc, "winerror", None)
    if err_no in (5, 13, 32):
        return True
    if win_err in (5, 32, 33):
        return True
    return False


def _robust_copy(src: str, dst: str, op_name: str = "copy") -> None:
    last_exc: Optional[Exception] = None
    for attempt, delay in enumerate([0.0, *AV_RETRY_DELAYS]):
        if delay > 0:
            time.sleep(delay)
        try:
            shutil.copy2(src, dst)
            if attempt > 0:
                detail_logger.debug(
                    f"{op_name}: erfolgreich nach {attempt} Retry(s)")
            return
        except Exception as e:
            last_exc = e
            if not _is_av_lock_error(e):
                raise
            detail_logger.debug(
                f"{op_name}: Versuch {attempt + 1} blockiert ({e}), warte ...")
    if last_exc is not None:
        raise last_exc


def _robust_move(src: str, dst: str, op_name: str = "move") -> None:
    last_exc: Optional[Exception] = None
    for attempt, delay in enumerate([0.0, *AV_RETRY_DELAYS]):
        if delay > 0:
            time.sleep(delay)
        try:
            shutil.move(src, dst)
            if attempt > 0:
                detail_logger.debug(
                    f"{op_name}: erfolgreich nach {attempt} Retry(s)")
            return
        except Exception as e:
            last_exc = e
            if not _is_av_lock_error(e):
                raise
            detail_logger.debug(
                f"{op_name}: Versuch {attempt + 1} blockiert ({e}), warte ...")
    if last_exc is not None:
        raise last_exc


def _robust_remove(path: str, op_name: str = "remove") -> None:
    last_exc: Optional[Exception] = None
    for attempt, delay in enumerate([0.0, *AV_RETRY_DELAYS]):
        if delay > 0:
            time.sleep(delay)
        try:
            os.remove(path)
            if attempt > 0:
                detail_logger.debug(
                    f"{op_name}: erfolgreich nach {attempt} Retry(s)")
            return
        except FileNotFoundError:
            return
        except Exception as e:
            last_exc = e
            if not _is_av_lock_error(e):
                raise
            detail_logger.debug(
                f"{op_name}: Versuch {attempt + 1} blockiert ({e}), warte ...")
    if last_exc is not None:
        raise last_exc


def _replace_file_atomic(src: str, dst: str,
                         op_name: str = "replace",
                         sd=None) -> None:
    """Ersetzt dst durch src OHNE Truncate-Fenster (analog 5_OCR_PDF).

    HINTERGRUND: shutil.move faellt auf Windows auf copy2 + unlink
    zurueck, sobald dst existiert oder der Move cross-volume laeuft -
    copy2 trunkiert die Zieldatei sofort auf 0 Bytes. Ein harter
    Abbruch in diesem Fenster (Netzwerk-Drop, AV-Scanner, Prozess-Kill)
    hinterliesse eine halbe Datei. Stattdessen wird src zunaechst als
    dst + '.tmp_new' AUF DAS ZIELVOLUME kopiert und dann per
    os.replace() atomar uebergeschoben: Das Ziel ist zu jedem Zeitpunkt
    entweder die alte oder die neue Datei. Backup-/Restore-Logik und
    .bak_-Reste (frueher _replace_file_with_backup) entfallen damit.
    """
    stage = dst + ".tmp_new"
    try:
        # Die Staging-Kopie gehoert INS try: stand sie davor, blieb bei
        # einem Abbruch mitten im Kopieren (Netzwerk-Drop, Strg+C) eine
        # halbe '<name>.tmp_new' in der Ablage liegen - das Aufraeumen
        # unten erfasste nur Fehler NACH der Kopie. Mit Attrappe
        # nachgemessen (copy2 bricht nach der Haelfte ab): vorher blieb
        # die halbe .tmp_new liegen, nachher nicht.
        _robust_copy(src, stage, op_name=f"{op_name}-Staging")
        # Schreibschutz auf dem Ziel entfernen: os.replace scheitert auf
        # NTFS an read-only-Zieldateien mit PermissionError.
        try:
            if os.path.exists(dst):
                dst_mode = os.stat(dst).st_mode
                if not (dst_mode & stat.S_IWRITE):
                    os.chmod(dst, dst_mode | stat.S_IWRITE)
        except Exception as _e:
            detail_logger.debug(f"_replace_file_atomic: Exception verworfen: {_e!r}")
        last_exc: Optional[Exception] = None
        for attempt, delay in enumerate([0.0, *AV_RETRY_DELAYS]):
            if delay > 0:
                time.sleep(delay)
            try:
                os.replace(stage, dst)
                last_exc = None
                if attempt > 0:
                    detail_logger.debug(
                        f"{op_name}: os.replace erfolgreich nach {attempt} Retry(s)")
                break
            except FileNotFoundError:
                raise
            except Exception as e:
                last_exc = e
                if not _is_av_lock_error(e):
                    raise
                detail_logger.debug(
                    f"{op_name}: os.replace Versuch {attempt + 1} blockiert ({e}), warte ...")
        if last_exc is not None:
            raise last_exc
    # BaseException, nicht Exception: Der Signal-Handler beendet sich mit
    # sys.exit(), und SystemExit erbt von BaseException. Bei Strg+C mitten im
    # Ersetzen blieb die Staging-Kopie '<name>.docx.tmp_new' sonst in der
    # Ablage liegen - kein Aufraeumpfad des Skripts erfasst sie je wieder.
    # Das abschliessende 'raise' reicht den Abbruch unveraendert weiter.
    except BaseException:
        try:
            if os.path.exists(stage):
                try:
                    os.chmod(stage, stat.S_IWRITE)
                except Exception as _e:
                    detail_logger.debug(f"_replace_file_atomic: Exception verworfen: {_e!r}")
                _robust_remove(stage, op_name=f"{op_name}-Staging-Cleanup")
        except Exception as _e:
            detail_logger.debug(f"_replace_file_atomic: Exception verworfen: {_e!r}")
        raise
    try:
        os.remove(src)
    except FileNotFoundError:
        pass
    except Exception as e_src:
        detail_logger.debug(f"{op_name}: Quelldatei nicht loeschbar ({e_src}): {src}")
    _apply_security_descriptor(_strip_long_path_prefix(dst), sd)


# ==================================================================
# Pfad-Hilfsfunktionen
# ==================================================================

def sanitize_path(raw: str) -> str:
    path = raw.strip().strip('"').strip("'")
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


def _strip_long_path_prefix(path: str) -> str:
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\"):
        return path[4:]
    return path


def _get_unique_target_path(target_path: str) -> str:
    if not os.path.exists(prepare_long_path(target_path)):
        return target_path
    base, ext = os.path.splitext(target_path)
    candidate = f"{base} (konvertiert){ext}"
    if not os.path.exists(prepare_long_path(candidate)):
        return candidate
    i = 2
    while True:
        candidate = f"{base} (konvertiert {i}){ext}"
        if not os.path.exists(prepare_long_path(candidate)):
            return candidate
        i += 1


# ==================================================================
# Resume-Datei (Neustart-Faehigkeit grosser Laeufe)
# ==================================================================

def get_resume_file_path(target_dir: str) -> str:
    key = hashlib.md5(
        os.path.normcase(os.path.abspath(target_dir)).encode("utf-8")
    ).hexdigest()[:12]
    return f"{RESUME_FILE_PREFIX}_{key}.txt"


def load_resume_set(resume_path: str) -> set:
    if not os.path.exists(resume_path):
        return set()
    try:
        with open(resume_path, "r", encoding="utf-8") as fh:
            return {line.strip() for line in fh if line.strip()}
    except Exception as e:
        detail_logger.warning(f"Resume-Datei nicht lesbar: {e}")
        return set()


def append_resume(resume_path: str, file_path: str) -> None:
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


# Statuswerte, die in die Resume-Datei eingetragen werden: dauerhaft
# erledigte bzw. dauerhaft nicht verarbeitbare Dateien. Transiente
# Zustaende (gesperrt) und Fehler werden beim naechsten Lauf erneut
# versucht.
RESUME_STATUSES = ("SUCCESS", "UNCHANGED", "SKIPPED_ENCRYPTED", "SKIPPED_PASSWORD")


# ==================================================================
# Konvertierungs-CSV (Nachweisliste .doc/.dot -> .docx/.dotx)
# ==================================================================


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

def _log_conversion(old_path: str, new_path: str) -> None:
    _protokoll(new_path or old_path, "konvertiert", "OK",
               f"aus {old_path}")
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


# ==================================================================
# Run-Summary
# ==================================================================

def write_run_summary(stats: dict, directory: str, font_name: str,
                      metadata_types: list, start_time, end_time) -> None:
    try:
        total = (stats.get("SUCCESS", 0) + stats.get("UNCHANGED", 0)
                 + stats.get("SKIPPED", 0) + stats.get("ERROR", 0))
        with open(RUN_SUMMARY_FILE, "w", encoding="utf-8-sig") as fh:
            fh.write("WORD FONT-ERSETZUNG – LETZTER LAUF\n")
            fh.write("=" * 60 + "\n")
            fh.write(f"Verzeichnis:    {directory}\n")
            fh.write(f"Schriftart:     {font_name}\n")
            fh.write(f"Metadaten:      "
                     f"{('Ja (' + str(len(metadata_types)) + ' Typen)') if metadata_types else 'Nein'}\n")
            fh.write(f"Start:          {start_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            fh.write(f"Ende:           {end_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            fh.write(f"Dauer:          {str(end_time - start_time).split('.')[0]}\n")
            fh.write("\n")
            fh.write(f"Erfolgreich:    {stats.get('SUCCESS', 0)}\n")
            fh.write(f"Unverändert:    {stats.get('UNCHANGED', 0)}\n")
            fh.write(f"Übersprungen:   {stats.get('SKIPPED', 0)}\n")
            for label, key in (("  davon gesperrt:          ", "SKIPPED_LOCKED"),
                               ("  davon OLE-verschlüsselt: ", "SKIPPED_ENCRYPTED"),
                               ("  davon passwortgeschützt: ", "SKIPPED_PASSWORD")):
                if stats.get(key, 0) > 0:
                    fh.write(f"{label}{stats[key]}\n")
            fh.write(f"Fehler:         {stats.get('ERROR', 0)}\n")
            fh.write(f"Gesamt:         {total}\n")
    except Exception as e:
        detail_logger.warning(f"Run-Summary nicht schreibbar: {e}")


# ==================================================================
# OLE-Verschlüsselungserkennung  (.doc/.dot)
# ==================================================================

def _is_doc_encrypted(file_path: str) -> bool:
    try:
        with open(prepare_long_path(file_path), "rb") as f:
            data = f.read(65536)
        if len(data) < 8:
            return False
        if not (data[0] == 0xD0 and data[1] == 0xCF and
                data[2] == 0x11 and data[3] == 0xE0):
            return False

        utf16_enc_pkg  = "EncryptedPackage".encode("utf-16-le")
        utf16_enc_info = "EncryptionInfo".encode("utf-16-le")
        if utf16_enc_pkg in data or utf16_enc_info in data:
            return True

        # Klassische .doc-Verschluesselung (XOR/RC4) traegt keinen der
        # obigen Stream-Namen, sondern setzt das fEncrypted-Bit im FIB
        # (Byte-Offset 0x0B des WordDocument-Streams, Maske 0x01). Der
        # Stream beginnt in der Praxis an einer Sektorgrenze; wIdent =
        # 0xA5EC (little-endian EC A5) markiert den FIB-Anfang.
        for off in range(512, min(len(data), 65536) - 12, 512):
            if data[off] == 0xEC and data[off + 1] == 0xA5:
                if data[off + 11] & 0x01:
                    return True
                break

        ascii_content = data.decode("ascii", errors="replace")
        return ("EncryptedPackage" in ascii_content
                or "EncryptionInfo" in ascii_content)
    except Exception:
        return False


# ==================================================================
# Interaktiver Start
# ==================================================================

def ask_directory() -> str:
    desktop_path   = _get_user_shell_folder("Desktop",       "Desktop")
    downloads_path = _get_user_shell_folder(DOWNLOADS_GUID,  "Downloads")

    print("\nZielverzeichnis auswählen:")
    print("  [1] Q:\\")
    print("  [2] R:\\")
    print("  [3] G:\\Geteilte Ablagen")
    print("  [4] G:\\Meine Ablage")
    print("  [5] \\\\server\\dfs\\dm")
    print(f"  [6] Desktop          ({desktop_path or 'nicht gefunden'})")
    print(f"  [7] Downloads        ({downloads_path or 'nicht gefunden'})")
    print("  [8] Eigenen Pfad eingeben")
    print()
    print("  Hinweis: Bei Start 'als Administrator' sind Netzlaufwerks-")
    print("  Buchstaben (Q:, R:) oft nicht verbunden – dann Option [5]")
    print("  (UNC-Pfad) verwenden.")
    print()
    while True:
        try:
            choice = input("Auswahl [1-8]: ").strip()
        except (EOFError, RuntimeError, OSError):
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
            if desktop_path is None:
                print("  Desktop-Ordner konnte nicht ermittelt werden.")
                continue
            path = desktop_path
        elif choice == "7":
            if downloads_path is None:
                print("  Downloads-Ordner konnte nicht ermittelt werden.")
                continue
            path = downloads_path
        elif choice == "8":
            try:
                raw = input("Pfad eingeben: ")
            except (EOFError, RuntimeError, OSError):
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
            print("     (Bei Start 'als Administrator' sind Netzlaufwerke oft")
            print("      nicht verbunden – dann Option [5] UNC-Pfad verwenden.)")
        elif choice in ("3", "4"):
            print(f"  Pfad nicht erreichbar: {path}")
            print("     (Google Drive for Desktop aktiv?)")
        elif choice == "5":
            print(f"  Netzwerkpfad nicht erreichbar: {path}")
            print("     (VPN aktiv? Netzwerkverbindung prüfen)")
        elif choice in ("6", "7"):
            print(f"  Pfad nicht erreichbar: {path}")
        else:
            print(f"  Verzeichnis nicht gefunden: '{path}'")


def ask_font() -> str:
    try:
        val = input(f"\nZiel-Schriftart [Standard: {NEW_FONT_NAME}]: ").strip()
    except (EOFError, RuntimeError, OSError):
        return NEW_FONT_NAME
    return val if val else NEW_FONT_NAME


def _is_tty() -> bool:
    """Haengt stdin an einer echten Konsole? (wie in 4b)"""
    try:
        return sys.stdin.isatty()
    except Exception:
        return False


def _warte_auf_taste(auto_mode: bool = False) -> None:
    """Abschliessendes 'Beliebige Taste' - nur wenn wirklich jemand zusieht.

    Ohne Konsole (pythonw, Taskplaner 'Unabhaengig von der Benutzeranmeldung')
    ist sys.stdin None und input() wirft RuntimeError('lost sys.stdin') - KEIN
    EOFError, der frueher allein abgefangen wurde. Das Skript endete dann nach
    vollstaendig geleisteter Arbeit mit einem Traceback und Exit-Code 1.
    Haengt stdin an einer offenen Pipe, blockierte der Prozess dauerhaft - und
    weil die Einzelinstanz-Sperre erst per atexit faellt, blockierte er damit
    auch jeden weiteren geplanten Lauf.
    """
    if auto_mode or not _is_tty():
        return
    try:
        input("\nBeliebige Taste drücken, um das Fenster zu schließen ...")
    except (EOFError, RuntimeError, OSError):
        pass


def ask_yes_no(prompt: str, default_yes: bool = True) -> bool:
    hint = "J/n" if default_yes else "j/N"
    while True:
        try:
            answer = input(f"{prompt} [{hint}]: ").strip().lower()
        except (EOFError, RuntimeError, OSError):
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
        except (EOFError, RuntimeError, OSError):
            return (False, False)
        if choice in ("", "1"):
            return (False, False)
        if choice == "2":
            return (True, True)
        if choice == "3":
            return (True, False)
        print("  Bitte 1, 2 oder 3 eingeben.")


def ask_metadata_detail() -> list:
    print()
    print("  Folgende Metadaten-Kategorien werden entfernt:")
    safe_types = [k for k in METADATA_TYPES
                  if k != 2 and k not in METADATA_TYPES_KOMMENTARE]
    for t in safe_types:
        print(f"    wdRDI {t:2d}: {METADATA_TYPES[t]}")

    selected = list(safe_types)

    print()
    if ask_yes_no("  Außerdem: Kommentar-Inhalte (Typ 1, samt @-Erwähnungen 18 "
                  "und Aufgaben 19) entfernen?"):
        selected.extend(METADATA_TYPES_KOMMENTARE)
    if ask_yes_no("  Außerdem: Überarbeitungen/Tracked Changes (Typ 2) entfernen?"):
        selected.append(2)

    return sorted(selected)


# ==================================================================
# Datei-Generator  (netzwerktauglich, Long-Path-fähig)
# ==================================================================

def file_generator(directory: str):
    extensions = {".docx", ".docm", ".dotx", ".dotm", ".doc", ".dot"}
    start = os.path.abspath(directory)
    # Eigenen Arbeitsordner ausnehmen: liegt das Ziel z.B. auf Documents,
    # wuerde der Generator sonst die eigenen Stage-Dateien einsammeln,
    # waehrend Word sie offen hat.
    temp_base_nc = os.path.normcase(os.path.abspath(TEMP_BASE_PATH))
    queue = [start]
    while queue:
        current_dir = queue.pop()
        scan_dir = prepare_long_path(current_dir)
        try:
            with os.scandir(scan_dir) as it:
                entries = list(it)
        except (PermissionError, OSError) as e_dir:
            detail_logger.warning(
                f"Verzeichnis übersprungen ({type(e_dir).__name__}): "
                f"{current_dir} – {e_dir}")
            continue
        for entry in entries:
            try:
                clean_path = _strip_long_path_prefix(entry.path)
                if entry.is_symlink():
                    continue
                if entry.is_dir(follow_symlinks=False):
                    # Papierkorb/Systemordner/NAS-Snapshots nie betreten.
                    if entry.name.lower() in EXCLUDE_DIR_NAMES:
                        continue
                    # Junctions/Mount-Points nicht folgen: is_symlink()
                    # erkennt unter Windows nur echte Symlinks; Junction-
                    # Zyklen fuehrten sonst in eine Endlosschleife.
                    try:
                        attrs = entry.stat(follow_symlinks=False).st_file_attributes
                        if attrs & FILE_ATTRIBUTE_REPARSE_POINT:
                            continue
                    except OSError:
                        pass
                    cp_nc = os.path.normcase(os.path.abspath(clean_path))
                    if cp_nc == temp_base_nc or cp_nc.startswith(temp_base_nc + os.sep):
                        continue
                    queue.append(clean_path)
                elif entry.is_file(follow_symlinks=False):
                    nl = entry.name.lower()
                    if nl.startswith("~$") or nl.startswith("._"):
                        continue
                    if any(nl.endswith(ext) for ext in extensions):
                        yield clean_path
            except (PermissionError, OSError) as e_entry:
                detail_logger.warning(
                    f"Eintrag übersprungen ({type(e_entry).__name__}): "
                    f"{entry.path} – {e_entry}")
                continue


# ==================================================================
# Font-Ersetzung auf einzelnen Ranges/Shapes (COM)
# ==================================================================

# Bis 29.09.2026 setzte 4a die Zielschrift pauschal - auch ueber Text in
# Wingdings/Symbol. Aus einem Haekchen wurde dabei ein "ü". Jetzt merkt sich
# _set_font_on_range vorher alle Stellen in Symbolschrift (Word-Suche mit
# Format) und schreibt ihre vier Schriftnamen danach zurueck.

# Obergrenze je Bereich und Schrift - schuetzt vor einer Endlosschleife,
# falls Word an einer Stelle nicht weitersucht.
SYMBOL_STELLEN_MAX = 20000

# Symbolschriften, die das gerade bearbeitete Dokument ueberhaupt kennt
# (None = unbekannt, dann werden alle gesucht). Gemessen am 29.09.2026: jede
# Suche kostet je Textbereich einen COM-Aufruf; alle 22 Symbolschriften
# ueber Haupttext, Kopf-/Fusszeilen, Tabellen und Textfelder machten aus
# 10 s pro Dokument ueber eine Minute.
_SYMBOL_KANDIDATEN = None


def _symbolschriften_im_dokument(doc):
    """Symbolschriften aus word/fontTable.xml der geoeffneten Datei, oder None."""
    try:
        with zipfile.ZipFile(prepare_long_path(doc.FullName)) as z:
            xml = z.read("word/fontTable.xml").decode("utf-8", errors="ignore")
    except Exception as _e:
        detail_logger.debug(f"_symbolschriften_im_dokument: Exception verworfen: {_e!r}")
        return None
    namen = set(re.findall(r'<w:font\s+w:name="([^"]+)"', xml))
    return sorted(n for n in namen if _ist_symbolschrift(n))


def _symbolstellen_finden(rng) -> list:
    """Stellen in Symbolschrift: (Start, Ende, Name, Ascii, Other, Bi, FarEast)."""
    stellen = []
    try:
        anfang, ende = rng.Start, rng.End
    except Exception:
        return stellen
    kandidaten = (_SYMBOL_KANDIDATEN if _SYMBOL_KANDIDATEN is not None
                  else _symbolschrift_namen())
    for name in kandidaten:
        try:
            r = rng.Duplicate
            f = r.Find
            f.ClearFormatting()
            f.Text = ""
            f.Replacement.Text = ""
            f.Format = True
            f.Forward = True
            f.Wrap = 0                      # wdFindStop
            f.MatchWildcards = False
            f.Font.Name = name
            n = 0
            while f.Execute() and n < SYMBOL_STELLEN_MAX:
                if r.Start >= ende or r.End <= anfang or r.End <= r.Start:
                    break
                s, e = max(r.Start, anfang), min(r.End, ende)
                fo = r.Font
                stellen.append((s, e, fo.Name, fo.NameAscii, fo.NameOther,
                                fo.NameBi, fo.NameFarEast))
                n += 1
                r.Collapse(0)               # wdCollapseEnd
        except Exception as _e:
            detail_logger.debug(f"_symbolstellen_finden: Exception verworfen: {_e!r}")
    return stellen


def _symbolstellen_zurueck(rng, stellen: list) -> None:
    for s, e, name, ascii_, other, bi, fe in stellen:
        try:
            r = rng.Duplicate
            r.SetRange(s, e)
            fo = r.Font
            for attr, wert in (("Name", name), ("NameAscii", ascii_),
                               ("NameOther", other), ("NameBi", bi),
                               ("NameFarEast", fe)):
                if wert:
                    try:
                        setattr(fo, attr, wert)
                    except Exception as _e:
                        detail_logger.debug(f"_symbolstellen_zurueck: Exception verworfen: {_e!r}")
        except Exception as _e:
            detail_logger.debug(f"_symbolstellen_zurueck: Exception verworfen: {_e!r}")


def _set_all_font_names(font, font_name: str) -> None:
    try:
        if font.Name != font_name:
            font.Name = font_name
    except Exception as _e:
        detail_logger.debug(f"_set_all_font_names: Exception verworfen: {_e!r}")
    for attr in ("NameAscii", "NameOther", "NameBi", "NameFarEast"):
        try:
            if getattr(font, attr) != font_name:
                setattr(font, attr, font_name)
        except Exception as _e:
            detail_logger.debug(f"_set_all_font_names: Exception verworfen: {_e!r}")


def _set_font_on_range(rng, font_name: str) -> None:
    try:
        einheitlich = rng.Font.Name
    except Exception:
        einheitlich = ""
    if _ist_symbolschrift(einheitlich):
        return
    # Einheitliche, normale Schrift: darin steckt keine Symbolschrift, die
    # Suche entfaellt. Word meldet bei gemischten Schriften einen leeren Namen.
    stellen = [] if einheitlich else _symbolstellen_finden(rng)
    try:
        _set_all_font_names(rng.Font, font_name)
    except Exception as _e:
        detail_logger.debug(f"_set_font_on_range: Exception verworfen: {_e!r}")
    if stellen:
        _symbolstellen_zurueck(rng, stellen)


def _apply_smartart_node(node, font_name: str) -> None:
    try:
        if not _ist_symbolschrift(node.TextFrame2.TextRange.Font.Name):
            _set_all_font_names(node.TextFrame2.TextRange.Font, font_name)
    except Exception as _e:
        detail_logger.debug(f"_apply_smartart_node: Exception verworfen: {_e!r}")
    try:
        for child in node.ChildNodes:
            try:
                _apply_smartart_node(child, font_name)
            except Exception as _e:
                detail_logger.debug(f"_apply_smartart_node: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_apply_smartart_node: Exception verworfen: {_e!r}")


def _set_font_on_shape(shape, font_name: str) -> None:
    try:
        shape_type = shape.Type
    except Exception:
        return

    if shape_type == MSO_GROUP:
        try:
            for item in shape.GroupItems:
                try:
                    _set_font_on_shape(item, font_name)
                except Exception as _e:
                    detail_logger.debug(f"_set_font_on_shape: Exception verworfen: {_e!r}")
        except Exception as _e:
            detail_logger.debug(f"_set_font_on_shape: Exception verworfen: {_e!r}")
        return

    if shape_type == MSO_SMARTART:
        try:
            for node in shape.SmartArt.Nodes:
                try:
                    _apply_smartart_node(node, font_name)
                except Exception as _e:
                    detail_logger.debug(f"_set_font_on_shape: Exception verworfen: {_e!r}")
        except Exception as _e:
            detail_logger.debug(f"_set_font_on_shape: Exception verworfen: {_e!r}")
        return

    try:
        tef = shape.TextEffect
        if tef is not None and not _ist_symbolschrift(tef.FontName):
            tef.FontName = font_name
    except Exception as _e:
        detail_logger.debug(f"_set_font_on_shape: Exception verworfen: {_e!r}")

    try:
        if shape.HasTextFrame:
            _set_font_on_range(shape.TextFrame.TextRange, font_name)
    except Exception as _e:
        detail_logger.debug(f"_set_font_on_shape: Exception verworfen: {_e!r}")


def _process_footnotes_endnotes(doc, font_name: str) -> None:
    try:
        for fn in doc.Footnotes:
            try:
                _set_font_on_range(fn.Range, font_name)
            except Exception as _e:
                detail_logger.debug(f"_process_footnotes_endnotes: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_process_footnotes_endnotes: Exception verworfen: {_e!r}")
    try:
        for en in doc.Endnotes:
            try:
                _set_font_on_range(en.Range, font_name)
            except Exception as _e:
                detail_logger.debug(f"_process_footnotes_endnotes: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_process_footnotes_endnotes: Exception verworfen: {_e!r}")


def _process_content_controls(doc, font_name: str) -> None:
    try:
        for cc in doc.ContentControls:
            try:
                _set_font_on_range(cc.Range, font_name)
            except Exception as _e:
                detail_logger.debug(f"_process_content_controls: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_process_content_controls: Exception verworfen: {_e!r}")


def _process_revisions(doc, font_name: str) -> None:
    try:
        for rev in doc.Revisions:
            try:
                _set_font_on_range(rev.Range, font_name)
            except Exception as _e:
                detail_logger.debug(f"_process_revisions: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_process_revisions: Exception verworfen: {_e!r}")


# ==================================================================
# Pre-Flight Lock-Check (Sperrdatei-Erkennung anderer Nutzer)
# ==================================================================

ERROR_SHARING_VIOLATION = 32


def _besitzerdatei_kandidaten(path: str):
    """Moegliche Namen der Word-Besitzerdatei '~$...' (siehe is_locked_by_other)."""
    d = os.path.dirname(path)
    b = os.path.basename(path)
    if not d or not b:
        return []
    namen = {"~$" + b}
    if len(b) > 1:
        namen.add("~$" + b[1:])
    if len(b) > 2:
        namen.add("~$" + b[2:])
    return [os.path.join(d, n) for n in namen]


def _eigene_besitzerdatei_entfernen(path: str) -> None:
    """Entfernt die Besitzerdatei, die eine vom Waechter beendete Word-Instanz
    neben path hinterlassen hat. Am Anfang von replace_fonts_in_document_com
    stellt is_locked_by_other sicher, dass es fuer diese Datei KEINE gab - eine
    jetzt vorhandene stammt also von uns. Gemessen 29.09.2026: nach zwei
    Stage-2-Timeouts lagen '~$sswort.docx' und '~$hreibschutz.docx' im
    Bestand; jeder Folgelauf meldete die Dateien als "in Bearbeitung", und die
    Reste waeren mit nach Google Drive gewandert. Eine von einer LEBENDEN
    Word-Sitzung offengehaltene Besitzerdatei laesst sich nicht loeschen und
    bleibt unangetastet."""
    for owner in _besitzerdatei_kandidaten(path):
        try:
            lp = prepare_long_path(owner)
            if os.path.exists(lp):
                try:
                    os.chmod(lp, stat.S_IWRITE)
                except Exception as _e:
                    detail_logger.debug(f"_eigene_besitzerdatei_entfernen: {_e!r}")
                os.remove(lp)
                detail_logger.warning(f"Verwaiste Besitzerdatei entfernt: {owner}")
        except Exception as e:
            detail_logger.warning(f"Besitzerdatei nicht entfernbar: {owner} – {e!r}")


def _kopie_ohne_schreibkennwort(quelle: str, ziel: str) -> None:
    """Kopie einer .docx/.docm/.dotx/.dotm ohne <w:writeProtection>.

    Entscheidung des Anwenders vom 30.09.2026: Schreibkennwoerter werden
    entfernt, nicht uebersprungen. Word selbst kann ein so geschuetztes
    Dokument nur schreibgeschuetzt oeffnen; Stage 2 speichert aber teils per
    Save() an Ort und Stelle. Deshalb auf Dateiebene, in einer Arbeitskopie,
    die wie eine Langpfad-Kopie behandelt und am Ende zurueckgeschoben wird.
    Die Empfehlung "schreibgeschuetzt oeffnen" ohne Kennwort geht dabei mit.
    """
    muster = re.compile(rb"<w:writeProtection\b[^>]*/>|<w:writeProtection\b[^>]*>.*?</w:writeProtection>", re.S)
    with zipfile.ZipFile(prepare_long_path(quelle)) as zi, \
            zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            daten = zi.read(info.filename)
            if info.filename == "word/settings.xml":
                daten = muster.sub(b"", daten)
            zo.writestr(info, daten)


def _ooxml_kennwort_grund(path: str) -> str:
    """'' wenn die OOXML-Datei ohne Kennwort zu oeffnen ist, sonst den Grund.

    Hintergrund (gemessen 29.09.2026 mit Word 2024): Stage 2 oeffnet
    schreibend. Bei einer verschluesselten .docx (Oeffnungskennwort) und bei
    einer mit Schreibkennwort (w:writeProtection mit Kennwort-Hash) wartet
    Word unsichtbar auf das Kennwort - jede solche Datei kostete den vollen
    180-s-Waechter plus Word-Neustart und hinterliess eine Besitzerdatei. 2a
    und 3a erkennen verschluesselte Dateien vorab; hier fehlte das.
    - Verschluesselt: OOXML-Datei im CFB-Container statt ZIP.
    - Schreibkennwort: <w:writeProtection> mit Hash-Attribut in
      word/settings.xml (die reine Empfehlung "schreibgeschuetzt oeffnen"
      ohne Kennwort bleibt erlaubt)."""
    try:
        lp = prepare_long_path(path)
        with open(lp, "rb") as f:
            kopf = f.read(8)
        if kopf.startswith(b"\xD0\xCF\x11\xE0"):
            return "verschluesselt"
        if not kopf.startswith(b"PK"):
            return ""
        with zipfile.ZipFile(lp) as z:
            try:
                settings = z.read("word/settings.xml").decode("utf-8", errors="ignore")
            except KeyError:
                return ""
        m = re.search(r"<w:writeProtection\b[^>]*>", settings)
        if m and re.search(r"w:(hashValue|hash|cryptProviderType|salt)=", m.group(0)):
            return "schreibkennwort"
    except Exception as _e:
        detail_logger.debug(f"_ooxml_kennwort_grund: {_e!r}")
    return ""


def is_locked_by_other(path: str) -> bool:
    try:
        d = os.path.dirname(path)
        b = os.path.basename(path)
        if d and b:
            # Word stellt das Praefix "~$" nicht einfach voran, sondern
            # ERSETZT je nach Namenslaenge das erste bzw. die ersten beiden
            # Zeichen des Dateinamens (aus "Vertrag.docx" wird
            # "~$rtrag.docx"). Nur Excel haengt "~$" vor den vollen Namen.
            # Daher alle drei Varianten pruefen.
            candidates = {"~$" + b}
            if len(b) > 1:
                candidates.add("~$" + b[1:])
            if len(b) > 2:
                candidates.add("~$" + b[2:])
            for cand in candidates:
                owner = os.path.join(d, cand)
                try:
                    if os.path.exists(prepare_long_path(owner)):
                        return True
                except Exception as _e:
                    detail_logger.debug(f"is_locked_by_other: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"is_locked_by_other: Exception verworfen: {_e!r}")

    try:
        h = win32file.CreateFile(
            prepare_long_path(path),
            win32file.GENERIC_READ,
            win32file.FILE_SHARE_READ
            | win32file.FILE_SHARE_WRITE
            | win32file.FILE_SHARE_DELETE,
            None,
            win32file.OPEN_EXISTING,
            0,
            None,
        )
        try:
            h.Close()
        except Exception as _e:
            detail_logger.debug(f"is_locked_by_other: Exception verworfen: {_e!r}")
        return False
    except pywintypes.error as e:
        try:
            if e.winerror == ERROR_SHARING_VIOLATION:
                return True
        except Exception as _e:
            detail_logger.debug(f"is_locked_by_other: Exception verworfen: {_e!r}")
        return False
    except Exception:
        return False


# ==================================================================
# Font-Ersetzung & Metadaten-Entfernung (COM)
# ==================================================================

def replace_fonts_in_document_com(
    file_path: str,
    word_app: win32com.client.CDispatch,
    pbar: tqdm,
    font_name: str,
    metadata_types: list,
) -> str:
    # word_app_global wird hier ggf. auf None gesetzt, wenn der Watchdog
    # die Word-Instanz killt. Aufrufer prueft danach und startet Word neu.
    global word_app_global, _SYMBOL_KANDIDATEN

    pbar.write(f"Prüfe: {os.path.basename(file_path)}")
    detail_logger.info(f"=== Starte: {file_path} ===")

    original_path       = file_path
    is_temp_copy        = False
    was_converted       = False
    schema_migrated     = False
    doc                 = None
    new_path            = None
    temp_stage1_path    = None
    # temp_stage2_path haelt den Pfad zur Stage-2-SaveAs2-Datei, in die
    # die finalen Font-Ersetzungen geschrieben werden (statt direkt ins
    # Original). Erst nach Close + erfolgreichem _robust_move wird das
    # Original ueberschrieben. Schuetzt vor Datenverlust bei Crash
    # (Watchdog-Kill, Netzwerk-Drop, COM-Fehler) waehrend Save.
    temp_stage2_path    = None
    ext                 = os.path.splitext(file_path)[1].lower()

    if is_locked_by_other(original_path):
        pbar.write(f"  ->  ÜBERSPRUNGEN (in Bearbeitung): "
                   f"{os.path.basename(original_path)}")
        detail_logger.info(
            f"Übersprungen (Sperrdatei/Sharing-Violation): {original_path}")
        return "SKIPPED_LOCKED"

    # ── Zeitstempel sichern ────────────────────────────────────────────────
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
    # wird nach jeder Datei-Ersetzung wieder angewendet (Admin-Kontext:
    # vollständig inkl. Owner; Nutzer-Kontext: DACL).
    orig_sd = _get_security_descriptor(original_path)

    # ── OLE-Verschlüsselungserkennung für .doc/.dot ────────────────────────
    if ext in (".doc", ".dot") and _is_doc_encrypted(file_path):
        pbar.write(f"  ->  ÜBERSPRUNGEN (OLE-verschlüsselt): "
                   f"{os.path.basename(original_path)}")
        detail_logger.info(f"Übersprungen (OLE-verschlüsselt): {original_path}")
        return "SKIPPED_ENCRYPTED"

    # ── Kennwort bei OOXML vorab erkennen (siehe _ooxml_kennwort_grund) ────
    if ext in (".docx", ".docm", ".dotx", ".dotm"):
        _kw_grund = _ooxml_kennwort_grund(file_path)
        if _kw_grund == "verschluesselt":
            pbar.write(f"  ->  ÜBERSPRUNGEN (verschlüsselt): "
                       f"{os.path.basename(original_path)}")
            detail_logger.info(f"Übersprungen (OOXML verschlüsselt): {original_path}")
            return "SKIPPED_ENCRYPTED"
    else:
        _kw_grund = ""

    # ── Schreibschutz-Prüfung (OS-Level) ───────────────────────────────────
    was_read_only = False
    try:
        safe_check = prepare_long_path(file_path)
        file_attrs = os.stat(safe_check).st_mode
        if not (file_attrs & stat.S_IWRITE):
            os.chmod(safe_check, file_attrs | stat.S_IWRITE)
            was_read_only = True
            detail_logger.debug("Schreibschutz temporär entfernt")
    except Exception as e_ro:
        detail_logger.debug(f"Schreibschutz-Prüfung fehlgeschlagen: {e_ro}")

    # ── Long-Path-Behandlung (mit AV-Retry beim Kopieren) ──────────────────
    if len(file_path) > MAX_PATH_LEN:
        detail_logger.debug("Long-Path erkannt, erstelle Temp-Kopie")
        try:
            uid         = uuid.uuid4().hex
            basename    = os.path.basename(file_path)
            unique_name = f"{uid}_{basename}"
            temp_path   = os.path.join(TEMP_PROCESS_PATH, unique_name)

            if len(temp_path) >= WINDOWS_MAX_PATH:
                unique_name = f"{uid}{os.path.splitext(basename)[1].lower()}"
                temp_path   = os.path.join(TEMP_PROCESS_PATH, unique_name)
                detail_logger.debug(
                    f"Dateiname zu lang für Temp-Pfad, kürze auf UUID: {temp_path}")

            _robust_copy(
                prepare_long_path(file_path),
                temp_path,
                op_name="Long-Path-Tempkopie")

            try:
                current_mode = os.stat(temp_path).st_mode
                os.chmod(temp_path, current_mode | stat.S_IWRITE)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
            file_path    = temp_path
            is_temp_copy = True
            detail_logger.debug(f"Temp-Kopie: {temp_path}")
        except Exception as e:
            log_error(original_path, Exception(f"Long-Path-Kopie fehlgeschlagen: {e}"))
            # Schreibschutz wiederherstellen: dieser Fruehausstieg liegt
            # noch vor dem try/finally, das den Schutz sonst zuruecksetzt.
            if was_read_only:
                try:
                    safe_ro = prepare_long_path(original_path)
                    current_mode = os.stat(safe_ro).st_mode
                    os.chmod(safe_ro, current_mode & ~stat.S_IWRITE)
                except Exception as _e:
                    detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
            return "ERROR"

    # ── Schreibkennwort (OOXML): Arbeitskopie ohne Kennwort ──────────────
    if _kw_grund == "schreibkennwort":
        try:
            os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)
            ohne = os.path.join(TEMP_PROCESS_PATH, f"ohne_kennwort_{uuid.uuid4().hex}{ext}")
            _kopie_ohne_schreibkennwort(file_path, ohne)
            if is_temp_copy:
                try:
                    os.remove(file_path)
                except OSError as _e:
                    detail_logger.debug(f"Alte Temp-Kopie nicht entfernbar: {_e!r}")
            file_path    = ohne
            is_temp_copy = True
            pbar.write(f"  ->  Schreibkennwort wird entfernt: {os.path.basename(original_path)}")
            detail_logger.info(f"Schreibkennwort entfernt (Arbeitskopie): {original_path}")
        except Exception as e:
            log_error(original_path, Exception(f"Kopie ohne Schreibkennwort fehlgeschlagen: {e}"))
            if was_read_only:
                try:
                    safe_ro = prepare_long_path(original_path)
                    os.chmod(safe_ro, os.stat(safe_ro).st_mode & ~stat.S_IWRITE)
                except Exception as _e:
                    detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
            return "ERROR"

    try:
        # ── STAGE 1: .doc/.dot → temp .docx/.dotx ─────────────────────────
        if ext in (".doc", ".dot"):
            # Watchdog-geschuetzter Open: ohne diesen wuerde ein unsichtbarer
            # Modal-Dialog (DRM-Popup, Verknuepfungs-Update,
            # Makro-Sicherheitswarnung trotz DisplayAlerts=False) den Thread
            # unbegrenzt blockieren. Der Watchdog killt nach
            # WORD_OPEN_TIMEOUT die Word-Instanz; die Aufrufer-Schleife
            # erkennt anhand word_app_global=None den toten Zustand und
            # startet beim naechsten File neu.
            # ReadOnly: Stage 1 liest nur und schreibt per SaveAs2 in eine
            # neue Datei. Schreibend oeffnen fragte bei einer .doc mit
            # Schreibkennwort unsichtbar nach dem Kennwort (Waechter);
            # schreibgeschuetzt oeffnet Word sie ohne Abfrage (gemessen
            # 30.09.2026), und SaveAs2 mit WritePassword="" laesst das
            # Kennwort weg - gewollt (Entscheidung vom 30.09.2026).
            def _do_open_stage1():
                return word_app.Documents.Open(
                    file_path,
                    ConfirmConversions = COM_FALSE,
                    ReadOnly          = COM_TRUE,
                    AddToRecentFiles  = COM_FALSE,
                    OpenAndRepair     = COM_FALSE,
                    NoEncodingDialog  = COM_TRUE,
                    PasswordDocument  = "",
                    PasswordTemplate  = "",
                )
            try:
                doc = _word_call_with_watchdog(
                    "Documents.Open (Stage 1)",
                    WORD_OPEN_TIMEOUT, word_pid_global, _do_open_stage1,
                    word_create_time_global
                )
            except TimeoutError:
                # Word-Instanz ist nach Watchdog-Kill tot. Globalstate
                # bereinigen, damit der Aufrufer beim naechsten File
                # _start_word() neu aufruft.
                word_app_global = None
                _eigene_besitzerdatei_entfernen(file_path)
                if file_path != original_path:
                    _eigene_besitzerdatei_entfernen(original_path)
                pbar.write(f"  ✕  TIMEOUT bei Stage-1-Open ({WORD_OPEN_TIMEOUT:.0f}s) – Word blockiert")
                log_error(original_path, Exception(
                    f"Timeout bei Documents.Open (Stage 1) nach {WORD_OPEN_TIMEOUT:.0f}s "
                    f"(vermutlich Modal-Dialog: DRM/Verknuepfung/Makro)"
                ))
                return "ERROR"
            detail_logger.debug("Dokument geöffnet (Stage 1 – Formatkonvertierung)")

            has_macros       = False
            vba_check_failed = False
            try:
                has_macros = doc.HasVBProject
            except Exception as e_vba:
                try:
                    comps      = doc.VBProject.VBComponents
                    has_macros = any(
                        comps.Item(i).CodeModule.CountOfLines > 0
                        for i in range(1, comps.Count + 1)
                    )
                    detail_logger.debug(
                        "HasVBProject fehlgeschlagen – VBProject-Fallback erfolgreich")
                except Exception:
                    vba_check_failed = True
                    has_macros       = True
                    detail_logger.warning(
                        f"HasVBProject-Prüfung und VBProject-Fallback fehlgeschlagen "
                        f"(Trust-Center korrekt konfiguriert? → "
                        f"Makroeinstellungen + VBA-Projektzugriff prüfen): {e_vba}\n"
                        f"  Datei wird im Makro-Format gespeichert (konservativ).")

            if ext == ".dot":
                new_format = WD_FORMAT_DOTM if has_macros else WD_FORMAT_DOTX
                new_ext    = ".dotm"         if has_macros else ".dotx"
            else:
                new_format = WD_FORMAT_DOCM if has_macros else WD_FORMAT_DOCX
                new_ext    = ".docm"         if has_macros else ".docx"

            if vba_check_failed:
                pbar.write(
                    f"  ⚠  VBA-Prüfung nicht möglich → als {new_ext} gesichert: "
                    f"{os.path.basename(original_path)}")

            temp_stage1_name = f"{uuid.uuid4().hex}{new_ext}"
            temp_stage1_path = os.path.join(TEMP_PROCESS_PATH, temp_stage1_name)

            # Watchdog-geschuetzter SaveAs2: Word kann beim Speichern in
            # die gleichen Modal-Dialoge wie beim Open laufen.
            try:
                _word_call_with_watchdog(
                    "Documents.SaveAs2 (Stage 1)",
                    WORD_SAVE_TIMEOUT, word_pid_global,
                    lambda: doc.SaveAs2(temp_stage1_path, FileFormat=new_format,
                                        AddToRecentFiles=COM_FALSE,
                                        WritePassword=""),
                    word_create_time_global
                )
            except TimeoutError:
                word_app_global = None
                pbar.write(f"  ✕  TIMEOUT bei Stage-1-SaveAs2 ({WORD_SAVE_TIMEOUT:.0f}s) – Word blockiert")
                log_error(original_path, Exception(
                    f"Timeout bei SaveAs2 (Stage 1) nach {WORD_SAVE_TIMEOUT:.0f}s"
                ))
                # doc-Objekt unbrauchbar (Word tot), finally-Cleanup soll
                # nicht erneut crashen.
                doc = None
                return "ERROR"
            detail_logger.debug(
                f"Stage 1: Als {new_ext} gespeichert → {temp_stage1_path}")

            doc.Close(SaveChanges=COM_FALSE)
            doc = None

            if is_temp_copy and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception as _e:
                    detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
                is_temp_copy = False

            file_path     = temp_stage1_path
            new_path      = os.path.splitext(original_path)[0] + new_ext
            was_converted = True
            # Gemessen am 30.09.2026: die aus einer .doc mit Schreibkennwort
            # erzeugte .docx trug das Kennwort trotz WritePassword="" weiter,
            # und Stage 2 (schreibend) hing in der unsichtbaren Abfrage bis
            # zum Waechter. Deshalb hier auf Dateiebene nachziehen.
            if _ooxml_kennwort_grund(temp_stage1_path) == "schreibkennwort":
                _ohne = temp_stage1_path + ".ohne_kennwort"
                _kopie_ohne_schreibkennwort(temp_stage1_path, _ohne)
                os.replace(_ohne, temp_stage1_path)
                detail_logger.info(f"Schreibkennwort entfernt (nach Stage 1): {original_path}")
            detail_logger.debug(
                "Stage 1 abgeschlossen – öffne für Font-Verarbeitung (Stage 2)")

        # ── STAGE 2: Dokument öffnen ──────────────────────────────────────
        # Watchdog-geschuetzt - selbe Begruendung wie in Stage 1. Bei einer
        # .doc/.dot wurde der Pfad in Stage 1 auf das frisch erzeugte
        # temp_stage1_path umgebogen, bei .docx/.dotx ist es der Original-
        # (bzw. Long-Path-Temp-)Pfad.
        def _do_open_stage2():
            return word_app.Documents.Open(
                file_path,
                ConfirmConversions = COM_FALSE,
                ReadOnly          = COM_FALSE,
                AddToRecentFiles  = COM_FALSE,
                OpenAndRepair     = COM_FALSE,
                NoEncodingDialog  = COM_TRUE,
                PasswordDocument  = "",
                PasswordTemplate  = "",
            )
        try:
            doc = _word_call_with_watchdog(
                "Documents.Open (Stage 2)",
                WORD_OPEN_TIMEOUT, word_pid_global, _do_open_stage2,
                word_create_time_global
            )
        except TimeoutError:
            word_app_global = None
            _eigene_besitzerdatei_entfernen(file_path)
            if file_path != original_path:
                _eigene_besitzerdatei_entfernen(original_path)
            pbar.write(f"  ✕  TIMEOUT bei Stage-2-Open ({WORD_OPEN_TIMEOUT:.0f}s) – Word blockiert")
            log_error(original_path, Exception(
                f"Timeout bei Documents.Open (Stage 2) nach {WORD_OPEN_TIMEOUT:.0f}s"
            ))
            return "ERROR"
        detail_logger.debug("Dokument geöffnet (Hauptverarbeitung)")

        try:
            doc.RemovePersonalInformation = False
        except Exception as _e:
            detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")

        orig_track_revisions = False
        try:
            orig_track_revisions = doc.TrackRevisions
            doc.TrackRevisions   = False
        except Exception as _e:
            detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")

        # Setup-Aenderungen (RemovePersonalInformation=False,
        # TrackRevisions=False) neutralisieren: Ab hier zaehlt das
        # Word-Dirty-Flag nur noch ECHTE Aenderungen (AcceptAllRevisions,
        # Convert, Font-Ersetzungen, Metadaten-Entfernung). Grundlage
        # fuer die Unveraendert-Erkennung vor dem Speichern.
        try:
            doc.Saved = True
        except Exception as _e:
            detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")

        if 2 in metadata_types:
            try:
                doc.AcceptAllRevisions()
                detail_logger.debug(
                    "AcceptAllRevisions() ausgeführt (Metadata-Typ 2 gewählt)")
            except Exception as e_accept:
                detail_logger.debug(f"AcceptAllRevisions fehlgeschlagen: {e_accept}")

        # ── Schema-Migration (CompatibilityMode) ──────────────────────────
        try:
            compat_mode = doc.CompatibilityMode
            detail_logger.debug(f"CompatibilityMode: {compat_mode}")
            if compat_mode < 15:
                doc.Saved = True
                doc.Convert()
                if doc.Saved:
                    detail_logger.debug(
                        f"Schema bereits aktuell (Modus {compat_mode})")
                else:
                    schema_migrated = True
                    detail_logger.info(
                        f"Schema-Migration durchgeführt (Modus {compat_mode} → aktuell): "
                        f"{original_path}")
            else:
                detail_logger.debug(
                    f"Schema bereits aktuell (Modus {compat_mode}), Convert() übersprungen")
        except Exception as e_conv:
            detail_logger.debug(f"doc.Convert() nicht möglich: {e_conv}")

        _SYMBOL_KANDIDATEN = _symbolschriften_im_dokument(doc)
        detail_logger.debug(f"Symbolschriften im Dokument: {_SYMBOL_KANDIDATEN}")

        # ── 1. Formatvorlagen (Styles) ────────────────────────────────────
        for style in doc.Styles:
            try:
                if _ist_symbolschrift(style.Font.Name):
                    continue
                _set_all_font_names(style.Font, font_name)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")

        # ── 2. Story Ranges + ShapeRange + InlineShapes ───────────────────
        # Bekannter Word-Fehler: StoryRanges/NextStoryRange ueberspringen
        # Kopf-/Fusszeilen spaeterer Abschnitte, wenn die erste Kopfzeile
        # leer bzw. noch nie angesprochen ist. Dokumentierter Vorgriff (Word-
        # MVP-Muster "lngJunk = ...Sections(1).Headers(1).Range.StoryType"):
        # die Kopfzeile einmal ansprechen, BEVOR StoryRanges gelesen wird.
        # Zusaetzlich werden unten alle Abschnitte mit ihren drei Kopf- und
        # Fusszeilen direkt bearbeitet - doppelt Bearbeitetes ist harmlos
        # (_set_all_font_names setzt nur Abweichendes).
        try:
            _ = doc.Sections(1).Headers(1).Range.StoryType
        except Exception as _e:
            detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")

        def _story_bearbeiten(current):
            _set_font_on_range(current, font_name)

            try:
                for shape in current.ShapeRange:
                    try:
                        _set_font_on_shape(shape, font_name)
                    except Exception as _e:
                        detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")

            try:
                for ishape in current.InlineShapes:
                    try:
                        if ishape.HasSmartArt:
                            for node in ishape.SmartArt.Nodes:
                                try:
                                    _apply_smartart_node(node, font_name)
                                except Exception as _e:
                                    detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
                        else:
                            try:
                                if ishape.HasTextFrame:
                                    _set_font_on_range(
                                        ishape.TextFrame.TextRange, font_name)
                            except Exception as _e:
                                detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
                            try:
                                _set_font_on_range(ishape.Range, font_name)
                            except Exception as _e:
                                detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
                    except Exception as _e:
                        detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")

        for story in doc.StoryRanges:
            current = story
            while current is not None:
                _story_bearbeiten(current)
                try:
                    current = current.NextStoryRange
                except Exception:
                    current = None

        # Zusaetzlich jede Kopf-/Fusszeile jedes Abschnitts direkt
        # (wdHeaderFooterPrimary/FirstPage/EvenPages = 1/2/3) - greift auch,
        # falls der Vorgriff oben in einer Word-Fassung nicht genuegt.
        # Verknuepfte (LinkToPrevious) zeigen denselben Inhalt wie im
        # Vorabschnitt und werden nicht erneut bearbeitet.
        try:
            for section in doc.Sections:
                for sammlung in (section.Headers, section.Footers):
                    for idx in (1, 2, 3):
                        try:
                            hf = sammlung(idx)
                            if hf.Exists and not hf.LinkToPrevious:
                                _story_bearbeiten(hf.Range)
                        except Exception as _e:
                            detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
        except Exception as _e:
            detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")

        # ── 3. Tabellen ───────────────────────────────────────────────────
        for table in doc.Tables:
            try:
                _set_font_on_range(table.Range, font_name)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")

        # ── 4. Fuß- und Endnoten ──────────────────────────────────────────
        _process_footnotes_endnotes(doc, font_name)

        # ── 5. Content Controls ───────────────────────────────────────────
        _process_content_controls(doc, font_name)

        # ── 6. Revisionen (nur wenn nicht via Typ 2 entfernt) ─────────────
        if 2 not in metadata_types:
            _process_revisions(doc, font_name)

        # ── 7. Kommentare (nur wenn nicht via Typ 1 entfernt) ─────────────
        if 1 not in metadata_types:
            for comment in doc.Comments:
                try:
                    _set_font_on_range(comment.Range, font_name)
                except Exception as _e:
                    detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")

        detail_logger.debug("Schriftarten ersetzt")

        # ── 8. Metadaten entfernen ────────────────────────────────────────
        if metadata_types and doc is not None:
            for info_type in metadata_types:
                try:
                    doc.RemoveDocumentInformation(info_type)
                except Exception as e:
                    detail_logger.debug(f"InfoType {info_type} nicht entfernt: {e}")

        # ── 9. TrackRevisions zurücksetzen (NACH Metadaten, VOR Save) ─────
        try:
            doc.TrackRevisions = orig_track_revisions
        except Exception as _e:
            detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")

        # ── 9.5. Unveraendert-Erkennung ───────────────────────────────────
        # Word pflegt das Dirty-Flag selbst: Ist nach Font-Pass und
        # Metadaten-Entfernung nichts geaendert (doc.Saved noch True),
        # wird ohne Speichern geschlossen - die Datei auf der Ablage
        # bleibt byte-identisch unberuehrt. Macht Re-Runs ueber bereits
        # bearbeitete Bestaende schnell und nicht-destruktiv.
        # Konservative Ausnahmen: Stage-1-Konvertierungen muessen immer
        # abgeschlossen werden; das TrackRevisions-Reset auf True (Zeile
        # oben) macht Dokumente mit aktivem Aenderungsmodus dirty -> die
        # werden weiterhin gespeichert (harmlos, nur nicht optimiert).
        if not was_converted:
            doc_unchanged = False
            try:
                doc_unchanged = bool(doc.Saved)
            except Exception:
                doc_unchanged = False
            if doc_unchanged:
                doc.Close(SaveChanges=COM_FALSE)
                doc = None
                pbar.write(f"  =  unverändert: {os.path.basename(original_path)}")
                detail_logger.info(
                    f"Unveraendert – kein Speichern noetig: {original_path}")
                return "UNCHANGED"

        # ── 10. Speichern ─────────────────────────────────────────────────
        # Watchdog-geschuetzt - Save kann ebenfalls in Modal-Dialoge laufen
        # ("Datei in neuerem Format speichern?", Lizenz-Popup,
        # Druckereinrichtungs-Dialog bei nicht-erreichbarem Default-Drucker).
        #
        # DATENVERLUST-SCHUTZ:
        #  - was_converted=True : Word arbeitet auf temp_stage1_path
        #                         (Stage-1-SaveAs2). Save() ist safe.
        #  - is_temp_copy=True  : Word arbeitet auf Long-Path-Temp-Kopie.
        #                         Save() ist safe (Block 13 schiebt zurueck).
        #  - sonst              : Word wuerde direkt das ORIGINAL ueberschreiben.
        #                         Bei Crash mittendrin (Watchdog-Kill,
        #                         Netzwerk-Drop, COM-Fehler) waere das
        #                         Original korrupt/leer. Daher: SaveAs2 in
        #                         eine Stage-2-Tempdatei, dann nach Close
        #                         per _robust_move atomar zum Original.
        save_in_place_safe = was_converted or is_temp_copy

        try:
            if save_in_place_safe:
                _word_call_with_watchdog(
                    "Documents.Save",
                    WORD_SAVE_TIMEOUT, word_pid_global,
                    lambda: doc.Save(),
                    word_create_time_global
                )
                detail_logger.debug("Gespeichert (Fonts"
                                    + (" & Metadaten)" if metadata_types else ")"))
            else:
                # Stage-2-SaveAs2 in Temp-Pfad. Format-Erhalt durch
                # FileFormat=Original-Format (doc.SaveFormat liest den
                # aktuellen FileFormat-Code der geoeffneten Datei).
                temp_stage2_name = f"stage2_{uuid.uuid4().hex}{ext}"
                temp_stage2_path = os.path.join(TEMP_PROCESS_PATH, temp_stage2_name)
                try:
                    save_format = doc.SaveFormat
                except Exception:
                    save_format = None

                def _do_save_stage2():
                    if save_format is not None:
                        doc.SaveAs2(temp_stage2_path,
                                    FileFormat=save_format,
                                    AddToRecentFiles=COM_FALSE)
                    else:
                        # Fallback: ohne FileFormat - Word waehlt anhand
                        # der Extension. Funktioniert fuer .docx/.dotx/.docm/.dotm.
                        doc.SaveAs2(temp_stage2_path,
                                    AddToRecentFiles=COM_FALSE)

                _word_call_with_watchdog(
                    "Documents.SaveAs2 (Stage 2)",
                    WORD_SAVE_TIMEOUT, word_pid_global,
                    _do_save_stage2,
                    word_create_time_global
                )
                detail_logger.debug(
                    f"Stage 2: SaveAs2 -> {temp_stage2_path}"
                    + (" (Fonts & Metadaten)" if metadata_types else " (Fonts)")
                )
        except TimeoutError:
            word_app_global = None
            pbar.write(f"  ✕  TIMEOUT bei Save ({WORD_SAVE_TIMEOUT:.0f}s) – Word blockiert")
            log_error(original_path, Exception(
                f"Timeout bei Save nach {WORD_SAVE_TIMEOUT:.0f}s"
            ))
            doc = None
            return "ERROR"

        # ── 11. Dokument schließen ────────────────────────────────────────
        if doc is not None:
            doc.Close(SaveChanges=COM_FALSE)
            doc = None
            detail_logger.debug("Dokument geschlossen")

        # ── 11.5. Stage-2-Tempdatei → Original zurueckschieben ─────────────
        # Nur wenn der direct-on-original-Save-Pfad genommen wurde
        # (was_converted=False, is_temp_copy=False). In den anderen Faellen
        # wird der Move durch Block 12 (Stage-1-Konvertierung) oder
        # Block 13 (Long-Path-Restore) erledigt.
        if temp_stage2_path is not None:
            try:
                safe_orig = prepare_long_path(original_path)
                if not (os.path.exists(temp_stage2_path)
                        and os.path.getsize(temp_stage2_path) > 0):
                    raise RuntimeError(
                        f"Stage-2-Tempdatei ist leer/fehlt: {temp_stage2_path}")
                # Schreibschutz auf dem Ziel kurzzeitig entfernen, falls
                # gesetzt - sonst kann der Move auf NTFS fehlschlagen.
                if os.path.exists(safe_orig):
                    try:
                        current_mode = os.stat(safe_orig).st_mode
                        os.chmod(safe_orig, current_mode | stat.S_IWRITE)
                    except Exception as _e:
                        detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
                # Atomare Ersetzung: Staging-Kopie auf dem Zielvolume +
                # os.replace. Das Original bleibt bei jedem Fehlschlag
                # unveraendert erhalten.
                _replace_file_atomic(temp_stage2_path, safe_orig,
                                     op_name="Stage-2-Move",
                                     sd=orig_sd)
                temp_stage2_path = None
                detail_logger.debug(
                    f"Stage-2-Tempdatei -> Original verschoben: {original_path}")
            except Exception as e_s2_move:
                # Ersetzung fehlgeschlagen. Dank os.replace ist das
                # Original unveraendert intakt. Stage-2-Tempdatei
                # zusaetzlich als RESCUE_-Datei sichern.
                rescue_name = f"RESCUE_{uuid.uuid4().hex}{ext}"
                rescue_path = os.path.join(TEMP_PROCESS_PATH, rescue_name)
                detail_logger.warning(
                    f"Stage-2-Move fehlgeschlagen: {e_s2_move}\n"
                    f"  Versuche Rescue-Kopie nach: {rescue_path}\n"
                    f"  Original unveraendert intakt: {original_path}"
                )
                try:
                    if temp_stage2_path and os.path.exists(temp_stage2_path):
                        _robust_copy(temp_stage2_path, rescue_path,
                                     op_name="Stage-2-Rescue")
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
                return "ERROR"

        # ── 12. Stage-1-Tempdatei → Zielort verschieben ───────────────────
        if was_converted and temp_stage1_path is not None:
            try:
                unique_new_path = _get_unique_target_path(new_path)
                if unique_new_path != new_path:
                    detail_logger.info(
                        f"Zielname kollidiert – weiche aus auf: {unique_new_path}")
                    pbar.write(
                        f"  ℹ  Zielname existiert bereits – speichere als: "
                        f"{os.path.basename(unique_new_path)}")
                new_path = unique_new_path
                safe_new = prepare_long_path(new_path)

                # Staging auf dem Zielvolume + os.replace statt shutil.move:
                # Temp (Dokumente, C:) und Ablage (Q:/R:/UNC) liegen auf
                # verschiedenen Volumes, shutil.move kopiert dann per copy2
                # direkt unter dem Zielnamen - ein Abbruch mittendrin liess
                # eine halbe X.docx neben der X.doc zurueck (der Folgelauf
                # wich dann auf "X (konvertiert).docx" aus). Jetzt entsteht
                # X.docx erst mit dem atomaren os.replace; eine halbe
                # Staging-Datei raeumt _replace_file_atomic selbst ab.
                # ACL folgt unten (sd=None hier, sonst doppelt).
                _replace_file_atomic(temp_stage1_path, safe_new,
                                     op_name="Stage-1-Move", sd=None)
                temp_stage1_path = None

                if not (os.path.exists(safe_new) and os.path.getsize(safe_new) > 0):
                    raise RuntimeError(
                        f"Konvertierte Datei nicht am Ziel angekommen: {new_path}")

                # ACL/Owner des Originals auf die konvertierte Datei
                # übertragen (Admin-Kontext: vollständig; Nutzer: DACL).
                _apply_security_descriptor(new_path, orig_sd)

                # Nachweisliste: alte/neue Pfade der Konvertierung in CSV.
                _log_conversion(original_path, new_path)

                safe_orig = prepare_long_path(original_path)
                if os.path.exists(safe_orig):
                    try:
                        try:
                            current_mode = os.stat(safe_orig).st_mode
                            os.chmod(safe_orig, current_mode | stat.S_IWRITE)
                        except Exception as _e:
                            detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
                        _robust_remove(safe_orig, op_name="Alte-doc-Loeschung")
                        detail_logger.debug(f"Alte Datei gelöscht: {original_path}")
                    except Exception as e_del:
                        detail_logger.warning(
                            f"Alte Datei nicht löschbar ({e_del}) "
                            f"– new_path ist gültig: {new_path}")
                        pbar.write(
                            f"  ⚠  Alte Datei nicht entfernbar (Konvertierung OK): "
                            f"{os.path.basename(original_path)}")
            except Exception as e_s1_move:
                rescue_ext  = os.path.splitext(new_path)[1] if new_path else ".docx"
                rescue_name = f"RESCUE_{uuid.uuid4().hex}{rescue_ext}"
                rescue_path = os.path.join(TEMP_PROCESS_PATH, rescue_name)
                detail_logger.warning(
                    f"Stage-1-Move fehlgeschlagen: {e_s1_move}\n"
                    f"  Versuche Rescue-Kopie nach: {rescue_path}"
                )
                try:
                    src_for_rescue = None
                    if temp_stage1_path and os.path.exists(temp_stage1_path):
                        src_for_rescue = temp_stage1_path
                    elif new_path and os.path.exists(prepare_long_path(new_path)):
                        src_for_rescue = prepare_long_path(new_path)
                    if src_for_rescue:
                        _robust_copy(src_for_rescue, rescue_path,
                                     op_name="Stage-1-Rescue")
                        detail_logger.warning(
                            f"Stage-1-Rescue gesichert: {rescue_path}")
                        pbar.write(
                            f"  ⚠  RESCUE (Konvertierung): "
                            f"{os.path.basename(rescue_path)}")
                except Exception as e_s1_rescue:
                    detail_logger.error(
                        f"Stage-1-Rescue fehlgeschlagen: {e_s1_rescue}")
                    if temp_stage1_path and os.path.exists(temp_stage1_path):
                        _preserved_temp_files.add(temp_stage1_path)
                        detail_logger.error(
                            f"DATENVERLUST-SCHUTZ: Stage-1-Temp wird NICHT gelöscht: "
                            f"{temp_stage1_path}\n"
                            f"  → Datei manuell prüfen und sichern!")
                log_error(
                    original_path,
                    Exception(f"Stage-1-Move fehlgeschlagen: {e_s1_move}"))
                pbar.write(
                    f"  ✗  FEHLER (Stage-1-Move): "
                    f"{os.path.basename(original_path)}")
                return "ERROR"

        # ── 13. Long-Path-Tempdatei → Original zurückschieben ─────────────
        if is_temp_copy:
            try:
                safe_orig = prepare_long_path(original_path)
                if os.path.exists(safe_orig):
                    try:
                        current_mode = os.stat(safe_orig).st_mode
                        os.chmod(safe_orig, current_mode | stat.S_IWRITE)
                    except Exception as _e:
                        detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
                # Atomare Ersetzung: das Original bleibt auch bei einem
                # Abbruch mitten im Kopiervorgang unveraendert erhalten.
                _replace_file_atomic(file_path, safe_orig,
                                     op_name="Long-Path-Restore",
                                     sd=orig_sd)
                detail_logger.debug("Temp-Datei zurückverschoben")
                is_temp_copy = False
            except Exception as e_move:
                rescue_ext  = os.path.splitext(os.path.basename(original_path))[1]
                rescue_name = f"RESCUE_{uuid.uuid4().hex}{rescue_ext}"
                rescue_path = os.path.join(TEMP_PROCESS_PATH, rescue_name)
                detail_logger.warning(
                    f"Zurückschieben fehlgeschlagen: {e_move}\n"
                    f"  Versuche Rescue-Kopie nach: {rescue_path}"
                )
                try:
                    _robust_copy(file_path, rescue_path, op_name="Long-Path-Rescue")
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
                return "ERROR"

        # ── 14. Zeitstempel wiederherstellen (mit exponentiellem Backoff) ─
        if orig_times:
            ts_target = new_path if (was_converted and new_path) else original_path
            if ts_target and os.path.exists(prepare_long_path(ts_target)):
                ts_last_exc = None
                for attempt, delay in enumerate([0.0, *TS_RETRY_DELAYS]):
                    if delay > 0:
                        time.sleep(delay)
                    try:
                        h_dst = win32file.CreateFile(
                            prepare_long_path(ts_target),
                            FILE_WRITE_ATTRIBUTES,
                            win32file.FILE_SHARE_READ
                                | win32file.FILE_SHARE_WRITE
                                | win32file.FILE_SHARE_DELETE,
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
                        ts_last_exc = None
                        break
                    except Exception as e_utime:
                        ts_last_exc = e_utime
                        # Nur bei Lock-artigen Fehlern lohnt ein Retry -
                        # permanente Fehler nicht ~17 s lang wiederholen.
                        if not _is_av_lock_error(e_utime):
                            break
                if ts_last_exc is not None:
                    # Rueckfall auf os.utime: erhaelt Zugriffs- und
                    # Aenderungszeit (nicht die Erstellungszeit), aber das
                    # ist deutlich besser als der vollstaendige Verlust
                    # des Datums. Bisher hatte nur 5_OCR_PDF.py diesen
                    # Rueckfall.
                    if not _utime_rueckfall(ts_target, orig_times):
                        detail_logger.warning(
                            f"Zeitstempel nicht wiederherstellbar: {ts_last_exc}")
                    else:
                        detail_logger.info(
                            "Zeitstempel über os.utime-Rückfall gesetzt "
                            "(ohne Erstellungszeit).")

        _flags = []
        if was_converted:
            _flags.append("konvertiert")
        if schema_migrated:
            _flags.append("Schema aktualisiert")
        _suffix = f" [{', '.join(_flags)}]" if _flags else ""
        pbar.write(f"  ✓  {os.path.basename(original_path)}{_suffix}")
        detail_logger.info(f"Erfolgreich: {original_path}{_suffix}")
        return "SUCCESS"

    except Exception as e:
        is_password_error = _ist_kennwortfehler(
            e, (original_path, file_path, new_path,
                temp_stage1_path, temp_stage2_path))

        if is_password_error:
            pbar.write(f"  ->  ÜBERSPRUNGEN (Passwort): {os.path.basename(original_path)}")
            detail_logger.info(f"Übersprungen (Passwort): {original_path}")
            return "SKIPPED_PASSWORD"

        # Toter COM-Server (Word gestorben ohne Watchdog-Timeout):
        # word_app_global=None triggert den reaktiven Neustart im
        # Aufrufer - sonst wuerde jede Folgedatei am toten COM-Objekt
        # scheitern, bis der 200er-Intervall-Neustart greift.
        if _is_rpc_dead_error(e):
            word_app_global = None
            detail_logger.warning(
                f"RPC-Verbindung zu Word verloren ({e}) – "
                f"reaktiver Neustart wird ausgeloest")

        log_error(original_path, e)
        pbar.write(f"  ✗  FEHLER: {os.path.basename(original_path)}")
        return "ERROR"

    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
        if is_temp_copy and os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
        if (temp_stage1_path is not None
                and os.path.exists(temp_stage1_path)
                and temp_stage1_path not in _preserved_temp_files):
            try:
                os.remove(temp_stage1_path)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
        if (temp_stage2_path is not None
                and os.path.exists(temp_stage2_path)
                and temp_stage2_path not in _preserved_temp_files):
            try:
                os.remove(temp_stage2_path)
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_document_com: Exception verworfen: {_e!r}")
        if was_read_only:
            ro_candidates = []
            if was_converted and new_path:
                ro_candidates.append(new_path)
            ro_candidates.append(original_path)
            for ro_path in ro_candidates:
                try:
                    safe_ro = prepare_long_path(ro_path)
                    if os.path.exists(safe_ro):
                        current_mode = os.stat(safe_ro).st_mode
                        os.chmod(safe_ro, current_mode & ~stat.S_IWRITE)
                        detail_logger.debug(
                            f"Schreibschutz wiederhergestellt: {ro_path}")
                        break
                except Exception as e_ro_restore:
                    detail_logger.debug(
                        f"Schreibschutz-Wiederherstellung fehlgeschlagen "
                        f"({ro_path}): {e_ro_restore}")


# ==================================================================
# Probelauf (Dry-Run) – ohne Word, ohne Aenderungen
# ==================================================================

def dry_run_directory(directory: str) -> dict:
    """Listet auf, was der Echtlauf tun WUERDE - insbesondere welche
    .doc/.dot konvertiert (und deren Originale ersetzt) wuerden.
    Es wird nichts geoeffnet, gespeichert oder geloescht; Word wird
    nicht gestartet."""
    stats = {"VERARBEITEN": 0, "KONVERTIEREN": 0,
             "GESPERRT": 0, "VERSCHLUESSELT": 0}
    print("\nPROBELAUF – es wird nichts geändert.\n")
    print("-" * 66)
    for file_path in file_generator(directory):
        ext = os.path.splitext(file_path)[1].lower()
        if is_locked_by_other(file_path):
            stats["GESPERRT"] += 1
            tag = "GESPERRT     "
        elif ext in (".doc", ".dot") and _is_doc_encrypted(file_path):
            stats["VERSCHLUESSELT"] += 1
            tag = "VERSCHLÜSSELT"
        elif ext in (".doc", ".dot"):
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
    print("    Würden konvertiert (.doc/.dot → neu,")
    print(f"                        Original ersetzt!):  {stats['KONVERTIEREN']}")
    print(f"    Gesperrt (würden übersprungen):          {stats['GESPERRT']}")
    print(f"    OLE-verschlüsselt (würden übersprungen): {stats['VERSCHLUESSELT']}")
    print(f"    GESAMT:                                  {sum(stats.values())}")
    return stats


# ==================================================================
# Verzeichnis-Verarbeitung
# ==================================================================

def process_directory(
    directory: str,
    font_name: str,
    metadata_types: list,
    count_files_first: bool,
    show_progress: bool = True,
    resume_path: Optional[str] = None,
    resume_set: Optional[set] = None,
) -> dict:
    global TEMP_PROCESS_PATH
    TEMP_PROCESS_PATH = os.path.join(TEMP_BASE_PATH, uuid.uuid4().hex)
    os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)

    resume_set = resume_set or set()

    if count_files_first:
        print("\nIndiziere Dateien (kann einige Minuten dauern) ...")
        files_list = [fp for fp in file_generator(directory)
                      if fp not in resume_set]
        total      = len(files_list)
        iterable   = files_list
        bar_fmt    = "{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]"
        print(f"{total} Dateien gefunden"
              + (f" ({len(resume_set)} via Resume übersprungen)" if resume_set else "")
              + ". Starte Verarbeitung ...\n")
    else:
        if resume_set:
            iterable = (fp for fp in file_generator(directory)
                        if fp not in resume_set)
        else:
            iterable = file_generator(directory)
        total    = None
        bar_fmt  = "{desc}: {n_fmt} Dateien [{elapsed}, {rate_fmt}]"
        print("\nStarte direkte Verarbeitung (ohne Vorab-Zählung) ...\n")

    print("-" * 66)

    stats = {"SUCCESS": 0, "UNCHANGED": 0, "ERROR": 0, "SKIPPED": 0}
    run_completed = False

    global word_app_global
    word = None

    _original_screen_updating   = True
    _original_pagination        = True
    _original_spell_check       = True
    _original_grammar_check     = True
    _original_background_save   = True
    _original_update_links      = False
    _original_no_prompt_convert = False

    def _start_word():
        nonlocal word
        nonlocal _original_screen_updating, _original_pagination
        nonlocal _original_spell_check, _original_grammar_check
        nonlocal _original_background_save
        nonlocal _original_update_links, _original_no_prompt_convert
        global word_app_global, word_pid_global, word_create_time_global

        try:
            _word_pids_before = {
                p.pid for p in psutil.process_iter(["name"])
                if p.info["name"] and p.info["name"].upper() == "WINWORD.EXE"
            }
        except Exception:
            _word_pids_before = set()

        word            = win32com.client.DispatchEx("Word.Application")
        word_app_global = word

        try:
            detail_logger.info(
                f"Word {word.Version} / Build {word.Build}")
        except Exception as e_ver:
            detail_logger.debug(f"Word-Version nicht ermittelbar: {e_ver}")

        # Der fruehere "Hwnd-Fallback" (word.Hwnd) scheiterte immer:
        # Word.Application hat keine Hwnd-Eigenschaft (MSWORD.OLB). War der
        # Snapshot mehrdeutig, lief der Hauptlauf ohne PID weiter - ohne
        # Watchdog-Kill und ohne Zombie-Schutz, ein Haenger blieb
        # unbegrenzt. Jetzt: _ermittle_word_pid, und ohne PID kein Lauf.
        word_pid_global = _ermittle_word_pid(word, _word_pids_before)
        if word_pid_global is None:
            try:
                word.Quit(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
            word            = None
            word_app_global = None
            raise RuntimeError(
                "PID der Word-Instanz nicht ermittelbar – ohne sie kann der "
                "Watchdog einen Hänger nicht beenden; Lauf abgebrochen")

        # Erstellungszeit der Word-Instanz sichern (PID-Recycling-Schutz
        # fuer Watchdog-Kill und _kill_specific_word).
        word_create_time_global = None
        if word_pid_global is not None:
            try:
                word_create_time_global = psutil.Process(
                    word_pid_global).create_time()
            except Exception as e_ct:
                detail_logger.debug(
                    f"Word-Erstellungszeit nicht ermittelbar: {e_ct}")

        try:
            _original_screen_updating = word.ScreenUpdating
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
        try:
            _original_pagination = word.Options.Pagination
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
        try:
            _original_spell_check = word.Options.CheckSpellingAsYouType
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
        try:
            _original_grammar_check = word.Options.CheckGrammarAsYouType
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
        try:
            _original_background_save = word.Options.BackgroundSave
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
        try:
            _original_update_links = word.Options.UpdateLinksAtOpen
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
        try:
            _original_no_prompt_convert = word.Options.DoNotPromptForConvert
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")

        word.Visible            = COM_FALSE
        word.DisplayAlerts      = COM_FALSE
        word.ScreenUpdating     = COM_FALSE
        word.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        try:
            word.Options.Pagination = False
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
        try:
            word.Options.CheckSpellingAsYouType = False
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
        try:
            word.Options.CheckGrammarAsYouType = False
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
        try:
            word.Options.BackgroundSave = False
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
        # Dialog-Quellen deaktivieren, vor denen der Watchdog schuetzt:
        # automatisches Verknuepfungs-Update und der Konvertierungs-Dialog
        # beim Oeffnen alter Formate. Bisher setzte das nur der Smoke-Test
        # - im Hauptlauf kostete jeder Treffer 180/240 s Timeout plus
        # Word-Neustart. Restore via _restore_word_settings.
        try:
            word.Options.UpdateLinksAtOpen = False
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")
        try:
            word.Options.DoNotPromptForConvert = True
        except Exception as _e:
            detail_logger.debug(f"_start_word: Exception verworfen: {_e!r}")

    try:
        _start_word()

        with tqdm(total=total, desc="Verarbeite", unit="Datei",
                  bar_format=bar_fmt, disable=not show_progress) as pbar:
            file_counter = 0
            for file_path in iterable:
                if file_counter > 0 and file_counter % WORD_RESTART_INTERVAL == 0:
                    pbar.write(
                        f"  ↻  Geplanter Word-Neustart nach {file_counter} Dateien ...")
                    detail_logger.info(
                        f"Geplanter Word-Neustart nach {file_counter} Dateien")
                    if word is not None:
                        _restore_word_settings(
                            word,
                            _original_screen_updating,
                            _original_pagination,
                            _original_spell_check,
                            _original_grammar_check,
                            _original_background_save,
                            _original_update_links,
                            _original_no_prompt_convert,
                        )
                        try:
                            word.Quit()
                        except Exception as _e:
                            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                    _kill_specific_word()
                    word = None
                    word_app_global = None
                    # Speicherbereinigungs-Neustart: idealer Zeitpunkt fuer
                    # Cache-Cleanup analog 3a (Word tot, frischer Neustart).
                    _cleanup_word_inetcache()
                    _cleanup_user_recent()
                    time.sleep(1)
                    try:
                        _start_word()
                    except Exception as e_planned:
                        detail_logger.error(
                            f"Geplanter Word-Neustart fehlgeschlagen: {e_planned}")
                        pbar.write(
                            f"  ✗  Word-Neustart fehlgeschlagen – "
                            f"Verarbeitung abgebrochen: {e_planned}")
                        break

                try:
                    result = replace_fonts_in_document_com(
                        file_path, word, pbar, font_name, metadata_types)
                    base = "SKIPPED" if result.startswith("SKIPPED") else result
                    stats[base] = stats.get(base, 0) + 1
                    if result != base:
                        stats[result] = stats.get(result, 0) + 1
                    if resume_path and result in RESUME_STATUSES:
                        append_resume(resume_path, file_path)
                except Exception as e:
                    log_error(file_path, e)
                    stats["ERROR"] += 1
                    pbar.write(f"  ✗  KRITISCH: {os.path.basename(file_path)}")
                finally:
                    pbar.update(1)
                    file_counter += 1

                # Toter-Word-Erkennung: Stirbt Word OHNE Watchdog-Timeout
                # (Absturz, AV-/EDR-Kill, OOM, Task-Manager), bliebe
                # word_app_global sonst gesetzt und jede Folgedatei
                # scheiterte am toten COM-Objekt, bis der 200er-Intervall-
                # Neustart zufaellig greift.
                if word_app_global is not None and not _word_process_alive():
                    detail_logger.warning(
                        "Word-Prozess ist tot (ohne Watchdog-Timeout) – "
                        "reaktiver Neustart wird ausgeloest")
                    word_app_global = None

                # Reaktiver Restart: replace_fonts_in_document_com setzt
                # word_app_global=None, wenn der Watchdog Word killen
                # musste. Ohne diesen Neustart wuerden alle Folge-Dateien
                # mit toten COM-Objekten crashen, bis der periodische
                # WORD_RESTART_INTERVAL-Neustart greift.
                if word_app_global is None and word is not None:
                    pbar.write("  ↻  Word reagiert nicht mehr – starte neu...")
                    detail_logger.info("Reaktiver Word-Neustart nach Watchdog-Kill")
                    _kill_specific_word()
                    word = None
                    # Cache-Cleanup nach Watchdog-Kill (Word ist garantiert tot).
                    _cleanup_word_inetcache()
                    _cleanup_user_recent()
                    time.sleep(1)
                    try:
                        _start_word()
                    except Exception as e_reactive:
                        detail_logger.error(
                            f"Reaktiver Word-Neustart fehlgeschlagen: {e_reactive}")
                        pbar.write(
                            f"  ✗  Word-Neustart fehlgeschlagen – "
                            f"Verarbeitung abgebrochen: {e_reactive}")
                        break
            else:
                # for-else: Schleife lief ohne break durch -> Lauf komplett.
                run_completed = True

        if run_completed and resume_path:
            delete_resume_file(resume_path)
            detail_logger.info(
                "Lauf vollstaendig abgeschlossen – Resume-Datei geloescht.")

    except Exception as e:
        print(f"\nKRITISCHER FEHLER: Word konnte nicht gestartet werden: {e}")
        log_error("GLOBAL", e)

    finally:
        if word is not None:
            _restore_word_settings(
                word,
                _original_screen_updating,
                _original_pagination,
                _original_spell_check,
                _original_grammar_check,
                _original_background_save,
                _original_update_links,
                _original_no_prompt_convert,
            )
            try:
                word.Quit()
                time.sleep(1)
            except Exception as _e:
                detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
        _kill_specific_word()
        _safe_cleanup_temp_process()
        # Office-INetCache (Content.Word/MSO) und Windows-Recent (.lnk)
        # aufraeumen - analog zu 3a/3b/3c. Beide best effort, fehlertolerant.
        try:
            _cleanup_word_inetcache()
        except Exception as _e:
            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
        try:
            _cleanup_user_recent()
        except Exception as _e:
            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")

    return stats



# ==================================================================
# Kommandozeile (Paritaet zu 4b/4c)
# ==================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="Word Font-Ersetzung – ersetzt Schriftarten in Word-Dateien",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--auto", action="store_true",
                        help="Trust-Center-Check und interaktive Rückfragen überspringen")
    parser.add_argument("--dir", dest="directory", default=None,
                        help="Zielverzeichnis")
    parser.add_argument("--font", default=None,
                        help=f"Ziel-Schriftart (Standard: {NEW_FONT_NAME})")
    parser.add_argument("--count-first", action="store_true",
                        help="Dateien vorab zählen (ETA-Anzeige)")
    parser.add_argument("--no-meta", action="store_true",
                        help="Metadaten nicht entfernen")
    parser.add_argument("--dry-run", action="store_true",
                        help="Probelauf: listet auf, was passieren würde. "
                             "Word wird nicht gestartet, nichts geöffnet, "
                             "gespeichert oder gelöscht.")
    parser.add_argument("--resume", dest="resume", default=None,
                        help="Resume-Datei: bereits fertige Dateien überspringen "
                             "(für geplante Tasks)")
    return parser.parse_args()


# ==================================================================
# Einstiegspunkt
# ==================================================================

if __name__ == "__main__":
    args      = parse_args()
    auto_mode = args.auto

    print()
    print("╔══════════════════════════════════════════════════════════════╗")
    print("║                     WORD FONT-ERSETZUNG                      ║")
    print("╚══════════════════════════════════════════════════════════════╝")
    print(f"  Start: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    # Vor dem ersten COM-Zugriff auf bereits laufende Sitzungen hinweisen.
    warn_running_word(auto_mode)


    print("\n" + "!" * 66)
    print("  WICHTIGER CHECK: TRUST-CENTER EINSTELLUNGEN")
    print("!" * 66)
    print("  Word öffnen → Datei → Optionen → Trust Center →")
    print("  'Einstellungen für das Trust Center...'")
    print()
    print("  1. GESCHÜTZTE ANSICHT:")
    print("     ☐ Geschützte Ansicht für Dateien aus dem Internet aktivieren")
    print("     ☐ Geschützte Ansicht für Dateien an potenziell unsicheren Speicherorten aktivieren")
    print("     ☐ Geschützte Ansicht für Outlook-Anlagen aktivieren")
    print("     → ALLE 3 DEAKTIVIEREN")
    print()
    print("  2. MAKROEINSTELLUNGEN:")
    print("     ⦿ Alle Makros aktivieren")
    print("     ☑ Zugriff auf das VBA-Projektobjektmodell vertrauen")
    print()
    print("  3. VERTRAUENSWÜRDIGE SPEICHERORTE:")
    print("     Folgende Speicherorte hinzufügen (☑ Unterordner ebenfalls vertrauenswürdig):")
    print(f"     • {TEMP_BASE_PATH}")
    print("       (zwingend – temporärer Arbeitsordner für Long-Path-Dateien)")
    print(r"     • Quellen auf Q:\, R:\, G:\ oder \\server\dfs:")
    print("       jeweiligen Pfad eintragen")
    print("       ☑ Vertrauenswürdige Speicherorte im Netzwerk zulassen")
    print("     • Quellen auf Desktop / Downloads / lokal:")
    print("       keine zusätzliche Eintragung nötig")
    print("!" * 66)

    if not auto_mode and not ask_yes_no(
        "Wurden diese Einstellungen in Word vorgenommen?",
        default_yes=False
    ):
        print("\n❌ Abbruch. Bitte konfigurieren Sie erst das Trust-Center in Word.")
        sys.exit(0)

    # --- Konfiguration: CLI hat Vorrang, sonst interaktive Abfrage ---
    if args.directory:
        start_dir = sanitize_path(args.directory)
        if not os.path.isdir(prepare_long_path(start_dir)):
            print(f"\n❌ Verzeichnis nicht erreichbar: '{start_dir}'")
            sys.exit(1)
    elif auto_mode:
        print("\n❌ Im Modus --auto ist --dir erforderlich.")
        sys.exit(1)
    else:
        start_dir = ask_directory()

    if auto_mode or args.count_first:
        show_progress = True
        count_first   = args.count_first
    else:
        show_progress, count_first = ask_progress_mode()

    NEW_FONT_NAME = args.font if args.font else (
        NEW_FONT_NAME if auto_mode else ask_font())

    metadata_types = []
    if args.no_meta:
        remove_meta = False
    elif auto_mode:
        # Ohne explizite Angabe im Automatikbetrieb bewusst nichts entfernen.
        remove_meta = False
    else:
        remove_meta = ask_yes_no("\nMetadaten entfernen?")
        if remove_meta:
            metadata_types = ask_metadata_detail()

    # --- Probelauf (Dry-Run) ---
    if args.dry_run:
        dry_run = True
    elif auto_mode:
        dry_run = False
    else:
        dry_run = ask_yes_no(
            "\nProbelauf (Dry-Run)? Zeigt nur, was passieren würde –\n"
            "  es wird nichts geöffnet, gespeichert oder gelöscht",
            default_yes=False)

    # --- Resume-Datei (nur Echtlauf) ---
    resume_path = args.resume if args.resume else get_resume_file_path(start_dir)
    resume_set: set = set()
    if not dry_run:
        resume_set = load_resume_set(resume_path)
        if resume_set and not auto_mode and not args.resume:
            print(f"\n♻  Resume-Datei eines früheren Laufs gefunden: "
                  f"{len(resume_set)} bereits verarbeitete Datei(en).")
            if not ask_yes_no("  Lauf fortsetzen (J) oder von vorn beginnen (n)?"):
                delete_resume_file(resume_path)
                resume_set = set()

    print()
    print("=" * 66)
    print("  KONFIGURATION")
    print("=" * 66)
    print(f"  Verzeichnis:          {start_dir}")
    print(f"  Neue Schriftart:      {NEW_FONT_NAME}")
    print(f"  Modus:                {'Vorab zählen (ETA)' if count_first else 'Generator (kein ETA)'}")
    print(f"  Fortschrittsbalken:   {'an' if show_progress else 'aus (inline)'}")
    if dry_run:
        print("  PROBELAUF:            JA – es wird nichts geändert")
    if resume_set:
        print(f"  Resume:               {len(resume_set)} Datei(en) werden übersprungen")
    if metadata_types:
        print(f"  Metadaten entfernen:  Ja ({len(metadata_types)} Typen)")
        for t in metadata_types:
            print(f"    wdRDI {t:2d}: {METADATA_TYPES[t]}")
    else:
        print("  Metadaten entfernen:  Nein")
    print("  Verarbeitungsumfang:")
    print("    * Formatvorlagen (Styles)")
    print("    * StoryRanges (Haupttext, Kopf-/Fußzeilen, Textrahmen)")
    print("    * Shapes in Kopf-/Fußzeilen (via ShapeRange je Story)")
    print("    * Tabellen")
    print("    * Fuß- und Endnoten")
    print("    * Content Controls (Inhaltssteuerelemente)")
    print("    * Revisionen (Track Changes Ranges)")
    print("    * Kommentare (sofern nicht via Metadaten-Typ 1 gelöscht)")
    print("    * SmartArt (via AllNodes / TextFrame2)")
    print("    * WordArt (via TextEffectFormat)")
    print("    * Multi-Script-Fonts (NameAscii / NameOther / NameBi / NameFarEast)")
    print("    * Konvertierung: .doc/.dot -> .docx/.dotx")
    print("=" * 66)

    if not auto_mode and not ask_yes_no("\nJetzt starten?"):
        print("Abgebrochen.")
        sys.exit(0)

    # --- Probelauf: ohne Word, ohne Aenderungen, dann Ende ---
    if dry_run:
        dr_start = datetime.now()
        try:
            dry_run_directory(start_dir)
        except Exception as e:
            print(f"\nKRITISCHER FEHLER im Probelauf: {e}")
            log_error("GLOBAL", e)
        print(f"\n  Dauer:       {str(datetime.now() - dr_start).split('.')[0]}")
        print(f"  Detail-Log:  {os.path.abspath(DETAILED_LOG_FILE)}")
        _warte_auf_taste(auto_mode)
        sys.exit(0)

    # --- Word Smoke-Test (nach Bestaetigung, vor Hauptlauf) ---
    # Prueft mit einer kurzlebigen Word-Instanz, ob Word eine .docx im
    # TEMP_BASE_PATH ohne Geschuetzte-Ansicht-Hänger oeffnen kann. Bei
    # Hänger killt der Watchdog Word nach 25 s. Ohne diesen Test wuerde
    # ein falsch konfiguriertes Trust-Center die spaeteren Open-Aufrufe
    # pro Datei minutenlang blockieren.
    # Kategorien: "com" = COM-Subsystem-Problem, "wd" = Word/Trust-Center.
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
            print("    • Antivirus / EDR blockiert COM-Aufrufe auf WINWORD.EXE")
            print("=" * 66)
            file_logger.error(f"COM-Subsystem nicht verfuegbar: {smoke_msg}")
            log_warn_text = "Abbruch durch Benutzer nach COM-Subsystem-Fehler."
            log_ignore_text = "COM-Subsystem-Fehler vom Benutzer ignoriert – Fortsetzung."
        else:
            print("  ⚠  WORD-SMOKE-TEST FEHLGESCHLAGEN")
            print("=" * 66)
            print(f"  Grund: {smoke_msg}")
            print()
            print("  Mögliche Ursachen:")
            print("    • Temp-Ordner nicht als vertrauenswürdiger Speicherort eingetragen")
            print(f"      ({TEMP_BASE_PATH})")
            print("    • Geschützte Ansicht für unsichere Speicherorte noch aktiv")
            print("    • Word/Office-Profil beschädigt oder fehlende Desktop-Ordner")
            print("=" * 66)
            file_logger.error(f"Word-Smoke-Test fehlgeschlagen: {smoke_msg}")
            log_warn_text = "Abbruch durch Benutzer nach Smoke-Test-Warnung."
            log_ignore_text = "Smoke-Test-Warnung vom Benutzer ignoriert – Fortsetzung."

        # Im Automatikmodus darf hier NICHT gefragt werden. ask_yes_no()
        # blockiert an input(): bei einem geplanten Task mit geschlossenem
        # stdin lieferte es default_yes=False und beendete den Lauf mit
        # sys.exit(0) - der Taskplaner meldete Erfolg, obwohl keine einzige
        # Datei angefasst wurde. Haengt stdin dagegen an einer Konsole oder
        # Pipe, blieb der Task unbegrenzt an der Rueckfrage stehen. Beides
        # widerspricht der eigenen Hilfe zu --auto ('Trust-Center-Check und
        # interaktive Rueckfragen ueberspringen'). Die Schwesterskripte
        # machen es richtig: 3a steigt mit exit 2 aus, 4b kapselt den Block.
        # Exit-Code 2 (nicht 0), damit der Taskplaner den Fehlschlag sieht.
        if auto_mode:
            print("\n❌  Abbruch im AUTO-Modus nach fehlgeschlagenem Trust-Center-Test.")
            file_logger.error("Abbruch (auto_mode) nach fehlgeschlagenem Trust-Center-Test.")
            sys.exit(2)

        if not ask_yes_no("\nTrotzdem fortfahren? (Timeout-Risiko pro Datei!)",
                          default_yes=False):
            print("Abgebrochen.")
            file_logger.error(log_warn_text)
            sys.exit(0)
        file_logger.error(log_ignore_text)
    else:
        detail_logger.info("Word-Smoke-Test erfolgreich.")
        print("  → Smoke-Test OK.")

    signal.signal(signal.SIGINT,  _signal_handler)
    signal.signal(signal.SIGTERM, _signal_handler)

    # --- Einzelinstanz-Schutz ---
    # Diese Sperre gab es bisher nur in 3a-3c. Ohne sie konnten zwei
    # Laeufe gleichzeitig ueber denselben Bestand gehen und sich
    # gegenseitig die Temp-Kopien und Zieldateien wegziehen.
    # Freigabe ueber atexit, damit sie auch bei sys.exit greift.
    _sperre = None
    if gem is not None:
        _sperre = gem.Einzelinstanz("4a_ersetze_font_in_word")
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

    try:
        if gem is not None and hasattr(gem, "recent_wurzel_hinzufuegen"):
            gem.recent_wurzel_hinzufuegen(start_dir)
        stats_result = process_directory(
            start_dir, NEW_FONT_NAME, metadata_types, count_first,
            show_progress=show_progress,
            resume_path=resume_path, resume_set=resume_set)
    except Exception as e:
        print(f"\nKRITISCHER FEHLER: {e}")
        log_error("GLOBAL", e)
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
        _total = (stats_result.get("SUCCESS", 0)
                  + stats_result.get("UNCHANGED", 0)
                  + stats_result.get("SKIPPED", 0)
                  + stats_result.get("ERROR", 0))
        print("  STATISTIK:")
        print(f"    [OK] Erfolgreich:    {stats_result.get('SUCCESS', 0)}")
        print(f"    [==] Unverändert:    {stats_result.get('UNCHANGED', 0)}")
        print(f"    [->] Übersprungen:   {stats_result.get('SKIPPED', 0)}")
        for _label, _key in (("in Bearbeitung (gesperrt)", "SKIPPED_LOCKED"),
                             ("OLE-verschlüsselt",         "SKIPPED_ENCRYPTED"),
                             ("passwortgeschützt",         "SKIPPED_PASSWORD")):
            if stats_result.get(_key, 0) > 0:
                print(f"          • {_label}: {stats_result[_key]}")
        print(f"    [!!] Fehler:         {stats_result.get('ERROR', 0)}")
        print(f"         GESAMT:          {_total}")

        write_run_summary(stats_result, start_dir, NEW_FONT_NAME,
                          metadata_types, start_time, datetime.now())

    print()
    print("!" * 66)
    print("HINWEIS: TRUST-CENTER EINSTELLUNGEN ZURÜCKSETZEN")
    print("Die zu Beginn vorgenommenen Änderungen im Trust-Center sollten")
    print("jetzt wieder rückgängig gemacht werden:")
    print("Word öffnen → Datei → Optionen → Trust Center → "
          "'Einstellungen für das Trust Center...'")
    print("1. GESCHÜTZTE ANSICHT: alle 3 Optionen wieder AKTIVIEREN")
    print("2. MAKROEINSTELLUNGEN: auf ursprüngliche Einstellung zurücksetzen")
    print("   (z.B. 'Alle Makros mit Benachrichtigung deaktivieren')")
    print("3. VERTRAUENSWÜRDIGE SPEICHERORTE:")
    print("   - Temp-Ordner des Skripts wieder ENTFERNEN:")
    print(f"     {TEMP_BASE_PATH}")
    print("   - ☐ 'Vertrauenswürdige Speicherorte im Netzwerk zulassen'")
    print("     wieder DEAKTIVIEREN (falls vorher nicht gesetzt)")
    print("!" * 66)

    print()
    print(f"  Fehler-Log:  {os.path.abspath(LOG_FILE)}")
    print(f"  Detail-Log:  {os.path.abspath(DETAILED_LOG_FILE)}")
    print(f"  Run-Summary: {os.path.abspath(RUN_SUMMARY_FILE)}")
    if os.path.exists(CONVERSIONS_CSV):
        print("  Konvertierungs-CSV (alte → neue Pfade):")
        print(f"               {os.path.abspath(CONVERSIONS_CSV)}")
    if resume_path and os.path.exists(resume_path):
        print(f"  Resume-Log:  {os.path.abspath(resume_path)}")
        print("               (erhalten – Neustart setzt dort fort)")
    print("=" * 66)

    _safe_cleanup_temp_process()
    _cleanup_windows_temp()
    _warte_auf_taste(auto_mode)
