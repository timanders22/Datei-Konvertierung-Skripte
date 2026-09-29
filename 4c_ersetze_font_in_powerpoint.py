# -*- coding: utf-8 -*-
# ==================================================================
# POWERPOINT FONT-ERSETZUNG
# ==================================================================
# ANLEITUNG:
# 1. Python installieren: https://www.python.org/
# 2. In der Konsole (CMD): pip install pywin32 tqdm psutil
# 3. Falls pywin32 erstmals installiert: python Scripts/pywin32_postinstall.py -install
# 4. In der Konsole (CMD): python 4c_ersetze_font_in_powerpoint.py
#
# WICHTIG: POWERPOINT SICHERHEITSEINSTELLUNGEN (TRUST CENTER)
# ------------------------------------------------------------------
# PowerPoint öffnen -> Datei -> Optionen -> Trust Center ->
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
#    (Erforderlich für korrekte .pptm/.pptx-Unterscheidung)
#
# 3. VERTRAUENSWÜRDIGE SPEICHERORTE:
#    ☑ Vertrauenswürdige Speicherorte im Netzwerk zulassen
#    Folgende Speicherorte hinzufügen:
#    • %LOCALAPPDATA%\Dateimigration-Arbeitskopien (temporärer Arbeitsordner des Skripts)
#    • Q:\ oder \\server\dfs (Quelldateien)
#    Jeweils: ☑ Unterordner ... sind ebenfalls vertrauenswürdig
#
# 4. DATENSCHUTZOPTIONEN:
#    ☐ Beim Öffnen automatisch verknüpfte Daten aktualisieren
#    → DEAKTIVIEREN (verhindert Popups bei eingebetteten Objekten)
#
# HINWEIS FÜR PyInstaller-Build (Umwandlung in .exe):
#   # Einmalig: PyInstaller installieren
#   pip install pyinstaller
#
#   # Build als ein einzeiliger Befehl (am sichersten in jeder Shell):
#   python -m PyInstaller --onefile --noupx --noconfirm --clean --console --icon="python_icon.ico" --name "4c_ersetze_font_in_powerpoint" --collect-submodules win32com --hidden-import pywintypes --hidden-import pythoncom --hidden-import win32api --hidden-import win32com.client --hidden-import win32process --hidden-import win32file --hidden-import win32con --hidden-import win32timezone --hidden-import winreg --hidden-import psutil --hidden-import tqdm 4c_ersetze_font_in_powerpoint.py
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
#     Interpreters (64-Bit-Python -> x64-EXE; 32-Bit-Python -> x86-EXE). PowerPoint
#     registriert seinen COM-Server als LocalServer32 (out-of-process,
#     POWERPNT.EXE); das Windows-COM-Subsystem marshallt Aufrufe zwischen
#     x64-Aufrufer und x86-Server (und umgekehrt) automatisch über DCOM/LRPC.
#     Eine x64-EXE arbeitet daher auch mit 32-Bit-PowerPoint zusammen (und umgekehrt).
#     Bitness-Match liefert minimal bessere Performance, ist aber nicht
#     erforderlich. 64-Bit-Python ist eine sichere Default-Wahl, da modernes
#     Office (2019+, M365) standardmäßig x64 ist.
#
# Stand: 11.06.2026
# ==================================================================

import os
import sys
import argparse
import signal
import shutil
import logging
import time
import uuid
import threading
import tempfile
import importlib.util
from datetime import datetime
from typing import Optional, Tuple

try:
    import winreg
except ImportError:
    winreg = None

# ==================================================================
# Konsolen-Encoding (UTF-8) für Box-Drawing-Zeichen & Umlaute
# ==================================================================
if sys.platform == "win32":
    try:
        os.system("chcp 65001 >nul 2>&1")
    except Exception:
        pass
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

import pythoncom
import win32com.client
import win32file
import win32process
import win32api
import win32con
import pywintypes
import stat
import psutil
from tqdm import tqdm

# ==================================================================
# COM-Konstanten
# ==================================================================
COM_TRUE  = -1
COM_FALSE =  0

MSO_TYPE_GROUP = 6
MSO_TYPE_TEXT_EFFECT = 15      # klassisches WordArt

PP_FORMAT_PPTX = 24
PP_FORMAT_PPTM = 25
PP_FORMAT_POTX = 26
PP_FORMAT_POTM = 27
PP_FORMAT_PPSX = 28
PP_FORMAT_PPSM = 29

PP_WINDOW_NORMAL    = 1
PP_WINDOW_MINIMIZED = 2
PP_ALERTS_NONE      = 1

MSO_FEATURE_INSTALL_NONE    = 0
MSO_AUTOMATION_SECURITY_LOW = 1

RPC_E_CALL_REJECTED   = -2147418111
STG_E_SHAREVIOLATION  = -2147286784
E_SHARING_VIOLATION   = -2147024864
COM_RETRY_CODES = {
    RPC_E_CALL_REJECTED,
    STG_E_SHAREVIOLATION,
    E_SHARING_VIOLATION,
}

OPENXML_FORMAT_MAP = {
    ".pptx": PP_FORMAT_PPTX,
    ".pptm": PP_FORMAT_PPTM,
    ".potx": PP_FORMAT_POTX,
    ".potm": PP_FORMAT_POTM,
    ".ppsx": PP_FORMAT_PPSX,
    ".ppsm": PP_FORMAT_PPSM,
}

# ==================================================================
# Metadaten-Optionen
# ==================================================================
BUILTIN_PROPS_TO_CLEAR = {
    "Author":              "Autor (Ersteller)",
    "Last author":         "Letzter Bearbeiter",
    "Title":               "Titel",
    "Subject":             "Betreff",
    "Keywords":            "Stichwörter",
    "Comments":            "Notizen/Kommentare (Eigenschaft)",
    "Company":             "Firma",
    "Manager":             "Vorgesetzter",
    "Category":            "Kategorie",
    "Hyperlink base":      "Hyperlink-Basis",
    "Presentation format": "Präsentationsformat",
}

METADATA_OPTIONS = {
    "props":    "Integrierte Dokumenteigenschaften (Autor, Titel, Firma ...)",
    "custom":   "Benutzerdefinierte Dokumenteigenschaften (alle löschen)",
    "comments": "Folien-Kommentare / Anmerkungen (INHALT wird gelöscht!)",
}

# ==================================================================
# Konfiguration (statische Konstanten)
# ==================================================================
DEFAULT_FONT_NAME    = "Arial"
COM_RESTART_INTERVAL = 25

MAX_PATH_LEN     = 240
WINDOWS_MAX_PATH = 260

# Watchdog-Timeouts fuer COM-Aufrufe in der Hauptverarbeitung. Werte sind
# grosszuegig bemessen, um grosse Praesentationen (viele Folien, eingebettete
# Medien) nicht voreilig abzuschiessen, fangen aber haengende Modal-Dialoge
# zuverlaessig ab (Reparatur-Prompt, Verknuepfungs-Update, DRM-Schutz,
# Trust-Center-Popup trotz AutomationSecurity=ForceDisable). Smoke-Test
# (25s) bleibt unveraendert, weil dort nur eine winzige Test-Datei
# verarbeitet wird.
PPT_OPEN_TIMEOUT = 180.0   # Sekunden fuer Presentations.Open
PPT_SAVE_TIMEOUT = 240.0   # Sekunden fuer SaveAs / Save (kann bei grossen
                           # Praesentationen laenger dauern als Open)

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

_SCRIPT_DIR = _resolve_script_directory()
_RUN_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

LOG_FILE          = os.path.join(
    _SCRIPT_DIR, f"4c_ersetze_font_in_powerpoint_{_RUN_TIMESTAMP}.log")
DETAILED_LOG_FILE = os.path.join(
    _SCRIPT_DIR, f"4c_ersetze_font_in_powerpoint_detailed_{_RUN_TIMESTAMP}.log")
# Die Done-Liste traegt bewusst KEINEN Zeitstempel: sie dient als
# Resume-Liste ueber mehrere Laeufe hinweg.
DONE_FILE         = os.path.join(_SCRIPT_DIR, "4c_ersetze_font_in_powerpoint_done.log")

AV_SCANNER_RETRIES = 10
AV_SCANNER_DELAY   = 0.5

MAX_GROUP_DEPTH = 20

REG_USER_SHELL_FOLDERS = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
REG_DOWNLOADS_GUID     = "{374DE290-123F-4565-9164-39C4925E467B}"


# ==================================================================
# Pfad-Hilfsfunktionen
# ==================================================================

def _long_path(path: str) -> str:
    if path.startswith("\\\\?\\"):
        return path
    path = os.path.abspath(path)
    if path.startswith("\\\\"):
        return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path


def sanitize_path(raw: str) -> str:
    path = raw.strip().strip('"').strip("'")
    if len(path) == 2 and path[1] == ":":
        path = path + "\\"
    return path


# ==================================================================
# Known-Folder-Resolver (Registry mit Fallback, OneDrive-kompatibel)
# ==================================================================

def _get_known_folder(reg_name: str, fallback: str) -> str:
    if winreg is not None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_USER_SHELL_FOLDERS) as key:
                value, _ = winreg.QueryValueEx(key, reg_name)
                value = os.path.expandvars(value)
                if os.path.isdir(value):
                    return value
        except (OSError, FileNotFoundError):
            pass
    return fallback


def resolve_documents_folder() -> str:
    return _get_known_folder(
        "Personal",
        os.path.join(os.path.expanduser("~"), "Documents"),
    )


def resolve_desktop_folder() -> str:
    return _get_known_folder(
        "Desktop",
        os.path.join(os.path.expanduser("~"), "Desktop"),
    )


def resolve_downloads_folder() -> str:
    return _get_known_folder(
        REG_DOWNLOADS_GUID,
        os.path.join(os.path.expanduser("~"), "Downloads"),
    )


# ==================================================================
# Abgeleitete Pfade (zur Importzeit aufgelöst)
# ==================================================================
# Gemeinsamer Arbeitsordner der Office-Skripte. Bis 29.09.2026 lag er in
# "Dokumente"; bei OneDrive-Ordnersicherung wanderte so jede Arbeitskopie
# in die Cloud. %LOCALAPPDATA% wird nie umgeleitet.
ARBEITS_BASIS = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP")
    or resolve_documents_folder(),
    "Dateimigration-Arbeitskopien")
TEMP_BASE_PATH = os.path.join(ARBEITS_BASIS, "4c_ersetze_font_in_powerpoint")
# Run-spezifischer Unterordner (UUID) erlaubt parallele Laeufe. Als
# vertrauenswuerdiger Speicherort wird der stabile TEMP_BASE_PATH
# eingetragen (mit Unterordnern).
TEMP_PROCESS_PATH = os.path.join(TEMP_BASE_PATH, uuid.uuid4().hex)

# Verzeichnisse, die bei der Suche NICHT betreten werden. Ohne diese Liste
# wurden auch Praesentationen im Papierkorb und in Schattenkopien
# verarbeitet - ihre Schriftart geaendert und das Original ersetzt.
# Identisch zu 4a/4b.
EXCLUDE_DIR_NAMES = {"$recycle.bin", "system volume information",
                     "~snapshot", ".snapshot"}

FILE_ATTRIBUTE_REPARSE_POINT = 0x0400


# ==================================================================
# Logging (Dual-Logger)
# ==================================================================

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
detail_logger.info(
    f"=== Skript-Start | PID {os.getpid()} | "
    f"Skript-Verzeichnis: {_SCRIPT_DIR} ==="
)


# ==================================================================
# Globale PPT-Referenz (Signal-Handler / Ctrl+C)
# ==================================================================
ppt_app_global: Optional[win32com.client.CDispatch] = None
ppt_pid_global: Optional[int] = None
# Prozess-Erstellungszeit der Skript-PowerPoint-Instanz. Schuetzt
# zusammen mit dem Prozessnamen-Check vor PID-Recycling: Stirbt
# PowerPoint waehrend einer langen Watchdog-Wartezeit und vergibt
# Windows die PID an eine neue, vom Benutzer geoeffnete PowerPoint-
# Sitzung, unterscheidet nur die Erstellungszeit die beiden Prozesse.
ppt_create_time_global: Optional[float] = None
ppt_orig_security_global = None
# Die gerade vom Skript bearbeitete Praesentation (fuer den Signal-Handler:
# erst sie schliessen, dann Presentations.Count pruefen).
_eigene_praesentation_global = None


# ==================================================================
# Signal-Handler (Ctrl+C / SIGTERM)
# ==================================================================

def _signal_handler(sig, frame) -> None:
    print("\n\n*** ABBRUCH durch Benutzer (Ctrl+C) - räume auf ...")

    # Erst die EIGENE, gerade bearbeitete Praesentation ohne Speichern
    # schliessen, dann ueber _quit_ppt_instance beenden: das beendet nur,
    # wenn danach keine Praesentation mehr offen ist (Einzelinstanz - der
    # Anwender kann waehrend des Laufs Dateien in DIESER Instanz geoeffnet
    # haben). Frueher: Quit() + Kill ohne jede Pruefung.
    if _eigene_praesentation_global is not None:
        try:
            _eigene_praesentation_global.Close()
        except Exception as _e:
            detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")
    _quit_ppt_instance(ppt_app_global, ppt_pid_global,
                       ppt_orig_security_global, ppt_create_time_global)

    # Best effort Cache-Cleanup beim Strg+C-Abbruch (analog 3c).
    try:
        _cleanup_ppt_inetcache()
    except Exception as _e:
        detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")
    try:
        _cleanup_user_recent()
    except Exception as _e:
        detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")

    _safe_cleanup_temp()

    try:
        pythoncom.CoUninitialize()
    except Exception as _e:
        detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")

    sys.exit(1)


# ==================================================================
# Hilfsfunktionen
# ==================================================================

def log_error(file_path: str, exc: Exception) -> None:
    file_logger.error(f"Datei: {file_path}\n  -> {exc}\n")
    detail_logger.error(f"Datei: {file_path}\n  -> {exc}\n")


# ==================================================================
# Startwarnung: laufende PowerPoint-Sitzungen
# ==================================================================
def find_running_powerpoint_pids() -> list:
    """PIDs aller PowerPoint-Prozesse des angemeldeten Benutzers."""
    pids = []
    try:
        me = psutil.Process().username().lower().split("\\")[-1].split("@")[0]
    except Exception:
        return pids
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            nm = proc.info.get("name") or ""
            if "POWERPNT.EXE" not in nm.upper():
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


def _alle_powerpoint_pids() -> set:
    """PIDs ALLER POWERPNT.EXE-Prozesse (jeder Benutzer) - Momentaufnahme
    unmittelbar vor DispatchEx. Was hier auftaucht, ist per Definition
    NICHT die eigene Instanz. Im Zweifel (Lesefehler) lieber zu viel
    als zu wenig aufnehmen - ein zu Unrecht fremd erklaerter Prozess
    fuehrt nur zum Abbruch, nie zu einem Kill."""
    found = set()
    try:
        for p in psutil.process_iter(["pid", "name"]):
            try:
                nm = p.info.get("name") or ""
                if "POWERPNT.EXE" in nm.upper():
                    found.add(p.info["pid"])
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
    except Exception as _e:
        detail_logger.debug(f"_alle_powerpoint_pids: Exception verworfen: {_e!r}")
    return found


class FremdePowerPointSitzung(RuntimeError):
    """DispatchEx hat eine PowerPoint-Instanz geliefert, die nicht sicher
    die eigene ist. Sie wird weder verstellt noch beendet; der Lauf bricht ab."""


# Wird gesetzt, wenn der Hauptlauf wegen einer fremden PowerPoint-Sitzung
# abbrechen musste - fuer einen Exitcode != 0 am Skriptende (--auto).
_abbruch_fremde_sitzung = False


def pruefe_powerpoint_geschlossen(auto_mode: bool = False) -> None:
    """Bricht ab, solange eine PowerPoint-Sitzung des Anwenders laeuft.

    PowerPoint ist eine Einzelinstanz-Anwendung. Vom Auftraggeber
    gemessen: Bei laufender Sitzung liefert DispatchEx("PowerPoint.
    Application") GENAU DIESE Sitzung, ihr HWND ergibt deren PID. Bisher
    wurde dann trotzdem weitergearbeitet: die Anwendersitzung wurde
    minimiert, AutomationSecurity/DisplayAlerts verstellt, alle 25 Dateien
    per Quit() beendet und bei Timeout per kill() abgeschossen - mitsamt
    ungespeicherter Arbeit. Die im alten Hinweis versprochene
    "Momentaufnahme der fremden Prozess-IDs" gab es in 4c nie.

    Deshalb wird bei laufendem PowerPoint gar nicht per COM gearbeitet:
      interaktiv: Hinweis, PowerPoint schliessen lassen, erneut pruefen
                  (oder abbrechen)
      --auto:     Abbruch mit Exitcode 2 - der Aufgabenplaner sieht den
                  Fehlschlag.
    Nach dem Start schuetzt zusaetzlich die Momentaufnahme in
    _start_ppt_instance (PID lief schon vorher -> fremd -> Abbruch).
    """
    while True:
        pids = find_running_powerpoint_pids()
        if not pids:
            return
        pids_str = ", ".join(str(p) for p in pids)

        if auto_mode:
            print()
            print(f"[X] PowerPoint laeuft bereits (PID: {pids_str}) - Abbruch.")
            print("    PowerPoint ist eine Einzelinstanz: das Skript wuerde sich an")
            print("    die laufende Sitzung haengen. Bitte PowerPoint schliessen und")
            print("    den Lauf neu starten.")
            file_logger.error(
                f"Abbruch (auto_mode): PowerPoint laeuft bereits (PID: {pids_str}).")
            detail_logger.error(
                f"Abbruch (auto_mode): PowerPoint laeuft bereits (PID: {pids_str}).")
            sys.exit(2)

        print()
        print("=" * 66)
        print("  PowerPoint laeuft bereits - so kann das Skript NICHT starten")
        print("=" * 66)
        print(f"  Gefundene PowerPoint-Prozesse (PID): {pids_str}")
        print()
        print("  PowerPoint ist eine Einzelinstanz-Anwendung: Das Skript wuerde")
        print("  sich an Ihre laufende Sitzung haengen, sie minimieren, Warnungen")
        print("  abschalten und sie beim Neustart bzw. Timeout beenden - mit")
        print("  Verlust ungespeicherter Arbeit.")
        print()
        print("  Bitte speichern Sie Ihre Praesentationen und schliessen Sie")
        print("  PowerPoint vollstaendig (auch im Task-Manager pruefen).")
        print("=" * 66)
        print()
        if not ask_yes_no("PowerPoint geschlossen - erneut pruefen?",
                          default_yes=False):
            print("Abgebrochen. Bitte PowerPoint schliessen und neu starten.")
            detail_logger.info(
                f"Abbruch durch Benutzer: PowerPoint laeuft (PID: {pids_str}).")
            sys.exit(0)


def _kill_powerpoint_by_pid(pid: Optional[int],
                            create_time: Optional[float] = None) -> None:
    if pid is None:
        return
    # Kill NUR mit bekannter Erstellungszeit: Ohne sie laesst sich die
    # eigene Instanz nicht von einer spaeter gestarteten Anwendersitzung
    # unterscheiden, die dieselbe (recycelte) PID bekommen hat. Frueher
    # wurde bei create_time=None ungeprueft getoetet.
    if create_time is None:
        detail_logger.warning(
            f"PPT-Prozess PID {pid}: Erstellungszeit unbekannt – kein Kill")
        return
    try:
        proc = psutil.Process(pid)
        if not proc.is_running() or "POWERPNT" not in (proc.name() or "").upper():
            return
        # Erstellungszeit-Check schuetzt zusaetzlich zum Namens-Check
        # vor PID-Recycling auf eine NEUE PowerPoint-Sitzung des
        # Benutzers, die der Namens-Check allein nicht erkennen wuerde.
        try:
            if proc.create_time() != create_time:
                detail_logger.debug(
                    f"PID {pid} hat abweichende Erstellungszeit "
                    f"– kein Kill (PID-Recycling)")
                return
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            return
        proc.kill()
        try:
            proc.wait(timeout=5)
        except psutil.TimeoutExpired:
            detail_logger.warning(
                f"PPT-Prozess PID {pid} reagiert nicht auf kill – erneuter Versuch")
            try:
                proc.terminate()
                proc.wait(timeout=5)
            except Exception as _e:
                detail_logger.debug(f"_kill_powerpoint_by_pid: Exception verworfen: {_e!r}")
        try:
            if proc.is_running():
                detail_logger.error(
                    f"PPT-Prozess PID {pid} konnte nicht beendet werden – "
                    f"Geister-Instanz möglich")
            else:
                detail_logger.debug(f"PPT-Prozess beendet: PID {pid}")
        except psutil.NoSuchProcess:
            detail_logger.debug(f"PPT-Prozess beendet: PID {pid}")
    except psutil.NoSuchProcess:
        pass
    except psutil.AccessDenied:
        detail_logger.warning(f"Kein Zugriff auf PPT-Prozess PID {pid}")
    except Exception as e:
        detail_logger.warning(f"PPT-Prozess PID {pid} nicht beendbar: {e}")


def _get_proc_create_time(pid: Optional[int]) -> Optional[float]:
    if pid is None:
        return None
    try:
        return psutil.Process(pid).create_time()
    except Exception:
        return None


def _safe_cleanup_temp() -> None:
    """Loescht den Run-spezifischen Temp-Ordner - AUSSER es liegen
    RESCUE_-Dateien darin (Datenverlust-Schutz: sie enthalten die
    einzige Kopie fehlgeschlagener Stage-2-Ergebnisse)."""
    if not os.path.exists(TEMP_PROCESS_PATH):
        return
    try:
        rescue_files = [
            f for f in os.listdir(TEMP_PROCESS_PATH)
            if f.startswith("RESCUE_")
        ]
        if rescue_files:
            detail_logger.warning(
                f"GESCHÜTZTE DATEIEN IM TEMP-ORDNER – Ordner wird NICHT gelöscht!\n"
                f"  Pfad: {TEMP_PROCESS_PATH}\n"
                f"  Dateien ({len(rescue_files)}):\n"
                + "\n".join(f"    * {f}" for f in rescue_files)
                + "\n  → Bitte manuell sichern, dann Ordner löschen.")
            print("\n  GESCHÜTZTE DATEIEN vorhanden – Temp-Ordner NICHT gelöscht:")
            print(f"   {TEMP_PROCESS_PATH}")
            for f in rescue_files:
                print(f"   * {f}")
            print("   → Bitte manuell sichern, dann Ordner löschen.")
        else:
            shutil.rmtree(TEMP_PROCESS_PATH)
    except Exception as e:
        detail_logger.warning(f"Aufräumen fehlgeschlagen: {e}")


def check_required_modules() -> None:
    required = {
        "win32com.client": "pywin32",
        "win32file":       "pywin32",
        "psutil":          "psutil",
        "tqdm":            "tqdm",
    }
    missing = [pkg for mod, pkg in required.items()
               if not importlib.util.find_spec(mod.split(".")[0])]
    if missing:
        print("=" * 66)
        print("FEHLENDE MODULE:")
        for pkg in dict.fromkeys(missing):
            print(f"   pip install {pkg}")
        print("=" * 66)
        sys.exit(1)


def _powerpoint_exe_aus_registry() -> Optional[str]:
    """Pfad zu POWERPNT.EXE aus der Registry - OHNE PowerPoint zu starten.

    Quellen: COM-Registrierung (PowerPoint.Application -> CLSID ->
    LocalServer32), sonst App Paths. Nachgemessen auf dem Arbeitsplatz
    (Office 2024, Klick-und-Los): LocalServer32 =
    '...\\Office16\\POWERPNT.EXE /AUTOMATION', App Paths = derselbe Pfad."""
    if winreg is None:
        return None
    ansichten = [0]
    if hasattr(winreg, "KEY_WOW64_64KEY"):
        ansichten += [winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY]

    def _lies(root, pfad, flag):
        try:
            with winreg.OpenKey(root, pfad, 0, winreg.KEY_READ | flag) as k:
                return winreg.QueryValueEx(k, "")[0]
        except OSError:
            return None

    kandidaten = []
    for flag in ansichten:
        clsid = _lies(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID", flag)
        if clsid:
            server = _lies(winreg.HKEY_CLASSES_ROOT,
                           rf"CLSID\{clsid}\LocalServer32", flag)
            if server:
                kandidaten.append(server)
        app = _lies(winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe",
                    flag)
        if app:
            kandidaten.append(app)
    for roh in kandidaten:
        s = os.path.expandvars(str(roh)).strip()
        if s.startswith('"'):
            s = s[1:].split('"', 1)[0]
        else:
            idx = s.lower().find(".exe")
            if idx >= 0:
                s = s[:idx + 4]
        if s and os.path.isfile(s):
            return s
    return None


def check_powerpoint_installed() -> None:
    # Installationspruefung NUR ueber Registry/Dateisystem. Frueher lief
    # hier DispatchEx + Quit(): bei laufender Anwendersitzung liefert
    # DispatchEx (Einzelinstanz) genau diese - das Quit() traf dann die
    # Sitzung des Anwenders, noch VOR der Laufpruefung und auch im
    # Probelauf, der laut Hilfetext PowerPoint gar nicht startet. Ob
    # PowerPoint per COM wirklich arbeitet, prueft spaeter der Smoke-Test
    # (nach der Laufpruefung).
    exe = _powerpoint_exe_aus_registry()
    if exe is None:
        print("=" * 66)
        print("FEHLER: Microsoft PowerPoint wurde nicht gefunden.")
        print("Bitte stellen Sie sicher, dass Microsoft PowerPoint installiert ist")
        print("(Office 2019 oder 2024, deutsch oder englisch).")
        print("Geprueft: Registry PowerPoint.Application (LocalServer32) und App Paths.")
        print("=" * 66)
        sys.exit(1)
    detail_logger.info(f"PowerPoint gefunden (Registry): {exe}")


def get_file_timestamps(file_path: str) -> Tuple[float, float, float]:
    try:
        st = os.stat(_long_path(file_path))
        return st.st_atime, st.st_mtime, st.st_ctime
    except Exception as e:
        t = time.time()
        detail_logger.warning(f"Zeitstempel nicht lesbar: {file_path} – {e}")
        return t, t, t


def set_file_timestamps(file_path: str, atime: float, mtime: float, ctime: float) -> bool:
    try:
        handle = win32file.CreateFile(
            _long_path(file_path),
            win32file.GENERIC_WRITE,
            win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
            None,
            win32file.OPEN_EXISTING,
            win32file.FILE_ATTRIBUTE_NORMAL,
            None,
        )
        try:
            win32file.SetFileTime(
                handle,
                pywintypes.Time(ctime),
                pywintypes.Time(atime),
                pywintypes.Time(mtime),
            )
        finally:
            handle.Close()
        detail_logger.debug(f"Zeitstempel wiederhergestellt: {file_path}")
        return True
    except Exception as e:
        detail_logger.warning(f"Zeitstempel nicht setzbar: {file_path} – {e}")
        return False


def _get_ppt_pid(ppt_app) -> Optional[int]:
    try:
        hwnd = ppt_app.HWND
        if hwnd:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return pid
    except Exception as _e:
        detail_logger.debug(f"_get_ppt_pid: Exception verworfen: {_e!r}")
    return None


def _com_call_with_retry(func, *args, retries: int = 5, delay: float = 1.0, **kwargs):
    for attempt in range(retries):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as e:
            hresult = getattr(e, "hresult", None)
            if hresult in COM_RETRY_CODES and attempt < retries - 1:
                detail_logger.debug(
                    f"COM-Retry ({hresult}), Versuch {attempt + 1}/{retries}")
                time.sleep(delay * (attempt + 1))
                continue
            raise
    raise RuntimeError("_com_call_with_retry: unerwartetes Ende ohne Ergebnis")


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
    return _long_path(path) if len(path) > MAX_PATH_LEN else path


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
# Datei-Sperren / AV-Scanner-Wartelogik
# ==================================================================

def _check_file_locked(file_path: str) -> bool:
    """Prueft, ob eine Datei von einem anderen Prozess gesperrt ist.

    READ-ONLY-TOLERANT: Frueher hat diese Funktion bei schreibgeschuetzten
    Dateien faelschlicherweise true zurueckgegeben (CreateFile mit
    GENERIC_WRITE schlaegt bei Read-Only-Files mit AccessDenied fehl).
    Folge: legitime, intakte schreibgeschuetzte .pptx wurden als "gesperrt"
    abgewiesen und nie verarbeitet.

    Strategie: Vor dem Write-Test pruefen, ob das Read-Only-Attribut
    gesetzt ist. Falls ja, faellt der Lock-Check auf einen reinen
    Read-Test zurueck (eine echte Sperre durch einen anderen Prozess
    verhindert auch das Lesen mit FILE_SHARE_NONE).
    """
    long_path = _long_path(file_path)

    # 1. Read-Only-Attribut pruefen - bei R/O nicht GENERIC_WRITE testen
    is_read_only = False
    try:
        attrs = win32api.GetFileAttributes(long_path)
        if attrs != -1 and (attrs & win32con.FILE_ATTRIBUTE_READONLY):
            is_read_only = True
    except Exception as _e:
        # Wenn GetFileAttributes scheitert, dem alten Pfad folgen
        # (Write-Test wird scheitern, aber das ist dann nicht durch R/O)
        detail_logger.debug(f"_check_file_locked: Exception verworfen: {_e!r}")

    try:
        if is_read_only:
            # Read-Test ohne Sharing: eine echte Sperre durch einen anderen
            # Prozess (z.B. offene PowerPoint-Instanz) wuerde auch das
            # Lesen blockieren. False Positives durch Read-Only sind so
            # ausgeschlossen.
            handle = win32file.CreateFile(
                long_path,
                win32file.GENERIC_READ,
                0,
                None,
                win32file.OPEN_EXISTING,
                win32file.FILE_ATTRIBUTE_NORMAL,
                None,
            )
        else:
            # Voller Read+Write-Test (Original-Verhalten fuer normale Dateien)
            handle = win32file.CreateFile(
                long_path,
                win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                0,
                None,
                win32file.OPEN_EXISTING,
                win32file.FILE_ATTRIBUTE_NORMAL,
                None,
            )
        handle.Close()
        return False
    except pywintypes.error:
        return True


def _is_locked_by_other_user(file_path: str) -> bool:
    try:
        d = os.path.dirname(file_path)
        b = os.path.basename(file_path)
        if not d or not b:
            return False
        owner = os.path.join(d, "~$" + b)
        try:
            return os.path.exists(_long_path(owner))
        except Exception:
            return False
    except Exception:
        return False


def _wait_until_unlocked(file_path: str,
                         retries: int = AV_SCANNER_RETRIES,
                         delay: float = AV_SCANNER_DELAY) -> bool:
    for attempt in range(retries):
        if not _check_file_locked(file_path):
            return True
        if attempt == 0:
            detail_logger.debug(f"Datei gesperrt (AV-Scanner?) – warte: {file_path}")
        time.sleep(delay)
    return not _check_file_locked(file_path)


def _safe_remove_with_retry(path: str,
                            retries: int = AV_SCANNER_RETRIES,
                            delay: float = AV_SCANNER_DELAY) -> bool:
    long_path = _long_path(path)
    # Read-Only-Attribut vorab entfernen: os.remove wirft auf Windows
    # bei schreibgeschuetzten Dateien IMMER PermissionError - ohne chmod
    # wuerden saemtliche Retries sinnlos fehlschlagen.
    try:
        os.chmod(long_path, stat.S_IWRITE)
    except Exception as _e:
        detail_logger.debug(f"_safe_remove_with_retry: Exception verworfen: {_e!r}")
    for attempt in range(retries):
        try:
            os.remove(long_path)
            return True
        except FileNotFoundError:
            return True
        except (PermissionError, OSError):
            if attempt < retries - 1:
                time.sleep(delay)
                continue
    return False


def _safe_move(src: str, dst: str) -> None:
    src_long = _long_path(src)
    dst_long = _long_path(dst)
    shutil.copy2(src_long, dst_long)
    _wait_until_unlocked(dst)
    if not _safe_remove_with_retry(src):
        detail_logger.warning(f"Quelldatei nicht löschbar nach Move: {src}")


def _replace_file_with_backup(src: str, dst: str) -> None:
    """Ersetzt dst durch src mit Backup-Schutz fuer die Zieldatei.

    HINTERGRUND: _safe_move basiert auf shutil.copy2, das die Zieldatei
    sofort im Modus 'wb' oeffnet und damit auf 0 Bytes trunkiert, BEVOR
    Daten fliessen. Bricht der Kopiervorgang ab (Netzwerk-Drop,
    AV-Scanner), waere die Zieldatei ohne Backup unwiderruflich
    zerstoert. Daher: Ziel per atomarem os.replace (Rename auf demselben
    Laufwerk) auf .bak sichern, bei Fehlschlag zurueckrollen, bei Erfolg
    .bak loeschen. Bleibt ein .bak nach fehlgeschlagenem Rollback
    liegen, steht der Pfad im Detail-Log."""
    dst_long = _long_path(dst)
    bak_long = None
    if os.path.exists(dst_long):
        try:
            os.chmod(dst_long, stat.S_IWRITE)
        except Exception as _e:
            detail_logger.debug(f"_replace_file_with_backup: Exception verworfen: {_e!r}")
        bak_long = _long_path(dst + f"_{uuid.uuid4().hex[:6]}.bak")
        os.replace(dst_long, bak_long)
    try:
        _safe_move(src, dst)
    # BaseException, nicht Exception: Der Signal-Handler beendet sich mit
    # sys.exit(), und das loest SystemExit aus - eine BaseException. Bei
    # Strg+C oder SIGTERM mitten im Kopiervorgang wurde der Rollback deshalb
    # UEBERSPRUNGEN, obwohl das Original oben bereits per os.replace auf den
    # .bak-Namen umbenannt war. Zurueck blieb am Originalnamen der von
    # shutil.copy2 sofort auf 0 Byte trunkierte Torso, waehrend der einzige
    # vollstaendige Bestand unter einem .bak-Namen lag, den weder der
    # Anwender noch ein spaeterer Lauf findet oder aufraeumt. Nachgestellt:
    # 'Vortrag.pptx' 2 Byte statt 1400, daneben 'Vortrag.pptx_99688a.bak'.
    # Das abschliessende 'raise' reicht SystemExit unveraendert weiter, der
    # Abbruch wirkt also wie bisher - nur eben mit intakter Datei.
    except BaseException:
        if bak_long and os.path.exists(bak_long):
            try:
                if os.path.exists(dst_long):
                    _safe_remove_with_retry(dst)
                os.replace(bak_long, dst_long)
                detail_logger.warning(
                    f"Move fehlgeschlagen – Zieldatei aus Backup "
                    f"wiederhergestellt: {dst}")
            except Exception as e_restore:
                detail_logger.error(
                    f"Zieldatei konnte NICHT aus Backup wiederhergestellt "
                    f"werden: {e_restore}\n  Backup liegt unter: {bak_long}")
        raise
    if bak_long and os.path.exists(bak_long):
        if not _safe_remove_with_retry(bak_long):
            detail_logger.warning(f"Backup-Datei nicht löschbar: {bak_long}")


def _cleanup_ppt_inetcache() -> None:
    """Raeumt den PowerPoint/Office-INetCache-Ordner auf
    (%LOCALAPPDATA%\\Microsoft\\Windows\\INetCache\\Content.MSO).

    Hinweis zum Pfadnamen: PowerPoint hat KEINEN eigenen Content.PowerPoint-
    Ordner. PowerPoint, Excel und andere Office-Anwendungen teilen sich den
    gemeinsamen MSO-Cache. Nur Word (Content.Word) und Outlook
    (Content.Outlook) haben eigene Cache-Ordner.

    Best effort: gelockte Dateien werden still uebersprungen (DEBUG).

    DARF NUR AUFGERUFEN WERDEN, WENN DAS SKRIPT-POWERPOINT NICHT LAEUFT.
    """
    local_appdata = os.environ.get("LOCALAPPDATA")
    if not local_appdata:
        return

    cache_dir = os.path.join(
        local_appdata, "Microsoft", "Windows", "INetCache", "Content.MSO")

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
                        detail_logger.debug(f"_cleanup_ppt_inetcache: Exception verworfen: {_e!r}")
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


# Whitelist fuer den Windows-Temp-Cleanup: NUR Eintraege mit diesen
# Praefixen (Office-Reste sowie eigene Skript-Artefakte) duerfen
# geloescht werden. NIEMALS pauschal leeren: Fremdprozesse legen aktive
# Daten ohne Lock in %TEMP% ab; blindes Loeschen zerstoert sie.
# gen_py (pywin32-COM-Cache) und excel8.0 gehoeren bewusst NICHT hierher
# (angeglichen an 4a): gen_py wird von parallel laufenden COM-Skripten
# (3a/3b/3c/4a/4b) aktiv genutzt - rmtree mitten im Lauf zerstoert deren
# Typbibliotheks-Cache; excel8.0 ist ein Excel-Artefakt.
_WINDOWS_TEMP_WHITELIST_PREFIXES = (
    "~$", "~df", "vbe",
    "4c_ersetze_font_in_powerpoint",
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
    # Gesperrte/in-Nutzung-Dateien werden stillschweigend uebersprungen.
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
# PowerPoint-Instanz-Lifecycle
# ==================================================================

def _start_ppt_instance():
    """Startet die Skript-Instanz und prueft, ob sie WIRKLICH die eigene ist.

    PowerPoint ist Einzelinstanz: laeuft (noch oder schon wieder) eine
    Anwendersitzung, liefert DispatchEx genau diese (vom Auftraggeber
    gemessen). Deshalb Momentaufnahme aller POWERPNT-PIDs unmittelbar
    VOR DispatchEx; lief die PID aus Application.HWND schon vorher, ist es
    nicht die eigene Instanz. Ebenso bei unbekannter PID oder schon offenen
    Praesentationen (eine frische Instanz hat keine). In diesen Faellen
    wird an der Instanz NICHTS verstellt (kein AutomationSecurity, kein
    Minimieren, kein DisplayAlerts), sie wird weder per Quit() noch per
    Kill beendet - nur die COM-Referenz wird freigegeben und der Lauf
    bricht mit FremdePowerPointSitzung ab.

    Rueckgabe: (ppt, pid, orig_settings) - orig_settings ist ein dict mit
    den Ausgangswerten von AutomationSecurity und DisplayAlerts fuer
    _quit_ppt_instance."""
    vorher = _alle_powerpoint_pids()
    ppt = win32com.client.DispatchEx("PowerPoint.Application")

    pid = _get_ppt_pid(ppt)
    grund = None
    if pid is None:
        # HWND nicht lesbar: nur eine EINDEUTIG neue PID gilt als eigene.
        neu = _alle_powerpoint_pids() - vorher
        if len(neu) == 1:
            pid = next(iter(neu))
            detail_logger.info(f"PowerPoint-PID via Momentaufnahme: {pid}")
        else:
            grund = (f"PID nicht ermittelbar (HWND fehlgeschlagen, "
                     f"{len(neu)} neue POWERPNT-Prozesse)")
    if grund is None and pid in vorher:
        grund = (f"PID {pid} lief bereits vor dem Start - das ist eine "
                 f"bestehende PowerPoint-Sitzung, nicht die eigene Instanz")
    if grund is None:
        try:
            offen = int(ppt.Presentations.Count)
        except Exception as _e:
            offen = 0
            detail_logger.debug(f"_start_ppt_instance: Exception verworfen: {_e!r}")
        if offen > 0:
            grund = (f"PowerPoint (PID {pid}) hat bereits {offen} offene "
                     f"Praesentation(en) - nicht die eigene, frische Instanz")
    if grund is not None:
        # Nichts verstellen, nicht beenden - nur loslassen.
        ppt = None
        detail_logger.error(f"Fremde PowerPoint-Sitzung erkannt: {grund}")
        raise FremdePowerPointSitzung(
            f"{grund}. Das Skript arbeitet nicht in fremden PowerPoint-"
            f"Sitzungen - bitte PowerPoint vollstaendig schliessen und neu starten.")

    orig_settings = {"AutomationSecurity": None, "DisplayAlerts": None}
    try:
        orig_settings["AutomationSecurity"] = ppt.AutomationSecurity
        ppt.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
    except Exception:
        detail_logger.warning("AutomationSecurity konnte nicht gesetzt werden")
    try:
        orig_settings["DisplayAlerts"] = ppt.DisplayAlerts
    except Exception as _e:
        detail_logger.debug(f"_start_ppt_instance: Exception verworfen: {_e!r}")

    ppt.Visible = COM_TRUE

    try:
        ppt.WindowState = PP_WINDOW_MINIMIZED
    except Exception as _e:
        detail_logger.debug(f"_start_ppt_instance: Exception verworfen: {_e!r}")
    try:
        ppt.DisplayAlerts = PP_ALERTS_NONE
    except Exception as _e:
        detail_logger.debug(f"_start_ppt_instance: Exception verworfen: {_e!r}")
    try:
        ppt.FeatureInstall = MSO_FEATURE_INSTALL_NONE
    except Exception as _e:
        detail_logger.debug(f"_start_ppt_instance: Exception verworfen: {_e!r}")

    try:
        detail_logger.info(
            f"PowerPoint gestartet (PID: {pid}, Version: {ppt.Version})")
    except Exception:
        detail_logger.debug(f"PowerPoint gestartet (PID: {pid})")
    return ppt, pid, orig_settings


def _offene_praesentationen(ppt) -> Optional[int]:
    """Presentations.Count + ProtectedViewWindows.Count; None = nicht lesbar."""
    try:
        n = int(_com_call_with_retry(lambda: ppt.Presentations.Count, retries=3))
    except Exception as _e:
        detail_logger.debug(f"_offene_praesentationen: Exception verworfen: {_e!r}")
        return None
    try:
        n += int(ppt.ProtectedViewWindows.Count)
    except Exception as _e:
        detail_logger.debug(f"_offene_praesentationen: Exception verworfen: {_e!r}")
    return n


def _quit_ppt_instance(ppt, ppt_pid: Optional[int], orig_security=None,
                       create_time: Optional[float] = None) -> bool:
    """Beendet die Skript-Instanz - aber NUR, wenn danach nichts verloren geht.

    Einzelinstanz: oeffnet der Anwender waehrend des Laufs eine Datei
    (Doppelklick), landet sie in DIESER Instanz. Bisher beendeten der
    Neustart alle 25 Dateien und das Laufende sie dann per Quit() (bei
    DisplayAlerts=Keine ohne Rueckfrage) und Kill. Jetzt: Quit/Kill nur,
    wenn nach dem Schliessen der eigenen Praesentation Presentations.Count
    == 0 ist. Sonst werden AutomationSecurity/DisplayAlerts zurueckgesetzt,
    das Fenster wiederhergestellt und gewarnt; die Instanz bleibt offen.
    Ist die Zahl nicht lesbar, bleibt eine noch lebende Instanz ebenfalls
    offen (ein offener Dialog des Anwenders blockiert COM genauso) - nur
    eine nicht mehr lebende braucht kein Beenden.

    Rueckgabe: True = beendet bzw. nicht mehr vorhanden,
               False = bewusst NICHT beendet (Instanz weiter benutzbar)."""
    if ppt is not None:
        einstellungen = orig_security if isinstance(orig_security, dict) else {
            "AutomationSecurity": orig_security, "DisplayAlerts": None}
        offen = _offene_praesentationen(ppt)
        lebt = True
        if offen is None and ppt_pid is not None and create_time is not None:
            try:
                p = psutil.Process(ppt_pid)
                lebt = (p.is_running() and p.create_time() == create_time)
            except psutil.NoSuchProcess:
                lebt = False
            except Exception as _e:
                detail_logger.debug(f"_quit_ppt_instance: Exception verworfen: {_e!r}")
        for attr in ("AutomationSecurity", "DisplayAlerts"):
            if einstellungen.get(attr) is not None:
                try:
                    setattr(ppt, attr, einstellungen[attr])
                except Exception as _e:
                    detail_logger.debug(f"_quit_ppt_instance: Exception verworfen: {_e!r}")
        if (offen is None and lebt) or (offen is not None and offen > 0):
            try:
                ppt.WindowState = PP_WINDOW_NORMAL
            except Exception as _e:
                detail_logger.debug(f"_quit_ppt_instance: Exception verworfen: {_e!r}")
            text = (f"PowerPoint (PID {ppt_pid}) wird NICHT beendet: "
                    + (f"{offen} Praesentation(en) sind noch offen (vermutlich "
                       f"vom Anwender waehrend des Laufs geoeffnet)."
                       if offen else
                       "Anzahl offener Praesentationen nicht lesbar."))
            detail_logger.warning(text)
            try:
                tqdm.write(f"  ⚠  {text}")
            except Exception as _e:
                detail_logger.debug(f"_quit_ppt_instance: Exception verworfen: {_e!r}")
            return False
        try:
            ppt.Quit()
            time.sleep(1)
        except Exception as _e:
            detail_logger.debug(f"_quit_ppt_instance: Exception verworfen: {_e!r}")
    _kill_powerpoint_by_pid(ppt_pid, create_time)
    return True


# ==================================================================
# Watchdog (Timeout-Schutz fuer COM-Aufrufe)
# ==================================================================

def _ppt_call_with_watchdog(call_label: str, timeout: float, ppt_pid,
                             call_fn, ppt_create_time=None):
    """
    Watchdog-Wrapper fuer COM-Calls (analog zu 3a/3b/3c/4a/4b).
    Fuehrt call_fn() aus und beendet den PowerPoint-Prozess (per
    psutil.kill), falls der Aufruf nicht innerhalb von 'timeout'
    Sekunden zurueckkehrt.

    Hintergrund:
      Trust-Center-Probleme oder Geschuetzte-Ansicht-Dialoge in einer
      PowerPoint-Instanz koennen den Aufrufer ohne diesen Schutz
      unbegrenzt blockieren. Mit Watchdog liefert der Aufruf nach
      'timeout' Sekunden eine TimeoutError zurueck, der Prozess wird
      hart beendet und der Smoke-Test/Hauptlauf kann sauber reagieren.

    call_label: Klartext fuer Fehlermeldungen.
    ppt_pid:    PID der PowerPoint-Instanz (von win32process.GetWindowThreadProcessId).
    call_fn:    Lambda mit dem eigentlichen COM-Aufruf.
    """
    done_event   = threading.Event()
    timeout_flag = [False]
    # Schloss + Fertig-Kennzeichen wie in 4a/3c: der Rueckgabewert von
    # call_fn() steht fest, BEVOR das finally done_event setzt. Laeuft der
    # Timeout genau dazwischen ab, toetete der Waechter PowerPoint, obwohl
    # der Aufruf gelungen war - der Aufrufer arbeitete dann mit einer toten
    # Instanz weiter. Nachgemessen mit Attrappe (Aufruf dauert genau so
    # lange wie der Timeout, 300 Laeufe): vorher 27 bzw. 31 Faelle "Erfolg
    # trotz Kill", nachher 0 - der Grenzfall wird jetzt eindeutig Timeout.
    state_lock   = threading.Lock()
    completed    = [False]

    def _watchdog():
        if not done_event.wait(timeout):
            with state_lock:
                if completed[0]:
                    return          # Aufruf war bereits fertig
                timeout_flag[0] = True
            if ppt_pid:
                # _kill_powerpoint_by_pid prueft Process-Name (POWERPNT)
                # UND Erstellungszeit vor kill() und schuetzt damit vor
                # PID-Recycling: Wenn PowerPoint waehrend des Aufrufs
                # abstuerzt (AV-Kill, interner Crash ohne COM-Exception)
                # gibt Windows die PID sofort frei und kann sie sehr
                # schnell an einen anderen Prozess vergeben - auch an
                # eine NEUE PowerPoint-Sitzung des Benutzers, die der
                # Namens-Check allein nicht erkennen wuerde.
                _kill_powerpoint_by_pid(ppt_pid, ppt_create_time)

    wd_thread = threading.Thread(target=_watchdog, daemon=True)
    wd_thread.start()

    try:
        _ergebnis = call_fn()
        with state_lock:
            if timeout_flag[0]:
                # Der Waechter hat bereits zugeschlagen: die Instanz ist
                # tot, das Ergebnis damit unbrauchbar. Als Timeout melden,
                # statt dem Aufrufer einen Erfolg vorzuspiegeln.
                raise TimeoutError(f"PowerPoint Timeout bei {call_label}")
            completed[0] = True
        return _ergebnis
    except Exception as e:
        done_event.set()
        if timeout_flag[0]:
            raise TimeoutError(f"PowerPoint Timeout bei {call_label}")
        raise e
    finally:
        done_event.set()
        wd_thread.join(timeout=1.0)


# ==================================================================
# PowerPoint Smoke-Test (deckt COM-Init und Trust-Center mit ab)
# ==================================================================
# Hintergrund:
#   Vor dem Hauptlauf wird mit einer kurzlebigen PowerPoint-Instanz
#   geprueft, ob der Workflow (Start -> SaveAs im Temp -> Open) ohne
#   Haenger durchlaeuft. Damit werden zwei Fehlerklassen frueh sichtbar:
#
#   1) COM-Subsystem (Kategorie "com"):
#      DispatchEx scheitert mit HRESULT 0x800401F0 (CO_E_NOTINITIALIZED)
#      oder verwandten Fehlern - typisch bei pywin32-Registrierungs-
#      problemen, fehlendem pywin32_postinstall oder AV-Eingriffen.
#
#   2) PowerPoint/Trust-Center (Kategorie "ppt"):
#      PowerPoint laeuft beim Open einer .pptx in einen unsichtbaren
#      Geschuetzte-Ansicht-Dialog, wenn der Pfad nicht als
#      vertrauenswuerdiger Speicherort eingetragen ist. In COM ist
#      dieser Dialog nicht sichtbar - der Aufruf haengt einfach.
#      Pro Datei waeren das minutenlange Haenger. Der Watchdog beendet
#      PowerPoint nach 25 s hart und liefert einen klaren Fehlertext.
#
#   Die Test-.pptx wird via Presentations.Add() + SaveAs() erzeugt
#   (OOXML fuer Praesentationen von Hand zu bauen ist nicht praktikabel
#   - SlideMaster, Layouts, Theme zu umfangreich); danach wird sie
#   wieder geoeffnet (das ist der eigentliche Trust-Center-Test).

def test_trust_center_smoke(timeout: float = 25.0) -> Tuple[bool, str, str]:
    """
    PowerPoint-Smoke-Test fuer 4c:
      Erzeugt eine kurzlebige PowerPoint-Instanz, baut eine Test-.pptx
      im TEMP_PROCESS_PATH, oeffnet sie und schliesst sie wieder. Im
      Erfolgsfall wird PowerPoint sauber beendet; bei Timeout per
      Watchdog hart gekillt.
      Rueckgabe: (ok, fehlertext, kategorie)
        kategorie: "ok"  - Test erfolgreich
                   "com" - COM-Subsystem-Problem (z.B. CO_E_NOTINITIALIZED)
                   "ppt" - PowerPoint/Trust-Center-Problem
                   "fremd" - DispatchEx lieferte eine bestehende Sitzung
                             (siehe _start_ppt_instance), nichts verstellt
    """
    try:
        os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)
    except Exception as e:
        return False, f"Temp-Ordner nicht erstellbar ({TEMP_PROCESS_PATH}): {e}", "ppt"

    test_path = os.path.join(
        TEMP_PROCESS_PATH, f"trustcheck_{uuid.uuid4().hex}.pptx"
    )

    ppt             = None
    ppt_pid         = None
    ppt_create_time = None
    orig_security   = None
    open_pres       = None
    pythoncom.CoInitialize()  # COM-Init fuer Main-Thread (sonst CO_E_NOTINITIALIZED bei DispatchEx)
    try:
        # Eigene kurzlebige PowerPoint-Instanz fuer den Test.
        try:
            ppt, ppt_pid, orig_security = _start_ppt_instance()
        except FremdePowerPointSitzung as e:
            # Kein "Trotzdem fortfahren" moeglich - siehe Hauptprogramm.
            return False, str(e), "fremd"
        except pywintypes.com_error as e:
            # HRESULT 0x800401F0 (-2147221008) = CO_E_NOTINITIALIZED -> COM-Subsystem, kein Trust-Center
            hresult = e.args[0] if e.args else None
            if hresult == -2147221008:
                return False, (f"COM-Subsystem nicht initialisiert "
                               f"(HRESULT 0x800401F0 CO_E_NOTINITIALIZED): {e}"), "com"
            return False, f"PowerPoint-Instanz konnte nicht gestartet werden: {e}", "ppt"
        except Exception as e:
            return False, f"PowerPoint-Instanz konnte nicht gestartet werden: {e}", "ppt"

        ppt_create_time = _get_proc_create_time(ppt_pid)

        # --- (1) Add + SaveAs (mit Watchdog) ---
        try:
            def _do_add_saveas():
                pres = ppt.Presentations.Add(WithWindow=COM_FALSE)
                # Auch bei scheiterndem SaveAs schliessen: _quit_ppt_instance
                # beendet nur bei Presentations.Count == 0 - eine liegen-
                # gebliebene Testpraesentation hielte die Instanz sonst offen.
                try:
                    pres.SaveAs(test_path, FileFormat=PP_FORMAT_PPTX)
                finally:
                    try:
                        pres.Close()
                    except Exception as _e:
                        detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
                return None

            _ppt_call_with_watchdog(
                "Presentations.Add+SaveAs (Smoke-Test)",
                timeout, ppt_pid, _do_add_saveas, ppt_create_time
            )
        except TimeoutError:
            ppt = None  # Watchdog hat den Prozess beendet
            return False, (f"TIMEOUT nach {timeout:.0f}s bei SaveAs "
                           "- Temp-Ordner blockiert oder PowerPoint haengt"), "ppt"
        except Exception as e:
            return False, f"Presentations.Add/SaveAs fehlgeschlagen: {e}", "ppt"

        if not os.path.exists(test_path):
            return False, "SaveAs hat keine Datei erzeugt", "ppt"

        # --- (2) Open der gespeicherten Datei (mit Watchdog) ---
        # Der eigentliche Trust-Center-Test: laeuft das Open ohne Geschuetzte-
        # Ansicht-Dialog durch? Ohne Watchdog haengt das hier minutenlang.
        open_holder = [None]
        def _do_open():
            open_holder[0] = ppt.Presentations.Open(
                test_path,
                ReadOnly=COM_TRUE,
                Untitled=COM_FALSE,
                WithWindow=COM_FALSE,
            )
            return None

        try:
            _ppt_call_with_watchdog(
                "Presentations.Open (Smoke-Test)",
                timeout, ppt_pid, _do_open, ppt_create_time
            )
        except TimeoutError:
            ppt = None  # Watchdog hat den Prozess beendet
            return False, (f"TIMEOUT nach {timeout:.0f}s bei Open "
                           "- Trust Center vermutlich nicht konfiguriert "
                           "oder Geschuetzte Ansicht aktiv"), "ppt"
        except Exception as e:
            return False, f"Presentations.Open fehlgeschlagen: {e}", "ppt"

        open_pres = open_holder[0]
        if open_pres is None:
            return False, "Presentations.Open lieferte kein Praesentations-Objekt", "ppt"

        return True, "", "ok"
    finally:
        if open_pres is not None:
            try: open_pres.Close()
            except Exception as _e:
                detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        if ppt is not None:
            _quit_ppt_instance(ppt, ppt_pid, orig_security, ppt_create_time)
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
# Resume-/Skip-Liste
# ==================================================================

def load_done_set() -> set:
    done = set()
    if os.path.exists(DONE_FILE):
        try:
            with open(DONE_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    p = line.strip()
                    if p:
                        done.add(p)
        except Exception as e:
            detail_logger.warning(f"Done-Liste nicht lesbar: {e}")
    return done



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

def append_done(file_path: str) -> None:
    _protokoll(file_path, "Schrift ersetzt", "OK")
    try:
        with open(DONE_FILE, "a", encoding="utf-8") as f:
            f.write(file_path + "\n")
    except Exception as e:
        detail_logger.warning(f"Done-Liste nicht schreibbar: {e}")


def reset_done_file() -> None:
    try:
        if os.path.exists(DONE_FILE):
            os.remove(DONE_FILE)
    except Exception as e:
        detail_logger.warning(f"Done-Liste nicht löschbar: {e}")


# ==================================================================
# Interaktiver Start
# ==================================================================

def ask_directory() -> str:
    desktop_path   = resolve_desktop_folder()
    downloads_path = resolve_downloads_folder()

    print("\nZielverzeichnis auswählen:")
    print("  [1] Q:\\")
    print("  [2] R:\\")
    print("  [3] G:\\Geteilte Ablagen")
    print("  [4] G:\\Meine Ablage")
    print("  [5] \\\\server\\dfs")
    print(f"  [6] {desktop_path}  (Desktop)")
    print(f"  [7] {downloads_path}  (Downloads)")
    print("  [8] Eigenen Pfad eingeben")
    print("  [0] Abbrechen")
    print()
    while True:
        try:
            choice = input("Auswahl [0-8]: ").strip()
        except (EOFError, RuntimeError, OSError):
            print("\nKeine Eingabe möglich (EOF) – Abbruch.")
            sys.exit(1)
        if choice == "0":
            print("Abgebrochen.")
            sys.exit(0)
        elif choice == "1":
            path = "Q:\\"
        elif choice == "2":
            path = "R:\\"
        elif choice == "3":
            path = "G:\\Geteilte Ablagen"
        elif choice == "4":
            path = "G:\\Meine Ablage"
        elif choice == "5":
            path = "\\\\server\\dfs"
        elif choice == "6":
            path = desktop_path
        elif choice == "7":
            path = downloads_path
        elif choice == "8":
            try:
                raw = input("Pfad eingeben: ")
            except (EOFError, RuntimeError, OSError):
                print("\nKeine Eingabe möglich (EOF) – Abbruch.")
                sys.exit(1)
            path = sanitize_path(raw)
        else:
            print("  Bitte 0-8 eingeben.")
            continue

        if os.path.isdir(path):
            return path

        if choice in ("1", "2"):
            print(f"  Laufwerk {path} nicht erreichbar – ist es gemappt?")
        elif choice in ("3", "4"):
            print(f"  Pfad nicht erreichbar: {path}")
            print("     (Google Drive gestartet? Laufwerk G: gemappt?)")
        elif choice == "5":
            print(f"  Netzwerkpfad nicht erreichbar: {path}")
            print("     (VPN aktiv? Netzwerkverbindung prüfen)")
        elif choice in ("6", "7"):
            print(f"  Pfad nicht gefunden: {path}")
        else:
            print(f"  Verzeichnis nicht gefunden: '{path}'")


def ask_font(default: str) -> str:
    try:
        val = input(f"\nZiel-Schriftart [Standard: {default}]: ").strip()
    except (EOFError, RuntimeError, OSError):
        return default
    return val if val else default


def _is_tty() -> bool:
    """Haengt stdin an einer echten Konsole? (wie in 4b)"""
    try:
        return sys.stdin.isatty()
    except Exception:
        return False


def _warte_auf_taste(auto_mode: bool = False) -> None:
    """Abschliessendes 'Beliebige Taste' - nur wenn wirklich jemand zusieht.

    Ohne diesen Schutz blockierte ein geplanter Lauf mit angehaengter Konsole
    unbegrenzt an der Eingabe: der Task lief nie zu Ende und hielt die
    Einzelinstanz-Sperre, so dass der Folgelauf am naechsten Tag mit
    'bereits aktiv' abbrach - die Migration stand still, ohne dass ein Fehler
    im Log erschien. Ist stdin ganz abgeloest (pythonw, Taskplaner ohne
    Benutzeranmeldung), wirft input() ausserdem RuntimeError('lost sys.stdin')
    oder OSError - beides KEIN EOFError, der frueher allein abgefangen wurde.
    """
    if auto_mode or not _is_tty():
        return
    try:
        input("Beliebige Taste drücken, um das Fenster zu schließen ...")
    except (EOFError, RuntimeError, OSError):
        pass


def ask_yes_no(prompt: str, default_yes: bool = False) -> bool:
    hint = "[J/n]" if default_yes else "[j/N]"
    while True:
        try:
            answer = input(f"{prompt} {hint}: ").strip().lower()
        except (EOFError, RuntimeError, OSError):
            return default_yes
        if answer == "" and default_yes:
            return True
        if answer == "" and not default_yes:
            return False
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


def ask_metadata_detail() -> dict:
    print()
    print("  Folgende Metadaten-Kategorien werden entfernt:")
    safe_keys = ["props", "custom"]
    for key in safe_keys:
        print(f"    * {METADATA_OPTIONS[key]}")

    selected = {k: True for k in safe_keys}

    print()
    print("  Folgende Option entfernt INHALTE (nicht nur Metadaten):")
    selected["comments"] = ask_yes_no(
        f"  {METADATA_OPTIONS['comments']} entfernen?")

    return selected


# ==================================================================
# Datei-Generator (netzwerktauglich)
# ==================================================================


def _ppt_kennwort_grund(path: str) -> str:
    """'' wenn PowerPoint die Datei ohne Kennwortabfrage oeffnet, sonst den Grund.

    Die COM-Schnittstelle kennt keinen Kennwort-Parameter: gemessen am
    29.09.2026 mit PowerPoint 2024 hing eine Praesentation mit
    Aenderungskennwort bis zum Waechter (180 s plus Neustart).
    - Verschluesselt: OOXML-Datei im CFB-Container statt ZIP; .ppt ueber
      msoffcrypto (falls installiert).
    - Aenderungskennwort: <p:modifyVerifier> in ppt/presentation.xml.
    """
    import re
    import zipfile
    ext = os.path.splitext(path)[1].lower()
    try:
        lp = _long_path(path)
        with open(lp, "rb") as f:
            kopf = f.read(8)
        ist_cfb = kopf.startswith(b"\xD0\xCF\x11\xE0")
        if ext in (".ppt", ".pps", ".pot"):
            if not ist_cfb:
                return ""
            try:
                import msoffcrypto
            except ImportError:
                return ""
            with open(lp, "rb") as f:
                return "verschluesselt" if msoffcrypto.OfficeFile(f).is_encrypted() else ""
        if ist_cfb:
            return "verschluesselt"
        if not kopf.startswith(b"PK"):
            return ""
        with zipfile.ZipFile(lp) as z:
            try:
                xml = z.read("ppt/presentation.xml").decode("utf-8", errors="ignore")
            except KeyError:
                return ""
        if re.search(r"<(?:\w+:)?modifyVerifier\b", xml):
            return "schreibkennwort"
    except Exception as _e:
        detail_logger.debug(f"_ppt_kennwort_grund: {_e!r}")
    return ""


def dry_run_directory(directory: str) -> dict:
    """Listet auf, was der Echtlauf tun WUERDE - insbesondere welche
    .ppt/.pps/.pot konvertiert (und deren Originale ersetzt) wuerden.
    Es wird nichts geoeffnet, gespeichert oder geloescht; PowerPoint wird
    nicht gestartet. Identisch aufgebaut zu 4a/4b.

    PowerPoint-Besonderheit: Die COM-Schnittstelle kennt keinen
    Password-Parameter; Dateien mit Kennwort erkennt _ppt_kennwort_grund
    vorab (seit 29.09.2026), der Echtlauf ueberspringt sie."""
    stats = {"VERARBEITEN": 0, "KONVERTIEREN": 0, "GESPERRT": 0,
             "KENNWORT": 0}
    print("\nPROBELAUF – es wird nichts geändert.\n")
    print("-" * 66)
    for file_path in file_generator(directory):
        ext = os.path.splitext(file_path)[1].lower()
        if _is_locked_by_other_user(file_path):
            stats["GESPERRT"] += 1
            tag = "GESPERRT     "
        elif _ppt_kennwort_grund(file_path):
            stats["KENNWORT"] += 1
            tag = "KENNWORT     "
        elif ext in (".ppt", ".pps", ".pot"):
            stats["KONVERTIEREN"] += 1
            tag = "KONVERTIEREN "
        else:
            stats["VERARBEITEN"] += 1
            tag = "VERARBEITEN  "
        print(f"  [{tag}] {file_path}")
        detail_logger.info(f"Probelauf [{tag.strip()}]: {file_path}")
    print("-" * 66)
    print("\n  PROBELAUF-ERGEBNIS (es wurde nichts geändert):")
    print(f"    Würden verarbeitet (Font/Schema):        {stats['VERARBEITEN']}")
    print( "    Würden konvertiert (.ppt/.pps/.pot → neu,")
    print(f"                        Original ersetzt!):  {stats['KONVERTIEREN']}")
    print(f"    Gesperrt (würden übersprungen):          {stats['GESPERRT']}")
    print(f"    Kennwortgeschützt (würden übersprungen): {stats['KENNWORT']}")
    print(f"    GESAMT:                                  {sum(stats.values())}")
    return stats


def file_generator(directory: str, skip_paths: Optional[set] = None):
    extensions = {
        ".pptx", ".pptm", ".ppsx", ".ppsm", ".potx", ".potm",
        ".ppt",  ".pps",  ".pot",
    }
    skip_paths = skip_paths or set()
    # Angeglichen an 4a (file_generator dort): scandir mit Langpfad-Praefix,
    # sonst scheiterten Ordner jenseits von 260 Zeichen still und ihre
    # Dateien fehlten im Lauf. Uebersprungene Ordner/Eintraege werden
    # protokolliert. Junctions/Mount-Points werden nicht betreten -
    # is_symlink() erkennt unter Windows nur echte Symlinks; nachgemessen
    # mit einer Junction auf den eigenen Elternordner: die alte Fassung
    # lief in die Schleife (Dateien mehrfach), die neue meldet jede Datei
    # genau einmal. Ausgegeben (und mit skip_paths verglichen) wird der
    # Pfad OHNE Praefix, zeichengleich zu bisher (os.path.join wie
    # DirEntry.path), damit die Done-Liste weiter greift.
    queue = [directory]
    t_nc = os.path.normcase(os.path.abspath(TEMP_BASE_PATH))
    while queue:
        current_dir = queue.pop()
        entries = None
        for scandir_attempt in range(2):
            try:
                with os.scandir(_long_path(current_dir)) as it:
                    entries = list(it)
                break
            except (PermissionError, OSError) as e_dir:
                if scandir_attempt == 0:
                    time.sleep(1)
                    continue
                detail_logger.warning(
                    f"Verzeichnis übersprungen ({type(e_dir).__name__}): "
                    f"{current_dir} – {e_dir}")
        if entries is None:
            continue
        for entry in entries:
            try:
                clean_path = os.path.join(current_dir, entry.name)
                if entry.is_symlink():
                    continue
                if entry.is_dir(follow_symlinks=False):
                    if entry.name.lower() in EXCLUDE_DIR_NAMES:
                        continue
                    try:
                        attrs = entry.stat(follow_symlinks=False).st_file_attributes
                        if attrs & FILE_ATTRIBUTE_REPARSE_POINT:
                            detail_logger.info(
                                f"Junction/Bereitstellungspunkt nicht betreten: {clean_path}")
                            continue
                    except OSError:
                        pass
                    # Eigenen Arbeitsordner nicht mitverarbeiten.
                    d_nc = os.path.normcase(os.path.abspath(clean_path))
                    if d_nc == t_nc or d_nc.startswith(t_nc + os.sep):
                        continue
                    queue.append(clean_path)
                elif entry.is_file(follow_symlinks=False):
                    nl = entry.name.lower()
                    if nl.startswith("~$") or nl.startswith("._"):
                        continue
                    if any(nl.endswith(ext) for ext in extensions):
                        if clean_path in skip_paths:
                            continue
                        yield clean_path
            except (PermissionError, OSError) as e_entry:
                detail_logger.warning(
                    f"Eintrag übersprungen ({type(e_entry).__name__}): "
                    f"{entry.path} – {e_entry}")
                continue


# ==================================================================
# Shape-Schriftart setzen (rekursiv, mit Tiefenlimit)
# ==================================================================

# Bis 29.09.2026 setzte 4c die Zielschrift pauschal auf jeden Textrahmen -
# auch ueber Text in Wingdings/Symbol; aus einem Haekchen wurde ein "ü".
# Jetzt werden Textrahmen abschnittsweise (Runs) bearbeitet und Abschnitte
# in Symbolschrift bleiben stehen.

def _com_eigenschaft(obj, name: str, *args):
    """Eigenschaft MIT Argumenten (Runs, Characters) in spaeter Bindung.

    Gemessen am 29.09.2026 (Excel, gleiche Office-Typbibliothek):
    TextRange2.Runs(1) wirft "Auflistung nicht unterstuetzt"; die Get-Form
    reicht die Argumente richtig durch.
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


def _textrange_setzen(tr, font_name: str) -> None:
    """PowerPoint-TextRange: Abschnitte in Symbolschrift bleiben stehen."""
    try:
        name = tr.Font.Name
    except Exception:
        name = ""
    if name:
        if not _ist_symbolschrift(name):
            tr.Font.Name = font_name
        return
    if tr.Length == 0:
        tr.Font.Name = font_name
        return
    for run in tr.Runs():
        try:
            if not _ist_symbolschrift(run.Font.Name):
                run.Font.Name = font_name
        except Exception as _e:
            detail_logger.debug(f"_textrange_setzen: Exception verworfen: {_e!r}")


def _textrange2_setzen(tr, font_name: str) -> None:
    """TextRange2 (SmartArt): wie _textrange_setzen, Runs ueber GetRuns."""
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


def _set_chart_fonts(chart, font_name: str) -> None:
    try:
        if chart.HasTitle:
            _font_name_setzen(chart.ChartTitle.Characters().Font, font_name)
    except Exception as _e:
        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")

    for axis_type in (1, 2, 3):
        for axis_group in (1, 2):
            try:
                ax = chart.Axes(axis_type, axis_group)
                try:
                    if ax.HasTitle:
                        _font_name_setzen(ax.AxisTitle.Characters().Font, font_name)
                except Exception as _e:
                    detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
                try:
                    _font_name_setzen(ax.TickLabels.Font, font_name)
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
        for series in chart.SeriesCollection():
            try:
                if series.HasDataLabels:
                    _font_name_setzen(series.DataLabels().Font, font_name)
            except Exception as _e:
                detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
            try:
                for trendline in series.Trendlines():
                    try:
                        if trendline.HasLabel:
                            _font_name_setzen(trendline.DataLabel.Font, font_name)
                    except Exception as _e:
                        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
            except Exception as _e:
                detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")

    try:
        _font_name_setzen(chart.ChartArea.Font, font_name)
    except Exception as _e:
        detail_logger.debug(f"_set_chart_fonts: Exception verworfen: {_e!r}")


def _set_wordart_font(shape, font_name: str) -> bool:
    # Nur klassisches WordArt. TextEffect gibt es bei JEDER Form mit Text;
    # TextEffect.FontName stellte bis 29.09.2026 den ganzen Rahmen auf einmal
    # um - samt Abschnitten in Symbolschrift, die _textrange_setzen danach
    # nicht mehr retten konnte. Textrahmen bearbeitet _process_shape_font.
    try:
        if shape.Type != MSO_TYPE_TEXT_EFFECT:
            return False
    except Exception:
        return False
    try:
        tef = shape.TextEffect
        if tef is not None and not _ist_symbolschrift(tef.FontName):
            tef.FontName = font_name
            detail_logger.debug("WordArt-Font gesetzt")
            return True
    except Exception as _e:
        detail_logger.debug(f"_set_wordart_font: Exception verworfen: {_e!r}")
    return False


def _process_shape_font(shape, font_name: str, depth: int = 0) -> None:
    if depth > MAX_GROUP_DEPTH:
        detail_logger.warning(
            f"Gruppen-Rekursionstiefe > {MAX_GROUP_DEPTH} – abgebrochen")
        return

    try:
        # --- WordArt ---
        try:
            _set_wordart_font(shape, font_name)
        except Exception as _e:
            detail_logger.debug(f"_process_shape_font: Exception verworfen: {_e!r}")

        # --- Textrahmen ---
        if shape.HasTextFrame:
            try:
                _textrange_setzen(shape.TextFrame.TextRange, font_name)
            except Exception as _e:
                detail_logger.debug(f"_process_shape_font: Exception verworfen: {_e!r}")

        # --- Tabelle ---
        if shape.HasTable:
            try:
                tbl = shape.Table
                for row in range(1, tbl.Rows.Count + 1):
                    for col in range(1, tbl.Columns.Count + 1):
                        try:
                            cell = tbl.Cell(row, col)
                            if cell.Shape.HasTextFrame:
                                try:
                                    _textrange_setzen(cell.Shape.TextFrame.TextRange, font_name)
                                except Exception as _e:
                                    detail_logger.debug(f"_process_shape_font: Exception verworfen: {_e!r}")
                        except Exception as _e:
                            detail_logger.debug(f"_process_shape_font: Exception verworfen: {_e!r}")
            except Exception as _e:
                detail_logger.debug(f"_process_shape_font: Exception verworfen: {_e!r}")

        # --- Gruppe (rekursiv) ---
        if shape.Type == MSO_TYPE_GROUP:
            try:
                for sub_shape in shape.GroupItems:
                    _process_shape_font(sub_shape, font_name, depth + 1)
            except Exception as _e:
                detail_logger.debug(f"_process_shape_font: Exception verworfen: {_e!r}")

        # --- SmartArt ---
        try:
            sa = shape.SmartArt
            if sa is not None:
                smartart_failed_nodes = 0
                smartart_total_nodes  = 0

                def _process_smartart_node(node):
                    nonlocal smartart_failed_nodes, smartart_total_nodes
                    smartart_total_nodes += 1
                    font_set = False
                    try:
                        tf2 = node.TextFrame2
                        if tf2 is not None:
                            _textrange2_setzen(tf2.TextRange, font_name)
                            font_set = True
                    except Exception as _e:
                        detail_logger.debug(f"_process_smartart_node: Exception verworfen: {_e!r}")
                    if not font_set:
                        try:
                            tf = node.TextFrame
                            if tf is not None:
                                _textrange_setzen(tf.TextRange, font_name)
                                font_set = True
                        except Exception as _e:
                            detail_logger.debug(f"_process_smartart_node: Exception verworfen: {_e!r}")
                    if not font_set:
                        smartart_failed_nodes += 1

                for node in sa.AllNodes:
                    _process_smartart_node(node)
                if smartart_failed_nodes > 0:
                    detail_logger.warning(
                        f"SmartArt: {smartart_failed_nodes}/{smartart_total_nodes} "
                        f"Knoten nicht änderbar – ggf. manuell prüfen")
        except Exception as _e:
            detail_logger.debug(f"_process_shape_font: Exception verworfen: {_e!r}")

        # --- Chart ---
        if shape.HasChart:
            try:
                _set_chart_fonts(shape.Chart, font_name)
            except Exception as _e:
                detail_logger.debug(f"_process_shape_font: Exception verworfen: {_e!r}")

    except Exception as _e:
        detail_logger.debug(f"_process_shape_font: Exception verworfen: {_e!r}")


def _process_shapes_collection(shapes, font_name: str) -> None:
    try:
        count = shapes.Count
    except Exception:
        try:
            for shape in shapes:
                _process_shape_font(shape, font_name)
        except Exception as _e:
            detail_logger.debug(f"_process_shapes_collection: Exception verworfen: {_e!r}")
        return

    for i in range(1, count + 1):
        try:
            shape = shapes(i)
            _process_shape_font(shape, font_name)
        except Exception:
            continue


# ==================================================================
# Metadaten-Entfernung (PowerPoint-spezifisch)
# ==================================================================

def _remove_ppt_metadata(presentation, selected: dict) -> bool:
    any_removed = False

    if selected.get("props"):
        for prop_name, prop_label in BUILTIN_PROPS_TO_CLEAR.items():
            try:
                presentation.BuiltInDocumentProperties(prop_name).Value = ""
                detail_logger.debug(f"Eigenschaft geleert: {prop_label}")
                any_removed = True
            except Exception as e:
                detail_logger.debug(f"Eigenschaft '{prop_label}' nicht änderbar: {e}")

    if selected.get("custom"):
        try:
            props = presentation.CustomDocumentProperties
            for i in range(props.Count, 0, -1):
                try:
                    props(i).Delete()
                    any_removed = True
                except Exception as e:
                    detail_logger.debug(f"Custom-Prop {i} nicht löschbar: {e}")
            detail_logger.debug("Benutzerdefinierte Eigenschaften geleert")
        except Exception as e:
            detail_logger.debug(f"CustomDocumentProperties nicht verfügbar: {e}")

    if selected.get("comments"):
        try:
            slide_count = presentation.Slides.Count
            for i in range(1, slide_count + 1):
                try:
                    slide = presentation.Slides(i)
                    for j in range(slide.Comments.Count, 0, -1):
                        try:
                            slide.Comments(j).Delete()
                            any_removed = True
                        except Exception as _e:
                            detail_logger.debug(f"_remove_ppt_metadata: Exception verworfen: {_e!r}")
                except Exception as e:
                    detail_logger.debug(
                        f"Kommentare auf Folie {i} nicht löschbar: {e}")
            detail_logger.debug("Folien-Kommentare gelöscht")
        except Exception as e:
            detail_logger.debug(f"Comments-Collection nicht verfügbar: {e}")

    return any_removed


# ==================================================================
# Font-Ersetzung + Metadaten (COM)
# ==================================================================

def replace_fonts_in_presentation(
    file_path: str,
    ppt_app: win32com.client.CDispatch,
    pbar: tqdm,
    font_name: str,
    metadata_selected: dict,
) -> str:
    # ppt_app_global wird hier ggf. auf None gesetzt, wenn der Watchdog
    # die PowerPoint-Instanz killt. Aufrufer prueft danach und startet
    # PowerPoint via _start_ppt_instance neu.
    global ppt_app_global, ppt_pid_global, _eigene_praesentation_global

    pbar.write(f"Prüfe: {os.path.basename(file_path)}")
    detail_logger.info(f"=== Starte: {file_path} ===")

    original_path  = file_path
    is_temp_copy   = False
    was_converted  = False
    presentation   = None
    temp_conv_path = None
    # Final-Markierung ("Als abgeschlossen kennzeichnen"): wird zum
    # Bearbeiten aufgehoben und VOR dem letzten Speichern wiederhergestellt.
    final_aufgehoben = False
    # temp_stage2_path haelt den Pfad zur Stage-2-SaveAs-Datei, in die
    # die finalen Font-Ersetzungen geschrieben werden (statt direkt ins
    # Original). Erst nach Close + erfolgreichem _safe_move wird das
    # Original ueberschrieben. Schuetzt vor Datenverlust bei Crash
    # (Watchdog-Kill nach 240s, Netzwerk-Drop, COM-Fehler) waehrend Save.
    temp_stage2_path = None
    # Zieldatei einer Formatkonvertierung, die DIREKT in die Ablage geschrieben
    # wurde (kein Long-Path-Temp-Umweg). Bricht die Verarbeitung danach ab -
    # Watchdog-Timeout beim Save, harter Kill, COM-Fehler -, liegt dort eine
    # halbfertige Datei NEBEN dem unveraenderten Original. Kein Aufraeumpfad
    # erfasste sie: der finally-Block kennt nur Temp-Artefakte. Deshalb hier
    # mitfuehren und im finally entfernen, solange der Durchlauf nicht
    # erfolgreich abgeschlossen ist.
    direkt_geschriebenes_ziel = None
    lauf_erfolgreich          = False
    # was_read_only protokolliert, ob das Original schreibgeschuetzt war.
    # Wird am Ende nach erfolgreichem Move wiederhergestellt.
    was_read_only  = False
    ext            = os.path.splitext(file_path)[1].lower()

    if _is_locked_by_other_user(original_path):
        pbar.write(f"  ->  ÜBERSPRUNGEN (in Bearbeitung): "
                   f"{os.path.basename(original_path)}")
        detail_logger.info(
            f"Übersprungen (Sperrdatei anderes Nutzers): {original_path}")
        return "SKIPPED"

    _grund = _ppt_kennwort_grund(original_path)
    if _grund:
        _text = "verschlüsselt" if _grund == "verschluesselt" else "Änderungskennwort"
        pbar.write(f"  ->  ÜBERSPRUNGEN ({_text}): "
                   f"{os.path.basename(original_path)}")
        detail_logger.info(f"Übersprungen ({_text}): {original_path}")
        return "SKIPPED"

    orig_atime, orig_mtime, orig_ctime = get_file_timestamps(original_path)
    detail_logger.debug(
        f"Zeitstempel gesichert: atime={orig_atime}, mtime={orig_mtime}, ctime={orig_ctime}")

    # NTFS-Sicherheitsinfo (Owner/Group/DACL) des Originals sichern –
    # wird nach der Datei-Ersetzung wieder angewendet (Admin-Kontext:
    # vollständig inkl. Owner; Nutzer-Kontext: DACL).
    orig_sd = _get_security_descriptor(original_path)

    # --- Long-Path-Behandlung ---
    if len(file_path) > MAX_PATH_LEN:
        if _check_file_locked(file_path):
            pbar.write(f"  ->  ÜBERSPRUNGEN (gesperrt): "
                       f"{os.path.basename(original_path)}")
            detail_logger.info(f"Übersprungen (Dateisperre): {original_path}")
            return "SKIPPED"

        is_temp_copy = True
        detail_logger.debug("Long-Path erkannt, erstelle Temp-Kopie")
        try:
            os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)
            uid         = uuid.uuid4().hex
            unique_name = f"{uid}_{os.path.basename(file_path)}"
            temp_path   = os.path.join(TEMP_PROCESS_PATH, unique_name)

            if len(temp_path) >= WINDOWS_MAX_PATH:
                unique_name = f"{uid}{os.path.splitext(file_path)[1].lower()}"
                temp_path   = os.path.join(TEMP_PROCESS_PATH, unique_name)
                detail_logger.debug(f"Temp-Pfad zu lang, kürze: {temp_path}")

            shutil.copy(_long_path(file_path), temp_path)
            file_path = temp_path
            detail_logger.debug(f"Temp-Kopie: {temp_path}")

            # --- AV-Scanner-Fenster abwarten ---
            if not _wait_until_unlocked(temp_path):
                detail_logger.warning(
                    f"Temp-Kopie bleibt gesperrt (AV-Scanner?): {temp_path}")
                pbar.write(f"  ->  ÜBERSPRUNGEN (AV-Sperre): "
                           f"{os.path.basename(original_path)}")
                return "SKIPPED"
        except Exception as e:
            log_error(original_path, Exception(f"Long-Path-Kopie fehlgeschlagen: {e}"))
            return "ERROR"

    try:
        if _check_file_locked(file_path):
            if not _wait_until_unlocked(file_path):
                pbar.write(f"  ->  ÜBERSPRUNGEN (gesperrt): "
                           f"{os.path.basename(original_path)}")
                detail_logger.info(f"Übersprungen (Dateisperre): {original_path}")
                return "SKIPPED"

        # Watchdog-geschuetzter Open: _com_call_with_retry schuetzt nur vor
        # zurueckgeworfenen COM-Fehlern (HRESULT-Retries), aber nicht vor
        # haengenden Aufrufen. Bei unsichtbaren Modal-Dialogen (Reparatur-
        # Prompt, Verknuepfungs-Update, DRM-Schutz) wuerde der Open ohne
        # Watchdog unendlich blockieren. Der Watchdog killt nach
        # PPT_OPEN_TIMEOUT die PowerPoint-Instanz; die nachgelagerte
        # if-ppt-is-None-Recovery im Aufrufer startet neu.
        def _do_open():
            return _com_call_with_retry(
                ppt_app.Presentations.Open,
                file_path,
                COM_FALSE,
                COM_FALSE,
                COM_FALSE,
                retries=5,
            )
        try:
            presentation = _ppt_call_with_watchdog(
                "Presentations.Open", PPT_OPEN_TIMEOUT, ppt_pid_global, _do_open,
                ppt_create_time_global
            )
        except TimeoutError:
            # PowerPoint-Instanz ist nach Watchdog-Kill tot. Globalstate
            # leeren, damit der Aufrufer beim naechsten File via
            # "if ppt is None"-Recovery _start_ppt_instance() neu aufruft.
            ppt_app_global = None
            ppt_pid_global = None
            pbar.write(f"  ✕  TIMEOUT bei Open ({PPT_OPEN_TIMEOUT:.0f}s) – PowerPoint blockiert")
            log_error(original_path, Exception(
                f"Timeout bei Presentations.Open nach {PPT_OPEN_TIMEOUT:.0f}s "
                f"(vermutlich Modal-Dialog: Reparatur/Verknuepfung/DRM)"
            ))
            return "ERROR"
        detail_logger.debug("Präsentation geöffnet")
        _eigene_praesentation_global = presentation

        try:
            if presentation.Final:
                presentation.Final = False
                final_aufgehoben = True
                detail_logger.debug("Final-Markierung aufgehoben")
        except Exception as _e:
            detail_logger.debug(f"replace_fonts_in_presentation: Exception verworfen: {_e!r}")

        try:
            presentation.RemovePersonalInformation = False
            detail_logger.debug("RemovePersonalInformation deaktiviert")
        except Exception as _e:
            detail_logger.debug(f"replace_fonts_in_presentation: Exception verworfen: {_e!r}")

        # --- Schema-Migration ---
        new_path = None

        if ext in (".ppt", ".pps", ".pot"):
            try:
                # Im Zweifel Makroformat (wie 4a): Scheiterte HasVBProject
                # (Trust Center: kein VBA-Projektzugriff), galt die Datei
                # bisher als makrofrei - eine .ppt MIT Makros wurde dann als
                # .pptx gespeichert und die Makros waren still verloren.
                has_macros = False
                try:
                    has_macros = bool(presentation.HasVBProject)
                except Exception as e_vba:
                    try:
                        comps      = presentation.VBProject.VBComponents
                        has_macros = any(
                            comps.Item(i).CodeModule.CountOfLines > 0
                            for i in range(1, comps.Count + 1)
                        )
                        detail_logger.debug(
                            "HasVBProject fehlgeschlagen – VBProject-Fallback erfolgreich")
                    except Exception:
                        has_macros = True
                        detail_logger.warning(
                            f"HasVBProject-Prüfung und VBProject-Fallback fehlgeschlagen "
                            f"(Trust-Center: Makroeinstellungen + VBA-Projektzugriff "
                            f"prüfen): {e_vba}\n"
                            f"  Datei wird im Makro-Format gespeichert (konservativ).")
                        pbar.write(
                            f"  ⚠  VBA-Prüfung nicht möglich → Makro-Format: "
                            f"{os.path.basename(original_path)}")

                if ext == ".pps":
                    new_format = PP_FORMAT_PPSM if has_macros else PP_FORMAT_PPSX
                    new_ext    = ".ppsm" if has_macros else ".ppsx"
                elif ext == ".pot":
                    new_format = PP_FORMAT_POTM if has_macros else PP_FORMAT_POTX
                    new_ext    = ".potm" if has_macros else ".potx"
                else:
                    new_format = PP_FORMAT_PPTM if has_macros else PP_FORMAT_PPTX
                    new_ext    = ".pptm" if has_macros else ".pptx"

                new_path = os.path.splitext(original_path)[0] + new_ext

                # Namenskollisions-Schutz: existiert das Ziel bereits
                # (fremde Datei gleichen Namens), wird ein nummerierter
                # Ausweichname gewaehlt, statt die fremde Datei via
                # SaveAs stillschweigend zu ueberschreiben (analog 4b).
                if os.path.exists(_long_path(new_path)):
                    base_no_ext = os.path.splitext(original_path)[0]
                    counter = 1
                    while os.path.exists(_long_path(
                            f"{base_no_ext}_{counter}{new_ext}")) and counter < 100:
                        counter += 1
                    new_path = f"{base_no_ext}_{counter}{new_ext}"
                    if os.path.exists(_long_path(new_path)):
                        raise RuntimeError(
                            "Mehr als 100 Namenskollisionen! Abbruch zum "
                            "Schutz fremder Dateien.")
                    pbar.write(
                        f"  ⚠  Ziel existiert bereits – speichere als "
                        f"'{os.path.basename(new_path)}'")
                    detail_logger.info(
                        f"Namenskollision gelöst: {original_path} → {new_path}")

                # Watchdog-geschuetzter SaveAs. TimeoutError wird ausserhalb
                # des try/except Exception-Blocks behandelt, damit eine tote
                # PowerPoint-Instanz nicht still als "Konvertierung
                # fehlgeschlagen" durchrutscht.
                if is_temp_copy:
                    temp_conv_path = os.path.join(
                        TEMP_PROCESS_PATH, uuid.uuid4().hex + new_ext)
                    _saveas_target = temp_conv_path
                else:
                    _saveas_target = new_path
                    # Ohne Temp-Umweg schreibt SaveAs direkt in die Ablage -
                    # ab hier ist eine unfertige Zieldatei moeglich.
                    direkt_geschriebenes_ziel = new_path

                try:
                    _ppt_call_with_watchdog(
                        "Presentations.SaveAs (Schema)",
                        PPT_SAVE_TIMEOUT, ppt_pid_global,
                        lambda: presentation.SaveAs(_saveas_target, new_format),
                        ppt_create_time_global
                    )
                except TimeoutError:
                    ppt_app_global = None
                    ppt_pid_global = None
                    pbar.write(f"  ✕  TIMEOUT bei Schema-SaveAs ({PPT_SAVE_TIMEOUT:.0f}s)")
                    log_error(original_path, Exception(
                        f"Timeout bei SaveAs (Schema-Migration) nach {PPT_SAVE_TIMEOUT:.0f}s"
                    ))
                    presentation = None
                    return "ERROR"

                was_converted = True
                detail_logger.debug(f"Schema-Migration: {ext} -> {new_ext}")

            except Exception as e:
                # Gescheiterte Konvertierung ist ein FEHLER. Bisher nur eine
                # Warnung: die Praesentation blieb an die Original-.ppt
                # gebunden, der Font-Save schrieb IN DAS ORIGINAL, es wurde
                # SUCCESS gezaehlt und in die Done-Liste eingetragen (kein
                # Wiederholungslauf), und ein Torso new_path blieb neben dem
                # Original liegen (lauf_erfolgreich=True verhinderte das
                # Aufraeumen). Jetzt: ERROR, nichts speichern - das finally
                # schliesst ohne Speichern und entfernt den Torso
                # (direkt_geschriebenes_ziel) bzw. die Temp-Konvertierung.
                log_error(original_path, Exception(
                    f"Konvertierung {ext} fehlgeschlagen: {e}"))
                pbar.write(f"  !!  FEHLER (Konvertierung): "
                           f"{os.path.basename(original_path)} – Original unverändert")
                return "ERROR"

        else:
            target_format = OPENXML_FORMAT_MAP.get(ext)
            if target_format is not None:
                # DATENVERLUST-SCHUTZ:
                #  - is_temp_copy=True  : SaveAs auf file_path schreibt in
                #                         die Long-Path-Tempkopie. Original
                #                         (auf der Original-Freigabe) bleibt
                #                         unberuehrt.
                #  - is_temp_copy=False : file_path == ORIGINAL. Ein nacktes
                #                         SaveAs(file_path,...) wuerde das
                #                         Original direkt ueberschreiben.
                #                         Bei Crash mittendrin (Watchdog-
                #                         Kill nach 240s, Netzwerk-Drop,
                #                         COM-Fehler) waere das Original
                #                         korrupt. Daher: SaveAs in
                #                         temp_stage2_path, PowerPoint
                #                         arbeitet ab dann auf der Temp-
                #                         Datei, finaler Move ins Original
                #                         erfolgt nach Close.
                if is_temp_copy:
                    _reserialize_target = file_path
                else:
                    os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)
                    temp_stage2_path = os.path.join(
                        TEMP_PROCESS_PATH, f"stage2_{uuid.uuid4().hex}{ext}")
                    _reserialize_target = temp_stage2_path

                try:
                    _ppt_call_with_watchdog(
                        "Presentations.SaveAs (Re-Serialisierung)",
                        PPT_SAVE_TIMEOUT, ppt_pid_global,
                        lambda: presentation.SaveAs(_reserialize_target, target_format),
                        ppt_create_time_global
                    )
                    if temp_stage2_path is not None:
                        detail_logger.debug(
                            f"Re-Serialisierung -> Stage-2-Temp: {temp_stage2_path}")
                    else:
                        detail_logger.debug(f"Re-Serialisierung: {ext}")
                except TimeoutError:
                    ppt_app_global = None
                    ppt_pid_global = None
                    pbar.write(f"  ✕  TIMEOUT bei Re-Serialisierung ({PPT_SAVE_TIMEOUT:.0f}s)")
                    log_error(original_path, Exception(
                        f"Timeout bei SaveAs (Re-Serialisierung) nach {PPT_SAVE_TIMEOUT:.0f}s"
                    ))
                    presentation = None
                    return "ERROR"
                except Exception as e:
                    if temp_stage2_path is None:
                        # Long-Path-Kopie: der Font-Save trifft nur die
                        # Temp-Kopie, das Original ersetzt spaeter
                        # _replace_file_with_backup - unveraendert.
                        detail_logger.warning(f"Re-Serialisierung fehlgeschlagen: {e}")
                    else:
                        # Scheitert der Stage-2-SaveAs, ist die Praesentation
                        # noch an das ORIGINAL auf der Ablage gebunden. Bisher
                        # wurde temp_stage2_path=None gesetzt und weiter-
                        # gemacht: der Font-Save schrieb dann ungeschuetzt
                        # direkt ins Original und zaehlte SUCCESS - genau das,
                        # was Stage 2 verhindern soll. Jetzt: ERROR, Original
                        # unberuehrt. temp_stage2_path bleibt gesetzt, damit
                        # das finally einen Rest entfernt.
                        log_error(original_path, Exception(
                            f"Re-Serialisierung (Stage-2-SaveAs) fehlgeschlagen: {e}"))
                        pbar.write(f"  !!  FEHLER (Re-Serialisierung): "
                                   f"{os.path.basename(original_path)} – Original unverändert")
                        return "ERROR"
            else:
                # Kein OOXML-Format-Mapping (selten - alte Formate ohne
                # Open-XML-Pendant). Trotzdem temp_stage2 vorbereiten,
                # damit der finale Save in Temp landet statt im Original.
                if not is_temp_copy:
                    os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)
                    temp_stage2_path = os.path.join(
                        TEMP_PROCESS_PATH, f"stage2_{uuid.uuid4().hex}{ext}")
                    try:
                        _ppt_call_with_watchdog(
                            "Presentations.SaveAs (Stage-2)",
                            PPT_SAVE_TIMEOUT, ppt_pid_global,
                            lambda: presentation.SaveAs(temp_stage2_path),
                            ppt_create_time_global
                        )
                        detail_logger.debug(
                            f"Stage-2-SaveAs (kein OOXML-Mapping) -> {temp_stage2_path}")
                    except TimeoutError:
                        ppt_app_global = None
                        ppt_pid_global = None
                        pbar.write(
                            f"  ✕  TIMEOUT bei Stage-2-SaveAs ({PPT_SAVE_TIMEOUT:.0f}s)")
                        log_error(original_path, Exception(
                            f"Timeout bei Stage-2-SaveAs nach {PPT_SAVE_TIMEOUT:.0f}s"
                        ))
                        presentation = None
                        return "ERROR"
                    except Exception as e:
                        # Wie oben: ohne Stage-2-Datei wuerde der Font-Save
                        # direkt ins Original schreiben -> ERROR, Original
                        # unberuehrt; ein Rest in temp_stage2_path wird im
                        # finally entfernt.
                        log_error(original_path, Exception(
                            f"Stage-2-SaveAs fehlgeschlagen: {e}"))
                        pbar.write(f"  !!  FEHLER (Stage-2-SaveAs): "
                                   f"{os.path.basename(original_path)} – Original unverändert")
                        return "ERROR"

        # --- Folienmaster + Layouts ---
        try:
            masters_count = presentation.SlideMasters.Count
            for i in range(1, masters_count + 1):
                try:
                    master = presentation.SlideMasters(i)
                    _process_shapes_collection(master.Shapes, font_name)
                    try:
                        layouts_count = master.CustomLayouts.Count
                        for j in range(1, layouts_count + 1):
                            try:
                                layout = master.CustomLayouts(j)
                                _process_shapes_collection(layout.Shapes, font_name)
                            except Exception:
                                continue
                    except Exception as _e:
                        detail_logger.debug(f"replace_fonts_in_presentation: Exception verworfen: {_e!r}")
                except Exception:
                    continue
        except Exception as e:
            detail_logger.debug(f"SlideMasters: {e}")

        # --- Handzettel-Master ---
        try:
            _process_shapes_collection(
                presentation.HandoutMaster.Shapes, font_name)
        except Exception as e:
            detail_logger.debug(f"HandoutMaster: {e}")

        # --- Notizen-Master ---
        try:
            _process_shapes_collection(
                presentation.NotesMaster.Shapes, font_name)
        except Exception as e:
            detail_logger.debug(f"NotesMaster: {e}")

        # --- Folien + Notizseiten ---
        try:
            slide_count = presentation.Slides.Count
            for i in range(1, slide_count + 1):
                try:
                    slide = presentation.Slides(i)
                    _process_shapes_collection(slide.Shapes, font_name)
                    try:
                        if slide.HasNotesPage:
                            _process_shapes_collection(
                                slide.NotesPage.Shapes, font_name)
                    except Exception as e:
                        detail_logger.debug(f"NotesPage Folie {i}: {e}")
                except Exception as e:
                    detail_logger.debug(f"Folie {i}: {e}")
                    continue
        except Exception as e:
            detail_logger.debug(f"Slides: {e}")

        detail_logger.debug("Schriftarten ersetzt")

        # --- Speichern (Fonts) ---
        try:
            _ppt_call_with_watchdog(
                "Presentations.Save (Fonts)",
                PPT_SAVE_TIMEOUT, ppt_pid_global,
                lambda: presentation.Save(),
                ppt_create_time_global
            )
        except TimeoutError:
            ppt_app_global = None
            ppt_pid_global = None
            pbar.write(f"  ✕  TIMEOUT bei Save (Fonts) ({PPT_SAVE_TIMEOUT:.0f}s)")
            log_error(original_path, Exception(
                f"Timeout bei Save (Fonts) nach {PPT_SAVE_TIMEOUT:.0f}s"
            ))
            presentation = None
            return "ERROR"
        detail_logger.debug("Gespeichert (Fonts)")

        # --- Metadaten (optional) ---
        if any(metadata_selected.values()):
            try:
                removed = _remove_ppt_metadata(presentation, metadata_selected)
                if removed:
                    try:
                        _ppt_call_with_watchdog(
                            "Presentations.Save (Metadaten)",
                            PPT_SAVE_TIMEOUT, ppt_pid_global,
                            lambda: presentation.Save(),
                            ppt_create_time_global
                        )
                        detail_logger.debug("Gespeichert (Metadaten)")
                    except TimeoutError:
                        ppt_app_global = None
                        ppt_pid_global = None
                        pbar.write(f"  ✕  TIMEOUT bei Save (Metadaten) ({PPT_SAVE_TIMEOUT:.0f}s)")
                        log_error(original_path, Exception(
                            f"Timeout bei Save (Metadaten) nach {PPT_SAVE_TIMEOUT:.0f}s"
                        ))
                        presentation = None
                        return "ERROR"
                    except Exception as e_save:
                        detail_logger.warning(
                            f"Speichern nach Metadaten fehlgeschlagen: {e_save}")
            except Exception as e_meta:
                detail_logger.warning(
                    f"Metadaten-Bereinigung fehlgeschlagen: {e_meta}")

        # --- Final-Markierung wiederherstellen (vor dem letzten Speichern) ---
        # Bisher blieb Final=False gespeichert: die Kennzeichnung "Als
        # abgeschlossen" war nach dem Lauf dauerhaft weg. Eigenes, letztes
        # Save NACH allen Aenderungen, damit Font- und Metadaten-Pass nicht
        # an einer als abgeschlossen markierten Praesentation arbeiten.
        # Scheitert es, wird die Datei NICHT uebernommen (ERROR, Original
        # unveraendert) - lieber gar nicht als halb gespeichert.
        # Ungemessen (kein Office hier): ob PowerPoint das Save nach
        # Final=True annimmt - bitte mit einer markierten Datei pruefen.
        if final_aufgehoben:
            try:
                presentation.Final = True
                _ppt_call_with_watchdog(
                    "Presentations.Save (Final-Markierung)",
                    PPT_SAVE_TIMEOUT, ppt_pid_global,
                    lambda: presentation.Save(),
                    ppt_create_time_global
                )
                final_aufgehoben = False
                detail_logger.debug("Final-Markierung wiederhergestellt und gespeichert")
            except TimeoutError:
                ppt_app_global = None
                ppt_pid_global = None
                pbar.write(f"  ✕  TIMEOUT bei Save (Final) ({PPT_SAVE_TIMEOUT:.0f}s)")
                log_error(original_path, Exception(
                    f"Timeout bei Save (Final-Markierung) nach {PPT_SAVE_TIMEOUT:.0f}s"
                ))
                presentation = None
                return "ERROR"
            except Exception as e_final:
                log_error(original_path, Exception(
                    f"Final-Markierung nicht wiederherstellbar: {e_final}"))
                pbar.write(f"  !!  FEHLER (Final-Markierung): "
                           f"{os.path.basename(original_path)} – Original unverändert")
                return "ERROR"

        # --- Schließen ---
        presentation.Close()
        presentation = None
        detail_logger.debug("Präsentation geschlossen")

        # --- Verschieben + Zeitstempel ---
        if was_converted:
            if is_temp_copy and temp_conv_path:
                _safe_move(temp_conv_path, new_path)
                detail_logger.debug(f"Konvertierte Datei verschoben: {new_path}")

            new_path_long = _long_path(new_path)
            new_path_safe = (
                os.path.exists(new_path_long)
                and os.path.getsize(new_path_long) > 0
            )
            if new_path_safe:
                _apply_security_descriptor(new_path, orig_sd)
                set_file_timestamps(new_path, orig_atime, orig_mtime, orig_ctime)
                if os.path.exists(_long_path(original_path)):
                    if not _safe_remove_with_retry(original_path):
                        detail_logger.warning(
                            f"Alte Datei nicht löschbar (gesperrt): {original_path}")
                    else:
                        detail_logger.debug(f"Alte Datei gelöscht: {original_path}")
            else:
                detail_logger.error(
                    f"Konvertiertes Ziel fehlt oder leer: {new_path} "
                    f"– Original wird NICHT gelöscht.")
                log_error(original_path, Exception(
                    f"SaveAs scheinbar erfolgreich, aber {new_path} "
                    f"fehlt oder hat 0 Bytes – Datenverlust verhindert"))
                pbar.write(f"  !!  FEHLER (Zieldatei leer): "
                           f"{os.path.basename(original_path)}")
                return "ERROR"
        else:
            if is_temp_copy:
                # Move mit Backup-Schutz: shutil.copy2 (Basis von
                # _safe_move) wuerde das Original beim Zurueckschieben
                # sonst sofort auf 0 Bytes trunkieren.
                _replace_file_with_backup(file_path, original_path)
                is_temp_copy = False
            elif temp_stage2_path is not None:
                # Stage-2: PowerPoint hat in temp_stage2_path geschrieben.
                # Original ist noch unberuehrt. Jetzt atomar druebermoven.
                try:
                    # Read-Only kurzzeitig entfernen, sonst schlaegt der
                    # Move auf NTFS fehl. was_read_only fuer Wiederher-
                    # stellung merken.
                    orig_long = _long_path(original_path)
                    if os.path.exists(orig_long):
                        try:
                            attrs = win32api.GetFileAttributes(orig_long)
                            if attrs != -1 and (attrs & win32con.FILE_ATTRIBUTE_READONLY):
                                was_read_only = True
                                os.chmod(orig_long, stat.S_IWRITE)
                        except Exception as _e:
                            detail_logger.debug(f"replace_fonts_in_presentation: Exception verworfen: {_e!r}")

                    if not (os.path.exists(_long_path(temp_stage2_path))
                            and os.path.getsize(_long_path(temp_stage2_path)) > 0):
                        raise RuntimeError(
                            f"Stage-2-Tempdatei ist leer/fehlt: {temp_stage2_path}")
                    # Move mit Backup-Schutz: shutil.copy2 (Basis von
                    # _safe_move) trunkiert das Original sofort auf
                    # 0 Bytes ('wb'). Ohne .bak-Backup waere es bei einem
                    # Abbruch mitten im Kopiervorgang zerstoert.
                    _replace_file_with_backup(temp_stage2_path, original_path)
                    temp_stage2_path = None
                    detail_logger.debug(
                        f"Stage-2-Tempdatei -> Original verschoben: {original_path}")
                except Exception as e_s2_move:
                    # Move fehlgeschlagen - _replace_file_with_backup hat
                    # das Original aus dem .bak wiederhergestellt (bzw.
                    # das Backup liegt neben dem Original, falls auch das
                    # Rollback scheiterte - siehe Detail-Log). Stage-2-
                    # Tempdatei zusaetzlich als RESCUE_-Datei sichern.
                    rescue_name = f"RESCUE_{uuid.uuid4().hex}{ext}"
                    rescue_path = os.path.join(TEMP_PROCESS_PATH, rescue_name)
                    detail_logger.warning(
                        f"Stage-2-Move fehlgeschlagen: {e_s2_move}\n"
                        f"  Versuche Rescue-Kopie nach: {rescue_path}\n"
                        f"  Original wiederhergestellt bzw. Backup vorhanden: "
                        f"{original_path}"
                    )
                    try:
                        if temp_stage2_path and os.path.exists(
                                _long_path(temp_stage2_path)):
                            shutil.copy2(
                                _long_path(temp_stage2_path),
                                _long_path(rescue_path))
                            detail_logger.warning(
                                f"Stage-2-Rescue gesichert: {rescue_path}")
                            pbar.write(
                                f"  ⚠  RESCUE (Font-Replace): "
                                f"{os.path.basename(rescue_path)} "
                                f"(Original geschützt)")
                    except Exception as e_s2_rescue:
                        detail_logger.error(
                            f"Stage-2-Rescue fehlgeschlagen: {e_s2_rescue}")
                    log_error(
                        original_path,
                        Exception(f"Stage-2-Move fehlgeschlagen: {e_s2_move}"))
                    pbar.write(
                        f"  !!  FEHLER (Stage-2-Move): "
                        f"{os.path.basename(original_path)}")
                    return "ERROR"

            # Read-Only-Attribut wiederherstellen, falls vorher gesetzt
            if was_read_only:
                try:
                    orig_long = _long_path(original_path)
                    if os.path.exists(orig_long):
                        current = win32api.GetFileAttributes(orig_long)
                        if current != -1:
                            win32api.SetFileAttributes(
                                orig_long,
                                current | win32con.FILE_ATTRIBUTE_READONLY)
                            detail_logger.debug(
                                f"Schreibschutz wiederhergestellt: {original_path}")
                except Exception as e_ro:
                    detail_logger.debug(
                        f"Schreibschutz-Wiederherstellung fehlgeschlagen: {e_ro}")
            _apply_security_descriptor(original_path, orig_sd)
            set_file_timestamps(original_path, orig_atime, orig_mtime, orig_ctime)

        pbar.write(f"  OK  {os.path.basename(original_path)}"
                   + (" [konvertiert]" if was_converted else ""))
        detail_logger.info(f"Erfolgreich: {original_path}")
        append_done(original_path)
        lauf_erfolgreich = True
        return "SUCCESS"

    except Exception as e:
        err_msg = str(e).lower()

        is_password_error = any(kw in err_msg for kw in (
            "password", "passwort", "kennwort",
            "protected", "geschützt", "geschuetzt",
            "0x800a11fd",
            "0x800a01a8",
            "0x800a11ff",
            "0x80048240",
            "encrypted",
        ))

        if is_password_error:
            pbar.write(f"  ->  ÜBERSPRUNGEN (Passwort): "
                       f"{os.path.basename(original_path)}")
            detail_logger.info(f"Übersprungen (Passwort): {original_path}")
            return "SKIPPED"

        log_error(original_path, e)
        pbar.write(f"  !!  FEHLER: {os.path.basename(original_path)}")
        return "ERROR"

    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception as _e:
                detail_logger.debug(f"replace_fonts_in_presentation: Exception verworfen: {_e!r}")
        _eigene_praesentation_global = None
        if is_temp_copy and os.path.exists(_long_path(file_path)):
            _safe_remove_with_retry(file_path)
        if temp_conv_path and os.path.exists(_long_path(temp_conv_path)):
            _safe_remove_with_retry(temp_conv_path)
        # Stage-2-Tempdatei: nur loeschen wenn sie noch da ist und nicht
        # erfolgreich verschoben wurde. Bei Fehler wurde sie bereits zur
        # RESCUE_-Datei kopiert; das hier ist nur Aufraeumarbeit.
        if temp_stage2_path and os.path.exists(_long_path(temp_stage2_path)):
            _safe_remove_with_retry(temp_stage2_path)
        # Halbfertige Zieldatei einer direkt in die Ablage geschriebenen
        # Konvertierung entfernen. Sonst laegen Original und Teilkonvertat
        # nebeneinander: der Anwender sieht zwei Dateien und kann nicht
        # erkennen, welche brauchbar ist - und ein Wiederholungslauf legt
        # wegen der Namenskollision eine dritte an (_1).
        # Das Original bleibt unangetastet; nur die neue Datei faellt weg.
        if (not lauf_erfolgreich
                and direkt_geschriebenes_ziel
                and direkt_geschriebenes_ziel.lower() != original_path.lower()
                and os.path.exists(_long_path(direkt_geschriebenes_ziel))):
            if _safe_remove_with_retry(direkt_geschriebenes_ziel):
                detail_logger.warning(
                    f"Unfertige Zieldatei nach Abbruch entfernt: "
                    f"{direkt_geschriebenes_ziel}")
            else:
                detail_logger.error(
                    f"Unfertige Zieldatei NICHT entfernbar - bitte pruefen: "
                    f"{direkt_geschriebenes_ziel}")


# ==================================================================
# Verzeichnis-Verarbeitung
# ==================================================================

def process_directory(
    directory: str,
    font_name: str,
    metadata_selected: dict,
    count_files_first: bool,
    skip_paths: Optional[set] = None,
    show_progress: bool = True,
) -> dict:
    os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)
    skip_paths = skip_paths or set()

    if count_files_first:
        print("\nIndiziere Dateien (kann einige Minuten dauern) ...")
        files_list = []
        with tqdm(
            desc="Indiziere",
            unit="Datei",
            bar_format="{desc}: {n_fmt} Dateien [{elapsed}, {rate_fmt}]",
            disable=not show_progress,
        ) as idx_bar:
            for f in file_generator(directory, skip_paths):
                files_list.append(f)
                idx_bar.update(1)
        total    = len(files_list)
        iterable = files_list
        bar_fmt  = "{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]"
        print(f"\n{total} Dateien gefunden. Starte Verarbeitung ...\n")
    else:
        iterable = file_generator(directory, skip_paths)
        total    = None
        bar_fmt  = "{desc}: {n_fmt} Dateien [{elapsed}, {rate_fmt}]"
        print("\nStarte direkte Verarbeitung (ohne Vorab-Zählung) ...\n")

    print("-" * 66)

    stats = {"SUCCESS": 0, "ERROR": 0, "SKIPPED": 0}

    global ppt_app_global, ppt_pid_global, ppt_orig_security_global
    global ppt_create_time_global, _abbruch_fremde_sitzung
    ppt             = None
    ppt_pid         = None
    ppt_create_time = None
    orig_security   = None
    file_counter    = 0

    try:
        ppt, ppt_pid, orig_security = _start_ppt_instance()
        ppt_create_time          = _get_proc_create_time(ppt_pid)
        ppt_app_global           = ppt
        ppt_pid_global           = ppt_pid
        ppt_create_time_global   = ppt_create_time
        ppt_orig_security_global = orig_security

        with tqdm(total=total, desc="Verarbeite", unit="Datei",
                  bar_format=bar_fmt, disable=not show_progress) as pbar:
            for file_path in iterable:
                # Einzelinstanz: oeffnet der Anwender waehrend des Laufs eine
                # Praesentation, landet sie in der Instanz des Skripts (dort
                # sind Warnhinweise abgeschaltet, und der Waechter wuerde sie
                # bei einer Zeitueberschreitung mit beenden). Zwischen zwei
                # Dateien ist keine eigene Praesentation offen - alles, was
                # jetzt offen ist, gehoert dem Anwender. Bis 29.09.2026
                # arbeitete 4c dann weiter (nur der Neustart entfiel); jetzt
                # endet der Lauf wie in 3c, und das finally unten laesst die
                # Instanz offen (_quit_ppt_instance beendet nichts, solange
                # etwas offen ist).
                if ppt is not None:
                    _offen = _offene_praesentationen(ppt)
                    if _offen:
                        meldung = (f"{_offen} Praesentation(en) des Anwenders in der "
                                   f"PowerPoint-Instanz des Skripts geoeffnet - Lauf wird "
                                   f"beendet, PowerPoint bleibt offen.")
                        pbar.write(f"  ⚠  {meldung}")
                        detail_logger.warning(meldung)
                        _abbruch_fremde_sitzung = True
                        break
                try:
                    # --- Geplanter Restart ---
                    # Nur wenn _quit_ppt_instance wirklich beendet hat: sind
                    # in der Instanz noch Praesentationen offen (Einzel-
                    # instanz - vom Anwender waehrend des Laufs geoeffnet),
                    # entfaellt der Neustart und die Instanz arbeitet weiter.
                    if (file_counter > 0 and file_counter % COM_RESTART_INTERVAL == 0
                            and ppt is not None
                            and not _quit_ppt_instance(ppt, ppt_pid, orig_security,
                                                       ppt_create_time)):
                        detail_logger.warning(
                            f"PPT-Neustart nach {file_counter} Dateien entfaellt "
                            f"- offene Praesentationen in der Instanz")
                        # Einstellungen wurden fuer die Uebergabe zurueck-
                        # gesetzt; fuer die Weiterarbeit wieder setzen.
                        for _attr, _wert in (("AutomationSecurity", MSO_AUTOMATION_SECURITY_LOW),
                                             ("DisplayAlerts", PP_ALERTS_NONE)):
                            try:
                                setattr(ppt, _attr, _wert)
                            except Exception as _e:
                                detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                    elif file_counter > 0 and file_counter % COM_RESTART_INTERVAL == 0:
                        detail_logger.info(
                            f"PPT-Neustart nach {file_counter} Dateien")
                        ppt             = None
                        ppt_pid         = None
                        ppt_create_time = None
                        orig_security   = None
                        ppt_app_global           = None
                        ppt_pid_global           = None
                        ppt_create_time_global   = None
                        ppt_orig_security_global = None
                        # Speicherbereinigungs-Neustart: idealer Zeitpunkt fuer
                        # Cache-Cleanup analog 3c (PowerPoint tot, frischer Neustart).
                        try:
                            _cleanup_ppt_inetcache()
                        except Exception as _e:
                            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                        try:
                            _cleanup_user_recent()
                        except Exception as _e:
                            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                        time.sleep(1)

                    # --- Recovery, falls PPT-Instanz fehlt ---
                    if ppt is None:
                        try:
                            ppt, ppt_pid, orig_security = _start_ppt_instance()
                            ppt_create_time          = _get_proc_create_time(ppt_pid)
                            ppt_app_global           = ppt
                            ppt_pid_global           = ppt_pid
                            ppt_create_time_global   = ppt_create_time
                            ppt_orig_security_global = orig_security
                        except FremdePowerPointSitzung:
                            raise
                        except Exception as e_start:
                            log_error("PPT_RESTART", e_start)
                            pbar.write(
                                "  !!  PPT-Neustart fehlgeschlagen, "
                                "warte 5s vor erneutem Versuch")
                            time.sleep(5)
                            try:
                                ppt, ppt_pid, orig_security = _start_ppt_instance()
                                ppt_create_time          = _get_proc_create_time(ppt_pid)
                                ppt_app_global           = ppt
                                ppt_pid_global           = ppt_pid
                                ppt_create_time_global   = ppt_create_time
                                ppt_orig_security_global = orig_security
                            except FremdePowerPointSitzung:
                                raise
                            except Exception as e_retry:
                                log_error("PPT_RESTART_RETRY", e_retry)
                                stats["ERROR"] += 1
                                pbar.write(
                                    f"  !!  KRITISCH: PPT nicht startbar, "
                                    f"überspringe: {os.path.basename(file_path)}")
                                continue

                    result = replace_fonts_in_presentation(
                        file_path, ppt, pbar, font_name, metadata_selected)
                    stats[result] += 1

                    # Reaktiver Restart-Check: replace_fonts_in_presentation
                    # setzt ppt_app_global=None, wenn der Watchdog die
                    # PowerPoint-Instanz killen musste. Ohne diese Spiegelung
                    # auf die lokale ppt-Variable wuerde der naechste
                    # Iterationsstart den toten ppt-Handle weiterverwenden
                    # und alle Folgedateien crashen, bis der periodische
                    # COM_RESTART_INTERVAL greift.
                    if ppt_app_global is None and ppt is not None:
                        detail_logger.info(
                            "Reaktiver PPT-Neustart nach Watchdog-Kill")
                        pbar.write("  ↻  PowerPoint reagiert nicht mehr – starte neu...")
                        # _quit_ppt_instance versucht graceful Quit, dann
                        # Kill via PID. Da Watchdog die Instanz bereits
                        # gekillt hat, ist es im Wesentlichen ein Cleanup.
                        try:
                            _quit_ppt_instance(ppt, ppt_pid, orig_security,
                                               ppt_create_time)
                        except Exception as _e:
                            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                        ppt             = None
                        ppt_pid         = None
                        ppt_create_time = None
                        orig_security   = None
                        # Cache-Cleanup nach Watchdog-Kill (PowerPoint tot).
                        try:
                            _cleanup_ppt_inetcache()
                        except Exception as _e:
                            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                        try:
                            _cleanup_user_recent()
                        except Exception as _e:
                            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                except FremdePowerPointSitzung:
                    raise
                except Exception as e:
                    log_error(file_path, e)
                    stats["ERROR"] += 1
                    pbar.write(f"  !!  KRITISCH: {os.path.basename(file_path)}")
                finally:
                    file_counter += 1
                    pbar.update(1)

    except FremdePowerPointSitzung as e_fremd:
        # Abbruch statt Weiterarbeit in einer fremden Sitzung (die Instanz
        # wurde in _start_ppt_instance weder verstellt noch beendet).
        _abbruch_fremde_sitzung = True
        print(f"\nABBRUCH: {e_fremd}")
        log_error("GLOBAL", e_fremd)

    except Exception as e:
        print(f"\nKRITISCHER FEHLER: PowerPoint konnte nicht gestartet werden: {e}")
        log_error("GLOBAL", e)

    finally:
        if not _quit_ppt_instance(ppt, ppt_pid, orig_security, ppt_create_time):
            print("\n  HINWEIS: PowerPoint wurde NICHT beendet, weil darin noch")
            print("  Praesentationen offen sind (vermutlich waehrend des Laufs")
            print("  geoeffnet). Bitte dort speichern und PowerPoint selbst schliessen.")
        ppt_app_global           = None
        ppt_pid_global           = None
        ppt_create_time_global   = None
        ppt_orig_security_global = None
        # RESCUE_-Dateien duerfen das Aufraeumen ueberleben - ein blindes
        # rmtree wuerde die einzige Kopie geretteter Stage-2-Ergebnisse
        # vernichten.
        _safe_cleanup_temp()
        # Office-INetCache (Content.MSO) und Windows-Recent (.lnk)
        # aufraeumen - analog zu 3a/3b/3c/4a/4b. Beide best effort.
        try:
            _cleanup_ppt_inetcache()
        except Exception as _e:
            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
        try:
            _cleanup_user_recent()
        except Exception as _e:
            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")

    return stats


# ==================================================================
# Einstiegspunkt
# ==================================================================


# ==================================================================
# Kommandozeile (Paritaet zu 4a/4b)
# ==================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="PowerPoint Font-Ersetzung – ersetzt Schriftarten in "
                    "PowerPoint-Dateien",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--auto", action="store_true",
                        help="Trust-Center-Check und interaktive Rückfragen überspringen "
                             "(läuft PowerPoint bereits: Abbruch mit Exitcode 2)")
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
                             "PowerPoint wird nicht gestartet, nichts geöffnet, "
                             "gespeichert oder gelöscht.")
    parser.add_argument("--reset-done", action="store_true",
                        help="Resume-Liste (Done-Datei) vor dem Lauf zurücksetzen")
    return parser.parse_args()


if __name__ == "__main__":
    args      = parse_args()
    auto_mode = args.auto

    check_required_modules()
    check_powerpoint_installed()

    # --- Header-Block ---
    _title_inner    = "POWERPOINT FONT-ERSETZUNG"
    _subtitle_inner = "Stand: 11.06.2026"
    _box_width      = 64

    print()
    print("╔" + "═" * _box_width + "╗")
    print("║" + _title_inner.center(_box_width)    + "║")
    print("║" + _subtitle_inner.center(_box_width) + "║")
    print("╚" + "═" * _box_width + "╝")
    # Laeuft PowerPoint, wird nicht per COM gearbeitet (Einzelinstanz).
    # Frueh pruefen, damit niemand erst alle Fragen beantwortet; der
    # Probelauf braucht kein PowerPoint und wird nicht aufgehalten.
    # Verbindlich ist die zweite Pruefung direkt vor dem Smoke-Test.
    if not args.dry_run:
        pruefe_powerpoint_geschlossen(auto_mode)

    print(f"  Start: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    print("  - Schema-Migration: alle Dateien mit aktueller Engine neu serialisiert")
    print("  - Zeitstempel werden beibehalten (atime, mtime, ctime)")
    print("  - SlideMaster/Layouts/Handout/Notizen werden verarbeitet")
    print("  - WordArt wird unterstützt")
    print("=" * 66)

    # --- Resume-Prüfung ---
    done_set = load_done_set()
    if args.reset_done and done_set:
        done_set = set()
        reset_done_file()
        print("\n  -> Done-Liste zurückgesetzt (--reset-done).")
    elif done_set:
        print()
        print(f"  Gefunden: {len(done_set)} bereits verarbeitete Dateien "
              f"({DONE_FILE})")
        if auto_mode or ask_yes_no(
            "  Diese Dateien bei der erneuten Verarbeitung überspringen?",
            default_yes=True,
        ):
            print(f"  -> {len(done_set)} Dateien werden übersprungen.")
        else:
            done_set = set()
            reset_done_file()
            print("  -> Done-Liste zurückgesetzt.")

    # --- Konfiguration: CLI hat Vorrang, sonst interaktive Abfrage ---
    if args.directory:
        start_dir = sanitize_path(args.directory)
        if not os.path.isdir(_long_path(start_dir)):
            print(f"\n[X] Verzeichnis nicht erreichbar: '{start_dir}'")
            sys.exit(1)
    elif auto_mode:
        print("\n[X] Im Modus --auto ist --dir erforderlich.")
        sys.exit(1)
    else:
        start_dir = ask_directory()

    if auto_mode or args.count_first:
        show_progress = True
        count_first   = args.count_first
    else:
        show_progress, count_first = ask_progress_mode()

    font_name = args.font if args.font else (
        DEFAULT_FONT_NAME if auto_mode else ask_font(DEFAULT_FONT_NAME))

    metadata_selected = {}
    if args.no_meta or auto_mode:
        # Ohne explizite Angabe im Automatikbetrieb bewusst nichts entfernen.
        remove_meta = False
    else:
        remove_meta = ask_yes_no("\nMetadaten entfernen?")
        if remove_meta:
            metadata_selected = ask_metadata_detail()

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

    # --- Zusammenfassung ---
    print()
    print("=" * 66)
    print("  KONFIGURATION")
    print("=" * 66)
    print(f"  Verzeichnis:          {start_dir}")
    print(f"  Neue Schriftart:      {font_name}")
    print(f"  Temp-Arbeitsordner:   {TEMP_PROCESS_PATH}")
    print("  Zeitstempel:          BLEIBEN ERHALTEN")
    print(f"  Modus:                {'Vorab zählen (ETA)' if count_first else 'Generator (kein ETA)'}")
    print(f"  Fortschrittsbalken:   {'an' if show_progress else 'aus (inline)'}")
    print(f"  Resume-Liste:         {len(done_set)} Dateien werden übersprungen"
          if done_set else "  Resume-Liste:         (leer)")
    if metadata_selected and any(metadata_selected.values()):
        print("  Metadaten entfernen:")
        for key, active in metadata_selected.items():
            if active:
                print(f"    * {METADATA_OPTIONS[key]}")
    else:
        print("  Metadaten entfernen:  Nein")
    print("  Verarbeitungsumfang:")
    print("    * Schema-Migration: Re-Serialisierung ALLER Dateien")
    print("    * Konvertierung: .ppt/.pps/.pot -> .pptx/.ppsx/.potx")
    print("    * SlideMasters + Layouts (inkl. H/F-Platzhalter-Shapes)")
    print("    * HandoutMaster + NotesMaster")
    print("    * Alle Folien: Shapes, Tabellen, Gruppen, Notizseiten")
    print("    * SmartArt-Knoten (via AllNodes, rekursiv)")
    print("    * Diagramme (ChartTitle, Achsen, Legende, DataLabels, Trendlinien)")
    print("    * WordArt (via TextEffectFormat)")
    print("=" * 66)

    # --- Trust-Center Vorab-Check ---
    print("\n" + "!" * 66)
    print("  WICHTIGER CHECK: TRUST-CENTER-EINSTELLUNGEN")
    print("!" * 66)
    print("  PowerPoint öffnen → Datei → Optionen → Trust Center")
    print("  → 'Einstellungen für das Trust Center...'")
    print()
    print("  1. GESCHÜTZTE ANSICHT:")
    print("     ☐ Dateien aus dem Internet in geschützter Ansicht öffnen")
    print("     ☐ Dateien an potenziell unsicheren Orten ... öffnen")
    print("     ☐ Anlagen aus Outlook in geschützter Ansicht öffnen")
    print("     → ALLE 3 DEAKTIVIEREN")
    print()
    print("  2. MAKROEINSTELLUNGEN:")
    print("     ⦿ Alle Makros aktivieren")
    print("     ☑ Zugriff auf das VBA-Projektobjektmodell vertrauen")
    print()
    print("  3. VERTRAUENSWÜRDIGE SPEICHERORTE:")
    print("     ☑ Vertrauenswürdige Speicherorte im Netzwerk zulassen")
    print("     Folgende Speicherorte hinzufügen:")
    print(f"     • {TEMP_BASE_PATH} (temporärer Arbeitsordner des Skripts)")
    print(r"     • Q:\ oder \\server\dfs (Quelldateien)")
    print("     Jeweils: ☑ Unterordner ... sind ebenfalls vertrauenswürdig")
    print()
    print("  4. DATENSCHUTZOPTIONEN:")
    print("     ☐ Beim Öffnen automatisch verknüpfte Daten aktualisieren")
    print("     → DEAKTIVIEREN (verhindert Popups bei eingebetteten Objekten)")
    print("!" * 66)
    if not auto_mode and not dry_run and not ask_yes_no(
        "Wurden diese Einstellungen in PowerPoint vorgenommen?",
        default_yes=False,
    ):
        print("\n[X] Abbruch. Bitte konfigurieren Sie erst das Trust-Center in PowerPoint.")
        sys.exit(0)

    if not auto_mode and not ask_yes_no("\nJetzt starten?"):
        print("Abgebrochen.")
        sys.exit(0)

    # --- Probelauf: ohne PowerPoint, ohne Aenderungen, dann Ende ---
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

    # --- PowerPoint Smoke-Test (nach Bestaetigung, vor Hauptlauf) ---
    # Prueft mit einer kurzlebigen PowerPoint-Instanz, ob PowerPoint im
    # TEMP_PROCESS_PATH eine .pptx erzeugen, oeffnen und schliessen kann.
    # Bei Geschuetzte-Ansicht-Hänger killt der Watchdog PowerPoint nach
    # 25 s. Ohne diesen Test wuerde ein falsch konfiguriertes Trust-
    # Center die spaeteren Open-Aufrufe pro Datei minutenlang blockieren.
    # Kategorien: "com" = COM-Subsystem-Problem, "ppt" = PowerPoint/Trust-Center.
    # Unmittelbar vor dem ersten COM-Zugriff erneut: zwischen Start und
    # hier kann der Anwender PowerPoint geoeffnet haben.
    pruefe_powerpoint_geschlossen(auto_mode)
    print("\nPrüfe COM-Subsystem und Trust-Center ...")
    smoke_ok, smoke_msg, smoke_cat = test_trust_center_smoke(timeout=25.0)
    if not smoke_ok and smoke_cat == "fremd":
        # Kein "Trotzdem fortfahren": der Hauptlauf bekaeme dieselbe
        # fremde Sitzung. Sie wurde weder verstellt noch beendet.
        print()
        print("=" * 66)
        print("  ⚠  POWERPOINT-SITZUNG IST NICHT DIE EIGENE - ABBRUCH")
        print("=" * 66)
        print(f"  Grund: {smoke_msg}")
        print("=" * 66)
        file_logger.error(f"Abbruch: fremde PowerPoint-Sitzung: {smoke_msg}")
        sys.exit(2)
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
            print("    • Antivirus / EDR blockiert COM-Aufrufe auf POWERPNT.EXE")
            print("=" * 66)
            file_logger.error(f"COM-Subsystem nicht verfuegbar: {smoke_msg}")
            log_warn_text = "Abbruch durch Benutzer nach COM-Subsystem-Fehler."
            log_ignore_text = "COM-Subsystem-Fehler vom Benutzer ignoriert – Fortsetzung."
        else:
            print("  ⚠  POWERPOINT-SMOKE-TEST FEHLGESCHLAGEN")
            print("=" * 66)
            print(f"  Grund: {smoke_msg}")
            print()
            print("  Mögliche Ursachen:")
            print("    • Temp-Ordner nicht als vertrauenswürdiger Speicherort eingetragen")
            print(f"      ({TEMP_BASE_PATH})")
            print("    • Geschützte Ansicht für unsichere Speicherorte noch aktiv")
            print("    • PowerPoint/Office-Profil beschädigt oder fehlende Desktop-Ordner")
            print("=" * 66)
            file_logger.error(f"PowerPoint-Smoke-Test fehlgeschlagen: {smoke_msg}")
            log_warn_text = "Abbruch durch Benutzer nach Smoke-Test-Warnung."
            log_ignore_text = "Smoke-Test-Warnung vom Benutzer ignoriert – Fortsetzung."

        # Im Automatikmodus NICHT fragen. Der Hilfetext zu --auto sagt
        # 'Trust-Center-Check und interaktive Rueckfragen ueberspringen', und
        # alle anderen Rueckfragen sind entsprechend geklammert - dieser Block
        # war es nicht. Mit angehaengter Konsole blockierte input() unbegrenzt
        # (der Task lief nie zu Ende und hielt die Einzelinstanz-Sperre); ohne
        # stdin lieferte ask_yes_no den Vorgabewert False und das Skript
        # beendete sich mit sys.exit(0), also Erfolgs-Exitcode - der
        # Aufgabenplaner meldete 'erfolgreich', obwohl keine einzige Datei
        # verarbeitet wurde. Exit-Code 2 macht den Fehlschlag sichtbar.
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
        detail_logger.info("PowerPoint-Smoke-Test erfolgreich.")
        print("  → Smoke-Test OK.")

    # --- Verarbeitung ---
    signal.signal(signal.SIGINT, _signal_handler)
    if hasattr(signal, "SIGTERM"):
        try:
            signal.signal(signal.SIGTERM, _signal_handler)
        except (ValueError, OSError):
            pass

    # --- Einzelinstanz-Schutz ---
    # Diese Sperre gab es bisher nur in 3a-3c. Ohne sie konnten zwei
    # Laeufe gleichzeitig ueber denselben Bestand gehen und sich
    # gegenseitig die Temp-Kopien und Zieldateien wegziehen.
    # Freigabe ueber atexit, damit sie auch bei sys.exit greift.
    _sperre = None
    if gem is not None:
        _sperre = gem.Einzelinstanz("4c_ersetze_font_in_powerpoint")
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
            start_dir, font_name, metadata_selected, count_first, done_set,
            show_progress=show_progress)
    except Exception as e:
        print(f"\nKRITISCHER FEHLER: {e}")
        log_error("GLOBAL", e)
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    duration = datetime.now() - start_time

    # --- Abschluss ---
    print()
    print("=" * 66)
    print("  VERARBEITUNG ABGESCHLOSSEN")
    print("=" * 66)
    print(f"  Ende:  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Dauer: {str(duration).split('.')[0]}")
    print()

    if stats_result:
        total_sum = sum(stats_result.values())
        print("  STATISTIK:")
        print(f"    [OK] Erfolgreich:    {stats_result.get('SUCCESS', 0)}")
        print(f"    [->] Übersprungen:   {stats_result.get('SKIPPED', 0)}")
        print(f"    [!!] Fehler:         {stats_result.get('ERROR', 0)}")
        print(f"         GESAMT:          {total_sum}")

    print("\n" + "!" * 66)
    print("WICHTIGER HINWEIS ZUR SICHERHEIT:")
    print("Bitte denken Sie daran, die TRUST-CENTER EINSTELLUNGEN in PowerPoint")
    print("aus Sicherheitsgründen wieder auf Ihre Standardwerte zurückzusetzen,")
    print("insbesondere die 'Geschützte Ansicht' und die 'Makroeinstellungen'.")
    print("\nDer temporäre Ordner kann entfernt werden:")
    print(f"  {TEMP_BASE_PATH}")
    print("!" * 66)

    print()
    print(f"  Fehler-Log:   {os.path.abspath(LOG_FILE)}")
    print(f"  Detail-Log:   {os.path.abspath(DETAILED_LOG_FILE)}")
    print(f"  Done-Liste:   {os.path.abspath(DONE_FILE)}")
    print("=" * 66)

    print()
    # _cleanup_windows_temp() VOR die Wartezeile: schliesst der Anwender das
    # Fenster ueber das Kreuz statt Enter zu druecken, unterblieb die
    # Temp-Bereinigung sonst vollstaendig (~$- und 4c-Reste in
    # %TEMP%). In 4a steht der Aufruf ebenfalls davor.
    _cleanup_windows_temp()
    _warte_auf_taste(auto_mode)
    # Abbruch wegen fremder PowerPoint-Sitzung: fuer den Aufgabenplaner
    # als Fehlschlag sichtbar machen (--auto).
    if _abbruch_fremde_sitzung:
        sys.exit(2)
