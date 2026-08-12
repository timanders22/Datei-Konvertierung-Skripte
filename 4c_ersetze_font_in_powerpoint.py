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
#    • C:\Users\<Benutzername>\Documents (temporärer Arbeitsordner des Skripts)
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
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

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

PP_FORMAT_PPTX = 24
PP_FORMAT_PPTM = 25
PP_FORMAT_POTX = 26
PP_FORMAT_POTM = 27
PP_FORMAT_PPSX = 28
PP_FORMAT_PPSM = 29

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
TEMP_BASE_PATH = os.path.join(
    resolve_documents_folder(), "4c_ersetze_font_in_powerpoint"
)
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


# ==================================================================
# Signal-Handler (Ctrl+C / SIGTERM)
# ==================================================================

def _signal_handler(sig, frame) -> None:
    print("\n\n*** ABBRUCH durch Benutzer (Ctrl+C) - räume auf ...")

    if ppt_app_global is not None:
        if ppt_orig_security_global is not None:
            try:
                ppt_app_global.AutomationSecurity = ppt_orig_security_global
            except Exception:
                pass
        try:
            ppt_app_global.Quit()
            time.sleep(1)
        except Exception:
            pass

    _kill_powerpoint_by_pid(ppt_pid_global, ppt_create_time_global)

    # Best effort Cache-Cleanup beim Strg+C-Abbruch (analog 3c).
    try:
        _cleanup_ppt_inetcache()
    except Exception:
        pass
    try:
        _cleanup_user_recent()
    except Exception:
        pass

    _safe_cleanup_temp()

    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

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


def warn_running_powerpoint(auto_mode: bool = False) -> None:
    """Warnt vor bereits laufenden PowerPoint-Sitzungen.

    PowerPoint wird per COM automatisiert. Laeuft bereits eine Sitzung des
    Anwenders, hat das zwei Auswirkungen, die man vorher kennen sollte:

      1. Die Automatisierung kann sich an die vorhandene Sitzung haengen.
         Warnhinweise werden dann abgeschaltet und das Fenster ausgeblendet
         - fuer den Anwender sieht das aus, als sei PowerPoint abgestuerzt.
      2. Die Aufraeumroutinen des Skripts koennen die eigene Instanz nicht
         mehr sicher von der fremden unterscheiden.

    Die Sitzung wird NICHT beendet (dafuer sorgt die Momentaufnahme der
    fremden Prozess-IDs beim Start), aber ein sauberer Lauf setzt ein
    geschlossenes PowerPoint voraus.
    """
    pids = find_running_powerpoint_pids()
    if not pids:
        return
    pids_str = ", ".join(str(p) for p in pids)

    if auto_mode:
        detail_logger.warning(
            f"Automatikmodus: laufende PowerPoint-Sitzungen (PID: {pids_str}) - "
            f"sie werden geschuetzt, aber nicht geschlossen."
        )
        return

    print()
    print("=" * 66)
    print("  WARNUNG: PowerPoint laeuft bereits")
    print("=" * 66)
    print(f"  Gefundene PowerPoint-Prozesse (PID): {pids_str}")
    print()
    print("  Ihre Sitzung wird vom Skript NICHT beendet. Waehrend des Laufs")
    print("  kann sie aber ausgeblendet werden und Warnhinweise sind")
    print("  abgeschaltet - das wirkt wie ein Absturz.")
    print("  Ausserdem laesst sich die eigene Automatisierungs-Instanz dann")
    print("  nicht mehr zuverlaessig von Ihrer Sitzung unterscheiden.")
    print()
    print("  EMPFEHLUNG: PowerPoint jetzt schliessen und das Skript neu starten.")
    print("=" * 66)
    print()
    if not ask_yes_no("Trotzdem fortfahren?"):
        print("Abgebrochen. Bitte PowerPoint schliessen und neu starten.")
        sys.exit(0)


def _kill_powerpoint_by_pid(pid: Optional[int],
                            create_time: Optional[float] = None) -> None:
    if pid is None:
        return
    try:
        proc = psutil.Process(pid)
        if not proc.is_running() or "POWERPNT" not in (proc.name() or "").upper():
            return
        # Erstellungszeit-Check schuetzt zusaetzlich zum Namens-Check
        # vor PID-Recycling auf eine NEUE PowerPoint-Sitzung des
        # Benutzers, die der Namens-Check allein nicht erkennen wuerde.
        if create_time is not None:
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
            except Exception:
                pass
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


def check_powerpoint_installed() -> None:
    pythoncom.CoInitialize()
    ppt = None
    try:
        try:
            ppt = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                detail_logger.info(f"PowerPoint-Version erkannt: {ppt.Version}")
            except Exception:
                pass
        except pywintypes.com_error as e:
            print("=" * 66)
            print("FEHLER: PowerPoint konnte nicht gestartet werden.")
            print("Bitte stellen Sie sicher, dass Microsoft PowerPoint installiert ist")
            print("(Office 2019 oder 2024, deutsch oder englisch).")
            print(f"COM-Details: {e}")
            print("=" * 66)
            sys.exit(1)
        except Exception as e:
            print("=" * 66)
            print("FEHLER: PowerPoint konnte nicht gestartet werden.")
            print(f"Details: {e}")
            print("=" * 66)
            sys.exit(1)
    finally:
        if ppt is not None:
            try:
                ppt.Quit()
            except Exception:
                pass
            time.sleep(0.5)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


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
    except Exception:
        pass
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
            except Exception:
                pass
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
    if dacl is not None:
        try:
            dacl_flags = win32security.DACL_SECURITY_INFORMATION
            try:
                ctrl, _rev = sd.GetSecurityDescriptorControl()
                if ctrl & win32security.SE_DACL_PROTECTED:
                    dacl_flags |= win32security.PROTECTED_DACL_SECURITY_INFORMATION
                else:
                    dacl_flags |= win32security.UNPROTECTED_DACL_SECURITY_INFORMATION
            except Exception:
                pass
            win32security.SetNamedSecurityInfo(
                p, win32security.SE_FILE_OBJECT, dacl_flags,
                None, None, dacl, None)
            detail_logger.debug(f"DACL wiederhergestellt: {path}")
        except Exception as e:
            detail_logger.warning(
                f"DACL-Wiederherstellung fehlgeschlagen ({path}): {e}")

    try:
        owner = sd.GetSecurityDescriptorOwner()
        group = None
        try:
            group = sd.GetSecurityDescriptorGroup()
        except Exception:
            pass
        if owner is not None:
            sec_flags = win32security.OWNER_SECURITY_INFORMATION
            if group is not None:
                sec_flags |= win32security.GROUP_SECURITY_INFORMATION
            win32security.SetNamedSecurityInfo(
                p, win32security.SE_FILE_OBJECT, sec_flags,
                owner, group, None, None)
            detail_logger.debug(f"Owner wiederhergestellt: {path}")
    except Exception as e:
        if _restore_privileges_enabled:
            detail_logger.warning(
                f"Owner-Wiederherstellung fehlgeschlagen ({path}): {e}")
        else:
            detail_logger.debug(
                f"Owner nicht gesetzt (kein Admin-Privileg – im "
                f"Nutzer-Kontext unkritisch): {path} – {e}")


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
    except Exception:
        # Wenn GetFileAttributes scheitert, dem alten Pfad folgen
        # (Write-Test wird scheitern, aber das ist dann nicht durch R/O)
        pass

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
    except Exception:
        pass
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
        except Exception:
            pass
        bak_long = _long_path(dst + f"_{uuid.uuid4().hex[:6]}.bak")
        os.replace(dst_long, bak_long)
    try:
        _safe_move(src, dst)
    except Exception:
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
                    except Exception:
                        pass
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
    """Raeumt den Windows-Recent-Ordner (%USERPROFILE%\\Recent) auf.

    Hier liegen .lnk-Verknuepfungen aller kuerzlich geoeffneten Dateien
    (nicht nur PowerPoint). Nur .lnk-Dateien werden geloescht.

    BEWUSSTE DESIGN-ENTSCHEIDUNG: Verknuepfungen sind systemweit, nicht
    nur PowerPoint. Akzeptabel im Kontext eines Massen-Verarbeitungs-Skripts.
    """
    # Seit Windows Vista liegt der Recent-Ordner unter
    # %APPDATA%\Microsoft\Windows\Recent. %USERPROFILE%\Recent ist nur
    # noch eine Kompatibilitaets-Junction mit Deny-List-ACL (Auflisten
    # wirft PermissionError) und dient hier lediglich als Fallback.
    candidates = []
    appdata = os.environ.get("APPDATA")
    if appdata:
        candidates.append(
            os.path.join(appdata, "Microsoft", "Windows", "Recent"))
    user_profile = os.environ.get("USERPROFILE")
    if user_profile:
        candidates.append(os.path.join(user_profile, "Recent"))

    recent_dir = None
    for cand in candidates:
        if cand.lower().endswith(os.sep + "recent") and os.path.isdir(cand):
            recent_dir = cand
            break
    if recent_dir is None:
        return

    files_deleted = 0
    files_skipped = 0
    bytes_freed   = 0

    try:
        for entry in os.listdir(recent_dir):
            if not entry.lower().endswith(".lnk"):
                continue
            fpath = os.path.join(recent_dir, entry)
            if not os.path.isfile(fpath):
                continue
            try:
                fsize = os.path.getsize(fpath)
            except OSError:
                fsize = 0
            try:
                try:
                    os.chmod(fpath, stat.S_IWRITE)
                except Exception:
                    pass
                os.remove(fpath)
                files_deleted += 1
                bytes_freed   += fsize
            except (OSError, PermissionError) as e:
                files_skipped += 1
                detail_logger.debug(
                    f"Recent: .lnk gesperrt, uebersprungen: {fpath} ({e})")
    except Exception as e:
        detail_logger.debug(f"Recent-Cleanup unerwarteter Fehler: {e}")
        return

    if files_deleted > 0 or files_skipped > 0:
        kb = bytes_freed / 1024
        detail_logger.info(
            f"Windows-Recent aufgeraeumt: "
            f"{files_deleted} .lnk-Dateien geloescht ({kb:.1f} KB), "
            f"{files_skipped} gesperrt/uebersprungen")


# Whitelist fuer den Windows-Temp-Cleanup: NUR Eintraege mit diesen
# Praefixen (Office/Python/COM-Reste sowie eigene Skript-Artefakte)
# duerfen geloescht werden. NIEMALS pauschal leeren: Fremdprozesse
# legen aktive Daten ohne Lock in %TEMP% ab; blindes Loeschen
# zerstoert sie.
_WINDOWS_TEMP_WHITELIST_PREFIXES = (
    "~$", "~df", "gen_py", "vbe", "excel8.0",
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
                except Exception:
                    pass
            elif entry.is_dir():
                try:
                    shutil.rmtree(entry.path, ignore_errors=True)
                    if not os.path.exists(entry.path):
                        removed_dirs += 1
                except Exception:
                    pass
        except Exception:
            pass
    try:
        detail_logger.info(
            f"Windows-Temp-Cleanup (Whitelist): {removed_files} Dateien, "
            f"{removed_dirs} Ordner aus {win_temp} entfernt, "
            f"{skipped} fremde Einträge unangetastet.")
    except Exception:
        pass


# ==================================================================
# PowerPoint-Instanz-Lifecycle
# ==================================================================

def _start_ppt_instance():
    ppt = win32com.client.DispatchEx("PowerPoint.Application")

    orig_security = None
    try:
        orig_security = ppt.AutomationSecurity
        ppt.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
    except Exception:
        detail_logger.warning("AutomationSecurity konnte nicht gesetzt werden")

    ppt.Visible = COM_TRUE

    try:
        ppt.WindowState = PP_WINDOW_MINIMIZED
    except Exception:
        pass
    try:
        ppt.DisplayAlerts = PP_ALERTS_NONE
    except Exception:
        pass
    try:
        ppt.FeatureInstall = MSO_FEATURE_INSTALL_NONE
    except Exception:
        pass

    pid = _get_ppt_pid(ppt)
    try:
        detail_logger.info(
            f"PowerPoint gestartet (PID: {pid}, Version: {ppt.Version})")
    except Exception:
        detail_logger.debug(f"PowerPoint gestartet (PID: {pid})")
    return ppt, pid, orig_security


def _quit_ppt_instance(ppt, ppt_pid: Optional[int], orig_security=None,
                       create_time: Optional[float] = None) -> None:
    if ppt is not None:
        if orig_security is not None:
            try:
                ppt.AutomationSecurity = orig_security
            except Exception:
                pass
        try:
            ppt.Quit()
            time.sleep(1)
        except Exception:
            pass
    _kill_powerpoint_by_pid(ppt_pid, create_time)


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

    def _watchdog():
        if not done_event.wait(timeout):
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
        return call_fn()
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
                pres.SaveAs(test_path, FileFormat=PP_FORMAT_PPTX)
                pres.Close()
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
            except Exception: pass
        if ppt is not None:
            _quit_ppt_instance(ppt, ppt_pid, orig_security, ppt_create_time)
        try:
            if os.path.exists(test_path):
                os.remove(test_path)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


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


def append_done(file_path: str) -> None:
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
        except EOFError:
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
            except EOFError:
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
    except EOFError:
        return default
    return val if val else default


def ask_yes_no(prompt: str, default_yes: bool = False) -> bool:
    hint = "[J/n]" if default_yes else "[j/N]"
    while True:
        try:
            answer = input(f"{prompt} {hint}: ").strip().lower()
        except EOFError:
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


def dry_run_directory(directory: str) -> dict:
    """Listet auf, was der Echtlauf tun WUERDE - insbesondere welche
    .ppt/.pps/.pot konvertiert (und deren Originale ersetzt) wuerden.
    Es wird nichts geoeffnet, gespeichert oder geloescht; PowerPoint wird
    nicht gestartet. Identisch aufgebaut zu 4a/4b.

    PowerPoint-Besonderheit: Die COM-Schnittstelle kennt keinen
    Password-Parameter, verschluesselte Dateien laufen im Echtlauf in
    einen Timeout. Ein zuverlaessiger Vorab-Check dafuer existiert hier
    nicht, daher entfaellt die Kategorie VERSCHLUESSELT."""
    stats = {"VERARBEITEN": 0, "KONVERTIEREN": 0, "GESPERRT": 0}
    print("\nPROBELAUF – es wird nichts geändert.\n")
    print("-" * 66)
    for file_path in file_generator(directory):
        ext = os.path.splitext(file_path)[1].lower()
        if _is_locked_by_other_user(file_path):
            stats["GESPERRT"] += 1
            tag = "GESPERRT     "
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
    print(f"    GESAMT:                                  {sum(stats.values())}")
    return stats


def file_generator(directory: str, skip_paths: Optional[set] = None):
    extensions = {
        ".pptx", ".pptm", ".ppsx", ".ppsm", ".potx", ".potm",
        ".ppt",  ".pps",  ".pot",
    }
    skip_paths = skip_paths or set()
    queue = [directory]
    while queue:
        current_dir = queue.pop()
        entries = None
        for scandir_attempt in range(2):
            try:
                with os.scandir(current_dir) as it:
                    entries = list(it)
                break
            except (PermissionError, OSError):
                if scandir_attempt == 0:
                    time.sleep(1)
                    continue
        if entries is None:
            continue
        for entry in entries:
            try:
                if entry.is_symlink():
                    continue
                if entry.is_dir(follow_symlinks=False):
                    if entry.name.lower() in EXCLUDE_DIR_NAMES:
                        continue
                    # Eigenen Arbeitsordner nicht mitverarbeiten.
                    d_nc = os.path.normcase(os.path.abspath(entry.path))
                    t_nc = os.path.normcase(os.path.abspath(TEMP_BASE_PATH))
                    if d_nc == t_nc or d_nc.startswith(t_nc + os.sep):
                        continue
                    queue.append(entry.path)
                elif entry.is_file(follow_symlinks=False):
                    nl = entry.name.lower()
                    if nl.startswith("~$") or nl.startswith("._"):
                        continue
                    if any(nl.endswith(ext) for ext in extensions):
                        if entry.path in skip_paths:
                            continue
                        yield entry.path
            except (PermissionError, OSError):
                continue


# ==================================================================
# Shape-Schriftart setzen (rekursiv, mit Tiefenlimit)
# ==================================================================

def _set_chart_fonts(chart, font_name: str) -> None:
    try:
        if chart.HasTitle:
            chart.ChartTitle.Characters().Font.Name = font_name
    except Exception:
        pass

    for axis_type in (1, 2, 3):
        for axis_group in (1, 2):
            try:
                ax = chart.Axes(axis_type, axis_group)
                try:
                    if ax.HasTitle:
                        ax.AxisTitle.Characters().Font.Name = font_name
                except Exception:
                    pass
                try:
                    ax.TickLabels.Font.Name = font_name
                except Exception:
                    pass
            except Exception:
                pass

    try:
        if chart.HasLegend:
            chart.Legend.Font.Name = font_name
    except Exception:
        pass

    try:
        for series in chart.SeriesCollection():
            try:
                if series.HasDataLabels:
                    series.DataLabels().Font.Name = font_name
            except Exception:
                pass
            try:
                for trendline in series.Trendlines():
                    try:
                        if trendline.HasLabel:
                            trendline.DataLabel.Font.Name = font_name
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass

    try:
        chart.ChartArea.Font.Name = font_name
    except Exception:
        pass


def _set_wordart_font(shape, font_name: str) -> bool:
    try:
        tef = shape.TextEffect
        if tef is not None:
            tef.FontName = font_name
            detail_logger.debug("WordArt-Font gesetzt")
            return True
    except Exception:
        pass
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
        except Exception:
            pass

        # --- Textrahmen ---
        if shape.HasTextFrame:
            tr = shape.TextFrame.TextRange
            try:
                tr.Font.Name = font_name
            except Exception:
                pass
            try:
                if tr.Length > 0:
                    for run in tr.Runs():
                        try:
                            run.Font.Name = font_name
                        except Exception:
                            pass
            except Exception:
                pass

        # --- Tabelle ---
        if shape.HasTable:
            try:
                tbl = shape.Table
                for row in range(1, tbl.Rows.Count + 1):
                    for col in range(1, tbl.Columns.Count + 1):
                        try:
                            cell = tbl.Cell(row, col)
                            if cell.Shape.HasTextFrame:
                                tr = cell.Shape.TextFrame.TextRange
                                try:
                                    tr.Font.Name = font_name
                                except Exception:
                                    pass
                                try:
                                    if tr.Length > 0:
                                        for run in tr.Runs():
                                            try:
                                                run.Font.Name = font_name
                                            except Exception:
                                                pass
                                except Exception:
                                    pass
                        except Exception:
                            pass
            except Exception:
                pass

        # --- Gruppe (rekursiv) ---
        if shape.Type == MSO_TYPE_GROUP:
            try:
                for sub_shape in shape.GroupItems:
                    _process_shape_font(sub_shape, font_name, depth + 1)
            except Exception:
                pass

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
                            tf2.TextRange.Font.Name = font_name
                            font_set = True
                    except Exception:
                        pass
                    if not font_set:
                        try:
                            tf = node.TextFrame
                            if tf is not None:
                                tf.TextRange.Font.Name = font_name
                                font_set = True
                        except Exception:
                            pass
                    if not font_set:
                        smartart_failed_nodes += 1

                for node in sa.AllNodes:
                    _process_smartart_node(node)
                if smartart_failed_nodes > 0:
                    detail_logger.warning(
                        f"SmartArt: {smartart_failed_nodes}/{smartart_total_nodes} "
                        f"Knoten nicht änderbar – ggf. manuell prüfen")
        except Exception:
            pass

        # --- Chart ---
        if shape.HasChart:
            try:
                _set_chart_fonts(shape.Chart, font_name)
            except Exception:
                pass

    except Exception:
        pass


def _process_shapes_collection(shapes, font_name: str) -> None:
    try:
        count = shapes.Count
    except Exception:
        try:
            for shape in shapes:
                _process_shape_font(shape, font_name)
        except Exception:
            pass
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
                        except Exception:
                            pass
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
    global ppt_app_global, ppt_pid_global

    pbar.write(f"Prüfe: {os.path.basename(file_path)}")
    detail_logger.info(f"=== Starte: {file_path} ===")

    original_path  = file_path
    is_temp_copy   = False
    was_converted  = False
    presentation   = None
    temp_conv_path = None
    # temp_stage2_path haelt den Pfad zur Stage-2-SaveAs-Datei, in die
    # die finalen Font-Ersetzungen geschrieben werden (statt direkt ins
    # Original). Erst nach Close + erfolgreichem _safe_move wird das
    # Original ueberschrieben. Schuetzt vor Datenverlust bei Crash
    # (Watchdog-Kill nach 240s, Netzwerk-Drop, COM-Fehler) waehrend Save.
    temp_stage2_path = None
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

        try:
            if presentation.Final:
                presentation.Final = False
                detail_logger.debug("Final-Markierung aufgehoben")
        except Exception:
            pass

        try:
            presentation.RemovePersonalInformation = False
            detail_logger.debug("RemovePersonalInformation deaktiviert")
        except Exception:
            pass

        # --- Schema-Migration ---
        new_path = None

        if ext in (".ppt", ".pps", ".pot"):
            try:
                has_macros = False
                try:
                    has_macros = presentation.HasVBProject
                except Exception:
                    pass

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
                detail_logger.warning(f"Konvertierung fehlgeschlagen: {e}")

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
                    detail_logger.warning(f"Re-Serialisierung fehlgeschlagen: {e}")
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
                        detail_logger.warning(
                            f"Stage-2-SaveAs fehlgeschlagen: {e}")
                        # Wenn das nicht klappt, ist temp_stage2_path ungueltig -
                        # reset, damit der finally-Block ihn nicht versucht zu loeschen.
                        temp_stage2_path = None

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
                    except Exception:
                        pass
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
                        except Exception:
                            pass

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
            except Exception:
                pass
        if is_temp_copy and os.path.exists(_long_path(file_path)):
            _safe_remove_with_retry(file_path)
        if temp_conv_path and os.path.exists(_long_path(temp_conv_path)):
            _safe_remove_with_retry(temp_conv_path)
        # Stage-2-Tempdatei: nur loeschen wenn sie noch da ist und nicht
        # erfolgreich verschoben wurde. Bei Fehler wurde sie bereits zur
        # RESCUE_-Datei kopiert; das hier ist nur Aufraeumarbeit.
        if temp_stage2_path and os.path.exists(_long_path(temp_stage2_path)):
            _safe_remove_with_retry(temp_stage2_path)


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
    global ppt_create_time_global
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
                try:
                    # --- Geplanter Restart ---
                    if file_counter > 0 and file_counter % COM_RESTART_INTERVAL == 0:
                        detail_logger.info(
                            f"PPT-Neustart nach {file_counter} Dateien")
                        _quit_ppt_instance(ppt, ppt_pid, orig_security,
                                           ppt_create_time)
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
                        except Exception:
                            pass
                        try:
                            _cleanup_user_recent()
                        except Exception:
                            pass
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
                        except Exception:
                            pass
                        ppt             = None
                        ppt_pid         = None
                        ppt_create_time = None
                        orig_security   = None
                        # Cache-Cleanup nach Watchdog-Kill (PowerPoint tot).
                        try:
                            _cleanup_ppt_inetcache()
                        except Exception:
                            pass
                        try:
                            _cleanup_user_recent()
                        except Exception:
                            pass
                except Exception as e:
                    log_error(file_path, e)
                    stats["ERROR"] += 1
                    pbar.write(f"  !!  KRITISCH: {os.path.basename(file_path)}")
                finally:
                    file_counter += 1
                    pbar.update(1)

    except Exception as e:
        print(f"\nKRITISCHER FEHLER: PowerPoint konnte nicht gestartet werden: {e}")
        log_error("GLOBAL", e)

    finally:
        _quit_ppt_instance(ppt, ppt_pid, orig_security, ppt_create_time)
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
        except Exception:
            pass
        try:
            _cleanup_user_recent()
        except Exception:
            pass

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
    # Vor dem ersten COM-Zugriff auf bereits laufende Sitzungen hinweisen.
    warn_running_powerpoint(auto_mode)

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
        if not auto_mode:
            try:
                input("\nBeliebige Taste drücken, um das Fenster zu schließen ...")
            except EOFError:
                pass
        sys.exit(0)

    # --- PowerPoint Smoke-Test (nach Bestaetigung, vor Hauptlauf) ---
    # Prueft mit einer kurzlebigen PowerPoint-Instanz, ob PowerPoint im
    # TEMP_PROCESS_PATH eine .pptx erzeugen, oeffnen und schliessen kann.
    # Bei Geschuetzte-Ansicht-Hänger killt der Watchdog PowerPoint nach
    # 25 s. Ohne diesen Test wuerde ein falsch konfiguriertes Trust-
    # Center die spaeteren Open-Aufrufe pro Datei minutenlang blockieren.
    # Kategorien: "com" = COM-Subsystem-Problem, "ppt" = PowerPoint/Trust-Center.
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

    pythoncom.CoInitialize()

    # Restore-Privilegien (Admin-Kontext) für ACL/Owner-Erhalt aktivieren.
    _enable_restore_privileges()

    start_time   = datetime.now()
    stats_result = None

    try:
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
    try:
        input("Beliebige Taste drücken, um das Fenster zu schließen ...")
    except EOFError:
        pass
    _cleanup_windows_temp()
