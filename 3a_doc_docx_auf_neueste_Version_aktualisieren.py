# ==================================================================
# WORD KOMPATIBILITÄTSMODUS UPDATER
# ==================================================================
# Datum: 13.06.2026
# ANLEITUNG:
# 1. Python installieren: https://www.python.org/
# 2. In der Konsole (CMD): pip install pywin32 tqdm psutil msoffcrypto-tool
# 3. Falls pywin32 erstmals installiert: python Scripts/pywin32_postinstall.py -install
# 4. In der Konsole (CMD): python 3a_doc_docx_auf_neueste_Version_aktualisieren.py
#
# WICHTIG: WORD SICHERHEITSEINSTELLUNGEN (TRUST CENTER)
# ------------------------------------------------------------------
# Word öffnen → Datei → Optionen → Trust Center →
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
#    (Erforderlich für korrekte .docm/.docx-Unterscheidung)
#
# 3. VERTRAUENSWÜRDIGE SPEICHERORTE:
#    ☑ Vertrauenswürdige Speicherorte im Netzwerk zulassen
#    Folgende Speicherorte hinzufügen:
#    • C:\Users\<Benutzername>\Documents (temporärer Arbeitsordner des Skripts)
#    • Q:\ oder \\server\dfs (Quelldateien)
#    Jeweils: ☑ Unterordner ... sind ebenfalls vertrauenswürdig
#
# HINWEIS FÜR PyInstaller-Build (Umwandlung in .exe):
#   # Einmalig: PyInstaller installieren
#   pip install pyinstaller
#
#   # Build als ein einzeiliger Befehl (am sichersten in jeder Shell):
#   python -m PyInstaller --onefile --noupx --noconfirm --clean --console --icon="python_icon.ico" --name "3a_doc_docx_auf_neueste_Version_aktualisieren" --hidden-import pywintypes --hidden-import pythoncom --hidden-import win32api --hidden-import win32file --hidden-import win32process --hidden-import win32con --hidden-import win32com.client --hidden-import msoffcrypto --hidden-import psutil --hidden-import tqdm 3a_doc_docx_auf_neueste_Version_aktualisieren.py
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
# ==================================================================

import argparse
import email.mime.multipart
import email.mime.text
import hashlib
import json
import os
import smtplib
import stat
import sys
import shutil
import signal
import logging
import importlib.util
import gc
import re
from collections import defaultdict
import time
import uuid
from datetime import datetime
from typing import Optional
import threading

# Die Drittanbieter-Module hier gebuendelt und mit Klartext-Meldung. Frueher
# standen sie als nackte Importe im Modulkopf: fehlte eines, brach das Skript
# mit einem ModuleNotFoundError-Traceback ab, BEVOR die eigene Pruefung
# check_required_modules() ueberhaupt lief - sie konnte deshalb nie ausloesen
# und ihre verstaendliche 'pip install'-Anleitung nie erscheinen.
try:
    import pythoncom
    import win32com.client
    import win32process
    import psutil
    from tqdm import tqdm
    import msoffcrypto
except ImportError as _e_imp:
    _paket = {
        "pythoncom": "pywin32", "win32com": "pywin32", "win32process": "pywin32",
        "psutil": "psutil", "tqdm": "tqdm", "msoffcrypto": "msoffcrypto-tool",
    }.get((getattr(_e_imp, "name", "") or "").split(".")[0], "")
    print("=" * 66)
    print("❌ FEHLENDES MODUL:", getattr(_e_imp, "name", _e_imp))
    if _paket:
        print(f"   pip install {_paket}")
    print("=" * 66)
    raise SystemExit(1)

# ==================================================================
# Skript-Metadaten
# ==================================================================
SCRIPT_VERSION = "1.2.0"
SCRIPT_DATE    = "14.06.2026"

# ==================================================================
# Konsolen-Encoding (UTF-8) – muss VOR jedem print stehen
# ==================================================================
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
# Fehlt _gemeinsam.py, laeuft alles unveraendert weiter.
try:
    _eigener_ordner = os.path.dirname(os.path.abspath(__file__))
    if _eigener_ordner not in sys.path:
        sys.path.insert(0, _eigener_ordner)
    import _gemeinsam as gem
except Exception:
    gem = None


# ==================================================================
# COM- und Format-Konstanten
# ==================================================================
COM_TRUE  = -1
COM_FALSE =  0

# WdSaveFormat
WD_FORMAT_DOCX = 12
WD_FORMAT_DOCM = 13
WD_FORMAT_DOTX = 14
WD_FORMAT_DOTM = 15

# HRESULT-Konstanten für Fehler-Klassifizierung
HR_RPC_E_CALL_REJECTED     = -2147418111  # 0x80010001 – Word beschäftigt
HR_RPC_E_SERVERCALL_RETRY  = -2147417846  # 0x8001010A
HR_RPC_E_SERVERCALL_RETRY2 = -2147417845  # 0x8001010B
HR_E_ACCESSDENIED          = -2147024891  # 0x80070005
HR_E_FILE_NOT_FOUND        = -2147024894  # 0x80070002

HR_TRANSIENT = {HR_RPC_E_CALL_REJECTED, HR_RPC_E_SERVERCALL_RETRY, HR_RPC_E_SERVERCALL_RETRY2}
HR_PERMANENT = {HR_E_ACCESSDENIED, HR_E_FILE_NOT_FOUND}

HR_VBA_NO_ACCESS = {-2146822880, -2146823146, -2146823136}
HR_OPEN_ESCALATE = {HR_RPC_E_CALL_REJECTED, -2146823146, -2146823136}

ERROR_SHARING_VIOLATION = 32

# ==================================================================
# Konfiguration
# ==================================================================
TEMP_PROCESS_PARENT = os.path.join(os.path.expanduser("~"), "Documents")
TEMP_PROCESS_PREFIX = "3a_doc_docx_auf_neueste_Version_aktualisieren_"
TEMP_PROCESS_PATH   = os.path.join(
    TEMP_PROCESS_PARENT, f"{TEMP_PROCESS_PREFIX}{os.getpid()}"
)

MAX_PATH_LEN           = 240
DEFAULT_PASSWORD       = ""
LOG_BASENAME           = "3a_doc_docx_auf_neueste_Version_aktualisieren"
MAX_RETRIES            = 3
RETRY_DELAY            = 2
CRASH_DELAY            = 10
WORD_RESTART_INTERVAL  = 200
SAVEAS_TIMEOUT         = 240.0
# Open-Timeout grosszuegig: eine grosse .doc (50+ MB, eingebettete
# Objekte) ueber VPN/DFS braucht beim Open leicht > 20 s. Ein zu
# knapper Wert killt gesundes Word, der Retry laeuft erneut ins
# Timeout -> ERROR trotz intakter Datei. Der Smoke-Test bleibt bei
# 25 s (winzige Test-Datei).
OPEN_TIMEOUT           = 180.0
PER_FILE_WARN_SECONDS  = 90.0

# Verzeichnisnamen, die nie betreten werden: geloeschte Dokumente im
# $RECYCLE.BIN wuerden sonst mitkonvertiert (und ihr Original geloescht);
# ~snapshot/.snapshot sind read-only NAS-Schattenkopien.
EXCLUDE_DIR_NAMES = {"$recycle.bin", "system volume information",
                     "~snapshot", ".snapshot"}

FILE_ATTRIBUTE_REPARSE_POINT = 0x0400

# AppVersion-Schwelle (siehe get_ooxml_app_version() weiter unten):
#   16.0 = Office 2016 / 2019 / 2021 / 2024 / 365
#   15.0 = Office 2013
TARGET_APP_VERSION_DEFAULT = 16.0

# AV-Wartelogik nach Schreiben in den Temp-Ordner
AV_WAIT_TIMEOUT  = 15.0
AV_WAIT_INTERVAL = 0.3

LOCK_FILE = os.path.join(
    os.environ.get("TEMP", os.environ.get("TMP", r"C:\tmp")),
    "3a_doc_docx_auf_neueste_Version_aktualisieren.lock"
)

# ==================================================================
# Globale Referenzen
# ==================================================================
word_app_global: Optional[win32com.client.CDispatch] = None
word_pid_global: Optional[int] = None
# Profil-persistente Word-Options, die das Skript umstellt. Beim ERSTEN
# _init_word_app gesichert (ueberdauert die vielen Word-Neustarts eines
# Laufs) und im finalen finally wiederhergestellt, damit die Word-
# Einstellungen des Benutzers nach dem Lauf unveraendert sind.
_orig_word_options: dict = {}
log_dir_global: Optional[str] = None
log_file_path_global: Optional[str] = None
detailed_log_file_path_global: Optional[str] = None
summary_json_path_global: Optional[str] = None
conversions_csv_path_global: Optional[str] = None
run_timestamp_global: str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# ==================================================================
# Logger-Platzhalter (Handler in _setup_logging gesetzt)
# ==================================================================
file_logger   = logging.getLogger("FileLogger")
detail_logger = logging.getLogger("DetailLogger")
file_logger.setLevel(logging.ERROR)
detail_logger.setLevel(logging.DEBUG)


def _setup_logging(log_dir: str) -> tuple:
    global log_file_path_global, detailed_log_file_path_global, summary_json_path_global
    global conversions_csv_path_global

    os.makedirs(log_dir, exist_ok=True)

    ts = run_timestamp_global
    log_file          = os.path.join(log_dir, f"{LOG_BASENAME}_{ts}.log")
    detailed_log_file = os.path.join(log_dir, f"{LOG_BASENAME}_detailed_{ts}.log")
    summary_json      = os.path.join(log_dir, f"{LOG_BASENAME}_summary_{ts}.json")
    conversions_csv   = os.path.join(log_dir, f"{LOG_BASENAME}_konvertierungen_{ts}.csv")

    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    fh = logging.FileHandler(log_file, mode="w", encoding="utf-8-sig")
    fh.setFormatter(fmt)
    file_logger.addHandler(fh)

    dh = logging.FileHandler(detailed_log_file, mode="w", encoding="utf-8-sig")
    dh.setFormatter(fmt)
    detail_logger.addHandler(dh)

    log_file_path_global          = log_file
    detailed_log_file_path_global = detailed_log_file
    summary_json_path_global      = summary_json
    conversions_csv_path_global   = conversions_csv
    return log_file, detailed_log_file



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
    # Nachweisliste der .doc/.dot -> .docx/.dotx-Konvertierungen (und
    # Ausweichnamen bei Kollision). Die JSON-Summary enthaelt nur Zaehler;
    # diese CSV haelt die konkreten Pfad-Paare fuers Archiv fest.
    if not conversions_csv_path_global:
        return
    try:
        import csv as _csv
        needs_header = not os.path.exists(conversions_csv_path_global)
        with open(conversions_csv_path_global, "a", encoding="utf-8-sig", newline="") as fh:
            writer = _csv.writer(fh, delimiter=";")
            if needs_header:
                writer.writerow(["timestamp", "alter_pfad", "neuer_pfad"])
            writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                old_path, new_path,
            ])
    except Exception as e:
        detail_logger.warning(f"Konvertierungs-CSV nicht schreibbar: {e}")


def log_error(file_path: str, exc: Exception) -> None:
    file_logger.error(f"Datei: {file_path}\n  -> {exc}\n")
    detail_logger.error(f"Datei: {file_path}\n  -> {exc}\n")


# ==================================================================
# Signal-Handler
# ==================================================================
def _signal_handler(sig, frame) -> None:
    print("\n\n*** ABBRUCH durch Benutzer – raeume auf ...")

    global word_app_global, word_pid_global
    if word_app_global is not None:
        # Profil-persistente Word-Optionen VOR dem Quit zuruecksetzen. Sie
        # ueberdauern das Skript im Benutzerprofil; nach einem Strg+C blieben
        # sie dauerhaft veraendert, weil der Handler die Instanz beendet und
        # das finale finally danach keine lebende COM-Instanz mehr vorfindet.
        try:
            _restore_word_options(word_app_global)
        except Exception as _e:
            detail_logger.debug(f"_signal_handler: Restore verworfen: {_e!r}")
        try:
            word_app_global.Quit(SaveChanges=COM_FALSE)
        except Exception as _e:
            detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")
        word_app_global = None
        word_pid_global = None

    _kill_orphaned_word()

    if os.path.exists(TEMP_PROCESS_PATH):
        try:
            shutil.rmtree(TEMP_PROCESS_PATH)
            print(f"  Temp-Verzeichnis bereinigt: {TEMP_PROCESS_PATH}")
        except Exception as e:
            print(f"  Temp-Bereinigung fehlgeschlagen: {e}")

    try:
        pythoncom.CoUninitialize()
    except Exception as _e:
        detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")

    _release_lock()
    sys.exit(1)


# ==================================================================
# Long-Path- und Datei-Hilfsfunktionen
# ==================================================================

def prepare_long_path(path: str) -> str:
    if path.startswith("\\\\?\\"):
        return path
    path = os.path.abspath(os.path.normpath(path))
    if path.startswith("\\\\"):
        return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path


def long_path(path: str) -> str:
    return prepare_long_path(path) if len(path) > MAX_PATH_LEN else path


def _strip_lp(path: str) -> str:
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\"):
        return path[4:]
    return path


def _is_under_temp_dir(path: str) -> bool:
    try:
        p = os.path.normcase(os.path.abspath(path))
        t = os.path.normcase(os.path.abspath(TEMP_PROCESS_PATH))
        return p.startswith(t + os.sep) or p == t
    except Exception:
        return False


def _clear_readonly(path: str) -> None:
    try:
        import win32api, win32con, win32file
        attrs = win32file.GetFileAttributes(path)
        if attrs != -1 and (attrs & win32con.FILE_ATTRIBUTE_READONLY):
            win32api.SetFileAttributes(path, attrs & ~win32con.FILE_ATTRIBUTE_READONLY)
    except Exception:
        try:
            os.chmod(path, stat.S_IWRITE)
        except Exception as _e:
            detail_logger.debug(f"_clear_readonly: Exception verworfen: {_e!r}")


def safe_remove(path: str) -> bool:
    try:
        p = long_path(path)
        if not os.path.exists(p):
            return True
        _clear_readonly(p)
        os.remove(p)
        detail_logger.debug(f"Gelöscht: {path}")
        return True
    except Exception as e:
        detail_logger.warning(f"Löschen fehlgeschlagen: {path} – {e}")
        return False



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

def safe_exists(path: str) -> bool:
    try:
        return os.path.exists(long_path(path))
    except Exception:
        return False


def _lp_for_reserve(p: str) -> str:
    try:
        return long_path(p)
    except Exception:
        return p


def reserve_unique_path(base: str, ext: str, max_tries: int = 100) -> Optional[str]:
    """Reserviert einen freien Ausweichnamen ATOMAR.

    'Pruefen und danach benutzen' laesst ein Zeitfenster offen, in dem ein
    parallel laufender Durchgang - oder ein Nutzer, der gerade speichert -
    denselben Namen belegen kann; die Datei wuerde beim anschliessenden
    Schreiben ueberschrieben.

    os.open(..., O_CREAT | O_EXCL) schlaegt fehl, wenn die Datei bereits
    existiert. Das Anlegen der 0-Byte-Platzhalterdatei IST damit die
    Reservierung. Der spaetere Schreibvorgang ueberschreibt den Platzhalter;
    kommt es nicht dazu, entfernt release_unique_path() ihn wieder.

    Rueckgabe: reservierter Pfad, oder None wenn alle Varianten belegt sind.
    """
    for counter in range(2, max_tries + 1):
        candidate = f"{base}_{counter}{ext}"
        try:
            fd = os.open(_lp_for_reserve(candidate),
                         os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o666)
            os.close(fd)
            return candidate
        except FileExistsError:
            continue
        except OSError:
            continue
    # Letzte Rettung: zufaelliger Name (wie in 3b/3c bereits ueblich)
    for _ in range(10):
        candidate = f"{base}_{uuid.uuid4().hex[:8]}{ext}"
        try:
            fd = os.open(_lp_for_reserve(candidate),
                         os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o666)
            os.close(fd)
            return candidate
        except OSError:
            continue
    return None


def release_unique_path(path: Optional[str]) -> None:
    """Entfernt einen reservierten Platzhalter - aber NUR wenn er 0 Byte hat.

    Eine Datei mit Inhalt wird unter keinen Umstaenden angefasst.
    """
    if not path:
        return
    try:
        p = _lp_for_reserve(path)
        if os.path.isfile(p) and os.path.getsize(p) == 0:
            os.remove(p)
    except Exception as _e:
        detail_logger.debug(f"release_unique_path: Exception verworfen: {_e!r}")


def safe_getsize(path: str) -> int:
    try:
        return os.path.getsize(long_path(path))
    except Exception:
        return 0


def verify_file(path: str, min_size: int = 100, wait_retries: int = 3) -> bool:
    for _ in range(wait_retries):
        if safe_exists(path):
            if safe_getsize(path) >= min_size:
                return True
        time.sleep(0.5)
    detail_logger.warning(f"Datei fehlt oder zu klein (Timeout): {path}")
    return False


# ==================================================================
# AV-Scanner-Wartelogik
# ==================================================================

def wait_for_file_available(path: str,
                            timeout: float = AV_WAIT_TIMEOUT,
                            interval: float = AV_WAIT_INTERVAL) -> bool:
    try:
        import win32file
        import win32con
        import pywintypes
    except Exception as e:
        detail_logger.debug(f"AV-Wait: pywin32 nicht verfügbar – {e}")
        return True

    deadline = time.time() + timeout
    waited   = False
    while time.time() < deadline:
        try:
            h = win32file.CreateFile(
                long_path(path),
                win32con.GENERIC_READ,
                0,
                None, win32con.OPEN_EXISTING, 0, None
            )
            try:
                h.Close()
            except Exception as _e:
                detail_logger.debug(f"wait_for_file_available: Exception verworfen: {_e!r}")
            if waited:
                detail_logger.debug(f"AV-Wait: Datei freigegeben nach Wartezeit – {path}")
            return True
        except pywintypes.error as e:
            if e.winerror == ERROR_SHARING_VIOLATION:
                waited = True
                time.sleep(interval)
                continue
            detail_logger.debug(f"AV-Wait: unerwarteter Fehler ({e.winerror}) – {path}")
            return False
        except Exception as e:
            detail_logger.debug(f"AV-Wait: Exception – {path}: {e}")
            return False

    detail_logger.warning(f"AV-Wait-Timeout ({timeout:.0f}s) – fortfahren ohne Bestätigung: {path}")
    return False


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
# Robuste Kopier-/Verschiebe-Funktionen
# ==================================================================

def robust_copy(src: str, dst: str, max_retries: int = MAX_RETRIES) -> bool:
    for attempt in range(max_retries):
        try:
            src_s = long_path(src)
            dst_s = long_path(dst)
            if os.path.exists(dst_s):
                _clear_readonly(dst_s)
            shutil.copy2(src_s, dst_s)
            time.sleep(0.5)
            if verify_file(dst):
                detail_logger.debug(f"Kopiert: {src} → {dst}")
                if _is_under_temp_dir(dst):
                    wait_for_file_available(dst)
                return True
        except Exception as e:
            detail_logger.warning(f"Kopieren Versuch {attempt+1}/{max_retries}: {e}")
        if attempt < max_retries - 1:
            time.sleep(RETRY_DELAY)
    return False


def robust_move(src: str, dst: str, max_retries: int = MAX_RETRIES) -> bool:
    if _is_under_temp_dir(src):
        wait_for_file_available(src, timeout=AV_WAIT_TIMEOUT / 2)

    src_s = long_path(src)
    dst_s = long_path(dst)

    try:
        if os.path.exists(dst_s):
            _clear_readonly(dst_s)
        os.replace(src_s, dst_s)
        detail_logger.debug(f"Ersetzt (replace): {src} → {dst}")
        return True
    except OSError as e_rep:
        detail_logger.debug(f"os.replace fehlgeschlagen, fallback staging+replace: {e_rep}")

    # Cross-Volume-Fallback OHNE Truncate-Fenster: src wird zuerst als
    # dst + '.tmp_new' AUF DAS ZIELVOLUME kopiert und dann per os.replace
    # atomar uebergeschoben. shutil.copy2 direkt auf dst wuerde das Ziel
    # sofort auf 0 Bytes trunkieren - ein Abbruch in diesem Fenster
    # (Netzwerk-Drop, AV, Prozess-Kill) hinterliesse eine halbe Datei.
    # Mit Staging ist dst zu jedem Zeitpunkt entweder alt oder neu.
    stage_s = long_path(dst + ".tmp_new")
    for attempt in range(max_retries):
        try:
            shutil.copy2(src_s, stage_s)
            time.sleep(0.5)
            if not verify_file(dst + ".tmp_new"):
                raise Exception("Verifizierung der Staging-Kopie fehlgeschlagen")
            if os.path.exists(dst_s):
                _clear_readonly(dst_s)
            os.replace(stage_s, dst_s)
            safe_remove(src)
            detail_logger.debug(f"Verschoben (staging+replace): {src} → {dst}")
            return True
        # BaseException, nicht Exception: Der Signal-Handler beendet sich mit
        # sys.exit() und loest damit SystemExit aus - das erbt von
        # BaseException. Bei Strg+C mitten im Kopieren blieb die Staging-Kopie
        # '<Ziel>.tmp_new' deshalb im ZIELVERZEICHNIS liegen, also auf der
        # Ablage - und kein Aufraeumpfad des Skripts erfasst sie je wieder.
        except BaseException as e:
            if isinstance(e, Exception):
                detail_logger.warning(f"Verschieben Versuch {attempt+1}/{max_retries}: {e}")
            try:
                if os.path.exists(stage_s):
                    _clear_readonly(stage_s)
                    os.remove(stage_s)
            except Exception as _e:
                detail_logger.debug(f"robust_move: Exception verworfen: {_e!r}")
            # Abbruch nach dem Aufraeumen unveraendert weiterreichen, sonst
            # wuerde Strg+C zu einem blossen 'Versuch fehlgeschlagen'.
            if not isinstance(e, Exception):
                raise
        if attempt < max_retries - 1:
            time.sleep(RETRY_DELAY)
    return False


# ==================================================================
# Lock-Detection für Pre-Flight (Word-Owner-Datei + Probe-Open)
# ==================================================================

def is_locked_by_other(path: str) -> bool:
    try:
        d = os.path.dirname(path)
        b = os.path.basename(path)
        if d and b:
            owner_candidates = ["~$" + b]
            if len(b) > 2:
                owner_candidates.append("~$" + b[2:])
            if len(b) > 1:
                owner_candidates.append("~$" + b[1:])
            for oc in owner_candidates:
                if safe_exists(os.path.join(d, oc)):
                    return True
    except Exception as _e:
        detail_logger.debug(f"is_locked_by_other: Exception verworfen: {_e!r}")

    try:
        import win32file
        import win32con
        import pywintypes
    except Exception:
        return False

    try:
        h = win32file.CreateFile(
            long_path(path),
            win32con.GENERIC_READ,
            win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
            None, win32con.OPEN_EXISTING, 0, None
        )
        try:
            h.Close()
        except Exception as _e:
            detail_logger.debug(f"is_locked_by_other: Exception verworfen: {_e!r}")
        return False
    except Exception as e:
        try:
            if isinstance(e, pywintypes.error) and e.winerror == ERROR_SHARING_VIOLATION:
                return True
        except Exception as _e:
            detail_logger.debug(f"is_locked_by_other: Exception verworfen: {_e!r}")
        return False


# ==================================================================
# Fehler-Klassifizierung
# ==================================================================

def is_transient_error(exc: Exception) -> bool:
    if isinstance(exc, TimeoutError):
        return True
    if isinstance(exc, pythoncom.com_error):
        try:
            if exc.hresult in HR_TRANSIENT:
                return True
        except Exception as _e:
            detail_logger.debug(f"is_transient_error: Exception verworfen: {_e!r}")
    s = str(exc).lower()
    if any(kw in s for kw in ("busy", "timeout", "sharing violation",
                              "network error", "wird gerade verwendet",
                              "freigabeverletzung")):
        return True
    return False


def is_permanent_error(exc: Exception) -> bool:
    if isinstance(exc, (PermissionError, FileNotFoundError)):
        return True
    if isinstance(exc, pythoncom.com_error):
        try:
            if exc.hresult in HR_PERMANENT:
                return True
        except Exception as _e:
            detail_logger.debug(f"is_permanent_error: Exception verworfen: {_e!r}")
    s = str(exc).lower()
    if any(kw in s for kw in ("access denied", "permission denied",
                              "zugriff verweigert", "nicht gefunden",
                              "not found")):
        return True
    return False


# ==================================================================
# Word-Lifecycle
# ==================================================================

def _normalize_user(u: str) -> set:
    u = u.lower().strip()
    forms = {u}
    if "\\" in u:
        forms.add(u.split("\\")[-1])
    if "@" in u:
        forms.add(u.split("@")[0])
    return forms


# ==================================================================
# Schutz fremder Word-Sitzungen
# ==================================================================
# _kill_orphaned_word() beendete trotz seines Namens JEDEN Word-Prozess des
# angemeldeten Benutzers - ohne jede Waisen-Pruefung. Hatte der Anwender
# Word mit ungespeicherter Arbeit offen, waren diese Dokumente beim
# ersten Aufraeumen (auch beim Abbruch mit Strg+C) verloren.
#
# Vor dem Start der eigenen COM-Instanz wird deshalb einmal festgehalten,
# welche Word-Prozesse es bereits gab. Diese gelten dauerhaft als fremd
# und werden nie beendet.
_FOREIGN_WORD_PIDS: set = set()


def snapshot_foreign_word_pids() -> None:
    """Merkt sich alle Word-Prozesse, die vor dem Skriptstart liefen."""
    global _FOREIGN_WORD_PIDS
    found = set()
    try:
        for p in psutil.process_iter(["pid", "name"]):
            try:
                nm = p.info.get("name") or ""
                if "WINWORD.EXE" in nm.upper():
                    found.add(p.info["pid"])
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
    except Exception as _e:
        # Im Zweifel lieber zu viel schuetzen als eine fremde Sitzung killen.
        detail_logger.debug(f"snapshot_foreign_word_pids: Exception verworfen: {_e!r}")
    _FOREIGN_WORD_PIDS = found
    if found:
        detail_logger.info(
            f"{len(found)} bereits laufende(r) Word-Prozess(e) erkannt "
            f"- diese werden nicht beendet: {sorted(found)}")


def is_foreign_word_pid(pid) -> bool:
    return pid in _FOREIGN_WORD_PIDS


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


def _kill_orphaned_word() -> None:
    try:
        cur_user_forms = _normalize_user(psutil.Process().username())
    except Exception:
        detail_logger.warning(
            "_kill_orphaned_word: Benutzerermittlung fehlgeschlagen – "
            "Prozesse werden nicht beendet.")
        return

    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if not (proc.info["name"] and "WINWORD.EXE" in proc.info["name"].upper()):
                continue
            try:
                proc_user_forms = _normalize_user(proc.username())
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
            if not (cur_user_forms & proc_user_forms):
                continue
            if is_foreign_word_pid(proc.info["pid"]):
                continue
            try:
                proc.kill()
                proc.wait(timeout=3)
                detail_logger.debug(f"Word-Prozess beendet: PID {proc.info['pid']}")
            except psutil.AccessDenied:
                detail_logger.warning(f"Kein Zugriff auf PID {proc.info['pid']}")
            except psutil.NoSuchProcess:
                pass
            except psutil.TimeoutExpired:
                detail_logger.warning(f"Prozess reagiert nicht: PID {proc.info['pid']}")
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue


def _find_newest_word_pid() -> Optional[int]:
    try:
        cur_user_forms = _normalize_user(psutil.Process().username())
    except Exception:
        return None
    candidates = []
    for p in psutil.process_iter(["pid", "name", "create_time"]):
        try:
            if not (p.info["name"] and "WINWORD.EXE" in p.info["name"].upper()):
                continue
            try:
                u = _normalize_user(p.username())
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
            if cur_user_forms & u:
                ct = p.info.get("create_time") or 0.0
                candidates.append((ct, p.info["pid"]))
        except Exception:
            continue
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]
    return None


def _cleanup_word_inetcache() -> None:
    local_appdata = os.environ.get("LOCALAPPDATA")
    if not local_appdata:
        return

    cache_subdirs = [
        os.path.join("Microsoft", "Windows", "INetCache", "Content.Word"),
        os.path.join("Microsoft", "Windows", "INetCache", "Content.MSO"),
    ]

    for sub in cache_subdirs:
        cache_dir = os.path.join(local_appdata, sub)

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
                except Exception as _e:
                    detail_logger.debug(f"_cleanup_user_recent: Exception verworfen: {_e!r}")
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


def _init_word_app() -> tuple:
    app = win32com.client.DispatchEx("Word.Application")
    time.sleep(0.3)

    pid = None
    try:
        hwnd = app.Hwnd
        if hwnd:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid == 0:
                pid = None
    except Exception:
        pid = None

    if not pid:
        pid = _find_newest_word_pid()

    app.Visible            = COM_FALSE
    app.DisplayAlerts      = COM_FALSE
    app.AutomationSecurity = 3

    # Profil-persistente Options EINMALIG sichern (vor der ersten
    # Aenderung), damit der finale Restore den Ausgangszustand kennt.
    # Kein 'global' noetig: das Dict wird nur mutiert, nie neu zugewiesen.
    if not _orig_word_options:
        for opt in ("UpdateLinksAtOpen", "DoNotPromptForConvert"):
            try:
                _orig_word_options[opt] = getattr(app.Options, opt)
            except Exception as _e:
                detail_logger.debug(f"_init_word_app: Exception verworfen: {_e!r}")

    try:
        app.Options.UpdateLinksAtOpen = False
    except Exception as _e:
        detail_logger.debug(f"_init_word_app: Exception verworfen: {_e!r}")
    # Konvertierungs-Dialog beim Oeffnen alter Formate unterdruecken -
    # das ist der Kern-Use-Case dieses Skripts und ohne dieses Flag eine
    # Haenger-Quelle, die sonst nur der Watchdog (teuer) abfaengt.
    try:
        app.Options.DoNotPromptForConvert = True
    except Exception as _e:
        detail_logger.debug(f"_init_word_app: Exception verworfen: {_e!r}")
    try:
        app.Options.NoPromptForTemplateID = True
    except Exception as _e:
        detail_logger.debug(f"_init_word_app: Exception verworfen: {_e!r}")
    try:
        app.DisplayRecentFiles = False
    except Exception as _e:
        detail_logger.debug(f"_init_word_app: Exception verworfen: {_e!r}")
    return app, pid


def _restore_word_options(app) -> None:
    # Profil-persistente Options auf den Ausgangswert zuruecksetzen.
    if app is None or not _orig_word_options:
        return
    for opt, val in _orig_word_options.items():
        try:
            setattr(app.Options, opt, val)
        except Exception as _e:
            detail_logger.debug(f"_restore_word_options: Exception verworfen: {_e!r}")


def _clear_recent_files(word_app: win32com.client.CDispatch) -> None:
    try:
        count = word_app.RecentFiles.Count
        for i in range(count, 0, -1):
            try:
                word_app.RecentFiles(i).Delete()
            except Exception as _e:
                detail_logger.debug(f"_clear_recent_files: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"_clear_recent_files: Exception verworfen: {_e!r}")


def _restart_word_engine(old_app=None, pbar=None) -> tuple:
    global word_app_global, word_pid_global
    if old_app is not None:
        try:
            old_app.Quit(SaveChanges=COM_FALSE)
        except Exception as _e:
            detail_logger.debug(f"_restart_word_engine: Exception verworfen: {_e!r}")
    word_app_global = None
    word_pid_global = None
    gc.collect()
    _kill_orphaned_word()
    _cleanup_word_inetcache()
    _cleanup_user_recent()
    time.sleep(CRASH_DELAY)
    app, pid = _init_word_app()
    word_app_global = app
    word_pid_global = pid
    if pbar is not None:
        pbar.write("  ↻  Word-Engine neu gestartet.")
    else:
        detail_logger.info("Word-Engine neu gestartet.")
    return app, pid


def determine_max_compat(word_app: win32com.client.CDispatch) -> int:
    fallback_mode = 15
    try:
        ver   = str(word_app.Version)
        major = int(ver.split(".")[0])
        if major >= 15:
            return 15
        if major == 14:
            return 14
        if major == 12:
            return 12
        if major == 11:
            return 11
    except Exception as e:
        detail_logger.warning(f"Version-Check fehlgeschlagen: {e}")

    try:
        temp_doc = word_app.Documents.Add()
        try:
            mode = temp_doc.CompatibilityMode
            return max(mode, fallback_mode)
        finally:
            temp_doc.Close(SaveChanges=COM_FALSE)
    except Exception:
        return fallback_mode


# ==================================================================
# Watchdog-gestützte COM-Calls (Open / SaveAs)
# ==================================================================

def _word_call_with_watchdog(call_label: str, timeout: float, word_pid,
                             call_fn, restore_fn=None):
    done_event   = threading.Event()
    timeout_flag = [False]

    # Erstellungszeit SYNCHRON im Main-Thread erfassen, BEVOR der Watchdog-
    # Thread startet. Wuerde sie erst im Thread gelesen, koennte Word in den
    # Millisekunden zwischen Thread-Spawn und Lesen abstuerzen, Windows die
    # PID an eine neue (vom Benutzer geoeffnete) Word-Sitzung vergeben - und
    # der Watchdog merkte sich deren create_time als "erwartet" und wuerde
    # nach Timeout den falschen, unbeteiligten Prozess killen.
    expected_ct = None
    if word_pid:
        try:
            expected_ct = psutil.Process(word_pid).create_time()
        except Exception:
            expected_ct = None

    def watchdog():
        if not done_event.wait(timeout):
            timeout_flag[0] = True
            if word_pid:
                try:
                    proc = psutil.Process(word_pid)
                    try:
                        proc_name = proc.name().upper()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        proc_name = ""
                    if proc_name != "WINWORD.EXE":
                        return
                    if expected_ct is not None:
                        try:
                            if abs(proc.create_time() - expected_ct) > 0.001:
                                return
                        except Exception:
                            return
                    proc.kill()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
                except Exception as _e:
                    detail_logger.debug(f"watchdog: Exception verworfen: {_e!r}")

    wd_thread = threading.Thread(target=watchdog, daemon=True)
    wd_thread.start()

    try:
        return call_fn()
    except Exception as e:
        done_event.set()
        if timeout_flag[0]:
            raise TimeoutError(f"Word Timeout bei {call_label}")
        raise e
    finally:
        done_event.set()
        if restore_fn is not None:
            try:
                restore_fn()
            except Exception as _e:
                detail_logger.debug(f"_word_call_with_watchdog: Exception verworfen: {_e!r}")


def safe_word_open(word_app, file_path, pw, is_binary,
                   timeout: float = OPEN_TIMEOUT, word_pid=None):
    _prev_alerts = None
    try:
        _prev_alerts = word_app.DisplayAlerts
    except Exception as _e:
        detail_logger.debug(f"safe_word_open: Exception verworfen: {_e!r}")
    try:
        word_app.DisplayAlerts = COM_FALSE
    except Exception as _e:
        detail_logger.debug(f"safe_word_open: Exception verworfen: {_e!r}")

    def _open():
        return word_app.Documents.Open(
            file_path,
            ConfirmConversions    = COM_FALSE,
            ReadOnly              = COM_TRUE if is_binary else COM_FALSE,
            OpenAndRepair         = COM_FALSE,
            AddToRecentFiles      = COM_FALSE,
            NoEncodingDialog      = COM_TRUE,
            PasswordDocument      = pw,
            PasswordTemplate      = pw,
            WritePasswordDocument = "" if is_binary else pw,
            WritePasswordTemplate = "" if is_binary else pw,
        )

    def _restore():
        if _prev_alerts is not None:
            try:
                word_app.DisplayAlerts = _prev_alerts
            except Exception as _e:
                detail_logger.debug(f"_restore: Exception verworfen: {_e!r}")

    return _word_call_with_watchdog(
        "Documents.Open (Passwort-Dialog/Netzwerk)",
        timeout, word_pid, _open, _restore
    )


def safe_word_saveas(word_app, doc, target_path, file_format, compat_mode,
                     timeout: float = SAVEAS_TIMEOUT, word_pid=None):
    try:
        word_app.DisplayAlerts = COM_FALSE
    except Exception as _e:
        detail_logger.debug(f"safe_word_saveas: Exception verworfen: {_e!r}")

    def _save():
        doc.SaveAs2(
            target_path,
            FileFormat          = file_format,
            CompatibilityMode   = compat_mode,
            AddToRecentFiles    = COM_FALSE,
            Password            = "",
            WritePassword       = "",
            ReadOnlyRecommended = COM_FALSE,
        )

    _word_call_with_watchdog(
        "SaveAs2 (Netzwerk)",
        timeout, word_pid, _save
    )


# ==================================================================
# Trust-Center Smoke-Test
# ==================================================================

def _build_minimal_docx(target_path: str) -> None:
    import zipfile

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
        zf.writestr("_rels/.rels",          rels)
        zf.writestr("word/document.xml",    document)


def test_trust_center_smoke(word_app, word_pid: Optional[int],
                            timeout: float = 25.0) -> tuple:
    os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)

    test_path = os.path.join(
        TEMP_PROCESS_PATH,
        f"trustcheck_{uuid.uuid4().hex}.docx"
    )

    try:
        _build_minimal_docx(test_path)
    except Exception as e:
        return (False, f"Test-DOCX konnte nicht erstellt werden: {e}")

    doc = None
    try:
        def _open_test():
            return word_app.Documents.Open(
                test_path,
                ConfirmConversions = COM_FALSE,
                ReadOnly           = COM_TRUE,
                AddToRecentFiles   = COM_FALSE,
                NoEncodingDialog   = COM_TRUE,
            )

        try:
            doc = _word_call_with_watchdog(
                "Trust-Center-Smoke-Test (Documents.Open)",
                timeout, word_pid, _open_test
            )
        except TimeoutError:
            return (False, f"TIMEOUT nach {timeout:.0f}s "
                            "- Trust Center vermutlich nicht konfiguriert "
                            "oder Geschützte Ansicht aktiv")
        except Exception as e:
            return (False, f"Documents.Open fehlgeschlagen: {e}")

        if doc is None:
            return (False, "Documents.Open lieferte kein Document-Objekt")

        return (True, "")
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try:
            if os.path.exists(test_path):
                os.remove(test_path)
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")


# ==================================================================
# OOXML-AppVersion-Pruefung
# ==================================================================
def get_ooxml_app_version(file_path: str) -> Optional[float]:
    import zipfile

    ext = os.path.splitext(file_path)[1].lower()
    if ext not in (".docx", ".docm", ".dotx", ".dotm"):
        return None

    zip_path = prepare_long_path(file_path) if len(file_path) > MAX_PATH_LEN else file_path

    try:
        with zipfile.ZipFile(zip_path, "r") as z:
            if "docProps/app.xml" not in z.namelist():
                return None
            with z.open("docProps/app.xml") as f:
                content = f.read(65536).decode("utf-8", errors="ignore")
                m = re.search(r"<(?:\w+:)?AppVersion>(\d+\.\d+)</(?:\w+:)?AppVersion>", content)
                if m:
                    return float(m.group(1))
    except (zipfile.BadZipFile, KeyError, ValueError):
        return None
    except PermissionError:
        detail_logger.warning(
            f"AppVersion nicht lesbar - Datei ist gesperrt: {file_path}")
        return None
    except Exception as e:
        detail_logger.warning(f"AppVersion-Lesen fehlgeschlagen ({file_path}): {e}")
        return None

    return None


def _resolve_target_app_version(word_app: win32com.client.CDispatch) -> float:
    try:
        major = int(str(word_app.Version).split(".")[0])
        if major >= 16:
            return 16.0
        if major == 15:
            return 15.0
        if major == 14:
            return 14.0
        if major == 12:
            return 12.0
        return float(major)
    except Exception:
        return TARGET_APP_VERSION_DEFAULT


# ==================================================================
# Modul-Check
# ==================================================================

def check_required_modules() -> None:
    required = {
        "win32com.client": "pywin32",
        "psutil":          "psutil",
        "tqdm":            "tqdm",
        "msoffcrypto":     "msoffcrypto-tool",
    }
    missing = [pkg for mod, pkg in required.items()
               if not importlib.util.find_spec(mod.split(".")[0])]
    if missing:
        print("=" * 66)
        print("❌ FEHLENDE MODULE:")
        for pkg in missing:
            print(f"   pip install {pkg}")
        print("=" * 66)
        sys.exit(1)


# ==================================================================
# Aufräumen verwaister Temp-Verzeichnisse aus früheren Abstürzen
# ==================================================================

def cleanup_orphaned_temp_dirs() -> None:
    try:
        if not os.path.isdir(TEMP_PROCESS_PARENT):
            return
        own_pid = os.getpid()
        for entry in os.listdir(TEMP_PROCESS_PARENT):
            if not entry.startswith(TEMP_PROCESS_PREFIX):
                continue
            full = os.path.join(TEMP_PROCESS_PARENT, entry)
            if not os.path.isdir(full):
                continue
            pid_str = entry[len(TEMP_PROCESS_PREFIX):]
            try:
                pid = int(pid_str)
            except ValueError:
                continue
            if pid == own_pid:
                continue
            try:
                if psutil.pid_exists(pid):
                    continue
            except Exception:
                continue
            try:
                shutil.rmtree(full, ignore_errors=True)
                detail_logger.debug(f"Verwaister Temp-Ordner entfernt: {full}")
            except Exception as _e:
                detail_logger.debug(f"cleanup_orphaned_temp_dirs: Exception verworfen: {_e!r}")
    except Exception as e:
        detail_logger.debug(f"cleanup_orphaned_temp_dirs: {e}")


# Whitelist fuer den Windows-Temp-Cleanup: NUR Word-eigene Reste
# (~$ Owner-Dateien, ~DF*.tmp, VBE-Reste) und eigene Skript-Artefakte.
# NIEMALS pauschal leeren: Fremdprozesse legen aktive Daten ohne Lock
# in %TEMP% ab; blindes Loeschen zerstoert sie. gen_py (pywin32-COM-
# Cache) und excel8.0 gehoeren bewusst NICHT hierher: gen_py wird von
# parallel laufenden COM-Skripten (3b/3c/4a) aktiv genutzt, excel8.0
# ist ein Excel-Artefakt.
_WINDOWS_TEMP_WHITELIST_PREFIXES = (
    "~$", "~df", "vbe",
    "3a_doc_docx_auf_neueste_version_aktualisieren",
)


def _is_whitelisted_temp_entry(name: str) -> bool:
    nl = name.lower()
    return any(nl.startswith(p) for p in _WINDOWS_TEMP_WHITELIST_PREFIXES)


def cleanup_windows_temp() -> None:
    # Best-effort Bereinigung von %LOCALAPPDATA%\Temp am Skriptende -
    # ausschliesslich per Whitelist bekannter Praefixe.
    # Gesperrte/in-Nutzung-Dateien werden stillschweigend uebersprungen.
    win_temp = os.environ.get("TEMP") or os.environ.get("TMP")
    if not win_temp or not os.path.isdir(win_temp):
        return
    try:
        for name in os.listdir(win_temp):
            if not _is_whitelisted_temp_entry(name):
                continue
            full = os.path.join(win_temp, name)
            try:
                if os.path.isdir(full):
                    shutil.rmtree(full, ignore_errors=True)
                else:
                    os.remove(full)
            except Exception as _e:
                detail_logger.debug(f"cleanup_windows_temp: Exception verworfen: {_e!r}")
    except Exception as _e:
        detail_logger.debug(f"cleanup_windows_temp: Exception verworfen: {_e!r}")


# ==================================================================
# Einzelinstanz-Schutz (Lock-File mit PID + EXE-Name)
# ==================================================================

def _own_process_name() -> str:
    try:
        return psutil.Process(os.getpid()).name().lower()
    except Exception:
        return ""


def _acquire_lock() -> bool:
    own_name = _own_process_name()

    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, "r", encoding="utf-8") as lf:
                content = [l.strip() for l in lf.read().splitlines() if l.strip()]
            if content:
                old_pid  = int(content[0])
                old_name = content[1].lower() if len(content) > 1 else ""
                # Dritte Zeile (optional, neues Format): Erstellungszeit des
                # sperrenden Prozesses. Schuetzt vor PID-Recycling: auf
                # Servern mit langer Uptime kann die alte PID laengst an ein
                # ANDERES Python-Skript vergeben sein - ohne create_time-
                # Abgleich wuerde 'python' im Namen den Updater dauerhaft
                # aussperren.
                old_ct = None
                if len(content) > 2:
                    try:
                        old_ct = float(content[2])
                    except ValueError:
                        old_ct = None
                if old_pid != os.getpid() and psutil.pid_exists(old_pid):
                    try:
                        proc = psutil.Process(old_pid)
                        proc_name = proc.name().lower()
                        name_match = (proc_name == own_name
                                      or (old_name and proc_name == old_name)
                                      or "python" in proc_name)
                        if name_match:
                            if old_ct is not None:
                                # Nur sperren, wenn es WIRKLICH derselbe
                                # Prozess ist (Name + Erstellungszeit).
                                try:
                                    if abs(proc.create_time() - old_ct) < 1.0:
                                        return False
                                except (psutil.NoSuchProcess, psutil.AccessDenied):
                                    pass
                            else:
                                # Altes Lock-Format ohne create_time -
                                # konservativ wie bisher behandeln.
                                return False
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
        except (ValueError, OSError, IndexError):
            pass

    try:
        d = os.path.dirname(LOCK_FILE)
        if d:
            os.makedirs(d, exist_ok=True)
        own_ct = ""
        try:
            own_ct = repr(psutil.Process(os.getpid()).create_time())
        except Exception:
            own_ct = ""
        with open(LOCK_FILE, "w", encoding="utf-8") as lf:
            lf.write(f"{os.getpid()}\n{own_name}\n{own_ct}\n")
    except OSError as e:
        detail_logger.warning(f"Lock-File konnte nicht erstellt werden: {e}")
    return True


def _release_lock() -> None:
    try:
        if os.path.exists(LOCK_FILE):
            with open(LOCK_FILE, "r", encoding="utf-8") as lf:
                content = [l.strip() for l in lf.read().splitlines() if l.strip()]
            stored_pid = int(content[0]) if content else 0
            if stored_pid == os.getpid():
                os.remove(LOCK_FILE)
    except Exception as _e:
        detail_logger.debug(f"_release_lock: Exception verworfen: {_e!r}")


# ==================================================================
# Pfad-Sanitisierung
# ==================================================================

def sanitize_path(raw: str) -> str:
    path = raw.strip().strip('"').strip("'")
    if len(path) == 2 and path[1] == ":":
        path = path + "\\"
    return path


# ==================================================================
# Dateinamens-Bereinigung
# ==================================================================

_RESERVED_NAMES = frozenset({
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
})

_ILLEGAL_CHARS = re.compile(r'[<>:"/\\|?*]')

_CONTROL_AND_INVISIBLE = re.compile(
    r'[\x00-\x1f\x7f'
    r'\u200b\u200c\u200d\u2060\ufeff'
    r'\u00ad\u034f\u061c'
    r'\u115f\u1160\u17b4\u17b5'
    r'\u2000-\u200f\u202a-\u202e\u2066-\u2069]'
)


def sanitize_filename(name: str) -> str:
    stem, ext = os.path.splitext(name)

    stem = _CONTROL_AND_INVISIBLE.sub("", stem)
    ext  = _CONTROL_AND_INVISIBLE.sub("", ext)

    stem = _ILLEGAL_CHARS.sub("-", stem)

    stem = stem.strip()
    stem = stem.rstrip(".")

    if not stem:
        stem = "_bereinigt"

    if stem.upper() in _RESERVED_NAMES:
        stem = f"_{stem}"

    return stem + ext


def sanitize_file_on_disk(full_path: str, pbar=None, detail_log=None) -> str:
    directory = os.path.dirname(full_path)
    old_name  = os.path.basename(full_path)
    new_name  = sanitize_filename(old_name)

    if new_name == old_name:
        return full_path

    new_path = os.path.join(directory, new_name)

    reserved_placeholder = None
    if safe_exists(new_path):
        stem, ext = os.path.splitext(new_name)
        # Atomar reservieren statt pruefen-und-spaeter-benutzen. Die alte
        # Schleife brach beim Zaehlerlimit ab UND benutzte den nie
        # geprueften Namen '_100' - anders als die zweite Kollisionsstelle
        # weiter unten, die danach zusaetzlich prueft.
        reserved = reserve_unique_path(os.path.join(directory, stem), ext)
        if not reserved:
            detail_log and detail_log.warning(
                f"Kein freier Ausweichname fuer {new_name} - Umbenennung uebersprungen")
            return full_path
        reserved_placeholder = reserved
        new_path = reserved
        new_name = os.path.basename(new_path)

    try:
        old_p = long_path(full_path)
        new_p = long_path(new_path)
        # os.replace statt os.rename: liegt am Ziel unser eigener 0-Byte-
        # Platzhalter aus reserve_unique_path, wuerde os.rename mit
        # FileExistsError scheitern.
        os.replace(old_p, new_p)
        reserved_placeholder = None
        msg = f"  📝 Umbenannt: '{old_name}' → '{new_name}'"
        if pbar is not None:
            pbar.write(msg)
        if detail_log:
            detail_log.info(f"Dateiname bereinigt: {full_path} → {new_path}")
        return new_path
    except Exception as e:
        msg = f"  ⚠  Umbenennung fehlgeschlagen: '{old_name}' – {e}"
        if pbar is not None:
            pbar.write(msg)
        if detail_log:
            detail_log.warning(msg)
        return full_path
    finally:
        # Reservierten Platzhalter freigeben, falls das Umbenennen
        # nicht bis zum Ende kam (entfernt nur 0-Byte-Dateien).
        release_unique_path(reserved_placeholder)


# ==================================================================
# Known-Folder-Aufloesung via Registry
# ==================================================================

DOWNLOADS_KNOWN_FOLDER_GUID = "{374DE290-123F-4565-9164-39C4925E467B}"


def _get_known_folder(name_or_guid: str) -> Optional[str]:
    """
    Liest einen Pfad aus 'User Shell Folders' und expandiert
    Umgebungsvariablen. Gibt None zurueck, wenn der Wert fehlt
    oder die Registry nicht erreichbar ist.
    """
    try:
        import winreg
    except ImportError:
        return None
    try:
        key_path = (r"Software\Microsoft\Windows\CurrentVersion"
                    r"\Explorer\User Shell Folders")
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            raw, _ = winreg.QueryValueEx(key, name_or_guid)
    except (FileNotFoundError, OSError) as e:
        detail_logger.debug(
            f"Known-Folder '{name_or_guid}' nicht in Registry: {e}")
        return None

    expanded = os.path.expandvars(raw) if raw else None
    if not expanded or "%" in expanded:
        detail_logger.debug(
            f"Known-Folder '{name_or_guid}' nicht aufloesbar: '{raw}'")
        return None
    return expanded


def _resolve_desktop() -> str:
    path = _get_known_folder("Desktop")
    if path and os.path.isdir(path):
        return path
    return os.path.join(os.path.expanduser("~"), "Desktop")


def _resolve_downloads() -> str:
    path = _get_known_folder(DOWNLOADS_KNOWN_FOLDER_GUID)
    if path and os.path.isdir(path):
        return path
    return os.path.join(os.path.expanduser("~"), "Downloads")


# ==================================================================
# Interaktive Abfragen
# ==================================================================

def ask_directory() -> str:
    desktop_path   = _resolve_desktop()
    downloads_path = _resolve_downloads()

    print("\nZielverzeichnis auswählen:")
    print("  [1] Q:\\")
    print("  [2] R:\\")
    print("  [3] G:\\Geteilte Ablagen")
    print("  [4] G:\\Meine Ablage")
    print("  [5] \\\\server\\dfs")
    print(f"  [6] Desktop ({desktop_path})")
    print(f"  [7] Downloads ({downloads_path})")
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
            print("\n⚠ Keine Eingabe möglich. Starte mit --dir und --auto-start.")
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
            path = "\\\\server\\dfs"
        elif choice == "6":
            path = desktop_path
        elif choice == "7":
            path = downloads_path
        elif choice == "8":
            try:
                raw = input("Pfad eingeben: ")
            except (EOFError, RuntimeError, OSError):
                print("\n⚠ Keine Eingabe möglich.")
                sys.exit(1)
            path = sanitize_path(raw)
        else:
            print("  Bitte 1 bis 8 eingeben.")
            continue

        if os.path.isdir(prepare_long_path(path)):
            return path
        print(f"  ❌ Verzeichnis nicht erreichbar: '{path}'")
        print("     Bitte erneut wählen.")


def _is_tty() -> bool:
    """Haengt stdin an einer echten Konsole? (wie in 4b)"""
    try:
        return sys.stdin.isatty()
    except Exception:
        return False


def _warte_auf_taste(auto_mode: bool = False) -> None:
    """Abschliessende Enter-Abfrage - nur wenn wirklich jemand zusieht.

    Ohne diesen Schutz blockierte ein geplanter Lauf mit angehaengter Konsole
    unbegrenzt an der Eingabe. Ist stdin ganz abgeloest (pythonw, Taskplaner
    ohne Benutzeranmeldung), wirft input() ausserdem RuntimeError('lost
    sys.stdin') oder OSError - beides KEIN EOFError, der frueher allein
    abgefangen wurde; das Skript endete dann nach vollstaendig geleisteter
    Arbeit mit einem Traceback.
    """
    if auto_mode or not _is_tty():
        return
    try:
        input("Drücken Sie Enter, um das Fenster zu schließen ...")
    except (EOFError, RuntimeError, OSError):
        pass


def ask_yes_no(prompt: str, default_yes: bool = False) -> bool:
    hint = "[J/n]" if default_yes else "[j/N]"
    while True:
        try:
            answer = input(f"{prompt} {hint}: ").strip().lower()
        except (EOFError, RuntimeError, OSError):
            print("\n⚠ Keine Eingabe möglich (Skript läuft im Hintergrund?). Nutze Standardwert.")
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


def ask_passwords() -> list:
    print()
    print("Passwörter zum Öffnen geschützter Word-Dateien:")
    print("  Jedes Passwort einzeln eingeben und mit Enter bestätigen (max. 3).")
    print("  Kein Komma oder Semikolon – ein Passwort pro Zeile!")
    print("  Leere Eingabe bei Passwort 1 → keine Passwörter hinterlegt.")
    passwords = []
    for slot in range(1, 4):
        try:
            pw = input(f"  Passwort {slot}: ").strip()
        except (EOFError, RuntimeError, OSError):
            break
        if pw:
            passwords.append(pw)
        else:
            break
    if passwords:
        print(f"  → {len(passwords)} Passwort(wörter) hinterlegt.")
        print("    Die Passwörter werden der Reihe nach ausprobiert.")
        print("    Keines passend → Datei wird als SKIPPED geloggt, Skript läuft weiter.")
    else:
        print("  → Kein Passwort hinterlegt.")
        print("    Geschützte Dateien werden automatisch als SKIPPED geloggt.")
    return passwords


def load_pwdfile(path: str) -> list:
    result = []
    try:
        with open(path, encoding="utf-8-sig", errors="replace") as f:
            for line in f:
                pw = line.strip()
                if pw and not pw.startswith("#"):
                    result.append(pw)
        print(f"  → {len(result)} Passwort/Passwörter aus Datei geladen: {path}")
    except Exception as e:
        print(f"  ❌ Passwortdatei nicht lesbar: {e}")
        sys.exit(1)
    return result


def ask_progress_mode() -> tuple:
    """
    Returns (show_progress: bool, count_first: bool).
      [1] Standard – kein tqdm-Balken, jede Datei wird inline geloggt.
      [2] tqdm mit ETA – benötigt Vorab-Scan aller Dateien.
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


# ==================================================================
# Berichts-Funktionen (E-Mail, JSON-Summary)
# ==================================================================

def _build_summary_html(
    stats: dict,
    duration,
    start_dir: str,
    info_counters: dict = None,
    mode_distribution: dict = None,
    error_lines: list = None,
) -> str:
    # Schlanke, inline-gestylte HTML-Tabellenstruktur. Bewusst ohne externe
    # CSS/Bilder (Mail-Clients blocken die ohnehin) - nur Inline-Styles,
    # damit der Bericht auch auf mobilen Clients lesbar bleibt.
    import html as _html

    def esc(v):
        return _html.escape(str(v))

    def rows(pairs):
        out = []
        for label, value in pairs:
            out.append(
                f'<tr><td style="padding:3px 12px 3px 0;">{esc(label)}</td>'
                f'<td style="padding:3px 0;text-align:right;'
                f'font-weight:bold;">{esc(value)}</td></tr>')
        return "".join(out)

    stat_rows = rows([
        ("Aktualisiert",    stats.get("UPDATED", 0)),
        ("Bereits aktuell", stats.get("ALREADY_CURRENT", 0)),
        ("Übersprungen",    stats.get("SKIPPED", 0)),
        ("Resume",          stats.get("RESUMED", 0)),
        ("Fehler",          stats.get("ERROR", 0)),
        ("Gesamt",          sum(stats.values())),
    ])

    info_html = ""
    if info_counters:
        info_pairs = [
            ("Mode-Upgrade",          info_counters.get("MODE_UPGRADED", 0)),
            ("Binär-Konvertierung",   info_counters.get("BINARY_CONVERTED", 0)),
            ("Mit Makros",            info_counters.get("MACROS_FOUND", 0)),
            ("Verschlüsselt",         info_counters.get("ENCRYPTED_SKIPPED", 0)),
            ("Geschützt",             info_counters.get("PROTECTED_SKIPPED", 0)),
        ]
        for key, label in (("WORD_AUTO_REPAIRED", "Word-Auto-Reparatur"),
                           ("APPVERSION_UPDATED", "AppVersion aktualisiert"),
                           ("FILENAME_RENAMED",   "Dateinamen bereinigt"),
                           ("LOCKED_SKIPPED",     "Gesperrt (in Nutzung)"),
                           ("MACRO_SIGNATURE_INVALIDATED",
                            "VBA-Signaturen ungültig (Makros)")):
            if info_counters.get(key, 0) > 0:
                info_pairs.append((label, info_counters[key]))
        info_html = (
            '<h3 style="margin:16px 0 4px;">Details</h3>'
            f'<table style="border-collapse:collapse;font-size:14px;">'
            f'{rows(info_pairs)}</table>')

    mode_html = ""
    if mode_distribution:
        mode_pairs = [(f"Modus {m}", mode_distribution[m])
                      for m in sorted(mode_distribution.keys())]
        mode_html = (
            '<h3 style="margin:16px 0 4px;">Ursprüngliche Compat-Modi</h3>'
            f'<table style="border-collapse:collapse;font-size:14px;">'
            f'{rows(mode_pairs)}</table>')

    error_html = ""
    if error_lines:
        items = "".join(
            f'<li style="font-family:monospace;font-size:12px;">{esc(l.rstrip())}</li>'
            for l in error_lines)
        error_html = (
            '<h3 style="margin:16px 0 4px;color:#b00;">'
            'Fehler-Auszug (letzte 50 Einträge)</h3>'
            f'<ul style="margin:0;padding-left:18px;">{items}</ul>')

    return (
        '<html><body style="font-family:Segoe UI,Arial,sans-serif;'
        'color:#222;font-size:14px;">'
        f'<h2 style="margin:0 0 8px;">Word Kompatibilitätsmodus Updater '
        f'{esc(SCRIPT_VERSION)} – Laufbericht</h2>'
        f'<p style="margin:0 0 4px;"><b>Verzeichnis:</b> {esc(start_dir)}</p>'
        f'<p style="margin:0 0 12px;"><b>Laufzeit:</b> '
        f'{esc(str(duration).split(".")[0])}</p>'
        '<h3 style="margin:8px 0 4px;">Statistik</h3>'
        f'<table style="border-collapse:collapse;font-size:14px;">{stat_rows}</table>'
        f'{info_html}{mode_html}{error_html}'
        '</body></html>')


def send_summary_mail(
    to_addr: str,
    from_addr: str,
    smtp_host: str,
    stats: dict,
    duration,
    start_dir: str,
    log_path: str,
    info_counters: dict = None,
    mode_distribution: dict = None,
    smtp_port: int = 25,
    smtp_user: str = None,
    smtp_pass: str = None,
    smtp_starttls: bool = False,
) -> None:
    subject = (
        f"Word-Updater abgeschlossen – "
        f"{stats.get('UPDATED', 0)} aktualisiert, "
        f"{stats.get('ERROR', 0)} Fehler"
    )

    error_excerpt = ""
    error_lines = []
    try:
        with open(log_path, encoding="utf-8-sig", errors="replace") as f:
            for line in f:
                if "ERROR" in line or "FEHLER" in line:
                    error_lines.append(line)
                    if len(error_lines) > 50:
                        error_lines.pop(0)
        if error_lines:
            error_excerpt = "\nFehler-Auszug (letzte 50 Einträge):\n" + "".join(error_lines)
    except Exception as _e:
        detail_logger.debug(f"send_summary_mail: Exception verworfen: {_e!r}")

    info_section = ""
    if info_counters:
        repair_line = ""
        if info_counters.get("WORD_AUTO_REPAIRED", 0) > 0:
            repair_line = f"  Word-Auto-Reparatur: {info_counters['WORD_AUTO_REPAIRED']}\n"
        appver_line = ""
        if info_counters.get("APPVERSION_UPDATED", 0) > 0:
            appver_line = f"  AppVersion aktualisiert: {info_counters['APPVERSION_UPDATED']}\n"
        rename_line = ""
        if info_counters.get("FILENAME_RENAMED", 0) > 0:
            rename_line = f"  Dateinamen bereinigt: {info_counters['FILENAME_RENAMED']}\n"
        locked_line = ""
        if info_counters.get("LOCKED_SKIPPED", 0) > 0:
            locked_line = f"  Gesperrt (in Nutzung): {info_counters['LOCKED_SKIPPED']}\n"
        info_section = (
            f"\nDETAILS:\n"
            f"  Mode-Upgrade:     {info_counters.get('MODE_UPGRADED', 0)}\n"
            f"  Binär-Konvert.:   {info_counters.get('BINARY_CONVERTED', 0)}\n"
            f"  Mit Makros:       {info_counters.get('MACROS_FOUND', 0)}\n"
            f"  Verschlüsselt:    {info_counters.get('ENCRYPTED_SKIPPED', 0)}\n"
            f"  Geschützt:        {info_counters.get('PROTECTED_SKIPPED', 0)}\n"
            f"{repair_line}"
            f"{appver_line}"
            f"{rename_line}"
            f"{locked_line}"
        )

    mode_section = ""
    if mode_distribution:
        mode_section = "\nURSPRÜNGLICHE COMPAT-MODI:\n"
        for mode in sorted(mode_distribution.keys()):
            mode_section += f"  Modus {mode}: {mode_distribution[mode]}\n"

    body = (
        f"Word Kompatibilitätsmodus Updater {SCRIPT_VERSION} – Laufbericht\n"
        f"{'=' * 60}\n"
        f"Verzeichnis:        {start_dir}\n"
        f"Laufzeit:           {str(duration).split('.')[0]}\n"
        f"\nSTATISTIK:\n"
        f"  Aktualisiert:     {stats.get('UPDATED', 0)}\n"
        f"  Bereits aktuell:  {stats.get('ALREADY_CURRENT', 0)}\n"
        f"  Übersprungen:     {stats.get('SKIPPED', 0)}\n"
        f"  Resume:           {stats.get('RESUMED', 0)}\n"
        f"  Fehler:           {stats.get('ERROR', 0)}\n"
        f"  Gesamt:           {sum(stats.values())}\n"
        f"{info_section}"
        f"{mode_section}"
        f"{error_excerpt}"
    )

    try:
        # multipart/alternative: Plain-Text zuerst, HTML danach. Mail-Clients
        # bevorzugen den letzten (HTML-)Part und zeigen die Tabellenstruktur;
        # Clients ohne HTML fallen sauber auf den Plain-Text zurueck.
        msg = email.mime.multipart.MIMEMultipart("alternative")
        msg["From"]    = from_addr
        msg["To"]      = to_addr
        msg["Subject"] = subject
        msg.attach(email.mime.text.MIMEText(body, "plain", "utf-8"))
        try:
            html_body = _build_summary_html(
                stats, duration, start_dir,
                info_counters=info_counters,
                mode_distribution=mode_distribution,
                error_lines=error_lines)
            msg.attach(email.mime.text.MIMEText(html_body, "html", "utf-8"))
        except Exception as e_html:
            detail_logger.debug(f"HTML-Mail-Teil konnte nicht erzeugt werden: {e_html}")
        recipient_list = [addr.strip() for addr in to_addr.split(",")]
        with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as server:
            if smtp_starttls or smtp_user:
                try:
                    server.ehlo()
                    server.starttls()
                    server.ehlo()
                except smtplib.SMTPException as e_tls:
                    if smtp_starttls:
                        raise
                    detail_logger.debug(f"STARTTLS nicht verfügbar, fahre unverschlüsselt fort: {e_tls}")
            if smtp_user:
                server.login(smtp_user, smtp_pass or "")
            server.sendmail(msg["From"], recipient_list, msg.as_string())
        print(f"  ✓  Zusammenfassung gesendet an: {to_addr}")
    except Exception as e:
        print(f"  ⚠  E-Mail-Versand fehlgeschlagen: {e}")


def write_json_summary(
    stats: dict,
    info_counters: dict,
    start_dir: str,
    duration,
    mode_distribution: dict,
) -> Optional[str]:
    if not summary_json_path_global:
        return None
    try:
        summary = {
            "timestamp":                 datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "script_version":            SCRIPT_VERSION,
            "start_dir":                 start_dir,
            "duration_seconds":          int(duration.total_seconds()),
            "duration_human":            str(duration).split('.')[0],
            "stats":                     dict(stats),
            "info_counters":             dict(info_counters),
            "original_mode_distribution": {str(k): v for k, v in mode_distribution.items()},
        }
        with open(summary_json_path_global, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        return summary_json_path_global
    except Exception as e:
        detail_logger.warning(f"JSON-Zusammenfassung fehlgeschlagen: {e}")
        return None


# ==================================================================
# Auto-Resume (interaktiver Modus)
# ==================================================================

def get_auto_resume_path(start_dir: str, log_dir: str) -> str:
    key = hashlib.md5(
        os.path.normcase(os.path.abspath(start_dir)).encode("utf-8")
    ).hexdigest()[:12]
    return os.path.join(log_dir, f"{LOG_BASENAME}_resume_{key}.txt")


def _resume_has_entries(resume_path: str) -> int:
    if not os.path.exists(resume_path):
        return 0
    try:
        with open(resume_path, encoding="utf-8-sig") as fh:
            return sum(1 for l in fh if l.strip())
    except Exception:
        return 0


# ==================================================================
# Datei-Generator
# ==================================================================

def _is_reparse_point(path: str) -> bool:
    # Junctions/Mount-Points erkennen: os.walk(followlinks=False) folgt
    # zwar keinen echten Symlinks, traversiert unter Windows aber
    # Junctions als normale Ordner -> Zyklus-Gefahr (Endlosschleife).
    try:
        attrs = os.stat(path, follow_symlinks=False).st_file_attributes
        return bool(attrs & FILE_ATTRIBUTE_REPARSE_POINT)
    except (OSError, AttributeError):
        return False


def _path_matches_exclude(root_lower: str, exclude_patterns: list) -> bool:
    # Grenzanker-Matching auf Ordnerebene: Das Muster muss eine komplette
    # Pfadkomponente (oder eine zusammenhaengende Komponentenfolge) sein.
    # Ein reines Substring-Matching wie 'pat in root_lower' wuerde
    # faelschlich quer ueber Komponentengrenzen treffen (z.B. '--exclude-dir
    # alt' wuerde 'Verwaltung' ausschliessen). Durch das Einrahmen mit os.sep
    # auf beiden Seiten matcht 'alt' nur eine echte Komponente 'alt',
    # waehrend Mehr-Segment-Muster ('Archiv\\Alt') weiterhin funktionieren.
    if not exclude_patterns:
        return False
    sep = os.sep
    anchored = sep + root_lower.strip(sep) + sep
    for pat in exclude_patterns:
        p = pat.strip().strip("\\/").replace("/", sep).lower()
        if p and (sep + p + sep) in anchored:
            return True
    return False


def file_generator(directory: str, exclude_patterns: list = None):
    extensions = {".docx", ".docm", ".dotx", ".dotm", ".doc", ".dot"}
    exclude_patterns = [p.lower() for p in (exclude_patterns or [])]

    safe_dir = prepare_long_path(directory)
    # Eigenen Arbeitsordner ausnehmen: liegt das Ziel z.B. auf Documents,
    # wuerde der Generator sonst die eigenen convert_*/longpath_*-Stage-
    # Dateien einsammeln, waehrend Word sie gerade schreibt.
    temp_nc = os.path.normcase(os.path.abspath(TEMP_PROCESS_PATH))

    for root, dirs, files in os.walk(safe_dir):
        root_lower = root.lower()
        if _path_matches_exclude(root_lower, exclude_patterns):
            dirs[:] = []
            continue
        # Papierkorb/Systemordner/NAS-Snapshots, eigenen Temp-Ordner und
        # Junctions in-place aus der Traversierung entfernen.
        pruned = []
        for d in dirs:
            if d.lower() in EXCLUDE_DIR_NAMES:
                continue
            full_d = os.path.join(root, d)
            d_nc = os.path.normcase(os.path.abspath(_strip_lp(full_d)))
            if d_nc == temp_nc or d_nc.startswith(temp_nc + os.sep):
                continue
            if _is_reparse_point(full_d):
                continue
            pruned.append(d)
        dirs[:] = pruned
        for f in files:
            nl = f.lower()
            if nl.startswith("~$") or nl.startswith("._"):
                continue
            if any(nl.endswith(ext) for ext in extensions):
                yield os.path.join(root, f)


# ==================================================================
# List-Only-Modus (kein Word-Start, nur Pfade auflisten)
# ==================================================================

def run_list_only(start_dir: str, log_dir: str, exclude_patterns: list) -> str:
    list_path = os.path.join(log_dir, f"{LOG_BASENAME}_list_only_{run_timestamp_global}.txt")
    os.makedirs(log_dir, exist_ok=True)

    print(f"\nList-Only: Sammle Dateien unter '{start_dir}' ...")
    count = 0
    total_size = 0
    ext_counter = defaultdict(int)

    with open(list_path, "w", encoding="utf-8") as f:
        f.write(f"# Word-Updater {SCRIPT_VERSION} – List-Only-Lauf\n")
        f.write(f"# Zeitpunkt: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"# Verzeichnis: {start_dir}\n")
        if exclude_patterns:
            f.write(f"# Excludes: {', '.join(exclude_patterns)}\n")
        f.write("# Spalten: Pfad | Größe (Bytes) | Extension\n")
        f.write("# " + "=" * 78 + "\n")
        for full_path in file_generator(start_dir, exclude_patterns=exclude_patterns):
            size = safe_getsize(full_path)
            ext  = os.path.splitext(full_path)[1].lower()
            f.write(f"{full_path} | {size} | {ext}\n")
            count        += 1
            total_size   += size
            ext_counter[ext] += 1
            if count % 500 == 0:
                print(f"  ... {count} Dateien gesammelt")

        f.write("# " + "=" * 78 + "\n")
        f.write(f"# Gesamt: {count} Dateien, {total_size:,} Bytes\n")
        for ext in sorted(ext_counter.keys()):
            f.write(f"# {ext}: {ext_counter[ext]}\n")

    print(f"\n  ✓  {count} Dateien aufgelistet ({total_size:,} Bytes)")
    print(f"  Datei: {list_path}")
    return list_path


# ==================================================================
# Kern-Logik: Kompatibilitätsmodus beenden / Konvertierung
# ==================================================================

def end_compatibility_mode(
    file_path: str,
    word_app: win32com.client.CDispatch,
    max_compat_mode: int,
    pbar: tqdm,
    passwords: list,
    info_counters: dict = None,
    mode_distribution: dict = None,
    target_app_version: float = TARGET_APP_VERSION_DEFAULT,
    dry_run: bool = False,
) -> tuple:
    pbar.write(f"Prüfe: {os.path.basename(file_path)}")
    detail_logger.info(f"=== Starte: {file_path} ===")

    file_start_time = time.time()

    original_path  = file_path
    is_temp_copy   = False
    backup_path    = None
    temp_save_path = None
    target_path    = None
    temp_path      = None
    doc            = None
    # Reservierter Zielname (0-Byte-Platzhalter) und ob er schon durch die
    # echte Datei ersetzt wurde - siehe Freigabe im finally.
    target_was_reserved = False
    converted_ok        = False
    ext            = os.path.splitext(file_path)[1].lower()
    is_template    = ext in (".dot", ".dotx", ".dotm")
    new_ext        = ext

    # --- Pre-flight: Datei aktuell durch anderen Prozess gesperrt? ---
    if is_locked_by_other(original_path):
        pbar.write("  → ÜBERSPRUNGEN: Datei wird gerade von einem anderen Benutzer bearbeitet.")
        detail_logger.info(f"Skip (Sperrdatei/Sharing-Violation): {original_path}")
        if info_counters is not None:
            info_counters["LOCKED_SKIPPED"] += 1
        return "SKIPPED_LOCKED", original_path

    # --- Long-Path-Behandlung ---
    if len(file_path) > MAX_PATH_LEN:
        try:
            temp_name = f"longpath_{uuid.uuid4().hex}{ext}"
            temp_path = os.path.join(TEMP_PROCESS_PATH, temp_name)
            if not robust_copy(file_path, temp_path):
                raise Exception("Kopieren fehlgeschlagen")
            if not verify_file(temp_path):
                raise Exception("Temp-Kopie leer/ungültig")
            wait_for_file_available(temp_path)
            file_path = temp_path
            is_temp_copy = True
            detail_logger.debug(f"Temp-Kopie: {temp_path}")
        except Exception as e:
            log_error(original_path, Exception(f"Long-Path-Kopie fehlgeschlagen: {e}"))
            return "ERROR", original_path

    if not is_temp_copy:
        if file_path.startswith("\\\\?\\UNC\\"):
            file_path = "\\\\" + file_path[8:]
        elif file_path.startswith("\\\\?\\"):
            file_path = file_path[4:]

    # --- Original-Zeitstempel merken ---
    orig_times = None
    try:
        import win32file
        FILE_READ_ATTRIBUTES = 0x0080
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
    except Exception as _e:
        detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")

    # NTFS-Sicherheitsinfo (Owner/Group/DACL) des Originals sichern -
    # wird nach der Ersetzung auf die neue Datei uebertragen.
    orig_sd = _get_security_descriptor(original_path)

    try:
        is_encrypted         = False
        needs_password_check = False
        is_binary            = ext in (".doc", ".dot")

        # --- AppVersion aus docProps/app.xml lesen (vor COM-Open!) ---
        original_app_version = get_ooxml_app_version(file_path) if not is_binary else None

        # --- Verschlüsselungs-Check ---
        if ext in (".docx", ".docm", ".dotx", ".dotm"):
            try:
                with open(file_path, "rb") as f_check:
                    office_file = msoffcrypto.OfficeFile(f_check)
                    is_encrypted = office_file.is_encrypted()
            except PermissionError:
                detail_logger.warning(
                    f"Datei gesperrt, Krypto-Check nicht möglich – übersprungen: {original_path}")
                file_logger.error(f"Datei: {original_path}\n  -> Übersprungen (gesperrt/kein Zugriff)\n")
                if info_counters is not None:
                    info_counters["LOCKED_SKIPPED"] = info_counters.get("LOCKED_SKIPPED", 0) + 1
                return "SKIPPED_LOCKED", original_path
            except Exception as e_crypt:
                detail_logger.warning(
                    f"msoffcrypto-Check fehlgeschlagen für {original_path}: {e_crypt}")
                try:
                    import zipfile
                    if not zipfile.is_zipfile(file_path):
                        is_encrypted = True
                    else:
                        with zipfile.ZipFile(file_path) as zf:
                            names = zf.namelist()
                            if "EncryptedPackage" in names or "EncryptionInfo" in names:
                                is_encrypted = True
                            elif "word/document.xml" not in names and \
                                 "[Content_Types].xml" in names:
                                is_encrypted = True
                            else:
                                is_encrypted = False
                except Exception as e_zip:
                    detail_logger.warning(
                        f"ZIP-Fallback ebenfalls fehlgeschlagen: {e_zip}")
                    is_encrypted = True
        elif ext in (".doc", ".dot"):
            needs_password_check = True
            try:
                with open(file_path, "rb") as f_ole:
                    ole_header = f_ole.read(8)
                    if ole_header == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                        content = f_ole.read(65528)
                        if b'EncryptedPackage' in content or b'EncryptionInfo' in content:
                            is_encrypted = True
                            needs_password_check = False
                        else:
                            # Klassische .doc-Verschluesselung (XOR/RC4)
                            # traegt keine der obigen Stream-Namen, sondern
                            # setzt das fEncrypted-Bit im FIB (Offset +11
                            # ab wIdent 0xA5EC, Maske 0x01). Der WordDocument-
                            # Stream beginnt an einer Sektorgrenze.
                            full = ole_header + content
                            for off in range(512, len(full) - 12, 512):
                                if full[off] == 0xEC and full[off + 1] == 0xA5:
                                    if full[off + 11] & 0x01:
                                        is_encrypted = True
                                        needs_password_check = False
                                    break
            except Exception as e_ole:
                detail_logger.debug(f"OLE-Header-Check fehlgeschlagen: {e_ole}")

        # --- Passwort-Kandidaten bestimmen ---
        if not is_encrypted and not needs_password_check:
            passwords_to_try = [""]
        elif is_encrypted and not passwords:
            pbar.write("  → ÜBERSPRUNGEN: Kennwortgeschützt (kein Passwort hinterlegt).")
            detail_logger.warning(f"Übersprungen (Kennwort): {original_path}")
            file_logger.error(f"Datei: {original_path}\n  -> Übersprungen (Passwortgeschützt)\n")
            if info_counters is not None:
                info_counters["ENCRYPTED_SKIPPED"] += 1
            return "SKIPPED", original_path
        elif needs_password_check:
            passwords_to_try = [""] + passwords
        else:
            passwords_to_try = passwords

        # --- Word-Interaktion runterfahren ---
        _prev_display_alerts = None
        try:
            _prev_display_alerts = word_app.DisplayAlerts
            word_app.DisplayAlerts = COM_FALSE
        except Exception as _e:
            detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")
        try:
            word_app.Interactive = COM_FALSE
        except Exception as _e:
            detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")

        # --- Öffnen mit Retry und Passwort-Liste ---
        doc_opened       = False
        timeout_occurred = False
        last_exception   = None

        for open_attempt in range(MAX_RETRIES):
            for pw in passwords_to_try:
                try:
                    doc = safe_word_open(
                        word_app, file_path, pw, is_binary,
                        timeout=OPEN_TIMEOUT, word_pid=word_pid_global)
                    doc_opened = True
                    detail_logger.debug(f"Geöffnet (Passwort: {'[LEER]' if pw == '' else '***'})")
                    break
                except TimeoutError:
                    timeout_occurred = True
                    break
                except pythoncom.com_error as open_err:
                    last_exception = open_err
                    hr = getattr(open_err, "hresult", None)
                    if hr in HR_OPEN_ESCALATE:
                        raise
                    if word_app.ProtectedViewWindows.Count > 0:
                        target_name = os.path.basename(file_path).lower()
                        pvw = None
                        pv_count = word_app.ProtectedViewWindows.Count
                        for pv_i in range(1, pv_count + 1):
                            try:
                                cand = word_app.ProtectedViewWindows(pv_i)
                                try:
                                    src = os.path.basename(str(cand.SourceName or "")).lower()
                                except Exception:
                                    src = ""
                                if src == target_name:
                                    pvw = cand
                                    break
                            except Exception:
                                continue
                        if pvw is not None:
                            try:
                                doc = pvw.Edit()
                                doc_opened = True
                                break
                            except Exception:
                                try:
                                    pvw.Close()
                                except Exception as _e:
                                    detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")
                        else:
                            detail_logger.warning(
                                f"Kein passendes ProtectedView-Fenster für {original_path} "
                                f"(Count={pv_count}) – kein blinder Zugriff auf Index 1.")
                except Exception as e:
                    last_exception = e

            if doc_opened or timeout_occurred:
                break
            if open_attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)

        try:
            word_app.Interactive = COM_TRUE
        except Exception as _e:
            detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")
        try:
            if _prev_display_alerts is not None:
                word_app.DisplayAlerts = _prev_display_alerts
        except Exception as _e:
            detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")

        if timeout_occurred:
            pbar.write("  ✗  FEHLER: Kennwort-Dialog blockiert (Timeout) – Word-Prozess beendet.")
            detail_logger.warning(f"Timeout (Word-Prozess getötet): {original_path}")
            file_logger.error(f"Datei: {original_path}\n  -> Timeout (Passwort-Prompt, Word-Prozess getötet)\n")
            return "ERROR", original_path

        if not doc_opened or doc is None:
            # Zwischen DAUERHAFT und VORUEBERGEHEND unterscheiden. Frueher
            # ging jeder gescheiterte Open als 'SKIPPED' zurueck, und der
            # Aufrufer vermerkt 'SKIPPED' in der Resume-Datei - eine Datei,
            # die nur gerade gesperrt war oder deren Netzpfad kurz weg war,
            # galt danach als dauerhaft erledigt und wurde in JEDEM Folgelauf
            # uebersprungen. Das widerspricht dem Kommentar am Resume-Filter,
            # der ausdruecklich nur dauerhaft erledigte Dateien vorsieht.
            # Ein fehlendes Passwort ist dauerhaft (SKIPPED), ein
            # Zugriffs-/Sperrfehler nicht (SKIPPED_TRANSIENT).
            if needs_password_check and passwords and last_exception:
                pbar.write("  → ÜBERSPRUNGEN: Passwort erforderlich (keines der hinterlegten Passwörter passt).")
                detail_logger.warning(f"Übersprungen (Passwort): {original_path}")
                if info_counters is not None:
                    info_counters["PROTECTED_SKIPPED"] += 1
                return "SKIPPED", original_path

            pbar.write("  → ÜBERSPRUNGEN: Zugriff verweigert (wird beim nächsten Lauf erneut versucht).")
            detail_logger.warning(
                f"Übersprungen (kein Zugriff, voruebergehend): {original_path}")
            if info_counters is not None:
                info_counters["LOCKED_SKIPPED"] += 1
            return "SKIPPED_TRANSIENT", original_path

        # --- Kompatibilitätsmodus lesen ---
        original_mode = None
        try:
            original_mode = doc.CompatibilityMode
            if mode_distribution is not None and original_mode is not None:
                mode_distribution[original_mode] += 1
        except Exception as _e:
            detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")

        # --- Bearbeitungsschutz aufheben ---
        try:
            if doc.ProtectionType != -1:
                unprotect_succeeded = False
                try:
                    doc.Unprotect()
                    unprotect_succeeded = True
                except Exception:
                    for pw in passwords:
                        try:
                            doc.Unprotect(pw)
                            unprotect_succeeded = True
                            detail_logger.debug(
                                "Bearbeitungsschutz mit bekanntem Passwort entfernt.")
                            break
                        except Exception as _e:
                            detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")
                if not unprotect_succeeded:
                    detail_logger.warning(
                        f"Bearbeitungsschutz konnte nicht entfernt werden "
                        f"(kein passendes Passwort): {original_path}")
        except Exception as _e:
            detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")

        # --- Migration-First: doc.Convert() ---
        needs_convert_call = (
            not is_binary
            and (original_mode is None or original_mode < max_compat_mode)
        )
        if needs_convert_call:
            try:
                doc.Convert()
            except Exception as e_conv:
                detail_logger.debug(f"doc.Convert() fehlgeschlagen: {e_conv}")

        # --- AppVersion-Check: ist die Datei nur metadatenseitig veraltet? ---
        # Eine FEHLENDE AppVersion (None) bei einer OOXML-Datei bedeutet,
        # dass die Datei nicht von einer aktuellen Word-Engine geschrieben
        # wurde (Drittanbieter wie PDF-Konverter, python-docx, LibreOffice
        # oder beschaedigte docProps). Solche Dateien sollen neu serialisiert
        # werden - genau das ist der Zweck dieses Skripts und entspricht dem
        # Verhalten des Excel-Geschwisterskripts 3b (dort: "Keine
        # OOXML-Versionsinformationen -> Update"). Binaerformate (.doc/.dot)
        # sind separat ueber is_binary abgedeckt; fuer sie ist
        # original_app_version ohnehin None, was hier korrekt zu "veraltet"
        # fuehrt, die is_binary-First-Verzweigung der Statusmeldung aber
        # nicht stoert.
        app_version_outdated = (
            (not is_binary and original_app_version is None)
            or (original_app_version is not None
                and original_app_version < target_app_version)
        )

        # --- Speichern nötig? ---
        needs_save = (
            is_binary
            or (original_mode is None)
            or (original_mode < max_compat_mode)
            or app_version_outdated
            or (not doc.Saved)
        )

        if not needs_save:
            av_str = (f"AppVersion {original_app_version}"
                      if original_app_version is not None else "AppVersion -")
            pbar.write(f"  ✓  Bereits aktuell (Modus {original_mode}, {av_str})")
            doc.Close(SaveChanges=COM_FALSE)
            doc = None
            return "ALREADY_CURRENT", original_path

        # --- Statusmeldung ---
        if is_binary:
            pbar.write(f"  → Binäres {ext.upper()} erkannt. Konvertiere ...")
            if info_counters is not None:
                info_counters["BINARY_CONVERTED"] += 1
        elif original_mode is not None and original_mode < max_compat_mode:
            pbar.write(f"  → Modus {original_mode} < {max_compat_mode}. Aktualisiere ...")
            if info_counters is not None:
                info_counters["MODE_UPGRADED"] += 1
        elif original_mode is not None and original_mode >= max_compat_mode and app_version_outdated:
            _av_reason = (f"AppVersion {original_app_version} < {target_app_version}"
                          if original_app_version is not None
                          else "Keine AppVersion (Drittanbieter/beschädigt)")
            pbar.write(
                f"  → {_av_reason} "
                f"(Modus {original_mode} bleibt). Aktualisiere Metadaten ...")
            if info_counters is not None:
                info_counters["APPVERSION_UPDATED"] = info_counters.get("APPVERSION_UPDATED", 0) + 1
        elif original_mode is not None and original_mode >= max_compat_mode:
            pbar.write(
                f"  → Word-Auto-Reparatur erkannt (Modus {original_mode}, "
                f"Dokument wurde beim Oeffnen still modifiziert)")
            if info_counters is not None:
                info_counters["WORD_AUTO_REPAIRED"] += 1

        # --- Makro-Erkennung ---
        has_macros = False
        if ext in (".docm", ".dotm"):
            has_macros = True
        elif ext in (".docx", ".dotx"):
            has_macros = False
        else:
            try:
                has_macros = doc.HasVBProject
            except pythoncom.com_error as e_vba:
                has_macros = True
                try:
                    if e_vba.hresult in HR_VBA_NO_ACCESS:
                        detail_logger.warning(
                            f"VBA-Objektmodell-Zugriff verweigert (Trust Center) – "
                            f"nehme Makros an: {original_path}")
                    else:
                        detail_logger.warning(
                            f"HasVBProject COM-Fehler ({e_vba.hresult}) – "
                            f"nehme Makros an: {original_path}")
                except Exception:
                    detail_logger.warning(
                        f"HasVBProject nicht abfragbar – nehme Makros an: {original_path}")
            except Exception as e_vba2:
                has_macros = True
                detail_logger.warning(
                    f"HasVBProject unerwartet: {e_vba2} – nehme Makros an: {original_path}")

        if has_macros and info_counters is not None:
            info_counters["MACROS_FOUND"] += 1

        # Hinweis: Re-Serialisierung durch SaveAs2 schreibt die Datei
        # physisch neu. Eine evtl. vorhandene digitale Signatur des
        # VBA-Projekts (Makro-Code-Signing) wird dadurch ungueltig.
        # Nur als Info ins Detail-Log - hilft bei spaeterer Fehlersuche,
        # wenn ein signiertes Makro nach der Migration als "nicht signiert"
        # gemeldet wird.
        if has_macros:
            detail_logger.info(
                f"Makros enthalten – nach Re-Serialisierung sind bestehende "
                f"VBA-Projekt-Signaturen ungueltig: {original_path}")
            if info_counters is not None:
                info_counters["MACRO_SIGNATURE_INVALIDATED"] = (
                    info_counters.get("MACRO_SIGNATURE_INVALIDATED", 0) + 1)

        # --- DRY-RUN: hier ist Schluss ───────────────────────────────────
        # An diesem Punkt steht fest, dass die Datei aktualisiert/konvertiert
        # WUERDE (needs_save war True, sonst waere oben bereits
        # ALREADY_CURRENT zurueckgegeben worden). Im Probelauf wird NICHTS
        # gespeichert, verschoben oder geloescht - das Dokument wird nur
        # geschlossen. Das Original auf der Platte ist zu keinem Zeitpunkt
        # angefasst worden (geoeffnet wurde read-only bzw. aus der Temp-Kopie).
        if dry_run:
            pbar.write("  🔎 [PROBELAUF] Würde aktualisiert/konvertiert "
                       "(kein Speichern)")
            detail_logger.info(f"[DRY-RUN] WOULD_UPDATE: {original_path}")
            try:
                doc.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")
            doc = None
            return "WOULD_UPDATE", original_path

        # --- Zielformat und Erweiterung bestimmen ---
        if is_template:
            new_format = WD_FORMAT_DOTM if has_macros else WD_FORMAT_DOTX
            new_ext    = ".dotm"        if has_macros else ".dotx"
        else:
            new_format = WD_FORMAT_DOCM if has_macros else WD_FORMAT_DOCX
            new_ext    = ".docm"        if has_macros else ".docx"

        temp_save_name = f"convert_{uuid.uuid4().hex}{new_ext}"
        temp_save_path = os.path.join(TEMP_PROCESS_PATH, temp_save_name)

        # --- SaveAs2 mit Watchdog ---
        try:
            safe_word_saveas(
                word_app, doc, temp_save_path, new_format, max_compat_mode,
                timeout=SAVEAS_TIMEOUT, word_pid=word_pid_global)
        except Exception as e_save:
            err_msg = str(e_save).lower()
            if ("macro" in err_msg or "makro" in err_msg) and new_ext in (".docx", ".dotx"):
                new_ext       = ".docm" if not is_template else ".dotm"
                new_format    = WD_FORMAT_DOCM if not is_template else WD_FORMAT_DOTM
                safe_remove(temp_save_path)
                temp_save_path = os.path.splitext(temp_save_path)[0] + new_ext
                safe_word_saveas(
                    word_app, doc, temp_save_path, new_format, max_compat_mode,
                    timeout=SAVEAS_TIMEOUT, word_pid=word_pid_global)
            else:
                raise

        doc.Close(SaveChanges=COM_FALSE)
        doc = None
        time.sleep(0.3)

        if not verify_file(temp_save_path):
            raise Exception("Verifizierung der konvertierten Datei fehlgeschlagen")

        # --- AV-Wartelogik nach Schreiben in Temp-Ordner ---
        wait_for_file_available(temp_save_path)

        # --- Zielpfad bestimmen (Namenskollision bei Ext-Wechsel beachten) ---
        if new_ext != ext:
            target_path = os.path.splitext(original_path)[0] + new_ext
        else:
            target_path = original_path

        if new_ext != ext:
            # Den Wunschnamen IMMER atomar belegen, nicht nur wenn er schon
            # besetzt ist. Bisher lief bei freiem Namen 'pruefen und danach
            # benutzen' - genau das Zeitfenster, gegen das
            # reserve_unique_path laut eigenem Docstring eingefuehrt wurde:
            # ein parallel laufender Durchgang oder ein Nutzer, der gerade
            # speichert, kann den Namen dazwischen belegen, und das
            # anschliessende Schreiben ueberschreibt die fremde Datei. Der
            # Fall 'noch frei' ist dabei der HAEUFIGERE.
            # os.open(O_CREAT|O_EXCL) legt einen 0-Byte-Platzhalter an; das
            # IST die Reservierung. Der finally-Block gibt ihn wieder frei,
            # wenn es nicht zum Schreiben kommt.
            reserviert_direkt = False
            try:
                fd = os.open(_lp_for_reserve(target_path),
                             os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o666)
                os.close(fd)
                reserviert_direkt = True
                target_was_reserved = True
            except FileExistsError:
                pass
            except OSError as e_res:
                detail_logger.debug(
                    f"Direktreservierung nicht moeglich ({target_path}): {e_res}")

            if not reserviert_direkt:
                base, new_ext_part = os.path.splitext(target_path)
                reserved = reserve_unique_path(base, new_ext_part)
                if not reserved:
                    raise Exception(
                        "Kein freier Ausweichname! Abbruch zum Schutz fremder Dateien."
                    )
                target_path = reserved
                target_was_reserved = True
                pbar.write(
                    f"  ⚠  Ziel existiert bereits – speichere als "
                    f"'{os.path.basename(target_path)}'"
                )
                detail_logger.info(
                    f"Formatkonflikt gelöst: '{original_path}' → '{target_path}'"
                )

        # --- Backup der Zieldatei ---
        # Ein frisch reservierter Platzhalter ist 0 Byte gross und braucht
        # kein Backup - sonst wuerde eine leere Datei gesichert und ein
        # evtl. vorhandenes echtes .bak beiseite gelegt.
        if safe_exists(target_path) and not target_was_reserved:
            backup_path = f"{target_path}.bak"
            if len(backup_path) > MAX_PATH_LEN:
                backup_path = prepare_long_path(backup_path)
            if safe_exists(backup_path):
                # Ein zurueckgebliebenes .bak stammt aus einem harten
                # Abbruch eines frueheren Laufs und kann die einzige
                # intakte Kopie der Zieldatei sein (Crash mitten im
                # Cross-Volume-Move trunkiert das Ziel). Daher beiseite
                # legen statt loeschen.
                stale_bak = f"{target_path}.bak_{uuid.uuid4().hex[:6]}"
                try:
                    os.replace(long_path(backup_path), long_path(stale_bak))
                    detail_logger.warning(
                        f"Veraltetes Backup beiseite gelegt (nicht gelöscht): {stale_bak}")
                except Exception:
                    safe_remove(backup_path)
                    detail_logger.warning(f"Veraltetes Backup gelöscht: {backup_path}")
            if not robust_copy(target_path, backup_path):
                safe_remove(backup_path)
                backup_path = None
                raise Exception(
                    "Backup fehlgeschlagen. Abbruch zum Schutz der Zieldatei."
                )

        # --- Long-Path-Temp-Kopie aufräumen ---
        if is_temp_copy and safe_exists(file_path):
            safe_remove(file_path)
            is_temp_copy = False

        # --- Konvertierte Datei an den Zielort verschieben ---
        if not robust_move(temp_save_path, target_path):
            raise Exception("Verschieben der konvertierten Datei fehlgeschlagen")
        temp_save_path = None
        # Ab hier steht die echte Datei am Zielpfad - der reservierte
        # Platzhalter ist damit ersetzt und darf nicht freigegeben werden.
        converted_ok = True

        # ACL/Owner des Originals auf die neue Datei uebertragen
        # (Admin-Kontext: vollstaendig; Nutzer-Kontext: DACL).
        _apply_security_descriptor(target_path, orig_sd)

        # --- Zeitstempel wiederherstellen (VOR dem Löschen des Originals) ---
        if orig_times and safe_exists(target_path):
            import win32file
            FILE_WRITE_ATTRIBUTES = 0x0100
            for av_retry in range(2):
                try:
                    h_dst = win32file.CreateFile(
                        prepare_long_path(target_path),
                        FILE_WRITE_ATTRIBUTES,
                        win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE,
                        None, win32file.OPEN_EXISTING, 0, None
                    )
                    try:
                        win32file.SetFileTime(h_dst, orig_times[0], orig_times[1], orig_times[2])
                    finally:
                        h_dst.Close()
                    break
                except Exception as e_utime:
                    if av_retry == 1:
                        if not _utime_rueckfall(target_path, orig_times):
                            detail_logger.warning(f"Konnte Zeitstempel nicht wiederherstellen: {e_utime}")
                        else:
                            detail_logger.info(
                                "Zeitstempel ueber os.utime-Rueckfall gesetzt "
                                "(ohne Erstellungszeit).")
                    else:
                        time.sleep(0.5)

        # --- Alte Quelldatei löschen (bei Namenswechsel) – ERST JETZT ---
        if original_path.lower() != target_path.lower():
            if not safe_remove(original_path):
                safe_remove(target_path)
                raise Exception(
                    "Löschen der Originaldatei fehlgeschlagen (Rechte?). "
                    "Konvertierung rückgängig gemacht."
                )

        if backup_path and safe_exists(backup_path):
            safe_remove(backup_path)
            backup_path = None

        duration_sec = time.time() - file_start_time
        if duration_sec > PER_FILE_WARN_SECONDS:
            detail_logger.warning(f"Lange Verarbeitung ({duration_sec:.1f}s): {original_path}")
        detail_logger.info(f"Dauer: {duration_sec:.1f}s für {original_path}")

        # Nachweisliste: Pfadwechsel (z.B. .doc -> .docx, oder
        # Ausweichname bei Kollision) in die Konvertierungs-CSV.
        if original_path.lower() != target_path.lower():
            log_conversion(original_path, target_path)

        pbar.write(f"  ✓  Aktualisiert: '{os.path.basename(original_path)}'")
        detail_logger.info(f"Erfolgreich: {original_path}")
        return "UPDATED", target_path

    except Exception:
        pbar.write(f"  ✗  FEHLER: '{os.path.basename(original_path)}'")

        if temp_save_path and safe_exists(temp_save_path):
            safe_remove(temp_save_path)
            temp_save_path = None

        if backup_path and safe_exists(backup_path) and target_path:
            if robust_copy(backup_path, target_path):
                _apply_security_descriptor(target_path, orig_sd)
                safe_remove(backup_path)
                pbar.write("    → Backup wiederhergestellt")
            else:
                pbar.write(
                    "    ⚠ KRITISCH: Rollback fehlgeschlagen! "
                    "Backup-Datei aus Sicherheitsgründen bewahrt."
                )
        elif not backup_path and target_path and safe_exists(target_path) \
                and target_path.lower() != original_path.lower():
            safe_remove(target_path)
            pbar.write("    → Korruptes Dateifragment nach Abbruch sicher entfernt.")

        raise

    finally:
        # Reservierten Zielnamen freigeben, falls die Konvertierung nicht
        # bis zum Schreiben kam. Entfernt ausschliesslich 0-Byte-Dateien -
        # ein real geschriebenes Ziel wird nie angefasst.
        if target_was_reserved and not converted_ok:
            release_unique_path(target_path)
        try:
            word_app.Interactive = COM_TRUE
        except Exception as _e:
            detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")

        if doc is not None:
            try:
                doc.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")
            doc = None
        try:
            count = word_app.ProtectedViewWindows.Count
            for i in range(count, 0, -1):
                try:
                    word_app.ProtectedViewWindows(i).Close()
                except Exception as _e:
                    detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")
        except Exception as _e:
            detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")
        try:
            for i in range(word_app.Documents.Count, 0, -1):
                try:
                    word_app.Documents(i).Close(SaveChanges=COM_FALSE)
                except Exception as _e:
                    detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")
        except Exception as _e:
            detail_logger.debug(f"end_compatibility_mode: Exception verworfen: {_e!r}")
        _clear_recent_files(word_app)
        if temp_path and not is_temp_copy and safe_exists(temp_path):
            safe_remove(temp_path)
        if is_temp_copy and safe_exists(file_path):
            safe_remove(file_path)
        if temp_save_path and safe_exists(temp_save_path):
            safe_remove(temp_save_path)


# ==================================================================
# Retry-Wrapper
# ==================================================================

def process_file_with_retries(
    full_path: str,
    word_app: win32com.client.CDispatch,
    max_compat_mode: int,
    passwords: list,
    pbar: tqdm,
    info_counters: dict = None,
    mode_distribution: dict = None,
    target_app_version: float = TARGET_APP_VERSION_DEFAULT,
    dry_run: bool = False,
) -> tuple:
    result, final_path = "ERROR", full_path
    for attempt in range(MAX_RETRIES):
        # Die Zaehler je Versuch auf einer KOPIE fuehren und erst beim
        # erfolgreichen Versuch uebernehmen. Frueher bekam
        # end_compatibility_mode die echten Dicts: saemtliche Info-Zaehler
        # werden dort INNERHALB der Funktion gesetzt, ein gescheiterter
        # Versuch hinterliess sie also - und der Wiederholungsversuch zaehlte
        # dieselbe Datei erneut. Bei MAX_RETRIES=3 stand am Ende bis zum
        # Dreifachen in der Auswertung, obwohl die Datei einmal verarbeitet
        # wurde. Dasselbe gilt fuer die Modus-Verteilung.
        versuch_counters = dict(info_counters) if info_counters is not None else None
        versuch_modes    = mode_distribution.copy() if mode_distribution is not None else None
        try:
            result, final_path = end_compatibility_mode(
                full_path, word_app, max_compat_mode, pbar, passwords,
                versuch_counters, versuch_modes,
                target_app_version=target_app_version,
                dry_run=dry_run)
            if info_counters is not None:
                info_counters.clear()
                info_counters.update(versuch_counters)
            if mode_distribution is not None:
                mode_distribution.clear()
                mode_distribution.update(versuch_modes)
            break
        except Exception as e:
            log_error(full_path, e)

            if is_permanent_error(e):
                pbar.write(
                    f"  ✗  Permanenter Fehler ({type(e).__name__}) – kein Retry.")
                break

            if attempt >= MAX_RETRIES - 1:
                break

            transient = is_transient_error(e)
            if isinstance(e, pythoncom.com_error):
                pbar.write(
                    f"  → COM-Fehler (Versuch {attempt + 1}/{MAX_RETRIES}"
                    f"{', transient' if transient else ''}) – prüfe Word-Instanz ...")
            else:
                pbar.write(
                    f"  → Fehler (Versuch {attempt + 1}/{MAX_RETRIES}"
                    f"{', transient' if transient else ''}) – prüfe Word-Instanz ...")

        # --- Health-Check ---
        word_alive = False
        try:
            _ = word_app.Version
            word_alive = True
        except Exception as _e:
            detail_logger.debug(f"process_file_with_retries: Exception verworfen: {_e!r}")

        if not word_alive:
            pbar.write("  ↻  Word-Instanz nicht mehr erreichbar – Neustart ...")
            try:
                word_app, _ = _restart_word_engine(old_app=word_app, pbar=pbar)
            except Exception as fatal_e:
                pbar.write(
                    f"  ✕  Kritisch: Word konnte nicht neu gestartet "
                    f"werden: {fatal_e}")
                break
        else:
            time.sleep(RETRY_DELAY)

    try:
        _ = word_app.Version
    except Exception:
        pbar.write("  ↻  Word-Instanz nach Verarbeitung nicht erreichbar – Neustart ...")
        try:
            word_app, _ = _restart_word_engine(old_app=word_app, pbar=pbar)
        except Exception as fatal_e:
            pbar.write(
                f"  ✕  Word-Neustart vor naechster Datei fehlgeschlagen: {fatal_e}")

    return result, final_path, word_app


# ==================================================================
# Verzeichnis-Verarbeitung
# ==================================================================

def process_directory(
    directory: str,
    word_app: win32com.client.CDispatch,
    max_compat_mode: int,
    passwords: list,
    count_files_first: bool,
    resume_file: Optional[str] = None,
    exclude_patterns: list = None,
    stop_on_error: bool = False,
    show_progress: bool = True,
    target_app_version: float = TARGET_APP_VERSION_DEFAULT,
    dry_run: bool = False,
) -> tuple:
    os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)

    gen = file_generator(directory, exclude_patterns=exclude_patterns)

    if count_files_first:
        print("\nIndiziere Dateien (kann einige Minuten dauern) ...")
        files_list = list(gen)
        total      = len(files_list)
        iterable   = files_list
        bar_fmt    = "{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]"
        print(f"{total} Dateien gefunden. Starte Verarbeitung ...\n")
    else:
        iterable = gen
        total    = None
        bar_fmt  = "{desc}: {n_fmt} Dateien [{elapsed}, {rate_fmt}]"
        print("\nStarte direkte Verarbeitung ...\n")

    print("-" * 66)

    stats             = defaultdict(int)
    info_counters     = defaultdict(int)
    mode_distribution = defaultdict(int)
    for k in ("UPDATED", "ALREADY_CURRENT", "SKIPPED", "ERROR", "RESUMED"):
        stats[k] = 0
    for k in ("MACROS_FOUND", "ENCRYPTED_SKIPPED", "PROTECTED_SKIPPED",
              "BINARY_CONVERTED", "MODE_UPGRADED", "APPVERSION_UPDATED",
              "WORD_AUTO_REPAIRED", "FILENAME_RENAMED", "LOCKED_SKIPPED"):
        info_counters[k] = 0

    processed_count = 0

    # --- Resume-Datei laden ---
    already_done: set = set()
    if resume_file and os.path.exists(resume_file):
        try:
            with open(resume_file, encoding="utf-8-sig") as rf:
                already_done = {l.strip().lower() for l in rf if l.strip()}
            print(f"  Resume: {len(already_done)} bereits verarbeitete Dateien geladen.")
        except Exception as e:
            print(f"  ⚠  Resume-Datei nicht lesbar: {e}")

    # --- Resume-Datei zum Anhängen öffnen (utf-8 ohne BOM) ---
    rf_handle = open(resume_file, "a", encoding="utf-8") if resume_file else None
    run_completed = False
    try:
        with tqdm(total=total, desc="Verarbeite", unit="Datei",
                  bar_format=bar_fmt, disable=not show_progress) as pbar:
            for full_path in iterable:
                if full_path.lower() in already_done:
                    stats["RESUMED"] += 1
                    pbar.update(1)
                    processed_count += 1
                    continue

                # Im Probelauf NICHT umbenennen (sanitize_file_on_disk wuerde
                # die Datei auf der Platte umbenennen) - der Dry-Run laesst
                # das Dateisystem voellig unangetastet.
                if not dry_run:
                    cleaned_path = sanitize_file_on_disk(
                        full_path, pbar=pbar, detail_log=detail_logger)
                    if cleaned_path != full_path:
                        info_counters["FILENAME_RENAMED"] += 1
                        full_path = cleaned_path

                result, final_path, word_app = process_file_with_retries(
                    full_path, word_app, max_compat_mode, passwords, pbar,
                    info_counters, mode_distribution,
                    target_app_version=target_app_version,
                    dry_run=dry_run)
                # Unter-Status (z.B. SKIPPED_LOCKED) auf den Basis-Status
                # fuer die Statistik abbilden.
                base_result = "SKIPPED" if result.startswith("SKIPPED") else result
                stats[base_result] = stats.get(base_result, 0) + 1

                # Resume nur fuer DAUERHAFT erledigte Dateien fortschreiben.
                # Gesperrte Dateien (SKIPPED_LOCKED) sind transient – sie
                # wuerden sonst nach einem Neustart fuer immer uebersprungen,
                # obwohl der sperrende Nutzer sie laengst geschlossen hat.
                resume_eligible = result in (
                    "UPDATED", "ALREADY_CURRENT", "SKIPPED")
                if resume_eligible:
                    already_done.add(full_path.lower())
                    if final_path:
                        already_done.add(final_path.lower())
                    if rf_handle:
                        try:
                            rf_handle.write(full_path + "\n")
                            if final_path and final_path != full_path:
                                rf_handle.write(final_path + "\n")
                            # Flush NACH Inkrement pruefen (siehe unten),
                            # daher hier processed_count+1 verwenden.
                            if (processed_count + 1) % 50 == 0:
                                rf_handle.flush()
                                os.fsync(rf_handle.fileno())
                        except Exception as _e:
                            detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")

                pbar.update(1)
                processed_count += 1

                if stop_on_error and result == "ERROR":
                    pbar.write("  ✕  --stop-on-error aktiv: Verarbeitung wird abgebrochen.")
                    break

                if processed_count % WORD_RESTART_INTERVAL == 0:
                    pbar.write(
                        f"  ↻  Speicher-Reset: Starte Word nach "
                        f"{WORD_RESTART_INTERVAL} Dateien neu..."
                    )
                    try:
                        word_app, _ = _restart_word_engine(
                            old_app=word_app, pbar=pbar)
                    except Exception as e_restart:
                        pbar.write(f"  ✗  Fehler beim Word-Neustart: {e_restart}")
            else:
                # for-else: Schleife lief ohne break vollstaendig durch.
                run_completed = True

    finally:
        if rf_handle:
            try:
                rf_handle.flush()
                rf_handle.close()
            except Exception as _e:
                detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
    detail_logger.info(f"Verarbeitung abgeschlossen: {dict(stats)}")
    return stats, info_counters, mode_distribution, run_completed


# ==================================================================
# Einstiegspunkt
# ==================================================================

if __name__ == "__main__":
    check_required_modules()

    # Fremde Word-Sitzungen erfassen, BEVOR eine eigene COM-Instanz
    # entsteht - danach waere die eigene nicht mehr unterscheidbar.
    snapshot_foreign_word_pids()

    start_time = datetime.now()

    print()
    print("╔══════════════════════════════════════════════════════════════╗")
    print("║              WORD KOMPATIBILITÄTSMODUS UPDATER               ║")
    print("╚══════════════════════════════════════════════════════════════╝")
    print(f"  Start: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # --- Kommandozeilenparameter ---
    parser = argparse.ArgumentParser(
        description="Word-Updater: DOC/DOCX Kompatibilitätsmodus aktualisieren",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Beispiele:\n"
            "  Interaktiv:   3a_doc_docx_auf_neueste_Version_aktualisieren.exe\n"
            "  Automatisch:  3a_... --dir \\\\server\\freigabe --auto-start\n"
            "  Mit Optionen: 3a_... --dir C:\\Daten --count-first\n"
            "  Auflisten:    3a_... --dir C:\\Daten --list-only\n"
            "  Mit Excludes: 3a_... --dir Q:\\ --exclude-dir Archiv --exclude-dir Temp\n"
            "\n"
            "WICHTIG – --auto-start:\n"
            "  Ausschließlich für automatisierte Dienste (Task Scheduler, SCCM, CI/CD).\n"
            "  Das Skript beendet im Automatikmodus ALLE Word-Prozesse des ausführenden\n"
            "  Benutzers ohne Rückfrage und ohne Vorwarnung. Ungespeicherte Änderungen\n"
            "  gehen unwiederbringlich verloren. Niemals interaktiv mit offenen\n"
            "  Word-Fenstern verwenden!"
        ),
    )
    parser.add_argument("--version", action="version",
        version=f"Word-Updater {SCRIPT_VERSION} ({SCRIPT_DATE})")
    parser.add_argument("--dir", metavar="PFAD",
        help="Zu verarbeitendes Verzeichnis (überspringt ask_directory)")
    parser.add_argument("--auto-start", action="store_true",
        help="Alle Bestätigungsabfragen überspringen und sofort starten")
    parser.add_argument("--password", metavar="PW", default=None,
        help="Standardpasswort zum Öffnen verschlüsselter Dateien")
    parser.add_argument("--count-first", action="store_true",
        help="Dateien vorab zählen (ETA-Anzeige, aber langsamerer Start)")
    parser.add_argument("--pwdfile", metavar="DATEI", default=None,
        help="Textdatei mit Passwörtern (ein Passwort pro Zeile)")
    parser.add_argument("--resume", metavar="DATEI", default=None,
        help="Resume-Datei: bereits fertige Dateien überspringen")
    parser.add_argument("--dry-run", action="store_true",
        help="Probelauf: Word startet, öffnet jede Datei und prüft "
             "CompatibilityMode/AppVersion/Passwort, SPEICHERT aber NICHTS. "
             "Liefert exakte Statistik (inkl. JSON), ohne eine Datei zu ändern.")
    parser.add_argument("--mailto", metavar="ADRESSE", default=None,
        help="E-Mail-Adresse für Laufbericht nach Abschluss")
    parser.add_argument("--mailfrom", metavar="ADRESSE", default=None,
        help="Absenderadresse für --mailto (Standard: word-updater@<smtp-host>)")
    parser.add_argument("--smtp", metavar="HOST", default="localhost",
        help="SMTP-Server für --mailto (Standard: localhost)")
    parser.add_argument("--smtp-port", metavar="PORT", type=int, default=25,
        help="SMTP-Port (Standard: 25; für STARTTLS meist 587)")
    parser.add_argument("--smtp-user", metavar="USER", default=None,
        help="SMTP-Benutzername (aktiviert Login + STARTTLS, für Office 365/Gmail)")
    parser.add_argument("--smtp-pass", metavar="PASS", default=None,
        help="SMTP-Passwort für --smtp-user. ACHTUNG: als Kommandozeilen-"
             "Argument in der Prozessliste sichtbar – besser per "
             "Umgebungsvariable SMTP_PASS setzen.")
    parser.add_argument("--smtp-starttls", action="store_true",
        help="STARTTLS erzwingen, auch ohne --smtp-user")
    parser.add_argument("--log-dir", metavar="PFAD", default=None,
        help="Verzeichnis für Log- und Summary-Dateien (Standard: Skript-Verzeichnis)")
    parser.add_argument("--exclude-dir", metavar="MUSTER", action="append", default=[],
        help="Teilpfad-Muster, das vom Durchlauf ausgeschlossen wird (mehrfach möglich)")
    parser.add_argument("--list-only", action="store_true",
        help="Nur Dateien auflisten (kein Word-Start, keine Änderungen)")
    parser.add_argument("--stop-on-error", action="store_true",
        help="Beim ersten Fehler sofort abbrechen (Standard: weiterlaufen)")
    args = parser.parse_args()

    auto_mode = args.auto_start

    # Vor dem ersten COM-Zugriff auf bereits laufende Sitzungen hinweisen.
    warn_running_word(auto_mode)

    # --- Log-Verzeichnis bestimmen und Logger konfigurieren ---
    if args.log_dir:
        log_dir = os.path.abspath(args.log_dir)
    else:
        if getattr(sys, "frozen", False):
            script_dir = os.path.dirname(os.path.abspath(sys.executable))
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        log_dir = script_dir if script_dir else os.getcwd()
    log_dir_global = log_dir
    log_file_path, detailed_log_file_path = _setup_logging(log_dir)

    # Restore-Privilegien (Admin-Kontext) fuer ACL/Owner-Erhalt aktivieren.
    _enable_restore_privileges()

    # --- Aufräumen verwaister Temp-Ordner aus früheren Läufen ---
    cleanup_orphaned_temp_dirs()

    # --- List-Only-Modus: kein Word, kein Lock, kein Trust-Center ---
    if args.list_only:
        if auto_mode and not args.dir:
            print("\n❌  Fehler: --dir fehlt im Automatikmodus.")
            sys.exit(1)
        if args.dir:
            start_dir = sanitize_path(args.dir)
            if not os.path.isdir(prepare_long_path(start_dir)):
                print(f"\n❌  Fehler: Verzeichnis existiert nicht oder ist nicht erreichbar: '{start_dir}'")
                sys.exit(1)
        else:
            start_dir = ask_directory()
        run_list_only(start_dir, log_dir, args.exclude_dir)
        sys.exit(0)

    # --- Vorab-Check der Trust-Center-Einstellungen ---
    if not auto_mode:
        docs_folder = os.path.join(os.path.expanduser("~"), "Documents")
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
        print("     ☑ Vertrauenswürdige Speicherorte im Netzwerk zulassen")
        print("     Folgende Speicherorte hinzufügen:")
        print(f"     • {docs_folder} (temporärer Arbeitsordner des Skripts)")
        print(r"     • Q:\ oder \\server\dfs (Quelldateien)")
        print("     Jeweils: ☑ Unterordner ... sind ebenfalls vertrauenswürdig")
        print("!" * 66)

        if not ask_yes_no("\nWurden diese Einstellungen in Word vorgenommen?"):
            print("\n❌  Abbruch. Bitte konfigurieren Sie erst das Trust-Center in Word.")
            sys.exit(0)

    # --- Konfiguration: CLI oder interaktiv ---
    if auto_mode and not args.dir:
        print("\n❌  Fehler: --dir fehlt im Automatikmodus.")
        print("    Verwendung: 3a_... --dir \\\\server\\freigabe --auto-start")
        sys.exit(1)

    if args.dir:
        start_dir = sanitize_path(args.dir)
        if not os.path.isdir(prepare_long_path(start_dir)):
            print(f"\n❌  Fehler: Verzeichnis existiert nicht oder ist nicht erreichbar: '{start_dir}'")
            sys.exit(1)
    else:
        start_dir = ask_directory()

    if auto_mode:
        count_first   = args.count_first
        show_progress = True
        passwords     = [args.password] if args.password else []
        if args.pwdfile:
            passwords = load_pwdfile(args.pwdfile) + passwords
    else:
        # Interaktiv: erst Passwörter, dann 3-Optionen-Fortschrittsmenü.
        if args.pwdfile:
            passwords = load_pwdfile(args.pwdfile)
            if args.password:
                passwords = [args.password] + passwords
        elif args.password:
            passwords = [args.password]
        else:
            passwords = ask_passwords()

        if args.count_first:
            show_progress = True
            count_first   = True
        else:
            show_progress, count_first = ask_progress_mode()

    # --- Resume-Datei bestimmen ---
    # CLI --resume hat Vorrang und behaelt die bewusst persistente,
    # inkrementelle Semantik (geplante Tasks). Ohne --resume wird im
    # interaktiven Modus eine verzeichnisspezifische Resume-Datei
    # angeboten und nach vollstaendigem Lauf geloescht.
    resume_file       = args.resume
    auto_resume_owned = False
    # Im Probelauf wird KEINE Resume-Datei verwendet oder geschrieben -
    # ein Dry-Run soll den echten Lauf nicht praejudizieren.
    if args.dry_run:
        resume_file = None
    if not auto_mode and not args.dry_run:
        args.dry_run = ask_yes_no(
            "\nProbelauf (Dry-Run)? Es wird geprüft, aber NICHTS gespeichert.\n"
            "  (Empfohlen vor dem ersten Echt-Lauf auf einer neuen Ablage)"
        )

    if resume_file is None and not args.dry_run and not auto_mode:
        candidate = get_auto_resume_path(start_dir, log_dir)
        n_done = _resume_has_entries(candidate)
        if n_done > 0:
            print(f"\n♻  Resume-Datei eines früheren Laufs gefunden: "
                  f"{n_done} bereits verarbeitete Datei(en).")
            if ask_yes_no("  Lauf fortsetzen (J) oder von vorn beginnen (n)?",
                          default_yes=True):
                resume_file = candidate
            else:
                try:
                    os.remove(candidate)
                except Exception:
                    pass
                resume_file = candidate
        else:
            resume_file = candidate
        auto_resume_owned = True

    # --- COM initialisieren VOR Signal-Handlern ---
    pythoncom.CoInitialize()
    signal.signal(signal.SIGINT,  _signal_handler)
    signal.signal(signal.SIGTERM, _signal_handler)
    if hasattr(signal, "SIGBREAK"):
        signal.signal(signal.SIGBREAK, _signal_handler)

    # --- Einzelinstanz-Schutz ---
    if not _acquire_lock():
        print("\n❌  Fehler: Es läuft bereits eine Instanz dieses Skripts.")
        print("    Falls das nicht stimmt, Lock-Datei manuell löschen:")
        print(f"    {LOCK_FILE}")
        sys.exit(1)

    word_app          = None
    stats_result      = None
    info_result       = None
    mode_dist_result  = None
    max_compat        = 15

    try:
        word_app, word_pid_global = _init_word_app()
        word_app_global = word_app

        max_compat         = determine_max_compat(word_app)
        target_app_version = _resolve_target_app_version(word_app)

        office_ver = word_app.Version

        # --- Trust-Center Smoke-Test ---
        print("\nPrüfe COM-Subsystem und Trust-Center ...")
        smoke_ok, smoke_msg = test_trust_center_smoke(
            word_app, word_pid_global, timeout=25.0
        )
        if not smoke_ok:
            print()
            print("=" * 66)
            print("  ⚠  WORD-SMOKE-TEST FEHLGESCHLAGEN")
            print("=" * 66)
            print(f"  Grund: {smoke_msg}")
            print()
            print("  Mögliche Ursachen:")
            print("    • Dokumentenordner nicht als vertrauenswürdiger Speicherort eingetragen")
            print("    • Geschützte Ansicht für unsichere Speicherorte noch aktiv")
            print("    • Word/Office-Profil beschädigt oder fehlende Desktop-Ordner")
            print("      (siehe Trust-Center-Hinweis weiter oben)")
            print("=" * 66)
            file_logger.error(f"Trust-Center-Smoke-Test fehlgeschlagen: {smoke_msg}")

            if auto_mode:
                print("\n❌  Abbruch im AUTO-Modus nach fehlgeschlagenem Trust-Center-Test.")
                file_logger.error("Abbruch (auto_mode) nach fehlgeschlagenem Trust-Center-Test.")
                sys.exit(2)

            if not ask_yes_no("\nTrotzdem fortfahren? (Timeout-Risiko pro Datei!)"):
                print("Abgebrochen.")
                file_logger.error("Abbruch durch Benutzer nach Trust-Center-Warnung.")
                sys.exit(0)
            file_logger.error("Trust-Center-Warnung vom Benutzer ignoriert – Fortsetzung.")
            try:
                word_app.Quit(SaveChanges=COM_FALSE)
            except Exception:
                pass
            word_app          = None
            word_app_global   = None
            word_pid_global   = None
        else:
            detail_logger.info("Trust-Center-Smoke-Test erfolgreich.")
            print("  → Smoke-Test OK.")

        # --- Konfigurations-Zusammenfassung ---
        compat_lines = [
            (15, "Word 2013 / 2016 / 2019 / 2021 / 2024 / 365"),
            (14, "Word 2010"),
            (12, "Word 2007"),
            (11, "Word 2003"),
        ]
        target_label = next(
            (lbl for m, lbl in compat_lines if m == max_compat),
            f"CompatibilityMode {max_compat}",
        )

        appver_lines = [
            (16.0, "Office 2016 / 2019 / 2021 / 2024 / 365"),
            (15.0, "Office 2013"),
            (14.0, "Office 2010"),
            (12.0, "Office 2007"),
        ]
        target_appver_label = next(
            (lbl for v, lbl in appver_lines if v == target_app_version),
            f"AppVersion {target_app_version}",
        )

        print()
        print("=" * 66)
        print("  KONFIGURATION")
        print("=" * 66)
        print(f"  Verzeichnis:           {start_dir}")
        print("  Office-Versionen:      Office 2013 / 2016 / 2019 / 2021 / 2024 / 365 (DE/EN)")
        print(f"  Lokal installiert:     {office_ver}")
        print(f"  Ziel-Kompatmodus:      {max_compat} ({target_label})")
        print(f"  Ziel-AppVersion:       {target_app_version} ({target_appver_label})")
        print(f"  Modus:                 {'Vorab zählen (ETA)' if count_first else 'Generator (kein ETA)'}")
        print(f"  Fortschrittsbalken:    {'an' if show_progress else 'aus (inline)'}")
        print(f"  Log-Verzeichnis:       {log_dir}")
        print("  Passwortdateien:       Überspringen (als SKIPPED loggen)")
        pw_display = f"{len(passwords)} hinterlegt" if passwords else "keines (Überspringen)"
        print(f"  Passwörter:            {pw_display}")
        if resume_file:
            _resume_label = resume_file + ("  (automatisch)" if auto_resume_owned else "")
            print(f"  Resume-Datei:          {_resume_label}")
        if args.exclude_dir:
            print(f"  Ausgeschlossen:        {', '.join(args.exclude_dir)}")
        if args.stop_on_error:
            print("  Stop-on-Error:         AKTIV – Abbruch beim ersten Fehler")
        if args.mailto:
            print(f"  E-Mail-Bericht:        {args.mailto} via {args.smtp}")
        print(f"  Word-Neustart alle:    {WORD_RESTART_INTERVAL} Dateien (Speicherbereinigung)")
        if auto_mode:
            print("  Ausführungsmodus:      AUTOMATISCH (keine Bestätigung erforderlich)")
        print()
        print("  Verarbeitungslogik:")
        print("    • .doc/.dot  → immer konvertieren zu .docx/.dotx")
        print(f"    • CompatibilityMode < {max_compat} → auf aktuellen Modus aktualisieren")
        print(f"    • CompatibilityMode ≥ {max_compat} und AppVersion < {target_app_version}")
        print("      → Re-Save zur Aktualisierung der Metadaten (z. B. von Word 2013")
        print(f"        gespeicherte Datei wird neu mit AppVersion {target_app_version} geschrieben)")
        print(f"    • CompatibilityMode ≥ {max_compat} und AppVersion ≥ {target_app_version}")
        print("      → bereits aktuell, keine Änderung")
        print("    • Verschlüsselte Dateien werden mit Passwort-Liste probiert,")
        print("      bei Fehlschlag als SKIPPED geloggt (kein Abbruch)")
        print()
        print("  CompatibilityMode-Referenz (Word.Document.CompatibilityMode):")
        for mode, label in compat_lines:
            arrow = "  ← Ziel" if mode == max_compat else ""
            print(f"    {mode:>2} = {label}{arrow}")
        print()
        print("  AppVersion-Referenz (docProps/app.xml in .docx/.docm/.dotx/.dotm):")
        for v, label in appver_lines:
            arrow = "  ← Ziel" if v == target_app_version else ""
            print(f"    {v} = {label}{arrow}")
        print("=" * 66)

        if not auto_mode:
            print("\n⚠ WARNUNG: Alle offenen Word-Fenster werden OHNE SPEICHERN geschlossen!")
            if not ask_yes_no("\nJetzt starten?"):
                print("Abgebrochen.")
                sys.exit(0)

        # --- Frisches Word starten ---
        if word_app_global is not None:
            try:
                word_app_global.Quit(SaveChanges=COM_FALSE)
            except Exception:
                pass
            word_app_global = None
            word_pid_global = None

        _kill_orphaned_word()

        word_app, word_pid_global = _init_word_app()
        word_app_global = word_app

        if args.dry_run:
            print(f"\nPROBELAUF (Dry-Run) – es wird NICHTS gespeichert: '{start_dir}'")
        else:
            print(f"\nVerarbeite: '{start_dir}'")
        print("-" * 66)

        stats_result, info_result, mode_dist_result, run_completed = process_directory(
            start_dir, word_app, max_compat, passwords, count_first,
            resume_file=resume_file,
            exclude_patterns=args.exclude_dir,
            stop_on_error=args.stop_on_error,
            show_progress=show_progress,
            target_app_version=target_app_version,
            dry_run=args.dry_run,
        )

        # Auto-Resume-Datei nach vollstaendigem Lauf entfernen (eine vom
        # Nutzer via --resume uebergebene Datei bleibt unangetastet).
        if auto_resume_owned and run_completed and resume_file:
            try:
                if os.path.exists(resume_file):
                    os.remove(resume_file)
                    detail_logger.info("Auto-Resume-Datei nach vollständigem Lauf gelöscht.")
            except Exception:
                pass

    except Exception as e:
        print(f"\n❌ KRITISCHER FEHLER: {e}")
        log_error("GLOBAL", e)

    finally:
        if word_app_global is not None:
            # Profil-persistente Word-Options auf den Ausgangswert
            # zuruecksetzen, BEVOR Word beendet wird.
            _restore_word_options(word_app_global)
            try:
                word_app_global.Quit(SaveChanges=COM_FALSE)
            except Exception:
                pass
        try:
            time.sleep(0.5)
            _kill_orphaned_word()
            _cleanup_word_inetcache()
            _cleanup_user_recent()
        except Exception:
            pass
        if os.path.exists(TEMP_PROCESS_PATH):
            try:
                time.sleep(1)
                shutil.rmtree(TEMP_PROCESS_PATH)
            except OSError:
                pass
        cleanup_windows_temp()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        _release_lock()

    duration = datetime.now() - start_time

    print()
    print("=" * 66)
    print("  PROBELAUF ABGESCHLOSSEN (es wurde NICHTS geändert)" if args.dry_run
          else "  VERARBEITUNG ABGESCHLOSSEN")
    print("=" * 66)
    print(f"  Ende:  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Dauer: {str(duration).split('.')[0]}")
    print()

    if stats_result:
        total_sum = sum(stats_result.values())
        print("  STATISTIK:")
        if args.dry_run:
            print(f"    [✓] Würde aktualisiert/konv.: {stats_result.get('WOULD_UPDATE', 0)}")
        else:
            print(f"    [✓] Aktualisiert / Konv.:   {stats_result.get('UPDATED', 0)}")
        print(f"    [=] Bereits aktuell:        {stats_result.get('ALREADY_CURRENT', 0)}")
        print(f"    [→] Übersprungen:           {stats_result.get('SKIPPED', 0)}")
        print(f"    [⏩] Resume (übersprungen): {stats_result.get('RESUMED', 0)}")
        print(f"    [✗] Fehler:                 {stats_result.get('ERROR', 0)}")
        print(f"        GESAMT:                 {total_sum}")
        if info_result:
            print()
            print("  DETAILS:")
            print(f"    [↑] Mode-Upgrade:           {info_result.get('MODE_UPGRADED', 0)}")
            print(f"    [⇒] Binär-Konvertierung:   {info_result.get('BINARY_CONVERTED', 0)}")
            print(f"    [M] Mit Makros:            {info_result.get('MACROS_FOUND', 0)}")
            print(f"    [🔒] Verschlüsselt (skip): {info_result.get('ENCRYPTED_SKIPPED', 0)}")
            print(f"    [🛡 ] Geschützt (skip):    {info_result.get('PROTECTED_SKIPPED', 0)}")
            if info_result.get("LOCKED_SKIPPED", 0) > 0:
                print(f"    [🔐] In Nutzung (skip):    {info_result['LOCKED_SKIPPED']}")
            if info_result.get("WORD_AUTO_REPAIRED", 0) > 0:
                print(f"    [⚙] Word-Auto-Reparatur:   {info_result['WORD_AUTO_REPAIRED']}")
            if info_result.get("APPVERSION_UPDATED", 0) > 0:
                print(f"    [v] AppVersion aktualisiert: {info_result['APPVERSION_UPDATED']}")
            if info_result.get("FILENAME_RENAMED", 0) > 0:
                print(f"    [📝] Dateinamen bereinigt:  {info_result['FILENAME_RENAMED']}")
        if mode_dist_result:
            print()
            print("  URSPRÜNGLICHE COMPAT-MODI:")
            for mode in sorted(mode_dist_result.keys()):
                print(f"    Modus {mode}: {mode_dist_result[mode]}")
        file_logger.error(
            f"VERARBEITUNG ABGESCHLOSSEN | "
            f"Verzeichnis: {start_dir} | "
            f"Dauer: {str(duration).split('.')[0]} | "
            f"Aktualisiert: {stats_result.get('UPDATED', 0)} | "
            f"Aktuell: {stats_result.get('ALREADY_CURRENT', 0)} | "
            f"Übersprungen: {stats_result.get('SKIPPED', 0)} | "
            f"Fehler: {stats_result.get('ERROR', 0)} | "
            f"Gesamt: {total_sum}"
        )

        # --- JSON-Summary schreiben ---
        json_path = write_json_summary(
            stats_result, info_result, start_dir,
            duration, mode_dist_result
        )
        if json_path:
            print(f"  JSON-Summary: {json_path}")

    print()
    print(f"  Fehler-Log:  {os.path.abspath(log_file_path)}")
    print(f"  Detail-Log:  {os.path.abspath(detailed_log_file_path)}")
    if conversions_csv_path_global and os.path.exists(conversions_csv_path_global):
        print("  Konvertierungs-CSV (alte → neue Pfade):")
        print(f"               {os.path.abspath(conversions_csv_path_global)}")
    if (resume_file and auto_resume_owned and not auto_mode
            and os.path.exists(resume_file)):
        print(f"  Resume-Log:  {os.path.abspath(resume_file)}")
        print("               (erhalten – Neustart setzt dort fort)")
    print("=" * 66)

    if not auto_mode:
        docs_folder = os.path.join(os.path.expanduser("~"), "Documents")
        print()
        print("!" * 66)
        print("  HINWEIS: TRUST-CENTER EINSTELLUNGEN ZURÜCKSETZEN (falls gewünscht)")
        print("!" * 66)
        print("  Die für dieses Skript gesetzten Trust-Center-Optionen können nach")
        print("  Abschluss zurückgesetzt werden. Beachten: einige Einstellungen")
        print("  (z. B. VBA-Objektmodell-Zugriff) werden evtl. von anderen Tools")
        print("  ebenfalls benötigt – vorher prüfen, bevor abgeschaltet wird.")
        print()
        print("  Empfohlenes Sicherheitsprofil für den Normalbetrieb:")
        print("  • Geschützte Ansicht: alle 3 Häkchen wieder aktivieren")
        print("  • Makros:             'Alle Makros mit Benachrichtigung deaktivieren'")
        print(f"  • Vertrauenswürdige Speicherorte: '{docs_folder}' wieder entfernen,")
        print("    falls keine andere Automation diesen benötigt")
        print("!" * 66)

    # WICHTIG: logging.shutdown() NICHT vor send_summary_mail aufrufen.
    # send_summary_mail loggt im STARTTLS-Fallback und bei Versandfehlern -
    # nach einem Shutdown waeren die FileHandler geschlossen und diese
    # Diagnose-Eintraege gingen verloren (je nach Python-Version drohte
    # frueher sogar ein "I/O operation on closed file"). Der Shutdown
    # erfolgt daher erst ganz am Ende.
    if args.mailto and stats_result:
        _from = args.mailfrom or f"word-updater@{args.smtp}"
        # Passwort bevorzugt aus Umgebungsvariable (nicht in Prozessliste
        # sichtbar); --smtp-pass als Fallback.
        _smtp_pass = os.environ.get("SMTP_PASS") or args.smtp_pass
        send_summary_mail(
            args.mailto, _from, args.smtp, stats_result,
            duration, start_dir, os.path.abspath(log_file_path),
            info_counters=info_result,
            mode_distribution=mode_dist_result,
            smtp_port=args.smtp_port,
            smtp_user=args.smtp_user,
            smtp_pass=_smtp_pass,
            smtp_starttls=args.smtp_starttls,
        )

    logging.shutdown()

    print()
    _warte_auf_taste(auto_mode)
