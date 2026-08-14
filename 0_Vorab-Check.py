# ==================================================================
# VORABCHECK FÜR NOTWENDIGE PROGRAMME, PATHS und ADD-INS
# ==================================================================
# Datum: 11.06.2026
# Prüft auf einem Zielrechner, ob alle für die Word-/OCR-Skripte
# nötigen Programme, Python-Pakete, DLLs, Office-Settings und
# Umgebungsvariablen vorhanden sind. Fehlendes wird soweit
# möglich automatisch aus einer Deploy-Quelle (Standard:
# Q:\deploy) installiert bzw. kopiert.
#
# Two-Stage-Admin-Modell:
#   Stage 1 (User – ohne Adminrechte):
#     - Office-Check, Python, pip-Pakete, lokale Binaries
#     - HKCU-ENV-Variablen, Trust-Center-/AccessVBOM-Einstellungen
#     - Stage-Kopie aller Admin-Installer aus Deploy nach
#       C:\tmp_skripte\stage  (Q:\-Zugriff nicht mehr nötig)
#     - Optional UAC-Re-Launch nach Stage 2 + Wait auf Beendigung
#     - Danach erneute Tool-Erkennung + ENV-Setzung in HKCU
#   Stage 2 (Admin – per UAC-Re-Launch, --admin-stage <dir>):
#     - Installiert systemweite Tools ausschließlich aus
#       Stage-Kopie (kein Q:\-Zugriff erforderlich)
#     - Schreibt Ergebnis-Manifest, beendet sich danach
#
# Deploy-Quelle überschreiben (optional):
#   als CLI-Argument:   0_Vorab-Check.exe "D:\alternative\quelle"
#   als ENV-Variable:   set VORABCHECK_DEPLOY=D:\alternative\quelle
#   interaktiv:         beim Start bestätigen oder anderen Pfad eingeben
#
# Weitere Schalter:
#   --yes / -y     alle Rückfragen mit "ja" beantworten (unbeaufsichtigt)
#   --no-input     keine Rückfragen, jeweils sichere Vorgabe verwenden
#   --no-pause     kein "ENTER zum Beenden" am Schluss
#   --log <datei>  Logfile explizit setzen (Stage 2 erbt es von Stage 1)
#
# Exit-Codes:
#   0 = alles OK
#   1 = manuelle Aktionen nötig
#   2 = Fehler aufgetreten
#   3 = unerwarteter Abbruch
#
# PyInstaller-Build:
#   python -m PyInstaller --onefile --console --clean -y --name "0_Vorab-Check" --icon "python_icon.ico" --hidden-import pywintypes --hidden-import win32api --hidden-import win32com --hidden-import win32com.client --hidden-import win32process --hidden-import win32file --hidden-import winreg --collect-submodules win32com 0_Vorab-Check.py
#
#   WICHTIG: --uac-admin NICHT setzen – Stage 1 muss ohne
#   Adminrechte starten können (Two-Stage-Modell).
#
# ==================================================================

from datetime import datetime

import sys
import os
import re
import json
import time
import glob
import shutil
import fnmatch
import hashlib
import logging
import platform
import atexit
import tempfile
import threading
import subprocess

# ------------------------------------------------------------------
# Plattform-abhängige Imports
# ------------------------------------------------------------------

IS_WINDOWS = (os.name == "nt")

if IS_WINDOWS:
    import ctypes
    from ctypes import wintypes
    import winreg
else:
    ctypes = None
    wintypes = None
    winreg = None

# =============================================================================
# Konstanten
# =============================================================================

DEPLOY_SOURCE_DEFAULT = r"Q:\deploy"
DEPLOY_SOURCE         = DEPLOY_SOURCE_DEFAULT

OCR_DIR             = r"C:\OCR"
VERAPDF_DIR         = r"C:\OCR\verapdf"
TMP_DIR             = r"C:\tmp_skripte"
BIN_DIR             = os.path.join(TMP_DIR, "bin")
STAGE_DIR           = os.path.join(TMP_DIR, "stage")
LOG_FILE            = os.path.join(TMP_DIR, f"0_Vorab-Check_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
STAGE_MANIFEST_FILE = os.path.join(STAGE_DIR, "stage_manifest.json")
STAGE_RESULT_FILE   = os.path.join(STAGE_DIR, "stage_result.json")

ADMIN_STAGE_FLAG = "--admin-stage"
LOG_FLAG         = "--log"

# Flags ohne Wert / mit Wert (fuer die Argument-Auswertung)
BOOL_FLAGS  = {"--yes", "-y", "--no-input", "--no-pause"}
VALUE_FLAGS = {ADMIN_STAGE_FLAG, LOG_FLAG}

# Nur diese Dateiendungen werden in der Deploy-Quelle indiziert.
# Spart bei grossen Netzlaufwerken sehr viel Zeit. Erweitern, falls
# neue find_in_deploy()-Muster andere Endungen benoetigen.
DEPLOY_INDEX_EXTENSIONS = {
    ".exe", ".msi", ".msix", ".dll", ".bat", ".cmd",
    ".xml", ".traineddata", ".whl", ".zip",
}

OFFICE_PACKAGES = {
    "tqdm":             "tqdm",
    "psutil":           "psutil",
    "msoffcrypto-tool": "msoffcrypto",
}

OCR_PACKAGES = {
    "ocrmypdf":  "ocrmypdf",
    "pikepdf":   "pikepdf",
    "pymupdf":   "fitz",
    "psutil":    "psutil",
    "pypdfium2": "pypdfium2",
    "tqdm":      "tqdm",
    "watchdog":  "watchdog",
}

REQUIRED_TESS_LANGS = ["deu", "eng", "fra", "spa"]

PYTHON_INSTALLERS = {
    "manager_patterns": ["python-manager-*.msix"],
    "classic_patterns": ["python-*-amd64.exe", "python-3*.exe", "python-*.exe"],
    "classic_exclude":  ["embed", "webinstall"],
}

TOOL_INSTALLERS = {
    "tesseract": {
        "patterns":   ["tesseract-ocr-w64-setup-*.exe", "tesseract*setup*.exe"],
        "silent_exe": ["/S"],
        "silent_msi": None,
        "admin":      True,
    },
    "ghostscript": {
        "patterns":   ["gs*w64.exe", "gs*w32.exe", "ghostscript*.exe"],
        "silent_exe": ["/S"],
        "silent_msi": None,
        "admin":      True,
    },
    "java": {
        "patterns":   ["OpenJDK25*.msi", "OpenJDK*jdk*.msi", "Temurin*.msi",
                       "jdk*.msi", "jre*.msi", "jre-8*windows*.exe"],
        "silent_exe": ["/s"],
        "silent_msi": ["/quiet", "/norestart",
                       "ADDLOCAL=FeatureMain,FeatureEnvironment,FeatureJarFileRunWith,FeatureJavaHome"],
        "admin":      True,
    },
}

TOOL_BINARIES = {
    "jbig2.exe":    {"target_dir": OCR_DIR},
    "pngquant.exe": {"target_dir": OCR_DIR},
    "unpaper.exe":  {"target_dir": OCR_DIR},
}

DOWNLOAD_LINKS = {
    "python":      "https://www.python.org/downloads/ bzw. https://aka.ms/python (Python Install Manager im Windows Store)",
    "tesseract":   "https://github.com/UB-Mannheim/tesseract/wiki bzw. https://github.com/tesseract-ocr/tesseract/releases (tesseract-ocr-w64-setup-...exe)",
    "tessdata":    "https://github.com/tesseract-ocr/tessdata (deu.traineddata, eng.traineddata, fra.traineddata, spa.traineddata)",
    "ghostscript": "https://ghostscript.com/releases/gsdnld.html (gs...w64.exe)",
    "java":        "https://adoptium.net/ (OpenJDK 21+ MSI-Installer wählen)",
    "verapdf":     "https://software.verapdf.org/releases/ (verapdf-installer...zip)",
    "jbig2":       "https://github.com/pts/pdfsizeopt-jbig2/releases",
    "pngquant":    "https://pngquant.org/pngquant-windows.zip",
    "unpaper":     "https://github.com/Inc44/unpaper_windows",
    "leptonica":   "wird normalerweise mit Tesseract (UB-Mannheim) ausgeliefert. Quelle: http://www.leptonica.org/source/README.html",
    "office":      "Microsoft Office wird NICHT automatisch installiert – Setup über die IT-Abteilung anfordern.",
}

OFFICE_RELEASE_LABELS = {
    "ProPlus2019Volume":    "Office 2019 ProPlus (Volume)",
    "Standard2019Volume":   "Office 2019 Standard (Volume)",
    "ProPlus2021Volume":    "Office 2021 ProPlus (Volume)",
    "Standard2021Volume":   "Office 2021 Standard (Volume)",
    "ProPlus2024Volume":    "Office 2024 ProPlus (Volume)",
    "Standard2024Volume":   "Office 2024 Standard (Volume)",
    "O365ProPlusRetail":    "Microsoft 365 Apps for Enterprise",
    "O365BusinessRetail":   "Microsoft 365 Apps for Business",
    "O365HomePremRetail":   "Microsoft 365 (Home/Personal)",
}

OFFICE_TRUST_APPS = ("Word", "Excel", "PowerPoint")

# Erfolgs-Codes für Installer
MSI_SUCCESS_CODES = (0, 1641, 3010)
EXE_SUCCESS_CODES = (0, 1641, 3010)

# =============================================================================
# Globaler Zustand
# =============================================================================

STATE = {
    "created_dirs":       [],
    "installed_packages": [],
    "installed_tools":    [],
    "copied_files":       [],
    "env_set":            [],
    "path_added":         [],
    "warnings":           [],
    "errors":             [],
    "manual_actions":     [],
    "admin_pending":      [],
    "stage2_summary":     [],
}

PYTHON_EXE        = None
TOOLS_FOUND       = {}
DEPLOY_INDEX      = {}
DEPLOY_INDEX_DIRS = set()

RUN_MODE   = "stage1"
STAGE_ROOT = STAGE_DIR

# Interaktivitaet (per CLI steuerbar: --yes / --no-input / --no-pause)
ASSUME_YES = False
NO_INPUT   = False
NO_PAUSE   = False

DEPLOY_INDEX_EXCLUDES = {
    "__pycache__", ".git", ".svn", ".hg", "node_modules",
    "$recycle.bin", "system volume information",
}

# =============================================================================
# Einheitliches Subprocess-Ergebnis
# =============================================================================

class CmdResult:
    __slots__ = ("returncode", "stdout", "stderr", "timed_out", "not_found", "error")

    def __init__(self, returncode=-1, stdout="", stderr="",
                 timed_out=False, not_found=False, error=""):
        self.returncode = returncode
        self.stdout     = stdout or ""
        self.stderr     = stderr or ""
        self.timed_out  = timed_out
        self.not_found  = not_found
        self.error      = error

    @property
    def ok(self):
        return (not self.timed_out) and (not self.not_found) and self.returncode == 0

    @property
    def msi_ok(self):
        return (not self.timed_out) and (not self.not_found) and self.returncode in MSI_SUCCESS_CODES

    @property
    def exe_ok(self):
        return (not self.timed_out) and (not self.not_found) and self.returncode in EXE_SUCCESS_CODES

# =============================================================================
# Konsole und Logging
# =============================================================================

def setup_console():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception as _e:
        logging.debug(f"setup_console: Exception verworfen: {_e!r}")

def setup_logging():
    root = logging.getLogger()
    # DEBUG statt INFO: An zwoelf Stellen werden verschluckte Ausnahmen mit
    # logging.debug() protokolliert - genau die Information, mit der sich ein
    # 'es passiert einfach nichts' spaeter aufklaeren laesst. Mit INFO auf dem
    # Root-Logger wurden sie samt und sonders verworfen, die catch-Bloecke
    # waren also stumm. Das Level gehoert an den Handler, nicht an den Logger.
    root.setLevel(logging.DEBUG)
    root.handlers.clear()
    try:
        os.makedirs(os.path.dirname(LOG_FILE) or TMP_DIR, exist_ok=True)
        fh = logging.FileHandler(LOG_FILE, mode="a", encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        fh.setLevel(logging.DEBUG)
        root.addHandler(fh)
    except Exception as e:
        # Ohne NullHandler wuerde logging auf stderr ausweichen -> doppelte Ausgabe.
        root.addHandler(logging.NullHandler())
        print(f"⚠️ Logfile konnte nicht geöffnet werden: {e}")

def ask(prompt, default=False):
    """Ja/Nein-Abfrage. Respektiert --yes / --no-input und nicht-interaktive Konsolen."""
    if ASSUME_YES:
        return True
    if NO_INPUT or not sys.stdin or not sys.stdin.isatty():
        return default
    try:
        answer = input(prompt).strip().lower()
    except (EOFError, KeyboardInterrupt):
        return default
    if not answer:
        return default
    return answer in ("j", "ja", "y", "yes")

def pause(prompt):
    if NO_PAUSE or NO_INPUT or ASSUME_YES:
        return
    try:
        input(prompt)
    except Exception as _e:
        logging.debug(f"pause: Exception verworfen: {_e!r}")

def log_info(msg):
    print(msg)
    logging.info(msg)

def log_warn(msg):
    print(msg)
    logging.warning(msg)
    STATE["warnings"].append(msg)

def log_err(msg):
    print(msg)
    logging.error(msg)
    STATE["errors"].append(msg)

def log_download_hint(tool_key):
    link = DOWNLOAD_LINKS.get(tool_key)
    if link:
        log_warn(f"   👉 Download: {link}")
        if tool_key != "office":
            log_warn(f"   👉 Heruntergeladene Datei nach '{DEPLOY_SOURCE}' kopieren.")
            log_warn("   👉 Danach dieses Skript erneut ausführen.")

# =============================================================================
# System-Helfer
# =============================================================================

def is_admin():
    if not IS_WINDOWS:
        try:
            return os.geteuid() == 0
        except AttributeError:
            return False
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False

def get_current_username():
    return (os.environ.get("USERNAME") or os.environ.get("USER") or "").lower()

def get_startupinfo():
    if not IS_WINDOWS:
        return None
    try:
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = subprocess.SW_HIDE
        return si
    except Exception:
        return None

def _decode_bytes(b):
    if b is None:
        return ""
    if isinstance(b, str):
        return b
    for enc in ("utf-8", "cp1252", "cp850", "latin-1"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", errors="replace")

# Ausgabedateien, die sich noch nicht loeschen liessen. Ein vom Kind
# abgesetzter Enkel erbt das Ausgabe-Handle und haelt es offen, solange er
# laeuft - die Datei ist dann bis zu seinem Ende gesperrt. Statt sie liegen
# zu lassen, wird bei jedem weiteren Aufruf und beim Programmende erneut
# aufgeraeumt.
_PENDING_TEMP_DELETES = []

def _drop_temp_file(path):
    if not path:
        return
    try:
        os.remove(path)
    except FileNotFoundError:
        pass
    except Exception:
        _PENDING_TEMP_DELETES.append(path)

def _sweep_pending_temp_files():
    for path in list(_PENDING_TEMP_DELETES):
        try:
            os.remove(path)
            _PENDING_TEMP_DELETES.remove(path)
        except FileNotFoundError:
            _PENDING_TEMP_DELETES.remove(path)
        except Exception as _e:
            logging.debug(f"_sweep_pending_temp_files({path}): noch gesperrt: {_e!r}")

atexit.register(_sweep_pending_temp_files)

def _read_output_file(path):
    try:
        with open(path, "rb") as fh:
            return _decode_bytes(fh.read())
    except Exception as _e:
        logging.debug(f"_read_output_file({path}): Exception verworfen: {_e!r}")
        return ""

def _kill_process_tree(proc):
    """Den gesamten Prozessbaum beenden, nicht nur den direkten Kindprozess.

    Nach einem Timeout muss auch ein abgesetzter Enkel weg (Installer, der
    sich per 'start' oder ueber einen Java-Launcher selbstaendig gemacht
    hat). Sonst laeuft er unbeaufsichtigt weiter, waehrend das Skript ihn
    als abgebrochen protokolliert und den naechsten Installer startet -
    msiexec-Sperre und halbfertige Installationen sind die Folge.
    """
    if proc is None:
        return
    try:
        subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                       capture_output=True, timeout=30,
                       startupinfo=get_startupinfo())
    except Exception as _e:
        logging.debug(f"_kill_process_tree: taskkill fehlgeschlagen: {_e!r}")
    try:
        proc.kill()
    except Exception as _e:
        logging.debug(f"_kill_process_tree: kill verworfen: {_e!r}")
    try:
        proc.wait(timeout=10)
    except Exception as _e:
        logging.debug(f"_kill_process_tree: wait nach kill: {_e!r}")

def run_cmd(args, timeout=30, cwd=None, extra_env=None):
    """Programm ausfuehren und Ausgabe einsammeln - mit bindendem Timeout.

    Bewusst NICHT subprocess.run(capture_output=True): dessen timeout ist
    auf Windows nicht bindend. Laeuft die Zeit ab, toetet run() nur den
    DIREKTEN Kindprozess und ruft danach communicate() OHNE timeout auf, um
    die restliche Ausgabe noch einzusammeln (siehe CPython-Quelltext,
    subprocess.run, Zweig '_mswindows'). Diese zweite Runde wartet auf das
    Pipe-Ende. Ein vom Kind gestarteter Enkel - bei .bat/cmd.exe/Installern
    der Normalfall ('start ...', Java-Launcher, nachgeladene Setup-Stufen) -
    hat die Pipe-Handles geerbt und haelt sie offen; der Aufruf kehrt dann
    erst zurueck, wenn der Enkel VON SICH AUS endet.
    Nachgestellt: .bat mit 30-s-Enkel, timeout=5 -> TimeoutExpired erst nach
    30,5 s. Damit stand der komplette Zwei-Stufen-Lauf still, weil Stage 1
    in relaunch_as_admin_and_wait() unbegrenzt auf Stage 2 wartet.

    Deshalb: Ausgabe in DATEIEN statt in Pipes umlenken - dann haelt kein
    geerbtes Handle den Aufruf auf -, mit wait(timeout=...) nur auf den
    direkten Kindprozess warten und bei Ablauf den ganzen Prozessbaum
    beenden.
    """
    _sweep_pending_temp_files()
    out_fd = err_fd = None
    out_path = err_path = None
    proc = None
    try:
        env = None
        if extra_env:
            env = os.environ.copy()
            env.update(extra_env)
        out_fd, out_path = tempfile.mkstemp(prefix="vorabcheck_out_")
        err_fd, err_path = tempfile.mkstemp(prefix="vorabcheck_err_")
        timed_out = False
        with os.fdopen(out_fd, "wb") as f_out, os.fdopen(err_fd, "wb") as f_err:
            out_fd = err_fd = None      # Besitz an fdopen uebergeben
            proc = subprocess.Popen(
                args,
                stdout=f_out,
                stderr=f_err,
                shell=False,
                cwd=cwd,
                startupinfo=get_startupinfo(),
                env=env,
            )
            try:
                proc.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                timed_out = True
                _kill_process_tree(proc)
        # Erst nach dem Schliessen lesen, sonst haelt der eigene
        # Schreib-Handle die Datei noch.
        stdout_text = _read_output_file(out_path)
        stderr_text = _read_output_file(err_path)
        if timed_out:
            return CmdResult(timed_out=True, error="Timeout",
                             stdout=stdout_text, stderr=stderr_text)
        return CmdResult(proc.returncode, stdout_text, stderr_text)
    except FileNotFoundError:
        return CmdResult(not_found=True, error="FileNotFound")
    except Exception as e:
        return CmdResult(error=str(e))
    finally:
        for fd in (out_fd, err_fd):
            if fd is not None:
                try:
                    os.close(fd)
                except Exception as _e:
                    logging.debug(f"run_cmd: close verworfen: {_e!r}")
        for pth in (out_path, err_path):
            _drop_temp_file(pth)

def long_path(p):
    if not IS_WINDOWS or not p:
        return p
    try:
        p_abs = os.path.abspath(p)
    except Exception:
        return p
    if p_abs.startswith("\\\\?\\"):
        return p_abs
    if p_abs.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p_abs[2:]
    return "\\\\?\\" + p_abs

def _strip_long_prefix(p):
    if not p:
        return p
    if p.startswith("\\\\?\\UNC\\"):
        return "\\\\" + p[8:]
    if p.startswith("\\\\?\\"):
        return p[4:]
    return p

def isdir_with_timeout(path, timeout=8.0):
    result = {"ok": False, "done": False}
    def check():
        try:
            result["ok"] = os.path.isdir(path)
        except Exception as _e:
            logging.debug(f"check: Exception verworfen: {_e!r}")
        finally:
            result["done"] = True
    t = threading.Thread(target=check, daemon=True)
    t.start()
    t.join(timeout)
    if not result["done"]:
        return None
    return result["ok"]

def same_file(a, b):
    try:
        return os.path.normcase(os.path.realpath(a)) == os.path.normcase(os.path.realpath(b))
    except Exception:
        return False

def version_tuple_from_name(name, prefix):
    m = re.search(re.escape(prefix) + r"[^0-9]*(\d+)(?:[._](\d+))?(?:[._](\d+))?", name.lower())
    if not m:
        return (0, 0, 0)
    parts = [int(g or 0) for g in m.groups()]
    # Kompakte Schreibweise ohne Trenner normalisieren:
    # "Python312" -> (3,12,0), "Python39" -> (3,9,0).
    # Sonst waere (312,0,0) nicht mit (3,12,0) aus "Python 3.12" vergleichbar.
    if m.group(2) is None and parts[0] >= 30:
        s = str(parts[0])
        parts = [int(s[0]), int(s[1:]), 0]
    return tuple(parts)

def natural_version_key(name):
    """Sortierschluessel aus allen Zahlen im Namen.
    Verhindert, dass 'gs9.56' lexikografisch vor 'gs10.03' einsortiert wird."""
    nums = re.findall(r"\d+", os.path.basename(str(name or "")))[:4]
    return tuple(int(n) for n in nums) + (0,) * (4 - len(nums))

def sha256_of(path):
    try:
        h = hashlib.sha256()
        with open(long_path(path), "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None

# Platzhalter fuer eine Datei, deren Hash nicht gebildet werden konnte.
# Kein gueltiger SHA-256-Wert, faellt also in jedem Vergleich auf und kann
# nie versehentlich als "passt" durchgehen.
HASH_UNREADABLE = "<nicht lesbar>"

def hash_tree(root):
    """SHA-256 aller Dateien unterhalb von root, Schluessel = relativer Pfad.

    Wird fuer Stage-Ordner gebraucht (z. B. veraPDF): dort startet Stage 2 eine
    .bat, die ihrerseits weitere Dateien aus demselben Ordner aufruft. Ein Hash
    nur auf die .bat wuerde einen Austausch der JAR/EXE nicht bemerken.
    """
    out = {}
    try:
        root_abs = os.path.abspath(root)
    except Exception:
        return out
    try:
        for dirpath, dirnames, filenames in os.walk(long_path(root_abs)):
            for name in filenames:
                full = _strip_long_prefix(os.path.join(dirpath, name))
                try:
                    rel = os.path.relpath(full, root_abs)
                except Exception:
                    continue
                # Nicht lesbare Dateien NICHT stillschweigend weglassen: sie
                # verschwaenden sonst aus der Erwartungsliste und blieben in
                # Stage 2 unbemerkt ungeprueft.
                digest = sha256_of(full)
                out[rel.replace("\\", "/").lower()] = digest or HASH_UNREADABLE
    except Exception as _e:
        logging.debug(f"hash_tree: Exception verworfen: {_e!r}")
    return out

# =============================================================================
# Two-Stage-Admin-Modell
# =============================================================================

def parse_run_mode():
    global RUN_MODE, STAGE_ROOT, STAGE_MANIFEST_FILE, STAGE_RESULT_FILE
    global LOG_FILE, ASSUME_YES, NO_INPUT, NO_PAUSE

    argv = sys.argv[1:]
    ASSUME_YES = any(a in ("--yes", "-y") for a in argv)
    NO_INPUT   = "--no-input" in argv
    NO_PAUSE   = "--no-pause" in argv or NO_INPUT

    if ADMIN_STAGE_FLAG in sys.argv:
        idx = sys.argv.index(ADMIN_STAGE_FLAG)
        if idx + 1 < len(sys.argv):
            STAGE_ROOT = os.path.abspath(sys.argv[idx + 1].strip().strip('"\''))
        RUN_MODE = "stage2"
    else:
        RUN_MODE = "stage1"
        STAGE_ROOT = STAGE_DIR

    # Stage 2 schreibt in dasselbe Logfile wie Stage 1 (--log <pfad>)
    if LOG_FLAG in sys.argv:
        idx = sys.argv.index(LOG_FLAG)
        if idx + 1 < len(sys.argv):
            candidate = sys.argv[idx + 1].strip().strip('"\'')
            if candidate:
                LOG_FILE = os.path.abspath(candidate)

    STAGE_MANIFEST_FILE = os.path.join(STAGE_ROOT, "stage_manifest.json")
    STAGE_RESULT_FILE   = os.path.join(STAGE_ROOT, "stage_result.json")

def get_user_arg_index_for_deploy():
    skip = False
    for i, a in enumerate(sys.argv[1:], start=1):
        if skip:
            skip = False
            continue
        if a in VALUE_FLAGS:
            skip = True
            continue
        if a in BOOL_FLAGS or a.startswith("-"):
            continue
        return i
    return None

def relaunch_as_admin_and_wait(stage_dir):
    """Rueckgabe: Exit-Code (int) bei erfolgreichem Lauf, sonst None.

    Vorher wurden False/None/int gemischt zurueckgegeben und mit 'rc is False'
    verglichen – ein regulaerer Exit-Code 0 war davon nicht unterscheidbar.
    """
    if not IS_WINDOWS:
        return None
    SEE_MASK_NOCLOSEPROCESS = 0x00000040

    class SHELLEXECUTEINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize",       wintypes.DWORD),
            ("fMask",        wintypes.ULONG),
            ("hwnd",         wintypes.HWND),
            ("lpVerb",       wintypes.LPCWSTR),
            ("lpFile",       wintypes.LPCWSTR),
            ("lpParameters", wintypes.LPCWSTR),
            ("lpDirectory",  wintypes.LPCWSTR),
            ("nShow",        ctypes.c_int),
            ("hInstApp",     wintypes.HINSTANCE),
            ("lpIDList",     ctypes.c_void_p),
            ("lpClass",      wintypes.LPCWSTR),
            ("hkeyClass",    wintypes.HKEY),
            ("dwHotKey",     wintypes.DWORD),
            ("hIcon",        wintypes.HANDLE),
            ("hProcess",     wintypes.HANDLE),
        ]

    # argtypes/restype explizit setzen: ohne sie behandelt ctypes HANDLE-Werte
    # als C-int (32 Bit) und kann 64-Bit-Handles abschneiden.
    shell32  = ctypes.windll.shell32
    kernel32 = ctypes.windll.kernel32
    shell32.ShellExecuteExW.argtypes  = [ctypes.POINTER(SHELLEXECUTEINFO)]
    shell32.ShellExecuteExW.restype   = wintypes.BOOL
    kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    kernel32.WaitForSingleObject.restype  = wintypes.DWORD
    kernel32.GetExitCodeProcess.argtypes  = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    kernel32.GetExitCodeProcess.restype   = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype  = wintypes.BOOL

    extra = []
    if NO_INPUT or NO_PAUSE:
        extra.append("--no-pause")
    if ASSUME_YES:
        extra.append("--yes")
    extra_str = (" " + " ".join(extra)) if extra else ""

    info = SHELLEXECUTEINFO()
    info.cbSize = ctypes.sizeof(info)
    info.fMask  = SEE_MASK_NOCLOSEPROCESS
    info.lpVerb = "runas"
    info.nShow  = 1

    if getattr(sys, "frozen", False):
        info.lpFile = sys.executable
        info.lpParameters = (f'{ADMIN_STAGE_FLAG} "{stage_dir}" '
                             f'{LOG_FLAG} "{LOG_FILE}"{extra_str}')
    else:
        info.lpFile = sys.executable
        script = os.path.abspath(__file__)
        info.lpParameters = (f'"{script}" {ADMIN_STAGE_FLAG} "{stage_dir}" '
                             f'{LOG_FLAG} "{LOG_FILE}"{extra_str}')

    try:
        ok = shell32.ShellExecuteExW(ctypes.byref(info))
    except Exception as e:
        log_err(f"   ❌ ShellExecuteEx fehlgeschlagen: {e}")
        return None
    if not ok or not info.hProcess:
        try:
            err = ctypes.GetLastError()
        except Exception:
            err = 0
        if err == 1223:   # ERROR_CANCELLED
            log_warn("   ⚠️ UAC-Abfrage wurde vom Benutzer abgebrochen.")
        else:
            log_warn(f"   ⚠️ UAC-Re-Launch nicht möglich (ShellExecuteEx-Fehler {err}).")
        return None

    handle = info.hProcess
    try:
        # Frueher wurde hier mit INFINITE gewartet. Haengt Stage 2 - etwa an
        # einem unsichtbar hinter anderen Fenstern liegenden Installer-Dialog -,
        # stand Stage 1 ohne jede Ausgabe still und war von aussen nicht von
        # einem Absturz zu unterscheiden.
        # Ein harter Timeout loest das NICHT: Stage 1 laeuft auf mittlerer
        # Integritaetsstufe und kann den elevierten Stage-2-Prozess nicht
        # beenden. Ein Abbruch wuerde ihn nur verwaisen lassen, waehrend beide
        # Stufen weiter in dieselbe Logdatei schreiben und der Anwender einen
        # zweiten Lauf startet. Deshalb weiterhin unbegrenzt warten - aber in
        # Minutenschritten und mit sichtbarem Lebenszeichen, damit erkennbar
        # bleibt, worauf gewartet wird.
        WAIT_TIMEOUT_RC = 0x00000102
        WAIT_FAILED_RC  = 0xFFFFFFFF
        wartete_min = 0
        while True:
            rc = kernel32.WaitForSingleObject(handle, 60_000)
            if rc != WAIT_TIMEOUT_RC:
                if rc == WAIT_FAILED_RC:
                    log_err("   ❌ Warten auf Admin-Prozess fehlgeschlagen (WAIT_FAILED).")
                    return None
                break
            wartete_min += 1
            if wartete_min % 5 == 0:
                log_info(f"   ⏳ Stage 2 läuft noch ({wartete_min} min). Falls nichts "
                         f"vorangeht: nach einem offenen Installer-Fenster sehen.")
        exit_code = wintypes.DWORD()
        if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
            log_err("   ❌ Exit-Code des Admin-Prozesses nicht lesbar.")
            return None
        return int(exit_code.value)
    except Exception as e:
        log_err(f"   ❌ Warten auf Admin-Prozess fehlgeschlagen: {e}")
        return None
    finally:
        try:
            kernel32.CloseHandle(handle)
        except Exception as _e:
            logging.debug(f"relaunch_as_admin_and_wait: Exception verworfen: {_e!r}")

def harden_stage_dir():
    """Stage-Verzeichnis gegen Manipulation durch andere lokale Konten schuetzen.

    Stage 2 fuehrt Installer aus diesem Verzeichnis mit Adminrechten aus.
    Bliebe es fuer 'Jeder'/'Benutzer' beschreibbar, koennte ein anderer
    Standardbenutzer die Dateien zwischen Stage 1 und Stage 2 austauschen
    (lokale Rechteausweitung). Best effort – Fehler sind nicht fatal.
    """
    if not IS_WINDOWS:
        return
    try:
        os.makedirs(STAGE_ROOT, exist_ok=True)
        user = os.environ.get("USERNAME") or ""
        args = ["icacls", STAGE_ROOT, "/inheritance:r",
                "/grant", "*S-1-5-32-544:(OI)(CI)F",   # Administratoren
                "/grant", "*S-1-5-18:(OI)(CI)F"]       # SYSTEM
        if user:
            args += ["/grant", f"{user}:(OI)(CI)M"]
        run_cmd(args, timeout=60)
    except Exception as _e:
        logging.debug(f"harden_stage_dir: Exception verworfen: {_e!r}")

def harden_program_dir(path):
    """Programmverzeichnis gegen Manipulation durch andere lokale Konten schuetzen.

    Aus C:\\OCR startet 5_OCR_PDF.py tesseract.exe, gswin64c.exe, jbig2.exe
    und verapdf.bat und laedt von dort DLLs - mit Vorrang vor
    C:\\Program Files und gegebenenfalls mit Adminrechten. Ein direkt unter
    C:\\ angelegter Ordner erbt aber die Standard-ACL von C:\\ und ist damit
    fuer jeden authentifizierten Benutzer beschreibbar. Ohne diese Haertung
    koennte ein Standardbenutzer dort ein eigenes Programm ablegen, das beim
    naechsten Admin-Lauf mit Administratorrechten ausgefuehrt wird.

    Rueckgabe True, wenn die ACL gesetzt werden konnte. Schlaegt es fehl
    (kein Adminrecht, kein Besitz), wird das als offene Handarbeit vermerkt
    statt still uebergangen - es ist eine Sicherheitszusage, kein Komfort.
    """
    if not IS_WINDOWS or not os.path.isdir(path):
        return False
    # SIDs statt Klartextnamen: 'Benutzer'/'Users' ist sprachabhaengig.
    args = ["icacls", path, "/inheritance:r",
            "/grant", "*S-1-5-32-544:(OI)(CI)F",    # Administratoren
            "/grant", "*S-1-5-18:(OI)(CI)F",        # SYSTEM
            "/grant", "*S-1-5-32-545:(OI)(CI)RX"]   # Benutzer: nur lesen/starten
    try:
        if run_cmd(args, timeout=60).ok:
            log_info(f"✅ Zugriffsrechte gehärtet: {path}")
            return True
    except Exception as _e:
        logging.debug(f"harden_program_dir: Exception verworfen: {_e!r}")
    log_warn(f"   ⚠️ Zugriffsrechte für {path} konnten nicht gehärtet werden.")
    STATE["manual_actions"].append(
        f'Zugriffsrechte härten (als Administrator): icacls "{path}" '
        f'/inheritance:r /grant *S-1-5-32-544:(OI)(CI)F /grant *S-1-5-18:(OI)(CI)F '
        f'/grant *S-1-5-32-545:(OI)(CI)RX'
    )
    return False

def stage_copy_file(source, label=None):
    if not source or not os.path.isfile(long_path(source)):
        return None
    try:
        os.makedirs(STAGE_ROOT, exist_ok=True)
        target = os.path.join(STAGE_ROOT, os.path.basename(source))
        if not os.path.isfile(long_path(target)) or not same_file(source, target):
            shutil.copy2(long_path(source), long_path(target))
        if label:
            log_info(f"   📦 Stage-Kopie: {label} -> {target}")
        return target
    except Exception as e:
        log_err(f"   ❌ Stage-Kopie fehlgeschlagen ({source}): {e}")
        return None

def stage_copy_dir(source_dir, sub_name=None):
    if not source_dir or not os.path.isdir(long_path(source_dir)):
        return None
    try:
        target = os.path.join(STAGE_ROOT, sub_name or os.path.basename(source_dir.rstrip("\\/")))
        if os.path.isdir(long_path(target)):
            shutil.rmtree(long_path(target), ignore_errors=True)
        shutil.copytree(long_path(source_dir), long_path(target))
        log_info(f"   📦 Stage-Kopie (Ordner): {source_dir} -> {target}")
        return target
    except Exception as e:
        log_err(f"   ❌ Stage-Kopie (Ordner) fehlgeschlagen ({source_dir}): {e}")
        return None

def stage_ocr_binary_for_admin(source, target, label=None):
    staged = stage_copy_file(source, label=label or os.path.basename(target))
    if not staged:
        return False
    STATE["admin_pending"].append({
        "tool":      "ocr_binary",
        "installer": staged,
        "target":    target,
    })
    return True

def verify_staged_dir(entry):
    """Prueft einen komplett gestageten Ordner gegen die Hash-Liste im Manifest.

    Auch zusaetzliche Dateien fuehren zum Abbruch: sie koennten von der .bat
    mitgeladen werden oder eine erwartete Datei verdecken.
    """
    root = entry.get("dir")
    expected = entry.get("files") or {}
    if not root:
        return True
    # Ordner-Stage OHNE Hash-Liste ist kein "nichts zu pruefen", sondern ein
    # fehlender Nachweis - und damit ein Abbruchgrund. Frueher lieferte der
    # Zweig True und der komplette Baum lief ungeprueft mit Adminrechten.
    if not expected:
        log_err(f"   ❌ Keine Hash-Liste im Manifest für Stage-Ordner '{root}' – Abbruch.")
        return False
    unlesbar = sorted(k for k, v in expected.items() if v == HASH_UNREADABLE)
    if unlesbar:
        log_err(f"   ❌ Hash-Liste für '{root}' enthält {len(unlesbar)} nicht lesbare "
                f"Datei(en) – Abbruch.")
        for k in unlesbar[:10]:
            log_err(f"      nicht lesbar: {k}")
        return False
    if not os.path.isdir(long_path(root)):
        log_err(f"   ❌ Stage-Ordner fehlt: {root}")
        return False

    actual  = hash_tree(root)
    missing = sorted(k for k in expected if k not in actual)
    changed = sorted(k for k in expected if k in actual and actual[k] != expected[k])
    added   = sorted(k for k in actual if k not in expected)

    if not (missing or changed or added):
        return True

    log_err(f"   ❌ Stage-Ordner '{os.path.basename(root)}' wurde seit Stage 1 verändert – Abbruch.")
    for k in missing[:10]:
        log_err(f"      fehlt:     {k}")
    for k in changed[:10]:
        log_err(f"      geändert:  {k}")
    for k in added[:10]:
        log_err(f"      zusätzlich: {k}")
    rest = len(missing) + len(changed) + len(added) - min(len(missing), 10) \
           - min(len(changed), 10) - min(len(added), 10)
    if rest > 0:
        log_err(f"      ... und {rest} weitere Abweichungen")
    return False

def verify_staged_entry(entry):
    """Prueft die in Stage 1 notierten SHA-256-Summen vor der Admin-Ausfuehrung."""
    src = entry.get("installer")
    expected = entry.get("sha256")
    if not src or not os.path.isfile(long_path(src)):
        log_err(f"   ❌ Stage-Datei fehlt: {src}")
        return False
    # Fehlende Pruefsumme = fehlender Nachweis = Abbruch. Frueher wurde hier
    # nur gewarnt und die Datei mit Adminrechten trotzdem gestartet - genau
    # der Austausch zwischen Stage 1 und Stage 2, gegen den dieses Modell
    # gebaut ist, waere damit unbemerkt geblieben.
    if not expected:
        log_err(f"   ❌ Keine Prüfsumme im Manifest für {os.path.basename(src)} – "
                f"Ausführung abgebrochen.")
        return False
    else:
        actual = sha256_of(src)
        if actual != expected:
            log_err(f"   ❌ Prüfsumme von {os.path.basename(src)} weicht ab – Ausführung abgebrochen.")
            log_err(f"      erwartet: {expected}")
            log_err(f"      gefunden: {actual}")
            return False
    # Bei Ordner-Stages zusaetzlich den gesamten Baum pruefen.
    return verify_staged_dir(entry)

def write_stage_manifest(manifest):
    try:
        os.makedirs(STAGE_ROOT, exist_ok=True)
        with open(STAGE_MANIFEST_FILE, "w", encoding="utf-8") as f:
            json.dump(manifest, f, indent=2, ensure_ascii=False)
        log_info(f"   ✅ Stage-Manifest geschrieben: {STAGE_MANIFEST_FILE}")
        return True
    except Exception as e:
        log_err(f"   ❌ Stage-Manifest konnte nicht geschrieben werden: {e}")
        return False

def read_stage_manifest():
    try:
        with open(STAGE_MANIFEST_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        log_err(f"❌ Stage-Manifest konnte nicht gelesen werden: {e}")
        return None

def write_stage_result(data):
    try:
        os.makedirs(STAGE_ROOT, exist_ok=True)
        with open(STAGE_RESULT_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        log_err(f"   ❌ Stage-Ergebnis konnte nicht geschrieben werden: {e}")

def read_stage_result():
    try:
        with open(STAGE_RESULT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

# =============================================================================
# Deploy-Quelle und Deploy-Index
# =============================================================================

def resolve_deploy_source():
    global DEPLOY_SOURCE

    user_idx = get_user_arg_index_for_deploy()
    if user_idx is not None:
        arg = sys.argv[user_idx].strip().strip('"\'')
        if arg:
            DEPLOY_SOURCE = arg
            return

    env = (os.environ.get("VORABCHECK_DEPLOY", "") or "").strip().strip('"\'')
    if env:
        DEPLOY_SOURCE = env
        return

    if ASSUME_YES or NO_INPUT or not sys.stdin or not sys.stdin.isatty():
        DEPLOY_SOURCE = DEPLOY_SOURCE_DEFAULT
        return

    print()
    print(f"  Deploy-Quelle (Standard): {DEPLOY_SOURCE_DEFAULT}")
    print("  ENTER zum Übernehmen oder anderen Pfad eingeben:")
    try:
        answer = input("  > ").strip().strip('"\'')
    except (EOFError, KeyboardInterrupt):
        answer = ""
    DEPLOY_SOURCE = answer if answer else DEPLOY_SOURCE_DEFAULT

def _is_excluded_dir(name):
    return name.lower() in DEPLOY_INDEX_EXCLUDES

def build_deploy_index(root):
    DEPLOY_INDEX.clear()
    DEPLOY_INDEX_DIRS.clear()
    root_path = long_path(root)
    count_files = 0
    count_dirs  = 0
    last_print  = time.time()

    try:
        for dirpath, dirnames, filenames in os.walk(root_path):
            dirnames[:] = [d for d in dirnames if not _is_excluded_dir(d)]
            DEPLOY_INDEX_DIRS.add(_strip_long_prefix(dirpath))
            count_dirs += 1
            for name in filenames:
                key = name.lower()
                if os.path.splitext(key)[1] not in DEPLOY_INDEX_EXTENSIONS:
                    continue
                DEPLOY_INDEX.setdefault(key, []).append(
                    _strip_long_prefix(os.path.join(dirpath, name))
                )
                count_files += 1
            now = time.time()
            if now - last_print >= 2.0:
                print(f"   … indiziert: {count_files} Dateien in {count_dirs} Ordnern")
                last_print = now
    except Exception as e:
        log_warn(f"   ⚠️ Fehler beim Indizieren der Deploy-Quelle: {e}")
    return count_files

def check_deploy_source():
    log_info("--- Prüfe Deploy-Quelle ---")
    log_info(f"   Pfad: {DEPLOY_SOURCE}")
    status = isdir_with_timeout(DEPLOY_SOURCE, timeout=10.0)
    if status is None:
        log_err(f"❌ Deploy-Quelle reagiert nicht (Timeout): {DEPLOY_SOURCE}")
        log_warn("   Netzlaufwerk eventuell nicht verbunden.")
        print("")
        return False
    if not status:
        log_err(f"❌ Deploy-Quelle nicht erreichbar: {DEPLOY_SOURCE}")
        log_warn("   Ohne diese Quelle können Tools/Dateien nicht automatisch kopiert/installiert werden.")
        print("")
        return False
    log_info("✅ Deploy-Quelle erreichbar.")
    log_info("   Indiziere Inhalte (einmalig)...")
    log_info("   Hinweis: Bei Netzlaufwerken (UNC) und sehr vielen Dateien")
    log_info("   (>100.000) kann das Indizieren 10–30 Minuten dauern.")
    t0 = time.time()
    count = build_deploy_index(DEPLOY_SOURCE)
    dt = time.time() - t0
    log_info(f"   {count} relevante Dateien in {len(DEPLOY_INDEX_DIRS)} Ordnern indiziert ({dt:.1f}s).")
    print("")
    return True

def find_in_deploy(pattern):
    pattern_lower = pattern.lower()
    matches = []
    for name_lower, paths in DEPLOY_INDEX.items():
        if fnmatch.fnmatch(name_lower, pattern_lower):
            matches.extend(paths)
    def sortkey(p):
        try:
            return os.path.getmtime(long_path(p))
        except Exception:
            return 0
    return sorted(set(matches), key=sortkey, reverse=True)

def find_dir_in_deploy(pattern):
    pattern_lower = pattern.lower()
    matches = [d for d in DEPLOY_INDEX_DIRS
               if fnmatch.fnmatch(os.path.basename(d).lower(), pattern_lower)]
    def sortkey(p):
        try:
            return os.path.getmtime(long_path(p))
        except Exception:
            return 0
    return sorted(matches, key=sortkey, reverse=True)

def find_wheels_dirs():
    out = []
    for d in DEPLOY_INDEX_DIRS:
        bn = os.path.basename(d).lower()
        if bn in ("wheels", "wheel", "pip-wheels", "pip_wheels"):
            out.append(d)
    return out

# =============================================================================
# Registry und Umgebungsvariablen
# =============================================================================

HWND_BROADCAST    = 0xFFFF
WM_SETTINGCHANGE  = 0x001A
SMTO_ABORTIFHUNG  = 0x0002

def broadcast_env_change():
    if not IS_WINDOWS:
        return
    try:
        result = ctypes.c_long()
        ctypes.windll.user32.SendMessageTimeoutW(
            HWND_BROADCAST, WM_SETTINGCHANGE, 0,
            ctypes.c_wchar_p("Environment"),
            SMTO_ABORTIFHUNG, 5000, ctypes.byref(result),
        )
    except Exception as _e:
        logging.debug(f"broadcast_env_change: Exception verworfen: {_e!r}")

def get_user_env(var_name):
    if not IS_WINDOWS:
        return None, None
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_READ) as key:
            value, regtype = winreg.QueryValueEx(key, var_name)
            return value, regtype
    except OSError:
        return None, None

def get_system_env(var_name):
    if not IS_WINDOWS:
        return None, None
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment",
                            0, winreg.KEY_READ) as key:
            value, regtype = winreg.QueryValueEx(key, var_name)
            return value, regtype
    except OSError:
        return None, None

def set_user_env(var_name, value, regtype=None):
    if not IS_WINDOWS:
        return False
    if regtype is None:
        regtype = winreg.REG_SZ
    existing, existing_type = get_user_env(var_name)
    # Nur ueberspringen, wenn Wert UND Typ bereits stimmen. Die urspruengliche
    # Bedingung war durch "or existing_type is not None" immer wahr, ein
    # falscher Registry-Typ wurde daher nie korrigiert.
    if existing == value and existing_type == regtype:
        os.environ[var_name] = value
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_ALL_ACCESS) as key:
            winreg.SetValueEx(key, var_name, 0, regtype, value)
        os.environ[var_name] = value
        broadcast_env_change()
        STATE["env_set"].append(f"{var_name} = {value}")
        return True
    except Exception as e:
        log_err(f"   ❌ Konnte {var_name} nicht setzen: {e}")
        return False

def append_user_path(new_entry):
    if not IS_WINDOWS:
        return False
    if not new_entry or not os.path.isdir(new_entry):
        return False
    current, regtype = get_user_env("Path")
    if current is None:
        current = ""
        regtype = winreg.REG_EXPAND_SZ
    entries = [p.strip() for p in current.split(os.pathsep) if p.strip()]
    entries_norm = {os.path.normcase(os.path.normpath(p)) for p in entries}
    new_norm = os.path.normcase(os.path.normpath(new_entry))
    env_path_norm = {os.path.normcase(os.path.normpath(p))
                     for p in os.environ.get("PATH", "").split(os.pathsep) if p}
    if new_norm in entries_norm:
        if new_norm not in env_path_norm:
            os.environ["PATH"] = os.environ.get("PATH", "") + os.pathsep + new_entry
        return False
    entries.append(new_entry)
    new_value = os.pathsep.join(entries)
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_ALL_ACCESS) as key:
            winreg.SetValueEx(key, "Path", 0, regtype, new_value)
        if new_norm not in env_path_norm:
            os.environ["PATH"] = os.environ.get("PATH", "") + os.pathsep + new_entry
        broadcast_env_change()
        STATE["path_added"].append(new_entry)
        return True
    except Exception as e:
        log_err(f"   ❌ Konnte PATH nicht ergänzen: {e}")
        return False

def append_system_path(new_entry):
    if not IS_WINDOWS or not is_admin():
        return False
    if not new_entry or not os.path.isdir(new_entry):
        return False
    current, regtype = get_system_env("Path")
    if current is None:
        current = ""
        regtype = winreg.REG_EXPAND_SZ
    entries = [p.strip() for p in current.split(os.pathsep) if p.strip()]
    entries_norm = {os.path.normcase(os.path.normpath(p)) for p in entries}
    new_norm = os.path.normcase(os.path.normpath(new_entry))
    if new_norm in entries_norm:
        return False
    entries.append(new_entry)
    new_value = os.pathsep.join(entries)
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment",
                            0, winreg.KEY_ALL_ACCESS) as key:
            winreg.SetValueEx(key, "Path", 0, regtype, new_value)
        os.environ["PATH"] = os.environ.get("PATH", "") + os.pathsep + new_entry
        broadcast_env_change()
        STATE["path_added"].append(f"[System] {new_entry}")
        return True
    except Exception as e:
        log_err(f"   ❌ Konnte System-PATH nicht ergänzen: {e}")
        return False

def refresh_env_from_registry():
    if not IS_WINDOWS:
        return
    saved_extra = []
    cur_path = os.environ.get("PATH", "")

    sys_path, _   = get_system_env("Path")
    user_path, _  = get_user_env("Path")
    sys_path  = sys_path or ""
    user_path = user_path or ""

    reg_norm = set()
    for p in (sys_path + os.pathsep + user_path).split(os.pathsep):
        if p.strip():
            reg_norm.add(os.path.normcase(os.path.normpath(p)))

    for entry in cur_path.split(os.pathsep):
        e = entry.strip()
        if not e:
            continue
        n = os.path.normcase(os.path.normpath(e))
        if n not in reg_norm:
            saved_extra.append(e)

    parts = [p for p in (sys_path, user_path) if p]
    if not parts and not saved_extra:
        return
    merged = os.pathsep.join(parts + saved_extra)
    try:
        merged = os.path.expandvars(merged)
    except Exception as _e:
        logging.debug(f"refresh_env_from_registry: Exception verworfen: {_e!r}")
    os.environ["PATH"] = merged

# =============================================================================
# Verzeichnisse und Kopier-Funktionen
# =============================================================================

def ensure_directories():
    log_info("--- Prüfe und erstelle Standard-Verzeichnisse ---")
    targets = [OCR_DIR, TMP_DIR, BIN_DIR]
    if RUN_MODE == "stage1":
        targets.append(STAGE_DIR)
    for d in targets:
        if os.path.isdir(d):
            log_info(f"✅ Verzeichnis vorhanden: {d}")
        else:
            try:
                os.makedirs(d, exist_ok=True)
                log_info(f"✅ Verzeichnis angelegt: {d}")
                STATE["created_dirs"].append(d)
            except Exception as e:
                log_err(f"❌ Verzeichnis konnte nicht angelegt werden: {d} ({e})")
    # Stage-Verzeichnis sofort absichern, nicht erst kurz vor dem UAC-Start:
    # sonst liegen die Installer zwischenzeitlich in einem Ordner, in den
    # andere lokale Konten schreiben duerfen.
    if RUN_MODE == "stage1" and os.path.isdir(STAGE_DIR):
        harden_stage_dir()
    # C:\OCR ist Programmverzeichnis, kein Arbeitsordner: 5_OCR_PDF.py startet
    # daraus Programme und laedt DLLs, mit Vorrang vor C:\Program Files. Als
    # direkt unter C:\ angelegter Ordner waere es sonst fuer jeden
    # authentifizierten Benutzer beschreibbar - bei einem spaeteren Admin-Lauf
    # eine lokale Rechteausweitung. Nur im Admin-Kontext versuchen; ohne
    # Adminrechte schlaegt icacls ohnehin fehl und es entstuende bei jedem
    # Lauf ein Falscheintrag in den offenen Handarbeiten.
    if os.path.isdir(OCR_DIR) and is_admin():
        harden_program_dir(OCR_DIR)
    print("")

def _ensure_parent_dir(target):
    parent = os.path.dirname(target)
    if parent and not os.path.isdir(parent):
        try:
            os.makedirs(parent, exist_ok=True)
        except Exception:
            try:
                os.makedirs(long_path(parent), exist_ok=True)
            except Exception as _e:
                logging.debug(f"_ensure_parent_dir: Exception verworfen: {_e!r}")

def copy_file_safe(source, target, retries=2, silent=False, soft=False):
    """soft=True: Fehlschlag nur als Warnung protokollieren.

    Wird genutzt, wenn ein Fehlschlag erwartbar ist (z. B. Schreibversuch nach
    C:\\OCR ohne Adminrechte) und anschliessend auf Stage 2 ausgewichen wird.
    Sonst wuerde der Exit-Code faelschlich auf 2 (Fehler) statt 1 stehen.
    """
    last_err = None
    for attempt in range(retries + 1):
        try:
            if os.path.isfile(long_path(source)) and os.path.isfile(long_path(target)) and same_file(source, target):
                return True
            _ensure_parent_dir(target)
            shutil.copy2(long_path(source), long_path(target))
            if not silent:
                suffix = f" (Versuch {attempt+1})" if attempt > 0 else ""
                log_info(f"   ✅ Kopiert{suffix}: {os.path.basename(source)} -> {target}")
            STATE["copied_files"].append(target)
            return True
        except shutil.SameFileError:
            return True
        except Exception as e:
            last_err = e
            if attempt < retries:
                time.sleep(0.8 * (attempt + 1))
    msg = f"   ❌ Kopieren fehlgeschlagen ({source} -> {target}): {last_err}"
    if soft:
        log_warn(msg)
    else:
        log_err(msg)
    return False

def copy_tree_safe(src_dir, dst_dir):
    try:
        if os.path.isdir(long_path(dst_dir)) and not same_file(src_dir, dst_dir):
            shutil.rmtree(long_path(dst_dir), ignore_errors=True)
        shutil.copytree(long_path(src_dir), long_path(dst_dir))
        log_info(f"   ✅ Ordner kopiert: {src_dir} -> {dst_dir}")
        STATE["copied_files"].append(dst_dir)
        return True
    except Exception as e:
        log_err(f"   ❌ Ordner kopieren fehlgeschlagen ({src_dir} -> {dst_dir}): {e}")
        return False

# =============================================================================
# Office-Erkennung und Trust-Center
# =============================================================================

def _detect_office_versions():
    if not IS_WINDOWS:
        return []
    versions = []
    candidates = ["16.0", "15.0", "14.0"]

    # 64-Bit- und 32-Bit-View prüfen (32-Bit-Office MSI liegt im WOW6432Node)
    registry_views = (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY)

    c2r = {"products": [], "language": None, "platform": None, "version_to_report": None}
    for view_flag in registry_views:
        if c2r["products"] or c2r["version_to_report"]:
            break
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                                0, winreg.KEY_READ | view_flag) as key:
                try:
                    pids, _ = winreg.QueryValueEx(key, "ProductReleaseIds")
                    c2r["products"] = [p.strip() for p in pids.split(",") if p.strip()]
                except OSError:
                    pass
                try:
                    c2r["language"], _ = winreg.QueryValueEx(key, "ClientCulture")
                except OSError:
                    pass
                try:
                    c2r["platform"], _ = winreg.QueryValueEx(key, "Platform")
                except OSError:
                    pass
                try:
                    c2r["version_to_report"], _ = winreg.QueryValueEx(key, "VersionToReport")
                except OSError:
                    pass
        except OSError:
            pass

    # Major-Version aus VersionToReport ableiten (z. B. "16.0.10417.20117" -> "16.0")
    c2r_major = ""
    if c2r["version_to_report"]:
        m = re.match(r"(\d+\.\d+)", c2r["version_to_report"])
        if m:
            c2r_major = m.group(1)

    for ver in candidates:
        info = {"version": ver, "products": [], "language": None,
                "platform": None, "version_to_report": None,
                "install_root": None}

        # Versionsspezifischer Word-InstallRoot
        for view_flag in registry_views:
            if info["install_root"]:
                break
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                    rf"SOFTWARE\Microsoft\Office\{ver}\Word\InstallRoot",
                                    0, winreg.KEY_READ | view_flag) as key:
                    p, _ = winreg.QueryValueEx(key, "Path")
                    info["install_root"] = p
            except OSError:
                pass

        # ClickToRun-Daten nur zuordnen, wenn der globale Major mit dem Kandidaten übereinstimmt
        if c2r_major == ver:
            info["products"]          = list(c2r["products"])
            info["language"]          = c2r["language"]
            info["platform"]          = c2r["platform"]
            info["version_to_report"] = c2r["version_to_report"]

        if info["install_root"] or info["products"]:
            versions.append(info)
    return versions

def _detect_office_ui_language(ver):
    out = []
    for hive, label in ((winreg.HKEY_CURRENT_USER, "HKCU"),
                        (winreg.HKEY_LOCAL_MACHINE, "HKLM")):
        try:
            with winreg.OpenKey(hive,
                                rf"SOFTWARE\Microsoft\Office\{ver}\Common\LanguageResources",
                                0, winreg.KEY_READ) as key:
                for name in ("PreferredUILanguage", "UILanguage", "InstallLanguage"):
                    try:
                        v, _ = winreg.QueryValueEx(key, name)
                        out.append((label, name, str(v)))
                    except OSError:
                        continue
        except OSError:
            continue
    return out

def _label_for_product(p):
    return OFFICE_RELEASE_LABELS.get(p, p)

def _check_office_trust_center(ver):
    log_info(f"   --- Trust-Center / Add-In-Settings (Office {ver}) ---")
    issues = []
    for app in OFFICE_TRUST_APPS:
        sec_path = rf"Software\Microsoft\Office\{ver}\{app}\Security"
        access_vbom = None
        vba_warnings = None
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sec_path, 0, winreg.KEY_READ) as key:
                try:
                    access_vbom, _ = winreg.QueryValueEx(key, "AccessVBOM")
                except OSError:
                    access_vbom = None
                try:
                    vba_warnings, _ = winreg.QueryValueEx(key, "VBAWarnings")
                except OSError:
                    vba_warnings = None
        except OSError:
            pass

        if access_vbom == 1:
            log_info(f"      ✅ {app}.AccessVBOM = 1 (Zugriff auf VBA-Projektmodell vertraut)")
        else:
            log_warn(f"      ⚠️ {app}.AccessVBOM nicht gesetzt – Skripte mit VBA-Zugriff können scheitern.")
            issues.append((app, "AccessVBOM", access_vbom))

        if vba_warnings is None:
            log_info(f"      ℹ️ {app}.VBAWarnings nicht gesetzt (Standardverhalten).")
        else:
            # VBAWarnings: 1=alle aktiv (Sicherheitsrisiko), 2=Benachrichtigung,
            # 3=nur signierte, 4=alle deaktiviert
            vba_labels = {
                1: "alle Makros aktiv – Sicherheitsrisiko",
                2: "Benachrichtigung für alle Makros",
                3: "nur digital signierte Makros",
                4: "alle Makros deaktiviert",
            }
            label = vba_labels.get(vba_warnings, "unbekannter Wert")
            if vba_warnings == 1:
                log_warn(f"      ⚠️ {app}.VBAWarnings = 1  ({label})")
            else:
                log_info(f"      ℹ️ {app}.VBAWarnings = {vba_warnings}  ({label})")

    if issues:
        log_warn("   👉 Tipp: AccessVBOM kann pro App auf 1 gesetzt werden (HKCU)")
        log_warn("            – wird unten optional gesetzt.")
    return issues

def _set_office_access_vbom(ver, apps):
    if not IS_WINDOWS:
        return
    for app in apps:
        sec_path = rf"Software\Microsoft\Office\{ver}\{app}\Security"
        try:
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, sec_path, 0, winreg.KEY_ALL_ACCESS) as key:
                cur = None
                try:
                    cur, _ = winreg.QueryValueEx(key, "AccessVBOM")
                except OSError:
                    pass
                if cur == 1:
                    continue
                winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
                log_info(f"      ✅ {app}.AccessVBOM = 1 gesetzt (HKCU\\{sec_path}).")
                STATE["env_set"].append(f"HKCU:{sec_path}\\AccessVBOM = 1")
        except Exception as e:
            log_err(f"      ❌ Konnte {app}.AccessVBOM nicht setzen: {e}")

def check_office():
    log_info("--- Prüfe Microsoft Office Installation ---")
    if not IS_WINDOWS:
        log_warn("   Office-Check nur unter Windows verfügbar.")
        print("")
        return

    versions = _detect_office_versions()
    if not versions:
        log_err("❌ Keine Microsoft-Office-Installation erkannt.")
        log_warn("   Word/Excel/PowerPoint sind erforderlich für die Folge-Skripte.")
        log_download_hint("office")
        STATE["manual_actions"].append("Microsoft Office (z. B. 2019 dt./engl. oder 2024) installieren lassen.")
        print("")
        return

    for info in versions:
        ver  = info["version"]
        prod = info["products"]
        lang = info["language"] or ""
        plat = info["platform"] or ""
        vrep = info["version_to_report"] or ""
        root = info["install_root"] or ""

        if prod:
            for p in prod:
                log_info(f"✅ Office gefunden: {_label_for_product(p)}")
        else:
            log_info(f"✅ Office gefunden: Version {ver} (Klassische MSI-Installation)")

        if root:
            log_info(f"   Pfad:    {root}")
        if vrep:
            log_info(f"   Build:   {vrep}")
        if plat:
            log_info(f"   Bitness: {plat}")
            if plat.lower() == "x86" and sys.maxsize > 2**32:
                log_warn("   ⚠️ 32-Bit Office auf 64-Bit-System erkannt – Skripte mit nativen DLL-Aufrufen können scheitern.")
        if lang:
            log_info(f"   Setup-Sprache: {lang}")
            ll = lang.lower()
            if not (ll.startswith("de") or ll.startswith("en")):
                log_warn(f"   ⚠️ Setup-Sprache '{lang}' ist weder Deutsch noch Englisch – Folge-Skripte ggf. anpassen.")

        ui_langs = _detect_office_ui_language(ver)
        if ui_langs:
            for label, name, val in ui_langs:
                log_info(f"   UI-Sprache ({label}\\{name}): {val}")

        issues = _check_office_trust_center(ver)
        vbom_apps = [app for app, key, _ in issues if key == "AccessVBOM"]
        if not vbom_apps:
            log_info("   ✅ AccessVBOM ist für alle geprüften Apps bereits gesetzt.")
        elif ask(f"   AccessVBOM für {', '.join(vbom_apps)} (Office {ver}) automatisch auf 1 setzen? [j/N] ",
                 default=False):
            _set_office_access_vbom(ver, vbom_apps)
        else:
            log_info("   ℹ️ AccessVBOM wurde nicht geändert.")
            STATE["manual_actions"].append(
                f"AccessVBOM in Office {ver} für {', '.join(vbom_apps)} manuell auf 1 setzen "
                f"(Trust Center: 'Zugriff auf das VBA-Projektobjektmodell vertrauen')."
            )

    print("")

# =============================================================================
# Python-Interpreter
# =============================================================================

def find_python_exe():
    own = sys.executable or ""
    own_name = os.path.basename(own).lower()
    is_frozen = getattr(sys, "frozen", False)

    if not is_frozen and own_name in ("python.exe", "pythonw.exe"):
        return own

    py_launcher = shutil.which("py")
    if py_launcher:
        r = run_cmd([py_launcher, "-3", "-c", "import sys; print(sys.executable)"], timeout=10)
        if r.ok:
            p = r.stdout.strip()
            if os.path.isfile(p):
                return p

    def _py_sort_key(p):
        return version_tuple_from_name(os.path.basename(os.path.dirname(p)), "python")

    localappdata = os.environ.get("LOCALAPPDATA", "")
    if localappdata:
        for sub in ("Python", os.path.join("Programs", "Python")):
            candidates = sorted(
                set(glob.glob(os.path.join(localappdata, sub, "*", "python.exe")) +
                    glob.glob(os.path.join(localappdata, sub, "pythoncore-*-64", "python.exe"))),
                key=_py_sort_key, reverse=True,
            )
            if candidates:
                return candidates[0]

    for prog in [os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")]:
        if not prog:
            continue
        candidates = sorted(
            glob.glob(os.path.join(prog, "Python*", "python.exe")),
            key=_py_sort_key, reverse=True,
        )
        if candidates:
            return candidates[0]

    return shutil.which("python")

def _appx_install_cmd(installer):
    """PowerShell-Aufruf, der Fehler auch wirklich als Exit-Code meldet.

    'powershell -Command Add-AppxPackage ...' liefert sonst auch beim
    Scheitern des Cmdlets Exit-Code 0.
    """
    quoted = str(installer).replace("'", "''")
    ps_cmd = ("$ErrorActionPreference='Stop'; "
              f"try {{ Add-AppxPackage -Path '{quoted}' -ErrorAction Stop; exit 0 }} "
              "catch { Write-Error $_.Exception.Message; exit 1 }")
    return ["powershell", "-ExecutionPolicy", "Bypass", "-NoProfile",
            "-NonInteractive", "-Command", ps_cmd]

def install_python_manager(installer):
    if not IS_WINDOWS:
        return False
    log_info(f"   Installiere Python-Manager: {os.path.basename(installer)}")
    cmd = _appx_install_cmd(installer)
    r = run_cmd(cmd, timeout=600)
    if r.ok:
        log_info("   ✅ Python-Manager installiert.")
        STATE["installed_tools"].append("python-manager")
        return True
    log_warn(f"   ⚠️ Python-Manager-Installation fehlgeschlagen (Exit {r.returncode}).")
    if r.stderr.strip():
        log_warn(f"       {r.stderr.strip().splitlines()[-1]}")
    return False

def install_python_classic(installer):
    log_info(f"   Installiere Python (klassisch): {os.path.basename(installer)}")
    cmd = [installer, "/quiet",
           "InstallAllUsers=0", "PrependPath=1",
           "Include_launcher=1", "Include_pip=1",
           "Include_test=0", "Include_doc=0",
           "InstallLauncherAllUsers=0"]
    r = run_cmd(cmd, timeout=900)
    if r.exe_ok:
        rc_info = "" if r.returncode == 0 else f" (Exit {r.returncode}, Reboot empfohlen)"
        log_info(f"   ✅ Python installiert{rc_info}.")
        STATE["installed_tools"].append("python")
        if r.returncode == 3010:
            STATE["manual_actions"].append("Windows-Neustart wird vom Python-Installer empfohlen.")
        return True
    log_err(f"   ❌ Python-Installation fehlgeschlagen (Exit {r.returncode}).")
    return False

def install_python_from_deploy():
    for pat in PYTHON_INSTALLERS["manager_patterns"]:
        hits = find_in_deploy(pat)
        if hits and install_python_manager(hits[0]):
            return True

    for pat in PYTHON_INSTALLERS["classic_patterns"]:
        hits = find_in_deploy(pat)
        hits = [h for h in hits
                if not any(ex in os.path.basename(h).lower()
                           for ex in PYTHON_INSTALLERS["classic_exclude"])]
        if hits and install_python_classic(hits[0]):
            return True

    log_err("   ❌ Kein Python-Installer in der Deploy-Quelle gefunden.")
    return False

def ensure_python(deploy_available):
    global PYTHON_EXE
    PYTHON_EXE = find_python_exe()
    if PYTHON_EXE:
        log_info(f"✅ Ziel-Python-Interpreter: {PYTHON_EXE}")
        print("")
        return True

    log_warn("❌ Kein Python-Interpreter erkannt.")
    if not deploy_available:
        log_warn("   Keine Deploy-Quelle verfügbar – Python kann nicht automatisch installiert werden.")
        log_download_hint("python")
        print("")
        return False

    log_info("   Versuche automatische Python-Installation aus Deploy-Quelle...")
    if install_python_from_deploy():
        refresh_env_from_registry()
        PYTHON_EXE = find_python_exe()
        if PYTHON_EXE:
            log_info(f"✅ Ziel-Python-Interpreter: {PYTHON_EXE}")
            print("")
            return True

    log_err("❌ Python-Installation fehlgeschlagen – Python-bezogene Prüfungen werden übersprungen.")
    log_download_hint("python")
    print("")
    return False

# =============================================================================
# Python-Pakete (Office + OCR)
# =============================================================================

def module_installed_in_target(imp):
    if not PYTHON_EXE:
        return None
    # find_spec statt "import": deutlich schneller (ocrmypdf/pikepdf brauchen
    # beim echten Import mehrere Sekunden) und ohne Import-Nebenwirkungen.
    probe = ("import importlib.util, sys\n"
             "try:\n"
             f"    sys.exit(0 if importlib.util.find_spec({imp!r}) else 1)\n"
             "except Exception:\n"
             "    sys.exit(1)\n")
    r = run_cmd([PYTHON_EXE, "-c", probe], timeout=60)
    if r.not_found or r.timed_out:
        return None
    return r.returncode == 0

def _pip_extra_args():
    args = []
    wheels = find_wheels_dirs()
    for w in wheels:
        args += ["--find-links", w]
    return args

def install_python_package(pkg):
    if not PYTHON_EXE:
        log_err(f"   ❌ Kein Ziel-Python vorhanden, '{pkg}' kann nicht installiert werden.")
        return False
    log_info(f"   Installiere '{pkg}' via pip im Ziel-Python...")
    args = [PYTHON_EXE, "-m", "pip", "install", "--upgrade",
            "--disable-pip-version-check", "--no-input",
            "--retries", "2", "--timeout", "30", pkg]
    args += _pip_extra_args()

    r = run_cmd(args, timeout=600)
    if r.not_found or r.timed_out:
        log_err(f"   ❌ pip-Aufruf für '{pkg}' fehlgeschlagen ({r.error or 'unbekannt'}).")
        return False
    if r.ok:
        log_info(f"   ✅ '{pkg}' erfolgreich installiert.")
        STATE["installed_packages"].append(pkg)
        return True
    last = ((r.stderr or r.stdout) or "").strip().splitlines()
    log_err(f"   ❌ pip-Installation von '{pkg}' fehlgeschlagen (Exit {r.returncode}).")
    if last:
        log_err(f"      {last[-1]}")
    return False

def check_and_install_package(pkg, imp):
    status = module_installed_in_target(imp)
    if status is True:
        log_info(f"✅ Python-Modul '{pkg}' ist installiert.")
        return True
    if status is None:
        log_err(f"❌ Ziel-Python nicht erreichbar – '{pkg}' kann nicht geprüft werden.")
        return False
    log_warn(f"❌ Python-Modul '{pkg}' fehlt.")
    if install_python_package(pkg):
        return module_installed_in_target(imp) is True
    return False

def check_pywin32():
    log_info("--- Prüfe pywin32 Konfiguration ---")
    if not PYTHON_EXE:
        log_err("❌ Kein Ziel-Python gefunden, pywin32-Prüfung nicht möglich.\n")
        return False

    probe = "import win32api, win32com.client, pythoncom, pywintypes; print(pywintypes.__file__)"
    r = run_cmd([PYTHON_EXE, "-c", probe], timeout=30)

    if r.ok:
        log_info("✅ pywin32 Basis-Module importierbar.")
        out = (r.stdout or "").strip().splitlines()
        if out:
            log_info(f"✅ pywintypes DLL-Pfad: {out[-1]}")
        log_info("✅ pywin32 korrekt installiert.\n")
        return True

    log_warn("❌ pywin32 fehlt oder ist fehlerhaft konfiguriert.")
    if not install_python_package("pywin32"):
        print("")
        return False

    py_dir = os.path.dirname(PYTHON_EXE)
    candidates = [
        os.path.join(py_dir, "Scripts", "pywin32_postinstall.py"),
        os.path.join(py_dir, "Lib", "site-packages", "pywin32_system32", "pywin32_postinstall.py"),
        os.path.join(py_dir, "Lib", "site-packages", "win32", "scripts", "pywin32_postinstall.py"),
    ]
    # Bei "pip install --user" liegt das Skript unter %APPDATA%\Python\PythonXY\...
    rs = run_cmd([PYTHON_EXE, "-c",
                  "import sysconfig;print(sysconfig.get_path('scripts'));"
                  "print(sysconfig.get_path('purelib'));"
                  "print(sysconfig.get_path('scripts', 'nt_user'));"
                  "print(sysconfig.get_path('purelib', 'nt_user'))"], timeout=30)
    if rs.ok:
        for line in rs.stdout.splitlines():
            base = line.strip()
            if not base:
                continue
            candidates.append(os.path.join(base, "pywin32_postinstall.py"))
            candidates.append(os.path.join(base, "pywin32_system32", "pywin32_postinstall.py"))
            candidates.append(os.path.join(base, "win32", "scripts", "pywin32_postinstall.py"))

    for cand in candidates:
        if os.path.isfile(cand):
            log_info("   Führe pywin32_postinstall.py aus...")
            rp = run_cmd([PYTHON_EXE, cand, "-install"], timeout=120)
            if rp.ok:
                log_info("   ✅ pywin32 post-install erfolgreich.")
            elif not is_admin():
                STATE["admin_pending"].append({
                    "tool":      "pywin32_post",
                    "installer": cand,
                    "python":    PYTHON_EXE,
                })
                log_warn("   ⚠️ pywin32 post-install für Stage 2 (Admin) vorgemerkt.")
            else:
                log_warn("   ⚠️ pywin32 post-install fehlgeschlagen (trotz Adminrechten).")
                STATE["manual_actions"].append(
                    f"pywin32 post-install als Administrator ausführen: {PYTHON_EXE} {cand} -install"
                )
            break

    r2 = run_cmd([PYTHON_EXE, "-c", probe], timeout=30)
    print("")
    return r2.ok

def check_office_dependencies():
    log_info("--- Prüfe Python-Abhängigkeiten für Office-Konvertierung ---")
    all_ok = True
    for pkg, imp in OFFICE_PACKAGES.items():
        if not check_and_install_package(pkg, imp):
            all_ok = False
    if all_ok:
        log_info("✅ Alle Office-Module vorhanden.\n")
    else:
        log_warn("❌ Es fehlen noch Office-Module.\n")

def check_ocr_dependencies():
    log_info("--- Prüfe Python-Abhängigkeiten für OCR von PDFs ---")
    all_ok = True
    for pkg, imp in OCR_PACKAGES.items():
        if not check_and_install_package(pkg, imp):
            all_ok = False
    if all_ok:
        log_info("✅ Alle OCR-Module vorhanden.\n")
    else:
        log_warn("❌ Es fehlen noch OCR-Module.\n")

# =============================================================================
# Externe System-Tools – Installer-Ausführung
# =============================================================================

def run_installer(installer_path, cfg):
    ext = os.path.splitext(installer_path)[1].lower()
    log_info(f"   Führe Installer aus: {os.path.basename(installer_path)}")

    if ext == ".msi":
        args = cfg.get("silent_msi") or ["/quiet", "/norestart"]
        cmd = ["msiexec", "/i", installer_path] + list(args)
        r = run_cmd(cmd, timeout=1200)
        ok = r.msi_ok
    elif ext == ".msix":
        r = run_cmd(_appx_install_cmd(installer_path), timeout=1200)
        ok = r.ok
    else:
        args = cfg.get("silent_exe") or ["/S"]
        cmd = [installer_path] + list(args)
        r = run_cmd(cmd, timeout=1200)
        ok = r.exe_ok

    if r.timed_out:
        log_err(f"   ❌ Installer-Timeout: {os.path.basename(installer_path)}")
        return False
    if r.not_found:
        log_err(f"   ❌ Installer nicht ausführbar: {os.path.basename(installer_path)}")
        return False
    if ok:
        rc_info = "" if r.returncode == 0 else f" (Exit {r.returncode})"
        log_info(f"   ✅ Installation erfolgreich{rc_info}: {os.path.basename(installer_path)}")
        if r.returncode == 3010:
            log_warn("   ⚠️ Neustart des Rechners wird empfohlen (Installer meldet Reboot Required).")
            STATE["manual_actions"].append("Windows-Neustart (durch Installer angefordert).")
        return True
    log_err(f"   ❌ Installer-Exit-Code {r.returncode}: {os.path.basename(installer_path)}")
    err = (r.stderr or r.stdout or "").strip().splitlines()
    if err:
        log_err(f"      {err[-1]}")
    return False

def pick_installer_in_deploy(tool_key):
    cfg = TOOL_INSTALLERS.get(tool_key) or {}
    for pat in cfg.get("patterns", []):
        hits = find_in_deploy(pat)
        if hits:
            return hits[0], cfg
    return None, cfg

# ENTFERNT: pick_installer_in_stage(tool_key)
#
# Die Funktion suchte im Stage-Ordner per Dateinamen-Muster nach einem
# Installer und lieferte den ERSTEN Treffer aus os.listdir. Genutzt wurde
# sie in Stage 2, also im Admin-Kontext - und zwar ohne Abgleich mit der
# Pruefsumme aus dem Manifest. Damit konnten die gepruefte und die
# tatsaechlich ausgefuehrte Datei auseinanderfallen.
#
# Stage 2 installiert jetzt ausschliesslich ueber stage2_install_pending()
# und fuehrt dort genau den Pfad aus, dessen SHA-256 zuvor gegen das
# Manifest geprueft wurde. Eine Suche nach Namensmuster gibt es im
# Admin-Kontext nicht mehr.

def stage_admin_tool(tool_key, deploy_available):
    if not deploy_available:
        return None
    installer, cfg = pick_installer_in_deploy(tool_key)
    if not installer:
        log_warn(f"   ⚠️ Kein Installer für '{tool_key}' unter {DEPLOY_SOURCE} gefunden.")
        return None
    staged = stage_copy_file(installer, label=tool_key)
    if staged:
        STATE["admin_pending"].append({
            "tool": tool_key,
            "installer": staged,
        })
    return staged

def try_install_tool(tool_key, deploy_available):
    cfg = TOOL_INSTALLERS.get(tool_key) or {}
    needs_admin = cfg.get("admin", False)

    if RUN_MODE == "stage2":
        # In Stage 2 wird ausschliesslich ueber stage2_install_pending()
        # installiert - dort steht die Pruefsumme aus dem Manifest zur
        # Verfuegung und genau die gepruefte Datei wird ausgefuehrt.
        #
        # Ein Namensmuster-Suchlauf im Stage-Ordner waere hier eine Luecke:
        # er koennte eine andere, ungepruefte Datei erwischen und sie mit
        # Adminrechten starten. Deshalb keine Ausfuehrung an dieser Stelle.
        log_warn(f"   ⚠️ '{tool_key}' wird in Stage 2 nur über das geprüfte "
                 f"Manifest installiert – kein ungeprüfter Suchlauf.")
        return False

    if needs_admin and not is_admin():
        staged = stage_admin_tool(tool_key, deploy_available)
        if staged:
            log_warn(f"   ⚠️ Auto-Install von '{tool_key}' benötigt Admin-Rechte – wird in Stage 2 ausgeführt.")
        else:
            STATE["manual_actions"].append(
                f"Als Administrator erneut ausführen, um '{tool_key}' zu installieren."
            )
        return False

    if not deploy_available:
        return False
    installer, cfg = pick_installer_in_deploy(tool_key)
    if not installer:
        log_warn(f"   ⚠️ Kein Installer für '{tool_key}' unter {DEPLOY_SOURCE} gefunden.")
        return False
    ok = run_installer(installer, cfg)
    if ok:
        STATE["installed_tools"].append(tool_key)
        refresh_env_from_registry()
    return ok

# =============================================================================
# Externe Tools – Erkennung (PATH / Program Files / LOCALAPPDATA)
# =============================================================================

def cache_tool(name, aliases=None):
    for a in (aliases or [name]):
        p = shutil.which(a)
        if p:
            TOOLS_FOUND[name] = p
            return p
    TOOLS_FOUND[name] = None
    return None

def _program_dirs():
    out = []
    for v in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        p = os.environ.get(v, "")
        if p and p not in out:
            out.append(p)
    return out

def _user_install_dirs():
    out = []
    la = os.environ.get("LOCALAPPDATA", "")
    if la:
        out.append(la)
        out.append(os.path.join(la, "Programs"))
    ap = os.environ.get("APPDATA", "")
    if ap:
        out.append(ap)
    return out

def _search_fallback_exe(subpaths):
    for base in _program_dirs() + _user_install_dirs():
        if not base or not os.path.isdir(base):
            continue
        for sub in subpaths:
            cand = os.path.join(base, sub)
            matches = glob.glob(cand)
            for m in matches:
                if os.path.isfile(m):
                    return m
    return None

def check_system_tools(deploy_available):
    log_info("--- Prüfe externe System-Tools ---")
    cache_tool("tesseract")
    cache_tool("ghostscript", ["gswin64c", "gswin32c"])
    cache_tool("java")
    cache_tool("verapdf")
    cache_tool("jbig2")
    cache_tool("pngquant")
    cache_tool("unpaper")

    _check_tesseract(deploy_available)
    _check_ghostscript(deploy_available)
    _check_single_binaries(deploy_available)
    _check_leptonica(deploy_available)
    _check_java(deploy_available)
    _check_verapdf(deploy_available)
    print("")

# =============================================================================
# Externe Tools – Tesseract
# =============================================================================

def _detect_missing_langs_at(tess_exe, tessdata_dir=None):
    extra = {}
    if tessdata_dir:
        extra["TESSDATA_PREFIX"] = tessdata_dir
    r = run_cmd([tess_exe, "--list-langs"], timeout=15, extra_env=extra or None)
    if not r.ok and not r.stderr and not r.stdout:
        return list(REQUIRED_TESS_LANGS)
    out = (r.stdout or "") + "\n" + (r.stderr or "")
    installed = [line.strip() for line in out.splitlines()
                 if re.match(r"^[a-z]{3}(_[A-Za-z]+)?$", line.strip())]
    return [l for l in REQUIRED_TESS_LANGS if l not in installed]

def _tesseract_version(tess_exe):
    r = run_cmd([tess_exe, "--version"], timeout=15)
    text = ((r.stdout or "") + "\n" + (r.stderr or "")).lower()
    m = re.search(r"tesseract\s+v?(\d+)\.(\d+)(?:\.(\d+))?", text)
    if not m:
        return None
    return tuple(int(g or 0) for g in m.groups())

def _tesseract_fallback_paths():
    return _search_fallback_exe([
        os.path.join("Tesseract-OCR", "tesseract.exe"),
        os.path.join("Programs", "Tesseract-OCR", "tesseract.exe"),
    ])

def _tessdata_writable_target(tess_exe):
    primary = os.path.join(os.path.dirname(tess_exe), "tessdata")
    try:
        os.makedirs(primary, exist_ok=True)
        testfile = os.path.join(primary, ".write_test_vorabcheck")
        with open(testfile, "w", encoding="utf-8") as f:
            f.write("x")
        os.remove(testfile)
        return primary, False
    except Exception as _e:
        logging.debug(f"_tessdata_writable_target: Exception verworfen: {_e!r}")
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        alt = os.path.join(local, "Tesseract-OCR", "tessdata")
        try:
            os.makedirs(alt, exist_ok=True)
            if alt not in STATE["created_dirs"]:
                STATE["created_dirs"].append(alt)
            return alt, True
        except Exception as e:
            log_err(f"   ❌ Konnte tessdata-Fallback nicht anlegen: {e}")
    return primary, False

def _check_tesseract(deploy_available):
    tess_exe = TOOLS_FOUND.get("tesseract")
    if not tess_exe:
        fb = _tesseract_fallback_paths()
        if fb:
            append_user_path(os.path.dirname(fb))
            TOOLS_FOUND["tesseract"] = fb
            tess_exe = fb

    if not tess_exe:
        log_warn("❌ Tesseract nicht im PATH gefunden.")
        if try_install_tool("tesseract", deploy_available):
            fb = _tesseract_fallback_paths()
            if fb:
                append_user_path(os.path.dirname(fb))
                TOOLS_FOUND["tesseract"] = fb
                tess_exe = fb
        if not tess_exe:
            log_download_hint("tesseract")
            return

    log_info(f"✅ Tesseract gefunden: {tess_exe}")

    ver = _tesseract_version(tess_exe)
    if ver:
        log_info(f"✅ Tesseract-Version: {'.'.join(str(v) for v in ver if v is not None)}")
        if ver[0] < 5:
            log_warn(f"⚠️ Tesseract {ver[0]}.x erkannt – empfohlen ist 5.x oder neuer.")

    tess_dir = os.path.dirname(tess_exe)
    primary_tessdata = os.path.join(tess_dir, "tessdata")

    tdp = (os.environ.get("TESSDATA_PREFIX") or "").strip('"\'').strip()
    tessdata_dir = primary_tessdata

    if tdp:
        if os.path.isdir(tdp):
            tessdata_dir = tdp
            log_info(f"✅ TESSDATA_PREFIX: '{tdp}'")
        elif os.path.isdir(os.path.join(tdp, "tessdata")):
            tessdata_dir = os.path.join(tdp, "tessdata")
            log_info(f"✅ TESSDATA_PREFIX (unter tessdata/): '{tessdata_dir}'")
        else:
            log_warn(f"⚠️ TESSDATA_PREFIX verweist auf ungültigen Pfad: '{tdp}' – wird überschrieben.")

    missing = _detect_missing_langs_at(tess_exe)
    use_fallback = False
    if missing:
        write_target, use_fallback = _tessdata_writable_target(tess_exe)
        if use_fallback and os.path.isdir(primary_tessdata):
            for f in glob.glob(os.path.join(glob.escape(primary_tessdata), "*.traineddata")):
                tgt = os.path.join(write_target, os.path.basename(f))
                if not os.path.isfile(tgt):
                    copy_file_safe(f, tgt, silent=True)
        tessdata_dir = write_target

        if deploy_available:
            for lang in list(missing):
                hits = find_in_deploy(f"{lang}.traineddata")
                if hits and copy_file_safe(hits[0], os.path.join(tessdata_dir, f"{lang}.traineddata")):
                    missing.remove(lang)

        if use_fallback:
            set_user_env("TESSDATA_PREFIX", tessdata_dir)
            log_info(f"✅ TESSDATA_PREFIX auf Fallback gesetzt: '{tessdata_dir}'")
        else:
            if not tdp or not os.path.isdir(tdp):
                set_user_env("TESSDATA_PREFIX", tessdata_dir)
                log_info(f"✅ TESSDATA_PREFIX gesetzt: '{tessdata_dir}'")

        missing = _detect_missing_langs_at(tess_exe, tessdata_dir)
    else:
        if not tdp and os.path.isdir(primary_tessdata):
            set_user_env("TESSDATA_PREFIX", primary_tessdata)
            log_info(f"✅ TESSDATA_PREFIX neu gesetzt: '{primary_tessdata}'")

    if missing:
        log_warn(f"❌ Fehlende Tesseract-Sprachen: {', '.join(missing)}")
        log_warn(f"   Ziel: '{tessdata_dir}'")
        log_download_hint("tessdata")
    else:
        log_info(f"✅ Alle Tesseract-Sprachen vorhanden: {', '.join(REQUIRED_TESS_LANGS)}")

# =============================================================================
# Externe Tools – Ghostscript
# =============================================================================

def _ghostscript_fallback_paths():
    for base in _program_dirs() + _user_install_dirs():
        if not base:
            continue
        gs_base = os.path.join(base, "gs")
        if not os.path.isdir(gs_base):
            continue
        # Numerisch sortieren: lexikografisch stuende "gs9.56" vor "gs10.03".
        for entry in sorted(os.listdir(gs_base), key=natural_version_key, reverse=True):
            gs_bin = os.path.join(gs_base, entry, "bin")
            cands = glob.glob(os.path.join(glob.escape(gs_bin), "gswin*c.exe"))
            if cands:
                return cands[0]
    return None

def _compute_gs_lib(gs_exe):
    if not gs_exe:
        return []
    gs_root = os.path.dirname(os.path.dirname(gs_exe))
    parts = []
    for sub in ("lib", "Resource", os.path.join("Resource", "Init"), "fonts"):
        p = os.path.join(gs_root, sub)
        if os.path.isdir(p):
            parts.append(p)
    return parts

def _set_gs_lib(gs_exe):
    parts = _compute_gs_lib(gs_exe)
    if not parts:
        return
    desired = os.pathsep.join(parts)
    current = (os.environ.get("GS_LIB") or "").strip('"\'').strip()
    if not current:
        if set_user_env("GS_LIB", desired):
            log_info(f"✅ GS_LIB gesetzt: '{desired}'")
        return
    current_parts = [p.strip() for p in current.split(os.pathsep) if p.strip()]
    if not any(os.path.isdir(p) or os.path.isfile(p) for p in current_parts):
        log_warn(f"⚠️ GS_LIB enthält nur ungültige Pfade: '{current}' – wird überschrieben.")
        if set_user_env("GS_LIB", desired):
            log_info(f"✅ GS_LIB korrigiert: '{desired}'")
        return

    # GS_LIB ist gueltig – zeigt es aber auf eine ANDERE Ghostscript-Installation
    # als die gefundene EXE, passen Programmversion und Ressourcen nicht zusammen
    # (fehlende Fonts, Fehler in Resource/Init). Nicht stillschweigend
    # ueberschreiben: die Einstellung kann bewusst gesetzt worden sein.
    gs_root_norm = os.path.normcase(os.path.normpath(os.path.dirname(os.path.dirname(gs_exe))))

    def _under_gs_root(p):
        try:
            n = os.path.normcase(os.path.normpath(os.path.abspath(p)))
        except Exception:
            return False
        return n == gs_root_norm or n.startswith(gs_root_norm + os.sep)

    if all(_under_gs_root(p) for p in current_parts):
        log_info("✅ GS_LIB passt zur gefundenen Ghostscript-Installation.")
        return

    log_warn("⚠️ GS_LIB zeigt auf eine andere Ghostscript-Installation als die gefundene EXE.")
    log_warn(f"   aktuell:  {current}")
    log_warn(f"   passend:  {desired}")
    log_warn("   Unterschiedliche Versionen von Programm und Ressourcen führen zu")
    log_warn("   schwer auffindbaren Ghostscript-Fehlern (fehlende Fonts, Resource/Init).")
    if ask("   GS_LIB jetzt auf die gefundene Installation umstellen? [j/N] ", default=False):
        if set_user_env("GS_LIB", desired):
            log_info(f"✅ GS_LIB korrigiert: '{desired}'")
    else:
        log_info("   ℹ️ GS_LIB wurde nicht geändert.")
        STATE["manual_actions"].append(
            "GS_LIB prüfen: zeigt auf eine andere Ghostscript-Installation als die "
            f"erkannte EXE ({gs_exe}). Passender Wert wäre: {desired}"
        )

def _check_ghostscript(deploy_available):
    gs_exe = TOOLS_FOUND.get("ghostscript")
    if not gs_exe:
        fb = _ghostscript_fallback_paths()
        if fb:
            append_user_path(os.path.dirname(fb))
            TOOLS_FOUND["ghostscript"] = fb
            gs_exe = fb

    if not gs_exe:
        log_warn("❌ Ghostscript nicht im PATH gefunden.")
        if try_install_tool("ghostscript", deploy_available):
            fb = _ghostscript_fallback_paths()
            if fb:
                append_user_path(os.path.dirname(fb))
                TOOLS_FOUND["ghostscript"] = fb
                gs_exe = fb
        if not gs_exe:
            log_download_hint("ghostscript")
            return

    log_info(f"✅ Ghostscript gefunden: {gs_exe}")
    name = os.path.splitext(os.path.basename(gs_exe))[0].lower()
    if name.startswith("gswin64"):
        log_info("✅ Ghostscript-Architektur: 64-Bit")
    elif name.startswith("gswin32"):
        log_warn("⚠️ Ghostscript-Architektur: 32-Bit")
        if sys.maxsize > 2**32:
            log_warn("   Empfehlung: 64-Bit Ghostscript installieren (bei großen PDFs).")
            log_download_hint("ghostscript")
    r = run_cmd([gs_exe, "--version"], timeout=15)
    if r.ok:
        v = (r.stdout or "").strip()
        if v:
            log_info(f"✅ Ghostscript-Version: {v.splitlines()[0]}")

    _set_gs_lib(gs_exe)

# =============================================================================
# Externe Tools – Einzel-Binaries (jbig2, pngquant, unpaper)
# =============================================================================

def _check_single_binaries(deploy_available):
    for tool_name, cfg in TOOL_BINARIES.items():
        key = os.path.splitext(tool_name)[0]
        exe = TOOLS_FOUND.get(key)
        if exe:
            log_info(f"✅ {tool_name} gefunden: {exe}")
            continue
        local_candidate = os.path.join(cfg["target_dir"], tool_name)
        if os.path.isfile(local_candidate):
            append_user_path(cfg["target_dir"])
            TOOLS_FOUND[key] = local_candidate
            log_info(f"✅ {tool_name} gefunden (lokal): {local_candidate}")
            continue
        log_warn(f"❌ {tool_name} nicht im PATH gefunden.")
        if not deploy_available:
            log_download_hint(key)
            continue
        hits = find_in_deploy(tool_name)
        if not hits:
            log_warn(f"   Nicht in {DEPLOY_SOURCE} gefunden.")
            log_download_hint(key)
            continue
        source_exe = hits[0]
        source_dir = os.path.dirname(source_exe)
        target_dir = cfg["target_dir"]
        target_exe = os.path.join(target_dir, tool_name)
        if copy_file_safe(source_exe, target_exe, soft=not is_admin()):
            for dll in glob.glob(os.path.join(glob.escape(source_dir), "*.dll")):
                copy_file_safe(dll, os.path.join(target_dir, os.path.basename(dll)),
                               silent=True, soft=True)
            append_user_path(target_dir)
            TOOLS_FOUND[key] = target_exe
        elif not is_admin():
            # Kopieren scheiterte (vermutlich kein Schreibrecht unter C:\):
            # exe + zugehoerige DLLs fuer Stage 2 (Admin) vormerken.
            if stage_ocr_binary_for_admin(source_exe, target_exe, label=tool_name):
                for dll in glob.glob(os.path.join(glob.escape(source_dir), "*.dll")):
                    stage_ocr_binary_for_admin(dll, os.path.join(target_dir, os.path.basename(dll)),
                                               label=os.path.basename(dll))
                append_user_path(target_dir)
                log_warn(f"   ⚠️ {tool_name} für Stage 2 (Admin) nach {target_dir} vorgemerkt.")

# =============================================================================
# Externe Tools – Leptonica
# =============================================================================

def _lept_version_key(path):
    m = re.search(r"leptonica[-_]?(\d+)(?:\.(\d+))?(?:\.(\d+))?", os.path.basename(path).lower())
    if not m:
        return (0, 0, 0)
    return tuple(int(g or 0) for g in m.groups())

def _check_leptonica(deploy_available):
    path_dirs = [p.strip().strip('"\'') for p in os.environ.get("PATH", "").split(os.pathsep) if p.strip()]
    tess_exe = TOOLS_FOUND.get("tesseract")
    if tess_exe:
        td = os.path.dirname(tess_exe)
        if td not in path_dirs:
            path_dirs.insert(0, td)
    search_dirs = list(path_dirs) + [OCR_DIR]

    hits = []
    for d in search_dirs:
        if not d or not os.path.isdir(d):
            continue
        for pat in ["leptonica-*.dll", "*leptonica*.dll", "liblept*.dll"]:
            hits.extend(glob.glob(os.path.join(glob.escape(d), pat)))
    hits = sorted(set(hits), key=_lept_version_key, reverse=True)

    if hits:
        found = hits[0]
        log_info(f"✅ Leptonica DLL gefunden: {found}")
        target = os.path.join(OCR_DIR, os.path.basename(found))
        if not same_file(found, target) and not os.path.isfile(target):
            copy_file_safe(found, target)
        return

    if deploy_available:
        d_hits = []
        for pat in ["leptonica-*.dll", "*leptonica*.dll", "liblept*.dll"]:
            d_hits.extend(find_in_deploy(pat))
        d_hits = sorted(set(d_hits), key=_lept_version_key, reverse=True)
        if d_hits:
            target = os.path.join(OCR_DIR, os.path.basename(d_hits[0]))
            if copy_file_safe(d_hits[0], target, soft=not is_admin()):
                log_info(f"✅ Leptonica DLL kopiert nach: {target}")
                return
            if not is_admin() and stage_ocr_binary_for_admin(d_hits[0], target,
                                                             label=os.path.basename(d_hits[0])):
                log_warn(f"   ⚠️ Leptonica DLL für Stage 2 (Admin) nach {OCR_DIR} vorgemerkt.")
                return

    log_warn("❌ Leptonica DLL nicht gefunden.")
    log_download_hint("leptonica")

# =============================================================================
# Externe Tools – Java
# =============================================================================

def _java_fallback_paths():
    for base in _program_dirs() + _user_install_dirs():
        if not base:
            continue
        for vendor in ["Eclipse Adoptium", "Java", "Zulu", "Microsoft", "Amazon Corretto"]:
            vdir = os.path.join(base, vendor)
            if not os.path.isdir(vdir):
                continue
            # Numerisch sortieren: lexikografisch stuende "jdk-8" vor "jdk-21".
            entries = sorted(os.listdir(vdir), key=natural_version_key, reverse=True)
            for s in entries:
                jh = os.path.join(vdir, s)
                j = os.path.join(jh, "bin", "java.exe")
                if os.path.isfile(j):
                    return j, jh
    return None, None

def _check_java(deploy_available):
    java_exe  = TOOLS_FOUND.get("java")
    java_home = (os.environ.get("JAVA_HOME") or "").strip('"\'').strip()

    if not java_exe:
        fb, jh = _java_fallback_paths()
        if fb:
            set_user_env("JAVA_HOME", jh)
            append_user_path(os.path.join(jh, "bin"))
            TOOLS_FOUND["java"] = fb
            java_exe  = fb
            java_home = jh

    if not java_exe:
        log_warn("❌ Java nicht im PATH gefunden.")
        if try_install_tool("java", deploy_available):
            fb, jh = _java_fallback_paths()
            if fb:
                set_user_env("JAVA_HOME", jh)
                append_user_path(os.path.join(jh, "bin"))
                TOOLS_FOUND["java"] = fb
                java_exe  = fb
                java_home = jh
        if not java_exe:
            log_download_hint("java")
            return

    log_info(f"✅ Java gefunden: {java_exe}")

    def _looks_like_java_home(p):
        return bool(p) and os.path.isdir(p) and os.path.isfile(os.path.join(p, "bin", "java.exe"))

    if java_home:
        if _looks_like_java_home(java_home):
            log_info(f"✅ JAVA_HOME: '{java_home}'")
        else:
            log_warn(f"⚠️ JAVA_HOME zeigt auf ungültigen Pfad: '{java_home}'")
            new_jh = os.path.dirname(os.path.dirname(java_exe))
            if not _looks_like_java_home(new_jh):
                # z. B. Oracle java8path-Shim: dirname/dirname ist kein echtes JAVA_HOME
                _, fb_jh = _java_fallback_paths()
                new_jh = fb_jh if _looks_like_java_home(fb_jh) else None
            if new_jh:
                set_user_env("JAVA_HOME", new_jh)
                log_info(f"✅ JAVA_HOME korrigiert: '{new_jh}'")
                java_home = new_jh
            else:
                log_warn("   ⚠️ Kein gültiges JAVA_HOME automatisch ermittelbar – bitte manuell setzen.")
                STATE["manual_actions"].append(
                    "JAVA_HOME manuell auf das Installationsverzeichnis eines JDK/JRE setzen "
                    "(muss bin\\java.exe enthalten)."
                )
    else:
        new_jh = os.path.dirname(os.path.dirname(java_exe))
        if not _looks_like_java_home(new_jh):
            _, fb_jh = _java_fallback_paths()
            new_jh = fb_jh if _looks_like_java_home(fb_jh) else None
        if new_jh:
            set_user_env("JAVA_HOME", new_jh)
            log_info(f"✅ JAVA_HOME neu gesetzt: '{new_jh}'")
            java_home = new_jh

    r = run_cmd([java_exe, "-version"], timeout=30)
    out = ((r.stderr or "") + "\n" + (r.stdout or "")).strip()
    if out:
        first = out.splitlines()[0]
        log_info(f"✅ Java-Version: {first}")
        low = out.lower()
        is_64 = any(m in low for m in ["64-bit", "x86_64", "amd64", "aarch64"])
        if not is_64 and sys.maxsize > 2**32:
            log_warn("⚠️ 32-Bit Java auf 64-Bit System – 64-Bit JRE/JDK empfohlen.")

        m = re.search(r'version "(\d+)(?:\.(\d+))?', out)
        if m:
            major = int(m.group(1))
            # Legacy-Schema "1.X.Y" (Java ≤ 8) -> Major ist das zweite Feld
            if major == 1 and m.group(2):
                major = int(m.group(2))
            if major < 11:
                log_warn(f"⚠️ Java {major} ist alt (JRE 8 nur Fallback). Empfohlen: JDK 17 oder neuer.")

    if java_home:
        path_java = os.path.normcase(os.path.realpath(java_exe))
        home_java_file = os.path.join(java_home, "bin", "java.exe")
        if os.path.isfile(home_java_file):
            home_java = os.path.normcase(os.path.realpath(home_java_file))
            if path_java != home_java:
                low_path = path_java.replace("/", "\\").lower()
                is_oracle_shim = (
                    "\\oracle\\java\\javapath\\"   in low_path or
                    "\\oracle\\java\\java8path\\" in low_path or
                    "\\oracle\\java\\javafxpath\\" in low_path
                )
                if is_oracle_shim:
                    log_info("ℹ️ Java im PATH ist ein Oracle-Launcher-Stub "
                             "(java8path/javapath); ruft die in der Registry "
                             "eingetragene Default-Java-Installation auf.")
                else:
                    log_warn("⚠️ Java im PATH und JAVA_HOME zeigen auf unterschiedliche Installationen.")
                    log_warn(f"   PATH:      {java_exe}")
                    log_warn(f"   JAVA_HOME: {home_java_file}")

# =============================================================================
# Externe Tools – veraPDF
# =============================================================================

def _verapdf_candidate_paths():
    return [
        os.path.join(VERAPDF_DIR, "verapdf.bat"),
        os.path.join(VERAPDF_DIR, "verapdf-gui.bat"),
        os.path.join(r"C:\Program Files\verapdf", "verapdf.bat"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "verapdf", "verapdf.bat"),
    ]

def _install_verapdf_silent(deploy_available, entry=None):
    if RUN_MODE == "stage2":
        # Nur den im Manifest vermerkten - und damit von verify_staged_entry
        # geprueften - Pfad verwenden.
        #
        # Vorher wurde die .bat im Stage-Verzeichnis per Namensvergleich neu
        # gesucht, mit Rueckfall auf STAGE_ROOT, wenn der Unterordner fehlte.
        # verify_staged_dir prueft aber nur den im Manifest genannten Ordner:
        # bei einem Manifest ohne 'dir'-Angabe waere so eine ungepruefte
        # verapdf-install.bat aus STAGE_ROOT mit Adminrechten gestartet worden.
        bat = (entry or {}).get("installer")
        if not bat or not os.path.isfile(long_path(bat)):
            log_warn("   ⚠️ Geprüfter veraPDF-Installer aus dem Manifest nicht vorhanden.")
            return None
        search_dir = os.path.dirname(bat)
        xml_file = None
        for f in os.listdir(search_dir):
            if f.lower() == "auto-install.xml":
                xml_file = os.path.join(search_dir, f)
                break
        if not xml_file:
            log_warn("   ⚠️ auto-install.xml in Stage-Verzeichnis nicht gefunden.")
            return None
        if not TOOLS_FOUND.get("java") and not shutil.which("java"):
            log_err("   ❌ veraPDF-Installation benötigt Java – zuerst Java einrichten.")
            return None
        log_info(f"   Starte veraPDF-Silent-Install ({os.path.basename(bat)})")
        log_info(f"   Profil: {xml_file}")
        r = run_cmd([bat, xml_file], timeout=900, cwd=search_dir)
        if r.ok:
            log_info("   ✅ veraPDF installiert.")
            STATE["installed_tools"].append("verapdf")
            STATE["stage2_summary"].append("installed:verapdf")
            for cand in _verapdf_candidate_paths():
                if os.path.isfile(cand):
                    return cand
            return None
        log_err(f"   ❌ veraPDF-Silent-Install fehlgeschlagen (Exit {r.returncode}).")
        return None

    if not deploy_available:
        return None
    bat_hits = find_in_deploy("verapdf-install.bat")
    if not bat_hits:
        log_warn("   ⚠️ verapdf-install.bat nicht in Deploy-Quelle gefunden.")
        return None
    bat = bat_hits[0]
    src_dir = os.path.dirname(bat)

    xml_local = os.path.join(src_dir, "auto-install.xml")
    xml_hits = find_in_deploy("auto-install.xml") if not os.path.isfile(xml_local) else [xml_local]
    if not xml_hits:
        log_warn("   ⚠️ auto-install.xml nicht gefunden – interaktive Installation nötig.")
        STATE["manual_actions"].append(
            f"veraPDF manuell installieren: {bat}  (oder auto-install.xml bereitstellen)"
        )
        return None
    xml_file = xml_hits[0]

    if not is_admin():
        staged = stage_copy_dir(src_dir, "verapdf_installer")
        if staged:
            try:
                if os.path.dirname(os.path.abspath(xml_file)).lower() != os.path.abspath(src_dir).lower():
                    shutil.copy2(long_path(xml_file),
                                 long_path(os.path.join(staged, "auto-install.xml")))
            except Exception as e:
                log_err(f"   ❌ auto-install.xml konnte nicht bereitgestellt werden: {e}")
            STATE["admin_pending"].append({
                "tool":      "verapdf",
                "installer": os.path.join(staged, os.path.basename(bat)),
                # ganzer Ordner wird gehasht: die .bat startet Java mit dem
                # IzPack-JAR aus demselben Verzeichnis
                "dir":       staged,
            })
            log_warn("   ⚠️ veraPDF-Installation für Stage 2 (Admin) vorgemerkt.")
        else:
            log_err("   ❌ veraPDF-Installer-Verzeichnis konnte nicht für Stage 2 bereitgestellt werden.")
            STATE["manual_actions"].append(
                f"veraPDF manuell als Administrator installieren: {bat}"
            )
        return None

    if not TOOLS_FOUND.get("java") and not shutil.which("java"):
        log_err("   ❌ veraPDF-Installation benötigt Java – zuerst Java einrichten.")
        return None

    log_info(f"   Starte veraPDF-Silent-Install ({os.path.basename(bat)})")
    log_info(f"   Profil: {xml_file}")
    r = run_cmd([bat, xml_file], timeout=900, cwd=src_dir)
    if r.ok:
        log_info("   ✅ veraPDF installiert.")
        STATE["installed_tools"].append("verapdf")
        for cand in _verapdf_candidate_paths():
            if os.path.isfile(cand):
                return cand
        return None
    log_err(f"   ❌ veraPDF-Silent-Install fehlgeschlagen (Exit {r.returncode}).")
    err = (r.stderr or r.stdout or "").strip().splitlines()
    if err:
        log_err(f"      {err[-1]}")
    return None

def _check_verapdf(deploy_available):
    v_exe = TOOLS_FOUND.get("verapdf")

    if not v_exe:
        for cand in _verapdf_candidate_paths():
            if cand and os.path.isfile(cand):
                v_exe = cand
                TOOLS_FOUND["verapdf"] = cand
                append_user_path(os.path.dirname(cand))
                set_user_env("VERAPDF_HOME", os.path.dirname(cand))
                break

    if not v_exe:
        log_warn("❌ veraPDF nicht im PATH/Standard-Pfaden gefunden.")
        v_exe = _install_verapdf_silent(deploy_available)
        if v_exe:
            TOOLS_FOUND["verapdf"] = v_exe
            append_user_path(os.path.dirname(v_exe))
            set_user_env("VERAPDF_HOME", os.path.dirname(v_exe))

    if not v_exe:
        log_download_hint("verapdf")
        return

    log_info(f"✅ veraPDF gefunden: {v_exe}")
    r = run_cmd([v_exe, "--version"], timeout=60)
    if r.ok:
        v = (r.stdout.strip() or r.stderr.strip())
        if v:
            log_info(f"✅ veraPDF-Version: {v.splitlines()[0]}")

# =============================================================================
# ENV-Übersicht und Multi-User-PATH-Hinweis
# =============================================================================

def check_environment_summary():
    log_info("--- Übersicht: Relevante Umgebungsvariablen ---")
    env_vars = [
        ("PATH",             "Suchpfad für ausführbare Dateien",                 False),
        ("TESSDATA_PREFIX",  "Tesseract-Sprachdaten-Verzeichnis",                False),
        ("JAVA_HOME",        "Java-Installationsverzeichnis",                    False),
        ("GS_LIB",           "Ghostscript-Bibliothekspfad (optional)",           True),
        ("VERAPDF_HOME",     "veraPDF-Installationsverzeichnis (optional)",      True),
    ]
    for var, desc, optional in env_vars:
        value = (os.environ.get(var, "") or "").strip('"\'').strip()
        if var == "PATH":
            n = len([p for p in value.split(os.pathsep) if p.strip()])
            log_info(f"  {var:<16} ✅  {desc} ({n} Einträge)")
        elif not value:
            if optional:
                log_info(f"  {var:<16} ℹ️  {desc} (nicht gesetzt – optional)")
            else:
                log_warn(f"  {var:<16} ❌  {desc} (nicht gesetzt)")
        elif var == "GS_LIB" and os.pathsep in value:
            parts = [p.strip() for p in value.split(os.pathsep) if p.strip()]
            if parts and all(os.path.isdir(p) or os.path.isfile(p) for p in parts):
                log_info(f"  {var:<16} ✅  {desc} = {value}")
            else:
                log_warn(f"  {var:<16} ⚠️  {desc} = {value} (ungültig)")
        elif os.path.isdir(value) or os.path.isfile(value):
            log_info(f"  {var:<16} ✅  {desc} = {value}")
        else:
            log_warn(f"  {var:<16} ⚠️  {desc} = {value} (ungültig)")
    print("")
    _check_multi_user_path()

def _check_multi_user_path():
    current = get_current_username()
    if not current:
        return
    path_value = os.environ.get("PATH", "")
    path_dirs = [p.strip().strip('"\'') for p in path_value.split(os.pathsep) if p.strip()]
    users_root = (os.environ.get("SystemDrive", "C:") or "C:") + r"\Users"
    # (?:\\|$) statt \\ : sonst wird "C:\Users\meier" ohne Trailing-Backslash nicht erkannt.
    pattern = re.compile(r"^" + re.escape(users_root) + r"\\([^\\]+)(?:\\|$)", re.IGNORECASE)
    ignore = {"public", "default", "default user", "all users", "alluser"}
    foreign = []
    for d in path_dirs:
        m = pattern.match(d)
        if not m:
            continue
        user_in_path = m.group(1).lower()
        if user_in_path == current or user_in_path in ignore:
            continue
        foreign.append((d, user_in_path))

    if not foreign:
        return

    log_warn("")
    log_warn("⚠️ PATH enthält Einträge fremder Benutzerprofile, die für den aktuellen")
    log_warn(f"   Nutzer ({current}) nicht zugreifbar sind:")
    for p, u in foreign:
        log_warn(f"   - {p}   [Profil: {u}]")
    log_warn("")
    log_warn("   Diese Einträge werden NICHT automatisch entfernt. Manuelle Bereinigung:")
    log_warn("   1. Windows-Taste drücken, 'Umgebungsvariablen' tippen,")
    log_warn("      'Systemumgebungsvariablen bearbeiten' öffnen.")
    log_warn("   2. Im Fenster 'Systemeigenschaften' auf 'Umgebungsvariablen...' klicken.")
    log_warn("   3. Im unteren Block 'Systemvariablen' den Eintrag 'Path' markieren")
    log_warn("      und auf 'Bearbeiten...' klicken.")
    log_warn("   4. Die oben aufgelisteten Zeilen mit fremden Benutzernamen jeweils")
    log_warn("      markieren und mit 'Löschen' entfernen.")
    log_warn("   5. Dreimal mit 'OK' bestätigen.")
    log_warn("   6. Falls die Einträge stattdessen im oberen Block 'Benutzervariablen'")
    log_warn("      (unter 'Path') liegen, dort entsprechend entfernen.")
    log_warn("   7. Alle offenen Terminals / Editoren / VS Code neu starten.")
    STATE["manual_actions"].append(
        f"PATH manuell bereinigen – {len(foreign)} fremde Benutzerprofil-Pfade entfernen."
    )

# =============================================================================
# Stage-2-Spezial: Tools mit Stage-Kopie installieren
# =============================================================================

def stage2_install_pending(manifest):
    log_info("--- Stage 2: Installiere vorgemerkte Tools mit Adminrechten ---")
    pending = manifest.get("admin_pending", [])
    if not pending:
        log_info("   Keine ausstehenden Admin-Installationen.")
        return
    for entry in pending:
        tool = entry.get("tool")
        if not tool:
            continue
        if not verify_staged_entry(entry):
            STATE["manual_actions"].append(
                f"'{tool}' wurde NICHT installiert (Prüfsumme/Datei ungültig) – Stage 1 erneut ausführen."
            )
            continue
        if tool == "verapdf":
            # entry mitgeben: der Installer wird aus dem geprueften
            # Manifest-Pfad gestartet, nicht per Namenssuche.
            _install_verapdf_silent(deploy_available=False, entry=entry)
        elif tool == "ocr_binary":
            src = entry.get("installer")
            tgt = entry.get("target")
            if src and tgt:
                if copy_file_safe(src, tgt):
                    log_info(f"   ✅ {os.path.basename(tgt)} nach {os.path.dirname(tgt)} kopiert.")
                    STATE["stage2_summary"].append(f"installed:{os.path.basename(tgt)}")
                else:
                    log_err(f"   ❌ {os.path.basename(tgt)} konnte nicht nach {os.path.dirname(tgt)} kopiert werden.")
            else:
                log_warn(f"   ⚠️ Stage-Datei für {os.path.basename(str(tgt))} nicht gefunden.")
        elif tool == "pywin32_post":
            cand = entry.get("installer")
            py = entry.get("python") or PYTHON_EXE or find_python_exe()
            if not py:
                log_warn("   ⚠️ Kein Python-Interpreter für pywin32 post-install gefunden.")
            elif cand and os.path.isfile(cand):
                log_info("   Führe pywin32_postinstall.py (Admin) aus...")
                rp = run_cmd([py, cand, "-install"], timeout=120)
                if rp.ok:
                    log_info("   ✅ pywin32 post-install (Admin) erfolgreich.")
                    STATE["stage2_summary"].append("installed:pywin32_post")
                else:
                    log_err(f"   ❌ pywin32 post-install (Admin) fehlgeschlagen (Exit {rp.returncode}).")
            else:
                log_warn("   ⚠️ pywin32_postinstall.py nicht gefunden.")
        else:
            # WICHTIG: genau die Datei ausfuehren, die oben geprueft wurde.
            #
            # Vorher stand hier 'try_install_tool(tool, ...)'. Das suchte den
            # Installer im Stage-Ordner NEU ueber ein Dateinamen-Muster
            # (pick_installer_in_stage). Geprueft wurde damit entry['installer'],
            # ausgefuehrt aber, was os.listdir als erstes passend lieferte -
            # und diese Reihenfolge ist nicht zugesichert.
            #
            # Der Stage-Ordner ist zwar gehaertet, das aufrufende Konto behaelt
            # aber Schreibrecht (Modify). Wer dort zwischen Stage 1 und Stage 2
            # eine ZWEITE, ebenfalls zum Muster passende Datei ablegt, konnte
            # so an der Pruefsumme vorbei eine beliebige EXE mit Adminrechten
            # starten lassen - genau die lokale Rechteausweitung, gegen die
            # das ganze Stage-Modell gebaut ist.
            verified = entry.get("installer")
            cfg = TOOL_INSTALLERS.get(tool) or {}
            if not verified or not os.path.isfile(long_path(verified)):
                log_err(f"   ❌ Geprüfter Installer für '{tool}' nicht mehr vorhanden – übersprungen.")
                STATE["manual_actions"].append(
                    f"'{tool}' wurde NICHT installiert (Stage-Datei fehlt) – Stage 1 erneut ausführen."
                )
                continue
            if run_installer(verified, cfg):
                STATE["installed_tools"].append(tool)
                STATE["stage2_summary"].append(f"installed:{tool}")
    print("")

# =============================================================================
# Zusammenfassung
# =============================================================================

def print_summary():
    print("")
    log_info("=" * 70)
    if RUN_MODE == "stage2":
        log_info(" Zusammenfassung (Stage 2 – Admin)")
    else:
        log_info(" Zusammenfassung")
    log_info("=" * 70)

    def section(title, items):
        if items:
            log_info(f"\n {title}:")
            for i in items:
                log_info(f"   - {i}")

    section("Verzeichnisse angelegt",         STATE["created_dirs"])
    section("Python-Pakete installiert",      STATE["installed_packages"])
    section("Tools installiert",              STATE["installed_tools"])
    section("Dateien kopiert",                STATE["copied_files"])
    section("ENV-Variablen gesetzt",          STATE["env_set"])
    section("PATH-Einträge ergänzt",          STATE["path_added"])
    section("Fehler",                         STATE["errors"])
    section("Manuelle Aktionen erforderlich", STATE["manual_actions"])

    if RUN_MODE == "stage1":
        section("Für Admin-Lauf vorgemerkt",
                [f"{e['tool']} ({os.path.basename(e['installer'])})" for e in STATE["admin_pending"]])

    log_info("")
    log_info(f" Warnungen:  {len(STATE['warnings'])}")
    log_info(f" Fehler:     {len(STATE['errors'])}")

    # 'warnings' MUSS mit hinein. Es gibt Pfade, die ein fehlendes
    # Pflichtwerkzeug ausschliesslich als Warnung ablegen (z. B. veraPDF ohne
    # erfolgreiche Installation). Ohne diesen Eintrag meldete der Lauf
    # 'Alles bereits korrekt eingerichtet' und endete mit Exit-Code 0 -
    # obwohl ein benoetigtes Programm fehlt.
    nothing = not any(STATE[k] for k in (
        "created_dirs", "installed_packages", "installed_tools",
        "copied_files", "env_set", "path_added", "errors", "manual_actions",
        "admin_pending", "warnings",
    ))
    if nothing:
        log_info("\n ✅ Alles bereits korrekt eingerichtet – keine Änderungen nötig.")
    elif STATE["warnings"] and not (STATE["errors"] or STATE["manual_actions"]):
        log_info("")
        log_info(" ⚠️  Der Lauf ist ohne Fehler beendet, aber mit Warnungen –")
        log_info("    bitte die Liste oben durchsehen. Ein fehlendes Werkzeug")
        log_info("    fällt sonst erst im späteren Verarbeitungsschritt auf.")
    elif STATE["env_set"] or STATE["path_added"]:
        log_info("")
        log_info(" ℹ️  Hinweis: Offene Terminals/Editoren/VS Code neu starten,")
        log_info("    damit PATH- und ENV-Änderungen dort wirksam werden.")
        log_info("    Für systemweite Übernahme ggf. ab- und wieder anmelden.")

    log_info("")
    log_info(f" Logfile: {LOG_FILE}")

# =============================================================================
# Main – Stage-1 (User) und Stage-2 (Admin)
# =============================================================================

def main_stage1():
    print()
    print("╔══════════════════════════════════════════════════════════════════════╗")
    print("║          Vorab-Check und Installation fehlender Programme            ║")
    print("║                Stage 1 – Anwender (ohne Adminrechte)                 ║")
    print("╚══════════════════════════════════════════════════════════════════════╝")
    print(f"  Start:  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Python: {sys.version.split()[0]}")
    print(f"  System: {platform.system()} {platform.version()[:40]}")

    resolve_deploy_source()

    logging.info("=== Vorab-Check Stage 1 gestartet ===")
    logging.info(f"Python: {sys.version.split()[0]} | System: {platform.system()} {platform.version()[:40]}")
    logging.info(f"Deploy-Quelle: {DEPLOY_SOURCE}")

    print()
    if is_admin():
        log_info("✅ Skript läuft mit Administrator-Rechten – Stage 1 inkl. Admin-Schritte.")
    else:
        log_info("ℹ️ Skript läuft OHNE Administrator-Rechte (Two-Stage-Modus aktiv).")
        log_info("   Admin-pflichtige Tools werden in eine lokale Stage-Kopie ausgelagert")
        log_info("   und am Ende über UAC-Re-Launch (Stage 2) installiert.")
    log_info("")

    ensure_directories()
    deploy_available = check_deploy_source()

    check_office()

    ensure_python(deploy_available)

    if PYTHON_EXE:
        check_pywin32()
        check_office_dependencies()
        check_ocr_dependencies()

    check_system_tools(deploy_available)
    check_environment_summary()

    if STATE["admin_pending"] and not is_admin() and IS_WINDOWS:
        log_info("--- Stage-Kopien für Admin-Installation vorbereitet ---")
        for entry in STATE["admin_pending"]:
            log_info(f"   • {entry['tool']:<12} {os.path.basename(entry['installer'])}")
        print()
        if ask("Jetzt Stage 2 mit Adminrechten starten (UAC-Abfrage)? [J/n] ",
               default=not NO_INPUT):
            # Ergebnisdatei eines frueheren Laufs entfernen, sonst wuerde ein
            # abgebrochener Stage-2-Lauf mit altem Ergebnis "bestaetigt".
            try:
                if os.path.isfile(STAGE_RESULT_FILE):
                    os.remove(STAGE_RESULT_FILE)
            except Exception as _e:
                logging.debug(f"main_stage1: Exception verworfen: {_e!r}")

            harden_stage_dir()
            # Ohne belastbare Pruefsumme darf ein Eintrag NICHT ins Manifest:
            # Stage 2 wuerde ihn sonst mit Adminrechten ungeprueft ausfuehren.
            # sha256_of() liefert bei jedem Fehler None (gesperrte Datei,
            # Virenscanner, offenes Handle), hash_tree() ein leeres dict -
            # genau in dem Moment also, in dem mit der Stage-Datei etwas nicht
            # stimmt, faellt der Schutz weg, auf dem das ganze Two-Stage-Modell
            # beruht. Ein Schutz, der bei Fehlern aufmacht statt zumacht, ist
            # die falsche Richtung: solche Eintraege werden verworfen und als
            # Handarbeit gemeldet.
            geprueft = []
            for entry in STATE["admin_pending"]:
                name = os.path.basename(entry.get("installer") or entry.get("dir") or "?")
                digest = sha256_of(entry.get("installer"))
                if not digest:
                    log_err(f"   ❌ Prüfsumme für {name} nicht bildbar – Eintrag wird NICHT "
                            f"an Stage 2 übergeben (Datei gesperrt oder nicht lesbar?).")
                    STATE["manual_actions"].append(
                        f"{entry.get('tool', name)}: Prüfsumme nicht bildbar, Installation "
                        f"von Hand nachholen ({entry.get('installer')})."
                    )
                    continue
                entry["sha256"] = digest
                if entry.get("dir"):
                    files = hash_tree(entry["dir"])
                    unlesbar = sorted(k for k, v in files.items() if v == HASH_UNREADABLE)
                    if not files or unlesbar:
                        grund = ("keine Datei lesbar" if not files
                                 else f"{len(unlesbar)} Datei(en) nicht lesbar")
                        log_err(f"   ❌ Hash-Liste für Ordner '{name}' unvollständig "
                                f"({grund}) – Eintrag wird NICHT an Stage 2 übergeben.")
                        for k in unlesbar[:5]:
                            log_err(f"      nicht lesbar: {k}")
                        STATE["manual_actions"].append(
                            f"{entry.get('tool', name)}: Hash-Liste unvollständig, Installation "
                            f"von Hand nachholen ({entry.get('dir')})."
                        )
                        continue
                    entry["files"] = files
                geprueft.append(entry)
            STATE["admin_pending"] = geprueft

            manifest = {
                "deploy_source":  DEPLOY_SOURCE,
                "stage_root":     STAGE_ROOT,
                "admin_pending":  STATE["admin_pending"],
                "user":           get_current_username(),
                "created":        datetime.now().isoformat(timespec="seconds"),
            }
            if write_stage_manifest(manifest):
                log_info("   Starte UAC-Re-Launch für Stage 2 ...")
                rc = relaunch_as_admin_and_wait(STAGE_ROOT)
                if rc is None:
                    log_err("   ❌ Stage 2 wurde nicht gestartet (UAC verweigert/abgebrochen).")
                    STATE["manual_actions"].append(
                        f"Stage 2 als Administrator nachholen (Stage-Verzeichnis: {STAGE_ROOT})."
                    )
                elif rc == 0:
                    log_info("   ✅ Stage 2 erfolgreich beendet (Exit 0).")
                else:
                    log_warn(f"   ⚠️ Stage 2 beendet mit Exit-Code {rc} – Details siehe Stage-2-Ausgabe/Log.")
                refresh_env_from_registry()
                _post_stage2_recheck()
        else:
            STATE["manual_actions"].append(
                "Stage 2 später als Administrator nachholen "
                f"(Stage-Verzeichnis: {STAGE_ROOT})."
            )

    print_summary()

    if STATE["errors"]:
        return 2
    if STATE["manual_actions"]:
        return 1
    # Auch reine Warnungen sichtbar machen: sie koennen ein fehlendes
    # Pflichtwerkzeug bedeuten (veraPDF & Co.), und Exit-Code 0 haette dem
    # Aufrufer - Aufgabenplanung oder Wrapper-Skript - einen sauberen Lauf
    # gemeldet.
    if STATE["warnings"]:
        return 1
    return 0

def _post_stage2_recheck():
    log_info("")
    log_info("--- Re-Check nach Stage 2 ---")

    result = read_stage_result()
    if result:
        installed = result.get("installed") or []
        if installed:
            log_info(f"   Stage 2 hat installiert: {', '.join(installed)}")
            for t in installed:
                if t not in STATE["installed_tools"]:
                    STATE["installed_tools"].append(f"{t} (via Stage 2)")
        for err in (result.get("errors") or []):
            if err not in STATE["errors"]:
                STATE["errors"].append(f"[Stage 2] {err}")

    TOOLS_FOUND.clear()
    cache_tool("tesseract")
    cache_tool("ghostscript", ["gswin64c", "gswin32c"])
    cache_tool("java")
    cache_tool("verapdf")

    if not TOOLS_FOUND.get("verapdf"):
        for cand in _verapdf_candidate_paths():
            if cand and os.path.isfile(cand):
                TOOLS_FOUND["verapdf"] = cand
                break

    for tool, label in [("tesseract", "Tesseract"),
                        ("ghostscript", "Ghostscript"),
                        ("java", "Java"),
                        ("verapdf", "veraPDF")]:
        p = TOOLS_FOUND.get(tool)
        if p:
            log_info(f"   ✅ {label}: {p}")
        else:
            log_warn(f"   ⚠️ {label}: weiterhin nicht gefunden.")

    tess = TOOLS_FOUND.get("tesseract")
    if tess and not (os.environ.get("TESSDATA_PREFIX") or "").strip():
        primary = os.path.join(os.path.dirname(tess), "tessdata")
        if os.path.isdir(primary):
            set_user_env("TESSDATA_PREFIX", primary)

    java = TOOLS_FOUND.get("java")
    if java and not (os.environ.get("JAVA_HOME") or "").strip():
        new_jh = os.path.dirname(os.path.dirname(java))
        if os.path.isdir(new_jh):
            set_user_env("JAVA_HOME", new_jh)
            append_user_path(os.path.join(new_jh, "bin"))

    verapdf = TOOLS_FOUND.get("verapdf")
    if verapdf:
        append_user_path(os.path.dirname(verapdf))
        if not (os.environ.get("VERAPDF_HOME") or "").strip():
            set_user_env("VERAPDF_HOME", os.path.dirname(verapdf))

    gs = TOOLS_FOUND.get("ghostscript")
    if gs:
        _set_gs_lib(gs)

    print("")

def main_stage2():
    print()
    print("╔══════════════════════════════════════════════════════════════════════╗")
    print("║              Stage 2 – Administrator-Installation                    ║")
    print("║         (gestartet via UAC-Re-Launch durch Stage 1 / User)           ║")
    print("╚══════════════════════════════════════════════════════════════════════╝")
    print(f"  Start:  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Stage:  {STAGE_ROOT}")

    logging.info("=== Vorab-Check Stage 2 (Admin) gestartet ===")

    if not IS_WINDOWS:
        log_err("❌ Stage 2 ist nur unter Windows sinnvoll.")
        return 2

    if not is_admin():
        log_err("❌ Stage 2 wurde gestartet, aber der Prozess hat keine Admin-Rechte.")
        log_warn("   Bitte Skript regulär (Stage 1) starten und UAC-Re-Launch zulassen.")
        return 2

    manifest = read_stage_manifest()
    if not manifest:
        log_err("❌ Kein gültiges Stage-Manifest – Stage 2 wird abgebrochen.")
        return 2

    log_info(f"   Manifest erstellt: {manifest.get('created')}")
    log_info(f"   Ursprünglicher User: {manifest.get('user')}")
    log_info("")

    ensure_directories()

    cache_tool("tesseract")
    cache_tool("ghostscript", ["gswin64c", "gswin32c"])
    cache_tool("java")
    cache_tool("verapdf")

    stage2_install_pending(manifest)

    refresh_env_from_registry()
    cache_tool("tesseract")
    cache_tool("ghostscript", ["gswin64c", "gswin32c"])
    cache_tool("java")
    cache_tool("verapdf")

    for tool, dirpath in [
        ("tesseract",   _tesseract_fallback_paths()),
        ("ghostscript", _ghostscript_fallback_paths()),
    ]:
        if dirpath:
            append_system_path(os.path.dirname(dirpath))

    write_stage_result({
        "installed":      STATE["installed_tools"],
        "errors":         STATE["errors"],
        "warnings":       STATE["warnings"],
        "stage2_summary": STATE["stage2_summary"],
        "ended":          datetime.now().isoformat(timespec="seconds"),
    })

    print_summary()

    log_info("")
    log_info(" Stage 2 abgeschlossen – Fenster schließt nach Bestätigung.")
    pause(" ENTER zum Beenden (Stage 1 läuft danach automatisch weiter)...")

    if STATE["errors"]:
        return 2
    if STATE["manual_actions"]:
        return 1
    # Auch reine Warnungen sichtbar machen: sie koennen ein fehlendes
    # Pflichtwerkzeug bedeuten (veraPDF & Co.), und Exit-Code 0 haette dem
    # Aufrufer - Aufgabenplanung oder Wrapper-Skript - einen sauberen Lauf
    # gemeldet.
    if STATE["warnings"]:
        return 1
    return 0

def main():
    setup_console()
    # parse_run_mode() vor setup_logging(): Stage 2 uebernimmt per --log das
    # Logfile von Stage 1, statt eine zweite Datei anzulegen.
    parse_run_mode()
    try:
        os.makedirs(TMP_DIR, exist_ok=True)
    except Exception as _e:
        logging.debug(f"main: Exception verworfen: {_e!r}")
    setup_logging()

    if RUN_MODE == "stage2":
        return main_stage2()
    return main_stage1()

if __name__ == "__main__":
    exit_code = 0
    try:
        exit_code = main() or 0
    except KeyboardInterrupt:
        print("\nAbgebrochen durch Benutzer.")
        exit_code = 3
    except Exception as e:
        print(f"\n❌ Unerwarteter Fehler: {e}")
        try:
            logging.exception("Unerwarteter Fehler im Hauptablauf")
        except Exception:
            pass
        exit_code = 3
    finally:
        if RUN_MODE == "stage1":
            pause("\nDrücke ENTER zum Beenden...")
    sys.exit(exit_code)
