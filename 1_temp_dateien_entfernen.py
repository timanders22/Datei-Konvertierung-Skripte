# ================================================================================
# TEMP-FILE CLEANER
# ================================================================================
# Löscht temporäre und unnötige Dateien in einem Verzeichnis und allen
# Unterverzeichnissen. Unterstützt Netzlaufwerke (UNC), Long-Paths und
# optionale Entfernung leerer Ordner.
#
# Kompatibel mit Microsoft Office 2019 (DE/EN) und Office 2024 – die relevanten
# Sperr- und Recovery-Patterns (~$, ~wr, .asd, .wbk, .xlk, .~tmp, .~rf) sind
# sprach- und versionsunabhängig.
#
# NUTZUNG:
#   python 1_temp_dateien_entfernen.py
#
# HINWEIS FÜR PyInstaller-Build (Umwandlung in .exe):
#   python -m PyInstaller --onefile --console --icon="python_icon.ico" ^
#       --hidden-import winreg ^
#       1_temp_dateien_entfernen.py
#
# Stand: 11.06.2026
# ================================================================================

import os
import sys
import shutil
import stat
import time
import logging
import threading
import fnmatch
from datetime import datetime
from typing import List, Set, Tuple, Dict, Generator, Optional, Any
from concurrent.futures import ThreadPoolExecutor

IS_WINDOWS = os.name == "nt"

# ================================================================================
# Konfiguration
# ================================================================================

LOG_FILE_PREFIX   = "1_temp_dateien_entfernen"
MAX_WORKERS       = 8
SUBMIT_BUFFER     = MAX_WORKERS * 50
PROGRESS_INTERVAL = 0.5
# Retry-Parameter fuer Datei-Loeschungen: bewusst aggressiv-kurz gehalten,
# weil unter MAX_WORKERS Threads ein langer Backoff den ganzen Pool ausbremsen
# kann. Temp-Dateien sind entweder direkt loeschbar oder dauerhaft gelockt
# (z.B. von einem laufenden Prozess - dann hilft auch 7s Warten nichts).
# Worst-case-Wartezeit pro gelockter Datei: 0.1 + 0.2 = 0.3s (vorher: 7.5s).
RETRY_BASE_DELAY  = 0.1
MAX_DELETE_RETRIES = 3

# Mindestalter in Stunden, das eine Datei haben muss, bevor sie geloescht wird.
# Schuetzt gegen den gefaehrlichsten Fall auf einem gemeinsam genutzten Share:
# ~$-Sperrdateien GERADE GEOEFFNETER Office-Dokumente, laufende Speichervorgaenge
# (Word legt beim Speichern ~WRDxxxx.tmp an und benennt sie danach um) und
# unfertige Downloads (.part/.crdownload). Ohne diesen Filter raeumt das Skript
# mitten in aktive Schreibvorgaenge hinein.
# 0 = Filter deaktiviert (nicht empfohlen bei Mehrbenutzer-Shares).
MIN_AGE_HOURS     = 24

# Verzeichnis-Presets fuer die Startauswahl. Hier die im eigenen Netz
# gebraeuchlichen Laufwerke/Shares eintragen. Format: (Pfad, Hinweistext)
DIRECTORY_PRESETS: Tuple[Tuple[str, str], ...] = (
    ("Q:\\",                 ""),
    ("R:\\",                 ""),
    ("G:\\Meine Ablage",     ""),
    ("G:\\Geteilte Ablagen", ""),
    ("\\\\server\\dfs",      "Platzhalter – eigenen Dateiserver eintragen"),
)

# Pfade, die niemals geloescht werden (z. B. das eigene Logfile). Wird in
# main() befuellt, sobald der Log-Pfad feststeht.
_PROTECTED_PATHS: Set[str] = set()

# Modulweiter Lock zur Synchronisation aller Konsolen-Ausgaben zwischen
# Logger und ProgressReporter. Beide schreiben sonst unkoordiniert ans
# selbe Terminal (Logger ueber sys.stderr, ProgressReporter ueber
# sys.stdout) und koennen sich gegenseitig zerreissen. Wird sowohl im
# Custom-Loghandler als auch in ProgressReporter._render verwendet.
CONSOLE_LOCK = threading.Lock()

# ================================================================================
# Kritische Dateien (Whitelist – NIEMALS löschen)
# ================================================================================

CRITICAL_FILES: Set[str] = {
    "ntuser.dat", "ntuser.dat.log", "ntuser.ini",
    "sam", "system", "software", "security",
    "boot.ini", "bootmgr", "ntldr",
    "pagefile.sys", "hiberfil.sys", "swapfile.sys",
    "autorun.inf",
}

# ================================================================================
# Regel-Tabellen (Präfix / Suffix / Exakt / Wildcard)
# ================================================================================

_PREFIX_RULES: Tuple[Tuple[str, str], ...] = (
    ("._",            "mac"),
    ("~$",            "windows"),
    ("~wr",           "windows"),
    (".~",            "windows"),
    (".nfs",          "linux"),
    (".fuse_hidden",  "linux"),
    (".trash-",       "linux"),
)
# Entfernt: ("tmp.", "other") und ("bridge cache", "app").
# Beides waren Praefix-Regeln und trafen damit jede Datei, die so BEGINNT –
# also auch "tmp.docx" oder "bridge cache Notizen.txt". Praezisere Fassungen
# stehen jetzt in _WILDCARD_RULES.

_SUFFIX_RULES_BASE: Tuple[Tuple[str, str], ...] = (
    (".tmp",          "windows"),
    (".temp",         "windows"),
    (".~tmp",         "windows"),
    (".~rf",          "windows"),
    (".dmp",          "windows"),
    (".chk",          "windows"),
    (".asd",          "windows"),
    (".wbk",          "windows"),
    (".xlk",          "windows"),
    (".syd",          "windows"),
    (".swp",          "linux"),
    (".swo",          "linux"),
    (".swn",          "linux"),
    (".dwl",          "app"),
    (".dwl2",         "app"),
    (".sv$",          "app"),
    (".prv",          "app"),
    (".psd.lock",     "app"),
    (".ai.lock",      "app"),
    (".idlk",         "app"),
    (".pyc",          "dev"),
    (".pyo",          "dev"),
    (".class",        "dev"),
    (".suo",          "dev"),
    (".user",         "dev"),
    (".ncb",          "dev"),
    (".sdf",          "dev"),
    (".opensdf",      "dev"),
    (".vspscc",       "dev"),
    (".vssscc",       "dev"),
    (".crdownload",   "other"),
    (".download",     "other"),
    (".tmpdownload",  "other"),
    (".part",         "other"),
    (".partial",      "other"),
    (".incomplete",   "other"),
    (".!ut",          "other"),
    (".!qb",          "other"),
)

_OPTIONAL_BAK_LOG_SUFFIXES: Tuple[Tuple[str, str], ...] = (
    (".bak", "windows"),
    (".log", "other"),
)

_EXACT_RULES: Dict[str, str] = {
    "thumbs.db":               "windows",
    "thumbs.db:encryptable":   "windows",
    "ehthumbs.db":             "windows",
    "ehthumbs_vista.db":       "windows",
    "msdownld.tmp":            "windows",
    "mscreate.dir":            "windows",
    "folder.htt":              "windows",
    ".ds_store":               "mac",
    ".localized":              "mac",
    "icon\r":                  "mac",
    ".apdisk":                 "mac",
    ".directory":              "linux",
    ".xdg-volume-info":        "linux",
}

# desktop.ini ist KEINE Temp-Datei: darin stehen Ordnersymbole, lokalisierte
# Ordnernamen und Ordnervorlagen. Auf Shares und in Profilordnern ist der
# Verlust nicht wiederherstellbar – deshalb nur auf ausdruecklichen Wunsch.
_OPTIONAL_DESKTOP_INI: Dict[str, str] = {
    "desktop.ini": "windows",
}

_WILDCARD_RULES: Tuple[Tuple[str, str], ...] = (
    # tmp.<mind. 6 Hex-Zeichen> – der Zufallsname echter Temp-Dateien.
    # Trifft NICHT "tmp.docx", "tmp.dbf" o. ae.
    ("tmp.[0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f]*", "other"),
    # Adobe-Bridge-Caches heissen konkret "bridge cache.bc" / ".bct"
    ("bridge cache*.bc",  "app"),
    ("bridge cache*.bct", "app"),
)

# ================================================================================
# Ausschluss-Verzeichnisse
# ================================================================================

MAC_JUNK_DIRS: Set[str] = {
    ".trashes", ".fseventsd", ".spotlight-v100", ".temporaryitems",
    ".documentrevisions-v100", ".mobilebackups", ".appledouble",
    "network trash folder", "temporary items",
}
_MAC_JUNK_DIRS_LOWER: Set[str] = {d.lower() for d in MAC_JUNK_DIRS}

EXCLUDE_DIRS: Set[str] = {
    "$RECYCLE.BIN", "System Volume Information",
    "Windows", "Program Files", "Program Files (x86)",
    "AppData",
    ".git", ".svn", ".vs", ".idea", ".vscode",
    "__pycache__", "node_modules", "vendor",
    ".cache", ".Trash", "lost+found",
}
_EXCLUDE_DIRS_LOWER: Set[str] = {d.lower() for d in EXCLUDE_DIRS}

# ================================================================================
# Konsolen-Encoding und Log-Pfad
# ================================================================================

def _configure_console_encoding() -> None:
    if not IS_WINDOWS:
        return
    for stream_name in ("stdout", "stderr", "stdin"):
        stream = getattr(sys, stream_name, None)
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass


def _get_application_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _get_log_path() -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_name  = f"{LOG_FILE_PREFIX}_{timestamp}.log"

    primary = os.path.join(_get_application_dir(), log_name)
    try:
        with open(primary, "a", encoding="utf-8-sig"):
            pass
        return primary
    except OSError:
        pass

    # Fallback-Kandidaten der Reihe nach auf Schreibbarkeit testen. Vorher
    # wurde der erste Kandidat ungeprueft zurueckgegeben – war auch der nicht
    # beschreibbar, starb das Skript beim Anlegen des FileHandlers.
    import tempfile
    fallback_root = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    for candidate_dir in (os.path.join(fallback_root, "TempCleaner"),
                          os.path.expanduser("~"),
                          tempfile.gettempdir()):
        try:
            os.makedirs(candidate_dir, exist_ok=True)
            candidate = os.path.join(candidate_dir, log_name)
            with open(candidate, "a", encoding="utf-8-sig"):
                pass
            return candidate
        except OSError:
            continue
    return os.path.join(tempfile.gettempdir(), log_name)

# ================================================================================
# Logging
# ================================================================================

def _setup_logging(log_path: str) -> None:
    file_handler = logging.FileHandler(log_path, mode="w", encoding="utf-8-sig")
    file_handler.setLevel(logging.INFO)

    # Custom-Handler, der die laufende ProgressReporter-Zeile sauber leert,
    # bevor eine Log-Meldung in die Konsole geschrieben wird. Schuetzt
    # GEGEN ZWEI Szenarien:
    #  - Logger feuert mitten in eine Progress-Zeile: clear leert sie zuerst
    #  - Progress feuert mitten in einen Log-Output: CONSOLE_LOCK serialisiert
    # Das CONSOLE_LOCK muss MODULWEIT geteilt werden, damit Logger und
    # ProgressReporter sich gegenseitig blockieren - ein lokaler Lock im
    # Reporter (frueher: self.print_lock) wuerde den Logger nicht aussperren.
    class _ProgressAwareStreamHandler(logging.StreamHandler):
        def emit(self, record):
            with CONSOLE_LOCK:
                try:
                    # Aktuelle Zeile mit ausreichend Spaces leeren, dann \r,
                    # damit der nachfolgende Log-Output am Zeilenanfang
                    # beginnt. 120 als grosszuegige Standardbreite; die
                    # exakte Konsolenbreite muss hier nicht ermittelt
                    # werden, weil das \r die Zeile in jedem Fall an den
                    # Anfang setzt.
                    self.stream.write("\r" + " " * 120 + "\r")
                    self.stream.flush()
                except Exception:
                    pass
                super().emit(record)

    console_handler = _ProgressAwareStreamHandler()
    console_handler.setLevel(logging.ERROR)
    console_handler.setFormatter(
        logging.Formatter("%(levelname)s - %(message)s")
    )

    # force=True: basicConfig ist sonst ein No-op, sobald irgendein Modul
    # bereits einen Root-Handler gesetzt hat – dann landet nichts im Logfile.
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[file_handler, console_handler],
        force=True,
    )

# ================================================================================
# Pfad-Hilfsfunktionen
# ================================================================================

def prepare_long_path(path: str) -> str:
    if not IS_WINDOWS:
        return os.path.abspath(path)
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


def safe_exists(path: str) -> bool:
    try:
        return os.path.exists(prepare_long_path(path))
    except Exception:
        return False


def safe_remove_with_retry(path: str, max_retries: int = MAX_DELETE_RETRIES) -> bool:
    target = prepare_long_path(path)
    delay  = RETRY_BASE_DELAY
    for attempt in range(max_retries):
        try:
            os.remove(target)
            return True
        except FileNotFoundError:
            return False
        except PermissionError:
            if attempt < max_retries - 1:
                try:
                    os.chmod(target, stat.S_IWRITE)
                except OSError:
                    pass
                time.sleep(delay)
                delay *= 2
        except OSError:
            if attempt < max_retries - 1:
                time.sleep(delay)
                delay *= 2
    return False


def get_file_size(path: str) -> int:
    try:
        return os.path.getsize(prepare_long_path(path))
    except Exception:
        return 0


def get_file_stat(path: str) -> Optional[os.stat_result]:
    """Ein einziger stat-Aufruf fuer Groesse UND Alter (spart einen Roundtrip
    pro Loeschkandidat, was auf Netzlaufwerken spuerbar ist)."""
    try:
        return os.stat(prepare_long_path(path))
    except Exception:
        return None


def is_old_enough(st: os.stat_result, min_age_seconds: float) -> bool:
    """True, wenn die Datei alt genug zum Loeschen ist.

    Bei Uhrzeit-Versatz zwischen Client und Fileserver kann mtime in der
    Zukunft liegen; das Ergebnis ist dann negativ und die Datei gilt als zu
    jung. Das ist die sichere Richtung.
    """
    if min_age_seconds <= 0:
        return True
    try:
        return (time.time() - st.st_mtime) >= min_age_seconds
    except Exception:
        return False


def register_protected_path(path: str) -> None:
    try:
        _PROTECTED_PATHS.add(os.path.normcase(os.path.abspath(path)))
    except Exception:
        pass


def is_protected_path(path: str) -> bool:
    """Schuetzt u. a. das eigene Logfile.

    Ohne diese Pruefung loescht sich das Skript bei aktivierter *.log-Option
    das gerade laufende Log selbst weg, sobald es im Zielverzeichnis liegt
    (Standardfall: .exe und Log liegen auf demselben Netzlaufwerk).
    """
    if not _PROTECTED_PATHS:
        return False
    try:
        return os.path.normcase(os.path.abspath(_strip_long_prefix(path))) in _PROTECTED_PATHS
    except Exception:
        return False


def _strip_long_prefix(path: str) -> str:
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\"):
        return path[4:]
    return path


def is_excluded_dir(dir_name: str) -> bool:
    dl = dir_name.lower()
    return dl in _EXCLUDE_DIRS_LOWER or dl in _MAC_JUNK_DIRS_LOWER


def is_critical_file(filename: str) -> bool:
    return filename.lower() in CRITICAL_FILES

# ================================================================================
# User Shell Folders (Registry-Lookup mit Fallback)
# ================================================================================

def _get_user_shell_folder(value_name: str, fallback_subdir: str) -> str:
    if IS_WINDOWS:
        try:
            import winreg
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                raw_value, _ = winreg.QueryValueEx(key, value_name)
                expanded = winreg.ExpandEnvironmentStrings(raw_value)
                if expanded and os.path.isdir(expanded):
                    return os.path.normpath(expanded)
        except (OSError, ImportError, FileNotFoundError):
            pass
    return os.path.join(os.path.expanduser("~"), fallback_subdir)


def get_desktop_path() -> str:
    return _get_user_shell_folder("Desktop", "Desktop")


def get_downloads_path() -> str:
    return _get_user_shell_folder(
        "{374DE290-123F-4565-9164-39C4925E467B}", "Downloads"
    )

# ================================================================================
# Interaktive Eingaben
# ================================================================================

def _is_valid_dir(path: str) -> bool:
    # try/except gegen OSError/ValueError aus prepare_long_path bzw.
    # os.path.isdir bei ungueltigen Zeichen (?, <, >, |, \x00) oder
    # kaputten Long-Path-Praefixen. Ohne diesen Schutz crasht das
    # Skript bei Tippfehlern in der Pfad-Eingabe komplett ab.
    try:
        return os.path.isdir(prepare_long_path(path))
    except (OSError, ValueError):
        return False


def ask_directory() -> str:
    desktop_path   = get_desktop_path()
    downloads_path = get_downloads_path()

    # (Anzeigetext, Pfad, Hinweis)
    entries: List[Tuple[str, str, str]] = [
        (path, path, hint) for path, hint in DIRECTORY_PRESETS
    ]
    entries.append((f"Desktop:    {desktop_path}",   desktop_path,   ""))
    entries.append((f"Downloads:  {downloads_path}", downloads_path, ""))
    custom = len(entries) + 1

    print("\nZielverzeichnis auswählen:")
    for i, (label, _p, hint) in enumerate(entries, start=1):
        print(f"  [{i}] {label}" + (f"   ({hint})" if hint else ""))
    print(f"  [{custom}] Eigenen Pfad eingeben")
    print()

    while True:
        try:
            choice = input(f"Auswahl [1-{custom}]: ").strip()
        except EOFError:
            print("\nKeine Eingabe möglich – Abbruch.")
            raise KeyboardInterrupt

        if choice.isdigit() and 1 <= int(choice) <= len(entries):
            path = entries[int(choice) - 1][1]
            # Presets sofort pruefen: sonst faellt ein nicht verbundenes
            # Laufwerk erst nach allen weiteren Rueckfragen und der
            # 'JA'-Bestaetigung auf.
            if _is_valid_dir(path):
                return path
            print(f"  ❌ Nicht erreichbar: '{path}'")
            print("     (Laufwerk verbunden? Netzwerkpfad erreichbar?)")
            continue

        if choice == str(custom):
            try:
                raw = input("Pfad eingeben: ")
            except EOFError:
                print("\nKeine Eingabe möglich – Abbruch.")
                raise KeyboardInterrupt
            path = sanitize_path(raw)
            if _is_valid_dir(path):
                return path
            print(f"  ❌ Verzeichnis nicht gefunden: '{path}'")
            print("     (Hinweis: Anführungszeichen werden automatisch entfernt)")
        else:
            print(f"  Bitte 1-{custom} eingeben.")


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

# ================================================================================
# Regel-Matching
# ================================================================================

def should_delete_file(
    filename: str,
    exact_rules: Dict[str, str],
    suffix_rules: Tuple[Tuple[str, str], ...],
    wildcard_rules: Tuple[Tuple[str, str], ...],
) -> Tuple[bool, str, str]:
    if not filename:
        return False, "Leer", ""

    if is_critical_file(filename):
        return False, "KRITISCH", ""

    file_lower = filename.lower()

    for prefix, cat in _PREFIX_RULES:
        if file_lower.startswith(prefix):
            return True, f"Präfix: {prefix}", cat

    for suffix, cat in suffix_rules:
        if file_lower.endswith(suffix):
            return True, f"Suffix: {suffix}", cat

    if file_lower in exact_rules:
        return True, f"Exakt: {file_lower}", exact_rules[file_lower]

    for pattern, cat in wildcard_rules:
        if fnmatch.fnmatch(file_lower, pattern):
            return True, f"Pattern: {pattern}", cat

    return False, "Behalten", ""

# ================================================================================
# rmtree-Error-Handling (Python <3.12 vs >=3.12)
# ================================================================================

def _rmtree_handle(func, path: str, exc: BaseException) -> None:
    if isinstance(exc, PermissionError):
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
            return
        except Exception:
            pass
    logging.error(f"rmtree-Fehler (nicht behebbar): {path} – {exc}")


def _rmtree_onerror(func, path: str, excinfo) -> None:
    _rmtree_handle(func, path, excinfo[1])


def _rmtree_onexc(func, path: str, exc: BaseException) -> None:
    _rmtree_handle(func, path, exc)


_RMTREE_KWARGS: Dict[str, Any] = (
    {"onexc": _rmtree_onexc}
    if sys.version_info >= (3, 12)
    else {"onerror": _rmtree_onerror}
)

# ================================================================================
# Kombinierter Walk: Mac-Junk-Dirs räumen + Dateien yielden
# ================================================================================

def walk_and_clean_junk(
    base_dir: str,
    stats: "CleanupStats",
    dry_run: bool = False,
) -> Generator[str, None, None]:
    def _walk_error(err: OSError) -> None:
        logging.warning(f"Verzeichnis nicht lesbar (fehlende Rechte?) – übersprungen: {err.filename}")

    for root, dirs, files in os.walk(base_dir, topdown=True, onerror=_walk_error):
        junk = [d for d in dirs if d.lower() in _MAC_JUNK_DIRS_LOWER]
        dirs[:] = [d for d in dirs if d not in junk and not is_excluded_dir(d)]

        for d in junk:
            full = os.path.join(root, d)
            if dry_run:
                logging.info(f"[SIMULATION] Würde Mac-Junk entfernen: {full}")
                stats.add_mac_junk()
                continue
            try:
                shutil.rmtree(full, **_RMTREE_KWARGS)
                logging.info(f"Mac-Junk entfernt: {full}")
                stats.add_mac_junk()
            except Exception as e:
                logging.error(f"Mac-Junk-Fehler {full}: {e}")
                stats.add_mac_junk_error()

        for f in files:
            yield os.path.join(root, f)

# ================================================================================
# Leere Ordner entfernen
# ================================================================================

def remove_empty_dirs(base_dir: str, dry_run: bool = False) -> Tuple[int, int]:
    removed = 0
    errors  = 0
    candidates: List[str] = []
    norm_base = os.path.normpath(base_dir)

    def _walk_error(err: OSError) -> None:
        logging.warning(f"Verzeichnis nicht lesbar – übersprungen: {err.filename}")

    for root, dirs, files in os.walk(base_dir, topdown=True, onerror=_walk_error):
        dirs[:] = [d for d in dirs if not is_excluded_dir(d)]
        if os.path.normpath(root) == norm_base:
            continue
        candidates.append(root)

    for root in reversed(candidates):
        try:
            if not os.listdir(root):
                if dry_run:
                    # In der Simulation wird nichts entfernt, deshalb werden
                    # auch keine Ordner "nachtraeglich leer". Die reale Zahl
                    # liegt also tendenziell hoeher als hier angezeigt.
                    logging.info(f"[SIMULATION] Würde leeren Ordner entfernen: {root}")
                    removed += 1
                    continue
                try:
                    os.rmdir(root)
                    logging.info(f"Leerer Ordner entfernt: {root}")
                    removed += 1
                except OSError as e:
                    logging.error(f"Ordner nicht löschbar: {root} – {e}")
                    errors += 1
        except Exception as e:
            logging.error(f"Fehler beim Ordner-Check {root}: {e}")
            errors += 1

    return removed, errors

# ================================================================================
# Statistik
# ================================================================================

class CleanupStats:
    def __init__(self) -> None:
        self.lock             = threading.Lock()
        self.deleted          = 0
        self.skipped          = 0
        self.errors           = 0
        self.too_young        = 0
        self.total_size       = 0
        self.processed        = 0
        self.dirs_removed     = 0
        self.dirs_errors      = 0
        self.mac_junk_removed = 0
        self.mac_junk_errors  = 0
        self.start_time       = time.monotonic()
        self.categories: Dict[str, int] = {
            "windows": 0, "mac": 0, "linux": 0,
            "app": 0, "dev": 0, "other": 0,
        }

    def add_deleted(self, size: int, category: str) -> None:
        with self.lock:
            self.deleted    += 1
            self.total_size += size
            self.processed  += 1
            if category in self.categories:
                self.categories[category] += 1
            else:
                self.categories["other"]  += 1

    def add_skipped(self) -> None:
        with self.lock:
            self.skipped   += 1
            self.processed += 1

    def add_too_young(self) -> None:
        with self.lock:
            self.too_young += 1
            self.skipped   += 1
            self.processed += 1

    def add_error(self) -> None:
        with self.lock:
            self.errors    += 1
            self.processed += 1

    def add_mac_junk(self) -> None:
        with self.lock:
            self.mac_junk_removed += 1

    def add_mac_junk_error(self) -> None:
        with self.lock:
            self.mac_junk_errors += 1

    def print_summary(self, dry_run: bool = False) -> None:
        elapsed = time.monotonic() - self.start_time
        size_mb = self.total_size / (1024 * 1024)
        verb    = "Würde entfernen" if dry_run else "Entfernt"

        print("\n" + "=" * 70)
        print("  CLEANUP ZUSAMMENFASSUNG" + ("  (SIMULATION)" if dry_run else ""))
        print("=" * 70)
        print(f"  Dauer:              {elapsed:.1f} s")
        print()
        print(f"  Dateien geprüft:    {self.processed:,}")
        print(f"  {verb + ':':<19} {self.deleted:,}")
        print(f"  Übersprungen:       {self.skipped:,}")
        if self.too_young:
            print(f"    davon zu jung:    {self.too_young:,}  (< {MIN_AGE_HOURS} h – bewusst geschont)")
        print(f"  Fehler (Dateien):   {self.errors:,}")
        print(f"  {'Speicher (theor.):' if dry_run else 'Speicher freig.:':<19} {size_mb:.2f} MB")

        if self.mac_junk_removed or self.mac_junk_errors:
            print()
            print(f"  Mac-Junk {'(würde)' if dry_run else 'entfernt'}:  {self.mac_junk_removed:,}")
            print(f"  Mac-Junk-Fehler:    {self.mac_junk_errors:,}")

        if self.dirs_removed or self.dirs_errors:
            print()
            print(f"  Leere Ordner {'(würde)' if dry_run else 'entf.'}: {self.dirs_removed:,}")
            print(f"  Ordner-Fehler:      {self.dirs_errors:,}")

        if self.deleted > 0:
            print()
            print("  Kategorien:")
            for cat, count in self.categories.items():
                if count > 0:
                    print(f"    {cat.upper():10} {count:6,}")

        print("=" * 70)

# ================================================================================
# Fortschritts-Anzeige (eigener Daemon-Thread)
# ================================================================================

class ProgressReporter:
    def __init__(self, stats: CleanupStats, base_dir: str,
                 base_dir_display: str,
                 total_files: Optional[int] = None) -> None:
        self.stats            = stats
        self.base_dir         = base_dir
        self.base_dir_display = base_dir_display
        self.total_files      = total_files
        self.current_path     = ""
        self.path_lock        = threading.Lock()
        # Bewusst KEIN eigener print_lock mehr - alle Konsolen-Ausgaben
        # synchronisieren ueber den modulweiten CONSOLE_LOCK, sodass auch
        # Logger und ProgressReporter sich gegenseitig nicht zerreissen.
        self.stop_event       = threading.Event()
        self.thread: Optional[threading.Thread] = None
        # Terminal-Breite cachen: shutil.get_terminal_size oeffnet bei jedem
        # Aufruf einen ioctl/Win32-Call. Bei ~2 Renders pro Sekunde ueber
        # einen langen Lauf summiert sich das. Refresh alle 5 Sekunden.
        self._cols_cache      = 120
        self._cols_refresh_at = 0.0

    def update_path(self, path: str) -> None:
        with self.path_lock:
            self.current_path = path

    def start(self) -> None:
        self.thread = threading.Thread(target=self._loop, daemon=True)
        self.thread.start()

    def stop(self) -> None:
        self.stop_event.set()
        if self.thread:
            self.thread.join(timeout=1.5)
        self._render()
        print()

    def _loop(self) -> None:
        while not self.stop_event.wait(PROGRESS_INTERVAL):
            self._render()

    def _render(self) -> None:
        with self.path_lock:
            path = self.current_path

        path_info = ""
        if path:
            try:
                base     = self.base_dir
                base_len = len(base) if base.endswith(os.sep) else len(base) + 1
                remainder = path[base_len:]

                first_slash = remainder.find(os.sep)
                if first_slash != -1:
                    subfolder  = remainder[:first_slash]
                    short_path = os.path.join(self.base_dir_display, subfolder) + os.sep + "[...]"
                else:
                    short_path = self.base_dir_display + os.sep + "[...]"

                if len(short_path) > 45:
                    short_path = "..." + short_path[-42:]
                path_info = f" | 📂 {short_path}"
            except Exception:
                pass

        proc    = self.stats.processed
        deleted = self.stats.deleted

        # Terminal-Breite cachen (alle 5s refresh). Bewusst AUSSERHALB des
        # CONSOLE_LOCK, weil get_terminal_size potenziell einen ioctl-Call
        # macht - im Lock waere das unnoetig blockierend.
        now = time.monotonic()
        if now >= self._cols_refresh_at:
            try:
                self._cols_cache = shutil.get_terminal_size((120, 24)).columns
            except Exception:
                self._cols_cache = 120
            self._cols_refresh_at = now + 5.0
        cols = self._cols_cache

        # Statuszeile als EIN print()-Aufruf vorbereiten und unter CONSOLE_LOCK
        # ausgeben. Frueher: zwei separate prints (clear + content) - auch mit
        # Lock fuer den Logger eine Race-Quelle, weil zwischen den beiden
        # prints ein anderer Lock-Inhaber (z.B. ein parallel feuerndes print
        # aus Print-Hauptthread) die Konsole anfassen koennte. Mit einem
        # einzelnen print + flush ist das ausgeschlossen.
        clear   = "\r" + " " * max(cols - 1, 0) + "\r"
        if self.total_files is not None and self.total_files > 0:
            perc = min(proc / self.total_files * 100, 100.0)
            content = (f"  ⏳ {proc:,}/{self.total_files:,} "
                       f"({perc:.0f}%)  –  "
                       f"{deleted:,} entfernt{path_info}")
        else:
            content = (f"  ⏳ {proc:,} Dateien geprüft  –  "
                       f"{deleted:,} entfernt{path_info}")

        with CONSOLE_LOCK:
            print(clear + content, end="", flush=True)

# ================================================================================
# Verarbeitung einer einzelnen Datei
# ================================================================================

def process_file(
    filepath: str,
    stats: CleanupStats,
    exact_rules: Dict[str, str],
    suffix_rules: Tuple[Tuple[str, str], ...],
    wildcard_rules: Tuple[Tuple[str, str], ...],
    abort_flag: threading.Event,
    progress: ProgressReporter,
    dry_run: bool = False,
    min_age_seconds: float = 0.0,
) -> None:
    if abort_flag.is_set():
        return

    progress.update_path(filepath)

    try:
        filename = os.path.basename(filepath)
        should, reason, category = should_delete_file(
            filename, exact_rules, suffix_rules, wildcard_rules
        )

        if not should:
            stats.add_skipped()
            return

        if is_protected_path(filepath):
            logging.info(f"Geschützt (nicht gelöscht): {filepath}")
            stats.add_skipped()
            return

        st = get_file_stat(filepath)
        if st is None:
            # Datei zwischenzeitlich weg oder nicht lesbar – nicht als Fehler
            # werten, das passiert bei Temp-Dateien im laufenden Betrieb staendig.
            stats.add_skipped()
            return

        if not is_old_enough(st, min_age_seconds):
            logging.info(f"Zu jung, geschont: {filepath}  ({reason})")
            stats.add_too_young()
            return

        file_size = st.st_size

        if dry_run:
            logging.info(f"[SIMULATION] Würde löschen: {filepath}  ({reason}, {file_size/1024:.1f} KB)")
            stats.add_deleted(file_size, category)
            return

        if safe_remove_with_retry(filepath):
            logging.info(f"Gelöscht: {filepath}  ({reason}, {file_size/1024:.1f} KB)")
            stats.add_deleted(file_size, category)
        else:
            logging.error(f"Fehler/Lock: {filepath}")
            stats.add_error()

    except Exception as e:
        logging.error(f"Fehler bei {filepath}: {e}")
        stats.add_error()

# ================================================================================
# Vorab-Zählung mit Live-Anzeige
# ================================================================================

def _count_files_with_progress(base_dir: str) -> int:
    def _walk_error(err: OSError) -> None:
        logging.warning(f"Verzeichnis nicht lesbar: {err.filename}")

    count      = 0
    last_print = time.monotonic()
    try:
        for root, dirs, files in os.walk(base_dir, topdown=True, onerror=_walk_error):
            dirs[:] = [d for d in dirs
                       if not is_excluded_dir(d)
                       and d.lower() not in _MAC_JUNK_DIRS_LOWER]
            count += len(files)
            now = time.monotonic()
            if now - last_print >= PROGRESS_INTERVAL:
                print(f"\r   Indiziere ... {count:,} Dateien", end="", flush=True)
                last_print = now
    except KeyboardInterrupt:
        raise
    except Exception as e:
        logging.error(f"Fehler beim Zählen: {e}")
    return count

# ================================================================================
# Hauptablauf
# ================================================================================

def run_cleanup(
    base_dir: str,
    base_dir_display: str,
    remove_empty: bool,
    count_first: bool,
    exact_rules: Dict[str, str],
    suffix_rules: Tuple[Tuple[str, str], ...],
    wildcard_rules: Tuple[Tuple[str, str], ...],
    dry_run: bool = False,
    min_age_seconds: float = 0.0,
) -> CleanupStats:
    stats = CleanupStats()

    print(f"\n{'=' * 70}")
    print("  TEMP-FILE CLEANER  –  " + ("SIMULATION (es wird nichts gelöscht)" if dry_run else "LÖSCHEN"))
    print(f"{'=' * 70}")
    print(f"  Verzeichnis: {base_dir_display}")
    print(f"  Start:       {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    if not safe_exists(base_dir):
        print(f"\n❌ Verzeichnis nicht gefunden: {base_dir_display}")
        return stats

    total_files: Optional[int] = None
    if count_first:
        print("\n🔍 Indiziere Dateien (kann dauern) ...")
        try:
            total_files = _count_files_with_progress(base_dir)
        except KeyboardInterrupt:
            print("\n\n⚠️  Indizierung abgebrochen!")
            return stats
        cols = shutil.get_terminal_size((120, 24)).columns
        print("\r" + " " * max(cols - 1, 0), end="")
        print(f"\r   Gesamt: {total_files:,} Dateien")
    else:
        print("  (Generator-Modus – Start ohne Vorab-Zählung)")

    print("\n🚀 Starte Verarbeitung ...\n")
    print("-" * 70)

    abort_flag       = threading.Event()
    submit_semaphore = threading.BoundedSemaphore(SUBMIT_BUFFER)
    progress         = ProgressReporter(stats, base_dir, base_dir_display, total_files)
    executor         = ThreadPoolExecutor(max_workers=MAX_WORKERS)

    def _worker_wrapper(fp: str) -> None:
        try:
            process_file(
                fp, stats,
                exact_rules, suffix_rules, wildcard_rules,
                abort_flag, progress,
                dry_run, min_age_seconds,
            )
        except Exception as exc:
            logging.error(f"Thread-Ausführung fehlgeschlagen: {exc}")
            stats.add_error()
        finally:
            submit_semaphore.release()

    progress.start()

    try:
        try:
            for filepath in walk_and_clean_junk(base_dir, stats, dry_run):
                if abort_flag.is_set():
                    break

                submit_semaphore.acquire()
                if abort_flag.is_set():
                    submit_semaphore.release()
                    break

                try:
                    executor.submit(_worker_wrapper, filepath)
                except Exception:
                    submit_semaphore.release()
                    raise

            executor.shutdown(wait=True)

        except KeyboardInterrupt:
            abort_flag.set()
            if sys.version_info >= (3, 9):
                executor.shutdown(wait=False, cancel_futures=True)
            else:
                executor.shutdown(wait=False)
            raise
        except Exception:
            abort_flag.set()
            if sys.version_info >= (3, 9):
                executor.shutdown(wait=False, cancel_futures=True)
            else:
                executor.shutdown(wait=False)
            raise
        finally:
            progress.stop()

        if stats.mac_junk_removed or stats.mac_junk_errors:
            wort = "Würde entfernen" if dry_run else "Entfernt"
            print(f"  🍎 {wort}: {stats.mac_junk_removed:,} Mac-Junk-Verzeichnis(se), "
                  f"{stats.mac_junk_errors} Fehler")

        if remove_empty:
            print()
            print("-" * 70)
            print("  🗂  " + ("Suche leere Ordner (Simulation) ..." if dry_run
                              else "Entferne leere Ordner ..."))
            removed, errors = remove_empty_dirs(base_dir, dry_run)
            stats.dirs_removed = removed
            stats.dirs_errors  = errors
            wort = "leere Ordner gefunden" if dry_run else "leere Ordner entfernt"
            print(f"  → {removed} {wort}, {errors} Fehler")

    except KeyboardInterrupt:
        print("\n\n⚠️  Abbruch durch Benutzer!")
        logging.warning("Abgebrochen durch Benutzer")
        stats.print_summary(dry_run)
        return stats
    except Exception as e:
        print(f"\n\n❌ Kritischer Fehler: {e}")
        logging.error(f"Kritischer Fehler: {e}", exc_info=True)
        stats.print_summary(dry_run)
        return stats

    stats.print_summary(dry_run)
    return stats

# ================================================================================
# Einstiegspunkt
# ================================================================================

def _build_suffix_rules(include_bak_log: bool) -> Tuple[Tuple[str, str], ...]:
    if include_bak_log:
        return _SUFFIX_RULES_BASE + _OPTIONAL_BAK_LOG_SUFFIXES
    return _SUFFIX_RULES_BASE


def _build_exact_rules(include_desktop_ini: bool) -> Dict[str, str]:
    if include_desktop_ini:
        rules = dict(_EXACT_RULES)
        rules.update(_OPTIONAL_DESKTOP_INI)
        return rules
    return _EXACT_RULES


def main() -> None:
    _configure_console_encoding()

    log_path = _get_log_path()
    _setup_logging(log_path)
    # Eigenes Logfile vor dem Selbstlöschen schützen (relevant, sobald die
    # *.log-Option aktiv ist und Skript/Log im Zielverzeichnis liegen).
    register_protected_path(log_path)

    if not IS_WINDOWS:
        print("\n⚠️  Dieses Skript ist für Windows/Netzlaufwerke konzipiert.")
        print("   Auf diesem System (os.name=%r) funktionieren Long-Path-Präfixe," % os.name)
        print("   UNC-Pfade und einige Löschmuster möglicherweise nicht korrekt.\n")

    print()
    print("╔══════════════════════════════════════════════════════════════╗")
    print("║                       TEMP-FILE CLEANER                      ║")
    print("╚══════════════════════════════════════════════════════════════╝")
    print(f"  Start: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    target_dir_display = ask_directory()
    target_dir         = prepare_long_path(target_dir_display)

    print()
    remove_empty = ask_yes_no(
        "Leere Ordner nach dem Löschen ebenfalls entfernen?",
        default_yes=False,
    )

    count_first = ask_yes_no(
        "\nDateien vorab zählen? (Ja = ETA-Anzeige, kostet Zeit beim Start)\n"
        "  (Hinweis: Bei Netzlaufwerken (UNC) und sehr vielen Dateien\n"
        "   (>100.000) kann die Zählung 10–30 Minuten dauern)",
        default_yes=False,
    )

    delete_bak_log = ask_yes_no(
        "\n*.bak und *.log Dateien ebenfalls löschen?\n"
        "  (⚠️  Riskant auf Netzlaufwerken – können Backups und Anwendungslogs sein)",
        default_yes=False,
    )

    delete_desktop_ini = ask_yes_no(
        "\ndesktop.ini ebenfalls löschen?\n"
        "  (⚠️  Enthält Ordnersymbole und lokalisierte Ordnernamen –\n"
        "   der Verlust lässt sich nicht rückgängig machen)",
        default_yes=False,
    )

    dry_run = not ask_yes_no(
        "\nECHT-Modus starten? (Nein = Simulation, es wird nichts gelöscht)",
        default_yes=False,
    )

    suffix_rules    = _build_suffix_rules(delete_bak_log)
    exact_rules     = _build_exact_rules(delete_desktop_ini)
    wildcard_rules  = _WILDCARD_RULES
    min_age_seconds = max(MIN_AGE_HOURS, 0) * 3600.0

    print()
    print("=" * 66)
    print("  KONFIGURATION")
    print("=" * 66)
    print(f"  Verzeichnis:         {target_dir_display}")
    if dry_run:
        print("  Modus:               🔍 SIMULATION (es wird nichts gelöscht)")
    else:
        print("  Modus:               ⚠️  ECHT-MODUS (löscht Dateien!)")
    if min_age_seconds > 0:
        print(f"  Mindestalter:        {MIN_AGE_HOURS} h (jüngere Dateien bleiben unangetastet)")
    else:
        print("  Mindestalter:        ⚠️  kein Filter – auch aktive Sperrdateien werden gelöscht")
    print(f"  Leere Ordner:        {'entfernen' if remove_empty else 'behalten'}")
    print(f"  Vorab zählen:        {'Ja (ETA)' if count_first else 'Nein (Generator)'}")
    print(f"  Threads:             {MAX_WORKERS}")
    print(f"  Lösch-Retries:       {MAX_DELETE_RETRIES} (Basis-Delay {RETRY_BASE_DELAY}s, exp. Backoff)")
    print()
    print("  Was wird gelöscht:")
    print("    • ~$-Sperrdateien (Word, Excel, PowerPoint – alle Office-Versionen)")
    print("    • ~wr*-Workfiles (Word)")
    print("    • ._-Artefakte (macOS auf Windows-Shares)")
    print("    • .ds_store, thumbs.db")
    print("    • *.tmp / *.temp / *.asd / *.wbk / *.xlk")
    print("    • *.crdownload, *.part (unvollst. Downloads)")
    print("    • *.swp / *.swo (Editor-Sperrdateien)")
    if delete_bak_log:
        print("    • *.bak / *.log  (auf Nutzerwunsch aktiviert)")
    else:
        print("    ○ *.bak / *.log  (nicht aktiv – zu riskant auf Netzlaufwerken)")
    if delete_desktop_ini:
        print("    • desktop.ini    (auf Nutzerwunsch aktiviert)")
    else:
        print("    ○ desktop.ini    (nicht aktiv – Ordnersymbole/-namen)")
    print("=" * 66)

    if dry_run:
        print()
        print("🔍 SIMULATION: Es wird nur protokolliert, nichts gelöscht.")
    else:
        print()
        print("⚠️  ECHT-MODUS: Dateien werden UNWIDERRUFLICH gelöscht!")
        try:
            confirm = input("   Tippe 'JA' (Großbuchstaben) zum Fortfahren: ").strip()
        except EOFError:
            confirm = ""
        if confirm != "JA":
            print("Abgebrochen.")
            return

    run_cleanup(
        target_dir, target_dir_display, remove_empty, count_first,
        exact_rules, suffix_rules, wildcard_rules,
        dry_run, min_age_seconds,
    )

    print()
    print(f"  Log gespeichert: {log_path}")
    try:
        input("\nEnter zum Beenden ...")
    except Exception:
        pass


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n❌ Abgebrochen (Ctrl+C)")
    except Exception as e:
        print(f"\n❌ UNBEKANNTER FEHLER: {e}")
        import traceback
        traceback.print_exc()
        try:
            input("\nEnter zum Beenden ...")
        except Exception:
            pass
