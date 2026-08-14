#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ==================================================================
# DATEINAMEN-BEREINIGUNG FÜR GOOGLE DRIVE FOR DESKTOP
# ==================================================================
# Stand: 11.06.2026
# ANLEITUNG:
# 1. Python installieren: https://www.python.org/
# 2. In der Konsole (CMD): pip install tqdm
# 3. In der Konsole (CMD): python 10_dateinamen_bereinigen.py
#
# FUNKTIONSWEISE:
# Google Drive for Desktop synchronisiert Dateien zwischen lokalen
# Laufwerken und der Cloud. Dateinamen, die unter Windows oder in
# der Google-Cloud ungültige Zeichen enthalten, führen zu Sync-
# Fehlern, beschädigten Pfaden oder unsichtbaren Dateinamen.
#
# Dieses Skript durchsucht ein Verzeichnis rekursiv und bereinigt
# alle Datei- UND Ordnernamen nach folgenden Regeln:
#
#   1. Mojibake-Reparatur (cp437/cp850/cp1252 → UTF-8 Round-Trip)
#      Heilt Namen wie 'u╠ê', 'ΓÇ░', '┬░', '╞Æ', die durch falsche
#      Codepage-Interpretation in NTFS gelandet sind. Generischer
#      Heiler + domänenspezifische Token-Whitelist (z. B. 'D‰mm' →
#      'Dämm'). Greift nur bei Verdachts-Codepoints und nur wenn
#      die Reparatur den Verdacht messbar reduziert.
#   2. Unicode-NFC-Normalisierung (macOS-NFD → Windows-NFC)
#   3. Unsichtbare Unicode-Zeichen entfernen
#      (Zero-Width Spaces, BOM, Soft Hyphen, Bidi-Marker)
#   4. Steuerzeichen durch Leerzeichen ersetzen (Tab, LF, CR)
#   5. Unzulässige Windows-Zeichen ersetzen
#      (< > : " / \ | ? * → '-')
#   6. Mehrfach-Leerzeichen zu einem Leerzeichen zusammenfassen
#   7. Nachfolgende Leerzeichen und Punkte entfernen
#   8. Reservierte MS-DOS-Namen mit '_' präfixieren
#      (CON, PRN, AUX, NUL, COM1-COM9, LPT1-LPT9)
#   9. Leere Namen nach Bereinigung → '_bereinigt'
#  10. Optional: Namen > 255 Zeichen kürzen mit Hash-Suffix
#
# Die Originaldateien werden NICHT verändert – nur der Dateiname
# auf der Platte wird angepasst. Der Dateiinhalt bleibt unberührt.
#
# Bei AV-Scanner- oder Indexer-Locks wird die Umbenennung mehrfach
# mit exponentiellem Backoff wiederholt, bevor ein Eintrag als
# Fehler markiert wird.
#
# UMWANDLUNG IN .EXE (PyInstaller):
#   pip install pyinstaller
#   python -m PyInstaller --onefile --console --clean --noconfirm --icon=python_icon.ico 10_dateinamen_bereinigen.py
# ==================================================================

from __future__ import annotations

import argparse
import hashlib
import logging
import os
import re
import sys
import time
import unicodedata
import uuid
from collections import Counter
from datetime import datetime
from enum import Enum, auto

if os.name == "nt":
    import winreg

# ==================================================================
# UTF-8-Ausgabe erzwingen
# ==================================================================
if sys.stdout is not None and getattr(sys.stdout, "encoding", None) \
        and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass
if sys.stderr is not None and getattr(sys.stderr, "encoding", None) \
        and sys.stderr.encoding.lower() != "utf-8":
    try:
        sys.stderr.reconfigure(encoding="utf-8")
    except AttributeError:
        pass

# ==================================================================
# tqdm-Import
# ==================================================================
try:
    from tqdm import tqdm
except ImportError:
    print("=" * 66)
    print("❌ FEHLENDES MODUL:")
    print("   pip install tqdm")
    print("=" * 66)
    sys.exit(1)


# ==================================================================
# Status-Enum
# ==================================================================
class RenameStatus(Enum):
    RENAMED       = auto()
    CLEAN         = auto()
    SKIPPED       = auto()
    ERROR         = auto()
    WOULD_RENAME  = auto()   # Probelauf: waere umbenannt worden
    EXCLUDED      = auto()   # in einem ausgeschlossenen Verzeichnis


# Nur diese Ergebnisse gehoeren in die Fortschrittsdatei: sie sind dauerhaft.
# ERROR (z. B. WinError 32 durch AV-Scanner) und SKIPPED (alle Ausweichnamen
# belegt) koennen sich beim naechsten Lauf aufloesen - wer sie vermerkt,
# ueberspringt sie fuer immer.
_RESUME_STATUSES = frozenset({
    RenameStatus.RENAMED,
    RenameStatus.CLEAN,
    RenameStatus.EXCLUDED,
    RenameStatus.WOULD_RENAME,
})


# ==================================================================
# Konfiguration
# ==================================================================

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

MAX_PATH_LEN = 259
MAX_NAME_LEN = 255
HASH_LEN     = 8

RENAME_RETRIES    = 5
RENAME_BASE_DELAY = 0.4
# Getrennte, kleinere Retry-Zahl fuer ERROR_ACCESS_DENIED (WinError 5).
# Ein AV-Scanner liefert diesen Code kurzzeitig, eine fehlende ACL dauerhaft
# - und im zweiten Fall sind fuenf Versuche mit exponentiellem Backoff
# 6 Sekunden pro Eintrag. Bei 100.000 nicht beschreibbaren Dateien waren
# das ueber 160 Stunden reine Wartezeit.
DENIED_RETRIES    = 2
EXISTS_RETRIES    = 2
EXISTS_DELAY      = 0.2

SKIP_FILES         = {".ds_store", "thumbs.db", "desktop.ini"}
SKIP_FILE_PREFIXES = ("~$", "._")
SKIP_DIR_PREFIXES  = ("._",)

# Verzeichnisse, deren Inhalt NIE umbenannt werden darf. Ohne diese Liste
# lief das Skript auch durch Versionsverwaltung und Systemordner. Das ist
# nicht bloss unnoetig, sondern zerstoerend: in einem unter Linux oder
# macOS erzeugten Git-Repository heissen Referenzen unter '.git/refs/'
# durchaus 'feature:name' - das Doppelpunkt-Ersetzen durch '-' macht das
# Repository unbrauchbar. Gleiches gilt fuer Paketverzeichnisse, deren
# Dateinamen Teil eines Hashes oder einer Sperrdatei sind.
EXCLUDE_DIR_NAMES = {
    ".git", ".svn", ".hg", ".bzr",
    "$recycle.bin", "system volume information", "recycler",
    "node_modules", "__pycache__", ".venv", "venv",
    ".vs", ".idea",
}

# Verzeichnis-Praesets fuer die Startauswahl. Leere Eintraege werden im
# Menue ausgeblendet, die Nummerierung bleibt trotzdem lueckenlos.
DIRECTORY_PRESETS = (
    "Q:\\",
    "R:\\",
    "G:\\Geteilte Ablagen",
    "G:\\Meine Ablage",
    "\\\\server\\dfs",
)

if getattr(sys, "frozen", False):
    _base_dir = os.path.dirname(sys.executable)
else:
    _base_dir = os.path.dirname(os.path.abspath(__file__))
_log_ts  = time.strftime("%Y-%m-%d_%H%M%S")


def _pick_output_dir(preferred: str) -> str:
    """Erstes beschreibbares Verzeichnis aus der Kandidatenliste.

    Vorher wurde ungeprueft neben das Skript geschrieben. Lag es auf einer
    schreibgeschuetzten Freigabe, warf logging.FileHandler in setup_logging()
    ungefangen - der Anwender bekam direkt nach dem Banner einen Traceback,
    obwohl das Protokoll gar nicht dort liegen muss. Der Schreibtest ist
    echt (Datei anlegen und loeschen), denn ein vorhandenes Verzeichnis
    sagt nichts ueber das Schreibrecht aus.
    """
    candidates = [preferred]
    local = os.environ.get("LOCALAPPDATA")
    if local:
        candidates.append(os.path.join(local, "FilenameCleanup"))
    tmp = os.environ.get("TEMP") or os.environ.get("TMP")
    if tmp:
        candidates.append(tmp)
    for cand in candidates:
        if not cand:
            continue
        try:
            os.makedirs(cand, exist_ok=True)
            probe = os.path.join(cand, f".writetest_{uuid.uuid4().hex}.tmp")
            with open(probe, "w", encoding="utf-8") as fh:
                fh.write("x")
            os.remove(probe)
            return cand
        except OSError:
            continue
    return preferred


OUTPUT_DIR = _pick_output_dir(_base_dir)
LOG_FILE = os.path.join(OUTPUT_DIR, f"10_dateinamen_bereinigen_{_log_ts}.log")

RENAME_LOG_FILE = os.path.join(
    OUTPUT_DIR, f"10_dateinamen_bereinigen_{_log_ts}_umbenennungen.csv"
)

logger = logging.getLogger("FilenameLogger")

# Maschinenlesbares Protokoll aller Umbenennungen. Das Textlog haelt die
# Vorgaenge nur in Prosa fest - bei mehreren hunderttausend Umbenennungen
# liess sich daraus kein Rueckgaengigmachen ableiten. Die CSV enthaelt
# Verzeichnis, alten und neuen Namen jeweils als eigenes Feld und ist
# damit direkt fuer ein Rollback-Skript verwendbar.
_rename_log_handle = None


def open_rename_log() -> None:
    global _rename_log_handle
    try:
        # errors="surrogatepass": Dateinamen koennen halbe Surrogatpaar-
        # Haelften enthalten (\udc80 & Co.) - Windows laesst solche Namen zu,
        # und das Skript behandelt sie weiter unten ausdruecklich. Mit der
        # strikten Vorgabe warf schon das Schreiben der Protokollzeile einen
        # UnicodeEncodeError, und genau fuer diese Umbenennung fehlte danach
        # die einzige Ruecknahmequelle.
        _rename_log_handle = open(
            RENAME_LOG_FILE, "w", encoding="utf-8-sig", newline="",
            errors="surrogatepass"
        )
        _rename_log_handle.write("Status;Typ;Verzeichnis;AlterName;NeuerName\n")
    except OSError as e:
        _rename_log_handle = None
        logger.warning(f"Umbenennungsprotokoll nicht schreibbar: {e}")


def close_rename_log() -> None:
    global _rename_log_handle
    if _rename_log_handle is not None:
        try:
            _rename_log_handle.close()
        except OSError:
            pass
        _rename_log_handle = None


def _csv_field(value: str) -> str:
    v = (value or "").replace("\r", " ").replace("\n", " ")
    if ";" in v or '"' in v:
        return '"' + v.replace('"', '""') + '"'
    return v


# ==================================================================
# Gemeinsames Laufprotokoll (migration.jsonl)
# ==================================================================
# Jedes Skript schreibt sein eigenes Format. Der Gesamtfortschritt ueber
# die elf Schritte liess sich damit nicht auswerten. Diese Zeile
# ERGAENZT das bestehende Protokoll, ersetzt es nicht.
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


def log_rename(status: str, is_dir: bool, directory: str,
               old_name: str, new_name: str) -> None:
    _protokoll(os.path.join(directory, old_name),
               "umbenannt" if new_name else "geprueft",
               status,
               f"neu: {new_name}" if new_name else "")
    if _rename_log_handle is None:
        return
    row = ";".join(_csv_field(x) for x in (
        status,
        "Ordner" if is_dir else "Datei",
        _display_path(directory),
        old_name,
        new_name,
    ))
    try:
        _rename_log_handle.write(row + "\n")
    except Exception as e:
        # Bewusst 'Exception', nicht 'OSError': ein UnicodeEncodeError ist
        # eine Unterklasse von ValueError und lief hier frueher vorbei - bis
        # hinauf in den Umbenennungs-try, der die bereits ausgefuehrte
        # Umbenennung dann als Fehler meldete.
        logger.warning(f"Protokollzeile nicht schreibbar: {e!r}")


# ==================================================================
# Logging
# ==================================================================
def setup_logging() -> None:
    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    logger.setLevel(logging.DEBUG)
    try:
        # errors="surrogatepass" aus demselben Grund wie beim CSV: Dateinamen
        # mit halben Surrogatpaar-Haelften duerfen die Protokollierung nicht
        # sprengen.
        fh = logging.FileHandler(LOG_FILE, mode="a", encoding="utf-8",
                                 errors="surrogatepass")
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    except OSError as e:
        # Ohne Datei-Handler weiterlaufen ist besser als abbrechen, aber der
        # Anwender muss wissen, dass es kein Protokoll gibt.
        print(f"\n⚠  WARNUNG: Logdatei nicht schreibbar ({e}).")
        print("   Der Lauf wird NICHT protokolliert.")
        logger.addHandler(logging.NullHandler())
        return
    logger.info("=" * 66)
    logger.info("NEUE SESSION GESTARTET")
    logger.info("=" * 66)


# ==================================================================
# Long-Path-Hilfsfunktionen
# ==================================================================
def prepare_long_path(path: str) -> str:
    if os.name != "nt":
        return os.path.abspath(path)
    if path.startswith("\\\\?\\"):
        return path
    path = os.path.abspath(os.path.normpath(path))
    if path.startswith("\\\\"):
        return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path


def _lp(path: str) -> str:
    if os.name != "nt":
        return path
    if path.startswith("\\\\?\\"):
        return path
    if len(path) > MAX_PATH_LEN:
        return prepare_long_path(path)
    return path


def _display_path(path: str) -> str:
    if not path:
        return path
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\"):
        return path[4:]
    return path


def safe_exists(path: str) -> bool:
    # os.path.exists() schluckt OSError stillschweigend (gibt False zurück),
    # was bei Netzwerkstörungen zu falscher Kollisionsauflösung führt.
    # os.stat() reicht OSErrors korrekt durch — nur FileNotFoundError ist
    # die saubere "existiert nicht"-Antwort.
    last_err = None
    for attempt in range(EXISTS_RETRIES):
        try:
            os.stat(_lp(path))
            return True
        except FileNotFoundError:
            return False
        except OSError as e:
            last_err = e
            time.sleep(EXISTS_DELAY)
    # Bei anhaltender Unsicherheit pessimistisch annehmen, die Datei
    # existiere — das löst Kollisionsauflösung mit Suffix aus, statt
    # blind in einen evtl. besetzten Zielnamen zu rennen.
    logger.debug(f"safe_exists fehlgeschlagen für {_display_path(path)}: {last_err}")
    return True


# ==================================================================
# User Shell Folders (Registry-Lookup mit Fallback)
# ==================================================================
def _resolve_user_shell_folder(value_name: str, fallback_subdir: str) -> str | None:
    home = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    fallback = os.path.join(home, fallback_subdir) if home else None

    if os.name == "nt":
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
            ) as key:
                value, _ = winreg.QueryValueEx(key, value_name)
                expanded = os.path.expandvars(value)
                if os.path.isdir(expanded):
                    return expanded
        except OSError:
            pass

    if fallback and os.path.isdir(fallback):
        return fallback
    return None


def get_desktop_path() -> str | None:
    return _resolve_user_shell_folder("Desktop", "Desktop")


def get_downloads_path() -> str | None:
    return _resolve_user_shell_folder(
        "{374DE290-123F-4565-9164-39C4925E467B}", "Downloads"
    )


# ==================================================================
# Pfad-Hilfsfunktionen
# ==================================================================
def safe_isdir(path: str) -> bool:
    """Verzeichnispruefung, die nie eine Ausnahme nach aussen laesst.

    Vorher war diese Absicherung nur im interaktiven Zweig vorhanden. Der
    Aufruf mit --dir pruefte direkt mit os.path.isdir(prepare_long_path(...)),
    und bei ungueltigen Systemzeichen oder Geraeteformaten (z.B. '?:\\x')
    werfen abspath/isdir auf API-Ebene - der Anwender bekam einen
    Traceback statt einer Fehlermeldung.
    """
    if not path:
        return False
    try:
        return os.path.isdir(prepare_long_path(path))
    except (OSError, ValueError):
        return False


def sanitize_path(raw: str) -> str:
    path = raw.strip().strip('"').strip("'")
    if len(path) == 2 and path[1] == ":":
        path = path + "\\"
    return path


def is_excluded_path(path: str) -> bool:
    """True, wenn eine Komponente des Pfads in EXCLUDE_DIR_NAMES steht.

    Die Pruefung laeuft ueber den ganzen Pfad, nicht nur den letzten
    Bestandteil: os.walk() wird hier mit topdown=False aufgerufen, und in
    diesem Modus hat das Kuerzen der 'dirs'-Liste keine Wirkung mehr - der
    Abstieg ist zu diesem Zeitpunkt schon erfolgt. Ausgeschlossene Baeume
    werden daher nicht uebersprungen, sondern nur von jeder Umbenennung
    ausgenommen.
    """
    p = _display_path(path)
    for part in p.replace("/", "\\").split("\\"):
        if part.lower() in EXCLUDE_DIR_NAMES:
            return True
    return False


def ask_directory() -> str:
    desktop_dir   = get_desktop_path()
    downloads_dir = get_downloads_path()

    choices: list[tuple[str, str]] = []
    for preset in DIRECTORY_PRESETS:
        if preset and preset.strip():
            choices.append((preset, preset))
    if desktop_dir:
        choices.append((desktop_dir, f"Desktop          ({desktop_dir})"))
    if downloads_dir:
        choices.append((downloads_dir, f"Downloads        ({downloads_dir})"))
    manual_idx = len(choices) + 1

    print("\nZielverzeichnis auswählen:")
    for idx, (_, label) in enumerate(choices, start=1):
        print(f"  [{idx}] {label}")
    print(f"  [{manual_idx}] Eigenen Pfad eingeben")
    print()

    while True:
        try:
            choice = input(f"Auswahl [1-{manual_idx}]: ").strip()
        except EOFError:
            print("\n❌ Fehler: Kein interaktives Terminal. Bitte --dir verwenden.")
            sys.exit(1)

        if choice.isdigit() and 1 <= int(choice) <= len(choices):
            selected = choices[int(choice) - 1][0]
            # Preset-Pfade koennen auf diesem Rechner fehlen (Laufwerk nicht
            # gemappt). Vorher wurden sie ungeprueft zurueckgegeben und das
            # Skript brach erst spaeter mit einer unklaren Meldung ab.
            if not safe_isdir(selected):
                print(f"  ❌ '{selected}' existiert nicht oder ist kein Verzeichnis.")
                print("     Hinweis: Bei erhoehten Rechten sind gemappte Netzlaufwerke")
                print("     oft ausgeblendet - dann den UNC-Pfad manuell eingeben.")
                continue
            return selected
        if choice == str(manual_idx):
            try:
                raw = input("Pfad eingeben: ")
            except EOFError:
                print("\n❌ Fehler: Kein interaktives Terminal. Bitte --dir verwenden.")
                sys.exit(1)
            path = sanitize_path(raw)
            if not path:
                # Leere Eingabe würde via prepare_long_path("") →
                # abspath("") zum aktuellen Arbeitsverzeichnis aufgelöst
                # und isdir gäbe True zurück - das Skript würde sein
                # eigenes Ausführungsverzeichnis bearbeiten. Hart ablehnen.
                print("  ❌ Eingabe darf nicht leer sein.")
                continue
            if safe_isdir(path):
                return path
            print(f"  ❌ Verzeichnis nicht gefunden oder ungültig: '{path}'")
            continue
        print(f"  Bitte 1–{manual_idx} eingeben.")


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


# ==================================================================
# Dateinamen-Bereinigung
# ==================================================================
_CONTROL_RE = re.compile(r"[\x00-\x1f\x7f]")

_INVISIBLE_RE = re.compile(
    r"[\u200b\u200c\u200d\u2060"
    r"\ufeff"
    r"\u00ad"
    r"\u200e\u200f"
    r"\u202a-\u202e"
    r"\u2066-\u2069"
    r"]"
)

_INVALID_CHARS_RE = re.compile(r'[<>:"/\\|?*]+')

# Piktogramm-Symbole, die in seriellen Datei- und Ordnernamen kein
# legitimes Vorkommen haben und vom Skript ohne Ersatz entfernt werden.
# Decken Misc Symbols (☎, ☂, ★, ♥, ...), Dingbats (✂, ✈, ✓, ✗, ...),
# Pfeile, Geometrische Formen, Misc Pictographs, Emoji-Bereich und die
# Private Use Area (Wingdings/Symbol-Font-Glyphen wie U+F028=Telefon,
# U+F022=Schere) ab. Latein-Zeichen, Akzente, Mathematik, ©®™, Währung,
# Bruchzahlen und Box-Drawing (für Mojibake-Zwischenformen) bleiben
# unangetastet.
_PICTOGRAM_RE = re.compile(
    r"[\u2600-\u26ff"        # Misc Symbols (☎ ☂ ★ ♥ ...)
    r"\u2700-\u27bf"          # Dingbats (✂ ✈ ✓ ✗ ...)
    r"\ue000-\uf8ff"          # Private Use Area (Wingdings: Telefon, Schere, ...)
    r"\U0001f300-\U0001f9ff"  # Misc Pictographs + Emoticons + Transport + Symbols
    r"\U0001fa00-\U0001faff"  # Symbols and Pictographs Extended-A
    r"]"
)

_MULTI_SPACE_RE = re.compile(r" {2,}")

_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


# ------------------------------------------------------------------
# Mojibake-Heiler
# ------------------------------------------------------------------
# Verdachts-Codepoints für die ANFANGS-Erkennung ('soll Heiler überhaupt
# laufen?'). Großzügig: schließt auch legitime deutsche Umlaute mit ein,
# was OK ist weil dieser Trigger nur entscheidet ob der Heiler aktiv wird.
_MOJIBAKE_HINT_CHARCLASS = (
    r"\u0080-\u009f"        # C1-Steuerbereich
    r"\u00a0-\u00bf"         # Latin-1 Punctuation/Symbols (¦, ©, ®, °, ·, ±, …)
    r"\u00c0-\u00ff"         # Latin-1-Sup (in Mojibake-Kombis)
    r"\u0192"                # ƒ (Florin)
    r"\u0391-\u03a9"         # Griech. Großbuchstaben (Γ, Σ, ...)
    r"\u2030"                # ‰ (Promille)
    r"\u2190-\u21ff"         # Pfeile
    r"\u2500-\u259f"         # Box-Drawing + Block
)
_MOJIBAKE_HINT_RE = re.compile(f"[{_MOJIBAKE_HINT_CHARCLASS}]")

# Strikterer Hint für die Akzeptanz-Prüfung von Heilungs-Ergebnissen.
# Identisch mit _MOJIBAKE_HINT_RE, aber legitime deutsche Umlaute (ÄÖÜß äöü)
# zählen NICHT als Verdacht. Damit akzeptiert _try_heal Ergebnisse wie 'Ä',
# 'ü', 'ö', 'ß' als saubere Heilung — vorher wurden sie als 'noch verdächtig'
# abgewiesen.
_MOJIBAKE_RESULT_HINT_RE = re.compile(
    r"[\u0080-\u009f"        # C1
    r"\u00a0-\u00bf"          # Latin-1 Punctuation/Symbols
    r"\u00c0-\u00c3"          # À, Á, Â, Ã (NICHT Ä)
    r"\u00c5-\u00d5"          # Å bis Õ (NICHT Ö)
    r"\u00d7-\u00db"          # × bis Û (NICHT Ü)
    r"\u00dd-\u00de"          # Ý, Þ (NICHT ß)
    r"\u00e0-\u00e3"          # à, á, â, ã (NICHT ä)
    r"\u00e5-\u00f5"          # å bis õ (NICHT ö)
    r"\u00f7-\u00fb"          # ÷ bis û (NICHT ü)
    r"\u00fd-\u00ff"          # ý, þ, ÿ
    r"\u0192"                 # ƒ
    r"\u0391-\u03a9"          # Griech. Großbuchstaben
    r"\u2030"                 # ‰
    r"\u2190-\u21ff"          # Pfeile
    r"\u2500-\u259f]"         # Box-Drawing + Block
)

# Cluster-Susp-Charclass für das CLUSTER-MATCHING. Schließt legitime
# deutsche Umlaute (Ä Ö Ü ß ä ö ü) explizit aus, damit ein Mojibake-
# Cluster wie '├ƒ' nicht versehentlich ein legitimes 'ö' im Wort 'rö├ƒ'
# mitgrabbt und dadurch unheilbar wird.
_CLUSTER_SUSP_CHARCLASS = (
    r"\u0080-\u009f"        # C1
    r"\u00a0-\u00bf"         # Latin-1 Punctuation/Symbols
    r"\u00c0-\u00c3"         # À, Á, Â, Ã (NICHT Ä=U+00C4)
    r"\u00c5-\u00d5"         # Å bis Õ (NICHT Ö=U+00D6)
    r"\u00d7-\u00db"         # × bis Û (NICHT Ü=U+00DC)
    r"\u00dd-\u00de"         # Ý, Þ (NICHT ß=U+00DF)
    r"\u00e0-\u00e3"         # à, á, â, ã (NICHT ä=U+00E4)
    r"\u00e5-\u00f5"         # å bis õ (NICHT ö=U+00F6)
    r"\u00f7-\u00fb"         # ÷ bis û (NICHT ü=U+00FC)
    r"\u00fd-\u00ff"         # ý, þ, ÿ
    r"\u0192"                # ƒ
    r"\u0391-\u03a9"         # Griech. Großbuchstaben
    r"\u2013-\u2014"         # En-Dash, Em-Dash (UTF-8-als-Latin1 Folgebytes)
    r"\u2018-\u201f"         # Smart Quotes („"…) – Folgebytes für Ä Å Ç É etc.
    r"\u2030"                # ‰
    r"\u2039-\u203a"         # ‹ › (UTF-8-als-Latin1 Folgebytes)
    r"\u20a0-\u20cf"         # Currency Symbols (₧, € — als Mojibake-Folgebytes)
    r"\u2190-\u21ff"         # Pfeile
    r"\u2500-\u259f"         # Box-Drawing + Block
)

# Cluster: optional ein ASCII-Buchstabe (für NFD-Combining-Marks, deren
# semantischer Träger der Vorgängerbuchstabe ist) plus ≥1 verdächtiger
# Codepoint. Die Alternation erlaubt zusätzlich '╠X' bzw. '¦X' (X = Latin-1-
# Sup-Zeichen) als Bestandteil eines Clusters, weil das NFD-Combining-Mark-
# Vervollständigungen sind (e╠ü → é, a╠ê → ä, o¦ê → ö). Außerhalb solcher
# Combining-Sequenzen werden Latin-1-Zeichen (insb. legit. ä ö ü) NICHT
# mitkonsumiert — wichtig zum Schutz vor "rö├ƒ"-Greedy-Mismatches.
_MOJIBAKE_CLUSTER_RE = re.compile(
    f"(?P<prefix>[A-Za-z])?"
    f"(?P<susp>(?:[\u2560\u00a6][\u00a0-\u00ff]|[{_CLUSTER_SUSP_CHARCLASS}])+)"
)

# Reparaturketten in Reihenfolge der Wahrscheinlichkeit. cp1252 ZUERST,
# weil das der korrekte Pfad für UTF-8-als-Latin1-Mojibake ist (Ã¼→ü,
# Ã¤→ä). cp437/cp850 fangen Box-Drawing-Mojibake auf — bei denen schlägt
# cp1252 ohnehin fehl, weil Box-Drawing-Zeichen nicht in cp1252 codierbar
# sind. Reihenfolge falsch herum würde dazu führen, dass cp850 zufällig
# "akzeptable" aber falsche Latin-Extended-B-Zeichen produziert (z. B.
# Ã¼ → Ǭ statt ü).
_REPAIR_CHAINS = (("cp1252", "utf-8"), ("cp437", "utf-8"), ("cp850", "utf-8"))

# Akzeptanz-Whitelist für Heilungs-Ergebnisse: westeuropäisches Latin
# plus übliche typografische Symbole. Schließt Latin Extended-B aus
# (U+0180-U+024F außer ƒ), weil dort selten gebrauchte Zeichen wie Ǭ oder
# Ǆ liegen, die cp850 fälschlich aus utf-8-Mojibake produziert. Mojibake-
# Zwischenformen ƒ (U+0192) und ‰ (U+2030) sind explizit erlaubt, weil sie
# durch die Whitelist (Stufe 2) weiter zu deutschen Umlauten heilen.
_ACCEPTABLE_RESULT_RE = re.compile(
    r"^["
    r"\u0020-\u007e"        # druckbares ASCII
    r"\u00a0-\u017f"         # Latin-1 Sup + Latin Extended-A
    r"\u0192"                # ƒ (Mojibake-Zwischenform für ä/ü)
    r"\u0300-\u036f"         # Combining Marks (NFD-Zwischenform)
    r"\u2000-\u206f"         # Allgemeine Interpunktion
    r"\u2030"                # ‰ (Mojibake-Zwischenform für E)
    r"\u2070-\u209f"         # Sub-/Superscript
    r"\u20a0-\u20cf"         # Währungssymbole
    r"\u2100-\u214f"         # Letterlike (™, ©, …)
    r"\u2150-\u218f"         # Zahlenformen
    r"\u2200-\u22ff"         # Mathematik
    r"\u2300-\u23ff"         # Misc Technical
    r"\u2500-\u259f"         # Box-Drawing (Mojibake-Zwischenform)
    r"\u25a0-\u26ff"         # Geom Shapes + Misc Symbols
    r"]*$"
)

# Domänenspezifische Reste, die der generische Heiler nicht eindeutig
# auflösen kann (mehrstufiges Mojibake aus Drittsoftware, z. B. CAD/DTP).
# Token-genau, NICHT zeichenweise, damit legitime '‰'/'ƒ'/'·' überleben.
# Sowohl Klein- als auch Großschreibung, weil .replace() Case-Sensitive ist.
# Bei Bedarf erweitern.
_POST_REPAIR_TOKENS = {
    # --- Diaeresis-Substitut U+00A6 + ê (NFD-Variante, generisch nicht heilbar) ---
    "a¦ê": "ä", "A¦ê": "Ä",
    "o¦ê": "ö", "O¦ê": "Ö",
    "u¦ê": "ü", "U¦ê": "Ü",

    # --- Drittsoftware-Substitut '+§' für 'ä' (nach Phase A: '├§')
    # Pfad nicht durch Codepage-Roundtrip erklärbar, aber semantisch eindeutig.
    "├§": "ä",

    # --- Domänen-Tokens: 'Dämm...' aus drei Mojibake-Quellen (.pat-Dateien) ---
    "D‰mm":  "Dämm",  "D‰MM":  "DÄMM",
    "Dƒmm":  "Dämm",  "DƒMM":  "DÄMM",
    "d╟╧mm": "dämm",  "D╟╧mm": "Dämm",  "D╟╧MM": "DÄMM",

    # --- 'ƒ' direkt als Umlaut-Substitut (Drittsoftware-Mutation, NICHT cp437→utf-8) ---
    # Konflikt: dasselbe ƒ steht für ä in Dämm-Wörtern und für ü in Brück-Wörtern.
    # Daher Token-genau (NICHT generisch), längere Tokens haben Vorrang im Dict.
    "Brƒck":  "Brück",   # Brücke, Brücken, Brückenbau
    "brƒck":  "brück",   # Severinsbrücke, Müngstenerbrücke (Compounds)
    "Mƒngst": "Müngst",  # Müngstener Brücke

    # --- '‰' allein als 'E'-Substitut (Drittsoftware-Mutation) ---
    "‰rd":    "Erd",     # Erdzeitalter
    "‰RD":    "ERD",

    # --- 'Ñ'/'Å' als 'ä'/'ü'-Substitut (Drittsoftware-Mutation, gleicher
    # Pfad wie 'fÅr'→'für'). Token-genau, NICHT als generisches Ñ→ä,
    # weil in skandinavischen/spanischen Namen Ñ und Å legitim sind
    # (España, Niño, Åke, Århus). Hier nur deutsche Wort-Tokens.
    "AuftrÑge":         "Aufträge",
    "AUFTRÑGE":         "AUFTRÄGE",
    "GaststreitkrÑfte": "Gaststreitkräfte",
    "krÑfte":           "kräfte",
    "KrÑfte":           "Kräfte",

    # --- '┬╧' Drittsoftware-Substitut für 'ß' (Oßberger) ---
    "┬╧": "ß",

    # --- Stufe-1-Heilung ergibt das falsche Zeichen (Drittsoftware-Mutation) ---
    # 'VerstN╠ârker' heilt zu 'VerstÑrker' → semantisch 'Verstärker'
    "VerstÑrker": "Verstärker",
    # 'fA╠èr' heilt zu 'fÅr' → semantisch deutsches 'für'
    "fÅr":        "für",
    # --- 'Ma┬Àe' heilt zu 'Ma·e' → semantisch 'Maße' (Abmessungen)
    "Ma·e":       "Maße",

    # --- '„' (U+201E, deutsches Anführungszeichen unten) als 'ä'-Substitut.
    # Entsteht aus 'ΓÇ₧' nach generischer cp437→utf-8-Heilung. NICHT
    # generisch heilen, weil „ in Dateinamen oft als legitimes deutsches
    # Anführungszeichen genutzt wird (z. B. 'Wanderausstellung „energie"').
    # Stattdessen kontextspezifische Wort-Tokens für die im Bestand
    # vorkommenden Mojibake-Wörter.
    "Fl„ch":     "Fläch",     # Fläche, Flächen, Metalloberflächen
    "fl„ch":     "fläch",     # Compounds
    "Pl„n":      "Plän",      # Pläne
    "Durchg„ng": "Durchgäng", # Durchgänge
    "Infos„ul":  "Infosäul",  # Infosäule

    # --- '─' (U+2500, Box-Drawing) als 'Ä'-Substitut (Drittsoftware-
    # Mutation). Im Bestand systematisch am Wortanfang (─nderungen,
    # ─NDERUNGEN). Token mit Buchstaben-Kontext, weil U+2500 theoretisch
    # auch in ASCII-Art-Dateinamen legitim sein könnte.
    # '_─' und ' ─' hatten entgegen dem Kommentar KEINEN Buchstaben-Kontext:
    # sie trafen jedes U+2500 nach Unterstrich oder Leerzeichen, also auch
    # eine reine Trennlinie ('Bericht ───── Anhang.pdf' wurde zu
    # 'Bericht Ä──── Anhang.pdf'). Deshalb den folgenden Buchstaben
    # mitfordern - genau wie bei den Tokens darunter.
    "_─n": "_Än",   # _─nderungen → _Änderungen
    "_─N": "_ÄN",
    " ─n": " Än",   # mit Leerzeichen davor
    " ─N": " ÄN",
    "─n": "Än",   # ─nderungen am Stringanfang
    "─N": "ÄN",   # ─NDERUNGEN
    "─ä": "Ää",   # falls am Anfang vor Wortteil — sehr unwahrscheinlich aber harmlos

    # --- Replacement-Char-Substitut '∩┐╜' (= UTF-8-Bytes von U+FFFD,
    # 'I lost data') als 'ü' in Wort-Position. Token-genau für deutsche
    # Wörter, weil U+FFFD theoretisch alles bedeuten kann.
    "t∩┐╜r":    "tür",
    "T∩┐╜r":    "Tür",

    # --- '▌ê' Drittsoftware-Substitut für 'Ü' (Überfluss).
    # Nur am Wortanfang sinnvoll; Token mit Großbuchstabe vorn schützt.
    "U▌ê": "Ü",

    # --- '┬ä'/'┬ü' Drittsoftware-Substitute. Doppeldeutig: in manchen
    # Fällen ä/ü, in anderen Müll, der weg soll. Daher kontextspezifische
    # Wort-Tokens, NICHT generisch.
    "Eigenerklae┬ärung":  "Eigenerklärung",
    "Anha╠ê┬änge":         "Anhänge",
    "f┬üuer":              "für",
    "R┬üM":                "RM",     # ┬ü als Müll, einfach weg
}


def _try_heal(fragment: str) -> str | None:
    """Versucht ein Mojibake-Fragment durch Round-Trip zu heilen.
    Liefert das geheilte Fragment (NFC) zurück, wenn alle Akzeptanz-
    Kriterien erfüllt sind, sonst None.
    """
    for src, dst in _REPAIR_CHAINS:
        try:
            cand = fragment.encode(src).decode(dst)
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        cand_nfc = unicodedata.normalize("NFC", cand)
        # Akzeptanzkriterien: keine Verdachts-Codepoints mehr übrig,
        # kein Replacement-Char, keine Steuerzeichen, Ergebnis bleibt
        # im erlaubten Latin-Bereich.
        if (not _MOJIBAKE_RESULT_HINT_RE.search(cand_nfc)
                and "\ufffd" not in cand_nfc
                and not _CONTROL_RE.search(cand_nfc)
                and _ACCEPTABLE_RESULT_RE.match(cand_nfc)):
            return cand_nfc
    return None


# Legitime deutsche Umlaute. Diese sieben Zeichen sind in Dateinamen
# extrem häufig legitim und dürfen daher die Plus-Substitution (Phase A)
# nicht auslösen. Damit bleiben Konstrukte wie 'Einbruch+Überfall' oder
# 'Schimmel+Asbest+Öle' (in denen '+' für "und" steht) unverändert.
_LEGIT_GERMAN_UMLAUTS = set("äöüÄÖÜß")


def _pre_substitute_plus(text: str) -> str:
    """Substitut für '+' als Box-Drawing-Substitut (Drittsoftware-Mutation).

    Manche Konvertierungs-Software konnte Box-Drawing-Zeichen nicht anzeigen
    und ersetzte sie durch ASCII-'+'. Wir machen das rückgängig, aber nur
    wenn der Name andere Mojibake-Indikatoren trägt (sonst bleiben legitime
    'C++_Tutorial.pdf' und '1+1=2.txt' unangetastet).

    Phase A: isolierte '+' direkt neben Verdachts-Zeichen → '├'.
             Heilt z. B. '+ƒ'→ß, '+£'→Ü, '+ñ'→ä, '+Â'→ö, '+û'→Ö.
             AUSGENOMMEN: legitime deutsche Umlaute (äöüÄÖÜß) als Nachbar.
             Dort steht '+' meist für "und" (z. B. 'Einbruch+Überfall' =
             'Einbruch und Überfall'). Da diese Umlaute massenhaft legitim
             vorkommen, wird das '+' unangetastet gelassen. Andere nicht-
             ASCII-Zeichen (ƒ, £, ñ, Â, û, ...) sind in Dateinamen
             praktisch immer Mojibake-Indikatoren.

    Phase B: '++' zwischen ASCII-Buchstaben → '├╝' (= UTF-8-Bytes für 'ü').
             Heilt z. B. 'Fl++gel'→Flügel, 'k++hler'→kühler.
             Annahme: '++' steht im Bestand für ü (cp437-Pfad). Falls 'ö'
             gemeint ist, kann der User das per Whitelist-Eintrag korrigieren.
    """
    if not _MOJIBAKE_HINT_RE.search(text):
        return text

    # Phase A: isolierte + neben Verdachts-Zeichen (NICHT in ++-Sequenzen)
    chars = list(text)
    n = len(chars)

    def _is_mojibake_neighbor(ch: str | None) -> bool:
        # Nachbar ist nur dann ein Mojibake-Indikator, wenn er
        # nicht-ASCII ist UND kein legitimer deutscher Umlaut ist.
        if ch is None or ord(ch) < 0x80:
            return False
        return ch not in _LEGIT_GERMAN_UMLAUTS

    for i, c in enumerate(chars):
        if c != "+":
            continue
        # Skip wenn Teil einer ++-Sequenz (Phase B kümmert sich darum)
        if (i > 0 and chars[i-1] == "+") or (i < n-1 and chars[i+1] == "+"):
            continue
        left  = chars[i-1] if i > 0 else None
        right = chars[i+1] if i < n-1 else None
        if _is_mojibake_neighbor(left) or _is_mojibake_neighbor(right):
            chars[i] = "\u251C"  # ├
    text2 = "".join(chars)

    # Phase B: ++ zwischen ASCII-Buchstaben → ├╝ (= ü in cp437→utf-8)
    #
    # Eigener, STRENGERER Torwaechter als der Gate am Funktionsanfang: dort
    # steht _MOJIBAKE_HINT_RE, dessen Zeichenklasse laut eigenem Kommentar
    # bewusst auch legitime deutsche Umlaute umfasst. Damit war der Gate bei
    # praktisch jedem deutschen Dateinamen mit Umlaut offen, und Phase B
    # wandelte jedes '++' zwischen zwei ASCII-Buchstaben um - obwohl der
    # Docstring das Gegenteil zusichert. Nachgestellt: 'Ausflug++Gruesse.pdf'
    # und 'Oel++Wasser.txt' wurden beide veraendert, nur weil anderswo im
    # Namen ein Umlaut stand. _MOJIBAKE_RESULT_HINT_RE nimmt die legitimen
    # Umlaute ausdruecklich aus und verlangt damit einen echten Beleg.
    if _MOJIBAKE_RESULT_HINT_RE.search(text2):
        text2 = re.sub(
            r"(?<=[A-Za-z])\+\+(?=[A-Za-z])",
            "\u251C\u255D",  # ├╝
            text2,
        )

    return text2


# Box-Drawing (U+2500-257F) und Blockelemente (U+2580-259F): ─ bis ▟.
# Bleibt eines davon nach dem Heilen stehen, war der Durchlauf erfolglos.
_BOXDRAW_RE = re.compile(r"[─-▟]")


def _repair_mojibake_core(text: str, plus_substitution: bool = True) -> str:
    if not _MOJIBAKE_HINT_RE.search(text):
        return text

    def _heal_match(m: re.Match[str]) -> str:
        prefix = m.group("prefix") or ""
        susp   = m.group("susp")
        full   = prefix + susp

        # 1. Versuch: Prefix einbeziehen (für NFD-Combining-Marks, bei
        #    denen der Vorgängerbuchstabe semantisch dazugehört)
        if prefix:
            healed = _try_heal(full)
            if healed is not None:
                return healed

        # 2. Versuch: nur den verdächtigen Teil heilen, Prefix bleibt
        healed = _try_heal(susp)
        if healed is not None:
            return prefix + healed

        # 3. Mehrstufige Heilung (z. B. ΓÇ░ → ‰): Stufe 1 produziert
        #    ein Zwischenergebnis, das selbst noch Verdacht enthält,
        #    aber im erlaubten Latin-Bereich liegt. Behalten — die
        #    Whitelist (Stufe 2) fängt es danach ab.
        for src, dst in _REPAIR_CHAINS:
            try:
                stage1 = full.encode(src).decode(dst)
            except (UnicodeEncodeError, UnicodeDecodeError):
                continue
            stage1_nfc = unicodedata.normalize("NFC", stage1)
            if (_ACCEPTABLE_RESULT_RE.match(stage1_nfc)
                    and len(_MOJIBAKE_HINT_RE.findall(stage1_nfc))
                        < len(_MOJIBAKE_HINT_RE.findall(full))):
                return stage1_nfc

        # Keine Heilung möglich: Original behalten (Whitelist greift später)
        return full

    # Heiler-Schleife mit Konvergenz-Check. Heterogene Cluster wie
    # 'o¦ê+ƒ' (NFD-Mojibake + UTF-8-Mojibake gemischt) heilen schrittweise:
    # erst macht die Whitelist 'o¦ê' → 'ö', danach kann '├ƒ' im nächsten
    # Durchlauf generisch zu 'ß' geheilt werden.
    current = text
    for _ in range(4):
        prev = current

        # Stufe 0: Plus-Substitution für Box-Drawing-ASCII-Substitute
        # (spekulativ - siehe _repair_mojibake unten)
        if plus_substitution:
            current = _pre_substitute_plus(current)

        # Stufe 1: Cluster-für-Cluster heilen (verkraftet Mixed-Strings,
        # weil saubere Zeichen wie ein bereits korrektes 'ü' nicht mit-encodet
        # werden — sie liegen außerhalb der Cluster-Matches).
        current = _MOJIBAKE_CLUSTER_RE.sub(_heal_match, current)

        # Zwischen-NFC, damit Whitelist-Tokens auf zusammengesetzten Buchstaben
        # matchen (NFD-Combining wurde in _try_heal schon normalisiert, aber
        # der Rest des Strings könnte noch NFD sein).
        current = unicodedata.normalize("NFC", current)

        # Stufe 2: explizite Token-Ersetzungen für nicht-generische Reste
        for src_tok, dst_tok in _POST_REPAIR_TOKENS.items():
            if src_tok in current:
                current = current.replace(src_tok, dst_tok)

        if current == prev:
            break  # Konvergenz erreicht

    return current


def _repair_mojibake(text: str) -> str:
    """Mojibake heilen; die Plus-Substitution dabei nur SPEKULATIV anwenden.

    _pre_substitute_plus ersetzt '+' durch '├' und ist laut eigenem Docstring
    ausdruecklich nur eine Vorstufe fuer den Heiler. Sie wurde aber
    unbedingt vorgenommen und nie zurueckgenommen, wenn danach keine Heilung
    zustande kam - das '├' blieb dann endgueltiger Bestandteil des neuen
    Namens und wurde so auf die Platte umbenannt.

    Betroffen sind legitime Namen mit einem Akzentzeichen neben einem Plus:
    _is_mojibake_neighbor haelt JEDES Nicht-ASCII-Zeichen ausser aeoeueAEOEUEss
    fuer einen Mojibake-Indikator, also auch e-Akut, a-Grave, ©, €, ° und µ,
    die in Archivbestaenden massenhaft legitim vorkommen. Nachgestellt:
    'André+Partner.pdf' -> 'André├Partner.pdf', 'Größe 10€+MwSt.pdf' ->
    'Größe 10€├MwSt.pdf'.

    Deshalb: erst mit Substitution heilen. Bleibt danach ein Box-Drawing-
    Zeichen stehen, hat sie nicht geheilt, sondern nur zerstoert - dann zaehlt
    der Durchlauf OHNE Substitution, in dem das '+' unangetastet bleibt.
    Da die Substitution Box-Drawing-Zeichen nur hinzufuegen kann, ist dieser
    zweite Durchlauf nie schlechter.
    """
    if not _MOJIBAKE_HINT_RE.search(text):
        return text

    mit = _repair_mojibake_core(text, plus_substitution=True)
    if "+" not in text or not _BOXDRAW_RE.search(mit):
        return mit
    return _repair_mojibake_core(text, plus_substitution=False)


# Anführungszeichen-Strip am Komponent-Anfang/-Ende. Entfernt typografische
# Doppel- und Einzel-Quotes, aber NUR wenn beide Enden Quotes sind (matched
# wrap, z. B. '"foo"' → 'foo' oder '»foo«' → 'foo'). Ein einzelner Quote an
# nur einem Ende bleibt stehen, sonst entstehen unbalanced-Reste wie
# '„Concorso d'Eleganza Villa d'Este.pdf' (öffnendes „ in der Mitte, das
# schließende " am Stem-Ende wird isoliert entfernt).
# Mittendrin stehende Quotes bleiben grundsaetzlich unangetastet
# (z. B. 'Bar „Foo" Baz.pdf').
# Das gerade '"' (U+0022) wird bereits durch _INVALID_CHARS_RE gefiltert.
_QUOTE_STRIP_CHARS = (
    "\u201c\u201d\u201e\u201f"  # " " „ ‟ Doppel-Quotes
    "\u2018\u2019\u201a\u201b"  # ' ' ‚ ‛ Einzel-Quotes
    "\u00ab\u00bb"              # « » französische Guillemets
    "\u2039\u203a"              # ‹ › einfache Guillemets
)
_QUOTE_STRIP_SET = frozenset(_QUOTE_STRIP_CHARS)


def _clean_component(text: str) -> str:
    text = _repair_mojibake(text)            # vor NFC heilen
    text = unicodedata.normalize("NFC", text)
    text = _INVISIBLE_RE.sub("", text)
    text = _PICTOGRAM_RE.sub("", text)       # Telefon, Schere, Wingdings, ... ohne Ersatz
    text = _CONTROL_RE.sub(" ", text)
    text = _INVALID_CHARS_RE.sub("-", text)
    text = _MULTI_SPACE_RE.sub(" ", text)
    # Symmetrischer Quote-Strip: nur entfernen, wenn BEIDE Enden Quotes sind.
    # Schleife frisst auch verschachteltes Wrapping wie '""foo""' → 'foo'.
    text = text.strip(" ")
    while (len(text) >= 2
           and text[0] in _QUOTE_STRIP_SET
           and text[-1] in _QUOTE_STRIP_SET):
        text = text[1:-1].strip(" ")
    return text


def sanitize_filename(name: str) -> str:
    name = _clean_component(name)
    name = name.strip().rstrip(". ").strip()
    check = name.split(".", 1)[0] if "." in name else name
    if check.upper() in _RESERVED_NAMES:
        name = f"_{name}"
    if not name:
        name = "_bereinigt"
    return name


def sanitize_extension(ext: str) -> str:
    if not ext:
        return ext
    dot = ext[0]
    suffix = ext[1:]
    suffix = _clean_component(suffix)
    suffix = suffix.strip().rstrip(". ").strip()
    if not suffix:
        return ""
    return dot + suffix


def truncate_name(basename: str, is_dir: bool) -> str:
    if len(basename) <= MAX_NAME_LEN:
        return basename
    if is_dir:
        stem, ext = basename, ""
    else:
        stem, ext = os.path.splitext(basename)
    # errors="replace": Windows/NTFS erlauben kaputte UTF-16-Surrogate-
    # Hälften (z.B. \udc00) in Dateinamen. Die landen via os.walk im
    # basename und würden bei .encode("utf-8") sonst einen harten
    # UnicodeEncodeError werfen und die Verarbeitung dieser Datei abbrechen.
    name_hash = hashlib.md5(basename.encode("utf-8", errors="replace")).hexdigest()[:HASH_LEN]
    if len(ext) > MAX_NAME_LEN - HASH_LEN - 1:
        return f"{name_hash}{ext[:MAX_NAME_LEN - HASH_LEN]}"
    available = MAX_NAME_LEN - len(ext) - 1 - HASH_LEN
    return f"{stem[:available]}_{name_hash}{ext}"


# ==================================================================
# Rename mit Retry (AV-Scanner / Indexer / SMB)
# ==================================================================
def _rename_with_retry(src: str, dst: str) -> None:
    # PermissionError ist eine Unterklasse von OSError, das Tuple war
    # redundant. FileExistsError ist ebenfalls OSError (WinError 183) und
    # faellt korrekt in den nicht-transienten Zweig, wo der Aufrufer es
    # gezielt abfaengt.
    last_err: Exception | None = None
    for attempt in range(RENAME_RETRIES):
        try:
            os.rename(src, dst)
            return
        except OSError as e:
            last_err = e
            winerr = getattr(e, "winerror", None)
            if winerr in (32, 33):
                # Sharing-/Lock-Violation: klassischer AV-Scanner-Konflikt,
                # loest sich in der Regel nach kurzem Warten.
                max_tries = RENAME_RETRIES
            elif winerr == 5 or (winerr is None and isinstance(e, PermissionError)):
                # Zugriff verweigert: meist eine fehlende Berechtigung und
                # damit dauerhaft. Nur kurz nachfassen.
                max_tries = DENIED_RETRIES
            else:
                raise
            if attempt >= max_tries - 1:
                raise
            time.sleep(RENAME_BASE_DELAY * (2 ** attempt))
    if last_err:
        raise last_err


# ==================================================================
# Sanitize-Eintrag
# ==================================================================
def sanitize_entry_on_disk(
    entry_path: str,
    pbar: tqdm,
    is_dir: bool,
    truncate_long: bool,
    dry_run: bool = False,
) -> RenameStatus:
    directory = os.path.dirname(entry_path)
    basename  = os.path.basename(entry_path)

    # Versionsverwaltung und Systemordner bleiben unangetastet.
    if is_excluded_path(entry_path):
        return RenameStatus.EXCLUDED

    if is_dir:
        new_name = sanitize_filename(basename)
    else:
        stem, ext  = os.path.splitext(basename)
        clean_stem = sanitize_filename(stem)
        clean_ext  = sanitize_extension(ext)
        new_name   = f"{clean_stem}{clean_ext}"

    if truncate_long:
        new_name = truncate_name(new_name, is_dir)

    if new_name == basename:
        return RenameStatus.CLEAN

    new_path = os.path.join(directory, new_name)

    # Der Gesamtpfad ist das eigentliche Google-Drive-Limit, nicht der
    # einzelne Name. Bisher wurde ausschliesslich der Name auf 255 Zeichen
    # geprueft, waehrend die Kollisionsauflaesung ('_2', UUID-Suffix) und
    # der Hash-Suffix den Pfad sogar verlaengern. Umbenannt wird trotzdem
    # (der neue Name ist ja der bessere), aber der Anwender erfaehrt es.
    if len(_display_path(new_path)) > MAX_PATH_LEN:
        logger.warning(
            f"Zielpfad ueber {MAX_PATH_LEN} Zeichen "
            f"({len(_display_path(new_path))}): {_display_path(new_path)}"
        )

    # Kollisionsauflösung
    if safe_exists(new_path) and new_path.lower() != entry_path.lower():
        resolved = False
        if is_dir:
            base_for_collision = new_name
            ext_for_collision  = ""
        else:
            base_for_collision, ext_for_collision = os.path.splitext(new_name)

        for counter in range(2, 100):
            candidate_name = f"{base_for_collision}_{counter}{ext_for_collision}"
            if truncate_long:
                candidate_name = truncate_name(candidate_name, is_dir)
            candidate = os.path.join(directory, candidate_name)
            if not safe_exists(candidate):
                new_path = candidate
                new_name = os.path.basename(candidate)
                resolved = True
                break

        # UUID-Fallback
        if not resolved:
            short_id = uuid.uuid4().hex[:8]
            fallback_name = f"{base_for_collision}_{short_id}{ext_for_collision}"
            if truncate_long:
                fallback_name = truncate_name(fallback_name, is_dir)
            candidate = os.path.join(directory, fallback_name)
            if not safe_exists(candidate):
                new_path = candidate
                new_name = fallback_name
                resolved = True

        if not resolved:
            logger.warning(f"Alle Zielnamen belegt für {_display_path(entry_path)}")
            pbar.write(f"  ⚠  Übersprungen (Kollision): '{basename}'")
            return RenameStatus.SKIPPED

    entry_type = "Ordner" if is_dir else "Datei"

    if dry_run:
        pbar.write(f"  [PROBELAUF] [{entry_type}] '{basename}' → '{new_name}'")
        logger.info(
            f"PROBELAUF: {_display_path(entry_path)} → {_display_path(new_path)}"
        )
        log_rename("WUERDE_UMBENENNEN", is_dir, directory, basename, new_name)
        return RenameStatus.WOULD_RENAME

    # Umbenennung und Protokollierung strikt trennen: frueher standen
    # log_rename() und logger.info() im selben try wie _rename_with_retry.
    # Ein Dateiname mit halber Surrogatpaar-Haelfte (\udc80) - deren Existenz
    # dieses Skript weiter oben selbst dokumentiert - loeste beim Schreiben
    # der Protokollzeile einen UnicodeEncodeError aus. Der ist eine
    # Unterklasse von ValueError, nicht von OSError, lief also am 'except
    # OSError' in log_rename vorbei und wurde unten von 'except Exception'
    # gefangen. Ergebnis: os.rename war bereits gelaufen, die Datei trug den
    # neuen Namen - gemeldet wurde 'Umbenennung fehlgeschlagen', gezaehlt
    # ERROR statt RENAMED, und die CSV-Zeile fehlte. Damit fehlte fuer genau
    # diese Umbenennung die einzige Ruecknahmequelle, die das Skript dem
    # Anwender zusichert.
    erfolg_status = None
    erfolg_name   = None
    erfolg_pfad   = None
    try:
        _rename_with_retry(_lp(entry_path), _lp(new_path))
        erfolg_status, erfolg_name, erfolg_pfad = "UMBENANNT", new_name, new_path
    except FileExistsError:
        # TOCTOU-Race: Zwischen der safe_exists-Kollisionsprüfung oben und
        # dem os.rename hat ein Dritter (paralleler Drive-Sync, AV-Scanner)
        # den Zielnamen belegt. os.rename wirft dann FileExistsError
        # (WinError 183). Statt das als harten Fehler zu werten, einmal mit
        # einem garantiert eindeutigen UUID-Suffix neu auflösen.
        short_id = uuid.uuid4().hex[:8]
        if is_dir:
            race_name = f"{new_name}_{short_id}"
        else:
            race_stem, race_ext = os.path.splitext(new_name)
            race_name = f"{race_stem}_{short_id}{race_ext}"
        if truncate_long:
            race_name = truncate_name(race_name, is_dir)
        race_path = os.path.join(directory, race_name)
        try:
            _rename_with_retry(_lp(entry_path), _lp(race_path))
            erfolg_status, erfolg_name, erfolg_pfad = "UMBENANNT_RACE", race_name, race_path
        except Exception as e2:
            pbar.write(f"  ✗  Umbenennung fehlgeschlagen: '{basename}' – {e2}")
            logger.warning(
                f"Umbenennung fehlgeschlagen (nach Race-Auflösung): {_display_path(entry_path)} – {e2}"
            )
            return RenameStatus.ERROR
    except Exception as e:
        pbar.write(f"  ✗  Umbenennung fehlgeschlagen: '{basename}' – {e}")
        logger.warning(
            f"Umbenennung fehlgeschlagen: {_display_path(entry_path)} – {e}"
        )
        return RenameStatus.ERROR

    # Ab hier steht der neue Name auf der Platte. Alles Folgende ist reine
    # Protokollierung und darf den Erfolg nicht mehr umdeuten.
    try:
        zusatz = " (Kollision aufgelöst)" if erfolg_status == "UMBENANNT_RACE" else ""
        pbar.write(f"  📝 [{entry_type}] '{basename}' → '{erfolg_name}'{zusatz}")
        logger.info(
            f"Umbenannt{zusatz}: {_display_path(entry_path)} → {_display_path(erfolg_pfad)}"
        )
    except Exception as e_log:
        try:
            logger.warning(f"Umbenennung erfolgt, Meldung fehlgeschlagen: {e_log!r}")
        except Exception:
            pass
    # log_rename faengt seine Fehler selbst ab und ist die Ruecknahmequelle -
    # bewusst ausserhalb des obigen try, damit es auch dann laeuft, wenn
    # schon die Bildschirmausgabe scheitert.
    log_rename(erfolg_status, is_dir, directory, basename, erfolg_name)
    return RenameStatus.RENAMED


# ==================================================================
# Eintrags-Generator (Bottom-Up)
# ==================================================================
def entry_generator(directory: str):
    def _walk_error(err: OSError) -> None:
        logger.warning(
            f"Verzeichnis nicht lesbar (fehlende Rechte?) – übersprungen: "
            f"{_display_path(err.filename or '')}"
        )

    lp_directory = prepare_long_path(directory)

    for root, dirs, files in os.walk(lp_directory, topdown=False, onerror=_walk_error):
        for f in files:
            nl = f.lower()
            if nl in SKIP_FILES or nl.startswith(SKIP_FILE_PREFIXES):
                continue
            yield (os.path.join(root, f), False)

        for d in dirs:
            dl = d.lower()
            if dl.startswith(SKIP_DIR_PREFIXES):
                continue
            yield (os.path.join(root, d), True)


# ==================================================================
# Zwei-Pass-Zählung (speicherfreundlich für Millionen Einträge)
# ==================================================================
def count_entries(directory: str) -> int:
    print("\nIndiziere Dateien und Ordner (Pass 1 — kann einige Minuten dauern) ...")
    spinner = "|/-\\"
    count = 0
    for _ in entry_generator(directory):
        count += 1
        if count % 5000 == 0:
            print(f"\r  {spinner[count // 5000 % 4]}  {count:,} Einträge gefunden ...",
                  end="", flush=True)
    print(f"\r  ✓  {count:,} Einträge gefunden. Starte Bereinigung (Pass 2) ...\n")
    return count


# ==================================================================
# Wiederaufnahme nach Abbruch
# ==================================================================
# Das Skript laeuft ueber Bestaende in Millionenhoehe und fing nach
# einem Abbruch bisher von vorne an - besonders aergerlich, weil der
# optionale Zaehldurchlauf den Baum bereits einmal komplett abgeht.
# Muster uebernommen aus 3a bis 3c (append_resume / load_resume_set).
#
# Vermerkt wird der Pfad VOR der Umbenennung. Damit ueberspringt ein
# zweiter Lauf genau die Eintraege, die schon geprueft wurden - die
# Analyse je Eintrag (Mojibake-Erkennung, Normalisierung, Kollisions-
# pruefung) faellt dann weg.
def get_resume_file_path(directory: str) -> str:
    kennung = hashlib.sha1(os.path.abspath(directory).encode("utf-8",
                                                             "surrogatepass")).hexdigest()[:12]
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or "."
    ordner = os.path.join(basis, "Dateinamen-Bereinigung")
    try:
        os.makedirs(ordner, exist_ok=True)
    except Exception:
        ordner = basis
    return os.path.join(ordner, f"resume_{kennung}.txt")


def load_resume_set(pfad: str) -> set:
    if not pfad or not os.path.exists(pfad):
        return set()
    try:
        with open(pfad, "r", encoding="utf-8", errors="replace") as fh:
            return {z.rstrip("\n") for z in fh if z.strip()}
    except Exception as e:
        logger.warning(f"Resume-Datei nicht lesbar ({e}) - beginne von vorne.")
        return set()


def append_resume(pfad: str, eintrag: str) -> None:
    if not pfad:
        return
    try:
        with open(pfad, "a", encoding="utf-8") as fh:
            fh.write(eintrag + "\n")
    except Exception as _e:
        # Ein nicht schreibbarer Fortschrittsvermerk darf den Lauf nicht
        # anhalten - er kostet dann nur die Wiederaufnahme.
        logger.debug(f"append_resume: Exception verworfen: {_e!r}")


def delete_resume_file(pfad: str) -> None:
    try:
        if pfad and os.path.exists(pfad):
            os.remove(pfad)
    except Exception as _e:
        logger.debug(f"delete_resume_file: Exception verworfen: {_e!r}")


# ==================================================================
# Verzeichnis-Verarbeitung
# ==================================================================
def process_directory(
    directory: str,
    count_files_first: bool,
    truncate_long: bool,
    dry_run: bool = False,
    resume_path: Optional[str] = None,
) -> Counter:
    if count_files_first:
        total = count_entries(directory)
        bar_fmt = "{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]"
    else:
        total = None
        bar_fmt = "{desc}: {n_fmt} Einträge [{elapsed}, {rate_fmt}]"
        print("\nStarte direkte Prüfung ...\n")

    print("-" * 66)

    stats: Counter = Counter({s: 0 for s in RenameStatus})

    # Im Probelauf wird nichts vermerkt: er aendert nichts, also gibt es
    # auch nichts fortzusetzen.
    erledigt = load_resume_set(resume_path) if (resume_path and not dry_run) else set()
    if erledigt:
        print(f"  ♻  Wiederaufnahme: {len(erledigt):,} bereits geprüfte Einträge werden übersprungen\n")

    with tqdm(total=total, desc="Prüfe", unit="Eintrag", bar_format=bar_fmt) as pbar:
        for entry_path, is_dir in entry_generator(directory):
            if erledigt and entry_path in erledigt:
                pbar.update(1)
                continue
            try:
                status = sanitize_entry_on_disk(
                    entry_path, pbar, is_dir, truncate_long, dry_run
                )
                stats[status] += 1
                # NUR dauerhaft erledigte Eintraege vermerken. Frueher lief
                # append_resume ohne jede Statusabfrage: ein Eintrag, dessen
                # Umbenennung an WinError 32 gescheitert war (AV-Scanner,
                # Indexer - genau der Fall, fuer den RENAME_RETRIES ueberhaupt
                # existiert), galt danach als erledigt und wurde in JEDEM
                # Folgelauf uebersprungen. Dasselbe fuer SKIPPED, wo alle
                # Ausweichnamen belegt waren - auch das kann sich aufloesen.
                if resume_path and not dry_run and status in _RESUME_STATUSES:
                    append_resume(resume_path, entry_path)
            except Exception as e:
                stats[RenameStatus.ERROR] += 1
                msg = f"Unerwarteter Fehler bei {_display_path(entry_path)}: {e}"
                pbar.write(f"  ✗  {msg}")
                logger.error(msg)
            pbar.update(1)

    summary = ", ".join(f"{k.name}={v}" for k, v in stats.items())
    logger.info(f"Verarbeitung abgeschlossen: {summary}")
    return stats


# ==================================================================
# Einstiegspunkt
# ==================================================================
if __name__ == "__main__":
    start_time = datetime.now()

    print()
    print("╔══════════════════════════════════════════════════════════════╗")
    print("║     DATEINAMEN-BEREINIGUNG FÜR GOOGLE DRIVE FOR DESKTOP      ║")
    print("╚══════════════════════════════════════════════════════════════╝")
    print(f"  Start: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")

    parser = argparse.ArgumentParser(
        description="Dateinamen-Bereinigung für Google Drive for Desktop",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Beispiele:\n"
            "  Interaktiv:   python 10_dateinamen_bereinigen.py\n"
            "  Probelauf:    python 10_dateinamen_bereinigen.py --dir Q:\\Abteilung --dry-run\n"
            "  Automatisch:  python 10_dateinamen_bereinigen.py --dir \\\\server\\share --auto-start\n"
        ),
    )
    parser.add_argument("--dir", metavar="PFAD",
                        help="Zu verarbeitendes Verzeichnis")
    parser.add_argument("--auto-start", action="store_true",
                        help="Alle Bestätigungsabfragen überspringen und sofort starten")
    parser.add_argument("--count-first", action="store_true",
                        help="Einträge vorab zählen (ETA-Anzeige, aber langsamerer Start)")
    parser.add_argument("--truncate-long", action="store_true",
                        help="Namen > 255 Zeichen kürzen (Google-Drive-Limit)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Probelauf: nur anzeigen, was umbenannt würde – nichts ändern")
    parser.add_argument("--no-resume", action="store_true",
                        help="Fortschrittsvermerk ignorieren und von vorne beginnen")
    args = parser.parse_args()

    setup_logging()

    auto_mode     = args.auto_start
    truncate_long = args.truncate_long
    dry_run       = args.dry_run

    if auto_mode and not args.dir:
        print("\n❌  Fehler: --dir fehlt im Automatikmodus.")
        print("    Verwendung: python 10_... --dir \\\\server\\freigabe --auto-start")
        sys.exit(1)

    if args.dir:
        start_dir = sanitize_path(args.dir)
    else:
        start_dir = ask_directory()

    # safe_isdir statt os.path.isdir(prepare_long_path(...)): Letzteres wirft
    # bei ungueltigen Systemzeichen oder Geraeteformaten und lieferte einen
    # Traceback statt einer Fehlermeldung.
    if not safe_isdir(start_dir):
        print(f"\n❌  Fehler: Verzeichnis existiert nicht oder ist nicht erreichbar: '{start_dir}'")
        sys.exit(1)

    if auto_mode:
        count_first = args.count_first
    else:
        if args.count_first:
            count_first = True
        else:
            count_first = ask_yes_no(
                "\nGesamtzahl der Einträge vorab ermitteln?\n"
                "  (Ja = längerer Start, dafür exakte ETA-Anzeige.\n"
                "   Bei Netzlaufwerken (UNC) und sehr vielen Einträgen\n"
                "   (>100.000) kann die Zählung 10–30 Minuten dauern.\n"
                "   Empfohlen ab mehreren hunderttausend Einträgen NEIN.)"
            )
        if not dry_run:
            # Umbenennungen lassen sich nur ueber das CSV-Protokoll
            # zurueckdrehen. Vor dem ersten Lauf auf einer Freigabe ist ein
            # Probelauf daher der sinnvolle Standard - vorher gab es diese
            # Moeglichkeit gar nicht.
            dry_run = ask_yes_no(
                "\nProbelauf durchführen (nichts wird umbenannt)?\n"
                "  (Empfohlen beim ersten Lauf auf einem Verzeichnis.\n"
                "   Zeigt jede geplante Umbenennung an und schreibt sie in\n"
                "   die CSV-Datei, ändert aber nichts auf der Platte.)",
                default_yes=True,
            )
        if not truncate_long:
            truncate_long = ask_yes_no(
                "\nNamen länger als 255 Zeichen kürzen?\n"
                "  (Google Drive akzeptiert max. 255 Zeichen pro Name.\n"
                "   Gekürzte Namen erhalten einen 8-stelligen Hash-Suffix\n"
                "   für Eindeutigkeit. Das Limit gilt pro Datei/Ordner,\n"
                "   nicht für den Gesamtpfad.)"
            )

    print()
    print("=" * 66)
    print("  KONFIGURATION")
    print("=" * 66)
    print(f"  Verzeichnis:   {start_dir}")
    if dry_run:
        print("  Modus:         PROBELAUF – es wird NICHTS umbenannt")
    else:
        print("  Modus:         SCHARF – Dateien und Ordner werden umbenannt")
    print(f"  Zählung:       {'Vorab zählen (ETA)' if count_first else 'Generator (kein ETA)'}")
    print(f"  Längen-Check:  {'JA (Kürzung > 255 Zeichen mit Hash)' if truncate_long else 'NEIN'}")
    print(f"  Ausgeschlossen: {', '.join(sorted(EXCLUDE_DIR_NAMES))}")
    if auto_mode:
        print("  Steuerung:     AUTOMATISCH (keine Bestätigung)")
    print()
    print("  Bereinigungsregeln:")
    print("    1. Mojibake-Reparatur (cp437/cp850/cp1252 → UTF-8 + Whitelist)")
    print("    2. Unicode-NFC-Normalisierung (macOS-NFD → Windows-NFC)")
    print("    3. Unsichtbare Unicode-Zeichen (ZWSP, BOM, Bidi) entfernen")
    print("    4. Steuerzeichen (Tab, LF, CR, ...) durch Leerzeichen ersetzen")
    print("    5. Unzulässige Zeichen (< > : \" / \\ | ? *) durch '-' ersetzen")
    print("    6. Mehrfach-Leerzeichen zu einem Leerzeichen zusammenfassen")
    print("    7. Nachfolgende Leerzeichen und Punkte entfernen")
    print("    8. Reservierte MS-DOS-Namen (CON, PRN, ...) mit '_' präfixieren")
    print("    9. Leere Namen nach Bereinigung → '_bereinigt'")
    if truncate_long:
        print("   10. Namen > 255 Zeichen kürzen mit deterministischem Hash-Suffix")
    print("=" * 66)

    if not auto_mode and not ask_yes_no("\nJetzt starten?"):
        print("Abgebrochen.")
        sys.exit(0)

    logger.info(
        f"Konfiguration — Verzeichnis: {start_dir} | "
        f"Count-First: {count_first} | Truncate: {truncate_long} | "
        f"Probelauf: {dry_run}"
    )
    open_rename_log()

    print(f"\nVerarbeite: '{start_dir}'")
    print("-" * 66)

    resume_path = None if (args.no_resume or dry_run) else get_resume_file_path(start_dir)

    try:
        stats_result = process_directory(
            start_dir, count_first, truncate_long, dry_run, resume_path
        )
        # Nur ein vollstaendig durchgelaufener Durchgang loescht den
        # Vermerk. Nach einem Abbruch bleibt er liegen - genau dafuer
        # ist er da.
        delete_resume_file(resume_path)
    except KeyboardInterrupt:
        print("\n\n*** ABBRUCH durch Benutzer.")
        close_rename_log()
        print(f"    Bisherige Umbenennungen: {RENAME_LOG_FILE}")
        if resume_path:
            print(f"    Fortschritt vermerkt in: {resume_path}")
            print("    Ein erneuter Start setzt dort fort (--no-resume beginnt neu).")
        sys.exit(1)
    finally:
        close_rename_log()

    duration = datetime.now() - start_time

    print()
    print("=" * 66)
    print("  VERARBEITUNG ABGESCHLOSSEN")
    print("=" * 66)
    print(f"  Ende:  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Dauer: {str(duration).split('.')[0]}")
    print()

    total_sum = sum(stats_result.values())
    print("  STATISTIK:")
    if dry_run:
        print(f"    [PROBELAUF] Würde bereinigt: {stats_result[RenameStatus.WOULD_RENAME]}")
    else:
        print(f"    [📝] Bereinigt:           {stats_result[RenameStatus.RENAMED]}")
    print(f"    [✓] Bereits sauber:       {stats_result[RenameStatus.CLEAN]}")
    if stats_result[RenameStatus.EXCLUDED]:
        print(f"    [–] Ausgeschlossen:       {stats_result[RenameStatus.EXCLUDED]}")
    print(f"    [⚠] Übersprungen:         {stats_result[RenameStatus.SKIPPED]}")
    print(f"    [✗] Fehler:               {stats_result[RenameStatus.ERROR]}")
    print(f"        GESAMT:               {total_sum}")
    logger.info(
        f"VERARBEITUNG ABGESCHLOSSEN | "
        f"Verzeichnis: {start_dir} | "
        f"Probelauf: {dry_run} | "
        f"Dauer: {str(duration).split('.')[0]} | "
        f"Bereinigt: {stats_result[RenameStatus.RENAMED]} | "
        f"Wuerde bereinigt: {stats_result[RenameStatus.WOULD_RENAME]} | "
        f"Sauber: {stats_result[RenameStatus.CLEAN]} | "
        f"Ausgeschlossen: {stats_result[RenameStatus.EXCLUDED]} | "
        f"Übersprungen: {stats_result[RenameStatus.SKIPPED]} | "
        f"Fehler: {stats_result[RenameStatus.ERROR]} | "
        f"Gesamt: {total_sum}"
    )

    if dry_run and stats_result[RenameStatus.WOULD_RENAME]:
        print()
        print("  PROBELAUF – es wurde nichts geändert.")
        print("  Für den scharfen Lauf dasselbe Kommando ohne --dry-run aufrufen.")

    print()
    print(f"  Log-Datei:     {os.path.abspath(LOG_FILE)}")
    if os.path.exists(RENAME_LOG_FILE):
        print(f"  Umbenennungen: {os.path.abspath(RENAME_LOG_FILE)}")
        print("                 (Spalten Verzeichnis/AlterName/NeuerName –")
        print("                  damit lässt sich der Lauf rückgängig machen)")
    print("=" * 66)
    print()

    # isatty() wirft ValueError auf einem geschlossenen Stream - das lief
    # vorher ungefangen und beendete das Skript mit einem Traceback, nachdem
    # die eigentliche Arbeit bereits erledigt war.
    try:
        interactive = sys.stdin is not None and sys.stdin.isatty()
    except (ValueError, OSError):
        interactive = False
    if not auto_mode and interactive:
        try:
            input("Beliebige Taste drücken, um das Fenster zu schließen ...")
        except (EOFError, KeyboardInterrupt):
            pass

    sys.exit(0 if stats_result[RenameStatus.ERROR] == 0 else 1)
