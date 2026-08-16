# ==================================================================
# EXCEL-UPDATER: XLS/XLSX/XLSM AUF NEUESTE APPVERSION
# ==================================================================
# Datum: 13.06.2026
# Unterstützte Office-Versionen: 2019 / 2021 / 2024 / 365  (DE / EN)
# ==================================================================
# ⚠ DESTRUKTIVES VERHALTEN (gewollt):
#   Bei JEDER Datei werden entfernt:
#     • Druckbereiche (PageSetup.PrintArea + _xlnm.print_area)
#     • Drucktitel (PrintTitleRows/Columns + _xlnm.print_titles)
#     • AutoFilter-Bereiche (_xlnm._FilterDatabase)
#     • Konsolidierungsquellen (_xlnm.consolidate_area)
#     • Kriterien-/Extraktionsbereiche
#   Grund: Legacy-Dateien haben oft duplizierte Sheet/Workbook-scope
#   Namen und Excel zeigt beim SaveAs einen modalen Konfliktdialog.
#   Pauschales Entfernen verhindert das Hängen im Stapelbetrieb.
#   → Wer Druckbereiche behalten muss, muss sie nach dem Lauf neu setzen.
# ==================================================================
# ANLEITUNG:
# 1. Python installieren: https://www.python.org/
# 2. In der Konsole (CMD): pip install pywin32 tqdm psutil msoffcrypto-tool
# 3. Falls pywin32 erstmals installiert: python Scripts/pywin32_postinstall.py -install
# 4. In der Konsole (CMD): python 3b_xls_xlsx_auf_neueste_Version_aktualisieren.py
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
#    • C:\Users\<Benutzername>\Documents (temporärer Arbeitsordner des Skripts)
#    • Q:\ oder \\server\dfs (Quelldateien)
#    Jeweils: ☑ Unterordner ... sind ebenfalls vertrauenswürdig
#
# HINWEIS FÜR PyInstaller-Build (Umwandlung in .exe):
#   # Einmalig: PyInstaller installieren
#   pip install pyinstaller
#
#   # Build als ein einzeiliger Befehl (am sichersten in jeder Shell):
#   python -m PyInstaller --onefile --noupx --noconfirm --clean --console --icon="python_icon.ico" --name "3b_xls_xlsx_auf_neueste_Version_aktualisieren" --hidden-import pywintypes --hidden-import pythoncom --hidden-import win32api --hidden-import win32file --hidden-import win32process --hidden-import win32con --hidden-import win32com.client --hidden-import winreg --hidden-import psutil --hidden-import tqdm --collect-submodules msoffcrypto 3b_xls_xlsx_auf_neueste_Version_aktualisieren.py
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
#     --collect-submodules msoffcrypto sammelt die OOXML/CFBF-Backends ein,
#                      die msoffcrypto via Plugin-Mechanismus lädt.
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
# ==================================================================

import argparse
import hashlib
import os
import signal
import smtplib
import sys
import shutil
import tempfile
import logging
import importlib.util
import email.mime.multipart
import email.mime.text
import pythoncom
import pywintypes
import win32com.client
import psutil
from tqdm import tqdm
import time
import uuid
import zipfile
import msoffcrypto
import re
import unicodedata
import winreg
from datetime import datetime
from typing import Optional, Tuple, Set, List
import threading
import traceback
import win32file
import win32process
import win32api
import win32con

# ==================================================================
# UTF-8 Konsole
# ==================================================================
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
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
# COM-Konstanten
# ==================================================================
COM_TRUE  = -1
COM_FALSE =  0

XL_XLSX  = 51
XL_XLSM  = 52
XL_XLTX  = 54
XL_XLTM  = 53
XL_XLAM  = 55
XL_XLSB  = 50

XL_NO_RESTRICTIONS = 0

XL_EXCEL4_MACRO_SHEET = 3

XL_LOCAL_SESSION_CHANGES = 2

XL_CORRUPT_NORMAL  = 0
XL_CORRUPT_REPAIR  = 1
XL_CORRUPT_EXTRACT = 2

XL_LINK_TYPE_EXCEL = 1
XL_LINK_TYPE_OLE   = 2

HRESULT_SAVEAS_MACRO_CONFLICT = {
    -2146827284,
    -2146826259,
}

HRESULT_EXCEL_BUSY_OR_CRASH = (
    -2147418111, -2147023174, -2147023170, -2146823683,
    -2147417848, -2147352567, -2146823146, -2146823136,
)

FILE_READ_ATTRIBUTES  = 0x0080
FILE_WRITE_ATTRIBUTES = 0x0100

WIN_ERROR_SHARING_VIOLATION = 32
WIN_ERROR_LOCK_VIOLATION    = 33

XL_OLD_FORMAT_CODES = {56, 18, 6, 43, 17, 4}

TARGET_APP_VERSION = 16.0
TARGET_LAST_EDITED = 7

OFFICE_VERSIONS_SUPPORTED = "Office 2019 / 2021 / 2024 / 365 (DE/EN)"

# Protokoll-Objekte fruehzeitig binden: _resolve_documents_dir() laeuft schon
# beim Import und 'verwirft' Fehler mit detail_logger.debug(...) - der Name
# wird aber erst am Ende von _setup_logging() belegt. Im Fehlerfall (Documents
# nicht erreichbar, umgeleitetes Profil) loeste der Fehlerschlucker deshalb
# selbst einen NameError aus und das Skript startete gar nicht.
# getLogger liefert dieselben Objekte, die _setup_logging() spaeter mit
# Handlern versieht - die Zuweisung dort bleibt unveraendert gueltig.
file_logger   = logging.getLogger("FileLogger")
detail_logger = logging.getLogger("DetailLogger")


# ==================================================================
# Konfiguration
# ==================================================================
def _resolve_documents_dir() -> str:
    base = os.path.join(os.path.expanduser("~"), "Documents")
    try:
        if os.path.isdir(base):
            return win32api.GetLongPathName(base)
    except Exception as _e:
        detail_logger.debug(f"_resolve_documents_dir: Exception verworfen: {_e!r}")
    return base

_docs_base = _resolve_documents_dir()

TEMP_PROCESS_PATH = os.path.join(_docs_base, f"3b_xls_xlsx_auf_neueste_Version_aktualisieren_{os.getpid()}")
# Kleinbuchstaben-Präfix der eigenen Temp-Ordner, für das gezielte
# Whitelisting-Cleanup in _cleanup_user_temp().
TEMP_PROCESS_PREFIX_LOWER = "3b_xls_xlsx_auf_neueste_version_aktualisieren_"
MAX_PATH_LEN      = 240
LOG_FILE: str           = ""
DETAILED_LOG_FILE: str  = ""
RUN_SUMMARY_FILE: str   = ""
CONVERSIONS_CSV: str    = ""
MAX_RETRIES       = 3
RETRY_DELAY       = 2

EXCEL_RESTART_EVERY = 200

# Open-Timeout fuer Workbooks.Open. BEWUSST KURZ (anders als das
# Word-Pendant 3a): er ist hier tragend fuer die definedName-Pre-Clean-
# Maschinerie - laeuft Open in den modalen Namenskonflikt-Dialog, soll
# der Watchdog schnell zuschlagen, damit der OOXML-Pre-Clean und der
# Re-Open greifen koennen. Ein langer Wert wuerde die Konflikt-Erkennung
# pro Datei um ein Vielfaches verzoegern.
OPEN_TIMEOUT         = 20.0

# Bewusst deutlich groesser als OPEN_TIMEOUT. Der kurze Open-Wert dient
# der schnellen Konflikt-Erkennung; beim Speichern gibt es nichts zu
# erkennen, dafuer aber echte Arbeit: eine grosse Arbeitsmappe ueber eine
# langsame Freigabe zu schreiben darf dauern. Der Waechter soll hier nur
# den echten Haenger abfangen (modaler Dialog, DFS-Ausfall), nicht einen
# langsamen, aber laufenden Schreibvorgang abwuergen.
SAVE_TIMEOUT         = 300.0

AV_WAIT_TIMEOUT      = 15.0
AV_WAIT_POLL         = 0.3
DISPLAY_PATH_MAX_LEN = 80

# Verzeichnisnamen, die nie betreten werden: geloeschte Arbeitsmappen im
# $RECYCLE.BIN wuerden sonst mitkonvertiert (und ihr Original ersetzt);
# ~snapshot/.snapshot sind read-only NAS-Schattenkopien.
EXCLUDE_DIR_NAMES = {"$recycle.bin", "system volume information",
                     "~snapshot", ".snapshot"}

FILE_ATTRIBUTE_REPARSE_POINT = 0x0400

# ==================================================================
# Logging
# ==================================================================
def _resolve_script_directory() -> str:
    try:
        if getattr(sys, "frozen", False):
            return os.path.dirname(os.path.abspath(sys.executable))
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()

def _select_log_directory() -> str:
    candidates = [_resolve_script_directory(), _docs_base, tempfile.gettempdir()]
    for d in candidates:
        try:
            os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, f".write_probe_{os.getpid()}")
            with open(probe, "w", encoding="utf-8") as f:
                f.write("")
            os.remove(probe)
            return d
        except Exception:
            continue
    return os.getcwd()

def _setup_logging() -> tuple:
    global LOG_FILE, DETAILED_LOG_FILE, RUN_SUMMARY_FILE, CONVERSIONS_CSV, _log_dir_global

    log_dir = _select_log_directory()
    _log_dir_global = log_dir
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    LOG_FILE          = os.path.join(log_dir, f"3b_xls_xlsx_auf_neueste_Version_aktualisieren_{ts}.log")
    DETAILED_LOG_FILE = os.path.join(log_dir, f"3b_xls_xlsx_auf_neueste_Version_aktualisieren_detailed_{ts}.log")
    RUN_SUMMARY_FILE  = os.path.join(log_dir, "3b_xls_xlsx_auf_neueste_Version_aktualisieren_last_run.txt")
    CONVERSIONS_CSV   = os.path.join(log_dir, f"3b_xls_xlsx_auf_neueste_Version_aktualisieren_konvertierungen_{ts}.csv")

    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    fh = logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8")
    fh.setFormatter(fmt)
    fl = logging.getLogger("FileLogger")
    fl.setLevel(logging.WARNING)
    fl.addHandler(fh)

    dh = logging.FileHandler(DETAILED_LOG_FILE, mode="w", encoding="utf-8")
    dh.setFormatter(fmt)
    dl = logging.getLogger("DetailLogger")
    dl.setLevel(logging.DEBUG)
    dl.addHandler(dh)

    return fl, dl

_log_dir_global: str = ""
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
    # Nachweisliste der Format-Konvertierungen (.xls/.xlt/.xla -> neu, und
    # Ausweichnamen bei Kollision) mit altem und neuem Pfad fuers Archiv.
    if not CONVERSIONS_CSV:
        return
    try:
        import csv as _csv
        needs_header = not os.path.exists(CONVERSIONS_CSV)
        with open(CONVERSIONS_CSV, "a", encoding="utf-8-sig", newline="") as fh:
            writer = _csv.writer(fh, delimiter=";")
            if needs_header:
                writer.writerow(["timestamp", "alter_pfad", "neuer_pfad"])
            writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                old_path, new_path,
            ])
    except Exception as e:
        detail_logger.warning(f"Konvertierungs-CSV nicht schreibbar: {e}")


def write_run_summary(stats: dict, directory: str, duration,
                      force_update: bool, break_links: bool) -> None:
    if not RUN_SUMMARY_FILE:
        return
    try:
        total = sum(v for k, v in stats.items() if k != "RENAMED")
        with open(RUN_SUMMARY_FILE, "w", encoding="utf-8-sig") as fh:
            fh.write("EXCEL-UPDATER – LETZTER LAUF\n")
            fh.write("=" * 60 + "\n")
            fh.write(f"Verzeichnis:    {directory}\n")
            fh.write(f"Force-Update:   {'Ja' if force_update else 'Nein'}\n")
            fh.write(f"Externe Links:  {'entfernt' if break_links else 'behalten'}\n")
            fh.write(f"Dauer:          {str(duration).split('.')[0]}\n")
            fh.write("\n")
            fh.write(f"Aktualisiert:   {stats.get('UPDATED', 0)}\n")
            fh.write(f"Bereits aktuell:{stats.get('ALREADY_CURRENT', 0)}\n")
            fh.write(f"Übersprungen:   {stats.get('SKIPPED', 0)}\n")
            fh.write(f"Fehler:         {stats.get('ERROR', 0)}\n")
            fh.write(f"Dateinamen ber.:{stats.get('RENAMED', 0)}\n")
            fh.write(f"Gesamt:         {total}\n")
    except Exception as e:
        detail_logger.warning(f"Run-Summary nicht schreibbar: {e}")


# ==================================================================
# E-Mail-Bericht (SMTP) – Parität zum Word-Skript 3a
# ==================================================================
def _build_summary_html(stats: dict, duration, start_dir: str,
                        error_lines: list = None) -> str:
    import html as _html

    def esc(v):
        return _html.escape(str(v))

    def rows(pairs):
        return "".join(
            f'<tr><td style="padding:3px 12px 3px 0;">{esc(l)}</td>'
            f'<td style="padding:3px 0;text-align:right;font-weight:bold;">'
            f'{esc(v)}</td></tr>'
            for l, v in pairs)

    total = sum(v for k, v in stats.items() if k != "RENAMED")
    stat_rows = rows([
        ("Aktualisiert / Konvertiert", stats.get("UPDATED", 0)),
        ("Bereits aktuell",            stats.get("ALREADY_CURRENT", 0)),
        ("Übersprungen",               stats.get("SKIPPED", 0)),
        ("Fehler",                     stats.get("ERROR", 0)),
        ("Dateinamen bereinigt",       stats.get("RENAMED", 0)),
        ("Gesamt",                     total),
    ])
    error_html = ""
    if error_lines:
        items = "".join(
            f'<li style="font-family:monospace;font-size:12px;">'
            f'{esc(l.rstrip())}</li>' for l in error_lines)
        error_html = (
            '<h3 style="margin:16px 0 4px;color:#b00;">'
            'Fehler-Auszug (letzte 50 Einträge)</h3>'
            f'<ul style="margin:0;padding-left:18px;">{items}</ul>')
    return (
        '<html><body style="font-family:Segoe UI,Arial,sans-serif;'
        'color:#222;font-size:14px;">'
        '<h2 style="margin:0 0 8px;">Excel-Updater – Laufbericht</h2>'
        f'<p style="margin:0 0 4px;"><b>Verzeichnis:</b> {esc(start_dir)}</p>'
        f'<p style="margin:0 0 12px;"><b>Laufzeit:</b> '
        f'{esc(str(duration).split(".")[0])}</p>'
        '<h3 style="margin:8px 0 4px;">Statistik</h3>'
        f'<table style="border-collapse:collapse;font-size:14px;">{stat_rows}</table>'
        f'{error_html}</body></html>')


def send_summary_mail(
    to_addr: str,
    from_addr: str,
    smtp_host: str,
    stats: dict,
    duration,
    start_dir: str,
    log_path: str,
    smtp_port: int = 25,
    smtp_user: str = None,
    smtp_pass: str = None,
    smtp_starttls: bool = False,
) -> None:
    total = sum(v for k, v in stats.items() if k != "RENAMED")
    subject = (f"Excel-Updater abgeschlossen – "
               f"{stats.get('UPDATED', 0)} aktualisiert, "
               f"{stats.get('ERROR', 0)} Fehler")

    error_lines = []
    error_excerpt = ""
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

    body = (
        f"Excel-Updater – Laufbericht\n"
        f"{'=' * 60}\n"
        f"Verzeichnis:        {start_dir}\n"
        f"Laufzeit:           {str(duration).split('.')[0]}\n"
        f"\nSTATISTIK:\n"
        f"  Aktualisiert:     {stats.get('UPDATED', 0)}\n"
        f"  Bereits aktuell:  {stats.get('ALREADY_CURRENT', 0)}\n"
        f"  Übersprungen:     {stats.get('SKIPPED', 0)}\n"
        f"  Fehler:           {stats.get('ERROR', 0)}\n"
        f"  Dateinamen ber.:  {stats.get('RENAMED', 0)}\n"
        f"  Gesamt:           {total}\n"
        f"{error_excerpt}"
    )

    try:
        msg = email.mime.multipart.MIMEMultipart("alternative")
        msg["From"]    = from_addr
        msg["To"]      = to_addr
        msg["Subject"] = subject
        msg.attach(email.mime.text.MIMEText(body, "plain", "utf-8"))
        try:
            html_body = _build_summary_html(stats, duration, start_dir,
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
                    detail_logger.debug(
                        f"STARTTLS nicht verfügbar, fahre unverschlüsselt fort: {e_tls}")
            if smtp_user:
                server.login(smtp_user, smtp_pass or "")
            server.sendmail(msg["From"], recipient_list, msg.as_string())
        print(f"  ✓  Zusammenfassung gesendet an: {to_addr}")
    except Exception as e:
        print(f"  ⚠  E-Mail-Versand fehlgeschlagen: {e}")


# ==================================================================
# Globale Excel-Referenz
# ==================================================================
excel_app_global: Optional[win32com.client.CDispatch] = None
excel_app_pid:    Optional[int]                       = None

# ==================================================================
# Forward-Deklarationen für Signal-Handler
# ==================================================================
def _restore_excel_settings(excel) -> None:
    try:
        excel.ScreenUpdating   = COM_TRUE
        excel.EnableEvents     = COM_TRUE
        excel.AskToUpdateLinks = COM_TRUE
    except Exception as _e:
        detail_logger.debug(f"_restore_excel_settings: Exception verworfen: {_e!r}")

def _normalize_username(name: str) -> str:
    return name.casefold() if name else ""

# ==================================================================
# Schutz fremder Excel-Sitzungen
# ==================================================================
# _kill_orphaned_excel() beendete trotz seines Namens JEDEN Excel-Prozess des
# angemeldeten Benutzers - ohne jede Waisen-Pruefung. Hatte der Anwender
# Excel mit ungespeicherter Arbeit offen, waren diese Dokumente beim
# ersten Aufraeumen (auch beim Abbruch mit Strg+C) verloren.
#
# Vor dem Start der eigenen COM-Instanz wird deshalb einmal festgehalten,
# welche Excel-Prozesse es bereits gab. Diese gelten dauerhaft als fremd
# und werden nie beendet.
_FOREIGN_EXCEL_PIDS: set = set()

# Der Schnappschuss allein genuegt nicht: er wird genau einmal vor dem Start
# gefuellt (snapshot_foreign_excel_pids). Jede Excel-Sitzung, die der Anwender
# WAEHREND des Laufs oeffnet, steht nicht darin und galt damit als 'verwaist'
# - _kill_orphaned_excel() beendete sie mit proc.kill(), also ohne
# Speichern-Rueckfrage. Und die Funktion laeuft nicht einmal, sondern bei
# jedem periodischen Neustart (alle EXCEL_RESTART_EVERY Dateien), bei jedem
# Pre-Clean-Neustart, im Signal-Handler und im finally; bei einem Lauf ueber
# eine grosse Ablage ist das faktisch ein Dauerzustand. Die Startwarnung sagt
# dem Anwender ausdruecklich das Gegenteil zu ('Ihre Sitzung wird vom Skript
# NICHT beendet').
#
# Deshalb wird die Frage umgedreht: Statt zu erraten, was fremd ist, wird
# mitgeschrieben, welche Excel-Prozesse dieses Skript SELBST gestartet hat.
# Nur die duerfen beendet werden - alles andere gehoert dem Anwender.
_OWN_EXCEL_PIDS: set = set()


def remember_own_excel_pid(pid) -> None:
    if pid:
        _OWN_EXCEL_PIDS.add(pid)


def snapshot_foreign_excel_pids() -> None:
    """Merkt sich alle Excel-Prozesse, die vor dem Skriptstart liefen."""
    global _FOREIGN_EXCEL_PIDS
    found = set()
    try:
        for p in psutil.process_iter(["pid", "name"]):
            try:
                nm = p.info.get("name") or ""
                if "EXCEL.EXE" in nm.upper():
                    found.add(p.info["pid"])
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
    except Exception as _e:
        # Im Zweifel lieber zu viel schuetzen als eine fremde Sitzung killen.
        detail_logger.debug(f"snapshot_foreign_excel_pids: Exception verworfen: {_e!r}")
    _FOREIGN_EXCEL_PIDS = found
    if found:
        detail_logger.info(
            f"{len(found)} bereits laufende(r) Excel-Prozess(e) erkannt "
            f"- diese werden nicht beendet: {sorted(found)}")


def is_foreign_excel_pid(pid) -> bool:
    return pid in _FOREIGN_EXCEL_PIDS


def _kill_orphaned_excel() -> None:
    try:
        current_user = _normalize_username(psutil.Process().username())
    except Exception:
        detail_logger.warning(
            "_kill_orphaned_excel: Benutzerermittlung fehlgeschlagen – "
            "Prozesse werden nicht beendet (Sicherheitsabbruch)."
        )
        return

    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if not (proc.info["name"] and "EXCEL.EXE" in proc.info["name"].upper()):
                continue
            try:
                if _normalize_username(proc.username()) != current_user:
                    continue
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
            if is_foreign_excel_pid(proc.info["pid"]):
                continue
            # Nur eigene Instanzen beenden. Ein Excel, das dieses Skript nicht
            # selbst gestartet hat, gehoert dem Anwender - auch wenn es erst
            # nach dem Startschnappschuss aufgemacht wurde.
            if proc.info["pid"] not in _OWN_EXCEL_PIDS:
                detail_logger.debug(
                    f"Excel-Prozess PID {proc.info['pid']} nicht vom Skript gestartet "
                    f"- bleibt unangetastet.")
                continue
            try:
                proc.kill()
                proc.wait(timeout=3)
                detail_logger.debug(f"Excel-Prozess beendet: PID {proc.info['pid']}")
            except psutil.AccessDenied:
                detail_logger.warning(f"Kein Zugriff auf PID {proc.info['pid']}")
            except psutil.NoSuchProcess:
                pass
            except psutil.TimeoutExpired:
                detail_logger.warning(f"Prozess reagiert nicht: PID {proc.info['pid']}")
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

# ==================================================================
# Lock-File: parallele Läufe auf demselben Zielverzeichnis verhindern
# ==================================================================
def _lock_file_path(target_dir: str) -> str:
    digest = hashlib.sha1(target_dir.encode("utf-8")).hexdigest()[:16]
    return os.path.join(_docs_base, f"excel_updater_{digest}.lock")

def _own_process_create_time() -> str:
    """Startzeit des eigenen Prozesses als stabile Kennung."""
    try:
        return f"{psutil.Process(os.getpid()).create_time():.6f}"
    except Exception as _e:
        detail_logger.debug(f"_own_process_create_time: Exception verworfen: {_e!r}")
        return ""


def _acquire_lock(target_dir: str) -> Optional[str]:
    """Verhindert parallele Laeufe auf demselben Zielverzeichnis.

    Die Lebendpruefung des Lock-Halters erfolgt ueber PID + Startzeit, NICHT
    ueber den Prozessnamen. Frueher galt ein Prozess nur dann als lebend, wenn
    sein Name 'EXCEL' oder 'PYTHON' enthielt. Im dokumentierten
    Auslieferungsweg (PyInstaller, --name
    '3b_xls_xlsx_auf_neueste_Version_aktualisieren') heisst der Prozess aber
    '3B_XLS_XLSX_AUF_NEUESTE_VERSION_AKTUALISIEREN.EXE' und enthaelt weder das
    eine noch das andere - der Lock des laufenden Nachbarn wurde also als
    'Stale-Lock entfernt' protokolliert und der zweite Lauf startete. Die
    Startzeit schliesst zusaetzlich den Fall aus, dass das Betriebssystem die
    PID inzwischen neu vergeben hat.
    """
    lock_path = _lock_file_path(target_dir)
    if os.path.exists(lock_path):
        try:
            with open(lock_path, "r", encoding="utf-8") as f:
                zeilen = [z.strip() for z in f.readlines()]
            stale_pid = int(zeilen[0]) if zeilen and zeilen[0] else -1
            # Zeile 4 traegt die Startzeit; aeltere Lock-Dateien haben sie
            # nicht - dann zaehlt allein die PID-Existenz (konservativ).
            stale_ctime = zeilen[3] if len(zeilen) > 3 else ""

            if stale_pid == os.getpid():
                # Eigener Lock aus diesem Prozess - kein Fremdlauf.
                return lock_path

            if psutil.pid_exists(stale_pid):
                lebt = True
                if stale_ctime:
                    try:
                        aktuell = f"{psutil.Process(stale_pid).create_time():.6f}"
                        lebt = (aktuell == stale_ctime)
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        # Kein Zugriff heisst nicht 'tot'. Im Zweifel den Lock
                        # respektieren - ein verweigerter Start kostet nichts,
                        # ein Parallellauf kann Dateien beschaedigen.
                        lebt = True
                if lebt:
                    detail_logger.warning(
                        f"Lock wird von PID {stale_pid} gehalten – Lauf wird nicht gestartet.")
                    return None
            # PID tot oder inzwischen neu vergeben – Stale-Lock entfernen.
            os.remove(lock_path)
            detail_logger.warning(f"Stale-Lock entfernt (PID {stale_pid}): {lock_path}")
        except Exception as e:
            detail_logger.warning(f"Lock-Check fehlgeschlagen: {e}")
            return None
    try:
        with open(lock_path, "w", encoding="utf-8") as f:
            f.write(f"{os.getpid()}\n{datetime.now().isoformat()}\n{target_dir}\n"
                    f"{_own_process_create_time()}\n")
        return lock_path
    except Exception as e:
        detail_logger.warning(f"Lock-Datei nicht schreibbar: {e}")
        return None

def _release_lock(lock_path: Optional[str]) -> None:
    # Nur den EIGENEN Lock entfernen. Frueher wurde die Datei bedingungslos
    # geloescht - der zuerst fertige Lauf raeumte damit den Lock eines noch
    # laufenden Nachbarn ab.
    if lock_path and os.path.exists(lock_path):
        try:
            with open(lock_path, "r", encoding="utf-8") as f:
                besitzer = int(f.readline().strip())
            if besitzer != os.getpid():
                detail_logger.warning(
                    f"Lock gehoert PID {besitzer}, nicht diesem Prozess "
                    f"({os.getpid()}) – bleibt bestehen.")
                return
        except Exception as _e:
            detail_logger.debug(f"_release_lock: Besitzpruefung verworfen: {_e!r}")
        try:
            os.remove(lock_path)
        except Exception as _e:
            detail_logger.debug(f"_release_lock: Exception verworfen: {_e!r}")

active_lock_path: Optional[str] = None

# Dateien (z.B. nicht wiederherstellbare Backups), die das Temp-Aufraeumen
# ueberleben muessen - sie sind im Fehlerfall die einzige intakte Kopie
# der Originaldatei.
_preserved_temp_files: set = set()


def _safe_cleanup_temp() -> None:
    """Loescht den Run-spezifischen Temp-Ordner - AUSSER es liegen
    geschuetzte Dateien darin (Datenverlust-Schutz)."""
    if not os.path.exists(TEMP_PROCESS_PATH):
        return
    try:
        preserved = [
            p for p in _preserved_temp_files
            if os.path.normcase(os.path.abspath(p)).startswith(
                os.path.normcase(os.path.abspath(TEMP_PROCESS_PATH)))
            and os.path.exists(p)
        ]
        if preserved:
            detail_logger.warning(
                "GESCHÜTZTE DATEIEN IM TEMP-ORDNER – Ordner wird NICHT gelöscht!\n"
                f"  Pfad: {TEMP_PROCESS_PATH}\n"
                + "\n".join(f"    * {os.path.basename(p)}" for p in preserved)
                + "\n  → Bitte manuell sichern, dann Ordner löschen.")
            print("\n  GESCHÜTZTE DATEIEN vorhanden – Temp-Ordner NICHT gelöscht:")
            print(f"   {TEMP_PROCESS_PATH}")
            for p in preserved:
                print(f"   * {os.path.basename(p)}")
            print("   → Bitte manuell sichern, dann Ordner löschen.")
        else:
            shutil.rmtree(TEMP_PROCESS_PATH)
    except OSError:
        pass
    except Exception as e:
        detail_logger.warning(f"Aufräumen fehlgeschlagen: {e}")

# ==================================================================
# Signal-Handler
# ==================================================================
def _signal_handler(sig, frame) -> None:
    print("\n\n⚠️  Abbruch – räume auf ...")
    if excel_app_global is not None:
        try:
            _restore_excel_settings(excel_app_global)
        except Exception as _e:
            detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")
        try:
            excel_app_global.Quit()
            time.sleep(1)
        except Exception as _e:
            detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")
    _kill_orphaned_excel()
    try:
        _cleanup_excel_inetcache()
    except Exception as _e:
        detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")
    try:
        _cleanup_user_recent()
    except Exception as _e:
        detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")
    _safe_cleanup_temp()
    _release_lock(active_lock_path)
    sys.exit(130)

# ==================================================================
# Excel-PID-Erkennung über psutil-Diff
# ==================================================================
def _snapshot_excel_pids() -> Set[int]:
    pids: Set[int] = set()
    try:
        current_user = _normalize_username(psutil.Process().username())
    except Exception:
        return pids
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if not (proc.info["name"] and "EXCEL.EXE" in proc.info["name"].upper()):
                continue
            try:
                if _normalize_username(proc.username()) == current_user:
                    pids.add(proc.info["pid"])
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return pids

# ==================================================================
# Pre-Run-Excel-Check
# ==================================================================
def _check_excel_running_warning(auto_mode: bool = False, no_kill: bool = False) -> None:
    try:
        current_user = _normalize_username(psutil.Process().username())
    except Exception:
        return

    running_pids = []
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if not (proc.info["name"] and "EXCEL.EXE" in proc.info["name"].upper()):
                continue
            try:
                if _normalize_username(proc.username()) == current_user:
                    running_pids.append(proc.info["pid"])
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    if not running_pids:
        return

    pids_str = ', '.join(str(p) for p in running_pids)

    if no_kill:
        print()
        print("❌  FEHLER: Excel läuft – Abbruch durch --no-kill-excel.")
        print(f"    Gefundene Excel-Prozesse (PID): {pids_str}")
        print("    Bitte alle Excel-Fenster schließen und Skript neu starten.")
        detail_logger.warning(
            f"Abbruch --no-kill-excel: laufende Excel-Prozesse (PID: {pids_str})"
        )
        file_logger.warning(
            f"Abbruch durch --no-kill-excel: laufende Excel-Prozesse (PID: {pids_str})"
        )
        sys.exit(1)

    if auto_mode:
        detail_logger.warning(
            f"Automatikmodus: laufende Excel-Prozesse gefunden (PID: {pids_str}) – "
            f"werden durch _kill_orphaned_excel() beendet."
        )
        return

    # Hinweis zum Text: Frueher stand hier, das Skript schliesse ALLE
    # Excel-Fenster - das traf zu, solange _kill_orphaned_excel() jeden
    # Excel-Prozess des Benutzers beendete. Seit snapshot_foreign_excel_pids()
    # die vorher laufenden Sitzungen schuetzt, stimmt das nicht mehr. Eine
    # Warnung, die mehr androht als eintritt, kostet Vertrauen - und eine,
    # die zu wenig androht, ist gefaehrlich. Daher praezise formuliert.
    print()
    print("=" * 66)
    print("  WARNUNG: Excel laeuft bereits")
    print("=" * 66)
    print(f"  Gefundene Excel-Prozesse (PID): {pids_str}")
    print()
    print("  Ihre Sitzung wird vom Skript NICHT beendet - auch eine, die Sie")
    print("  erst waehrend des Laufs oeffnen, bleibt unangetastet. Waehrend")
    print("  des Laufs kann sie aber ausgeblendet werden und Warnhinweise")
    print("  sind abgeschaltet - das wirkt wie ein Absturz.")
    print("  Ausserdem laesst sich die eigene Automatisierungs-Instanz dann")
    print("  nicht mehr zuverlaessig von Ihrer Sitzung unterscheiden.")
    print()
    print("  EMPFEHLUNG: Excel jetzt schliessen und das Skript neu starten.")
    print("=" * 66)
    print()
    if not ask_yes_no("Trotzdem fortfahren?"):
        print("Abgebrochen. Bitte Excel schliessen und neu starten.")
        sys.exit(0)

# ==================================================================
# Modul-Check
# ==================================================================
def check_required_modules() -> None:
    if getattr(sys, "frozen", False):
        return

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
# Long-Path-Hilfsfunktionen
# ==================================================================
def prepare_long_path(path: str) -> str:
    if path.startswith("\\\\?\\"):
        return path
    path = os.path.abspath(os.path.normpath(path))
    if path.startswith("\\\\"):
        return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path

def _truncate_display(path: str, max_len: int = DISPLAY_PATH_MAX_LEN) -> str:
    if len(path) <= max_len:
        return path
    half = (max_len - 3) // 2
    return path[:half] + "..." + path[-half:]

# ==================================================================
# AV-Scanner / File-Lock-Wartelogik
# ==================================================================
def wait_for_file_unlocked(
    path: str,
    timeout: float = AV_WAIT_TIMEOUT,
    poll_interval: float = AV_WAIT_POLL,
    require_exclusive: bool = False,
) -> bool:
    if not safe_exists(path):
        return False

    p = prepare_long_path(path) if len(path) > MAX_PATH_LEN else path
    share_mode = 0 if require_exclusive else win32file.FILE_SHARE_READ

    end = time.monotonic() + timeout
    last_err = None

    while time.monotonic() < end:
        try:
            h = win32file.CreateFile(
                p,
                win32con.GENERIC_READ,
                share_mode,
                None,
                win32file.OPEN_EXISTING,
                0,
                None,
            )
            try:
                h.Close()
            except Exception as _e:
                detail_logger.debug(f"wait_for_file_unlocked: Exception verworfen: {_e!r}")
            return True
        except pywintypes.error as e:
            last_err = e
            if e.winerror in (WIN_ERROR_SHARING_VIOLATION, WIN_ERROR_LOCK_VIOLATION):
                time.sleep(poll_interval)
                continue
            return False
        except Exception as e:
            last_err = e
            return False

    detail_logger.warning(
        f"wait_for_file_unlocked: Timeout nach {timeout}s ({path}) – letzter Fehler: {last_err}"
    )
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
# Datei-Hilfsfunktionen
# ==================================================================
def safe_remove(path: str) -> bool:
    try:
        p = prepare_long_path(path) if len(path) > MAX_PATH_LEN else path
        if os.path.exists(p):
            try:
                win32api.SetFileAttributes(p, win32con.FILE_ATTRIBUTE_NORMAL)
            except Exception as _e:
                detail_logger.debug(f"safe_remove: Exception verworfen: {_e!r}")
            os.remove(p)
            detail_logger.debug(f"Gelöscht: {path}")
            return True
        return False
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
        p = prepare_long_path(path) if len(path) > MAX_PATH_LEN else path
        return os.path.exists(p)
    except Exception:
        return False

def _lp_for_reserve(p: str) -> str:
    try:
        return prepare_long_path(p)
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
        p = prepare_long_path(path) if len(path) > MAX_PATH_LEN else path
        return os.path.getsize(p)
    except Exception:
        return 0

def verify_file(path: str, min_size: int = 100) -> bool:
    if not safe_exists(path):
        return False
    size = safe_getsize(path)
    if size < min_size:
        detail_logger.warning(f"Datei zu klein ({size} Bytes): {path}")
        return False

    ext = os.path.splitext(path)[1].lower()
    p = prepare_long_path(path) if len(path) > MAX_PATH_LEN else path

    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam"):
        try:
            if not zipfile.is_zipfile(p):
                detail_logger.warning(f"Keine gültige ZIP-Struktur: {path}")
                return False
        except Exception as e:
            detail_logger.warning(f"ZIP-Prüfung fehlgeschlagen: {path} – {e}")
            return False

    elif ext in (".xls", ".xla", ".xlt", ".xlsb"):
        try:
            with open(p, "rb") as f_check:
                header = f_check.read(8)
                if header != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                    detail_logger.warning(f"Kein gültiger OLE2-Header: {path}")
                    return False
        except Exception as e:
            detail_logger.warning(f"OLE2-Header-Prüfung fehlgeschlagen: {path} – {e}")
            return False

    return True

def robust_copy(src: str, dst: str, max_retries: int = MAX_RETRIES) -> bool:
    for attempt in range(max_retries):
        try:
            s = prepare_long_path(src) if len(src) > MAX_PATH_LEN else src
            d = prepare_long_path(dst) if len(dst) > MAX_PATH_LEN else dst
            if os.path.exists(d):
                try:
                    win32api.SetFileAttributes(d, win32con.FILE_ATTRIBUTE_NORMAL)
                except Exception:
                    os.chmod(d, 0o666)
            shutil.copy2(s, d)
            wait_for_file_unlocked(dst, timeout=AV_WAIT_TIMEOUT, require_exclusive=True)
            if verify_file(dst):
                detail_logger.debug(f"Kopiert: {src} → {dst}")
                return True
        except Exception as e:
            detail_logger.warning(f"Kopieren Versuch {attempt+1}/{max_retries}: {e}")
        if attempt < max_retries - 1:
            time.sleep(RETRY_DELAY)
    return False

def robust_move(src: str, dst: str, max_retries: int = MAX_RETRIES) -> bool:
    s = prepare_long_path(src) if len(src) > MAX_PATH_LEN else src
    d = prepare_long_path(dst) if len(dst) > MAX_PATH_LEN else dst

    # Schnellpfad: gleiches Volume -> atomares os.replace ohne Kopie.
    try:
        if os.path.exists(d):
            try:
                win32api.SetFileAttributes(d, win32con.FILE_ATTRIBUTE_NORMAL)
            except Exception:
                os.chmod(d, 0o666)
        os.replace(s, d)
        detail_logger.debug(f"Ersetzt (replace): {src} → {dst}")
        return True
    except OSError as e_rep:
        detail_logger.debug(
            f"os.replace fehlgeschlagen, Fallback staging+replace: {e_rep}")

    # Cross-Volume-Fallback OHNE Truncate-Fenster: src wird zuerst als
    # dst + '.tmp_new' AUF DAS ZIELVOLUME kopiert und dann per os.replace
    # atomar uebergeschoben. shutil.copy2 direkt auf dst wuerde das Ziel
    # sofort trunkieren - ein Abbruch in diesem Fenster (Netzwerk-Drop,
    # AV, Prozess-Kill) hinterliesse eine halbe Datei. Mit Staging ist
    # dst zu jedem Zeitpunkt entweder die alte oder die neue Datei.
    stage      = dst + ".tmp_new"
    stage_long = prepare_long_path(stage) if len(stage) > MAX_PATH_LEN else stage
    for attempt in range(max_retries):
        try:
            shutil.copy2(s, stage_long)
            wait_for_file_unlocked(stage, timeout=AV_WAIT_TIMEOUT, require_exclusive=True)
            if not verify_file(stage):
                raise Exception("Verifizierung der Staging-Kopie fehlgeschlagen")
            if os.path.exists(d):
                try:
                    win32api.SetFileAttributes(d, win32con.FILE_ATTRIBUTE_NORMAL)
                except Exception:
                    os.chmod(d, 0o666)
            os.replace(stage_long, d)
            if not safe_remove(src):
                detail_logger.warning(
                    f"Verschieben: Quelldatei konnte nicht entfernt werden (Duplikat bleibt): {src}"
                )
            detail_logger.debug(f"Verschoben (staging+replace): {src} → {dst}")
            return True
        # BaseException, nicht Exception: Der Signal-Handler beendet sich mit
        # sys.exit() (SystemExit erbt von BaseException). Bei Strg+C mitten im
        # Kopieren blieb die Staging-Datei '<Ziel>.tmp_new' sonst auf der
        # Freigabe liegen, und kein Aufraeumpfad erfasst sie je wieder.
        except BaseException as e:
            if isinstance(e, Exception):
                detail_logger.warning(f"Verschieben Versuch {attempt+1}/{max_retries}: {e}")
            try:
                if os.path.exists(stage_long):
                    safe_remove(stage)
            except Exception as _e:
                detail_logger.debug(f"robust_move: Exception verworfen: {_e!r}")
            # Abbruch nach dem Aufraeumen unveraendert weiterreichen.
            if not isinstance(e, Exception):
                raise
        if attempt < max_retries - 1:
            time.sleep(RETRY_DELAY)
    return False

def create_backup(original_path: str) -> Optional[str]:
    if not safe_exists(original_path):
        return None
    ext = os.path.splitext(original_path)[1]
    backup_path = os.path.join(TEMP_PROCESS_PATH, f"backup_{uuid.uuid4().hex}{ext}")
    if safe_exists(backup_path):
        safe_remove(backup_path)
    if robust_copy(original_path, backup_path):
        detail_logger.debug(f"Backup erstellt: {backup_path} (Original: {original_path})")
        return backup_path
    detail_logger.warning(f"Backup nicht möglich: {original_path}")
    return None

def restore_backup(backup_path: str, target_path: str) -> bool:
    if not safe_exists(backup_path):
        return False
    try:
        if robust_copy(backup_path, target_path):
            detail_logger.info(f"Backup wiederhergestellt: {backup_path} → {target_path}")
            return True
        return False
    except Exception as e:
        detail_logger.error(f"Backup-Wiederherstellung fehlgeschlagen: {e}")
        return False

# ==================================================================
# User-Shell-Folder-Auflösung
# ==================================================================
DOWNLOADS_GUID = "{374DE290-123F-4565-9164-39C4925E467B}"

def _query_user_shell_folder(value_name: str) -> Optional[str]:
    try:
        sub_key = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as k:
            raw, _ = winreg.QueryValueEx(k, value_name)
        try:
            expanded = win32api.ExpandEnvironmentStrings(raw)
        except Exception:
            expanded = os.path.expandvars(raw)
        return expanded
    except Exception:
        return None

def get_desktop_path() -> Optional[str]:
    candidate = _query_user_shell_folder("Desktop")
    if candidate and os.path.isdir(candidate):
        return candidate
    fallback = os.path.join(os.path.expanduser("~"), "Desktop")
    if os.path.isdir(fallback):
        return fallback
    return None

def get_downloads_path() -> Optional[str]:
    candidate = _query_user_shell_folder(DOWNLOADS_GUID)
    if candidate and os.path.isdir(candidate):
        return candidate
    fallback = os.path.join(os.path.expanduser("~"), "Downloads")
    if os.path.isdir(fallback):
        return fallback
    return None

# ==================================================================
# Pfad-Hilfsfunktionen
# ==================================================================
def sanitize_path(raw: str) -> str:
    path = raw.strip().strip('"').strip("'").strip()
    if len(path) == 2 and path[1] == ":":
        path = path + "\\"
    try:
        path = os.path.normpath(path)
    except Exception as _e:
        detail_logger.debug(f"sanitize_path: Exception verworfen: {_e!r}")
    return path

def ask_directory() -> str:
    desktop_path   = get_desktop_path()
    downloads_path = get_downloads_path()

    print("\nZielverzeichnis auswählen:")
    print("  [1] Q:\\")
    print("  [2] R:\\")
    print("  [3] G:\\Geteilte Ablagen")
    print("  [4] G:\\Meine Ablage")
    print("  [5] \\\\server\\dfs")
    print(f"  [6] Desktop                {('('+desktop_path+')')   if desktop_path   else '(nicht gefunden)'}")
    print(f"  [7] Downloads              {('('+downloads_path+')') if downloads_path else '(nicht gefunden)'}")
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
            return "Q:\\"
        elif choice == "2":
            return "R:\\"
        elif choice == "3":
            return "G:\\Geteilte Ablagen"
        elif choice == "4":
            return "G:\\Meine Ablage"
        elif choice == "5":
            return "\\\\server\\dfs"
        elif choice == "6":
            if desktop_path:
                return desktop_path
            print("  ❌ Desktop-Verzeichnis konnte nicht ermittelt werden.")
        elif choice == "7":
            if downloads_path:
                return downloads_path
            print("  ❌ Downloads-Verzeichnis konnte nicht ermittelt werden.")
        elif choice == "8":
            try:
                raw = input("Pfad eingeben: ")
            except EOFError:
                print("\nKeine Eingabe möglich (EOF) – Abbruch.")
                sys.exit(1)
            path = sanitize_path(raw)
            if os.path.isdir(path):
                return path
            print(f"  ❌ Verzeichnis nicht gefunden: '{path}'")
            print("     (Hinweis: Anführungszeichen werden automatisch entfernt)")
        else:
            print("  Bitte 1-8 eingeben.")

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

def ask_passwords() -> list:
    print()
    print("Passwörter zum Öffnen geschützter Excel-Dateien:")
    print("  Jedes Passwort einzeln eingeben und mit Enter bestätigen (max. 3).")
    print("  Kein Komma oder Semikolon – ein Passwort pro Zeile!")
    print("  Leere Eingabe bei Passwort 1 → keine Passwörter hinterlegt.")
    passwords = []
    for slot in range(1, 4):
        try:
            pw = input(f"  Passwort {slot}: ").strip()
        except EOFError:
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

def ask_progress_mode() -> tuple:
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

# ==================================================================
# Logging-Helfer
# ==================================================================
def log_error(file_path: str, exc: Exception) -> None:
    file_logger.error(f"Datei: {file_path}\n  -> {exc}\n")
    tb = traceback.format_exc()
    detail_logger.error(f"Datei: {file_path}\n  -> {exc}\n{tb}")

# ==================================================================
# Allgemeiner Excel-Watchdog
# ==================================================================
def _excel_call_with_watchdog(call_label: str, timeout: float, excel_pid,
                              call_fn, restore_fn=None):
    """Fuehrt einen blockierenden COM-Aufruf mit Zeitwaechter aus.

    Bisher war in diesem Skript ausschliesslich das Oeffnen abgesichert.
    Das Speichern lief ungeschuetzt - und genau dort haengt Excel im
    Alltag: grosse Arbeitsmappe auf einem Netzlaufwerk, modaler Dialog
    hinter unsichtbarem Fenster, DFS-Timeout, Virenscanner. Ohne Waechter
    blockierte der Lauf dort unbegrenzt, und zwar an der Stelle, an der
    bereits eine Temp-Kopie existiert und das Original ersetzt werden
    soll. 3a und 3c sichern alle COM-Aufrufe so ab; diese Fassung ist die
    Excel-Entsprechung von _word_call_with_watchdog aus 3a.
    """
    done_event   = threading.Event()
    timeout_flag = [False]
    # Schloss + Fertig-Kennzeichen gegen ein schmales, aber echtes
    # Zeitfenster: der Rueckgabewert von call_fn() steht fest, BEVOR das
    # finally done_event setzt. Laeuft der Timeout genau dazwischen ab,
    # toetet der Waechter Excel, obwohl der Aufruf gelungen ist - der
    # Aufrufer bekaeme ein Ergebnis und arbeitete danach mit einer toten
    # COM-Instanz weiter. Schloss ALLEIN genuegt nicht (nachgemessen:
    # 43 -> 30 von 300 Faellen); erst die unteilbare Pruefung auf dem
    # Erfolgspfad unten macht Toeten und Erfolg eindeutig (0 von 300).
    state_lock   = threading.Lock()
    completed    = [False]

    # Erstellungszeit SYNCHRON im Main-Thread erfassen, BEVOR der Watchdog-
    # Thread startet - sonst koennte Windows die PID zwischenzeitlich an
    # eine neue Excel-Sitzung des Benutzers vergeben und der Waechter
    # wuerde nach Timeout den falschen Prozess killen.
    expected_ct = None
    if excel_pid:
        try:
            expected_ct = psutil.Process(excel_pid).create_time()
        except Exception:
            expected_ct = None

    def watchdog():
        if not done_event.wait(timeout):
            with state_lock:
                if completed[0]:
                    return          # Aufruf war bereits fertig
                timeout_flag[0] = True
            if excel_pid:
                try:
                    proc = psutil.Process(excel_pid)
                    try:
                        proc_name = proc.name().upper()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        proc_name = ""
                    if proc_name != "EXCEL.EXE":
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
        _ergebnis = call_fn()
        with state_lock:
            if timeout_flag[0]:
                # Der Waechter hat bereits zugeschlagen: die Instanz ist
                # tot, das Ergebnis damit unbrauchbar. Als Timeout melden,
                # statt dem Aufrufer einen Erfolg vorzuspiegeln.
                raise TimeoutError(f"Excel Timeout bei {call_label}")
            completed[0] = True
        return _ergebnis
    except Exception as e:
        done_event.set()
        if timeout_flag[0]:
            raise TimeoutError(f"Excel Timeout bei {call_label}")
        raise e
    finally:
        done_event.set()
        if restore_fn is not None:
            try:
                restore_fn()
            except Exception as _e:
                detail_logger.debug(f"_excel_call_with_watchdog: Exception verworfen: {_e!r}")


# ==================================================================
# Excel-Open mit Watchdog
# ==================================================================
def safe_excel_open(excel_app, file_path, pw, corrupt_load=XL_CORRUPT_NORMAL,
                    timeout=OPEN_TIMEOUT, read_only=True):
    excel_pid = excel_app_pid

    open_event   = threading.Event()
    timeout_flag = [False]
    state_lock   = threading.Lock()
    completed    = [False]

    # Erstellungszeit SYNCHRON im Main-Thread erfassen, BEVOR der Watchdog-
    # Thread startet. Wuerde sie erst im Thread gelesen, koennte die Skript-
    # Excel-Instanz in den Millisekunden zwischen Thread-Spawn und Lesen
    # abstuerzen, Windows die PID an eine neue (vom Benutzer geoeffnete)
    # Excel-Sitzung vergeben - und der Watchdog merkte sich deren create_time
    # als "erwartet" und killte nach Timeout den falschen Prozess.
    expected_ct = None
    if excel_pid:
        try:
            expected_ct = psutil.Process(excel_pid).create_time()
        except Exception:
            expected_ct = None

    def watchdog():
        if not open_event.wait(timeout):
            with state_lock:
                if completed[0]:
                    return
                timeout_flag[0] = True
                if excel_pid:
                    try:
                        proc = psutil.Process(excel_pid)
                        if "excel" not in proc.name().lower():
                            proc = None
                        if proc is not None and expected_ct is not None:
                            try:
                                if abs(proc.create_time() - expected_ct) > 0.001:
                                    proc = None
                            except Exception:
                                proc = None
                        if proc is not None:
                            proc.kill()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
                    except Exception as _e:
                        detail_logger.debug(f"watchdog: Exception verworfen: {_e!r}")

    wd_thread = threading.Thread(target=watchdog, daemon=True)
    wd_thread.start()

    _prev_alerts = None
    try:
        _prev_alerts = excel_app.DisplayAlerts
    except Exception as _e:
        detail_logger.debug(f"safe_excel_open: Exception verworfen: {_e!r}")
    try:
        excel_app.DisplayAlerts = COM_FALSE
    except Exception as _e:
        detail_logger.debug(f"safe_excel_open: Exception verworfen: {_e!r}")

    wb = None
    try:
        wb = excel_app.Workbooks.Open(
            file_path,
            UpdateLinks               = 0,
            ReadOnly                  = COM_TRUE if read_only else COM_FALSE,
            IgnoreReadOnlyRecommended = COM_TRUE,
            Notify                    = COM_FALSE,
            AddToMru                  = COM_FALSE,
            Password                  = pw,
            WriteResPassword          = "",
            CorruptLoad               = corrupt_load,
        )
        with state_lock:
            completed[0] = True
    except Exception:
        with state_lock:
            completed[0] = True
        open_event.set()
        if timeout_flag[0]:
            raise TimeoutError(
                "Excel Timeout (modaler Dialog: Passwort, Namenskonflikt, "
                "Wiederherstellung oder Netzwerkhänger)"
            )
        raise
    finally:
        open_event.set()
        wd_thread.join(timeout=1.0)
        try:
            if _prev_alerts is not None:
                excel_app.DisplayAlerts = _prev_alerts
        except Exception as _e:
            detail_logger.debug(f"safe_excel_open: Exception verworfen: {_e!r}")

    return wb

# ==================================================================
# Trust-Center Smoke-Test
# ==================================================================
def _build_minimal_xlsx(target_path: str) -> None:
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


def test_trust_center_smoke(excel_app, timeout: float = 25.0) -> Tuple[bool, str]:
    os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)

    test_path = os.path.join(
        TEMP_PROCESS_PATH,
        f"trustcheck_{uuid.uuid4().hex}.xlsx"
    )

    try:
        _build_minimal_xlsx(test_path)
    except Exception as e:
        return (False, f"Test-XLSX konnte nicht erstellt werden: {e}")

    wb = None
    try:
        try:
            wb = safe_excel_open(
                excel_app, test_path, "",
                corrupt_load=XL_CORRUPT_NORMAL, timeout=timeout
            )
        except TimeoutError:
            return (False, f"TIMEOUT nach {timeout:.0f}s "
                            "- Trust Center vermutlich nicht konfiguriert "
                            "oder Geschützte Ansicht aktiv")
        except Exception as e:
            return (False, f"Workbooks.Open fehlgeschlagen: {e}")

        if wb is None:
            return (False, "Workbooks.Open lieferte kein Workbook-Objekt")

        return (True, "")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")
        try:
            if os.path.exists(test_path):
                os.remove(test_path)
        except Exception as _e:
            detail_logger.debug(f"test_trust_center_smoke: Exception verworfen: {_e!r}")

# ==================================================================
# OOXML-Versionsprüfung
# ==================================================================
def get_ooxml_version_info(file_path: str) -> dict:
    info = {
        "app_version":   None,
        "last_edited":   None,
        "lowest_edited": None,
        "rup_build":     None,
    }

    ext = os.path.splitext(file_path)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam"):
        return info

    zip_path = prepare_long_path(file_path) if len(file_path) > MAX_PATH_LEN else file_path

    try:
        with zipfile.ZipFile(zip_path, "r") as z:
            names = z.namelist()

            if "docProps/app.xml" in names:
                with z.open("docProps/app.xml") as f:
                    content = f.read(65536).decode("utf-8", errors="ignore")
                    m = re.search(r"<(?:\w+:)?AppVersion>(\d+\.\d+)</(?:\w+:)?AppVersion>", content)
                    if m:
                        info["app_version"] = float(m.group(1))

            if "xl/workbook.xml" in names:
                with z.open("xl/workbook.xml") as f:
                    content = f.read(524288).decode("utf-8", errors="ignore")
                    m = re.search(r"lastEdited\s*=\s*\"(\d+)\"", content)
                    if m:
                        info["last_edited"] = int(m.group(1))
                    m = re.search(r"lowestEdited\s*=\s*\"(\d+)\"", content)
                    if m:
                        info["lowest_edited"] = int(m.group(1))
                    m = re.search(r"rupBuild\s*=\s*\"(\d+)\"", content)
                    if m:
                        info["rup_build"] = int(m.group(1))

    except PermissionError:
        detail_logger.warning(
            f"OOXML-Info nicht lesbar – Datei ist gesperrt: {file_path}"
        )
    except Exception as e:
        detail_logger.warning(f"OOXML-Info-Lesen fehlgeschlagen ({file_path}): {e}")

    return info

def needs_version_update(
    file_path: str,
    wb_format: Optional[int],
    force_update: bool,
    ooxml_info: dict,
) -> Tuple[bool, str]:
    ext = os.path.splitext(file_path)[1].lower()
    app_version  = ooxml_info.get("app_version")
    last_edited  = ooxml_info.get("last_edited")

    if ext in (".xls", ".xla", ".xlt"):
        return True, "Altes Binärformat"

    if ext == ".xlsb":
        if wb_format in XL_OLD_FORMAT_CODES:
            return True, f"Altes FileFormat ({wb_format})"
        if force_update:
            return True, "Force Update (.xlsb Binärformat)"
        return False, "Binärformat (.xlsb) – keine OOXML-Prüfung möglich"

    if wb_format in XL_OLD_FORMAT_CODES:
        return True, f"Altes FileFormat ({wb_format})"

    if app_version is not None and app_version < TARGET_APP_VERSION:
        return True, f"AppVersion {app_version} < {TARGET_APP_VERSION}"

    if last_edited is not None and last_edited < TARGET_LAST_EDITED:
        return True, (
            f"Schema-Migration (lastEdited={last_edited} < "
            f"{TARGET_LAST_EDITED}, AppVersion {app_version})"
        )

    if app_version is None and last_edited is None:
        return True, "Keine OOXML-Versionsinformationen (Drittanbieter/beschädigt)"
    if app_version is None:
        return True, f"Keine AppVersion (lastEdited={last_edited})"
    if last_edited is None:
        return True, f"Kein lastEdited (AppVersion {app_version}, Drittanbieter?)"

    if force_update:
        return True, (
            f"Force Update (AppVersion {app_version}, "
            f"lastEdited={last_edited})"
        )

    return False, (
        f"Aktuell (AppVersion {app_version}, "
        f"lastEdited={last_edited})"
    )

# ==================================================================
# OOXML-Pre-Clean: Namenskonflikte VOR COM-Open entfernen
# ==================================================================

_BUILTIN_CONFLICT_NAMES_OOXML = {
    "_xlnm.print_area",
    "_xlnm.print_titles",
    "_xlnm._filterdatabase",
    "_xlnm.criteria",
    "_xlnm.extract",
    "_xlnm.consolidate_area",
    "print_area",
    "print_titles",
    "_filterdatabase",
    "criteria",
    "extract",
    "consolidate_area",
    "druckbereich",
    "drucktitel",
    "filterdatenbank",
    "kriterien",
    "extrahieren",
    "konsolidierungsbereich",
}


def scan_definedName_conflicts(file_path: str) -> int:
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam"):
        return 0
    zip_path = prepare_long_path(file_path) if len(file_path) > MAX_PATH_LEN else file_path
    try:
        with zipfile.ZipFile(zip_path, "r") as zin:
            if "xl/workbook.xml" not in zin.namelist():
                return 0
            with zin.open("xl/workbook.xml") as f:
                wb_xml = f.read().decode("utf-8", errors="ignore")
    except Exception:
        return 0
    m = re.search(r"<definedNames\b[^>]*>(.*?)</definedNames>", wb_xml, re.DOTALL)
    if not m:
        return 0
    inner = m.group(1)
    entries = re.findall(
            # Selbstschliessende Form ZUERST: <definedName .../> ohne
            # Inhalt kommt in Excel-Dateien regelmaessig vor. Ohne
            # diese Alternative verschmolz das Muster einen solchen
            # Eintrag mit dem NAECHSTEN bis zu dessen </definedName> -
            # der Pre-Clean traf damit den falschen Bereich.
            r"<definedName[^>]*/>|<definedName[^>]*?>.*?</definedName>",
            inner, re.DOTALL)
    count = 0
    for entry in entries:
        name_m = re.search(r'\bname\s*=\s*"([^"]*)"', entry)
        if not name_m:
            continue
        name_lc = name_m.group(1).lower()
        if name_lc in _BUILTIN_CONFLICT_NAMES_OOXML or name_lc.startswith("_xlnm."):
            count += 1
    return count


def pre_clean_definedNames(file_path: str) -> Tuple[bool, int]:
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam"):
        return (False, 0)

    zip_path = prepare_long_path(file_path) if len(file_path) > MAX_PATH_LEN else file_path
    tmp_clean = zip_path + ".cleantmp"

    try:
        with zipfile.ZipFile(zip_path, "r") as zin:
            if "xl/workbook.xml" not in zin.namelist():
                return (False, 0)
            with zin.open("xl/workbook.xml") as f:
                wb_xml = f.read().decode("utf-8", errors="ignore")

        m = re.search(r"<definedNames\b[^>]*>(.*?)</definedNames>", wb_xml, re.DOTALL)
        if not m:
            return (False, 0)

        original_block = m.group(0)
        inner          = m.group(1)

        entries = re.findall(
            # Selbstschliessende Form ZUERST: <definedName .../> ohne
            # Inhalt kommt in Excel-Dateien regelmaessig vor. Ohne
            # diese Alternative verschmolz das Muster einen solchen
            # Eintrag mit dem NAECHSTEN bis zu dessen </definedName> -
            # der Pre-Clean traf damit den falschen Bereich.
            r"<definedName[^>]*/>|<definedName[^>]*?>.*?</definedName>",
            inner, re.DOTALL)
        if not entries:
            return (False, 0)

        kept       = []
        seen_keys  = set()
        removed    = 0

        for entry in entries:
            name_m = re.search(r'\bname\s*=\s*"([^"]*)"', entry)
            if not name_m:
                removed += 1
                continue
            name_lc = name_m.group(1).lower()

            scope_m = re.search(r'\blocalSheetId\s*=\s*"([^"]*)"', entry)
            scope   = scope_m.group(1) if scope_m else "global"

            if name_lc in _BUILTIN_CONFLICT_NAMES_OOXML or name_lc.startswith("_xlnm."):
                removed += 1
                continue

            if "#REF!" in entry or "#BEZUG!" in entry:
                removed += 1
                continue

            key = (name_lc, scope)
            if key in seen_keys:
                removed += 1
                continue
            seen_keys.add(key)

            kept.append(entry)

        if removed == 0:
            return (False, 0)

        if kept:
            new_block = f"<definedNames>{''.join(kept)}</definedNames>"
        else:
            new_block = ""

        new_wb_xml = wb_xml.replace(original_block, new_block, 1)

        with zipfile.ZipFile(zip_path, "r") as zin:
            with zipfile.ZipFile(tmp_clean, "w", zipfile.ZIP_DEFLATED,
                                 allowZip64=True) as zout:
                for item in zin.infolist():
                    if item.filename == "xl/workbook.xml":
                        zout.writestr(item, new_wb_xml.encode("utf-8"))
                    else:
                        with zin.open(item, "r") as src, \
                             zout.open(item, "w", force_zip64=True) as dst:
                            shutil.copyfileobj(src, dst, length=1024 * 1024)

        os.replace(tmp_clean, zip_path)
        wait_for_file_unlocked(file_path, timeout=AV_WAIT_TIMEOUT)

        detail_logger.warning(
            f"Pre-Clean definedNames: {removed} Einträge entfernt aus {file_path}"
        )
        return (True, removed)

    except Exception as e:
        detail_logger.warning(f"Pre-Clean definedNames fehlgeschlagen ({file_path}): {e}")
        try:
            if os.path.exists(tmp_clean):
                os.remove(tmp_clean)
        except Exception as _e:
            detail_logger.debug(f"pre_clean_definedNames: Exception verworfen: {_e!r}")
        return (False, 0)

# ==================================================================
# Dateinamen-Bereinigung
# ==================================================================
_RESERVED_NAMES = frozenset({
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
})

_INVALID_CHARS = re.compile(r'[<>:"/\\|?*]')

_INVISIBLE_RE = re.compile(
    r'[\x00-\x1f\x7f'
    r'\u200b\u200c\u200d\u200e'
    r'\u200f\u2028\u2029'
    r'\u202a-\u202e'
    r'\u2060\u2061\u2062\u2063'
    r'\u2064\u2066-\u2069'
    r'\ufeff\ufffe\uffff'
    r'\u00ad'
    r'\u034f\u061c\u115f\u1160'
    r'\u17b4\u17b5'
    r'\u180e\uffa0]'
)

def sanitize_filename(name: str) -> str:
    stem, ext = os.path.splitext(name)

    stem = _INVISIBLE_RE.sub("", stem)
    stem = "".join(
        ch for ch in stem
        if unicodedata.category(ch) not in ("Cc", "Cf")
    )

    stem = _INVALID_CHARS.sub("-", stem)

    stem = stem.strip()
    stem = stem.rstrip(".")

    if stem.upper() in _RESERVED_NAMES:
        stem = f"_{stem}"

    if not stem:
        stem = "_bereinigt"

    return stem + ext

def _resolve_collision(directory: str, base_name: str, source_path: str = "") -> str:
    if not safe_exists(os.path.join(directory, base_name)):
        return base_name

    stem, ext = os.path.splitext(base_name)

    # Sonderfall zuerst: zeigt der Ausweichname auf die Quelldatei selbst,
    # ist das keine Kollision - dann darf NICHT reserviert werden.
    for counter in range(2, 1000):
        candidate_path = os.path.join(directory, f"{stem}_{counter}{ext}")
        if source_path and candidate_path.lower() == source_path.lower():
            return f"{stem}_{counter}{ext}"
        if not safe_exists(candidate_path):
            break

    # Atomar reservieren (siehe reserve_unique_path).
    reserved = reserve_unique_path(os.path.join(directory, stem), ext, max_tries=999)
    if reserved:
        return os.path.basename(reserved)
    return f"{stem}_{uuid.uuid4().hex[:8]}{ext}"

def sanitize_file_on_disk(
    file_path: str,
    pbar: tqdm,
) -> Tuple[str, bool]:
    directory = os.path.dirname(file_path)
    old_name  = os.path.basename(file_path)
    new_name  = sanitize_filename(old_name)

    if new_name == old_name:
        return file_path, False

    new_path = os.path.join(directory, new_name)

    if safe_exists(new_path) and new_path.lower() != file_path.lower():
        new_name = _resolve_collision(directory, new_name, source_path=file_path)
        new_path = os.path.join(directory, new_name)

    try:
        src = prepare_long_path(file_path) if len(file_path) > MAX_PATH_LEN else file_path
        dst = prepare_long_path(new_path)  if len(new_path)  > MAX_PATH_LEN else new_path
        # os.replace statt os.rename: liegt am Ziel unser eigener 0-Byte-
        # Platzhalter aus reserve_unique_path, wuerde os.rename scheitern.
        os.replace(src, dst)
        pbar.write(f"  📝 Dateiname bereinigt: '{old_name}' → '{new_name}'")
        detail_logger.info(f"Dateiname bereinigt: '{file_path}' → '{new_path}'")
        return new_path, True
    except Exception as e:
        detail_logger.warning(f"Dateiname-Bereinigung fehlgeschlagen: {file_path} – {e}")
        release_unique_path(new_path)
        return file_path, False

# ==================================================================
# Auto-Resume (Neustart-Faehigkeit grosser Laeufe)
# ==================================================================
def get_auto_resume_path(start_dir: str, log_dir: str) -> str:
    key = hashlib.md5(
        os.path.normcase(os.path.abspath(start_dir)).encode("utf-8")
    ).hexdigest()[:12]
    return os.path.join(log_dir, f"3b_xls_xlsx_auf_neueste_Version_aktualisieren_resume_{key}.txt")


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


# Nur DAUERHAFT erledigte Status ins Resume schreiben. SKIPPED ist in 3b
# mehrdeutig (Timeout/Recovery/temporaerer Lock) und wird bewusst beim
# naechsten Lauf erneut versucht.
RESUME_STATUSES = ("UPDATED", "ALREADY_CURRENT")


# ==================================================================
# Datei-Generator
# ==================================================================
def _is_reparse_point(path: str) -> bool:
    # Junctions/Mount-Points erkennen: os.walk folgt zwar keinen echten
    # Symlinks, traversiert unter Windows aber Junctions als normale
    # Ordner -> Zyklus-Gefahr (Endlosschleife).
    try:
        attrs = os.stat(path, follow_symlinks=False).st_file_attributes
        return bool(attrs & FILE_ATTRIBUTE_REPARSE_POINT)
    except (OSError, AttributeError):
        return False


def _temp_file_is_active(path: str) -> bool:
    # Schneller Lock-Test fuer ~$-Owner-/._-Stub-Dateien: laesst sich die
    # Datei NICHT exklusiv oeffnen, ist sie aktiv (jemand hat die zugehoerige
    # Mappe gerade offen) und darf NICHT geloescht werden. Nur verwaiste
    # (entsperrte) Reste werden bereinigt.
    p = prepare_long_path(path) if len(path) > MAX_PATH_LEN else path
    try:
        h = win32file.CreateFile(
            p, win32con.GENERIC_READ, 0, None,
            win32file.OPEN_EXISTING, 0, None)
        try:
            h.Close()
        except Exception as _e:
            detail_logger.debug(f"_temp_file_is_active: Exception verworfen: {_e!r}")
        return False
    except pywintypes.error as e:
        if e.winerror in (WIN_ERROR_SHARING_VIOLATION, WIN_ERROR_LOCK_VIOLATION):
            return True
        return False
    except Exception:
        return False


def _strip_lp(path: str) -> str:
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\"):
        return path[4:]
    return path


def _path_matches_exclude(root_lower: str, exclude_patterns: list) -> bool:
    # Grenzanker-Matching auf Ordnerebene: Das Muster muss eine komplette
    # Pfadkomponente (oder eine zusammenhaengende Komponentenfolge) sein.
    # Ein reines Substring-Matching ('pat in root_lower') wuerde quer ueber
    # Komponentengrenzen treffen (z.B. '--exclude-dir alt' wuerde
    # 'Verwaltung' ausschliessen). Durch das Einrahmen mit os.sep auf beiden
    # Seiten matcht 'alt' nur eine echte Komponente, Mehr-Segment-Muster
    # ('Archiv\\Alt') funktionieren weiter.
    if not exclude_patterns:
        return False
    sep = os.sep
    anchored = sep + root_lower.strip(sep) + sep
    for pat in exclude_patterns:
        p = pat.strip().strip("\\/").replace("/", sep).lower()
        if p and (sep + p + sep) in anchored:
            return True
    return False


def file_generator(directory: str, clean_temp: bool = True,
                   exclude_patterns: list = None,
                   clean_appledouble: bool = False):
    """Excel-Dateien liefern und dabei optional verwaiste Office-Reste raeumen.

    clean_temp=False raeumt NICHTS - zwingend fuer den Probelauf, der
    zusichert, das Dateisystem unangetastet zu lassen.

    clean_appledouble steuert die '._*'-Dateien getrennt und ist bewusst
    standardmaessig aus: Auf NAS-/SMB-Freigaben mit Mac-Clients sind das
    AppleDouble-Container (Finder-Metadaten, Resource-Forks) und keine
    Excel-Reste. Sie wurden frueher ohne Endungsfilter im gesamten Baum
    geloescht, also auch '._Foto.jpg' und '._Bericht.docx' - Dateien, mit
    denen dieses Skript nichts zu tun hat.
    """
    extensions = {
        ".xls", ".xlsx", ".xlsm", ".xlsb",
        ".xla", ".xlam",
        ".xlt", ".xltx", ".xltm",
    }
    exclude_patterns = [p.lower() for p in (exclude_patterns or [])]

    def _walk_error(err: OSError) -> None:
        detail_logger.warning(
            f"Verzeichnis nicht lesbar (fehlende Rechte?) – übersprungen: {err.filename}"
        )

    directory = prepare_long_path(directory)
    # Eigenen Arbeitsordner ausnehmen: liegt das Ziel z.B. auf Documents,
    # wuerde der Generator sonst die eigenen convert_*/backup_*/preclean_*-
    # Dateien einsammeln, waehrend Excel sie gerade schreibt.
    temp_nc = os.path.normcase(os.path.abspath(TEMP_PROCESS_PATH))

    for root, dirs, files in os.walk(directory, onerror=_walk_error):
        # Per --exclude-dir ausgeschlossene Teilpfade gar nicht erst betreten.
        if _path_matches_exclude(_strip_lp(root).lower(), exclude_patterns):
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

        # '._'-Dateien werden weiterhin von den Zieldateien getrennt (sie sind
        # keine verarbeitbaren Mappen), aber nur dann zum Loeschen vorgemerkt,
        # wenn das ausdruecklich verlangt ist.
        temp_files   = [f for f in files
                        if f.startswith("~$") or (clean_appledouble and f.startswith("._"))]
        target_files = [f for f in files if not (f.startswith("~$") or f.startswith("._"))]

        if clean_temp:
            for f in temp_files:
                full_t = os.path.join(root, f)
                try:
                    # Aktive Owner-/Stub-Dateien (zugehoerige Mappe offen)
                    # NICHT loeschen - nur verwaiste Reste bereinigen.
                    if _temp_file_is_active(full_t):
                        continue
                    safe_remove(full_t)
                except Exception as _e:
                    detail_logger.debug(f"file_generator: Exception verworfen: {_e!r}")

        for f in target_files:
            nl = f.lower()
            if any(nl.endswith(ext) for ext in extensions):
                yield _strip_lp(os.path.join(root, f))

# ==================================================================
# Schutz und Verknüpfungen entfernen
# ==================================================================
def remove_protection(
    wb: win32com.client.CDispatch,
    file_path_display: str,
    break_links: bool,
) -> bool:
    changed = False

    try:
        if wb.MultiUserEditing:
            try:
                is_read_only = bool(wb.ReadOnly)
            except Exception:
                is_read_only = True
            if is_read_only:
                detail_logger.debug(
                    f"Legacy-Freigabe (Shared Workbook) erkannt, "
                    f"Mappe ReadOnly - ExclusiveAccess uebersprungen: "
                    f"{file_path_display}"
                )
            else:
                wb.ExclusiveAccess()
                changed = True
                detail_logger.info(f"Legacy-Freigabe (Shared Workbook) aufgehoben: {file_path_display}")
    except Exception as e:
        detail_logger.warning(f"Fehler beim Aufheben der Arbeitsmappen-Freigabe: {e}")

    try:
        if wb.ProtectStructure or wb.ProtectWindows:
            try:
                wb.Unprotect(Password="")
                changed = True
                detail_logger.info(f"Workbook-Schutz entfernt: {file_path_display}")
            except Exception as e:
                detail_logger.debug(f"Workbook-Unprotect fehlgeschlagen: {e}")
    except Exception as e:
        detail_logger.debug(f"ProtectStructure-Check fehlgeschlagen: {e}")

    try:
        for sheet in wb.Worksheets:
            try:
                if (sheet.ProtectContents
                        or sheet.ProtectDrawingObjects
                        or sheet.ProtectScenarios):
                    sheet.Unprotect(Password="")
                    sheet.EnableSelection = XL_NO_RESTRICTIONS
                    changed = True
                    detail_logger.debug(f"Sheet-Schutz entfernt: {sheet.Name}")
            except Exception as e:
                detail_logger.debug(f"Sheet-Unprotect '{sheet.Name}': {e}")
    except Exception as e:
        detail_logger.warning(f"Sheet-Iteration fehlgeschlagen: {e}")

    if break_links:
        for link_type in (XL_LINK_TYPE_EXCEL, XL_LINK_TYPE_OLE):
            try:
                links = wb.LinkSources(link_type)
                if links:
                    for link in links:
                        try:
                            wb.BreakLink(link, link_type)
                            changed = True
                        except Exception as _e:
                            detail_logger.debug(f"remove_protection: Exception verworfen: {_e!r}")
                    detail_logger.warning(
                        f"Externe Verknüpfungen entfernt (Type {link_type}): {file_path_display}"
                    )
            except Exception as _e:
                detail_logger.debug(f"remove_protection: Exception verworfen: {_e!r}")

    return changed

# ==================================================================
# Namen-Konfliktbereinigung (DESTRUKTIV – siehe Header)
# ==================================================================
_SYSTEM_NAMES_EXACT = {
    "_xlnm.print_area",
    "_xlnm.print_titles",
    "_xlnm._filterdatabase",
    "_xlnm.criteria",
    "_xlnm.extract",
    "_xlnm.consolidate_area",
    "print_area",
    "print_titles",
    "_filterdatabase",
    "criteria",
    "extract",
    "consolidate_area",
    "druckbereich",
    "drucktitel",
    "filterdatenbank",
    "kriterien",
    "extrahieren",
    "konsolidierungsbereich",
}

def _extract_local_name(full_name: str) -> str:
    name = full_name
    if "!" in name:
        name = name.split("!")[-1]
    name = name.strip("'\" \t")
    return name.lower()

def resolve_name_conflicts(
    wb: win32com.client.CDispatch,
) -> bool:
    changed = False

    try:
        for ws in wb.Worksheets:
            try:
                ps = ws.PageSetup
                if ps.PrintArea:
                    ps.PrintArea = ""
                    changed = True
                    detail_logger.debug(f"PageSetup.PrintArea gelöscht: {ws.Name}")
                if ps.PrintTitleRows:
                    ps.PrintTitleRows = ""
                    changed = True
                    detail_logger.debug(f"PageSetup.PrintTitleRows gelöscht: {ws.Name}")
                if ps.PrintTitleColumns:
                    ps.PrintTitleColumns = ""
                    changed = True
                    detail_logger.debug(f"PageSetup.PrintTitleColumns gelöscht: {ws.Name}")
            except Exception as e:
                detail_logger.debug(f"PageSetup nicht bereinigbar ({ws.Name}): {e}")
    except Exception as e:
        detail_logger.warning(f"Fehler bei PageSetup-Bereinigung: {e}")

    def _try_delete_name(nm) -> bool:
        should_delete = False
        try:
            local_en = _extract_local_name(nm.Name)
            try:
                local_de = _extract_local_name(nm.NameLocal)
            except Exception:
                local_de = ""

            should_delete = (
                local_en in _SYSTEM_NAMES_EXACT
                or local_de in _SYSTEM_NAMES_EXACT
            )

            if not should_delete:
                should_delete = "_xlnm." in local_en or "_xlnm." in local_de

        except Exception:
            should_delete = True

        if not should_delete:
            try:
                refers = str(nm.RefersTo).lower()
                if "#ref!" in refers or "#bezug!" in refers:
                    should_delete = True
                    detail_logger.debug(
                        f"Korrupter Bezug zur Löschung vorgemerkt: "
                        f"{getattr(nm, 'Name', '<unlesbar>')}"
                    )
            except Exception as _e:
                detail_logger.debug(f"_try_delete_name: Exception verworfen: {_e!r}")

        if should_delete:
            try:
                nm.Visible = COM_TRUE
            except Exception as _e:
                detail_logger.debug(f"_try_delete_name: Exception verworfen: {_e!r}")

            try:
                nm.Delete()
                return True
            except Exception as e_del:
                try:
                    safe_name = f"Z_Konflikt_{str(uuid.uuid4())[:6]}"
                    nm.Name = safe_name
                    detail_logger.debug(
                        f"Delete fehlgeschlagen ({e_del}), "
                        f"umbenannt in: {safe_name}"
                    )
                    return True
                except Exception as e_ren:
                    detail_logger.warning(
                        f"Delete und Umbenennen fehlgeschlagen für "
                        f"'{getattr(nm, 'Name', '<unlesbar>')}': {e_ren}"
                    )
        return False

    for _pass in range(3):
        deleted_this_pass = 0

        try:
            for i in range(wb.Names.Count, 0, -1):
                try:
                    nm = wb.Names(i)
                    if _try_delete_name(nm):
                        deleted_this_pass += 1
                        changed = True
                        detail_logger.debug(
                            f"[Pass {_pass+1}] Systemname Mappenebene gelöscht (Index {i})"
                        )
                except Exception as _e:
                    detail_logger.debug(f"resolve_name_conflicts: Exception verworfen: {_e!r}")
        except Exception as e:
            detail_logger.warning(f"Fehler bei Mappen-Namen (Pass {_pass+1}): {e}")

        try:
            for ws in wb.Sheets:
                try:
                    for i in range(ws.Names.Count, 0, -1):
                        try:
                            nm = ws.Names(i)
                            if _try_delete_name(nm):
                                deleted_this_pass += 1
                                changed = True
                                detail_logger.debug(
                                    f"[Pass {_pass+1}] Systemname Blattebene "
                                    f"'{getattr(ws, 'Name', '?')}' gelöscht (Index {i})"
                                )
                        except Exception as _e:
                            detail_logger.debug(f"resolve_name_conflicts: Exception verworfen: {_e!r}")
                except Exception as e:
                    detail_logger.debug(
                        f"Fehler bei Blatt-Namen "
                        f"'{getattr(ws, 'Name', '?')}' (Pass {_pass+1}): {e}"
                    )
        except Exception as e:
            detail_logger.warning(f"Fehler bei Sheet-Iteration (Pass {_pass+1}): {e}")

        if deleted_this_pass == 0:
            break

    return changed

# ==================================================================
# SaveAs mit Makro-Fallback
# ==================================================================
def _is_macro_format_error(exc: Exception, target_ext: str) -> bool:
    if target_ext not in (".xlsx", ".xltx"):
        return False
    hresult = getattr(exc, "hresult", None)
    if hresult in HRESULT_SAVEAS_MACRO_CONFLICT:
        return True
    e_text = str(exc).lower()
    return any(kw in e_text for kw in ("macro", "makro", "vba", "0x800a03ec"))

def _save_as_workbook(
    wb: win32com.client.CDispatch,
    temp_save_path: str,
    new_format: int,
    new_ext: str,
    is_template: bool,
    open_password: str = "",
) -> Tuple[str, int, str]:
    """Mappe unter neuem Format speichern.

    open_password ist das Kennwort, mit dem die Mappe geoeffnet werden konnte.
    Es MUSS beim SaveAs wieder mitgegeben werden: 'Password=""' ist in
    Excel-COM kein 'unveraendert lassen', sondern 'ohne Kennwort speichern'.
    Nachgestellt mit echtem Excel - eine mit Kennwort erzeugte .xlsx kam nach
    SaveAs(Password="") unverschluesselt heraus und liess sich ohne Kennwort
    oeffnen; mit SaveAs(Password=<Kennwort>) blieb die Verschluesselung
    erhalten. Da die entstandene Datei das Original ersetzt und das Backup bei
    Erfolg geloescht wird, waere der Vertraulichkeitsverlust endgueltig.
    """
    def _resolve_save_method(workbook):
        method = getattr(workbook, "SaveAs", None)
        if callable(method):
            return workbook, method
        detail_logger.warning(
            f"wb.SaveAs nicht aufrufbar (Typ: {type(method).__name__}) – "
            f"versuche neuen pywin32-Dispatch-Wrapper über _oleobj_"
        )
        try:
            import win32com.client.dynamic as _dyn
            fresh_wb = _dyn.Dispatch(workbook._oleobj_)
            fresh_method = getattr(fresh_wb, "SaveAs", None)
            if callable(fresh_method):
                detail_logger.debug("Dispatch-Rebind erfolgreich – SaveAs jetzt callable")
                return fresh_wb, fresh_method
        except Exception as e_rebind:
            detail_logger.warning(f"Dispatch-Rebind fehlgeschlagen: {e_rebind}")
        return workbook, None

    wb, save_method = _resolve_save_method(wb)
    if save_method is None:
        raise RuntimeError(
            "wb.SaveAs ist nicht aufrufbar (auch nach Dispatch-Rebind). "
            "Wahrscheinliche Ursache: COM-Method-Resolution-Konflikt "
            "(z.B. VBA-Funktion mit Namen 'SaveAs' im Workbook)."
        )

    # Der SaveAs laeuft jetzt unter Zeitwaechter (siehe
    # _excel_call_with_watchdog). Vorher war in diesem Skript nur das
    # Oeffnen abgesichert; ein Haenger beim Speichern blockierte den
    # gesamten Lauf unbegrenzt.
    def _do_save(path, fmt):
        return _excel_call_with_watchdog(
            f"SaveAs {os.path.basename(path)}",
            SAVE_TIMEOUT,
            excel_app_pid,
            lambda: save_method(
                path,
                FileFormat           = fmt,
                Password             = open_password,
                WriteResPassword     = "",
                ReadOnlyRecommended  = COM_FALSE,
                AddToMru             = COM_FALSE,
                ConflictResolution   = XL_LOCAL_SESSION_CHANGES,
            ),
        )

    try:
        _do_save(temp_save_path, new_format)
        detail_logger.debug(f"SaveAs → {temp_save_path}")
        return temp_save_path, new_format, new_ext

    except Exception as e_save:
        if not _is_macro_format_error(e_save, new_ext):
            raise

        if os.path.exists(temp_save_path):
            try:
                safe_remove(temp_save_path)
            except Exception as _e:
                detail_logger.debug(f"_save_as_workbook: Exception verworfen: {_e!r}")

        new_ext        = ".xlsm" if not is_template else ".xltm"
        new_format     = XL_XLSM if not is_template else XL_XLTM
        temp_save_path = os.path.splitext(temp_save_path)[0] + new_ext
        detail_logger.warning(
            f"Makro-Fallback ausgelöst: Speichere als {new_ext}"
        )
        wb, save_method = _resolve_save_method(wb)
        if save_method is None:
            raise RuntimeError(
                "wb.SaveAs (Makro-Fallback) ist nicht aufrufbar."
            )
        _do_save(temp_save_path, new_format)
        detail_logger.debug(f"SaveAs (Makro-Fallback) → {temp_save_path}")
        return temp_save_path, new_format, new_ext

# ==================================================================
# Kern-Logik: Konvertierung / Aktualisierung
# ==================================================================
def convert_excel_file(
    file_path: str,
    excel_app: win32com.client.CDispatch,
    pbar: tqdm,
    passwords: list,
    force_update: bool,
    break_links: bool,
    dry_run: bool = False,
) -> str:
    pbar.write(f"Prüfe: {_truncate_display(os.path.basename(file_path))}")
    detail_logger.info(f"=== Starte: {file_path} ===")

    original_path       = file_path
    is_temp_copy        = False
    backup_path         = None
    temp_save_path      = None
    target_path         = None
    wb                  = None
    target_file_created = False
    move_started        = False
    # Reservierter Zielname (0-Byte-Platzhalter) - siehe Freigabe im finally.
    target_was_reserved = False
    ext            = os.path.splitext(file_path)[1].lower()
    is_template    = ext in (".xlt", ".xltx", ".xltm")
    is_addin       = ext in (".xla", ".xlam")
    is_binary      = ext == ".xlsb"

    if len(file_path) > MAX_PATH_LEN:
        is_temp_copy = True
        try:
            temp_name = f"longpath_{uuid.uuid4().hex}{ext}"
            temp_path = os.path.join(TEMP_PROCESS_PATH, temp_name)
            if not robust_copy(file_path, temp_path):
                raise Exception("Kopieren fehlgeschlagen")
            if not verify_file(temp_path):
                raise Exception("Temp-Kopie leer/ungültig")
            # require_exclusive: die Temp-Kopie wird gleich von Excel SCHREIBEND
            # geoeffnet. Ein reiner Read-Share-Check liefert sofort True, obwohl
            # ein AV-Scanner noch einen Write-Deny-Lock haelt (AV erlaubt fast
            # immer Read-Sharing) -> Excel-Open liefe in eine Sharing-Violation.
            wait_for_file_unlocked(temp_path, timeout=AV_WAIT_TIMEOUT,
                                   require_exclusive=True)
            file_path = temp_path
            detail_logger.debug(f"Temp-Kopie: {temp_path}")
        except Exception as e:
            log_error(original_path, Exception(f"Long-Path-Kopie fehlgeschlagen: {e}"))
            return "ERROR"

    pre_read_ooxml_info = get_ooxml_version_info(file_path)

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

    # NTFS-Sicherheitsinfo (Owner/Group/DACL) des Originals sichern -
    # wird nach der Ersetzung auf die neue Datei uebertragen.
    orig_sd = _get_security_descriptor(original_path)

    try:
        is_encrypted         = False
        needs_password_check = False

        if ext in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb"):
            try:
                with open(file_path, "rb") as f_check:
                    office_file = msoffcrypto.OfficeFile(f_check)
                    is_encrypted = office_file.is_encrypted()
            except Exception as e_crypt:
                detail_logger.debug(f"msoffcrypto-Check fehlgeschlagen: {e_crypt}")
                is_encrypted = not zipfile.is_zipfile(file_path)
        elif ext in (".xls", ".xla", ".xlt"):
            needs_password_check = True
            # Klassische BIFF-Verschluesselung (.xls, FilePass-Record) traegt
            # keine OOXML-Marker. msoffcrypto erkennt sie zuverlaessig –
            # damit faellt der teure "blinde" Excel-Open-Versuch mit
            # Passwort-Dialog weg. Manueller Marker-Scan bleibt als Fallback,
            # falls msoffcrypto das Legacy-Format nicht parst.
            crypto_decided = False
            try:
                with open(file_path, "rb") as f_check:
                    office_file = msoffcrypto.OfficeFile(f_check)
                    if office_file.is_encrypted():
                        is_encrypted         = True
                        needs_password_check = False
                    crypto_decided = True
            except Exception as e_crypt:
                detail_logger.debug(f"msoffcrypto-Check (.xls) fehlgeschlagen: {e_crypt}")
            if not crypto_decided:
                try:
                    with open(file_path, "rb") as f_ole:
                        ole_header = f_ole.read(8)
                        if ole_header == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                            content = f_ole.read(4088)
                            if b'EncryptedPackage' in content or b'EncryptionInfo' in content:
                                is_encrypted = True
                                needs_password_check = False
                except Exception as e_ole:
                    detail_logger.debug(f"OLE-Header-Check fehlgeschlagen: {e_ole}")

        if not is_encrypted and not needs_password_check:
            passwords_to_try = [""]
        elif is_encrypted and not passwords:
            pbar.write("  → ÜBERSPRUNGEN: Kennwortgeschützt (kein Passwort hinterlegt).")
            detail_logger.warning(f"Übersprungen (Kennwort): {original_path}")
            file_logger.warning(f"Datei: {original_path}\n  -> Übersprungen (Passwortgeschützt)\n")
            return "SKIPPED"
        elif needs_password_check:
            passwords_to_try = [""] + passwords
        else:
            passwords_to_try = passwords

        _prev_display_alerts = None
        try:
            _prev_display_alerts = excel_app.DisplayAlerts
            excel_app.DisplayAlerts = COM_FALSE
        except Exception as _e:
            detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")
        try:
            excel_app.Interactive = COM_FALSE
        except Exception as _e:
            detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")

        try:
            win32api.SetFileAttributes(prepare_long_path(file_path), win32con.FILE_ATTRIBUTE_NORMAL)
            detail_logger.debug(f"FILE_ATTRIBUTE_READONLY entfernt: {file_path}")
        except Exception as _e:
            detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")

        wait_for_file_unlocked(file_path, timeout=AV_WAIT_TIMEOUT)

        conflict_count = scan_definedName_conflicts(file_path)
        if conflict_count > 0:
            detail_logger.info(
                f"definedName-Konflikt vorab erkannt "
                f"({conflict_count} Eintraege): {original_path}"
            )
            if not is_temp_copy:
                try:
                    temp_name = f"preclean_{uuid.uuid4().hex}{ext}"
                    temp_path = os.path.join(TEMP_PROCESS_PATH, temp_name)
                    if robust_copy(file_path, temp_path) and verify_file(temp_path):
                        # exklusiv: Temp-Kopie wird gleich schreibend geoeffnet
                        wait_for_file_unlocked(temp_path, timeout=AV_WAIT_TIMEOUT,
                                               require_exclusive=True)
                        file_path    = temp_path
                        is_temp_copy = True
                    else:
                        safe_remove(temp_path)
                        detail_logger.warning(
                            "Prophylaktischer Pre-Clean: Temp-Kopie fehlgeschlagen"
                        )
                except Exception as e_pp:
                    detail_logger.warning(
                        f"Prophylaktischer Pre-Clean: Setup fehlgeschlagen: {e_pp}"
                    )
            if is_temp_copy:
                cleaned, removed = pre_clean_definedNames(file_path)
                if cleaned:
                    pbar.write(
                        f"  ⚠  {removed} OOXML-Namenskonflikte vorab entfernt"
                    )

        wb_opened           = False
        timeout_occurred    = False
        last_exception      = None
        pre_clean_attempted = False

        # Kennwort merken, mit dem das Oeffnen gelungen ist. Es muss beim
        # SaveAs wieder mitgegeben werden, sonst schreibt Excel die Mappe
        # unverschluesselt zurueck (siehe _save_as_workbook).
        open_password = ""
        for open_attempt in range(MAX_RETRIES):
            timeout_occurred = False

            for pw in passwords_to_try:
                try:
                    wb = safe_excel_open(excel_app, file_path, pw,
                                         corrupt_load=XL_CORRUPT_NORMAL, timeout=OPEN_TIMEOUT)
                    wb_opened = True
                    open_password = pw
                    detail_logger.debug(f"Geöffnet (Normal, Passwort: {'[LEER]' if pw == '' else '***'})")
                    break
                except TimeoutError:
                    timeout_occurred = True
                    break
                except Exception as e:
                    last_exception = e

            if wb_opened:
                break

            if timeout_occurred and not pre_clean_attempted:
                pre_clean_attempted = True
                detail_logger.warning(
                    f"Open-Timeout – starte OOXML-Pre-Clean: {original_path}"
                )

                if not is_temp_copy:
                    try:
                        temp_name = f"preclean_{uuid.uuid4().hex}{ext}"
                        temp_path = os.path.join(TEMP_PROCESS_PATH, temp_name)
                        if not robust_copy(file_path, temp_path):
                            detail_logger.warning("Pre-Clean: Kopie nach TEMP fehlgeschlagen")
                            break
                        if not verify_file(temp_path):
                            detail_logger.warning("Pre-Clean: Temp-Kopie ungültig")
                            safe_remove(temp_path)
                            break
                        # exklusiv: Temp-Kopie wird gleich schreibend geoeffnet
                        wait_for_file_unlocked(temp_path, timeout=AV_WAIT_TIMEOUT,
                                               require_exclusive=True)
                        file_path    = temp_path
                        is_temp_copy = True
                    except Exception as e_copy:
                        detail_logger.warning(f"Pre-Clean Setup fehlgeschlagen: {e_copy}")
                        break

                cleaned, removed = pre_clean_definedNames(file_path)
                if cleaned:
                    pbar.write(
                        f"  ⚠  {removed} OOXML-Namenskonflikte vorab entfernt – versuche erneut zu öffnen"
                    )
                    detail_logger.warning(
                        "Excel nach Watchdog-Kill verworfen – starte neue Instanz vor Retry."
                    )
                    try:
                        _restore_excel_settings(excel_app)
                    except Exception as _e:
                        detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")
                    _quit_excel_app(excel_app)
                    _kill_orphaned_excel()
                    _cleanup_excel_inetcache()
                    _cleanup_user_recent()
                    time.sleep(1)
                    excel_app = _create_excel_app()
                    try:
                        excel_app.Interactive = COM_FALSE
                    except Exception as _e:
                        detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")
                    _prev_display_alerts = None
                    pbar.write("  ↻  Excel-Instanz neu gestartet – versuche Re-Open ...")
                    continue

                detail_logger.warning(
                    "Pre-Clean ohne Wirkung – Timeout-Ursache vermutlich Passwort/Recovery"
                )
                break

            if open_attempt == 0 and not timeout_occurred:
                for pw in passwords_to_try:
                    try:
                        wb = safe_excel_open(excel_app, file_path, pw,
                                             corrupt_load=XL_CORRUPT_REPAIR, timeout=OPEN_TIMEOUT)
                        wb_opened = True
                        open_password = pw
                        detail_logger.debug(f"Geöffnet (Repair-Modus, Passwort: {'[LEER]' if pw == '' else '***'})")
                        break
                    except TimeoutError:
                        timeout_occurred = True
                        break
                    except Exception as e:
                        last_exception = e

            if wb_opened or timeout_occurred:
                break

            detail_logger.debug(f"Öffnen Versuch {open_attempt+1} fehlgeschlagen.")
            if open_attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)

        if (not wb_opened
                and not timeout_occurred
                and not pre_clean_attempted
                and last_exception is not None
                and ext in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam")):

            err_text = str(last_exception).lower()
            looks_like_name_issue = (
                "name" in err_text or "konflikt" in err_text
                or "definedname" in err_text or "0x800a03ec" in err_text
            )

            if looks_like_name_issue:
                pre_clean_attempted = True
                detail_logger.warning(
                    f"Retry-Erschöpfung mit '{type(last_exception).__name__}' – "
                    f"versuche Hail-Mary Pre-Clean: {original_path}"
                )

                if not is_temp_copy:
                    try:
                        temp_name = f"preclean_{uuid.uuid4().hex}{ext}"
                        temp_path = os.path.join(TEMP_PROCESS_PATH, temp_name)
                        if (robust_copy(file_path, temp_path)
                                and verify_file(temp_path)):
                            # exklusiv: Temp-Kopie wird gleich schreibend geoeffnet
                            wait_for_file_unlocked(temp_path, timeout=AV_WAIT_TIMEOUT,
                                                   require_exclusive=True)
                            file_path    = temp_path
                            is_temp_copy = True
                        else:
                            safe_remove(temp_path)
                    except Exception as e_copy:
                        detail_logger.warning(f"Hail-Mary Setup fehlgeschlagen: {e_copy}")

                if is_temp_copy:
                    cleaned, removed = pre_clean_definedNames(file_path)
                    if cleaned:
                        pbar.write(
                            f"  ⚠  {removed} OOXML-Namenskonflikte vorab entfernt (Hail-Mary)"
                        )
                        try:
                            _restore_excel_settings(excel_app)
                        except Exception as _e:
                            detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")
                        _quit_excel_app(excel_app)
                        _kill_orphaned_excel()
                        _cleanup_excel_inetcache()
                        _cleanup_user_recent()
                        time.sleep(1)
                        excel_app = _create_excel_app()
                        try:
                            excel_app.Interactive = COM_FALSE
                        except Exception as _e:
                            detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")
                        _prev_display_alerts = None

                        for pw in passwords_to_try:
                            try:
                                wb = safe_excel_open(excel_app, file_path, pw,
                                                     corrupt_load=XL_CORRUPT_NORMAL,
                                                     timeout=OPEN_TIMEOUT)
                                wb_opened = True
                                open_password = pw
                                pbar.write("  ✓  Re-Open nach Hail-Mary erfolgreich.")
                                break
                            except TimeoutError:
                                timeout_occurred = True
                                break
                            except Exception as e_retry:
                                last_exception = e_retry

        try:
            excel_app.Interactive = COM_TRUE
        except Exception as _e:
            detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")
        try:
            if _prev_display_alerts is not None:
                excel_app.DisplayAlerts = _prev_display_alerts
        except Exception as _e:
            detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")

        if timeout_occurred:
            if pre_clean_attempted:
                pbar.write("  → ÜBERSPRUNGEN: Modaler Dialog blockiert (Timeout) – Pre-Clean ohne Wirkung.")
                detail_logger.warning(f"Übersprungen (Timeout nach Pre-Clean): {original_path}")
                file_logger.warning(
                    f"Datei: {original_path}\n"
                    f"  -> Übersprungen (Timeout nach Pre-Clean – Passwort/Recovery-Dialog?)\n"
                )
            else:
                pbar.write("  → ÜBERSPRUNGEN: Modaler Dialog blockiert (Timeout).")
                detail_logger.warning(f"Übersprungen (Timeout): {original_path}")
                file_logger.warning(
                    f"Datei: {original_path}\n"
                    f"  -> Übersprungen (Modal-Dialog: Passwort, Namenskonflikt oder Recovery)\n"
                )
            return "SKIPPED"

        if not wb_opened or wb is None:
            if needs_password_check and passwords and last_exception:
                pbar.write("  → ÜBERSPRUNGEN: Passwort erforderlich (keines der hinterlegten Passwörter passt).")
            else:
                pbar.write("  → ÜBERSPRUNGEN: Zugriff verweigert.")
            detail_logger.warning(f"Übersprungen (kein Zugriff): {original_path}")
            file_logger.warning(f"Datei: {original_path}\n  -> Übersprungen (Passwort/kein Zugriff)\n")
            return "SKIPPED"

        try:
            wb.Activate()
        except Exception as _e:
            detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")

        original_format = None
        try:
            original_format = wb.FileFormat
            detail_logger.debug(f"FileFormat: {original_format}")
        except Exception as _e:
            detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")

        try:
            _need_writable_reopen = bool(wb.MultiUserEditing) and bool(wb.ReadOnly)
        except Exception:
            _need_writable_reopen = False

        if _need_writable_reopen and dry_run:
            # Im Probelauf wird das Original nicht angefasst: kein Backup,
            # kein Reopen mit Schreibzugriff, kein ExclusiveAccess.
            pbar.write("  🔎 [PROBELAUF] Freigegebene Mappe – würde exklusiv geöffnet")
            detail_logger.info(f"[DRY-RUN] WOULD_TAKE_EXCLUSIVE: {original_path}")
            _need_writable_reopen = False

        if _need_writable_reopen:
            if backup_path is None:
                backup_path = create_backup(original_path)
                if backup_path is None:
                    log_error(
                        original_path,
                        Exception("Backup vor ExclusiveAccess fehlgeschlagen – "
                                  "Datei wird zum Schutz übersprungen.")
                    )
                    try:
                        wb.Close(SaveChanges=COM_FALSE)
                    except Exception as _e:
                        detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")
                    return "SKIPPED"
            try:
                _reopen_path = wb.FullName
            except Exception:
                _reopen_path = file_path
            try:
                wb.Close(SaveChanges=COM_FALSE)
            except Exception as _e_close:
                detail_logger.warning(
                    f"Close vor Shared-Reopen fehlgeschlagen: {_e_close}"
                )
            try:
                wb = safe_excel_open(
                    excel_app, _reopen_path, pw,
                    corrupt_load=XL_CORRUPT_NORMAL, timeout=OPEN_TIMEOUT,
                    read_only=False,
                )
                detail_logger.info(
                    f"Shared-Workbook mit Schreibzugriff erneut geoeffnet "
                    f"fuer ExclusiveAccess: {original_path}"
                )
            except Exception as _e_rw:
                detail_logger.warning(
                    f"Shared-Reopen (ReadOnly=False) fehlgeschlagen: {_e_rw} - "
                    f"Fallback ReadOnly-Reopen"
                )
                try:
                    wb = safe_excel_open(
                        excel_app, _reopen_path, pw,
                        corrupt_load=XL_CORRUPT_NORMAL, timeout=OPEN_TIMEOUT,
                        read_only=True,
                    )
                except Exception as _e_ro:
                    log_error(
                        original_path,
                        Exception(f"Reopen komplett fehlgeschlagen "
                                  f"(writable + readonly): {_e_ro}")
                    )
                    return "ERROR"

        # DATENVERLUST-SCHUTZ vor ExclusiveAccess (shared + SCHREIBBAR, direkt
        # aus dem Original geoeffnet): remove_protection() ruft fuer freige-
        # gebene Mappen wb.ExclusiveAccess() auf - und dieser COM-Befehl
        # SPEICHERT die Mappe sofort in-place, veraendert also das ORIGINAL.
        # Der bestehende Backup-Pfad oben deckt nur den shared+ReadOnly-Fall
        # ab. Fuer shared+writable-aus-Original daher hier ein Backup anlegen
        # und das Original als "mutiert" markieren (-> Rollback im except).
        # Wurde die Mappe aus einer Temp-Kopie geoeffnet (is_temp_copy),
        # trifft ExclusiveAccess nur die Temp-Kopie - dann kein Backup noetig.
        original_mutated = bool(_need_writable_reopen)
        try:
            _shared_writable = bool(wb.MultiUserEditing) and not bool(wb.ReadOnly)
        except Exception:
            _shared_writable = False
        if _shared_writable and not is_temp_copy and dry_run:
            pbar.write("  🔎 [PROBELAUF] Freigegebene Mappe – würde exklusiv geöffnet")
            detail_logger.info(f"[DRY-RUN] WOULD_TAKE_EXCLUSIVE: {original_path}")
            _shared_writable = False

        if _shared_writable and not is_temp_copy:
            if backup_path is None:
                backup_path = create_backup(original_path)
                if backup_path is None:
                    log_error(
                        original_path,
                        Exception("Backup vor ExclusiveAccess fehlgeschlagen – "
                                  "Datei wird zum Schutz übersprungen.")
                    )
                    try:
                        wb.Close(SaveChanges=COM_FALSE)
                    except Exception as _e:
                        detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")
                    return "SKIPPED"
            original_mutated = True

        protection_removed = remove_protection(wb, original_path, break_links)

        names_resolved = resolve_name_conflicts(wb)

        update_needed, reason = needs_version_update(
            original_path, original_format, force_update, pre_read_ooxml_info)

        if not update_needed and not protection_removed and not names_resolved:
            pbar.write(f"  ✓  Aktuell ({reason})")
            try:
                wb.Close(SaveChanges=COM_FALSE)
            except Exception as e_close:
                detail_logger.warning(
                    f"wb.Close fehlgeschlagen (ALREADY_CURRENT) – "
                    f"finally übernimmt: {e_close}"
                )
            wb = None
            return "ALREADY_CURRENT"

        # --- DRY-RUN: hier ist Schluss ---------------------------------
        # An diesem Punkt steht fest, dass die Mappe aktualisiert bzw.
        # konvertiert WUERDE (sonst waere oben ALREADY_CURRENT zurueck-
        # gegeben worden). Im Probelauf wird nichts gespeichert, verschoben
        # oder geloescht - die Mappe wird nur geschlossen. Aenderungen aus
        # remove_protection()/resolve_name_conflicts() bestehen nur im
        # Arbeitsspeicher und gehen mit SaveChanges=False verloren.
        if dry_run:
            grund = reason if update_needed else (
                "Schutz würde entfernt" if protection_removed
                else "Namenskonflikte würden bereinigt")
            pbar.write(f"  🔎 [PROBELAUF] Würde aktualisiert/konvertiert ({grund})")
            detail_logger.info(f"[DRY-RUN] WOULD_UPDATE: {original_path} – {grund}")
            try:
                wb.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")
            wb = None
            return "WOULD_UPDATE"

        if update_needed:
            pbar.write(f"  → {reason}. Aktualisiere ...")
        elif protection_removed:
            pbar.write(f"  → Schutz entfernt. Speichere neu ... ({reason})")
        elif names_resolved:
            pbar.write(f"  → Namenskonflikte bereinigt. Speichere neu ... ({reason})")

        has_macros = True
        vbproject_check_failed = False
        try:
            has_macros = bool(wb.HasVBProject)
        except Exception as e_vba:
            vbproject_check_failed = True
            detail_logger.warning(
                f"HasVBProject nicht lesbar – defaultet auf True (Makro-Schutz): {e_vba}"
            )
            pbar.write(
                "  ⚠  HasVBProject-Zugriff verweigert – Datei wird vorsorglich als .xlsm gespeichert."
            )
            pbar.write(
                "      (Trust Center: 'Zugriff auf VBA-Projektobjektmodell vertrauen' aktivieren)"
            )
        if not has_macros:
            try:
                for sheet in wb.Sheets:
                    if sheet.Type == XL_EXCEL4_MACRO_SHEET:
                        has_macros = True
                        detail_logger.info(
                            f"Excel 4.0 Makroblatt erkannt: "
                            f"'{sheet.Name}' in {original_path}"
                        )
                        break
            except Exception as _e:
                detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")

        if is_addin:
            new_format = XL_XLAM
            new_ext    = ".xlam"
        elif is_binary:
            new_format = XL_XLSB
            new_ext    = ".xlsb"
        elif is_template:
            new_format = XL_XLTM if has_macros else XL_XLTX
            new_ext    = ".xltm" if has_macros else ".xltx"
        else:
            new_format = XL_XLSM if has_macros else XL_XLSX
            new_ext    = ".xlsm" if has_macros else ".xlsx"

        temp_save_name = f"convert_{uuid.uuid4().hex}{new_ext}"
        temp_save_path = os.path.join(TEMP_PROCESS_PATH, temp_save_name)

        temp_save_path, new_format, new_ext = _save_as_workbook(
            wb, temp_save_path, new_format, new_ext, is_template,
            open_password=open_password,
        )

        try:
            wb.Close(SaveChanges=COM_FALSE)
        except Exception as e_close:
            detail_logger.warning(
                f"wb.Close fehlgeschlagen nach SaveAs – "
                f"finally übernimmt: {e_close}"
            )
        wb = None
        time.sleep(0.3)

        # exklusiv: die soeben per SaveAs geschriebene Datei wird gleich
        # per robust_move verschoben (gelesen UND die Quelle geloescht);
        # exklusiver Check stellt sicher, dass der AV-Scanner sie freigegeben hat.
        wait_for_file_unlocked(temp_save_path, timeout=AV_WAIT_TIMEOUT,
                               require_exclusive=True)

        if not verify_file(temp_save_path):
            raise Exception("Verifizierung der konvertierten Datei fehlgeschlagen")

        if ext != new_ext:
            target_path = os.path.splitext(original_path)[0] + new_ext
        else:
            target_path = original_path

        if ext != new_ext and safe_exists(target_path):
            target_dir  = os.path.dirname(target_path)
            target_base = os.path.basename(target_path)
            free_name   = _resolve_collision(target_dir, target_base)
            target_path = os.path.join(target_dir, free_name)
            target_was_reserved = True
            pbar.write(
                f"  ⚠  Ziel existiert bereits – speichere als "
                f"'{os.path.basename(target_path)}'"
            )
            detail_logger.info(
                f"Formatkonflikt gelöst: '{original_path}' → '{target_path}'"
            )

        if backup_path is None:
            backup_path = create_backup(original_path)
            # Ein frisch reservierter Platzhalter ist 0 Byte gross - der
            # Schutzabbruch unten gilt nur fuer echte Zieldateien.
            if backup_path is None and safe_exists(target_path) and not target_was_reserved:
                # Ohne Backup darf eine existierende Zieldatei nicht
                # ueberschrieben werden: robust_move trunkiert das Ziel
                # sofort via copy2 - bei einem Abbruch mitten im
                # Kopiervorgang waere es ohne Backup unwiederbringlich
                # zerstoert.
                raise Exception(
                    "Backup fehlgeschlagen – Abbruch zum Schutz der Zieldatei.")

        if is_temp_copy and safe_exists(file_path):
            safe_remove(file_path)
            is_temp_copy = False

        move_started = True
        if not robust_move(temp_save_path, target_path):
            raise Exception("Verschieben der konvertierten Datei fehlgeschlagen")
        target_file_created = True

        # ACL/Owner des Originals auf die neue Datei uebertragen
        # (Admin-Kontext: vollstaendig; Nutzer-Kontext: DACL).
        _apply_security_descriptor(target_path, orig_sd)

        if original_path.lower() != target_path.lower():
            if not safe_remove(original_path):
                if safe_exists(target_path):
                    safe_remove(target_path)
                    log_error(
                        original_path,
                        Exception(
                            f"Originaldatei konnte nicht gelöscht werden (fehlende Rechte?) – "
                            f"Zieldatei '{os.path.basename(target_path)}' wurde zur Vermeidung "
                            f"von Dateiduplizierung ebenfalls entfernt."
                        )
                    )
                raise Exception(
                    f"Konvertierung erfolgreich, aber Originaldatei konnte nicht gelöscht "
                    f"werden (fehlende Rechte?): '{original_path}'"
                )

        if backup_path and safe_exists(backup_path):
            safe_remove(backup_path)
            backup_path = None

        if orig_times and safe_exists(target_path):
            for av_retry in range(5):
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
                    if av_retry == 4:
                        if not _utime_rueckfall(target_path, orig_times):
                            detail_logger.warning(f"Zeitstempel nicht wiederherstellbar: {e_utime}")
                        else:
                            detail_logger.info(
                                "Zeitstempel ueber os.utime-Rueckfall gesetzt "
                                "(ohne Erstellungszeit).")
                    else:
                        time.sleep(0.5)

        # Nachweisliste: Pfadwechsel (z.B. .xls -> .xlsx, oder Ausweichname
        # bei Kollision) in die Konvertierungs-CSV.
        if original_path.lower() != target_path.lower():
            log_conversion(original_path, target_path)

        pbar.write(f"  ✓  Aktualisiert: '{_truncate_display(os.path.basename(original_path))}'")
        detail_logger.info(f"Erfolgreich: {original_path}")
        return "UPDATED"

    except Exception:
        display_name = os.path.basename(locals().get("original_path", file_path))
        pbar.write(f"  ✗  FEHLER: '{_truncate_display(display_name)}'")

        if wb is not None:
            try:
                wb.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")
            del wb
            wb = None

        rollback_target = locals().get("target_path")
        if (move_started
                and rollback_target
                and rollback_target != locals().get("original_path", file_path)):
            if safe_exists(rollback_target):
                safe_remove(rollback_target)
                detail_logger.warning(
                    f"Rollback: unvollständige Zieldatei entfernt: {rollback_target}"
                )

        inplace_overwrite = (
            move_started
            and locals().get("target_path") is not None
            and locals().get("target_path", "").lower() == original_path.lower()
        )
        shared_reopen_done = bool(locals().get("_need_writable_reopen"))
        # original_mutated wird gesetzt, sobald ExclusiveAccess() das
        # Original in-place gespeichert hat (shared+writable). Ohne diese
        # Bedingung wuerde der Rollback bei einem spaeteren Fehler (z.B.
        # nicht loeschbares Original beim Formatwechsel) das Backup verwerfen
        # und die mutierte (entfreigegebene) Originaldatei liegen lassen.
        original_mutated = bool(locals().get("original_mutated"))
        needs_restore = (
            not safe_exists(original_path) or inplace_overwrite
            or shared_reopen_done or original_mutated
        )
        if backup_path and safe_exists(backup_path) and needs_restore:
            if restore_backup(backup_path, original_path):
                _apply_security_descriptor(original_path, orig_sd)
                safe_remove(backup_path)
                detail_logger.warning(f"Backup wiederhergestellt: {original_path}")
            else:
                # Restore fehlgeschlagen: Das Backup ist jetzt die einzige
                # intakte Kopie - es darf KEINESFALLS geloescht werden und
                # muss auch das Temp-Aufraeumen ueberleben.
                _preserved_temp_files.add(backup_path)
                detail_logger.error(
                    f"KRITISCH: Backup-Restore fehlgeschlagen – Backup bleibt "
                    f"erhalten: {backup_path} (Original: {original_path})")
                pbar.write(
                    "    ⚠ KRITISCH: Rollback fehlgeschlagen! "
                    "Backup-Datei wird aus Sicherheitsgründen bewahrt.")
        elif backup_path and safe_exists(backup_path):
            safe_remove(backup_path)
            detail_logger.warning(
                f"Backup gelöscht (Original unverändert vorhanden): {original_path}"
            )

        raise

    finally:
        # Reservierten Zielnamen freigeben, falls die Konvertierung nicht
        # bis zum Schreiben kam. Entfernt ausschliesslich 0-Byte-Dateien.
        if target_was_reserved and not target_file_created:
            release_unique_path(target_path)
        try:
            excel_app.Interactive = COM_TRUE
        except Exception as _e:
            detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")

        if wb is not None:
            try:
                wb.Close(SaveChanges=COM_FALSE)
            except Exception as _e:
                detail_logger.debug(f"convert_excel_file: Exception verworfen: {_e!r}")
            del wb
            wb = None
        if is_temp_copy and safe_exists(file_path):
            safe_remove(file_path)
        if temp_save_path and safe_exists(temp_save_path):
            safe_remove(temp_save_path)

# ==================================================================
# Wrapper mit Retry
# ==================================================================
def process_file_with_retries(
    full_path: str,
    excel_app: win32com.client.CDispatch,
    pbar: tqdm,
    passwords: list,
    force_update: bool,
    break_links: bool,
    dry_run: bool = False,
) -> str:
    result = "ERROR"
    for attempt in range(MAX_RETRIES):
        # Aktuelle Instanz aus excel_app_global nachziehen. convert_excel_file
        # startet Excel bei einem Absturz intern neu, kann die Zuweisung aber
        # nur an seinen eigenen Parameter machen - die hier gehaltene Referenz
        # zeigt danach auf die per _quit_excel_app beendete Instanz. Der
        # naechste Versuch lief damit gegen ein totes COM-Objekt.
        # _create_excel_app() setzt excel_app_global; das ist die verlaessliche
        # Quelle fuer die gerade gueltige Instanz.
        if excel_app_global is not None and excel_app_global is not excel_app:
            excel_app = excel_app_global
            detail_logger.debug(
                "process_file_with_retries: Excel-Instanz nach Neustart nachgezogen.")
        try:
            result = convert_excel_file(
                full_path, excel_app, pbar,
                passwords, force_update, break_links, dry_run=dry_run)
            break
        except pythoncom.com_error as e:
            if e.hresult in HRESULT_EXCEL_BUSY_OR_CRASH and attempt < MAX_RETRIES - 1:
                pbar.write(f"  → Excel beschäftigt oder abgestürzt, warte {RETRY_DELAY}s ...")
                time.sleep(RETRY_DELAY)
            else:
                log_error(full_path, e)
                break
        except (OSError, IOError, pywintypes.error, RuntimeError) as e:
            if attempt < MAX_RETRIES - 1:
                pbar.write(f"  → Dateisystem-Fehler, warte {RETRY_DELAY}s ...")
                time.sleep(RETRY_DELAY)
            else:
                log_error(full_path, e)
                break
        except Exception as e:
            log_error(full_path, e)
            pbar.write(f"  ✗  Unerwarteter Fehler ({type(e).__name__}): siehe Detail-Log.")
            break
    return result

# ==================================================================
# Excel-Instanz-Fabrik
# ==================================================================
def _create_excel_app() -> win32com.client.CDispatch:
    global excel_app_global, excel_app_pid

    before = _snapshot_excel_pids()
    app    = win32com.client.DispatchEx("Excel.Application")
    app.Visible            = COM_FALSE
    app.DisplayAlerts      = COM_FALSE
    app.AskToUpdateLinks   = COM_FALSE
    app.AutomationSecurity = 3
    app.EnableEvents       = COM_FALSE
    app.ScreenUpdating     = COM_FALSE

    time.sleep(0.5)
    after    = _snapshot_excel_pids()
    new_pids = after - before

    # ALLE neu entstandenen PIDs als eigene vermerken, nicht nur die
    # ausgewaehlte: aus den uebrigen entstehen genau die Waisen, die
    # _kill_orphaned_excel() aufraeumen soll. Was hier nicht steht, hat das
    # Skript nicht gestartet und wird nie beendet.
    for _p in new_pids:
        remember_own_excel_pid(_p)

    if len(new_pids) == 1:
        excel_app_pid = new_pids.pop()
    elif len(new_pids) > 1:
        excel_app_pid = max(new_pids)
        detail_logger.warning(
            f"Mehrere neue EXCEL.EXE-Prozesse gefunden ({new_pids}) – nehme jüngste PID {excel_app_pid}."
        )
    else:
        excel_app_pid = None
        try:
            _, hwnd_pid = win32process.GetWindowThreadProcessId(app.Hwnd)
            excel_app_pid = hwnd_pid or None
        except Exception as _e:
            detail_logger.debug(f"_create_excel_app: Exception verworfen: {_e!r}")
        remember_own_excel_pid(excel_app_pid)
        if not excel_app_pid:
            detail_logger.warning(
                "Excel-PID konnte nicht ermittelt werden – Watchdog-Timeout fällt aus."
            )

    excel_app_global = app
    detail_logger.debug(f"Excel-Instanz erzeugt (PID: {excel_app_pid})")
    return app

def _quit_excel_app(app) -> None:
    global excel_app_global, excel_app_pid

    pid_to_wait = excel_app_pid
    try:
        app.Quit()
    except Exception as _e:
        detail_logger.debug(f"_quit_excel_app: Exception verworfen: {_e!r}")

    if pid_to_wait:
        proc = None
        try:
            proc = psutil.Process(pid_to_wait)
            if "excel" not in proc.name().lower():
                proc = None
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            proc = None
        except Exception:
            proc = None

        if proc is not None:
            try:
                proc.wait(timeout=5)
            except psutil.NoSuchProcess:
                pass
            except psutil.TimeoutExpired:
                try:
                    if "excel" in proc.name().lower():
                        proc.kill()
                        proc.wait(timeout=3)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
                except Exception as _e:
                    detail_logger.debug(f"_quit_excel_app: Exception verworfen: {_e!r}")
            except Exception as _e:
                detail_logger.debug(f"_quit_excel_app: Exception verworfen: {_e!r}")

    if excel_app_global is app:
        excel_app_global = None
        excel_app_pid    = None

def _cleanup_user_temp() -> None:
    """Räumt den User-TEMP-Ordner (C:\\Users\\<user>\\AppData\\Local\\Temp) best
    effort auf. Gesperrte oder fremde Dateien werden übersprungen, Fehler
    nicht eskaliert. Skips landen nur auf DEBUG-Level, um das Detail-Log
    nicht mit hunderten Permission-Fehlern zu fluten.

    Begründung: Excel hinterlässt unter manchen Bedingungen (Crash, Watchdog-
    Kill) Stub-Dateien (~$*.xlsx, Office-Recovery-Stubs, gen_py-Reste) im
    User-TEMP, die sich über lange Läufe massiv anhäufen.
    """
    temp_dir = tempfile.gettempdir()
    if not temp_dir or not os.path.isdir(temp_dir):
        return

    norm = os.path.normcase(os.path.abspath(temp_dir))
    looks_safe = (
        "\\appdata\\local\\temp" in norm
        or norm.endswith("\\temp")
    )
    if not looks_safe:
        detail_logger.debug(f"TEMP-Cleanup übersprungen (unsicherer Pfad): {temp_dir}")
        return

    removed_files = 0
    removed_dirs  = 0
    try:
        entries = os.listdir(temp_dir)
    except OSError as e:
        detail_logger.debug(f"TEMP-Cleanup: listdir fehlgeschlagen: {e}")
        return

    def _is_office_python_temp(nm: str) -> bool:
        low = nm.lower()
        if low.startswith("~$") or low.startswith("~df"):
            return True
        # gen_py (pywin32-COM-Cache) bewusst NICHT: wird von parallel
        # laufenden COM-Skripten (3a/3c/4a) aktiv genutzt. vbe/excel8.0/cvr
        # sind Excel-eigene Reste und bleiben.
        if low.startswith("vbe") or low.startswith("excel8.0"):
            return True
        if low.startswith("cvr") and low.endswith(".tmp"):
            return True
        if low.startswith(TEMP_PROCESS_PREFIX_LOWER):
            return True
        return False

    for name in entries:
        if not _is_office_python_temp(name):
            continue
        full = os.path.join(temp_dir, name)
        try:
            if os.path.isdir(full) and not os.path.islink(full):
                shutil.rmtree(full)
                removed_dirs += 1
            else:
                try:
                    os.chmod(full, 0o666)
                except OSError:
                    pass
                os.remove(full)
                removed_files += 1
        except (PermissionError, OSError):
            continue
        except Exception:
            continue

    detail_logger.debug(
        f"TEMP-Cleanup ({temp_dir}): {removed_files} Dateien, {removed_dirs} Ordner entfernt"
    )


def _cleanup_excel_inetcache() -> None:
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
                        os.chmod(fpath, 0o666)
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
                    os.chmod(fpath, 0o666)
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


# ==================================================================
# Verzeichnis-Verarbeitung
# ==================================================================
def process_directory(
    directory: str,
    excel_app: win32com.client.CDispatch,
    passwords: list,
    force_update: bool,
    break_links: bool,
    count_files_first: bool,
    show_progress: bool = True,
    resume_path: Optional[str] = None,
    resume_set: Optional[set] = None,
    exclude_patterns: list = None,
    dry_run: bool = False,
) -> Tuple[dict, win32com.client.CDispatch, bool]:
    os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)

    resume_set = resume_set or set()
    # clean_temp=not dry_run: Der Probelauf sichert an drei Stellen zu, das
    # Dateisystem unangetastet zu lassen ('laesst das Dateisystem vollstaendig
    # unangetastet', 'SPEICHERT aber NICHTS', 'es wurde NICHTS geaendert').
    # Mit dem frueher fest verdrahteten clean_temp=True loeschte gerade der
    # Modus, den man vor dem ersten Echt-Lauf auf einer fremden Ablage waehlt,
    # bereits Dateien im gesamten Baum (nachgestellt: 4 von 6 Probedateien).
    gen = file_generator(directory, clean_temp=not dry_run,
                         exclude_patterns=exclude_patterns)
    if resume_set:
        _src_gen = gen
        gen = (f for f in _src_gen if f.lower() not in resume_set)

    if count_files_first:
        print("\nIndiziere und bereinige Dateien (kann einige Minuten dauern) ...")
        files_list: List[str] = []
        for i, f in enumerate(gen, 1):
            files_list.append(f)
            if i % 1000 == 0:
                print(f"  ... {i:>7} Dateien indiziert", flush=True)
        total      = len(files_list)
        iterable   = files_list
        bar_fmt    = "{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]"
        print(f"\n{total} Dateien gefunden"
              + (f" ({len(resume_set)} via Resume übersprungen)" if resume_set else "")
              + ". Starte Verarbeitung ...\n")
    else:
        iterable = gen
        total    = None
        bar_fmt  = "{desc}: {n_fmt} Dateien [{elapsed}, {rate_fmt}]"
        print("\nStarte direkte Verarbeitung (Bereinigung Ordner für Ordner) ...\n")

    print("-" * 66)

    stats = {"UPDATED": 0, "ALREADY_CURRENT": 0, "SKIPPED": 0, "ERROR": 0,
             "RENAMED": 0, "WOULD_UPDATE": 0}

    files_processed = 0
    current_excel   = excel_app
    run_completed   = False

    with tqdm(total=total, desc="Verarbeite", unit="Datei",
              bar_format=bar_fmt, disable=not show_progress) as pbar:
        for full_path in iterable:

            # Im Probelauf NICHT umbenennen - sanitize_file_on_disk wuerde
            # die Datei auf der Platte umbenennen. Der Probelauf laesst das
            # Dateisystem vollstaendig unangetastet.
            if not dry_run:
                full_path, was_renamed = sanitize_file_on_disk(full_path, pbar)
                if was_renamed:
                    stats["RENAMED"] += 1

            needs_restart        = False
            restart_is_periodic  = False

            try:
                _ = current_excel.Version
            except Exception:
                detail_logger.warning("Excel-Instanz reagiert nicht (z. B. nach Timeout). Starte neu...")
                needs_restart = True

            if (not needs_restart
                    and files_processed > 0
                    and files_processed % EXCEL_RESTART_EVERY == 0):
                needs_restart       = True
                restart_is_periodic = True

            if needs_restart:
                if restart_is_periodic:
                    pbar.write(
                        f"  ↻  Excel-Neustart nach {files_processed} Dateien "
                        f"(Speicherbereinigung) ..."
                    )
                try:
                    _restore_excel_settings(current_excel)
                except Exception as _e:
                    detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                _quit_excel_app(current_excel)
                _kill_orphaned_excel()
                _cleanup_excel_inetcache()
                _cleanup_user_recent()
                time.sleep(2)
                current_excel = _create_excel_app()

            result = process_file_with_retries(
                full_path, current_excel, pbar,
                passwords, force_update, break_links, dry_run=dry_run)
            # Wurde Excel waehrend der Verarbeitung neu gestartet, zeigt
            # current_excel sonst weiter auf die beendete Instanz - jede
            # weitere Datei liefe dann gegen ein totes COM-Objekt, und der
            # Lebendtest weiter oben wuerde bei JEDER Datei einen erneuten
            # Neustart ausloesen.
            if excel_app_global is not None and excel_app_global is not current_excel:
                current_excel = excel_app_global
                detail_logger.debug(
                    "process_directory: Excel-Instanz nach Neustart nachgezogen.")
            stats[result] = stats.get(result, 0) + 1
            # Resume im Probelauf NICHT fortschreiben: sonst gaelten die
            # Dateien beim spaeteren Echt-Lauf als bereits erledigt.
            if resume_path and result in RESUME_STATUSES and not dry_run:
                append_resume(resume_path, full_path)
            files_processed += 1
            pbar.update(1)
        else:
            # for-else: Schleife lief ohne break vollstaendig durch.
            run_completed = True

    detail_logger.info(f"Verarbeitung abgeschlossen: {stats}")
    return stats, current_excel, run_completed

# ==================================================================
# Einstiegspunkt
# ==================================================================
if __name__ == "__main__":
    check_required_modules()

    # Fremde Excel-Sitzungen erfassen, BEVOR eine eigene COM-Instanz
    # entsteht - danach waere die eigene nicht mehr unterscheidbar.
    snapshot_foreign_excel_pids()

    start_time = datetime.now()

    print()
    print("╔══════════════════════════════════════════════════════════════╗")
    print("║        EXCEL-UPDATER: XLS/XLSX AUF NEUESTE APPVERSION        ║")
    print("╚══════════════════════════════════════════════════════════════╝")
    print(f"  Start:    {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Versions: {OFFICE_VERSIONS_SUPPORTED}")

    parser = argparse.ArgumentParser(
        description="Excel-Updater: XLS/XLSX auf aktuelle AppVersion konvertieren",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Beispiele:\n"
            "  Interaktiv:   python updater.py\n"
            "  Automatisch:  python updater.py --dir \\\\server\\freigabe --auto-start\n"
            "  Mit Optionen: python updater.py --dir C:\\Daten --count-first\n"
            "\n"
            "WICHTIG – --auto-start:\n"
            "  Ausschließlich für automatisierte Dienste (Task Scheduler, SCCM, CI/CD).\n"
            "  Das Skript beendet im Automatikmodus ALLE Excel-Prozesse des ausführenden\n"
            "  Benutzers ohne Rückfrage und ohne Vorwarnung. Ungespeicherte Änderungen\n"
            "  gehen unwiederbringlich verloren. Niemals interaktiv mit offenen\n"
            "  Excel-Fenstern verwenden!"
        ),
    )
    parser.add_argument(
        "--dir", metavar="PFAD",
        help="Zu verarbeitendes Verzeichnis (überspringt ask_directory)"
    )
    parser.add_argument(
        "--auto-start", action="store_true",
        help="Alle Bestätigungsabfragen überspringen und sofort starten"
    )
    parser.add_argument(
        "--password", metavar="PW", default=None,
        help="Standardpasswort zum Öffnen verschlüsselter Dateien"
    )
    parser.add_argument(
        "--break-links", action="store_true",
        help="Externe Verknüpfungen entfernen (Vorsicht: Datenverlust!)"
    )
    parser.add_argument(
        "--exclude-dir", metavar="MUSTER", action="append", default=[],
        help="Teilpfad-Muster (Ordnername), das vom Durchlauf ausgeschlossen "
             "wird; mehrfach möglich (z.B. --exclude-dir Archiv --exclude-dir _Alt)"
    )
    parser.add_argument("--dry-run", action="store_true",
        help="Probelauf: Excel startet, oeffnet jede Mappe und prueft "
             "Format/Version/Schutz/Namenskonflikte, SPEICHERT aber NICHTS. "
             "Liefert exakte Statistik, ohne eine Datei zu aendern.")
    parser.add_argument(
        "--mailto", metavar="ADRESSE", default=None,
        help="E-Mail-Adresse für Laufbericht nach Abschluss (Komma-getrennt für mehrere)"
    )
    parser.add_argument(
        "--mailfrom", metavar="ADRESSE", default=None,
        help="Absenderadresse für --mailto (Standard: excel-updater@<smtp-host>)"
    )
    parser.add_argument(
        "--smtp", metavar="HOST", default="localhost",
        help="SMTP-Server für --mailto (Standard: localhost)"
    )
    parser.add_argument(
        "--smtp-port", metavar="PORT", type=int, default=25,
        help="SMTP-Port (Standard: 25; für STARTTLS meist 587)"
    )
    parser.add_argument(
        "--smtp-user", metavar="USER", default=None,
        help="SMTP-Benutzername (aktiviert Login + STARTTLS, für Office 365/Gmail)"
    )
    parser.add_argument(
        "--smtp-pass", metavar="PASS", default=None,
        help="SMTP-Passwort für --smtp-user. ACHTUNG: als Kommandozeilen-Argument "
             "in der Prozessliste sichtbar – besser per Umgebungsvariable SMTP_PASS setzen."
    )
    parser.add_argument(
        "--smtp-starttls", action="store_true",
        help="STARTTLS erzwingen, auch ohne --smtp-user"
    )
    parser.add_argument(
        "--force-update", action="store_true",
        help="Alle .xlsx/.xlsm ohne Versionscheck neu speichern"
    )
    parser.add_argument(
        "--count-first", action="store_true",
        help="Dateien vorab zählen (ETA-Anzeige, aber langsamerer Start)"
    )
    parser.add_argument(
        "--no-kill-excel", action="store_true",
        help="Kein Kill laufender Excel-Prozesse – Abbruch, falls Excel läuft"
    )
    parser.add_argument(
        "--resume", metavar="DATEI", default=None,
        help="Resume-Datei: bereits fertige Dateien überspringen (für geplante Tasks)"
    )
    args = parser.parse_args()

    auto_mode = args.auto_start

    if not auto_mode:
        print()
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
        print(f"     • {_docs_base} (temporärer Arbeitsordner des Skripts)")
        print(r"     • Q:\ oder \\server\dfs (Quelldateien)")
        print("     Jeweils: ☑ Unterordner ... sind ebenfalls vertrauenswürdig")
        print("!" * 66)
        print()
        if not ask_yes_no(
            "Wurden diese Einstellungen in Excel vorgenommen?",
            default_yes=False
        ):
            print("\n❌  Abbruch. Bitte konfigurieren Sie erst das Trust-Center in Excel.")
            sys.exit(0)

    if auto_mode and not args.dir:
        print("\n❌  Fehler: --dir fehlt im Automatikmodus.")
        print("    Verwendung: python updater.py --dir \\\\server\\freigabe --auto-start")
        sys.exit(1)

    if args.dir:
        start_dir = sanitize_path(args.dir)
        if not os.path.isdir(start_dir):
            print(f"\n❌  Fehler: Verzeichnis existiert nicht oder ist nicht erreichbar: '{start_dir}'")
            sys.exit(1)
    else:
        start_dir = ask_directory()

    if auto_mode:
        count_first   = args.count_first
        show_progress = True
        passwords     = [args.password] if args.password else []
        break_links   = args.break_links
        force_upd     = args.force_update
    else:
        if args.password:
            passwords = [args.password]
        else:
            passwords = ask_passwords()

        if args.count_first:
            show_progress = True
            count_first   = True
        else:
            show_progress, count_first = ask_progress_mode()

        if args.break_links:
            break_links = True
        else:
            break_links = ask_yes_no(
                "\n⚠️  Externe Verknüpfungen entfernen? (VORSICHT: Datenverlust möglich!)"
            )

        if args.force_update:
            force_upd = True
        else:
            force_upd = ask_yes_no(
                "\nForce Update: ALLE .xlsx/.xlsm auch ohne Versions-Check neu speichern?\n"
                "  (Nur nötig wenn Kompatibilitätsmodus bei alten Office-2007-Dateien bleibt)"
            )

        if not args.dry_run:
            args.dry_run = ask_yes_no(
                "\nProbelauf (Dry-Run)? Es wird geprüft, aber NICHTS gespeichert.\n"
                "  (Empfohlen vor dem ersten Echt-Lauf auf einer neuen Ablage)"
            )

    # --- Resume-Datei bestimmen ---
    # CLI --resume hat Vorrang (bewusst persistent fuer geplante Tasks).
    # Ohne --resume wird im interaktiven Modus eine verzeichnisspezifische
    # Resume-Datei angeboten und nach vollstaendigem Lauf geloescht.
    resume_file       = args.resume
    auto_resume_owned = False
    resume_set: set   = set()
    if resume_file:
        resume_set = load_resume_set(resume_file)
    elif not auto_mode:
        candidate = get_auto_resume_path(start_dir, _log_dir_global)
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

    print()
    print("=" * 66)
    print("  KONFIGURATION")
    print("=" * 66)
    print(f"  Verzeichnis:           {start_dir}")
    print(f"  Office-Versionen:      {OFFICE_VERSIONS_SUPPORTED}")
    print(f"  Ziel-AppVersion:       {TARGET_APP_VERSION} (Excel 2016/2019/2021/2024/365)")
    print(f"  Ziel-Schema-Level:     lastEdited ≥ {TARGET_LAST_EDITED} (Excel 2019/2021/2024/365)")
    print(f"  Modus:                 {'Vorab zählen (ETA)' if count_first else 'Generator (kein ETA)'}")
    print(f"  Fortschrittsbalken:    {'an' if show_progress else 'aus (inline)'}")
    # Die Zeile behauptete unabhaengig von der Lage 'Überspringen' und stand
    # damit im Widerspruch zur direkt folgenden Zeile 'N hinterlegt'. Mit
    # hinterlegten Kennwoertern werden verschluesselte Mappen sehr wohl
    # geoeffnet und verarbeitet - nur bleiben sie jetzt auch verschluesselt.
    if passwords:
        print(f"  Passwortdateien:       Öffnen mit hinterlegtem Kennwort")
        print(f"                         (Verschlüsselung bleibt erhalten)")
        print(f"  Passwörter:            {len(passwords)} hinterlegt")
    else:
        print("  Passwortdateien:       Überspringen (als SKIPPED loggen)")
        print("  Passwörter:            keines hinterlegt")
    print(f"  Externe Links:         {'ENTFERNEN ⚠️' if break_links else 'Behalten'}")
    print(f"  Force Update:          {'Ja (alle xlsx/xlsm)' if force_upd else 'Nein'}")
    if args.exclude_dir:
        print(f"  Ausgeschlossen:        {', '.join(args.exclude_dir)}")
    if resume_file and (resume_set or not auto_resume_owned):
        _rlabel = ("automatisch" if auto_resume_owned else resume_file)
        _rcount = f"{len(resume_set)} übersprungen" if resume_set else "neu"
        print(f"  Resume:                {_rcount} ({_rlabel})")
    if args.mailto:
        print(f"  E-Mail-Bericht:        {args.mailto} via {args.smtp}")
    print(f"  Excel-Neustart alle:   {EXCEL_RESTART_EVERY} Dateien (Speicherbereinigung)")
    if auto_mode:
        print("  Modus:                 AUTOMATISCH (keine Bestätigung erforderlich)")
    print()
    print("  Verarbeitungslogik:")
    print("    • .xls/.xlt/.xla  → immer zu .xlsx/.xltx konvertieren")
    print(f"    • AppVersion < {TARGET_APP_VERSION} → auf aktuelle Version aktualisieren")
    print(f"    • lastEdited < {TARGET_LAST_EDITED} → Schema-Migration durch SaveAs")
    print("    • Workbook- und Sheet-Schutz wird entfernt (ohne Passwort)")
    print("    • Druckbereiche/Drucktitel/AutoFilter-Bereiche werden entfernt")
    print("       (verhindert modale Konfliktdialoge bei legacy Dateien)")
    print()
    print("  lastEdited-Referenz (xl/workbook.xml):")
    print(f"      7 = Excel 2019 / 2021 / 2024 / 365  ← Ziel (Schema-Level {TARGET_LAST_EDITED})")
    print("      6 = Excel 2013 / 2016")
    print("      5 = Excel 2010")
    print("      4 = Excel 2007")
    print()
    print("  AppVersion-Referenz (docProps/app.xml):")
    print("    16.0 = Excel 2016 / 2019 / 2021 / 2024 / 365")
    print("    15.0 = Excel 2013")
    print("    14.0 = Excel 2010")
    print("    12.0 = Excel 2007")
    print("     3.x = Drittanbieter (z.B. openpyxl)")
    print("=" * 66)

    if not auto_mode and not ask_yes_no("\nJetzt starten?"):
        print("Abgebrochen.")
        sys.exit(0)

    signal.signal(signal.SIGINT, _signal_handler)
    if hasattr(signal, "SIGBREAK"):
        signal.signal(signal.SIGBREAK, _signal_handler)

    _check_excel_running_warning(auto_mode, no_kill=args.no_kill_excel)

    lock_path = _acquire_lock(start_dir)
    if lock_path is None:
        print()
        print("❌  FEHLER: Es läuft bereits ein Skript-Lauf auf diesem Verzeichnis.")
        print(f"    Lock-Datei: {_lock_file_path(start_dir)}")
        print("    Wenn Du sicher bist, dass kein anderer Lauf aktiv ist,")
        print("    lösche die Lock-Datei manuell und starte erneut.")
        sys.exit(1)
    active_lock_path = lock_path

    _kill_orphaned_excel()
    pythoncom.CoInitialize()

    # Restore-Privilegien (Admin-Kontext) fuer ACL/Owner-Erhalt aktivieren.
    _enable_restore_privileges()

    excel_app    = None
    stats_result = None

    try:
        excel_app = _create_excel_app()

        excel_version = "Unbekannt"
        try:
            excel_version = excel_app.Version
        except Exception:
            pass
        print(f"\n  Excel-Version: {excel_version}")

        # --- Trust-Center Smoke-Test ---
        print("\n  Prüfe COM-Subsystem und Trust-Center ...")
        smoke_ok, smoke_msg = test_trust_center_smoke(excel_app, timeout=25.0)
        if not smoke_ok:
            print()
            print("=" * 66)
            print("  ⚠  EXCEL-SMOKE-TEST FEHLGESCHLAGEN")
            print("=" * 66)
            print(f"  Grund: {smoke_msg}")
            print()
            print("  Mögliche Ursachen:")
            print("    • Dokumentenordner nicht als vertrauenswürdiger Speicherort eingetragen")
            print("    • Geschützte Ansicht für unsichere Speicherorte noch aktiv")
            print("    • Excel/Office-Profil beschädigt oder fehlende Desktop-Ordner")
            print("      (siehe Trust-Center-Hinweis im Skript-Header)")
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
                _quit_excel_app(excel_app)
            except Exception:
                pass
            _kill_orphaned_excel()
            excel_app = _create_excel_app()
        else:
            detail_logger.info("Trust-Center-Smoke-Test erfolgreich.")
            print("    → Smoke-Test OK.")

        if args.dry_run:
            print(f"\nPROBELAUF (Dry-Run) – es wird NICHTS gespeichert: '{start_dir}'")
        else:
            print(f"\nVerarbeite: '{start_dir}'")
        print("-" * 66)

        stats_result, excel_app, run_completed = process_directory(
            start_dir, excel_app,
            passwords, force_upd, break_links, count_first,
            show_progress=show_progress,
            resume_path=resume_file, resume_set=resume_set,
            exclude_patterns=args.exclude_dir,
            dry_run=args.dry_run)

        # Auto-Resume-Datei nach vollstaendigem Lauf entfernen (eine vom
        # Nutzer via --resume uebergebene Datei bleibt unangetastet).
        if auto_resume_owned and run_completed and resume_file:
            delete_resume_file(resume_file)
            detail_logger.info("Auto-Resume-Datei nach vollständigem Lauf gelöscht.")

    except Exception as e:
        print(f"\n❌ KRITISCHER FEHLER: {e}")
        detail_logger.error(
            f"KRITISCHER FEHLER ({type(e).__name__}): {e}\n{traceback.format_exc()}"
        )
        log_error("GLOBAL", e)

    finally:
        active_excel = excel_app_global if excel_app_global is not None else excel_app
        if active_excel is not None:
            try:
                _restore_excel_settings(active_excel)
            except Exception:
                pass
            _quit_excel_app(active_excel)
        _kill_orphaned_excel()
        if os.path.exists(TEMP_PROCESS_PATH):
            time.sleep(1)
            _safe_cleanup_temp()
        try:
            _cleanup_user_temp()
        except Exception:
            pass
        try:
            _cleanup_excel_inetcache()
        except Exception:
            pass
        try:
            _cleanup_user_recent()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        _release_lock(lock_path)

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
        renamed_count = stats_result.get("RENAMED", 0)
        total_sum = sum(v for k, v in stats_result.items() if k != "RENAMED")
        print("  STATISTIK:")
        if args.dry_run:
            print(f"    [✓] Würde aktualisiert/konv.: {stats_result.get('WOULD_UPDATE', 0)}")
        else:
            print(f"    [✓] Aktualisiert / Konv.: {stats_result.get('UPDATED', 0)}")
        print(f"    [=] Bereits aktuell:       {stats_result.get('ALREADY_CURRENT', 0)}")
        print(f"    [→] Übersprungen:          {stats_result.get('SKIPPED', 0)}")
        print(f"    [✗] Fehler:                {stats_result.get('ERROR', 0)}")
        print(f"        GESAMT:                {total_sum}")
        if renamed_count > 0:
            print(f"    [📝] Dateinamen bereinigt: {renamed_count}")
        file_logger.error(
            f"VERARBEITUNG ABGESCHLOSSEN | "
            f"Verzeichnis: {start_dir} | "
            f"Dauer: {str(duration).split('.')[0]} | "
            f"Aktualisiert: {stats_result.get('UPDATED', 0)} | "
            f"Aktuell: {stats_result.get('ALREADY_CURRENT', 0)} | "
            f"Übersprungen: {stats_result.get('SKIPPED', 0)} | "
            f"Fehler: {stats_result.get('ERROR', 0)} | "
            f"Dateinamen bereinigt: {renamed_count} | "
            f"Gesamt: {total_sum}"
        )

        write_run_summary(stats_result, start_dir, duration, force_upd, break_links)

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

    if not auto_mode:
        print()
        print("!" * 66)
        print("  WICHTIG: TRUST-CENTER-EINSTELLUNGEN ZURÜCKSETZEN")
        print("!" * 66)
        print("  Die vor dem Lauf gelockerten Sicherheitseinstellungen")
        print("  sollten jetzt wieder auf die Unternehmensstandards")
        print("  zurückgesetzt werden:")
        print()
        print("  Excel → Datei → Optionen → Trust Center →")
        print("  'Einstellungen für das Trust Center...'")
        print()
        print("  1. GESCHÜTZTE ANSICHT:")
        print("     ☑ Alle 3 Optionen wieder AKTIVIEREN")
        print("  2. MAKROEINSTELLUNGEN:")
        print("     ⦿ Deaktivieren von VBA-Makros mit Benachrichtigung")
        print("     ☐ Zugriff auf das VBA-Projektobjektmodell wieder DEAKTIVIEREN")
        print("  3. VERTRAUENSWÜRDIGE SPEICHERORTE:")
        print(f"     → {_docs_base}")
        print("       kann aus der Liste entfernt werden")
        print("     → 'Vertrauenswürdige Speicherorte im Netzwerk zulassen'")
        print("       ggf. wieder DEAKTIVIEREN")
        print("!" * 66)

    # --- E-Mail-Bericht (für --auto-start-Serverbetrieb) ---
    if args.mailto and stats_result:
        _from = args.mailfrom or f"excel-updater@{args.smtp}"
        # Passwort bevorzugt aus Umgebungsvariable (nicht in Prozessliste sichtbar).
        _smtp_pass = os.environ.get("SMTP_PASS") or args.smtp_pass
        send_summary_mail(
            args.mailto, _from, args.smtp, stats_result,
            duration, start_dir, os.path.abspath(LOG_FILE),
            smtp_port=args.smtp_port,
            smtp_user=args.smtp_user,
            smtp_pass=_smtp_pass,
            smtp_starttls=args.smtp_starttls,
        )

    print()
    try:
        input("Zum Beenden Eingabetaste drücken ...")
    except EOFError:
        pass
