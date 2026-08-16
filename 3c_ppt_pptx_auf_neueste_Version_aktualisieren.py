# -*- coding: utf-8 -*-
# ==================================================================
# POWERPOINT-UPDATER: ALLE PPT/PPTX DURCH AKTUELLE ENGINE NEU SPEICHERN
# ==================================================================
# Datum: 13.06.2026
# ANLEITUNG:
# 1. Python installieren: https://www.python.org/
# 2. In der Konsole (CMD): pip install pywin32 tqdm psutil
# 3. Falls pywin32 erstmals installiert: python Scripts/pywin32_postinstall.py -install
# 4. In der Konsole (CMD): python 3c_ppt_pptx_auf_neueste_Version_aktualisieren.py
#
# FUNKTIONSWEISE:
# PowerPoint besitzt – anders als Word – KEINE Presentation.Convert()-Methode.
# Der Kompatibilitätsmodus und Legacy-Strukturen lassen sich nicht programmgesteuert
# erkennen oder gezielt migrieren. Der einzige zuverlässige Weg: Presentation.SaveAs()
# mit dem Zielformat. Dabei serialisiert PowerPoint die gesamte Datei neu und migriert
# alle internen Schemas auf die aktuelle Engine.
#
# Deshalb wird JEDE Datei ausnahmslos durch SaveAs geschleust – unabhängig
# von der AppVersion in docProps/app.xml.
#
# WICHTIG: POWERPOINT SICHERHEITSEINSTELLUNGEN (TRUST CENTER)
# ------------------------------------------------------------------
# PowerPoint öffnen → Datei → Optionen → Trust Center →
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
# WICHTIG: AUTOMATISIERUNG ALS NT-AUTORITÄT\SYSTEM (SCCM / TASK SCHEDULER)
# ------------------------------------------------------------------
# Wenn das Skript über einen Dienst im SYSTEM-Kontext läuft (headless),
# schlägt win32com.client.DispatchEx("PowerPoint.Application") mit einem
# DCOM-Fehler fehl (0x80070005 / 0x80080005), weil Office-Anwendungen eine
# Desktop-Struktur im Benutzerprofil benötigen.
#
# Lösung (einmalig auf dem ausführenden Server als Administrator):
#   mkdir "C:\Windows\System32\config\systemprofile\Desktop"
#   mkdir "C:\Windows\SysWOW64\config\systemprofile\Desktop"   (64-Bit)
#
# Danach funktioniert --auto-start im SYSTEM-Kontext zuverlässig.
#
# HINWEIS FÜR PyInstaller-Build (Umwandlung in .exe):
#   # Einmalig: PyInstaller installieren
#   pip install pyinstaller
#
#   # Build als ein einzeiliger Befehl (am sichersten in jeder Shell):
#   python -m PyInstaller --onefile --noupx --noconfirm --clean --console --icon="python_icon.ico" --name "3c_ppt_pptx_auf_neueste_Version_aktualisieren" --hidden-import pywintypes --hidden-import pythoncom --hidden-import win32api --hidden-import win32file --hidden-import win32process --hidden-import win32con --hidden-import win32com.client --hidden-import winreg --hidden-import psutil --hidden-import tqdm 3c_ppt_pptx_auf_neueste_Version_aktualisieren.py
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
#     Interpreters (64-Bit-Python -> x64-EXE; 32-Bit-Python -> x86-EXE). PowerPoint
#     registriert seinen COM-Server als LocalServer32 (out-of-process,
#     POWERPNT.EXE); das Windows-COM-Subsystem marshallt Aufrufe zwischen
#     x64-Aufrufer und x86-Server (und umgekehrt) automatisch über DCOM/LRPC.
#     Eine x64-EXE arbeitet daher auch mit 32-Bit-PowerPoint zusammen (und umgekehrt).
#     Bitness-Match liefert minimal bessere Performance, ist aber nicht
#     erforderlich. 64-Bit-Python ist eine sichere Default-Wahl, da modernes
#     Office (2019+, M365) standardmäßig x64 ist.
#
# ==================================================================

import argparse
import os
import stat
import sys
import shutil
import signal
import smtplib
import logging
import importlib.util
import tempfile
import time
import uuid
import zipfile
import re
import csv
import hashlib
import winreg
import threading
import email.mime.multipart
import email.mime.text
from datetime import datetime
from typing import Optional, Tuple


# ==================================================================
# Konsolen-Encoding (UTF-8) - muss VOR jedem print stehen
# ==================================================================
# Ohne diesen Block bricht die erste Ausgabe mit Rahmenzeichen oder Emoji
# unter der Windows-Standardcodepage (cp850/cp1252) mit UnicodeEncodeError
# ab - auch bei Umleitung in eine Datei. Gleiche Fassung wie in 3a/3b.
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



IS_FROZEN = getattr(sys, "frozen", False)


def check_required_modules() -> None:
    if IS_FROZEN:
        return
    required = {
        "win32com.client": "pywin32",
        "win32file":       "pywin32",
        "win32api":        "pywin32",
        "win32con":        "pywin32",
        "psutil":          "psutil",
        "tqdm":            "tqdm",
        "msoffcrypto":     "msoffcrypto-tool",
    }
    missing = sorted({pkg for mod, pkg in required.items()
                      if not importlib.util.find_spec(mod.split(".")[0])})
    if missing:
        print("=" * 66)
        print("❌ FEHLENDE MODULE:")
        for pkg in missing:
            print(f"   pip install {pkg}")
        print("=" * 66)
        sys.exit(1)


check_required_modules()

import pythoncom
import win32com.client
import win32api
import win32con
import win32file
import win32process
import psutil
from tqdm import tqdm
import msoffcrypto

# ==================================================================
# COM-Konstanten
# ==================================================================
COM_TRUE  = -1   # VARIANT_TRUE
COM_FALSE =  0   # VARIANT_FALSE

# PpSaveAsFileType
PP_FORMAT_PPTX = 24   # ppSaveAsOpenXMLPresentation
PP_FORMAT_PPTM = 25   # ppSaveAsOpenXMLPresentationMacroEnabled
PP_FORMAT_PPSX = 28   # ppSaveAsOpenXMLShow
PP_FORMAT_PPSM = 29   # ppSaveAsOpenXMLShowMacroEnabled
PP_FORMAT_POTX = 26   # ppSaveAsOpenXMLTemplate
PP_FORMAT_POTM = 27   # ppSaveAsOpenXMLTemplateMacroEnabled
PP_FORMAT_PPAM = 30   # ppSaveAsOpenXMLAddin

# Alte PPT-FileFormat-Codes (ppSaveAsPresentation=1, ppSaveAsDefault=11)
PP_OLD_FORMAT_CODES = {1, 11}

# Ziel-AppVersion (PowerPoint 2016/2019/2021/365)
TARGET_APP_VERSION = 16.0

# Verzeichnisnamen, die nie betreten werden: geloeschte Praesentationen
# im $RECYCLE.BIN wuerden sonst mitkonvertiert (und ihr Original ersetzt);
# ~snapshot/.snapshot sind read-only NAS-Schattenkopien.
EXCLUDE_DIR_NAMES = {"$recycle.bin", "system volume information",
                     "~snapshot", ".snapshot"}

FILE_ATTRIBUTE_REPARSE_POINT = 0x0400

# ==================================================================
# Protokoll-Objekte fruehzeitig binden: _resolve_documents_dir() laeuft schon
# beim Import und 'verwirft' Fehler mit detail_logger.debug(...) - der Name
# wird aber erst am Ende von _setup_logging() belegt. Im Fehlerfall (Documents
# nicht erreichbar, umgeleitetes Profil) loeste der Fehlerschlucker deshalb
# selbst einen NameError aus und das Skript startete gar nicht.
# getLogger liefert dieselben Objekte, die _setup_logging() spaeter mit
# Handlern versieht - die Zuweisung dort bleibt unveraendert gueltig.
file_logger   = logging.getLogger("FileLogger")
detail_logger = logging.getLogger("DetailLogger")


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

def _resolve_script_directory() -> str:
    try:
        if getattr(sys, "frozen", False):
            return os.path.dirname(os.path.abspath(sys.executable))
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()

_docs_base = _resolve_documents_dir()

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

TEMP_PROCESS_PATH = os.path.join(
    _docs_base, f"3c_ppt_pptx_auf_neueste_Version_aktualisieren_{os.getpid()}"
)
_TEMP_PROCESS_PREFIX_LOWER = "3c_ppt_pptx_auf_neueste_version_aktualisieren_"
MAX_PATH_LEN           = 240
LOG_FILE: str          = ""
DETAILED_LOG_FILE: str = ""
RUN_SUMMARY_FILE: str  = ""
CONVERSIONS_CSV: str   = ""
_LOG_DIR: str          = ""
MAX_RETRIES            = 3
RETRY_DELAY            = 2     # Sekunden
PPT_RESTART_INTERVAL   = 25    # PowerPoint-Neustart nach N Dateien (RAM-Freigabe)

# AV-Scanner-Toleranz (Datei-Sperre durch Virenscanner abwarten)
AV_MAX_RETRIES = 15
AV_RETRY_DELAY = 1.0           # Sekunden

# Watchdog-Timeouts fuer COM-Aufrufe im Hauptlauf. 
PPT_OPEN_TIMEOUT   = 180.0     # Sekunden fuer Presentations.Open
PPT_SAVEAS_TIMEOUT = 240.0     # Sekunden fuer SaveAs (kann bei grossen
                               # Dateien laenger dauern als Open)

# ==================================================================
# Logging  (Dual-Logger)
# ==================================================================
def _setup_logging() -> tuple:
    global LOG_FILE, DETAILED_LOG_FILE, RUN_SUMMARY_FILE, CONVERSIONS_CSV, _LOG_DIR

    log_dir = _select_log_directory()
    _LOG_DIR = log_dir
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    LOG_FILE          = os.path.join(
        log_dir,
        f"3c_ppt_pptx_auf_neueste_Version_aktualisieren_{ts}.log"
    )
    DETAILED_LOG_FILE = os.path.join(
        log_dir,
        f"3c_ppt_pptx_auf_neueste_Version_aktualisieren_detailed_{ts}.log"
    )
    RUN_SUMMARY_FILE  = os.path.join(
        log_dir, "3c_ppt_pptx_auf_neueste_Version_aktualisieren_last_run.txt"
    )
    CONVERSIONS_CSV   = os.path.join(
        log_dir, f"3c_ppt_pptx_auf_neueste_Version_aktualisieren_konvertierungen_{ts}.csv"
    )

    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    fh = logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8")
    fh.setFormatter(fmt)
    fl = logging.getLogger("FileLogger")
    fl.setLevel(logging.INFO)
    fl.addHandler(fh)

    dh = logging.FileHandler(DETAILED_LOG_FILE, mode="w", encoding="utf-8")
    dh.setFormatter(fmt)
    dl = logging.getLogger("DetailLogger")
    dl.setLevel(logging.DEBUG)
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
    # Nachweisliste der Format-Konvertierungen (.ppt/.pps/.pot -> neu, und
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


def write_run_summary(stats: dict, directory: str, duration) -> None:
    if not RUN_SUMMARY_FILE:
        return
    try:
        total = sum(v for k, v in stats.items() if k != "RENAMED")
        with open(RUN_SUMMARY_FILE, "w", encoding="utf-8-sig") as fh:
            fh.write("POWERPOINT-UPDATER – LETZTER LAUF\n")
            fh.write("=" * 60 + "\n")
            fh.write(f"Verzeichnis:    {directory}\n")
            fh.write(f"Dauer:          {str(duration).split('.')[0]}\n")
            fh.write("\n")
            fh.write(f"Aktualisiert:   {stats.get('UPDATED', 0)}\n")
            fh.write(f"Übersprungen:   {stats.get('SKIPPED', 0)}\n")
            fh.write(f"Fehler:         {stats.get('ERROR', 0)}\n")
            fh.write(f"Dateinamen ber.:{stats.get('RENAMED', 0)}\n")
            fh.write(f"Gesamt:         {total}\n")
    except Exception as e:
        detail_logger.warning(f"Run-Summary nicht schreibbar: {e}")


# ==================================================================
# E-Mail-Bericht (SMTP) – Parität zu 3a/3b, für Headless-Betrieb
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
            f'{esc(v)}</td></tr>' for l, v in pairs)

    total = sum(v for k, v in stats.items() if k != "RENAMED")
    stat_rows = rows([
        ("Aktualisiert / Konvertiert", stats.get("UPDATED", 0)),
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
        '<h2 style="margin:0 0 8px;">PowerPoint-Updater – Laufbericht</h2>'
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
    subject = (f"PPT-Updater abgeschlossen – "
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
        f"PowerPoint-Updater – Laufbericht\n"
        f"{'=' * 60}\n"
        f"Verzeichnis:        {start_dir}\n"
        f"Laufzeit:           {str(duration).split('.')[0]}\n"
        f"\nSTATISTIK:\n"
        f"  Aktualisiert:     {stats.get('UPDATED', 0)}\n"
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
# Auto-Resume (Neustart-Faehigkeit grosser Laeufe)
# ==================================================================
def get_auto_resume_path(start_dir: str, log_dir: str) -> str:
    key = hashlib.md5(
        os.path.normcase(os.path.abspath(start_dir)).encode("utf-8")
    ).hexdigest()[:12]
    return os.path.join(
        log_dir, f"3c_ppt_pptx_auf_neueste_Version_aktualisieren_resume_{key}.txt")


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
# Globale PPT-Referenz (single source of truth)
# ==================================================================
ppt_app_global: Optional[win32com.client.CDispatch] = None
ppt_app_pid:    Optional[int]                       = None

# ==================================================================
# Long-Path- / Anzeige-Hilfen (vorgezogen, da unten genutzt)
# ==================================================================

def prepare_long_path(path: str) -> str:
    if path.startswith("\\\\?\\"):
        return path
    path = os.path.abspath(os.path.normpath(path))
    if path.startswith("\\\\"):
        return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path


def strip_com_path(path: str) -> str:
    if not path:
        return path
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\"):
        return path[4:]
    return path


# Lesbare Pfad-Darstellung für Konsole/Log
display_path = strip_com_path


# ==================================================================
# Einzelinstanz-Lock je Zielverzeichnis (Schutz gegen parallele Laeufe)
# ==================================================================
# Verhindert, dass zwei parallele Laeufe (z.B. doppelt gestarteter
# Scheduled Task) gleichzeitig dieselben Dateien konvertieren und sich
# gegenseitig os.replace/Loeschungen zerschiessen. Der Lock haengt am
# Zielverzeichnis (Hash), nicht am Temp-Ordner (der ist ohnehin PID-
# getrennt). PID + Name + Erstellungszeit schuetzen gegen PID-Recycling:
# eine veraltete PID, die laengst an ein anderes Python-Skript vergeben
# wurde, sperrt den Updater dadurch NICHT dauerhaft aus.
active_lock_path: Optional[str] = None


def _lock_file_path(target_dir: str) -> str:
    digest = hashlib.sha1(
        os.path.normcase(os.path.abspath(target_dir)).encode("utf-8")
    ).hexdigest()[:16]
    return os.path.join(_docs_base, f"3c_ppt_updater_{digest}.lock")


def _acquire_lock(target_dir: str) -> Optional[str]:
    lock_path = _lock_file_path(target_dir)
    if os.path.exists(lock_path):
        try:
            with open(lock_path, "r", encoding="utf-8") as f:
                content = [l.strip() for l in f.read().splitlines() if l.strip()]
            old_pid = int(content[0]) if content else 0
            old_ct  = None
            if len(content) > 2:
                try:
                    old_ct = float(content[2])
                except ValueError:
                    old_ct = None
            if old_pid and old_pid != os.getpid() and psutil.pid_exists(old_pid):
                try:
                    p = psutil.Process(old_pid)
                    pname = p.name().lower()
                    looks_like_us = "powerpnt" in pname or "python" in pname or pname.startswith("3c_")
                    if looks_like_us:
                        if old_ct is not None:
                            # Nur sperren, wenn es WIRKLICH derselbe Prozess
                            # ist (Name + Erstellungszeit) - schuetzt vor
                            # PID-Recycling auf ein fremdes Python-Skript.
                            try:
                                if abs(p.create_time() - old_ct) < 1.0:
                                    return None
                            except (psutil.NoSuchProcess, psutil.AccessDenied):
                                pass
                        else:
                            return None  # altes Lock-Format ohne create_time
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            # PID tot/fremd -> Stale-Lock entfernen.
            os.remove(lock_path)
            detail_logger.warning(f"Stale-Lock entfernt (PID {old_pid}): {lock_path}")
        except Exception as e:
            detail_logger.warning(f"Lock-Check fehlgeschlagen: {e}")
            return None
    try:
        own_ct = ""
        try:
            own_ct = repr(psutil.Process(os.getpid()).create_time())
        except Exception:
            own_ct = ""
        with open(lock_path, "w", encoding="utf-8") as f:
            f.write(f"{os.getpid()}\n{datetime.now().isoformat()}\n{own_ct}\n{target_dir}\n")
        return lock_path
    except Exception as e:
        detail_logger.warning(f"Lock-Datei nicht schreibbar: {e}")
        return None


def _release_lock(lock_path: Optional[str]) -> None:
    if lock_path and os.path.exists(lock_path):
        try:
            os.remove(lock_path)
        except Exception as _e:
            detail_logger.debug(f"_release_lock: Exception verworfen: {_e!r}")


# ==================================================================
# Signal-Handler (Ctrl+C / SIGTERM)
# ==================================================================
def _signal_handler(sig, frame) -> None:
    print("\n\n*** ABBRUCH durch Benutzer – raeume auf ...")

    if ppt_app_global is not None:
        try:
            ppt_app_global.Quit()
            time.sleep(1)
        except Exception as _e:
            detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")

    _kill_user_powerpoint()

    try:
        _cleanup_ppt_inetcache()
    except Exception as _e:
        detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")
    try:
        _cleanup_user_recent()
    except Exception as _e:
        detail_logger.debug(f"_signal_handler: Exception verworfen: {_e!r}")

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

    _release_lock(active_lock_path)
    os._exit(1)


# ==================================================================
# Hilfsfunktionen
# ==================================================================

def log_error(file_path: str, exc: Exception) -> None:
    file_logger.error(f"Datei: {display_path(file_path)}\n  -> {exc}\n")
    detail_logger.error(
        f"Datei: {display_path(file_path)}\n  -> {exc}\n", exc_info=True
    )


# ==================================================================
# Schutz fremder PowerPoint-Sitzungen
# ==================================================================
# _kill_user_powerpoint() beendete trotz seines Namens JEDEN PowerPoint-Prozess des
# angemeldeten Benutzers - ohne jede Waisen-Pruefung. Hatte der Anwender
# PowerPoint mit ungespeicherter Arbeit offen, waren diese Dokumente beim
# ersten Aufraeumen (auch beim Abbruch mit Strg+C) verloren.
#
# Vor dem Start der eigenen COM-Instanz wird deshalb einmal festgehalten,
# welche PowerPoint-Prozesse es bereits gab. Diese gelten dauerhaft als fremd
# und werden nie beendet.
_FOREIGN_POWERPOINT_PIDS: set = set()


def snapshot_foreign_powerpoint_pids() -> None:
    """Merkt sich alle PowerPoint-Prozesse, die vor dem Skriptstart liefen."""
    global _FOREIGN_POWERPOINT_PIDS
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
        # Im Zweifel lieber zu viel schuetzen als eine fremde Sitzung killen.
        detail_logger.debug(f"snapshot_foreign_powerpoint_pids: Exception verworfen: {_e!r}")
    _FOREIGN_POWERPOINT_PIDS = found
    if found:
        detail_logger.info(
            f"{len(found)} bereits laufende(r) PowerPoint-Prozess(e) erkannt "
            f"- diese werden nicht beendet: {sorted(found)}")


def is_foreign_powerpoint_pid(pid) -> bool:
    return pid in _FOREIGN_POWERPOINT_PIDS


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


def _kill_user_powerpoint() -> None:
    try:
        current_user = psutil.Process().username()
    except Exception:
        detail_logger.warning(
            "_kill_user_powerpoint: Benutzerermittlung fehlgeschlagen "
            "– Prozesse werden nicht beendet.")
        return
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if proc.info["name"] and "POWERPNT.EXE" in proc.info["name"].upper():
                try:
                    proc_user = proc.username().lower().split("\\")[-1].split("@")[0]
                    cur_user  = current_user.lower().split("\\")[-1].split("@")[0]
                    if proc_user != cur_user:
                        continue
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue
                if is_foreign_powerpoint_pid(proc.info["pid"]):
                    continue
                try:
                    proc.kill()
                    proc.wait(timeout=3)
                    detail_logger.debug(f"PPT-Prozess beendet: PID {proc.info['pid']}")
                except psutil.AccessDenied:
                    detail_logger.warning(f"Kein Zugriff auf PID {proc.info['pid']}")
                except psutil.NoSuchProcess:
                    pass
                except psutil.TimeoutExpired:
                    detail_logger.warning(f"Prozess reagiert nicht: PID {proc.info['pid']}")
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue


def _configure_ppt_instance(ppt_app) -> None:
    global ppt_app_pid

    ppt_app.Visible = COM_TRUE
    try:
        ppt_app.WindowState = 2   # ppWindowMinimized
    except Exception as _e:
        detail_logger.debug(f"_configure_ppt_instance: Exception verworfen: {_e!r}")
    try:
        ppt_app.DisplayAlerts = 1   # ppAlertsNone
    except Exception as _e:
        detail_logger.debug(f"_configure_ppt_instance: Exception verworfen: {_e!r}")
    try:
        ppt_app.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
    except Exception as _e:
        detail_logger.debug(f"_configure_ppt_instance: Exception verworfen: {_e!r}")

    ppt_app_pid = None
    try:
        hwnd_attr = ppt_app.HWND
        hwnd_val  = hwnd_attr() if callable(hwnd_attr) else hwnd_attr
        if hwnd_val:
            _, pid = win32process.GetWindowThreadProcessId(int(hwnd_val))
            if pid:
                ppt_app_pid = int(pid)
    except Exception as e:
        detail_logger.debug(f"PID-Ermittlung via HWND fehlgeschlagen: {e}")

    if ppt_app_pid is None:
        try:
            current_user = psutil.Process().username().lower().split("\\")[-1]
            candidates = []
            for proc in psutil.process_iter(["pid", "name", "create_time"]):
                try:
                    name = proc.info["name"]
                    if not name or "POWERPNT.EXE" not in name.upper():
                        continue
                    proc_user = proc.username().lower().split("\\")[-1].split("@")[0]
                    if proc_user == current_user:
                        candidates.append((proc.info["create_time"], proc.info["pid"]))
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue
            if candidates:
                ppt_app_pid = max(candidates)[1]   # juengster Prozess
                detail_logger.debug(f"PID via psutil-Fallback: {ppt_app_pid}")
        except Exception as e:
            detail_logger.debug(f"PID via psutil-Fallback fehlgeschlagen: {e}")


def _ppt_call_with_watchdog(call_label: str, timeout: float, ppt_pid,
                             call_fn, restore_fn=None):
    done_event   = threading.Event()
    timeout_flag = [False]
    # Schloss + Fertig-Kennzeichen gegen ein schmales, aber echtes
    # Zeitfenster: der Rueckgabewert von call_fn() steht fest, BEVOR das
    # finally done_event setzt. Laeuft der Timeout genau dazwischen ab,
    # toetet der Waechter PowerPoint, obwohl der Aufruf gelungen ist - der
    # Aufrufer bekaeme ein Ergebnis und arbeitete danach mit einer toten
    # COM-Instanz weiter. Schloss ALLEIN genuegt nicht (nachgemessen:
    # 43 -> 30 von 300 Faellen); erst die unteilbare Pruefung auf dem
    # Erfolgspfad unten macht Toeten und Erfolg eindeutig (0 von 300).
    state_lock   = threading.Lock()
    completed    = [False]

    # Erstellungszeit SYNCHRON im Main-Thread erfassen, BEVOR der Watchdog-
    # Thread startet. Wuerde sie erst im Thread gelesen, koennte die Skript-
    # PowerPoint-Instanz in den Millisekunden zwischen Thread-Spawn und Lesen
    # abstuerzen, Windows die PID an eine neue (vom Benutzer geoeffnete)
    # PowerPoint-Sitzung vergeben - und der Watchdog merkte sich deren
    # create_time als "erwartet" und killte nach Timeout den falschen Prozess.
    expected_ct = None
    if ppt_pid:
        try:
            expected_ct = psutil.Process(ppt_pid).create_time()
        except Exception:
            expected_ct = None

    def _watchdog():
        if not done_event.wait(timeout):
            with state_lock:
                if completed[0]:
                    return          # Aufruf war bereits fertig
                timeout_flag[0] = True
            if ppt_pid:
                try:
                    proc = psutil.Process(ppt_pid)
                    if "powerpnt" not in proc.name().lower():
                        proc = None
                    if proc is not None and expected_ct is not None:
                        try:
                            if abs(proc.create_time() - expected_ct) > 0.001:
                                proc = None  # recycelt -> fremde Instanz
                        except Exception:
                            proc = None
                    if proc is not None:
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
        if restore_fn is not None:
            try:
                restore_fn()
            except Exception as _e:
                detail_logger.debug(f"_ppt_call_with_watchdog: Exception verworfen: {_e!r}")


def _verify_trust_center_for_temp(ppt_app, ppt_pid: Optional[int],
                                  timeout: float = 25.0) -> Tuple[bool, str]:
    try:
        os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)
    except Exception as e:
        return False, f"Ordner nicht erstellbar: {e}"

    test_path = os.path.join(TEMP_PROCESS_PATH, f"trust_test_{uuid.uuid4().hex}.pptx")

    # --- (1) Add + SaveAs (mit Watchdog) ---
    saveas_pres = [None]
    def _do_saveas():
        saveas_pres[0] = ppt_app.Presentations.Add(WithWindow=COM_FALSE)
        saveas_pres[0].SaveAs(test_path, FileFormat=PP_FORMAT_PPTX)
        saveas_pres[0].Close()
        saveas_pres[0] = None

    try:
        _ppt_call_with_watchdog(
            "SaveAs (Trust-Center-Test)", timeout, ppt_pid, _do_saveas
        )
    except TimeoutError:
        if saveas_pres[0] is not None:
            try: saveas_pres[0].Close()
            except Exception as _e:
                detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        try: _remove_from_mru(ppt_app, test_path)
        except Exception as _e:
            detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        if os.path.exists(test_path):
            try: os.remove(test_path)
            except Exception as _e:
                detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        return False, (f"TIMEOUT nach {timeout:.0f}s bei SaveAs - "
                       "Temp-Ordner blockiert oder PowerPoint haengt")
    except Exception as e:
        if saveas_pres[0] is not None:
            try: saveas_pres[0].Close()
            except Exception as _e:
                detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        try: _remove_from_mru(ppt_app, test_path)
        except Exception as _e:
            detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        if os.path.exists(test_path):
            try: os.remove(test_path)
            except Exception as _e:
                detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        return False, f"SaveAs fehlgeschlagen: {e}"

    if not os.path.exists(test_path):
        try: _remove_from_mru(ppt_app, test_path)
        except Exception as _e:
            detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        return False, "SaveAs hat keine Datei erzeugt"

    # --- (2) Open der gespeicherten Datei (mit Watchdog) ---
    open_pres = [None]
    def _do_open():
        open_pres[0] = ppt_app.Presentations.Open(
            test_path,
            ReadOnly=COM_TRUE,
            Untitled=COM_FALSE,
            WithWindow=COM_FALSE,
        )

    try:
        _ppt_call_with_watchdog(
            "Open (Trust-Center-Test)", timeout, ppt_pid, _do_open
        )
    except TimeoutError:
        # Watchdog hat den Prozess beendet; Datei aufraeumen und melden.
        if os.path.exists(test_path):
            try: os.remove(test_path)
            except Exception as _e:
                detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        return False, (f"TIMEOUT nach {timeout:.0f}s bei Open - "
                       "Trust Center vermutlich nicht konfiguriert oder "
                       "Geschuetzte Ansicht aktiv")
    except Exception as e:
        if open_pres[0] is not None:
            try: open_pres[0].Close()
            except Exception as _e:
                detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        try: _remove_from_mru(ppt_app, test_path)
        except Exception as _e:
            detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        if os.path.exists(test_path):
            try: os.remove(test_path)
            except Exception as _e:
                detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        return False, f"Open fehlgeschlagen: {e}"

    if open_pres[0] is None:
        try: _remove_from_mru(ppt_app, test_path)
        except Exception as _e:
            detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        if os.path.exists(test_path):
            try: os.remove(test_path)
            except Exception as _e:
                detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
        return False, "Presentations.Open lieferte kein Praesentations-Objekt"

    # --- (3) Cleanup ---
    try: open_pres[0].Close()
    except Exception as _e:
        detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
    try: _remove_from_mru(ppt_app, test_path)
    except Exception as _e:
        detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
    if os.path.exists(test_path):
        try: os.remove(test_path)
        except Exception as _e:
            detail_logger.debug(f"_verify_trust_center_for_temp: Exception verworfen: {_e!r}")
    return True, ""


# Sorgt dafuer, dass der folgende Hinweis nur EINMAL pro Lauf im Protokoll
# steht - er wiederholt sich sonst bei jeder verarbeiteten Datei.
#
# Der frueher hier stehende Name `_MRU_AVAILABLE` samt Kommentar "wird auf
# False gesetzt, sobald RecentFiles erstmals nicht erreichbar war" las sich wie
# das Ergebnis einer Faehigkeitspruefung. Eine solche Pruefung gab es nie: die
# Funktion hat noch nie etwas an ppt_app oder paths angefasst, sondern schaltet
# beim ersten Aufruf bedingungslos um. Der Name sagt jetzt, was die Variable
# wirklich bedeutet.
#
# Die Aussage im Hinweis selbst wurde nachgemessen (PowerPoint 16.0, COM):
# Application.RecentFiles gibt es dort tatsaechlich nicht - der Zugriff
# scheitert mit DISP_E_UNKNOWNNAME (0x80020006). Word und Excel kennen die
# Eigenschaft, PowerPoint nicht. Es gibt hier also nichts zu bereinigen.
_MRU_HINWEIS_GEZEIGT = False


def _remove_from_mru(ppt_app, *paths) -> None:
    # ppt_app und paths bleiben bewusst ungenutzt: die Signatur haelt die
    # Aufrufstellen zu den Schwesterskripten (3a/Word) deckungsgleich.
    global _MRU_HINWEIS_GEZEIGT
    if not _MRU_HINWEIS_GEZEIGT:
        _MRU_HINWEIS_GEZEIGT = True
        detail_logger.debug(
            "MRU-Bereinigung übersprungen: PowerPoint exponiert kein "
            "Application.RecentFiles über COM (Microsoft-Limitierung). "
            "Temporäre Pfade können später ggf. manuell aus der Registry "
            "entfernt werden: HKCU\\Software\\Microsoft\\Office\\16.0\\"
            "PowerPoint\\File MRU"
        )


def _cleanup_user_temp() -> None:
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
            return True                         # Office Owner-/Temp-Stubs
        # gen_py (win32com-Dispatch-Cache) bewusst NICHT: wird von parallel
        # laufenden COM-Skripten (3a/3b/4a) aktiv genutzt; Loeschen waehrend
        # deren Lauf kann sie stoeren.
        if low.startswith("vbe"):
            return True                         # VBA-Reste
        if low.startswith("ppt") and low.endswith(".tmp"):
            return True                         # PowerPoint-Temp (pptXXXX.tmp)
        if low.startswith("cvr") and low.endswith(".tmp"):
            return True                         # Office-Crash-Recovery-Stubs
        if low.startswith(_TEMP_PROCESS_PREFIX_LOWER):
            return True                         # eigene verwaiste Arbeitsordner
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
            # Gesperrte oder fremde Dateien → still übergehen.
            continue
        except Exception:
            continue

    detail_logger.debug(
        f"TEMP-Cleanup ({temp_dir}): {removed_files} Dateien, {removed_dirs} Ordner entfernt"
    )


def _cleanup_ppt_inetcache() -> None:
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
                        detail_logger.debug(f"_cleanup_ppt_inetcache: Exception verworfen: {_e!r}")
                    os.remove(fpath)
                    files_deleted += 1
                    bytes_freed   += fsize
                except (OSError, PermissionError) as e:
                    files_skipped += 1
                    detail_logger.debug(
                        f"INetCache: Datei gesperrt, uebersprungen: "
                        f"{fpath} ({e})")
            # Leere Unterverzeichnisse mitnehmen (best effort)
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
        if not os.path.exists(p):
            return True
        try:
            win32api.SetFileAttributes(p, win32con.FILE_ATTRIBUTE_NORMAL)
        except Exception:
            os.chmod(p, stat.S_IWRITE)
        os.remove(p)
        detail_logger.debug(f"Gelöscht: {display_path(path)}")
        return True
    except Exception as e:
        detail_logger.warning(f"Löschen fehlgeschlagen: {display_path(path)} – {e}")
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
        detail_logger.warning(f"Datei zu klein ({size} Bytes): {display_path(path)}")
        return False
    return True


def _wait_for_file_unlock(path: str,
                          max_retries: int = AV_MAX_RETRIES,
                          retry_delay: float = AV_RETRY_DELAY) -> bool:
    if not safe_exists(path):
        return False
    p = prepare_long_path(path) if len(path) > MAX_PATH_LEN else path
    for attempt in range(max_retries):
        try:
            h = win32file.CreateFile(
                p,
                win32file.GENERIC_READ,
                0,
                None,
                win32file.OPEN_EXISTING,
                0,
                None,
            )
            try:
                h.Close()
            except Exception as _e:
                detail_logger.debug(f"_wait_for_file_unlock: Exception verworfen: {_e!r}")
            if attempt > 0:
                detail_logger.debug(
                    f"Datei nach {attempt * retry_delay:.1f}s freigegeben: "
                    f"{display_path(path)}"
                )
            return True
        except Exception as e:
            detail_logger.debug(
                f"Datei gesperrt (Versuch {attempt+1}/{max_retries}): "
                f"{display_path(path)} – {e}"
            )
            time.sleep(retry_delay)
    detail_logger.warning(
        f"Datei nach {max_retries * retry_delay:.0f}s noch gesperrt "
        f"(AV-Scanner?): {display_path(path)}"
    )
    return False


def robust_copy(src: str, dst: str, max_retries: int = MAX_RETRIES) -> bool:
    for attempt in range(max_retries):
        try:
            s = prepare_long_path(src) if len(src) > MAX_PATH_LEN else src
            d = prepare_long_path(dst) if len(dst) > MAX_PATH_LEN else dst
            if os.path.exists(d):
                try:
                    win32api.SetFileAttributes(d, win32con.FILE_ATTRIBUTE_NORMAL)
                except Exception:
                    os.chmod(d, stat.S_IWRITE)
            shutil.copy2(s, d)
            time.sleep(0.5)
            if verify_file(dst):
                detail_logger.debug(f"Kopiert: {display_path(src)} → {display_path(dst)}")
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
                os.chmod(d, stat.S_IWRITE)
        os.replace(s, d)
        detail_logger.debug(f"Ersetzt (replace): {display_path(src)} → {display_path(dst)}")
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
            src_size = os.path.getsize(s)
            shutil.copy2(s, stage_long)
            time.sleep(0.5)
            if not safe_exists(stage):
                raise Exception("Staging-Datei existiert nicht nach Kopieren")
            if safe_getsize(stage) != src_size:
                raise Exception(
                    f"Größenabweichung Staging: Quelle={src_size}, "
                    f"Ziel={safe_getsize(stage)}"
                )
            if os.path.exists(d):
                try:
                    win32api.SetFileAttributes(d, win32con.FILE_ATTRIBUTE_NORMAL)
                except Exception:
                    os.chmod(d, stat.S_IWRITE)
            os.replace(stage_long, d)
            if not safe_remove(src):
                detail_logger.warning(
                    f"Verschieben: Zieldatei OK, aber Quelle nicht löschbar: "
                    f"{display_path(src)}"
                )
            detail_logger.debug(
                f"Verschoben (staging+replace): {display_path(src)} → {display_path(dst)}"
            )
            return True
        # BaseException, nicht Exception: Der Signal-Handler beendet sich mit
        # sys.exit() (SystemExit erbt von BaseException). Bei Strg+C mitten im
        # Kopieren blieb die Staging-Datei '<Ziel>.tmp_new' sonst auf der
        # Ablage liegen, und kein Aufraeumpfad erfasst sie je wieder.
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


# ==================================================================
# Dateinamen-Bereinigung (Google Drive for Desktop u. a.)
# ==================================================================

_INVISIBLE_RE = re.compile(
    r"[\x00-\x1f\x7f"
    r"\u200b\u200c\u200d\u2060"
    r"\ufeff"
    r"\u00ad"
    r"\u200e\u200f"
    r"\u202a-\u202e"
    r"\u2066-\u2069"
    r"]"
)

_INVALID_CHARS_RE = re.compile(r'[<>:"/\\|?*]')

_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def sanitize_filename(name: str) -> str:
    name = _INVISIBLE_RE.sub("", name)
    name = _INVALID_CHARS_RE.sub("-", name)
    name = name.strip().rstrip(".")
    base_token = name.split(".")[0].upper() if name else ""
    if base_token in _RESERVED_NAMES:
        name = f"_{name}"
    if not name:
        name = "_bereinigt"
    return name


def sanitize_file_on_disk(file_path: str, pbar: tqdm) -> Optional[str]:
    directory = os.path.dirname(file_path)
    basename  = os.path.basename(file_path)
    stem, ext = os.path.splitext(basename)

    clean_stem = sanitize_filename(stem)

    if clean_stem == stem:
        return None

    new_name = f"{clean_stem}{ext}"
    new_path = os.path.join(directory, new_name)

    reserved_placeholder = None
    if safe_exists(new_path) and new_path.lower() != file_path.lower():
        # Atomar reservieren (siehe reserve_unique_path) statt pruefen und
        # spaeter benutzen.
        candidate = reserve_unique_path(
            os.path.join(directory, clean_stem), ext, max_tries=999)
        if candidate is None:
            detail_logger.warning(
                f"Dateinamen-Bereinigung: kein freier Ausweichname für "
                f"{display_path(file_path)}")
            return None
        reserved_placeholder = candidate
        new_path = candidate
        new_name = os.path.basename(candidate)

    try:
        src = prepare_long_path(file_path) if len(file_path) > MAX_PATH_LEN else file_path
        dst = prepare_long_path(new_path)  if len(new_path)  > MAX_PATH_LEN else new_path
        # os.replace statt os.rename: liegt am Ziel unser eigener 0-Byte-
        # Platzhalter, wuerde os.rename unter Windows scheitern.
        os.replace(src, dst)
        reserved_placeholder = None
        pbar.write(f"  📝 Dateiname bereinigt: '{basename}' → '{new_name}'")
        detail_logger.info(
            f"Dateiname bereinigt: {display_path(file_path)} → {display_path(new_path)}"
        )
        file_logger.info(
            f"Dateiname bereinigt: '{basename}' → '{new_name}' in {display_path(directory)}"
        )
        return new_path
    except Exception as e:
        detail_logger.warning(
            f"Dateinamen-Bereinigung fehlgeschlagen: {display_path(file_path)} – {e}"
        )
        return None
    finally:
        release_unique_path(reserved_placeholder)


# ==================================================================
# Pfad-Hilfsfunktionen (User-Shell-Folders + sanitize_path + 8-Wege-Abfrage)
# ==================================================================

def sanitize_path(raw: str) -> str:
    path = raw.strip().strip('"').strip("'")
    if len(path) == 2 and path[1] == ":":
        path = path + "\\"
    return path


def _get_user_shell_folder(value_name: str, fallback_subdir: str) -> str:
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        ) as key:
            value, _ = winreg.QueryValueEx(key, value_name)
            expanded = os.path.expandvars(value)
            if os.path.isdir(expanded):
                return expanded
    except Exception as e:
        detail_logger.debug(
            f"Registry-Lookup für '{value_name}' fehlgeschlagen, "
            f"nutze Fallback: {e}"
        )
    return os.path.join(os.path.expanduser("~"), fallback_subdir)


def _get_desktop_dir() -> str:
    return _get_user_shell_folder("Desktop", "Desktop")


def _get_downloads_dir() -> str:
    return _get_user_shell_folder(
        "{374DE290-123F-4565-9164-39C4925E467B}", "Downloads"
    )


def ask_directory() -> str:
    desktop_dir   = _get_desktop_dir()
    downloads_dir = _get_downloads_dir()

    print("\nZielverzeichnis auswählen:")
    print("  [1] Q:\\")
    print("  [2] R:\\")
    print("  [3] G:\\Geteilte Ablagen")
    print("  [4] G:\\Meine Ablage")
    print("  [5] \\\\server\\dfs")
    print(f"  [6] Desktop des Benutzers   ({desktop_dir})")
    print(f"  [7] Downloads des Benutzers ({downloads_dir})")
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
            print("\n❌ Fehler: Kein interaktives Terminal. Bitte --dir und --auto-start verwenden.")
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
            if os.path.isdir(prepare_long_path(desktop_dir)):
                return desktop_dir
            print(f"  ❌ Desktop-Verzeichnis nicht gefunden: '{desktop_dir}'")
        elif choice == "7":
            if os.path.isdir(prepare_long_path(downloads_dir)):
                return downloads_dir
            print(f"  ❌ Downloads-Verzeichnis nicht gefunden: '{downloads_dir}'")
        elif choice == "8":
            try:
                raw = input("Pfad eingeben: ")
            except EOFError:
                print("\n❌ Fehler: Kein interaktives Terminal. Bitte --dir und --auto-start verwenden.")
                sys.exit(1)
            path = sanitize_path(raw)
            if os.path.isdir(prepare_long_path(path)):
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
# AppVersion-Prüfung  (ZIP-Ebene, ohne COM)
# ==================================================================

def get_app_version(file_path: str) -> Optional[float]:
    ext = os.path.splitext(file_path)[1].lower()
    xml_formats = {".pptx", ".pptm", ".ppsx", ".ppsm", ".potx", ".potm", ".ppam"}
    if ext not in xml_formats:
        return None

    p = prepare_long_path(file_path) if len(file_path) > MAX_PATH_LEN else file_path

    try:
        with zipfile.ZipFile(p, "r") as z:
            if "docProps/app.xml" not in z.namelist():
                return None
            with z.open("docProps/app.xml") as f:
                content = f.read(65536).decode("utf-8", errors="ignore")
                match = re.search(r"<(?:\w+:)?AppVersion>([0-9]+(?:\.[0-9]+)?)</(?:\w+:)?AppVersion>",
                                  content)
                if match:
                    return float(match.group(1))
    except Exception as e:
        detail_logger.debug(
            f"AppVersion-Lesen fehlgeschlagen ({display_path(file_path)}): {e}"
        )

    return None


def get_conversion_info(file_path: str, pres_format: Optional[int]) -> str:
    ext = os.path.splitext(file_path)[1].lower()

    if ext in (".ppt", ".pps", ".pot"):
        return "Altes Binärformat → Formatkonvertierung"

    if pres_format in PP_OLD_FORMAT_CODES:
        return f"Altes FileFormat ({pres_format}) → Formatkonvertierung"

    app_version = get_app_version(file_path)
    if app_version is None:
        return "Keine AppVersion (Drittanbieter/beschädigt) → Neuserialierung"
    if app_version < TARGET_APP_VERSION:
        return f"AppVersion {app_version} < {TARGET_APP_VERSION} → Neuserialierung"

    return f"AppVersion {app_version} → Neuserialierung durch aktuelle Engine"


# ==================================================================
# Zielformat bestimmen (Format + Dateiendung)
# ==================================================================

def get_target_format(ext: str, has_macros: bool) -> Tuple[int, str]:
    addins     = {".ppam"}
    templates  = {".pot", ".potx", ".potm"}
    slideshows = {".pps", ".ppsx", ".ppsm"}

    if ext in addins:
        return PP_FORMAT_PPAM, ".ppam"
    elif ext in templates:
        if has_macros:
            return PP_FORMAT_POTM, ".potm"
        return PP_FORMAT_POTX, ".potx"
    elif ext in slideshows:
        if has_macros:
            return PP_FORMAT_PPSM, ".ppsm"
        return PP_FORMAT_PPSX, ".ppsx"
    else:
        if has_macros:
            return PP_FORMAT_PPTM, ".pptm"
        return PP_FORMAT_PPTX, ".pptx"


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
    # Schneller Lock-Test fuer ~$-/._-Stub-Dateien: laesst sich die Datei
    # NICHT exklusiv oeffnen, ist sie aktiv (zugehoerige Datei gerade offen)
    # und darf NICHT geloescht werden. Nur verwaiste Reste werden bereinigt.
    p = prepare_long_path(path) if len(path) > MAX_PATH_LEN else path
    try:
        h = win32file.CreateFile(
            p, win32file.GENERIC_READ, 0, None,
            win32file.OPEN_EXISTING, 0, None)
        try:
            h.Close()
        except Exception as _e:
            detail_logger.debug(f"_temp_file_is_active: Exception verworfen: {_e!r}")
        return False
    except Exception:
        # Sharing-/Lock-Violation oder anderer Zugriffsfehler -> als aktiv
        # behandeln und in Ruhe lassen.
        return True


def _path_matches_exclude(root_lower: str, exclude_patterns: list) -> bool:
    # Grenzanker-Matching auf Ordnerebene: Das Muster muss eine komplette
    # Pfadkomponente (oder zusammenhaengende Komponentenfolge) sein. Ein
    # reines Substring-Matching ('pat in root') wuerde quer ueber
    # Komponentengrenzen treffen (z.B. '--exclude-dir alt' wuerde
    # 'Verwaltung' ausschliessen). Durch Einrahmen mit os.sep matcht 'alt'
    # nur eine echte Komponente; Mehr-Segment-Muster ('Archiv\\Alt') gehen.
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
    """PowerPoint-Dateien liefern und optional verwaiste Office-Reste raeumen.

    clean_temp=False raeumt NICHTS - zwingend fuer den Probelauf, der
    zusichert, das Dateisystem unangetastet zu lassen.

    clean_appledouble steuert die '._*'-Dateien getrennt und ist bewusst
    standardmaessig aus: Auf Freigaben mit Mac-Clients sind das
    AppleDouble-Container (Finder-Metadaten, Resource-Forks) und keine
    Office-Reste. Sie wurden ohne Endungsfilter im gesamten Baum geloescht,
    also auch '._Foto.jpg' - Dateien, mit denen dieses Skript nichts zu tun
    hat.
    """
    extensions = {
        ".ppt",  ".pptx", ".pptm", ".ppam",
        ".pps",  ".ppsx", ".ppsm",
        ".pot",  ".potx", ".potm",
    }
    exclude_patterns = [p.lower() for p in (exclude_patterns or [])]

    def _walk_error(err: OSError) -> None:
        detail_logger.warning(
            f"Verzeichnis nicht lesbar (fehlende Rechte?) – übersprungen: "
            f"{display_path(err.filename) if err.filename else err}"
        )

    directory = prepare_long_path(directory)
    # Eigenen Arbeitsordner ausnehmen: liegt das Ziel z.B. auf Documents,
    # wuerde der Generator sonst die eigenen convert_*/longpath_*-Dateien
    # einsammeln, waehrend PowerPoint sie gerade schreibt.
    temp_nc = os.path.normcase(os.path.abspath(TEMP_PROCESS_PATH))

    for root, dirs, files in os.walk(directory, onerror=_walk_error):
        # Per --exclude-dir ausgeschlossene Teilpfade gar nicht erst betreten.
        if _path_matches_exclude(strip_com_path(root).lower(), exclude_patterns):
            dirs[:] = []
            continue
        # Papierkorb/Systemordner/NAS-Snapshots, eigenen Temp-Ordner und
        # Junctions in-place aus der Traversierung entfernen.
        pruned = []
        for d in dirs:
            if d.lower() in EXCLUDE_DIR_NAMES:
                continue
            full_d = os.path.join(root, d)
            d_nc = os.path.normcase(os.path.abspath(strip_com_path(full_d)))
            if d_nc == temp_nc or d_nc.startswith(temp_nc + os.sep):
                continue
            if _is_reparse_point(full_d):
                continue
            pruned.append(d)
        dirs[:] = pruned

        if clean_temp:
            for f in files:
                if f.startswith("~$") or (clean_appledouble and f.startswith("._")):
                    full_t = os.path.join(root, f)
                    try:
                        # Aktive Stub-Dateien (zugehoerige Datei offen) NICHT
                        # loeschen - nur verwaiste Reste bereinigen.
                        if _temp_file_is_active(full_t):
                            continue
                        safe_remove(full_t)
                    except Exception as _e:
                        detail_logger.debug(f"file_generator: Exception verworfen: {_e!r}")

        for f in files:
            nl = f.lower()
            if nl.startswith("~$") or nl.startswith("._"):
                continue
            if any(nl.endswith(ext) for ext in extensions):
                yield strip_com_path(os.path.join(root, f))


# ==================================================================
# Final-Status entfernen
# ==================================================================

def remove_protection(pres: win32com.client.CDispatch,
                      file_path_display: str) -> bool:
    changed = False
    try:
        if pres.Final:
            pres.Final = COM_FALSE
            changed = True
            detail_logger.info(f"Final-Status entfernt: {display_path(file_path_display)}")
    except Exception as e:
        detail_logger.debug(f"Final-Check fehlgeschlagen: {e}")
    return changed


# ==================================================================
# Eindeutigen Zielpfad finden (bei Kollision)
# ==================================================================

def _resolve_unique_path(target_path: str) -> str:
    if not safe_exists(target_path):
        return target_path
    base, ext_part = os.path.splitext(target_path)
    # Atomar reservieren statt pruefen-und-spaeter-benutzen (siehe
    # reserve_unique_path). Der Platzhalter wird vom spaeteren Schreiben
    # ueberschrieben bzw. von release_unique_path wieder entfernt.
    reserved = reserve_unique_path(base, ext_part, max_tries=999)
    if reserved:
        return reserved
    raise Exception(f"Kein freier Zielname findbar für: {display_path(target_path)}")


# ==================================================================
# Kern-Logik: Konvertierung / Aktualisierung
# ==================================================================

def convert_ppt_file(
    file_path: str,
    ppt_app: win32com.client.CDispatch,
    pbar: tqdm,
    skip_password: bool,
    dry_run: bool = False,
) -> str:
    pbar.write(f"Prüfe: {os.path.basename(file_path)}")
    detail_logger.info(f"=== Starte: {display_path(file_path)} ===")

    original_path  = file_path
    is_temp_copy   = False
    backup_path    = None
    # Beiseitegelegtes veraltetes Backup eines frueheren Laufs (siehe unten).
    # Wird bei erfolgreichem Abschluss zusammen mit backup_path geloescht,
    # damit sich solche Reste nicht dauerhaft auf der Ablage anhaeufen.
    stale_bak      = None
    temp_save_path = None
    target_path    = None
    pres           = None
    # Reservierter Zielname (0-Byte-Platzhalter) - siehe Freigabe im finally.
    target_was_reserved = False
    converted_ok        = False
    ext            = os.path.splitext(file_path)[1].lower()
    update_complete = False
    # Die Sicherung ist NACHWEISLICH gelungen. Nur dann darf der
    # Fehlerpfad aus backup_path zurueckschreiben: scheitert robust_copy
    # mitten in der Kopie, liegt am Backup-Pfad ein Fragment (shutil.copy2
    # trunkiert das Ziel sofort), und verify_file prueft nur >= 100 Byte -
    # ein solches Fragment wuerde ungeprueft ueber das Original laufen.
    backup_ok      = False
    # Die Zieldatei wurde in diesem Lauf angefasst (robust_move begonnen).
    # Ist sie es nicht, ist target_path das unveraenderte Original und darf
    # im Fehlerpfad weder ueberschrieben noch entfernt werden.
    target_touched = False

    # PowerPoint-Add-Ins (.ppam) lassen sich NICHT via Presentations.Open
    # als Praesentation oeffnen (PowerPoint blockiert das strikt) und auch
    # nicht via SaveAs aktualisieren. Frueher wurde der Open trotzdem
    # versucht und der com_error abgefangen - das kostete pro Datei einen
    # unnoetigen (potenziell haengenden) COM-Open. Daher hier sofort und
    # ohne COM-Eingriff ueberspringen; bleibt als SKIPPED in der Statistik
    # sichtbar.
    if ext == ".ppam":
        pbar.write("  → ÜBERSPRUNGEN: PowerPoint-Add-In (.ppam) – nicht via SaveAs aktualisierbar.")
        detail_logger.info(
            f"Add-In übersprungen ohne COM-Open (.ppam): {display_path(original_path)}")
        return "SKIPPED"

    mru_paths_to_clean = []

    # --- Long-Path-Behandlung ---
    if len(file_path) > MAX_PATH_LEN:
        is_temp_copy = True
        try:
            temp_name = f"longpath_{uuid.uuid4().hex}{ext}"
            temp_path = os.path.join(TEMP_PROCESS_PATH, temp_name)
            if not robust_copy(file_path, temp_path):
                raise Exception("Kopieren fehlgeschlagen")
            _wait_for_file_unlock(temp_path)
            if not verify_file(temp_path):
                raise Exception("Temp-Kopie leer/ungültig")
            file_path = temp_path
            detail_logger.debug(f"Temp-Kopie: {display_path(temp_path)}")
        except Exception as e:
            log_error(original_path, Exception(f"Long-Path-Kopie fehlgeschlagen: {e}"))
            return "ERROR"

    orig_times = None
    try:
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
    except Exception:
        orig_times = None

    # NTFS-Sicherheitsinfo (Owner/Group/DACL) des Originals sichern -
    # wird nach der Ersetzung auf die neue Datei uebertragen.
    orig_sd = _get_security_descriptor(original_path)

    try:
        # --- Verschluesselungs-Check (vor COM-Open) ---
        is_encrypted = False
        if ext in (".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm"):
            try:
                with open(file_path, "rb") as f_check:
                    office_file = msoffcrypto.OfficeFile(f_check)
                    is_encrypted = office_file.is_encrypted()
            except Exception as e_crypt:
                detail_logger.warning(
                    f"msoffcrypto-Check fehlgeschlagen für {original_path}: {e_crypt}")
                try:
                    if not zipfile.is_zipfile(file_path):
                        is_encrypted = True
                    else:
                        with zipfile.ZipFile(file_path) as zf:
                            names = zf.namelist()
                            if "EncryptedPackage" in names or "EncryptionInfo" in names:
                                is_encrypted = True
                            elif "ppt/presentation.xml" not in names and \
                                 "[Content_Types].xml" in names:
                                is_encrypted = True
                            else:
                                is_encrypted = False
                except Exception as e_zip:
                    detail_logger.warning(
                        f"ZIP-Fallback ebenfalls fehlgeschlagen: {e_zip}")
                    is_encrypted = True
        elif ext in (".ppt", ".pps", ".pot"):
            try:
                with open(file_path, "rb") as f_ole:
                    ole_header = f_ole.read(8)
                    if ole_header == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                        content = f_ole.read(4088)
                        if b'EncryptedPackage' in content or b'EncryptionInfo' in content:
                            is_encrypted = True
            except Exception as e_ole:
                detail_logger.debug(f"OLE-Header-Check fehlgeschlagen: {e_ole}")

        if is_encrypted:
            if skip_password:
                pbar.write("  → ÜBERSPRUNGEN: Passwortgeschützt.")
                detail_logger.warning(
                    f"Übersprungen (Passwort): {display_path(original_path)}"
                )
                return "SKIPPED"
            log_error(
                original_path,
                Exception("Passwortgeschützt – kein Zugriff (PowerPoint COM unterstützt kein Passwort)")
            )
            return "ERROR"

        # --- Datei öffnen (mit Retry, Passwort-Abbruch) ---
        com_path    = strip_com_path(file_path)
        mru_paths_to_clean.append(com_path)

        try:
            p_attr = prepare_long_path(file_path) if len(file_path) > MAX_PATH_LEN else file_path
            attrs  = win32api.GetFileAttributes(p_attr)
            if attrs & win32con.FILE_ATTRIBUTE_READONLY:
                win32api.SetFileAttributes(
                    p_attr, attrs & ~win32con.FILE_ATTRIBUTE_READONLY
                )
                detail_logger.debug(
                    f"Read-Only-Attribut entfernt: {display_path(file_path)}"
                )
        except Exception as e:
            detail_logger.debug(f"Read-Only-Check fehlgeschlagen: {e}")

        pres_opened = False
        for open_attempt in range(MAX_RETRIES):
            try:
                pres_holder = [None]
                def _do_open():
                    # Im Probelauf ZWINGEND schreibgeschuetzt oeffnen.
                    # Nachgemessen an einer .ppt: PowerPoint schreibt eine
                    # schreibend geoeffnete Datei im Altformat schon beim
                    # Oeffnen neu auf die Platte - 258560 -> 260608 Bytes,
                    # bei jedem Lauf ein anderer Hash. Weder ein blosses
                    # Close() noch Saved=True verhindern das (beides
                    # gemessen); nur ReadOnly=True tut es. Der Zeitstempel
                    # wird spaeter wiederhergestellt, die Aenderung war
                    # deshalb unsichtbar - der Probelauf hat seine Zusage
                    # "es wurde NICHTS geaendert" gebrochen.
                    pres_holder[0] = ppt_app.Presentations.Open(
                        com_path,
                        ReadOnly   = COM_TRUE if dry_run else COM_FALSE,
                        Untitled   = COM_FALSE,
                        WithWindow = COM_FALSE,
                    )
                _ppt_call_with_watchdog(
                    "Open", PPT_OPEN_TIMEOUT, ppt_app_pid, _do_open
                )
                pres = pres_holder[0]
                pres_opened = True
                detail_logger.debug("Geöffnet")
                break
            except TimeoutError:
                pbar.write(f"  ✕  TIMEOUT bei Open ({PPT_OPEN_TIMEOUT:.0f}s) – PowerPoint blockiert")
                log_error(original_path, Exception(
                    f"Timeout bei Presentations.Open nach {PPT_OPEN_TIMEOUT:.0f}s "
                    f"(vermutlich modaler Dialog: Reparatur/Verknuepfung/Passwort)"
                ))
                return "ERROR"
            except pythoncom.com_error as e:
                detail_logger.debug(f"Öffnen Versuch {open_attempt+1}: {e}")
                if ext == ".ppam":
                    pbar.write("  → ÜBERSPRUNGEN: PowerPoint-Add-In (.ppam) – nicht via SaveAs aktualisierbar.")
                    detail_logger.warning(
                        f"Add-In übersprungen (.ppam kann nicht als Praesentation "
                        f"geoeffnet/konvertiert werden): {display_path(original_path)}"
                    )
                    return "SKIPPED"
                is_pwd_error = (
                    (hasattr(e, "hresult") and e.hresult == -2147024809)
                    or "password" in str(e).lower()
                    or "kennwort" in str(e).lower()
                )
                if is_pwd_error:
                    if skip_password:
                        pbar.write("  → ÜBERSPRUNGEN: Passwortgeschützt.")
                        detail_logger.warning(
                            f"Übersprungen (Passwort): {display_path(original_path)}"
                        )
                        return "SKIPPED"
                    log_error(
                        original_path,
                        Exception(f"Passwortgeschützt – kein Zugriff: {e}")
                    )
                    return "ERROR"
                if open_attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY)

        if not pres_opened or pres is None:
            if skip_password:
                # Vor dem SKIPPED pruefen, ob die COM-Instanz ueberhaupt noch
                # lebt. Stuerzt PowerPoint zwischen zwei Dateien ab - womit der
                # Code ausdruecklich rechnet -, scheitert Presentations.Open in
                # allen Versuchen, und mit skip_password ging das als 'SKIPPED'
                # zurueck. Der Wiederanlauf der Engine haengt aber allein an
                # 'ERROR': er lief nie an, und JEDE weitere Datei scheiterte
                # gegen dieselbe tote Instanz.
                try:
                    _ = ppt_app.Version
                except Exception:
                    pbar.write("  ✗  FEHLER: PowerPoint reagiert nicht mehr.")
                    detail_logger.warning(
                        f"COM-Instanz tot beim Oeffnen: {display_path(original_path)}"
                    )
                    log_error(original_path,
                              Exception("PowerPoint-Instanz nicht mehr erreichbar"))
                    return "ERROR"
                pbar.write("  → ÜBERSPRUNGEN: Zugriff verweigert.")
                detail_logger.warning(
                    f"Übersprungen (kein Zugriff): {display_path(original_path)}"
                )
                return "SKIPPED"
            else:
                log_error(original_path, Exception("Öffnen fehlgeschlagen (Defekt?)"))
                return "ERROR"

        pres_format = None
        try:
            pres_format = pres.FileFormat
            detail_logger.debug(f"FileFormat: {pres_format}")
        except Exception as _e:
            detail_logger.debug(f"convert_ppt_file: Exception verworfen: {_e!r}")

        # Im Probelauf entfaellt das Aufheben des Schutzes: die Praesentation
        # ist dort schreibgeschuetzt geoeffnet (siehe _do_open), und gespeichert
        # wird ohnehin nicht. Das Ergebnis des Probelaufs haengt nicht daran -
        # er meldet fuer jede geoeffnete Datei WOULD_UPDATE.
        protection_removed = False
        if not dry_run:
            protection_removed = remove_protection(pres, original_path)

        reason = get_conversion_info(original_path, pres_format)
        pbar.write(f"  → {reason}")
        if protection_removed:
            pbar.write("  → Final-Status entfernt")

        has_macros = False
        if ext in (".pptm", ".ppsm", ".potm", ".ppam"):
            has_macros = True
        else:
            try:
                has_macros = pres.HasVBProject
            except Exception:
                has_macros = True
                detail_logger.warning(
                    f"HasVBProject nicht lesbar (Trust Center?) – "
                    f"Makros werden sicherheitshalber angenommen: "
                    f"{display_path(original_path)}"
                )

        new_format, new_ext = get_target_format(ext, has_macros)

        # --- DRY-RUN: hier ist Schluss ---------------------------------
        # Ab hier wuerde gespeichert, verschoben und ggf. das Original
        # geloescht. Anders als Word/Excel kennt PowerPoint keinen
        # Kompatibilitaetsmodus: jede geoeffnete Datei wuerde neu
        # serialisiert, daher gibt es hier kein ALREADY_CURRENT.
        #
        # Die Datei ist im Probelauf schreibgeschuetzt geoeffnet (siehe
        # _do_open) - erst das macht die Zusage "es wurde NICHTS geaendert"
        # wahr. Frueher stand hier die Annahme, ein blosses Close() verwerfe
        # alles, weil die Aenderungen "nur im Arbeitsspeicher" bestuenden.
        if dry_run:
            pbar.write(f"  🔎 [PROBELAUF] Würde neu serialisiert → {new_ext}")
            detail_logger.info(
                f"[DRY-RUN] WOULD_UPDATE: {display_path(original_path)} → {new_ext}")
            try:
                pres.Close()
            except Exception as _e:
                detail_logger.debug(f"convert_ppt_file: Exception verworfen: {_e!r}")
            pres = None
            return "WOULD_UPDATE"

        temp_save_name = f"convert_{uuid.uuid4().hex}{new_ext}"
        temp_save_path = os.path.join(TEMP_PROCESS_PATH, temp_save_name)
        mru_paths_to_clean.append(temp_save_path)

        try:
            pres.Password      = ""
            pres.WritePassword = ""
        except Exception as _e:
            detail_logger.debug(f"convert_ppt_file: Exception verworfen: {_e!r}")
        try:
            _ppt_call_with_watchdog(
                "SaveAs",
                PPT_SAVEAS_TIMEOUT,
                ppt_app_pid,
                lambda: pres.SaveAs(temp_save_path, FileFormat=new_format)
            )
            detail_logger.debug(f"SaveAs → {display_path(temp_save_path)}")
        except TimeoutError:
            pbar.write(f"  ✕  TIMEOUT bei SaveAs ({PPT_SAVEAS_TIMEOUT:.0f}s) – PowerPoint blockiert")
            log_error(original_path, Exception(
                f"Timeout bei SaveAs nach {PPT_SAVEAS_TIMEOUT:.0f}s "
                f"(vermutlich modaler Dialog beim Speichern)"
            ))
            pres = None
            return "ERROR"
        except Exception as e_save:
            error_msg = str(e_save).lower()
            if ("macro" in error_msg or "makro" in error_msg) and not has_macros:
                new_format, new_ext = get_target_format(ext, has_macros=True)
                safe_remove(temp_save_path)
                temp_save_path = os.path.splitext(temp_save_path)[0] + new_ext
                mru_paths_to_clean.append(temp_save_path)
                try:
                    _ppt_call_with_watchdog(
                        "SaveAs (Makro-Fallback)",
                        PPT_SAVEAS_TIMEOUT,
                        ppt_app_pid,
                        lambda: pres.SaveAs(temp_save_path, FileFormat=new_format)
                    )
                    detail_logger.debug(f"SaveAs (Makro-Fallback) → {display_path(temp_save_path)}")
                except TimeoutError:
                    pbar.write(f"  ✕  TIMEOUT bei SaveAs-Makro-Fallback ({PPT_SAVEAS_TIMEOUT:.0f}s)")
                    log_error(original_path, Exception(
                        f"Timeout bei SaveAs-Makro-Fallback nach {PPT_SAVEAS_TIMEOUT:.0f}s"
                    ))
                    pres = None
                    return "ERROR"
            else:
                raise

        pres.Close()
        pres = None
        _remove_from_mru(ppt_app, *mru_paths_to_clean)
        time.sleep(0.3)

        _wait_for_file_unlock(temp_save_path)

        if not verify_file(temp_save_path):
            raise Exception("Verifizierung der konvertierten Datei fehlgeschlagen")

        # --- Zielpfad bestimmen (bei Formatänderung ggf. eindeutigen Namen) ---
        if ext != new_ext:
            target_path = os.path.splitext(original_path)[0] + new_ext
            if safe_exists(target_path):
                new_target = _resolve_unique_path(target_path)
                target_was_reserved = (new_target != target_path)
                pbar.write(
                    f"  ⚠  Ziel existiert bereits – speichere als "
                    f"'{os.path.basename(new_target)}'"
                )
                target_path = new_target
        else:
            target_path = original_path

        # --- Backup der Zieldatei (nur wenn Ziel bereits existiert) ---
        # Ein frisch reservierter Platzhalter ist 0 Byte gross und braucht
        # kein Backup.
        if safe_exists(target_path) and not target_was_reserved:
            backup_path = f"{target_path}.bak"
            if len(backup_path) > MAX_PATH_LEN:
                backup_path = prepare_long_path(backup_path)
            if safe_exists(backup_path):
                # Ein zurueckgebliebenes .bak stammt aus einem harten
                # Abbruch eines frueheren Laufs und kann die einzige
                # intakte Kopie der Zieldatei sein (Crash mitten im
                # Cross-Volume-Move trunkiert das Ziel). Daher beiseite
                # legen statt loeschen - und nach erfolgreichem Abschluss
                # dieses Laufs (Zieldatei ist dann wieder valide) raeumen.
                stale_bak = f"{target_path}.bak_{uuid.uuid4().hex[:6]}"
                try:
                    os.replace(
                        prepare_long_path(backup_path) if len(backup_path) > MAX_PATH_LEN else backup_path,
                        prepare_long_path(stale_bak) if len(stale_bak) > MAX_PATH_LEN else stale_bak,
                    )
                    detail_logger.warning(
                        f"Veraltetes Backup beiseite gelegt (nicht gelöscht): "
                        f"{display_path(stale_bak)}"
                    )
                except Exception:
                    safe_remove(backup_path)
                    detail_logger.warning(
                        f"Veraltetes Backup gelöscht: {display_path(backup_path)}"
                    )
            if not robust_copy(target_path, backup_path):
                # Am Backup-Pfad kann ein Torso liegen (Kopie mittendrin
                # abgerissen). Er darf den Fehlerpfad unten nicht als
                # Rueckfallquelle erreichen - entfernen und backup_path
                # zuruecksetzen, BEVOR die Ausnahme geworfen wird.
                bad_backup = backup_path
                backup_path = None
                safe_remove(bad_backup)
                raise Exception(
                    f"Backup fehlgeschlagen: {display_path(bad_backup)}. "
                    "Abbruch zum Schutz der Zieldatei."
                )
            backup_ok = True

        if is_temp_copy and safe_exists(file_path):
            safe_remove(file_path)
            is_temp_copy = False

        _wait_for_file_unlock(temp_save_path)

        # Ab hier kann die Zieldatei veraendert sein - auch wenn robust_move
        # scheitert (Cross-Volume-Rueckfall schreibt in Etappen).
        target_touched = True
        if not robust_move(temp_save_path, target_path):
            raise Exception("Verschieben der konvertierten Datei fehlgeschlagen")
        # Ab hier steht die echte Datei am Zielpfad - der reservierte
        # Platzhalter ist damit ersetzt und darf nicht freigegeben werden.
        converted_ok = True

        # ACL/Owner des Originals auf die neue Datei uebertragen
        # (Admin-Kontext: vollstaendig; Nutzer-Kontext: DACL).
        _apply_security_descriptor(target_path, orig_sd)

        if original_path.lower() != target_path.lower():
            if not safe_remove(original_path):
                safe_remove(target_path)
                raise Exception(
                    "Löschen der Originaldatei fehlgeschlagen (Rechte?). "
                    "Konvertierung rückgängig gemacht."
                )

        update_complete = True

        if backup_path and safe_exists(backup_path):
            safe_remove(backup_path)
            backup_path = None

        # Veraltetes Backup eines frueheren Laufs ist nach erfolgreichem
        # Abschluss obsolet (die Zieldatei ist jetzt wieder valide) - sonst
        # bliebe es als Datenmuell dauerhaft auf der Ablage liegen.
        if stale_bak and safe_exists(stale_bak):
            if safe_remove(stale_bak):
                detail_logger.debug(
                    f"Veraltetes Backup nach Erfolg entfernt: {display_path(stale_bak)}")
            stale_bak = None

        # --- Zeitstempel wiederherstellen ---
        if orig_times and safe_exists(target_path):
            FILE_WRITE_ATTRIBUTES = 0x0100
            for av_retry in range(5):
                try:
                    h_dst = win32file.CreateFile(
                        prepare_long_path(target_path),
                        FILE_WRITE_ATTRIBUTES,
                        win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | 4,
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

        # Nachweisliste: Pfadwechsel (z.B. .ppt -> .pptx, oder Ausweichname
        # bei Kollision) in die Konvertierungs-CSV.
        if original_path.lower() != target_path.lower():
            log_conversion(original_path, target_path)

        pbar.write(f"  ✓  Aktualisiert: '{os.path.basename(original_path)}'")
        detail_logger.info(f"Erfolgreich: {display_path(original_path)}")
        return "UPDATED"

    except Exception as e:
        pbar.write(f"  ✗  FEHLER: '{os.path.basename(original_path)}'")

        # Nur eine nachweislich gelungene Sicherung ist eine Rueckfallquelle.
        have_backup = bool(backup_ok and backup_path and safe_exists(backup_path))

        if have_backup and target_path and target_touched:
            if robust_copy(backup_path, target_path):
                _apply_security_descriptor(target_path, orig_sd)
                safe_remove(backup_path)
                pbar.write("    → Backup wiederhergestellt")
            else:
                pbar.write(
                    "    ⚠ KRITISCH: Rollback fehlgeschlagen! "
                    "Backup-Datei aus Sicherheitsgründen bewahrt."
                )
        elif have_backup and target_path:
            # Die Zieldatei wurde nie angefasst - sie IST das Original.
            # Zurueckkopieren waere ein Risiko ohne jeden Nutzen.
            safe_remove(backup_path)
            pbar.write("    → Zieldatei unverändert – kein Rollback nötig")
        elif target_touched and target_path and safe_exists(target_path):
            if update_complete:
                pbar.write(
                    "    ⚠  Spaeter Fehler nach erfolgreicher Konvertierung – "
                    "Zieldatei wird BEHALTEN."
                )
                detail_logger.warning(
                    f"Spaete Exception nach update_complete=True - "
                    f"target_path bleibt erhalten: {display_path(target_path)} ({e})"
                )
            else:
                safe_remove(target_path)
                pbar.write("    → Korruptes Dateifragment nach Abbruch sicher entfernt.")

        raise

    finally:
        # Reservierten Zielnamen freigeben, falls die Konvertierung nicht
        # bis zum Schreiben kam. Entfernt ausschliesslich 0-Byte-Dateien.
        if target_was_reserved and not converted_ok:
            release_unique_path(target_path)
        if pres is not None:
            try:
                pres.Close()
            except Exception as _e:
                detail_logger.debug(f"convert_ppt_file: Exception verworfen: {_e!r}")
            pres = None
        try:
            _remove_from_mru(ppt_app, *mru_paths_to_clean)
        except Exception as _e:
            detail_logger.debug(f"convert_ppt_file: Exception verworfen: {_e!r}")
        if is_temp_copy and safe_exists(file_path):
            safe_remove(file_path)
        if temp_save_path and safe_exists(temp_save_path):
            safe_remove(temp_save_path)


# ==================================================================
# Wrapper mit Retry (für COM-Fehler "PowerPoint beschäftigt")
# ==================================================================

def process_file_with_retries(
    full_path: str,
    ppt_app: win32com.client.CDispatch,
    pbar: tqdm,
    skip_password: bool,
    dry_run: bool = False,
) -> str:
    result = "ERROR"
    for attempt in range(MAX_RETRIES):
        try:
            result = convert_ppt_file(
                full_path, ppt_app, pbar, skip_password, dry_run=dry_run)
            break
        except pythoncom.com_error as e:
            if e.hresult in (
                -2147418111, -2147023174, -2147023170, -2146823683,
                -2147417848, -2147352567, -2146823146, -2146823136
            ) and attempt < MAX_RETRIES - 1:
                pbar.write(f"  → PowerPoint beschäftigt oder abgestürzt, warte {RETRY_DELAY}s ...")
                time.sleep(RETRY_DELAY)
            else:
                log_error(full_path, e)
                break
        except Exception as e:
            if attempt < MAX_RETRIES - 1:
                pbar.write(f"  → Dateisystem-Fehler, warte {RETRY_DELAY}s ...")
                time.sleep(RETRY_DELAY)
            else:
                log_error(full_path, e)
                break
    return result


# ==================================================================
# Verzeichnis-Verarbeitung
# ==================================================================

def process_directory(
    directory: str,
    skip_password: bool,
    count_files_first: bool,
    show_progress: bool = True,
    resume_path: Optional[str] = None,
    resume_set: Optional[set] = None,
    exclude_patterns: list = None,
    dry_run: bool = False,
) -> Tuple[dict, bool]:
    global ppt_app_global

    os.makedirs(TEMP_PROCESS_PATH, exist_ok=True)

    resume_set = resume_set or set()
    # clean_temp=not dry_run: Der Probelauf sichert zu, das Dateisystem
    # unangetastet zu lassen. Mit dem frueher fest verdrahteten
    # clean_temp=True loeschte gerade der Modus, den man vor dem ersten
    # Echtlauf auf einer fremden Ablage waehlt, bereits Dateien im ganzen
    # Baum. Gleiche Ursache und gleiche Korrektur wie in 3b.
    gen = file_generator(directory, clean_temp=not dry_run,
                         exclude_patterns=exclude_patterns)
    if resume_set:
        _src_gen = gen
        gen = (f for f in _src_gen if f.lower() not in resume_set)

    if count_files_first:
        print("\nIndiziere und bereinige Dateien (kann einige Minuten dauern) ...")
        files_list = list(gen)
        total      = len(files_list)
        iterable   = files_list
        bar_fmt    = "{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]"
        print(f"{total} Dateien gefunden"
              + (f" ({len(resume_set)} via Resume übersprungen)" if resume_set else "")
              + ". Starte Verarbeitung ...\n")
    else:
        iterable = gen
        total    = None
        bar_fmt  = "{desc}: {n_fmt} Dateien [{elapsed}, {rate_fmt}]"
        print("\nStarte direkte Verarbeitung (Bereinigung Ordner für Ordner) ...\n")

    print("-" * 66)

    stats           = {"UPDATED": 0, "SKIPPED": 0, "ERROR": 0, "RENAMED": 0,
                       "WOULD_UPDATE": 0}
    processed_count = 0
    run_completed   = False

    with tqdm(total=total, desc="Verarbeite", unit="Datei",
              bar_format=bar_fmt, disable=not show_progress) as pbar:
        for full_path in iterable:
            # Im Probelauf NICHT umbenennen - das Dateisystem bleibt
            # vollstaendig unangetastet.
            if not dry_run:
                sanitized_path = sanitize_file_on_disk(full_path, pbar)
                if sanitized_path is not None:
                    full_path = sanitized_path
                    stats["RENAMED"] += 1

            result = process_file_with_retries(
                full_path, ppt_app_global, pbar, skip_password, dry_run=dry_run)
            stats[result] += 1
            # Resume nur fuer erfolgreich neu serialisierte Dateien
            # fortschreiben. SKIPPED ist in 3c mehrdeutig (Passwort/Add-In/
            # kein Zugriff) und wird beim naechsten Lauf erneut versucht –
            # die Pruefung dafuer ist billig (Erkennung vor/bei Open).
            # Resume im Probelauf NICHT fortschreiben: sonst gaelten die
            # Dateien beim spaeteren Echt-Lauf als bereits erledigt.
            if resume_path and result == "UPDATED" and not dry_run:
                append_resume(resume_path, full_path)

            if result == "ERROR":
                try:
                    _ = ppt_app_global.Version if ppt_app_global is not None else None
                    if ppt_app_global is None:
                        raise AttributeError("ppt_app_global is None")
                except (pythoncom.com_error, AttributeError):
                    pbar.write("  ↻  PowerPoint reagiert nicht mehr. Starte PPT-Engine neu...")
                    _kill_user_powerpoint()
                    _cleanup_ppt_inetcache()
                    _cleanup_user_recent()
                    try:
                        ppt_app_global = win32com.client.DispatchEx("PowerPoint.Application")
                        _configure_ppt_instance(ppt_app_global)
                    except Exception as fatal_e:
                        pbar.write(f"  ✕  Kritisch: PPT konnte nicht neu gestartet werden: {fatal_e}")
                        ppt_app_global = None
                        break

            pbar.update(1)
            processed_count += 1

            if processed_count % PPT_RESTART_INTERVAL == 0:
                pbar.write(
                    f"  ↻  Speicher-Reset: Starte PowerPoint nach "
                    f"{PPT_RESTART_INTERVAL} Dateien neu..."
                )
                try:
                    if ppt_app_global is not None:
                        ppt_app_global.Quit()
                except Exception as _e:
                    detail_logger.debug(f"process_directory: Exception verworfen: {_e!r}")
                _kill_user_powerpoint()
                _cleanup_ppt_inetcache()
                _cleanup_user_recent()
                time.sleep(2)
                try:
                    ppt_app_global = win32com.client.DispatchEx("PowerPoint.Application")
                    _configure_ppt_instance(ppt_app_global)
                except Exception as e_restart:
                    pbar.write(f"  ✗  Fehler beim PPT-Neustart: {e_restart}")
                    ppt_app_global = None
                    break
        else:
            # for-else: Schleife lief ohne break vollstaendig durch.
            run_completed = True

    detail_logger.info(f"Verarbeitung abgeschlossen: {stats}")
    return stats, run_completed


# ==================================================================
# Einstiegspunkt
# ==================================================================

if __name__ == "__main__":
    # Fremde PowerPoint-Sitzungen erfassen, BEVOR eine eigene COM-Instanz
    # entsteht - danach waere die eigene nicht mehr unterscheidbar.
    snapshot_foreign_powerpoint_pids()

    exit_code  = 0
    start_time = datetime.now()

    PROG_NAME  = os.path.basename(sys.argv[0]) or "ppt_updater"
    INVOCATION = PROG_NAME if IS_FROZEN else f"python {PROG_NAME}"

    print()
    print("╔══════════════════════════════════════════════════════════════╗")
    print("║ POWERPOINT-UPDATER: AKTUALISIERUNG DURCH INSTALLIERTE ENGINE ║")
    print("╚══════════════════════════════════════════════════════════════╝")
    print(f"  Start: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # --- Kommandozeilenparameter ---
    parser = argparse.ArgumentParser(
        prog=PROG_NAME,
        description="PPT-Updater: Alle PowerPoint-Dateien durch aktuelle Engine neu serialisieren",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            f"Beispiele:\n"
            f"  Interaktiv:   {INVOCATION}\n"
            f"  Automatisch:  {INVOCATION} --dir \\\\server\\freigabe --auto-start\n"
            f"  Mit Optionen: {INVOCATION} --dir C:\\Daten --skip-pwd --count-first\n"
            f"\n"
            f"WICHTIG – --auto-start:\n"
            f"  Ausschließlich für automatisierte Dienste (Task Scheduler, SCCM, CI/CD).\n"
            f"  Das Skript beendet im Automatikmodus ALLE PowerPoint-Prozesse des\n"
            f"  ausführenden Benutzers ohne Rückfrage und ohne Vorwarnung. Ungespeicherte\n"
            f"  Änderungen gehen unwiederbringlich verloren. Niemals interaktiv mit\n"
            f"  offenen PowerPoint-Fenstern verwenden!"
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
    pwd_group = parser.add_mutually_exclusive_group()
    pwd_group.add_argument(
        "--skip-pwd", action="store_true",
        help="Passwortgeschützte Dateien überspringen"
    )
    pwd_group.add_argument(
        "--no-skip-pwd", action="store_true",
        help="Passwortgeschützte Dateien als Fehler loggen"
    )
    parser.add_argument(
        "--count-first", action="store_true",
        help="Dateien vorab zählen (ETA-Anzeige, aber langsamerer Start)"
    )
    parser.add_argument(
        "--resume", metavar="DATEI", default=None,
        help="Resume-Datei: bereits fertige Dateien überspringen (für geplante Tasks)"
    )
    parser.add_argument(
        "--exclude-dir", metavar="MUSTER", action="append", default=[],
        help="Teilpfad-Muster (Ordnername), das vom Durchlauf ausgeschlossen "
             "wird; mehrfach möglich (z.B. --exclude-dir Archiv --exclude-dir Templates)"
    )
    parser.add_argument("--dry-run", action="store_true",
        help="Probelauf: PowerPoint startet, oeffnet jede Datei und prueft "
             "Format/Makros/Schutz, SPEICHERT aber NICHTS. Liefert exakte "
             "Statistik, ohne eine Datei zu aendern.")
    parser.add_argument(
        "--mailto", metavar="ADRESSE", default=None,
        help="E-Mail-Adresse für Laufbericht nach Abschluss (Komma-getrennt für mehrere)"
    )
    parser.add_argument(
        "--mailfrom", metavar="ADRESSE", default=None,
        help="Absenderadresse für --mailto (Standard: ppt-updater@<smtp-host>)"
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
        help="SMTP-Benutzername (aktiviert Login + STARTTLS)"
    )
    parser.add_argument(
        "--smtp-pass", metavar="PASS", default=None,
        help="SMTP-Passwort für --smtp-user. ACHTUNG: in der Prozessliste "
             "sichtbar – besser per Umgebungsvariable SMTP_PASS setzen."
    )
    parser.add_argument(
        "--smtp-starttls", action="store_true",
        help="STARTTLS erzwingen, auch ohne --smtp-user"
    )
    args = parser.parse_args()

    auto_mode = args.auto_start

    # Vor dem ersten COM-Zugriff auf bereits laufende Sitzungen hinweisen.
    warn_running_powerpoint(auto_mode)

    # --- Trust-Center-Erinnerung ---
    if not auto_mode:
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
        print(f"     • {TEMP_PROCESS_PATH} (temporärer Arbeitsordner des Skripts)")
        print(r"     • Q:\ oder \\server\dfs (Quelldateien)")
        print("     Jeweils: ☑ Unterordner ... sind ebenfalls vertrauenswürdig")
        print()
        print("  4. DATENSCHUTZOPTIONEN:")
        print("     ☐ Beim Öffnen automatisch verknüpfte Daten aktualisieren")
        print("     → DEAKTIVIEREN (verhindert Popups bei eingebetteten Objekten)")
        print("!" * 66)

        if not ask_yes_no(
            "\n  Wurden diese Einstellungen in PowerPoint vorgenommen?",
            default_yes=False,
        ):
            print("\n❌  Abbruch. Bitte zuerst das Trust Center in PowerPoint konfigurieren.")
            sys.exit(0)

    # --- Konfiguration: CLI oder interaktiv ---
    if auto_mode and not args.dir:
        print("\n❌  Fehler: --dir fehlt im Automatikmodus.")
        print(f"    Verwendung: {INVOCATION} --dir \\\\server\\freigabe --auto-start")
        sys.exit(1)

    if args.dir:
        start_dir = sanitize_path(args.dir)
        if not os.path.isdir(prepare_long_path(start_dir)):
            print(f"\n❌  Fehler: Verzeichnis existiert nicht oder ist nicht erreichbar: '{start_dir}'")
            sys.exit(1)
    else:
        start_dir = ask_directory()

    # --- skip_pwd: einheitliche Auswertung ---
    if args.skip_pwd:
        skip_pwd = True
    elif args.no_skip_pwd:
        skip_pwd = False
    elif auto_mode:
        skip_pwd = True
    else:
        if not args.dry_run:
            args.dry_run = ask_yes_no(
                "\nProbelauf (Dry-Run)? Es wird geprüft, aber NICHTS gespeichert.\n"
                "  (Empfohlen vor dem ersten Echt-Lauf auf einer neuen Ablage)"
            )
        skip_pwd = ask_yes_no(
            "\nPasswortgeschützte / nicht öffenbare Dateien überspringen?",
            default_yes=True,
        )

    # --- Fortschrittsanzeige (3-Optionen-Menu) ---
    if args.count_first:
        show_progress = True
        count_first   = True
    elif auto_mode:
        show_progress = True
        count_first   = False
    else:
        show_progress, count_first = ask_progress_mode()

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
        candidate = get_auto_resume_path(start_dir, _LOG_DIR)
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

    # --- Konfiguration anzeigen + "Jetzt starten?" (vor PPT-Eingriff) ---
    appver_lines = [
        (16.0, "Office 2016 / 2019 / 2021 / 2024 / 365"),
        (15.0, "Office 2013"),
        (14.0, "Office 2010"),
        (12.0, "Office 2007"),
    ]
    target_appver_label = next(
        (lbl for v, lbl in appver_lines if v == TARGET_APP_VERSION),
        f"AppVersion {TARGET_APP_VERSION}",
    )

    print()
    print("=" * 66)
    print("  KONFIGURATION")
    print("=" * 66)
    print(f"  Verzeichnis:           {start_dir}")
    print("  Office-Versionen:      Office 2016 / 2019 / 2021 / 2024 / 365 (DE/EN)")
    print(f"  Ziel-AppVersion:       {TARGET_APP_VERSION} ({target_appver_label})")
    print(f"  Modus:                 {'Vorab zählen (ETA)' if count_first else 'Generator (kein ETA)'}")
    print(f"  Fortschrittsbalken:    {'an' if show_progress else 'aus (inline)'}")
    print(f"  Passwortdateien:       {'Überspringen (als SKIPPED loggen)' if skip_pwd else 'Fehler loggen'}")
    if args.exclude_dir:
        print(f"  Ausgeschlossen:        {', '.join(args.exclude_dir)}")
    if resume_file and (resume_set or not auto_resume_owned):
        _rlabel = ("automatisch" if auto_resume_owned else resume_file)
        _rcount = f"{len(resume_set)} übersprungen" if resume_set else "neu"
        print(f"  Resume:                {_rcount} ({_rlabel})")
    if args.mailto:
        print(f"  E-Mail-Bericht:        {args.mailto} via {args.smtp}")
    print(f"  PowerPoint-Neustart:   alle {PPT_RESTART_INTERVAL} Dateien (Speicherbereinigung)")
    if auto_mode:
        print("  Ausführungsmodus:      AUTOMATISCH (keine Bestätigung erforderlich)")
    print()
    print("  Verarbeitungslogik:")
    print("    PowerPoint hat keine Convert()-Methode und kein lastEdited/")
    print("    CompatibilityMode-Feld. Der einzige Weg, Legacy-Strukturen")
    print("    zu migrieren, ist Presentation.SaveAs() durch die aktuelle")
    print("    Engine. Daher wird JEDE Datei neu serialisiert – auch wenn")
    print("    AppVersion bereits ≥ Ziel ist.")
    print()
    print("    • .ppt   → .pptx / .pptm   (Formatkonvertierung)")
    print("    • .pps   → .ppsx / .ppsm   (Slideshow bleibt Slideshow)")
    print("    • .pot   → .potx / .potm   (Template bleibt Template)")
    print("    • .ppam  → übersprungen     (Add-In, nicht via SaveAs aktualisierbar)")
    print("    • .pptx / .pptm / .ppsx / .ppsm / .potx / .potm")
    print("                              → Neuserialierung (AppVersion-Update)")
    print("    • 'Als Final' Status wird entfernt")
    print()
    print("  AppVersion-Referenz (docProps/app.xml in OOXML-PowerPoint-Dateien):")
    for v, label in appver_lines:
        arrow = "  ← Ziel" if v == TARGET_APP_VERSION else ""
        print(f"    {v} = {label}{arrow}")
    print()
    print("  ⚠️  PowerPoint-Fenster:")
    print("     PowerPoint MUSS sichtbar laufen (technische Einschränkung).")
    print("     Das Fenster wird minimiert. NICHT schließen!")
    print("     ALLE laufenden PowerPoint-Sitzungen des Benutzers werden")
    print("     beendet, sobald der Lauf bestätigt wurde.")
    print("=" * 66)

    if not auto_mode and not ask_yes_no("\nJetzt starten?"):
        print("Abgebrochen.")
        sys.exit(0)

    # --- Einzelinstanz-Lock je Zielverzeichnis (vor jedem PPT-Eingriff) ---
    # Schuetzt gegen parallele Laeufe auf demselben Ordner (doppelter
    # Scheduled Task / ungeduldiger Admin), die sich sonst gegenseitig die
    # Dateien zerschiessen wuerden.
    active_lock_path = _acquire_lock(start_dir)
    if active_lock_path is None:
        print()
        print("❌  FEHLER: Es läuft bereits ein Skript-Lauf auf diesem Verzeichnis.")
        print(f"    Lock-Datei: {_lock_file_path(start_dir)}")
        print("    Wenn sicher kein anderer Lauf aktiv ist, Lock-Datei manuell löschen.")
        sys.exit(1)

    # --- Erst NACH Bestätigung: PowerPoint-Eingriff ---
    signal.signal(signal.SIGINT,  _signal_handler)
    signal.signal(signal.SIGTERM, _signal_handler)
    pythoncom.CoInitialize()

    # Restore-Privilegien (Admin-Kontext) fuer ACL/Owner-Erhalt aktivieren.
    _enable_restore_privileges()

    stats_result = None

    try:
        _kill_user_powerpoint()

        ppt_app_global = win32com.client.DispatchEx("PowerPoint.Application")
        _configure_ppt_instance(ppt_app_global)

        ppt_version = "Unbekannt"
        try:
            ppt_version = ppt_app_global.Version
        except Exception:
            pass
        print(f"\n  PowerPoint-Version:   {ppt_version}")

        # --- Trust-Center-Praxistest (mit Watchdog-Timeout) ---
        ok, err_msg = _verify_trust_center_for_temp(
            ppt_app_global, ppt_app_pid, timeout=25.0
        )
        if not ok:
            print()
            print("!" * 66)
            print("  ❌ VORAB-TEST FEHLGESCHLAGEN: Temp-Ordner nicht nutzbar")
            print("!" * 66)
            print(f"  Pfad:    {TEMP_PROCESS_PATH}")
            print(f"  Fehler:  {err_msg}")
            print()
            print("  Mögliche Ursachen:")
            print("    • Ordner ist NICHT als vertrauenswürdiger Speicherort in")
            print("      PowerPoint eingetragen (Trust Center → Vertrauenswürdige")
            print("      Speicherorte → Neuen Speicherort hinzufügen)")
            print("    • Geschützte Ansicht für unsichere Speicherorte ist noch aktiv")
            print("    • Keine Schreibrechte auf dem Ordner")
            print("    • Antivirus blockiert den Schreibzugriff")
            print("!" * 66)
            log_error("TRUST_CENTER_TEST", Exception(err_msg))
            if "TIMEOUT" in err_msg:
                ppt_app_global = None
                ppt_app_pid    = None
            exit_code = 1
            raise RuntimeError("Trust-Center-Test fehlgeschlagen")

        if args.dry_run:
            print(f"\nPROBELAUF (Dry-Run) – es wird NICHTS gespeichert: '{start_dir}'")
        else:
            print(f"\nVerarbeite: '{start_dir}'")
        print("-" * 66)

        stats_result, run_completed = process_directory(
            start_dir, skip_pwd, count_first,
            show_progress=show_progress,
            resume_path=resume_file, resume_set=resume_set,
            exclude_patterns=args.exclude_dir,
            dry_run=args.dry_run)

        # Auto-Resume-Datei nach vollstaendigem Lauf entfernen (eine vom
        # Nutzer via --resume uebergebene Datei bleibt unangetastet).
        if auto_resume_owned and run_completed and resume_file:
            delete_resume_file(resume_file)
            detail_logger.info("Auto-Resume-Datei nach vollständigem Lauf gelöscht.")

    except SystemExit:
        raise
    except Exception as e:
        print(f"\n❌ KRITISCHER FEHLER: {e}")
        log_error("GLOBAL", e)
        exit_code = 1

    finally:
        if ppt_app_global is not None:
            try:
                ppt_app_global.Quit()
            except Exception:
                pass
        _kill_user_powerpoint()
        if os.path.exists(TEMP_PROCESS_PATH):
            try:
                time.sleep(1)
                shutil.rmtree(TEMP_PROCESS_PATH)
            except OSError:
                pass
        try:
            _cleanup_user_temp()
        except Exception:
            pass
        try:
            _cleanup_ppt_inetcache()
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
        _release_lock(active_lock_path)

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
        total_sum = sum(v for k, v in stats_result.items() if k != "RENAMED")
        print("  STATISTIK:")
        if args.dry_run:
            print(f"    [✓] Würde neu serialisiert: {stats_result.get('WOULD_UPDATE', 0)}")
        else:
            print(f"    [✓] Aktualisiert / Konv.: {stats_result.get('UPDATED', 0)}")
        print(f"    [→] Übersprungen:          {stats_result.get('SKIPPED', 0)}")
        print(f"    [✗] Fehler:                {stats_result.get('ERROR', 0)}")
        print(f"        GESAMT:                {total_sum}")
        if stats_result.get("RENAMED", 0) > 0:
            print(f"    [📝] Dateinamen bereinigt: {stats_result['RENAMED']}")
        file_logger.info(
            f"VERARBEITUNG ABGESCHLOSSEN | "
            f"Verzeichnis: {start_dir} | "
            f"Dauer: {str(duration).split('.')[0]} | "
            f"Aktualisiert: {stats_result.get('UPDATED', 0)} | "
            f"Übersprungen: {stats_result.get('SKIPPED', 0)} | "
            f"Fehler: {stats_result.get('ERROR', 0)} | "
            f"Dateinamen bereinigt: {stats_result.get('RENAMED', 0)} | "
            f"Gesamt: {total_sum}"
        )

        write_run_summary(stats_result, start_dir, duration)

    print()
    if not auto_mode:
        print("  ⚠️  WICHTIG: TRUST-CENTER-EINSTELLUNGEN ZURÜCKSETZEN")
        print("     Bitte die in Schritt 1–4 geänderten Einstellungen in")
        print("     PowerPoint → Datei → Optionen → Trust Center wieder")
        print("     auf die ursprünglichen Werte zurücksetzen.")
        print()
        print("     Insbesondere:")
        print("     • Geschützte Ansicht wieder aktivieren")
        print("     • Makroeinstellungen zurücksetzen")
        print(f"     • Vertrauenswürdigen Speicherort entfernen: {TEMP_PROCESS_PATH}")
        print("     • 'Vertrauenswürdige Speicherorte im Netzwerk zulassen'")
        print("       nur deaktivieren, falls es vorher nicht aktiv war")
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

    # --- E-Mail-Bericht (für --auto-start-Serverbetrieb) ---
    if args.mailto and stats_result:
        _from = args.mailfrom or f"ppt-updater@{args.smtp}"
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

    if not auto_mode:
        try:
            input("\nBeliebige Taste drücken zum Beenden...")
        except EOFError:
            pass

    sys.exit(exit_code)
