# -*- coding: utf-8 -*-
"""
Gemeinsame Grundfunktionen der Migrations-Skripte (Python-Teil).

WOZU
====
Die Sammlung ist durch Kopieren gewachsen. Dieselbe Aufgabe wurde dadurch
mehrfach geloest - und die Korrekturen wanderten nicht mit. Konkrete
Faelle aus der Durchsicht:

  * Das Umstellen der Konsole auf UTF-8 fehlte in 3c, 5 und
    pdf_bildcheck.py. Die drei brachen beim ersten Rahmenzeichen mit
    UnicodeEncodeError ab, waehrend acht Geschwister-Skripte den
    passenden Block seit jeher hatten.
  * Die Einzelinstanz-Sperre gab es nur in 3a, 3b und 3c. In 4a-4c und 5
    konnten zwei Laeufe gleichzeitig ueber denselben Bestand gehen.
  * Der Zeitstempel-Erhalt lief ueberall ueber win32file.SetFileTime,
    aber nur 5_OCR_PDF.py hatte einen Rueckfall auf os.utime.

Dieses Modul haelt die kanonische Fassung. Die Skripte binden es mit
RUECKFALL ein (siehe unten) - fehlt die Datei, arbeiten sie mit ihrer
eingebauten Kopie weiter. Damit bleibt jedes Skript einzeln lauffaehig
und die ps2exe-/PyInstaller-Uebersetzung funktioniert unveraendert.

EINBINDEN
=========
    from _gemeinsam import lade_grundbibliothek
    gem = lade_grundbibliothek()          # None, wenn nicht verfuegbar

oder direkt, mit Rueckfall:

    try:
        import _gemeinsam as gem
    except Exception:
        gem = None

Stand: 14.08.2026
"""

from __future__ import annotations

import json
import os
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

__all__ = [
    "konsole_auf_utf8",
    "Einzelinstanz",
    "zeitstempel_lesen",
    "zeitstempel_schreiben",
    "lade_pfade",
    "preset_liste",
    "Laufprotokoll",
    "schreibbares_verzeichnis",
    "recent_momentaufnahme",
    "recent_wurzel_hinzufuegen",
    "recent_ist_eigen",
    "recent_eigene_entfernen",
    "symbolschriften",
    "symbolschrift_namen",
    "ist_symbolschrift",
]

VERSION = "1.0.0"


# ==================================================================
# Konsole
# ==================================================================
def konsole_auf_utf8() -> bool:
    """Stellt stdout/stderr auf UTF-8 um. Muss VOR jedem print laufen.

    Ohne das bricht die erste Ausgabe mit Rahmenzeichen oder Emoji unter
    der Windows-Standardcodepage (cp850/cp1252) mit UnicodeEncodeError ab
    - im Konsolenfenster ebenso wie bei Umleitung in eine Datei.

    errors="replace" statt des strengen Standards: eine unuebersetzbare
    Ausgabe soll ein Fragezeichen erzeugen, nicht den Lauf abbrechen.
    """
    ok = True
    for stream in (sys.stdout, sys.stderr):
        try:
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            ok = False
    return ok


# ==================================================================
# Einzelinstanz-Sperre
# ==================================================================
class Einzelinstanz:
    """Lock-Datei mit PID, Prozessname und Erstellungszeit.

    Die Erstellungszeit ist der Kern: auf Rechnern mit langer Laufzeit
    vergibt Windows eine freigewordene PID binnen Minuten neu. Ohne den
    Abgleich sperrt ein beliebiges anderes Python-Skript mit passender
    PID das Werkzeug dauerhaft aus.

    Verwendung:
        sperre = Einzelinstanz("4a_ersetze_font_in_word")
        if not sperre.belegen():
            print(sperre.hinweis()); sys.exit(1)
        try:
            ...
        finally:
            sperre.freigeben()
    """

    def __init__(self, name: str, verzeichnis: Optional[str] = None) -> None:
        basis = verzeichnis or os.environ.get("TEMP") or os.environ.get("TMP") or r"C:\tmp"
        self.pfad = os.path.join(basis, f"{name}.lock")
        self._belegt = False

    # -- intern ----------------------------------------------------
    @staticmethod
    def _prozessname(pid: int) -> str:
        try:
            import psutil
            return psutil.Process(pid).name().lower()
        except Exception:
            return ""

    @staticmethod
    def _erstellungszeit(pid: int) -> Optional[float]:
        try:
            import psutil
            return psutil.Process(pid).create_time()
        except Exception:
            return None

    @staticmethod
    def _lebt(pid: int) -> bool:
        try:
            import psutil
            return psutil.pid_exists(pid)
        except Exception:
            # Ohne psutil lieber nicht sperren, als faelschlich aussperren.
            return False

    def _inhalt_lesen(self, versuche: int = 5) -> Optional[str]:
        """Inhalt der Sperrdatei; None, wenn sie (nicht mehr) da ist.

        Ein soeben per O_EXCL angelegter, noch leerer Sperrsatz gehoert
        einem Lauf, der gerade schreibt - kurz nachfassen, statt ihn als
        verwaist zu werten.
        """
        inhalt = ""
        for _ in range(max(1, versuche)):
            try:
                with open(self.pfad, "r", encoding="utf-8") as fh:
                    inhalt = fh.read()
            except FileNotFoundError:
                return None
            except (OSError, ValueError):
                inhalt = ""
            if inhalt.strip():
                return inhalt
            time.sleep(0.2)
        return inhalt

    def _ist_fremd_belegt(self, inhalt: str, eigener_name: str) -> bool:
        try:
            zeilen = [z.strip() for z in inhalt.splitlines() if z.strip()]
            if not zeilen:
                return False
            alte_pid = int(zeilen[0])
            alter_name = zeilen[1].lower() if len(zeilen) > 1 else ""
            alte_ct: Optional[float] = None
            if len(zeilen) > 2:
                try:
                    alte_ct = float(zeilen[2])
                except ValueError:
                    alte_ct = None

            if alte_pid != os.getpid() and self._lebt(alte_pid):
                name_jetzt = self._prozessname(alte_pid)
                passt = (
                    name_jetzt == eigener_name
                    or (alter_name and name_jetzt == alter_name)
                    or "python" in name_jetzt
                )
                if passt:
                    if alte_ct is not None:
                        ct_jetzt = self._erstellungszeit(alte_pid)
                        if ct_jetzt is not None and abs(ct_jetzt - alte_ct) < 1.0:
                            return True
                    else:
                        # Altes Format ohne Erstellungszeit -
                        # konservativ als belegt behandeln.
                        return True
        except (ValueError, IndexError):
            pass
        return False

    # -- oeffentlich -----------------------------------------------
    def belegen(self) -> bool:
        """True, wenn die Sperre uns gehoert; False, wenn ein anderer Lauf laeuft.

        Frueher: os.path.exists pruefen, dann open(..., "w"). Zwei
        gleichzeitig gestartete Laeufe sahen beide keine Sperrdatei und
        schrieben beide - der zweite ueberschrieb den ersten, beide liefen
        (nachgestellt 29.09.: sechs Prozesse mit gemeinsamem Startzeitpunkt,
        fuenf Runden - jedes Mal belegten alle sechs).
        Jetzt legt os.open(O_CREAT|O_EXCL) die Datei atomar an: genau einer
        gewinnt. Eine vorhandene Sperre wird wie bisher auf ihren Besitzer
        geprueft (PID, Prozessname, Erstellungszeit); ist sie verwaist,
        wird sie entfernt - aber nur, wenn ihr Inhalt unmittelbar vorher
        noch derselbe ist - und das atomare Anlegen erneut versucht.
        Restrisiko: zwei Laeufe, die DIESELBE verwaiste Sperre im selben
        Augenblick entfernen, koennen sich in einem Fenster von Mikro-
        sekunden noch ueberholen.
        """
        eigener_name = self._prozessname(os.getpid())
        eigene_ct = self._erstellungszeit(os.getpid())
        satz = (f"{os.getpid()}\n{eigener_name}\n"
                f"{repr(eigene_ct) if eigene_ct is not None else ''}\n")

        try:
            ordner = os.path.dirname(self.pfad)
            if ordner:
                os.makedirs(ordner, exist_ok=True)
        except OSError:
            pass

        for _ in range(5):
            try:
                fd = os.open(self.pfad, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o644)
            except FileExistsError:
                inhalt = self._inhalt_lesen()
                if inhalt is None:
                    continue            # inzwischen entfernt - neu versuchen
                if self._ist_fremd_belegt(inhalt, eigener_name):
                    return False
                # Verwaist (tote PID, fremder Prozess, eigene PID, leer oder
                # unlesbar): nur entfernen, wenn noch derselbe Inhalt drinsteht.
                if self._inhalt_lesen(versuche=1) == inhalt:
                    try:
                        os.remove(self.pfad)
                    except FileNotFoundError:
                        pass
                    except OSError:
                        break
                continue
            except OSError:
                break
            try:
                os.write(fd, satz.encode("utf-8"))
            except OSError:
                pass
            finally:
                os.close(fd)
            self._belegt = True
            return True

        # Keine Sperrdatei anlegbar: lieber weiterarbeiten als blockieren -
        # die Sperre ist eine Vorsichtsmassnahme, keine Voraussetzung (wie
        # bisher). Ein lebender Besitzer hat oben bereits False geliefert.
        self._belegt = True
        return True

    def freigeben(self) -> None:
        try:
            if not os.path.exists(self.pfad):
                return
            with open(self.pfad, "r", encoding="utf-8") as fh:
                zeilen = [z.strip() for z in fh.read().splitlines() if z.strip()]
            if zeilen and int(zeilen[0]) == os.getpid():
                os.remove(self.pfad)
        except Exception:
            pass

    def hinweis(self) -> str:
        return ("\nEs läuft bereits eine Instanz dieses Skripts.\n"
                "Falls das nicht stimmt, Sperrdatei löschen:\n"
                f"    {self.pfad}")


# ==================================================================
# Zeitstempel
# ==================================================================
def zeitstempel_lesen(pfad: str) -> Optional[Dict[str, Any]]:
    """Liest Erstellungs-, Zugriffs- und Aenderungszeit.

    Bevorzugt win32file (erfasst auch die Erstellungszeit), faellt sonst
    auf os.stat zurueck. Diesen Rueckfall hatte bisher nur 5_OCR_PDF.py;
    ohne pywin32 verloren alle anderen Skripte die Zeitstempel
    stillschweigend.
    """
    try:
        import win32file
        handle = win32file.CreateFile(
            pfad, win32file.GENERIC_READ,
            win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
            None, win32file.OPEN_EXISTING, 0, None)
        try:
            erstellt, zugriff, geaendert = win32file.GetFileTime(handle)
            return {"art": "win32", "werte": (erstellt, zugriff, geaendert)}
        finally:
            handle.Close()
    except Exception:
        pass

    try:
        st = os.stat(pfad)
        return {"art": "stat", "werte": (st.st_atime, st.st_mtime)}
    except Exception:
        return None


def zeitstempel_schreiben(pfad: str, daten: Optional[Dict[str, Any]]) -> bool:
    """Schreibt die von zeitstempel_lesen gelieferten Werte zurueck."""
    if not daten:
        return False

    if daten.get("art") == "win32":
        try:
            import win32file
            import win32con
            handle = win32file.CreateFile(
                pfad, win32file.GENERIC_WRITE,
                win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
                None, win32file.OPEN_EXISTING,
                win32con.FILE_ATTRIBUTE_NORMAL, None)
            try:
                erstellt, zugriff, geaendert = daten["werte"]
                win32file.SetFileTime(handle, erstellt, zugriff, geaendert)
                return True
            finally:
                handle.Close()
        except Exception:
            # Weiter zum os.utime-Rueckfall: die Aenderungszeit ist das,
            # worauf es im Ergebnis ankommt.
            try:
                _, zugriff, geaendert = daten["werte"]
                os.utime(pfad, (zugriff.timestamp(), geaendert.timestamp()))
                return True
            except Exception:
                return False

    if daten.get("art") == "stat":
        try:
            atime, mtime = daten["werte"]
            os.utime(pfad, (atime, mtime))
            return True
        except Exception:
            return False

    return False


# ==================================================================
# Zentrale Pfad-Presets
# ==================================================================
_PFADE_CACHE: Optional[Dict[str, Any]] = None

_PFADE_STANDARD: Dict[str, Any] = {
    "presets": [
        "Q:\\",
        "R:\\",
        "G:\\Geteilte Ablagen",
        "G:\\Meine Ablage",
        "\\\\server\\dfs",
    ],
    "hunspell": {
        "nupkg": "https://www.nuget.org/api/v2/package/NHunspell/1.2.5554.16953",
        "dic": "https://raw.githubusercontent.com/LibreOffice/dictionaries/master/de/de_DE_frami.dic",
        "aff": "https://raw.githubusercontent.com/LibreOffice/dictionaries/master/de/de_DE_frami.aff",
    },
}


def lade_pfade(neu_laden: bool = False) -> Dict[str, Any]:
    """Liest pfade.json neben dem Skript; faellt auf die Vorgaben zurueck.

    Die Presets standen als Platzhalter in acht Skripten - mal mit, mal
    ohne abschliessenden Backslash, in zwei verschiedenen Reihenfolgen.
    Wer sie an seine Umgebung anpasst, musste acht Stellen pflegen.
    """
    global _PFADE_CACHE
    if _PFADE_CACHE is not None and not neu_laden:
        return _PFADE_CACHE

    daten = dict(_PFADE_STANDARD)
    kandidat = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pfade.json")
    try:
        if os.path.exists(kandidat):
            with open(kandidat, "r", encoding="utf-8-sig") as fh:
                aus_datei = json.load(fh)
            if isinstance(aus_datei, dict):
                daten.update(aus_datei)
    except Exception:
        # Kaputte Konfiguration darf keinen Lauf verhindern.
        pass

    _PFADE_CACHE = daten
    return daten


def preset_liste() -> List[str]:
    """Die Verzeichnis-Presets als Liste, leere Eintraege entfernt."""
    presets = lade_pfade().get("presets") or []
    return [str(p) for p in presets if str(p).strip()]


# ==================================================================
# Gemeinsames Laufprotokoll (JSON Lines)
# ==================================================================
class Laufprotokoll:
    """Eine Zeile je Datei und Schritt, ueber alle Skripte hinweg.

    Jedes Skript schreibt bisher sein eigenes Format: CSV mit Semikolon,
    CSV mit Komma, .log, XLSX, teils UTF-8 mit BOM, teils ohne. Damit
    laesst sich der Gesamtfortschritt ueber die elf Schritte nicht
    auswerten - etwa die Frage, welche Dateien in Schritt 2 liegen
    blieben und in Schritt 7 wieder auftauchen.

    Diese Datei ERSETZT die bestehenden Protokolle nicht, sie ergaenzt
    sie. Angehaengt wird zeilenweise, damit parallele Laeufe und harte
    Abbrueche nichts zerstoeren.
    """

    def __init__(self, skript: str, pfad: Optional[str] = None) -> None:
        self.skript = skript
        self.aktiv = True
        # Abschnitt "protokoll" aus pfade.json. Die Datei beschreibt ihn
        # ("datei": leer = neben den Skripten, "aktiv"), gelesen wurde er
        # aber nie - weder "aktiv": false noch ein eigener Pfad hatten eine
        # Wirkung. Ein ausdruecklich uebergebener pfad geht weiter vor.
        einstellung: Dict[str, Any] = {}
        try:
            roh = lade_pfade().get("protokoll")
            if isinstance(roh, dict):
                einstellung = roh
        except Exception:
            einstellung = {}
        aktiv = einstellung.get("aktiv", True)
        if aktiv in (False, 0) or str(aktiv).strip().lower() in ("false", "nein", "0"):
            self.aktiv = False
        datei = str(einstellung.get("datei") or "").strip()
        if pfad:
            self.pfad = pfad
        elif datei:
            datei = os.path.expandvars(datei)
            if not os.path.isabs(datei):
                datei = os.path.join(os.path.dirname(os.path.abspath(__file__)), datei)
            if os.path.isdir(datei):
                datei = os.path.join(datei, "migration.jsonl")
            self.pfad = datei
        else:
            basis = os.path.dirname(os.path.abspath(__file__))
            if not os.access(basis, os.W_OK):
                basis = (os.environ.get("LOCALAPPDATA")
                         or os.environ.get("TEMP") or ".")
            self.pfad = os.path.join(basis, "migration.jsonl")

    def schreibe(self, pfad_datei: str, aktion: str, status: str,
                 detail: str = "", **weitere: Any) -> None:
        if not self.aktiv:
            return
        satz = {
            "ts": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "skript": self.skript,
            "pfad": pfad_datei,
            "aktion": aktion,
            "status": status,
        }
        if detail:
            satz["detail"] = detail
        satz.update(weitere)
        try:
            with open(self.pfad, "a", encoding="utf-8") as fh:
                fh.write(json.dumps(satz, ensure_ascii=False) + "\n")
        except Exception:
            # Ein nicht schreibbares Gesamtprotokoll darf den Lauf nicht
            # anhalten - die skripteigenen Protokolle laufen weiter.
            self.aktiv = False


# ==================================================================
# Schreibbares Verzeichnis finden
# ==================================================================
def schreibbares_verzeichnis(kandidaten: List[Optional[str]]) -> Optional[str]:
    """Erster Kandidat, in den sich wirklich schreiben laesst.

    Der Test ist echt (Datei anlegen und wieder loeschen): ein
    existierendes Verzeichnis sagt nichts ueber das Schreibrecht aus -
    und genau daran scheiterten Laeufe erst am Ende, nachdem stundenlang
    gearbeitet worden war.
    """
    import uuid
    for kandidat in kandidaten:
        if not kandidat or not str(kandidat).strip():
            continue
        try:
            os.makedirs(kandidat, exist_ok=True)
            probe = os.path.join(kandidat, f".writetest_{uuid.uuid4().hex}.tmp")
            with open(probe, "w", encoding="utf-8") as fh:
                fh.write("x")
            os.remove(probe)
            return kandidat
        except Exception:
            continue
    return None


# ==================================================================
# Zuletzt verwendet: nur die eigenen Eintraege entfernen
# ==================================================================
# Bis 29.09.2026 loeschten 3a-4c JEDE .lnk-Datei in
# %APPDATA%\Microsoft\Windows\Recent und 3a zusaetzlich die ganze
# Word-Liste "Zuletzt verwendet" - auch die Eintraege, die der Anwender
# selbst angelegt hatte. Jetzt gilt:
#   * recent_momentaufnahme() merkt sich beim Skriptstart, welche
#     Verknuepfungen schon da waren. Die bleiben immer stehen.
#   * recent_eigene_entfernen() loescht nur Verknuepfungen, die NACH der
#     Momentaufnahme entstanden sind UND deren Ziel in einem eigenen
#     Ordner liegt: im bearbeiteten Verzeichnis (recent_wurzel_hinzufuegen)
#     oder im Arbeitsordner der Skripte bzw. %TEMP%.
# Ohne Momentaufnahme wird nichts geloescht. Oeffnet der Anwender waehrend
# des Laufs selbst eine Datei aus dem bearbeiteten Verzeichnis, faellt
# deren neuer Eintrag mit weg - von aussen nicht unterscheidbar.

_RECENT_VORHER: Optional[set] = None
_RECENT_WURZELN: List[str] = []


def _recent_ordner() -> Optional[str]:
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None
    d = os.path.join(appdata, "Microsoft", "Windows", "Recent")
    return d if os.path.isdir(d) else None


def _pfad_norm(p: str) -> str:
    p = os.path.abspath(p)
    # %TEMP% steht oft in 8.3-Form (C:\Users\ABCDEF~1\...), das Ziel einer
    # Verknuepfung aber in Langform - ohne Angleichen passt nichts zusammen.
    if "~" in p:
        try:
            import ctypes
            puffer = ctypes.create_unicode_buffer(32768)
            if ctypes.windll.kernel32.GetLongPathNameW(p, puffer, 32768):
                p = puffer.value
        except Exception:
            pass
    p = os.path.normcase(p)
    if p.startswith("\\\\?\\unc\\"):
        p = "\\\\" + p[8:]
    elif p.startswith("\\\\?\\"):
        p = p[4:]
    return p.rstrip("\\")


def _eigene_wurzeln() -> List[str]:
    wurzeln = list(_RECENT_WURZELN)
    lokal = os.environ.get("LOCALAPPDATA")
    if lokal:
        wurzeln.append(_pfad_norm(os.path.join(lokal, "Dateimigration-Arbeitskopien")))
    temp = os.environ.get("TEMP")
    if temp:
        wurzeln.append(_pfad_norm(temp))
    return [w for w in wurzeln if w]


def recent_momentaufnahme() -> None:
    """Merkt sich die vorhandenen Recent-Verknuepfungen (nur beim ersten Aufruf)."""
    global _RECENT_VORHER
    if _RECENT_VORHER is not None:
        return
    d = _recent_ordner()
    if d is None:
        _RECENT_VORHER = set()
        return
    try:
        _RECENT_VORHER = {n.lower() for n in os.listdir(d)}
    except OSError:
        # Unbekannter Ausgangszustand: lieber gar nichts loeschen.
        _RECENT_VORHER = None


def recent_wurzel_hinzufuegen(pfad: Optional[str]) -> None:
    """Meldet ein bearbeitetes Verzeichnis als eigene Wurzel an."""
    if pfad:
        try:
            w = _pfad_norm(pfad)
        except Exception:
            return
        if w and w not in _RECENT_WURZELN:
            _RECENT_WURZELN.append(w)


def recent_ist_eigen(pfad: Optional[str]) -> bool:
    """True, wenn pfad in einer eigenen Wurzel liegt (oder sie selbst ist)."""
    if not pfad:
        return False
    try:
        p = _pfad_norm(pfad)
    except Exception:
        return False
    for w in _eigene_wurzeln():
        if p == w or p.startswith(w + "\\"):
            return True
    return False


def _lnk_ziel_lesen(pythoncom, shell, lnk: str) -> Optional[str]:
    link = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
        shell.IID_IShellLink)
    link.QueryInterface(pythoncom.IID_IPersistFile).Load(lnk, 0)
    return link.GetPath(0)[0] or None


def _lnk_ziel(lnk: str) -> Optional[str]:
    try:
        import pythoncom
        from win32com.shell import shell
    except Exception:
        return None
    # Die Skripte haben COM im Hauptthread laengst initialisiert. Ein
    # eigenes CoInitialize/CoUninitialize-Paar stoerte dort die noch
    # lebenden Office-Objekte (pywin32 meldete beim Freigeben einen
    # Win32-Fehler) - deshalb nur in einem Thread ohne COM selbst
    # initialisieren.
    try:
        return _lnk_ziel_lesen(pythoncom, shell, lnk)
    except pythoncom.com_error as e:
        if e.hresult != -2147221008:        # CO_E_NOTINITIALIZED
            return None
    except Exception:
        return None
    try:
        pythoncom.CoInitialize()
    except Exception:
        return None
    try:
        return _lnk_ziel_lesen(pythoncom, shell, lnk)
    except Exception:
        return None
    finally:
        pythoncom.CoUninitialize()


def recent_eigene_entfernen() -> Tuple[int, int]:
    """Loescht neue Recent-Verknuepfungen auf eigene Pfade.

    Rueckgabe: (geloescht, gesperrt).
    """
    if _RECENT_VORHER is None:
        return 0, 0
    d = _recent_ordner()
    if d is None:
        return 0, 0
    geloescht = gesperrt = 0
    try:
        namen = os.listdir(d)
    except OSError:
        return 0, 0
    for name in namen:
        if not name.lower().endswith(".lnk") or name.lower() in _RECENT_VORHER:
            continue
        voll = os.path.join(d, name)
        if not os.path.isfile(voll) or not recent_ist_eigen(_lnk_ziel(voll)):
            continue
        try:
            os.remove(voll)
            geloescht += 1
        except OSError:
            gesperrt += 1
    return geloescht, gesperrt


# ==================================================================
# Symbolschriften (fuer 4a/4b/4c)
# ==================================================================
# In Wingdings, Symbol & Co. steht hinter jedem Zeichencode ein Bild
# (Haekchen, Pfeil, Kaestchen). Stellt ein Skript solchen Text auf Arial um,
# bleibt der Code stehen und aus dem Haekchen wird ein Buchstabe ("ü" statt
# des Hakens). Diese Schriften duerfen deshalb nie ersetzt werden.

SYMBOLSCHRIFTEN_FEST_NAMEN = (
    "Symbol", "Wingdings", "Wingdings 2", "Wingdings 3", "Webdings",
    "Marlett", "MT Extra", "Bookshelf Symbol 7", "MS Reference Specialty",
    "MS Outlook", "ZapfDingbats", "Zapf Dingbats", "ITC Zapf Dingbats",
    "Monotype Sorts", "Segoe MDL2 Assets", "Segoe Fluent Icons",
    "Holo MDL2 Assets", "Wingdings-Regular",
)
SYMBOLSCHRIFTEN_FEST = frozenset(n.lower() for n in SYMBOLSCHRIFTEN_FEST_NAMEN)

_SYMBOLSCHRIFTEN_CACHE: Optional[frozenset] = None
_SYMBOLSCHRIFTEN_NAMEN: List[str] = []


def symbolschriften() -> frozenset:
    """Namen (klein) aller Symbolschriften: feste Liste + installierte.

    Installierte Schriften mit SYMBOL_CHARSET liefert GDI
    (EnumFontFamiliesExW); damit sind auch Symbolschriften erfasst, die in
    der festen Liste fehlen.
    """
    global _SYMBOLSCHRIFTEN_CACHE
    if _SYMBOLSCHRIFTEN_CACHE is not None:
        return _SYMBOLSCHRIFTEN_CACHE
    namen = {n.lower(): n for n in SYMBOLSCHRIFTEN_FEST_NAMEN}
    try:
        import ctypes
        from ctypes import wintypes

        class LOGFONTW(ctypes.Structure):
            _fields_ = [("lfHeight", wintypes.LONG), ("lfWidth", wintypes.LONG),
                        ("lfEscapement", wintypes.LONG), ("lfOrientation", wintypes.LONG),
                        ("lfWeight", wintypes.LONG), ("lfItalic", wintypes.BYTE),
                        ("lfUnderline", wintypes.BYTE), ("lfStrikeOut", wintypes.BYTE),
                        ("lfCharSet", wintypes.BYTE), ("lfOutPrecision", wintypes.BYTE),
                        ("lfClipPrecision", wintypes.BYTE), ("lfQuality", wintypes.BYTE),
                        ("lfPitchAndFamily", wintypes.BYTE),
                        ("lfFaceName", wintypes.WCHAR * 32)]

        SYMBOL_CHARSET = 2
        rueckruf_typ = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.POINTER(LOGFONTW),
                                          ctypes.c_void_p, wintypes.DWORD, wintypes.LPARAM)

        def _rueckruf(lf, _tm, _typ, _param):
            if lf.contents.lfCharSet == SYMBOL_CHARSET:
                name = lf.contents.lfFaceName.lstrip("@").strip()
                if name:
                    namen.setdefault(name.lower(), name)
            return 1

        gdi32 = ctypes.windll.gdi32
        user32 = ctypes.windll.user32
        user32.GetDC.restype = wintypes.HDC
        gdi32.EnumFontFamiliesExW.argtypes = [wintypes.HDC, ctypes.POINTER(LOGFONTW),
                                              rueckruf_typ, wintypes.LPARAM, wintypes.DWORD]
        hdc = user32.GetDC(None)
        try:
            lf = LOGFONTW()
            lf.lfCharSet = SYMBOL_CHARSET
            gdi32.EnumFontFamiliesExW(hdc, ctypes.byref(lf), rueckruf_typ(_rueckruf), 0, 0)
        finally:
            user32.ReleaseDC(None, hdc)
    except Exception:
        pass      # feste Liste genuegt als Rueckfall
    _SYMBOLSCHRIFTEN_NAMEN[:] = sorted(namen.values(), key=str.lower)
    _SYMBOLSCHRIFTEN_CACHE = frozenset(namen)
    return _SYMBOLSCHRIFTEN_CACHE


def symbolschrift_namen() -> List[str]:
    """Dieselben Schriften in Originalschreibung (fuer Suchen mit Format)."""
    symbolschriften()
    return list(_SYMBOLSCHRIFTEN_NAMEN)


def ist_symbolschrift(name: Optional[str]) -> bool:
    return bool(name) and str(name).strip().lower() in symbolschriften()


# ==================================================================
# Bequemer Rueckfall-Loader
# ==================================================================
def lade_grundbibliothek():
    """Gibt dieses Modul zurueck - Gegenstueck fuer den try/except-Import."""
    return sys.modules[__name__]


if __name__ == "__main__":
    konsole_auf_utf8()
    print(f"Grundbibliothek der Migrations-Skripte, Version {VERSION}")
    print(f"  Presets       : {preset_liste()}")
    print(f"  Protokollpfad : {Laufprotokoll('selbsttest').pfad}")
    print("  Dieses Modul wird von den Skripten importiert, nicht direkt gestartet.")
