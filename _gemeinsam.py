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

    # -- oeffentlich -----------------------------------------------
    def belegen(self) -> bool:
        """True, wenn die Sperre uns gehoert; False, wenn ein anderer Lauf laeuft."""
        eigener_name = self._prozessname(os.getpid())

        if os.path.exists(self.pfad):
            try:
                with open(self.pfad, "r", encoding="utf-8") as fh:
                    zeilen = [z.strip() for z in fh.read().splitlines() if z.strip()]
                if zeilen:
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
                                    return False
                            else:
                                # Altes Format ohne Erstellungszeit -
                                # konservativ als belegt behandeln.
                                return False
            except (ValueError, OSError, IndexError):
                pass

        try:
            ordner = os.path.dirname(self.pfad)
            if ordner:
                os.makedirs(ordner, exist_ok=True)
            eigene_ct = self._erstellungszeit(os.getpid())
            with open(self.pfad, "w", encoding="utf-8") as fh:
                fh.write(f"{os.getpid()}\n{eigener_name}\n"
                         f"{repr(eigene_ct) if eigene_ct is not None else ''}\n")
            self._belegt = True
        except OSError:
            # Keine Sperrdatei schreibbar: lieber weiterarbeiten als
            # blockieren - die Sperre ist eine Vorsichtsmassnahme,
            # keine Voraussetzung.
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
        if pfad:
            self.pfad = pfad
        else:
            basis = os.path.dirname(os.path.abspath(__file__))
            if not os.access(basis, os.W_OK):
                basis = (os.environ.get("LOCALAPPDATA")
                         or os.environ.get("TEMP") or ".")
            self.pfad = os.path.join(basis, "migration.jsonl")
        self.aktiv = True

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
