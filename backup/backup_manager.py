"""
Backup-Manager
Erstellt und verwaltet Datenbank-Backups als JSON.
Enthält außerdem Funktionen für ZIP-Backups und ZIP-Restore des gesamten Nesk3-Ordners.
"""
import os
import sys
import glob
import json
import shutil
import zipfile
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import BACKUP_DIR, BACKUP_MAX_KEEP, BASE_DIR


def _ensure_backup_dir() -> str:
    """Erstellt das Backup-Verzeichnis falls nicht vorhanden."""
    path = os.path.join(BASE_DIR, BACKUP_DIR)
    os.makedirs(path, exist_ok=True)
    return path


def create_backup(typ: str = "manuell") -> str:
    """
    Erstellt ein vollständiges Backup aller Tabellen als JSON.
    Gibt den Dateipfad zurück.
    """
    # TODO: Implementierung folgt
    return ""


def list_backups() -> list[dict]:
    """Gibt eine Liste aller vorhandenen Backups zurück."""
    backup_dir = _ensure_backup_dir()
    backups = []
    for fname in sorted(os.listdir(backup_dir), reverse=True):
        if fname.endswith(".json"):
            fpath = os.path.join(backup_dir, fname)
            size  = os.path.getsize(fpath)
            mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
            backups.append({
                "dateiname":  fname,
                "pfad":       fpath,
                "groesse_kb": round(size / 1024, 1),
                "erstellt":   mtime.strftime("%d.%m.%Y %H:%M"),
            })
    return backups


def restore_backup(filepath: str) -> int:
    """
    Stellt ein Backup wieder her.
    Gibt die Anzahl der wiederhergestellten Datensätze zurück.
    """
    # TODO: Implementierung folgt
    return 0


def _cleanup_old_backups(backup_dir: str):
    """Löscht ältere Backups wenn MAX_KEEP überschritten."""
    files = sorted(
        [f for f in os.listdir(backup_dir) if f.endswith(".json")]
    )
    while len(files) > BACKUP_MAX_KEEP:
        os.remove(os.path.join(backup_dir, files.pop(0)))


# ---------------------------------------------------------------------------
# Automatische Startup-DB-Backups (SQLite .db-Dateien, täglich angelegt)
# Speicherort: database SQL/Backup Data/db_backups/YYYY-MM-DD/
# ---------------------------------------------------------------------------

def _db_backup_root() -> str:
    from config import DB_PATH
    return os.path.join(os.path.dirname(DB_PATH), "Backup Data", "db_backups")


def _format_datum(tag: str) -> str:
    try:
        return datetime.strptime(tag, "%Y-%m-%d").strftime("%d.%m.%Y")
    except Exception:
        return tag


def _snapshots_fuer_tag(tag_pfad: str) -> list[dict]:
    """Gibt alle Snapshots (Zeitstempel-Gruppen) eines Tages zurück."""
    snapshots: dict[str, dict] = {}
    for f in sorted(glob.glob(os.path.join(tag_pfad, "*.db"))):
        fname = os.path.basename(f)
        # Kein _wiederherstellung-Unterordner
        parts = fname.rsplit("_", 1)
        if len(parts) != 2:
            continue
        ts_raw = parts[1].replace(".db", "")
        if len(ts_raw) != 6 or not ts_raw.isdigit():
            continue
        uhrzeit = f"{ts_raw[0:2]}:{ts_raw[2:4]} Uhr"
        if ts_raw not in snapshots:
            snapshots[ts_raw] = {"zeit": uhrzeit, "ts": ts_raw, "dateien": []}
        snapshots[ts_raw]["dateien"].append({
            "name": parts[0],
            "pfad": f,
            "groesse_kb": round(os.path.getsize(f) / 1024, 1),
        })
    return sorted(snapshots.values(), key=lambda x: x["ts"])


def list_db_backups() -> list[dict]:
    """
    Listet alle automatisch angelegten Startup-DB-Backups auf.
    Gibt eine Liste von Tages-Einträgen (neueste zuerst) zurück.
    """
    basis = _db_backup_root()
    if not os.path.isdir(basis):
        return []
    result = []
    for tag in sorted(os.listdir(basis), reverse=True):
        tag_pfad = os.path.join(basis, tag)
        # Nur echte Tages-Ordner (YYYY-MM-DD), keine _wiederherstellung etc.
        if not os.path.isdir(tag_pfad) or len(tag) != 10 or tag.count("-") != 2:
            continue
        db_dateien = glob.glob(os.path.join(tag_pfad, "*.db"))
        if not db_dateien:
            continue
        gesamt = sum(os.path.getsize(f) for f in db_dateien)
        snapshots = _snapshots_fuer_tag(tag_pfad)
        db_namen = {os.path.basename(f).rsplit("_", 1)[0] for f in db_dateien}
        result.append({
            "datum":             tag,
            "datum_anzeige":     _format_datum(tag),
            "pfad":              tag_pfad,
            "anzahl_dbs":        len(db_namen),
            "anzahl_snapshots":  len(snapshots),
            "groesse_mb":        round(gesamt / (1024 * 1024), 1),
            "snapshots":         snapshots,
        })
    return result


def restore_db_backup_as_copy(tag_pfad: str, ts: str | None = None) -> dict:
    """
    Kopiert DB-Backup-Dateien eines Snapshots in einen geschützten Unterordner.
    Die Live-Datenbanken werden NICHT verändert.
    Turso hat keinen Zugriff auf diesen Ordner.

    Parameters
    ----------
    tag_pfad : Pfad zum Tages-Ordner des Backups
    ts       : Zeitstempel (HHMMSS) des gewünschten Snapshots; None = neuester

    Returns
    -------
    dict mit {'erfolg', 'ziel', 'anzahl', 'meldung'}
    """
    if ts is None:
        # Neuesten Snapshot bestimmen
        alle = sorted(glob.glob(os.path.join(tag_pfad, "*.db")))
        if not alle:
            return {"erfolg": False, "ziel": "", "anzahl": 0, "meldung": "Keine Backup-Dateien gefunden."}
        letzter_ts = os.path.basename(alle[-1]).rsplit("_", 1)[-1].replace(".db", "")
        if len(letzter_ts) != 6:
            return {"erfolg": False, "ziel": "", "anzahl": 0, "meldung": "Zeitstempel ungültig."}
        ts = letzter_ts

    muster = glob.glob(os.path.join(tag_pfad, f"*_{ts}.db"))
    if not muster:
        return {"erfolg": False, "ziel": "", "anzahl": 0, "meldung": f"Snapshot {ts} nicht gefunden."}

    # Zielordner: _wiederherstellung/<YYYY-MM-DD_HHMMSS>/
    tag_name = os.path.basename(tag_pfad)
    ziel_basis = os.path.join(_db_backup_root(), "_wiederherstellung")
    ziel_name  = f"{tag_name}_{ts}"
    ziel_ordner = os.path.join(ziel_basis, ziel_name)

    if os.path.exists(ziel_ordner):
        # Bereits kopiert – einfach Pfad zurückgeben
        vorh = glob.glob(os.path.join(ziel_ordner, "*.db"))
        return {
            "erfolg": True, "ziel": ziel_ordner, "anzahl": len(vorh),
            "meldung": (
                f"Kopie bereits vorhanden ({len(vorh)} Datenbank-Datei(en)).\n\n"
                f"Speicherort:\n{ziel_ordner}\n\n"
                "Die Live-Datenbanken wurden NICHT verändert."
            ),
        }

    os.makedirs(ziel_ordner, exist_ok=True)
    kopiert = 0
    for src in sorted(muster):
        fname  = os.path.basename(src)
        # name_HHMMSS.db  →  name.db
        teile  = fname.rsplit("_", 1)
        zielname = teile[0] + ".db" if len(teile) == 2 else fname
        shutil.copy2(src, os.path.join(ziel_ordner, zielname))
        kopiert += 1

    uhrzeit = f"{ts[0:2]}:{ts[2:4]} Uhr"
    return {
        "erfolg": True,
        "ziel":   ziel_ordner,
        "anzahl": kopiert,
        "meldung": (
            f"{kopiert} Datenbank-Kopie(n) vom {_format_datum(tag_name)} {uhrzeit} gesichert.\n\n"
            f"Speicherort (kein Turso-Zugriff):\n{ziel_ordner}\n\n"
            "Die Live-Datenbanken wurden NICHT verändert.\n"
            "Im Notfall können die Dateien von dort manuell zurückgespielt werden."
        ),
    }


def list_restored_copies() -> list[dict]:
    """Listet alle bereits erstellten Wiederherstellungs-Kopien auf."""
    basis = os.path.join(_db_backup_root(), "_wiederherstellung")
    if not os.path.isdir(basis):
        return []
    result = []
    for name in sorted(os.listdir(basis), reverse=True):
        pfad = os.path.join(basis, name)
        if not os.path.isdir(pfad):
            continue
        dateien = glob.glob(os.path.join(pfad, "*.db"))
        groesse = sum(os.path.getsize(f) for f in dateien)
        result.append({
            "name":       name,
            "pfad":       pfad,
            "anzahl":     len(dateien),
            "groesse_mb": round(groesse / (1024 * 1024), 1),
        })
    return result


# ---------------------------------------------------------------------------
# ZIP-Backup  /  ZIP-Restore  (gesamter Nesk3-Quellcode-Ordner)
# ---------------------------------------------------------------------------

_CODE_BACKUP_DIR = os.path.join(BASE_DIR, "Backup Data")

# Ordner/Muster die beim ZIP-Backup NICHT einbezogen werden sollen
_ZIP_EXCLUDE_DIRS  = {'__pycache__', '.git', 'Backup Data', 'backup', 'build_tmp', 'Exe'}
_ZIP_EXCLUDE_EXTS  = {'.pyc', '.pyo'}


def create_zip_backup() -> str:
    """
    Erstellt ein vollständiges ZIP-Backup des Nesk3-Ordners (alle .py, .db, .ini, .json Dateien).
    Speichert das ZIP unter 'Backup Data/Nesk3_backup_<timestamp>.zip'.
    Gibt den vollständigen ZIP-Pfad zurück.
    """
    os.makedirs(_CODE_BACKUP_DIR, exist_ok=True)
    stamp    = datetime.now().strftime('%Y%m%d_%H%M%S')
    zip_name = f"Nesk3_backup_{stamp}.zip"
    zip_path = os.path.join(_CODE_BACKUP_DIR, zip_name)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(BASE_DIR):
            # Ausgeschlossene Ordner überspringen (in-place modifizieren)
            dirs[:] = [d for d in dirs if d not in _ZIP_EXCLUDE_DIRS]
            for fname in files:
                if os.path.splitext(fname)[1].lower() in _ZIP_EXCLUDE_EXTS:
                    continue
                full_path = os.path.join(root, fname)
                arcname   = os.path.relpath(full_path, BASE_DIR)
                zf.write(full_path, arcname)

    return zip_path


def list_zip_backups() -> list[dict]:
    """
    Gibt eine Liste aller ZIP-Backups im Backup-Data-Ordner zurück.
    Jedes Element: {'dateiname', 'pfad', 'groesse_kb', 'erstellt'}
    """
    if not os.path.isdir(_CODE_BACKUP_DIR):
        return []
    result = []
    for fname in sorted(os.listdir(_CODE_BACKUP_DIR), reverse=True):
        if fname.lower().endswith('.zip'):
            fpath = os.path.join(_CODE_BACKUP_DIR, fname)
            size  = os.path.getsize(fpath)
            mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
            result.append({
                'dateiname':  fname,
                'pfad':       fpath,
                'groesse_kb': round(size / 1024, 1),
                'erstellt':   mtime.strftime('%d.%m.%Y %H:%M'),
            })
    return result


def restore_from_zip(zip_path: str, ziel_ordner: str = None) -> dict:
    """
    Stellt einen Nesk3-Quellcode-Backup aus einer ZIP-Datei wieder her.

    Parameters
    ----------
    zip_path     : Vollständiger Pfad zur ZIP-Datei
    ziel_ordner  : Zielordner; Standard = BASE_DIR (= aktueller Nesk3-Ordner)

    Returns
    -------
    dict mit {'erfolg': bool, 'dateien': int, 'meldung': str}
    """
    if ziel_ordner is None:
        ziel_ordner = BASE_DIR

    if not os.path.isfile(zip_path):
        return {'erfolg': False, 'dateien': 0, 'meldung': f'ZIP nicht gefunden: {zip_path}'}

    if not zipfile.is_zipfile(zip_path):
        return {'erfolg': False, 'dateien': 0, 'meldung': 'Keine gültige ZIP-Datei.'}

    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            namelist = zf.namelist()
            # Nur .py / .db / .ini / .json / .txt Dateien wiederherstellen; niemals Backup Data selbst
            restore_names = [
                n for n in namelist
                if not n.replace('\\', '/').startswith('Backup Data/')
                and os.path.splitext(n)[1].lower() not in _ZIP_EXCLUDE_EXTS
            ]
            for member in restore_names:
                target = os.path.join(ziel_ordner, member)
                os.makedirs(os.path.dirname(target), exist_ok=True)
                with zf.open(member) as src, open(target, 'wb') as dst:
                    shutil.copyfileobj(src, dst)

        return {
            'erfolg':  True,
            'dateien': len(restore_names),
            'meldung': f'{len(restore_names)} Dateien aus {os.path.basename(zip_path)} wiederhergestellt.',
        }
    except Exception as e:
        return {'erfolg': False, 'dateien': 0, 'meldung': f'Fehler beim Wiederherstellen: {e}'}
