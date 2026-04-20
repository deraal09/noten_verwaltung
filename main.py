"""
Main-Einstiegspunkt für Notenverwaltung
"""

import os
import logging
from typing import Optional

from constants import DEFAULT_GEWICHTUNG, DATA_FILE, HALBJAHRE
from constants import get_ns_aliases as _get_ns_aliases
from models import NotenVerwaltung
from dialogs import PasswordDialog

logger = logging.getLogger(__name__)


def _migrate_old_md() -> Optional[NotenVerwaltung]:
    """Migriert alte noten.md Dateien zum neuen Format."""
    md_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "noten.md")
    if not os.path.exists(md_file) or os.path.exists(DATA_FILE):
        return None

    old_data = NotenVerwaltung()
    old_data.gewichtung_muendlich = DEFAULT_GEWICHTUNG
    cur_sj = cur_kl = cur_sk = cur_hj = None

    try:
        with open(md_file, "r", encoding="utf-8") as f:
            for z in f:
                z = z.strip()
                if z.startswith("Gewichtung Unterrichtsleistung:") or z.startswith("Gewichtung Mündlich:"):
                    try:
                        w = int(z.split(":")[1].replace("%", "").strip())
                        if 0 <= w <= 100:
                            old_data.gewichtung_muendlich = w
                    except ValueError:
                        pass
                elif z.startswith("## Schuljahr "):
                    cur_sj = z[13:].strip()
                    old_data.schuljahre[cur_sj] = {}
                    cur_kl = cur_sk = cur_hj = None
                elif z.startswith("### Klasse ") and cur_sj:
                    cur_kl = z[11:].strip()
                    old_data.schuljahre[cur_sj][cur_kl] = {
                        "notenschluessel": "IHK", "notenschluessel_csv": "",
                        "schuelerinnen": {}, "faecher": {}}
                    cur_sk = cur_hj = None
                elif z.startswith("#### ") and cur_sj and cur_kl:
                    t = z[5:].split(",", 1)
                    nn = t[0].strip()
                    vn = t[1].strip() if len(t) > 1 else ""
                    cur_sk = NotenVerwaltung._key(nn, vn)
                    old_data.schuljahre[cur_sj][cur_kl]["schuelerinnen"][cur_sk] = {
                        "nachname": nn, "vorname": vn}
                    cur_hj = None
                elif z.startswith("##### ") and cur_sk:
                    cur_hj = z[6:].strip()
                elif z.startswith("- Unterrichtsleistung (manuell):") and cur_sk and cur_hj:
                    ns = z.split(":", 1)[1].strip()
                    if ns:
                        fach = old_data.schuljahre[cur_sj][cur_kl]["faecher"].setdefault(
                            "Allgemein", {"halbjahre": {}, "klausuren": {}, "unterrichtsleistungen": {}})
                        hj_data = fach["halbjahre"].setdefault(cur_hj, {"noten": {}})
                        hj_data["noten"].setdefault(cur_sk, {"muendlich": [], "schriftlich": []})
                        hj_data["noten"][cur_sk]["muendlich"] = [int(x.strip()) for x in ns.split(",") if x.strip()]
                elif z.startswith("- Mündlich:") and cur_sk and cur_hj:
                    ns = z[11:].strip()
                    if ns:
                        fach = old_data.schuljahre[cur_sj][cur_kl]["faecher"].setdefault(
                            "Allgemein", {"halbjahre": {}, "klausuren": {}, "unterrichtsleistungen": {}})
                        hj_data = fach["halbjahre"].setdefault(cur_hj, {"noten": {}})
                        hj_data["noten"].setdefault(cur_sk, {"muendlich": [], "schriftlich": []})
                        hj_data["noten"][cur_sk]["muendlich"] = [int(x.strip()) for x in ns.split(",") if x.strip()]
                elif z.startswith("- Schriftlich:") and cur_sk and cur_hj:
                    ns = z[14:].strip()
                    if ns:
                        fach = old_data.schuljahre[cur_sj][cur_kl]["faecher"].setdefault(
                            "Allgemein", {"halbjahre": {}, "klausuren": {}, "unterrichtsleistungen": {}})
                        hj_data = fach["halbjahre"].setdefault(cur_hj, {"noten": {}})
                        hj_data["noten"].setdefault(cur_sk, {"muendlich": [], "schriftlich": []})
                        hj_data["noten"][cur_sk]["schriftlich"] = [int(x.strip()) for x in ns.split(",") if x.strip()]
    except Exception as e:
        logger.error("Migration fehlgeschlagen: %s", e)
        return None

    return old_data


def main() -> None:
    """Hauptfunktion zum Starten der Anwendung."""
    import tkinter as tk

    root = tk.Tk()
    root.title("Notenverwaltung")
    root.geometry("820x510")
    root.minsize(710, 435)

    # Prüfe auf Migration von alter noten.md Datei
    migrated_data = _migrate_old_md()
    first_time = not os.path.exists(DATA_FILE)

    dlg = PasswordDialog(root, title="Passwort setzen" if first_time else "Passwort eingeben",
                         first_time=first_time)
    if dlg.result is None:
        root.destroy()
        return

    password = dlg.result

    # Import app here to avoid circular imports
    from app import NotenVerwaltungApp

    app = NotenVerwaltungApp(root, password)

    # Wenn Migration erfolgreich war, übertrage die Daten
    if migrated_data and migrated_data.schuljahre:
        migrated_data.speichern_verschluesselt(password, app.data_file)
        app.daten = migrated_data
        app._refresh_sj()

    root.mainloop()


if __name__ == "__main__":
    main()