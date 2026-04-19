#!/usr/bin/env python3
"""
Notenverwaltung - Programm zur Verwaltung von Schülerinnennoten
Daten werden verschlüsselt gespeichert (.ndat)
Export als Markdown (.md) oder CSV möglich
Notenschlüssel: BG (0-15) und IHK (1-6)
Klausuren/Unterrichtsleistungen mit Punkte-System, Notenschlüssel-CSV und prozentualer Gewichtung
Fächer mit eigenen Noten und Klausuren pro Klasse
Gewichtung: Jede Klausur/UL hat einen %-Anteil an der Gesamtnote,
der aus dem Kategorie-Pool (UL%/Schriftlich%) kommt.
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import json
import zlib
import base64
import hashlib
import os
import sys
import csv
import logging
from typing import Optional

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

if getattr(sys, 'frozen', False):
    _APP_DIR = os.path.dirname(sys.executable)
else:
    _APP_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_FILE = os.path.join(_APP_DIR, "noten.ndat")
HALBJAHRE = ["1. Halbjahr", "2. Halbjahr"]
DEFAULT_GEWICHTUNG = 60
ITERATIONS = 100000
NOTENSCHLUESSEL = {"IHK": (1, 6), "BG": (0, 15)}
DEFAULT_NS_CSV = {
    "IHK": "100,1;99,1.1;98,1.1;97,1.2;96,1.2;95,1.3;94,1.3;93,1.4;92,1.4;91,1.5;90,1.6;89,1.7;88,1.8;87,1.9;86,2;85,2;84,2.1;83,2.2;82,2.3;81,2.4;80,2.5;79,2.6;78,2.7;77,2.7;76,2.8;75,2.9;74,2.9;73,3;72,3.1;71,3.1;70,3.2;69,3.3;68,3.3;67,3.4;66,3.5;65,3.6;64,3.6;63,3.7;62,3.7;61,3.8;60,3.9;59,3.9;58,4;57,4;56,4.1;55,4.1;54,4.2;53,4.3;52,4.3;51,4.4;50,4.4;49,4.5;48,4.6;47,4.6;46,4.7;45,4.7;44,4.8;43,4.8;42,4.9;41,4.9;40,5;39,5;38,5;37,5.1;36,5.1;35,5.2;34,5.2;33,5.3;32,5.3;31,5.4;30,5.4;29,5.5;28,5.6;27,5.6;26,5.6;25,5.6;24,5.6;23,5.6;22,5.7;21,5.7;20,5.7;19,5.7;18,5.7;17,5.7;16,5.8;15,5.8;14,5.8;13,5.8;12,5.8;11,5.9;10,5.9;9,5.9;8,5.9;7,5.9;6,5.9;5,6;4,6;3,6;2,6;1,6;0,6",
    "BG": "100,15;99,15;98,15;97,15;96,15;95,15;94,14;93,14;92,14;91,14;90,14;89,13;88,13;87,13;86,13;85,13;84,12;83,12;82,12;81,12;80,12;79,11;78,11;77,11;76,11;75,11;74,10;73,10;72,10;71,10;70,10;69,9;68,9;67,9;66,9;65,9;64,8;63,8;62,8;61,8;60,8;59,7;58,7;57,7;56,7;55,7;54,6;53,6;52,6;51,6;50,6;49,5;48,5;47,5;46,5;45,5;44,4;43,4;42,4;41,4;40,4;39,3;38,3;37,3;36,3;35,3;34,3;33,3;32,2;31,2;30,2;29,2;28,2;27,2;26,1;25,1;24,1;23,1;22,1;21,1;20,1;19,0;18,0;17,0;16,0;15,0;14,0;13,0;12,0;11,0;10,0;9,0;8,0;7,0;6,0;5,0;4,0;3,0;2,0;1,0;0,0",
}
_NS_ALIASES = {"Berufsschule": "IHK", "Berufliches Gymnasium": "BG", "IHK": "IHK", "BG": "BG"}

# UI-Konstanten
MSG_WAEHLEN_SJ_KL = "Bitte Schuljahr und Klasse auswählen!"
MSG_WAEHLEN_SJ_KL_FACH = "Bitte Schuljahr, Klasse und Fach auswählen!"
MSG_WAEHLEN_KL_FACH_SK = "Bitte Klasse, Fach und Schülerin auswählen!"
MSG_EXISTIERT_BEREITS = "'{}' existiert bereits!"
MSG_WAEHLEN_ELEMENT = "Bitte ein Element auswählen!"
AUTO_SAVE_MS = 60000  # 60 Sekunden


def _derive_key(password: str, salt: bytes) -> bytes:
    return hashlib.pbkdf2_hmac('sha256', password.encode(), salt, ITERATIONS)

def _xor_encrypt(data_bytes: bytes, key: bytes) -> bytes:
    key_len = len(key)
    return bytes(b ^ key[i % key_len] for i, b in enumerate(data_bytes))

def encrypt_data(data_dict: dict, password: str) -> bytes:
    json_bytes = json.dumps(data_dict, ensure_ascii=False).encode('utf-8')
    compressed = zlib.compress(json_bytes)
    salt = os.urandom(16)
    key = _derive_key(password, salt)
    encrypted = _xor_encrypt(compressed, key)
    verify = hashlib.sha256(key).digest()
    return base64.b64encode(salt + verify + encrypted)

def decrypt_data(raw_bytes: bytes, password: str) -> Optional[dict]:
    try:
        raw = base64.b64decode(raw_bytes)
    except Exception:
        logger.warning("Base64-Dekodierung fehlgeschlagen")
        return None
    if len(raw) < 48:
        logger.warning("Daten zu kurz für Entschlüsselung")
        return None
    salt = raw[:16]
    verify = raw[16:48]
    encrypted = raw[48:]
    key = _derive_key(password, salt)
    if hashlib.sha256(key).digest() != verify:
        logger.warning("Passwort-Verifikation fehlgeschlagen")
        return None
    compressed = _xor_encrypt(encrypted, key)
    try:
        json_bytes = zlib.decompress(compressed)
        return json.loads(json_bytes.decode('utf-8'))
    except (zlib.error, json.JSONDecodeError, UnicodeDecodeError) as e:
        logger.error("Dekompression/Deserialisierung fehlgeschlagen: %s", e)
        return None


# ---------------------------------------------------------------------------
# Datenmodell
# ---------------------------------------------------------------------------
class NotenVerwaltung:
    def __init__(self) -> None:
        self.schuljahre: dict = {}
        self.gewichtung_muendlich: int = DEFAULT_GEWICHTUNG

    # ---- Hilfsmethoden ----
    def _get_klasse(self, sj: str, k: str) -> Optional[dict]:
        return self.schuljahre.get(sj, {}).get(k)

    def get_schueler_dict(self, sj: str, k: str) -> Optional[dict]:
        kl = self._get_klasse(sj, k)
        return kl.get("schuelerinnen", {}) if kl else None

    def _get_fach(self, sj: str, k: str, fach: str) -> Optional[dict]:
        kl = self._get_klasse(sj, k)
        if kl is None:
            return None
        return kl.get("faecher", {}).get(fach)

    def get_notenschluessel(self, sj: str, k: str) -> str:
        kl = self._get_klasse(sj, k)
        ns = kl.get("notenschluessel", "IHK") if kl else "IHK"
        return _NS_ALIASES.get(ns, ns)

    def get_notenbereich(self, sj: str, k: str) -> tuple:
        return NOTENSCHLUESSEL.get(self.get_notenschluessel(sj, k), (1, 6))

    @property
    def ul_prozent(self) -> int:
        return self.gewichtung_muendlich

    @property
    def schriftlich_prozent(self) -> int:
        return 100 - self.gewichtung_muendlich

    # ---- Notenschlüssel CSV ----
    @staticmethod
    def ns_csv_lookup(prozent: int, csv_str: str) -> Optional[float]:
        entries = NotenVerwaltung.ns_csv_parse(csv_str)
        if not entries:
            return None
        for p, n in entries:
            if prozent >= p:
                return n
        return entries[-1][1]

    @staticmethod
    def ns_csv_parse(csv_str: str) -> list:
        if not csv_str:
            return []
        entries = []
        for pair in csv_str.split(";"):
            parts = pair.strip().split(",")
            if len(parts) == 2:
                try:
                    entries.append((float(parts[0].strip()), float(parts[1].strip())))
                except ValueError:
                    pass
        entries.sort(key=lambda x: x[0], reverse=True)
        return entries

    def get_ns_csv(self, sj: str, k: str) -> str:
        kl = self._get_klasse(sj, k)
        if kl is None:
            return DEFAULT_NS_CSV["IHK"]
        csv_str = kl.get("notenschluessel_csv", "")
        if not csv_str:
            ns = _NS_ALIASES.get(kl.get("notenschluessel", "IHK"), "IHK")
            return DEFAULT_NS_CSV.get(ns, DEFAULT_NS_CSV["IHK"])
        return csv_str

    def set_ns_csv(self, sj: str, k: str, csv_str: str) -> None:
        kl = self._get_klasse(sj, k)
        if kl is not None:
            kl["notenschluessel_csv"] = csv_str

    # ---- Serialisierung ----
    def to_dict(self) -> dict:
        sj = {}
        for s, klasses in self.schuljahre.items():
            sj[s] = {}
            for k, kl_data in klasses.items():
                sk_dict = {}
                for sk, d in kl_data.get("schuelerinnen", {}).items():
                    sk_dict[sk] = {"nachname": d["nachname"], "vorname": d["vorname"]}
                faecher = {}
                for fn, fd in kl_data.get("faecher", {}).items():
                    faecher[fn] = {
                        "halbjahre": fd.get("halbjahre", {}),
                        "klausuren": fd.get("klausuren", {}),
                        "unterrichtsleistungen": fd.get("unterrichtsleistungen", {}),
                    }
                sj[s][k] = {
                    "notenschluessel": _NS_ALIASES.get(kl_data.get("notenschluessel", "IHK"),
                                                        kl_data.get("notenschluessel", "IHK")),
                    "notenschluessel_csv": kl_data.get("notenschluessel_csv", ""),
                    "schuelerinnen": sk_dict,
                    "faecher": faecher,
                }
        return {"gewichtung_muendlich": self.gewichtung_muendlich, "schuljahre": sj}

    def from_dict(self, data: dict) -> None:
        self.gewichtung_muendlich = data.get("gewichtung_muendlich", DEFAULT_GEWICHTUNG)
        self.schuljahre = {}
        for s, klasses in data.get("schuljahre", {}).items():
            self.schuljahre[s] = {}
            for k, kl_data in klasses.items():
                if not isinstance(kl_data, dict):
                    continue
                schueler = {}
                ns = _NS_ALIASES.get(kl_data.get("notenschluessel", "IHK"),
                                     kl_data.get("notenschluessel", "IHK"))
                ns_csv = kl_data.get("notenschluessel_csv", "")
                for sk, d in kl_data.get("schuelerinnen", {}).items():
                    if isinstance(d, dict):
                        schueler[sk] = {"nachname": d.get("nachname", ""),
                                        "vorname": d.get("vorname", "")}
                faecher = {}
                for fn, fd in kl_data.get("faecher", {}).items():
                    faecher[fn] = self._parse_fach(fd)
                if not kl_data.get("faecher"):
                    faecher = self._migrate_old_format(kl_data, schueler)
                self.schuljahre[s][k] = {
                    "notenschluessel": ns,
                    "notenschluessel_csv": ns_csv,
                    "schuelerinnen": schueler,
                    "faecher": faecher,
                }
        # BUG-FIX: Auto-Verteilung für alte Daten ohne Gewichtung
        self._auto_distribute_all()

    def _parse_fach(self, fd: dict) -> dict:
        halbjahre = {}
        for hj, hj_data in fd.get("halbjahre", {}).items():
            noten = {}
            for sk, sk_noten in hj_data.get("noten", {}).items():
                noten[sk] = {
                    "muendlich": sk_noten.get("muendlich", []),
                    "schriftlich": sk_noten.get("schriftlich", []),
                }
            halbjahre[hj] = {"noten": noten}
        klausuren = {}
        for hj, klist in fd.get("klausuren", {}).items():
            fixed = []
            for klausur in (klist if isinstance(klist, list) else []):
                if isinstance(klausur, dict):
                    fixed.append({
                        "name": klausur.get("name", ""),
                        "max_punkte_pro_aufgabe": klausur.get("max_punkte_pro_aufgabe", []),
                        "ergebnisse": klausur.get("ergebnisse", {}),
                        "gewichtung": klausur.get("gewichtung", 0),
                    })
            klausuren[hj] = fixed
        unterrichtsleistungen = {}
        for hj, ulist in fd.get("unterrichtsleistungen", {}).items():
            fixed = []
            for ul in (ulist if isinstance(ulist, list) else []):
                if isinstance(ul, dict):
                    fixed.append({
                        "name": ul.get("name", ""),
                        "max_punkte_pro_aufgabe": ul.get("max_punkte_pro_aufgabe", []),
                        "ergebnisse": ul.get("ergebnisse", {}),
                        "gewichtung": ul.get("gewichtung", 0),
                    })
            unterrichtsleistungen[hj] = fixed
        return {"halbjahre": halbjahre, "klausuren": klausuren,
                "unterrichtsleistungen": unterrichtsleistungen}

    def _migrate_old_format(self, kl_data: dict, schueler: dict) -> dict:
        fach = {"halbjahre": {}, "klausuren": {}, "unterrichtsleistungen": {}}
        for sk, d in (kl_data.get("schuelerinnen", kl_data) if isinstance(kl_data, dict) else {}).items():
            if not isinstance(d, dict):
                continue
            for hj in HALBJAHRE:
                hj_data = d.get("halbjahre", {}).get(hj, {})
                if hj not in fach["halbjahre"]:
                    fach["halbjahre"][hj] = {"noten": {}}
                fach["halbjahre"][hj]["noten"][sk] = {
                    "muendlich": hj_data.get("muendlich", []),
                    "schriftlich": hj_data.get("schriftlich", []),
                }
        for hj, klist in kl_data.get("klausuren", {}).items():
            fixed = []
            for klausur in (klist if isinstance(klist, list) else []):
                if isinstance(klausur, dict):
                    fixed.append({
                        "name": klausur.get("name", ""),
                        "max_punkte_pro_aufgabe": klausur.get("max_punkte_pro_aufgabe", []),
                        "ergebnisse": klausur.get("ergebnisse", {}),
                        "gewichtung": klausur.get("gewichtung", 0),
                    })
            fach["klausuren"][hj] = fixed
        return {"Allgemein": fach}

    def _auto_distribute_all(self) -> None:
        """Auto-Verteilung für alle Klausuren/ULs mit Gewichtung=0 nach dem Laden."""
        for sj, klasses in self.schuljahre.items():
            for k, kl_data in klasses.items():
                for fn, fach in kl_data.get("faecher", {}).items():
                    for hj in HALBJAHRE:
                        klausuren = fach.get("klausuren", {}).get(hj, [])
                        if klausuren and all(kl.get("gewichtung", 0) == 0 for kl in klausuren):
                            self._auto_distribute_klausuren(sj, k, fn, hj)
                        uls = fach.get("unterrichtsleistungen", {}).get(hj, [])
                        if uls and all(ul.get("gewichtung", 0) == 0 for ul in uls):
                            self._auto_distribute_ul(sj, k, fn, hj)

    def speichern_verschluesselt(self, password: str, filepath: Optional[str] = None) -> None:
        encrypted = encrypt_data(self.to_dict(), password)
        fp = filepath or DATA_FILE
        tmp_file = fp + ".tmp"
        try:
            with open(tmp_file, "wb") as f:
                f.write(encrypted)
            os.replace(tmp_file, fp)
            logger.info("Daten gespeichert: %s", fp)
        except OSError as e:
            logger.error("Fehler beim Speichern: %s", e)
            raise

    def laden_verschluesselt(self, password: str, filepath: Optional[str] = None) -> bool:
        fp = filepath or DATA_FILE
        if not os.path.exists(fp):
            return True
        try:
            with open(fp, "rb") as f:
                raw = f.read()
        except OSError as e:
            logger.error("Fehler beim Laden: %s", e)
            return False
        data = decrypt_data(raw, password)
        if data is None:
            return False
        self.from_dict(data)
        logger.info("Daten geladen: %s", fp)
        return True

    # ---- Export ----
    def export_markdown(self, filepath: str) -> None:
        z = ["# Notenverwaltung", "",
             f"Gewichtung Unterrichtsleistung: {self.ul_prozent}%",
             f"Gewichtung Schriftlich: {self.schriftlich_prozent}%", ""]
        for sj in sorted(self.schuljahre):
            z.append(f"## Schuljahr {sj}")
            z.append("")
            for kn in sorted(self.schuljahre[sj]):
                ns = self.get_notenschluessel(sj, kn)
                nb = self.get_notenbereich(sj, kn)
                z.append(f"### Klasse {kn} [{ns} (Noten {nb[0]}-{nb[1]})]")
                z.append("")
                for fn in sorted(self.schuljahre[sj][kn].get("faecher", {})):
                    z.append(f"#### Fach: {fn}")
                    z.append("")
                    fach = self.schuljahre[sj][kn]["faecher"][fn]
                    for hj in HALBJAHRE:
                        klausuren = fach.get("klausuren", {}).get(hj, [])
                        if klausuren:
                            z.append(f"**Klausuren {hj}:**")
                            for kl in klausuren:
                                max_p = sum(kl["max_punkte_pro_aufgabe"])
                                gw = kl.get("gewichtung", 0)
                                z.append(f"- {kl['name']} ({gw}% der Gesamtnote, max. {max_p} Punkte)")
                            z.append("")
                        uls = fach.get("unterrichtsleistungen", {}).get(hj, [])
                        if uls:
                            z.append(f"**Unterrichtsleistungen {hj}:**")
                            for ul in uls:
                                max_p = sum(ul["max_punkte_pro_aufgabe"])
                                gw = ul.get("gewichtung", 0)
                                z.append(f"- {ul['name']} ({gw}% der Gesamtnote, max. {max_p} Punkte)")
                            z.append("")
                    for sk in self.schuelerin_sortiert(sj, kn):
                        d = self.schuljahre[sj][kn]["schuelerinnen"][sk]
                        for hj in HALBJAHRE:
                            noten = fach.get("halbjahre", {}).get(hj, {}).get("noten", {}).get(sk, {})
                            m = noten.get("muendlich", [])
                            s_manual = noten.get("schriftlich", [])
                            kn_notes = self.get_klausur_noten_gewichtet(sj, kn, fn, sk, hj)
                            ul_notes = self.get_ul_noten_gewichtet(sj, kn, fn, sk, hj)
                            if m or s_manual or kn_notes or ul_notes:
                                sum_kl_gw = sum(kl.get("gewichtung", 0) for kl in klausuren)
                                sum_ul_gw = sum(ul.get("gewichtung", 0) for ul in uls)
                                remaining_ul = max(0, self.ul_prozent - sum_ul_gw)
                                remaining_schr = max(0, self.schriftlich_prozent - sum_kl_gw)
                                z.append(f"**{d['nachname']}, {d['vorname']}** – {fn} – {hj}")
                                if m:
                                    z.append(f"- Unterrichtsleistung manuell ({remaining_ul}%): {', '.join(str(n) for n in m)}")
                                if ul_notes:
                                    z.append(f"- Unterrichtsleistung bewertet: {', '.join(f'{n} ({g}%)' for n, g in ul_notes)}")
                                if s_manual:
                                    z.append(f"- Schriftlich manuell ({remaining_schr}%): {', '.join(str(n) for n in s_manual)}")
                                if kn_notes:
                                    z.append(f"- Schriftlich Klausuren: {', '.join(f'{n} ({g}%)' for n, g in kn_notes)}")
                                gn = self.gesamtnote_hj(sj, kn, fn, sk, hj)
                                if gn is not None:
                                    z.append(f"- **Gesamtnote: {gn:.2f}**")
                                z.append("")
                    z.append("")
                z.append("")
            z.append("")
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write("\n".join(z))
            logger.info("Markdown exportiert: %s", filepath)
        except OSError as e:
            logger.error("Markdown-Export fehlgeschlagen: %s", e)
            raise

    def export_csv(self, filepath: str) -> None:
        try:
            with open(filepath, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["Schuljahr", "Klasse", "Fach", "Notenschlüssel", "Nachname", "Vorname",
                            "Halbjahr", "UL manuell", "UL bewertet", "Schriftlich manuell",
                            "Schriftlich Klausuren", "Gesamtnote Halbjahr"])
                for sj in sorted(self.schuljahre):
                    for kn in sorted(self.schuljahre[sj]):
                        ns = self.get_notenschluessel(sj, kn)
                        for fn in sorted(self.schuljahre[sj][kn].get("faecher", {})):
                            for sk in self.schuelerin_sortiert(sj, kn):
                                d = self.schuljahre[sj][kn]["schuelerinnen"][sk]
                                for hj in HALBJAHRE:
                                    gn = self.gesamtnote_hj(sj, kn, fn, sk, hj)
                                    w.writerow([
                                        sj, kn, fn, ns, d["nachname"], d["vorname"], hj,
                                        " | ".join(str(n) for n in self.get_muendlich(sj, kn, fn, sk, hj)),
                                        " | ".join(f"{n}({g}%)" for n, g in self.get_ul_noten_gewichtet(sj, kn, fn, sk, hj)),
                                        " | ".join(str(n) for n in self.get_schriftlich(sj, kn, fn, sk, hj)),
                                        " | ".join(f"{n}({g}%)" for n, g in self.get_klausur_noten_gewichtet(sj, kn, fn, sk, hj)),
                                        f"{gn:.2f}" if gn else "-",
                                    ])
            logger.info("CSV exportiert: %s", filepath)
        except OSError as e:
            logger.error("CSV-Export fehlgeschlagen: %s", e)
            raise

    # ---- CRUD Schuljahr / Klasse / Schülerin ----
    def schuljahr_hinzufuegen(self, s: str) -> bool:
        s = s.strip()
        if not s or s in self.schuljahre:
            return False
        self.schuljahre[s] = {}
        return True

    def schuljahr_loeschen(self, s: str) -> bool:
        if s in self.schuljahre:
            del self.schuljahre[s]
            return True
        return False

    def klasse_hinzufuegen(self, sj: str, k: str, notenschluessel: str = "IHK") -> bool:
        k = k.strip()
        if not k or sj not in self.schuljahre or k in self.schuljahre[sj]:
            return False
        self.schuljahre[sj][k] = {
            "notenschluessel": notenschluessel,
            "notenschluessel_csv": "",
            "schuelerinnen": {},
            "faecher": {},
        }
        return True

    def klasse_loeschen(self, sj: str, k: str) -> bool:
        if sj in self.schuljahre and k in self.schuljahre[sj]:
            del self.schuljahre[sj][k]
            return True
        return False

    def klasse_uebertragen(self, sj_quelle: str, k: str, sj_ziel: str) -> bool:
        kl = self._get_klasse(sj_quelle, k)
        if kl is None or sj_ziel not in self.schuljahre or k in self.schuljahre[sj_ziel]:
            return False
        schueler = {sk: {"nachname": d["nachname"], "vorname": d["vorname"]}
                    for sk, d in kl.get("schuelerinnen", {}).items()}
        faecher = {}
        for fn in kl.get("faecher", {}):
            faecher[fn] = {"halbjahre": {h: {"noten": {}} for h in HALBJAHRE},
                           "klausuren": {}, "unterrichtsleistungen": {}}
        self.schuljahre[sj_ziel][k] = {
            "notenschluessel": kl.get("notenschluessel", "IHK"),
            "notenschluessel_csv": kl.get("notenschluessel_csv", ""),
            "schuelerinnen": schueler,
            "faecher": faecher,
        }
        return True

    @staticmethod
    def _key(nn: str, vn: str) -> str:
        return f"{nn}, {vn}"

    def schuelerin_hinzufuegen(self, sj: str, k: str, nn: str, vn: str) -> bool:
        nn, vn = nn.strip(), vn.strip()
        if not nn or not vn:
            return False
        sd = self.get_schueler_dict(sj, k)
        if sd is None:
            return False
        key = self._key(nn, vn)
        if key in sd:
            return False
        sd[key] = {"nachname": nn, "vorname": vn}
        return True

    def schuelerin_loeschen(self, sj: str, k: str, sk: str) -> bool:
        sd = self.get_schueler_dict(sj, k)
        if sd is not None and sk in sd:
            del sd[sk]
            return True
        return False

    def schuelerin_sortiert(self, sj: str, k: str) -> list:
        sd = self.get_schueler_dict(sj, k)
        if sd is None:
            return []
        return sorted(sd.keys(), key=lambda x: (sd[x]["nachname"].lower(), sd[x]["vorname"].lower()))

    # ---- CRUD Fächer ----
    def fach_hinzufuegen(self, sj: str, k: str, fach: str) -> bool:
        fach = fach.strip()
        if not fach:
            return False
        kl = self._get_klasse(sj, k)
        if kl is None:
            return False
        if "faecher" not in kl:
            kl["faecher"] = {}
        if fach in kl["faecher"]:
            return False
        kl["faecher"][fach] = {
            "halbjahre": {h: {"noten": {}} for h in HALBJAHRE},
            "klausuren": {},
            "unterrichtsleistungen": {},
        }
        return True

    def fach_loeschen(self, sj: str, k: str, fach: str) -> bool:
        kl = self._get_klasse(sj, k)
        if kl is None:
            return False
        faecher = kl.get("faecher", {})
        if fach in faecher:
            del faecher[fach]
            return True
        return False

    def fach_sortiert(self, sj: str, k: str) -> list:
        kl = self._get_klasse(sj, k)
        if kl is None:
            return []
        return sorted(kl.get("faecher", {}).keys())

    # ---- Noten (pro Fach) – öffentliche API ----
    def get_muendlich(self, sj: str, k: str, fach: str, sk: str, hj: str) -> list:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return []
        return f.get("halbjahre", {}).get(hj, {}).get("noten", {}).get(sk, {}).get("muendlich", [])

    def get_schriftlich(self, sj: str, k: str, fach: str, sk: str, hj: str) -> list:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return []
        return f.get("halbjahre", {}).get(hj, {}).get("noten", {}).get(sk, {}).get("schriftlich", [])

    def _ensure_noten_dict(self, sj: str, k: str, fach: str, sk: str, hj: str) -> Optional[dict]:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return None
        hj_data = f.setdefault("halbjahre", {}).setdefault(hj, {"noten": {}})
        if "noten" not in hj_data:
            hj_data["noten"] = {}
        sk_noten = hj_data["noten"].setdefault(sk, {"muendlich": [], "schriftlich": []})
        if "muendlich" not in sk_noten:
            sk_noten["muendlich"] = []
        if "schriftlich" not in sk_noten:
            sk_noten["schriftlich"] = []
        return sk_noten

    def note_hinzufuegen(self, sj: str, k: str, fach: str, sk: str, hj: str,
                         typ: str, note: float) -> bool:
        nb = self.get_notenbereich(sj, k)
        if not (nb[0] <= note <= nb[1]):
            return False
        sk_noten = self._ensure_noten_dict(sj, k, fach, sk, hj)
        if sk_noten is not None:
            sk_noten[typ].append(note)
            return True
        return False

    def note_loeschen(self, sj: str, k: str, fach: str, sk: str, hj: str,
                      typ: str, idx: int) -> bool:
        sk_noten = self._ensure_noten_dict(sj, k, fach, sk, hj)
        if sk_noten is not None:
            n = sk_noten[typ]
            if 0 <= idx < len(n):
                n.pop(idx)
                return True
        return False

    # ---- Gewichtung-Verteilung ----
    def _auto_distribute_klausuren(self, sj: str, k: str, fach: str, hj: str) -> None:
        klausuren = self.get_klausuren(sj, k, fach, hj)
        if not klausuren:
            return
        each = self.schriftlich_prozent / len(klausuren)
        for kl in klausuren:
            kl["gewichtung"] = round(each, 1)
        total = sum(kl["gewichtung"] for kl in klausuren)
        diff = self.schriftlich_prozent - total
        if diff != 0 and klausuren:
            klausuren[-1]["gewichtung"] = round(klausuren[-1]["gewichtung"] + diff, 1)

    def _auto_distribute_ul(self, sj: str, k: str, fach: str, hj: str) -> None:
        uls = self.get_unterrichtsleistungen(sj, k, fach, hj)
        if not uls:
            return
        each = self.ul_prozent / len(uls)
        for ul in uls:
            ul["gewichtung"] = round(each, 1)
        total = sum(ul["gewichtung"] for ul in uls)
        diff = self.ul_prozent - total
        if diff != 0 and uls:
            uls[-1]["gewichtung"] = round(uls[-1]["gewichtung"] + diff, 1)

    def get_remaining_ul_pct(self, sj: str, k: str, fach: str, hj: str) -> float:
        uls = self.get_unterrichtsleistungen(sj, k, fach, hj)
        used = sum(ul.get("gewichtung", 0) for ul in uls)
        return max(0, self.ul_prozent - used)

    def get_remaining_schriftlich_pct(self, sj: str, k: str, fach: str, hj: str) -> float:
        klausuren = self.get_klausuren(sj, k, fach, hj)
        used = sum(kl.get("gewichtung", 0) for kl in klausuren)
        return max(0, self.schriftlich_prozent - used)

    def get_total_klausur_gewichtung(self, sj: str, k: str, fach: str, hj: str,
                                      exclude_idx: int = -1) -> float:
        """Summe der Klausur-Gewichtungen, optional ohne einen Index."""
        klausuren = self.get_klausuren(sj, k, fach, hj)
        return sum(kl.get("gewichtung", 0) for i, kl in enumerate(klausuren) if i != exclude_idx)

    def get_total_ul_gewichtung(self, sj: str, k: str, fach: str, hj: str,
                                 exclude_idx: int = -1) -> float:
        """Summe der UL-Gewichtungen, optional ohne einen Index."""
        uls = self.get_unterrichtsleistungen(sj, k, fach, hj)
        return sum(ul.get("gewichtung", 0) for i, ul in enumerate(uls) if i != exclude_idx)

    # ---- CRUD Klausuren (pro Fach) ----
    def klausur_hinzufuegen(self, sj: str, k: str, fach: str, hj: str,
                            name: str, max_punkte_pro_aufgabe: list) -> bool:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return False
        if "klausuren" not in f:
            f["klausuren"] = {}
        if hj not in f["klausuren"]:
            f["klausuren"][hj] = []
        for klausur in f["klausuren"][hj]:
            if klausur["name"] == name:
                return False
        f["klausuren"][hj].append({
            "name": name,
            "max_punkte_pro_aufgabe": max_punkte_pro_aufgabe,
            "ergebnisse": {},
            "gewichtung": 0,
        })
        self._auto_distribute_klausuren(sj, k, fach, hj)
        return True

    def klausur_loeschen(self, sj: str, k: str, fach: str, hj: str, idx: int) -> bool:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return False
        klist = f.get("klausuren", {}).get(hj, [])
        if 0 <= idx < len(klist):
            klist.pop(idx)
            self._auto_distribute_klausuren(sj, k, fach, hj)
            return True
        return False

    def get_klausuren(self, sj: str, k: str, fach: str, hj: str) -> list:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return []
        return f.get("klausuren", {}).get(hj, [])

    def klausur_punkte_setzen(self, sj: str, k: str, fach: str, hj: str,
                               kidx: int, sk: str, punkte: list) -> bool:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return False
        klist = f.get("klausuren", {}).get(hj, [])
        if not (0 <= kidx < len(klist)):
            return False
        klausur = klist[kidx]
        if len(punkte) != len(klausur["max_punkte_pro_aufgabe"]):
            return False
        for i, p in enumerate(punkte):
            if p is not None and (p < 0 or p > klausur["max_punkte_pro_aufgabe"][i]):
                return False
        klausur["ergebnisse"][sk] = punkte
        return True

    def klausur_gewichtung_setzen(self, sj: str, k: str, fach: str, hj: str,
                                   kidx: int, gewichtung: float) -> bool:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return False
        klist = f.get("klausuren", {}).get(hj, [])
        if not (0 <= kidx < len(klist)):
            return False
        if gewichtung < 0:
            return False
        klist[kidx]["gewichtung"] = gewichtung
        return True

    @staticmethod
    def _round_pct(prozent: float) -> int:
        return int(prozent + 0.5)

    def klausur_note_berechnen(self, sj: str, k: str, fach: str, hj: str,
                                kidx: int, sk: str) -> Optional[float]:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return None
        klist = f.get("klausuren", {}).get(hj, [])
        if not (0 <= kidx < len(klist)):
            return None
        klausur = klist[kidx]
        if sk not in klausur["ergebnisse"]:
            return None
        punkte = klausur["ergebnisse"][sk]
        if any(p is None for p in punkte):
            return None
        max_p = sum(klausur["max_punkte_pro_aufgabe"])
        if max_p == 0:
            return None
        prozent = self._round_pct(sum(punkte) / max_p * 100)
        return self.ns_csv_lookup(prozent, self.get_ns_csv(sj, k))

    def klausur_durchschnitt_berechnen(self, sj: str, k: str, fach: str, hj: str,
                                        kidx: int) -> Optional[float]:
        """Berechnet den Durchschnitt aller Schülerinnennoten für eine Klausur."""
        f = self._get_fach(sj, k, fach)
        if f is None:
            return None
        klist = f.get("klausuren", {}).get(hj, [])
        if not (0 <= kidx < len(klist)):
            return None
        klausur = klist[kidx]
        ergebnisse = klausur.get("ergebnisse", {})

        noten = []
        for sk in ergebnisse:
            punkte = ergebnisse[sk]
            if any(p is None for p in punkte):
                continue
            max_p = sum(klausur["max_punkte_pro_aufgabe"])
            if max_p == 0:
                continue
            prozent = self._round_pct(sum(punkte) / max_p * 100)
            note = self.ns_csv_lookup(prozent, self.get_ns_csv(sj, k))
            if note is not None:
                noten.append(note)

        return round(sum(noten) / len(noten), 2) if noten else None

    def klausur_nicht_bestanden_count(self, sj: str, k: str, fach: str, hj: str,
                                       kidx: int) -> tuple:
        """Zählt die nicht bestandenen Noten für eine Klausur.

        Returns:
            tuple: (anzahl_nicht_bestanden, gesamt_anzahl, prozent_nicht_bestanden, warnung)
        """
        f = self._get_fach(sj, k, fach)
        if f is None:
            return (0, 0, 0.0, False)
        klist = f.get("klausuren", {}).get(hj, [])
        if not (0 <= kidx < len(klist)):
            return (0, 0, 0.0, False)

        klausur = klist[kidx]
        ergebnisse = klausur.get("ergebnisse", {})
        ns = self.get_notenschluessel(sj, k)

        nicht_bestanden = 0
        gesamt = 0

        for sk in ergebnisse:
            punkte = ergebnisse[sk]
            if any(p is None for p in punkte):
                continue
            max_p = sum(klausur["max_punkte_pro_aufgabe"])
            if max_p == 0:
                continue

            prozent = self._round_pct(sum(punkte) / max_p * 100)
            note = self.ns_csv_lookup(prozent, self.get_ns_csv(sj, k))

            if note is not None:
                gesamt += 1
                # IHK: Note > 4.5 ist nicht bestanden
                # BG: Note < 4 ist nicht bestanden
                if ns == "BG":
                    if note < 4:
                        nicht_bestanden += 1
                else:  # IHK
                    if note > 4.5:
                        nicht_bestanden += 1

        prozent = (nicht_bestanden / gesamt * 100) if gesamt > 0 else 0.0
        warnung = prozent > 30.0

        return (nicht_bestanden, gesamt, round(prozent, 1), warnung)

    def get_klausur_noten_gewichtet(self, sj: str, k: str, fach: str,
                                     sk: str, hj: str) -> list:
        klausuren = self.get_klausuren(sj, k, fach, hj)
        result = []
        for i in range(len(klausuren)):
            note = self.klausur_note_berechnen(sj, k, fach, hj, i, sk)
            if note is not None:
                result.append((note, klausuren[i].get("gewichtung", 0)))
        return result

    # ---- CRUD Unterrichtsleistungen (pro Fach) ----
    def ul_hinzufuegen(self, sj: str, k: str, fach: str, hj: str,
                       name: str, max_punkte_pro_aufgabe: list) -> bool:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return False
        if "unterrichtsleistungen" not in f:
            f["unterrichtsleistungen"] = {}
        if hj not in f["unterrichtsleistungen"]:
            f["unterrichtsleistungen"][hj] = []
        for ul in f["unterrichtsleistungen"][hj]:
            if ul["name"] == name:
                return False
        f["unterrichtsleistungen"][hj].append({
            "name": name,
            "max_punkte_pro_aufgabe": max_punkte_pro_aufgabe,
            "ergebnisse": {},
            "gewichtung": 0,
        })
        self._auto_distribute_ul(sj, k, fach, hj)
        return True

    def ul_loeschen(self, sj: str, k: str, fach: str, hj: str, idx: int) -> bool:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return False
        ulist = f.get("unterrichtsleistungen", {}).get(hj, [])
        if 0 <= idx < len(ulist):
            ulist.pop(idx)
            self._auto_distribute_ul(sj, k, fach, hj)
            return True
        return False

    def get_unterrichtsleistungen(self, sj: str, k: str, fach: str, hj: str) -> list:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return []
        return f.get("unterrichtsleistungen", {}).get(hj, [])

    def ul_punkte_setzen(self, sj: str, k: str, fach: str, hj: str,
                          ulidx: int, sk: str, punkte: list) -> bool:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return False
        ulist = f.get("unterrichtsleistungen", {}).get(hj, [])
        if not (0 <= ulidx < len(ulist)):
            return False
        ul = ulist[ulidx]
        if len(punkte) != len(ul["max_punkte_pro_aufgabe"]):
            return False
        for i, p in enumerate(punkte):
            if p is not None and (p < 0 or p > ul["max_punkte_pro_aufgabe"][i]):
                return False
        ul["ergebnisse"][sk] = punkte
        return True

    def ul_gewichtung_setzen(self, sj: str, k: str, fach: str, hj: str,
                              ulidx: int, gewichtung: float) -> bool:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return False
        ulist = f.get("unterrichtsleistungen", {}).get(hj, [])
        if not (0 <= ulidx < len(ulist)):
            return False
        if gewichtung < 0:
            return False
        ulist[ulidx]["gewichtung"] = gewichtung
        return True

    def ul_note_berechnen(self, sj: str, k: str, fach: str, hj: str,
                           ulidx: int, sk: str) -> Optional[float]:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return None
        ulist = f.get("unterrichtsleistungen", {}).get(hj, [])
        if not (0 <= ulidx < len(ulist)):
            return None
        ul = ulist[ulidx]
        if sk not in ul["ergebnisse"]:
            return None
        punkte = ul["ergebnisse"][sk]
        if any(p is None for p in punkte):
            return None
        max_p = sum(ul["max_punkte_pro_aufgabe"])
        if max_p == 0:
            return None
        prozent = self._round_pct(sum(punkte) / max_p * 100)
        return self.ns_csv_lookup(prozent, self.get_ns_csv(sj, k))

    def get_ul_noten_gewichtet(self, sj: str, k: str, fach: str,
                                sk: str, hj: str) -> list:
        uls = self.get_unterrichtsleistungen(sj, k, fach, hj)
        result = []
        for i in range(len(uls)):
            note = self.ul_note_berechnen(sj, k, fach, hj, i, sk)
            if note is not None:
                result.append((note, uls[i].get("gewichtung", 0)))
        return result

    # ---- Berechnungen ----
    @staticmethod
    def durchschnitt(noten: list) -> Optional[float]:
        return round(sum(noten) / len(noten), 2) if noten else None

    def gesamtnote_hj(self, sj: str, k: str, fach: str, sk: str, hj: str) -> Optional[float]:
        total = 0.0
        has_any = False

        # Bewertete ULs
        uls = self.get_unterrichtsleistungen(sj, k, fach, hj)
        for i, ul in enumerate(uls):
            note = self.ul_note_berechnen(sj, k, fach, hj, i, sk)
            gw = ul.get("gewichtung", 0)
            if note is not None and gw > 0:
                total += note * (gw / 100)
                has_any = True

        # Manuelle UL-Noten
        manual_ul = self.get_muendlich(sj, k, fach, sk, hj)
        remaining_ul = self.get_remaining_ul_pct(sj, k, fach, hj)
        if manual_ul and remaining_ul > 0:
            avg = sum(manual_ul) / len(manual_ul)
            total += avg * (remaining_ul / 100)
            has_any = True

        # Bewertete Klausuren
        klausuren = self.get_klausuren(sj, k, fach, hj)
        for i, kl in enumerate(klausuren):
            note = self.klausur_note_berechnen(sj, k, fach, hj, i, sk)
            gw = kl.get("gewichtung", 0)
            if note is not None and gw > 0:
                total += note * (gw / 100)
                has_any = True

        # Manuelle schriftliche Noten
        manual_schr = self.get_schriftlich(sj, k, fach, sk, hj)
        remaining_schr = self.get_remaining_schriftlich_pct(sj, k, fach, hj)
        if manual_schr and remaining_schr > 0:
            avg = sum(manual_schr) / len(manual_schr)
            total += avg * (remaining_schr / 100)
            has_any = True

        return round(total, 2) if has_any else None

    def gesamtnote_jahr(self, sj: str, k: str, fach: str, sk: str) -> Optional[float]:
        notes = []
        for hj in HALBJAHRE:
            gn = self.gesamtnote_hj(sj, k, fach, sk, hj)
            if gn is not None:
                notes.append(gn)
        return round(sum(notes) / len(notes), 2) if notes else None


# ---------------------------------------------------------------------------
# Dialoge
# ---------------------------------------------------------------------------
class _CenteredToplevel(tk.Toplevel):
    def _center(self) -> None:
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw - self.winfo_width()) // 2}+{(sh - self.winfo_height()) // 2}")


class PasswordDialog(_CenteredToplevel):
    def __init__(self, parent, title: str = "Passwort eingeben", first_time: bool = False):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result: Optional[str] = None
        self.transient(parent)
        self.grab_set()
        f = ttk.Frame(self, padding=15)
        f.pack()
        msg = "Bitte vergeben Sie ein Passwort:" if first_time else "Bitte Passwort eingeben:"
        ttk.Label(f, text=msg).grid(row=0, column=0, columnspan=2, pady=(0, 10))
        ttk.Label(f, text="Passwort:").grid(row=1, column=0, sticky="w", pady=(0, 5))
        self.pw = ttk.Entry(f, show="*", width=25)
        self.pw.grid(row=1, column=1, pady=(0, 5), padx=(5, 0))
        self.pw.focus_set()
        self.pw2 = None
        if first_time:
            ttk.Label(f, text="Bestätigen:").grid(row=2, column=0, sticky="w", pady=(0, 10))
            self.pw2 = ttk.Entry(f, show="*", width=25)
            self.pw2.grid(row=2, column=1, pady=(0, 10), padx=(5, 0))
        bf = ttk.Frame(f)
        bf.grid(row=3, column=0, columnspan=2)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.LEFT, padx=5)
        self.pw.bind("<Return>", lambda e: self.pw2.focus_set() if self.pw2 else self._ok())
        if self.pw2:
            self.pw2.bind("<Return>", lambda e: self._ok())
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _ok(self) -> None:
        pw = self.pw.get()
        if not pw:
            return
        if self.pw2 and pw != self.pw2.get():
            messagebox.showwarning("Fehler", "Passwörter stimmen nicht überein!", parent=self)
            return
        self.result = pw
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class KlasseDialog(_CenteredToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Klasse hinzufügen")
        self.resizable(False, False)
        self.result = None
        self.transient(parent)
        self.grab_set()
        f = ttk.Frame(self, padding=15)
        f.pack()
        ttk.Label(f, text="Klassenname:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.e_name = ttk.Entry(f, width=25)
        self.e_name.grid(row=0, column=1, pady=(0, 5), padx=(5, 0))
        self.e_name.focus_set()
        ttk.Label(f, text="Notenschlüssel:").grid(row=1, column=0, sticky="w", pady=(0, 5))
        self.ns_var = tk.StringVar(value="IHK")
        ns_frame = ttk.Frame(f)
        ns_frame.grid(row=1, column=1, sticky="w", pady=(0, 5), padx=(5, 0))
        for ns, nb in NOTENSCHLUESSEL.items():
            ttk.Radiobutton(ns_frame, text=f"{ns} (Noten {nb[0]}-{nb[1]})",
                            variable=self.ns_var, value=ns).pack(anchor="w")
        ttk.Label(f, text="BG = Berufliches Gymnasium (0-15)\nIHK = Berufsschule (1-6)",
                  foreground="gray").grid(row=2, column=0, columnspan=2, pady=(0, 5))
        bf = ttk.Frame(f)
        bf.grid(row=3, column=0, columnspan=2)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.LEFT, padx=5)
        self.e_name.bind("<Return>", lambda e: self._ok())
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _ok(self) -> None:
        name = self.e_name.get().strip()
        ns = self.ns_var.get()
        if name:
            self.result = (name, ns)
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class FachDialog(_CenteredToplevel):
    def __init__(self, parent, title: str = "Fach hinzufügen"):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.transient(parent)
        self.grab_set()
        f = ttk.Frame(self, padding=15)
        f.pack()
        ttk.Label(f, text="Fachname:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.e_name = ttk.Entry(f, width=30)
        self.e_name.grid(row=0, column=1, pady=(0, 5), padx=(5, 0))
        self.e_name.focus_set()
        bf = ttk.Frame(f)
        bf.grid(row=1, column=0, columnspan=2)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.LEFT, padx=5)
        self.e_name.bind("<Return>", lambda e: self._ok())
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _ok(self) -> None:
        name = self.e_name.get().strip()
        if name:
            self.result = name
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class SchuelerinDialog(_CenteredToplevel):
    def __init__(self, parent, title: str = "Schülerin hinzufügen"):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.transient(parent)
        self.grab_set()
        f = ttk.Frame(self, padding=15)
        f.pack()
        ttk.Label(f, text="Nachname:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.e_nn = ttk.Entry(f, width=30)
        self.e_nn.grid(row=0, column=1, pady=(0, 5), padx=(5, 0))
        self.e_nn.focus_set()
        ttk.Label(f, text="Vorname:").grid(row=1, column=0, sticky="w", pady=(0, 10))
        self.e_vn = ttk.Entry(f, width=30)
        self.e_vn.grid(row=1, column=1, pady=(0, 10), padx=(5, 0))
        bf = ttk.Frame(f)
        bf.grid(row=2, column=0, columnspan=2)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.LEFT, padx=5)
        self.e_nn.bind("<Return>", lambda e: self.e_vn.focus_set())
        self.e_vn.bind("<Return>", lambda e: self._ok())
        self._center()
        self.wait_window()

    def _ok(self) -> None:
        nn, vn = self.e_nn.get().strip(), self.e_vn.get().strip()
        if nn and vn:
            self.result = (nn, vn)
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class SchuelerlisteDialog(_CenteredToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Schülerliste hinzufügen")
        self.geometry("500x450")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.result = None
        f = ttk.Frame(self, padding=15)
        f.pack(fill=tk.BOTH, expand=True)
        ttk.Label(f, text="Schüler als Liste eingeben",
                  font=("TkDefaultFont", 11, "bold")).pack(anchor="w")
        ttk.Label(f, text="Format: Nachname, Vorname (eine Schülerin pro Zeile)",
                  foreground="gray").pack(anchor="w", pady=(0, 10))
        self.text = tk.Text(f, height=15, width=50, font=("Courier", 10))
        self.text.pack(fill=tk.BOTH, expand=True, pady=(5, 10))
        ttk.Label(f, text="Tipp: Liste kann aus Excel/CSV kopiert werden",
                  foreground="gray", font=("TkDefaultFont", 8)).pack(anchor="w")
        bf = ttk.Frame(f)
        bf.pack(fill=tk.X, pady=(15, 0))
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.RIGHT, padx=5)
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _ok(self) -> None:
        content = self.text.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Warnung", "Bitte mindestens eine Schülerin eingeben!", parent=self)
            return
        schueler_liste = []
        for line in content.split("\n"):
            line = line.strip()
            if not line:
                continue
            if "," in line:
                parts = line.split(",", 1)
                nn = parts[0].strip()
                vn = parts[1].strip() if len(parts) > 1 else ""
            elif "\t" in line:
                parts = line.split("\t", 1)
                nn = parts[0].strip()
                vn = parts[1].strip() if len(parts) > 1 else ""
            else:
                nn = line
                vn = ""
            if nn:
                schueler_liste.append((nn, vn))
        if not schueler_liste:
            return
        self.result = schueler_liste
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class KlausurDialog(_CenteredToplevel):
    def __init__(self, parent, title: str = "Klausur hinzufügen"):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.transient(parent)
        self.grab_set()
        f = ttk.Frame(self, padding=15)
        f.pack()
        ttk.Label(f, text="Name:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.e_name = ttk.Entry(f, width=25)
        self.e_name.grid(row=0, column=1, pady=(0, 5), padx=(5, 0))
        self.e_name.focus_set()
        ttk.Label(f, text="Anzahl Aufgaben:").grid(row=1, column=0, sticky="w", pady=(0, 5))
        self.e_anz = ttk.Spinbox(f, from_=1, to=20, width=5)
        self.e_anz.grid(row=1, column=1, sticky="w", pady=(0, 5), padx=(5, 0))
        self.e_anz.set(3)
        ttk.Label(f, text="Max. Punkte pro Aufgabe\n(kommagetrennt):",
                  foreground="gray").grid(row=2, column=0, columnspan=2, sticky="w", pady=(5, 2))
        self.e_punkte = ttk.Entry(f, width=25)
        self.e_punkte.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        self.e_punkte.insert(0, "10,10,10")
        self.e_anz.bind("<Return>", lambda e: self.e_punkte.focus_set())
        self.e_punkte.bind("<Return>", lambda e: self._ok())
        bf = ttk.Frame(f)
        bf.grid(row=5, column=0, columnspan=2)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.LEFT, padx=5)
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _ok(self) -> None:
        name = self.e_name.get().strip()
        if not name:
            messagebox.showwarning("Warnung", "Bitte einen Namen eingeben!", parent=self)
            return
        try:
            anz = int(self.e_anz.get())
        except ValueError:
            messagebox.showwarning("Warnung", "Ungültige Anzahl!", parent=self)
            return
        try:
            max_p = [int(x.strip()) for x in self.e_punkte.get().strip().split(",") if x.strip()]
        except ValueError:
            messagebox.showwarning("Warnung", "Ungültige Punkte-Eingabe!", parent=self)
            return
        if len(max_p) != anz:
            messagebox.showwarning("Warnung",
                                   f"Anzahl der Punkte-Einträge ({len(max_p)}) stimmt nicht mit Aufgabenanzahl ({anz}) überein!",
                                   parent=self)
            return
        if any(p <= 0 for p in max_p):
            messagebox.showwarning("Warnung", "Punkte müssen > 0 sein!", parent=self)
            return
        self.result = (name, max_p)
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class GewichtungDialog(_CenteredToplevel):
    """Dialog zum Ändern der prozentualen Gewichtung einer Klausur/UL."""

    def __init__(self, parent, title: str, name: str, current_gewichtung: float,
                 category_total: int, other_gewichtung: float):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.category_total = category_total
        self.other_gewichtung = other_gewichtung
        self.transient(parent)
        self.grab_set()
        f = ttk.Frame(self, padding=15)
        f.pack()
        ttk.Label(f, text=f"Gewichtung für '{name}':",
                  font=("TkDefaultFont", 10, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 10))
        ttk.Label(f, text="Anteil an Gesamtnote (%):").grid(row=1, column=0, sticky="w", pady=(0, 5))
        self.e_gw = ttk.Spinbox(f, from_=0, to=100, width=6, increment=5)
        self.e_gw.grid(row=1, column=1, sticky="w", pady=(0, 5), padx=(5, 0))
        self.e_gw.set(current_gewichtung)
        max_allowed = category_total - other_gewichtung
        info = (f"Kategorie-Gesamt: {category_total}%\n"
                f"Andere bereits vergeben: {other_gewichtung:.1f}%\n"
                f"Maximal möglich: {max_allowed:.1f}%")
        ttk.Label(f, text=info, foreground="gray").grid(row=2, column=0, columnspan=2, pady=(5, 5))
        bf = ttk.Frame(f)
        bf.grid(row=3, column=0, columnspan=2)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.LEFT, padx=5)
        self.e_gw.bind("<Return>", lambda e: self._ok())
        self.e_gw.focus_set()
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _ok(self) -> None:
        try:
            gewichtung = float(self.e_gw.get())
        except ValueError:
            messagebox.showwarning("Warnung", "Ungültige Gewichtung!", parent=self)
            return
        if gewichtung < 0:
            messagebox.showwarning("Warnung", "Gewichtung muss ≥ 0 sein!", parent=self)
            return
        total_with_new = self.other_gewichtung + gewichtung
        if total_with_new > self.category_total:
            messagebox.showwarning("Warnung",
                                   f"Summe der Gewichtungen ({total_with_new:.1f}%) "
                                   f"übersteigt Kategorie-Gesamt ({self.category_total}%).\n"
                                   f"Maximal erlaubt: {self.category_total - self.other_gewichtung:.1f}%",
                                   parent=self)
            return
        self.result = gewichtung
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class PunkteDialog(_CenteredToplevel):
    """Dialog zum Bearbeiten von Punkten für Klausuren oder Unterrichtsleistungen."""

    def __init__(self, parent, title: str, name: str, max_punkte_pro_aufgabe: list,
                 schuelerinnen: list, existing_ergebnisse: dict, ns_csv_str: str,
                 notenschluessel_typ: str = "IHK"):
        super().__init__(parent)
        self.title(title)
        self.geometry("850x550")
        self.minsize(700, 450)
        self.transient(parent)
        self.grab_set()
        self.max_punkte = max_punkte_pro_aufgabe
        self.schuelerinnen = schuelerinnen
        self.existing_ergebnisse = existing_ergebnisse
        self.ns_csv = ns_csv_str
        self.ns_typ = notenschluessel_typ  # "IHK" oder "BG"
        self.result = None
        self.num_cols = len(max_punkte_pro_aufgabe)
        self.ges_max = sum(max_punkte_pro_aufgabe)
        hf = ttk.Frame(self, padding=5)
        hf.pack(fill=tk.X)
        ttk.Label(hf, text=f"{name}", font=("TkDefaultFont", 11, "bold")).pack(side=tk.LEFT)
        ttk.Label(hf, text=f"  |  Max. Punkte: {self.ges_max} ({', '.join(str(p) for p in max_punkte_pro_aufgabe)})",
                  foreground="gray").pack(side=tk.LEFT)
        container = ttk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, padx=5)
        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)
        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        canvas.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))
        # Headers
        headers = ["Schülerin"] + [f"A{i+1} (/{p})" for i, p in enumerate(max_punkte_pro_aufgabe)] + ["Gesamt", "%", "Note"]
        for c, h in enumerate(headers):
            ttk.Label(self.inner, text=h, font=("TkDefaultFont", 9, "bold")).grid(
                row=0, column=c, padx=2, pady=2, sticky="w")
        self.entries = {}
        self.labels_ges = {}
        self.labels_pct = {}
        self.labels_note = {}
        for r, (sk, nn, vn) in enumerate(schuelerinnen, start=1):
            ttk.Label(self.inner, text=f"{nn}, {vn}").grid(row=r, column=0, sticky="w", padx=2, pady=1)
            row_entries = []
            existing = existing_ergebnisse.get(sk, [])
            for c, max_p in enumerate(max_punkte_pro_aufgabe):
                e = ttk.Entry(self.inner, width=6)
                e.grid(row=r, column=c + 1, padx=1, pady=1)
                if c < len(existing) and existing[c] is not None:
                    e.insert(0, str(existing[c]))
                e.bind("<KeyRelease>", lambda ev, row=r: self._update_row(row))
                row_entries.append(e)
            self.entries[sk] = row_entries
            lbl_g = ttk.Label(self.inner, text="", width=7)
            lbl_g.grid(row=r, column=self.num_cols + 1, padx=2)
            lbl_p = ttk.Label(self.inner, text="", width=7)
            lbl_p.grid(row=r, column=self.num_cols + 2, padx=2)
            lbl_n = ttk.Label(self.inner, text="", width=7, font=("TkDefaultFont", 9, "bold"))
            lbl_n.grid(row=r, column=self.num_cols + 3, padx=2)
            self.labels_ges[r] = lbl_g
            self.labels_pct[r] = lbl_p
            self.labels_note[r] = lbl_n
            self._update_row(r)
        # Abstand vor "Unterm Strich"
        ttk.Frame(self.inner, height=10).grid(row=len(schuelerinnen) + 1, column=0, pady=5)
        # "Unterm Strich" Zeile
        self.unter_strich_row = len(schuelerinnen) + 2
        ttk.Separator(self.inner, orient=tk.HORIZONTAL).grid(
            row=self.unter_strich_row, column=0, columnspan=self.num_cols + 4, sticky="ew", pady=(5, 5))
        ttk.Label(self.inner, text="⸺ Unterm Strich:", font=("TkDefaultFont", 10, "bold")).grid(
            row=self.unter_strich_row + 1, column=0, sticky="w", padx=2, pady=5)
        # Durchschnitt pro Aufgabe
        self.labels_avg_per_task = []
        for c in range(self.num_cols):
            lbl = ttk.Label(self.inner, text="", width=7, font=("TkDefaultFont", 9, "bold"), foreground="#2a5da8")
            lbl.grid(row=self.unter_strich_row + 1, column=c + 1, padx=2)
            self.labels_avg_per_task.append(lbl)
        # Gesamt, %, Note unterm Strich
        self.lbl_ul_ges = ttk.Label(self.inner, text="", width=10, font=("TkDefaultFont", 10, "bold"), foreground="#2a5da8")
        self.lbl_ul_ges.grid(row=self.unter_strich_row + 1, column=self.num_cols + 1, padx=2)
        self.lbl_ul_pct = ttk.Label(self.inner, text="", width=10, font=("TkDefaultFont", 10, "bold"), foreground="#2a5da8")
        self.lbl_ul_pct.grid(row=self.unter_strich_row + 1, column=self.num_cols + 2, padx=2)
        self.lbl_ul_note = ttk.Label(self.inner, text="", width=10, font=("TkDefaultFont", 11, "bold"), foreground="#c44")
        self.lbl_ul_note.grid(row=self.unter_strich_row + 1, column=self.num_cols + 3, padx=2)
        bf = ttk.Frame(self, padding=5)
        bf.pack(fill=tk.X)
        ttk.Button(bf, text="Alle berechnen", command=self._update_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.RIGHT, padx=5)
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _get_row_points(self, r: int) -> list:
        sk = self.schuelerinnen[r - 1][0]
        pts = []
        for e in self.entries[sk]:
            v = e.get().strip()
            if v == "":
                pts.append(None)
            else:
                try:
                    pts.append(int(v))
                except ValueError:
                    try:
                        pts.append(float(v))
                    except ValueError:
                        pts.append(None)
        return pts

    def _update_row(self, r: int) -> None:
        pts = self._get_row_points(r)
        if all(p is not None for p in pts):
            ges = sum(pts)
            max_p = sum(self.max_punkte)
            pct_raw = ges / max_p * 100 if max_p > 0 else 0
            pct = NotenVerwaltung._round_pct(pct_raw)
            note = NotenVerwaltung.ns_csv_lookup(pct, self.ns_csv)
            self.labels_ges[r].config(text=f"{ges}/{max_p}")
            self.labels_pct[r].config(text=f"{pct}%")
            note_str = (f"{note:.0f}" if note is not None and note == int(note)
                        else (f"{note:.1f}" if note is not None else "-"))
            self.labels_note[r].config(text=note_str)
            # Nicht bestanden: IHK Note > 4.4, BG Note < 3.5
            if note is not None:
                if self.ns_typ == "IHK" and note > 4.4:
                    self.labels_note[r].config(foreground="#c44")
                elif self.ns_typ == "BG" and note < 3.5:
                    self.labels_note[r].config(foreground="#c44")
                else:
                    self.labels_note[r].config(foreground="black")
            else:
                self.labels_note[r].config(foreground="black")
        else:
            self.labels_ges[r].config(text="")
            self.labels_pct[r].config(text="")
            self.labels_note[r].config(text="")

    def _update_all(self) -> None:
        for r in range(1, len(self.schuelerinnen) + 1):
            self._update_row(r)
        self._update_unter_strich()

    def _update_unter_strich(self) -> None:
        """Berechnet und aktualisiert die 'Unterm Strich' Zeile mit Durchschnitten."""
        all_points_per_task = [[] for _ in range(self.num_cols)]
        all_total_points = []
        all_pcts = []

        for r in range(1, len(self.schuelerinnen) + 1):
            pts = self._get_row_points(r)
            if all(p is not None for p in pts):
                for c, p in enumerate(pts):
                    all_points_per_task[c].append(p)
                ges = sum(pts)
                all_total_points.append(ges)
                max_p = sum(self.max_punkte)
                pct_raw = ges / max_p * 100 if max_p > 0 else 0
                pct = NotenVerwaltung._round_pct(pct_raw)
                all_pcts.append(pct)

        # Durchschnitt pro Aufgabe in Prozent
        for c in range(self.num_cols):
            if all_points_per_task[c]:
                avg = sum(all_points_per_task[c]) / len(all_points_per_task[c])
                pct_avg = NotenVerwaltung._round_pct(avg / self.max_punkte[c] * 100) if self.max_punkte[c] > 0 else 0
                self.labels_avg_per_task[c].config(text=f"Ø{pct_avg}%")
            else:
                self.labels_avg_per_task[c].config(text="")

        # Gesamt, %, Note unterm Strich
        if all_total_points and all_pcts:
            ges_avg = sum(all_total_points) / len(all_total_points)
            pct_avg = sum(all_pcts) / len(all_pcts)
            pct_rounded = NotenVerwaltung._round_pct(pct_avg)
            note = NotenVerwaltung.ns_csv_lookup(pct_rounded, self.ns_csv)

            self.lbl_ul_ges.config(text=f"Ø{ges_avg:.1f}/{self.ges_max}")
            self.lbl_ul_pct.config(text=f"Ø{pct_rounded}%")
            note_str = (f"{note:.0f}" if note is not None and float(note).is_integer()
                        else (f"{note:.1f}" if note is not None else "-"))
            self.lbl_ul_note.config(text=f"Ø{note_str}")
        else:
            self.lbl_ul_ges.config(text="")
            self.lbl_ul_pct.config(text="")
            self.lbl_ul_note.config(text="")

    def _ok(self) -> None:
        self.result = {}
        for sk, nn, vn in self.schuelerinnen:
            pts = []
            all_filled = True
            for i, e in enumerate(self.entries[sk]):
                v = e.get().strip()
                if v == "":
                    pts.append(None)
                    all_filled = False
                else:
                    try:
                        p = float(v)
                        if p < 0 or p > self.max_punkte[i]:
                            messagebox.showwarning("Warnung",
                                                   f"Punkte für {nn}, {vn} Aufgabe {i+1} müssen "
                                                   f"zwischen 0 und {self.max_punkte[i]} liegen!",
                                                   parent=self)
                            return
                        pts.append(p)
                    except ValueError:
                        messagebox.showwarning("Warnung",
                                               f"Ungültige Eingabe für {nn}, {vn} Aufgabe {i+1}!",
                                               parent=self)
                        return
            if all_filled:
                self.result[sk] = pts
            elif any(p is not None for p in pts):
                existing = self.existing_ergebnisse.get(sk, [])
                merged = []
                for i in range(len(self.max_punkte)):
                    if i < len(pts) and pts[i] is not None:
                        merged.append(pts[i])
                    elif i < len(existing) and existing[i] is not None:
                        merged.append(existing[i])
                    else:
                        merged.append(None)
                self.result[sk] = merged
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class _UebertragenDialog(_CenteredToplevel):
    def __init__(self, parent, sj_quelle: str, kl_name: str, alle_sj: dict):
        super().__init__(parent)
        self.title("Klasse übertragen")
        self.resizable(False, False)
        self.result = None
        self.transient(parent)
        self.grab_set()
        f = ttk.Frame(self, padding=15)
        f.pack()
        ttk.Label(f, text=f"Klasse '{kl_name}' aus Schuljahr '{sj_quelle}'\nübertragen in Schuljahr:",
                  justify="left").grid(row=0, column=0, columnspan=2, pady=(0, 10))
        andere_sj = sorted(s for s in alle_sj if s != sj_quelle)
        ttk.Label(f, text="Ziel-Schuljahr:").grid(row=1, column=0, sticky="w", pady=(0, 5))
        self.sj_var = tk.StringVar()
        self.sj_cb = ttk.Combobox(f, textvariable=self.sj_var, width=18, values=andere_sj)
        self.sj_cb.grid(row=1, column=1, pady=(0, 5), padx=(5, 0))
        if andere_sj:
            self.sj_var.set(andere_sj[0])
        self.sj_cb.bind("<Return>", lambda e: self._ok())
        bf = ttk.Frame(f)
        bf.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Neues SJ anlegen", command=self._new_sj, width=14).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.LEFT, padx=5)
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _ok(self) -> None:
        ziel = self.sj_var.get().strip()
        if ziel:
            self.result = ziel
        self.destroy()

    def _new_sj(self) -> None:
        ns = simpledialog.askstring("Neues Schuljahr", "Schuljahr (z.B. 2026/27):", parent=self)
        if ns and ns.strip():
            self.sj_var.set(ns.strip())
            cur = list(self.sj_cb['values']) + [ns.strip()]
            self.sj_cb['values'] = sorted(set(cur))

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class NotenschluesselCsvDialog(_CenteredToplevel):
    def __init__(self, parent, current_csv: str, notenschluessel_typ: str, alle_klassen: list):
        super().__init__(parent)
        self.title("Notenschlüssel bearbeiten")
        self.resizable(True, True)
        self.geometry("550x500")
        self.transient(parent)
        self.grab_set()
        self.alle_klassen = alle_klassen
        self.result = None
        f = ttk.Frame(self, padding=10)
        f.pack(fill=tk.BOTH, expand=True)
        ttk.Label(f, text="Notenschlüssel im CSV-Format:",
                  font=("TkDefaultFont", 10, "bold")).pack(anchor="w")
        ttk.Label(f, text="Format: prozent,note;prozent,note;...  (absteigend nach %)",
                  foreground="gray").pack(anchor="w", pady=(0, 5))
        std_frame = ttk.Frame(f)
        std_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(std_frame, text="Standard laden:").pack(side=tk.LEFT, padx=(0, 5))
        for ns in NOTENSCHLUESSEL:
            ttk.Button(std_frame, text=f"{ns}",
                       command=lambda n=ns: self._load_default(n), width=8).pack(side=tk.LEFT, padx=2)
        self.text = tk.Text(f, height=5, width=60, font=("Courier", 10))
        self.text.pack(fill=tk.X, pady=(0, 5))
        self.text.insert("1.0", current_csv)
        ttk.Label(f, text="Vorschau:", font=("TkDefaultFont", 10, "bold")).pack(anchor="w", pady=(5, 2))
        self.preview = tk.Text(f, height=8, width=60, font=("Courier", 9),
                               state="disabled", background="#f0f0f0")
        self.preview.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        ttk.Button(f, text="Vorschau aktualisieren", command=self._update_preview).pack(anchor="w", pady=(0, 10))
        ttk.Label(f, text="Auf andere Klassen übertragen:",
                  font=("TkDefaultFont", 10, "bold")).pack(anchor="w", pady=(5, 2))
        self.transfer_vars = {}
        if alle_klassen:
            tf = ttk.Frame(f)
            tf.pack(fill=tk.X, pady=(0, 5))
            for sj, k in alle_klassen:
                var = tk.BooleanVar(value=False)
                self.transfer_vars[(sj, k)] = var
                ttk.Checkbutton(tf, text=f"{k} ({sj})", variable=var).pack(anchor="w")
        else:
            ttk.Label(f, text="(Keine anderen Klassen vorhanden)", foreground="gray").pack(anchor="w")
        bf = ttk.Frame(f)
        bf.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.RIGHT, padx=5)
        self._update_preview()
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _load_default(self, ns: str) -> None:
        self.text.delete("1.0", tk.END)
        self.text.insert("1.0", DEFAULT_NS_CSV.get(ns, ""))
        self._update_preview()

    def _update_preview(self) -> None:
        csv_str = self.text.get("1.0", tk.END).strip()
        entries = NotenVerwaltung.ns_csv_parse(csv_str)
        self.preview.config(state="normal")
        self.preview.delete("1.0", tk.END)
        if not entries:
            self.preview.insert(tk.END, "Ungültiges Format oder leer.")
        else:
            self.preview.insert(tk.END, f"{'Prozent ≥':>12} | {'Note':>6}\n")
            self.preview.insert(tk.END, "-" * 25 + "\n")
            for p, n in entries:
                self.preview.insert(tk.END, f"{p:>10.1f}% | {n:>6}\n")
        self.preview.config(state="disabled")

    def _ok(self) -> None:
        csv_str = self.text.get("1.0", tk.END).strip()
        entries = NotenVerwaltung.ns_csv_parse(csv_str)
        if not entries:
            messagebox.showwarning("Warnung", "Ungültiges Format!", parent=self)
            return
        transfer = [(sj, k) for (sj, k), var in self.transfer_vars.items() if var.get()]
        self.result = (csv_str, transfer)
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


# ---------------------------------------------------------------------------
# Hauptanwendung
# ---------------------------------------------------------------------------
class NotenVerwaltungApp:
    def __init__(self, root: tk.Tk, password: str, data_file: Optional[str] = None):
        self.root = root
        self.root.title("Notenverwaltung")
        self.root.geometry("1100x680")
        self.root.minsize(950, 580)
        self.password = password
        self.data_file = data_file or DATA_FILE
        self.daten = NotenVerwaltung()
        self._init_failed = False
        if os.path.exists(self.data_file):
            if not self.daten.laden_verschluesselt(self.password, self.data_file):
                messagebox.showerror("Fehler", "Falsches Passwort oder beschädigte Daten!")
                self._init_failed = True
                self.root.after(100, self.root.destroy)
                return
        else:
            self.daten.speichern_verschluesselt(self.password, self.data_file)
        self._build_gui()
        self._refresh_sj()
        self._update_title()
        self._start_auto_save()

    def _save(self) -> None:
        self.daten.speichern_verschluesselt(self.password, self.data_file)

    def _update_title(self) -> None:
        fname = os.path.basename(self.data_file) if self.data_file else "noten.ndat"
        self.root.title(f"Notenverwaltung – {fname}")

    def _start_auto_save(self) -> None:
        """Periodisches Auto-Save alle 60 Sekunden."""
        self._auto_save()

    def _auto_save(self) -> None:
        try:
            self._save()
        except Exception as e:
            logger.error("Auto-Save fehlgeschlagen: %s", e)
        self.root.after(AUTO_SAVE_MS, self._auto_save)

    def _build_gui(self) -> None:
        gm = self.daten.gewichtung_muendlich
        gs = 100 - gm
        sty = ttk.Style()
        sty.configure("TButton", font=("TkDefaultFont", 12), padding=(10, 6))
        sty.configure("TLabel", font=("TkDefaultFont", 12))
        sty.configure("TNotebook.Tab", font=("TkDefaultFont", 12, "bold"), padding=(18, 8))
        sty.configure("H.TLabel", font=("TkDefaultFont", 13, "bold"))
        sty.configure("G.TLabel", font=("TkDefaultFont", 14, "bold"))
        sty.configure("J.TLabel", font=("TkDefaultFont", 13, "bold"), foreground="#2a5da8")
        sty.configure("I.TLabel", font=("TkDefaultFont", 10), foreground="gray")
        sty.configure("NS.TLabel", font=("TkDefaultFont", 11, "bold"), foreground="#c44")
        sty.configure("TLabelframe.Label", font=("TkDefaultFont", 12, "bold"))
        sty.configure("TSpinbox", font=("TkDefaultFont", 12), padding=(4, 4))
        sty.configure("Treeview", font=("TkDefaultFont", 12), rowheight=30)
        sty.configure("Treeview.Heading", font=("TkDefaultFont", 11, "bold"))
        # Menubar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        fm = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Datei", menu=fm)
        fm.add_command(label="Öffnen...", command=self._file_open, accelerator="Strg+O")
        fm.add_command(label="Speichern unter...", command=self._file_save_as, accelerator="Strg+Shift+S")
        fm.add_separator()
        fm.add_command(label="Passwort ändern...", command=self._change_password)
        fm.add_separator()
        fm.add_command(label="Export Markdown...", command=self._export_md)
        fm.add_command(label="Export CSV (Excel)...", command=self._export_csv)
        fm.add_separator()
        fm.add_command(label="Beenden", command=self.root.quit)
        # Hauptframe
        hf = ttk.Frame(self.root, padding=10)
        hf.pack(fill=tk.BOTH, expand=True)
        # Einstellungen
        top = ttk.LabelFrame(hf, text="Einstellungen", padding=5)
        top.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        # Zeile 1: Schuljahr, Halbjahr, Gewichtung
        row1 = ttk.Frame(top)
        row1.pack(fill=tk.X, pady=(0, 3))
        ttk.Label(row1, text="Schuljahr:").pack(side=tk.LEFT, padx=(0, 3))
        self.sj_var = tk.StringVar()
        self.sj_cb = ttk.Combobox(row1, textvariable=self.sj_var, state="readonly", width=12)
        self.sj_cb.pack(side=tk.LEFT, padx=(0, 5))
        self.sj_cb.bind("<<ComboboxSelected>>", self._on_sj)
        ttk.Button(row1, text="+", width=3, command=self._sj_add).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(row1, text="−", width=3, command=self._sj_del).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Separator(row1, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Label(row1, text="Halbjahr:").pack(side=tk.LEFT, padx=(5, 3))
        self.hj_var = tk.StringVar()
        self.hj_cb = ttk.Combobox(row1, textvariable=self.hj_var, state="readonly", width=12)
        self.hj_cb.pack(side=tk.LEFT, padx=(0, 10))
        self.hj_cb['values'] = HALBJAHRE
        self.hj_var.set(HALBJAHRE[0])
        self.hj_cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_all())
        # Zeile 2: Klasse, Fach
        row2 = ttk.Frame(top)
        row2.pack(fill=tk.X)
        ttk.Label(row2, text="Klasse:").pack(side=tk.LEFT, padx=(0, 3))
        self.kl_var = tk.StringVar()
        self.kl_cb = ttk.Combobox(row2, textvariable=self.kl_var, state="readonly", width=12)
        self.kl_cb.pack(side=tk.LEFT, padx=(0, 2))
        self.kl_cb.bind("<<ComboboxSelected>>", self._on_kl)
        kl_menu = ttk.Frame(row2)
        kl_menu.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(kl_menu, text="+", width=3, command=self._kl_add).pack(side=tk.LEFT, padx=(0, 1))
        ttk.Button(kl_menu, text="→", width=3, command=self._kl_uebertragen).pack(side=tk.LEFT, padx=(0, 1))
        ttk.Button(kl_menu, text="−", width=3, command=self._kl_del).pack(side=tk.LEFT)
        ttk.Separator(row2, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Label(row2, text="Fach:").pack(side=tk.LEFT, padx=(5, 3))
        self.fach_var = tk.StringVar()
        self.fach_cb = ttk.Combobox(row2, textvariable=self.fach_var, state="readonly", width=14)
        self.fach_cb.pack(side=tk.LEFT, padx=(0, 2))
        self.fach_cb.bind("<<ComboboxSelected>>", self._on_fach)
        ttk.Button(row2, text="+", width=3, command=self._fach_add).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(row2, text="−", width=3, command=self._fach_del).pack(side=tk.LEFT)
        # Schülerinnen
        sf = ttk.LabelFrame(hf, text="Schülerinnen", padding=5)
        sf.grid(row=1, column=0, sticky="nsew", padx=(0, 3))
        self.sk_lb = tk.Listbox(sf, height=18, width=22, exportselection=False,
                                 font=("TkDefaultFont", 12), selectbackground="#4a90d9")
        self.sk_lb.pack(fill=tk.BOTH, expand=True)
        self.sk_lb.bind("<<ListboxSelect>>", self._on_sk)
        bf2 = ttk.Frame(sf)
        bf2.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(bf2, text="Neu", command=self._sk_add).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        ttk.Button(bf2, text="Liste", command=self._sk_list_add).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(bf2, text="Löschen", command=self._sk_del).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))
        # Notebook
        self.nb = ttk.Notebook(hf)
        self.nb.grid(row=1, column=1, sticky="nsew", padx=(3, 0))
        self._build_noten_tab(gm, gs)
        self._build_ul_tab()
        self._build_klausuren_tab()
        hf.columnconfigure(0, weight=1)
        hf.columnconfigure(1, weight=5)
        hf.rowconfigure(1, weight=1)
        self.root.bind("<Control-o>", lambda e: self._file_open())
        self.root.bind("<Control-O>", lambda e: self._file_open())
        self.root.bind("<Control-Shift-s>", lambda e: self._file_save_as())
        self.root.bind("<Control-Shift-S>", lambda e: self._file_save_as())

    def _build_noten_tab(self, gm: int, gs: int) -> None:
        nf = ttk.Frame(self.nb, padding=5)
        self.nb.add(nf, text="  Noten  ")
        self.info_lbl = ttk.Label(nf, text="Bitte eine Schülerin auswählen", style="H.TLabel")
        self.info_lbl.pack(anchor="w", pady=(0, 2))
        self.ns_lbl = ttk.Label(nf, text="", style="NS.TLabel")
        self.ns_lbl.pack(anchor="w", pady=(0, 5))
        self.m_frame = ttk.LabelFrame(nf, text="Unterrichtsleistung", padding=5)
        self.m_frame.pack(fill=tk.BOTH, expand=True)
        gw_bar = ttk.Frame(self.m_frame)
        gw_bar.pack(fill=tk.X, pady=(0, 3))
        ttk.Label(gw_bar, text="Gewichtung:").pack(side=tk.LEFT, padx=(0, 3))
        self.gw_var = tk.StringVar(value=str(gm))
        self.gw_sb = ttk.Spinbox(gw_bar, from_=0, to=100, width=4, textvariable=self.gw_var,
                                  command=self._on_gw)
        self.gw_sb.pack(side=tk.LEFT, padx=(0, 2))
        self.gw_ul_lbl = ttk.Label(gw_bar, text=f"{gm}%")
        self.gw_ul_lbl.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(gw_bar, text="Schriftlich:").pack(side=tk.LEFT, padx=(0, 3))
        self.gw_sl = ttk.Label(gw_bar, text=f"{gs}%")
        self.gw_sl.pack(side=tk.LEFT)
        self.gw_sb.bind("<Return>", lambda e: self._on_gw())
        self.m_lb = tk.Listbox(self.m_frame, height=5, exportselection=False,
                                font=("TkDefaultFont", 12), selectbackground="#4a90d9")
        self.m_lb.pack(fill=tk.BOTH, expand=True)
        mbf = ttk.Frame(self.m_frame)
        mbf.pack(fill=tk.X, pady=(5, 0))
        self.m_sp = ttk.Spinbox(mbf, from_=1, to=6, width=5)
        self.m_sp.pack(side=tk.LEFT, padx=(0, 5))
        self.m_sp.set(1)
        ttk.Button(mbf, text="Note eintragen",
                   command=lambda: self._note_add("muendlich")).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(mbf, text="Löschen",
                   command=lambda: self._note_del("muendlich")).pack(side=tk.LEFT)
        self.m_avg = ttk.Label(self.m_frame, text="")
        self.m_avg.pack(anchor="w", pady=(0, 2))
        self.m_count_lbl = ttk.Label(self.m_frame, text="", style="I.TLabel")
        self.m_count_lbl.pack(anchor="w", pady=(0, 0))
        self.s_frame = ttk.LabelFrame(nf, text=f"Schriftliche Noten ({gs}%)", padding=5)
        self.s_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        self.s_lb = tk.Listbox(self.s_frame, height=5, exportselection=False,
                                font=("TkDefaultFont", 12), selectbackground="#4a90d9")
        self.s_lb.pack(fill=tk.BOTH, expand=True)
        sbf = ttk.Frame(self.s_frame)
        sbf.pack(fill=tk.X, pady=(5, 0))
        self.s_sp = ttk.Spinbox(sbf, from_=1, to=6, width=5)
        self.s_sp.pack(side=tk.LEFT, padx=(0, 5))
        self.s_sp.set(1)
        ttk.Button(sbf, text="Note eintragen",
                   command=lambda: self._note_add("schriftlich")).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(sbf, text="Löschen",
                   command=lambda: self._note_del("schriftlich")).pack(side=tk.LEFT)
        self.s_avg = ttk.Label(self.s_frame, text="")
        self.s_avg.pack(anchor="w", pady=(0, 2))
        self.s_count_lbl = ttk.Label(self.s_frame, text="", style="I.TLabel")
        self.s_count_lbl.pack(anchor="w", pady=(0, 0))
        self.g_lbl = ttk.Label(nf, text="Gesamtnote: -", style="G.TLabel")
        self.g_lbl.pack(anchor="w", pady=(10, 0))
        self.j_lbl = ttk.Label(nf, text="Jahresnote: -", style="J.TLabel")
        self.j_lbl.pack(anchor="w", pady=(8, 0))
        ttk.Label(nf, text="(Gesamtnote über beide Halbjahre)", style="I.TLabel").pack(anchor="w")

    def _build_bewertung_tab(self, parent_frame: ttk.Frame, title: str) -> tuple:
        """Generischer Aufbau für Klausuren/UL-Tabs. Gibt (listbox, btn_frame, tree) zurück."""
        ttk.Label(parent_frame, text=f"{title}:", style="H.TLabel").pack(anchor="w", pady=(0, 5))
        lb = tk.Listbox(parent_frame, height=5, exportselection=False,
                         font=("TkDefaultFont", 12), selectbackground="#4a90d9")
        lb.pack(fill=tk.X, pady=(0, 5))
        btn_frame = ttk.Frame(parent_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 5))
        tree_frame = ttk.Frame(parent_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(tree_frame, columns=("info",), show="headings", height=8)
        tree.heading("info", text=f"Keine {title} ausgewählt")
        tree.column("info", width=400)
        tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=tree_scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        return lb, btn_frame, tree

    def _build_klausuren_tab(self) -> None:
        kf = ttk.Frame(self.nb, padding=5)
        self.nb.add(kf, text="  Klausuren  ")
        self.kl_klausur_lb, btn_frame, self.kl_tree = self._build_bewertung_tab(kf, "Klausuren")
        self.kl_klausur_lb.bind("<<ListboxSelect>>", self._on_klausur_select)
        ttk.Button(btn_frame, text="Hinzufügen", command=self._klausur_add).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Löschen", command=self._klausur_del).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Punkte bearbeiten", command=self._klausur_punkte).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Gewichtung", command=self._klausur_gewichtung).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Notenschlüssel", command=self._ns_csv_edit).pack(side=tk.LEFT, padx=(0, 5))

    def _build_ul_tab(self) -> None:
        uf = ttk.Frame(self.nb, padding=5)
        self.nb.add(uf, text="  Unterrichtsleistungen  ")
        self.ul_lb, btn_frame, self.ul_tree = self._build_bewertung_tab(uf, "Unterrichtsleistungen")
        self.ul_lb.bind("<<ListboxSelect>>", self._on_ul_select)
        ttk.Button(btn_frame, text="Hinzufügen", command=self._ul_add).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Löschen", command=self._ul_del).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Punkte bearbeiten", command=self._ul_punkte).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Gewichtung", command=self._ul_gewichtung).pack(side=tk.LEFT, padx=(0, 5))

    # ---- Password / Export ----
    def _file_open(self) -> None:
        fp = filedialog.askopenfilename(
            defaultextension=".ndat",
            filetypes=[("Notendatei", "*.ndat"), ("Alle Dateien", "*.*")],
            title="Notendatei öffnen",
            initialdir=os.path.dirname(self.data_file) if self.data_file else _APP_DIR)
        if not fp:
            return
        dlg = PasswordDialog(self.root, title="Passwort eingeben", first_time=False)
        if dlg.result is None:
            return
        new_data = NotenVerwaltung()
        if not new_data.laden_verschluesselt(dlg.result, fp):
            messagebox.showerror("Fehler", "Falsches Passwort oder beschädigte Daten!")
            return
        self.password = dlg.result
        self.data_file = fp
        self.daten = new_data
        self._update_title()
        self._refresh_sj()
        self._refresh_noten(None, None)
        self._refresh_klausuren()
        self._refresh_ul()

    def _file_save_as(self) -> None:
        fp = filedialog.asksaveasfilename(
            defaultextension=".ndat",
            filetypes=[("Notendatei", "*.ndat"), ("Alle Dateien", "*.*")],
            title="Notendatei speichern unter",
            initialdir=os.path.dirname(self.data_file) if self.data_file else _APP_DIR,
            initialfile=os.path.basename(self.data_file) if self.data_file else "noten.ndat")
        if not fp:
            return
        self.data_file = fp
        self._save()
        self._update_title()

    def _change_password(self) -> None:
        dlg = PasswordDialog(self.root, title="Passwort ändern", first_time=False)
        if dlg.result is None:
            return
        self.password = dlg.result
        self._save()
        messagebox.showinfo("OK", "Passwort geändert und Daten gespeichert.")

    def _export_md(self) -> None:
        fp = filedialog.asksaveasfilename(defaultextension=".md",
                                           filetypes=[("Markdown", "*.md")], title="Export Markdown")
        if fp:
            try:
                self.daten.export_markdown(fp)
                messagebox.showinfo("Export", f"Markdown exportiert nach:\n{fp}")
            except Exception as e:
                messagebox.showerror("Fehler", f"Export fehlgeschlagen: {e}")

    def _export_csv(self) -> None:
        fp = filedialog.asksaveasfilename(defaultextension=".csv",
                                           filetypes=[("CSV", "*.csv")], title="Export CSV")
        if fp:
            try:
                self.daten.export_csv(fp)
                messagebox.showinfo("Export", f"CSV exportiert nach:\n{fp}")
            except Exception as e:
                messagebox.showerror("Fehler", f"Export fehlgeschlagen: {e}")

    # ---- Getters ----
    def _sj(self) -> Optional[str]:
        return self.sj_var.get() or None

    def _hj(self) -> str:
        return self.hj_var.get() or HALBJAHRE[0]

    def _kl(self) -> Optional[str]:
        v = self.kl_var.get()
        return v if v else None

    def _fach(self) -> Optional[str]:
        v = self.fach_var.get()
        return v if v else None

    def _sk(self) -> Optional[str]:
        s = self.sk_lb.curselection()
        return self.sk_lb.get(s[0]) if s else None

    @staticmethod
    def _parse_kl_name(kl_display: Optional[str]) -> Optional[str]:
        if kl_display is None:
            return None
        if " [" in kl_display:
            return kl_display.rsplit(" [", 1)[0]
        return kl_display

    # ---- Refresh ----
    def _refresh_all(self) -> None:
        self._refresh_noten(self._kl(), self._sk())
        self._refresh_klausuren()
        self._refresh_ul()

    def _refresh_sj(self) -> None:
        sl = sorted(self.daten.schuljahre.keys())
        self.sj_cb['values'] = sl
        if sl:
            if self._sj() not in sl:
                self.sj_var.set(sl[0])
            self._refresh_kl()
        else:
            self.sj_var.set("")
            self.kl_var.set("")
            self.kl_cb['values'] = []
            self._refresh_sk(None)
            self._refresh_noten(None, None)

    def _refresh_kl(self) -> None:
        sj = self._sj()
        current = self.kl_var.get()
        if sj and sj in self.daten.schuljahre:
            values = [f"{k} [{self.daten.get_notenschluessel(sj, k)}]"
                      for k in sorted(self.daten.schuljahre[sj])]
            self.kl_cb['values'] = values
            if current not in values:
                self.kl_var.set(values[0] if values else "")
        else:
            self.kl_cb['values'] = []
            self.kl_var.set("")

    def _refresh_fach(self, kl: Optional[str]) -> None:
        sj = self._sj()
        kl_name = self._parse_kl_name(kl) if kl else None
        current = self.fach_var.get()
        if sj and kl_name and kl_name in self.daten.schuljahre.get(sj, {}):
            faecher = self.daten.fach_sortiert(sj, kl_name)
            self.fach_cb['values'] = faecher
            if current not in faecher:
                self.fach_var.set(faecher[0] if faecher else "")
                self._on_fach(None)
        else:
            self.fach_cb['values'] = []
            self.fach_var.set("")

    def _refresh_sk(self, kl: Optional[str]) -> None:
        self.sk_lb.delete(0, tk.END)
        sj = self._sj()
        kl_name = self._parse_kl_name(kl) if kl else None
        if sj and kl_name and kl_name in self.daten.schuljahre.get(sj, {}):
            for sk in self.daten.schuelerin_sortiert(sj, kl_name):
                self.sk_lb.insert(tk.END, sk)

    def _refresh_noten(self, kl: Optional[str], sk: Optional[str]) -> None:
        self.m_lb.delete(0, tk.END)
        self.s_lb.delete(0, tk.END)
        fach = self._fach()
        if not kl or not sk or not fach:
            self.info_lbl.config(text="Bitte Klasse, Fach und Schülerin auswählen")
            self.ns_lbl.config(text="")
            self.m_avg.config(text="")
            self.m_count_lbl.config(text="")
            self.s_avg.config(text="")
            self.s_count_lbl.config(text="")
            self.g_lbl.config(text="Gesamtnote: -")
            self.j_lbl.config(text="Jahresnote: -")
            return
        sj = self._sj()
        hj = self._hj()
        kl_name = self._parse_kl_name(kl)
        sd = self.daten.get_schueler_dict(sj, kl_name)
        if sd is None or sk not in sd:
            return
        d = sd[sk]
        ns = self.daten.get_notenschluessel(sj, kl_name)
        nb = self.daten.get_notenbereich(sj, kl_name)
        self.info_lbl.config(text=f"{d['nachname']}, {d['vorname']} – {fach} ({kl_name}) — {hj}")
        self.ns_lbl.config(text=f"Notenschlüssel: {ns} (Noten {nb[0]}–{nb[1]})")
        # UL
        remaining_ul = self.daten.get_remaining_ul_pct(sj, kl_name, fach, hj)
        muendlich = self.daten.get_muendlich(sj, kl_name, fach, sk, hj)
        for i, n in enumerate(muendlich):
            self.m_lb.insert(tk.END, f"{i+1}. Note: {n}")
        ul_notes_gw = self.daten.get_ul_noten_gewichtet(sj, kl_name, fach, sk, hj)
        uls = self.daten.get_unterrichtsleistungen(sj, kl_name, fach, hj)
        for i, (n, gw) in enumerate(ul_notes_gw):
            uname = uls[i]["name"] if i < len(uls) else f"UL{i+1}"
            n_str = f"{n:.0f}" if float(n).is_integer() else f"{n:.1f}"
            self.m_lb.insert(tk.END, f"  UL: {uname} ({gw}%) → {n_str}")
        ul_info_parts = []
        if muendlich:
            avg_m = sum(muendlich) / len(muendlich)
            ul_info_parts.append(f"Manuell Ø {avg_m:.1f} ({remaining_ul:.0f}%)")
        if ul_notes_gw:
            for i, (n, gw) in enumerate(ul_notes_gw):
                uname = uls[i]["name"] if i < len(uls) else f"UL{i+1}"
                ul_info_parts.append(f"{uname} ({gw}%)")
        self.m_avg.config(text=" | ".join(ul_info_parts) if ul_info_parts else "")
        # Anzahl UL-Noten anzeigen
        ul_count = len(muendlich)
        ul_bewertet_count = len(ul_notes_gw)
        if ul_count > 0 or ul_bewertet_count > 0:
            parts = []
            if ul_count > 0:
                parts.append(f"{ul_count} manuelle Note(n)")
            if ul_bewertet_count > 0:
                parts.append(f"{ul_bewertet_count} bewertete UL(s)")
            self.m_count_lbl.config(text=f"Anzahl: {' | '.join(parts)}")
        else:
            self.m_count_lbl.config(text="Anzahl: 0")
        # Schriftlich
        remaining_schr = self.daten.get_remaining_schriftlich_pct(sj, kl_name, fach, hj)
        schriftlich = self.daten.get_schriftlich(sj, kl_name, fach, sk, hj)
        for i, n in enumerate(schriftlich):
            self.s_lb.insert(tk.END, f"{i+1}. Note: {n}")
        kn_notes_gw = self.daten.get_klausur_noten_gewichtet(sj, kl_name, fach, sk, hj)
        klausuren = self.daten.get_klausuren(sj, kl_name, fach, hj)
        for i, (n, gw) in enumerate(kn_notes_gw):
            kname = klausuren[i]["name"] if i < len(klausuren) else f"K{i+1}"
            n_str = f"{n:.0f}" if float(n).is_integer() else f"{n:.1f}"
            self.s_lb.insert(tk.END, f"  K: {kname} ({gw}%) → {n_str}")
        schr_info_parts = []
        if schriftlich:
            avg_s = sum(schriftlich) / len(schriftlich)
            schr_info_parts.append(f"Manuell Ø {avg_s:.1f} ({remaining_schr:.0f}%)")
        if kn_notes_gw:
            for i, (n, gw) in enumerate(kn_notes_gw):
                kname = klausuren[i]["name"] if i < len(klausuren) else f"K{i+1}"
                schr_info_parts.append(f"{kname} ({gw}%)")
        self.s_avg.config(text=" | ".join(schr_info_parts) if schr_info_parts else "")
        # Anzahl schriftliche Noten anzeigen
        schr_count = len(schriftlich)
        klausur_count = len(kn_notes_gw)
        if schr_count > 0 or klausur_count > 0:
            parts = []
            if schr_count > 0:
                parts.append(f"{schr_count} manuelle Note(n)")
            if klausur_count > 0:
                parts.append(f"{klausur_count} Klausur(n)")
            self.s_count_lbl.config(text=f"Anzahl: {' | '.join(parts)}")
        else:
            self.s_count_lbl.config(text="Anzahl: 0")
        gn = self.daten.gesamtnote_hj(sj, kl_name, fach, sk, hj)
        self.g_lbl.config(text=f"Gesamtnote ({hj}): {gn:.2f}" if gn is not None else f"Gesamtnote ({hj}): -")
        jn = self.daten.gesamtnote_jahr(sj, kl_name, fach, sk)
        self.j_lbl.config(text=f"Jahresnote: {jn:.2f}" if jn is not None else "Jahresnote: -")
        self.m_sp.config(from_=nb[0], to=nb[1])
        self.m_sp.set(nb[0])
        self.s_sp.config(from_=nb[0], to=nb[1])
        self.s_sp.set(nb[0])

    def _refresh_bewertung(self, typ: str, keep_selection: Optional[int] = None) -> None:
        """Generischer Refresh für Klausuren- oder UL-Tab."""
        is_klausur = typ == "klausur"
        refreshing_attr = '_refreshing_klausuren' if is_klausur else '_refreshing_ul'
        lb = self.kl_klausur_lb if is_klausur else self.ul_lb
        tree = self.kl_tree if is_klausur else self.ul_tree
        get_list = self.daten.get_klausuren if is_klausur else self.daten.get_unterrichtsleistungen
        note_berechnen = self.daten.klausur_note_berechnen if is_klausur else self.daten.ul_note_berechnen
        durchschnitt_berechnen = self.daten.klausur_durchschnitt_berechnen if is_klausur else None
        nicht_bestanden_count = self.daten.klausur_nicht_bestanden_count if is_klausur else None
        label_name = "Klausuren" if is_klausur else "Unterrichtsleistungen"

        setattr(self, refreshing_attr, True)
        try:
            if keep_selection is None:
                sel = lb.curselection()
                keep_selection = sel[0] if sel else None
            lb.delete(0, tk.END)
            sj = self._sj()
            hj = self._hj()
            kl = self._kl()
            fach = self._fach()
            kl_name = self._parse_kl_name(kl) if kl else None
            if not sj or not kl_name or not fach:
                return
            items = get_list(sj, kl_name, fach, hj)
            for i, item in enumerate(items):
                max_p = sum(item["max_punkte_pro_aufgabe"])
                anzahl = len(item["max_punkte_pro_aufgabe"])
                gw = item.get("gewichtung", 0)
                # Durchschnittsnote für Klausuren berechnen und anzeigen
                durchschnitt_str = ""
                nicht_bestanden_str = ""
                if is_klausur:
                    if durchschnitt_berechnen:
                        durchschnitt = durchschnitt_berechnen(sj, kl_name, fach, hj, i)
                        if durchschnitt is not None:
                            durchschnitt_str = f" [Ø {durchschnitt:.2f}]"
                    # Nicht bestandene Noten zählen und Warnung anzeigen
                    if nicht_bestanden_count:
                        nb_count, nb_gesamt, nb_prozent, nb_warnung = nicht_bestanden_count(sj, kl_name, fach, hj, i)
                        if nb_gesamt > 0:
                            nicht_bestanden_str = f" [{nb_count}/{nb_gesamt} nicht best. ({nb_prozent}%)"
                            if nb_warnung:
                                nicht_bestanden_str += " ⚠️ GENEHMIGUNG!]"
                            else:
                                nicht_bestanden_str += "]"
                lb.insert(tk.END, f"{item['name']} ({gw}%, max {max_p} P., {anzahl} Aufg.){durchschnitt_str}{nicht_bestanden_str}")
            if keep_selection is not None and keep_selection < len(items):
                lb.selection_set(keep_selection)
                lb.see(keep_selection)
            for t_item in tree.get_children():
                tree.delete(t_item)
            if items:
                sel = lb.curselection()
                if sel:
                    idx = sel[0]
                    item = items[idx]
                    max_p = item["max_punkte_pro_aufgabe"]
                    cols = ["schuelerin"] + [f"a{i}" for i in range(len(max_p))] + ["gesamt", "prozent", "note"]
                    tree["columns"] = cols
                    tree.heading("schuelerin", text="Schülerin")
                    tree.column("schuelerin", width=150)
                    for i, mp in enumerate(max_p):
                        tree.heading(f"a{i}", text=f"A{i+1} (/{mp})")
                        tree.column(f"a{i}", width=60, anchor="center")
                    tree.heading("gesamt", text="Gesamt")
                    tree.column("gesamt", width=60, anchor="center")
                    tree.heading("prozent", text="%")
                    tree.column("prozent", width=55, anchor="center")
                    tree.heading("note", text="Note")
                    tree.column("note", width=55, anchor="center")
                    csv_str = self.daten.get_ns_csv(sj, kl_name)
                    ges_max = sum(max_p)
                    for sk in self.daten.schuelerin_sortiert(sj, kl_name):
                        d = self.daten.get_schueler_dict(sj, kl_name)[sk]
                        vals = [f"{d['nachname']}, {d['vorname']}"]
                        ergebnis = item["ergebnisse"].get(sk, [])
                        for i in range(len(max_p)):
                            vals.append(str(ergebnis[i]) if i < len(ergebnis) and ergebnis[i] is not None else "-")
                        if ergebnis and all(p is not None for p in ergebnis):
                            ges = sum(ergebnis)
                            pct_raw = ges / ges_max * 100 if ges_max > 0 else 0
                            pct = NotenVerwaltung._round_pct(pct_raw)
                            note = NotenVerwaltung.ns_csv_lookup(pct, csv_str)
                            vals.append(f"{ges}/{ges_max}")
                            vals.append(f"{pct}")
                            n_str = (f"{note:.0f}" if note is not None and float(note).is_integer()
                                     else (f"{note:.1f}" if note is not None else "-"))
                            vals.append(n_str)
                        else:
                            vals.extend(["-", "-", "-"])
                        tree.insert("", tk.END, values=vals)
            else:
                tree["columns"] = ["info"]
                tree.heading("info", text=f"Keine {label_name} vorhanden")
                tree.column("info", width=400)
        finally:
            setattr(self, refreshing_attr, False)

    def _refresh_klausuren(self, keep_selection: Optional[int] = None) -> None:
        self._refresh_bewertung("klausur", keep_selection)

    def _refresh_ul(self, keep_selection: Optional[int] = None) -> None:
        self._refresh_bewertung("ul", keep_selection)

    def _refresh_gw_labels(self) -> None:
        gm = self.daten.gewichtung_muendlich
        gs = 100 - gm
        self.gw_ul_lbl.config(text=f"{gm}%")
        self.gw_sl.config(text=f"{gs}%")
        self.s_frame.config(text=f"Schriftliche Noten ({gs}%)")

    # ---- Events ----
    def _on_sj(self, e) -> None:
        self._refresh_kl()
        self._refresh_fach(None)
        self._refresh_sk(None)
        self._refresh_noten(None, None)
        self._refresh_klausuren()
        self._refresh_ul()

    def _on_kl(self, e) -> None:
        self._refresh_fach(self._kl())
        self._refresh_sk(self._kl())
        self._refresh_noten(None, None)
        self._refresh_klausuren()
        self._refresh_ul()

    def _on_fach(self, e) -> None:
        self._refresh_noten(self._kl(), self._sk())
        self._refresh_klausuren()
        self._refresh_ul()

    def _on_sk(self, e) -> None:
        self._refresh_noten(self._kl(), self._sk())

    def _on_klausur_select(self, e) -> None:
        if not getattr(self, '_refreshing_klausuren', False):
            self._refresh_klausuren()

    def _on_ul_select(self, e) -> None:
        if not getattr(self, '_refreshing_ul', False):
            self._refresh_ul()

    def _on_gw(self) -> None:
        try:
            w = int(self.gw_var.get())
        except ValueError:
            self.gw_var.set(str(self.daten.gewichtung_muendlich))
            return
        if 0 <= w <= 100:
            self.daten.gewichtung_muendlich = w
        else:
            self.gw_var.set(str(self.daten.gewichtung_muendlich))
            return
        self.gw_sl.config(text=f"{100 - self.daten.gewichtung_muendlich}%")
        self._refresh_gw_labels()
        self._save()
        self._refresh_noten(self._kl(), self._sk())

    # ---- CRUD Schuljahr / Klasse / Schülerin ----
    def _sj_add(self) -> None:
        s = simpledialog.askstring("Schuljahr hinzufügen", "Schuljahr (z.B. 2025/26):", parent=self.root)
        if s is None:
            return
        s = s.strip()
        if not s:
            messagebox.showwarning("Warnung", "Bitte ein Schuljahr eingeben!")
            return
        if not self.daten.schuljahr_hinzufuegen(s):
            messagebox.showwarning("Warnung", MSG_EXISTIERT_BEREITS.format(f"Schuljahr '{s}'"))
            return
        self._save()
        self._refresh_sj()
        self.sj_var.set(s)
        self._on_sj(None)

    def _sj_del(self) -> None:
        s = self._sj()
        if not s:
            messagebox.showwarning("Warnung", "Bitte ein Schuljahr auswählen!")
            return
        if not messagebox.askyesno("Bestätigung", f"Schuljahr '{s}' wirklich löschen?\nAlle Daten gehen verloren!"):
            return
        self.daten.schuljahr_loeschen(s)
        self._save()
        self._refresh_sj()

    def _kl_add(self) -> None:
        sj = self._sj()
        if not sj:
            messagebox.showwarning("Warnung", "Bitte zuerst ein Schuljahr auswählen!")
            return
        dlg = KlasseDialog(self.root)
        if dlg.result is None:
            return
        k, ns = dlg.result
        if not self.daten.klasse_hinzufuegen(sj, k, ns):
            messagebox.showwarning("Warnung", MSG_EXISTIERT_BEREITS.format(f"Klasse '{k}'"))
            return
        self._save()
        self._refresh_kl()
        # Neue Klasse im Dropdown auswählen
        ns_display = self.daten.get_notenschluessel(sj, k)
        self.kl_var.set(f"{k} [{ns_display}]")
        self._on_kl(None)

    def _kl_del(self) -> None:
        sj = self._sj()
        kl = self._kl()
        if not sj or not kl:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL)
            return
        kl_name = self._parse_kl_name(kl)
        if not messagebox.askyesno("Bestätigung", f"Klasse '{kl_name}' wirklich löschen?"):
            return
        self.daten.klasse_loeschen(sj, kl_name)
        self._save()
        self._refresh_kl()
        self._refresh_fach(None)
        self._refresh_sk(None)
        self._refresh_noten(None, None)
        self._refresh_klausuren()
        self._refresh_ul()

    def _kl_uebertragen(self) -> None:
        sj = self._sj()
        kl = self._kl()
        if not sj or not kl:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL)
            return
        kl_name = self._parse_kl_name(kl)
        if len(self.daten.schuljahre) < 2:
            if messagebox.askyesno("Neues Schuljahr", "Kein anderes Schuljahr vorhanden.\nNeues Schuljahr anlegen?"):
                ns = simpledialog.askstring("Neues Schuljahr", "Schuljahr (z.B. 2026/27):", parent=self.root)
                if ns and ns.strip() and self.daten.schuljahr_hinzufuegen(ns.strip()):
                    self._save()
                    self._refresh_sj()
            return
        dlg = _UebertragenDialog(self.root, sj, kl_name, self.daten.schuljahre)
        if dlg.result is None:
            return
        ziel = dlg.result
        if ziel not in self.daten.schuljahre:
            self.daten.schuljahr_hinzufuegen(ziel)
            self._save()
            self._refresh_sj()
        if not self.daten.klasse_uebertragen(sj, kl_name, ziel):
            messagebox.showwarning("Warnung", MSG_EXISTIERT_BEREITS.format(f"Klasse '{kl_name}' in Schuljahr '{ziel}'"))
            return
        self._save()
        messagebox.showinfo("OK", f"Klasse '{kl_name}' nach '{ziel}' übertragen (Schüler + Fächer, ohne Noten).")
        self._refresh_sj()
        self.sj_var.set(ziel)
        self._on_sj(None)

    def _sk_add(self) -> None:
        sj = self._sj()
        kl = self._kl()
        if not sj or not kl:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL)
            return
        kl_name = self._parse_kl_name(kl)
        dlg = SchuelerinDialog(self.root)
        if dlg.result is None:
            return
        nn, vn = dlg.result
        if not self.daten.schuelerin_hinzufuegen(sj, kl_name, nn, vn):
            messagebox.showwarning("Warnung", MSG_EXISTIERT_BEREITS.format(f"Schülerin '{nn}, {vn}'"))
            return
        self._save()
        self._refresh_sk(kl)

    def _sk_list_add(self) -> None:
        sj = self._sj()
        kl = self._kl()
        if not sj or not kl:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL)
            return
        kl_name = self._parse_kl_name(kl)
        dlg = SchuelerlisteDialog(self.root)
        if dlg.result is None:
            return
        added = skipped = 0
        for nn, vn in dlg.result:
            if self.daten.schuelerin_hinzufuegen(sj, kl_name, nn, vn):
                added += 1
            else:
                skipped += 1
        if added > 0:
            self._save()
        msg = f"{added} Schülerin(nen) hinzugefügt."
        if skipped > 0:
            msg += f"\n{skipped} bereits vorhanden."
        messagebox.showinfo("Ergebnis", msg)
        self._refresh_sk(kl)

    def _sk_del(self) -> None:
        sj = self._sj()
        kl = self._kl()
        if not sj or not kl:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL)
            return
        kl_name = self._parse_kl_name(kl)
        sk = self._sk()
        if not sk:
            messagebox.showwarning("Warnung", "Bitte eine Schülerin auswählen!")
            return
        if not messagebox.askyesno("Bestätigung", f"Schülerin '{sk}' wirklich löschen?"):
            return
        self.daten.schuelerin_loeschen(sj, kl_name, sk)
        self._save()
        self._refresh_sk(kl)
        self._refresh_noten(None, None)

    def _fach_add(self) -> None:
        sj = self._sj()
        kl = self._kl()
        if not sj or not kl:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL)
            return
        kl_name = self._parse_kl_name(kl)
        dlg = FachDialog(self.root)
        if dlg.result is None:
            return
        if not self.daten.fach_hinzufuegen(sj, kl_name, dlg.result):
            messagebox.showwarning("Warnung", MSG_EXISTIERT_BEREITS.format(f"Fach '{dlg.result}'"))
            return
        self._save()
        self._refresh_fach(kl)

    def _fach_del(self) -> None:
        sj = self._sj()
        kl = self._kl()
        if not sj or not kl:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL)
            return
        kl_name = self._parse_kl_name(kl)
        fach = self._fach()
        if not fach:
            messagebox.showwarning("Warnung", "Bitte ein Fach auswählen!")
            return
        if not messagebox.askyesno("Bestätigung", f"Fach '{fach}' wirklich löschen?\nAlle Noten und Klausuren gehen verloren!"):
            return
        self.daten.fach_loeschen(sj, kl_name, fach)
        self._save()
        self._refresh_fach(kl)
        self._refresh_noten(None, None)
        self._refresh_klausuren()
        self._refresh_ul()

    def _note_add(self, typ: str) -> None:
        sj = self._sj()
        hj = self._hj()
        kl = self._kl()
        sk = self._sk()
        fach = self._fach()
        if not sj or not kl or not sk or not fach:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_KL_FACH_SK)
            return
        kl_name = self._parse_kl_name(kl)
        nb = self.daten.get_notenbereich(sj, kl_name)
        sp = self.m_sp if typ == "muendlich" else self.s_sp
        try:
            note = int(sp.get())
        except ValueError:
            messagebox.showwarning("Warnung", f"Bitte eine gültige Note ({nb[0]}-{nb[1]}) eingeben!")
            return
        if not (nb[0] <= note <= nb[1]):
            messagebox.showwarning("Warnung", f"Die Note muss zwischen {nb[0]} und {nb[1]} liegen!")
            return
        self.daten.note_hinzufuegen(sj, kl_name, fach, sk, hj, typ, note)
        self._save()
        self._refresh_noten(kl, sk)

    def _note_del(self, typ: str) -> None:
        sj = self._sj()
        hj = self._hj()
        kl = self._kl()
        sk = self._sk()
        fach = self._fach()
        if not sj or not kl or not sk or not fach:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_KL_FACH_SK)
            return
        kl_name = self._parse_kl_name(kl)
        lb = self.m_lb if typ == "muendlich" else self.s_lb
        s = lb.curselection()
        if not s:
            messagebox.showwarning("Warnung", "Bitte eine Note zum Löschen auswählen!")
            return
        item_text = lb.get(s[0])
        # Bewertete Noten können nur über ihren Tab gelöscht werden
        if item_text.strip().startswith("UL:"):
            messagebox.showinfo("Hinweis", "Bewertete Unterrichtsleistungen können nur über den Tab 'Unterrichtsleistungen' gelöscht werden.")
            return
        if item_text.strip().startswith("K:"):
            messagebox.showinfo("Hinweis", "Klausurnoten können nur über den Tab 'Klausuren' gelöscht werden.")
            return
        if not messagebox.askyesno("Bestätigung", "Diese Note wirklich löschen?"):
            return
        # Zähle nur manuelle Noten bis zur ausgewählten Position
        manual_count = 0
        for i in range(s[0]):
            txt = lb.get(i).strip()
            if not txt.startswith(("K:", "UL:")):
                manual_count += 1
        self.daten.note_loeschen(sj, kl_name, fach, sk, hj, typ, manual_count)
        self._save()
        self._refresh_noten(kl, sk)

    # ---- Generische CRUD für Klausuren/UL ----
    def _bewertung_add(self, typ: str) -> None:
        sj = self._sj()
        hj = self._hj()
        kl = self._kl()
        fach = self._fach()
        if not sj or not kl or not fach:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL_FACH)
            return
        kl_name = self._parse_kl_name(kl)
        is_klausur = typ == "klausur"
        dlg_title = "Klausur hinzufügen" if is_klausur else "Unterrichtsleistung hinzufügen"
        warn_name = "Klausur" if is_klausur else "Unterrichtsleistung"
        add_fn = self.daten.klausur_hinzufuegen if is_klausur else self.daten.ul_hinzufuegen
        get_fn = self.daten.get_klausuren if is_klausur else self.daten.get_unterrichtsleistungen
        refresh_fn = self._refresh_klausuren if is_klausur else self._refresh_ul

        dlg = KlausurDialog(self.root, title=dlg_title)
        if dlg.result is None:
            return
        name, max_p = dlg.result
        if not add_fn(sj, kl_name, fach, hj, name, max_p):
            messagebox.showwarning("Warnung", MSG_EXISTIERT_BEREITS.format(f"{warn_name} '{name}'"))
            return
        self._save()
        items = get_fn(sj, kl_name, fach, hj)
        refresh_fn(keep_selection=len(items) - 1 if items else None)
        self._refresh_noten(self._kl(), self._sk())

    def _bewertung_del(self, typ: str) -> None:
        sj = self._sj()
        hj = self._hj()
        kl = self._kl()
        fach = self._fach()
        if not sj or not kl or not fach:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL_FACH)
            return
        kl_name = self._parse_kl_name(kl)
        is_klausur = typ == "klausur"
        lb = self.kl_klausur_lb if is_klausur else self.ul_lb
        get_fn = self.daten.get_klausuren if is_klausur else self.daten.get_unterrichtsleistungen
        del_fn = self.daten.klausur_loeschen if is_klausur else self.daten.ul_loeschen
        refresh_fn = self._refresh_klausuren if is_klausur else self._refresh_ul
        warn_name = "Klausur" if is_klausur else "Unterrichtsleistung"

        sel = lb.curselection()
        if not sel:
            messagebox.showwarning("Warnung", f"Bitte eine {warn_name} auswählen!")
            return
        idx = sel[0]
        item = get_fn(sj, kl_name, fach, hj)[idx]
        if not messagebox.askyesno("Bestätigung", f"{warn_name} '{item['name']}' wirklich löschen?"):
            return
        del_fn(sj, kl_name, fach, hj, idx)
        self._save()
        refresh_fn()
        self._refresh_noten(self._kl(), self._sk())

    def _bewertung_punkte(self, typ: str) -> None:
        sj = self._sj()
        hj = self._hj()
        kl = self._kl()
        fach = self._fach()
        if not sj or not kl or not fach:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL_FACH)
            return
        kl_name = self._parse_kl_name(kl)
        is_klausur = typ == "klausur"
        lb = self.kl_klausur_lb if is_klausur else self.ul_lb
        get_fn = self.daten.get_klausuren if is_klausur else self.daten.get_unterrichtsleistungen
        punkte_fn = self.daten.klausur_punkte_setzen if is_klausur else self.daten.ul_punkte_setzen
        refresh_fn = self._refresh_klausuren if is_klausur else self._refresh_ul
        dlg_title_prefix = "Punkte bearbeiten: Klausur" if is_klausur else "Punkte bearbeiten: Unterrichtsleistung"
        warn_name = "Klausur" if is_klausur else "Unterrichtsleistung"

        sel = lb.curselection()
        if not sel:
            messagebox.showwarning("Warnung", f"Bitte eine {warn_name} auswählen!")
            return
        idx = sel[0]
        item = get_fn(sj, kl_name, fach, hj)[idx]
        schuelerinnen = []
        for sk in self.daten.schuelerin_sortiert(sj, kl_name):
            d = self.daten.get_schueler_dict(sj, kl_name)[sk]
            schuelerinnen.append((sk, d["nachname"], d["vorname"]))
        if not schuelerinnen:
            messagebox.showinfo("Hinweis", "Keine Schülerinnen in dieser Klasse!")
            return
        csv_str = self.daten.get_ns_csv(sj, kl_name)
        ns_typ = self.daten.get_notenschluessel(sj, kl_name)
        dlg = PunkteDialog(self.root, dlg_title_prefix, item["name"],
                           item["max_punkte_pro_aufgabe"], schuelerinnen, item["ergebnisse"], csv_str,
                           notenschluessel_typ=ns_typ)
        if dlg.result is None:
            return
        saved = 0
        for sk, punkte in dlg.result.items():
            if punkte_fn(sj, kl_name, fach, hj, idx, sk, punkte):
                saved += 1
        if saved == 0 and dlg.result:
            messagebox.showwarning("Fehler", "Punkte konnten nicht gespeichert werden!", parent=self.root)
        self._save()
        refresh_fn()
        self._refresh_noten(self._kl(), self._sk())

    def _bewertung_gewichtung(self, typ: str) -> None:
        sj = self._sj()
        hj = self._hj()
        kl = self._kl()
        fach = self._fach()
        if not sj or not kl or not fach:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL_FACH)
            return
        kl_name = self._parse_kl_name(kl)
        is_klausur = typ == "klausur"
        lb = self.kl_klausur_lb if is_klausur else self.ul_lb
        get_fn = self.daten.get_klausuren if is_klausur else self.daten.get_unterrichtsleistungen
        gewichtung_fn = self.daten.klausur_gewichtung_setzen if is_klausur else self.daten.ul_gewichtung_setzen
        refresh_fn = self._refresh_klausuren if is_klausur else self._refresh_ul
        dlg_title = "Gewichtung ändern: Klausur" if is_klausur else "Gewichtung ändern: Unterrichtsleistung"
        warn_name = "Klausur" if is_klausur else "Unterrichtsleistung"

        sel = lb.curselection()
        if not sel:
            messagebox.showwarning("Warnung", f"Bitte eine {warn_name} auswählen!")
            return
        idx = sel[0]
        item = get_fn(sj, kl_name, fach, hj)[idx]
        if is_klausur:
            category_total = self.daten.schriftlich_prozent
            other_gw = self.daten.get_total_klausur_gewichtung(sj, kl_name, fach, hj, exclude_idx=idx)
        else:
            category_total = self.daten.ul_prozent
            other_gw = self.daten.get_total_ul_gewichtung(sj, kl_name, fach, hj, exclude_idx=idx)

        dlg = GewichtungDialog(self.root, dlg_title, item["name"],
                               item.get("gewichtung", 0), category_total, other_gw)
        if dlg.result is None:
            return
        gewichtung_fn(sj, kl_name, fach, hj, idx, dlg.result)
        self._save()
        refresh_fn()
        self._refresh_noten(self._kl(), self._sk())

    # ---- Spezifische Aufrufe ----
    def _klausur_add(self) -> None:
        self._bewertung_add("klausur")

    def _klausur_del(self) -> None:
        self._bewertung_del("klausur")

    def _klausur_punkte(self) -> None:
        self._bewertung_punkte("klausur")

    def _klausur_gewichtung(self) -> None:
        self._bewertung_gewichtung("klausur")

    def _ul_add(self) -> None:
        self._bewertung_add("ul")

    def _ul_del(self) -> None:
        self._bewertung_del("ul")

    def _ul_punkte(self) -> None:
        self._bewertung_punkte("ul")

    def _ul_gewichtung(self) -> None:
        self._bewertung_gewichtung("ul")

    def _ns_csv_edit(self) -> None:
        sj = self._sj()
        kl = self._kl()
        if not sj or not kl:
            messagebox.showwarning("Warnung", MSG_WAEHLEN_SJ_KL)
            return
        kl_name = self._parse_kl_name(kl)
        current_csv = self.daten.get_ns_csv(sj, kl_name)
        ns_typ = self.daten.get_notenschluessel(sj, kl_name)
        alle_klassen = [(s, k) for s, klasses in self.daten.schuljahre.items()
                        for k in klasses if not (s == sj and k == kl_name)]
        dlg = NotenschluesselCsvDialog(self.root, current_csv, ns_typ, alle_klassen)
        if dlg.result is None:
            return
        csv_str, transfer = dlg.result
        self.daten.set_ns_csv(sj, kl_name, csv_str)
        for s, k in transfer:
            self.daten.set_ns_csv(s, k, csv_str)
        self._save()
        self._refresh_klausuren()
        self._refresh_ul()
        self._refresh_noten(self._kl(), self._sk())


# ---------------------------------------------------------------------------
# Migration & Main
# ---------------------------------------------------------------------------
def _migrate_old_md() -> Optional[NotenVerwaltung]:
    md_file = os.path.join(_APP_DIR, "noten.md")
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
    root = tk.Tk()
    root.title("Notenverwaltung")
    root.geometry("1100x680")
    root.minsize(950, 580)
    migrated_data = _migrate_old_md()
    first_time = not os.path.exists(DATA_FILE)
    dlg = PasswordDialog(root, title="Passwort setzen" if first_time else "Passwort eingeben",
                         first_time=first_time)
    if dlg.result is None:
        root.destroy()
        return
    password = dlg.result
    app = NotenVerwaltungApp(root, password)
    if migrated_data and migrated_data.schuljahre:
        migrated_data.speichern_verschluesselt(password, app.data_file)
        app.daten = migrated_data
        app._refresh_sj()
    root.mainloop()


if __name__ == "__main__":
    main()