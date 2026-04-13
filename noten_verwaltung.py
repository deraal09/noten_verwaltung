#!/usr/bin/env python3
"""
Notenverwaltung - Programm zur Verwaltung von Schülerinnennoten
Daten werden verschlüsselt gespeichert (.ndat)
Export als Markdown (.md) oder CSV möglich
Notenschlüssel: BG (0-15) und IHK (1-6)
Klausuren mit Punkte-System und Notenschlüssel-CSV
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

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

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
    "IHK": "100,1;99,1.1;97,1.2;95,1.3;93,1.4;91,1.5;90,1.6;89,1.7;88,1.8;87,1.9;85,2;84,2.1;83,2.2;82,2.3;81,2.4;80,2.5;79,2.6;77,2.7;76,2.8;74,2.9;73,3;71,3.1;70,3.2;68,3.3;67,3.4;66,3.5;64,3.6;62,3.7;61,3.8;59,3.9;57,4;55,4.1;54,4.2;52,4.3;50,4.4;49,4.5;47,4.6;45,4.7;43,4.8;41,4.9;39,5;36,5.1;34,5.2;32,5.3;30,5.4;29,5.5;23,5.6;20,5.7;12,5.8;6,5.9;0,6",
    "BG": "95,15;94,14;89,13;84,12;79,11;74,10;69,9;64,8;59,7;54,6;49,5;44,4;39,3;32,2;26,1;19,0;0,0",
}
# Aliase für Abwärtskompatibilität (alte Daten mit langen Namen)
_NS_ALIASES = {"Berufsschule": "IHK", "Berufliches Gymnasium": "BG", "IHK": "IHK", "BG": "BG"}


def _derive_key(password, salt):
    return hashlib.pbkdf2_hmac('sha256', password.encode(), salt, ITERATIONS)


def _xor_encrypt(data_bytes, key):
    key_len = len(key)
    return bytes(b ^ key[i % key_len] for i, b in enumerate(data_bytes))


def encrypt_data(data_dict, password):
    json_bytes = json.dumps(data_dict, ensure_ascii=False).encode('utf-8')
    compressed = zlib.compress(json_bytes)
    salt = os.urandom(16)
    key = _derive_key(password, salt)
    encrypted = _xor_encrypt(compressed, key)
    verify = hashlib.sha256(key).digest()
    return base64.b64encode(salt + verify + encrypted)


def decrypt_data(raw_bytes, password):
    raw = base64.b64decode(raw_bytes)
    salt = raw[:16]; verify = raw[16:48]; encrypted = raw[48:]
    key = _derive_key(password, salt)
    if hashlib.sha256(key).digest() != verify:
        return None
    compressed = _xor_encrypt(encrypted, key)
    try:
        json_bytes = zlib.decompress(compressed)
        return json.loads(json_bytes.decode('utf-8'))
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Datenmodell
# ---------------------------------------------------------------------------
class NotenVerwaltung:
    def __init__(self):
        self.schuljahre = {}
        self.gewichtung_muendlich = DEFAULT_GEWICHTUNG

    # ---- Hilfsmethoden ----
    def _get_klasse(self, sj, k):
        return self.schuljahre.get(sj, {}).get(k)

    def _get_schueler_dict(self, sj, k):
        kl = self._get_klasse(sj, k)
        return kl.get("schuelerinnen", {}) if kl else None

    def get_notenschluessel(self, sj, k):
        kl = self._get_klasse(sj, k)
        ns = kl.get("notenschluessel", "IHK") if kl else "IHK"
        return _NS_ALIASES.get(ns, ns)

    def get_notenbereich(self, sj, k):
        return NOTENSCHLUESSEL.get(self.get_notenschluessel(sj, k), (1, 6))

    # ---- Notenschlüssel CSV ----
    @staticmethod
    def ns_csv_lookup(prozent, csv_str):
        """Schlägt eine Note anhand des Prozentsatzes im CSV-Schlüssel nach."""
        if not csv_str:
            return None
        try:
            entries = []
            for pair in csv_str.split(";"):
                parts = pair.strip().split(",")
                if len(parts) == 2:
                    entries.append((float(parts[0].strip()), float(parts[1].strip())))
            entries.sort(key=lambda x: x[0], reverse=True)
            for p, n in entries:
                if prozent >= p:
                    return n
            return entries[-1][1] if entries else None
        except Exception:
            return None

    @staticmethod
    def ns_csv_parse(csv_str):
        """Parst CSV-Notenschlüssel und gibt Liste von (prozent, note) zurück."""
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

    def get_ns_csv(self, sj, k):
        kl = self._get_klasse(sj, k)
        if kl is None:
            return DEFAULT_NS_CSV["IHK"]
        csv_str = kl.get("notenschluessel_csv", "")
        if not csv_str:
            ns = kl.get("notenschluessel", "IHK")
            ns = _NS_ALIASES.get(ns, ns)
            return DEFAULT_NS_CSV.get(ns, DEFAULT_NS_CSV["IHK"])
        return csv_str

    def set_ns_csv(self, sj, k, csv_str):
        kl = self._get_klasse(sj, k)
        if kl is not None:
            kl["notenschluessel_csv"] = csv_str

    # ---- Serialisierung ----
    def to_dict(self):
        sj = {}
        for s, klasses in self.schuljahre.items():
            sj[s] = {}
            for k, kl_data in klasses.items():
                sk_dict = {}
                for sk, d in kl_data.get("schuelerinnen", {}).items():
                    sk_dict[sk] = {"nachname": d["nachname"], "vorname": d["vorname"], "halbjahre": d["halbjahre"]}
                sj[s][k] = {
                    "notenschluessel": _NS_ALIASES.get(kl_data.get("notenschluessel", "IHK"), kl_data.get("notenschluessel", "IHK")),
                    "notenschluessel_csv": kl_data.get("notenschluessel_csv", ""),
                    "schuelerinnen": sk_dict,
                    "klausuren": kl_data.get("klausuren", {}),
                }
        return {"gewichtung_muendlich": self.gewichtung_muendlich, "schuljahre": sj}

    def from_dict(self, data):
        self.gewichtung_muendlich = data.get("gewichtung_muendlich", DEFAULT_GEWICHTUNG)
        self.schuljahre = {}
        for s, klasses in data.get("schuljahre", {}).items():
            self.schuljahre[s] = {}
            for k, kl_data in klasses.items():
                # Bestimme Datenformat (alt vs. neu)
                if isinstance(kl_data, dict) and "schuelerinnen" in kl_data:
                    # Neues Format: {notenschluessel, schuelerinnen, klausuren, ...}
                    ns = _NS_ALIASES.get(kl_data.get("notenschluessel", "IHK"), kl_data.get("notenschluessel", "IHK"))
                    ns_csv = kl_data.get("notenschluessel_csv", "")
                    schueler_raw = kl_data["schuelerinnen"]
                    klausuren = kl_data.get("klausuren", {})
                elif isinstance(kl_data, dict):
                    # Altes Format: kl_data ist direkt {schueler_key: {nachname, vorname, halbjahre}}
                    # Prüfe ob es Schülerdaten sind (hat "halbjahre" in den Werten)
                    first_val = next(iter(kl_data.values()), None)
                    if isinstance(first_val, dict) and "halbjahre" in first_val:
                        ns = "IHK"; ns_csv = ""; schueler_raw = kl_data; klausuren = {}
                    else:
                        # Unbekanntes Format - überspringen
                        logging.warning(f"Unbekanntes Datenformat für Klasse '{k}' - wird übersprungen")
                        continue
                else:
                    logging.warning(f"Ungültige Daten für Klasse '{k}' - wird übersprungen")
                    continue

                schueler = {}
                try:
                    for sk, d in schueler_raw.items():
                        if not isinstance(d, dict):
                            continue
                        hjs = {}
                        for hj in HALBJAHRE:
                            default_hj = {"muendlich": [], "schriftlich": []}
                            hj_data = d.get("halbjahre", {}).get(hj, default_hj)
                            if not isinstance(hj_data, dict):
                                hj_data = default_hj
                            hjs[hj] = {
                                "muendlich": hj_data.get("muendlich", []),
                                "schriftlich": hj_data.get("schriftlich", [])
                            }
                        schueler[sk] = {
                            "nachname": d.get("nachname", ""),
                            "vorname": d.get("vorname", ""),
                            "halbjahre": hjs
                        }
                except Exception as e:
                    logging.error(f"Fehler beim Laden der Schülerdaten für Klasse '{k}': {e}")
                    schueler = {}

                # Klausuren migrieren: ensure correct structure
                fixed_klausuren = {}
                try:
                    for hj, klist in klausuren.items():
                        fixed = []
                        for klausur in (klist if isinstance(klist, list) else []):
                            if isinstance(klausur, dict):
                                fixed.append({
                                    "name": klausur.get("name", ""),
                                    "max_punkte_pro_aufgabe": klausur.get("max_punkte_pro_aufgabe", []),
                                    "ergebnisse": klausur.get("ergebnisse", {}),
                                })
                        fixed_klausuren[hj] = fixed
                except Exception as e:
                    logging.error(f"Fehler beim Laden der Klausurdaten für Klasse '{k}': {e}")
                    fixed_klausuren = {}

                self.schuljahre[s][k] = {
                    "notenschluessel": ns,
                    "notenschluessel_csv": ns_csv,
                    "schuelerinnen": schueler,
                    "klausuren": fixed_klausuren
                }

    def speichern_verschluesselt(self, password, filepath=None):
        encrypted = encrypt_data(self.to_dict(), password)
        fp = filepath or DATA_FILE
        tmp_file = fp + ".tmp"
        with open(tmp_file, "wb") as f:
            f.write(encrypted)
        os.replace(tmp_file, fp)

    def laden_verschluesselt(self, password, filepath=None):
        fp = filepath or DATA_FILE
        if not os.path.exists(fp):
            return True
        with open(fp, "rb") as f:
            raw = f.read()
        data = decrypt_data(raw, password)
        if data is None:
            return False
        self.from_dict(data)
        return True

    # ---- Export ----
    def export_markdown(self, filepath):
        z = ["# Notenverwaltung", "", f"Gewichtung Mündlich: {self.gewichtung_muendlich}%", ""]
        for sj in sorted(self.schuljahre):
            z.append(f"## Schuljahr {sj}"); z.append("")
            for kn in sorted(self.schuljahre[sj]):
                ns = self.get_notenschluessel(sj, kn); nb = self.get_notenbereich(sj, kn)
                z.append(f"### Klasse {kn} [{ns} (Noten {nb[0]}-{nb[1]})]"); z.append("")
                # Notenschlüssel CSV
                csv_str = self.get_ns_csv(sj, kn)
                z.append(f"**Notenschlüssel:** `{csv_str}`"); z.append("")
                # Klausuren
                for hj in HALBJAHRE:
                    klausuren = self.get_klausuren(sj, kn, hj)
                    if klausuren:
                        z.append(f"**Klausuren {hj}:**"); z.append("")
                        for i, kl in enumerate(klausuren):
                            max_p = sum(kl["max_punkte_pro_aufgabe"])
                            z.append(f"- {kl['name']} (max. {max_p} Punkte, {len(kl['max_punkte_pro_aufgabe'])} Aufgaben)")
                        z.append("")
                for sk in self.schuelerin_sortiert(sj, kn):
                    d = self.schuljahre[sj][kn]["schuelerinnen"][sk]
                    z.append(f"#### {d['nachname']}, {d['vorname']}")
                    for hj in HALBJAHRE:
                        hd = d["halbjahre"][hj]
                        z.append(f"##### {hj}")
                        z.append(f"- Mündlich: {', '.join(str(n) for n in hd['muendlich'])}")
                        # Schriftlich: manuell + Klausur
                        manual = hd['schriftlich']
                        klausur_notes = self.get_klausur_noten(sj, kn, sk, hj)
                        if klausur_notes:
                            kn_str = ", ".join(str(n) for n in klausur_notes)
                            z.append(f"- Schriftlich (manuell): {', '.join(str(n) for n in manual)}")
                            z.append(f"- Schriftlich (Klausuren): {kn_str}")
                        else:
                            z.append(f"- Schriftlich: {', '.join(str(n) for n in manual)}")
                    z.append("")
                z.append("")
            z.append("")
        with open(filepath, "w", encoding="utf-8") as f:
            f.write("\n".join(z))

    def export_csv(self, filepath):
        with open(filepath, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Schuljahr", "Klasse", "Notenschlüssel", "Nachname", "Vorname", "Halbjahr",
                         "Mündliche Noten", "Schriftl. Noten (manuell)", "Schriftl. Noten (Klausuren)",
                         "Durchschnitt Mündlich", "Durchschnitt Schriftlich", "Gesamtnote Halbjahr", "Jahresnote"])
            for sj in sorted(self.schuljahre):
                for kn in sorted(self.schuljahre[sj]):
                    ns = self.get_notenschluessel(sj, kn)
                    for sk in self.schuelerin_sortiert(sj, kn):
                        d = self.schuljahre[sj][kn]["schuelerinnen"][sk]
                        for hj in HALBJAHRE:
                            hd = d["halbjahre"][hj]
                            dm = self.durchschnitt(hd["muendlich"])
                            kn_notes = self.get_klausur_noten(sj, kn, sk, hj)
                            all_schr = self.get_schriftlich_all(sj, kn, sk, hj)
                            ds = self.durchschnitt(all_schr)
                            gh = self.gesamtnote_hj(sj, kn, sk, hj)
                            w.writerow([sj, kn, ns, d["nachname"], d["vorname"], hj,
                                        " | ".join(str(n) for n in hd["muendlich"]),
                                        " | ".join(str(n) for n in hd["schriftlich"]),
                                        " | ".join(str(n) for n in kn_notes),
                                        f"{dm:.2f}" if dm else "-", f"{ds:.2f}" if ds else "-",
                                        f"{gh:.2f}" if gh else "-", ""])
                        gj = self.gesamtnote_jahr(sj, kn, sk)
                        w.writerow([sj, kn, ns, d["nachname"], d["vorname"], "Gesamt", "", "", "", "", "",
                                    f"{gj:.2f}" if gj else "-", ""])

    # ---- CRUD Schuljahr / Klasse / Schülerin ----
    def schuljahr_hinzufuegen(self, s):
        s = s.strip()
        if not s or s in self.schuljahre: return False
        self.schuljahre[s] = {}; return True

    def schuljahr_loeschen(self, s):
        if s in self.schuljahre: del self.schuljahre[s]; return True
        return False

    def klasse_hinzufuegen(self, sj, k, notenschluessel="IHK"):
        k = k.strip()
        if not k or sj not in self.schuljahre or k in self.schuljahre[sj]: return False
        self.schuljahre[sj][k] = {"notenschluessel": notenschluessel, "notenschluessel_csv": "", "schuelerinnen": {}, "klausuren": {}}
        return True

    def klasse_loeschen(self, sj, k):
        if sj in self.schuljahre and k in self.schuljahre[sj]: del self.schuljahre[sj][k]; return True
        return False

    @staticmethod
    def _key(nn, vn): return f"{nn}, {vn}"

    def _hj_neu(self): return {h: {"muendlich": [], "schriftlich": []} for h in HALBJAHRE}

    def schuelerin_hinzufuegen(self, sj, k, nn, vn):
        nn, vn = nn.strip(), vn.strip()
        if not nn or not vn: return False
        sd = self._get_schueler_dict(sj, k)
        if sd is None: return False
        key = self._key(nn, vn)
        if key in sd: return False
        sd[key] = {"nachname": nn, "vorname": vn, "halbjahre": self._hj_neu()}
        return True

    def schuelerin_loeschen(self, sj, k, sk):
        sd = self._get_schueler_dict(sj, k)
        if sd is not None and sk in sd: del sd[sk]; return True
        return False

    def schuelerin_sortiert(self, sj, k):
        sd = self._get_schueler_dict(sj, k)
        if sd is None: return []
        return sorted(sd.keys(), key=lambda x: (sd[x]["nachname"].lower(), sd[x]["vorname"].lower()))

    # ---- CRUD Noten ----
    def note_hinzufuegen(self, sj, k, sk, hj, typ, note):
        nb = self.get_notenbereich(sj, k)
        if not (nb[0] <= note <= nb[1]): return False
        sd = self._get_schueler_dict(sj, k)
        if sd is not None and sk in sd:
            d = sd[sk]
            if hj in d["halbjahre"]: d["halbjahre"][hj][typ].append(note); return True
        return False

    def note_loeschen(self, sj, k, sk, hj, typ, idx):
        sd = self._get_schueler_dict(sj, k)
        if sd is not None and sk in sd:
            d = sd[sk]
            if hj in d["halbjahre"]:
                n = d["halbjahre"][hj][typ]
                if 0 <= idx < len(n): n.pop(idx); return True
        return False

    # ---- CRUD Klausuren ----
    def klausur_hinzufuegen(self, sj, k, hj, name, max_punkte_pro_aufgabe):
        kl = self._get_klasse(sj, k)
        if kl is None: return False
        if "klausuren" not in kl: kl["klausuren"] = {}
        if hj not in kl["klausuren"]: kl["klausuren"][hj] = []
        for klausur in kl["klausuren"][hj]:
            if klausur["name"] == name: return False
        kl["klausuren"][hj].append({"name": name, "max_punkte_pro_aufgabe": max_punkte_pro_aufgabe, "ergebnisse": {}})
        return True

    def klausur_loeschen(self, sj, k, hj, idx):
        kl = self._get_klasse(sj, k)
        if kl is None: return False
        klist = kl.get("klausuren", {}).get(hj, [])
        if 0 <= idx < len(klist): klist.pop(idx); return True
        return False

    def get_klausuren(self, sj, k, hj):
        kl = self._get_klasse(sj, k)
        if kl is None: return []
        return kl.get("klausuren", {}).get(hj, [])

    def klausur_punkte_setzen(self, sj, k, hj, kidx, sk, punkte):
        kl = self._get_klasse(sj, k)
        if kl is None: return False
        klist = kl.get("klausuren", {}).get(hj, [])
        if not (0 <= kidx < len(klist)): return False
        klausur = klist[kidx]
        if len(punkte) != len(klausur["max_punkte_pro_aufgabe"]): return False
        for i, p in enumerate(punkte):
            if p is not None and (p < 0 or p > klausur["max_punkte_pro_aufgabe"][i]): return False
        klausur["ergebnisse"][sk] = punkte
        return True

    def klausur_note_berechnen(self, sj, k, hj, kidx, sk):
        kl = self._get_klasse(sj, k)
        if kl is None: return None
        klist = kl.get("klausuren", {}).get(hj, [])
        if not (0 <= kidx < len(klist)): return None
        klausur = klist[kidx]
        if sk not in klausur["ergebnisse"]: return None
        punkte = klausur["ergebnisse"][sk]
        if any(p is None for p in punkte): return None
        max_p = sum(klausur["max_punkte_pro_aufgabe"])
        if max_p == 0: return None
        prozent = sum(punkte) / max_p * 100
        return self.ns_csv_lookup(prozent, self.get_ns_csv(sj, k))

    def get_klausur_noten(self, sj, k, sk, hj):
        klausuren = self.get_klausuren(sj, k, hj)
        noten = []
        for i in range(len(klausuren)):
            note = self.klausur_note_berechnen(sj, k, hj, i, sk)
            if note is not None: noten.append(note)
        return noten

    def get_schriftlich_all(self, sj, k, sk, hj):
        sd = self._get_schueler_dict(sj, k)
        if sd is None or sk not in sd: return []
        manual = sd[sk].get("halbjahre", {}).get(hj, {}).get("schriftlich", [])
        return manual + self.get_klausur_noten(sj, k, sk, hj)

    # ---- Berechnungen ----
    @staticmethod
    def durchschnitt(noten):
        return round(sum(noten) / len(noten), 2) if noten else None

    def _gf(self): return self.gewichtung_muendlich / 100, (100 - self.gewichtung_muendlich) / 100

    def _gesamt(self, sm, ss):
        fm, fs = self._gf()
        if sm is not None and ss is not None: return round(sm * fm + ss * fs, 2)
        return sm if sm is not None else ss

    def gesamtnote_hj(self, sj, k, sk, hj):
        sd = self._get_schueler_dict(sj, k)
        if sd is None or sk not in sd: return None
        d = sd[sk].get("halbjahre", {}).get(hj)
        if not d: return None
        muendlich = d.get("muendlich", [])
        schriftlich = self.get_schriftlich_all(sj, k, sk, hj)
        return self._gesamt(self.durchschnitt(muendlich), self.durchschnitt(schriftlich))

    def gesamtnote_jahr(self, sj, k, sk):
        sd = self._get_schueler_dict(sj, k)
        if sd is None or sk not in sd: return None
        d = sd[sk]; gm, gs = [], []
        for h in HALBJAHRE:
            gm.extend(d["halbjahre"][h]["muendlich"])
            gs.extend(self.get_schriftlich_all(sj, k, sk, h))
        return self._gesamt(self.durchschnitt(gm), self.durchschnitt(gs))


# ---------------------------------------------------------------------------
# Dialoge
# ---------------------------------------------------------------------------
class _CenteredToplevel(tk.Toplevel):
    """Basisklasse für zentrierte Dialoge."""
    def _center(self):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw - self.winfo_width()) // 2}+{(sh - self.winfo_height()) // 2}")


class PasswordDialog(_CenteredToplevel):
    def __init__(self, parent, title="Passwort eingeben", first_time=False):
        super().__init__(parent); self.title(title); self.resizable(False, False)
        self.result = None; self.transient(parent); self.grab_set()
        f = ttk.Frame(self, padding=15); f.pack()
        msg = "Bitte vergeben Sie ein Passwort:" if first_time else "Bitte Passwort eingeben:"
        ttk.Label(f, text=msg).grid(row=0, column=0, columnspan=2, pady=(0, 10))
        ttk.Label(f, text="Passwort:").grid(row=1, column=0, sticky="w", pady=(0, 5))
        self.pw = ttk.Entry(f, show="*", width=25); self.pw.grid(row=1, column=1, pady=(0, 5), padx=(5, 0)); self.pw.focus_set()
        self.pw2 = None
        if first_time:
            ttk.Label(f, text="Bestätigen:").grid(row=2, column=0, sticky="w", pady=(0, 10))
            self.pw2 = ttk.Entry(f, show="*", width=25); self.pw2.grid(row=2, column=1, pady=(0, 10), padx=(5, 0))
        bf = ttk.Frame(f); bf.grid(row=3, column=0, columnspan=2)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.LEFT, padx=5)
        self.pw.bind("<Return>", lambda e: self.pw2.focus_set() if self.pw2 else self._ok())
        if self.pw2: self.pw2.bind("<Return>", lambda e: self._ok())
        self._center(); self.protocol("WM_DELETE_WINDOW", self._cancel); self.wait_window()

    def _ok(self):
        pw = self.pw.get()
        if not pw: return
        if self.pw2 and pw != self.pw2.get():
            messagebox.showwarning("Fehler", "Passwörter stimmen nicht überein!", parent=self); return
        self.result = pw; self.destroy()

    def _cancel(self): self.result = None; self.destroy()


class KlasseDialog(_CenteredToplevel):
    def __init__(self, parent):
        super().__init__(parent); self.title("Klasse hinzufügen"); self.resizable(False, False)
        self.result = None; self.transient(parent); self.grab_set()
        f = ttk.Frame(self, padding=15); f.pack()
        ttk.Label(f, text="Klassenname:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.e_name = ttk.Entry(f, width=25); self.e_name.grid(row=0, column=1, pady=(0, 5), padx=(5, 0)); self.e_name.focus_set()
        ttk.Label(f, text="Notenschlüssel:").grid(row=1, column=0, sticky="w", pady=(0, 5))
        self.ns_var = tk.StringVar(value="IHK")
        ns_frame = ttk.Frame(f); ns_frame.grid(row=1, column=1, sticky="w", pady=(0, 5), padx=(5, 0))
        for ns, nb in NOTENSCHLUESSEL.items():
            ttk.Radiobutton(ns_frame, text=f"{ns} (Noten {nb[0]}-{nb[1]})", variable=self.ns_var, value=ns).pack(anchor="w")
        ttk.Label(f, text="BG = Berufliches Gymnasium (0-15)\nIHK = Berufsschule (1-6)", foreground="gray").grid(row=2, column=0, columnspan=2, pady=(0, 5))
        bf = ttk.Frame(f); bf.grid(row=3, column=0, columnspan=2)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.LEFT, padx=5)
        self.e_name.bind("<Return>", lambda e: self._ok())
        self._center(); self.protocol("WM_DELETE_WINDOW", self._cancel); self.wait_window()

    def _ok(self):
        name = self.e_name.get().strip(); ns = self.ns_var.get()
        if name: self.result = (name, ns)
        self.destroy()

    def _cancel(self): self.result = None; self.destroy()


class SchuelerinDialog(_CenteredToplevel):
    def __init__(self, parent, title="Schülerin hinzufügen"):
        super().__init__(parent); self.title(title); self.resizable(False, False)
        self.result = None; self.transient(parent); self.grab_set()
        f = ttk.Frame(self, padding=15); f.pack()
        ttk.Label(f, text="Nachname:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.e_nn = ttk.Entry(f, width=30); self.e_nn.grid(row=0, column=1, pady=(0, 5), padx=(5, 0)); self.e_nn.focus_set()
        ttk.Label(f, text="Vorname:").grid(row=1, column=0, sticky="w", pady=(0, 10))
        self.e_vn = ttk.Entry(f, width=30); self.e_vn.grid(row=1, column=1, pady=(0, 10), padx=(5, 0))
        bf = ttk.Frame(f); bf.grid(row=2, column=0, columnspan=2)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.LEFT, padx=5)
        self.e_nn.bind("<Return>", lambda e: self.e_vn.focus_set())
        self.e_vn.bind("<Return>", lambda e: self._ok())
        self._center(); self.wait_window()

    def _ok(self):
        nn, vn = self.e_nn.get().strip(), self.e_vn.get().strip()
        if nn and vn: self.result = (nn, vn)
        self.destroy()

    def _cancel(self): self.result = None; self.destroy()


class SchuelerlisteDialog(_CenteredToplevel):
    """Dialog zum Hinzufügen mehrerer Schüler als Liste."""
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Schülerliste hinzufügen")
        self.geometry("500x450")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.result = None  # Liste von (nachname, vorname)

        f = ttk.Frame(self, padding=15)
        f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(f, text="Schüler als Liste eingeben", font=("TkDefaultFont", 11, "bold")).pack(anchor="w")
        ttk.Label(f, text="Format: Nachname, Vorname (eine Schülerin pro Zeile)", foreground="gray").pack(anchor="w", pady=(0, 10))

        # Textfeld für Liste
        ttk.Label(f, text="Schülerinnen:").pack(anchor="w")
        self.text = tk.Text(f, height=15, width=50, font=("Courier", 10))
        self.text.pack(fill=tk.BOTH, expand=True, pady=(5, 10))
        self.text.insert("1.0", "Müller, Anna\nSchmidt, Berta\nFischer, Christoph\n")

        # Info
        ttk.Label(f, text="Tipp: Liste kann aus Excel/CSV kopiert werden", foreground="gray", font=("TkDefaultFont", 8)).pack(anchor="w")

        # Buttons
        bf = ttk.Frame(f)
        bf.pack(fill=tk.X, pady=(15, 0))
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.RIGHT, padx=5)

        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _ok(self):
        content = self.text.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Warnung", "Bitte mindestens eine Schülerin eingeben!", parent=self)
            return
        schueler_liste = []
        errors = []
        for i, line in enumerate(content.split("\n"), 1):
            line = line.strip()
            if not line:
                continue
            # Versuche Nachname, Vorname zu parsen
            if "," in line:
                parts = line.split(",", 1)
                nn = parts[0].strip()
                vn = parts[1].strip() if len(parts) > 1 else ""
            else:
                # Vielleicht Tab-getrennt
                if "\t" in line:
                    parts = line.split("\t", 1)
                    nn = parts[0].strip()
                    vn = parts[1].strip() if len(parts) > 1 else ""
                else:
                    # Nur ein Name - als Nachname verwenden
                    nn = line
                    vn = ""
            if not nn:
                errors.append(f"Zeile {i}: Leer")
                continue
            schueler_liste.append((nn, vn))
        if errors:
            messagebox.showwarning("Warnung", f"Fehler beim Parsen:\n{chr(10).join(errors)}\n\n{len(schueler_liste)} Schülerin(nen) werden trotzdem hinzugefügt.", parent=self)
        if not schueler_liste:
            return
        self.result = schueler_liste
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


class KlausurDialog(_CenteredToplevel):
    """Dialog zum Hinzufügen einer Klausur."""
    def __init__(self, parent):
        super().__init__(parent); self.title("Klausur hinzufügen"); self.resizable(False, False)
        self.result = None; self.transient(parent); self.grab_set()
        f = ttk.Frame(self, padding=15); f.pack()
        ttk.Label(f, text="Name der Klausur:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.e_name = ttk.Entry(f, width=25); self.e_name.grid(row=0, column=1, pady=(0, 5), padx=(5, 0)); self.e_name.focus_set()
        ttk.Label(f, text="Anzahl Aufgaben:").grid(row=1, column=0, sticky="w", pady=(0, 5))
        self.e_anz = ttk.Spinbox(f, from_=1, to=20, width=5); self.e_anz.grid(row=1, column=1, sticky="w", pady=(0, 5), padx=(5, 0)); self.e_anz.set(3)
        ttk.Label(f, text="Max. Punkte pro Aufgabe\n(kommagetrennt, z.B. 10,10,10):", foreground="gray").grid(row=2, column=0, columnspan=2, sticky="w", pady=(5, 2))
        self.e_punkte = ttk.Entry(f, width=25); self.e_punkte.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        self.e_punkte.insert(0, "10,10,10")
        self.e_anz.bind("<Return>", lambda e: self.e_punkte.focus_set())
        self.e_punkte.bind("<Return>", lambda e: self._ok())
        bf = ttk.Frame(f); bf.grid(row=4, column=0, columnspan=2)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.LEFT, padx=5)
        self._center(); self.protocol("WM_DELETE_WINDOW", self._cancel); self.wait_window()

    def _ok(self):
        name = self.e_name.get().strip()
        if not name: messagebox.showwarning("Warnung", "Bitte einen Namen eingeben!", parent=self); return
        try:
            anz = int(self.e_anz.get())
        except ValueError:
            messagebox.showwarning("Warnung", "Ungültige Anzahl!", parent=self); return
        punkte_str = self.e_punkte.get().strip()
        try:
            max_p = [int(x.strip()) for x in punkte_str.split(",") if x.strip()]
        except ValueError:
            messagebox.showwarning("Warnung", "Ungültige Punkte-Eingabe! Zahlen kommagetrennt eingeben.", parent=self); return
        if len(max_p) != anz:
            messagebox.showwarning("Warnung", f"Anzahl der Punkte-Einträge ({len(max_p)}) stimmt nicht mit Aufgabenanzahl ({anz}) überein!", parent=self); return
        if any(p <= 0 for p in max_p):
            messagebox.showwarning("Warnung", "Punkte müssen > 0 sein!", parent=self); return
        self.result = (name, max_p); self.destroy()

    def _cancel(self): self.result = None; self.destroy()


class KlausurPunkteDialog(_CenteredToplevel):
    """Dialog zum Eintragen der Klausurpunkte für alle Schülerinnen."""
    def __init__(self, parent, klausur_name, max_punkte_pro_aufgabe, schuelerinnen, existing_ergebnisse, ns_csv_str):
        super().__init__(parent)
        self.title(f"Punkte bearbeiten: {klausur_name}")
        self.geometry("750x500"); self.minsize(600, 400)
        self.transient(parent); self.grab_set()
        self.max_punkte = max_punkte_pro_aufgabe
        self.schuelerinnen = schuelerinnen  # [(key, nachname, vorname), ...]
        self.existing_ergebnisse = existing_ergebnisse
        self.ns_csv = ns_csv_str
        self.result = None  # dict: sk -> [punkte_liste]

        # Header
        hf = ttk.Frame(self, padding=5); hf.pack(fill=tk.X)
        ttk.Label(hf, text=f"Klausur: {klausur_name}", font=("TkDefaultFont", 11, "bold")).pack(side=tk.LEFT)
        ges_max = sum(max_punkte_pro_aufgabe)
        ttk.Label(hf, text=f"  |  Max. Punkte: {ges_max} ({', '.join(str(p) for p in max_punkte_pro_aufgabe)})", foreground="gray").pack(side=tk.LEFT)

        # Scrollable area
        container = ttk.Frame(self); container.pack(fill=tk.BOTH, expand=True, padx=5)
        canvas = tk.Canvas(container); scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)
        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # Mouse wheel
        def _on_mousewheel(event): canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
        # Linux Mausrad
        canvas.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        # Header row
        headers = ["Schülerin"] + [f"A{i+1} (/{p})" for i, p in enumerate(max_punkte_pro_aufgabe)] + ["Gesamt", "%", "Note"]
        for c, h in enumerate(headers):
            ttk.Label(self.inner, text=h, font=("TkDefaultFont", 9, "bold")).grid(row=0, column=c, padx=2, pady=2, sticky="w")

        # Data rows
        self.entries = {}  # sk -> [Entry, ...]
        self.labels_ges = {}; self.labels_pct = {}; self.labels_note = {}
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
            lbl_g = ttk.Label(self.inner, text="", width=7); lbl_g.grid(row=r, column=len(max_punkte_pro_aufgabe) + 1, padx=2)
            lbl_p = ttk.Label(self.inner, text="", width=7); lbl_p.grid(row=r, column=len(max_punkte_pro_aufgabe) + 2, padx=2)
            lbl_n = ttk.Label(self.inner, text="", width=7, font=("TkDefaultFont", 9, "bold")); lbl_n.grid(row=r, column=len(max_punkte_pro_aufgabe) + 3, padx=2)
            self.labels_ges[r] = lbl_g; self.labels_pct[r] = lbl_p; self.labels_note[r] = lbl_n
            self._update_row(r)

        # Buttons
        bf = ttk.Frame(self, padding=5); bf.pack(fill=tk.X)
        ttk.Button(bf, text="Alle berechnen", command=self._update_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.RIGHT, padx=5)

        self._center(); self.protocol("WM_DELETE_WINDOW", self._cancel); self.wait_window()

    def _get_row_points(self, r):
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
                    pts.append(None)
        return pts

    def _update_row(self, r):
        pts = self._get_row_points(r)
        if all(p is not None for p in pts):
            ges = sum(pts)
            max_p = sum(self.max_punkte)
            pct = ges / max_p * 100 if max_p > 0 else 0
            note = NotenVerwaltung.ns_csv_lookup(pct, self.ns_csv)
            self.labels_ges[r].config(text=f"{ges}/{max_p}")
            self.labels_pct[r].config(text=f"{pct:.1f}%")
            note_str = f"{note:.0f}" if note is not None and note == int(note) else (f"{note:.1f}" if note is not None else "-")
            self.labels_note[r].config(text=note_str)
        else:
            self.labels_ges[r].config(text=""); self.labels_pct[r].config(text=""); self.labels_note[r].config(text="")

    def _update_all(self):
        for r in range(1, len(self.schuelerinnen) + 1):
            self._update_row(r)

    def _ok(self):
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
                            messagebox.showwarning("Warnung", f"Punkte für {nn}, {vn} Aufgabe {i+1} müssen zwischen 0 und {self.max_punkte[i]} liegen!", parent=self)
                            return
                        pts.append(p)
                    except ValueError:
                        messagebox.showwarning("Warnung", f"Ungültige Eingabe für {nn}, {vn} Aufgabe {i+1}!", parent=self)
                        return
            if all_filled:
                self.result[sk] = pts
            elif any(p is not None for p in pts):
                # Teilausfüllung: bestehende Ergebnisse + neue Einträge zusammenführen
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

    def _cancel(self): self.result = None; self.destroy()


class NotenschluesselCsvDialog(_CenteredToplevel):
    """Dialog zum Bearbeiten des Notenschlüssel-CSV und Übertragen auf andere Klassen."""
    def __init__(self, parent, current_csv, notenschluessel_typ, alle_klassen):
        """
        alle_klassen: [(sj, k_name), ...] Liste aller anderen Klassen
        """
        super().__init__(parent); self.title("Notenschlüssel bearbeiten"); self.resizable(True, True)
        self.geometry("550x500"); self.transient(parent); self.grab_set()
        self.alle_klassen = alle_klassen
        self.result = None  # (csv_str, [(sj, k), ...] für Übertragung)

        f = ttk.Frame(self, padding=10); f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(f, text="Notenschlüssel im CSV-Format:", font=("TkDefaultFont", 10, "bold")).pack(anchor="w")
        ttk.Label(f, text="Format: prozent,note;prozent,note;...  (absteigend nach %)", foreground="gray").pack(anchor="w", pady=(0, 5))

        # Standard-Buttons
        std_frame = ttk.Frame(f); std_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(std_frame, text="Standard laden:").pack(side=tk.LEFT, padx=(0, 5))
        for ns in NOTENSCHLUESSEL:
            ttk.Button(std_frame, text=f"{ns}", command=lambda n=ns: self._load_default(n), width=8).pack(side=tk.LEFT, padx=2)

        self.text = tk.Text(f, height=5, width=60, font=("Courier", 10)); self.text.pack(fill=tk.X, pady=(0, 5))
        self.text.insert("1.0", current_csv)

        # Vorschau
        ttk.Label(f, text="Vorschau:", font=("TkDefaultFont", 10, "bold")).pack(anchor="w", pady=(5, 2))
        self.preview = tk.Text(f, height=8, width=60, font=("Courier", 9), state="disabled", background="#f0f0f0")
        self.preview.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        ttk.Button(f, text="Vorschau aktualisieren", command=self._update_preview).pack(anchor="w", pady=(0, 10))

        # Übertragen
        ttk.Label(f, text="Auf andere Klassen übertragen:", font=("TkDefaultFont", 10, "bold")).pack(anchor="w", pady=(5, 2))
        self.transfer_vars = {}
        if alle_klassen:
            tf = ttk.Frame(f); tf.pack(fill=tk.X, pady=(0, 5))
            for sj, k in alle_klassen:
                var = tk.BooleanVar(value=False)
                self.transfer_vars[(sj, k)] = var
                ttk.Checkbutton(tf, text=f"{k} ({sj})", variable=var).pack(anchor="w")
        else:
            ttk.Label(f, text="(Keine anderen Klassen vorhanden)", foreground="gray").pack(anchor="w")

        # Buttons
        bf = ttk.Frame(f); bf.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.RIGHT, padx=5)

        self._update_preview()
        self._center(); self.protocol("WM_DELETE_WINDOW", self._cancel); self.wait_window()

    def _load_default(self, ns):
        self.text.delete("1.0", tk.END)
        self.text.insert("1.0", DEFAULT_NS_CSV.get(ns, ""))
        self._update_preview()

    def _update_preview(self):
        csv_str = self.text.get("1.0", tk.END).strip()
        entries = NotenVerwaltung.ns_csv_parse(csv_str)
        self.preview.config(state="normal"); self.preview.delete("1.0", tk.END)
        if not entries:
            self.preview.insert(tk.END, "Ungültiges Format oder leer.")
        else:
            self.preview.insert(tk.END, f"{'Prozent ≥':>12} | {'Note':>6}\n")
            self.preview.insert(tk.END, "-" * 25 + "\n")
            for p, n in entries:
                self.preview.insert(tk.END, f"{p:>10.1f}% | {n:>6}\n")
        self.preview.config(state="disabled")

    def _ok(self):
        csv_str = self.text.get("1.0", tk.END).strip()
        # Validieren
        entries = NotenVerwaltung.ns_csv_parse(csv_str)
        if not entries:
            messagebox.showwarning("Warnung", "Ungültiges Format! Bitte im Format prozent,note;... eingeben.", parent=self); return
        transfer = [(sj, k) for (sj, k), var in self.transfer_vars.items() if var.get()]
        self.result = (csv_str, transfer); self.destroy()

    def _cancel(self): self.result = None; self.destroy()


# ---------------------------------------------------------------------------
# Hauptanwendung
# ---------------------------------------------------------------------------
class NotenVerwaltungApp:
    def __init__(self, root, password, data_file=None):
        self.root = root; self.root.title("Notenverwaltung")
        self.root.geometry("1100x680"); self.root.minsize(950, 580)
        self.password = password
        self.data_file = data_file or DATA_FILE
        self.daten = NotenVerwaltung()
        if os.path.exists(self.data_file):
            if not self.daten.laden_verschluesselt(self.password, self.data_file):
                messagebox.showerror("Fehler", "Falsches Passwort oder beschädigte Daten!")
                self.root.after(100, self.root.destroy); return
        else:
            self.daten.speichern_verschluesselt(self.password, self.data_file)
        self._build_gui(); self._refresh_sj()
        self._update_title()

    def _save(self):
        self.daten.speichern_verschluesselt(self.password, self.data_file)

    def _update_title(self):
        fname = os.path.basename(self.data_file) if self.data_file else "noten.ndat"
        self.root.title(f"Notenverwaltung – {fname}")

    def _build_gui(self):
        gm = self.daten.gewichtung_muendlich; gs = 100 - gm
        sty = ttk.Style()
        sty.configure("H.TLabel", font=("TkDefaultFont", 11, "bold"))
        sty.configure("G.TLabel", font=("TkDefaultFont", 12, "bold"))
        sty.configure("J.TLabel", font=("TkDefaultFont", 11, "bold"), foreground="#2a5da8")
        sty.configure("I.TLabel", font=("TkDefaultFont", 9), foreground="gray")
        sty.configure("NS.TLabel", font=("TkDefaultFont", 9, "bold"), foreground="#c44")

        # Menü
        menubar = tk.Menu(self.root); self.root.config(menu=menubar)
        fm = tk.Menu(menubar, tearoff=0); menubar.add_cascade(label="Datei", menu=fm)
        fm.add_command(label="Öffnen...", command=self._file_open, accelerator="Strg+O")
        fm.add_command(label="Speichern unter...", command=self._file_save_as, accelerator="Strg+Shift+S")
        fm.add_separator()
        fm.add_command(label="Passwort ändern...", command=self._change_password)
        fm.add_separator()
        fm.add_command(label="Export Markdown...", command=self._export_md)
        fm.add_command(label="Export CSV (Excel)...", command=self._export_csv)
        fm.add_separator()
        fm.add_command(label="Beenden", command=self.root.quit)

        hf = ttk.Frame(self.root, padding=10); hf.pack(fill=tk.BOTH, expand=True)

        # Oben: Einstellungen
        top = ttk.LabelFrame(hf, text="Einstellungen", padding=5)
        top.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 5))
        ttk.Label(top, text="Schuljahr:").pack(side=tk.LEFT, padx=(0, 3))
        self.sj_var = tk.StringVar(); self.sj_cb = ttk.Combobox(top, textvariable=self.sj_var, state="readonly", width=12)
        self.sj_cb.pack(side=tk.LEFT, padx=(0, 5)); self.sj_cb.bind("<<ComboboxSelected>>", self._on_sj)
        ttk.Button(top, text="+", width=3, command=self._sj_add).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(top, text="−", width=3, command=self._sj_del).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Label(top, text="Halbjahr:").pack(side=tk.LEFT, padx=(5, 3))
        self.hj_var = tk.StringVar(); self.hj_cb = ttk.Combobox(top, textvariable=self.hj_var, state="readonly", width=12)
        self.hj_cb.pack(side=tk.LEFT, padx=(0, 10)); self.hj_cb['values'] = HALBJAHRE; self.hj_var.set(HALBJAHRE[0])
        self.hj_cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_all())
        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Label(top, text="Gewichtung Mündlich:").pack(side=tk.LEFT, padx=(5, 3))
        self.gw_var = tk.StringVar(value=str(gm))
        self.gw_sb = ttk.Spinbox(top, from_=0, to=100, width=4, textvariable=self.gw_var, command=self._on_gw)
        self.gw_sb.pack(side=tk.LEFT, padx=(0, 2))
        ttk.Label(top, text="%  /  Schriftlich:").pack(side=tk.LEFT)
        self.gw_sl = ttk.Label(top, text=f"{gs}%"); self.gw_sl.pack(side=tk.LEFT, padx=(0, 5))
        self.gw_sb.bind("<Return>", lambda e: self._on_gw())
        self.gw_sb.bind("<FocusOut>", lambda e: self._on_gw())

        # Links: Klassen
        kf = ttk.LabelFrame(hf, text="Klassen", padding=5); kf.grid(row=1, column=0, sticky="nsew", padx=(0, 5))
        self.kl_lb = tk.Listbox(kf, height=18, width=22, exportselection=False)
        self.kl_lb.pack(fill=tk.BOTH, expand=True); self.kl_lb.bind("<<ListboxSelect>>", self._on_kl)
        bf = ttk.Frame(kf); bf.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(bf, text="Hinzufügen", command=self._kl_add).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        ttk.Button(bf, text="Löschen", command=self._kl_del).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))

        # Mitte: Schülerinnen
        sf = ttk.LabelFrame(hf, text="Schülerinnen", padding=5); sf.grid(row=1, column=1, sticky="nsew", padx=5)
        self.sk_lb = tk.Listbox(sf, height=18, width=28, exportselection=False)
        self.sk_lb.pack(fill=tk.BOTH, expand=True); self.sk_lb.bind("<<ListboxSelect>>", self._on_sk)
        bf2 = ttk.Frame(sf); bf2.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(bf2, text="Hinzufügen", command=self._sk_add).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        ttk.Button(bf2, text="Liste", command=self._sk_list_add).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(bf2, text="Löschen", command=self._sk_del).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))

        # Rechts: Notebook mit Tabs
        self.nb = ttk.Notebook(hf); self.nb.grid(row=1, column=2, sticky="nsew", padx=(5, 0))
        self._build_noten_tab(gm, gs)
        self._build_klausuren_tab()

        hf.columnconfigure(0, weight=1); hf.columnconfigure(1, weight=2); hf.columnconfigure(2, weight=4)
        hf.rowconfigure(1, weight=1)

        # Tastenkürzel
        self.root.bind("<Control-o>", lambda e: self._file_open())
        self.root.bind("<Control-O>", lambda e: self._file_open())
        self.root.bind("<Control-Shift-s>", lambda e: self._file_save_as())
        self.root.bind("<Control-Shift-S>", lambda e: self._file_save_as())

    # ---- Noten-Tab ----
    def _build_noten_tab(self, gm, gs):
        nf = ttk.Frame(self.nb, padding=5); self.nb.add(nf, text="  Noten  ")
        self.info_lbl = ttk.Label(nf, text="Bitte eine Schülerin auswählen", style="H.TLabel")
        self.info_lbl.pack(anchor="w", pady=(0, 2))
        self.ns_lbl = ttk.Label(nf, text="", style="NS.TLabel")
        self.ns_lbl.pack(anchor="w", pady=(0, 5))

        self.m_frame = ttk.LabelFrame(nf, text=f"Mündliche Noten ({gm}%)", padding=5)
        self.m_frame.pack(fill=tk.BOTH, expand=True)
        self.m_lb = tk.Listbox(self.m_frame, height=5, exportselection=False)
        self.m_lb.pack(fill=tk.BOTH, expand=True)
        mbf = ttk.Frame(self.m_frame); mbf.pack(fill=tk.X, pady=(5, 0))
        self.m_sp = ttk.Spinbox(mbf, from_=1, to=6, width=5); self.m_sp.pack(side=tk.LEFT, padx=(0, 5)); self.m_sp.set(1)
        ttk.Button(mbf, text="Note eintragen", command=lambda: self._note_add("muendlich")).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(mbf, text="Löschen", command=lambda: self._note_del("muendlich")).pack(side=tk.LEFT)
        self.m_avg = ttk.Label(self.m_frame, text="Durchschnitt: -"); self.m_avg.pack(anchor="w", pady=(5, 0))

        self.s_frame = ttk.LabelFrame(nf, text=f"Schriftliche Noten ({gs}%)", padding=5)
        self.s_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        self.s_lb = tk.Listbox(self.s_frame, height=5, exportselection=False)
        self.s_lb.pack(fill=tk.BOTH, expand=True)
        sbf = ttk.Frame(self.s_frame); sbf.pack(fill=tk.X, pady=(5, 0))
        self.s_sp = ttk.Spinbox(sbf, from_=1, to=6, width=5); self.s_sp.pack(side=tk.LEFT, padx=(0, 5)); self.s_sp.set(1)
        ttk.Button(sbf, text="Note eintragen", command=lambda: self._note_add("schriftlich")).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(sbf, text="Löschen", command=lambda: self._note_del("schriftlich")).pack(side=tk.LEFT)
        ttk.Label(sbf, text="(Manuell – Klausurnoten siehe Tab Klausuren)", foreground="gray").pack(side=tk.LEFT, padx=(10, 0))
        self.s_avg = ttk.Label(self.s_frame, text="Durchschnitt: -"); self.s_avg.pack(anchor="w", pady=(5, 0))

        self.g_lbl = ttk.Label(nf, text="Gesamtnote: -", style="G.TLabel"); self.g_lbl.pack(anchor="w", pady=(10, 0))
        self.gw_info = ttk.Label(nf, text=f"({gm}% mündlich + {gs}% schriftlich)", style="I.TLabel"); self.gw_info.pack(anchor="w")
        self.j_lbl = ttk.Label(nf, text="Jahresnote: -", style="J.TLabel"); self.j_lbl.pack(anchor="w", pady=(8, 0))
        ttk.Label(nf, text="(Gesamtnote über beide Halbjahre)", style="I.TLabel").pack(anchor="w")

    # ---- Klausuren-Tab ----
    def _build_klausuren_tab(self):
        kf = ttk.Frame(self.nb, padding=5); self.nb.add(kf, text="  Klausuren  ")

        # Oben: Klausurliste
        top = ttk.Frame(kf); top.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(top, text="Klausuren:", style="H.TLabel").pack(side=tk.LEFT, padx=(0, 10))
        self.kl_klausur_lb = tk.Listbox(top, height=6, exportselection=False, width=40)
        self.kl_klausur_lb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.kl_klausur_lb.bind("<<ListboxSelect>>", self._on_klausur_select)
        btn_frame = ttk.Frame(top); btn_frame.pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Hinzufügen", command=self._klausur_add).pack(fill=tk.X, pady=1)
        ttk.Button(btn_frame, text="Löschen", command=self._klausur_del).pack(fill=tk.X, pady=1)
        ttk.Button(btn_frame, text="Punkte\nbearbeiten", command=self._klausur_punkte).pack(fill=tk.X, pady=1)

        # Notenschlüssel-Button
        ttk.Button(top, text="Noten-\nschlüssel", command=self._ns_csv_edit).pack(side=tk.LEFT, padx=(5, 0))

        # Unten: Treeview mit Ergebnissen
        self.kl_tree = ttk.Treeview(kf, columns=("info",), show="headings", height=10)
        self.kl_tree.heading("info", text="Keine Klausur ausgewählt")
        self.kl_tree.column("info", width=400)
        tree_scroll = ttk.Scrollbar(kf, orient=tk.VERTICAL, command=self.kl_tree.yview)
        self.kl_tree.configure(yscrollcommand=tree_scroll.set)
        self.kl_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # ---- Password / Export ----
    def _file_open(self):
        """Andere Datendatei öffnen."""
        fp = filedialog.askopenfilename(
            defaultextension=".ndat",
            filetypes=[("Notendatei", "*.ndat"), ("Alle Dateien", "*.*")],
            title="Notendatei öffnen",
            initialdir=os.path.dirname(self.data_file) if self.data_file else _APP_DIR
        )
        if not fp: return
        # Passwort für die neue Datei abfragen
        dlg = PasswordDialog(self.root, title="Passwort eingeben", first_time=False)
        if dlg.result is None: return
        new_pw = dlg.result
        new_data = NotenVerwaltung()
        if not new_data.laden_verschluesselt(new_pw, fp):
            messagebox.showerror("Fehler", "Falsches Passwort oder beschädigte Daten!")
            return
        self.password = new_pw
        self.data_file = fp
        self.daten = new_data
        self._update_title()
        self._refresh_sj()
        self._refresh_noten(None, None)
        self._refresh_klausuren()

    def _file_save_as(self):
        """Daten unter neuem Dateinamen speichern."""
        fp = filedialog.asksaveasfilename(
            defaultextension=".ndat",
            filetypes=[("Notendatei", "*.ndat"), ("Alle Dateien", "*.*")],
            title="Notendatei speichern unter",
            initialdir=os.path.dirname(self.data_file) if self.data_file else _APP_DIR,
            initialfile=os.path.basename(self.data_file) if self.data_file else "noten.ndat"
        )
        if not fp: return
        self.data_file = fp
        self._save()
        self._update_title()

    def _change_password(self):
        dlg = PasswordDialog(self.root, title="Passwort ändern", first_time=False)
        if dlg.result is None: return
        self.password = dlg.result; self._save(); messagebox.showinfo("OK", "Passwort geändert und Daten gespeichert.")

    def _export_md(self):
        fp = filedialog.asksaveasfilename(defaultextension=".md", filetypes=[("Markdown", "*.md")], title="Export Markdown")
        if fp: self.daten.export_markdown(fp); messagebox.showinfo("Export", f"Markdown exportiert nach:\n{fp}")

    def _export_csv(self):
        fp = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], title="Export CSV")
        if fp: self.daten.export_csv(fp); messagebox.showinfo("Export", f"CSV exportiert nach:\n{fp}")

    # ---- Getters ----
    def _sj(self): return self.sj_var.get() or None
    def _hj(self): return self.hj_var.get() or HALBJAHRE[0]
    def _kl(self):
        s = self.kl_lb.curselection(); return self.kl_lb.get(s[0]) if s else None
    def _sk(self):
        s = self.sk_lb.curselection(); return self.sk_lb.get(s[0]) if s else None

    @staticmethod
    def _parse_kl_name(kl_display):
        if kl_display is None: return None
        if " [" in kl_display: return kl_display[:kl_display.rfind(" [")]
        return kl_display

    # ---- Refresh ----
    def _refresh_all(self):
        self._refresh_noten(self._kl(), self._sk())
        self._refresh_klausuren()

    def _refresh_sj(self):
        sl = sorted(self.daten.schuljahre.keys()); self.sj_cb['values'] = sl
        if sl:
            if self._sj() not in sl: self.sj_var.set(sl[0])
            self._refresh_kl()
        else:
            self.sj_var.set(""); self.kl_lb.delete(0, tk.END); self._refresh_sk(None); self._refresh_noten(None, None)

    def _refresh_kl(self):
        self.kl_lb.delete(0, tk.END); sj = self._sj()
        if sj and sj in self.daten.schuljahre:
            for k in sorted(self.daten.schuljahre[sj]):
                ns = self.daten.get_notenschluessel(sj, k)
                self.kl_lb.insert(tk.END, f"{k} [{ns}]")

    def _refresh_sk(self, kl):
        self.sk_lb.delete(0, tk.END); sj = self._sj()
        kl_name = self._parse_kl_name(kl) if kl else None
        if sj and kl_name and kl_name in self.daten.schuljahre.get(sj, {}):
            for sk in self.daten.schuelerin_sortiert(sj, kl_name): self.sk_lb.insert(tk.END, sk)

    def _refresh_noten(self, kl, sk):
        self.m_lb.delete(0, tk.END); self.s_lb.delete(0, tk.END)
        if not kl or not sk:
            self.info_lbl.config(text="Bitte eine Schülerin auswählen"); self.ns_lbl.config(text="")
            self.m_avg.config(text="Durchschnitt: -"); self.s_avg.config(text="Durchschnitt: -")
            self.g_lbl.config(text="Gesamtnote: -"); self.j_lbl.config(text="Jahresnote: -"); return
        sj, hj = self._sj(), self._hj()
        kl_name = self._parse_kl_name(kl)
        sd = self.daten._get_schueler_dict(sj, kl_name)
        if sd is None or sk not in sd: return
        d = sd[sk]; ns = self.daten.get_notenschluessel(sj, kl_name); nb = self.daten.get_notenbereich(sj, kl_name)
        hd = d["halbjahre"][hj]
        self.info_lbl.config(text=f"Noten für: {d['nachname']}, {d['vorname']} ({kl_name}) — {hj}")
        self.ns_lbl.config(text=f"Notenschlüssel: {ns} (Noten {nb[0]}–{nb[1]})")
        # Mündlich
        for i, n in enumerate(hd["muendlich"]): self.m_lb.insert(tk.END, f"{i+1}. Note: {n}")
        sm = self.daten.durchschnitt(hd["muendlich"])
        self.m_avg.config(text=f"Durchschnitt: {sm:.2f}" if sm is not None else "Durchschnitt: -")
        # Schriftlich: manuell + Klausur
        for i, n in enumerate(hd["schriftlich"]): self.s_lb.insert(tk.END, f"{i+1}. Note: {n}")
        kn_notes = self.daten.get_klausur_noten(sj, kl_name, sk, hj)
        klausuren = self.daten.get_klausuren(sj, kl_name, hj)
        for i, n in enumerate(kn_notes):
            kname = klausuren[i]["name"] if i < len(klausuren) else f"K{i+1}"
            n_str = f"{n:.0f}" if n == int(n) else f"{n:.1f}"
            self.s_lb.insert(tk.END, f"  K: {kname} → {n_str}")
        all_schr = self.daten.get_schriftlich_all(sj, kl_name, sk, hj)
        ss = self.daten.durchschnitt(all_schr)
        self.s_avg.config(text=f"Durchschnitt: {ss:.2f}" if ss is not None else "Durchschnitt: -")
        # Gesamtnote
        gn = self.daten.gesamtnote_hj(sj, kl_name, sk, hj)
        self.g_lbl.config(text=f"Gesamtnote ({hj}): {gn:.2f}" if gn is not None else f"Gesamtnote ({hj}): -")
        jn = self.daten.gesamtnote_jahr(sj, kl_name, sk)
        self.j_lbl.config(text=f"Jahresnote: {jn:.2f}" if jn is not None else "Jahresnote: -")
        # Spinbox
        self.m_sp.config(from_=nb[0], to=nb[1]); self.m_sp.set(nb[0])
        self.s_sp.config(from_=nb[0], to=nb[1]); self.s_sp.set(nb[0])

    def _refresh_klausuren(self, keep_selection=None):
        """Aktualisiert die Klausurliste und -tabelle.
        keep_selection: Index der Klausur, die ausgewählt bleiben soll (oder None für aktuelle Auswahl)
        """
        self._refreshing_klausuren = True
        try:
            self._refresh_klausuren_impl(keep_selection)
        finally:
            self._refreshing_klausuren = False

    def _refresh_klausuren_impl(self, keep_selection):
        # Aktuelle Auswahl merken
        if keep_selection is None:
            sel = self.kl_klausur_lb.curselection()
            keep_selection = sel[0] if sel else None

        self.kl_klausur_lb.delete(0, tk.END)
        sj, hj, kl = self._sj(), self._hj(), self._kl()
        kl_name = self._parse_kl_name(kl) if kl else None
        if not sj or not kl_name: return
        klausuren = self.daten.get_klausuren(sj, kl_name, hj)
        for i, k in enumerate(klausuren):
            max_p = sum(k["max_punkte_pro_aufgabe"])
            anzahl = len(k["max_punkte_pro_aufgabe"])
            self.kl_klausur_lb.insert(tk.END, f"{k['name']} (max {max_p} P., {anzahl} Aufg.)")

        # Auswahl wiederherstellen
        if keep_selection is not None and keep_selection < len(klausuren):
            self.kl_klausur_lb.selection_set(keep_selection)
            self.kl_klausur_lb.see(keep_selection)

        # Treeview leeren
        for item in self.kl_tree.get_children():
            self.kl_tree.delete(item)
        # Spalten konfigurieren
        if klausuren:
            sel = self.kl_klausur_lb.curselection()
            if sel:
                kidx = sel[0]
                klausur = klausuren[kidx]
                max_p = klausur["max_punkte_pro_aufgabe"]
                cols = ["schuelerin"] + [f"a{i}" for i in range(len(max_p))] + ["gesamt", "prozent", "note"]
                self.kl_tree["columns"] = cols
                self.kl_tree.heading("schuelerin", text="Schülerin")
                self.kl_tree.column("schuelerin", width=150)
                for i, mp in enumerate(max_p):
                    self.kl_tree.heading(f"a{i}", text=f"A{i+1} (/{mp})")
                    self.kl_tree.column(f"a{i}", width=60, anchor="center")
                self.kl_tree.heading("gesamt", text="Gesamt")
                self.kl_tree.column("gesamt", width=60, anchor="center")
                self.kl_tree.heading("prozent", text="%")
                self.kl_tree.column("prozent", width=55, anchor="center")
                self.kl_tree.heading("note", text="Note")
                self.kl_tree.column("note", width=55, anchor="center")
                # Daten einfügen
                csv_str = self.daten.get_ns_csv(sj, kl_name)
                ges_max = sum(max_p)
                for sk in self.daten.schuelerin_sortiert(sj, kl_name):
                    d = self.daten._get_schueler_dict(sj, kl_name)[sk]
                    vals = [f"{d['nachname']}, {d['vorname']}"]
                    ergebnis = klausur["ergebnisse"].get(sk, [])
                    for i in range(len(max_p)):
                        vals.append(str(ergebnis[i]) if i < len(ergebnis) and ergebnis[i] is not None else "-")
                    if ergebnis and all(p is not None for p in ergebnis):
                        ges = sum(ergebnis)
                        pct = ges / ges_max * 100 if ges_max > 0 else 0
                        note = NotenVerwaltung.ns_csv_lookup(pct, csv_str)
                        vals.append(f"{ges}/{ges_max}")
                        vals.append(f"{pct:.1f}")
                        n_str = f"{note:.0f}" if note is not None and note == int(note) else (f"{note:.1f}" if note is not None else "-")
                        vals.append(n_str)
                    else:
                        vals.extend(["-", "-", "-"])
                    self.kl_tree.insert("", tk.END, values=vals)
        else:
            self.kl_tree["columns"] = ["info"]
            self.kl_tree.heading("info", text="Keine Klausuren vorhanden")
            self.kl_tree.column("info", width=400)

    def _refresh_gw_labels(self):
        gm = self.daten.gewichtung_muendlich; gs = 100 - gm
        self.m_frame.config(text=f"Mündliche Noten ({gm}%)")
        self.s_frame.config(text=f"Schriftliche Noten ({gs}%)")
        self.gw_info.config(text=f"({gm}% mündlich + {gs}% schriftlich)")

    # ---- Events ----
    def _on_sj(self, e): self._refresh_kl(); self._refresh_sk(None); self._refresh_noten(None, None); self._refresh_klausuren()
    def _on_kl(self, e): self._refresh_sk(self._kl()); self._refresh_noten(None, None); self._refresh_klausuren()
    def _on_sk(self, e): self._refresh_noten(self._kl(), self._sk())
    def _on_klausur_select(self, e):
        if not getattr(self, '_refreshing_klausuren', False):
            self._refresh_klausuren()

    def _on_gw(self):
        try:
            w = int(self.gw_var.get())
            if 0 <= w <= 100: self.daten.gewichtung_muendlich = w
            else: self.gw_var.set(str(self.daten.gewichtung_muendlich)); return
        except ValueError: self.gw_var.set(str(self.daten.gewichtung_muendlich)); return
        self.gw_sl.config(text=f"{100 - self.daten.gewichtung_muendlich}%")
        self._refresh_gw_labels(); self._save(); self._refresh_noten(self._kl(), self._sk())

    # ---- CRUD Schuljahr / Klasse / Schülerin ----
    def _sj_add(self):
        s = simpledialog.askstring("Schuljahr hinzufügen", "Schuljahr (z.B. 2025/26):", parent=self.root)
        if s is None: return; s = s.strip()
        if not s: messagebox.showwarning("Warnung", "Bitte ein Schuljahr eingeben!"); return
        if not self.daten.schuljahr_hinzufuegen(s): messagebox.showwarning("Warnung", f"Schuljahr '{s}' existiert bereits!"); return
        self._save(); self._refresh_sj(); self.sj_var.set(s); self._on_sj(None)

    def _sj_del(self):
        s = self._sj()
        if not s: messagebox.showwarning("Warnung", "Bitte ein Schuljahr auswählen!"); return
        if not messagebox.askyesno("Bestätigung", f"Schuljahr '{s}' wirklich löschen?\nAlle Daten gehen verloren!"): return
        self.daten.schuljahr_loeschen(s); self._save(); self._refresh_sj()

    def _kl_add(self):
        sj = self._sj()
        if not sj: messagebox.showwarning("Warnung", "Bitte zuerst ein Schuljahr auswählen!"); return
        dlg = KlasseDialog(self.root)
        if dlg.result is None: return
        k, ns = dlg.result
        if not k: messagebox.showwarning("Warnung", "Bitte einen Klassennamen eingeben!"); return
        if not self.daten.klasse_hinzufuegen(sj, k, ns): messagebox.showwarning("Warnung", f"Klasse '{k}' existiert bereits!"); return
        self._save(); self._refresh_kl()

    def _kl_del(self):
        sj, kl = self._sj(), self._kl()
        if not sj: messagebox.showwarning("Warnung", "Bitte zuerst ein Schuljahr auswählen!"); return
        if not kl: messagebox.showwarning("Warnung", "Bitte eine Klasse auswählen!"); return
        kl_name = self._parse_kl_name(kl)
        if not messagebox.askyesno("Bestätigung", f"Klasse '{kl_name}' wirklich löschen?\nAlle Schülerinnen und Noten gehen verloren!"): return
        self.daten.klasse_loeschen(sj, kl_name); self._save(); self._refresh_kl(); self._refresh_sk(None); self._refresh_noten(None, None); self._refresh_klausuren()

    def _sk_add(self):
        sj, kl = self._sj(), self._kl()
        if not sj: messagebox.showwarning("Warnung", "Bitte zuerst ein Schuljahr auswählen!"); return
        if not kl: messagebox.showwarning("Warnung", "Bitte zuerst eine Klasse auswählen!"); return
        kl_name = self._parse_kl_name(kl)
        dlg = SchuelerinDialog(self.root)
        if dlg.result is None: return
        nn, vn = dlg.result
        if not self.daten.schuelerin_hinzufuegen(sj, kl_name, nn, vn):
            messagebox.showwarning("Warnung", f"Schülerin '{nn}, {vn}' existiert bereits!"); return
        self._save(); self._refresh_sk(kl)

    def _sk_list_add(self):
        """Mehrere Schülerinnen als Liste hinzufügen."""
        sj, kl = self._sj(), self._kl()
        if not sj: messagebox.showwarning("Warnung", "Bitte zuerst ein Schuljahr auswählen!"); return
        if not kl: messagebox.showwarning("Warnung", "Bitte zuerst eine Klasse auswählen!"); return
        kl_name = self._parse_kl_name(kl)
        dlg = SchuelerlisteDialog(self.root)
        if dlg.result is None: return
        added = 0
        skipped = 0
        for nn, vn in dlg.result:
            if self.daten.schuelerin_hinzufuegen(sj, kl_name, nn, vn):
                added += 1
            else:
                skipped += 1
        if added > 0:
            self._save()
        if skipped > 0:
            messagebox.showinfo("Ergebnis", f"{added} Schülerin(nen) hinzugefügt.\n{skipped} bereits vorhanden (übersprungen).")
        elif added > 0:
            messagebox.showinfo("Ergebnis", f"{added} Schülerin(nen) hinzugefügt.")
        else:
            messagebox.showinfo("Ergebnis", "Keine neuen Schülerinnen hinzugefügt.")
        self._refresh_sk(kl)

    def _sk_del(self):
        sj, kl = self._sj(), self._kl()
        if not sj: messagebox.showwarning("Warnung", "Bitte zuerst ein Schuljahr auswählen!"); return
        if not kl: messagebox.showwarning("Warnung", "Bitte zuerst eine Klasse auswählen!"); return
        kl_name = self._parse_kl_name(kl); sk = self._sk()
        if not sk: messagebox.showwarning("Warnung", "Bitte eine Schülerin auswählen!"); return
        if not messagebox.askyesno("Bestätigung", f"Schülerin '{sk}' wirklich löschen?\nAlle Noten gehen verloren!"): return
        self.daten.schuelerin_loeschen(sj, kl_name, sk); self._save(); self._refresh_sk(kl); self._refresh_noten(None, None)

    # ---- CRUD Noten ----
    def _note_add(self, typ):
        sj, hj, kl, sk = self._sj(), self._hj(), self._kl(), self._sk()
        if not sj or not kl or not sk: messagebox.showwarning("Warnung", "Bitte Schuljahr, Klasse und Schülerin auswählen!"); return
        kl_name = self._parse_kl_name(kl)
        nb = self.daten.get_notenbereich(sj, kl_name)
        sp = self.m_sp if typ == "muendlich" else self.s_sp
        try: note = int(sp.get())
        except ValueError: messagebox.showwarning("Warnung", f"Bitte eine gültige Note ({nb[0]}-{nb[1]}) eingeben!"); return
        if not (nb[0] <= note <= nb[1]):
            messagebox.showwarning("Warnung", f"Die Note muss zwischen {nb[0]} und {nb[1]} liegen!"); return
        self.daten.note_hinzufuegen(sj, kl_name, sk, hj, typ, note); self._save(); self._refresh_noten(kl, sk)

    def _note_del(self, typ):
        sj, hj, kl, sk = self._sj(), self._hj(), self._kl(), self._sk()
        if not sj or not kl or not sk: messagebox.showwarning("Warnung", "Bitte Schuljahr, Klasse und Schülerin auswählen!"); return
        kl_name = self._parse_kl_name(kl)
        lb = self.m_lb if typ == "muendlich" else self.s_lb
        s = lb.curselection()
        if not s: messagebox.showwarning("Warnung", "Bitte eine Note zum Löschen auswählen!"); return
        # Prüfen ob es eine Klausurnote ist (kann nicht gelöscht werden)
        item_text = lb.get(s[0])
        if item_text.strip().startswith("K:"):
            messagebox.showinfo("Hinweis", "Klausurnoten können nur über den Tab 'Klausuren' gelöscht werden."); return
        if not messagebox.askyesno("Bestätigung", "Diese Note wirklich löschen?"): return
        # Index in der manuellen Liste berechnen
        manual_count = 0
        for i in range(s[0]):
            if not lb.get(i).strip().startswith("K:"): manual_count += 1
        self.daten.note_loeschen(sj, kl_name, sk, hj, typ, manual_count); self._save(); self._refresh_noten(kl, sk)

    # ---- CRUD Klausuren ----
    def _klausur_add(self):
        sj, hj, kl = self._sj(), self._hj(), self._kl()
        if not sj or not kl: messagebox.showwarning("Warnung", "Bitte Schuljahr und Klasse auswählen!"); return
        kl_name = self._parse_kl_name(kl)
        dlg = KlausurDialog(self.root)
        if dlg.result is None: return
        name, max_p = dlg.result
        if not self.daten.klausur_hinzufuegen(sj, kl_name, hj, name, max_p):
            messagebox.showwarning("Warnung", f"Klausur '{name}' existiert bereits!"); return
        self._save()
        klausuren = self.daten.get_klausuren(sj, kl_name, hj)
        new_idx = len(klausuren) - 1 if klausuren else None
        self._refresh_klausuren(keep_selection=new_idx)
        self._refresh_noten(self._kl(), self._sk())

    def _klausur_del(self):
        sj, hj, kl = self._sj(), self._hj(), self._kl()
        if not sj or not kl: messagebox.showwarning("Warnung", "Bitte Schuljahr und Klasse auswählen!"); return
        kl_name = self._parse_kl_name(kl)
        sel = self.kl_klausur_lb.curselection()
        if not sel: messagebox.showwarning("Warnung", "Bitte eine Klausur auswählen!"); return
        kidx = sel[0]
        klausur = self.daten.get_klausuren(sj, kl_name, hj)[kidx]
        if not messagebox.askyesno("Bestätigung", f"Klausur '{klausur['name']}' wirklich löschen?"): return
        self.daten.klausur_loeschen(sj, kl_name, hj, kidx); self._save(); self._refresh_klausuren(); self._refresh_noten(self._kl(), self._sk())

    def _klausur_punkte(self):
        sj, hj, kl = self._sj(), self._hj(), self._kl()
        if not sj or not kl: messagebox.showwarning("Warnung", "Bitte Schuljahr und Klasse auswählen!"); return
        kl_name = self._parse_kl_name(kl)
        sel = self.kl_klausur_lb.curselection()
        if not sel: messagebox.showwarning("Warnung", "Bitte eine Klausur auswählen!"); return
        kidx = sel[0]
        klausur = self.daten.get_klausuren(sj, kl_name, hj)[kidx]
        # Schülerinnenliste vorbereiten
        schuelerinnen = []
        for sk in self.daten.schuelerin_sortiert(sj, kl_name):
            d = self.daten._get_schueler_dict(sj, kl_name)[sk]
            schuelerinnen.append((sk, d["nachname"], d["vorname"]))
        if not schuelerinnen:
            messagebox.showinfo("Hinweis", "Keine Schülerinnen in dieser Klasse!"); return
        csv_str = self.daten.get_ns_csv(sj, kl_name)
        dlg = KlausurPunkteDialog(self.root, klausur["name"], klausur["max_punkte_pro_aufgabe"],
                                  schuelerinnen, klausur["ergebnisse"], csv_str)
        if dlg.result is None: return
        saved = 0
        for sk, punkte in dlg.result.items():
            if self.daten.klausur_punkte_setzen(sj, kl_name, hj, kidx, sk, punkte):
                saved += 1
            else:
                logging.warning(f"Punkte für {sk} konnten nicht gespeichert werden (Validierungsfehler)")
        if saved == 0 and dlg.result:
            messagebox.showwarning("Fehler", "Punkte konnten nicht gespeichert werden! Bitte Eingaben prüfen.", parent=self.root)
        self._save(); self._refresh_klausuren(); self._refresh_noten(self._kl(), self._sk())

    def _ns_csv_edit(self):
        sj, kl = self._sj(), self._kl()
        if not sj or not kl: messagebox.showwarning("Warnung", "Bitte Schuljahr und Klasse auswählen!"); return
        kl_name = self._parse_kl_name(kl)
        current_csv = self.daten.get_ns_csv(sj, kl_name)
        ns_typ = self.daten.get_notenschluessel(sj, kl_name)
        # Alle anderen Klassen sammeln
        alle_klassen = []
        for s, klasses in self.daten.schuljahre.items():
            for k in klasses:
                if not (s == sj and k == kl_name):
                    alle_klassen.append((s, k))
        dlg = NotenschluesselCsvDialog(self.root, current_csv, ns_typ, alle_klassen)
        if dlg.result is None: return
        csv_str, transfer = dlg.result
        self.daten.set_ns_csv(sj, kl_name, csv_str)
        for s, k in transfer:
            self.daten.set_ns_csv(s, k, csv_str)
        self._save(); self._refresh_klausuren(); self._refresh_noten(self._kl(), self._sk())


# ---------------------------------------------------------------------------
# Migration & Main
# ---------------------------------------------------------------------------
def _migrate_old_md():
    md_file = os.path.join(_APP_DIR, "noten.md")
    if not os.path.exists(md_file) or os.path.exists(DATA_FILE):
        return False
    old_data = NotenVerwaltung(); old_data.gewichtung_muendlich = DEFAULT_GEWICHTUNG
    cur_sj = cur_kl = cur_sk = cur_hj = None
    with open(md_file, "r", encoding="utf-8") as f:
        for z in f:
            z = z.strip()
            if z.startswith("Gewichtung Mündlich:"):
                try:
                    w = int(z.replace("Gewichtung Mündlich:", "").replace("%", "").strip())
                    if 0 <= w <= 100: old_data.gewichtung_muendlich = w
                except ValueError: pass
            elif z.startswith("## Schuljahr "):
                cur_sj = z[13:].strip(); old_data.schuljahre[cur_sj] = {}; cur_kl = cur_sk = cur_hj = None
            elif z.startswith("### Klasse ") and cur_sj:
                cur_kl = z[11:].strip()
                old_data.schuljahre[cur_sj][cur_kl] = {"notenschluessel": "IHK", "notenschluessel_csv": "", "schuelerinnen": {}, "klausuren": {}}
                cur_sk = cur_hj = None
            elif z.startswith("#### ") and cur_sj and cur_kl:
                t = z[5:].split(",", 1); nn = t[0].strip(); vn = t[1].strip() if len(t) > 1 else ""
                cur_sk = NotenVerwaltung._key(nn, vn)
                old_data.schuljahre[cur_sj][cur_kl]["schuelerinnen"][cur_sk] = {"nachname": nn, "vorname": vn, "halbjahre": old_data._hj_neu()}
                cur_hj = None
            elif z.startswith("##### ") and cur_sk:
                cur_hj = z[6:].strip()
                if cur_hj not in old_data.schuljahre[cur_sj][cur_kl]["schuelerinnen"][cur_sk]["halbjahre"]:
                    old_data.schuljahre[cur_sj][cur_kl]["schuelerinnen"][cur_sk]["halbjahre"][cur_hj] = {"muendlich": [], "schriftlich": []}
            elif z.startswith("- Mündlich:") and cur_sk and cur_hj:
                ns = z[12:].strip()
                if ns: old_data.schuljahre[cur_sj][cur_kl]["schuelerinnen"][cur_sk]["halbjahre"][cur_hj]["muendlich"] = [int(x.strip()) for x in ns.split(",") if x.strip()]
            elif z.startswith("- Schriftlich:") and cur_sk and cur_hj:
                ns = z[14:].strip()
                if ns: old_data.schuljahre[cur_sj][cur_kl]["schuelerinnen"][cur_sk]["halbjahre"][cur_hj]["schriftlich"] = [int(x.strip()) for x in ns.split(",") if x.strip()]
    return old_data


def main():
    root = tk.Tk()
    root.title("Notenverwaltung"); root.geometry("1100x680"); root.minsize(950, 580)
    migrated_data = _migrate_old_md()
    first_time = not os.path.exists(DATA_FILE)
    if first_time:
        dlg = PasswordDialog(root, title="Passwort setzen", first_time=True)
    else:
        dlg = PasswordDialog(root, title="Passwort eingeben", first_time=False)
    if dlg.result is None: root.destroy(); return
    password = dlg.result
    app = NotenVerwaltungApp(root, password)
    if migrated_data and migrated_data.schuljahre:
        migrated_data.speichern_verschluesselt(password, app.data_file)
        app.daten = migrated_data; app._refresh_sj()
    root.mainloop()


if __name__ == "__main__":
    main()