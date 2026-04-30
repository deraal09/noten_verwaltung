"""
Datenmodell für Notenverwaltung
"""

import os
import csv
import logging
from typing import Optional, Dict, Any, List, Tuple

from constants import (
    HALBJAHRE, DEFAULT_GEWICHTUNG, DEFAULT_NS_CSV, NOTENSCHLUESSEL
)
from constants import get_ns_aliases as _get_ns_aliases
from encryption import encrypt_data, decrypt_data

logger = logging.getLogger(__name__)


class NotenVerwaltung:
    """Datenmodell für die Verwaltung von Noten."""

    def __init__(self) -> None:
        self.schuljahre: Dict[str, Dict[str, Dict[str, Any]]] = {}
        self.gewichtung_muendlich: int = DEFAULT_GEWICHTUNG
        self.letztes_schuljahr: Optional[str] = None
        self.letztes_halbjahr: Optional[str] = None

    # ---- Hilfsmethoden ----
    def _get_klasse(self, sj: str, k: str) -> Optional[Dict[str, Any]]:
        return self.schuljahre.get(sj, {}).get(k)

    def get_schueler_dict(self, sj: str, k: str) -> Optional[Dict[str, Dict[str, str]]]:
        kl = self._get_klasse(sj, k)
        return kl.get("schuelerinnen", {}) if kl else None

    def _get_fach(self, sj: str, k: str, fach: str) -> Optional[Dict[str, Any]]:
        kl = self._get_klasse(sj, k)
        if kl is None:
            return None
        return kl.get("faecher", {}).get(fach)

    def get_notenschluessel(self, sj: str, k: str) -> str:
        kl = self._get_klasse(sj, k)
        ns = kl.get("notenschluessel", "IHK") if kl else "IHK"
        ns_aliases = _get_ns_aliases()
        return ns_aliases.get(ns, ns)

    def set_notenschluessel(self, sj: str, k: str, ns_typ: str) -> None:
        """Setzt den Notenschlüssel-Typ für eine Klasse."""
        kl = self._get_klasse(sj, k)
        if kl:
            kl["notenschluessel"] = ns_typ

    def get_notenbereich(self, sj: str, k: str) -> Tuple[int, int]:
        ns = self.get_notenschluessel(sj, k)
        return NOTENSCHLUESSEL.get(ns, (1, 6))

    @property
    def ul_prozent(self) -> int:
        return self.gewichtung_muendlich

    @property
    def schriftlich_prozent(self) -> int:
        return 100 - self.gewichtung_muendlich

    def set_letztes_schuljahr(self, sj: Optional[str]) -> None:
        """Speichert das letzte ausgewählte Schuljahr."""
        self.letztes_schuljahr = sj

    def set_letztes_halbjahr(self, hj: Optional[str]) -> None:
        """Speichert das letzte ausgewählte Halbjahr."""
        self.letztes_halbjahr = hj

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
    def ns_csv_parse(csv_str: str) -> List[Tuple[float, float]]:
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
            ns_aliases = _get_ns_aliases()
            ns = ns_aliases.get(kl.get("notenschluessel", "IHK"), "IHK")
            return DEFAULT_NS_CSV.get(ns, DEFAULT_NS_CSV["IHK"])
        return csv_str

    def set_ns_csv(self, sj: str, k: str, csv_str: str) -> None:
        kl = self._get_klasse(sj, k)
        if kl is not None:
            kl["notenschluessel_csv"] = csv_str

    # ---- Serialisierung ----
    def to_dict(self) -> Dict[str, Any]:
        ns_aliases = _get_ns_aliases()
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
                    "notenschluessel": ns_aliases.get(kl_data.get("notenschluessel", "IHK"),
                                                      kl_data.get("notenschluessel", "IHK")),
                    "notenschluessel_csv": kl_data.get("notenschluessel_csv", ""),
                    "schuelerinnen": sk_dict,
                    "faecher": faecher,
                }
        return {"gewichtung_muendlich": self.gewichtung_muendlich,
                "letztes_schuljahr": self.letztes_schuljahr,
                "letztes_halbjahr": self.letztes_halbjahr,
                "schuljahre": sj}

    def from_dict(self, data: Dict[str, Any]) -> None:
        ns_aliases = _get_ns_aliases()
        self.gewichtung_muendlich = data.get("gewichtung_muendlich", DEFAULT_GEWICHTUNG)
        self.letztes_schuljahr = data.get("letztes_schuljahr")
        self.letztes_halbjahr = data.get("letztes_halbjahr")
        self.schuljahre = {}
        for s, klasses in data.get("schuljahre", {}).items():
            self.schuljahre[s] = {}
            for k, kl_data in klasses.items():
                if not isinstance(kl_data, dict):
                    continue
                schueler = {}
                ns = ns_aliases.get(kl_data.get("notenschluessel", "IHK"),
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
        # Auto-Verteilung für alte Daten ohne Gewichtung
        self._auto_distribute_all()

    def _parse_fach(self, fd: Dict[str, Any]) -> Dict[str, Any]:
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

    def _migrate_old_format(self, kl_data: Dict[str, Any], schueler: Dict[str, Dict[str, str]]) -> Dict[str, Any]:
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
        from constants import DATA_FILE
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
        from constants import DATA_FILE
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
        ns_aliases = _get_ns_aliases()
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

    def schuelerin_sortiert(self, sj: str, k: str) -> List[str]:
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

    def fach_sortiert(self, sj: str, k: str) -> List[str]:
        kl = self._get_klasse(sj, k)
        if kl is None:
            return []
        return sorted(kl.get("faecher", {}).keys())

    # ---- Noten (pro Fach) – öffentliche API ----
    def get_muendlich(self, sj: str, k: str, fach: str, sk: str, hj: str) -> List[float]:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return []
        return f.get("halbjahre", {}).get(hj, {}).get("noten", {}).get(sk, {}).get("muendlich", [])

    def get_schriftlich(self, sj: str, k: str, fach: str, sk: str, hj: str) -> List[float]:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return []
        return f.get("halbjahre", {}).get(hj, {}).get("noten", {}).get(sk, {}).get("schriftlich", [])

    def _ensure_noten_dict(self, sj: str, k: str, fach: str, sk: str, hj: str) -> Optional[Dict[str, List[float]]]:
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
                            name: str, max_punkte_pro_aufgabe: List[float]) -> bool:
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

    def get_klausuren(self, sj: str, k: str, fach: str, hj: str) -> List[Dict[str, Any]]:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return []
        return f.get("klausuren", {}).get(hj, [])

    def klausur_punkte_setzen(self, sj: str, k: str, fach: str, hj: str,
                               kidx: int, sk: str, punkte: List[float]) -> bool:
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
                                       kidx: int) -> Tuple[int, int, float, bool]:
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
                                     sk: str, hj: str) -> List[Tuple[float, float]]:
        klausuren = self.get_klausuren(sj, k, fach, hj)
        result = []
        for i in range(len(klausuren)):
            note = self.klausur_note_berechnen(sj, k, fach, hj, i, sk)
            if note is not None:
                result.append((note, klausuren[i].get("gewichtung", 0)))
        return result

    # ---- CRUD Unterrichtsleistungen (pro Fach) ----
    def ul_hinzufuegen(self, sj: str, k: str, fach: str, hj: str,
                       name: str, max_punkte_pro_aufgabe: List[float]) -> bool:
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

    def get_unterrichtsleistungen(self, sj: str, k: str, fach: str, hj: str) -> List[Dict[str, Any]]:
        f = self._get_fach(sj, k, fach)
        if f is None:
            return []
        return f.get("unterrichtsleistungen", {}).get(hj, [])

    def ul_punkte_setzen(self, sj: str, k: str, fach: str, hj: str,
                          ulidx: int, sk: str, punkte: List[float]) -> bool:
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
                                sk: str, hj: str) -> List[Tuple[float, float]]:
        uls = self.get_unterrichtsleistungen(sj, k, fach, hj)
        result = []
        for i in range(len(uls)):
            note = self.ul_note_berechnen(sj, k, fach, hj, i, sk)
            if note is not None:
                result.append((note, uls[i].get("gewichtung", 0)))
        return result

    # ---- Berechnungen ----
    @staticmethod
    def durchschnitt(noten: List[float]) -> Optional[float]:
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

    def fehlende_punkte_bis_naechste_note(self, sj: str, k: str, fach: str, sk: str, hj: str) -> Optional[Tuple[float, int]]:
        """Berechnet die fehlenden Punkte bis zur nächsten besseren Note.

        Returns:
            tuple: (naechste_note, fehlende_punkte) oder None wenn keine Berechnung möglich
        """
        gn = self.gesamtnote_hj(sj, k, fach, sk, hj)
        if gn is None:
            return None

        ns_typ = self.get_notenschluessel(sj, k)
        csv_str = self.get_ns_csv(sj, k)
        entries = self.ns_csv_parse(csv_str)
        if not entries:
            return None

        # Berechne tatsächlich erreichte Punkte und Maximalpunkte
        max_punkte = 0
        aktuelle_punkte = 0
        for ul in self.get_unterrichtsleistungen(sj, k, fach, hj):
            max_punkte += sum(ul.get("max_punkte_pro_aufgabe", []))
            ergebnisse = ul.get("ergebnisse", {}).get(sk, [])
            if ergebnisse:
                aktuelle_punkte += sum(p for p in ergebnisse if p is not None)
        for kl in self.get_klausuren(sj, k, fach, hj):
            max_punkte += sum(kl.get("max_punkte_pro_aufgabe", []))
            ergebnisse = kl.get("ergebnisse", {}).get(sk, [])
            if ergebnisse:
                aktuelle_punkte += sum(p for p in ergebnisse if p is not None)
        aktuelle_pct = aktuelle_punkte / max_punkte * 100 if max_punkte > 0 else 0

        if ns_typ == "BG":
            # BG: höhere Note ist besser, aber Noten können bei mehreren % gleich bleiben
            # Erstelle reduzierten Schlüssel nur mit Notensprüngen
            reduced = []
            last_note = None
            for p, n in sorted(entries, key=lambda x: x[0], reverse=True):
                if n != last_note:
                    reduced.append((p, n))
                    last_note = n
            reduced.sort(key=lambda x: x[0])  # Aufsteigend sortieren

            # Finde aktuellen Schwellenwert (wo die aktuelle Note beginnt)
            current_pct = None
            for p, n in reduced:
                if abs(n - gn) < 0.01:
                    current_pct = p
                    break
            if current_pct is None:
                return None

            # Finde nächste bessere Note
            naechste_note = None
            naechste_pct = None
            for p, n in reduced:
                if n > gn:
                    naechste_note = n
                    naechste_pct = p
                    break

            if naechste_note is None:
                return None  # Bereits beste Note

            # BG: fehlende Punkte = nächster Schwellenwert - tatsächlich erreichte Punkte
            pct_diff = naechste_pct - aktuelle_pct
        else:
            # IHK: niedrigere Note ist besser, höhere % = bessere Note
            sorted_entries = sorted(entries, key=lambda x: x[0])

            # Finde die nächste bessere Note
            naechste_note = None
            naechste_pct = None
            for p, n in sorted_entries:
                if n < gn:
                    naechste_note = n
                    naechste_pct = p
                    break

            if naechste_note is None:
                return None  # Bereits beste Note

            # IHK: fehlende Punkte = nächster Schwellenwert - tatsächlich erreichte Punkte
            pct_diff = naechste_pct - aktuelle_pct

        if pct_diff <= 0:
            return None

        # Fehlende Punkte basierend auf der tatsächlichen Gesamtpunktzahl
        max_punkte = 0
        for ul in self.get_unterrichtsleistungen(sj, k, fach, hj):
            max_punkte += sum(ul.get("max_punkte_pro_aufgabe", []))
        for kl in self.get_klausuren(sj, k, fach, hj):
            max_punkte += sum(kl.get("max_punkte_pro_aufgabe", []))

        if max_punkte > 0:
            fehlende = int(pct_diff * max_punkte / 100 + 0.5)
        else:
            fehlende = int(pct_diff + 0.5)

        # Mindestens 1 Punkt anzeigen wenn eine bessere Note möglich ist
        if fehlende == 0:
            fehlende = 1

        return (naechste_note, fehlende)
