"""
Dialog-Klassen für Notenverwaltung
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from typing import Optional, List, Tuple, Dict, Any

from constants import NOTENSCHLUESSEL, DEFAULT_NS_CSV
from constants import get_ns_aliases as _get_ns_aliases
from models import NotenVerwaltung


# ---------------------------------------------------------------------------
# Basis-Dialog
# ---------------------------------------------------------------------------
class _CenteredToplevel(tk.Toplevel):
    """Basis-Toplevel mit Zentrierungsfunktion."""

    def _center(self) -> None:
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw - self.winfo_width()) // 2}+{(sh - self.winfo_height()) // 2}")


# ---------------------------------------------------------------------------
# Passwort-Dialog
# ---------------------------------------------------------------------------
class PasswordDialog(_CenteredToplevel):
    """Dialog zur Passworteingabe (auch für Erstpasswort)."""

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


# ---------------------------------------------------------------------------
# Klasse-Dialog
# ---------------------------------------------------------------------------
class KlasseDialog(_CenteredToplevel):
    """Dialog zum Hinzufügen einer Klasse."""

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


# ---------------------------------------------------------------------------
# Fach-Dialog
# ---------------------------------------------------------------------------
class FachDialog(_CenteredToplevel):
    """Dialog zum Hinzufügen eines Fachs."""

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


# ---------------------------------------------------------------------------
# Schülerin-Dialog
# ---------------------------------------------------------------------------
class SchuelerinDialog(_CenteredToplevel):
    """Dialog zum Hinzufügen einer Schülerin."""

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


# ---------------------------------------------------------------------------
# Schülerliste-Dialog
# ---------------------------------------------------------------------------
class SchuelerlisteDialog(_CenteredToplevel):
    """Dialog zum Hinzufügen mehrerer Schülerinnen."""

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
        ttk.Label(f, text="Format: Nachname, Vorname (durch Komma, Leerzeichen oder Tab getrennt)",
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
            elif " " in line:
                parts = line.split(" ", 1)
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


# ---------------------------------------------------------------------------
# Klausur/UL-Dialog
# ---------------------------------------------------------------------------
class KlausurDialog(_CenteredToplevel):
    """Dialog zum Hinzufügen einer Klausur oder Unterrichtsleistung."""

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


# ---------------------------------------------------------------------------
# Gewichtung-Dialog
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# Punkte-Dialog
# ---------------------------------------------------------------------------
class PunkteDialog(_CenteredToplevel):
    """Dialog zum Bearbeiten von Punkten für Klausuren oder Unterrichtsleistungen."""

    def __init__(self, parent, title: str, name: str, max_punkte_pro_aufgabe: List[float],
                 schuelerinnen: List[Tuple[str, str, str]], existing_ergebnisse: Dict[str, List[float]],
                 ns_csv_str: str, notenschluessel_typ: str = "IHK"):
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
        headers = ["Schülerin"] + [f"A{i+1} (/{p})" for i, p in enumerate(max_punkte_pro_aufgabe)] + ["Gesamt", "%", "Note", "+ Pkt."]
        for c, h in enumerate(headers):
            ttk.Label(self.inner, text=h, font=("TkDefaultFont", 9, "bold")).grid(
                row=0, column=c, padx=2, pady=2, sticky="w")
        self.entries = {}
        self.labels_ges = {}
        self.labels_pct = {}
        self.labels_note = {}
        self.labels_bis_note = {}
        self.entry_grid = []  # Für Navigation mit Pfeiltasten
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
                e.bind("<Key>", self._handle_entry_key)
                row_entries.append(e)
            self.entries[sk] = row_entries
            self.entry_grid.append(row_entries)
            lbl_g = ttk.Label(self.inner, text="", width=7)
            lbl_g.grid(row=r, column=self.num_cols + 1, padx=2)
            lbl_p = ttk.Label(self.inner, text="", width=7)
            lbl_p.grid(row=r, column=self.num_cols + 2, padx=2)
            lbl_n = ttk.Label(self.inner, text="", width=7, font=("TkDefaultFont", 9, "bold"))
            lbl_n.grid(row=r, column=self.num_cols + 3, padx=2)
            lbl_bn = ttk.Label(self.inner, text="", width=8, foreground="#2a5da8")
            lbl_bn.grid(row=r, column=self.num_cols + 4, padx=2)
            self.labels_ges[r] = lbl_g
            self.labels_pct[r] = lbl_p
            self.labels_note[r] = lbl_n
            self.labels_bis_note[r] = lbl_bn
            self._update_row(r)
        # Abstand vor "Unterm Strich"
        ttk.Frame(self.inner, height=10).grid(row=len(schuelerinnen) + 1, column=0, pady=5)
        # "Unterm Strich" Zeile
        self.unter_strich_row = len(schuelerinnen) + 2
        ttk.Separator(self.inner, orient=tk.HORIZONTAL).grid(
            row=self.unter_strich_row, column=0, columnspan=self.num_cols + 5, sticky="ew", pady=(5, 5))
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
        # Leere Zelle für "bis Note" Spalte unterm Strich
        ttk.Label(self.inner, text="", width=8).grid(row=self.unter_strich_row + 1, column=self.num_cols + 4, padx=2)
        bf = ttk.Frame(self, padding=5)
        bf.pack(fill=tk.X)
        ttk.Button(bf, text="Alle berechnen", command=self._update_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="OK", command=self._ok, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bf, text="Abbrechen", command=self._cancel, width=10).pack(side=tk.RIGHT, padx=5)
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _get_row_points(self, r: int) -> List[float]:
        sk = self.schuelerinnen[r - 1][0]
        pts = []
        for e in self.entries[sk]:
            v = e.get().strip().replace(",", ".")  # Komma als Dezimaltrennzeichen akzeptieren
            if v == "":
                pts.append(None)
            else:
                try:
                    pts.append(float(v))
                except ValueError:
                    pts.append(None)
        return pts

    def _handle_entry_key(self, event) -> str:
        """Behandelt Navigation mit Pfeiltasten in den Eingabefeldern."""
        if not self.entry_grid:
            return

        # Finde aktuelles Entry
        current = event.widget
        current_row = None
        current_col = None
        for r, row in enumerate(self.entry_grid):
            for c, entry in enumerate(row):
                if entry == current:
                    current_row = r
                    current_col = c
                    break
            if current_row is not None:
                break

        if current_row is None or current_col is None:
            return

        num_rows = len(self.entry_grid)
        num_cols = len(self.entry_grid[0]) if num_rows > 0 else 0

        new_row, new_col = current_row, current_col

        if event.keysym == "Left":
            new_col = max(0, current_col - 1)
        elif event.keysym == "Right":
            new_col = min(num_cols - 1, current_col + 1)
        elif event.keysym == "Up":
            new_row = max(0, current_row - 1)
        elif event.keysym == "Down":
            new_row = min(num_rows - 1, current_row + 1)
        elif event.keysym == "Return" or event.keysym == "KP_Enter":
            # Enter: nächste Zelle oder erste Zelle der nächsten Zeile
            if current_col < num_cols - 1:
                new_col = current_col + 1
            elif current_row < num_rows - 1:
                new_row = current_row + 1
                new_col = 0
            else:
                # Letzte Zelle: OK-Button fokussieren
                for child in self.pack_slaves():
                    if isinstance(child, ttk.Frame):
                        for btn in child.pack_slaves():
                            if isinstance(btn, ttk.Button) and btn.cget("text") == "OK":
                                btn.focus()
                                return
        elif event.keysym == "Tab":
            # Tab: nächste Zelle
            if current_col < num_cols - 1:
                new_col = current_col + 1
            elif current_row < num_rows - 1:
                new_row = current_row + 1
                new_col = 0
            else:
                return "break"  # Tab normal weiterleiten
        else:
            return  # Andere Tasten normal weiterleiten

        if new_row != current_row or new_col != current_col:
            new_entry = self.entry_grid[new_row][new_col]
            new_entry.focus()
            new_entry.select_range(0, tk.END)
            return "break"

        return

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
            # Punkte bis zur nächsten Note berechnen (mit unrounded pct für Genauigkeit)
            self._update_bis_note(r, ges, max_p, pct_raw, note)
        else:
            self.labels_ges[r].config(text="")
            self.labels_pct[r].config(text="")
            self.labels_note[r].config(text="")
            self.labels_bis_note[r].config(text="")

    def _update_bis_note(self, r: int, ges: int, max_p: int, pct: float, current_note: Optional[float]) -> None:
        """Berechnet und zeigt, wie viele Punkte bis zur nächsten Note fehlen."""
        entries = NotenVerwaltung.ns_csv_parse(self.ns_csv)
        if not entries:
            self.labels_bis_note[r].config(text="")
            return

        if self.ns_typ == "BG":
            # BG: höhere Note ist besser, aber Noten können bei mehreren % gleich bleiben
            # Erstelle reduzierten Schlüssel nur mit Notensprüngen
            reduced = []
            last_note = None
            for p, n in sorted(entries, key=lambda x: x[0], reverse=True):
                if n != last_note:
                    reduced.append((p, n))
                    last_note = n
            reduced.sort(key=lambda x: x[0])  # Aufsteigend sortieren

            # Finde aktuellen Schwellenwert
            current_threshold_pct = None
            for p, n in reduced:
                if abs(n - current_note) < 0.01:
                    current_threshold_pct = p
                    break

            if current_threshold_pct is None:
                self.labels_bis_note[r].config(text="")
                return

            # Finde nächste bessere Note
            next_note_pct = None
            next_note_note = None
            for p, n in reduced:
                if n > current_note:
                    next_note_pct = p
                    next_note_note = n
                    break

            if next_note_pct is None:
                self.labels_bis_note[r].config(text="✓ beste")
                return

            # BG: höhere % = bessere Note, also pct_diff = next - current
            pct_diff = next_note_pct - current_threshold_pct
        else:
            # IHK: niedrigere Note ist besser, höhere % = bessere Note
            sorted_entries = sorted(entries, key=lambda x: x[0])

            # Finde den Schwellenwert der aktuellen Note
            current_threshold_pct = None
            for p, n in sorted_entries:
                if abs(n - current_note) < 0.01:
                    current_threshold_pct = p
                    break

            if current_threshold_pct is None:
                self.labels_bis_note[r].config(text="")
                return

            # IHK: bessere Note = kleinere Note (niedrigere Zahl)
            next_note_pct = None
            next_note_note = None

            for p, n in sorted_entries:
                if n < current_note:
                    next_note_pct = p
                    next_note_note = n
                    break

            if next_note_pct is None:
                self.labels_bis_note[r].config(text="✓ beste")
                return

            # IHK: pct_diff basierend auf tatsächlichem Prozent
            pct_diff = next_note_pct - pct

        # Berechne fehlende Punkte (auf 0.5 gerundet)
        if pct_diff <= 0:
            note_str = str(int(current_note)) if current_note == int(current_note) else f"{current_note:.1f}"
            self.labels_bis_note[r].config(text="✓ " + note_str)
        else:
            missing = round(pct_diff * max_p / 100 * 2) / 2  # Auf 0.5er runden
            missing_int = max(0.5, missing) if missing > 0 else 1
            next_note_str = str(int(next_note_note)) if next_note_note == int(next_note_note) else f"{next_note_note:.1f}"
            self.labels_bis_note[r].config(text=f"+{missing_int} → {next_note_str}")

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


# ---------------------------------------------------------------------------
# Übertragen-Dialog
# ---------------------------------------------------------------------------
class _UebertragenDialog(_CenteredToplevel):
    """Dialog zum Übertragen einer Klasse in ein anderes Schuljahr."""

    def __init__(self, parent, sj_quelle: str, kl_name: str, alle_sj: Dict[str, Any]):
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


# ---------------------------------------------------------------------------
# Notenschlüssel CSV Dialog
# ---------------------------------------------------------------------------
class NotenschluesselCsvDialog(_CenteredToplevel):
    """Dialog zum Bearbeiten des Notenschlüssel CSV."""

    def __init__(self, parent, current_csv: str, notenschluessel_typ: str, alle_klassen: List[Tuple[str, str]]):
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