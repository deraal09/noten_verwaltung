"""
Hauptanwendung für Notenverwaltung
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import logging
from typing import Optional, Tuple

from constants import HALBJAHRE, AUTO_SAVE_MS, MSG_WAEHLEN_SJ_KL, MSG_WAEHLEN_SJ_KL_FACH, MSG_WAEHLEN_KL_FACH_SK, MSG_EXISTIERT_BEREITS
from constants import get_ns_aliases as _get_ns_aliases
from constants import DATA_FILE as DEFAULT_DATA_FILE
from models import NotenVerwaltung
from dialogs import (
    PasswordDialog, KlasseDialog, FachDialog, SchuelerinDialog,
    SchuelerlisteDialog, KlausurDialog, GewichtungDialog,
    PunkteDialog, _UebertragenDialog, NotenschluesselCsvDialog
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Hauptanwendung
# ---------------------------------------------------------------------------
class NotenVerwaltungApp:
    """GUI-Anwendung für die Notenverwaltung."""

    def __init__(self, root: tk.Tk, password: str, data_file: Optional[str] = None):
        self.root = root
        self.root.title("Notenverwaltung")
        self.root.geometry("960x600")
        self.root.minsize(710, 435)
        self.password = password
        self.data_file = data_file or DEFAULT_DATA_FILE
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
        sty.configure("TButton", font=("TkDefaultFont", 9), padding=(7, 4))
        sty.configure("TLabel", font=("TkDefaultFont", 9))
        sty.configure("TNotebook.Tab", font=("TkDefaultFont", 9, "bold"), padding=(14, 6))
        sty.configure("H.TLabel", font=("TkDefaultFont", 10, "bold"))
        sty.configure("G.TLabel", font=("TkDefaultFont", 11, "bold"))
        sty.configure("J.TLabel", font=("TkDefaultFont", 10, "bold"), foreground="#2a5da8")
        sty.configure("I.TLabel", font=("TkDefaultFont", 8), foreground="gray")
        sty.configure("NS.TLabel", font=("TkDefaultFont", 9, "bold"), foreground="#c44")
        sty.configure("TLabelframe.Label", font=("TkDefaultFont", 9, "bold"))
        sty.configure("TSpinbox", font=("TkDefaultFont", 9), padding=(3, 3))
        sty.configure("Treeview", font=("TkDefaultFont", 9), rowheight=22)
        sty.configure("Treeview.Heading", font=("TkDefaultFont", 8, "bold"))
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
        # Filter-Menü (umbenannt in "Schuljahr")
        self.filter_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="▾ Schuljahr", menu=self.filter_menu)
        self.filter_menu.add_command(label="Schuljahr: (bitte auswählen)", command=lambda: None, state="disabled")
        self._sj_menu_var = tk.StringVar()
        for sj in sorted(self.daten.schuljahre.keys()):
            self.filter_menu.add_radiobutton(label=f"  {sj}", variable=self._sj_menu_var,
                                              value=sj, command=self._on_sj_menu)
        self.filter_menu.add_command(label="+ Neues Schuljahr...", command=self._sj_add_menu)
        self.filter_menu.add_separator()
        self.filter_menu.add_command(label="Halbjahr:", command=lambda: None, state="disabled")
        self._hj_menu_var = tk.StringVar(value=HALBJAHRE[0])
        for hj in HALBJAHRE:
            self.filter_menu.add_radiobutton(label=f"  {hj}", variable=self._hj_menu_var,
                                              value=hj, command=self._on_hj_menu)
        # Hauptframe
        hf = ttk.Frame(self.root, padding=10)
        hf.pack(fill=tk.BOTH, expand=True)

        # SJ/HJ Variablen
        self.sj_var = tk.StringVar()
        self.hj_var = tk.StringVar()

        # SJ/HJ Anzeige (nur Text, Auswahl über Menü)
        self._sj_hj_label = ttk.Label(hf, text="",
                                       font=("TkDefaultFont", 10, "bold"))
        self._sj_hj_label.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))

        # Linke Spalte: Klasse und Fach über Schülerinnen
        left_col = ttk.Frame(hf)
        left_col.grid(row=1, column=0, sticky="nsew", padx=(0, 3))

        # Klasse und Fach Auswahl
        kf_frame = ttk.LabelFrame(left_col, text="Klasse & Fach", padding=5)
        kf_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(kf_frame, text="Klasse:").pack(anchor="w")
        self.kl_var = tk.StringVar()
        self.kl_cb = ttk.Combobox(kf_frame, textvariable=self.kl_var, state="readonly", width=12)
        self.kl_cb.pack(fill=tk.X, pady=(2, 2))
        self.kl_cb.bind("<<ComboboxSelected>>", self._on_kl)
        kl_btn_frame = ttk.Frame(kf_frame)
        kl_btn_frame.pack(fill=tk.X)
        ttk.Button(kl_btn_frame, text="+", width=5, command=self._kl_add).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(kl_btn_frame, text="→", width=5, command=self._kl_uebertragen).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(kl_btn_frame, text="−", width=5, command=self._kl_del).pack(side=tk.LEFT)

        ttk.Label(kf_frame, text="Fach:").pack(anchor="w", pady=(5, 0))
        self.fach_var = tk.StringVar()
        self.fach_cb = ttk.Combobox(kf_frame, textvariable=self.fach_var, state="readonly", width=14)
        self.fach_cb.pack(fill=tk.X, pady=(2, 2))
        self.fach_cb.bind("<<ComboboxSelected>>", self._on_fach)
        fach_btn_frame = ttk.Frame(kf_frame)
        fach_btn_frame.pack(fill=tk.X)
        ttk.Button(fach_btn_frame, text="+", width=5, command=self._fach_add).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(fach_btn_frame, text="−", width=5, command=self._fach_del).pack(side=tk.LEFT)

        # Schülerinnen
        sf = ttk.LabelFrame(left_col, text="Schülerinnen", padding=5)
        sf.pack(fill=tk.BOTH, expand=True)
        self.sk_lb = tk.Listbox(sf, height=10, width=18, exportselection=False,
                                 font=("TkDefaultFont", 9), selectbackground="#4a90d9")
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
        self._build_notenschluessel_tab()
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
        self.m_lb = tk.Listbox(self.m_frame, height=4, exportselection=False,
                                font=("TkDefaultFont", 9), selectbackground="#4a90d9")
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
        self.s_lb = tk.Listbox(self.s_frame, height=4, exportselection=False,
                                font=("TkDefaultFont", 9), selectbackground="#4a90d9")
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
        self.fp_lbl = ttk.Label(nf, text="", style="FP.TLabel")
        self.fp_lbl.pack(anchor="w", pady=(0, 0))
        self.j_lbl = ttk.Label(nf, text="Jahresnote: -", style="J.TLabel")
        self.j_lbl.pack(anchor="w", pady=(8, 0))
        ttk.Label(nf, text="(Gesamtnote über beide Halbjahre)", style="I.TLabel").pack(anchor="w")

    def _build_bewertung_tab(self, parent_frame: ttk.Frame, title: str) -> Tuple[tk.Listbox, ttk.Frame, ttk.Treeview]:
        """Generischer Aufbau für Klausuren/UL-Tabs. Gibt (listbox, btn_frame, tree) zurück."""
        ttk.Label(parent_frame, text=f"{title}:", style="H.TLabel").pack(anchor="w", pady=(0, 5))
        lb = tk.Listbox(parent_frame, height=4, exportselection=False,
                         font=("TkDefaultFont", 9), selectbackground="#4a90d9")
        lb.pack(fill=tk.X, pady=(0, 5))
        btn_frame = ttk.Frame(parent_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 5))
        tree_frame = ttk.Frame(parent_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(tree_frame, columns=("info",), show="headings", height=6)
        tree.heading("info", text=f"Keine {title} ausgewählt")
        tree.column("info", width=300)
        tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=tree_scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        return lb, btn_frame, tree

    def _build_notenschluessel_tab(self) -> None:
        """Tab für Anzeige und Bearbeitung des Notenschlüssels."""
        nf = ttk.Frame(self.nb, padding=10)
        self.nb.add(nf, text="  Notenschlüssel  ")

        # Info-Text
        info_frame = ttk.Frame(nf)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(info_frame, text="Notenschlüssel für Klasse:",
                  font=("TkDefaultFont", 10, "bold")).pack(anchor="w")
        self.ns_klasse_lbl = ttk.Label(info_frame, text="(Keine Klasse ausgewählt)",
                                       style="I.TLabel")
        self.ns_klasse_lbl.pack(anchor="w", pady=(2, 5))

        # Buttons zum Laden der Standards
        btn_frame = ttk.Frame(nf)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(btn_frame, text="IHK-Standard laden",
                   command=self._ns_load_ihk).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="BG-Standard laden",
                   command=self._ns_load_bg).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Bearbeiten",
                   command=self._ns_edit).pack(side=tk.LEFT)

        # Vorschau-Tabelle mit Info zu allen Fächern
        preview_frame = ttk.LabelFrame(nf, text="Vorschau (gilt für alle Fächer der Klasse)", padding=5)
        preview_frame.pack(fill=tk.BOTH, expand=True)

        # Treeview für die Tabelle
        columns = ("prozent", "note")
        self.ns_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=15)
        self.ns_tree.heading("prozent", text="Prozent ≥")
        self.ns_tree.heading("note", text="Note")
        self.ns_tree.column("prozent", width=120, anchor="center")
        self.ns_tree.column("note", width=100, anchor="center")

        ns_scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.ns_tree.yview)
        self.ns_tree.configure(yscrollcommand=ns_scroll_y.set)

        self.ns_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ns_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # Aktuelles CSV speichern für die Bearbeitung
        self._ns_current_csv = ""
        self._ns_current_kl = None

    def _ns_load_ihk(self) -> None:
        """Lädt den IHK-Standard-Notenschlüssel."""
        sj = self._sj()
        kl = self._kl()
        if not sj or not kl:
            messagebox.showwarning("Warnung", "Bitte Schuljahr und Klasse auswählen!")
            return
        kl_name = self._parse_kl_name(kl)
        from constants import DEFAULT_NS_CSV
        self.daten.set_ns_csv(sj, kl_name, DEFAULT_NS_CSV["IHK"])
        self._save()
        self._refresh_notenschluessel()

    def _ns_load_bg(self) -> None:
        """Lädt den BG-Standard-Notenschlüssel."""
        sj = self._sj()
        kl = self._kl()
        if not sj or not kl:
            messagebox.showwarning("Warnung", "Bitte Schuljahr und Klasse auswählen!")
            return
        kl_name = self._parse_kl_name(kl)
        from constants import DEFAULT_NS_CSV
        self.daten.set_ns_csv(sj, kl_name, DEFAULT_NS_CSV["BG"])
        self._save()
        self._refresh_notenschluessel()

    def _ns_edit(self) -> None:
        """Öffnet den Bearbeitungsdialog für den Notenschlüssel."""
        sj = self._sj()
        kl = self._kl()
        if not sj or not kl:
            messagebox.showwarning("Warnung", "Bitte Schuljahr und Klasse auswählen!")
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
        self._refresh_notenschluessel()
        self._refresh_klausuren()
        self._refresh_ul()
        self._refresh_noten(self._kl(), self._sk())


    def _build_klausuren_tab(self) -> None:
        kf = ttk.Frame(self.nb, padding=5)
        self.nb.add(kf, text="  Klausuren  ")
        self.kl_klausur_lb, btn_frame, self.kl_tree = self._build_bewertung_tab(kf, "Klausuren")
        self.kl_klausur_lb.bind("<<ListboxSelect>>", self._on_klausur_select)
        ttk.Button(btn_frame, text="Hinzufügen", command=self._klausur_add).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Löschen", command=self._klausur_del).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Punkte bearbeiten", command=self._klausur_punkte).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Gewichtung", command=self._klausur_gewichtung).pack(side=tk.LEFT, padx=(0, 5))

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
            initialdir=os.path.dirname(self.data_file) if self.data_file else os.path.dirname(os.path.abspath(__file__)))
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
            initialdir=os.path.dirname(self.data_file) if self.data_file else os.path.dirname(os.path.abspath(__file__)),
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
        # Aktualisiere Filter-Menü
        self.filter_menu.delete(1, tk.END)
        self.filter_menu.add_command(label="Schuljahr: (bitte auswählen)", command=lambda: None, state="disabled")
        self._sj_menu_var = tk.StringVar()
        for sj in sl:
            self.filter_menu.add_radiobutton(label=f"  {sj}", variable=self._sj_menu_var,
                                              value=sj, command=self._on_sj_menu)
        self.filter_menu.add_command(label="+ Neues Schuljahr...", command=self._sj_add_menu)
        self.filter_menu.add_separator()
        self.filter_menu.add_command(label="Halbjahr:", command=lambda: None, state="disabled")
        self._hj_menu_var = tk.StringVar(value=self.hj_var.get() or HALBJAHRE[0])
        for hj in HALBJAHRE:
            self.filter_menu.add_radiobutton(label=f"  {hj}", variable=self._hj_menu_var,
                                              value=hj, command=self._on_hj_menu)
        if sl:
            # Versuche das zuletzt gespeicherte Schuljahr zu verwenden
            letztes_sj = self.daten.letztes_schuljahr
            if letztes_sj and letztes_sj in sl:
                self.sj_var.set(letztes_sj)
                self._sj_menu_var.set(letztes_sj)
            elif self._sj() not in sl:
                self.sj_var.set(sl[0])
                self._sj_menu_var.set(sl[0])
            # Versuche das zuletzt gespeicherte Halbjahr zu verwenden
            letztes_hj = self.daten.letztes_halbjahr
            if letztes_hj and letztes_hj in HALBJAHRE:
                self.hj_var.set(letztes_hj)
                self._hj_menu_var.set(letztes_hj)
            self._update_sj_hj_label()
            self._refresh_kl()
        else:
            self.sj_var.set("")
            self.kl_var.set("")
            self.kl_cb['values'] = []
            self._sj_hj_label.config(text="")
            self._refresh_sk(None)
            self._refresh_noten(None, None)

    def _update_sj_hj_label(self) -> None:
        """Aktualisiert die SJ/HJ-Anzeige."""
        sj = self.sj_var.get()
        hj = self.hj_var.get()
        if sj and hj:
            self._sj_hj_label.config(text=f"Schuljahr: {sj} | Halbjahr: {hj}")
        elif sj:
            self._sj_hj_label.config(text=f"Schuljahr: {sj}")
        else:
            self._sj_hj_label.config(text="")

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

    def _refresh_notenschluessel(self) -> None:
        """Aktualisiert die Notenschlüssel-Vorschau."""
        # Treeview leeren
        for item in self.ns_tree.get_children():
            self.ns_tree.delete(item)

        sj = self._sj()
        kl = self._kl()

        if not sj or not kl:
            self.ns_klasse_lbl.config(text="(Keine Klasse ausgewählt)")
            return

        kl_name = self._parse_kl_name(kl)
        ns = self.daten.get_notenschluessel(sj, kl_name)
        csv_str = self.daten.get_ns_csv(sj, kl_name)
        entries = NotenVerwaltung.ns_csv_parse(csv_str)

        self.ns_klasse_lbl.config(text=f"{kl_name} ({sj}) — Notenschlüssel: {ns}")
        self._ns_current_csv = csv_str
        self._ns_current_kl = (sj, kl_name)

        # Einträge einfügen
        for p, n in entries:
            note_str = f"{n:.0f}" if n == int(n) else f"{n:.1f}"
            self.ns_tree.insert("", tk.END, values=(f"{p:.0f}%", note_str))

    def _refresh_notenschluessel_tab(self) -> None:
        """Aktualisiert die Notenschlüssel-Vorschau (Alias)."""
        self._refresh_notenschluessel()

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
            self.fp_lbl.config(text="")
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

        # Fehlende Punkte bis zur nächsten Note
        fp = self.daten.fehlende_punkte_bis_naechste_note(sj, kl_name, fach, sk, hj)
        if fp is not None:
            naechste_note, fehlende = fp
            note_str = f"{naechste_note:.0f}" if naechste_note == int(naechste_note) else f"{naechste_note:.1f}"
            self.fp_lbl.config(text=f"→ {fehlende} Pkt. bis Note {note_str}")
        else:
            self.fp_lbl.config(text="")

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
        self.daten.set_letztes_schuljahr(self._sj())
        self._refresh_kl()
        self._refresh_fach(None)
        self._refresh_sk(None)
        self._refresh_noten(None, None)
        self._refresh_notenschluessel()
        self._refresh_klausuren()
        self._refresh_ul()

    def _on_hj(self, e=None) -> None:
        self.daten.set_letztes_halbjahr(self._hj())
        self._refresh_all()

    def _on_sj_menu(self) -> None:
        sj = self._sj_menu_var.get()
        self.sj_var.set(sj)
        self.daten.set_letztes_schuljahr(sj)
        self._update_sj_hj_label()
        self._refresh_kl()
        self._refresh_fach(None)
        self._refresh_sk(None)
        self._refresh_noten(None, None)
        self._refresh_notenschluessel()
        self._refresh_klausuren()
        self._refresh_ul()

    def _on_hj_menu(self) -> None:
        hj = self._hj_menu_var.get()
        self.hj_var.set(hj)
        self.daten.set_letztes_halbjahr(hj)
        self._update_sj_hj_label()
        self._refresh_all()

    def _sj_add_menu(self) -> None:
        """Öffnet Dialog zum Hinzufügen eines neuen Schuljahres."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Neues Schuljahr")
        dlg.transient(self.root)
        dlg.grab_set()
        ttk.Label(dlg, text="Bezeichnung (z.B. 2024/25):").pack(padx=10, pady=(10, 3))
        entry = ttk.Entry(dlg, width=15)
        entry.pack(padx=10, pady=(0, 10))
        entry.focus()

        def save():
            name = entry.get().strip()
            if not name:
                messagebox.showwarning("Fehler", "Bitte eine Bezeichnung eingeben!", parent=dlg)
                return
            if name in self.daten.schuljahre:
                messagebox.showwarning("Fehler", "Schuljahr existiert bereits!", parent=dlg)
                return
            self.daten.schuljahre[name] = {}
            self.daten.letztes_schuljahr = name
            self._save()
            self._refresh_sj()
            dlg.destroy()

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(pady=(0, 10))
        ttk.Button(btn_frame, text="Speichern", command=save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Abbrechen", command=dlg.destroy).pack(side=tk.LEFT)
        dlg.bind("<Return>", lambda e: save())
        dlg.bind("<Escape>", lambda e: dlg.destroy())

    def _on_kl(self, e) -> None:
        self._refresh_fach(self._kl())
        self._refresh_sk(self._kl())
        self._refresh_noten(None, None)
        self._refresh_notenschluessel()
        self._refresh_klausuren()
        self._refresh_ul()

    def _on_fach(self, e) -> None:
        self._refresh_noten(self._kl(), self._sk())
        self._refresh_notenschluessel()
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

