#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""

"""

from __future__ import annotations
import json
import logging
import threading
import os, sys, subprocess
import tkinter.simpledialog as simpledialog
import time
from pathlib import Path
from datetime import datetime, time as dt_time, timedelta
from typing import Optional, Dict
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

# watchdog opzionale
try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
except Exception:
    Observer = None  # type: ignore

# --- Backend hooks ---
from Swarky import Config, run_once, setup_logging, count_tif_files

# --- Tema ---
LIGHT_BG = "#eef3f9"
NAVY_BG  = "#000080"
NAVY_SEL = "#133869"
FG_LIGHT = "#ffffff"
FG_DARK  = "#1f2937"


def _open_path(path: Path) -> None:
    try:
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


class SwarkyApp:
    def __init__(self) -> None:
        self.json_path = Path("config.json")

        # Root
        self.root = tk.Tk()
        try:
            self.root.iconbitmap(default="esse.ico")
        except Exception:
            pass
        self.root.title("Swarky")
        self.root.configure(bg=LIGHT_BG)
        
        # Icona finestra anche in exe (PyInstaller onefile)
        def resource_path(rel: str) -> str:
            import sys, os
            from pathlib import Path
            base = getattr(sys, "_MEIPASS", os.getcwd())  # cartella temporanea del bundle
            return str(Path(base) / rel)

        try:
            self.root.iconbitmap(default=resource_path("esse.ico"))
        except Exception as e:
            logging.debug("iconbitmap fallita: %s", e)

        self._run_error_notified = False
        self._run_in_progress = False
        self._run_lock = threading.Lock()

        # Flag: blocca gli scan di Plotter durante il batch backend
        self._scan_plotter_disabled = False
        # Debounce id per refresh Plotter
        self._refresh_plotter_after_id: Optional[str] = None
        self._phase_tick_id: Optional[str] = None

        # Config boot
        self._ensure_default_config()
        self.cfg = self._build_cfg_from_json(self._load_config_json(silent=True))
        setup_logging(self.cfg)

        # Tema + UI
        self._setup_theme()
        self._build_layout()

        # Watchers
        self.plotter_observer: Optional[Observer] = None
        self.watch_thread: Optional[threading.Thread] = None
        self.watch_stop_event: Optional[threading.Event] = None

        # Bootstrap
        self.refresh_plotter()
        self._schedule_if_ready()
        self.update_clock()
        self.periodic_plotter_refresh()
        self.start_plotter_watcher()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
    # ---------------- Tabellari ----------------
    def open_tabellari(self) -> None:
        """
        1) Chiede 'Numero Disegno' (es: DXX123456)
        2) Chiede 'Quante posizioni vuoi inserire?' (es: 7)
        3) Genera un TXT in DIR_TABELLARI con una singola riga:
           DXX12345601,DXX12345602, ...
        """
        try:
            # --- Input 1: prefisso ---
            prefix = simpledialog.askstring(
                "Tabellari", "Inserisci Numero Disegno:", parent=self.root
            )
            if prefix is None:  # annulla
                return
            prefix = prefix.strip().upper()

            # Validazione semplice: D + 2 lettere + 6 cifre
            import re
            if not re.fullmatch(r"D[A-Z]{2}\d{6}", prefix):
                messagebox.showerror("Tabellari", "Formato non valido.\nAtteso: D + due lettere + 6 cifre.")
                return

            # --- Input 2: numero posizioni ---
            n_str = simpledialog.askstring(
                "Tabellari", "Quante posizioni vuoi inserire?", parent=self.root
            )
            if n_str is None:
                return
            n_str = n_str.strip()
            if not n_str.isdigit():
                messagebox.showerror("Tabellari", "Inserisci un numero intero valido.")
                return

            n = int(n_str)
            if n <= 0 or n > 99:
                messagebox.showerror("Tabellari", "Il numero deve essere tra 1 e 99.")
                return

            # --- Generazione stringhe: 01..NN a 2 cifre ---
            items = [f"{prefix}{i:02d}" for i in range(1, n + 1)]
            line = ",".join(items) + "\n"

            # --- Scrittura file ---
            out_dir = self.cfg.DIR_TABELLARI
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / f"{prefix}_tabellari.txt"

            with out_path.open("w", encoding="utf-8") as f:
                f.write(line)

            messagebox.showinfo("Tabellari", f"File creato:\n{out_path}")
            # opzionale: apri la cartella
            # _open_path(out_dir)

        except Exception as e:
            logging.exception("Errore Tabellari")
            messagebox.showerror("Tabellari", f"Errore imprevisto:\n{e}")
                
    # ---------------- CONFIG (solo JSON) ----------------
    def _ensure_default_config(self) -> None:
        if self.json_path.exists():
            return
        default = {
            "paths": {
                "hplotter": ".",
                "archivio": "./ArchivioDisegni",
                "error_dir": "./Rivedere",
                "pari_rev": "./Pari_Revisione",
                "plm": "./plm",
                "storico": "./Storico",
                "iss": "./iss",
                "fiv": "./FIVloading",
                "heng": "./Hengelo",
                "error_plm": "./Plm_error",
                "tab": "./tabellari",
                "log_dir": None
            },
            "AUTO_TIME": "",
            "LOG_LEVEL": "INFO",
            "ACCEPT_PDF": True,
            "LOG_PHASES": True
        }
        try:
            self.json_path.write_text(json.dumps(default, indent=2), encoding="utf-8")
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile creare config.json: {e}")

    def _load_config_json(self, silent: bool = False) -> dict:
        try:
            return json.loads(self.json_path.read_text(encoding="utf-8"))
        except Exception as e:
            if not silent:
                messagebox.showwarning("Config JSON non leggibile", f"Dettagli: {e}")
            return {}

    def _build_cfg_from_json(self, data: dict) -> Config:
        paths = data.get("paths", {}) if isinstance(data, dict) else {}

        def _p(x):
            return Path(x).expanduser() if x else None

        # passa anche ACCEPT_PDF al dataclass Config del backend
        return Config(
            DIR_HPLOTTER      = _p(paths.get("hplotter")),
            ARCHIVIO_DISEGNI  = _p(paths.get("archivio")),
            ERROR_DIR         = _p(paths.get("error_dir")),
            PARI_REV_DIR      = _p(paths.get("pari_rev")),
            PLM_DIR           = _p(paths.get("plm")),
            ARCHIVIO_STORICO  = _p(paths.get("storico")),
            DIR_ISS           = _p(paths.get("iss")),
            DIR_FIV_LOADING   = _p(paths.get("fiv")),
            DIR_HENGELO       = _p(paths.get("heng")),
            DIR_PLM_ERROR     = _p(paths.get("error_plm")),
            DIR_TABELLARI     = _p(paths.get("tab")),
            LOG_DIR           = _p(paths.get("log_dir")),
            LOG_LEVEL         = logging.INFO if data.get("LOG_LEVEL","INFO")=="INFO" else logging.DEBUG,
            ACCEPT_PDF        = bool(data.get("ACCEPT_PDF", True)),
            LOG_PHASES        = bool(data.get("LOG_PHASES", True)),
        )

    def _reload_cfg(self) -> None:
        self.cfg = self._build_cfg_from_json(self._load_config_json(silent=True))
        try:
            setup_logging(self.cfg)
        except Exception:
            pass

    # ---------------- TEMA / UI ----------------
    def _setup_theme(self) -> None:
        for fname in ("TkDefaultFont","TkTextFont","TkMenuFont","TkHeadingFont"):
            try:
                tkfont.nametofont(fname).configure(family="Consolas")
            except tk.TclError:
                pass
        style = ttk.Style(self.root)
        try: style.theme_use("clam")
        except tk.TclError: pass
        style.configure(".", font=("Consolas", 9), background=LIGHT_BG, foreground=FG_DARK)
        style.configure("TFrame", background=LIGHT_BG)
        style.configure("TLabelframe", background=LIGHT_BG, bordercolor=NAVY_BG)
        style.configure("TLabelframe.Label", background=LIGHT_BG, foreground=FG_DARK)
        style.configure("Navy.Treeview", background=NAVY_BG, fieldbackground=NAVY_BG,
                        foreground=FG_LIGHT, rowheight=24, borderwidth=0)
        style.map("Navy.Treeview",
                  background=[("selected", NAVY_SEL)],
                  foreground=[("selected", FG_LIGHT)])
        style.configure("Navy.Treeview.Heading", background="white", foreground="black", relief="flat")

    def _build_layout(self) -> None:
        for c in range(3):
            self.root.columnconfigure(c, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.lbl_plm_errors_var = tk.StringVar(value="N° PLM errors: 0")
        self.lbl_check_var      = tk.StringVar(value="N° Check Dwgs: 0")
        self.lbl_same_var       = tk.StringVar(value="N° Same Rev.: 0")
        self.lbl_drawings_var   = tk.StringVar(value="N° Drawings: 0")
        self.plotter_lbl        = tk.StringVar(value="plotter")
        self.plotter_frame   = ttk.LabelFrame(self.root, text="Plotter")
        self.anomaly_frame   = ttk.LabelFrame(self.root, text="Anomalie")
        self.processed_frame = ttk.LabelFrame(self.root, text="File processati")
        ttk.Label(self.plotter_frame, textvariable=self.lbl_plm_errors_var, anchor="w").pack(side="bottom", fill="x")
        ttk.Label(self.plotter_frame, textvariable=self.lbl_check_var,      anchor="w").pack(side="bottom", fill="x")
        ttk.Label(self.plotter_frame, textvariable=self.lbl_same_var,       anchor="w").pack(side="bottom", fill="x")
        ttk.Label(self.plotter_frame, textvariable=self.lbl_drawings_var,   anchor="w").pack(side="bottom", fill="x")
        self.controls        = ttk.Frame(self.root)
        self.plotter_frame.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8,8))
        self.anomaly_frame.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        self.processed_frame.grid(row=0, column=2, sticky="nsew", padx=8, pady=8)
        self.controls.grid(row=1, column=0, columnspan=3, sticky="ew", padx=8, pady=(0,8))

        # Plotter list
        self.plotter_list = tk.Listbox(self.plotter_frame, highlightthickness=0, borderwidth=0,
                                       bg="navy", fg="white", selectbackground="#1d4ed8",
                                       selectforeground="white",
                                       font=("Consolas",9)
                                       )
        self.plotter_list.pack(fill="both", expand=True)
        self.plotter_list.bind("<Double-Button-1>", self._open_selected_plotter)
        
        # Anomalie
        self.anomaly_tree = ttk.Treeview(
            self.anomaly_frame, 
            columns=("data","ora","file","errore"),
            show="headings",
            style="Navy.Treeview"
        )
        for col, head, w in (
            ("data","Data",90),
            ("ora","Ora",70),
            ("file","File",160),
            ("errore","Errore",180)
        ):
            self.anomaly_tree.heading(col, text=head, anchor="w")
            self.anomaly_tree.column(col, width=w, anchor="w", stretch=True)
        self.anomaly_tree.pack(fill="both", expand=True)

        # Processati
        self.processed_tree = ttk.Treeview(
            self.processed_frame,
            columns=("data","ora","file","proc","dest","conf"),
            show="headings",
            style="Navy.Treeview"
        )
        for col, head, w in (
            ("data","Data",90),
            ("ora","Ora",70),
            ("file","File",160),
            ("proc","Processo",160),
            ("dest","Destinazione",150),
            ("conf","Confronto",180)
        ):
            self.processed_tree.heading(col, text=head, anchor="w")
            self.processed_tree.column(col, width=w, anchor="w", stretch=True)
        self.processed_tree.pack(fill="both", expand=True)

        # Min sizes
        self.root.update_idletasks()
        min_anom = sum(self.anomaly_tree.column(c,'width') for c in ("data","ora","file","errore")) + 16
        self.root.grid_columnconfigure(1, minsize=min_anom, weight=1)
        min_proc = sum(self.processed_tree.column(c,'width') for c in ("data","ora","file","proc","dest","conf")) + 16
        self.root.grid_columnconfigure(2, minsize=min_proc, weight=1)
        plotter_min = tkfont.nametofont("TkDefaultFont").measure("M"*25)
        self.root.grid_columnconfigure(0, minsize=plotter_min, weight=1)
        self.root.minsize(plotter_min + min_anom + min_proc + 48, 480)

        # Controls
        self.interval_var = tk.StringVar(value="60")
        ttk.Label(self.controls, text="Intervallo (s):").pack(side="left")
        ttk.Entry(self.controls, textvariable=self.interval_var, width=6).pack(side="left", padx=(0,8))
        self.btn_swarky = ttk.Button(self.controls, text="Swarky", command=self.run_once_thread)
        self.btn_swarky.pack(side="left", padx=(20,4))        
        self.btn_start = ttk.Button(self.controls, text="Avvia watch", command=self.start_watch)
        self.btn_stop  = ttk.Button(self.controls, text="Ferma watch", state="disabled", command=self.stop_watch)
        self.btn_start.pack(side="left", padx=4)
        self.btn_stop.pack(side="left", padx=4)
        ttk.Button(self.controls, text="Pulisci", command=self._clear_tables).pack(side="left", padx=4)
        ttk.Button(self.controls, text="Plotter", command=self._open_plotter_folder).pack(side="left", padx=4)
        ttk.Button(self.controls, text="PariRev", command=self.open_parirev).pack(side="left", padx=4)
        ttk.Button(self.controls, text="Tabellari", command=self.open_tabellari).pack(side="left", padx=4)
        ttk.Button(self.controls, text="Settings", command=self.open_settings).pack(side="left", padx=4)

        # Orologio (larghezza fissa: "dd/mm/yyyy hh:mm:ss" = 19 char)
        self.clock_label = ttk.Label(self.controls, width=19, anchor="e")
        self.clock_label.pack(side="right", padx=5)

        # Label stato fasi (fase + ms) con larghezza fissa
        self.phase_var = tk.StringVar(value="Pronto.")
        self.phase_label = ttk.Label(self.controls, textvariable=self.phase_var, width=36, anchor="e")
        self.phase_label.pack(side="right", padx=8)

        # Logging → GUI
        self.tree_handler = _TreeviewHandler(self)
        logging.getLogger().handlers.insert(0, self.tree_handler)

    # ---------------- Inserimento righe ----------------
    def insert_processed(self, data, ora, file, proc, dest, conf) -> None:
        self.processed_tree.insert("", "end", values=(data, ora, file, proc, dest, conf))

    def insert_anomaly(self, data, ora, file, errore) -> None:
        self.anomaly_tree.insert("", "end", values=(data, ora, file, errore))
        
    # ---------------- Gestione contatori ----------------        
    def update_counters(self) -> None:
        try:
            stats = count_tif_files(self.cfg)
        except Exception:
            stats = {}
        drawings = 0
        try:
            drawings = self.plotter_list.size()
        except Exception:
            pass
        self.lbl_drawings_var.set(f"N° Drawings: {drawings}")
        self.lbl_same_var.set(f"N° Same Rev.: {stats.get('Same Rev Dwg', 0)}")
        self.lbl_check_var.set(f"N° Check Dwgs: {stats.get('Check Dwg', 0)}")
        self.lbl_plm_errors_var.set(f"N° PLM errors: {stats.get('Plm error Dwg', 0)}")
        
    # ---------------- Plotter ----------------
    def refresh_plotter(self) -> None:
        """Full scan della cartella Plotter (usare con parsimonia)."""
        self.plotter_list.delete(0, tk.END)
        if getattr(self.cfg, "ACCEPT_PDF", True):
            patterns = ("*.tif","*.TIF","*.pdf","*.PDF")
        else:
            patterns = ("*.tif","*.TIF")
        try:
            base = self.cfg.DIR_HPLOTTER
            files = {p.name.lower(): p.name for pat in patterns for p in base.glob(pat) if p.is_file()}
        except Exception:
            files = {}
        for name in sorted(files.values(), key=str.lower):
            self.plotter_list.insert(tk.END, name)
        self.update_counters()

    def request_plotter_refresh(self, delay_ms: int = 300) -> None:
        """Debounce: pianifica un refresh_plotter unico entro delay_ms."""
        if self._scan_plotter_disabled:
            return
        # coalesca più richieste ravvicinate
        if self._refresh_plotter_after_id is not None:
            try:
                self.root.after_cancel(self._refresh_plotter_after_id)
            except Exception:
                pass
            self._refresh_plotter_after_id = None
        self._refresh_plotter_after_id = self.root.after(delay_ms, self._do_debounced_refresh)

    def _do_debounced_refresh(self) -> None:
        self._refresh_plotter_after_id = None
        if not self._scan_plotter_disabled:
            self.refresh_plotter()

    def _open_selected_plotter(self, _=None) -> None:
        sel = self.plotter_list.curselection()
        if not sel:
            return
        p = (self.cfg.DIR_HPLOTTER / self.plotter_list.get(sel[0]))
        if p.exists():
            _open_path(p)
            
    def _refresh_parirev(self) -> None:
        """Se la finestra PariRev è aperta, aggiorna la sua listbox."""
        try:
            win = getattr(self, "_parirev_win", None)
            if win and win.winfo_exists():
                win.refresh_list()
        except Exception:
            pass

    # ---------------- AUTO_TIME ----------------
    def _read_auto_time_from_file(self) -> Optional[dt_time]:
        try:
            at = self._load_config_json(silent=True).get("AUTO_TIME", "")
            if at and ":" in at:
                hh, mm = map(int, at.split(":", 1))
                if 0 <= hh <= 23 and 0 <= mm <= 59:
                    return dt_time(hh, mm)
        except Exception:
            pass
        return None

    # ---------------- Timer & clock ----------------
    def periodic_plotter_refresh(self) -> None:
        """Aggiorna periodicamente, rispettando il blocco scan e usando debounce."""
        if not self._scan_plotter_disabled:
            self.request_plotter_refresh(delay_ms=300)
        self._refresh_parirev()
        self.root.after(10000 if self.plotter_observer else 1000, self.periodic_plotter_refresh)

    def update_clock(self) -> None:
        self.clock_label.config(text=datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        self.root.after(1000, self.update_clock)

    # ---------------- Run & scheduler ----------------
    def run_once_thread(self) -> None:
        if self._run_in_progress:
            return
        self._run_in_progress = True
        try:
            self.btn_swarky.state(["disabled"])
        except Exception:
            pass
        threading.Thread(target=self._run_once_worker, daemon=True).start()

    def _run_once_worker(self) -> None:
        try:
            if not self._run_lock.acquire(blocking=False):
                self.root.after(0, lambda: messagebox.showinfo("Attendi", "Un'elaborazione è già in corso."))
                return

            # Blocca gli scan di Plotter durante il batch
            self._scan_plotter_disabled = True
            # Cancella un eventuale debounce in coda
            if self._refresh_plotter_after_id is not None:
                try:
                    self.root.after_cancel(self._refresh_plotter_after_id)
                except Exception:
                    pass
                self._refresh_plotter_after_id = None

            result = run_once(self.cfg)
            self.root.after(0, lambda: messagebox.showinfo(
                "Swarky",
                "Esecuzione completata." + ("" if result else "\nNessun file da processare.")
            ))

        except Exception as e:
            if not self._run_error_notified:
                self._run_error_notified = True
                err = str(e)
                self.root.after(0, lambda msg=err: messagebox.showwarning(
                    "Esecuzione interrotta",
                    "Errore durante run_once (config incompleta o percorsi non validi?).\n\n"
                    f"Dettagli: {msg}\nControlla i percorsi in Settings."
                ))

        finally:
            if self._run_lock.locked():
                try:
                    self._run_lock.release()
                except Exception:
                    pass
            self._run_in_progress = False

            # Riabilita scan e fai UN solo refresh finale
            def _after_batch():
                self._scan_plotter_disabled = False
                self.refresh_plotter()
                self._refresh_parirev()
                self._phase_end("Pronto.")
                try:
                    self.btn_swarky.state(["!disabled"])
                except Exception:
                    pass
            self.root.after(0, _after_batch)

    def _schedule_if_ready(self) -> None:
        if hasattr(self, "_schedule_id") and self._schedule_id is not None:
            try: self.root.after_cancel(self._schedule_id)
            except Exception: pass
            self._schedule_id = None

        target_time = self._read_auto_time_from_file()
        if not target_time:
            return
        now = datetime.now()
        target = datetime.combine(now.date(), target_time)
        if now >= target:
            target += timedelta(days=1)
        delay_ms = int((target - now).total_seconds() * 1000)
        self._schedule_id = self.root.after(delay_ms, self._scheduled_run)

    def _scheduled_run(self) -> None:
        self.run_once_thread()
        self._schedule_if_ready()

    # ---------------- Watchdog FS ----------------
    def start_plotter_watcher(self) -> None:
        if Observer is None:
            return
        app = self
        class Handler(FileSystemEventHandler):
            def _refresh(self, event):
                if getattr(event, "is_directory", False):
                    return
                # Rispetta il blocco scan; usa debounce
                if not app._scan_plotter_disabled:
                    app.root.after(0, app.request_plotter_refresh)
                    app.root.after(0, app._refresh_parirev)
            on_created = on_moved = on_deleted = _refresh
        try:
            observer = Observer()
            observer.schedule(Handler(), str(self.cfg.DIR_HPLOTTER), recursive=False)
            observer.daemon = True
            observer.start()
            self.plotter_observer = observer
        except Exception:
            self.plotter_observer = None

    # ---------------- Periodic watch ----------------
    def _watch_worker(self, interval: int, stop_event: threading.Event) -> None:
        while not stop_event.is_set():
            try:
                if not self._run_lock.acquire(blocking=False):
                    pass
                else:
                    try:
                        # Blocca gli scan nel thread di watch per la durata del batch
                        self._scan_plotter_disabled = True
                        if self._refresh_plotter_after_id is not None:
                            try:
                                self.root.after_cancel(self._refresh_plotter_after_id)
                            except Exception:
                                pass
                            self._refresh_plotter_after_id = None

                        result = run_once(self.cfg)
                        msg = "Completato." if result else "Nessun file."
                        logging.info("Watch: %s", msg)
                        self.root.after(0, lambda m=msg: self.clock_label.config(text=f"{datetime.now():%H:%M:%S} • {m}"))
                    finally:
                        try:
                            self._run_lock.release()
                        except Exception:
                            pass
                        # Riabilita scan e un refresh finale
                        def _after_watch_batch():
                            self._scan_plotter_disabled = False
                            self.refresh_plotter()
                            self._refresh_parirev()
                            self._phase_end("Pronto.")
                        self.root.after(0, _after_watch_batch)
            except Exception as e:
                if not self._run_error_notified:
                    self._run_error_notified = True
                    err = str(e)
                    self.root.after(0, lambda msg=err: messagebox.showwarning(
                        "Watch: errore",
                        "Errore durante l'esecuzione periodica (config incompleta o percorsi non validi?).\n\n"
                        f"Dettagli: {msg}\nIl watch continuerà a provare."
                    ))

            # attesa intervallo, interrotta se arriva stop_event
            for _ in range(interval):
                if stop_event.is_set():
                    break
                threading.Event().wait(1)

    def start_watch(self) -> None:
        if self.watch_thread and self.watch_thread.is_alive():
            return
        try:
            interval = max(1, int(self.interval_var.get()))
        except ValueError:
            interval = 60
        self.watch_stop_event = threading.Event()
        self.watch_thread = threading.Thread(target=self._watch_worker,
                                             args=(interval, self.watch_stop_event), daemon=True)
        self.watch_thread.start()
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")

    def stop_watch(self) -> None:
        if self.watch_stop_event:
            self.watch_stop_event.set()
        if self.watch_thread:
            self.watch_thread.join(timeout=2)
        self.watch_thread = None
        self.watch_stop_event = None
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")

    # ---------------- Utility controls ----------------
    def _clear_tables(self) -> None:
        for tv in (self.anomaly_tree, self.processed_tree):
            for i in tv.get_children():
                tv.delete(i)

    def _open_plotter_folder(self) -> None:
        _open_path(self.cfg.DIR_HPLOTTER)

    def open_settings(self) -> None:
        SettingsDialog(self)
        
    def open_parirev(self) -> None:
        try:
            if getattr(self, "_parirev_win", None) and self._parirev_win.winfo_exists():
                self._parirev_win.focus_set()
                return
        except Exception:
            pass
        try:
            from Gui_Parirev import PariRevWindow
            self._parirev_win = PariRevWindow(self.root, self.cfg)
        except Exception as e:
            messagebox.showerror("PariRev", f"Impossibile aprire l'interfaccia PariRev:\n{e}")
 
 # ---------------- Phase timer ----------------
    def _phase_start(self, text: str) -> None:
        """Avvia/Resetta timer per la fase corrente e mostra testo+ms."""
        # stop di un possibile ticker precedente
        try:
            if hasattr(self, "_phase_tick_id") and self._phase_tick_id:
                self.root.after_cancel(self._phase_tick_id)
                self._phase_tick_id = None
        except Exception:
            pass

        self._phase_text = text
        self._phase_t0 = time.perf_counter()
        self.phase_var.set(f"{text} • 0 ms")
        self._phase_tick_id = self.root.after(50, self._phase_tick)

    def _phase_tick(self) -> None:
        try:
            if not hasattr(self, "_phase_t0"):
                return
            ms = int((time.perf_counter() - self._phase_t0) * 1000)
            self.phase_var.set(f"{self._phase_text} • {ms} ms")
            self._phase_tick_id = self.root.after(50, self._phase_tick)
        except Exception:
            pass

    def _phase_end(self, final_text: str | None = None) -> None:
        """Ferma il timer fase e mostra eventualmente un testo finale."""
        try:
            if hasattr(self, "_phase_tick_id") and self._phase_tick_id:
                self.root.after_cancel(self._phase_tick_id)
                self._phase_tick_id = None
        except Exception:
            pass
        self.phase_var.set(final_text or "Pronto.")
        for attr in ("_phase_t0", "_phase_text"):
            if hasattr(self, attr):
                delattr(self, attr)
            

    # ---------------- Close ----------------
    def _on_close(self) -> None:
        # spegne eventuale ticker/label di fase attivo
        try:
            self._phase_end()
        except Exception:
            pass

        try:
            if self.plotter_observer:
                self.plotter_observer.stop()
                self.plotter_observer.join(timeout=2)
        except Exception:
            pass
        self.stop_watch()
        if hasattr(self, "_schedule_id") and self._schedule_id is not None:
            try: self.root.after_cancel(self._schedule_id)
            except Exception: pass
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


class SettingsDialog(tk.Toplevel):
    """Impostazioni: paths + AUTO_TIME (HH:MM) + ACCEPT_PDF."""
    PATH_FIELDS = [
        ("hplotter",  "Cartella Plotter"),
        ("archivio",  "Archivio Disegni"),
        ("error_dir", "Rivedere (error_dir)"),
        ("pari_rev",  "Pari Revisione"),
        ("plm",       "PLM"),
        ("storico",   "Storico"),
        ("iss",       "ISS"),
        ("fiv",       "FIV loading"),
        ("heng",      "Hengelo"),
        ("error_plm", "PLM_Error"),
        ("tab",       "Tabellari"),
        ("log_dir",   "Log (opzionale)"),
    ]

    def __init__(self, app: SwarkyApp):
        super().__init__(app.root)
        self.app = app
        self.title("Settings")
        self.configure(bg=LIGHT_BG)
        self.resizable(False, True)
        self.grab_set()

        self.vars: Dict[str, tk.StringVar] = {}
        self._load_config_dict()

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        for r, (key, label) in enumerate(self.PATH_FIELDS):
            ttk.Label(frm, text=f"{label}:").grid(row=r, column=0, sticky="w", pady=4)
            self.vars[key] = tk.StringVar(value=self._paths.get(key, "") if self._paths.get(key) else "")
            ttk.Entry(frm, textvariable=self.vars[key], width=56).grid(row=r, column=1, sticky="ew", pady=4)
            ttk.Button(frm, text="Sfoglia…", command=lambda k=key: self._browse_dir(k)).grid(row=r, column=2, padx=(6,0))

        r = len(self.PATH_FIELDS)
        ttk.Label(frm, text="Avvio automatico (HH:MM):").grid(row=r, column=0, sticky="w", pady=8)
        self.time_var = tk.StringVar(value=self._auto_time or "")
        ttk.Entry(frm, textvariable=self.time_var, width=8).grid(row=r, column=1, sticky="w", pady=8)

        # Checkbox ACCEPT_PDF
        self.accept_pdf_var = tk.BooleanVar(value=bool(self._accept_pdf))
        ttk.Checkbutton(frm, text="Accetta PDF nella cartella Plotter", variable=self.accept_pdf_var).grid(
            row=r+1, column=0, columnspan=2, sticky="w", pady=(0,8)
        )
        
        # Checkbox LOG_PHASES
        self.log_phases_var = tk.BooleanVar(value=bool(self._log_phases))
        ttk.Checkbutton(frm, text="Mostra fasi (inizio/fine) nel file di log", variable=self.log_phases_var).grid(
            row=r+2, column=0, columnspan=2, sticky="w", pady=(0,8)
        )

        btns = ttk.Frame(frm)
        btns.grid(row=r+3, column=0, columnspan=3, sticky="e", pady=(12,0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right", padx=6)
        ttk.Button(btns, text="Salva", command=self._save).pack(side="right")

        frm.columnconfigure(1, weight=1)

    def _load_config_dict(self):
        data = {}
        try:
            data = json.loads(self.app.json_path.read_text(encoding="utf-8"))
        except Exception:
            pass
        self._paths = data.get("paths", {})
        self._auto_time = data.get("AUTO_TIME", "")
        self._log_level = data.get("LOG_LEVEL", "INFO")
        self._accept_pdf = data.get("ACCEPT_PDF", True)
        self._log_phases = data.get("LOG_PHASES", True)

    def _browse_dir(self, key: str) -> None:
        start = self.vars[key].get().strip() or str(Path.cwd())
        d = filedialog.askdirectory(initialdir=start, title=f"Scegli cartella per {key}")
        if d:
            self.vars[key].set(d)

    def _save(self) -> None:
        new_paths: Dict[str, str|None] = {}
        for key, _ in self.PATH_FIELDS:
            v = self.vars[key].get().strip()
            if key != "log_dir" and not v:
                messagebox.showerror("Errore", f"Il percorso '{key}' non può essere vuoto.")
                return
            if v:
                p = Path(v).expanduser()
                if not p.exists() or not p.is_dir():
                    messagebox.showerror("Errore", f"La cartella '{v}' non esiste o non è una directory.")
                    return
                new_paths[key] = str(p)
            else:
                new_paths[key] = None if key == "log_dir" else ""

        auto_time = (self.time_var.get() or "").strip()
        if auto_time:
            try:
                hh, mm = map(int, auto_time.split(":", 1))
                if not (0 <= hh <= 23 and 0 <= mm <= 59):
                    raise ValueError
                auto_time = f"{hh:02d}:{mm:02d}"
            except Exception:
                messagebox.showerror("Errore", "Orario non valido. Usa HH:MM (es. 08:30) o lascia vuoto.")
                return

        data_out = {
            "paths": new_paths,
            "AUTO_TIME": auto_time,
            "LOG_LEVEL": self._log_level,
            "ACCEPT_PDF": bool(self.accept_pdf_var.get()),
            "LOG_PHASES": bool(self.log_phases_var.get())
        }
        try:
            self.app.json_path.write_text(json.dumps(data_out, indent=2), encoding="utf-8")
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile salvare config.json:\n{e}")
            return

        self.app._reload_cfg()
        self.app._schedule_if_ready()
        self.app.refresh_plotter()

        messagebox.showinfo("OK", "Impostazioni salvate e applicate.")
        self.destroy()

class _TreeviewHandler(logging.Handler):
    def __init__(self, app: SwarkyApp):
        super().__init__()
        self.app = app

    def _remove_from_plotter_listbox(self, file_name: str) -> None:
        """Rimuove un item dalla listbox senza scandire la cartella (evita round-trip)."""
        try:
            lb = self.app.plotter_list
            names = lb.get(0, 'end')
            if file_name in names:
                idx = names.index(file_name)
                lb.delete(idx)
                self.app.update_counters()
        except Exception:
            pass

    def emit(self, record: logging.LogRecord) -> None:
        ui = getattr(record, "ui", None)
        if not ui:
            return

        kind = ui[0]
        ts = datetime.fromtimestamp(record.created)

        if kind == "processed":
            # ui = ("processed", file_name, process, compare, dest)
            file_name = ui[1]
            process   = ui[2] if len(ui) > 2 else ""
            compare   = ui[3] if len(ui) > 3 else ""
            dest      = ui[4] if len(ui) > 4 else ""
            def _add():
                self.app.insert_processed(ts.strftime("%d.%b.%Y"),
                                          ts.strftime("%H:%M:%S"),
                                          file_name, process, dest, compare)
                self._remove_from_plotter_listbox(file_name)
            self.app.root.after(0, _add)

        elif kind == "anomaly":
            # ui = ("anomaly", file_name, msg)
            file_name = ui[1]
            msg       = ui[2] if len(ui) > 2 else ""
            def _add():
                self.app.insert_anomaly(ts.strftime("%d.%b.%Y"),
                                        ts.strftime("%H:%M:%S"),
                                        file_name, msg)
                self._remove_from_plotter_listbox(file_name)
            self.app.root.after(0, _add)

        elif kind == "phase":
            # ui = ("phase", "Testo fase corrente")
            phase_text = ui[1] if len(ui) > 1 else ""
            self.app.root.after(0, lambda: self.app._phase_start(phase_text))

        elif kind in ("phase_end", "phase_done"):
            # ui ("phase_end", "Testo finale?")  oppure ("phase_done", elapsed_ms)
            def _end():
                final_text = None
                if kind == "phase_end":
                    final_text = ui[1] if len(ui) > 1 else None
                elif kind == "phase_done":
                    try:
                        elapsed_ms = int(ui[1]) if len(ui) > 1 else None
                        if elapsed_ms is not None:
                            final_text = f"Completato • {elapsed_ms} ms"
                    except Exception:
                        final_text = None
                self.app._phase_end(final_text)
            self.app.root.after(0, _end)

def main() -> None:
    SwarkyApp().run()

if __name__ == "__main__":
    main()
