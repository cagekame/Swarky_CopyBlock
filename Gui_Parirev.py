# Gui_Parirev.py
from __future__ import annotations
import os, sys, subprocess, shutil
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
from Swarky import BASE_NAME, map_location, _docno_from_match

LIGHT_BG = "#eef3f9"
NAVY_BG  = "#000080"
NAVY_SEL = "#133869"
FG_LIGHT = "light gray"
FG_WHITE = "white"

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

class PariRevWindow(tk.Toplevel):
    def __init__(self, master: tk.Misc, cfg) -> None:
        super().__init__(master)
        self.title("FSR")
        self.configure(bg=LIGHT_BG)
        self.resizable(True, True)
        self.cfg = cfg

        # ===== griglia finestra: 2 colonne sopra + info + LOG sotto =====
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)  # il LOG si espande

        PAD = 8
        LABEL_PADY = (0,4)
        BTN_W = 20

        # ----- colonna sinistra -----
        left = tk.Frame(self, bg=LIGHT_BG)
        left.grid(row=0, column=0, sticky="nsew", padx=PAD, pady=PAD)
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)

        tk.Label(left, text="Same Revision", bg=LIGHT_BG).grid(row=0, column=0, sticky="w", pady=LABEL_PADY)

        self.lst_srfolder = tk.Listbox(
            left,
            bg=NAVY_BG, fg=FG_LIGHT, selectbackground=NAVY_SEL, selectforeground=FG_WHITE,
            width=30, exportselection=False, selectmode="browse", highlightthickness=0, borderwidth=0
        )
        self.lst_srfolder.grid(row=1, column=0, sticky="nsew")
        self.lst_srfolder.bind("<Double-Button-1>", self._open_selected)
        self.lst_srfolder.bind("<<ListboxSelect>>", self._on_select)

        # ----- colonna destra -----
        right = tk.Frame(self, bg=LIGHT_BG)
        right.grid(row=0, column=1, sticky="nsew", padx=PAD, pady=PAD)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        tk.Label(right, text="Azioni", bg=LIGHT_BG).grid(row=0, column=0, sticky="w", pady=LABEL_PADY)

        btns = tk.Frame(right, bg=LIGHT_BG)
        btns.grid(row=1, column=0, sticky="nsew")

        self.btn_sr_go     = tk.Button(btns, text="Start Process",  width=BTN_W, command=self._start_process_worker)
        self.btn_getnumber = tk.Button(btns, text="Get Number",     width=BTN_W, command=self._not_implemented)
        self.btn_goto      = tk.Button(btns, text="GoTo Folder",    width=BTN_W, command=self._goto_dest_folder)
        self.btn_srdir     = tk.Button(btns, text="Goto Sr Folder", width=BTN_W, command=self._goto_sr_folder)

        for i, b in enumerate((self.btn_sr_go, self.btn_getnumber, self.btn_goto, self.btn_srdir)):
            pady = (0,3) if i == 0 else (3,0) if i == 3 else 3
            b.pack(fill="x", expand=True, pady=pady)

        # ----- info dimensione disegno -----
        self._size_var = tk.StringVar(value="Drawing size (Kilobyte): 0")
        tk.Label(self, textvariable=self._size_var, bg=LIGHT_BG).grid(
            row=1, column=0, columnspan=2, sticky="w", padx=PAD, pady=(0,PAD)
        )

        # ----- LOG sotto -----
        logf = tk.Frame(self, bg=LIGHT_BG)
        logf.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=PAD, pady=(0, PAD))
        logf.columnconfigure(0, weight=1)
        logf.rowconfigure(1, weight=1)

        tk.Label(logf, text="LOG", bg=LIGHT_BG).grid(row=0, column=0, sticky="w", pady=(0,4))
        self.lst_log = tk.Listbox(logf, bg=NAVY_BG, fg=FG_LIGHT, height=10, highlightthickness=0, borderwidth=0,
                                  selectbackground=NAVY_SEL, selectforeground=FG_WHITE)
        self.lst_log.grid(row=1, column=0, sticky="nsew")

        # centro e popolo
        self.transient(master)
        self._center_on_parent()
        self.refresh_list()

        # Blocca dimensione minima sull’attuale (solo su questa finestra)
        self.update_idletasks()
        self.minsize(self.winfo_width(), self.winfo_height())

    # -------- listbox refresh (pausa se focus, ripristina selezione) --------
    def refresh_list(self) -> None:
        lb = self.lst_srfolder
        try:
            if self.focus_get() is lb:
                return
        except Exception:
            pass

        saved = None
        sel = lb.curselection()
        if sel:
            saved = lb.get(sel[0])

        try:
            base = self.cfg.PARI_REV_DIR
            patterns = ["*.tif", "*.TIF"]
            if getattr(self.cfg, "ACCEPT_PDF", True):
                patterns += ["*.pdf", "*.PDF"]
            names = sorted({p.name for pat in patterns for p in base.glob(pat) if p.is_file()}, key=str.lower)
        except Exception:
            names = []

        old = list(lb.get(0, tk.END))
        if old == names:
            return

        lb.delete(0, tk.END)
        for nm in names:
            lb.insert(tk.END, nm)

        if saved in names:
            idx = names.index(saved)
            lb.selection_set(idx)
            lb.see(idx)

        self._update_size_label()

    # -------- utils/log --------
    def _log(self, msg: str) -> None:
        self.lst_log.insert(tk.END, msg)
        self.lst_log.see(tk.END)

    def _pretty_loc(self, loc: dict) -> str:
        return f"{loc.get('log_name','?')} / {loc.get('arch_tif_loc','?')}"

    def _update_size_label(self) -> None:
        sel = self.lst_srfolder.curselection()
        if not sel:
            self._size_var.set("Drawing size (Kilobyte): 0")
            return
        name = self.lst_srfolder.get(sel[0])
        p = Path(self.cfg.PARI_REV_DIR) / name
        try:
            size_kb = int(p.stat().st_size / 1024)
        except Exception:
            size_kb = 0
        self._size_var.set(f"Drawing size (Kilobyte): {size_kb}")

    def _on_select(self, _evt=None) -> None:
        self._copy_docno_prefix()
        self._update_size_label()

    # -------- azioni --------
    def _open_selected(self, _evt=None) -> None:
        sel = self.lst_srfolder.curselection()
        if not sel:
            return
        p = Path(self.cfg.PARI_REV_DIR) / self.lst_srfolder.get(sel[0])
        if p.exists():
            _open_path(p)

    def _copy_docno_prefix(self, _evt=None) -> None:
        sel = self.lst_srfolder.curselection()
        if not sel:
            return
        name = self.lst_srfolder.get(sel[0])
        m = BASE_NAME.fullmatch(name)
        if not m:
            return
        try:
            docno = _docno_from_match(m)
            self.clipboard_clear()
            self.clipboard_append(docno)
        except Exception:
            pass

    def _goto_sr_folder(self) -> None:
        _open_path(self.cfg.PARI_REV_DIR)

    def _goto_dest_folder(self) -> None:
        sel = self.lst_srfolder.curselection()
        if not sel:
            messagebox.showinfo("FSR", "Seleziona un file nella lista.")
            return
        name = self.lst_srfolder.get(sel[0])
        m = BASE_NAME.fullmatch(name)
        if not m:
            messagebox.showwarning("FSR", "Nome file non valido.")
            return
        loc = map_location(m, self.cfg)
        dest_dir = loc["dir_tif_loc"]
        dest_dir.mkdir(parents=True, exist_ok=True)
        _open_path(dest_dir)

    def _start_process_worker(self) -> None:
        try:
            self.btn_sr_go.config(state="disabled")
            sel = self.lst_srfolder.curselection()
            if len(sel) != 1:
                messagebox.showinfo("FSR", "Seleziona **un solo** file.")
                return

            nm = self.lst_srfolder.get(sel[0])
            src = self.cfg.PARI_REV_DIR / nm
            if not src.exists():
                messagebox.showwarning("FSR", f"Il file non esiste più in Pari Revisione:\n{nm}")
                return

            m = BASE_NAME.fullmatch(nm)
            if not m:
                messagebox.showwarning("FSR", f"Nome file non valido (regex):\n{nm}")
                return

            try:
                loc = map_location(m, self.cfg)
                target_dir = loc["dir_tif_loc"]
                human_loc = self._pretty_loc(loc)
            except Exception as e:
                messagebox.showerror("FSR", f"map_location fallita:\n{e}")
                return

            dest = target_dir / nm
            if not dest.exists():
                messagebox.showwarning("FSR", f"NON presente in Archivio (non aggiornato):\n{nm}\n→ {human_loc}")
                self._log(f"{nm} → {human_loc}: assente in archivio")
                return

            try:
                target_dir.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dest)  # overwrite
                self._log(f"{nm} → {human_loc}: copiato (overwrite)")
            except Exception as e:
                self._log(f"{nm} → {human_loc}: ERRORE copia → {e}")
                messagebox.showerror("FSR", f"Errore durante la copia:\n{e}")
        finally:
            self.btn_sr_go.config(state="normal")

    # -------- window helpers --------
    def _center_on_parent(self) -> None:
        try:
            self.update_idletasks()
            px = self.master.winfo_rootx()
            py = self.master.winfo_rooty()
            pw = self.master.winfo_width()
            ph = self.master.winfo_height()
            w  = self.winfo_width()
            h  = self.winfo_height()
            x = px + (pw - w)//2
            y = py + (ph - h)//2
            self.geometry(f"+{max(x,0)}+{max(y,0)}")
        except Exception:
            pass

    def _not_implemented(self) -> None:
        messagebox.showinfo("PariRev", "Funzione non ancora implementata.")
