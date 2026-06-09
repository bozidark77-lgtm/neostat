"""
app.py  —  RECONSTRUCTED SOURCE (from PyInstaller .pyc, Python 3.13)

Reconstructed from extracted bytecode. Widget options, geometry, colors, Serbian
labels, file-dialog titles and message text are EXACT. Layout/grid call order is
reconstructed from referenced names; verify against a 3.13 decompile of `app.pyc`.

Entry point of "AMS – Analiza izveštaja" (powered by Neostat™). Lets the user
pick the supplier .xlsx and the BREZA .xlsx, choose an output, then runs
convert.convert_xlsx_to_csv on both and analyze.generate_report on the results.
"""

import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from analyze import generate_report
from convert import convert_xlsx_to_csv


class ConvertLauncherApp:
    @staticmethod
    def _resource_path(base_dir, filename):
        """Resolve a bundled resource whether running frozen (PyInstaller) or not."""
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            p = Path(meipass) / filename
            if p.exists():
                return p
        return Path(base_dir) / filename

    def __init__(self, root):
        self.root = root
        root.title("AMS - Analiza izveštaja")
        root.geometry("900x430")
        root.resizable(0, 0)
        root.configure(bg="#EEF3FA")

        self.base_dir = Path(__file__).resolve().parent
        self.convert_script = self.base_dir / "convert.py"
        self.analyze_script = self.base_dir / "analyze.py"

        # window icon / logo (bundled resources)
        try:
            ico = self._resource_path(self.base_dir, "neostat.ico")
            if ico.exists():
                root.iconbitmap(default=str(ico))
            png = self._resource_path(self.base_dir, "neostat_logo.png")
            if png.exists():
                self._window_icon = tk.PhotoImage(file=str(png))
                root.iconphoto(True, self._window_icon)
        except Exception:
            pass

        self.supplier_var = tk.StringVar(value="")
        self.ams_var = tk.StringVar(value="")
        self.output_var = tk.StringVar(value="")
        self._build_ui()

    def _build_ui(self):
        outer = tk.Frame(self.root, bg="#EEF3FA", padx=20, pady=18)
        outer.pack(fill="both", expand=True)

        card = tk.Frame(outer, bg="white", bd=1, relief="solid", padx=16, pady=16)
        card.pack(fill="both", expand=True)

        tk.Label(card, text="AMS \u2013 Analiza izveštaja",
                 font=("Segoe UI", 16, "bold"), fg="#1F3A5F", bg="white", anchor="w") \
            .grid(row=0, column=0, columnspan=3, sticky="we")
        tk.Label(card, text="Izaberite ulazne Excel fajlove i lokaciju izlaza, pa kliknite Analiziraj.",
                 fg="#4A607A", bg="white", anchor="w") \
            .grid(row=1, column=0, columnspan=3, sticky="we", pady=(2, 12))

        def row(r, label_text, var, command):
            tk.Label(card, text=label_text, anchor="w", bg="white", fg="#23364D",
                     font=("Segoe UI", 9, "bold")).grid(row=r, column=0, sticky="w")
            tk.Entry(card, textvariable=var, width=78, relief="solid", bd=1) \
                .grid(row=r + 1, column=0, columnspan=2, padx=(0, 8), pady=(2, 10), sticky="we")
            tk.Button(card, text="Pretraži...", width=12, command=command,
                      bg="#E7EEF8", activebackground="#D9E6F7", relief="flat", cursor="hand2") \
                .grid(row=r + 1, column=2, pady=(2, 10), sticky="we")

        row(2, "IZVEŠTAJ DOBAVLJAČA", self.supplier_var, self._pick_supplier)
        row(4, "BREZA", self.ams_var, self._pick_ams)

        tk.Label(card, text="IZLAZNI FAJL", anchor="w", bg="white", fg="#23364D",
                 font=("Segoe UI", 9, "bold")).grid(row=6, column=0, sticky="w")
        tk.Entry(card, textvariable=self.output_var, width=78, relief="solid", bd=1) \
            .grid(row=7, column=0, columnspan=2, padx=(0, 8), pady=(2, 10), sticky="we")
        tk.Button(card, text="Sačuvaj kao...", width=12, command=self._pick_output,
                  bg="#E7EEF8").grid(row=7, column=2, pady=(2, 10), sticky="we")

        btns = tk.Frame(card, bg="white")
        btns.grid(row=8, column=0, columnspan=3, sticky="we")
        tk.Button(btns, text="Analiziraj", width=16, command=self._analyze,
                  bg="#1E5AA8", fg="white", activebackground="#2D6BBE", activeforeground="white",
                  relief="flat", cursor="hand2", font=("Segoe UI", 10, "bold")) \
            .pack(side="left", padx=(0, 8))
        tk.Button(btns, text="Izlaz", width=12, command=self.root.destroy,
                  bg="#F1F4F8", activebackground="#E6ECF5", relief="flat").pack(side="left")

        card.columnconfigure(0, weight=1)
        tk.Label(self.root, text="powered by Neostat\u2122", font=("Segoe UI", 8),
                 fg="#6E7C8F", bg="#EEF3FA") \
            .place(relx=1.0, rely=1.0, anchor="se", x=-12, y=-8)

    def _pick_supplier(self):
        p = filedialog.askopenfilename(
            title="Izaberi IZVESTAJ DOBAVLJACA xlsx", initialdir=str(self.base_dir),
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if p:
            self.supplier_var.set(p)

    def _pick_ams(self):
        p = filedialog.askopenfilename(
            title="Izaberi BREZA xlsx", initialdir=str(self.base_dir),
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if p:
            self.ams_var.set(p)

    def _pick_output(self):
        current = self.output_var.get().strip()
        initialdir = str(Path(current).parent) if current and Path(current).parent.exists() else str(self.base_dir)
        initialfile = Path(current).name if current and Path(current).suffix.lower() == ".xlsx" else "Nalaz.xlsx"
        p = filedialog.asksaveasfilename(
            title="Izaberi gde da sačuvaš izlazni fajl", initialdir=initialdir,
            initialfile=initialfile, defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if p:
            self.output_var.set(p)

    @staticmethod
    def _validate_xlsx(path_str, label):
        p = Path(path_str).expanduser().resolve()
        if not p.exists():
            raise ValueError(label + ": fajl ne postoji.")
        if not p.is_file() or p.suffix.lower() != ".xlsx":
            raise ValueError(label + ": mora biti .xlsx fajl.")
        return p

    def _run_convert_for_file(self, xlsx_path):
        csv_path = xlsx_path.with_suffix(".csv")
        convert_xlsx_to_csv(str(xlsx_path), str(csv_path))
        print("Konvertovano: " + xlsx_path.name + " -> " + csv_path.name)
        return csv_path

    def _analyze(self):
        temp_csvs = []
        try:
            sup_xlsx = self._validate_xlsx(self.supplier_var.get().strip(), "IZVESTAJ DOBAVLJACA")
            ams_xlsx = self._validate_xlsx(self.ams_var.get().strip(), "BREZA")

            out_str = self.output_var.get().strip()
            out_path = Path(out_str).expanduser().resolve() if out_str else (self.base_dir / "Nalaz.xlsx")
            if out_path.suffix.lower() != ".xlsx":
                out_path = out_path.with_suffix(".xlsx")
            out_path.parent.mkdir(parents=True, exist_ok=True)
        except ValueError as e:
            messagebox.showerror("Greška", str(e))
            return

        self.root.config(cursor="wait")
        self.root.update_idletasks()
        try:
            sup_csv = self._run_convert_for_file(sup_xlsx)
            ams_csv = self._run_convert_for_file(ams_xlsx)
            temp_csvs += [sup_csv, ams_csv]
            generate_report(str(sup_csv), str(ams_csv), str(out_path))
            messagebox.showinfo("Uspeh", "Analiza uspešna.\n\nKreiran izlazni fajl:\n" + str(out_path))
            self.root.destroy()
        except Exception as e:
            messagebox.showerror("Greška",
                                 "Konverzija je uspela, ali analiza nije uspela.\n\nDetalji:\n" + str(e))
        finally:
            self.root.config(cursor="")
            for c in temp_csvs:
                try:
                    if c.exists():
                        c.unlink()
                except OSError:
                    pass


def main():
    root = tk.Tk()
    ConvertLauncherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
