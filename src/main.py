"""
main.py – TN-Doku-Konfigurator
Tkinter-GUI für die automatische Erstellung von Teilnehmer-Ablagesystemen.
"""
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Lokale Module
from build_info import BUILD_INFO
from core import get_app_dir, run_all
from csv_reader import read_participants

APP_TITLE = f"TN-Doku-Konfigurator v{BUILD_INFO['version']}"
WINDOW_W = 780
WINDOW_H = 700
USER_DOC_FILE = "DOKUMENTATION_ANWENDER.md"
TECH_DOC_FILE = "DOKUMENTATION_TECHNIK.md"


def find_excel_executable() -> str | None:
    """Ermittelt den von Windows registrierten Pfad zu Microsoft Excel."""
    try:
        import winreg
    except ImportError:
        return None

    key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
    registry_views = (
        0,
        winreg.KEY_WOW64_64KEY,
        winreg.KEY_WOW64_32KEY,
    )
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for registry_view in registry_views:
            try:
                with winreg.OpenKey(
                    hive, key_path, 0, winreg.KEY_READ | registry_view
                ) as key:
                    excel_path, _ = winreg.QueryValueEx(key, "")
                excel_path = os.path.expandvars(str(excel_path)).strip('"')
                if os.path.isfile(excel_path):
                    return excel_path
            except OSError:
                continue
    return None


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self._app_dir = get_app_dir()
        self.title(APP_TITLE)
        self.resizable(True, True)
        self.minsize(620, 480)
        self._center_window(WINDOW_W, WINDOW_H)
        self._set_icon()

        self._participants: list[dict] = []
        self._loaded_csv_path: str | None = None
        self._csv_is_being_edited = False
        self._build()
        self._load_defaults()

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def _build(self):
        # ── Rahmen: obere Eingaben ──────────────────────────────────────
        frm_top = ttk.LabelFrame(self, text="Konfiguration", padding=8)
        frm_top.pack(fill="x", padx=10, pady=(10, 4))
        frm_top.columnconfigure(1, weight=1)

        # CSV-Datei
        ttk.Label(frm_top, text="CSV-Datei:").grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=3
        )
        self._csv_var = tk.StringVar()
        self._csv_entry = ttk.Entry(frm_top, textvariable=self._csv_var)
        self._csv_entry.grid(row=0, column=1, sticky="ew", pady=3)
        self._csv_entry.bind("<FocusOut>", self._on_csv_entry_focus_out)
        self._csv_entry.bind("<Return>", self._on_csv_entry_return)
        self._edit_csv_btn = ttk.Button(
            frm_top, text="Bearbeiten", command=self._edit_csv, width=10
        )
        self._edit_csv_btn.grid(
            row=0, column=2, padx=(4, 0), pady=3
        )

        # Ausgabe-Ordner
        ttk.Label(frm_top, text="Ausgabe-Ordner:").grid(
            row=1, column=0, sticky="w", padx=(0, 6), pady=3
        )
        self._out_var = tk.StringVar()
        self._out_entry = ttk.Entry(frm_top, textvariable=self._out_var)
        self._out_entry.grid(row=1, column=1, sticky="ew", pady=3)
        ttk.Button(frm_top, text="Ordner …", command=self._browse_output, width=10).grid(
            row=1, column=2, padx=(4, 0), pady=3
        )

        # Ablagesystem-Ordner
        ttk.Label(frm_top, text="Ablagesystem-Ordner:").grid(
            row=2, column=0, sticky="w", padx=(0, 6), pady=3
        )
        self._lists_var = tk.StringVar()
        self._lists_entry = ttk.Entry(frm_top, textvariable=self._lists_var)
        self._lists_entry.grid(row=2, column=1, sticky="ew", pady=3)
        ttk.Button(frm_top, text="Ordner …", command=self._browse_lists, width=10).grid(
            row=2, column=2, padx=(4, 0), pady=3
        )

        # ── Teilnehmer-Vorschau ─────────────────────────────────────────
        frm_list = ttk.LabelFrame(self, text="Teilnehmer (0)", padding=8)
        frm_list.pack(fill="both", expand=True, padx=10, pady=4)
        self._frm_list = frm_list

        cols = ("Name", "Maßnahme", "Maßnahmekürzel")
        self._tree = ttk.Treeview(frm_list, columns=cols, show="headings", height=8)
        for col in cols:
            self._tree.heading(col, text=col)
        self._tree.column("Name", width=260)
        self._tree.column("Maßnahme", width=140)
        self._tree.column("Maßnahmekürzel", width=120)
        sb = ttk.Scrollbar(frm_list, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # ── Aktions-Schaltflächen ───────────────────────────────────────
        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", padx=10, pady=4)

        ttk.Button(
            frm_btn, text="Doku_Anwender", command=self._open_user_doc
        ).pack(side="left", padx=(0, 6))

        ttk.Button(
            frm_btn, text="Doku_Technik", command=self._open_tech_doc
        ).pack(side="left", padx=(0, 6))

        self._run_btn = tk.Button(
            frm_btn,
            text="Dokumentation erstellen",
            command=self._start_run,
            state="disabled",
            font=("Segoe UI", 10, "bold"),
            bg="#1f7a1f",
            fg="white",
            activebackground="#2a8f2a",
            activeforeground="white",
            relief="raised",
            bd=1,
            padx=10,
        )
        self._run_btn.pack(side="left")

        ttk.Button(
            frm_btn, text="Ergebnis-Ordner öffnen",
            command=self._open_result
        ).pack(side="right")

        # ── Fortschritts-Leiste ─────────────────────────────────────────
        self._progress = ttk.Progressbar(self, mode="indeterminate")
        self._progress.pack(fill="x", padx=10, pady=(0, 2))

        # ── Log-Ausgabe ─────────────────────────────────────────────────
        frm_log = ttk.LabelFrame(self, text="Protokoll", padding=4)
        frm_log.pack(fill="x", padx=10, pady=(0, 8))

        self._log_text = tk.Text(
            frm_log, height=9, state="disabled",
            font=("Consolas", 9), wrap="word",
            background="#1e1e1e", foreground="#d4d4d4",
            insertbackground="white",
        )
        log_sb = ttk.Scrollbar(frm_log, orient="vertical", command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=log_sb.set)
        self._log_text.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

        # Statusleiste
        self._status_var = tk.StringVar(value="Bereit.")
        ttk.Label(self, textvariable=self._status_var, anchor="w").pack(
            fill="x", padx=10, pady=(0, 4)
        )

        # Ergebnis-Pfad für "Ordner öffnen"
        self._result_path: str | None = None

    # ------------------------------------------------------------------
    # Standard-Pfade
    # ------------------------------------------------------------------

    def _load_defaults(self):
        # CSV: bevorzugt relativ in data/, Legacy-Fallback im App-Verzeichnis
        default_csv_rel = os.path.join("data", "Teilnehmer_Beginn.CSV")
        legacy_csv_rel = "Teilnehmer_Beginn.CSV"
        if os.path.isfile(self._resolve_path(default_csv_rel)):
            self._csv_var.set(default_csv_rel)
        elif os.path.isfile(self._resolve_path(legacy_csv_rel)):
            self._csv_var.set(legacy_csv_rel)
        else:
            self._csv_var.set(default_csv_rel)

        # Ausgabe: Standard relativ zum EXE-Ordner
        self._out_var.set("output")

        # Ablagesystem: bevorzugt relativ in data/, Legacy-Fallback im App-Verzeichnis
        default_lists_rel = os.path.join("data", "Ablagesystem")
        legacy_lists_rel = "_Listen"
        if os.path.isdir(self._resolve_path(default_lists_rel)):
            self._lists_var.set(default_lists_rel)
        elif os.path.isdir(self._resolve_path(legacy_lists_rel)):
            self._lists_var.set(legacy_lists_rel)
        else:
            self._lists_var.set(default_lists_rel)

        # Ggf. direkt laden
        if os.path.isfile(self._resolve_path(self._csv_var.get())):
            self._load_csv(silent=True)

    def _resolve_path(self, path_value: str) -> str:
        """Löst relative Pfade relativ zum EXE-/App-Verzeichnis auf."""
        path_value = (path_value or "").strip()
        if not path_value:
            return ""
        expanded = os.path.expandvars(os.path.expanduser(path_value))
        if not os.path.isabs(expanded):
            expanded = os.path.join(self._app_dir, expanded)
        return os.path.normpath(expanded)

    # ------------------------------------------------------------------
    # Datei-/Ordner-Dialoge
    # ------------------------------------------------------------------

    def _edit_csv(self):
        """Öffnet die konfigurierte CSV in einer eigenen Excel-Instanz."""
        if self._csv_is_being_edited:
            return

        csv_path = self._resolve_path(self._csv_var.get())
        if not csv_path or not os.path.isfile(csv_path):
            messagebox.showerror(
                "CSV-Datei nicht gefunden",
                "Bitte geben Sie eine vorhandene CSV-Datei im Eingabefeld an.",
            )
            return

        excel_path = find_excel_executable()
        if not excel_path:
            messagebox.showerror(
                "Microsoft Excel nicht gefunden",
                "Zum Bearbeiten der CSV-Datei wird Microsoft Excel benötigt.",
            )
            return

        self._csv_is_being_edited = True
        self._edit_csv_btn.configure(state="disabled")
        self._run_btn.configure(state="disabled")
        self._status_var.set("CSV wird in Excel bearbeitet …")

        def worker():
            try:
                # /x startet eine eigene Excel-Instanz. Daher endet der Prozess,
                # sobald die bearbeitete CSV geschlossen wird.
                subprocess.run([excel_path, "/x", csv_path], check=False)
                self.after(0, self._on_csv_editor_closed)
            except OSError as e:
                self.after(0, self._on_csv_editor_error, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def _on_csv_editor_closed(self):
        """Lädt die Teilnehmerdaten nach dem Schließen der Excel-Datei neu."""
        self._csv_is_being_edited = False
        self._edit_csv_btn.configure(state="normal")
        self._load_csv(silent=False)
        if self._participants:
            self._status_var.set(
                f"CSV nach Bearbeitung neu geladen: "
                f"{os.path.basename(self._loaded_csv_path or '')}"
            )

    def _on_csv_editor_error(self, message: str):
        self._csv_is_being_edited = False
        self._edit_csv_btn.configure(state="normal")
        self._run_btn.configure(state="normal" if self._participants else "disabled")
        self._status_var.set("Excel konnte nicht gestartet werden.")
        messagebox.showerror("Excel konnte nicht gestartet werden", message)

    def _on_csv_entry_focus_out(self, _event=None):
        """Lädt CSV still im Hintergrund, wenn das Feld verlassen wird."""
        self._load_csv(silent=True)

    def _on_csv_entry_return(self, _event=None):
        """Lädt CSV explizit bei Enter im CSV-Eingabefeld."""
        self._load_csv(silent=False)

    def _browse_output(self):
        path = filedialog.askdirectory(
            title="Ausgabe-Ordner auswählen",
            initialdir=self._resolve_path(self._out_var.get()) or self._app_dir,
        )
        if path:
            self._out_var.set(os.path.abspath(path))

    def _browse_lists(self):
        path = filedialog.askdirectory(
            title="Ablagesystem-Ordner auswählen",
            initialdir=self._resolve_path(self._lists_var.get()) or self._app_dir,
        )
        if path:
            self._lists_var.set(os.path.abspath(path))

    # ------------------------------------------------------------------
    # CSV laden & Vorschau
    # ------------------------------------------------------------------

    def _load_csv(self, silent: bool = False):
        csv_input = self._csv_var.get().strip()
        if not csv_input:
            if not silent:
                messagebox.showwarning("Kein Pfad", "Bitte zuerst eine CSV-Datei auswählen.")
            return
        csv_path = self._resolve_path(csv_input)
        try:
            self._participants = read_participants(csv_path)
            self._loaded_csv_path = csv_path
            self._refresh_tree()
            self._run_btn.configure(state="normal")
            self._status_var.set(
                f"{len(self._participants)} Teilnehmer geladen aus: {os.path.basename(csv_path)}"
            )
        except Exception as e:
            self._participants = []
            self._loaded_csv_path = None
            self._refresh_tree()
            self._run_btn.configure(state="disabled")
            if not silent:
                messagebox.showerror("CSV-Fehler", str(e))
            self._status_var.set(f"Fehler: {e}")

    def _refresh_tree(self):
        self._tree.delete(*self._tree.get_children())
        for p in self._participants:
            self._tree.insert(
                "",
                "end",
                values=(
                    p.get("Name", ""),
                    p.get("Maßnahme", ""),
                    p.get("Maßnahmekürzel", ""),
                ),
            )
        count = len(self._participants)
        self._frm_list.configure(text=f"Teilnehmer ({count})")

    # ------------------------------------------------------------------
    # Hauptlauf
    # ------------------------------------------------------------------

    def _start_run(self):
        if self._csv_is_being_edited:
            messagebox.showinfo(
                "CSV wird bearbeitet",
                "Bitte schließen Sie die CSV-Datei in Excel. Anschließend wird sie "
                "automatisch neu geladen.",
            )
            return

        current_csv_path = self._resolve_path(self._csv_var.get())
        if current_csv_path and current_csv_path != self._loaded_csv_path:
            self._load_csv(silent=False)

        if not self._participants:
            messagebox.showwarning("Keine Daten", "Bitte zuerst eine CSV-Datei laden.")
            return
        out_dir = self._resolve_path(self._out_var.get())
        if not out_dir:
            messagebox.showerror("Ungültiger Pfad", "Bitte einen gültigen Ausgabe-Ordner wählen.")
            return
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror(
                "Ungültiger Pfad",
                f"Der Ausgabe-Ordner konnte nicht erstellt werden:\n{out_dir}\n\n{e}"
            )
            return
        lists_dir = self._resolve_path(self._lists_var.get())
        if not lists_dir or not os.path.isdir(lists_dir):
            messagebox.showerror(
                "Fehlender Ablagesystem-Ordner",
                f"Der Ablagesystem-Ordner wurde nicht gefunden:\n{lists_dir}\n\n"
                "Bitte den Pfad manuell konfigurieren."
            )
            return

        self._run_btn.configure(state="disabled")
        self._progress.start(12)
        self._result_path = None
        self._log_clear()
        self._status_var.set("Wird verarbeitet …")

        participants_snapshot = list(self._participants)

        def worker():
            try:
                result = run_all(
                    participants=participants_snapshot,
                    output_dir=out_dir,
                    lists_dir=lists_dir,
                    log=self._log_append_thread,
                )
                self.after(0, self._on_success, result)
            except Exception as e:
                self.after(0, self._on_error, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, result_path: str):
        self._progress.stop()
        self._run_btn.configure(state="normal")
        self._result_path = result_path
        self._status_var.set(f"Fertig: {result_path}")
        messagebox.showinfo(
            "Fertig",
            f"Dokumentation erfolgreich erstellt:\n\n{result_path}",
        )

    def _on_error(self, message: str):
        self._progress.stop()
        self._run_btn.configure(state="normal")
        self._status_var.set(f"Fehler aufgetreten.")
        messagebox.showerror("Fehler", message)

    # ------------------------------------------------------------------
    # Ergebnis-Ordner öffnen
    # ------------------------------------------------------------------

    def _open_result(self):
        path = self._resolve_path(self._out_var.get())
        if path and os.path.isdir(path):
            os.startfile(path)
        else:
            messagebox.showinfo(
                "Ordner nicht gefunden",
                "Der konfigurierte Ausgabe-Ordner wurde nicht gefunden."
            )

    def _open_user_doc(self):
        self._open_doc(USER_DOC_FILE)

    def _open_tech_doc(self):
        self._open_doc(TECH_DOC_FILE)

    def _open_doc(self, file_name: str):
        app_dir = self._app_dir
        candidates = [
            os.path.join(app_dir, "docs", file_name),
            os.path.join(app_dir, file_name),
        ]
        for path in candidates:
            if os.path.isfile(path):
                try:
                    os.startfile(path)
                    return
                except Exception as e:
                    messagebox.showerror("Dokument konnte nicht geöffnet werden", str(e))
                    return
        messagebox.showwarning(
            "Dokument nicht gefunden",
            f"Die Datei '{file_name}' wurde weder im Ordner 'docs' noch im Anwendungsverzeichnis gefunden."
        )

    # ------------------------------------------------------------------
    # Log-Hilfsmethoden
    # ------------------------------------------------------------------

    def _log_clear(self):
        self._log_text.configure(state="normal")
        self._log_text.delete("1.0", "end")
        self._log_text.configure(state="disabled")

    def _log_append(self, msg: str):
        self._log_text.configure(state="normal")
        self._log_text.insert("end", msg + "\n")
        self._log_text.see("end")
        self._log_text.configure(state="disabled")

    def _log_append_thread(self, msg: str):
        """Thread-sicherer Log-Aufruf via after()."""
        self.after(0, self._log_append, msg)

    # ------------------------------------------------------------------
    # Fensterzentrierungen
    # ------------------------------------------------------------------

    def _center_window(self, w: int, h: int):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _set_icon(self):
        """Setzt das Fenster-Icon aus app_icon.ico."""
        if getattr(sys, 'frozen', False):
            # EXE-Modus: PyInstaller entpackt datas nach _MEIPASS
            meipass_dir = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
            icon_path = os.path.join(meipass_dir, "app_icon.ico")
        else:
            # Entwicklungsmodus: Icon liegt im src/-Verzeichnis neben main.py
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.ico")
        if os.path.isfile(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Einstiegspunkt
# ---------------------------------------------------------------------------

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
