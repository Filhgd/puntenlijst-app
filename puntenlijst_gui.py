#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
puntenlijst_gui.py
==================

Grafische app rond puntenlijst_core.py.

Gebruik: sleep een map (of losse PDF's) op het venster, of klik op een van
de knoppen. Het Excel-bestand komt in dezelfde map als de PDF's en wordt
daarna met één klik geopend.

Werkt op macOS en Windows. Drag & drop via tkinterdnd2; zonder die
bibliotheek werkt de app ook, maar dan enkel met de knoppen.
"""

import os
import re
import sys
import queue
import threading
import subprocess
import traceback

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import puntenlijst_core as core

# Drag & drop is optioneel: zonder tkinterdnd2 blijven de knoppen werken.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    HAS_DND = False


APP_TITLE = "Puntenlijst-generator"
BG = "#F4F6FA"
ACCENT = "#305496"
DROP_BG = "#E8EDF7"
DROP_BG_ACTIVE = "#D2DDF2"
OK_COLOR = "#1E7B34"
WARN_COLOR = "#B58900"
ERR_COLOR = "#B00020"


def split_dnd_paths(data):
    """Zet de ruwe drag&drop-string om in een lijst paden.

    Tk levert paden met spaties tussen accolades: {/pad/met spaties/x.pdf}.
    """
    tokens = re.findall(r"\{[^}]*\}|\S+", data)
    return [t[1:-1] if t.startswith("{") and t.endswith("}") else t
            for t in tokens]


def open_path(path):
    """Open een bestand of map met het standaardprogramma."""
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def reveal_in_folder(path):
    """Toon het bestand in Finder/Verkenner."""
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path])
        elif os.name == "nt":
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(path)])
    except Exception:
        pass


class App:
    def __init__(self, root):
        self.root = root
        self.busy = False
        self.result_path = None
        self.msg_queue = queue.Queue()

        root.title(APP_TITLE)
        root.geometry("720x560")
        root.minsize(560, 460)
        root.configure(bg=BG)

        # ---- kop ----
        tk.Label(
            root, text=APP_TITLE, font=("Helvetica", 20, "bold"),
            bg=BG, fg=ACCENT,
        ).pack(pady=(16, 2))
        tk.Label(
            root,
            text="Maakt één Excel met alle examenpunten uit de deliberatie-PDF's",
            font=("Helvetica", 12), bg=BG, fg="#555555",
        ).pack()

        # ---- dropzone ----
        self.drop = tk.Label(
            root,
            text=self._drop_text(),
            font=("Helvetica", 14),
            bg=DROP_BG, fg=ACCENT,
            relief="ridge", bd=2,
            justify="center",
        )
        self.drop.pack(fill="x", padx=24, pady=(14, 8), ipady=34)

        if HAS_DND:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self.on_drop)
            self.drop.dnd_bind("<<DropEnter>>",
                               lambda e: self.drop.configure(bg=DROP_BG_ACTIVE))
            self.drop.dnd_bind("<<DropLeave>>",
                               lambda e: self.drop.configure(bg=DROP_BG))

        # ---- knoppen ----
        btns = tk.Frame(root, bg=BG)
        btns.pack(pady=4)
        self.btn_folder = tk.Button(
            btns, text="Kies een map…", font=("Helvetica", 12),
            command=self.pick_folder, padx=12, pady=4,
        )
        self.btn_folder.pack(side="left", padx=6)
        self.btn_files = tk.Button(
            btns, text="Kies PDF-bestanden…", font=("Helvetica", 12),
            command=self.pick_files, padx=12, pady=4,
        )
        self.btn_files.pack(side="left", padx=6)
        self.btn_open = tk.Button(
            btns, text="Open Excel", font=("Helvetica", 12, "bold"),
            command=self.open_result, padx=12, pady=4, state="disabled",
        )
        self.btn_open.pack(side="left", padx=6)
        self.btn_reveal = tk.Button(
            btns, text="Toon in map", font=("Helvetica", 12),
            command=self.reveal_result, padx=12, pady=4, state="disabled",
        )
        self.btn_reveal.pack(side="left", padx=6)

        # ---- statusregel ----
        self.status = tk.Label(root, text="Klaar om te starten.",
                               font=("Helvetica", 12, "bold"), bg=BG, fg="#333333")
        self.status.pack(pady=(6, 0))

        # ---- logvenster ----
        self.log_box = scrolledtext.ScrolledText(
            root, font=("Courier", 11), height=12, state="disabled",
            bg="white", relief="sunken", bd=1,
        )
        self.log_box.pack(fill="both", expand=True, padx=24, pady=(6, 16))

        self.root.after(100, self._poll_queue)

    def _drop_text(self):
        if HAS_DND:
            return ("Sleep hier de map met deliberatie-PDF's\n"
                    "(of losse PDF-bestanden)")
        return ("Gebruik de knoppen hieronder om de map\n"
                "of de PDF-bestanden te kiezen")

    # ------------------------------------------------------------------ log
    def log(self, msg):
        self.msg_queue.put(("log", msg))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self.log_box.configure(state="normal")
                    self.log_box.insert("end", payload + "\n")
                    self.log_box.see("end")
                    self.log_box.configure(state="disabled")
                elif kind == "done":
                    self._on_done(payload)
                elif kind == "error":
                    self._on_error(payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # ------------------------------------------------------------- invoer
    def on_drop(self, event):
        self.drop.configure(bg=DROP_BG)
        paths = split_dnd_paths(event.data)
        self.start(paths)

    def pick_folder(self):
        folder = filedialog.askdirectory(title="Kies de map met de PDF's")
        if folder:
            self.start([folder])

    def pick_files(self):
        files = filedialog.askopenfilenames(
            title="Kies de deliberatie-PDF's",
            filetypes=[("PDF-bestanden", "*.pdf")],
        )
        if files:
            self.start(list(files))

    # ---------------------------------------------------------- verwerking
    def start(self, paths):
        if self.busy:
            messagebox.showinfo(APP_TITLE, "Er loopt al een verwerking, even geduld.")
            return
        pdfs = core.collect_pdfs(paths)
        if not pdfs:
            messagebox.showwarning(
                APP_TITLE,
                "Geen PDF-bestanden gevonden.\n\n"
                "Sleep de map met de deliberatie-PDF's op het venster,\n"
                "of kies de PDF's via de knoppen.",
            )
            return

        self.busy = True
        self.result_path = None
        self.btn_open.configure(state="disabled")
        self.btn_reveal.configure(state="disabled")
        self.status.configure(text=f"Bezig… {len(pdfs)} PDF's verwerken",
                              fg=WARN_COLOR)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self.log(f"{len(pdfs)} PDF-bestanden gevonden.")
        self.log("")

        t = threading.Thread(target=self._worker, args=(pdfs,), daemon=True)
        t.start()

    def _worker(self, pdfs):
        try:
            result = core.generate(pdfs, log=self.log)
            self.msg_queue.put(("done", result))
        except Exception as e:
            traceback.print_exc()
            self.msg_queue.put(("error", str(e)))

    def _on_done(self, result):
        self.busy = False
        self.result_path = result["out_path"]
        self.btn_open.configure(state="normal")
        self.btn_reveal.configure(state="normal")
        if result["n_bad"]:
            self.status.configure(
                text=(f"Klaar: {result['n_students']} studenten - "
                      f"{result['n_bad']} te controleren (zie tabblad 'Controle')"),
                fg=WARN_COLOR,
            )
        else:
            self.status.configure(
                text=(f"Klaar: {result['n_students']} studenten, "
                      "alle controles OK"),
                fg=OK_COLOR,
            )

    def _on_error(self, msg):
        self.busy = False
        self.status.configure(text="Er ging iets mis.", fg=ERR_COLOR)
        self.log("")
        self.log(f"FOUT: {msg}")
        messagebox.showerror(APP_TITLE, f"Er ging iets mis:\n\n{msg}")

    # ------------------------------------------------------------- output
    def open_result(self):
        if self.result_path and os.path.exists(self.result_path):
            open_path(self.result_path)

    def reveal_result(self):
        if self.result_path and os.path.exists(self.result_path):
            reveal_in_folder(self.result_path)


def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
