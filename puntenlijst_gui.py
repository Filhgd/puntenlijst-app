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
import math
import wave
import queue
import struct
import random
import tempfile
import threading
import subprocess
import traceback

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import puntenlijst_core as core
from version import __version__, APP_NAME, DEVELOPER, DEVELOPER_EMAIL, ORGANISATION

# Drag & drop is optioneel: zonder tkinterdnd2 blijven de knoppen werken.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    HAS_DND = False


APP_TITLE = APP_NAME


def resource_path(rel):
    """Pad naar een meegeleverd bestand, ook binnen een PyInstaller-app."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)
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


# ---------------------------------------------------------------------------
# Geluid: korte retro (MIDI-achtige) deuntjes, ter plekke gesynthetiseerd.
# Geen extra bibliotheken nodig; afspelen via afplay (Mac) / winsound (Windows).
# ---------------------------------------------------------------------------
SUCCESS_NOTES = [(523, 0.11), (659, 0.11), (784, 0.11), (1047, 0.32)]  # C-E-G-C
ERROR_NOTES = [(311, 0.16), (233, 0.16), (155, 0.42)]                  # dalend


def _make_wav(path, notes):
    """Schrijf een klein WAV-bestand met een blokgolf-melodietje (chiptune)."""
    rate = 22050
    frames = bytearray()
    for freq, dur in notes:
        n = int(rate * dur)
        for i in range(n):
            t = i / rate
            v = 1.0 if math.sin(2 * math.pi * freq * t) >= 0 else -1.0
            attack = min(1.0, i / (0.02 * n + 1))          # zachte start
            release = min(1.0, (n - i) / (0.35 * n + 1))   # uitsterven
            frames += struct.pack("<h", int(v * attack * release * 11000))
    with wave.open(path, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(rate)
        w.writeframes(bytes(frames))


def play_sound(kind):
    """Speel het succes- of foutdeuntje af (asynchroon, faalt stil)."""
    try:
        path = os.path.join(tempfile.gettempdir(), f"puntenlijst_{kind}.wav")
        if not os.path.exists(path):
            _make_wav(path, SUCCESS_NOTES if kind == "ok" else ERROR_NOTES)
        if sys.platform == "darwin":
            subprocess.Popen(["afplay", path])
        elif os.name == "nt":
            import winsound
            winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        else:
            subprocess.Popen(["aplay", "-q", path])
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Vuurwerk: overlay-animatie op een donker canvas over het hele venster.
# ---------------------------------------------------------------------------
class Fireworks:
    COLORS = ["#FFD700", "#FF6B6B", "#4ECDC4", "#95E1D3", "#F38181",
              "#AA96DA", "#FCE38A", "#7EC8E3", "#FFFFFF"]
    DURATION_MS = 4500

    def __init__(self, root, title, subtitle):
        self.root = root
        self.alive = True
        self.parts = []
        self.canvas = tk.Canvas(root, bg="#1B1B2F", highlightthickness=0)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.canvas.bind("<Button-1>", lambda e: self.stop())
        root.update_idletasks()
        self.w = max(root.winfo_width(), 400)
        self.h = max(root.winfo_height(), 300)
        self.canvas.create_text(self.w / 2, self.h * 0.42, text=title,
                                font=("Helvetica", 30, "bold"), fill="#FFD700")
        self.canvas.create_text(self.w / 2, self.h * 0.53, text=subtitle,
                                font=("Helvetica", 14), fill="white",
                                justify="center")
        self.canvas.create_text(self.w / 2, self.h * 0.94,
                                text="(klik om te sluiten)",
                                font=("Helvetica", 10), fill="#777799")
        for delay in (0, 350, 700, 1100, 1500, 1900, 2400):
            root.after(delay, self._burst)
        root.after(self.DURATION_MS, self.stop)
        self._tick()

    def _burst(self):
        if not self.alive:
            return
        cx = random.uniform(0.15, 0.85) * self.w
        cy = random.uniform(0.10, 0.45) * self.h
        color = random.choice(self.COLORS)
        for _ in range(26):
            angle = random.uniform(0, 2 * math.pi)
            speed = random.uniform(1.5, 5.5)
            item = self.canvas.create_oval(cx - 2, cy - 2, cx + 2, cy + 2,
                                           fill=color, outline="")
            self.parts.append({
                "id": item,
                "vx": math.cos(angle) * speed,
                "vy": math.sin(angle) * speed,
                "life": random.randint(22, 40),
            })

    def _tick(self):
        if not self.alive:
            return
        dead = []
        for p in self.parts:
            p["vy"] += 0.12                      # zwaartekracht
            self.canvas.move(p["id"], p["vx"], p["vy"])
            p["life"] -= 1
            if p["life"] == 8:
                self.canvas.itemconfigure(p["id"], fill="#555577")
            if p["life"] <= 0:
                self.canvas.delete(p["id"])
                dead.append(p)
        for p in dead:
            self.parts.remove(p)
        self.root.after(30, self._tick)

    def stop(self):
        if self.alive:
            self.alive = False
            self.canvas.destroy()


def shake_window(root):
    """Schud het venster kort heen en weer (fout-effect)."""
    try:
        root.update_idletasks()
        x, y = root.winfo_x(), root.winfo_y()
        seq = [14, -14, 11, -11, 8, -8, 5, -5, 2, -2, 0]

        def step(i=0):
            if i < len(seq):
                root.geometry(f"+{x + seq[i]}+{y}")
                root.after(35, step, i + 1)
        step()
    except Exception:
        pass


class App:
    def __init__(self, root):
        self.root = root
        self.busy = False
        self.result_path = None
        self.msg_queue = queue.Queue()

        root.title(f"{APP_TITLE} v{__version__}")
        root.geometry("720x620")
        root.minsize(560, 500)
        root.configure(bg=BG)

        # ---- logo faculteit (indien meegeleverd) ----
        self.logo_img = None
        try:
            logo_file = resource_path(os.path.join("assets", "logo_faculteit.png"))
            if os.path.exists(logo_file):
                self.logo_img = tk.PhotoImage(file=logo_file)
                tk.Label(root, image=self.logo_img, bg=BG).pack(pady=(14, 0))
        except Exception:
            self.logo_img = None

        # ---- kop ----
        tk.Label(
            root, text=APP_TITLE, font=("Helvetica", 20, "bold"),
            bg=BG, fg=ACCENT,
        ).pack(pady=(10 if self.logo_img else 16, 2))
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
        self.log_box.pack(fill="both", expand=True, padx=24, pady=(6, 4))

        # ---- voettekst: versie en ontwikkelaar ----
        tk.Label(
            root,
            text=(f"v{__version__}  ·  Ontwikkeld door {DEVELOPER} "
                  f"({DEVELOPER_EMAIL})\n{ORGANISATION}"),
            font=("Helvetica", 10), bg=BG, fg="#888888", justify="center",
        ).pack(pady=(0, 10))

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
            subtitle = (f"{result['n_students']} studenten verwerkt\n"
                        f"Let op: {result['n_bad']} te controleren "
                        f"(tabblad 'Controle')")
        else:
            self.status.configure(
                text=(f"Klaar: {result['n_students']} studenten, "
                      "alle controles OK"),
                fg=OK_COLOR,
            )
            subtitle = (f"{result['n_students']} studenten verwerkt - "
                        "alle controles OK")
        play_sound("ok")
        Fireworks(self.root, "GELUKT!", subtitle)

    def _on_error(self, msg):
        self.busy = False
        self.status.configure(text="Er ging iets mis.", fg=ERR_COLOR)
        self.log("")
        self.log(f"FOUT: {msg}")
        play_sound("fout")
        shake_window(self.root)
        self.root.after(450, lambda: messagebox.showerror(
            APP_TITLE, f"Er ging iets mis:\n\n{msg}"))

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
