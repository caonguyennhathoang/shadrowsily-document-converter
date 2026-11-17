"""Shared Shadrowsily visuals for the desktop and Explorer windows."""

from pathlib import Path
from ctypes import windll
import tkinter as tk
from tkinter import ttk


APP_NAME = "Shadrovert"
ASSETS = Path(__file__).resolve().parent / "assets"
NAVY = "#0F1C2D"
SURFACE = "#182A40"
CYAN = "#75D5EC"
TEXT = "#E0EBF7"
MUTED = "#A4B7CD"


def apply_theme(window):
    windll.shell32.SetCurrentProcessExplicitAppUserModelID("Shadrowsily.DocumentConverter")
    window.title(APP_NAME)
    window.configure(background=NAVY)
    window.iconbitmap(default=str(ASSETS / "app.ico"))
    style = ttk.Style(window)
    style.theme_use("clam")
    style.configure(".", background=NAVY, foreground=TEXT, font=("Segoe UI", 10))
    style.configure("TFrame", background=NAVY)
    style.configure("TLabel", background=NAVY, foreground=TEXT)
    style.configure("Muted.TLabel", foreground=MUTED)
    style.configure("TButton", background=SURFACE, foreground=TEXT, borderwidth=0, padding=(12, 8))
    style.map("TButton", background=[("active", "#27425E")], foreground=[("disabled", "#71839A")])
    style.configure("Accent.TButton", background=CYAN, foreground=NAVY)
    style.map("Accent.TButton", background=[("disabled", SURFACE), ("active", "#A4E8F6")],
              foreground=[("disabled", "#71839A"), ("active", NAVY)])
    style.configure("TCombobox", fieldbackground=SURFACE, background=SURFACE,
                    foreground=TEXT, arrowcolor=CYAN, bordercolor="#314761", padding=7)
    style.map("TCombobox", fieldbackground=[("readonly", SURFACE)],
              foreground=[("readonly", TEXT), ("disabled", "#71839A")])
    style.configure("TProgressbar", background=CYAN, troughcolor=SURFACE, borderwidth=0, thickness=6)
    style.configure("Vertical.TScrollbar", background="#314761", troughcolor=SURFACE, borderwidth=0)
    window.option_add("*TCombobox*Listbox.background", SURFACE)
    window.option_add("*TCombobox*Listbox.foreground", TEXT)
    window.option_add("*TCombobox*Listbox.selectBackground", "#315A80")


def brand_header(parent, compact=False):
    header = ttk.Frame(parent)
    header.pack(fill="x", pady=(0, 18 if not compact else 10))
    logo = tk.PhotoImage(file=str(ASSETS / ("brand-compact.png" if compact else "brand-header.png")))
    label = ttk.Label(header, image=logo)
    label.image = logo
    label.pack(side="left")
    if not compact:
        ttk.Label(header, text="Xử lý trên máy", style="Muted.TLabel").pack(side="right")
    return header
