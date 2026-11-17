"""Compact conversion workflow for the File Explorer entry point."""

from pathlib import Path
from queue import Empty, Queue
from threading import Event, Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from branding import apply_theme, brand_header

WORD_SUFFIXES = (".doc", ".docx")
IMAGE_SUFFIXES = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".gif", ".jp2", ".avif")
IMAGE_FORMATS = ("PNG", "JPEG", "TIFF", "BMP", "WEBP")

def conversion_mode(source_suffix: str, target: str) -> str:
    if target == "pdf":
        if source_suffix in WORD_SUFFIXES:
            return "word_pdf"
        if source_suffix in IMAGE_SUFFIXES:
            return "images_pdf"
    if source_suffix == ".pdf":
        if target == "docx":
            return "pdf_word"
        if target.upper() in IMAGE_FORMATS:
            return "pdf_images"
    raise ValueError("Định dạng đầu vào hoặc đầu ra không được hỗ trợ.")


def choose_output(window, source: Path, suffix: str, name: str | None = None) -> Path | None:
    options = {"parent": window, "initialdir": str(source.parent)}
    if suffix:
        selected = filedialog.asksaveasfilename(
            **options, title="Lưu tệp chuyển đổi", defaultextension=suffix,
            filetypes=[(suffix[1:].upper(), f"*{suffix}")],
            initialfile=f"{name or source.stem}{suffix}", confirmoverwrite=True,
        )
        if selected and Path(selected).suffix.lower() != suffix:
            messagebox.showerror("Sai định dạng", f"Tên tệp đầu ra phải kết thúc bằng {suffix}.", parent=window)
            return None
    else:
        selected = filedialog.askdirectory(**options, title="Chọn nơi lưu thư mục ảnh", mustexist=True)
    return Path(selected) if selected else None


def convert_from_explorer(target: str, source: Path) -> int:
    window = tk.Tk()
    window.withdraw()
    apply_theme(window)
    try:
        source = source.absolute()
        if not source.is_file():
            raise ValueError("Không tìm thấy tệp đã chọn.")
        mode = conversion_mode(source.suffix.lower(), target)
    except (OSError, ValueError) as exc:
        messagebox.showerror("Không thể chuyển đổi", str(exc), parent=window)
        window.destroy()
        return 1

    events = Queue()
    output = None
    output_ready = Event()

    def convert():
        try:
            from converter import convert_to_file, pdf_to_images, word_application

            if mode == "word_pdf":
                # The COM instance is created, used and closed on the same worker thread.
                with word_application() as word:
                    output_ready.wait()
                    if output is None:
                        return
                    convert_to_file(mode, [source], output, word=word)
            else:
                output_ready.wait()
                if output is None:
                    return
                if mode == "pdf_images":
                    pdf_to_images(source, output, target.upper(), 300,
                                  lambda page, total: events.put(("page", (page, total))))
                else:
                    convert_to_file(mode, [source], output)
        except Exception as exc:
            events.put(("error", str(exc)))
        else:
            events.put(("done", None))

    # A non-daemon worker also closes Word if the user cancels the save dialog.
    worker = Thread(target=convert)

    # Give the native picker a mapped topmost owner even when Explorer launched us in the background.
    window.attributes("-alpha", 0.0)
    window.attributes("-topmost", True)
    window.deiconify()
    window.update_idletasks()
    window.focus_force()
    try:
        worker.start()
        output = choose_output(window, source, "" if mode == "pdf_images" else f".{target}")
    finally:
        window.attributes("-topmost", False)
        output_ready.set()
    if output is None:
        window.destroy()
        return 0

    window.resizable(False, False)
    window.protocol("WM_DELETE_WINDOW", window.iconify)
    body = ttk.Frame(window, padding=18)
    body.pack(fill="both", expand=True)
    brand_header(body, compact=True)
    label = ttk.Label(body, text="Đang chuyển đổi…", wraplength=330)
    label.pack(anchor="w", pady=(0, 8))
    bar = ttk.Progressbar(body, length=330, mode="indeterminate")
    bar.pack(fill="x")
    window.update_idletasks()
    window.attributes("-alpha", 1.0)
    bar.start(15)

    exit_code = 0

    def poll():
        nonlocal exit_code
        while True:
            try:
                kind, value = events.get_nowait()
            except Empty:
                window.after(70, poll)
                return
            bar.stop()
            if kind == "page":
                page, total = value
                label.configure(text=f"Đang xuất ảnh trang {page}/{total}…")
                bar.configure(mode="determinate", maximum=total, value=page - 1)
                continue
            if kind == "error":
                exit_code = 1
                messagebox.showerror("Không thể chuyển đổi", value, parent=window)
                window.destroy()
            else:
                label.configure(text="Đã lưu kết quả")
                bar.configure(mode="determinate", maximum=100, value=100)
                window.after(400, window.destroy)
            return

    window.after(70, poll)
    window.mainloop()
    return exit_code
