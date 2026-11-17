"""Vietnamese desktop interface for four local file conversions."""

import os
from pathlib import Path
from queue import Empty, Queue
import sys
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from explorer_menu import IMAGE_FORMATS, IMAGE_SUFFIXES, WORD_SUFFIXES, choose_output, convert_from_explorer
from branding import CYAN, SURFACE, TEXT, apply_theme, brand_header


MODES = {
    "Word → PDF": ("word_pdf", "Tệp Word", " ".join(f"*{suffix}" for suffix in WORD_SUFFIXES), ".pdf"),
    "PDF → Word": ("pdf_word", "Tệp PDF", "*.pdf", ".docx"),
    "Ảnh → PDF": ("images_pdf", "Tệp ảnh", " ".join(f"*{suffix}" for suffix in IMAGE_SUFFIXES), ".pdf"),
    "PDF → Ảnh": ("pdf_images", "Tệp PDF", "*.pdf", ""),
}


class ConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.sources = []
        self.result = None
        self.events = Queue()
        self.busy = False

        apply_theme(root)
        root.geometry("720x570")
        root.minsize(620, 530)
        root.protocol("WM_DELETE_WINDOW", self.close)

        body = ttk.Frame(root, padding=24)
        body.pack(fill="both", expand=True)
        brand_header(body)
        ttk.Label(body, text="Chuyển đổi tài liệu", font=("Segoe UI", 23, "bold")).pack(anchor="w")
        ttk.Label(body, text="Word, PDF và ảnh", style="Muted.TLabel").pack(anchor="w", pady=(4, 18))

        self.mode = ttk.Combobox(body, values=list(MODES), state="readonly")
        self.mode.current(0)
        self.mode.pack(fill="x")
        self.mode.bind("<<ComboboxSelected>>", self.reset)

        toolbar = ttk.Frame(body)
        toolbar.pack(fill="x", pady=(10, 5))
        self.choose_button = ttk.Button(toolbar, text="+  Chọn tệp…", command=self.choose)
        self.choose_button.pack(side="left")
        self.order_buttons = [
            ttk.Button(toolbar, text="Lên", command=lambda: self.move(-1)),
            ttk.Button(toolbar, text="Xuống", command=lambda: self.move(1)),
            ttk.Button(toolbar, text="Bỏ ảnh", command=self.remove),
        ]
        for button in self.order_buttons:
            button.pack(side="left", padx=(5, 0))

        file_list = ttk.Frame(body)
        file_list.pack(fill="both", expand=True)
        scrollbar = ttk.Scrollbar(file_list, orient="vertical")
        scrollbar.pack(side="right", fill="y")
        self.listbox = tk.Listbox(
            file_list, height=5, yscrollcommand=scrollbar.set, exportselection=False,
            background=SURFACE, foreground=TEXT, selectbackground="#315A80", selectforeground=TEXT,
            relief="flat", borderwidth=0, highlightthickness=1, highlightbackground="#314761",
            highlightcolor=CYAN, font=("Segoe UI", 10), activestyle="none",
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self.refresh_controls)
        scrollbar.configure(command=self.listbox.yview)

        self.image_options = ttk.Frame(body)
        ttk.Label(self.image_options, text="Định dạng ảnh").pack(side="left")
        self.image_format = ttk.Combobox(self.image_options, values=IMAGE_FORMATS, state="readonly", width=8)
        self.image_format.current(0)
        self.image_format.pack(side="left", padx=(6, 14))
        ttk.Label(self.image_options, text="DPI").pack(side="left")
        self.dpi = ttk.Combobox(self.image_options, values=("72", "150", "200", "300", "600"), state="readonly", width=5)
        self.dpi.current(3)
        self.dpi.pack(side="left", padx=6)

        self.actions = ttk.Frame(body)
        self.actions.pack(fill="x", pady=(12, 6))
        self.run_button = ttk.Button(self.actions, text="Chuyển đổi…", command=self.start, style="Accent.TButton")
        self.run_button.pack(side="left")
        self.open_button = ttk.Button(self.actions, text="Mở kết quả", command=self.open_result)
        self.open_button.pack(side="left", padx=8)
        self.status = tk.StringVar()
        ttk.Label(body, textvariable=self.status, style="Muted.TLabel", wraplength=610).pack(anchor="w", pady=(6, 0))
        self.reset()

    def reset(self, _event=None):
        self.sources.clear()
        self.result = None
        self.show_sources()
        mode = MODES[self.mode.get()][0]
        self.image_options.pack_forget()
        if mode == "pdf_images":
            self.image_options.pack(fill="x", pady=(8, 0), before=self.actions)
        hints = {
            "word_pdf": "Chọn tệp DOC hoặc DOCX. Cần Microsoft Word.",
            "pdf_word": "Chọn tệp PDF. Kết quả có thể khác bố cục gốc.",
            "images_pdf": "Chọn một hoặc nhiều ảnh. Thứ tự trong danh sách là thứ tự trang.",
            "pdf_images": "Mỗi trang PDF được lưu thành một ảnh riêng.",
        }
        self.status.set(hints[mode])

    def refresh_controls(self, _event=None):
        is_images = MODES[self.mode.get()][0] == "images_pdf"
        selected = self.listbox.curselection()
        index = selected[0] if selected else -1
        can_move_up = index > 0
        can_move_down = 0 <= index < len(self.sources) - 1
        for button, enabled in zip(self.order_buttons, (can_move_up, can_move_down, index >= 0)):
            button.configure(state="normal" if enabled and is_images and not self.busy else "disabled")
        self.choose_button.configure(state="disabled" if self.busy else "normal")
        self.run_button.configure(state="normal" if self.sources and not self.busy else "disabled")
        self.open_button.configure(state="normal" if self.result and not self.busy else "disabled")
        for dropdown in (self.mode, self.image_format, self.dpi):
            dropdown.configure(state="disabled" if self.busy else "readonly")

    def show_sources(self, index=None):
        self.listbox.delete(0, "end")
        for path in self.sources:
            self.listbox.insert("end", str(path))
        if index is not None and self.sources:
            index = min(index, len(self.sources) - 1)
            self.listbox.selection_set(index)
            self.listbox.see(index)
        self.refresh_controls()

    def choose(self):
        mode, label, pattern, _ = MODES[self.mode.get()]
        options = {"parent": self.root, "filetypes": [(label, pattern)]}
        if self.sources:
            options["initialdir"] = self.sources[-1].parent
        if mode == "images_pdf":
            paths = filedialog.askopenfilenames(**options)
            if not paths:
                return
            index = len(self.sources)
            self.sources.extend(Path(path) for path in paths)
        else:
            path = filedialog.askopenfilename(**options)
            if not path:
                return
            index = 0
            self.sources = [Path(path)]
        self.result = None
        self.status.set(f"Đã chọn {len(self.sources)} tệp.")
        self.show_sources(index)

    def move(self, delta):
        index = self.listbox.curselection()[0]
        target = index + delta
        self.sources[index], self.sources[target] = self.sources[target], self.sources[index]
        self.result = None
        self.status.set("Đã đổi thứ tự ảnh.")
        self.show_sources(target)

    def remove(self):
        index = self.listbox.curselection()[0]
        self.sources.pop(index)
        self.result = None
        self.status.set(f"Đã chọn {len(self.sources)} tệp.")
        self.show_sources(index)

    def start(self):
        mode, _, _, suffix = MODES[self.mode.get()]
        sources = self.sources.copy()
        first = sources[0]
        output = choose_output(self.root, first, suffix, "anh-da-ghep" if len(sources) > 1 else None)
        if output is not None:
            self.run_conversion(mode, sources, output)

    def run_conversion(self, mode, sources, output):
        self.busy = True
        self.result = None
        self.status.set("Đang chuyển đổi…")
        self.refresh_controls()
        Thread(
            target=self.convert,
            args=(mode, sources, output, self.image_format.get(), int(self.dpi.get())),
            daemon=True,
        ).start()
        self.root.after(80, self.poll)

    def convert(self, mode, sources, output, image_format, dpi):
        def progress(page, total):
            self.events.put(("progress", f"Đang xử lý trang {page}/{total}…"))

        try:
            from converter import convert_to_file, pdf_to_images

            if mode == "pdf_images":
                result = pdf_to_images(sources[0], output, image_format, dpi, progress)
            else:
                result = convert_to_file(mode, sources, output)
        except Exception as exc:
            self.events.put(("error", str(exc)))
        else:
            self.events.put(("done", result))

    def poll(self):
        while self.busy:
            try:
                kind, value = self.events.get_nowait()
            except Empty:
                self.root.after(80, self.poll)
                return
            if kind == "progress":
                self.status.set(value)
                continue
            self.busy = False
            if kind == "done":
                self.result = value
                self.status.set(f"Đã lưu {value.name}")
            else:
                self.status.set("Chuyển đổi thất bại.")
                messagebox.showerror("Không thể chuyển đổi", value, parent=self.root)
            self.refresh_controls()

    def open_result(self):
        try:
            os.startfile(self.result)
        except OSError as exc:
            messagebox.showerror("Không thể mở kết quả", str(exc), parent=self.root)

    def close(self):
        if self.busy:
            messagebox.showinfo("Đang chuyển đổi", "Hãy đợi chuyển đổi hoàn tất rồi đóng ứng dụng.", parent=self.root)
        else:
            self.root.destroy()


if __name__ == "__main__":
    if len(sys.argv) == 4 and sys.argv[1] == "--convert":
        sys.exit(convert_from_explorer(sys.argv[2], Path(sys.argv[3])))
    window = tk.Tk()
    ConverterApp(window)
    window.mainloop()
