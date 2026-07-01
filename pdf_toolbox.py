#!/usr/bin/env python3
"""PDF 工具箱"""

import ctypes
import os
import sys

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES

# ── 色票 ──────────────────────────────────────────────────────────────────────
BG        = "#f1f5f9"
SURFACE   = "#ffffff"
PRIMARY   = "#2563eb"
DANGER    = "#ef4444"
TEXT1     = "#1e293b"
TEXT2     = "#64748b"
BORDER    = "#e2e8f0"
AMBER     = "#d97706"
HOVER     = "#dbeafe"
WORD_BLUE = "#1d4ed8"


def ensure_pypdf() -> bool:
    try:
        import pypdf  # noqa: F401
        return True
    except ImportError:
        return False


def ensure_win32com() -> bool:
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def ensure_pymupdf() -> bool:
    try:
        import fitz  # noqa: F401
        return True
    except ImportError:
        return False


def ensure_pdf2docx() -> bool:
    try:
        import pdf2docx  # noqa: F401
        return True
    except ImportError:
        return False


def pdf_to_jpgs(pdf_path: str, out_dir: str, dpi: int = 300,
                on_page=None) -> list[str]:
    """把 PDF 每頁轉成 JPG，on_page(current, total) 用於進度回呼。"""
    import fitz
    doc = fitz.open(pdf_path)
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    total = len(doc)
    paths = []
    for i, page in enumerate(doc, 1):
        pix = page.get_pixmap(matrix=mat, alpha=False)
        out_path = os.path.join(out_dir, f"{base}_p{i:03d}.jpg")
        pix.save(out_path)
        paths.append(out_path)
        if on_page:
            on_page(i, total)
    doc.close()
    return paths


def docx_to_pdf(docx_path: str, pdf_path: str) -> None:
    """用 Word COM 把 docx 轉 PDF，路徑必須是絕對路徑。"""
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(docx_path)
        doc.SaveAs2(pdf_path, FileFormat=17)   # 17 = wdFormatPDF
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


WORD_EXTS = {".doc", ".docx"}
IMG_EXTS  = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}


def imgs_to_pdf(img_paths: list[str], out_path: str) -> None:
    """把多張圖片依序合成一份 PDF（Pillow）。"""
    from PIL import Image
    pages: list[Image.Image] = []
    for p in img_paths:
        img = Image.open(p)
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGB")
        elif img.mode == "RGBA":
            bg = Image.new("RGB", img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[3])
            img = bg
        pages.append(img)
    if not pages:
        raise ValueError("沒有可用的圖片")
    pages[0].save(out_path, save_all=True, append_images=pages[1:])


class SplashScreen(tk.Toplevel):
    W, H = 480, 240

    def __init__(self, master):
        super().__init__(master)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=SURFACE)
        self.config(highlightbackground=BORDER, highlightthickness=1)

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{self.W}x{self.H}+{(sw - self.W) // 2}+{(sh - self.H) // 2}")

        self._img_ref = None
        img_path = self._find_asset("splash.png")
        if img_path:
            try:
                raw = tk.PhotoImage(file=img_path)
                sx = max(1, raw.width() // 200)
                sy = max(1, raw.height() // 80)
                factor = max(sx, sy)
                self._img_ref = raw.subsample(factor) if factor > 1 else raw
                tk.Label(self, image=self._img_ref, bg=SURFACE).pack(pady=(20, 6))
            except Exception:
                self._draw_title()
        else:
            self._draw_title()

        self._status_var = tk.StringVar(value="載入中…")
        tk.Label(self, textvariable=self._status_var, bg=SURFACE, fg=TEXT2,
                 font=("Segoe UI", 9)).pack()

        self._pvar = tk.DoubleVar(value=0)
        ttk.Progressbar(
            self, variable=self._pvar, maximum=100,
            mode="determinate", length=self.W - 80,
            style="Horizontal.TProgressbar",
        ).pack(pady=(10, 0))

        self.update_idletasks()
        # Tkinter splash 就緒後立刻關掉 PyInstaller 的 bootloader splash
        try:
            import pyi_splash  # type: ignore
            pyi_splash.close()
        except ImportError:
            pass

    def _draw_title(self):
        tk.Label(self, text="PDF 工具箱", bg=SURFACE, fg=TEXT1,
                 font=("Segoe UI", 18, "bold")).pack(pady=(20, 6))

    @staticmethod
    def _find_asset(name: str) -> str | None:
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(base, name)
        return path if os.path.exists(path) else None

    def set_progress(self, pct: float, text: str = ""):
        self._pvar.set(pct)
        if text:
            self._status_var.set(text)
        self.update_idletasks()


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.title("PDF 工具箱")
        self.geometry("800x640")
        self.resizable(True, True)
        self.configure(bg=BG)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.files: list[dict] = []

        self._setup_styles()
        splash = SplashScreen(self)
        splash.set_progress(60, "建立介面…")
        self._build_ui()
        splash.set_progress(100, "完成！")
        self.after(300, lambda: (splash.destroy(), self._show_main()))

    def _show_main(self):
        self.deiconify()
        self.lift()
        self.focus_force()
        self._show_disclaimer()

    def _show_disclaimer(self):
        dlg = tk.Toplevel(self)
        dlg.title("使用聲明")
        dlg.resizable(False, False)
        dlg.configure(bg=SURFACE)
        dlg.attributes("-topmost", True)

        W, H = 520, 280
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        dlg.geometry(f"{W}x{H}+{(sw - W) // 2}+{(sh - H) // 2}")

        dlg.grab_set()
        dlg.focus_set()
        dlg.protocol("WM_DELETE_WINDOW", self.destroy)

        pad = ttk.Frame(dlg, padding=(28, 22, 28, 20))
        pad.pack(fill="both", expand=True)

        ttk.Label(pad, text="⚠  使用聲明",
                  font=("Segoe UI", 13, "bold"),
                  foreground=TEXT1, background=SURFACE).pack(anchor="w")
        ttk.Separator(pad, orient="horizontal").pack(fill="x", pady=(10, 16))

        tk.Label(pad,
                 text="本軟體禁止台北市私立立人國際中小學使用，\n除非取得著作權人同意。",
                 bg=SURFACE, fg=DANGER,
                 font=("Segoe UI", 12, "bold"),
                 justify="left", wraplength=460).pack(anchor="w")

        tk.Label(pad,
                 text="© 2026 陳冠廷．著作權所有，未經授權不得於特定機構使用。",
                 bg=SURFACE, fg=TEXT2,
                 font=("Segoe UI", 9),
                 justify="left").pack(anchor="w", pady=(10, 0))

        btn_row = ttk.Frame(pad)
        btn_row.pack(fill="x", pady=(20, 0))
        ttk.Button(btn_row, text="關閉程式",
                   command=self.destroy,
                   style="Danger.TButton").pack(side="right", padx=(8, 0))
        ttk.Button(btn_row, text="我了解，繼續使用",
                   command=dlg.destroy,
                   style="Primary.TButton").pack(side="right")

        dlg.wait_window()

    # ── 樣式 ──────────────────────────────────────────────────────────────────

    def _setup_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", background=BG, foreground=TEXT1, font=("Segoe UI", 10))
        s.configure("TFrame",  background=BG)
        s.configure("TLabel",  background=BG, foreground=TEXT1)
        s.configure("Title.TLabel", background=BG, foreground=TEXT1,
                    font=("Segoe UI", 15, "bold"))
        s.configure("Sub.TLabel",   background=BG, foreground=TEXT2,
                    font=("Segoe UI", 9))
        s.configure("Info.TLabel",  background=BG, foreground=TEXT2,
                    font=("Segoe UI", 9))

        for name, bg, fg, hover_bg in [
            ("Primary", PRIMARY, "white", "#1d4ed8"),
            ("Danger",  DANGER,  "white", "#dc2626"),
            ("Muted",   BORDER,  TEXT1,   "#cbd5e1"),
            ("Ghost",   SURFACE, TEXT1,   HOVER),
        ]:
            s.configure(f"{name}.TButton",
                        background=bg, foreground=fg,
                        font=("Segoe UI", 10, "bold" if name == "Primary" else "normal"),
                        padding=(12, 7), relief="flat", borderwidth=0)
            s.map(f"{name}.TButton", background=[("active", hover_bg)])

        s.configure("Treeview",
                    background=SURFACE, foreground=TEXT1,
                    fieldbackground=SURFACE, rowheight=38,
                    font=("Segoe UI", 10), borderwidth=0, relief="flat")
        s.configure("Treeview.Heading",
                    background=BG, foreground=TEXT2,
                    font=("Segoe UI", 9, "bold"), relief="flat", padding=(8, 7))
        s.map("Treeview",
              background=[("selected", HOVER)],
              foreground=[("selected", TEXT1)])
        s.configure("TCheckbutton", background=BG, foreground=TEXT1,
                    font=("Segoe UI", 10))
        s.configure("TSeparator", background=BORDER)
        s.configure("Green.TButton",
                    background="#059669", foreground="white",
                    font=("Segoe UI", 10), padding=(12, 7),
                    relief="flat", borderwidth=0)
        s.map("Green.TButton", background=[("active", "#047857")])
        s.configure("WordBlue.TButton",
                    background=WORD_BLUE, foreground="white",
                    font=("Segoe UI", 10), padding=(12, 7),
                    relief="flat", borderwidth=0)
        s.map("WordBlue.TButton", background=[("active", "#1e3a8a")])
        s.configure("Horizontal.TProgressbar",
                    troughcolor=BORDER, background=PRIMARY,
                    thickness=10, borderwidth=0, relief="flat")

    # ── UI 建構 ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        wrap = ttk.Frame(self, padding=(20, 16, 20, 18))
        wrap.grid(sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(1, weight=1)

        # Header
        hdr = ttk.Frame(wrap)
        hdr.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        ttk.Label(hdr, text="PDF 工具箱", style="Title.TLabel").pack(side="left")
        ttk.Label(hdr, text="   合併 ／ 圖片→PDF ／ Word→PDF ／ PDF→JPG ／ PDF→Word ／ 分割",
                  style="Sub.TLabel").pack(side="left", pady=(6, 0))
        ttk.Label(hdr, text="© 2026 陳冠廷",
                  style="Sub.TLabel").pack(side="right", pady=(6, 0))

        # Content
        body = ttk.Frame(wrap)
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        card = tk.Frame(body, bg=SURFACE,
                        highlightbackground=BORDER, highlightthickness=1)
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            card,
            columns=("no", "name", "pages", "status"),
            show="headings", selectmode="browse",
        )
        for col, text, w, anchor, stretch in [
            ("no",     "#",      44,  "center", False),
            ("name",   "檔案名稱", 400, "w",      True),
            ("pages",  "頁數",    62,  "center", False),
            ("status", "狀態",    110, "center", False),
        ]:
            self.tree.heading(col, text=text, anchor=anchor)
            self.tree.column(col, width=w, minwidth=w, stretch=stretch, anchor=anchor)

        self.tree.tag_configure("odd",  foreground=AMBER)
        self.tree.tag_configure("even", foreground=TEXT1)
        self.tree.tag_configure("word", foreground=WORD_BLUE)

        vsb = ttk.Scrollbar(card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self.empty = tk.Label(
            card,
            text="尚無檔案\n點擊「＋ 新增檔案」，或直接拖曳檔案到此區域\n支援 PDF、Word（.docx / .doc）",
            bg=SURFACE, fg="#94a3b8",
            font=("Segoe UI", 11), justify="center",
        )

        # 拖曳進入區域時高亮
        card.drop_target_register(DND_FILES)
        card.dnd_bind("<<Drop>>", self._on_drop)
        card.dnd_bind("<<DragEnter>>", lambda e: card.config(
            highlightbackground=PRIMARY, highlightthickness=2))
        card.dnd_bind("<<DragLeave>>", lambda e: card.config(
            highlightbackground=BORDER, highlightthickness=1))

        # Button column
        bc = ttk.Frame(body)
        bc.grid(row=0, column=1, sticky="n")

        def b(text, cmd, style, pady=(0, 6)):
            ttk.Button(bc, text=text, command=cmd,
                       style=style, width=11).pack(fill="x", pady=pady)

        b("＋ 新增檔案", self.add_files,      "Primary.TButton", (0, 8))
        b("✕ 移除選取", self.remove_selected, "Danger.TButton",  (0, 20))
        ttk.Separator(bc, orient="horizontal").pack(fill="x", pady=(0, 10))
        b("↑  上移",   self.move_up,   "Ghost.TButton", (0, 4))
        b("↓  下移",   self.move_down, "Ghost.TButton", (0, 18))
        ttk.Separator(bc, orient="horizontal").pack(fill="x", pady=(0, 10))
        b("清空",      self.clear_all, "Muted.TButton", (0, 0))

        # Options
        opt = ttk.Frame(wrap)
        opt.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        self.var_blank = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opt,
            text=" 單數頁結尾補空白頁（讓下一份 PDF 從正面起始，適合雙面列印）",
            variable=self.var_blank, command=self._refresh_status,
            bg=BG, fg=TEXT1, font=("Segoe UI", 10),
            activebackground=BG, cursor="hand2",
        ).pack(side="left")

        # 進度條
        prog_row = ttk.Frame(wrap)
        prog_row.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        prog_row.columnconfigure(0, weight=1)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(
            prog_row, variable=self.progress_var,
            maximum=100, mode="determinate",
        )
        self.progress.grid(row=0, column=0, sticky="ew")
        self.prog_label = ttk.Label(prog_row, text="", style="Info.TLabel")
        self.prog_label.grid(row=1, column=0, sticky="w", pady=(3, 0))

        # Status + 操作按鈕（兩行）
        bar = ttk.Frame(wrap)
        bar.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        bar.columnconfigure(0, weight=1)

        # 第一行：轉換工具
        conv_row = ttk.Frame(bar)
        conv_row.grid(row=0, column=0, sticky="e", pady=(0, 4))
        ttk.Label(conv_row, text="轉換：", style="Info.TLabel").pack(side="left", padx=(0, 6))
        ttk.Button(conv_row, text="Word → PDF",
                   command=self.convert_words,
                   style="Ghost.TButton").pack(side="left", padx=(0, 4))
        ttk.Button(conv_row, text="圖片 → PDF",
                   command=self.convert_imgs_to_pdf,
                   style="Ghost.TButton").pack(side="left", padx=(0, 4))
        ttk.Button(conv_row, text="PDF → JPG",
                   command=self.convert_to_jpg,
                   style="Green.TButton").pack(side="left", padx=(0, 4))
        ttk.Button(conv_row, text="PDF → Word",
                   command=self.convert_to_word,
                   style="WordBlue.TButton").pack(side="left", padx=(0, 0))

        # 第二行：主操作
        btn_row = ttk.Frame(bar)
        btn_row.grid(row=1, column=0, sticky="e", pady=(0, 4))
        ttk.Button(btn_row, text="PDF 分割",
                   command=self.split_pdf,
                   style="Muted.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(btn_row, text="  合併 PDF  →",
                   command=self.merge,
                   style="Primary.TButton").pack(side="left")

        self.status_var = tk.StringVar(value="尚未新增任何檔案")
        ttk.Label(bar, textvariable=self.status_var,
                  style="Info.TLabel").grid(row=1, column=0, sticky="w")

        self._show_empty()

    # ── 狀態刷新 ──────────────────────────────────────────────────────────────

    def _show_empty(self):
        if self.files:
            self.empty.place_forget()
        else:
            self.empty.place(relx=0.5, rely=0.5, anchor="center")

    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, f in enumerate(self.files, 1):
            if f.get("ftype") == "word":
                self.tree.insert("", "end",
                                 values=(i, f["name"], "—", "📝 Word（待轉換）"),
                                 tags=("word",))
            else:
                odd = f["pages"] % 2 == 1
                self.tree.insert("", "end",
                                 values=(i, f["name"], f["pages"],
                                         "⚠ 單數頁" if odd else "✓ 偶數頁"),
                                 tags=("odd" if odd else "even",))
        self._show_empty()
        self._refresh_status()

    def _refresh_status(self):
        if not self.files:
            self.status_var.set("尚未新增任何檔案")
            return
        pdf_files  = [f for f in self.files if f.get("ftype") != "word"]
        word_files = [f for f in self.files if f.get("ftype") == "word"]
        total_pdf  = sum(f["pages"] for f in pdf_files)
        n_odd      = sum(1 for f in pdf_files if f["pages"] % 2 == 1)
        extra      = n_odd if self.var_blank.get() else 0

        parts = [f"{len(self.files)} 個檔案"]
        if pdf_files:
            parts.append(f"PDF 共 {total_pdf} 頁")
        if word_files:
            parts.append(f"Word {len(word_files)} 份（頁數待轉換）")
        if n_odd:
            parts.append(f"{n_odd} 個單數頁")
        if extra:
            parts.append(f"補空白後 PDF 共 {total_pdf + extra} 頁")
        self.status_var.set("　".join(parts))

    def _set_progress(self, pct: float, detail: str = ""):
        self.progress_var.set(pct)
        self.prog_label.config(text=detail)
        self.update_idletasks()

    def _reset_progress(self):
        self.progress_var.set(0)
        self.prog_label.config(text="")

    # ── 選取 helpers ──────────────────────────────────────────────────────────

    def _sel_idx(self) -> int | None:
        sel = self.tree.selection()
        if not sel:
            return None
        return list(self.tree.get_children()).index(sel[0])

    def _set_sel(self, idx: int):
        kids = self.tree.get_children()
        if 0 <= idx < len(kids):
            self.tree.selection_set(kids[idx])
            self.tree.see(kids[idx])

    # ── 動作 ──────────────────────────────────────────────────────────────────

    @staticmethod
    def _parse_dnd_paths(data: str) -> list[str]:
        """解析 tkinterdnd2 回傳的路徑字串（含空格路徑用 {} 包住）。"""
        paths, i = [], 0
        data = data.strip()
        while i < len(data):
            if data[i] == "{":
                end = data.index("}", i)
                paths.append(data[i + 1:end])
                i = end + 2
            elif data[i] == " ":
                i += 1
            else:
                end = i
                while end < len(data) and data[end] not in (" ", "{"):
                    end += 1
                paths.append(data[i:end])
                i = end
        return [p for p in paths if p]

    def _on_drop(self, event):
        from pypdf import PdfReader
        from tkinter import messagebox
        paths = self._parse_dnd_paths(event.data)
        for p in paths:
            ext  = os.path.splitext(p)[1].lower()
            name = os.path.basename(p)
            if ext in WORD_EXTS:
                self.files.append({"path": p, "name": name, "pages": 0, "ftype": "word"})
            elif ext == ".pdf":
                try:
                    pages = len(PdfReader(p).pages)
                    self.files.append({"path": p, "name": name, "pages": pages, "ftype": "pdf"})
                except Exception as e:
                    messagebox.showerror("讀取失敗", f"{name}\n{e}")
            else:
                messagebox.showwarning("不支援的格式",
                    f"{name}\n拖曳區只接受 PDF / Word，圖片請用「圖片→PDF」按鈕。")
        # 拖曳結束恢復邊框顏色
        event.widget.config(highlightbackground=BORDER, highlightthickness=1)
        self._reset_progress()
        self._refresh_tree()

    def convert_imgs_to_pdf(self):
        """開對話框選圖片，合成一份 PDF 並儲存到使用者指定位置。"""
        paths = filedialog.askopenfilenames(
            title="選擇圖片（可多選，依選取順序排列）",
            filetypes=[
                ("圖片檔案", "*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.webp"),
                ("所有檔案", "*.*"),
            ],
        )
        if not paths:
            return
        out = filedialog.asksaveasfilename(
            title="儲存為 PDF",
            defaultextension=".pdf",
            filetypes=[("PDF 檔案", "*.pdf")],
            initialfile="圖片合併.pdf",
        )
        if not out:
            return
        total = len(paths)
        try:
            for i, p in enumerate(paths, 1):
                self._set_progress(i / total * 100, f"處理圖片 {i}/{total}：{os.path.basename(p)}")
            imgs_to_pdf(list(paths), out)
            self._set_progress(100, f"完成！已儲存 → {os.path.basename(out)}")
            messagebox.showinfo("完成", f"已將 {total} 張圖片合成 PDF：\n{out}")
        except Exception as e:
            messagebox.showerror("轉換失敗", str(e))
        finally:
            self._reset_progress()

    def add_files(self):
        from pypdf import PdfReader
        paths = filedialog.askopenfilenames(
            title="選擇要合併的檔案",
            filetypes=[
                ("支援格式", "*.pdf *.docx *.doc"),
                ("PDF 檔案", "*.pdf"),
                ("Word 文件", "*.docx *.doc"),
                ("所有檔案", "*.*"),
            ],
        )
        total = len(paths)
        for i, p in enumerate(paths, 1):
            ext  = os.path.splitext(p)[1].lower()
            name = os.path.basename(p)
            self._set_progress(i / total * 100 if total else 100,
                               f"讀取 {i}/{total}：{name}")
            if ext in WORD_EXTS:
                self.files.append({"path": p, "name": name, "pages": 0, "ftype": "word"})
            else:
                try:
                    pages = len(PdfReader(p).pages)
                    self.files.append({"path": p, "name": name, "pages": pages, "ftype": "pdf"})
                except Exception as e:
                    messagebox.showerror("讀取失敗", f"{name}\n{e}")
        self._reset_progress()
        self._refresh_tree()

    def remove_selected(self):
        idx = self._sel_idx()
        if idx is not None:
            self.files.pop(idx)
            self._refresh_tree()
            self._set_sel(min(idx, len(self.files) - 1))

    def move_up(self):
        idx = self._sel_idx()
        if idx is None or idx == 0:
            return
        self.files[idx - 1], self.files[idx] = self.files[idx], self.files[idx - 1]
        self._refresh_tree()
        self._set_sel(idx - 1)

    def move_down(self):
        idx = self._sel_idx()
        if idx is None or idx >= len(self.files) - 1:
            return
        self.files[idx], self.files[idx + 1] = self.files[idx + 1], self.files[idx]
        self._refresh_tree()
        self._set_sel(idx + 1)

    def clear_all(self):
        if not self.files:
            return
        if messagebox.askyesno("確認清空", "移除所有已新增的檔案？"):
            self.files.clear()
            self._refresh_tree()

    def convert_words(self):
        """將清單中所有 Word 檔轉成 PDF，存到原始資料夾，並更新清單。"""
        word_files = [f for f in self.files if f.get("ftype") == "word"]
        if not word_files:
            messagebox.showinfo("提示", "清單中沒有 Word 檔案")
            return
        if not ensure_win32com():
            messagebox.showerror(
                "缺少套件",
                "無法安裝 pywin32\n請手動執行：\n\npip install pywin32\n\n"
                "（需要已安裝 Microsoft Word）",
            )
            return

        from pypdf import PdfReader

        total  = len(word_files)
        errors = []
        for idx, f in enumerate(word_files):
            out_pdf = os.path.splitext(f["path"])[0] + ".pdf"
            try:
                self._set_progress(idx / total * 100,
                                   f"轉換 Word ({idx + 1}/{total})：{f['name']}")
                docx_to_pdf(os.path.abspath(f["path"]), os.path.abspath(out_pdf))
                pages = len(PdfReader(out_pdf).pages)
                f["path"]  = out_pdf
                f["name"]  = os.path.basename(out_pdf)
                f["pages"] = pages
                f["ftype"] = "pdf"
            except Exception as e:
                errors.append(f"{f['name']}：{e}")
        self._set_progress(100, f"✓ {total} 個 Word 檔案轉換完成")
        self._refresh_tree()

        if errors:
            messagebox.showerror("部分轉換失敗", "\n".join(errors))
        else:
            messagebox.showinfo(
                "轉換完成",
                f"✓ {len(word_files)} 個 Word 檔案已轉換為 PDF\n"
                f"存放位置：原始檔案所在資料夾"
            )
        self._reset_progress()

    def convert_to_jpg(self):
        """將清單中所有 PDF 每頁匯出為 JPG，存到使用者指定的資料夾。"""
        pdf_files = [f for f in self.files if f.get("ftype") == "pdf"]
        if not pdf_files:
            messagebox.showwarning("提示", "清單中沒有 PDF 檔案\n（Word 檔請先點「轉換 Word → PDF」）")
            return
        if not ensure_pymupdf():
            messagebox.showerror(
                "缺少套件",
                "無法安裝 pymupdf\n請手動執行：\n\npip install pymupdf",
            )
            return

        out_dir = filedialog.askdirectory(title="選擇 JPG 儲存資料夾")
        if not out_dir:
            return

        total_pages = sum(f["pages"] for f in pdf_files)
        done        = [0]
        errors      = []
        total_imgs  = 0

        for f in pdf_files:
            try:
                fname = f["name"]

                def on_page(cur, file_total, _n=fname, _d=done):
                    _d[0] += 1
                    pct = _d[0] / total_pages * 100 if total_pages > 0 else 100
                    self._set_progress(pct,
                        f"轉換第 {_d[0]}/{total_pages} 頁：{_n}")

                imgs = pdf_to_jpgs(f["path"], out_dir, dpi=300, on_page=on_page)
                total_imgs += len(imgs)
            except Exception as e:
                errors.append(f"{f['name']}：{e}")

        if errors:
            messagebox.showerror("部分轉換失敗", "\n".join(errors))
        else:
            messagebox.showinfo(
                "轉換完成",
                f"✓ {len(pdf_files)} 個 PDF 轉換完成\n"
                f"共 {total_imgs} 張 JPG（300 DPI）\n\n"
                f"儲存位置：{out_dir}",
            )
        self._reset_progress()
        self._refresh_status()

    def convert_to_word(self):
        """將清單中所有 PDF 轉成 Word (.docx)，存到原始資料夾。"""
        pdf_files = [f for f in self.files if f.get("ftype") == "pdf"]
        if not pdf_files:
            messagebox.showwarning(
                "提示", "清單中沒有 PDF 檔案\n（Word 檔請先點「轉換 Word → PDF」）"
            )
            return
        if not ensure_pdf2docx():
            messagebox.showerror(
                "缺少套件",
                "無法安裝 pdf2docx\n請手動執行：\n\npip install pdf2docx",
            )
            return

        from pdf2docx import parse

        total  = len(pdf_files)
        errors = []
        for idx, f in enumerate(pdf_files):
            out_docx = os.path.splitext(f["path"])[0] + ".docx"
            try:
                self._set_progress(
                    idx / total * 100,
                    f"轉換 Word ({idx + 1}/{total})：{f['name']}",
                )
                parse(f["path"], out_docx)
            except Exception as e:
                errors.append(f"{f['name']}：{e}")

        self._set_progress(100, f"✓ {total} 個 PDF 轉換完成")
        self._reset_progress()

        if errors:
            messagebox.showerror("部分轉換失敗", "\n".join(errors))
        else:
            messagebox.showinfo(
                "轉換完成",
                f"✓ {len(pdf_files)} 個 PDF 已轉換為 Word\n"
                f"存放位置：各原始檔所在資料夾",
            )

    def split_pdf(self):
        """將清單中選取的 PDF 逐頁拆成獨立 PDF 檔案。"""
        idx = self._sel_idx()
        if idx is None:
            messagebox.showwarning("提示", "請先在清單中選取一個 PDF 檔案")
            return
        f = self.files[idx]
        if f.get("ftype") == "word":
            messagebox.showwarning("提示", "請先將 Word 轉換為 PDF")
            return

        out_dir = filedialog.askdirectory(title="選擇分割後 PDF 的儲存資料夾")
        if not out_dir:
            return

        from pypdf import PdfReader, PdfWriter

        name = f["name"]
        base = os.path.splitext(name)[0]
        try:
            reader = PdfReader(f["path"])
            total  = len(reader.pages)
            for i, page in enumerate(reader.pages, 1):
                writer = PdfWriter()
                writer.add_page(page)
                out_path = os.path.join(out_dir, f"{base}_p{i:03d}.pdf")
                with open(out_path, "wb") as fh:
                    writer.write(fh)
                self._set_progress(
                    i / total * 100,
                    f"分割第 {i}/{total} 頁：{name}",
                )

            self._reset_progress()
            messagebox.showinfo(
                "分割完成",
                f"✓ {name} 已分割為 {total} 個 PDF\n\n儲存位置：{out_dir}",
            )
        except PermissionError:
            self._reset_progress()
            messagebox.showerror(
                "儲存失敗",
                f"無法寫入到 {out_dir}\n（請確認資料夾有寫入權限）",
            )
        except Exception as e:
            self._reset_progress()
            messagebox.showerror("分割失敗", str(e))

        self._refresh_status()

    def merge(self):
        if not self.files:
            messagebox.showwarning("提示", "請先新增 PDF 檔案")
            return

        word_files = [f for f in self.files if f.get("ftype") == "word"]
        if word_files:
            names = "\n".join(f"  • {f['name']}" for f in word_files)
            messagebox.showwarning(
                "尚有 Word 檔案未轉換",
                f"請先點擊「轉換 Word → PDF」完成轉換：\n\n{names}"
            )
            return

        out = filedialog.asksaveasfilename(
            title="儲存合併後的 PDF",
            defaultextension=".pdf",
            filetypes=[("PDF 檔案", "*.pdf")],
        )
        if not out:
            return

        from pypdf import PdfReader, PdfWriter

        total_pages = sum(f["pages"] for f in self.files)
        done_pages  = 0
        writer      = PdfWriter()
        try:
            for f in self.files:
                reader = PdfReader(f["path"])
                pages  = len(reader.pages)
                for page in reader.pages:
                    writer.add_page(page)
                    done_pages += 1
                    pct = done_pages / total_pages * 100 if total_pages > 0 else 100
                    self._set_progress(pct,
                        f"合併 {done_pages}/{total_pages} 頁：{f['name']}")
                if self.var_blank.get() and pages % 2 == 1:
                    last = reader.pages[-1]
                    writer.add_blank_page(
                        width=float(last.mediabox.width),
                        height=float(last.mediabox.height),
                    )

            self._set_progress(100, "正在寫入檔案…")
            with open(out, "wb") as fh:
                writer.write(fh)

            self._reset_progress()
            messagebox.showinfo(
                "合併完成",
                f"✓ 儲存成功\n"
                f"檔案：{os.path.basename(out)}\n"
                f"總頁數：{len(writer.pages)} 頁\n\n"
                f"路徑：{out}",
            )
            self._refresh_status()

        except PermissionError:
            self._reset_progress()
            messagebox.showerror("儲存失敗",
                                 f"無法寫入 {out}\n（檔案可能正在被其他程式開啟）")
        except Exception as e:
            self._reset_progress()
            messagebox.showerror("合併失敗", str(e))


def main():
    if not ensure_pypdf():
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("缺少套件",
                             "無法自動安裝 pypdf\n請手動執行：\n\npip install pypdf")
        root.destroy()
        return
    App().mainloop()


if __name__ == "__main__":
    main()
