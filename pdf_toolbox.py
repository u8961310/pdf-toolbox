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

import asyncio
import io
import threading

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
GREEN     = "#059669"


# ── 套件檢查 ──────────────────────────────────────────────────────────────────

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


def ensure_pil() -> bool:
    try:
        from PIL import Image, ImageTk  # noqa: F401
        return True
    except ImportError:
        return False


def ensure_winsdk() -> bool:
    try:
        import winrt.windows.media.ocr  # noqa: F401
        return True
    except ImportError:
        return False


# ── 工具函式 ──────────────────────────────────────────────────────────────────

def pdf_to_jpgs(pdf_path: str, out_dir: str, dpi: int = 300,
                on_page=None) -> list[str]:
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
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(docx_path)
        doc.SaveAs2(pdf_path, FileFormat=17)
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


WORD_EXTS = {".doc", ".docx"}
IMG_EXTS  = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}


def imgs_to_pdf(img_paths: list[str], out_path: str) -> None:
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


# ── Splash ────────────────────────────────────────────────────────────────────

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


# ── 主應用程式 ────────────────────────────────────────────────────────────────

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.title("PDF 工具箱")
        self.geometry("960x700")
        self.minsize(860, 600)
        self.resizable(True, True)
        self.configure(bg=BG)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=0)
        self.rowconfigure(1, weight=1)

        # 合併分割用
        self.files: list[dict] = []
        # OCR 用
        self.ocr_files: list[str] = []
        self.ocr_result: str = ""
        self._ocr_running: bool = False
        # 比對用
        self.cmp_path_a: str = ""
        self.cmp_path_b: str = ""
        self.cmp_doc_a = None
        self.cmp_doc_b = None
        self.cmp_page: int = 0
        self.cmp_diffs: dict[int, list] = {}
        self._cmp_img_a = None
        self._cmp_img_b = None

        self._setup_styles()
        splash = SplashScreen(self)
        splash.set_progress(60, "建立介面…")
        self._build_ui()
        splash.set_progress(100, "完成！")
        self.after(300, lambda: (splash.destroy(), self._show_main()))
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        if self.cmp_doc_a:
            self.cmp_doc_a.close()
        if self.cmp_doc_b:
            self.cmp_doc_b.close()
        self.destroy()

    def _show_main(self):
        self.deiconify()
        self.lift()
        self.focus_force()
        self._show_disclaimer()

    # ── 使用聲明 ──────────────────────────────────────────────────────────────

    def _show_disclaimer(self):
        dlg = tk.Toplevel(self)
        dlg.title("使用聲明")
        dlg.resizable(False, False)
        dlg.configure(bg=SURFACE)
        dlg.attributes("-topmost", True)

        W, H = 560, 380
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        dlg.geometry(f"{W}x{H}+{(sw - W) // 2}+{(sh - H) // 2}")

        dlg.grab_set()
        dlg.focus_set()
        dlg.protocol("WM_DELETE_WINDOW", self.destroy)

        CJK = "Microsoft JhengHei"

        tk.Frame(dlg, bg=DANGER, height=4).pack(fill="x")
        tk.Frame(dlg, bg="#fef2f2", height=6).pack(fill="x")

        pad = tk.Frame(dlg, bg=SURFACE, padx=28, pady=20)
        pad.pack(fill="both", expand=True)

        title_row = tk.Frame(pad, bg=SURFACE)
        title_row.pack(fill="x", anchor="w")
        tk.Label(title_row, text="⚠", bg=SURFACE, fg=DANGER,
                 font=(CJK, 16)).pack(side="left", padx=(0, 10))
        tk.Label(title_row, text="使用聲明", bg=SURFACE, fg=TEXT1,
                 font=(CJK, 15, "bold")).pack(side="left")

        tk.Frame(pad, bg=BORDER, height=1).pack(fill="x", pady=(14, 16))

        warn_box = tk.Frame(pad, bg="#fef2f2",
                            highlightbackground="#fca5a5", highlightthickness=1)
        warn_box.pack(fill="x")
        tk.Label(warn_box,
                 text="本軟體依使用授權協議提供。茲聲明台北市私立立人國際中小學不在本軟體授權使用人範圍內。"
                      "任何屬於或代表該機構之人員，在未獲著作權人書面授權前使用本軟體，即構成著作權侵害；"
                      "著作權人保留依中華民國著作權法第六章規定提起民刑事訴訟之權利。",
                 bg="#fef2f2", fg="#b91c1c",
                 font=(CJK, 10, "bold"),
                 justify="left", wraplength=480,
                 padx=14, pady=12).pack(anchor="w", fill="x")

        tk.Label(pad,
                 text="© 2026 陳冠廷 · 著作權所有，未經授權不得於特定機構使用。",
                 bg=SURFACE, fg=TEXT2,
                 font=(CJK, 9),
                 justify="left").pack(anchor="w", pady=(14, 0))

        def _btn(parent, text, cmd, bg, hover):
            b = tk.Button(parent, text=text, command=cmd,
                          bg=bg, fg="white", relief="flat",
                          font=(CJK, 10), padx=16, pady=8,
                          cursor="hand2", activebackground=hover,
                          activeforeground="white", bd=0)
            b.bind("<Enter>", lambda _: b.config(bg=hover))
            b.bind("<Leave>", lambda _: b.config(bg=bg))
            return b

        btn_row = tk.Frame(pad, bg=SURFACE)
        btn_row.pack(fill="x", pady=(18, 0))
        _btn(btn_row, "關閉程式",       self.destroy,  DANGER,   "#dc2626").pack(side="right", padx=(6, 0))
        _btn(btn_row, "我了解，繼續使用", dlg.destroy, PRIMARY, "#1d4ed8").pack(side="right")

        dlg.wait_window()

    # ── 樣式 ──────────────────────────────────────────────────────────────────

    def _setup_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        CJK = "Microsoft JhengHei"
        s.configure(".", background=BG, foreground=TEXT1, font=(CJK, 10))
        s.configure("TFrame",  background=BG)
        s.configure("TLabel",  background=BG, foreground=TEXT1)
        s.configure("Title.TLabel", background=BG, foreground=TEXT1,
                    font=(CJK, 15, "bold"))
        s.configure("Sub.TLabel",   background=BG, foreground=TEXT2,
                    font=(CJK, 9))
        s.configure("Info.TLabel",  background=BG, foreground=TEXT2,
                    font=(CJK, 9))

        for name, bg, fg, hover_bg in [
            ("Primary", PRIMARY,   "white", "#1d4ed8"),
            ("Danger",  DANGER,    "white", "#dc2626"),
            ("Muted",   BORDER,    TEXT1,   "#cbd5e1"),
            ("Ghost",   SURFACE,   TEXT1,   HOVER),
            ("Green",   GREEN,     "white", "#047857"),
            ("Word",    WORD_BLUE, "white", "#1e3a8a"),
        ]:
            s.configure(f"{name}.TButton",
                        background=bg, foreground=fg,
                        font=(CJK, 10, "bold" if name == "Primary" else "normal"),
                        padding=(12, 7), relief="flat", borderwidth=0)
            s.map(f"{name}.TButton", background=[("active", hover_bg)])

        s.configure("Treeview",
                    background=SURFACE, foreground=TEXT1,
                    fieldbackground=SURFACE, rowheight=36,
                    font=(CJK, 10), borderwidth=0, relief="flat")
        s.configure("Treeview.Heading",
                    background=BG, foreground=TEXT2,
                    font=(CJK, 9, "bold"), relief="flat", padding=(8, 7))
        s.map("Treeview",
              background=[("selected", HOVER)],
              foreground=[("selected", TEXT1)])
        s.configure("TCheckbutton", background=BG, foreground=TEXT1, font=(CJK, 10))
        s.configure("TSeparator", background=BORDER)
        s.configure("Horizontal.TProgressbar",
                    troughcolor=BORDER, background=PRIMARY,
                    thickness=8, borderwidth=0, relief="flat")
        s.configure("TNotebook", background=BG, borderwidth=0, tabmargins=0)
        s.configure("TNotebook.Tab", background=BG, foreground=TEXT2,
                    font=(CJK, 10), padding=(14, 8))
        s.map("TNotebook.Tab",
              background=[("selected", SURFACE)],
              foreground=[("selected", TEXT1)])

        # 舊樣式別名，保持向下相容
        s.configure("Green.TButton",    background=GREEN,     foreground="white",
                    font=(CJK, 10), padding=(12, 7), relief="flat", borderwidth=0)
        s.map("Green.TButton",    background=[("active", "#047857")])
        s.configure("WordBlue.TButton", background=WORD_BLUE, foreground="white",
                    font=(CJK, 10), padding=(12, 7), relief="flat", borderwidth=0)
        s.map("WordBlue.TButton", background=[("active", "#1e3a8a")])

    # ── UI 建構 ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        CJK = "Microsoft JhengHei"

        tk.Frame(self, bg=PRIMARY, height=3).grid(row=0, column=0, sticky="ew")

        wrap = ttk.Frame(self, padding=(20, 14, 20, 16))
        wrap.grid(row=1, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(1, weight=1)

        # Header 卡片
        hdr_card = tk.Frame(wrap, bg=SURFACE,
                            highlightbackground=BORDER, highlightthickness=1)
        hdr_card.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        tk.Frame(hdr_card, bg=PRIMARY, width=4).pack(side="left", fill="y")
        hdr_inner = tk.Frame(hdr_card, bg=SURFACE, padx=16, pady=11)
        hdr_inner.pack(side="left", fill="both", expand=True)
        tk.Label(hdr_inner, text="PDF 工具箱", bg=SURFACE, fg=TEXT1,
                 font=(CJK, 16, "bold")).pack(side="left")
        tk.Label(hdr_inner, text="  合併 · 分割 · 格式轉換 · OCR · 文件比對",
                 bg=SURFACE, fg=TEXT2, font=(CJK, 9)).pack(side="left", pady=(8, 0))
        tk.Label(hdr_inner, text="© 2026 陳冠廷",
                 bg=SURFACE, fg=TEXT2, font=(CJK, 9)).pack(side="right", pady=(8, 0))

        # Notebook
        nb = ttk.Notebook(wrap)
        nb.grid(row=1, column=0, sticky="nsew")

        tab1 = ttk.Frame(nb)
        tab2 = ttk.Frame(nb)
        tab3 = ttk.Frame(nb)
        nb.add(tab1, text="  📄  合併 & 分割  ")
        nb.add(tab2, text="  🔍  OCR 辨識  ")
        nb.add(tab3, text="  📊  文件比對  ")

        self._build_tab_merge(tab1, CJK)
        self._build_tab_ocr(tab2, CJK)
        self._build_tab_compare(tab3, CJK)

    # ── Tab 1：合併 & 分割 ────────────────────────────────────────────────────

    def _build_tab_merge(self, parent, CJK):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        body = ttk.Frame(parent)
        body.grid(row=0, column=0, sticky="nsew", padx=12, pady=(10, 0))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        # 檔案清單卡片
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
            text="📂\n\n尚無檔案\n\n點擊「＋ 新增檔案」或拖曳檔案到此處\nPDF、Word（.docx / .doc）",
            bg=SURFACE, fg="#94a3b8",
            font=(CJK, 11), justify="center",
        )

        card.drop_target_register(DND_FILES)
        card.dnd_bind("<<Drop>>", self._on_drop)
        card.dnd_bind("<<DragEnter>>", lambda e: card.config(
            highlightbackground=PRIMARY, highlightthickness=2))
        card.dnd_bind("<<DragLeave>>", lambda e: card.config(
            highlightbackground=BORDER, highlightthickness=1))

        # 側邊按鈕
        bc = ttk.Frame(body)
        bc.grid(row=0, column=1, sticky="n")

        def grp(text):
            tk.Label(bc, text=text, bg=BG, fg=TEXT2,
                     font=(CJK, 8)).pack(anchor="w", pady=(0, 4))

        def b(text, cmd, style, pady=(0, 6)):
            ttk.Button(bc, text=text, command=cmd,
                       style=style, width=11).pack(fill="x", pady=pady)

        grp("檔案管理")
        b("＋ 新增檔案", self.add_files,      "Primary.TButton", (0, 6))
        b("✕ 移除選取", self.remove_selected, "Danger.TButton",  (0, 16))
        ttk.Separator(bc, orient="horizontal").pack(fill="x", pady=(0, 10))
        grp("排列順序")
        b("↑  上移",   self.move_up,   "Ghost.TButton", (0, 4))
        b("↓  下移",   self.move_down, "Ghost.TButton", (0, 16))
        ttk.Separator(bc, orient="horizontal").pack(fill="x", pady=(0, 10))
        b("清空全部",  self.clear_all, "Muted.TButton", (0, 0))

        # 選項
        opt = ttk.Frame(parent)
        opt.grid(row=1, column=0, sticky="ew", padx=12, pady=(8, 0))
        self.var_blank = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opt,
            text=" 單數頁結尾補空白頁（讓下一份 PDF 從正面起始，適合雙面列印）",
            variable=self.var_blank, command=self._refresh_status,
            bg=BG, fg=TEXT1, font=(CJK, 10),
            activebackground=BG, cursor="hand2",
        ).pack(side="left")

        # 進度條
        prog_row = ttk.Frame(parent)
        prog_row.grid(row=2, column=0, sticky="ew", padx=12, pady=(6, 0))
        prog_row.columnconfigure(0, weight=1)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(
            prog_row, variable=self.progress_var, maximum=100, mode="determinate")
        self.progress.grid(row=0, column=0, sticky="ew")
        self.prog_label = ttk.Label(prog_row, text="", style="Info.TLabel")
        self.prog_label.grid(row=1, column=0, sticky="w", pady=(3, 0))

        # 底部操作列
        bar = ttk.Frame(parent)
        bar.grid(row=3, column=0, sticky="ew", padx=12, pady=(8, 10))
        bar.columnconfigure(0, weight=1)

        right = ttk.Frame(bar)
        right.grid(row=0, column=1, rowspan=2, sticky="se")

        conv_wrap = tk.Frame(right, bg=BG)
        conv_wrap.pack(anchor="e", pady=(0, 6))
        tk.Label(conv_wrap, text="格式轉換", bg=BG, fg=TEXT2,
                 font=(CJK, 8)).pack(side="left", padx=(0, 8))
        for text, cmd, style in [
            ("Word→PDF",  self.convert_words,       "Ghost.TButton"),
            ("圖片→PDF",  self.convert_imgs_to_pdf, "Ghost.TButton"),
            ("PDF→JPG",   self.convert_to_jpg,       "Green.TButton"),
            ("PDF→Word",  self.convert_to_word,      "WordBlue.TButton"),
        ]:
            ttk.Button(conv_wrap, text=text, command=cmd,
                       style=style).pack(side="left", padx=(0, 4))

        main_btns = ttk.Frame(right)
        main_btns.pack(anchor="e")
        ttk.Button(main_btns, text="PDF 分割",
                   command=self.split_pdf,
                   style="Muted.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(main_btns, text="  合併 PDF  →",
                   command=self.merge,
                   style="Primary.TButton").pack(side="left")

        self.status_var = tk.StringVar(value="尚未新增任何檔案")
        ttk.Label(bar, textvariable=self.status_var,
                  style="Info.TLabel").grid(row=1, column=0, sticky="w")

        self._show_empty()

    # ── Tab 2：OCR 辨識 ───────────────────────────────────────────────────────

    def _build_tab_ocr(self, parent, CJK):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        parent.rowconfigure(3, weight=2)

        # 操作列
        ctrl = tk.Frame(parent, bg=BG)
        ctrl.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        ttk.Button(ctrl, text="＋ 選擇 PDF", command=self.ocr_add_files,
                   style="Primary.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(ctrl, text="✕ 清空", command=self.ocr_clear,
                   style="Danger.TButton").pack(side="left")
        ttk.Button(ctrl, text="▶  開始辨識", command=self.run_ocr,
                   style="Green.TButton").pack(side="right")

        # 檔案清單
        lf = tk.Frame(parent, bg=SURFACE,
                      highlightbackground=BORDER, highlightthickness=1)
        lf.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 6))
        lf.columnconfigure(0, weight=1)
        lf.rowconfigure(0, weight=1)
        self.ocr_listbox = tk.Listbox(lf, bg=SURFACE, fg=TEXT1, font=(CJK, 10),
                                       selectmode="extended", relief="flat", bd=0,
                                       height=4)
        self.ocr_listbox.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(lf, orient="vertical", command=self.ocr_listbox.yview)
        self.ocr_listbox.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        # 進度
        prog = ttk.Frame(parent)
        prog.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 4))
        prog.columnconfigure(0, weight=1)
        self.ocr_pvar = tk.DoubleVar(value=0)
        ttk.Progressbar(prog, variable=self.ocr_pvar, maximum=100,
                        mode="determinate").grid(row=0, column=0, sticky="ew")
        self.ocr_status_var = tk.StringVar(value="選擇 PDF 後點「開始辨識」")
        ttk.Label(prog, textvariable=self.ocr_status_var,
                  style="Info.TLabel").grid(row=1, column=0, sticky="w", pady=(3, 0))

        # 辨識結果
        rf = tk.Frame(parent, bg=SURFACE,
                      highlightbackground=BORDER, highlightthickness=1)
        rf.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 6))
        rf.columnconfigure(0, weight=1)
        rf.rowconfigure(0, weight=1)
        self.ocr_text = tk.Text(rf, bg=SURFACE, fg=TEXT1, font=(CJK, 10),
                                 wrap="word", relief="flat", state="disabled",
                                 padx=10, pady=8)
        self.ocr_text.grid(row=0, column=0, sticky="nsew")
        vsb2 = ttk.Scrollbar(rf, orient="vertical", command=self.ocr_text.yview)
        self.ocr_text.configure(yscrollcommand=vsb2.set)
        vsb2.grid(row=0, column=1, sticky="ns")

        # 儲存按鈕
        save_row = ttk.Frame(parent)
        save_row.grid(row=4, column=0, sticky="e", padx=12, pady=(0, 10))
        ttk.Button(save_row, text="存成 PDF",
                   command=lambda: self.save_ocr_result("pdf"),
                   style="Ghost.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(save_row, text="存成 Word",
                   command=lambda: self.save_ocr_result("word"),
                   style="WordBlue.TButton").pack(side="left")

    # ── Tab 3：文件比對（所見即所得）────────────────────────────────────────────

    def _build_tab_compare(self, parent, CJK):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(2, weight=1)

        # 左右文件選擇卡片
        def file_card(col, label, var_attr, side_color, cmd):
            px_l = 12 if col == 0 else 6
            px_r = 6  if col == 0 else 12
            card = tk.Frame(parent, bg=SURFACE,
                            highlightbackground=BORDER, highlightthickness=1)
            card.grid(row=0, column=col, sticky="ew",
                      padx=(px_l, px_r), pady=(10, 6))
            tk.Frame(card, bg=side_color, width=4).pack(side="left", fill="y")
            inner = tk.Frame(card, bg=SURFACE, padx=12, pady=10)
            inner.pack(side="left", fill="both", expand=True)
            tk.Label(inner, text=label, bg=SURFACE, fg=TEXT2,
                     font=(CJK, 8, "bold")).pack(anchor="w")
            var = tk.StringVar(value="尚未選擇")
            setattr(self, var_attr, var)
            tk.Label(inner, textvariable=var, bg=SURFACE, fg=TEXT1,
                     font=(CJK, 10), wraplength=320,
                     justify="left").pack(anchor="w", pady=(4, 6))
            ttk.Button(inner, text="選擇檔案", command=cmd,
                       style="Ghost.TButton").pack(anchor="e")

        file_card(0, "文件 A", "cmp_var_a", PRIMARY, lambda: self.cmp_pick("a"))
        file_card(1, "文件 B", "cmp_var_b", GREEN,   lambda: self.cmp_pick("b"))

        # 控制列（頁面導航 ＋ 比對按鈕）
        ctrl = tk.Frame(parent, bg=BG)
        ctrl.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 6))

        nav = tk.Frame(ctrl, bg=BG)
        nav.pack(side="left")
        ttk.Button(nav, text=" ◀ ", command=self.cmp_prev_page,
                   style="Ghost.TButton").pack(side="left")
        self.cmp_page_var = tk.StringVar(value="請先選擇文件")
        tk.Label(nav, textvariable=self.cmp_page_var,
                 bg=BG, fg=TEXT1, font=(CJK, 10), width=16).pack(side="left", padx=4)
        ttk.Button(nav, text=" ▶ ", command=self.cmp_next_page,
                   style="Ghost.TButton").pack(side="left")

        self.cmp_status_var = tk.StringVar(value="選好兩份文件後點「開始比對」")
        tk.Label(ctrl, textvariable=self.cmp_status_var,
                 bg=BG, fg=TEXT2, font=(CJK, 9)).pack(side="left", padx=14)

        ttk.Button(ctrl, text="  🔎  開始比對  ",
                   command=self.run_compare,
                   style="Primary.TButton").pack(side="right")

        # 雙 Canvas（左=文件A，右=文件B）
        ca_frame = tk.Frame(parent, bg=BG)
        ca_frame.grid(row=2, column=0, columnspan=2,
                      sticky="nsew", padx=12, pady=(0, 4))
        ca_frame.columnconfigure(0, weight=1)
        ca_frame.columnconfigure(1, weight=1)
        ca_frame.rowconfigure(0, weight=1)

        def make_pane(col):
            pane = tk.Frame(ca_frame, bg=SURFACE,
                            highlightbackground=BORDER, highlightthickness=1)
            pane.grid(row=0, column=col, sticky="nsew",
                      padx=(0, 5) if col == 0 else (5, 0))
            pane.columnconfigure(0, weight=1)
            pane.rowconfigure(0, weight=1)
            c = tk.Canvas(pane, bg="#f8fafc", bd=0, highlightthickness=0)
            vsb = ttk.Scrollbar(pane, orient="vertical",   command=c.yview)
            hsb = ttk.Scrollbar(pane, orient="horizontal", command=c.xview)
            c.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            c.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            hsb.grid(row=1, column=0, sticky="ew")
            c.bind("<MouseWheel>",
                   lambda e, _c=c: _c.yview_scroll(-1 if e.delta > 0 else 1, "units"))
            return c

        self.cmp_canvas_a = make_pane(0)
        self.cmp_canvas_b = make_pane(1)

        # 圖例
        legend = tk.Frame(parent, bg=BG)
        legend.grid(row=3, column=0, columnspan=2,
                    sticky="w", padx=12, pady=(2, 8))
        for color, text in [
            ("#ef4444", "■ 字型 / 大小 / 顏色差異"),
            (PRIMARY,   "■ 僅出現於文件 A"),
            (GREEN,     "■ 僅出現於文件 B"),
        ]:
            tk.Label(legend, text=text, bg=BG, fg=color,
                     font=(CJK, 9)).pack(side="left", padx=(0, 18))

    # ── 狀態刷新（合併分割用）────────────────────────────────────────────────

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

    # ── 合併分割動作 ──────────────────────────────────────────────────────────

    @staticmethod
    def _parse_dnd_paths(data: str) -> list[str]:
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
        event.widget.config(highlightbackground=BORDER, highlightthickness=1)
        self._reset_progress()
        self._refresh_tree()

    def convert_imgs_to_pdf(self):
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
            title="儲存為 PDF", defaultextension=".pdf",
            filetypes=[("PDF 檔案", "*.pdf")], initialfile="圖片合併.pdf",
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
            self._set_progress(i / total * 100 if total else 100, f"讀取 {i}/{total}：{name}")
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
        word_files = [f for f in self.files if f.get("ftype") == "word"]
        if not word_files:
            messagebox.showinfo("提示", "清單中沒有 Word 檔案")
            return
        if not ensure_win32com():
            messagebox.showerror("缺少套件",
                "無法安裝 pywin32\n請手動執行：\n\npip install pywin32\n\n（需要已安裝 Microsoft Word）")
            return
        from pypdf import PdfReader
        total, errors = len(word_files), []
        for idx, f in enumerate(word_files):
            out_pdf = os.path.splitext(f["path"])[0] + ".pdf"
            try:
                self._set_progress(idx / total * 100,
                                   f"轉換 Word ({idx + 1}/{total})：{f['name']}")
                docx_to_pdf(os.path.abspath(f["path"]), os.path.abspath(out_pdf))
                pages = len(PdfReader(out_pdf).pages)
                f["path"] = out_pdf
                f["name"] = os.path.basename(out_pdf)
                f["pages"] = pages
                f["ftype"] = "pdf"
            except Exception as e:
                errors.append(f"{f['name']}：{e}")
        self._set_progress(100, f"✓ {total} 個 Word 檔案轉換完成")
        self._refresh_tree()
        if errors:
            messagebox.showerror("部分轉換失敗", "\n".join(errors))
        else:
            messagebox.showinfo("轉換完成",
                f"✓ {len(word_files)} 個 Word 檔案已轉換為 PDF\n存放位置：原始檔案所在資料夾")
        self._reset_progress()

    def convert_to_jpg(self):
        pdf_files = [f for f in self.files if f.get("ftype") == "pdf"]
        if not pdf_files:
            messagebox.showwarning("提示", "清單中沒有 PDF 檔案\n（Word 檔請先點「轉換 Word → PDF」）")
            return
        if not ensure_pymupdf():
            messagebox.showerror("缺少套件", "請手動執行：\n\npip install pymupdf")
            return
        out_dir = filedialog.askdirectory(title="選擇 JPG 儲存資料夾")
        if not out_dir:
            return
        total_pages = sum(f["pages"] for f in pdf_files)
        done, errors, total_imgs = [0], [], 0
        for f in pdf_files:
            try:
                fname = f["name"]
                def on_page(cur, file_total, _n=fname, _d=done):
                    _d[0] += 1
                    pct = _d[0] / total_pages * 100 if total_pages > 0 else 100
                    self._set_progress(pct, f"轉換第 {_d[0]}/{total_pages} 頁：{_n}")
                imgs = pdf_to_jpgs(f["path"], out_dir, dpi=300, on_page=on_page)
                total_imgs += len(imgs)
            except Exception as e:
                errors.append(f"{f['name']}：{e}")
        if errors:
            messagebox.showerror("部分轉換失敗", "\n".join(errors))
        else:
            messagebox.showinfo("轉換完成",
                f"✓ {len(pdf_files)} 個 PDF 轉換完成\n共 {total_imgs} 張 JPG（300 DPI）\n\n儲存位置：{out_dir}")
        self._reset_progress()
        self._refresh_status()

    def convert_to_word(self):
        pdf_files = [f for f in self.files if f.get("ftype") == "pdf"]
        if not pdf_files:
            messagebox.showwarning("提示", "清單中沒有 PDF 檔案")
            return
        if not ensure_pdf2docx():
            messagebox.showerror("缺少套件", "請手動執行：\n\npip install pdf2docx")
            return
        from pdf2docx import parse
        total, errors = len(pdf_files), []
        for idx, f in enumerate(pdf_files):
            out_docx = os.path.splitext(f["path"])[0] + ".docx"
            try:
                self._set_progress(idx / total * 100,
                                   f"轉換 Word ({idx + 1}/{total})：{f['name']}")
                parse(f["path"], out_docx)
            except Exception as e:
                errors.append(f"{f['name']}：{e}")
        self._set_progress(100, f"✓ {total} 個 PDF 轉換完成")
        self._reset_progress()
        if errors:
            messagebox.showerror("部分轉換失敗", "\n".join(errors))
        else:
            messagebox.showinfo("轉換完成",
                f"✓ {len(pdf_files)} 個 PDF 已轉換為 Word\n存放位置：各原始檔所在資料夾")

    def split_pdf(self):
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
            total = len(reader.pages)
            for i, page in enumerate(reader.pages, 1):
                writer = PdfWriter()
                writer.add_page(page)
                out_path = os.path.join(out_dir, f"{base}_p{i:03d}.pdf")
                with open(out_path, "wb") as fh:
                    writer.write(fh)
                self._set_progress(i / total * 100, f"分割第 {i}/{total} 頁：{name}")
            self._reset_progress()
            messagebox.showinfo("分割完成",
                f"✓ {name} 已分割為 {total} 個 PDF\n\n儲存位置：{out_dir}")
        except PermissionError:
            self._reset_progress()
            messagebox.showerror("儲存失敗", f"無法寫入到 {out_dir}\n（請確認資料夾有寫入權限）")
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
            messagebox.showwarning("尚有 Word 檔案未轉換",
                f"請先點擊「轉換 Word → PDF」完成轉換：\n\n{names}")
            return
        out = filedialog.asksaveasfilename(
            title="儲存合併後的 PDF", defaultextension=".pdf",
            filetypes=[("PDF 檔案", "*.pdf")],
        )
        if not out:
            return
        from pypdf import PdfReader, PdfWriter
        total_pages = sum(f["pages"] for f in self.files)
        done_pages = 0
        writer = PdfWriter()
        try:
            for f in self.files:
                reader = PdfReader(f["path"])
                pages = len(reader.pages)
                for page in reader.pages:
                    writer.add_page(page)
                    done_pages += 1
                    pct = done_pages / total_pages * 100 if total_pages > 0 else 100
                    self._set_progress(pct, f"合併 {done_pages}/{total_pages} 頁：{f['name']}")
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
            messagebox.showinfo("合併完成",
                f"✓ 儲存成功\n檔案：{os.path.basename(out)}\n"
                f"總頁數：{len(writer.pages)} 頁\n\n路徑：{out}")
            self._refresh_status()
        except PermissionError:
            self._reset_progress()
            messagebox.showerror("儲存失敗", f"無法寫入 {out}\n（檔案可能正在被其他程式開啟）")
        except Exception as e:
            self._reset_progress()
            messagebox.showerror("合併失敗", str(e))

    # ── OCR 動作 ──────────────────────────────────────────────────────────────

    def ocr_add_files(self):
        paths = filedialog.askopenfilenames(
            title="選擇 PDF 檔案",
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")],
        )
        for p in paths:
            if p not in self.ocr_files:
                self.ocr_files.append(p)
                self.ocr_listbox.insert("end", os.path.basename(p))

    def ocr_clear(self):
        self.ocr_files.clear()
        self.ocr_listbox.delete(0, "end")
        self.ocr_text.config(state="normal")
        self.ocr_text.delete("1.0", "end")
        self.ocr_text.config(state="disabled")
        self.ocr_pvar.set(0)
        self.ocr_status_var.set("選擇 PDF 後點「開始辨識」")

    def run_ocr(self):
        if self._ocr_running:
            return
        if not self.ocr_files:
            messagebox.showwarning("提示", "請先選擇 PDF 檔案")
            return
        if not ensure_pymupdf():
            messagebox.showerror("缺少套件", "請手動執行：\n\npip install pymupdf")
            return
        if not ensure_winsdk():
            messagebox.showerror("OCR 不可用",
                "找不到 Windows OCR 元件（winsdk）。\n"
                "此功能需要 Windows 10 1803 以上版本。")
            return

        self._ocr_running = True
        self.ocr_pvar.set(2)
        self.ocr_status_var.set("準備 Windows OCR…")
        threading.Thread(target=self._ocr_worker,
                         args=(list(self.ocr_files),), daemon=True).start()

    def _ocr_worker(self, pdf_paths: list[str]):
        import fitz
        from PIL import Image

        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        all_text: list[str] = []

        total_pages = 0
        for p in pdf_paths:
            try:
                total_pages += len(fitz.open(p))
            except Exception:
                pass

        done = 0
        for pdf_path in pdf_paths:
            name = os.path.basename(pdf_path)
            all_text.append(f"{'═' * 38}\n{name}\n{'═' * 38}")
            try:
                doc = fitz.open(pdf_path)
                for i, page in enumerate(doc, 1):
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                    pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    try:
                        text = loop.run_until_complete(self._win_ocr_page(pil_img))
                    except Exception as e:
                        text = f"[辨識錯誤：{e}]"
                    all_text.append(f"\n── 第 {i} 頁 ──\n{text}")
                    done += 1
                    pct = done / total_pages * 100 if total_pages else 100
                    status = f"辨識中 {done}/{total_pages} 頁：{name} 第 {i} 頁"
                    self.after(0, lambda p=pct, s=status: (
                        self.ocr_pvar.set(p), self.ocr_status_var.set(s)))
                doc.close()
            except Exception as e:
                all_text.append(f"[錯誤] {name}：{e}")

        loop.close()
        result = "\n\n".join(all_text)
        self.after(0, lambda: self._ocr_worker_done(result, total_pages))

    async def _win_ocr_page(self, pil_img):
        """使用 Windows.Media.OCR 辨識單頁圖片（在 background thread 的 asyncio loop 中執行）。"""
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.globalization import Language
        from winrt.windows.graphics.imaging import (
            BitmapDecoder, BitmapPixelFormat, BitmapAlphaMode, SoftwareBitmap,
        )
        from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter

        buf = io.BytesIO()
        pil_img.save(buf, format="PNG")
        img_bytes = buf.getvalue()

        stream = InMemoryRandomAccessStream()
        writer = DataWriter(stream.get_output_stream_at(0))
        writer.write_bytes(img_bytes)
        await writer.store_async()
        writer.detach_stream()
        stream.seek(0)

        decoder = await BitmapDecoder.create_async(stream)
        sb = await decoder.get_software_bitmap_async()
        if sb.bitmap_pixel_format != BitmapPixelFormat.BGRA8:
            sb = SoftwareBitmap.convert(
                sb, BitmapPixelFormat.BGRA8, BitmapAlphaMode.PREMULTIPLIED)

        lang = Language("zh-Hant")
        engine = (OcrEngine.try_create_from_language(lang)
                  if OcrEngine.is_language_supported(lang)
                  else OcrEngine.try_create_from_user_profile_languages())
        if not engine:
            return "[無法建立 OCR 引擎，請確認已安裝繁體中文語言包]"

        result = await engine.recognize_async(sb)
        return "\n".join(line.text for line in result.lines)

    def _ocr_worker_done(self, result: str, total_pages: int):
        self._ocr_running = False
        self.ocr_result = result
        self.ocr_text.config(state="normal")
        self.ocr_text.delete("1.0", "end")
        self.ocr_text.insert("1.0", result)
        self.ocr_text.config(state="disabled")
        self.ocr_pvar.set(100)
        self.ocr_status_var.set(f"✓ 辨識完成，共 {total_pages} 頁")

    def save_ocr_result(self, fmt: str):
        if not self.ocr_result:
            messagebox.showwarning("提示", "尚未執行 OCR 辨識")
            return

        if fmt == "word":
            out = filedialog.asksaveasfilename(
                title="存成 Word", defaultextension=".docx",
                filetypes=[("Word 文件", "*.docx")],
            )
            if not out:
                return
            try:
                from docx import Document
                from docx.shared import Pt
                doc = Document()
                style = doc.styles["Normal"]
                style.font.name = "微軟正黑體"
                style.font.size = Pt(12)
                for line in self.ocr_result.split("\n"):
                    doc.add_paragraph(line)
                doc.save(out)
                messagebox.showinfo("完成", f"✓ 已存成 Word：\n{out}")
            except Exception as e:
                messagebox.showerror("儲存失敗", str(e))

        elif fmt == "pdf":
            out = filedialog.asksaveasfilename(
                title="存成 PDF", defaultextension=".pdf",
                filetypes=[("PDF 檔案", "*.pdf")],
            )
            if not out:
                return
            try:
                import fitz
                doc = fitz.open()
                lines = self.ocr_result.split("\n")
                chunk_size = 55
                for i in range(0, len(lines), chunk_size):
                    page = doc.new_page(width=595, height=842)
                    page.insert_textbox(
                        fitz.Rect(50, 50, 545, 800),
                        "\n".join(lines[i:i + chunk_size]),
                        fontname="china-t", fontsize=11,
                        color=(0, 0, 0),
                    )
                doc.save(out)
                doc.close()
                messagebox.showinfo("完成", f"✓ 已存成 PDF：\n{out}")
            except Exception as e:
                messagebox.showerror("儲存失敗", str(e))

    # ── 文件比對動作（所見即所得）────────────────────────────────────────────

    def cmp_pick(self, which: str):
        path = filedialog.askopenfilename(
            title=f"選擇文件 {'A' if which == 'a' else 'B'}",
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")],
        )
        if not path:
            return
        if not ensure_pymupdf():
            messagebox.showerror("缺少套件", "請手動執行：\n\npip install pymupdf")
            return
        import fitz
        try:
            doc = fitz.open(path)
        except Exception as e:
            messagebox.showerror("開啟失敗", str(e))
            return

        if which == "a":
            if self.cmp_doc_a:
                self.cmp_doc_a.close()
            self.cmp_doc_a = doc
            self.cmp_path_a = path
            self.cmp_var_a.set(os.path.basename(path))
        else:
            if self.cmp_doc_b:
                self.cmp_doc_b.close()
            self.cmp_doc_b = doc
            self.cmp_path_b = path
            self.cmp_var_b.set(os.path.basename(path))

        self.cmp_diffs = {}
        self.cmp_page = 0
        self._render_compare_page()

    def _get_page_spans(self, doc, page_i: int) -> list[dict]:
        """從已開啟的 fitz doc 取出第 page_i 頁所有 span。"""
        if page_i >= len(doc):
            return []
        spans = []
        for block in doc[page_i].get_text("dict", flags=0)["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    if not text:
                        continue
                    c = span["color"]
                    color_hex = f"#{(c>>16)&0xFF:02X}{(c>>8)&0xFF:02X}{c&0xFF:02X}"
                    flags = span["flags"]
                    spans.append({
                        "text":  text,
                        "bbox":  span["bbox"],
                        "font":  span["font"].split("+")[-1],
                        "size":  round(span["size"], 1),
                        "color": color_hex,
                        "style": ("粗" if flags & 16 else "") + ("斜" if flags & 2 else ""),
                    })
        return spans

    def run_compare(self):
        if not self.cmp_doc_a or not self.cmp_doc_b:
            messagebox.showwarning("提示", "請先選擇兩份文件")
            return

        max_pages = max(len(self.cmp_doc_a), len(self.cmp_doc_b))
        all_diffs: dict[int, list] = {}

        for page_i in range(max_pages):
            spans_a = self._get_page_spans(self.cmp_doc_a, page_i)
            spans_b = self._get_page_spans(self.cmp_doc_b, page_i)
            lookup_a: dict[str, dict] = {}
            for s in spans_a:
                lookup_a.setdefault(s["text"], s)
            lookup_b: dict[str, dict] = {}
            for s in spans_b:
                lookup_b.setdefault(s["text"], s)

            page_diffs: list[dict] = []
            for text, sa in lookup_a.items():
                if text not in lookup_b:
                    page_diffs.append({"type": "only_a", "bbox": sa["bbox"]})
                else:
                    sb = lookup_b[text]
                    if (sa["font"] != sb["font"]
                            or sa["size"] != sb["size"]
                            or sa["color"] != sb["color"]):
                        page_diffs.append({
                            "type": "diff",
                            "bbox_a": sa["bbox"],
                            "bbox_b": sb["bbox"],
                        })
            for text, sb in lookup_b.items():
                if text not in lookup_a:
                    page_diffs.append({"type": "only_b", "bbox": sb["bbox"]})
            all_diffs[page_i] = page_diffs

        self.cmp_diffs = all_diffs
        total_diffs = sum(len(v) for v in all_diffs.values())
        n_pages_hit = sum(1 for v in all_diffs.values() if v)
        self.cmp_status_var.set(
            f"✓ 比對完成：{total_diffs} 處差異，分布於 {n_pages_hit}/{max_pages} 頁"
            if total_diffs else "✓ 兩份文件格式完全一致")
        self.cmp_page = 0
        self._render_compare_page()

    def _render_compare_page(self):
        """把兩份 PDF 當前頁渲染到 Canvas，在差異處畫彩色框。"""
        if not ensure_pymupdf() or not ensure_pil():
            return
        import fitz
        from PIL import Image, ImageTk

        SCALE = 1.5
        mat = fitz.Matrix(SCALE, SCALE)
        page_diffs = self.cmp_diffs.get(self.cmp_page, [])

        def render_one(doc, canvas, img_attr, which):
            canvas.delete("all")
            if not doc or self.cmp_page >= len(doc):
                canvas.create_text(
                    16, 16, anchor="nw", text="（此文件無此頁）",
                    fill=TEXT2, font=("Microsoft JhengHei", 11))
                return
            pix = doc[self.cmp_page].get_pixmap(matrix=mat, alpha=False)
            photo = ImageTk.PhotoImage(
                Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
            setattr(self, img_attr, photo)
            canvas.config(scrollregion=(0, 0, pix.width, pix.height))
            canvas.create_image(0, 0, anchor="nw", image=photo)
            for d in page_diffs:
                dtype = d["type"]
                if dtype == "diff":
                    raw = d.get(f"bbox_{which}")
                    if raw:
                        x0, y0, x1, y1 = [v * SCALE for v in raw]
                        canvas.create_rectangle(
                            x0, y0, x1, y1,
                            outline="#ef4444", width=2, fill="#fecaca")
                elif dtype == f"only_{which}":
                    x0, y0, x1, y1 = [v * SCALE for v in d["bbox"]]
                    fill  = "#bfdbfe" if which == "a" else "#bbf7d0"
                    color = PRIMARY   if which == "a" else GREEN
                    canvas.create_rectangle(
                        x0, y0, x1, y1,
                        outline=color, width=2, fill=fill)

        render_one(self.cmp_doc_a, self.cmp_canvas_a, "_cmp_img_a", "a")
        render_one(self.cmp_doc_b, self.cmp_canvas_b, "_cmp_img_b", "b")

        total_a = len(self.cmp_doc_a) if self.cmp_doc_a else 0
        total_b = len(self.cmp_doc_b) if self.cmp_doc_b else 0
        total = max(total_a, total_b, 1)
        self.cmp_page_var.set(f"第 {self.cmp_page + 1} / {total} 頁")

    def cmp_prev_page(self):
        if self.cmp_page > 0:
            self.cmp_page -= 1
            self._render_compare_page()

    def cmp_next_page(self):
        total_a = len(self.cmp_doc_a) if self.cmp_doc_a else 0
        total_b = len(self.cmp_doc_b) if self.cmp_doc_b else 0
        if self.cmp_page < max(total_a, total_b) - 1:
            self.cmp_page += 1
            self._render_compare_page()


def main():
    if not ensure_pypdf():
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("缺少套件",
                             "無法找到 pypdf\n請手動執行：\n\npip install pypdf")
        root.destroy()
        return
    App().mainloop()


if __name__ == "__main__":
    main()
