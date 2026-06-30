# PDF 工具箱新功能實作計畫

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在現有 PDF 工具箱新增「PDF 轉 Word」與「PDF 逐頁分割」兩個功能，並重排底部按鈕列以容納共五顆操作按鈕。

**Architecture:** 單一檔案 `pdf_toolbox.py` 新增兩個 helper function（`ensure_pdf2docx`）、兩個 App method（`convert_to_word`、`split_pdf`）、一個新 Button 樣式（`WordBlue.TButton`），最後重排 `_build_ui()` 底部 bar 為兩行。

**Tech Stack:** Python 3.10+、tkinter、pypdf（既有）、pymupdf（既有）、pdf2docx（新增，自動 pip 安裝）、pywin32（既有）

## Global Constraints

- 單一檔案 `pdf_toolbox.py`，不建立任何新 Python 檔案
- 沿用既有 `_auto_install()` 模式處理套件依賴
- 色票全用已定義的常數（`WORD_BLUE = "#1d4ed8"`、`BG`、`TEXT1` 等）
- 沿用 `self._set_progress()` / `self._reset_progress()` 處理進度條
- 分割輸出命名格式：`{basename}_p{N:03d}.pdf`
- PDF→Word 輸出到與原 PDF 相同目錄，副檔名 `.docx`

---

## 檔案結構

只修改一個檔案：

| 檔案 | 動作 | 說明 |
|------|------|------|
| `pdf_toolbox.py` | 修改 | 新增函式、方法、樣式、重排 UI |

---

### Task 1：`ensure_pdf2docx()` + `WordBlue.TButton` 樣式

**Files:**
- Modify: `pdf_toolbox.py:65`（`ensure_pymupdf` 之後插入）
- Modify: `pdf_toolbox.py:163`（`Green.TButton` 之後插入）

**Interfaces:**
- Produces: `ensure_pdf2docx() -> bool`（Task 2 呼叫）
- Produces: style name `"WordBlue.TButton"`（Task 4 使用）

- [ ] **Step 1：在 `ensure_pymupdf()` 之後（第 65 行後）插入 `ensure_pdf2docx()`**

在 `def ensure_pymupdf():` 整個函式結尾（`return _auto_install("pymupdf")` 那行）的下方空一行後加入：

```python
def ensure_pdf2docx() -> bool:
    try:
        import pdf2docx  # noqa: F401
        return True
    except ImportError:
        return _auto_install("pdf2docx")
```

- [ ] **Step 2：在 `_setup_styles()` 裡 `Green.TButton` 區塊之後（第 163 行後）加入 `WordBlue.TButton`**

在 `s.map("Green.TButton", ...)` 那行的正下方加入：

```python
        s.configure("WordBlue.TButton",
                    background=WORD_BLUE, foreground="white",
                    font=("Segoe UI", 10), padding=(12, 7),
                    relief="flat", borderwidth=0)
        s.map("WordBlue.TButton", background=[("active", "#1e3a8a")])
```

- [ ] **Step 3：驗證 app 能正常啟動**

```bash
cd /d/code/repos/pdf-toolbox
python pdf_toolbox.py
```

預期：視窗正常開啟，無 traceback，功能與之前完全相同（尚未加新按鈕）。

- [ ] **Step 4：Commit**

```bash
git add pdf_toolbox.py
git commit -m "feat: add ensure_pdf2docx helper and WordBlue button style"
```

---

### Task 2：`App.convert_to_word()` 方法

**Files:**
- Modify: `pdf_toolbox.py:507`（在 `convert_to_jpg` 結尾後、`merge` 開始前插入）

**Interfaces:**
- Consumes: `ensure_pdf2docx() -> bool`（Task 1）
- Consumes: `self.files: list[dict]`（`ftype`, `path`, `name` 欄位）
- Consumes: `self._set_progress(pct, detail)`、`self._reset_progress()`
- Produces: `App.convert_to_word(self)` — Task 4 的按鈕 command

- [ ] **Step 1：在 `convert_to_jpg()` 結尾後（約第 507 行）、`def merge(self):` 前插入 `convert_to_word`**

```python
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
```

- [ ] **Step 2：驗證 app 能正常啟動（方法存在但尚未連接按鈕）**

```bash
python pdf_toolbox.py
```

預期：視窗正常開啟，無 traceback。

- [ ] **Step 3：Commit**

```bash
git add pdf_toolbox.py
git commit -m "feat: add convert_to_word method (PDF -> docx via pdf2docx)"
```

---

### Task 3：`App.split_pdf()` 方法

**Files:**
- Modify: `pdf_toolbox.py`（緊接在 `convert_to_word` 之後、`def merge` 之前）

**Interfaces:**
- Consumes: `self._sel_idx() -> int | None`（取得選取列的索引）
- Consumes: `self.files: list[dict]`（`ftype`, `path`, `name`, `pages` 欄位）
- Consumes: `self._set_progress(pct, detail)`、`self._reset_progress()`、`self._refresh_status()`
- Consumes: `pypdf.PdfReader`、`pypdf.PdfWriter`（既有依賴）
- Produces: `App.split_pdf(self)` — Task 4 的按鈕 command

- [ ] **Step 1：在 `convert_to_word()` 結尾後、`def merge(self):` 前插入 `split_pdf`**

```python
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
```

- [ ] **Step 2：驗證 app 能正常啟動**

```bash
python pdf_toolbox.py
```

預期：視窗正常開啟，無 traceback。

- [ ] **Step 3：Commit**

```bash
git add pdf_toolbox.py
git commit -m "feat: add split_pdf method (one PDF per page)"
```

---

### Task 4：UI 重排（底部按鈕兩行 + Header 副標題）

**Files:**
- Modify: `pdf_toolbox.py:180-181`（Header 副標題）
- Modify: `pdf_toolbox.py:268-285`（底部 bar 區塊）

**Interfaces:**
- Consumes: `App.convert_to_word`（Task 2）
- Consumes: `App.split_pdf`（Task 3）
- Consumes: `"WordBlue.TButton"` 樣式（Task 1）

- [ ] **Step 1：更新 Header 副標題（第 180-181 行）**

將：
```python
        ttk.Label(hdr, text="   合併 ／ Word 轉 PDF ／ PDF 轉 JPG",
                  style="Sub.TLabel").pack(side="left", pady=(6, 0))
```

改為：
```python
        ttk.Label(hdr, text="   合併 ／ Word 轉 PDF ／ PDF 轉 JPG ／ PDF 轉 Word ／ PDF 分割",
                  style="Sub.TLabel").pack(side="left", pady=(6, 0))
```

- [ ] **Step 2：將底部 bar 區塊（第 268-283 行）改為兩行版**

將整個 `# Status + 操作按鈕` 區塊：

```python
        # Status + 操作按鈕
        bar = ttk.Frame(wrap)
        bar.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        bar.columnconfigure(0, weight=1)
        self.status_var = tk.StringVar(value="尚未新增任何檔案")
        ttk.Label(bar, textvariable=self.status_var,
                  style="Info.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(bar, text="轉換 Word → PDF",
                   command=self.convert_words,
                   style="Ghost.TButton").grid(row=0, column=1, padx=(0, 8))
        ttk.Button(bar, text="PDF → JPG",
                   command=self.convert_to_jpg,
                   style="Green.TButton").grid(row=0, column=2, padx=(0, 8))
        ttk.Button(bar, text="  合併 PDF  →",
                   command=self.merge,
                   style="Primary.TButton").grid(row=0, column=3)
```

替換為：

```python
        # Status + 操作按鈕（兩行）
        bar = ttk.Frame(wrap)
        bar.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        bar.columnconfigure(0, weight=1)

        btn_row = ttk.Frame(bar)
        btn_row.grid(row=0, column=0, sticky="e", pady=(0, 4))
        ttk.Button(btn_row, text="轉換 Word → PDF",
                   command=self.convert_words,
                   style="Ghost.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(btn_row, text="PDF → JPG",
                   command=self.convert_to_jpg,
                   style="Green.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(btn_row, text="PDF → Word",
                   command=self.convert_to_word,
                   style="WordBlue.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(btn_row, text="PDF 分割",
                   command=self.split_pdf,
                   style="Muted.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(btn_row, text="  合併 PDF  →",
                   command=self.merge,
                   style="Primary.TButton").pack(side="left")

        self.status_var = tk.StringVar(value="尚未新增任何檔案")
        ttk.Label(bar, textvariable=self.status_var,
                  style="Info.TLabel").grid(row=1, column=0, sticky="w")
```

- [ ] **Step 3：啟動 app 做完整手動驗證**

```bash
python pdf_toolbox.py
```

驗證清單：
1. 視窗標題列 Header 副標題顯示「合併 ／ Word 轉 PDF ／ PDF 轉 JPG ／ PDF 轉 Word ／ PDF 分割」
2. 底部共五顆按鈕依序出現：`轉換 Word → PDF`（灰）→ `PDF → JPG`（綠）→ `PDF → Word`（深藍）→ `PDF 分割`（淡灰）→ `合併 PDF →`（藍）
3. 狀態文字在按鈕列下方獨立一行
4. 點「PDF → Word」無選取 PDF 時 → 顯示 warning「清單中沒有 PDF 檔案」
5. 點「PDF 分割」無選取時 → 顯示 warning「請先在清單中選取一個 PDF 檔案」
6. 新增一個 PDF 到清單，選取後點「PDF 分割」→ 選資料夾 → 完成後在資料夾看到 `{name}_p001.pdf` 等檔案
7. 新增一個 PDF 到清單，點「PDF → Word」→ 選資料夾 → 進度條運行 → 完成後在原目錄看到同名 `.docx`
8. 既有三個功能（合併、Word→PDF、PDF→JPG）行為不變

- [ ] **Step 4：Commit**

```bash
git add pdf_toolbox.py
git commit -m "feat: restructure bottom bar and add PDF→Word, PDF split buttons"
```
