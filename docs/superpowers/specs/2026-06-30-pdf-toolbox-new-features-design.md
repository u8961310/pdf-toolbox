# PDF 工具箱新功能設計規格

- **日期**：2026-06-30
- **狀態**：待實作
- **影響範圍**：`pdf_toolbox.py`（單一檔案）

---

## 背景

PDF 工具箱（`repos/pdf-toolbox/pdf_toolbox.py`）是一個 Python tkinter 桌面工具，現有三個功能：PDF 合併、Word → PDF 轉換、PDF → JPG 轉換。本次新增兩個功能：**PDF → Word 轉換**與 **PDF 逐頁分割**。

---

## 新功能概覽

| 功能 | 新增依賴 | 操作對象 |
|------|---------|---------|
| PDF → Word | `pdf2docx`（自動 pip 安裝） | 清單中所有 PDF |
| PDF 分割（逐頁） | 無（使用既有 `pypdf`） | 清單中選取的單一 PDF |

---

## 一、PDF 轉 Word

### 核心函式

**`ensure_pdf2docx() -> bool`**
- 同 `ensure_pypdf` / `ensure_pymupdf` 模式
- 先 `import pdf2docx`，若 ImportError 則呼叫 `_auto_install("pdf2docx")`
- 回傳 bool 表示是否可用

**`App.convert_to_word()`**
- 過濾 `self.files` 中 `ftype == "pdf"` 的項目
- 若無 PDF → `messagebox.showwarning("提示", "清單中沒有 PDF 檔案\n（Word 檔請先點「轉換 Word → PDF」）")`
- 呼叫 `ensure_pdf2docx()`，失敗 → 顯示手動安裝指令 `pip install pdf2docx`
- 逐檔執行：
  ```python
  from pdf2docx import parse
  parse(pdf_path, docx_path)
  ```
  - 輸出路徑：`os.path.splitext(f["path"])[0] + ".docx"`（與原 PDF 同目錄）
  - 進度：`self._set_progress(idx / total * 100, f"轉換 Word ({idx+1}/{total})：{f['name']}")`
- 部分失敗：成功的繼續執行，最後統一顯示失敗清單
- 完成 messagebox：`✓ N 個 PDF 已轉換為 Word\n存放位置：各原始檔所在資料夾`

### 輸出命名

`{原始PDF basename}.docx`，存於原 PDF 所在資料夾。

---

## 二、PDF 分割（逐頁）

### 核心函式

**`App.split_pdf()`**
- 取得清單中目前選取的項目 `self._sel_idx()`
- 若無選取 → `messagebox.showwarning("提示", "請先在清單中選取一個 PDF 檔案")`
- 若選取的是 Word 檔 → `messagebox.showwarning("提示", "請先將 Word 轉換為 PDF")`
- 彈出 `filedialog.askdirectory(title="選擇分割後 PDF 的儲存資料夾")`，使用者取消則直接返回
- 使用已有的 `pypdf.PdfReader` + `pypdf.PdfWriter`，逐頁寫出：
  ```python
  for i, page in enumerate(reader.pages, 1):
      writer = PdfWriter()
      writer.add_page(page)
      out_path = os.path.join(out_dir, f"{base}_p{i:03d}.pdf")
      with open(out_path, "wb") as fh:
          writer.write(fh)
      self._set_progress(i / total * 100, f"分割第 {i}/{total} 頁：{name}")
  ```
- 完成 messagebox：`✓ {檔名} 已分割為 N 個 PDF\n儲存位置：{out_dir}`
- 分割後原檔仍保留在清單中，不做任何變更

### 輸出命名

`{原始PDF basename}_p001.pdf`、`_p002.pdf`⋯ 零填充三位數。

---

## 三、UI 調整

### 底部按鈕列重排（兩行）

**調整前（單行）：**
```
[狀態文字] [Word→PDF] [PDF→JPG] [合併PDF→]
```

**調整後（兩行）：**
```
行 0（按鈕列）：[Word→PDF]  [PDF→JPG]  [PDF→Word]  [PDF 分割]  [合併 PDF →]
行 1（狀態列）：[狀態文字.............................................]
```

- 按鈕順序：`轉換 Word → PDF`（Ghost）→ `PDF → JPG`（Green）→ `PDF → Word`（Word Blue 新色）→ `PDF 分割`（Muted）→ `合併 PDF →`（Primary，最右）
- `PDF → Word` 按鈕新增 `WordBlue.TButton` 樣式：`background="#1d4ed8"`、`foreground="white"`（沿用已定義的 `WORD_BLUE` 色票）

### Header 副標題

```
合併 ／ Word 轉 PDF ／ PDF 轉 JPG ／ PDF 轉 Word ／ PDF 分割
```

---

## 四、錯誤處理原則

| 情境 | 處理方式 |
|------|---------|
| `pdf2docx` 安裝失敗 | showwarning + 顯示 `pip install pdf2docx` |
| 個別 PDF 轉 Word 失敗 | 繼續其他檔案，最後 showerror 列出失敗清單 |
| 分割時無選取 | showwarning 提示 |
| 分割時選取 Word 檔 | showwarning 提示先轉換 |
| 分割寫檔 PermissionError | showerror 提示檔案被佔用 |

---

## 五、不在本次範圍內

- PDF 分割的「自訂頁碼範圍」或「固定每 N 頁」模式
- PDF → Word 的品質設定（DPI、圖片解析度等）
- 分割後自動將拆出的檔案加回清單
